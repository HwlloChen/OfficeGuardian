import numpy as np
import queue
import threading
import time
import logging
from comtypes import CLSCTX_ALL
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume, IAudioMeterInformation
from collections import deque
import array


class AudioAnalyzer:
    """负责分析音频输出的响度大小"""

    def __init__(self, config):
        self.config = config
        self.logger = logging.getLogger('OfficeGuardian.AudioAnalyzer')
        self.stop_event = threading.Event()
        self.analysis_thread = None
        self.current_db = -100.0
        self.is_audio_playing = False
        self.over_max_duration = 0
        self.under_min_duration = 0
        self.last_check_time = time.time()
        self.callback = None
        self.db_history = deque(maxlen=60)  # 3秒的历史数据(60个采样点，采样率20Hz)
        self.last_average_update = time.time()
        self.current_average_db = -100.0

    def start_analyzing(self, callback=None):
        """开始分析音频输出"""
        try:
            # 获取默认音频输出会话
            devices = AudioUtilities.GetSpeakers()
            interface = devices.Activate(
                IAudioMeterInformation._iid_, CLSCTX_ALL, None)
            self.meter = interface.QueryInterface(IAudioMeterInformation)

            self.callback = callback
            self.stop_event.clear()
            self.analysis_thread = threading.Thread(target=self._analysis_loop)
            self.analysis_thread.daemon = True
            self.analysis_thread.start()

            self.logger.info("音频分析已启动")

        except Exception as e:
            self.logger.error(f"启动音频分析失败: {e}")
            raise

    def stop_analyzing(self):
        """停止分析音频输出"""
        self.stop_event.set()
        if self.analysis_thread and self.analysis_thread.is_alive():
            self.analysis_thread.join(timeout=1.0)
        self.meter = None
        self.logger.info("音频分析已停止")

    def _analysis_loop(self):
        """分析音频的线程函数"""
        while not self.stop_event.is_set():
            try:
                # 获取真实响度
                real_db = self.get_real_db()
                # 获取输出响度（现在是平均值）
                output_db = self.get_current_db()
                self.current_db = output_db

                # 使用真实响度判断是否有音频播放
                if real_db > self.config.audio_threshold:
                    self.is_audio_playing = True
                    current_time = time.time()
                    time_diff = current_time - self.last_check_time

                    # 使用平均输出响度进行阈值判断
                    if output_db > self.config.max_db:
                        self.over_max_duration += time_diff
                        self.under_min_duration = 0
                    elif output_db < self.config.min_db:
                        self.under_min_duration += time_diff
                        self.over_max_duration = 0
                    else:
                        self.over_max_duration = 0
                        self.under_min_duration = 0

                    self.last_check_time = current_time

                    # 使用较短的时间阈值，因为现在使用的是平均值
                    if self.over_max_duration >= self.config.interval_max and self.callback:
                        self.callback("over_max", output_db)
                        self.over_max_duration = 0
                    elif self.under_min_duration >= self.config.interval_min and self.callback:
                        self.callback("under_min", output_db)
                        self.under_min_duration = 0
                else:
                    self.is_audio_playing = False
                    self.over_max_duration = 0
                    self.under_min_duration = 0

                time.sleep(0.05)  # 20Hz采样率

            except Exception as e:
                self.logger.error(f"音频分析错误: {e}")
                time.sleep(0.1)

    def get_current_db(self):
        """获取当前分贝值（经过系统音量调节后的输出响度）"""
        try:
            # 获取原始峰值
            peak = self.meter.GetPeakValue()
            # 获取当前系统音量
            volume = AudioUtilities.GetSpeakers().Activate(
                IAudioEndpointVolume._iid_, CLSCTX_ALL, None
            ).QueryInterface(IAudioEndpointVolume).GetMasterVolumeLevelScalar()

            if peak > 0:
                # 原始分贝值
                original_db = 20 * np.log10(peak)
                # 补偿系统音量的影响（音量越小，削减越多）
                volume_db = 20 * np.log10(volume) if volume > 0 else -100.0
                output_db = original_db + volume_db

                # 更新历史数据
                self.db_history.append(output_db)

                # 每50ms更新一次平均值
                current_time = time.time()
                if current_time - self.last_average_update >= 0.05:
                    self._update_average_db()
                    self.last_average_update = current_time

                return self.current_average_db
            return -100.0
        except:
            return -100.0

    def _update_average_db(self):
        """更新平均分贝值"""
        if len(self.db_history) > 0:
            # 将分贝值转换回能量域进行平均
            energies = [10 ** (db/20) for db in self.db_history]
            avg_energy = sum(energies) / len(energies)
            # 将平均能量转换回分贝域
            self.current_average_db = 20 * np.log10(avg_energy)

    def get_real_db(self):
        """获取真实响度（未经音量调节的原始分贝值）"""
        try:
            peak = self.meter.GetPeakValue()
            if peak > 0:
                return 20 * np.log10(peak)
            return -100.0
        except:
            return -100.0

    def is_playing(self):
        """判断当前是否有音频播放"""
        return self.is_audio_playing
