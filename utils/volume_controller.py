import logging
import platform
from ctypes import cast, POINTER

# Windows音量控制
if platform.system() == 'Windows':
    from comtypes import CLSCTX_ALL
    from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume


class VolumeController:
    """控制系统音量"""

    def __init__(self, config):
        self.config = config
        self.logger = logging.getLogger('OfficeGuardian.VolumeController')
        self.os_type = platform.system()
        self.current_volume = 0
        self.device_id = config.device_id
        self.volume = None
        self._initialize_volume_controller()

    def _initialize_volume_controller(self):
        """初始化音量控制接口"""
        try:
            if self.os_type == 'Windows':
                if self.device_id is None:
                    # 使用默认设备
                    speakers = AudioUtilities.GetSpeakers()
                    self.logger.debug("使用默认音频设备")
                else:
                    # 尝试找到指定ID的设备
                    device_enumerator = AudioUtilities.GetDeviceEnumerator()
                    try:
                        speakers = device_enumerator.GetDevice(self.device_id)
                        self.logger.debug(f"使用设备ID: {self.device_id}")
                    except Exception as e:
                        self.logger.warning(
                            f"找不到指定设备({self.device_id})，使用默认设备: {e}")
                        speakers = AudioUtilities.GetSpeakers()

                interface = speakers.Activate(
                    IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
                self.volume = cast(interface, POINTER(IAudioEndpointVolume))
                self.current_volume = self.volume.GetMasterVolumeLevelScalar()
                self.device_id = speakers.GetId()  # 更新设备ID
                self.logger.debug("音量控制接口初始化成功")
            else:
                self.logger.error(f"不支持的操作系统: {self.os_type}")
                raise NotImplementedError(f"不支持的操作系统: {self.os_type}")
        except Exception as e:
            self.logger.error(f"初始化音量控制失败: {e}")
            raise
        finally:
            self.logger.debug("音量控制接口初始化完成")

    def set_device(self, device_id):
        """切换音频设备"""
        if self.device_id != device_id:
            self.device_id = device_id
            self._initialize_volume_controller()
            self.logger.info(f"已切换到设备: {device_id}")

    def get_volume(self):
        """获取当前系统音量 (0.0 到 1.0)"""
        try:
            if self.os_type == 'Windows':
                self.current_volume = self.volume.GetMasterVolumeLevelScalar()
                # 特殊情况：当音量小于0.8%时，将音量调整为1%
                if self.current_volume < 0.008:
                    self.set_volume(0.01)
                    self.logger.warning("音量过低，自动调整至1%")
                return self.current_volume
        except Exception as e:
            self.logger.error(f"获取音量失败: {e}")
            return 0.0

    def set_volume(self, volume_level):
        """
        设置系统音量

        Args:
            volume_level: 音量级别 (0.0 到 1.0)
        """
        # 确保音量在有效范围内
        volume_level = max(0.0, min(1.0, volume_level))

        try:
            if self.os_type == 'Windows':
                self.volume.SetMasterVolumeLevelScalar(volume_level, None)
                self.current_volume = volume_level
                self.logger.info(f"系统音量已设置为: {volume_level:.2f}")
            else:
                self.logger.error(f"不支持的操作系统: {self.os_type}")
        except Exception as e:
            self.logger.error(f"设置音量失败: {e}")

    def increase_volume(self):
        """增加系统音量"""
        current = self.get_volume()
        self.set_volume(current + 0.02)
        return self.get_volume()

    def decrease_volume(self):
        """减小系统音量"""
        current = self.get_volume()
        self.set_volume(current - 0.02)
        return self.get_volume()

    def adjust_volume_for_db(self, current_db, target_db):
        """
        根据当前分贝值和目标分贝值调整系统音量

        Args:
            current_db: 当前音频输出的分贝值
            target_db: 目标分贝值
        """
        # 简单线性调整 - 可以根据实际情况优化算法
        db_diff = target_db - current_db

        # 转换为音量调整比例 - 这里使用经验公式，可能需要根据测试调整
        # 假设分贝与音量大致是对数关系
        if abs(db_diff) < 1.0:  # 如果差异很小，不做调整
            return self.get_volume()

        # 根据分贝差异计算音量调整
        volume_change = (db_diff / 20.0) * \
            self.config.volume_change_k  # 简单比例关系，需要调整

        # 获取当前音量并应用变化
        current_volume = self.get_volume()
        new_volume = current_volume + volume_change

        # 应用新的音量
        self.set_volume(new_volume)
        self.logger.info(
            f"调整音量: 当前 {current_db:.2f}dB, 目标 {target_db:.2f}dB, 音量从 {current_volume:.2f} 调整到 {new_volume:.2f}")

        return new_volume
