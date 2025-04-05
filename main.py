import sys
import os
import logging
import argparse
from PySide6.QtWidgets import QApplication
from PySide6.QtCore import QThread, Signal, QObject
import time

from audio_analyzer import AudioAnalyzer
from volume_controller import VolumeController
from config import Config
from logger import setup_logger
from gui import MainWindow
from service_manager import ServiceManager


class OfficeGuardianWorker(QObject):
    """音频均衡器工作线程"""

    volume_changed = Signal(float)  # 音量变化信号

    def __init__(self, audio_analyzer, volume_controller, config):
        super().__init__()
        self.logger = logging.getLogger('OfficeGuardian.Worker')
        self.audio_analyzer = audio_analyzer
        self.volume_controller = volume_controller
        self.config = config
        self.running = False

    def start(self):
        """开始音频均衡处理"""
        self.running = True
        self.audio_analyzer.start_analyzing(callback=self.on_audio_event)
        self.logger.info("音频均衡处理已启动")

    def stop(self):
        """停止音频均衡处理"""
        self.running = False
        self.audio_analyzer.stop_analyzing()
        self.logger.info("音频均衡处理已停止")

    def on_audio_event(self, event_type, current_db):
        """
        音频事件回调

        Args:
            event_type: 事件类型 ('over_max' 或 'under_min')
            current_db: 当前分贝值
        """
        if not self.running:
            return

        if event_type == "over_max":
            # 音频响度超过最大阈值，降低音量
            self.logger.info(
                f"响度过高: {current_db:.2f} dB > {self.config.max_db:.2f} dB")
            new_volume = self.volume_controller.adjust_volume_for_db(
                current_db, self.config.max_db)
            self.volume_changed.emit(new_volume)

        elif event_type == "under_min":
            # 音频响度低于最小阈值，提高音量
            self.logger.info(
                f"响度过低: {current_db:.2f} dB < {self.config.min_db:.2f} dB")
            new_volume = self.volume_controller.adjust_volume_for_db(
                current_db, self.config.min_db)
            self.volume_changed.emit(new_volume)


def main():
    """程序主入口"""
    # 解析命令行参数
    parser = argparse.ArgumentParser(description='办公室的大盾 - 音频响度均衡器')
    parser.add_argument('--minimized', action='store_true', help='以最小化方式启动')
    parser.add_argument('--service', action='store_true', help='以服务方式启动')
    args = parser.parse_args()

    # 加载配置
    config = Config()

    # 设置日志
    logger = setup_logger(config)
    logger.info("音频响度均衡器启动中...")

    # 创建QApplication实例
    app = QApplication(sys.argv)
    app.setApplicationName("OfficeGuardian")
    app.setQuitOnLastWindowClosed(False)  # 关闭窗口时不退出应用

    # 创建服务管理器
    service_manager = ServiceManager()

    # 创建音频分析器和音量控制器
    try:
        volume_controller = VolumeController(config)
        audio_analyzer = AudioAnalyzer(config)

        # 创建工作线程
        worker = OfficeGuardianWorker(
            audio_analyzer, volume_controller, config)
        worker_thread = QThread()
        worker.moveToThread(worker_thread)
        worker_thread.started.connect(worker.start)
        worker_thread.start()

        # 创建并显示主窗口
        main_window = MainWindow(
            audio_analyzer, volume_controller, config, service_manager)

        # 根据参数和配置决定是否最小化启动
        if args.minimized or (config.start_minimized and not args.service):
            logger.info("程序以最小化方式启动")
        else:
            main_window.show()

        # 检查开机自启动设置
        if config.auto_start:
            if not service_manager.add_to_startup():
                logger.warning("添加开机启动项失败")

        # 开始应用主循环
        exit_code = app.exec()

        # 清理资源
        worker.stop()
        worker_thread.quit()
        worker_thread.wait()

        logger.info("程序正常退出")
        return exit_code

    except Exception as e:
        logger.error(f"程序启动失败: {e}")
        return 1


if __name__ == "__main__":
    sys.exit(main())
