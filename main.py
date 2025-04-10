import sys, os
import logging
import argparse
import wx
from utils.audio_analyzer import AudioAnalyzer
from utils.volume_controller import VolumeController
from utils.config import Config
from utils.logger import setup_logger
from utils.gui import MainFrame
from utils.service_manager import ServiceManager

class OfficeGuardianWorker:
    """音频均衡器工作类"""

    def __init__(self, audio_analyzer, volume_controller, config, gui=None):
        self.logger = logging.getLogger('OfficeGuardian.Worker')
        self.audio_analyzer = audio_analyzer
        self.volume_controller = volume_controller
        self.config = config
        self.gui = gui
        self.running = False

    def start(self):
        """开始音频均衡处理"""
        self.running = True
        self.audio_analyzer.start_analyzing(callback=self.on_audio_event)
        self.logger.debug("音频均衡处理已启动")

    def stop(self):
        """停止音频均衡处理"""
        self.running = False
        self.audio_analyzer.stop_analyzing()
        self.logger.info("音频均衡处理已停止")

    def on_audio_event(self, event_type, current_db):
        """音频事件回调"""
        if not self.running:
            return

        if event_type == "over_max":
            self.logger.debug(
                f"响度过高: {current_db:.2f} dB > {self.config.max_db:.2f} dB")
            new_volume = self.volume_controller.adjust_volume_for_db(
                current_db, self.config.max_db)
            if self.gui:
                wx.CallAfter(self.gui.update_volume, new_volume)

        elif event_type == "under_min":
            self.logger.debug(
                f"响度过低: {current_db:.2f} dB < {self.config.min_db:.2f} dB")
            new_volume = self.volume_controller.adjust_volume_for_db(
                current_db, self.config.min_db)
            if self.gui:
                wx.CallAfter(self.gui.update_volume, new_volume)

def main():
    """程序主入口"""
    # 解析命令行参数
    parser = argparse.ArgumentParser(description='办公室的大盾 - 音频响度均衡器')
    parser.add_argument('--minimized', action='store_true', help='以最小化方式启动')
    parser.add_argument('--service', action='store_true', help='以服务方式启动')
    args = parser.parse_args()

    # 加载配置
    config = Config(os.path.dirname(os.path.abspath(__file__)))

    # 设置日志
    logger = setup_logger(config)
    logger.info("音频响度均衡器启动中...")

    # 创建wxPython应用实例
    app = wx.App()

    try:
        # 创建必要的组件
        volume_controller = VolumeController(config)
        audio_analyzer = AudioAnalyzer(config)
        service_manager = ServiceManager()

        # 创建主窗口
        frame = MainFrame(None, audio_analyzer, volume_controller, config, service_manager)
        app.SetTopWindow(frame)

        # 创建工作线程
        worker = OfficeGuardianWorker(audio_analyzer, volume_controller, config, frame)
        frame.set_worker(worker)  # 设置 worker 实例
        worker.start()

        # 根据参数和配置决定是否最小化启动
        if args.minimized or (config.start_minimized and not args.service):
            logger.info("程序以最小化方式启动")
            frame.Hide()
        else:
            frame.Show()

        # 检查开机自启动设置
        if config.auto_start:
            if not service_manager.add_to_startup():
                logger.warning("添加开机启动项失败")

        # 开始应用主循环
        exit_code = app.MainLoop()

        # 清理资源
        worker.stop()
        logger.info("程序正常退出")
        return exit_code

    except Exception as e:
        logger.critical(f"程序启动失败: {e}", exc_info=True)
        return 1
    finally:
        logger.info("程序退出")

if __name__ == "__main__":
    sys.exit(main())
