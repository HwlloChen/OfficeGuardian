import os
import logging
from logging.handlers import RotatingFileHandler
from pathlib import Path


def setup_logger(config):
    """设置日志系统"""

    # 获取日志级别
    log_level_str = getattr(config, 'logging_level', 'INFO')
    log_level = getattr(logging, log_level_str, logging.INFO)

    # 创建日志目录
    log_dir = os.path.join(str(Path.home()), '.audio_equalizer', 'logs')
    if not os.path.exists(log_dir):
        os.makedirs(log_dir)

    # 配置根日志记录器
    logger = logging.getLogger('OfficeGuardian')
    logger.setLevel(log_level)

    # 清除现有处理器
    if logger.handlers:
        logger.handlers = []

    # 创建控制台处理器
    console_handler = logging.StreamHandler()
    console_handler.setLevel(log_level)

    # 创建文件处理器 (轮换日志文件，限制大小为1MB，保留5个备份)
    file_handler = RotatingFileHandler(
        os.path.join(log_dir, 'audio_equalizer.log'),
        maxBytes=1024*1024,
        backupCount=5
    )
    file_handler.setLevel(log_level)

    # 创建格式化器
    formatter = logging.Formatter(
        '%(asctime)s - %(name)s - %(levelname)s - %(message)s')
    console_handler.setFormatter(formatter)
    file_handler.setFormatter(formatter)

    # 添加处理器到日志记录器
    logger.addHandler(console_handler)
    logger.addHandler(file_handler)

    logger.info(f"日志系统初始化完成，级别: {log_level_str}")

    return logger
