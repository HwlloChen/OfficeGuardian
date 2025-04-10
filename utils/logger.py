import logging


def setup_logger(config):
    """设置日志系统"""
    # 获取日志级别
    log_level_str = getattr(config, 'logging_level', 'INFO')
    log_level = getattr(logging, log_level_str, logging.INFO)

    # 配置根日志记录器
    logger = logging.getLogger('OfficeGuardian')
    logger.setLevel(log_level)

    # 清除现有处理器
    if logger.handlers:
        logger.handlers = []

    # 创建控制台处理器
    console_handler = logging.StreamHandler()
    console_handler.setLevel(log_level)

    # 创建格式化器
    formatter = logging.Formatter(
        '%(asctime)s - %(name)s - %(levelname)s - %(message)s')
    console_handler.setFormatter(formatter)

    # 添加处理器到日志记录器
    logger.addHandler(console_handler)

    logger.info(f"日志系统初始化完成，级别: {log_level_str}")

    return logger
