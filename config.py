import os
import json
import logging
from pathlib import Path


class Config:
    """配置管理类"""

    # 默认配置
    DEFAULT_CONFIG = {
        'max_db': -10.0,          # 最大响度阈值（分贝）
        'min_db': -40.0,          # 最小响度阈值（分贝）
        'audio_threshold': -60.0,  # 有音频判断阈值（分贝）
        'auto_start': False,      # 开机自启动
        'start_minimized': False,  # 启动时最小化
        'check_interval': 0.5,    # 检查间隔（秒）
        'was_calibrated': False,  # 是否已校准
        'logging_level': 'INFO',   # 日志级别
        'interval_max': 2,         # 音量过大调整间隔（秒）
        'interval_min': 8,         # 音量过小调整间隔（秒）
        'volume_change_k': 0.2,        # 渐进式音量调整系数k
        'device_id': None,       # 设备ID
    }

    def __init__(self):
        self.logger = logging.getLogger('OfficeGuardian.Config')
        self.config_dir = self._get_config_dir()
        self.config_file = os.path.join(self.config_dir, 'config.json')
        self._load_config()

    def _get_config_dir(self):
        """获取配置文件目录"""
        # 在程序所在目录下创建配置目录
        base_dir = os.path.dirname(os.path.abspath(__file__))
        config_dir = os.path.join(base_dir, 'config')
        if not os.path.exists(config_dir):
            os.makedirs(config_dir)
        return config_dir

    def _load_config(self):
        """加载配置文件"""
        # 如果配置文件存在，从文件加载
        if os.path.exists(self.config_file):
            try:
                with open(self.config_file, 'r') as f:
                    config_data = json.load(f)

                # 更新默认配置
                merged_config = self.DEFAULT_CONFIG.copy()
                merged_config.update(config_data)

                # 应用配置
                for key, value in merged_config.items():
                    setattr(self, key, value)

                self.logger.debug("配置已从文件加载")
            except Exception as e:
                self.logger.error(f"加载配置文件失败: {e}, 使用默认配置")
                self._apply_default_config()
            finally:
                self.logger.debug("配置加载完成")
        else:
            # 使用默认配置
            self._apply_default_config()
            self.logger.info("使用默认配置")

    def _apply_default_config(self):
        """应用默认配置"""
        for key, value in self.DEFAULT_CONFIG.items():
            setattr(self, key, value)

    def save_config(self):
        """保存配置到文件"""
        config_data = {}
        for key in self.DEFAULT_CONFIG.keys():
            config_data[key] = getattr(self, key)

        try:
            with open(self.config_file, 'w') as f:
                json.dump(config_data, f, indent=4)
            self.logger.info("配置已保存到文件")
            return True
        except Exception as e:
            self.logger.error(f"保存配置文件失败: {e}")
            return False

    def update(self, **kwargs):
        """更新配置"""
        for key, value in kwargs.items():
            if key in self.DEFAULT_CONFIG:
                setattr(self, key, value)
                self.logger.info(f"配置已更新: {key}={value}")
            else:
                self.logger.warning(f"未知配置项: {key}")

        # 保存更新后的配置
        return self.save_config()

    def reset_to_default(self):
        """重置为默认配置"""
        self._apply_default_config()
        self.logger.info("配置已重置为默认值")
        return self.save_config()
