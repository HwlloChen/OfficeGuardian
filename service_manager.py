import os
import sys
import logging
import subprocess
from pathlib import Path
import shutil

class ServiceManager:
    """Windows服务管理类，用于添加/删除开机启动项"""
    
    def __init__(self):
        self.logger = logging.getLogger('AudioEqualizer.ServiceManager')
        self.app_name = "AudioEqualizer"
        self.exe_path = sys.executable
        
        # 获取当前用户的启动文件夹路径
        self.startup_folder = str(Path.home() / "AppData" / "Roaming" / "Microsoft" / "Windows" / "Start Menu" / "Programs" / "Startup")
        
        if self.exe_path.endswith('python.exe'):
            # 如果是以Python解释器运行，则使用脚本路径
            self.app_path = os.path.abspath(sys.argv[0])
            # 创建启动脚本
            self.startup_script = os.path.join(self.startup_folder, f"{self.app_name}.bat")
        else:
            # 如果是打包后的可执行文件
            self.app_path = self.exe_path
            # 创建快捷方式
            self.startup_link = os.path.join(self.startup_folder, f"{self.app_name}.lnk")
            
    def create_shortcut(self):
        """创建快捷方式"""
        try:
            import win32com.client
            shell = win32com.client.Dispatch("WScript.Shell")
            shortcut = shell.CreateShortCut(self.startup_link)
            shortcut.Targetpath = self.app_path
            shortcut.Arguments = "--minimized"
            shortcut.WorkingDirectory = os.path.dirname(self.app_path)
            shortcut.save()
            return True
        except Exception as e:
            self.logger.error(f"创建快捷方式失败: {e}")
            return False
            
    def create_batch_script(self):
        """创建启动脚本"""
        try:
            with open(self.startup_script, 'w') as f:
                f.write(f'@echo off\n')
                f.write(f'start "" "{self.exe_path}" "{self.app_path}" --minimized\n')
            return True
        except Exception as e:
            self.logger.error(f"创建启动脚本失败: {e}")
            return False
            
    def add_to_startup(self):
        """添加到开机启动"""
        try:
            if self.exe_path.endswith('python.exe'):
                # Python脚本模式：创建批处理文件
                success = self.create_batch_script()
            else:
                # 可执行文件模式：创建快捷方式
                success = self.create_shortcut()
                
            if success:
                self.logger.info(f"已添加到开机启动: {self.app_path}")
            return success
            
        except Exception as e:
            self.logger.error(f"添加开机启动失败: {e}")
            return False
            
    def remove_from_startup(self):
        """从开机启动中移除"""
        try:
            # 删除可能存在的启动项
            startup_files = [
                os.path.join(self.startup_folder, f"{self.app_name}.lnk"),
                os.path.join(self.startup_folder, f"{self.app_name}.bat")
            ]
            
            for file in startup_files:
                if os.path.exists(file):
                    os.remove(file)
                    self.logger.info(f"已从开机启动中移除: {file}")
                    
            return True
        except Exception as e:
            self.logger.error(f"移除开机启动失败: {e}")
            return False
