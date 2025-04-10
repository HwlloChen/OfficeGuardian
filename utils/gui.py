import wx
import wx.adv
import logging
from utils.calibration import CalibrationDialog
from pycaw.pycaw import AudioUtilities
from utils.about_dialog import AboutDialog

class LogHandler(logging.Handler):
    """自定义日志处理器，支持带颜色的日志显示"""
    def __init__(self, text_ctrl):
        super().__init__()
        self.text_ctrl = text_ctrl
        
        # 为不同级别定义更易区分的颜色
        self.colors = {
            logging.DEBUG: wx.Colour(120, 120, 120),    # 较深的灰色
            logging.INFO: wx.Colour(0, 128, 255),       # 天蓝色
            logging.WARNING: wx.Colour(255, 128, 0),    # 橙色
            logging.ERROR: wx.Colour(255, 0, 0),        # 红色
            logging.CRITICAL: wx.Colour(153, 0, 0)      # 深红色
        }

    def emit(self, record):
        try:
            msg = self.format(record)
            wx.CallAfter(self._write_log, record, msg)
        except Exception:
            self.handleError(record)

    def _write_log(self, record, msg):
        """在GUI线程中写入日志"""
        try:
            # 设置文本属性
            attr = wx.TextAttr()
            attr.SetTextColour(self.colors.get(record.levelno, wx.Colour(0, 0, 0)))
            
            # 根据日志级别设置字体样式
            if record.levelno >= logging.ERROR:
                font = wx.Font(wx.NORMAL_FONT.GetPointSize(), wx.FONTFAMILY_DEFAULT,
                             wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_BOLD)
            else:
                font = wx.Font(wx.NORMAL_FONT.GetPointSize(), wx.FONTFAMILY_DEFAULT,
                             wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_NORMAL)
                
            attr.SetFont(font)
            
            # 保存当前位置
            current_pos = self.text_ctrl.GetLastPosition()
            
            # 写入日志文本
            self.text_ctrl.AppendText(msg + '\n')
            
            # 应用样式到新添加的文本
            self.text_ctrl.SetStyle(current_pos, self.text_ctrl.GetLastPosition(), attr)
            
            # 确保最新的日志可见
            self.text_ctrl.ShowPosition(self.text_ctrl.GetLastPosition())
            
            # 添加额外的换行以提高可读性
            self.text_ctrl.AppendText('\n')
            
            # 如果文本过长，保留最后的1000行
            self._trim_log()
            
        except Exception as e:
            print(f"Error writing log: {e}")  # 用于调试

    def _trim_log(self):
        """保持日志在合理的长度范围内"""
        try:
            max_lines = 1000
            text = self.text_ctrl.GetValue()
            lines = text.split('\n')
            
            if len(lines) > max_lines:
                # 保留最后的max_lines行
                new_text = '\n'.join(lines[-max_lines:])
                self.text_ctrl.SetValue(new_text)
                # 移动光标到末尾
                self.text_ctrl.SetInsertionPointEnd()
        except Exception:
            pass

class MainFrame(wx.Frame):
    """主窗口"""

    def __init__(self, parent, audio_analyzer, volume_controller, config, service_manager):
        super().__init__(parent, title="办公室的大盾 - 音频响度均衡器", 
                        size=(1000, 650))
        
        self.audio_analyzer = audio_analyzer
        self.volume_controller = volume_controller
        self.config = config
        self.service_manager = service_manager
        self.logger = logging.getLogger('OfficeGuardian.GUI')
        self.worker = None  # 添加 worker 属性

        # 创建菜单栏
        self._create_menu_bar()

        # 初始化界面
        self._init_ui()
        self._create_tray_icon()
        
        # 创建定时器用于更新显示
        self.update_timer = wx.Timer(self)
        self.Bind(wx.EVT_TIMER, self._on_timer, self.update_timer)
        self.update_timer.Start(100)  # 100ms更新一次

        # 绑定关闭事件
        self.Bind(wx.EVT_CLOSE, self.on_close)

        # 设置日志处理器
        log_handler = LogHandler(self.log_text)
        log_handler.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s - %(message)s'))
        logging.getLogger('OfficeGuardian').addHandler(log_handler)

    def set_worker(self, worker):
        """设置 worker 实例"""
        self.worker = worker

    def _create_menu_bar(self):
        """创建菜单栏"""
        menubar = wx.MenuBar()
        
        # 文件菜单
        file_menu = wx.Menu()
        exit_item = file_menu.Append(wx.ID_EXIT, "退出\tCtrl+Q", "退出程序")
        self.Bind(wx.EVT_MENU, self._on_exit, exit_item)
        menubar.Append(file_menu, "文件")
        
        # 帮助菜单
        help_menu = wx.Menu()
        about_item = help_menu.Append(wx.ID_ABOUT, "关于\tF1", "关于本程序")
        self.Bind(wx.EVT_MENU, self._on_about, about_item)
        menubar.Append(help_menu, "帮助")
        
        self.SetMenuBar(menubar)

    def _init_ui(self):
        """初始化UI"""
        panel = wx.Panel(self)
        main_sizer = wx.BoxSizer(wx.HORIZONTAL)

        # 左侧控制面板
        left_panel = self._create_left_panel(panel)
        main_sizer.Add(left_panel, 3, wx.EXPAND | wx.ALL, 5)

        # 右侧日志面板
        right_panel = self._create_right_panel(panel)
        main_sizer.Add(right_panel, 2, wx.EXPAND | wx.ALL, 5)

        panel.SetSizer(main_sizer)

    def _create_left_panel(self, parent):
        """创建左侧控制面板"""
        panel = wx.Panel(parent)
        sizer = wx.BoxSizer(wx.VERTICAL)

        # 设备选择
        device_box = wx.StaticBox(panel, label="音频设备选择")
        device_sizer = wx.StaticBoxSizer(device_box, wx.VERTICAL)
        self.device_combo = wx.Choice(panel)
        self._populate_device_list()
        device_sizer.Add(self.device_combo, 0, wx.EXPAND | wx.ALL, 5)
        sizer.Add(device_sizer, 0, wx.EXPAND | wx.ALL, 5)

        # 修改音频状态部分
        status_box = wx.StaticBox(panel, label="音频状态")
        status_sizer = wx.StaticBoxSizer(status_box, wx.VERTICAL)
        
        # 自动调节开关
        auto_adjust_sizer = wx.BoxSizer(wx.HORIZONTAL)
        self.auto_adjust_toggle = wx.ToggleButton(panel, label="自动音量调节")
        self.auto_adjust_toggle.SetValue(True)  # 默认开启
        self.auto_adjust_toggle.Bind(wx.EVT_TOGGLEBUTTON, self._on_auto_adjust_toggle)
        auto_adjust_sizer.Add(self.auto_adjust_toggle, 0, wx.ALL, 5)
        
        # 自动调节状态
        self.auto_adjust_status = wx.StaticText(panel, label="状态: 已启用")
        self.auto_adjust_status.SetForegroundColour(wx.Colour(0, 128, 0))  # 深绿色
        auto_adjust_sizer.Add(self.auto_adjust_status, 0, wx.ALL | wx.ALIGN_CENTER_VERTICAL, 5)
        
        status_sizer.Add(auto_adjust_sizer, 0, wx.EXPAND)
        
        # 当前响度
        self.db_label = wx.StaticText(panel, label="当前响度: -60.0 dB")
        status_sizer.Add(self.db_label, 0, wx.ALL, 5)
        
        # 系统音量
        self.volume_label = wx.StaticText(panel, label="系统音量: 50%")
        status_sizer.Add(self.volume_label, 0, wx.ALL, 5)
        
        # 音频状态
        self.status_label = wx.StaticText(panel, label="音频状态: 无音频")
        status_sizer.Add(self.status_label, 0, wx.ALL, 5)
        
        sizer.Add(status_sizer, 0, wx.EXPAND | wx.ALL, 5)

        # 参数设置
        settings_box = wx.StaticBox(panel, label="参数设置")
        settings_sizer = wx.StaticBoxSizer(settings_box, wx.VERTICAL)

        # 最大响度
        max_db_sizer = wx.BoxSizer(wx.HORIZONTAL)
        max_db_sizer.Add(wx.StaticText(panel, label="最大响度:"), 0, wx.ALIGN_CENTER_VERTICAL)
        self.max_db_spin = wx.SpinCtrlDouble(panel, value=f"{self.config.max_db:.1f}",
                                           min=-60.0, max=0.0, inc=0.4)
        self.max_db_spin.SetDigits(1)  # 显示一位小数
        self.max_db_spin.Bind(wx.EVT_SPINCTRLDOUBLE, self._on_max_db_changed)
        max_db_sizer.Add(self.max_db_spin, 1, wx.LEFT, 5)
        settings_sizer.Add(max_db_sizer, 0, wx.EXPAND | wx.ALL, 5)

        # 最小响度
        min_db_sizer = wx.BoxSizer(wx.HORIZONTAL)
        min_db_sizer.Add(wx.StaticText(panel, label="最小响度:"), 0, wx.ALIGN_CENTER_VERTICAL)
        self.min_db_spin = wx.SpinCtrlDouble(panel, value=f"{self.config.min_db:.1f}",
                                           min=-80.0, max=-10.0, inc=0.4)
        self.min_db_spin.SetDigits(1)  # 显示一位小数
        self.min_db_spin.Bind(wx.EVT_SPINCTRLDOUBLE, self._on_min_db_changed)
        min_db_sizer.Add(self.min_db_spin, 1, wx.LEFT, 5)
        settings_sizer.Add(min_db_sizer, 0, wx.EXPAND | wx.ALL, 5)

        # 音频检测阈值
        threshold_sizer = wx.BoxSizer(wx.HORIZONTAL)
        threshold_sizer.Add(wx.StaticText(panel, label="音频检测阈值:"), 0, wx.ALIGN_CENTER_VERTICAL)
        self.threshold_spin = wx.SpinCtrlDouble(panel, value=f"{self.config.audio_threshold:.1f}",
                                              min=-90.0, max=-30.0, inc=0.4)
        self.threshold_spin.SetDigits(1)  # 显示一位小数
        self.threshold_spin.Bind(wx.EVT_SPINCTRLDOUBLE, self._on_threshold_changed)
        threshold_sizer.Add(self.threshold_spin, 1, wx.LEFT, 5)
        settings_sizer.Add(threshold_sizer, 0, wx.EXPAND | wx.ALL, 5)

        # 音量过大调整间隔
        interval_max_sizer = wx.BoxSizer(wx.HORIZONTAL)
        interval_max_sizer.Add(wx.StaticText(panel, label="音量过大调整间隔(秒):"), 0, wx.ALIGN_CENTER_VERTICAL)
        self.interval_max_spin = wx.SpinCtrlDouble(panel, value=str(self.config.interval_max),
                                                 min=0.5, max=10.0, inc=0.5)
        self.interval_max_spin.Bind(wx.EVT_SPINCTRLDOUBLE, self._on_interval_max_changed)
        interval_max_sizer.Add(self.interval_max_spin, 1, wx.LEFT, 5)
        settings_sizer.Add(interval_max_sizer, 0, wx.EXPAND | wx.ALL, 5)

        # 音量过小调整间隔
        interval_min_sizer = wx.BoxSizer(wx.HORIZONTAL)
        interval_min_sizer.Add(wx.StaticText(panel, label="音量过小调整间隔(秒):"), 0, wx.ALIGN_CENTER_VERTICAL)
        self.interval_min_spin = wx.SpinCtrlDouble(panel, value=str(self.config.interval_min),
                                                 min=1.0, max=30.0, inc=1.0)
        self.interval_min_spin.Bind(wx.EVT_SPINCTRLDOUBLE, self._on_interval_min_changed)
        interval_min_sizer.Add(self.interval_min_spin, 1, wx.LEFT, 5)
        settings_sizer.Add(interval_min_sizer, 0, wx.EXPAND | wx.ALL, 5)

        # 音量调整系数
        volume_k_sizer = wx.BoxSizer(wx.HORIZONTAL)
        volume_k_sizer.Add(wx.StaticText(panel, label="音量调整系数:"), 0, wx.ALIGN_CENTER_VERTICAL)
        self.volume_k_spin = wx.SpinCtrlDouble(panel, value=str(self.config.volume_change_k),
                                             min=0.1, max=0.4, inc=0.02)
        self.volume_k_spin.Bind(wx.EVT_SPINCTRLDOUBLE, self._on_volume_k_changed)
        volume_k_sizer.Add(self.volume_k_spin, 1, wx.LEFT, 5)
        settings_sizer.Add(volume_k_sizer, 0, wx.EXPAND | wx.ALL, 5)

        # 开机自启动和启动时最小化选项
        checkbox_sizer = wx.BoxSizer(wx.VERTICAL)
        self.autostart_check = wx.CheckBox(panel, label="开机自启动")
        self.autostart_check.SetValue(self.config.auto_start)
        self.autostart_check.Bind(wx.EVT_CHECKBOX, self._on_autostart_changed)
        
        self.minimize_check = wx.CheckBox(panel, label="启动时最小化")
        self.minimize_check.SetValue(self.config.start_minimized)
        self.minimize_check.Bind(wx.EVT_CHECKBOX, self._on_minimize_changed)
        
        checkbox_sizer.Add(self.autostart_check, 0, wx.ALL, 5)
        checkbox_sizer.Add(self.minimize_check, 0, wx.ALL, 5)
        settings_sizer.Add(checkbox_sizer, 0, wx.EXPAND)

        sizer.Add(settings_sizer, 0, wx.EXPAND | wx.ALL, 5)

        # 按钮区
        button_sizer = wx.BoxSizer(wx.HORIZONTAL)
        calibrate_btn = wx.Button(panel, label="重新校准")
        calibrate_btn.Bind(wx.EVT_BUTTON, self.show_calibration_dialog)
        reset_btn = wx.Button(panel, label="恢复默认设置")
        reset_btn.Bind(wx.EVT_BUTTON, self._on_reset_clicked)
        
        button_sizer.Add(calibrate_btn, 1, wx.RIGHT, 5)
        button_sizer.Add(reset_btn, 1)
        sizer.Add(button_sizer, 0, wx.EXPAND | wx.ALL, 5)

        panel.SetSizer(sizer)
        return panel

    def _create_right_panel(self, parent):
        """创建右侧日志面板"""
        panel = wx.Panel(parent)
        sizer = wx.BoxSizer(wx.VERTICAL)

        # 日志文本框
        log_box = wx.StaticBox(panel, label="运行日志")
        log_sizer = wx.StaticBoxSizer(log_box, wx.VERTICAL)
        
        # 使用等宽字体来改善日志显示效果
        font = wx.Font(9, wx.FONTFAMILY_TELETYPE, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_NORMAL)
        self.log_text = wx.TextCtrl(panel, style=wx.TE_MULTILINE | wx.TE_READONLY | wx.TE_RICH2)
        self.log_text.SetFont(font)
        self.log_text.SetBackgroundColour(wx.Colour(250, 250, 250))  # 略微灰白的背景色
        
        log_sizer.Add(self.log_text, 1, wx.EXPAND)
        sizer.Add(log_sizer, 1, wx.EXPAND | wx.ALL, 5)

        panel.SetSizer(sizer)
        return panel

    def _create_tray_icon(self):
        """创建系统托盘图标"""
        self.tbicon = wx.adv.TaskBarIcon()
        
        # 绑定托盘事件
        self.tbicon.Bind(wx.adv.EVT_TASKBAR_LEFT_DOWN, self._on_tray_left_click)
        self.tbicon.Bind(wx.adv.EVT_TASKBAR_LEFT_DCLICK, self._on_tray_left_dclick)
        self.tbicon.Bind(wx.adv.EVT_TASKBAR_RIGHT_UP, self._on_tray_right_click)
        
        # 创建托盘菜单
        self.tray_menu = wx.Menu()
        self._create_tray_menu()
        
        # 设置图标
        icon = self._create_app_icon()
        self.tbicon.SetIcon(icon, "音频响度均衡器")
        self.SetIcon(icon)

    def _create_app_icon(self):
        """创建应用图标"""
        bmp = wx.Bitmap(32, 32)
        dc = wx.MemoryDC(bmp)
        dc.SetBackground(wx.Brush(wx.Colour(0, 120, 215)))  # Windows蓝
        dc.Clear()
        
        # 绘制简单的音量图标
        dc.SetBrush(wx.WHITE_BRUSH)
        dc.SetPen(wx.WHITE_PEN)
        dc.DrawRectangle(8, 12, 6, 8)  # 音量喇叭
        dc.DrawPolygon([(14, 16), (20, 8), (20, 24), (14, 16)])  # 声波
        
        dc.SelectObject(wx.NullBitmap)
        icon = wx.Icon()
        icon.CopyFromBitmap(bmp)
        return icon

    def _create_tray_menu(self):
        """创建托盘菜单"""
        # 删除旧菜单项
        if hasattr(self, 'tray_menu'):
            for item in self.tray_menu.GetMenuItems():
                self.tray_menu.Remove(item.GetId())
        else:
            self.tray_menu = wx.Menu()

        # 重新创建菜单项并绑定事件
        self.menu_toggle_window = self.tray_menu.Append(
            wx.ID_ANY, 
            "隐藏主窗口" if self.IsShown() else "显示主窗口"
        )
        self.tray_menu.AppendSeparator()
        
        # 添加校准选项
        self.menu_calibrate = self.tray_menu.Append(wx.ID_ANY, "音频校准")
        self.tray_menu.AppendSeparator()
        
        # 添加状态开关
        self.menu_enabled = self.tray_menu.AppendCheckItem(wx.ID_ANY, "启用音量调节")
        self.menu_enabled.Check(True)  # 默认启用状态
        self.tray_menu.AppendSeparator()
        
        self.menu_about = self.tray_menu.Append(wx.ID_ANY, "关于")
        self.tray_menu.AppendSeparator()
        self.menu_exit = self.tray_menu.Append(wx.ID_ANY, "退出")

        # 绑定事件到TaskBarIcon实例
        self.tbicon.Bind(wx.EVT_MENU, self._on_toggle_window, self.menu_toggle_window)
        self.tbicon.Bind(wx.EVT_MENU, self.show_calibration_dialog, self.menu_calibrate)
        self.tbicon.Bind(wx.EVT_MENU, self._on_toggle_enabled, self.menu_enabled)
        self.tbicon.Bind(wx.EVT_MENU, self._on_about, self.menu_about)
        self.tbicon.Bind(wx.EVT_MENU, self._on_exit, self.menu_exit)

    def _on_toggle_window(self, event):
        """切换窗口显示状态"""
        if self.IsShown():
            self.Hide()
            self.menu_toggle_window.SetItemLabel("显示主窗口")
        else:
            self.Show()
            self.Restore()
            self.Raise()
            self.menu_toggle_window.SetItemLabel("隐藏主窗口")

    def _on_auto_adjust_toggle(self, event):
        """自动调节开关事件处理"""
        enabled = event.IsChecked()
        if enabled:
            self.audio_analyzer.start_analyzing(self.worker.on_audio_event)
            self.auto_adjust_status.SetLabel("状态: 已启用")
            self.auto_adjust_status.SetForegroundColour(wx.Colour(0, 128, 0))
            self.menu_enabled.Check(True)  # 同步更新托盘菜单状态
            self.logger.info("音量自动调节已启用")
        else:
            self.audio_analyzer.stop_analyzing()
            self.auto_adjust_status.SetLabel("状态: 已禁用")
            self.auto_adjust_status.SetForegroundColour(wx.Colour(128, 128, 128))
            self.menu_enabled.Check(False)  # 同步更新托盘菜单状态
            self.logger.info("音量自动调节已禁用")

    def _on_toggle_enabled(self, event):
        """托盘菜单中切换音量调节启用状态"""
        enabled = self.menu_enabled.IsChecked()
        # 同步更新主界面状态
        self.auto_adjust_toggle.SetValue(enabled)
        if enabled:
            self.audio_analyzer.start_analyzing(self.worker.on_audio_event)
            self.auto_adjust_status.SetLabel("状态: 已启用")
            self.auto_adjust_status.SetForegroundColour(wx.Colour(0, 128, 0))
            self.logger.info("音量自动调节已启用")
        else:
            self.audio_analyzer.stop_analyzing()
            self.auto_adjust_status.SetLabel("状态: 已禁用")
            self.auto_adjust_status.SetForegroundColour(wx.Colour(128, 128, 128))
            self.logger.info("音量自动调节已禁用")

    def _on_tray_left_click(self, event):
        """托盘左键单击"""
        # 显示简单的状态提示
        current_volume = self.volume_controller.get_volume()
        current_db = self.audio_analyzer.get_current_db()
        self.tbicon.ShowBalloon(
            "音频响度状态",
            f"系统音量: {int(current_volume * 100)}%\n"
            f"当前响度: {current_db:.1f} dB",
            2000
        )

    def _on_tray_left_dclick(self, event):
        """托盘左键双击"""
        if self.IsShown():
            self.Hide()
        else:
            self.Show()
            self.Restore()
            self.Raise()

    def _on_tray_right_click(self, event):
        """托盘右键菜单"""
        # 更新窗口显示状态的菜单项文本
        self.menu_toggle_window.SetItemLabel(
            "隐藏主窗口" if self.IsShown() else "显示主窗口"
        )
        self.tbicon.PopupMenu(self.tray_menu)

    def _on_about(self, event):
        """显示关于对话框"""
        dlg = AboutDialog(self)
        dlg.ShowModal()
        dlg.Destroy()

    def _on_exit(self, event):
        """退出程序"""
        dlg = wx.MessageDialog(self, "确定要退出程序吗？",
                             "确认", wx.YES_NO | wx.NO_DEFAULT | wx.ICON_QUESTION)
        if dlg.ShowModal() == wx.ID_YES:
            self.tbicon.Destroy()
            self.Destroy()
        dlg.Destroy()

    def _populate_device_list(self):
        """填充设备列表"""
        self.audio_devices = AudioUtilities.GetAllDevices()
        self.device_combo.Clear()
        default_index = 0
        for i, device in enumerate(self.audio_devices):
            self.device_combo.Append(device.FriendlyName, device.id)
            if device.id == self.config.device_id:
                default_index = i
        self.device_combo.SetSelection(default_index)
        self.device_combo.Bind(wx.EVT_CHOICE, self._on_device_changed)

    def _on_device_changed(self, event):
        """设备选择改变"""
        device_id = self.device_combo.GetClientData(event.GetSelection())
        self.config.update(device_id=device_id)
        self.audio_analyzer.set_device(device_id)
        self.logger.info(f"选择设备: {device_id}")

    def _on_max_db_changed(self, event):
        """最大响度设置改变"""
        self.config.update(max_db=event.GetValue())

    def _on_min_db_changed(self, event):
        """最小响度设置改变"""
        self.config.update(min_db=event.GetValue())

    def _on_threshold_changed(self, event):
        self.config.update(audio_threshold=event.GetValue())

    def _on_interval_max_changed(self, event):
        self.config.update(interval_max=event.GetValue())

    def _on_interval_min_changed(self, event):
        self.config.update(interval_min=event.GetValue())

    def _on_volume_k_changed(self, event):
        self.config.update(volume_change_k=event.GetValue())

    def _on_autostart_changed(self, event):
        autostart = event.IsChecked()
        if autostart and not self.service_manager.add_to_startup():
            wx.MessageBox("添加开机启动项失败，请以管理员权限运行程序。", "错误",
                         wx.OK | wx.ICON_ERROR)
            self.autostart_check.SetValue(False)
            return
        self.config.update(auto_start=autostart)

    def _on_minimize_changed(self, event):
        self.config.update(start_minimized=event.IsChecked())

    def _on_timer(self, event):
        """定时器事件，更新显示"""
        current_db = self.audio_analyzer.get_current_db()
        current_volume = self.volume_controller.get_volume()
        
        # 更新显示
        self.db_label.SetLabel(f"当前响度: {current_db:.1f} dB")
        self.volume_label.SetLabel(f"系统音量: {int(current_volume * 100)}%")
        
        # 更新音频状态
        if self.audio_analyzer.is_playing():
            if current_db > self.config.max_db:
                status = "响度过高"
                color = wx.Colour(255, 0, 0)
            elif current_db < self.config.min_db:
                status = "响度过低"
                color = wx.Colour(255, 165, 0)
            else:
                status = "正常"
                color = wx.Colour(0, 255, 0)
        else:
            status = "无音频"
            color = wx.Colour(128, 128, 128)
            
        self.status_label.SetLabel(f"音频状态: {status}")
        self.status_label.SetForegroundColour(color)

    def _on_reset_clicked(self, event):
        """重置按钮点击"""
        dlg = wx.MessageDialog(self, "确定要恢复默认设置吗？",
                             "确认", wx.YES_NO | wx.NO_DEFAULT | wx.ICON_QUESTION)
        
        if dlg.ShowModal() == wx.ID_YES:
            self.config.reset_to_default()
            self._update_ui_from_config()
            wx.MessageBox("已恢复默认设置。", "提示")
        dlg.Destroy()

    def _on_minimize(self, event):
        """最小化事件"""
        self.Hide()
        self.tbicon.ShowBalloon("办公室的大盾 - 音频响度均衡器",
                               "程序已最小化到系统托盘，双击图标可以重新打开。",
                               2000)

    def update_volume(self, volume):
        """更新音量显示"""
        self.volume_label.SetLabel(f"系统音量: {int(volume * 100)}%")

    def show_calibration_dialog(self, event):
        """显示校准对话框"""
        dlg = CalibrationDialog(self, self.audio_analyzer, self.volume_controller)
        if dlg.ShowModal() == wx.ID_OK:
            max_db, min_db, audio_threshold = dlg.get_calibration_results()
            self.config.update(
                max_db=max_db,
                min_db=min_db,
                audio_threshold=audio_threshold,
                was_calibrated=True
            )
            self._update_ui_from_config()
            wx.MessageBox(f"校准完成！\n\n"
                        f"最大响度: {max_db:.1f} dB\n"
                        f"最小响度: {min_db:.1f} dB\n"
                        f"音频检测阈值: {audio_threshold:.1f} dB",
                        "校准完成")
        dlg.Destroy()

    def _update_ui_from_config(self):
        """从配置更新UI"""
        self.max_db_spin.SetValue(self.config.max_db)
        self.min_db_spin.SetValue(self.config.min_db)
        self.threshold_spin.SetValue(self.config.audio_threshold)
        self.interval_max_spin.SetValue(self.config.interval_max)
        self.interval_min_spin.SetValue(self.config.interval_min)
        self.volume_k_spin.SetValue(self.config.volume_change_k)
        self.autostart_check.SetValue(self.config.auto_start)
        self.minimize_check.SetValue(self.config.start_minimized)
        
        # 更新设备选择
        for i in range(self.device_combo.GetCount()):
            if self.device_combo.GetClientData(i) == self.config.device_id:
                self.device_combo.SetSelection(i)
                break

    def on_close(self, event):
        """关闭窗口事件"""
        if event.CanVeto():
            event.Veto()
            self.Hide()
            self.tbicon.ShowBalloon(
                "办公室的大盾 - 音频响度均衡器",
                "程序已最小化到系统托盘，双击图标可以重新打开。",
                2000
            )
        else:
            self.tbicon.Destroy()
            self.Destroy()
