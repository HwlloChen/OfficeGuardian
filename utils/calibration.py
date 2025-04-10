import logging
import numpy as np
import wx

class CalibrationDialog(wx.Dialog):
    """校准对话框，用于帮助用户设置最大和最小响度阈值"""

    def __init__(self, parent, audio_analyzer, volume_controller):
        super().__init__(parent, title="音频响度校准", size=(600, 470))
        self.logger = logging.getLogger('OfficeGuardian.Calibration')
        self.audio_analyzer = audio_analyzer
        self.volume_controller = volume_controller

        # 初始默认值
        self.max_db = -10.0
        self.min_db = -40.0
        self.audio_threshold = -60.0

        # 状态变量
        self.current_step = 0  # 0=准备, 1=校准最大响度, 2=校准最小响度, 3=校准音频阈值
        self.collected_db_values = []

        self._init_ui()
        self.start_timer()

    def _init_ui(self):
        """初始化UI"""
        main_sizer = wx.BoxSizer(wx.VERTICAL)

        # 说明文本
        self.info_text = wx.TextCtrl(self, style=wx.TE_MULTILINE | wx.TE_READONLY | wx.TE_NO_VSCROLL)
        self.info_text.SetMinSize((480, 140))
        main_sizer.Add(self.info_text, 0, wx.ALL | wx.EXPAND, 5)

        # 进度条
        self.progress = wx.Gauge(self, range=3)
        main_sizer.Add(self.progress, 0, wx.ALL | wx.EXPAND, 5)

        # 音量控制
        volume_box = wx.StaticBox(self, label="系统音量")
        volume_sizer = wx.StaticBoxSizer(volume_box, wx.VERTICAL)
        
        self.volume_slider = wx.Slider(self, value=int(self.volume_controller.get_volume() * 100),
                                     minValue=0, maxValue=100,
                                     style=wx.SL_HORIZONTAL)
        self.volume_label = wx.StaticText(self, label="当前音量: 50%")
        self.volume_slider.Bind(wx.EVT_SLIDER, self.on_volume_changed)
        
        volume_sizer.Add(self.volume_slider, 0, wx.ALL | wx.EXPAND, 5)
        volume_sizer.Add(self.volume_label, 0, wx.ALL, 5)
        main_sizer.Add(volume_sizer, 0, wx.ALL | wx.EXPAND, 5)

        # 响度显示
        db_box = wx.StaticBox(self, label="当前响度")
        db_sizer = wx.StaticBoxSizer(db_box, wx.VERTICAL)
        
        self.db_gauge = wx.Gauge(self, range=80)
        self.db_label = wx.StaticText(self, label="-60 dB")
        
        db_sizer.Add(self.db_gauge, 0, wx.ALL | wx.EXPAND, 5)
        db_sizer.Add(self.db_label, 0, wx.ALL, 5)
        main_sizer.Add(db_sizer, 0, wx.ALL | wx.EXPAND, 5)

        # 按钮
        button_sizer = wx.BoxSizer(wx.HORIZONTAL)
        self.start_button = wx.Button(self, label="开始校准")
        self.next_button = wx.Button(self, label="下一步")
        self.cancel_button = wx.Button(self, label="取消")

        self.start_button.Bind(wx.EVT_BUTTON, self.on_start)
        self.next_button.Bind(wx.EVT_BUTTON, self.on_next)
        self.cancel_button.Bind(wx.EVT_BUTTON, self.on_cancel)

        button_sizer.Add(self.cancel_button, 0, wx.ALL, 5)
        button_sizer.Add(self.start_button, 0, wx.ALL, 5)
        button_sizer.Add(self.next_button, 0, wx.ALL, 5)
        
        self.next_button.Disable()
        main_sizer.Add(button_sizer, 0, wx.ALL | wx.ALIGN_RIGHT, 5)

        self.SetSizer(main_sizer)
        self._update_step_ui()

    def start_timer(self):
        """启动定时器"""
        self.timer = wx.Timer(self)
        self.Bind(wx.EVT_TIMER, self.on_timer)
        self.timer.Start(100)  # 100ms

    def on_timer(self, event):
        """定时器事件处理"""
        current_db = self.audio_analyzer.get_current_db()
        self.db_label.SetLabel(f"{current_db:.1f} dB")
        self.db_gauge.SetValue(int(current_db + 80))  # 转换到0-80范围

        if self.current_step in [1, 2, 3] and self.audio_analyzer.is_playing():
            self.collected_db_values.append(current_db)

    def on_volume_changed(self, event):
        """音量滑块改变事件"""
        volume = self.volume_slider.GetValue() / 100.0
        self.volume_controller.set_volume(volume)
        self.volume_label.SetLabel(f"当前音量: {int(volume * 100)}%")

    def on_start(self, event):
        """开始校准"""
        self.current_step = 1
        self.progress.SetValue(1)
        self.start_button.Disable()
        self.next_button.Enable()
        self._update_step_ui()

    def on_next(self, event):
        """下一步"""
        if self.current_step == 1:
            # 完成最大响度校准
            if len(self.collected_db_values) > 0:
                self.max_db = np.percentile(self.collected_db_values, 95)
            self.collected_db_values = []
            self.current_step = 2

        elif self.current_step == 2:
            # 完成最小响度校准
            if len(self.collected_db_values) > 0:
                self.min_db = np.percentile(self.collected_db_values, 50)
            self.collected_db_values = []
            self.current_step = 3

        elif self.current_step == 3:
            # 完成音频阈值校准
            if len(self.collected_db_values) > 0:
                self.audio_threshold = np.percentile(self.collected_db_values, 5)
            self.timer.Stop()
            self.EndModal(wx.ID_OK)
            return

        self.progress.SetValue(self.current_step)
        self._update_step_ui()

    def on_cancel(self, event):
        """取消校准"""
        self.timer.Stop()
        self.EndModal(wx.ID_CANCEL)

    def _update_step_ui(self):
        """更新步骤UI显示"""
        if self.current_step == 0:
            info = ("欢迎使用音频响度均衡器校准向导。\n"
                   "本向导将帮助您设置适合您环境的音频响度阈值。\n"
                   "整个过程分为三个步骤：\n"
                   "1. 校准最大响度\n"
                   "2. 校准最小响度\n"
                   "3. 校准音频检测阈值\n\n"
                   "准备好后，请点击\"开始校准\"按钮。")
        elif self.current_step == 1:
            info = ("步骤1: 校准最大响度\n\n"
                   "请播放您通常会听的音频，并将音量调整到您认为的最大舒适音量。\n"
                   "这将帮助程序确定不应超过的最大响度。\n\n"
                   "播放音频并调整好音量后，点击\"下一步\"继续。")
        elif self.current_step == 2:
            info = ("步骤2: 校准最小响度\n\n"
                   "请播放您通常会听的音频，并将音量调整到您认为的最小可听音量。\n"
                   "这将帮助程序确定不应低于的最小响度。\n\n"
                   "播放音频并调整好音量后，点击\"下一步\"继续。")
        else:
            info = ("步骤3: 校准音频检测阈值\n\n"
                   "请播放任何音频，系统将自动检测背景噪音和实际音频的区别。\n"
                   "这将帮助程序确定是否有音频正在播放。\n\n"
                   "播放音频一段时间后，点击\"下一步\"完成校准。")

        self.info_text.SetValue(info)

    def get_calibration_results(self):
        """获取校准结果"""
        return self.max_db, self.min_db, self.audio_threshold
