import time
import logging
import numpy as np
from PySide6.QtWidgets import (QDialog, QVBoxLayout, QLabel, QPushButton,
                               QSlider, QHBoxLayout, QProgressBar)
from PySide6.QtCore import Qt, QTimer, Signal, Slot


class CalibrationDialog(QDialog):
    """校准对话框，用于帮助用户设置最大和最小响度阈值"""

    calibration_complete = Signal(float, float, float)

    def __init__(self, audio_analyzer, volume_controller, parent=None):
        super().__init__(parent)
        self.logger = logging.getLogger('AudioEqualizer.Calibration')
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

    def _init_ui(self):
        """初始化UI"""
        self.setWindowTitle("音频响度校准")
        self.setMinimumSize(500, 400)

        # 主布局
        layout = QVBoxLayout(self)

        # 说明标签
        self.info_label = QLabel(
            "欢迎使用音频响度均衡器校准向导。\n"
            "本向导将帮助您设置适合您环境的音频响度阈值。\n"
            "整个过程分为三个步骤：\n"
            "1. 校准最大响度\n"
            "2. 校准最小响度\n"
            "3. 校准音频检测阈值\n\n"
            "准备好后，请点击\"开始校准\"按钮。"
        )
        self.info_label.setWordWrap(True)
        layout.addWidget(self.info_label)

        # 进度条
        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 3)
        self.progress_bar.setValue(0)
        layout.addWidget(self.progress_bar)

        # 当前音量标签
        self.volume_layout = QHBoxLayout()
        self.volume_label = QLabel("系统音量:")
        self.volume_slider = QSlider(Qt.Horizontal)
        self.volume_slider.setRange(0, 100)
        self.volume_slider.setValue(
            int(self.volume_controller.get_volume() * 100))
        self.volume_slider.valueChanged.connect(self._on_volume_changed)
        self.volume_value_label = QLabel(f"{self.volume_slider.value()}%")

        self.volume_layout.addWidget(self.volume_label)
        self.volume_layout.addWidget(self.volume_slider)
        self.volume_layout.addWidget(self.volume_value_label)
        layout.addLayout(self.volume_layout)

        # 当前分贝值显示
        self.db_layout = QHBoxLayout()
        self.db_label = QLabel("当前响度:")
        self.db_value_label = QLabel("-60 dB")
        self.db_bar = QProgressBar()
        self.db_bar.setRange(-80, 0)
        self.db_bar.setValue(-60)

        self.db_layout.addWidget(self.db_label)
        self.db_layout.addWidget(self.db_bar)
        self.db_layout.addWidget(self.db_value_label)
        layout.addLayout(self.db_layout)

        # 按钮
        self.button_layout = QHBoxLayout()
        self.start_button = QPushButton("开始校准")
        self.start_button.clicked.connect(self._start_calibration)
        self.next_button = QPushButton("下一步")
        self.next_button.clicked.connect(self._next_step)
        self.next_button.setEnabled(False)
        self.cancel_button = QPushButton("取消")
        self.cancel_button.clicked.connect(self.reject)

        self.button_layout.addWidget(self.cancel_button)
        self.button_layout.addWidget(self.start_button)
        self.button_layout.addWidget(self.next_button)
        layout.addLayout(self.button_layout)

        # 设置定时器更新分贝显示
        self.update_timer = QTimer(self)
        self.update_timer.timeout.connect(self._update_db_display)
        self.update_timer.start(100)  # 每100ms更新一次

    def _update_db_display(self):
        """更新分贝显示"""
        current_db = self.audio_analyzer.get_current_db()
        self.db_value_label.setText(f"{current_db:.1f} dB")
        self.db_bar.setValue(int(current_db))

        # 收集校准数据
        if self.current_step in [1, 2, 3] and self.audio_analyzer.is_playing():
            self.collected_db_values.append(current_db)

    def _on_volume_changed(self, value):
        """音量滑块改变事件"""
        self.volume_controller.set_volume(value / 100.0)
        self.volume_value_label.setText(f"{value}%")

    def _start_calibration(self):
        """开始校准"""
        self.current_step = 1
        self.progress_bar.setValue(1)
        self.start_button.setEnabled(False)
        self.next_button.setEnabled(True)
        self._update_step_ui()

    def _next_step(self):
        """进入下一步"""
        if self.current_step == 1:
            # 完成最大响度校准
            if len(self.collected_db_values) > 0:
                # 使用95百分位作为最大响度
                self.max_db = np.percentile(self.collected_db_values, 95)
                self.logger.info(f"最大响度校准: {self.max_db:.2f} dB")
            self.collected_db_values = []
            self.current_step = 2

        elif self.current_step == 2:
            # 完成最小响度校准
            if len(self.collected_db_values) > 0:
                # 使用中位数作为最小响度
                self.min_db = np.percentile(self.collected_db_values, 50)
                self.logger.info(f"最小响度校准: {self.min_db:.2f} dB")
            self.collected_db_values = []
            self.current_step = 3

        elif self.current_step == 3:
            # 完成音频阈值校准
            if len(self.collected_db_values) > 0:
                # 使用5百分位作为音频检测阈值
                self.audio_threshold = np.percentile(
                    self.collected_db_values, 5)
                self.logger.info(f"音频检测阈值校准: {self.audio_threshold:.2f} dB")

            # 完成全部校准
            self.calibration_complete.emit(
                self.max_db, self.min_db, self.audio_threshold)
            self.accept()
            return

        # 更新UI显示
        self.progress_bar.setValue(self.current_step)
        self._update_step_ui()

    def _update_step_ui(self):
        """更新当前步骤的UI显示"""
        if self.current_step == 1:
            self.info_label.setText(
                "步骤1: 校准最大响度\n\n"
                "请播放您通常会听的音频，并将音量调整到您认为的最大舒适音量。\n"
                "这将帮助程序确定不应超过的最大响度。\n\n"
                "播放音频并调整好音量后，点击\"下一步\"继续。"
            )
        elif self.current_step == 2:
            self.info_label.setText(
                "步骤2: 校准最小响度\n\n"
                "请播放您通常会听的音频，并将音量调整到您认为的最小可听音量。\n"
                "这将帮助程序确定不应低于的最小响度。\n\n"
                "播放音频并调整好音量后，点击\"下一步\"继续。"
            )
        elif self.current_step == 3:
            self.info_label.setText(
                "步骤3: 校准音频检测阈值\n\n"
                "请播放任何音频，系统将自动检测背景噪音和实际音频的区别。\n"
                "这将帮助程序确定是否有音频正在播放。\n\n"
                "播放音频一段时间后，点击\"下一步\"完成校准。"
            )

    def closeEvent(self, event):
        """关闭窗口事件"""
        self.update_timer.stop()
        event.accept()
