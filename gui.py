from PySide6.QtWidgets import (QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
                               QLabel, QPushButton, QSystemTrayIcon, QMenu,
                               QSlider, QCheckBox, QGroupBox, QTextEdit,
                               QMessageBox, QSpinBox, QDoubleSpinBox, QApplication,
                               QComboBox, QSplitter)  # 添加 QTextEdit, QSplitter
from PySide6.QtCore import Qt, QTimer, Signal, Slot, QSize
from PySide6.QtGui import QIcon, QPixmap, QFont, QAction, QTextCharFormat, QColor, QTextCursor
import logging
from calibration import CalibrationDialog
from pycaw.pycaw import AudioUtilities


class LogHandler(logging.Handler):
    """自定义日志处理器，支持颜色显示"""

    def __init__(self, text_edit):
        super().__init__()
        self.text_edit = text_edit
        self.format = logging.Formatter(
            '%(asctime)s - %(levelname)s - %(message)s')

        # 定义不同级别的日志颜色
        self.colors = {
            logging.DEBUG: QColor(128, 128, 128),  # 灰色
            logging.INFO: QColor(0, 174, 255),    # 蓝色
            logging.WARNING: QColor(255, 165, 0),  # 橙色
            logging.ERROR: QColor(255, 0, 0),     # 红色
            logging.CRITICAL: QColor(139, 0, 0)   # 深红色
        }

    def emit(self, record):
        msg = self.format.format(record)
        cursor = self.text_edit.textCursor()
        fmt = QTextCharFormat()
        fmt.setForeground(self.colors.get(record.levelno, QColor(0, 0, 0)))

        # 插入日志文本
        cursor.movePosition(QTextCursor.End)
        cursor.insertText(msg + '\n', fmt)

        # 插入间隔
        spacer_fmt = QTextCharFormat()
        spacer_fmt.setFontPointSize(8)  # 设置间隔的字体大小
        cursor.insertText('\n', spacer_fmt)

        # 滚动到最新日志
        self.text_edit.verticalScrollBar().setValue(
            self.text_edit.verticalScrollBar().maximum())


class MainWindow(QMainWindow):
    """主窗口"""

    update_volume = Signal(float, float)  # 要更新的音量信号

    def __init__(self, audio_analyzer, volume_controller, config, service_manager):
        super().__init__()
        self.logger = logging.getLogger('OfficeGuardian.GUI')
        self.audio_analyzer = audio_analyzer
        self.volume_controller = volume_controller
        self.config = config
        self.service_manager = service_manager
        self.audio_devices = []  # 用于存储音频设备列表

        # 设置日志处理器
        self.log_text_edit = QTextEdit()
        self.log_text_edit.setReadOnly(True)
        self.log_handler = LogHandler(self.log_text_edit)
        self.log_handler.setFormatter(
            logging.Formatter('%(asctime)s - %(levelname)s - %(message)s'))
        logging.getLogger('OfficeGuardian').addHandler(self.log_handler)

        self._init_ui()
        self._create_tray_icon()

        # 定时器，用于更新UI显示
        self.update_timer = QTimer(self)
        self.update_timer.timeout.connect(self._update_display)
        self.update_timer.start(100)  # 每100ms更新一次

        # 检查是否需要首次校准
        if not self.config.was_calibrated:
            QTimer.singleShot(500, self.show_calibration_dialog)

    def _init_ui(self):
        """初始化UI"""
        self.setWindowTitle("办公室的大盾 - 音频响度均衡器")
        self.setMinimumSize(800, 500)  # 增加窗口大小以适应日志显示区域

        # 创建主分割器
        main_splitter = QSplitter(Qt.Horizontal)
        self.setCentralWidget(main_splitter)

        # 左侧控制面板
        left_widget = QWidget()
        left_layout = QVBoxLayout(left_widget)

        # --- 音频设备选择 ---
        device_group = QGroupBox("音频设备选择")
        device_layout = QHBoxLayout(device_group)

        device_label = QLabel("选择设备:")
        self.device_combo = QComboBox()
        self._populate_device_list()  # 填充设备列表
        self.device_combo.currentIndexChanged.connect(self._on_device_changed)

        device_layout.addWidget(device_label)
        device_layout.addWidget(self.device_combo)
        left_layout.addWidget(device_group)

        # --- 音频状态组 ---
        status_group = QGroupBox("音频状态")
        status_layout = QVBoxLayout(status_group)

        # 当前响度
        db_layout = QHBoxLayout()
        db_label = QLabel("当前响度:")
        self.db_value_label = QLabel("-60.0 dB")
        self.db_value_label.setMinimumWidth(80)
        db_layout.addWidget(db_label)
        db_layout.addWidget(self.db_value_label)
        status_layout.addLayout(db_layout)

        # 系统音量
        volume_layout = QHBoxLayout()
        volume_label = QLabel("系统音量:")
        self.volume_value_label = QLabel("50%")
        self.volume_value_label.setMinimumWidth(80)
        volume_layout.addWidget(volume_label)
        volume_layout.addWidget(self.volume_value_label)
        status_layout.addLayout(volume_layout)

        # 音频状态
        audio_status_layout = QHBoxLayout()
        audio_status_label = QLabel("音频状态:")
        self.audio_status_value_label = QLabel("无音频")
        self.audio_status_value_label.setMinimumWidth(80)
        audio_status_layout.addWidget(audio_status_label)
        audio_status_layout.addWidget(self.audio_status_value_label)
        status_layout.addLayout(audio_status_layout)

        left_layout.addWidget(status_group)

        # --- 参数设置组 ---
        settings_group = QGroupBox("参数设置")
        settings_layout = QVBoxLayout(settings_group)

        # 最大响度设置
        max_db_layout = QHBoxLayout()
        max_db_label = QLabel("最大响度:")
        self.max_db_spin = QDoubleSpinBox()
        self.max_db_spin.setRange(-60.0, 0.0)
        self.max_db_spin.setValue(self.config.max_db)
        self.max_db_spin.setSingleStep(1.0)
        self.max_db_spin.valueChanged.connect(self._on_max_db_changed)
        max_db_layout.addWidget(max_db_label)
        max_db_layout.addWidget(self.max_db_spin)
        settings_layout.addLayout(max_db_layout)

        # 最小响度设置
        min_db_layout = QHBoxLayout()
        min_db_label = QLabel("最小响度:")
        self.min_db_spin = QDoubleSpinBox()
        self.min_db_spin.setRange(-80.0, -10.0)
        self.min_db_spin.setValue(self.config.min_db)
        self.min_db_spin.setSingleStep(1.0)
        self.min_db_spin.valueChanged.connect(self._on_min_db_changed)
        min_db_layout.addWidget(min_db_label)
        min_db_layout.addWidget(self.min_db_spin)
        settings_layout.addLayout(min_db_layout)

        # 音频检测阈值设置
        threshold_layout = QHBoxLayout()
        threshold_label = QLabel("音频检测阈值:")
        self.threshold_spin = QDoubleSpinBox()
        self.threshold_spin.setRange(-90.0, -30.0)
        self.threshold_spin.setValue(self.config.audio_threshold)
        self.threshold_spin.setSingleStep(1.0)
        self.threshold_spin.valueChanged.connect(self._on_threshold_changed)
        threshold_layout.addWidget(threshold_label)
        threshold_layout.addWidget(self.threshold_spin)
        settings_layout.addLayout(threshold_layout)

        # 音量过大调整间隔
        interval_max_layout = QHBoxLayout()
        interval_max_label = QLabel("音量过大调整间隔(秒):")
        self.interval_max_spin = QDoubleSpinBox()
        self.interval_max_spin.setRange(0.5, 10.0)
        self.interval_max_spin.setValue(self.config.interval_max)
        self.interval_max_spin.setSingleStep(0.5)
        self.interval_max_spin.valueChanged.connect(
            self._on_interval_max_changed)
        interval_max_layout.addWidget(interval_max_label)
        interval_max_layout.addWidget(self.interval_max_spin)
        settings_layout.addLayout(interval_max_layout)

        # 音量过小调整间隔
        interval_min_layout = QHBoxLayout()
        interval_min_label = QLabel("音量过小调整间隔(秒):")
        self.interval_min_spin = QDoubleSpinBox()
        self.interval_min_spin.setRange(1.0, 30.0)
        self.interval_min_spin.setValue(self.config.interval_min)
        self.interval_min_spin.setSingleStep(1.0)
        self.interval_min_spin.valueChanged.connect(
            self._on_interval_min_changed)
        interval_min_layout.addWidget(interval_min_label)
        interval_min_layout.addWidget(self.interval_min_spin)
        settings_layout.addLayout(interval_min_layout)

        # 音量调整系数
        volume_change_k_layout = QHBoxLayout()
        volume_change_k_label = QLabel("渐进式音量调整系数k:")
        self.volume_change_k_spin = QDoubleSpinBox()
        self.volume_change_k_spin.setRange(0.1, 0.4)
        self.volume_change_k_spin.setValue(self.config.volume_change_k)
        self.volume_change_k_spin.setSingleStep(0.02)
        self.volume_change_k_spin.valueChanged.connect(
            self._on_volume_change_k_changed)
        volume_change_k_layout.addWidget(volume_change_k_label)
        volume_change_k_layout.addWidget(self.volume_change_k_spin)
        settings_layout.addLayout(volume_change_k_layout)

        # 开机自启动
        autostart_layout = QHBoxLayout()
        self.autostart_checkbox = QCheckBox("开机自启动")
        self.autostart_checkbox.setChecked(self.config.auto_start)
        self.autostart_checkbox.stateChanged.connect(
            self._on_autostart_changed)
        autostart_layout.addWidget(self.autostart_checkbox)
        settings_layout.addLayout(autostart_layout)

        # 启动时最小化
        minimize_layout = QHBoxLayout()
        self.minimize_checkbox = QCheckBox("启动时最小化")
        self.minimize_checkbox.setChecked(self.config.start_minimized)
        self.minimize_checkbox.stateChanged.connect(self._on_minimize_changed)
        minimize_layout.addWidget(self.minimize_checkbox)
        settings_layout.addLayout(minimize_layout)

        left_layout.addWidget(settings_group)

        # --- 按钮组 ---
        button_layout = QHBoxLayout()

        # 校准按钮
        self.calibrate_button = QPushButton("重新校准")
        self.calibrate_button.clicked.connect(self.show_calibration_dialog)
        button_layout.addWidget(self.calibrate_button)

        # 重置按钮
        self.reset_button = QPushButton("恢复默认设置")
        self.reset_button.clicked.connect(self._on_reset_clicked)
        button_layout.addWidget(self.reset_button)

        left_layout.addLayout(button_layout)

        # 右侧日志显示区域
        right_widget = QWidget()
        right_layout = QVBoxLayout(right_widget)

        log_group = QGroupBox("运行日志")
        log_layout = QVBoxLayout(log_group)
        log_layout.addWidget(self.log_text_edit)
        right_layout.addWidget(log_group)

        # 将左右两侧添加到分割器
        main_splitter.addWidget(left_widget)
        main_splitter.addWidget(right_widget)

        # 设置分割器比例
        main_splitter.setStretchFactor(0, 3)  # 左侧占比
        main_splitter.setStretchFactor(1, 2)  # 右侧占比

        # 状态栏
        self.statusBar().showMessage("音频响度均衡器已就绪")

    def _create_tray_icon(self):
        """创建系统托盘图标"""
        # 尝试创建一个简单的图标
        icon = QIcon()
        icon_pixmap = QPixmap(32, 32)
        icon_pixmap.fill(Qt.blue)  # 简单的蓝色图标
        icon.addPixmap(icon_pixmap)

        self.tray_icon = QSystemTrayIcon(self)
        self.tray_icon.setIcon(icon)
        self.tray_icon.setToolTip("音频响度均衡器")

        # 创建托盘菜单
        tray_menu = QMenu()

        show_action = QAction("显示", self)
        show_action.triggered.connect(self.show)
        tray_menu.addAction(show_action)

        hide_action = QAction("隐藏", self)
        hide_action.triggered.connect(self.hide)
        tray_menu.addAction(hide_action)

        tray_menu.addSeparator()

        exit_action = QAction("退出", self)
        exit_action.triggered.connect(self._on_exit_clicked)
        tray_menu.addAction(exit_action)

        self.tray_icon.setContextMenu(tray_menu)
        self.tray_icon.show()

        # 托盘图标单击事件
        self.tray_icon.activated.connect(self._on_tray_activated)

    def _update_display(self):
        """更新显示数据"""
        # 更新响度显示
        current_db = self.audio_analyzer.get_current_db()
        self.db_value_label.setText(f"{current_db:.1f} dB")

        # 更新音量显示
        current_volume = self.volume_controller.get_volume()
        self.volume_value_label.setText(f"{int(current_volume * 100)}%")

        # 更新音频状态
        if self.audio_analyzer.is_playing():
            if current_db > self.config.max_db:
                self.audio_status_value_label.setText("响度过高")
                self.audio_status_value_label.setStyleSheet("color: red;")
            elif current_db < self.config.min_db:
                self.audio_status_value_label.setText("响度过低")
                self.audio_status_value_label.setStyleSheet("color: orange;")
            else:
                self.audio_status_value_label.setText("正常")
                self.audio_status_value_label.setStyleSheet("color: green;")
        else:
            self.audio_status_value_label.setText("无音频")
            self.audio_status_value_label.setStyleSheet("")

    def _populate_device_list(self):
        """填充设备列表"""
        self.audio_devices = AudioUtilities.GetAllDevices()
        self.device_combo.clear()
        default_index = 0
        for i, device in enumerate(self.audio_devices):
            self.device_combo.addItem(device.FriendlyName, device.id)
            if device.id == self.config.device_id:
                default_index = i
        self.device_combo.setCurrentIndex(default_index)

    def _on_device_changed(self, index):
        """设备选择改变"""
        device_id = self.device_combo.itemData(index)
        self.config.update(device_id=device_id)
        self.audio_analyzer.set_device(device_id)
        self.logger.info(f"选择设备: {device_id}")

    def _on_max_db_changed(self, value):
        """最大响度设置改变"""
        self.config.update(max_db=value)

    def _on_min_db_changed(self, value):
        """最小响度设置改变"""
        self.config.update(min_db=value)

    def _on_threshold_changed(self, value):
        """音频检测阈值设置改变"""
        self.config.update(audio_threshold=value)

    def _on_interval_max_changed(self, value):
        """音量过大调整间隔改变"""
        self.config.update(interval_max=value)

    def _on_interval_min_changed(self, value):
        """音量过小调整间隔改变"""
        self.config.update(interval_min=value)

    def _on_volume_change_k_changed(self, value):
        """音量调整系数改变"""
        self.config.update(volume_change_k=value)

    def _on_autostart_changed(self, state):
        """开机自启动设置改变"""
        autostart = self.autostart_checkbox.isChecked()
        self.config.update(auto_start=autostart)

        if autostart:
            if not self.service_manager.add_to_startup():
                QMessageBox.warning(self, "警告", "添加开机启动项失败，请以管理员权限运行程序。")
                self.autostart_checkbox.setChecked(False)
        else:
            self.service_manager.remove_from_startup()

    def _on_minimize_changed(self, state):
        """启动时最小化设置改变"""
        self.config.update(start_minimized=self.minimize_checkbox.isChecked())

    def _on_reset_clicked(self):
        """重置按钮点击"""
        reply = QMessageBox.question(
            self, "确认", "确定要恢复默认设置吗？",
            QMessageBox.Yes | QMessageBox.No, QMessageBox.No
        )

        if reply == QMessageBox.Yes:
            self.config.reset_to_default()
            self._update_ui_from_config()
            QMessageBox.information(self, "提示", "已恢复默认设置。")

    def _on_exit_clicked(self):
        """退出按钮点击"""
        reply = QMessageBox.question(
            self, "确认", "确定要退出程序吗？",
            QMessageBox.Yes | QMessageBox.No, QMessageBox.No
        )

        if reply == QMessageBox.Yes:
            QApplication.quit()

    def _on_tray_activated(self, reason):
        """托盘图标激活"""
        if reason == QSystemTrayIcon.ActivationReason.Trigger:
            if self.isVisible():
                self.hide()
            else:
                self.show()
                self.activateWindow()

    def _update_ui_from_config(self):
        """从配置更新UI"""
        self.max_db_spin.setValue(self.config.max_db)
        self.min_db_spin.setValue(self.config.min_db)
        self.threshold_spin.setValue(self.config.audio_threshold)
        self.autostart_checkbox.setChecked(self.config.auto_start)
        self.minimize_checkbox.setChecked(self.config.start_minimized)
        self.interval_max_spin.setValue(self.config.interval_max)
        self.interval_min_spin.setValue(self.config.interval_min)
        self.volume_change_k_spin.setValue(self.config.volume_change_k)
        # 更新设备选择
        for i in range(self.device_combo.count()):
            if self.device_combo.itemData(i) == self.config.device_id:
                self.device_combo.setCurrentIndex(i)
                break

    def show_calibration_dialog(self):
        """显示校准对话框"""
        dialog = CalibrationDialog(
            self.audio_analyzer, self.volume_controller, self)
        dialog.calibration_complete.connect(self._on_calibration_complete)
        dialog.exec()

    @Slot(float, float, float)
    def _on_calibration_complete(self, max_db, min_db, audio_threshold):
        """校准完成回调"""
        self.config.update(
            max_db=max_db,
            min_db=min_db,
            audio_threshold=audio_threshold,
            was_calibrated=True
        )
        self._update_ui_from_config()
        QMessageBox.information(self, "校准完成",
                                f"校准完成！\n\n"
                                f"最大响度: {max_db:.1f} dB\n"
                                f"最小响度: {min_db:.1f} dB\n"
                                f"音频检测阈值: {audio_threshold:.1f} dB")

    def closeEvent(self, event):
        """关闭窗口事件"""
        # 移除日志处理器
        logging.getLogger('OfficeGuardian').removeHandler(self.log_handler)
        event.ignore()
        self.hide()
        self.tray_icon.showMessage(
            "办公室的大盾 - 音频响度均衡器",
            "程序已最小化到系统托盘，双击图标可以重新打开。",
            QSystemTrayIcon.Information,
            2000
        )
