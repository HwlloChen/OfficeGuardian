# 办公室的大盾 - 音频响度均衡器

这是一个自动调节系统音量以保持音频响度在特定范围内的应用程序。它通过分析电脑的音频输出响度，自动调整系统音量，使实际响度控制在预设的范围内。

**特别场景：班级电脑老师们的视频音量大小不一 / 电脑音量被同学调来调去**

## 功能特点

- 实时分析电脑音频输出响度
- 自动调整系统音量以保持音频响度在设定范围内
- 用户友好的GUI界面，可显示实时响度和音量信息
- 可最小化到系统托盘运行
- 支持开机自启动
- 首次使用时引导用户进行校准

## 系统要求

- Windows 7/8/10/11
- Python 3.7+
- 必要的Python库: PySide6, numpy, sounddevice, pycaw, comtypes, pywin32

## 安装方法

1. 克隆或下载本仓库
   ```bash
   git clone https://github.com/HwlloChen/OfficeGuardian.git
   ```
2. 安装必要的依赖:

   ```bash
   pip install -r requirements.txt
   ```

## 使用方法

运行主程序:

```bash
python main.py
```

参数选项:
- `--minimized`: 以最小化方式启动程序
- `--service`: 以服务方式启动程序

## 校准

首次启动时，程序会引导您进行校准。校准过程包括三个步骤:

1. 校准最大响度: 播放您通常会听的音频，并将音量调整到最大舒适音量
2. 校准最小响度: 播放音频，并将音量调整到最小可听音量
3. 校准音频检测阈值: 播放音频，程序将自动检测背景噪音和实际音频的区别

## 配置选项

程序提供以下配置选项:

- 最大响度: 音频不应超过的最大分贝值
- 最小响度: 音频不应低于的最小分贝值
- 音频检测阈值: 判断是否有音频播放的分贝阈值
- 开机自启动: 是否在系统启动时自动运行程序
- 启动时最小化: 是否以最小化方式启动程序

## 项目结构

- `main.py`: 程序入口点
- `audio_analyzer.py`: 音频分析模块
- `volume_controller.py`: 系统音量控制模块
- `gui.py`: PySide6 GUI界面
- `calibration.py`: 校准模块
- `service_manager.py`: Windows服务管理模块
- `config.py`: 配置管理模块
- `logger.py`: 日志记录模块

## 许可证

GNU GENERAL PUBLIC LICENSE v3 - 详情请参阅LICENSE文件