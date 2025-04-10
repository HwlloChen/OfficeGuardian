import wx
import wx.adv


class AboutDialog(wx.Dialog):
    """关于对话框"""
    def __init__(self, parent):
        super().__init__(parent, title="关于", size=(500, 300))
        
        sizer = wx.BoxSizer(wx.VERTICAL)
        
        # 标题
        title = wx.StaticText(self, label="办公室的大盾 - 音频响度均衡器")
        title.SetFont(wx.Font(14, wx.FONTFAMILY_DEFAULT, wx.FONTSTYLE_NORMAL, wx.FONTWEIGHT_BOLD))
        sizer.Add(title, 0, wx.ALL | wx.CENTER, 10)
        
        # 版本信息
        version = wx.StaticText(self, label="版本 1.14.514")
        sizer.Add(version, 0, wx.ALL | wx.CENTER, 5)
        
        # 分割线
        line = wx.StaticLine(self)
        sizer.Add(line, 0, wx.EXPAND | wx.ALL, 5)
        
        # 描述文本
        description = wx.TextCtrl(self, style=wx.TE_MULTILINE | wx.TE_READONLY | wx.NO_BORDER)
        description.SetValue(
            "办公室的大盾是一款智能音频响度均衡器，它能自动调节系统音量，"
            "确保音频始终保持在舒适的响度范围内。\n\n"
            "主要功能：\n"
            "• 实时监控音频响度\n"
            "• 自动调节系统音量\n"
            "• 支持多音频设备\n"
            "• 自定义响度范围\n"
            "• 开机自启动\n\n"
            "作者：Chen\n"
            "版权所有 © 2024"
        )
        sizer.Add(description, 1, wx.EXPAND | wx.ALL, 10)
        
        # Github链接
        github_link = wx.adv.HyperlinkCtrl(
            self, 
            wx.ID_ANY, 
            "在 Github 上查看项目",
            "https://github.com/HwlloChen/OfficeGuardian"
        )
        sizer.Add(github_link, 0, wx.ALL | wx.CENTER, 5)
        
        # 确定按钮
        btn = wx.Button(self, wx.ID_OK, "确定")
        sizer.Add(btn, 0, wx.ALL | wx.CENTER, 10)
        
        self.SetSizer(sizer)