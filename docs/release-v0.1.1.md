# obj2pdf v0.1.1

Windows PowerPoint 对象转 PDF 插件，支持完整对象、组合和多选合并导出，自动裁切页面并提供留白设置。

下载附件 `obj2pdf-Setup.exe`，关闭 PowerPoint 后双击安装，再打开 PowerPoint 的 obj2pdf 选项卡使用。更新可直接运行安装包。

本版本增加导出错误诊断：记录 PowerPoint 拒绝的对象位置、对象名称和属性数值，并修正组合对象负线宽值可能影响边界计算的问题。

需要 Windows 桌面版 Microsoft PowerPoint 和 .NET Framework 4.8。已验证 64 位 PowerPoint 16.0；32 位尚未实机验收。安装包未签名。

已通过现有尺寸、旋转、小对象、字体及源文件保护测试。此前报告的“指定的值超出了范围”问题尚未在原问题文稿上确认解决，此版本适合试用和收集诊断信息。

源码不包含本地测试文件、测试目录及编译产物。
