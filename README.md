# obj2pdf

在 Windows 桌面版 PowerPoint 中，将选中的完整对象导出为独立 PDF。多个对象合并为一页，保留相对位置，PDF 页面按对象边界自动裁切。

**选中需要的图 → 点击导出 → 得到尺寸匹配的 PDF，无需再裁切整页幻灯片。**

[下载安装包](https://github.com/cx-333/ppt-obj2pdf/releases/tag/v0.1.1) · [安装与使用](#安装与使用) · [卸载教程](#卸载教程)

## 这个项目能做什么？

在 PowerPoint 中画好一张流程图、模型框图或排好一组图文后，通常只想把这部分放进论文、报告或其他文档。obj2pdf 可以直接导出选中的内容，避免先导出整页 PPT，再手工裁掉大片空白。

| 你选中的内容 | 得到的结果 |
| --- | --- |
| 一个形状、图片或完整文本框 | 以对象边界加留白为页面大小的单页 PDF |
| 一个包含多个元素的组合 | 整个组合导出，保留内部布局与层叠关系 |
| 同一页上的多个对象 | 合并到一页 PDF，保留对象之间的相对位置和间距 |
| 带旋转的对象或很小的对象 | 根据旋转后的外接边界调整 PDF 页面尺寸 |

支持设置四周留白；默认白底，也可保留幻灯片背景。文字和形状采用 PowerPoint 原生 PDF 渲染，能保留的文本和矢量不会先转成整张截图。

## 快照与导出示例

### 示例一：只导出左侧组合，不带上右侧形状

下面是示例幻灯片的真实内容快照：左侧的蓝色矩形与橙色三角形是一个组合，右侧还有一个独立的圆角矩形。

![导出前：幻灯片中有左侧组合和右侧独立圆角矩形](docs/images/source-slide.png)

**选中左侧整个组合，再点击「obj2pdf → 选中对象转 PDF」。** 结果只包含这个组合，右侧圆角矩形和整页的大片空白不会进入 PDF：

<img src="docs/images/group-export.png" alt="导出后：PDF 只包含蓝色矩形与橙色三角形组成的完整组合" width="555">

### 示例二：多个对象合并导出，保留文字与间距

同时选中矩形与文本框，可以得到一个包含两者的 PDF，无需先将它们组合。下面是实际导出结果，文本仍保留为 PDF 文本，字号检查为 18pt。

![多选导出：矩形和文本框保留相对位置、间距与文字大小](docs/images/multiple-export.png)

以上图片是文档展示用 PNG：第一张由 PowerPoint 导出幻灯片内容，后两张来自实际生成的 PDF 页面渲染，并非功能区界面截图。图片为了适配 README 显示而缩放，不代表 PDF 的物理尺寸；插件实际输出的是 PDF 文件。

## 安装与使用

1. 保存文稿并关闭所有 PowerPoint 窗口。
2. 从本仓库的 **Releases** 页面下载 `obj2pdf-Setup.exe`，双击并确认安装。无需下载源码，无需 Python、开发工具或 VBA 宏权限。
3. 打开 PowerPoint 和 PPTX，选择一个对象、完整组合，或按 Ctrl 多选。
4. 点击 **obj2pdf → 选中对象转 PDF**，选择保存位置。
5. 在 **导出设置** 中调整四周留白，或选择保留幻灯片背景。默认白底、0.7 毫米留白。

导出包含尚未保存的编辑；不会保存或修改原文稿，也不占用剪贴板。成功后显示输出位置。

安装目标是当前用户的 `%LOCALAPPDATA%\Obj2Pdf`。再次运行安装包可更新插件。

## 卸载教程

卸载前，请保存文稿并关闭所有 PowerPoint 窗口。使用安装插件时的 Windows 用户账户操作，无需管理员权限。

### 方法一：通过 Windows 设置卸载

1. 打开 Windows **设置 → 应用 → 已安装的应用**（Windows 10 中为“应用和功能”）。
2. 搜索 **obj2pdf**，找到 **obj2pdf - PowerPoint 对象转 PDF**。
3. 点击该应用的 **卸载**，在弹出的“卸载 obj2pdf 插件？”窗口中点击 **确定**。
4. 看到“已卸载插件。”后，重新打开 PowerPoint，确认 **obj2pdf** 选项卡已消失。

### 方法二：使用安装程序卸载

在下载的安装程序所在目录打开 PowerShell，执行：

```powershell
.\obj2pdf-Setup.exe /uninstall
```

如果原安装包已删除，也可以在 PowerShell 中运行安装目录里的程序：

```powershell
& "$env:LOCALAPPDATA\Obj2Pdf\Uninstall.exe" /uninstall
```

在确认窗口点击 **确定**，等待卸载完成。必须带上 `/uninstall` 参数；直接双击 `Uninstall.exe` 会进入安装流程。

### 卸载后的文件与常见问题

- 卸载会移除插件注册项、程序版本目录及导出设置，不会删除你的 PPT 文稿或已导出的 PDF。
- 安装目录可能留下 `Uninstall.exe` 和 `startup.log`。卸载程序退出后，如不需要保留日志，可在文件资源管理器地址栏输入 `%LOCALAPPDATA%\Obj2Pdf`，确认目录内容后手动删除残留文件。
- 如果提示 PowerPoint 仍在运行，请确认文稿已保存并完全退出 PowerPoint，再重试；必要时在任务管理器中检查是否仍有 `POWERPNT.EXE` 进程。
- 仅在 PowerPoint 的“COM 加载项”中取消勾选属于停用，不会卸载程序。需要彻底卸载时，请使用以上任一方法。

## 兼容性与范围

- 需要 Windows 桌面版 Microsoft PowerPoint 和 .NET Framework 4.8。
- 当前已在本机 PowerPoint 16.0、64 位 Office 实测。安装器注册 32/64 位 Office，32 位 Office 尚未实机验收。
- 不支持 PowerPoint 网页版、Mac 版或 WPS。
- 支持完整形状、文本框、组合及多选。组合内部的子对象选择会提示改选整个组合。
- 图片、图表、表格等走 PowerPoint 原生渲染；媒体仅能导出静态显示内容。
- 按旋转后的几何外接矩形及顶层线宽计算边界，不做像素级去白边。阴影、发光、箭头、组合内粗线或文字溢出时应增加留白。
- 文本框内部空白属于对象本身，会保留。对象间距同样会保留。
- “保留背景”保留背景填充，不包含母版装饰形状；背景图案会按新的画布尺寸布局。
- PDF 保留 PowerPoint 能输出的文字和矢量；特殊字体、艺术字和效果是否栅格化取决于 Office。
- 首版安装包未做商业代码签名。企业环境要求签名加载项时，需要签名和管理员部署流程。

## 开发

需要 Windows、.NET SDK（本机使用 9.0.203）。首次构建通过 NuGet 获取参考程序集和 PDFsharp 6.2.4。

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\build.ps1
```

输出：`dist/obj2pdf-Setup.exe`。依赖版本及内容哈希记录在各项目的 `packages.lock.json` 中。

测试脚本、测试文稿及测试输出仅保留在本地，不包含在此仓库中。编译产物不提交到源码仓库；安装 EXE 通过 Releases 分发。

## 文件说明

| 文件 | 用途 |
| --- | --- |
| `src/Obj2Pdf` | COM 插件、功能区、设置、临时文稿与对象导出 |
| `src/Crop` | 独立 PDF 裁切程序，隔离 PDF 库依赖 |
| `src/Setup` | 当前用户安装、更新、卸载 |
| `build.ps1` | 构建并打包单文件安装器 |
| `docs/technical-plan.md` | 技术方案 |
| `docs/images` | README 使用的幻灯片快照与导出效果图 |

## 排错

导出错误记录在 `%LOCALAPPDATA%\Obj2Pdf\last-error.log`，启动记录在同目录 `startup.log`；安装失败记录在 `%TEMP%\obj2pdf-setup-error.log`。

找不到选项卡时，先重启 PowerPoint，再检查“文件 → 选项 → 加载项 → COM 加载项”中是否有 obj2pdf。若 Office 或组织策略禁用了加载项，应按其提示处理。

受保护视图、受权限管理限制的文稿、无法复制的对象或损坏的文件可能无法导出。已有同名 PDF 被其他软件锁定时，关闭占用后再试。
