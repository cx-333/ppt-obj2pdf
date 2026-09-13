# obj2pdf 技术方案

## 目标

用户在 PowerPoint 中选中一个完整对象或多个对象，直接保存为边界匹配的单页 PDF。保留对象原始尺寸、相对位置和堆叠关系，省去整页导出后的手工裁切。交付可双击安装的 Windows EXE。

## 技术选型

采用 C# / .NET Framework 4.8 COM 加载项，使用 `IDTExtensibility2` 接入生命周期、`IRibbonExtensibility` 提供功能区。通过 COM late binding 访问 PowerPoint，不依赖用户安装 Office PIA 或 VSTO 开发工具。

选择桌面 COM 接口是为了完整访问当前选区、未保存内容和 PowerPoint 原生 PDF 导出。首版明确面向 Windows，不引入网页宿主、云转换服务或 VBA 宏配置。

`IDTExtensibility2`、`IRibbonExtensibility` 必须使用正确的 dual 接口 ABI；生命周期 `custom` 参数显式以 `SAFEARRAY(VARIANT)*` 传递。错误的 COM 封送可能导致 Office 宿主崩溃，不能仅以 DLL 编译通过作为验收。

## 导出流程

1. 读取活动窗口中的 ShapeRange，记录当前幻灯片索引及顶层 Shape.Id；拒绝组合内部子对象。
2. 通过 SaveCopyAs 保存包含未保存编辑的临时 PPTX，不改变源文件状态。
3. 从源对象几何尺寸、旋转角度和线宽计算联合外接矩形，加上用户设置的毫米留白。
4. 只修改临时 PPTX 中 `ppt/presentation.xml` 的 `p:sldSz`。使用 12700 EMU/pt 换算，不通过 COM PageSetup 调整尺寸，避免 PowerPoint 自动缩放字体。
5. 隐藏打开临时 PPTX，删除非目标幻灯片及非选中形状，隐藏母版装饰对象；按设置使用白色或原背景填充。
6. 将保留对象平移至新画布边界，不调整宽高、字体或旋转。保留组合、图片、主题、图表关系和层叠顺序。
7. 使用 `Presentation.SaveAs(..., ppSaveAsPDF)` 生成原生 PDF。
8. 启动独立的 `Obj2Pdf.Crop.exe`，通过 PDFsharp 修改 MediaBox/CropBox/TrimBox/BleedBox/ArtBox。PowerPoint 最小页面为 72pt，PDF 裁切支持更小对象。
9. 在目标目录生成中间文件，成功后原子替换已有目标；失败时保留原 PDF。关闭临时文稿并清理文件。

PDF 裁切程序独立运行，有自己的 .NET binding redirects，避免 PDFsharp 的 Microsoft.Extensions 依赖与 Office 及其他插件冲突。裁切超时为 60 秒。Office 原生 COM 保存操作为同步调用，首版不提供中途取消。

## 安装与卸载

单个 AnyCPU Windows 安装 EXE 内嵌 ZIP，含插件、裁切程序、配置及第三方许可。安装前检查 PowerPoint 已关闭、PowerPoint 存在和 .NET Framework 4.8 已安装。

解压至 `%LOCALAPPDATA%\Obj2Pdf\app-版本-时间戳`，在 HKCU 注册 COM 类和 PowerPoint Addins 键（LoadBehavior=3），为 32/64 位视图注册；加入 Windows 卸载列表。无需管理员权限。

更新使用独立目录，成功注册后清理旧版本目录。首版未实现安装中断时的完整注册表事务回滚；重新运行安装器可以修复。卸载仅删除本产品的注册项和版本目录，保留正在执行的卸载 EXE。

## 首版边界与后续路线

| 项目 | 首版 | 后续方向 |
| --- | --- | --- |
| 多选 | 合并一页 | 每对象分别导出、批量命名 |
| 裁切 | 旋转几何边界＋可调留白 | 自动识别阴影、发光和文字溢出 |
| 质量 | 原生 PDF 内容 | 特殊字体与复杂效果兼容性矩阵 |
| 平台 | Windows PowerPoint，64 位已验收 | 32 位实机矩阵、Mac 单独方案 |
| 安装 | 当前用户、未签名 EXE | 商业签名、企业 MSI、升级回滚 |
| 交互 | 功能区导出按钮 | 右键菜单、快捷键、预览 |

## 依据

- [Microsoft：COM 插件自定义 Office 功能区](https://github.com/MicrosoftDocs/VBA-Docs/blob/main/Library-Reference/Concepts/customize-the-office-fluent-ribbon-by-using-a-managed-com-add-in.md)
- [Microsoft：PowerPoint SaveAs](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.presentation.saveas)
- [Microsoft：PowerPoint 固定格式导出](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.presentation.exportasfixedformat)
- [Microsoft：PowerPoint 页面尺寸](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.pagesetup)
- [PDFsharp：.NET Framework 支持](https://docs.pdfsharp.net/General/Overview/Dotnet-Framework-Support.html)
