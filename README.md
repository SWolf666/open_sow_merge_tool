# SOW Excel Merge Tool

Windows Excel 合并与多分支 SVN 提交工具，最终交付物为
`sow_merge_tool.exe`。

- [工程说明](工程说明.md)：项目结构、开发、测试、构建和发布约定。
- [使用说明](使用说明.md)：安装、右键菜单、Excel 合并和多分支提交操作。

开发直接在 `master` 进行，提交说明使用中文，不创建 Pull Request。

快速开始：

```powershell
.\tools\bootstrap.ps1
.\tools\test.ps1 -Profile Fast
.\tools\test.ps1 -Profile Integration
.\tools\release.ps1 -DeployPath 'C:\sow_main\excel\excel_merge_tool'
```

Fast/Integration/Adversarial/Full 默认不打开可视窗口；需要 Tk/Win32 截图和 DPI 抽检时显式运行 `.\tools\test.ps1 -Profile Native` 或 `.\tools\test.ps1 -Profile Visual`。

Feishu 需求文档和真实业务配置在验证期间保持只读。
