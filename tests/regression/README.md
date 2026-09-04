# 回归脚本

本目录保存 Excel 合并、SVN 提交和 Native 工作台回归脚本。它们不是用户运行时入口，统一由仓库根目录的 `tools/test.ps1` 调度。

常用命令：

```powershell
.\tools\test.ps1 -Profile Fast
.\tools\test.ps1 -Profile Full
.\tools\test.ps1 -Profile Integration
.\tools\test.ps1 -Profile Adversarial
.\tools\test.ps1 -Profile Native

日常门禁（Fast、Integration、Adversarial、Full）默认只运行无界面 harness、
pytest 和 headless SVN 仓库，不创建真实 Tk/比较器窗口。需要实际渲染时显式运行：

```powershell
.\tools\test.ps1 -Profile Native   # 关键 10%–20%：启动中心、配对列表、DPI 和工作台抽检
.\tools\test.ps1 -Profile Visual   # 关键 Tk/Win32 可视抽检，运行前确认没有业务实例
```

Native/Visual 只使用临时工作簿/工作副本，并在启动前检查已有
`sow_merge_tool.exe`；测试脚本不会强杀业务进程。无界面会话测试通过
`ComparisonSessionManager`/`ComparisonListModel` 注入假的 Popen 和状态轮询，覆盖
双击、单活动子进程、启动失败、关闭列表不杀子进程和会话状态迁移。
历史 UI smoke（文件名以 `_smoke_test_` 且直接构造 `SowMergeApp`）已标记为
Visual-only，不由默认 profile 调度；需要逐项截图时可单独执行对应脚本。
```

`Fast` 只运行关键烟测；`Full` 才运行本目录全部 43 个烟测。真实 SVN 仓库测试位于 `tests/integration`，由 `Integration`、`Full` 和 `Adversarial` 调度；不要把它重新加入本目录造成重复执行。
