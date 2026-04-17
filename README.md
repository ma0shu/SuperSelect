# SuperSelect

轻量级 WPF 工具：监听系统文件对话框，在其下方弹出一行快速选择面板，支持：

- Everything 关键词搜索（多排序）
- 最近文件（按修改时间倒序）
- 文件托盘（拖拽加入、持久化）
- 当前资源管理器路径快速跳转

## 技术栈

- C# / .NET 8
- WPF
- WinEventHook
- UI Automation
- Everything SDK（`Everything64.dll`）

## 运行前置

1. 安装并启动 Everything。
2. 准备 `Everything64.dll`（Everything SDK）并确保可被应用加载：
   1. 放到可执行文件目录；或
   2. 所在目录加入系统 `PATH`。

## 构建

```powershell
dotnet build SuperSelect.slnx -v minimal
```

## 使用逻辑

- 打开任意程序的“打开/保存”文件对话框时，面板会自动出现在对话框下方。
- 单击结果：先跳转到文件所在目录，再把文件名写入“文件名”输入框。
- 双击结果或按 Enter：直接写入完整路径并触发“打开/保存”。
- 在面板上直接拖入文件可加入“托盘”。

