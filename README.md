# timing_diagram_generator
时序图系统

## 打包成单一 Windows EXE

项目提供 `scripts\build_windows_exe.bat` 用于在 Windows 上通过 PyInstaller 打包。

使用步骤：

1. 安装 Python 3.10 或更新版本，并确保 `python` 可在命令行中运行。
2. 在项目根目录运行：

   ```bat
   scripts\build_windows_exe.bat
   ```

3. 打包完成后输出文件在 `dist` 目录：

   - `dist\时序图自动生成工具.exe`
   - `dist\常用动作时间.json`

`常用动作时间.json` 是用户可编辑配置文件，需要和 exe 放在同一个目录。程序以 exe 所在目录作为配置目录；如果配置文件不存在，首次运行会在 exe 旁边生成默认配置。

如果运行 bat 时出现类似 `'锘緻echo' 不是内部或外部命令` 的错误，说明 bat 文件被保存成了带 BOM 的 UTF-8。请使用本仓库的 `scripts\build_windows_exe.bat`，或把 bat 另存为“UTF-8 无 BOM”。

如果 `.venv\Scripts\activate.bat` 不存在，通常是虚拟环境创建不完整；删除 `.venv` 后重新执行脚本即可。

## 常用动作公式编辑

在“常用动作时间”窗口中：

- “使用者”页只用于选择动作、填写参数值并应用到动作，不会修改配置。
- “编辑者”页用于新增、修改、删除动作配置。
- 点击“编辑公式/参数”会打开独立公式编辑窗口，窗口右侧固定显示“可用公式说明”。
- 在公式编辑窗口点击“保存公式”后，会校验公式和参数，并同步保存到 `常用动作时间.json`。
- 自定义公式支持参数名、`if(条件, 真值, 假值)`、`abs`、`min`、`max`、`round`、`ceil`、`floor`、`sqrt`、`pi`。
