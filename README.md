# 本地资金流向链条分析工具

这是一个给非技术用户使用的 Windows 桌面小工具，完全本地运行，不联网、不上传数据。软件读取 `.xlsx` 或 `.csv` 交易流水，自动分析资金流向相邻关系和链条。

## 普通用户下载使用

1. 打开 GitHub Releases 页面。
2. 下载 `fund-flow-analyst-windows.zip`。
3. 解压 zip。
4. 双击 `fund-flow-analyst.exe` 运行。

注意：不要下载 GitHub 自动显示的 `Source code (zip)` 或 `Source code (tar.gz)`，那是源码包，里面不会有 exe。

exe 版本不需要安装 Python，不需要安装第三方库，也不需要配置 pip、清华源或其他镜像源。

## 开发者本地运行

```powershell
cd excel
python fund_flow_app.py
```

当前程序只使用 Python 标准库。开发者只有在打包 exe 时才需要安装 PyInstaller。

## 开发者本地打包

双击：

```text
excel\打包成EXE_开发者用.bat
```

打包完成后，exe 会生成在：

```text
excel\dist\资金流向分析工具.exe
```

## GitHub 自动发布

这个仓库包含 GitHub Actions 发布流程。推送版本 tag 后会自动构建 Windows exe，并在 GitHub Release 中上传：

```text
fund-flow-analyst-windows.zip
```

发布新版本示例：

```powershell
git tag v下一个版本号
git push origin v下一个版本号
```
