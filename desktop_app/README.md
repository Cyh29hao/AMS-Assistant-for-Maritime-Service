# AMS Assistant Desktop

这一层是给普通用户准备的桌面应用壳子，目标是把 req1、req2、req3 收到一个本地 GUI 里。

## 主要职责

- 提供图形界面
- 固定工作区结构
- 保存用户设置
- 让设置和工作区跨版本保留
- 提供 release 打包脚本
- 给结果文件提供固定“最新入口”文件名，降低普通用户找文件的成本

## 关键文件

- `app.py`
  桌面 GUI 主界面
- `runtime.py`
  设置存储、工作区路径、脚本桥接、更新检查
- `build_release.py`
  PyInstaller 打包 + release 预览目录生成
- `release_assets/应用使用说明.html`
  给普通用户看的主要说明页

## 当前发布策略

当前采用的是“程序本体”和“用户数据”分离的思路：

- 程序本体放在 release 目录下的 `app/`
- 用户设置与工作区放在 `user-data/`

优点：

- 更新应用时，不会顺手把用户数据删掉
- 方便做本地 release 预览
- 方便以后把不同版本的应用本体替换掉，而保留同一份用户数据

## 为什么要保留环境变量覆盖

`runtime.py` 支持通过环境变量覆盖设置目录和默认工作区：

- `AMS_ASSISTANT_SETTINGS_DIR`
- `AMS_ASSISTANT_DEFAULT_WORKSPACE`
- `AMS_REQ2_BROWSER`

这样做的目的，是为了让：

- 源码模式
- 打包后的 exe
- release 预览目录

都能共用同一套逻辑，而不需要写三套路径处理。

## 打包命令

在仓库根目录运行：

```powershell
python .\desktop_app\build_release.py
```

生成结果：

- `普通用户体验区-桌面应用版-release预览/`
- `desktop_release_build/AMS-Assistant-Desktop-v0.1.0.zip`

## 自检思路

`launch_ams_desktop_app.py` 支持隐藏的 `--self-test` 模式。这个模式的用途是验证：

- 打包后的 exe 能不能初始化工作区
- 打包后的 exe 能不能调用 req1 生成合同
- 打包后的 exe 能不能找到内置说明页

这可以帮助我们确认“源码能跑”和“打包后也能跑”不是两回事。
