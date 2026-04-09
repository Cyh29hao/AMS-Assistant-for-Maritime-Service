# AMS Assistant for Maritime Service

AMS Assistant for Maritime Service 是一个面向航运业务场景的本地自动化实施仓库。

它当前更接近：

- 内部业务自动化原型仓库
- 落地实施仓库
- 桌面工具与业务脚本并行演进的工程仓库

它当前还不等于：

- 默认宽松开源的通用产品仓库
- 不加边界即可直接对外商用的交付包
- 不需要脱敏和权限审查即可公开传播的一切材料集合

## 现在有什么

### 已可运行

1. `req1 系统出合同`
   - Excel / JSON 输入
   - Word 合同生成
   - 示例、自检、模板映射
2. `req2 自动查通关`
   - 网站登录态保存
   - 检查登录态
   - 单票测试
   - 整表自动回填
3. `desktop_app`
   - 本地 GUI
   - 工作区固定
   - 设置持久化
   - release 打包脚本
4. `req3`
   - 入口已预留

## 如果你是普通用户

优先走 release，而不是直接研究源码。

### 推荐路径

1. 打开 [Releases 页面](https://github.com/Cyh29hao/AMS-Assistant-for-Maritime-Service/releases)
2. 下载最新桌面版压缩包
3. 解压
4. 双击 `1-启动AMS桌面应用.bat` 或 `Start AMS Assistant.bat`

如果你只是想先了解桌面版实现思路，看这里：

- [desktop_app/README.md](./desktop_app/README.md)

## 如果你是维护者或评估者

你更适合从这几层理解仓库：

1. `maritime-service/`
   - 业务脚本、批处理、模板、示例、业务文档
2. `desktop_app/`
   - 桌面 GUI、设置、打包、用户引导
3. `docs/`
   - 项目级材料和公开说明
4. `our_requirements/`
   - 历史需求原始材料

## 对外公开前，先看这些

如果这个仓库要继续对外传播、演示、交给更多人下载，请先看这组文件：

### 许可证

- [LICENSE](./LICENSE)

### 公开使用边界

- [docs/public-use/README.md](./docs/public-use/README.md)
- [docs/public-use/data-sanitization.md](./docs/public-use/data-sanitization.md)
- [docs/public-use/template-sources.md](./docs/public-use/template-sources.md)
- [docs/public-use/automation-security-boundaries.md](./docs/public-use/automation-security-boundaries.md)

### 历史材料目录说明

- [our_requirements/README.md](./our_requirements/README.md)

## 最短阅读顺序

### 面向普通用户

1. [Releases 页面](https://github.com/Cyh29hao/AMS-Assistant-for-Maritime-Service/releases)
2. [desktop_app/README.md](./desktop_app/README.md)

### 面向业务实施

1. [maritime-service/README.md](./maritime-service/README.md)
2. [maritime-service/开始看这里.md](./maritime-service/%E5%BC%80%E5%A7%8B%E7%9C%8B%E8%BF%99%E9%87%8C.md)
3. [maritime-service/docs/00-如何验收这套skill.md](./maritime-service/docs/00-%E5%A6%82%E4%BD%95%E9%AA%8C%E6%94%B6%E8%BF%99%E5%A5%97skill.md)
4. [maritime-service/docs/10-req2-如何验收.md](./maritime-service/docs/10-req2-%E5%A6%82%E4%BD%95%E9%AA%8C%E6%94%B6.md)

## 仓库结构

```text
.
├─ maritime-service/
│  业务脚本、批处理、模板、示例、业务文档
├─ desktop_app/
│  桌面 GUI、设置持久化、release 打包、用户说明
├─ docs/
│  项目级材料与公开边界说明
├─ our_requirements/
│  历史需求原始材料
├─ launch_ams_desktop_app.py
│  桌面应用源码入口
└─ 启动AMS桌面应用.bat
   Windows 下的源码启动器
```

## 当前公开定位

一句话概括：

这个仓库已经足够公开展示、公开评估、公开下载桌面版 release，但默认仍应按“保守公开、授权使用、先脱敏再传播”的方式对待。
