<p align="center">
  <img src="./assets/readme/ams-hero.svg" alt="AMS Assistant Hero" width="100%" />
</p>

<h1 align="center">AMS Assistant for Maritime Service</h1>

<p align="center">
  一个面向航运业务团队的本地桌面自动化工作台。<br />
  把 Excel、Word、网页登录查询、结果回填和后续扩展，收进同一个更容易交付、更容易验收的本地应用里。
</p>

<p align="center">
  <a href="https://github.com/Cyh29hao/AMS-Assistant-for-Maritime-Service/releases">
    <img alt="Release" src="https://img.shields.io/github/v/release/Cyh29hao/AMS-Assistant-for-Maritime-Service?display_name=tag&label=release&color=0f766e" />
  </a>
  <img alt="Platform" src="https://img.shields.io/badge/platform-Windows-0f172a?logo=windows&logoColor=white" />
  <img alt="Workflows" src="https://img.shields.io/badge/workflows-Req1%20%7C%20Req2%20%7C%20Req3-0b6bcb" />
  <img alt="Mode" src="https://img.shields.io/badge/mode-local--first%20desktop-1d4ed8" />
  <img alt="Status" src="https://img.shields.io/badge/status-prototype%20to%20delivery-f59e0b" />
</p>

<p align="center">
  <a href="https://github.com/Cyh29hao/AMS-Assistant-for-Maritime-Service/releases"><strong>下载桌面版</strong></a>
  ·
  <a href="./desktop_app/README.md"><strong>桌面版说明</strong></a>
  ·
  <a href="./maritime-service/README.md"><strong>业务脚本层</strong></a>
  ·
  <a href="./docs/public-use/README.md"><strong>公开使用边界</strong></a>
</p>

> [!IMPORTANT]
> 这个仓库当前更接近“内部业务自动化原型 + 实施仓库”，已经可以公开展示、公开下载桌面版、公开评估实现思路，但它默认不是“无边界的通用开源 SaaS 产品”。

## 这是什么

AMS Assistant 的目标不是做一个只会聊天的 maritime copilot，而是做一个能真正落到业务动作上的本地桌面工作台：

- `Req1 系统出合同`：在 Excel 填好数据后，一键生成 Word 合同。
- `Req2 自动查通关`：复用本地保存的网页登录态，查状态并把结果回填到工作簿。
- `Req3 船期表`：入口已预留，后续继续接入实际工作流。

它更像一个把“人手在 Excel、Word、浏览器之间搬运数据”的过程压缩起来的本地应用，而不是单独一条脚本。

## 一眼看懂这个产品

<p align="center">
  <img src="./assets/readme/ams-workflows.svg" alt="AMS Assistant Workflows" width="100%" />
</p>

## 现在已经能做什么

| 模块 | 普通用户能感受到什么 | 当前状态 |
| --- | --- | --- |
| `Req1 系统出合同` | 填 Excel，点一次，得到 Word 合同和固定命名的最新结果文件 | 已可用 |
| `Req2 自动查通关` | 保存一次登录态后，按提单号批量查询并回填业务表 | 已可用 |
| `Req2 网站查询层` | 登录检查、单票测试、整表更新、结果留档 | 已可用 |
| `Req3 船期表` | 入口、目录和扩展位置已经预留 | 待继续实现 |
| `Desktop App` | 通过 GUI 统一进入 req1 / req2 / req3，而不是翻文件夹找脚本 | 已可用 |

## 如果你只是想直接使用

### 最短路径

1. 打开 [Releases 页面](https://github.com/Cyh29hao/AMS-Assistant-for-Maritime-Service/releases)。
2. 下载最新的桌面版压缩包。
3. 解压后，双击 `1-启动AMS桌面应用.bat`。
4. 在桌面应用里体验 `Req1 出合同` 和 `Req2 查通关`。

### 你会得到什么体验

- 不用先研究源码目录结构。
- 不用自己拼接脚本命令。
- 常用入口集中在一个本地 GUI 里。
- 配置和工作区与应用本体分离，后续更新更容易保留既有设置。

## 为什么这套东西比“几个零散脚本”更值得继续做

- `本地优先`：业务文件、登录态、工作区都留在本机，更贴近实际业务场景。
- `面向交付`：不仅有脚本，还有批处理入口、桌面 GUI、示例、验收路径和文档。
- `面向更新`：桌面版已经按“程序本体”和“用户数据”分离的方向组织，后续更适合做 release、做升级、做配置保留。
- `面向真实工作流`：Req1 和 Req2 都不是只停留在 demo 文案，而是已经能跑通真实输入到真实输出。

## 仓库结构怎么读最清楚

| 目录 | 建议谁看 | 作用 |
| --- | --- | --- |
| [`desktop_app/`](./desktop_app/README.md) | 普通用户、产品视角、交付视角 | 桌面 GUI、设置持久化、打包、更新检查 |
| [`maritime-service/`](./maritime-service/README.md) | 实施同学、脚本维护者 | 业务脚本、模板、示例、批处理、详细教程 |
| [`docs/public-use/`](./docs/public-use/README.md) | 对外展示、准备公开的人 | 许可证、脱敏、模板来源、自动化安全边界 |
| [`our_requirements/`](./our_requirements/README.md) | 需求回溯、内部评估 | 历史需求原始材料，不是普通用户入口 |

## 对外公开前，先看这些边界

- [LICENSE](./LICENSE)
- [公开使用边界总览](./docs/public-use/README.md)
- [数据脱敏说明](./docs/public-use/data-sanitization.md)
- [模板来源说明](./docs/public-use/template-sources.md)
- [账号与站点自动化安全边界](./docs/public-use/automation-security-boundaries.md)
- [历史材料目录说明](./our_requirements/README.md)

## 推荐阅读顺序

### 面向普通用户

1. [Releases 页面](https://github.com/Cyh29hao/AMS-Assistant-for-Maritime-Service/releases)
2. [desktop_app/README.md](./desktop_app/README.md)

### 面向实施与维护

1. [maritime-service/README.md](./maritime-service/README.md)
2. [maritime-service/开始看这里.md](./maritime-service/%E5%BC%80%E5%A7%8B%E7%9C%8B%E8%BF%99%E9%87%8C.md)
3. [maritime-service/docs/00-如何验收这套skill.md](./maritime-service/docs/00-%E5%A6%82%E4%BD%95%E9%AA%8C%E6%94%B6%E8%BF%99%E5%A5%97skill.md)
4. [maritime-service/docs/10-req2-如何验收.md](./maritime-service/docs/10-req2-%E5%A6%82%E4%BD%95%E9%AA%8C%E6%94%B6.md)

## 当前定位，一句话说清

这是一个已经开始具备“可下载、可演示、可验收、可继续落地”的航运业务自动化桌面产品雏形，同时也保留了足够完整的实施仓库结构，方便继续扩 Req3、继续接新站点、继续把原型往真实业务交付推进。
