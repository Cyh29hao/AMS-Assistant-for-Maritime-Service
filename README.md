# AMS Assistant for Maritime Service

AMS（Assistant for Maritime Service）正在从“航运自动化脚本集合”收成一套更像产品的本地桌面工具。

目前这套仓库已经同时具备两条使用方式：

- `脚本/批处理模式`
  适合先快速跑通 req1、req2、教程与示例
- `桌面应用模式`
  适合普通用户通过一个本地 GUI 使用 req1、req2，并为 req3 预留统一入口

## 现在已经能做什么

### 已可用

1. `req1 系统出合同`
   - 支持 Excel / JSON 输入
   - 自动选择模板
   - 生成 Word 合同和 JSON 摘要
2. `req2 自动查通关`
   - 支持 Excel 回填
   - 支持网站登录态保存、检查登录态、单票测试、整表自动回填
3. `桌面应用原型`
   - 一个本地 GUI 入口
   - 工作区固定
   - 配置可保存
   - 检查更新入口
   - 为普通用户准备的 HTML 使用说明
4. `req3`
   - 入口已预留
   - 后续可继续接船期表自动化

## 如果你是普通用户

最推荐的路径不是直接研究源码，而是：

1. 去 GitHub Releases 页面下载 release 包
2. 解压后打开桌面应用
3. 通过 GUI 使用 req1 和 req2

如果当前还没有发布好的 release 包，也可以自己从源码生成本地 release 预览：

```powershell
python -m pip install -r .\desktop_app\requirements-desktop.txt
python .\desktop_app\build_release.py
```

生成后会得到：

- `普通用户体验区-桌面应用版-release预览/`
- `desktop_release_build/AMS-Assistant-Desktop-v0.1.0.zip`

桌面应用相关说明：

- [desktop_app/README.md](./desktop_app/README.md)
- [desktop_app/发布评估与后续方案.md](./desktop_app/%E5%8F%91%E5%B8%83%E8%AF%84%E4%BC%B0%E4%B8%8E%E5%90%8E%E7%BB%AD%E6%96%B9%E6%A1%88.md)

## 如果你是想先快速验功能的人

### 方式一：直接跑原来的批处理入口

进入 [maritime-service](./maritime-service)，然后：

1. 双击 `0-安装依赖.bat`
2. 双击 `1-一键生成全部示例合同.bat`
3. 双击 `5-一键自检.bat`

如果要试 req2：

1. 双击 `7-一键生成req2示例结果.bat`
2. 双击 `10-一键自检req2.bat`
3. 如果要接真网站，再双击 `11-首次登录并保存req2网站登录态.bat`

### 方式二：用源码直接启动桌面应用

```powershell
python -m pip install -r .\desktop_app\requirements-desktop.txt
python .\launch_ams_desktop_app.py
```

如果你在 Windows 上，希望保留一个最简单的启动入口，也可以用：

- [启动AMS桌面应用.bat](./%E5%90%AF%E5%8A%A8AMS%E6%A1%8C%E9%9D%A2%E5%BA%94%E7%94%A8.bat)

## 仓库结构

```text
.
├─ maritime-service/
│  业务脚本、批处理、模板、示例、文档
├─ desktop_app/
│  桌面 GUI、设置存储、release 打包、用户说明
├─ docs/
│  项目级材料
├─ our_requirements/
│  原始需求材料
├─ launch_ams_desktop_app.py
│  桌面应用 Python 入口
└─ 启动AMS桌面应用.bat
   Windows 下的源码模式启动器
```

## 当前推荐的阅读顺序

1. [desktop_app/README.md](./desktop_app/README.md)
2. [maritime-service/开始看这里.md](./maritime-service/%E5%BC%80%E5%A7%8B%E7%9C%8B%E8%BF%99%E9%87%8C.md)
3. [maritime-service/docs/00-如何验收这套skill.md](./maritime-service/docs/00-%E5%A6%82%E4%BD%95%E9%AA%8C%E6%94%B6%E8%BF%99%E5%A5%97skill.md)
4. [maritime-service/docs/10-req2-如何验收.md](./maritime-service/docs/10-req2-%E5%A6%82%E4%BD%95%E9%AA%8C%E6%94%B6.md)

## 维护建议

这套仓库现在比较适合继续按三层结构往前扩：

1. `maritime-service/scripts`
   负责业务 workflow
2. `desktop_app/runtime.py`
   负责把 workflow 封装成桌面动作
3. `desktop_app/app.py`
   负责把动作挂到 GUI

这样后面继续做 req3、继续增强 req2、或者替换界面形态，都会更稳。

更详细的验收步骤见：

- [`maritime-service/docs/00-如何验收这套skill.md`](./maritime-service/docs/00-%E5%A6%82%E4%BD%95%E9%AA%8C%E6%94%B6%E8%BF%99%E5%A5%97skill.md)

## `maritime-service` 子目录说明

`maritime-service/` 是这个仓库里真正“拿来干活”的部分。

它的定位不是聊天机器人，而是一套航运业务自动化 skill 包。当前重点是 `req1 系统出合同`，也就是：

1. 接收业务输入
2. 识别模板类型
3. 规范化字段
4. 把内容写进既有 Word 模板
5. 导出最终合同和结构化摘要

建议继续阅读：

- [`maritime-service/README.md`](./maritime-service/README.md)
- [`maritime-service/SKILL.md`](./maritime-service/SKILL.md)
- [`maritime-service/references/contract-generation.md`](./maritime-service/references/contract-generation.md)

## 输入与输出

### 当前支持的输入

1. JSON 请求文件
2. Excel 工作簿

### 当前输出

1. Word 合同 `.docx`
2. 结构化生成摘要 `.json`
3. 自检报告

### 当前内置模板

1. `domestic_forwarder`
2. `domestic_logistics`
3. `overseas_astar`
4. `overseas_rta`

模板注册表位于：

- [`maritime-service/scripts/contract_template_registry.json`](./maritime-service/scripts/contract_template_registry.json)

## 适合谁用

### 业务同事

如果你主要工作是出合同，不想自己改代码，只想：

1. 点批处理
2. 填 Excel
3. 拖文件生成 Word

那你主要看：

- [`maritime-service/开始看这里.md`](./maritime-service/%E5%BC%80%E5%A7%8B%E7%9C%8B%E8%BF%99%E9%87%8C.md)
- [`maritime-service/docs/01-小学生级别使用教程.md`](./maritime-service/docs/01-%E5%B0%8F%E5%AD%A6%E7%94%9F%E7%BA%A7%E5%88%AB%E4%BD%BF%E7%94%A8%E6%95%99%E7%A8%8B.md)
- [`maritime-service/docs/06-常见报错与处理.md`](./maritime-service/docs/06-%E5%B8%B8%E8%A7%81%E6%8A%A5%E9%94%99%E4%B8%8E%E5%A4%84%E7%90%86.md)

### 维护者

如果你要继续扩展这套系统，重点看：

1. 模板如何新增或修改
2. Excel 字段如何映射
3. 以后如何接网站自动化和本地网页

建议看：

- [`maritime-service/docs/02-系统出合同实施方案.md`](./maritime-service/docs/02-%E7%B3%BB%E7%BB%9F%E5%87%BA%E5%90%88%E5%90%8C%E5%AE%9E%E6%96%BD%E6%96%B9%E6%A1%88.md)
- [`maritime-service/docs/03-开发者维护说明.md`](./maritime-service/docs/03-%E5%BC%80%E5%8F%91%E8%80%85%E7%BB%B4%E6%8A%A4%E8%AF%B4%E6%98%8E.md)
- [`maritime-service/docs/07-如何新增或调整模板.md`](./maritime-service/docs/07-%E5%A6%82%E4%BD%95%E6%96%B0%E5%A2%9E%E6%88%96%E8%B0%83%E6%95%B4%E6%A8%A1%E6%9D%BF.md)
- [`maritime-service/docs/08-字段填写对照表.md`](./maritime-service/docs/08-%E5%AD%97%E6%AE%B5%E5%A1%AB%E5%86%99%E5%AF%B9%E7%85%A7%E8%A1%A8.md)
- [`docs/AMS-网页与技术方案.md`](./docs/AMS-%E7%BD%91%E9%A1%B5%E4%B8%8E%E6%8A%80%E6%9C%AF%E6%96%B9%E6%A1%88.md)

## 已知依赖

运行合同工作流至少需要：

1. Python 3.12 或更高版本
2. `python-docx`
3. `openpyxl`

可以直接执行：

```powershell
cd .\maritime-service
python -m pip install -r .\requirements.txt
```

## 关于中文乱码

为了降低 Windows 环境里的乱码概率，这个仓库已经做了这些处理：

1. Markdown 文档统一用 UTF-8 保存
2. 批处理文件改成纯 ASCII 内容，避免 `cmd` 直接解释中文行
3. 文档内容尽量放在 `.md` 中，而不是直接写在 `.bat` 中

如果你仍然看到 Markdown 乱码，优先尝试：

1. 用 VS Code 打开
2. 用支持 UTF-8 的编辑器打开
3. 不要用旧版记事本或某些会强制 ANSI 的查看器

## 后续开发建议

推荐按这个顺序推进：

1. 把 `req1` 的模板和字段映射继续贴近真实业务
2. 给 `req1` 增加更细的校验和异常提示
3. 再落地 `req2` 的网站登录、查询、回填
4. 再落地 `req3` 的 LINE UP 解析、区域查询和矩阵导出
5. 最后再做本地网页界面和定时任务

不建议一开始就上 GUI 自动点击一切。优先顺序应该是：

1. 先能用
2. 再稳定
3. 再自动
4. 最后才是“完全不用人看”

## 相关文档

- [`docs/AMS-PRD.md`](./docs/AMS-PRD.md)
- [`docs/AMS-网页与技术方案.md`](./docs/AMS-%E7%BD%91%E9%A1%B5%E4%B8%8E%E6%8A%80%E6%9C%AF%E6%96%B9%E6%A1%88.md)
- [`our_requirements/需求整理.md`](./our_requirements/%E9%9C%80%E6%B1%82%E6%95%B4%E7%90%86.md)
- [`maritime-service/README.md`](./maritime-service/README.md)

## 许可证与说明

这个仓库当前更接近内部业务自动化原型与实施仓库。

如果你准备对外公开使用，建议补充：

1. 明确许可证
2. 数据脱敏说明
3. 模板来源说明
4. 账号与站点自动化的安全边界说明
