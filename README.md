# AMS Assistant for Maritime Service

AMS（Assistant for Maritime Service）是一套面向航运业务场景的本地自动化工作台原型。

这个仓库现在不是“只写了想法”的空架子，而是已经包含一条真正可运行的业务流程：

1. `req1 系统出合同` 已经落地成可执行脚本
2. 支持用 `Excel` 或 `JSON` 作为输入
3. 支持自动选择合同模板并生成 `Word .docx`
4. 支持一键跑示例、一键自检、一键生成空白 Excel 模板
5. `req2 通关查询`、`req3 船期汇总` 已经整理了需求与后续落地方案，适合作为第二阶段开发

如果你第一次来到这个仓库，最推荐的阅读顺序是：

1. 看 [`maritime-service/开始看这里.md`](./maritime-service/%E5%BC%80%E5%A7%8B%E7%9C%8B%E8%BF%99%E9%87%8C.md)
2. 看 [`maritime-service/docs/00-如何验收这套skill.md`](./maritime-service/docs/00-%E5%A6%82%E4%BD%95%E9%AA%8C%E6%94%B6%E8%BF%99%E5%A5%97skill.md)
3. 直接双击 `maritime-service` 目录里的批处理文件开始跑

## 当前交付状态

### 已完成

1. `maritime-service/` 目录已经整理成一套可直接运行的 skill 包
2. 合同工作流已经可以生成 4 份示例合同
3. 已提供空白 Excel 模板、示例 JSON、示例 Excel、详细教程、开发者说明
4. 已修复 Windows `cmd` 下批处理的中文编码问题
5. 已统一 Markdown 文件编码，减少中文乱码风险

### 已规划但未落地为生产功能

1. `req2` 网站登录后自动查通关并回填 Excel
2. `req3` 读取 LINE UP 后自动查询网站并生成矩阵型船期表
3. 本地网页化界面
4. 任务日志面板和定时调度面板

## 仓库结构

```text
.
├─ index.html / styles.css / script.js
│  根目录说明页，可作为项目展示入口
├─ docs/
│  项目级 PRD 与网页/技术方案
├─ our_requirements/
│  原始需求材料、邮件、样例附件
└─ maritime-service/
   ├─ SKILL.md
   ├─ README.md
   ├─ 0-安装依赖.bat
   ├─ 1-一键生成全部示例合同.bat
   ├─ 2-一键生成空白Excel模板.bat
   ├─ 3-把Excel拖到这里生成合同.bat
   ├─ 4-把JSON拖到这里生成合同.bat
   ├─ 5-一键自检.bat
   ├─ 6-打开教程.bat
   ├─ assets/
   │  合同模板
   ├─ examples/
   │  示例 JSON 与 Excel
   ├─ docs/
   │  用户教程、实施方案、维护说明
   ├─ references/
   │  架构设计、路线图、流程目录
   └─ scripts/
      核心 Python 脚本
```

## 最短上手路径

这一段是写给“只想马上试一下的人”的。

### 方式一：完全不看代码，直接点批处理

1. 先进入 [`maritime-service`](./maritime-service)
2. 双击 `0-安装依赖.bat`
3. 双击 `1-一键生成全部示例合同.bat`
4. 去 `maritime-service/output/contracts/` 看有没有生成 4 份 Word 合同
5. 双击 `5-一键自检.bat`
6. 如果看到 4 条 `通过`，说明核心流程已经跑通

### 方式二：自己填 Excel 再生成

1. 双击 `2-一键生成空白Excel模板.bat`
2. 打开 `maritime-service/examples/workbooks/blank-contract-template.xlsx`
3. 按教程填写一份测试业务
4. 把这个 Excel 文件拖到 `3-把Excel拖到这里生成合同.bat`
5. 去 `maritime-service/output/contracts/` 查看生成结果

### 方式三：开发者命令行运行

```powershell
cd .\maritime-service
python -m pip install -r .\requirements.txt
python .\scripts\contract_workflow.py build-examples
python .\scripts\contract_workflow.py verify-examples
python .\scripts\contract_workflow.py make-workbook-template --output .\examples\workbooks\blank-contract-template.xlsx
python .\scripts\contract_workflow.py from-json --input .\examples\contract_requests\domestic-forwarder-mu-chuang-362.json
python .\scripts\contract_workflow.py from-workbook --input .\examples\workbooks\blank-contract-template.xlsx
```

## 5 分钟验收清单

如果你要判断“这个仓库现在是不是能交付给别人试”，就按下面这套最短验收：

1. 成功执行 `0-安装依赖.bat`
2. 成功执行 `1-一键生成全部示例合同.bat`
3. `maritime-service/output/contracts/` 下出现 4 份 `.docx`
4. 成功执行 `5-一键自检.bat`
5. 自检结果显示所有示例 `通过`
6. 人工打开任意一个 Word，确认这些字段已经写进去：
   - 合同编号
   - 日期
   - 船名
   - 货描
   - 装港
   - 卸港
   - laycan
   - 费用条款
   - 货物明细行

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
