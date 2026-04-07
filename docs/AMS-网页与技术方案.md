# AMS 网页与技术方案

更新时间：2026-04-08

## 当前落地情况

截至 2026-04-08：

1. 仓库中已经有一个可直接运行的 `maritime-service/` 技术实现
2. 当前最完整的流程是 `req1 系统出合同`
3. 现阶段的主要执行方式是批处理入口加 Python 脚本
4. “本地网页 + 本地服务”的方案仍然成立，但属于下一阶段工程化目标

## 1. 总体方案

AMS 建议长期采用“本地网页 + 本地服务”的形态：

1. 用户在浏览器访问一个本地 HTML 页面
2. 页面负责上传文件、展示进度、下载结果
3. Python 后端负责 Excel 解析、Word 生成、网页登录自动化、定时任务

这比“纯静态 HTML”更适合当前需求，因为纯前端无法稳定完成：

1. Word 合同生成
2. Excel 文件读写
3. 网站自动登录与查询
4. 本地定时任务
5. 任务日志记录

## 2. 推荐技术栈

### 2.1 前端

1. HTML
2. CSS
3. 原生 JavaScript

说明：

本项目用户群体有限、页面数量有限，前期没有必要引入复杂前端框架。原生前端足够做出清晰稳定的内部工具界面。

### 2.2 后端

1. Python 3.12+
2. FastAPI
3. Uvicorn

说明：

FastAPI 适合：

1. 处理文件上传下载
2. 提供 JSON 接口
3. 和定时任务、数据库、自动化脚本整合

### 2.3 自动化与文档处理

1. Playwright：网页登录、查询、抓取结果
2. openpyxl：Excel 读写
3. pandas：复杂表格整理
4. python-docx 或 docxtpl：Word 模板填充

推荐：

1. 合同若主要是替换变量，优先 `docxtpl`
2. 合同若需要保留复杂格式并操作表格段落，配合 `python-docx`

### 2.4 调度与存储

1. APScheduler：定时任务
2. SQLite：任务记录、配置、日志索引
3. 本地文件目录：保存上传文件和导出文件

## 3. 信息架构

## 3.1 页面结构

建议网页导航只保留 5 个一级入口：

1. 首页
2. 合同中心
3. 通关查询
4. 船期汇总
5. 任务日志

## 3.2 首页

### 目标

让用户一眼看到今天最常用的动作。

### 内容

1. 三张功能卡片：
   - 生成合同
   - 查询通关
   - 生成船期表
2. 最近任务列表
3. 今日执行摘要
4. 常见问题入口

## 3.3 合同中心

### 页面区域

1. 上传区
2. 记录预览区
3. 模板识别区
4. 字段确认区
5. 结果下载区

### 用户流程

1. 上传业务表格
2. 查看识别出的记录
3. 选择某条或多条记录
4. 系统自动匹配模板
5. 用户确认或补充字段
6. 点击生成合同
7. 下载 `.docx`

## 3.4 通关查询

### 页面区域

1. 上传待回填表格
2. 查询配置
3. 执行进度
4. 结果列表
5. 导出按钮

### 用户流程

1. 上传业务表
2. 系统筛出未通关提单
3. 用户点击“开始查询”
4. 页面展示当前查询进度
5. 任务完成后展示成功/失败统计
6. 下载更新后的 Excel

## 3.5 船期汇总

### 页面区域

1. 上传 LINE UP
2. 识别船名预览
3. 区域选择
4. 查询执行状态
5. 矩阵结果预览
6. 导出结果

### 用户流程

1. 上传 LINE UP 文件
2. 系统解析船名
3. 用户选择只查九区/十区或全查
4. 点击开始查询
5. 系统生成矩阵表
6. 用户导出 Excel

## 3.6 任务日志

### 页面区域

1. 任务筛选
2. 列表表格
3. 详情抽屉或弹窗

### 字段建议

1. 任务类型
2. 开始时间
3. 结束时间
4. 状态
5. 成功数
6. 失败数
7. 输入文件
8. 输出文件

## 4. 页面原型建议

## 4.1 设计原则

1. 偏业务工具风格，强调清楚、稳定、直接
2. 重点信息前置，不做花哨视觉
3. 每个页面一个主动作按钮
4. 所有结果都能下载

## 4.2 交互原则

1. 上传后立即预览
2. 长任务必须显示进度
3. 失败必须显示原因
4. 用户可以回到上一步重新执行

## 5. 系统模块设计

## 5.1 前端模块

1. `dashboard`
2. `contract-generator`
3. `customs-status`
4. `lineup-report`
5. `task-history`

## 5.2 后端模块

1. `api`
2. `services`
3. `jobs`
4. `storage`
5. `automation`
6. `parsers`

建议目录结构：

```text
AMS/
  app/
    main.py
    api/
      routes_contracts.py
      routes_customs.py
      routes_lineup.py
      routes_tasks.py
    services/
      contract_service.py
      customs_service.py
      lineup_service.py
      task_service.py
    automation/
      browser.py
      customs_query.py
      lineup_query.py
    parsers/
      excel_parser.py
      contract_mapping.py
      lineup_parser.py
    templates/
      contracts/
    static/
      index.html
      css/
      js/
    db/
      models.py
      sqlite.py
    scheduler/
      jobs.py
  data/
    uploads/
    outputs/
    logs/
```

## 6. 核心数据流

## 6.1 合同生成数据流

```text
业务表格 -> Excel 解析 -> 记录识别 -> 模板选择 -> 字段映射 -> Word 生成 -> 输出文件
```

## 6.2 通关查询数据流

```text
业务表格 -> 筛出未通关记录 -> Playwright 登录网站 -> 逐票查询 -> 结果解析 -> Excel 回填 -> 输出文件
```

## 6.3 船期汇总数据流

```text
LINE UP -> 解析船名 -> Playwright 查询 -> 结果归类 -> 生成矩阵表 -> 输出 Excel
```

## 7. API 设计建议

## 7.1 合同接口

1. `POST /api/contracts/preview`
   - 上传表格并返回可识别记录
2. `POST /api/contracts/generate`
   - 生成合同
3. `GET /api/contracts/download/{task_id}`
   - 下载结果

## 7.2 通关接口

1. `POST /api/customs/preview`
   - 预览可查询记录
2. `POST /api/customs/run`
   - 执行查询回填
3. `GET /api/customs/download/{task_id}`
   - 下载结果

## 7.3 船期接口

1. `POST /api/lineup/preview`
   - 解析船名
2. `POST /api/lineup/run`
   - 执行查询并生成报表
3. `GET /api/lineup/download/{task_id}`
   - 下载结果

## 7.4 任务接口

1. `GET /api/tasks`
2. `GET /api/tasks/{task_id}`
3. `GET /api/tasks/{task_id}/logs`

## 8. 数据库设计建议

### 8.1 tasks

记录一次任务执行。

字段建议：

1. `id`
2. `task_type`
3. `status`
4. `input_file`
5. `output_file`
6. `started_at`
7. `finished_at`
8. `success_count`
9. `failure_count`
10. `error_summary`

### 8.2 task_logs

记录任务步骤日志。

字段建议：

1. `id`
2. `task_id`
3. `level`
4. `message`
5. `created_at`

### 8.3 settings

保存本地配置。

字段建议：

1. `key`
2. `value`
3. `updated_at`

## 9. 合同模块的实现建议

## 9.1 模板策略

建议不要直接依赖“红字识别”作为核心机制，而是建立一个字段映射层。

更稳的方案：

1. 先人工梳理每个模板需要的字段清单
2. 将模板变量显式命名
3. 由程序把 Excel 字段映射到模板变量

### 原因

1. Word 里颜色不一定稳定
2. 红字规则难以长期维护
3. 模板升级时更容易校验

## 9.2 建议的合同生成流程

1. 读取 Excel 行
2. 标准化字段名
3. 聚合同一合同下的多条货物明细
4. 根据业务规则匹配模板
5. 生成中间结构化数据
6. 渲染 Word 模板

## 10. 浏览器自动化建议

## 10.1 选择 Playwright 的原因

1. 对现代网页兼容更好
2. 等待机制比传统脚本更稳
3. 对登录、表单、表格抓取支持更成熟

## 10.2 自动化策略

1. 登录逻辑独立封装
2. 查询逻辑按“通关查询”和“船期查询”拆开
3. 尽量使用稳定选择器
4. 保存失败截图和 HTML 片段以便排查

## 10.3 风险兜底

1. 登录失败时中断任务并提示重新登录
2. 页面字段缺失时记日志，不直接崩溃
3. 单条查询失败继续处理后续记录

## 11. 定时任务设计

### 11.1 首期支持

1. 通关查询每日定时执行

### 11.2 建议行为

1. 定时任务运行前先检查上次任务是否未结束
2. 若已有执行中的同类任务，则跳过
3. 每次定时执行都生成一条任务记录

## 12. 开发优先级

### 第一阶段

1. FastAPI 基础框架
2. 静态 HTML 页面
3. 文件上传下载
4. 合同模块 MVP

### 第二阶段

1. 船期模块 MVP
2. Playwright 查询封装
3. 任务日志页面

### 第三阶段

1. 通关模块
2. 定时任务
3. 更完整的失败重试和截图日志

## 13. 建议的 MVP 页面清单

1. 首页
2. 合同生成页
3. 船期汇总页
4. 任务日志页

说明：

通关查询页可以在第二轮或第三轮加入，但合同页建议最先做。

## 14. 当前最重要的补充材料

1. 合同输入 Excel 样本
2. 合同模板变量清单
3. 通关回填 Excel 样本
4. 网站账号登录流程
5. 船期表输出模板 Excel

## 15. 一句话结论

AMS 最合适的落地方式，是一个本地运行的业务自动化网页工具：

前端轻、后端稳、文件处理强、自动化可追踪，并且先从合同生成这个最清晰的 MVP 切入。
