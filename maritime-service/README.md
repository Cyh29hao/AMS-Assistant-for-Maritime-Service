# maritime-service

这是 AMS 仓库里真正负责“执行业务动作”的 skill 目录。

当前已经落地的第一条真实流程是：

1. `req1 系统出合同`

它支持：

1. 从 JSON 生成合同
2. 从 Excel 生成合同
3. 批量生成示例合同
4. 自动校验示例输出
5. 生成空白 Excel 填写模板

## 先看哪里

### 普通使用者

1. 看 [`开始看这里.md`](./%E5%BC%80%E5%A7%8B%E7%9C%8B%E8%BF%99%E9%87%8C.md)
2. 看 [`docs/01-小学生级别使用教程.md`](./docs/01-%E5%B0%8F%E5%AD%A6%E7%94%9F%E7%BA%A7%E5%88%AB%E4%BD%BF%E7%94%A8%E6%95%99%E7%A8%8B.md)
3. 直接点批处理文件

### 维护者

1. 看 [`SKILL.md`](./SKILL.md)
2. 看 [`references/contract-generation.md`](./references/contract-generation.md)
3. 看 [`docs/03-开发者维护说明.md`](./docs/03-%E5%BC%80%E5%8F%91%E8%80%85%E7%BB%B4%E6%8A%A4%E8%AF%B4%E6%98%8E.md)

## 目录说明

```text
maritime-service/
├─ assets/contract_templates/
│  合同模板
├─ docs/
│  用户教程、维护说明、实施方案
├─ examples/
│  示例 JSON、示例 Excel
├─ output/
│  运行后输出的合同与报告
├─ references/
│  架构、路线图、工作流设计
├─ scripts/
│  主要 Python 脚本
├─ 0-安装依赖.bat
├─ 1-一键生成全部示例合同.bat
├─ 2-一键生成空白Excel模板.bat
├─ 3-把Excel拖到这里生成合同.bat
├─ 4-把JSON拖到这里生成合同.bat
├─ 5-一键自检.bat
├─ 6-打开教程.bat
└─ SKILL.md
```

## 最常用命令

```powershell
cd .\maritime-service
python -m pip install -r .\requirements.txt
python .\scripts\contract_workflow.py build-examples
python .\scripts\contract_workflow.py verify-examples
python .\scripts\contract_workflow.py make-workbook-template --output .\examples\workbooks\blank-contract-template.xlsx
python .\scripts\contract_workflow.py from-json --input .\examples\contract_requests\domestic-forwarder-mu-chuang-362.json
```

## Windows 用户建议

如果你不想敲命令，直接用这些文件：

1. `0-安装依赖.bat`
2. `1-一键生成全部示例合同.bat`
3. `2-一键生成空白Excel模板.bat`
4. `3-把Excel拖到这里生成合同.bat`
5. `4-把JSON拖到这里生成合同.bat`
6. `5-一键自检.bat`
7. `6-打开教程.bat`
