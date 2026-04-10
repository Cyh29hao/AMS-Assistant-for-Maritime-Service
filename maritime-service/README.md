# maritime-service

这一层是 AMS Assistant 的业务脚本层。

如果把整个仓库理解成一套本地桌面工具，那么：

- `maritime-service/`
  负责“真正做事”
- `desktop_app/`
  负责“让普通用户更容易用”

## 公开边界提醒

如果你是从公开仓库直接进入这一层，请先知道：

- 这里更偏实施层与业务脚本层
- 其中的模板、样本和原始需求材料并不自动等于可无限制对外再分发
- 涉及账号、网页登录和查询的站点自动化，只能在已授权场景中使用

建议配合阅读：

- [../LICENSE](../LICENSE)
- [../docs/public-use/README.md](../docs/public-use/README.md)

## 当前包含的内容

### 合同生成中心

- Excel / JSON 输入
- 合同模板注册表
- Word 合同生成
- 示例与自检

### 通关查询中心

- Excel 模板与回填
- 示例与自检
- 网站登录态保存
- 登录态检查
- 单票查询
- 整表自动回填

### 船期与港区矩阵

- 已预留目录和入口
- 后续可以继续接自动化流程

## 常用入口

### 合同生成中心

- `0-安装依赖.bat`
- `1-一键生成全部示例合同.bat`
- `2-一键生成空白Excel模板.bat`
- `3-把Excel拖到这里生成合同.bat`
- `5-一键自检.bat`

### 通关查询中心

- `7-一键生成req2示例结果.bat`
- `10-一键自检req2.bat`
- `11-首次登录并保存req2网站登录态.bat`
- `12-检查req2网站登录态.bat`
- `13-把req2工作簿拖到这里从网站自动查询并回填.bat`
- `14-输入提单号测试req2网站查询.bat`

## 文档入口

- [开始看这里.md](./%E5%BC%80%E5%A7%8B%E7%9C%8B%E8%BF%99%E9%87%8C.md)
- [docs/00-如何验收这套skill.md](./docs/00-%E5%A6%82%E4%BD%95%E9%AA%8C%E6%94%B6%E8%BF%99%E5%A5%97skill.md)
- [docs/01-小学生级别使用教程.md](./docs/01-%E5%B0%8F%E5%AD%A6%E7%94%9F%E7%BA%A7%E5%88%AB%E4%BD%BF%E7%94%A8%E6%95%99%E7%A8%8B.md)
- [docs/09-req2-自动查通关使用教程.md](./docs/09-req2-%E8%87%AA%E5%8A%A8%E6%9F%A5%E9%80%9A%E5%85%B3%E4%BD%BF%E7%94%A8%E6%95%99%E7%A8%8B.md)
- [docs/11-req2网页登录使用教程.md](./docs/11-req2%E7%BD%91%E9%A1%B5%E7%99%BB%E5%BD%95%E4%BD%BF%E7%94%A8%E6%95%99%E7%A8%8B.md)

## 命令行运行

```powershell
cd .\maritime-service
python -m pip install -r .\requirements.txt
python .\scripts\contract_workflow.py build-examples
python .\scripts\contract_workflow.py verify-examples
python .\scripts\clearance_workflow.py build-examples
python .\scripts\clearance_workflow.py verify-examples
```

## 说明

如果你是普通用户，优先考虑用仓库根目录的桌面应用入口，而不是长期停留在这一层逐个点脚本。
