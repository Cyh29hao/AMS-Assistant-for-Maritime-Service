# 私密业务包示例目录

这个目录现在既是结构示例，也是一个可以直接打包的演示私密包来源目录。

它默认只覆盖两类内容：

- `req2` 的站点配置示例
- 帮助中心里的私密包说明页示例

这样做的好处是：

- 你可以先直接打一个演示包，测试导入流程
- 不会因为缺少真实 Word 模板而影响 `req1`
- 后续你只需要在这个基础上替换成自己的真实模板和真实配置

如果你不想用命令行，也可以直接在桌面应用里：

1. 进入 `设置`
2. 点击 `复制示例到工作区`
3. 回到同一区域点击 `制作私密包`

你以后真正要打包时，建议复制这个目录，填入你自己的私密模板和配置，再运行：

```powershell
python .\desktop_app\build_private_pack.py --source .\private_pack_source.example --output .\private-pack-dist\ams-private-pack-demo.amspack
```

## 最小结构

```text
private_pack_source.example/
  private-pack.json
  maritime-service/
    scripts/
      clearance_site_config.json
  desktop_app/
    release_assets/
      help/
        private-pack.html
```

## 说明

- `private-pack.json`
  私密包自己的说明信息
- `maritime-service/scripts/clearance_site_config.json`
  私密站点配置示例
- `desktop_app/release_assets/help/`
  可选：私密版帮助页示例

## 如果你要扩展成完整私密包

你可以在这个目录里继续加入：

- `maritime-service/assets/contract_templates/`
- `maritime-service/scripts/contract_template_registry.json`

这样导入后，`req1` 也会切换到私密模板和私密规则。

打好的 `.amspack` 发给内部人即可。
