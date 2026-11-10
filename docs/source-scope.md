# 源码公开范围与运行说明

## 来源与处理

核心源码来自项目维护者提供的 2026 年 11 月 10 日源码快照。公开副本保留原始工具与插件结构，替换注释中的具体客户名称、企业标识，以及导出表格中的供应商品牌文本。部分表头识别正则中的企业简称也同步替换。

`source-manifest.json` 记录原始文件和公开副本的 SHA256。演示脚本、虚构目录和回归测试为本次整理新增；历史评测结果不是这个虚构演示跑出的成绩。

## 已公开

- 三个 Python 核心工具：选型与检查、比单流程、经验构建。
- 自检插件与插件元数据。兼容性标识 `fyys-selfcheck`、`FYYS_WORKSPACE` 和结果字段保留，便于对照原始接口。
- 虚构产品目录、本地演示与关键规则回归测试。

## 未公开

- 具体客户的生成、修订、补价脚本，以及方案文档。
- 含业务实例、价格规则和内部路径的技能文件、记忆与实战笔记。
- 生产配置、内部平台前后端接口及同步任务。它们依赖原系统的权限、数据库和目录结构，不能单独作为可运行服务发布。
- 真实产品库、价格、人工配单经验和评测答案。
- Hermes 框架本身与模型凭据。

## 本地演示

```bash
pip install -r requirements.txt
python examples/demo.py --output demo-output
python -m unittest discover -s tests -v
```

演示在临时目录创建隔离工作区，复制核心工具，放入虚构产品目录，再生成两行需求单。4K 会议摄像机行选择虚构 4K 产品，门禁控制器行说明无对应产品。导出文件保存在指定输出目录，临时目录自动清理。

演示使用预先指定的选择调用 `pick`，目的是验证工具流程；没有运行自主 Agent，也不展示模型准确率。

## 使用自己的资料

核心工具默认从工作区 `references/category_products_full_with_internal_model.json` 读取分类、子分类和产品数据，结构可参考 `examples/catalog.json`。该目录被 Git 忽略，以减少误上传业务资料的可能。

准备资料后可运行：

```bash
python scripts/kb_tool.py search "会议摄像机" "4K，12倍光学变焦"
python scripts/bidan.py draft 需求单.xlsx
python scripts/bidan.py show 方案表格/需求单_底稿.json
python scripts/bidan.py pick 方案表格/需求单_底稿.json 1=你的型号
python scripts/bidan.py export 方案表格/需求单_底稿.json
```

`pick` 使用草稿编号而非 Excel 行号。脚本包含特定产品系列与业务约束，接入其他厂商前需要检查并调整。

## 插件边界

插件依赖兼容的 Hermes 钩子接口，不是独立服务。`FYYS_WORKSPACE` 应指向包含 `scripts/kb_tool.py` 的工作区。插件会执行该脚本并扫描工作区中的表格，应只用于可信工作区。

插件在工具结果中反馈检查问题，并在最终回复提示遗留错误；它不构成权限隔离，也不保证阻止文件被下载或交付。插件异常采用 fail-open 行为，完整 Hermes 环境与并发会话未在此公开版完成集成测试。

## 验证范围

本次验证使用 Python 3.11 与 openpyxl，覆盖虚构需求单的完整工具链、4K 参数检查、停单候选过滤、非主营识别、无线连接类型区分和检查器异常反馈。未调用线上系统或付费模型，也未重跑历史准确率评测。
