# 金十数据新闻提取程序

这是一个轻量级 Python 脚本，用于从金十数据接口抓取新闻/快讯，并仅导出两个 `Word` 摘要日报。

## 功能
- 抓取金十新闻快讯（支持翻页）
- 去重
- 按最近小时数过滤（默认最近 24 小时）
- 按关键词过滤
- 仅输出两个 Word 文档：昨日总结 + 今日最新

## 环境要求
- Python 3.9+

## 安装依赖
```bash
pip install -r requirements.txt
```

## 使用方法
在当前目录执行：

```bash
python jin10_news_collector.py
```

常见参数：

```bash
python jin10_news_collector.py --pages 5 --since-hours 48 --keyword 美联储 --output-dir output
```

参数说明：
- `--pages`：抓取页数，默认 `3`
- `--since-hours`：仅保留最近 N 小时新闻，默认 `24`
- `--keyword`：关键词过滤（可选）
- `--output-dir`：输出目录，默认 `output`
- `--timeout`：请求超时秒数，默认 `15`
- `--delay`：分页请求间隔秒数，默认 `0.8`

## 输出
程序会在输出目录生成以下 Word 文档：
- `reports/jin10_yesterday_summary_YYYYMMDD.docx`
- `reports/jin10_today_latest_YYYYMMDD.docx`
- `reports/latest_yesterday_summary.docx`
- `reports/latest_today_latest.docx`

日报内容按主题自动分组：
- 宏观政策
- 外汇贵金属
- 能源与大宗
- 股市公司
- 加密市场
- 其他

## 说明
- 金十接口字段可能调整，脚本里做了多种字段兼容解析。
- 若接口限制增强（例如鉴权升级），可在脚本中更新请求头。
