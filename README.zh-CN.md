# ⚽ 澳彩足球赔率更新工具

[English documentation](README.md)

抓取澳彩足球赔率 JSON 数据，并按预期的 Excel 样表格式增量更新赔率。项目提供命令行脚本和 Tkinter 图形界面。

## 📦 环境要求

- Python 3.10 或更高版本
- 可访问 `https://www.macauslot.com`

### macOS 环境初始化

使用 Homebrew 安装 Python 3.13 和 Tk 9：

```bash
brew install python-tk@3.13
```

创建项目虚拟环境并安装依赖：

```bash
cd /Users/niuhui/PycharmProjects/macauslotCrawler
/opt/homebrew/bin/python3.13 -m venv .venv
source .venv/bin/activate
python -m pip install -r requirements.txt
```

后续运行时，只需激活已有虚拟环境：

```bash
cd /Users/niuhui/PycharmProjects/macauslotCrawler
source .venv/bin/activate
```

需要退出虚拟环境时执行：

```bash
deactivate
```

### 通用安装

如果 Python 和 Tk 已经配置完成，安装依赖：

```bash
python3 -m pip install -r requirements.txt
```

## 🖥️ 图形界面

启动 GUI：

```bash
cd /Users/niuhui/PycharmProjects/macauslotCrawler
source .venv/bin/activate
python macauslot_odds_gui.py
```

界面支持：

- 立即更新全场三合一赔率
- 按分钟定时更新，并增加随机延迟
- 配置 HTTP/HTTPS 代理
- 在可信代理环境下关闭 SSL 证书校验
- 查看实时抓取日志

默认输出文件：

```text
outputs/清洗后的表格.xlsx
```

## ⌨️ 命令行

更新默认 Excel 文件：

```bash
python scrape_macauslot_odds.py
```

导出全部已知玩法的原始字段：

```bash
python scrape_macauslot_odds.py --format raw
```

指定代理：

```bash
python scrape_macauslot_odds.py --proxy-host 127.0.0.1 --proxy-port 7890
```

查看完整参数：

```bash
python scrape_macauslot_odds.py --help
```

## 📊 可选 Excel 样表

程序会尝试读取：

```text
assets/sample_template.xlsx
```

该文件用于复用表格样式和列宽。文件不存在时，程序会使用内置样式继续生成 Excel，不影响基本抓取和更新功能。

## 📁 项目文件

```text
scrape_macauslot_odds.py  抓取、数据清洗和 Excel 增量更新逻辑
macauslot_odds_gui.py     Tkinter 图形界面和定时任务
requirements.txt          Python 第三方依赖
README.md                 英文主文档
```
