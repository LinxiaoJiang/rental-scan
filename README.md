# Rental Scan

房东招租助手技能：每天扫描 Facebook 租房群和小红书的新帖，挑出与房源匹配的求租人，汇总给房东确认后再联系，并用 Excel 台账记录。

> 脱敏公开版：不含任何账号、群组 ID、房源信息或客户数据。

## 文件
- `SKILL.md` — 工作流程与文案模板（按占位符填入自己的房源信息）
- `tools/rental_log.py` — Excel 台账脚本
- `ledger_template.csv` — 台账列名

## 依赖
- Python 3 + `requirements.txt`（openpyxl）
- 运行环境相关（可替换）：一个支持技能文件的 AI 助理（原作者用 Muse）；Facebook 扫描用的 `facebook-cli` 是该环境自带的命令行工具，换成你环境里的等价工具即可。

## 安装与使用
```bash
pip install -r requirements.txt
python3 tools/rental_log.py --file "Rental Scan.xlsx" list
python3 tools/rental_log.py --file "Rental Scan.xlsx" add --name "Jane" --type A --url "<帖子链接>"
python3 tools/rental_log.py --file "Rental Scan.xlsx" clear --yes   # 清空记录，必须带 --yes
```
把 `SKILL.md` 和 `tools/` 放进你的 AI 助理技能目录，再配一个每日定时任务即可。

## 注意
联系任何人前都需要房东确认。自动化操作请遵守 Facebook 和小红书的平台条款。

## License
MIT
