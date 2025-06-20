from docx import Document # type: ignore
doc = Document("2025年度中国人民大学科研标兵评审表.docx")

# 表三正好是第 3 张表（索引 2）
tbl = doc.tables[2]

# 假设已有三列表头，两列留空
projects = [
    ("NQ202201", "国产开源 LLM 可信验证平台", 320),
    ("NQ202305", "基于 SQL 的多模态问答引擎", 260),
]

for code, title, budget in projects:
    row_cells = tbl.add_row().cells        # 在尾部插行
    row_cells[0].text = code
    row_cells[1].text = title
    row_cells[2].text = str(budget)
    # 其余单元格不写则保持空白

doc.save("科研标兵评审表_已填写.docx")
