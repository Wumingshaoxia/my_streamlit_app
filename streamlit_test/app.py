import streamlit as st
from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.text import WD_BREAK
from docx.shared import Pt
from docx.oxml.ns import qn
import pandas as pd
import io
import zipfile
from datetime import timedelta
from copy import deepcopy
import os

st.title("Hi！这里可以生成催缴函/回执函")

# =============================
# 提供 Excel 模板下载
# =============================
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
with open(os.path.join(BASE_DIR, "催缴函-template.xlsx"), "rb") as f:
    st.download_button(
        "📥 下载 Excel 模板（催缴函-template.xlsx）",
        data=f,
        file_name="催缴函-template.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

# =============================
# 上传 Excel
# =============================
excel_file = st.file_uploader("上传已填写的Excel模板", type="xlsx")

# 选择生成类型
doc_type = st.selectbox("请选择生成类型：", ["催缴函", "回执函"])

# 日期选择器
if doc_type == "催缴函":
    send_date = st.date_input("请选择发函日期")
    stop_date = st.date_input("请选择支付欠费截止日期")
    end_date = stop_date + timedelta(days=1)
else:
    receipt_date = st.date_input("请选择回执日期")

# ---------------------------
# 替换占位符函数
# ---------------------------
def replace_placeholder(doc, placeholders: dict, font_name=None, font_size=None):
    for p in doc.paragraphs:
        for key, value in placeholders.items():
            if key in p.text:
                for run in p.runs:
                    if key in run.text:
                        run.text = run.text.replace(key, str(value))
                        if font_name:
                            run.font.name = font_name
                            run._element.rPr.rFonts.set(qn('w:eastAsia'), font_name)
                        if font_size:
                            run.font.size = Pt(font_size)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for key, value in placeholders.items():
                    if key in cell.text:
                        for p in cell.paragraphs:
                            for run in p.runs:
                                if key in run.text:
                                    run.text = run.text.replace(key, str(value))
                                    if font_name:
                                        run.font.name = font_name
                                        run._element.rPr.rFonts.set(qn('w:eastAsia'), font_name)
                                    if font_size:
                                        run.font.size = Pt(font_size)

# ---------------------------
# 复制文档内容
# ---------------------------
def append_doc(target, source):
    for element in source.element.body:
        target.element.body.append(deepcopy(element))

if excel_file:
    st.success("Excel 上传成功！")
    df = pd.read_excel(excel_file)

    # 选择生成模式
    mode = st.radio(
        "请选择生成方式：",
        ("每个集团单独生成一个 Word", "合并所有集团到一个 Word")
    )

    if st.button("生成 Word"):
        TEMPLATE1_PATH = os.path.join(BASE_DIR, "template1.docx")
        TEMPLATE2_PATH = os.path.join(BASE_DIR, "template2.docx")
        TEMPLATE_PATH = TEMPLATE1_PATH if doc_type == "催缴函" else TEMPLATE2_PATH

        # ==========================
        # 单独生成
        # ==========================
        if mode == "每个集团单独生成一个 Word":
            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zipf:
                for _, row in df.iterrows():
                    doc = Document(TEMPLATE_PATH)
                    if doc_type == "催缴函":
                        placeholders = {
                            "{{集团名称}}": row["集团名称"],
                            "{{客户经理}}": row["客户经理"],
                            "{{客户经理手机号}}": row["客户经理手机号"],
                            "{{逾期欠费金额}}": row["逾期欠费金额"],
                            "{{违约金}}": row["违约金"],
                            "{{共计欠费}}": row["共计欠费"],
                            "{{发函日期}}": send_date.strftime("%Y年%m月%d日"),
                            "{{支付欠费截止日期}}": stop_date.strftime("%Y年%m月%d日"),
                            "{{终止业务日期}}": end_date.strftime("%Y年%m月%d日"),
                        }
                        replace_placeholder(doc, placeholders)
                    else:
                        placeholders = {
                            "{{集团名称}}": row["集团名称"],
                            "{{客户经理}}": row["客户经理"],
                            "{{客户经理手机号}}": row["客户经理手机号"],
                            "{{共计欠费}}": row["共计欠费"],
                            "{{回执日期}}": receipt_date.strftime("%Y年%m月%d日"),
                        }
                        replace_placeholder(doc, placeholders, font_name="宋体", font_size=13)

                    file_buffer = io.BytesIO()
                    doc.save(file_buffer)
                    file_buffer.seek(0)
                    filename = f"{doc_type}_{row['集团名称']}.docx"
                    zipf.writestr(filename, file_buffer.getvalue())

            zip_buffer.seek(0)
            st.success("生成成功！点击下载 ZIP 文件👇")
            st.download_button(
                f"下载全部 {doc_type} Word（ZIP）",
                data=zip_buffer,
                file_name=f"{doc_type}合集.zip",
                mime="application/zip",
            )

        # ==========================
        # 合并生成
        # ==========================
        else:
            combined_doc = Document()  # 新建空文档
            first = True
            for _, row in df.iterrows():
                doc = Document(TEMPLATE_PATH)
                if doc_type == "催缴函":
                    placeholders = {
                        "{{集团名称}}": row["集团名称"],
                        "{{客户经理}}": row["客户经理"],
                        "{{客户经理手机号}}": row["客户经理手机号"],
                        "{{逾期欠费金额}}": row["逾期欠费金额"],
                        "{{违约金}}": row["违约金"],
                        "{{共计欠费}}": row["共计欠费"],
                        "{{发函日期}}": send_date.strftime("%Y年%m月%d日"),
                        "{{支付欠费截止日期}}": stop_date.strftime("%Y年%m月%d日"),
                        "{{终止业务日期}}": end_date.strftime("%Y年%m月%d日"),
                    }
                    replace_placeholder(doc, placeholders)
                else:
                    placeholders = {
                        "{{集团名称}}": row["集团名称"],
                        "{{客户经理}}": row["客户经理"],
                        "{{客户经理手机号}}": row["客户经理手机号"],
                        "{{共计欠费}}": row["共计欠费"],
                        "{{回执日期}}": receipt_date.strftime("%Y年%m月%d日"),
                    }
                    replace_placeholder(doc, placeholders, font_name="宋体", font_size=13)

                if not first:
                    # 自动分页
                    combined_doc.add_paragraph().add_run().add_break(WD_BREAK.PAGE)
                first = False

                append_doc(combined_doc, doc)

            output_buffer = io.BytesIO()
            combined_doc.save(output_buffer)
            output_buffer.seek(0)
            st.success(f"合并 {doc_type} Word 生成成功！点击下载👇")
            st.download_button(
                f"下载合并版 {doc_type} Word",
                data=output_buffer,
                file_name=f"合并{doc_type}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            )
