import streamlit as st
from docx import Document
import pandas as pd
import io
import zipfile
from datetime import timedelta
from copy import deepcopy
from docx.enum.section import WD_SECTION
from docx.shared import Pt
from docx.oxml.ns import qn
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

if excel_file:
    st.success("Excel 上传成功！")
    df = pd.read_excel(excel_file)

# 选择生成模式
mode = st.radio(
    "请选择生成方式：",
    ("每个集团单独生成一个 Word", "合并所有集团到一个 Word")
)

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

# 删除前 N 段落
def remove_first_n_paragraphs(doc, n):
    removed = 0
    while removed < n and len(doc.paragraphs) > 0:
        p = doc.paragraphs[0]
        p._element.getparent().remove(p._element)
        removed += 1

# 删除前两个 section（前两页）
def remove_first_two_sections(doc):
    if len(doc.sections) > 1:
        first_sec = doc.sections[0]
        for p in list(doc.paragraphs):
            if p._element.getroottree().getpath(p._element).startswith(
                first_sec._sectPr.getroottree().getpath(first_sec._sectPr)):
                p._element.getparent().remove(p._element)
        if len(doc.sections) > 2:
            second_sec = doc.sections[1]
            for p in list(doc.paragraphs):
                if p._element.getroottree().getpath(p._element).startswith(
                    second_sec._sectPr.getroottree().getpath(second_sec._sectPr)):
                    p._element.getparent().remove(p._element)

# 删除回执函开头第一个表格
def remove_first_table(doc):
    if doc.tables:
        tbl = doc.tables[0]._element
        tbl.getparent().remove(tbl)

# ---------------------------
# 点击生成按钮
# ---------------------------
if st.button("生成 Word"):

    # 根据类型和模式选择模板
    if doc_type == "催缴函" and mode == "每个集团单独生成一个 Word":
        TEMPLATE_PATH = os.path.join(BASE_DIR, "template1.docx")
    elif doc_type == "催缴函" and mode == "合并所有集团到一个 Word":
        TEMPLATE_PATH = os.path.join(BASE_DIR, "template1_2.docx")
    elif doc_type == "回执函" and mode == "每个集团单独生成一个 Word":
        TEMPLATE_PATH = os.path.join(BASE_DIR, "template2.docx")
    else:  # 回执函 & 合并
        TEMPLATE_PATH = os.path.join(BASE_DIR, "template2_2.docx")

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
    else:
        # 合并模式
        combined_doc = Document(TEMPLATE_PATH)
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
                combined_doc.add_section(WD_SECTION.NEW_PAGE)
            first = False
            append_doc(combined_doc, doc)

        # 根据类型单独设置删除行数
        if doc_type == "催缴函":
            remove_first_two_sections(combined_doc)
            remove_first_n_paragraphs(combined_doc, n=len(df)+8)  # 可调整
        else:
            remove_first_table(combined_doc)
            remove_first_two_sections(combined_doc)
            remove_first_n_paragraphs(combined_doc, n=len(df)+4)  # 可调整


        #if len(combined_doc.paragraphs) > 0:
        #    last_p = combined_doc.paragraphs[-1]._element
        #   last_p.getparent().remove(last_p)
        # 删除最后多余空白段落
        while len(combined_doc.paragraphs) > 0:
            last_p = combined_doc.paragraphs[-1]
            if last_p.text.strip() == "":
                last_p._element.getparent().remove(last_p._element)
            else:
                break
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

import streamlit as st
import pandas as pd
import io
import zipfile
import os  # 用于处理文件名和后缀

st.title("这里可以批量重命名")

# ==========================
# 1️⃣ 提供 Excel 模板下载
# ==========================
import os
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
with open(os.path.join(BASE_DIR, "Rename_template.xlsx"), "rb") as f:

    st.download_button(
        "📥 下载Excel 模板（Rename_template.xlsx）",
        data=f,
        file_name="Rename_template.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
st.markdown("Tips:按新名顺序扫描，扫描设置使用自动命名为1、2、3……这样文件原名只需填1、2、3下拉即可（像这样↓）")
import os
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
st.image(os.path.join(BASE_DIR, "example.png"))

# ==========================
# 2️⃣ 上传 Excel
# ==========================
excel_file = st.file_uploader("上传已填写的 Excel 模板", type="xlsx")

if excel_file:
    df = pd.read_excel(excel_file)
    st.success("Excel 上传成功！")
    
    # 检查必须列
    if "文件原名" not in df.columns or "新名" not in df.columns:
        st.error("Excel 必须包含列：'文件原名' 和 '新名'")
    else:
        # 转成字符串并去掉空格和前导单引号，确保匹配成功
        df["文件原名"] = df["文件原名"].astype(str).str.strip().str.lstrip("'")
        df["新名"] = df["新名"].astype(str).str.strip().str.lstrip("'")

        # ==========================
        # 3️⃣ 用户选择需要改名的文件
        # ==========================
        files_to_rename = st.file_uploader(
            "选择需要重命名的文件（可以多选）",
            accept_multiple_files=True
        )

        if files_to_rename:
            st.write("已选择文件：", [f.name for f in files_to_rename])

            if st.button("开始批量重命名"):
                zip_buffer = io.BytesIO()
                renamed_count = 0

                with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zipf:
                    for f in files_to_rename:
                        # 分离文件名和后缀
                        file_base, file_ext = os.path.splitext(f.name)
                        file_base = file_base.strip().lstrip("'")  # 去掉空格和单引号

                        # 匹配 Excel 中的原名
                        match_row = df[df["文件原名"] == file_base]
                        if not match_row.empty:
                            new_base_name = str(match_row["新名"].values[0]).strip().lstrip("'")
                            new_name = new_base_name + file_ext  # 拼回原来的后缀
                            zipf.writestr(new_name, f.getbuffer())
                            renamed_count += 1
                        else:
                            st.warning(f"文件 '{f.name}' 在 Excel 中没有找到对应新名")

                zip_buffer.seek(0)
                st.success(f"重命名完成，共 {renamed_count} 个文件被重命名")
                st.download_button(
                    "📥 下载重命名后的文件（ZIP）",
                    data=zip_buffer,
                    file_name="重命名后的文件.zip",
                    mime="application/zip"
                )



