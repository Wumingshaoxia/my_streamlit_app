import streamlit as st
import pandas as pd
import io
import zipfile
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
import os

st.title("Hi！这里可以生成催缴函/回执函 PDF（Streamlit Cloud 可用）")

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
excel_file = st.file_uploader("上传已填写的 Excel 模板", type="xlsx")

doc_type = st.selectbox("请选择生成类型：", ["催缴函", "回执函"])

if doc_type == "催缴函":
    send_date = st.date_input("请选择发函日期")
    stop_date = st.date_input("请选择支付欠费截止日期")
else:
    receipt_date = st.date_input("请选择回执日期")

if excel_file:
    df = pd.read_excel(excel_file)
    st.success("Excel 上传成功！")

    mode = st.radio(
        "请选择生成方式：",
        ("每个集团单独生成 PDF", "合并所有集团生成 PDF")
    )

    # ============================
    # PDF 生成函数
    # ============================
    def generate_pdf(path, placeholders, doc_type):
        """
        用 ReportLab 生成 PDF
        path: 保存路径
        placeholders: 字典，替换内容
        """
        c = canvas.Canvas(path, pagesize=A4)
        width, height = A4

        # 注册中文字体（需要项目目录下有 SimSun.ttf 或其他中文字体）
        font_path = os.path.join(BASE_DIR, "SimSun.ttf")
        pdfmetrics.registerFont(TTFont("SimSun", font_path))
        c.setFont("SimSun", 12)

        y = height - 100  # 顶部开始

        c.drawString(50, y, f"{doc_type}")
        y -= 40
        for k, v in placeholders.items():
            c.drawString(50, y, f"{k}：{v}")
            y -= 25

        c.showPage()
        c.save()

    # ============================
    # 点击生成按钮
    # ============================
    if st.button("生成 PDF"):
        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zipf:
            if mode == "每个集团单独生成 PDF":
                for _, row in df.iterrows():
                    placeholders = {
                        "集团名称": row["集团名称"],
                        "客户经理": row["客户经理"],
                        "客户经理手机号": row["客户经理手机号"],
                        "共计欠费": row["共计欠费"]
                    }
                    if doc_type == "催缴函":
                        placeholders.update({
                            "逾期欠费金额": row["逾期欠费金额"],
                            "违约金": row["违约金"],
                            "发函日期": send_date.strftime("%Y年%m月%d日"),
                            "支付欠费截止日期": stop_date.strftime("%Y年%m月%d日")
                        })
                    else:
                        placeholders.update({
                            "回执日期": receipt_date.strftime("%Y年%m月%d日")
                        })

                    pdf_name = f"{doc_type}_{row['集团名称']}.pdf"
                    pdf_path = os.path.join("/tmp", pdf_name)
                    generate_pdf(pdf_path, placeholders, doc_type)
                    zipf.write(pdf_path, pdf_name)

            else:  # 合并模式
                pdf_name = f"合并{doc_type}.pdf"
                pdf_path = os.path.join("/tmp", pdf_name)
                c = canvas.Canvas(pdf_path, pagesize=A4)
                width, height = A4
                font_path = os.path.join(BASE_DIR, "SimSun.ttf")
                pdfmetrics.registerFont(TTFont("SimSun", font_path))
                c.setFont("SimSun", 12)

                for _, row in df.iterrows():
                    y = height - 100
                    c.drawString(50, y, f"{doc_type}")
                    y -= 40
                    placeholders = {
                        "集团名称": row["集团名称"],
                        "客户经理": row["客户经理"],
                        "客户经理手机号": row["客户经理手机号"],
                        "共计欠费": row["共计欠费"]
                    }
                    if doc_type == "催缴函":
                        placeholders.update({
                            "逾期欠费金额": row["逾期欠费金额"],
                            "违约金": row["违约金"],
                            "发函日期": send_date.strftime("%Y年%m月%d日"),
                            "支付欠费截止日期": stop_date.strftime("%Y年%m月%d日")
                        })
                    else:
                        placeholders.update({
                            "回执日期": receipt_date.strftime("%Y年%m月%d日")
                        })
                    for k, v in placeholders.items():
                        c.drawString(50, y, f"{k}：{v}")
                        y -= 25
                    c.showPage()
                c.save()
                zipf.write(pdf_path, pdf_name)

        zip_buffer.seek(0)
        st.success("PDF 生成成功！点击下载 ZIP 文件👇")
        st.download_button(
            f"下载全部 {doc_type} PDF（ZIP）",
            data=zip_buffer,
            file_name=f"{doc_type}_合集.zip",
            mime="application/zip"
        )
