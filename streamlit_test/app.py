        else:
            # ============================
            # ⭐ 改为 PDF 合并模式（回执函专用）
            # ============================
            from tempfile import NamedTemporaryFile
            from docx2pdf import convert
            from PyPDF2 import PdfMerger

            pdf_files = []  # 存储每个集团生成的 PDF 路径

            for _, row in df.iterrows():

                # 1️⃣ 加载模板
                doc = Document(TEMPLATE_PATH)

                # 2️⃣ 填充占位符
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

                # 3️⃣ 保存单个 Word
                with NamedTemporaryFile(delete=False, suffix=".docx") as tmp_word:
                    tmp_word_path = tmp_word.name
                    doc.save(tmp_word_path)

                # 4️⃣ Word → PDF
                pdf_path = tmp_word_path.replace(".docx", ".pdf")
                convert(tmp_word_path, pdf_path)
                pdf_files.append(pdf_path)

            # ============================
            # ⭐ 合并所有 PDF
            # ============================
            merger = PdfMerger()
            for pdf in pdf_files:
                merger.append(pdf)

            # 输出最终合并 PDF
            merged_pdf_path = os.path.join(BASE_DIR, f"合并{doc_type}.pdf")
            merger.write(merged_pdf_path)
            merger.close()

            # 提供下载
            with open(merged_pdf_path, "rb") as f:
                st.success(f"合并 {doc_type} PDF 生成成功！格式不会串行！")
                st.download_button(
                    f"下载合并版 {doc_type}（PDF）",
                    data=f,
                    file_name=f"合并{doc_type}.pdf",
                    mime="application/pdf",
                )
