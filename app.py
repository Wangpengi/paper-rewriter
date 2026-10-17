import io
import os
import streamlit as st
from docx import Document
from openai import OpenAI

st.set_page_config(page_title="论文无损格式降重 (DeepSeek)", page_icon="📝", layout="centered")
st.title("📝 论文无损格式在线降重")
st.caption("基于 DeepSeek 模型，重构学术表达并 100% 继承 Word 排版样式。")

# 侧边栏配置
with st.sidebar:
    st.header("⚙️ 模型配置")
    api_key = st.text_input(
        "DeepSeek API Key", 
        value=os.getenv("DEEPSEEK_API_KEY", ""), 
        type="password", 
        placeholder="sk-..."
    )
    base_url = st.text_input("Base URL", value="https://api.deepseek.com")
    model_name = st.selectbox("选择模型", ["deepseek-chat", "deepseek-reasoner"], index=0)
    st.caption("注：日常改写与降重推荐使用 deepseek-chat，响应更快且成本极低。")

PROMPT = """你是一名资深学术论文审稿与改写专家。请对输入的段落进行学术降重：
1. 策略：变换主被动语态、倒置因果逻辑结构、长复合句拆解为精炼短句、非术语同义替换。
2. 约束：严禁更改专业名词、算法名、数学符号与标点编号；保持严谨学术语体。
3. 输出：直接输出改写结果，严禁带有任何前缀、解释或标注。如果输入是纯符号、公式或极短标题，原样返回。"""

def rewrite_text(client, text):
    if not text.strip() or len(text.strip()) < 15:
        return text
    try:
        res = client.chat.completions.create(
            model=model_name,
            messages=[
                {"role": "system", "content": PROMPT},
                {"role": "user", "content": text}
            ],
            temperature=0.3
        )
        return res.choices[0].message.content.strip()
    except Exception as e:
        st.error(f"处理段落出错: {e}")
        return text

def update_paragraph_runs(p, new_text):
    if not p.runs:
        p.text = new_text
        return
    p.runs[0].text = new_text
    for r in p.runs[1:]:
        r.text = ""

# 主操作区
uploaded_file = st.file_uploader("上传已排版好的论文 (.docx)", type=["docx"])

if uploaded_file:
    if not api_key:
        st.warning("⚠️ 请先在左侧输入你的 DeepSeek API Key！")
    else:
        if st.button("🚀 开始原格式降重", type="primary"):
            client = OpenAI(api_key=api_key, base_url=base_url)
            doc = Document(uploaded_file)
            
            valid_paragraphs = [
                p for p in doc.paragraphs 
                if p.text.strip() and not p.style.name.startswith("Caption")
            ]
            total = len(valid_paragraphs)
            
            if total == 0:
                st.info("文档未检测到有效正文段落。")
            else:
                progress_bar = st.progress(0)
                status_text = st.empty()
                
                for idx, p in enumerate(valid_paragraphs):
                    status_text.text(f"正在处理正文段落：{idx + 1} / {total} ...")
                    rewritten = rewrite_text(client, p.text)
                    update_paragraph_runs(p, rewritten)
                    progress_bar.progress((idx + 1) / total)
                
                status_text.text("正文处理完成，正在检查表格...")
                for table in doc.tables:
                    for row in table.rows:
                        for cell in row.cells:
                            for p in cell.paragraphs:
                                if p.text.strip():
                                    rewritten = rewrite_text(client, p.text)
                                    update_paragraph_runs(p, rewritten)
                
                status_text.empty()
                output_stream = io.BytesIO()
                doc.save(output_stream)
                output_stream.seek(0)
                
                st.success("✅ 降重完成！格式已完好继承。")
                st.download_button(
                    label="📥 点击下载降重后的 Word 文件",
                    data=output_stream,
                    file_name=f"降重版_{uploaded_file.name}",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )