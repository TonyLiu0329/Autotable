import streamlit as st
import os
import logging
import tempfile
import shutil
from autotable import AutoTable
from extraction import extract_tables_from_docx, extract_content_to_json
from llm_clients import APIClient, OllamaClient
import config

import socket
from datetime import datetime

def get_local_ip():
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        # 连接外部地址以获取准确的局域网IP（不会实际发送数据）
        s.connect(("8.8.8.8", 80))
        ip = s.getsockname()[0]
        s.close()
        return ip
    except Exception:
        return "127.0.0.1"

def setup_logging():
    # 配置根日志记录器
    root_logger = logging.getLogger()
    root_logger.setLevel(logging.INFO)
    
    # 确保有终端输出
    if not any(isinstance(h, logging.StreamHandler) for h in root_logger.handlers):
        console_handler = logging.StreamHandler()
        console_handler.setLevel(logging.INFO)
        root_logger.addHandler(console_handler)

def save_to_history(source_path, target_filename, history_dir="history", max_records=20):
    """保存文件到历史记录，并自动清理旧记录"""
    if not os.path.exists(history_dir):
        os.makedirs(history_dir)
    
    # 复制新文件
    target_path = os.path.join(history_dir, target_filename)
    shutil.copy(source_path, target_path)
    
    # 获取所有 .docx 文件
    files = [f for f in os.listdir(history_dir) if f.endswith(".docx")]
    
    # 如果超过限制，删除最老的
    if len(files) > max_records:
        # 按修改时间排序，最老的在前
        files.sort(key=lambda x: os.path.getmtime(os.path.join(history_dir, x)))
        
        # 计算需要删除的数量
        num_to_delete = len(files) - max_records
        
        for i in range(num_to_delete):
            file_to_delete = files[i]
            try:
                os.remove(os.path.join(history_dir, file_to_delete))
                logging.info(f"Deleted old history file: {file_to_delete}")
            except Exception as e:
                logging.error(f"Failed to delete old history file {file_to_delete}: {e}")

def load_css():
    st.markdown("""
        <style>
        /* 入场动画 */
        @keyframes fadeIn {
            0% { opacity: 0; transform: translateY(20px); }
            100% { opacity: 1; transform: translateY(0); }
        }
        .stApp {
            font-family: 'Source Sans Pro', sans-serif;
            animation: fadeIn 0.8s ease-out;
        }
        /* 标题样式 */
        h1 {
            color: #1E88E5;
            text-align: center;
            font-weight: 700;
            padding-bottom: 20px;
        }
        /* 主按钮样式增强 */
        .stButton>button[kind="primary"] {
            background-color: #1E88E5;
            border: none;
            border-radius: 8px;
            height: 50px;
            font-size: 18px;
            font-weight: 600;
            transition: all 0.3s ease;
        }
        .stButton>button[kind="primary"]:hover {
            background-color: #1565C0;
            box-shadow: 0 4px 6px rgba(0,0,0,0.1);
        }
        /* 下载按钮样式 */
        .stDownloadButton>button {
            border-radius: 8px;
            border: 1px solid #4CAF50;
            color: #4CAF50;
            background-color: white;
            transition: all 0.3s;
        }
        .stDownloadButton>button:hover {
            background-color: #E8F5E9;
            border-color: #2E7D32;
            color: #2E7D32;
        }
        /* 隐藏页脚 */
        footer {visibility: hidden;}
        /* 卡片容器微调 */
        [data-testid="stVerticalBlock"] > [style*="flex-direction: column;"] > [data-testid="stVerticalBlock"] {
            gap: 1rem;
        }
        </style>
    """, unsafe_allow_html=True)

def main():
    st.set_page_config(
        page_title="智能填表助手", 
        page_icon="🤖",
        layout="wide",
        initial_sidebar_state="expanded"
    )
    
    # 加载自定义CSS
    load_css()
    
    # 顶部标题区
    st.title("智能填表助手")
    st.markdown("""
    <div style='text-align: center; color: #666; margin-bottom: 30px;'>
        基于大语言模型的自动化文档填充工具，支持 Word/Excel 智能数据提取与回填
    </div>
    """, unsafe_allow_html=True)

    # 使用说明折叠区
    with st.expander("📖 使用指南 (点击展开)", expanded=False):
        st.markdown("""
        **如何使用：**
        1. **配置模型**：在左侧栏设置 LLM (API 或 Ollama)。
        2. **上传文件**：上传 Word 模板和 Excel/Word 知识库。
        3. **开始处理**：点击按钮，等待 AI 自动分析并填充表格。
        4. **下载结果**：处理完成后下载生成的 Word 文档。
        """)
    
    # 初始化 session state
    if 'processed_file' not in st.session_state:
        st.session_state.processed_file = None # 存储最终结果 (filename, data)
    if 'extracted_file' not in st.session_state:
        st.session_state.extracted_file = None # 存储中间结果 (filename, data)
    
    # --- 侧边栏：配置设置 ---
    with st.sidebar:
        st.header("⚙️ 系统设置")
        
        with st.expander("🧠 LLM 模型配置", expanded=True):
            run_mode = st.radio(
                "运行模式",
                ("api", "ollama"),
                index=0 if config.RUN_MODE == "api" else 1,
                help="选择使用在线 API (如 OpenAI/DeepSeek) 或本地 Ollama 模型"
            )
            
            if run_mode == "api":
                api_base_url = st.text_input("API Base URL", value=config.API_BASE_URL, help="例如: https://api.openai.com/v1")
                api_key = st.text_input("API Key", value=config.API_KEY, type="password", help="在此输入您的 API 密钥")
                api_model = st.text_input("Model Name", value=config.API_MODEL_NAME, help="例如: gpt-4o, deepseek-chat")
            else:
                ollama_host = st.text_input("Ollama Host", value=config.OLLAMA_HOST, help="本地 Ollama 服务地址，通常为 http://localhost:11434")
                ollama_model = st.text_input("Ollama Model", value=config.OLLAMA_MODEL_NAME, help="已拉取的 Ollama 模型名称，如 qwen2.5:14b")
        
        st.info("💡 提示：修改配置后无需重启，直接点击开始处理即可生效。")
        
        st.divider()
        local_ip = get_local_ip()
        st.success(f"📡 **局域网共享已开启**\n\可通过以下地址访问：\n**http://{local_ip}:8501**")
            
    # 设置日志系统
    setup_logging()

    # === 在线上传处理区域 ===
    with st.container(border=True):
        st.subheader("🌐 在线上传处理")
        st.info("ℹ️ 请上传您的文件，处理完成后即可下载结果。")
        
        # 将单选框移至列布局上方，确保下方两个文件上传框对齐
        kb_source_type = st.radio("📚 知识库来源类型", ("上传 Excel 文件", "从 Word 文档提取"), horizontal=True)
        
        col_up1, col_up2 = st.columns(2)
        with col_up1:
            uploaded_word = st.file_uploader("📤 上传 Word 模版 (目标)", type=["docx"])
        with col_up2:
            if kb_source_type == "上传 Excel 文件":
                uploaded_kb = st.file_uploader("📤 上传 Excel 知识库", type=["xlsx"])
                uploaded_kb_is_docx = False
            else:
                uploaded_kb = st.file_uploader("📤 上传 Word 来源文档", type=["docx"], key="upload_kb_docx")
                uploaded_kb_is_docx = True
        
        st.markdown("###")
        start_btn_web = st.button("🚀 开始处理并生成下载", type="primary", use_container_width=True)
    
    # 处理结果显示区域
    result_container = st.container()
    
    if start_btn_web:
        # 重置之前的状态
        st.session_state.processed_file = None
        st.session_state.extracted_file = None
        
        if not uploaded_word or not uploaded_kb:
            st.error("⚠️ 请确保已上传 Word 模版和知识库文件！")
        else:
            try:
                # 创建临时目录
                with tempfile.TemporaryDirectory() as temp_dir:
                    # 保存 Word 模版
                    temp_word_path = os.path.join(temp_dir, uploaded_word.name)
                    with open(temp_word_path, "wb") as f:
                        f.write(uploaded_word.getbuffer())
                    
                    # 处理知识库
                    kb_path = ""
                    if uploaded_kb_is_docx:
                        # 保存来源 Word
                        temp_source_docx = os.path.join(temp_dir, "source.docx")
                        with open(temp_source_docx, "wb") as f:
                            f.write(uploaded_kb.getbuffer())
                        
                        # 提取为 Excel 或 JSON
                        
                        # 初始化 Client (提前初始化，因为提取也可能需要 LLM)
                        if run_mode == "api":
                            client = APIClient(api_base_url, api_key, api_model)
                        else:
                            client = OllamaClient(ollama_host, ollama_model)
                        
                        temp_extracted_kb = os.path.join(temp_dir, "extracted_knowledge.json") # 默认改为 JSON
                        
                        with st.status("🔍 正在智能分析文档...", expanded=True) as status:
                            st.write("正在读取源文档...")
                            # 使用新的智能提取函数
                            extract_success = extract_content_to_json(temp_source_docx, temp_extracted_kb, client)
                            
                            if not extract_success:
                                status.update(label="❌ 数据提取失败", state="error")
                                st.error("从 Word 文档提取数据失败！")
                                st.stop()
                            
                            kb_path = temp_extracted_kb
                            st.write("✅ 数据提取完成，准备填表...")
                            
                            # 读取提取的文件用于下载
                            with open(temp_extracted_kb, "rb") as f:
                                extracted_data = f.read()
                            st.session_state.extracted_file = ("extracted_knowledge.json", extracted_data)
                                
                            st.write("正在填充目标表格...")
                            
                            temp_output_dir = os.path.join(temp_dir, "output")
                                
                            # 运行 AutoTable
                            at = AutoTable(
                                knowledge_base_path=kb_path,
                                word_template_path=temp_word_path,
                                llm_client=client,
                                output_folder=temp_output_dir
                            )
                            success = at.run()
                            
                            if success:
                                status.update(label="✅ 处理完成！", state="complete", expanded=False)
                                # 查找生成的文件
                                generated_files = [f for f in os.listdir(temp_output_dir) if f.endswith(".docx")]
                                if generated_files:
                                        result_file = generated_files[0]
                                        result_path = os.path.join(temp_output_dir, result_file)
                                        
                                        # 保存到历史记录
                                        save_to_history(result_path, result_file)
                                        
                                        # 读取文件用于下载
                                        with open(result_path, "rb") as f:
                                            file_data = f.read()
                                        st.session_state.processed_file = (result_file, file_data)
                                        
                                        st.balloons()
                                        st.success("✅ 文档已生成，请点击下方按钮下载。")
                                else:
                                    status.update(label="❌ 未生成文件", state="error")
                                    st.error("❌ 未找到生成的文件。")
                            else:
                                status.update(label="❌ 处理失败", state="error")
                                st.error("❌ 处理失败，请检查文件内容是否规范。")

                    else:
                        # Excel 流程 (保持 simpler spinner)
                        kb_path = os.path.join(temp_dir, uploaded_kb.name)
                        with open(kb_path, "wb") as f:
                            f.write(uploaded_kb.getbuffer())
                        
                        if run_mode == "api":
                            client = APIClient(api_base_url, api_key, api_model)
                        else:
                            client = OllamaClient(ollama_host, ollama_model)
                        
                        temp_output_dir = os.path.join(temp_dir, "output")
                        
                        with st.status("🔄 正在处理表格...", expanded=True) as status:
                            at = AutoTable(
                                knowledge_base_path=kb_path,
                                word_template_path=temp_word_path,
                                llm_client=client,
                                output_folder=temp_output_dir
                            )
                            success = at.run()
                            
                            if success:
                                status.update(label="✅ 处理完成！", state="complete", expanded=False)
                                generated_files = [f for f in os.listdir(temp_output_dir) if f.endswith(".docx")]
                                if generated_files:
                                    result_file = generated_files[0]
                                    result_path = os.path.join(temp_output_dir, result_file)
                                    
                                    # 保存到历史记录
                                    save_to_history(result_path, result_file)

                                    with open(result_path, "rb") as f:
                                        file_data = f.read()
                                    st.session_state.processed_file = (result_file, file_data)
                                    st.balloons()
                                else:
                                    st.error("❌ 未找到生成的文件。")
                            else:
                                status.update(label="❌ 处理失败", state="error")
                                st.error("❌ 处理失败。")
                            
            except Exception as e:
                st.error(f"处理过程中发生异常: {str(e)}")
    
    # 在主循环中渲染下载按钮（持久化显示）
    if st.session_state.extracted_file or st.session_state.processed_file:
        st.markdown("---")
        st.subheader("📥 结果下载")
        dl_col1, dl_col2 = st.columns(2)
        
        with dl_col1:
            if st.session_state.extracted_file:
                fname, data = st.session_state.extracted_file
                st.download_button(
                    label=f"⬇️ 下载提取的中间数据\n({os.path.splitext(fname)[1]})",
                    data=data,
                    file_name=fname,
                    mime="application/json" if fname.endswith(".json") else "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="btn_dl_extracted",
                    use_container_width=True
                )
            
        with dl_col2:
            if st.session_state.processed_file:
                fname, data = st.session_state.processed_file
                st.download_button(
                    label=f"⬇️ 下载最终结果文档\n{fname}",
                    data=data,
                    file_name=fname,
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    key="btn_dl_final",
                    use_container_width=True,
                    type="primary" 
                )
    
    st.markdown("---")
    with st.expander("📜 历史生成记录 (点击展开)", expanded=False):
        history_dir = "history"
        if not os.path.exists(history_dir):
            os.makedirs(history_dir)
            
        files = [f for f in os.listdir(history_dir) if f.endswith(".docx")]
        # 按修改时间倒序
        files.sort(key=lambda x: os.path.getmtime(os.path.join(history_dir, x)), reverse=True)
        
        if not files:
            st.info("暂无历史记录")
        else:
            st.write(f"共找到 {len(files)} 条记录")
            # 表格展示：文件名 | 大小 | 时间 | 下载
            for f in files:
                file_path = os.path.join(history_dir, f)
                col1, col2, col3 = st.columns([3, 1, 1])
                with col1:
                    st.write(f"📄 {f}")
                with col2:
                    # 显示时间
                    mtime = datetime.fromtimestamp(os.path.getmtime(file_path)).strftime('%Y-%m-%d %H:%M')
                    st.caption(mtime)
                with col3:
                    with open(file_path, "rb") as file:
                        st.download_button(
                            label="⬇️ 下载",
                            data=file,
                            file_name=f,
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                            key=f"dl_hist_{f}"
                        )

if __name__ == "__main__":
    main()
