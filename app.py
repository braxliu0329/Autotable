import streamlit as st
import os
import json
import logging
import tempfile
import shutil
from autotable import AutoTable
from extraction import extract_tables_from_docx, extract_content_to_json, save_aspose_json
from llm_clients import APIClient, OllamaClient
import config

import socket
from datetime import datetime
import time

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

def save_to_history(source_path, target_filename, history_dir="history", max_records=10):
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

class StreamlitLogHandler(logging.Handler):
    def __init__(self, placeholder):
        super().__init__()
        self.placeholder = placeholder
        self.log_buffer = []

    def emit(self, record):
        # 剔除失败日志 (ERROR及以上级别)
        if record.levelno >= logging.ERROR:
            return
            
        msg = self.format(record)
        
        # 额外过滤：如果日志内容包含 'error' (忽略大小写)，也不显示
        if "error" in msg.lower():
            return

        self.log_buffer.append(msg)
        # 实时更新显示最新的 50 行日志，避免页面卡顿
        content = "\n".join(self.log_buffer[-50:])
        self.placeholder.code(content, language="text")

def load_css():
    st.markdown("""
        <style>
        /* 动画定义 */
        @keyframes fadeIn {
            from { opacity: 0; transform: translateY(10px); }
            to { opacity: 1; transform: translateY(0); }
        }
        
        @keyframes slideInRight {
            from { opacity: 0; transform: translateX(20px); }
            to { opacity: 1; transform: translateX(0); }
        }

        /* 增强的容器样式 - 适配主题 */
        .step-container {
            animation: slideInRight 0.4s ease-out;
            padding: 30px;
            background-color: var(--secondary-background-color);
            border-radius: 15px;
            margin-bottom: 25px;
            border: 1px solid rgba(128, 128, 128, 0.2);
            box-shadow: 0 4px 15px rgba(0,0,0,0.1);
        }

        /* 标题样式 */
        h1 {
            color: var(--primary-color);
            text-align: center;
            font-weight: 800;
            padding-bottom: 10px;
            font-size: 2.5rem;
            text-shadow: 0 2px 4px rgba(0,0,0,0.1);
        }
        
        .description-text {
            text-align: center;
            color: var(--text-color);
            opacity: 0.8;
            margin-bottom: 40px;
            font-size: 1.1rem;
        }

        /* 步骤指示器美化 */
        .step-indicator {
            display: flex;
            justify-content: center;
            margin-bottom: 40px;
            font-weight: 600;
            color: var(--text-color);
            opacity: 0.6;
            position: relative;
        }
        
        /* 连接线 */
        .step-indicator::before {
            content: '';
            position: absolute;
            top: 50%;
            left: 20%;
            right: 20%;
            height: 2px;
            background-color: var(--text-color);
            opacity: 0.2;
            z-index: 0;
            transform: translateY(-50%);
        }

        .step-indicator .step {
            margin: 0 30px;
            padding: 10px 20px;
            position: relative;
            z-index: 1;
            background-color: var(--background-color);
            border-radius: 20px;
            transition: all 0.3s;
            border: 1px solid rgba(128, 128, 128, 0.3);
        }
        
        .step-indicator .active {
            color: var(--primary-color);
            background-color: var(--secondary-background-color);
            border: 1px solid var(--primary-color);
            box-shadow: 0 0 10px rgba(77, 166, 255, 0.3);
        }
        
        .step-indicator .completed {
            color: #4caf50;
            background-color: var(--secondary-background-color);
            border: 1px solid #4caf50;
        }
        
        /* 按钮增强 */
        .stButton>button {
            border-radius: 10px;
            height: 50px;
            font-weight: 600;
            transition: all 0.2s;
            font-size: 16px;
        }
        .stButton>button:hover {
            transform: translateY(-2px);
            box-shadow: 0 4px 12px rgba(0,0,0,0.2);
        }

        /* 隐藏页脚 */
        footer {visibility: hidden;}
        </style>
    """, unsafe_allow_html=True)

def render_step_indicator(current_step):
    steps = [
        {"id": 1, "label": "1. 选择来源"},
        {"id": 2, "label": "2. 上传知识库"},
        {"id": 3, "label": "3. 填表生成"}
    ]
    
    html = '<div class="step-indicator">'
    for step in steps:
        status_class = ""
        icon = ""
        if current_step == step["id"]:
            status_class = "active"
            icon = "🔷"
        elif current_step > step["id"]:
            status_class = "completed"
            icon = "✅"
        else:
            icon = "⚪"
        
        html += f'<div class="step {status_class}">{icon} {step["label"]}</div>'
    html += '</div>'
    
    st.markdown(html, unsafe_allow_html=True)

def main():
    st.set_page_config(
        page_title="智能填表助手", 
        page_icon="🤖",
        layout="wide",
        initial_sidebar_state="expanded"
    )
    
    load_css()
    
    # 初始化 session state
    if 'current_step' not in st.session_state:
        st.session_state.current_step = 1
    if 'kb_source_type' not in st.session_state:
        st.session_state.kb_source_type = "上传 Excel 文件"
    if 'kb_file_data' not in st.session_state:
        st.session_state.kb_file_data = None # {'name': str, 'data': bytes}
    if 'processed_file' not in st.session_state:
        st.session_state.processed_file = None 
    if 'intermediate_files' not in st.session_state:
        st.session_state.intermediate_files = {} # {'filename': bytes} 
        
    setup_logging()

    # --- 侧边栏 ---
    with st.sidebar:
        st.header("⚙️ 系统设置")
        with st.expander("🧠 LLM 模型配置", expanded=True):
            run_mode = st.radio(
                "运行模式",
                ("api", "ollama"),
                index=0 if config.RUN_MODE == "api" else 1,
                help="选择使用在线 API 或本地 Ollama 模型"
            )
            
            if run_mode == "api":
                api_base_url = st.text_input("API Base URL", value=config.API_BASE_URL)
                api_key = st.text_input("API Key", value=config.API_KEY, type="password")
                api_model = st.text_input("Model Name", value=config.API_MODEL_NAME)
            else:
                ollama_host = st.text_input("Ollama Host", value=config.OLLAMA_HOST)
                ollama_model = st.text_input("Ollama Model", value=config.OLLAMA_MODEL_NAME)
            
            # 超时设置
            timeout_seconds = st.slider(
                "请求超时时间 (秒)", 
                min_value=60, 
                max_value=600, 
                value=300, 
                step=30,
                help="设置 API 请求的最大等待时间。处理大表格或大知识库时，建议调大此值。"
            )
        
        st.divider()
        local_ip = get_local_ip()
        st.success(f"📡 局域网访问地址：\n**http://{local_ip}:8501**")
        
        
        st.markdown("---")
        with st.expander("📖 使用指南", expanded=False):
            st.markdown("""
            1. **选择来源**：Excel 适合结构化数据，Word 适合提取简历等非结构化文档。
            2. **上传知识库**：上传包含数据的文件。
            3. **上传模板**：上传需要填充的 Word 模板 (.docx)，系统会自动识别下划线和表格进行填充。
            """)

    # --- 主体区域 ---
    st.title("智能填表助手")
    st.markdown("""
    <div class='description-text'>
        基于大语言模型的自动化文档填充工具，支持 Word/Excel 智能数据提取与回填<br>
        让 AI 帮你完成繁琐的表格填写工作
    </div>
    """, unsafe_allow_html=True)
    
    render_step_indicator(st.session_state.current_step)

    # 容器用于页面切换
    placeholder = st.empty()

    # === STEP 1: 选择来源 ===
    if st.session_state.current_step == 1:
        with placeholder.container():
            # 修正方案：Streamlit 原生 st.container(border=True) 是最佳选择，能产生带边框的容器。
            # 配合自定义 CSS 修改这个原生容器的样式。
            
            with st.container(border=True):
                st.subheader("步骤 1: 选择知识库来源")
                st.info("💡 请选择您的数据来源格式。Excel 适合结构化数据，Word 适合非结构化文档提取。")
                
                kb_type = st.radio(
                    "知识库类型", 
                    ("上传 Excel 文件", "从 Word 文档提取"), 
                    index=0 if st.session_state.kb_source_type == "上传 Excel 文件" else 1,
                    horizontal=True
                )
                
                st.markdown("<br>", unsafe_allow_html=True)
                col1, col2 = st.columns([1, 4])
                with col1:
                    if st.button("下一步 ➡️", type="primary", use_container_width=True):
                        st.session_state.kb_source_type = kb_type
                        st.session_state.current_step = 2
                        st.rerun()

    # === STEP 2: 上传知识库 ===
    elif st.session_state.current_step == 2:
        with placeholder.container():
            with st.container(border=True):
                st.subheader(f"步骤 2: {st.session_state.kb_source_type}")
                
                uploaded_kb = None
                if st.session_state.kb_source_type == "上传 Excel 文件":
                    uploaded_kb = st.file_uploader("📤 上传 Excel (.xlsx) 文件", type=["xlsx"])
                else:
                    uploaded_kb = st.file_uploader("📤 上传 Word (.docx) 来源文档", type=["docx"])

                st.markdown("<br>", unsafe_allow_html=True)
                col_back, col_next = st.columns([1, 4])
                
                with col_back:
                    if st.button("⬅️ 上一步", use_container_width=True):
                        st.session_state.current_step = 1
                        st.rerun()
                
                with col_next:
                    # 检查是否已有文件
                    has_file = uploaded_kb is not None
                    if st.button("下一步 ➡️", type="primary", disabled=not has_file, use_container_width=True):
                        if uploaded_kb:
                            # 保存文件内容到 session state
                            st.session_state.kb_file_data = {
                                "name": uploaded_kb.name,
                                "data": uploaded_kb.getvalue(),
                                "type": "docx" if st.session_state.kb_source_type == "从 Word 文档提取" else "xlsx"
                            }
                            st.session_state.current_step = 3
                            st.rerun()
                
                if not uploaded_kb and st.session_state.kb_file_data:
                    st.info(f"✅ 已缓存文件: {st.session_state.kb_file_data['name']}")

    # === STEP 3: 上传模板并运行 ===
    elif st.session_state.current_step == 3:
        with placeholder.container():
            with st.container(border=True):
                st.subheader("步骤 3: 上传模板并生成")
                
                # 显示已就绪的知识库
                if st.session_state.kb_file_data:
                    st.success(f"✅ 知识库已就绪: {st.session_state.kb_file_data['name']}")
                else:
                    st.error("❌ 知识库丢失，请返回重新上传")

                uploaded_template = st.file_uploader("📤 上传 Word (.docx) 模板文件", type=["docx"])
                
                st.markdown("<br>", unsafe_allow_html=True)
                col_back, col_run = st.columns([1, 4])
                
                with col_back:
                    if st.button("⬅️ 上一步", use_container_width=True):
                        st.session_state.current_step = 2
                        st.rerun()
                
                with col_run:
                    run_disabled = uploaded_template is None or st.session_state.kb_file_data is None
                    if st.button("🚀 开始处理", type="primary", disabled=run_disabled, use_container_width=True):
                        # 执行处理逻辑
                        with st.spinner("正在初始化环境..."):
                            try:
                                with tempfile.TemporaryDirectory() as temp_dir:
                                    # 1. 恢复知识库文件
                                    kb_info = st.session_state.kb_file_data
                                    kb_path = os.path.join(temp_dir, kb_info["name"])
                                    with open(kb_path, "wb") as f:
                                        f.write(kb_info["data"])
                                    
                                    # 2. 保存模板文件
                                    temp_word_path = os.path.join(temp_dir, uploaded_template.name)
                                    with open(temp_word_path, "wb") as f:
                                        f.write(uploaded_template.getbuffer())

                                    # 3. 初始化 LLM
                                    if run_mode == "api":
                                        client = APIClient(api_base_url, api_key, api_model, timeout=timeout_seconds)
                                    else:
                                        client = OllamaClient(ollama_host, ollama_model, timeout=timeout_seconds)

                                    # 4. 如果是 Word 知识库，先提取
                                    final_kb_path = kb_path
                                    
                                    # 4.1 始终生成知识库的中间文件（如果是 Word）
                                    if kb_info["type"] == "docx":
                                        json_kb_path = os.path.join(temp_dir, "kb_extracted.json")
                                        aspose_kb_path = os.path.join(temp_dir, "kb_aspose.json")
                                        
                                        with st.status("🔍 正在分析知识库内容...", expanded=False) as status:
                                            # 生成 Aspose 结构文件 (供下载)
                                            save_aspose_json(kb_path, aspose_kb_path)
                                            
                                            # 智能提取内容 (供填表用)
                                            extract_success = extract_content_to_json(kb_path, json_kb_path, client)
                                            if not extract_success:
                                                status.update(label="❌ 提取失败", state="error")
                                                st.error("知识库提取失败")
                                                st.stop()
                                            final_kb_path = json_kb_path
                                            
                                            # 保存到 Session State 供下载
                                            if os.path.exists(aspose_kb_path):
                                                with open(aspose_kb_path, "rb") as f:
                                                    st.session_state.intermediate_files["kb_aspose.json"] = f.read()
                                            if os.path.exists(json_kb_path):
                                                with open(json_kb_path, "rb") as f:
                                                    st.session_state.intermediate_files["kb_content.json"] = f.read()

                                            # 实时显示提取的 JSON 数据
                                            try:
                                                with open(final_kb_path, "r", encoding="utf-8") as f:
                                                    json_data = json.load(f)
                                                    st.markdown("### 📄 提取到的数据")
                                                    st.json(json_data, expanded=False)
                                            except Exception as e:
                                                st.warning(f"无法显示中间数据: {e}")

                                    # 4.2 生成模板的中间文件
                                    template_aspose_path = os.path.join(temp_dir, "template_aspose.json")
                                    save_aspose_json(temp_word_path, template_aspose_path)
                                    if os.path.exists(template_aspose_path):
                                        with open(template_aspose_path, "rb") as f:
                                            st.session_state.intermediate_files["template_aspose.json"] = f.read()

                                    # 5. 运行 AutoTable
                                    temp_output_dir = os.path.join(temp_dir, "output")
                                    with st.status("🤖 正在智能填表...", expanded=False) as status:
                                        # 创建日志显示区域
                                        st.write("📋 执行日志")
                                        log_placeholder = st.empty()
                                        
                                        # 配置日志处理器
                                        log_handler = StreamlitLogHandler(log_placeholder)
                                        log_handler.setLevel(logging.INFO)
                                        formatter = logging.Formatter('%(asctime)s - %(message)s', datefmt='%H:%M:%S')
                                        log_handler.setFormatter(formatter)
                                        
                                        # 添加到根日志记录器
                                        root_logger = logging.getLogger()
                                        root_logger.addHandler(log_handler)
                                        
                                        try:
                                            at = AutoTable(final_kb_path, temp_word_path, client, temp_output_dir)
                                            if at.run():
                                                status.update(label="✅ 完成！", state="complete")
                                                # 处理结果
                                                generated_files = [f for f in os.listdir(temp_output_dir) if f.endswith(".docx")]
                                                if generated_files:
                                                    result_file = generated_files[0]
                                                    result_path = os.path.join(temp_output_dir, result_file)
                                                    save_to_history(result_path, result_file)
                                                    with open(result_path, "rb") as f:
                                                        st.session_state.processed_file = (result_file, f.read())
                                                else:
                                                    st.error("未生成文件")
                                            else:
                                                status.update(label="❌ 失败", state="error")
                                                st.error("填表过程出错")
                                        finally:
                                            # 清理日志处理器
                                            root_logger.removeHandler(log_handler)
                                            
                            except Exception as e:
                                st.error(f"发生错误: {str(e)}")

            
            # 显示下载区域 (仅在 Step 3 显示)
            if st.session_state.processed_file:
                with st.container(border=True):
                    st.success("✅ 文档生成成功！")
                    
                    # 结果下载
                    fname, data = st.session_state.processed_file
                    st.download_button(
                        label=f"⬇️ 下载结果: {fname}",
                        data=data,
                        file_name=fname,
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        type="primary",
                        use_container_width=True
                    )
                    
                    # 中间文件下载区域
                    if st.session_state.intermediate_files:
                        st.markdown("---")
                        st.subheader("🛠️ 调试/中间文件下载")
                        
                        # 使用 columns 展示下载按钮
                        int_cols = st.columns(len(st.session_state.intermediate_files))
                        
                        for idx, (filename, file_data) in enumerate(st.session_state.intermediate_files.items()):
                            # 根据文件类型选择列
                            col_idx = idx % 3
                            with int_cols[col_idx]:
                                label_prefix = "📄"
                                if "aspose" in filename:
                                    label_prefix = "🏗️" # 结构文件
                                elif "content" in filename:
                                    label_prefix = "📝" # 内容文件
                                    
                                st.download_button(
                                    label=f"{label_prefix} {filename}",
                                    data=file_data,
                                    file_name=filename,
                                    mime="application/json",
                                    use_container_width=True
                                )

    # --- 底部历史记录 (始终显示) ---
    st.markdown("---")
    with st.expander("📜 历史生成记录", expanded=False):
        history_dir = "history"
        if os.path.exists(history_dir):
            files = [f for f in os.listdir(history_dir) if f.endswith(".docx")]
            files.sort(key=lambda x: os.path.getmtime(os.path.join(history_dir, x)), reverse=True)
            for f in files:
                col1, col2 = st.columns([4, 1])
                col1.text(f"📄 {f}")
                with open(os.path.join(history_dir, f), "rb") as file:
                    col2.download_button("下载", file, file_name=f, key=f"hist_{f}")

if __name__ == "__main__":
    main()
