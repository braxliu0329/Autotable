import os
import pandas as pd
import logging
import subprocess
import re
import json
from io import StringIO

logger = logging.getLogger(__name__)

def get_node_wrapper_path():
    """
    Get the path to the Node.js wrapper script.
    """
    current_dir = os.path.dirname(os.path.abspath(__file__))
    project_root = os.path.dirname(current_dir)
    return os.path.join(
        project_root,
        "skills", "custom", "docx-extraction",
        "scripts", "server", "index.cjs"
    )

def extract_structure_using_aspose(docx_path):
    """
    Call the Node.js wrapper to extract STRUCTURED content from DOCX (JSON mode).
    Returns a Python dict containing 'outline', 'pages', 'tables' etc.
    """
    wrapper_path = get_node_wrapper_path()
    if not os.path.exists(wrapper_path):
        raise FileNotFoundError(f"Node.js wrapper not found at {wrapper_path}")
        
    try:
        # Call node wrapper with "json" argument
        # Note: Using shell=True for Windows compatibility with node command
        # Ensure 'node' is in system PATH
        # logger.debug(f"Running node {wrapper_path} {docx_path} --json")
        print(f"[Skill] Invoking Docx Extraction Skill: {wrapper_path}")  # Console output
        result = subprocess.run(
            ["node", wrapper_path, docx_path, "--json"], 
            capture_output=True, 
            text=True, 
            encoding='utf-8'
        )
        
        if result.returncode != 0:
            logger.error(f"Node stderr: {result.stderr}")
            raise Exception(f"Extraction failed with code {result.returncode}: {result.stderr}")
            
        try:
            # logger.debug(f"Node stdout preview: {result.stdout[:200]}")
            raw_data = json.loads(result.stdout)
            
            # Unpack response wrapper if present
            if isinstance(raw_data, dict) and 'data' in raw_data and 'code' in raw_data:
                if raw_data['code'] != 200:
                    raise Exception(f"Aspose service error: {raw_data.get('message')}")
                return raw_data['data']
                
            return raw_data
        except json.JSONDecodeError as je:

            raise Exception(f"Failed to parse Aspose JSON output: {je}. Output preview: {result.stdout[:200]}")
        
    except Exception as e:
        logger.error(f"Aspose structural extraction failed: {str(e)}")
        raise e

def extract_text_using_aspose(docx_path):
    """
    Re-implemented to use the structural extraction (Node.js wrapper)
    and flatten the content to text.
    """
    try:
        structure = extract_structure_using_aspose(docx_path)
        text_lines = []
        tables = structure.get("tables", [])
        
        for page in structure.get("pages", []):
            for line in page.get("textContent", []):
                # Check for table placeholder
                match = re.match(r"\[TABLE_(\d+)\]", line)
                if match:
                    # In text mode, we might want to skip tables or just keep the placeholder,
                    # or maybe try to convert HTML to text? 
                    # For now, let's keep the placeholder to indicate a table was there.
                    # Or better, append a marker.
                    text_lines.append(line)
                else:
                    text_lines.append(line)
        return text_lines
    except Exception as e:
        logger.error(f"Aspose text extraction failed: {str(e)}")
        raise e

def extract_table_structure_using_aspose(docx_path):
    """
    Re-implemented to use structural extraction and return HTML tables.
    Note: The return format is now a list of HTML strings, not the old JSON format.
    The caller (extract_tables_from_docx) needs to be updated to handle this.
    """
    try:
        structure = extract_structure_using_aspose(docx_path)
        return structure.get("tables", [])
    except Exception as e:
        logger.error(f"Aspose table extraction failed: {str(e)}")
        raise e


def save_aspose_json(docx_path, output_json_path):
    """
    Extract structure using Aspose and save directly to a JSON file.
    """
    try:
        structure = extract_structure_using_aspose(docx_path)
        with open(output_json_path, 'w', encoding='utf-8') as f:
            json.dump(structure, f, ensure_ascii=False, indent=4)
        logger.info(f"Successfully saved Aspose structure to {output_json_path}")
        return True
    except Exception as e:
        logger.error(f"Failed to save Aspose JSON: {e}")
        return False


def extract_content_to_json(docx_path, output_json_path, llm_client):
    """
    使用 LLM 智能提取 Word 文档内容，生成扁平化的 JSON 知识库。
    支持 HTML 表格结构的语义理解。
    """
    try:
        try:
            # 1. 获取结构化数据 (Text + HTML Tables)
            structure = extract_structure_using_aspose(docx_path)
            html_tables = structure.get("tables", [])
            
            full_content = []
            for page in structure.get("pages", []):
                for line in page.get("textContent", []):
                    # Check for table placeholder
                    match = re.match(r"\[TABLE_(\d+)\]", line)
                    if match:
                        table_idx = int(match.group(1)) - 1
                        if 0 <= table_idx < len(html_tables):
                            # Append the full HTML table string
                            full_content.append(html_tables[table_idx])
                        else:
                            full_content.append(line)
                    else:
                        full_content.append(line)
                        
            # full_content contains strings (lines) and HTML strings (tables)
            
        except Exception as e:
            logger.error(f"Aspose extraction failed, falling back or aborting: {e}")
            raise e
        
        def _split_chunks(items, max_chars=6000): # Increased chunk size for HTML overhead
            chunks = []
            buf = []
            length = 0
            for item in items:
                s = item if isinstance(item, str) else str(item)
                # If a single item (e.g. table) is too huge, we might have issues.
                # But assume LLM can handle it or we just split it anyway (which breaks HTML).
                # Ideally, we should not split inside HTML table.
                
                if length + len(s) + 1 > max_chars and buf:
                    chunks.append("\n".join(buf))
                    buf = []
                    length = 0
                
                buf.append(s)
                length += len(s) + 1
                
            if buf:
                chunks.append("\n".join(buf))
            return chunks

        def _parse_json(text):
            try:
                # Find the outer-most JSON object
                start = text.find('{')
                end = text.rfind('}') + 1
                if start != -1 and end != -1:
                    return json.loads(text[start:end])
                raise ValueError("未找到JSON内容")
            except Exception:
                raise

        def _merge(a, b):
            for k, v in b.items():
                if k not in a:
                    a[k] = v
                else:
                    av = a[k]
                    if isinstance(av, dict) and isinstance(v, dict):
                        _merge(av, v)
                    elif isinstance(av, list):
                        if isinstance(v, list):
                            a[k] = av + [x for x in v]
                        else:
                            a[k] = av + [v]
                    else:
                        if av == v:
                            a[k] = av
                        else:
                            a[k] = [av, v] if not isinstance(av, list) else av + [v]
            return a

        chunks = _split_chunks(full_content)
        total = len(chunks)
        merged = {}
        raw_fallback = {}

        for idx, chunk in enumerate(chunks, start=1):
            prompt = f"""
            请分析以下文档内容片段（第{idx}/{total}段），将其提取为结构化 JSON 数据。
            
            文档内容片段：
            {chunk}
            
            要求：
            1. **核心目标**：提取文档中的关键事实、长文本描述以及**列表型数据**。
            2. **列表/对象提取（关键）**：
               - 如果遇到“主要完成人”、“主要完成单位”、“获奖情况”等包含多个实体的信息，**请务必提取为对象列表（List of Objects）**，而不是扁平化的字符串。
               - 例如：`"主要完成人": [{{"姓名": "张三", "职称": "教授", "单位": "X大"}}, {{"姓名": "李四", "职称": "讲师"}}]`。
               - **严禁**简单地将人名合并为 `"完成人": "张三、李四"`，必须保留每个人的详细属性（如性别、出生年月、职称、工作单位等）。
            3. **表格处理**：文档中包含 HTML 格式表格，请根据表格的行列关系提取每一行的数据。如果是人员名单表，请按行提取为人员对象列表。
            4. **长文本**：对于“成果简介”、“创新点”等长文本，保留原文段落。
            5. **键名**：使用中文键名，保持语义清晰（如“主要完成人”、“主要完成单位”、“成果简介”）。
            6. **合并策略**：如果一段内容是上一段的延续，请尽量保持结构的一致性以便后续合并。
            7. **忠实原文**：提取的值必须忠实于原文，不要进行推测或无中生有的补充。如果原文没有提及某项信息，不要生成对应的键值对。
            
            仅返回 JSON 对象。
            """
            messages = [
                {"role": "system", "content": "你是一个专业的数据提取专家，擅长将非结构化文档（含HTML表格）转化为结构化数据。"},
                {"role": "user", "content": prompt}
            ]
            try:
                response = llm_client.chat_completion(messages, temperature=0.1)
                data = _parse_json(response)
                merged = _merge(merged, data)
            except Exception as e:
                raw_fallback[f"Raw_Content_Chunk_{idx}"] = chunk[:500] + "..." # Truncate for log
                raw_fallback[f"Error_Chunk_{idx}"] = str(e)

        final_data = merged if merged else {"Error": "JSON解析失败"}
        if raw_fallback:
            final_data.update({"_extraction_errors": raw_fallback})

        with open(output_json_path, 'w', encoding='utf-8') as f:
            json.dump(final_data, f, ensure_ascii=False, indent=4)
        logger.info(f"成功使用 LLM 提取数据到 {output_json_path}")
        return True

    except Exception as e:
        logger.error(f"智能提取失败: {str(e)}")
        return False

def clean_cell_text(text):
    if not text:
        return ""
    text = text.strip()
    pattern = r'\s+([\u4e00-\u9fa5]{2,10}[：:])'
    text = re.sub(pattern, r'\n\1', text)
    return text

def extract_tables_from_docx(docx_path, output_excel_path):
    """
    从Word文档中提取表格和文本，保存为Excel文件 (Aspose Version)
    Re-implemented to use structural HTML tables.
    """
    try:
        # 1. 提取表格 (Using Aspose structural HTML output)
        structure = extract_structure_using_aspose(docx_path)
        html_tables = structure.get("tables", [])
        tables_data = []
        
        for html in html_tables:
            if not html.strip():
                continue
            try:
                # Wrap in StringIO to silence warning
                dfs = pd.read_html(StringIO(html))
                if dfs:
                    df = dfs[0]
                    # Clean cell text
                    # Apply cleaning to all cells. Convert to string first to avoid errors.
                    # Use map on each series to avoid deprecated applymap
                    df = df.astype(str).apply(lambda col: col.map(clean_cell_text))
                    tables_data.append(df)
            except Exception as e:
                logger.warning(f"Failed to parse HTML table: {e}")

        # 2. 提取文本 (Using our re-implemented text extraction)
        # Re-use extract_text_using_aspose which now returns lines with placeholders
        raw_lines = extract_text_using_aspose(docx_path)
        text_content = []
        
        for line in raw_lines:
            line = line.strip()
            if not line: continue
            
            # Check for table placeholder
            if re.match(r"\[TABLE_\d+\]", line):
                # This is a table, we skip it in "Text_Content" sheet as it's in Table sheets
                continue
                
            if line.startswith("[页眉]"):
                text_content.append({"Content": line.replace("[页眉]", "").strip(), "Type": "Header"})
                continue
            if line.startswith("[页脚]"):
                text_content.append({"Content": line.replace("[页脚]", "").strip(), "Type": "Footer"})
                continue
            
            text_content.append({"Content": line, "Type": "Text"})

        # 3. 写入Excel
        with pd.ExcelWriter(output_excel_path, engine='openpyxl') as writer:
            # 写入文本内容
            if text_content:
                df_text = pd.DataFrame(text_content)
                df_text.to_excel(writer, sheet_name="Text_Content", index=False)
            
            # 写入表格
            for i, df in enumerate(tables_data):
                sheet_name = f"Table_{i+1}"
                # Truncate sheet name if too long (Excel limit 31 chars)
                if len(sheet_name) > 31:
                    sheet_name = sheet_name[:31]
                df.to_excel(writer, sheet_name=sheet_name, index=False, header=False)
                
        logger.info(f"成功从 {docx_path} 提取数据到 {output_excel_path}")
        return True
    except Exception as e:
        logger.error(f"提取数据失败: {str(e)}")
        return False

def agent_extract_tables(docx_path, llm_client):
    """
    Agentic extraction: LLM decides to call the extraction tool to process the document.
    Demonstrates the "LLM uses Skill" pattern.
    """
    tools = [
        {
            "type": "function",
            "function": {
                "name": "extract_docx_structure",
                "description": "Extract structural content (tables, paragraphs, outline) from a DOCX file using Aspose backend. Returns JSON representation of the document.",
                "parameters": {
                    "type": "object",
                    "properties": {
                        "file_path": {
                            "type": "string",
                            "description": "The absolute path to the DOCX file."
                        }
                    },
                    "required": ["file_path"]
                }
            }
        }
    ]

    messages = [
        {"role": "system", "content": "You are a helpful assistant capable of extracting information from documents using provided tools. When asked to extract tables, use the available tool."},
        {"role": "user", "content": f"Please extract the table structure from this file: {docx_path}"}
    ]

    logger.info(f"Agent starts: {messages[-1]['content']}")
    
    try:
        # First turn: LLM decides to call tool
        response_msg = llm_client.chat_with_tools(messages, tools=tools)
        
        # Handle dict response (APIClient/OllamaClient returns dict now)
        if isinstance(response_msg, dict):
            messages.append(response_msg)
            tool_calls = response_msg.get("tool_calls")
        else:
            # Fallback for simple string response (should not happen with updated client)
            return response_msg

        if tool_calls:
            for tool_call in tool_calls:
                function_name = tool_call["function"]["name"]
                # Arguments are typically a JSON string
                arguments_str = tool_call["function"]["arguments"]
                try:
                    arguments = json.loads(arguments_str)
                except:
                    arguments = {}
                
                if function_name == "extract_docx_structure":
                    file_path = arguments.get("file_path")
                    logger.info(f"LLM invoking tool: {function_name} with path {file_path}")
                    
                    try:
                        # Call the actual function
                        structure = extract_structure_using_aspose(file_path)
                        
                        # Generate a summary for the LLM to avoid context overflow
                        table_count = len(structure.get("tables", []))
                        page_count = len(structure.get("pages", []))
                        outline_count = len(structure.get("outline", []))
                        
                        tool_output = json.dumps({
                            "status": "success",
                            "summary": f"Extracted {page_count} pages, {table_count} tables, {outline_count} outline items.",
                            "info": "Full structure is available in memory."
                        })
                        
                    except Exception as e:
                        tool_output = f"Error: {str(e)}"
                    
                    messages.append({
                        "role": "tool",
                        "tool_call_id": tool_call["id"],
                        "name": function_name,
                        "content": tool_output
                    })

            # Second turn: LLM summarizes result
            final_response = llm_client.chat_with_tools(messages, tools=tools)
            return final_response["content"]
        
        return response_msg.get("content", "")
        
    except Exception as e:
        logger.error(f"Agent execution failed: {str(e)}")
        return f"Agent execution failed: {str(e)}"
