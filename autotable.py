import os
import pandas as pd
import logging
import json
import subprocess
import shutil
from datetime import datetime
import re
from itertools import groupby

logger = logging.getLogger(__name__)

# Aspose EXE Path
ASPOSE_EXE = os.path.abspath(os.path.join(
    os.path.dirname(__file__), 
    "..", 
    "Aspose.Words for .NET v25.7.0", 
    "AsposeLicenseDemo", 
    "bin", "Debug", "net9.0", 
    "AsposeLicenseDemo.exe"
))

class AutoTable:
    """自动化填表处理核心类 (Aspose Backend)"""
    def __init__(self, knowledge_base_path, word_template_path, llm_client, output_folder="output"):
        self.knowledge_base_path = knowledge_base_path
        self.word_template_path = word_template_path
        self.output_folder = output_folder
        self.llm_client = llm_client
        self.knowledge_base = None
        self.knowledge_dict = None
        
        # New state variables for Aspose workflow
        self.table_structures = []
        self.pending_instructions = []
        self.doc_path = word_template_path # Store path instead of object

        if not os.path.exists(output_folder):
            os.makedirs(output_folder)

    def load_knowledge_base(self):
        try:
            logger.info(f"正在加载知识库: {self.knowledge_base_path}")
            if self.knowledge_base_path.endswith('.xlsx'):
                dfs = pd.read_excel(self.knowledge_base_path, sheet_name=None, header=None)
                structured_data = {}
                total_records = 0
                for sheet_name, df in dfs.items():
                    df = df.where(pd.notnull(df), "")
                    matrix_data = df.values.tolist()
                    structured_data[sheet_name] = matrix_data
                    total_records += len(matrix_data)
                self.knowledge_dict = structured_data
                logger.info(f"Excel知识库加载完成，共读取 {len(dfs)} 个工作表，合计 {total_records} 条数据")
                return True
            elif self.knowledge_base_path.endswith('.json'):
                with open(self.knowledge_base_path, 'r', encoding='utf-8') as f:
                    self.knowledge_dict = json.load(f)
                logger.info(f"JSON知识库加载完成: {self.knowledge_base_path}")
                return True
            else:
                logger.error("目前只支持 .xlsx 或 .json 格式知识库")
                return False
        except Exception as e:
            logger.error(f"知识库加载异常: {str(e)}")
            return False

    def load_template(self):
        try:
            logger.info(f"正在分析Word模板结构 (Aspose): {self.word_template_path}")
            if not os.path.exists(ASPOSE_EXE):
                logger.error(f"Aspose Executable not found at: {ASPOSE_EXE}")
                return False

            print(f"[Skill] Invoking Aspose Backend for Table Extraction: {ASPOSE_EXE}") # Console output
            cmd = [ASPOSE_EXE, self.word_template_path, "extract-tables"]
            result = subprocess.run(cmd, capture_output=True, text=True, encoding='utf-8')
            
            if result.returncode != 0:
                logger.error(f"Aspose extraction failed: {result.stderr}")
                return False
                
            self.table_structures = json.loads(result.stdout)
            logger.info(f"模板分析完成，包含{len(self.table_structures)}个表格")
            return True
        except Exception as e:
            logger.error(f"模板加载失败: {str(e)}")
            return False

    def _convert_aspose_row_to_python_style(self, aspose_rows):
        """Convert PascalCase Aspose JSON to snake_case python-docx style structure"""
        python_rows = []
        for r_idx, row in enumerate(aspose_rows):
            py_row = []
            for c_idx, cell in enumerate(row):
                # Aspose keys: Text, IsMerged, IsEmpty
                cell_info = {
                    "id": f"{r_idx}_{c_idx}",
                    "text": cell.get("Text", ""),
                    "is_merged": cell.get("IsMerged", False),
                    "is_empty": cell.get("IsEmpty", False)
                }
                if cell_info["is_merged"]:
                    del cell_info["id"]
                py_row.append(cell_info)
            python_rows.append(py_row)
        return python_rows

    def analyze_tables_with_llm(self, table_structure, knowledge_context, used_contexts=None):
        # Same logic as before
        data_format_desc = "扁平化的 JSON 键值对（Key-Value）" if isinstance(knowledge_context, dict) else "按 Sheet（来源）分组的二维数组（矩阵）格式"
        
        used_context_desc = ""
        if used_contexts and len(used_contexts) > 0:
            used_context_desc = f"""
        **上下文去重约束（重要）**：
        当前文档中包含多个结构相似的表格，用于填写不同实体（人员/项目）的信息。
        以下实体标识（如姓名、项目名）**已被前面的表格使用过**：
        {json.dumps(used_contexts, ensure_ascii=False)}
        
        **请务必从知识库中选择一个【未使用过】的新实体数据进行填充。**
        """

        prompt = f"""
        请分析以下表格结构（JSON格式），并结合提供的知识库数据，直接在 JSON 数据上修改，将知识库中的值填入对应的单元格。
        
        **重要提示**：输出必须是严格合法的 JSON 格式。注意转义字符，不要有尾随逗号。
        
        知识库数据（{data_format_desc}）：
        {json.dumps(knowledge_context, ensure_ascii=False, default=str)}
        
        表格结构（JSON List of Tables）：
        {json.dumps(table_structure, ensure_ascii=False)}
        
        {used_context_desc}
        
        **任务说明**：
        0. **核心原则（防幻觉）**：你只能使用提供的“知识库数据”来填充表格。如果知识库中没有找到对应的信息，**请保持单元格原样**（或留空），**绝对不要编造**数据。不要尝试计算或推测日期、数字等，除非知识库里有明确依据。
        1. 你可以直接修改上述“表格结构”JSON中的 `text` 字段。
        2. **优先填写空白格**：请优先寻找与 Label 相邻的 **空白单元格** (`"is_empty": true`) 进行填写，而不是直接修改 Label 所在的单元格。
           - 例如：遇到 `[[{{"text": "姓名："}}, {{"text": "", "is_empty": true}}]]`，请将第二个单元格修改为 `{{ "text": "张三" }}`。
           - 只有当没有相邻空白格时（即 Label 和下划线在同一个单元格内），才修改 Label 单元格。
        3. **不要修改** `id` 和 `is_merged` 字段。
        4. **Label/表头处理（严禁重复）**：
           - **短Label（如“姓名：”）**：请直接在Label后填入内容。例如：“姓名：张三”。
           - **复杂排版（如“起始：____ 年 __ 月”）**：请只替换下划线或空白部分，**绝对不要**重复“起始：”这个词。例如：“起始：1998 年 1 月”。
           - **多字段合并（如“起始... 完成...”）**：如果单元格包含多个需要填写的时间点（如起始时间、完成时间），请确保填入所有值。**严禁**在填好的内容后重复保留未填写的模板（如“完成：   年   月”）。正确示例：“起始：2020年9月  完成：2023年6月”。
           - **长文本题（如“1. 成果简介...”、“3. 成果的创新点...”）**：这不仅是标题，也是填空区域。请务必将知识库中对应的长文本内容**追加**到该单元格的标题下方（换行）。**不要**因为它是标题就跳过！
        5. **避免重复**：请仔细检查，如果单元格内已经包含了标题（如“本人签名：”或“起始：”），你填入的内容**不要**再次包含该标题字样。
        6. **列表实体匹配**：如果知识库中包含“主要完成人”等列表数据，请根据表格的顺序或去重逻辑依次选择实体填入。如果是“第( )完成人”，请同时填入对应的序号（如“1”）。
        7. **自动增行（重要）**：如果知识库中的列表数据（如获奖情况、成员名单）数量超过了表格现有的行数，请直接在 JSON 结构中**复制并追加**新的 Row 对象。系统会自动识别这些新增的行并在文档中插入。请确保新增行的格式与该列表区域的其他行保持一致。
        
        **返回格式**：
        请返回**修改后的完整 JSON 表格结构**。
        格式必须与输入的“表格结构”完全一致（包含 TableIndex 等），只是部分 `text` 字段被更新了。
        
        为了方便去重，请在 JSON 列表的**最前面**添加一个特殊对象：
        [
            {{ "__identity__": "本次使用的实体标识（如张三）" }},
            {{ "TableIndex": 1, "Rows": [...] }},
            {{ "TableIndex": 2, "Rows": [...] }},
            ...
        ]
        """
        try:
            messages = [
                {"role": "system", "content": "你是一个专业的表格填充助手。"},
                {"role": "user", "content": prompt}
            ]
            result = self.llm_client.chat_completion(messages, temperature=0.1)
            return self._parse_llm_json(result)
        except Exception as e:
            logger.error(f"表格分析失败: {str(e)}")
            return []


    def _parse_llm_json(self, text):
        text = text.strip()
        
        # 1. 尝试提取 Markdown 代码块 (最可靠)
        match = re.search(r'```(?:json)?\s*(.*?)```', text, re.DOTALL)
        if match:
            text = match.group(1).strip()
        
        # 2. 尝试直接解析
        try:
            return json.loads(text)
        except json.JSONDecodeError:
            pass
            
        # 3. 尝试寻找最外层的 [] 或 {}
        # 找到第一个 [ 或 {
        start_idx = -1
        end_idx = -1
        
        for i, char in enumerate(text):
            if char in ['[', '{']:
                start_idx = i
                break
        
        if start_idx != -1:
            # 找到最后一个 ] 或 }
            for i in range(len(text) - 1, start_idx, -1):
                if text[i] in [']', '}']:
                    end_idx = i + 1
                    break
            
            if end_idx != -1:
                candidate = text[start_idx:end_idx]
                try:
                    return json.loads(candidate)
                except json.JSONDecodeError:
                    pass
        
        # 4. 如果还是失败，尝试修复常见错误（如中文引号）
        text_fixed = text.replace("“", '"').replace("”", '"')
        try:
            return json.loads(text_fixed)
        except:
            pass
            
        # 5. 记录失败的片段以便调试
        error_snippet = text[:200] + "..." if len(text) > 200 else text
        raise ValueError(f"无法解析 JSON (len={len(text)}). Snippet: {error_snippet}")

    def fill_document(self):
        if not self.table_structures or self.knowledge_dict is None:
            logger.error("文档结构或知识库未正确初始化")
            return False
        
        filled_count = 0
        used_identities = []
        self.pending_instructions = [] # Clear previous instructions
        
        # Sort tables by PageIndex to ensure groupby works correctly
        sorted_tables = sorted(self.table_structures, key=lambda x: x.get("PageIndex", 0))
        
        for page_idx, page_tables in groupby(sorted_tables, key=lambda x: x.get("PageIndex", 0)):
            page_tables_list = list(page_tables)
            logger.info(f"正在处理第 {page_idx} 页，共 {len(page_tables_list)} 个表格")
            
            page_tables_payload = []
            for tbl in page_tables_list:
                rows_python = self._convert_aspose_row_to_python_style(tbl.get("Rows", []))
                page_tables_payload.append({
                    "TableIndex": tbl.get("TableIndex"),
                    "Rows": rows_python
                })
            
            if not page_tables_payload:
                continue
                
            # LLM Analysis (Page Chunk)
            modified_page_tables = self.analyze_tables_with_llm(page_tables_payload, self.knowledge_dict, used_contexts=used_identities)
            
            if not isinstance(modified_page_tables, list):
                logger.warning(f"第 {page_idx} 页 LLM 返回格式错误")
                continue
                
            # Identity extraction
            if modified_page_tables and isinstance(modified_page_tables[0], dict) and "__identity__" in modified_page_tables[0]:
                identity = modified_page_tables[0]["__identity__"]
                used_identities.append(identity)
                logger.info(f"页面 {page_idx} 使用了实体: {identity}")
                modified_page_tables = modified_page_tables[1:]
            
            # Process results for each table in the page
            for mod_table in modified_page_tables:
                 if not isinstance(mod_table, dict) or "TableIndex" not in mod_table:
                     continue
                 
                 t_idx = mod_table["TableIndex"]
                 mod_rows = mod_table.get("Rows", [])
                 
                 # Find original table data to compare
                 original_table_data = next((t for t in page_tables_list if t["TableIndex"] == t_idx), None)
                 if not original_table_data:
                     continue
                     
                 original_rows = self._convert_aspose_row_to_python_style(original_table_data.get("Rows", []))
                 
                 filled_count = self._process_table_instructions(t_idx, original_rows, mod_rows, filled_count)

        logger.info(f"完成文档分析，生成 {len(self.pending_instructions)} 条填充指令")
        return True

    def _process_table_instructions(self, table_idx, original_structure, modified_rows, filled_count):
        """
        Smart comparison of original structure and LLM modified rows to handle:
        1. Content Updates (using strict Label Protection & Fallback)
        2. Row Insertions (when modified_rows has more items)
        """
        orig_idx = 0
        mod_idx = 0
        
        while mod_idx < len(modified_rows):
            mod_row = modified_rows[mod_idx]
            
            # Case 1: All original rows consumed. Treat remaining modified rows as Insertions (Append).
            if orig_idx >= len(original_structure):
                self._generate_insert_instructions(table_idx, mod_idx, mod_row)
                filled_count += 1
                mod_idx += 1
                continue

            orig_row = original_structure[orig_idx]
            
            # Helper to check if rows are "compatible" (structurally or anchor-wise)
            def is_row_match(r_orig, r_mod):
                if not isinstance(r_mod, list): return False
                # If length mismatch is large, likely different row type
                if abs(len(r_orig) - len(r_mod)) > 2: return False 
                
                match_score = 0
                total_anchors = 0
                
                for c_o, c_m in zip(r_orig, r_mod):
                    txt_o = c_o.get("text", "").strip()
                    txt_m = c_m.get("text", "").strip() if isinstance(c_m, dict) else ""
                    
                    # If original has text (Anchor), modified should have it too
                    if txt_o and len(txt_o) > 2 and not c_o.get("is_empty"):
                        # Heuristic: Is this a label?
                        if re.match(r'^[^：:]+[：:]$', txt_o) or txt_o in txt_m:
                            match_score += 1
                        total_anchors += 1
                
                if total_anchors > 0:
                    return match_score >= (total_anchors * 0.5) # At least 50% anchors match
                return True # If no anchors (empty row), assume match

            # Case 2: Match Found
            if is_row_match(orig_row, mod_row):
                filled_count = self._generate_update_instructions(table_idx, mod_idx, orig_row, mod_row, filled_count)
                orig_idx += 1
                mod_idx += 1
                continue
            
            # Case 3: Mismatch. Is it an Insertion?
            # Lookahead: Does `orig_row` appear later in modified_rows?
            found_orig_at = -1
            for k in range(1, 10):
                if mod_idx + k < len(modified_rows):
                    if is_row_match(orig_row, modified_rows[mod_idx + k]):
                        found_orig_at = mod_idx + k
                        break
            
            if found_orig_at != -1:
                # Found original row later. So [mod_idx ... found_orig_at - 1] are Insertions.
                self._generate_insert_instructions(table_idx, mod_idx, mod_row)
                filled_count += 1
                mod_idx += 1
                # Do NOT increment orig_idx
            else:
                # Original row NOT found later. Assume Update (or heavy modification).
                filled_count = self._generate_update_instructions(table_idx, mod_idx, orig_row, mod_row, filled_count)
                orig_idx += 1
                mod_idx += 1
        
        return filled_count

    def _generate_insert_instructions(self, table_idx, target_row_idx, mod_row_data):
        if not isinstance(mod_row_data, list): return
        # 1. Emit "insert_row"
        self.pending_instructions.append({
            "TableIndex": table_idx,
            "RowIndex": target_row_idx,
            "Action": "insert_row",
            "CellIndex": 0,
            "Text": ""
        })
        # 2. Emit updates for this new row's cells
        for c_idx, mod_cell in enumerate(mod_row_data):
            if isinstance(mod_cell, dict):
                text = mod_cell.get("text", "")
                if text:
                    self.pending_instructions.append({
                        "TableIndex": table_idx,
                        "RowIndex": target_row_idx,
                        "CellIndex": c_idx,
                        "Text": text
                    })

    def _generate_update_instructions(self, table_idx, r_idx, row_data, mod_row, filled_count):
        # Note: r_idx is the target row index in the dynamic table
        if not isinstance(mod_row, list): return filled_count
        
        for c_idx, cell_info in enumerate(row_data):
            if c_idx >= len(mod_row): break
            mod_cell = mod_row[c_idx]
            
            if cell_info.get("is_merged"): continue
                
            original_text = cell_info.get("text", "").strip()
            new_text = mod_cell.get("text", original_text) if isinstance(mod_cell, dict) else original_text
            
            # 过滤 LLM 可能返回的 "(Merged)" 占位符
            if str(new_text).strip() == "(Merged)":
                new_text = original_text
                
            operation_desc = "LLM Generate"
            
            # --- Fallback Strategy for Long Text Fields ---
            # 如果 LLM 未修改内容，且是已知的长文本字段，尝试从知识库直接匹配
            if new_text == original_text and isinstance(self.knowledge_dict, dict):
                # 改为允许任何字典结构（不再强制扁平化），只要能找到对应 Key 即可
                matched_kb_value = None
                
                # 1. 定义 Header 关键词与 Knowledge Key 的映射关系 (Fuzzy Match)
                # Key: Header 关键词, Value: Knowledge Key (优先尝试)
                header_mapping = [
                    (["成果简介", "解决的教学问题"], ["成果简介", "主要解决的教学问题", "成果简介及主要解决的教学问题"]),
                    (["方法", "解决教学问题的方法"], ["成果解决教学问题的方法", "解决教学问题的方法", "方法"]),
                    (["创新点", "成果的创新点"], ["成果的创新点", "创新点"]),
                    (["推广应用", "应用效果"], ["成果的推广应用效果", "推广应用效果", "推广应用"]),
                ]
                
                found_key = None
                # 尝试匹配 Header
                for header_keywords, kb_candidate_keys in header_mapping:
                    if any(hk in original_text for hk in header_keywords):
                        # 尝试在 Knowledge Base 中寻找对应的 Key
                        for candidate in kb_candidate_keys:
                            if candidate in self.knowledge_dict:
                                found_key = candidate
                                break
                    if found_key:
                        break
                
                if found_key:
                    raw_value = self.knowledge_dict[found_key]
                    
                    # 2. 格式化提取的值 (Handle List/Dict/String)
                    extracted_text = ""
                    
                    if isinstance(raw_value, str):
                        extracted_text = raw_value
                    elif isinstance(raw_value, list):
                        # 列表可能是 ["str1", "str2"] 或 [{"序号":1, "内容":...}, ...]
                        lines = []
                        for item in raw_value:
                            if isinstance(item, str):
                                lines.append(item)
                            elif isinstance(item, dict):
                                # 尝试提取 value 中包含“描述”、“内容”等的字段
                                desc = ""
                                # 优先找长文本字段
                                for k, v in item.items():
                                    if any(x in k for x in ["描述", "内容", "创新点", "方法", "问题"]) and isinstance(v, str):
                                        desc = v
                                        break
                                # 没找到就找第一个 value 是 string 的
                                if not desc:
                                    for v in item.values():
                                        if isinstance(v, str) and len(v) > 5:
                                            desc = v
                                            break
                                
                                # 尝试提取序号
                                seq = str(item.get("序号", ""))
                                if seq and desc:
                                    lines.append(f"{seq}. {desc}")
                                elif desc:
                                    lines.append(desc)
                        extracted_text = "\n".join(lines)
                        
                    elif isinstance(raw_value, dict):
                        # 字典可能是 {"总体描述": "...", "具体内容": [...]}
                        parts = []
                        if "总体描述" in raw_value:
                            parts.append(raw_value["总体描述"])
                        if "具体内容" in raw_value and isinstance(raw_value["具体内容"], list):
                            for item in raw_value["具体内容"]:
                                if isinstance(item, dict):
                                    seq = str(item.get("序号", ""))
                                    content = item.get("内容", "") or item.get("详细描述", "")
                                    if seq and content:
                                        parts.append(f"{seq}. {content}")
                                    elif content:
                                        parts.append(content)
                        extracted_text = "\n".join(parts)

                    if extracted_text:
                        matched_kb_value = extracted_text

                if matched_kb_value:
                    logger.info(f"触发规则兜底: 自动填充 '{original_text[:10]}...' using Key='{found_key}'")
                    new_text = original_text.strip() + "\n" + matched_kb_value.strip()
                    operation_desc = "Fallback Rule"

            if new_text != original_text and new_text:
                is_pure_label = re.match(r'^[^:：]{2,8}[:：]$', original_text)
                if is_pure_label and len(str(new_text)) > len(original_text):
                    if c_idx + 1 < len(row_data):
                        right_cell_info = row_data[c_idx + 1]
                        if right_cell_info.get("is_empty"):
                            redirect_value = str(new_text).replace(original_text, "").strip()
                            if redirect_value:
                                logger.info(f"[Fill Op] Table {table_idx+1} ({r_idx},{c_idx+1}) | Header: '{original_text[:15]}...' | Op: Redirect to Right | Content: '{redirect_value[:20]}...'")
                                self.pending_instructions.append({
                                    "TableIndex": table_idx,
                                    "RowIndex": r_idx,
                                    "CellIndex": c_idx + 1,
                                    "Text": redirect_value
                                })
                                filled_count += 1
                                continue 
                orig_norm = re.sub(r'\s+', '', original_text)
                new_norm = re.sub(r'\s+', '', str(new_text))
                is_label_like = bool(re.match(r'^[\u4e00-\u9fa5\s]{1,8}(?:[:：])?$', original_text)) or ("姓名" in original_text and len(original_text) <= 10)
                if (str(new_text).startswith(original_text) or new_norm.startswith(orig_norm)) and is_label_like:
                    if c_idx + 1 < len(row_data):
                        right_cell_info = row_data[c_idx + 1]
                        appended = str(new_text)[len(original_text):].strip()
                        if not appended and new_norm.startswith(orig_norm):
                            appended = new_norm[len(orig_norm):].strip()
                        if right_cell_info.get("is_empty") and appended:
                            logger.info(f"[Fill Op] Table {table_idx+1} ({r_idx},{c_idx+1}) | Header: '{original_text[:15]}...' | Op: Append to Right | Content: '{appended[:20]}...'")
                            self.pending_instructions.append({
                                "TableIndex": table_idx,
                                "RowIndex": r_idx,
                                "CellIndex": c_idx + 1,
                                "Text": appended
                            })
                            filled_count += 1
                            continue 

                # --- Smart Logic (Simplified) ---
                original_new_text = str(new_text)
                
                # 1. Clean Label Prefix
                # 智能判断是否为复杂标题（如 "1. 成果简介"、"3.创新点(800字)"）
                # 如果是复杂标题，则不执行前缀清除，保留 LLM 返回的完整 "标题+内容"
                is_numbered_header = bool(re.match(r'^[\(\（]?\d+[.\、\）\)]', original_text))
                is_instruction_header = any(k in original_text for k in ["不超过", "字)", "字）", "简介", "创新点", "应用效果", "存在问题"])
                is_long_header = len(original_text) > 15
                
                should_skip_clean = is_numbered_header or is_instruction_header or is_long_header
                
                if str(new_text).startswith(original_text) and len(str(new_text)) > len(original_text):
                    if not should_skip_clean:
                        new_text = str(new_text)[len(original_text):].strip()
                elif not should_skip_clean:
                    label_match = re.match(r'^([^:：]{1,5}[:：])', original_text)
                    if label_match:
                        prefix = label_match.group(1)
                        if str(new_text).startswith(prefix) and len(str(new_text)) > len(prefix):
                            new_text = str(new_text)[len(prefix):].strip()

                # 2. Label Protection
                label_match = re.match(r'^([^:：]{1,10}[:：])', original_text)
                if label_match:
                    label_prefix = label_match.group(1)
                    # Fuzzy check: ignore spaces and punctuation to prevent duplicates
                    # 模糊匹配：忽略空格和标点，防止出现 "起始年月：起始年月 2023" 这种重复
                    clean_punc_pattern = r'[\s:：_]+'
                    clean_label = re.sub(clean_punc_pattern, '', label_prefix)
                    clean_new = re.sub(clean_punc_pattern, '', str(new_text))
                    
                    if not clean_new.startswith(clean_label):
                        new_text = label_prefix + str(new_text)
                
                # 2.5 Fill Blank Cleaning
                # 修复：对于“年 月”类型的填空，LLM 可能会在填完后重复附带未填的模板，需要清除
                # Improved cleaning logic to handle multi-line garbage and different formats
                is_fill_blank_check = re.search(r'[:：]\s*[_\u005F\u2013-\u2017\s]{2,}', original_text) or ("年" in original_text and "月" in original_text)
                
                if is_fill_blank_check:
                        lines = str(new_text).splitlines()
                        cleaned_lines = []
                        
                        # Regex 1: "Label: ... Year ... Month" (standard date template)
                        # Matches start of line (ignoring leading spaces), optional label, optional colon, Year, Month
                        p1 = re.compile(r'^\s*(?:[\u4e00-\u9fa5\w\(\)（）\s]{2,50})?[:：]?\s*年\s*月\s*$')
                        
                        # Regex 2: "Label: ... Year" (year only template)
                        # Must have underscores or multiple spaces before Year to distinguish from filled "2014 Year"
                        p2 = re.compile(r'^\s*(?:[\u4e00-\u9fa5\w\(\)（）\s]{2,50})?[:：]?\s*[_\u005F\u2013-\u2017\s]{2,}\s*年\s*$')

                        for line in lines:
                            # Skip lines that look like unfilled templates
                            if p1.match(line):
                                logger.info(f"[Cleaner] Removed garbage line (Type 1): '{line.strip()}'")
                                continue
                            if p2.match(line):
                                logger.info(f"[Cleaner] Removed garbage line (Type 2): '{line.strip()}'")
                                continue
                            cleaned_lines.append(line)
                        
                        cleaned_text = "\n".join(cleaned_lines).strip()
                        
                        if len(cleaned_text) < len(str(new_text).strip()):
                            new_text = cleaned_text

                # 3. Smart Append
                is_cleaned = len(str(new_text)) < len(original_new_text) if 'original_new_text' in locals() else False
                has_content = len(str(new_text).strip()) > 0
                
                if has_content and not is_cleaned:
                    is_fill_blank = is_fill_blank_check
                    is_long_header = len(original_text) > 20 and not is_fill_blank
                    is_keyword_header = any(k in original_text for k in ["简介", "方法", "概述", "说明", "问题"])
                    
                    if (is_long_header or is_keyword_header):
                        original_fingerprint_clean = re.sub(r'\s+', '', original_text[:15])
                        new_text_clean = re.sub(r'\s+', '', str(new_text))
                        
                        # 增强判断：检查是否保留了原来的标签头部（Savior Logic）
                        # 如果新文本以原文本的头部开头（忽略空白和下划线），则认为是填充而非覆盖，不需要追加原文本
                        clean_pattern = r'[\s_]+'
                        original_clean_full = re.sub(clean_pattern, '', original_text)
                        new_clean_full = re.sub(clean_pattern, '', str(new_text))
                        
                        # 取前 8 个有效字符作为前缀指纹 (例如 "起始：", "1.成果简介")
                        prefix_len = min(8, len(original_clean_full))
                        original_prefix = original_clean_full[:prefix_len]
                        starts_with_prefix = new_clean_full.startswith(original_prefix) if original_prefix else False
                        
                        # 添加调试日志
                        logger.info(f"[SmartAppend] Cell({r_idx},{c_idx}) Long/Key={is_long_header}/{is_keyword_header}")
                        logger.info(f"[SmartAppend] OrigPrefix: {original_prefix} | NewStart: {new_clean_full[:10] if new_clean_full else ''} | Match: {starts_with_prefix}")

                        if original_fingerprint_clean and original_fingerprint_clean not in new_text_clean and not starts_with_prefix:
                            logger.info(f"检测到长标题被覆盖，正在尝试恢复追加模式: {original_text[:10]}...")
                            new_text = original_text.strip() + "\n" + str(new_text).strip()
                            operation_desc = "Smart Append (Recovery)"

                    # Add Instruction
                    logger.info(f"[Fill Op] Table {table_idx+1} ({r_idx},{c_idx}) | Header: '{original_text[:15]}...' | Op: {operation_desc} | Content: '{str(new_text)[:20].replace(chr(10), ' ')}...'")
                    self.pending_instructions.append({
                        "TableIndex": table_idx,
                        "RowIndex": r_idx,
                        "CellIndex": c_idx,
                        "Text": str(new_text)
                    })
                    filled_count += 1
        return filled_count

    def save_document(self, filename=None):
        if not self.pending_instructions:
            logger.warning("没有可执行的填充指令")
            return False
            
        if not filename:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            base_name = os.path.splitext(os.path.basename(self.word_template_path))[0]
            filename = f"{base_name}_filled_{timestamp}.docx"
        
        output_path = os.path.abspath(os.path.join(self.output_folder, filename))
        instructions_path = os.path.abspath(os.path.join(self.output_folder, "fill_tasks.json"))
        
        try:
            # Save instructions
            with open(instructions_path, "w", encoding='utf-8') as f:
                json.dump(self.pending_instructions, f, ensure_ascii=False, indent=2)
            
            # Execute Aspose Fill
            logger.info(f"调用 Aspose 执行填充: {instructions_path} -> {output_path}")
            print(f"[Skill] Invoking Aspose Backend for Document Filling: {ASPOSE_EXE}") # Console output
            cmd = [ASPOSE_EXE, self.word_template_path, "fill", instructions_path, output_path]
            
            result = subprocess.run(cmd, capture_output=True, text=True, encoding='utf-8')
            if result.returncode != 0:
                logger.error(f"Aspose Fill failed: {result.stderr}")
                logger.error(f"Stdout: {result.stdout}")
                return False
                
            logger.info(f"文档已保存至: {output_path}")
            return True
        except Exception as e:
            logger.error(f"文档保存失败: {str(e)}")
            return False

    def run(self):
        logger.info("启动自动化填表流程 (Aspose版)")
        if all([
            self.load_knowledge_base(),
            self.load_template(),
            self.fill_document(),
            self.save_document()
        ]):
            logger.info("流程执行成功")
            return True
        logger.error("流程执行过程中发生错误")
        return False
