import os
import pandas as pd
import logging
from docx import Document
import json
from datetime import datetime

import re

logger = logging.getLogger(__name__)

class AutoTable:
    """自动化填表处理核心类"""
    def __init__(self, knowledge_base_path, word_template_path, llm_client, output_folder="output"):
        self.knowledge_base_path = knowledge_base_path
        self.word_template_path = word_template_path
        self.output_folder = output_folder
        self.llm_client = llm_client
        self.knowledge_base = None
        self.knowledge_dict = None
        self.doc = None

        if not os.path.exists(output_folder):
            os.makedirs(output_folder)

    def load_knowledge_base(self):
        try:
            logger.info(f"正在加载知识库: {self.knowledge_base_path}")
            if self.knowledge_base_path.endswith('.xlsx'):
                # 读取所有工作表，不将第一行作为表头（header=None）
                # 这样可以处理 Key-Value 型的表格，也可以避免错误列名的问题
                dfs = pd.read_excel(self.knowledge_base_path, sheet_name=None, header=None)
                
                # 构建结构化知识库：{ "Sheet名": [ [行1数据], [行2数据], ... ], ... }
                structured_data = {}
                total_records = 0
                
                for sheet_name, df in dfs.items():
                    # 处理 NaN 值为 None/空字符串
                    df = df.where(pd.notnull(df), "")
                    
                    # 转换为二维列表 (List of Lists)
                    # 这种格式最通用，既适合列表型表格，也适合 KV 型表格
                    # LLM 可以根据数据分布自行推断行列关系
                    matrix_data = df.values.tolist()
                    
                    # 特殊处理 Text_Content: 如果它是原来的格式（有 Header），pd.read_excel(header=None) 会把 Header 也读成第一行数据
                    # 但 Text_Content 比较简单，即使把 'Content', 'Type' 当作数据也不影响理解
                    
                    structured_data[sheet_name] = matrix_data
                    total_records += len(matrix_data)
                
                self.knowledge_dict = structured_data
                logger.info(f"Excel知识库加载完成，共读取 {len(dfs)} 个工作表，合计 {total_records} 条数据")
                
                # 输出结构化后的字典供调试/查看
                logger.info(f"结构化处理后的知识库字典: {json.dumps(self.knowledge_dict, ensure_ascii=False, default=str)}")
                
                return True
            elif self.knowledge_base_path.endswith('.json'):
                # 直接读取 JSON 格式的知识库
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
            logger.info(f"正在加载Word模板: {self.word_template_path}")
            self.doc = Document(self.word_template_path)
            logger.info(f"模板加载完成，包含{len(self.doc.tables)}个表格")
            return True
        except Exception as e:
            error_msg = str(e)
            if "no relationship of type" in error_msg or "File is not a zip file" in error_msg:
                logger.error(f"模板加载失败: 文件格式不正确。请确保您上传的是真正的 .docx 文件，而不是直接修改后缀名的 .doc 文件。请尝试用 Word 打开该文件并'另存为' .docx 格式。({error_msg})")
            else:
                logger.error(f"模板加载失败: {error_msg}")
            return False

    def _is_potential_slot(self, text):
        """判断文本是否为填空位"""
        if not text:
            return True
            
        # 移除常见的不可见字符和空白
        clean_text = text.strip().replace('\u200b', '').replace('\u3000', ' ')
        if not clean_text:
            return True
            
        if set(clean_text).issubset(set(" _()（）")):
            return True
        
        # 正则增强规则
        # 包含连续下划线
        if re.search(r'[_]{2,}', clean_text):
            return True
        # 包含空括号 或 提示性括号
        elif re.search(r'[(\uff08](?:\s*|.*?(?:填写|输入|粘贴|限|字|内容).*?)[)\uff09]', clean_text):
            return True
        # 包含 "年" 和 "月" 的日期格式
        elif '年' in clean_text and '月' in clean_text:
            if not re.match(r'^\d', clean_text):
                 if re.search(r'[_]+|\s+', clean_text):
                     return True
        # 启发式规则：以冒号结尾的 Prompt
        elif re.search(r'[:：]\s*$', clean_text):
            return True
            
        return False

    def _preprocess_paragraphs(self, paragraphs):
        """
        预处理段落列表：识别填空位
        返回: (markdown_text, anchor_map, id_to_text_map)
        anchor_map: {anchor_id: paragraph_index}
        """
        anchor_map = {}
        id_to_text_map = {}
        markdown_lines = []
        anchor_counter = 1
        
        for idx, para in enumerate(paragraphs):
            text = para.text.strip()
            if self._is_potential_slot(text):
                anchor_id = f"{{{{ID_{anchor_counter:03d}}}}}"
                anchor_map[anchor_id] = idx
                id_to_text_map[anchor_id] = f"原内容: '{text}'"
                
                # 在Markdown中展示
                # 尝试更智能的展示：如果段落很短，直接展示ID；如果长，展示部分上下文？
                # 暂时简单处理：直接用 ID 替换原文本展示给 LLM
                markdown_lines.append(f"- {anchor_id} (原: {text})")
                anchor_counter += 1
            else:
                # 如果不是填空位，但也包含在文档中，是否给LLM看？
                # 如果是纯文本，可能包含上下文信息。
                # 最好还是提供一些上下文。
                if len(text) > 0:
                    markdown_lines.append(f"- {text}")
        
        markdown_text = "\n".join(markdown_lines)
        return markdown_text, anchor_map, id_to_text_map

    def _preprocess_table(self, table):
        """
        预处理表格：识别填空位，生成带锚点的Markdown文本，并记录锚点映射。
        返回: (markdown_text, anchor_map, id_to_text_map)
        """
        anchor_map = {}
        id_to_text_map = {}
        markdown_lines = []
        anchor_counter = 1
        
        # 使用列表存储已处理的单元格 _tc 对象，以处理合并单元格
        processed_tcs = [] # Store (tc_object, anchor_id)
        
        # 遍历每一行
        for row_idx, row in enumerate(table.rows):
            row_cells_text = []
            for col_idx, cell in enumerate(row.cells):
                cell_text = cell.text.strip()
                
                # 检查该单元格是否已处理过
                current_tc = cell._tc
                existing_anchor_id = None
                for seen_tc, seen_id in processed_tcs:
                    if current_tc == seen_tc:
                        existing_anchor_id = seen_id
                        break
                
                if existing_anchor_id:
                    # 如果已处理过，直接使用之前的 ID
                    row_cells_text.append(existing_anchor_id)
                    # 注意：我们不需要再次添加到 anchor_map，因为之前已经加过了
                    # 但我们需要确保在 Markdown 中显示这个 ID，以便 LLM 理解表格结构
                    continue

                # 判断是否是填空位
                is_potential_slot = self._is_potential_slot(cell_text)
                
                if is_potential_slot:
                    anchor_id = f"{{{{ID_{anchor_counter:03d}}}}}"
                    anchor_map[anchor_id] = (row_idx, col_idx)
                    
                    # 记录为已处理
                    processed_tcs.append((current_tc, anchor_id))
                    
                    # 尝试获取上下文提示（Label）
                    context_hint = ""
                    try:
                        # 尝试获取左侧单元格文本作为提示
                        if col_idx > 0:
                            left_text = row.cells[col_idx-1].text.strip()
                            if left_text and len(left_text) < 20:
                                context_hint = f" | 左侧Label: {left_text}"
                        
                        # 如果左侧为空，尝试获取上方单元格文本（针对上下结构的表格）
                        if not context_hint and row_idx > 0:
                            top_text = table.cell(row_idx-1, col_idx).text.strip()
                            if top_text and len(top_text) < 20:
                                context_hint = f" | 上方Label: {top_text}"
                    except Exception:
                        pass

                    # 记录原始文本和上下文提示，供LLM参考
                    id_to_text_map[anchor_id] = f"原内容: '{cell_text}'{context_hint}"
                    row_cells_text.append(anchor_id)
                    anchor_counter += 1
                else:
                    # 调试日志：记录为什么这个单元格没有被选中（采样打印）
                    # if len(cell_text) > 0 and len(cell_text) < 10:
                    #    logger.debug(f"跳过单元格: '{cell_text}' - 未匹配规则")
                    
                    # 记录非填空位单元格为已处理（虽然没有ID，但我们不想重复处理）
                    # 不过，非填空位通常有内容，如果重复出现，在Markdown里重复显示内容是正确的
                    # 所以这里不需要记录 processed_tcs，除非我们想去重显示？
                    # 对于Markdown表格，如果合并单元格跨列，我们通常希望显示 | Content | Content | Content |
                    # 这样结构是对齐的。所以这里保持原样。
                    
                    # 清理一下换行符，以免破坏Markdown表格结构
                    clean_text = cell_text.replace('\n', '<br>')
                    row_cells_text.append(clean_text)
            
            markdown_lines.append("| " + " | ".join(row_cells_text) + " |")
            
        # 组合成Markdown表格字符串
        markdown_text = "\n".join(markdown_lines)
        return markdown_text, anchor_map, id_to_text_map

    def analyze_tables_with_llm(self, table_markdown, knowledge_context, id_to_text_map):
        # 动态判断知识库格式，生成不同的 Prompt 描述
        data_format_desc = "扁平化的 JSON 键值对（Key-Value）" if isinstance(knowledge_context, dict) else "按 Sheet（来源）分组的二维数组（矩阵）格式"
        
        prompt = f"""
        请分析以下带有锚点（格式如 {{{{ID_XXX}}}}）的文档内容（表格或段落），并结合提供的知识库数据，将正确的值填入对应的锚点。
        
        知识库数据（{data_format_desc}）：
        {json.dumps(knowledge_context, ensure_ascii=False, default=str)}
        
        文档内容结构（Markdown）：
        {table_markdown}
        
        锚点对应的原始文本（参考用，可能包含提示信息）：
        {json.dumps(id_to_text_map, ensure_ascii=False)}
        
        请仔细思考字段的对应关系。
        注意：知识库数据可能为以下两种格式之一：
        1. 按 Sheet（来源）分组的二维数组（矩阵）格式。
        2. 扁平化的 JSON 键值对（Key-Value）。
        
        请根据上下文自动推断。
        - 如果遇到包含换行符被拼接的字段名（如“现从事工作及专长”），请尝试模糊匹配。
        - **必须**参考提供的“锚点对应的原始文本”中的上下文线索（如“左侧Label”或“上方Label”）来确定填入内容。
        
        特别注意以下字段的提取和填充：
        1. **多键合并与分发**：
           - **合并**：如果表格中只有一个对应字段的单元格，但知识库中有多个相关键（如 "奖项_1", "奖项_2"），请将它们合并为一个完整段落（如 "1. A奖；2. B奖"）。
           - **分发**：如果表格中针对同一属性（如“爱好”）预留了**多个独立**的单元格（即多个锚点），且知识库中有对应的多条数据，请将数据**分散填入**不同的锚点，不要重复。
              - 例如：表格“爱好”下有 `{{ID_001}}` 和 `{{ID_002}}`，知识库有 `爱好_1: A`, `爱好_2: B`。
              - 正确：`"{{ID_001}}": "A"`, `"{{ID_002}}": "B"`
              - 错误：`"{{ID_001}}": "A, B"`, `"{{ID_002}}": "A, B"`
              - **注意**：如果数据条数少于单元格数（如只有 `爱好_1: A`），请只填入第一个锚点，其余锚点**不要**在返回的JSON中出现（即留空）。
        2. **复杂文本字段**：如“现从事工作及专长”、“何时何地受何奖励”。这些内容可能较长，包含多行，请务必提取完整。
        3. **基础信息字段**：如“工作单位”、“电子信箱”、“通讯地址”。这些信息可能散落在不同位置，请仔细查找。
        4. **格式处理**：如果源数据中包含换行符（\n），请根据目标表格的语境合理保留或替换为逗号/空格。

        如果一个字段在多个地方出现，请优先选择最匹配上下文的值。
        
        **禁止行为**：
        - 严禁在表格末尾或其他空位自动生成“总结”、“备注”或“额外说明”，除非表格中有明确的“备注”或“总结”标签指示。
        - 如果锚点没有明确的上下文指示（Label），且无法确定其对应关系，请保持为空，不要强行填入剩余的知识库信息。

        重要：请返回**完整**的填入内容。
        如果原始文本是占位符（如 "____"），直接返回填入值。
        如果原始文本包含提示信息且你需要保留（例如 "姓名：____"），请返回完整内容（例如 "姓名：张三"）。
        如果原始文本是日期格式（如 "____年__月__日"），请返回填充好的完整日期字符串（例如 "2024年1月1日"）。
        通常情况下，对于包含下划线的单元格，用户希望你填充内容并覆盖原有占位符。
        
        返回 JSON 格式：
        {{
            "{{{{ID_001}}}}": "填入的值1",
            "{{{{ID_002}}}}": "填入的值2",
            ...
        }}
        请确保只返回JSON格式数据，不要包含其他内容。
        """
        try:
            messages = [
                {"role": "system", "content": "你是一个专业的文档填充助手，擅长处理复杂表格和多层级数据映射。"},
                {"role": "user", "content": prompt}
            ]
            result = self.llm_client.chat_completion(messages, temperature=0.1)
            json_str = self._extract_json(result)
            return json.loads(json_str)
        except Exception as e:
            logger.error(f"表格分析失败: {str(e)}")
            return {}

    def _extract_json(self, text):
        try:
            json.loads(text)
            return text
        except json.JSONDecodeError:
            import re
            # 移除 replace('\n', '') 以保留 JSON 字符串中的换行符
            json_match = re.search(r'({.*})', text, re.DOTALL)
            if json_match:
                return json_match.group(1)
            raise ValueError("未找到有效JSON内容")

    def fill_document(self):
        if not self.doc or self.knowledge_dict is None:
            logger.error("文档或知识库未正确初始化")
            return False
        
        filled_count = 0
        
        # --- 1. 处理正文段落 ---
        logger.info("正在处理正文段落...")
        para_markdown, para_anchor_map, para_id_to_text_map = self._preprocess_paragraphs(self.doc.paragraphs)
        
        if para_anchor_map:
            logger.info(f"发现 {len(para_anchor_map)} 个段落填空位，正在请求 LLM 分析...")
            fill_map = self.analyze_tables_with_llm(para_markdown, self.knowledge_dict, para_id_to_text_map)
            
            for anchor_id, value in fill_map.items():
                if anchor_id in para_anchor_map:
                    try:
                        idx = para_anchor_map[anchor_id]
                        para = self.doc.paragraphs[idx]
                        
                        # 简单替换：全段替换
                        # 注意：这会丢失段落内的部分格式（如加粗），但保留段落整体样式
                        style = para.style
                        para.text = str(value)
                        para.style = style
                        
                        filled_count += 1
                        logger.debug(f"段落锚点 {anchor_id} 已填充值: {value}")
                    except Exception as e:
                        logger.error(f"段落填充异常: {anchor_id} - {str(e)}")
                else:
                    logger.warning(f"LLM返回了不存在的段落锚点ID: {anchor_id}")
        else:
            logger.info("正文段落中未发现填空位")

        # --- 2. 处理表格 ---
        for table_idx, table in enumerate(self.doc.tables):
            logger.info(f"正在处理第{table_idx + 1}个表格")
            
            # 第一阶段：Word 模板的“数字化”预处理
            table_markdown, anchor_map, id_to_text_map = self._preprocess_table(table)
            
            if not anchor_map:
                logger.info("该表格未发现填空位，跳过")
                continue

            # 第二阶段：LLM 分析并直接返回填充映射
            # 将知识库数据作为上下文传给 LLM
            fill_map = self.analyze_tables_with_llm(table_markdown, self.knowledge_dict, id_to_text_map)
            
            for anchor_id, value in fill_map.items():
                if anchor_id in anchor_map:
                    try:
                        row, col = anchor_map[anchor_id]
                        cell = table.cell(row, col)
                        
                        # 检查是否需要追加模式 (Append Mode)
                        # 如果原单元格包含类似 "1. xxx (不超过xx字)" 的指令性文本，且 LLM 返回的内容不包含该头信息，则追加
                        original_text = cell.text.strip()
                        should_append = False
                        
                        # 识别指令性表头特征：
                        # 1. 以数字开头 (1. / 1、 / 1 )
                        # 2. 包含字数限制说明 (不超过...字)
                        if len(original_text) > 0 and (
                            re.match(r'^\d+(\.\d+)*[、. ]', original_text) or 
                            re.search(r'[（(].*?不超过.*?字.*?[)）]', original_text)
                        ):
                            # 进一步检查：如果 LLM 返回的值已经包含了原文本（或者原文本的前一部分），则不需要追加，直接覆盖即可
                            # 简单的模糊检查
                            clean_val = str(value).strip()
                            if clean_val.startswith(original_text[:min(10, len(original_text))]):
                                should_append = False
                            else:
                                should_append = True
                        
                        if should_append:
                            logger.info(f"检测到指令性表头，采用追加模式: {anchor_id}")
                            # 追加新段落
                            # 注意：add_paragraph 会在单元格末尾添加新段落
                            cell.add_paragraph(str(value))
                        else:
                            # 尝试保留原有样式
                            if cell.paragraphs:
                                # 获取第一段的样式（如果有）
                                first_para = cell.paragraphs[0]
                                style = first_para.style
                                
                                # 清空并重写
                                cell.text = str(value)
                                
                                # 重新应用样式（简单尝试）
                                if cell.paragraphs:
                                    cell.paragraphs[0].style = style
                            else:
                                # 确保值为字符串
                                cell.text = str(value)
                            
                        filled_count += 1
                        logger.debug(f"表格锚点 {anchor_id} 已填充值: {value}")
                    except IndexError:
                        logger.error(f"无效的单元格位置: {anchor_map.get(anchor_id)}")
                    except Exception as e:
                        logger.error(f"字段填写异常: {anchor_id} - {str(e)}")
                else:
                    logger.warning(f"LLM返回了不存在的表格锚点ID: {anchor_id}")
                    
        logger.info(f"完成文档填充，共填写{filled_count}个字段")
        return True

    def save_document(self, filename=None):
        if not self.doc:
            logger.error("文档实例未初始化")
            return False
        if not filename:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            base_name = os.path.splitext(os.path.basename(self.word_template_path))[0]
            filename = f"{base_name}_filled_{timestamp}.docx"
        output_path = os.path.join(self.output_folder, filename)
        try:
            self.doc.save(output_path)
            logger.info(f"文档已保存至: {output_path}")
            return True
        except Exception as e:
            logger.error(f"文档保存失败: {str(e)}")
            return False

    def run(self):
        logger.info("启动自动化填表流程")
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