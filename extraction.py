import os
import pandas as pd
from docx import Document
import logging

import re

import json

logger = logging.getLogger(__name__)

def extract_content_to_json(docx_path, output_json_path, llm_client):
    """
    使用 LLM 智能提取 Word 文档内容，生成扁平化的 JSON 知识库
    """
    try:
        try:
            doc = Document(docx_path)
        except Exception as e:
            error_msg = str(e)
            if "no relationship of type" in error_msg or "File is not a zip file" in error_msg:
                raise ValueError(f"文件格式不正确。请确保您上传的是真正的 .docx 文件，而不是直接修改后缀名的 .doc 文件。({error_msg})")
            raise e
            
        full_text = []

        # 1. 收集所有文本（页眉、页脚、正文、表格）
        # 页眉
        for section in doc.sections:
            for para in section.header.paragraphs:
                if para.text.strip():
                    full_text.append(f"[页眉] {para.text.strip()}")
            for para in section.footer.paragraphs:
                if para.text.strip():
                    full_text.append(f"[页脚] {para.text.strip()}")

        # 正文
        for para in doc.paragraphs:
            if para.text.strip():
                full_text.append(para.text.strip())

        # 表格
        for i, table in enumerate(doc.tables):
            full_text.append(f"\n[表格 {i+1}]")
            for row in table.rows:
                # 使用 ' // ' 替换换行符，以便在保持单行结构的同时保留换行信息
                row_text = " | ".join([cell.text.strip().replace('\n', ' // ') for cell in row.cells if cell.text.strip()])
                if row_text:
                    full_text.append(row_text)
        
        doc_content = "\n".join(full_text)

        # 2. 调用 LLM 进行结构化提取
        prompt = f"""
        请分析以下文档内容，提取所有信息，将其整理为**扁平化的 JSON 键值对**。
        
        文档内容：
        {doc_content}
        
        要求：
        1. **区分事实与段落**：
           - 对于**短事实**（如姓名、电话、时间），请进行**细粒度拆分**。
             例如：“张三，男，1990年出生” -> {{"姓名": "张三", "性别": "男", "出生年份": "1990"}}
           - 对于**长段落/描述性内容**（如“成果简介”、“主要解决的教学问题”、“工作描述”），请**完整保留原话**，不要拆分，也不要过度总结。
             例如：{{"成果简介": "21世纪是创新创业的世纪...（保留完整段落）"}}
        2. **处理列表**：对于“主要贡献”、“奖励情况”等列表性内容，建议保留为带序号的完整文本，或拆分为 {{"奖励_1": "...", "奖励_2": "..."}}。
        3. **标准化键名**：键名请尽量简洁、规范，但对于特定的小标题（如“成果简介”），直接使用原标题作为键名。
        4. **不要遗漏**：请提取文档中的所有有效信息，包括页眉页脚中的联系方式。
        5. **格式还原**：文档中如果出现 ' // '，表示原文的换行符。在生成的 JSON 值中，请将其还原为标准的换行符（\\n）或保留合理的段落结构。
        
        请直接返回 JSON 对象，不要包含 Markdown 格式标记或其他废话。
        """

        messages = [
            {"role": "system", "content": "你是一个专业的数据提取专家，擅长将非结构化文档转化为结构化数据。"},
            {"role": "user", "content": prompt}
        ]

        response = llm_client.chat_completion(messages, temperature=0.1)
        
        # 清洗 JSON
        try:
            # 尝试找到第一个 { 和最后一个 }
            start = response.find('{')
            end = response.rfind('}') + 1
            if start != -1 and end != -1:
                json_str = response[start:end]
                data = json.loads(json_str)
            else:
                raise ValueError("未找到JSON内容")
        except Exception as e:
            logger.error(f"LLM 返回的 JSON 解析失败: {response}")
            # 兜底：如果解析失败，把原始内容存进去
            data = {"Raw_Content": doc_content, "Error": "JSON解析失败"}

        # 3. 保存为 JSON 文件
        with open(output_json_path, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=4)
            
        logger.info(f"成功使用 LLM 提取数据到 {output_json_path}")
        return True

    except Exception as e:
        logger.error(f"智能提取失败: {str(e)}")
        return False

def clean_cell_text(text):
    """
    清洗单元格文本：
    1. 去除首尾空白
    2. 尝试将混在一行的 Key-Value 对（如 '工作单位：xxx 职务：yyy'）拆分为多行
    """
    if not text:
        return ""
    
    text = text.strip()
    
    # 正则策略：查找 "空格 + 中文键名 + 冒号" 的模式，将其前面的空格替换为换行符
    # 键名限制为 2-6 个汉字，避免误伤普通句子
    # 处理中文冒号和英文冒号
    pattern = r'\s+([\u4e00-\u9fa5]{2,10}[：:])'
    text = re.sub(pattern, r'\n\1', text)
    
    return text

def extract_tables_from_docx(docx_path, output_excel_path):
    """
    从Word文档中提取表格和文本，保存为Excel文件
    """
    try:
        try:
            doc = Document(docx_path)
        except Exception as e:
            error_msg = str(e)
            if "no relationship of type" in error_msg or "File is not a zip file" in error_msg:
                raise ValueError(f"文件格式不正确。请确保您上传的是真正的 .docx 文件，而不是直接修改后缀名的 .doc 文件。({error_msg})")
            raise e

        # 1. 提取所有文本段落 (包括页眉页脚)
        text_content = []
        
        # 提取页眉页脚
        for section in doc.sections:
            # 页眉
            for para in section.header.paragraphs:
                 if para.text.strip():
                    text_content.append({"Content": para.text.strip(), "Type": "Header"})
            # 页脚
            for para in section.footer.paragraphs:
                 if para.text.strip():
                    text_content.append({"Content": para.text.strip(), "Type": "Footer"})

        # 正文段落
        for para in doc.paragraphs:
            if para.text.strip():
                text_content.append({"Content": para.text.strip(), "Type": "Text"})
        
        # 2. 提取表格
        tables_data = []
        for i, table in enumerate(doc.tables):
            # 尝试提取表格数据
            # 假设第一行是表头
            rows = []
            if len(table.rows) > 0:
                # 获取所有行数据
                for row in table.rows:
                    # 保留换行符，以便LLM更好地理解多行内容（如“主要贡献”、“奖励情况”）
                    # 同时尝试智能拆分混在一起的 KV 对
                    cell_data = [clean_cell_text(cell.text) for cell in row.cells]
                    rows.append(cell_data)
                
                # 直接转换为 DataFrame，不假设第一行是表头
                # 这样可以保留 Key-Value 类型的表格结构（如简历、登记表）
                # 同时也解决了表头包含换行符导致的问题
                df = pd.DataFrame(rows)
                tables_data.append(df)

        # 3. 写入Excel
        with pd.ExcelWriter(output_excel_path, engine='openpyxl') as writer:
            # 写入文本内容
            if text_content:
                df_text = pd.DataFrame(text_content)
                df_text.to_excel(writer, sheet_name="Text_Content", index=False)
            
            # 写入表格 (不写入 Header，不写入 Index)
            for i, df in enumerate(tables_data):
                sheet_name = f"Table_{i+1}"
                df.to_excel(writer, sheet_name=sheet_name, index=False, header=False)
                
        logger.info(f"成功从 {docx_path} 提取数据到 {output_excel_path}")
        return True
    except Exception as e:
        logger.error(f"提取数据失败: {str(e)}")
        return False
