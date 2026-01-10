import os
import json
import fitz  # PyMuPDF
import re
import pandas as pd
import shutil
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, simpledialog, ttk, Canvas
import threading
import subprocess
from datetime import datetime, timedelta
import math
import numpy as np
from PIL import Image
# import easyocr  # 移除这里的导入
from utils.common import log
import tempfile  # 添加临时文件模块


# ========== 日志写入函数 ==========
# 添加一个全局变量来跟踪是否是首次调用
_log_initialized = False

def write_log(message):
    """写入日志到文件"""
    global _log_initialized
    try:
        log_file = os.path.join(os.path.dirname(os.path.abspath(__file__)), "ocr_debug.log")
        
        # 如果是首次调用，清空日志文件
        mode = "w" if not _log_initialized else "a"
        if not _log_initialized:
            _log_initialized = True
        
        with open(log_file, mode, encoding="utf-8") as f:
            timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            f.write(f"[{timestamp}] {message}\n")
    except Exception as e:
        print(f"写入日志失败: {e}")


# ========== 基金代码修正函数 ==========
def correct_fund_code(raw_code):
    """修正OCR识别错误的基金代码
    
    Args:
        raw_code: 原始识别的基金代码
        
    Returns:
        corrected_code: 修正后的基金代码
    """
    if not raw_code or len(raw_code) < 6:
        return raw_code
    
    corrected_code = raw_code
    
    # 1. 修正首字母：如果第一个字符是8，改为B
    if corrected_code[0] == '8':
        corrected_code = 'B' + corrected_code[1:]
        print(f"基金代码首字母修正: 8 -> B")
    
    # 2. 修正数字1被识别为小写l的情况
    # 检查从第二个字符开始的所有字符
    for i in range(1, len(corrected_code)):
        if corrected_code[i] == 'l':
            corrected_code = corrected_code[:i] + '1' + corrected_code[i+1:]
            print(f"基金代码数字修正: 位置{i}的l -> 1")
    
    if corrected_code != raw_code:
        print(f"基金代码修正: {raw_code} -> {corrected_code}")
    
    return corrected_code

# ========== PDF转图像并OCR识别 ==========
def extract_text_with_easyocr(pdf_path):
    """使用 EasyOCR 提取PDF文本"""
    try:
        # 在函数内部导入easyocr，避免打包时的导入错误
        write_log(f"开始处理PDF: {pdf_path}")
        
        # 检查文件是否存在
        if not os.path.exists(pdf_path):
            write_log(f"错误：PDF文件不存在: {pdf_path}")
            return "", []
        
        write_log("正在导入easyocr...")
        import easyocr
        write_log("easyocr导入成功")
        
        # 初始化 EasyOCR（首次运行会下载模型）
        write_log("初始化 EasyOCR...")
        write_log("首次运行会下载模型文件，请稍等...")
        
        try:
            reader = easyocr.Reader(['ch_sim', 'en'], gpu=False)  # 中文简体和英文，不使用GPU
            write_log("EasyOCR 初始化成功")
        except Exception as e:
            write_log(f"EasyOCR初始化失败: {e}")
            import traceback
            traceback_info = traceback.format_exc()
            write_log(f"详细错误信息: {traceback_info}")
            return "", []
        
        write_log("正在打开PDF文件...")
        doc = fitz.open(pdf_path)
        write_log(f"PDF打开成功，共{len(doc)}页")
        
        all_text = []
        
        for page_num in range(len(doc)):
            write_log(f"正在处理页面 {page_num + 1}/{len(doc)}...")
            
            try:
                # 获取页面
                page = doc.load_page(page_num)
                write_log(f"页面{page_num + 1}加载成功")
                
                # 将页面转换为图像
                mat = fitz.Matrix(2.5, 2.5)  # 2.5倍缩放
                pix = page.get_pixmap(matrix=mat)
                write_log(f"页面{page_num + 1}转换为图像成功")
                
                # 使用系统临时目录保存临时文件
                with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as tmp:
                    temp_img_path = tmp.name
                
                write_log(f"临时文件路径: {temp_img_path}")
                
                # 保存图像到临时文件
                pix.save(temp_img_path)
                write_log(f"临时图像保存成功: {temp_img_path}")
                
                # 检查临时文件是否存在
                if not os.path.exists(temp_img_path):
                    write_log(f"错误：临时图像文件未生成: {temp_img_path}")
                    continue
                
                # 检查文件大小
                file_size = os.path.getsize(temp_img_path)
                write_log(f"临时文件大小: {file_size} 字节")
                
                # 使用 EasyOCR 识别
                write_log(f"开始OCR识别页面{page_num + 1}...")
                results = reader.readtext(temp_img_path)
                write_log(f"OCR识别完成，获得{len(results)}个结果")
                
                # 提取文本
                page_text = []
                for (bbox, text, prob) in results:
                    if prob > 0.3:  # 置信度阈值
                        write_log(f"识别文本: {text} (置信度: {prob:.2f})")
                        page_text.append(text)
                
                all_text.extend(page_text)
                write_log(f"页面 {page_num + 1} 完成，识别到 {len(page_text)} 行文本")
                
                # 删除临时文件
                try:
                    if os.path.exists(temp_img_path):
                        os.remove(temp_img_path)
                        write_log(f"临时文件已删除: {temp_img_path}")
                except Exception as e:
                    write_log(f"删除临时文件失败: {e}")
                    
            except Exception as e:
                write_log(f"处理页面{page_num + 1}时出错: {e}")
                import traceback
                traceback_info = traceback.format_exc()
                write_log(f"详细错误信息: {traceback_info}")
                continue
        
        doc.close()
        write_log(f"PDF处理完成，总共提取到{len(all_text)}行文本")
        return '\n'.join(all_text), all_text
        
    except Exception as e:
        write_log(f"EasyOCR识别失败: {str(e)}")
        import traceback
        traceback_info = traceback.format_exc()
        write_log(f"详细错误信息: {traceback_info}")
        return "", []

# ========== 提取函数 ==========
def extract_manual_fields_ocr(text, lines):
    """从OCR识别的文本中提取字段"""
    # 初始化返回值
    fund_market_code = ''  # 证券代码（从基金代码提取）
    amount = ''  # 申购金额
    
    # 将文本转换为单行，便于正则匹配
    full_text = ' '.join(lines).replace('\n', ' ')
    full_text_clean = full_text.replace(',', '').replace(' ', '')
    print(f"\n完整文本: {full_text[:500]}...")  # 打印前500个字符用于调试
    
    try:
        # 1. 提取产品名称（自动填充）

        
        # 2. 提取证券代码（从基金代码字段）
        fund_patterns = [
            r'基金代码[：:\s]*([B8]\d{5})',  # 允许B或8开头
            r'基金代码[：:\s]*([B8][0-9l]{5})',  # 允许数字1被识别为l
        ]
        for i, pattern in enumerate(fund_patterns):
            match = re.search(pattern, full_text)
            if match:
                raw_fund_code = match.group(1).strip()
                print(f"原始证券代码匹配成功 (模式{i+1}): {raw_fund_code}")
                
                # 应用修正函数
                fund_market_code = correct_fund_code(raw_fund_code)
                print(f"修正后证券代码: {fund_market_code}")
                break
        
        # 3. 提取确认金额
        amount_patterns = [
            r'确认金额[：:\s]*([\d, ]+\.?\d*)',
        ]
        for i, pattern in enumerate(amount_patterns):
            matches = re.findall(pattern, full_text)
            if matches:
                # 把所有逗号和空格去掉
                amounts = [m.replace(',', '').replace(' ', '') for m in matches]
                # 只保留能转为float的
                amounts = [float(a) for a in amounts if re.match(r'^\d+(\.\d+)?$', a)]
                if amounts:
                    amount = f"{max(amounts):.2f}"
                    print(f"确认金额匹配成功 (模式{i+1}): {amount}")
                    break
        
        print(f"\n最终提取结果:")
        print(f"证券代码: {fund_market_code}")
        print(f"确认金额: {amount}")
    
    except Exception as e:
        print(f"字段提取过程中出错: {str(e)}")
    
    return fund_market_code, amount

# ========== 红利除权提取主逻辑 ==========
def run_manual_dividend_extract(folder_path, json_path, log_text):
    # 1. 日期
    current_year = datetime.now().year
    today = datetime.now()
    today_str = today.strftime('%Y%m%d')
    yesterday_str = (today - timedelta(days=1)).strftime('%Y%m%d')
    target_cols = ['账套编号', '产品代码','市场代码', '凭证日期', '登记日期','派送金额','产品名称']

    # 3. 读取产品代码
    try:
        with open(json_path, 'r', encoding='utf-8') as f:
            product_code_dict = json.load(f)
    except Exception as e:
        log(f"产品代码加载失败: {e}", log_text)
        return False

    # 4. 遍历红利除权文件夹
    target_df = pd.DataFrame(columns=target_cols)
    target_path = os.path.join(folder_path, str(current_year), today_str, "1场外开基")
    log(f"正在查找分红单子目录: {target_path}", log_text)
    if not os.path.isdir(target_path):
        log(f"目标路径不存在！", log_text)
        return False

    # 添加文件计数变量
    total_files = 0
    processed_files = 0
    failed_files = []

    # 使用 os.walk 递归遍历 target_path 及其子文件夹
    has_target_folder = False
    for root, dirs, files in os.walk(target_path):
        # 仅处理路径名中包含“分红”的文件夹
        if not ("分红" in root):
            continue
            
        has_target_folder = True
        log(f"扫描目录: {root}", log_text)

        # 筛选出包含"万事如意"的PDF文件
        pdf_files = [f for f in files if f.lower().endswith('.pdf') and "万事如意" in f]
        total_files += len(pdf_files)
        
        for file in pdf_files:
            # 文件的完整路径现在使用 root
            file_path = os.path.join(root, file)
            log(f"正在处理文件: {file}", log_text)
            
            try:
                # 使用OCR提取文本
                text, lines = extract_text_with_easyocr(file_path)
                
                if not text:
                    log(f"文件 {file} 未能提取到任何文本", log_text)
                    failed_files.append(file_path)
                    continue

                # 判断是否为手工申购单据
                is_manual = ('万事如意分红' in file) or ('万事如意' in file)

                if is_manual:
                    fund_market_code, amount = extract_manual_fields_ocr(text, lines)
                    records = [(fund_market_code, amount)]
                else:
                    continue

                for fund_market_code, amount in records:
                    # 验证必要字段
                    if not fund_market_code or not amount:
                        log(f"文件 {file} 缺少必要字段，跳过", log_text)
                        continue
                    
                    temp_df = pd.DataFrame([{
                        '市场代码': fund_market_code,
                        '派送金额': amount
                    }])

                    # 强制转换为数值型
                    try:
                        temp_df['派送金额'] = pd.to_numeric(temp_df['派送金额'], errors='coerce').round(2)
                    except:
                        log(f"文件 {file} 金额格式转换失败", log_text)
                        continue
                    
                    # 映射账套编号
                    temp_df['产品名称'] = '万联资管万事如意FOF1号单一资产管理计划'
                    temp_df['账套编号'] = temp_df['产品名称'].map(product_code_dict)
                    # 填充其它字段
                    temp_df['凭证日期'] = yesterday_str
                    temp_df['登记日期'] = yesterday_str
                    temp_df['产品代码'] =''
                    # 调整列顺序，合并进总表
                    temp_df = temp_df[target_cols]
                    target_df = pd.concat([target_df, temp_df], ignore_index=True)

                # 成功处理后增加计数
                processed_files += 1
                log(f"文件 {file} 处理完成", log_text)

            except Exception as e:
                # 记录处理失败的文件
                failed_files.append(file_path)
                log(f"处理PDF失败: {file_path}", log_text)
                log(f"错误信息: {e}", log_text)
                continue

    if not has_target_folder and total_files == 0:
        log("没有找到包含'分红'的子文件夹。", log_text)
        return False

    # 显示最终处理结果
    if failed_files:
        log(f"有 {len(failed_files)} 个文件处理失败:", log_text)
        for failed_file in failed_files:
            log(f"- {failed_file}", log_text)
    else:
        log(f"所有 {total_files} 个文件都已成功处理", log_text)

    if target_df.empty:
        log("没有提取到任何有效数据。", log_text)
        return False

    # 保存输出
    output_folder = os.path.join(folder_path, str(current_year), today_str)
    if not os.path.isdir(output_folder):
        os.makedirs(output_folder, exist_ok=True)
    output_file = os.path.join(output_folder, "【境内理财产品】红利除权.xlsx")
    try:
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            target_df.to_excel(writer, sheet_name='Sheet1', index=False)
        log(f"已汇总输出到: {output_file}", log_text)
        return output_folder
    except Exception as e:
        log(f"写入Excel失败: {e}", log_text)
        return False

