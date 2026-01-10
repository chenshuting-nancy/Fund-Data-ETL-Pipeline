import os
import json
import pdfplumber
import re
import pandas as pd
from datetime import datetime, timedelta
from utils.common import log

def run_conversion_extract(folder_path, json_path, log_text):
    """运行超级转换确认单提取
    
    Args:
        folder_path: 文件夹路径
        json_path: 产品代码映射JSON文件路径
        log_text: 日志文本框对象
    """
    # 1. 日期
    current_year = datetime.now().year
    today = datetime.now()
    today_str = today.strftime('%Y%m%d')
    yesterday_str = (today - timedelta(days=1)).strftime('%Y%m%d')
    target_cols = ['产品代码',	'转出基金市场代码',	'转出基金交易市场',	'转出确认日期',	'转出份额',	'转出金额',	'转出费用',	'转入基金市场代码',	'转入基金交易市场',	'转入份额',	
                   '转入金额',	'资金账户',	'股东代码',	'席位代码',	'转入费用',	'退补款交收日',	'转入确认日期',	'产品名称','平台'
                    ]
    
    # 3. 读取产品代码
    # 3. 读取产品代码 - 使用专门的转换单映射文件
    conversion_json_path = json_path.replace('product_codes.json', 'product_codes_conversion.json')
    
    # 如果转换单映射文件不存在，尝试使用默认路径
    if not os.path.exists(conversion_json_path):
        conversion_json_path = os.path.join(os.path.dirname(json_path), 'product_codes_conversion.json')
    
    try:
        with open(conversion_json_path, 'r', encoding='utf-8') as f:
            product_code_dict = json.load(f)
        log(f"成功加载转换单产品代码映射文件: {conversion_json_path}", log_text)
    except Exception as e:
        log(f"转换单产品代码加载失败: {e}", log_text)
        log(f"请确保存在文件: {conversion_json_path}", log_text)
        return False
    

    # 4. 平台提取函数
    # ========== 定义按基金平台提取函数 ===============
    #京东肯特瑞
    def extract_jd_fields(lines):
        """
        从京东肯特瑞基金超级转换确认单中提取信息
        """
        # 1. 客户名称（产品名称）
        product_name = ''
        for line in lines:
            if '客户名称' in line:
                match = re.search(r'客户名称\s+(.*)', line)
                if match:
                    product_name = match.group(1).strip()
                    break
        
        # 2. 转出基金市场代码
        out_fund_code = ''
        for line in lines:
            if '转出基金代码' in line:
                match = re.search(r'转出基金代码\s+(\d{6})', line)
                if match:
                    out_fund_code = match.group(1)
                    break
        
        # 3. 转出金额
        out_amount = ''
        for line in lines:
            if '转出基金确认金额' in line:
                match = re.search(r'转出基金确认金额\s+([\d,]+\.\d+)', line)
                if match:
                    out_amount = match.group(1).replace(',', '')
                    break
        
        # 4. 转出份额
        out_shares = ''
        for line in lines:
            if '转出基金确认份额' in line:
                match = re.search(r'转出基金确认份额\s+([\d,]+\.\d+)', line)
                if match:
                    out_shares = match.group(1).replace(',', '')
                    break
        
        # 5. 转入基金市场代码
        in_fund_code = ''
        for line in lines:
            if '转入基金代码' in line:
                match = re.search(r'转入基金代码\s+(\d{6})', line)
                if match:
                    in_fund_code = match.group(1)
                    break
        
        # 6. 转入金额
        in_amount = ''
        for line in lines:
            if '转入基金确认金额' in line:
                match = re.search(r'转入基金确认金额\s+([\d,]+\.\d+)', line)
                if match:
                    in_amount = match.group(1).replace(',', '')
                    break
        
        # 7. 转入份额
        in_shares = ''
        for line in lines:
            if '转入基金确认份额' in line:
                match = re.search(r'转入基金确认份额\s+([\d,]+\.\d+)', line)
                if match:
                    in_shares = match.group(1).replace(',', '')
                    break

        #8. 转换手续费
        in_fee = ''
        for line in lines:
            if '转换手续费' in line:
                match = re.search(r'转换手续费\s+([\d,]+\.\d+)', line)
                if match:
                    in_fee = match.group(1).replace(',', '')
                    break
        
        return product_name, out_fund_code, out_amount, out_shares, in_fund_code, in_fee, in_amount, in_shares,"京东肯特瑞"

    # 天天基金提取函数
    # 天天基金提取函数 (修正手续费错行问题)
    def extract_tiantian_fields(lines):
        """
        从天天基金超级转换确认单中提取信息 (返回8个值 + 平台标识)
        """
        # 初始化字段
        product_name = ''
        out_fund_code = ''
        out_amount_str = '0'   # 暂存字符串用于后续计算
        out_shares = ''
        in_fund_code = ''
        in_fee_str = '0'       # 暂存字符串用于后续计算
        in_amount = ''
        in_shares = ''
        
        # --- 1. 提取产品名称 (处理断行) ---
        name_part1 = ''
        name_part2 = ''
        for i, line in enumerate(lines[:10]): 
            if '万联' in line and not name_part1:
                name_part1 = line.strip()
                for j in range(1, 4): 
                    if i+j < len(lines):
                        next_line = lines[i+j]
                        if '计划' in next_line or next_line.startswith('合资产'):
                            name_part2 = next_line.strip()
                            break
                break
        product_name = name_part1 + name_part2
        
        # --- 2. 提取转出信息 ---
        for i, line in enumerate(lines):
            if '转出基金代码' in line:
                match = re.search(r'转出基金代码\s+(\d{6})', line)
                if match:
                    out_fund_code = match.group(1)
            
            if '转出基金确认' in line and '金额' not in line: 
                if i + 1 < len(lines):
                    val_line = lines[i+1]
                    vals = re.findall(r'([\d,]+\.\d+)', val_line)
                    if len(vals) >= 2:
                        out_shares = vals[0].replace(',', '') 
                        out_amount_str = vals[1].replace(',', '') 
        
        # --- 3. 提取转入信息与手续费 (重点修改部分) ---
        for i, line in enumerate(lines):
            if '转入基金代码' in line:
                match = re.search(r'转入基金代码\s+(\d{6})', line)
                if match:
                    in_fund_code = match.group(1)
            
            # === 修改开始：增强的手续费提取逻辑 ===
            if '手续费' in line:
                fee_found = False
                # 策略 A: 在当前行查找 "手续费 123.45" 或 "手续费123.45"
                match = re.search(r'手续费\s*([\d,.]+)', line)
                if match:
                    # 排除掉只有"手续费"三个字后面没有数字的情况
                    # 有时候可能是 "手续费 转入..." 导致匹配不到数字
                    pass 
                
                # 如果当前行没找到数字，或者想更精准匹配
                # 尝试提取当前行所有的数字，看看是否合理
                current_line_vals = re.findall(r'([\d,]+\.\d+)', line)
                if current_line_vals:
                     in_fee_str = current_line_vals[0].replace(',', '')
                     fee_found = True
                
                # 策略 B: 如果当前行没有找到费用，且 i > 0，去上一行找
                # 针对案例：Line 20: '719.97(转换费：0,补差费'
                if not fee_found and i > 0:
                    prev_line = lines[i-1]
                    # 匹配模式：数字紧接着左括号 -> 123.45(
                    match_prev = re.search(r'([\d,.]+)\s*[\(（]', prev_line)
                    if match_prev:
                        in_fee_str = match_prev.group(1).replace(',', '')
            # === 修改结束 ===

            if '转入基金确认' in line and '份额' not in line:
                if i + 1 < len(lines):
                    val_line = lines[i+1]
                    vals = re.findall(r'([\d,]+\.\d+)', val_line)
                    if len(vals) >= 1:
                        in_shares = vals[0].replace(',', '') 

        # --- 4. 计算转入金额 ---
        try:
            o_amt_f = float(out_amount_str)
            i_fee_f = float(in_fee_str) if in_fee_str else 0.0
            in_amt_f = o_amt_f - i_fee_f
            in_amount = "{:.2f}".format(in_amt_f)
        except ValueError:
            in_amount = out_amount_str 

        return product_name, out_fund_code, out_amount_str, out_shares, in_fund_code, in_fee_str, in_amount, in_shares, "天天基金"

    # 5. 遍历确认单文件夹
    target_df = pd.DataFrame(columns=target_cols)
    target_path = os.path.join(folder_path, str(current_year), today_str, "1场外开基")
    log(f"正在查找确认单子目录: {target_path}", log_text)
    if not os.path.isdir(target_path):
        log(f"目标路径不存在！", log_text)
        return False

    # 添加文件计数变量
    total_files = 0
    processed_files = 0
    failed_files = []

    # 使用 os.walk 递归遍历 target_path 下的所有子文件夹
    # 递归遍历 target_path 下所有目录
    for root, dirs, files in os.walk(target_path):
        # 这一步是为了确保我们只关注包含确认单的路径，但现在 root 已经是完整路径
        # 我们应该检查 root 路径中是否包含 "确认"
        if "确认" not in root:
             continue
        log(f"扫描目录: {root}", log_text)
        
        pdf_files = [f for f in files if f.lower().endswith('.pdf') and ("超级" in f or "转换" in f)]
        total_files += len(pdf_files)
        
        for file in pdf_files:
            file_path = os.path.join(root, file)

            try:
                # 尝试处理PDF文件
                with pdfplumber.open(file_path) as pdf:
                    text = ''
                    for page in pdf.pages:
                        page_text = page.extract_text()
                        if page_text:
                            text += page_text
                lines = text.split('\n')

                # 判断平台，调用相应函数
                is_jd = ('肯特瑞基金' in file) or any('肯特瑞' in l for l in lines[:2])
                is_tiantian = ('天天基金' in file)
                
                records = [] # 初始化 records

                if is_jd:
                    # 京东的逻辑保持不变
                    p_name, o_code, o_amt, o_sh, i_code, in_fee, i_amt, i_sh, platform = extract_jd_fields(lines)
                    records = [(p_name, o_code, o_amt, o_sh, i_code, in_fee, i_amt, i_sh ,platform)]
                    
                elif is_tiantian:
                    # 调用天天基金提取
                    p_name, o_code, o_amt, o_sh, i_code, in_fee, i_amt, i_sh, platform = extract_tiantian_fields(lines)
                    records = [(p_name, o_code, o_amt, o_sh, i_code, in_fee, i_amt, i_sh ,platform)]
                else:
                    continue

                for product_name, out_fund_code, out_amount, out_shares, in_fund_code, in_fee, in_amount, in_shares, platform in records:
                    temp_df = pd.DataFrame([{
                        '产品名称': product_name,
                        '转出基金市场代码': out_fund_code,
                        '转出份额': out_shares,
                        '转出金额': out_amount,
                        '转入基金市场代码': in_fund_code,
                        '转入份额': in_shares,
                        '转入金额': in_amount,
                        '转入费用': in_fee,
                        '平台': platform

                    }])

                    #强制转换为数值型
                    temp_df['转出份额'] = pd.to_numeric(temp_df['转出份额'], errors='coerce').round(2)
                    temp_df['转出金额'] = pd.to_numeric(temp_df['转出金额'], errors='coerce').round(2)
                    temp_df['转入份额'] = pd.to_numeric(temp_df['转入份额'], errors='coerce').round(2)
                    temp_df['转入金额'] = pd.to_numeric(temp_df['转入金额'], errors='coerce').round(2)
                    temp_df['转入费用'] = pd.to_numeric(temp_df['转入费用'], errors='coerce').round(2)
                    # 映射账套编号
                    temp_df['产品代码'] = temp_df['产品名称'].map(product_code_dict)
                    # 填充其它字段
                    temp_df['转出基金交易市场'] = '国内银行间'
                    temp_df['转入基金交易市场'] = '国内银行间'
                    temp_df['转出确认日期'] = today_str
                    temp_df['转出费用'] = ''
                    temp_df['资金账户'] = ''
                    temp_df['股东代码'] = ''
                    temp_df['席位代码'] = ''
                    temp_df['退补款交收日'] = ''
                    temp_df['转入确认日期'] = ''

                    # 调整列顺序，合并进总表
                    temp_df = temp_df[target_cols]
                    target_df = pd.concat([target_df, temp_df], ignore_index=True)

                # 成功处理后增加计数
                processed_files += 1

            except Exception as e:
                # 记录处理失败的文件
                failed_files.append(file_path)
                log(f"处理PDF失败: {file_path}", log_text)
                log(f"错误信息: {e}", log_text)
                continue

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
    output_file = os.path.join(output_folder, "【境内基金业务】超级转换确认.xls")
    try:
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            target_df.to_excel(writer, sheet_name='Sheet1', index=False)
        log(f"已汇总输出到: {output_file}", log_text)
        return output_folder
    except Exception as e:
        log(f"写入Excel失败: {e}", log_text)
        return False