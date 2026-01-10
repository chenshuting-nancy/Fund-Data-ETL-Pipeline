import os
import json
import pdfplumber
import re
import pandas as pd
from datetime import datetime, timedelta
from utils.common import log

def run_purchase_confirm_extract(folder_path, json_path, log_text):
    """运行申购确认单提取
    
    Args:
        folder_path: 文件夹路径
        json_path: 产品代码映射JSON文件路径
        log_text: 日志文本框对象
    """
    # 1. 日期
    current_year = datetime.now().year
    today = datetime.now()
    today_str = today.strftime('%Y%m%d')
    target_cols = ['账套编号', '基金市场代码', '交易市场', '日期','业务类别','数量','金额', '手续费','佣金','交易对手','资金账户','赎回到账日期', '股东账户','席位号','产品名称','基金平台']
    
    # 3. 读取产品代码
    try:
        with open(json_path, 'r', encoding='utf-8') as f:
            product_code_dict = json.load(f)
    except Exception as e:
        log(f"产品代码加载失败: {e}", log_text)
        return False

    # 4. 平台提取函数
    # ========== 定义按基金平台提取函数 ===============
    #好买
    def extract_haomai_fields(text, lines):
        # 1. 账户名称（产品名称）
        product_name = ''
        for i, line in enumerate(lines):
            if '账户名称' in line:
                prev_line = lines[i-1].strip() if i > 0 else ''
                next_line = lines[i+1].strip() if i+1 < len(lines) else ''
                if prev_line and '制单人' not in prev_line and '好买基金' not in prev_line:
                    product_name += prev_line
                if next_line and '证件类型' not in next_line and '产品代码' not in next_line:
                    product_name += next_line
                break
        product_name = product_name.replace(' ', '').replace('\u3000', '')

        # 2. 产品代码
        m2 = re.search(r'产品代码[：: ]*([0-9]{6})', text)
        fund_market_code = m2.group(1).strip() if m2 else ''

        # 3. 确认金额
        m3 = re.search(r'确认金额[：: ]*([\d,]+\.\d+)', text)
        amount = m3.group(1).replace(',', '') if m3 else ''

        # 4. 确认份额
        m4 = re.search(r'确认份额[：: ]*([\d,]+\.\d+)', text)
        shares = m4.group(1).replace(',', '') if m4 else ''
        
        # 5. 手续费
        m6 = re.search(r'手续费[：: ]*([\d,]+\.\d+)', text)
        fee = m6.group(1).replace(',', '') if m6 else ''
   
        return product_name, fund_market_code, amount, shares, fee, '好买基金'

    #天天
    def extract_tiantian_fields(lines):

        # 1. 产品名称（账户户名，跨行拼接）
        product_name = ''
        for i, line in enumerate(lines):
            if '账户户名' in line:
                # 账户户名一般在本行和上一行
                # 先看上一行
                prev_line = lines[i-1].strip() if i > 0 else ''
                this_name = line.split('账户户名')[0].strip()
                # 再拼接下方的“产管理计划”等
                next_line = lines[i+1].strip() if i+1 < len(lines) else ''
                # 合成
                # 判断方案：一般产品名在上一行+本行（账户户名前）+下行
                cand = ''
                if prev_line and '确认单' not in prev_line:
                    cand += prev_line
                if this_name:
                    cand += this_name
                # “产管理计划”大概率在下行
                if ('产管理计划' in next_line) or (next_line and '账户类型' not in next_line):
                    cand += next_line
                product_name = cand.replace(' ', '').replace('\u3000', '')
                break

        # 2. 基金市场代码（基金代码，文本格式）
        m2 = re.search(r'基金代码[：: ]*([0-9]{6})', text)
        fund_market_code = m2.group(1).strip() if m2 else ''

        # 3. 确认金额
        m3 = re.search(r'确认金额[：: ]*([\d,]+\.\d+)', text)
        amount = m3.group(1).replace(',', '') if m3 else ''

        # 4. 确认份额
        m4 = re.search(r'确认份额[：: ]*([\d,]+\.\d+)', text)
        shares = m4.group(1).replace(',', '') if m4 else ''

        # 5. 手续费
        m6 = re.search(r'确认费用[：: ]*([\d,]+\.\d+)', text)
        fee = m6.group(1).replace(',', '') if m6 else ''

        return product_name, fund_market_code, amount, shares, fee, '天天基金'

    #利得
    def extract_lide_fields(lines):

        # 1. 产品名称（投资者姓名/名称）
        product_name = ''
        for line in lines:
            if '投资者姓名/名称' in line:
                match = re.search(r'投资者姓名/名称[:：]\s*(.*)', line)
                if match:
                    product_name = match.group(1).replace(' ', '').replace('\u3000', '')
                    break

        # 2. 基金市场代码（基金代码）
        m2 = re.search(r'基金代码[：: ]*([0-9]{6})', text)
        fund_market_code = m2.group(1).strip() if m2 else ''

        # 3. 确认金额
        m3 = re.search(r'确认金额（元）[：: ]*([\d,]+\.\d+)', text)
        amount = m3.group(1).replace(',', '') if m3 else ''

        # 4. 确认份额
        m4 = re.search(r'确认份额（份）[：: ]*([\d,]+\.\d+)', text)
        shares = m4.group(1).replace(',', '') if m4 else ''
        
        # 5. 手续费
        m6 = re.search(r'交易费用（元）[：: ]*([\d,]+\.\d+)', text)
        fee = m6.group(1).replace(',', '') if m6 else ''
        

        return product_name, fund_market_code, amount, shares, fee, '利得基金'

    #长量
    def extract_changliang_fields(lines):

        # 1. 产品名称（投资者名称）
        product_name = ''
        for line in lines:
            if '投资者名称' in line:
                match = re.search(r'投资者名称\s*(.*)', line)
                if match:
                    # 去掉可能的多余空格
                    product_name = match.group(1).strip().replace(' ', '').replace('\u3000', '')
                    break

        # 2. 基金市场代码（基金代码，文本格式）
        m2 = re.search(r'基金代码[：: ]*([0-9]{6})', text)
        fund_market_code = m2.group(1).strip() if m2 else ''

        # 3. 确认金额 - 注意长量基金的格式包括"(元)"
        m3 = re.search(r'确认金额[：: ]*([\d,]+\.\d+).*?\(元\)', text)
        amount = m3.group(1).replace(',', '') if m3 else ''

        # 4. 确认份额 - 注意长量基金的格式包括"(份)"
        m4 = re.search(r'确认份额[：: ]*([\d,]+\.\d+).*?\(份\)', text)
        shares = m4.group(1).replace(',', '') if m4 else ''

        # 5. 手续费
        m6 = re.search(r'手续费[：: ]*([\d,]+\.\d+).*?\(元\)', text)
        fee = m6.group(1).replace(',', '') if m6 else ''
        
        return product_name, fund_market_code, amount, shares, fee, '长量基金'

    #盈米
    def extract_yingmi_fields(lines):
        """
        从盈米基金赎回确认PDF解析行中提取所有赎回确认记录。
        返回: 列表，每个元素为 (产品名称, 基金市场代码, 金额, 份额, 到账日期)
        """
        # 1. 提取产品名称（投资者名称）
        product_name = ''
        for idx, line in enumerate(lines):
            if ('投资者名称' in line) and ('投资者类型' in line):
                name_parts = []
                if idx-1 >= 0:
                    prev = lines[idx-1].strip()
                    if prev and ('公司' not in prev and '信息' not in prev):
                        name_parts.append(prev)
                if idx+1 < len(lines):
                    next_ = lines[idx+1].strip()
                    if next_ and (len(next_) < 25):  # 经验过滤
                        name_parts.append(next_)
                if name_parts:
                    product_name = ''.join(name_parts)
                    break
        if not product_name:
            for line in lines:
                if '投资者名称' in line:
                    match = re.search(r'投资者名称\s*([^\s]+)', line)
                    if match:
                        product_name = match.group(1).replace(' ', '').replace('\u3000', '')
                        break

        # 2. 提取所有确认记录
        results = []
        
        # 找到所有交易序号行，作为记录的开始
        transaction_start_indices = []
        for i, line in enumerate(lines):
            if '交易序号' in line and '交易类型' in line and '申购' in line:
                transaction_start_indices.append(i)
                print(f"找到交易序号行: {i}: {line}")
        
        # 如果没有找到交易序号行，尝试其他方式查找
        if not transaction_start_indices:
            for i, line in enumerate(lines):
                if ('交易类型' in line and '申购' in line) or ('交易类型：' in line and '申购' in line):
                    transaction_start_indices.append(i)
                    print(f"找到交易类型行: {i}: {line}")
        
        # 处理每条交易记录
        for start_idx in transaction_start_indices:
            # 定义处理范围（从当前索引到下一个交易序号或文件结束）
            end_idx = len(lines)
            for next_start in transaction_start_indices:
                if next_start > start_idx:
                    end_idx = next_start
                    break
            
            # 在当前记录范围内搜索所需信息
            record_lines = lines[start_idx:end_idx]
            record_text = '\n'.join(record_lines)
            
            # 初始化变量
            fund_market_code = ''
            amount = ''
            shares = ''
            fee = ''
            est_date = ''
            
            # 提取基金代码
            code_match = re.search(r'基金代码[:：]\s*(\d{6})', record_text)
            if code_match:
                fund_market_code = code_match.group(1)
            
            # 提取确认金额
            amount_match = re.search(r'确认金额[:：]?\s*([\d,]+\.\d+)', record_text)
            if amount_match:
                amount = amount_match.group(1).replace(',', '')
            
            # 提取确认份额
            shares_match = re.search(r'确认份额[:：]?\s*([\d,]+\.\d+)', record_text)
            if shares_match:
                shares = shares_match.group(1).replace(',', '')
            
            # 提取手续费
            fee_match = re.search(r'手续费[:：]?\s*([\d,]+\.\d+)', record_text)
            if fee_match:
                fee = fee_match.group(1).replace(',', '')

            
            print(f"记录 {start_idx}: 代码={fund_market_code}, 金额={amount}, 份额={shares}, 手续费={fee}")
            
            # 如果找到了基金代码和金额/份额，添加到结果
            if fund_market_code and (amount or shares):
                results.append((product_name, fund_market_code, amount, shares, fee, '盈米基金'))
            else:
                print(f"记录 {start_idx} 缺少关键信息，跳过")
        
        print(f"共找到 {len(results)} 条有效记录")
        return results

    #交行
    def extract_jiaohang_fields(lines):
    
        # 1. 产品名称（从投资者信息字段提取）
        product_name = ''
        for line in lines:
            if '投资者信息' in line:
                # 提取投资者信息后的产品名称
                parts = line.split('投资者信息')
                if len(parts) > 1:
                    product_name = parts[1].strip()
                    break
        
        # 清理产品名称中的多余空格和特殊字符
        product_name = product_name.replace(' ', '').replace('\u3000', '').replace('\n', '')

        # 2. 基金市场代码（从产品代码字段提取）
        fund_market_code = ''
        for line in lines:
            if '产品代码' in line:
                # 提取6位数字的基金代码
                match = re.search(r'产品代码\s*([0-9]{6})', line)
                if match:
                    fund_market_code = match.group(1)
                    break

        # 3. 确认金额（从确认金额字段提取）
        amount = ''
        for line in lines:
            if '确认金额' in line:
                # 提取金额数值，支持带小数点的格式
                match = re.search(r'确认金额\s*([\d,]+\.?\d*)', line)
                if match:
                    amount = match.group(1).replace(',', '')
                    break

        # 4. 确认份额（从确认份额字段提取）
        shares = ''
        for line in lines:
            if '确认份额' in line:
                # 提取份额数值，支持带小数点的格式
                match = re.search(r'确认份额\s*([\d,]+\.?\d*)', line)
                if match:
                    shares = match.group(1).replace(',', '')
                    break

        # 5. 手续费
        fee = ''
        for line in lines:
            if '认申购手续费' in line:
                # 提取金额数值，支持带小数点的格式
                match = re.search(r'认申购手续费\s*([\d,]+\.?\d*)', line)
                if match:
                    fee = match.group(1).replace(',', '')
                    break

        return product_name, fund_market_code, amount, shares, fee, '交e通'

    #京东肯特瑞
    def extract_jd_fields(lines):
    
        # 1. 产品名称（从客户名称字段提取）
        product_name = ''
        for line in lines:
            if '客户名称' in line:
                # 提取客户名称后的产品名称
                parts = line.split('客户名称')
                if len(parts) > 1:
                    product_name = parts[1].strip()
                    break
        
        # 清理产品名称中的多余空格和特殊字符
        product_name = product_name.replace(' ', '').replace('\u3000', '').replace('\n', '')

        # 2. 基金市场代码（从基金代码字段提取）
        fund_market_code = ''
        for line in lines:
            if '基金代码' in line:
                # 提取6位数字的基金代码
                match = re.search(r'基金代码\s*([0-9]{6})', line)
                if match:
                    fund_market_code = match.group(1)
                    break

        # 3. 确认金额（从确认金额字段提取，去除逗号）
        amount = ''
        for line in lines:
            if '确认金额' in line:
                # 提取金额，支持带逗号的格式，并去除逗号
                match = re.search(r'确认金额\s*([\d,]+\.?\d*)', line)
                if match:
                    amount = match.group(1).replace(',', '')
                    break

        # 4. 确认份额（从确认份额字段提取，去除逗号）
        shares = ''
        for line in lines:
            if '确认份额' in line:
                # 提取份额，支持带逗号的格式，并去除逗号
                match = re.search(r'确认份额\s*([\d,]+\.?\d*)', line)
                if match:
                    shares = match.group(1).replace(',', '')
                    break

        # 5. 手续费
        fee = ''
        for line in lines:
            if '手续费' in line :
                match = re.search(r'手续费\s*([\d,]+\.?\d*)', line)
                if match:
                    fee = match.group(1).replace(',', '')
                    break

        return product_name, fund_market_code, amount, shares, fee, '京东肯特瑞'

    #网金基金
    def extract_wangjin_fields(lines):
        
        # 1. 产品名称（从投资者名称字段提取）
        product_name = ''
        for line in lines:
            if '投资者名称' in line:
                # 提取投资者名称后的产品名称
                parts = line.split('投资者名称')
                if len(parts) > 1:
                    product_name = parts[1].strip()
                    break
        
        # 清理产品名称中的多余空格和特殊字符
        product_name = product_name.replace(' ', '').replace('\u3000', '').replace('\n', '')

        # 2. 基金市场代码（从基金代码字段提取）
        fund_market_code = ''
        for line in lines:
            if '基金代码' in line:
                # 提取6位数字的基金代码
                match = re.search(r'基金代码\s*([0-9]{6})', line)
                if match:
                    fund_market_code = match.group(1)
                    break

        # 3. 确认金额（从确认金额字段提取）
        amount = ''
        for i, line in enumerate(lines):
            if '申购金额' in line and ('小写' in line):
                # 尝试多种匹配模式
                # 模式1：申购金额（小写） 38,000,000.00
                match1 = re.search(r'申购金额[（(]?小写[）)]?\s*([0-9,]+\.?[0-9]*)', line)
                # 模式2：申购金额小写.壹 (后面跨行显示金额)
                match2 = re.search(r'申购金额小写[^0-9]*([0-9,]+\.?[0-9]*)', line)
                
                if match1:
                    amount = match1.group(1).replace(',', '')
                elif match2:
                    amount = match2.group(1).replace(',', '')
                else:
                    # 如果当前行没有完整金额，检查下一行
                    if i + 1 < len(lines):
                        next_line = lines[i + 1].strip()
                        # 检查下一行是否是金额格式
                        amount_match = re.match(r'^([0-9,]+\.?[0-9]*)$', next_line)
                        if amount_match:
                            amount = amount_match.group(1).replace(',', '')
                break

        # 4. 确认份额（从"确认净额"行提取第一个数字）
        shares = ''
        for line in lines:
            if '确认净额' in line:  
                # 提取这一行中的所有数字
                matches = re.findall(r'([0-9,]+\.?[0-9]*)', line)
                if matches:
                    # 第一个数字是确认份额
                    shares = matches[0].replace(',', '')
                    break

        # 5. 手续费
        fee = ''
        for line in lines:
            if '费开户' in line:
                # 提取手续费
                match = re.search(r'费开户\s*([\d,]+\.?\d*)', line)
                if match:
                    fee = match.group(1).replace(',', '')
                    break

        return product_name, fund_market_code, amount, shares, fee, '网金基金'

    #平安行E通
    def extract_pingan_fields(lines):
        
        # 1. 产品名称（从账户名称字段提取，处理跨多行情况）
        product_name = ''
        for i, line in enumerate(lines):
            if '账户名称' in line:
                # 先获取当前行账户名称后的内容
                parts = line.split('账户名称')
                if len(parts) > 1:
                    current_part = parts[1].strip()
                    product_name += current_part
                
                # 检查后续行是否包含产品名称的剩余部分
                j = i + 1
                while j < len(lines) and j < i + 5:  # 最多检查后续4行
                    next_line = lines[j].strip()
                    # 如果遇到其他字段标识，停止拼接
                    if (any(keyword in next_line for keyword in 
                        ['开户行名称', '投资主体产品名称', '基金代码', '申请日期', 
                            '确认金额', '手续费', '交易状态', '经办人', '特别说明']) or
                        len(next_line) == 0):
                        break
                    # 拼接产品名称
                    product_name += next_line
                    j += 1
                break
        
        # 清理产品名称中的多余空格和特殊字符
        product_name = product_name.replace(' ', '').replace('\u3000', '').replace('\n', '')

        # 2. 基金市场代码（从基金代码字段提取）
        fund_market_code = ''
        for line in lines:
            if '基金代码' in line:
                # 提取6位数字的基金代码
                match = re.search(r'基金代码\s*([0-9]{6})', line)
                if match:
                    fund_market_code = match.group(1)
                    break

        # 3. 确认金额（从确认金额字段提取，去除逗号和"元"字符）
        amount = ''
        for line in lines:
            if '确认金额' in line:
                # 提取金额，支持带逗号的格式，并去除逗号和"元"字符
                match = re.search(r'确认金额\s*([\d,]+\.?\d*)元?', line)
                if match:
                    amount = match.group(1).replace(',', '')
                    break

        # 4. 确认份额（从确认份额字段提取，去除逗号和"份"字符）
        shares = ''
        for line in lines:
            if '确认份额' in line:
                # 提取份额，支持带逗号的格式，并去除逗号和"份"字符
                match = re.search(r'确认份额\s*([\d,]+\.?\d*)份?', line)
                if match:
                    shares = match.group(1).replace(',', '')
                    break

        # 5. 手续费
        fee = ''
        for line in lines:
            if '手续费' in line:
                match = re.search(r'手续费\s*([\d,]+\.?\d*)元?', line)
                if match:
                    fee = match.group(1).replace(',', '')
                    break

        return product_name, fund_market_code, amount, shares, fee, '平安行E通'

    #建行直销
    def extract_jianhang_fields(lines):
      
        # 1. 产品名称（从客户名称字段提取）
        product_name = ''
        for line in lines:
            if '客 户 名 称' in line:
                # 提取客户名称后的产品名称
                match = re.search(r'客\s*户\s*名\s*称\s*[：:]\s*(.*)', line)
                if match:
                    # 去除可能的框线字符和多余空格
                    product_name = match.group(1).strip().replace('┃', '').replace(' ', '').replace('\u3000', '')
                    break
        
        # 清理产品名称中的多余空格和特殊字符
        product_name = product_name.replace(' ', '').replace('\u3000', '').replace('\n', '')

        # 2. 基金市场代码（从基金代码字段提取）
        fund_market_code = ''
        for line in lines:
            if '基 金 代 码' in line:
                # 提取6位数字的基金代码
                match = re.search(r'基\s*金\s*代\s*码\s*[：:]\s*([0-9]{6})', line)
                if match:
                    fund_market_code = match.group(1)
                    break

        # 3. 确认金额（从确 认 金 额字段提取）
        amount = ''
        for line in lines:
            if '确 认 金 额' in line:
                # 提取确 认 金 额数值
                match = re.search(r'确\s*认\s*金\s*额\s*[：:]?\s*([\d,]+\.?\d*)', line)
                if match:
                    amount = match.group(1).replace(',', '')
                    break
        
        # 4. 确认份额（从确 认 份 额字段提取）
        shares = ''
        for line in lines:
            if '确 认 份 额' in line:
                # 提取确 认 份 额数值
                match = re.search(r'确\s*认\s*份\s*额\s*[：:]?\s*([\d,]+\.?\d*)', line)
                if match:
                    shares = match.group(1).replace(',', '')
                    break
        
        # 5. 手续费（从手   续  费字段提取）
        fee = ''
        for line in lines:
            # 去除所有空格后再判断
            if '手续费' in line.replace(' ', ''):
                # 提取这一行中的数字
                match = re.search(r'([\d,]+\.?\d*)', line)
                if match:
                    fee = match.group(1).replace(',', '')
                    break

        return product_name, fund_market_code, amount, shares, fee, '建行直销'

    #腾元基金   
    def extract_tengyuan_fields(lines):
        
        # 1. 产品名称（从客户名称字段提取）
        product_name = ''
        for line in lines:
            if '客户名称' in line:
                # 处理表格格式，提取│后的内容
                parts = line.split('│')
                if len(parts) >= 2:
                    # 获取客户名称部分，去除可能的表格符号
                    product_name = parts[1].strip()
                    # 去除可能的结尾表格符号
                    product_name = product_name.replace('┃', '').strip()
                    break
        
        # 清理产品名称中的多余空格和特殊字符
        product_name = product_name.replace(' ', '').replace('\u3000', '')
        
        # 2. 基金市场代码（从基金代码字段提取）
        fund_market_code = ''
        for line in lines:
            if '基金代码' in line:
                # 处理表格格式，提取│后的6位数字
                parts = line.split('│')
                for part in parts:
                    # 查找6位数字的基金代码
                    match = re.search(r'(\d{6})', part)
                    if match:
                        fund_market_code = match.group(1)
                        break
                if fund_market_code:
                    break
        
        # 3. 确认金额和确认份额（在同一行）
        amount = ''
        shares = ''
        for line in lines:
            if '确认金额' in line and '确认份额' in line and '│' in line:
                # 第18行: '┃确认金额 │8,000,000.00 │确认份额 │6,932,743.98 ┃'
                parts = line.split('│')
                if len(parts) >= 4:
                    # 确认金额在第2个部分
                    amount_match = re.search(r'([\d,]+\.?\d*)', parts[1])
                    if amount_match:
                        amount = amount_match.group(1).replace(',', '')
                    
                    # 确认份额在第4个部分
                    shares_match = re.search(r'([\d,]+\.?\d*)', parts[3])
                    if shares_match:
                        shares = shares_match.group(1).replace(',', '')
                break
            
        # 4. 手续费（从手续费字段提取）
        fee = ''
        for line in lines:
            if '手' in line and '续' in line and '费' in line and '│' in line:
                # 第20行: '┃单位净值 │1.15380 │手 续 费 │1,000.00 ┃'
                parts = line.split('│')
                if len(parts) >= 4:
                    # 手续费在第4个部分
                    fee_match = re.search(r'([\d,]+\.?\d*)', parts[3])
                    if fee_match:
                        fee = fee_match.group(1).replace(',', '')
                        break
        
        return product_name, fund_market_code, amount, shares, fee, '腾元基金'

    #联泰基金
    def extract_liantai_fields(lines):
        product_name = ''
        for line in lines:
            if '投资账户' in line:
                    match = re.search(r'投资账户\s*([^\s]+)', line)
                    if match:
                        product_name = match.group(1).replace(' ', '').replace('\u3000', '')
                        break

        results = []
        i = 0
        N = len(lines)
        while i < N:
            line = lines[i]
            # 联泰基金使用 "交易信息（X/Y）" 来标记每条记录的开始
            if '交易信息' in line:
                fund_market_code = ''
                amount = ''
                shares = ''
                fee = ''
                est_date = ''
                
                # 在接下来的几行中查找基金代码和金额、份额
                lookahead = 1
                while (i + lookahead) < N and lookahead <= 8:  # 增加查找范围
                    subline = lines[i + lookahead]
                    
                    # 查找基金代码
                    if '基金代码' in subline:
                        # 提取基金代码（支持6位数字）
                        match = re.search(r'基金代码\s+([0-9]{6})', subline)
                        if match:
                            fund_market_code = match.group(1)
                    
                    # 查找确认金额
                    if '确认金额(元)' in subline:
                        match = re.search(r'确认金额\(元\)\s*([\d,]+\.?\d*)', subline)
                        if match:
                            amount = match.group(1).replace(',', '')

                    # 查找确认份额
                    if '确认份额(份)' in subline:
                        match = re.search(r'确认份额\(份\)\s*([\d,]+\.?\d*)', subline)
                        if match:
                            shares = match.group(1).replace(',', '')

                    # 查找手续费
                    if '手续费(元)' in subline:
                        match = re.search(r'手续费\(元\)\s*([\d,]+\.?\d*)', subline)
                        if match:
                            fee = match.group(1).replace(',', '')
                            
                    # 如果遇到下一个交易信息块，停止查找
                    if lookahead > 1 and '交易信息' in subline:
                        break
                        
                    lookahead += 1
                
                # 如果找到了基金代码和红利份额，添加到结果中
                if fund_market_code and amount :
                    results.append((product_name, fund_market_code, amount, shares, fee, '联泰基金'))
            
            i += 1
        
        return results

    #融联创同业交易平台
    def extract_ronglianchuang_fields(lines):
        # 1. 产品名称（来款账号名称，可能跨行）
        product_name = ''
        for i, line in enumerate(lines):
            if '来款账号名称' in line:
                # 当前行“来款账号名称”后的部分
                parts = line.split('来款账号名称')
                if len(parts) > 1:
                    product_name = parts[1].strip().lstrip(':：')
                # 继续拼接下一行（比如“计划”）
                if i + 1 < len(lines):
                    next_line = lines[i + 1].strip()
                    if next_line and not any(k in next_line for k in ['大额支付行号', '产品代码']):
                        product_name += next_line
                break
        product_name = product_name.replace(' ', '').replace('\u3000', '')

        # 2. 基金市场代码（产品代码）
        fund_market_code = ''
        for line in lines:
            if '产品代码' in line:
                match = re.search(r'产品代码\s*[:：]?\s*([0-9]{6})', line)
                if match:
                    fund_market_code = match.group(1)
                    break

        # 3. 确认金额（去掉逗号，去掉“元”）
        amount = ''
        for line in lines:
            if '确认金额' in line:
                match = re.search(r'确认金额\s*[:：]?\s*([\d,]+\.\d+)', line)
                if match:
                    amount = match.group(1).replace(',', '')
                    break

        # 4. 确认份额（去掉逗号）
        shares = ''
        for line in lines:
            if '确认份额' in line:
                match = re.search(r'确认份额\s*[:：]?\s*([\d,]+\.\d+)', line)
                if match:
                    shares = match.group(1).replace(',', '')
                    break

        # 5. 手续费（去掉逗号，去掉“元”）
        fee = ''
        for line in lines:
            if '手续费' in line:
                match = re.search(r'手续费\s*[:：]?\s*([\d,]+\.\d+)', line)
                if match:
                    fee = match.group(1).replace(',', '')
                    break

        return product_name, fund_market_code, amount, shares, fee, '融联创同业交易平台'

    #民生同业e+
    def extract_minsheng_fields(lines):
        
        # 1. 产品名称（从客户名称字段提取）
        product_name = ''
        for line in lines:
            if '客户名称' in line:
                # 提取客户名称后的产品名称，支持冒号和中文冒号
                match = re.search(r'客户名称[：:]\s*(.*)', line)
                if match:
                    product_name = match.group(1).strip()
                    break
        
        # 清理产品名称中的多余空格和特殊字符
        product_name = product_name.replace(' ', '').replace('\u3000', '').replace('\n', '')

        # 2. 基金市场代码（从产品代码字段提取）
        fund_market_code = ''
        for line in lines:
            if '产品代码' in line:
                # 提取产品代码，支持6位数字代码
                match = re.search(r'产品代码[：:]\s*([0-9]{6})', line)
                if match:
                    fund_market_code = match.group(1)
                    break

        # 3. 确认金额（从确认金额（元）字段提取，去除逗号）
        amount = ''
        for line in lines:
            if '确认金额（元）' in line:
                # 提取确认金额，支持带逗号的格式，并去除逗号
                match = re.search(r'确认金额（元）[：:]\s*([\d,]+\.?\d*)', line)
                if match:
                    amount = match.group(1).replace(',', '')
                    break

        # 4. 确认份额（从确认份额（份）字段提取，去除逗号）
        shares = ''
        for line in lines:
            if '确认份额（份）' in line:
                # 提取确认份额，支持带逗号的格式，并去除逗号
                match = re.search(r'确认份额（份）[：:]\s*([\d,]+\.?\d*)', line)
                if match:
                    shares = match.group(1).replace(',', '')
                    break

        # 5. 手续费（从手续费（元）字段提取，去除逗号）
        fee = ''
        for line in lines:
            if '手续费（元）' in line:
                # 提取手续费，支持带逗号的格式，并去除逗号
                match = re.search(r'手续费（元）[：:]\s*([\d,]+\.?\d*)', line)
                if match:
                    fee = match.group(1).replace(',', '')
                    break

        return product_name, fund_market_code, amount, shares, fee, '民生同业e+'

    #和讯科技
    def extract_hexun_fields(lines):
       
        # 1. 产品名称（从账户名称字段提取，处理跨行情况）
        product_name = ''
        for i, line in enumerate(lines):
            if '账户名称' in line:
                # 先获取当前行账户名称后的内容
                parts = line.split('账户名称')
                if len(parts) > 1:
                    # 提取账户名称和账户类型之间的内容
                    current_part = parts[1].strip()
                    # 去除"账户类型"及其后面的内容
                    if '账户类型' in current_part:
                        current_part = current_part.split('账户类型')[0].strip()
                    product_name += current_part
                
                # 检查后续行是否包含产品名称的剩余部分
                j = i + 1
                while j < len(lines) and j < i + 5:  # 最多检查后续4行
                    next_line = lines[j].strip()
                    # 如果遇到其他字段标识，停止拼接
                    if (any(keyword in next_line for keyword in 
                        ['交易账号', '基金账号', '确认工作日', '业务类型', 
                        '确认单号', '基金代码', '基金名称']) or
                        len(next_line) == 0 or
                        next_line.isdigit()):  # 跳过纯数字行
                        break
                    # 拼接产品名称
                    product_name += next_line
                    j += 1
                break
        
        # 清理产品名称中的多余空格和特殊字符
        product_name = product_name.replace(' ', '').replace('\u3000', '').replace('\n', '')

        # 2. 基金市场代码（从基金代码字段提取）
        fund_market_code = ''
        for line in lines:
            if '基金代码' in line:
                # 提取6位数字的基金代码
                match = re.search(r'基金代码\s+([0-9]{6})', line)
                if match:
                    fund_market_code = match.group(1)
                    break

        # 3. 确认金额（从确认金额字段提取，去除逗号）
        amount = ''
        for line in lines:
            if '确认金额' in line:
                # 提取确认金额，支持带逗号的格式，并去除逗号
                match = re.search(r'确认金额\s+([\d,]+\.?\d*)', line)
                if match:
                    amount = match.group(1).replace(',', '')
                    break

        # 4. 确认份额（从确认份额字段提取，去除逗号）
        shares = ''
        for line in lines:
            if '确认份额' in line:
                # 提取确认份额，支持带逗号的格式，并去除逗号
                match = re.search(r'确认份额\s+([\d,]+\.?\d*)', line)
                if match:
                    shares = match.group(1).replace(',', '')
                    break

        # 5. 确认费用（从确认费用字段提取，去除逗号）
        fee = ''
        for line in lines:
            if '确认费用' in line:
                # 提取确认费用，支持带逗号的格式，并去除逗号
                match = re.search(r'确认费用\s+([\d,]+\.?\d*)', line)
                if match:
                    fee = match.group(1).replace(',', '')
                    break

        return product_name, fund_market_code, amount, shares, fee, '和讯基金'

    #招赢通基金
    def extract_zhaoyingtong_fields(lines):
        
        # 1. 产品名称（从投资者名称字段提取，可能跨行）
        product_name = ''
        for i, line in enumerate(lines):
            if '投资者名称' in line:
                # 先获取当前行投资者名称后的内容
                parts = line.split('投资者名称')
                if len(parts) > 1:
                    current_part = parts[1].strip()
                    product_name += current_part
                
                # 检查后续行是否包含产品名称的剩余部分
                j = i + 1
                while j < len(lines) and j < i + 5:  # 最多检查后续4行
                    next_line = lines[j].strip()
                    # 如果遇到其他字段标识，停止拼接
                    if (any(keyword in next_line for keyword in 
                        ['证件类型', '证件号码', '基金账号', '基金交易账号', 
                        '产品信息', '产品类型', '产品管理人', '产品代码']) or
                        len(next_line) == 0):
                        break
                    # 拼接产品名称
                    product_name += next_line
                    j += 1
                break
        
        # 清理产品名称中的多余空格和特殊字符
        product_name = product_name.replace(' ', '').replace('\u3000', '').replace('\n', '')

        # 2. 基金市场代码（从产品代码字段提取）
        fund_market_code = ''
        for line in lines:
            if '产品代码' in line:
                # 提取6位数字的基金代码
                match = re.search(r'产品代码\s+([0-9]{6})', line)
                if match:
                    fund_market_code = match.group(1)
                    break

        # 3. 确认金额（从确认金额字段提取，去除逗号和CNY）
        amount = ''
        for line in lines:
            if '确认金额' in line:
                # 提取确认金额，去除CNY和逗号
                match = re.search(r'确认金额\s+CNY\s+([\d,]+\.?\d*)', line)
                if match:
                    amount = match.group(1).replace(',', '')
                    break

        # 4. 确认份额（从确认份额字段提取，去除逗号）
        shares = ''
        for line in lines:
            if '确认份额' in line:
                # 提取确认份额，支持带逗号的格式，并去除逗号
                match = re.search(r'确认份额\s+([\d,]+\.?\d*)', line)
                if match:
                    shares = match.group(1).replace(',', '')
                    break

        # 5. 交易费用（从交易费用字段提取，去除逗号和CNY）
        fee = ''
        for line in lines:
            if '交易费用' in line:
                # 提取交易费用，去除CNY和逗号
                match = re.search(r'交易费用\s+CNY\s+([\d,]+\.?\d*)', line)
                if match:
                    fee = match.group(1).replace(',', '')
                    break

        return product_name, fund_market_code, amount, shares, fee, '招赢通'

    #兴证全球基金
    def extract_xingzheng_fields(lines):
        
        # 1. 产品名称（从账 号 名 称字段提取）
        product_name = ''
        for line in lines:
            if '账 号 名 称' in line:
                # 提取账号名称后的产品名称，支持冒号和中文冒号
                match = re.search(r'账\s*号\s*名\s*称\s*[：:]\s*(.*)', line)
                if match:
                    product_name = match.group(1).strip()
                    break
        
        # 清理产品名称中的多余空格和特殊字符
        product_name = product_name.replace(' ', '').replace('\u3000', '').replace('\n', '')

        # 2. 基金市场代码（从基 金 代 码字段提取）
        fund_market_code = ''
        for line in lines:
            if '基 金 代 码' in line:
                # 提取6位数字的基金代码
                match = re.search(r'基\s*金\s*代\s*码\s*[：:]\s*([0-9]{6})', line)
                if match:
                    fund_market_code = match.group(1)
                    break

        # 3. 确认金额（从确 认 金 额字段提取，去除逗号）
        amount = ''
        for line in lines:
            if '确 认 金 额' in line:
                # 提取确认金额，支持带逗号的格式，并去除逗号
                match = re.search(r'确\s*认\s*金\s*额\s*[：:]\s*([\d,]+\.?\d*)', line)
                if match:
                    amount = match.group(1).replace(',', '')
                    break

        # 4. 确认份额（从确 认 份 额字段提取，去除逗号）
        shares = ''
        for line in lines:
            if '确 认 份 额' in line:
                # 提取确认份额，支持带逗号的格式，并去除逗号
                match = re.search(r'确\s*认\s*份\s*额\s*[：:]\s*([\d,]+\.?\d*)', line)
                if match:
                    shares = match.group(1).replace(',', '')
                    break

        # 5. 手续费（从手 续 费字段提取，去除逗号）
        fee = ''
        for line in lines:
            if '手 续 费' in line:
                # 提取手续费，支持带逗号的格式，并去除逗号
                match = re.search(r'手\s*续\s*费\s*[：:]\s*([\d,]+\.?\d*)', line)
                if match:
                    fee = match.group(1).replace(',', '')
                    break

        return product_name, fund_market_code, amount, shares, fee, '兴证全球基金'

    #邮储银行
    def extract_youchu_fields(lines):
        
        # 1. 产品名称（从客户名称字段提取，处理跨行情况）
        product_name = ''
        for i, line in enumerate(lines):
            if '客户名称:' in line:
                # 先获取当前行客户名称后的内容
                parts = line.split('客户名称:')
                if len(parts) > 1:
                    current_part = parts[1].strip()
                    # 去除可能的其他字段信息，只保留客户名称部分
                    if '证件类型:' in current_part:
                        current_part = current_part.split('证件类型:')[0].strip()
                    product_name += current_part
                
                # 检查后续行是否包含产品名称的剩余部分
                j = i + 1
                while j < len(lines) and j < i + 5:  # 最多检查后续4行
                    next_line = lines[j].strip()
                    # 如果遇到其他字段标识，停止拼接
                    if (any(keyword in next_line for keyword in 
                        ['基金账号:', '证件号码:', '交易账号:', '产品信息', 
                        '基金公司:', '产品代码:', '产品名称:', '交易信息']) or
                        len(next_line) == 0):
                        break
                    # 拼接产品名称
                    product_name += next_line
                    j += 1
                break
        
        # 清理产品名称中的多余空格和特殊字符
        product_name = product_name.replace(' ', '').replace('\u3000', '').replace('\n', '')

        # 2. 基金市场代码（从产品代码字段提取）
        fund_market_code = ''
        for line in lines:
            if '产品代码:' in line:
                # 提取6位数字的基金代码
                match = re.search(r'产品代码:\s*([0-9]{6})', line)
                if match:
                    fund_market_code = match.group(1)
                    break

        # 3. 确认金额（从确认金额（元）字段提取）
        amount = ''
        for line in lines:
            if '确认金额（元）:' in line:
                # 提取确认金额数值
                match = re.search(r'确认金额（元）:\s*([\d,]+\.?\d*)', line)
                if match:
                    amount = match.group(1).replace(',', '')
                    break

        # 4. 确认份额（从确认份额（份）字段提取）
        shares = ''
        for line in lines:
            if '确认份额（份）:' in line:
                # 提取确认份额数值
                match = re.search(r'确认份额（份）:\s*([\d,]+\.?\d*)', line)
                if match:
                    shares = match.group(1).replace(',', '')
                    break

        # 5. 手续费（从手续费（元）字段提取）
        fee = ''
        for line in lines:
            if '手续费（元）:' in line:
                # 提取手续费数值
                match = re.search(r'手续费（元）:\s*([\d,]+\.?\d*)', line)
                if match:
                    fee = match.group(1).replace(',', '')
                    break

        return product_name, fund_market_code, amount, shares, fee, '邮储银行'

    #基煜基金
    def extract_jiyu_fields(lines):
        
        # 1. 产品名称（从账户名称字段提取）
        product_name = ''
        for line in lines:
            if '账户名称' in line:
                # 提取账户名称后的产品名称
                parts = line.split('账户名称')
                if len(parts) > 1:
                    product_name = parts[1].strip()
                    break
        
        # 清理产品名称中的多余空格和特殊字符
        product_name = product_name.replace(' ', '').replace('\u3000', '').replace('\n', '')

        # 2. 基金市场代码（从产品代码字段提取）
        fund_market_code = ''
        for line in lines:
            if '产品代码' in line:
                # 提取6位数字的基金代码
                match = re.search(r'产品代码\s*([0-9]{6})', line)
                if match:
                    fund_market_code = match.group(1)
                    break

        # 3. 确认金额（从确认金额字段提取，去除逗号和"元"）
        amount = ''
        for line in lines:
            if '确认金额' in line:
                # 提取确认金额，支持带逗号的格式，并去除逗号和"元"
                match = re.search(r'确认金额\s*([\d,]+\.?\d*)元?', line)
                if match:
                    amount = match.group(1).replace(',', '')
                    break

        # 4. 确认份额（从确认份额字段提取，去除逗号和"份"）
        shares = ''
        for line in lines:
            if '确认份额' in line:
                # 提取确认份额，支持带逗号的格式，并去除逗号和"份"
                match = re.search(r'确认份额\s*([\d,]+\.?\d*)份?', line)
                if match:
                    shares = match.group(1).replace(',', '')
                    break

        # 5. 手续费（从手续费字段提取，去除逗号和"元"）
        fee = ''
        for line in lines:
            if '手续费' in line:
                # 提取手续费，支持带逗号的格式，并去除逗号和"元"
                match = re.search(r'手续费\s*([\d,]+\.?\d*)元?', line)
                if match:
                    fee = match.group(1).replace(',', '')
                    break

        return product_name, fund_market_code, amount, shares, fee, '基煜基金'

    #宁波银行
    def extract_ningboBank_fields(lines):
        
        # 1. 产品名称（从客户名称字段提取）
        product_name = ''
        for line in lines:
            if '客户名称' in line:
                # 提取客户名称后的产品名称
                # 处理可能在同一行的情况
                parts = line.split('客户名称')
                if len(parts) > 1:
                    # 去除可能的其他字段（如基金账号等）
                    name_part = parts[1].strip()
                    # 如果同一行有其他字段标识，只取前面的部分
                    if '基金账号' in name_part:
                        name_part = name_part.split('基金账号')[0].strip()
                    product_name = name_part
                    break
        
        # 清理产品名称中的多余空格和特殊字符
        product_name = product_name.replace(' ', '').replace('\u3000', '').replace('\n', '')

        # 2. 基金市场代码（从产品代码字段提取）
        fund_market_code = ''
        for line in lines:
            if '产品代码' in line:
                # 提取6位数字的基金代码
                match = re.search(r'产品代码\s*([0-9]{6})', line)
                if match:
                    fund_market_code = match.group(1)
                    break

        # 3. 确认金额（从确认金额（元）字段提取，去除逗号）
        amount = ''
        for line in lines:
            if '确认金额（元）' in line:
                # 提取金额，支持带逗号的格式，并去除逗号
                match = re.search(r'确认金额（元）\s*([\d,]+\.?\d*)', line)
                if match:
                    amount = match.group(1).replace(',', '')
                    break

        # 4. 确认份额（从确认份额（份）字段提取，去除逗号）
        shares = ''
        for line in lines:
            if '确认份额（份）' in line:
                # 提取份额，支持带逗号的格式，并去除逗号
                match = re.search(r'确认份额（份）\s*([\d,]+\.?\d*)', line)
                if match:
                    shares = match.group(1).replace(',', '')
                    break

        # 5. 手续费（从申购费用（元）字段提取，去除逗号）
        fee = ''
        for line in lines:
            if '申购费用（元）' in line:
                # 提取费用，支持带逗号的格式，并去除逗号
                match = re.search(r'申购费用（元）\s*([\d,]+\.?\d*)', line)
                if match:
                    fee = match.group(1).replace(',', '')
                    break

        return product_name, fund_market_code, amount, shares, fee, '宁波银行'

    #国信嘉利基金
    def extract_guoxinjiali_fields(lines):
        
        # 1. 产品名称（从账户名称字段提取）
        product_name = ''
        for line in lines:
            if '账户名称' in line:
                # 提取账户名称后的产品名称
                parts = line.split('账户名称')
                if len(parts) > 1:
                    product_name = parts[1].strip()
                    break
        
        # 清理产品名称中的多余空格和特殊字符
        product_name = product_name.replace(' ', '').replace('\u3000', '').replace('\n', '')

        # 2. 基金市场代码（从产品代码字段提取）
        fund_market_code = ''
        for line in lines:
            if '产品代码' in line:
                # 提取6位数字的基金代码
                match = re.search(r'产品代码\s*([0-9]{6})', line)
                if match:
                    fund_market_code = match.group(1)
                    break

        # 3. 确认金额（从确认金额字段提取，去除逗号和"元"）
        amount = ''
        for line in lines:
            if '确认金额' in line:
                # 提取确认金额，支持带逗号的格式，并去除逗号和"元"
                match = re.search(r'确认金额\s*([\d,]+\.?\d*)元?', line)
                if match:
                    amount = match.group(1).replace(',', '')
                    break

        # 4. 确认份额（从确认份额字段提取，去除逗号和"份"）
        shares = ''
        for line in lines:
            if '确认份额' in line:
                # 提取确认份额，支持带逗号的格式，并去除逗号和"份"
                match = re.search(r'确认份额\s*([\d,]+\.?\d*)份?', line)
                if match:
                    shares = match.group(1).replace(',', '')
                    break

        # 5. 手续费（从手续费字段提取）
        fee = ''
        for line in lines:
            if '手续费' in line and '手续费（元）' not in line:  # 避免与其他平台的格式混淆
                # 提取手续费，支持带逗号的格式
                match = re.search(r'手续费\s*([\d,]+\.?\d*)', line)
                if match:
                    fee = match.group(1).replace(',', '')
                    break

        return product_name, fund_market_code, amount, shares, fee, '国信嘉利基金'

    #攀赢基金
    def extract_panying_fields(lines):
        
        # 1. 产品名称（从客户名称字段提取）
        product_name = ''
        for line in lines:
            if '客户名称' in line:
                # 提取客户名称后的产品名称
                parts = line.split('客户名称')
                if len(parts) > 1:
                    product_name = parts[1].strip()
                    break
        
        # 清理产品名称中的多余空格和特殊字符
        product_name = product_name.replace(' ', '').replace('\u3000', '').replace('\n', '')

        # 2. 基金市场代码（从产品代码字段提取）
        fund_market_code = ''
        for line in lines:
            if '产品代码' in line:
                # 提取6位数字的基金代码
                match = re.search(r'产品代码\s*([0-9]{6})', line)
                if match:
                    fund_market_code = match.group(1)
                    break

        # 3. 确认金额（从确认金额字段提取，去除逗号和"元"）
        amount = ''
        for line in lines:
            if '确认金额' in line:
                # 提取确认金额，支持带逗号的格式，并去除逗号和"元"
                # 这一行通常包含"申请金额"和"确认金额"，正则会匹配"确认金额"后面的数字
                match = re.search(r'确认金额\s*([\d,]+\.?\d*)元?', line)
                if match:
                    amount = match.group(1).replace(',', '')
                    break

        # 4. 确认份额（从确认份额字段提取，去除逗号和"份"）
        shares = ''
        for line in lines:
            if '确认份额' in line:
                # 提取确认份额，支持带逗号的格式，并去除逗号和"份"
                match = re.search(r'确认份额\s*([\d,]+\.?\d*)份?', line)
                if match:
                    shares = match.group(1).replace(',', '')
                    break

        # 5. 手续费（从手续费字段提取，去除逗号和"元"）
        fee = ''
        for line in lines:
            if '手续费' in line:
                # 提取手续费，支持带逗号的格式，并去除逗号和"元"
                match = re.search(r'手续费\s*([\d,]+\.?\d*)元?', line)
                if match:
                    fee = match.group(1).replace(',', '')
                    break

        return product_name, fund_market_code, amount, shares, fee, '攀赢基金'

    # 证达通基金（单笔申购确认）
    def extract_zdt_fields(lines):
        product_name = ''
        fund_market_code = ''
        amount = ''
        shares = ''
        fee = ''

        for line in lines:
            line = line.strip()

            # 1. 提取产品名称
            # 样本：投资者名称：万联资管广鑫66号集合资产管理计划 投资者类型：产品
            if '投资者名称' in line:
                # 使用非贪婪匹配，直到遇到“投资者类型”或行尾
                match = re.search(r'投资者名称[：:]\s*(.+?)(?:\s+投资者类型|$)', line)
                if match:
                    product_name = match.group(1).strip()

            # 2. 提取基金代码
            # 样本：基金名称：英大现金宝A 基金代码：000912
            if '基金代码' in line:
                match = re.search(r'基金代码[：:]\s*([0-9]{6})', line)
                if match:
                    fund_market_code = match.group(1)

            # 3. 提取确认金额
            # 样本：确认金额：100,000,000.00元
            if '确认金额' in line:
                match = re.search(r'确认金额[：:]\s*([\d,]+\.?\d*)', line)
                if match:
                    amount = match.group(1).replace(',', '')

            # 4. 提取确认份额
            # 样本：确认份额：100,000,000.00份
            if '确认份额' in line:
                match = re.search(r'确认份额[：:]\s*([\d,]+\.?\d*)', line)
                if match:
                    shares = match.group(1).replace(',', '')

            # 5. 提取手续费
            # 样本：基金净值：1.0000 手续费：0.00元
            if '手续费' in line:
                match = re.search(r'手续费[：:]\s*([\d,]+\.?\d*)', line)
                if match:
                    fee = match.group(1).replace(',', '')

        # 统一清理产品名称中的空格和特殊字符
        product_name = product_name.replace(' ', '').replace('\u3000', '').replace('\n', '')

        return product_name, fund_market_code, amount, shares, fee, '证达通基金'

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

        pdf_files = []
        for f in files:
            if not f.lower().endswith('.pdf'):
                continue
            # 排除“强行调”类文件
            if "强行调" in f:
                continue
            # 排除“调增”类文件
            if "调增" in f:
                continue
            # 排除“超级转换”类文件
            if "超级转换" in f:
                continue
            # 排除“转换”类文件
            if "转换" in f:
                continue
             # 排除“分红方式”类文件
            if "分红方式" in f:
                continue
            # 如果包含“赎回”，但文件名或文件内容表明它其实是申购确认单，就保留
            if "赎回" in f and not (("江苏银行" in f) or ("融联创" in f)):
                continue
            pdf_files.append(f)

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
                records = []  # 先初始化为空列表

                is_haomai = any('好买基金' in l for l in lines[:2]) and not any('转换' in l for l in lines)
                is_tiantian = ('天天基金' in file) or any('天天基金' in l for l in lines[3:]) and not any('转换' in l for l in lines)
                is_lide = any('利得基金' in l for l in lines[3:])
                is_changliang = any('长量基金' in l for l in lines[:2])
                is_yingmi = ('盈米' in file) or any('盈米' in l for l in lines[:3])
                is_jiaohang = ('交e通' in file) or any('交通银行' in l for l in lines[:2])
                is_jd = any('肯特瑞' in l for l in lines[:2]) and any('申购确认' in l for l in lines[:2])
                is_wangjin = ('网金' in file) or any('网金基金' in l for l in lines[5:]) 
                is_pingan = any('行E通' in l for l in lines[5:])
                is_jianhang = ('建行' in file) or any('客 户 名 称' in l for l in lines) #比较脆弱，考虑"红股"
                is_liantai = (('北极星' in file) or any('联泰' in l for l in lines[:2])) and any('申购' in l for l in lines[:20])
                is_tengyuan = ('腾元' in file) or any('腾元基金' in l for l in lines[5:])
                is_ronglianchuang = (('江苏银行' in file) or any('融联创' in l for l in lines[:2])) and any('申购' in l for l in lines[:5])
                is_minsheng = ('民生同业e+' in file) or any('同业e+' in l for l in lines[2:])
                is_hexun = ('和讯' in file) or any('和讯信息科技有限公司' in l for l in lines[3:])
                is_zhaoyingtong = ('招赢通' in file) or any('招赢通' in l for l in lines[:2])
                is_xingzheng = ('兴证' in file) or any('兴证全球基金' in l for l in lines[:2])
                is_youchu = ('邮储' in file)
                is_jiyu = any('基煜基金' in l for l in lines[:2])
                is_ningboBank = (('宁波' in file and '北极星' not in file) or (any('宁波银行' in l for l in lines[15:]) and not any('联泰' in l for l in lines[:5])))
                is_guoxinjiali = any('国信嘉利基金' in l for l in lines[:2])
                is_panying = ('攀赢' in file) or any('攀赢' in l for l in lines[:2])
                is_zdt = (any('证达通' in l for l in lines) and any('申购确认单' in l for l in lines))

                if is_haomai:
                    product_name, fund_market_code, amount, shares, fee, platform = extract_haomai_fields(text, lines)
                    records = [(product_name, fund_market_code, amount, shares, fee, platform)]
                elif is_tiantian:
                    product_name, fund_market_code, amount, shares, fee, platform = extract_tiantian_fields(lines)
                    records = [(product_name, fund_market_code, amount, shares, fee, platform)]
                elif is_lide:
                    product_name, fund_market_code, amount, shares, fee, platform = extract_lide_fields(lines)
                    records = [(product_name, fund_market_code, amount, shares, fee, platform)]
                elif is_changliang:
                    product_name, fund_market_code, amount, shares, fee, platform = extract_changliang_fields(lines)
                    records = [(product_name, fund_market_code, amount, shares, fee, platform)]
                elif is_jiaohang:
                    product_name, fund_market_code, amount, shares, fee, platform = extract_jiaohang_fields(lines)
                    records = [(product_name, fund_market_code, amount, shares, fee, platform)]
                elif is_jd:
                    product_name, fund_market_code, amount, shares, fee, platform = extract_jd_fields(lines)
                    records = [(product_name, fund_market_code, amount, shares, fee, platform)]
                elif is_wangjin:
                    product_name, fund_market_code, amount, shares, fee, platform = extract_wangjin_fields(lines)
                    records = [(product_name, fund_market_code, amount, shares, fee, platform)]
                elif is_pingan:
                    product_name, fund_market_code, amount, shares, fee, platform = extract_pingan_fields(lines)
                    records = [(product_name, fund_market_code, amount, shares, fee, platform)]
                elif is_jianhang:
                    product_name, fund_market_code, amount, shares, fee, platform = extract_jianhang_fields(lines)
                    records = [(product_name, fund_market_code, amount, shares, fee, platform)]
                elif is_tengyuan:
                    product_name, fund_market_code, amount, shares, fee, platform = extract_tengyuan_fields(lines)
                    records = [(product_name, fund_market_code, amount, shares, fee, platform)]
                elif is_ronglianchuang:
                    product_name, fund_market_code, amount, shares, fee, platform = extract_ronglianchuang_fields(lines)
                    records = [(product_name, fund_market_code, amount, shares, fee, platform)]
                elif is_minsheng:
                    product_name, fund_market_code, amount, shares, fee, platform = extract_minsheng_fields(lines)
                    records = [(product_name, fund_market_code, amount, shares, fee, platform)]
                elif is_hexun:
                    product_name, fund_market_code, amount, shares, fee, platform = extract_hexun_fields(lines)
                    records = [(product_name, fund_market_code, amount, shares, fee, platform)]
                elif is_zhaoyingtong:
                    product_name, fund_market_code, amount, shares, fee, platform = extract_zhaoyingtong_fields(lines)
                    records = [(product_name, fund_market_code, amount, shares, fee, platform)]
                elif is_xingzheng:
                    product_name, fund_market_code, amount, shares, fee, platform = extract_xingzheng_fields(lines)
                    records = [(product_name, fund_market_code, amount, shares, fee, platform)]
                elif is_youchu:
                    product_name, fund_market_code, amount, shares, fee, platform = extract_youchu_fields(lines)
                    records = [(product_name, fund_market_code, amount, shares, fee, platform)]
                elif is_jiyu:
                    product_name, fund_market_code, amount, shares, fee, platform = extract_jiyu_fields(lines)
                    records = [(product_name, fund_market_code, amount, shares, fee, platform)]
                elif is_ningboBank:
                    product_name, fund_market_code, amount, shares, fee, platform = extract_ningboBank_fields(lines)
                    records = [(product_name, fund_market_code, amount, shares, fee, platform)]
                elif is_guoxinjiali:
                    product_name, fund_market_code, amount, shares, fee, platform = extract_guoxinjiali_fields(lines)
                    records = [(product_name, fund_market_code, amount, shares, fee, platform)]
                elif is_panying:
                    product_name, fund_market_code, amount, shares, fee, platform = extract_panying_fields(lines)
                    records = [(product_name, fund_market_code, amount, shares, fee, platform)]
                elif is_zdt:
                    product_name, fund_market_code, amount, shares, fee, platform = extract_zdt_fields(lines)
                    records = [(product_name, fund_market_code, amount, shares, fee, platform)]
                elif is_yingmi:
                    records = extract_yingmi_fields(lines)
                elif is_liantai:
                    records = extract_liantai_fields(lines)

                # 只有当 records 不为空时才处理数据
                if not records:
                    continue

                for product_name, fund_market_code, amount, shares, fee, platform in records:
                    temp_df = pd.DataFrame([{
                        '产品名称': product_name,
                        '基金市场代码': fund_market_code, 
                        '金额': amount,
                        '数量': shares,
                        '手续费': fee,
                        '基金平台': platform
                    }])

                    #强制转换为数值型
                    temp_df['金额'] = pd.to_numeric(temp_df['金额'], errors='coerce').round(2)
                    temp_df['手续费'] = pd.to_numeric(temp_df['手续费'], errors='coerce').round(2)
                    temp_df['数量'] = pd.to_numeric(temp_df['数量'], errors='coerce').round(2)
                    # 映射账套编号
                    temp_df['账套编号'] = temp_df['产品名称'].map(product_code_dict)
                    # 填充其它字段
                    temp_df['交易市场'] = '国内银行间'
                    temp_df['业务类别'] = '基金申购确认'
                    temp_df['日期'] = today_str
                    temp_df['佣金'] = ''
                    temp_df['交易对手'] = ''
                    temp_df['资金账户'] = ''
                    temp_df['赎回到账日期'] = ''
                    temp_df['股东账户'] = ''
                    temp_df['席位号'] = ''
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
    output_file = os.path.join(output_folder, "【境内基金业务】申购确认.xls")
    try:
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            target_df.to_excel(writer, sheet_name='Sheet1', index=False)
        log(f"已汇总输出到: {output_file}", log_text)
        return output_folder
    except Exception as e:
        log(f"写入Excel失败: {e}", log_text)
        return False