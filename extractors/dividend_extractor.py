import os
import json
import pdfplumber
import re
import pandas as pd
from datetime import datetime, timedelta
from utils.common import log

def run_dividend_extract(folder_path, json_path, log_text):
    """运行分红单提取
    
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
    target_cols = ['账套编号','产品代码', '基金市场代码','交易市场','日期', '派送份额', '派送金额', '红利截止日期', '持仓分类','产品名称','基金平台']

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

        m2 = re.search(r'产品代码[：: ]*([0-9]{6})', text)
        fund_market_code = m2.group(1).strip() if m2 else ''

        m3 = re.search(r'确认金额[：: ]*([\d,]+\.\d+)', text)
        dividend_amount = m3.group(1).replace(',', '') if m3 else ''

        m4 = re.search(r'确认份额[：: ]*([\d,]+\.\d+)', text)
        dividend_shares = m4.group(1).replace(',', '') if m4 else ''

        return product_name, fund_market_code, dividend_amount, dividend_shares, '好买基金'

    #天天
    def extract_tiantian_fields(lines):
        product_name = ''
        for i, line in enumerate(lines):
            if '账户户名' in line:
                prev_line = lines[i-1].strip() if i > 0 else ''
                this_name = line.split('账户户名')[0].strip()
                next_line = lines[i+1].strip() if i+1 < len(lines) else ''
                cand = ''
                if prev_line and '确认单' not in prev_line:
                    cand += prev_line
                if this_name:
                    cand += this_name
                if ('产管理计划' in next_line) or (next_line and '账户类型' not in next_line):
                    cand += next_line
                product_name = cand.replace(' ', '').replace('\u3000', '')
                break

        fund_market_code = ''
        for line in lines:
            if '基金代码' in line:
                match = re.search(r'基金代码\s*([0-9]{6})', line)
                if match:
                    fund_market_code = match.group(1)
                    break

        dividend_amount = ''
        for i, line in enumerate(lines):
            if '红利资金（元' in line:
                next_line = lines[i+1].strip() if i+1 < len(lines) else ''
                nums = re.findall(r'[\d,]+\.\d+', next_line)
                if nums:
                    dividend_amount = nums[-1].replace(',', '')
                break

        dividend_shares = ''
        for i, line in enumerate(lines):
            if '红利再投资基' in line:
                for offset in range(1, 3):
                    idx = i + offset
                    if idx < len(lines):
                        nums = re.findall(r'[\d,]+\.\d+', lines[idx])
                        if nums:
                            dividend_shares = nums[-1].replace(',', '')
                            break
                if dividend_shares:
                    break

        return product_name, fund_market_code, dividend_amount, dividend_shares, '天天基金'

    #利得
    def extract_lide_fields(lines):
        product_name = ''
        for line in lines:
            if '投资者姓名/名称' in line:
                match = re.search(r'投资者姓名/名称[:：]\s*(.*)', line)
                if match:
                    product_name = match.group(1).replace(' ', '').replace('\u3000', '')
                    break

        fund_market_code = ''
        for line in lines:
            if '基金代码' in line:
                match = re.search(r'基金代码\s*([0-9]{6})', line)
                if match:
                    fund_market_code = match.group(1)
                    break

        dividend_amount = ''
        for line in lines:
            if '红利总金额' in line:
                match = re.search(r'红利总金额（元）\s*([\d,]+\.\d+)', line)
                if match:
                    dividend_amount = match.group(1).replace(',', '')
                    break

        dividend_shares = ''
        for line in lines:
            if '红利再投份额' in line:
                match = re.search(r'红利再投份额（份）\s*([\d,]+\.\d+)', line)
                if match:
                    dividend_shares = match.group(1).replace(',', '')
                    break

        return product_name, fund_market_code, dividend_amount, dividend_shares, '利得基金'

    #长量
    def extract_changliang_fields(lines):
        product_name = ''
        for line in lines:
            if '投资者名称' in line:
                match = re.search(r'投资者名称\s*(.*)', line)
                if match:
                    product_name = match.group(1).strip().replace(' ', '').replace('\u3000', '')
                    break

        fund_market_code = ''
        for line in lines:
            if '基金代码' in line:
                match = re.search(r'基金代码\s*([0-9]{6})', line)
                if match:
                    fund_market_code = match.group(1)
                    break

        dividend_amount = ''
        for line in lines:
            if '红利转投份额' in line:
                match = re.search(r'红利转投份额\s*([\d,]+\.\d+)', line)
                if match:
                    dividend_amount = match.group(1).replace(',', '')
                    break

        dividend_shares = dividend_amount

        return product_name, fund_market_code, dividend_amount, dividend_shares, '长量基金'

    #兴证全球
    def extract_xingzheng_fields(lines):
        product_name = ''
        for line in lines:
            if '账 号 名 称' in line:
                match = re.search(r'账\s*号\s*名\s*称\s*[:：]\s*(.*)', line)
                if match:
                    product_name = match.group(1).replace(' ', '').replace('\u3000', '')
                    break

        fund_market_code = ''
        for line in lines:
            if '基 金 代 码' in line:
                match = re.search(r'基\s*金\s*代\s*码\s*[:：]\s*([0-9]{6})', line)
                if match:
                    fund_market_code = match.group(1)
                    break

        dividend_amount = ''
        for line in lines:
            if '再投资份额' in line:
                match = re.search(r'再投资份额\s*[:：]?\s*([\d,]+\.\d+)', line)
                if match:
                    dividend_amount = match.group(1).replace(',', '')
                    break

        dividend_shares = dividend_amount

        return product_name, fund_market_code, dividend_amount, dividend_shares,'兴证全球基金'

    #盈米
    def extract_yingmi_fields(lines):
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

        results = []
        i = 0
        N = len(lines)
        while i < N:
            line = lines[i]
            if '序号:' in line and '基金代码:' in line:
                fund_market_code = ''
                match = re.search(r'基金代码[:：]\s*([0-9A-Za-z]+)', line)
                if match:
                    fund_market_code = match.group(1)
                dividend_amount = ''
                dividend_shares = ''
                lookahead = 1
                while (i + lookahead) < N and lookahead <= 4:
                    subline = lines[i + lookahead]
                    if '分红金额' in subline:
                        match = re.search(r'分红金额[:：]?\s*([\d,\.]+)', subline)
                        if match:
                            dividend_amount = match.group(1).replace(',', '')
                    if '红利再投份额' in subline:
                        match = re.search(r'红利再投份额[:：]?\s*([\d,\.]+)', subline)
                        if match:
                            dividend_shares = match.group(1).replace(',', '')
                    lookahead += 1
                if fund_market_code and dividend_amount and dividend_shares:
                    results.append((product_name, fund_market_code, dividend_amount, dividend_shares, '盈米基金'))
            i += 1
        return results

    #招赢通基金
    def extract_zhaoyingtong_fields(lines):
        # 1. 产品名称（投资者名称）
        product_name = ''
        for line in lines:
            if '投资者名称' in line:
                match = re.search(r'投资者名称\s*(.*)', line)
                if match:
                    product_name = match.group(1).strip().replace(' ', '').replace('\u3000', '')
                    break

        # 2. 基金市场代码（产品代码，文本格式）
        fund_market_code = ''
        for line in lines:
            if '产品代码' in line:
                match = re.search(r'产品代码\s*([0-9A-Za-z]+)', line)
                if match:
                    fund_market_code = match.group(1).strip()
                    break

        # 3. 派送金额（分红金额，去除CNY）
        dividend_amount = ''
        for line in lines:
            if '分红金额' in line and 'CNY' in line:
                match = re.search(r'CNY\s*([\d,\.]+)', line)
                if match:
                    dividend_amount = match.group(1).replace(',', '')
                    break

        # 4. 派送份额（转投份额(份)）
        dividend_shares = ''
        for line in lines:
            if '转投份额' in line:
                match = re.search(r'转投份额\(份\)\s*([\d,\.]+)', line)
                if match:
                    dividend_shares = match.group(1).replace(',', '')
                    break

        return product_name, fund_market_code, dividend_amount, dividend_shares,'招赢通基金'

    #邮储银行
    def extract_youchu_fields(lines):
        # 1. 产品名称（客户名称，可能跨两三行拼接）
        product_name = ''
        for i, line in enumerate(lines):
            if '客户名称' in line:
                # 取本行"客户名称:"后的内容
                match = re.search(r'客户名称[:：]?\s*(\S+)', line)
                if match:
                    name_part = match.group(1)
                else:
                    name_part = ''
                # 有些产品名被拆分到下2~3行
                ext = ''
                for j in range(1, 4):
                    idx = i + j
                    if idx < len(lines):
                        ext_line = lines[idx].strip()
                        # 只提取含"集合资产管"或"理计划"等关键字的行
                        if ext_line and (('集合资产管' in ext_line) or ('理计划' in ext_line)):
                            ext += ext_line
                product_name = (name_part + ext).replace(' ', '').replace('\u3000', '')
                break

        # 2. 基金市场代码（产品代码）
        fund_market_code = ''
        for line in lines:
            if '产品代码' in line:
                match = re.search(r'产品代码[:：]?\s*([0-9A-Za-z]+)', line)
                if match:
                    fund_market_code = match.group(1).strip()
                    break

        # 3. 派送金额（再投资金额，去除"元"）
        dividend_amount = ''
        for line in lines:
            if '再投资金额' in line:
                match = re.search(r'再投资金额[:：]?\s*([\d,\.]+)', line)
                if match:
                    dividend_amount = match.group(1).replace(',', '')
                    break

        # 4. 派送份额（红股，去除"份"）
        dividend_shares = ''
        for line in lines:
            if '红股' in line:
                match = re.search(r'红股[:：]?\s*([\d,\.]+)', line)
                if match:
                    dividend_shares = match.group(1).replace(',', '')
                    break

        return product_name, fund_market_code, dividend_amount, dividend_shares,'邮储银行'

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

        # 3. 红利金额和红利份额（都从确认份额(份)字段提取）
        dividend_amount = ''
        dividend_shares = ''
        for line in lines:
            if '确认份额(份)' in line:
                # 提取确认份额数值，支持带逗号的格式
                match = re.search(r'确认份额\(份\)\s*([\d,]+\.?\d*)', line)
                if match:
                    value = match.group(1).replace(',', '')
                    dividend_amount = value  # 红利金额
                    dividend_shares = value  # 红利份额
                    break

        return product_name, fund_market_code, dividend_amount, dividend_shares,'平安行E通'

    #交通银行交e通
    def extract_jiaohang_fields(lines):
        
        # 1. 产品名称（投资者信息）
        product_name = ''
        for line in lines:
            if '投资者信息' in line:
                # 提取"投资者信息"后面的内容
                match = re.search(r'投资者信息\s*(.*)', line)
                if match:
                    product_name = match.group(1).strip().replace(' ', '').replace('\u3000', '')
                    break

        # 2. 基金市场代码（产品代码）
        fund_market_code = ''
        for line in lines:
            if '产品代码' in line:
                match = re.search(r'产品代码\s*([0-9A-Za-z]+)', line)
                if match:
                    fund_market_code = match.group(1).strip()
                    break

        # 3. 派送金额和派送份额（转投份额）
        dividend_amount = ''
        dividend_shares = ''
        for line in lines:
            if '转投份额' in line:
                match = re.search(r'转投份额\s*([\d,\.]+)', line)
                if match:
                    amount = match.group(1).replace(',', '')
                    dividend_amount = amount
                    dividend_shares = amount
                    break

        # 4. 红利截止日期（确认日期的前一日）
        dividend_end_date = ''
        for line in lines:
            if '确认日期' in line:
                match = re.search(r'确认日期\s*(\d{8})', line)
                if match:
                    confirm_date_str = match.group(1)
                    try:
                        # 将字符串转换为日期对象
                        confirm_date = datetime.strptime(confirm_date_str, '%Y%m%d')
                        # 计算前一日
                        dividend_end_date = (confirm_date - timedelta(days=1)).strftime('%Y%m%d')
                    except ValueError:
                        dividend_end_date = ''
                    break

        return product_name, fund_market_code, dividend_amount, dividend_shares,'交e通', dividend_end_date

    #和讯科技
    def extract_hexun_fields(lines):
    
        # 1. 产品名称（从账户名称字段提取，需要跨行拼接）
        product_name = ''
        for i, line in enumerate(lines):
            if '账户名称' in line:
                # 提取账户名称后的内容
                parts = line.split('账户名称')
                if len(parts) > 1:
                    # 去除"账户类型"及其后的内容
                    name_part = parts[1].split('账户类型')[0].strip()
                    product_name += name_part
                
                # 检查后续几行是否包含产品名称的剩余部分
                j = i + 1
                while j < len(lines) and j < i + 5:  # 最多检查后续4行
                    next_line = lines[j].strip()
                    # 如果行包含关键字段，停止拼接
                    if (any(keyword in next_line for keyword in 
                        ['交易账号', '确认工作日', '基金代码', '红利基数', '重要提示']) or
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

        # 3. 红利金额（从红利资金(元)字段提取）
        dividend_amount = ''
        for line in lines:
            if '红利资金(元)' in line:
                # 提取红利金额
                match = re.search(r'红利资金\(元\)\s*([\d,]+\.?\d*)', line)
                if match:
                    dividend_amount = match.group(1).replace(',', '')
                    break

        # 4. 红利份额（从红利再投资确认份额字段提取，可能跨行）
        dividend_shares = ''
        for i, line in enumerate(lines):
            if '红利再投资确认份' in line:
                # 先在当前行查找数值
                match = re.search(r'红利再投资确认份[额]?\s*([\d,]+\.?\d*)', line)
                if match:
                    dividend_shares = match.group(1).replace(',', '')
                    break
                # 如果当前行没有找到，检查下一行
                elif i + 1 < len(lines):
                    next_line = lines[i + 1].strip()
                    # 在下一行查找数值（通常是"额"字段拆分的情况）
                    match = re.search(r'^([\d,]+\.?\d*)', next_line)
                    if match:
                        dividend_shares = match.group(1).replace(',', '')
                        break

        return product_name, fund_market_code, dividend_amount, dividend_shares, '和讯科技'

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

        # 3. 红利金额和红利份额（都从红股字段提取）
        dividend_amount = ''
        dividend_shares = ''
        for line in lines:
            if '红 股' in line:
                # 提取红股数值
                match = re.search(r'红\s*股\s*[：:]?\s*([\d,]+\.?\d*)', line)
                if match:
                    value = match.group(1).replace(',', '')
                    dividend_amount = value  # 红利金额
                    dividend_shares = value  # 红利份额
                    break

        return product_name, fund_market_code, dividend_amount, dividend_shares, '建行直销'

    #腾元基金
    def extract_tengyuan_fields(lines):
     
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

        # 3. 红利金额和红利份额（都从红利再投份额字段提取）
        dividend_amount = ''
        dividend_shares = ''
        for line in lines:
            if '红利再投份额' in line:
                # 提取红利再投份额数值，去除逗号
                match = re.search(r'红利再投份额\s*([\d,]+\.?\d*)', line)
                if match:
                    value = match.group(1).replace(',', '')
                    dividend_amount = value  # 红利金额
                    dividend_shares = value  # 红利份额
                    break

        return product_name, fund_market_code, dividend_amount, dividend_shares, '腾元基金'

    #网金基金
    def extract_wangjin_fields(lines):
        # 判断是否为第二种格式（包含分隔线）
        is_format2 = any('─────' in line for line in lines)
        
        # 1. 产品名称（从客户名称字段提取）
        product_name = ''
        
        if is_format2:
            # 第二种格式：客户名称可能在第1行
            for line in lines[:3]:  # 只检查前3行
                if '客户名称' in line:
                    # 提取客户名称和网点名称之间的内容
                    match = re.search(r'客户名称\s*([^网点名称]+)', line)
                    if match:
                        product_name = match.group(1).strip()
                        # 检查是否需要拼接下一行（如果产品名称被截断）
                        if not product_name.endswith('计划'):
                            # 查找下一行是否包含"理计划"等结尾
                            line_idx = lines.index(line)
                            if line_idx + 1 < len(lines):
                                next_line = lines[line_idx + 1].strip()
                                if '理计划' in next_line or '管理计划' in next_line:
                                    product_name += next_line
                        break
        else:
            # 第一种格式：原有逻辑
            for i, line in enumerate(lines):
                if '客户名称' in line:
                    parts = line.split('客户名称')
                    if len(parts) > 1:
                        current_part = parts[1].strip()
                        product_name += current_part
                    
                    # 检查后续几行是否包含产品名称的剩余部分
                    j = i + 1
                    while j < len(lines) and j < i + 5:
                        next_line = lines[j].strip()
                        if ('理计划' in next_line or '管理计划' in next_line):
                            product_name += next_line
                            break
                        elif (next_line and 
                            '基金账号' not in next_line and 
                            '交易账号' not in next_line and
                            '交易类别' not in next_line and
                            '基金代码' not in next_line and
                            '─────' not in next_line and
                            len(next_line) > 3):
                            product_name += next_line
                        else:
                            break
                        j += 1
                    break
        
        # 清理产品名称
        product_name = product_name.replace(' ', '').replace('\u3000', '').replace('\n', '')
        
        # 2. 基金市场代码（从基金代码字段提取）
        fund_market_code = ''
        for line in lines:
            if '基金代码' in line:
                match = re.search(r'基金代码\s*([0-9]{6})', line)
                if match:
                    fund_market_code = match.group(1)
                    break
        
        # 3. 红利金额和红利份额（从红利再投份额字段提取）
        dividend_amount = ''
        dividend_shares = ''
        
        if is_format2:
            # 第二种格式：数值可能包含逗号，且在同一行
            for line in lines:
                if '红利再投份额' in line:
                    # 提取红利再投份额后的数值（可能包含逗号）
                    match = re.search(r'红利再投份额\s*([\d,]+\.?\d*)', line)
                    if match:
                        value = match.group(1).replace(',', '')
                        dividend_amount = value
                        dividend_shares = value
                        break
        else:
            # 第一种格式：原有逻辑（可能跨行）
            for i, line in enumerate(lines):
                if '红利再投份额' in line:
                    # 先在当前行查找数值
                    match = re.search(r'红利再投份额\s*([\d,]+\.?\d*)', line)
                    if match:
                        value = match.group(1).replace(',', '')
                        dividend_amount = value
                        dividend_shares = value
                        break
                    # 如果当前行没有找到，检查下一行
                    elif i + 1 < len(lines):
                        next_line = lines[i + 1].strip()
                        match = re.search(r'^([\d,]+\.?\d*)', next_line)
                        if match:
                            value = match.group(1).replace(',', '')
                            dividend_amount = value
                            dividend_shares = value
                            break
        
        return product_name, fund_market_code, dividend_amount, dividend_shares, '网金基金'

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

        # 3. 红利金额（从红利再投金额字段提取）
        dividend_amount = ''
        for line in lines:
            if '红利再投金额' in line:
                # 提取红利再投金额数值，去除逗号
                match = re.search(r'红利再投金额\s*([\d,]+\.?\d*)', line)
                if match:
                    dividend_amount = match.group(1).replace(',', '')
                    break

        # 4. 红利份额（从红利再投份额字段提取）
        dividend_shares = ''
        for line in lines:
            if '红利再投份额' in line:
                # 提取红利再投份额数值，去除逗号
                match = re.search(r'红利再投份额\s*([\d,]+\.?\d*)', line)
                if match:
                    dividend_shares = match.group(1).replace(',', '')
                    break

        return product_name, fund_market_code, dividend_amount, dividend_shares, '京东肯特瑞'

    #融联创同业交易平台
    def extract_ronglianchuang_fields(lines):
        
        # 1. 产品名称（从投资主体产品名称字段提取，处理跨行情况）
        product_name = ''
        for i, line in enumerate(lines):
            if '投资主体产品名称' in line:
                # 提取投资主体产品名称后的内容，但需要排除银行账号等字段
                parts = line.split('投资主体产品名称')
                if len(parts) > 1:
                    current_part = parts[1].strip()
                    # 如果当前行包含银行账号，只取银行账号之前的部分
                    if '银行账号' in current_part:
                        current_part = current_part.split('银行账号')[0].strip()
                    product_name += current_part
                
                # 检查后续行是否包含产品名称的剩余部分（如"管理计划"）
                j = i + 1
                while j < len(lines) and j < i + 3:  # 最多检查后续2行
                    next_line = lines[j].strip()
                    # 如果遇到其他字段标识，停止拼接
                    if (any(keyword in next_line for keyword in 
                        ['银行账号', '基金账号', '平台交易账号', '产品信息', '基金代码', '基金名称']) or
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
                # 提取基金代码（支持6位数字）
                match = re.search(r'基金代码\s*([0-9]{6})', line)
                if match:
                    fund_market_code = match.group(1)
                    break

        # 3. 红利金额和红利份额（都从再投资份额字段提取）
        dividend_amount = ''
        dividend_shares = ''
        for line in lines:
            if '再投资份额（份）' in line:
                # 提取再投资份额数值，去除逗号
                match = re.search(r'再投资份额（份）\s*([\d,]+\.?\d*)', line)
                if match:
                    value = match.group(1).replace(',', '')
                    dividend_amount = value  # 红利金额
                    dividend_shares = value  # 红利份额
                    break

        return product_name, fund_market_code, dividend_amount, dividend_shares, '融联创同业交易平台'

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
                dividend_amount = ''
                dividend_shares = ''
                
                # 在接下来的几行中查找基金代码和红利再投份额
                lookahead = 1
                while (i + lookahead) < N and lookahead <= 8:  # 增加查找范围
                    subline = lines[i + lookahead]
                    
                    # 查找基金代码
                    if '基金代码' in subline:
                        # 提取基金代码（支持6位数字）
                        match = re.search(r'基金代码\s+([0-9]{6})', subline)
                        if match:
                            fund_market_code = match.group(1)
                    
                    # 查找红利再投份额
                    if '红利再投份额(份)' in subline:
                        match = re.search(r'红利再投份额\(份\)\s*([\d,]+\.?\d*)', subline)
                        if match:
                            dividend_shares = match.group(1).replace(',', '')
                            dividend_amount = dividend_shares  # 联泰基金的分红金额等于份额
                    
                    # 如果遇到下一个交易信息块，停止查找
                    if lookahead > 1 and '交易信息' in subline:
                        break
                        
                    lookahead += 1
                
                # 如果找到了基金代码和红利份额，添加到结果中
                if fund_market_code and dividend_amount and dividend_shares:
                    results.append((product_name, fund_market_code, dividend_amount, dividend_shares, '联泰基金'))
            
            i += 1
        
        return results

    #民生同业e+
    def extract_minsheng_fields(lines):
       
        # 1. 产品名称（从客户名称字段提取）
        product_name = ''
        for line in lines:
            if '客户名称' in line:
                # 提取客户名称后的产品名称
                match = re.search(r'客户名称[：:]\s*(.*)', line)
                if match:
                    product_name = match.group(1).strip().replace(' ', '').replace('\u3000', '')
                    break
        
        # 2. 基金市场代码（从产品代码字段提取）
        fund_market_code = ''
        for line in lines:
            if '产品代码' in line:
                # 提取产品代码（支持6位数字或字母数字组合）
                match = re.search(r'产品代码[：:]\s*([0-9A-Za-z]+)', line)
                if match:
                    fund_market_code = match.group(1).strip()
                    break
        
        # 3. 派送金额和派送份额（都从确认份额字段提取）
        dividend_amount = ''
        dividend_shares = ''
        for line in lines:
            if '确认份额（份）' in line:
                # 提取确认份额数值，去除逗号
                match = re.search(r'确认份额（份）[：:]?\s*([\d,]+\.?\d*)', line)
                if match:
                    value = match.group(1).replace(',', '')
                    dividend_amount = value  # 派送金额
                    dividend_shares = value  # 派送份额
                    break
        
        return product_name, fund_market_code, dividend_amount, dividend_shares, '民生同业e+'

    #证达通基金
    def extract_zdt_fields(lines):
        """
        提取证达通基金分红确认单的字段
        兼容：
        1. 汇总列表格式 - 标准行（包含基金名称）
        2. 汇总列表格式 - 紧凑行（因换行导致基金名称缺失，账号直连代码）
        3. 单笔确认单格式（保留原逻辑）
        """
        # 1. 公共部分：提取产品名称（从投资者名称字段提取）
        product_name = ''
        for line in lines:
            if '投资者名称' in line:
                match = re.search(r'投资者名称[：:]\s*(.+?)(?:\s+生成时间|$)', line)
                if match:
                    product_name = match.group(1).strip().replace(' ', '').replace('\u3000', '')
                else:
                    match = re.search(r'投资者名称[：:]\s*(\S+)', line)
                    if match:
                        product_name = match.group(1).strip().replace(' ', '').replace('\u3000', '')
                break
        
        results = []
        is_tabular = False
        
        # 2. 遍历行进行汇总列表正则匹配
        for line in lines:
            line = line.strip()
            
            # --- 方案 A：优先尝试匹配【换行导致的紧凑格式】 ---
            # 特征：账号(\d{10,}) 后面紧接着就是 代码([0-9]{6})，中间没有基金名
            # 解决您遇到的：'3 0000000011738 018655 红利再投资...'
            match_compact = re.match(
                r'^\s*(\d+)\s+'           # 序号
                r'(\d{10,})\s+'           # 交易账号
                r'([0-9]{6})\s+'          # 基金代码 (直接接代码)
                r'红利再投资\s+'          # 分红方式
                r'([\d,]+\.?\d*)\s+'      # 分红金额
                r'([\d,]+\.?\d*)',        # 分红份额
                line
            )
            
            # --- 方案 B：匹配【标准汇总格式】 ---
            # 特征：账号和代码中间有文字（基金名称）
            match_standard = re.match(
                r'^\s*(\d+)\s+'           # 序号
                r'(\d{10,})\s+'           # 交易账号
                r'(.+?)\s+'               # 基金名称 (非贪婪匹配)
                r'([0-9]{6})\s+'          # 基金代码
                r'红利再投资\s+'          # 分红方式
                r'([\d,]+\.?\d*)\s+'      # 分红金额
                r'([\d,]+\.?\d*)',        # 分红份额
                line
            )
            
            if match_compact:
                is_tabular = True
                fund_market_code = match_compact.group(3)
                dividend_amount = match_compact.group(4).replace(',', '')
                dividend_shares = match_compact.group(5).replace(',', '')
                results.append((product_name, fund_market_code, dividend_amount, dividend_shares, '证达通基金'))
            
            elif match_standard:
                is_tabular = True
                fund_market_code = match_standard.group(4)
                dividend_amount = match_standard.group(5).replace(',', '')
                dividend_shares = match_standard.group(6).replace(',', '')
                results.append((product_name, fund_market_code, dividend_amount, dividend_shares, '证达通基金'))
        
        # 3. 如果没有检测到汇总列表格式，则尝试解析格式2：单笔确认单格式
        # 【注意：这部分完全保持了原脚本的逻辑】
        if not is_tabular:
            fund_market_code = ''
            dividend_shares = ''
            
            for line in lines:
                # 提取基金代码
                if '基金代码' in line:
                    match = re.search(r'基金代码[：:]\s*([0-9]{6})', line)
                    if match:
                        fund_market_code = match.group(1)
                
                # 提取分红份额
                if '分红份额' in line:
                    match = re.search(r'分红份额[：:]\s*([\d,]+\.?\d*)', line)
                    if match:
                        dividend_shares = match.group(1).replace(',', '')
            
            # 只有当关键字段都提取到时才添加结果
            if fund_market_code and dividend_shares:
                # 根据指示：单笔格式下，金额和份额均取“分红份额”的值
                dividend_amount = dividend_shares
                results.append((product_name, fund_market_code, dividend_amount, dividend_shares, '证达通基金'))
        
        return results
    
    #基煜基金
    def extract_jiyu_fields(lines):

        # 1. 产品名称（从客户名称字段提取）
        product_name = ''
        for line in lines:
            if '客户名称' in line:
                # 提取客户名称后的内容
                match = re.search(r'客户名称\s*(.*)', line)
                if match:
                    # 去除首尾空格及内部空格
                    product_name = match.group(1).strip().replace(' ', '').replace('\u3000', '')
                    break

        # 2. 基金市场代码（从产品代码字段提取）
        fund_market_code = ''
        for line in lines:
            if '产品代码' in line:
                # 提取数字或字母组合
                match = re.search(r'产品代码\s*([0-9A-Za-z]+)', line)
                if match:
                    fund_market_code = match.group(1).strip()
                    break

        # 3. 派送金额（从再投资金额提取，自动忽略"元"）
        dividend_amount = ''
        for line in lines:
            if '再投资金额' in line:
                # 匹配数字部分，支持逗号分隔
                match = re.search(r'再投资金额\s*([\d,]+\.?\d*)', line)
                if match:
                    dividend_amount = match.group(1).replace(',', '')
                    break

        # 4. 派送份额（从再投资份额提取，自动忽略"份"）
        dividend_shares = ''
        for line in lines:
            if '再投资份额' in line:
                # 匹配数字部分，支持逗号分隔
                match = re.search(r'再投资份额\s*([\d,]+\.?\d*)', line)
                if match:
                    dividend_shares = match.group(1).replace(',', '')
                    break

        return product_name, fund_market_code, dividend_amount, dividend_shares, '基煜基金'

    #宁波银行
    def extract_ningboBank_fields(lines):
        
        # 1. 产品名称（从客户名称字段提取）
        product_name = ''
        for line in lines:
            if '客户名称' in line:
                # 提取客户名称后的内容
                match = re.search(r'客户名称\s*(.*)', line)
                if match:
                    # 去除空格
                    product_name = match.group(1).strip().replace(' ', '').replace('\u3000', '')
                    break

        # 2. 基金市场代码（从产品代码字段提取）
        fund_market_code = ''
        for line in lines:
            if '产品代码' in line:
                # 提取产品代码后的数字
                match = re.search(r'产品代码\s*([0-9A-Za-z]+)', line)
                if match:
                    fund_market_code = match.group(1).strip()
                    break

        # 3. 派送金额和派送份额（均从红利份额字段提取）
        # 注意：用户指定金额和份额都取"红利份额（份）"的值
        dividend_amount = ''
        dividend_shares = ''
        for line in lines:
            if '红利份额（份）' in line:
                # 提取紧跟在"红利份额（份）"后的数值
                match = re.search(r'红利份额（份）\s*([\d,]+\.?\d*)', line)
                if match:
                    value = match.group(1).replace(',', '')
                    dividend_shares = value
                    dividend_amount = value # 按照指示，金额也取份额的值
                    break

        return product_name, fund_market_code, dividend_amount, dividend_shares, '宁波银行'

    # 国信嘉利基金
    def extract_guoxinjiali_fields(lines, filename):
        """
        提取国信嘉利基金分红确认单字段
        样本结构 (跨行):
        Line N-1: '万联资管臻 2025122 富安达现金通货'
        Line N:   '选3号FOF集 3 710501 币A 分红 - - - - 0.00 87.09 0.00 ...'
        策略：优先从文件名提取全称，否则尝试从内容拼接并补全
        """
        product_name = ''
        fund_market_code = ''
        dividend_amount = ''
        dividend_shares = ''

        # --- 策略1：优先从文件名提取产品名称 ---
        # 原始文件名格式: "万联资管臻选3号FOF集合资产管理计划_交易确认单_2025-12-23.pdf"
        # 逻辑：提取开头 到 "_交易确认单" 之前的内容
        if filename:
            try:
                base_name = os.path.basename(filename)
                # 查找 "_交易确认单" 的位置
                end_idx = base_name.find('_交易确认单')
                
                if end_idx != -1:
                    extracted_name = base_name[:end_idx]
                    # 额外清洗：如果文件名确实偶尔包含【】，可以做一个去除去除操作，保证纯净
                    if '】' in extracted_name:
                        extracted_name = extracted_name.split('】')[-1]
                        
                    product_name = extracted_name
            except Exception:
                pass # 如果文件名解析出错，保持为空，后续逻辑会处理

        for i, line in enumerate(lines):
            # 关键定位点是 "分红"
            if '分红' in line:
                parts = line.split()
                # 找到 "分红" 所在的索引位置
                try:
                    div_idx = parts.index('分红')
                except ValueError:
                    continue

                # 1. 提取基金代码 (fund_market_code)
                # 代码通常在 "分红" 之前，且为6位数字
                # 在样本中：['选3号FOF集', '3', '710501', '币A', '分红'...]
                # 从分红位置向前寻找
                for j in range(div_idx - 1, -1, -1):
                    token = parts[j]
                    if re.match(r'^\d{6}$', token):
                        fund_market_code = token
                        break

                # 2. 提取红利金额和份额
                # 样本逻辑：分红 [申请金额] [申请份额] [确认金额] [确认份额]
                # 样本数据：分红 - - - - 0.00 87.09
                # 索引推算：分红(idx) -> -(idx+1) -> -(idx+2) -> -(idx+3) -> -(idx+4) -> 0.00(idx+5) -> 87.09(idx+6)
                # 用户要求：金额和份额均取 "确认份额" (即 87.09)
                target_idx = div_idx + 6
                if target_idx < len(parts):
                    value = parts[target_idx].replace(',', '')
                    dividend_shares = value
                    dividend_amount = value # 按指示，金额也取份额的值

                # 3. 提取产品名称 (product_name) - 处理跨行
                # 产品名称的第一部分在上一行 (i-1) 的第一个元素
                # 产品名称的第二部分在当前行 (i) 的第一个元素
                 # --- 策略2：如果文件名没提取到，则从内容提取并补全 ---
                if not product_name and i > 0:
                    prev_line_parts = lines[i-1].split()
                    if prev_line_parts and parts:
                        part1 = prev_line_parts[0] # 万联资管臻
                        part2 = parts[0]           # 选3号FOF集
                        raw_name = part1 + part2
                        
                        # 自动补全逻辑
                        # --- 自动补全逻辑 ---
                        # 针对文件名过长被截断或OCR识别不全的情况进行修复
                        # 目标后缀1: 集合资产管理计划
                        # 目标后缀2: 单一资产管理计划
                        
                        # 1. 处理“集合资产管理计划”的残缺情况
                        if raw_name.endswith('集'):
                            product_name = raw_name + '合资产管理计划'
                        elif raw_name.endswith('集合'):
                            product_name = raw_name + '资产管理计划'
                        elif raw_name.endswith('集合资'):
                            product_name = raw_name + '产管理计划'
                        elif raw_name.endswith('集合资产'):
                            product_name = raw_name + '管理计划'
                        
                        # 2. 处理“单一资产管理计划”的残缺情况
                        elif raw_name.endswith('单'):
                            product_name = raw_name + '一资产管理计划'
                        elif raw_name.endswith('单一'):
                            product_name = raw_name + '资产管理计划'
                        elif raw_name.endswith('单一资'):
                            product_name = raw_name + '产管理计划'
                        elif raw_name.endswith('单一资产'):
                            product_name = raw_name + '管理计划'
                            
                        # 3. 处理通用的“资产管理计划”残缺情况 (如果前面没有集/单，或者被截断在“计划”之前)
                        elif raw_name.endswith('资产管理计'):
                            product_name = raw_name + '划'
                        elif raw_name.endswith('资产管理'):
                            product_name = raw_name + '计划'
                            
                        else:
                            product_name = raw_name # 无法补全则原样输出
                
                # 找到一条有效记录后即可退出 (假设单文件单记录)
                break

        return product_name, fund_market_code, dividend_amount, dividend_shares, '国信嘉利基金'

    # 攀赢基金
    def extract_panying_fields(lines):
        product_name = ''
        fund_market_code = ''
        dividend_amount = ''
        dividend_shares = ''
        
        for i, line in enumerate(lines):
            line = line.strip()
            
            # 1. 提取产品名称（从客户名称字段提取）
            # 样本：客户名称 万联资管民利2号集合资产管理计划
            if '客户名称' in line:
                match = re.search(r'客户名称\s*[:：]?\s*(.*)', line)
                if match:
                    # 去除空格和全角空格
                    product_name = match.group(1).strip().replace(' ', '').replace('\u3000', '')
            
            # 2. 提取基金代码（从产品代码字段提取）
            # 样本：产品代码 004179 产品名称
            if '产品代码' in line:
                match = re.search(r'产品代码\s*[:：]?\s*([0-9]{6})', line)
                if match:
                    fund_market_code = match.group(1)
            
            # 3. 提取红利金额（从所得现金字段提取）
            # 样本：所得现金（元） 1,154.93
            if '所得现金' in line:
                # 兼容中文括号（）和英文括号()
                match = re.search(r'所得现金[（(]元[）)]\s*[:：]?\s*([\d,]+\.?\d*)', line)
                if match:
                    dividend_amount = match.group(1).replace(',', '')
            
            # 4. 提取红利份额（从所得份额字段提取）
            # 样本：所得份额（份）\n1,154.93 (即第一行只有标签，第二行开头是数值)
            if '所得份额' in line:
                # 策略A：先在当前行查找（防止数值其实在同一行）
                match = re.search(r'所得份额[（(]份[）)]\s*[:：]?\s*([\d,]+\.?\d*)', line)
                if match:
                    dividend_shares = match.group(1).replace(',', '')
                # 策略B：如果当前行没找到数值，且还有下一行，检查下一行开头是否为数字
                elif i + 1 < len(lines):
                    next_line = lines[i+1].strip()
                    # 匹配行首的数字
                    match_next = re.match(r'^([\d,]+\.?\d*)', next_line)
                    if match_next:
                        dividend_shares = match_next.group(1).replace(',', '')

        return product_name, fund_market_code, dividend_amount, dividend_shares, '攀赢基金'

    # 5. 遍历分红文件夹
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

    # 使用 os.walk 递归遍历 target_path 下的所有子文件夹
    # 递归遍历 target_path 下所有目录
    for root, dirs, files in os.walk(target_path):
        # 我们应该检查 root 路径中是否包含 "分红"
        if "分红" not in root:
             continue
        log(f"扫描目录: {root}", log_text)
    
        pdf_files = [f for f in files if f.lower().endswith('.pdf')]
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
                is_haomai = any('好买基金' in l for l in lines[:2])
                is_tiantian = ('天天基金' in file) or any('天天基金' in l for l in lines[3:])
                is_xingzheng = any('兴证全球基金' in l for l in lines[:2])
                is_lide = any('利得基金' in l for l in lines[3:])
                is_changliang = any('长量基金' in l for l in lines[:2])
                is_yingmi = ('盈米' in file) or any('盈米' in l for l in lines[:3])
                is_zhaoyingtong = any('招赢通' in l for l in lines[:2])
                is_youchu = ('邮储' in file)
                is_pingan = any('行E通' in l for l in lines[5:])
                is_jiaohang = ('交e通' in file) or any('交通银行' in l for l in lines[:2])
                is_hexun = any('和讯信息科技有限公司' in l for l in lines[3:])
                is_jianhang = ('建行' in file) or any('客 户 名 称' in l for l in lines) #比较脆弱，考虑"红股"
                is_tengyuan = ('腾元' in file) or any('腾元基金' in l for l in lines[5:])
                is_wangjin = ('网金' in file) or any('网金基金' in l for l in lines[5:]) #最后五行是否包含
                is_jd = ('肯特瑞基金' in file) or any('肯特瑞' in l for l in lines[:2])
                is_ronglianchuang = any('融联创' in l for l in lines[:2])
                is_liantai = ('北极星' in file) or any('联泰' in l for l in lines[:2])
                is_minsheng = ('民生同业e+' in file) or any('同业e+' in l for l in lines[2:])
                is_zdt = any('证达通' in l for l in lines)
                is_jiyu= any('基煜基金' in l for l in lines[:2])
                is_ningboBank = ('宁波' in file) or any('同业客户付款账户信息' in l for l in lines[5:])
                is_guoxinjiali = any('国信嘉利基金' in l for l in lines[:2])
                is_panying = ('攀赢' in file) or any('攀赢' in l for l in lines[:2])


                if is_haomai:
                    product_name, fund_market_code, dividend_amount, dividend_shares, platform = extract_haomai_fields(text, lines)
                    records = [(product_name, fund_market_code, dividend_amount, dividend_shares, platform, None)]  # 添加None作为红利截止日期
                elif is_tiantian:
                    product_name, fund_market_code, dividend_amount, dividend_shares, platform = extract_tiantian_fields(lines)
                    records = [(product_name, fund_market_code, dividend_amount, dividend_shares, platform, None)]
                elif is_xingzheng:
                    product_name, fund_market_code, dividend_amount, dividend_shares, platform = extract_xingzheng_fields(lines)
                    records = [(product_name, fund_market_code, dividend_amount, dividend_shares, platform, None)]
                elif is_lide:
                    product_name, fund_market_code, dividend_amount, dividend_shares, platform = extract_lide_fields(lines)
                    records = [(product_name, fund_market_code, dividend_amount, dividend_shares, platform, None)]
                elif is_changliang:
                    product_name, fund_market_code, dividend_amount, dividend_shares, platform = extract_changliang_fields(lines)
                    records = [(product_name, fund_market_code, dividend_amount, dividend_shares, platform, None)]
                elif is_yingmi:
                    yingmi_records = extract_yingmi_fields(lines)
                    records = [(pn, fmc, da, ds, platform, None) for pn, fmc, da, ds, platform in yingmi_records]  # 为盈米的多条记录添加None
                elif is_zhaoyingtong:
                    product_name, fund_market_code, dividend_amount, dividend_shares, platform = extract_zhaoyingtong_fields(lines)
                    records = [(product_name, fund_market_code, dividend_amount, dividend_shares, platform, None)]
                elif is_youchu:
                    product_name, fund_market_code, dividend_amount, dividend_shares, platform = extract_youchu_fields(lines)
                    records = [(product_name, fund_market_code, dividend_amount, dividend_shares, platform, None)]
                elif is_pingan:
                    product_name, fund_market_code, dividend_amount, dividend_shares, platform = extract_pingan_fields(lines)
                    records = [(product_name, fund_market_code, dividend_amount, dividend_shares, platform, None)]
                elif is_jiaohang:
                    product_name, fund_market_code, dividend_amount, dividend_shares, platform, dividend_end_date = extract_jiaohang_fields(lines)
                    records = [(product_name, fund_market_code, dividend_amount, dividend_shares, platform, dividend_end_date)]
                elif is_hexun:
                    product_name, fund_market_code, dividend_amount, dividend_shares, platform = extract_hexun_fields(lines)
                    records = [(product_name, fund_market_code, dividend_amount, dividend_shares, platform, None)]
                elif is_jianhang:
                    product_name, fund_market_code, dividend_amount, dividend_shares, platform = extract_jianhang_fields(lines)
                    records = [(product_name, fund_market_code, dividend_amount, dividend_shares, platform, None)]
                elif is_tengyuan:
                    product_name, fund_market_code, dividend_amount, dividend_shares, platform = extract_tengyuan_fields(lines)
                    records = [(product_name, fund_market_code, dividend_amount, dividend_shares, platform, None)]
                elif is_wangjin:
                    product_name, fund_market_code, dividend_amount, dividend_shares, platform = extract_wangjin_fields(lines)
                    records = [(product_name, fund_market_code, dividend_amount, dividend_shares, platform, None)]
                elif is_jd:
                    product_name, fund_market_code, dividend_amount, dividend_shares, platform = extract_jd_fields(lines)
                    records = [(product_name, fund_market_code, dividend_amount, dividend_shares, platform, None)]
                elif is_ronglianchuang:
                    product_name, fund_market_code, dividend_amount, dividend_shares, platform = extract_ronglianchuang_fields(lines)
                    records = [(product_name, fund_market_code, dividend_amount, dividend_shares, platform, None)]
                elif is_minsheng:
                    product_name, fund_market_code, dividend_amount, dividend_shares, platform = extract_minsheng_fields(lines)
                    records = [(product_name, fund_market_code, dividend_amount, dividend_shares, platform, None)]
                elif is_zdt:
                    zdt_records = extract_zdt_fields(lines)
                    records = [(pn, fmc, da, ds, platform, None) for pn, fmc, da, ds, platform in zdt_records]
                elif is_liantai:
                    liantai_records = extract_liantai_fields(lines)
                    records = [(pn, fmc, da, ds, platform, None) for pn, fmc, da, ds, platform in liantai_records]  # 为联泰的多条记录添加None
                elif is_jiyu:
                    product_name, fund_market_code, dividend_amount, dividend_shares, platform = extract_jiyu_fields(lines)
                    records = [(product_name, fund_market_code, dividend_amount, dividend_shares, platform, None)]
                elif is_ningboBank:
                    product_name, fund_market_code, dividend_amount, dividend_shares, platform = extract_ningboBank_fields(lines)
                    records = [(product_name, fund_market_code, dividend_amount, dividend_shares, platform, None)]
                elif is_guoxinjiali:
                    product_name, fund_market_code, dividend_amount, dividend_shares, platform = extract_guoxinjiali_fields(lines, file)
                    records = [(product_name, fund_market_code, dividend_amount, dividend_shares, platform, None)]
                elif is_panying:
                    product_name, fund_market_code, dividend_amount, dividend_shares, platform = extract_panying_fields(lines)
                    records = [(product_name, fund_market_code, dividend_amount, dividend_shares, platform, None)]
                else:
                    continue

                for product_name, fund_market_code, dividend_amount, dividend_shares, platform, custom_end_date in records:
                    temp_df = pd.DataFrame([{
                        '产品名称': product_name,
                        '基金市场代码': fund_market_code,
                        '派送金额': dividend_amount,
                        '派送份额': dividend_shares,
                        '基金平台': platform
                    }])
                    temp_df['派送金额'] = pd.to_numeric(temp_df['派送金额'], errors='coerce').round(2)
                    temp_df['派送份额'] = pd.to_numeric(temp_df['派送份额'], errors='coerce').round(2)
                    temp_df['账套编号'] = temp_df['产品名称'].map(product_code_dict)
                    temp_df['交易市场'] = '国内银行间'
                    temp_df['日期'] = today_str
                    # 根据是否有自定义红利截止日期来设置该字段
                    if custom_end_date is not None:
                        temp_df['红利截止日期'] = custom_end_date  # 使用交通银行返回的红利截止日期
                    else:
                        temp_df['红利截止日期'] = yesterday_str    # 使用默认的昨天日期
                        
                    temp_df['持仓分类'] = ''
                    temp_df['产品代码'] = ''
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
    output_file = os.path.join(output_folder, "【境内基金业务】红利再投.xls")
    try:
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            target_df.to_excel(writer, sheet_name='Sheet1', index=False)
        log(f"已汇总输出到: {output_file}", log_text)

        # ====== 新增：分组合并并输出 ======
        merge_cols = ['账套编号', '基金市场代码']
        sum_cols = ['派送份额', '派送金额']
        # 其他列取第一条，但基金平台需要特殊处理
        first_cols = [col for col in target_cols if col not in merge_cols + sum_cols + ['基金平台']]
        grouped = target_df.groupby(merge_cols, as_index=False)

        # 定义合并基金平台的函数
        def merge_platforms(platform_series):
            # 去重并用顿号连接
            unique_platforms = platform_series.dropna().unique()
            return '、'.join(unique_platforms)
        
        # 分别处理各种聚合方式
        agg_dict = {
            '派送份额': 'sum',
            '派送金额': 'sum',
            '基金平台': merge_platforms,  # 使用自定义函数合并平台
            **{col: 'first' for col in first_cols}
        }

        merged_df = grouped.agg(agg_dict)
        
        merged_output_file = os.path.join(output_folder, "【境内基金业务】红利再投_合并后.xls")
        # 调整列顺序为target_cols
        merged_df = merged_df[target_cols]
        with pd.ExcelWriter(merged_output_file, engine='openpyxl') as writer:
            merged_df.to_excel(writer, sheet_name='Sheet1', index=False)
        log(f"合并后数据已输出到: {merged_output_file}", log_text)
        # ====== 新增结束 ======

        return output_folder
    except Exception as e:
        log(f"写入Excel失败: {e}", log_text)
        return False