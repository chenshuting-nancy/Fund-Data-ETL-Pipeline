import os
import json
import pdfplumber
import re
import pandas as pd
from datetime import datetime, timedelta
from utils.common import log

def run_purchase_extract(folder_path, json_path, log_text):
    """运行申购申请单提取
    
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
    target_cols = ['账套编号', '基金市场代码', '交易市场', '日期','业务类别','数量','金额', '手续费','佣金','交易对手','资金账户','赎回到账日期', '股东账户','席位号','产品名称','基金平台']

    # 3. 读取产品代码
    try:
        with open(json_path, 'r', encoding='utf-8') as f:
            product_code_dict = json.load(f)
    except Exception as e:
        log(f"产品代码加载失败: {e}", log_text)
        return False

    # 4. 平台提取函数
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

        # 3. 申请金额
        m3 = re.search(r'申请金额小写[：: ]*([\d,]+\.\d+)', text)
        amount = m3.group(1).replace(',', '') if m3 else ''

        return product_name, fund_market_code, amount, '好买基金'

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
        fund_market_code = ''
        for line in lines:
            if '基金代码' in line:
                match = re.search(r'基金代码\s*([0-9]{6})', line)
                if match:
                    fund_market_code = match.group(1)
                    break

        # 3. 申请金额
        amount = ''
        for line in lines:
            if '申请金额' in line:
                match = re.search(r'申请金额\s*([\d,]+\.\d+)', line)
                if match:
                    amount = match.group(1).replace(',', '')
                    break

        return product_name, fund_market_code, amount, '天天基金'

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
        fund_market_code = ''
        for line in lines:
            if '基金代码' in line:
                # 基金代码可能出现在“基金名称 ... 基金代码 ...”这一行
                match = re.search(r'基金代码\s*([0-9]{6})', line)
                if match:
                    fund_market_code = match.group(1)
                    break

        # 3. 申请金额（元）
        amount = ''
        for line in lines:
            if '申请金额（元）' in line:        
                match = re.search(r'申请金额（元）\s*([\d,]+\.\d+)', line)
                if match:
                    amount = match.group(1).replace(',', '')
                    break

        return product_name, fund_market_code, amount, '利得基金'
    
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
        fund_market_code = ''
        for line in lines:
            if '基金代码' in line:
                # 可能有内容如“基金代码 210013 基金名称...”
                match = re.search(r'基金代码\s*([0-9]{6})', line)
                if match:
                    fund_market_code = match.group(1)
                    break

        # 3. 申请金额（申请金额，去掉(元)等）
        amount = ''
        for line in lines:
            if '申请金额' in line:
                match = re.search(r'申请金额\s*([\d,]+\.\d+)', line)
                if match:
                    amount = match.group(1).replace(',', '')
                    break

        return product_name, fund_market_code, amount, '长量基金'

    #盈米
    def extract_yingmi_fields(lines):
        """
        从盈米基金分红PDF解析行中提取所有申购申请记录。
        返回: 列表，每个元素为 (产品名称, 基金市场代码, 金额)
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

        # 2. 提取所有申购记录
        results = []
        for i, line in enumerate(lines):
            # 检测交易类型行 - 修改匹配条件以适应PDF实际格式
            if ('交易类型' in line and '申购' in line) or ('交易类型：' in line and '申购' in line):
                
                # 在附近行查找基金代码
                fund_market_code = ''
                amount = ''
                
                # 查找基金代码 - 可能在当前行或后续几行
                for j in range(i, min(i+5, len(lines))):
                    if '基金代码' in lines[j]:
                        # 尝试两种匹配模式
                        match1 = re.search(r'基金代码[:：]\s*([0-9]{6})', lines[j])
                        match2 = re.search(r'基金代码\s*([0-9]{6})', lines[j])
                        if match1:
                            fund_market_code = match1.group(1)
                        elif match2:
                            fund_market_code = match2.group(1)
                        
                        break
                
                # 查找申请金额 - 可能在当前行或后续几行
                for j in range(i, min(i+5, len(lines))):
                    if '申请金额' in lines[j]:
                        # 尝试几种匹配模式
                        match1 = re.search(r'申请金额[:：]\s*([\d,\.]+)', lines[j])
                        match2 = re.search(r'申请金额[:：]?\s*([\d,]+\.\d+)', lines[j])
                        
                        if match1:
                            amount = match1.group(1).replace(',', '')
                        elif match2:
                            amount = match2.group(1).replace(',', '')               
                        break
                
                # 如果找到了基金代码和金额，添加到结果
                if fund_market_code and amount:
                    results.append((product_name, fund_market_code, amount, '盈米基金'))

        return results

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

        # 3. 申请金额（从申请金额字段提取，去除"元"字符）
        amount = ''
        for line in lines:
            if '申请金额' in line:
                # 提取金额，支持带逗号的格式，并去除"元"字符
                match = re.search(r'申请金额\s*([\d,]+\.?\d*)元?', line)
                if match:
                    amount = match.group(1).replace(',', '')
                    break

        return product_name, fund_market_code, amount,'平安行E通'

    #交通银行交e通
    def extract_jiaohang_fields(lines):

        # 1. 产品名称（从投资者信息字段提取）
        product_name = ''
        for i, line in enumerate(lines):
            if '投资者信息' in line:
                # 投资者信息后的下一行通常是产品名称
                if i + 1 < len(lines):
                    next_line = lines[i + 1].strip()
                    if (next_line and 
                        '基金账户' not in next_line and 
                        '产品信息' not in next_line and
                        '客户信息' not in next_line and
                        len(next_line) > 5):  # 避免单个字符的干扰
                        product_name = next_line
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

        # 3. 申请金额（从申请金额/份额字段提取）
        amount = ''
        for line in lines:
            if '申请金额/份额' in line:
                # 提取金额，支持带小数点的格式
                match = re.search(r'申请金额/份额\s*([\d,]+\.?\d*)', line)
                if match:
                    amount = match.group(1).replace(',', '')
                    break

        return product_name, fund_market_code, amount,'交e通'

    #网金基金
    def extract_wangjin_fields(lines):
    
        # 1. 产品名称（从投资者名称字段提取，需要处理跨行情况）
        product_name = ''
        for i, line in enumerate(lines):
            if '投资者名称' in line:
                # 先检查当前行是否有产品名称
                parts = line.split('投资者名称')
                if len(parts) > 1 and parts[1].strip():
                    product_name = parts[1].strip()
                else:
                    # 如果当前行没有，检查前一行
                    if i > 0:
                        prev_line = lines[i-1].strip()
                        # 排除一些明显不是产品名称的行
                        if (prev_line and 
                            '机构名称' not in prev_line and
                            '网点名称' not in prev_line and
                            '基金' not in prev_line and
                            '─' not in prev_line and  # 排除分隔线
                            len(prev_line) > 10):  # 产品名称通常较长
                            product_name = prev_line
                break
        
        # 清理产品名称中的多余空格和特殊字符
        product_name = product_name.replace(' ', '').replace('\u3000', '').replace('\n', '')

        # 2. 基金市场代码（从基金代码字段提取，可能在同一行或跨行）
        fund_market_code = ''
        for i, line in enumerate(lines):
            if '基金代码' in line:
                # 先检查当前行是否有代码（格式1：基金代码 005151 基金名称）
                match = re.search(r'基金代码\s*([0-9]{6})', line)
                if match:
                    fund_market_code = match.group(1)
                else:
                    # 如果当前行没有，检查下一行（格式2：跨行显示）
                    if i + 1 < len(lines):
                        next_line = lines[i + 1].strip()
                        # 检查下一行是否是6位数字
                        if re.match(r'^[0-9]{6}$', next_line):
                            fund_market_code = next_line
                break

        # 3. 申购金额（支持两种格式的字段名）
        amount = ''
        # 格式1：申购金额小写
        # 格式2：申购金额（小写）
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

        return product_name, fund_market_code, amount, '网金基金'

    #腾元基金
    def extract_tengyuan_fields(lines):
        
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

        # 3. 申购金额（从申购金额（小写）字段提取，去除逗号）
        amount = ''
        for line in lines:
            if '申购金额（小写）' in line:
                # 提取金额，支持带逗号的格式，并去除逗号
                match = re.search(r'申购金额（小写）\s*([\d,]+\.?\d*)', line)
                if match:
                    amount = match.group(1).replace(',', '')
                    break

        return product_name, fund_market_code, amount, '腾元基金'

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
                        ['交易账号', '申请工作日', '基金代码', '申请金额', '重要提示']) or
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

        # 3. 申请金额（从申请金额字段提取，去除逗号）
        amount = ''
        for line in lines:
            if '申请金额' in line:
                # 提取金额，支持带逗号的格式，并去除逗号
                match = re.search(r'申请金额\s*([\d,]+\.?\d*)', line)
                if match:
                    amount = match.group(1).replace(',', '')
                    break

        return product_name, fund_market_code, amount, '和讯科技'

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

        # 3. 申请金额（从申请金额(元)字段提取，去除逗号）
        amount = ''
        for line in lines:
            if '申请金额(元)' in line:
                # 提取金额，支持带逗号的格式，并去除逗号
                match = re.search(r'申请金额\(元\)\s*([\d,]+\.?\d*)', line)
                if match:
                    amount = match.group(1).replace(',', '')
                    break

        return product_name, fund_market_code, amount, '京东肯特瑞'

    #民生同业e+
    def extract_minsheng_fields(lines):
        
        # 1. 产品名称（从客户名称字段提取）
        product_name = ''
        for line in lines:
            if '客户名称：' in line:
                # 提取客户名称后、交易类型前的内容
                parts = line.split('客户名称：')
                if len(parts) > 1:
                    # 进一步分割，去掉交易类型部分
                    name_part = parts[1].split('交易类型：')[0].strip()
                    product_name = name_part
                    break
        
        # 清理产品名称中的多余空格和特殊字符
        product_name = product_name.replace(' ', '').replace('\u3000', '').replace('\n', '')

        # 2. 基金市场代码（从产品代码字段提取）
        fund_market_code = ''
        for line in lines:
            if '产品代码：' in line:
                # 提取产品代码后的6位数字
                match = re.search(r'产品代码：\s*([0-9]{6})', line)
                if match:
                    fund_market_code = match.group(1)
                    break

        # 3. 申购金额（从委托金额/委托份额字段提取，去除逗号）
        amount = ''
        for line in lines:
            if '委托金额/委托份额：' in line:
                # 提取金额，支持带逗号的格式，并去除逗号
                match = re.search(r'委托金额/委托份额：\s*([\d,]+\.?\d*)', line)
                if match:
                    amount = match.group(1).replace(',', '')
                    break

        return product_name, fund_market_code, amount, '民生同业e+'

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

        # 3. 申请金额（分红金额，去除CNY）
        amount = ''
        for line in lines:
            if '申请金额' in line and 'CNY' in line:
                match = re.search(r'CNY\s*([\d,\.]+)', line)
                if match:
                    amount = match.group(1).replace(',', '')
                    break

        return product_name, fund_market_code, amount, '招赢通基金'

    #融联创同业交易平台
    def extract_ronglianchuang_fields(lines):
        
        # 1. 产品名称（从投资者名称字段提取，处理跨行情况）
        product_name = ''
        for i, line in enumerate(lines):
            if '投资者名称' in line:
                # 提取投资者名称后的内容
                parts = line.split('投资者名称')
                if len(parts) > 1:
                    current_part = parts[1].strip()
                    product_name += current_part
                
                # 检查后续行是否包含产品名称的剩余部分（如"管理计划"）
                j = i + 1
                while j < len(lines) and j < i + 3:  # 最多检查后续2行
                    next_line = lines[j].strip()
                    # 如果遇到其他字段标识，停止拼接
                    if (any(keyword in next_line for keyword in 
                        ['银行账号', '开户行名称', '基金代码', '基金名称', '申请日期', '申请金额']) or
                        len(next_line) == 0):
                        break
                    # 拼接产品名称
                    product_name += next_line
                    j += 1
                break
        
        # 清理产品名称中的多余空格和特殊字符
        product_name = product_name.replace(' ', '').replace('\u3000', '').replace('\n', '')
        
        # 特殊处理：如果产品名称以"管理计划"开头，需要重新排列
        if product_name.startswith('管理计划'):
            # 找到"管理计划"的位置，重新排列
            parts = product_name.split('管理计划')
            if len(parts) > 1:
                # 重新组合：产品名称 + 管理计划
                product_name = parts[1] + '管理计划'

        # 2. 基金市场代码（从基金代码字段提取）
        fund_market_code = ''
        for line in lines:
            if '基金代码' in line:
                # 提取基金代码（支持6位数字）
                match = re.search(r'基金代码\s*([0-9]{6})', line)
                if match:
                    fund_market_code = match.group(1)
                    break

        # 3. 申请金额（从申请金额字段提取，去除逗号和"元"字符）
        amount = ''
        for line in lines:
            if '申请金额' in line:
                # 提取申请金额，支持带逗号的格式，并去除逗号和"元"字符
                match = re.search(r'申请金额\s*([\d,]+\.?\d*)元?', line)
                if match:
                    amount = match.group(1).replace(',', '')
                    break

        return product_name, fund_market_code, amount, '融联创同业交易平台'

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
                
                # 在接下来的几行中查找基金代码和申购金额
                lookahead = 1
                while (i + lookahead) < N and lookahead <= 8:  # 增加查找范围
                    subline = lines[i + lookahead]
                    
                    # 查找基金代码
                    if '基金代码' in subline:
                        # 提取基金代码（支持6位数字）
                        match = re.search(r'基金代码\s+([0-9]{6})', subline)
                        if match:
                            fund_market_code = match.group(1)
                    
                    # 查找申购金额
                    if '申请金额(元)' in subline:
                        match = re.search(r'申请金额\(元\)\s*([\d,]+\.?\d*)', subline)
                        if match:
                            amount = match.group(1).replace(',', '')
                            
                    # 如果遇到下一个交易信息块，停止查找
                    if lookahead > 1 and '交易信息' in subline:
                        break
                        
                    lookahead += 1
                
                # 如果找到了基金代码和红利份额，添加到结果中
                if fund_market_code and amount :
                    results.append((product_name, fund_market_code, amount, '联泰基金'))
            
            i += 1
        
        return results

    #基煜基金
    def extract_jiyu_fields(lines):
        
        # 1. 产品名称（从账户名称字段提取）
        product_name = ''
        for line in lines:
            if '账户名称：' in line and '付款账户' not in line:  # 排除付款账户信息部分的账户名称
                # 提取账户名称后的产品名称
                match = re.search(r'账户名称：\s*(.*)', line)
                if match:
                    product_name = match.group(1).strip()
                    break
        
        # 清理产品名称中的多余空格和特殊字符
        product_name = product_name.replace(' ', '').replace('\u3000', '').replace('\n', '')

        # 2. 基金市场代码（从产品代码字段提取）
        fund_market_code = ''
        for line in lines:
            if '产品代码：' in line:
                # 提取6位数字的基金代码
                match = re.search(r'产品代码：\s*([0-9]{6})', line)
                if match:
                    fund_market_code = match.group(1)
                    break

        # 3. 申购金额（从申购金额(小写)字段提取，去除逗号和"元"）
        amount = ''
        for line in lines:
            if '申购金额(小写)：' in line or '申购金额（小写）：' in line:
                # 提取金额，支持带逗号的格式，并去除逗号和"元"
                match = re.search(r'申购金额[（(]小写[）)]：\s*([\d,]+\.?\d*)元?', line)
                if match:
                    amount = match.group(1).replace(',', '')
                    break

        return product_name, fund_market_code, amount, '基煜基金'

    #宁波银行
    def extract_ningboBank_fields(lines):
        
        # 1. 产品名称（从客户名称字段提取）
        product_name = ''
        for line in lines:
            if '客户名称' in line:
                parts = line.split('客户名称')
                if len(parts) > 1:
                    # 去除可能存在的其他字段（如交易账号等）
                    name_part = parts[1].strip()
                    # 如果同一行有其他字段标识，只取前面的部分
                    if '交易账号' in name_part:
                        name_part = name_part.split('交易账号')[0].strip()
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

        # 3. 申请金额（从申请金额（元）字段提取，去除逗号）
        amount = ''
        for line in lines:
            if '申请金额（元）' in line:
                # 提取金额，支持带逗号的格式，并去除逗号
                match = re.search(r'申请金额（元）\s*([\d,]+\.?\d*)', line)
                if match:
                    amount = match.group(1).replace(',', '')
                    break

        return product_name, fund_market_code, amount, '宁波银行'

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

        # 3. 申请金额（从申请金额字段提取，去除逗号和"元"）
        amount = ''
        for line in lines:
            if '申请金额' in line:
                # 提取金额，支持带逗号的格式，并去除逗号和"元"
                match = re.search(r'申请金额\s*([\d,]+\.?\d*)元?', line)
                if match:
                    amount = match.group(1).replace(',', '')
                    break

        return product_name, fund_market_code, amount, '国信嘉利基金'

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

        # 3. 申购金额（从申购金额（小写）字段提取，去除逗号和"元"）
        amount = ''
        for line in lines:
            if '申购金额' in line and '小写' in line:
                # 提取金额，支持带逗号的格式，并去除逗号和"元"
                # 匹配格式如：申购金额（小写） 49,990,000.00元
                match = re.search(r'申购金额[（(]小写[）)]\s*([\d,]+\.?\d*)元?', line)
                if match:
                    amount = match.group(1).replace(',', '')
                    break

        return product_name, fund_market_code, amount, '攀赢基金'

    # 证达通基金（兼容汇总格式与单笔格式）
    def extract_zdt_fields(lines):
        # 1. 公共部分：提取产品名称（投资者名称）
        product_name = ''
        for line in lines:
            if '投资者名称' in line:
                # 提取名称，非贪婪匹配直到遇到“生成时间”或行尾
                match = re.search(r'投资者名称[：:]\s*(.+?)(?:\s+生成时间|$)', line)
                if match:
                    product_name = match.group(1).strip().replace(' ', '').replace('\u3000', '')
                break
        
        results = []
        
        # 判断是否为单笔格式（特征：有"申购受理单"且通常没有"汇总"字样）
        is_single_mode = any('申购受理单' in l for l in lines[:2]) and not any('汇总' in l for l in lines[:2])
        
        if is_single_mode:
            # === 策略 A：单笔格式提取 ===
            fund_market_code = ''
            amount = ''
            
            for line in lines:
                # 提取基金代码 (例如：基金代码：583101)
                if '基金代码' in line:
                    match = re.search(r'基金代码[：:]\s*([0-9]{6})', line)
                    if match:
                        fund_market_code = match.group(1)
                
                # 提取申购金额 (例如：申购金额（小写）：60,000,000.00元)
                if '申购金额' in line and '小写' in line:
                    # 匹配括号（兼容中英文）、冒号、金额、去除元
                    match = re.search(r'申购金额[（(]小写[）)][：:]\s*([\d,]+\.?\d*)', line)
                    if match:
                        amount = match.group(1).replace(',', '')
            
            if fund_market_code and amount:
                results.append((product_name, fund_market_code, amount, '证达通基金'))
                
        else:
            # === 策略 B：汇总格式提取 (原有逻辑) ===
            # 策略：以“6位基金代码”为锚点寻找数据，兼容跨行
            for i, line in enumerate(lines):
                # 查找当前行中所有的6位数字代码
                code_iter = re.finditer(r'([0-9]{6})', line)
                
                for match in code_iter:
                    fund_market_code = match.group(1)
                    amount = ''
                    
                    # 情况1：金额在同一行（代码之后）
                    rest_of_line = line[match.end():]
                    amount_match = re.search(r'([\d,]+\.\d+)', rest_of_line)
                    if amount_match:
                        amount = amount_match.group(1).replace(',', '')
                    
                    # 情况2：金额在下一行（跨行）
                    elif i + 1 < len(lines):
                        next_line = lines[i+1].strip()
                        # 匹配下一行开头的金额
                        next_amount_match = re.search(r'^([\d,]+\.\d+)', next_line)
                        if next_amount_match:
                            amount = next_amount_match.group(1).replace(',', '')
                    
                    if fund_market_code and amount:
                        results.append((product_name, fund_market_code, amount, '证达通基金'))

        return results

    # 5. 遍历申购申请文件夹
    target_df = pd.DataFrame(columns=target_cols)
    target_path = os.path.join(folder_path, str(current_year), today_str, "1场外开基")
    log(f"正在查找受理单子目录: {target_path}", log_text)
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
        # 这一步是为了确保我们只关注包含受理单/申请单的路径，但现在 root 已经是完整路径
        # 我们应该检查 root 路径中是否包含 "受理"或"申请"
        if not ("受理" in root or "申请" in root):
             continue
        
        log(f"扫描目录: {root}", log_text)

        # 修改这一行，筛选出不包含"赎回"和"转换"的PDF文件
        pdf_files = [f for f in files if f.lower().endswith('.pdf') and "赎回" not in f and "超级" not in f and "转换" not in f and "分红方式" not in f and "分红设置" not in f and "失效" not in f]
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
                is_lide = any('利得基金' in l for l in lines[3:])
                is_changliang = any('长量基金' in l for l in lines[:2])
                is_yingmi = ('盈米' in file) or any('盈米' in l for l in lines[:3])
                is_pingan = any('行E通' in l for l in lines[5:])
                is_jiaohang = ('交e通' in file) or any('交通银行' in l for l in lines[:2])
                is_wangjin = ('网金' in file) or any('网金基金' in l for l in lines[5:]) #最后五行是否包含
                is_tengyuan = ('腾元' in file) or any('腾元基金' in l for l in lines[5:])
                is_hexun = any('和讯信息科技有限公司' in l for l in lines[3:]) 
                is_jd = ('肯特瑞基金' in file) or any('肯特瑞' in l for l in lines[:2])
                is_minsheng = ('民生同业e+' in file) or any('同业e+' in l for l in lines[2:])
                is_zhaoyingtong = any('招赢通' in l for l in lines[:2])
                is_ronglianchuang = any('融联创' in l for l in lines[8:])
                is_liantai = ('北极星' in file) or any('联泰' in l for l in lines[:2])
                is_jiyu = any('基煜基金' in l for l in lines[:2])
                is_ningboBank = (('宁波' in file and '北极星' not in file) or (any('宁波银行' in l for l in lines[15:]) and not any('联泰' in l for l in lines[:5])))
                is_guoxinjiali = any('国信嘉利基金' in l for l in lines[:2])
                is_panying = ('攀赢' in file) or any('攀赢' in l for l in lines[:2])
                is_zdt = any('证达通' in l for l in lines) and any('赎回交易（合计0笔，共计0.00份）' in l for l in lines) and not any('超级' in l for l in lines) 
                # 证达通判断：满足 汇总格式条件 OR 单笔格式条件
                # 汇总条件：含'证达通' + '赎回交易（合计0笔...' + 无'超级'
                # 单笔条件：含'证达通' + '申购受理单'
                is_zdt = (
                    (any('证达通' in l for l in lines) and any('赎回交易（合计0笔，共计0.00份）' in l for l in lines) and not any('超级' in l for l in lines)) 
                    or 
                    (any('证达通' in l for l in lines) and any('申购受理单' in l for l in lines))
                )

                if is_haomai:
                    product_name, fund_market_code, amount, platform = extract_haomai_fields(text, lines)
                    records = [(product_name, fund_market_code, amount, platform)]
                elif is_tiantian:
                    product_name, fund_market_code, amount, platform = extract_tiantian_fields(lines)
                    records = [(product_name, fund_market_code, amount, platform)]
                elif is_lide:
                    product_name, fund_market_code, amount, platform = extract_lide_fields(lines)
                    records = [(product_name, fund_market_code, amount, platform)]
                elif is_changliang:
                    product_name, fund_market_code, amount, platform = extract_changliang_fields(lines)
                    records = [(product_name, fund_market_code, amount, platform)]
                elif is_pingan:
                    product_name, fund_market_code, amount, platform = extract_pingan_fields(lines)
                    records = [(product_name, fund_market_code, amount, platform)]
                elif is_jiaohang:
                    product_name, fund_market_code, amount, platform = extract_jiaohang_fields(lines)
                    records = [(product_name, fund_market_code, amount, platform)]
                elif is_wangjin:
                    product_name, fund_market_code, amount, platform = extract_wangjin_fields(lines)
                    records = [(product_name, fund_market_code, amount, platform)]
                elif is_tengyuan:
                    product_name, fund_market_code, amount, platform = extract_tengyuan_fields(lines)
                    records = [(product_name, fund_market_code, amount, platform)]
                elif is_hexun:
                    product_name, fund_market_code, amount, platform = extract_hexun_fields(lines)
                    records = [(product_name, fund_market_code, amount, platform)]
                elif is_jd:
                    product_name, fund_market_code, amount, platform = extract_jd_fields(lines)
                    records = [(product_name, fund_market_code, amount, platform)]
                elif is_minsheng:
                    product_name, fund_market_code, amount, platform = extract_minsheng_fields(lines)
                    records = [(product_name, fund_market_code, amount, platform)]
                elif is_zhaoyingtong:
                    product_name, fund_market_code, amount, platform = extract_zhaoyingtong_fields(lines)
                    records = [(product_name, fund_market_code, amount, platform)]
                elif is_ronglianchuang:
                    product_name, fund_market_code, amount, platform = extract_ronglianchuang_fields(lines)
                    records = [(product_name, fund_market_code, amount, platform)]
                elif is_jiyu:
                    product_name, fund_market_code, amount, platform = extract_jiyu_fields(lines)
                    records = [(product_name, fund_market_code, amount, platform)]
                elif is_ningboBank:
                    product_name, fund_market_code, amount, platform = extract_ningboBank_fields(lines)
                    records = [(product_name, fund_market_code, amount, platform)]
                elif is_guoxinjiali:
                    product_name, fund_market_code, amount, platform = extract_guoxinjiali_fields(lines)
                    records = [(product_name, fund_market_code, amount, platform)]
                elif is_panying:
                    product_name, fund_market_code, amount, platform = extract_panying_fields(lines)
                    records = [(product_name, fund_market_code, amount, platform)]
                elif is_yingmi:
                    records = extract_yingmi_fields(lines)
                elif is_liantai:
                    records = extract_liantai_fields(lines)
                elif is_zdt:
                    records = extract_zdt_fields(lines)
                else:
                    continue

                for product_name, fund_market_code, amount, platform in records:
                    temp_df = pd.DataFrame([{
                        '产品名称': product_name,
                        '基金市场代码': fund_market_code,
                        '金额': amount,
                        '基金平台': platform
                    }])

                    #强制转换为数值型
                    temp_df['金额'] = pd.to_numeric(temp_df['金额'], errors='coerce').round(2)
                    # 映射账套编号
                    temp_df['账套编号'] = temp_df['产品名称'].map(product_code_dict)
                    # 填充其它字段
                    temp_df['交易市场'] = '国内银行间'
                    temp_df['业务类别'] = '基金申购申请'
                    temp_df['日期'] = today_str
                    temp_df['数量'] = ''
                    temp_df['手续费'] = ''
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
    output_file = os.path.join(output_folder, "【境内基金业务】申购申请.xls")
    try:
        with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
            target_df.to_excel(writer, sheet_name='Sheet1', index=False)
        log(f"已汇总输出到: {output_file}", log_text)
        return output_folder
    except Exception as e:
        log(f"写入Excel失败: {e}", log_text)
        return False