#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
快手订单对账与利润统计系统
基于 Flask 的网页版订单比对工具
"""

import os
import re
import uuid
from datetime import datetime
from flask import Flask, render_template, request, jsonify, send_file
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from werkzeug.utils import secure_filename

app = Flask(__name__)

# 配置
UPLOAD_FOLDER = 'uploads'
EXPORT_FOLDER = 'exports'
ALLOWED_EXTENSIONS = {'xlsx', 'xls', 'et'}
MAX_CONTENT_LENGTH = 100 * 1024 * 1024  # 100MB

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['EXPORT_FOLDER'] = EXPORT_FOLDER
app.config['MAX_CONTENT_LENGTH'] = MAX_CONTENT_LENGTH

# 确保目录存在
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(EXPORT_FOLDER, exist_ok=True)


def allowed_file(filename):
    """检查文件扩展名是否允许"""
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


def clean_order_id(order_id):
    """清洗订单号：去除空格、特殊字符，只保留数字和字母"""
    if pd.isna(order_id):
        return ''
    order_id = str(order_id).strip()
    # 去除所有空格和特殊字符，只保留数字和字母
    order_id = re.sub(r'[^\w]', '', order_id)
    return order_id.upper()


def detect_columns(df, is_official=True):
    """自动检测列名"""
    columns = df.columns.tolist()
    result = {}
    
    # 订单号可能的列名
    order_patterns = ['订单号', '订单编号', '订单ID', 'order', 'order_id', '订单']
    # 状态可能的列名
    status_patterns = ['状态', '订单状态', 'status', 'state']
    # 商品名可能的列名
    product_patterns = ['商品', '商品名称', '商品名', 'product', '商品标题', '标题']
    # 金额可能的列名
    amount_patterns = ['金额', '销售额', '实付金额', '支付金额', '总价', 'amount', 'price', '价格']
    # 成本可能的列名
    cost_patterns = ['成本', '成本价', '进货价', 'cost', '进价', '采购价']
    
    def find_column(patterns, columns):
        """查找匹配的列名"""
        for pattern in patterns:
            for col in columns:
                if pattern.lower() in col.lower():
                    return col
        return columns[0] if columns else None
    
    result['order_id'] = find_column(order_patterns, columns)
    result['status'] = find_column(status_patterns, columns) if is_official else None
    result['product'] = find_column(product_patterns, columns)
    result['amount'] = find_column(amount_patterns, columns)
    result['cost'] = find_column(cost_patterns, columns)
    
    return result, columns


def process_official_data(df, column_mapping):
    """处理官方订单数据"""
    try:
        # 提取需要的列
        order_col = column_mapping.get('order_id')
        status_col = column_mapping.get('status')
        product_col = column_mapping.get('product')
        amount_col = column_mapping.get('amount')
        cost_col = column_mapping.get('cost')
        
        if not all([order_col, status_col, product_col, amount_col]):
            raise ValueError("官方表缺少必要的列映射")
        
        # 创建新 DataFrame
        processed = pd.DataFrame()
        processed['订单号'] = df[order_col].apply(clean_order_id)
        processed['原始状态'] = df[status_col].astype(str)
        processed['商品名称'] = df[product_col].astype(str)
        processed['销售额'] = pd.to_numeric(df[amount_col], errors='coerce').fillna(0)
        
        # 如果有成本列，使用它，否则为0
        if cost_col and cost_col in df.columns:
            processed['成本'] = pd.to_numeric(df[cost_col], errors='coerce').fillna(0)
        else:
            processed['成本'] = 0
        
        # 过滤空订单号
        processed = processed[processed['订单号'] != '']
        
        # 去重（保留第一条）
        processed = processed.drop_duplicates(subset=['订单号'], keep='first')
        
        # 状态过滤：只保留完全等于"交易成功"或"已发货"的订单
        valid_status = ['交易成功', '已发货']
        processed = processed[processed['原始状态'].isin(valid_status)]
        
        return processed
    except Exception as e:
        raise Exception(f"处理官方数据时出错: {str(e)}")


def process_customer_data(df, column_mapping):
    """处理客服统计数据"""
    try:
        # 提取需要的列
        order_col = column_mapping.get('order_id')
        product_col = column_mapping.get('product')
        amount_col = column_mapping.get('amount')
        cost_col = column_mapping.get('cost')
        
        if not all([order_col, product_col, amount_col]):
            raise ValueError("客服表缺少必要的列映射")
        
        # 创建新 DataFrame
        processed = pd.DataFrame()
        processed['订单号'] = df[order_col].apply(clean_order_id)
        processed['商品名称'] = df[product_col].astype(str)
        processed['销售额'] = pd.to_numeric(df[amount_col], errors='coerce').fillna(0)
        
        # 成本列必须存在
        if cost_col and cost_col in df.columns:
            processed['成本'] = pd.to_numeric(df[cost_col], errors='coerce').fillna(0)
        else:
            processed['成本'] = 0
        
        # 过滤空订单号
        processed = processed[processed['订单号'] != '']
        
        # 去重检测（记录重复的）
        duplicates = processed[processed.duplicated(subset=['订单号'], keep=False)]
        duplicate_order_ids = duplicates['订单号'].unique().tolist()
        
        # 去重（保留第一条）
        processed = processed.drop_duplicates(subset=['订单号'], keep='first')
        
        return processed, duplicate_order_ids
    except Exception as e:
        raise Exception(f"处理客服数据时出错: {str(e)}")


def compare_orders(official_df, customer_df):
    """对比订单数据"""
    # 获取订单号集合
    official_orders = set(official_df['订单号'].tolist())
    customer_orders = set(customer_df['订单号'].tolist())
    
    # 匹配成功的订单
    matched_orders = official_orders & customer_orders
    
    # 官方有但客服没有（漏单）
    missing_in_customer = official_orders - customer_orders
    
    # 客服有但官方没有
    extra_in_customer = customer_orders - official_orders
    
    return {
        'matched': matched_orders,
        'missing': missing_in_customer,
        'extra': extra_in_customer
    }


def generate_result_data(official_df, customer_df, comparison_result):
    """生成结果数据"""
    matched_orders = comparison_result['matched']
    missing_orders = comparison_result['missing']
    extra_orders = comparison_result['extra']
    
    result_rows = []
    
    # 1. 匹配成功的订单
    for order_id in matched_orders:
        official_row = official_df[official_df['订单号'] == order_id].iloc[0]
        customer_row = customer_df[customer_df['订单号'] == order_id].iloc[0]
        
        # 优先使用客服表的成本，如果没有则使用官方表
        sales_amount = customer_row['销售额'] if customer_row['销售额'] > 0 else official_row['销售额']
        cost_amount = customer_row['成本'] if customer_row['成本'] > 0 else official_row['成本']
        profit = sales_amount - cost_amount
        
        result_rows.append({
            '订单号': order_id,
            '商品名称': customer_row['商品名称'],
            '销售额': sales_amount,
            '成本': cost_amount,
            '利润': profit,
            '状态': '正常',
            '备注': ''
        })
    
    # 2. 官方有但客服没有的订单（漏单）
    for order_id in missing_orders:
        official_row = official_df[official_df['订单号'] == order_id].iloc[0]
        result_rows.append({
            '订单号': order_id,
            '商品名称': official_row['商品名称'],
            '销售额': official_row['销售额'],
            '成本': official_row['成本'],
            '利润': official_row['销售额'] - official_row['成本'],
            '状态': '异常',
            '备注': '【客服漏记】'
        })
    
    # 3. 客服有但官方没有的订单
    for order_id in extra_orders:
        customer_row = customer_df[customer_df['订单号'] == order_id].iloc[0]
        result_rows.append({
            '订单号': order_id,
            '商品名称': customer_row['商品名称'],
            '销售额': customer_row['销售额'],
            '成本': customer_row['成本'],
            '利润': customer_row['销售额'] - customer_row['成本'],
            '状态': '异常',
            '备注': '【客服有官方没有】'
        })
    
    return pd.DataFrame(result_rows)


def generate_summary(result_df):
    """生成汇总统计"""
    normal_df = result_df[result_df['状态'] == '正常']
    abnormal_df = result_df[result_df['状态'] == '异常']
    
    summary = {
        'total_sales': result_df['销售额'].sum(),
        'total_cost': result_df['成本'].sum(),
        'total_profit': result_df['利润'].sum(),
        'total_orders': len(result_df),
        'normal_orders': len(normal_df),
        'abnormal_orders': len(abnormal_df),
        'missing_orders': len(result_df[result_df['备注'] == '【客服漏记】']),
        'extra_orders': len(result_df[result_df['备注'] == '【客服有官方没有】'])
    }
    
    return summary


def export_to_excel(result_df, summary, filename):
    """导出结果到 Excel"""
    try:
        # 确保数值列是数字类型
        for col in ['销售额', '成本', '利润']:
            if col in result_df.columns:
                result_df[col] = pd.to_numeric(result_df[col], errors='coerce').fillna(0)
        
        # 重置索引并添加序号列
        result_df = result_df.reset_index(drop=True)
        result_df.insert(0, '序号', range(1, len(result_df) + 1))
        
        # 创建工作簿
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "对账结果"
        
        # 设置列宽
        column_widths = {
            'A': 8,   # 序号
            'B': 25,  # 订单号
            'C': 40,  # 商品名称
            'D': 12,  # 销售额
            'E': 12,  # 成本
            'F': 12,  # 利润
            'G': 12,  # 状态
            'H': 20   # 备注
        }
        for col, width in column_widths.items():
            ws.column_dimensions[col].width = width
        
        # 定义样式
        header_fill = PatternFill(start_color='4472C4', end_color='4472C4', fill_type='solid')
        header_font = Font(name='微软雅黑', size=11, bold=True, color='FFFFFF')
        normal_font = Font(name='微软雅黑', size=10)
        red_font = Font(name='微软雅黑', size=10, color='FF0000')
        abnormal_fill = PatternFill(start_color='FFF2CC', end_color='FFF2CC', fill_type='solid')
        
        # 边框样式
        thin_border = Border(
            left=Side(style='thin'),
            right=Side(style='thin'),
            top=Side(style='thin'),
            bottom=Side(style='thin')
        )
        
        # 写入表头
        headers = ['序号', '订单号', '商品名称', '销售额', '成本', '利润', '状态', '备注']
        for col_idx, header in enumerate(headers, 1):
            cell = ws.cell(row=1, column=col_idx, value=header)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal='center', vertical='center')
            cell.border = thin_border
        
        # 写入数据 - 按固定列顺序写入
        col_order = ['序号', '订单号', '商品名称', '销售额', '成本', '利润', '状态', '备注']
        
        for row_idx in range(len(result_df)):
            # 获取当前行数据为字典
            row_dict = result_df.iloc[row_idx].to_dict()
            
            is_abnormal = str(row_dict.get('状态', '')) == '异常'
            is_loss = float(row_dict.get('利润', 0)) < 0
            
            # 按固定顺序写入每一列
            for col_idx, col_name in enumerate(col_order, 1):
                value = row_dict.get(col_name, '')
                cell = ws.cell(row=row_idx + 2, column=col_idx)
                cell.border = thin_border
                
                # 设置值和格式
                if col_name == '序号':
                    cell.value = int(value) if value else row_idx + 1
                    cell.alignment = Alignment(horizontal='center', vertical='center')
                    cell.font = normal_font
                elif col_name == '订单号':
                    cell.value = str(value)
                    cell.alignment = Alignment(horizontal='left', vertical='center')
                    cell.font = normal_font
                elif col_name == '商品名称':
                    cell.value = str(value)
                    cell.alignment = Alignment(horizontal='left', vertical='center')
                    cell.font = normal_font
                elif col_name in ['销售额', '成本', '利润']:
                    cell.value = float(value) if value else 0
                    cell.number_format = '¥#,##0.00'
                    cell.alignment = Alignment(horizontal='center', vertical='center')
                    if col_name == '利润' and is_loss:
                        cell.font = red_font
                    else:
                        cell.font = normal_font
                elif col_name == '状态':
                    cell.value = str(value)
                    cell.alignment = Alignment(horizontal='center', vertical='center')
                    cell.font = normal_font
                elif col_name == '备注':
                    cell.value = str(value)
                    cell.alignment = Alignment(horizontal='left', vertical='center')
                    cell.font = normal_font
                
                # 异常行高亮
                if is_abnormal and col_name not in ['销售额', '成本', '利润']:
                    cell.fill = abnormal_fill
                elif is_abnormal and col_name == '利润' and not is_loss:
                    cell.fill = abnormal_fill
        
        # 添加空行
        current_row = len(result_df) + 3
        
        # 写入汇总统计
        summary_data = [
            ['汇总统计', '', '', '', '', '', '', ''],
            ['项目', '数值', '', '', '', '', '', ''],
            ['总销售额', summary['total_sales'], '', '', '', '', '', ''],
            ['总成本', summary['total_cost'], '', '', '', '', '', ''],
            ['总利润', summary['total_profit'], '', '', '', '', '', ''],
            ['订单总数', summary['total_orders'], '', '', '', '', '', ''],
            ['正常订单', summary['normal_orders'], '', '', '', '', '', ''],
            ['异常订单', summary['abnormal_orders'], '', '', '', '', '', ''],
            ['客服漏记', summary['missing_orders'], '', '', '', '', '', ''],
            ['客服有官方没有', summary['extra_orders'], '', '', '', '', '', '']
        ]
        
        summary_fill = PatternFill(start_color='E7E6E6', end_color='E7E6E6', fill_type='solid')
        summary_font = Font(name='微软雅黑', size=10, bold=True)
        
        for row_data in summary_data:
            for col_idx, value in enumerate(row_data, 1):
                cell = ws.cell(row=current_row, column=col_idx, value=value)
                cell.border = thin_border
                cell.alignment = Alignment(horizontal='left' if col_idx == 1 else 'center', vertical='center')
                
                if current_row == len(result_df) + 3:  # 汇总统计标题行
                    cell.fill = header_fill
                    cell.font = header_font
                elif current_row == len(result_df) + 4:  # 项目行
                    cell.fill = summary_fill
                    cell.font = summary_font
                else:
                    cell.font = normal_font
                    if col_idx == 2 and isinstance(value, (int, float)):  # 数值列
                        cell.number_format = '¥#,##0.00' if current_row <= len(result_df) + 7 else '#,##0'
            
            current_row += 1
        
        # 设置行高
        ws.row_dimensions[1].height = 25
        for row in range(2, current_row):
            ws.row_dimensions[row].height = 20
        
        # 保存文件
        filepath = os.path.join(EXPORT_FOLDER, filename)
        wb.save(filepath)
        return filepath
    except Exception as e:
        raise Exception(f"导出 Excel 时出错: {str(e)}")


# ============== API 路由 ==============

@app.route('/')
def index():
    """主页"""
    return render_template('index.html')


@app.route('/api/upload', methods=['POST'])
def upload_file():
    """上传文件并检测列"""
    try:
        if 'file' not in request.files:
            return jsonify({'success': False, 'error': '没有文件'})
        
        file = request.files['file']
        file_type = request.form.get('type', 'official')  # official 或 customer
        
        if file.filename == '':
            return jsonify({'success': False, 'error': '文件名为空'})
        
        if not allowed_file(file.filename):
            return jsonify({'success': False, 'error': '不支持的文件格式，请上传 .xlsx 或 .xls 文件'})
        
        # 保存文件
        unique_id = str(uuid.uuid4())[:8]
        filename = f"{file_type}_{unique_id}_{secure_filename(file.filename)}"
        filepath = os.path.join(UPLOAD_FOLDER, filename)
        file.save(filepath)
        
        # 读取文件检测列
        try:
            if filename.endswith('.et'):
                # WPS .et 格式需要特殊处理，提示用户转换
                return jsonify({
                    'success': False, 
                    'error': '检测到 WPS .et 格式，请另存为 .xlsx 格式后重新上传'
                })
            
            df = pd.read_excel(filepath, engine='openpyxl')
            
            if df.empty:
                return jsonify({'success': False, 'error': '文件为空'})
            
            # 自动检测列
            detected_columns, all_columns = detect_columns(df, is_official=(file_type == 'official'))
            
            return jsonify({
                'success': True,
                'filename': filename,
                'filepath': filepath,
                'columns': all_columns,
                'detected': detected_columns,
                'row_count': len(df)
            })
            
        except Exception as e:
            # 删除上传的文件
            if os.path.exists(filepath):
                os.remove(filepath)
            return jsonify({'success': False, 'error': f'读取文件失败: {str(e)}'})
            
    except Exception as e:
        return jsonify({'success': False, 'error': f'上传失败: {str(e)}'})


@app.route('/api/compare', methods=['POST'])
def compare():
    """执行对比"""
    try:
        data = request.json
        
        official_file = data.get('official_file')
        customer_file = data.get('customer_file')
        official_mapping = data.get('official_mapping')
        customer_mapping = data.get('customer_mapping')
        
        if not all([official_file, customer_file, official_mapping, customer_mapping]):
            return jsonify({'success': False, 'error': '缺少必要的参数'})
        
        # 读取文件
        official_df = pd.read_excel(official_file, engine='openpyxl')
        customer_df = pd.read_excel(customer_file, engine='openpyxl')
        
        # 处理数据
        official_processed = process_official_data(official_df, official_mapping)
        customer_processed, duplicate_orders = process_customer_data(customer_df, customer_mapping)
        
        # 对比订单
        comparison_result = compare_orders(official_processed, customer_processed)
        
        # 生成结果
        result_df = generate_result_data(official_processed, customer_processed, comparison_result)
        
        # 生成汇总
        summary = generate_summary(result_df)
        summary['duplicate_orders'] = len(duplicate_orders)
        
        # 转换为前端格式
        result_data = result_df.to_dict('records')
        
        return jsonify({
            'success': True,
            'data': result_data,
            'summary': summary,
            'duplicates': duplicate_orders
        })
        
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)})


@app.route('/api/export', methods=['POST'])
def export():
    """导出 Excel"""
    try:
        data = request.json
        result_data = data.get('data', [])
        summary = data.get('summary', {})
        
        if not result_data:
            return jsonify({'success': False, 'error': '没有数据可导出'})
        
        result_df = pd.DataFrame(result_data)
        
        # 生成文件名
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        filename = f"对账结果_{timestamp}.xlsx"
        
        # 导出
        filepath = export_to_excel(result_df, summary, filename)
        
        return jsonify({
            'success': True,
            'filename': filename,
            'download_url': f'/api/download/{filename}'
        })
        
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)})


@app.route('/api/download/<filename>')
def download(filename):
    """下载文件"""
    try:
        filepath = os.path.join(EXPORT_FOLDER, filename)
        if not os.path.exists(filepath):
            return jsonify({'success': False, 'error': '文件不存在'})
        
        return send_file(filepath, as_attachment=True, download_name=filename)
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)})


@app.route('/api/cleanup', methods=['POST'])
def cleanup():
    """清理临时文件"""
    try:
        data = request.json
        files = data.get('files', [])
        
        for filepath in files:
            if os.path.exists(filepath):
                os.remove(filepath)
        
        return jsonify({'success': True})
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)})


if __name__ == '__main__':
    print("=" * 50)
    print("快手订单对账与利润统计系统")
    print("=" * 50)
    print("访问地址: http://localhost:5000")
    print("=" * 50)
    app.run(host='0.0.0.0', port=5000, debug=True)
