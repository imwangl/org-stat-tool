import os
import zipfile
import shutil
import openpyxl
import xlrd
from flask import Flask, render_template, request, send_file, jsonify
from collections import defaultdict
from datetime import datetime, timedelta
import uuid

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = '/tmp/org_stat_uploads'
app.config['OUTPUT_FOLDER'] = '/tmp/org_stat_outputs'

os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
os.makedirs(app.config['OUTPUT_FOLDER'], exist_ok=True)


def find_file(data_dir, keyword):
    """模糊匹配文件名"""
    for f in os.listdir(data_dir):
        if f.endswith(('.xlsx', '.xls')) and keyword in f:
            return f"{data_dir}/{f}"
    return None

def process_org_stat(data_dir, output_dir, user_time):
    """执行机构数统计"""
    
    # 动态查找文件
    template_file = find_file(data_dir, "模板")
    template2_file = find_file(data_dir, "模版2")
    corp_file = find_file(data_dir, "企业用户数")
    org_file = find_file(data_dir, "机构用户")
    precharge_file = find_file(data_dir, "预充值")
    history_file = find_file(data_dir, "历史客户")
    market_file = find_file(data_dir, "客户数据汇总表-市场部")
    strategy_file = find_file(data_dir, "客户数据汇总表-战略增长中心")
    
    if not template_file:
        return "错误: 未找到模板文件"
    
    # 复制模板
    template_path = f"{output_dir}/模板.xlsx"
    shutil.copy(template_file, template_path)
    
    # 修改A1单元格标题，添加时间
    wb = openpyxl.load_workbook(template_path)
    ws = wb.active
    ws.cell(1, 1).value = f"客户数据汇总表（{user_time}）未与财务核对版"
    wb.save(template_path)
    
    # 步骤1-2: 财务表数据
    finance_data = []
    for f in [market_file, strategy_file] if f:
        try:
            wb = xlrd.open_workbook(f)
            ws = wb.sheet_by_index(0)
            for row in range(3, ws.nrows):
                name = ws.cell_value(row, 3)
                if name:
                    finance_data.append({
                        '部门': ws.cell_value(row, 1),
                        '组别': ws.cell_value(row, 2),
                        '姓名': name,
                        '发生月份': user_time
                    })
        except Exception as e:
            print(f"读取{f}失败: {e}")
    
    wb = openpyxl.load_workbook(template_path)
    ws = wb.active
    from openpyxl.styles import Alignment
    
    for i, data in enumerate(finance_data):
        row_num = 4 + i
        ws.cell(row_num, 1).value = i + 1
        ws.cell(row_num, 2).value = data['部门']
        ws.cell(row_num, 3).value = data['组别']
        ws.cell(row_num, 4).value = data['姓名']
        ws.cell(row_num, 5).value = data['发生月份']
        ws.cell(row_num, 12).value = data['发生月份']
        ws.row_dimensions[row_num].height = 15
    
    for row in range(4, 4 + len(finance_data)):
        for col in range(1, 26):
            ws.cell(row, col).alignment = Alignment(horizontal='left', vertical='center')
    
    wb.save(template_path)
    
    # 步骤5: 企业用户数
    corp_data = {}
    ws = wb.active
    for row in range(3, ws.max_row + 1):
        name = ws.cell(row, 1).value
        count = ws.cell(row, 4).value
        if name and count is not None:
            corp_data[str(name).strip()] = count
    
    wb = openpyxl.load_workbook(template_path)
    ws = wb.active
    
    for row in range(4, 4 + len(finance_data)):
        name = ws.cell(row, 4).value
        if name and str(name).strip() in corp_data:
            val = corp_data[str(name).strip()]
            ws.cell(row, 6).value = int(val)
        else:
            ws.cell(row, 6).value = 0
        ws.cell(row, 6).number_format = '0'
    
    wb.save(template_path)
    
    # 结果a
    result_a = f"{output_dir}/结果a.xlsx"
    wb = openpyxl.load_workbook(result_a)
    ws = wb.active
    ws.cell(1, 17).value = "认证企业"
    ws.cell(1, 18).value = "认证企业2"
    ws.cell(1, 19).value = "企业名称"
    for row in range(2, ws.max_row + 1):
        ws.cell(row, 17).value = ws.cell(row, 9).value

    ws_p = wb_p.active
    for i in range(2, min(ws_p.max_row + 1, ws.max_row + 1)):
        ws.cell(i, 18).value = ws_p.cell(i, 8).value

    ws_h = wb_h.active
    for i in range(2, min(ws_h.max_row + 1, ws.max_row + 1)):
        company = ws_h.cell(i, 5).value
        if company:
            ws.cell(i, 19).value = company
    wb.save(result_a)
    
    # 结果b
    result_b = f"{output_dir}/结果b.xlsx"
    shutil.copy(result_a, result_b)
    wb = openpyxl.load_workbook(result_b)
    ws = wb.active
    col_values = defaultdict(list)
    for row in range(2, ws.max_row + 1):
        for col in [17, 18, 19]:
            val = ws.cell(row, col).value
            if val and str(val).strip():
                col_values[str(val).strip()].append(row)
    rows_to_delete = set()
    for row in range(2, ws.max_row + 1):
        val = ws.cell(row, 17).value
        if val and str(val).strip() and len(col_values[str(val).strip()]) > 1:
            rows_to_delete.add(row)
    for row in sorted(rows_to_delete, reverse=True):
        ws.delete_rows(row)
    wb.save(result_b)
    
    # 结果c
    result_c = f"{output_dir}/结果c.xlsx"
    shutil.copy(result_b, result_c)
    wb.save(result_c)
    
    # 步骤15-17: 补贴统计
    wb = openpyxl.load_workbook(result_c)
    ws = wb.active
    stats = defaultdict(lambda: {'-0.5': 0, '-0.1': 0})
    for row in range(2, ws.max_row + 1):
        name = ws.cell(row, 2).value
        ratio = ws.cell(row, 5).value
        if name and ratio:
            r = str(ratio).strip()
            if r in ['-0.5', '-0.50']:
                stats[str(name).strip()]['-0.5'] += 1
            elif r in ['-0.1', '-0.10']:
                stats[str(name).strip()]['-0.1'] += 1
    
    new_wb = openpyxl.Workbook()
    new_ws = new_wb.active
    new_ws.cell(1, 1).value = "用户名称"
    new_ws.cell(1, 2).value = "-0.5计数"
    new_ws.cell(1, 3).value = "-0.1计数"
    for i, (name, counts) in enumerate(sorted(stats.items())):
        new_ws.cell(i+2, 1).value = name
        new_ws.cell(i+2, 2).value = counts['-0.5']
        new_ws.cell(i+2, 3).value = counts['-0.1']
    subsidy_path = f"{output_dir}/补贴比例统计.xlsx"
    new_wb.save(subsidy_path)
    
    # 步骤16-17
    wb_sub = openpyxl.load_workbook(subsidy_path)
    ws_sub = wb_sub.active
    wb = openpyxl.load_workbook(template_path)
    ws = wb.active
    all_names = [(row, str(ws.cell(row, 4).value).strip()) for row in range(4, 4 + len(finance_data)) if ws.cell(row, 4).value]
    subsidy_50 = {str(ws_sub.cell(r, 1).value).strip(): ws_sub.cell(r, 2).value or 0 for r in range(2, ws_sub.max_row + 1) if ws_sub.cell(r, 1).value}
    subsidy_10 = {str(ws_sub.cell(r, 1).value).strip(): ws_sub.cell(r, 3).value or 0 for r in range(2, ws_sub.max_row + 1) if ws_sub.cell(r, 1).value}
    for row, name in all_names:
        ws.cell(row, 18).value = subsidy_50.get(name, 0)
        ws.cell(row, 19).value = subsidy_10.get(name, 0)
    
    ws_h = wb_h.active
    name_counts = {}
    for row in range(2, ws_h.max_row + 1):
        name = ws_h.cell(row, 2).value
        if name:
            name_counts[str(name).strip()] = name_counts.get(str(name).strip(), 0) + 1
    for row, name in all_names:
        ws.cell(row, 22).value = name_counts.get(name, 0)
    wb.save(template_path)
    
    # 保存主输出
    final_output = f"{output_dir}/未与财务核对版本.xlsx"
    shutil.copy(template_path, final_output)
    
    # 生成与财务核对版
    template2_path = template2_file
    final_verify_path = f"{output_dir}/与财务核对版（{user_time}）.xlsx"
    shutil.copy(template2_path, final_verify_path)
    
    # 修改与财务核对版A1
    wb2 = openpyxl.load_workbook(final_verify_path)
    ws2 = wb2.active
    ws2.cell(1, 1).value = f"客户数据汇总表（{user_time}）与财务核对版"
    wb2.save(final_verify_path)
    
    # 读取源数据
    source_wb = openpyxl.load_workbook(final_output)
    source_ws = source_wb.active
    
    wb = openpyxl.load_workbook(final_verify_path)
    ws = wb.active
    
    # 复制1-6列和12列
    for row in range(4, 4 + len(finance_data)):
        for col in [1, 2, 3, 4, 5, 6, 12]:
            ws.cell(row, col).value = source_ws.cell(row, col).value
    
    # 复制R、S、V列
    for row in range(4, 4 + len(finance_data)):
        ws.cell(row, 20).value = source_ws.cell(row, 18).value
        ws.cell(row, 21).value = source_ws.cell(row, 19).value
        ws.cell(row, 24).value = source_ws.cell(row, 22).value
    
    # 财务表P列匹配
    finance_p_data = {}
    for f in [market_file, strategy_file] if f:
        wb_f = xlrd.open_workbook(f)
        ws_f = wb_f.sheet_by_index(0)
        for row in range(3, ws_f.nrows):
            name = ws_f.cell_value(row, 3)
            p_val = ws_f.cell_value(row, 15)
            if name:
                finance_p_data[str(name).strip()] = p_val
    
    for row in range(4, 4 + len(finance_data)):
        name = ws.cell(row, 4).value
        if name and str(name).strip() in finance_p_data:
            ws.cell(row, 16).value = finance_p_data[str(name).strip()]
    
    # 财务表R列匹配
    finance_r_data = {}
    for f in [market_file, strategy_file] if f:
        wb_f = xlrd.open_workbook(f)
        ws_f = wb_f.sheet_by_index(0)
        for row in range(3, ws_f.nrows):
            name = ws_f.cell_value(row, 3)
            r_val = ws_f.cell_value(row, 17)
            if name:
                finance_r_data[str(name).strip()] = r_val
    
    for row in range(4, 4 + len(finance_data)):
        name = ws.cell(row, 4).value
        if name and str(name).strip() in finance_r_data:
            ws.cell(row, 19).value = finance_r_data[str(name).strip()]
    
    # 新增步骤: 列O = 列P + 列R + 列X
    for row in range(4, 4 + len(finance_data)):
        p_val = ws.cell(row, 16).value or 0
        r_val = ws.cell(row, 19).value or 0
        x_val = ws.cell(row, 24).value or 0
        try:
            ws.cell(row, 15).value = float(p_val) + float(r_val) + float(x_val)
        except:
            ws.cell(row, 15).value = 0
    
    # 新增步骤: 列R = 列S + 列T
    for row in range(4, 4 + len(finance_data)):
        s_val = ws.cell(row, 19).value or 0
        t_val = ws.cell(row, 20).value or 0
        try:
            ws.cell(row, 18).value = float(s_val) + float(t_val)
        except:
            ws.cell(row, 18).value = 0
    
    # 新增步骤: 列M = 列O
    for row in range(4, 4 + len(finance_data)):
        o_val = ws.cell(row, 15).value or 0
        ws.cell(row, 13).value = o_val
    
    wb.save(final_verify_path)
    
    return final_output, final_verify_path

@app.route('/')
def index():
    return render_template('index.html')


@app.route('/upload_single', methods=['POST'])
def upload_single():
    """单文件上传处理 - 11个独立文件上传"""
    user_time = request.form.get('time', '3月')
    
    # 创建任务目录
    task_id = str(uuid.uuid4())
    task_dir = os.path.join(app.config['UPLOAD_FOLDER'], task_id)
    output_dir = os.path.join(app.config['OUTPUT_FOLDER'], task_id)
    os.makedirs(task_dir, exist_ok=True)
    os.makedirs(output_dir, exist_ok=True)
    
    # 文件名映射 (表单字段名 -> 期望的文件名)
    # 注意：process_org_stat 用 find_file 模糊匹配，所以文件名包含关键词即可
    file_map = {
        'org_user': '机构用户.xlsx',        # 机构用户
        'market': '客户数据市场部.xlsx',    # 客户数据汇总表-市场部
        'strategy': '客户数据战略中心.xlsx', # 客户数据汇总表-战略增长中心
        'history': '历史客户.xlsx',         # 历史客户
        'template1': '模板.xlsx',          # 未与财务核对模板
        'template2': '模版2.xlsx',         # 与财务核对模板
        'company': '企业用户数.xlsx',       # 企业用户数
        'all_history': '历史客户_全部.xlsx', # 全部历史客户
        'prepaid': '预充值.xlsx',           # 预充值
        'monthly': '月结用户.xlsx',         # 月结用户
    }
    
    # 保存上传的文件
    for key, filename in file_map.items():
        f = request.files.get(key)
        if f and f.filename:
            save_name = filename
            f.save(os.path.join(task_dir, save_name))
            print(f"保存文件: {save_name}")
    
    # 处理
    try:
        final_output, final_verify = process_org_stat(task_dir, output_dir, user_time)
        return jsonify({
            'success': True,
            'task_id': task_id,
            'files': [
                {'name': '未与财务核对版本.xlsx', 'path': final_output},
                {'name': '与财务核对版.xlsx', 'path': final_verify},
                {'name': '补贴比例统计.xlsx', 'path': f"{output_dir}/补贴比例统计.xlsx"}
            ]
        })
    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500

@app.route('/upload', methods=['POST'])
def upload():
    if 'files' not in request.files:
        return jsonify({'error': '请上传Excel文件'}), 400
    
    files = request.files.getlist('files')
    user_time = request.form.get('time', '3月')
    
    if not files or files[0].filename == '':
        return jsonify({'error': '未选择文件'}), 400
    
    # 创建任务目录
    task_id = str(uuid.uuid4())
    task_dir = os.path.join(app.config['UPLOAD_FOLDER'], task_id)
    output_dir = os.path.join(app.config['OUTPUT_FOLDER'], task_id)
    os.makedirs(task_dir, exist_ok=True)
    os.makedirs(output_dir, exist_ok=True)
    
    # 保存上传的文件
    for f in files:
        f.save(os.path.join(task_dir, f.filename))
    
    # 找到数据目录
    data_dir = None
    for root, dirs, files_list in os.walk(task_dir):
        xlsx_files = [f for f in files_list if not f.startswith('._') and (f.endswith('.xlsx') or f.endswith('.xls'))]
        if xlsx_files:
            data_dir = root
            print(f"数据目录: {root}, 文件: {xlsx_files}")
            break
    
    if not data_dir:
        return jsonify({'error': '未找到Excel文件'}), 400
    
    # 处理
    try:
        final_output, final_verify = process_org_stat(data_dir, output_dir, user_time)
        return jsonify({
            'success': True,
            'task_id': task_id,
            'files': [
                {'name': '未与财务核对版本.xlsx', 'path': final_output},
                {'name': '与财务核对版.xlsx', 'path': final_verify},
                {'name': '补贴比例统计.xlsx', 'path': f"{output_dir}/补贴比例统计.xlsx"}
            ]
        })
    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({'error': str(e)}), 500

@app.route('/download/<filename>')
def download(filename):
    for root, dirs, files in os.walk(app.config['OUTPUT_FOLDER']):
        if filename in files:
            return send_file(os.path.join(root, filename), as_attachment=True)
    return "文件未找到", 404

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)