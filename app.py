# 版本1 app.py
# 只适用于为同一专业或同一计算规则的学生设定计算规则

import os
import pandas as pd
from flask import Flask, request, render_template, jsonify, send_from_directory
import time

# --- 配置 ---
# 创建 uploads 和 output 文件夹（如果不存在）
if not os.path.exists('uploads'):
    os.makedirs('uploads')
if not os.path.exists('output'):
    os.makedirs('output')

UPLOAD_FOLDER = 'uploads/'
OUTPUT_FOLDER = 'output/'
ALLOWED_EXTENSIONS = {'xlsx'}

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['OUTPUT_FOLDER'] = OUTPUT_FOLDER
# 增加文件缓存时间，防止浏览器缓存旧文件
app.config['SEND_FILE_MAX_AGE_DEFAULT'] = 0 

# --- 辅助函数 ---
def allowed_file(filename):
    """检查文件扩展名是否合法"""
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def perform_calculation(filepath, rules):
    """
    根据动态规则执行加权平均计算
    
    参数:
    filepath (str): 原始Excel文件路径
    rules (dict): 包含课程名和对应学分（权重）的字典
    
    返回:
    pd.DataFrame: 包含结果的DataFrame
    """
    df = pd.read_excel(filepath)
    
    def calculate_weighted_average(student_row):
        weighted_sum = 0
        total_credits = 0
        
        for course, credit in rules.items():
            # 检查学生是否有该课程的成绩，且成绩是数字
            if course in student_row and pd.to_numeric(student_row[course], errors='coerce') == student_row[course]:
                weighted_sum += student_row[course] * credit
                total_credits += credit
        
        if total_credits == 0:
            return 0.0
            
        return round(weighted_sum / total_credits, 2)

    df['加权平均分'] = df.apply(calculate_weighted_average, axis=1)
    
    # 筛选输出列，这里假设原始文件有'学号'和'姓名'列
    output_cols = []
    if '学号' in df.columns: output_cols.append('学号')
    if '姓名' in df.columns: output_cols.append('姓名')
    output_cols.append('加权平均分')
    # 输出选择的课程
    for course in rules.keys():
        if course in df.columns:
            output_cols.append(course)

    return df[output_cols]

# --- 路由（API接口） ---

@app.route('/')
def index():
    """渲染主页面"""
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    """处理文件上传，并返回课程列表"""
    if 'file' not in request.files:
        return jsonify({'error': 'No file part'}), 400
    file = request.files['file']
    if file.filename == '':
        return jsonify({'error': 'No selected file'}), 400
    if file and allowed_file(file.filename):
        # 使用时间戳确保文件名唯一
        timestamp = str(int(time.time()))
        filename = timestamp + "_" + file.filename
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(filepath)
        
        # 使用 pandas 读取列名（课程名）
        try:
            df_headers = pd.read_excel(filepath, nrows=0) # nrows=0 只读表头，效率高
            # 假设学号、姓名等非课程列需要排除
            course_columns = [col for col in df_headers.columns if col not in ['学号', '姓名', '专业']]
            return jsonify({'courses': course_columns, 'filename': filename})
        except Exception as e:
            return jsonify({'error': f'Error processing file: {e}'}), 500
            
    return jsonify({'error': 'File type not allowed'}), 400

@app.route('/calculate', methods=['POST'])
def calculate():
    """接收计算规则并执行计算"""
    data = request.get_json()
    filename = data.get('filename')
    rules_list = data.get('rules') # rules 是一个列表，如 [{'course': '语文', 'credit': 4.0}, ...]
    
    if not filename or not rules_list:
        return jsonify({'error': 'Missing filename or rules'}), 400
        
    # 将前端发来的列表转换为 {'课程名': 学分} 的字典
    rules_dict = {item['course']: float(item['credit']) for item in rules_list}

    input_filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
    
    try:
        result_df = perform_calculation(input_filepath, rules_dict)
        
        output_filename = 'result_' + filename
        output_filepath = os.path.join(app.config['OUTPUT_FOLDER'], output_filename)
        result_df.to_excel(output_filepath, index=False)
        
        # 返回可供下载的文件名
        return jsonify({'download_file': output_filename})
    except Exception as e:
        return jsonify({'error': f'Calculation failed: {e}'}), 500

@app.route('/download/<filename>')
def download_file(filename):
    """提供文件下载"""
    return send_from_directory(app.config['OUTPUT_FOLDER'], filename, as_attachment=True)


if __name__ == '__main__':
    app.run(debug=True) # debug=True 模式方便开发调试
