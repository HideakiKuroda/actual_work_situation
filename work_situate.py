from flask import Flask, request, render_template, flash, session,redirect, url_for, jsonify
from openpyxl import Workbook, load_workbook
from utility import fill_dates_and_weekdays,generate_new_filename,show_dailyreports,copy_data_to_work_record,copy_dates_to_new_sheet
from datetime import datetime
import os
import logging

app = Flask(__name__)
app.secret_key = 'secret_key8902083508##'
app.config.update(
    SESSION_COOKIE_SECURE=True,
    SESSION_COOKIE_SAMESITE='None',
)
# ロガーの設定
logger = logging.getLogger(__name__)
logger.setLevel(logging.INFO)

# ファイルハンドラーの設定
file_handler = logging.FileHandler('app.log', encoding='utf-8')
file_handler.setLevel(logging.INFO)

# フォーマッターの設定
formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
file_handler.setFormatter(formatter)

# ロガーにハンドラーを追加
logger.addHandler(file_handler)

#最初の画面を開く
@app.route('/', methods=['GET', 'POST'])
def index():
    # message = "「年月日」と「勤務」を入力して作成ボタンを押してください。"
    # file_created = False
    session.clear
    # templatesフォルダ内のindex.htmlをレンダリングして返す
    return render_template("index.html")

@app.route('/create_report', methods=['GET', 'POST'])
def create_report():
	session['monthSelect'] = request.form['monthSelect']
	new_filename = f"{session['monthSelect']}「あさか丸」勤務実態表.xlsx"
	template_path = 'work_record.xlsx'
	new_file_path = generate_new_filename(os.path.join('work_records', new_filename))
	workbook = load_workbook(template_path)
	#テンプレートから新しいファイルを作成、日付と曜日を入れる
	fill_dates_and_weekdays(workbook, session['monthSelect'],template_path)
    # sessionのキーに対応するリストを初期化（これはアプリケーションのどこかで行う必要があります）
	session['dailyreports'] = show_dailyreports(session['monthSelect'])
	#ファイルの保存
	workbook.save(new_file_path) 
	flash('新しい勤務実態表を作成しました! %s' % new_filename)
    # 'work_records'フォルダ内のファイルリストを取得
	files = os.listdir('work_records')
	#files_path = ['work_records/' + file for file in files]  # パスを含むファイル名のリスト
	start_date = datetime.strptime(session['monthSelect']+ "-01", "%Y-%m-%d")
	session['start_date'] = start_date.strftime('%Y年%m月') 
	session['files'] = files  # セッションにファイルリストを保存
	session['new_filename'] = os.path.basename(new_file_path)
	return redirect(url_for('show_files'))

@app.route('/show_files')
def show_files():
    files = session.get('files', [])  # セッションからファイルリストを取得
    monthSelect = session.get('start_date','') 
    selected_file = session.get('new_filename','')
    dailyreports = session.get('dailyreports','')
	
    return render_template("show_files.html", files=files, monthSelect=monthSelect, selected_file=selected_file, dailyreports=dailyreports)

@app.route('/edit_report', methods=['GET', 'POST'])
def edit_files():
    selected_files = request.form.getlist('selected_files')
    copy_data_to_work_record(selected_files, session.get('new_filename',''))
    return render_template("index.html") 

@app.route('/create_crew_shift', methods=['GET', 'POST'])
def create_crew_shift():
    if request.method == 'POST':
        try:
            data = request.get_json()
            month  = data.get('month', None)
            session['monthSelect'] = month
            logger.info('monthSelect: %s', month) 
            new_filename = f"{session['monthSelect']}「あさか丸」乗組員勤務表.xlsx"
            template_path = 'crew_shift_template.xlsx'
            new_file_path = generate_new_filename(os.path.join('crew_shift', new_filename))
            workbook = load_workbook(template_path)
            # テンプレートから新しいファイルを作成し、日付と曜日を入れる
            fill_dates_and_weekdays(workbook, session['monthSelect'], template_path)
            # 'work_records'ディレクトリ内のファイルを確認
            work_records_path = 'work_records'
            month_str = session['monthSelect']
            files = os.listdir(work_records_path)
            # logger.info('month_str: %s', month_str) 
            # logger.info('files: %s', files) 
            if not any(file.startswith(month_str) for file in files):
                # logger.info('files_None: %s', files) 
                # flash("対応する「勤務実態表」がありません:")
                # return redirect(url_for('index')) 
                return jsonify({'error': "対応する「勤務実態表」がありません"}), 404
            # 'name.txt'にリストされた名前でシートを作成
            with open('name.txt', 'r', encoding='utf-8') as names_file:
                names = names_file.read().splitlines()
                for name in names:
                    if name not in workbook.sheetnames:
                        workbook.create_sheet(title=name)
                        # ここで、新しく作成したシートに日付や曜日などをコピーする処理を追加する必要があります。
                        copy_dates_to_new_sheet(workbook, 'name', name)
            # logger.info('names: %s', names)             
            # # 元にあった'name'という名前のシートを削除
            if 'name' in workbook.sheetnames:
                std = workbook['name']
                workbook.remove(std)
            
            workbook.save(new_file_path)
            return jsonify({'message': "乗組員の勤務表が正常に作成されました。"}), 200     
        except Exception as e:
            app.logger.error(f'Error occurred: {e}')
            return jsonify({'error': "エラーが発生しました！処理をやり直してください。"}), 404

logging.basicConfig(filename='app.log', level=logging.INFO, 
                    format='%(asctime)s %(levelname)s:%(message)s')

if __name__ == '__main__':
    app.run(debug=True)