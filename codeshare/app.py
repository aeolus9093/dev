from flask import Flask, render_template, request, redirect, url_for, send_from_directory, flash
import os
import shutil
import datetime
import socket
import zipfile
from werkzeug.utils import secure_filename

app = Flask(__name__)
app.secret_key = 'codeshare_secret_key'  # 세션을 위한 시크릿 키

# 공유 폴더 경로 설정
SHARE_FOLDER = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'shared_files')
ALLOWED_EXTENSIONS = {'py', 'js', 'html', 'css', 'txt', 'md', 'json', 'xml', 'csv', 'zip'}

# 공유 폴더가 없으면 생성
if not os.path.exists(SHARE_FOLDER):
    os.makedirs(SHARE_FOLDER)

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def get_local_ip():
    try:
        # 로컬 IP 주소 가져오기
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.connect(('8.8.8.8', 80))
        ip = s.getsockname()[0]
        s.close()
        return ip
    except:
        return '127.0.0.1'

@app.route('/')
def index():
    files = []
    for filename in os.listdir(SHARE_FOLDER):
        file_path = os.path.join(SHARE_FOLDER, filename)
        if os.path.isfile(file_path):
            file_size = os.path.getsize(file_path)
            modified_time = datetime.datetime.fromtimestamp(os.path.getmtime(file_path))
            files.append({
                'name': filename,
                'size': file_size,
                'modified': modified_time.strftime('%Y-%m-%d %H:%M:%S')
            })
    return render_template('index.html', files=files)

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        flash('파일이 선택되지 않았습니다.', 'error')
        return redirect(request.url)
    
    files = request.files.getlist('file')
    
    for file in files:
        if file.filename == '':
            flash('파일이 선택되지 않았습니다.', 'error')
            continue
        
        if file and allowed_file(file.filename):
            filename = secure_filename(file.filename)
            file.save(os.path.join(SHARE_FOLDER, filename))
            flash(f'{filename} 파일이 업로드되었습니다.', 'success')
        else:
            flash(f'{file.filename}은(는) 허용되지 않는 파일 형식입니다.', 'error')
    
    return redirect(url_for('index'))

@app.route('/download/<filename>')
def download_file(filename):
    return send_from_directory(SHARE_FOLDER, filename, as_attachment=True)

@app.route('/delete/<filename>', methods=['POST'])
def delete_file(filename):
    file_path = os.path.join(SHARE_FOLDER, filename)
    if os.path.exists(file_path):
        os.remove(file_path)
        flash(f'{filename} 파일이 삭제되었습니다.', 'success')
    else:
        flash(f'{filename} 파일을 찾을 수 없습니다.', 'error')
    return redirect(url_for('index'))

@app.route('/view/<filename>')
def view_file(filename):
    file_path = os.path.join(SHARE_FOLDER, filename)
    if not os.path.exists(file_path):
        flash(f'{filename} 파일을 찾을 수 없습니다.', 'error')
        return redirect(url_for('index'))
    
    file_ext = filename.rsplit('.', 1)[1].lower() if '.' in filename else ''
    
    if file_ext in ['py', 'js', 'html', 'css', 'txt', 'md', 'json', 'xml', 'csv']:
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                content = f.read()
            return render_template('view.html', filename=filename, content=content, file_ext=file_ext)
        except:
            flash(f'{filename} 파일을 읽을 수 없습니다.', 'error')
            return redirect(url_for('index'))
    else:
        flash(f'{filename} 파일은 미리보기를 지원하지 않습니다.', 'error')
        return redirect(url_for('index'))

@app.route('/zip_download', methods=['POST'])
def zip_download():
    selected_files = request.form.getlist('selected_files')
    if not selected_files:
        flash('다운로드할 파일을 선택해주세요.', 'error')
        return redirect(url_for('index'))
    
    zip_filename = f'codeshare_{datetime.datetime.now().strftime("%Y%m%d_%H%M%S")}.zip'
    zip_path = os.path.join(SHARE_FOLDER, zip_filename)
    
    with zipfile.ZipFile(zip_path, 'w') as zipf:
        for filename in selected_files:
            file_path = os.path.join(SHARE_FOLDER, filename)
            if os.path.exists(file_path):
                zipf.write(file_path, filename)
    
    flash(f'선택한 파일들이 {zip_filename}으로 압축되었습니다.', 'success')
    return redirect(url_for('download_file', filename=zip_filename))

if __name__ == '__main__':
    ip = get_local_ip()
    print(f"* 코드 공유 서버가 시작되었습니다.")
    print(f"* 로컬 네트워크에서 접속: http://{ip}:5000")
    app.run(host='0.0.0.0', port=5000, debug=True)
