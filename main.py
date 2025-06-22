from flask import Flask, render_template, request, send_from_directory
import os
from werkzeug.utils import secure_filename
import cloudconvert
import requests

app = Flask(__name__)
API_KEY = "SEU_API_KEY"
cloudconvert.configure(api_key=API_KEY)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_files():
    file = request.files['file']
    input_path = os.path.join('uploads', secure_filename(file.filename))
    file.save(input_path)

    # Cria o job
    job = cloudconvert.Job.create(payload={
        "tasks": {
            'import-my-file': {
                'operation': 'import/upload'
            },
            'convert-my-file': {
                'operation': 'convert',
                'input': 'import-my-file',
                'input_format': 'pptx',
                'output_format': 'pdf',
            },
            'export-my-file': {
                'operation': 'export/url',
                'input': 'convert-my-file'
            }
        }
    })

    # Imprime a resposta completa do job para diagnóstico
    print("Resposta do Job:", job)

    # Verifica se 'tasks' existe na resposta
    if 'tasks' not in job:
        return f"Erro: 'tasks' não encontrado na resposta do job. Resposta completa: {job}"

    # Pega a tarefa de upload
    upload_task_id = job['tasks'][0]['id']
    upload_task = cloudconvert.Task.find(id=upload_task_id)

    # Faz o upload do arquivo
    upload_url = upload_task['result']['form']['url']
    upload_params = upload_task['result']['form']['parameters']
    with open(input_path, 'rb') as f:
        files = {'file': (file.filename, f)}
        requests.post(upload_url, data=upload_params, files=files)

    # Espera a conversão terminar
    job = cloudconvert.Job.wait(id=job['id'])

    # Verifica novamente se 'tasks' existe após o processamento
    if 'tasks' not in job:
        return f"Erro: 'tasks' não encontrado após a conversão. Resposta completa: {job}"

    # Pega o link de download
    tasks = cloudconvert.Job.find(id=job['id'])['tasks']
    export_task = next(task for task in tasks if task['name'] == 'export-my-file')
    file_url = export_task['result']['files'][0]['url']

    return f'<a href="{file_url}" target="_blank">Download PDF</a>'

if __name__ == '__main__':
    app.run(debug=True)
