from flask import Flask, request, render_template, send_file
import sqlite3
import pandas as pd
from docxtpl import DocxTemplate
from pyngrok import ngrok

app = Flask(__name__)

def init_db():
    conn = sqlite3.connect('database.db')
    c = conn.cursor()
    c.execute('''
        CREATE TABLE IF NOT EXISTS disciplinas (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            nome TEXT,
            periodo INTEGER,
            ementa TEXT,
            bibliografia_basica TEXT,
            bibliografia_complementar TEXT
        )
    ''')
    conn.commit()
    conn.close()

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/add', methods=['POST'])
def add_disciplina():
    nome = request.form['nome']
    periodo = request.form['periodo']
    ementa = request.form['ementa']
    bibliografia_basica = request.form['bibliografia_basica']
    bibliografia_complementar = request.form['bibliografia_complementar']

    conn = sqlite3.connect('database.db')
    c = conn.cursor()
    c.execute('''
        INSERT INTO disciplinas (nome, periodo, ementa, bibliografia_basica, bibliografia_complementar)
        VALUES (?, ?, ?, ?, ?)
    ''', (nome, periodo, ementa, bibliografia_basica, bibliografia_complementar))
    conn.commit()
    conn.close()

    return 'Disciplina adicionada com sucesso!'

@app.route('/report')
def generate_report():
    conn = sqlite3.connect('database.db')
    df = pd.read_sql_query("SELECT * FROM disciplinas ORDER BY periodo, nome", conn)
    conn.close()

    semestres = {}
    for _, row in df.iterrows():
        periodo = row['periodo']
        if periodo not in semestres:
            semestres[periodo] = []
        semestres[periodo].append(row)

    doc = DocxTemplate("templates/template.docx")
    doc.render({'semestres': semestres})
    doc.save("output.docx")

    return send_file("output.docx", as_attachment=True)

if __name__ == '__main__':
    init_db()
    # Expor o servidor local via ngrok
    public_url = ngrok.connect(5000)
    print(f"ngrok URL: {public_url}")
    app.run(port=5000)
