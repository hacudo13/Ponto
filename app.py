from flask import Flask, request, jsonify, render_template, send_file
from flask_sqlalchemy import SQLAlchemy
from flask_cors import CORS
from datetime import datetime, date
import pandas as pd
import io

app = Flask(__name__)
app.config['SQLALCHEMY_DATABASE_URI'] = 'sqlite:///ponto.db'
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False
db = SQLAlchemy(app)
CORS(app)

class Employee(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(80), unique=True, nullable=False)
    position = db.Column(db.String(80), nullable=True)
    time_records = db.relationship('TimeRecord', backref='employee', lazy=True)

    def __repr__(self):
        return f'<Employee {self.name}>'

class TimeRecord(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    employee_id = db.Column(db.Integer, db.ForeignKey('employee.id'), nullable=False)
    date = db.Column(db.Date, nullable=False)
    check_in = db.Column(db.DateTime, nullable=False)
    check_out = db.Column(db.DateTime, nullable=True)
    break_start = db.Column(db.DateTime, nullable=True)
    break_end = db.Column(db.DateTime, nullable=True)

    def __repr__(self):
        return f'<TimeRecord {self.employee_id} {self.date}>'

@app.route('/employees', methods=['POST'])
def add_employee():
    data = request.get_json()
    new_employee = Employee(name=data['name'], position=data.get('position'))
    db.session.add(new_employee)
    db.session.commit()
    return jsonify({'message': 'Employee added successfully!'})

@app.route('/employees', methods=['GET'])
def get_employees():
    employees = Employee.query.all()
    output = []
    for employee in employees:
        output.append({'id': employee.id, 'name': employee.name, 'position': employee.position})
    return jsonify({'employees': output})

@app.route('/time_records', methods=['POST'])
def add_time_record():
    data = request.get_json()
    employee_id = data['employee_id']
    record_type = data['type'] # 'check_in', 'check_out', 'break_start', 'break_end'

    employee = Employee.query.get(employee_id)
    if not employee:
        return jsonify({'message': 'Employee not found!'}), 404

    today = datetime.now().date()
    latest_record = TimeRecord.query.filter_by(employee_id=employee_id, date=today).order_by(TimeRecord.id.desc()).first()

    if record_type == 'check_in':
        if latest_record and latest_record.check_out is None:
            return jsonify({'message': 'Employee already checked in!'}), 400
        new_record = TimeRecord(employee_id=employee_id, date=today, check_in=datetime.now())
        db.session.add(new_record)
    elif record_type == 'check_out':
        if not latest_record or latest_record.check_out is not None:
            return jsonify({'message': 'Employee not checked in or already checked out!'}), 400
        latest_record.check_out = datetime.now()
    elif record_type == 'break_start':
        if not latest_record or latest_record.check_out is not None:
            return jsonify({'message': 'Employee not checked in!'}), 400
        if latest_record.break_start is not None and latest_record.break_end is None:
            return jsonify({'message': 'Break already started!'}), 400
        latest_record.break_start = datetime.now()
    elif record_type == 'break_end':
        if not latest_record or latest_record.break_start is None or latest_record.break_end is not None:
            return jsonify({'message': 'Break not started or already ended!'}), 400
        latest_record.break_end = datetime.now()
    else:
        return jsonify({'message': 'Invalid record type!'}), 400

    db.session.commit()
    return jsonify({'message': 'Time record added successfully!'})

@app.route('/time_records/<int:employee_id>', methods=['GET'])
def get_time_records(employee_id):
    records = TimeRecord.query.filter_by(employee_id=employee_id).all()
    output = []
    for record in records:
        output.append({
            'id': record.id,
            'employee_id': record.employee_id,
            'date': record.date.isoformat(),
            'check_in': record.check_in.isoformat() if record.check_in else None,
            'check_out': record.check_out.isoformat() if record.check_out else None,
            'break_start': record.break_start.isoformat() if record.break_start else None,
            'break_end': record.break_end.isoformat() if record.break_end else None
        })
    return jsonify({'time_records': output})

@app.route('/generate_report')
def generate_report():
    employee_id = request.args.get('employee_id')
    start_date = request.args.get('start_date')
    end_date = request.args.get('end_date')
    
    # Converter strings de data para objetos date
    try:
        start_date_obj = datetime.strptime(start_date, '%Y-%m-%d').date()
        end_date_obj = datetime.strptime(end_date, '%Y-%m-%d').date()
    except ValueError:
        return jsonify({'message': 'Formato de data inválido'}), 400
    
    # Query base
    query = TimeRecord.query.filter(
        TimeRecord.date >= start_date_obj,
        TimeRecord.date <= end_date_obj
    )
    
    # Filtrar por funcionário se especificado
    if employee_id:
        query = query.filter(TimeRecord.employee_id == employee_id)
    
    # Executar query e incluir dados do funcionário
    records = query.join(Employee).all()
    
    # Preparar dados para a planilha
    data = []
    for record in records:
        # Calcular horas trabalhadas
        worked_hours = 0
        break_duration = 0
        
        if record.check_in and record.check_out:
            total_time = (record.check_out - record.check_in).total_seconds() / 3600
            
            if record.break_start and record.break_end:
                break_duration = (record.break_end - record.break_start).total_seconds() / 3600
            
            worked_hours = total_time - break_duration
        
        data.append({
            'Funcionário': record.employee.name,
            'Cargo': record.employee.position or 'Não informado',
            'Data': record.date.strftime('%d/%m/%Y'),
            'Entrada': record.check_in.strftime('%H:%M') if record.check_in else '',
            'Início Intervalo': record.break_start.strftime('%H:%M') if record.break_start else '',
            'Fim Intervalo': record.break_end.strftime('%H:%M') if record.break_end else '',
            'Saída': record.check_out.strftime('%H:%M') if record.check_out else '',
            'Horas Trabalhadas': f'{worked_hours:.2f}' if worked_hours > 0 else '0.00',
            'Duração Intervalo (h)': f'{break_duration:.2f}' if break_duration > 0 else '0.00'
        })
    
    # Criar DataFrame
    df = pd.DataFrame(data)
    
    if df.empty:
        return jsonify({'message': 'Nenhum registro encontrado para o período especificado'}), 404
    
    # Criar arquivo Excel em memória
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='Relatório de Ponto', index=False)
        
        # Ajustar largura das colunas
        worksheet = writer.sheets['Relatório de Ponto']
        for column in worksheet.columns:
            max_length = 0
            column_letter = column[0].column_letter
            for cell in column:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
            adjusted_width = min(max_length + 2, 50)
            worksheet.column_dimensions[column_letter].width = adjusted_width
    
    output.seek(0)
    
    # Nome do arquivo
    filename = f'relatorio_ponto_{start_date}_a_{end_date}.xlsx'
    if employee_id:
        employee = Employee.query.get(employee_id)
        if employee:
            filename = f'relatorio_ponto_{employee.name}_{start_date}_a_{end_date}.xlsx'
    
    return send_file(
        output,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        as_attachment=True,
        download_name=filename
    )

@app.route('/')
def index():
    return render_template('index.html')

if __name__ == '__main__':
    with app.app_context():
        db.create_all()
    app.run(debug=True, host='0.0.0.0')


