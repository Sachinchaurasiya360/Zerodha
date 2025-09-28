from flask import Flask, render_template, request, send_file, Response
import sqlite3
from datetime import datetime
import pandas as pd
import os
import io

app = Flask(__name__)

@app.route('/')
def index():
    return render_template('index.html', selected_date='', no_data=False)

@app.route('/attendance', methods=['POST'])
def attendance():
    selected_date = request.form.get('selected_date')
    selected_date_obj = datetime.strptime(selected_date, '%Y-%m-%d')
    formatted_date = selected_date_obj.strftime('%Y-%m-%d')

    conn = sqlite3.connect('attendance.db')
    cursor = conn.cursor()

    cursor.execute("SELECT name, time FROM attendance WHERE date = ?", (formatted_date,))
    attendance_data = cursor.fetchall()

    conn.close()

    if not attendance_data:
        return render_template('index.html', selected_date=selected_date, no_data=True)
    
    return render_template('index.html', selected_date=selected_date, attendance_data=attendance_data)

@app.route('/export', methods=['POST'])
def export():
    export_date = request.form.get('export_date')
    export_date_obj = datetime.strptime(export_date, '%Y-%m-%d')
    formatted_date = export_date_obj.strftime('%Y-%m-%d')
    
    # Connect to the database and fetch the attendance data
    conn = sqlite3.connect('attendance.db')
    
    # Create a pandas DataFrame from the database query
    query = f"SELECT name, time FROM attendance WHERE date = '{formatted_date}'"
    df = pd.read_sql_query(query, conn)
    
    # Add a serial number column
    df.insert(0, 'S.No', range(1, len(df) + 1))
    
    # Add a date column
    df['Date'] = formatted_date
    
    # Reorder columns
    df = df[['S.No', 'Date', 'Name', 'Time']]
    
    conn.close()
    
    if df.empty:
        return render_template('index.html', selected_date=export_date, no_data=True)
    
    # Create an Excel file in memory
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Attendance Data')
        
        # Auto-adjust columns' width
        workbook = writer.book
        worksheet = writer.sheets['Attendance Data']
        for i, col in enumerate(df.columns):
            max_len = max(df[col].astype(str).map(len).max(), len(col)) + 2
            worksheet.column_dimensions[chr(65+i)].width = max_len
    
    output.seek(0)
    
    # Generate filename with date
    filename = f"Attendance_{formatted_date.replace('-', '_')}.xlsx"
    
    # Return the Excel file as a downloadable attachment
    return Response(
        output,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f"attachment;filename={filename}"}
    )

if __name__ == '__main__':
    app.run(debug=True)
