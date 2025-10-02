from flask import Flask, render_template, request, send_file, Response
import sqlite3
from datetime import datetime
import pandas as pd
import os
import io

app = Flask(__name__)

@app.route('/')
def index():
    # Get some basic statistics
    conn = sqlite3.connect('attendance.db')
    cursor = conn.cursor()
    
    # Get total records count
    cursor.execute("SELECT COUNT(*) FROM attendance")
    total_records = cursor.fetchone()[0]
    
    # Get unique dates count
    cursor.execute("SELECT COUNT(DISTINCT date) FROM attendance")
    total_days = cursor.fetchone()[0]
    
    # Get unique people count
    cursor.execute("SELECT COUNT(DISTINCT name) FROM attendance")
    total_people = cursor.fetchone()[0]
    
    conn.close()
    
    stats = {
        'total_records': total_records,
        'total_days': total_days,
        'total_people': total_people
    }
    
    return render_template('index.html', selected_date='', no_data=False, stats=stats)

@app.route('/attendance', methods=['POST'])
def attendance():
    selected_date = request.form.get('selected_date')
    if not selected_date:
        return render_template('index.html', selected_date='', no_data=True)
        
    selected_date_obj = datetime.strptime(selected_date, '%Y-%m-%d')
    formatted_date = selected_date_obj.strftime('%Y-%m-%d')

    conn = sqlite3.connect('attendance.db')
    cursor = conn.cursor()

    cursor.execute("SELECT name, time FROM attendance WHERE date = ?", (formatted_date,))
    attendance_data = cursor.fetchall()
    
    # Get statistics
    cursor.execute("SELECT COUNT(*) FROM attendance")
    total_records = cursor.fetchone()[0]
    
    cursor.execute("SELECT COUNT(DISTINCT date) FROM attendance")
    total_days = cursor.fetchone()[0]
    
    cursor.execute("SELECT COUNT(DISTINCT name) FROM attendance")
    total_people = cursor.fetchone()[0]

    conn.close()
    
    stats = {
        'total_records': total_records,
        'total_days': total_days,
        'total_people': total_people
    }

    if not attendance_data:
        return render_template('index.html', selected_date=selected_date, no_data=True, stats=stats)
    
    return render_template('index.html', selected_date=selected_date, attendance_data=attendance_data, stats=stats)

@app.route('/export', methods=['POST'])
def export():
    # Connect to the database and fetch ALL attendance data
    conn = sqlite3.connect('attendance.db')
    
    # Create a pandas DataFrame from the database query for all records
    query = "SELECT name, time, date FROM attendance ORDER BY date DESC, time ASC"
    df = pd.read_sql_query(query, conn)
    
    conn.close()
    
    if df.empty:
        return render_template('index.html', selected_date='', no_data=True)
    
    # Add a serial number column
    df.insert(0, 'S.No', range(1, len(df) + 1))
    
    # Reorder columns for better presentation
    df = df[['S.No', 'Date', 'Name', 'Time']]
    
    # Rename columns for better presentation
    df.columns = ['S.No', 'Date', 'Name', 'Time']
    
    # Create an Excel file in memory with multiple sheets
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        # Write all data to main sheet
        df.to_excel(writer, index=False, sheet_name='All Attendance Records')
        
        # Create summary sheet with date-wise attendance count
        summary_df = df.groupby('Date').agg({
            'Name': 'count',
            'Time': ['min', 'max']
        }).reset_index()
        
        # Flatten column names
        summary_df.columns = ['Date', 'Total Present', 'First Arrival', 'Last Arrival']
        summary_df.insert(0, 'S.No', range(1, len(summary_df) + 1))
        
        summary_df.to_excel(writer, index=False, sheet_name='Daily Summary')
        
        # Auto-adjust columns' width for both sheets
        for sheet_name in ['All Attendance Records', 'Daily Summary']:
            worksheet = writer.sheets[sheet_name]
            if sheet_name == 'All Attendance Records':
                current_df = df
            else:
                current_df = summary_df
                
            for i, col in enumerate(current_df.columns):
                max_len = max(current_df[col].astype(str).map(len).max(), len(col)) + 2
                worksheet.column_dimensions[chr(65+i)].width = max_len
    
    output.seek(0)
    
    # Generate filename with current date
    current_date = datetime.now().strftime('%Y_%m_%d')
    filename = f"Complete_Attendance_Report_{current_date}.xlsx"
    
    # Return the Excel file as a downloadable attachment
    return Response(
        output,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f"attachment;filename={filename}"}
    )

@app.route('/export_date', methods=['POST'])
def export_date():
    export_date = request.form.get('export_date')
    if not export_date:
        return render_template('index.html', selected_date='', no_data=True)
        
    export_date_obj = datetime.strptime(export_date, '%Y-%m-%d')
    formatted_date = export_date_obj.strftime('%Y-%m-%d')
    
    # Connect to the database and fetch the attendance data for specific date
    conn = sqlite3.connect('attendance.db')
    
    # Create a pandas DataFrame from the database query
    query = f"SELECT name, time, date FROM attendance WHERE date = '{formatted_date}' ORDER BY time ASC"
    df = pd.read_sql_query(query, conn)
    
    conn.close()
    
    if df.empty:
        return render_template('index.html', selected_date=export_date, no_data=True)
    
    # Add a serial number column
    df.insert(0, 'S.No', range(1, len(df) + 1))
    
    # Reorder columns for better presentation
    df = df[['S.No', 'Date', 'Name', 'Time']]
    
    # Create an Excel file in memory
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name=f'Attendance_{formatted_date}')
        
        # Auto-adjust columns' width
        worksheet = writer.sheets[f'Attendance_{formatted_date}']
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
