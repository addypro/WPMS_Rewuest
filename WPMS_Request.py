import os
from flask import Flask, request, render_template, redirect, url_for, flash
from werkzeug.utils import secure_filename
import psycopg2
import pandas as pd


template_dir = os.path.abspath('D:\Support (Python)\WPMS_Request_Project')

app = Flask(__name__, template_folder=template_dir)
app.secret_key = 'supersecretkey'  # Needed for session management for flash messages
app.config['UPLOAD_FOLDER'] = 'D:\Support (Python)\WPMS_Request_Project\Web_Tests'
app.config['ALLOWED_EXTENSIONS'] = {'xlsx', 'xls', 'csv'}
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16 MB limit

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in app.config['ALLOWED_EXTENSIONS']

def insert_requestor(requestor):
    sql = """INSERT INTO requestor(name, company, email, phone, submission_date, submission_time)
             VALUES(%s, %s, %s, %s, current_date, current_time) RETURNING requestor_id;"""
    conn = psycopg2.connect(
        dbname="WPMS_Requests", user="postgres", password="Merck123", host="localhost"
    )
    cur = conn.cursor()
    cur.execute(sql, requestor)
    requestor_id = cur.fetchone()[0]
    conn.commit()
    cur.close()
    conn.close()
    return requestor_id

def insert_request(data):
    sql = """INSERT INTO requests(
                requestor_id, request_type, comments, disregard_begin, disregard_remind, 
                point, building, location, equipment_ctu_id, measurement_sensor, gxp, 
                operating_units_states, signal_source, source_address, excursion_delay, 
                operating_setpoint, low_excursion_limit, high_excursion_limit, pm_task_id, 
                equipment_asset_id, department_cost_center, dispatch_trade_number, dispatch_delay, 
                owner_notify_delay, owners, area_coordinator, other_fields
             )
             VALUES(
                %s, %s, %s, %s, %s, 
                %s, %s, %s, %s, %s, %s, 
                %s, %s, %s, %s, %s, 
                %s, %s, %s, %s, %s, 
                %s, %s, %s, %s, %s, %s
             )"""
    conn = psycopg2.connect(
        dbname="WPMS_Requests", user="postgres",  password="Merck123", host="localhost"
    )
    cur = conn.cursor()
    cur.executemany(sql, data)  # Support bulk inserts
    conn.commit()
    cur.close()
    conn.close()

def insert_bulk_requests(data):
    sql = """INSERT INTO requests(
                requestor_id, request_type, comments, disregard_begin, disregard_remind, point)
             VALUES (%s, %s, %s, %s, %s, %s)"""
    conn = psycopg2.connect(
        dbname="WPMS_Requests", user="postgres", password="Merck123", host="localhost"
    )
    cur = conn.cursor()
    cur.executemany(sql, data)
    conn.commit()
    cur.close()
    conn.close()

@app.route('/', methods=['GET', 'POST'])
def index():
    app.logger.debug('Debug message: Inside the index route')
    if request.method == 'POST':
        app.logger.info('hi')
        try:
            requestor_name = request.form['requestor_name']
            requestor_company = request.form['requestor_company']
            requestor_email = request.form['requestor_email']
            requestor_phone = request.form['requestor_phone']
            request_type = request.form['request_type']
            comments = request.form.get('comments')
            requestor_id = insert_requestor((requestor_name, requestor_company, requestor_email, requestor_phone))

            print(f"Requestor: {requestor_name}, {requestor_company}, {requestor_email}, {requestor_phone}")
            print(f"Request Type: {request_type}")

            # Handle Disregard, Resume, and Remove/Spare request types
            if request_type in ['Disregard', 'Resume', 'Remove/Spare']:
                points_data = []

                if 'bulk_points' in request.form and request.form['bulk_points']:
                    points = [point.strip() for point in request.form['bulk_points'].split(',')]
                    disregard_begin = request.form.get('bulk_disregard_begin')
                    disregard_remind = request.form.get('bulk_disregard_remind')

                    for point in points:
                        if request_type == 'Disregard':
                            points_data.append((requestor_id, request_type, comments, disregard_begin, disregard_remind, point))
                        else:
                            points_data.append((requestor_id, request_type, comments, None, None, point))

                elif 'bulk_file' in request.files and request.files['bulk_file']:
                    file = request.files['bulk_file']
                    print(f"File received: {file.filename}")
                    if allowed_file(file.filename):
                        filepath = os.path.join(app.config['UPLOAD_FOLDER'], secure_filename(file.filename))
                        file.save(filepath)
                        df = pd.read_excel(filepath) if file.filename.endswith('.xlsx') else pd.read_csv(filepath)
                        print(f"DataFrame loaded: {df.head()}")
                        for _, row in df.iterrows():
                            if request_type == 'Disregard':
                                points_data.append((requestor_id, request_type, comments, row['Disregard Begin'], row['Disregard Remind'], row['Point Number']))
                            else:
                                points_data.append((requestor_id, request_type, comments, None, None, row['Point Number']))

                else:
                    points = request.form.getlist('points[]')
                    disregard_begins = request.form.getlist('disregard_begin[]')
                    disregard_reminds = request.form.getlist('disregard_remind[]')

                    for i, point in enumerate(points):
                        if request_type == 'Disregard':
                            points_data.append((requestor_id, request_type, comments, disregard_begins[i], disregard_reminds[i], point))
                        else:
                            points_data.append((requestor_id, request_type, comments, None, None, point))

                insert_bulk_requests(points_data)
                return redirect(url_for('confirmation'))

            # Handle other request types like Add, Edit, Repurpose with file upload
            elif request_type in ['Add', 'Edit', 'Repurpose']:
                if 'file' not in request.files:
                    print(f"File received: {file.filename}")
                    flash('No file part', 'error')
                    return redirect(request.url)
                file = request.files['file']
                if file and allowed_file(file.filename):
                    filename = secure_filename(file.filename)
                    filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
                    file.save(filepath)

                    print(f"File saved to: {filepath}")

                    # Process the uploaded file
                    df = pd.read_excel(filepath) if filename.endswith('.xlsx') else pd.read_csv(filepath)

                    print("Uploaded file content preview:")
                    print(df.head())

                    # Insert into the database
                    for _, row in df.iterrows():
                        data = (
                            requestor_id, request_type, comments, None, None, 
                            row.get('Point'), row.get('Building'), row.get('Location'), 
                            row.get('Equipment/CTU ID'), row.get('Measurement/Sensor'), 
                            row.get('GxP') == 'Yes', row.get('Operating Units/States'), 
                            row.get('Signal Source'), row.get('Source Address (I/O or OPC)'), 
                            row.get('Excursion Delay'), row.get('Operating Setpoint'), 
                            row.get('Low Excursion Limit'), row.get('High Excursion Limit'), 
                            row.get('PM Task ID'), row.get('Equipment Asset ID'), 
                            row.get('Department or Cost Center'), row.get('Dispatch Trade #'), 
                            row.get('Dispatch Delay'), row.get('Owner Notify Delay'), 
                            row.get('Owner 1 Name'), row.get('Owner 1 Mobile #'), 
                            row.get('Owner 1 Work #'), row.get('Owner 2 Name'), 
                            row.get('Owner 2 Mobile #'), row.get('Owner 2 Work #'), 
                            row.get('Owner 3 Name'), row.get('Owner 3 Mobile #'), 
                            row.get('Owner 3 Work #'), row.get('Area Coordinator Name'), 
                            row.get('Area Coordinator Mobile #'), row.get('Area Coordinator Work #')
                        )
                        insert_request([data])  # Updated to pass as a list for executemany
                
                    return redirect(url_for('confirmation'))
                else:
                    flash('Invalid file type. Please upload an Excel or CSV file.', 'error')
                    return redirect(request.url)

        except Exception as e:
            print(f"An error occurred: {e}")
            flash('An error occurred while processing the request. Please try again.', 'error')
            return redirect(request.url)
    
    return render_template('index.html')

@app.route('/confirmation')
def confirmation():
    return render_template('confirmation.html')

if __name__ == '__main__':
    os.environ['FLASK_APP'] = 'app'
    app.run(debug=True)
