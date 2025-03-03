from io import BytesIO
from flask import Flask,render_template,request,redirect, send_file,url_for,jsonify,flash,session

from dotenv import load_dotenv
import os
from openpyxl import load_workbook
from flask_sqlalchemy import SQLAlchemy
from googleapiclient.discovery import build
from google.oauth2 import service_account
import json
from datetime import datetime
import openpyxl
from sqlalchemy import case, extract, func, or_
import requests
from google_auth_oauthlib.flow import Flow
from google.oauth2.credentials import Credentials
os.environ['OAUTHLIB_INSECURE_TRANSPORT'] = '0'
app=Flask(__name__)


CLIENT_ID = os.environ.get('CLIENT_ID')
CLIENT_SECRET = os.environ.get('CLIENT_SECRET')
REDIRECT_URI='http://127.0.0.1:5000/call_back'
SCOPES = ['openid', 'https://www.googleapis.com/auth/userinfo.email', 'https://www.googleapis.com/auth/userinfo.profile']


flow = Flow.from_client_config(
    {
        "web": {
            "client_id": CLIENT_ID,
            "client_secret": CLIENT_SECRET,
            "redirect_uris": [REDIRECT_URI],
            "auth_uri": "https://accounts.google.com/o/oauth2/auth",
            "token_uri": "https://oauth2.googleapis.com/token",
        }
    },
    scopes=SCOPES
)

excel_path = os.path.join(app.static_folder,'files','DMAX-2024-Live.xlsx')
workbook = load_workbook(excel_path)
sheet = workbook.active
def find_next_available_row(sheet):
    for row in range(1, sheet.max_row + 1):
        if all([cell.value in [None, ""] for cell in sheet[row]]):
            
            return row
        
    return sheet.max_row + 1 

def credentials_to_dict(credentials):
    return {
        'token': credentials.token,
        'refresh_token': credentials.refresh_token,
        'token_uri': credentials.token_uri,
        'client_id': credentials.client_id,
        'client_secret': credentials.client_secret,
        'scopes': credentials.scopes
    }

app.config['SECRET_KEY'] = 'your_secret_key'
app.config['SQLALCHEMY_DATABASE_URI'] = 'sqlite:///employees.db'
app.config['SQLALCHEMY_BINDS']={
    'dform':'sqlite:///dform.db',
    'emp_info':'sqlite:///empinfo.db',
    'op_excellence':'sqlite:///opexcellence.db',
    'target_columns':'sqlite:///target_columns.db',
    'project_targets':'sqlite:///project_targets.db',
    'project_targets':'sqlite:///project_targets.db',
    'dmax_approval':'sqlite:///dmax_approval.db'
}
db = SQLAlchemy(app)


###############################################  Helper Functions ###############################################

def get_logged_in_user_details():
    """
    Retrieves the logged-in user's name from the Employee table based on session data.
    """
    # Check if the user logged in with username/password
    if 'username' in session:
        username = session['username']
        user = Employee.query.filter_by(emp_id=username).first()  # Match emp_id with the username
        if user:
            return {"name": user.name, "role": user.role,"email":user.email}

    # Check if the user logged in with Google Sign-In (using email)
    if 'email' in session:
        email = session['email']
        user = Employee.query.filter_by(email=email).first()  # Match email with the logged-in user's email
        if user:
            return {"name": user.name, "role": user.role,"email":user.email}

    # If no user is found, return None or an appropriate message
    return None

def get_filtered_employees(base_query, search_query, selected_month, selected_date):
    if search_query:
        base_query = base_query.filter(func.lower(Dform.employee_name) == search_query)
    if selected_month:
        base_query = base_query.filter(extract('month', Dform.today_date) == int(selected_month))
    if selected_date:
        base_query = base_query.filter_by(today_date=selected_date)
    return base_query.all()

def get_first_filtered_employees(base_query, search_query, selected_month, selected_date,selected_year):
    if search_query:
        base_query = base_query.filter(func.lower(Dform.employee_name) == search_query)
    if selected_month:
        base_query = base_query.filter(extract('month', Dform.today_date) == int(selected_month))
    if selected_date:
        try:
            selected_date = datetime.strptime(selected_date, "%Y-%m-%d").date()  # Convert to date
            base_query = base_query.filter(Dform.today_date == selected_date)  # ✅ Correct filtering
        except ValueError:
            print("Invalid date format:", selected_date)  # Debugging log
    if selected_year:  # Add year filtering
        base_query = base_query.filter(extract('year', Dform.today_date) == int(selected_year))    
    return base_query

def get_date_range_for_month(month):
    current_year = datetime.now().year
    next_year = current_year
    month = int(month)  # Convert to integer
    if month == 1:  
        
        prev_month=12
        next_year=current_year+1
    else:
        prev_month=month-1   
        next_year=current_year 
    start_date = datetime(current_year, prev_month, 26)
    end_date = datetime(next_year, month, 25)
    
    return start_date.strftime("%Y-%m-%d"), end_date.strftime("%Y-%m-%d")

def get_averages_for_filtered_employees(filtered_entries):
    if filtered_entries:
        averages = filtered_entries.with_entities(
            func.avg(Dform.target).label('avg_target'),
            func.avg(Dform.actual).label('avg_actual'),
            func.avg(Dform.production).label('avg_production'),
            func.avg(Dform.quality).label('avg_quality'),
            func.avg(Dform.attendance).label('avg_attendance'),
            func.avg(Dform.skill).label('avg_skill'),
            func.avg(Dform.new_initiatives).label('avg_new_initiatives'),
            func.avg(Dform.Dmax_score).label('avg_Dmax_score')
        ).first()
        return {
                "avg_target": round(averages.avg_target, 2) if averages.avg_target else 0,
                "avg_actual": round(averages.avg_actual, 2) if averages.avg_actual else 0,
                "avg_production": round(averages.avg_production, 2) if averages.avg_production else 0,
                "avg_quality": round(averages.avg_quality, 2) if averages.avg_quality else 0,
                "avg_attendance": round(averages.avg_attendance, 2) if averages.avg_attendance else 0,
                "avg_skill": round(averages.avg_skill, 2) if averages.avg_skill else 0,
                "avg_new_initiatives": round(averages.avg_new_initiatives, 2) if averages.avg_new_initiatives else 0,
                "avg_Dmax_score": round(averages.avg_Dmax_score, 2) if averages.avg_Dmax_score else 0
            }
    return None


def calculate_attendance(designation, attendance_input):
    if designation in ["Intern", "Jr.QA Engineer"]:
        
        attendance = int((attendance_input * 5 / 100) * 100)
    elif designation in ["QA Engineer", "Sr.QA Engineer", "QA Lead"]:
        attendance = int((attendance_input * 5 / 100) * 100)
    else:
        attendance = 0  # Default case

    return attendance*10

def generate_excel_from_template(employees):
    directory = os.path.abspath("static/files")
    filename = "DMAX-sample.xlsx"
    sample_file_path = os.path.join(directory, filename)
    wb = load_workbook(sample_file_path)
    ws = wb.active  # Get the active sheet
    
    start_row = 4  

    # List of database columns corresponding to template headers
    ALLOWED_COLUMNS = [
            "employee_name","employee_id","employee_email","today_date","project","designation",
            "test_case_creation_target","test_case_creation_actual",
            "test_case_updation_target", "test_case_updation_actual",
            "test_case_execution_target", "test_case_execution_actual", 
            "defects_found_target", "defects_found_actual","defects_verification_target", "defects_verification_actual", 
            "test_scripts_creation_target", "test_scripts_creation_actual","test_scripts_updation_target", "test_scripts_updation_actual",
            "test_scripts_execution_target","test_scripts_execution_actual","project_doc_target",
            "project_doc_actual", "internal_Review_target", "internal_Review_actual", "regression_cycle_target",
            "regression_cycle_actual", "req_anal_target", "req_anal_actual", "end_cases_exec_target", "end_cases_exec_actual",
            "site_Scrub_target", "site_Scrub_actual", 
            "task_coverage_score_target", "task_coverage_score_actual",
            "assessment_score_target", "assessment_score_actual", "assessment_re_score_target",
            "assessment_re_score_actual", "cert_score_target", "cert_score_actual", "cert_re_score_target","cert_re_score_actual",
            "new_features_imp_target", "new_features_imp_actual", "defects_fixed_target",
            "defects_fixed_actual", "enhancements_target", "enhancements_actual", "fig_desgns_target",
            "fig_desgns_actual", "doc_update_target", "doc_update_actual", "research_target", "research_actual",
            "inv_defs", "spel_errors", "client_esc", "tst_cases_missing", "attendance", "skill", "new_initiatives", "target", "actual", "production",
            "quality", "attendance", "skill", "new_initiatives", "Dmax_score"
    ]
    for row_num, emp in enumerate(employees, start=start_row):
        for col_num, column_name in enumerate(ALLOWED_COLUMNS, start=1):
            ws.cell(row=row_num, column=col_num, value=getattr(emp, column_name, ''))

    # Save the modified file to memory (without changing the original)
    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return send_file(output, download_name="filtered_employees.xlsx", as_attachment=True, mimetype="application/octet-stream")

def has_previous_month_entry(employee_id, month):
    """ Check if the employee has at least one entry for the previous month. """
    previous_month = month - 1 if month > 1 else 12
    

    return db.session.query(Dform).filter(
        Dform.employee_id == employee_id,
        extract('month', Dform.today_date) == previous_month  # Extract month from today_date
    ).first() is not None
###############################################  Month dictionary ###############################################

monthsDict = {
    "01": "January",
    "02": "February",
    "03": "March",
    "04": "April",
    "05": "May",
    "06": "June",
    "07": "July",
    "08": "August",
    "09": "September",
    "10": "October",
    "11": "November",
    "12": "December"
  }

monthsDict_2 = {
    "January": 1, "February": 2, "March": 3, "April": 4,
    "May": 5, "June": 6, "July": 7, "August": 8,
    "September": 9, "October": 10, "November": 11, "December": 12
}

corrections = {
                    "SrQAEngineer": "Sr.QA Engineer",
                    "Sr QA Engineer": "Sr.QA Engineer",
                    "QAEngineer": "QA Engineer",
                    "JrQAEngineer": "Jr.QA Engineer",
                    "Jr QA Engineer": "Jr.QA Engineer",
                    "QALead": "QA Lead"
                    }

###############################################  Helper Variables ###############################################

current_date = datetime.now().date()
current_month=current_date.month
current_year=current_date.year
last_ten_years = [current_year +1 - i for i in range(11)]
projects=['Akyrian','Auxo','Avanti','Bench','Fora Travels','Indihood','IPS','IQHive','LevelBlue','Web Development','Opus Clip','Training']
designations=['Intern','Jr.QA Engineer','QA Engineer','Sr.QA Engineer','QA Lead']

###############################################  Database classes ###############################################



class Dform(db.Model):
    __tablename__ = 'login'
    __bind_key__="dform"
    id = db.Column(db.Integer, primary_key=True)
    employee_name=db.Column(db.String(100),nullable=False)
    employee_id=db.Column(db.String(100),nullable=False)
    employee_email=db.Column(db.String(100),nullable=False)
    today_date=db.Column(db.String(100),nullable=False)
    project=db.Column(db.String(100),nullable=False)
    designation=db.Column(db.String(100),nullable=False)
    test_case_creation_target= db.Column(db.Integer)
    test_case_creation_actual=db.Column(db.Integer)
    test_case_updation_target=db.Column(db.Integer)
    test_case_updation_actual=db.Column(db.Integer)
    test_case_execution_target=db.Column(db.Integer)
    test_case_execution_actual=db.Column(db.Integer)
    defects_found_target=db.Column(db.Integer)
    defects_found_actual=db.Column(db.Integer)
    test_scripts_creation_target=db.Column(db.Integer)
    test_scripts_creation_actual=db.Column(db.Integer)
    test_scripts_updation_target=db.Column(db.Integer)
    test_scripts_updation_actual=db.Column(db.Integer)
    test_scripts_execution_target=db.Column(db.Integer)
    test_scripts_execution_actual=db.Column(db.Integer)
    site_Scrub_target=db.Column(db.Integer)
    site_Scrub_actual=db.Column(db.Integer)
    project_doc_target=db.Column(db.Integer)
    project_doc_actual=db.Column(db.Integer)
    internal_Review_target=db.Column(db.Integer)
    internal_Review_actual=db.Column(db.Integer)
    regression_cycle_target=db.Column(db.Integer)
    regression_cycle_actual=db.Column(db.Integer)
    req_anal_target=db.Column(db.Integer)
    req_anal_actual=db.Column(db.Integer)
    end_cases_exec_target=db.Column(db.Integer)
    end_cases_exec_actual=db.Column(db.Integer)
    task_coverage_score_target=db.Column(db.Integer)
    task_coverage_score_actual=db.Column(db.Integer)
    assessment_score_target=db.Column(db.Integer)
    assessment_score_actual=db.Column(db.Integer)
    assessment_re_score_target=db.Column(db.Integer)
    assessment_re_score_actual=db.Column(db.Integer)
    cert_score_target=db.Column(db.Integer)
    cert_score_actual=db.Column(db.Integer)
    cert_re_score_target=db.Column(db.Integer)
    cert_re_score_actual=db.Column(db.Integer)
    new_features_imp_target=db.Column(db.Integer)
    new_features_imp_actual=db.Column(db.Integer)
    defects_fixed_target=db.Column(db.Integer)
    defects_fixed_actual=db.Column(db.Integer)
    enhancements_target=db.Column(db.Integer)
    enhancements_actual=db.Column(db.Integer)
    fig_desgns_target=db.Column(db.Integer)
    fig_desgns_actual=db.Column(db.Integer)
    doc_update_target=db.Column(db.Integer)
    doc_update_actual=db.Column(db.Integer)
    research_target=db.Column(db.Integer)
    research_actual=db.Column(db.Integer)
    inv_defs=db.Column(db.Integer)
    spel_errors=db.Column(db.Float)
    client_esc=db.Column(db.Integer)
    tst_cases_missing=db.Column(db.Integer)
    att=db.Column(db.Integer)
    dtouch=db.Column(db.Integer)
    new_init=db.Column(db.Integer)
    defects_verification_target=db.Column(db.Integer)
    defects_verification_actual=db.Column(db.Integer)
    target=db.Column(db.Integer)
    actual=db.Column(db.Integer)
    production=db.Column(db.Integer)
    quality=db.Column(db.Integer)
    attendance=db.Column(db.Integer)
    skill=db.Column(db.Integer)
    new_initiatives=db.Column(db.Integer)
    Dmax_score=db.Column(db.Integer)
    

# Employee Model
class Employee(db.Model):
    id = db.Column(db.Integer, primary_key=True)
    emp_id = db.Column(db.String(100), unique=True, nullable=False)
    password = db.Column(db.String(200), nullable=False)
    role = db.Column(db.String(100), nullable=False)
    email= db.Column(db.String(100), nullable=False)
    name= db.Column(db.String(100), nullable=False)
    is_approved = db.Column(db.Boolean, default=False, nullable=False)

class Employee_information(db.Model):
    __bind_key__="emp_info"
    id = db.Column(db.Integer, primary_key=True)
    emp_name = db.Column(db.String(100),nullable=False)
    emp_id = db.Column(db.String(100),nullable=False)
    emp_email = db.Column(db.String(100),nullable=False)
    emp_date = db.Column(db.String(100),nullable=False)
    emp_project = db.Column(db.String(100),nullable=False)
    emp_designation = db.Column(db.String(100),nullable=False)
    test_case_creation_target= db.Column(db.Integer)
    test_case_updation_target=db.Column(db.Integer)
    test_case_execution_target=db.Column(db.Integer)
    defects_found_target=db.Column(db.Integer)
    test_scripts_creation_target=db.Column(db.Integer)
    test_scripts_updation_target=db.Column(db.Integer)
    test_scripts_execution_target=db.Column(db.Integer)
    site_Scrub_target=db.Column(db.Integer)
    project_doc_target=db.Column(db.Integer)
    internal_Review_target=db.Column(db.Integer)
    regression_cycle_target=db.Column(db.Integer)
    req_anal_target=db.Column(db.Integer)
    end_cases_exec_target=db.Column(db.Integer)
    task_coverage_score_target=db.Column(db.Integer)
    assessment_score_target=db.Column(db.Integer)
    assessment_re_score_target=db.Column(db.Integer)
    cert_score_target=db.Column(db.Integer)
    cert_re_score_target=db.Column(db.Integer)
    new_features_imp_target=db.Column(db.Integer)
    defects_fixed_target=db.Column(db.Integer)
    enhancements_target=db.Column(db.Integer)
    fig_desgns_target=db.Column(db.Integer)
    doc_update_target=db.Column(db.Integer)
    research_target=db.Column(db.Integer)
    inv_defs=db.Column(db.Integer)
    spel_errors=db.Column(db.Float)
    client_esc=db.Column(db.Integer)
    tst_cases_missing=db.Column(db.Integer)
    att=db.Column(db.Integer)
    dtouch=db.Column(db.Integer)
    new_init=db.Column(db.Integer)
    defects_verification_target=db.Column(db.Integer)
    reporting_manager=db.Column(db.String(100))
    actual_reporting_manager=db.Column(db.String(100))
    target_month=db.Column(db.String(100))
    
class Target_columns(db.Model):
    __bind_key__="target_columns"
    id = db.Column(db.Integer, primary_key=True)
    emp_id = db.Column(db.String(100),nullable=False)
    test_case_creation_target= db.Column(db.Integer)
    test_case_updation_target=db.Column(db.Integer)
    test_case_execution_target=db.Column(db.Integer)
    defects_found_target=db.Column(db.Integer)
    test_scripts_creation_target=db.Column(db.Integer)
    test_scripts_updation_target=db.Column(db.Integer)
    test_scripts_execution_target=db.Column(db.Integer)
    site_Scrub_target=db.Column(db.Integer)
    project_doc_target=db.Column(db.Integer)
    internal_Review_target=db.Column(db.Integer)
    regression_cycle_target=db.Column(db.Integer)
    req_anal_target=db.Column(db.Integer)
    end_cases_exec_target=db.Column(db.Integer)
    task_coverage_score_target=db.Column(db.Integer)
    assessment_score_target=db.Column(db.Integer)
    assessment_re_score_target=db.Column(db.Integer)
    cert_score_target=db.Column(db.Integer)
    cert_re_score_target=db.Column(db.Integer)
    new_features_imp_target=db.Column(db.Integer)
    defects_fixed_target=db.Column(db.Integer)
    enhancements_target=db.Column(db.Integer)
    fig_desgns_target=db.Column(db.Integer)
    doc_update_target=db.Column(db.Integer)
    research_target=db.Column(db.Integer)
    inv_defs=db.Column(db.Integer)
    spel_errors=db.Column(db.Float)
    client_esc=db.Column(db.Integer)
    tst_cases_missing=db.Column(db.Integer)
    att=db.Column(db.Integer)
    dtouch=db.Column(db.Integer)
    new_init=db.Column(db.Integer)
    defects_verification_target=db.Column(db.Integer)
    target_month=db.Column(db.String(100))
    target_year=db.Column(db.String(100))
    status=db.Column(db.String(100))

class OperationalExcellence(db.Model):
    __bind_key__="op_excellence"
    
    id = db.Column(db.Integer, primary_key=True)
    emp_id = db.Column(db.String(100), nullable=False, unique=True)
    month=db.Column(db.String(100), nullable=False)
    dtouch_score = db.Column(db.Float,default=0.0)
    new_init_score = db.Column(db.Float,default=0.0)
    start_date = db.Column(db.String(100))
    end_date = db.Column(db.String(100))

class ProjectTargets(db.Model):
    __bind_key__="project_targets"
    
    id = db.Column(db.Integer, primary_key=True)
    Project = db.Column(db.String(100), nullable=False, unique=True)   
    Lead = db.Column(db.String(100), nullable=False)
    ApprovalManager= db.Column(db.String(100), nullable=False)

class DmaxApprovals(db.Model):
    __bind_key__="dmax_approval"
    id = db.Column(db.Integer, primary_key=True)
    employee_email = db.Column(db.String, nullable=False)
    approved_month = db.Column(db.Integer, nullable=False)  # Month (1-12)
    approved_year = db.Column(db.Integer, nullable=False)   # Year (2024, etc.)
    status= db.Column(db.String(100), nullable=False)

###############################################  app context processors ###############################################

@app.context_processor
def custom_global_variable():
    role = None

    # Check if the user is logged in (email is in session)
    if 'email' in session:
        user_email = session['email']

        # Query the user role from the database
        user = Employee.query.filter_by(email=user_email).first()
        if user:
            role = user.role  # Assuming `role` is a column in your User model
            print("role",role)

    if 'username' in session :
        username = session['username']
        user = Employee.query.filter_by(emp_id=username).first()
        if user:
            role=user.role
            print("role",user)         
    # Return the role to all templates as a global variable
    return {'user_role': role}

@app.template_filter('get_attr')
def get_attr(obj, attr):
    """Fetches an attribute from an object safely."""
    return getattr(obj, attr, 'N/A')



###############################################  app routes ###############################################

@app.route('/form',methods=["GET","POST"])
def home():
    if 'username' in session :
        username = session['username']
        employee = Employee.query.filter_by(emp_id=username).first()
        if employee:
            role = employee.role
            
    elif 'email' in session:
        email=session['email']
        employee=Employee.query.filter_by(email=email).first()
        if employee:
            role=employee.role
                              
    else:
       
        return redirect(url_for('sign'))
        

    if request.method=="POST":
        
        workbook = load_workbook(excel_path)
        sheet = workbook.active
        
        next_row = find_next_available_row(sheet)
         
        field_to_column = {
            "employee_name": 'A',
            "employee_id": 'B',
            "employee_email": 'C',
            "today_date": 'D',
            "project": 'E',
            "designation": 'F',
            "test_case_creation_target": 'G',
            "test_case_creation_actual": 'H',
            "test_case_updation_target": 'I',
            "test_case_updation_actual": 'J',
            "test_case_execution_target": 'K',
            "test_case_execution_actual": 'L',
            "defects_found_target":'M',
            "defects_found_actual":'N',
            "defects_verification_target":'O',
            "defects_verification_actual":'P',
            "test_scripts_creation_target":'Q',
            "test_scripts_creation_actual":'R',
            "test_scripts_updation_target":'S',
            "test_scripts_updation_actual":'T',
            "test_scripts_execution_target":'U',
            "test_scripts_execution_actual":'V',
            "site_Scrub_target":'AG',
            "site_Scrub_actual":'AH',
            "project_doc_target":'W',
            "project_doc_actual":'X',
            "internal_Review_target":'Y',
            "internal_Review_actual":'Z',
            "regression_cycle_target":'AA',
            "regression_cycle_actual":'AB',
            "req_anal_target":'AC',
            "req_anal_actual":'AD',
            "end_cases_exec_target":'AE',
            "end_cases_exec_actual":'AF',
            "task_coverage_score_target":'AI',
            "task_coverage_score_actual":'AJ',
            "assessment_score_target":'AK',
            "assessment_score_actual":'AL',
            "assessment_re_score_target":'AM',
            "assessment_re_score_actual":'AN',
            "cert_score_target":"AO",
            "cert_score_actual":'AP',
            "cert_re_score_target":'AQ',
            "cert_re_score_actual":'AR',
            "new_features_imp_target":'AS',
            "new_features_imp_actual":'AT',
            "defects_fixed_target":'AU',
            "defects_fixed_actual":'AV',
            "enhancements_target":'AW',
            "enhancements_actual":'AX',
            "fig_desgns_target":'AY',
            "fig_desgns_actual":'AZ',
            "doc_update_target":'BA',
            "doc_update_actual":'BB',
            "research_target":'BC',
            "research_actual":'BD',
            "inv_defs":'BE',
            "spel_errors":'BF',
            "client_esc":'BG',
            "tst_cases_missing":'BH',
            "att":'BI',
            "dtouch":'BJ',
            "new_init":'BK',    
        }
        form_data = {}
        row_values = []
        for field, column in field_to_column.items():
            value = request.form.get(field)
            value = value.strip() if value else ''
            if value and value.replace('.', '', 1).isdigit():
                value = float(value)
            form_data[field] = value
            # sheet[f'{column}{next_row}'] = value
            # row_values.append(value)
        # dictionary for mapping actual to target    
        actual_to_target_mapping = {}

        for key in field_to_column.keys():
            if key.endswith('_actual'):
                target_key = key.replace('_actual', '_target')  # Replace '_actual' with '_target'
                if target_key in field_to_column:  # Check if target_key exists
                    actual_to_target_mapping[key] = target_key   
        
        results = {}
        employee_id = form_data['employee_id']
        form_today_date = datetime.strptime(form_data['today_date'], '%Y-%m-%d')
        existing_entry = Dform.query.filter_by(today_date=form_data['today_date'], employee_email=form_data['employee_email']).first()
        if existing_entry:
            
            flash("An entry for this date already exists!", "warning")
            return redirect(request.referrer)
        month = form_today_date.month
        previous_month = month - 1 if month > 1 else 12
        month_str=str(month).zfill(2)
        previous_month_str = str(previous_month).zfill(2)
        first_entry = not db.session.query(Dform).filter_by(employee_id=employee_id).first()
        form_today_date = datetime.strptime(form_data['today_date'], '%Y-%m-%d')
        month = form_today_date.month
        year = form_today_date.year
        previous_month = month - 1 if month > 1 else 12
        previous_year = year if month > 1 else year - 1
        if first_entry:
            # If first entry, only allow submission for the current month
            if month != datetime.today().month or year != datetime.today().year:
                flash("You can only submit the form for the current month as this is your first entry.", "error")
                return redirect(request.referrer)
        else:
            if month != datetime.today().month or year != datetime.today().year:
                if month > 1:
                    previous_approval = db.session.query(DmaxApprovals).filter_by(
                        employee_email=form_data['employee_id'],
                        approved_month=previous_month,
                        approved_year=previous_year,
                        status="Approved"
                    ).first()
                    if not previous_approval:
                        flash(f"Please get {monthsDict[previous_month_str]}'s data approved before proceeding.", "error")
                        return redirect(request.referrer)
        # if not first_entry and not has_previous_month_entry(employee_id,month):
        #     flash(f"Please complete current month before proceeding to {monthsDict[month_str]}", "error")
        #     return redirect(request.referrer)
        designation = form_data.get("designation", "")
        attendance_input = form_data.get("att", 0)
        if designation == "Intern":
            attendance = int((attendance_input * 10 / 100) * 100)
        elif designation == "Jr.QA Engineer":
            attendance = int((attendance_input * 10 / 100) * 100)
        elif designation == "QA Engineer":   
            attendance = int((attendance_input * 5 / 100) * 100) 
        elif designation=="Sr.QA Engineer":
            attendance = int((attendance_input * 5 / 100) * 100)
        elif designation=="QA Lead":
            attendance = int((attendance_input * 5 / 100) * 100)  
        else:
            attendance = 0    
        results['BP'] = attendance      
        # Initialize the 'BL' sum as 0
        operational_excellence=OperationalExcellence.query.filter_by(emp_id=form_data["employee_id"]).first()
        if operational_excellence and operational_excellence.start_date and operational_excellence.end_date:
            start_date = datetime.strptime(operational_excellence.start_date, "%Y-%m-%d")
            end_date = datetime.strptime(operational_excellence.end_date, "%Y-%m-%d")
            
            # Convert form_data['today_date'] to a datetime object
            today_date = datetime.strptime(form_data['today_date'], "%Y-%m-%d")
            if start_date <= today_date <= end_date:

            # results['BP']=operational_excellence.attendance_score
                results['BQ'] = operational_excellence.dtouch_score
                results['BR'] = operational_excellence.new_init_score
                print("startdate",start_date, "enddate",end_date,"today_date",today_date)
            else:
                  results['BQ'] = 0
                  results['BR'] = 0  
        else:
            # results['BP']=0   
            results['BQ'] = 0
            results['BR'] = 0 
        print(results['BP'])     
        results['BL'] = 0
        results['BM'] = 0
        # Loop through the actual-to-target mapping and apply the formula
        for actual_field, target_field in actual_to_target_mapping.items():
            if form_data[actual_field] > 0:
                results['BL'] +=int(form_data[target_field])
                results['BM'] += int(form_data[actual_field])  
        if results['BM'] != 0 and results['BL'] != 0:  # Check if both BM and BL are not zero
            if designation == "Intern":
                results['BN'] = ((results['BM'] / results['BL']) * 30 / 100) * 100
            elif designation == "Sr.QA Engineer":
                results['BN'] = ((results['BM'] / results['BL']) * 30 / 100) * 100
            elif designation == "Jr.QA Engineer":
                results['BN'] = ((results['BM'] / results['BL']) * 40 / 100) * 100
            elif designation == "QA Engineer":
                results['BN'] = ((results['BM'] / results['BL']) * 35 / 100) * 100
            elif designation == "QA Lead":
                results['BN'] = ((results['BM'] / results['BL']) * 20 / 100) * 100
            else:
                results['BN'] = 0 
        else:
            results['BN'] = 0     
        if form_data['client_esc'] == 1:  # Check if BG (Client Escalations) is 1
            results['BO'] = 0  # Set BO to 0 if BG is 1
        else:
            if results['BN']==0:
                results['BO']=0
            else:    
                sum_invalid_defects_to_test_cases = (
                    form_data['inv_defs'] +  # BE: Invalid Defects
                    form_data['spel_errors'] +  # BF: Spelling Errors
                    form_data['client_esc'] +  # BG: Client Escalations
                    form_data['tst_cases_missing']  # BH: Test Cases Missing
                )       
                results['BO'] = ((100 - sum_invalid_defects_to_test_cases) * 0.4 / 100) * 100
            # results['BP'] = int((form_data['att'] * 1 * 10 / 100) * 100) 
            # results['BP']=0
            # results['BQ'] = int(((form_data['dtouch'] * 10 / 100 / 100) * 100)*100)
            # results['BQ'] = 0
            # results['BR'] =int(((form_data['new_init'] * 10 / 100 / 100) * 100)*100)  
            # results['BR'] = 0   
            results['BS'] = sum(
                                results[key] for key in [ 'BN', 'BO', 'BP', 'BQ', 'BR']
                            )

        if results['BP']==0:
            results['BN']=0
            results['BO']=0
            results['BS']=0
            results['BQ']=0
            results['BR']=0

        
            

        new_entry = Dform(
            employee_name=form_data['employee_name'],
            employee_id=form_data['employee_id'],
            employee_email=form_data['employee_email'],
            today_date=form_data['today_date'],
            project=form_data['project'],
            designation=form_data['designation'],
            test_case_creation_target=form_data.get('test_case_creation_target'),
            test_case_creation_actual=form_data.get('test_case_creation_actual'),
            test_case_updation_target=form_data.get('test_case_updation_target'),
            test_case_updation_actual=form_data.get('test_case_updation_actual'),
            test_case_execution_target=form_data.get('test_case_execution_target'),
            test_case_execution_actual=form_data.get('test_case_execution_actual'),
            defects_found_target=form_data.get('defects_found_target'),
            defects_found_actual=form_data.get('defects_found_actual'),
            test_scripts_creation_target=form_data.get('test_scripts_creation_target'),
            test_scripts_creation_actual=form_data.get('test_scripts_creation_actual'),
            test_scripts_updation_target=form_data.get('test_scripts_updation_target'),
            test_scripts_updation_actual=form_data.get('test_scripts_updation_actual'),
            test_scripts_execution_target=form_data.get('test_scripts_execution_target'),
            test_scripts_execution_actual=form_data.get('test_scripts_execution_actual'),
            site_Scrub_target=form_data.get('site_Scrub_target'),
            site_Scrub_actual=form_data.get('site_Scrub_actual'),
            project_doc_target=form_data.get('project_doc_target'),
            project_doc_actual=form_data.get('project_doc_actual'),
            internal_Review_target=form_data.get('internal_Review_target'),
            internal_Review_actual=form_data.get('internal_Review_actual'),
            regression_cycle_target=form_data.get('regression_cycle_target'),
            regression_cycle_actual=form_data.get('regression_cycle_actual'),
            req_anal_target=form_data.get('req_anal_target'),
            req_anal_actual=form_data.get('req_anal_actual'),
            end_cases_exec_target=form_data.get('end_cases_exec_target'),
            end_cases_exec_actual=form_data.get('end_cases_exec_actual'),
            task_coverage_score_target=form_data.get('task_coverage_score_target'),
            task_coverage_score_actual=form_data.get('task_coverage_score_actual'),
            assessment_score_target=form_data.get('assessment_score_target'),
            assessment_score_actual=form_data.get('assessment_score_actual'),
            assessment_re_score_target=form_data.get('assessment_re_score_target'),
            assessment_re_score_actual=form_data.get('assessment_re_score_actual'),
            cert_score_target=form_data.get('cert_score_target'),
            cert_score_actual=form_data.get('cert_score_actual'),
            cert_re_score_target=form_data.get('cert_re_score_target'),
            cert_re_score_actual=form_data.get('cert_re_score_actual'),
            new_features_imp_target=form_data.get('new_features_imp_target'),
            new_features_imp_actual=form_data.get('new_features_imp_actual'),
            defects_fixed_target=form_data.get('defects_fixed_target'),
            defects_fixed_actual=form_data.get('defects_fixed_actual'),
            defects_verification_target=form_data.get('defects_verification_target'),
            defects_verification_actual=form_data.get('defects_verification_actual'),
            enhancements_target=form_data.get('enhancements_target'),
            enhancements_actual=form_data.get('enhancements_actual'),
            fig_desgns_target=form_data.get('fig_desgns_target'),
            fig_desgns_actual=form_data.get('fig_desgns_actual'),
            doc_update_target=form_data.get('doc_update_target'),
            doc_update_actual=form_data.get('doc_update_actual'),
            research_target=form_data.get('research_target'),
            research_actual=form_data.get('research_actual'),
            inv_defs=form_data.get('inv_defs'),
            spel_errors=form_data.get('spel_errors'),
            client_esc=form_data.get('client_esc'),
            tst_cases_missing=form_data.get('tst_cases_missing'),
            att=form_data.get('att'),
            dtouch=form_data.get('dtouch'),
            new_init=form_data.get('new_init'),
            target=results['BL'],
            actual=results['BM'],
            production=results['BN'],
            quality=results['BO'],
            
            # Attendance (BO) and Skill (BP)
            attendance=results['BP'],
            skill=results['BQ'],
            
            # New Initiatives (BQ) and Dmax Score (BS)
            new_initiatives=results['BR'],
            Dmax_score=results['BS'],
        )
        
        # Add to DB and commit the session
        db.session.add(new_entry)
        db.session.commit()
        flash("Form submitted sucessfully!","success")
        
        return redirect(url_for('home'))
    return render_template('index.html',role=role)

  
    
@app.route('/',methods=["GET","POST"])
def sign():
    
    if request.method=="POST":
        username = request.form["username"]  
        password = request.form["password"]
        employee = Employee.query.filter_by(emp_id=username).first()
        if employee:
            if not employee.is_approved:  
                flash("You do not have access please contact admin", "danger")
                
                return redirect(url_for('sign'))  

        if employee and employee.password == password:
            session.clear()
            
            session['username'] = username
            role=employee.role
            if role=="super_admin":
                return redirect(url_for('dmax_table'))
            if role=="admin":
                return redirect(url_for('dmax_table'))
            if role=="manager":
                return redirect(url_for('home'))
            if role=="crewmate":
                return redirect(url_for('view_dscore'))

        else:
            flash("Invalid credentials. Please try again!", "danger")
        
        
    return render_template('sign.html')     

@app.route('/view_dmax')
def view_dmax():
    return "hello"
@app.route('/login')
def login():
    flow.redirect_uri = REDIRECT_URI  
    authorization_url, state = flow.authorization_url()
    session['state'] = state
    
    return redirect(authorization_url) 

@app.route('/call_back')
def google_sign_in():
     
    flow.fetch_token(authorization_response=request.url)
    
    # Store the credentials in the session
    credentials = flow.credentials
    
    session['credentials'] = credentials_to_dict(credentials)
    
    credentials = Credentials.from_authorized_user_info(session['credentials'])
    response = requests.get('https://www.googleapis.com/oauth2/v3/userinfo', headers={'Authorization': f'Bearer {credentials.token}'})
    if response.status_code == 200:
        user_info = response.json()
        
        user_email = user_info.get('email')  # Extract email from the response
        employee = Employee.query.filter_by(email=user_email).first()
        
        if not employee:
            flash("Invalid credentials. Please try again!", "danger")
            return redirect(url_for('sign'))
        # Store the email in the session or perform any other actions
        
        if employee:
            if not employee.is_approved:  
                flash("You do not have access please contact admin", "danger")
                    
                return redirect(url_for('sign')) 
               
        session['email'] = user_email
        if employee.role == 'crewmate':
            return redirect(url_for('view_dscore'))
        if employee.role=='super_admin' or employee.role=="admin":
            return redirect(url_for('dmax_table'))
        
    else:
        print("Failed to fetch user info")
        
    return redirect(url_for('home'))
    
@app.route('/read_excel')
def read_excel():
    # Construct the path to the Excel file in the static folder
    

    # Load the Excel workbook
    
    
    # Select the active sheet
    sheet = workbook.active
    
    # Example: Reading data from the first row, first column (A1)
    first_cell_value = sheet['A3'].value

    # Optionally: Process the data further and return it to the template
    return f"Value in A1: {first_cell_value}"


@app.route('/search', methods=['POST'])
def search_employee():
    month_order = case(
        {month: i for i, (month, _) in enumerate(monthsDict_2.items(), 1)},  # Map month names to numeric values (1 for "January", 2 for "February", etc.)
        value=Target_columns.target_month,  # The month column in your Target_columns table
        else_=0  # Default value for non-matching months (if any)
    )
    data = request.json
    employee_name = data.get('employee_name')
    logged_in_user = get_logged_in_user_details()
    logged_in_user_role = logged_in_user.get("role")
    logged_in_user_name = logged_in_user.get("name")
    current_date = datetime.now().date()
    current_month_day = (current_date.month, current_date.day)
    print("current_date",current_month_day)
    current_month=datetime.now().strftime('%B')
    current_year = str(datetime.now().year)
    target_month = None
    for month in range(1, 13):  # Loop through all months
        start_date, end_date = get_date_range_for_month(month)
        
        # Convert start_date and end_date (strings) to datetime.date objects
        start_date = datetime.strptime(start_date, "%Y-%m-%d").date()
        end_date = datetime.strptime(end_date, "%Y-%m-%d").date()
        start_month_day = (start_date.month, start_date.day)
        end_month_day = (end_date.month, end_date.day)
        print("start_date",start_month_day)
        print("end_date",end_month_day)
        if start_month_day <= end_month_day:
            # Normal range within the same year
            if start_month_day <= current_month_day <= end_month_day:
                target_month = str(month).zfill(2)  # Store the matching month (e.g., "02")
                print("target_month", target_month)
                break
        else:
            # Range spans across the end of the year
            if current_month_day >= start_month_day or current_month_day <= end_month_day:
                target_month = str(month).zfill(2)  # Store the matching month (e.g., "02")
                print("target_month", target_month)
                break
    if target_month:
        target_month_string = monthsDict.get(target_month, "Unknown Month")
        
    employees_list = []
    # Query the Login table where employee_name matches (case-insensitive search)
    matched_employees = Employee_information.query.filter(Employee_information.emp_name.ilike(f'%{employee_name}%')).all() 
    
    for emp in matched_employees:
        print(f"Processing employee: {emp.emp_name} (emp_id: {emp.emp_id})")   
    # Create a list of dictionaries containing employee details
        if emp.reporting_manager == logged_in_user_name:
            # Include employee details in the response if the employee's reporting manager matches
            employee_details = {column.name: getattr(emp, column.name) for column in Employee_information.__table__.columns}
            if "emp_designation" in employee_details:
                employee_details["emp_designation"] = corrections.get(employee_details["emp_designation"], employee_details["emp_designation"])
            matched_targets = Target_columns.query.filter(
                Target_columns.emp_id == emp.emp_id,
                Target_columns.status == "approved",
                Target_columns.target_month == current_month,
                Target_columns.target_year == current_year
            ).first()
            print(f"Debug: matched_targets for {emp.emp_name} = {matched_targets}")
            if matched_targets:
                # Update employee details with target values
                for column in Target_columns.__table__.columns:
                    if column.name.endswith('_target'):
                        employee_details[column.name] = getattr(matched_targets, column.name)
            if not matched_targets:
                # Set all target fields to 0 if no matched targets are found
                if "error" not in employee_details:  # Prevents duplicate execution
                    employee_details["error"] = f"Set targets first for {current_month} {current_year}"        
                    
            employees_list.append(employee_details)
    
    return jsonify({"employees": employees_list})

@app.route('/register', methods=['GET','POST'])
def register():
    if request.method=="POST":
        username=request.form["username"]
        password=request.form["password"]
        role=request.form["role"]
        email=request.form["email"]
        name=request.form["name"]
        existing_employee = Employee.query.filter(
            or_(
                Employee.emp_id == username,
                Employee.email == email,
                Employee.name == name
            )
        ).first()
        if existing_employee:
            if existing_employee.emp_id == username:
                flash('Employee with that ID already exists', "danger")
            elif existing_employee.email == email:
                flash('Employee with that email already exists', "danger")
            elif existing_employee.name == name:
                flash('Employee with that name already exists', "danger")
            return redirect(url_for('register'))
        new_employee = Employee(emp_id=username, password=password, role=role,email=email,name=name)
        db.session.add(new_employee)
        db.session.commit()
        flash('Employee registered successfully!', 'success')
        
        return redirect(url_for('sign'))
    return render_template("register.html")

@app.route("/no-acess")
def no_access():
    return render_template("no_acess.html")

@app.route("/logout")
def logout():
    session.pop('username',None)
    return redirect(url_for('sign'))

@app.route("/employee_info",methods=["GET","POST"])
def employee_info():
        

    employees = Employee_information.query.all()

    return render_template("employee_info.html",employees=employees)

@app.route('/edit_employee/<int:employee_id>', methods=['GET', 'POST'])
def edit_employee(employee_id):
    # Retrieve employee data from the database based on employee_id
    
    employee =Employee_information.query.get(employee_id)
    if employee:
        employee.emp_designation = corrections.get(employee.emp_designation, employee.emp_designation)
    projects=['Akyrian','Auxo','Avanti','Bench','Fora Travels','Indihood','IPS','IQHive','LevelBlue','Web Development','Opus Clip','Training']
    designations=['Intern','Jr.QA Engineer','QA Engineer','Sr.QA Engineer','QA Lead']
    if request.method == 'POST':
        print
        # Update the employee data with form values
        employee.emp_name = request.form['emp_name']
        employee.emp_id=request.form['emp_id']
        employee.emp_email = request.form['emp_email']
        employee.emp_project = request.form['emp_project']
        employee.emp_designation = corrections.get(request.form['emp_designation'], request.form['emp_designation'])
        employee.reporting_manager=request.form['rep_manager']
        # Save the updated data back to the database
        
        db.session.commit()
        
        return redirect(url_for('employee_info'))
        
    # Render the edit form with existing values
    return render_template('edit_employee.html', employee=employee,projects=projects,designations=designations)



@app.route("/employee_upload" ,methods=['GET', 'POST'])
def employee_upload():
    if request.method=="POST":
        try:
            file = request.files['file']  # Get the uploaded file

            # Load the workbook directly from the file object
            wb = load_workbook(file)  # No need for BytesIO here
            ws = wb.active
            headers = [cell.value for cell in ws[1]]
            column_mapping = {
                "Employee Name":"emp_name",	
                "Employee ID":	"emp_id",
                "Employee Email":"emp_email", 
                "Today's Date":"emp_date",	
                "Select your Project":"emp_project",	
                "Designation":"emp_designation",	 
                "Lead/Spocs":"reporting_manager",
                "Reporting manager":"actual_reporting_manager"
            # Add more mappings as per your Excel file
            }
            required_fields = list(column_mapping.values())
            mapped_columns = {column_mapping[h]: idx for idx, h in enumerate(headers) if h in column_mapping}
            
            processed_employees = set()
            # Process the rows and add to the database
            for row in ws.iter_rows(min_row=2, values_only=True):
                
                employee_data = {db_field: row[idx] for db_field, idx in mapped_columns.items()}
                if all(value is None for value in employee_data.values()):
                    continue  
                filled_fields = [field for field in required_fields if employee_data.get(field)]

                if 0 < len(filled_fields) < len(required_fields):  # If only some required fields are filled
                    flash(f"Row skipped: Incomplete data provided. Please fill all required fields.", "danger")
                    continue 
                if "emp_designation" in employee_data:
                    designation = employee_data["emp_designation"]				                    
                    
                    # Standardizing incorrect spellings
                    corrections = {
                        "SrQAEngineer": "Sr.QA Engineer",
                        "Sr QA Engineer": "Sr.QA Engineer",
                        "JrQAEngineer": "Jr.QA Engineer",
                        "Jr QA Engineer": "Jr.QA Engineer"
                    }
                    
                    # If the designation needs correction, apply it
                    
                    emp_email = employee_data.get("emp_email", "")
                    emp_id = employee_data.get("emp_id", "")
                    emp_name=employee_data.get("emp_name", "")            
                                    
                    existing_employee_data = (
                        db.session.query(Employee_information)
                        .filter(
                            or_(
                                func.lower(Employee_information.emp_email) == emp_email.lower(),
                                Employee_information.emp_id == emp_id,
                                func.lower(Employee_information.emp_name) == emp_name.lower(),
                            )
                        )
                        .first()
                    )

                    if existing_employee_data:
                        flash(f"Employee with email '{emp_email}', ID '{emp_id}', or name '{emp_name}' already exists. Skipping entry.", "danger")
                        continue 
                    employee_data["emp_designation"] = corrections.get(designation, designation)
                if 'emp_date' in employee_data:
                    emp_date = employee_data['emp_date']

                    if isinstance(emp_date, datetime):
                        emp_date = emp_date.date()  # Removes the time and keeps only the date
                        print("Date without time:", emp_date)

                    # If emp_date is a string, convert it to a date object
                    elif isinstance(emp_date, str):
                        try:
                            emp_date = datetime.strptime(emp_date, '%Y-%m-%d').date()  # Convert to date
                            print("Successfully parsed the date:", emp_date)
                        except ValueError:
                            print("Incorrect date format")
                            emp_date = None   
                   
                emp_name = employee_data.get("emp_name")
                emp_email = employee_data.get("emp_email")
                actual_reporting_manager = employee_data.get("actual_reporting_manager")
                if not emp_name or not emp_email:
                    continue
                        
                if 'emp_name' in employee_data:
                    emp_name = employee_data['emp_name']
                    actual_reporting_manager = employee_data.get('actual_reporting_manager')
                    name_exists = False 

                    if emp_name and emp_date and actual_reporting_manager:
                        # Extract year and month from the new employee entry
                        new_year, new_month = emp_date.year, emp_date.month

                        # Query the most recent entry of the employee
                        last_entry = (
                            db.session.query(Employee_information)
                            .filter(func.lower(Employee_information.emp_name) == emp_name.lower())
                            .order_by(Employee_information.emp_date.desc())  # Get most recent entry
                            .first()
                        )

                        if last_entry:
                            last_date = datetime.strptime(last_entry.emp_date, "%Y-%m-%d").date()
                            last_year, last_month = last_date.year, last_date.month
                            last_manager = last_entry.actual_reporting_manager

                            # Check if it's a duplicate entry
                            is_duplicate = (new_year == last_year and new_month == last_month and actual_reporting_manager == last_manager)

                            if is_duplicate:
                                if emp_name not in processed_employees:
                                    flash(f"Employee '{emp_name}' already exists for {new_month}-{new_year} under the same manager. Skipping entry.", "danger")
                                    processed_employees.add(emp_name)
                                continue  # Skip the duplicate entry

                
                    # if emp_date:
                    #     try:
                    #         db.session.query(Employee_information).update({"emp_date": emp_date})
                    #         db.session.commit()
                    #         flash(f"Updated emp_date for all employees to {emp_date}.", "success")
                    #     except Exception as e:
                    #         print(e)
                    #         flash("Failed to update emp_date for existing employees.", "danger")          

                employee = Employee_information(**employee_data)
                # emp_id = employee_data['emp_id']
                # target_month = employee_data['target_month']
                # target_fields = {key: value for key, value in employee_data.items() if key.endswith('_target')}
                # existing_target = db.session.query(Target_columns).filter_by(emp_id=emp_id, target_month=target_month).first()
                # if existing_target:
                #     # Update the existing record
                #     for field, value in target_fields.items():
                #         setattr(existing_target, field, value)
                #     try:
                #         db.session.commit()
                #         flash(f"Updated targets for emp_id '{emp_id}' for the month '{target_month}'.", "success")
                #     except Exception as e:
                #         print(e)
                #         flash(f"Failed to update targets for emp_id '{emp_id}'.", "danger")
                # else:        
                #     new_target = Target_columns(
                #     emp_id=employee_data['emp_id'],
                #     target_month=employee_data['target_month'],
                #     **target_fields  # Dynamically add target fields
                #     )
                #     db.session.add(new_target)
                # Add the employee record to the session
                
                db.session.add(employee)
                
            # # Commit all changes to the database
            db.session.commit() 
            
        except Exception as e:
            print("e",e)   
            flash("Failed to upload employees. Please check the sample file.", "danger")
    return render_template("employee_upload.html")

@app.route("/dmax_table", methods=["GET", "POST"])

def dmax_table():
    default_page_size = 10
    page_size_options = [10, 20, 30, 'All']
    current_month=str(datetime.now().month)
    
    if request.method == "POST":
        selected_page_size = request.form.get("page_size", default_page_size)
        # Set page_size to None for 'All', or convert to int if numeric
        if selected_page_size == 'All':
            session['page_size'] = None
        else:
            session['page_size'] = int(selected_page_size)
        # Redirect to page 1 with the current search term to avoid form resubmission
        return redirect(url_for('dmax_table', page=1, search_term=request.args.get('search_term', '')))
    # designation_counts = (
    #     db.session.query(Employee_information.emp_designation, func.count(Employee_information.emp_designation))
    #     .group_by(Employee_information.emp_designation)
    #     .all()
    # )
    designation_counts = db.session.query(
            Employee_information.emp_designation, 
            func.count(Employee_information.emp_designation)
        )
    
    page_size = session.get('page_size', default_page_size)
    search_term = request.args.get("search_term", "").strip()
    selected_designation=request.args.get("designation")
    selected_project=request.args.get("project")
    selected_month = request.args.get("month", current_month)
    if not selected_month:
        current_month = datetime.now().month
        selected_month = str(current_month)
    page = request.args.get("page", 1, type=int)
    projects=["Akyrian","Auxo","Avanti","Bench","Fora Travels","Indihood","IPS","IQHive","LevelBlue","Web Development","Opus Clip","Training"]
    designations=["Intern","Jr.QA Engineer","QA Engineer","Sr.QA Engineer","QA Lead"]
    query=Dform.query
    
    if search_term:
        query = query.filter(func.lower(Dform.employee_name) == search_term.lower())
        designation_counts = designation_counts.filter(func.lower(Employee_information.emp_name).ilike(f"%{search_term.lower()}%"))
    if selected_designation:
        query=query.filter(Dform.designation == selected_designation)
        designation_counts = designation_counts.filter(Employee_information.emp_designation == selected_designation)
    if selected_project:
        query = query.filter(func.lower(Dform.project) == selected_project.strip().lower())
        designation_counts = designation_counts.filter(Employee_information.emp_project == selected_project)
    if selected_month:
        query=query.filter(extract('month', Dform.today_date) == int(selected_month))  
    
    designation_counts = designation_counts.group_by(Employee_information.emp_designation).all()
    labels = [row[0] for row in designation_counts]  # Designation names
    scores = [row[1] for row in designation_counts]  # Counts
    if page_size:  # If 'All' is not selected, paginate based on page size
        paginated_entries = query.paginate(page=page, per_page=page_size, error_out=False)
    else:
        # Show all results if page size is 'All'
        paginated_entries = query.paginate(page=1, per_page=query.count(), error_out=False)

    # Prepare a list to hold the data for all entries
    data_list = []
    
    # Check if there are any entries
    # if paginated_entries.items:
    #     # Use SQLAlchemy's metadata to get the columns in the order they are defined
    #     columns =['employee_name','employee_id','project','designation','production', 'quality', 'attendance', 'skill','Dmax_score']  # Use the first entry to get column names
        
    #     # Loop through each entry to create a dictionary for each row
    #     for entry in paginated_entries.items:
    #         data = {field: getattr(entry, field) for field in columns}  # Create a dict for each entry
    #         data_list.append(data)  # Add to the list

    top_scores = (
    query.with_entities(Dform.employee_id, Dform.employee_name, func.avg(Dform.Dmax_score).label('avg_score'))
    .group_by(Dform.employee_id, Dform.employee_name)
    .order_by(func.avg(Dform.Dmax_score).desc())  # Order by avg_score in descending order
    .limit(4)  # Limit to top 4
    .all()
    )
    chart_data = [
    {"name": row[1], "score": round(row[2], 2)}  # Use employee_name and avg_score
    for row in top_scores
    ]
    avg_query = query.with_entities(
    Dform.employee_id,
    Dform.employee_name,
    func.avg(Dform.Dmax_score).label('avg_score')
    ).group_by(Dform.employee_id, Dform.employee_name)
    if page_size:  # Paginate if a page size is specified
        paginated_avg_scores = avg_query.paginate(page=page, per_page=page_size, error_out=False)
    else:
        paginated_avg_scores = avg_query.paginate(page=1, per_page=avg_query.count(), error_out=False)
    avg_scores_list = [
        {"employee_id": row[0], "employee_name": row[1], "avg_score": round(row[2], 2)}
        for row in paginated_avg_scores.items
    ]


    return render_template('dmax_table.html', data_list=data_list,pagination=paginated_avg_scores,avg_scores_list=avg_scores_list, search_term=search_term,page_size=page_size,page_size_options=page_size_options,chart_data=chart_data,selected_month=selected_month,selected_designation=selected_designation,selected_project=selected_project,projects=projects,designations=designations,labels=labels, scores=scores)

@app.route("/team_dmax_table", methods=["GET", "POST"])
def team_dmax_table():
    user_details = get_logged_in_user_details()
    if user_details:
        user_name=user_details['name']
        user_email=user_details['email']
        user_name=user_name.lower()
        role=user_details['role']
        selected_month = request.args.get('month')
        selected_year = request.args.get('year')
        selected_year = int(selected_year) if selected_year else None
        current_month = datetime.now().strftime("%m")
        current_year = datetime.now().strftime("%Y")
        selected_project = request.args.get('project')
        print("selected_project",selected_project)
        if not selected_month:
            selected_month = current_month
        if not selected_year:
            selected_year= int(current_year)    
        selected_month_name = monthsDict.get(selected_month)
        employees_under_projects = []
        if role=="admin" or role=="super_admin":
            projects_under_manager = ProjectTargets.query.with_entities(ProjectTargets.Project).filter(
                func.lower(ProjectTargets.ApprovalManager) == user_name.lower()
            ).all()
            project_names = [proj[0] for proj in projects_under_manager]
            if project_names:
                employees_under_manager = Employee_information.query.filter(
                    Employee_information.emp_project.in_(project_names)
                )
            else:
                employees_under_manager = Employee_information.query.filter(False)
        if role=="manager":    
            employees_under_manager = Employee_information.query.filter(func.lower(Employee_information.reporting_manager)==user_name)
        elif role == "crewmate":
            employees_under_manager = Employee_information.query.filter(
                func.lower(Employee_information.emp_email) == func.lower(user_email)
            ) 
            
        if selected_project:
            employees_under_manager = employees_under_manager.filter(
                func.lower(Employee_information.emp_project) == func.lower(selected_project)
            )

        employees_under_manager = employees_under_manager.all()   
        print(employees_under_manager)    
        projects_led_by_user = ProjectTargets.query.filter(func.lower(ProjectTargets.Lead)==user_name)
        if selected_project:
            projects_led_by_user = projects_led_by_user.filter(
                func.lower(ProjectTargets.Project) == func.lower(selected_project)
            )

        projects_led_by_user = projects_led_by_user.all()
        employees_under_actual_approval_manager=Employee_information.query.filter(func.lower(Employee_information.actual_reporting_manager)==user_name)
        if selected_project:
            employees_under_actual_approval_manager = employees_under_actual_approval_manager.filter(
                func.lower(Employee_information.emp_project) == func.lower(selected_project)
            )
        employees_under_actual_approval_manager = employees_under_actual_approval_manager.all()    
        if projects_led_by_user:
            # If the user is a lead, find all employees working on the same project
            for project in projects_led_by_user:
                employees_in_project = Employee_information.query.filter_by(emp_project=project.Project).all()
                employees_under_projects.extend(employees_in_project)
        user_manager  = ProjectTargets.query.filter(func.lower(ProjectTargets.ApprovalManager)==user_name)
        if user_manager:
            if selected_project:
                user_manager = user_manager.filter(
                    func.lower(ProjectTargets.Project) == func.lower(selected_project)
                )
            for project in user_manager:
                employees_in_project = Employee_information.query.filter_by(emp_project=project.Project).all()
                employees_under_projects.extend(employees_in_project)
        all_accessible_employees = {emp.emp_email: emp for emp in employees_under_manager + employees_under_projects  + employees_under_actual_approval_manager}.values()
        all_accessible_emp_ids = [emp.emp_id for emp in all_accessible_employees]
        targets = Target_columns.query.filter(Target_columns.emp_id.in_(all_accessible_emp_ids)).all()
        employees_with_pending_targets = {
            target.emp_id: {"month": target.target_month, "year": target.target_year}  
            for target in targets if target.status == "waiting for approval"
        }
        
        approved_targets = Target_columns.query.filter(
                    Target_columns.emp_id.in_([emp.emp_id for emp in all_accessible_employees]),  # Only check targets for managed employees
                    Target_columns.target_month == selected_month_name,
                    Target_columns.target_year == selected_year,
                    Target_columns.status == "approved"
                ).all()   
        
        
        employees_with_approved_targets = set(target.emp_id for target in approved_targets)
        waiting_for_approval_targets = Target_columns.query.filter(
            Target_columns.emp_id.in_([emp.emp_id for emp in all_accessible_employees]),  
            Target_columns.target_month == selected_month_name,
            Target_columns.target_year == selected_year,
            Target_columns.status == "waiting for approval"
        ).all()
        employees_with_pending_targets = {
            target.emp_id: {"month": target.target_month, "year": target.target_year}
            for target in waiting_for_approval_targets
        }
        
        filtered_employees = []
        for emp in all_accessible_employees:
                filtered_employees.append({
                    "name": emp.emp_name,
                    "project":emp.emp_project,
                    "designation":emp.emp_designation,
                    "date":emp.emp_date,
                    "emp_id":emp.emp_id,
                    "reporting_manager":emp.reporting_manager,
                    "actual_reporting_manager":emp.actual_reporting_manager,
                    "id":emp.id,
                    "role": (
                                "Project Lead" if emp.emp_project in [proj.Project for proj in projects_led_by_user]
                                else ("Approval Manager" if emp.emp_project in [proj.Project for proj in user_manager]
                                else "Employee")
                            ) , # Determine role,
                    "has_approved_target": (emp.emp_id in employees_with_approved_targets) if employees_with_approved_targets else False , # Store whether they have an approved target
                    "has_pending_target": employees_with_pending_targets.get(emp.emp_id, {})        
                                        })
                
        return render_template('team_dmax_table.html',employees=filtered_employees,user_name=user_name,years=last_ten_years,selected_month=selected_month, current_month=current_month,monthsDict=monthsDict,current_year=current_year,selected_year=selected_year,selected_month_name=selected_month_name,projects=projects,selected_project=selected_project)
    
    return "No user found or not logged in."  

@app.route("/view_dscore", methods=["GET", "POST"])
def view_dscore():
    user_name = get_logged_in_user_details()
    ALLOWED_COLUMNS = [
        "id","employee_name","target", "actual", "production","project",
        "quality", "attendance", "skill", "new_initiatives", "Dmax_score","att"
    ]
    current_year=datetime.now().year
    current_month=str(datetime.now().month)
    status_mapping = {
        "approved": "Approved",
        "waiting for approval": "In Process"
    }
    years = [current_year - i for i in range(11)]
    # ALLOWED_COLUMNS = [
    #     "employee_name", "today_date", "test_case_creation_target",
    #     "test_case_creation_actual", "test_case_updation_target", "test_case_updation_actual",
    #     "test_case_execution_target", "test_case_execution_actual", "defects_found_target",
    #     "defects_found_actual","defects_verification_target", "defects_verification_actual", "test_scripts_creation_target", "test_scripts_creation_actual",
    #     "test_scripts_execution_target","test_scripts_execution_actual","test_scripts_updation_target", "test_scripts_updation_actual",
    #     "site_Scrub_target", "site_Scrub_actual", "project_doc_target",
    #     "project_doc_actual", "internal_Review_target", "internal_Review_actual", "regression_cycle_target",
    #     "regression_cycle_actual", "req_anal_target", "req_anal_actual", "end_cases_exec_target",
    #     "end_cases_exec_actual", "task_coverage_score_target", "task_coverage_score_actual",
    #     "assessment_score_target", "assessment_score_actual", "assessment_re_score_target",
    #     "assessment_re_score_actual", "cert_score_target", "cert_score_actual", "cert_re_score_target",
    #     "cert_re_score_actual", "new_features_imp_target", "new_features_imp_actual", "defects_fixed_target",
    #     "defects_fixed_actual", "enhancements_target", "enhancements_actual", "fig_desgns_target",
    #     "fig_desgns_actual", "doc_update_target", "doc_update_actual", "research_target", "research_actual",
    #     "inv_defs", "spel_errors", "client_esc", "tst_cases_missing", "att", "dtouch", "new_init", "target", "actual", "production",
    #     "quality", "attendance", "skill", "new_initiatives", "Dmax_score"
    # ]
    
    if user_name:
        
        role = user_name['role']
        email=user_name['email']
        user_name=user_name['name'].lower()
        search_query = request.args.get('search', '').strip().lower()
        selected_date = request.args.get('date') 
        selected_month = request.args.get('month',current_month)
        
        selected_year=request.args.get('year',current_year)
        if selected_month.isdigit():  # Only convert if it's numeric
            current_month_formatted = datetime.strptime(selected_month, "%m").strftime("%B")
        try:
            selected_month_int = int(selected_month)  # Convert string "2" to integer 2
            current_month_formatted = datetime.strptime(str(selected_month_int), "%m").strftime("%B")  # "2" → "February"
        except ValueError:
            current_month_formatted = datetime.now().strftime("%B")  # Fallback if invalid input
        current_year_formatted = str(selected_year)    
        
        if role =="manager":    
            employees_under_manager = Employee_information.query.filter( or_(
                func.lower(Employee_information.reporting_manager) == user_name,
                func.lower(Employee_information.actual_reporting_manager) == user_name
            )).all()
            is_actual_manager = False
            filtered_employees=[]
            for emp in employees_under_manager:
                matched_employees = get_first_filtered_employees(
                    Dform.query.filter_by(employee_email=emp.emp_email),
                    search_query,
                    selected_month,
                    selected_date,
                    selected_year
                )
                emp_id = db.session.query(Employee_information.emp_id).filter(
                    Employee_information.emp_email == emp.emp_email
                ).first()
                if emp_id:
                    emp_id = emp_id[0]
                    
                # project_status = db.session.query(Target_columns.status).filter(
                #     ProjectTargets.employee_email == emp.emp_email,
                #     ProjectTargets.month == selected_month,
                #     ProjectTargets.year == selected_year
                # ).scalar()
                project_status = db.session.query(Target_columns.status).filter(
                    Target_columns.emp_id == emp_id,  # Query for a single employee
                    Target_columns.target_month == current_month_formatted,
                    Target_columns.target_year == current_year_formatted
                ).first()
                project_status = project_status[0] if project_status else None
                project_status=status_mapping.get(project_status, "Targets not set")
                
                approval_record = db.session.query(DmaxApprovals).filter_by(
                    employee_email=emp.emp_id,
                    approved_month=selected_month,
                    approved_year=selected_year
                ).first()

                if approval_record:
                    approval_status = approval_record.status.capitalize()  # e.g., "Approved", "Rejected", etc.
                else:
                    approval_status = None
                is_actual_manager = emp.reporting_manager.lower() == user_name.lower()
                print(is_actual_manager,approval_status)
                # query = Dform.query.filter_by(employee_email=emp.emp_email)
                # if search_query:
                #     # Use = for exact match (case-sensitive)
                #     query = query.filter(func.lower(Dform.employee_name) == search_query)
                    
                # if selected_month:
                #     # Extract the month from today_date (assuming today_date is a datetime field)
                #     # please import extract in godaddy
                #    query = query.filter(extract('month', Dform.today_date) == int(selected_month))
                # if selected_date:
                #     query = query.filter_by(today_date=selected_date)
                # matched_employees = query.all()    
                if matched_employees.count() > 0:
                    # for matched in matched_employees:
                    #     filtered_employees.append(
                    #         {
                    #             column: getattr(matched, column, None)  # Use getattr to get the attribute dynamically
                    #             for column in ALLOWED_COLUMNS         # Filter by allowed columns
                    #         }
                    #     )
                    averages = get_averages_for_filtered_employees(matched_employees)
                    if averages:
                        
                        filtered_employees.append(
                        {  "id": emp.id,
                            "employee_id": emp.emp_id,
                            "employee_name": emp.emp_name,
                            "target": averages["avg_target"],  # ✅ Correct way to access dictionary values
                            "actual": averages["avg_actual"],
                            "production": averages["avg_production"],
                            "quality": averages["avg_quality"],
                            "attendance": averages["avg_attendance"],
                            "skill": averages["avg_skill"],
                            "new_initiatives": averages["avg_new_initiatives"],
                            "Dmax_score": averages["avg_Dmax_score"],
                            "project":emp.emp_project,
                            "project_status":project_status,
                            "approval_status":approval_status,
                            "is_actual_manager":is_actual_manager
                        }
                        )
                        
            return render_template("view_dscore.html",employees=filtered_employees,role=role,search_query=search_query, selected_month=selected_month, selected_date=selected_date,selected_year=int(selected_year),years=years,is_actual_manager=is_actual_manager)        

        if role=="crewmate":
             
            matched_employees = get_first_filtered_employees(
                Dform.query.filter_by(employee_email=email),
                search_query,
                selected_month,
                selected_date,
                selected_year
            )
            filtered_employees=[]
            if matched_employees.count() > 0:
                first_entry = matched_employees.first()
                if first_entry:
                    averages = get_averages_for_filtered_employees(matched_employees)  # Compute averages ONCE
                    
                    if averages:
                        emp_id = db.session.query(Employee_information.emp_id).filter(
                            Employee_information.emp_email == email
                        ).scalar()
                        project_status = db.session.query(Target_columns.status).filter(
                            Target_columns.emp_id == emp_id,
                            Target_columns.target_month == current_month_formatted,
                            Target_columns.target_year == current_year_formatted
                        ).scalar()
                        project_status=status_mapping.get(project_status, "Targets not set")
                        approval_record = db.session.query(DmaxApprovals).filter_by(
                            employee_email=first_entry.employee_id,  # ✅ Use employee_id instead of emp_id
                            approved_month=selected_month,
                            approved_year=selected_year
                        ).first()
                        if approval_record:
                            approval_status = approval_record.status.capitalize()  # e.g., "Approved", "Rejected", etc.
                        else:
                            approval_status = None
                        filtered_employees.append(
                            {
                                "id": first_entry.id if first_entry else None,  # No employee ID needed for crewmates, or use an appropriate field
                                "employee_name": first_entry.employee_name,
                                "employee_id": first_entry.employee_id,  # Assuming email identifies the crewmate
                                "target": averages["avg_target"],  # ✅ Correct way to access dictionary values
                                "actual": averages["avg_actual"],
                                "production": averages["avg_production"],
                                "quality": averages["avg_quality"],
                                "attendance": averages["avg_attendance"],
                                "skill": averages["avg_skill"],
                                "new_initiatives": averages["avg_new_initiatives"],
                                "Dmax_score": averages["avg_Dmax_score"],
                                "project":first_entry.project,
                                "project_status": project_status,
                                "approval_status":approval_status
                            }
                        )
                    

            return render_template("view_dscore.html",employees=filtered_employees,role=role,search_query=search_query, selected_month=selected_month, selected_date=selected_date,selected_year=int(selected_year),years=years)                
                
        if role == "admin" or role == "super_admin":
            employees_under_manager = Employee_information.query.all()  # Get all employees under the manager
            filtered_employees = []

            for emp in employees_under_manager:
                # Filter Dform by employee email and other search parameters
                matched_employees = get_first_filtered_employees(
                    Dform.query.filter_by(employee_email=emp.emp_email),  # Filter by the employee's email
                    search_query,
                    selected_month,
                    selected_date,
                    selected_year
                )
                emp_id = db.session.query(Employee_information.emp_id).filter(
                    Employee_information.emp_email == emp.emp_email
                ).first()
                if emp_id:
                    emp_id = emp_id[0]
                project_status = db.session.query(Target_columns.status).filter(
                    Target_columns.emp_id == emp_id,  # Query for a single employee
                    Target_columns.target_month == current_month_formatted,
                    Target_columns.target_year == current_year_formatted
                ).first()
                project_status = project_status[0] if project_status else None
                project_status=status_mapping.get(project_status, "Targets not set")
                project_name = db.session.query(Employee_information.emp_project).filter(
                    Employee_information.emp_id == emp_id
                ).first() if emp_id else None
                
                
                project_name = project_name[0] if project_name else None
                approval_manager_row = db.session.query(ProjectTargets.ApprovalManager).filter(
                    ProjectTargets.Project == project_name
                ).first()
                approval_manager_row = approval_manager_row[0] if approval_manager_row else None
                is_approval_manager = user_name.lower() == approval_manager_row.lower() if approval_manager_row else None
                 
                approval_record = db.session.query(DmaxApprovals).filter_by(
                    employee_email=emp.emp_id,
                    approved_month=selected_month,
                    approved_year=selected_year
                ).first()
                if approval_record:
                    approval_status=approval_record.status
                else:
                    approval_status=None    
                print(emp_id,project_name,approval_manager_row,approval_status)
                
                # If matched employees exist, calculate averages
                if matched_employees.count() > 0:
                    averages = get_averages_for_filtered_employees(matched_employees)
                    
                    # If averages are found, append to the filtered employees list
                    if averages:
                         # Debugging line to check the results

                        # Append the employee data with averages to the result list
                        filtered_employees.append(
                            {
                                "id": emp.id,  # Employee ID
                                "employee_name": emp.emp_name,  # Employee's name
                                "employee_id": emp.emp_id,
                                "target": averages["avg_target"],  # ✅ Correct way to access dictionary values
                                "actual": averages["avg_actual"],
                                "production": averages["avg_production"],
                                "quality": averages["avg_quality"],
                                "attendance": averages["avg_attendance"],
                                "skill": averages["avg_skill"],
                                "new_initiatives": averages["avg_new_initiatives"],
                                "Dmax_score": averages["avg_Dmax_score"],
                                "project":emp.emp_project,
                                "project_status": project_status,
                                "approval_status": approval_status,
                                "is_approval_manager": is_approval_manager
                            }
                        )

            # filtered_employees = []
            # matched_employees = get_filtered_employees(
            #     Dform.query,
            #     search_query,
            #     selected_month,
            #     selected_date
            # )
            # if matched_employees:
            #     for matched in matched_employees:
            #         filtered_employees.append(
            #                 {
            #                     column: getattr(matched, column, None)  # Use getattr to get the attribute dynamically
            #                     for column in ALLOWED_COLUMNS         # Filter by allowed columns
            #                 }
            #             )
            return render_template("view_dscore.html", employees=filtered_employees, role=role, search_query=search_query, selected_month=selected_month, selected_date=selected_date,selected_year=int(selected_year),years=years)

@app.route('/delete_employee/<int:id>', methods=['POST'])
def delete_employee(id):
    employee = Employee_information.query.get(id)
    if employee:
        db.session.delete(employee)
        db.session.commit()
        flash('Employee deleted successfully!', 'success')
    else:
        flash('Employee not found.', 'danger')
    return jsonify({'success': True})

@app.route('/operational_excellence/<string:emp_id>', methods=['GET', 'POST'])
def operational_excellence(emp_id):
    employee_info = Employee_information.query.filter_by(emp_id=emp_id).first()
    op_excellence=OperationalExcellence.query.filter_by(emp_id=emp_id).first()
    if not op_excellence:
        op_excellence = OperationalExcellence(
            emp_id=emp_id,
            month=0,
            
            dtouch_score=0,
            new_init_score=0
        )
        db.session.add(op_excellence)
        db.session.commit()
    
    if employee_info:
        designation = employee_info.emp_designation
        if request.method == 'POST':
            month=request.form.get('month')
            
            # attendance=request.form.get('attendance')
            dtouch = request.form.get('dtouch')
            new_init=request.form.get('newInitiatives')
            # attendance =int(attendance)
            dtouch=int(dtouch)   
            new_init=int(new_init) 
            if designation == "Intern":
                # attendance = int((attendance * 10 / 100) * 100)
                dtouch = int(((dtouch * 10 / 100 / 100) * 100) * 100)
                new_init=0
            if designation=="Jr.QA Engineer":
                # attendance = int((attendance * 5 / 100) * 100)
                dtouch = int(((dtouch * 10 / 100 / 100) * 100) * 100)
                new_init=0
            if designation=="QA Engineer":
                # attendance = int((attendance * 5 / 100) * 100)
                dtouch = int(((dtouch * 5 / 100 / 100) * 100) * 100)
                new_init = int(((new_init * 15 / 100 / 100) * 100) * 100)
            if designation=="Sr.QA Engineer":
                # attendance = int((attendance * 5 / 100) * 100)
                dtouch = (((dtouch * 5 / 100 / 100) * 100) )
                new_init = ((new_init * 20 / 100 / 100) * 100)
                
            if designation=="QA Lead":
                # attendance = int((attendance * 5 / 100) * 100)
                dtouch = (((dtouch * 5 / 100 / 100) * 100) )    
                new_init = ((new_init * 20 / 100 / 100) * 100)
            start_date,end_date = get_date_range_for_month(month)    
            # op_excellence.attendance_score = attendance
            op_excellence.dtouch_score = dtouch
            op_excellence.new_init_score = new_init    
            op_excellence.month = month
            op_excellence.start_date = start_date
            op_excellence.end_date = end_date
            db.session.commit()    

            
    return render_template('operational_excellence.html',emp_id=emp_id,months_dict=monthsDict)

@app.route("/full_table_view/<string:id>", methods=['GET'])
def full_table_view(id):
    user_name = get_logged_in_user_details()
    if user_name:
        
        role = user_name['role']
        
    project=request.args.get('project')
    
    base_query = Dform.query.filter_by(employee_id=id)
    current_year=datetime.now().year
    selected_date = request.args.get('date')
    current_month = str(datetime.now().strftime("%m"))
    
    selected_month = request.args.get('month', current_month).zfill(2)
    selected_month_formatted=int(request.args.get('month', current_month).zfill(1))
    selected_month_name = monthsDict.get(selected_month)
    
    
     
    selected_year = request.args.get('year', current_year)
    
    selected_year_formatted = selected_year
    approval_exists = db.session.query(DmaxApprovals.id).filter_by(
        employee_email=id,  # Replace `id` with the actual employee identifier
        approved_month=selected_month,
        approved_year=selected_year,
        status="Approved"# If you need to check specifically for an approved status
    ).first()
    approved = "Yes" if approval_exists else "No"
    filtered_query=get_first_filtered_employees(base_query, None, selected_month, selected_date, selected_year)
    employee = filtered_query.all() 
    if request.args.get('download_excel') == '1':
        return generate_excel_from_template(employee)
    years = [current_year - i for i in range(11)]
    selected_year=request.args.get('year',current_year)
    CATEGORY_TO_COLUMNS = {
        "Testcase Creation": [
            "test_case_creation_target", "test_case_creation_actual"
        ],
        "Testcase Updation": [
            "test_case_updation_target", "test_case_updation_actual"
        ],
        "Testcase Execution": [
            "test_case_execution_target", "test_case_execution_actual"
        ],
        "Defects (5/day)": [
            "defects_found_target", "defects_found_actual"
        ],
        "Issue Verification": [
            "defects_verification_target", "defects_verification_actual"
        ],
        "Testscripts Creation": [
            "test_scripts_creation_target", "test_scripts_creation_actual"
        ],
        "Testscripts Updation": [
            "test_scripts_updation_target", "test_scripts_updation_actual"
        ],
        "Testscripts Execution": [
            "test_scripts_execution_target", "test_scripts_execution_actual"
        ],
        "Site Scrub": [
            "site_Scrub_target", "site_Scrub_actual"
        ],
        "Project Documentation": [
            "project_doc_target", "project_doc_actual"
        ],
        "Internal review":[
            "internal_Review_target", "internal_Review_actual"
        ],
        "Regression cycle":[
            "regression_cycle_target", "regression_cycle_actual"
        ],
        "Requirement analyzing/writing testcondition":[
            "regression_cycle_target", "regression_cycle_actual"
        ],
        "End-End test cases executed":[
            "end_cases_exec_target", "end_cases_exec_actual"
        ],
        "Task Achivement/Coverage score":[
            "task_coverage_score_target","task_coverage_score_actual"
        ],
        "Assessment Test score":[
           "assessment_score_target","assessment_score_actual" 
        ],
        "Assessment Retest score":[
          "assessment_re_score_target","assessment_re_score_actual"  
        ],
        "Certification Test score":[
           "cert_score_target","cert_score_actual" 
        ],
        "Certification Retest score":[
           "cert_re_score_target","cert_re_score_actual" 
        ],   
        "New Features Implemented":[
           "new_features_imp_target","new_features_imp_actual" 
        ],   
        "New Features Implemented":[
           "new_features_imp_target","new_features_imp_actual" 
        ], 
        "Defects Fixed":[
            "defects_fixed_target","defects_fixed_actual"
        ],
        "Enhancements Target":[
            "enhancements_target","enhancements_actual"
        ],
        "Figma Designs Created": [
            "fig_desgns_target", "fig_desgns_actual"
        ],
        "Project Documentation Updation": [
            "doc_update_target", "doc_update_actual"
        ],
        "Research": [
            "research_target", "research_actual"
        ]
        # Add other mappings here as needed
    }
    PROJECT_SELECTION = {
            "Akyrian": [
                "Testcase Creation",  # Only these categories should be included
                "Testcase Updation",
                "Testcase Execution",
                "Defects (5/day)",
                "Issue Verification",
                "Testscripts Creation",
                "Testscripts Updation",
                "Testscripts Execution",
                "Site Scrub",
                "Project Documentation",
                "Internal review",
                "Regression cycle",
                "Requirement analyzing/writing testcondition",
                "End-End test cases executed"
            ],
            "Akyrian": [
                "Testcase Creation",  # Only these categories should be included
                "Testcase Updation",
                "Testcase Execution",
                "Defects (5/day)",
                "Issue Verification",
                "Testscripts Creation",
                "Testscripts Updation",
                "Testscripts Execution",
                "Site Scrub",
                "Project Documentation",
                "Internal review",
                "Regression cycle",
                "Requirement analyzing/writing testcondition",
                "End-End test cases executed"
            ],
            "Auxo": [
                "Testcase Creation",  # Only these categories should be included
                "Testcase Updation",
                "Testcase Execution",
                "Defects (5/day)",
                "Issue Verification",
                "Testscripts Creation",
                "Testscripts Updation",
                "Testscripts Execution",
                "Site Scrub",
                "Project Documentation",
                "Internal review",
                "Regression cycle",
                "Requirement analyzing/writing testcondition",
                "End-End test cases executed"
            ],
            "Avanti": [
                "Testcase Creation",  # Only these categories should be included
                "Testcase Updation",
                "Testcase Execution",
                "Defects (5/day)",
                "Issue Verification",
                "Site Scrub",
                "Project Documentation",
                "Internal review",
                "Regression cycle",
                "Requirement analyzing/writing testcondition",
                "End-End test cases executed"
            ],
            "Bench" :[
                "Task Achivement/Coverage score",  # Only these categories should be included
                "Assessment Test score",
                "Assessment Retest score",
                "Certification Test score",
                "Certification Retest score"
            ],
            "Fora Travels": [
                "Testcase Creation",  # Only these categories should be included
                "Testcase Updation",
                "Testcase Execution",
                "Defects (5/day)",
                "Issue Verification",
                "Testscripts Creation",
                "Testscripts Updation",
                "Testscripts Execution",
                "Site Scrub",
                "Project Documentation",
                "Internal review",
                "Regression cycle",
                "Requirement analyzing/writing testcondition",
                "End-End test cases executed"
            ],
            "Indihood": [
                "Testcase Creation",  # Only these categories should be included
                "Testcase Updation",
                "Testcase Execution",
                "Defects (5/day)",
                "Issue Verification",
                "Testscripts Creation",
                "Testscripts Updation",
                "Testscripts Execution",
                "Site Scrub",
                "Project Documentation",
                "Internal review",
                "Regression cycle",
                "Requirement analyzing/writing testcondition",
                "End-End test cases executed"
            ],
            "IPS":[
                "Testcase Creation",  # Only these categories should be included
                "Testcase Updation",
                "Testcase Execution",
                "Defects (5/day)",
                "Issue Verification",
                "Testscripts Creation",
                "Testscripts Updation",
                "Testscripts Execution",
                "Site Scrub",
                "Project Documentation",
                "Internal review",
                "Regression cycle",
                "Requirement analyzing/writing testcondition",
                "End-End test cases executed"
            ],
            "IQHive":[
                "Testcase Creation",  # Only these categories should be included
                "Testcase Updation",
                "Testcase Execution",
                "Defects (5/day)",
                "Issue Verification",
                "Testscripts Creation",
                "Testscripts Updation",
                "Testscripts Execution",
                "Site Scrub",
                "Project Documentation",
                "Internal review",
                "Regression cycle",
                "Requirement analyzing/writing testcondition",
                "End-End test cases executed"
            ],
            "IQHive":[
                "Testcase Creation",  # Only these categories should be included
                "Testcase Updation",
                "Testcase Execution",
                "Defects (5/day)",
                "Issue Verification",
                "Testscripts Creation",
                "Testscripts Updation",
                "Testscripts Execution",
                "Site Scrub",
                "Project Documentation",
                "Internal review",
                "Regression cycle",
                "Requirement analyzing/writing testcondition",
                "End-End test cases executed"
            ],
            "LevelBlue":[
                "Testcase Creation",  # Only these categories should be included
                "Testcase Updation",
                "Testcase Execution",
                "Defects (5/day)",
                "Issue Verification",
                "Testscripts Creation",
                "Testscripts Updation",
                "Testscripts Execution",
                "Site Scrub",
                "Project Documentation",
                "Internal review",
                "Regression cycle",
                "Requirement analyzing/writing testcondition",
                "End-End test cases executed"
            ],
            "Web Development":[
                "New Features Implemented",
                "Defects Fixed",
                "Enhancements Target",
                "Figma Designs Created",
                "Project Documentation Updation",
                "Research"
            ],
            "Opus Clip":[
                "Testcase Creation",  # Only these categories should be included
                "Testcase Updation",
                "Testcase Execution",
                "Defects (5/day)",
                "Issue Verification",
                "Site Scrub",
                "Project Documentation",
                "Internal review",
                "Regression cycle",
                "Requirement analyzing/writing testcondition",
                "End-End test cases executed"
            ],
            "Bench" :[
                "Task Achivement/Coverage score",  # Only these categories should be included
                "Assessment Test score",
                "Assessment Retest score",
                "Certification Test score",
                "Certification Retest score"
            ]

            # Add more projects with specific selections here
        }
    TABLE_HEADERS = {
        "Production": {
            "Testcase Creation": ["Target", "Actual"],
            "Testcase Updation": ["Target", "Actual"],
            "Testcase Execution": ["Target", "Actual"],
            "Defects (5/day)": ["Target", "Actual"],
            "Issue Verification": ["Target", "Actual"],
            "Testscripts Creation": ["Target", "Actual"],
            "Testscripts Execution": ["Target", "Actual"],
            "Testscripts Updation": ["Target", "Actual"],
            "Project Documentation": ["Target", "Actual"],
            "Internal review": ["Target", "Actual"],
            "Regression cycle": ["Target", "Actual"],
            "Requirement analyzing/writing testcondition": ["Target", "Actual"],
            "End-End test cases executed": ["Target", "Actual"],
            "Site Scrub": ["Target", "Actual"],
            "Task Achivement/Coverage score": ["Target", "Actual"],
            "Assessment Test score": ["Target", "Actual"],
            "Assessment Retest score": ["Target", "Actual"],
            "Certification Test score": ["Target", "Actual"],
            "Certification Retest score": ["Target", "Actual"],
            "New Features Implemented": ["Target", "Actual"],
            "Defects Fixed": ["Target", "Actual"],
            "Enhancements Target": ["Target", "Actual"],
            "Figma Designs Created": ["Target", "Actual"],
            "Project Documentation Updation": ["Target", "Actual"],
            "Research": ["Target", "Actual"]
        },
        "Quality": {
            "No. of invalid defects": None,
            "Spelling/Typo errors": None,
            "Client escalations": None,
            "Testcase missing": None,
        },
        "Attendance": None,
        "Skill": None,
        "New initiatives":None,
        "Production%":{
            "Target": None,
            "Actual": None,
            "production %": None,
        },
        
        "Quality%": None,
        
        "attendance":None,
        "Skill (%)":None,
        "new initiatives(%)":None,
        "Dmax score":None
    }
    selected_categories = PROJECT_SELECTION.get(project, [])
    if project in PROJECT_SELECTION:
    # Get the production headers
        production_headers = TABLE_HEADERS.get("Production", {})
        for category in list(production_headers.keys()):
            if category not in selected_categories:
                # Remove the unwanted category
                del production_headers[category]

    
    # Only keep the categories that are selected for this project
    
    allowed_columns=[]
    # for category in selected_categories:
    #     if category in CATEGORY_TO_COLUMNS:
    #         allowed_columns.extend(CATEGORY_TO_COLUMNS[category])
    # print(allowed_columns)        
    
    ALLOWED_COLUMNS = [
        "employee_name", "today_date", "test_case_creation_target",
        "test_case_creation_actual", "test_case_updation_target", "test_case_updation_actual",
        "test_case_execution_target", "test_case_execution_actual", "defects_found_target",
        "defects_found_actual","defects_verification_target", "defects_verification_actual", "test_scripts_creation_target", "test_scripts_creation_actual",
        "test_scripts_execution_target","test_scripts_execution_actual","test_scripts_updation_target", "test_scripts_updation_actual",
         "project_doc_target","project_doc_actual","internal_Review_target", "internal_Review_actual",
        
          "regression_cycle_target","regression_cycle_actual","req_anal_target", "req_anal_actual",
          "end_cases_exec_target","end_cases_exec_actual","site_Scrub_target", "site_Scrub_actual",
         "task_coverage_score_target", "task_coverage_score_actual","assessment_score_target", "assessment_score_actual",
          "assessment_re_score_target","assessment_re_score_actual","cert_score_target", "cert_score_actual","cert_re_score_target","cert_re_score_actual",
           "new_features_imp_target", "new_features_imp_actual","defects_fixed_target", "defects_fixed_actual",
           "enhancements_target", "enhancements_actual", "fig_desgns_target", "fig_desgns_actual","doc_update_target", "doc_update_actual",
           "research_target", "research_actual",
          "inv_defs",  "spel_errors",  "client_esc", "tst_cases_missing","att","skill","new_initiatives",
          "target","actual","production","quality","attendance","skill","new_initiatives","Dmax_score"
            
        # #     
        # ,"Dmax_score",new_init
        # "quality", "attendance", "skill",  
    ]
    core_columns = [
        "employee_name", "today_date","inv_defs",  "spel_errors",  "client_esc", "tst_cases_missing","att","skill","new_initiatives",
          "target","actual","production","quality","attendance","skill","new_initiatives","Dmax_score"
    ]
    filtered_columns = []
    for column in ALLOWED_COLUMNS:
        if column in core_columns:
            filtered_columns.append(column)
            continue  # Skip to the next column

        for category in selected_categories:
            if category in CATEGORY_TO_COLUMNS:
                if column in CATEGORY_TO_COLUMNS[category]:  
                    filtered_columns.append(column)
                  
    if employee:
        return render_template("full_table_view.html", employee=employee, ALLOWED_COLUMNS=filtered_columns,TABLE_HEADERS=TABLE_HEADERS,years=years,selected_year=int(selected_year),selected_date=selected_date,monthsDict=monthsDict,current_month=current_month,selected_month=selected_month,project=project,role=role,approved=approved)
    return "No data found"
    
@app.route('/approve_users')
def approve_users():
    pending_users = Employee.query.filter_by(is_approved=False).all()
    approved_users=Employee.query.filter_by(is_approved=True).all()
    return render_template('approve_users.html', pending_users=pending_users,approved_users=approved_users)

@app.route('/approve_selected', methods=['POST'])
def approve_selected():
    print("Approving selected users")
    data = request.get_json()
    emp_ids = data.get('emp_ids', [])
    if not emp_ids:
        return jsonify({"success": False, "message": "No users selected"}), 400

    try:
        # Update is_approved to True for selected users
        Employee.query.filter(Employee.emp_id.in_(emp_ids)).update({"is_approved": True}, synchronize_session=False)
        db.session.commit()
        return jsonify({"success": True})
    
    except Exception as e:
        db.session.rollback()  # Rollback if something goes wrong
        return jsonify({"success": False, "message": str(e)}), 500
        
@app.route('/form_bulk_upload', methods=['GET','POST'])
def form_bulk_upload():
    
    # field_to_column = {
    #         "employee_name": 'A',
    #         "employee_id": 'B',
    #         "employee_email": 'C',
    #         "today_date": 'D',
    #         "project": 'E',
    #         "designation": 'F',
    #         "test_case_creation_target": 'G',
    #         "test_case_creation_actual": 'H',
    #         "test_case_updation_target": 'I',
    #         "test_case_updation_actual": 'J',
    #         "test_case_execution_target": 'K',
    #         "test_case_execution_actual": 'L',
    #         "defects_found_target":'M',
    #         "defects_found_actual":'N',
    #         "defects_verification_target":'O',
    #         "defects_verification_actual":'P',
    #         "test_scripts_creation_target":'Q',
    #         "test_scripts_creation_actual":'R',
    #         "test_scripts_updation_target":'S',
    #         "test_scripts_updation_actual":'T',
    #         "test_scripts_execution_target":'U',
    #         "test_scripts_execution_actual":'V',
    #         "site_Scrub_target":'AG',
    #         "site_Scrub_actual":'AH',
    #         "project_doc_target":'W',
    #         "project_doc_actual":'X',
    #         "internal_Review_target":'Y',
    #         "internal_Review_actual":'Z',
    #         "regression_cycle_target":'AA',
    #         "regression_cycle_actual":'AB',
    #         "req_anal_target":'AC',
    #         "req_anal_actual":'AD',
    #         "end_cases_exec_target":'AE',
    #         "end_cases_exec_actual":'AF',
    #         "task_coverage_score_target":'AI',
    #         "task_coverage_score_actual":'AJ',
    #         "assessment_score_target":'AK',
    #         "assessment_score_actual":'AL',
    #         "assessment_re_score_target":'AM',
    #         "assessment_re_score_actual":'AN',
    #         "cert_score_target":"AO",
    #         "cert_score_actual":'AP',
    #         "cert_re_score_target":'AQ',
    #         "cert_re_score_actual":'AR',
    #         "new_features_imp_target":'AS',
    #         "new_features_imp_actual":'AT',
    #         "defects_fixed_target":'AU',
    #         "defects_fixed_actual":'AV',
    #         "enhancements_target":'AW',
    #         "enhancements_actual":'AX',
    #         "fig_desgns_target":'AY',
    #         "fig_desgns_actual":'AZ',
    #         "doc_update_target":'BA',
    #         "doc_update_actual":'BB',
    #         "research_target":'BC',
    #         "research_actual":'BD',
    #         "inv_defs":'BE',
    #         "spel_errors":'BF',
    #         "client_esc":'BG',
    #         "tst_cases_missing":'BH',
    #         "att":'BI',
    #         "dtouch":'BJ',
    #         "new_init":'BK',  
    #         "target":'BL' 
    #     }
    # actual_to_target_mapping = {}

    # for key in field_to_column.keys():
    #     if key.endswith('_actual'):
    #         target_key = key.replace('_actual', '_target')  # Replace '_actual' with '_target'
    #         if target_key in field_to_column:  # Check if target_key exists
    #             actual_to_target_mapping[key] = target_key
    # if request.method == 'POST':
    #     file = request.files['file']
    #     wb = load_workbook(file, data_only=True)
    #     ws=wb.active
    #     data_list = []

    #     for row in ws.iter_rows(min_row=3, values_only=True):  # Skip header row
    #         row_data = {}
    #         for field, column in field_to_column.items():
    #             col_index = openpyxl.utils.column_index_from_string(column) - 1
    #             row_data[field] = row[col_index] if col_index < len(row) else None
                
    #             designation = row_data.get("designation", "")
    #             attendance_input = row_data.get("att", 0)
    #             if designation == "Intern":
    #                 attendance = int((attendance_input * 10 / 100) * 100)
    #             elif designation == "Jr.QA Engineer":
    #                 attendance = int((attendance_input * 10 / 100) * 100)
    #             elif designation == "QA Engineer":   
    #                 attendance = int((attendance_input * 5 / 100) * 100) 
    #             elif designation=="Sr.QA Engineer":
    #                 attendance = int((attendance_input * 5 / 100) * 100)
    #             elif designation=="QA Lead":
    #                 attendance = int((attendance_input * 5 / 100) * 100)  
    #             else:
    #                 attendance = 0    
    #             row_data["attendance"] = attendance
    #             row_data["target"] = 0
    #             row_data["actual"] = 0
    #         if all(value not in (None, "") for value in row_data.values()):  
    #             data_list.append(row_data)
    #     for row_data in data_list:
    #         # Fetch operational excellence details for each employee
    #         operational_excellence = OperationalExcellence.query.filter_by(emp_id=row_data.get("employee_id")).first()
            
    #         if operational_excellence and operational_excellence.start_date and operational_excellence.end_date:
    #             # Convert start_date and end_date to datetime objects
    #             start_date = datetime.strptime(operational_excellence.start_date, "%Y-%m-%d")
    #             end_date = datetime.strptime(operational_excellence.end_date, "%Y-%m-%d")

    #             # Convert 'today_date' in row_data to a datetime object (ensure the field exists)
    #             today_date_str = row_data.get('today_date', '')
    #             print("today_date_str",today_date_str)
    #             print("start_date",start_date)
    #             print("end_date",end_date)
    #             if today_date_str:
    #                 if isinstance(today_date_str, str):  # Check if it's a string
    #                     today_date = datetime.strptime(today_date_str, "%Y-%m-%d")
    #                 else:
    #                     today_date = today_date_str

    #                 # Check if today_date is within the operational excellence date range
    #                 if start_date <= today_date <= end_date:
    #                     # Update the skill and new_initiatives in the row_data
    #                     row_data['skill'] = operational_excellence.dtouch_score
    #                     row_data['new_initiatives'] = operational_excellence.new_init_score
    #                     print("startdate", start_date, "enddate", end_date, "today_date", today_date)
    #                 else:
    #                     # If today's date is not in the range, set default values
    #                     row_data['skill'] = 0
    #                     row_data['new_initiatives'] = 0
    #             else:
    #                 # Handle case where 'today_date' is missing or invalid
    #                 row_data['skill'] = 0
    #                 row_data['new_initiatives'] = 0
    #         else:
    #             # If no operational excellence record is found, set default values
    #             row_data['skill'] = 0
    #             row_data['new_initiatives'] = 0
            
    #         if row_data.get('client_esc', 0) == 1:  # Check if BG (Client Escalations) is 1
    #             row_data['quality'] = 0  # Set quality to 0 if BG is 1
    #         else:
    #             sum_invalid_defects_to_test_cases = (
    #                 row_data.get('inv_defs', 0) +  # BE: Invalid Defects
    #                 row_data.get('spel_errors', 0) +  # BF: Spelling Errors
    #                 row_data.get('client_esc', 0) +  # BG: Client Escalations
    #                 row_data.get('tst_cases_missing', 0)  # BH: Test Cases Missing
    #             )       
    #             row_data['quality'] = ((100 - sum_invalid_defects_to_test_cases) * 0.4 / 100) * 100
            
    #         for actual_field, target_field in actual_to_target_mapping.items():
    #             for row in data_list: 
    #                 if actual_field in row and row[actual_field] is not None and row[actual_field] > 0:
    #                     row['target'] += int(row.get(target_field, 0))  # Use .get() to avoid KeyError
    #                     row['actual'] += int(row.get(actual_field, 0))  

    #                 if row['target'] != 0 and row['actual'] != 0:  
    #                     row['production'] = ((row['actual'] / row['target']) * 40 / 100) *100
    #                 else:
    #                     row['production'] = 0   
    #         row_data['Dmax_score'] = sum([
    #             row_data.get('production', 0),
    #             row_data.get('quality', 0),
    #             row_data.get('attendance', 0),
    #             row_data.get('new_initiatives', 0),
    #             row_data.get('skill', 0)
    #         ])      
    #         if row_data["attendance"]==0:      
    #             row_data['Dmax_score'] =0
    #             row_data['production'] =0
    #             row_data['quality'] =0
            
    #         new_entry = Dform(
    #         employee_name=row_data['employee_name'],
    #         employee_id=row_data['employee_id'],
    #         employee_email=row_data['employee_email'],
    #         today_date=row_data['today_date'],
    #         project=row_data['project'],
    #         designation=row_data['designation'],
    #         test_case_creation_target=row_data.get('test_case_creation_target'),
    #         test_case_creation_actual=row_data.get('test_case_creation_actual'),
    #         test_case_updation_target=row_data.get('test_case_updation_target'),
    #         test_case_updation_actual=row_data.get('test_case_updation_actual'),
    #         test_case_execution_target=row_data.get('test_case_execution_target'),
    #         test_case_execution_actual=row_data.get('test_case_execution_actual'),
    #         defects_found_target=row_data.get('defects_found_target'),
    #         defects_found_actual=row_data.get('defects_found_actual'),
    #         test_scripts_creation_target=row_data.get('test_scripts_creation_target'),
    #         test_scripts_creation_actual=row_data.get('test_scripts_creation_actual'),
    #         test_scripts_updation_target=row_data.get('test_scripts_updation_target'),
    #         test_scripts_updation_actual=row_data.get('test_scripts_updation_actual'),
    #         test_scripts_execution_target=row_data.get('test_scripts_execution_target'),
    #         test_scripts_execution_actual=row_data.get('test_scripts_execution_actual'),
    #         site_Scrub_target=row_data.get('site_Scrub_target'),
    #         site_Scrub_actual=row_data.get('site_Scrub_actual'),
    #         project_doc_target=row_data.get('project_doc_target'),
    #         project_doc_actual=row_data.get('project_doc_actual'),
    #         internal_Review_target=row_data.get('internal_Review_target'),
    #         internal_Review_actual=row_data.get('internal_Review_actual'),
    #         regression_cycle_target=row_data.get('regression_cycle_target'),
    #         regression_cycle_actual=row_data.get('regression_cycle_actual'),
    #         req_anal_target=row_data.get('req_anal_target'),
    #         req_anal_actual=row_data.get('req_anal_actual'),
    #         end_cases_exec_target=row_data.get('end_cases_exec_target'),
    #         end_cases_exec_actual=row_data.get('end_cases_exec_actual'),
    #         task_coverage_score_target=row_data.get('task_coverage_score_target'),
    #         task_coverage_score_actual=row_data.get('task_coverage_score_actual'),
    #         assessment_score_target=row_data.get('assessment_score_target'),
    #         assessment_score_actual=row_data.get('assessment_score_actual'),
    #         assessment_re_score_target=row_data.get('assessment_re_score_target'),
    #         assessment_re_score_actual=row_data.get('assessment_re_score_actual'),
    #         cert_score_target=row_data.get('cert_score_target'),
    #         cert_score_actual=row_data.get('cert_score_actual'),
    #         cert_re_score_target=row_data.get('cert_re_score_target'),
    #         cert_re_score_actual=row_data.get('cert_re_score_actual'),
    #         new_features_imp_target=row_data.get('new_features_imp_target'),
    #         new_features_imp_actual=row_data.get('new_features_imp_actual'),
    #         defects_fixed_target=row_data.get('defects_fixed_target'),
    #         defects_fixed_actual=row_data.get('defects_fixed_actual'),
    #         defects_verification_target=row_data.get('defects_verification_target'),
    #         defects_verification_actual=row_data.get('defects_verification_actual'),
    #         enhancements_target=row_data.get('enhancements_target'),
    #         enhancements_actual=row_data.get('enhancements_actual'),
    #         fig_desgns_target=row_data.get('fig_desgns_target'),
    #         fig_desgns_actual=row_data.get('fig_desgns_actual'),
    #         doc_update_target=row_data.get('doc_update_target'),
    #         doc_update_actual=row_data.get('doc_update_actual'),
    #         research_target=row_data.get('research_target'),
    #         research_actual=row_data.get('research_actual'),
    #         inv_defs=row_data.get('inv_defs'),
    #         spel_errors=row_data.get('spel_errors'),
    #         client_esc=row_data.get('client_esc'),
    #         tst_cases_missing=row_data.get('tst_cases_missing'),
    #         att=row_data.get('att'),
    #         target=row_data['target'],
    #         actual=row_data['actual'],
    #         production=row_data['production'],
    #         quality=row_data['quality'],
    #         attendance=row_data['attendance'],
    #         skill=row_data['skill'],
    #         new_initiatives=row_data['new_initiatives'],
    #         Dmax_score=row_data['Dmax_score'],
    #     )

    #     db.session.add(new_entry)
    #     db.session.commit()

    #     return data_list
    field_to_column = {
        "employee_name": 0,
        "employee_id": 1,
        "employee_email": 2,
        "today_date": 3,
        "project": 4,
        "designation": 5,
        "test_case_creation_target": 6,
        "test_case_creation_actual": 7,
        "test_case_updation_target": 8,
        "test_case_updation_actual": 9,
        "test_case_execution_target": 10,
        "test_case_execution_actual": 11,
        "defects_found_target": 12,
        "defects_found_actual": 13,
        "defects_verification_target": 14,
        "defects_verification_actual": 15,
        "test_scripts_creation_target": 16,
        "test_scripts_creation_actual": 17,
        "test_scripts_updation_target": 18,
        "test_scripts_updation_actual": 19,
        "test_scripts_execution_target": 20,
        "test_scripts_execution_actual": 21,
        "site_Scrub_target": 22,
        "site_Scrub_actual": 23,
        "project_doc_target": 24,
        "project_doc_actual": 25,
        "internal_Review_target": 26,
        "internal_Review_actual": 27,
        "regression_cycle_target": 28,
        "regression_cycle_actual": 29,
        "req_anal_target": 30,
        "req_anal_actual": 31,
        "end_cases_exec_target": 32,
        "end_cases_exec_actual": 33,
        "task_coverage_score_target": 34,
        "task_coverage_score_actual": 35,
        "assessment_score_target": 36,
        "assessment_score_actual": 37,
        "assessment_re_score_target": 38,
        "assessment_re_score_actual": 39,
        "cert_score_target": 40,
        "cert_score_actual": 41,
        "cert_re_score_target": 42,
        "cert_re_score_actual": 43,
        "new_features_imp_target": 44,
        "new_features_imp_actual": 45,
        "defects_fixed_target": 46,
        "defects_fixed_actual": 47,
        "enhancements_target": 48,
        "enhancements_actual": 49,
        "fig_desgns_target": 50,
        "fig_desgns_actual": 51,
        "doc_update_target": 52,
        "doc_update_actual": 53,
        "research_target": 54,
        "research_actual": 55,
        "inv_defs": 56,
        "spel_errors": 57,
        "client_esc": 58,
        "tst_cases_missing": 59,
        "att": 60,
        "dtouch": 61,
        "new_init": 62,
        "target": 63,
        "actual":64,
        "production":65,
        "quality":66,
        "attendance":67,
        "skill":68,
        "new_initiatives":69,
        "Dmax_score":70
    }
    if request.method == 'POST':
        file = request.files["file"]
        wb = load_workbook(file, data_only=True)
        ws=wb.active
        all_row_data = []
        if file.filename == "":
            return "No file selected", 400
        for row in ws.iter_rows(min_row=4, values_only=True):  # Use values_only=False to access the cells directly
            if not any(row):  # Skip empty rows
                continue
            row_data = {}
            
            # Map the columns to the appropriate fields and get the .value for each formula
            for field, col_index in field_to_column.items():
                cell_value = row[col_index]  # Directly get the value
                if field == "today_date" and isinstance(cell_value, datetime):
                    row_data[field] = cell_value.date().strftime("%Y-%m-%d")    
                if field in ["production", "quality", "attendance", "skill", "new_initiatives", "Dmax_score"]:
                    if isinstance(cell_value, (int, float)):  # Ensure it's numeric before multiplying
                        row_data[field] = round(cell_value * 100, 2)  # Convert to percentage
                    else:
                        row_data[field] = cell_value  # Keep as is if not numeric
                else:
                    row_data[field] = cell_value
            
            # Check if any field in the row is None, and skip the row if it contains null values
            new_entry = Dform(**row_data)
            db.session.add(new_entry)
        db.session.commit()
              
    return render_template("form_bulk_upload.html")

@app.route("/set_targets/<string:emp_id>",methods=["GET","POST"])
def set_targets(emp_id):
    role=request.args.get('role')
    target_month = request.form.get("target_month")
    target_year = request.form.get("target_year")
    
    current_year=datetime.now().year
    current_year=int(current_year)
    current_month_check=datetime.now().strftime("%B")
    current_year_check = str(datetime.now().year)
    
    
    
    # current_month_formatted=datetime.now().strftime('%B')
    # current_year_formatted = str(datetime.now().year)
    month_order = case(
        {month: i for i, (month, _) in enumerate(monthsDict_2.items(), 1)},  # Map month name to numeric value (1 for "January", 2 for "February", etc.)
        value=Target_columns.target_month,
        else_=0
    )
    years=last_ten_years
    employee=Employee_information.query.filter_by(emp_id=emp_id).first()
    first_entry_check = Target_columns.query.filter_by(emp_id=emp_id).first()
    # if first_entry_check is None:  # This means it's their first entry
    #     if target_year != current_year_check or target_month != current_month_check:
    #         flash(f"For the first entry, you can only set a target for {current_month_check} {current_year_check}.", "warning")
    #         return redirect(url_for("set_targets", emp_id=emp_id, role=role))
    #     previous_entry = None  
    if target_month and target_year and first_entry_check:  
        target_year = int(target_year)  
        target_month_num = monthsDict_2.get(target_month)  

        if target_month_num == 1:
            previous_entry=None
        else:

            prev_month_num = target_month_num - 1
            prev_year = target_year

            
        
            prev_month_name = [month for month, num in monthsDict_2.items() if num == prev_month_num][0]
            
            previous_entry = Target_columns.query.filter_by(
                emp_id=emp_id, target_month=prev_month_name, target_year=str(prev_year),status="approved"
            ).first()
    target_user = Target_columns.query.filter_by(emp_id=emp_id, status="waiting for approval")\
    .order_by(Target_columns.target_year.desc(),
              month_order.desc(),  # Sort by the numeric value of the month
    ).first()
    target_month_user=Target_columns.query.filter_by(emp_id=emp_id,target_month=target_month,target_year=target_year).first()
    values={
        "emp_id":employee.emp_id,
        "emp_name":employee.emp_name,
        "project":employee.emp_project
    }
    
    fields = {}
    if target_user:
        fields = {column.name: getattr(target_user, column.name) for column in Target_columns.__table__.columns}
    current_year=current_year
        
    if request.method=="POST":
        role=request.args.get('role')
        month_formatted=request.form.get("target_month")
        
        year_formatted=str(request.form.get("target_year"))
        target_month = request.form.get("target_month")
        target_month_num = monthsDict_2.get(target_month)  
        if target_month_num != 1 and first_entry_check is None:
            if year_formatted != current_year_check or month_formatted != current_month_check:
                flash(f"For the first target entry,Please set a target for {current_month_check} {current_year_check}.", "warning")
                return redirect(url_for("set_targets", emp_id=emp_id, role=role))
            previous_entry = None 
        existing_approved_entry = Target_columns.query.filter_by(
            emp_id=emp_id,
            target_month=month_formatted,
            target_year=year_formatted,
            status="approved"
        ).scalar()
        if existing_approved_entry:
            flash("An approved Dmax score already exists for this month", "warning")
            return redirect(url_for("set_targets",emp_id=emp_id,role="Project Lead"))  # Redirect to the same page

        if target_month_num != 1 and first_entry_check and not previous_entry:
            existing_entry = Target_columns.query.filter_by(
                emp_id=emp_id,
                target_month=month_formatted,
                target_year=year_formatted
            ).first()
            if not existing_entry:
                flash(f"{prev_month_name}'s target has to be approved first!", "warning")
                return redirect(url_for("set_targets", emp_id=emp_id, role=role))
        if target_month_user:
            for field, value in request.form.items():
                
                if hasattr(target_user, field):
                    setattr(target_user, field, value)
            target_user.status = "waiting for approval"        
             
                   
        else:
            target_data = {field: int(value) if value.isdigit() else value for field, value in request.form.items() if hasattr(Target_columns, field)}
            target_data["emp_id"] = emp_id  # Ensure emp_id is included
            target_data["status"] = "waiting for approval"
            new_target_entry = Target_columns(**target_data)
            db.session.add(new_target_entry)            
        db.session.commit()        
    return render_template("set_targets.html", employee=employee,values=values,current_year=current_year,fields=fields,role=role,monthsDict=monthsDict,years=years)

@app.route("/project_targets", methods=["GET", "POST"])
def project_targets():
    column_mapping = {
        "Approval Manager": "ApprovalManager"  # Only for non-matching columns
    }
    
    project_targets = ProjectTargets.query.all()
    if request.method == "POST":
        file = request.files['file']
        wb = load_workbook(file, data_only=True)
        ws=wb.active
        excel_headers = [cell.value for cell in ws[1]]
        db_columns = [column_mapping.get(header, header) for header in excel_headers]
        for row in ws.iter_rows(min_row=2, values_only=True):
            if not any(row):  # Skip empty rows
                continue
            if any(cell is None or cell == "" for cell in row):  # Check if any column is empty
                flash(f"Skipping row with missing data", "danger")
                continue
            data_dict = dict(zip(db_columns, row))
            project_name = data_dict.get("Project")
            if project_name:
                existing_project = ProjectTargets.query.filter_by(Project=project_name).first()
            if existing_project:
                    # Update existing record
                    existing_project.Lead = data_dict.get("Lead", existing_project.Lead)
                    existing_project.ApprovalManager = data_dict.get("ApprovalManager", existing_project.ApprovalManager)    
            else:
                employee = ProjectTargets(**data_dict)
                db.session.add(employee)
        
        db.session.commit()
        
    return render_template("projects_upload.html",project_targets=project_targets)

@app.route("/approve_targets/<string:emp_id>/<string:role>", methods=["GET", "POST"])
def approve_targets(emp_id, role):
    # Handle the logic for approving targets
    if request.method == "POST":
        employee_id = request.form.get('employee_id')  # This returns an iterable of (key, value) pairs
        month_order = case(
            {month: i for i, (month, _) in enumerate(monthsDict_2.items(), 1)},  # Map month name to numeric value (1 for "January", 2 for "February", etc.)
            value=Target_columns.target_month,
            else_=0
        )
        target_entry = Target_columns.query.filter_by(emp_id=employee_id, status="waiting for approval")\
            .order_by(Target_columns.target_year.desc(), 
                      month_order.desc(),  # Order by month (latest month first)
                      ).first() # Order by creation date (latest first) 
        if target_entry:
            # Mark the found target entry as approved
            target_entry.status = 'approved'
            db.session.commit()
            print(f"Approved target entry for {employee_id} (Year: {target_entry.target_year}, Month: {target_entry.target_month})")
        # approved_entry = Target_columns.query.filter_by(emp_id=employee_id, status="approved").first()
        
        # if approved_entry:
        #     db.session.delete(approved_entry)  # Delete the existing approved target
        #     db.session.commit()
        #     print(f"Deleted existing approved entry for {employee_id}")
        # target_entry = Target_columns.query.filter_by(emp_id=employee_id).first()
        # if target_entry:
        #     target_entry.status = 'approved'
        #     db.session.commit()
        return redirect(url_for('team_dmax_table')) 

@app.route('/view_targets')
def view_targets():
    emp_id = request.args.get('emp_id')
    month = request.args.get('month')
    year=request.args.get('year')
    target = Target_columns.query.filter_by(
        emp_id=emp_id, 
        target_month=month, 
        target_year=year
    ).first()
    employee=Employee_information.query.filter_by(emp_id=emp_id).first()
    values={
        "emp_id":employee.emp_id,
        "emp_name":employee.emp_name,
        "project":employee.emp_project
    }
    if target is None:
        return "No target found for this employee for the specified month and year.", 404
    fields = {column.name: getattr(target, column.name) for column in Target_columns.__table__.columns}
    return render_template('view_targets.html', fields=fields,monthsDict=monthsDict,values=values)

@app.route('/profile')
def profile():
    user=get_logged_in_user_details()
    if user:
        print(user)
    if not user or "email" not in user:
        return "User not found", 404  # Handle case where user details are missing
    print(user)
    # Query the Employee table using the logged-in user's email
    employee = Employee.query.filter_by(email=user["email"]).first() 
    
    if not employee:
        return "Employee record not found", 404  # Handle case where employee data is missing

    return render_template("profile.html", employee=employee)
    
@app.route('/delete_employee_data', methods=['POST'])
def delete_employee_data():
    data = request.get_json()
    ids_to_delete = data.get("ids", [])
    
    if not ids_to_delete:
        return jsonify({"success": False, "message": "No IDs received"})    
    Dform.query.filter(Dform.id.in_(ids_to_delete)).delete()
    db.session.commit()
    return jsonify({"success": True})

@app.route('/approve_employee', methods=['POST'])
def approve_employee():
    
    data = request.json 
    email = data.get("email")
    month = data.get("month")
    
    year = data.get("year")
    first_entry_check = DmaxApprovals.query.filter_by(employee_email=email).count() == 0
    if not first_entry_check and int(month) > 1:
        prev_month = int(month) - 1  # Convert month to integer before subtraction
        prev_year = int(year)
        previous_approval=DmaxApprovals.query.filter_by(
            employee_email=email,
            approved_month=prev_month,
            approved_year=prev_year
        ).first()
        if previous_approval and previous_approval.status != "Approved":
            return jsonify({"success": False, "message": f"Approval for {prev_month}-{prev_year} is not completed yet!"}), 400
    existing_approval = DmaxApprovals.query.filter_by(
        employee_email=email,
        approved_month=month,
        approved_year=year
    ).first()

    if existing_approval:
       existing_approval.status = "Approved"

    # Add new approval entry
    new_approval = DmaxApprovals(
        employee_email=email,
        approved_month=month,
        approved_year=year,
        status="waiting for approval"
    )
    db.session.add(new_approval)
    db.session.commit()

    return jsonify({"success": True, "message": "Approval recorded"})


with app.app_context():
        
        db.create_all()
        


if __name__ == "__main__":
    app.run(debug=True)

