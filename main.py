import os
import uuid
import io
import json
from datetime import datetime, timedelta
from fastapi import FastAPI, Depends, HTTPException, status, UploadFile, File, Form, BackgroundTasks
from fastapi.security import OAuth2PasswordBearer, OAuth2PasswordRequestForm
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import Response, FileResponse
from sqlalchemy import create_engine, Column, Integer, String, Float, Date, DateTime, Boolean, ForeignKey, Text, func
from sqlalchemy.ext.declarative import declarative_base
from sqlalchemy.orm import sessionmaker, Session
from passlib.context import CryptContext
from jose import JWTError, jwt
from typing import Optional, List
import cloudinary
import cloudinary.uploader
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from pydantic import BaseModel

# ================= Configuration =================
DATABASE_URL = os.environ.get("DATABASE_URL")
SECRET_KEY = os.environ.get("SECRET_KEY", "your-secret-key-change-in-production")
ALGORITHM = "HS256"
ACCESS_TOKEN_EXPIRE_MINUTES = 60 * 24  # 24 hours

cloudinary.config(
    cloud_name=os.environ.get("CLOUDINARY_CLOUD_NAME"),
    api_key=os.environ.get("CLOUDINARY_API_KEY"),
    api_secret=os.environ.get("CLOUDINARY_API_SECRET"),
    secure=True
)

# ================= Database setup =================
engine = create_engine(DATABASE_URL)
SessionLocal = sessionmaker(autocommit=False, autoflush=False, bind=engine)
Base = declarative_base()

# ================= Models =================
class User(Base):
    __tablename__ = "users"
    id = Column(Integer, primary_key=True, index=True)
    username = Column(String, unique=True, index=True)
    password = Column(String)
    full_name = Column(String)
    role = Column(String)
    enabled = Column(Boolean, default=True)
    last_login = Column(DateTime)
    created_at = Column(DateTime, default=datetime.utcnow)

class Project(Base):
    __tablename__ = "projects"
    id = Column(Integer, primary_key=True, index=True)
    project_code = Column(String, unique=True)
    project_name = Column(String)
    phase = Column(String)
    start_date = Column(Date)
    end_date = Column(Date)
    budget = Column(Float)
    client_name = Column(String)
    status = Column(String, default="Active")

class Activity(Base):
    __tablename__ = "activities"
    id = Column(Integer, primary_key=True, index=True)
    activity_code = Column(String, unique=True)
    activity_name = Column(String)
    discipline = Column(String)
    total_quantity = Column(Float, default=0)
    unit = Column(String)
    baseline_daily_qty = Column(Float, default=0)
    critical = Column(Integer, default=0)
    budget_cost = Column(Float, default=0)
    duration_days = Column(Integer, default=1)
    weight = Column(Float, default=1)
    area = Column(String)
    system = Column(String)
    constraint_type = Column(String)
    constraint_date = Column(Date)
    actual_start = Column(Date)
    actual_finish = Column(Date)
    planned_start = Column(Date)
    planned_finish = Column(Date)
    project_id = Column(Integer, ForeignKey("projects.id"))

class DailyProgress(Base):
    __tablename__ = "daily_progress"
    id = Column(Integer, primary_key=True, index=True)
    report_date = Column(Date)
    activity_id = Column(Integer, ForeignKey("activities.id"))
    engineer_name = Column(String)
    planned_quantity = Column(Float, default=0)
    actual_quantity = Column(Float, default=0)
    cumulative_planned = Column(Float, default=0)
    cumulative_actual = Column(Float, default=0)
    remaining_quantity = Column(Float, default=0)
    planned_percent = Column(Float, default=0)
    actual_percent = Column(Float, default=0)
    variance = Column(Float, default=0)
    manpower = Column(Integer, default=0)
    equipment = Column(String)
    material = Column(String)
    issues = Column(Text)
    next_day_plan = Column(Text)
    status = Column(String)
    submitted_at = Column(DateTime, default=datetime.utcnow)
    user_id = Column(Integer, ForeignKey("users.id"))
    submitted_by = Column(String)

class LookaheadTask(Base):
    __tablename__ = "lookahead_tasks"
    id = Column(Integer, primary_key=True, index=True)
    activity_id = Column(Integer, ForeignKey("activities.id"))
    planned_start = Column(Date)
    planned_finish = Column(Date)
    priority = Column(String)
    constraint_text = Column(Text)
    owner = Column(String)
    status = Column(String)

class ReportPhoto(Base):
    __tablename__ = "report_photos"
    id = Column(Integer, primary_key=True, index=True)
    progress_id = Column(Integer, ForeignKey("daily_progress.id"))
    cloudinary_public_id = Column(String)
    photo_url = Column(String)
    original_filename = Column(String)
    uploaded_at = Column(DateTime, default=datetime.utcnow)
    uploaded_by = Column(String)

class License(Base):
    __tablename__ = "license"
    id = Column(Integer, primary_key=True, default=1)
    expiry_date = Column(Date)
    last_extended_by = Column(String)
    extended_at = Column(DateTime)

class AuditLog(Base):
    __tablename__ = "audit_log"
    id = Column(Integer, primary_key=True, index=True)
    username = Column(String)
    action = Column(String)
    timestamp = Column(DateTime, default=datetime.utcnow)

class SystemSetting(Base):
    __tablename__ = "system_settings"
    key = Column(String, primary_key=True)
    value = Column(String)

# ================= Pydantic schemas =================
class ActivityCreate(BaseModel):
    activity_code: str
    activity_name: str
    discipline: str
    total_quantity: float = 0
    unit: str = ""
    baseline_daily_qty: float = 0
    critical: int = 0
    budget_cost: float = 0
    duration_days: int = 1
    weight: float = 1
    area: str = ""
    system: str = ""
    constraint_type: str = ""
    constraint_date: Optional[str] = None
    actual_start: Optional[str] = None
    actual_finish: Optional[str] = None
    planned_start: Optional[str] = None
    planned_finish: Optional[str] = None
    project_id: int = 1

class DailyProgressUpdate(BaseModel):
    planned_quantity: float
    actual_quantity: float
    manpower: int
    equipment: str = ""
    material: str = ""
    issues: str = ""
    next_day_plan: str = ""

class LookaheadCreate(BaseModel):
    activity_code: str
    planned_start: str
    planned_finish: str
    priority: str
    constraint_text: str = ""
    owner: str = ""

class ProjectCreate(BaseModel):
    project_code: str
    project_name: str
    client_name: str = "SE Saudi Energy"
    start_date: str
    end_date: str
    status: str = "Active"

class ProjectUpdate(BaseModel):
    project_code: str
    project_name: str
    client_name: str
    start_date: str
    end_date: str
    status: str

class UserCreate(BaseModel):
    username: str
    password: str
    full_name: str
    role: str

class UserUpdate(BaseModel):
    full_name: str
    role: str
    password: Optional[str] = None
    enabled: bool

class SettingsUpdate(BaseModel):
    edit_window_hours: int
    monthly_report_cycle_type: str
    monthly_report_start_day: int
    allow_planner_edit_activities: bool
    working_days: List[int]  # 0=Sat,1=Sun,...,6=Fri
    reports_folder: str = ""
    photos_base_folder: str = ""
    backup_folder: str = ""
    backup_frequency: str = "on_submit"
    backup_retention: int = 30

# ================= Helper functions =================
pwd_context = CryptContext(schemes=["pbkdf2_sha256"], deprecated="auto")

def verify_password(plain_password, hashed_password):
    return pwd_context.verify(plain_password, hashed_password)

def get_password_hash(password):
    return pwd_context.hash(password)

def authenticate_user(db: Session, username: str, password: str):
    user = db.query(User).filter(User.username == username, User.enabled == True).first()
    if not user:
        return False
    if not verify_password(password, user.password):
        return False
    return user

def create_access_token(data: dict, expires_delta: timedelta = None):
    to_encode = data.copy()
    if expires_delta:
        expire = datetime.utcnow() + expires_delta
    else:
        expire = datetime.utcnow() + timedelta(minutes=15)
    to_encode.update({"exp": expire})
    encoded_jwt = jwt.encode(to_encode, SECRET_KEY, algorithm=ALGORITHM)
    return encoded_jwt

def get_db():
    db = SessionLocal()
    try:
        yield db
    finally:
        db.close()

oauth2_scheme = OAuth2PasswordBearer(tokenUrl="token")

async def get_current_user(token: str = Depends(oauth2_scheme), db: Session = Depends(get_db)):
    credentials_exception = HTTPException(status_code=401, detail="Could not validate credentials")
    try:
        payload = jwt.decode(token, SECRET_KEY, algorithms=[ALGORITHM])
        username: str = payload.get("sub")
        if username is None:
            raise credentials_exception
    except JWTError:
        raise credentials_exception
    user = db.query(User).filter(User.username == username).first()
    if user is None:
        raise credentials_exception
    return user

def log_audit(db: Session, username: str, action: str):
    log = AuditLog(username=username, action=action, timestamp=datetime.utcnow())
    db.add(log)
    db.commit()

def calculate_progress(installed, total):
    if total <= 0:
        return 0
    return round((installed / total) * 100, 2)

def get_status_from_cumulative_variance(actual, planned, total):
    if actual >= total and total > 0:
        return "Completed"
    var = actual - planned
    if var >= 0:
        return "On Track"
    elif var >= -10:
        return "Delayed"
    else:
        return "Critical"

def calculate_cumulative_and_status(activity_id: int, db: Session):
    entries = db.query(DailyProgress).filter(DailyProgress.activity_id == activity_id).order_by(DailyProgress.report_date).all()
    act = db.query(Activity).filter(Activity.id == activity_id).first()
    total_qty = act.total_quantity if act else 0
    cum_planned = 0
    cum_actual = 0
    for e in entries:
        cum_planned += e.planned_quantity
        cum_actual += e.actual_quantity
        e.cumulative_planned = cum_planned
        e.cumulative_actual = cum_actual
        e.remaining_quantity = max(total_qty - cum_actual, 0)
        e.planned_percent = calculate_progress(cum_planned, total_qty)
        e.actual_percent = calculate_progress(cum_actual, total_qty)
        e.variance = e.actual_percent - e.planned_percent
        e.status = get_status_from_cumulative_variance(e.cumulative_actual, e.cumulative_planned, total_qty)
    db.commit()

def format_worksheet(ws, header_row):
    thin = Side(style='thin', color='D9D9D9')
    for row in ws.iter_rows():
        for cell in row:
            cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
            cell.border = Border(left=thin, right=thin, top=thin, bottom=thin)
    for col in range(1, ws.max_column + 1):
        cell = ws.cell(row=header_row, column=col)
        cell.font = Font(bold=True, color='FFFFFF', size=11)
        cell.fill = PatternFill('solid', fgColor='4472C4')
    for col_cells in ws.columns:
        max_len = 0
        for cell in col_cells:
            try:
                max_len = max(max_len, len(str(cell.value or '')))
            except:
                pass
        ws.column_dimensions[get_column_letter(col_cells[0].column)].width = min(max_len + 5, 55)

def get_setting(db: Session, key: str, default: str = ""):
    setting = db.query(SystemSetting).filter(SystemSetting.key == key).first()
    return setting.value if setting else default

def set_setting(db: Session, key: str, value: str):
    setting = db.query(SystemSetting).filter(SystemSetting.key == key).first()
    if setting:
        setting.value = value
    else:
        setting = SystemSetting(key=key, value=value)
        db.add(setting)
    db.commit()

# ================= FastAPI app =================
app = FastAPI(title="AShawkyDPSR API", version="1.0")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

@app.on_event("startup")
def startup():
    Base.metadata.create_all(bind=engine)
    db = SessionLocal()
    # Create default admin
    admin = db.query(User).filter(User.username == "admin").first()
    if not admin:
        hashed = get_password_hash("asd4cats!@")
        new_admin = User(username="admin", password=hashed, full_name="Administrator", role="Admin", enabled=True, created_at=datetime.utcnow())
        db.add(new_admin)
        db.commit()
    # Create default license
    lic = db.query(License).first()
    if not lic:
        default_expiry = (datetime.utcnow() + timedelta(days=7)).date()
        lic = License(expiry_date=default_expiry)
        db.add(lic)
        db.commit()
    # Create default project
    proj = db.query(Project).filter(Project.project_code == "RFC-001").first()
    if not proj:
        proj = Project(project_code="RFC-001", project_name="Rabigh Fuel Conversion", client_name="SE Saudi Energy", status="Active")
        db.add(proj)
        db.commit()
    db.close()

# ================= Auth endpoints =================
@app.post("/token")
async def login(form_data: OAuth2PasswordRequestForm = Depends(), db: Session = Depends(get_db)):
    user = authenticate_user(db, form_data.username, form_data.password)
    if not user:
        raise HTTPException(status_code=401, detail="Incorrect username or password")
    access_token_expires = timedelta(minutes=ACCESS_TOKEN_EXPIRE_MINUTES)
    access_token = create_access_token(data={"sub": user.username}, expires_delta=access_token_expires)
    user.last_login = datetime.utcnow()
    db.commit()
    log_audit(db, user.username, f"Logged in")
    return {"access_token": access_token, "token_type": "bearer", "role": user.role, "user_id": user.id, "full_name": user.full_name}

@app.get("/users/me")
async def read_users_me(current_user: User = Depends(get_current_user)):
    return {"id": current_user.id, "username": current_user.username, "full_name": current_user.full_name, "role": current_user.role}

# ================= User Management (Admin only) =================
@app.get("/users")
async def get_users(current_user: User = Depends(get_current_user), db: Session = Depends(get_db)):
    if current_user.role != "Admin":
        raise HTTPException(status_code=403, detail="Not allowed")
    users = db.query(User).all()
    return [{"id": u.id, "username": u.username, "full_name": u.full_name, "role": u.role, "enabled": u.enabled, "last_login": u.last_login} for u in users]

@app.post("/users")
async def create_user(user: UserCreate, current_user: User = Depends(get_current_user), db: Session = Depends(get_db)):
    if current_user.role != "Admin":
        raise HTTPException(status_code=403, detail="Not allowed")
    existing = db.query(User).filter(User.username == user.username).first()
    if existing:
        raise HTTPException(status_code=400, detail="Username already exists")
    hashed = get_password_hash(user.password)
    new_user = User(username=user.username, password=hashed, full_name=user.full_name, role=user.role, enabled=True, created_at=datetime.utcnow())
    db.add(new_user)
    db.commit()
    log_audit(db, current_user.username, f"Created user {user.username}")
    return {"message": "User created"}

@app.put("/users/{user_id}")
async def update_user(user_id: int, user: UserUpdate, current_user: User = Depends(get_current_user), db: Session = Depends(get_db)):
    if current_user.role != "Admin":
        raise HTTPException(status_code=403, detail="Not allowed")
    db_user = db.query(User).filter(User.id == user_id).first()
    if not db_user:
        raise HTTPException(status_code=404, detail="User not found")
    db_user.full_name = user.full_name
    db_user.role = user.role
    db_user.enabled = user.enabled
    if user.password:
        db_user.password = get_password_hash(user.password)
    db.commit()
    log_audit(db, current_user.username, f"Updated user {db_user.username}")
    return {"message": "User updated"}

@app.delete("/users/{user_id}")
async def delete_user(user_id: int, current_user: User = Depends(get_current_user), db: Session = Depends(get_db)):
    if current_user.role != "Admin":
        raise HTTPException(status_code=403, detail="Not allowed")
    db_user = db.query(User).filter(User.id == user_id).first()
    if not db_user:
        raise HTTPException(status_code=404, detail="User not found")
    if db_user.username == "admin":
        raise HTTPException(status_code=400, detail="Cannot delete admin")
    db.delete(db_user)
    db.commit()
    log_audit(db, current_user.username, f"Deleted user {db_user.username}")
    return {"message": "User deleted"}

# ================= License endpoints =================
@app.get("/license/status")
async def license_status(db: Session = Depends(get_db)):
    lic = db.query(License).first()
    if not lic:
        return {"valid": False, "expiry": None}
    today = datetime.utcnow().date()
    valid = lic.expiry_date >= today
    return {"valid": valid, "expiry": lic.expiry_date.isoformat()}

@app.post("/license/extend")
async def extend_license(expiry_date: str, current_user: User = Depends(get_current_user), db: Session = Depends(get_db)):
    if current_user.role != "Admin":
        raise HTTPException(status_code=403, detail="Only Admin can extend license")
    try:
        new_expiry = datetime.strptime(expiry_date, "%Y-%m-%d").date()
        lic = db.query(License).first()
        if not lic:
            lic = License(id=1)
            db.add(lic)
        lic.expiry_date = new_expiry
        lic.last_extended_by = current_user.username
        lic.extended_at = datetime.utcnow()
        db.commit()
        log_audit(db, current_user.username, f"Extended license to {expiry_date}")
        return {"message": f"License extended to {expiry_date}"}
    except:
        raise HTTPException(status_code=400, detail="Invalid date format")

# ================= Projects CRUD =================
@app.get("/projects")
async def get_projects(db: Session = Depends(get_db)):
    projects = db.query(Project).filter(Project.status == "Active").all()
    return [{"id": p.id, "project_code": p.project_code, "project_name": p.project_name, "client_name": p.client_name, "start_date": p.start_date.isoformat() if p.start_date else None, "end_date": p.end_date.isoformat() if p.end_date else None, "status": p.status} for p in projects]

@app.post("/projects")
async def create_project(proj: ProjectCreate, current_user: User = Depends(get_current_user), db: Session = Depends(get_db)):
    if current_user.role != "Admin":
        raise HTTPException(status_code=403, detail="Not allowed")
    existing = db.query(Project).filter(Project.project_code == proj.project_code).first()
    if existing:
        raise HTTPException(status_code=400, detail="Project code exists")
    new_proj = Project(
        project_code=proj.project_code,
        project_name=proj.project_name,
        client_name=proj.client_name,
        start_date=datetime.strptime(proj.start_date, "%Y-%m-%d").date(),
        end_date=datetime.strptime(proj.end_date, "%Y-%m-%d").date(),
        status=proj.status
    )
    db.add(new_proj)
    db.commit()
    log_audit(db, current_user.username, f"Created project {proj.project_code}")
    return {"message": "Project created"}

@app.put("/projects/{project_code}")
async def update_project(project_code: str, proj: ProjectUpdate, current_user: User = Depends(get_current_user), db: Session = Depends(get_db)):
    if current_user.role != "Admin":
        raise HTTPException(status_code=403, detail="Not allowed")
    db_proj = db.query(Project).filter(Project.project_code == project_code).first()
    if not db_proj:
        raise HTTPException(status_code=404, detail="Project not found")
    db_proj.project_code = proj.project_code
    db_proj.project_name = proj.project_name
    db_proj.client_name = proj.client_name
    db_proj.start_date = datetime.strptime(proj.start_date, "%Y-%m-%d").date()
    db_proj.end_date = datetime.strptime(proj.end_date, "%Y-%m-%d").date()
    db_proj.status = proj.status
    db.commit()
    log_audit(db, current_user.username, f"Updated project {project_code}")
    return {"message": "Project updated"}

@app.delete("/projects/{project_code}")
async def delete_project(project_code: str, current_user: User = Depends(get_current_user), db: Session = Depends(get_db)):
    if current_user.role != "Admin":
        raise HTTPException(status_code=403, detail="Not allowed")
    if project_code == "RFC-001":
        raise HTTPException(status_code=400, detail="Cannot delete default project")
    db_proj = db.query(Project).filter(Project.project_code == project_code).first()
    if not db_proj:
        raise HTTPException(status_code=404, detail="Project not found")
    db.delete(db_proj)
    db.commit()
    log_audit(db, current_user.username, f"Deleted project {project_code}")
    return {"message": "Project deleted"}

# ================= Activities CRUD =================
@app.get("/activities")
async def get_activities(db: Session = Depends(get_db)):
    acts = db.query(Activity).all()
    return [{"id": a.id, "activity_code": a.activity_code, "activity_name": a.activity_name, "discipline": a.discipline, "total_quantity": a.total_quantity, "unit": a.unit, "baseline_daily_qty": a.baseline_daily_qty, "critical": a.critical, "planned_start": a.planned_start.isoformat() if a.planned_start else None, "planned_finish": a.planned_finish.isoformat() if a.planned_finish else None} for a in acts]

@app.post("/activities")
async def create_activity(act: ActivityCreate, db: Session = Depends(get_db), current_user: User = Depends(get_current_user)):
    if current_user.role not in ["Admin", "Planner"]:
        raise HTTPException(status_code=403, detail="Permission denied")
    # Check for planner permission setting
    if current_user.role == "Planner":
        allow = get_setting(db, "allow_planner_edit_activities", "false")
        if allow.lower() != "true":
            raise HTTPException(status_code=403, detail="Planner activity editing not allowed by admin")
    existing = db.query(Activity).filter(Activity.activity_code == act.activity_code).first()
    if existing:
        raise HTTPException(status_code=400, detail="Activity code already exists")
    new_act = Activity(**act.dict())
    db.add(new_act)
    db.commit()
    db.refresh(new_act)
    log_audit(db, current_user.username, f"Created activity {act.activity_code}")
    return {"id": new_act.id}

@app.put("/activities/{activity_code}")
async def update_activity(activity_code: str, act: ActivityCreate, db: Session = Depends(get_db), current_user: User = Depends(get_current_user)):
    if current_user.role not in ["Admin", "Planner"]:
        raise HTTPException(status_code=403, detail="Permission denied")
    if current_user.role == "Planner":
        allow = get_setting(db, "allow_planner_edit_activities", "false")
        if allow.lower() != "true":
            raise HTTPException(status_code=403, detail="Planner activity editing not allowed by admin")
    db_act = db.query(Activity).filter(Activity.activity_code == activity_code).first()
    if not db_act:
        raise HTTPException(status_code=404, detail="Activity not found")
    for key, value in act.dict().items():
        setattr(db_act, key, value)
    db.commit()
    log_audit(db, current_user.username, f"Updated activity {activity_code}")
    return {"message": "Updated"}

@app.delete("/activities/{activity_code}")
async def delete_activity(activity_code: str, db: Session = Depends(get_db), current_user: User = Depends(get_current_user)):
    if current_user.role != "Admin":
        raise HTTPException(status_code=403, detail="Only Admin can delete")
    db_act = db.query(Activity).filter(Activity.activity_code == activity_code).first()
    if not db_act:
        raise HTTPException(status_code=404, detail="Not found")
    db.delete(db_act)
    db.commit()
    log_audit(db, current_user.username, f"Deleted activity {activity_code}")
    return {"message": "Deleted"}

# ================= Daily Progress endpoints =================
@app.post("/daily")
async def submit_daily(
    report_date: str = Form(...),
    activity_code: str = Form(...),
    planned_quantity: float = Form(...),
    actual_quantity: float = Form(...),
    manpower: int = Form(...),
    equipment: str = Form(""),
    material: str = Form(""),
    issues: str = Form(""),
    next_day_plan: str = Form(""),
    photos: List[UploadFile] = File(None),
    db: Session = Depends(get_db),
    current_user: User = Depends(get_current_user)
):
    act = db.query(Activity).filter(Activity.activity_code == activity_code).first()
    if not act:
        raise HTTPException(status_code=404, detail="Activity not found")
    new_progress = DailyProgress(
        report_date=datetime.strptime(report_date, "%Y-%m-%d").date(),
        activity_id=act.id,
        engineer_name=current_user.full_name or current_user.username,
        planned_quantity=planned_quantity,
        actual_quantity=actual_quantity,
        manpower=manpower,
        equipment=equipment,
        material=material,
        issues=issues,
        next_day_plan=next_day_plan,
        user_id=current_user.id,
        submitted_by=current_user.username,
        submitted_at=datetime.utcnow()
    )
    db.add(new_progress)
    db.commit()
    db.refresh(new_progress)
    calculate_cumulative_and_status(act.id, db)
    if photos:
        for photo in photos:
            upload_result = cloudinary.uploader.upload(photo.file, folder="AShawkyDPSR", public_id=f"{act.activity_code}_{uuid.uuid4().hex}")
            photo_record = ReportPhoto(
                progress_id=new_progress.id,
                cloudinary_public_id=upload_result['public_id'],
                photo_url=upload_result['secure_url'],
                original_filename=photo.filename,
                uploaded_by=current_user.username
            )
            db.add(photo_record)
        db.commit()
    log_audit(db, current_user.username, f"Submitted daily report for {activity_code}")
    return {"message": "Daily report submitted", "id": new_progress.id}

@app.get("/daily/recent")
async def get_recent_entries(limit: int = 50, db: Session = Depends(get_db), current_user: User = Depends(get_current_user)):
    query = db.query(DailyProgress).join(Activity).order_by(DailyProgress.report_date.desc())
    if current_user.role not in ["Admin", "Planner"]:
        query = query.filter(DailyProgress.user_id == current_user.id)
    entries = query.limit(limit).all()
    result = []
    for e in entries:
        result.append({
            "id": e.id,
            "report_date": e.report_date.isoformat(),
            "activity_code": e.activity.activity_code,
            "activity_name": e.activity.activity_name,
            "engineer_name": e.engineer_name,
            "planned_quantity": e.planned_quantity,
            "actual_quantity": e.actual_quantity,
            "unit": e.activity.unit,
            "cumulative_actual": e.cumulative_actual,
            "status": e.status
        })
    return result

@app.get("/daily/entry/{entry_id}")
async def get_entry(entry_id: int, db: Session = Depends(get_db), current_user: User = Depends(get_current_user)):
    entry = db.query(DailyProgress).filter(DailyProgress.id == entry_id).first()
    if not entry:
        raise HTTPException(status_code=404, detail="Not found")
    if current_user.role not in ["Admin", "Planner"] and entry.user_id != current_user.id:
        raise HTTPException(status_code=403, detail="Not allowed")
    return {
        "id": entry.id,
        "report_date": entry.report_date.isoformat(),
        "activity_code": entry.activity.activity_code,
        "planned_quantity": entry.planned_quantity,
        "actual_quantity": entry.actual_quantity,
        "manpower": entry.manpower,
        "equipment": entry.equipment,
        "material": entry.material,
        "issues": entry.issues,
        "next_day_plan": entry.next_day_plan,
        "submitted_at": entry.submitted_at.isoformat()
    }

@app.put("/daily/entry/{entry_id}")
async def update_entry(entry_id: int, data: DailyProgressUpdate, db: Session = Depends(get_db), current_user: User = Depends(get_current_user)):
    entry = db.query(DailyProgress).filter(DailyProgress.id == entry_id).first()
    if not entry:
        raise HTTPException(status_code=404, detail="Not found")
    if current_user.role not in ["Admin", "Planner"]:
        if entry.user_id != current_user.id:
            raise HTTPException(status_code=403, detail="Not allowed")
        # Check edit window
        edit_window_hours = int(get_setting(db, "edit_window_hours", "24"))
        time_diff = datetime.utcnow() - entry.submitted_at
        if time_diff.total_seconds() > edit_window_hours * 3600:
            raise HTTPException(status_code=403, detail=f"Edit window ({edit_window_hours} hours) expired")
    entry.planned_quantity = data.planned_quantity
    entry.actual_quantity = data.actual_quantity
    entry.manpower = data.manpower
    entry.equipment = data.equipment
    entry.material = data.material
    entry.issues = data.issues
    entry.next_day_plan = data.next_day_plan
    db.commit()
    calculate_cumulative_and_status(entry.activity_id, db)
    log_audit(db, current_user.username, f"Updated daily report entry {entry_id}")
    return {"message": "Entry updated"}

@app.delete("/daily/entry/{entry_id}")
async def delete_entry(entry_id: int, db: Session = Depends(get_db), current_user: User = Depends(get_current_user)):
    if current_user.role != "Admin":
        raise HTTPException(status_code=403, detail="Only Admin can delete")
    entry = db.query(DailyProgress).filter(DailyProgress.id == entry_id).first()
    if not entry:
        raise HTTPException(status_code=404, detail="Not found")
    db.delete(entry)
    db.commit()
    calculate_cumulative_and_status(entry.activity_id, db)
    log_audit(db, current_user.username, f"Deleted daily report entry {entry_id}")
    return {"message": "Deleted"}

# ================= Progress Monitor =================
@app.get("/progress/activities")
async def get_activity_progress(db: Session = Depends(get_db), discipline: str = None):
    query = db.query(Activity)
    if discipline and discipline != "All":
        query = query.filter(Activity.discipline == discipline)
    activities = query.all()
    result = []
    for act in activities:
        latest_progress = db.query(DailyProgress).filter(DailyProgress.activity_id == act.id).order_by(DailyProgress.report_date.desc()).first()
        cum_actual = latest_progress.cumulative_actual if latest_progress else 0
        cum_planned = latest_progress.cumulative_planned if latest_progress else 0
        total_qty = act.total_quantity
        prog = calculate_progress(cum_actual, total_qty) if total_qty > 0 else 0
        if cum_actual >= total_qty:
            status = "✅ Completed"
        elif cum_actual >= cum_planned:
            status = "🟢 On Track"
        elif (cum_planned - cum_actual) / max(cum_planned, 1) > 0.1:
            status = "🔴 Critical"
        else:
            status = "🟡 Delayed"
        result.append({
            "activity_code": act.activity_code,
            "activity_name": act.activity_name,
            "discipline": act.discipline,
            "progress": prog,
            "installed": cum_actual,
            "total": total_qty,
            "unit": act.unit,
            "remaining": max(total_qty - cum_actual, 0),
            "planned_finish": act.planned_finish.isoformat() if act.planned_finish else None,
            "status": status
        })
    return result

# ================= Lookahead =================
@app.get("/lookahead")
async def get_lookahead_tasks(db: Session = Depends(get_db)):
    tasks = db.query(LookaheadTask).join(Activity).all()
    return [{"id": t.id, "activity_code": t.activity.activity_code, "activity_name": t.activity.activity_name, "planned_start": t.planned_start.isoformat() if t.planned_start else None, "planned_finish": t.planned_finish.isoformat() if t.planned_finish else None, "priority": t.priority, "constraint_text": t.constraint_text, "owner": t.owner, "status": t.status} for t in tasks]

@app.post("/lookahead")
async def create_lookahead_task(task: LookaheadCreate, db: Session = Depends(get_db), current_user: User = Depends(get_current_user)):
    if current_user.role not in ["Admin", "Planner"]:
        raise HTTPException(status_code=403, detail="Permission denied")
    act = db.query(Activity).filter(Activity.activity_code == task.activity_code).first()
    if not act:
        raise HTTPException(status_code=404, detail="Activity not found")
    new_task = LookaheadTask(
        activity_id=act.id,
        planned_start=datetime.strptime(task.planned_start, "%Y-%m-%d").date(),
        planned_finish=datetime.strptime(task.planned_finish, "%Y-%m-%d").date(),
        priority=task.priority,
        constraint_text=task.constraint_text,
        owner=task.owner,
        status="Planned"
    )
    db.add(new_task)
    db.commit()
    log_audit(db, current_user.username, f"Added lookahead task for {task.activity_code}")
    return {"message": "Task added"}

# ================= Audit Log =================
@app.get("/audit")
async def get_audit_logs(limit: int = 200, db: Session = Depends(get_db), current_user: User = Depends(get_current_user)):
    if current_user.role != "Admin":
        raise HTTPException(status_code=403, detail="Not allowed")
    logs = db.query(AuditLog).order_by(AuditLog.timestamp.desc()).limit(limit).all()
    return [{"id": l.id, "username": l.username, "action": l.action, "timestamp": l.timestamp.isoformat()} for l in logs]

# ================= Settings =================
@app.get("/settings")
async def get_settings(db: Session = Depends(get_db)):
    settings = {}
    keys = ["edit_window_hours", "monthly_report_cycle_type", "monthly_report_start_day", "allow_planner_edit_activities",
            "working_days", "reports_folder", "photos_base_folder", "backup_folder", "backup_frequency", "backup_retention"]
    for key in keys:
        val = get_setting(db, key, "")
        if key in ["edit_window_hours", "monthly_report_start_day", "backup_retention"]:
            try:
                val = int(val)
            except:
                val = 24 if key == "edit_window_hours" else 1 if key == "monthly_report_start_day" else 30
        elif key in ["monthly_report_cycle_type"]:
            if not val:
                val = "calendar"
        elif key in ["allow_planner_edit_activities"]:
            val = val.lower() == "true" if val else False
        elif key in ["working_days"]:
            if val:
                try:
                    val = json.loads(val)
                except:
                    val = [0,1,2,3,4]  # Sat-Thu
            else:
                val = [0,1,2,3,4]
        settings[key] = val
    return settings

@app.post("/settings")
async def update_settings(settings: SettingsUpdate, current_user: User = Depends(get_current_user), db: Session = Depends(get_db)):
    if current_user.role != "Admin":
        raise HTTPException(status_code=403, detail="Not allowed")
    set_setting(db, "edit_window_hours", str(settings.edit_window_hours))
    set_setting(db, "monthly_report_cycle_type", settings.monthly_report_cycle_type)
    set_setting(db, "monthly_report_start_day", str(settings.monthly_report_start_day))
    set_setting(db, "allow_planner_edit_activities", str(settings.allow_planner_edit_activities).lower())
    set_setting(db, "working_days", json.dumps(settings.working_days))
    set_setting(db, "reports_folder", settings.reports_folder)
    set_setting(db, "photos_base_folder", settings.photos_base_folder)
    set_setting(db, "backup_folder", settings.backup_folder)
    set_setting(db, "backup_frequency", settings.backup_frequency)
    set_setting(db, "backup_retention", str(settings.backup_retention))
    log_audit(db, current_user.username, "Updated system settings")
    return {"message": "Settings saved"}

# ================= Reports (Excel) =================
def get_progress_data(db: Session, start_date: str, end_date: str, aggregate=False):
    query = db.query(DailyProgress).join(Activity).filter(
        DailyProgress.report_date >= datetime.strptime(start_date, "%Y-%m-%d").date(),
        DailyProgress.report_date <= datetime.strptime(end_date, "%Y-%m-%d").date()
    )
    entries = query.all()
    if not entries:
        return pd.DataFrame()
    data = []
    for e in entries:
        data.append({
            "report_date": e.report_date.isoformat(),
            "activity_code": e.activity.activity_code,
            "activity_name": e.activity.activity_name,
            "discipline": e.activity.discipline,
            "engineer_name": e.engineer_name,
            "planned_quantity": e.planned_quantity,
            "actual_quantity": e.actual_quantity,
            "unit": e.activity.unit,
            "cumulative_planned": e.cumulative_planned,
            "cumulative_actual": e.cumulative_actual,
            "remaining_quantity": e.remaining_quantity,
            "planned_percent": e.planned_percent,
            "actual_percent": e.actual_percent,
            "variance": e.variance,
            "manpower": e.manpower,
            "equipment": e.equipment,
            "material": e.material,
            "issues": e.issues,
            "next_day_plan": e.next_day_plan,
            "status": e.status
        })
    df = pd.DataFrame(data)
    if aggregate:
        # Aggregate per activity: sum planned/actual, last cumulative
        agg = df.groupby(['activity_code','activity_name','discipline','unit']).agg({
            'planned_quantity': 'sum',
            'actual_quantity': 'sum',
            'cumulative_planned': 'last',
            'cumulative_actual': 'last',
            'remaining_quantity': 'last',
            'planned_percent': 'last',
            'actual_percent': 'last',
            'variance': 'last',
            'manpower': 'mean',
            'status': 'last'
        }).reset_index()
        agg['manpower'] = agg['manpower'].round(0)
        agg['productivity'] = agg['actual_quantity'] / agg['manpower'].fillna(1)
        agg['productivity'] = agg['productivity'].round(2)
        agg.columns = ['Activity Code','Activity Name','Discipline','Unit',
                       'Planned Qty','Actual Qty','Cumulative Planned','Cumulative Actual',
                       'Remaining','Planned %','Actual %','Variance %','Avg Manpower','Status','Productivity']
        return agg
    else:
        df['Productivity'] = (df['actual_quantity'] / df['manpower'].replace(0, 1)).round(2)
        df.rename(columns={
            'report_date': 'Date','activity_code': 'Activity Code','activity_name': 'Activity Name',
            'discipline': 'Discipline','engineer_name': 'Engineer','planned_quantity': 'Planned Qty',
            'actual_quantity': 'Actual Qty','unit': 'Unit','cumulative_planned': 'Cumulative Planned',
            'cumulative_actual': 'Cumulative Actual','remaining_quantity': 'Remaining',
            'planned_percent': 'Planned %','actual_percent': 'Actual %','variance': 'Variance %',
            'manpower': 'Manpower','equipment': 'Equipment','material': 'Material','issues': 'Issues',
            'next_day_plan': 'Next Day Plan','status': 'Status'
        }, inplace=True)
        return df

def export_excel(rtype, df, period):
    if df.empty:
        return None
    wb = Workbook()
    ws = wb.active
    ws.merge_cells('A1:Z1'); ws['A1'] = 'RABIGH FUEL CONVERSION PROJECT'
    ws.merge_cells('A2:Z2'); ws['A2'] = f'{rtype.upper()} PROGRESS REPORT'
    ws.merge_cells('A3:Z3'); ws['A3'] = 'Planning & Controls Department'
    ws['A4'] = 'Contractor'; ws['B4'] = 'DOOSAN'
    ws['D4'] = 'Client'; ws['E4'] = 'SE Saudi Energy'
    ws['G4'] = 'Project'; ws['H4'] = 'RFC'
    ws['A5'] = 'Generated By:'; ws['B5'] = 'Eng. Ahmed Shawky'
    ws['D5'] = 'Period:'; ws['E5'] = period
    ws['G5'] = 'Report Date:'; ws['H5'] = datetime.now().strftime('%d/%m/%Y')
    header = 7
    for col_idx, col_name in enumerate(df.columns, 1):
        ws.cell(row=header, column=col_idx, value=col_name)
    for row_idx, row in enumerate(df.itertuples(index=False), header+1):
        for col_idx, val in enumerate(row, 1):
            ws.cell(row=row_idx, column=col_idx, value=val)
    format_worksheet(ws, header)
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output

@app.post("/reports/daily")
async def export_daily_report(date: str, db: Session = Depends(get_db)):
    df = get_progress_data(db, date, date, aggregate=False)
    if df.empty:
        raise HTTPException(status_code=404, detail="No data for this date")
    output = export_excel("Daily", df, date)
    return Response(content=output.getvalue(), media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", headers={"Content-Disposition": f"attachment; filename=daily_report_{date}.xlsx"})

@app.post("/reports/weekly")
async def export_weekly_report(db: Session = Depends(get_db)):
    today = datetime.now().date()
    # Monday to Friday of current week
    monday = today - timedelta(days=today.weekday())
    friday = monday + timedelta(days=4)
    df = get_progress_data(db, monday.isoformat(), friday.isoformat(), aggregate=True)
    if df.empty:
        raise HTTPException(status_code=404, detail="No data for this week")
    period = f"{monday.strftime('%d/%m/%Y')} to {friday.strftime('%d/%m/%Y')}"
    output = export_excel("Weekly", df, period)
    return Response(content=output.getvalue(), media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", headers={"Content-Disposition": "attachment; filename=weekly_report.xlsx"})

@app.post("/reports/monthly")
async def export_monthly_report(db: Session = Depends(get_db)):
    today = datetime.now().date()
    first_day = today.replace(day=1)
    if today.month == 12:
        last_day = today.replace(year=today.year+1, month=1, day=1) - timedelta(days=1)
    else:
        last_day = today.replace(month=today.month+1, day=1) - timedelta(days=1)
    df = get_progress_data(db, first_day.isoformat(), last_day.isoformat(), aggregate=True)
    if df.empty:
        raise HTTPException(status_code=404, detail="No data for this month")
    period = f"{first_day.strftime('%d/%m/%Y')} to {last_day.strftime('%d/%m/%Y')}"
    output = export_excel("Monthly", df, period)
    return Response(content=output.getvalue(), media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", headers={"Content-Disposition": "attachment; filename=monthly_report.xlsx"})

@app.post("/reports/full")
async def export_full_database(db: Session = Depends(get_db)):
    df = get_progress_data(db, "2020-01-01", "2030-12-31", aggregate=False)
    if df.empty:
        raise HTTPException(status_code=404, detail="No data")
    output = export_excel("Full Database", df, "All records")
    return Response(content=output.getvalue(), media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", headers={"Content-Disposition": "attachment; filename=full_export.xlsx"})

@app.post("/reports/pdf")
async def export_pdf_report(db: Session = Depends(get_db)):
    # Simple PDF report – can be expanded
    try:
        output = io.BytesIO()
        doc = SimpleDocTemplate(output, pagesize=landscape(A4))
        styles = getSampleStyleSheet()
        story = []
        story.append(Paragraph("RABIGH FUEL CONVERSION PROJECT", ParagraphStyle('CustomTitle', parent=styles['Title'], fontSize=16, textColor=colors.HexColor('#1a3a5c'))))
        story.append(Paragraph("Progress Report", styles['Heading2']))
        story.append(Spacer(1,20))
        # Get activity summary
        rows = db.query(Activity).all()
        data = [['Code','Activity','Discipline','Total Qty','Unit']]
        for a in rows:
            data.append([a.activity_code, a.activity_name[:30], a.discipline, a.total_quantity, a.unit])
        t = Table(data)
        t.setStyle(TableStyle([('BACKGROUND',(0,0),(-1,0),colors.HexColor('#1a3a5c')),('TEXTCOLOR',(0,0),(-1,0),colors.whitesmoke),('ALIGN',(0,0),(-1,-1),'CENTER'),('FONTNAME',(0,0),(-1,0),'Helvetica-Bold'),('FONTSIZE',(0,0),(-1,0),10),('BOTTOMPADDING',(0,0),(-1,0),12),('BACKGROUND',(0,1),(-1,-1),colors.beige),('GRID',(0,0),(-1,-1),1,colors.black)]))
        story.append(t)
        doc.build(story)
        output.seek(0)
        return Response(content=output.getvalue(), media_type="application/pdf", headers={"Content-Disposition": "attachment; filename=report.pdf"})
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

# ================= AI Forecast =================
@app.get("/forecast")
async def get_forecast(db: Session = Depends(get_db)):
    rows = db.query(DailyProgress.report_date, func.sum(DailyProgress.actual_quantity).label("qty")).group_by(DailyProgress.report_date).order_by(DailyProgress.report_date).all()
    if len(rows) < 2:
        return {"message": "Not enough data for forecast"}
    qty_data = [r.qty for r in rows]
    rate = sum(qty_data[-min(14, len(qty_data)):]) / min(14, len(qty_data))
    total = db.query(func.sum(Activity.total_quantity)).scalar() or 0
    installed = db.query(func.sum(DailyProgress.cumulative_actual)).scalar() or 0
    remaining = max(total - installed, 0)
    days = int(remaining / rate) if rate > 0 else 365
    finish = datetime.now() + timedelta(days=days)
    return {"rate": rate, "remaining": remaining, "days": days, "completion_date": finish.strftime("%Y-%m-%d")}

# ================= Import / Export Activities =================
@app.post("/activities/import")
async def import_activities(file: UploadFile = File(...), current_user: User = Depends(get_current_user), db: Session = Depends(get_db)):
    if current_user.role not in ["Admin", "Planner"]:
        raise HTTPException(status_code=403, detail="Permission denied")
    if current_user.role == "Planner":
        allow = get_setting(db, "allow_planner_edit_activities", "false")
        if allow.lower() != "true":
            raise HTTPException(status_code=403, detail="Planner activity editing not allowed")
    df = pd.read_excel(file.file)
    imported = 0
    for _, row in df.iterrows():
        code = row.get("Activity Code")
        if not code:
            continue
        exists = db.query(Activity).filter(Activity.activity_code == code).first()
        if not exists:
            act = Activity(
                activity_code=code,
                activity_name=row.get("Activity Name", ""),
                discipline=row.get("Discipline", ""),
                total_quantity=row.get("Quantity", 0),
                unit=row.get("Unit", "%"),
                baseline_daily_qty=row.get("Baseline Daily Qty", 0),
                critical=1 if row.get("Critical") == "Yes" else 0,
                planned_start=row.get("Planned Start") if pd.notna(row.get("Planned Start")) else None,
                planned_finish=row.get("Planned Finish") if pd.notna(row.get("Planned Finish")) else None,
                project_id=1
            )
            db.add(act)
            imported += 1
    db.commit()
    log_audit(db, current_user.username, f"Imported {imported} activities")
    return {"message": f"Imported {imported} activities"}

@app.get("/activities/export")
async def export_activities(db: Session = Depends(get_db)):
    acts = db.query(Activity).all()
    data = [{
        "Activity Code": a.activity_code,
        "Activity Name": a.activity_name,
        "Discipline": a.discipline,
        "Quantity": a.total_quantity,
        "Unit": a.unit,
        "Baseline Daily Qty": a.baseline_daily_qty,
        "Critical": "Yes" if a.critical else "No",
        "Planned Start": a.planned_start.isoformat() if a.planned_start else "",
        "Planned Finish": a.planned_finish.isoformat() if a.planned_finish else ""
    } for a in acts]
    df = pd.DataFrame(data)
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name="Activities", index=False)
    output.seek(0)
    return Response(content=output.getvalue(), media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", headers={"Content-Disposition": "attachment; filename=activities.xlsx"})

# ================= Run =================
if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=10000)
