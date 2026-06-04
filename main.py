# ================= COMPLETE BACKEND CODE =================
# Copy everything from here to the end

import os
import uuid
from datetime import datetime, timedelta
from fastapi import FastAPI, Depends, HTTPException, status, UploadFile, File, Form
from fastapi.security import OAuth2PasswordBearer, OAuth2PasswordRequestForm
from fastapi.middleware.cors import CORSMiddleware
from sqlalchemy import create_engine, Column, Integer, String, Float, Date, DateTime, Boolean, ForeignKey, Text
from sqlalchemy.ext.declarative import declarative_base
from sqlalchemy.orm import sessionmaker, Session
from passlib.context import CryptContext
from jose import JWTError, jwt
from typing import Optional, List
import cloudinary
import cloudinary.uploader
import pandas as pd
from openpyxl import Workbook
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
import io
from fastapi.responses import Response
from pydantic import BaseModel

# ================= Configuration =================
DATABASE_URL = os.environ.get("DATABASE_URL")
SECRET_KEY = os.environ.get("SECRET_KEY", "your-super-secret-key-change-in-production")
ALGORITHM = "HS256"
ACCESS_TOKEN_EXPIRE_MINUTES = 60 * 24

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

# ================= Helper functions =================
pwd_context = CryptContext(schemes=["bcrypt"], deprecated="auto")

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

async def get_current_user(token: str = Depends(OAuth2PasswordBearer(tokenUrl="token")), db: Session = Depends(get_db)):
    from fastapi import HTTPException
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

def get_db():
    db = SessionLocal()
    try:
        yield db
    finally:
        db.close()

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
    admin = db.query(User).filter(User.username == "admin").first()
    if not admin:
        hashed = get_password_hash("asd4cats!@")
        new_admin = User(username="admin", password=hashed, full_name="Administrator", role="Admin", enabled=True, created_at=datetime.utcnow())
        db.add(new_admin)
        db.commit()
    lic = db.query(License).first()
    if not lic:
        default_expiry = (datetime.utcnow() + timedelta(days=7)).date()
        lic = License(expiry_date=default_expiry)
        db.add(lic)
        db.commit()
    db.close()

@app.post("/token")
async def login(form_data: OAuth2PasswordRequestForm = Depends(), db: Session = Depends(get_db)):
    user = authenticate_user(db, form_data.username, form_data.password)
    if not user:
        raise HTTPException(status_code=401, detail="Incorrect username or password")
    access_token_expires = timedelta(minutes=ACCESS_TOKEN_EXPIRE_MINUTES)
    access_token = create_access_token(data={"sub": user.username}, expires_delta=access_token_expires)
    user.last_login = datetime.utcnow()
    db.commit()
    return {"access_token": access_token, "token_type": "bearer", "role": user.role, "user_id": user.id, "full_name": user.full_name}

@app.get("/users/me")
async def read_users_me(current_user: User = Depends(get_current_user)):
    return {"id": current_user.id, "username": current_user.username, "full_name": current_user.full_name, "role": current_user.role}

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
        return {"message": f"License extended to {expiry_date}"}
    except:
        raise HTTPException(status_code=400, detail="Invalid date format")

@app.get("/activities")
async def get_activities(db: Session = Depends(get_db)):
    acts = db.query(Activity).all()
    return [{"id": a.id, "activity_code": a.activity_code, "activity_name": a.activity_name, "discipline": a.discipline, "total_quantity": a.total_quantity, "unit": a.unit, "critical": a.critical, "planned_start": a.planned_start.isoformat() if a.planned_start else None, "planned_finish": a.planned_finish.isoformat() if a.planned_finish else None} for a in acts]

@app.post("/activities")
async def create_activity(act: ActivityCreate, db: Session = Depends(get_db), current_user: User = Depends(get_current_user)):
    if current_user.role not in ["Admin", "Planner"]:
        raise HTTPException(status_code=403, detail="Permission denied")
    existing = db.query(Activity).filter(Activity.activity_code == act.activity_code).first()
    if existing:
        raise HTTPException(status_code=400, detail="Activity code already exists")
    new_act = Activity(**act.dict())
    db.add(new_act)
    db.commit()
    db.refresh(new_act)
    return {"id": new_act.id}

@app.put("/activities/{activity_code}")
async def update_activity(activity_code: str, act: ActivityCreate, db: Session = Depends(get_db), current_user: User = Depends(get_current_user)):
    if current_user.role not in ["Admin", "Planner"]:
        raise HTTPException(status_code=403, detail="Permission denied")
    db_act = db.query(Activity).filter(Activity.activity_code == activity_code).first()
    if not db_act:
        raise HTTPException(status_code=404, detail="Activity not found")
    for key, value in act.dict().items():
        setattr(db_act, key, value)
    db.commit()
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
    return {"message": "Deleted"}

def calculate_progress(installed, total):
    if total <= 0: return 0
    return round((installed / total) * 100, 2)

def get_status_from_cumulative_variance(actual, planned, total):
    if actual >= total and total > 0: return "Completed"
    var = actual - planned
    if var >= 0: return "On Track"
    elif var >= -10: return "Delayed"
    else: return "Critical"

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
async def update_entry(entry_id: int, data: dict, db: Session = Depends(get_db), current_user: User = Depends(get_current_user)):
    entry = db.query(DailyProgress).filter(DailyProgress.id == entry_id).first()
    if not entry:
        raise HTTPException(status_code=404, detail="Not found")
    if current_user.role not in ["Admin", "Planner"]:
        if entry.user_id != current_user.id:
            raise HTTPException(status_code=403, detail="Not allowed")
        time_diff = datetime.utcnow() - entry.submitted_at
        if time_diff.total_seconds() > 24 * 3600:
            raise HTTPException(status_code=403, detail="Edit window (24h) expired")
    entry.planned_quantity = data.get("planned_quantity", entry.planned_quantity)
    entry.actual_quantity = data.get("actual_quantity", entry.actual_quantity)
    entry.manpower = data.get("manpower", entry.manpower)
    entry.equipment = data.get("equipment", entry.equipment)
    entry.material = data.get("material", entry.material)
    entry.issues = data.get("issues", entry.issues)
    entry.next_day_plan = data.get("next_day_plan", entry.next_day_plan)
    db.commit()
    calculate_cumulative_and_status(entry.activity_id, db)
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
    return {"message": "Deleted"}

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
        status = "✅ Completed" if cum_actual >= total_qty else "🟢 On Track" if cum_actual >= cum_planned else "🔴 Critical" if (cum_planned - cum_actual) / max(cum_planned,1) > 0.1 else "🟡 Delayed"
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

# ================= Reports (Excel) =================
@app.post("/reports/daily")
async def export_daily_report(date: str, db: Session = Depends(get_db)):
    report_date = datetime.strptime(date, "%Y-%m-%d").date()
    entries = db.query(DailyProgress).filter(DailyProgress.report_date == report_date).all()
    if not entries:
        raise HTTPException(status_code=404, detail="No data")
    data = []
    for e in entries:
        data.append({
            "Date": e.report_date.isoformat(),
            "Activity Code": e.activity.activity_code,
            "Activity Name": e.activity.activity_name,
            "Discipline": e.activity.discipline,
            "Engineer": e.engineer_name,
            "Planned Qty": e.planned_quantity,
            "Actual Qty": e.actual_quantity,
            "Unit": e.activity.unit,
            "Cumulative Planned": e.cumulative_planned,
            "Cumulative Actual": e.cumulative_actual,
            "Remaining": e.remaining_quantity,
            "Planned %": e.planned_percent,
            "Actual %": e.actual_percent,
            "Variance %": e.variance,
            "Manpower": e.manpower,
            "Productivity": e.actual_quantity / e.manpower if e.manpower else 0,
            "Equipment": e.equipment,
            "Material": e.material,
            "Issues": e.issues,
            "Next Day Plan": e.next_day_plan,
            "Status": e.status
        })
    df = pd.DataFrame(data)
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name="Daily Report", index=False)
    output.seek(0)
    return Response(content=output.read(), media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", headers={"Content-Disposition": f"attachment; filename=daily_report_{date}.xlsx"})

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)