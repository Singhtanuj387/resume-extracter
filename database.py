import os
from datetime import datetime
from sqlalchemy import create_engine, Column, Integer, String, Text, DateTime
from sqlalchemy.orm import declarative_base, sessionmaker
from sqlalchemy.engine import make_url


def get_clean_database_url():
    """
    Cleans and validates the DATABASE_URL environment variable.
    Handles Render/Heroku postgres:// -> postgresql://, surrounding quotes, psql prefix,
    and automatically falls back to SQLite if the variable is empty or invalid.
    """
    raw_url = os.getenv("DATABASE_URL")
    if not raw_url:
        return "sqlite:///./resumes.db"

    url = raw_url.strip().strip("\"'").strip()
    if not url:
        return "sqlite:///./resumes.db"

    # If copied with the CLI command e.g. "psql postgresql://..."
    if url.startswith("psql "):
        url = url[5:].strip().strip("\"'").strip()

    # Fix legacy Render / Heroku prefix postgres:// -> postgresql://
    if url.startswith("postgres://"):
        url = url.replace("postgres://", "postgresql://", 1)

    try:
        make_url(url)
        return url
    except Exception as e:
        print(f"Warning: Unable to parse DATABASE_URL '{raw_url}': {e}. Falling back to SQLite.")
        return "sqlite:///./resumes.db"


DATABASE_URL = get_clean_database_url()
print(f"Connecting to database: {DATABASE_URL.split('@')[-1] if '@' in DATABASE_URL else DATABASE_URL}")

connect_args = {}
if DATABASE_URL.startswith("sqlite"):
    connect_args = {"check_same_thread": False}

engine = create_engine(DATABASE_URL, connect_args=connect_args, pool_pre_ping=True)
SessionLocal = sessionmaker(autocommit=False, autoflush=False, bind=engine)
Base = declarative_base()


class Resume(Base):
    __tablename__ = "resumes"

    id = Column(Integer, primary_key=True, index=True, autoincrement=True)
    job_id = Column(String(64), index=True, nullable=True)
    name = Column(String(255), index=True, nullable=True)
    contact_no = Column(String(100), nullable=True)
    email = Column(String(255), index=True, nullable=True)
    role = Column(String(100), index=True, nullable=True)
    skills = Column(Text, nullable=True)
    education = Column(Text, nullable=True)
    projects = Column(Text, nullable=True)
    experience = Column(Text, nullable=True)
    area_of_interest = Column(Text, nullable=True)
    awards = Column(Text, nullable=True)
    extra_curriculars = Column(Text, nullable=True)
    registration_no = Column(String(100), index=True, nullable=True)
    linkedin = Column(String(255), nullable=True)
    github = Column(String(255), nullable=True)
    source_file = Column(String(255), nullable=True)
    created_at = Column(DateTime, default=datetime.utcnow)

    def to_dict(self):
        """Converts model to dictionary matching Excel generator keys."""
        return {
            "id": self.id,
            "Name": self.name or "",
            "Contact No": self.contact_no or "",
            "Email": self.email or "",
            "Role": self.role or "Uncategorized",
            "Skills": self.skills or "",
            "Education": self.education or "",
            "Projects": self.projects or "",
            "Experience & Internships": self.experience or "",
            "Area of Interest / Objective": self.area_of_interest or "",
            "Awards & Achievements": self.awards or "",
            "Extra Curriculars & Leadership": self.extra_curriculars or "",
            "Registration No": self.registration_no or "",
            "LinkedIn": self.linkedin or "",
            "GitHub": self.github or "",
            "Source File": self.source_file or "",
            "created_at": self.created_at.strftime("%Y-%m-%d %H:%M") if self.created_at else "",
        }


def init_db():
    """Initializes database tables."""
    Base.metadata.create_all(bind=engine)


def get_db():
    """FastAPI dependency for obtaining a database session."""
    db = SessionLocal()
    try:
        yield db
    finally:
        db.close()
