from sqlalchemy import create_engine
from sqlalchemy.orm import sessionmaker
import urllib

params = urllib.parse.quote_plus(
    "Driver={ODBC Driver 18 for SQL Server};"
    "Server=tcp:subcruncher.database.windows.net,1433;"
    "Database=SubCruncherDB;"
    "Uid=subcruncher;"
    "Pwd=5BBE8766-47F2-4A58-82CD-87F8255DF8A0;"
    "Encrypt=yes;"
    "TrustServerCertificate=no;"
    "Connection Timeout=30;"
)

DATABASE_URL = "mssql+pyodbc:///?odbc_connect=" + params

engine = create_engine(DATABASE_URL, echo=True)
SessionLocal = sessionmaker(bind=engine, autocommit=False, autoflush=False)

def get_db():
    db = SessionLocal()
    try:
        yield db
    finally:
        db.close()