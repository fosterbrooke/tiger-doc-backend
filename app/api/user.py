from fastapi import APIRouter, Depends, HTTPException, status
from sqlalchemy.ext.asyncio import AsyncSession
from sqlalchemy.future import select
from app.models.user import User
from app.schemas.user import UserSignup, UserSignin
from app.db.database import get_db
from app.utils.userUtils import hash_password, verify_password

router = APIRouter()

@router.post("/signup")
async def signup_user_endpoint(user: UserSignup, db: AsyncSession = Depends(get_db)):
    result = db.execute(select(User).where(User.email == user.email))
    existing_user = result.scalar_one_or_none()
    if existing_user:
        raise HTTPException(status_code=400, detail="Email already registered")
    
    new_user = User(
        name=user.name,
        email=user.email,
        password=hash_password(user.password)
    )
    db.add(new_user)
    db.commit()
    return {"message": "Signup successful"}

@router.post("/signin")
async def signin_user_endpoint(user: UserSignin, db: AsyncSession = Depends(get_db)):
    result = db.execute(select(User).where(User.email == user.email))
    db_user = result.scalar_one_or_none()
    if not db_user or not verify_password(user.password, db_user.password):
        raise HTTPException(status_code=401, detail="Invalid credentials")
    return {"message": "Signin successful", "user_id": db_user.id}