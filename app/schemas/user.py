from pydantic import BaseModel, EmailStr

class UserSignup(BaseModel):
    name: str
    email: EmailStr
    password: str

class UserSignin(BaseModel):
    email: EmailStr
    password: str