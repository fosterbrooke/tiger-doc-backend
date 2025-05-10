from fastapi import APIRouter, FastAPI
from fastapi.middleware.cors import CORSMiddleware
from app.api import process, user
# from app.db.database import connect_to_mongo, close_mongo_connection

app = FastAPI()

@app.on_event("startup")
async def startup_db_client():
    # await connect_to_mongo()
    print('startup')

@app.on_event("shutdown")
async def shutdown_db_client():
    # await close_mongo_connection()
    print('shutdown')

# Add CORS middleware
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

backend_router = APIRouter()

backend_router.include_router(user.router, prefix="/users", tags=["users"])
backend_router.include_router(process.router, prefix="/process", tags=["process"])

app.include_router(backend_router)