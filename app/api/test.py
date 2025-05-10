from fastapi import APIRouter

router = APIRouter()

@router.post("")
async def test_endpoint():
    return {"status": "success"}