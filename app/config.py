from pydantic_settings import BaseSettings

class Settings(BaseSettings):
    MONGO_URL: str
    OPENAI_API_KEY: str

    class Config:
        env_file = ".env"

settings = Settings()