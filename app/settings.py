# # API Key settings

# from pydantic import BaseModel
# import os

# class Settings(BaseModel):
#     api_key: str = os.getenv("API_KEY", "change-me")
#     max_upload_mb: int = int(os.getenv("MAX_UPLOAD_MB", "500"))  # per file

# settings = Settings()