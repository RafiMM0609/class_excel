from minio import Minio, S3Error
from minio.datatypes import Object
from config import (
    MINIO_ENPOINT,
    MINIO_ACCESS_KEY,
    MINIO_SECRET_KEY,
    MINIO_SECURE,
    MINIO_BUCKET,
    STORAGE,
)
import os
from PIL import Image as PILImage

# Using Minio
minio_client = Minio(
    MINIO_ENPOINT,
    access_key=MINIO_ACCESS_KEY,
    secret_key=MINIO_SECRET_KEY,
    secure=MINIO_SECURE,
)

def download_file_to_bytes(
    path:str,
):
    if STORAGE == 'local':
        current_directory = os.getcwd()
        image_data = f'{current_directory}/example.png'
    else:
        response = minio_client.get_object(MINIO_BUCKET, path)
        image_data = response.read()
        response.close()
        response.release_conn()
    return image_data