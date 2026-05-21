import os
import uuid

def save_uploaded_file(file, upload_folder):
    ext = os.path.splitext(file.filename)[-1]
    unique_name = f"{uuid.uuid4().hex}{ext}"
    file_path = os.path.join(upload_folder, unique_name)
    file.save(file_path)
    return file_path

def get_file_path(filename, folder):
    return os.path.join(folder, filename)
