def read_file(file_path):
    with open(file_path, 'r') as file:
        return file.read()

def write_file(file_path, data):
    with open(file_path, 'w') as file:
        file.write(data)

def create_temp_file(suffix='', delete=True):
    import tempfile
    return tempfile.NamedTemporaryFile(suffix=suffix, delete=delete)

def normalize_file_path(file_path):
    import os
    return os.path.normpath(file_path)

def get_file_extension(file_path):
    import os
    return os.path.splitext(file_path)[1]