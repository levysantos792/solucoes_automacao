import os

class envload():
    def __init__(self, file_path='.env'):
        self.load_env_file(file_path)

    def load_env_file(self, file_path):
        if os.path.exists(file_path):
            with open(file_path) as f:
                for line in f:
                    if line.strip() and not line.startswith('#'):
                        key, value = line.strip().split('=', 1)
                        os.environ[key] = value
