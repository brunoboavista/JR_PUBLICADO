import shutil
import os

# from_path = 'C:/Users/Bruno/Desktop/JR/STORAGE/'
# to_path = 'D:/Bruno/OneDrive/STORAGE/'

def copy_and_overwrite(from_path, to_path):
    if os.path.exists(to_path):
        shutil.rmtree(to_path)
    shutil.copytree(from_path, to_path)
        
copy_and_overwrite(from_path,to_path)