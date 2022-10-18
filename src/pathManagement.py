from pathlib import Path
import os
def createFolder(path,force):
    try:
        if not os.path.exists(path):
            print("folder dont exits")
            if force:
                os.makedirs(path)
                print("folder created")
            return False
        else:
            print("Folder already exists")
            return True
    except OSError:
        print ('Error: Creating directory. ' +  path)
def get_current_path():
    currentSrcPath = os.getcwd()
    currentPath = Path(currentSrcPath)
    return currentPath 

#get_current_path()