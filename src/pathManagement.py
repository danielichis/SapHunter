from pathlib import Path
import os
def get_current_path():
    currentSrcPath = os.getcwd()
    path = Path(currentSrcPath)
    return path 
    
#get_current_path()