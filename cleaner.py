import os
import shutil


def folderCleaner():
    dirPath = "Images\\"
    #fileList = os.listdir(dirPath)
    #[os.remove(os.path.abspath(os.path.join(dirPath, fileName))) for fileName in fileList]
    if os.path.exists('unzip_dir\\'):
        shutil.rmtree('unzip_dir\\')
    if os.path.exists('Evidence\\'):
        shutil.rmtree('Evidence\\')
    for filename in os.listdir(dirPath):
        filepath = dirPath + filename
        if os.path.exists(filepath):
            os.unlink(filepath)
