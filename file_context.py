import sys

class Context():
    def __init__(self, fileName, filePath, status):
        self.fileName = fileName
        self.filePath = filePath
        self.status = status

    def reset(self):
        self.status = ""
        self.fileName = "[Untitled]"
        self.filePath = ""
