from pathlib import Path

def findTemplates(templatesPath):
        templatesPath = Path(templatesPath)
        if templatesPath.exists():
            return templatesPath
        if Path("../templates").exists():
            return Path("../templates")
        if Path("./templates").exists():
            return Path("./templates")
        
        return templatesPath

def findData(DataPath):
        DataPath = Path(DataPath)
        if DataPath.exists():
            return DataPath
        if Path("../data").exists():
            return Path("../data")
        if Path("./data").exists():
            return Path("./data")
        
        return DataPath