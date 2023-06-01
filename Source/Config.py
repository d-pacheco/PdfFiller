import yaml, requests
from yaml.parser import ParserError
from rich import print
from pathlib import Path

class Config:
    """
    A class that loads and stores the configurations
    """

    def __init__(self, configPath: str) -> None:
        self.files = {}
        try:
            configPath = self.__findConfig(configPath)
            with open(configPath, "r", encoding='utf-8') as f:
                config = yaml.safe_load(f)
                self.files = config.get("files") 
        except FileNotFoundError as ex:
            print(f"[red]CRITICAL ERROR: The configuration file cannot be found at {configPath}\nHave you extacted the ZIP archive and edited the configuration file?")
            print("Press any key to exit...")
            input()
            raise ex
        except (ParserError, KeyError) as ex:
            print(f"[red]CRITICAL ERROR: The configuration file does not have a valid format.\nPlease, check it for extra spaces and other characters.")
            print("Press any key to exit...")
            input()
            raise ex


    def GetFilesFromConfig(self):
        return list(self.files.keys())

        
    def GetFieldsForFile(self, fileName):
        return list(self.files[fileName].keys())
    

    def GetSheetAndCell(self, fileName, fieldName):
        try:
            value = self.files[fileName][fieldName]
            sheet, cell = value.split(':')
            return sheet, cell
        except ValueError as ex:
            print(f'[red]CRITICAL ERROR: The configuration file does not have a valid format.\nPlease, check all entries have format "sheet:cell"')
            print("Press any key to exit...")
            input()
            raise ex
        except Exception as ex:
            print(f'[red]CRITICAL ERROR: Something went wrong when getting sheet and cell')
            print("Press any key to exit...")
            input()
            raise ex
    
    
    def __findConfig(self, configPath):
        """
        Try to find configuartion file in alternative locations.

        :param configPath: user suplied configuartion file path
        :return: pathlib.Path, path to the configuration file
        """
        configPath = Path(configPath)
        if configPath.exists():
            return configPath
        if Path("../config/config.yaml").exists():
            return Path("../config/config.yaml")
        if Path("config/config.yaml").exists():
            return Path("config/config.yaml")
        
        return configPath