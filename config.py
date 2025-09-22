import configparser
import os

def Checkfile(pathfile):
    if os.path.exists(pathfile):
        print(f"Config - {pathfile} file is exist.")
        return True
    else:
        # logging.info("file is not exist")
        raise FileNotFoundError(f"Config - {pathfile} file is not exist")


def loadvar(pathfile):
    if Checkfile(pathfile):
        #Read properties file
        config=configparser.ConfigParser()
        config.read(pathfile)

        name=config.get('DEFAULT','name')
        email=config.get('DEFAULT','email_address')
    else:
        name='Default'
        email='sample@outlook.com'

    return name,email


