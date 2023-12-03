import sys
from os import listdir, mkdir, system, getcwd, remove
from zipfile import ZipFile

import requests
from json import *

UNINSTALLERS_PATH = fr"{getcwd()}\uninstall"


def checkApi():
    response = None
    try:
        response = requests.get("https://raw.githubusercontent.com/PyCoreTeam/OfficeT/main/api.json", verify=False)
    except:
        pass
    if response.ok:
        return 'OK'
    del response

def getAPI():
    f = None
    try:
        f = open("./apis/api.json", 'w')
        f.write(
            str(requests.get("https://raw.githubusercontent.com/PyCoreTeam/OfficeT/main/api.json", verify=False).text))
    except Exception as e:
        print(e)
        sys.exit(0)
    finally:
        f.close()
        del f


def dumpApi():
    f = open("./apis/api.json", 'r')
    f.close()
    return load(f)


def downloadODP(api):
    r = requests.get(api, verify=False)
    f = open("./office.zip", 'wb')
    f.write(r.content)
    f.close()
    del f


def getConfigs(p=False):
    if not p:

        return listdir("./configs")
    else:
        v = 0
        for i in listdir("./configs"):
            v += 1
            print(str(v), '.', i.replace(".xml", ''))
        return listdir("./configs")


def getUninstalls(api):
    r = requests.get(api, verify=False)
    try:
        mkdir("uninstall")
    except:
        pass
    f = open("./uninstall/u.zip", 'wb')
    f.write(r.content)
    f.close()
    with ZipFile("./uninstall/u.zip", 'r') as f:
        f.extractall("./uninstall")
        f.close()
    del f


def initOft(step: int):
    match step:
        case 1:
            try:
                mkdir("./apis")
            except:
                pass
        case 2:
            try:
                remove("./office.zip")
            except:
                pass
        case _:
            try:
                remove("./office/config.xml")
            except:
                pass


def operationUninstaller(index: int):
    match index:
        case 1:
            i = int(input("Office版本:\n1.2003\n2.2007\n3.2010"))
            match i:
                case 1:
                    system(fr"{UNINSTALLERS_PATH}\EasyFix\MicrosoftEasyFix20054.mini.diagcab")
                case 2:
                    system(fr"{UNINSTALLERS_PATH}\EasyFix\MicrosoftEasyFix20052.mini.diagcab")
                case 3:
                    system(fr"{UNINSTALLERS_PATH}\EasyFix\MicrosoftEasyFix20055.mini.diagcab")
                case _:
                    return
        case 2:
            system(fr"{UNINSTALLERS_PATH}\o15-ctrremove.diagcab")
        case 3:
            system(fr"{UNINSTALLERS_PATH}\ClickOnce\SaraSetup.exe")
        case 4:
            system(fr"{UNINSTALLERS_PATH}\uninstall\Troubleshoot\OffScrubC2R.vbs")
        case _:
            return
