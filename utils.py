import sys
from os import listdir, mkdir, system, getcwd, remove
from zipfile import ZipFile

import requests
from json import *

UNINSTALLERS_PATH = fr"{getcwd()}\uninstall"
YAOTYPES : list[str] = ['C2R_DogfoodDevMain','C2R_InsiderFast']

def initOft(step: int):
    """初始化程序"""
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


def checkApi():
    """检测Github API是否可用"""
    response = None
    try:
        response = requests.get("https://raw.githubusercontent.com/PyCoreTeam/OfficeT/main/api.json", verify=False)
    except:
        pass
    if response.ok:
        return 'OK'
    del response


def getAPI():
    """获取API"""
    f = None
    try:
        with open("./apis/api.json", 'w') as f:
            f.write(
                str(requests.get("https://raw.githubusercontent.com/PyCoreTeam/OfficeT/main/api.json",
                                 verify=False).text))

    except Exception as e:
        print(e)
        sys.exit(0)
    finally:
        f.close()
        del f


def dumpApi():
    """将API转化为Python字典"""
    with open("./apis/api.json", 'r') as f:
        _f = f
        f.close()
        return load(_f)


def downloadODP(api):
    """下载Office Development Pack"""
    r = requests.get(api, verify=False)
    with open("./office.zip", 'wb') as f:
        f.write(r.content)
        f.close()
    del f, r


def downloadUninstalls(api):
    """下载卸载工具并解压"""
    r = requests.get(api, verify=False)
    try:
        mkdir("uninstall")
    except:
        pass
    with open("./uninstall/u.zip", 'wb') as f:
        f.write(r.content)
        f.close()
        del f
    with ZipFile("./uninstall/u.zip", 'r') as f:
        f.extractall("./uninstall")
        f.close()
        del f


def getIstOfficeConfig(p=False):
    """获取安装Office的配置"""
    if not p:

        return listdir("./configs")
    else:
        v = 0
        for i in listdir("./configs"):
            v += 1
            print(str(v), '.', i.replace(".xml", ''))
        return listdir("./configs")


def operationUninstaller(index: int):
    """运行合适的卸载工具"""
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


def getYaoCurlConfig(p=False):
    """获取YAOCTRU的配置"""
    if not p:
        return listdir("./YAOCTRU/yaoctru_curls")
    else:
        v = 0
        for i in listdir("./YAOCTRU/yaoctru_curls"):
            v += 1
            print(str(v), '.', i.replace(".bat", ''))
        # return listdir("./YAOCTRU/yaoctru_curls")
