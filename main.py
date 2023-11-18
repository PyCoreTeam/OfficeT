from utils import *
from lang import *
from os import mkdir, remove, path, rename, system, getcwd
from shutil import move, copy
from zipfile import *
import subprocess
try:
    mkdir("./apis")
except:
    pass
try:
    remove("./office/office.zip")
except:
    pass
try:
    remove("./office.zip")
except:
    pass
try:
    remove("./office/config.xml")
except:
    pass

print("API Status:", checkApi(), "\nPyCore Office tweak\n可用语言:zh 简体中文 en English")
lang: object
if input("语言(Language): ").lower() == "zh":
    lang = Chinese
else:
    lang = English
getAPI()
api = dumpApi()
if not path.exists("./office/setup.exe"):
    downloadODP(api['ODP_EN'])
    move("./office.zip","./office/office.zip")
    with ZipFile("./office/office.zip", 'r') as zip_ref:
        zip_ref.extractall("./office")
        zip_ref.close()
clist = getConfigs(p=False)
getConfigs(p=True)
index = int(input("可用配置列表(输入数字): ")) - 1

try:
    fn = clist[index]
    copy(f"./configs/{fn}","./office/")
    rename(f"./office/{fn}","./office/config.xml")
except Exception as e:
    pass
if input("现在安装？(y or n)").lower() in ['y','yes','ok']:
    with open("./office/install.bat", 'w') as f:
        f.write("cd ./office\nsetup /download config.xml\nsetup /configure config.xml\npause")
        f.close()
    print(getcwd())
    system(fr'{getcwd()}\office\install.bat')
