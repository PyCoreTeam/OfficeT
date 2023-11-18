from utils import *
from lang import *
from os import *
from shutil import move
from zipfile import *
try:
    mkdir("./apis")
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
downloadODP(api['ODP_EN'])
move("./office.zip","./office/office.zip")
with ZipFile("./office/office.zip", 'r') as zip_ref:
    zip_ref.extractall("./office")
