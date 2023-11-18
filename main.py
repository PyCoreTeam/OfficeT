from utils import *
from lang import *
from os import *
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
path = str(input(lang().ASK_PATH))
getAPI()
