import sys

import requests
from json import *


def checkApi():
    response: requests.Response
    try:
        response = requests.get("https://raw.githubusercontent.com/PyCoreTeam/OfficeT/main/api.json",verify=False)
    except:
        pass
    if response.ok:
        return 'OK'
def getAPI():
    try:
        f = open("./apis/api.json",'w')
        f.write(str(requests.get("https://raw.githubusercontent.com/PyCoreTeam/OfficeT/main/api.json", verify=False).text))
    except Exception as e:
        print(e)
        sys.exit(0)
def dumpApi():
    f = open("./apis/api.json", 'r')
    return load(f)
def downloadODP(api):
    f = open("odp.exe",'wb')
    f.write(requests.get(api).content)
