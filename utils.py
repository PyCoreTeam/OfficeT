import sys
from os import listdir, mkdir
from zipfile import ZipFile

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
    from tqdm import tqdm
    r = requests.get(api,verify=False)
    total = int(r.headers.get('content-length'))
    progress_bar = tqdm(total=total, desc="ODP - ")
    for chunk in r.iter_content(chunk_size=8192):
        progress_bar.update(len(chunk))
    f = open("./office.zip",'wb')
    f.write(r.content)
    progress_bar.close()
def getConfigs(p=False):

    if p == False:

        return listdir("./configs")
    else:
        v = 0
        for i in listdir("./configs"):
            v+=1
            print(str(v),'.', i.replace(".xml", ''))
        return listdir("./configs")
def getUninstalls(api):
    r = requests.get(api, verify=False)
    try:
        mkdir("uninstall")
    except:
        pass
    f = open("./uninstall/u.zip",'wb')
    f.write(r.content)
    f.close()
    with ZipFile("./uninstall/u.zip",'r') as f:
        f.extractall("./uninstall")
        f.close()



