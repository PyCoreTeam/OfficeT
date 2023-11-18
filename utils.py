import sys

import requests



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
