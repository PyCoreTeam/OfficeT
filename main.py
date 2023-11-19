from utils import *
from lang import *
from os import mkdir, remove, path, rename, system, getcwd
from shutil import move, copy
from zipfile import *
clist = getConfigs(p=False)
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
m = int(input("1.安装Office\n2.卸载Office\n"))
if m == 1:
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
elif m == 2:
    print("Office UnInstall\n卸载需谨慎\n")
    if not path.exists("./uninstall/o15-ctrremove.diagcab"):
        if input("你还未下载Uninstall组件。是否下载？(y or n)").lower() in ['ok','y','yes']:
            getUninstalls(clist['uninstall'])
    index = int(input("选择卸载工具\n1.Easy Fix (office2003, 2007, 2010，暂不支持win7及更低版本)\n2.o15ctrremove.diagcab(兼容，全版本)\n3.SaRa("
                      "兼容，全版本，建议使用)\n4.Troubleshoot(全版本)"))
    if index == 1:
        i = int(input("Office版本:\n1.2003\n2.2007\n3.2010"))
        if i == 1:
            system(fr"{getcwd()}\uninstall\EasyFix\MicrosoftEasyFix20054.mini.diagcab")
        elif i == 2:
            system(fr"{getcwd()}\uninstall\EasyFix\MicrosoftEasyFix20052.mini.diagcab")
        elif i == 3:
            system(fr"{getcwd()}\uninstall\EasyFix\MicrosoftEasyFix20055.mini.diagcab")
    elif index == 2:
        system(fr"{getcwd()}\uninstall\o15-ctrremove.diagcab")
    elif index == 3:
        system(fr"{getcwd()}\uninstall\ClickOnce\SaraSetup.exe")
    elif index == 4:
        system(fr"{getcwd()}\uninstall\Troubleshoot\OffScrubC2R.vbs")
