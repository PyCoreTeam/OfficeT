from os import path, rename
from shutil import move, copy

from utils import *

clist = getIstOfficeConfig(p=False)
for step in range(3):
    initOft(step)
m = int(input("1.安装Office\n2.卸载Office\n"))
match m:
    case 1:
        print("API状态" + checkApi())
        getAPI()
        api = dumpApi()
        if not path.exists("./office/setup.exe"):
            downloadODP(api['ODP_EN'])
            move("./office.zip", "./office/office.zip")
            with ZipFile("./office/office.zip", 'r') as zip_ref:
                zip_ref.extractall("./office")
                zip_ref.close()

        getIstOfficeConfig(p=True)
        index = int(input("可用配置列表(输入数字): ")) - 1

        try:
            fn = clist[index]
            copy(f"./configs/{fn}", "./office/")
            rename(f"./office/{fn}", "./office/config.xml")
        except Exception as e:
            pass
        if input("现在安装？(y or n)").lower() in ['y', 'yes', 'ok']:
            with open("./office/install.bat", 'w') as f:
                f.write("cd ./office\nsetup /download config.xml\nsetup /configure config.xml\npause")
                f.close()
            print(getcwd())
            system(fr'{getcwd()}\office\install.bat')
            del f
    case 2:
        print("Office UnInstall\n卸载需谨慎\n")
        if not path.exists("./uninstall/o15-ctrremove.diagcab"):
            if input("你还未下载Uninstall组件。是否下载？(y or n)").lower() in ['ok', 'y', 'yes']:
                downloadUninstalls(clist['uninstall'])
        index = int(
            input("选择卸载工具\n1.Easy Fix (office2003, 2007, 2010，暂不支持win7及更低版本)\n2.o15ctrremove.diagcab(兼容，全版本)\n3.SaRa("
                  "兼容，全版本，建议使用)\n4.Troubleshoot(全版本)"))
        operationUninstaller(index)
        del index
    case _:
        print("未知命令。")
