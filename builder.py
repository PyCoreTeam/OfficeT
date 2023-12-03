from os import system


def buildOft(_mode: int):
    """0, --onefile
    1, --standalone"""
    with open("./build.bat", 'w') as f:
        f.write(
            "nuitka" + " --onefile --remove-output "
                       "--windows-icon-from-ico=./icon.ico main.py" if _mode == 0 else " --standalone" + "--remove"
                                                                                                         "-output "

                                                                                                         "--windows"
                                                                                                         "-icon-from"
                                                                                                         "-ico=./icon"
                                                                                                         ".ico main.py")
        f.close()
    try:
        system(
            "nuitka" + " --onefile --remove-output "
                       "--windows-icon-from-ico=./icon.ico main.py" if _mode == 0 else " --standalone" + "--remove"
                                                                                                         "-output "

                                                                                                         "--windows"
                                                                                                         "-icon-from"
                                                                                                         "-ico=./icon"
                                                                                                         ".ico main.py")
    except Exception as e:
        print(e)


mode = int(input("0, --onefile\n1, --standalone\n"))
if input("Build now?(y or n)").lower() in ['y', 'yes']:
    buildOft(mode)
else:
    exit(0)
