import os
import re
import sys
import subprocess
from pathlib import Path
from win32com.shell import shell, shellcon  # pycharm cannot find those import but it will be interpreted

from PIL import Image
import win32api
import win32con
import win32gui
import win32ui

EXE_SCRIPT = 'import subprocess'

SHGFI_ICON = 0x000000100
SHGFI_ICONLOCATION = 0x000001000
SHGFI_USEFILEATTRIBUTES = 0x000000010
#SHIL_SIZE = 0x00001
SHIL_SIZE = 0x00002

def get_icon(target_file):
    ret, info = shell.SHGetFileInfo(target_file, 0, SHGFI_ICONLOCATION | SHGFI_ICON | SHIL_SIZE | SHGFI_USEFILEATTRIBUTES)
    hIcon, iIcon, dwAttr, name, typeName = info
    ico_x = win32api.GetSystemMetrics(win32con.SM_CXICON)
    hdc = win32ui.CreateDCFromHandle(win32gui.GetDC(0))
    hbmp = win32ui.CreateBitmap()
    hbmp.CreateCompatibleBitmap(hdc, ico_x, ico_x)
    hdc = hdc.CreateCompatibleDC()
    hdc.SelectObject(hbmp)
    hdc.DrawIcon((0, 0), hIcon)
    win32gui.DestroyIcon(hIcon)

    bmpinfo = hbmp.GetInfo()
    bmpstr = hbmp.GetBitmapBits(True)

    img = Image.frombuffer(
        "RGBA",
        (bmpinfo["bmWidth"], bmpinfo["bmHeight"]),
        bmpstr, "raw", "BGRA", 0, 1
    )
    return img



#/*****************************************************************************/
#/* ファイルタイプから、関連づいているアプリケーションを返す                  */
#/*****************************************************************************/
def get_ftype_exe(file_type):
    assoc_prog = ''
    cmd_text = f'ftype {file_type}'
    returncode = subprocess.run(cmd_text, shell=True, capture_output=True, text=True)

    if (result := re.match(fr'{file_type}=(\")*(.+\.exe)(\")*', returncode.stdout)):
        assoc_prog = result.group(2)
        print(f'{file_type} の関連アプリケーションは {assoc_prog}です')
    else:
        print(f'{file_type} の関連アプリケーションが見つかりませんでした')

    return assoc_prog


#/*****************************************************************************/
#/* 拡張子からファイルタイプを取得し、関連づいているアプリケーションを返す    */
#/*****************************************************************************/
def get_assoc_exe(ext):
    assoc_prog = ''
    cmd_text = f'assoc {ext}'
    returncode = subprocess.run(cmd_text, shell=True, capture_output=True, text=True)
#   print(f'{returncode.stdout}')

    if (result := re.match(f'{ext}=(.+)', returncode.stdout)):
        file_type = result.group(1)
        print(f'{ext}に紐づくのは、{file_type}')
        get_ftype_exe(file_type)
    else:
        print(f'{returncode.stdout}')

    return assoc_prog


#/*****************************************************************************/
#/* メイン関数                                                                */
#/*****************************************************************************/
def main():
    arg_0 = sys.argv.pop(0)

#   print(os.path.basename(arg_0))
#   print(os.path.basename(__file__))
    if (len(sys.argv) != 1):
        print(f'引数で対象ファイルを指定してください')
        return

#   print(sys.platform)
    target_file = sys.argv.pop(0)
    target_abs_path = Path(target_file).resolve()
    target_name = os.path.splitext(os.path.basename(target_file))[0]
    target_ext  = os.path.splitext(os.path.basename(target_file))[1]
    if (target_ext):
        prog_path = get_assoc_exe(target_ext)

    if (os.path.isfile(target_file) or os.path.isdir(target_file)):
        if (str(target_abs_path) == __file__):
            print(f'自分自身を指定することはできません')
            return

        print(f'{target_name} を起動するEXEを作成します')
        target_script = target_name + '.py'
        target_icon   = target_name + '.ico'
        py_file = open(target_script, "w", encoding='utf-8')
        print(f'import subprocess', file=py_file)

        cmd_text = f"cmd = r'start \"\" \"{target_abs_path}\"'"            #/* start と起動したいファイルの間に、""を入れておかないと、半角スペースを含むパスのファイルを開けない */
        print(cmd_text, file=py_file)
        print('returncode = subprocess.Popen(cmd, shell=True)', file=py_file)
        py_file.close()

        img = get_icon(target_file)
        img.save(target_icon, format="ICO", sizes=[(32, 32)])

        cmd_text = f'pyinstaller "{target_script}" --onefile --noconsole --clean --icon="{target_icon}"'
        returncode = subprocess.run(cmd_text, shell=True, capture_output=True, text=True)
        print(returncode)
        print('EXE生成が完了しました。')

    else:
        print(f'{target_file} は有効なファイルではありません')

    return


if __name__ == "__main__":
    main()


