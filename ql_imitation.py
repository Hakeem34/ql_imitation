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

RE_URL         = re.compile(r'https?://.+')
RE_ENVIRON_VAL = re.compile(r'(\%[^%]+\%)')

SHGFI_ICON = 0x000000100
SHGFI_ICONLOCATION = 0x000001000
SHGFI_USEFILEATTRIBUTES = 0x000000010
#SHIL_SIZE = 0x00001
SHIL_SIZE = 0x00002


#/*****************************************************************************/
#/* ICON画像データの取得と保存                                                */
#/*****************************************************************************/
def save_icon(target_file, target_icon):
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

    img.save(target_icon, format="ICO", sizes=[(32, 32)])
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
#       print(f'{file_type} の関連アプリケーションは {assoc_prog}です')
    else:
#       print(f'{file_type} の関連アプリケーションが見つかりませんでした')
        pass

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
#       print(f'{ext}に紐づくのは、{file_type}')
        assoc_prog = get_ftype_exe(file_type)
    else:
        print(f'{returncode.stdout}')

    return assoc_prog



#/*****************************************************************************/
#/* 実行ファイルの元になるpyファイルの作成                                    */
#/*****************************************************************************/
def make_script(target_script, start_target):
    py_file = open(target_script, "w", encoding='utf-8')
    print(f'import subprocess', file=py_file)

    cmd_text = f"cmd = r'start \"\" \"{start_target}\"'"            #/* start と起動したいファイルの間に、""を入れておかないと、半角スペースを含むパスのファイルを開けない */
    print(cmd_text, file=py_file)
    print('returncode = subprocess.Popen(cmd, shell=True)', file=py_file)
    py_file.close()


#/*****************************************************************************/
#/* 実行ファイルの作成                                                        */
#/*****************************************************************************/
def make_executable(target_file, exe_file):
    target_name = os.path.splitext(os.path.basename(target_file))[0]

    ql_path = '.\\ql_' + target_name
    os.makedirs(os.path.join(ql_path), exist_ok = True)
    os.chdir(ql_path)

    target_script = target_name + '.py'
    target_icon   = target_name + '.ico'

    make_script(target_script, target_file)
    img = save_icon(exe_file, target_icon)

    cmd_text = f'pyinstaller "{target_script}" --onefile --noconsole --clean --icon="{target_icon}"'
    returncode = subprocess.run(cmd_text, shell=True, capture_output=True, text=True)
#   print(returncode)
    os.chdir('..')
    print('EXE生成が完了しました。')

    #/* 生成したEXEファイルの絶対パスを返す */
    return Path(ql_path + '\\dist').resolve()



#/*****************************************************************************/
#/* 環境変数の置き換え                                                        */
#/*****************************************************************************/
def convert_environment_values(value):
    ret_value = value
    while (result := RE_ENVIRON_VAL.search(ret_value)):
        env_val      = result.group(1)
        env_val_name = env_val[1:-1]
        ret_value = ret_value.replace(env_val, os.environ[env_val_name])

    return ret_value


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
    open_path   = ''
    target_file = sys.argv.pop(0)
    target_abs_path = Path(target_file).resolve()
    target_name = os.path.splitext(os.path.basename(target_file))[0]
    target_ext  = os.path.splitext(os.path.basename(target_file))[1]
    if (target_ext):
        prog_path = get_assoc_exe(target_ext)

    if (os.path.isfile(target_file)):
        print(f'ファイル {target_file} を開くEXEを作成します')
        open_path = make_executable(target_abs_path, target_file)                  #/* 通常ファイルは絶対パスにして、ファイルそのものをexeファイルに指定する                        */
    elif  (os.path.isdir(target_file)):
        print(f'フォルダ {target_file} を開くEXEを作成します')
        sys_root = os.environ['SystemRoot']

        explorer = get_ftype_exe('folder')
        explorer = convert_environment_values(explorer)
        open_path = make_executable(target_abs_path, target_file)                  #/* フォルダは絶対パスにして、エクスプローラのexeを指定する                                      */
    elif  (result := RE_URL.match(target_file)):
        print(f'URL {target_file} を開くEXEを作成します')
        chrome = get_ftype_exe('ChromeHTML')
        edge   = get_ftype_exe('html')
        if (chrome != ''):
            open_path = make_executable(target_file, chrome)                       #/* URLはそのまま渡して、ChromeかEdgeのexeを指定する(デフォルトブラウザを調べるのはめんどくさい) */
        else:
            open_path = make_executable(target_file, edge)
    else:
        print(f'{target_file} は有効なファイルではありません')


    #/* 最後にできあがったexeファイルのあるパスを開く（タスクバーへの登録は手動でやってもらう） */
    if (open_path != ''):
        cmd_text = f'start \"\" \"{open_path}\"'
        returncode = subprocess.run(cmd_text, shell=True, capture_output=True, text=True)
        
    return


if __name__ == "__main__":
    main()


