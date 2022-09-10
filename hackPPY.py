import re
import winreg
from rich import print

VERSION = re.compile(r'\d+\.\d+')


def get_all_ppt_ver() -> list:
    reg = winreg.ConnectRegistry(None, winreg.HKEY_CURRENT_USER)
    keys = winreg.OpenKey(
        reg,
        r'Software\Microsoft\Office',
        0,
        winreg.KEY_ALL_ACCESS)
    versions = []
    i = 0
    while 1:
        try:
            s = winreg.EnumKey(keys, i)
            if VERSION.match(s):
                versions.append(s)
        except:
            break
        finally:
            i += 1
    print('[green]Find PPT versions:', " ".join(versions))
    return versions


def hack_ppt(dpi=300, ppt_ver='16.0') -> bool:
    reg_dir = r'Software\Microsoft\Office\$\PowerPoint\Options'.replace(
        '$', ppt_ver)
    print(f'\n[bold yellow]------- Hacking PPT {ppt_ver} -------')
    print(f'path: [cyan]{reg_dir}')
    reg = winreg.ConnectRegistry(None, winreg.HKEY_CURRENT_USER)
    try:
        keys = winreg.OpenKey(
            reg,
            reg_dir,
            0,
            winreg.KEY_ALL_ACCESS)
    except Exception:
        print(f'status: [bold yellow]Pass[/], version {ppt_ver} not found!')
        return False
    try:
        winreg.QueryValueEx(keys, 'ExportBitmapResolution')
    except FileNotFoundError:
        print(f"status: [bold yellow]Pass[/], PPT version {ppt_ver} Already hacked!")
        return False

    winreg.CreateKey(keys, 'ExportBitmapResolution')
    winreg.SetValueEx(keys, 'ExportBitmapResolution', 0, winreg.REG_DWORD, dpi)
    winreg.FlushKey(keys)
    winreg.CloseKey(keys)
    reg.Close()
    print(f"status: [bold green]Success[/], PPT version {ppt_ver}")
    return True


if __name__ == '__main__':
    # if admin():
    print('[green]Input [bold]DPI [/][yellow](default 300): ', end='')
    dpi = input()
    if not dpi:
        dpi = 300
    else:
        dpi = int(dpi)
    for i in get_all_ppt_ver():
        hack_ppt(dpi=dpi, ppt_ver=i)
    print('\n[bold green]Press enter to exit...', end=' ')
    input()
