import os
import time
from bs4 import BeautifulSoup
import pywinauto as pw


def main():
    xml_folder_name = "xml"
    user = "кто"
    working_folder = "D:\ржака"

    # Создаём папку для всех извлечённых xml файлов
    os.mkdir(f"C:\\Users\\{user}\\Desktop\\{xml_folder_name}")
    # Поочерёдно запускаем каждый dc4 файл и извлекаем оттуда xml
    for root, dirs, files in os.walk(working_folder):
        for file in files:
            if file.endswith(".dc4"):
                open_file(os.path.join(root, file))
                create_xml(xml_folder_name)

    data = collect_info(xml_folder_name, user)
    print(f"collected data: {data}")


def collect_info(path: str, user: str):
    data = []
    for root, dirs, files in os.walk(f"C:\\Users\\{user}\\Desktop\\{path}"):
        for file in files:
            if file.endswith(".xml") and find_deduction(os.path.join(root, file)):
                data.append(get_info(os.path.join(root, file)))
                print(f"<log> {root}\\{file} -- 1")
            # else:
                # print(f"<log> {file} -- 0")
    return data


def open_file(file: str):
    file = file.replace("%", "{%}").replace("^", "{^}").replace("+", "{+}")
    pw.Application().start('explorer.exe "C:\\Program Files"')
    app = pw.Application(backend="uia").connect(path="explorer.exe", title="Program Files")
    dlg = app["Program Files"]
    dlg['Address: C:\\Program Files'].wait("ready")
    dlg.type_keys("%D")
    dlg.type_keys(file, with_spaces=True)
    dlg.type_keys("{ENTER}")
    time.sleep(1)
    app.kill()


def create_xml(folder_path: str):
    app = pw.Application(backend="uia").connect(path="C:\\АО ГНИВЦ\\Декларация 2024\\Decl2024.exe")
    dlg_spec = app.window(title_re=".* - Декларация 2024$")
    actionable_dlg = dlg_spec.wait('visible')
    dlg_spec.CoolBarButtons.click_input()
    expl = pw.Desktop(backend="uia")["Обзор папок"]
    dlg_expl = expl
    dlg_expl_save_folder = dlg_expl.window(class_name="SysTreeView32").window(title="Рабочий стол")
    dlg_expl_save_folder.double_click_input()
    dlg_expl_save_folder.window(title=folder_path).click_input()
    dlg_expl.window(title="ОК").click_input()
    pw.Desktop(backend="uia")["Декларация 2024"].window(title="ОК").click_input()
    app.kill()


def find_deduction(path: str) -> bool:
    with open(path) as f:
        p = f.read()
        pos = p.find("ОстИВБезПроц")
        if pos > 0:
            p = p[pos:pos + 30].split()[0]
            p = int(p[14:-4])
            if p:
                return True
        return False


def get_info(path: str) -> list:
    with open(path) as f:
        soup = BeautifulSoup(f, features='xml')
        fio = soup.find("ФИОФЛ")
        fio = list(fio.attrs.values())
        phone = soup.find("НПФЛ3")
        phone = [phone.attrs["Тлф"]]
        return fio + phone


if __name__ == '__main__':
    t1 = time.time()
    main()
    t2 = time.time()
    print(f"time consumed: {t2 - t1}")
