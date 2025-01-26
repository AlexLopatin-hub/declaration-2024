import os
import time
from bs4 import BeautifulSoup


def main(working_folder):
    import pywinauto as pw
    # Во избежание некорректной работы pywinauto
    try:
        pw.Application(backend="uia").connect(path="C:\\АО ГНИВЦ\\Декларация 2024\\Decl2024.exe")
        raise RuntimeError("Закройте приложение Декларация 2024 перед тем как запускать программу")
    except pw.application.ProcessNotFoundError:
        pass

    xml_folder_name = "xml"
    user = os.getlogin()

    # Создаём папку для всех извлечённых xml файлов
    os.mkdir(f"C:\\Users\\{user}\\Desktop\\{xml_folder_name}")

    # Поочерёдно запускаем каждый dc4 файл и извлекаем оттуда xml
    for root, dirs, files in os.walk(working_folder):
        for file in files:
            if file.endswith(".dc4"):
                open_file(os.path.join(root, file))
                create_xml(xml_folder_name)

    # Пробегаемся по xml-файлам, находим нужные и записываем данные клиентов в текстовый файл
    for root, dirs, files in os.walk(f"C:\\Users\\{user}\\Desktop\\{xml_folder_name}"):
        for file in files:
            data = collect_info(os.path.join(root, file))
            if data:
                with open(f"C:\\Users\\{user}\\Desktop\\{xml_folder_name}\\result.txt", "a") as f:
                    f.write(" ".join(data) + "\n")


def collect_info(file: str) -> list:
    data = []
    if file.endswith(".xml") and find_deduction(file):
        with open(file) as f:
            soup = BeautifulSoup(f, features='xml')
            fio = soup.find("ФИОФЛ")
            fio = list(fio.attrs.values())
            phone = soup.find("НПФЛ3")
            phone = [phone.attrs["Тлф"]]
            data = fio + phone
        print(f"<log> {file} -- 1")
    else:
        print(f"<log> {file} -- 0")
    return data


def open_file(file: str):
    import pywinauto as pw
    file = file.replace("%", "{%}").replace("^", "{^}").replace("+", "{+}")
    pw.Application().start('explorer.exe "C:\\Program Files"')
    app = pw.Application(backend="uia").connect(path="explorer.exe", title="Program Files")
    dlg = app["Program Files"]
    dlg['Address: C:\\Program Files'].wait("ready")
    time.sleep(0.1)
    dlg.type_keys("%D")
    dlg.type_keys(file, with_spaces=True)
    dlg.type_keys("{ENTER}")
    time.sleep(1)
    app.kill()


def create_xml(folder_path: str):
    import pywinauto as pw
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
        f = f.read()
        pos = f.find("ОстИВБезПроц")
        if pos > 0:
            ost = f[pos:pos + 70].split()
            try:
                ost = int(ost[0][14:-4]) + int(ost[1].split(">")[0][14:-4])
            except ValueError:
                print("<log> Ошибка чтения xml-файла")
                return False
            if ost:
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
