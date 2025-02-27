import os
import time
from bs4 import BeautifulSoup


def main(working_folder: str, alternate = 0) -> str:
    user = os.getlogin()

    if alternate == 1:
        # if os.path.isdir(working_folder):
        #     raise FileNotFoundError("Не найдена указанная папка")
        for root, dirs, files in os.walk(working_folder):
            for file in files:
                data = collect_info(os.path.join(root, file))
                if data:
                    with open(f"{working_folder}\\result.txt", "a") as f:
                        f.write(" ".join(data) + "\n")
        return working_folder

    else:
        import pywinauto as pw
        # Во избежание некорректной работы pywinauto
        try:
            pw.Application(backend="uia").connect(path="C:\\АО ГНИВЦ\\Декларация 2024\\Decl2024.exe")
            raise RuntimeError("Закройте приложение Декларация 2024 перед тем как запускать программу")
        except pw.application.ProcessNotFoundError:
            pass

        # Создаём папку для всех извлечённых xml файлов
        for i in range(100):
            xml_folder_name = "xml" + str(i).replace("0", "")
            xml_folder_path = f"C:\\{xml_folder_name}"
            if not os.path.isdir(xml_folder_path):
                break
        else:
            raise FileExistsError
        # os.mkdir(f"C:\\Users\\{user}\\Desktop\\{xml_folder_name}")
        os.mkdir(xml_folder_path)

        # Поочерёдно запускаем каждый dc4 файл и извлекаем оттуда xml
        for root, dirs, files in os.walk(working_folder):
            for file in files:
                if file.endswith(".dc4"):
                    open_file(os.path.join(root, file))
                    create_xml(xml_folder_name)

        # Пробегаемся по xml-файлам, находим нужные и записываем данные клиентов в текстовый файл
        for root, dirs, files in os.walk(xml_folder_path):
            for file in files:
                data = collect_info(os.path.join(root, file))
                if data:
                    with open(f"{xml_folder_path}\\result.txt", "a") as f:
                        f.write(" ".join(data) + "\n")

        # Возвращаем путь созданной папки
        # return f"C:\\Users\\{user}\\Desktop\\xml"
        return xml_folder_path


def open_file(file: str) -> None:
    import pywinauto as pw
    pw.Application().start(f'explorer.exe "{file}"')

def create_xml(folder_name: str) -> None:
    import pywinauto as pw
    app = pw.Application(backend="uia").connect(path="C:\\АО ГНИВЦ\\Декларация 2024\\Decl2024.exe")
    dlg_spec = app.window(title_re=".* - Декларация 2024$")
    time.sleep(2)
    actionable_dlg = dlg_spec.wait('visible')
    dlg_spec.children()[2].click_input()
    # dlg_spec.CoolBarMenu.window(title="Декларация").click_input()
    # pw.keyboard.send_keys("{DOWN 2}{ENTER}")
    expl = pw.Desktop(backend="uia")["Обзор папок"]
    dlg_expl = expl
    dlg_expl_save_folder = dlg_expl.window(class_name="SysTreeView32")
    dlg_expl_save_folder.window(title="Этот компьютер").window(title_re=".*(C:)").double_click_input()

    dlg_expl_save_folder.wheel_mouse_input(wheel_dist = -3)
    try:
        dlg_expl_save_folder.window(title=folder_name).click_input()
    except pw.findwindows.ElementNotFoundError:
        dlg_expl_save_folder.wheel_mouse_input(wheel_dist = -3)
        try:
            dlg_expl_save_folder.window(title=folder_name).click_input()
        except pw.findwindows.ElementNotFoundError:
            dlg_expl_save_folder.wheel_mouse_input(wheel_dist = -3)
            dlg_expl_save_folder.window(title=folder_name).click_input()

    dlg_expl.window(title="ОК").click_input()
    pw.Desktop(backend="uia")["Декларация 2024"].window(title="ОК").click_input()
    app.kill()

def collect_info(file: str) -> list:
    data = []
    if file.endswith(".xml") and find_deduction(file):
        with open(file) as f:
            soup = BeautifulSoup(f, features='xml')
            fio = soup.find("ФИОФЛ")
            fio = list(fio.attrs.values())
            fio[-1] += ","
            phone = soup.find("НПФЛ3")
            phone = [phone.attrs["Тлф"]]
            data = fio + phone
        print(f"<log> {file} -- 1")
    else:
        print(f"<log> {file} -- 0")
    return data

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
    main(".")
    t2 = time.time()
    print(f"time consumed: {t2 - t1}")
