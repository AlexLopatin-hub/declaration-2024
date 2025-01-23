import os
from time import time
from bs4 import BeautifulSoup
from lxml.objectify import parse


def main():
    data = []
    for root, dirs, files in os.walk("D:\ржака"):
        for file in files:
            if file.endswith(".xml") and find_deduction(os.path.join(root, file)):
                data.append(get_info(os.path.join(root, file)))
                print(f"{root} {file} -- 1")
            # else:
                # print(f"{file} -- 0")
    print(data)


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
        phone = phone.attrs["Тлф"]
        fio.append(phone)
        return fio                                      # возвращает всю инфу


if __name__ == '__main__':
    t1 = time()
    main()
    t2 = time()
    print(t2 - t1)
