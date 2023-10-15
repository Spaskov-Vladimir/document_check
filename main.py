import sys

from PyQt5 import uic
from PyQt5.QtWidgets import QMainWindow, QApplication, QFileDialog
import enchant
import pandas as pd
from PyPDF2 import PdfReader
import camelot
import re

from datetime import datetime

dictionary = enchant.Dict("ru_Ru")


def is_correct_word(word):
    return dictionary.check(word)


def correct_word(word):
    return ", ".join(dictionary.suggest(word))


months = {
    "января": ".01.",
    "февраля": ".02.",
    "марта": ".03.",
    "апреля": ".04.",
    "мая": ".05.",
    "июня": ".06.",
    "июля": ".07.",
    "августа": ".08.",
    "сентября": ".09.",
    "октября": ".10.",
    "ноября": ".11.",
    "декабря": ".12."
}


def validate_date(date_text):
    date_text = re.sub(r'\s', "", date_text.replace("\"", ""))
    for month in months:
        date_text = date_text.replace(month, months[month])
    try:
        if len(date_text.split(".")[-1]) == 4:
            datetime.strptime(date_text, '%d.%m.%Y')
        else:
            datetime.strptime(date_text, '%d.%m.%y')
        return True
    except ValueError:
        return False


class MyProject(QMainWindow):
    def __init__(self):
        super().__init__()
        self.reader = None
        self.path_to_file = ""
        self.data = ""
        uic.loadUi('form.ui', self)
        self.btn_check_doc.clicked.connect(self.load_data)
        self.btn_correct_doc.clicked.connect(self.correct_text)

    def clear_text(self):
        self.txt_doc.setText("")

    def add_text(self, text):
        old_text = self.txt_doc.toPlainText()
        self.txt_doc.setText(old_text + "\n" + text if old_text != "" else text)

    def correct_text(self):
        self.clear_text()
        self.path_to_file = QFileDialog.getOpenFileName(self,
                                                        "Выбрать документ",
                                                        "",
                                                        "Файлы документов (*.pdf)")[0]
        self.reader = PdfReader(self.path_to_file)
        self.data = ""
        for page in self.reader.pages:
            self.data += page.extract_text()
        for date in re.findall(r'\b(\d{2}\s*\.\s*\d{2}\s*\.\s*\d{2,4})|(["]\s*\d{1,2}\s*["]\s*(?:января|февраля|марта|апреля|мая|июня|июля|августа|сентября|октября|ноября|декабря)\s*\d{4})\b', self.data):
            if not validate_date("".join(date)):
                self.add_text("Дата " + "".join(date) + " не корректна.")
        for word in re.findall(r"\b(?![А-ЯЁ]{2,}\b)[а-яёА-ЯЁ]{2,}\b", self.data.replace("\n", " ")):
            if not is_correct_word(word):
                self.add_text("Ошибка в слове \""
                              + word + "\"."
                              + " Возможные исправдения: \""
                              + correct_word(word).replace(", ", "\", \"")
                              + "\".")

    def load_data(self):
        self.path_to_file = QFileDialog.getOpenFileName(self,
                                                        "Выбрать документ",
                                                        "",
                                                        "Файлы документов (*.pdf)")[0]
        self.reader = PdfReader(self.path_to_file)
        self.data = ""
        for page in self.reader.pages:
            self.data += page.extract_text()
        extract_data = self.data.split("\n")[0]
        self.clear_text()
        for date in re.findall(r'\b(\d{2}\s*\.\s*\d{2}\s*\.\s*\d{2,4})|(["]\s*\d{1,2}\s*["]\s*(?:января|февраля|марта|апреля|мая|июня|июля|августа|сентября|октября|ноября|декабря)\s*\d{4})\b', self.data):
            if not validate_date("".join(date)):
                self.add_text("Ошибка: дата некорректна")
                break
        if "М-11" in extract_data.replace(" ", ""):
            self.m_11()
        elif "ФМУ-76" in extract_data.replace(" ", ""):
            self.fmu_76()

    def organization_name_find(self, split_data):
        first_pos = split_data[6].find("Организация") + 10
        second_pos = split_data[6].find("по ОКПО")
        if (second_pos - first_pos) < 5:
            self.add_text("Ошибка: имя организации не заполнено")
        # ОКПО
        if len(split_data[6]) - (split_data[6].find("ОКПО") + 3) < 3:
            self.add_text("Ошибка: ОКПО не заполнен")

    def structural_unit(self, data, table_0):
        if table_0[0][3] == "":
            self.add_text("Ошибка! БЕ не заполнено")

        line_data = data.replace("\n", " ")
        if len(line_data.split("подразделение")[1].split("БЕ")[0].replace(" ", "")) == 0:
            self.add_text("Ошибка! Структурное подразделение не заполнено")

    def m_11(self):
        split_data = self.data.split("\n")
        tables = camelot.read_pdf(self.path_to_file)

        table = tables[0]
        table_0 = table.df

        table_1 = tables[1]
        table_1 = table_1.df.applymap(lambda x: x.replace('\n', ' '))

        table = tables[2]
        table_2 = table.df.applymap(lambda x: x.replace('\n', ' '))

        if len(split_data[3].split(" ")) != 3:
            self.add_text("Ошибка: номер документа не заполнен!")

        # Организация
        self.organization_name_find(split_data)

        # Структурное подразделение
        self.structural_unit(self.data, table_0)

        # Дата составления
        if table_1[0][2] == "":
            self.add_text("Ошибка! Дата не заполнена!")

        # Отправитель
        if table_1[2][2] == "":
            self.add_text("Ошибка! Отправитель не заполнен")

        if table_1[4][2] == "":
            self.add_text("Ошибка! Получатель не заполнен")

        if len(self.data.split("Затребовал")[1].split("Разрешил")[0].replace(" ", "")) == 0:
            self.add_text("Ошибка! Поле \"Затребовал\" не заполнено!")

        if len(self.data.split("Разрешил")[1].split("Материальные")[0].replace(" ", "")) == 0:
            self.add_text('Ошибка! Поле "Разрешил" не заполнено!')

        if len(self.data.split("Через кого")[1].split("Затребовал")[0].replace(" ", "")) == 0:
            self.add_text('Ошибка! Поле "Через кого" не заполнено!')

        for i in range(3, table_2.shape[0]):
            if table_2[1][i] == "" or table_2[2][i] == "":
                self.add_text("Ошибка! Наименование или номенклатурный номер не заполнены!")
                break

        for i in range(3, table_2.shape[0]):
            if table_2[7][i] == "" or table_2[8][i] == "":
                self.add_text("Ошибка! Единицы измерения не заполнены!")
                break

        for i in range(3, table_2.shape[0]):
            if table_2[9][i] == "" or table_2[10][i] == "":
                self.add_text('Ошибка! Поля "Затребовано" или "Отпущено" не заполнено!')
                break

        if len(self.data.split("Отпустил")[1].split("Получил")[0].replace("\n", " ")
                       .replace("(должность)электронная подпись (подпись)", "")
                       .replace("(расшифровка подписи)", "")
                       .replace(" ", "")) < 5:
            self.add_text('Ошибка! Поле "Отпустил" или "Получил" не заполнено')

    def fmu_76(self):
        line_data = self.data.replace("\n", " ")
        split_data = self.data.split("\n")
        if len(" ".join(line_data.split("по ОКУД")[1].split("по ОКПО")[0].split(" ")[2:])) < 3:
            self.add_text("Ошибка! Название организации не заполнено!")
        if len(line_data.split("организация")[1].split("БЕ")[0].replace(" ", "")) < 5:
            self.add_text("Ошибка! Название структурного подразделения не заполнено!")

        numbers = list(range(1, len(self.reader.pages) + 1))
        all_pages = ','.join(map(str, numbers))
        tables = camelot.read_pdf(self.path_to_file, pages=all_pages)

        if len(line_data.split("ответственное лицо")[1].split("Направление расхода")[0].replace(" ", "")) < 5:
            self.add_text("Ошибка! Материально ответственное лицо не заполнено!")

        if len(line_data.split("Направление расхода")[1].split("Инвентарный номер")[0].replace(" ", "")) < 5:
            self.add_text("Ошибка! Направление расхода не заполнено")

        if len(line_data.split("Комиссия в составе:")[1].split("составила настоящий")[0].replace(" ", "")) < 5:
            self.add_text("Ошибка! Состав комиссии не заполнен")

        list_of_tables = []
        void_flag = True

        for element in tables:
            list_of_tables.append(element.df)

        #         for i,j in element.df.shape:
        #             self.add_text(f"--{element[i][j]}--")
        #         if not element.df.isna().all().all():

        #     for index, row in list_of_tables2[-1].iterrows():
        #         for column in list_of_tables2[-1].columns:
        #             value = row[column]
        #             self.add_text(f"Значение в строке {index}, столбце '{column}': {value}")

        if len(list_of_tables) > 2:
            tables_all = pd.concat(list_of_tables[1:-1], axis=0, ignore_index=True)
        else:
            tables_all = list_of_tables[1]
        if len(list_of_tables[0][0][2].replace(" ", "")) < 5:
            self.add_text("Ошибка! Цех структурного подразделения не заполнен!")

        for i in range(0, tables_all.shape[0]):
            if len(tables_all[4][i]) < 5 or len(tables_all[5][i]) < 5:
                self.add_text("Ошибка! Материальные ценности не заполнены!")
                break

        for i in range(0, tables_all.shape[0]):
            if not tables_all[7][i].isdigit() or len(tables_all[8][i]) < 2:
                self.add_text("Ошибка! Еденицы измерения не заполнены!")
                break

        for i in range(0, tables_all.shape[0]):
            if len(tables_all[9][i]) < 3:
                self.add_text("Ошибка! Нормативное количество не заполнено!")
                break

        for i in range(0, tables_all.shape[0]):
            if len(tables_all[10][i]) < 3:
                self.add_text("Ошибка! Поле 'Фактически израсходованно' не заполнено!")
                break

        for i in range(0, tables_all.shape[0]):
            if len(tables_all[13][i]) < 3:
                self.add_text("Ошибка! Поле 'Отклонение фактического расхода от нормы' не заполнено!")
                break

        for i in range(0, tables_all.shape[0]):
            if len(tables_all[14][i]) < 3:
                self.add_text("Ошибка! Поле 'Вид работ' не заполнено!")
                break

        for i in range(0, tables_all.shape[0]):
            if len(tables_all[10][i]) < 3:
                self.add_text("Ошибка! Поле 'Фактически израсходованно' не заполнены!")
                break

        if len(list_of_tables[0][2][2]) < 3:
            self.add_text("Ошибка! Статья расходов/носитель затрат заполнен неправильно")

        for i in range(0, tables_all.shape[0]):
            if len(tables_all[0][i]) < 10:
                self.add_text("Ошибка! Технический счет не заполнен!")
                break

        for i in range(0, tables_all.shape[0]):
            if len(tables_all[1][i]) < 12:
                self.add_text("Ошибка! Производственный заказ не заполнен!")
                break


if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = MyProject()
    ex.show()
    sys.exit(app.exec_())

