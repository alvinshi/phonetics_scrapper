import requests
import re
from openpyxl import Workbook
from openpyxl import load_workbook


class Scrapper:
    SUCCESS = 200
    WORD_PLACEHOLDER = "{word}"
    URL_TEMPLATE = "https://en.oxforddictionaries.com/definition/{}".format(WORD_PLACEHOLDER)
    OXFORD_REGEX = re.compile('<span class="phoneticspelling">(.*?)</span>')
    INPUT_PATH = "input.xlsx"
    OUTPUT_PATH = "output.xlsx"

    @staticmethod
    def _retrieve_html(url):
        response = requests.get(url)
        return response.status_code, response.content

    @staticmethod
    def _get_cell_index(row, col):
        return col.upper() + str(row)

    def _get_words(self):
        wb = load_workbook(filename= self.INPUT_PATH)
        print(wb)
        sheet = wb['Sheet1']
        row = 1
        words = []
        end_reached = False
        while not end_reached:
            cell = self._get_cell_index(row, "A")
            word = sheet[cell].value
            if word and len(word) > 0:
                words.append(word)
                row += 1
            else:
                end_reached = True
        return words

    def _get_phonetics(self, url):
        _, content = self._retrieve_html(url)
        m = self.OXFORD_REGEX.search(content)
        if m:
            return m.group(1)
        else:
            return ""

    def run(self):
        wb = Workbook()
        worksheet = wb.active
        words = self._get_words()
        urls = list(map(lambda word: self.URL_TEMPLATE.replace(self.WORD_PLACEHOLDER, word), words))
        for index, url in enumerate(urls):
            phonetics = self._get_phonetics(url)
            worksheet.cell(row=index+1, column=1, value=words[index])
            worksheet.cell(row=index+1, column=2, value=phonetics)
            if index % 50 == 0:
                print("{} done".format(index))
                wb.save(self.OUTPUT_PATH)
        wb.save(self.OUTPUT_PATH)
        print("All {} words have been matched with their oxford phonetics".format(len(words)))


def main():
    Scrapper().run()


if __name__ == "__main__":
    main()