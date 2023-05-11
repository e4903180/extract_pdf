from tqdm import tqdm
import fitz
from ExtractPdf_0503 import *
import os

class Sinopac(Adviser):
    '''Handle 永豐投顧(SinoPac) pdf

        Args :
            directory_path : (str) pdf path
        
        Return :
            rating : (str) recommend
    '''
    def __init__(self, file_path):
        super().__init__(file_path)
        self.file_path = file_path
        self.possible_rating = ['買進', '中立']

    def get_rating_process(self):
        with fitz.open(self.file_path) as doc:
            page = doc.load_page(0)
            rect = page.rect
            if self.check_source(doc)=='old':
                self.advisor = 'Sinopac'
                self.get_rating_old_version_1(page)
                self.get_rating_old_version_2(page)
            elif self.check_source(doc)=='new':
                self.advisor = 'Sinopac'
                clip_check_report = fitz.Rect(0, 0, rect.width, 150)
                text_check_report = page.get_text(clip=clip_check_report, sort=True).strip()
                if self.check_report(text_check_report):
                    self.get_rating_new_version_1(page)
                    self.get_rating_new_version_2(page)
        return self.check_rating(self.rating_1, self.rating_2, self.possible_rating)

    def get_target_price_process(self):
        with fitz.open(self.file_path) as doc:
            page = doc.load_page(0)
            rect = page.rect
            if self.check_source(doc)=='old':
                self.advisor = 'Sinopac'
                self.get_target_price_old_version_1(page)
                self.get_target_price_old_version_2(page)
            elif self.check_source(doc)=='new':
                self.advisor = 'Sinopac'
                clip_check_report = fitz.Rect(0, 0, rect.width, 150)
                text_check_report = page.get_text(clip=clip_check_report, sort=True).strip()
                if self.check_report(text_check_report):
                    self.get_target_price_new_version_1(page)
                    self.get_target_price_new_version_2(page)
        return self.check_targrt_price(self.tp_1, self.tp_2)
    
    def get_author_process(self):
        with fitz.open(self.file_path) as doc:
            page = doc.load_page(0)
            rect = page.rect
            if self.check_source(doc)=='old':
                self.advisor = 'Sinopac'
                self.get_author_old_version_1(page, rect)
                self.get_author_old_version_2(page)
            elif self.check_source(doc)=='new':
                self.advisor = 'Sinopac'
                clip_check_report = fitz.Rect(0, 0, rect.width, 150)
                text_check_report = page.get_text(clip=clip_check_report, sort=True).strip()
                if self.check_report(text_check_report):
                    self.get_author_new_version_1(page, rect)
                    self.get_author_new_version_2(page)
        return self.check_author(self.author_1, self.author_2)

    def get_summary_process(self):
        with fitz.open(self.file_path) as doc:
            page = doc.load_page(0)
            rect = page.rect
            if self.check_source(doc)=='old':
                self.advisor = 'Sinopac'
                self.get_summary_old_version_1(page, rect)
                self.get_summary_old_version_2(page)
            elif self.check_source(doc)=='new':
                self.advisor = 'Sinopac'
                clip_check_report = fitz.Rect(0, 0, rect.width, 150)
                text_check_report = page.get_text(clip=clip_check_report, sort=True).strip()
                if self.check_report(text_check_report):
                    self.get_summary_new_version_1(page, rect)
                    self.get_summary_new_version_2(page)
        return self.check_summary(self.summary_1, self.summary_2)


    def check_source(self, doc):
        page_check_source = doc.load_page(-1)
        text_check_source = page_check_source.get_text()
        if '永豐證券投資顧問股份有限公司' in text_check_source:
            return 'old'
        elif 'SinoPac Securities' in text_check_source:
            return 'new'

    def check_report(self, text_check_report):
        if '個股聚焦' in text_check_report:
            return True
    
    def get_rating_old_version_1(self, page):
        clip_old_version_1 = fitz.Rect(220, 80, 560, 140)
        text_old_version_1 = page.get_text(clip=clip_old_version_1, sort=True).strip()
        try:
            text_old_version_1 = text_old_version_1.split('）')[1].strip()
            self.rating_1 = text_old_version_1.split('\n')[1].strip()
        except:
            self.rating_1 = 'NULL'

    def get_rating_old_version_2(self, page):
        clip_old_version_2 = fitz.Rect(425, 90, 560, 130)
        self.rating_2 = page.get_text(clip=clip_old_version_2, sort=True).strip()
    
    def get_rating_new_version_1(self, page):
        clip_new_version_1 = fitz.Rect(0, 0, 200, 400)
        text_new_version_1 = page.get_text(clip=clip_new_version_1, sort=True).strip()
        try:
            text_new_version_1 = text_new_version_1.split('投資建議')[1]
            self.rating_1 = text_new_version_1.split('\n')[0].strip()
        except:
            self.rating_1 = 'NULL'
    
    def get_rating_new_version_2(self, page):
        clip_new_version_2 = fitz.Rect(75, 200, 120, 235)
        text_new_version_2 = page.get_text(clip=clip_new_version_2, sort=True).strip()
        self.rating_2 = text_new_version_2  

    def get_target_price_old_version_1(self, page):
        clip_old_version_1 = fitz.Rect(130, 100, 200, 160)
        text_old_version_1 = page.get_text(clip=clip_old_version_1, sort=True).strip()
        try:
            text_old_version_1 = text_old_version_1.split('NT$')[1].strip()
            self.tp_1 = text_old_version_1.split('\n')[0].strip()
        except:
            self.tp_1 = 'NULL'
    
    def get_target_price_old_version_2(self, page):
        clip_old_version_2 = fitz.Rect(130, 140, 200, 160)
        text_old_version_2 = page.get_text(clip=clip_old_version_2).strip()
        try:
            text_old_version_2 = text_old_version_2.split('NT$')[1].strip()
            self.tp_2 = text_old_version_2.split('\n')[0].strip()
        except:
            self.tp_2 = 'NULL'
    
    def get_target_price_new_version_1(self, page):
        clip_new_version_1 = fitz.Rect(115, 0, 200, 300)
        text_new_version_1 = page.get_text(clip=clip_new_version_1, sort=True).strip()
        try:
            text_new_version_1 = text_new_version_1.split('NT$')[1].strip()
            self.tp_1 = text_new_version_1.split('\n')[0].strip()
        except:
            self.tp_1 = 'NULL'
    
    def get_target_price_new_version_2(self, page):
        clip_new_version_2 = fitz.Rect(115, 235, 200, 270)
        text_new_version_2 = page.get_text(clip=clip_new_version_2, sort=True).strip()
        try:
            text_new_version_2 = text_new_version_2.split('NT$')[1].strip()
            self.tp_2 = text_new_version_2.split('\n')[0].strip()
        except:
            self.tp_2 = 'NULL'

    def get_author_old_version_1(self, page, rect):
        clip_old_version_1 = fitz.Rect(0, 700, 200, rect.height)
        text_old_version_1 = page.get_text(clip=clip_old_version_1, sort=True).strip()
        try:
            text_old_version_1 = text_old_version_1.split('詳見最後頁聲明')[0].strip()
            self.author_1 = text_old_version_1.split(' ')[0].strip()
        except:
            self.author_1 = 'NULL'
    
    def get_author_old_version_2(self, page):
        clip_old_version_2 = fitz.Rect(0, 700, 200, 780)
        text_old_version_2 = page.get_text(clip=clip_old_version_2).strip()
        try:
            self.author_2 = text_old_version_2.split(' ')[0].strip()
        except:
            self.author_2 = 'NULL'

    def get_author_new_version_1(self, page, rect):
        clip_new_version_1 = fitz.Rect(0, 700, 200, rect.height)
        text_new_version_1 = page.get_text(clip=clip_new_version_1, sort=True).strip()
        try:
            text_new_version_1 = text_new_version_1.split('永豐證券投資顧問股份有限公司')[0].strip()
            self.author_1 = text_new_version_1.split(' ')[0].strip()
        except:
            self.author_1 = None
    
    def get_author_new_version_2(self, page):
        clip_new_version_2 = fitz.Rect(0, 700, 200, 780)
        text_new_version_2 = page.get_text(clip=clip_new_version_2).strip()
        try:
            self.author_2 = text_new_version_2.split(' ')[0].strip()
        except:
            self.author_2 = None

    def get_summary_old_version_1(self, page, rect):
        clip_old_version_1 = fitz.Rect(0, 0, rect.width, 170)
        text_old_version_1 = page.get_text(clip=clip_old_version_1, sort=True).strip()
        try:
            text_old_version_1 = text_old_version_1.split('目標價')[1].strip()
            self.summary_1 = text_old_version_1.split('\n')[0].strip()
        except:
            self.summary_1 = 'NULL'

    def get_summary_old_version_2(self, page):
        clip_old_version_2 = fitz.Rect(220, 125, 550, 155)
        text_old_version_2 = page.get_text(clip=clip_old_version_2).strip()
        try:
            self.summary_2 = text_old_version_2.split('\n')[0].strip()
        except:
            self.summary_2 = 'NULL'

    def get_summary_new_version_1(self, page, rect):
        clip_new_version_1 = fitz.Rect(200, 170, rect.width, 250)
        text_new_version_1 = page.get_text(clip=clip_new_version_1, sort=True).strip()
        try:
            text_new_version_1 = text_new_version_1.split(')')[1].strip()
            self.summary_1 = text_new_version_1.split('\n')[0].strip()
        except:
            self.summary_1 = None

    def get_summary_new_version_2(self, page):
        clip_new_version_2 = fitz.Rect(200, 200, 560, 250)
        text_new_version_2 = page.get_text(clip=clip_new_version_2).strip()
        try:
            self.summary_2 = text_new_version_2.split('\n')[0].strip()
        except:
            self.summary_2 = None

if __name__ == '__main__' :
    sample_file_root = r'./extract_pdf/extract_pdf/sample_file/1305_華夏_永豐_買進.pdf'
    pdfReader = Sinopac(sample_file_root)
    rating = pdfReader.get_rating_process()
    target_price = pdfReader.get_target_price_process()
    author = pdfReader.get_author_process()
    summary = pdfReader.get_summary_process()
    print(rating)
    print(target_price)
    print(author)
    print(summary)
    pass
