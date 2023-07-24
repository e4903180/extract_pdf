from tqdm import tqdm
import fitz
from ExtractPdf import *
import os

class Fubon(Adviser):
    '''Handle 富邦投顧(Fubon) pdf

        Args :
            directory_path : (str) pdf path
        
        Return :
            rating : (str) recommend

        Todo :
            fix get tp and author 
    '''
    def __init__(self, file_path):
        super().__init__(file_path)
        self.file_path = file_path
        self.possible_rating = ['增加持股', '未評等', '中立', '買進', '降低持股', 'Buy', 'Neutral']

    def get_rating_process(self):
        with fitz.open(self.file_path) as doc:
            page = doc.load_page(0)
            if self.check_source(doc):
                self.advisor = self.__class__.__name__
                self.get_rating_version_1(page)
                self.get_rating_version_2(page)
        return self.check_rating(self.rating_1, self.rating_2, self.possible_rating)

    def get_target_price_process(self):
        with fitz.open(self.file_path) as doc:
            page = doc.load_page(0)
            if self.check_source(doc):
                self.advisor = self.__class__.__name__
                self.get_target_price_version_1(page)
                self.get_target_price_version_2(page)
        return self.check_targrt_price(self.tp_1, self.tp_2)
    
    def get_author_process(self):
        with fitz.open(self.file_path) as doc:
            page = doc.load_page(0)
            rect = page.rect
            if self.check_source(doc):
                self.advisor = self.__class__.__name__
                self.get_author_version_1(page, rect)
                self.get_author_version_2(page)
        return self.check_author(self.author_1, self.author_2)

    def get_summary_process(self):
        with fitz.open(self.file_path) as doc:
            page = doc.load_page(0)
            rect = page.rect
            if self.check_source(doc):
                self.advisor = self.__class__.__name__
                self.get_summary_version_1(page, rect)
                self.get_summary_version_2(page, rect)
        return self.check_summary(self.summary_1, self.summary_2)

    def check_source(self, doc):
        page_check_source = doc.load_page(-1)
        text_check_source = page_check_source.get_text()
        if '富邦投顧' in text_check_source :
            return True
        else:
            return False

    def check_report(self, text_check_report):
        pass
    
    def get_rating_version_1(self, page):
        clip_version_1 = fitz.Rect(50, 120, 200, 200)
        text_version_1 = page.get_text(clip=clip_version_1, sort=True).strip()
        if text_version_1.split('\n')[0].strip() == '_':
            self.rating_1 = text_version_1.split('\n')[1].strip()
        else :
            self.rating_1 = text_version_1.split('\n')[0].strip()

    def get_rating_version_2(self, page):
        clip_version_2 = fitz.Rect(50, 140, 210, 165)
        self.rating_2 = page.get_text(clip=clip_version_2, sort=True).strip()

    def get_target_price_version_1(self, page):
        clip_version_1 = fitz.Rect(50, 120, 205, 220)
        text_version_1 = page.get_text(clip=clip_version_1, sort=True).strip()
        try:
            text_version_1 = text_version_1.split('$')[2].strip()
            self.tp_1 = text_version_1.split('\n')[0].strip()
        except:
            self.tp_1 = 'NULL'
    
    def get_target_price_version_2(self, page):
        clip_version_2 = fitz.Rect(140, 150, 205, 177)
        text_version_2 = page.get_text(clip=clip_version_2).strip()
        try:
            self.tp_2 = text_version_2.split('$')[1].strip()
        except:
            self.tp_2 = 'NULL'

    def get_author_version_1(self, page, rect):
        clip_version_1 = fitz.Rect(0, 760, 200, rect.height)
        text_version_1 = page.get_text(clip=clip_version_1, sort=True).strip()
        try:
            text_version_1 = text_version_1.split('-')[0].strip()
            self.author_1 = text_version_1.split('\n')[0].strip()
        except:
            self.author_1 = 'NULL'
    
    def get_author_version_2(self, page):
        clip_version_2 = fitz.Rect(30, 765, 100, 777)
        text_version_2 = page.get_text(clip=clip_version_2).strip()
        self.author_2 = text_version_2.split('\n')[0].strip()

    def get_summary_version_1(self, page, rect):
        clip_version_1 = fitz.Rect(220, 60, rect.width, 140)
        text_version_1 = page.get_text(clip=clip_version_1, sort=True).strip()
        try:
            text_version_1 = text_version_1.split(')')[1].strip()
            self.summary_1 = text_version_1.split('\n')[0].strip()
        except:
            self.summary_1 = 'NULL'

    def get_summary_version_2(self, page, rect):
        clip_version_2 = fitz.Rect(220, 105, rect.width, 140)
        text_version_2 = page.get_text(clip=clip_version_2).strip()
        self.summary_2 = text_version_2.split('\n')[0].strip()

if __name__ == '__main__' :
    sample_file_root = r'./extract_pdf/extract_pdf/sample_file/1312_國喬_富邦_中立.pdf'
    pdfReader = Fubon(sample_file_root)
    rating = pdfReader.get_rating_process()
    target_price = pdfReader.get_target_price_process()
    author = pdfReader.get_author_process()
    summary = pdfReader.get_summary_process()
    print(rating)
    print(target_price)
    print(author)
    print(summary)
    pass
