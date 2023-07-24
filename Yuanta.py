from tqdm import tqdm
import fitz
from ExtractPdf import *
import os

class Yuanta(Adviser):
    '''Handle 元大投顧(Yuanta) pdf

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
        self.possible_rating = ['持有-超越同業 (維持評等)', '買進 (維持評等)', '持有-落後同業', '持有-落後同業 (維持評等)', 
                 '買進 (調升評等)', '買進 (重新納入研究範圍)',
                 '持有-超越同業 (調降評等)', '買進 (研究員異動)', '買進  (初次報告)', 
                 '買進 (初次報告)', '持有-超越同業', '持有-落後同業(維持評等)', '賣出 (維持評等)', 
                 '持有-超越大盤(維持評等)', '持有-超越大盤 (維持評等)', '買進', '持有-落後大盤']

    def get_rating_process(self):
        with fitz.open(self.file_path) as doc:
            page = doc.load_page(0)
            rect = page.rect
            if self.check_source(doc):
                self.advisor = self.__class__.__name__
                if self.check_report(doc, rect):
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
        if '元大證券投資顧問' in text_check_source :
            return True
        else:
            return False

    def check_report(self, doc, rect):
        page_check_report = doc.load_page(0)
        clip_check_report = fitz.Rect(0, 0, rect.width, 70)
        text_check_report = page_check_report.get_text(clip=clip_check_report).strip()
        if '更新報告' in text_check_report or '初次報告' in text_check_report or 'Top Call 首選' in text_check_report:
            return True
        else:
            return False
    
    def get_rating_version_1(self, page):
        clip_version_1 = fitz.Rect(0, 0, 210, 230)
        text_version_1 = page.get_text(clip=clip_version_1, sort=True).strip()
        text_version_1 = text_version_1.split('目標價')[0].strip()
        self.rating_1 = text_version_1.split('\n')[-1].strip()

    def get_rating_version_2(self, page):
        clip_version_2 = fitz.Rect(0, 115, 210, 145)
        text_new_version_1 = page.get_text(clip=clip_version_2).strip()
        self.rating_2 = text_new_version_1.split('\n')[0].strip()

    def get_target_price_version_1(self, page):
        clip_version_1 = fitz.Rect(0, 0, 210, 230)
        text_version_1 = page.get_text(clip=clip_version_1, sort=True).strip()
        try:
            text_version_1 = text_version_1.split('目標價')[1].strip()
            text_version_1 = text_version_1.split('NT$')[1].strip()
            self.tp_1 = text_version_1.split('\n')[0].strip()
        except:
            self.tp_1 = 'NULL'
    
    def get_target_price_version_2(self, page):
        clip_version_2 = fitz.Rect(30, 135, 120, 170)
        text_version_2 = page.get_text(clip=clip_version_2).strip()
        try:
            text_version_2 = text_version_2.split('NT$')[1].strip()
            self.tp_2 = text_version_2.split('\n')[0].strip()
        except:
            self.tp_2 = 'NULL'

    def get_author_version_1(self, page, rect):
        clip_version_1 = fitz.Rect(0, 725, 220, rect.height)
        text_version_1 = page.get_text(clip=clip_version_1, sort=True).strip()
        try:
            text_version_1 = text_version_1.split('@')[0].strip()
            self.author_1 = text_version_1.split('\n')[0].strip()
        except:
            self.author_1 = 'NULL'
    
    def get_author_version_2(self, page):
        clip_version_2 = fitz.Rect(30, 725, 220, 740)
        text_version_2 = page.get_text(clip=clip_version_2).strip()
        self.author_2 = text_version_2.split('\n')[0].strip()

    def get_summary_version_1(self, page, rect):
        clip_version_1 = fitz.Rect(225, 65, rect.width, 200)
        text_version_1 = page.get_text(clip=clip_version_1, sort=True).strip()
        try:
            self.summary_1 = text_version_1.split('►')[0].strip()
        except:
            self.summary_1 = 'NULL'

    def get_summary_version_2(self, page, rect):
        # clip_version_2 = fitz.Rect(225, 65, rect.width, 95)
        # text_version_2 = page.get_text(clip=clip_version_2).strip()
        # self.summary_2 = text_version_2.split('\n')[0].strip()
        self.summary_2 = self.summary_1

if __name__ == '__main__' :
    sample_file_root = r'./extract_pdf/extract_pdf/sample_file/2049_上銀_元大_買進.pdf'
    pdfReader = Yuanta(sample_file_root)
    rating = pdfReader.get_rating_process()
    target_price = pdfReader.get_target_price_process()
    author = pdfReader.get_author_process()
    summary = pdfReader.get_summary_process()
    print(rating)
    print(target_price)
    print(author)
    print(summary)
    pass
