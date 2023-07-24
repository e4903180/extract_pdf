from tqdm import tqdm
import fitz
from ExtractPdf import *
import os

class Ctbc(Adviser):
    '''Handle 中信託(CTBC) pdf

        Args :
            directory_path : (str) pdf path
        
        Return :
            rating : (str) recommend
    '''


    def __init__(self, file_path):
        super().__init__(file_path)
        self.possible_rating = ['中立', '買進', '增加持股(Overweight)', '中立(Neutral)', 
                              '買進(Buy)', '增加持股', '-', '降低持股(Underweight)', '未評等']
        self.version = None
        self.doc = None

    def run(self):
        self.open()
        self.init_version()
        self.get_rating_process()
        self.get_target_price_process()
        self.get_author_process()
        self.get_summary_process()
        self.close()

    def open(self):
        doc = fitz.open(self.file_path)
        page = doc.load_page(0)
        self.doc = doc
        self.page = page
        self.rect = page.rect

    def close(self):
        if type(self.doc) == fitz.fitz.Document:
            self.doc.close()

    def init_version(self):
        page_check_source = self.doc.load_page(-1)
        text_check_source = page_check_source.get_text()
        if self.check_source(text_check_source):
            self.advisor = self.__class__.__name__
            text_check_report = text_check_source
            if self.check_report(text_check_report):
                clip_check_version= fitz.Rect(370, 80, 450, 200)
                text_check_version = self.page.get_text(clip=clip_check_version, sort=True).strip()
                if self.check_report(text_check_version) == 'old':
                    self.version = 'old'
                else:
                    self.version = 'new'
        
    def check_source(self, text_check_source):
        check_source = ['中國信託金融控股', '中信投顧投資分析報告']
        return True if any(keyword in text_check_source for keyword in check_source) else False
    
    def check_report(self, text_check_report):
        return True if '個股報告' in text_check_report else False
     
    def check_version(self, text_check_version):
        return 'old' if '投資評等' in text_check_version else 'new'

    def get_rating_process(self):
        if self.version == 'old':
            self.get_rating_old_version_1()
            self.get_rating_old_version_2()
        elif self.version == 'new':
            self.get_rating_new_version_1()
            self.get_rating_new_version_2()
        self.rating = self.check_rating(self.rating_1, self.rating_2, self.possible_rating)

    def get_target_price_process(self):
        self.get_target_price_old_version_1()
        self.get_target_price_old_version_2()
        self.tp =  self.check_targrt_price(self.tp_1, self.tp_2)
    
    def get_author_process(self):
        self.get_author_old_version_1()
        self.get_author_old_version_2()
        self.author = self.check_author(self.author_1, self.author_2)

    def get_summary_process(self):
        clip_check_version= fitz.Rect(370, 80, 450, 200)
        text_check_version = self.page.get_text(clip=clip_check_version, sort=True).strip()
        if self.version == 'old':
            self.get_summary_old_version_1()
        elif self.version == 'new':
            self.get_summary_new_version_1()
        self.summary = self.check_summary(self.summary_1, self.summary_2)
    
    def get_rating_old_version_1(self):
        clip_old_version_1= fitz.Rect(370, 80, 450, 200)
        text_old_version_1 = self.page.get_text(clip=clip_old_version_1, sort=True).strip()
        try:
            text_old_version_1 = text_old_version_1.split('投資評等')[1].strip()
            self.rating_1 = text_old_version_1.split('\n')[0].strip()
        except:
            self.rating_1 = 'NULL'

    def get_rating_old_version_2(self):
        clip_old_version_2 = fitz.Rect(370, 120, 430, 150)
        self.rating_2 = self.page.get_text(clip=clip_old_version_2, sort=True).strip()
    
    def get_rating_new_version_1(self):
        clip_new_version_1 = fitz.Rect(200, 0, self.rect.width, 200)
        text_new_version_1 = self.page.get_text(clip=clip_new_version_1, sort=True).strip()
        try:
            text_new_version_1 = text_new_version_1.split('評 等')[1]
            self.rating_1 = text_new_version_1.split('\n')[1].strip()
        except:
            self.rating_1 = 'NULL'
        
    def get_rating_new_version_2(self):
        clip_new_version_2 = fitz.Rect(350, 115, 570, 200)
        text_new_version_1 = self.page.get_text(clip=clip_new_version_2, sort=True).strip()
        self.rating_2 = text_new_version_1.split('\n')[0].strip()

    def get_target_price_old_version_1(self):
        clip_old_version_1 = fitz.Rect(0, 100, 200, 200)
        text_old_version_1 = self.page.get_text(clip=clip_old_version_1, sort=True).strip()
        try:
            text_old_version_1 = text_old_version_1.split('目 標 價')[1].strip()
            self.tp_1 = text_old_version_1.split('元')[0].strip()
        except:
            self.tp_1 = 'NULL'
    
    def get_target_price_old_version_2(self):
        clip_old_version_2 = fitz.Rect(0, 165, 200, 175)
        text_old_version_2 = self.page.get_text(clip=clip_old_version_2).strip()
        try:
            text_old_version_2 = text_old_version_2.split('目 標 價')[1].strip()
            self.tp_2 = text_old_version_2.split('元')[0].strip()
        except:
            self.tp_2 = 'NULL'

    def get_author_old_version_1(self):
        clip_old_version_1 = fitz.Rect(0, 555, 200, 620)
        text_old_version_1 = self.page.get_text(clip=clip_old_version_1, sort=True).strip()
        try:
            text_old_version_1 = text_old_version_1.split('@')[0].strip()
            text_old_version_1 = text_old_version_1.split('\n')[0].strip()
            self.author_1 = text_old_version_1.split('(')[0].strip()
        except:
            self.author_1 = 'NULL'
    
    def get_author_old_version_2(self):
        clip_old_version_2 = fitz.Rect(0, 0, 200, self.rect.height)
        text_old_version_2 = self.page.get_text(clip=clip_old_version_2).strip()
        try:
            text_old_version_2 = text_old_version_2.split('個股與大盤走勢')[1].strip()
            text_old_version_2 = text_old_version_2.split('@')[0].strip()
            text_old_version_2 = text_old_version_2.split('\n')[-2].strip()
            self.author_2 = text_old_version_2.split('(')[0].strip()
        except:
            self.author_2 = 'NULL'

    def get_summary_old_version_1(self):
        clip_old_version_1 = fitz.Rect(210, 145, self.rect.width, 300)
        text_old_version_1 = self.page.get_text(clip=clip_old_version_1, sort=True).strip()
        try:
            text_old_version_1 = text_old_version_1.split('結論與建議')[1].strip()
            text_old_version_1 = text_old_version_1.split('：')[0].strip()
            self.summary_1 = ''.join(text_old_version_1.split('\n')[:-1]).strip()
        except:
            self.summary_1 = 'NULL'

    def get_summary_new_version_1(self):
        clip_new_version_1 = fitz.Rect(210, 145, self.rect.width, 300)
        text_new_version_1 = self.page.get_text(clip=clip_new_version_1, sort=True).strip()
        try:
            self.summary_1 = text_new_version_1.split('\uf06c')[0].strip()
        except:
            self.summary_1 = 'NULL'

if __name__ == '__main__' :
    sample_file_root = r'./extract_pdf/extract_pdf/sample_file/1513_中興電_中信託_買進.pdf'
    pdfReader = Ctbc(sample_file_root)
    pdfReader.run()
    # rating = pdfReader.get_rating_process()
    # tp = pdfReader.get_target_price_process()
    # author = pdfReader.get_author_process()
    # summary = pdfReader.get_summary_process()
    print(pdfReader.rating)
    print(pdfReader.tp)
    print(pdfReader.author)
    print(pdfReader.summary)

    pass