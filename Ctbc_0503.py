from tqdm import tqdm
import fitz
from ExtractPdf_0503 import *
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

    def get_rating_process(self):
        with fitz.open(self.file_path) as doc:
            page = doc.load_page(0)
            rect = page.rect
            self.check_source(doc, page, rect)
        return self.check_rating(self.rating_1, self.rating_2, self.possible_rating)

    def check_source(self, doc, page, rect):
        page_check_source = doc.load_page(-1)
        text_check_source = page_check_source.get_text()
        check_source = ['中國信託金融控股', '中信投顧投資分析報告']
        if any(keyword in text_check_source for keyword in check_source):
            self.advisor = 'Ctbc'
            text_check_report = text_check_source
            self.check_report(page, rect, text_check_report)
    
    def check_report(self, page, rect, text_check_report):
        if '個股報告' in text_check_report:
            self.check_version(page, rect)
     
    def check_version(self, page, rect):
        clip_check_version= fitz.Rect(370, 80, 450, 200)
        text_check_version = page.get_text(clip=clip_check_version, sort=True).strip()
        if '投資評等' in text_check_version:
            self.get_rating_old_version_1(text_check_version)
            self.get_rating_old_version_2(page)
        else:
            self.get_rating_new_version_1(page, rect)
            self.get_rating_new_version_2(page)

    def get_rating_old_version_1(self, text_check_version):
        text_old_version_1 = text_check_version
        try:
            text_old_version_1 = text_old_version_1.split('投資評等')[1].strip()
            self.rating_1 = text_old_version_1.split('\n')[0].strip()
        except:
            self.rating_1 = 'NULL'

    def get_rating_old_version_2(self, page):
        clip_old_version_2 = fitz.Rect(370, 120, 430, 150)
        self.rating_2 = page.get_text(clip=clip_old_version_2, sort=True).strip()
    
    def get_rating_new_version_1(self, page, rect):
        clip_new_version_1 = fitz.Rect(200, 0, rect.width, 200)
        text_new_version_1 = page.get_text(clip=clip_new_version_1, sort=True).strip()
        try:
            text_new_version_1 = text_new_version_1.split('評 等')[1]
            self.rating_1 = text_new_version_1.split('\n')[1].strip()
        except:
            self.rating_1 = 'NULL'
        
    def get_rating_new_version_2(self, page):
        clip_new_version_2 = fitz.Rect(350, 115, 570, 200)
        text_new_version_1 = page.get_text(clip=clip_new_version_2, sort=True).strip()
        self.rating_2 = text_new_version_1.split('\n')[0].strip()

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
    sample_file_root = r'./hw_pdf/class/sample_file/1513_中興電_中信託_買進.pdf'
    pdfReader = Ctbc(sample_file_root)
    rating = pdfReader.get_rating_process()
    print(rating)
    pass