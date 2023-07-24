from tqdm import tqdm
import fitz
from ExtractPdf import *
import os
import datetime

class Capital(Adviser):
    '''Handle 群益(Capital) pdf

        Args :
            directory_path : (str) pdf path
        
        Return :
            rating : (str) recommend
    '''


    def __init__(self, file_path):
        super().__init__(file_path)
        self.possible_rating = ['中立', '中立', '區間操作', '買進']

    def get_all(self):
        doc, page = self.open()
        advisor, version = self.get_advisor_and_version(doc, page)
        stock = self.get_stock(page, version)
        date = self.get_date(page, version)
        rating_1, rating_2 = self.get_rating_1_and_2(page, version)
        rating = self.check_rating(rating_1, rating_2, self.possible_rating)
        tp_1, tp_2 = self.get_tp_1_and_2(page, version)
        tp = self.check_targrt_price(tp_1, tp_2)
        author_1, author_2 = self.get_author_1_and_2(page, version)
        author = self.check_author(author_1, author_2)
        summary_1, summary_2 = self.get_summary_1_and_2(page, version)
        summary = self.check_summary(summary_1, summary_2)
        self.close(doc)
        return advisor, version, stock, date, rating_1, rating_2, rating,\
            tp_1, tp_2, tp, author_1, author_2, author, \
                summary_1, summary_2, summary
        # return advisor, version, stock, date, rating, tp, author, summary

    def get_advisor_and_version(self, doc, page):
        advisor, version = 'NULL', 'NULL'
        page_check_source = doc.load_page(-1)
        clip_check_source = fitz.Rect(0, 0.9*page.rect.height, page.rect.width, page.rect.height)
        text_check_source = page_check_source.get_text(clip=clip_check_source)
        if self.check_source(text_check_source):
            advisor = self.__class__.__name__
            clip_check_version= fitz.Rect(page.rect.width/2, 0, page.rect.width, 70)
            text_check_version = page.get_text(clip=clip_check_version, sort=True).strip()
            version = self.check_version(text_check_version)
        return advisor, version
        
    def check_source(self, text_check_source):
        check_source = ['群益投顧']
        return True if any(keyword in text_check_source for keyword in check_source) else False
     
    def check_version(self, text_check_version):
        check_version = ['個股報告']
        return 'report' if any(keyword in text_check_version for keyword in check_version) else 'NULL'

    # stock
    def get_stock(self, page, version):
        if version == 'report':
            return self.get_stock_report_version(page)
        else:
            return 'NULL'

    def get_stock_report_version(self, page):
        clip_report_version_1 = fitz.Rect(200, 105, page.rect.width, 160)
        text_report_version_1 = page.get_text(clip=clip_report_version_1, sort=True).strip()
        try:
            if '（' in text_report_version_1:
                text_report_version_1 = text_report_version_1.split('（')[1].strip()
            elif '(' in text_report_version_1:
                text_report_version_1 = text_report_version_1.split('(')[1].strip()
            if '）' in text_report_version_1:
                text_report_version_1 = text_report_version_1.split('）')[0].strip()
            elif ')' in text_report_version_1:
                text_report_version_1 = text_report_version_1.split(')')[0].strip()
            return text_report_version_1.split('\n')[0].strip()
        except:
            return 'NULL'
    
    # date
    def get_date(self, page, version):
        if version == 'report':
            return self.get_date_report_version(page)
        else:
            return 'NULL'
    
    def get_date_report_version(self, page):
        clip_report_version_1 = fitz.Rect(0, 0, page.rect.width, 100)
        text_report_version_1 = page.get_text(clip=clip_report_version_1, sort=True).strip()
        try:
            text_report_version_1 = text_report_version_1.split('個股報告')[1].strip()
            Date_1 = text_report_version_1.split('\n')[0].strip()
            return datetime.datetime.strptime(Date_1, '%Y 年 %m 月 %d 日').date()
        except:
            return 'NULL'

    # rating
    def get_rating_1_and_2(self, page, version):
        rating_1, rating_2 = 'NULL', 'NULL'
        if version == 'report':
            rating_1 = self.get_rating_report_version_1(page)
            rating_2 = self.get_rating_report_version_2(page)
        return rating_1, rating_2
    
    def get_rating_report_version_1(self, page):
        clip_report_version_1= fitz.Rect(200, 105, page.rect.width, 160)
        text_report_version_1 = page.get_text(clip=clip_report_version_1, sort=True).strip()
        try:
            text_report_version_1 = text_report_version_1.split(')')[1].strip()
            rate_1 = text_report_version_1.split('\n')[0].strip()
            if '立' in str(rate_1):
                str(rate_1).replace('立', '立') # '立'會導致改檔名失敗
            return rate_1
        except:
            return 'NULL'

    def get_rating_report_version_2(self, page):
        clip_report_version_2 = fitz.Rect(400, 105, 565, 160)
        rate_2 = page.get_text(clip=clip_report_version_2, sort=True).strip()
        if '立' in str(rate_2):
            str(rate_2).replace('立', '立') # '立'會導致改檔名失敗
        return rate_2

    # target_price
    def get_tp_1_and_2(self, page, version):
        tp_1, tp_2 = 'NULL', 'NULL'
        if version == 'report':
            tp_1 = self.get_tp_report_version_1(page)
            tp_2 = self.get_tp_report_version_2(page)
        return tp_1, tp_2
    
    def get_tp_report_version_1(self, page):
        clip_report_version_1 = fitz.Rect(0, 80, 195, 170)
        text_report_version_1 = page.get_text(clip=clip_report_version_1, sort=True).strip()
        try:
            text_report_version_1 = text_report_version_1.split('目 標 價')[1].strip()
            text_report_version_1 = text_report_version_1.split('3 個 月')[1].strip()
            return text_report_version_1.split('元')[0].strip()
        except:
            return 'NULL'
    
    def get_tp_report_version_2(self, page):
        clip_report_version_2 = fitz.Rect(130, 140, 195, 155)
        text_report_version_2 = page.get_text(clip=clip_report_version_2).strip()
        try:
            return text_report_version_2.split('元')[0].strip()
        except:
            return 'NULL'

    # author   
    def get_author_1_and_2(self, page, version):
        author_1, author_2 = 'NULL', 'NULL'
        if version == 'report':
            author_1 = self.get_author_report_version_1(page)
            author_2 = self.get_author_report_version_2(page)
        return author_1, author_2

    def get_author_report_version_1(self, page):
        clip_report_version_1 = fitz.Rect(0, 80, 195, 170)
        text_report_version_1 = page.get_text(clip=clip_report_version_1, sort=True).strip()
        try:
            text_report_version_1 = text_report_version_1.split('研究員：')[1].strip()
            text_report_version_1 = text_report_version_1.split('@')[0].strip()
            return text_report_version_1.split(' ')[0].strip()
        except:
            return 'NULL'
    
    def get_author_report_version_2(self, page):
        clip_report_version_2 = fitz.Rect(70, 105, 92, 117)
        text_report_version_2 = page.get_text(clip=clip_report_version_2).strip()
        return text_report_version_2.split(' ')[0].strip()
        
    # summary   
    def get_summary_1_and_2(self, page, version):
        summary_1, summary_2 = 'NULL', 'NULL'
        if version == 'report':
            summary_1 = self.get_summary_report_version_1(page)
            summary_2 = self.get_summary_report_version_2(page)
        return summary_1, summary_2
    
    def get_summary_report_version_1(self, page):
        clip_report_version_1 = fitz.Rect(200, 160, page.rect.width, 240)
        text_report_version_1 = page.get_text(clip=clip_report_version_1, sort=True).strip()
        try:
            return text_report_version_1.split('\n投 資 建 議')[0].strip()
        except:
            return 'NULL'
        
    def get_summary_report_version_2(self, page):
        clip_report_version_2 = fitz.Rect(200, 160, page.rect.width, 190)
        return page.get_text(clip=clip_report_version_2, sort=True).strip()
