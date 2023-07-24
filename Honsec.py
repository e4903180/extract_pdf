from tqdm import tqdm
import fitz
from ExtractPdf import *
import os
import datetime

class Honsec(Adviser):
    '''Handle 宏遠(Honsec) pdf

        Args :
            directory_path : (str) pdf path
        
        Return :
            rating : (str) recommend
    '''


    def __init__(self, file_path):
        super().__init__(file_path)
        self.possible_rating = ['買進', '區間操作', '買進（調升）', '強力買進', '區間→買進',
                 '中立', '買進（維持）', '中立（調降）', '區間操作（調降）', '區間', '維持買進']

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
        if self.check_source(page_check_source):
            advisor = self.__class__.__name__
            version = self.check_version(page)
        return advisor, version
        
    def check_source(self, page_check_source):
        check_source = ['宏遠投顧']
        rect = page_check_source.rect
        clip_check_source = fitz.Rect(0, 0.9*rect.height, rect.width, rect.height)
        text_check_source = page_check_source.get_text(clip=clip_check_source)
        return True if any(keyword in text_check_source for keyword in check_source) else False
     
    def check_version(self, page):
        check_version = ['公司研究報告']
        clip_check_version= fitz.Rect(page.rect.width/2, 0, page.rect.width, 100)
        text_check_version = page.get_text(clip=clip_check_version, sort=True).strip()
        return 'report' if any(keyword in text_check_version for keyword in check_version) else 'NULL'

    # stock
    def get_stock(self, page, version):
        if version == 'report':
            return self.get_stock_report_version(page)
        else:
            return 'NULL'

    def get_stock_report_version(self, page):
        clip_report_version_1 = fitz.Rect(220, 100, page.rect.width, 150)
        text_report_version_1 = page.get_text(clip=clip_report_version_1, sort=True).strip()
        try:
            text_report_version_1 = page.get_text(clip=clip_report_version_1, sort=True).strip()
            return text_report_version_1.split(' ')[0].strip()
        except:
            return 'NULL'
    
    # date
    def get_date(self, page, version):
        if version == 'report':
            return self.get_date_report_version(page)
        else:
            return 'NULL'
    
    def get_date_report_version(self, page):
        clip_report_version_1 = fitz.Rect(220, 150, page.rect.width, 200)
        text_report_version_1 = page.get_text(clip=clip_report_version_1, sort=True).strip()
        try:
            text_report_version_1 = text_report_version_1.split('報告日期：')[1].strip()
            Date_1 = text_report_version_1.split('\n')[0].strip()
            return datetime.datetime.strptime(Date_1, '%Y/%m/%d').date()
        except:
            return 'NULL'

    # rating
    def get_rating_1_and_2(self, page, version):
        rating_1 = 'NULL'
        if version == 'report':
            rating_1 = self.get_rating_report_version_1(page)
        return rating_1, 'NULL'
    
    def get_rating_report_version_1(self, page):
        clip_report_version_1= fitz.Rect(0, 0, 220, page.rect.height)
        text_report_version_1 = page.get_text(clip=clip_report_version_1, sort=True).strip()
        try:
            if '投資評等:' in text_report_version_1:
                text_report_version_1 = text_report_version_1.split('投資評等:')[1].strip()
            else:
                text_report_version_1 = text_report_version_1.split('投資評等：')[1].strip()
            return text_report_version_1.split('\n')[0].strip()
        except:
            return 'NULL'

    # target_price
    def get_tp_1_and_2(self, page, version):
        tp_1 = 'NULL'
        if version == 'report':
            tp_1 = self.get_tp_report_version_1(page)
        return tp_1, 'NULL'
    
    def get_tp_report_version_1(self, page):
        clip_report_version_1 = fitz.Rect(0, 0, 220, 500)
        text_report_version_1 = page.get_text(clip=clip_report_version_1, sort=True).strip()
        try:
            if '目標價位' in text_report_version_1:
                text_report_version_1 = text_report_version_1.split('目標價位')[1].strip()
            elif '操作區間' in text_report_version_1:
                text_report_version_1 = text_report_version_1.split('操作區間')[1].strip()
            elif '目標區間' in text_report_version_1:
                text_report_version_1 = text_report_version_1.split('目標區間')[1].strip()
            elif '區間價位' in text_report_version_1:
                text_report_version_1 = text_report_version_1.split('區間價位')[1].strip()
            else:
                text_report_version_1 = text_report_version_1.split('區間價')[1].strip()
            text_report_version_1 = text_report_version_1.split('\n')[0].strip()
            if ':' in text_report_version_1:
                text_report_version_1 = text_report_version_1.replace(':', '').strip()
            elif '：' in text_report_version_1:
                text_report_version_1 = text_report_version_1.replace('：', '').strip()
            return text_report_version_1.split('元')[0].strip()
        except:
            return 'NULL'

    # author   
    def get_author_1_and_2(self, page, version):
        author_1 = 'NULL'
        if version == 'report':
            author_1 = self.get_author_report_version_1(page)
        return author_1, 'NULL'

    def get_author_report_version_1(self, page):
        clip_report_version_1 = fitz.Rect(220, 150, page.rect.width, 200)
        text_report_version_1 = page.get_text(clip=clip_report_version_1, sort=True).strip()
        try:
            if '研 究 員' in text_report_version_1:
                text_report_version_1 = text_report_version_1.split('研 究 員')[1].strip()
            else:
                text_report_version_1 = text_report_version_1.split('研究員')[1].strip()
            text_report_version_1 = text_report_version_1.split('\n')[0].strip()
            if ':' in text_report_version_1:
                text_report_version_1 = text_report_version_1.replace(':', '').strip()
            elif '：' in text_report_version_1:
                text_report_version_1 = text_report_version_1.replace('：', '').strip()
            return text_report_version_1.split('\n')[0].strip()
        except:
            return 'NULL'
        
    # summary   
    def get_summary_1_and_2(self, page, version):
        summary_1 = 'NULL'
        if version == 'report':
            summary_1 = self.get_summary_report_version_1(page)
        return summary_1, 'NULL'
    
    def get_summary_report_version_1(self, page):
        clip_report_version_1 = fitz.Rect(220, 190, page.rect.width, page.rect.height)
        text_report_version_1 = page.get_text(clip=clip_report_version_1, sort=True).strip()
        try:
            title = ['一.', '二.', '三.', '四.', '五.']
            content = list()
            for num in title:
                if num in text_report_version_1:
                    text_report_version_1 = text_report_version_1.split(num)[1].strip()
                    content.insert(-1, text_report_version_1.split('\n')[0].strip())
            return '。'.join(content)
        except:
            return 'NULL'
