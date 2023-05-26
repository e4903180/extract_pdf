from tqdm import tqdm
import fitz
from ExtractPdf import *


class Ibf(Adviser):
    '''Handle 國票 pdf

        Args :
            directory_path : (str) pdf path
        
        Return :
            rating : (str) recommend
    '''
    def __init__(self, file_path):
        super().__init__(file_path)
        self.possible_rating = ['買進', '區間操作', '強力買進']

    def get_rating_process(self):
        with fitz.open(self.file_path) as doc:
            page = doc.load_page(0)
            rect = page.rect
            if self.check_source(doc, rect):
                if self.check_version(page, rect) =='old':
                    self.get_rating_old_version_1(page, rect)
                    self.get_rating_old_version_2(page)
                else:
                    self.get_rating_new_version_1(page)
                    self.get_rating_new_version_2(page)
        return self.check_rating(self.rating_1, self.rating_2, self.possible_rating)

    def get_target_price_process(self):
        with fitz.open(self.file_path) as doc:
            page = doc.load_page(0)
            rect = page.rect
            if self.check_source(doc, rect):
                if self.check_version(page, rect) =='old':
                    self.get_target_price_old_version_1(page, rect)
                    self.get_target_price_old_version_2(page)
                else:
                    self.get_target_price_new_version_1(page)
                    self.get_target_price_new_version_2(page)
        return self.check_targrt_price(self.tp_1, self.tp_2)
    
    def get_author_process(self):
        with fitz.open(self.file_path) as doc:
            page = doc.load_page(0)
            rect = page.rect
            if self.check_source(doc, rect):
                self.get_author_version_1(page, rect)
                self.get_author_version_2(page, rect)
        return self.check_author(self.author_1, self.author_2)

    def get_summary_process(self):
        with fitz.open(self.file_path) as doc:
            page = doc.load_page(0)
            rect = page.rect
            if self.check_source(doc, rect):
                self.get_summary_version_1(page, rect)
                self.get_summary_version_2(page, rect)
        return self.check_summary(self.summary_1, self.summary_2)

    def check_source(self, doc, rect):
        page_check_source = doc.load_page(-1)
        clip_check_source = fitz.Rect(0, 0, rect.width, rect.height)
        text_check_source = page_check_source.get_text(clip=clip_check_source, sort=True)
        if '國票投顧所有' in text_check_source:
            self.advisor = 'Ibf'
            return True
        else:
            return False
        
    def check_version(self, page, rect):
        clip_check_version = fitz.Rect(40, 0, rect.width, 400)
        text_check_version = page.get_text(clip=clip_check_version, sort=True).strip()
        if '國票觀點' in text_check_version:
            return 'old'
     
    def get_rating_old_version_1(self, page, rect):
        clip_old_version_1 = fitz.Rect(380, 0, rect.width, 400)
        text_old_version_1 = page.get_text(clip=clip_old_version_1, sort=True).strip()
        try:
            if '目標價' in text_old_version_1:
                text_old_version_1 = text_old_version_1.split('目標價')[1].strip()
                self.rating_1 = text_old_version_1.split('\n')[0].strip()
            elif '區間價位' in text_old_version_1:
                text_old_version_1 = text_old_version_1.split('區間價位')[1].strip()
                self.rating_1 = text_old_version_1.split('\n')[0].strip()
            elif '操作區間' in text_old_version_1:
                text_old_version_1 = text_old_version_1.split('操作區間')[1].strip()
                self.rating_1 = text_old_version_1.split('\n')[0].strip()
            elif '/買進' in text_old_version_1:
                text_old_version_1 = text_old_version_1.split('/買進')[1].strip()
                self.rating_1 = text_old_version_1.split('\n')[0].strip()
        except:
            self.rating_1 = 'NULL'

    def get_rating_old_version_2(self, page):
        clip_old_version_2 = fitz.Rect(380, 200, 470, 270)
        text_old_version_2 = page.get_text(clip=clip_old_version_2, sort=True).strip()
        if '買進' in text_old_version_2:
            self.rating_2 = '買進'
        elif '區間操作' in text_old_version_2:
            self.rating_2 = '區間操作'
        elif '賣出' in text_old_version_2:
            self.rating_2 = '賣出'
    
    def get_rating_new_version_1(self, page):
        clip_new_version_1 = fitz.Rect(30, 200, 220, 400)
        text_new_version_1 = page.get_text(clip=clip_new_version_1, sort=True).strip()
        try:
            if '目標價' in text_new_version_1:
                text_new_version_1 = text_new_version_1.split('目標價')[1].strip()
                self.rating_1 = text_new_version_1.split('\n')[0].strip()  
            elif '區間價位' in text_new_version_1:
                text_new_version_1 = text_new_version_1.split('區間價位')[1].strip()
                self.rating_1 = text_new_version_1.split('\n')[0].strip()
            elif '操作區間' in text_new_version_1:
                text_new_version_1 = text_new_version_1.split('操作區間')[1].strip()
                self.rating_1 = text_new_version_1.split('\n')[0].strip()  
        except:
            self.rating_1 = 'NULL'
        
    def get_rating_new_version_2(self, page):
        clip_new_version_2 = fitz.Rect(40, 200, 120, 400)
        text_new_version_2 = page.get_text(clip=clip_new_version_2, sort=True).strip()
        if '強力買進' in text_new_version_2:
            self.rating_2 = '強力買進'
        elif '買進' in text_new_version_2:
            self.rating_2 = '買進'
        elif '區間操作' in text_new_version_2:
            self.rating_2 = '區間操作'
        elif '賣出' in text_new_version_2:
            self.rating_2 = '賣出'

    def get_target_price_old_version_1(self, page, rect):
        clip_old_version_1 = fitz.Rect(380, 0, rect.width, 400)
        text_old_version_1 = page.get_text(clip=clip_old_version_1).strip()
        try:
            if '目標價' in text_old_version_1:
                text_old_version_1 = text_old_version_1.split('目標價')[1].strip()
                self.tp_1 = text_old_version_1.split('\n')[1].strip()
            elif '區間價位' in text_old_version_1:
                text_old_version_1 = text_old_version_1.split('區間價位')[1].strip()
                self.tp_1 = text_old_version_1.split('\n')[1].strip()
            elif '操作區間' in text_old_version_1:
                text_old_version_1 = text_old_version_1.split('操作區間')[1].strip()
                self.tp_1 = text_old_version_1.split('\n')[1].strip()
            elif '/買進' in text_old_version_1:
                text_old_version_1 = text_old_version_1.split('/買進')[1].strip()
                self.tp_1 = text_old_version_1.split('\n')[1].strip()
            else:
                self.tp_1 = 'NULL'
        except:
            self.tp_1 = 'NULL'
    
    def get_target_price_old_version_2(self, page):
        clip_old_version_2 = fitz.Rect(470, 245, 560, 265)
        self.tp_2 = page.get_text(clip=clip_old_version_2).strip()
    
    def get_target_price_new_version_1(self, page):
        clip_new_version_1 = fitz.Rect(30, 200, 200, 400)
        text_new_version_1 = page.get_text(clip=clip_new_version_1, sort=True).strip()
        try:
            if '目標價' in text_new_version_1:
                text_new_version_1 = text_new_version_1.split('目標價')[1].strip()
                self.tp_1 = text_new_version_1.split('\n')[1].strip()  
            elif '區間價位' in text_new_version_1:
                text_new_version_1 = text_new_version_1.split('區間價位')[1].strip()
                self.tp_1 = text_new_version_1.split('\n')[1].strip()  
            elif '操作區間' in text_new_version_1:
                text_new_version_1 = text_new_version_1.split('操作區間')[1].strip()
                self.tp_1 = text_new_version_1.split('\n')[1].strip()  
            else:
                self.tp_1 = 'NULL'
        except:
            self.tp_1 = 'NULL'
    
    def get_target_price_new_version_2(self, page):
        clip_new_version_2 = fitz.Rect(30, 200, 200, 400)
        text_new_version_2 = page.get_text(clip=clip_new_version_2, sort=True).strip() 
        try:
            text_new_version_2 = text_new_version_2.split(self.rating_1)[1].strip()
            self.tp_2 = text_new_version_2.split('\n')[0].strip()
        except:
            self.tp_2 = 'NULL'

    def get_author_version_1(self, page, rect):
        clip_new_version_1 = fitz.Rect(400, 0, rect.width, 80)
        text_new_version_1 = page.get_text(clip=clip_new_version_1, sort=True).strip()
        try:
            text_new_version_1 = text_new_version_1.split('\n')[1].strip()
            self.author_1 = text_new_version_1.split(' ')[0].strip()  
        except:
            self.author_1 = 'NULL'
    
    def get_author_version_2(self, page, rect):
        clip_new_version_2 = fitz.Rect(400, 60, rect.width, 80)
        text_new_version_2 = page.get_text(clip=clip_new_version_2, sort=True).strip() 
        self.author_2 = text_new_version_2.split(' ')[0].strip()

    def get_summary_version_1(self, page, rect):
        clip_new_version_1 = fitz.Rect(0, 80, rect.width, 190)
        text_new_version_1 = page.get_text(clip=clip_new_version_1, sort=True).strip()
        try:
            self.summary_1 = text_new_version_1.split('\n')[-1].strip()  
        except:
            self.summary_1 = 'NULL'

    def get_summary_version_2(self, page, rect):
        clip_new_version_2 = fitz.Rect(0, 125, rect.width, 190)
        self.summary_2 = page.get_text(clip=clip_new_version_2, sort=True).strip() 

if __name__ == '__main__' :
    sample_file_root = r'./extract_pdf/extract_pdf/sample_file/1308_亞聚_國票_區間操作.pdf'
    pdfReader = Ibf(sample_file_root)
    rating = pdfReader.get_rating_process()
    target_price = pdfReader.get_target_price_process()
    author = pdfReader.get_author_process()
    summary = pdfReader.get_summary_process()
    print(rating)
    print(target_price)
    print(author)
    print(summary)
    pass