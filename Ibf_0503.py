from tqdm import tqdm
import fitz
from ExtractPdf_0503 import *


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

    def get_rating_process(self, file_path):
        with fitz.open(file_path) as doc:
            page = doc.load_page(0)
            rect = page.rect
            self.check_source(doc, page, rect)
        return self.check_rating(self.rating_1, self.rating_2, self.possible_rating)

    def check_source(self, doc, page, rect):
        page_check_source = doc.load_page(-1)
        clip_check_source = fitz.Rect(0, 0, rect.width, rect.height)
        text_check_source = page_check_source.get_text(clip=clip_check_source, sort=True)
        if '國票投顧所有' in text_check_source:
            self.advisor = 'Ibf'
            self.check_version(page, rect)
        
    def check_version(self, page, rect):
        clip_check_version = fitz.Rect(40, 0, rect.width, 400)
        text_check_version = page.get_text(clip=clip_check_version, sort=True).strip()
        if '國票觀點' in text_check_version:
            self.get_rating_old_version_1(page, rect)
            self.get_rating_old_version_2(page)
        else:  
            self.get_rating_new_version_1(page)
            self.get_rating_new_version_2(page)
     
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


if __name__ == '__main__' :
    # sample_file_root = r'./hw_pdf/class/sample_file/1308_亞聚_國票_區間操作.pdf'
    # pdfReader = Ibf(sample_file_root)
    # rating = pdfReader.get_rating_process(sample_file_root)
    # print(rating)
    pass