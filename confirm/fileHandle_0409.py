import json
import datetime
import os
import sys
from tqdm import tqdm
import fitz

root_path = json.load(open("../../root_path.json"))



class ExtractPdfRate():
    """Create handle for vary investment companies
    """
    def __init__(self):
        pass

    def check_rate(self, rate_1:str, rate_2:str, possible_rate:list):
        '''Check that the extracted ratings are correct
        
            Args :
                rate_1 : (str) extracted by method 1
                rate_2 : (str) extracted by method 2
                possible_rate : (list) all possible ratings
            Return :
                rate : (str) recommend
        '''
        if rate_1 == rate_2:
            return rate_1 if rate_1 != 'NULL' else 'NULL'
        for rate in possible_rate:
            if rate == rate_1:
                return rate
            elif rate == rate_2:
                return rate
        return 'NULL'

    def sinopac(self, directory_path:str):
        '''Handle 永豐投顧(SinoPac) pdf

            Args :
                directory_path : (str) pdf path
            
            Return :
                rate : (str) recommend
        '''
        possible_rate = ['買進', '中立']
        rate_1, rate_2 = 'NULL'
        with fitz.open(directory_path) as doc:
            page = doc.load_page(0)
            rect = page.rect
            # 檢查是否由永豐投顧出版
            page_check_source = doc.load_page(-1)
            text_check_source = page_check_source.get_text()
            if '永豐證券投資顧問股份有限公司' in text_check_source:
                # 報告為舊版
                # 提取評價的第一種方法
                clip_old_version_1 = fitz.Rect(220, 80, 560, 140)
                text_old_version_1 = page.get_text(clip=clip_old_version_1, sort=True).strip()
                text_old_version_1 = text_old_version_1.split('）')[1].strip()
                rate_1 = text_old_version_1.split('\n')[1].strip()
                # 提取評價的第二種方法
                clip_old_version_2 = fitz.Rect(425, 90, 560, 130)
                rate_2 = page.get_text(clip=clip_old_version_2, sort=True).strip()
            elif 'SinoPac Securities' in text_check_source:
                # 報告為新版
                # 檢查報告版本
                clip_check_report = fitz.Rect(0, 0, rect.width, 150)
                text_check_report = page.get_text(clip=clip_check_report, sort=True).strip()
                if '個股聚焦' in text_check_report:
                    # 提取評價的第一種方法
                    clip_new_version_1 = fitz.Rect(0, 0, 200, 400)
                    text_new_version_1 = page.get_text(clip=clip_new_version_1, sort=True).strip()
                    text_new_version_1 = text_new_version_1.split('投資建議')[1]
                    rate_1 = text_new_version_1.split('\n')[0].strip()
                    # 提取評價的第二種方法
                    clip_new_version_2 = fitz.Rect(75, 200, 120, 235)
                    text_new_version_1 = page.get_text(clip=clip_new_version_2, sort=True).strip()
                    rate_2 = text_new_version_1          
        return self.check_rate(rate_1, rate_2, possible_rate)
    
    def ibf(self, directory_path:str):
        '''Handle 國票 pdf

            Args :
                directory_path : (str) pdf path
            
            Return :
                rate : (str) recommend
        '''
        possible_rate = ['買進', '區間操作', '強力買進']
        rate_1, rate_2 = 'NULL'
        with fitz.open(directory_path) as doc:
            page = doc.load_page(0)
            rect = page.rect
            page_check_source = doc.load_page(-1)
            # 檢查是否由國票投顧出版
            clip_check_source = fitz.Rect(0, 0, rect.width, rect.height)
            text_check_source = page_check_source.get_text(clip=clip_check_source, sort=True)
            if '國票投顧所有' in text_check_source:
                clip_check_version = fitz.Rect(40, 0, rect.width, 400)
                text_check_version = page.get_text(clip=clip_check_version, sort=True).strip()
                if '國票觀點' in text_check_version:
                    # 報告為舊版
                    # 提取評價的第一種方法
                    clip_old_version_1 = fitz.Rect(380, 0, rect.width, 400)
                    text_old_version_1 = page.get_text(clip=clip_old_version_1, sort=True).strip()
                    if '目標價' in text_old_version_1:
                        text_old_version_1 = text_old_version_1.split('目標價')[1].strip()
                        rate_1 = text_old_version_1.split('\n')[0].strip()
                    elif '區間價位' in text_old_version_1:
                        text_old_version_1 = text_old_version_1.split('區間價位')[1].strip()
                        rate_1 = text_old_version_1.split('\n')[0].strip()
                    elif '操作區間' in text_old_version_1:
                        text_old_version_1 = text_old_version_1.split('操作區間')[1].strip()
                        rate_1 = text_old_version_1.split('\n')[0].strip()
                    elif '/買進' in text_old_version_1:
                        text_old_version_1 = text_old_version_1.split('/買進')[1].strip()
                        rate_1 = text_old_version_1.split('\n')[0].strip()
                    # 提取評價的第二種方法
                    clip_old_version_2 = fitz.Rect(380, 200, 470, 270)
                    text_old_version_2 = page.get_text(clip=clip_old_version_2, sort=True).strip()
                    if '買進' in text_old_version_2:
                        rate_2 = '買進'
                    elif '區間操作' in text_old_version_2:
                        rate_2 = '區間操作'
                    elif '賣出' in text_old_version_2:
                        rate_2 = '賣出'
                else:
                    # 報告為新版
                    # 提取評價的第一種方法
                    clip_new_version_1 = fitz.Rect(30, 200, 220, 400)
                    text_new_version_1 = page.get_text(clip=clip_new_version_1, sort=True).strip()
                    if '目標價' in text_new_version_1:
                        text_new_version_1 = text_new_version_1.split('目標價')[1].strip()
                        rate_1 = text_new_version_1.split('\n')[0].strip()  
                    elif '區間價位' in text_new_version_1:
                        text_new_version_1 = text_new_version_1.split('區間價位')[1].strip()
                        rate_1 = text_new_version_1.split('\n')[0].strip()  
                    elif '操作區間' in text_new_version_1:
                        text_new_version_1 = text_new_version_1.split('操作區間')[1].strip()
                        rate_1 = text_new_version_1.split('\n')[0].strip()  
                    # 提取評價的第二種方法
                    clip_new_version_2 = fitz.Rect(40, 200, 120, 400)
                    text_new_version_2 = page.get_text(clip=clip_new_version_2, sort=True).strip()
                    if '強力買進' in text_new_version_2:
                            rate_2 = '強力買進'
                    elif '買進' in text_new_version_2:
                        rate_2 = '買進'
                    elif '區間操作' in text_new_version_2:
                        rate_2 = '區間操作'
                    elif '賣出' in text_new_version_2:
                        rate_2 = '賣出'
        return self.check_rate(rate_1, rate_2, possible_rate)

    def ctbc(self, directory_path:str):
        '''Handle 中信託(CTBC) pdf

            Args :
                directory_path : (str) pdf path
            
            Return :
                rate : (str) recommend
        '''
        possible_rate = ['中立', '買進', '增加持股(Overweight)', '中立(Neutral)', 
                        '買進(Buy)', '增加持股', '-', '降低持股(Underweight)', '未評等']
        rate_1, rate_2 = 'NULL'
        with fitz.open(directory_path) as doc:
            page = doc.load_page(0)
            rect = page.rect
            # 檢查是否由中信投顧出版
            page_check_source = doc.load_page(-1)
            text_check_source = page_check_source.get_text()
            if '中國信託金融控股' or '中信投顧投資分析報告' in text_check_source :
                if '個股報告' in text_check_source :
                    clip_check_version= fitz.Rect(370, 80, 450, 200)
                    text_check_version = page.get_text(clip=clip_check_version, sort=True).strip()
                    if '投資評等' in text_check_version:
                        # 報告為舊版
                        # 提取評價的第一種方法
                        text_old_version_1 = text_check_version
                        text_old_version_1 = text_old_version_1.split('投資評等')[1].strip()
                        rate_1 = text_old_version_1.split('\n')[0].strip()
                        # 提取評價的第二種方法
                        clip_old_version_2 = fitz.Rect(370, 120, 430, 150)
                        rate_2 = page.get_text(clip=clip_old_version_2, sort=True).strip()
                    else:
                        # 報告為新版
                        # 提取評價的第一種方法
                        clip_new_version_1 = fitz.Rect(200, 0, rect.width, 200)
                        text_new_version_1 = page.get_text(clip=clip_new_version_1, sort=True).strip()
                        text_new_version_1 = text_new_version_1.split('評 等')[1]
                        rate_1 = text_new_version_1.split('\n')[1].strip()
                        # 提取評價的第二種方法
                        clip_new_version_2 = fitz.Rect(350, 115, 570, 200)
                        text_new_version_1 = page.get_text(clip=clip_new_version_2, sort=True).strip()
                        rate_2 = text_new_version_1.split('\n')[0].strip()
        return self.check_rate(rate_1, rate_2, possible_rate)

    def fubon(self, directory_path:str):
        '''Handle 富邦(Fubon) pdf
        
            Args :
                directory_path : (str) pdf path
            Return :
                rate : (str) recommend
        '''
        possible_rate = ['增加持股', '未評等', '中立', '買進', '降低持股', 'Buy', 'Neutral']
        rate_1, rate_2 = 'NULL'
        with fitz.open(directory_path) as doc:
            page = doc.load_page(0)
            # 檢查是否由富邦投顧出版
            clip_check_source = fitz.Rect(30, 70, 200, 140)
            text_check_source = page.get_text(clip=clip_check_source, sort=True)
            if 'Fubon' in text_check_source :
                # 提取評價的第一種方法
                clip_1 = fitz.Rect(50, 120, 200, 200)
                text_1 = page.get_text(clip=clip_1, sort=True).strip()
                rate_1 = text_1.split('\n')[0].strip()
                # 提取評價的第二種方法
                clip_2 = fitz.Rect(50, 140, 210, 170)
                rate_2 = page.get_text(clip=clip_2, sort=True).strip()
        return self.check_rate(rate_1, rate_2, possible_rate)      

    def yuanta(self, directory_path:str):
        '''Handle 元大(yuanta) pdf
        
            Args :
                directory_path : (str) pdf path
            
            Return :
                rate : (str) recommend
        '''
        possible_rate = ['持有-超越同業 (維持評等)', '買進 (維持評等)', '持有-落後同業', '持有-落後同業 (維持評等)', 
                        '買進 (調升評等)', '買進 (重新納入研究範圍)',
                        '持有-超越同業 (調降評等)', '買進 (研究員異動)', '買進  (初次報告)', 
                        '買進 (初次報告)', '持有-超越同業', '持有-落後同業(維持評等)', '賣出 (維持評等)', 
                        '持有-超越大盤(維持評等)', '持有-超越大盤 (維持評等)', '買進', '持有-落後大盤']
        rate_1, rate_2 = 'NULL'
        with fitz.open(directory_path) as doc:
            page = doc.load_page(0)
            rect = page.rect
            # 檢查是否由元大投顧出版
            page_check_source = doc.load_page(-1)
            text_check_source = page_check_source.get_text()
            if '元大證券投資顧問' in text_check_source:
                # 檢查是否為個股報告
                page_check_report = doc.load_page(0)
                clip_check_report = fitz.Rect(0, 0, rect.width, 70)
                text_check_report = page_check_report.get_text(clip=clip_check_report, sort=True).strip()
                if '更新報告' in text_check_report or '初次報告' in text_check_report:
                        # 提取評價的第一種方法
                        clip_new_version_1 = fitz.Rect(0, 0, 210, 230)
                        text_new_version_1 = page.get_text(clip=clip_new_version_1, sort=True).strip()
                        text_new_version_1 = text_new_version_1.split('目標價')[0].strip()
                        rate_1 = text_new_version_1.split('\n')[-1].strip()
                        # 提取評價的第二種方法
                        clip_new_version_2 = fitz.Rect(0, 115, 210, 145)
                        text_new_version_1 = page.get_text(clip=clip_new_version_2, sort=True).strip()
                        rate_2 = text_new_version_1.split('\n')[0].strip()
        return self.check_rate(rate_1, rate_2, possible_rate)

    def honsec(self, directory_path:str):
        '''Handle 宏遠(honsec) pdf
        
            Args :
                directory_path : (str) pdf path
            
            Return :
                rate : (str) recommend
        '''
        possible_rate = ['買進', '區間操作', '買進（調升）', '強力買進', '區間→買進',
                    '中立', '買進（維持）', '中立（調降）', '區間操作（調降）', '區間', '維持買進']
        rate_1, rate_2 = 'NULL'
        with fitz.open(directory_path) as doc:
            page = doc.load_page(0)
            rect = page.rect
            # 檢查是否由宏遠投顧出版
            clip_check_source = fitz.Rect(0, 0.9*rect.height, rect.width, rect.height)
            text_check_source = page.get_text(clip=clip_check_source, sort=True)
            if '宏遠投顧' in text_check_source :
                # 提取評價的第一種方法
                clip_new_version_1 = fitz.Rect(0, 0, 220, rect.height)
                text_new_version_1 = page.get_text(clip=clip_new_version_1, sort=True).strip()
                if '投資評等:' in text_new_version_1:
                    text_new_version_1 = text_new_version_1.split('投資評等:')[1]
                else:
                    text_new_version_1 = text_new_version_1.split('投資評等：')[1]
                rate_1 = text_new_version_1.split('\n')[0].strip()
        return self.check_rate(rate_1, rate_2, possible_rate)

    def taishin(self, directory_path:str):
        '''Handle 台新(taishin) pdf
    
            Args :
                directory_path : (str) pdf path
            
            Return :
                rate : (str) recommend
        '''
        possible_rate = ['長期持有', '中立', '買進', '強力買進', '買 進']
        rate_1, rate_2 = 'NULL'
        with fitz.open(directory_path) as doc:
            page = doc.load_page(0)
            # 檢查是否由台新投顧出版
            page_check_source = doc.load_page(-1)
            text_check_source = page_check_source.get_text()
            if '台新證券投資顧問' in text_check_source :
                # 提取評價的第一種方法
                clip_new_version_1 = fitz.Rect(0, 0, 230, 180)
                text_new_version_1 = page.get_text(clip=clip_new_version_1, sort=True).strip()
                text_new_version_1 = text_new_version_1.split('投資評等')[1].strip()
                rate_1 = text_new_version_1.split('\n')[0].strip()
                # 提取評價的第二種方法
                clip_new_version_2 = fitz.Rect(110, 85, 220, 120)
                text_new_version_1 = page.get_text(clip=clip_new_version_2, sort=True).strip()
                rate_2 = text_new_version_1.split('\n')[0].strip()
        return self.check_rate(rate_1, rate_2, possible_rate)

    def pscnet(self, directory_path:str):
        '''Handle 統一(pscnet) pdf
        
            Args :
                directory_path : (str) pdf path
            
            Return :
                rate : (str) recommend
        '''
        possible_rate = ['中立(維持評等)', '買進(調升評等)', '買進(維持評等)', '降低持股(調降評等)', '中立(調降評等)', '強力買進(調升評等)',
                        '強力買進(維持評等)', '未評等', '中立(初次評等)', '中立 (維持評等)', '強力買進(上調評等)', '買進(初次評等)', '買進 (維持評等)',
                        '買進(調降目標價)', '中立(降低評等)', '買進 (調升評等)', '買進', '賣出(調降評等)', '中立(調升評等)', '中立(下修評等)', 
                        '中立', '未評等(調整評等)', '中立 (調降評等)', '未評等(初次評等)', '強力買進(初次評等)']
        rate_1, rate_2 = 'NULL'
        with fitz.open(directory_path) as doc:
            page = doc.load_page(0)
            rect = page.rect
            # 檢查是否由統一投顧出版
            page_check_source = doc.load_page(-1)
            text_check_source = page_check_source.get_text()
            if '統一證券投資顧問' in text_check_source :
                # 檢查版本
                clip_check_new_report = fitz.Rect(rect.width/2, 0, rect.width, 150)
                text_check_new_report = page.get_text(clip=clip_check_new_report, sort=True)
                if '投資速報' in text_check_new_report :
                    # 報告為投資速報提取評價的第一種方法
                    clip_old_report_1 = fitz.Rect(0, 0, 270, 270)
                    text_old_report_1 = page.get_text(clip=clip_old_report_1, sort=True).strip()
                    text_old_report_1 = text_old_report_1.split('出刊緣由')[0].strip()
                    rate_1 = text_old_report_1.split('\n')[0].strip()
                    # 報告為舊版個股報告 提取評價的第二種方法
                    clip_old_report_2 = fitz.Rect(70, 130, 265, 160)
                    rate_2 = page.get_text(clip=clip_old_report_2, sort=True).strip()
                elif '訪談報告' in text_check_new_report :
                    # 報告為訪談報告提取評價的第一種方法
                    clip_new_report_1 = fitz.Rect(375, 120, 565, 190)
                    text_new_report_1 = page.get_text(clip=clip_new_report_1, sort=True).strip()
                    text_new_report_1 = text_new_report_1.split('出刊緣由')[0].strip()
                    rate_1 = text_new_report_1.split('\n')[0].strip()
                    # 報告為訪談報提取評價的第二種方法
                    clip_new_report_2 = fitz.Rect(375, 120, 565, 160)
                    rate_2 = page.get_text(clip=clip_new_report_2, sort=True).strip()
                elif '初次報告' in text_check_new_report :
                    # 報告為初次報告提取評價的第一種方法
                    clip_new_report_1 = fitz.Rect(365, 190, 565, 380)
                    text_new_report_1 = page.get_text(clip=clip_new_report_1, sort=True).strip()
                    text_new_report_1 = text_new_report_1.split('目標價')[0].strip()
                    rate_1 = text_new_report_1.split('\n')[0].strip()
                    # 報告為初次報告提取評價的第二種方法
                    clip_new_report_2 = fitz.Rect(365, 190, 565, 220)
                    rate_2 = page.get_text(clip=clip_new_report_2, sort=True).strip()
        return self.check_rate(rate_1, rate_2, possible_rate)

    def capital(self, directory_path:str):
        '''Handle 群益(capital) pdf
        
            Args :
                directory_path : (str) pdf path
            
            Return :
                rate : (str) recommend
        '''
        possible_rate = ['中立', '中立', '區間操作', '買進']
        rate_1, rate_2 = 'NULL'
        with fitz.open(directory_path) as doc:
            page = doc.load_page(0)
            rect = page.rect
            # 檢查是否由群益投顧出版
            clip_check_source = fitz.Rect(0, 0.9*rect.height, rect.width, rect.height)
            text_check_source = page.get_text(clip=clip_check_source, sort=True)
            if '群益投顧' in text_check_source :
                clip_check_new_report = fitz.Rect(rect.width/2, 0, rect.width, 70)
                text_check_new_report = page.get_text(clip=clip_check_new_report, sort=True)
                if '個股報告' in text_check_new_report :
                    # 報告為新版個股報告 提取評價的第一種方法
                    clip_new_report_1 = fitz.Rect(190, 105, rect.width, 160)
                    text_new_report_1 = page.get_text(clip=clip_new_report_1, sort=True).strip()
                    text_new_report_1 = text_new_report_1.split(')')[1].strip()
                    rate_1 = text_new_report_1.split('\n')[0].strip()
                    # 報告為新版個股報告 提取評價的第二種方法
                    clip_new_report_2 = fitz.Rect(400, 105, 565, 160)
                    rate_2 = page.get_text(clip=clip_new_report_2, sort=True).strip()
        rate = self.check_rate(rate_1, rate_2, possible_rate)
        # '立'因unicode不同有時會造成無法壓縮的錯誤
        if '立' in rate:
            rate.replace('立', '立')
        return rate

    def masterlink(self, directory_path:str):
        '''Handle 元富(masterlink) pdf
        
            Args :
                directory_path : (str) pdf path
            
            Return :
                rate : (str) recommend
        '''
        rate_1, rate_2 = 'NULL'
        possible_rate = ['維持買進', 'BUY', '評等中立', '維持中立', 'HOLD', '中立轉買進', 
                        'Upgrade to BUY', '買進轉中立', 'Downgrade to HOLD', '評等買進', 
                        'UPGRADE TO BUY', 'Upgrade To BUY', '買進轉強力買進', '維持強力買進', 
                        'STRONG BUY', 'Upgarde to BUY']
        with fitz.open(directory_path) as doc:
            page = doc.load_page(0)  # 加載文檔的第0頁
            rect = page.rect
            # 檢查是否由元富投顧出版
            clip_check_source = fitz.Rect(0, rect.height/2, rect.width, rect.height)
            text_check_source = page.get_text(clip=clip_check_source, sort=True)
            if '本刊載之報告為元富投顧於特定日期之分析' in text_check_source or 'MasterLink Research reports' in text_check_source :
                # 檢查版本
                clip_check_new_report = fitz.Rect(rect.width/2, 0, rect.width, 70)
                text_check_new_report = page.get_text(clip=clip_check_new_report, sort=True)
                clip_check_old_fast_report = fitz.Rect(rect.width/2, 50, rect.width, 150)
                text_check_old_report = text_check_new_report
                text_check_old_fast_report = page.get_text(clip=clip_check_old_fast_report, sort=True)
                if '公司訪談報告' in text_check_old_report or '評等調整' in text_check_old_report or '訪談報告' in text_check_old_report :
                    # 報告為舊版個股報告 提取評價的第一種方法
                    clip_old_report_1 = fitz.Rect(0, 0, rect.width, 120)
                    text_old_report_1 = page.get_text(clip=clip_old_report_1, sort=True).strip()
                    text_old_report_1 = text_old_report_1.split(')')[1].strip()
                    rate_1 = text_old_report_1.split('\n')[0].strip()
                    # 報告為舊版個股報告 提取評價的第二種方法
                    clip_old_report_2 = fitz.Rect(250, 60, rect.width, 110)
                    rate_2 = page.get_text(clip=clip_old_report_2, sort=True).strip()
                elif '公司拜訪快報' in text_check_old_fast_report :
                    # 報告為舊版公司拜訪快報 提取評價的第一種方法
                    clip_old_fast_report_1 = fitz.Rect(0, 0, rect.width, 200)
                    text_old_fast_report_1 = page.get_text(clip=clip_old_fast_report_1, sort=True).strip()
                    text_old_fast_report_1 = text_old_fast_report_1.split('建議：')[1].strip()
                    rate_1 = text_old_fast_report_1.split('\n')[0].strip()
                    # 報告為舊版公司拜訪快報 提取評價的第二種方法
                    clip_old_fast_report_2 = fitz.Rect(70, 140, rect.width, 170)
                    rate_2 = page.get_text(clip=clip_old_fast_report_2, sort=True).strip()
                elif '公司拜訪報告' in text_check_new_report or '個股報告' in text_check_new_report or '評等調升報告' in text_check_new_report:
                    # 報告為新版個股報告 提取評價的第一種方法
                    clip_new_report_1 = fitz.Rect(0, 0, rect.width, 120)
                    text_new_report_1 = page.get_text(clip=clip_new_report_1, sort=True).strip()
                    text_new_report_1 = text_new_report_1.split(')')[1].strip()
                    rate_1 = text_new_report_1.split('\n')[0].strip()
                    # 報告為新版個股報告 提取評價的第二種方法
                    clip_new_report_2 = fitz.Rect(250, 60, rect.width, 110)
                    rate_2 = page.get_text(clip=clip_new_report_2, sort=True).strip()
                elif '公司拜訪快報' in text_check_new_report :
                    # 報告為新版公司拜訪快報 提取評價的第一種方法
                    clip_new_fast_report_1 = fitz.Rect(0, 0, rect.width, 150)
                    text_new_fast_report_1 = page.get_text(clip=clip_new_fast_report_1, sort=True).strip()
                    text_new_fast_report_1 = text_new_fast_report_1.split('建議：')[1].strip()
                    rate_1 = text_new_fast_report_1.split('\n')[0].strip()
                    # 報告為新版公司拜訪快報 提取評價的第二種方法
                    clip_new_fast_report_2 = fitz.Rect(80, 105, rect.width, 140)
                    rate_2 = page.get_text(clip=clip_new_fast_report_2, sort=True).strip()
                elif 'Company Report' in text_check_new_report :
                    # 報告為英文版個股報告 提取評價的第一種方法
                    clip_english_report_1 = fitz.Rect(0, 0, rect.width, 100)
                    text_english_report_1 = page.get_text(clip=clip_english_report_1, sort=True).strip()
                    text_english_report_1 = text_english_report_1.split(')')[-1].strip()
                    rate_1 = text_english_report_1.split('\n')[0].strip()
                    # 報告為新版個股報告 提取評價的第二種方法
                    clip_english_report_2 = fitz.Rect(320, 60, rect.width, 110)
                    rate_2 = page.get_text(clip=clip_english_report_2, sort=True).strip()
        return self.check_rate(rate_1, rate_2, possible_rate)

    def ffhc(self, directory_path:str):
        '''Handle 第一金(First Financial Holding Co.) pdf
        
            Args :
                directory_path : (str) pdf path
            
            Return :
                rate : (str) recommend
        '''
        rate_1, rate_2 = 'NULL', 'NULL'
        possible_rate = ['買進', '區間操作', '中立', 'Trading Buy', '區間', '強力買進', '- -', 'buy', 'Buy', 'Neutral']
        with fitz.open(directory_path) as doc:
            page = doc.load_page(0)
            rect = page.rect
            page_check_source = doc.load_page(-1)
            text_check_source = page_check_source.get_text()
            if '第一金證券投資顧問' in text_check_source :
                # 檢查版本                
                clip_check_report = fitz.Rect(rect.width/2, 0, rect.width, 150)
                text_check_report = page.get_text(clip=clip_check_report)
                check_new_report = ['個股報告', '動態更新', '近況更新', '新股報告', '新 股 報 告', '新 股 掛 牌', '新股掛牌', '新 股 報 告 ', '個股速報']
                if any(keyword in text_check_report for keyword in check_new_report):
                    # 提取評價的第一種方法
                    clip_new_report_1 = fitz.Rect(0, 130, 170, 300)
                    text_new_report_1 = page.get_text(clip=clip_new_report_1).strip()
                    rate_1 = text_new_report_1.split('\n')[0].strip()
                    # 提取評價的第二種方法
                    clip_new_report_2 = fitz.Rect(0, 140, 170, 300)
                    rate_2 = page.get_text(clip=clip_new_report_2).strip()
                    for keyword in possible_rate:
                        if keyword in rate_2:
                            rate_2 = keyword
                elif '個 股 速 報' in text_check_report :
                    # 提取評價的第一種方法
                    clip_old_report_1 = fitz.Rect(0, 130, 170, 300)
                    text_old_report_1 = page.get_text(clip=clip_old_report_1).strip()
                    rate_1 = text_old_report_1.split('\n')[0].strip()
                    # 提取評價的第二種方法
                    clip_old_report_2 = fitz.Rect(0, 140, 170, 300)
                    rate_2 = page.get_text(clip=clip_old_report_2).strip()
                    for keyword in possible_rate:
                        if keyword in rate_2:
                            rate_2 = keyword
        return self.check_rate(rate_1, rate_2, possible_rate)

    def jihsun(self, directory_path:str):
        '''Handle 日昇(JihSun) pdf
        
            Args :
                directory_path : (str) pdf path
            
            Return :
                rate : (str) recommend
        '''
        rate_1, rate_2 = 'NULL', 'NULL'
        possible_rate = ['買進', '持有', '中立', '買進買進', '中立中立']
        with fitz.open(directory_path) as doc:
            page = doc.load_page(0)
            rect = page.rect
            page_check_source = doc.load_page(-1)
            text_check_source = page_check_source.get_text()
            if '日盛證券投資顧問' in text_check_source :
                # 檢查版本                
                clip_check_report = fitz.Rect(rect.width/2, 0, rect.width, 150)
                text_check_report = page.get_text(clip=clip_check_report)
                if any(keyword in text_check_report for keyword in ['訪談報告', '公司速報']):
                    # 提取評價的第一種方法
                    clip_old_report_1 = fitz.Rect(100, 0, 165, 200)
                    text_old_report_1 = page.get_text(clip=clip_old_report_1, sort=True).strip()
                    try:
                        text_old_report_1 = text_old_report_1.split('投資評等')[1].strip()
                        rate_1 = text_old_report_1.split('\n')[0].strip()
                    except:
                        rate_1 = 'NULL'
                    # 提取評價的第二種方法
                    clip_old_report_2 = fitz.Rect(100, 120, 165, 145)
                    rate_2 = page.get_text(clip=clip_old_report_2).strip()
        return self.check_rate(rate_1, rate_2, possible_rate)

    def esun(self, directory_path:str):
        '''Handle 玉山(ESun) pdf
        
            Args :
                directory_path : (str) pdf path
            
            Return :
                rate : (str) recommend
        '''
        rate_1, rate_2 = 'NULL', 'NULL'
        possible_rate = ['買進(維持)', '逢低買進(維持)', '買進(初次)']
        with fitz.open(directory_path) as doc:
            page = doc.load_page(0)
            rect = page.rect
            page_check_source = doc.load_page(-1)
            text_check_source = page_check_source.get_text()
            if '玉山證券投資顧問' in text_check_source :
                # 檢查版本                
                clip_check_report = fitz.Rect(0, 0, rect.width, 100)
                text_check_report = page.get_text(clip=clip_check_report)
                if any(keyword in text_check_report for keyword in ['研 究 報 告']):
                    # 提取評價的第一種方法
                    clip_old_report_1 = fitz.Rect(220, 80, rect.width, 180)
                    text_old_report_1 = page.get_text(clip=clip_old_report_1, sort=True).strip()
                    try:
                        text_old_report_1 = text_old_report_1.split('TP')[0].strip()
                        rate_1 = text_old_report_1.split('\n')[-1].strip()
                    except:
                        rate_1 = 'NILL'
                    # 提取評價的第二種方法
                    clip_old_report_2 = fitz.Rect(230, 120, 430, 160)
                    rate_2 = page.get_text(clip=clip_old_report_2).strip()
        return self.check_rate(rate_1, rate_2, possible_rate)

    def cathay(self, directory_path:str):
        '''Handle 國泰(Cathay) pdf
        
            Args :
                directory_path : (str) pdf path
            
            Return :
                rate : (str) recommend
        '''
        rate_1, rate_2 = 'NULL', 'NULL'
        possible_rate = ['中立 – 維持中立', '買進 – 維持買進', '買進', '中立 – 初次評等中立', '中立 – 買進轉中立', '中立 – 初次評等', '買進– 維持買進']
        with fitz.open(directory_path) as doc:   
            page = doc.load_page(0)
            rect = page.rect
            page_check_source = doc.load_page(-1)
            text_check_source = page_check_source.get_text()
            if '國泰金融控股公司' in text_check_source :
                # 檢查版本                
                clip_check_report = fitz.Rect(rect.width/2, 0, rect.width, 150)
                text_check_report = page.get_text(clip=clip_check_report)
                if any(keyword in text_check_report for keyword in ['個股報告']):
                    # 提取評價的第一種方法
                    clip_old_report_1 = fitz.Rect(350, 80, rect.width, 160)
                    text_old_report_1 = page.get_text(clip=clip_old_report_1, sort=True).strip()
                    rate_1 = text_old_report_1.split('\n')[0].strip()
                    # 提取評價的第二種方法
                    clip_old_report_2 = fitz.Rect(350, 80, 550, 130)
                    rate_2 = page.get_text(clip=clip_old_report_2).strip()
        return self.check_rate(rate_1, rate_2, possible_rate)

    def mega(self, directory_path:str):
        '''Handle 兆豐(Mega) pdf
        
            Args :
                directory_path : (str) pdf path
            
            Return :
                rate : (str) recommend
        '''
        rate_1, rate_2 = 'NULL', 'NULL'
        possible_rate = ['逢低買進', '區間操作', '買進']
        with fitz.open(directory_path) as doc:
            page = doc.load_page(0)
            rect = page.rect
            page_check_source = doc.load_page(-1)
            text_check_source = page_check_source.get_text()
            if '本刊所刊載之內容僅做為參考，惟已力求正確與完整' in text_check_source :
                # 檢查版本                
                clip_check_new_report = fitz.Rect(rect.width/2, 0, rect.width, 150)
                text_check_new_report = page.get_text(clip=clip_check_new_report)
                if '訪談速報' in text_check_new_report or '個股報告' in text_check_new_report:
                    # 提取評價的第一種方法
                    clip_old_report_1 = fitz.Rect(400, 0, rect.width, 200)
                    text_old_report_1 = page.get_text(clip=clip_old_report_1).strip()
                    try:
                        text_old_report_1 = text_old_report_1.split('目標價')[1].strip()
                        rate_1 = text_old_report_1.split('\n')[0].strip()
                    except:
                        rate_1 = 'NULL'
                # 提取評價的第二種方法
                clip_old_report_2 = fitz.Rect(410, 140, 490, 165)
                rate_2 = page.get_text(clip=clip_old_report_2).strip()
            return self.check_rate(rate_1, rate_2, possible_rate)

    def gfortune(self, directory_path:str):
        '''Handle 福邦(Grand Fortune) pdf
        
            Args :
                directory_path : (str) pdf path
            
            Return :
                rate : (str) recommend
        '''
        rate_1, rate_2 = 'NULL', 'NULL'
        possible_rate = ['中立', '優於大盤']
        with fitz.open(directory_path) as doc:
            page = doc.load_page(0)
            rect = page.rect
            page_check_source = doc.load_page(-1)
            text_check_source = page_check_source.get_text()
            if 'Grand Fortune Securities' in text_check_source :
                # 檢查版本                
                clip_check_report = fitz.Rect(0, 0, rect.width, 100)
                text_check_report = page.get_text(clip=clip_check_report)
                if any(keyword in text_check_report for keyword in ['股市個股早報']):
                    # 提取評價的第一種方法
                    clip_old_report_1 = fitz.Rect(0, 0, 200, 300)
                    text_old_report_1 = page.get_text(clip=clip_old_report_1, sort=True).strip()
                    try:
                        text_old_report_1 = text_old_report_1.split('投資評等')[1].strip()
                        rate_1 = text_old_report_1.split('\n')[0].strip()
                    except:
                        rate_1 = 'NULL'
                    # 提取評價的第二種方法
                    clip_old_report_2 = fitz.Rect(30, 200, 220, 250)
                    rate_2 = page.get_text(clip=clip_old_report_2).strip()
        return self.check_rate(rate_1, rate_2, possible_rate)

    def concords(self, directory_path:str):
        '''Handle 康和(Concord Securities) pdf
        
            Args :
                directory_path : (str) pdf path
            
            Return :
                rate : (str) recommend
        '''
        rate_1, rate_2 = 'NULL', 'NULL'
        possible_rate = ['買進 (維持評等)', '中性', '中性 (維持評等)', '買進(維持評等)', '中 性 (維 持 評 等 )', '中性 (調降評等)',
                    '買進 (調升評等)', '買進(首次評等)', '未評等', '未 評 等', '買進 (首次評等)', '新股掛牌 (未評等)', '新 股 掛 牌',
                    '買進-維持', '逢低買進-首次', '逢低買進', '逢低買進-維持', '買進-首次', '買進']
        with fitz.open(directory_path) as doc:
            page = doc.load_page(0)
            rect = page.rect
            page_check_source = doc.load_page(-1)
            text_check_source = page_check_source.get_text()
            if '康 和 投 資 顧 問' in text_check_source :
                # 檢查版本                
                clip_check_report = fitz.Rect(rect.width/2, 0, rect.width, 150)
                text_check_report = page.get_text(clip=clip_check_report)
                if any(keyword in text_check_report for keyword in ['個股報告', '個 股 報 告', '投  資  速  報', '投 資 速 報']):
                    # 提取評價的第一種方法
                    clip_old_report_1 = fitz.Rect(0, 100, 400, 160)
                    text_old_report_1 = page.get_text(clip=clip_old_report_1, sort=True).strip()
                    if '，' in text_old_report_1:
                        rate_1 = text_old_report_1.split('，')[0].strip()
                    elif '目標價' in text_old_report_1:
                        rate_1 = text_old_report_1.split('目標價')[0].strip()
                    else:
                        rate_1 = text_old_report_1.split('\n')[0].strip()
        return self.check_rate(rate_1, rate_2, possible_rate)