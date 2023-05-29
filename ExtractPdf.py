from abc import abstractmethod

class ExtractPdf(object):
    '''Extracts pdf's informations'''
    def __init__(self, file_path:str):
        '''
        Args :
            file_path : (str) target file's full path
            rating_1 : (str) rating extracted by method 1
            rating_2 : (str) rating extracted by method 2
            possible_rating : (list) all possible ratings in the adviser

        Todo:
            fix get tp, author process
        '''
        self.file_path = file_path
        self.rating_1 = 'NULL'
        self.rating_2 = 'NULL'
        self.possible_rating = list()
        self.advisor = 'Unknown'
        self.tp_1 = 'NULL'
        self.tp_2 = 'NULL'
        self.author_1 = 'NULL'
        self.author_2 = 'NULL'
        self.summary_1 = 'NULL'
        self.summary_2 = 'NULL'

    def check_rating(self, rating_1, rating_2, possible_rating):
        '''Check that the extracted ratings are correct

            Return :
                rating : (str) recommend rating
        '''
        if rating_1 == rating_2:
            return rating_1
        for rating in possible_rating:
            if rating == rating_1:
                return rating
            elif rating == rating_2:
                return rating
        return 'NULL'
    
    def check_targrt_price(self, tp_1, tp_2):
        '''Check that the extracted target_price are correct

            Return :
                tp : (str) recommend target price
        '''
        if tp_1 == tp_2:
            return tp_1
        elif tp_1 != 'NULL' and tp_2 == 'NULL':
            return tp_1
        elif tp_2 != 'NULL' and tp_1 == 'NULL':
            return tp_2
        else:
            return 'NULL'
        # else:
        #     if '.' in tp_1:
        #         try:
        #             int = int(tp_1.split('.')[0])
        #             int = int(tp_1.split('.')[1])
        #             return tp_1
        #         except:
        #             pass
        #     if '.' in tp_2:
        #         try:
        #             int = int(tp_2.split('.')[0])
        #             int = int(tp_2.split('.')[1])
        #             return tp_2
        #         except:
        #             pass
        #     return 'NULL'

    def check_author(self, author_1, author_2):
        '''Check that the extracted target_price are correct

            Return :
                tp : (str) recommend target price
        '''
        if author_1 == author_2:
            return author_1
        elif author_1 != 'NULL' and author_2 == 'NULL':
            return author_1
        elif author_2 != 'NULL' and author_1 == 'NULL':
            return author_2
        else:
            return 'NULL'
        # else:
        #     if len(author_1)>0 and len(author_1)<4:
        #         return author_1
        #     elif len(author_2)>0 and len(author_2)<4:
        #         return author_2
        #     else:
        #         return 'NULL'

    def check_summary(self, summary_1, summary_2):
        '''Check that the extracted target_price are correct

            Return :
                tp : (str) recommend target price
        '''
        if summary_1 == summary_2:
            return summary_1
        elif summary_1 != 'NULL' and summary_2 == 'NULL':
            return summary_1
        elif summary_2 != 'NULL' and summary_1 == 'NULL':
            return summary_2
        else:
            return 'NULL'   
    
    
class Adviser(ExtractPdf):
    '''Extracts pdf's adviser'''
    def __init__(self, file_path):
        super().__init__(file_path)

    @abstractmethod
    def get_rating_process(self):
        raise NotImplementedError("get_rating_process not implement")

    @abstractmethod
    def get_target_price_process(self):
        raise NotImplementedError("get_target_price_process not implement")

    @abstractmethod
    def get_author_process(self):
        raise NotImplementedError("get_author_process not implement")

    @abstractmethod
    def get_summary_process(self):
        raise NotImplementedError("get_summary_process not implement")

if __name__ == '__main__' :
    pass