from selenium import webdriver
import pandas as pd
import time
from google_sheet_handler import Google_sheet_handler
import logger_hander

class Send_whatsapp_msg:

    def get_connection(self):
        self.driver = webdriver.Chrome("./chromedriver")
        self.driver.get("https://web.whatsapp.com")
        print("Scan QR Code, And then Enter")
        input()
        logger.info("Whatsapp Logged In..!")
        return self.driver

    def send_msg(self,name, message):
        print("okkkkk",name,message)
        try:
            self.driver.find_element_by_xpath('//span[@title = "{}"]'.format(name)).click()
            self.text_box = self.driver.find_element_by_xpath('//*[@id="main"]/footer/div[1]/div[2]/div/div[2]')
            self.text_box.send_keys(message)
            self.driver.find_element_by_xpath('//*[@id="main"]/footer/div[1]/div[3]/button/span').click()
            time.sleep(10)
            logger.info("Send message to {}".format(name))
        except Exception as _:
            logger.error("{} User/ group not found...! try again".format(name))

    def check_cell_name_valid_or_not(self, sheet, List_cell_name):
        return Google_sheet_handler.find_cell(self, sheet, List_cell_name)

    def extract_data(self,sheet):
        List_of_cell_name = ['Name'	,'Track Score',	'Overall Comments']
        Name_list = []
        Track_score_list = []
        Overall_comments_list = []
        flag = self.check_cell_name_valid_or_not(sheet, List_of_cell_name)
        if flag:
            list_of_records = sheet.get_all_records()
            if len(list_of_records) != 0:
                for records in list_of_records:
                    Name_list.append(records.get('Name'))
                    Track_score_list.append(records.get('Track Score'))
                    Overall_comments_list.append(records.get('Overall Comments'))
                return Name_list, Track_score_list, Overall_comments_list
            else:
                logger.error("Sheet Records not available..!")
        else:
            logger.error("Sheet Header name Or Cell Name not valid..!")

    def send_messages(self, msg_dataframe):
        try:
            for ind in msg_dataframe.index:
                name = msg_dataframe['Name_list'][ind]
                score = str(msg_dataframe['Track_score_list'][ind])
                feedback = msg_dataframe['Overall_comments_list'][ind]
                message = "Hello {name}, your score is  {score} & Feedback is : {feedback}".format(name=name, score=score,feedback=feedback)
                self.send_msg(name,message)
            return True
        except Exception as e:
            excepName = type(e).__name__
            logger.error("Something Went Wrong while iterating Dataframe.. ! {}".format(excepName))
            return False

if __name__ == "__main__":
    sheet_handler = Google_sheet_handler()
    logger = logger_hander.set_logger()
    whatsapp_obj = Send_whatsapp_msg()

    sheet = sheet_handler.call_sheet("Engineers Details", "whatsapp_demo")
    if sheet != 'WorksheetNotFound':
        logger.info("got Sheet Handler..!")
        Name_list, Track_score_list, Overall_comments_list = whatsapp_obj.extract_data(sheet)
        dict = {'Name_list': Name_list, 'Track_score_list': Track_score_list,'Overall_comments_list': Overall_comments_list}
        msg_dataframe = pd.DataFrame(dict)
        driver = whatsapp_obj.get_connection()
        time.sleep(10)
        send_msg_flag = whatsapp_obj.send_messages(msg_dataframe)
        if send_msg_flag:
            driver.quit()
            logger.info("Whatsapp msg send successfully..!")
        else:
            logger.error("whatsapp msg Failed..!")

