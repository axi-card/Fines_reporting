import pandas as pd
pd.set_option('display.max_columns',500)
import ezodf
import datetime
from numpy import where
import os
import sys
from mailing import Mailing
from file_transfer import SFTP

class Reporting:

    def __init__ (self):
        """This is an obligatory form of dataframe"""
        self.template = pd.DataFrame(columns=['Proposal date','Decision date','Post date','Activation date','Limit','Status','Substatus','Comment','Comment date'])

        self.emsg = Mailing()


    def card_proposals_preparation(self):
        """Shape Card Proposals dataframe for further treatment"""
        try:
            self.card_proposals_df = pd.read_excel("DOWNLOADS/Proposals/gvProposals.xls")

            if not pd.Series(['Date of Proposal ','Date Credit Dept. ','PESEL','Limit','Names','Status']).isin(self.card_proposals_df.columns).all():
                raise Exception

        except FileNotFoundError:
            self.emsg.send_critical_message("There is no DOWNLOADS/Proposals/gvProposals.xls directory.\nReport preparation has stopped.")

            print("There is no DOWNLOADS/Proposals/gvProposals.xls directory.\nReport preparation has stopped.")
            sys.exit()

        except Exception:
            self.emsg.send_critical_message("gvProposals.xls file does not contain all necessary columns.\nReport preparation has stopped.")
            print("gvProposals.xls file does not contain all necessary columns.\nReport preparation has stopped.")
            sys.exit()


        #Limit records



        self.card_proposals_df = self.card_proposals_df[-self.card_proposals_df["Status"].isin(["Approved"])]

        #Rename and shape



        self.card_proposals_df = self.card_proposals_df.rename(columns={"Date of Proposal ": "Proposal date", "Date Credit Dept. ": "Decision date","Names": "Customer"})



        self.card_proposals_df['Post date'] = ""

        self.card_proposals_df['Activation date'] = ""

        self.card_proposals_df['Substatus'] = ""

        self.card_proposals_df['Comment'] = ""

        self.card_proposals_df['Comment date'] = ""

        #Order

        self.card_proposals_df = self.card_proposals_df[['Proposal date', 'Decision date', 'Post date', 'Activation date', 'PESEL','Phone', 'Limit', 'Status','Substatus', 'Comment','Comment date']]

    def credit_cards_preparation(self):

        try:
            self.credit_cards_df = pd.read_excel("DOWNLOADS/CreditCards/ASPxGridViewCreditCards.xls")

            if not pd.Series(['Date of Proposal','Approval Date','CID','Customer','Limit','Phone','Status']).isin(self.credit_cards_df.columns).all():
                raise Exception

        except FileNotFoundError:
            msg = "There is no DOWNLOADS/CreditCards/ASPxGridViewCreditCards.xls directory.\nReport preparation has stopped."
            self.emsg.send_critical_message(msg)
            print(msg)
            sys.exit()

        except Exception:
            msg = "ASPxGridViewCreditCards.xls file does not contain all necessary columns.\nReport preparation has stopped."
            self.emsg.send_critical_message(msg)
            print(msg)
            sys.exit()


        #Limit records


        self.credit_cards_df = self.credit_cards_df[self.credit_cards_df['Status'].isin(['With signed contract', 'Approved '])]

        # Prepare other files

        self.reports_cards_df = pd.read_excel("DOWNLOADS/ReportsCards/ASPxGridViewCards.xls")

        self.raport_do_cc_prep()


        try:

            self.processing_df = pd.read_excel("DOWNLOADS/Processing/processing.xlsx")

            if not pd.Series(['Komentarz','PESEL','Substatus','Data komentarza']).isin(self.processing_df.columns).all():
                raise Exception

        except FileNotFoundError:
            msg = "There is no DOWNLOADS/Processing/processing.xlsx directory.\nReport preparation has stopped."
            self.emsg.send_critical_message(msg)
            print(msg)
            sys.exit()

        except Exception:
            msg = "processing.xlsx file does not contain all necessary columns.\nReport preparation has stopped."
            self.emsg.send_critical_message(msg)
            print(msg)
            sys.exit()


        self.card_proposals_temp_df = pd.read_excel("DOWNLOADS/Proposals/gvProposals.xls")




        # Merge files


        self.credit_cards_df = self.credit_cards_df.merge(self.reports_cards_df[['PESEL', 'Date of Activation', 'CID']], on='CID',how='left')

        self.credit_cards_df = self.credit_cards_df.merge(self.raport_do_cc_df[['CID', 'Post date', 'Comments']], on='CID', how='left')

        self.credit_cards_df = self.credit_cards_df.merge(self.processing_df[['Komentarz', 'PESEL', 'Substatus','Data komentarza']], on='PESEL',how='left')

        self.credit_cards_df = self.credit_cards_df.merge(self.card_proposals_temp_df[['Phone', 'PESEL']], on='Phone', how='left')

        self.credit_cards_df['PESEL'] = where(self.credit_cards_df['PESEL_x'].isnull(), self.credit_cards_df['PESEL_y'],self.credit_cards_df['PESEL_x'])

        #Combine columns

        self.credit_cards_df['Comments'] = self.credit_cards_df['Comments'].fillna("")

        self.credit_cards_df['Komentarz'] = self.credit_cards_df['Komentarz'].fillna("")

        self.credit_cards_df['Komentarz'] = where(self.credit_cards_df['Komentarz'] == "", self.credit_cards_df['Comments'],self.credit_cards_df['Komentarz'] + " " + self.credit_cards_df['Comments'])

        # Rename and shape

        self.credit_cards_df["ID"] = ""

        self.credit_cards_df = self.credit_cards_df.rename(columns={"Approval Date": "Decision date", "Date of Activation": "Activation date","Date of Proposal":"Proposal date", "Komentarz":"Comment",'Data komentarza':'Comment date'})



        # Order

        self.credit_cards_df = self.credit_cards_df[['Proposal date', 'Decision date', 'Post date', 'Activation date', 'PESEL','Phone', 'Limit', 'Status', 'Substatus', 'Comment', 'Comment date']]



    def raport_do_cc_prep(self):

        try:
            self.raport_do_cc_df = self.read_ods("DOWNLOADS/Raport do CC/Raport do CC NEW.ods", 0)

            if not pd.Series(['CID','Post date','Comments']).isin(self.raport_do_cc_df.columns).all():
                raise Exception

        except FileNotFoundError:
            msg = "There is no DOWNLOADS/Raport do CC/Raport do CC NEW.ods directory.\nReport preparation has stopped."
            self.emsg.send_critical_message(msg)
            print(msg)
            sys.exit()

        except Exception:
            msg = "Raport do CC NEW.ods file does not contain all necessary columns.\nReport preparation has stopped."
            self.emsg.send_critical_message(msg)
            print(msg)
            sys.exit()



        self.raport_do_cc_df['CID'] = self.raport_do_cc_df['CID'].fillna(0)

        self.raport_do_cc_df['CID'] = self.raport_do_cc_df['CID'].astype('int')



    def read_ods(self,filename, sheet_no=0, header=0):
        tab = ezodf.opendoc(filename=filename).sheets[sheet_no]

        return pd.DataFrame({col[header].value: [x.value for x in col[header + 1:]] for col in tab.columns()})




    def concat_long_form(self):
        """Concatenate newly downloaded long form with former one"""



        self.long_form = pd.read_excel("DOWNLOADS/Dealzilla/long_form.xlsx")

        self.long_form_full = pd.read_excel("DOWNLOADS/Dealzilla/long_form_full.xlsx")

        self.long_form_full = pd.concat([self.long_form,self.long_form_full], ignore_index=True, sort=False)

        self.long_form_full = self.long_form_full.drop_duplicates(subset='id',keep='first')

        self.long_form_full['PESEL'] = self.long_form_full['identificationNumber']



        self.long_form_full['id'] = self.long_form_full['id'].replace(46885, 74078)

        self.long_form_full['id'] = self.long_form_full['id'].replace(46879, 74074)


        # remove older duplicates

        self.long_form_full = self.long_form_full.sort_values(by='id',ascending=False)

        self.long_form_full = self.long_form_full.drop_duplicates(subset=['PESEL'])

        self.long_form_full.to_excel("DOWNLOADS/Dealzilla/long_form_full.xlsx", index=False)


    def concat_dataframes(self):
        """Concatenate all dataframes and establish final order of the report"""

        self.long_form_full = pd.read_excel("DOWNLOADS/Dealzilla/long_form_full.xlsx",)

        self.final_df  = pd.concat([self.template, self.credit_cards_df, self.card_proposals_df], ignore_index=True, sort=False)

        #Edited - get Proposal date from Long Form (create_time)

        self.final_df = self.final_df.merge(self.long_form_full[['PESEL','id','create_time']], on='PESEL',how='left')

        self.final_df = self.final_df.drop_duplicates(subset=(['id','Proposal date','Status']))

        self.final_df = self.final_df.sort_values(by=['Proposal date'], ascending=False)

        self.final_df = self.final_df.drop_duplicates(subset=(['id']))


        self.final_df['Proposal date'] = self.final_df['create_time']

        self.final_df = self.final_df.dropna(subset=['id'])

        self.final_df['Proposal date'] = self.final_df['Proposal date'].map(self.correct_date_long_form)

        #Change to datetime format

        self.final_df['Post date'] = self.final_df['Post date'].astype('datetime64[ns]')

        try:
            self.final_df['Proposal date'] = self.final_df['Proposal date'].astype('datetime64[ns]')
        except:
            pass

        self.final_df['Status'] = self.final_df['Status'].str.strip()

        self.final_df = self.final_df.sort_values(by=['Proposal date'], ascending=False)

        # Order and remove \n

        self.final_df['Comment'] = self.final_df['Comment'].str.replace("\n", " ")


        # These ids will not be in the report

        self.final_df = self.final_df[-self.final_df['id'].isin([['44494','44495','44496','44497','44498','44520','44512','44521','45303']])]


        self.final_df = self.final_df[['id','Proposal date', 'Decision date', 'Post date', 'Activation date', 'Limit', 'Status','Substatus', 'Comment', 'Comment date', 'PESEL']]

    def create_file(self):
        """Write prepared file as new report"""

        try:
            os.remove("DOWNLOADS/Compare/old.xlsx")
        except:
            print("COULD NOT FIND OLD.XLSX")

        try:
            os.rename("DOWNLOADS/Compare/new.xlsx","DOWNLOADS/Compare/old.xlsx")
        except:
            print("COULD NOT FIND NEW.XLSX\nPYTHON SCRIP HAS TO STOP\nPRESS ENTER")
            sys.quit()

        self.final_df.to_excel("DOWNLOADS/Compare/new.xlsx", index=False)

    def compare_files(self):
        """Compare both files and return differences report"""

        try:
            old_df = pd.read_excel("DOWNLOADS/Compare/old.xlsx")
        except Exception as e:

            print(e)
            self.emsg.send_critical_message(e)


        try:
            new_df = pd.read_excel("DOWNLOADS/Compare/new.xlsx")

        except Exception as e:
            
            print(e)
            self.emsg.send_critical_message(e)



        new_df = new_df.merge(old_df[['id','Status','Post date']], on="id", how='left', suffixes=('','_previous',))



        #get only rows where the difference between statuses and post date occured

        new_df['Post date'] = new_df['Post date'].fillna("")

        new_df['Post date_previous'] = new_df['Post date_previous'].fillna("")

        new_df['Difference'] = where((new_df['Post date'] != new_df['Post date_previous']) | (new_df['Status'] != new_df['Status_previous']),"Yes","No")

        report_name = self.get_report_name() + " - full.csv"

        new_df.to_csv("J:/Public/tymczasowe/Raporty FINES - full/{0}".format(report_name), index=False, encoding='windows-1250')

        new_df = new_df[new_df['Difference'] == "Yes"]

        #new_df = new_df.drop(columns='Difference')

        # in case if with signed contract was erased due to lack of id

        new_df['Status_previous'] = where(new_df['Status'] == 'With signed contract', 'Approved',new_df['Status_previous'])

        new_df = new_df[['id','Proposal date','Decision date','Post date','Activation date','Limit','Status']]


        report_name = self.get_report_name() + ".csv"

        new_df.to_csv("OUTPUT/{0}".format(report_name), index=False, encoding='windows-1250')

        new_df.to_csv("J:/Public/tymczasowe/Raporty FINES/{0}".format(report_name), index=False, encoding='windows-1250')

        try:
            sftp = SFTP("OUTPUT/{0}".format(report_name),'/home/fines/public_html/{0}'.format(report_name))
            sftp.send_file()
            self.emsg.send_success_message()

        except Exception as e:
            self.emsg.send_critical_message(e)



    def correct_date_long_form(self,col):
        """Corrects create_time from long form"""
        if str(col) == "":
            return ""

        d = str(col)[7:11]

        d += '-'

        if str(col)[3:6] == 'Jan':
            d += '01'
        elif str(col)[3:6] == 'Feb':
            d += '02'
        elif str(col)[3:6] == 'Mar':
            d += '03'
        elif str(col)[3:6] == 'Apr':
            d += '04'
        elif str(col)[3:6] == 'May':
            d += '05'
        elif str(col)[3:6] == 'Jun':
            d += '06'
        elif str(col)[3:6] == 'Jul':
            d += '07'
        elif str(col)[3:6] == 'Aug':
            d += '08'
        elif str(col)[3:6] == 'Sep':
            d += '09'
        elif str(col)[3:6] == 'Oct':
            d += '10'
        elif str(col)[3:6] == 'Nov':
            d += '11'
        elif str(col)[3:6] == 'Dec':
            d += '12'

        d += '-'

        d += str(col)[:2]

        d += " "

        d += str(col)[12:17]

        d += ":00"

        return d

    def get_report_name(self):

        n = datetime.datetime.now()

        report_name = "FINES " + str(n.year) + "_"

        if len(str(n.month)) == 1:
            report_name += "0" + str(n.month) + "_"
        else:
            report_name += str(n.month) + "_"


        if len(str(n.day)) == 1:
            report_name += "0" + str(n.day) + "_"
        else:
            report_name += str(n.day) + "_"


        if len(str(n.hour)) == 1:
            report_name += "0" + str(n.hour) + "_"
        else:
            report_name += str(n.hour) + "_"


        if len(str(n.minute)) == 1:
            report_name += "0" + str(n.minute) + "_"
        else:
            report_name += str(n.minute) + "_"


        if len(str(n.second)) == 1:
            report_name += "0" + str(n.second)
        else:
            report_name += str(n.second)


        return report_name







