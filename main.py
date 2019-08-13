import datetime
from download_files import Downloads
from report_prep import Reporting



def time_between(low,high):
    """Allows to decide when the code runs"""
    tm = datetime.datetime.now()
    if tm.hour >= low  and tm.hour <=high:
        return True
    return False

def not_weekend():
    if datetime.datetime.today().weekday() in [5,6]:
        return False
    return True


def download_FINES(code):

    downloads = Downloads()

    downloads.download_Proposals_DC1(code)

    downloads.download_Proposals_DC2(code)

    downloads.download_Credit_Cards_DC1(code)

    downloads.download_Credit_Cards_DC2(code)

    downloads.concat_Proposals()

    downloads.concat_Credit_Cards()

    downloads.download_Reports_Cards()


    if time_between(7, 7):

        downloads.download_Raport_do_CC()

    downloads.download_Processing()

    downloads.download_Dealzilla(5)

    downloads.convert_xls_to_xlsx()

def prepare_FINES():


    report = Reporting()


    report.card_proposals_preparation()


    report.credit_cards_preparation()


    report.concat_long_form()


    report.concat_dataframes()


    report.create_file()


    report.compare_files()



if __name__ == "__main__":


    code = 'Fines'

    if time_between(6,18) and not_weekend():

        download_FINES(code)

        prepare_FINES()

