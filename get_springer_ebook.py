import openpyxl 
import threading
import sys
from bs4 import BeautifulSoup as BS 
import requests
from pathlib import Path 
import pickle
import traceback
import time



def thread_get_book(row,saved_title):
    try:
        if row[0] == None : return 
        if row[0] in saved_title : return  #already has this title 
        
        path = "./ebook/"+row[11] # use category as sub directory 
        Path(path).mkdir(parents=True,exist_ok=True)
        path_file = path+"/"+row[0] + ".pdf"    #use title as file name 
        file_book = Path(path_file)
        if not file_book.exists():      # check if this ebook has already been downloaded 
            book_page = requests.get( row[18] )      #go to the book page
            soup = BS(book_page.text,"html.parser")
            item = soup.find_all("a", attrs= {"title":"Download this book in PDF format"})
            req = requests.get("https://link.springer.com"+ item[0]["href"])  # this is the url of pdf 
            f = open(path_file,"wb")
            f.write(req.content)
            f.close()
            saved_title.add(row[0])
            print( " Downloaded ... Title : {} |  Category : {} |  Link : {} |  url = {}".format(row[0],row[11],row[18], item[0]["href"]))
        #else :
        #    print("This file already exists :",path_file)
    except:
        traceback.print_exc()




if __name__ == '__main__':
    wb = openpyxl.load_workbook("./english_textbooks.xlsx",data_only=True)
    sheet = wb[wb.sheetnames[0]]
    c = 0
    list_tb = list(sheet.values )
    header = list_tb[0]
    p_file = Path("saved_title.pickle")
    MAX_THREADS = 10 # maximum number of threads that we will create
    if p_file.exists(): # if this is not the first time we run 
        saved_title = pickle.loads(open("saved_title.pickle","rb").read())
    else: 
        saved_title = set()


    thread_pool = []

    try:
        for row in list_tb[1:len(list_tb)]:
            a_thread = threading.Thread(target=thread_get_book, args=(row,saved_title))
            while  threading.activeCount() > MAX_THREADS :  # if the max is reached , just wait for 0.1 sec 
                time.sleep(0.1)
            a_thread.start()
            thread_pool.append(a_thread)
    except:
        traceback.print_exc()
    finally:    
        for th in thread_pool:
            th.join()
        pic_file = open("saved_title.pickle","wb")
        pic_file.write(pickle.dumps(saved_title))
        pic_file.close()

