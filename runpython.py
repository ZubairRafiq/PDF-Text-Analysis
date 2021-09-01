from sys import flags
import xlwings as xw
import pandas as pd
import numpy as np
import PyPDF2
import re
import os
from pathlib import Path 
from tika import parser # pip install tika

from typing import Dict
import fitz  # pip install pymupdf

import nltk
nltk.download('stopwords')
from nltk.corpus import stopwords
from nltk.tokenize import word_tokenize
stop_words = set(stopwords.words('german')) #Get german Stop Words
newStopWords = {'ungheinrich'}
stop_words = stop_words.union(newStopWords)


# In[303]:

def getFilename(wb):
    sheet = xlsxwriter.Workbook('runpython.xlsm')
    #global fname
    fname = sheet.range('A2').value
    sheet["A3"].value = fname
    return fname


def getDetails(wb, pdf_file):               #Function for displaying details in Excel
    sheet = wb.sheets['PDF_Details']
    
    def getRowCell():               #Function for getting next empty cell number
            Y = wb.sheets['PDF_Details'].range('D' + str(wb.sheets['PDF_Details'].cells.last_cell.row)).end('up').row + 1
            return Y

    lower = getRowCell()        # Get next empty cell in Row D

    
    if(getDetails.flag == True ):           #Check for flag to display header in first iteration
        getDetails.flag = False
        #sheet.range('A2:Q10000').clear_contents()
      
        sheet["A" + str(lower-1)].value = "Filename"
        sheet["B" + str(lower-1)].value = "Type"
        sheet["C" + str(lower-1)].value = "Position"
        sheet["D" + str(lower-1)].value = "Content"
        sheet["E" + str(lower-1)].value = "Frequency"
    #else:
        #getDetails.i += 30
        #sheet["A"+ str(getDetails.i-1)].value = "Filename = " + str(pdf_file.name)

    sheet["D" + str(lower)].value = df.head(10)     #Display list of keywords

    middle = getRowCell()       # Get empty cell after keywords in Row D

    sheet["C" + str(middle)].value = get_bookmarks(str(pdf_file))       #Display list of bookmarks


    

def main():
    wb = xw.Book.caller()
    
    #getFilename(wb)
    #getDetails(wb)

if __name__ == "__main__":
    xw.Book("runpython.xlsm").set_mock_caller()
    main()
   



wb = xw.Book.caller()
sheet = wb.sheets[0]

#dir_name = r"C:\Users\zubai\runpython\files"
#base_filename = getFilename(wb)
#suffix = ".pdf"
#os.path.join(dir_name, str(base_filename) + suffix)

PATH_PDF_FILES = sheet.range('source_dir').value            #Get folder path from Excel Cell (Soucre Dir)

pdf_files = list(Path(PATH_PDF_FILES).glob('*.pdf'))        #List all PDF file paths in a List


getDetails.flag = True
getDetails.i = 4

for pdf_file in pdf_files:                                  # iterate through each file
    raw = parser.from_file(str(pdf_file))                   
    text = raw['content']
    
    keywords = re.findall(r'[a-zA-Z]\w+',text)              # Get keywords without digits and special characters
    keywords = [w for w in keywords if not w.lower() in stop_words and w.split() if len(w)>3]       # remove German stop words and words less than 3 chars

    df = pd.DataFrame(list(set(keywords)),columns=['keywords'])  #Dataframe with unique keywords to avoid repetition in rows

    def weightage(word,text,number_of_documents=1):
        word_list = re.findall(word,text)
        number_of_times_word_appeared =len(word_list)
        tf = number_of_times_word_appeared/float(len(text))
        idf = np.log((number_of_documents)/float(number_of_times_word_appeared))
        tf_idf = tf*idf
        return number_of_times_word_appeared,tf,idf ,tf_idf

    df['Frequency'] = df['keywords'].apply(lambda x: weightage(x,text)[0])  

    df = df.sort_values('Frequency',ascending=False)            # Add sorted frequecy list 
    #df.to_csv('keywords.csv')
    df = df.set_index('keywords')
    df.head(10)

    def get_bookmarks(filepath: str) -> Dict[int, str]: # WARNING! One page can have multiple bookmarks!
        bookmarks = {}
        with fitz.open(filepath) as doc:
            toc = doc.getToC()  # [[lvl, title, page, …], …]
            for level, title, page in toc:
                bookmarks[page] = title
        return bookmarks
    
    getDetails(wb, pdf_file)




# In[2]:





# In[332]:




""" raw = parser.from_file(os.path.join(dir_name, str(base_filename) + suffix))
text = raw['content'] """



# In[333]:


#keywords = re.findall(r'[a-zA-Z]\w+',text)



# In[334]:

#keywords = [w for w in keywords if not w.lower() in stop_words and w.split() if len(w)>3]



# In[308]:


#print(keywords)


# In[309]:


#df = pd.DataFrame(list(set(keywords)),columns=['keywords'])  #Dataframe with unique keywords to avoid repetition in rows


# In[310]:


""" def weightage(word,text,number_of_documents=1):
    word_list = re.findall(word,text)
    number_of_times_word_appeared =len(word_list)
    tf = number_of_times_word_appeared/float(len(text))
    idf = np.log((number_of_documents)/float(number_of_times_word_appeared))
    tf_idf = tf*idf
    return number_of_times_word_appeared,tf,idf ,tf_idf """


# In[311]:


#df['Frequency'] = df['keywords'].apply(lambda x: weightage(x,text)[0])


# In[312]:


#df = df.sort_values('Frequency',ascending=False)
#df.to_csv('keywords.csv')
#df.head(25)

#df1 = pd.read_excel(r"C:\Users\zubai\runpython\runpython.xlsm", sheet_name='PDF_Details')
#names = df['keywords'][0:6]
#values = df['Frequency'][0:6]

#fig = px.pie(df, values = values, names = names, title = 'Results' )

#fig.update_traces(textposition = 'inside', textinfo = 'percent+label')

#fig.update_layout(title_font_size = 42)

#t = time.localtime()
#timestamp = time.strftime('%Y-%m-%d_%H%M', t)

#plotly.offline.plot(fig, filename = f'Piechart_{timestamp}.html')

#output_path = str(Path(__file__).parent / 'myplot.html')
#plotly.offline.plot(fig,filename=output_path)

#plotly.offline.plot(fig, filename ='Piechart.html')

#names 
# In[1]:





""" def get_bookmarks(filepath: str) -> Dict[int, str]:
    # WARNING! One page can have multiple bookmarks!
    bookmarks = {}
    with fitz.open(filepath) as doc:
        toc = doc.getToC()  # [[lvl, title, page, …], …]
        for level, title, page in toc:
            bookmarks[page] = title
    return bookmarks


print(get_bookmarks(os.path.join(dir_name, str(base_filename) + suffix))) """


# In[ ]:
# In[338]:




""" 
def show_tree(bookmark_list, indent=0):
    for item in bookmark_list:
        if isinstance(item, list):
            # recursive call with increased indentation
            show_tree(item, indent + 4)
        else:
            print(" " * indent + item.title)
            


reader = PyPDF2.PdfFileReader(os.path.join(dir_name, str(base_filename) + suffix))

show_tree(reader.getOutlines()) """


# In[3]:


""" wb = xw.Book.caller()
#filename(wb)
getDetails(wb) """

""" def getDetails(wb):
    sheet = wb.sheets['PDF_Details']
    sheet.range('1:50').clear()
    sheet["A1"].value = df.head(25)
    sheet["E1"].value = get_bookmarks(os.path.join(dir_name, base_filename + suffix)) """


""" def main():
    wb = xw.Book.caller()
    sheet = wb.sheets['PDF_Details']
    filename(wb)
    sheet.range('1:50').clear()
    sheet["A1"].value = df.head(25)
    sheet["E1"].value = get_bookmarks(os.path.join(dir_name, base_filename + suffix)) """
  




""" if __name__ == "__main__":
    xw.Book("runpython.xlsm").set_mock_caller()
    main() """
# In[ ]:




