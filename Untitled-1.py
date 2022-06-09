# import sys
# from bs4 import BeautifulSoup
# import requests
# import re
# import spacy
# import textwrap
# importing required modules
from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.converter import TextConverter
from pdfminer.layout import LAParams
from pdfminer.pdfpage import PDFPage
from io import StringIO
import docx2txt
from striprtf.striprtf import rtf_to_text

def convert_rtf_to_text(path):
    with open(path) as file:
         content = file.read()
    text = rtf_to_text(content)
    return text

def convert_docx_to_txt(path):
    text = docx2txt.process(path)
    return text

def convert_pdf_to_txt(path):
    rsrcmgr = PDFResourceManager()
    retstr = StringIO()
    codec = 'utf-8'
    laparams = LAParams()
    device = TextConverter(rsrcmgr, retstr, codec=codec, laparams=laparams)
    fp = open(path, 'rb')
    interpreter = PDFPageInterpreter(rsrcmgr, device)
    maxpages = 0
    caching = True
    pagenos=set()

    for page in PDFPage.get_pages(fp, pagenos, caching=caching, check_extractable=True):
        interpreter.process_page(page)
    text = retstr.getvalue()
    fp.close()
    device.close()
    retstr.close()
    return text

filepath='C:/Users/userpc/Documents/file-sample_100kB.rtf'
text=""
file_extns = filepath.split(".")
if(file_extns[-1]=="pdf"):
    text = convert_pdf_to_txt(filepath)
elif(file_extns[-1]=="rtf"):
    text=convert_rtf_to_text(filepath)
else:
    text=convert_docx_to_txt(filepath)
print(text)


# !{sys.executable} -m spacy download en_core_web_sm
# spacy_model = 'en_core_web_sm'

# # URL = input("Enter URL :")
# URL = "https://www.fool.com/earnings/call-transcripts/2022/05/17/walmart-inc-wmt-q1-2023-earnings-call-transcript/"
# page = requests.get(URL)

# # soup = BeautifulSoup(page.content, 'html.parser')
# # print(soup.prettify())
 
# regex = re.compile(r'Prepared Remarks:\s*(\n*.*)\s*Duration:',re.DOTALL)
# matches = regex.finditer(page.text)
# plain_text=''
# for match in matches:
#     # print(match)
#     plain_text = BeautifulSoup(match.group(1), 'html.parser').get_text(separator='') # remove all the html
#     plain_text = re.sub("\n|\r", ".", plain_text, flags=re.MULTILINE) # remove new line
#     plain_text = re.sub("\s\s+|\t", " ", plain_text) # replace multiple spaces with single space
#  # printing the plain text   
# # wrap_list = textwrap.wrap(plain_text, 120)
# # print('\n'.join(wrap_list))

# #load spacy model
# nlp = spacy.load(spacy_model)

# #removing stop words
# additional_stop_words = ['hi', 'earning', 'conference', 'speaker', 'analyst', 'operator', 'welcome', \
#                          'think', 'cost', 'result', 'primarily', 'overall', 'line', 'general', \
#                           'thank', 'see', 'alphabet', 'google', 'facebook', 'amazon', 'microsoft',\
#                         'business', 'customer', 'revenue', 'question', 'lady', 'gentleman', \
#                         'continue', 'continuing', 'continued', 'focus', 'participant', 'see', 'seeing', \
#                         'user', 'work', 'lot', 'day',  'like', 'looking', 'look', 'come', 'yes', 'include', \
#                         'investor', 'director', 'expense', 'manager', 'founder', 'chairman', \
#                          'chief', 'operating', 'officer', 'executive', 'financial', 'senior', 'vice', 'president', \
#                         'opportunity', 'go', 'expect', 'increase', 'quarter', 'stand', 'instructions', \
#                         'obviously', 'thing', 'important', 'help', 'bring', 'mention', 'yeah', 'get', 'proceed', \
#                         'currency', 'example', 'believe'] 

# for stopword in additional_stop_words:
#     nlp.vocab[stopword].is_stop = True

# cleaned_words = []

# doc = nlp(plain_text)
# with doc.retokenize() as retokenizer:
#     for ent in doc.ents:
#         # print(ent.text, ent.label_)
#         retokenizer.merge(doc[ent.start:ent.end], attrs={"LEMMA": ent.text})
# #print('-------------------')
# for word in doc:
#     # print(word, word.lemma_, word.ent_type_)
#     if word.is_alpha and word.is_ascii and not word.is_stop and \
#         word.ent_type_ not in ['PERSON','DATE', 'TIME', 'ORDINAL', 'CARDINAL'] and \
#         word.text.lower() not in additional_stop_words and \
#         word.lemma_.lower() not in additional_stop_words:
#             #print(word)
#             cleaned_words.append(word.lemma_.lower())
# print(cleaned_words)