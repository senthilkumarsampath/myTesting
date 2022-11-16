# import dstetl
# from dstetl.extractors.docx import extract_docx_properties

# docs = 'RL_01_AGHDACI_C001.docx'
# print (extract_docx_properties(docs))

# from email import header
# from pydoc import doc
# from re import T
# from turtle import clear
import docx
import csv
from docx import *
import os
# document = Document('C:/Users/senthilps/Downloads/RL_docx (2)/RL_docx/RL_01_Amavilah_C001.docx')
inlineTags = ['TBLCIT','title','gt','gd','DAY','MONTH','YEAR','Department','Country','CITY','INSTITUTION','STATE','AFFLABEL','email','URL','fnlink','AFFCIT','SURNAME','GIVENNAME','suffix','forename','prefix','degrees','AUTHOR','CORCIT','on-behalf-of','Bubble','REFCIT','FIGCIT','inlineequation','token','link','inline','chartlabel','chartcaption','CORAUTHOR','phone','fax','Rec','ACC','Rev','term','def','Speaker','line','SECCIT','figlabel','figcaption','figsource','figpara','figsource_break','alttext','SUPPCIT','seelink','FMT_Sub','Date','volume','issue','doi','edition','Editedby','imagelabel','imagecaption','Imprint','see','seealso','KW','KeywordTitle','ChartCIT','copyright','publicationdate','fpage','lpage','BOXCIT','CHAPCIT','PARTCIT','SUPPData','SchemesCIT','MapCIT','PlatesCIT','PhotoCIT','ImageCIT','Chapter','photolabel','photocaption','plateslabel','platescaption','page','type','REFID','journaltitle','pages','pubmed','crossref','chapter-title','publisher','labeltext','uri','season','misc','comment','supp','isbn','editor','booktitle','loc','collab','schemeslabel','schemescaption','seelabel','alt-text','tblfnlink','tbllabel','tblcaption','Abbreviation','Ack','AFFS','BACKMATTER','BL','BODYMATTER','Box_BulletList','Box_Group','Box_Group1','Box_NumberedList','Box_UnorderedList','CHAPTERBACKMATTER','CHAPTERFRONTMATTER','Contributors','Corresp-Group','Dialogue','Example_BulletList','Example_Group','Example_NumberedList','Example_UnorderedList','Extract','Extract_BulletList','Extract_Group','Extract_NumberedList','Extract_UnorderedList','Floats','Forward','FRONTMATTER','funding_group','glossary','glossary_text','Imprint_Group','Index','Index_Group','List_of_ILL','Math_BulletList','Math_Group','Math_NumberedList','Math_UnorderedList','NL','Part','Preface','Problem_BulletList','Problem_Group','Problem_NumberedList','Problem_UnorderedList','Programlisting','REF-LIST','Series_Group','SL','TOC','Vignette_BulletList','Vignette_Group','Vignette_NumberedList','Vignette_UnorderedList',]

def pTagExtraction(docs):
    dName, fName = os.path.split(docs)
    c, ext = os.path.splitext(fName)
    csvFileName = dName+"/"+c+".csv"
    print(csvFileName)
    aDoc = Document(docs)
    stdTags = aDoc._element.xpath('.//ancestor::w:sdt//w:tag/@w:val')
    docTags = [] 
    isfound = False
    for tags in stdTags:
        isfound = False
        for itag in inlineTags:
            if (tags.lower()==itag.lower()):
                isfound=True
                break
        if(isfound == False):
            docTags.append(tags)
    
    paraID = 1

    with open(csvFileName, 'w', newline='') as pCSV:
        fTags = csv.writer(pCSV)
        fTags.writerow(["paraID", "pTags"]) 
        for prtag in docTags:
            fTags.writerow([paraID, prtag])
            paraID+=1
   
path = "C:/Users/senthilps/Downloads/RL_docx (2)/RL_docx/"  

# collecting all the files in the folder
# print(path)
# pTagExtraction(path)
dir_list = os.listdir(path)

# print(dir_list[0])

for p in dir_list:
    aPath = path+"/"+p
    # print(aPath)
    pTagExtraction(aPath)
    

