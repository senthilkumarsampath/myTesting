import re
from docx import Document
Doc = Document("/Users/senthil/Desktop/Senthil/myTesting/Test_Docx_clean-up.docx")
# regex for finding the figure, table and text box placeholder with numerial citations
spaces_pattern = re.compile ("^ +$")
break_char = "^(?:\<[Bb]r\>?)(?:eak)?(?:\>)(?: +)?$"
placeholder = "^[iI]nsert\s([fF]ig\.?(?:ure)?(?:s)?|[tT]able(?:s)?|(?:[Tt](?:ext))?\s?[bB](?:ox)(?:es)?)\s\d+(\.?\d+)?"
placeholder2 = re.compile ("(?i)^(?: +)?\<(?:insert )?[ft]?(?:igure|able)? ?\d+[.-]\d+(?: here)?\>")
for paragraphs in Doc.paragraphs:
    if re.search(spaces_pattern,paragraphs.text) or re.search(placeholder,paragraphs.text) or paragraphs.text.count("\n") == len(paragraphs.text) or re.search(break_char,paragraphs.text) or re.search(placeholder2,paragraphs.text):
        print(paragraphs.text)
        para = paragraphs._element
        para.getparent().remove(para)
        para._para = para._element = None
        # print(start_space)
Doc.save("/Users/senthil/Desktop/Senthil/myTesting/C01-test.docx")