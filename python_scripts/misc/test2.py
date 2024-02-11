""" Module for converting docx to xml for element prediction"""
import os
import re
import uuid
import logging
from collections import Counter
import docx
import docx.package
import docx.parts.document
from docx.document import Document
import docx.parts.numbering
from docx.oxml.table import CT_Tbl
from docx.oxml.text.paragraph import CT_P
from docx.oxml.numbering import CT_Num
from docx.table import _Cell, Table
from docx.text.paragraph import Paragraph
from docx.oxml import parse_xml
from docx.oxml.ns import nsmap
from lxml import etree
import pandas as pd

# define info log
logging.basicConfig(
    format='%(asctime)s %(levelname)-8s %(message)s',
    level=logging.INFO,
    datefmt='%Y-%m-%d %H:%M:%S')

# health check 
def health_check():
    """Function to perform a health check (Used for testing that the package gets loaded)"""
    output_str = "Health Check OK"
    return output_str


# extract docx properties
def extract_docx_properties(docx_path):
    """ Function to create an XML document that can be passed to Element Prediction
    Args:
        docx_path (str): String containing the path to the docx that should be used for extraction
    Returns:
        [str]: An xml formatted string that contains extracted properties from the docx
    """
    # invoke create_document_object() to create document obj
    logging.info('Started creating the docx object')
    document, numbering_pd, theme_dict = create_document_object(docx_path)

    # Default value, empty string
    para_properties_xml = ""
    # Create list that will hold all the properties which should be treated as CDATA
    c_data_tags = ["ParaContent"]
    # Create a list of dictionaries with the properties of each paragraph
    logging.info('Extracting the properties of the document')
    para_properties_list = extract_properties_to_list(
        document,
        numbering_pd,
        theme_dict)

    logging.info('Writing the extracted properties into XML')
    # Convert list of dictionaries to an XML string
    para_properties_xml = create_structured_xml(
        para_properties_list,
        c_data_tags)

    return para_properties_xml


# create document object
def create_document_object(docx_path):
    """ Function to create a python-docx Document object
    Args:
        docx_path (str): String containing the path to the docx file
    Returns:
        python-docx Document: A python-docx Document object of the docx file
    """
    document = None
    numbering_pd = None
    theme_dict = None

    # First check that a string with characters has been passed
    if len(docx_path) > 0:
        # Check to see if the file exists
        if os.path.exists(docx_path):
            docx_package = docx.package.Package.open(docx_path)
            # Create the Numbering DataFrame
            try:
                numbering = docx_package.main_document_part.numbering_part

            except (RuntimeError, TypeError, NameError, AttributeError, NotImplementedError):
                numbering = None

            if not numbering is None:
                numbering_pd = create_numbering_pd(numbering_part=numbering)

            # Create the Theme Dictionary
            theme_dict = get_theme_data(docx_package)

            document = docx.Document(docx_path)

    return document, numbering_pd, theme_dict


# create numbering pd
def create_numbering_pd(numbering_part):
    """ Function to create a Pandas Dataframe representing the numbering.xml file in a docx
    Args:
        numbering_part (python-docx parts object): A python-docx parts object representing
            the numbering.xml file
    Returns:
        [pandas Dataframe]: A converted pandas DataFrame of the numbering.xml file for easy access
    """
    namespace = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
    num_list = []
    abstract_num_list = []

    # Loop through each part of numbering.xml
    if numbering_part._element.getchildren():
        for num in numbering_part._element:
            # The abstractNum part
            if not isinstance(num, CT_Num):
                # Loop through each child, adding value to a dict
                for abstract_element in num.iterchildren():
                    abstract_num_dict = {
                        "abstract_num_id" : num.get(namespace + "abstractNumId")
                    }
                    # Get MultilevelType
                    if abstract_element.tag == (namespace + "multiLevelType"):
                        abstract_num_dict["multi_level_type"] = abstract_element.get(namespace + "val")
                    if abstract_element.tag == (namespace + "lvl"):
                        # Add Dictionary for level based info
                        abstract_num_level_dict = abstract_num_dict
                        abstract_num_level_dict["level"] = abstract_element.get(namespace + "ilvl")

                        for level_element in abstract_element.iterchildren():
                            if level_element.tag == (namespace + "start"):
                                abstract_num_level_dict["level_start"] = level_element.get(namespace + "val")
                            if level_element.tag == (namespace + "numFmt"):
                                abstract_num_level_dict["level_num_format"] = level_element.get(namespace + "val")
                            if level_element.tag == (namespace + "lvlText"):
                                abstract_num_level_dict["level_text"] = level_element.get(namespace + "val")
                            if level_element.tag == (namespace + "pPr"):
                                for level_para_prop_element in level_element.iterchildren():
                                    if level_para_prop_element.tag == (namespace + "ind"):
                                        abstract_num_level_dict["level_para_prop_left"] = level_para_prop_element.get(namespace + "left")
                                        abstract_num_level_dict["level_para_prop_hanging"] = level_para_prop_element.get(namespace + "hanging")
                        abstract_num_list.append(abstract_num_level_dict)
            else:
                # The numID part that is in document.xml
                num_dict = {
                    "num_id" : num.get(namespace + "numId"),
                    "abstract_num_id" : str(num.abstractNumId.val)}
                # append into num_list
                num_list.append(num_dict)

    # Convert Lists to DataFrame
    if num_list:
        num_id_pd = pd.DataFrame(num_list)
        abstract_num_pd = pd.DataFrame(abstract_num_list)

    # Merge into a single Data Frame
        numbering_pd = pd.merge(
            num_id_pd,
            abstract_num_pd,
            how = "left",
            on = "abstract_num_id")

        return numbering_pd


# get data from numbering pd
def get_data_from_numbering_pd(numbering_pd, num_id, level, column):
    """ Function return spacific data from the numbering_pd DataFrame
    Args:
        numbering_pd (pandas DataFrame): The pandas Dataframe corresponding to numbering.xml, see
            create_numbering_pd()
        num_id (int): An integer corresponding to the List Para Style to be retrieved
        level (int): An integer corresponding to the level that the list is at
        column (str): String corresponding to the column name to be returned from the numbering_pd
            dataframe
    Returns:
        [str]: The found data value corresponding to the inputs
    """

    output = ""
    # check if numbering_pd is not defined
    if not numbering_pd is None:
        # check if it is not empty if its defined
        if len(numbering_pd.index) > 0 and \
            column in numbering_pd.columns and \
            int(num_id) >= 0:

            numbering_filtered = numbering_pd[numbering_pd.num_id == num_id]
            if int(level) > 0:
                numbering_filtered = numbering_filtered[numbering_filtered.level == level]
                output = numbering_filtered[column].to_list()[0]

    return output


# get fonts found into theme xml into docx 
def get_theme_data(docx_package):
    """ Function to create a dictionary containing the major and minor fonts found in theme.xml
    Args:
        docx_package (python-docx docx.package.Package object): A python-docx Package object
    Returns:
        [dict]: A dictionary containing two keys containing string values, major and minor fonts
    """
    output = {
        "major_font" : "",
        "minor_font" : ""
    }
    # Iterate through each of the parts in the docx package
    for part in docx_package.parts:
        if part.partname.startswith("/word/theme/"):
            theme_xml = parse_xml(part.blob)

            for major_font in theme_xml.xpath("//a:majorFont/a:latin/@typeface",
                namespaces = nsmap):
                output["major_font"] = major_font

            for minor_font in theme_xml.xpath("//a:minorFont/a:latin/@typeface",
                namespaces = nsmap):
                output["minor_font"] = minor_font

    return output


""" 
Function to iterate through paragraphs of a document and return each paragraph in a list
Input:
- document: python-docx document object, see create_document_object
Output:
- para_properties_list: A list containing dictionaries which contain properties of the paragraph
"""


# extract docx properties into list
def extract_properties_to_list(document, numbering_pd, theme_dict):
    """ Function to iterate through paragraphs of a document, extract relevant information
        about the paragraph, store as a dictionary, and append to a list
    Args:
        document (python-docx docx.Document): A python-docx Document object representing the .docx file
        numbering_pd (pandas DataFrame): A pandas DataFrame representing the numbering.xml
        theme_dict (dict): A dictionary for the major and minor fonts
    Returns:
        [list]: A list containing dictionaries which store the information on each paragraph
    """

    block_id = 1
    document_properties_list = []
    logging.info('Started processing the components to extract the properties')
    for document_block in iter_block_items(document):
        # Check to see if the paragraph has a text box
        para_contains_text_box = para_contains_xpath(
            document_block,
            xpath_string = ".//v:textbox/w:txbxContent")

        # Check to see if paragraph contains an inline image
        para_contains_linked_image = para_contains_xpath(
            document_block,
            xpath_string = ".//w:drawing/wp:inline/a:graphic/a:graphicData/pic:pic"
        )

        # Check to see if paragraph contains an word shape
        para_contains_shape = para_contains_xpath(
            document_block,
            xpath_string = ".//w:drawing/wp:inline/a:graphic/a:graphicData/dgm:relIds"
        )

        ##########################
        ######## Table ###########
        ##########################
        if isinstance(document_block, Table):
            pass
            # Iterate through the rows, cells and paragraphs in the table
            # for table_row in document_block.rows:
            #     for cell in table_row.cells:
            #         for para in cell.paragraphs:
            #             # Uncommented this line to exclude any paragraphs with blank or only new lines
            #             document_properties_list.append(create_paragraph_properties(
            #                 document,
            #                 para,
            #                 block_id,
            #                 block_type = "table_cell_paragraph",
            #                 numbering_pd = numbering_pd,
            #                 theme_dict = theme_dict))

            #             block_id += 1

        # Check to see if the current block is a paragraph and contains textbox
        elif isinstance(document_block, Paragraph) and para_contains_text_box:
            ################################
            ### Paragraphs with TextBox ####
            ################################
            text_box_para = etree.ElementBase.xpath(
                document_block._element,
                '//v:textbox/w:txbxContent/w:p',
                namespaces = document_block._element.nsmap)

            # Iterate through all the paragraphs in the textboxes
            for txt_box_para in text_box_para:
                # Convert the paragraph (w:p) to a Paragraph Class
                txt_box_para_class = Paragraph(
                    txt_box_para,
                    document_block
                )

                # Identify the properties
                document_properties_list.append(create_paragraph_properties(
                    document,
                    txt_box_para_class,
                    block_id,
                    block_type = "text_box_paragraph",
                    numbering_pd = numbering_pd,
                    theme_dict = theme_dict))

                block_id += 1

        elif isinstance(document_block, Paragraph):
            block_type = ""
            if para_contains_linked_image:
                block_type = "linked_image_paragraph"
            elif para_contains_shape:
                block_type = "shape_paragraph"
            else:
                block_type = "paragraph"

            ####################################
            ### Paragraph with Linked Image ####
            ####################################
            document_properties_list.append(create_paragraph_properties(
                document,
                document_block,
                block_id,
                block_type = block_type,
                numbering_pd = numbering_pd,
                theme_dict = theme_dict))

            block_id += 1

    return document_properties_list


# iterate over document
def iter_block_items(parent):
    """
    Yield each paragraph and table child within *parent*, in document
    order. Each returned value is an instance of either Table or
    Paragraph. *parent* would most commonly be a reference to a main
    Document object, but also works for a _Cell object, which itself can
    contain paragraphs and tables.
    """
    if isinstance(parent, Document):
        parent_elm = parent.element.body
    elif isinstance(parent, _Cell):
        parent_elm = parent._tc
    else:
        raise ValueError("something's not right")

    for child in parent_elm.iterchildren():
        if isinstance(child, CT_P):
            yield Paragraph(child, parent)
        elif isinstance(child, CT_Tbl):
            yield Table(child, parent)


# check if pata contains specified xpath
def para_contains_xpath(para, xpath_string):
    """ Function to check if an XML block contains a certain xpath query
    Args:
        para (python-docx Paragraph): Python-docx paragraph object to check for a certain xpath
        xpath_string (str): String containing the xpath to be queried
    Returns:
        boolean: True/False indicating if the para XML contains the xpath
    """
    # Combine Namespaces from the package and the paragraph
    custom_nsmap = dict(list(nsmap.items()) + list(para._element.nsmap.items()))
    xml_value = etree.ElementBase.xpath(
        para._element,
        xpath_string,
        namespaces = custom_nsmap)

    return len(xml_value) > 0


# create paragraph properties
def create_paragraph_properties(document, para, para_id, block_type, numbering_pd, theme_dict):
    """ Function to create a paragraph properties dictionary
    Args:
        document (python_docx Document): Python-docx Document object
        para (python-docx Paragraph): Python-docx Paragraph object
        para_id (int): Sequential integer indicating the order in which the document block
            occurs in the document
        block_type (str): Type of document block, example, paragraph, text_box_paragraph,
            table_cell_paragraph
        numbering_pd (pandas DataFrame): A pandas DataFrame corresponding to numbering.xml
        theme_dict (dict): A dictionary containing keys for the major and minor fonts
    Returns:
        dict: Dictionary containing the paragraph properties for the document block
    """

    para_prop_dict = {}
    para_prop_dict["ParaID"] = para_id
    para_prop_dict["ParaObjectType"] = block_type
    para_prop_dict["ParaHexId"]=retreive_para_hex_id(para)
    para_prop_dict["ParaCleanedContent"] = transform_para_content(get_para_content(para))
    para_prop_dict["ParaContent"] = get_para_content(para)
    para_prop_dict["ParaContentTabStart"] = get_para_content_tab_start_count(para)
    para_prop_dict["ParaFontFamily"] = get_para_font_family(document, para, theme_dict)
    para_prop_dict["ParaBold"] = get_para_bold(document, para)
    para_prop_dict["ParaItalic"] = get_para_italic(document, para)
    para_prop_dict["ParaFontSize"] = get_para_font_size(document, para)
    para_prop_dict["ParaStyle"] = get_para_style(para)
    para_prop_dict["ParaListStyle"] = get_para_list_style(document, para, numbering_pd)
    para_prop_dict["ParaLeftIndent"] = get_para_left_indent(document, para, numbering_pd)
    para_prop_dict["ParaRightIndent"] = get_para_right_indent(document, para)
    para_prop_dict["ParaFirstLineIndent"] = get_para_first_line_indent(document, para)
    para_prop_dict["ParaAlignment"] = get_para_alignment(document, para)
    para_prop_dict["ParaLineSpace"] = get_para_line_space(document, para)
    para_prop_dict["ParaAboveSpace"] = get_para_space_above(document, para)
    para_prop_dict["ParaBelowSpace"] = get_para_space_below(document, para)

    para_prop_dict = get_para_border(para_prop_dict, para)

    para_prop_dict = get_para_shading(para_prop_dict, para)

    para_prop_dict["ParaSingleStrike"] = get_para_single_strike(para)
    para_prop_dict["ParaDoubleStrike"] = get_para_double_strike(para)
    para_prop_dict["ParaUnderline"] = get_para_underline(document, para)
    para_prop_dict["ParaSmallCaps"] = get_para_small_caps(para)

    return para_prop_dict


# get xml attribute
def get_xml_attribute(
    para,
    tag_to_find,
    tag_parent,
    tag_attribute,
    default_value = 0):
    """ Function to retrieve an xml attribute for a given tag and parent tag
    Args:
        para (python-docx paragraph object): Python-docx paragraph object
            found through iterating the document
        tag_to_find (str): The tag to retreive attribute values for
        tag_parent (str): The parent tag_to_find should have. Found using .getparent().tag
        tag_attribute (str): The attribute value to be retrieved from the element
        default_value (int, optional): A default value to be returned in the
            event no value is found. Defaults to 0.
    Returns:
        str: The attribute value found in the XML
    """
    # Set-up the return value to be the default value
    attribute_value = default_value
    # Convert para element xml to an etree data structure
    para_xml = etree.XML(para._element.xml)

    # Iterate directly to the tag that should be found
    for element in para_xml.iter(tag_to_find):
        # Check if the tag is the same and the parent is as expected
        if element.tag == tag_to_find and element.getparent().tag == tag_parent:
            # Check if the attribute to be retrieved is actually there
            if tag_attribute in element.keys():
                # Retrieve the attribute values from the dict
                attribute_value = element.get(tag_attribute)
                # Exit the loop as we have got what we needed
                break

    return attribute_value


# retrieve para hex id
def retreive_para_hex_id(para):
    """Function to generate random hex id for each paras.

    Args:
        para (docx obj): docx.Document obj.

    Returns:
        Generated para ID.
    """
    # define wordml xpath
    word14_namespace_ml = "{http://schemas.microsoft.com/office/word/2010/wordml}"
    # generate para ID into input_doc
    if (len(para._p.xpath("@w14:paraId")) < 1):
        para._p.set(
            word14_namespace_ml + "paraId", gen_id())
        return para._p.xpath("@w14:paraId")[0]
    else:
        return para._p.xpath("@w14:paraId")[0]


# get para content
def get_para_content(para):
    """ Function to retrieve the paragraph content across all runs
    Input:
    - para: paragraph object from document
    Output:
    - para_content: String containing the text for the paragraph
    """
    para_content = ""
    # Check to ensure that the para is not None
    if not para is None:
        # Retrieve the text from all the runs in the paragraph
        para_content = para.text

    return para_content


# transform para contents
def transform_para_content(para_content):
    """ Function to transform raw paragraph content into a cleaned version
    Input:
    - para_content: String containing the raw paragrah content, see get_para_content
    Output:
    - para_content_transformed: String containing the cleaned/transformed paragraph content
    """

    para_content_transformed = para_content

    # Replace ` with a '
    para_content_transformed = re.sub('`', "'", para_content_transformed)

    return para_content_transformed


# get number of tab characters ar start of a str of paragraphs
def get_para_content_tab_start_count(para):
    """ Function to count the number of tab characters at the start of a string of paragraph of text
    Args:
        para (python-docx Paragraph object): A python-docx Paragraph corresponding to a paragraph in a word document
    Returns:
        [int]: The number of tab characters at the start of a paragraph of text, default to 0
    """
    output = 0
    try:
        if len(para.text) > 0:
            i = 0
            while True:
                if para.text[i] == "\t" and i < len(para.text):
                    i += 1
                else:
                    break
            output = i
        return output
    except Exception as error:
        print(f"Error occured while taking the tab start count: {error}")


# get para font family
def get_para_font_family(document, para, theme_dict):
    """ Function to retrieve the font family for a specific paragraph, the run is
        also included to check for any run specific fonts
    Args:
        document (python-docx document object): Python-docx document object
        para (python-docx Paragraph object): A python-docx object of the paragraph
        theme_dict (dict): A dictionary containing keys for the major and minor fonts
    Returns:
        [str]: The font family identified for the paragraph
    """
    try:
        # Set the default font_family to be Word default of Calibir
        font_family = "Default"
        # Check to esnure the paragraph object is not None
        if not para is None:
            # Direct font family name for the run and style values for the run
            run_font_values = []
            run_style_font_values = []
            # Iterate through each run to populate the two above lists
            for run in para.runs:
                if not run.text == "" and not run.text == "\n":
                    run_font_values.append(run.font.name)
                    if run.style is not None:
                        run_style_font_values.append(run.style.font.name)

            # Check the paragraph style font family name
            para_style_font_name = para.style.font.name

            # Identify the paragraph style
            para_style = get_para_style(para)
            # Check the style from the document styles
            if para_style is not None:
                document_para_style = document.styles[para_style].font.name

                # Check the theme_dict for a font at theme level
                theme_font = None
                if para_style.lower().find("heading") >= 0 and "major_font" in theme_dict.keys():
                    theme_font = theme_dict.get("major_font")
                elif para_style.lower().find("heading") < 0 and "minor_font" in theme_dict.keys():
                    theme_font = theme_dict.get("minor_font")


            # Check which of the above variables should be used, based on the following logic
            # - The font directly applied to the paragraph
            # - The font for any style applied to the paragraph
            # - The majority font for any style applied to the runs within the paragraph
            # - The majority font directly applied to any run within the paragraph
            # - The Major/Minor Font used in theme.xml, depending on the heading
            if para_style_font_name is None:
                if document_para_style is None:
                    # Condition to check if all values are None in the list
                    if all(run_font_value is None for run_font_value in run_style_font_values):
                        if all(run_font_value is None for run_font_value in run_font_values):
                            if not theme_font is None:
                                font_family = theme_font
                        else:
                            run_values_counter = Counter(run_font_values)
                            font_family = run_values_counter.most_common(1)[0][0]
                    else:
                        run_values_counter = Counter(run_style_font_values)
                        font_family = run_values_counter.most_common(1)[0][0]
                else:
                    font_family = document_para_style
            else:
                font_family = para_style_font_name

        return font_family
    except Exception as error:
        print(f"Error while fetching the para font family information: {error}")


# check if para is bold
def get_para_bold(document, para):
    """ Function to identify if a paragraph contains bold content
    Input:
    - document: Python-docx document object
    - para: Paragraph object
    Output
    - is_bold: Boolean indicating if the paragraph is bold
    """
    try:
        # Default is False (not bold)
        is_bold = False
        # Check to see if the para object is None
        if not para is None:
            # Create a list to store the bold values for each of the runs
            run_values = []
            for run in para.runs:
                if not run.bold is None:
                    run_values.append(run.bold)


            # Check paragraph style
            para_bold = para.style.font.bold

            # Check document style
            document_style_bold = None
            para_style = get_para_style(para)
            # Check the style from the document styles
            if para_style is not None:
                document_style_bold = document.styles[para_style].font.bold

            # Update is_bold variable based on the following logic:
            # - check if the paragraph style is bold
            # - Check the default value for the paragraph style at document level
            # - check if all runs are bold, if they are assign is_bold to True
            if para_bold is None:
                if document_style_bold is None:
                    if len(run_values) > 0:
                        if all(run_bold is True for run_bold in run_values):
                            is_bold = True
                else:
                    is_bold = document_style_bold
            else:
                is_bold = para_bold

        return is_bold
    except Exception as error:
        print(f"Error while fetching the bold information from para: {error}")

# check if para is italic
def get_para_italic(document, para):
    """ Function to identify if a paragraph contains italic content
    Input:
    - document: Python-docx document object
    - para: Paragraph object
    Output
    - is_italic: Boolean indicating if the paragraph is italic
    """
    try:
        # Default is False (not italic)
        is_italic = False
        full_run_complete = True
        # Check to see if the para object is None
        if not para is None:
            # Create a list to store the italic values for each of the runs
            run_values = []
            for run in para.runs:
                if not run.italic is None:
                    run_values.append(run.italic)
                else:
                    full_run_complete = False
                    break


            # Check paragraph style
            para_italic = para.style.font.italic

            # Check document style
            document_style_italic = None
            para_style = get_para_style(para)
            # Check the style from the document styles
            if para_style is not None:
                document_style_italic = document.styles[para_style].font.italic

            # Update is_italic variable based on the following logic:
            # - check if the paragraph style is italic
            # - Check the default value for the paragraph style at document level
            # - check if all runs are italic, if they are assign is_italic to True
            if para_italic is None:
                if document_style_italic is None:
                    if len(run_values) > 0 and full_run_complete:
                        if all(run is True for run in run_values):
                            is_italic = True
                else:
                    is_italic = document_style_italic
            else:
                is_italic = para_italic

        return is_italic
    except Exception as error:
        print(f"Error while fetching the italic information from para: {error}")



# get para font size 
def get_para_font_size(document, para):
    """ Function to retrieve the font size for a paragraph based on the style hierarchy
    Input:
    - document: document object for the entire docx file
    - para: paragraph object for the current paragraph
    Output:
    - font_size: Number indicating the font size, default to 11 - Word default font size
    """
    try:
        # Default font size to be 11
        font_size = 11
        if not document is None and not para is None:
            ## Gather all the raw data, follow style hierarchy as defined by
            #https://stackoverflow.com/questions/64031644/how-to-get-a-style-value-by-traversing
            #-from-bottom-run-to-top-docdefaults

            # Direct font values for the run and style values for the run
            run_font_values = []
            run_style_font_values = []
            # Iterate through each run to populate the two above lists
            for run in para.runs:
                if not run.text == "" and not run.text == "\n":
                    run_font_values.append(run.font.size)
                    if run.style is not None:
                        run_style_font_values.append(run.style.font.size)
            # Check the font size from the paragraph style
            paragraph_style_size = para.style.font.size

            # Identify the paragraph style
            document_para_style = None
            para_style = get_para_style(para)
            # Check the style from the document styles
            if para_style is not None:
                document_para_style = document.styles[para_style].font.size

            # Check if all the run font values are None
            if all(run is None for run in run_font_values):
                # Check if all the run style font values are None
                if all(run is None for run in run_style_font_values):
                    # Check if the paragraph style is none
                    if paragraph_style_size is None:
                        # Check if the style as defined in styles.xml has a size
                        if not document_para_style is None:
                            font_size = document_para_style.pt
                    else:
                        font_size = paragraph_style_size.pt
                else:
                    # Get unique set of font sizes identified in the run
                    font_sizes = [font_size for font_size in run_style_font_values
                                    if not font_size is None]
                    if font_sizes:
                        font_size = font_sizes[0].pt
            else:
                # Get unique set of font sizes identified in the run
                font_sizes = [font_size for font_size in run_font_values if not font_size is None]
                if font_sizes:
                    font_size = font_sizes[0].pt

        return font_size
    except Exception as error:
        print(f"Error while fetching the font size information from para: {error}")

# get para style
def get_para_style(para):
    """ Function to retrieve the font paragraph style for a paragraph based on the style hierarchy
    Input:
    - para: paragraph object for the current paragraph
    Output:
    - para_style: The style that is applied to the paragraph, defaults to 'Normal'
    """
    # Set the default to be Normal style
    para_style = "Normal"

    # Check the para style object
    if not para.style is None:
        para_style = para.style.name

    return para_style


# get para list style 
def get_para_list_style(document, para, numbering_pd):
    """ Function to return the paragraph list style, default to '' for now """
    try:
        para_list_style = ""
        document_pPr = None

        # Retrieve the paragraph style
        para_style = get_para_style(para)

        # Get any paragraph properties associated with the paragraph, either at the style
        # or paragraph level
        if para_style is not None:
            document_pPr = document.styles[para_style]._element.pPr
        para_pPr = para._p.pPr

        num_id = -1
        level = 0
        # Check the paragraph level first
        # Extract out the number id for the list
        if not para_pPr is None:
            if not para_pPr.numPr is None:
                num_id = str(para_pPr.numPr.numId.val)
                # Identify the Level, if any
                if not para_pPr.numPr.ilvl is None:
                    level = str(para_pPr.numPr.ilvl.val)

            elif not document_pPr is None:
                if not document_pPr.numPr is None:
                    num_id = str(document_pPr.numPr.numId.val)
                    # Identify the Level, if any
                    if not document_pPr.numPr.ilvl is None:
                        level = str(document_pPr.numPr.ilvl.val)


        # Identify the paragraph list style from the numbering_pd DataFrame
        numbering_para_list_style = get_data_from_numbering_pd(
            numbering_pd,
            num_id, level,
            column = "level_num_format"
        )

        # if len(numbering_para_list_style) > 0:
        #     para_list_style = numbering_para_list_style
        
        #######  this is for testing the project 500200 #####
        # if len(str(numbering_para_list_style)) > 0:
        para_list_style = numbering_para_list_style

        return para_list_style
    except Exception as error:
        print(f"Error while fetching the list style information from para: {error}")

# get para left indentation
def get_para_left_indent(document, para, numbering_pd):
    """ Function to find the left indent for a paragraph, default to 0
    Input:
    - document: Python-docx document object
    - para: Paragraph object for the current paragraph
    Output:
    - para_left_indent: Number indicating the left indent
    """
    try:
        para_left_indent = 0
        document_para_style = None
        document_pPr = None
        # Get the paragraph style
        para_style = get_para_style(para)
        # Check the style from the document styles
        if para_style is not None:
            document_para_style = document.styles[para_style].paragraph_format.left_indent
        # Check for numbering.xml when lists are used in Document
            document_pPr = document.styles[para_style]._element.pPr
        para_pPr = para._p.pPr

        numbering_left_indent = get_left_indent_from_numbering_pd(
            numbering_pd, document_pPr, para_pPr
        )

        if not para is None:
            if not para.paragraph_format.left_indent is None:
                para_left_indent = para.paragraph_format.left_indent.pt
                if not numbering_left_indent is None:
                    para_left_indent -= numbering_left_indent
            elif not para.style.paragraph_format.left_indent is None:
                para_left_indent = para.style.paragraph_format.left_indent.pt
            elif not document_para_style is None:
                para_left_indent = document.styles[para_style].paragraph_format.left_indent.pt
            elif not numbering_left_indent is None:
                para_left_indent = numbering_left_indent

        return para_left_indent
    except Exception as error:
        print(f"Error while fetching the left indent information from para: {error}")


# get para left indentation from numbering_pd
def get_left_indent_from_numbering_pd(numbering_pd, document_pPr, para_pPr):
    """ Function to retrieve the left_indent from the numbering_pd
    Args:
        numbering_pd (pandas DataFrame): Pandas DataFrame representing the numbering.xml file
        document_pPr (python-docx document object): Python-docx document object
        para_pPr (python-docx Paragraph properties object): python-docx Paragraph properties object
    Returns:
        [int]: Integer representing the left_indent for a paragraph, or None if none found
    """
    try:
        para_left_indent = None

        if not para_pPr is None:
            # Extract out the number id for the list
            num_id = -1
            level = 0
            if not para_pPr.numPr is None:
                num_id = str(para_pPr.numPr.numId.val)
                # Identify the Level, if any
                if not para_pPr.numPr.ilvl is None:
                    level = str(para_pPr.numPr.ilvl.val)
            elif not document_pPr is None:
                # Extract out the number id for the list
                if not document_pPr.numPr is None:
                    num_id = str(document_pPr.numPr.numId.val)
                    # Identify the Level, if any
                    if not document_pPr.numPr.ilvl is None:
                        level = str(document_pPr.numPr.ilvl.val)

            # Identify the left_indent from the numbering_pd DataFrame
            numbering_left_indent = get_data_from_numbering_pd(
                numbering_pd,
                num_id, level,
                column = "level_para_prop_left")

            # Convert to Pts if there is any value
            if len(str(numbering_left_indent)) > 0:
                # para_left_indent = int(numbering_left_indent) / 20
                #######  this is for testing the project 500200 #####
                para_left_indent = float(numbering_left_indent) / 20
            

        return para_left_indent
    except Exception as error:
        print(f"Error while fetching the left indent information from list para: {error}")


# get para right indentation
def get_para_right_indent(document, para):
    """ Function to find the right indent for a paragraph, default to 0
    Input:
    - document: Python-docx document object
    - para: Paragraph object for the current paragraph
    Output:
    - para_right_indent: Number indicating the right indent
    """
    try:
        para_right_indent = 0
        document_para_style = None
        # Get the paragraph style
        para_style = get_para_style(para)
        # Check the style from the document styles
        if para_style is not None:
            document_para_style = document.styles[para_style].paragraph_format.right_indent

        if not para is None:
            if not para.paragraph_format.right_indent is None:
                para_right_indent = para.paragraph_format.right_indent.pt
            elif not para.style.paragraph_format.right_indent is None:
                para_right_indent = para.style.paragraph_format.right_indent.pt
            elif not document_para_style is None:
                para_right_indent = document.styles[para_style].paragraph_format.right_indent.pt

        return para_right_indent
    except Exception as error:
        print(f"Error while fetching the right indent information from para: {error}")


# get para first line indentation
def get_para_first_line_indent(document, para):
    """ Function to retrieve the first line indent for a paragraph
    Input:
    - document: Python-docx document object
    - para: paragraph object for the current paragraph
    Output:
    - first_line_indent: The length of the first line indent, default to 0, expressed in points
    """
    try:
        first_line_indent = 0
        document_para_style = None

        # Get the paragraph style
        para_style = get_para_style(para)
        # Check the style from the document styles
        if para_style is not None:
            document_para_style = document.styles[para_style].paragraph_format.first_line_indent

        if not para is None:
            if not para.paragraph_format.first_line_indent is None:
                first_line_indent = para.paragraph_format.first_line_indent.pt
            elif not para.style.paragraph_format.first_line_indent is None:
                first_line_indent = para.style.paragraph_format.first_line_indent.pt
            elif not document_para_style is None:
                first_line_indent = document.styles[para_style].paragraph_format.first_line_indent.pt

        return first_line_indent
    except Exception as error:
            print(f"Error while fetching the right indent information from para: {error}")


# get para alignment
def get_para_alignment(document, para):
    """ Function to identify the alignment of the para
    Input:
    - document: Python-docx document object
    - para: Paragraph object in which the alignment is to be identified
    Output:
    - para_alignment: String containing the alignment of the paragraph para, default to "LEFT"
    """
    para_alignment = "left"

    if not para is None:
        # Paragraph alignment direct to the paragraph format
        para_format_alignment = para.paragraph_format.alignment
        # Paragraph alignment direct
        para_alignment_xml = para.alignment
        # Paragraph alignment from style
        para_style_alignment = para.style.paragraph_format.alignment
        # Document level style default
        document_para_style = document.styles[get_para_style(para)].paragraph_format.alignment

        if not para_format_alignment is None:
            para_alignment = para_format_alignment
        elif not para_alignment_xml is None:
            para_alignment = para_alignment_xml
        elif not para_style_alignment is None:
            para_alignment = para_style_alignment
        elif not document_para_style is None:
            para_alignment = document_para_style

        # Iterate through the inbuilt python docx values to get the original XML value
        for class_object in docx.enum.text.WD_PARAGRAPH_ALIGNMENT.__members__:
            if class_object.value == para_alignment:
                para_alignment = class_object.xml_value

    return para_alignment


# get para line space
def get_para_line_space(document, para):
    """ Function to identify the line spacing of a paragraph
    Input:
    - document: Python-docx document object
    - para: Paragraph object in which the alignment is to be identified
    Output:
    - para_line_space: Numeric indicating the line spacing
    """
    para_line_space = 1.15

    if not para is None:
        # Paragraph alignment direct to the paragraph format
        para_format = para.paragraph_format.line_spacing
        # Paragraph alignment from style
        para_style = para.style.paragraph_format.line_spacing
        # Document level style default
        document_para_style = document.styles[get_para_style(para)].paragraph_format.line_spacing

        if not para_format is None:
            para_line_space = para_format
        elif not para_style is None:
            para_line_space = para_style
        elif not document_para_style is None:
            para_line_space = document_para_style

    return para_line_space


# get space above para
def get_para_space_above(document, para):
    """ Function to identify the spacing above of a paragraph
    Input:
    - document: Python-docx document object
    - para: Paragraph object in which the alignment is to be identified
    Output:
    - para_space_above: Numeric indicating the space above
    """
    para_space_above = 0

    if not para is None:
        # Paragraph alignment direct to the paragraph format
        para_format = para.paragraph_format.space_before
        # Paragraph alignment from style
        para_style = para.style.paragraph_format.space_before
        # Document level style default
        document_para_style = document.styles[get_para_style(para)].paragraph_format.space_before

        if not para_format is None:
            para_space_above = para_format.pt
        elif not para_style is None:
            para_space_above = para_style.pt
        elif not document_para_style is None:
            para_space_above = document_para_style.pt

    return para_space_above


# get space below para
def get_para_space_below(document, para):
    """ Function to identify the spacing below of a paragraph
    Input:
    - document: Python-docx document object
    - para: Paragraph object in which the alignment is to be identified
    Output:
    - para_space_below: Numeric indicating the space below
    """
    para_space_below = 0

    if not para is None:
        # Paragraph alignment direct to the paragraph format
        para_format = para.paragraph_format.space_after
        # Paragraph alignment from style
        para_style = para.style.paragraph_format.space_after
        # Document level style default
        document_para_style = document.styles[get_para_style(para)].paragraph_format.space_after

        if not para_format is None:
            para_space_below = para_format.pt
        elif not para_style is None:
            para_space_below = para_style.pt
        elif not document_para_style is None:
            para_space_below = document_para_style.pt

    return para_space_below


# get para border
def get_para_border(para_prop_dict, para):
    """Function to retrieve the Paragraph Border Properties from XML
    Args:
        para_prop_dict (dict): Dictionary results should be appended to
        para (object): Paragraph object from the document object
    """
    ## Border Values - Top
    para_prop_dict["ParaBorderTopVal"] = get_xml_attribute(para,
        tag_to_find = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}top",
        tag_parent = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}pBdr",
        tag_attribute = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val",
        default_value = 0)

    para_prop_dict["ParaBorderTopSz"] = get_xml_attribute(para,
        tag_to_find = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}top",
        tag_parent = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}pBdr",
        tag_attribute = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}sz",
        default_value = 0)

    para_prop_dict["ParaBorderTopSpace"] = get_xml_attribute(para,
        tag_to_find = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}top",
        tag_parent = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}pBdr",
        tag_attribute = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}space",
        default_value = 0)

    para_prop_dict["ParaBorderTopColor"] = get_xml_attribute(para,
        tag_to_find = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}top",
        tag_parent = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}pBdr",
        tag_attribute = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}color",
        default_value = -1)

    ## Border Values - Left
    para_prop_dict["ParaBorderLeftVal"] = get_xml_attribute(para,
        tag_to_find = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}left",
        tag_parent = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}pBdr",
        tag_attribute = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val",
        default_value = 0)

    para_prop_dict["ParaBorderLeftSz"] = get_xml_attribute(para,
        tag_to_find = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}left",
        tag_parent = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}pBdr",
        tag_attribute = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}sz",
        default_value = 0)

    para_prop_dict["ParaBorderLeftSpace"] = get_xml_attribute(para,
        tag_to_find = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}left",
        tag_parent = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}pBdr",
        tag_attribute = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}space",
        default_value = 0)

    para_prop_dict["ParaBorderLeftColor"] = get_xml_attribute(para,
        tag_to_find = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}left",
        tag_parent = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}pBdr",
        tag_attribute = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}color",
        default_value = -1)

    ## Border Values - Bottom
    para_prop_dict["ParaBorderBottomVal"] = get_xml_attribute(para,
        tag_to_find = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}bottom",
        tag_parent = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}pBdr",
        tag_attribute = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val",
        default_value = 0)

    para_prop_dict["ParaBorderBottomSz"] = get_xml_attribute(para,
        tag_to_find = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}bottom",
        tag_parent = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}pBdr",
        tag_attribute = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}sz",
        default_value = 0)

    para_prop_dict["ParaBorderBottomSpace"] = get_xml_attribute(para,
        tag_to_find = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}bottom",
        tag_parent = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}pBdr",
        tag_attribute = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}space",
        default_value = 0)

    para_prop_dict["ParaBorderBottomColor"] = get_xml_attribute(para,
        tag_to_find = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}bottom",
        tag_parent = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}pBdr",
        tag_attribute = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}color",
        default_value = -1)

    ## Border Values - Right
    para_prop_dict["ParaBorderRightVal"] = get_xml_attribute(para,
        tag_to_find = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}right",
        tag_parent = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}pBdr",
        tag_attribute = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val",
        default_value = 0)

    para_prop_dict["ParaBorderRightSz"] = get_xml_attribute(para,
        tag_to_find = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}right",
        tag_parent = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}pBdr",
        tag_attribute = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}sz",
        default_value = 0)

    para_prop_dict["ParaBorderRightSpace"] = get_xml_attribute(para,
        tag_to_find = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}right",
        tag_parent = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}pBdr",
        tag_attribute = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}space",
        default_value = 0)

    para_prop_dict["ParaBorderRightColor"] = get_xml_attribute(para,
        tag_to_find = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}right",
        tag_parent = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}pBdr",
        tag_attribute = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}color",
        default_value = -1)

    ## Border Values - Between
    para_prop_dict["ParaBorderBetweenVal"] = get_xml_attribute(para,
        tag_to_find = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}between",
        tag_parent = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}pBdr",
        tag_attribute = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val",
        default_value = 0)

    para_prop_dict["ParaBorderBetweenSz"] = get_xml_attribute(para,
        tag_to_find = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}between",
        tag_parent = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}pBdr",
        tag_attribute = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}sz",
        default_value = 0)

    para_prop_dict["ParaBorderBetweenSpace"] = get_xml_attribute(para,
        tag_to_find = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}between",
        tag_parent = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}pBdr",
        tag_attribute = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}space",
        default_value = 0)

    para_prop_dict["ParaBorderBetweenColor"] = get_xml_attribute(para,
        tag_to_find = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}between",
        tag_parent = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}pBdr",
        tag_attribute = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}color",
        default_value = -1)

    return para_prop_dict


# get para shading
def get_para_shading(para_prop_dict, para):
    """Function to retrieve the Paragraph Shading Properties from XML
    Args:
        para_prop_dict (dict): Dictionary results should be appended to
        para (object): Paragraph object from the document object
    """
    if not para is None:
        # Get the value of the amount of percent
        para_prop_dict["ParaShadingVal"] = get_xml_attribute(para,
            tag_to_find = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}shd",
            tag_parent = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}pPr",
            tag_attribute = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val"
        )

        # Get the colour of the shading
        para_prop_dict["ParaShadingColor"] = get_xml_attribute(para,
            tag_to_find = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}shd",
            tag_parent = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}pPr",
            tag_attribute = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}color"
        )

        # Get the fill of the shading
        para_prop_dict["ParaShadingFill"] = get_xml_attribute(para,
            tag_to_find = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}shd",
            tag_parent = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}pPr",
            tag_attribute = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}fill"
        )

    return para_prop_dict


# get para single strike
def get_para_single_strike(para):
    """ Function to return if a paragraph is all single striked through
    All runs within the paragraph should the single striked through in order to
    return True. Only runs which have text are considered.
    Args:
        para (python-docx Paragraph): Python-docx paragraph object
    Returns:
        boolean: True/False wheather all runs which contain text in the paragraph are
        single striked throug
    """
    # empty list to store the run strike
    run_strike = []
    # iterate over each para runs 
    for run in para.runs:
        # check if its not empty
        if not run.text == "":
            # append into run strike list
            run_strike.append(run.font.strike)

    return all(run_strike)


# get para double strike
def get_para_double_strike(para):
    """ Function to return if a paragraph is all double striked through
    All runs within the paragraph should the double striked through in order to
    return True. Only runs which have text are considered.
    Args:
        para (python-docx Paragraph): Python-docx paragraph object
    Returns:
        boolean: True/False wheather all runs which contain text in the paragraph are
        double striked throug
    """
    # empty list to store the run strike
    run_strike = []
    # iterate over each para runs 
    for run in para.runs:
        # check if its not empty
        if not run.text == "":
            # append into run strike list
            run_strike.append(run.font.double_strike)

    return all(run_strike)


# get underline para
def get_para_underline(document, para):
    """[summary]
    Args:
        document (python_docx Document): Python-docx Document object
        para (python-docx Paragraph): Python-docx Paragraph object
    Returns:
        [str]: String containing the underline value for the paragraph
    """
    # underline empty string as an placeholder
    underline = ""

    # Get the Underline values for each run in the paragraph
    # empty list to store run values
    run_values = []
    # iterate over each para runs
    for run in para.runs:
        # check if its not empty 
        if not run.text == "":
            # append into run_values list
            run_values.append(run.font.underline)

    # Check if there is only one run_value
    if len(run_values) == 1:
        # If this single value is True, set underline to be Single
        if run_values[0] is True:
            underline = "single"
        else:
            para_prop_underline = run_values[0]
            for underline_class in docx.enum.text.WD_UNDERLINE.__members__:
                if underline_class.value == para_prop_underline:
                    underline = underline_class.xml_value

    # If underline is still blank after checking the runs, check the style
    if underline == "":
        # Get the paragraph style
        style_underline = None
        para_style = get_para_style(para)

        # Check the font underline in the document style
        if para_style is not None:
            style_underline = document.styles[para_style].font.underline

        # If its not None, then update the return value
        if not style_underline is None:
            # Iterate throught the UNDERLINE class
            for underline_class in docx.enum.text.WD_UNDERLINE.__members__:
                # Check if the display value is equal to the value we have gotten
                if underline_class.value == style_underline:
                    # Update the output with the xml value
                    underline = underline_class.xml_value

    return underline


# get para small caption
def get_para_small_caps(para):
    """ Function to identify if the para has small caps enabled
    Args:
        para (python-docx Paragraph): Python-docx Paragraph object
    Returns:
        boolean: True/False if the paragraph has small caps enabled
    """
    # define output as placeholder
    output = "No Text"
    # empty list to store run values
    run_values = []

    # iterate over para.runs
    for run in para.runs:
        if not run.text == "":
            run_values.append(run.font.small_caps)

    # check if list is not empty
    if len(run_values) > 0:
        output = any(run_values)

    return output


# create structured xml after list of dict properties
def create_structured_xml(para_properties_list, c_data_tags):
    """ Function to transform a list of dictionary poperties into XML
    Input:
    - para_properties_list: List containing dictionary for each subelement, see
    create_para_properties_list
    - c_data_tags: List of strings containing the tags which should be treated as XML CDATA
    Output:
    - para_properties_xml: xml string of the para_properties_list
    """
    # para properties xml placeholder
    para_properties_xml = ""
    try:
        if len(para_properties_list) > 0:
            # Create root element
            root = etree.Element("ArrayOfParagraphProperties")

            # Loop through each dictionary in the list
            for para_dict in para_properties_list:
                # Create a parent node for the sub-child of root
                parent = etree.SubElement(root, "ParagraphProperties")

                # Loop through each key, value pair in the dictionary
                for key, value in para_dict.items():
                    # Check if the value is None, if it is then no text to be added
                    if value is None:
                        etree.SubElement(parent, key)
                    else:
                        # Condition when there is text to be added for a sub-child of parent
                        # Check if the current key is of CDATA type
                        if key in c_data_tags:
                            etree.SubElement(parent, key).text = etree.CDATA(str(value))
                        else:
                            etree.SubElement(parent, key).text = str(value)

                # Append the parent to the root
                root.append(parent)

        # Convert to a UTF-8 encoded string, and add in the initial xml declaration
        para_properties_xml = etree.tostring(root, xml_declaration = True,
                                            encoding = "UTF-8").decode("UTF-8")

        return para_properties_xml
    except Exception as error:
        print(f"Error occured while writing the property extraction into XML: {error}")

# generate random id
def gen_id():
    return str(uuid.uuid4().hex)[0:7].upper()


if __name__ == "__main__":
    doc_xml = extract_docx_properties("/Users/senthil/Downloads/RL_01_TRUT_C001.docx")
    
print (doc_xml)