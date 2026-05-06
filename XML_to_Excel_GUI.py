import tkinter as tk
import pandas as pd
import os
import logging
import xml.etree.ElementTree as ET
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
logging.basicConfig(filename='myProgramLog.txt', filemode='w', level=logging.DEBUG, format=' %(asctime)s - %(levelname)s- %(message)s')

window = tk.Tk()
window.title("XML to Excel Converter")
window.minsize(300, 120)
window.columnconfigure(0, weight=1, minsize=75)
window.rowconfigure(0, weight=1, minsize=50)

#funciton to convert to excel
def convert_to_excel():
    logging.info('Convert to Excel button clicked. Starting conversion process...')
    logging.debug('Checking validity of file path...')
    if(check_validity()):
        #build dataframe from XML based on list
        tree = ET.parse(ent_file.get())
        root = tree.getroot()

        logging.info('Parsing XML file and building dataframe...')
        dfList = [] #list of all lists to be written to CSV file
        headerList = []#list of all headers; Columns of CSV file
        get_headers(root, headerList) 

        #get shortened header names
        logging.info('Getting shortened header names...')
        headerMap = get_unique_headers(headerList)
        shortHeaderList = [''] * len(headerList)
        for fullPath, shortened in headerMap.items():
            logging.debug(f'Adding header {shortened} to headerList')
            shortHeaderList[headerList.index(fullPath)] = shortened

        logging.info('---------------Headers created, populating children-------------')
        i = 1
        for child in root:
            childList = [''] * len(headerList)
            process_child(child, '', childList, headerList,)
            logging.info(f'Block {i} processed: {childList}')
            i = i + 1
            dfList.append(childList)

        df = pd.DataFrame(dfList, columns=shortHeaderList)
        logging.debug(f'Dataframe created from XML file. Dateframe created: {df}')

        wb = Workbook()
        ws = wb.active
        logging.info('Populating workbook with data...')
        for row in dataframe_to_rows(df, index=False, header=True):
            ws.append(row)
        
        #get output file name from parent XML file and save workbook with that name
        XMLFileName = os.path.basename(ent_file.get())
        logging.info(f'Saving workbook with name {XMLFileName.split(".")[0]}.xlsx')
        wb.save(f'{XMLFileName.split(".")[0]}.xlsx')
        lbl_file_created["text"] = f"File created: {XMLFileName.split('.')[0]}.xlsx"
    else:
        lbl_valid["text"] = "Please enter a valid XML file path!"

#Checks validity of file path, returning true/false and logging any errors
def check_validity():
    directory = ent_file.get()
    validFile = False
    if(os.path.isfile(directory)):
        if(os.path.splitext(directory)[1].lower() == '.xml'):
            validFile = True
        else:
            basename = os.path.basename(directory)
            logging.warning('User tried using this file: %s.', basename)
    else:
        logging.warning('User gave an invalid file path: %s.', directory)

    return validFile

#XML parsing functions
def get_headers(treeRoot, headerList):
    logging.info('Getting Headers')
    for child in treeRoot:
        build_header_children(child, '', headerList)
    headerList = sorted(headerList)
    logging.debug(f'Header list: {headerList}')

#Processes a singular child to get headers; Allows for recursion
def build_header_children(root, path, headerList):
    currentPath = f'{path}/{root.tag}' if path else root.tag
    if(len(root) == 0): #if leaf node, process header
        if(currentPath not in headerList):
            logging.debug(f'Adding header {root.tag} to headerList')
            headerList.append(f'{currentPath}')
        else:
            #handle duplicate headers
            logging.debug(f'Duplicate tag found: {root.tag}')
    else:
        for child in root:
            #dont process any text element; skip to processing subtrees
            logging.debug(f'Navigating subtree {child.tag}; Text of parent left behind is {root.text} and tag is {root.tag}')
            build_header_children(child, currentPath, headerList)
        
#Processes a child element, recursively going though it's children as well
def process_child(root, path, childList, headerList):
    currentPath = f'{path}/{root.tag}' if path else root.tag
    if(len(root) == 0): #if leaf node, store info
        logging.debug(f'Logging leaf node {root.text} at {currentPath}')
        text = root.text.strip() if root.text is not None else ''
        if not text:
            logging.debug(f'Path was empty / whitespace, returning')
            return
        try:
            insertionIndex = headerList.index(currentPath)
        except ValueError as e:
            logging.error(f'{e}: header {currentPath} not found, but has leaf node. Continuing...')
            return
                    #if list at the index is empty, add text; If text is already there, instead append with new info
        childList[insertionIndex] = text if not childList[insertionIndex] else f'{childList[insertionIndex]}, {text}'
    else:
        for child in root:
            #dont process any text element; skip to processing subtrees
            logging.debug(f'Navigating subtree {child.tag}; Text of parent left behind is {root.text} and tag is {root.tag}')
            process_child(child, currentPath, childList, headerList)

def get_unique_headers(fullPaths):
    #Convert full paths to shortest unique suffixes.
    #For each path, find the shortest suffix that uniquely identifies it.
    headers = {}  # maps original full path to final header name
    
    for fullPath in fullPaths:
        parts = fullPath.split('/')
        
        # Try from shortest to longest suffix
        for suffix_length in range(1, len(parts) + 1):
            suffix = '/'.join(parts[-suffix_length:]) #joins last <suffix_length> elements of parts with slashes to create suffix
            
            # Check if another path would also generate this same suffix at this length
            conflict = False
            for otherPath in fullPaths:
                if otherPath == fullPath:
                    continue
                otherParts = otherPath.split('/')
                if suffix_length <= len(otherParts):
                    otherSuffix = '/'.join(otherParts[-suffix_length:])
                    if otherSuffix == suffix:
                        conflict = True
                        break
            if not conflict:
                headers[fullPath] = suffix
                break
        
        # If no unique suffix found, use full path as fallback
        if fullPath not in headers:
            headers[fullPath] = fullPath
    return headers

#set up outer frame
frm = tk.Frame(window, bg='lightgray')

#set up and insert widgets as they appear in the GUI
lbl_prompt_file = tk.Label(frm, text="Enter the name of the XML file to be summarized:", relief='raised', highlightbackground='black', highlightthickness=1)

ent_file = tk.Entry(frm)

lbl_valid = tk.Label(frm, text="")

btn_convert = tk.Button(frm, text="Convert to Excel", command=convert_to_excel)

lbl_file_created = tk.Label(frm, text="")

#insert into frames and window
lbl_prompt_file.grid(row=0, column=0)
ent_file.grid(row=1, column=0)
lbl_valid.grid(row=2, column=0, sticky="n")
btn_convert.grid(row=3, column=0)
lbl_file_created.grid(row=4, column=0)

frm.grid(row=0, column=0, padx=10, pady=10)

logging.debug('GUI setup complete. Starting main loop...')

window.mainloop()