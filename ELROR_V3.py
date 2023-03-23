#=== IMPORTS =================================================================
from __future__ import print_function
import warnings
import logging
import extract_msg
import pandas as pd
import hashlib
import datetime
import glob
import os
import shutil
from tqdm import tqdm 
import sys
import re
import zlib 
import struct
import io
import olefile
from oletools.thirdparty.tablestream import tablestream
from oletools import crypto, ftguess, olevba, mraptor, oleobj, ooxml
from oletools.common.log_helper import log_helper
from oletools.common.codepages import get_codepage_name
import datetime
import csv
from pathlib import Path
import PyPDF2
import docx2txt
from pptx import Presentation 


#=== IMPORTS =================================================================
def logo():
    print("")
    print("""
    ░▄▄▄▄░
    ▀▀▄██►
    ▀▀███►
    ░▀███►░█►
    ▒▄████▀▀                         
    """)
    print("")
logo()

# Step 1 
warnings.simplefilter(action='ignore', category=FutureWarning)
fail = 0
count_msg = 0 
count_attachment = 0

__version__ = '0.60.1' # Version to OLEID

msg_output = sys.argv[1]   
output_path =sys.argv[2]        

msg =  [f for f in os.listdir() if '.msg' in f.lower()]

if not os.path.exists(msg_output):
    os.makedirs(msg_output)

for image in msg:
    new_path = '' + image
    shutil.move(image, new_path)

# specify the directory and file extension
file_extension = '*.msg'

file_list = glob.glob(f'{msg_output}/{file_extension}')

 # Add where you want the output files here 

current_date = datetime.datetime.now()
current_month = current_date.strftime("%B")
current_year = current_date.strftime("%Y")

tool_name = "ELROR" 

# This saves the  files

output_folder_name = f"{tool_name}_{current_year}_{current_month}_CSV_Output"
output_folder_path = os.path.join(output_path, output_folder_name)
if not os.path.exists(output_folder_path):
    os.makedirs(output_folder_path)

# This is is the file path for the OLE TOOLS at the bottom of the script
csv_name_OLE = f"{tool_name}_{current_year}_{current_month}_Oletools_CSV_Output.csv"

# Creates the csv for the Url 
csv_name_Url = f"{tool_name}_{current_year}_{current_month}_Url_CSV_Output.csv"
csv_path_Url = os.path.join(output_folder_path, csv_name_Url)
# save the dataframe to the CSV file to the updated storage 
csv_path_OLE = os.path.join(output_folder_path, csv_name_OLE)

# get the number of files in the list
num_files = len(file_list)

# loop through each file and extracted the field values
columns = shutil.get_terminal_size().columns
print("Loading".center(columns))

# set up progress bar
file_list_len = len(file_list)
pbar = tqdm(total=file_list_len, desc='Processing files', unit='file')

# initialize variables
count_msg = 0
count_attachment = 0
fail = 0

# initialize data frame
email_data = pd.DataFrame(columns=["Sender", "Subject", "Hash(256)", "Attachments", "Attachment Hash(256)"])

for file in file_list:
    try :
        msg = extract_msg.Message(file)
        msg_sender = msg.sender
        msg_date = msg.date
        msg_subj = msg.subject
        msg_message = msg.body
        msg_attachment = msg.attachments
        
        # Hash the message content using the hashlib library
        with open(file, 'rb') as msg_file:
            msg_text = msg_file.read()
            
        h = hashlib.sha256(msg_text).hexdigest()
        
        # create folder with hash name
        folder_path = os.path.join(output_folder_path, h)
        if not os.path.exists(folder_path):
            os.makedirs(folder_path)

        # move message file to folder with hash name
        msg_file_name = os.path.basename(file)
        msg_file_path = os.path.join(folder_path, msg_file_name)
        shutil.copyfile(file, msg_file_path)
        count_msg +=1
        
        attachment_hash_list = []
        attachment_filename_list = []
        attachment_folder_path = os.path.join(folder_path, 'Attachments')
        
        if msg_attachment:
            if not os.path.exists(attachment_folder_path):
                #Makes folder for storing the attachment
                os.makedirs(attachment_folder_path)
                
            for attachment in msg.attachments:
                #for loop 
                
                try: # To combat the errors 
                    attachment_filename = attachment.longFilename
                    attachment_data = attachment.data
                    attachment_file_path = os.path.join(attachment_folder_path, attachment_filename)
                    with open(attachment_file_path, 'wb') as f:
                        f.write(attachment_data)
                    attachment_hash = hashlib.sha256(attachment_data).hexdigest()
                    attachment_hash_list.append(attachment_hash)
                    attachment_filename_list.append(attachment_filename)
                    count_attachment +=1
                except Exception:
                    # handle error here
                    fail += 1
                    
                    
        # add the row to the Data frame 
        email_data = email_data.append({"Sender": msg_sender,
                                        "Subject": msg_subj ,
                                        "Hash(256)": h,
                                        "Attachments": attachment_filename_list ,
                                        "Attachment Hash(256)": attachment_hash_list
                                        }, ignore_index=True)
        
        df2=email_data.dropna(axis=1) # This drops the N/A felids 
        
    except Exception :
        fail+=1
    
    pbar.update(1)

pbar.close()
        
print("")
print("STATS".center(columns))
print("Failed to save attachment")
print("")
print(fail)
print("")
print("Total Number of .msg saved")
print("")
print(count_msg)
print("")
print("Total Number of attachments saved")
print("")
print(count_attachment)
print("")
success_rate = ((count_msg - fail) / count_msg) * 100
print(f"The success rate of msg are {success_rate:.2f}%")
print("")

# create the CSV filename
csv_name = f"{tool_name}_{current_year}_{current_month}_CSV_Output.csv"

# save the dataframe to the CSV file to the updated storage 
csv_path = os.path.join(output_folder_path, csv_name)

df2.to_csv(csv_path, index=False)

_thismodule_dir = os.path.normpath(os.path.abspath(os.path.dirname(__file__)))
# print('_thismodule_dir = %r' % _thismodule_dir)
_parent_dir = os.path.normpath(os.path.join(_thismodule_dir, '..'))
# print('_parent_dir = %r' % _thirdparty_dir)
if _parent_dir not in sys.path:
    sys.path.insert(0, _parent_dir)


tool_name = "ELROR" 
current_date = datetime.datetime.now()
current_month = current_date.strftime("%B")
current_year = current_date.strftime("%Y")


# === LOGGING =================================================================

log = log_helper.get_or_create_silent_logger('oleid')

# === CONSTANTS ===============================================================

class RISK(object):
    """
    Constants for risk levels
    """
    HIGH = 'HIGH'
    MEDIUM = 'Medium'
    LOW = 'low'
    NONE = 'none'
    INFO = 'info'
    UNKNOWN = 'Unknown'
    ERROR = 'Error'  # if a check triggered an unexpected error

risk_color = {
    RISK.HIGH: 'red',
    RISK.MEDIUM: 'yellow',
    RISK.LOW: 'white',
    RISK.NONE: 'green',
    RISK.INFO: 'cyan',
    RISK.UNKNOWN: None,
    RISK.ERROR: None
}

#=== FUNCTIONS ===============================================================

def detect_flash(data):
    #TODO: report
    found = []
    for match in re.finditer(b'CWS|FWS', data):
        start = match.start()
        if start+8 > len(data):
            # header size larger than remaining data, this is not a SWF
            continue
        #TODO: one struct.unpack should be simpler
        # Read Header
        header = data[start:start+3]
        # Read Version
        ver = struct.unpack('<b', data[start+3:start+4])[0]
        # Error check for version above 20
        #TODO: is this accurate? (check SWF specifications)
        if ver > 20:
            continue
        # Read SWF Size
        size = struct.unpack('<i', data[start+4:start+8])[0]
        if start+size > len(data) or size < 1024:
            # declared size larger than remaining data, this is not a SWF
            # or declared size too small for a usual SWF
            continue
        # Read SWF into buffer. If compressed read uncompressed size.
        swf = data[start:start+size]
        compressed = False
        if b'CWS' in header:
            compressed = True
            # compressed SWF: data after header (8 bytes) until the end is
            # compressed with zlib. Attempt to decompress it to check if it is
            # valid
            compressed_data = swf[8:]
            try:
                zlib.decompress(compressed_data)
            except Exception:
                continue
        # else we don't check anything at this stage, we only assume it is a
        # valid SWF. So there might be false positives for uncompressed SWF.
        found.append((start, size, compressed))
        #print 'Found SWF start=%x, length=%d' % (start, size)
    return found


#=== CLASSES =================================================================

class Indicator(object):
    """
    Piece of information of an :py:class:`OleID` object.

    Contains an ID, value, type, name and description. No other functionality.
    """

    def __init__(self, _id, value=None, _type=bool, name=None,
                 description=None, risk=RISK.UNKNOWN, hide_if_false=True):
        self.id = _id
        self.value = value
        self.type = _type
        self.name = name
        if name == None:
            self.name = _id
        self.description = description
        self.risk = risk
        self.hide_if_false = hide_if_false


class OleID(object):
    def __init__(self, filename=None, data=None):
        if filename is None and data is None:
            raise ValueError('OleID requires either a file path or file data, or both')
        self.file_on_disk = False  # True = file on disk / False = file in memory
        if data is None:
            self.file_on_disk = True  # useful for some check that don't work in memory
            with open(filename, 'rb') as f:
                self.data = f.read()
        else:
            self.data = data
        self.data_bytesio = io.BytesIO(self.data)
        if isinstance(filename, olefile.OleFileIO):
            self.ole = filename
            self.filename = None
        else:
            self.filename = filename
            self.ole = None
        self.indicators = []
        self.suminfo_data = None

    def get_indicator(self, indicator_id):
        """Helper function: returns an indicator if present (or None)"""
        result = [indicator for indicator in self.indicators
                  if indicator.id == indicator_id]
        if result:
            return result[0]
        else:
            return None

    def check(self):
        self.ftg = ftguess.FileTypeGuesser(filepath=self.filename, data=self.data)
        ftype = self.ftg.ftype
        # if it's an unrecognized OLE file, display the root CLSID in description:
        if self.ftg.filetype == ftguess.FTYPE.GENERIC_OLE:
            description = 'Unrecognized OLE file. Root CLSID: {} - {}'.format(
                self.ftg.root_clsid, self.ftg.root_clsid_name)
        else:
            description = ''
        ft = Indicator('ftype', value=ftype.longname, _type=str, name='File format', risk=RISK.INFO,
                       description=description)
        self.indicators.append(ft)
        ct = Indicator('container', value=ftype.container, _type=str, name='Container format', risk=RISK.INFO,
                       description='Container type')
        self.indicators.append(ct)

        # check if it is actually an OLE file:
        if self.ftg.container == ftguess.CONTAINER.OLE:
       
            self.check_properties()
            self.check_encrypted()
            self.check_macros()
            self.check_external_relationships()
            self.check_object_pool()
            self.check_flash()
            if self.ole is not None:
                self.ole.close()
        return self.indicators

    def check_properties(self):
        if not self.ole:
            return None
        meta = self.ole.get_metadata()
        appname = Indicator('appname', meta.creating_application, _type=str,
                            name='Application name', description='Application name declared in properties',
                            risk=RISK.INFO)
        self.indicators.append(appname)
        codepage_name = None
        if meta.codepage is not None:
            codepage_name = '{}: {}'.format(meta.codepage, get_codepage_name(meta.codepage))
        codepage = Indicator('codepage', codepage_name, _type=str,
                      name='Properties code page', description='Code page used for properties',
                      risk=RISK.INFO)
        self.indicators.append(codepage)
        author = Indicator('author', meta.author, _type=str,
                      name='Author', description='Author declared in properties',
                      risk=RISK.INFO)
        self.indicators.append(author)
        return appname, codepage, author

    def check_encrypted(self):
        encrypted = Indicator('encrypted', False, name='Encrypted',
                              risk=RISK.NONE,
                              description='The file is not encrypted',
                              hide_if_false=False)
        self.indicators.append(encrypted)
        # Only OLE files can be encrypted (OpenXML files are encrypted in an OLE container):
        if not self.ole:
            return None
        try:
            if crypto.is_encrypted(self.ole):
                encrypted.value = True
                encrypted.risk = RISK.LOW
                encrypted.description = 'The file is encrypted. It may be decrypted with msoffcrypto-tool'
        except Exception as exception:
            # msoffcrypto-tool can trigger exceptions, such as "Unknown file format" for Excel 5.0/95
            encrypted.value = 'Error'
            encrypted.risk = RISK.ERROR
            encrypted.description = 'msoffcrypto-tool raised an error when checking if the file is encrypted: {}'.format(exception)
        return encrypted

    def check_external_relationships(self):
        ext_rels = Indicator('ext_rels', 0, name='External Relationships', _type=int,
                              risk=RISK.NONE,
                              description='External relationships such as remote templates, remote OLE objects, etc',
                              hide_if_false=False)
        self.indicators.append(ext_rels)
        # this check only works for OpenXML files
        if not self.ftg.is_openxml():
            return ext_rels
        # to collect relationship types:
        rel_types = set()
        # open an XmlParser, using a BytesIO instead of filename (to work in memory)
        xmlparser = ooxml.XmlParser(self.data_bytesio)
        for rel_type, target in oleobj.find_external_relationships(xmlparser):
            log.debug('External relationship: type={} target={}'.format(rel_type, target))
            rel_types.add(rel_type)
            ext_rels.value += 1
        if ext_rels.value > 0:
            ext_rels.description = 'External relationships found: {} - use oleobj for details'.format(
                ', '.join(rel_types))
            ext_rels.risk = RISK.HIGH
        return ext_rels

    def check_object_pool(self):
        objpool = Indicator(
            'ObjectPool', False, name='ObjectPool',
            description='Contains an ObjectPool stream, very likely to contain '
                        'embedded OLE objects or files. Use oleobj to check it.',
            risk=RISK.NONE)
        self.indicators.append(objpool)
        if not self.ole:
            return None
        if self.ole.exists('ObjectPool'):
            objpool.value = True
            objpool.risk = RISK.LOW
            # TODO: set risk to medium for OLE package if not executable
            # TODO: set risk to high for Package executable or object with CVE in CLSID
        return objpool

    def check_macros(self):
        vba_indicator = Indicator(_id='vba', value='No', _type=str, name='VBA Macros',
                                  description='This file does not contain VBA macros.',
                                  risk=RISK.NONE, hide_if_false=False)
        self.indicators.append(vba_indicator)
        xlm_indicator = Indicator(_id='xlm', value='No', _type=str, name='XLM Macros',
                                  description='This file does not contain Excel 4/XLM macros.',
                                  risk=RISK.NONE, hide_if_false=False)
        self.indicators.append(xlm_indicator)
        if self.ftg.filetype == ftguess.FTYPE.RTF:
            # For RTF we don't call olevba otherwise it triggers an error
            vba_indicator.description = 'RTF files cannot contain VBA macros'
            xlm_indicator.description = 'RTF files cannot contain XLM macros'
            return vba_indicator, xlm_indicator
        vba_parser = None  # flag in case olevba fails
        try:
            vba_parser = olevba.VBA_Parser(filename=self.filename, data=self.data)
            if vba_parser.detect_vba_macros():
                vba_indicator.value = 'Yes'
                vba_indicator.risk = RISK.MEDIUM
                vba_indicator.description = 'This file contains VBA macros. No suspicious keyword was found. Use olevba and mraptor for more info.'
                # check code with mraptor
                vba_code = vba_parser.get_vba_code_all_modules()
                m = mraptor.MacroRaptor(vba_code)
                m.scan()
                if m.suspicious:
                    vba_indicator.value = 'Yes, suspicious'
                    vba_indicator.risk = RISK.HIGH
                    vba_indicator.description = 'This file contains VBA macros. Suspicious keywords were found. Use olevba and mraptor for more info.'
        except Exception as e:
            vba_indicator.risk = RISK.ERROR
            vba_indicator.value = 'Error'
            vba_indicator.description = 'Error while checking VBA macros: %s' % str(e)
        finally:
            if vba_parser is not None:
                vba_parser.close()
            vba_parser = None
        # Check XLM macros only for Excel file types:
        if self.ftg.is_excel():
            # TODO: for now XLM detection only works for files on disk... So we need to reload VBA_Parser from the filename
            #       To be improved once XLMMacroDeobfuscator can work on files in memory
            if self.file_on_disk:
                try:
                    vba_parser = olevba.VBA_Parser(filename=self.filename)
                    if vba_parser.detect_xlm_macros():
                        xlm_indicator.value = 'Yes'
                        xlm_indicator.risk = RISK.MEDIUM
                        xlm_indicator.description = 'This file contains XLM macros. Use olevba to analyse them.'
                except Exception as e:
                    xlm_indicator.risk = RISK.ERROR
                    xlm_indicator.value = 'Error'
                    xlm_indicator.description = 'Error while checking XLM macros: %s' % str(e)
                finally:
                    if vba_parser is not None:
                        vba_parser.close()
            else:
                xlm_indicator.risk = RISK.UNKNOWN
                xlm_indicator.value = 'Unknown'
                xlm_indicator.description = 'For now, XLM macros can only be detected for files on disk, not in memory'

        return vba_indicator, xlm_indicator

    def check_flash(self):
        flash = Indicator(
            'flash', 0, _type=int, name='Flash objects',
            description='Number of embedded Flash objects (SWF files) detected '
                        'in OLE streams. Not 100% accurate, there may be false '
                        'positives.',
            risk=RISK.NONE)
        self.indicators.append(flash)
        if not self.ole:
            return None
        for stream in self.ole.listdir():
            data = self.ole.openstream(stream).read()
            found = detect_flash(data)
            # just add to the count of Flash objects:
            flash.value += len(found)
            #print stream, found
        if flash.value > 0:
            flash.risk = RISK.MEDIUM
        return flash
#=== MAIN =================================================================

def main():
    # Set the directory containing the files
    
    log_helper.enable_logging()

    # Loop over all files in the directory
    for filename in os.listdir(msg_output):
        macro_ext = [".docm", ".xlsm", ".pptm", ".ppsm"]
        if filename.endswith(".docx") or filename.endswith(".msg") or filename.endswith(".vba") or any(filename.endswith(ext) for ext in macro_ext):

            # Open the file and run OleID on it
            filepath = os.path.join(msg_output, filename)
            
            oleid = OleID(filepath)
            indicators = oleid.check()

            # Write the results to a CSV file

            with open(csv_path_OLE, 'a', newline='', encoding='utf-8') as csvfile:
                writer = csv.writer(csvfile)
                for indicator in indicators:
                    if not (indicator.hide_if_false and not indicator.value):
                        #print '%s: %s' % (indicator.name, indicator.value)
                        color = risk_color.get(indicator.risk, None)
                        writer.writerow([filename, indicator.name, indicator.value, indicator.risk, indicator.description])
        
    
    for path in Path(output_folder_path).rglob('Attachments/'):
            path = os.path.abspath(path)
            directory = path.replace(os.sep, '/')
          
            
            for filename in os.listdir(directory):
                macro_ext = [".docm", ".xlsm", ".pptm", ".ppsm", ".pdf" , ".zip" , ".7zip"]
                if filename.endswith(".docx") or filename.endswith(".bas") or filename.endswith(".vba") or any(filename.endswith(ext) for ext in macro_ext):

                    # Open the file and run OleID on it
                    filepath = os.path.join(directory, filename)
                    
                    oleid = OleID(filepath)
                    indicators = oleid.check()

                    # Write the results to a CSV file

                    with open(csv_path_OLE, 'a', newline='', encoding='utf-8') as csvfile:
                        writer = csv.writer(csvfile)
                        for indicator in indicators:
                            if not (indicator.hide_if_false and not indicator.value):
                                #print '%s: %s' % (indicator.name, indicator.value)
                                color = risk_color.get(indicator.risk, None)
                                writer.writerow([filename, indicator.name, indicator.value, indicator.risk, indicator.description])
if __name__ == '__main__':
    main()

with open(os.path.join(output_folder_path, csv_path_Url), "w", newline="") as csvfile:
    csvwriter = csv.writer(csvfile)
    csvwriter.writerow(["Filename", "URL"])

    for path in Path(output_folder_path).rglob("Attachments/*.*"):
        path = os.path.abspath(path)
        directory = path.replace(os.sep, "/")
        extension = os.path.splitext(directory)[1]
        if extension == ".pdf":
            # Do something with PDF file
            with open(directory, "rb") as pdf_file:
                # Create a PDF reader object
                pdf_reader = PyPDF2.PdfReader(pdf_file)

                # Loop through each page in the PDF
                for page in pdf_reader.pages:
                    # Get the text content of the page
                    text = page.extract_text()

                    # Use regex to find all URLs in the text 
                    urls = re.findall('http[s]?://(?:[a-zA-Z]|[0-9]|[$-_@.&+]|[!*\(\),]|(?:%[0-9a-fA-F][0-9a-fA-F]))+',text)

                    # Write the URLs found in the current page to CSV
                    for url in urls:
                        csvwriter.writerow([os.path.basename(directory), url])
        elif extension == ".doc" or extension == ".docx":
            # Do something with Word file
            text = docx2txt.process(directory)

            # Use regex to find all URLs in the text
            urls = re.findall('http[s]?://(?:[a-zA-Z]|[0-9]|[$-_@.&+]|[!*\(\),]|(?:%[0-9a-fA-F][0-9a-fA-F]))+',text)

            # Write the URLs found in the current file to CSV
            for url in urls:
                csvwriter.writerow([os.path.basename(directory), url])
###
        
        elif extension == ".ppt" or extension == ".pptx":
            
            # Add bit for ppt and macro enabled one 
            
            urls = re.findall('http[s]?://(?:[a-zA-Z]|[0-9]|[$-_@.&+]|[!*\(\),]|(?:%[0-9a-fA-F][0-9a-fA-F]))+',text)
           
            for slide in Presentation.slides:
                for shape in slide.shapes:
                    if hasattr(shape, urls):
                        print(shape.urls)
            
            for url in urls:
                csvwriter.writerow([os.path.basename(directory), url])
####        
        elif extension == ".png" or extension == ".zip":
            print("")
            print("Unable to process these files")
            print("")
            print(directory)
            print("")
            # only added in becasue of the dataset I was using ;)
        else:
            print("Unknown file type")
