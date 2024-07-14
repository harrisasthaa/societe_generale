import zipfile
import os
import re
from oletools.olevba import VBA_Parser

def extract_vba_code(file_path):
    vba_code = ""
    vba_temp_path = "vbaProject.bin"

    # Extract the vbaProject.bin file from the .xlsm file
    with zipfile.ZipFile(file_path, 'r') as z:
        for file_info in z.infolist():
            if file_info.filename.endswith('vbaProject.bin'):
                with z.open(file_info) as f:
                    with open(vba_temp_path, 'wb') as temp_file:
                        temp_file.write(f.read())

    # Read the vbaProject.bin file using oletools
    vba_parser = VBA_Parser(vba_temp_path)
    if vba_parser.detect_vba_macros():
        for (filename, stream_path, vba_filename, vba_code_chunk) in vba_parser.extract_macros():
            vba_code += vba_code_chunk

    # Clean up the temporary vbaProject.bin file
    try:
        os.remove(vba_temp_path)
    except PermissionError as e:
        print(f"Failed to remove temporary file {vba_temp_path}: {e}")
    
    return vba_code

def analyze_vba_code(vba_code):
    procedures = re.findall(r'Sub\s+(\w+)', vba_code, re.IGNORECASE)
    functions = re.findall(r'Function\s+(\w+)', vba_code, re.IGNORECASE)
    variables = re.findall(r'Dim\s+(\w+)', vba_code, re.IGNORECASE)
    return {
        "procedures": procedures,
        "functions": functions,
        "variables": variables
    }
