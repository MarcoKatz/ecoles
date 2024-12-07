import os
import re
import pandas as pd
from docx import Document

def extract_school_info(doc_path):
    document = Document(doc_path)
    schools = []
    school = {}
    current_section = None
    address_line_count = 0
    
    for para in document.paragraphs:
        text = para.text.strip()
        
        if not school or 'Name' not in school:
            school['Name'] = text
        
        elif text.lower().startswith("nom du chef d'établissement"):
            current_section = 'Director'
        
        elif text.lower().startswith('adresse'):
            current_section = 'Address'
            address_line_count = 0
        
        elif text.lower().startswith('type'):
            current_section = 'Type'
            school['Type'] = ''
        
        elif re.match(r'^(tél|Tél)', text, re.IGNORECASE):
            current_section = 'Telephone'
        
        elif re.match(r'^email', text, re.IGNORECASE):
            current_section = 'Email'
        
        elif re.match(r'^fax', text, re.IGNORECASE):
            current_section = None  # Skip the fax line
        
        elif re.match(r'^site internet', text, re.IGNORECASE):
            current_section = 'Website'
        
        elif re.match(r'^langue école', text, re.IGNORECASE):
            current_section = 'Langue'
        
        elif re.match(r'^transports en commun', text, re.IGNORECASE):
            current_section = 'Transports'
        
        elif re.match(r'^nombre d\'étudiant', text, re.IGNORECASE):
            current_section = None  # Skip the "Nombre d'étudiant" section
        
        elif current_section == 'Director':
            school['Director'] = text
            current_section = None
        
        elif current_section == 'Address':
            if address_line_count == 0:
                school['Address'] = text
                address_line_count += 1
            else:
                address_parts = text.split(' ')
                school['ZipCode'] = address_parts[0]
                school['Town'] = ' '.join(address_parts[1:])
                current_section = None
        
        elif current_section == 'Type':
            if school['Type']:
                school['Type'] += ' / ' + text
            else:
                school['Type'] = text
        
        elif current_section == 'Telephone':
            school['Telephone'] = text
            current_section = None
        
        elif current_section == 'Email':
            school['Email'] = text
            current_section = None
        
        elif current_section == 'Website':
            school['Website'] = text
            current_section = None
        
        elif current_section == 'Langue':
            school['Langue'] = text
            current_section = None
        
        elif text == "EN SAVOIR PLUS":
            if school:
                schools.append(school)
                school = {}
                current_section = None
    
    if school:
        schools.append(school)
    
    return schools

def save_to_excel(schools, output_path):
    df = pd.DataFrame(schools)
    df['FullAddress'] = df['Address']
    df[['Street', 'Number']] = df['FullAddress'].str.extract(r'(.+?), (\d+)$')
    df.drop(columns=['Address'], inplace=True)
    
    # Ensure all expected columns are present
    for col in ['Name', 'Director', 'FullAddress', 'Street', 'Number', 'ZipCode', 'Town', 'Type', 'Telephone', 'Email', 'Website', 'Langue']:
        if col not in df.columns:
            df[col] = None
    
    df = df[['Name', 'Director', 'FullAddress', 'Street', 'Number', 'ZipCode', 'Town', 'Type', 'Telephone', 'Email', 'Website', 'Langue']]
    df.to_excel(output_path, index=False)

if __name__ == "__main__":
    base_dir = os.path.dirname(os.path.abspath(__file__))
    doc_path = os.path.join(base_dir, 'Enseignement 1000 Bruxelles load.docx')
    output_path = os.path.join(base_dir, 'Ecoles Bruxelles 1000.xlsx')
    
    schools = extract_school_info(doc_path)
    save_to_excel(schools, output_path)