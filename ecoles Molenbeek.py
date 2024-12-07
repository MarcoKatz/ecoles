import os
import re
import pandas as pd
from docx import Document

def extract_school_info(doc_path):
    document = Document(doc_path)
    schools = []
    school = {}
    
    # Updated regex pattern to handle both accented and non-accented characters
    school_name_keywords = r'(é|e|É|E)cole|school|Ath(é|e)n(é|e)e|institut|lycee|paviljoen|GBS|academie|ACADEMIE|campus|scolaire|college|instituut|ISFSC|humaniora|centrum|atheneum'
    
    for para in document.paragraphs:
        text = para.text.strip()
        
        if re.search(school_name_keywords, text, re.IGNORECASE):
            if school:
                schools.append(school)
                school = {}
            school['Name'] = text
        
        elif text.lower().startswith('direction :'):
            school['Director'] = text.split(':', 1)[1].strip()
        
        elif re.search(r'\b(rue|place|avenue|chemin|square|chaussée)\b', text, re.IGNORECASE) and re.search(r'\d+', text):
            school['Address'] = text
        
        elif re.match(r'^\d{4}', text) and 'Bruxelles' in text:
            address_parts = text.split(' ')
            school['ZipCode'] = address_parts[0]
            school['Town'] = ' '.join(address_parts[1:])
        
        elif re.match(r'^(tél|Tél)', text):
            if 'Name' in school:
                school['Telephone'] = text.split(':', 1)[1].strip()
        
        elif re.match(r'^E-mail', text, re.IGNORECASE):
            if 'Name' in school:
                school['Email'] = text.split(':', 1)[1].strip()
        
        elif '@' in text:
            if 'Name' in school:
                school['Email'] = text.strip()
        
        elif text.startswith('http'):
            if 'Name' in school:
                school['Website'] = text
        
        elif re.match(r'^Fax', text, re.IGNORECASE):
            continue  # Skip the fax line
        
        elif text == "EN SAVOIR PLUS" or text == "Projet pédagogique et règlement":
            continue  # Ignore these lines
        
        else:
            if 'Address' in school:
                school['Address'] += ' ' + text
    
    if school:
        schools.append(school)
    
    return schools

def save_to_excel(schools, output_path):
    df = pd.DataFrame(schools)
    df['FullAddress'] = df['Address']
    df[['Street', 'Number']] = df['FullAddress'].str.extract(r'(.+?), (\d+)$')
    df.drop(columns=['Address'], inplace=True)
    
    # Ensure all expected columns are present
    for col in ['Name', 'Director', 'FullAddress', 'Street', 'Number', 'ZipCode', 'Town', 'Telephone', 'Email', 'Website']:
        if col not in df.columns:
            df[col] = None
    
    df = df[['Name', 'Director', 'FullAddress', 'Street', 'Number', 'ZipCode', 'Town', 'Telephone', 'Email', 'Website']]
    df.to_excel(output_path, index=False)

if __name__ == "__main__":
    base_dir = os.path.dirname(os.path.abspath(__file__))
    doc_path = os.path.join(base_dir, 'Ecoles Jette.docx')
    output_path = os.path.join(base_dir, 'ecoles Jette.xlsx')
    
    schools = extract_school_info(doc_path)
    save_to_excel(schools, output_path)