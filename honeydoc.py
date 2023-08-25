#!/usr/bin/env python
""" Insert tracking information into docx files """

import click
import lxml.etree as ET
import magic
from pathlib import Path
from tempfile import TemporaryDirectory
import zipfile

def ensure_png_content_type(content_type_path: str):
    ''' Ensure PNG is supported content type '''
    with open(content_type_path, 'r') as rels_file:
        tree = ET.parse(rels_file)

    root = tree.getroot()
    
    extensions = []
    for child in root.findall('{http://schemas.openxmlformats.org/package/2006/content-types}Default'):
        extensions.append(child.get('Extension'))
    
    if 'png' not in extensions:
        element = ET.XML('''<Default Extension="png" ContentType="image/png"/>''')
        root.insert(0, element)
        tree.write(content_type_path, encoding="UTF-8", xml_declaration=True)
    return

def generate_doc_tracker(relId: str):
    ''' Generate the 1x1 pixel tracking image '''
    # Hack to set namespaces
    name_spaces = {
        'w': "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
        'wp': "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
        'a': "http://schemas.openxmlformats.org/drawingml/2006/main",
        'pic': "http://schemas.openxmlformats.org/drawingml/2006/picture",
        'mc': "http://schemas.openxmlformats.org/markup-compatibility/2006",
        'r': "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

    }
    for prefix,uri in name_spaces.items():
        ET.register_namespace(prefix, uri)
    
    docpr_name = 'get rekt'
    pic_name = 'get_hacked'
    # Not sure if this can be simplified
    drawing_string = f'''
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" mc:Ignorable="w14 w15 w16se w16cid w16 w16cex w16sdtdh wp14">
    <w:r>
        <w:drawing mc:Ignorable="w14 wp14">
            <wp:inline distT="0" distB="0" distL="0" distR="0">
                <wp:extent cx="1" cy="1"/>
                <wp:effectExtent l="0" t="0" r="0" b="0"/>
                <wp:docPr id="4" name="{docpr_name}"/>
                <wp:cNvGraphicFramePr>
                    <a:graphicFrameLocks xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" noChangeAspect="1"/>
                </wp:cNvGraphicFramePr>
                <a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                    <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
                        <pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
                            <pic:nvPicPr>
                                <pic:cNvPr id="0" name="{pic_name}"/>
                                <pic:cNvPicPr/>
                            </pic:nvPicPr>
                            <pic:blipFill>
                                <a:blip r:link="{relId}"/>
                                <a:stretch>
                                    <a:fillRect/>
                                </a:stretch>
                            </pic:blipFill>
                            <pic:spPr>
                                <a:xfrm>
                                    <a:off x="0" y="0"/>
                                    <a:ext cx="1" cy="1"/>
                                </a:xfrm>
                                <a:prstGeom prst="rect">
                                    <a:avLst/>
                                </a:prstGeom>
                            </pic:spPr>
                        </pic:pic>
                    </a:graphicData>
                </a:graphic>
            </wp:inline>
        </w:drawing>
    </w:r>
</w:document>
'''
    # xml needs to be flat for docx format
    xml_string = drawing_string.replace('\n', '').replace('    ', '')
    root = ET.fromstring(xml_string)

    drawing = root.find('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}r')
    
    return drawing
    
def insert_doc_tracker(doc_path: str, doc_tracker: str):
    # Not sure if this is still needed
    namespaces = {}
    for event, elem in ET.iterparse(doc_path, events=('start-ns',)):
        prefix, uri = elem
        if prefix is None:
            prefix = ''
        namespaces[prefix] = uri
        ET.register_namespace(prefix, uri)
    
    tree = ET.parse(doc_path)
    root = tree.getroot()
    # Find first paragraph and insert the tracker into it
    paragraph = root.find('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}p')

    paragraph.append(doc_tracker)
    with open(doc_path, 'w') as doc_file:
        tree.write(doc_path, encoding="UTF-8", xml_declaration=True) # maybe hardcode xml_declaration
    return

def generate_rels_tracker(relID: str, target: str):
    ''' Given a set of parameters, generates a tracker rels object '''
    # Create the root Relashionship element
    element = ET.Element("Relationship")
    
    # Set the attributes
    element.set("Id", relID)
    element.set("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image")
    element.set("Target", target)
    element.set("TargetMode", "External")
    return element

def insert_rels_tracker(rels_path: str, rels_tracker: str):
    ''' Insert tracker image into Docx rels file '''
    # Parse the rel file
    with open(rels_path, 'r') as rels_file:
        tree = ET.parse(rels_file)
    root = tree.getroot()
    root.append(rels_tracker)

    with open(rels_path, 'w') as xml_file:
        #xml_file.write('<?xml version="1.0" encoding="UTF-8" standalone = "yes"?>\n')
        tree.write(rels_path, encoding="UTF-8", xml_declaration = True)
    return

def grab_max_rId(rels_path: str):
    ''' Given path to extracted Docx document relashionship file, find max rId '''
    with open(rels_path, 'r') as rels_file:
        tree = ET.parse(rels_file)

    root = tree.getroot()
    
    rIds = []
    for child in root.findall('{http://schemas.openxmlformats.org/package/2006/relationships}Relationship'):
        rId = int(child.get('Id').replace('rId', ''))
        rIds.append(rId)
    maxId = max(rIds)
    return maxId

def insert_token(docx_folder: str, target: str):
    ''' Takes a filepath to extracted docx file folder and inserts a tracker '''
    # 1. Ensure image content type is supported
    content_types = '[Content_Types].xml'
    content_types_path = docx_folder / content_types
    ensure_png_content_type(content_types_path)
    
    # 2. Add rels to external target
    rels = 'word/_rels/document.xml.rels'
    rels_path = docx_folder / rels
    
    maxId = grab_max_rId(rels_path)
    #click.echo(f'Max rId: {maxId}')
    
    relId = 'rId'+str(maxId+1)
    #relId = 'rId1337'
    #target = "http://127.0.0.1:8000/shot.png"
    
    rels_tracker = generate_rels_tracker(relId, target)
    insert_rels_tracker(rels_path, rels_tracker)
    
    # 3. Insert image into document
    document = 'word/document.xml'
    document_path = docx_folder / document
    doc_tracker = generate_doc_tracker(relId)
    insert_doc_tracker(document_path, doc_tracker)
    
    return

def extract_docx(docx_path: str, temp_path: str):
    with zipfile.ZipFile(docx_path, 'r') as doczip:
        doczip.extractall(temp_path)
    return

def package_docx(input_path: str, output_file:str):
    with zipfile.ZipFile(output_file, 'w', zipfile.ZIP_DEFLATED) as zipf:
        for file_path in Path(input_path).rglob('*'):
            # Hack to Force Timestamps to Epoch for Zip format
            info = zipfile.ZipInfo.from_file(file_path)
            info.date_time = (1980, 1, 1, 0, 0, 0)
            zipf.write(file_path, file_path.relative_to(input_path))        

@click.command()
@click.option('--file', required=True, help='Path to docx to trap')
@click.option('--url', required=True, help='URL of remote image on C2')
@click.option('--out', default='honeydoc.docx', show_default=True, help='Name of output file')
def docx_token(file: str, url:str, out: str):
    """ Insert tracking pixel into docx files """
    docx_path = Path(file)
    if out:
        out_path = Path(out).with_suffix('.docx')
    file_type = magic.from_file(str(docx_path))
    if not docx_path.exists():
        click.secho(f"Invalid File Path: {docx_path}", fg='red')
        return
    elif docx_path.suffix != '.docx':
        click.secho(f"Invalid File Extension: {docx_path.suffix}", fg='red')
        return
    elif file_type != 'Microsoft Word 2007+':
        click.secho(f"Invalid File Type: {file_type}", fg='red')
        return

    with TemporaryDirectory() as temp_dir:
        temp_path = Path(temp_dir)
        # 1. Extract the Docx to a temporary folder
        extract_docx(docx_path, temp_path)

        # 2. Insert the tracker by modifying XML files
        insert_token(temp_path, url)
        
        # 3. Repackage modified XML files into trapped Docx
        package_docx(temp_path, out_path) # change to different path?
        click.secho(f"Succesfully created HoneyToken Docx file at:\n{out_path.resolve()}", fg='green')
    return

if __name__ == "__main__":
    docx_token()