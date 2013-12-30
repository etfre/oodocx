import os
from lxml import etree
from oodocx import helper_functions

NSPREFIXES = {
    'mo': 'http://schemas.microsoft.com/office/mac/office/2008/main',
    'o':  'urn:schemas-microsoft-com:office:office',
    've': 'http://schemas.openxmlformats.org/markup-compatibility/2006',
    # Text Content
    'w':   'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    'w10': 'urn:schemas-microsoft-com:office:word',
    'wne': 'http://schemas.microsoft.com/office/word/2006/wordml',
    # Drawing
    'a':   'http://schemas.openxmlformats.org/drawingml/2006/main',
    'm':   'http://schemas.openxmlformats.org/officeDocument/2006/math',
    'mv':  'urn:schemas-microsoft-com:mac:vml',
    'pic': 'http://schemas.openxmlformats.org/drawingml/2006/picture',
    'v':   'urn:schemas-microsoft-com:vml',
    'wp':  'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
    # Properties (core and extended)
    'cp':  'http://schemas.openxmlformats.org/package/2006/metadata/core-properties',
    'dc':  'http://purl.org/dc/elements/1.1/',
    'ep':  'http://schemas.openxmlformats.org/officeDocument/2006/extended-properties',
    'xsi': 'http://www.w3.org/2001/XMLSchema-instance',
    # Content Types
    'ct':  'http://schemas.openxmlformats.org/package/2006/content-types',
    # Package Relationships
    'r':   'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'pr':  'http://schemas.openxmlformats.org/package/2006/relationships',
    # Dublin Core document properties
    'dcmitype': 'http://purl.org/dc/dcmitype/',
    'dcterms':  'http://purl.org/dc/terms/'}

def makeelement(tagname, tagtext=None, nsprefix='w', attributes=None, attrnsprefix=None):
    '''Create an element & return it'''
    # Deal with list of nsprefix by making namespacemap
    namespacemap = None
    if isinstance(nsprefix, list):
        namespacemap = {}
        for prefix in nsprefix:
            namespacemap[prefix] = NSPREFIXES[prefix]
        # FIXME: rest of code below expects a single prefix
        nsprefix = nsprefix[0]
    if nsprefix:
        namespace = '{' + NSPREFIXES[nsprefix] + '}'
    else:
        # For when namespace = None
        namespace = ''
    newelement = etree.Element(namespace+tagname, nsmap=namespacemap)
    # Add attributes with namespaces
    if attributes:
        # If they haven't bothered setting attribute namespace, use an empty string
        # (equivalent of no namespace)
        if not attrnsprefix:
           # Quick hack: it seems every element that has a 'w' nsprefix for its tag uses the same prefix for its attributes
            if nsprefix == 'w':
                attributenamespace = namespace
            else:
                attributenamespace = ''
        else:
            attributenamespace = '{'+NSPREFIXES[attrnsprefix]+'}'
        
        for tagattribute in attributes:
            newelement.set(attributenamespace+tagattribute, attributes[tagattribute])
    if tagtext:
        newelement.text = tagtext
    newelement.prefix
    return newelement

def write_rels():
    relationships = etree.fromstring(
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>')
    relationship3 = etree.fromstring('<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties"'
    ' Target="docProps/app.xml"/>')
    relationship2 = etree.fromstring('<Relationship Id="rId2" '
    'Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties"'
    ' Target="docProps/core.xml"/>')
    relationship1 = etree.fromstring('<Relationship Id="rId1" '
    'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"'
    ' Target="word/document.xml"/>')
    relationships.append(relationship3)
    relationships.append(relationship2)
    relationships.append(relationship1)
    return relationships

def write_content_types():
    content_types = etree.fromstring(
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"> '
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/> '
    '<Default Extension="xml" ContentType="application/xml"/> '
    '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/> '
    '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/> '
    '<Override PartName="/word/stylesWithEffects.xml" ContentType="application/vnd.ms-word.stylesWithEffects+xml"/> '
    '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/> '
    '<Override PartName="/word/webSettings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.webSettings+xml"/> '
    '<Override PartName="/word/fontTable.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.fontTable+xml"/> '
    '<Override PartName="/word/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/> '
    '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/> '
    '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/> '
    '<Default Extension="jpeg" ContentType="image/jpeg"/> '
    '<Default Extension="gif" ContentType="image/gif"/> '
    '<Default Extension="jpg" ContentType="image/jpeg"/> '
    '<Default Extension="png" ContentType="image/png"/> '
    '</Types>')
    # document.xmlfiles[document.comments] = os.path.join('word', 'comments.xml')
    return content_types
    
def setup_comments(document):
    if document.comments is None:
        document.comments = etree.fromstring(
        '<w:comments xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" '
        'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" '
        'xmlns:o="urn:schemas-microsoft-com:office:office" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
        'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" '
        'xmlns:v="urn:schemas-microsoft-com:vml" '
        'xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" '
        'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" '
        'xmlns:w10="urn:schemas-microsoft-com:office:word" '
        'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
        'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" '
        'xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" '
        'xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk" '
        'xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml" '
        'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" '
        'mc:Ignorable="w14 wp14"></w:comments>')
        next_id = '0'
        document.xmlfiles[document.comments] = os.path.join('word', 'comments.xml')
    else:
        next_id = str(len([element for element in document.comments if
        element.tag == '{' + NSPREFIXES['w'] + '}Comment']))
    add_content_override(document,  '/word/comments.xml',
                            'application/vnd.openxmlformats-officedocument'
                            '.wordprocessingml.comments+xml')
    helper_functions.add_relationship(document, 'comments.xml', 'http://schemas.'
    'openxmlformats.org/officeDocument/2006/relationships/comments')
    return next_id
    
def add_content_override(document, part_name, content_type): 
    '''checks Types element to see if comments element is included,
    adds it if not'''
    flat_contenttypes = sum(
    [child.items() for child in document.contenttypes.getchildren()], [])
    if part_name not in [child[1] for child in flat_contenttypes]:
        document.contenttypes.append(makeelement('Override', nsprefix=None,
        attributes={'PartName': part_name, 'ContentType': content_type}))