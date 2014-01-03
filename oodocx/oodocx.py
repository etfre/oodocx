"""
Open and modify Microsoft Word 2007 and 2010 docx files (called 'OpenXML' and
'Office OpenXML' by Microsoft)

https://github.com/evfredericksen/oodocx
See LICENSE for licensing information.
"""

import logging
import zipfile
import shutil
import re
import time
import datetime
import os
import io
import collections
import stat
import tempfile
from lxml import etree
from oodocx import helper_functions
from oodocx import write_files

log = logging.getLogger(__name__)
TEMPLATE_DIR = os.path.join(os.path.dirname(__file__), 'template')
BASE_DIR = tempfile.mkdtemp()
# All Word prefixes / namespace matches used in document.xml & core.xml.
# LXML doesn't actually use prefixes (just the real namespace) , but these
# make it easier to copy Word output more easily.
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
    'dcterms':  'http://purl.org/dc/terms/',
    # Special xml namespace
    'xml': 'http://www.w3.org/XML/1998/namespace'}

COLOR_MAP = {
    'black': '000000',
    'blue': '0000FF',
    'gray': '808080',
    'green': '00FF00',
    'grey': '808080',
    'orange': 'FFA500',
    'pink': 'FFCBDB',
    'purple': 'FF00FF',
    'red': 'FF0000',
    'silver': 'C0C0C0',
    'white': 'FFFFFF',
    'yellow': 'FFFF00'}


class Docx():
    def __init__(self, file=''):
        folder_number = 1
        self.write_dir = os.path.join(BASE_DIR, str(folder_number))
        while os.path.isdir(self.write_dir):
            folder_number += 1
            self.write_dir = os.path.join(BASE_DIR,  str(folder_number))
        # dictionary to connect element objects to their path in the docx file
        self.xmlfiles = {}
        try:
            shutil.rmtree(self.write_dir, onerror=helper_functions.remove_readonly)
        except:
            pass
        self.media_dir = os.path.join(self.write_dir, 'word', 'media')
        # Declare empty attributes, which may or may not be assigned to xml
        # elements later
        self.comments = None
        # self.xmlfiles[self.comments] = os.path.join('word/comments.xml')
        if file:
            os.mkdir(self.write_dir)
            zipdoc = zipfile.ZipFile(file)
            for filepath in zipdoc.namelist():
                zipdoc.extractall(self.write_dir)
        else:
            shutil.copytree(TEMPLATE_DIR, self.write_dir)
            
            self.rels = write_files.write_rels()
            self.xmlfiles[self.rels] = os.path.join('_rels', '.rels')
            
            self.contenttypes = write_files.write_content_types()
            self.xmlfiles[self.contenttypes] = '[Content_Types].xml'
            
        for root, dirs, filenames in os.walk(self.write_dir):
            for file in filenames:
                if file[-4:] == '.xml' or file[-5:] == '.rels':
                    absdir = os.path.abspath(os.path.join(root, file))
                    docstr = io.open(absdir, 'r', encoding='utf8')
                    relpath = os.path.relpath(absdir, self.write_dir)
                    xmlfile = (etree.fromstring(docstr.read().encode('utf_8')))
                    if file == '[Content_Types].xml':
                        self.contenttypes = xmlfile
                        self.xmlfiles[self.contenttypes] = relpath
                        # update self.contenttypes, as needed
                        filetypes = {'gif':  'image/gif',
                        'jpeg': 'image/jpeg',
                        'jpg':  'image/jpeg',
                        'png':  'image/png',
                        'rels': 'application/vnd.openxmlformats-package'
                                '.relationships+xml',
                        'xml':  'application/xml'}
                        default_elements = [child for child
                        in self.contenttypes.iterchildren()
                        if 'Default' in child.tag] 
                        for key, value in filetypes.items():
                            missing_filetype = True
                            for child in default_elements:
                                if key == child.items()[0][1]:
                                    missing_filetype = False
                            if missing_filetype:
                                default_element = makeelement('Default',
                                nsprefix=None,
                                attributes={'Extension': key,
                                'ContentType': value})
                                self.contenttypes.append(default_element)
                    elif file == 'app.xml':
                        self.app = xmlfile
                        self.xmlfiles[self.app] = relpath
                    elif file == 'comments.xml':
                        self.comments = xmlfile
                        self.xmlfiles[self.comments] = relpath
                    elif file == 'core.xml': 
                        self.core = xmlfile
                        self.xmlfiles[self.core] = relpath
                    elif file == 'document.xml': 
                        self.document = xmlfile
                        self.xmlfiles[self.document] = relpath
                    elif file == 'document.xml.rels': 
                        self.relationships = xmlfile
                        self.xmlfiles[self.relationships] = relpath	
                    elif file == 'fontTable.xml': 
                        self.fontTable = xmlfile
                        self.xmlfiles[self.fontTable] = relpath
                    elif file == 'settings.xml': 
                        self.styles = xmlfile
                        self.xmlfiles[self.styles] = relpath						
                    elif file == 'styles.xml': 
                        self.styles = xmlfile
                        self.xmlfiles[self.styles] = relpath
                    elif file == 'stylesWithEffects.xml': 
                        self.stylesWithEffects = xmlfile
                        self.xmlfiles[self.stylesWithEffects] = relpath		
                    elif file == 'webSettings.xml': 
                        self.webSettings = xmlfile
                        self.xmlfiles[self.webSettings] = relpath	
        self.body = self.document.xpath('/w:document/w:body',
                                        namespaces=NSPREFIXES)[0]
        
    def get_body(self):
        print('Warning: This method is deprecated and will be removed at some '
              'point in the future. Use the self.body attribute instead.')
        return self.document.xpath('/w:document/w:body',
                                   namespaces=NSPREFIXES)[0]
        
    def search(self, search, result_type='text', ignore_runs=True):
        '''Search each paragraph for a regex, returns first matching
        element object or None if nothing found. Will return the
        first element if match spans multiple text elements.'''
        searchre = re.compile(search)
        result = None
        if ignore_runs:
            for paragraph in self.document.iter('{' + NSPREFIXES['w'] + '}p'):
                text_positions = []
                start = 0
                paragraph_string = get_text(paragraph)
                for text_element in paragraph.iter('{' + NSPREFIXES['w'] +
                                                   '}t'):
                    if text_element.text:
                        text_positions.append({'start': start,
                        'end': start + len(text_element.text) - 1,
                        'element': text_element})
                        start += len(text_element.text)
                match = searchre.search(paragraph_string)
                if match:
                    for position in text_positions:
                        if match.start() in range(position['start'],
                                                  position['end'] + 1):
                            result = position['element']
                            break
                    break
        else:
            for element in self.document.iter('{' + NSPREFIXES['w'] + '}t'):
                if element.text and searchre.search(element.text):
                    result = element
                    break
        if result is not None:
            if result_type.lower() == 'paragraph':
                if (result.iterancestors('{' + NSPREFIXES['w'] + '}p')
                    is not None):
                    while not result.tag == '{' + NSPREFIXES['w'] + '}p':
                        result = result.getparent()
                else:
                    raise
            elif result_type.lower() == 'run':
                if (result.iterancestors('{' + NSPREFIXES['w'] + '}r')
                    is not None):
                    while not result.tag == '{' + NSPREFIXES['w'] + '}r':
                        result = result.getparent()
                else:
                    raise
        return result
        
    def replace(self, search, replace, ignore_runs=True):
        '''Replace all occurrences of string with a different string.
        If ignore_runs is true, the function will ignore separate run
        and text elements and instead search each paragraph text
        content as a single string. Note that this will also ignore
        formatting elements within a paragraph such as tabs, which
        may cause unexpected results. Set ignore_runs to false if you
        want a more conservative search.'''
        searchre = re.compile(search)
        if ignore_runs:
            for paragraph_element in self.document.iter('{' + NSPREFIXES['w'] +
                                                        '}p'):
                rundict = collections.OrderedDict()
                start = 0
                for run_element in paragraph_element.iter('{' + NSPREFIXES['w']
                                                          + '}r'):
                    run_string = get_text(run_element)
                    rundict[run_element] = [start, start + len(run_string),
                                            run_string]
                    start += len(run_string)
                match_slices = [match.span() for match in
                re.finditer(searchre, get_text(paragraph_element))]
                preliminary_runs = collections.OrderedDict()
                runs_to_exclude = set()
                for run, text_info in rundict.items():
                    for match in match_slices:
                        if ((match[0] < text_info[1] and
                        match[1] > text_info[0]) or
                        (match[0] >= text_info[0] and
                        match[1] <= text_info[1])):
                            preliminary_runs[run] = text_info
                            if (match[0] < text_info[0] and
                            match[1] > text_info[1]):
                                runs_to_exclude.add(run)
                runs_to_modify = collections.OrderedDict()
                for run, text_info in list(preliminary_runs.items()):
                    if run not in runs_to_exclude:
                        runs_to_modify[run] = text_info
                    # merge text from runs contained entirely inside
                    # matches with preceding run
                    else:
                        previous_run = list(runs_to_modify.items())[-1][0]
                        previous_text_info = list(
                        runs_to_modify.items())[-1][1]
                        previous_text = previous_run.find(
                        '{' + NSPREFIXES['w'] + '}t')
                        previous_text.text += text_info[2]
                        previous_text_info[1] += len(text_info[2])
                        previous_text_info[2] += text_info[2]
                        paragraph_element.remove(run)
                overflow = 0
                for index, (run, text_info) in enumerate(
                runs_to_modify.items()):
                    runshift = -overflow
                    overflow = 0
                    text_element = run.find('{' + NSPREFIXES['w'] + '}t')
                    newstring = text_element.text
                    for match in match_slices:
                        if match[0] in range(text_info[0], text_info[1]):
                            newstring = (newstring[:match[0] + runshift -
                            text_info[0]] + replace + 
                            text_element.text[match[1] - text_info[0]:])
                            try:
                                if ' ' in (newstring[0], newstring[-1]):
                                    text_element.set('{' + NSPREFIXES['xml'] +
                                    '}space', 'preserve')
                            except IndexError:
                                pass
                            # Difference between replace and search length.
                            # Executes for each match in match_slices to
                            # account for potential difference in lengths
                            # of matches due to regex search argument
                            runshift += len(replace) - (match[1] - match[0])
                            if match[1] > text_info[1]:
                                overflow = match[1] - text_info[1]
                                if index < len(runs_to_modify) - 1:
                                    next_run = list(
                                    runs_to_modify.keys())[index + 1]
                                    next_text = next_run.find(
                                    '{' + NSPREFIXES['w'] + '}t')
                                    next_text.text = next_text.text[overflow:]					
                    text_element.text = newstring
        else:
            for element in self.document.iter('{' + NSPREFIXES['w'] + '}t'):
                if element.text and searchre.search(element.text):
                    element.text = re.sub(search, replace, element.text)
                        
    def clean(self):
        '''Remove empty text and run elements'''
        for t in ('t', 'r'):
            remove_list = []
            for element in self.document.iter():
                if (element.tag == '{' + NSPREFIXES['w'] + '}t' and
                not element.text and not len(element)):
                    remove_list.append(element)
            for element in remove_list:
                element.getparent().remove(element)
        
    def add_style(self, styleId, type, default=None, name=None):
        if default:
            style = makeelement('style', attributes={'styleId': styleId, 
            'type': type, 'default': default})
        else:
            style = makeelement('style', attributes={'styleId': styleId, 
            'type': type})
        style.append(makeelement('pPr'))
        style.append(makeelement('rPr'))
        self.styles.append(style)
        return style
        
    def set_margins(self, left='', right='', top='', bottom='', header='',
    footer='', gutter=''):
        sectPr = self.get_section_properties()
        attributes_dict = {'left': left, 'right': right, 'top': top, 'bottom':
        bottom, 'header': header, 'footer': footer, 'gutter': gutter}
        pgMar = sectPr.find('{' + NSPREFIXES['w'] + '}pgMar')
        if pgMar is None:
            pgMar = makeelement('pgMar')
            sectPr.append(pgMar)
        for key, value in attributes_dict.items():
            if value:
                pgMar.set('{' + NSPREFIXES['w'] + '}' + str(key), value)
        
    def modify_paragraph_defaults(self, indent='default', spacing='default',
    pstyle='default', justification='default', modify_styles=False):
        elements_to_modify = []
        docdefaults = self.styles.find('{' + NSPREFIXES['w'] + '}docDefaults')
        if docdefaults is None:
            docdefaults = makeelement('docDefaults')
            self.styles.insert(0, docdefaults)
        pprdefault = self.styles.find('{' + NSPREFIXES['w'] + '}pPrDefault')
        if pprdefault is None:
            pprdefault = makeelement('pPrDefault')
            docdefaults.append(pprdefault)
        elements_to_modify.append(pprdefault)
        if modify_styles:
            elements_to_modify.extend([element for element in
            self.styles.iterchildren('{' + NSPREFIXES['w'] + '}style')])
        modify_paragraph(elements_to_modify, indent=indent,
        spacing=spacing, pstyle=pstyle, justification=justification)

    def modify_font_defaults(self, name='default', size='default', 
    underline='default', color='default', highlight='default',
    strikethrough='default', bold='default', subscript='default',
    superscript='default', italics='default', shadow='default',
    smallcaps='default', allcaps='default', hidden='default',
    modify_styles=False):
        elements_to_modify = []
        docdefaults = self.styles.find('{' + NSPREFIXES['w'] + '}docDefaults')
        if docdefaults is None:
            docdefaults = makeelement('docDefaults')
            self.styles.insert(0, docdefaults)
        rprdefault = self.styles.find('{' + NSPREFIXES['w'] + '}rPrDefault')
        if rprdefault is None:
            rprdefault = makeelement('rPrDefault')
            docdefaults.append(rprdefault)
        elements_to_modify.append(rprdefault)
        if modify_styles:
            elements_to_modify.extend([element for element in
            self.styles.iterchildren('{' + NSPREFIXES['w'] + '}style')])
        modify_font(elements_to_modify, name=name, size=size,
        underline=underline, color=color, highlight=highlight,
        strikethrough=strikethrough, subscript=subscript,
        superscript=superscript, bold=bold, italics=italics, shadow=shadow,
        smallcaps=smallcaps, allcaps=allcaps, hidden=hidden)
        
    def get_section_properties(self):
        '''Returns the sectPr element at the end of the body, creates
        the element first if one is not found'''
        sect_list = [child for child in
                     self.body.iterchildren('{' + NSPREFIXES['w'] + '}sectPr')]
        if len(sect_list) == 1:
            return sect_list[0]
        elif not sect_list:
            sect_props = makeelement('sectPr')
            self.body.append(sect_props)
            return sect_props
            
    def get_document_text(self):
        '''Return the raw text of a document, as a list of paragraphs.'''
        paratextlist = []
        # Compile a list of all paragraph (p) elements
        paralist = []
        for element in self.document.iter('{' + NSPREFIXES['w'] + '}p'):
            paralist.append(element)
        # Since a single sentence might be spread over multiple text elements,
        # iterate through each paragraph, appending all text (t) children to
        # that paragraph's text.
        for para in paralist:
            paratext = ''
            for element in para.iter('{' + NSPREFIXES['w'] + '}t'):
                # Find t (text) elements
                if element.text:
                    paratext += paratext
                elif element.tag == '{' + NSPREFIXES['w'] + '}tab':
                    paratext += '\t'
            # Add our completed paragraph text to the list of paragraph text
            if len(paratext):
                paratextlist.append(paratext)
        return paratextlist
    
    def merge(self, docpath, page_break=True):
        '''Appends a .docx to the end of this document. docpath can
        either be a Docx object or a file path. This method will likely
        break if both documents possess the same type of elements that
        require id mapping such as lists or comments. Pictures and other
        <w: drawing> elements, however, should work 100% of the time.'''
        if isinstance(docpath, Docx):
            fromdoc = docpath
        else:
            fromdoc = Docx(docpath)
        # Update relationship Ids
        for relationship in fromdoc.relationships:
            old_rId = relationship.values()[0]
            relationship_type = relationship.values()[1]
            relationship_target = relationship.values()[2]
            new_rId = helper_functions.add_relationship(self,
                                                        relationship_target,
                                                        relationship_type)
            if new_rId is None:
                new_rId = old_rId
            for element in fromdoc.body.iter():
                for item in element.items():
                    attribute = item[0]
                    value = item[1]
                    if old_rId == value:
                        element.set(attribute, new_rId)
        if not os.path.isdir(self.media_dir):
            os.mkdir(self.media_dir)
        tofiles = []
        for (dirpath, dirnames, filenames) in os.walk(self.write_dir):
            relpath = dirpath[len(self.write_dir) + 1:]
            for file in filenames:
                relfile = os.path.join(relpath, file)
                tofiles.append(relfile)
        for (dirpath, dirnames, filenames) in os.walk(fromdoc.write_dir):
            relpath = dirpath[len(fromdoc.write_dir) + 1:]
            head, tail = os.path.split(dirpath)
            if tail == 'media':
                for file in filenames:
                    if os.path.join(self.media_dir, file) not in tofiles:
                        shutil.copyfile(os.path.join(fromdoc.media_dir, file),
                                        os.path.join(self.media_dir, file))
                    else:  # Account for duplicate picture names
                        shutil.copyfile(os.path.join(fromdoc.media_dir,
                                                    'new_' + file),
                                        os.path.join(self.media_dir, file))
            else:
                for file in filenames:
                    if not os.path.isdir(os.path.join(self.write_dir, relpath)):
                        os.mkdir(os.path.join(self.write_dir, relpath))
                    if os.path.join(relpath, file) not in tofiles:  
                        shutil.copyfile(os.path.join(fromdoc.write_dir, relpath, file),
                        os.path.join(self.write_dir, relpath, file))
        # Update Content Types if necessary
        for type in fromdoc.contenttypes.iterchildren():
            type_string = etree.tostring(type)
            if type_string not in [etree.tostring(second_type) for second_type
                                   in self.contenttypes.iterchildren()]:
                self.contenttypes.append(type)
        first_sectpr = self.body.find('{' + NSPREFIXES['w'] + '}sectPr')
        if page_break:
            try:
                last_para = [para for para in
                      self.body.iterchildren('{' + NSPREFIXES['w'] + '}p')][-1]
            except IndexError:
                last_para = paragraph('')
                self.body.append(last_para)
            run = makeelement('r')
            last_para.append(run)
            run.append(makeelement('br', attributes={'type': 'page'}))
            first_para = fromdoc.body.find('{' + NSPREFIXES['w'] + '}p')
            if first_para is None:
                first_para = paragraph('')
                fromdoc.body.append(first_para)
            r = first_para.find('{' + NSPREFIXES['w'] + '}r')
            if r is None:
                r = makeelement('r')
                first_para.append(r)
            r.insert(0, makeelement('lastRenderedPageBreak'))
        self.body.extend(fromdoc.body.iterchildren())
      
    def save(self, output):
        '''Saves the Docx to the output path provided.'''
        docxfile = zipfile.ZipFile(output, mode='w',
        compression=zipfile.ZIP_DEFLATED)
        # Move to the template data path
        prev_dir = os.path.abspath('.')  # save previous working dir
        os.chdir(self.write_dir)
        # Write changes made to xml files in write directory between __init__()
        # and save()
        for xmlfile, relpath in self.xmlfiles.items():
            absolutepath = os.path.split(
                           os.path.join(self.write_dir, relpath))[0]
            if not os.path.isdir(absolutepath):
                os.mkdir(absolutepath)
            newdoc = io.open(relpath, 'w')
            newdoc.write(etree.tostring(
            xmlfile, xml_declaration=True).decode(encoding='UTF-8'))
            newdoc.close()
        files_to_ignore = ['.DS_Store']  # nuisance from some os's
        for dirpath, dirnames, filenames in os.walk('.'):
            for filename in filenames:
                if filename in files_to_ignore:
                    continue
                templatefile = os.path.join(dirpath, filename)
                archivename = templatefile[2:]
                docxfile.write(templatefile, archivename)
        docxfile.close()
        os.chdir(prev_dir)  # restore previous working dir
        try:
            shutil.rmtree(self.write_dir, onerror=helper_functions.remove_readonly)
        except FileNotFoundError:
            pass
    
def merge_text(run):
    '''Combines the text of all text elements in a run into a single
    text element, removes the other text elements.'''
    for index, child in enumerate(
    run.iterchildren('{' + NSPREFIXES['w'] + '}t')):
        if index == 0:
            child.text = get_text(run)
        else:
            run.remove(child)
    
def modify_font(elements, name='default', size='default', underline='default',
color='default', highlight='default', strikethrough='default', bold='default',
subscript='default', superscript='default', italics='default',
shadow='default', smallcaps='default', allcaps='default', hidden='default'):
    """Allows you to modify common font properties for all of the runs
    in an element or list/tuple of elements. Some notes:
    *The size parameter interprets its argument as being of the same
    measurement system as you see in the word editor, rather than the
    half-points that open xml uses behind the scenes
    *Bold, italics, subscript, superscript, and shadow will
    ignore any string arguments and will interpret True and 1 as true and
    False and 0 as false.
    *Underline will accept any (case-insensitive) strings that match the
    values in the "Member name" category of the following url:
    http://msdn.microsoft.com/en-us/library/documentformat.openxml.drawing.textunderlinevalues
    Otherwise, it behaves like the above parameters in that it will
    interpret True and 1 as a generic single underline, and False and 0
    as false"""
    underline_values = ('single', 'double', 'thick', 'dotted', 'dash',
    'dotDash', 'dotDotDash', 'wave', 'wavyHeavy', 'wavyDouble')
    highlight_values = ('yellow', 'green', 'cyan' 'magenta', 'blue', 'red',
    'darkBlue', 'darkCyan', 'darkGreen', 'darkMagenta', 'darkRed', 'darkRed',
    'lightGray', 'black')
    if isinstance(elements, (list, tuple)):
        run_list = []
        for element in elements:
            if element.tag in ('{' + NSPREFIXES['w'] + '}rPrDefault',
            '{' + NSPREFIXES['w'] + '}style'):
                run_list.append(element)
                continue
            for child in element.iter('{' + NSPREFIXES['w'] + '}r'):
                run_list.append(child)
    elif elements.tag == '{' + NSPREFIXES['w'] + '}r':
        run_list = [elements]
    else:
        run_list = [child for child in
        elements.iter('{' + NSPREFIXES['w'] + '}r')]
    for run in run_list:
        rpr = run.find('{' + NSPREFIXES['w'] + '}rPr')
        if rpr is None:
            rpr = makeelement('rPr')
            run.insert(0, rpr)
        if name != 'default':
            rfonts = rpr.find('{' + NSPREFIXES['w'] + '}rFonts')
            if rfonts is not None:
                rpr.remove(rfonts)
            rfonts = makeelement('rFonts', attributes={'ascii': name,
            'hAnsi': name})
            rpr.append(rfonts)
        if size != 'default':
            sz = rpr.find('{' + NSPREFIXES['w'] + '}sz')
            szCs = rpr.find('{' + NSPREFIXES['w'] + '}szCs')
            if sz is not None:
                rpr.remove(sz)
            sz = makeelement('sz', attributes={'val': str(int(size) * 2)})
            rpr.append(sz)
            if szCs is not None:
                rpr.remove(szCs)
            szCs = makeelement('szCs', attributes={'val': str(int(size) * 2)})
            rpr.append(szCs)
        if underline != 'default':
            valid_input = False
            u = rpr.find('{' + NSPREFIXES['w'] + '}u')
            for value in underline_values:
                if str(underline).lower() == value.lower():
                    valid_input = True
                    if u is not None:
                        rpr.remove(u)
                    u = makeelement('u', attributes={'val': value})
                    rpr.append(u)
                    break
            if underline in (1, True) or (not valid_input and
            isinstance(underline, str)):
                if u is not None:
                    rpr.remove(u)
                u = makeelement('u', attributes={'val': 'single'})
                rpr.append(u)
            if underline == 0 or underline == False:
                if u is not None:
                    rpr.remove(u)
        if color != 'default':
            color_element = rpr.find('{' + NSPREFIXES['w'] + '}color')
            if color_element is not None:
                rpr.remove(color_element)
            if isinstance(color, str):
                if color.lower() in COLOR_MAP:
                    color = COLOR_MAP[color.lower()]
                color_element = makeelement('color', attributes={'val': color})
                rpr.append(color_element)	
        if highlight != 'default':
            valid_input = False
            highlight_element = rpr.find('{' + NSPREFIXES['w'] + '}highlight')
            if isinstance(highlight, str) or highlight in (0, False):
                if highlight_element is not None:
                    rpr.remove(highlight_element)
                if isinstance(highlight, str):
                    for value in highlight_values:
                        if str(highlight).lower() == value.lower():
                            valid_input = True
                            highlight_element = makeelement('highlight',
                            attributes={'val': value})
                            break
                    if not valid_input:
                        highlight_element = makeelement('highlight',
                        attributes={'val': highlight})
                    rpr.append(highlight_element)
        if strikethrough != 'default':
            strike = rpr.find('{' + NSPREFIXES['w'] + '}strike')
            dstrike = rpr.find('{' + NSPREFIXES['w'] + '}dstrike')
            if (isinstance(strikethrough, str) or 
            strikethrough in (1, True, 0, False)):
                if strike is not None:
                    rpr.remove(strike)
                if dstrike is not None:
                    rpr.remove(dstrike)
                if isinstance(strikethrough, str):
                    if strikethrough.lower() in ('single', 'strike'):
                        rpr.append(makeelement('strike'))
                    elif strikethrough.lower() in ('double', 'dstrike'):
                        rpr.append(makeelement('dstrike'))
                    else:
                        rpr.append(makeelement(strikethrough))
                elif strikethrough in (1, True):
                    rpr.append(makeelement('strike'))
        if subscript != 'default':
            vertAlign = rpr.find('{' + NSPREFIXES['w'] + '}vertAlign')
            if vertAlign is not None and subscript in (0, False, 1, True):
                rpr.remove(vertAlign)
            if subscript in (1, True):
                vertAlign = makeelement('vertAlign',
                attributes={'val': 'subscript'})
                rpr.append(vertAlign)
        if superscript != 'default':
            vertAlign = rpr.find('{' + NSPREFIXES['w'] + '}vertAlign')
            if vertAlign is not None and superscript in (0, False, 1, True):
                rpr.remove(vertAlign)
            if superscript in (1, True):
                vertAlign = makeelement('vertAlign',
                attributes={'val': 'superscript'})
                rpr.append(vertAlign)
        bool_list = ((bold, 'b'), (italics, 'i'), (shadow, 'shadow'), 
        (allcaps, 'caps'), (smallcaps, 'smallCaps'), (hidden, 'vanish'))
        for key, value in bool_list:
            if key != 'default':
                element = rpr.find('{{{0}}}{1}'.format(NSPREFIXES['w'], value))
                if element is not None and key in (0, False, 1, True):
                    rpr.remove(element)
                if key in (1, True):
                    element = makeelement(value)
                    rpr.append(element)
                    
def modify_paragraph(elements, indent='default', spacing='default',
pstyle='default', justification='default'):
    if isinstance(elements, (list, tuple)):
        para_list = []
        for element in elements:
            if element.tag in ('{' + NSPREFIXES['w'] + '}pPrDefault',
            '{' + NSPREFIXES['w'] + '}style'):
                para_list.append(element)
                continue
            para_list.extend([child for child in element if child.tag
            == '{' + NSPREFIXES['w'] + '}p'])
    elif elements.tag == '{' + NSPREFIXES['w'] + '}p':
        para_list = [elements]
    else:
        para_list = [child for child in elements if child.tag
        == '{' + NSPREFIXES['w'] + '}p']
    for para in para_list:
        pPr = para.find('{' + NSPREFIXES['w'] + '}pPr')
        if pPr is None:
            pPr = makeelement('pPr')
            para.insert(0, pPr)
        if indent != 'default':
            ind = pPr.find('{' + NSPREFIXES['w'] + '}ind')
            if ind is not None:
                pPr.remove(ind)
            if isinstance(indent, dict):
                ind = makeelement('ind', attributes=indent)
                pPr.append(ind)
        if spacing != 'default':
            spacing_element = pPr.find('{' + NSPREFIXES['w'] + '}spacing')
            if spacing_element is not None:
                pPr.remove(spacing_element)
            if isinstance(spacing, dict):
                if 'lineRule' not in spacing.keys():
                    spacing['lineRule'] = 'auto'
                spacing_element = makeelement('spacing', attributes=spacing)
                pPr.append(spacing_element)
        if pstyle != 'default':
            pstyle_element = pPr.find('{' + NSPREFIXES['w'] + '}pStyle')
            if pstyle_element is not None:
                pPr.remove(pstyle_element)
            pstyle_element = makeelement('pStyle', attributes={'val': pstyle})
            pPr.append(pstyle_element)
        if justification != 'default':
            jc = para.find('{' + NSPREFIXES['w'] + '}jc')
            if jc is not None:
                jc.set('{' + NSPREFIXES['w'] + '}val', justification.lower())
            else:
                jc = makeelement('jc', attributes={'val': justification.lower()})
                pPr.append(jc)
            
def makeelement(tagname, tagtext=None, nsprefix='w', attributes=None,
                attrnsprefix=None):
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
    newelement = etree.Element(namespace + tagname, nsmap=namespacemap)
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
            attributenamespace = '{' + NSPREFIXES[attrnsprefix] + '}'
        for tagattribute in attributes:
            newelement.set(attributenamespace + tagattribute,
                           attributes[tagattribute])
    if tagtext is not None and len(tagtext):
        newelement.text = tagtext
    newelement.prefix
    return newelement
    
def paragraph(paratext, breakbefore=False, rprops=None, pprops=None):
    '''Make a new paragraph element, containing a run, and some text.
    Return the paragraph element. rprops modifies properties of all
    runs within paragraph, pprops modifies paragraph's overall
    properties (spacing, indentation, etc.)'''
    paragraph = makeelement('p')
    if isinstance(paratext, list):
        text = []
        for pt in paratext:
            if isinstance(pt, (list, tuple)):
                text.append([makeelement('t', tagtext=pt[0]), pt[1]])
            else:
                text.append([makeelement('t', tagtext=pt), ''])
    else:
        text = [[makeelement('t', tagtext=paratext), ''], ]
    pPr = makeelement('pPr')
    # Add paragraph properties
    if pprops:
        if isinstance(pprops, dict):
            for tag, atts in pprops.items():
                pPr.append(makeelement(tag, attributes=atts))
        elif isinstance(pprops, str):
            pPr.append(makeelement(pprops))
        else:
            raise TypeError("pprops argument must be of 'dict' or 'str' type")
    # Add the text to the run, and the run to the paragraph
    paragraph.append(pPr)
    for t in text:
        run = makeelement('r')
        rPr = makeelement('rPr')
        # Add run properties
        if rprops:
            if isinstance(rprops, dict):
                for tag, atts in rprops.items():
                    rPr.append(makeelement(tag, attributes=atts))
            elif isinstance(rprops, str):
                rPr.append(makeelement(rprops))
            else:
                raise TypeError("rprops argument must be of 'dict' or 'str' type")
        run.append(rPr)
        # Apply styles
        if t[1].find('b') > -1:
            b = makeelement('b')
            rPr.append(b)
        if t[1].find('u') > -1:
            u = makeelement('u', attributes={'val': 'single'})
            rPr.append(u)
        if t[1].find('i') > -1:
            i = makeelement('i')
            rPr.append(i)
        run.append(rPr)
        # Insert lastRenderedPageBreak for assistive technologies like
        # document narrators to know when a page break occurred.
        if breakbefore:
            lastRenderedPageBreak = makeelement('lastRenderedPageBreak')
            run.append(lastRenderedPageBreak)
        run.append(t[0])
        paragraph.append(run)
    # Return the combined paragraph
    return paragraph
    # Add the text to the run, and the run to the paragraph
    paragraph.append(pPr)
    for t in text:
        run = makeelement('r')
        rPr = makeelement('rPr')
        # Apply styles
        if t[1].find('b') > -1:
            b = makeelement('b')
            rPr.append(b)
        if t[1].find('u') > -1:
            u = makeelement('u', attributes={'val': 'single'})
            rPr.append(u)
        if t[1].find('i') > -1:
            i = makeelement('i')
            rPr.append(i)
        if run_properties:
            if isinstance(run_properties[0], str):
                a = makeelement(run_properties[0], 
                attributes={run_properties[1]: run_properties[2]})
                rPr.append(a)
            else:
                for element in run_properties:
                    a = makeelement(element[0],
                    attributes={element[1]: element[2]})
                    rPr.append(a)
        run.append(rPr)
        # Insert lastRenderedPageBreak for assistive technologies like
        # document narrators to know when a page break occurred.
        if breakbefore:
            lastRenderedPageBreak = makeelement('lastRenderedPageBreak')
            run.append(lastRenderedPageBreak)
        run.append(t[0])
        paragraph.append(run)
    # Return the combined paragraph
    return paragraph
    
def heading(headingtext, headinglevel=1, lang='en'):
    '''Make a new heading, return the heading element'''
    lmap = {'en': 'Heading', 'it': 'Titolo'}
    # Make our elements
    paragraph = makeelement('p')
    pr = makeelement('pPr')
    pStyle = makeelement('pStyle', attributes={'val': lmap[lang] + str(headinglevel)})
    run = makeelement('r')
    text = makeelement('t', tagtext=headingtext)
    # Add the text the run, and the run to the paragraph
    pr.append(pStyle)
    run.append(text)
    paragraph.append(pr)
    paragraph.append(run)
    # Return the combined paragraph
    return paragraph
    
def table(contents, heading=True, colw=None, cwunit='dxa', tblw=0, twunit='auto', borders={}, celstyle=None):
    """
    Return a table element based on specified parameters

    @param list contents: A list of lists describing contents. Every item in
                        the list can be a string or a valid XML element
                        itself. It can also be a list. In that case all the
                        listed elements will be merged into the cell.
    @param bool heading:  Tells whether first line should be treated as
                        heading or not
    @param list colw:     list of integer column widths specified in wunitS.
    @param str  cwunit:   Unit used for column width:
                            'pct'  : fiftieths of a percent
                            'dxa'  : twentieths of a point
                            'nil'  : no width
                            'auto' : automagically determined
    @param int  tblw:     Table width
    @param int  twunit:   Unit used for table width. Same possible values as
                        cwunit.
    @param dict borders:  Dictionary defining table border. Supported keys
                        are: 'top', 'left', 'bottom', 'right',
                        'insideH', 'insideV', 'all'.
                        When specified, the 'all' key has precedence over
                        others. Each key must define a dict of border
                        attributes:
                            color : The color of the border, in hex or
                                    'auto'
                            space : The space, measured in points
                            sz    : The size of the border, in eighths of
                                    a point
                            val   : The style of the border, see
                http://www.schemacentral.com/sc/ooxml/t-w_ST_Border.htm
    @param list celstyle: Specify the style for each colum, list of dicts.
                        supported keys:
                        'align' : specify the alignment, see paragraph
                                    documentation.
    @return lxml.etree:   Generated XML etree element
    """
    table = makeelement('tbl')
    columns = len(contents[0])
    # Table properties
    tableprops = makeelement('tblPr')
    tablestyle = makeelement('tblStyle', attributes={'val': ''})
    tableprops.append(tablestyle)
    tablewidth = makeelement('tblW', attributes={'w': str(tblw), 'type': str(twunit)})
    tableprops.append(tablewidth)
    if len(borders.keys()):
        tableborders = makeelement('tblBorders')
        for b in ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']:
            if b in borders.keys() or 'all' in borders.keys():
                k = 'all' if 'all' in borders.keys() else b
                attrs = {}
                for a in borders[k].keys():
                    try:
                        attrs[a] = unicode(borders[k][a])
                    except NameError:
                        attrs[a] = borders[k][a]
                borderelem = makeelement(b, attributes=attrs)
                tableborders.append(borderelem)
        tableprops.append(tableborders)
    tablelook = makeelement('tblLook', attributes={'val': '0400'})
    tableprops.append(tablelook)
    table.append(tableprops)
    # Table Grid
    tablegrid = makeelement('tblGrid')
    for i in range(columns):
        tablegrid.append(makeelement('gridCol', attributes={'w': str(colw[i]) if colw else '2390'}))
    table.append(tablegrid)
    # Heading Row
    row = makeelement('tr')
    rowprops = makeelement('trPr')
    cnfStyle = makeelement('cnfStyle', attributes={'val': '000000100000'})
    rowprops.append(cnfStyle)
    row.append(rowprops)
    if heading:
        for i, heading in enumerate(contents[0]):
            cell = makeelement('tc')
            # Cell properties
            cellprops = makeelement('tcPr')
            if colw:
                wattr = {'w': str(colw[i]), 'type': cwunit}
            else:
                wattr = {'w': '0', 'type': 'auto'}
            cellwidth = makeelement('tcW', attributes=wattr)
            cellprops.append(cellwidth)
            cell.append(cellprops)
            # Paragraph (Content)
            if not isinstance(heading, (list, tuple)):
                heading = [heading]
            for h in heading:
                if isinstance(h, etree._Element):
                    cell.append(h)
                else:
                    cell.append(paragraph(h, pprops={'jc':{'val': 'center'}}))
            row.append(cell)
        table.append(row)
    # Contents Rows
    for contentrow in contents[1 if heading else 0:]:
        row = makeelement('tr')
        for i, content in enumerate(contentrow):
            cell = makeelement('tc')
            # Properties
            cellprops = makeelement('tcPr')
            if colw:
                wattr = {'w': str(colw[i]), 'type': cwunit}
            else:
                wattr = {'w': '0', 'type': 'auto'}
            cellwidth = makeelement('tcW', attributes=wattr)
            cellprops.append(cellwidth)
            cell.append(cellprops)
            # Paragraph (Content)
            if not isinstance(content, (list, tuple)):
                content = [content]
            for c in content:
                if isinstance(c, etree._Element):
                    cell.append(c)
                else:
                    if celstyle and 'align' in celstyle[i].keys():
                        align = celstyle[i]['align']
                    else:
                        align = 'left'
                    cell.append(paragraph(c, pprops={'jc':{'val': 'center'}}))
            row.append(cell)
        table.append(row)
    return table
    
def picture(document, picpath, picdescription='', pixelwidth=None, pixelheight=None, nochangeaspect=True, nochangearrowheads=True):
    '''Take a document and a picture file path, and return a paragraph
    containing the image. The document argument is necessary because we
    need to update the Relationships element when a picture is added'''
    # http://openxmldeveloper.org/articles/462.aspx
    # Create an image. Size may be specified, otherwise it will based on the
    # pixel size of image.
    # Copy the file into the media dir
    if not os.path.isdir(document.media_dir):
        os.mkdir(document.media_dir)
    picname = os.path.basename(picpath)
    shutil.copyfile(picname, os.path.join(document.media_dir, picname))
    # Check if the user has specified a size
    if not pixelwidth or not pixelheight:
        pixelwidth, pixelheight = helper_functions.get_image_size(picpath)
    # OpenXML measures on-screen objects in English Metric Units
    # 1cm = 36000 EMUs
    emuperpixel = 12700
    width = str(pixelwidth * emuperpixel)
    height = str(pixelheight * emuperpixel)
    # Set relationship ID to the first available
    picid = '2'
    rId = helper_functions.add_relationship(document,
                                       os.path.join('media', picname),
                                       'http://schemas.openxmlformats.org/'
                                       'officeDocument/2006/relationships/'
                                       'image')
    # There are 3 main elements inside a picture
    # 1. The Blipfill - specifies how the image fills the picture area (stretch, tile, etc.)
    blipfill = makeelement('blipFill', nsprefix='pic')
    blipfill.append(makeelement('blip', nsprefix='a', attrnsprefix='r',
                    attributes={'embed': rId}))
    stretch = makeelement('stretch', nsprefix='a')
    stretch.append(makeelement('fillRect', nsprefix='a'))
    blipfill.append(makeelement('srcRect', nsprefix='a'))
    blipfill.append(stretch)

    # 2. The non visual picture properties
    nvpicpr = makeelement('nvPicPr', nsprefix='pic')
    cnvpr = makeelement('cNvPr', nsprefix='pic',
                        attributes={'id': '0', 'name': 'Picture 1', 'descr': picname})
    nvpicpr.append(cnvpr)
    cnvpicpr = makeelement('cNvPicPr', nsprefix='pic')
    cnvpicpr.append(makeelement('picLocks', nsprefix='a',
                    attributes={'noChangeAspect': str(int(nochangeaspect)),
                                'noChangeArrowheads': str(int(nochangearrowheads))}))
    nvpicpr.append(cnvpicpr)

    # 3. The Shape properties
    sppr = makeelement('spPr', nsprefix='pic', attributes={'bwMode': 'auto'})
    xfrm = makeelement('xfrm', nsprefix='a')
    xfrm.append(makeelement('off', nsprefix='a', attributes={'x': '0', 'y': '0'}))
    xfrm.append(makeelement('ext', nsprefix='a', attributes={'cx': width, 'cy': height}))
    prstgeom = makeelement('prstGeom', nsprefix='a', attributes={'prst': 'rect'})
    prstgeom.append(makeelement('avLst', nsprefix='a'))
    sppr.append(xfrm)
    sppr.append(prstgeom)

    # Add our 3 parts to the picture element
    pic = makeelement('pic', nsprefix='pic')
    pic.append(nvpicpr)
    pic.append(blipfill)
    pic.append(sppr)

    # Now make the supporting elements
    # The following sequence is just: make element, then add its children
    graphicdata = makeelement('graphicData', nsprefix='a',
                            attributes={'uri': 'http://schemas.openxmlforma'
                                                'ts.org/drawingml/2006/picture'})
    graphicdata.append(pic)
    graphic = makeelement('graphic', nsprefix='a')
    graphic.append(graphicdata)

    framelocks = makeelement('graphicFrameLocks', nsprefix='a',
                            attributes={'noChangeAspect': '1'})
    framepr = makeelement('cNvGraphicFramePr', nsprefix='wp')
    framepr.append(framelocks)
    docpr = makeelement('docPr', nsprefix='wp',
                        attributes={'id': picid, 'name': 'Picture 1',
                                    'descr': picdescription})
    effectextent = makeelement('effectExtent', nsprefix='wp',
                            attributes={'l': '25400', 't': '0', 'r': '0',
                                        'b': '0'})
    extent = makeelement('extent', nsprefix='wp',
                        attributes={'cx': width, 'cy': height})
    inline = makeelement('inline', attributes={'distT': "0", 'distB': "0",
                                            'distL': "0", 'distR': "0"},
                        nsprefix='wp')
    inline.append(extent)
    inline.append(effectextent)
    inline.append(docpr)
    inline.append(framepr)
    inline.append(graphic)
    drawing = makeelement('drawing')
    drawing.append(inline)
    run = makeelement('r')
    run.append(drawing)
    paragraph = makeelement('p')
    paragraph.append(run)
    return paragraph
    
def append_text(element, text):		
    if element.tag == '{' + NSPREFIXES['w'] + '}body':
        try:
            last_para = [child for child in element.iterchildren() if child.tag == '{' + NSPREFIXES['w'] + '}p'][-1]
        except IndexError:
            element.append(paragraph(text))
            return
        try:
            last_run = [child for child in last_para.iterchildren() if child.tag == '{' + NSPREFIXES['w'] + '}r'][-1]
        except IndexError:
            last_run = makeelement('r')
            last_para.append(last_run)
        try:
            last_text = [child for child in last_run.iterchildren() if child.tag == '{' + NSPREFIXES['w'] + '}t'][-1]
            last_text.text += text
        except IndexError:
            last_text = makeelement('t', tagtext=text)
            last_run.append(last_text)
    elif element.tag == '{' + NSPREFIXES['w'] + '}p':
        try:
            last_run = [child for child in element.iterchildren('{' + NSPREFIXES['w'] + '}r')][-1]
        except IndexError:
            last_run = makeelement('r')
            element.append(last_run)
        try:
            last_text = [child for child in last_run.iterchildren('{' + NSPREFIXES['w'] + '}t')][-1]
            last_text.text += text
        except IndexError:
            last_text = makeelement('t', tagtext=text)
            last_run.append(last_text)
    elif element.tag == '{' + NSPREFIXES['w'] + '}r':
        try:
            last_text = [child for child in element.iterchildren('{' + NSPREFIXES['w'] + '}t')][-1]
            last_text.text += text
        except IndexError:
            last_text = makeelement('t', tagtext=text)
            element.append(last_text)
    elif element.tag == '{' + NSPREFIXES['w'] + '}t':
        element.text += text
        
def numbered_list(start, end=None):
    '''Creates a numbered list containing all of the paragraphs between
    the start and end paragraph elements, inclusively'''
    if end is None:
        end = start
    body = start.getparent()
    if body != end.getparent():
        raise ValueError('start and end elements must have the same parent')
    if start.tag != '{' + NSPREFIXES['w'] + '}p':
        raise ValueError('start argument must be a paragraph element')
    if end.tag != '{' + NSPREFIXES['w'] + '}p':
        raise ValueError('end argument must be a paragraph element')
    if body.index(start) > body.index(end):
        raise ValueError('end paragraph cannot precede start paragraph')
    para_list = [para for para in body.iterchildren() if body.index(para) in
    range(body.index(start), body.index(end) + 1)]
    numId_set = set()
    for element in body.iter('{' + NSPREFIXES['w'] + '}numId'):
        for k, v in element.items():
            if k == '{' + NSPREFIXES['w'] + '}val':
                numId_set.add(v)
    numId_value = '1'
    while numId_value in numId_set:
        numId_value = str(int(numId_value) + 1)
    for para in para_list:
        pPr = makeelement('pPr')
        for child in para.iterchildren('{' + NSPREFIXES['w'] + '}pPr'):
            pPr = child
            break
        if pPr.getparent() is None:
            para.insert(0, pPr)
        numPr = makeelement('numPr')
        ilvl = makeelement('ilvl', attributes={'val': '0'})
        numId = makeelement('numId', attributes={'val': numId_value})
        numPr.append(ilvl)
        numPr.append(numId)
        pPr.insert(0, numPr)	

def add_comment(document, text, start, end=None, username='', initials=''):
    if end is None:
        end = start
    else:
        sparent = start.getparent()
        sgparent = start.getparent().getparent()
        sggparent = start.getparent().getparent().getparent()
        eparent = end.getparent()
        egparent = end.getparent().getparent()
        eggparent = end.getparent().getparent().getparent()
        if start.tag == end.tag:
            if start.tag == '{' + NSPREFIXES['w'] + '}p':
                if sparent.index(start) > eparent.index(end):
                    raise ValueError('end element cannot precede start '
                                     'element')
            elif start.tag == '{' + NSPREFIXES['w'] + '}r':		
                if sgparent.index(sparent) > egparent.index(eparent):
                    raise ValueError('end element cannot precede start '
                                     'element')
            elif start.tag == '{' + NSPREFIXES['w'] + '}t':		
                if sggparent.index(sgparent) > eggparent.index(egparent):
                    raise ValueError('end element cannot precede start '
                                     'element')
        elif sparent.tag == end.tag and (sgparent.index(sparent) >
        eparent.index(eparent)):
            raise ValueError('end element cannot precede start element')
        elif sgparent.tag == end.tag and (sggparent.index(sgparent) >
        eparent.index(eparent)):
            raise ValueError('end element cannot precede start element')
        elif start.tag == eparent.tag and (sparent.index(start) >
        egparent.index(eparent)):
            raise ValueError('end element cannot precede start element')
        elif start.tag == egparent.tag and (sparent.index(start) >
        eggparent.index(egparent)):
            raise ValueError('end element cannot precede start element')
    id_number = write_files.setup_comments(document)
    if start.tag == '{' + NSPREFIXES['w'] + '}p': # Insert commentRangeStart element
        start.insert(0, makeelement('commentRangeStart',
        attributes={'id': id_number}))
    elif start.tag == '{' + NSPREFIXES['w'] + '}r':
        paragraph = start.getparent()
        pos = paragraph.index(start)
        paragraph.insert(pos, makeelement('commentRangeStart',
        attributes={'id': id_number}))
    elif start.tag == '{' + NSPREFIXES['w'] + '}t':
        run = start.getparent()
        paragraph = run.getparent()
        text_pos = run.index(start)
        run_pos = paragraph.index(run)
        text_elements = [child for child in run.iterchildren() if child.tag ==
        '{' + NSPREFIXES['w'] + '}t']
        if text_pos != 0:
            preceding_run = makeelement('r')
            for element_text in text_elements[:text_pos - 1]:
                run.remove(element_text)
                preceding_run.append(element_text)
            paragraph.insert(run_pos, preceding_run)
            run_pos += 1
        paragraph.insert(run_pos, makeelement('commentRangeStart',
        attributes={'id': id_number}))
    if end.tag == '{' + NSPREFIXES['w'] + '}p': # Insert commentRangeEnd element
        end.append(makeelement('commentRangeEnd',
        attributes={'id': id_number}))
        run = makeelement('r')
        end.append(run)
        rPr = makeelement('rPr')
        run.append(rPr)
        rPr.append(makeelement('rStyle',
        attributes={'val': 'CommentReference'}))
        run.append(makeelement('commentReference',
        attributes={'id': id_number}))
    elif end.tag == '{' + NSPREFIXES['w'] + '}r':
        paragraph = end.getparent()
        pos = paragraph.index(end)
        paragraph.insert(pos + 1, makeelement('commentRangeEnd',
        attributes={'id': id_number}))
        run = makeelement('r')
        end.append(run)
        rPr = makeelement('rPr')
        run.append(rPr)
        rPr.append(makeelement('rStyle',
        attributes={'val': 'CommentReference'}))
        run.append(makeelement('commentReference',
        attributes={'id': id_number}))
    elif end.tag == '{' + NSPREFIXES['w'] + '}t':
        run = end.getparent()
        paragraph = run.getparent()
        text_pos = run.index(end)
        run_pos = paragraph.index(run)
        text_elements = [child for child in
                         run.iterchildren('{' + NSPREFIXES['w'] + '}t')]
        if text_pos != 0:
            preceding_run = makeelement('r')
            for element_text in text_elements[:text_pos - 1]:
                run.remove(element_text)
                preceding_run.append(element_text)
            paragraph.insert(run_pos, preceding_run)
            run_pos += 1
        paragraph.insert(run_pos + 1, makeelement('commentRangeEnd',
        attributes={'id': id_number}))
        run = makeelement('r')
        paragraph.insert(run_pos + 2, run)
        rPr = makeelement('rPr')
        run.append(rPr)
        rPr.append(makeelement('rStyle',
        attributes={'val': 'CommentReference'}))
        run.append(makeelement('commentReference',
        attributes={'id': id_number}))
    date = datetime.datetime.now()              # Content for comments.xml
    daystr = str(date.day)
    hourstr = str(date.hour)
    minutestr = str(date.minute)
    if len(daystr) == 1:
        daystr = '0' + daystr
    if len(hourstr) == 1:
        hourstr = '0' + hourstr
    if len(minutestr) == 1:
        minutestr = '0' + minutestr
    comment = makeelement('comment', attributes={'id': id_number, 
    'author': username, 'date': '{0}-{1}-{2}T{3}:{4}:00Z'.format(
    str(date.year), str(date.month), daystr, hourstr, minutestr),
    'initials': initials})
    para = makeelement('p')
    comment.append(para)
    pPr = makeelement('pPr')
    para.append(pPr)
    pPr.append(makeelement('pStyle', attributes={'val': 'CommentText'}))
    run_reference = makeelement('r')
    para.append(run_reference)
    rPr = makeelement('rPr')
    run_reference.append(rPr)
    rPr.append(makeelement('rStyle', attributes={'val': 'CommentReference'}))
    run_reference.append(makeelement('annotationRef'))
    run_text = makeelement('r')
    para.append(run_text)
    run_text.append(makeelement('t', tagtext=text))
    document.comments.append(comment)
    
    
def get_text(element):
    '''Returns a single string of text which is a concatenation of all
    of the text contained by text elements in the argument element'''
    if element.tag == '{' + NSPREFIXES['w'] + '}t':
        return element.text
    else:
        text_list = []
        for descendant in element.iter('{' + NSPREFIXES['w'] + '}t'):
            if descendant.text:
                text_list.append(descendant.text)
        return ''.join(text_list)
  
def remove_formatting(element):
    for descendant in element.iter():
        if (descendant.tag == '{' + NSPREFIXES['w'] + '}pPr' or 
        descendant.tag == '{' + NSPREFIXES['w'] + '}rPr'):
            descendant.getparent().remove(descendant)        
