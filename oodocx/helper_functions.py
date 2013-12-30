import struct
import imghdr
import stat
import os
from lxml import etree
	
def get_image_size(fname):
	'''Determine the image type of fhandle and return its size.
	from draco'''
	fhandle = open(fname, 'rb')
	head = fhandle.read(24)
	if len(head) != 24:
		return
	if imghdr.what(fname) == 'png':
		check = struct.unpack('>i', head[4:8])[0]
		if check != 0x0d0a1a0a:
			return
		width, height = struct.unpack('>ii', head[16:24])
	elif imghdr.what(fname) == 'gif':
		width, height = struct.unpack('<HH', head[6:10])
	elif imghdr.what(fname) == 'jpeg':
		try:
			fhandle.seek(0) # Read 0xff next
			size = 2
			ftype = 0
			while not 0xc0 <= ftype <= 0xcf:
				fhandle.seek(size, 1)
				byte = fhandle.read(1)
				while ord(byte) == 0xff:
					byte = fhandle.read(1)
				ftype = ord(byte)
				size = struct.unpack('>H', fhandle.read(2))[0] - 2
			# We are at a SOFn block
			fhandle.seek(1, 1)  # Skip `precision' byte.
			height, width = struct.unpack('>HH', fhandle.read(4))
		except Exception: #IGNORE:W0703
			return
	else:
		return
	return width, height
	
def remove_readonly(fn, path, excinfo):
    if fn is os.rmdir:
        os.chmod(path, stat.S_IWRITE)
        os.rmdir(path)
    elif fn is os.remove:
        os.chmod(path, stat.S_IWRITE)
        os.remove(path)
        
def add_relationship(document, target, type):
    '''checks Relationships element to see if element is included,
    adds it if not, returns element's rId or None'''
    relationship_items = [child.items() for child in document.relationships.getchildren()]
    flat_relationships = sum(relationship_items, [])
    id_numbers = sorted([int(item[1][3:]) for item in flat_relationships if item[0] == 'Id'])
    rId_number = len(id_numbers) + 1
    for count, number in enumerate(id_numbers, start=1):
        if count != number:
            rId_number = count + 1
            break
    if target not in [child[1] for child in flat_relationships] or 'media' in target:
        document.relationships.append(makeelement('Relationship', nsprefix=None,
        attributes={'Id': 'rId' + str(rId_number),
                    'Target': target,
                    'Type': type}))
        return 'rId' + str(rId_number)
    else:
        return None        

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
