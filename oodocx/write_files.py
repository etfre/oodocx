from lxml import etree

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