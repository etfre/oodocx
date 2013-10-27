<h1>oodocx</h1>

oodocx is a fork/modification/port of Mike MacCana's excellent <a href="https://github.com/mikemaccana/python-docx">python-docx</a> module.

1. Object Oriented. As the name suggests, oodocx has an object-oriented aspect.
The Docx class represents the various files that comprise a docx file,
which is simply a zip file that contains primarily xml files.
The document.xml file is the main document that holds most of the stuff that
you find in a word document. However, this module also gives you easy access to
most of the other files that typically comprise a docx file as attributes of
the Docx class. For example, the root element in styles.xml  of a Docx object
called "d" can be accessed with 
    d.styles
. This allows for greater control of the look, feel, and metadata
of a docx file
2. Python 3 compatible
3. Expanded functionality
4. 

<h2>Installation</h2>
Clone this repository, navigate your shell window to the oodocx directory folder that contains the setup.py file, and enter "python setup.py install" 
(or just "setup.py install", depending on how you execute python files on your computer)

<h2>How do I...</h2>
First of all, be sure to check out the /examples folder to
  <h3>Create a new document, insert a paragraph of text, and save it</h3>
  
    d = Docx()
    body = d.get_body()
    body.append(paragraph('Hello world!'))
    d.save('hello.docx')

  <h3>Open a document, insert a paragraph after the paragraph containing the word "apple", and save it somewhere else</h3>
  
    d = Docx(r'C:\users\appleman\applecart\apples.docx')
    body = d.get_body
    apple_para = d.search('apple', result_type='paragraph')
    pos = body.index(apple_para) + 1 #lxml
    body.insert(pos, paragraph('Bananas!')) #lxml
    d.save(r'C:\users\bananstand\there's always money here.docx')
    
Note that the index() and insert() methods in the fourth and fifth lines of the above code are from the underlying lxml module. Check out the documentation <a href='http://lxml.de/api/lxml.etree._Element-class.html'>here</a>.
    
    
