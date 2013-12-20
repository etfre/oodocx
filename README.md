<h1>oodocx</h1>

oodocx is a fork/modification/Python 3 port of Mike MacCana's excellent
<a href="https://github.com/mikemaccana/python-docx">python-docx</a> module.
Here are some of the differences of oodocx from python-docx:
<p>
1. Object Oriented. As the name suggests, oodocx has an object-oriented aspect.
The Docx class represents the various files that comprise a docx file,
which is simply a zip file that contains primarily xml files.
The document.xml file is the main document that holds most of the content that
you find in a word document. However, this module also gives you easy access to
most of the other files that typically comprise a docx file as attributes of
the Docx class. For example, the root element in styles.xml of a Docx object
named "d" can be accessed as d.styles.
This allows for greater control of the look, feel, and metadata
of a docx file.
</p>
<p>2. Python 3 compatible.</p>
<p>3. Expanded functionality, particularly for people with limited knowledge of xml
who want to get started quickly on their docx scripting projects. Some examples include
the modify_font and modify_paragraph functions, which allow for easy modification of 
common font and paragraph properties of an element or a list of elements.</p>
<p>4. oodocx keeps all of the files in the Docx zip, rather than just the 
document.xml file. This helps to preserve formatting and ensures that pictures, comments,
and other elements in a Docx won't break from save to save.</p>

<h2>Installation</h2>
First, ensure that you have the appropriate version of lxml installed. Then,
clone this repository, navigate your shell window to the oodocx directory
folder that contains the setup.py file, and enter "python setup.py install"
(or just "setup.py install", depending on how you execute python files on your
computer)

<h2>How do I...</h2>
First of all, be sure to check out the /examples folder for basic examples of this module
  <h3>Create a new document, insert a paragraph of text, and save it</h3>
  
    d = oodocx.Docx()
    d.body.append(oodocx.paragraph('Hello world!'))
    d.save('hello.docx')

  <h3>Open a document, insert a paragraph after the paragraph containing the word "apple", and save it somewhere else</h3>
  
    d = oodocx.Docx(r'C:\users\applecart\apples.docx')
    apple_para = d.search('apple', result_type='paragraph')
    pos = body.index(apple_para) + 1 #lxml
    d.body.insert(pos, oodocx.paragraph('Bananas!')) #lxml
    d.save(r'C:\users\bananstand\there's always money here.docx')
    
Note that the index() and insert() methods in the fourth and fifth lines of the above code are from the underlying lxml module. Check out the documentation <a href='http://lxml.de/api/lxml.etree._Element-class.html'>here</a>.
    
    
