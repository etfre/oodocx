<h1>oodocx</h1>

Note: I cut my programming teeth with this project. As a result, the code
is questionably written, poorly modularized, and lacks anything even remotely
resembling a testing suite. That said, everything works as it should,
and it contains several features that the superiorly constructed
<a href="https://github.com/python-openxml/python-docx">python-docx</a> currently
lacks. As a result, for the time being at least, this repository will remain available.
Here are some of the differences of oodocx from python-docx:

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
    
    
