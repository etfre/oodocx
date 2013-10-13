<h1>oodocx</h1>

oodocx is a fork/modification/port of Mike MacCana's excellent <a href="https://github.com/mikemaccana/python-docx">python-docx</a> module.

1. Object Oriented. As the name suggests, oodocx has an object-oriented aspect.
The Docx class acts as a wrapper for the ElementTree class from the underlying lxml module on which this module is built.
2. Python 3 compatible
3. Expanded functionality
4. 

<h2>Installation</h2>
pass

<h2>How do I...</h2>
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
    
    
