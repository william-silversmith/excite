"""Excite: External Citation processor

Takes documents such as Apple Pages .pages files, processes LaTeX
bibliographies embedded in the documents."""

import zipfile
import shutil
import re

from xml.etree import ElementTree
from collections import defaultdict

def alltext(node):
    """Get all text from an ElementTree node."""
    text = ""
    for t in node.itertext():
        text += t
    return text


class Bibliography(object):
    """Represents the bibliography. Notes the order of citations and renders the
    bibliography in that order."""

    def __init__(self, orderby='citation-first'):
        assert orderby in ('citation-first', 'reference-first')

        self.order = {}
        self.orderby = orderby
        self.citations = []
        self.references = defaultdict(str)

    def AddCitation(self, label):
        assert type(label) is str

        self.citations.append(label)

        if self.orderby == 'citation-first':
            self.__MaybeUpdateOrder(label)

    def AddReference(self, label, reference):
        assert type(label) is str

        if self.references.has_key(label):
            raise KeyError('Duplicate references.')

        self.references[label] = reference

        if self.orderby == 'reference-first':
            self.__MaybeUpdateOrder(label)

    def __MaybeUpdateOrder(self, label):
        """Associate the given label with the current maximum sequence number plus one if it has not
        already been encountered."""
        try:
            self.order[label]
        except KeyError:
            vals = self.order.values()
            vals.append(0)
            self.order[label] = max(vals) + 1

    def Index(self, label):
        """Return the index associated with a label."""
        return self.order[label]

    def GetReferenceByLabel(self, label):
        """Get (label, index, reference representation) by the reference's label."""
        return (label, self.Index(label), self.references[label])

    def GetReferenceByIndex(self, index):
        """Get (label, index, reference representation) by the reference's index."""
        assert index <= self.ItemCount()

        for label in self.order:
            if self.order[label] == index:
                return (label, index, self.references[label])

    def IsConsistent(self):
        """Returns True if the citations and references are consistent with each other.
        This method is intended to be called after a full pass of a document has been completed."""

        return set(self.order.keys()) == set(self.references.keys())

    def ItemCount(self):
        return max(len(self.order), len(self.references))

class WordProcessingDocument(object):
    """Represents a generic document that is supported by this system."""
    
    primarydocument = None
    supportedstyles = ()

    def __init__(self, filename):
        self.filename = filename
        self.citationformat = r"\\cite\{(\w+)\}"
        self.bibformat = r'\\bibitem\{(\w+)\} ?(.*)'

    def ProcessCitations(self, style, orderby):
        """Internally construct a version of the document that has the citations properly created."""
        raise NotImplementedError

    def Materialize(self, outputfile):
        """Generate a completed file at thie indicated outputfile."""
        raise NotImplementedError

class ApplePages(WordProcessingDocument):
    """Represents a processor for an Apple iWork Pages document."""
    
    primarydocument = 'index.xml' # Where in the zip archive is the primary XML document located
    supportedstyles = ('square-brace')

    # XML namespaces present in .pages XML
    xmlnamespaces = {
        "sf": "http://developer.apple.com/namespaces/sf",
        "sfa": "http://developer.apple.com/namespaces/sfa", 
        "xsi": "http://www.w3.org/2001/XMLSchema-instance",
        "sl": "http://developer.apple.com/namespaces/sl",
    }

    def __init__(self, filename):
        super(ApplePages, self).__init__(filename)

        for (identifier, namespace) in ApplePages.xmlnamespaces.items():
            ElementTree.register_namespace(identifier, namespace)

        with zipfile.ZipFile(filename, 'r') as pageszip:
            self.document = ElementTree.XML(pageszip.read(ApplePages.primarydocument))

    def ProcessCitations(self, style='square-brace', order='citation-first'):
        """Internally construct a version of the document that has the citations properly created."""
        assert style in self.supportedstyles
        assert order in ('citation-first', 'reference-first')

        bibliography = Bibliography()
        citationnodes = []
        bibnodes = []

        # ElementTree represents namespaces like so: sf:p -> {http://developer.apple.com/namespaces/sf}p
        for node in self.document.findall(self.ns('.//sf:text-body//sf:p')):
            searchtext = alltext(node)
            citationmatch = re.findall(self.citationformat, searchtext)
            
            if len(citationmatch):
                citationnodes.append(node)
                for label in citationmatch:
                    bibliography.AddCitation(label)

            bibitemmatch = re.search(self.bibformat, searchtext)

            if bibitemmatch:
                bibnodes.append(node)
                bibliography.AddReference(bibitemmatch.group(1), bibitemmatch.group(2).strip())

        if bibliography.IsConsistent() == False:
            raise ValueError("Citations and references are not one-to-one.")
        
        self.__ReplaceCitationMarkers(style=style, citationnodes=citationnodes, bibliography=bibliography)
        self.__ReplaceBibitemMarkers(style=style, bibnodes=bibnodes, bibliography=bibliography)

    def __ReplaceCitationMarkers(self, style, citationnodes, bibliography):
        """Helper method for ProcessCitations. Changes document to replace
        citations with the correct number rendered in a particular style.

        Requires:
            style: A supported rendering style
            citationnodes: A list of Element objects representing a node containing a citation.
            bibliography: A completed (consistent) Bibliography object
        
        Returns: Void
        """
        def replacecitation(text):
            if text == None:
                return None

            for label in re.findall(self.citationformat, text):
                replacementtext = self.RenderCitation(style=style, index=bibliography.Index(label))
                replacementpattern = r'\\cite\{' + label + r'\}'
                text = re.sub(replacementpattern, replacementtext, text)

            return text

        def traverse(node):
            node.text = replacecitation(node.text)
            node.tail = replacecitation(node.tail)
            for nextnode in node.getchildren():
                traverse(nextnode)

        for node in citationnodes:
            traverse(node)
    
    def __ReplaceBibitemMarkers(self, style, bibnodes, bibliography):
        """Helper method for ProcessCitations. Changes document to replace
        bibitems with the correct number and reference rendered in a particular style.

        Requires:
            style: A supported rendering style
            bibnodes: A list of Element objects representing a node containing a reference.
            bibliography: A completed (consistent) Bibliography object
        
        Returns: Void
        """
        sequencenumber = 1
        for node in bibnodes:
            (label, index, text) = bibliography.GetReferenceByIndex(sequencenumber)
            
            node.text = self.RenderReference(style=style, index=index, text=text)
            node.tail = None
            sequencenumber += 1

    def RenderCitation(self, style, index):
        """Produces the final XML that represents the citation."""

        assert style in ApplePages.supportedstyles

        if style == 'square-brace':
            return "[{0}]".format(index)

        return str(index)

    def RenderReference(self, style, index, text):
        """Produces the final XML that represents the reference."""

        return "{0}. {1}".format(index, text)

    def Materialize(self, outputfilename):
        """Serializes the internal representation into a new pages file."""
        if self.filename != outputfilename:
            shutil.copy2(self.filename, outputfilename)

        with zipfile.ZipFile(outputfilename, 'a') as pageszip:
            finalxml = ElementTree.tostring(element=self.document, encoding='us-ascii', method='xml')
            # finalxml should be bytes, which is why us-ascii was chosen as the encoding
            pageszip.writestr(ApplePages.primarydocument, finalxml) 

    def ns(self, string):
        """Provides transformation of namespaced tags into something ElementTree can understand."""
        for namespace in ApplePages.xmlnamespaces.keys():
            pattern = namespace + ":"
            string = re.sub(pattern, "{{{0}}}".format(ApplePages.xmlnamespaces[namespace]), string)

        return string



