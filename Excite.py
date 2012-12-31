"""Excite: External Citation processor

Takes documents such as Apple Pages .pages files, processes LaTeX
bibliographies embedded in the documents."""

import zipfile
import shutil
import re

from xml.etree import ElementTree
from collections import defaultdict

def AllText(node):
    """Get all text from an ElementTree node."""
    text = ""
    for t in node.itertext():
        text += t
    return text


class Bibliography(object):
    """Represents the bibliography. Notes the order of citations and renders the
    bibliography in that order."""

    def __init__(self):
        self.citationorder = {}
        self.references = defaultdict(str)

    def AddCitation(self, label):
        assert type(label) is str

        try:
            self.citationorder[label]
        except KeyError:
            vals = self.citationorder.values()
            vals.append(0)
            self.citationorder[label] = max(vals) + 1

    def CitationIndex(self, label):
        return self.citationorder[label]

    def AddReference(self, label, text):
        assert type(label) is str
        assert type(text) is str

        if self.references.has_key(label):
            raise KeyError('Duplicate references.')

        self.references[label] = text.strip()

    def GetReferenceByLabel(self, label):
        return (label, self.CitationIndex(label), self.references[label])

    def GetReferenceByIndex(self, index):
        assert index <= self.ItemCount()

        for label in self.citationorder:
            if self.citationorder[label] == index:
                return (label, index, self.references[label])

    def IsConsistent(self):
        """Returns True if the citations and references are consistent with each other.
        This method is intended to be called after a full pass of a document has been completed."""

        return set(self.citationorder.keys()) == set(self.references.keys())

    def ItemCount(self):
        return max(len(self.citationorder), self.references)

class WordProcessingDocument(object):
    """Represents a generic document that is supported by this system."""
    
    PRIMARYDOCUMENT = None
    supportedstyles = ()

    def __init__(self, filename):
        self.filename = filename
        self.citationformat = r"\\cite\{(\w+)\}"
        self.bibformat = r'\\bibitem\{(\w+)\} ?(.*)'

    def ProcessCitations(self, style):
        """Internally construct a version of the document that has the citations properly created."""
        raise NotImplementedError

    def Materialize(self, outputfile):
        """Generate a completed file at thie indicated outputfile."""
        raise NotImplementedError

class ApplePages(WordProcessingDocument):
    """Represents a processor for an Apple iWork Pages document."""
    
    PRIMARYDOCUMENT = 'index.xml' # Where in the zip archive is the primary XML document located
    supportedstyles = ('squarebrace')

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
            self.document = ElementTree.XML(pageszip.read(ApplePages.PRIMARYDOCUMENT))

    def ProcessCitations(self, style='squarebrace'):
        assert style in self.supportedstyles

        bibliography = Bibliography()
        citationnodes = []
        bibnodes = []

        # ElementTree represents namespaces like so: sf:p -> {http://developer.apple.com/namespaces/sf}p
        for node in self.document.findall(self.ns('.//sf:text-body//sf:p')):
            searchtext = AllText(node)
            citationmatch = re.findall(self.citationformat, searchtext)
            
            if len(citationmatch):
                citationnodes.append(node)
                for label in citationmatch:
                    bibliography.AddCitation(label)

            bibitemmatch = re.search(self.bibformat, searchtext)

            if bibitemmatch:
                bibnodes.append(node)
                bibliography.AddReference(bibitemmatch.group(1), bibitemmatch.group(2))

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
                replacementtext = self.RenderCitation(style=style, index=bibliography.CitationIndex(label))
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

        if style == 'squarebrace':
            return "[{0}]".format(index)

        return str(index)

    def RenderReference(self, style, index, text):
        """Produces the final XML that represents the reference."""

        return "{0}. {1}".format(index, text)

    def Materialize(self, outputfilename):
        if self.filename != outputfilename:
            shutil.copy2(self.filename, outputfilename)

        with zipfile.ZipFile(outputfilename, 'a') as pageszip:
            finalxml = ElementTree.tostring(element=self.document, encoding='us-ascii', method='xml')
            # finalxml should be bytes, which is why us-ascii was chosen as the encoding
            pageszip.writestr(ApplePages.PRIMARYDOCUMENT, finalxml) 

    def ns(self, string):
        """Provides transformation of namespaced tags into something ElementTree can understand."""
        for namespace in ApplePages.xmlnamespaces.keys():
            pattern = namespace + ":"
            string = re.sub(pattern, "{{{0}}}".format(ApplePages.xmlnamespaces[namespace]), string)

        return string



