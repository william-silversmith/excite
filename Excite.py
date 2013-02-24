"""Excite: External Citation processor

Takes documents such as Apple Pages .pages files, processes LaTeX
bibliographies embedded in the documents."""

import zipfile
import shutil
import re

from xml.etree import ElementTree
from collections import defaultdict

class MissingReferenceError(StandardError):
    def __init__(self, message, badcites):
        super(MissingReferenceError, self).__init__(message)
        self.message = message
        self.badcites = badcites

class DuplicateReferenceError(StandardError):
    def __init__(self, message, badrefs):
        super(DuplicateReferenceError, self).__init__(message)
        self.message = message
        self.badrefs = badrefs

def alltext(node):
    """Get all text from an ElementTree node."""
    text = u""
    for t in node.itertext():
        text += t
    return unicode(text)

def maybestr(obj):
    if obj is None:
        return u""
    return unicode(obj)

def copyelement(fromnode, tonode):
    """Replaces the information in an ElementTree.Element with another. This
    eases in-place insertion into a document tree when the parent and index
    of the child are unknown."""

    tonode.clear()
    tonode.tag = fromnode.tag
    tonode.text = fromnode.text
    tonode.tail = fromnode.tail
    tonode.attrib = fromnode.attrib

    for i, child in enumerate(fromnode):
        tonode.insert(i, child)

def traversetransform(node, transformingfunc):
    """Transforms the text of each node in an ElementTree.Element subtree via a transforming function
    that accepts a string and returns a transformed string."""
    node.text = transformingfunc(node.text)
    node.tail = transformingfunc(node.tail)
    for nextnode in node:
        traversetransform(nextnode, transformingfunc)

class Bibliography(object):
    """Represents the bibliography. Notes the order of citations and renders the
    bibliography in that order."""

    def __init__(self, orderby=u'citation-first'):
        assert orderby in (u'citation-first', u'reference-first')

        self.order = {}
        self.orderby = orderby
        self.citations = []
        self.references = defaultdict(str)

    def AddCitation(self, label):
        assert type(label) in (str, unicode)

        label = unicode(label)

        self.citations.append(label)

        if self.orderby == u'citation-first':
            self.__MaybeUpdateOrder(label)

    def AddReference(self, label, reference):
        assert type(label) in (str, unicode)
        label = unicode(label)

        if self.references.has_key(label):
            raise DuplicateReferenceError(message=u'Duplicate references found when adding a reference.', badrefs=[label])

        self.references[label] = reference

        if self.orderby == u'reference-first':
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
        return self.order[unicode(label)]

    def GetReferenceByLabel(self, label):
        """Get (label, index, reference representation) by the reference's label."""
        return (label, self.Index(label), self.references[unicode(label)])

    def GetReferenceByIndex(self, index):
        """Get (label, index, reference representation) by the reference's index."""
        assert index <= self.Count()

        for label in self.order:
            if self.order[label] == index:
                return (label, index, self.references[label])

    def IsConsistent(self):
        """Returns True if all the citations correspond to references.
        This method is intended to be called after a full pass of a document has been completed."""

        return set(self.citations).intersection(set(self.references.keys())) == set(self.citations)

    def Count(self):
        return max(len(self.order), len(self.references))

class WordProcessingDocument(object):
    """Represents a generic document that is supported by this system."""
    
    primarydocument = None
    supportedcitestyles = ()
    supportedbibstyles = ()

    def __init__(self, filename):
        self.filename = filename
        self.citationformat = ur"\\cite\{(\w+)\}"
        self.bibformat = ur'\\bibitem\{(\w+)\} ?(.*)'

    def ProcessCitations(self, style, orderby):
        """Internally construct a version of the document that has the citations properly created."""
        raise NotImplementedError

    def Materialize(self, outputfile):
        """Generate a completed file at thie indicated outputfile."""
        raise NotImplementedError

class ApplePages(WordProcessingDocument):
    """Represents a processor for an Apple iWork Pages document."""
    
    primarydocument = u'index.xml' # Where in the zip archive is the primary XML document located
    supportedcitestyles = (u'square-brace', u'superscript', u'parens')
    supportedbibstyles = (u'square-brace', u'digit-dot')

    styles = {
        # 50,000 is simply a high number as it's simpler and faster hope no one makes page that big than to have to figure out
        # a unique ID by parsing the tree. We can do that if necessary.
        'superscript': u"SFWPCharacterStyle-50000",
    }


    # XML namespaces present in .pages XML
    xmlnamespaces = {
        u"sf": u"http://developer.apple.com/namespaces/sf",
        u"sfa": u"http://developer.apple.com/namespaces/sfa", 
        u"xsi": u"http://www.w3.org/2001/XMLSchema-instance",
        u"sl": u"http://developer.apple.com/namespaces/sl",
    }

    def __init__(self, filename):
        super(ApplePages, self).__init__(filename)

        for (identifier, namespace) in ApplePages.xmlnamespaces.items():
            ElementTree.register_namespace(identifier, namespace)

        with zipfile.ZipFile(filename, 'r') as pageszip:
            self.document = ElementTree.XML(pageszip.read(ApplePages.primarydocument))

        self.__FixInsertionPoint()
        self.__AddStyles()

    def __AddStyles(self):
        """The XML document needs styles present in the header in order to be able to 
        handle e.g. superscripts."""

        # Superscript Example:
        # <sf:characterstyle sf:parent-ident="character-style-null" sfa:ID="SFWPCharacterStyle-10">
        #   <sf:property-map>
        #       <sf:superscript>
        #           <sf:number sfa:number="1" sfa:type="i"/>
        #       </sf:superscript>
        #   </sf:property-map>
        # </sf:characterstyle>

        styleparent = self.document.find(self.ns('.//sf:anon-styles'))

        style = ElementTree.SubElement(parent=styleparent, tag=self.ns('sf:characterstyle'), attrib={ 
            self.ns('sf:parent-ident'): u"character-style-null",
            self.ns('sfa:ID'): self.styles['superscript'], 
        })
        pmap = ElementTree.SubElement(parent=style, tag=self.ns('sf:property-map'))
        superscript = ElementTree.SubElement(parent=pmap, tag=self.ns('sf:superscript'))
        number = ElementTree.SubElement(parent=superscript, tag=self.ns('sf:number'), attrib={
            self.ns('sfa:number'): u"1",
            self.ns('sfa:type'): u"i",
        })

    def __FixInsertionPoint(self):
        """The insertion point breaks apart text nodes and can prevent correct parsing by
        breaking apart the markup."""

        insertionpointparent = self.document.find(self.ns('.//sf:insertion-point/..'))
        if insertionpointparent is not None:
            insertionpoint = self.document.find(self.ns('.//sf:insertion-point'))

            children = insertionpointparent.getchildren()
            previous = children[0]

            for child in children:
                if previous.tail is None:
                    previous.tail = u""

                if child is insertionpoint:
                    previous.tail += maybestr(insertionpoint.text) + maybestr(insertionpoint.tail)
                    insertionpointparent.remove(insertionpoint)
                    break
                else:
                    previous = child            

    def ProcessCitations(self, citestyle=u'square-brace', bibstyle=u'digit-dot', orderby=u'citation-first'):
        """Internally construct a version of the document that has the citations properly created."""
        assert citestyle in self.supportedcitestyles, citestyle + " is not a supported citation style."
        assert bibstyle in self.supportedbibstyles, bibstyle + " is not a supported bibliography style."
        assert orderby in (u'citation-first', u'reference-first'), orderby + " is not a supported ordering."

        bibliography = Bibliography(orderby=orderby)
        citationnodes = []
        bibnodes = []

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
                bibliography.AddReference(bibitemmatch.group(1), node)

        if bibliography.IsConsistent() == False:
            raise MissingReferenceError(
                message=u"Encountered citations that do not have a corresponding reference.", 
                badcites=set(bibliography.citations).difference(set(bibliography.references.keys()))
            )
        
        self.__ReplaceCitationMarkers(style=citestyle, citationnodes=citationnodes, bibliography=bibliography)
        self.__ReplaceBibitemMarkers(style=bibstyle, bibnodes=bibnodes, bibliography=bibliography)

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
                label = unicode(label)
                replacementtext = self.RenderCitation(style=style, index=bibliography.Index(label))
                replacementpattern = ur'\\cite\{' + label + ur'\}'
                text = re.sub(replacementpattern, replacementtext, text, flags=re.UNICODE)

            return text

        for node in citationnodes:
            xml = ElementTree.tostring(node, encoding="utf-8", method="xml")

            if type(xml) is str:
                xml = unicode(xml, encoding='utf8')

            xml = replacecitation(xml).encode('ascii', 'xmlcharrefreplace')
            modifiednode = ElementTree.XML(xml)
            copyelement(fromnode=modifiednode, tonode=node)

    def RenderCitation(self, style, index):
        """Produces the final XML that represents the citation."""

        assert style in ApplePages.supportedcitestyles

        if style == u'square-brace':
            return u"[{0}]".format(index)
        elif style == u'superscript':
            return u"<sf:span sf:style=\"{0}\">{1}</sf:span>".format(self.styles['superscript'], index)
        elif style == u'parens':
            return u"({0})".format(index)

        return unicode(index)    

    def __ReplaceBibitemMarkers(self, style, bibnodes, bibliography):
        """Helper method for ProcessCitations. Changes document to replace
        bibitems with the correct number and reference rendered in a particular style.

        Requires:
            style: A supported rendering style
            bibnodes: A list of Element objects representing a node containing a reference.
            bibliography: A completed (consistent) Bibliography object
        
        Returns: Void
        """

        renderednodes = []
        for sequencenumber in range(1, bibliography.Count() + 1):
            renderednodes.append(self.RenderReference(style=style, index=sequencenumber, bibliography=bibliography))

        for domnode, rendernode in zip(bibnodes, renderednodes):
            copyelement(fromnode=rendernode, tonode=domnode)

    def RenderReference(self, style, index, bibliography):
        """Produces the final XML that represents the reference."""
        assert style in self.supportedbibstyles

        (label, index, referencenode) = bibliography.GetReferenceByIndex(index)
        rendernode = ElementTree.Element(self.ns("sf:p"))
        copyelement(fromnode=referencenode, tonode=rendernode)

        if style == u'digit-dot':
            replacementtext = u"{0}. "
        elif style == u'square-brace':
            replacementtext = u"[{0}] "

        replacementtext = replacementtext.format(index)
        replacementpattern = ur'\\bibitem\{' + label + ur'\}'

        def replacebib(text):
            if text == None:
                return None

            for label in re.findall(replacementpattern, text):
                label = unicode(label)
                text = re.sub(replacementpattern, replacementtext, text)

            return text

        traversetransform(node=rendernode, transformingfunc=replacebib)

        return rendernode

    def Materialize(self, outputfilename):
        """Serializes the internal representation into a new pages file."""
        if self.filename != outputfilename:
            shutil.copy2(self.filename, outputfilename)

        with zipfile.ZipFile(outputfilename, 'a') as pageszip:
            finalxml = ElementTree.tostring(element=self.document, encoding='utf-8', method='xml')
            pageszip.writestr(ApplePages.primarydocument, finalxml) 

    def ns(self, string):
        """Provides transformation of namespaced tags into something ElementTree can understand."""

        # ElementTree represents namespaces like so: sf:p -> {http://developer.apple.com/namespaces/sf}p
        for namespace in ApplePages.xmlnamespaces:
            pattern = unicode(namespace + ":")
            string = re.sub(pattern, u"{{{0}}}".format(ApplePages.xmlnamespaces[namespace]), string)

        return string



