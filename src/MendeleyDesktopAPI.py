#!/usr/bin/python

# Mendeley Desktop API

# To run tests, open MendeleyDesktop and run this script
# (it knows that it's in test mode because "import unohelper"
#  will fail when run outside of the OpenOffice.org environment)

import re

try:
    import unohelper
    from com.sun.star.task import XJob
    testMode = False
except:
    testMode = True
    print "-- test mode --"
    from MendeleyHttpClient import *

    class unohelper():
        def __init__(self, ctx):
            pass

        class Base():
            def __init__(self, ctx):
                pass

    class XJob():
        def __init__(self, ctx):
            pass

#import imp
#import os
#import sys

#encoding = sys.getfilesystemencoding()
#sourceDir = os.path.dirname(unicode(__file__, encoding))
#sourceDir = os.path.dirname(unicode(sys.executable, encoding))
#MendeleyHttpClient = imp.load_source('MendeleyHttpClient', 'f:/code/openoffice-plugin/MendeleyEmptyExtension.oxt/Scripts/MendeleyHttpClient.py')
#MendeleyHttpClient = imp.load_source('MendeleyHttpClient', sourceDir + '/MendeleyHttpClient.py')

if not testMode:
    g_ImplementationHelper = unohelper.ImplementationHelper()

class MendeleyDesktopAPI(unohelper.Base, XJob):
    def __init__(self, ctx):
        self.closed = 0
        self.ctx = ctx # component context
        self._client = MendeleyHttpClient()
        
        self._formattedCitationsResponse = MendeleyHttpClient.FormattedCitationsAndBibliographyResponse()

        self.citationClusters = []
        self.citationStyleUrl = ""
        self.formattedBibliography = []
    
    def resetCitations(self, unused):
        self.citationClusters = []

    def addCitationCluster(self, fieldCode):
        # remove ADDIN and CSL_CITATION from start
        pattern = re.compile("CSL_CITATION[ ]*({.*$)")
        match = pattern.search(fieldCode)

        if match == None:
            self.citationClusters.append({"fieldCode" : fieldCode.decode('string_escape')})
        else:
            bareJson = match.group(1)
            citationCluster = json.loads(bareJson)
            self.citationClusters.append({"citationCluster" : citationCluster})

    def addFormattedCitation(self, formattedCitation):
        self.citationClusters[len(self.citationClusters)-1]["formattedText"] = formattedCitation
    
    def setCitationStyle(self, citationStyleUrl):
        self.citationStyleUrl = citationStyleUrl

    def getCitationStyleId(self, unused):
        return self.citationStyleUrl

    def formatCitationsAndBibliography(self, unused):
        self._formattedCitationsResponse = self._client.formattedCitationsAndBibliography_Interactive(
                 self.citationStyleUrl, self.citationClusters)

        return json.dumps(self._formattedCitationsResponse.__dict__)

    def getCitationCluster(self, index):
        return json.dumps(self._formattedCitationsResponse.citationClusters[int(index)]["citationCluster"])
#        return "citation cluster"

    def getFormattedCitation(self, index):
        return str(self._formattedCitationsResponse.citationClusters[int(index)]["formattedText"])
#        return json.dumps(self._formattedCitationsResponse.__dict__)

    def getFormattedBibliography(self):
        return self.formattedBibliography.join("<br/>")

    # for testing
    def setNumberTest(self, number):
        self.number = number.decode('string_escape')
        return "from hello " + str(number)

    # for testing
    def getNumberTest(self, unused):
        return "number = " + str(self.number)

    def execute(self, args):
        functionName = str(args[0].Value)
        functionArg  = str(args[1].Value)
        if hasattr(self, functionName):
            statement = 'self.' + functionName + '("' + functionArg.encode('string_escape').replace('"', '\\"') + '")'
            return str(eval(statement))
        else:
            raise Exception("ERROR: Function " + functionName + " doesn't exist")

if not testMode:
    g_ImplementationHelper.addImplementation(MendeleyDesktopAPI, "org.openoffice.pyuno.MendeleyDesktopAPI", ("com.sun.star.task.MendeleyDesktopAPI",),)
else:
    print "--running unit test--"
    
    print "set up http client"
    api = MendeleyDesktopAPI("component context")

    print "Format some citations and a bibliography"
    api.resetCitations("")
    api.setCitationStyle("http://www.zotero.org/styles/apa")
    api.addCitationCluster("ADDIN any old text can go here CSL_CITATION { \"citationItems\" : [ { \"id\" : \"ITEM-1\", \"itemData\" : { \"author\" : [ { \"family\" : \"Smith\", \"given\" : \"John\" }, { \"family\" : \"Jr\", \"given\" : \"John Smith\" } ], \"id\" : \"ITEM-1\", \"issued\" : { \"date-parts\" : [ [ \"2001\" ] ] }, \"title\" : \"Title01\", \"type\" : \"article\" }, \"uris\" : [ \"http://local/documents/?uuid=55ff8735-3f3c-4c9f-87c3-8db322ba3f74\" ] }, { \"id\" : \"ITEM-2\", \"itemData\" : { \"author\" : [ { \"family\" : \"Evans\", \"given\" : \"Gareth\" }, { \"family\" : \"Jr\", \"given\" : \"Gareth Evans\" } ], \"id\" : \"ITEM-2\", \"issued\" : { \"date-parts\" : [ [ \"2002\" ] ] }, \"title\" : \"Title02\", \"type\" : \"article\" }, \"uris\" : [ \"http://local/documents/?uuid=15d6d1e4-a9ff-4258-88b6-a6d6d6bdc0ed\" ] } ], \"mendeley\" : { \"previouslyFormattedCitation\" : \"(Evans & Jr, 2002; Smith & Jr, 2001)\" }, \"properties\" : { \"noteIndex\" : 0 }, \"schema\" : \"https://github.com/citation-style-language/schema/raw/master/csl-citation.json\" }")
    api.addFormattedCitation("(Evans & Jr, 2002; Smith & Jr, 2001)")
    api.addCitationCluster("Mendeley Citation{15d6d1e4-a9ff-4258-88b6-a6d6d6bdc0ed}")
    api.addFormattedCitation("test")
    print "formatted citation and bib: " + api.formatCitationsAndBibliography("")

    print "Returned citation JSON: " + api.getCitationCluster(0)
    print "Returned formatted citation: " + api.getFormattedCitation(0)
    print ""
    print "Returned citation JSON: " + api.getCitationCluster(1)
    print "Returned formatted citation: " + api.getFormattedCitation(1)
    

