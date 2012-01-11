#!/usr/bin/python

# Mendeley Desktop API

# This provides a wrapper for OpenOffice basic to use the
# HTTP/JSON Mendeley Desktop Word Processor API.
# It is responsible for building the python dictionaries which represent
# the JSON requests and providing accessor functions for the server responses.

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
    print "-- not running in OpenOffice environment --"
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

if not testMode:
    g_ImplementationHelper = unohelper.ImplementationHelper()

class MendeleyDesktopAPI(unohelper.Base, XJob):
    def __init__(self, ctx):
        self.closed = 0
        self.ctx = ctx # component context
        self._client = MendeleyHttpClient()
        
        self._formattedCitationsResponse = MendeleyHttpClient.ResponseBody()

        self.citationClusters = []
        self.citationStyleUrl = ""
        self.formattedBibliography = []

        # used by citation_update_interactive
        self.formattedText = ""
    
    def resetCitations(self, unused):
        self.citationClusters = []

    def _citationClusterFromFieldCode(self, fieldCode):
        # remove ADDIN and CSL_CITATION from start
        pattern = re.compile("CSL_CITATION[ ]*({.*$)")
        match = pattern.search(fieldCode)

        if match == None:
            result = {"fieldCode" : fieldCode.decode('string_escape')}
        else:
            bareJson = match.group(1)
            citationCluster = json.loads(bareJson)
            result = {"citationCluster" : citationCluster}
        return result

    def addCitationCluster(self, fieldCode):
        self.citationClusters.append(self._citationClusterFromFieldCode(fieldCode))

    def addFormattedCitation(self, formattedCitation):
        self.citationClusters[len(self.citationClusters)-1]["formattedText"] = formattedCitation
    
    def setCitationStyle(self, citationStyleUrl):
        self.citationStyleUrl = citationStyleUrl

    def getCitationStyleId(self, unused):
        return self.citationStyleUrl

    def formatCitationsAndBibliography(self, unused):
        self._formattedCitationsResponse = \
            self._client.formattedCitationsAndBibliography_Interactive(
            self.citationStyleUrl, self.citationClusters).body

        return json.dumps(self._formattedCitationsResponse.__dict__)

    def getCitationCluster(self, index):
        return "ADDIN CSL_CITATION " + json.dumps(self._formattedCitationsResponse.citationClusters[int(index)]["citationCluster"])
#        return "citation cluster"

    def getFormattedCitation(self, index):
        return self._formattedCitationsResponse.citationClusters[int(index)]["formattedText"]
#        return json.dumps(self._formattedCitationsResponse.__dict__)

    def getFormattedBibliography(self, unused):
        # a single string is interpreted as a file name
        if (type(self._formattedCitationsResponse.bibliography) == type(u"unicode string")
                or type(self._formattedCitationsResponse.bibliography) == type("string")):
            return self._formattedCitationsResponse.bibliography;
        else:
            return "<br/>".join(self._formattedCitationsResponse.bibliography)
        
    def getUserAccount(self, unused):
        return self._client.userAccount().body.account

    def citationStyle_choose_interactive(self, styleId):
        return self._client.citationStyle_choose_interactive(
            {"citationStyleUrl": styleId}).body.citationStyleUrl

    def citation_choose_interactive(self):
        response = self._client.citation_choose_interactive()
        try:
            assert(response.status==200)
            fieldCode = self._fieldCodeFromCitationCluster(response.body.citationCluster)
        except:
            raise MendeleyHttpClient.UnexpectedResponse(response)
        return fieldCode
    
    def citation_edit_interactive(self, fieldCode):
        response = self._client.citation_edit_interactive(
            self._citationClusterFromFieldCode(fieldCode))
        try:
            assert(response.status==200)
            fieldCode = self._fieldCodeFromCitationCluster(response.body.citationCluster)
        except:
            raise MendeleyHttpClient.UnexpectedResponse(response)
        return fieldCode
    
    def setDisplayedText(self, displayedText):
        self.formattedText = displayedText

    def citation_update_interactive(self, fieldCode):
        citationCluster = self._citationClusterFromFieldCode(fieldCode)
        citationCluster["formattedText"] = self.formattedText

        response = self._client.citation_update_interactive(citationCluster)
        try:
            assert(response.status==200)
            fieldCode = self._fieldCodeFromCitationCluster(response.body.citationCluster)
        except:
            raise MendeleyHttpClient.UnexpectedResponse(response)
        return fieldCode

    def getFieldCodeFromUuid(self, documentUuid):
        response = self._client.testMethods_citationCluster_getFromUuid(
            {"documentUuid": documentUuid})
        try:
            assert(response.status==200)
            fieldCode = self._fieldCodeFromCitationCluster(response.body.citationCluster)
        except:
            raise MendeleyHttpClient.UnexpectedResponse(response)
        return fieldCode

    def _fieldCodeFromCitationCluster(self, citationCluster):
        return "ADDIN CSL_CITATION " + json.dumps(citationCluster)

    # for testing
    def setNumberTest(self, number):
        self.number = number.decode('string_escape')
        return "from hello " + str(number)

    # for testing
    def getNumberTest(self, unused):
        return "number = " + str(self.number)

    #TODO: refactor to allow multiple numbers of arguments
    def execute(self, args):
        functionName = str(args[0].Value)
        functionArg  = str(args[1].Value)
        if hasattr(self, functionName):
            statement = 'self.' + functionName + '("' + functionArg.encode('string_escape').replace('"', '\\"') + '")'
            return eval(statement)
        else:
            raise Exception("ERROR: Function " + functionName + " doesn't exist")

if not testMode:
    g_ImplementationHelper.addImplementation(MendeleyDesktopAPI, "org.openoffice.pyuno.MendeleyDesktopAPI", ("com.sun.star.task.MendeleyDesktopAPI",),)
