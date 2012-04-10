#!/usr/bin/python

# Mendeley Desktop API

# This provides a wrapper for OpenOffice basic to use the
# HTTP/JSON Mendeley Desktop Word Processor API.
# It is responsible for building the python dictionaries which represent
# the JSON requests and providing accessor functions for the server responses.

# To run tests, open MendeleyDesktop and run this script
# (it knows that it's in test mode because "import unohelper"
#  will fail when run outside of the OpenOffice.org environment)

# simplejson is json 
try: import simplejson as json
except ImportError: import json

import os
import re

# if MENDELEY_UNIT_TEST environment variable exists:
# it doesn't try to use the unohelper package. Mendeley tests sets this
# variable when needed.
if os.environ.has_key('MENDELEY_UNIT_TEST'):
    # either unohelper or XJob modules are not available
    # these are only required when running within OpenOffice"
    from MendeleyHttpClient import MendeleyHttpClient
    testMode = True
    
    class unohelper():
        def __init__(self, ctx):
            pass

        class Base():
            def __init__(self, ctx):
                pass

    class XJob():
        def __init__(self, ctx):
            pass
    
else:
    import unohelper
    from com.sun.star.task import XJob
    testMode = False

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

        self._previousResultLength = 0

    def resetCitations(self):
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

    def getCitationStyleId(self):
        return self.citationStyleUrl

    def formatCitationsAndBibliography(self):
        self._formattedCitationsResponse = \
            self._client.formattedCitationsAndBibliography_Interactive(
            self.citationStyleUrl, self.citationClusters).body

        return json.dumps(self._formattedCitationsResponse.__dict__)

    def getCitationCluster(self, index):
        return "ADDIN CSL_CITATION " + json.dumps(self._formattedCitationsResponse.citationClusters[int(index)]["citationCluster"])
#        return "citation cluster"

    def getFormattedCitation(self, index):
        return self._formattedCitationsResponse.citationClusters[int(index)]["formattedText"]

    def getFormattedBibliography(self):
        # a single string is interpreted as a file name
        if (type(self._formattedCitationsResponse.bibliography) == type(u"unicode string")
                or type(self._formattedCitationsResponse.bibliography) == type("string")):
            return self._formattedCitationsResponse.bibliography;
        else:
            return "<br/>".join(self._formattedCitationsResponse.bibliography)
        
    def getUserAccount(self):
        response = self._client.userAccount()
        
        if (response.status != 200):
            raise MendeleyHttpClient.UnexpectedResponse(response)

        return response.body.account

    def citationStyle_choose_interactive(self, styleId):
        return self._client.citationStyle_choose_interactive(
            {"currentStyleUrl": styleId}).body.citationStyleUrl

    def citation_choose_interactive(self, hintText):
        response = self._client.citation_choose_interactive(
            {"citationEditorHint": hintText})
        try:
            assert(response.status == 200)
            fieldCode = self._fieldCodeFromCitationCluster(response.body.citationCluster)
        except:
            raise MendeleyHttpClient.UnexpectedResponse(response)

        return fieldCode
    
    def citation_edit_interactive(self, fieldCode, hintText):
        citationCluster = self._citationClusterFromFieldCode(fieldCode)
        citationCluster["citationEditorHint"] = hintText
        response = self._client.citation_edit_interactive(citationCluster)
        try:
            assert(response.status == 200)
            fieldCode = self._fieldCodeFromCitationCluster(response.body.citationCluster)
        except:
            raise MendeleyHttpClient.UnexpectedResponse(response)
        return fieldCode
    
    def setDisplayedText(self, displayedText):
        self.formattedText = displayedText

    def citation_update_interactive(self, fieldCode, formattedText):
        citationCluster = self._citationClusterFromFieldCode(fieldCode)
        citationCluster["formattedText"] = formattedText

        response = self._client.citation_update_interactive(citationCluster)
        try:
            assert(response.status == 200)
            fieldCode = self._fieldCodeFromCitationCluster(response.body.citationCluster)
        except:
            raise MendeleyHttpClient.UnexpectedResponse(response)
        return fieldCode

    def getFieldCodeFromUuid(self, documentUuid):
        response = self._client.testMethods_citationCluster_getFromUuid(
            {"documentUuid": documentUuid})
        try:
            assert(response.status == 200)
            fieldCode = self._fieldCodeFromCitationCluster(response.body.citationCluster)
        except:
            raise MendeleyHttpClient.UnexpectedResponse(response)
        return fieldCode

    def _fieldCodeFromCitationCluster(self, citationCluster):
        if ("citationItems" in citationCluster):
            if (len(citationCluster["citationItems"]) == 0):
                return ""

        return "ADDIN CSL_CITATION " + json.dumps(citationCluster, sort_keys=True)

    def citation_undoManualFormat(self, fieldCode):
        citationCluster = self._citationClusterFromFieldCode(fieldCode)
        response = self._client.citation_undoManualFormat(citationCluster)
        try:
            assert(response.status == 200)
            fieldCode = self._fieldCodeFromCitationCluster(response.body.citationCluster)
        except:
            raise MendeleyHttpClient.UnexpectedResponse(response)
        return fieldCode

    def citations_merge(self, *fieldCodes):
        clusters = []

        for fieldCode in fieldCodes:
            clusters.append(self._citationClusterFromFieldCode(fieldCode))

        response = self._client.citations_merge({"citationClusters": clusters})
        try:
            assert(response.status == 200)
            mergedFieldCode = \
                self._fieldCodeFromCitationCluster(response.body.citationCluster)
        except:
            raise MendeleyHttpClient.UnexpectedResponse(response)
        
        return mergedFieldCode

    def wordProcessor_set(self, wordProcessor, version):
        response = self._client.wordProcessor_set(
            {
                "wordProcessor": wordProcessor,
                "version": version
            })

        try:
            assert(response.status == 200)
        except:
            raise MendeleyHttpClient.UnexpectedResponse(response)

        return ""

    def mendeleyDesktopInfo(self):
        response = self._client.mendeleyDesktopInfo()
        try:
            assert(response.status == 200)
        except:
            raise MendeleyHttpClient.UnexpectedResponse(response)

	result = {"processId": response.body.processId}
	
        return result

    def isMendeleyDesktopRunningStr(self):
	try:
            response = self._client.mendeleyDesktopInfo()
            return str(response.status == 200)
	except:
	    return False

    # for testing
    def setNumberTest(self, number):
        self.number = number.decode('string_escape')
        return ""

    # for testing
    def getNumberTest(self):
        return str(self.number)

    # for testing
    def concatenateStringsTest(self, string1, string2):
        return str(string1) + str(string2)

    def previousSuccess(self):
        previousResponse = self._client.previousResponse
        
        return str(previousResponse.status == 200)

    def previousErrorMessage(self):
        previousResponse = self._client.previousResponse
        
        downloadInstructions = "Please download the latest version " + \
                "of Mendeley Desktop here: \n" + \
                "http://www.mendeley.com/download-mendeley-desktop"

        if (previousResponse.status == 406 or 
            previousResponse.status == 415):
            if (previousResponse.contentType.startswith(
                "application/vnd.mendeley.typeDeprecatedError")):
                # TODO: insert link to plugin download
                return "Deprecated type error. Please update this plugin to work with the " + \
                    "current version of Mendeley Desktop"
            else:
                return "Unknown type error. " + downloadInstructions

        if (previousResponse.status == 404):
            return "Page not found. " + downloadInstructions

        if (previousResponse.status != 200):
            return "Unknown error\n" + json.dumps(previousResponse.__dict__)

        return ""

    def previousResultLength(self):
        return self._previousResultLength

    def previousResponse(self):
        return json.dumps(self.previousResponse.__dict__)

    def execute(self, args):
        functionName = str(args[0].Value)
        statement = 'self.' + functionName + '('
        for arg in range(1, len(args)):
            statement += '"'
            statement += args[arg].Value.encode('unicode_escape').replace('"', '\\"')
            statement += '"'
            if arg < len(args) - 1:
                statement += ', '
        statement += ')'

        if hasattr(self, functionName):
            try:
                result = eval(statement)
                self._previousResultLength = len(unicode(result))
                return result
            except MendeleyHttpClient.UnexpectedResponse:
                return ""
        else:
            raise Exception("ERROR: Function " + functionName + " doesn't exist")

if not testMode:
    g_ImplementationHelper.addImplementation(MendeleyDesktopAPI,
        "org.openoffice.pyuno.MendeleyDesktopAPI",
        ("com.sun.star.task.MendeleyDesktopAPI",),)
