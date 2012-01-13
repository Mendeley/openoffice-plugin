import MendeleyDesktopAPI 
import unittest

# API functions that can't be tested without manual user interaction
class TestMendeleyDesktopAPI_Interactive(unittest.TestCase):

    def setUp(self):
        self.api = MendeleyDesktopAPI.MendeleyDesktopAPI("component context (unused)")

    def test_citationStyle_choose_interactive(self):
        chosenStyle = self.api.citationStyle_choose_interactive("http://www.zotero.org/styles/apa")
        print "chosen style = " + chosenStyle

    def test_citation_choose_interactive(self):
        chosenFieldCode = self.api.citation_choose_interactive("hint text")
        print "chosen field code = " + chosenFieldCode
    
    def test_citation_edit_interactive(self):
        editedCitation = self.api.citation_edit_interactive(
            "Mendeley Citation{15d6d1e4-a9ff-4258-88b6-a6d6d6bdc0ed}", "hint text")
        print "edited citation = " + editedCitation

    def test_formatCitations_withManualEdit(self):
        self.api.resetCitations()
        self.api.setCitationStyle("http://www.zotero.org/styles/apa")
        self.api.addCitationCluster("ADDIN any old text can go here CSL_CITATION { \"citationItems\" : [ { \"id\" : \"ITEM-1\", \"itemData\" : { \"author\" : [ { \"family\" : \"Smith\", \"given\" : \"John\" }, { \"family\" : \"Jr\", \"given\" : \"John Smith\" } ], \"id\" : \"ITEM-1\", \"issued\" : { \"date-parts\" : [ [ \"2001\" ] ] }, \"title\" : \"Title01\", \"type\" : \"article\" }, \"uris\" : [ \"http://local/documents/?uuid=55ff8735-3f3c-4c9f-87c3-8db322ba3f74\" ] }, { \"id\" : \"ITEM-2\", \"itemData\" : { \"author\" : [ { \"family\" : \"Evans\", \"given\" : \"Gareth\" }, { \"family\" : \"Jr\", \"given\" : \"Gareth Evans\" } ], \"id\" : \"ITEM-2\", \"issued\" : { \"date-parts\" : [ [ \"2002\" ] ] }, \"title\" : \"Title02\", \"type\" : \"article\" }, \"uris\" : [ \"http://local/documents/?uuid=15d6d1e4-a9ff-4258-88b6-a6d6d6bdc0ed\" ] } ], \"mendeley\" : { \"previouslyFormattedCitation\" : \"(Evans & Jr, 2002; Smith & Jr, 2001)\" }, \"properties\" : { \"noteIndex\" : 0 }, \"schema\" : \"https://github.com/citation-style-language/schema/raw/master/csl-citation.json\" }")
        self.api.addFormattedCitation("(Evans & Jr, 2002; Smith & Jr, 2001)")
        self.api.addCitationCluster("Mendeley Citation{15d6d1e4-a9ff-4258-88b6-a6d6d6bdc0ed}")
        self.api.addFormattedCitation("test")
        print "formatted: " + self.api.formatCitationsAndBibliography()

if __name__ == '__main__':
    unittest.main()
