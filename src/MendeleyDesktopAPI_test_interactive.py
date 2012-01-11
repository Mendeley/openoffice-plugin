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
        chosenFieldCode = self.api.citation_choose_interactive()
        print "chosen field code = " + chosenFieldCode
    
    def test_citation_edit_interactive(self):
        editedCitation = self.api.citation_edit_interactive(
            "Mendeley Citation{15d6d1e4-a9ff-4258-88b6-a6d6d6bdc0ed}")
        print "edited citation = " + editedCitation

if __name__ == '__main__':
    unittest.main()
