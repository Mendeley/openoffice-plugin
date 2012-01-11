
import MendeleyHttpClient
import json
import unittest

class TestMendeleyHttpClientInteractive(unittest.TestCase):

    def setUp(self):
        self.client = MendeleyHttpClient.MendeleyHttpClient()

    def test_citationStyle_choose_interactive(self):
        response = self.client.citationStyle_choose_interactive(
            {"citationStyleUrl": "http://www.zotero.org/styles/apa"}
            )

        print "chosen style = " + response.citationStyleUrl

    def test_citation_edit_interactive(self, citationCluster):
        response = self.client.citation_edit_interactive(

if __name__ == '__main__':
    unittest.main()

