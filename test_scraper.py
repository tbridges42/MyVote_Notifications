import unittest
import scraper

class TestScraper(unittest.TestCase):

    def test_get_creds(self):
        creds = scraper.get_creds()
        self.assertEquals("bridgt", creds['username'])

    def test_date(self):
        print(scraper.get_yesterday())


if __name__ == '__main__':
    unittest.main()