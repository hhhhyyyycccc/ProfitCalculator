import unittest
import my_package

class TestPackage(unittest.TestCase):
    def test_version_attribute(self):
        # ensure package can be imported and has __version__
        self.assertTrue(hasattr(my_package, "__version__"))

if __name__ == "__main__":
    unittest.main()
