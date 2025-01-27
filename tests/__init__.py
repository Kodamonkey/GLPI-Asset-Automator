import unittest
from scripts.interface_v3 import *

class TestInterfaceV3(unittest.TestCase):
    def test_functionality_one(self):
        self.assertEqual(functionality_one(args), expected_result)

    def test_functionality_two(self):
        self.assertTrue(functionality_two(args))

    def test_edge_case(self):
        with self.assertRaises(ExpectedException):
            edge_case_functionality(args)

if __name__ == '__main__':
    unittest.main()