import unittest
import pandas as pd
from formatting import creatingDict

class Testing(unittest.TestCase):

    def setUp(self):
        self.empDict = {}
        self.demoData = pd.DataFrame({
            'EmpID': [1, 1, 1, 2],
            'EmpName': ['Kevin Hart', 'Kevin Hart', 'Kevin Hart', 'Dwayne Johnson'],
            'StartTime': ['07:00', '07:00', '07:00', '07:00'],
            'EndTime': ['18:00', None, '19:00', '20:00'],
            'ShiftTime': ['660', 'N/A', '670', '680']
        })
        creatingDict(self.demoData, self.empDict)

    def test_creatingDict(self):                
        self.assertEqual(len(self.empDict['Kevin Hart']), 3)
        self.assertEqual(len(self.empDict['Dwayne Johnson']), 1)

if __name__ == '__main__':
    unittest.main()