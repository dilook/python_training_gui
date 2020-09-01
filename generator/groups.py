import getopt
import os
import random
import string
import sys

from comtypes.client import CreateObject
# noinspection PyUnresolvedReferences
from comtypes.gen.Excel import xlLocalSessionChanges


def random_string(prefix, maxlen):
    symbols = string.ascii_letters + string.digits + string.punctuation + " " * 10
    return prefix + "".join([random.choice(symbols) for i in range(random.randrange(maxlen))])


try:
    opts, args = getopt.getopt(sys.argv[1:], "n:f:", ["number of groups", "file"])
except getopt.GetoptError as err:
    print(err)
    print("\tUse -n to set number of generating groups\n"
          "\tUse -f to set path to destination file")
    sys.exit(2)

n = 5
f = "data/groups.xlsx"

for o, a in opts:
    if o == "-n":
        n = int(a)
    elif o == "-f":
        f = a

project_dir = os.path.dirname(os.path.dirname(os.path.realpath(__file__)))
file = os.path.join(project_dir, f)

xl = CreateObject("Excel.Application")
try:
    xl.Visible = 0
    xl.DisplayAlerts = False
    wb = xl.Workbooks.Add()
    for i in range(n):
        xl.Range[f"A{i + 1}"].Value[()] = random_string("group", 10)
    wb.SaveAs(file, ConflictResolution=xlLocalSessionChanges)
finally:
    xl.Quit()
