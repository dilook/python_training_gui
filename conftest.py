import os

import pytest
from comtypes.client import CreateObject

from fixture.application import Application


@pytest.fixture(scope="session")
def app(request):
    fixture = Application('C:\\Program Files (x86)\\GAS Softwares\\Free Address Book\\AddressBook.exe')
    request.addfinalizer(fixture.destroy)
    return fixture


def load_from_excel(file_name):
    file = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data", f"{file_name}.xlsx")
    return read_excel(file)


def pytest_generate_tests(metafunc):
    for fixture in metafunc.fixturenames:
        if fixture.startswith("excel_"):
            test_data = load_from_excel(fixture[6:])
            metafunc.parametrize(fixture, test_data)


def read_excel(filename):
    result = []
    xl = CreateObject("Excel.Application")
    try:
        xl.Visible = 0
        wb = xl.Workbooks.Open(filename, ReadOnly=True)
        sheet = wb.Sheets[1]
        for row in sheet.UsedRange.Rows:
            result.append(row.Value())
    finally:
        xl.Quit()
    return result
