import pytest

from fixture.application import Application


@pytest.fixture(scope="session")
def app(request):
    fixture = Application('C:\\Program Files (x86)\\GAS Softwares\\Free Address Book\\AddressBook.exe')
    request.addfinalizer(fixture.destroy)
    return fixture
