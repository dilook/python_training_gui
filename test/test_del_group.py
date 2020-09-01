import random


def test_add_group(app):
    if app.group.count() == 1:
        app.group.add_new_group("first_group")
    old_list = app.group.get_group_list()
    index = random.randrange(len(old_list))
    app.group.remove_group_by_index(index)
    assert len(old_list) - 1 == app.group.count()
    new_list = app.group.get_group_list()
    old_list[index: index + 1] = []
    assert sorted(old_list) == sorted(new_list)
