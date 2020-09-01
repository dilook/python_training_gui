

def test_add_group(app, excel_groups):
    data_groups = excel_groups
    old_list = app.group.get_group_list()
    app.group.add_new_group(data_groups)
    new_list = app.group.get_group_list()
    old_list.append(data_groups)
    assert sorted(old_list) == sorted(new_list)
