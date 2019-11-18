def flat_dict(nested_dict: dict, data: list, previous_key: str = '', row_id: str = ''):
    for key, value in nested_dict.items():
        if isinstance(value, dict):
            if previous_key == '':
                flat_dict(value, data, key, row_id)
            else:
                flat_dict(value, data, '{}__{}'.format(previous_key, key), row_id)
        elif isinstance(value, list):
            if previous_key == '':
                flat_list(value, data, key, row_id)
            else:
                flat_list(value, data, '{}__{}'.format(previous_key, key), row_id)
        else:
            if previous_key == '':
                if row_id == '':
                    row_id = '1'
                data.append('{}::{} : {}'.format(row_id, key, value))
            else:
                if row_id == '':
                    row_id = '1'
                key = '{}__{}'.format(previous_key, key)
                data.append('{}::{} : {}'.format(row_id, key, value))


def flat_list(nested_list: dict, data: list, list_key: str = '', row_id: str = ''):
    temp_list = []
    for i in range(len(nested_list)):
        if isinstance(nested_list[i], list):
            if row_id == '':
                flat_list(nested_list[i], data, list_key, row_id='{}'.format(i + 1))
            else:
                flat_list(nested_list[i], data, list_key, row_id='{}||{}'.format(row_id, i + 1))
        elif isinstance(nested_list[i], dict):
            if row_id == '':
                flat_dict(nested_list[i], data, list_key, '{}'.format(i + 1))
            else:
                flat_dict(nested_list[i], data, list_key, row_id='{}||{}'.format(row_id, i + 1))
        else:
            temp_list.append(str(nested_list[i]))
    if len(temp_list) > 0:
        if row_id == '':
            row_id = '1'
        data.append('{}::{} : {}'.format(row_id, list_key, '|||'.join(temp_list)))


def json_to_excel(json_file_object, destination_file_location: str, main_sheet_name: str = 'Main'):
    import pandas as pd
    data = list()
    if isinstance(json_file_object, list):
        flat_list(json_file_object, data)
    else:
        flat_dict(json_file_object, data)

    relationship_hierarchy = {0: [main_sheet_name]}

    for cell in data:
        level = cell.count('||')
        if level > 0:
            sheet_name = '__'.join(cell.split('__')[:level])
            sheet_name = sheet_name[sheet_name.find('::') + 2:]
            relationship_hierarchy.setdefault(level, []).append(sheet_name)

    for level, sheet_name_list in relationship_hierarchy.items():
        relationship_hierarchy[level] = list(set(sheet_name_list))

    flat_data = {}

    for level, sheet_name_list in relationship_hierarchy.items():
        for cell in data:
            if cell.count('||') == level:
                row_number = cell[:cell.find('::')]
                cell_header = cell[cell.find('::') + 2:].split(' : ')[0]
                cell_value = cell[cell.find('::') + 2:].split(' : ')[1]
                flat_data.setdefault(row_number, []).append({cell_header: cell_value})

    df_data = {}

    for hierarchy_level, sheet_name_list in relationship_hierarchy.items():
        if hierarchy_level == 0:
            sheet_name = relationship_hierarchy[hierarchy_level][0]
            for row_id, row in flat_data.items():
                sheet_row = {}
                if row_id.count('||') == hierarchy_level:
                    for column in row:
                        for header, value in column.items():
                            sheet_row.__setitem__(header, value)
                            sheet_row.__setitem__('{}_Row_ID'.format(main_sheet_name), row_id)
                    df_data.setdefault(sheet_name, []).append(sheet_row)
        else:
            for sheet_name in relationship_hierarchy[hierarchy_level]:
                for row_id, row in flat_data.items():
                    sheet_row = {}
                    if row_id.count('||') == hierarchy_level:
                        for column in row:
                            for header, value in column.items():
                                if '__'.join(header.split('__')[:hierarchy_level]) == sheet_name:
                                    sheet_row.__setitem__(header, value)
                                    sheet_row.__setitem__('{}_Row_ID'.format(main_sheet_name), row_id.split('||')[0])
                                    for i in range(len(sheet_name.split('__'))):
                                        sheet_row.__setitem__('{}_Row_ID'.format(sheet_name.split('__')[i]),
                                                              row_id.split('||')[i+1])
                        if len(sheet_row) != 0:
                            df_data.setdefault(sheet_name, []).append(sheet_row)
    # https://github.com/PyCQA/pylint/issues/3060 pylint: disable=abstract-class-instantiated
    with pd.ExcelWriter(destination_file_location) as writer:
        for sheet_name, data in df_data.items():
            df = pd.DataFrame(data)
            df.to_excel(writer, sheet_name=sheet_name, index=False)
