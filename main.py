import PySimpleGUI as sg
import PLM_BOM_transpose


def PopupQuickView(subassembly):
    if not any(subassembly[0] in i for i in products):
        sg.Popup("No such subassembly loaded")
        return None

    subassembly_parts = sorted([i[1:] for i in parts if subassembly[0] in i[0]])
    window_q = sg.Window("QuickView",
                         [[sg.Table(values=subassembly_parts, headings=headings, max_col_width=700,
                                    auto_size_columns=True,
                                    display_row_numbers=False,
                                    cols_justification=['l', 'c', 'l'],
                                    def_col_width=50,
                                    right_click_selects=True,
                                    num_rows=20,
                                    key='-TABLE_QUICKVIEW-',
                                    selected_row_colors='red on yellow',
                                    enable_events=True,
                                    expand_x=True,
                                    expand_y=True,
                                    enable_click_events=True,  # Comment out to not enable header and other clicks
                                    tooltip=f'This is BOM for {subassembly[0]}',
                                    right_click_menu=['', ['Quick view']])], ],
                         finalize=True, font=('Helvetica', 16),
                         resizable=True, size=(600, 200))
    while True:
        event_q, values_q = window_q.read()
        print(event_q, values_q)
        if event_q == sg.WIN_CLOSED:
            break
        if event_q == 'Quick view':
            window_q.close()
            if not PopupQuickView(window_q['-TABLE_QUICKVIEW-'].get()[values_q['-TABLE_QUICKVIEW-'][0]]):
                PopupQuickView(subassembly)
    return True


def compare(source):
    parts_right, parts_left = [], []

    if len(values['-IN_RIGHT-']) > 0:
        if source == "input":
            picked_right = values['-IN_RIGHT-'].replace("'", "").replace(",", "").replace("(", "").replace(")", "")
        else:
            picked_right = values['-BOX_RIGHT-'][0]
        parts_right = sorted([i[1:] for i in parts if i[0] == picked_right])
        window['-TABLE_RIGHT-'].update(parts_right)

    else:
        window['-TABLE_RIGHT-'].update([])
        window.refresh()

    if len(values['-IN_LEFT-']) > 0:
        if source == "input":
            picked_left = values['-IN_LEFT-'].replace("'", "").replace(",", "").replace("(", "").replace(")", "")
        else:
            picked_left = values['-BOX_LEFT-'][0]
        parts_left = sorted([i[1:] for i in parts if i[0] == picked_left])
        window['-TABLE_LEFT-'].update(parts_left)

    else:
        window['-TABLE_LEFT-'].update([])

    if len(window['-TABLE_RIGHT-'].get()) > 1 and len(window['-TABLE_LEFT-'].get()) > 1:
        for count, part in enumerate(parts_left):
            index = -1
            for i, obj in enumerate(parts_right):
                if obj[0] == part[0]:
                    if obj[1] == part[1]:
                        index = i
                    break
            if index == -1:
                window['-TABLE_LEFT-'].update(row_colors=((count, 'red'),))

        for count, part in enumerate(parts_right):
            index = -1
            for i, obj in enumerate(parts_left):
                if obj[0] == part[0]:
                    if obj[1] == part[1]:
                        index = i
                    break
            if index == -1:
                window['-TABLE_RIGHT-'].update(row_colors=((count, 'red'),))
    else:
        window['-TABLE_LEFT-'].update(row_colors=((0, ''),))
        window['-TABLE_RIGHT-'].update(row_colors=((0, ''),))


headings = ["Part", "Qty", "Description"]

col1 = sg.Column([[sg.Table(values=[[]], headings=headings, max_col_width=700,
                            auto_size_columns=True,
                            display_row_numbers=False,
                            # justification='center',
                            cols_justification=['l', 'c', 'l'],
                            def_col_width=50,
                            right_click_selects=True,
                            num_rows=20,
                            key='-TABLE_LEFT-',
                            selected_row_colors='red on yellow',
                            enable_events=True,
                            expand_x=True,
                            expand_y=True,
                            enable_click_events=True,  # Comment out to not enable header and other clicks
                            tooltip='This is BOM1',
                            right_click_menu=['', ['Quick view', 'Expand for this', 'Expand for all']])]],
                 pad=(0, 0), expand_x=True)

col2 = sg.Column([[sg.Table(values=[[]], headings=headings, max_col_width=50,
                            auto_size_columns=True,
                            display_row_numbers=False,
                            # justification='right',
                            cols_justification=['l', 'c', 'l'],
                            right_click_selects=True,
                            num_rows=20,
                            key='-TABLE_RIGHT-',
                            selected_row_colors='red on yellow',
                            enable_events=True,
                            expand_x=True,
                            expand_y=True,
                            enable_click_events=True,  # Comment out to not enable header and other clicks
                            tooltip='This is BOM2',
                            right_click_menu=['', ['Quick view', 'Expand for this', 'Expand for all']])]],
                 pad=(0, 0), expand_x=True)

col3 = sg.Column([[sg.Frame('Actions:',
                            [[sg.Column([[sg.Button('Load Data', key='-FILES-'),
                                          sg.Button('Clear All', key='-CLEAR_FILES-'), sg.Button('Clear Tables'), ]
                                         ],
                                        pad=(0, 0))]])]], pad=(0, 0))

col4 = sg.Column([
    [sg.Text('Kod motoru:')],
    [sg.Input(size=(20, 1), enable_events=True, key='-IN_LEFT-')],
    [sg.pin(sg.Col([[sg.Listbox(values=[], size=(20, 4), enable_events=True, key='-BOX_LEFT-',
                                select_mode=sg.LISTBOX_SELECT_MODE_SINGLE, no_scrollbar=True)]],
                   key='-BOX-CONTAINER_LEFT-', pad=(0, 0), visible=False))]])

col5 = sg.Column([
    [sg.Text('Kod motoru:')],
    [sg.Input(size=(20, 1), enable_events=True, key='-IN_RIGHT-')],
    [sg.pin(sg.Col([[sg.Listbox(values=[], size=(20, 4), enable_events=True, key='-BOX_RIGHT-',
                                select_mode=sg.LISTBOX_SELECT_MODE_SINGLE, no_scrollbar=True)]],
                   key='-BOX-CONTAINER_RIGHT-', pad=(0, 0), visible=False))]])

layout = [[sg.Push(), col3, sg.Push()],
          [col4, sg.Push(), col5],
          [col1, col2]]

window = sg.Window('Columns and Frames', layout, return_keyboard_events=True, finalize=True, font=('Helvetica', 16),
                   resizable=True, size=(1200, 800))
list_element_left: sg.Listbox = window.Element(
    '-BOX_LEFT-')  # store listbox element for easier access and to get to docstrings
list_element_right: sg.Listbox = window.Element(
    '-BOX_RIGHT-')  # store listbox element for easier access and to get to docstrings
prediction_list_left, input_text_left, sel_item_left = [], "", 0
prediction_list_right, input_text_right, sel_item_right = [], "", 0
products, parts = [], []

table_clicked_right, table_clicked_left = False, False

while True:
    event, values = window.read()
    print(event, values)
    if event == sg.WIN_CLOSED:
        break

    elif event == '-FILES-':
        try:
            files = sg.PopupGetFile('Select folder', no_window=True, multiple_files=True,
                                    file_types=(("Excel files", [".xls", ".xlsx"]),))
            products_t, parts_t = PLM_BOM_transpose.load_files(files)
            products += products_t
            parts += parts_t
            products = list(set(products))
            parts = list(set(parts))
        except Exception as e:
            print(e)
            sg.Popup('Wrong file', keep_on_top=True)

    elif event == "Quick view":
        if not table_clicked_right:
            try:
                PopupQuickView(window['-TABLE_LEFT-'].get()[table_clicked_left])
            except:
                continue
        elif not table_clicked_left:
            try:
                PopupQuickView(window['-TABLE_RIGHT-'].get()[table_clicked_right])
            except:
                continue

    elif isinstance(event, tuple):
        if event[0] == '-TABLE_LEFT-':
            table_clicked_left = event[2][0]
            table_clicked_right = False
        elif event[0] == '-TABLE_RIGHT-':
            table_clicked_right = event[2][0]
            table_clicked_left = False

    # if event and len(values['-IN_LEFT-']) > 0:
    #     # picked_left = values['-IN_LEFT-'].replace("'", "").replace(",", "").replace("(", "").replace(")", "")
    #     # parts_left = [i[1:] for i in parts if i[0] == picked_left]
    #     # window['-TABLE_LEFT-'].update(parts_left)
    #     compare("input")
    #
    # if event and len(values['-IN_RIGHT-']) > 0:
    #     # picked_right = values['-IN_RIGHT-'].replace("'", "").replace(",", "").replace("(", "").replace(")", "")
    #     # parts_right = [i[1:] for i in parts if i[0] == picked_right]
    #     # window['-TABLE_RIGHT-'].update(parts_right)
    #     compare("input")

    elif event == 'Clear Tables':
        window['-IN_LEFT-'].update('')
        window['-BOX-CONTAINER_LEFT-'].update(visible=False)
        window['-IN_RIGHT-'].update('')
        window['-BOX-CONTAINER_RIGHT-'].update(visible=False)
        window['-TABLE_LEFT-'].update([])
        window['-TABLE_RIGHT-'].update([])

    elif event == '-CLEAR_FILES-':
        products, parts = [], []
        window['-IN_LEFT-'].update('')
        window['-BOX-CONTAINER_LEFT-'].update(visible=False)
        window['-IN_RIGHT-'].update('')
        window['-BOX-CONTAINER_RIGHT-'].update(visible=False)
        window['-TABLE_LEFT-'].update([])
        window['-TABLE_RIGHT-'].update([])

    # if event.startswith('Escape'):
    #     window['-IN_LEFT-'].update('')
    #     window['-BOX-CONTAINER_LEFT-'].update(visible=False)
    #     window['-IN_RIGHT-'].update('')
    #     window['-BOX-CONTAINER_RIGHT-'].update(visible=False)
    #     print("teraz")
    # elif event.startswith('Down') and len(prediction_list_left):
    #     sel_item_left = (sel_item_left + 1) % len(prediction_list_left)
    #     list_element_left.update(set_to_index=sel_item_left, scroll_to_index=sel_item_left)
    # elif event.startswith('Up') and len(prediction_list_left):
    #     sel_item_left = (sel_item_left + (len(prediction_list_left) - 1)) % len(prediction_list_left)
    #     list_element_left.update(set_to_index=sel_item_left, scroll_to_index=sel_item_left)
    elif event == '\r':
        if len(values['-BOX_LEFT-']) > 0:
            window['-IN_LEFT-'].update(value=values['-BOX_LEFT-'])
            window['-BOX-CONTAINER_LEFT-'].update(visible=False)
        if len(values['-BOX_RIGHT-']) > 0:
            window['-IN_RIGHT-'].update(value=values['-BOX_RIGHT-'])
            window['-BOX-CONTAINER_RIGHT-'].update(visible=False)

    elif event == '-IN_LEFT-':
        text = values['-IN_LEFT-']
        if text == input_text_left:
            continue
        else:
            input_text_left = text
        prediction_list_left = []
        if text:
            prediction_list_left = [item for item in products if item.startswith(text)]

        list_element_left.update(values=prediction_list_left)
        sel_item_left = 0
        list_element_left.update(set_to_index=sel_item_left)

        if len(prediction_list_left) > 0:
            window['-BOX-CONTAINER_LEFT-'].update(visible=True)
        else:
            window['-BOX-CONTAINER_LEFT-'].update(visible=False)
        compare("input")

    elif event == '-IN_RIGHT-':
        text = values['-IN_RIGHT-']
        if text == input_text_right:
            continue
        else:
            input_text_right = text
        prediction_list_right = []
        if text:
            prediction_list_right = [item for item in products if item.startswith(text)]

        list_element_right.update(values=prediction_list_right)
        sel_item_right = 0
        list_element_right.update(set_to_index=sel_item_right)

        if len(prediction_list_right) > 0:
            window['-BOX-CONTAINER_RIGHT-'].update(visible=True)
        else:
            window['-BOX-CONTAINER_RIGHT-'].update(visible=False)

        compare("input")

    elif event == '-BOX_LEFT-':
        window['-IN_LEFT-'].update(value=values['-BOX_LEFT-'])
        window['-BOX-CONTAINER_LEFT-'].update(visible=False)
        # picked_left = values['-BOX_LEFT-'][0]
        # parts_left = [i[1:] for i in parts if i[0] == picked_left]
        # window['-TABLE_LEFT-'].update(parts_left)

        compare("box")

    elif event == '-BOX_RIGHT-':
        window['-IN_RIGHT-'].update(value=values['-BOX_RIGHT-'])
        window['-BOX-CONTAINER_RIGHT-'].update(visible=False)
        # picked_right = values['-BOX_RIGHT-'][0]
        # parts_right = [i[1:] for i in parts if i[0] == picked_right]
        # window['-TABLE_RIGHT-'].update(parts_right)
        compare("box")

window.close()
