import PySimpleGUI as sg
import PLM_BOM_transpose


def PopupQuickView(subassembly):
    if not any(subassembly[0] in i for i in products): #check if subassembly is loaded. #any because products is list of tuples and need to iterate throught it
        sg.Popup("No such subassembly loaded")
        return None

    subassembly_parts = sorted([i[1:] for i in parts if subassembly[0] in i[0]]) #create list of parts that are in picked subassembly. #i[1:] because first position in parts is product number which is not needed here
    #declaration for QuickView table
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
        if event_q == 'Quick view': #event for creating QuickView in QuickView
            window_q.close() #close primary window
            if not PopupQuickView(window_q['-TABLE_QUICKVIEW-'].get()[values_q['-TABLE_QUICKVIEW-'][0]]): #return None if picked part doesnt have subassembly
                PopupQuickView(subassembly) #create primary window if picked part doesnt have subassembly
    return True


def compare(source):
    parts_right, parts_left = [], []

    if len(values['-IN_RIGHT-']) > 0: #check if input is not empty
        if source == "input": #distinguish sources if writen by user or picked from BOX
            picked_right = values['-IN_RIGHT-'].replace("'", "").replace(",", "").replace("(", "").replace(")", "") #removes ',() from input
        else:
            picked_right = values['-BOX_RIGHT-'][0]
        parts_right = sorted([i[1:] for i in parts if i[0] == picked_right]) #create list of parts that are in picked product
        window['-TABLE_RIGHT-'].update(parts_right) #update Table

    else:
        window['-TABLE_RIGHT-'].update([]) #if user erased input set table to empty

    if len(values['-IN_LEFT-']) > 0: #check if input is not empty
        if source == "input": #distinguish sources if writen by user or picked from BOX
            picked_left = values['-IN_LEFT-'].replace("'", "").replace(",", "").replace("(", "").replace(")", "") #removes ',() from input
        else:
            picked_left = values['-BOX_LEFT-'][0]
        parts_left = sorted([i[1:] for i in parts if i[0] == picked_left]) #create list of parts that are in picked product
        window['-TABLE_LEFT-'].update(parts_left) #update Table

    else:
        window['-TABLE_LEFT-'].update([]) #if user erased input set table to empty

    if len(window['-TABLE_RIGHT-'].get()) > 1 and len(window['-TABLE_LEFT-'].get()) > 1: #if both inputs are not empty
        for count, part in enumerate(parts_left): #enumerate through parts from left table
            index = -1
            for i, obj in enumerate(parts_right): #enumerate through parts from right table
                if obj[0] == part[0]: #if part found in right table
                    if obj[1] == part[1]: #if it have same quantity
                        index = i
                        break
                    index = -2
                    break
            if index == -1:
                window['-TABLE_LEFT-'].update(row_colors=((count, 'red'),)) #if part if not found in right table set row color to red
            elif index == -2:
                window['-TABLE_LEFT-'].update(row_colors=((count, 'black', 'yellow'),)) #if part if found in right table but with different quantity set row color to yellow

        for count, part in enumerate(parts_right): #enumerate through parts from right table
            index = -1
            for i, obj in enumerate(parts_left): #enumerate through parts from right table
                if obj[0] == part[0]: #if part found in left table
                    if obj[1] == part[1]: #if it have same quantity
                        index = i
                        break
                    index = -2
                    break
            if index == -1:
                window['-TABLE_RIGHT-'].update(row_colors=((count, 'red'),)) #if part if not found in right table set row color to red
            elif index == -2:
                window['-TABLE_RIGHT-'].update(row_colors=((count, 'black', 'yellow'),)) #if part if found in right table but with different quantity set row color to yellow
    else: #ensure no rows are colored if both inputs are not picked
        window['-TABLE_LEFT-'].update(row_colors=((0, ''),))
        window['-TABLE_RIGHT-'].update(row_colors=((0, ''),))


headings = ["Part", "Qty", "Description"] #heading for tables

#column declaration with left table
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

#column declaration with right table
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

#column declaration with top buttons
col3 = sg.Column([[sg.Frame('Actions:',
                            [[sg.Column([[sg.Button('Load Data', key='-FILES-'), sg.Button('Clear All', key='-CLEAR_FILES-'), sg.Button('Clear Tables'), ]],
                    pad=(0, 0))]])]], pad=(0, 0))

#column declaration with left input
col4 = sg.Column([
    [sg.Text('Kod motoru:')],
    [sg.Input(size=(20, 1), enable_events=True, key='-IN_LEFT-')],
    [sg.pin(sg.Col([[sg.Listbox(values=[], size=(20, 4), enable_events=True, key='-BOX_LEFT-',
                                select_mode=sg.LISTBOX_SELECT_MODE_SINGLE, no_scrollbar=True)]],
                   key='-BOX-CONTAINER_LEFT-', pad=(0, 0), visible=False))]])

#column declaration with right input
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
        compare("box")

    elif event == '-BOX_RIGHT-':
        window['-IN_RIGHT-'].update(value=values['-BOX_RIGHT-'])
        window['-BOX-CONTAINER_RIGHT-'].update(visible=False)
        compare("box")

window.close()
