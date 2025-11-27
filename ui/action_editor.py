import os
import tkinter
from tkinter import ttk
import tkinter.messagebox as messagebox

from core.actions import parse_action_line, compose_action_line


def resolve_action_path(name: str) -> str:
    if not name:
        return ''
    p_actions = os.path.join('actions', name)
    if os.path.exists(p_actions):
        return p_actions
    return name if os.path.exists(name) else ''


def open_action_editor(root, state, action_file_name: str, refresh_current_action, on_saved=None):
    """Open .action file editor; callbacks supplied by caller for selection refresh."""
    path = resolve_action_path(action_file_name)
    if not path:
        messagebox.showinfo('No file', 'Please select an .action file first.')
        return
    editor = tkinter.Toplevel(root)
    editor.title(f'Edit Action - {os.path.basename(path)}')
    editor.geometry('800x500')
    editor.grab_set()

    editor.grid_rowconfigure(1, weight=1)
    editor.grid_columnconfigure(0, weight=1)

    frame_top = ttk.Frame(editor)
    frame_top.grid(row=0, column=0, sticky='ew', padx=10, pady=(10, 0))

    frame = ttk.Frame(editor)
    frame.grid(row=1, column=0, sticky='nsew', padx=10, pady=10)

    estyle = ttk.Style(editor)
    estyle.configure('Action.Treeview', rowheight=40, padding=0)
    estyle.configure('Action.Heading', padding=0)

    columns = ('line','type','op','vk','btn','act','x','y','nx','ny','dx','dy','ms','raw')
    headings = {
        'line':'#','type':'Type','op':'Op','vk':'VK','btn':'Btn','act':'Act',
        'x':'X','y':'Y','nx':'NX','ny':'NY','dx':'DX','dy':'DY','ms':'MS','raw':'Raw'
    }
    tree = ttk.Treeview(frame, columns=columns, show='headings', style='Action.Treeview', selectmode='extended')
    for c in columns:
        tree.heading(c, text=headings[c], anchor='center')
    center_cols = ['line','type','op','vk','btn','act','x','y','nx','ny','dx','dy','ms']
    for c in center_cols:
        tree.column(c, width=60, anchor='center', stretch=True)
    tree.column('raw', width=200, anchor='w', stretch=True)
    tree.column('line', width=40)

    vsb = ttk.Scrollbar(frame, orient='vertical', command=tree.yview)
    tree.configure(yscrollcommand=vsb.set)
    tree.grid(row=0, column=0, sticky='nsew')
    vsb.grid(row=0, column=1, sticky='ns')
    frame.rowconfigure(0, weight=1)
    frame.columnconfigure(0, weight=1)

    def allowed_columns_for(row_type, row_op):
        if row_type == 'VK':
            return {'op','vk','ms'}
        if row_type == 'MS':
            if (row_op or '').upper() == 'MOVE':
                return {'op','x','y','nx','ny','ms'}
            if (row_op or '').upper() == 'CLICK':
                return {'op','btn','act','x','y','nx','ny','ms'}
            if (row_op or '').upper() == 'SCROLL':
                return {'op','dx','dy','ms'}
            return {'op','ms'}
        return set()

    try:
        with open(path, 'r', encoding='utf-8') as f:
            lines = [ln.rstrip('\n') for ln in f.readlines()]
    except Exception as e:
        messagebox.showerror('Error', f'Failed to open file:\n{e}')
        editor.destroy(); return

    meta_lines = []
    meta_screen = ''
    meta_start = ''
    data_lines = []
    for text in lines:
        if text.startswith('META '):
            meta_lines.append(text)
            parts = text.split()
            if len(parts)>=3 and parts[1] == 'SCREEN' and len(parts)>=4:
                meta_screen = f"{parts[2]}x{parts[3]}"
            if len(parts)>=3 and parts[1] == 'START':
                meta_start = parts[2]
        elif text.startswith('#') or text.strip() == '':
            continue
        else:
            data_lines.append(text)

    meta_disp = f"Screen: {meta_screen or 'N/A'}    Start: {meta_start or 'N/A'}"
    meta_label = ttk.Label(frame_top, text=meta_disp, style='Biz.TLabel')
    meta_label.pack(anchor='w')

    for idx, text in enumerate(data_lines, start=1):
        d = parse_action_line(text)
        vals = [idx, d['type'], d['op'], d['vk'], d['btn'], d['act'], d['x'], d['y'], d['nx'], d['ny'], d['dx'], d['dy'], d['ms'], d['raw']]
        tree.insert('', 'end', values=tuple(vals))

    edit_entry = None
    edit_item = None
    edit_col = None
    edit_commit = None
    dragging_selection = []
    drag_insert_index = None
    drag_scroll_job = None
    last_drag_y = None

    def cancel_inline_editor(*_):
        nonlocal edit_entry, edit_item, edit_col
        try:
            if edit_entry:
                edit_entry.destroy()
        except Exception:
            pass
        edit_entry = None
        edit_item = None
        edit_col = None

    def spawn_editor(row, colname, x, y, w, h):
        nonlocal edit_entry, edit_item, edit_col, edit_commit
        row_type = tree.set(row, 'type')
        row_op = tree.set(row, 'op')
        if colname not in allowed_columns_for(row_type, row_op) and colname not in ('type','raw'):
            return
        val = tree.set(row, colname)
        edit_item = row
        edit_col = colname

        def finish_edit_value(new_val):
            nonlocal edit_entry, edit_item, edit_col, edit_commit
            if edit_entry and edit_item and edit_col:
                if edit_col == 'type':
                    new_val = 'VK' if new_val == 'VK' else 'MS'
                    tree.set(edit_item, 'type', new_val)
                    if new_val == 'VK':
                        tree.set(edit_item, 'op', 'DOWN' if tree.set(edit_item,'op') not in ('DOWN','UP') else tree.set(edit_item,'op'))
                        for c in ['btn','act','x','y','nx','ny','dx','dy']: tree.set(edit_item, c, '')
                    else:
                        tree.set(edit_item, 'op', 'MOVE' if tree.set(edit_item,'op') not in ('MOVE','CLICK','SCROLL') else tree.set(edit_item,'op'))
                        tree.set(edit_item, 'vk', '')
                elif edit_col == 'op':
                    t = tree.set(edit_item,'type')
                    if t == 'VK':
                        new_val = 'DOWN' if new_val not in ('DOWN','UP') else new_val
                        tree.set(edit_item,'op', new_val)
                    else:
                        new_val = new_val if new_val in ('MOVE','CLICK','SCROLL') else 'MOVE'
                        tree.set(edit_item,'op', new_val)
                        if new_val == 'MOVE':
                            for c in ['btn','act','dx','dy']: tree.set(edit_item,c,'')
                        elif new_val == 'CLICK':
                            for c in ['dx','dy']: tree.set(edit_item,c,'')
                            if tree.set(edit_item,'btn') not in ('left','right'): tree.set(edit_item,'btn','left')
                            if tree.set(edit_item,'act') not in ('DOWN','UP'): tree.set(edit_item,'act','DOWN')
                        elif new_val == 'SCROLL':
                            for c in ['btn','act','x','y','nx','ny']: tree.set(edit_item,c,'')
                else:
                    int_fields = {'vk','x','y','dx','dy','ms'}
                    float_fields = {'nx','ny'}
                    if edit_col in int_fields and new_val != '':
                        try:
                            int(new_val)
                        except Exception:
                            return
                    if edit_col in float_fields and new_val != '':
                        try:
                            float(new_val)
                        except Exception:
                            return
                    tree.set(edit_item, edit_col, new_val)
                cancel_inline_editor()
        edit_commit = finish_edit_value

        def on_enter(event):
            finish_edit_value(edit_entry.get())
        def on_focus_out(event):
            finish_edit_value(edit_entry.get())

        if colname in ('type','op'):
            values = ['VK','MS'] if colname == 'type' else (['DOWN','UP'] if row_type=='VK' else ['MOVE','CLICK','SCROLL'])
            edit_entry = ttk.Combobox(frame, values=values, state='readonly')
            edit_entry.set(val)
        else:
            edit_entry = ttk.Entry(frame)
            edit_entry.insert(0, val)
        edit_entry.place(in_=tree, x=x, y=y, width=w, height=h)
        edit_entry.focus_set()
        edit_entry.bind('<Return>', on_enter)
        edit_entry.bind('<FocusOut>', on_focus_out)

    def commit_inline_if_any():
        nonlocal edit_entry, edit_item, edit_col, edit_commit
        try:
            if edit_entry and edit_commit:
                edit_commit(edit_entry.get())
        except Exception:
            pass
        cancel_inline_editor()

    def begin_edit(event):
        region = tree.identify('region', event.x, event.y)
        if region != 'cell':
            return
        colid = tree.identify_column(event.x)
        if colid == '#1':
            return
        row = tree.identify_row(event.y)
        if not row:
            return
        bbox = tree.bbox(row, colid)
        if not bbox:
            return
        x, y, w, h = bbox
        spawn_editor(row, columns[int(colid[1:]) - 1], x, y, w, h)

    tree.bind('<Double-1>', begin_edit)

    def add_row():
        commit_inline_if_any()
        vals = {'type':'VK','op':'DOWN','vk':'0','btn':'','act':'','x':'','y':'','nx':'','ny':'','dx':'','dy':'','ms':'0','raw':''}
        sel = tree.selection()
        insert_index = tree.index(sel[-1]) + 1 if sel else 'end'
        iid = tree.insert('', insert_index, values=( '', vals['type'], vals['op'], vals['vk'], vals['btn'], vals['act'], vals['x'], vals['y'], vals['nx'], vals['ny'], vals['dx'], vals['dy'], vals['ms'], vals['raw']))
        renumber_lines()
        spawn_editor(iid, 'type', *tree.bbox(iid, '#2'))

    def delete_rows():
        commit_inline_if_any()
        sel = tree.selection()
        if not sel:
            return
        for iid in sel:
            tree.delete(iid)
        renumber_lines()

    def renumber_lines():
        for idx, iid in enumerate(tree.get_children(''), start=1):
            tree.set(iid, 'line', str(idx))

    def save_and_close():
        commit_inline_if_any()
        try:
            items = tree.get_children('')
            new_lines = []
            for i in items:
                vals = {c: tree.set(i, c) for c in columns}
                new_line = compose_action_line(vals)
                new_lines.append(new_line)
            with open(path, 'w', encoding='utf-8') as f:
                if meta_lines:
                    f.write('\n'.join(meta_lines) + '\n')
                f.write('\n'.join(new_lines) + '\n')
            if callable(refresh_current_action):
                refresh_current_action()
            if callable(on_saved):
                on_saved(path)
        except Exception as e:
            messagebox.showerror('Error', f'Failed to save file:\n{e}')
            return
        editor.destroy()

    def move_drag(event):
        nonlocal drag_insert_index, dragging_selection, last_drag_y, drag_scroll_job
        if not dragging_selection:
            sel_now = list(tree.selection()) or [tree.identify_row(event.y)]
            sel_now = [s for s in sel_now if s]
            children = tree.get_children('')
            idx_map = {iid: i for i, iid in enumerate(children)}
            dragging_selection = sorted(sel_now, key=lambda i: idx_map.get(i, 0))
            try:
                tree.selection_set(tuple(dragging_selection))
            except Exception:
                pass
        last_drag_y = event.y
        insert_at, line_y = compute_insert_at(event)
        drag_insert_index = insert_at
        try:
            insert_line.place(in_=tree, x=0, y=line_y, width=tree.winfo_width(), height=2)
        except Exception:
            pass

    def compute_insert_at(event):
        children = tree.get_children('')
        if not children:
            return 0, 0
        target = tree.identify_row(event.y)
        if target:
            idx = tree.index(target)
            bbox = tree.bbox(target)
            if bbox and len(bbox) == 4:
                x, y, w, h = bbox
                above = event.y < (y + h/2)
                insert_at = idx if above else idx + 1
                line_y = y if above else y + h
            else:
                h_widget = max(0, tree.winfo_height())
                insert_at = idx
                line_y = max(0, min(event.y, h_widget))
            return insert_at, line_y
        first_bbox = tree.bbox(children[0])
        last_bbox = tree.bbox(children[-1])
        fy = first_bbox[1] if first_bbox and len(first_bbox) == 4 else 0
        ly = last_bbox[1] if last_bbox and len(last_bbox) == 4 else max(0, tree.winfo_height() - 1)
        lh = last_bbox[3] if last_bbox and len(last_bbox) == 4 else 0
        if event.y < fy:
            return 0, fy
        else:
            return len(children), ly + lh

    def on_tree_press(event):
        nonlocal dragging_selection, drag_insert_index
        commit_inline_if_any(event)
        row = tree.identify_row(event.y)
        if not row:
            dragging_selection = []
            drag_insert_index = None
            clear_drag_highlight()
            return
        ctrl = (event.state & 0x0004) != 0
        shift = (event.state & 0x0001) != 0
        current_sel = list(tree.selection())
        if (row not in current_sel) and not (ctrl or shift):
            tree.selection_set((row,))
        drag_insert_index = None
        clear_drag_highlight()

    insert_line = tkinter.Frame(tree, height=2, background='#2563eb')
    def on_tree_motion(event):
        nonlocal drag_insert_index, dragging_selection, last_drag_y, drag_scroll_job
        if not dragging_selection:
            sel_now = list(tree.selection())
            if not sel_now:
                sel_now = [tree.identify_row(event.y)]
            sel_now = [s for s in sel_now if s]
            children = tree.get_children('')
            idx_map = {iid: i for i, iid in enumerate(children)}
            dragging_selection = sorted(sel_now, key=lambda i: idx_map.get(i, 0))
            try:
                tree.selection_set(tuple(dragging_selection))
            except Exception:
                pass
        last_drag_y = event.y
        insert_at, line_y = compute_insert_at(event)
        drag_insert_index = insert_at
        try:
            insert_line.place(in_=tree, x=0, y=line_y, width=tree.winfo_width(), height=2)
        except Exception:
            pass

    def clear_drag_highlight():
        try:
            insert_line.place_forget()
        except Exception:
            pass

    def on_tree_release(event):
        nonlocal dragging_selection, drag_insert_index
        try:
            insert_line.place_forget()
        except Exception:
            pass
        if not dragging_selection:
            return
        children = list(tree.get_children(''))
        idx_map = {iid: i for i, iid in enumerate(children)}
        if drag_insert_index is None:
            dragging_selection = []
            drag_insert_index = None
            clear_drag_highlight()
            return
        sel_sorted = sorted(dragging_selection, key=lambda i: idx_map.get(i, 0))
        base = [iid for iid in children if iid not in sel_sorted]
        insertion = drag_insert_index
        before_count = sum(1 for iid in sel_sorted if idx_map[iid] < drag_insert_index)
        insertion -= before_count
        insertion = max(0, min(insertion, len(base)))
        new_order = base[:insertion] + sel_sorted + base[insertion:]
        for pos, iid in enumerate(new_order):
            try:
                tree.move(iid, '', pos)
            except Exception:
                pass
        renumber_lines()
        try:
            tree.selection_set(tuple(sel_sorted))
            if sel_sorted:
                tree.focus(sel_sorted[0])
        except Exception:
            pass
        dragging_selection = []
        drag_insert_index = None
        clear_drag_highlight()

    tree.bind('<ButtonPress-1>', on_tree_press, add='+')
    tree.bind('<B1-Motion>', on_tree_motion, add='+')
    tree.bind('<ButtonRelease-1>', on_tree_release, add='+')

    btn_frame = ttk.Frame(editor)
    btn_frame.grid(row=2, column=0, sticky='ew', padx=10, pady=(0,10))

    save_btn = ttk.Button(btn_frame, text='Save', command=save_and_close)
    save_btn.pack(side='right', padx=6)
    close_btn = ttk.Button(btn_frame, text='Close', command=editor.destroy)
    close_btn.pack(side='right')
    add_btn = ttk.Button(btn_frame, text='Add', command=add_row)
    add_btn.pack(side='left')
    del_btn = ttk.Button(btn_frame, text='Delete', command=delete_rows)
    del_btn.pack(side='left', padx=6)

    return editor
