import tkinter as tk
from tkinter import ttk, filedialog, Menu, Text, StringVar
import binascii
from base64 import b64decode
from PIL import Image, ImageTk
import io

class VcfParser:
    def parse_vcf(self, vcf_content):
        lines = vcf_content.strip().split('\n')
        vcard_entries = []
        current_entry = []
        in_vcard = False

        for line in lines:
            line = line.strip()
            if line.startswith('BEGIN:VCARD'):
                if in_vcard:
                    vcard_entries.append(current_entry)
                current_entry = [line]
                in_vcard = True
            elif line.startswith('END:VCARD'):
                current_entry.append(line)
                vcard_entries.append(current_entry)
                current_entry = []
                in_vcard = False
            elif in_vcard:
                current_entry.append(line)

        if in_vcard and current_entry:
            vcard_entries.append(current_entry)

        contacts = []
        for entry in vcard_entries:
            processed_lines = []
            current_line = None

            for line in entry:
                if line.startswith('='):
                    if current_line is not None:
                        current_line += line[1:]
                else:
                    if current_line is not None:
                        processed_lines.append(current_line)
                    current_line = line

            if current_line is not None:
                processed_lines.append(current_line)

            name = None
            phones = []
            photo_data = None
            original_lines = entry.copy()
            has_photo = False

            for idx, line in enumerate(processed_lines):
                if line.startswith('END:VCARD'):
                    break
                if ':' not in line:
                    continue
                key, value = line.split(':', 1)
                key = key.upper()

                if key.startswith('N'):
                    if 'CHARSET=UTF-8' in key and 'ENCODING=QUOTED-PRINTABLE' in key:
                        value = value.replace('==', '=')
                        try:
                            decoded_bytes = binascii.a2b_qp(value)
                            name = decoded_bytes.decode('utf-8').replace(';', ' ')
                        except Exception as e:
                            name = f"Error decoding N: {e}"
                elif key.startswith('TEL'):
                    phones.append(value)
                elif key.startswith('PHOTO'):
                    has_photo = True
                    if 'BASE64' in key:
                        photo_lines = [value]
                        next_idx = idx + 1
                        while next_idx < len(processed_lines):
                            next_line = processed_lines[next_idx]
                            if next_line.startswith(' ') or ':' not in next_line:
                                photo_lines.append(next_line.strip())
                                next_idx += 1
                            else:
                                break
                        photo_data = ''.join(photo_lines).replace(' ', '').replace('\n', '')

            if name:
                main_phone = phones[0] if phones else None
                additional_phones = ', '.join(phones[1:]) if len(phones) > 1 else ''
                contacts.append({
                    'name': name,
                    'phone': main_phone,
                    'additional_phones': additional_phones,
                    'original_lines': original_lines,
                    'selected': tk.BooleanVar(value=False),
                    'has_photo': has_photo,
                    'photo_data': photo_data
                })

        return contacts

class ContactViewer:
    def __init__(self, root):
        self.root = root
        self.all_contacts = []
        self.contacts = []
        self.sort_column = None
        self.sort_reverse = False
        self.search_var = StringVar()
        self.create_widgets()
        self.tree.bind('<Button-1>', self.on_tree_click)
        self.tree.bind('<Double-1>', self.show_photo)
        self.current_image = None

    def create_widgets(self):
        menubar = Menu(self.root)
        filemenu = Menu(menubar, tearoff=0)
        filemenu.add_command(label="Import VCF", command=self.import_vcf)
        filemenu.add_command(label="Save VCF", command=self.save_vcf)
        filemenu.add_command(label="Delete Selected", command=self.delete_selected)
        filemenu.add_command(label="Delete Contacts Without Phone", command=self.delete_contacts_without_phone)
        menubar.add_cascade(label="File", menu=filemenu)
        self.root.config(menu=menubar)

        search_frame = tk.Frame(self.root)
        search_frame.pack(padx=10, pady=5, fill=tk.X)
        lbl_search = tk.Label(search_frame, text="Search:")
        lbl_search.pack(side=tk.RIGHT)
        self.entry_search = tk.Entry(search_frame, textvariable=self.search_var, width=30)
        self.entry_search.pack(side=tk.RIGHT, padx=5)
        self.entry_search.bind('<KeyRelease>', self.filter_contacts)
        btn_clear = tk.Button(search_frame, text="Clear", command=self.clear_search)
        btn_clear.pack(side=tk.LEFT)

        container = tk.Frame(self.root)
        container.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        self.tree = ttk.Treeview(container, 
                               columns=("#", "Name", "Phone", "Additional Phones", "Photo", "Select"), 
                               show="headings", 
                               selectmode='extended')
        
        columns_config = [
            ("#", 50, tk.CENTER),
            ("Name", 200, tk.W),
            ("Phone", 150, tk.W),
            ("Additional Phones", 200, tk.W),
            ("Photo", 60, tk.CENTER),
            ("Select", 40, tk.CENTER)
        ]

        for col, width, anchor in columns_config:
            if col in ["#", "Name", "Phone", "Additional Phones", "Photo", "Select"]:  # اضافه شد
                self.tree.heading(
                    col, 
                    text=col, 
                    command=lambda c=col: self.sort_contacts(c)
                )
            else:
                self.tree.heading(col, text=col)

        self.tree.tag_configure('selected', background='#d0f0d0')

        yscrollbar = ttk.Scrollbar(container, orient="vertical", command=self.tree.yview)
        xscrollbar = ttk.Scrollbar(container, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=yscrollbar.set, xscrollcommand=xscrollbar.set)

        self.tree.pack(side="left", fill="both", expand=True)
        yscrollbar.pack(side="right", fill="y")
        xscrollbar.pack(side="bottom", fill="x")

        self.error_log = Text(self.root, height=5, width=50)
        self.error_log.pack(fill=tk.X, padx=10, pady=10)
        self.error_log.configure(state='disabled')

        self.image_frame = tk.Frame(self.root)
        self.image_frame.pack(padx=10, pady=10)
        self.image_label = tk.Label(self.image_frame)
        self.image_label.pack()

    def import_vcf(self):
        file_path = filedialog.askopenfilename(filetypes=[("VCF files", "*.vcf")])
        if not file_path:
            return

        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                content = f.read()
            parser = VcfParser()
            self.all_contacts = parser.parse_vcf(content)
            self.contacts = self.all_contacts.copy()
            self.display_contacts()
        except Exception as e:
            self.log_error(f"Error importing VCF: {str(e)}")

    def display_contacts(self):
        for item in self.tree.get_children():
            self.tree.delete(item)

        for i, contact in enumerate(self.contacts, start=1):
            tags = ('selected',) if contact['selected'].get() else ()
            checkbox_state = "☑" if contact['selected'].get() else "☐"
            photo_icon = "🖼️" if contact['has_photo'] else ""
            self.tree.insert("", "end", 
                           values=(
                               str(i),
                               contact['name'],
                               contact['phone'] or 'No Phone',
                               contact['additional_phones'] or '-',
                               photo_icon,
                               checkbox_state
                           ), tags=tags)

    def on_tree_click(self, event):
        region = self.tree.identify_region(event.x, event.y)
        if region != "cell":
            return

        column = self.tree.identify_column(event.x)
        item = self.tree.identify_row(event.y)
        
        if column == "#6":
            current_value = self.tree.set(item, "Select")
            new_value = "☐" if current_value == "☑" else "☑"
            self.tree.set(item, "Select", new_value)
            
            index = int(self.tree.set(item, "#")) - 1
            self.contacts[index]['selected'].set(new_value == "☑")
            self.display_contacts()

    def sort_contacts(self, column):
        if self.sort_column == column:
            self.sort_reverse = not self.sort_reverse
        else:
            self.sort_column = column
            self.sort_reverse = False

        reverse_sort = self.sort_reverse
        
        if column == "Name":
            key = lambda x: x['name'].lower()
        elif column == "Phone":
            key = lambda x: x['phone'] or ''
        elif column == "#":
            key = lambda x: self.all_contacts.index(x)
        elif column == "Additional Phones":  # اضافه شده
            key = lambda x: len(x['additional_phones'].split(', ')) if x['additional_phones'] else 0
        elif column == "Photo":
            key = lambda x: not x['has_photo']
            reverse_sort = not reverse_sort
        elif column == "Select":
            key = lambda x: not x['selected'].get()
            reverse_sort = not reverse_sort
        else:
            return

        self.contacts.sort(key=key, reverse=reverse_sort)
        self.display_contacts()

    def delete_selected(self):
        selected_indices = [i for i, c in enumerate(self.contacts) if c['selected'].get()]
        for i in reversed(selected_indices):
            contact = self.contacts[i]
            if contact in self.all_contacts:
                self.all_contacts.remove(contact)
            del self.contacts[i]
        self.display_contacts()

    def delete_contacts_without_phone(self):
        initial_count = len(self.contacts)
        self.contacts = [c for c in self.contacts if c['phone'] is not None and c['phone'].strip() != '']
        removed_count = initial_count - len(self.contacts)
        
        if removed_count > 0:
            self.all_contacts = [c for c in self.all_contacts if c['phone'] is not None and c['phone'].strip() != '']
            self.display_contacts()
            self.log_error(f"Removed {removed_count} contacts without phone numbers")
        else:
            self.log_error("No contacts without phone numbers found")

    def save_vcf(self):
        if not self.contacts:
            self.log_error("No contacts to save")
            return

        file_path = filedialog.asksaveasfilename(defaultextension=".vcf", filetypes=[("VCF files", "*.vcf")])
        if not file_path:
            return

        try:
            with open(file_path, 'w', encoding='utf-8') as f:
                for contact in self.contacts:
                    in_photo = False
                    for line in contact['original_lines']:
                        stripped_line = line.strip()
                        
                        if stripped_line == '':
                            f.write('\n')
                            continue
                            
                        if stripped_line.upper().startswith('PHOTO'):
                            f.write(stripped_line + '\n')
                            in_photo = True
                        elif in_photo:
                            if stripped_line.startswith('END:VCARD'):
                                f.write(line + '\n')
                                in_photo = False
                            else:
                                f.write(' ' + line.lstrip() + '\n')
                        else:
                            f.write(line + '\n')
            self.log_error("VCF saved successfully")
        except Exception as e:
            self.log_error(f"Error saving VCF: {str(e)}")

    def show_photo(self, event):
        selected_items = self.tree.selection()
        if not selected_items:
            return
        
        item = selected_items[0]
        region = self.tree.identify_region(event.x, event.y)
        if region != "cell":
            return
        
        index = int(self.tree.set(item, "#")) - 1
        contact = self.contacts[index]
        
        if contact['photo_data']:
            try:
                data = contact['photo_data']
                missing_padding = len(data) % 4
                if missing_padding:
                    data += '=' * (4 - missing_padding)
                
                image_data = b64decode(data)
                
                header = image_data[:32].decode('ascii', errors='ignore')
                if 'JFIF' not in header and 'PNG' not in header:
                    raise ValueError("Invalid image format")
                
                image = Image.open(io.BytesIO(image_data))
                image.verify()
                image = Image.open(io.BytesIO(image_data))
                image.thumbnail((300, 300))
                self.current_image = ImageTk.PhotoImage(image)
                self.image_label.config(image=self.current_image)
            except Exception as e:
                self.log_error(f"Image Error: {str(e)}")
                self.image_label.config(image='')
        else:
            self.image_label.config(image='')
            self.log_error("No photo available")

    def filter_contacts(self, event=None):
        # تبدیل کاراکترهای عربی به فارسی قبل از جستجو
        search_term = self.search_var.get().lower()
        search_term = search_term.replace('ي', 'ی')  # تبدیل ي عربی به ی فارسی
        search_term = search_term.replace('ك', 'ک')  # تبدیل ك عربی به ک فارسی
        
        if not search_term:
            self.contacts = self.all_contacts.copy()
        else:
            self.contacts = [
                c for c in self.all_contacts
                if (search_term in c['name'].lower().replace('ي', 'ی') or 
                    (c['phone'] and search_term in c['phone']) or 
                    (c['additional_phones'] and search_term in c['additional_phones'].replace('ي', 'ی')))
            ]
        
        self.display_contacts()

    def clear_search(self):
        self.search_var.set('')
        self.contacts = self.all_contacts.copy()
        self.display_contacts()

    def log_error(self, message):
        self.error_log.configure(state='normal')
        self.error_log.insert(tk.END, message + '\n')
        self.error_log.configure(state='disabled')
        self.error_log.see(tk.END)

if __name__ == "__main__":
    root = tk.Tk()
    root.title("VCF Viewer")
    app = ContactViewer(root)
    root.mainloop()