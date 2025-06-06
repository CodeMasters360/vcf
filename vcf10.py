import tkinter as tk
from tkinter import ttk, filedialog, Menu, Text
import binascii

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
            phone = None
            original_lines = entry.copy()

            for line in processed_lines:
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
                    phone = value
                    break

            if name:
                contacts.append({
                    'name': name,
                    'phone': phone,
                    'original_lines': original_lines,
                    'selected': tk.BooleanVar(value=False)
                })

        return contacts

class ContactViewer:
    def __init__(self, root):
        self.root = root
        self.contacts = []
        self.sort_column = None
        self.sort_reverse = False
        self.create_widgets()
        self.tree.bind('<Button-1>', self.on_tree_click)

    def create_widgets(self):
        menubar = Menu(self.root)
        filemenu = Menu(menubar, tearoff=0)
        filemenu.add_command(label="Import VCF", command=self.import_vcf)
        filemenu.add_command(label="Save VCF", command=self.save_vcf)
        filemenu.add_command(label="Delete Selected", command=self.delete_selected)
        filemenu.add_command(label="Delete Contacts Without Phone", command=self.delete_contacts_without_phone)  # Added menu item
        menubar.add_cascade(label="File", menu=filemenu)
        self.root.config(menu=menubar)

        container = tk.Frame(self.root)
        container.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        self.tree = ttk.Treeview(container, 
                               columns=("#", "Name", "Phone", "Select"), 
                               show="headings", 
                               selectmode='extended')
        
        self.tree.heading("#", text="#", command=lambda: self.sort_contacts("#"))
        self.tree.heading("Name", text="Name", command=lambda: self.sort_contacts("Name"))
        self.tree.heading("Phone", text="Phone", command=lambda: self.sort_contacts("Phone"))
        self.tree.heading("Select", text="✓", 
                        command=lambda: self.sort_contacts("Select"),
                        anchor=tk.CENTER)

        self.tree.column("#", width=50, anchor=tk.CENTER)
        self.tree.column("Name", width=200)
        self.tree.column("Phone", width=150)
        self.tree.column("Select", width=40, anchor=tk.CENTER)

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

    def import_vcf(self):
        file_path = filedialog.askopenfilename(filetypes=[("VCF files", "*.vcf")])
        if not file_path:
            return

        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                content = f.read()
            parser = VcfParser()
            self.contacts = parser.parse_vcf(content)
            self.display_contacts()
        except Exception as e:
            self.log_error(f"Error importing VCF: {str(e)}")

    def display_contacts(self):
        for item in self.tree.get_children():
            self.tree.delete(item)

        for i, contact in enumerate(self.contacts, start=1):
            tags = ('selected',) if contact['selected'].get() else ()
            checkbox_state = "☑" if contact['selected'].get() else "☐"
            self.tree.insert("", "end", 
                           values=(str(i), contact['name'], contact['phone'], checkbox_state),
                           tags=tags)

    def on_tree_click(self, event):
        region = self.tree.identify_region(event.x, event.y)
        if region != "cell":
            return

        column = self.tree.identify_column(event.x)
        item = self.tree.identify_row(event.y)
        
        if column == "#4":
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
            key = lambda x: self.contacts.index(x)
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
            del self.contacts[i]
        self.display_contacts()

    # Added method for deleting contacts without phone
    def delete_contacts_without_phone(self):
        initial_count = len(self.contacts)
        self.contacts = [c for c in self.contacts if c['phone'] is not None and c['phone'].strip() != '']
        removed_count = initial_count - len(self.contacts)
        
        if removed_count > 0:
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