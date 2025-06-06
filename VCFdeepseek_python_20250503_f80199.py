import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import quopri

class VCFViewer:
    def __init__(self, root):
        self.root = root
        self.root.title("VCF Viewer")
        
        self.contacts = []
        self.current_sort = None
        
        self.create_widgets()
        
    def create_widgets(self):
        # Top Frame for Buttons
        top_frame = ttk.Frame(self.root)
        top_frame.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Button(top_frame, text="Open VCF", command=self.load_vcf).pack(side=tk.LEFT)
        ttk.Button(top_frame, text="Sort by Name", command=lambda: self.sort_contacts('name')).pack(side=tk.LEFT, padx=5)
        ttk.Button(top_frame, text="Sort by Tel", command=lambda: self.sort_contacts('tel')).pack(side=tk.LEFT)
        ttk.Button(top_frame, text="Delete Selected", command=self.delete_contacts).pack(side=tk.LEFT, padx=5)
        ttk.Button(top_frame, text="Save VCF", command=self.save_vcf).pack(side=tk.LEFT)
        
        # Treeview
        self.tree = ttk.Treeview(self.root, columns=('Select', 'Name', 'Tel'), show='headings')
        self.tree.heading('Select', text='Select')
        self.tree.heading('Name', text='Name')
        self.tree.heading('Tel', text='Tel')
        self.tree.column('Select', width=50, anchor='center')
        self.tree.column('Name', width=200)
        self.tree.column('Tel', width=150)
        
        vsb = ttk.Scrollbar(self.root, orient="vertical", command=self.tree.yview)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)
        self.tree.configure(yscrollcommand=vsb.set)
        
        self.tree.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        self.tree.bind('<Button-1>', self.on_tree_click)
        
        # Error Console
        self.error_console = scrolledtext.ScrolledText(self.root, height=8, state=tk.DISABLED)
        self.error_console.pack(fill=tk.BOTH, padx=5, pady=5)
        
        ttk.Button(self.root, text="Copy Errors", command=self.copy_errors).pack(pady=5)
        
    def load_vcf(self):
        path = filedialog.askopenfilename(filetypes=[("VCF Files", "*.vcf")])
        if not path:
            return
        
        try:
            with open(path, 'r', encoding='utf-8') as f:
                content = f.read()
            
            self.contacts = self.parse_vcf(content)
            self.update_treeview()
        except Exception as e:
            self.show_error(f"Error loading file: {str(e)}")
    
    def parse_vcf(self, content):
        contacts = []
        cards = content.split('BEGIN:VCARD')
        for card in cards[1:]:  # Skip empty first element
            card = 'BEGIN:VCARD' + card
            lines = []
            current_line = ''
            
            # Process line folding
            for line in card.splitlines():
                line = line.rstrip('\r\n')
                if line.startswith((' ', '\t')):
                    current_line += line[1:]
                else:
                    if current_line:
                        lines.append(current_line)
                    current_line = line
            if current_line:
                lines.append(current_line)
            
            # Process encoding and parameters
            contact = {'original': [], 'name': '', 'tel': ''}
            for line in lines:
                contact['original'].append(line)
                
                if line.startswith('N;') or line.startswith('FN;'):
                    parts = line.split(':', 1)
                    if len(parts) < 2:
                        continue
                    
                    # Decode Quoted-Printable
                    value = parts[1]
                    if 'QUOTED-PRINTABLE' in line.upper():
                        value = value.replace('=\n', '')
                        value = quopri.decodestring(value.encode('utf-8')).decode('utf-8', 'ignore')
                    
                    contact['name'] = value.replace('_', ' ').strip()
                
                elif line.startswith('TEL;'):
                    contact['tel'] = line.split(':', 1)[-1]
            
            contacts.append(contact)
        return contacts
    
    def update_treeview(self):
        for item in self.tree.get_children():
            self.tree.delete(item)
            
        for contact in self.contacts:
            self.tree.insert('', 'end', values=('☐', contact['name'], contact['tel']))
    
    def on_tree_click(self, event):
        region = self.tree.identify_region(event.x, event.y)
        if region != 'cell':
            return
        
        col = self.tree.identify_column(event.x)
        item = self.tree.identify_row(event.y)
        
        if col == '#1':  # Select column
            current_val = self.tree.item(item, 'values')[0]
            new_val = '☑' if current_val == '☐' else '☐'
            values = list(self.tree.item(item, 'values'))
            values[0] = new_val
            self.tree.item(item, values=values)
    
    def sort_contacts(self, key):
        self.contacts.sort(key=lambda x: x[key].lower())
        self.update_treeview()
    
    def delete_contacts(self):
        to_delete = []
        for item in self.tree.get_children():
            if self.tree.item(item, 'values')[0] == '☑':
                to_delete.append(item)
        
        for item in reversed(to_delete):
            index = self.tree.index(item)
            del self.contacts[index]
            self.tree.delete(item)
    
    def save_vcf(self):
        path = filedialog.asksaveasfilename(defaultextension=".vcf")
        if not path:
            return
        
        try:
            with open(path, 'w', encoding='utf-8') as f:
                for contact in self.contacts:
                    f.write('\n'.join(contact['original']))
                    f.write('\nEND:VCARD\n')
            messagebox.showinfo("Success", "File saved successfully")
        except Exception as e:
            self.show_error(f"Error saving file: {str(e)}")
    
    def show_error(self, message):
        self.error_console.config(state=tk.NORMAL)
        self.error_console.insert(tk.END, message + '\n')
        self.error_console.config(state=tk.DISABLED)
        self.error_console.see(tk.END)
    
    def copy_errors(self):
        self.root.clipboard_clear()
        self.root.clipboard_append(self.error_console.get('1.0', tk.END))

if __name__ == '__main__':
    root = tk.Tk()
    VCFViewer(root)
    root.mainloop()