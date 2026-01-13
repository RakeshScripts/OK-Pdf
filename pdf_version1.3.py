import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pikepdf
import threading
import time


class SecurePDFToolGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("OK PDF Merger & Splitter")
        self.root.geometry("800x650")
        self.root.config(bg="#2e2e2e")

        self.pdf_list = []
        self.split_pdf_path = None

        # Styles
        style = ttk.Style()
        style.theme_use('clam')
        style.configure("TNotebook", background="#2e2e2e", foreground="#fff")
        style.configure("TNotebook.Tab", background="#444", foreground="#fff", padding=[10, 5])
        style.map("TNotebook.Tab", background=[("selected", "#1e90ff")])

        # Tabs
        self.tab_control = ttk.Notebook(root)
        self.merge_tab = ttk.Frame(self.tab_control)
        self.split_tab = ttk.Frame(self.tab_control)
        self.tab_control.add(self.merge_tab, text='Merge PDFs')
        self.tab_control.add(self.split_tab, text='Split/Extract Options')
        self.tab_control.pack(expand=1, fill="both")

        self.setup_merge_tab()
        self.setup_split_tab()

        dev_label = tk.Label(root,
                             text="Developed by Rakesh Rokade (Body & Trims) | Email: rr927039.ttl@tatamotors.com",
                             bg="#2e2e2e", fg="#aaaaaa", font=("Times New Roman", 9))
        dev_label.pack(side=tk.BOTTOM, pady=5)

    def setup_merge_tab(self):
        frame = tk.Frame(self.merge_tab, bg="#2e2e2e")
        frame.pack(fill=tk.BOTH, expand=1, padx=20, pady=20)

        self.listbox = tk.Listbox(frame, font=("Times New Roman", 12), selectbackground="#1e90ff", bg="#1e1e1e",
                                  fg="white")
        self.listbox.pack(fill=tk.BOTH, expand=1, side=tk.LEFT)

        scrollbar = tk.Scrollbar(frame, command=self.listbox.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.listbox.config(yscrollcommand=scrollbar.set)

        btn_frame = tk.Frame(self.merge_tab, bg="#2e2e2e")
        btn_frame.pack(fill=tk.X, padx=20)

        tk.Button(btn_frame, text="Add PDFs", command=self.add_pdfs, bg="#1e90ff", fg="white", width=12).pack(
            side=tk.LEFT, padx=5, pady=10)
        tk.Button(btn_frame, text="Move Up", command=self.move_up, bg="#444", fg="white", width=10).pack(side=tk.LEFT,
                                                                                                         padx=5)
        tk.Button(btn_frame, text="Move Down", command=self.move_down, bg="#444", fg="white", width=10).pack(
            side=tk.LEFT, padx=5)
        tk.Button(btn_frame, text="Remove", command=self.remove_selected, bg="#dc3545", fg="white", width=10).pack(
            side=tk.LEFT, padx=5)
        tk.Button(btn_frame, text="Merge PDFs", command=lambda: threading.Thread(target=self.merge_pdfs).start(),
                  bg="#28a745", fg="white", width=15).pack(side=tk.RIGHT, padx=5)

        self.merge_progress = ttk.Progressbar(self.merge_tab, orient=tk.HORIZONTAL, length=500, mode='determinate')
        self.merge_progress.pack(pady=10)

    def setup_split_tab(self):
        container = tk.Frame(self.split_tab, bg="#2e2e2e")
        container.pack(fill=tk.BOTH, expand=1, padx=30, pady=20)

        tk.Button(container, text="1. Select Source PDF", command=self.select_pdf, bg="#1e90ff", fg="white",
                  font=("Times New Roman", 12, "bold")).pack(pady=10)

        self.file_label = tk.Label(container, text="No file selected", bg="#2e2e2e", fg="#fbff00",
                                   font=("Arial", 9, "italic"))
        self.file_label.pack()

        # Variable with Trace for UI updates
        self.split_mode = tk.IntVar(value=1)
        self.split_mode.trace_add("write", self.update_split_ui)

        mode_frame = tk.LabelFrame(container, text=" 2. Select Operation Mode ", bg="#2e2e2e", fg="white",
                                   font=("Times New Roman", 12, "bold"), padx=10, pady=10)
        mode_frame.pack(fill=tk.X, pady=20)

        tk.Radiobutton(mode_frame, text="Mode 1: Select Pages to Keep (Range A to B)", variable=self.split_mode,
                       value=1, bg="#2e2e2e", fg="white", selectcolor="#1e90ff", activebackground="#2e2e2e").pack(
            anchor=tk.W)
        tk.Radiobutton(mode_frame, text="Mode 2: Remove specific page (Keep others)", variable=self.split_mode, value=2,
                       bg="#2e2e2e", fg="white", selectcolor="#1e90ff", activebackground="#2e2e2e").pack(anchor=tk.W)
        tk.Radiobutton(mode_frame, text="Mode 3: Keep only one specific page", variable=self.split_mode, value=3,
                       bg="#2e2e2e", fg="white", selectcolor="#1e90ff", activebackground="#2e2e2e").pack(anchor=tk.W)

        # Input Area
        self.input_frame = tk.Frame(container, bg="#2e2e2e")
        self.input_frame.pack(pady=10)

        self.label1 = tk.Label(self.input_frame, text="Start Page:", bg="#2e2e2e", fg="white")
        self.label1.grid(row=0, column=0, padx=5)
        self.start_entry = tk.Entry(self.input_frame, width=10)
        self.start_entry.grid(row=0, column=1, padx=5)

        self.label2 = tk.Label(self.input_frame, text="End Page:", bg="#2e2e2e", fg="white")
        self.label2.grid(row=0, column=2, padx=5)
        self.end_entry = tk.Entry(self.input_frame, width=10)
        self.end_entry.grid(row=0, column=3, padx=5)

        tk.Button(container, text="Process & Save PDF",
                  command=lambda: threading.Thread(target=self.process_split).start(),
                  bg="#28a745", fg="white", font=("Times New Roman", 12, "bold"), height=2).pack(pady=20)

        self.split_progress = ttk.Progressbar(container, orient=tk.HORIZONTAL, length=500, mode='determinate')
        self.split_progress.pack(pady=10)

    def update_split_ui(self, *args):
        """Changes labels and enables/disables entries based on mode"""
        mode = self.split_mode.get()
        if mode == 1:
            self.label1.config(text="Start Page:")
            self.label2.grid()
            self.end_entry.grid()
        elif mode == 2:
            self.label1.config(text="Page to REMOVE:")
            self.label2.grid_remove()
            self.end_entry.grid_remove()
        elif mode == 3:
            self.label1.config(text="Page to KEEP:")
            self.label2.grid_remove()
            self.end_entry.grid_remove()

    def add_pdfs(self):
        files = filedialog.askopenfilenames(filetypes=[("PDF Files", "*.pdf")])
        for f in files:
            self.pdf_list.append(f)
            self.listbox.insert(tk.END, f.split("/")[-1])

    def remove_selected(self):
        idxs = self.listbox.curselection()
        for idx in reversed(idxs):
            self.pdf_list.pop(idx)
            self.listbox.delete(idx)

    def move_up(self):
        idxs = self.listbox.curselection()
        for idx in idxs:
            if idx == 0: continue
            self.pdf_list[idx], self.pdf_list[idx - 1] = self.pdf_list[idx - 1], self.pdf_list[idx]
            val = self.listbox.get(idx)
            self.listbox.delete(idx)
            self.listbox.insert(idx - 1, val)
            self.listbox.select_set(idx - 1)

    def move_down(self):
        idxs = self.listbox.curselection()
        for idx in reversed(idxs):
            if idx == len(self.pdf_list) - 1: continue
            self.pdf_list[idx], self.pdf_list[idx + 1] = self.pdf_list[idx + 1], self.pdf_list[idx]
            val = self.listbox.get(idx)
            self.listbox.delete(idx)
            self.listbox.insert(idx + 1, val)
            self.listbox.select_set(idx + 1)

    def merge_pdfs(self):
        if not self.pdf_list:
            messagebox.showwarning("Warning", "No PDFs selected")
            return
        save_path = filedialog.asksaveasfilename(defaultextension=".pdf")
        if not save_path: return
        try:
            merged_pdf = pikepdf.Pdf.new()
            total = len(self.pdf_list)
            for idx, pdf_file in enumerate(self.pdf_list, start=1):
                with pikepdf.Pdf.open(pdf_file) as src:
                    merged_pdf.pages.extend(src.pages)
                self.merge_progress['value'] = (idx / total) * 100
                self.root.update_idletasks()
            merged_pdf.save(save_path)
            merged_pdf.close()
            messagebox.showinfo("Success", "PDFs merged successfully!")
            self.merge_progress['value'] = 0
        except Exception as e:
            messagebox.showerror("Error", str(e))

    def select_pdf(self):
        path = filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf")])
        if path:
            self.split_pdf_path = path
            self.file_label.config(text=f"Selected: {path.split('/')[-1]}")

    def process_split(self):
        if not self.split_pdf_path:
            messagebox.showerror("Error", "Please select a source PDF first.")
            return
        try:
            mode = self.split_mode.get()
            val1 = int(self.start_entry.get())

            with pikepdf.Pdf.open(self.split_pdf_path) as src:
                total_pages = len(src.pages)
                new_pdf = pikepdf.Pdf.new()

                # Logic for different modes
                if mode == 1:  # Keep Range
                    val2 = int(self.end_entry.get())
                    pages_to_process = src.pages[val1 - 1: val2]
                elif mode == 2:  # Remove Specific Page
                    pages_to_process = [src.pages[i] for i in range(total_pages) if (i + 1) != val1]
                elif mode == 3:  # Keep Specific Page
                    pages_to_process = [src.pages[val1 - 1]]

                total = len(pages_to_process)
                for idx, pg in enumerate(pages_to_process, 1):
                    new_pdf.pages.append(pg)
                    self.split_progress['value'] = (idx / total) * 100
                    self.root.update_idletasks()

                save_path = filedialog.asksaveasfilename(defaultextension=".pdf")
                if save_path:
                    new_pdf.save(save_path)
                    new_pdf.close()
                    messagebox.showinfo("Success", "Action Completed!")
                self.split_progress['value'] = 0
        except Exception as e:
            messagebox.showerror("Error", f"Ensure page numbers are correct.\nDetails: {str(e)}")


def show_splash():
    splash = tk.Tk()
    splash.overrideredirect(True)
    splash.geometry("400x200+500+300")
    splash.config(bg="#1e90ff")
    tk.Label(splash, text="OK PDF TOOL", bg="#1e90ff", fg="white", font=("Arial", 20, "bold")).pack(expand=True)
    splash.update()
    time.sleep(2)
    splash.destroy()


if __name__ == "__main__":
    show_splash()
    root = tk.Tk()
    app = SecurePDFToolGUI(root)
    root.mainloop()
