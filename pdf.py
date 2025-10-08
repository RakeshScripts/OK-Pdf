import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pikepdf
import threading
import time


class SecurePDFToolGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("OK PDF Merger & Splitter")
        self.root.geometry("750x500")
        self.root.config(bg="#2e2e2e")  # Dark background

        self.pdf_list = []

        # Styles
        style = ttk.Style()
        style.theme_use('clam')
        style.configure("TNotebook", background="#2e2e2e", foreground="#fff")
        style.configure("TNotebook.Tab", background="#444", foreground="#fff", padding=[10, 5])
        style.map("TNotebook.Tab", background=[("selected", "#1e90ff")])
        style.configure("TButton", background="#1e90ff", foreground="#fff", font=("Times New Roman", 10, "bold"))

        # Tabs
        self.tab_control = ttk.Notebook(root)
        self.merge_tab = ttk.Frame(self.tab_control, style="TFrame")
        self.split_tab = ttk.Frame(self.tab_control, style="TFrame")
        self.tab_control.add(self.merge_tab, text='Merge PDFs')
        self.tab_control.add(self.split_tab, text='Split PDF')
        self.tab_control.pack(expand=1, fill="both")

        self.setup_merge_tab()
        self.setup_split_tab()

        # Developer name
        dev_label = tk.Label(root, text="Developed by Rakesh Rokade (Body & Trims) Email: rr927039.ttl@tatamotors.com", bg="#2e2e2e", fg="white",
                             font=("Times New Roman", 10, ))
        dev_label.pack(side=tk.BOTTOM, pady=5)

    # Merge Tab
    def setup_merge_tab(self):
        frame = tk.Frame(self.merge_tab, bg="#2e2e2e")
        frame.pack(fill=tk.BOTH, expand=1, padx=20, pady=20)

        self.listbox = tk.Listbox(frame, font=("Times New Roman", 12), selectbackground="#1e90ff", selectforeground="#fff")
        self.listbox.pack(fill=tk.BOTH, expand=1, side=tk.LEFT)

        scrollbar = tk.Scrollbar(frame, command=self.listbox.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.listbox.config(yscrollcommand=scrollbar.set)

        btn_frame = tk.Frame(self.merge_tab, bg="#2e2e2e")
        btn_frame.pack(fill=tk.X, padx=20)

        tk.Button(btn_frame, text="Add PDFs", command=self.add_pdfs, bg="#1e90ff", fg="white").pack(side=tk.LEFT,
                                                                                                    padx=5, pady=10)
        tk.Button(btn_frame, text="Move Up", command=self.move_up, bg="#1e90ff", fg="white").pack(side=tk.LEFT, padx=5)
        tk.Button(btn_frame, text="Move Down", command=self.move_down, bg="#1e90ff", fg="white").pack(side=tk.LEFT,
                                                                                                      padx=5)
        tk.Button(btn_frame, text="Merge PDFs", command=lambda: threading.Thread(target=self.merge_pdfs).start(),
                  bg="#28a745", fg="white").pack(side=tk.LEFT, padx=5)

        self.merge_progress = ttk.Progressbar(self.merge_tab, orient=tk.HORIZONTAL, length=500, mode='determinate')
        self.merge_progress.pack(pady=10)

    # Split Tab
    def setup_split_tab(self):
        tk.Button(self.split_tab, text="Select PDF", command=self.select_pdf, bg="#1e90ff", fg="white",
                  font=("Times New Roman", 12, "bold")).pack(pady=20)

        page_frame = tk.Frame(self.split_tab, bg="#2e2e2e")
        page_frame.pack(pady=10)

        tk.Label(page_frame, text="Start Page:", bg="#2e2e2e", fg="white", font=("Times New Roman", 12)).grid(row=0, column=0,
                                                                                                    padx=5)
        self.start_entry = tk.Entry(page_frame, width=10, font=("Times New Roman", 12))
        self.start_entry.grid(row=0, column=1, padx=5)

        tk.Label(page_frame, text="End Page:", bg="#2e2e2e", fg="white", font=("Times New Roman", 12)).grid(row=0, column=2,
                                                                                                  padx=5)
        self.end_entry = tk.Entry(page_frame, width=10, font=("Times New Roman", 12))
        self.end_entry.grid(row=0, column=3, padx=5)

        tk.Button(self.split_tab, text="Split PDF", command=lambda: threading.Thread(target=self.split_pdf).start(),
                  bg="#28a745", fg="white", font=("Times New Roman", 12, "bold")).pack(pady=10)

        self.split_progress = ttk.Progressbar(self.split_tab, orient=tk.HORIZONTAL, length=500, mode='determinate')
        self.split_progress.pack(pady=10)

    # Merge Functions
    def add_pdfs(self):
        files = filedialog.askopenfilenames(filetypes=[("PDF Files", "*.pdf")])
        for f in files:
            self.pdf_list.append(f)
            self.listbox.insert(tk.END, f.split("/")[-1])

    def move_up(self):
        idxs = self.listbox.curselection()
        for idx in idxs:
            if idx == 0:
                continue
            self.pdf_list[idx], self.pdf_list[idx - 1] = self.pdf_list[idx - 1], self.pdf_list[idx]
            self.listbox.insert(idx - 1, self.listbox.get(idx))
            self.listbox.delete(idx + 1)
            self.listbox.select_set(idx - 1)

    def move_down(self):
        idxs = self.listbox.curselection()
        for idx in reversed(idxs):
            if idx == len(self.pdf_list) - 1:
                continue
            self.pdf_list[idx], self.pdf_list[idx + 1] = self.pdf_list[idx + 1], self.pdf_list[idx]
            self.listbox.insert(idx + 2, self.listbox.get(idx))
            self.listbox.delete(idx)
            self.listbox.select_set(idx + 1)

    def merge_pdfs(self):
        if not self.pdf_list:
            messagebox.showwarning("Warning", "No PDFs selected")
            return
        save_path = filedialog.asksaveasfilename(defaultextension=".pdf")
        if save_path:
            merged_pdf = pikepdf.Pdf.new()
            try:
                total = len(self.pdf_list)
                for idx, pdf_file in enumerate(self.pdf_list, start=1):
                    src = pikepdf.Pdf.open(pdf_file)
                    merged_pdf.pages.extend(src.pages)
                    self.merge_progress['value'] = (idx / total) * 100
                    self.root.update_idletasks()
                    time.sleep(0.1)
                merged_pdf.save(save_path)
                merged_pdf.close()
                self.merge_progress['value'] = 0
                messagebox.showinfo("Success", f"PDFs merged into {save_path}")
            except Exception as e:
                messagebox.showerror("Error", str(e))

    # Split Functions
    def select_pdf(self):
        self.split_pdf_path = filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf")])
        if self.split_pdf_path:
            messagebox.showinfo("Selected PDF", f"Selected: {self.split_pdf_path}")

    def split_pdf(self):
        try:
            start = int(self.start_entry.get()) - 1
            end = int(self.end_entry.get())
            src = pikepdf.Pdf.open(self.split_pdf_path)
            split_pdf = pikepdf.Pdf.new()
            total = end - start
            for idx, page in enumerate(src.pages[start:end], start=1):
                split_pdf.pages.append(page)
                self.split_progress['value'] = (idx / total) * 100
                self.root.update_idletasks()
                time.sleep(0.1)
            save_path = filedialog.asksaveasfilename(defaultextension=".pdf")
            if save_path:
                split_pdf.save(save_path)
                split_pdf.close()
                self.split_progress['value'] = 0
                messagebox.showinfo("Success", f"PDF split saved as {save_path}")
        except Exception as e:
            messagebox.showerror("Error", str(e))


# Splash Screen
def show_splash():
    splash = tk.Tk()
    splash.overrideredirect(True)
    splash.geometry("500x300+500+200")
    splash.config(bg="#5dd1ff")

    label = tk.Label(splash, text="Welcome to OK PDF", bg="#5dd1ff", fg="white", font=("Times New Roman", 24, "bold"))
    label.pack(expand=True)

    splash.update()
    splash.after(3000, splash.destroy)  # Display for 3 seconds
    splash.mainloop()


# Run App
if __name__ == "__main__":
    show_splash()
    root = tk.Tk()
    app = SecurePDFToolGUI(root)
    root.mainloop()
