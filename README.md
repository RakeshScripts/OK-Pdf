OK PDF Merger & Splitter
A lightweight, secure, and user-friendly desktop application to merge and split PDF files. Developed specifically for engineering and administrative workflows.

🚀 Features
Merge PDFs: Combine multiple PDF files into one. Reorder files easily with 'Move Up' and 'Move Down' buttons.

Split Mode 1 (Range): Extract a specific range of pages (e.g., Pages 5 to 10).

Split Mode 2 (Remove): Remove a specific page and keep the rest of the document.

Split Mode 3 (Specific): Extract exactly one specific page into a new file.

Dynamic UI: The interface changes based on the selected mode to prevent input errors.

Dark Theme: Easy on the eyes for long working hours.

Multithreaded: The GUI remains responsive even during heavy processing.

🛠️ Installation & Dependencies
To run the source code, you need Python 3.x installed.

1. Clone the repository
Bash

git clone https://github.com/RakeshScripts/OK-Pdf.git
cd ok-pdf-tool
2. Install required libraries
This project uses pikepdf for high-performance PDF manipulation and tkinter (which comes standard with most Python installs).

Bash

pip install pikepdf pyinstaller
📦 How to create the EXE (Windows)
To share this tool with users who do not have Python installed, you can convert the script into a standalone .exe file.

Step 1: Prepare your Icon
Ensure your icon file (e.g., app_icon.ico) is in the same folder as your script.

Step 2: Run PyInstaller
Open your terminal in the project folder and run the following command:

Bash

pyinstaller --noconsole --onefile --add-data "app_icon.ico;." --icon="app_icon.ico" --name "OK_PDF_Tool" main.py
Flag Explanations:

--noconsole: Hides the black terminal window when the app opens.

--onefile: Bundles everything into a single executable.

--add-data "app_icon.ico;.": Embeds the icon file inside the EXE (required for the window title bar icon to work).

--icon="app_icon.ico": Sets the icon for the file in Windows Explorer.

--name: The final name of your application.

Step 3: Collect your EXE
Once finished, the final executable will be located in the dist/ folder.

🖥️ Usage
1) Select Tab: Choose between "Merge" or "Split".

2) Add Files: Use the "Add/Select" buttons to load your PDFs.

3) Choose Mode: (For Splitting) Use the radio buttons to pick your desired operation.

4) Process: Click the green "Process" button and choose where to save your new file.

👨‍💻 Developer
Rakesh Rokade
