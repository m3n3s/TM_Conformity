import tkinter as tk
from tkinter import ttk
from tkinter import filedialog as fd
import reporter

# create the root window
root = tk.Tk()
root.title('TM Conformity Reporter')
root.resizable(False, False)
root.geometry('300x200')

csvPath = ""
outDir = ""

def selectFile():
    filetypes = (
        ('csv files', '*.csv'),
        ('All files', '*.*')
    )

    global csvPath
    csvPath = fd.askopenfilename(
        title='Open a file',
        initialdir='/',
        filetypes=filetypes)
    
    print(csvPath)
    csv_label["text"] = csv_label["text"] + "\n" + csvPath

def selectDirectory():
    global outDir
    outDir = fd.askdirectory()

    print(dir)
    output_label["text"] = output_label["text"] + "\n" + outDir

email = tk.StringVar()
password = tk.StringVar()

# Entry frame
entry = ttk.Frame(root)
entry.pack(padx=10, pady=10, fill='x', expand=True)


# CSV file path
csv_label = ttk.Label(entry, text="Select a CSV file: ")
csv_label.pack(fill='x', expand=True)

csv_button = ttk.Button(
    entry,
    text='Select a File',
    command=selectFile
)
csv_button.pack(fill='x', expand=True)

# Output file directory
output_label = ttk.Label(entry, text="Select an output directory: ")
output_label.pack(fill='x', expand=True)

output_entry = ttk.Button(
    entry,
    text='Select a Directory',
    command=selectDirectory
)
output_entry.pack(fill='x', expand=True)

# Generate button
generate_button = ttk.Button(
    entry, 
    text="Generate",
    command=reporter.generatePPTX(csvPath, outDir),
)
generate_button.pack(pady=20, ipady=10)


# run the application
root.mainloop()
