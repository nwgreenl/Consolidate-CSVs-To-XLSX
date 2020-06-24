import datetime, os, re, sys, pandas as pd
from tkinter import filedialog

# get files via tkinter
cwd = os.getcwd()
files = filedialog.askopenfilenames(initialdir=cwd, title="Select File", filetypes=[("CSV Files", ".csv")])
if len(files) == 0:
    sys.exit("\nNo Files Selected... Goodbye!\n")

# output dir
outputDirName = "output"
if not os.path.exists("%s/%s" % (cwd, outputDirName)):
    os.mkdir("%s/%s" % (cwd, outputDirName))

# output file 
outputDate = datetime.datetime.now().strftime("%m-%d-%Y_%I-%M") 
outputFileName = "consolidated_%s" % outputDate
outputExt = "xlsx"
outputFile = "%s/%s/%s.%s" % (cwd, outputDirName, outputFileName, outputExt)

# sheet name regex
legalChars = re.compile("[^a-zA-Z0-9]")

# consolidate
print("Running...\n")
with pd.ExcelWriter(outputFile, engine="xlsxwriter") as writer:  
    for file in files:
        # sheet names must be <= 31 chars and cannot contain "\ / * ? : ,"
        # opting to remove any char that isn't a word or digit
        fileNameForSheet = legalChars.sub("", os.path.basename(file)).replace("csv", "")[:31]
        
        df = pd.read_csv(file)
        df.to_excel(writer, sheet_name=fileNameForSheet, index=False)

# success message
isInputPlural = "s" if len(files) > 1 else ""
print("Successfully created '%s' by consolidating the following CSV%s:\n  - " % (outputFileName, isInputPlural), end="")
print(*files, sep = "\n  - ")