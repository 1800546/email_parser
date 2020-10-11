# email_parser (Windows)

Parse ost or pst email datafile in working directory automatically via command line or manually.

This tool uses relative path and everything is based on the working directory (path of command line). The datafile must be named either "input.ost" or "input.pst" and be placed in the working directory. The "email_parser.exe" executable should be placed inside the "libraries" sub-folder in the working directory.

Multiple csv files will be generated inside the "emailtemp" sub-folder in working directory, depending on the structure of the input datafile.

#Sample codes (Python2 for Windows)
First, let's create the necessary sub-folder
```
os.makedirs("emailtemp")
```
Here is how to generate the output
```
if os.path.isdir("emailtemp"):
    os.system('\"'+abspath("libraries/email_parser.exe")+'\"')
```
Below is how to parse the multiple csv files generated
```
regexlineslist = []
for filename in os.listdir("emailtemp"):
    if filename.endswith(".csv"):
        with open(os.path.join("emailtemp", filename)) as fpregex:
            for lineregex in csv.reader(x.replace('\0', '') for x in fpregex):
                regexlineslist.append(re.sub(r'(\s|\u180B|\u200B|\u200C|\u200D|\u2060|\uFEFF)+', '',str(lineregex).strip())[2:-2])
```
The following regex can be used to match with emails addresses and timestamps in the list populated in the code above.
```
email = re.compile("^[0-9a-zA-Z_.-]+@[0-9a-zA-Z][0-9a-zA-Z.-]*[0-9a-zA-Z]\.[a-zA-Z]{2,64}$")
timestamps = re.compile("^\d{1,2}[A-Z][a-z]{2}\d{5,6}:\d{2}:\d{2}(a|p)m$")
```
However, before using the parsing or regex codes above, it is recommended to run the extraction multiple times to ensure that there was no silent errors the fist time. This is because the tool is not stable when automated and the results have to be verified. Fortunately, there is an easy way to verify the results as shown below. It is recommended to run it 20 times to be safe; but for time-critical situations, 10 times would probably suffice.
```
emailvar6 = 0
def askSequence2():
    global emailvar6
    emailvar6 = raw_input("Enter the number of times to scan the ost/pst file (Recommended min 20): ")
values2 = list("0123456789")
def remover2(my_string = ""):
  for item in my_string:
    if item not in values2:
      my_string = my_string.replace(item, "")
  return my_string
askSequence2()
while ((int(emailvar6) == 0)or(len(emailvar6) != len(remover2(emailvar6)))):
    print "Invalid input. Please try again."
    askSequence2()

emailvar7 = 1
emailvar8 = 0
while (emailvar7 <= int(emailvar6)):
    try:
        shutil.rmtree("emailtemp")
    except:
        pass
    os.makedirs("emailtemp")
    if os.path.isdir("emailtemp"):
        os.system('\"'+abspath("libraries/email_parser.exe")+'\"')
    emailpath, emaildirs, emailfiles = next(os.walk("emailtemp"))
    emailfilecount = len(emailfiles)
    print "Iteration "+str(emailvar7)+ " of " + emailvar6 +": "+str(emailfilecount)
    emailvar7 = emailvar7+1
    if (emailfilecount > emailvar8):
        emailvar8 = emailfilecount
print "The assumed number of files is " + str(emailvar8) + " files."

def getEFCount():
    emailpath, emaildirs, emailfiles = next(os.walk("emailtemp"))
    return len(emailfiles)

while (getEFCount() != emailvar8):
    try:
        shutil.rmtree("emailtemp")
    except:
        pass
    os.makedirs("emailtemp")
    if os.path.isdir("emailtemp"):
        os.system('\"'+abspath("libraries/email_parser.exe")+'\"')
    emailpath, emaildirs, emailfiles = next(os.walk("emailtemp"))
    emailfilecount = len(emailfiles)
    print "Current fetch: "+str(emailfilecount)+ " files."
print "CSV files successfully verified"
```
