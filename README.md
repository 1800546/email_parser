# email_parser (Windows)

Parse ost or pst email datafile in working directory automatically via command line or manually.

This tool uses relative path and everything is based on the working directory (path of command line). The datafile must be copied into the working directory and renamed to either "input.ost" or "input.pst". The "email_parser.exe" executable should be placed inside the "libraries" folder in the working directory.

Multiple csv files will be generated inside the "emailtemp" folder in working directory, which must be created before running the tool.

Since this tool and the sample codes provided are designed to recognise only this specific working directory structure, it is recommended to download/clone this repo entirely in order to avoid manual folder modifications. If you prefer manual modification, please refer to the "Working directory structure" section below.

# Working directory structure
Before running the tool
```
-libraries/
    -email_parser.exe
-input.xst
-emailtemp/
```
After running the tool
```
-libraries/
    -email_parser.exe
-input.xst
-emailtemp/
    -email_7iWqTD3MgLkFD0YGuNVKSxbSuG8Gfjvg.csv
    -email_9ocujoqrpIACjhP2ThEQMSqhxClU06HG.csv
    -email_b3Z8DrSTvdqvHmE348Kw9L0aHxApdey4.csv
    -email_cEbrK3mBPuOENRNvQaN2WUz7pAGcsWjv.csv
```
The number of files in the "emailtemp" folder depends on the structure of the input datafile. Be sure to run the tool multiple times to ensure accuracy, as mentioned in the "Limitations" section.

The tool can be used manually by direct execution with the layout above kept in mind, but it can also be automated. Some sample automation codes are provided in the "Sample automation codes" section below.

# Sample automation codes (Python2)

Before we start, please refer to "Annex A" section and ensure your windows time/date display configuration in metro mode (Settings app) and desktop mode (Control panel) are exactly the same as the image in Annex A. This sample code is optimised for this particular display configuration. Other display configurations will cause this sample code to not function, since the regex may not be able to adapt to locale-specific changes. If you are using other display configurations, you may have to create your own regex.

First, let's create the necessary folder
```
os.makedirs("emailtemp")
```
Here is how to generate the output
```
if os.path.isdir("emailtemp"):
    os.system('\"'+abspath("libraries/email_parser.exe")+'\"')
```
Below is how to read the multiple csv files generated
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
email = re.compile("^[0-9a-zA-Z_.-]+@[0-9a-zA-Z][0-9a-zA-Z.-]*[0-9a-zA-Z]\\.[a-zA-Z]{2,64}$")
timestamp = re.compile("^\\d{1,2}[A-Z][a-z]{2}\\d{5,6}:\\d{2}:\\d{2}(a|p)m$")
```
# Limitations
Before using the reading or parsing codes above, it is recommended to run the extraction multiple times to ensure that there was no silent errors the first time. This is because the tool is not stable when automated and the results have to be verified. Fortunately, there is an easy way to verify the results as shown below. It is recommended to run it 20 times to be safe; but for time-critical situations, 10 times would probably suffice.
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
# Testing

The datafile must be copied into the working directory and renamed to "input.ost" or "input.pst". Below are some ways you can obtain a datafile.

You can use your own outlook datafiles (ost or pst) which are found in the following folder:

C:\Users\ **your_user_name** \AppData\Local\Microsoft\Outlook\ **your_email@domain.com** .xst

Otherwise, a sample ost file can be downloaded for testing from:
```
https://www.cfreds.nist.gov/data_leakage_case/data-leakage-case.html
```
Click on the Personal Computer (PC) – 'DD' Image.

Once downloaded, use a tool like OSFMount to mount the image in writeable mode.
```
https://www.osforensics.com/tools/mount-disk-images.html
```
Inside the image, there is an outlook datafile at:
```
X:\Users\informant\AppData\Local\Microsoft\Outlook\iaman.informant@nist.gov.ost
```

# Annex A
This annex is for reference from the "Sample codes" section.
![Display Configuration](https://github.com/1800546/email_parser/raw/main/displayconfig.jpg)


