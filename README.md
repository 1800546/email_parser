# email_parser (Windows)

Parse ost or pst email datafile in working directory automatically via command line or manually.

This tool uses relative path and everything is based on the working directory (path of command line). The datafile must be named either "input.ost" or "input.pst" and be placed in the working directory. The executable should be placed inside the "libraries" folder in the working directory.

Below is a sample python2 code to process the csv files generated (generated inside the "emailtemp" folder in working directory)

'''
regexlineslist = []
for filename in os.listdir("emailtemp"):
    if filename.endswith(".csv"):
        with open(os.path.join("emailtemp", filename)) as fpregex:
            for lineregex in csv.reader(x.replace('\0', '') for x in fpregex):
                regexlineslist.append(re.sub(r'(\s|\u180B|\u200B|\u200C|\u200D|\u2060|\uFEFF)+', '',str(lineregex).strip())[2:-2])
'''

