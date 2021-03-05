import docx2txt
import re
import pandas as pd
textfile = docx2txt.process("Daatabase.docx")
#pattern = re.compile(r"[a-zA-Z0-9]+@[a-zA-Z]+\.(com|edu|net)")
#matches = pattern.finditer(textfile)
#for match in matches:
 #   print(match.group(0))
print(textfile)
fobj=open("test_data.txt","w")
fobj.write(textfile)
fobj.close()
fobj=open("test_data.txt")
fobj1=open("test_data1.txt","w")
test_str=""
for line in fobj:
    l1=line.strip()
    if l1!="":
        print(l1,file=fobj1)
        test_str+=l1+"\n"
fobj.close()
fobj1.close()
print(test_str)
test_dict={"Name":[],"Email":[]}
regex = r"^Name\W.*|Email-Id\W.*"
matches = re.finditer(regex, test_str, re.MULTILINE)
fobj=open("email_data.txt","w")
for matchNum, match in enumerate(matches, start=1):
    
    print ("Match {matchNum} was found at {start}-{end}: {match}".format(matchNum = matchNum, start = match.start(), end = match.end(), match = match.group()))
    print(match.group().split("\n")[1],file=fobj)
    if matchNum%2!=0:
        test_dict["Name"].append(match.group().split("\n")[1])
    else:
        test_dict["Email"].append(match.group().split("\n")[1])
    
#text = textfile[textfile.find("!Name")+1:textfile.find("Message")]
#print(text)
regex = r"^First Name\W.*|^Last Name\W.*|Email :\W.*"
matches = re.finditer(regex, test_str, re.MULTILINE)
for matchNum, match in enumerate(matches, start=1):
    
    print ("Match {matchNum} was found at {start}-{end}: {match}".format(matchNum = matchNum, start = match.start(), end = match.end(), match = match.group()))
    print(match.group().split(" : ")[1], file=fobj)
    if matchNum%3==1:
        test_dict["Name"].append(match.group().split(" : ")[1])
    elif matchNum%3==0:
        test_dict["Email"].append(match.group().split(" : ")[1])
unfobj.close()
print(test_dict)
print(len(test_dict["Name"]),len(test_dict["Email"]))
df = pd.DataFrame(test_dict)
print(df)
df.to_csv('details.csv', index=False)
Collapse



