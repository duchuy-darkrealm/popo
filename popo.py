import sys
from myopenpyxl import MyOpenpyxl
from openpyxl import Workbook

DATA_FILE_NAME = "data.xlsx"
TEMPLATE_FILE_NAME = "template.txt"
RESULT_FILE_NAME = "result.txt"

def getParamList(string, tag):
    string = string.replace("<" + tag,"")
    string = string.replace(">","")
    params = string.split("$")
    data = {}
    for param in params:
        values = param.split("=")
        if len(values)!= 2:
            continue
        data[values[0].strip()] = values[1].strip()
    return data

def isContainTag(string,tag):
    string = string.strip()
    start=string.find("<" + tag)
    end=string.find(">")
    if start != -1 and end!=-1 and start<end:
        return True
    return False
        
def replaceString(row,string):
    for key in row:
        string = string.replace("${" + str(key) + "}",str(row[key]))
        string = string.replace("${uppercase(" + str(key) + ")}",str(row[key]).upper())
        string = string.replace("${lowercase(" + str(key) + ")}",str(row[key]).lower())
        string = string.replace("${capitalize(" + str(key) + ")}",str(row[key]).capitalize())
        string = string.replace("${camelize(" + str(key) + ")}",camelize(str(row[key])))
        string = string.replace("${uppercamelize(" + str(key) + ")}",uppercamelize(str(row[key])))
        
    string = string.replace("${newline}","")
    return string

def camelize(string):
    words = string.split("_")
    result=""
    for i in range(len(words)):
        if i == 0:
            result += words[i]
        else:
            result += words[i].capitalize()
    return result

def uppercamelize(string):
    words = string.split("_")
    result = ""
    for i in range(len(words)):
        result += words[i].capitalize()
    return result

def findMarkup():
    result = []

    #find the end tag </loop>
    end = -1
    for i in range(0,len(lines)):
        if lines[i].find("</loop>") != -1:
            end = i
            break

    #if not found end tag, return as False
    if end == -1 :
        for line in lines:
            result.append(line)
        return result, False

    #if found end tag, find the nearest start tag <loop>
    start = -1
    for i in range(0,end-1):
        if isContainTag(lines[i],"loop"):
            start = i
            
    print("Execute loop in line: " + str(start) + "-" + str(end))
    #save the result with a loop from start to end
    #save from beginning of file to start point
    for i in range(0,start):
        result.append(lines[i])

    #loop from start point to end point with data get from param
    param = getParamList(lines[start],"loop")
    if "data" not in param :
        print("Parse error: not found $data in <loop>")
        sys.exit()
    
    filename = param["data"]
    rows = MyOpenpyxl().readValue(filename)
    for row in rows:
        for i in range(start+1,end):
            string = replaceString(row,lines[i])
            result.append(string)

    #save from end point to end of file
    for i in range(end+1,len(lines)):
        result.append(lines[i])

    return result, True

#load command line param
if len(sys.argv) == 1:
    print(" --------------------")
    print("|      popo.py       |")
    print(" --------------------")
    print("Description: Import load template and exel file, then save to result file")
    print("------------------")
    print("Argument 1: file name")
    print("------------------")
    print("Template Example:")
    print("")
    print("This is a list of animal:")
    print("<loop $data = data.xlsx>")
    print("	${name} has the weight: ${weight}")
    print("	${name} owner is ${capitalize(ownername)}")
    print("${newline}")
    print("</loop>")
    print("And that's all of my animal")
    print("------------------")
    print("written by Duchuy © 2020")
    print("------------------")
    TEMPLATE_FILE_NAME = input("Please enter template file name: ")
    print("Load template: " + TEMPLATE_FILE_NAME)
else:
    print ("Load template: " +  sys.argv[1])
    TEMPLATE_FILE_NAME = sys.argv[1]

f = open(TEMPLATE_FILE_NAME,"r",encoding="utf-8")
lines = f.readlines()
rows = MyOpenpyxl().readValue(DATA_FILE_NAME)

while True:
    lines, result = findMarkup()
    if result == False:
        break

pos = TEMPLATE_FILE_NAME.rfind(".")
filename = TEMPLATE_FILE_NAME[0:pos] + "_result.txt" 
f = open(filename,"w",encoding="utf-8")

for line in lines:
    f.write(line)
f.close()

print("Saved as " + filename)
print("Written by Duchuy © 2020")













