# OfficeMacroEncoder
Encoder for VBScript macros, helps evade obvious detection and warnings in office 2013/2016.

I was reading this excellent tutorial on using SET+Metasploit to craft a malicious document: http://null-byte.wonderhowto.com/how-to/create-obfuscate-virus-inside-microsoft-word-document-0167780/ . The encoder used is written in Java, so I just converted it to Python.

Usage:

$ python shellcodemacro.txt obfuscatedoutput.txt

The file takes the macro in shellcodemacro.txt and converts it into an obfuscated macro, stored in obfuscatedoutput.txt.  You may then insert the obfuscated version of the macro into your Word/Excel document, set up a meterpreter payload stage per the blog post above, and fire away your spear phish.

As of November 2016, this encoding method gets 6/55 detection on VirusTotal. Neither McAfee nor Symantec detect the malicious file.

**Legal Warning** 

This tool should be used for educational and research purposes only. The user is responsible for using it in accordance with all laws which apply to them.
