dim server
set server = WScript.CreateObject("OdfConverterHost.Entier")

dim a,b

a = 12
b = server.AuCarre(a)
WScript.Echo a & "^2=" & b

dim c,d
c = "C:\Documents and Settings\Administrateur\Mes documents\Invitation_EN.odt"
d = "C:\TestOOP\Test.docx"

set server = WScript.CreateObject("ConverterSimpleWrapper.Converter")

server.OdfToOox c, d, true
