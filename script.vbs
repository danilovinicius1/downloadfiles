on error resume next
dim dia 
dim tamanhodia
dia = day(now)
dim data

dim hora
dim minuto
dim segundo
dim tempo

dim nomearquivocoelba
dim nomearquivocelpe
dim nomearquivocosern

dim icoelba
dim icelpe
dim icosern

icoelba = 0
icelpe  = 0
icosern = 0


Function LPad (str, pad, length)
    LPad = String(length - Len(str), pad) & str
End Function

hora = LPad(Hour(Time), "0", 2)
minuto = LPad(Minute(Time), "0", 2)
segundo = LPad(Second(Time), "0", 2)

dim mes
dim tamanhomes
mes = month(now)

dim ano
ano = year(now)

tamanhodia = Len(dia)
tamanhomes = Len(mes)

if tamanhodia = 1 then
dia = "0" & day(now) 
else
dia = day(now)
end if

if tamanhomes = 1 then
mes = "0" & month(now)
else
mes = month(now)
end if

data = dia & "." & mes & "." & ano

tempo = hora & "." & minuto & "." & segundo


'msgbox data & "-" & tempo

dim xHttpcoelba
dim bStrmcoelba
set xHttpcoelba = createobject("Microsoft.XMLHTTP")
Set bStrmcoelba = createobject("Adodb.Stream")


do
icoelba = icoelba + 1
	xHttpcoelba.Open "GET", "portalweb", False
xHttpcoelba.Send
Wscript.sleep (5000)
loop until icoelba = 10
nomearquivocoelba =  "d:\Monitoramentofatmail\Coelba\" & "Monitoramento-Coelba" & "-" & data & "-" & tempo & ".pdf"


with bStrmcoelba
      .type = 1 '//binary
      .open
      .write xHttpcoelba.responseBody
      .savetofile nomearquivocoelba, 2 '//overwrite
end with

dim xHttpcelpe
dim bStrmcelpe
set xHttpcelpe = createobject("Microsoft.XMLHTTP")
Set bStrmcelpe = createobject("Adodb.Stream")

do
icelpe = icelpe + 1
xHttpcelpe.Open "GET", "portalweb", False
xHttpcelpe.Send
Wscript.sleep (5000)
loop until icelpe = 10
nomearquivocelpe =  "d:\Monitoramentofatmail\Celpe\" & "Monitoramento-Celpe" & "-" & data & "-" & tempo & ".pdf"


with bStrmcelpe
      .type = 1 '//binary
      .open
      .write xHttpcelpe.responseBody
      .savetofile nomearquivocelpe, 2 '//overwrite
end with

dim xHttpcosern
dim bStrmcosern
set xHttpcosern = createobject("Microsoft.XMLHTTP")
Set bStrmcosern = createobject("Adodb.Stream")

do
icosern = icosern + 1
xHttpcosern.Open "GET", "portalweb", False
xHttpcosern.Send
Wscript.sleep (5000)
loop until icosern = 10

nomearquivocosern =  "d:\Monitoramentofatmail\Cosern\" & "Monitoramento-Cosern" & "-" & data & "-" & tempo & ".pdf"


with bStrmcosern
      .type = 1 '//binary
      .open
      .write xHttpcosern.responseBody
      .savetofile nomearquivocosern, 2 '//overwrite
end with


set fsocoelba = createobject("Scripting.FileSystemObject") 
set arq = fsocoelba.GetFile(nomearquivocoelba)
if (fsocoelba.FileExists (nomearquivocoelba)) then
	'Wscript.Echo ("O Arquivo Coelba existe")
    'wscript.quit()
        if arq.size = 35008 then
            'msgbox "Arquivo Coelba OK"
            'goto saircoelba
        else
            call emailcoelba
        end if
else
	'Wscript.Echo("Arquivo Não Existe")
	call emailcoelba
	
end if


set fsocelpe = createobject("Scripting.FileSystemObject") 
set arq = fsocelpe.GetFile(nomearquivocelpe)
if (fsocelpe.FileExists (nomearquivocelpe)) then
	'Wscript.Echo ("O Arquivo Celpe existe")
    'wscript.quit()
        if arq.size = 37564 then
            'msgbox "Arquivo Celpe OK"
           'goto saircelpe
        else
            call emailcelpe
        end if
else
	'Wscript.Echo("Arquivo Não Existe")
	call emailcelpe
end if


set fsocosern = createobject("Scripting.FileSystemObject") 
set arq = fsocosern.GetFile(nomearquivocosern)
if (fsocosern.FileExists (nomearquivocosern)) then
	'Wscript.Echo ("O Arquivo Cosern existe")
    'wscript.quit()
        if arq.size = 53160 then
            'msgbox "Arquivo Cosern OK"
            'goto saircosern
        else
            call emailcosern
        end if
else
	'wscript.Echo("Arquivo Não Existe")
	call emailcosern
end if


wscript.quit()

sub emailcoelba()

Set MyEmail=CreateObject("CDO.Message")


Const cdoBasic=0 'Do not Authenticate
Const cdoAuth=1 'Basic Authentication

MyEmail.Subject = "Título do e-mail"
MyEmail.From    = "e-mail do remetente"
MyEmail.To      = "email dos destinatários"
MyEmail.TextBody= "Texto do e-mail"

MyEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendusing")=2

'SMTP Server
MyEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserver")="servidor smtp"

'SMTP Port
MyEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserverport")=25

MyEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 0


'Your UserID on the SMTP server
MyEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendusername") = "email"

'Your password on the SMTP server
MyEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "senha"

'Use SSL for the connection (False or True)
MyEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = False

'MyEmail.AddAttachment nomearquivocoelba

MyEmail.Configuration.Fields.Update
MyEmail.Send

Set MyEmail=nothing

End sub

sub emailcelpe()

Set MyEmail=CreateObject("CDO.Message")


Const cdoBasic=0 'Do not Authenticate
Const cdoAuth=1 'Basic Authentication

MyEmail.Subject = "Título do e-mail"
MyEmail.From    = "e-mail do remetente"
MyEmail.To      = "email dos destinatários"
MyEmail.TextBody= "Texto do e-mail"

MyEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendusing")=2

'SMTP Server
MyEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserver")="servidor smtp"

'SMTP Port
MyEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserverport")=25

MyEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 0


'Your UserID on the SMTP server
MyEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendusername") = "email"

'Your password on the SMTP server
MyEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "senha"

'Use SSL for the connection (False or True)
MyEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = False

'MyEmail.AddAttachment nomearquivocelpe

MyEmail.Configuration.Fields.Update
MyEmail.Send

Set MyEmail=nothing

End sub

sub emailcosern()

Set MyEmail=CreateObject("CDO.Message")


Const cdoBasic=0 'Do not Authenticate
Const cdoAuth=1 'Basic Authentication

MyEmail.Subject = "Título do e-mail"
MyEmail.From    = "e-mail do remetente"
MyEmail.To      = "email dos destinatários"
MyEmail.TextBody= "Texto do e-mail"

MyEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendusing")=2

'SMTP Server
MyEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserver")="servidor smtp"

'SMTP Port
MyEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserverport")=25

MyEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 0


'Your UserID on the SMTP server
MyEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendusername") = "email"

'Your password on the SMTP server
MyEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "senha"

'Use SSL for the connection (False or True)
MyEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = False

'MyEmail.AddAttachment nomearquivocosern

MyEmail.Configuration.Fields.Update
MyEmail.Send

Set MyEmail=nothing

End sub
