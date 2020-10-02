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



xHttpcoelba.Open "GET", "http://autoatendimento.coelba.com.br/NDP_DCSRUCES_D~home~neologw~sap.com/servlet/login.neoenergia.com.FaturaPorEmail?cc=5374600&f=200008080297&retpage=1", False
xHttpcoelba.Send

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

xHttpcelpe.Open "GET", "http://autoatendimento.celpe.com.br/NDP_DCSRUCES_D~home~neologw~sap.com/servlet/login.neoenergia.com.FaturaPorEmail?cc=7024458193&f=310052092140&retpage=1", False
xHttpcelpe.Send

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

xHttpcosern.Open "GET", "http://autoatendimento.cosern.com.br/NDP_DCSRUCES_D~home~neologw~sap.com/servlet/login.neoenergia.com.FaturaPorEmail?cc=7011116285&f=330078955259&retpage=1", False
xHttpcosern.Send

nomearquivocosern =  "d:\Monitoramentofatmail\Cosern\" & "Monitoramento-Cosern" & "-" & data & "-" & tempo & ".pdf"


with bStrmcosern
      .type = 1 '//binary
      .open
      .write xHttpcosern.responseBody
      .savetofile nomearquivocosern, 2 '//overwrite
end with

set fsocoelba = createobject("Scripting.FileSystemObject") 
if (fsocoelba.FileExists (nomearquivocoelba)) then
	Wscript.Echo ("O Arquivo Coelba existe")
	'wscript.quit()
else
	Script.Echo("Arquivo Não Existe")
end if

set fsocelpe = createobject("Scripting.FileSystemObject") 
if (fsocelpe.FileExists (nomearquivocelpe)) then
	Wscript.Echo ("O Arquivo Celpe existe")
	'wscript.quit()
else
	Script.Echo("Arquivo Não Existe")
end if

set fsocosern = createobject("Scripting.FileSystemObject") 
if (fsocosern.FileExists (nomearquivocosern)) then
	Wscript.Echo ("O Arquivo Cosern existe")
	'wscript.quit()
else
	Script.Echo("Arquivo Não Existe")
end if

Wscript.Echo ("Enviando E-mail")

Set MyEmail=CreateObject("CDO.Message")

Const cdoBasic=0 'Do not Authenticate
Const cdoAuth=1 'Basic Authentication

MyEmail.Subject = "Teste"
MyEmail.From    = "remetente@teste.com"
MyEmail.To      = "destinatario@teste.com"
MyEmail.TextBody= "Monitoramento e/ou Download de Arquivos"

MyEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendusing")=2

'SMTP Server
MyEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserver")="smtp.live.com" 'informar o servidor SMTP

'SMTP Port
MyEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpserverport")=25

MyEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1


'Your UserID on the SMTP server
MyEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendusername") = "seuemail@email.com"

'Your password on the SMTP server
MyEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "senha"

'Use SSL for the connection (False or True)
MyEmail.Configuration.Fields.Item ("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = True

MyEmail.Configuration.Fields.Update
MyEmail.Send

Set MyEmail=nothing

