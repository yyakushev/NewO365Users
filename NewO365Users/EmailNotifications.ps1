param (
    [string] $sourcefile = '20180119_User_PROD-RESET-PW-V05.csv'
)

$BodyHTML_greetings_en = @"
<html xmlns:v="urn:schemas-microsoft-com:vml"
xmlns:o="urn:schemas-microsoft-com:office:office"
xmlns:w="urn:schemas-microsoft-com:office:word"
xmlns:m="http://schemas.microsoft.com/office/2004/12/omml"
xmlns="http://www.w3.org/TR/REC-html40">

<head>
<meta http-equiv=Content-Type content="text/html; charset=unicode">
<meta name=ProgId content=Word.Document>
<meta name=Generator content="Microsoft Word 15">
<meta name=Originator content="Microsoft Word 15">
<link rel=File-List
href="WG%20Bitte%20um%20Feedback%20D365%20Usermail%20zum%20Livegang_files/filelist.xml">
<link rel=Edit-Time-Data
href="WG%20Bitte%20um%20Feedback%20D365%20Usermail%20zum%20Livegang_files/editdata.mso">
<link rel=themeData
href="WG%20Bitte%20um%20Feedback%20D365%20Usermail%20zum%20Livegang_files/themedata.thmx">
<link rel=colorSchemeMapping
href="WG%20Bitte%20um%20Feedback%20D365%20Usermail%20zum%20Livegang_files/colorschememapping.xml">
<style>
<!--
 /* Font Definitions */
 @font-face
	{font-family:"Cambria Math";
	panose-1:2 4 5 3 5 4 6 3 2 4;
	mso-font-charset:0;
	mso-generic-font-family:roman;
	mso-font-pitch:variable;
	mso-font-signature:-536869121 1107305727 33554432 0 415 0;}
@font-face
	{font-family:Calibri;
	panose-1:2 15 5 2 2 2 4 3 2 4;
	mso-font-charset:0;
	mso-generic-font-family:swiss;
	mso-font-pitch:variable;
	mso-font-signature:-536859905 -1073732485 9 0 511 0;}
@font-face
	{font-family:"Segoe UI";
	panose-1:2 11 5 2 4 2 4 2 2 3;
	mso-font-charset:0;
	mso-generic-font-family:swiss;
	mso-font-pitch:variable;
	mso-font-signature:-469750017 -1073683329 9 0 511 0;}
 /* Style Definitions */
 p.MsoNormal, li.MsoNormal, div.MsoNormal
	{mso-style-unhide:no;
	mso-style-qformat:yes;
	mso-style-parent:"";
	margin:0in;
	margin-bottom:.0001pt;
	mso-pagination:widow-orphan;
	font-size:12.0pt;
	font-family:"Times New Roman",serif;
	mso-fareast-font-family:Calibri;
	mso-fareast-theme-font:minor-latin;}
a:link, span.MsoHyperlink
	{mso-style-priority:99;
	color:blue;
	text-decoration:underline;
	text-underline:single;}
a:visited, span.MsoHyperlinkFollowed
	{mso-style-noshow:yes;
	mso-style-priority:99;
	color:purple;
	text-decoration:underline;
	text-underline:single;}
p
	{mso-style-priority:99;
	mso-margin-top-alt:auto;
	margin-right:0in;
	mso-margin-bottom-alt:auto;
	margin-left:0in;
	mso-pagination:widow-orphan;
	font-size:12.0pt;
	font-family:"Times New Roman",serif;
	mso-fareast-font-family:Calibri;
	mso-fareast-theme-font:minor-latin;}
p.msonormal0, li.msonormal0, div.msonormal0
	{mso-style-name:msonormal;
	mso-style-priority:99;
	mso-style-unhide:no;
	mso-margin-top-alt:auto;
	margin-right:0in;
	mso-margin-bottom-alt:auto;
	margin-left:0in;
	mso-pagination:widow-orphan;
	font-size:12.0pt;
	font-family:"Times New Roman",serif;
	mso-fareast-font-family:Calibri;
	mso-fareast-theme-font:minor-latin;}
span.EmailStyle19
	{mso-style-type:personal;
	mso-style-noshow:yes;
	mso-style-unhide:no;
	font-family:"Segoe UI",sans-serif;
	mso-ascii-font-family:"Segoe UI";
	mso-hansi-font-family:"Segoe UI";
	mso-bidi-font-family:"Segoe UI";
	color:black;}
span.EmailStyle20
	{mso-style-type:personal;
	mso-style-noshow:yes;
	mso-style-unhide:no;
	font-family:"Segoe UI",sans-serif;
	mso-ascii-font-family:"Segoe UI";
	mso-hansi-font-family:"Segoe UI";
	mso-bidi-font-family:"Segoe UI";
	color:black;}
.MsoChpDefault
	{mso-style-type:export-only;
	mso-default-props:yes;
	font-size:10.0pt;
	mso-ansi-font-size:10.0pt;
	mso-bidi-font-size:10.0pt;}
@page WordSection1
	{size:8.5in 11.0in;
	margin:70.85pt 70.85pt 56.7pt 70.85pt;
	mso-header-margin:.5in;
	mso-footer-margin:.5in;
	mso-paper-source:0;}
div.WordSection1
	{page:WordSection1;}
-->
</style>
<!--[if gte mso 10]>
<style>
 /* Style Definitions */
 table.MsoNormalTable
	{mso-style-name:"Table Normal";
	mso-tstyle-rowband-size:0;
	mso-tstyle-colband-size:0;
	mso-style-noshow:yes;
	mso-style-priority:99;
	mso-style-parent:"";
	mso-padding-alt:0in 5.4pt 0in 5.4pt;
	mso-para-margin:0in;
	mso-para-margin-bottom:.0001pt;
	mso-pagination:widow-orphan;
	font-size:10.0pt;
	font-family:"Times New Roman",serif;}
</style>
<![endif]--><!--[if gte mso 9]><xml>
 <o:shapedefaults v:ext="edit" spidmax="1026"/>
</xml><![endif]--><!--[if gte mso 9]><xml>
 <o:shapelayout v:ext="edit">
  <o:idmap v:ext="edit" data="1"/>
 </o:shapelayout></xml><![endif]-->
</head>

<body lang=EN-US link=blue vlink=purple style='tab-interval:.5in'>

<div class=WordSection1>

<p class=MsoNormal><span lang=DE style='font-size:7.5pt;font-family:"Arial",sans-serif;
color:blue;mso-ansi-language:DE'>German version see below</span><span lang=DE
style='mso-ansi-language:DE'> <br>
<br>
</span><b><span lang=DE style='font-size:10.0pt;font-family:"Arial",sans-serif;
mso-ansi-language:DE'>Dear 
"@

$BodyHTML_username_en = @"
</span></b><span lang=DE style='mso-ansi-language:
DE'><br>
<br>
</span><span lang=DE style='font-size:10.0pt;font-family:"Arial",sans-serif;
mso-ansi-language:DE'>From 22 January 2018 - 9:00 a.m. CET (UTC +1) you can
access the new CRM-system.</span><span lang=DE style='mso-ansi-language:DE'> <br>
<br>
</span><b><span lang=DE style='font-size:10.0pt;font-family:"Arial",sans-serif;
mso-ansi-language:DE'>Notice:</span></b><span lang=DE style='mso-ansi-language:
DE'> <o:p></o:p></span></p>

<p><span lang=DE style='font-size:10.0pt;font-family:Symbol;mso-ansi-language:
DE'>· </span><span lang=DE style='font-size:10.0pt;mso-ansi-language:DE'>&nbsp;</span><span
lang=DE style='font-size:10.0pt;font-family:Symbol;mso-ansi-language:DE'> </span><span
lang=DE style='font-size:10.0pt;mso-ansi-language:DE'>&nbsp;</span><span
lang=DE style='font-size:10.0pt;font-family:Symbol;mso-ansi-language:DE'> </span><span
lang=DE style='font-size:10.0pt;mso-ansi-language:DE'>&nbsp;</span><span
lang=DE style='font-size:10.0pt;font-family:Symbol;mso-ansi-language:DE'> </span><span
lang=DE style='font-size:10.0pt;mso-ansi-language:DE'>&nbsp;</span><span
lang=DE style='font-size:10.0pt;font-family:"Arial",sans-serif;mso-ansi-language:
DE'>Temporary passwords are valid for 90 days. The first time you log in, you
will be asked to change your password. </span><span lang=DE style='mso-ansi-language:
DE'><o:p></o:p></span></p>

<p><span lang=DE style='font-size:10.0pt;font-family:Symbol;mso-ansi-language:
DE'>· </span><span lang=DE style='font-size:10.0pt;mso-ansi-language:DE'>&nbsp;</span><span
lang=DE style='font-size:10.0pt;font-family:Symbol;mso-ansi-language:DE'> </span><span
lang=DE style='font-size:10.0pt;mso-ansi-language:DE'>&nbsp;</span><span
lang=DE style='font-size:10.0pt;font-family:Symbol;mso-ansi-language:DE'> </span><span
lang=DE style='font-size:10.0pt;mso-ansi-language:DE'>&nbsp;</span><span
lang=DE style='font-size:10.0pt;font-family:Symbol;mso-ansi-language:DE'> </span><span
lang=DE style='font-size:10.0pt;mso-ansi-language:DE'>&nbsp;</span><span
lang=DE style='font-size:10.0pt;font-family:"Arial",sans-serif;mso-ansi-language:
DE'>Please note that the password you choose complies with the LANXESS password
guidelines (at least 12 digits, at least one capital letter, lower case letter,
special characters and one digit). For more information on how to use LANXESS
IT safely, please visit the following Xnet link: </span><span lang=DE
style='mso-ansi-language:DE'><a
href="http://global.portal.lanxess/intranet/intranet_portal_r6.nsf/id/EN_Safe-use-of-LANXESS-IT"><span
style='font-size:10.0pt;font-family:"Arial",sans-serif'>Safe use of LANXESS IT</span></a>
<o:p></o:p></span></p>

<p><span lang=DE style='mso-ansi-language:DE'><br>
</span><b><span lang=DE style='font-size:10.0pt;font-family:"Arial",sans-serif;
mso-ansi-language:DE'>Access to the new CRM-system: </span></b><span lang=DE
style='mso-ansi-language:DE'><a href="https://lanxess.crm4.dynamics.com/"><b><span
style='font-size:10.0pt;font-family:"Arial",sans-serif'>https://lanxess.crm4.dynamics.com/</span></b></a>
<br>
<br>
</span><b><span lang=DE style='font-size:10.0pt;font-family:"Arial",sans-serif;
mso-ansi-language:DE'>User name: 
"@

$BodyHTML_password_en = @"
</span></b><span lang=DE style='mso-ansi-language:
DE'><br>
</span><b><span lang=DE style='font-size:10.0pt;font-family:"Arial",sans-serif;
mso-ansi-language:DE'>Temporary password: 
"@

$BodyHTML_end_en = @"
</span></b><span lang=DE
style='mso-ansi-language:DE'> <br>
<br>
</span><span lang=DE style='font-size:10.0pt;font-family:"Arial",sans-serif;
mso-ansi-language:DE'>Thank you for your attention.</span><span lang=DE
style='mso-ansi-language:DE'> <br>
<br>
</span><span lang=DE style='font-size:10.0pt;font-family:"Arial",sans-serif;
mso-ansi-language:DE'>With kind regards</span><span lang=DE style='mso-ansi-language:
DE'> <br>
<br>
</span><span lang=DE style='font-size:10.0pt;font-family:"Arial",sans-serif;
mso-ansi-language:DE'>Your CRM Team</span><span lang=DE style='mso-ansi-language:
DE'> <br>
<br>
</span><span lang=DE style='font-size:10.0pt;font-family:"Arial",sans-serif;
mso-ansi-language:DE'>--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------</span><span
lang=DE style='mso-ansi-language:DE'> <br>
<br>
</span>
"@

$BodyHTML_greetings_de = @"
<span lang=DE style='font-size:7.5pt;font-family:"Arial",sans-serif;
color:blue;mso-ansi-language:DE'>German version</span><span lang=DE
style='mso-ansi-language:DE'> <br>
<br>
</span><b><span lang=DE style='font-size:10.0pt;font-family:"Arial",sans-serif;
mso-ansi-language:DE'>Sehr geehrte/r, 
"@

$BodyHTML_username_de = @"
</span></b><span lang=DE style='mso-ansi-language:
DE'> <br>
<br>
</span><span lang=DE style='font-size:10.0pt;font-family:"Arial",sans-serif;
mso-ansi-language:DE'>Ab dem 22. Januar 2018 - 9:00 Uhr MEZ (UTC +1) können Sie
auf das neue CRM-System zugreifen. </span><span lang=DE style='mso-ansi-language:
DE'><br>
<br>
</span><b><span lang=DE style='font-size:10.0pt;font-family:"Arial",sans-serif;
mso-ansi-language:DE'>Hinweis:</span></b><span lang=DE style='mso-ansi-language:
DE'> <o:p></o:p></span></p>

<p><span lang=DE style='font-size:10.0pt;font-family:Symbol;mso-ansi-language:
DE'>· </span><span lang=DE style='font-size:10.0pt;mso-ansi-language:DE'>&nbsp;</span><span
lang=DE style='font-size:10.0pt;font-family:Symbol;mso-ansi-language:DE'> </span><span
lang=DE style='font-size:10.0pt;mso-ansi-language:DE'>&nbsp;</span><span
lang=DE style='font-size:10.0pt;font-family:Symbol;mso-ansi-language:DE'> </span><span
lang=DE style='font-size:10.0pt;mso-ansi-language:DE'>&nbsp;</span><span
lang=DE style='font-size:10.0pt;font-family:Symbol;mso-ansi-language:DE'> </span><span
lang=DE style='font-size:10.0pt;mso-ansi-language:DE'>&nbsp;</span><span
lang=DE style='font-size:10.0pt;font-family:"Arial",sans-serif;mso-ansi-language:
DE'>Temporäre Kennwörter sind 90 Tage gültig. Bei der ersten Anmeldung werden
Sie aufgefordert das Passwort zu ändern. </span><span lang=DE style='mso-ansi-language:
DE'><o:p></o:p></span></p>

<p><span lang=DE style='font-size:10.0pt;font-family:Symbol;mso-ansi-language:
DE'>· </span><span lang=DE style='font-size:10.0pt;mso-ansi-language:DE'>&nbsp;</span><span
lang=DE style='font-size:10.0pt;font-family:Symbol;mso-ansi-language:DE'> </span><span
lang=DE style='font-size:10.0pt;mso-ansi-language:DE'>&nbsp;</span><span
lang=DE style='font-size:10.0pt;font-family:Symbol;mso-ansi-language:DE'> </span><span
lang=DE style='font-size:10.0pt;mso-ansi-language:DE'>&nbsp;</span><span
lang=DE style='font-size:10.0pt;font-family:Symbol;mso-ansi-language:DE'> </span><span
lang=DE style='font-size:10.0pt;mso-ansi-language:DE'>&nbsp;</span><span
lang=DE style='font-size:10.0pt;font-family:"Arial",sans-serif;mso-ansi-language:
DE'>Bitte beachten Sie, dass das von Ihnen gewählte Passwort den LANXESS
Passwort Richtlinien entspricht (mindestens 12-stellig, mindestens je ein
Großbuchstabe, Kleinbuchstabe, Sonderzeichen sowie eine Ziffer). Weitere
Informationen zum sicheren Umgang mit LANXESS IT erhalten Sie unter folgendem
Xnet Link: </span><span lang=DE style='mso-ansi-language:DE'><a
href="http://global.portal.lanxess/intranet/intranet_portal_r6.nsf/id/8a271602c80dbb81c1258139005aa1d3"><span
style='font-size:10.0pt;font-family:"Arial",sans-serif'>Sicherer Umgang mit
LANXESS IT</span></a> <o:p></o:p></span></p>

<p><span lang=DE style='mso-ansi-language:DE'><br>
</span><b><span lang=DE style='font-size:10.0pt;font-family:"Arial",sans-serif;
mso-ansi-language:DE'>Zugang zum neuen CRM-System: </span></b><span lang=DE
style='mso-ansi-language:DE'><a href="https://lanxess.crm4.dynamics.com/"><b><span
style='font-size:10.0pt;font-family:"Arial",sans-serif'>https://lanxess.crm4.dynamics.com/</span></b></a>
<br>
<br>
</span><b><span lang=DE style='font-size:10.0pt;font-family:"Arial",sans-serif;
mso-ansi-language:DE'>Benutzername: 
"@

$BodyHTML_password_de = @"
</span></b><span lang=DE style='mso-ansi-language:
DE'> <br>
</span><b><span lang=DE style='font-size:10.0pt;font-family:"Arial",sans-serif;
mso-ansi-language:DE'>Temporäres Passwort: 
"@

$BodyHTML_end_de = @"
</span></b><span lang=DE
style='mso-ansi-language:DE'><br>
<br>
</span><span lang=DE style='font-size:10.0pt;font-family:"Arial",sans-serif;
mso-ansi-language:DE'>Vielen Dank für Ihre Aufmerksamkeit. </span><span
lang=DE style='mso-ansi-language:DE'><br>
<br>
</span><span lang=DE style='font-size:10.0pt;font-family:"Arial",sans-serif;
mso-ansi-language:DE'>Mit freundlichen Grüßen,</span><span lang=DE
style='mso-ansi-language:DE'> <br>
<br>
</span><span lang=DE style='font-size:10.0pt;font-family:"Arial",sans-serif;
mso-ansi-language:DE'>Ihr CRM Team</span><span lang=DE style='mso-ansi-language:
DE'> <o:p></o:p></span></p>

<p class=MsoNormal><span lang=DE style='mso-ansi-language:DE'><o:p>&nbsp;</o:p></span></p>

</div>

</body>

</html>
"@


$logfile = "ResetEmailProcessing.log"
$from = 'crm-lxs-support@itvt.de'
$smtp = '127.0.0.1:25'
$cred = Get-Credential
$subject = "New CRM-System - User Login Information"

$enc = [System.Text.Encoding]::UTF8

$users = import-csv $sourcefile -Delimiter ','

$users |ForEach-Object {  
    
        $msgHTML = $BodyHTML_greetings_en+$_.fullname`
                +$BodyHTML_username_en+$_.UPN`
                +$BodyHTML_password_en+$_.NewPassword`
                +$BodyHTML_end_en`
                +$BodyHTML_greetings_de+$_.fullname`
                +$BodyHTML_username_de+$_.UPN`
                +$BodyHTML_password_de+$_.NewPassword`
                +$BodyHTML_end_de
        try {
            Send-MailMessage    -SmtpServer '127.0.0.1' `
                                -BodyAsHtml $msgHTML `
                                -From $from `
                                -to $_.PrimaryEmail `
                                -Subject $subject `
                                -Encoding $enc
            "$(Get-date -Format "dd.MM.yy_HH.mm.ss"): Email was sent to user $($_.UPN)" | Out-File $logfile -Append -Encoding unicode
    
        } catch {

            "$(Get-date -Format "dd.MM.yy_HH.mm.ss"): ERROR: Email wasn't sent to user $($_.UPN)" | Out-File $logfile -Append -Encoding unicode
                            
        }
}


 
