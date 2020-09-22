<div align="center">

## Simple Mail Transfer Protocol


</div>

### Description

This code will teach you Step-by-Step how to make a simple emailer. I would recommend this if you are looking to make an emailer, anonymailer, or Ezine emailer.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Corey langdon \(has since resigned\)](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/corey-langdon-has-since-resigned.md)
**Level**          |Beginner
**User Rating**    |3.1 (34 globes from 11 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Coding Standards](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/coding-standards__1-43.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/corey-langdon-has-since-resigned-simple-mail-transfer-protocol__1-37957/archive/master.zip)





### Source Code

This code is a little hard to understand but by going over it a few timew you can EASILY understand SMTP.<p>
<b>First we need to connect to an SMTP server like this:</b><p>
Winsock1.Connect "SMTP server", "25"<p>
<b>Now that we have connected we need to start the transfer:</b><p>
Winsock1.Senddata "HELO " + Winsock1.LocalIP + vbCrLf<p>
<b>That starts the transfer. Then You need to Specify who is sending the email:</b><p>
Winsock1.SendData "MAIL FROM:<" + "MailFrom@blah.com" + " > " + vbCrLf<p>
<b>That will say who it is from. You then need to say who you want to send it to:</b><p>
Winsock1.SendData "RCPT TO:<" + "Mailto@blah.com" + " > " + vbCrLf<p>
<b>You have specified who to send it to, now you have to tell the server that you are ready to send the contents of the email:</b><p>
Winsock1.SendData "DATA" + vbCrLf<p>
<b>These are the contents of the email:</b><p>
 Winsock1.SendData "From: <" + "Mailfrom@blah.com" + ">" + vbCrLf + _<p>
 "To: " + "Mailto@blah.com" + vbCrLf + _<p>
 "Subject: " + "Subject" + vbCrLf <p>
<b>When you have done that you are ready to send the Message and Other data:</b><p>
"Mime-Version: 1.0" + vbCrLf + _<p>
 "Content-Type: text/html" + vbTab + "charset=us-ascii" + vbCrLf + vbCrLf & "Message of Email"<p>
 Winsock1.SendData vbCrLf + "." + vbCrLf<p>
<b>Now we need to tell the server that we are done sending the email:</b><p>
Winsock1.SendData "QUIT"<p>
<b>Please Vote for this bcuz i worked on this for awhile and got kicked offline while doing this. I will accept any critisizm as long as it is appropriate and not bagging on me. I would like it if i could get some votes. Also if you copy and paste this and go over it for awhile then you will have mastered this. It only takes 10 minutes. A couple of suggestions if you are going to make this an actual emailer:<p>
<br>
<b>-Change the "Mailto@blah.com" and "Mailfrom@blah.com" to a text box appropriatley fitting them so you can change the Mailer and Reciever.</b><p>
<b>-Also Put DoEvents: Doevents: Doevents: doevents under The winsock1.Connect "SMTP SERVER", "2"</b><p>
<b>-Change the "SMTP SERVER" to a combobox or textbox to change servers.</b><p>
<b>-Email me if you have any Questions.</b><p>
<b><u><i>P.S. I would be in your debt for Five globes :) !</i></u></b>

