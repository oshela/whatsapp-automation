Sub WebWhatsapp()
'Activate Selenium Type Library: Tools > References
 Dim bot As New WebDriver
 Dim ks As New Keys
 'init New Chrome Instance & Navigate To webwhatsapp
 bot.start "Chrome", "https://web.Whatsapp.com/"
 bot.Get"/"
 'ask user to scan QR code. once logged in, continue with the macro
 MsgBox "please scan the QR code, after you are logged in, please confirm by clicking"
 'determine the number of messages by identifying the number of last row in the column A
 lastrow = Cells(Rows.Count,1).end(xlup).row
 'search phonenumber/name, press enter, paste text into webwhatsapp,press enter to send message
 For i=2 To lastrow
 'get each text(phonenumber or name)from worksheet
 SearchText = Sheets(1).Range("A" & i).Value
 'click in the searchbox
 bot.findElementByXpath("//* [@id='side']/div[1]/div/Label/div[2].click").click
 'wait 500 ms
 bot.wait(500)
'insert search text (phonenumber or name)
bot.Sendkeys(SearchText)
'wait 500 ms
bot.wait(500)
'press enter to confirm search text
bot.Sendkeys(ks.Enter)
'wait 500 ms
bot.wait(500)
'load message into webwhatsapp
bot.Sendkeys(textmessage)
'wait 500ms
bot.wait(500)
'press enter to send the message
bot.Sendkeys(ks.Enter)
Next i
'get notification once done
MsgBox "done :)"