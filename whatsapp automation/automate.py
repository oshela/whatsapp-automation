'---------------------------------------------------------------------------------------
' Module     : mWhatsAppBOT
' Author     : oshela

' Email      : abimikuoshela@gmail.com
' Purpose    : Send (multiple) text messages, media and documents via WebWhatsApp
' Dependency : VBA Selenium Type Library | https://github.com/florentbr/SeleniumBasic
'---------------------------------------------------------------------------------------

Option Explicit

Public BOT As Selenium.WebDriver, By As Selenium.By, ks As Selenium.Keys
Public wb As Workbook
Public ws As Worksheet
Public i As Long

Public Sub WhatsAppBOT()

10        On Error GoTo ErrHandler
20        Application.ScreenUpdating = True
          
          Dim MessageValue As String
          Dim KeepLoginCredentials As String
          Dim MessageType As String
          Dim ImageCaption As String
          Dim SearchText As String
          Dim DefaultDelay As Long
          Dim RandomDelay As Long
          Dim LastRowStatus As Long
          Dim LastRowNumber As Long
          Dim LastRowText As Long
          Dim IsValidContact As Boolean

30        Set wb = ThisWorkbook
40        Set ws = wb.Worksheets("BOT")
          
50        KeepLoginCredentials = ws.Range("KeepLoginCredentials")

60        Set BOT = InitWebDriver(KeepLoginCredentials)
70        Set By = New Selenium.By
80        Set ks = New Selenium.Keys
          

          ''' Determine last rows
90        LastRowStatus = ws.Cells(Rows.Count, BotColumn.wcStatus).End(xlUp).Row
100       LastRowNumber = ws.Cells(Rows.Count, BotColumn.wcNumber).End(xlUp).Row
110       LastRowText = ws.Cells(Rows.Count, BotColumn.wcText).End(xlUp).Row

          ''' Clear status cells
120       If LastRowStatus > FirstRow - 1 Then
130           ws.Range(ws.Cells(FirstRow, BotColumn.wcStatus), _
                  ws.Cells(LastRowStatus, BotColumn.wcStatus)).ClearContents
140       End If

          ''' Init New Chrome instance & navigate to WebWhatsApp
150       BOT.Start "chrome"
160       BOT.Get "https://web.whatsapp.com/"

          ''' Ask user to scan the QR code. Once logged in, continue with the macro
170       MsgBox _
              "Please scan the QR code." & _
              "After you are logged in, please confirm this message box by clicking 'ok'", vbOKOnly, "WhatsApp Bot"

180       DefaultDelay = ws.Range("DefaultDelay").Value * 1000
        
          ''' Go to each link, paste text into WebWhatsApp and press enter to send the message
190       For i = FirstRow To LastRowNumber

200           MessageValue = ws.Cells(i, BotColumn.wcText).Value
210           MessageType = ws.Cells(i, BotColumn.wcType).Value
220           ImageCaption = ws.Cells(i, BotColumn.wcCaption).Value
230           SearchText = ws.Cells(i, BotColumn.wcNumber).Value
              
240           If SearchText = vbNullString Then
250               ws.Cells(i, BotColumn.wcStatus).Value = _
                      "Error: No number provided | " & Format(Now, "mm/dd/yyyy HH:mm:ss")
260               GoTo NextIteration
270           End If

280           RandomDelay = CalcRandomDelay()

290           IsValidContact = IsValidContactSavedNumber(i)
300           If IsValidContact = False Then GoTo NextIteration
              
310           BOT.FindElementByXPath(xPathSearchInputField).WaitDisplayed (True)
              
320           If MessageType = "Text" Then
330               Call SendTextmessage(MessageValue)
340           ElseIf MessageType = "Media" Or MessageType = "Document" Then
350               Call SendAttachment(MessageType, MessageValue, ImageCaption, DefaultDelay)
360           End If
              
370           BOT.Wait (DefaultDelay)
380           BOT.Wait (RandomDelay)
          
NextIteration:
390       Next i

400       BOT.Quit
410       MsgBox "Task completed!", vbOKOnly, "WhatsApp BOT"

EndIt:
420       If Not BOT Is Nothing Then BOT.Quit
430       Exit Sub
          
ErrHandler:
440       DisplayError Err.Source, Err.Description, "mWhatsAppBOT.WhatsAppBOT", Erl
450       Resume EndIt
          
End Sub


Private Function IsValidContactSavedNumber(ByVal i As Long) As Boolean

460       On Error GoTo ErrHandler
          ''' Insert searchtext (phone number / or name) in searchinput field
          Dim SearchText As String
470       SearchText = ws.Cells(i, BotColumn.wcNumber)
480       BOT.FindElementByXPath(xPathSearchInputField).WaitDisplayed(True).SendKeys (SearchText)
490       BOT.Wait (200)
500       BOT.SendKeys (ks.Enter)
510       BOT.Wait (200)
          
          ''' Check, if contact exists
520       If BOT.IsElementPresent(By.XPath(xPathNoContactFound)) Then
530           BOT.FindElementByXPath(xPathSearchInputField).Clear
540           ws.Cells(i, BotColumn.wcStatus).Value = _
                  "Error: No contact found | " & Format(Now, "mm/dd/yyyy HH:mm:ss")
550           IsValidContactSavedNumber = False
560       Else
570           IsValidContactSavedNumber = True
580       End If
          
590       Exit Function
          
ErrHandler:
600       RaiseError Err.Number, Err.Source, "mWhatsAppBOT.IsValidContactSavedNumber", Err.Description, Erl
End Function

'---------------------------------------------------------------------------------------
' Procedure : CalcRandomDelay
' Purpose   : Calculate Random Number using Application.WorksheetFunction.RandBetween
'---------------------------------------------------------------------------------------
'
Private Function CalcRandomDelay() As Long

610       On Error GoTo ErrHandler
          Dim minRandomDelay As Long
          Dim maxRandomDelay As Long
          Dim RandomDelay As Long
          Dim useRandomDelay As String
          
620       useRandomDelay = ws.Range("useRandomDelay").Value
630       If useRandomDelay = "Yes" Then
640           minRandomDelay = ws.Range("minRandomDelay").Value * 1000
650           maxRandomDelay = ws.Range("maxRandomDelay").Value * 1000
660           RandomDelay = Application.WorksheetFunction.RandBetween(minRandomDelay, maxRandomDelay)
670       Else
680           RandomDelay = 0
690       End If
          
700       CalcRandomDelay = RandomDelay
710       Exit Function
          
ErrHandler:
720       RaiseError Err.Number, Err.Source, "mWhatsAppBOT.CalcRandomDelay", Err.Description, Erl

End Function

'---------------------------------------------------------------------------------------
' Procedure : InitWebDriver
' Purpose   : Initial Chrome WebDriver. All settings are done in mGlobals module
'             Settings will overwrite the default timeout times
'---------------------------------------------------------------------------------------
'
Private Function InitWebDriver(KeepLoginCredentials As String) As WebDriver

730       On Error GoTo ErrHandler
          Dim UserDataDir As String
740       Set BOT = New Selenium.WebDriver
          
          ''' Set timeouts defined in mGlobals
750       BOT.Timeouts.ImplicitWait = ImplicitWait
760       BOT.Timeouts.PageLoad = PageLoad
770       BOT.Timeouts.Server = TimeoutServer

          ''' Add additional arguments
780       BOT.AddArgument "--disable-popup-blocking"
790       BOT.AddArgument "--disable-notifications"

          ''' Avoid Scanning the QR Code
800       If KeepLoginCredentials = "Yes" Then
810           UserDataDir = ThisWorkbook.Path & "\" & "ChromeUserData"
820           Call CreateDirectory(UserDataDir)
830           BOT.AddArgument "--user-data-dir=" & UserDataDir
840       End If
          
850       Set InitWebDriver = BOT
860       Exit Function
          
ErrHandler:
870       Select Case Err.Number
              ''' Could be because .Net Framework is not installed
              Case -2146232576
880               MsgBox _
                      "Oh, it looks like that you first need to activate/install your .NET Framework." & _
                      "But no worries! To fix this issue, follow the steps here:" & vbNewLine & _
                      "https://pythonandvba.com/automation-error" & vbNewLine & vbNewLine & _
                      Err.Number & ": " & Err.Description, , "WhatsApp Bot"
890               End
900           Case Else
910               RaiseError Err.Number, Err.Source, "mWhatsAppBOT.InitWebDriver", Err.Description, Erl
920       End Select
          
End Function

'---------------------------------------------------------------------------------------
' Procedure : SendTextmessage
' Purpose   : Send 'normal' text message. Errors will be returned in the respective row
'---------------------------------------------------------------------------------------
'
Private Sub SendTextmessage(MessageValue As String)

930       On Error GoTo ErrHandler
          Dim arrTextMessage As Variant
          Dim LenOfArray As Integer
          Dim Line As Long

940       On Error GoTo ErrHandler
          ''' Split text message based on "|" to identify new paragraph
950       arrTextMessage = Split(MessageValue, "|")
          
          '''' Length of variable. If only one line, it returns 1
960       LenOfArray = UBound(arrTextMessage) - LBound(arrTextMessage) + 1
          
          ''' Iterate over array and press Shift + Enter to create mew paragraph
970       For Line = LBound(arrTextMessage) To UBound(arrTextMessage)
              'Cursor should be aleady in the text input field
980           BOT.Wait (200)
990           BOT.SendKeys (arrTextMessage(Line))
1000          BOT.Wait (500)
1010          If LenOfArray > 1 Then
                  ''' Create a new line by pressing Shift & Enter
1020              BOT.Keyboard.KeyDown (ks.Shift)
1030              BOT.SendKeys (ks.Enter)
1040              BOT.Keyboard.KeyUp (ks.Shift)
1050              BOT.Wait (200)
1060          End If
1070      Next Line
1080      BOT.Wait (200)
1090      BOT.SendKeys (ks.Enter)
1100      ws.Cells(i, BotColumn.wcStatus).Value = "Sent: " & Format(Now, "mm/dd/yyyy HH:mm:ss")
1110      Exit Sub
          
ErrHandler:
1120      ws.Cells(i, BotColumn.wcStatus).Value = _
              "Error: " & Err.Number & "_" & Err.Description & ", " & Format(Now, "mm/dd/yyyy HH:mm:ss")

End Sub

'---------------------------------------------------------------------------------------
' Procedure : SendAttachment
' Purpose   : Send Media or Document attachment (instead of text message)
'             Errors will be returned in the respective row
'---------------------------------------------------------------------------------------
'
Private Sub SendAttachment(MessageType As String, MessageValue As String, ImageCaption As String, DefaultDelay As Long)
          
1130      On Error GoTo ErrHandler
          
          Dim FileSizeMB As Single
          Dim arrImageCaption As Variant
          Dim LenOfArray As Integer
          Dim Line As Long
          Dim fso As Object

          
1140      MessageValue = Replace(MessageValue, """", "")
          
          ''' Before trying to send anything, check if file exists
1150      Set fso = CreateObject("Scripting.FileSystemObject")
1160      If fso.FileExists(MessageValue) = False Then
1170          ws.Cells(i, BotColumn.wcStatus).Value = _
                  "Error: File does not exist | " & Format(Now, "mm/dd/yyyy HH:mm:ss")
1180          GoTo NextIteration
1190      End If

1200      BOT.FindElement(By.XPath(AttachmentButton)).WaitDisplayed(True).Click
1210      ws.Cells(i, BotColumn.wcText).Copy
          
1220      If MessageType = "Media" Then
1230          BOT.FindElement(By.XPath(AttachmentImage)).WaitDisplayed(True).Click
1240      ElseIf MessageType = "Document" Then
1250          BOT.FindElement(By.XPath(AttachmentDocument)).WaitDisplayed(True).Click
1260      End If
          
          ''' Paste filelink via Application (VBA)
1270      BOT.Wait (2000)
1280      Application.SendKeys (MessageValue)
1290      BOT.Wait (2000)
1300      Application.SendKeys ("{Enter}")
1310      BOT.Wait (DefaultDelay)
          
1320      If MessageType = "Media" Then
1330          If Not ImageCaption = vbNullString Then
                  ''' Split Imagecaption based on "|" to identify new paragraph
1340              arrImageCaption = Split(ImageCaption, "|")
          
                  ''' Length of variable. If only one line, it returns 1
1350              LenOfArray = UBound(arrImageCaption) - LBound(arrImageCaption) + 1
                  
                  ''' Loop over array and press Shift + Enter to create new paragraph
1360              For Line = LBound(arrImageCaption) To UBound(arrImageCaption)
                      ''' Insert caption. Cursor is already in 'Caption Field'
1370                  BOT.Wait (DefaultDelay)
1380                  BOT.SendKeys (arrImageCaption(Line))
1390                  BOT.Wait (200)
                       
1400                  If LenOfArray > 1 Then
                          ''' Create a new line by pressing Shift & Enter
1410                      BOT.Keyboard.KeyDown (ks.Shift)
1420                      BOT.SendKeys (ks.Enter)
1430                      BOT.Keyboard.KeyUp (ks.Shift)
1440                      BOT.Wait (200)
1450                  End If
1460              Next Line
1470          End If
1480      End If
          ''' Send attachment by pressing ENTER
1490      BOT.SendKeys (ks.Enter)
          ''' Increase Delay Time depending on filesize.
          '''' E.g. Filesize of 1 MB -> Increase of Delay time 2000 ms (1 * 2000)
1500      FileSizeMB = Application.WorksheetFunction.RoundUp(FileLen(MessageValue) / 1000000, 0)
1510      BOT.Wait (DefaultDelay + (FileSizeMB * DelayTimeAttachment))
1520      ws.Cells(i, BotColumn.wcStatus).Value = "Sent: " & Format(Now, "mm/dd/yyyy HH:mm:ss")

NextIteration:
1530      Set fso = Nothing
1540      Exit Sub
          
ErrHandler:
1550      ws.Cells(i, BotColumn.wcStatus).Value = _
              "Error: " & Err.Number & "_" & Err.Description & ", " & Format(Now, "mm/dd/yyyy HH:mm:ss")

End Sub

' ------------------------------------------------------
' Name:      CreateDirectory
' Purpose:   Create Directory for any given path (if it does not already exist)
' ------------------------------------------------------
Private Sub CreateDirectory(FolderPath)
          
1560      On Error GoTo ErrHandler
          
          'Using FileSystemObject raised sometimes an error:
          'Bad file name or number, Error code: 52
          'Therefore, I am using the following alternative

1570      If Dir(FolderPath, vbDirectory) = "" Then MkDir FolderPath
          
1580      Exit Sub
          
ErrHandler:
1590      RaiseError Err.Number, Err.Source, "mWhatsAppBOT.CreateDirectory", Err.Description, Erl

End Sub
