Attribute VB_Name = "WilshireEmail"
Sub SendOneEmailToAll()
    
    Dim miMgrEmail As MailItem, strSubject As String, strBody As String, strHTMLBody As String
    Dim dtLBD As Date, dt3BD As Date, strMonth As String, dtToday As Date
    Dim rMe As Recipient, iPastDays As Integer, iAddDays As Integer, iWDays As Integer
    
    'Calculate the dates: last business day of this month; third business day of next month; this month's long name
    dtToday = Date
    
    iPastDays = DatePart("d", dtToday)
    
    If iPastDays < 15 Then
        dtLBD = DateAdd("d", -1 * iPastDays, dtToday)
    Else
        dtLBD = DateAdd("d", -1 * iPastDays + 1, dtToday)
        dtLBD = DateAdd("m", 1, dtLBD)
        dtLBD = DateAdd("d", -1, dtLBD)
    End If
    
    Do Until IsBD(dtLBD)
        dtLBD = DateAdd("d", -1, dtLBD)
    Loop
    
    iAddDays = 0
    dt3BD = dtLBD
    Do
        dt3BD = DateAdd("d", 1, dt3BD)
        If IsBD(dt3BD) Then iAddDays = iAddDays + 1
    Loop While iAddDays < 3
    strMonth = MonthName(DatePart("m", dtLBD))
    
    'Create Subject string
    strSubject = "Russell Investments - " & strMonth & " Fixed Income Data Request"

    'Create Body string
    strHTMLBody = "<HTML><p><font size=2 face=Tahoma> Just a reminder that Russell Investments is expecting to receive your data for the fixed income portfolios your firm manages on our " _
    & "behalf in our prescribed template format as of " & Format(dtLBD, "mmmm dd, yyyy") & " <font color = red><u>no later than 4pm PDT (GMT-8) on " & WeekdayName(Weekday(dt3BD)) & ", " _
    & Format(dt3BD, "mmmm dd, yyyy") & "</font></u>.</font></p> <p><font size=2 face=Tahoma>Please make sure all dates within the " _
    & "template are in the specified format of yyyymmdd.  The date on the filename will stay the same as previous months. </font></p> <p><font size=2 face=Tahoma>As with previous months, " _
    & "templates should be sent to the FTP location.  Please let me know if you have any questions.</font></p>" _
    & "<p><font size = 2, face = Garamond><b>Shweta Sundararajan</b><br>Russell Investments | Enterprise Data Management<br>+1 206 505 3796</font></p></HTML>"
    
    'Create new mail item
    Set miMgrEmail = Application.CreateItem(olMailItem)
    
    'Change From: to the Fixed Income Analytics mailbox
    miMgrEmail.SentOnBehalfOfName = "Fixed Income Analytics (Mailbox)"
    
    'Change Recipients object to include all managers in BCC and yourself, Carrie in To
    Set rMe = miMgrEmail.Recipients.Add("Fixed Income Managers")
    rMe.Type = olBCC
    
    'Request read receipt
    miMgrEmail.ReadReceiptRequested = True
    
    'Set Subject
    miMgrEmail.Subject = strSubject
    
    'Set Body string
    With miMgrEmail
        .BodyFormat = olFormatHTML
        .HTMLBody = strHTMLBody
    End With
    
    'Attach template
    miMgrEmail.Attachments.Add ("I:\_Wilshire Procedurs Docs\Templates\20110930 Revised Manager Spreadsheet Template.xls")
    
    'Send
    miMgrEmail.Display
    
EndThisSub:
    Set rMe = Nothing
    Set miMgrEmail = Nothing
    
End Sub

Private Function IsBD(dt As Date) As Boolean
    Dim iWDay As Integer, vHolDates As Variant
    
    IsBD = True
    
    vHolDates = Array("1/2/2012", "1/16/2012", "2/20/2012", "5/28/2012", "7/4/2012", "9/3/2012", "11/22/2012", "11/23/2012", "12/25/2012")
    
    iWDay = Weekday(dt)
    If iWDay = 1 Or iWDay = 7 Then
        IsBD = False
    Else
        For iWDay = 0 To UBound(vHolDates)
            If dt = vHolDates(iWDay) Then
                IsBD = False
            End If
        Next iWDay
    End If
    
End Function

Sub SendToEach()
    
    Dim miMgrEmail As MailItem, strSubject As String, strBody As String, strHTMLBody As String
    Dim dtLBD As Date, dt3BD As Date, strMonth As String, dtToday As Date
    Dim iPastDays As Integer, iAddDays As Integer, iWDays As Integer
    Dim xlApp As Object ', wbSource As Workbook, wsSource As Worksheet, iCtr As Integer, strEMail As String
    
    'Calculate the dates: last business day of this month; third business day of next month; this month's long name
    dtToday = Date
    
    iPastDays = DatePart("d", dtToday)
    
    If iPastDays < 15 Then
        dtLBD = DateAdd("d", -1 * iPastDays, dtToday)
    Else
        dtLBD = DateAdd("d", -1 * iPastDays + 1, dtToday)
        dtLBD = DateAdd("m", 1, dtLBD)
        dtLBD = DateAdd("d", -1, dtLBD)
    End If
    
    Do Until IsBD(dtLBD)
        dtLBD = DateAdd("d", -1, dtLBD)
    Loop
    
    iAddDays = 0
    dt3BD = dtLBD
    Do
        dt3BD = DateAdd("d", 1, dt3BD)
        If IsBD(dt3BD) Then iAddDays = iAddDays + 1
    Loop While iAddDays < 3
    strMonth = MonthName(DatePart("m", dtLBD))
    
    'Create Subject string
    strSubject = "Russell Investments - " & strMonth & " Fixed Income Data Request"

    'Create Body string
    strHTMLBody = "<HTML><p><font size=2 face=Tahoma> Just a reminder that Russell Investments is expecting to receive your data for the fixed income portfolios your firm manages on our " _
    & "behalf in our prescribed template format as of " & Format(dtLBD, "mmmm dd, yyyy") & " <font color = red><u>no later than 4pm PDT (GMT-8) on " & WeekdayName(Weekday(dt3BD)) & ", " _
    & Format(dt3BD, "mmmm dd, yyyy") & "</font></u>.</font></p> <p><font size=2 face=Tahoma>Please make sure all dates within the " _
    & "template are in the specified format of yyyymmdd.  The date on the filename will stay the same as previous months. </font></p> <p><font size=2 face=Tahoma>As with previous months, " _
    & "templates should be sent to the FTP location.  Please let me know if you have any questions.</font></p>" _
    & "<p><font size = 2, face = Garamond><b>Shweta Sundararajan</b><br>Russell Investments | Enterprise Data Management<br>+1 206 505 3796</font></p></HTML>"
    
    Set xlApp = CreateObject("Excel.Application")
    
    Set wbSource = xlApp.workbooks.Open("\\frctc_fs-a\users$\SSundara\My Documents\Work\Wilshire\Info\ContactSpdsheet.xlsx")
    Set wsSource = wbSource.worksheets(1)
    
    iCtr = 1
    
    Do
    
        'Create new mail item
        Set miMgrEmail = Application.CreateItem(olMailItem)
        
        'Change From: to the Fixed Income Analytics mailbox
        miMgrEmail.SentOnBehalfOfName = "Fixed Income Analytics (Mailbox)"
        
        'Change Recipients object to include all managers in BCC and yourself, Carrie in To
        strEMail = wsSource.cells(iCtr + 1, 1)
        miMgrEmail.BCC = strEMail
        miMgrEmail.CC = "Luetzow, Carrie; Cavallari, Lisa"
        
        'Request read receipt
        miMgrEmail.ReadReceiptRequested = True
        
        'Set Subject
        miMgrEmail.Subject = strSubject
        
        'Set Body string
        With miMgrEmail
            .BodyFormat = olFormatHTML
            .HTMLBody = strHTMLBody
        End With
        
        Do
            'Attach template
            miMgrEmail.Attachments.Add ("\\frctc_fs-a\users$\SSundara\My Documents\Work\Wilshire\mmddyyyy\" & wsSource.cells(iCtr + 1, 2) & ".xls")
            iCtr = iCtr + 1
        Loop Until wsSource.cells(iCtr + 1, 1) <> strEMail
            
        'Send
        miMgrEmail.Display
    
    Loop Until wsSource.cells(iCtr + 1, 1) = ""
    
    
EndThisSub:
    wbSource.Close
    Set miMgrEmail = Nothing

End Sub
