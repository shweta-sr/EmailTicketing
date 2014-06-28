Attribute VB_Name = "AutoAttach"
Option Explicit

Sub SendToEach(strFile As String, strSubject As String)
    
    Dim miEmail As MailItem, strBody As String, strHTMLBody As String ', strSubject As String
    Dim xlApp As Excel.Application, wbSource As Excel.Workbook, wsSource As Excel.Worksheet, iCtr As Integer, strEMail As String
    ', strFile As String
        
    'Create Subject string
    'strSubject = "TEST - Tickets for Indus Creations' 'Ulle Veliye'"

    'Create Body string
    strHTMLBody = "<HTML><font size=2 face=Tahoma><p> Your tickets for <font size = 3 color = red>Ulle Veliye</font> and/or coupons for additional services are attached." _
        & " Please print it out and bring it with you on the day of the play.</p> " _
        & "(If you don't see any seat numbers, someone else probably bought your tickets for you. They will either have paper tickets for you or an e-mail like this that includes your seats.)" _
        & "<p><u>The show will start on time. </u></p><p>The Renton IKEA Performing Arts Center" _
        & " is about 30 minutes' drive from the Eastside. Give yourself enough time to get there, pick up tickets if you need to, " _
        & " socialise and find your seats.</p>Theatre address: 400 S Second Street, Renton, WA 98057. </p><p>Bing Maps link: <link>" _
        & "http://binged.it/NC8Ihg</link></p>" _
        & "<p>If you have any questions, please don't hesitate to contact us at" _
        & " tickets@induscreations.org. <p>Thank you for your support - we look forward to seeing you soon! </font></HTML>"
    
    Set xlApp = CreateObject("Excel.Application")
    
    Set wbSource = xlApp.workbooks.Open(strFile)
    Set wsSource = wbSource.worksheets(1)
    
    iCtr = 2
    
    Do
    
        'Create new mail item
        Set miEmail = Application.CreateItem(olMailItem)
                
        'Set Recipients
        strEMail = wsSource.cells(iCtr + 1, 2)
        miEmail.To = strEMail
        
        'Set Subject
        miEmail.Subject = strSubject
        
        'Set Body string
        With miEmail
            .BodyFormat = olFormatHTML
            .HTMLBody = strHTMLBody
        End With
        
        'Attach template
        strFile = wsSource.cells(iCtr + 1, 20)
        miEmail.Attachments.Add (strFile)
            
        'Send
        miEmail.Send
        'miEmail.Save
        'miEmail.Display
        
        iCtr = iCtr + 1
    
    Loop Until wsSource.cells(iCtr + 1, 1) = ""
    
    
EndThisSub:
    wbSource.Close
    xlApp.Application.Quit
    Set wbSource = Nothing
    Set xlApp = Nothing
    Set miEmail = Nothing

End Sub

Sub SendToAll()
    'Call SendToEach("C:\Users\Shweta\Documents\Indus2012\AutoEmail\List_2pm.xlsx", "2 P.M. Tickets for Indus Creations' Ulle Veliye")
    Call SendToEach("C:\Users\Shweta\Documents\Indus2012\AutoEmail\List_6pm.xlsx", "6 P.M. Tickets for Indus Creations' Ulle Veliye")
End Sub
