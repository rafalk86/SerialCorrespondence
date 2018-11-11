Sub Sending()

Dim filePath As String
Dim lastSlot As Integer
Dim addedSlots As Integer
Dim signature
Dim fso As Object
Dim txs As Object
Dim footer
Dim info
Dim answer

answer = MsgBox("Czy na pewno chcesz wys³aæ wiadomoœci?", vbYesNo + vbExclamation, "Potwierdzenie wys³ania")

    If answer = 6 Then

        filePath = ThisWorkbook.Path
        signature = filePath & "\" & "Coral_signature" & ".htm"

        lastSlot = ActiveSheet.Range("A" & ActiveSheet.Rows.Count).End(xlUp).Row
        For I = lastSlot To 2 Step -1
    
            If Cells(I, "C").Value <> "" Then
                
                If Dir(filePath & "\" & Cells(I, "C").Value & ".pdf") <> "" Then
                
                    Cells(I, "D") = filePath & "\" & Cells(I, "C").Value & ".pdf"
                    
                End If
            
            End If
            
            Cells(I, "B").Value = Replace(Cells(I, "B").Value, ",", ";")
            Cells(I, "B").Value = Replace(Cells(I, "B").Value, " ", "")
          
        Next
    
        addedSlots = Worksheets("Sending List").Range("D:D").Cells.SpecialCells(xlCellTypeConstants).Count
    
        If lastSlot = addedSlots Then
    
            Set fso = CreateObject("Scripting.FileSystemObject")
            Set txs = fso.GetFile(signature).OpenAsTextStream

            footer = txs.readall
    
            For I = lastSlot To 2 Step -1
    
                Set Mail = CreateObject("outlook.application")
                Set MyMessage = Mail.CreateItem(0)
        
                With MyMessage
                .To = Cells(I, "B").Value
                .CC = Cells(21, "E").Value + ";" + Cells(22, "E").Value
                .Subject = Cells(1, "E").Value & " ACC " & Cells(I, "C").Value
                .ReadReceiptRequested = False
                .OriginatorDeliveryReportRequested = False
                .htmlBody = "<font face = calibri><font size = 3>" & Cells(3, "E").Value & "<br>" & "<br>" & Cells(4, "E").Value & "</size></font>" & "<br>" & "<br>" & "<font face = arial><font size = 2>" & Cells(12, "E").Value & "<br>" & Cells(13, "E").Value & "<br>" & Cells(14, "E").Value & "<br>" & Cells(15, "E").Value & "<br>" & Cells(16, "E").Value & "<br>" & Cells(17, "E").Value & "<br>" & Cells(18, "E").Value & "<br>" & Cells(19, "E").Value & "<br>" & Cells(20, "E").Value & "<br>" & "</size></font>" & footer
                .Attachments.Add (Cells(I, "D").Value)
                .Send
                End With

                Set Mail = Nothing
                Set MyMessage = Nothing

            Next
            
            Range(Cells(2, "D"), Cells(lastSlot, "D")) = ""
            
            info = MsgBox("Wys³ano poprawnie. Liczba wiadomoœci " & lastSlot - 1 & ".", vbInformation, "Braki w plikach")
            
        Else
        'msg o braku plików
    
            info = MsgBox("Nie wszyscy kontrahenci maj¹ przypisane pliki z potwierdzeniem sald.", vbCritical, "Braki w plikach")
    
        End If

    Else
        
        'nie wysy³a wiadomoœci

    End If

End Sub

