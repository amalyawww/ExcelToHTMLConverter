Sub ConvertSheetToHTMLFormatWithNBSPAndSaveAsTextFile()
    Dim ws As Worksheet
    Dim htmlContent As String
    Dim lastRow As Long, lastCol As Long
    Dim rowNum As Long, colNum As Long
    Dim filePath As String
    Dim fileNum As Integer
    Dim fileName As String
    
    ' Set worksheet
    Set ws = ThisWorkbook.Sheets("Sheet1") ' Ganti dengan nama sheet yang sesuai
    
    ' Temukan baris dan kolom terakhir dengan data
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    
    ' Mulai membangun string HTML sebagai teks
    htmlContent = "<html><body><table border='1'>" & vbCrLf
    
    ' Loop melalui setiap baris
    For rowNum = 1 To lastRow
        htmlContent = htmlContent & "<tr>" & vbCrLf
        ' Loop melalui setiap kolom
        For colNum = 1 To lastCol
            ' Tambahkan data ke dalam tag <td> di tabel HTML sebagai teks
            htmlContent = htmlContent & "<td>" & Replace(ws.Cells(rowNum, colNum).Text, " ", "&nbsp;") & "</td>" & vbCrLf
        Next colNum
        htmlContent = htmlContent & "</tr>" & vbCrLf
    Next rowNum
    
    ' Akhiri tag HTML
    htmlContent = htmlContent & "</table></body></html>"
    
    ' Tentukan lokasi dan nama file untuk menyimpan teks
    fileName = "SheetContent_" & Format(Now, "YYYYMMDD_HHMMSS") & ".txt"
    filePath = Application.DefaultFilePath & "\" & fileName ' Folder default di komputer
    
    ' Buat file teks
    fileNum = FreeFile
    Open filePath For Output As fileNum
    Print #fileNum, htmlContent
    Close fileNum
    
    ' Tampilkan pesan bahwa file telah didownload
    MsgBox "Konten berhasil dikonversi ke format HTML dan disimpan di: " & filePath
End Sub