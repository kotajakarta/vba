Sub cetakallpdftranskip()
Dim a, b, X As Integer
Msg = "Anda akan mencetak Semua Halaman ke file PDF ????"
Style = vbOKCancel
Title = "--== CETAK ALL PDF==--"
jawab = MsgBox(Msg, Style, Title)
If jawab = vbOK Then
Range("a8") = Range("a8").Value
a = Range("a8")
b = Range("a9")
For X = a To b
    Simpankepdftranskip
Range("a8") = X + 1
Next X
MsgBox ("Proses Cetak Selesai !!!")
errorhandler:
Range("a8") = 1
Else
End If
End Sub

---------------------------------------

Function Simpankepdftranskip() As Boolean  ' Copies sheets into new PDF file for e-mailing
    Dim Thissheet As String, ThisFile As String, PathName As String
    Dim SvAs As String

Application.ScreenUpdating = False

' Get File Save Name
    Thissheet = ActiveSheet.Name
    ThisFile = ActiveWorkbook.Name
    PathName = ActiveWorkbook.Path
    SvAs = PathName & "\pdf-ulya\" & Range("a8") & "_" & Range("B8") & "_Transkip Ijazah_Ulya" & ".pdf"

'Set Print Quality
    On Error Resume Next
    ActiveSheet.PageSetup.PrintQuality = 600
    Err.Clear
    On Error GoTo 0

' Instruct user how to send
    On Error GoTo RefLibError
    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, Filename:=SvAs, Quality:=xlQualityStandard, IncludeDocProperties:=False, IgnorePrintAreas:=False, OpenAfterPublish:=False
    On Error GoTo 0
    
SaveOnly:
     
    Simpankepdftranskip = True
    GoTo EndMacro
    
RefLibError:
    MsgBox "Gagal membuat PDF. Lokasi file penyimpanan tidak ditemukan."
    Simpankepdftranskip = False
EndMacro:
End Function
'=============================
Function bFileExists(rsFullPath As String) As Boolean
  bFileExists = CBool(Len(Dir$(rsFullPath)) > 0)
End Function
'=============================