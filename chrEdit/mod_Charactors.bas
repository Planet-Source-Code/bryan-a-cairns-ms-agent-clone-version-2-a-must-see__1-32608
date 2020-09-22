Attribute VB_Name = "mod_Charactors"
Private Declare Function CoCreateGuid Lib "ole32.dll" (pGUID As Any) As Long


Public Function CreateGUID() As String
    Dim i As Long, b(0 To 15) As Byte


    If CoCreateGuid(b(0)) = 0 Then


        For i = 0 To 15
            CreateGUID = CreateGUID & Right$("00" & Hex$(b(i)), 2)
        Next i
    Else
        MsgBox "Error While creating GUID!"
    End If
End Function

