Attribute VB_Name = "Module5"
Function Host(REF As String) As String
    If InStr(1, REF, "instagram", vbTextCompare) > 0 Then
        Host = "Instagram"
    ElseIf InStr(1, REF, "facebook", vbTextCompare) > 0 Then
        Host = "Facebook"
    ElseIf InStr(1, REF, "google", vbTextCompare) > 0 Or InStr(1, REF, "banner", vbTextCompare) > 0 Then
        Host = "Google"
    ElseIf InStr(1, REF, "youtube", vbTextCompare) > 0 Then
        Host = "YouTube"
    Else
        Host = "Unknown"
    End If
End Function

