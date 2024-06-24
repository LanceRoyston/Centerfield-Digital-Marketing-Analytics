Attribute VB_Name = "Module2"
Function ConvRate(Orders, Leads)

If Leads > 0 Then
ConvRate = Orders / Leads

Else
ConvRate = 0
End If


End Function

