Attribute VB_Name = "Module4"
Function CostperConv(Cost, Conversions)

If Conversions > 0 Then
CostperConv = Cost / Conversions

Else
CostperConv = 0
End If

End Function
