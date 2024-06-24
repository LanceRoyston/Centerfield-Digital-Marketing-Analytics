Attribute VB_Name = "Module1"
Function Profit(Revenue, Cost)

If Revenue > 0 Then
Profit = Revenue - Cost

Else
Profit = -1 * Cost
End If

End Function
