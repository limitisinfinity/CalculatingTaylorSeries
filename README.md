'Taylor Series
' Passing Value from a function to another function to Sub;
'Function for: K Factorial, Equation for Taylor series (SUMf(x))
'Sub (Main) is outputing (calling) the answer to the input (x)

Sub Findx()
Dim ErrorLimit As Double
Dim X As Double
ErrorLimit = 0.000001
X = 0.98488
Sum = 1
i = 100
End Sub

Function f(ByVal Sum As Double, ByVal ErrorLimit As Double, ByVal X As Double, ByVal k As Integer)
Dim J As Double
J = Fact(ByVal k, J)
    For i = 1 To k
        If k = 0 Then
          f = Fact(ByVal k, J)
        Else
            f = X ^ k / Fact(ByVal k, J)
        End If
            
            If f(ByVal X, ByVal k, ByVal NewJ) < ErrorLimit Then
                Dim Value As Range
                Set Value = Range("B1:B10")
                    Value.Rows(1).Value = "Error Limit: " & ErrorLimit
                    Value.Rows(2).Value = "x: " & X
                    Value.Rows(3).Value = "k: " & i
                    Value.Rows(4).Value = "k!: " & J
                    Value.Rows(5).Value = "F(x): " & f(ByVal X, ByVal k, ByVal NewJ)
                    Value.Rows(6).Value = "Sum[F(x)]: " & SUMf
                i = 101
            End If
                Sum = Sum + f(ByVal Sum, ByVal ErrorLimit, ByVal X, ByVal k)
            If i = k Then
                Value.Rows(7) = "Please enter a larger value for k"
            End If
          End If
    Next i
Exit Function
End Function

Function Fact(ByVal k As Integer, ByRef J As Double)
 k = k - 1
'Function Fact(ByVal Prodc As Integer, ByVal N As Integer, ByVal ErrorLimit As Double, ByRef J As Double)
    If k = 0 Then
Fact = 1
    End If
Fact = Fact(ByVal k, J) * (k - 1)

Exit Function
J = Fact(ByVal k, J)
End Function
