# CalculatingTaylorSeries
'Calculate Taylor Series using a Factorial
' First Class Example
Sub FindF()
'My version
'Taylor Series Function
'f(x) = x^k/k!; Sum[f(x)]; for: n=1 to k(infinity) terms
'future program possible changes include:
'   Sum, n, a (At I = a To (n)),
'   F(x) function, and if F is Abs(F)
    Dim MyAnswer As Double
    'Dim Value As Range
    'Set Value = Range("B1:B10")
    MyAnswer = f(0.6684)
    'Value.Rows(6).Value = "Sum[F(x)]: " & MyAnswer
End Sub

Function f(X As Double) As Double
Dim ErrorLimit As Double
ErrorLimit = 0.00001
Fact = 1
Sum = 1
    For i = 1 To 100
        Fact = Fact * i
        f = X ^ i / Fact
        'This is f(x)
            If f < ErrorLimit Then
            Dim Value As Range
            Set Value = Range("B1:B10")
                Value.Rows(1).Value = "Error Limit: " & ErrorLimit
                Value.Rows(2).Value = "x: " & X
                Value.Rows(3).Value = "I: " & i
                Value.Rows(4).Value = "k!: " & Fact
                Value.Rows(5).Value = "F(x): " & f
                Value.Rows(6).Value = "Sum[F(x)]: " & Sum
'                Debug.Print "x: " & x
'                Debug.Print "I: " & I
'                Debug.Print "k!: " & Fact
'                Debug.Print "F(x): " & F
'                Debug.Print "Sum[F(x)]: " & Sum
            i = 101
            End If
        Sum = Sum + f
        f = Sum
            If i = 100 Then
            Value.Rows(7) = "Please enter a larger value for n"
            End If
    Next i
End Function
