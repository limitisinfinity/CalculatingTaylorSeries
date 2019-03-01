Sub TaylorSeries()

'Class Version - Dr. Luks 
'Worksheets.Select ("Sheet2")
'Taylor Series: SUM[F(x)] = SUM[(x)^k/k!]
Variable = 0.6684
ErrorLimit = 0.000001
Sum = 1
Fact = 1
    For k = 1 To 100
    '(2): Term 0 cannot equal 0 b/c then: !DIV/0 (see below)
    Fact = Fact * k
    Term = Variable ^ k / Fact

        If Term < ErrorLimit Then
            Debug.Print "k = " & k - 1
            ' B/C of (2), display "k" by calculating "k-1"^
            Debug.Print "k! = " & Fact
            ' Another way to write: MsgBox ("The Answer is" & Sum)
            Debug.Print "Term = " & Term
            Debug.Print "Sum = " & Sum
            
            ' End Loop by making k = # outside of the scope
            Dim Value As Range
            Set Value = Range("B1:B10")
            Value.Rows(1).Value = ErrorLimit
            Value.Rows(2).Value = Variable
            Value.Rows(3).Value = k
            Value.Rows(4).Value = Fact
            Value.Rows(5).Value = Term
            Value.Rows(6).Value = Sum
            k = 101
        End If
        Sum = Sum + Term
        If k = 100 Then
            Range("B3").Select
            ActiveCell.Value = "Formula did not converge in 100 terms; please enter in a large value for n"
        End If
    Next k
    Debug.Print Sum
    Debug.Print Term
End Sub
