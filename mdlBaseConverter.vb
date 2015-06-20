Option Explicit On

Module mdlBaseConversion
    Public Const beyondTen = "ABCDEFGHIJKLMNOPQRSTUVWXYZ!@#$%^&*(){}[]|/\:<>?,'`;+-_~"
    Public Function Convert(ByVal number As Integer, ByVal base1 As Integer, ByVal base2 As Integer) As Object
        Dim k As Integer, n As Integer, digit(100) As Integer, result As String, m As String

        If base2 <> 10 Then number = Decimalize(number, base2)
        result = ""
        For k = DigitCount(number, base1) To 0 Step -1
            digit(k) = 0
            For n = 1 To base1 - 1
                If Compare(number, k, base1) < 0 Then
                    digit(k) = digit(k) + 0
                    Exit For
                Else
                    number = Compare(number, k, base1)
                    digit(k) = n
                End If
            Next n
            m = CStr(digit(k))
            'If digit(DigitCount(number, base1)) = 0 Then result = ""
            If digit(k) > 9 Then m = Mid(beyondTen, digit(k) - 9, 1)
            result = result & m & " "
        Next k

        Convert = result

    End Function

    Private Function Compare(ByVal num, ByVal n, ByVal base) As Double
        Compare = num - base ^ n
    End Function

    Private Function DigitCount(ByVal num As Double, ByVal base As Integer)
        Dim i As Integer
        For i = 0 To 1000
            If Compare(num, i, base) < 0 Then
                Exit For
            Else
                num = Compare(num, i, base)
            End If
        Next i
        DigitCount = i
    End Function
    Public Function Decimalize(ByVal input As Integer, ByVal inbase As Integer)
        Dim dig(10) As Integer, bound As Integer, i As Integer, Tenbase As Integer
        bound = Len(CStr(input))
        Tenbase = 0
        For i = bound To 1 Step -1
            dig(i) = Val(Mid(input, bound - i + 1, 1))
            Tenbase = Tenbase + dig(i) * inbase ^ (i - 1)
        Next
        Decimalize = Tenbase
    End Function
End Module
