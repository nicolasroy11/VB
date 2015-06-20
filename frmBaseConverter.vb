Option Explicit On
'Imports System
'Imports System.IO
'Imports System.Collections
Public Class frmBaseConverter
    Public frmdft As New clsFormDefault 'used to call a class - clsFormDefault is the form color scheme

    Private Sub frmBaseConverter_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Call frmdft.FormDefault() : Me.BackColor = Color.MintCream
    End Sub

    Private Sub cmdExit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdExit.Click
        End
    End Sub

    Private Sub cmdConvert_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdConvert.Click
        Dim baseOut As Integer, baseIn As Integer, number As Integer
        number = Val(t9.Text)
        baseOut = Val(t12.Text)
        baseIn = Val(t10.Text)
        If InputValidation(number, baseIn) = 0 Then
            ErrorMsg(0)
        ElseIf t12.Text = "" Then
            ErrorMsg(4)
        Else
            t11.Text = CStr(mdlBaseConversion.Convert(number, baseOut, baseIn))
        End If
    End Sub

    Private Sub cmdConvertSeries_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdConvertSeries.Click
        Dim Fn As Integer, En As Integer, Inc As Integer, Fib As Integer
        Dim Eib As Integer, Fob As Integer, Eob As Integer
        Dim j As Integer, k As Integer, l As Integer, result As String, Dif As Integer, lstResult As String

        Fn = Val(t1.Text)
        En = Val(t2.Text)
        Inc = Val(t3.Text)
        Fib = Val(t4.Text)
        Eib = Val(t5.Text)
        Fob = Val(t6.Text)
        Eob = Val(t7.Text)
        If Fib > Eib Then ErrorMsg(1)
        If Fob > Eob Then ErrorMsg(2)
        If Fn > En Then ErrorMsg(3)
        'count = 0
        result = 0
        lstResult = ""
        t8.Text = ""
        For j = Fn To En Step Inc
            'count = count + 1
            For k = Fib To Eib
                For l = Fob To Eob
                    If Not InputValidation(j, k) Then GoTo 100
                    result = CStr(mdlBaseConversion.Convert(j, l, k))
                    Dif = j - Val(result)
                    'j = Fn + Inc * count
                    lstResult = lstResult & CStr(j) & vbTab & CStr(k) & vbTab & CStr(l) & vbTab & result & vbTab & CStr(Dif) & vbNewLine
                Next l
            Next k
100:    Next j
        t8.Text = lstResult
    End Sub

    Private Function InputValidation(ByVal n As Integer, ByVal b As Integer) As Boolean
        Dim i As Integer, parsedInp(0 To 15) As String, v As Boolean
        If cboInputValidation.Checked = False Then
            InputValidation = 1
            Exit Function
        Else
            For i = 1 To Len(CStr(n))
                parsedInp(i) = Mid(CStr(n), i, 1)
                If Val(parsedInp(i)) < b Then v = 1 Else v = 0
            Next i
        End If
        InputValidation = v

    End Function

    Private Sub ErrorMsg(ByVal i As Integer)
        Dim errmsg As String
        Select Case i
            Case 0 : errmsg = "Your input is beyond base tolerance. If you still wish to proceed with these parameters, check off the input validation option."
            Case 1 : errmsg = "Your input end base must be greater or equal to your start base."
            Case 2 : errmsg = "Your output end base must be greater or equal to your start base."
            Case 3 : errmsg = "Your first input number must be greater or equal to your end input number."
            Case 4 : errmsg = "An end conversion base is needed."
        End Select
        MsgBox(errmsg)
    End Sub
    Private Sub t1_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles t1.TextChanged
        If t1.Text <> "" Then t1.BackColor = Color.Gray _
            : l1.ForeColor = Color.LightSkyBlue Else t1.BackColor = Color.LightGray : l1.ForeColor = Color.DarkGray
    End Sub

    Private Sub t2_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles t2.TextChanged
        If t2.Text <> "" Then t2.BackColor = Color.Gray _
            : l2.ForeColor = Color.LightSkyBlue Else t2.BackColor = Color.LightGray : l2.ForeColor = Color.DarkGray
    End Sub

    Private Sub t3_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles t3.TextChanged
        If t3.Text <> "" Then t3.BackColor = Color.Gray _
            : l3.ForeColor = Color.LightSkyBlue Else t3.BackColor = Color.LightGray : l3.ForeColor = Color.DarkGray
    End Sub

    Private Sub t4_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles t4.TextChanged
        If t4.Text <> "" Then t4.BackColor = Color.Gray _
            : l4.ForeColor = Color.LightSkyBlue Else t4.BackColor = Color.LightGray : l4.ForeColor = Color.DarkGray
    End Sub

    Private Sub t5_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles t5.TextChanged
        If t5.Text <> "" Then t5.BackColor = Color.Gray _
            : l5.ForeColor = Color.LightSkyBlue Else t5.BackColor = Color.LightGray : l5.ForeColor = Color.DarkGray
    End Sub

    Private Sub t6_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles t6.TextChanged
        If t6.Text <> "" Then t6.BackColor = Color.Gray _
            : l6.ForeColor = Color.LightSkyBlue Else t6.BackColor = Color.LightGray : l6.ForeColor = Color.DarkGray
    End Sub

    Private Sub t7_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles t7.TextChanged
        If t7.Text <> "" Then t7.BackColor = Color.Gray _
            : l7.ForeColor = Color.LightSkyBlue Else t7.BackColor = Color.LightGray : l7.ForeColor = Color.DarkGray
    End Sub

    Private Sub t8_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles t8.TextChanged
        If t8.Text <> "" Then t8.BackColor = Color.Gray '_
        ': l8.ForeColor = Color.LightSkyBlue Else t8.BackColor = Color.LightGray : l8.ForeColor = Color.DarkGray
    End Sub
    Private Sub t9_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles t9.TextChanged
        If t9.Text <> "" Then t9.BackColor = Color.Gray _
          : l9.ForeColor = Color.LightSkyBlue Else t9.BackColor = Color.LightGray : l9.ForeColor = Color.DarkGray
    End Sub

    Private Sub t10_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles t10.TextChanged
        If t10.Text <> "" Then t10.BackColor = Color.Gray _
          : l10.ForeColor = Color.LightSkyBlue Else t10.BackColor = Color.LightGray : l10.ForeColor = Color.DarkGray
    End Sub

    Private Sub t11_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles t11.TextChanged
        If t11.Text <> "" Then t11.BackColor = Color.Gray _
          : l11.ForeColor = Color.LightSkyBlue Else t11.BackColor = Color.LightGray : l11.ForeColor = Color.DarkGray
    End Sub

    Private Sub t12_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles t12.TextChanged
        If t12.Text <> "" Then t12.BackColor = Color.Gray _
          : l12.ForeColor = Color.LightSkyBlue Else t12.BackColor = Color.LightGray : l12.ForeColor = Color.DarkGray
    End Sub
End Class