Public Class clsFormDefault

    Public Sub FormDefault()
        Dim ctrl As Control
        For Each ctrl In frmBaseConverter.Controls
            Select Case TypeName(ctrl)
                Case "TextBox"
                    With ctrl
                        .BackColor = Color.LightGray
                        .Cursor = Cursors.PanEast
                        .ForeColor = Color.White
                        .Font = New Font("Lucida Console", 9, FontStyle.Regular)
                    End With
                Case "Button"
                    With ctrl
                        .Cursor = Cursors.NoMove2D
                        .ForeColor = Color.DarkGray
                        .Font = New Font("Lucida Console", 9, FontStyle.Regular)
                    End With
                Case "Label"
                    With ctrl
                        .Font = New Font("Lucida Console", 9, FontStyle.Regular)
                        .BackColor = Color.Transparent
                        .ForeColor = Color.DarkGray
                    End With
                Case "Form"
                    With ctrl
                        .BackColor = Color.White
                        .Cursor = Cursors.PanNorth
                    End With
                Case "CheckBox"
                    With ctrl
                        .BackColor = Color.Transparent
                        .Cursor = Cursors.PanNorth
                        .ForeColor = Color.DarkGray
                    End With
            End Select
        Next
    End Sub

    Public Sub FormChange()
        For Each ctrl In frmBaseConverter.Controls
            ctrl.BackColor = Color.Gray
            ctrl.ForeColor = Color.LightSkyBlue
        Next
    End Sub
    'Private Sub t2_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles t2.TextChanged
    '    If t2.Text <> "" Then t2.BackColor = Color.Gray _
    '     : l2.ForeColor = Color.LightSkyBlue Else t2.BackColor = Color.LightGray : l2.ForeColor = Color.DarkGray
    'End Sub

    'Private Sub t3_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles t3.TextChanged
    '    If t3.Text <> "" Then t3.BackColor = Color.Gray _
    '      : l3.ForeColor = Color.LightSkyBlue Else t3.BackColor = Color.LightGray : l3.ForeColor = Color.DarkGray
    'End Sub

    'Private Sub t4_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles t4.TextChanged
    '    If t4.Text <> "" Then t4.BackColor = Color.Gray _
    '      : l4.ForeColor = Color.LightSkyBlue Else t4.BackColor = Color.LightGray : l4.ForeColor = Color.DarkGray
    'End Sub

    'Private Sub t5_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles t5.TextChanged
    '    If t5.Text <> "" Then t5.BackColor = Color.Gray _
    '     : l5.ForeColor = Color.LightSkyBlue Else t5.BackColor = Color.LightGray : l5.ForeColor = Color.DarkGray
    'End Sub

    'Private Sub t6_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles t6.TextChanged
    '    If t6.Text <> "" Then t6.BackColor = Color.Gray _
    '     : l6.ForeColor = Color.LightSkyBlue Else t6.BackColor = Color.LightGray : l6.ForeColor = Color.DarkGray
    'End Sub

    'Private Sub b2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles b2.Click

    '    t1.Text = ""
    '    t2.Text = ""
    '    t3.Text = ""
    '    t4.Text = ""
    '    t5.Text = ""
    '    t6.Text = ""



    'Private Sub b3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles b3.Click
    '    End
    'End Sub
End Class
