Private Const RED   = FALSE
Private Const BLACK = TRUE

Class Test
    Private Sub Class_Initialize
        Wscript.Echo "Init"
    End Sub

    Private Sub Class_Terminate
        Wscript.Echo "Term"
    End Sub
End Class

Function Nextnode(n)
    If Isempty(n) Then
        Set Nextnode = New Test
    Else
        Set Nextnode = n
    End If
End Function

' Dim T
' Set T = Nextnode(T)
Dim x, j
x = 1
j = 2
x = x = j
wscript.echo x
x = (x = j)
wscript.echo x
j = x
x = (x = j)
wscript.echo x
