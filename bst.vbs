Class BST
    Public Root

    Sub Class_Initialize
        Set Root = Nothing
    End Sub

    Function direc(a, b)
        direc = (a < b) + 1
    End Function

    Sub Class_Terminate
        Set Root = Nothing
    End Sub

    Function BST_Insert(n, v)
        Dim direction
        If n Is Nothing Then
            Set n = (New Node).Init(v)
        Else
            direction = direc(v, n.Data)
            Set n.Child(direction) = BST_Insert(n.Child(direction), v)
        End If
        Set BST_Insert = n
    End Function
End Class

Class Node
    Public Data
    Public Child(1)

    Sub Class_Initialize
        Set Child(0) = Nothing
        Set Child(1) = Nothing
    End Sub

    Function Init(v)
        Data = v
        Set Child(0) = Nothing
        Set Child(1) = Nothing
        Set Init = Me
    End Function

    Sub Class_Terminate
        Set Child(0) = Nothing
        Set Child(1) = Nothing
    End Sub
End Class

Dim t
Set t = New BST
Set t.Root = t.BST_Insert(t.Root, 5)
Set t.Root = t.BST_Insert(t.Root, 5)
Wscript.Echo t.direc(1,1), t.direc(1,2), t.direc(2,1)
