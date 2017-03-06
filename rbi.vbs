Option Explicit

Private Const RED    = TRUE
Private Const BLACK  = FALSE

Class RBTree
    Public Root

    Private Sub Class_Initialize
        Set Root = Nothing
    End Sub

    Public Function TreeAssert
        TreeAssert = Root Is Nothing Or TreeA(Root) > 0
    End Function

    Private Function TreeA(n)
        Dim lh, rh
        If n Is Nothing Then
            TreeA = 1
        ElseIf n Is Root And isRed(n) Then
            Wscript.Echo "Root is red "
            TreeA = 0
        ElseIf isLess(n, n.Lchild) Or isLess(n.Rchild, n) Then
            Wscript.Echo "Data doesn't match"
            TreeA = 0
        ElseIf isRed(n) And (isRed(n.Lchild) Or isRed(n.Rchild)) Then
            Wscript.Echo "Two red nodes in a row"
            TreeA = 0
        Else
            lh = TreeA(n.Lchild)
            rh = TreeA(n.Rchild)
            If lh <> 0 And rh <> 0 And lh <> rh Then
                Wscript.Echo "Black height violation"
                TreeA = 0
            ElseIf lh <> 0 and rh <> 0 Then
                If isRed(n) Then
                    TreeA = lh
                Else
                    TreeA = lh+1
                End If
            Else
                TreeA = 0
            End If
        End If
    End Function

    Private Function isRed(n)
        If n Is Nothing Then
            isRed = FALSE
        Else
            isRed = (n.Color = RED)
        End If
    End Function

    Private Function isLess(a, b)
        If a Is Nothing Or b Is Nothing Then
            isLess = FALSE
        Else
            isLess = a.Data < b.Data
        End If
    End Function

    Private Sub ColorFlip(n)
        n.Color = Not n.Color
        n.Lchild.Color = Not n.Lchild.Color
        n.Rchild.Color = Not n.Rchild.Color
    End Sub

    Public Sub RotateLeft(n)
        Dim x
        Wscript.Echo "Rotate Left ", n.Data
        Set x = n.Rchild
        Set n.Rchild = x.Lchild
        Set x.Lchild = n
        Set x.Parent = n.Parent
        Set n.Parent = x
        x.Color = n.Color
        n.Color = RED
        If x.Parent Is Nothing Then
            Set Root = x
        ElseIf x.Parent.Lchild Is n Then
            Set x.Parent.Lchild = x
        Else
            Set x.Parent.Rchild = x
        End If
    End Sub

    Private Sub RotateRight(n)
        Dim x
        Wscript.Echo "Rotate Right ", n.Data
        Set x = n.Lchild
        Set n.Lchild = x.Rchild
        Set x.Rchild = n
        Set x.Parent = n.Parent
        Set n.Parent = x
        x.Color = n.Color
        n.Color = RED
        If x.Parent Is Nothing Then
            Set Root = x
        ElseIf x.Parent.Lchild Is n Then
            Set x.Parent.Lchild = x
        Else
            Set x.Parent.Rchild = x
        End If
    End Sub

    Public Function NodeInsert(byval v)
        Dim n, p, f
        Set n = Root : Set p = Nothing
        Do Until n Is Nothing
            If v = n.Data Then ' Item already in the tree
                Set NodeInsert = n
                Exit Function
            ElseIf v < n.Data Then
                Set p = n
                Set n = n.Lchild
            Else
                Set p = n
                Set n = n.Rchild
            End If
        Loop
        Set f = (New Node).Init(v)
        Set f.Parent = p
        If p Is Nothing Then ' At the root
            Set Root = f
        Else
            If f.Data < p.Data Then
                Set p.Lchild = f
            Else
                Set p.Rchild = f
            End If
            InsertFixup f
        End If
        Root.Color = BLACK
        Set NodeInsert = f
    End Function

    Private Sub InsertFixup(n)
        Dim p, gp
        Set p = n.Parent
        Do
            If Not isRed(p) Then
                Exit Do
            End If

            Set gp = p.Parent
            If isRed(gp.Lchild) And isRed(gp.Rchild) Then ' Case 1
                ColorFlip gp
                Set p = gp.Parent
            ElseIf gp.Lchild Is p Then
                If isRed(gp.Lchild) And isRed(gp.Lchild.Rchild) Then ' Case 2
                    RotateLeft gp.Lchild
                End If
                If IsRed(gp.Lchild) And isRed(gp.Lchild.Lchild) Then ' Case 3
                    RotateRight gp
                End If
                Exit Do
            Else
                If isRed(gp.Rchild) And isRed(gp.Rchild.Lchild) Then ' Case 2
                    RotateRight p
                End If
                If IsRed(gp.Rchild) And isRed(gp.Rchild.Rchild) Then ' Case 3
                    RotateLeft gp
                End If
                Exit Do
            End If
        Loop Until FALSE
    End Sub

    Private Function SearchNode(n, byval v)
        If n Is Nothing Then
            SearchNode = FALSE
        ElseIf v < n.Data Then
            SearchNode = SearchNode(n.Lchild, v)
        ElseIf v > n.Data Then
            SearchNode = SearchNode(n.Rchild, v)
        Else
            SearchNode = TRUE
        End If
    End Function

    Public Function Search(byval v)
        Search = SearchNode(Me.Root, v)
    End Function

    Private Function InOrderSuccessor(n)
        Dim v
        Set v = n.Rchild
        Do While Not v.Lchild Is Nothing
            Set v = v.Lchild
        Loop
        Set InOrderSuccessor = v
    End Function

    Private Function DelFixupLeft(n) ' Double black is to the left
        Dim s ' sibling of Double Black
        If Not DeleteCompleted Then
            Wscript.Echo "DelFixupLeft"
            Set s = n.Rchild
            If Not s Is Nothing Then
                If Not isRed(s) And (isRed(s.Lchild) Or isRed(s.Rchild)) Then ' Case 1
                    Wscript.Echo "Case 1 - Sibling Black with a Red child"
                    If isRed(s.Rchild) Then
                        Set n = RotateLeft(n)
                    ElseIf isRed(s.Lchild) Then
                        Set n.Rchild = RotateRight(n.Rchild)
                        Set n = RotateLeft(n)
                    End If
                    n.Lchild.Color = BLACK
                    n.Rchild.Color = BLACK
                    DeleteCompleted = TRUE
                ElseIf Not isRed(s) And Not isRed(s.Lchild) And Not isRed(s.Rchild) Then ' Case 2
                    Wscript.Echo "Case 2 - Sibling Black with Black Children ", isRed(n)
                    DeleteCompleted = isRed(n) ' If parent RED we are done otherwise push it up a level
                    n.Color = BLACK
                    s.Color = RED
                ElseIf isRed(s) Then ' Case 3
                    Wscript.Echo "Case 3 - Sibling Red"
                    Set n = RotateLeft(n)
                    Set n.Lchild = DelFixupLeft(n.Lchild) ' Let's recursively fix this since it's now a previous case
                    DeleteCompleted = TRUE
                End If
            End If
        End If
        Set DelFixupLeft = n
    End Function

    Private Function DelFixupRight(n) ' Double black is to the right
        Dim s ' sibling of Double Black
        If Not DeleteCompleted Then
            Wscript.Echo "DelFixupRight"
            Set s = n.Lchild
            If Not s Is Nothing Then
                If Not isRed(s) And (isRed(s.Lchild) Or isRed(s.Rchild)) Then ' Case 1
                    Wscript.Echo "Case 1 - Sibling Black with a Red child"
                    If isRed(s.Lchild) Then
                        Set n = RotateRight(n)
                    ElseIf isRed(s.Rchild) Then
                        Set n.Lchild = RotateLeft(n.Lchild)
                        Set n = RotateRight(n)
                    End If
                    n.Lchild.Color = BLACK
                    n.Rchild.Color = BLACK
                    DeleteCompleted = TRUE
                ElseIf Not isRed(s) And Not isRed(s.Lchild) And Not isRed(s.Rchild) Then ' Case 2
                    Wscript.Echo "Case 2 - Sibling Black with Black Children ", isRed(n)
                    DeleteCompleted = isRed(n) ' If parent RED we are done otherwise push it up a level
                    n.Color = BLACK
                    s.Color = RED
                ElseIf isRed(s) Then ' Case 3
                    Wscript.Echo "Case 3 - Sibling Red"
                    Set n = RotateRight(n)
                    Set n.Rchild = DelFixupRight(n.Rchild) ' Let's recursively fix this since it's now a previous case
                    DeleteCompleted = TRUE
                End If
            End If
        End If
        Set DelFixupRight = n
    End Function

    Private Function DelNode(n, byval v)
        Dim t
        If Not n Is Nothing Then
            If v < n.Data Then
                Set n.Lchild = DelNode(n.Lchild, v)
                Set n = DelFixupLeft(n)
            ElseIf v > n.Data Then
                Set n.Rchild = DelNode(n.Rchild, v)
                Set n = DelFixupRight(n)
            Else
                If n.Lchild Is Nothing And n.Rchild Is Nothing Then
                    Wscript.Echo "Deleting Leaf ", isRed(n)
                    DeleteCompleted = isRed(n)
                    Set n = Nothing
                ElseIf n.Lchild Is Nothing Then
                    Wscript.Echo "Deleting One Child Leaf"
                    Set t = n
                    Set n = t.Rchild
                    n.Color = BLACK
                    Set t = Nothing
                    DeleteCompleted = TRUE
                ElseIf n.Rchild Is Nothing Then
                    Wscript.Echo "Deleting One Child Leaf"
                    Set t = n
                    Set n = t.Lchild
                    n.Color = BLACK
                    Set t = Nothing
                    DeleteCompleted = TRUE
                Else
                    Set t = InOrderSuccessor(n)
                    n.Data = t.Data
                    Set n.Rchild = DelNode(n.Rchild, t.Data)
                    Set n = DelFixupRight(n)
                End If
            End If
        End If
        Set DelNode = n
    End Function

    Public Function DeleteNode(byval v)
        Set Root = DelNode(Root, v)
        If Not Root Is Nothing Then
            Root.Color = BLACK
        End If
        DeleteNode = TRUE
    End Function

    Public Sub InsertRandomData(byval cnt)
        Dim i, rnum
        Randomize timer
        For i = 1 to cnt
            rnum = rnd*100 mod cnt
            Wscript.Echo rnum
            If Not Search(rnum) Then
                If NodeInsert(rnum) Is Nothing Then
                    Exit Sub
                End If
            End If
            PrintTree
        Next
    End Sub

    Public Sub PrintTree
        PrintNode(Me.Root)
        Wscript.Echo ""
    End Sub

    Private Sub PrintNode(n)
        If Not n Is Nothing Then
            If n.Lchild Is Nothing Then
                wscript.stdout.write "*,b "
            Else
                wscript.stdout.write n.Lchild.Data & "," & n.Lchild.ColorChar(n.Lchild) & " "
            End If
            wscript.stdout.write n.Data & "," & n.ColorChar(n)
            If n.Parent Is Nothing Then
                wscript.stdout.write ",* "
            Else
                wscript.stdout.write "," & n.Parent.Data & " "
            End If
            If n.Rchild Is Nothing Then
                wscript.stdout.write "*,b "
            Else
                wscript.stdout.write n.Rchild.Data & "," & n.Rchild.ColorChar(n.Rchild) & " "
            End If
            wscript.echo ""
            PrintNode(n.Lchild)
            PrintNode(n.Rchild)
        End If
    End Sub

    Private Sub Class_Terminate
        ' wscript.echo "Tree Term"
        Set Root = Nothing
    End Sub
End Class

Class Node
    Public Data
    Public Color
    Public Lchild
    Public Rchild
    Public Parent

    Private Sub Class_Initialize
        ' Wscript.Echo "Node Init"
        Data  = -1
        Color = RED
        Set Lchild = Nothing
        Set Rchild = Nothing
        Set Parent = Nothing
    End Sub

    Public Function Init(n)
        Data  = N
        Color = RED
        Set Lchild = Nothing
        Set Rchild = Nothing
        Set Parent = Nothing
        Set Init = Me
    End Function

    Public Function ColorChar(n)
        If n Is Nothing Then
            ColorChar = "b"
        ElseIf n.Color = RED Then
            ColorChar = "r"
        Else
            ColorChar = "b"
        End If
    End Function

    Public Function isRed
        isRed = (Me.Color = RED)
    End Function

    Private Sub Class_terminate
        ' Wscript.Echo "Node Term"
        Set Lchild = Nothing
        Set Rchild = Nothing
        Set Parent = Nothing
    End Sub
End Class

Dim T, n, i, S
Set T = New RBTree
T.InsertRandomData 15
T.PrintTree
If Not T.TreeAssert Then
    Wscript.Echo "Tree is not valid."
End If
