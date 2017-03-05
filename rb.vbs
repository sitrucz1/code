Option Explicit

Private Const RED    = TRUE
Private Const BLACK  = FALSE

Class Tree
    Public Root
    Private DeleteCompleted

    Private Sub Class_Initialize
        ' wscript.echo "Tree Init"
        Set Root = Nothing
        DeleteCompleted = TRUE
    End Sub

    Private Function isLess(byval a, byval b)
        If a Is Nothing Or b Is Nothing Then
            isLess = FALSE
        Else
            isLess = (a.Data - b.Data) < 0
        End If
    End Function

    Public Function TreeAssert(byval vmin, byval vmax)
        TreeAssert = isRoot And isBST(Me.Root, vmin, vmax) And is23(Me.Root) And isBalanced
    End Function

    Private Function isRoot
        If Me.Root Is Nothing Then
            isRoot = TRUE
        ElseIf Me.Root.isRed Then
            Wscript.Echo "Root is Red"
            isRoot = FALSE
        Else
            isRoot = TRUE
        End If
    End Function

    Private Function isBST(n, byval vmin, byval vmax)
        If n Is Nothing Then
            isBST = TRUE
        ElseIf n.Data < vmin or n.Data > vmax Then
            Wscript.Echo "Data value is out of range ", n.Data
            isBST = FALSE
        ElseIf isLess(n, n.Lchild) Or isLess(n.Rchild, n) Then
            Wscript.Echo "Data of children don't meet criteria of parent"
            isBST = FALSE
        Else
            isBST = isBST(n.Lchild, vmin, vmax) And isBST(n.Rchild, vmin, vmax)
        End If
    End Function

    Private Function is23(n)
        If n Is Nothing Then
            is23 = TRUE
        ElseIf isRed(n) And (isRed(n.Lchild) Or isRed(n.Rchild)) Then
            Wscript.Echo "Two reds in a row at ", n.Data
            is23 = FALSE
        Else
            is23 = is23(n.Lchild) And is23(n.Rchild)
        End If
    End Function

    Private Function isBalanced
        Dim black, n
        black = 0
        Set n = Me.Root
        Do While Not n Is Nothing
            If Not isRed(n) Then
                black = black+1
            End If
            Set n = n.Lchild
        Loop
        isBalanced = isBal(Me.Root, black)
    End Function

    Private Function isBal(n, byval black)
        If n Is Nothing Then
            If black <> 0 Then
                Wscript.Echo "Black nodes don't match ", black
            End If
            isBal = (black = 0)
        Else
            If Not isRed(n) Then
                black = black-1
            End If
            isBal = isBal(n.Lchild, black) And isBal(n.Rchild, black)
        End If
    End Function

    Private Function isRed(n)
        If n Is Nothing Then
            isRed = FALSE
        Else
            isRed = (n.Color = RED)
        End If
    End Function

    Private Sub ColorFlip(n)
        n.Color = Not n.Color
        n.Lchild.Color = Not n.Lchild.Color
        n.Rchild.Color = Not n.Rchild.Color
    End Sub

    Private Function RotateLeft(n)
        Dim x
        Wscript.Echo "Rotate Left ", n.Data
        Set x = n.Rchild
        Set n.Rchild = x.Lchild
        Set x.Lchild = n
        x.Color = n.Color
        n.Color = RED
        Set RotateLeft = x
    End Function

    Private Function RotateRight(n)
        Dim x
        Wscript.Echo "Rotate Right ", n.Data
        Set x = n.Lchild
        Set n.Lchild = x.Rchild
        Set x.Rchild = n
        x.Color = n.Color
        n.Color = RED
        Set RotateRight = x
    End Function

    Public Function NodeInsert(byval v)
        Set Root = NodeIns(Root, v)
        Root.Color = BLACK
        NodeInsert = TRUE
    End Function

    Private Function NodeIns(n, byval v)
        If n Is Nothing Then
            Set n = (New Node).Init(v)
        Else
            If isRed(n.Lchild) And isRed(n.Rchild) Then
                ColorFlip(n)
            End If

            If v < n.Data Then
                Set n.Lchild = NodeIns(n.Lchild, v)
            ElseIf v > n.Data Then
                Set n.Rchild = NodeIns(n.Rchild, v)
            Else
                ' n was found so already in tree.
            End If

            If isRed(n.Lchild) Then
                If isRed(n.Lchild.Rchild) Then
                    Set n.Lchild = RotateLeft(n.Lchild)
                    Set n = RotateRight(n)
                ElseIf isRed(n.Lchild.Lchild) Then
                    Set n = RotateRight(n)
                End If
            ElseIf isRed(n.Rchild) Then
                If isRed(n.Rchild.Lchild) Then
                    Set n.Rchild = RotateRight(n.Rchild)
                    Set n = RotateLeft(n)
                ElseIf isRed(n.Rchild.Rchild) Then
                    Set n = RotateLeft(n)
                End If
            End If
        End If
        Set NodeIns = n
    End Function

    Public Function NodeInsI(byval v)
        Dim n, p, gp, ggp, u, f, stk, head
        Set head = New Node : head.Color = BLACK : head.Lchild = Root ' Sentinel above root
        Set stk = New Stack : stk.Push head
        Set f = (New Node).Init(v) : Set p = head : Set n = Root
        Do Until n Is Nothing
            If v = n.Data Then ' Item already in the tree
                Set NodeInsI = n
                Exit Function
            ElseIf v < n.Data Then
                Set p = n
                Set n = n.Lchild
            Else
                Set p = n
                Set n = n.Rchild
            End If
            stk.Push p
        Loop
        If p Is head Then ' At the root
            Set Root = f
        Else
            If f.Data < p.Data Then
                Set p.Lchild = f
            Else
                Set p.Rchild = f
            End If
            ' Rebalance
            Set p = stk.Pop
            Do While isRed(p)
                Set gp = stk.Pop
                If isRed(gp.Lchild) And isRed(gp.Rchild) Then
                    ColorFlip gp
                    Set p = stk.Pop
                ElseIf isRed(gp.Lchild) Then
                    If isRed(gp.Lchild.Rchild) Then ' Left/Right
                        Set gp.Lchild = RotateLeft(gp.Lchild.Rchild)
                    End If
                    If isRed(gp.Lchild.Lchild) Then ' Left/Left
                        Set ggp = stk.Pop
                        If ggp.Lchild is gp Then
                            Set ggp.Lchild = RotateRight(gp)
                            Set p = ggp.Lchild
                        Else
                            Set ggp.Rchild = RotateRight(gp)
                            Set p = ggp.Rchild
                        End If
                    End If
                ElseIf isRed(gp.Rchild) Then
                    If isRed(gp.Rchild.Lchild) Then ' Right/Left
                        Set gp.Rchild = RotateRight(gp.Rchild.Lchild)
                    End If
                    If isRed(gp.Rchild.Rchild) Then ' Right/Right
                        Set ggp = stk.Pop
                        If ggp.Lchild is gp Then
                            Set ggp.Lchild = RotateLeft(gp)
                            Set p = ggp.Lchild
                        Else
                            Set ggp.Rchild = RotateLeft(gp)
                            Set p = ggp.Rchild
                        End If
                    End If
                End If
            Loop
        End If
        Root.Color = BLACK
        Set NodeInsI = f
    End Function

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

    Public Function SearchBlackLeaf(n)
        If n Is Nothing Then
            Set SearchBlackLeaf = n
        ElseIf isBlack(n) and n.Lchild Is Nothing and n.Rchild Is Nothing Then
            Set SearchBlackLeaf = n
        Else
            Dim t
            Set t = SearchBlackLeaf(n.Lchild)
            If t Is Nothing Then
                Set t = SearchBlackLeaf(n.Rchild)
            End If
            Set SearchBlackLeaf = t
        End If
    End Function

    Public Sub InsertRandomData(byval cnt)
        Dim i, rnum
        Randomize timer
        For i = 1 to cnt
            rnum = rnd*100 mod cnt
            If Not Search(rnum) Then
                If Not NodeInsert(rnum) Then
                    Exit Sub
                End If
            End If
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
            wscript.stdout.write n.Data & "," & n.ColorChar(n) & " "
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

    Private Sub Class_Initialize
        ' Wscript.Echo "Node Init"
        Data  = -1
        Color = RED
        Set Lchild = Nothing
        Set Rchild = Nothing
    End Sub

    Public Function Init(n)
        Data  = N
        Color = RED
        Set Lchild = Nothing
        Set Rchild = Nothing
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
    End Sub
End Class

Class StackNode
    Public Data
    Public Nex

    Private Sub Class_Initialize
        Set Data = Nothing
        Set Nex = Nothing
    End Sub

    Public Function Init(n)
        Set Data = n
        Set nex = Nothing
        Set Init = Me
    End Function
End Class

Class Stack
    Private Cur

    Sub Class_Initialize
        Set Cur = Nothing
    End Sub

    Function Push(n)
        Dim t
        If n Is Nothing Then
            Set Push = n
            Exit Function
        End If
        Set t = (New StackNode).Init(n)
        Set t.Nex = Cur
        Set Cur = t
        Set Push = t
    End Function

    Function Pop
        Dim t
        If cur Is Nothing Then
            Set Pop = Nothing
            Exit Function
        End If
        Set t = Cur
        Set Cur = Cur.nex
        Set Pop = t.Data
        Set t = Nothing
    End Function
End Class

Dim T, n, i, S
Set T = New Tree
' T.InsertRandomData 50
T.NodeInsI 5
T.NodeInsI 3
T.NodeInsI 7
T.NodeInsI 9
T.PrintTree
' Set S = New Stack
' S.push T.Root
' S.push T.Root.Lchild
' S.push T.Root.Rchild
' Wscript.echo s.pop.Data
' Set n = s.Pop
' Wscript.echo n.data
' Set n = s.Pop
' Wscript.echo n.data
' For i = 1 to 5
' Do
'     T.DeleteNode(T.Root.Data)
'     T.PrintTree
'     If Not T.TreeAssert(0, 1000) Then
'         Wscript.Echo "Tree is not valid."
'     End If
' Loop Until T.Root Is Nothing
' Next
Set T = Nothing
