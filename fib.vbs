Option Explicit
On Error Resume Next

Function Fib(n)
    If n < 2 Then
        Fib = n
    Else
        Fib = Fib(n-2) + Fib(n-1)
    End If
End Function

Function getArg
    If Wscript.Arguments.Count > 0 Then
        getArg = Wscript.Arguments.Item(0)
    Else
        getArg = 4
    End If
End Function

Wscript.Echo Fib(getArg)
