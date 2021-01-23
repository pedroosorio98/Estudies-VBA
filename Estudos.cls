Sub exemplo1()
'Passando tudo pra uppercase:
sstr = CStr(InputBox("Insira um string lowercase:"))
MsgBox (UCase(sstr))
End Sub
Sub exemplo2()
'Pegando o lado direito de um string:
sstr = CStr(InputBox("Insira um string lowercase:"))
MsgBox (Right(sstr, 5))
End Sub
Sub exemplo3()
'Mensagens de erro com vbCritical:
MsgBox ("ERRO"), vbCritical
End Sub
Sub exemplo4()
'Lbound e Ubound em arrays:
Dim MyArray(0 To 10, 10 To 20, 100 To 200)
MyArray(0, 10, 102) = 30
'Dá erro assim:
' MyArray(0, 4, 102) = 30 Pois 4 < 10

MsgBox ("1: " & Str(LBound(MyArray, 1)))
MsgBox ("2: " & Str(LBound(MyArray, 2)))
MsgBox ("3: " & Str(UBound(MyArray, 3)))

End Sub
Sub exemplo5()
'Replace(string,palavra,palavrareplace):
MsgBox (Replace("Pedro", "o", Str(0)))
End Sub
'Argumento opcional em função:
Function funcao(x, y, Optional z) 'Optional
    If z <> "" Then
        funcao = x + y + z
    Else
        funcao = x + y
    End If
End Function
Sub testardfuncao()
MsgBox (Str(funcao(10, 20, 30)))
MsgBox (Str(funcao(10, 20, "")))
End Sub
Sub exemplo6()
' Juntando os elementos de um array com Join():
Dim arr(10) As String
For i = 1 To 10
    arr(i) = Str(i)
Next i
MsgBox (Join(arr, "."))
End Sub
' Como usar Split:
Sub Splitar()
Dim vet(5) As Single
Sheet1.Cells(1, 1) = Split("P/E/D/R/O", "/") 'Isso vai fzr um array e vai retornar o elemento 0.
x = Split("P/E/D/R/O", "/")
Sheet1.Cells(2, 1) = Join(x, "") 'Juntando os elementos do split.
End Sub
'Usando left:
Sub leftright()
Dim p As String
Dim x As String
p = "PEDRO OSORIO"
x = left(p, 2) 'Pega as duas primeiras letras da esquerda.
y = Right(p, 2) 'Pega as duas primeiras letras da direita.
MsgBox ("x = " & x & " y = " & y)
End Sub
' Função para calcular o retorno acumulado de uma série de retornos
Function RetornoAcumulado(rng As Range) As Single
    Dim r As Single
    r = 1
    For Each cell In rng
        r = r * (1 + cell.Value)
    Next
RetornoAcumulado = r - 1
End Function
' Função que calcula o maximum drawdown de uma série
Function MaximumDrawdown(rng As Range) As Single
Dim MaxPrice As Single
Dim MaxDrawdown As Single
MaxPrice = -1
MaxDrawdown = 0
For Each cell In rng
    If cell.Value > MaxPrice Then
        MaxPrice = cell.Value
    End If
    If MaxDrawdown > cell.Value / MaxPrice - 1 Then
       MaxDrawdown = cell.Value / MaxPrice - 1
    End If
Next

MaximumDrawdown = MaxDrawdown

End Function
