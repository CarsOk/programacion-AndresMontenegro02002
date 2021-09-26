***programacion***

 Sub ejemplo()<br>
 MsgBox "hola andres"<br>
 MsgBox "buenas tardes "<br>
 MsgBox "chao andres"

End Sub

*** segunda programacion***<br>
visual basic<br>
Sub inicio()<br>
 a = Int(InputBox("escriba primera nota"))<br>
 b = Int(InputBox("escriba segunda nota"))<br>
 c = Int(InputBox("escriba trecera notaa"))<br>
 d = Int(InputBox("escriba cuarta nota"))<br>
 e = Int(InputBox("escriba quinta nota"))<br>
 h = a + b + c + d + e<br>
 i = MsgBox(h / 5)<br>

 j = Int(InputBox("escribir el resultado"))<br>
 If (j > 7) Then<br>
 MsgBox "el resultado " & " es mayor 7 asignatura aprovada"<br>
 Else<br>
 MsgBox "el resultado " & " es menor 7 asignatura perdida"<br>
 End If<br>

End Sub

***como crear una funcion***

Function misnotas(a, b, c, d, e)
 Dim mis As Double
 mis = (a + b + c + d + e) / 5
 misnotas = mis
End Function
