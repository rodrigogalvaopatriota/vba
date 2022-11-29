Function TiraAcento(Palavra)
 CAcento = "àáâãäèéêëìíîïòóôõöùúûüÀÁÂÃÄÈÉÊËÌÍÎÒÓÔÕÖÙÚÛÜçÇñÑ"
 SAcento = "aaaaaeeeeiiiiooooouuuuAAAAAEEEEIIIOOOOOUUUUcCnN"
 Texto = ""
 if Palavra <> "" then
 For X = 1 to Len(Palavra)
 Letra = mid(Palavra,X,1)
 Pos_Acento = inStr(CAcento,Letra)
 if Pos_Acento > 0 then
 Letra = mid(SAcento,Pos_Acento,1)
 end if
 Texto = Texto & Letra
 next
 TiraAcento = Texto
 end if
 end function

Function VerificaPalavra(atributo)

Dim i
 Dim id
 Dim Auxiliar
 Dim Resultado

Auxiliar = Split(Atributo, " ", - 1, vbBinaryCompare)

For i = LBound(Auxiliar) To Ubound(Auxiliar)
 Resultado = Resultado & " " & TiraAcento(Auxiliar(i))
 Next

VerificaPalavra = Trim(Resultado)
 end function
