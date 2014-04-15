<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Generar Claves</title>
</head>

<body>
<%
'Creando la funcion para generar
Function GenerarPassword(largo)
    Dim Resultado, Caracter, Password

    'Cargamos la matriz con números y letras
    caracter = Array("0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "a", "b", "c", "d", "e", "f", "g", "h", "i", "j", "k", "l", "m", "n", "o", "p", "q", "r", "s", "t", "u", "v", "w", "x", "y", "z")
    
    Randomize()
    Do While Len(Resultado) < largo
        Resultado = Resultado & Caracter(Int(36 * Rnd()))
    Loop
    GenerarPassword = Resultado
End Function

'Igualamos Variables a la funcion mas el parametro de longitud de la clave (5)
V1= GenerarPassword(30)
V2=GenerarPassword(12)
V3=GenerarPassword(15)
V4=GenerarPassword(12)
V5=GenerarPassword(2)	
%>
</body>
</html>
