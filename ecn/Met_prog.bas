Attribute VB_Name = "Met_prog"
Function fx_bis(x) 'Función de bisección f(x)
'fx_bis = Log(x ^ 2) - 0.7 '5.6
'fx_bis = Sin(x) - x ^ 2 '5.5
fx_bis = (9.81 * x / 15) * (1 - Exp(-15 * 10 / x)) - 36 '5.13
End Function
Sub Bis_met() 'Método de bisección
xl = Cells(2, 1)
xu = Cells(2, 2)
xr = Cells(2, 3)
max_i = Cells(2, 4)
max_e = Cells(2, 5)
er = 100
i = 0
While i < max_i
'While er > max_e
    xrold = xr
    xr = (xu + xl) / 2
    i = i + 1
    If fx_bis(xl) * fx_bis(xr) < 0 Then
        xu = xr
    End If
    If fx_bis(xl) * fx_bis(xr) > 0 Then
        xl = xr
    End If
    er = Abs((xr - xrold) / xr) * 100
Wend
Cells(2, 6) = xr
Cells(2, 7) = er
End Sub
Function fx_fal(x) 'Función de posición falsa f(x)
'fx_fal = Log(x ^ 2) - 0.7 '5.6
'fx_fal = Sin(x) - x ^ 2 '5.5
fx_fal = (9.81 * x / 15) * (1 - Exp(-15 * 10 / x)) - 36 '5.13
End Function
Sub Fal_met() 'Método de posición falsa
xl = Cells(2, 1)
xu = Cells(2, 2)
xr = Cells(3, 3)
max_i = Cells(2, 4)
max_e = Cells(2, 5)
er = 100
i = 0
fxl = fx_fal(xl)
fxu = fx_fal(xu)
While i < max_i
'While er > max_e
    xrold = xr
    i = i + 1
    If fxl * fxr < 0 Then
        xu = xr
        fxu = fx_fal(xu)
    End If
    If fx_fal(xl) * fx_fal(xr) > 0 Then
        xl = xr
        fxl = fx_fal(xl)
    End If
    xr = xu - fxu * (xl - xu) / (fxl - fxu)
    fxr = fx_fal(xr)
    er = Abs((xr - xrold) / xr) * 100
Wend
Cells(2, 8) = xr
Cells(2, 9) = er
End Sub
Function fx_new(x) 'Función de Newton Raphson f(x)
'fx_new = Log(x ^ 2) - 0.7 '5.6
'fx_new = Sin(x) - x ^ 2 '5.5
fx_new = (9.81 * x / 15) * (1 - Exp(-15 * 10 / x)) - 36 '5.13
End Function
Function dfx_new(x) 'Función de Newton Raphson df(x)
'dfx_new = 2 * x / (x ^ 2) '5.6
'dfx_new = Cos(x) - 2 * x '5.5
dfx_new = 1 - Exp(-150 / x) * (1 + 150 / x) '5.13
End Function
Sub New_met() 'Método de newton raphson
x = Cells(2, 1)
max_i = Cells(2, 4)
max_e = Cells(2, 5)
er = 100
i = 0
max_err = Cells(2, 5)
er = 100
While i < max_i
'While er > max_e
    i = i + 1
    fx = fx_new(x)
    dfx = dfx_new(x)
    xr = x - fx / dfx
    er = Abs((xr - x) / xr) * 100
    x = xr
Wend
Cells(2, 10) = xr
Cells(2, 11) = er
End Sub
Sub Cra_met() 'Método de Cramer
n = 5 'Número de datos
Dim a() 'creación de vectores
Dim b()
Dim s()
ReDim a(1 To n, 1 To n)
ReDim b(1 To n)
ReDim s(1 To n)
For i = 0 To n - 1 'creación matriz A, vector b
    For j = 0 To n - 1 'matriz A coeff
        a(i + 1, j + 1) = Cells(6 + i, 1 + j)
'        Cells(11 + i, 1 + j) = a(i + 1, j + 1)
    b(i + 1) = Cells(6 + i, 7)
    Next j
Next i
For k = 1 To n - 1 'solución del sistema por eliminación de Gauss hacia adelante
    For i = k + 1 To n
        factor = a(i, k) / a(k, k)
        For j = k + 1 To n
            a(i, j) = a(i, j) - factor * a(k, j)
        Next j
        b(i) = b(i) - factor * b(k)
    Next i
Next k
s(n) = b(n) / a(n, n) 'sustitución hacia atrás
For i = n - 1 To 1 Step -1
    Sum = b(i)
    For j = i + 1 To n
        Sum = Sum - a(i, j) * s(j)
    Next j
    s(i) = Sum / a(i, i)
Next i
For i = 0 To n - 1
    Cells(6 + i, 9) = s(i + 1)
'    Cells(11 + i, 7) = b(i + 1)
Next i
End Sub
Sub Gau_met() 'Método de Gauss Seidel
x1 = Cells(6, 7) / Cells(6, 1)
x2 = (Cells(7, 7) - x1 * Cells(7, 1)) / Cells(7, 2)
x3 = (Cells(8, 7) - x1 * Cells(8, 1) - x2 * Cells(8, 2)) / Cells(8, 3)
x4 = (Cells(9, 7) - x1 * Cells(9, 1) - x2 * Cells(9, 2) - x3 * Cells(9, 3)) / Cells(9, 4)
x5 = (Cells(10, 7) - x1 * Cells(10, 1) - x2 * Cells(10, 2) - x3 * Cells(10, 3) - x4 * Cells(10, 4)) / Cells(10, 5)
max_e = Cells(12, 1)

er = 100
While er > max_e
    x1n = (Cells(6, 7) - x2 * Cells(6, 2) - x3 * Cells(6, 3) - x4 * Cells(6, 4) - x5 * Cells(6, 5)) / Cells(6, 1)
    x2n = (Cells(7, 7) - x1n * Cells(7, 1) - x3 * Cells(7, 3) - x4 * Cells(7, 4) - x5 * Cells(7, 5)) / Cells(7, 2)
    x3n = (Cells(8, 7) - x1n * Cells(8, 1) - x2n * Cells(8, 2) - x4 * Cells(8, 4) - x5 * Cells(8, 5)) / Cells(8, 3)
    x4n = (Cells(9, 7) - x1n * Cells(9, 1) - x2n * Cells(9, 2) - x3n * Cells(9, 3) - x5 * Cells(9, 5)) / Cells(9, 4)
    x5n = (Cells(10, 7) - x1n * Cells(10, 1) - x2n * Cells(10, 2) - x3n * Cells(10, 3) - x4n * Cells(10, 4)) / Cells(10, 5)
    er = Abs((x1n - x1) / x1n) * 100
    x1 = x1n
    x2 = x2n
    x3 = x3n
    x4 = x4n
    x5 = x5n
Wend
Cells(7, 11) = er
Cells(6, 10) = x1
Cells(7, 10) = x2
Cells(8, 10) = x3
Cells(9, 10) = x4
Cells(10, 10) = x5
End Sub
