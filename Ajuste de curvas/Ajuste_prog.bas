Attribute VB_Name = "Ajuste_prog"
Sub Pol_int_Newton()
n = 6 'Numero de datos
'xx = Cells(9, 1) + 273.15 'Numero (x) donde se ajusta valor f(x)
'xx = Cells(9, 5) + 273.15 'Numero (x) donde se ajusta valor f(x)
xx = Cells(9, 9)
Dim x(6)
Dim y(6)
Dim yint(6)
Dim fdd(6, 6)
For i = 0 To n - 1 'Entrada de datos x,y
'    x(i) = Cells(3 + i, 1) + 273.15
'    x(i) = Cells(3 + i, 5) + 273.15
    x(i) = Cells(3 + i, 9)
'    y(i) = Cells(3 + i, 2)
'    y(i) = Cells(3 + i, 6)
    y(i) = Cells(3 + i, 10)
    fdd(i, 0) = y(i)
Next i
For j = 1 To n
    For i = 0 To n - j
        fdd(i, j) = (fdd(i + 1, j - 1) - fdd(i, j - 1)) / (x(i + j) - x(i))
    Next i
Next j
xterm = 1
yint(0) = fdd(0, 0)
For o = 1 To n
    xterm = xterm * (xx - x(o - 1))
    yint(o) = yint(o - 1) + fdd(0, o) * xterm
Next o
'Cells(9, 2) = yint(n - 1)
'Cells(9, 6) = yint(n - 1)
Cells(9, 10) = yint(n - 1)
End Sub

Sub Pol_int_LaGrange()
n = 6 - 1 'Numero de datos - 1
'xx = Cells(10, 1) 'Numero (x) donde se ajusta valor f(x)
'xx = Cells(10, 5) 'Numero (x) donde se ajusta valor f(x)
xx = Cells(10, 9) 'Numero (x) donde se ajusta valor f(x)
Dim x(6)
Dim y(6)
Sum = 0
For i = 0 To n 'Entrada de datos x,y
'    x(i) = Cells(3 + i, 1)
'    x(i) = Cells(3 + i, 5)
    x(i) = Cells(3 + i, 9)
'    y(i) = Cells(3 + i, 2)
'    y(i) = Cells(3 + i, 6)
    y(i) = Cells(3 + i, 10)
Next i
For i = 0 To n
    producto = y(i)
    For j = 0 To n
        If i <> j Then
            producto = producto * (xx - x(j)) / (x(i) - x(j))
        End If
    Next j
    Sum = Sum + producto
Next i
'Cells(10, 2) = Sum
'Cells(10, 6) = Sum
Cells(10, 10) = Sum
End Sub

Sub Ajuste_min_cuadrados()
'n = Cells(1, 1)
n = 6 'Número de datos
o = 5 'Orden del polinomio
'xx = Cells(11, 1) 'Numero (x) donde se ajusta valor f(x)
'xx = Cells(11, 5) 'Numero (x) donde se ajusta valor f(x)
xx = Cells(11, 9) 'Numero (x) donde se ajusta valor f(x)
Dim a() 'creación de vectores
Dim b()
Dim x()
Dim y()
Dim s()
ReDim a(1 To o + 1, 1 To o + 1)
ReDim b(1 To o + 1)
ReDim x(1 To n)
ReDim y(1 To n)
ReDim s(1 To o + 1)
For i = 1 To n  'Entrada de datos x,y
'    x(i) = Cells(2 + i, 1)
'    x(i) = Cells(2 + i, 5)
    x(i) = Cells(2 + i, 9)
'    y(i) = Cells(2 + i, 2)
'    y(i) = Cells(2 + i, 6)
    y(i) = Cells(2 + i, 10)
Next i
For i = 1 To o + 1 'creación matriz A, vector b
    For j = 1 To i 'matriz A coeff
        k = i + j - 2
        Sum = 0
        For L = 1 To n
            Sum = Sum + x(L) ^ k
        Next L
        a(i, j) = Sum
        a(j, i) = Sum
    Next j
    Sum = 0
    For L = 1 To n
        Sum = Sum + y(L) * x(L) ^ (i - 1)
    Next L
    b(i) = Sum
Next i
For k = 1 To o + 1 'solución del sistema por eliminación de Gauss hacia adelante
    For i = k + 1 To o + 1
        factor = a(i, k) / a(k, k)
        For j = 1 To o + 1
            a(i, j) = a(i, j) - factor * a(k, j)
        Next j
        b(i) = b(i) - factor * b(k)
    Next i
Next k
s(o + 1) = b(o + 1) / a(o + 1, o + 1) 'sustitución hacia atrás
For i = o To 1 Step -1
    Sum = b(i)
    For j = i + 1 To o + 1
        Sum = Sum - a(i, j) * s(j)
    Next j
    s(i) = Sum / a(i, i)
Next i
ym = b(1) / n 'cálculo de r2
st = 0 'cálculo de St, Sr
For i = 1 To n
    st = st + (y(i) - ym) ^ 2
Next i
sr = 0
For i = 1 To n
    Sum = y(i)
    For j = 1 To o + 1
        Sum = Sum - s(j) * x(i) ^ (j - 1)
    Next j
    sr = sr + (Sum) ^ 2
Next i
r2 = (st - sr) / st
yint = 0
For i = 1 To n 'Valor aproximado calculando polinomio
    yint = yint + s(i) * xx ^ (i - 1)
Next i
'Cells(11, 2) = yint
'Cells(11, 3) = r2
'Cells(11, 6) = yint
'Cells(11, 7) = r2
Cells(11, 10) = yint
Cells(11, 11) = r2
End Sub

