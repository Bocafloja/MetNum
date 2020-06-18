Attribute VB_Name = "EDO_met"
Function fx_euler1(x) 'Función dy/dx Euler
'fx_euler = -2 * x ^ 3 + 12 * x ^ 2 - 20 * x + 8.5
fx_euler1 = (175 - x - 0.15 * 12 * x ^ 2) / 12
End Function

Function fx_euler2(x) 'segunda derivada despejada de la edo a analizar
fx_euler2 = -2 * x ^ 3 + 12 * x ^ 2 - 20 * x + 8.5
'fx_euler2 = (175 - x - 0.15 * 12 * x ^ 2) / 12
'fx_euler2 = -1.5 * x
End Function
Sub EULER2_met()
xi = Cells(9, 2)
xf = Cells(10, 2)
dx = Cells(11, 2)
y0 = Cells(12, 2)
dy0 = Cells(13, 2)
nc = (xf - xi) / dx

xx = xi

q0 = y0
i0 = dy0
For i = 1 To nc
    q0 = q0 + i0 * dx
    i0 = i0 + fx_euler2(xx) * dx
    xx = xx + dx
'
    Cells(8 + i, 4) = xx
    Cells(8 + i, 5) = q0
    Cells(8 + i, 6) = i0
    If y0 < 0 Then
        Exit For
    End If
Next i
'Wend
End Sub
Sub EULER_met()
xi = Cells(9, 2)
xf = Cells(10, 2)
dx = Cells(11, 2)
y0 = Cells(12, 2)

nc = (xf - xi) / dx

xx = xi

For i = 1 To nc
    dy = fx_euler1(xx)
    y0 = y0 + dy * dx
    xx = xx + dx
'
    Cells(8 + i, 4) = xx
    Cells(8 + i, 5) = y0
    If y0 < 0 Then
        Exit For
    End If
Next i
'Wend
End Sub
Function fx_rk(x, y)
'fx_rk =
End Function
Sub RK41()
xi = Cells(9, 2)
xf = Cells(10, 2)
dx = Cells(11, 2)
y0 = Cells(12, 2)
nc = (xf - xi) / dx
xx = xi
yy = y0

For i = 1 To nc
    k1 = fx_rk(xx, yy)
    k2 = fx_rk(xx + 0.5 * dx, yy + 0.5 * dx * k1)
    k3 = fx_rk(xx + 0.5 * dx, yy + 0.5 * dx * k2)
    k4 = fc_rk(xx + dx, yy * dx * k3)
    yy = yy + (1 / 6) * h * (k1 + 2 * k2 + 2 * k3 + k4)
    xx = xx + dx
    Cells(8 + i, 4) = xx
    Cells(8 + i, 5) = yy
    If y0 < 0 Then
        Exit For
    End If
Next i
'Wend
End Sub

Function f2x_rk42(x, yp, y) 'segunda derivada despejada de la edo a analizar
'f2x_rk4 =
End Function

Sub RK42()
xi = Cells(9, 2)
xf = Cells(10, 2)
dx = Cells(11, 2)
y0 = Cells(12, 2)
dy0 = Cells(13, 2)
nc = (xf - xi) / dx
Dim funcion(2)
Dim z0(2)
Dim k1(2)
Dim k2(2)
Dim k3(2)
Dim k4(2)


xx = xi
z0(1) = y0
z0(2) = dy0
funcion(1) = w
funcion(2) = f2_rk42(x, yp, y)

For i = 1 To nc
    k1(1) = z0(2)
    k1(2) = f2_rk42(xi, z0(2), z0(1))
    
    k2(1) = z0(2) + (dx / 2) * k1(2)
    k2(2) = f2_rk42(xi + dx / 2, z0(2) + dx / 2 * k1(2), z0(1) + dx / 2 * k1(1))
    
    k3(1) = z0(2) + (dx / 2) * k2(2)
    k3(2) = f2_rk42(xi + dx / 2, z0(2) + dx / 2 * k2(2), z0(1) + dx / 2 * k2(1))
    
    k4(1) = z0(2) + dx * k3(2)
    k4(2) = f2_rk42(xi + dx, z0(2) + dx * k3(2), z0(1) + dx * k3(1))
    
    z0(1) = z0(1) + (1 / 6) * dx * (k1(1) + 2 * k2(1) + 2 * k3(1) + k4(1))
    
    z0(2) = z0(2) + (1 / 6) * dx * (k1(2) + 2 * k2(2) + 2 * k3(2) + k4(2))
    xx = xx + dx
    Cells(8 + i, 4) = xx
    Cells(8 + i, 5) = z0(1)
    Cells(8 + i, 6) = z0(2)
        Next i
        
'Wend
End Sub

Function edo1(t, x, y, z) 'ecuacion 1
'edo1 =
End Function

Function edo2(t, x, y, z) 'ecuacion 2
'edo2 =
End Function

Function edo3(t, x, y, z) 'ecuacion 3
'edo3=
End Function


Sub sisedo_met()
xi = Cells(9, 2)
yi = Cells(8, 2)
zi = Cells(7, 2)
ti = Cells(10, 2)
tf = Cells(11, 2)
dt = Cells(12, 2)
nc = (tf - ti) / dt
Dim x0(3)
Dim funcion(3)
t0 = ti
For i = 1 To nc
    funcion(1) = edo1(ti, x0(1), x0(2), x0(3))
    funcion(2) = edo2(ti, x0(1), x0(2), x0(3))
    funcion(3) = edo3(ti, x0(1), x0(2), x0(3))
       
    x0(1) = x0(1) + dt * funcion(1)
    x0(2) = x0(2) + dt * funcion(2)
    x0(3) = x0(3) + dt * funcion(3)
    
    t0 = t0 + dt
    Cells(8 + i, 4) = t0
    Cells(8 + i, 5) = x0(1)
    Cells(8 + i, 6) = x0(2)
    Cells(8 + i, 7) = x0(3)
    
     
Next i
'Wend
End Sub

Sub disparo()
yi = Cells(4, 2)
yf = Cells(5, 2)
xi = Cells(6, 2)
xf = Cells(7, 2)
disp1 = Cells(8, 2)
disp2 = Cells(9, 2)
dx = Cells(10, 2)

nc = (xf - xi) / dx

Dim zo(2)
Dim k1(2)
Dim k2(2)
Dim k3(2)
Dim k4(2)


xx = xi
zo(1) = y0
z0(2) = disp1

For i = 1 To nc
    k1(1) = z0(2)
    k1(2) = f2_rk42(xi, z0(2), z0(1))
    
    k2(1) = z0(2) + (dx / 2) * k1(2)
    k2(2) = f2_rk42(xi + dx / 2, z0(2) + dx / 2 * k1(2), z0(1) + dx / 2 * k1(1))
    
    k3(1) = z0(2) + (dx / 2) * k2(2)
    k3(2) = f2_rk42(xi + dx / 2, z0(2) + dx / 2 * k2(2), z0(1) + dx / 2 * k2(1))
    
    k4(1) = z0(2) + dx * k3(2)
    k4(2) = f2_rk42(xi + dx, z0(2) + dx * k3(2), z0(1) + dx * k3(1))
    
    z0(1) = z0(1) + (1 / 6) * dx * (k1(1) + 2 * k2(1) + 2 * k3(1) + k4(1))
    
    z0(2) = z0(2) + (1 / 6) * dx * (k1(2) + 2 * k2(2) + 2 * k3(2) + k4(2))
    xx = xx + dx
    Cells(8 + i, 4) = xx
    Cells(8 + i, 5) = z0(1)
    Cells(8 + i, 6) = z0(2)
        Next i

xx = xi
zo(1) = y0
z0(2) = disp2

For i = 1 To nc
    k1(1) = z0(2)
    k1(2) = f2_rk42(xi, z0(2), z0(1))
    
    k2(1) = z0(2) + (dx / 2) * k1(2)
    k2(2) = f2_rk42(xi + dx / 2, z0(2) + dx / 2 * k1(2), z0(1) + dx / 2 * k1(1))
    
    k3(1) = z0(2) + (dx / 2) * k2(2)
    k3(2) = f2_rk42(xi + dx / 2, z0(2) + dx / 2 * k2(2), z0(1) + dx / 2 * k2(1))
    
    k4(1) = z0(2) + dx * k3(2)
    k4(2) = f2_rk42(xi + dx, z0(2) + dx * k3(2), z0(1) + dx * k3(1))
    
    z0(1) = z0(1) + (1 / 6) * dx * (k1(1) + 2 * k2(1) + 2 * k3(1) + k4(1))
    
    z0(2) = z0(2) + (1 / 6) * dx * (k1(2) + 2 * k2(2) + 2 * k3(2) + k4(2))
    xx = xx + dx
    Cells(8 + i, 8) = xx
    Cells(8 + i, 9) = z0(1)
    Cells(8 + i, 10) = z0(2)
        Next i
sol1 = Cells(8 + np, 5)
sol2 = Cells(8 + np, 9)
xx = xi
z0(1) = y0
z0(2) = disp1 + (yf - sol1) * (disp2 - disp1) / (sol2 - sol1)
Cells(11, 2) = z0(2)
For i = 1 To nc
    k1(1) = z0(2)
    k1(2) = f2_rk42(xi, z0(2), z0(1))
    
    k2(1) = z0(2) + (dx / 2) * k1(2)
    k2(2) = f2_rk42(xi + dx / 2, z0(2) + dx / 2 * k1(2), z0(1) + dx / 2 * k1(1))
    
    k3(1) = z0(2) + (dx / 2) * k2(2)
    k3(2) = f2_rk42(xi + dx / 2, z0(2) + dx / 2 * k2(2), z0(1) + dx / 2 * k2(1))
    
    k4(1) = z0(2) + dx * k3(2)
    k4(2) = f2_rk42(xi + dx, z0(2) + dx * k3(2), z0(1) + dx * k3(1))
    
    z0(1) = z0(1) + (1 / 6) * dx * (k1(1) + 2 * k2(1) + 2 * k3(1) + k4(1))
    
    z0(2) = z0(2) + (1 / 6) * dx * (k1(2) + 2 * k2(2) + 2 * k3(2) + k4(2))
    xx = xx + dx
    Cells(8 + i, 12) = xx
    Cells(8 + i, 13) = z0(1)
    Cells(8 + i, 14) = z0(2)
        Next i
        
End Sub
