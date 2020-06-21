"""
Ésta función saca el Y en cada instánte; es necesario meter ésta función en el loop temporal e ir almacenando Y en otra lista
"""
def sistema5x5_RK4(f1, f2, f3, f4, f5, t, y1, y2, y3, y4, y5):
    
    k11 = h*f1(t, y1, y2, y3, y4, y5)
    k12 = h*f2(t, y1, y2, y3, y4, y5)
    k13 = h*f3(t, y1, y2, y3, y4, y5)
    k14 = h*f4(t, y1, y2, y3, y4, y5)
    k15 = h*f5(t, y1, y2, y3, y4, y5)
    
    k21 = h*f1(t + h/2, y1 + k11/2, y2 + k12/2, y3 + k13/2, y4 + k14/2, y5 + k15/2)
    k22 = h*f2(t + h/2, y1 + k11/2, y2 + k12/2, y3 + k13/2, y4 + k14/2, y5 + k15/2)
    k23 = h*f3(t + h/2, y1 + k11/2, y2 + k12/2, y3 + k13/2, y4 + k14/2, y5 + k15/2)
    k24 = h*f4(t + h/2, y1 + k11/2, y2 + k12/2, y3 + k13/2, y4 + k14/2, y5 + k15/2)
    k25 = h*f5(t + h/2, y1 + k11/2, y2 + k12/2, y3 + k13/2, y4 + k14/2, y5 + k15/2)
    
    k31 = h*f1(t + h/2, y1 + k21/2, y2 + k22/2, y3 + k23/2, y4 + k24/2, y5 + k25/2)
    k32 = h*f2(t + h/2, y1 + k21/2, y2 + k22/2, y3 + k23/2, y4 + k24/2, y5 + k25/2)
    k33 = h*f3(t + h/2, y1 + k21/2, y2 + k22/2, y3 + k23/2, y4 + k24/2, y5 + k25/2)
    k34 = h*f4(t + h/2, y1 + k21/2, y2 + k22/2, y3 + k23/2, y4 + k24/2, y5 + k25/2)
    k35 = h*f5(t + h/2, y1 + k21/2, y2 + k22/2, y3 + k23/2, y4 + k24/2, y5 + k25/2)
    
    k41 = h*f1(t + h, y1 + k31, y2 + k32, y3 + k33, y4 + k34, y5 + k35)
    k42 = h*f2(t + h, y1 + k31, y2 + k32, y3 + k33, y4 + k34, y5 + k35)
    k43 = h*f3(t + h, y1 + k31, y2 + k32, y3 + k33, y4 + k34, y5 + k35)
    k44 = h*f3(t + h, y1 + k31, y2 + k32, y3 + k33, y4 + k34, y5 + k35)
    k45 = h*f3(t + h, y1 + k31, y2 + k32, y3 + k33, y4 + k34, y5 + k35)
    
    y1 = y1 + (k11 + 2*k21 + 2*k31 + k41)/6
    y2 = y2 + (k12 + 2*k22 + 2*k32 + k42)/6
    y3 = y3 + (k13 + 2*k23 + 2*k33 + k43)/6
    y4 = y4 + (k14 + 2*k24 + 2*k34 + k44)/6
    y5 = y5 + (k15 + 2*k25 + 2*k35 + k45)/6
    
    Y = [y1, y2, y3, y4, y5]
    return Y
