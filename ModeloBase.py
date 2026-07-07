import gurobipy as gp
from gurobipy import GRB


# LECTURA DE DATOS
#funcion para leer y guardar los datos de un fichero de texto con el formato dado en el enunciado
def leer_datos(nombre_fichero):
    data = {}

    with open(nombre_fichero, "r") as f:
        lineas = f.readlines()

    i = 0
    while i < len(lineas):
        linea = lineas[i].strip()

        if linea == "" or linea.startswith("#"):
            i += 1
            continue

        if linea.startswith("NUM_ORDERS"):
            data["num_orders"] = int(linea.split()[1])

        elif linea.startswith("NUM_LINES"):
            data["num_lines"] = int(linea.split()[1])

        elif linea == "WEIGHT":
            i += 1
            weight = {}
            while i < len(lineas) and lineas[i].strip() != "":
                idx, val = lineas[i].split()
                weight[int(idx)] = float(val)
                i += 1
            data["weight"] = weight
            continue

        elif linea == "SETUP_IN":
            i += 1
            setup_in = {}
            while i < len(lineas) and lineas[i].strip() != "":
                idx, val = lineas[i].split()
                setup_in[int(idx)] = float(val)
                i += 1
            data["setup_in"] = setup_in
            continue

        elif linea == "SETUP_OUT":
            i += 1
            setup_out = {}
            while i < len(lineas) and lineas[i].strip() != "":
                idx, val = lineas[i].split()
                setup_out[int(idx)] = float(val)
                i += 1
            data["setup_out"] = setup_out
            continue

        elif linea == "INITIAL_STATE":
            i += 1
            initial = {}
            while i < len(lineas) and lineas[i].strip() != "":
                idx, val = lineas[i].split()
                initial[int(idx)] = float(val)
                i += 1
            data["initial"] = initial
            continue

        elif linea == "CAPACITY":
            i += 1
            cap = {}
            while i < len(lineas) and lineas[i].strip() != "":
                idx, val = lineas[i].split()
                cap[int(idx)] = float(val)
                i += 1
            data["capacity"] = cap
            continue

        i += 1

    return data


# CARGAR DATOS

data = leer_datos("input.txt") #llama a la funcion

#extrae los datos del diccionario y los guarda en variables para usarlos mas facilmente
num_orders = data["num_orders"] 
num_lines = data["num_lines"]

#crea listas con los indices de las ordenes y las lineas para usarlos en los bucles
orders = list(range(num_orders))
lines = list(range(num_lines))

#extrae los datos del diccionario y los guarda en variables para usarlos mas facilmente
weight = data["weight"]
setup_in = data["setup_in"]
setup_out = data["setup_out"]
initial_state = data["initial"]
capacity = data["capacity"]

I = "inicio"
F = "final"


# VALIDACIÓN DE DATOS
#compruba que todos los orders tengan los datos necesarios
for i in orders:
    if i not in weight:
        raise ValueError(f"Falta WEIGHT para order {i}")
    if i not in setup_in:
        raise ValueError(f"Falta SETUP_IN para order {i}")
    if i not in setup_out:
        raise ValueError(f"Falta SETUP_OUT para order {i}")

#compruba que todas las lineas tengan los datos necesarios
for l in lines:
    if l not in initial_state:
        raise ValueError(f"Falta INITIAL_STATE para línea {l}")
    if l not in capacity:
        raise ValueError(f"Falta CAPACITY para línea {l}")


# COSTES

coste = {}
for i in orders:
    for j in orders:
        if i != j:
            coste[(i, j)] = abs(setup_out[i] - setup_in[j])

coste_inicio = {}
for l in lines:
    for j in orders:
        coste_inicio[(j, l)] = abs(initial_state[l] - setup_in[j])


# MODELO
#crea el modelo de Gurobi
m = gp.Model("modelo_general")


# PARÁMETROS DE PARADA PERSONALIZADA

MIN_TIME = 180       # 3 minutos
MAX_TIME = 300       # 5 minutos
TARGET_GAP = 0.05    # 5%

#esta funcion se ejecuta mientras gurobi esta resolviendo el modelo
def parada_personalizada(model, where):
    if where == GRB.Callback.MIP:  #MIP:Programacion entera mixta
        runtime = model.cbGet(GRB.Callback.RUNTIME) #obtiene el tiempo transcurrido
        
        # Si se ha superado el tiempo máximo, se detiene el modelo
        if runtime >= MAX_TIME:
            model._parada_por_tiempo_maximo = True
            model.terminate()
            return

        obj_best = model.cbGet(GRB.Callback.MIP_OBJBST) #obtiene el mejor valor objetivo encontrado hasta el momento
        obj_bound = model.cbGet(GRB.Callback.MIP_OBJBND) #obtiene la mejor cota conocida

        if obj_best < GRB.INFINITY and abs(obj_best) > 1e-9:
            gap = abs(obj_best - obj_bound) / abs(obj_best) #calcula el gap 

            # Si se ha superado el tiempo mínimo y el gap es menor o igual al objetivo, se detiene el modelo
            if runtime >= MIN_TIME and gap <= TARGET_GAP:
                model._parada_por_gap_y_tiempo = True
                model.terminate()


m._parada_por_gap_y_tiempo = False
m._parada_por_tiempo_maximo = False


# ARCOS
#crea todos los arcos posibles y los guarda en una lista para luego crear las variables de decisión solo para esos arcos
arcos_orders = [
    (i, j, l)
    for i in orders
    for j in orders
    if i != j
    for l in lines
]

arcos_inicio = [
    (I, j, l)
    for j in orders
    for l in lines
]

arcos_final = [
    (i, F, l)
    for i in orders
    for l in lines
]

arcos = arcos_orders + arcos_inicio + arcos_final


# VARIABLES

x = m.addVars(arcos, vtype=GRB.BINARY, name="x") #x[i,j,l] = 1 si el arco de la orden i a la orden j en la línea l está activo, 0 en caso contrario

#variable auxiliar de las restricciones de MTZ para evitar subciclos. u[i,l] representa la posición de la orden i en la secuencia de la línea l
u = m.addVars(
    orders,
    lines,
    vtype=GRB.CONTINUOUS,
    lb=0,
    ub=num_orders,
    name="u"
)

#variable auxiliar para calcular la carga total de cada línea. y[i,l] = 1 si la orden i está asignada a la línea l, 0 en caso contrario
y = m.addVars(
    orders,
    lines,
    vtype=GRB.BINARY,
    name="assign"
)

carga = m.addVars(
    lines,
    vtype=GRB.CONTINUOUS, #continua porque las toneladas pueden ser decimales
    lb=0,
    name="load"
)


# FUNCIÓN OBJETIVO

m.setObjective(
    gp.quicksum(
        coste[i, j] * x[i, j, l]
        for i in orders
        for j in orders
        if i != j
        for l in lines
    )
    +
    gp.quicksum(
        coste_inicio[j, l] * x[I, j, l]
        for j in orders
        for l in lines
    ),
    GRB.MINIMIZE
)


# RESTRICCIONES
# Salida única
for i in orders:
    m.addConstr(
        gp.quicksum(
            x[i, j, l]
            for j in orders
            if i != j
            for l in lines
        )
        +
        gp.quicksum(
            x[i, F, l]
            for l in lines
        )
        == 1,
        name=f"salida_unica_{i}"
    )

# Entrada única
for j in orders:
    m.addConstr(
        gp.quicksum(
            x[i, j, l]
            for i in orders
            if i != j
            for l in lines
        )
        +
        gp.quicksum(
            x[I, j, l]
            for l in lines
        )
        == 1,
        name=f"entrada_unica_{j}"
    )

# Conservación de flujo
for l in lines:
    for k in orders:
        m.addConstr(
            gp.quicksum(
                x[i, k, l]
                for i in orders
                if i != k
            )
            +
            x[I, k, l]
            ==
            gp.quicksum(
                x[k, j, l]
                for j in orders
                if j != k
            )
            +
            x[k, F, l],
            name=f"flujo_linea_{l}_order_{k}"
        )

# Como mucho un inicio y un final por línea
for l in lines:
    m.addConstr(
        gp.quicksum(x[I, j, l] for j in orders) <= 1,
        name=f"un_inicio_linea_{l}"
    )

    m.addConstr(
        gp.quicksum(x[i, F, l] for i in orders) <= 1,
        name=f"un_final_linea_{l}"
    )

# Definición de asignación. Define si un order está en una linea.
for l in lines:
    for i in orders:
        m.addConstr(
            y[i, l]
            ==
            gp.quicksum(
                x[h, i, l]
                for h in orders
                if h != i
            )
            +
            x[I, i, l],
            name=f"def_assign_{i}_{l}"
        )

# Carga en toneladas
for l in lines:
    m.addConstr(
        carga[l]
        ==
        gp.quicksum(
            weight[i] * y[i, l]
            for i in orders
        ),
        name=f"def_load_{l}"
    )

# Capacidad máxima por línea
for l in lines:
    m.addConstr(
        carga[l] <= capacity[l],
        name=f"capacidad_linea_{l}"
    )

# MTZ por línea
n = num_orders

for i in orders:
    for j in orders:
        if i != j:
            for l in lines:
                m.addConstr(
                    u[i, l] - u[j, l] + n * x[i, j, l] <= n - 1,
                    name=f"mtz_{i}_{j}_{l}"
                )


# RESOLVER

m.optimize(parada_personalizada)


# OUTPUT: ARCHIVO + PANTALLA
#la funcion escribir escribe cada linea en el archivo de salida y por pantalla
def escribir(linea, f):
    print(linea)
    f.write(linea + "\n")


with open("output.txt", "w") as f:

    if m.SolCount > 0:

        if m.status == GRB.OPTIMAL:
            escribir("Solución óptima encontrada", f)

        elif m._parada_por_gap_y_tiempo:
            escribir("Solución válida encontrada", f)
            escribir("No se ha demostrado optimalidad.", f)
            escribir("El modelo se ha detenido porque han pasado al menos 3 minutos y el gap es menor o igual al 5%.", f)

        elif m._parada_por_tiempo_maximo:
            escribir("Solución válida encontrada", f)
            escribir("No se ha demostrado optimalidad.", f)
            escribir("El modelo se ha detenido porque se ha alcanzado el límite máximo de 5 minutos.", f)

        else:
            escribir("Solución válida encontrada", f)
            escribir("No se ha demostrado optimalidad.", f)
            escribir(f"Estado de parada de Gurobi: {m.status}", f)

        escribir(f"Valor objetivo: {m.objVal}", f)

        try:
            escribir(f"Gap final: {m.MIPGap}", f)
        except Exception:
            escribir("Gap final no disponible", f)

        escribir(f"Tiempo de resolución: {m.Runtime} segundos", f)
        escribir("", f)

        escribir("Carga por línea:", f)
        for l in lines:
            escribir(f"Línea {l}: {carga[l].x} toneladas / capacidad {capacity[l]}", f)

        for l in lines:
            escribir(f"\nLínea {l}:", f)

            inicio = None
            for j in orders:
                if x[I, j, l].x > 0.5: 
                    inicio = j
                    break

            if inicio is None:
                escribir("  Línea vacía", f)
                continue

            secuencia = [inicio]
            actual = inicio

            while True:
                if x[actual, F, l].x > 0.5: #si desde la bobina actual se va al nodo final, termina la línea
                    break

                siguiente = None

                for j in orders:
                    if j != actual and x[actual, j, l].x > 0.5: #
                        siguiente = j
                        break

                if siguiente is None:
                    break

                secuencia.append(siguiente)
                actual = siguiente

            escribir( #
                "  start -> "
                + " -> ".join(map(str, secuencia)) 
                + " -> end",
                f
            )

    else:
        escribir("No se encontró ninguna solución factible", f)
        escribir(f"Estado de parada de Gurobi: {m.status}", f)
        escribir(f"Tiempo de resolución: {m.Runtime} segundos", f)