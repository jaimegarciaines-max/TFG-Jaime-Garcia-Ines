import gurobipy as gp
from gurobipy import GRB
import math

# ============================================================
# MODELO DE VALIDACIÓN MDVRP - INSTANCIA CORDEAU P02
# ============================================================
# Este archivo es una versión simplificada del modelo exacto original.
#
# Objetivo:
#   Validar la parte estructural del modelo exacto:
#   - asignación única de clientes/orders,
#   - secuenciación,
#   - restricciones de capacidad,
#   - eliminación de subtours,
#   - cálculo de costes entre nodos.
#
# Diferencias respecto al modelo industrial original:
#   - No se usan SETUP_IN ni SETUP_OUT.
#   - El coste entre nodos se calcula como distancia euclídea.
#   - Se incluye el retorno final al depósito, por lo que se valida
#     como un MDVRP cerrado clásico.
#   - No se ejecuta ALNS.
#
# Input esperado:
#   Instancia Cordeau MDVRP en formato texto, por ejemplo p02.txt.
# ============================================================


# ------------------------------------------------------------
# PARÁMETROS DE EJECUCIÓN
# ------------------------------------------------------------

INPUT_FILE = "p02.txt"
OUTPUT_FILE = "output_validacion_p02.txt"

# Se deja la instancia completa.
# Si en algún momento se quisiera probar una subinstancia, poner por ejemplo 20.
MAX_CLIENTES = None

# Límite máximo de resolución: 4 horas = 14.400 segundos.
TIME_LIMIT = 14400


# ------------------------------------------------------------
# LECTURA DE INSTANCIA CORDEAU MDVRP
# ------------------------------------------------------------

def leer_cordeau_mdvrp(nombre_fichero):
    """
    Lee una instancia MDVRP en formato Cordeau.

    Estructura del archivo:
        Primera línea:
            tipo, vehículos_por_depósito, num_clientes, num_depósitos

        Siguientes num_depósitos líneas:
            duración_máxima, capacidad

        Siguientes num_clientes líneas:
            id, x, y, duración_servicio, demanda, ...

        Últimas num_depósitos líneas:
            id, x, y, duración_servicio, demanda, ...

    Devuelve un diccionario con:
        - tipo
        - vehículos_por_depósito
        - número de clientes
        - número de depósitos
        - coordenadas de clientes
        - coordenadas de depósitos
        - demanda de clientes
        - capacidad por depósito
    """

    with open(nombre_fichero, "r") as f:
        lineas = [line.strip() for line in f if line.strip()]

    primera = lineas[0].split()

    tipo = int(primera[0])
    vehiculos_por_deposito = int(primera[1])
    num_clientes = int(primera[2])
    num_depositos = int(primera[3])

    capacidad_deposito = {}
    for d in range(num_depositos):
        partes = lineas[1 + d].split()
        capacidad_deposito[d] = float(partes[1])

    clientes = {}
    demanda = {}

    offset_clientes = 1 + num_depositos

    for idx in range(num_clientes):
        partes = lineas[offset_clientes + idx].split()

        id_original = int(partes[0])
        x = float(partes[1])
        y = float(partes[2])
        dem = float(partes[4])

        # Se reindexan los clientes de 0 a num_clientes-1 para mantener
        # la misma lógica que en el modelo original.
        i = idx

        clientes[i] = {
            "id_original": id_original,
            "x": x,
            "y": y
        }

        demanda[i] = dem

    depositos = {}

    offset_depositos = offset_clientes + num_clientes

    for d in range(num_depositos):
        partes = lineas[offset_depositos + d].split()

        id_original = int(partes[0])
        x = float(partes[1])
        y = float(partes[2])

        depositos[d] = {
            "id_original": id_original,
            "x": x,
            "y": y
        }

    return {
        "tipo": tipo,
        "vehiculos_por_deposito": vehiculos_por_deposito,
        "num_clientes": num_clientes,
        "num_depositos": num_depositos,
        "clientes": clientes,
        "depositos": depositos,
        "demanda": demanda,
        "capacidad_deposito": capacidad_deposito
    }


# ------------------------------------------------------------
# FUNCIONES AUXILIARES
# ------------------------------------------------------------

def distancia(a, b):
    """
    Calcula la distancia euclídea entre dos puntos.
    Cada punto debe tener claves 'x' e 'y'.
    """
    return math.sqrt((a["x"] - b["x"]) ** 2 + (a["y"] - b["y"]) ** 2)


def extraer_solucion_desde_gurobi(x, orders, lines, I, F):
    """
    Convierte la solución de Gurobi en un diccionario:
        solucion[linea] = [cliente_1, cliente_2, ..., cliente_n]
    """

    solucion = {}

    for l in lines:
        inicio = None

        for j in orders:
            if x[I, j, l].x > 0.5:
                inicio = j
                break

        if inicio is None:
            solucion[l] = []
            continue

        secuencia = [inicio]
        actual = inicio

        while True:
            if x[actual, F, l].x > 0.5:
                break

            siguiente = None

            for j in orders:
                if j != actual and x[actual, j, l].x > 0.5:
                    siguiente = j
                    break

            if siguiente is None:
                break

            secuencia.append(siguiente)
            actual = siguiente

        solucion[l] = secuencia

    return solucion


def coste_solucion_mdvrp(solucion, coste, coste_inicio, coste_final):
    """
    Recalcula el coste total de una solución cerrada MDVRP:
        depósito -> clientes de la ruta -> depósito
    """

    total = 0.0

    for l, secuencia in solucion.items():
        if len(secuencia) == 0:
            continue

        total += coste_inicio[(secuencia[0], l)]

        for pos in range(len(secuencia) - 1):
            i = secuencia[pos]
            j = secuencia[pos + 1]
            total += coste[(i, j)]

        total += coste_final[(secuencia[-1], l)]

    return total


def escribir(linea, f):
    """
    Escribe una línea simultáneamente por pantalla y en archivo.
    """
    print(linea)
    f.write(linea + "\n")


# ------------------------------------------------------------
# CARGA Y PREPARACIÓN DE DATOS
# ------------------------------------------------------------

data = leer_cordeau_mdvrp(INPUT_FILE)

num_orders_total = data["num_clientes"]

if MAX_CLIENTES is None:
    num_orders = num_orders_total
else:
    num_orders = min(MAX_CLIENTES, num_orders_total)

orders = list(range(num_orders))

num_depositos = data["num_depositos"]
vehiculos_por_deposito = data["vehiculos_por_deposito"]

# En Cordeau MDVRP hay varios vehículos por depósito.
# Para p02: 4 depósitos x 2 vehículos por depósito = 8 rutas.
num_lines = num_depositos * vehiculos_por_deposito
lines = list(range(num_lines))

# Cada línea/ruta se asocia a uno de los depósitos.
# Ejemplo p02:
#   líneas 0-1 -> depósito 0
#   líneas 2-3 -> depósito 1
#   líneas 4-5 -> depósito 2
#   líneas 6-7 -> depósito 3
deposito_asociado = {
    l: l // vehiculos_por_deposito
    for l in lines
}

clientes = {
    i: data["clientes"][i]
    for i in orders
}

depositos = data["depositos"]

weight = {
    i: data["demanda"][i]
    for i in orders
}

# Cada ruta/línea hereda la capacidad del depósito al que pertenece.
capacity = {
    l: data["capacidad_deposito"][deposito_asociado[l]]
    for l in lines
}

I = "inicio"
F = "final"


# ------------------------------------------------------------
# VALIDACIÓN DE DATOS DE ENTRADA
# ------------------------------------------------------------

for i in orders:
    if i not in clientes:
        raise ValueError(f"Faltan coordenadas para el cliente/order {i}")
    if i not in weight:
        raise ValueError(f"Falta demanda para el cliente/order {i}")

for l in lines:
    if l not in capacity:
        raise ValueError(f"Falta capacidad para la línea/ruta {l}")
    d = deposito_asociado[l]
    if d not in depositos:
        raise ValueError(f"Falta depósito asociado a la línea/ruta {l}")


# ------------------------------------------------------------
# COSTES MDVRP CLÁSICO
# ------------------------------------------------------------

# Coste cliente-cliente.
coste = {}
for i in orders:
    for j in orders:
        if i != j:
            coste[(i, j)] = distancia(clientes[i], clientes[j])

# Coste depósito-cliente.
coste_inicio = {}
for l in lines:
    d = deposito_asociado[l]

    for j in orders:
        coste_inicio[(j, l)] = distancia(depositos[d], clientes[j])

# Coste cliente-depósito.
# Necesario para validar como MDVRP cerrado clásico.
coste_final = {}
for l in lines:
    d = deposito_asociado[l]

    for i in orders:
        coste_final[(i, l)] = distancia(clientes[i], depositos[d])


# ------------------------------------------------------------
# CREACIÓN DEL MODELO EXACTO
# ------------------------------------------------------------

m = gp.Model("validacion_mdvrp_cordeau_p02")

# Límite máximo de resolución: 4 horas.
m.setParam("TimeLimit", TIME_LIMIT)


# ------------------------------------------------------------
# ARCOS
# ------------------------------------------------------------

# Arcos entre clientes.
arcos_orders = [
    (i, j, l)
    for i in orders
    for j in orders
    if i != j
    for l in lines
]

# Arcos desde el depósito de cada ruta hasta el primer cliente.
arcos_inicio = [
    (I, j, l)
    for j in orders
    for l in lines
]

# Arcos desde el último cliente hasta el depósito de su ruta.
arcos_final = [
    (i, F, l)
    for i in orders
    for l in lines
]

arcos = arcos_orders + arcos_inicio + arcos_final


# ------------------------------------------------------------
# VARIABLES
# ------------------------------------------------------------

# x[i,j,l] = 1 si se usa el arco i -> j en la línea/ruta l.
x = m.addVars(arcos, vtype=GRB.BINARY, name="x")

# u[i,l] = posición auxiliar del cliente i en la ruta l.
# Se usa para restricciones MTZ de eliminación de subtours.
u = m.addVars(
    orders,
    lines,
    vtype=GRB.CONTINUOUS,
    lb=0,
    ub=num_orders,
    name="u"
)

# y[i,l] = 1 si el cliente/order i está asignado a la línea/ruta l.
y = m.addVars(
    orders,
    lines,
    vtype=GRB.BINARY,
    name="assign"
)

# carga[l] = demanda total asignada a la línea/ruta l.
carga = m.addVars(
    lines,
    vtype=GRB.CONTINUOUS,
    lb=0,
    name="load"
)


# ------------------------------------------------------------
# FUNCIÓN OBJETIVO
# ------------------------------------------------------------
# Minimizar:
#   distancia depósito -> primer cliente
# + distancias entre clientes consecutivos
# + distancia último cliente -> depósito

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
    )
    +
    gp.quicksum(
        coste_final[i, l] * x[i, F, l]
        for i in orders
        for l in lines
    ),
    GRB.MINIMIZE
)


# ------------------------------------------------------------
# RESTRICCIONES
# ------------------------------------------------------------

# Cada cliente/order tiene exactamente una salida:
# hacia otro cliente o hacia el nodo final.
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


# Cada cliente/order tiene exactamente una entrada:
# desde otro cliente o desde el depósito de una ruta.
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


# Conservación de flujo por línea/ruta.
# Si un cliente entra en una ruta, también debe salir de esa misma ruta.
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


# Cada línea/ruta puede tener como mucho un inicio y un final.
# Esto permite que algunas rutas queden vacías.
for l in lines:
    m.addConstr(
        gp.quicksum(x[I, j, l] for j in orders) <= 1,
        name=f"un_inicio_linea_{l}"
    )

    m.addConstr(
        gp.quicksum(x[i, F, l] for i in orders) <= 1,
        name=f"un_final_linea_{l}"
    )


# Definición de asignación:
# y[i,l] = 1 si el cliente/order i entra en la ruta l.
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


# Carga total de cada línea/ruta.
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


# Capacidad máxima de cada línea/ruta.
for l in lines:
    m.addConstr(
        carga[l] <= capacity[l],
        name=f"capacidad_linea_{l}"
    )


# Restricciones MTZ por línea/ruta para evitar subtours.
n = num_orders

for i in orders:
    for j in orders:
        if i != j:
            for l in lines:
                m.addConstr(
                    u[i, l] - u[j, l] + n * x[i, j, l] <= n - 1,
                    name=f"mtz_{i}_{j}_{l}"
                )


# ------------------------------------------------------------
# RESOLUCIÓN
# ------------------------------------------------------------

m.optimize()


# ------------------------------------------------------------
# OUTPUT
# ------------------------------------------------------------

with open(OUTPUT_FILE, "w") as f:
    escribir("VALIDACIÓN MDVRP - INSTANCIA CORDEAU P02", f)
    escribir(f"Archivo de entrada: {INPUT_FILE}", f)
    escribir(f"Clientes usados: {num_orders}", f)
    escribir(f"Depósitos originales: {num_depositos}", f)
    escribir(f"Vehículos por depósito: {vehiculos_por_deposito}", f)
    escribir(f"Líneas/rutas usadas: {num_lines}", f)
    escribir(f"Depósito asociado a cada línea: {deposito_asociado}", f)
    escribir(f"Capacidad por línea/ruta: {capacity}", f)
    escribir(f"Límite de tiempo: {TIME_LIMIT} segundos", f)
    escribir("", f)

    if m.SolCount > 0:
        if m.status == GRB.OPTIMAL:
            escribir("Solución óptima encontrada", f)
        elif m.status == GRB.TIME_LIMIT:
            escribir("Solución factible encontrada, pero se alcanzó el límite de tiempo", f)
        else:
            escribir("Solución factible encontrada, pero no se ha demostrado optimalidad", f)
            escribir(f"Estado Gurobi: {m.status}", f)

        escribir(f"Valor objetivo Gurobi: {m.objVal}", f)

        try:
            escribir(f"Gap final: {m.MIPGap}", f)
        except Exception:
            escribir("Gap final no disponible", f)

        escribir(f"Tiempo de resolución Gurobi: {m.Runtime} segundos", f)
        escribir("", f)

        solucion = extraer_solucion_desde_gurobi(x, orders, lines, I, F)
        coste_recalculado = coste_solucion_mdvrp(
            solucion,
            coste,
            coste_inicio,
            coste_final
        )

        escribir(f"Coste recalculado desde secuencias: {coste_recalculado}", f)
        escribir(f"Diferencia objetivo - recalc: {abs(m.objVal - coste_recalculado)}", f)
        escribir("", f)

        escribir("Carga por línea/ruta:", f)

        for l in lines:
            d = deposito_asociado[l]
            escribir(
                f"Línea/ruta {l} "
                f"(depósito original {depositos[d]['id_original']}): "
                f"{carga[l].x} / {capacity[l]}",
                f
            )

        escribir("", f)
        escribir("Secuencias:", f)

        for l in lines:
            secuencia = solucion[l]
            ids_originales = [clientes[i]["id_original"] for i in secuencia]

            d = deposito_asociado[l]
            deposito_original = depositos[d]["id_original"]

            escribir(
                f"\nLínea/ruta {l} "
                f"(depósito original {deposito_original}):",
                f
            )

            if len(secuencia) == 0:
                escribir("  Ruta vacía", f)
            else:
                escribir(
                    "  depot -> "
                    + " -> ".join(map(str, ids_originales))
                    + " -> depot",
                    f
                )

    else:
        escribir("No se encontró ninguna solución factible", f)
        escribir(f"Estado Gurobi: {m.status}", f)
        escribir(f"Tiempo de resolución: {m.Runtime} segundos", f)
