import gurobipy as gp
from gurobipy import GRB
import math

# ============================================================
# MODELO DE VALIDACIÓN CVRP - INSTANCIA E-n22-k4
# ============================================================
# Este archivo adapta la estructura del modelo exacto original para validar
# la parte principal de la formulación mediante una instancia CVRP estándar.
#
# Objetivo:
#   Validar la estructura de routing capacitado:
#   - asignación única de clientes/orders,
#   - secuenciación de rutas,
#   - restricciones de capacidad,
#   - conservación de flujo,
#   - eliminación de subtours mediante MTZ,
#   - cálculo de costes por distancia euclídea.
#
# Diferencias respecto al modelo industrial original:
#   - No se usan SETUP_IN ni SETUP_OUT.
#   - El coste entre nodos se calcula como distancia euclídea.
#   - Solo existe un depósito.
#   - Todas las rutas/vehículos salen del mismo depósito y vuelven a él.
#   - No se ejecuta ALNS.
#
# Instancia utilizada:
#   E-n22-k4.txt, descargada de CVRPLIB/VRPLIB:
#   CVRP con 21 clientes, 1 depósito, 4 vehículos/rutas y capacidad 6000.
#   Valor óptimo conocido para la instancia E-n22-k4: 375.
# ============================================================


# ------------------------------------------------------------
# PARÁMETROS DE EJECUCIÓN
# ------------------------------------------------------------

INPUT_FILE = "E-n22-k4.txt"
OUTPUT_FILE = "output_validacion_E-n22-k4.txt"

# Valor óptimo conocido de la instancia E-n22-k4.
OPTIMO_CONOCIDO = 375

# Número máximo de vehículos/rutas de la instancia E-n22-k4.
# El nombre k4 indica 4 vehículos.
NUM_VEHICULOS = 4

# Límite máximo de resolución.
# Para esta instancia pequeña debería bastar con mucho menos, pero se deja amplio.
# TIME_LIMIT = 3600


# ------------------------------------------------------------
# LECTURA DE INSTANCIA CVRPLIB / VRPLIB
# ------------------------------------------------------------

def leer_cvrplib(nombre_fichero):
    """
    Lee una instancia CVRP en formato CVRPLIB/VRPLIB.

    Se extraen:
        - nombre de la instancia,
        - dimensión total,
        - capacidad de los vehículos,
        - coordenadas de los nodos,
        - demandas,
        - depósito.

    En este formato:
        - la dimensión incluye depósito + clientes,
        - las coordenadas están en NODE_COORD_SECTION,
        - las demandas están en DEMAND_SECTION,
        - el depósito está en DEPOT_SECTION.
    """

    with open(nombre_fichero, "r") as f:
        lineas = [line.strip() for line in f if line.strip()]

    nombre = None
    dimension = None
    capacity = None

    coords_originales = {}
    demandas_originales = {}
    deposito_original = None

    seccion = None

    for linea in lineas:
        if linea.startswith("EOF"):
            break

        if linea.startswith("NODE_COORD_SECTION"):
            seccion = "coords"
            continue

        if linea.startswith("DEMAND_SECTION"):
            seccion = "demandas"
            continue

        if linea.startswith("DEPOT_SECTION"):
            seccion = "deposito"
            continue

        if seccion is None:
            if ":" in linea:
                clave, valor = [p.strip() for p in linea.split(":", 1)]

                if clave == "NAME":
                    nombre = valor
                elif clave == "DIMENSION":
                    dimension = int(valor)
                elif clave == "CAPACITY":
                    capacity = float(valor)

            continue

        if seccion == "coords":
            partes = linea.split()
            nodo = int(partes[0])
            x = float(partes[1])
            y = float(partes[2])

            coords_originales[nodo] = {
                "id_original": nodo,
                "x": x,
                "y": y
            }

        elif seccion == "demandas":
            partes = linea.split()
            nodo = int(partes[0])
            demanda = float(partes[1])
            demandas_originales[nodo] = demanda

        elif seccion == "deposito":
            valor = int(linea.split()[0])

            if valor == -1:
                seccion = None
            else:
                deposito_original = valor

    if nombre is None:
        raise ValueError("No se ha encontrado NAME en el archivo.")

    if dimension is None:
        raise ValueError("No se ha encontrado DIMENSION en el archivo.")

    if capacity is None:
        raise ValueError("No se ha encontrado CAPACITY en el archivo.")

    if deposito_original is None:
        raise ValueError("No se ha encontrado depósito en DEPOT_SECTION.")

    if deposito_original not in coords_originales:
        raise ValueError("El depósito no tiene coordenadas asociadas.")

    clientes_originales = [
        nodo for nodo in sorted(coords_originales.keys())
        if nodo != deposito_original
    ]

    clientes = {}
    demanda = {}

    for idx, nodo_original in enumerate(clientes_originales):
        clientes[idx] = {
            "id_original": nodo_original,
            "x": coords_originales[nodo_original]["x"],
            "y": coords_originales[nodo_original]["y"]
        }

        demanda[idx] = demandas_originales[nodo_original]

    deposito = {
        "id_original": deposito_original,
        "x": coords_originales[deposito_original]["x"],
        "y": coords_originales[deposito_original]["y"]
    }

    return {
        "nombre": nombre,
        "dimension": dimension,
        "num_clientes": len(clientes),
        "clientes": clientes,
        "deposito": deposito,
        "demanda": demanda,
        "capacidad": capacity
    }


# ------------------------------------------------------------
# FUNCIONES AUXILIARES
# ------------------------------------------------------------

def distancia(a, b):
    """
    Calcula la distancia EUC_2D estándar de CVRPLIB.

    En las instancias CVRPLIB con EDGE_WEIGHT_TYPE = EUC_2D,
    la distancia se calcula como:

        int(sqrt((x_i-x_j)^2 + (y_i-y_j)^2) + 0.5)

    Es decir, se redondea al entero más próximo siguiendo la
    convención clásica de TSPLIB/CVRPLIB.
    """
    return int(
        math.sqrt(
            (a["x"] - b["x"]) ** 2 +
            (a["y"] - b["y"]) ** 2
        ) + 0.5
    )


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


def coste_solucion_cvrp(solucion, coste, coste_inicio, coste_final):
    """
    Recalcula el coste total de una solución cerrada CVRP:
        depósito -> clientes de la ruta -> depósito
    """

    total = 0.0

    for l, secuencia in solucion.items():
        if len(secuencia) == 0:
            continue

        total += coste_inicio[secuencia[0]]

        for pos in range(len(secuencia) - 1):
            i = secuencia[pos]
            j = secuencia[pos + 1]
            total += coste[(i, j)]

        total += coste_final[secuencia[-1]]

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

data = leer_cvrplib(INPUT_FILE)

num_orders = data["num_clientes"]
orders = list(range(num_orders))

num_lines = NUM_VEHICULOS
lines = list(range(num_lines))

clientes = data["clientes"]
deposito = data["deposito"]

weight = data["demanda"]

capacity = {
    l: data["capacidad"]
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


# ------------------------------------------------------------
# COSTES CVRP CLÁSICO
# ------------------------------------------------------------

coste = {}
for i in orders:
    for j in orders:
        if i != j:
            coste[(i, j)] = distancia(clientes[i], clientes[j])

coste_inicio = {}
for j in orders:
    coste_inicio[j] = distancia(deposito, clientes[j])

coste_final = {}
for i in orders:
    coste_final[i] = distancia(clientes[i], deposito)


# ------------------------------------------------------------
# CREACIÓN DEL MODELO EXACTO
# ------------------------------------------------------------

m = gp.Model("validacion_cvrp_E_n22_k4")
# m.setParam("TimeLimit", TIME_LIMIT)


# ------------------------------------------------------------
# ARCOS
# ------------------------------------------------------------

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


# ------------------------------------------------------------
# VARIABLES
# ------------------------------------------------------------

x = m.addVars(arcos, vtype=GRB.BINARY, name="x")

u = m.addVars(
    orders,
    lines,
    vtype=GRB.CONTINUOUS,
    lb=0,
    ub=num_orders,
    name="u"
)

y = m.addVars(
    orders,
    lines,
    vtype=GRB.BINARY,
    name="assign"
)

carga = m.addVars(
    lines,
    vtype=GRB.CONTINUOUS,
    lb=0,
    name="load"
)


# ------------------------------------------------------------
# FUNCIÓN OBJETIVO
# ------------------------------------------------------------

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
        coste_inicio[j] * x[I, j, l]
        for j in orders
        for l in lines
    )
    +
    gp.quicksum(
        coste_final[i] * x[i, F, l]
        for i in orders
        for l in lines
    ),
    GRB.MINIMIZE
)


# ------------------------------------------------------------
# RESTRICCIONES
# ------------------------------------------------------------

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

for l in lines:
    m.addConstr(
        gp.quicksum(x[I, j, l] for j in orders) <= 1,
        name=f"un_inicio_linea_{l}"
    )

    m.addConstr(
        gp.quicksum(x[i, F, l] for i in orders) <= 1,
        name=f"un_final_linea_{l}"
    )

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

for l in lines:
    m.addConstr(
        carga[l] <= capacity[l],
        name=f"capacidad_linea_{l}"
    )

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
    escribir("VALIDACIÓN CVRP - INSTANCIA E-n22-k4", f)
    escribir(f"Archivo de entrada: {INPUT_FILE}", f)
    escribir(f"Nombre instancia: {data['nombre']}", f)
    escribir(f"Nodos totales en archivo: {data['dimension']}", f)
    escribir(f"Clientes usados: {num_orders}", f)
    escribir(f"Depósito original: {deposito['id_original']}", f)
    escribir(f"Vehículos/rutas usadas: {num_lines}", f)
    escribir(f"Capacidad por vehículo/ruta: {data['capacidad']}", f)
    # escribir(f"Límite de tiempo: {TIME_LIMIT} segundos", f)
    escribir(f"Óptimo conocido: {OPTIMO_CONOCIDO}", f)
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

        desviacion = 100 * (m.objVal - OPTIMO_CONOCIDO) / OPTIMO_CONOCIDO
        escribir(f"Desviación respecto al óptimo conocido: {desviacion} %", f)

        try:
            escribir(f"Gap final Gurobi: {m.MIPGap}", f)
        except Exception:
            escribir("Gap final Gurobi no disponible", f)

        escribir(f"Tiempo de resolución Gurobi: {m.Runtime} segundos", f)
        escribir("", f)

        solucion = extraer_solucion_desde_gurobi(x, orders, lines, I, F)
        coste_recalculado = coste_solucion_cvrp(
            solucion,
            coste,
            coste_inicio,
            coste_final
        )

        escribir(f"Coste recalculado desde secuencias: {coste_recalculado}", f)
        escribir(f"Diferencia objetivo - recalc: {abs(m.objVal - coste_recalculado)}", f)
        escribir("", f)

        escribir("Carga por vehículo/ruta:", f)

        for l in lines:
            escribir(
                f"Ruta {l}: {carga[l].x} / {capacity[l]}",
                f
            )

        escribir("", f)
        escribir("Secuencias:", f)

        for l in lines:
            secuencia = solucion[l]
            ids_originales = [clientes[i]["id_original"] for i in secuencia]

            escribir(f"\nRuta {l}:", f)

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
