import zipfile
from scipy.optimize import linprog
import numpy as np
import time
import pandas as pd


def parse_instance_from_zip(file):
    lines = file.readlines()
    instance_name = ""
    d = 0
    r = 0
    SCj = []
    Dk = []
    Cjk = []

    i = 0  # Index pentru a itera prin lines
    while i < len(lines):
        line = lines[i].decode('utf-8').strip()
        if line.startswith("instance_name"):
            instance_name = line.split('=')[1].strip().strip('";')
        elif line.startswith("d ="):
            d = int(line.split('=')[1].strip().strip(';'))
        elif line.startswith("r ="):
            r = int(line.split('=')[1].strip().strip(';'))
        elif line.startswith("SCj ="):
            SCj = list(map(int, line.split('=')[1].strip().strip('[];').split()))
        elif line.startswith("Dk ="):
            Dk = list(map(int, line.split('=')[1].strip().strip('[];').split()))
        elif line.startswith("Cjk ="):
            # Extrage toate elementele Cjk, indiferent de linia în care apar
            costs = []
            while not line.endswith("];"):
                # Curățăm orice element care nu este un număr întreg
                costs.extend(
                    [x for x in line.replace('[', '').replace(']', '').replace(';', '').split() if x.isdigit()])
                i += 1
                line = lines[i].decode('utf-8').strip()
            # Curățăm și ultima linie de Cjk
            costs.extend([x for x in line.replace('[', '').replace(']', '').replace(';', '').split() if
                          x.isdigit()])

            # Creează matricea Cjk
            Cjk = [list(map(int, costs[i:i + r])) for i in range(0, len(costs), r)]

            print(f"Dimensiunea lui Cjk citită: {len(Cjk)} x {len(Cjk[0])}")  # Debug: Verifică dimensiunea

            # Verificare dimensiune Cjk
            if len(Cjk) != d or any(len(row) != r for row in Cjk):
                raise ValueError(
                    f"Dimensiunea lui Cjk este incorectă: {len(Cjk)} x {len(Cjk[0])} (trebuie să fie {d} x {r})")

        i += 1  # Crește indexul pentru a trece la următoarea linie

    return instance_name, d, r, SCj, Dk, Cjk


def solve_transportation_problem(SCj, Dk, Cjk):
    d = len(SCj)
    r = len(Dk)

    # Funcția obiectivului (costurile sunt transformate într-un vector)
    c = np.array(Cjk).flatten()

    # Verificare dacă c are dimensiunea corectă
    assert c.shape[0] == d * r, f"Dimensiunea lui c trebuie să fie {d * r}, dar este {c.shape[0]}"

    # Constrângeri pentru capacitățile depozitelor
    A_ub = np.zeros((d, d * r))
    for i in range(d):
        A_ub[i, i * r:(i + 1) * r] = 1

    # Verificarea dimensiunii lui A_ub
    assert A_ub.shape == (d, d * r), f"Dimensiunea lui A_ub trebuie să fie ({d}, {d * r}), dar este {A_ub.shape}"

    b_ub = SCj

    # Constrângeri pentru cererea magazinelor
    A_eq = np.zeros((r, d * r))
    for j in range(r):
        A_eq[j, j::r] = 1
    b_eq = Dk

    # Limite pentru variabilele decizionale (fiecare cantitate trebuie să fie >= 0)
    bounds = [(0, None) for _ in range(d * r)]

    # Soluționarea problemei
    start_time = time.time()
    result = linprog(c, A_ub=A_ub, b_ub=b_ub, A_eq=A_eq, b_eq=b_eq, bounds=bounds, method='highs')
    end_time = time.time()

    # Rezultate
    is_solved = result.success
    optimal_cost = result.fun if is_solved else None
    num_iterations = result.nit if is_solved else None
    run_time = end_time - start_time

    return optimal_cost, num_iterations, run_time, is_solved


def save_results(instances_results, output_file):
    df = pd.DataFrame(instances_results, columns=['Instance', 'Optimal Cost', 'Iterations', 'Run Time (s)', 'Solved'])
    df.to_excel(output_file, index=False)


def process_zip_file(zip_path, output_file):
    instances_results = []

    # Deschidem arhiva zip
    with zipfile.ZipFile(zip_path, 'r') as archive:
        for filename in archive.namelist():
            if filename.endswith('.dat'):
                with archive.open(filename) as file:
                    instance_name, d, r, SCj, Dk, Cjk = parse_instance_from_zip(file)

                    # Rezolvăm problema de transport
                    optimal_cost, num_iterations, run_time, is_solved = solve_transportation_problem(SCj, Dk, Cjk)

                    # Adăugăm rezultatul la listă
                    instances_results.append([
                        instance_name,
                        optimal_cost,
                        num_iterations,
                        run_time,
                        is_solved
                    ])

    save_results(instances_results, output_file)


zip_path = r'C:\Users\deniv\PycharmProjects\probleme_transport\simple_instances.zip'  # Path către arhiva zip
output_file = r'C:\Users\deniv\PycharmProjects\probleme_transport\simple_results.xlsx'  # Path pentru salvarea rezultatelor
process_zip_file(zip_path, output_file)