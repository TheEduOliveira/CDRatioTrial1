import pandas as pd
from pyomo.environ import *
from pyomo.opt import SolverFactory
import numpy as np

def solve_production_problem(demand, line_capacity, production_rate):
    """
    Solves the production allocation problem.

    Args:
        demand (dict): Dictionary of demand data.
                      Keys: (Period, Category, Product)
                      Values: Demand (Kg)
        line_capacity (dict): Dictionary of line capacity.
                             Keys: (Period, Category, Line)
                             Values: Available Hours
        production_rate (dict): Dictionary of production rates.
                               Keys: (Period, Category, Line, Product)
                               Values: Production (Kg/h)

    Returns:
        pandas.DataFrame: DataFrame containing the production allocation results,
                        aggregated by Period, Category, Line, and Product.
    """

    Periods = sorted(list(set(p for p, _, _ in demand)))
    Categories = sorted(list(set(c for _, c, _ in demand)))
    Products = sorted(list(set(prod for _, _, prod in demand)))
    Lines = sorted(list(set(line for _, _, line in line_capacity)))

    # Adiciona linha fictícia (Fallback_Line)
    fictitious_line = "Fallback_Line"
    for (period, category, product) in demand:
        production_rate[(period, category, fictitious_line, product)] = 0.01
    for (period, category) in set((p, c) for p, c, _ in demand):
        line_capacity[(period, category, fictitious_line)] = float("inf")
    Lines.append(fictitious_line)

    # === Modelo Pyomo ===
    model = ConcreteModel()
    model.PERIODS = Set(initialize=Periods)
    model.CATEGORIES = Set(initialize=Categories)
    model.PRODUCTS = Set(initialize=Products)
    model.LINES = Set(initialize=Lines)

    model.X = Var(model.PERIODS, model.CATEGORIES, model.LINES, model.PRODUCTS, domain=NonNegativeReals)

    def demand_rule(m, p, c, prod):
        return sum(m.X[p, c, l, prod] * production_rate.get((p, c, l, prod), 0) for l in m.LINES) >= demand.get((p, c, prod), 0)
    model.DemandConstraint = Constraint(model.PERIODS, model.CATEGORIES, model.PRODUCTS, rule=demand_rule)

    def capacity_rule(m, p, c, l):
        return sum(m.X[p, c, l, prod] for prod in m.PRODUCTS if (p, c, l, prod) in production_rate) <= line_capacity.get((p, c, l), 0)
    model.CapacityConstraint = Constraint(model.PERIODS, model.CATEGORIES, model.LINES, rule=capacity_rule)

    model.Objective = Objective(
        expr=sum(model.X[p, c, l, prod] for p in model.PERIODS for c in model.CATEGORIES for l in model.LINES for prod in model.PRODUCTS),
        sense=minimize
    )

    # === Resolver ===
    solver = SolverFactory("appsi_highs")  # You might need to adjust the solver
    results = solver.solve(model, tee=False)  # Set tee=True for solver output

    # === Resultados ===
    data = []
    for p in Periods:
        for c in Categories:
            for l in Lines:
                for prod in Products:
                    val = model.X[p, c, l, prod].value
                    if val and val > 0:
                        kgs = val * production_rate.get((p, c, l, prod), 0)
                        data.append([p, c, l, prod, val, kgs])

    results_df = pd.DataFrame(data, columns=["Period", "Category", "Line", "Product", "Hours", "Kg_Produced"])

    # Redistribuir a fallback line
    fallback_name = "Fallback_Line"
    fallback_rows = results_df[results_df["Line"] == fallback_name]
    non_fallback_rows = results_df[results_df["Line"] != fallback_name]

    adjusted_rows = []

    for _, row in fallback_rows.iterrows():
        p, c, prod = row["Period"], row["Category"], row["Product"]
        total_kgs = row["Kg_Produced"]

        eligible_lines = [
            l for l in Lines if (p, c, l, prod) in production_rate and l != fallback_name
        ]
        if not eligible_lines:
            continue

        equal_kgs = total_kgs / len(eligible_lines)

        for l in eligible_lines:
            rate = production_rate.get((p, c, l, prod), None)
            if rate and rate > 0:
                adjusted_rows.append({
                    "Period": p,
                    "Category": c,
                    "Line": l,
                    "Product": prod,
                    "Kg_Produced": equal_kgs,
                    "Hours": equal_kgs / rate
                })

    results_df_adjusted = pd.concat([non_fallback_rows, pd.DataFrame(adjusted_rows)], ignore_index=True)

    # Agregar por Period, Category, Line, Product
    results_df_adjusted = results_df_adjusted.groupby(
        ["Period", "Category", "Line", "Product"], as_index=False
    ).sum()

    capacity_demand_data = []
    for (p, c, l), capacity in line_capacity.items():
        total_demand_hours = 0
        # Filter the results for the current period and line
        line_production = results_df_adjusted[
            (results_df_adjusted["Period"] == p) &
            (results_df_adjusted["Line"] == l)
        ]
        # Calculate the hours required for the produced Kg
        for _, row in line_production.iterrows():
            prod = row["Product"]
            produced_kg = row["Kg_Produced"]
            rate = production_rate.get((p, c, l, prod), 0)
            if rate > 0:
                total_demand_hours += produced_kg / rate
            else:
                total_demand_hours += 0  # Or handle this case as you see fit (e.g., raise an error, use a default rate)

        ratio = (capacity / (total_demand_hours + 1e-9)) if total_demand_hours > 0 else np.nan
        capacity_demand_data.append({"Period": p, "Category": c, "Line": l, "Capacity/Demand": ratio})

    capacity_demand_ratio_df = pd.DataFrame(capacity_demand_data)

    return results_df_adjusted, capacity_demand_ratio_df, results_df

if __name__ == '__main__':
    # === Example Usage (for testing) ===
    # Replace with your actual data loading
    file_path = "C://Users//KGQ2858//OneDrive - MDLZ//Downloads//solvercomstreamlit//InputData.xlsx"  # Or the path to your CSVs
    demand_df = pd.read_excel(file_path, sheet_name="Demand")
    line_capacity_df = pd.read_excel(file_path, sheet_name="Line_Capacity")
    production_df = pd.read_excel(file_path, sheet_name="Production")

    demand_df["Key"] = list(zip(demand_df.Period, demand_df.Category, demand_df.Product))
    demand = demand_df.set_index("Key")["Demand (Kg)"].to_dict()

    line_capacity_df["Key"] = list(zip(line_capacity_df.Period, line_capacity_df.Category, line_capacity_df.Line))
    line_capacity = line_capacity_df.set_index("Key")["Available Hours"].to_dict()

    production_df["Key"] = list(zip(production_df.Period, production_df.Category, production_df.Line, production_df.Product))
    production_rate = production_df.set_index("Key")["Production (Kg/h)"].to_dict()

    results = solve_production_problem(demand, line_capacity, production_rate)
    print(results)