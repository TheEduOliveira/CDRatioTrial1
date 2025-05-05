
import streamlit as st
import pandas as pd
import numpy as np
import pickle  # For saving/loading simulation data (optional)
from Solver import solve_production_problem  # Import your solver function
from openpyxl import load_workbook  # For saving to Excel

def load_dataframes(excel_file):
    """Loads the dataframes from an Excel file."""

    try:
        demand_df = pd.read_excel(excel_file, sheet_name="Demand")
        line_capacity_df = pd.read_excel(excel_file, sheet_name="Line_Capacity")
        production_df = pd.read_excel(excel_file, sheet_name="Production")
        return demand_df, line_capacity_df, production_df
    except ValueError as e:
        st.error(f"Error loading Excel file or sheets: {e}")
        return None, None, None

def create_editable_matrix(data, index_label, column_labels, key):
    """
    Creates an editable matrix using Streamlit's data_editor.

    Args:
        data (pd.DataFrame): Data to display in the matrix.
        index_label (str): Label for the index column.
        column_labels (list): List of column labels.
        key (str): A unique key for the data_editor.

    Returns:
        pd.DataFrame: The edited data.
    """

    edited_data = st.data_editor(
        data,
        key=key,
        hide_index=False,
        num_rows="dynamic",
    )
    return edited_data

def load_simulation_data(simulation_id, excel_file="InputData.xlsx"):
    """
    Loads data for a specific simulation from the 'InputData.xlsx' Excel file.

    Args:
        simulation_id (int): The ID of the simulation to load.
        excel_file (str, optional): The name of the Excel file. Defaults to "InputData.xlsx".

    Returns:
        dict: A dictionary containing the simulation data.
              Returns None if no data is found for the simulation_id.
    """

    try:
        demand_df = pd.read_excel(excel_file, sheet_name="Demand")
        line_capacity_df = pd.read_excel(excel_file, sheet_name="Line_Capacity")
        production_df = pd.read_excel(excel_file, sheet_name="Production")
    except ValueError as e:
        st.error(f"Error loading Excel file or sheets: {e}")
        return None

    # Filter dataframes by RelatedSimulation
    demand_df = demand_df[demand_df["RelatedSimulation"] == simulation_id]
    line_capacity_df = line_capacity_df[line_capacity_df["RelatedSimulation"] == simulation_id]
    production_df = production_df[production_df["RelatedSimulation"] == simulation_id]

    if demand_df.empty or line_capacity_df.empty or production_df.empty:
        return None

    periods = sorted(demand_df["Period"].unique())
    lines = sorted(line_capacity_df["Line"].unique())
    products = sorted(demand_df["Product"].unique())

    simulation_data = {
        "periods": periods,
        "line_capacity": {
            line: {
                period: line_capacity_df[(line_capacity_df["Line"] == line) & (line_capacity_df["Period"] == period)]["Available Hours"].values[0]
                if not line_capacity_df[(line_capacity_df["Line"] == line) & (line_capacity_df["Period"] == period)].empty
                else 0
                for period in periods
            }
            for line in lines
        },
        "demand": {
            product: {
                period: demand_df[(demand_df["Product"] == product) & (demand_df["Period"] == period)]["Demand (Kg)"].values[0]
                if not demand_df[(demand_df["Product"] == product) & (demand_df["Period"] == period)].empty
                else 0
                for period in periods
            }
            for product in products
        },
        "production_ratio": {}
    }

    for period in periods:
        simulation_data["production_ratio"][period] = {
            product: {
                line: production_df[
                    (production_df["Product"] == product)
                    & (production_df["Line"] == line)
                    & (production_df["Period"] == period)
                ]["Production (Kg/h)"].values[0]
                if not production_df[
                    (production_df["Product"] == product)
                    & (production_df["Line"] == line)
                    & (production_df["Period"] == period)
                ].empty
                else 0
                for line in lines
            }
            for product in products
        }

    #st.write("simulation_data:")  # ADD THIS LINE
    #st.write(simulation_data)      # AND THIS LINE
    return simulation_data

def simulation_form(initial_data=None):
    """
    Creates the form for the user to input simulation parameters.

    Args:
        initial_data (dict, optional):  A dictionary containing data to pre-fill the form.
                                      Defaults to None (create a new simulation).

    Returns:
        dict: A dictionary containing the simulation parameters.
    """

    simulation_data = {}

    # Step 1: Period Selection
    st.header("1. Select Period Range")
    all_periods = [
        "2025'Q1", "2025'Q2", "2025'Q3", "2025'Q4",
        "2026'Q1", "2026'Q2", "2026'Q3", "2026'Q4",
        "2027'Q1", "2027'Q2", "2027'Q3", "2027'Q4",
        "2028'Q1", "2028'Q2", "2028'Q3", "2028'Q4",
        "2029'Q1", "2029'Q2", "2029'Q3", "2029'Q4",
        "2030'Q1", "2030'Q2", "2030'Q3", "2030'Q4", "2031'Q1"
    ]
    if initial_data and "periods" in initial_data and initial_data["periods"]:
        start_period = st.selectbox("Start Period", all_periods, index=all_periods.index(initial_data["periods"][0]))
        end_period = st.selectbox("End Period", all_periods, index=all_periods.index(initial_data["periods"][-1]))
    else:
        start_period = st.selectbox("Start Period", all_periods, index=0)
        end_period = st.selectbox("End Period", all_periods, index=len(all_periods) - 1)
    simulation_data["periods"] = [period for period in all_periods if start_period <= period <= end_period]

    # Step 2: Line Capacity
    st.header("2. Line Capacity")
    available_lines = set()
    if initial_data and "line_capacity" in initial_data:
        available_lines.update(initial_data["line_capacity"].keys())
        initial_line_capacity_data = pd.DataFrame(initial_data["line_capacity"]).T.fillna(0)
    else:
        initial_line_capacity_data = pd.DataFrame({period: [0] for period in simulation_data["periods"]}, index=["Line1"])

    edited_line_capacity = create_editable_matrix(
        initial_line_capacity_data.reset_index().rename(columns={"index": "Line"}),
        "Line",
        simulation_data["periods"],
        key="line_capacity"
    )
    simulation_data["line_capacity"] = edited_line_capacity.set_index("Line").to_dict('index')

    # Step 3: Product Demand
    st.header("3. Product Demand")
    available_products = set()
    if initial_data and "demand" in initial_data:
        available_products.update(initial_data["demand"].keys())
        initial_demand_data = pd.DataFrame(initial_data["demand"]).T.fillna(0)
    else:
        initial_demand_data = pd.DataFrame({period: [0] for period in simulation_data["periods"]}, index=["Product1"])
    edited_demand_data = create_editable_matrix(
        initial_demand_data.reset_index().rename(columns={"index": "Product"}),
        "Product",
        simulation_data["periods"],
        key="product_demand"
    )
    simulation_data["demand"] = edited_demand_data.set_index("Product").to_dict('index')

    # Step 4: Production Ratio
    st.header("4. Production Ratio (Kg/h)")
    use_same_ratio = st.checkbox("Use same Kg/h for all periods", value=True)  # Default to True
    production_ratio_data = {}
    
    if use_same_ratio:
        st.subheader("All Periods")
        available_lines = edited_line_capacity["Line"].unique().tolist()  # Corrected
        available_products = edited_demand_data["Product"].unique().tolist()  # Corrected

        if initial_data and "production_ratio" in initial_data:
            # Create a DataFrame directly from the production_ratio data
            data = {}
            for period, prod_data in initial_data["production_ratio"].items():
                for product, line_data in prod_data.items():
                    for line, value in line_data.items():
                        try:
                            value = float(value)
                        except (ValueError, TypeError):
                            value = 0
                        if product not in data:
                            data[product] = {}
                        data[product][line] = value

            #initial_production_data = pd.DataFrame(data).fillna(0).reindex(index=list(available_products), columns=list(available_lines)).fillna(0)
            initial_production_data = pd.DataFrame(data).T
        else:
            initial_production_data = pd.DataFrame({line: [0] for line in available_lines}, index=available_products)

        edited_production_data = create_editable_matrix(
            initial_production_data.reset_index().rename(columns={"index": "Product"}),
            "Product",
            list(available_lines),
            key="production_ratio_all"
        )

        # Store the production ratio for all periods (same value)
        for period in simulation_data["periods"]:
            production_ratio_data[period] = edited_production_data.set_index("Product").to_dict('index')

    else:
        st.subheader("Per Period")
        for period in simulation_data["periods"]:
            st.write(f"**{period}**")

            available_lines = edited_line_capacity["Line"].unique().tolist()  # Corrected
            available_products = edited_demand_data["Product"].unique().tolist()  # Corrected

            if initial_data and "production_ratio" in initial_data and period in initial_data["production_ratio"]:

                data = {}
                for product, line_data in initial_data["production_ratio"][period].items():
                    for line, value in line_data.items():
                        try:
                            value = float(value)
                        except (ValueError, TypeError):
                            value = 0
                        if product not in data:
                            data[product] = {}
                        data[product][line] = value
                initial_production_data = pd.DataFrame(data).T  # Simplificado!
            else:
                initial_production_data = pd.DataFrame({line: [0] for line in available_lines}, index=available_products)


            edited_production_data = create_editable_matrix(
                initial_production_data.reset_index().rename(columns={"index": "Product"}),
                "Product",
                list(available_lines),
                key=f"production_ratio_{period}"
            )
            production_ratio_data[period] = edited_production_data.set_index("Product").to_dict('index')

    simulation_data["production_ratio"] = production_ratio_data

    return simulation_data


def color_capacity_demand_ratio(val):
    """Colors the capacity/demand ratio based on the defined conditions."""
    if val < 1:
        color = 'red'
    elif val >= 1.1:
        color = 'blue'
    else:
        color = 'green'
    return f'background-color: {color}'

def SaveResults(excel_file, simulation_data, results_df_adjusted, capacity_demand_ratio_df):
    st.subheader("Save Simulation")

    # Use st.session_state para persistir os valores entre as execuções
    if "related_simulation" not in st.session_state:
        st.session_state.related_simulation = "Simulation X"  # Valor padrão
    if "category" not in st.session_state:
        st.session_state.category = "Chocolate"  # Valor padrão

    related_simulation = st.text_input(
        "Related Simulation", value=st.session_state.related_simulation,
        key="related_simulation_input"  # Adicione uma chave
    )
    category = st.text_input(
        "Category", value=st.session_state.category,
        key="category_input"  # Adicione uma chave
    )

    # Atualiza o session state
    st.session_state.related_simulation = related_simulation
    st.session_state.category = category
    
    if st.button("Save Simulation", key="save_button"): #Adicionei key para o botão
        # Salva os dados da simulação no arquivo Excel
        try:
            book = load_workbook(excel_file)
            with pd.ExcelWriter(excel_file, engine='openpyxl', mode='a', if_sheet_exists='overlay') as writer:
                writer.book = book
                writer.sheets = dict((ws.title, ws) for ws in book.worksheets)
                
                results_df_adjusted.to_excel(writer, sheet_name="Output Prod", index=False)
                capacity_demand_ratio_df.to_excel(writer, sheet_name="Output CD", index=False)
                
                # Use os valores do session state aqui também
                demand_df = pd.DataFrame(simulation_data["demand"])
                demand_df["RelatedSimulation"] = st.session_state.related_simulation
                demand_df["Category"] = st.session_state.category
                demand_df.to_excel(writer, sheet_name="Demand", index=True)

                line_capacity_df = pd.DataFrame(simulation_data["line_capacity"])
                line_capacity_df["RelatedSimulation"] = st.session_state.related_simulation
                line_capacity_df["Category"] = st.session_state.category
                line_capacity_df.to_excel(writer, sheet_name="Line_Capacity", index=True)
                
                production_ratio_df_dict = {}
                for period, data in simulation_data["production_ratio"].items():
                    production_ratio_df_dict[period] = pd.DataFrame(data).T
                production_ratio_df = pd.concat(production_ratio_df_dict, keys=production_ratio_df_dict.keys(),
                                                names=['Period', 'Product', 'Line'])
                production_ratio_df["RelatedSimulation"] = st.session_state.related_simulation
                production_ratio_df["Category"] = st.session_state.category
                production_ratio_df.to_excel(writer, sheet_name="Production_Ratio", index=True)
                
                writer.save()  # Salva as mudanças no arquivo
            st.success(f"Simulation {st.session_state.related_simulation} saved successfully!")
        except Exception as e:
            st.error(f"Error saving simulation data: {e}")

def main():
    st.title("Production Demand Allocation Simulator")
    excel_file = "InputData.xlsx"

    menu = ["Create New Simulation", "Consult Simulations"]
    choice = st.sidebar.selectbox("Select Option", menu)

    if choice == "Create New Simulation":
        st.header("Create a New Simulation")
        simulation_data = simulation_form()

        if st.button("Run Simulation"):
            demand_dict = {(period, "Chocolate", product): data for product, periods_data in simulation_data["demand"].items() for period, data in periods_data.items()}
            line_capacity_dict = {(period, "Chocolate", line): data for line, periods_data in simulation_data["line_capacity"].items() for period, data in periods_data.items()}
            production_rate_dict = {(period, "Chocolate", line, product): ratio for period, products_data in simulation_data["production_ratio"].items() for product, lines_data in products_data.items() for line, ratio in lines_data.items()}

            results_df_adjusted, capacity_demand_ratio_df = solve_production_problem(
                demand_dict, line_capacity_dict, production_rate_dict
            )

            st.subheader("Simulation Results")
            st.dataframe(results_df_adjusted)

            st.subheader("Line Capacity vs. Actual Allocated Demand by Period")
            # Pivot the DataFrame for display
            capacity_demand_pivot_df = capacity_demand_ratio_df.pivot_table(
                index="Line", columns="Period", values="Capacity/Demand", aggfunc="first"  # or 'mean' if needed
            )
            st.dataframe(capacity_demand_pivot_df.style.applymap(color_capacity_demand_ratio).format("{:.2f}"))

            SaveResults(excel_file, simulation_data, results_df_adjusted, capacity_demand_ratio_df)

    elif choice == "Consult Simulations":
        st.header("Consult Existing Simulations")
        demand_df = pd.read_excel(excel_file, sheet_name="Demand")
        available_simulation_ids = sorted(demand_df["RelatedSimulation"].unique())

        if not available_simulation_ids:
            st.warning("No simulations found.")
        else:
            selected_simulation_id = st.selectbox("Select Simulation to Edit", available_simulation_ids)
            if selected_simulation_id:
                simulation_data = load_simulation_data(selected_simulation_id, excel_file)
                if simulation_data:
                    st.subheader(f"Edit Simulation: {selected_simulation_id}")
                    simulation_data = simulation_form(initial_data=simulation_data)
                    if st.button("Run Simulation"):
                        demand_dict = {(period, "Chocolate", product): data for product, periods_data in simulation_data["demand"].items() for period, data in periods_data.items()}
                        line_capacity_dict = {(period, "Chocolate", line): data for line, periods_data in simulation_data["line_capacity"].items() for period, data in periods_data.items()}
                        production_rate_dict = {(period, "Chocolate", line, product): ratio for period, products_data in simulation_data["production_ratio"].items() for product, lines_data in products_data.items() for line, ratio in lines_data.items()}

                        results_df_adjusted, capacity_demand_ratio_df, results = solve_production_problem(
                            demand_dict, line_capacity_dict, production_rate_dict
                        )

                        st.subheader("Simulation Results")
                        st.dataframe(results_df_adjusted)

                        st.write(results)

                        st.subheader("Line Capacity vs. Actual Allocated Demand by Period")
                        # Pivot the DataFrame for display
                        capacity_demand_pivot_df = capacity_demand_ratio_df.pivot_table(
                            index="Line", columns="Period", values="Capacity/Demand", aggfunc="first"  # or 'mean' if needed
                        )
                        st.dataframe(capacity_demand_pivot_df.style.applymap(color_capacity_demand_ratio).format("{:.2f}"))

                        SaveResults(excel_file, simulation_data, results_df_adjusted, capacity_demand_ratio_df)

if __name__ == "__main__":
    main()