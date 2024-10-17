import pandas as pd
import plotly.graph_objects as go
import networkx as nx

''' This file specifically plots the C. Elegans male chemical connectome.

    Adjacency matrices come from: https://wormwiring.org/pages/adjacency.html

    :)
'''

# Load the Excel file with adjacency matrices
file_path = 'SI 5 Connectome adjacency matrices, corrected July 2020.xlsx'
excel_data = pd.ExcelFile(file_path)

# Loading "male chemical" sheet
male_chemical_df = pd.read_excel(excel_data, sheet_name='male chemical')

# Function to clean the male chemical adjacency matrix (preparing the data)
def clean_male_adjacency_matrix(df):
    # Drop the first row and column which are not part of the actual data
    df_cleaned = df.iloc[1:, 3:]
    
    # Set the first column as index (neuron names)
    df_cleaned.index = df.iloc[1:, 2]
    
    # Set the first row as the column headers (neuron names)
    df_cleaned.columns = df.iloc[1, 3:]
    
    # Convert all values to numeric, forcing non-numeric entries to NaN
    df_cleaned = df_cleaned.apply(pd.to_numeric, errors='coerce')
    
    # Fill NaN values with 0 (assuming no connection where NaN is present)
    df_cleaned = df_cleaned.fillna(0)
    
    return df_cleaned

# "Cleaning" the male chemical adjacency matrix
male_chemical_cleaned = clean_male_adjacency_matrix(male_chemical_df)
# Interactive 3D visualization function using plotly
def plot_connectome_plotly(matrix):
    # Creating a digraph from  adjacency matrix
    G = nx.from_pandas_adjacency(matrix, create_using=nx.DiGraph)

    # Generating 3D layout for the nodes using spring layout
    pos = nx.spring_layout(G, dim=3, seed=42)

    # Extract node positions
    x_nodes = [pos[node][0] for node in G.nodes()]
    y_nodes = [pos[node][1] for node in G.nodes()]
    z_nodes = [pos[node][2] for node in G.nodes()]

    # Creating edges (connections)
    edge_trace = []
    for edge in G.edges():
        x0, y0, z0 = pos[edge[0]]
        x1, y1, z1 = pos[edge[1]]
        edge_trace.append(go.Scatter3d(
            x=[x0, x1, None],
            y=[y0, y1, None],
            z=[z0, z1, None],
            mode='lines',
            line=dict(color='blue', width=2),
            opacity=0.5
        ))

    # Creating nodes (neurons)
    node_trace = go.Scatter3d(
        x=x_nodes, y=y_nodes, z=z_nodes,
        mode='markers',
        marker=dict(
            size=5,
            color='red',
            opacity=0.8
        ),
        text=list(G.nodes()),
        hoverinfo='text'
    )

    # Creating 3D figure
    fig = go.Figure(data=edge_trace + [node_trace])
    fig.update_layout(
        title='3D Visualization of C. elegans Male Connectome',
        scene=dict(
            xaxis=dict(title='X Axis'),
            yaxis=dict(title='Y Axis'),
            zaxis=dict(title='Z Axis')
        )
    )
    fig.show()

# Running the visualization with the cleaned matrix
plot_connectome_plotly(male_chemical_cleaned)