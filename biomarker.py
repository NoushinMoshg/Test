import os
import pandas as pd
from scipy.stats import ttest_ind

def load_data(folder_path, patient_type):
    data_frames = []
    for file_name in os.listdir(folder_path):
        if file_name.endswith('.xls'):
            file_path = os.path.join(folder_path, file_name)
            df = pd.read_excel(file_path)
            df['Patient_Type'] = patient_type  # Add a column to indicate the patient type
            data_frames.append(df)
    return pd.concat(data_frames, ignore_index=True)

def normalize_rpm(data):
    # Implement normalization if needed
    return data

def convert_to_numeric(data):
    for column in data.columns:
        if column not in ['Patient_Type', 'contig_id']:
            data[column] = pd.to_numeric(data[column], errors='coerce')
    return data

def find_differential_genes(combined_data):
    breast_cancer_genes = combined_data[combined_data['Patient_Type'] == 'Breast_Cancer'].drop(columns=['Patient_Type'])
    normal_genes = combined_data[combined_data['Patient_Type'] == 'Normal'].drop(columns=['Patient_Type'])

    p_values = {}
    for gene in breast_cancer_genes.columns:
        # Ensure both series are numeric
        breast_cancer_values = pd.to_numeric(breast_cancer_genes[gene], errors='coerce').dropna()
        normal_values = pd.to_numeric(normal_genes[gene], errors='coerce').dropna()

        t_stat, p_val = ttest_ind(breast_cancer_values, normal_values)
        p_values[gene] = p_val
    
    significant_genes = [gene for gene, p_val in p_values.items() if p_val < 0.05]
    return significant_genes

def save_to_excel(significant_genes, output_file):
    significant_genes_df = pd.DataFrame(significant_genes, columns=['Significant_Genes'])
    significant_genes_df.to_excel(output_file, index=False)
    print(f"Significant genes saved to '{output_file}'.")

breast_cancer_folder = '/Users/noushinmoshgabadi/Desktop/Compare_expression_BC/Breast_Cancer'
normal_folder = '/Users/noushinmoshgabadi/Desktop/Compare_expression_BC/Breast_Normal'
output_file = '/Users/noushinmoshgabadi/Desktop/Compare_expression_BC/Significant_Genes.xlsx'

breast_cancer_data = load_data(breast_cancer_folder, 'Breast_Cancer')
normal_data = load_data(normal_folder, 'Normal')

combined_data = pd.concat([breast_cancer_data, normal_data])

combined_data = convert_to_numeric(combined_data)
combined_data = normalize_rpm(combined_data)

significant_genes = find_differential_genes(combined_data)
save_to_excel(significant_genes, output_file)
