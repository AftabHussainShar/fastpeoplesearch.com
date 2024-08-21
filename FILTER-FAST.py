import pandas as pd
input_file = 'output_results.csv' 
df = pd.read_csv(input_file)
df['Full Name'] = df['First Name'] + ' ' + df['Last Name']
results = []
for index, row in df.iterrows():
    full_name = row['Full Name']
    first_name = row['First Name']
    matched_number = None
    for i in range(1, 8): 
        result_name_col = f'Result {i} Name'
        result_number_col = f'Result {i} Number'
        
        if pd.notna(row[result_name_col]) and (full_name in row[result_name_col] or first_name in row[result_name_col]):
            matched_number = row[result_number_col]
            break
    
    results.append({
        'First Name': row['First Name'],
        'Last Name': row['Last Name'],
        'Matched Number': matched_number
    })

output_df = pd.DataFrame(results)
output_file = 'final_output.csv'  
output_df.to_csv(output_file, index=False)
print(f"Results saved to {output_file}")
