import pandas as pd

# Load the Excel file
input_file = './SLP30_DiaryNotes.xlsx'
output_file = './SLP30_DiaryNotes_Updated.xlsx'
# input_file = './SLP30_CustomerTransactions.xlsx'
# output_file = './SLP30_CustomerTransactions_Updated.xlsx'

# Load the sheet (assuming it's the first one)
df = pd.read_excel(input_file)

# Ensure TLNO30 is treated as a string
df['TLNO30'] = df['TLNO30'].astype(str)

# Prepare rows to insert
rows_to_insert = []

for idx, row in df.iterrows():
    text = str(row['TLIN30'])
    # Replace double quotes with single quotes
    # text = text.replace('"', "'")
    # text = text.replace('\\', "")
    stripped_text = text.strip()  # Remove leading/trailing whitespace for length check
    if len(stripped_text) > 60:
        # Find the last space before position 60 in the stripped text
        split_index = stripped_text.rfind(' ', 0, 60)
        if split_index == -1:
            # No space found, split last 5 characters
            first_part = stripped_text[:-5].strip()
            second_part = stripped_text[-5:].strip()
        else:
            first_part = stripped_text[:split_index].strip()
            second_part = stripped_text[split_index:].strip()

        # Update the current row with the first part (strip before saving)
        df.at[idx, 'TLIN30'] = first_part

        # Create a new row for the second part
        new_row = {
            'CUSN30': row['CUSN30'],
            'CUNO': row['CUNO'],
            'SUBJ30': row['SUBJ30'],
            'TLNO30': f"{row['TLNO30']}.1",
            'TLIN30': second_part
        }
        rows_to_insert.append((idx + 0.5, new_row))
    else:
        # Always strip the string even if not split
        df.at[idx, 'TLIN30'] = stripped_text

# Insert the new rows after each original one
for insert_idx, new_row in sorted(rows_to_insert, key=lambda x: x[0], reverse=True):
    df.loc[insert_idx] = new_row

# Sort and reset index
df = df.sort_index().reset_index(drop=True)

# Save to a new Excel file
df.to_excel(output_file, index=False)

print(f"Updated file saved as {output_file}")