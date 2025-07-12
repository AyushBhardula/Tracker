import pandas as pd
import random
import os

# === SETTINGS ===
excel_path = r"G:\My Drive\SharedData\users.xlsx"
sheet_name = "Employee_data"
output_path = r"G:\My Drive\IdPass_pass\ID_pass.xlsx"

# === Load Latest Employee Data ===
df_new = pd.read_excel(excel_path, sheet_name=sheet_name, engine='openpyxl')
df_new.columns = df_new.columns.str.strip()  # Clean headers

# === Create Unique Key to Identify Users (using email if available else first+last) ===
if 'Email Address' in df_new.columns:
    df_new["UniqueKey"] = df_new["Email Address"].str.strip().str.lower()
else:
    df_new["UniqueKey"] = (df_new["First Name"].str.strip().str.lower() + "_" +
                           df_new["Last Name"].str.strip().str.lower())

print("New UniqueKeys:")
print(df_new["UniqueKey"].tolist())
print("Any duplicates in new data?", df_new["UniqueKey"].duplicated().any())

# === Load Existing Data or Create Empty ===
if os.path.exists(output_path):
    df_existing = pd.read_excel(output_path, engine='openpyxl')
    df_existing.columns = df_existing.columns.str.strip()

    # Ensure UniqueKey exists and create if missing
    if "UniqueKey" not in df_existing.columns:
        if 'Email Address' in df_existing.columns:
            df_existing["UniqueKey"] = df_existing["Email Address"].str.strip().str.lower()
        else:
            df_existing["UniqueKey"] = (df_existing["First Name"].str.strip().str.lower() + "_" +
                                        df_existing["Last Name"].str.strip().str.lower())

    print("Existing UniqueKeys:")
    print(df_existing["UniqueKey"].tolist())
    print("Any duplicates in existing data?", df_existing["UniqueKey"].duplicated().any())
else:
    # Make empty DataFrame with necessary columns and proper index
    df_existing = pd.DataFrame(columns=list(df_new.columns) + ["Username", "Password", "UniqueKey"])
    df_existing.index = pd.RangeIndex(start=0, stop=0)  # Fix empty DataFrame index
    print("No existing data file found, starting fresh.")

# === Check for duplicate UniqueKeys in both DataFrames ===
if df_new["UniqueKey"].duplicated().any():
    raise ValueError("Duplicate UniqueKeys found in source employee data!")

if df_existing["UniqueKey"].duplicated().any():
    raise ValueError("Duplicate UniqueKeys found in existing user/password data!")

# === Find new users by UniqueKey ===
new_users = df_new[~df_new["UniqueKey"].isin(df_existing["UniqueKey"])]

print(f"New users found: {len(new_users)}")
print("New users UniqueKeys:", new_users["UniqueKey"].tolist())


# === Password Generator ===
def generate_password(first_name):
    special_chars = "!@#$"
    return "Smb" + first_name[:2].capitalize() + str(random.randint(0, 9)) + random.choice(special_chars)


# === Username Generator ===
def generate_username(first, last, index):
    return (first[:2] + last[:2]).lower() + f"{index:02}"


# === Generate for New Users ===
start_index = len(df_existing)
new_users = new_users.copy()

new_users["Username"] = [generate_username(row["First Name"], row["Last Name"], i + start_index)
                         for i, row in new_users.iterrows()]
new_users["Password"] = new_users["First Name"].apply(generate_password)

# === Reset index before concatenation ===
df_existing.reset_index(drop=True, inplace=True)
new_users.reset_index(drop=True, inplace=True)

print("Existing index:", df_existing.index.tolist())
print("New index:", new_users.index.tolist())

# === Concatenate safely ===
if df_existing.empty:
    df_final = new_users
else:
    df_final = pd.concat([df_existing, new_users], ignore_index=True)

# === Drop UniqueKey before saving ===
df_final.drop(columns=["UniqueKey"], inplace=True)

# === Save ===
df_final.to_excel(output_path, index=False)

print(f"✅ {len(new_users)} new user(s) added.")
print(f"✅ Usernames and passwords saved to:\n{output_path}")




