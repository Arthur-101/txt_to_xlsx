# db_manager.py
"""
Database Manager:
Handles saving, loading, and listing student result datasets
using SQLite. Stores not only marks and grades but also overall
percentage for each student.
"""

import sqlite3
import pandas as pd
from datetime import datetime
import json

# Database filename (SQLite DB in current directory)
DB_FILE = "results.db"

def init_db():
    """
    Initialize the database by creating required tables if they don't exist.
    - datasets: metadata of each stored dataset
    - results: detailed marks, grades, and percentage per student per subject
    """
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()

    # Table for dataset metadata
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS datasets (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        name TEXT NOT NULL,
        created_at TEXT NOT NULL
    )
    """)

    # Table for all student results (linked to datasets table)
    # Added 'percentage' column for each row (replicate for subjects in case subject-wise percentage needed)
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS results (
        dataset_id INTEGER,
        Roll TEXT,
        Gender TEXT,
        Name TEXT,
        subject TEXT,
        mark TEXT,
        grade TEXT,
        percentage TEXT,
        FOREIGN KEY(dataset_id) REFERENCES datasets(id)
    )
    """)

    conn.commit()
    conn.close()

def save_dataframe(dataset_name, df):
    """
    Save a processed results DataFrame into the database.

    Parameters:
    dataset_name (str): The name to store for this dataset.
    df (pd.DataFrame): The wide-format DataFrame containing Roll, Gender, Name,
                       marks per subject, Grade_subject columns, and possibly 'Percentage'.

    Notes:
    - This function flattens the DataFrame to a subject-wise row format.
    - Percentage is pulled from `Percentage` column if present in df.
    """
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()

    # Insert into datasets table
    created_at = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    # cursor.execute("INSERT INTO datasets (name, created_at) VALUES (?, ?)", (dataset_name, created_at))
    # dataset_id = cursor.lastrowid
    col_order_json = json.dumps(df.columns.tolist())  # Store exact column order
    cursor.execute(
        "INSERT INTO datasets (name, created_at) VALUES (?, ?)",
        (dataset_name, created_at)
    )
    dataset_id = cursor.lastrowid


    # Flatten DataFrame into rows for storage
    records = []
    for _, row in df.iterrows():
        roll = row["Roll"]
        gender = row["Gender"]
        name = row["Name"]

        # Retrieve overall percentage if available
        overall_percentage = str(row["Percentage"]) if "Percentage" in df.columns else ""

        # For each subject in DataFrame, insert marks & grade along with percentage
        for col in df.columns:
            if col not in ["Roll", "Gender", "Name", "Percentage"] and not col.startswith("Grade_"):
                mark = row[col]
                grade = row.get(f"Grade_{col}", "")
                records.append((dataset_id, roll, gender, name, col, str(mark), str(grade), overall_percentage))

    cursor.executemany("""
    INSERT INTO results (dataset_id, Roll, Gender, Name, subject, mark, grade, percentage)
    VALUES (?, ?, ?, ?, ?, ?, ?, ?)
    """, records)

    conn.commit()
    conn.close()

def list_datasets():
    """
    Returns:
        List of (id, name, created_at) tuples for all saved datasets ordered by recent first.
    """
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    cursor.execute("SELECT id, name, created_at FROM datasets ORDER BY created_at DESC")
    rows = cursor.fetchall()
    conn.close()
    return rows

def load_dataframe(dataset_id):
    """
    Load a previously saved dataset from the DB into wide DataFrame format.

    Parameters:
        dataset_id (int): ID from datasets table.

    Returns:
        pd.DataFrame in wide format:
            Roll | Gender | Name | Subject1 | Grade_Subject1 | ... | Percentage
    """
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()

    cursor.execute("SELECT name FROM datasets WHERE id=?", (dataset_id,))
    stored_cols_result = cursor.fetchone()
    stored_col_order = json.loads(stored_cols_result[0]) if stored_cols_result and stored_cols_result[0] else None

    cursor.execute("""
    SELECT Roll, Gender, Name, subject, mark, grade, percentage
    FROM results
    WHERE dataset_id = ?
    """, (dataset_id,))
    rows = cursor.fetchall()
    conn.close()

    if not rows:
        return pd.DataFrame()

    # Convert to DataFrame for pivot
    # df = pd.DataFrame(rows, columns=["Roll", "Gender", "Name", "Subject", "Mark", "Grade", "Percentage"])

    # Extract percentages separately (avoid duplication — percentage is same for each subject for a student)
    # percentage_map = df.groupby(["Roll"])["Percentage"].first().to_dict()
    df_long = pd.DataFrame(rows, columns=["Roll", "Gender", "Name", "Subject", "Mark", "Grade", "Percentage"])
    percentage_map = df_long.groupby("Roll")["Percentage"].first().to_dict()

    # Pivot marks & grades back into wide format
    df_pivot = df_long.pivot_table(
        index=["Roll", "Gender", "Name"],
        columns="Subject",
        values=["Mark", "Grade"],
        aggfunc="first"
    ).reset_index()

    # Flatten multi-index columns into simple column names
    df_pivot.columns = [
        col[1] if col[0] == "Mark" else
        f"Grade_{col[1]}" if col[0] == "Grade" else col
        for col in df_pivot.columns
    ]

    # Add Percentage column at the end
    df_pivot["Percentage"] = df_pivot["Roll"].map(percentage_map)

    if stored_col_order:
        current_cols = df_pivot.columns.tolist()
        ordered_cols = [c for c in stored_col_order if c in current_cols]
        # Ensure Percentage is last
        if "Percentage" in ordered_cols:
            ordered_cols.remove("Percentage")
        ordered_cols.append("Percentage")
        df_pivot = df_pivot[ordered_cols]


    return df_pivot.fillna("")
