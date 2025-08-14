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
    - datasets: metadata of each stored dataset (with column order json)
    - results: detailed marks, grades, and percentage per student per subject
    """
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()

    # Table for dataset metadata
    cursor.execute("""
    CREATE TABLE IF NOT EXISTS datasets (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        name TEXT NOT NULL,
        col_order TEXT, -- stores JSON column order
        created_at TEXT NOT NULL
    )
    """)

    # Table for all student results
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
        df (pd.DataFrame): Wide-format DataFrame containing Roll, Gender, Name,
                           subject marks, Grade_subject columns, and 'Percentage'.

    Notes:
        - Flattens DataFrame to subject-wise rows for storage.
        - Stores Percentage for each student in each subject row.
        - col_order is saved in datasets table for reconstruction on load.
    """
    if df is None or df.empty:
        raise ValueError("DataFrame is empty or None")

    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()

    created_at = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    col_order_json = json.dumps(df.columns.tolist())  # Store exact column order

    cursor.execute(
        "INSERT INTO datasets (name, col_order, created_at) VALUES (?, ?, ?)",
        (dataset_name, col_order_json, created_at)
    )
    dataset_id = cursor.lastrowid

    # Flatten DataFrame into subject-wise rows
    records = []
    for _, row in df.iterrows():
        roll = row["Roll"]
        gender = row["Gender"]
        name = row["Name"]
        overall_percentage = str(row["Percentage"]) if "Percentage" in df.columns else ""

        for col in df.columns:
            # Skip meta and percentage columns
            if col not in ["Roll", "Gender", "Name", "Percentage"] and not col.startswith("Grade_"):
                mark = row[col]
                grade_col = f"Grade_{col}"
                grade = row[grade_col] if grade_col in df.columns else ""
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

    cursor.execute("SELECT col_order FROM datasets WHERE id=?", (dataset_id,))
    col_order_result = cursor.fetchone()
    stored_col_order = json.loads(col_order_result[0]) if col_order_result and col_order_result[0] else None

    cursor.execute("""
    SELECT Roll, Gender, Name, subject, mark, grade, percentage
    FROM results
    WHERE dataset_id = ?
    """, (dataset_id,))
    rows = cursor.fetchall()
    conn.close()

    if not rows:
        return pd.DataFrame()

    # Convert to DataFrame
    # Build df_long with consistent column names
    df_long = pd.DataFrame(
        rows,
        columns=["Roll", "Gender", "Name", "Subject", "Mark", "Grade", "Percentage"]
    )

    # Ensure correct dtypes (strings are fine; keep as-is)
    # Create percentage map keyed by (Roll, Gender, Name) to be robust
    key_cols = ["Roll", "Gender", "Name"]
    df_long[key_cols] = df_long[key_cols].astype(str)
    percentage_map = df_long.drop_duplicates(key_cols + ["Percentage"]).set_index(key_cols)["Percentage"].to_dict()

    # Pivot marks & grades back into wide format
    df_pivot = df_long.pivot_table(
        index=key_cols,
        columns="Subject",
        values=["Mark", "Grade"],
        aggfunc="first"
    )

    # Bring index back to columns and flatten headers
    df_pivot = df_pivot.reset_index()

    # Flatten multiindex columns into single level: Mark->subject, Grade->Grade_subject
    new_cols = []
    for top, bottom in df_pivot.columns:
        if top in key_cols and bottom == "":
            new_cols.append(top)
        elif top == "Mark":
            new_cols.append(bottom)
        elif top == "Grade":
            new_cols.append(f"Grade_{bottom}")
        else:
            # Fallback for plain columns
            new_cols.append(top if bottom == "" else f"{top}_{bottom}")
    df_pivot.columns = new_cols

    # Add Percentage by matching on the full identity (Roll, Gender, Name)
    df_pivot["Percentage"] = df_pivot.apply(
        lambda r: percentage_map.get((str(r["Roll"]), str(r["Gender"]), str(r["Name"])), ""),
        axis=1
    )

    # Restore stored column order if available
    if stored_col_order:
        current = [c for c in stored_col_order if c in df_pivot.columns]
        # Ensure Percentage last
        if "Percentage" in current:
            current = [c for c in current if c != "Percentage"] + ["Percentage"]
        df_pivot = df_pivot.reindex(columns=current)

    return df_pivot.fillna("")
