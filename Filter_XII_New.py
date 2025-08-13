# Filter_XII_New.py
import re
import pandas as pd

class DataProcessor_XII_New:
    def __init__(self, input_file):
        self.input_file = input_file
        self.read_data()
        self.process_data()

    def read_data(self):
        with open(self.input_file, encoding="utf-8", errors="ignore") as f:
            self.reader = f.read()

    def process_data(self):
        """
        Expected format per student:
         - Line1: Roll Gender Name <subcode1> <subcode2> ... <subcodeN> [maybe RESULT token]
         - Line2: mark1 grade1 mark2 grade2 ... markN gradeN [maybe other trailing tokens]
        We parse tokens, pair subject codes with mark-grade pairs.
        """
        # simple helpers (token patterns)
        roll_token_pattern = re.compile(r'^\d+$')
        code_pattern = re.compile(r'^\d{3}$')           # subject codes like 041, 301, etc.
        mark_token_pattern = re.compile(r'^(?:\d{1,3}|AB)$')   # 2/3 digit or AB
        grade_token_pattern = re.compile(r'^[A-Z]?\d?$|^E$')   # A1, B2, C, E etc.

        # Subject dictionary: map codes to names (expand as needed)
        self.subject_dict = {
            "041": "Maths", "042": "Physics", "043": "Chemistry", "044": "Biology",
            "301": "English", "048": "Physical Education", "030": "Economics",
            "054": "Business Studies", "055": "Accountancy", "302": "Hindi",
            "322": "Sanskrit", "065": "I.P.", "049": "Painting", "241" : "Applied Maths"
        }

        # prepare
        lines = [ln.rstrip() for ln in self.reader.splitlines() if ln.strip() != ""]
        students = []
        seen_subject_codes_ordered = []  # preserve first-seen order

        i = 0
        while i < len(lines) - 0:
            line1 = lines[i].strip()
            # safety: need at least one more line for marks; otherwise break
            if i + 1 >= len(lines):
                break
            line2 = lines[i + 1].strip()

            # Tokenize
            toks1 = line1.split()
            toks2 = line2.split()

            # Basic validation: first token should be roll (digits), second token gender (M/F or single letter)
            if len(toks1) < 3 or not roll_token_pattern.match(toks1[0]):
                # not a student header; skip one line and continue
                i += 1
                continue

            roll = toks1[0]
            gender = toks1[1]
            # find first token index in toks1 that looks like a subject code (3-digit)
            first_code_idx = None
            for idx in range(2, len(toks1)):
                if code_pattern.match(toks1[idx]):
                    first_code_idx = idx
                    break

            if first_code_idx is None:
                # can't find subject code in this header; skip (malformed)
                i += 1
                continue

            # name is tokens from 2 .. first_code_idx-1
            name = " ".join(toks1[2:first_code_idx]).strip()
            # subject_codes = toks1[first_code_idx:]
            # drop trailing tokens in subject_codes that are not 3-digit (just in case)
            subject_codes = [c for c in toks1[first_code_idx:] if code_pattern.match(c)]

            # record order of subject codes
            for c in subject_codes:
                if c not in seen_subject_codes_ordered:
                    seen_subject_codes_ordered.append(c)

            # parse marks/grades from line2: iterate tokens and pair mark+grade when possible
            mark_grade_pairs = []
            j = 0
            while j < len(toks2):
                t_mark = toks2[j]
                if mark_token_pattern.match(t_mark):
                    # look ahead for a grade token
                    grade = ""
                    if j + 1 < len(toks2) and grade_token_pattern.match(toks2[j + 1]):
                        grade = toks2[j + 1]
                        j += 2
                    else:
                        # no grade found; still accept mark, advance by 1
                        j += 1
                    mark_grade_pairs.append((t_mark, grade))
                else:
                    # token is not a mark, skip
                    j += 1

            # Now pair subject_codes with mark_grade_pairs
            data = {"Roll": roll, "Gender": gender, "Name": name}
            pairs_to_use = min(len(subject_codes), len(mark_grade_pairs))
            for k in range(pairs_to_use):
                code = subject_codes[k]
                mark_token, grade_token = mark_grade_pairs[k]
                subj_name = self.subject_dict.get(code, code)  # fallback to code if not found
                mark_val = None
                if mark_token != "AB":
                    try:
                        mark_val = int(mark_token)
                    except:
                        mark_val = None
                data[subj_name] = mark_val
                data["Grade_" + subj_name] = grade_token

            # If some subjects exist but marks missing, leave them absent (NaN)
            # append student
            students.append(data)

            i += 2  # next student block (since you said each student is 2 lines)

        # Build columns from seen_subject_codes_ordered mapped to names
        columns = ["Roll", "Gender", "Name"]
        for code in seen_subject_codes_ordered:
            subj_name = self.subject_dict.get(code, code)
            columns.append(subj_name)
            columns.append("Grade_" + subj_name)

        # create dataframe; fill missing cols with NaN (pd will do this if columns provided)
        self.df_final = pd.DataFrame(students, columns=columns)
        # keep blank cells as empty strings for UI consistency
        self.df_final = self.df_final.fillna("")

            # === AUTO CALCULATE PERCENTAGE ===
        numeric_marks = self.df_final[[c for c in self.df_final.columns if not c.startswith("Grade_") and c not in ["Roll", "Gender", "Name"]]].apply(pd.to_numeric, errors='coerce')

        # Option 1: Simple average of all subjects taken
        # self.df_final["Percentage"] = (numeric_marks.sum(axis=1) / (numeric_marks.count(axis=1) * 100) * 100).round(2)

        # Option 2: Top 5 percentage instead (uncomment if needed)
        self.df_final["Percentage"] = numeric_marks.apply(lambda row: row.nlargest(5).sum() / 500 * 100, axis=1).round(2)

    def export_to_excel(self, path):
        self.df_final.to_excel(path, index=False)
