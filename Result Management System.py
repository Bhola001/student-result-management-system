import pandas as pd
import os

FILE_NAME = "Students.xlsx"
SYSTEM_COLUMNS = ["Name", "Class", "Age", "Percentage"]

# File load ya create
if os.path.exists(FILE_NAME):
    df = pd.read_excel(FILE_NAME, index_col=0)
else:
    df = pd.DataFrame(
        {
            "Name": ["Ram", "Shyam"],
            "Class": ["12th", "12th"]
        },
        index=[101, 102]
    )

    df.index.name = "Student ID"
    df.to_excel(FILE_NAME)

TOPPERS_FILE = "Class_Top_3_Toppers.xlsx"

def Export_Class_Top_3(df):

    if "Percentage" not in df.columns:
        print("Percentage column not found.")
        return

    toppers_data = []

    for cls in sorted(df["Class"].dropna().unique(), reverse=True):

        class_df = df[df["Class"] == cls]

        if class_df["Percentage"].isna().all():
            continue

        rank_df = class_df.sort_values(
            by="Percentage",
            ascending=False
        ).head(3)

        ranks = ["1st", "2nd", "3rd"]

        for i, (student_id, row) in enumerate(rank_df.iterrows()):

            toppers_data.append({
                "Class": cls,
                "Rank": ranks[i],
                "Student ID": student_id,
                "Name": row["Name"],
                "Percentage": row["Percentage"]
            })

    topper_df = pd.DataFrame(toppers_data)

    topper_df.to_excel(
        TOPPERS_FILE,
        index=False
    )
    

    print(f"{TOPPERS_FILE} created successfully.")

SUBJECT_TOPPERS_FILE = "Class_Wise_Subject_Toppers.xlsx"

def Export_Class_Wise_Subject_Toppers(df):

    toppers_data = []

    for cls in sorted(
        df["Class"].dropna().unique(),
        reverse=True
    ):

        class_df = df[df["Class"] == cls]

        for subject in class_df.columns:

            if subject in SYSTEM_COLUMNS:
                continue

            class_df[subject] = pd.to_numeric(
                class_df[subject],
                errors="coerce"
            )

            if class_df[subject].isna().all():
                continue

            topper_id = class_df[subject].idxmax()

            topper = class_df.loc[topper_id]

            toppers_data.append({
                "Class": cls,
                "Subject": subject,
                "Student ID": topper_id,
                "Name": topper["Name"],
                "Marks": topper[subject],
                "Percentage": topper["Percentage"]
            })

    if len(toppers_data) == 0:
        print("No subject topper data available.")
        return

    topper_df = pd.DataFrame(toppers_data)

    topper_df.to_excel(
        SUBJECT_TOPPERS_FILE,
        index=False
    )

    print(
        f"{SUBJECT_TOPPERS_FILE} created successfully."
    )

AVERAGE_REPORT_FILE = "Class_Average_Report.xlsx"
def Export_Class_Average_Report(df):

    if "Percentage" not in df.columns:
        print("Percentage column not found.")
        return

    report_data = []

    for cls in df["Class"].dropna().unique():

        class_df = df[df["Class"] == cls]

        if class_df["Percentage"].isna().all():
            continue

        avg_percentage = round(
            class_df["Percentage"].mean(),
            2
        )

        report_data.append(
            {
                "Class": cls,
                "Average Percentage": avg_percentage
            }
        )

    if len(report_data) == 0:
        print("No percentage data available.")
        return

    report_df = pd.DataFrame(report_data)

    report_df.to_excel(
        AVERAGE_REPORT_FILE,
        index=False
    )

    print(
        f"{AVERAGE_REPORT_FILE} created successfully."
    )    

REPORT_FILE = "School_Report.xlsx"

def Export_Pass_Students_Class_Wise(df):

    if "Percentage" not in df.columns:
        print("Percentage column not found.")
        return

    for cls in df["Class"].dropna().unique():

        class_df = df[df["Class"] == cls]

        pass_students = []

        for student_id, row in class_df.iterrows():

            fail_subjects = 0

            for col in df.columns:

                if col not in SYSTEM_COLUMNS:

                    try:
                        marks = float(row[col])

                        if marks < 33:
                            fail_subjects += 1

                    except:
                        pass

            # Pass Student
            if fail_subjects == 0:

                student_data = row.copy()
                student_data["Student ID"] = student_id

                pass_students.append(student_data)

        if len(pass_students) > 0:

            pass_df = pd.DataFrame(pass_students)

            file_name = f"{cls}_Pass_Students.xlsx"
            
            pass_df.to_excel(file_name, index=False)

            print(f"{file_name} created successfully.")

def Export_Compartment_Students_Class_Wise(df):

    if "Percentage" not in df.columns:
        print("Percentage column not found.")
        return

    for cls in df["Class"].dropna().unique():

        class_df = df[df["Class"] == cls]

        compartment_students = []

        for student_id, row in class_df.iterrows():

            fail_subjects = 0

            for col in df.columns:

                if col not in SYSTEM_COLUMNS:

                    try:
                        marks = float(row[col])

                        if marks < 33:
                            fail_subjects += 1

                    except:
                        pass

            # 1, 2, ya 3 subjects me fail = Compartment
            if 1 <= fail_subjects <= 3:
                student_data = row.copy()
                student_data["Student ID"] = student_id

                compartment_students.append(student_data)
        if compartment_students:

            compartment_df = pd.DataFrame(compartment_students)

            file_name = (
                f"{cls}_Compartment_Students.xlsx"
            )

            compartment_df.to_excel(file_name, index=False)

            print(
                f"{file_name} created successfully."
            )

def Export_Fail_Students_Class_Wise(df):

    if "Percentage" not in df.columns:
        print("Percentage column not found.")
        return

    for cls in df["Class"].dropna().unique():

        class_df = df[df["Class"] == cls]

        fail_students = []

        for student_id, row in class_df.iterrows():

            fail_subjects = 0

            for col in df.columns:

                if col not in SYSTEM_COLUMNS:

                    try:
                        marks = float(row[col])

                        if marks < 33:
                            fail_subjects += 1

                    except:
                        pass

            # 4 ya usse zyada subjects me fail = FAIL
            if fail_subjects > 3:
                student_data = row.copy()
                student_data["Student ID"] = student_id

                fail_students.append(student_data)

        if fail_students:

            fail_df = pd.DataFrame(fail_students)

            file_name = f"{cls}_Fail_Students.xlsx"

            fail_df.to_excel(file_name, index=False)

            print(f"{file_name} created successfully.")

SCHOOL_STATS_FILE = "School_Statistics.xlsx"

SCHOOL_STATS_FILE = "Class_Wise_Statistics.xlsx"

def Export_Class_Wise_Statistics(df):

    if "Percentage" not in df.columns:
        print("Percentage column not found.")
        return

    stats_data = []

    for cls in sorted(
        df["Class"].dropna().unique(),
        reverse=True
    ):

        class_df = df[df["Class"] == cls]

        total_students = len(class_df)

        average_percentage = round(
            class_df["Percentage"].mean(),
            2
        )

        pass_students = 0
        compartment_students = 0
        fail_students = 0

        for student_id, row in class_df.iterrows():

            fail_subjects = 0

            for col in class_df.columns:

                if col not in SYSTEM_COLUMNS:

                    try:
                        marks = float(row[col])

                        if marks < 33:
                            fail_subjects += 1

                    except:
                        pass

            if fail_subjects == 0:
                pass_students += 1

            elif fail_subjects <= 3:
                compartment_students += 1

            else:
                fail_students += 1

        school_topper = ""

        if not class_df["Percentage"].isna().all():

            topper = class_df.loc[
                class_df["Percentage"].idxmax()
            ]

            school_topper = (
                f"{topper['Name']} "
                f"({topper['Percentage']}%)"
            )

        stats_data.append({
            "Class": cls,
            "Total Students": total_students,
            "Average Percentage": average_percentage,
            "Topper": school_topper,
            "Pass Students": pass_students,
            "Compartment Students": compartment_students,
            "Fail Students": fail_students
        })

    stats_df = pd.DataFrame(stats_data)

    stats_df.to_excel(
        SCHOOL_STATS_FILE,
        index=False
    )

    print(
        f"{SCHOOL_STATS_FILE} created successfully."
    )

def Generate_Report(df):

    with pd.ExcelWriter(REPORT_FILE) as writer:

        # Main Data
        df.to_excel(writer, sheet_name="All_Students")

        # Class Wise Rank Lists
        if "Percentage" in df.columns:

            for cls in df["Class"].dropna().unique():

                class_df = df[df["Class"] == cls]

                rank_df = class_df.sort_values(
                    "Percentage",
                    ascending=False
                ).copy()

                rank_df.insert(
                    0,
                    "Rank",
                    range(1, len(rank_df) + 1)
                )

                rank_df.to_excel(
                    writer,
                    sheet_name=f"{cls}_Rank",
                    index=True
                )

        # Class Toppers
        toppers = []

        for cls in df["Class"].dropna().unique():

            class_df = df[df["Class"] == cls]

            if (
                "Percentage" in class_df.columns
                and not class_df["Percentage"].isna().all()
            ):
                topper = class_df.loc[
                    class_df["Percentage"].idxmax()
                ]

                toppers.append(topper)

        if toppers:
            pd.DataFrame(toppers).to_excel(
                writer,
                sheet_name="Class_Toppers"
            )

    print(f"{REPORT_FILE} generated successfully.")

def Save_Data(df):
    df.to_excel(FILE_NAME)
    Export_Class_Average_Report(df)
    Export_Pass_Students_Class_Wise(df)
    Export_Compartment_Students_Class_Wise(df)
    Export_Fail_Students_Class_Wise(df)
    Export_Class_Top_3(df)
    Export_Class_Wise_Subject_Toppers(df)
    Export_Class_Wise_Statistics(df)
    Generate_Report(df)

def generate_student_id(df): 
    ids = sorted(df.index) 
    new_id = 101 
    for i in ids: 
        if i == new_id: 
            new_id += 1 
        else: break 
    return new_id

def calculate_percentage(df, student_id):

    if "Percentage" not in df.columns:
        return

    total = 0
    count = 0

    for col in df.columns:
        if col not in SYSTEM_COLUMNS:
            try:
                total += float(df.loc[student_id, col])
                count += 1
            except:
                pass

    if count > 0:
        df.loc[student_id, "Percentage"] = round(total / count, 2)


def Add_New_Column(df):

    col_name = input("New column name: ")

    if col_name in df.columns:
        print("Column already exists.")
        return

    df[col_name] = 0.0

    if col_name.lower() == "percentage":

        for student_id in df.index:

            calculate_percentage(df, student_id)

    print(f"Column '{col_name}' added successfully!")

    Save_Data(df)

def Update_Column_Value(df):
    try:
        student_id = int(input("Student ID: "))
    except ValueError:
        print("Please enter a valid Student ID.")
        return

    if student_id not in df.index:
        print("Student ID not found!")
        return

    col_name = input("Column name: ")

    if col_name not in df.columns:
        print("Column not found!")
        return

    if col_name not in SYSTEM_COLUMNS:
        try:
            new_value = float(input("New value: "))
        except ValueError:
            print("Please enter a valid number.")
            return
    else:
        new_value = input("New value: ")

    df.loc[student_id, col_name] = new_value

    calculate_percentage(df, student_id)

    print("Data updated successfully!")

    Save_Data(df)

def Delete_column(df):
    column_name = input("Column Name to delete: ")

    if column_name in SYSTEM_COLUMNS:
        print("This column cannot be deleted.")
        return
        
    if column_name in df.columns:
        confirm = input("Enter (Yes/No) for Deleting:")
            
        if confirm.lower() == "yes" : 
                
            df.drop(column_name, axis=1, inplace=True)
            
            Save_Data(df)
            
            print("Column deleted.")
            
        elif confirm.lower() == "no":
            print("Column not deleted.")
            return

        else:
            print("Please a (Yes/No) only")

    else:
        print("Column not found.")

def Adding_New_Student_Detail(df):
    # Add New Student
    student_id = generate_student_id(df)
    print("Student ID:", student_id)

    name = input("Name: ")
    student_class = input("Class: ")
    
    try:
        age = int(input("Enter Age: "))

        if age <= 0:
            print("Invalid age")
            return

    except ValueError:
        print("Please enter only numbers.")
        return


    df.loc[student_id, "Name"] = name
    df.loc[student_id, "Class"] = student_class
    df.loc[student_id, "Age"] = age

    # Baaki columns ke liye value lo
    for col in df.columns:
        if col not in SYSTEM_COLUMNS:
            try:
                value = float(input(f"Enter {col}: "))
            except ValueError:
                print("Invalid marks.")
                return

            df.loc[student_id, col] = value

    calculate_percentage(df, student_id)

    print("Student added successfully!")

    Save_Data(df)

def Update_Single_Row(df):
    # Update Full Row
    try:
        student_id = int(input("Student ID: "))
    except ValueError:
        print("Please enter a valid Student ID.")
        return

    if student_id in df.index:

        for col in df.columns:
            if col != "Percentage":

                if col == "Age":
                    try:
                        value = int(input("Enter Age: "))
                    except ValueError:
                        print("Invalid Age")
                        return
                
                elif col not in SYSTEM_COLUMNS:
                    try:
                        value = float(input(f"Enter {col}: "))
                    except ValueError:
                        print("Please enter a valid number.")
                        return

                else:
                    value = input(f"Enter {col}: ")

                df.loc[student_id, col] = value

        calculate_percentage(df, student_id)    

        print("Row updated successfully!")

    else:
        print("Student ID not found!")

    Save_Data(df)

def Delete_Student_Detail(df):
    try:
        s_D_id = int(input("Enter Student ID: "))
    except ValueError:
        print("Please enter a valid Student ID.")
        return

    if s_D_id in df.index :
        confirm = input("Enter (Yes/No) for Deleting:")
            
        if confirm.lower() == "yes" : 
            df.drop(s_D_id, inplace=True)
            print("Students delete successfully.")

            print("Next Available Student ID:", generate_student_id(df))

        elif confirm.lower() == "no":
            print("Students not deleted.")
            return

        else:
            print("Please enter Yes or No only.")

    else:
        print("Student Id not found.")
    
    Save_Data(df)

def Check_School_Topper(df):
    
    if "Percentage" not in df.columns:
        print("Percentage column not found.")
        return

    if df["Percentage"].isna().all():
        print("No percentage data available.")
        return

    # Maximum percentage wala student
    top_student = df.loc[df["Percentage"].idxmax()]
        
    print("Student ID:", top_student.name)
    print("Name of Student:", top_student["Name"])
    print("Percentage of Student:", top_student["Percentage"])

    print("\nMarks:")
    for col in df.columns:
        if col not in SYSTEM_COLUMNS:
            print(f"{col}: {top_student[col]}")

def pass_fail_Compartment(df):
    try:
        student_id = int(input("Enter Student ID: "))
    except ValueError:
        print("Please enter only numbers.")
        return

    if student_id not in df.index:
        print("Student ID not found.")
        return 

    fail_subjects = 0

    for col in df.columns:
        if col not in SYSTEM_COLUMNS:

            try:
                marks = float(df.loc[student_id, col])

                if marks < 33:
                    fail_subjects += 1

            except (ValueError, TypeError):
                pass
    student_name = df.loc[student_id, "Name"]

    if fail_subjects == 0:
        print(
            f"Student {student_name} PASSED with "
            f"{df.loc[student_id, 'Percentage']}%"
        )

    elif fail_subjects <= 3:
        print(
            f"Student {student_name} has COMPARTMENT in "
            f"{fail_subjects} subject(s)."
        )

    else:
        print(
            f"Student {student_name} FAILED in "
            f"{fail_subjects} subject(s)."
        )

def Search_Student(df):
    try:
        student_id = int(input("Student ID: "))
    except ValueError:
        print("Please enter a valid Student ID.")
        return

    if student_id in df.index:
        print("\nStudent Details:")
        print(df.loc[student_id])
    else:
        print("Student ID not found.")

def Subject_Wise_Topper(df):

    subject = input("Enter Subject Name: ")

    if subject not in df.columns:
        print("Subject not found!")
        return

    if subject in SYSTEM_COLUMNS:
        print("Please enter a valid subject name.")
        return
    
    df[subject] = pd.to_numeric(df[subject], errors="coerce")

    if df[subject].isna().all():
        print("No marks available.")
        return

    topper = df.loc[df[subject].idxmax()]
    
    # df[subject].idxmax() find Student_ID 
    # df.loc[Student_ID] 
    # "It will find the details of the student 
    # whose Student ID is generated."
    
    
    print("\n=== Subject Wise Topper ===")
    print("Subject:", subject)
    print("Student ID:", topper.name)
    print("Name:", topper["Name"])
    print("Marks:", topper[subject])
    print("Percentage:", topper["Percentage"])

def Rank_List(df):
    
    if "Percentage" not in df.columns:
        print("Percentage column not found.")
        return
    
    if df["Percentage"].isna().all():
        print("No percentage data available.")
        return

    rank_df = df.sort_values("Percentage", ascending=False)

    print("\n=== School Rank List ===")

    rank = 1

    for student_id, row in rank_df.iterrows():

        print(
            f"Rank {rank} | "
            f"ID: {student_id} | "
            f"Name: {row['Name']} | "
            f"Percentage: {row['Percentage']}"
        )

        rank += 1

def Student_Report_Card(df):

    try:
        student_id = int(input("Enter Student ID: "))
    except ValueError:
        print("Invalid Student ID.")
        return

    if student_id not in df.index:
        print("Student not found.")
        return

    student = df.loc[student_id]

    print("\n===== REPORT CARD =====")
    print("Student ID:", student_id)
    print("Name:", student["Name"])
    print("Class:", student["Class"])
    print("Age:", student["Age"])

    print("\nMarks:")

    for col in df.columns:
        if col not in SYSTEM_COLUMNS:
            print(f"{col}: {student[col]}")

    print("\nPercentage:", student["Percentage"])

def Class_Topper(df):

    student_class = input("Enter Class: ")

    class_df = df[df["Class"] == student_class]

    if len(class_df) == 0:
        print("No students found.")
        return
    
    if "Percentage" not in class_df.columns:
        print("Percentage column not found.")
        return

    if class_df["Percentage"].isna().all():
        print("No percentage data available.")
        return

    topper = class_df.loc[class_df["Percentage"].idxmax()]

    print("\n=== Class Wise Topper ===")
    print("Class:", student_class)
    print("Student ID:", topper.name)
    print("Name:", topper["Name"])
    print("Percentage:", topper["Percentage"])


def Class_Wise_Rank(df):

    student_class = input("Enter Class: ")

    class_df = df[df["Class"] == student_class]

    if len(class_df) == 0:
        print("No students found.")
        return

    if "Percentage" not in class_df.columns:
        print("Percentage column not found.")
    
        return

    if class_df["Percentage"].isna().all():
        print("No percentage data available.")
        return

    rank_df = class_df.sort_values("Percentage", ascending=False)

    print(f"\n=== Rank List of Class {student_class} ===")

    rank = 1

    for student_id, row in rank_df.iterrows():

        print(
            f"Rank {rank} | "
            f"ID: {student_id} | "
            f"Name: {row['Name']} | "
            f"Percentage: {row['Percentage']}"
        )

        rank += 1



while True:
    
    print("\n=== Result Management System ===")
    print("1. Show Data")
    print("2. Add New Column")
    print("3. Update Column Value")
    print("4. Delete column")
    print("5. Adding New Student Detail")
    print("6. Update a Single Row")
    print("7. Delete Student Detail")
    print("8. Total Student")
    print("9. Export to CSV")
    print("10.Check_School_Topper")
    print("11.Check pass / fail / Compartment Students")
    print("12.Search Student")
    print("13.Subject Wise Topper")
    print("14.Rank List")
    print("15.Student Report Card")
    print("16.Class Topper")
    print("17.Class Wise Rank")
    print("18.Save & Exit")

    choice = input("Enter choice: ")

    if choice == "1":
        print(df)

    elif choice == "2":
        Add_New_Column(df)

    elif choice == "3":
        Update_Column_Value(df)

    elif choice == "4":
        Delete_column(df)

    elif choice == "5":
        Adding_New_Student_Detail(df)

    elif choice == "6":
        Update_Single_Row(df)

    elif choice == "7":
        Delete_Student_Detail(df)
    
    elif choice == "8":
        print("Total Students:", len(df))

    elif choice == "9":
        df.to_csv("Students.csv", index=True)
        print("CSV file created.")

    elif choice == "10":
        Check_School_Topper(df)

    elif choice == "11":
        pass_fail_Compartment(df)

    elif choice == "12":
        Search_Student(df)

    elif choice == "13":
        Subject_Wise_Topper(df)

    elif choice == "14":
        Rank_List(df)

    elif choice == "15":
        Student_Report_Card(df)

    elif choice == "16":
        Class_Topper(df)

    elif choice == "17":
        Class_Wise_Rank(df)

    elif choice == "18":
        Save_Data(df)
        print("Data saved successfully!")
        break

    else:
        print("Invalid choice!")