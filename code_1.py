import pandas as pd
import numpy as np
import re
from collections import Counter
import matplotlib.pyplot as plt
import seaborn as sns

# Read the spreadsheet into a pandas DataFrame
df = pd.read_excel("Coursera MASKED data.xlsx")

# Remove rows with missing values
df.replace("#", np.nan, inplace=True)

# Calculate the count of occurrences for each division, group, and department
division_counts = df["Division"].value_counts()
group_counts = df["Group"].value_counts()
department_counts = df["Department"].value_counts()

# Find the most popular division, group, and department
most_popular_division = division_counts.idxmax()
most_popular_group = group_counts.idxmax()
most_popular_department = department_counts.idxmax()

# Print the results
print("Most Popular Division:", most_popular_division)
print("Most Popular Group:", most_popular_group)
print("Most Popular Department:", most_popular_department)

plt.figure(figsize=(8, 6))
df["Division"].value_counts().plot(kind="bar")
plt.title("Most Popular Division")
plt.xlabel("Division")
plt.ylabel("Count")
plt.tight_layout()
plt.savefig("most_popular_division.png")
plt.close()

# Plotting the most popular group
plt.figure(figsize=(8, 6))
df["Group"].value_counts().plot(kind="bar")
plt.title("Most Popular Group")
plt.xlabel("Group")
plt.ylabel("Count")
plt.tight_layout()
plt.savefig("most_popular_group.png")
plt.close()

# Plotting the most popular department
plt.figure(figsize=(20, 12))
df["Department"].value_counts().plot(kind="bar")
plt.title("Most Popular Department")
plt.xlabel("Department")
plt.ylabel("Count")
plt.xticks(rotation=45, ha="right")
plt.tight_layout()
plt.savefig("most_popular_department.png")
plt.close()

df["Enrollment Time"] = pd.to_datetime(df["Enrollment Time"])

# Extract year and month from enrollment time
df["Enrollment Year"] = df["Enrollment Time"].dt.year
df["Enrollment Month"] = df["Enrollment Time"].dt.month

# Group the data by year and month and count the enrollments
enrollment_counts = (
    df.groupby(["Enrollment Year", "Enrollment Month"]).size().reset_index(name="Count")
)

# Plot the enrollment trend over time
plt.figure(figsize=(10, 6))
plt.plot(enrollment_counts.index, enrollment_counts["Count"], marker="o")
plt.title("Enrollment Trend Over Time")
plt.xlabel("Time")
plt.ylabel("Enrollment Count")
plt.xticks(
    range(len(enrollment_counts)),
    enrollment_counts["Enrollment Year"].astype(str)
    + "-"
    + enrollment_counts["Enrollment Month"].astype(str),
)
plt.xticks(rotation=45, ha="right")
plt.savefig("enrollment_trends.png")
plt.close()

# Calculate the completion rates
total_courses = df["Completed"].count()
df["Completed"] = df["Completed"].str.lower()
# Calculate the number of completed courses
completed_courses = df["Completed"].value_counts()["yes"]

# Calculate the completion rate
completion_rate = (completed_courses / total_courses) * 100

# Print the completion rate
print()
print(f"Completion Rate: {completion_rate:.2f}%")
print()

average_learning_hours = df["Learning Hours Spent"].mean()

# Calculate the total learning hours spent
total_learning_hours = df["Learning Hours Spent"].sum()

print("Average learning hours spent:", average_learning_hours)
print("Total learning hours spent:", total_learning_hours)
print()

# Remove any missing values
grades = df["Course Grade"].dropna()

# Calculate the grade distribution
grade_counts = grades.value_counts()

# Plot the grade distribution
plt.figure(figsize=(12, 6))  # Increase the figure size to accommodate the labels

# Customize the plot appearance
plt.bar(grade_counts.index, grade_counts.values)
plt.title("Course Grade Distribution")
plt.xlabel("Grade")
plt.ylabel("Count")
plt.xticks(rotation=45, ha="right")  # Rotate and align x-axis labels

plt.tight_layout()
plt.savefig("grade_distribution.png")
plt.close()

skills_column = df["Skills Learned"].str.lower()

# Combine all skills into a single string
all_skills = ";".join(skills_column.dropna().astype(str))

# Remove special characters and punctuation
all_skills = re.sub(r"[^a-zA-Z\s;]", "", all_skills)

# Split the string into individual skills
individual_skills = all_skills.split(";")

# Remove leading/trailing whitespaces from skills
individual_skills = [skill.strip() for skill in individual_skills]

# Remove empty skills (if any)
individual_skills = [skill for skill in individual_skills if skill]

# Count the frequency of each skill
skill_counts = Counter(individual_skills)

# Exclude "and" from the skill analysis
if "and" in skill_counts:
    del skill_counts["and"]

# Display the top 10 most frequently mentioned skills
top_skills = skill_counts.most_common(10)
print("Top 10 Skills Learned :\n")

for skill, count in top_skills:
    print(f"{skill}: {count}")

print()


division_insights = df.groupby("Division").agg(
    {
        "Duration(in hrs)": ["mean", "min", "max"],
        "Learning Hours Spent": ["sum"],
        "Course Grade": ["mean"],
    }
)

# Group the data by group and calculate insights
group_insights = df.groupby("Group").agg(
    {
        "Duration(in hrs)": ["mean", "min", "max"],
        "Learning Hours Spent": ["sum"],
        "Course Grade": ["mean"],
    }
)

# Display division-specific insights
print("Division-Specific Insights:")
print(division_insights)

# Display group-specific insights
print("\nGroup-Specific Insights:")
print(group_insights)


grouped_df = df.groupby("Department")

# Calculate insights for each department
for department, data in grouped_df:
    print(f"Department: {department}")

    # Calculate the total number of courses in each department
    total_courses = data["Course Name"].nunique()
    print(f"Total number of courses: {total_courses}")

    # Calculate the average duration of courses in each department
    avg_duration = data["Duration(in hrs)"].mean()
    print(f"Average duration of courses (in hours): {avg_duration:.2f}")

    # Calculate the percentage of completed courses in each department
    completed_courses = data[data["Completed"] == "yes"]
    completion_percentage = (completed_courses.shape[0] / data.shape[0]) * 100
    print(f"Completion percentage: {completion_percentage:.2f}%")

    # Calculate the total learning hours spent in each department
    total_learning_hours = data["Learning Hours Spent"].sum()
    print(f"Total learning hours spent: {total_learning_hours:.2f}")

    print("\n")


course_durations = df["Duration(in hrs)"]

# Drop missing values, if any
course_durations = course_durations.dropna()

# Calculate the range, average, and other summary statistics
duration_range = course_durations.max() - course_durations.min()
average_duration = course_durations.mean()
median_duration = course_durations.median()
mode_duration = course_durations.mode().values[0]

# Plot the distribution of course durations
plt.figure(figsize=(8, 6))
course_durations.hist(bins=10)
plt.title("Distribution of Course Durations")
plt.xlabel("Duration(in hrs)")
plt.ylabel("Frequency")
plt.grid(True)
plt.savefig("course_duration_analysis.png")
plt.close()

# Print the summary statistics
print("Course Duration Analysis:")
print("Range:", duration_range)
print("Average Duration:", average_duration)
print("Median Duration:", median_duration)
print("Mode Duration:", mode_duration)
print()

# Calculate course enrollment counts
course_enrollment_counts = df["Course Name"].value_counts()

# Calculate course completion counts
course_completion_counts = df[df["Completed"] == "Yes"]["Course Name"].value_counts()

# Calculate course completion rates
course_completion_rates = course_completion_counts / course_enrollment_counts

# Merge the completion rates into the original DataFrame
df["Enrollment Rate"] = df["Course Name"].map(course_enrollment_counts)
df["Completion Rate"] = df["Course Name"].map(course_completion_rates)

# Sort the DataFrame by enrollment rate in descending order
df_sorted_by_enrollment = df.sort_values(by="Enrollment Rate", ascending=False)

# Sort the DataFrame by completion rate in descending order
df_sorted_by_completion = df.sort_values(by="Completion Rate", ascending=False)

# Get the most popular course based on enrollment rate
most_popular_course_enrollment = df_sorted_by_enrollment.iloc[0]["Course Name"]

# Get the most popular course based on completion rate
most_popular_course_completion = df_sorted_by_completion.iloc[0]["Course Name"]

# Display the most popular course based on enrollment rate and completion rate
print("Most Popular Course based on Enrollment Rate:", most_popular_course_enrollment)
print("Most Popular Course based on Completion Rate:", most_popular_course_completion)
print()

# Convert Enrollment Time and Completion Time columns to datetime type
df["Enrollment Time"] = pd.to_datetime(df["Enrollment Time"])
df["Completion Time"] = pd.to_datetime(df["Completion Time"])

# Calculate the time taken for completion for each course in hours
df["Time Taken for Completion"] = (
    df["Completion Time"] - df["Enrollment Time"]
).dt.total_seconds() / 3600

# Calculate the average and median time taken for completion per program
average_time_per_program = df.groupby("Program Name")[
    "Time Taken for Completion"
].mean()
median_time_per_program = df.groupby("Program Name")[
    "Time Taken for Completion"
].median()

# Save results for programs to separate CSV files with adjusted column width
average_time_per_program.to_frame().to_csv(
    "average_time_per_program.csv",
    header=["Average Time Taken for Completion"],
    index_label="Program Name",
    float_format="%.2f",
)
median_time_per_program.to_frame().to_csv(
    "median_time_per_program.csv",
    header=["Median Time Taken for Completion"],
    index_label="Program Name",
    float_format="%.2f",
)

# Calculate the average and median time taken for completion per department
average_time_per_department = df.groupby("Department")[
    "Time Taken for Completion"
].mean()
median_time_per_department = df.groupby("Department")[
    "Time Taken for Completion"
].median()

# Save results for departments to separate CSV files with adjusted column width
average_time_per_department.to_frame().to_csv(
    "average_time_per_department.csv",
    header=["Average Time Taken for Completion"],
    index_label="Department",
    float_format="%.2f",
)
median_time_per_department.to_frame().to_csv(
    "median_time_per_department.csv",
    header=["Median Time Taken for Completion"],
    index_label="Department",
    float_format="%.2f",
)

# Replace 'yes' and 'no' in the 'Completed' column with binary values
df["Completed"] = df["Completed"].map({"Yes": 1, "No": 0})

# Convert 'Completion Time' and 'Course Grade' columns to numeric
df["Completion Time"] = pd.to_numeric(df["Completion Time"], errors="coerce")
df["Course Grade"] = pd.to_numeric(df["Course Grade"], errors="coerce")

# Drop rows with missing values in 'Completion Time' and 'Course Grade' columns
df.dropna(subset=["Completion Time", "Course Grade"], inplace=True)

# Calculate the average completion time and course grade for each course
course_data = (
    df.groupby("Course Name")
    .agg({"Completion Time": "mean", "Course Grade": "mean"})
    .reset_index()
)

# Calculate the correlation coefficient between completion time and course grade
correlation = course_data["Completion Time"].corr(course_data["Course Grade"])

# Plot the scatter plot using Seaborn
plt.figure(figsize=(8, 6))
sns.scatterplot(x="Completion Time", y="Course Grade", data=course_data)
plt.title("Correlation between Completion Time and Course Grade")
plt.xlabel("Average Completion Time")
plt.ylabel("Average Course Grade")
plt.grid(True)
plt.tight_layout()
plt.savefig("correlation_completion_time_course_grade.png")
plt.close()

print(f"Correlation between Completion Time and Course Grade: {correlation:.2f}")
print()

# Replace 'Yes' and 'No' with True and False in the Completed column
df["Completed"] = df["Completed"].replace({"Yes": True, "No": False})

# Group the data by Program Name and count the enrollments
enrollment_by_program = df.groupby("Program Name")["Name"].count().reset_index()

# Sort the programs by enrollment in descending order
enrollment_by_program = enrollment_by_program.sort_values(by="Name", ascending=False)

# Program with highest enrollment
highest_enrollment_program = enrollment_by_program.iloc[0]["Program Name"]
highest_enrollment_count = enrollment_by_program.iloc[0]["Name"]

# Program with lowest enrollment
lowest_enrollment_program = enrollment_by_program.iloc[-1]["Program Name"]
lowest_enrollment_count = enrollment_by_program.iloc[-1]["Name"]

# Print the results
print("Program with the Highest Enrollment: ", highest_enrollment_program)
print("Enrollment Count: ", highest_enrollment_count)

print("\nProgram with the Lowest Enrollment: ", lowest_enrollment_program)
print("Enrollment Count: ", lowest_enrollment_count)
print()
