# EXCEL_PROJECT-DATA_ANALYTICS
# Global Data Job Market Analysis & Dashboard
# Excel Salary Dashboard
<img width="1900" height="982" alt="2026-06-18 (4)" src="https://github.com/user-attachments/assets/85a596da-3aae-40e0-9d34-4cc8feb9fd12" />


## Introduction

This data jobs salary dashboard was created to help job seekers investigate salaries for their desired jobs and ensure they are being adequately compensated. 

The data is from my Excel course, which provides a foundation in analyzing data using this powerful tool. The data contains detailed information on job titles, salaries, locations, and essential skills that are presented here.

### Dashboard File
My final dashboard is in [1_Salary_Dashboard.xlsx](1_Salary_Dashboard.xlsx).

### Excel Skills Used

The following Excel skills were utilized for analysis:

- **📉 Charts**
- **🧮 Formulas and Functions**
- **❎ Data Validation**

### Data Jobs Dataset

The dataset used for this project contains real-world data science job information from 2023. The dataset is available via my Excel course, which provides a foundation for analyzing data using Excel. It includes detailed information on:

- **👨‍💼 Job titles**
- **💰 Salaries**
- **📍 Locations**
- **🛠️ Skills**

## Dashboard Build

### 📉 Charts

#### 📊 Data Science Job Salaries - Bar Chart

<img width="599" height="375" alt="job title" src="https://github.com/user-attachments/assets/26b630ea-160d-4b24-9192-7a56fe90c6ca" />



- 🛠️ **Excel Features:** Utilized bar chart feature (with formatted salary values) and optimized layout for clarity.
- 🎨 **Design Choice:** Horizontal bar chart for visual comparison of median salaries.
- 📉 **Data Organization:** Sorted job titles by descending salary for improved readability.
- 💡 **Insights Gained:** This enables quick identification of salary trends, noting that Senior roles and Engineers are higher-paying than Analyst roles.

#### 🗺️ Country Median Salaries - Map Chart

<img width="511" height="360" alt="country" src="https://github.com/user-attachments/assets/b58d3ea0-44d0-48ca-b354-ffa996808c54" />


- 🛠️ **Excel Features:** Utilized Excel's map chart feature to plot median salaries globally.
- 🎨 **Design Choice:** Color-coded map to visually differentiate salary levels across regions.
- 📊 **Data Representation:** Plotted median salary for each country with available data.
- 👁️ **Visual Enhancement:** Improved readability and immediate understanding of geographic salary trends.
- 💡 **Insights Gained:** Enables quick grasp of global salary disparities and highlights high/low salary regions.

### 🧮 Formulas and Functions

#### 💰 Median Salary by Job Titles

```
=MEDIAN(
IF(
    (jobs[job_title_short]=$B2)*
    (jobs[job_country]=COUNTRY)*
    (ISNUMBER(SEARCH(TYPE,jobs[job_schedule_type])))*
    (jobs[salary_year_avg]<>0),
    jobs[salary_year_avg]
)
)
```

- 🔍 **Multi-Criteria Filtering:** Checks job title, country, schedule type, and excludes blank salaries.
- 📊 **Array Formula:** Utilizes `MEDIAN()` function with nested `IF()` statement to analyze an array.
- 🎯 **Tailored Insights:** Provides specific salary information for job titles, regions, and schedule types.
- **🔢 Formula Purpose:** This formula populates the table below, returning the median salary based on job title, country, and type specified.

🍽️ Background Table

<img width="1042" height="457" alt="median" src="https://github.com/user-attachments/assets/c525d3f3-c6d2-42a8-96d5-76f0e4c1e20b" />

📉 Dashboard Implementation

<img width="547" height="450" alt="type" src="https://github.com/user-attachments/assets/e4b2e54d-500b-4589-a651-c17a22256ec1" />


#### ⏰ Count of Job Schedule Type

```
=FILTER(A2#,(NOT(ISNUMBER(SEARCH("and",A2#))+ISNUMBER(SEARCH(",",A2#))))*(A2#<>0))
```

- 🔍 **Unique List Generation:** This Excel formula below employs the `FILTER()` function to exclude entries containing "and" or commas, and omit zero values.
- **🔢 Formula Purpose:** This formula populates the table below, which gives us a list of unique job schedule types.

🍽️ Background Table

<img width="774" height="724" alt="filter" src="https://github.com/user-attachments/assets/37eab13a-20bc-4a46-a4ed-08b53ed994c5" />


📉 Dashboard Implementation:

<img width="572" height="752" alt="type im" src="https://github.com/user-attachments/assets/6fc40bcf-384b-472c-b632-7589c3623866" />

### ❎ Data Validation

#### 🔍 Filtered List

- 🔒 **Enhanced Data Validation:** Implementing the filtered list as a data validation rule under the `Job Title`, `Country`, and `Type` option in the Data tab ensures:
    - 🎯 User input is restricted to predefined, validated schedule types
    - 🚫 Incorrect or inconsistent entries are prevented
    - 👥 Overall usability of the dashboard is enhanced

<img width="318" height="303" alt="validation" src="https://github.com/user-attachments/assets/52411cee-6e6c-4e93-b7ff-415459f58ac4" />

## Conclusion

I created this dashboard to showcase insights into salary trends across various data-related job titles. Utilizing data from my Excel course, this dashboard allows users to make informed decisions about their career paths. Exploring the functionalities to understand how location and job type influence salaries. 
# Project 2 Analysis

## Introduction

As a former job seeker, I’ve always been surprised by the lack of data exploring the most optimal jobs and skills in the data science market. I set out to understand what skills top employers request and how to land more pay.

### Questions to Analyze

To understand the data science job market, I asked the following:

1. **What are the top skills of data science?**
2. **Which categories yield the highest return relative to their market frequency?**
3. **What are the top skills of data professionals?**
4. **What’s the pay for the top 10 skills?**

### Excel Skills Used

The following Excel skills were utilized for analysis:

- **📊 Pivot Tables**
- **📈 Pivot Charts**
- **🧮 DAX (Data Analysis Expressions)**
- **🔍 Power Query**
- **💪 Power Pivot**

### Data Jobs Dataset

The dataset used for this project contains real-world data science job information from 2023. The dataset is available via my Excel course, which provides a foundation for analyzing data using Excel. 

It includes detailed information on:

- **👨‍💼 Job titles**
- **💰 Salaries**
- **📍 Locations**
- **🛠️ Skills**

## 1️⃣ What are the top skills of data science?

### 🔍 Skill: Power Query (ETL)

#### 📥 Extract

- I extracted the data using Power Query, where I appended all twelve monthly salary sheets into a single, consolidated dataset and create two queries:
    - 🗃️ First one with all the data jobs information.
    - 🔧 The second listing the skills for each job ID.

#### 🔄 Transform

- Then, I transformed each query by changing column types, removing unnecessary columns, cleaning text to eliminate specific words, and trimming excess whitespace.
    - 📊 data_jobs_all
    <img width="226" height="230" alt="powerquerym1" src="https://github.com/user-attachments/assets/81dd940d-7fa6-482a-b0d9-36ad9ef1f329" />


    - 🛠️ data_job_skills

       <img width="263" height="311" alt="powerquery12" src="https://github.com/user-attachments/assets/572f59ec-3251-46f5-87fd-6a7ec749863e" />


#### 🔗 Load

- Finally, I loaded both transformed queries into the workbook, setting the foundation for my subsequent analysis.
    - 📊 data_jobs_all

        <img width="1920" height="1042" alt="powerquerym" src="https://github.com/user-attachments/assets/a12759d7-abf2-43b0-bdb6-4c7d87861934" />


    - 🛠️ data_job_skills

      <img width="1920" height="1037" alt="powerquery" src="https://github.com/user-attachments/assets/4b5bd766-9993-4d86-8450-a98b39213c81" />


### 📊 Analysis

#### 💡 Insights

- 👑 The Core Duopoly: SQL (18,500 jobs) and Python (17,689 jobs) heavily dominate the data science ecosystem, appearing in nearly triple the number of job postings compared to the next closest tool.

📊 The Mid-Tier Essentials: A highly competitive group composed of Tableau (7,013 jobs), R (6,929 jobs), and AWS (6,842 jobs) follows, showcasing a balanced market demand across data visualization, statistical programming, and cloud infrastructure.

📈 Legacy vs. Big Data: Traditional data tools like Excel (6,258 jobs) still maintain a strong baseline presence, but are closely accompanied by specialized enterprise big data frameworks like Spark (5,274 jobs) and Azure (4,760 jobs).

   <img width="623" height="394" alt="ANAP21" src="https://github.com/user-attachments/assets/cf795a1f-3818-4f59-af48-69e8513e30d9" />


#### 🤔 So What
This chart serves as a direct prioritization roadmap for a data professional's learning path. It clearly demonstrates that while mastering niche or advanced cloud architectures is highly valuable, building an airtight foundation in SQL and Python is an absolute non-negotiable step for baseline marketability, as they form the operational backbone of the vast majority of open market roles.

## 2️⃣ Which categories yield the highest return relative to their market frequency?

### 🧮 Skills: PivotTables & DAX

#### 📈Pivot Table

- 🔢 I created a PivotTable using the Data Model I created with Power Pivot.
- 📊 I moved the `job_title_short` to the rows area.
- 🧮 Then I added new measure to calculate the median salary 
    
#### 🧮 DAX

- To calculate the median year salary I used DAX.

    ```
    Median Salary := MEDIAN(data_jobs_all[salary_year_avg])
    ```
###📈 PivotChart

I created a combo PivotChart to plot median salary and job count by job title from my PivotTable.
🪙 Primary Axis: Median Salary (as a Clustered Column)
👍 Secondary Axis: Job Count (as Diamond Markers)
To customize the chart, I added axis titles, removed the connecting lines for the job count series to isolate the individual data points, and sorted the job titles in descending order by median salary to clearly display the highest-paying roles.

###📊 Analysis

####💡 Insights

-💎 The Seniority Premium: Senior Data Scientist ($155,000) and Senior Data Engineer ($148,750) command the highest median salaries in the dataset, illustrating the substantial financial reward for advanced, specialized expertise.

⚖️ The Volume "Sweet Spot": Data Scientist ($127,500) and Data Engineer ($125,000) roles offer the most optimal balance in the market, delivering top-tier compensation while maintaining high job availability (between 6,700 and 8,500+ postings).

📉 High Demand, Baseline Pay: Data Analyst roles dominate the market in sheer volume with 9,606 postings, yet they sit near the lower end of the salary spectrum at a median of $90,000, proving that high market frequency does not always equal premium value.

<img width="961" height="494" alt="ANA2" src="https://github.com/user-attachments/assets/7b601f78-e914-41ee-a3e3-b475b856612a" />


#### **🤔 So What**

- This chart serves as a career roadmap for data professionals. It visually proves that while generalist entry-level positions like Data Analyst offer the most abundant job openings, long-term career growth and maximum earning potential depend on deliberate upskilling into specialized domains like Data Engineering and Data Science

## 3️⃣ What are the top skills of data professionals?

### 🔧 Skill:Power query & Power Pivot
- I extracted the data using Power Query, where I appended all twelve monthly salary sheets into a single, consolidated dataset and create two queries:
    - 🗃️ First one with all the data jobs information.
    - 🔧 The second listing the skills for each job ID.

#### 🔄 Transform

- Then, I transformed each query by changing column types, removing unnecessary columns, cleaning text to eliminate specific words, and trimming excess whitespace.
    - 📊 data_jobs_all
    <img width="226" height="230" alt="powerquerym1" src="https://github.com/user-attachments/assets/81dd940d-7fa6-482a-b0d9-36ad9ef1f329" />


    - 🛠️ data_job_skills

       <img width="263" height="311" alt="powerquery12" src="https://github.com/user-attachments/assets/572f59ec-3251-46f5-87fd-6a7ec749863e" />


#### 🔗 Load

- Finally, I loaded both transformed queries into the workbook, setting the foundation for my subsequent analysis.
    - 📊 data_jobs_all

        <img width="1920" height="1042" alt="powerquerym" src="https://github.com/user-attachments/assets/a12759d7-abf2-43b0-bdb6-4c7d87861934" />


    - 🛠️ data_job_skills

      <img width="1920" height="1037" alt="powerquery" src="https://github.com/user-attachments/assets/4b5bd766-9993-4d86-8450-a98b39213c81" />


#### 💪 Power Pivot

- 🔗 I created a data model by integrating the `data_jobs_all` and `data_jobs_skills` tables into one model.
- 🧹 Since I had already cleaned the data using Power Query; Power Pivot created a relationship between these two tables.

#### 🔗 Data Model

- I created a relationship between my two tables using the `job_id` column.

<img width="708" height="645" alt="powerpivot" src="https://github.com/user-attachments/assets/b82d7b64-5e65-404a-a404-18dc86d48f9f" />

#### 📃 Power Pivot Menu

- The Power Pivot menu was used to refine my data model and makes it easy to create measures.

 <img width="1920" height="1017" alt="powerpivot2" src="https://github.com/user-attachments/assets/654d5d54-3127-4250-8c66-0f758f8292b8" />


### 📊Analysis

#### 💡Insights

- 💻 SQL and Python dominate as top skills in data-related jobs, reflecting their foundational role in data processing and analysis.
- ☁️ Emerging technologies like AWS and Azure also show significant presence, underlining the industry's shift towards cloud services and big data technologies.

   <img width="623" height="394" alt="ANAP21" src="https://github.com/user-attachments/assets/67225187-bf24-4a6e-8361-443c473f8867" />



#### 🤔So What

- Understanding prevalent skills in the industry not only helps professionals stay competitive but also guides training and educational programs to focus on the most impactful technologies.

## 4️⃣ What’s the pay of the top 10 skills?

### 📊 Skill: Advanced Charts (Pivot Chart)

#### 📈 PivotChart

- I created a combo PivotChart to plot median salary and skill likelihood (%) from my PivotTable.
    - 🪙 **Primary Axis:** Median Salary (as a Clustered Column)
    - 👍 **Secondary Axis:** Skill Likelihood (as a Line with Markers)
- To customize the chart, I added a title axis title, removed the lines (skill likelihood), and changed the markers to diamonds.

### 📊 Analysis

#### 💡Insights

- 💰 Higher median salaries are associated with skills like Python, Oracle, and SQL, suggesting their critical role in high-paying tech jobs.
- 📉 Skills like PowerPoint and Word have the lowest median salaries and likelihood, indicating less specialization and demand in high-salary sectors.

  <img width="614" height="364" alt="ASAP3" src="https://github.com/user-attachments/assets/b12743b2-bd09-49ca-b96b-3eda4b278d89" />


### 🤔So What

- This chart highlights the importance of investing time in learning high-value skills like Python and SQL, which are evidently tied to higher paying roles, especially for those looking to maximize their salary in the tech industry.
