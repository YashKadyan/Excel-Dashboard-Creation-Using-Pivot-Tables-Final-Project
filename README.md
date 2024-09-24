## EXCEL ANALYTICS PROJECT:
Excel Dashboard Creation Using VLOOKUP, Pivot Tables and Timeline in MS-Excel.


## STUDENT INFORMATION:

**Name:** Yash Sanjay Kadyan


# INTRODUCTION

In the realm of data analysis and business intelligence, Excel remains a powerful tool for transforming raw data into actionable insights. This project focuses on the creation of nine comprehensive dashboards using advanced Excel functionalities such as Pivot Tables, Timelines, and VLOOKUP. By leveraging these features, the dashboards provide a dynamic and interactive way to visualize data, uncover trends, and support decision-making processes.

The S10-Project.xlsx file serves as the foundational dataset for this project, containing a wealth of information that is meticulously organized and analyzed to produce the dashboards. Each dashboard is designed with specific objectives in mind, ranging from performance tracking and sales analysis to operational efficiency and financial summaries. The use of Pivot Tables allows for efficient data summarization and segmentation, while Timelines facilitate time-based filtering, enabling users to explore data across different periods seamlessly. Additionally, VLOOKUP functions are employed to integrate and cross-reference data from multiple sources within the file, ensuring comprehensive and accurate reporting.

This project demonstrates the power of Excel as a versatile tool for dashboard creation, showcasing its potential to drive business insights and strategic planning.


# Implementation

 **Pivot Tables:**

• Create Pivot Tables to summarize and analyze data from the ‘Compiled Data’, ‘Employees’, ‘Categories’, ‘Customers’, ‘EmployeeTerritories’, ‘Order_Details’, ‘Orders’, ‘Products’, ‘Region’, ‘Shippers’, ‘Suppliers’ and ‘Territories’ sheets.


 **VLOOKUP & Data Cross-Referencing:**

• Implement VLOOKUP formulas to link data across sheets. For example, cross-reference Order IDs in the ‘Order_Details’ sheet with Employees in the ‘Employees’ sheet to display the Employees Full Names.


 **Timeline Visualization:**

• Utilize the Timeline feature in Excel to create an interactive visual of the project’s analysis Year wise. This will help in tracking the project’s progress Year wise and identifying any delays in real-time.

• Apply a filter within the timeline to highlight the project's progress for a particular Year based on the Year filter.


 **Dashboard Integration:**

• Build an integrated dashboard combining insights from all sheets. Use charts and graphs to display key metrics such as monthly sales trend yearwise, Yearly Sales and overall project progress.

• Implement interactive controls like dropdowns and slicers for dynamic data filtering, allowing users to customize their view of the project data.


# PROJECT EXPLANATION

The S10-Project.xlsx file is designed to create a comprehensive and interactive Excel dashboard that provides a clear visualization and summary of data using advanced Excel features such as Pivot Tables, Timelines and VLOOKUP. The dashboard offers an insightful analysis of the project's key metrics, enabling efficient decision-making and progress tracking.


## Data Structure:

 **Sheet 1: Compiled Data**

• Contains the primary dataset used for analysis. This sheet includes information such as Order IDs, Sales, Freight, Employee IDs and Customer IDs.

• The data is organized in tabular form with each column representing a specific attribute, such as "OrderID", "OrderDate", "ShipVia", "Shipper CompanyName", "ProductID", etc.

• This raw data serves as the source for creating Pivot Tables and performing VLOOKUP operations to derive calculated metrics.


## Pivot Tables:

Pivot Tables are used extensively in this project to summarize and analyze the raw data. They provide dynamic views of the data, allowing users to slice and dice the information in various ways. Key implementations include:


**1. Employee-wise Sales Analysis Pivot Table:**

• Summarizes Employees by their Sum of Sales.				

• Displays the Sum of Sales and their respective Employees, enabling quick identification of Employees with the Lowest Sum of Sales.

• Allows filtering by Employees to monitor individual workload and progress.


**2. Category-wise Sales Pivot Table:**

• Shows the distribution of Sum of Sales among Categories, indicating who has the lowest percentage among all the Categories.

• Displays the Sum of Sales and their respective Categories, enabling quick identification of Categories with the Lowest Sum of Sales Percentage-wise.


**3. Shipper Company-wise Freight Pivot Table:**

• Tracks the Freight of each Shipper Company.

• Displays the Freight for each Shipper Company, helping the project managers assess progress.


## VLOOKUP:

The VLOOKUP function is used to cross-reference data between different sheets, enabling the integration of related information for comprehensive analysis:


**1. Employee Sales Mapping:**

• Uses VLOOKUP to pull detailed Employee information (e.g., EmployeeID,Employee FullName) from the "Employees" sheet into the main "Compiled Data" sheet.

• Allows the dashboard to display contextual information for each Employee, such as the Employee's Sales Year-wise.


**2. Sum of Sales Per Order ID:**
	
• VLOOKUP is used to fetch information about Sum of Sales that require immediate attention.

• This helps in highlighting Sum of Sales for a particular year, thus prioritizing them in the dashboard.


## Timelines:

Timelines are used to provide a visual representation of the project's schedule, offering an easy way to monitor progress over time:


**1. Timeline in Years:**

• Displays the timeline in Years using a Gantt chart format.

• The timeline is interactive, allowing users to filter by time period (e.g., yearly) to view Sum of Sales scheduled within a specific range. (i.e. Time Period in Years)


## Dashboard Design:

The dashboard integrates multiple elements like Pivot Tables, Charts, and Timelines to present a cohesive view of the project data. Key components include:


**1. Interactive Charts:**

• Bar charts and pie charts are used to visually represent data from the Pivot Tables, such as the distribution of Sum of Sales by Employees and Categories.

• These charts update dynamically when filters are applied, providing real-time insights.


**2. Filters:**

• Dropdown filters offer additional flexibility in data selection, making the dashboard highly interactive and user-friendly.


## User Experience Enhancements:


**1. Dynamic Updates:**

• The dashboard is designed to update automatically when new data is added to the "Compiled Data" sheet, using dynamic ranges and named ranges to ensure that Pivot Tables and Charts reflect the most current information.

• This eliminates the need for manual updates, improving efficiency and reducing the risk of errors.


**2. Interactive Elements:**

• Users can interact with the dashboard through clickable elements like buttons and dropdowns, enabling a customized view of the data based on specific requirements.

• Timelines and charts are linked to these interactive controls, providing an intuitive and engaging user experience.

This project leverages advanced Excel functionalities to create an interactive, insightful and user-friendly dashboard that effectively communicates the Sum of Sales of the Project Year-wise.


# USAGE


**A. Viewing Sales Summary:**

• Navigate to the “Compiled Data” sheet to view a summary of all project tasks.

• Hover over the bar or pie charts for detailed data points, such as the exact number of Sales for each Region.


**B. Analyzing Sum of Sales:**

• View the Pivot Table to get the Sum of Sales assigned to each Region, along with a visual representation in the accompanying chart.


**C. Exploring Data with Timelines:**

• Use the Timeline control to filter data by specific year ranges, such as viewing the sum of sales completed within the last year.

• Drag the Timeline slider to adjust the year range dynamically, and watch the Pivot Tables and charts update in real time.


**D. Utilizing VLOOKUP for Detailed Information:**

• VLOOKUP is used in various sheets to pull in additional data. For example, in the “Compiled Data” sheet, employee details are pulled from a separate sheet: “Employees” to provide context on who is responsible for each Sum of Sales.

• To see a detailed Sum of Sales information for a specific Employee, hover over or click on the relevant Employee Full Name entry in the Pivot Table.


**E. Updating the Dashboard:**

• When new data is added to the “Compiled Data” sheet, refresh all Pivot Tables to incorporate the latest information.

• Use the “Refresh All” button in the Data tab or a custom macro button if provided, to update all data connections, ensuring the dashboard reflects the most current status of the project.


**F. Exporting the Dashboard:**

• To share the dashboard with stakeholders, export it as a PDF or print it directly from Excel.

• Use the “File” menu to choose “Save As” and select PDF format, ensuring all visible sheets and charts are included in the export.


This usage guide provides a step-by-step walkthrough of how to effectively interact with and gain insights from the Excel Dashboard created using Pivot Tables, Timeline, and VLOOKUP in the S10-Project.xlsx file.


# CONCLUSION

In conclusion, the Excel Dashboard project developed using Pivot Tables, VLOOKUP and Timelines has successfully provided an advanced and interactive solution for monitoring and managing project data within the S10-Project.xlsx file. By transforming raw data into meaningful insights, the dashboard facilitates effective decision-making and project progress tracking.

The use of Pivot Tables enables dynamic data summarization and analysis, while VLOOKUP seamlessly integrates cross-sheet data to provide comprehensive insights. The Timelines feature offers a clear visual representation of project sales, enhancing the ability to monitor project timelines and identify potential delays.

This project has met its objectives by delivering a user-friendly, visually appealing, and highly functional dashboard that caters to the diverse analytical needs of project managers and stakeholders. As a robust tool for project management and reporting, it offers a clear, consolidated view of project performance, making it easier to identify trends, track progress, and optimize resource utilization.

With its flexible design and powerful analytical capabilities, the dashboard is well-positioned to evolve and adapt to future project needs, ensuring its continued relevance and value in facilitating data-driven decision-making and project oversight.


# PROJECT SCREENSHOTS

![List of Dashboards from 1 to 6 and Compiled Data Sheet](https://github.com/user-attachments/assets/cdd39de4-757e-4466-b752-4b4f1bb3dc0b)

![List of Dashboards from 7 to 9 and Compiled Data Sheet](https://github.com/user-attachments/assets/a34fa3b2-934f-49b2-bb2f-3a6f7dfd8426)

![List of Dashboards from 9 to Not Possible to Plot and Compiled Data Sheet](https://github.com/user-attachments/assets/548166f5-f4f3-41b1-b8fb-cc3dbde1dbfc)

![Dashboard 1 for S10-Project](https://github.com/user-attachments/assets/2908fdb7-e0dc-41d6-a945-71d4e78c9dfa)

![Dashboard 1 for S10-Project for the year 1996](https://github.com/user-attachments/assets/581dcc5c-6799-4e03-a2cc-e1e4c2d6c421)

![Dashboard 1 for S10-Project for the year 1997](https://github.com/user-attachments/assets/229786e8-680d-4452-85f5-39e80b9f2803)

![Dashboard 1 for S10-Project for the year 1998](https://github.com/user-attachments/assets/0c73e66a-3e22-4c7c-9a7f-227802036bbc)

![Dashboard 2 for S10-Project](https://github.com/user-attachments/assets/3d140481-7d02-4ebb-ad04-869127458b27)

![Dashboard 2 for S10-Project for the year 1996](https://github.com/user-attachments/assets/e7e4c50c-9640-49e1-8cda-244d1316f143)

![Dashboard 2 for S10-Project for the year 1997](https://github.com/user-attachments/assets/32732bbc-459b-4dd0-b59d-3a41d1dfafc2)

![Dashboard 2 for S10-Project for the year 1998](https://github.com/user-attachments/assets/b5c69c22-71f1-4aac-9add-f32e20a47eff)

![Dashboard 3 for S10-Project - 1](https://github.com/user-attachments/assets/af911950-784b-440c-ad25-d58597503b2c)

![Dashboard 3 for S10-Project - 2](https://github.com/user-attachments/assets/7157a438-ad1e-4834-8bb6-1b71e93fecc0)

![Dashboard 3 for S10-Project for the year 1996](https://github.com/user-attachments/assets/81fc1c80-e7a7-4e1b-ad17-9596c27120f8)

![Dashboard 3 for S10-Project for the year 1997](https://github.com/user-attachments/assets/f6792990-c61f-4823-999f-db522b0df5ad)

![Dashboard 3 for S10-Project for the year 1998](https://github.com/user-attachments/assets/15bbeb95-5921-4d0e-8a3f-19265beb4660)

![Dashboard 4 for S10-Project - 1](https://github.com/user-attachments/assets/03ffc0d3-2270-44b1-94ef-6f1158ce2523)

![Dashboard 4 for S10-Project - 2](https://github.com/user-attachments/assets/1c2abdd2-c5c2-4298-ae6f-be210264e0b7)

![Dashboard 4 for S10-Project for the year 1996](https://github.com/user-attachments/assets/a0f4c78b-7f93-4190-9646-66bc3701aa78)

![Dashboard 4 for S10-Project for the year 1997](https://github.com/user-attachments/assets/6ac994b9-8249-4fd8-888b-f77e79f1cd8e)

![Dashboard 4 for S10-Project for the year 1998](https://github.com/user-attachments/assets/f4b2c0b6-ae8f-4ece-ab03-0745da47a115)

![Dashboard 5 for S10-Project - 1](https://github.com/user-attachments/assets/df14e273-253c-4a3a-8b62-9cc5bf87b6dc)

![Dashboard 5 for S10-Project - 2](https://github.com/user-attachments/assets/27a78d2b-1717-4191-86a1-0510a49db15b)

![Dashboard 5 for S10-Project for the year 1996](https://github.com/user-attachments/assets/099c6b3b-0c5b-4296-9ab8-823926b88b48)

![Dashboard 5 for S10-Project for the year 1997](https://github.com/user-attachments/assets/738b01f5-5755-4385-83ce-6621bb20925f)

![Dashboard 5 for S10-Project for the year 1998](https://github.com/user-attachments/assets/00ce88ea-9b18-4a4a-add6-cdbc013d0206)

![Dashboard 6 for S10-Project](https://github.com/user-attachments/assets/bb63e7ca-1ccf-4694-aa52-46f36797b151)

![Dashboard 6 for S10-Project for the year 1996](https://github.com/user-attachments/assets/fdeffa55-af5d-46d6-8889-07495bc216b3)

![Dashboard 6 for S10-Project for the year 1997](https://github.com/user-attachments/assets/debdabd4-27f2-4942-ae17-e2ed6862c3c4)

![Dashboard 6 for S10-Project for the year 1998](https://github.com/user-attachments/assets/ceb4c752-a30d-4506-ad69-03b326e5a9dc)

![Dashboard 7 for S10-Project](https://github.com/user-attachments/assets/ce1d1b84-bff4-4051-9206-de556ad4abef)

![Dashboard 7 for S10-Project for the year 1996](https://github.com/user-attachments/assets/c9befd74-2ca1-4410-946a-7e21bf28884f)

![Dashboard 7 for S10-Project for the year 1997](https://github.com/user-attachments/assets/5f6f1f76-1807-4666-aa73-05e2eb2bef0f)

![Dashboard 7 for S10-Project for the year 1998](https://github.com/user-attachments/assets/73a01f71-2201-44b9-98ee-1190e6683f4d)

![Dashboard 8 for S10-Project - 1](https://github.com/user-attachments/assets/bbe4071a-980e-4b5e-932e-77bcd0e8c641)

![Dashboard 8 for S10-Project - 2](https://github.com/user-attachments/assets/d0d23142-e06e-49d0-8b66-34abeaf60ce9)

![Dashboard 8 for S10-Project for the year 1996](https://github.com/user-attachments/assets/56221986-e01d-4232-a920-09415b3cf48f)

![Dashboard 8 for S10-Project for the year 1997](https://github.com/user-attachments/assets/2bd0a203-063a-45ec-9b87-4eb38784a733)

![Dashboard 8 for S10-Project for the year 1998](https://github.com/user-attachments/assets/fe6af18e-5ecc-47ed-a44c-52b5d3418c29)

![Dashboard 9 for S10-Project](https://github.com/user-attachments/assets/b36b33bb-e9d1-4ff4-8b00-da5c92192215)

![Dashboard 9 for S10-Project for the year 1996](https://github.com/user-attachments/assets/e49de3ee-1b5d-41cb-9c63-5e2bd932e2a6)

![Dashboard 9 for S10-Project for the year 1997](https://github.com/user-attachments/assets/fea966ef-bb90-4eb5-9c87-aaccdfd2dd6c)

![Dashboard 9 for S10-Project for the year 1998](https://github.com/user-attachments/assets/75008d47-d135-4585-b0eb-b55a35f8a6d5)
