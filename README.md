## PROJECT:
Excel Dashboard Creation Using VLOOKUP, Pivot Tables and Timeline in MS-Excel.


## STUDENT INFORMATION:

**Name:** Yash Sanjay Kadyan


# INTRODUCTION

In the realm of data analysis and business intelligence, Excel remains a powerful tool for transforming raw data into actionable insights. This project focuses on the creation of nine comprehensive dashboards using advanced Excel functionalities such as Pivot Tables, Timelines, and VLOOKUP. By leveraging these features, the dashboards provide a dynamic and interactive way to visualize data, uncover trends, and support decision-making processes.

The S10-Project.xlsx file serves as the foundational dataset for this project, containing a wealth of information that is meticulously organized and analyzed to produce the dashboards. Each dashboard is designed with specific objectives in mind, ranging from performance tracking and sales analysis to operational efficiency and financial summaries. The use of Pivot Tables allows for efficient data summarization and segmentation, while Timelines facilitate time-based filtering, enabling users to explore data across different periods seamlessly. Additionally, VLOOKUP functions are employed to integrate and cross-reference data from multiple sources within the file, ensuring comprehensive and accurate reporting.

This project demonstrates the power of Excel as a versatile tool for dashboard creation, showcasing its potential to drive business insights and strategic planning.


# Implementation

 Pivot Tables:

• Create Pivot Tables to summarize and analyze data from the ‘Compiled Data’, ‘Employees’, ‘Categories’, ‘Customers’, ‘EmployeeTerritories’, ‘Order_Details’, ‘Orders’, ‘Products’, ‘Region’, ‘Shippers’, ‘Suppliers’ and ‘Territories’ sheets.


 VLOOKUP & Data Cross-Referencing:

• Implement VLOOKUP formulas to link data across sheets. For example, cross-reference Order IDs in the ‘Order_Details’ sheet with Employees in the ‘Employees’ sheet to display the Employees Full Names.


 Timeline Visualization:

• Utilize the Timeline feature in Excel to create an interactive visual of the project’s analysis Year wise. This will help in tracking the project’s progress Year wise and identifying any delays in real-time.

• Apply a filter within the timeline to highlight the project's progress for a particular Year based on the Year filter.


 Dashboard Integration:

• Build an integrated dashboard combining insights from all sheets. Use charts and graphs to display key metrics such as monthly sales trend yearwise, Yearly Sales and overall project progress.

• Implement interactive controls like dropdowns and slicers for dynamic data filtering, allowing users to customize their view of the project data.


# PROJECT EXPLANATION

The S10-Project.xlsx file is designed to create a comprehensive and interactive Excel dashboard that provides a clear visualization and summary of data using advanced Excel features such as Pivot Tables, Timelines and VLOOKUP. The dashboard offers an insightful analysis of the project's key metrics, enabling efficient decision-making and progress tracking.


## Data Structure:

	• Sheet 1: Compiled Data

		• Contains the primary dataset used for analysis. This sheet includes information such as Order IDs, Sales, Freight, Employee IDs and Customer IDs.
		• The data is organized in tabular form with each column representing a specific attribute, such as "OrderID", "OrderDate", "ShipVia", "Shipper CompanyName", "ProductID", etc.
		• This raw data serves as the source for creating Pivot Tables and performing VLOOKUP operations to derive calculated metrics.


## Pivot Tables:

Pivot Tables are used extensively in this project to summarize and analyze the raw data. They provide dynamic views of the data, allowing users to slice and dice the information in various ways. Key implementations include:

	• Employee-wise Sales Analysis Pivot Table:

		• Summarizes Employees by their Sum of Sales.				• Displays the Sum of Sales and their respective Employees, enabling quick identification of Employees with the Lowest Sum of Sales.
		• Allows filtering by Employees to monitor individual workload and progress.

	• Category-wise Sales Pivot Table:

		• Shows the distribution of Sum of Sales among Categories, indicating who has the lowest percentage among all the Categories.
		• Displays the Sum of Sales and their respective Categories, enabling quick identification of Categories with the Lowest Sum of Sales Percentage-wise.

	• Shipper Company-wise Freight Pivot Table:

		• Tracks the Freight of each Shipper Company.
		• Displays the Freight for each Shipper Company, helping the project managers assess progress.


## VLOOKUP:

The VLOOKUP function is used to cross-reference data between different sheets, enabling the integration of related information for comprehensive analysis:

	• Employee Sales Mapping:

		• Uses VLOOKUP to pull detailed Employee information (e.g., EmployeeID,Employee FullName) from the "Employees" sheet into the main "Compiled Data" sheet.
		• Allows the dashboard to display contextual information for each Employee, such as the Employee's Sales Year-wise.

	• Sum of Sales Per Order ID:
	
		• VLOOKUP is used to fetch information about Sum of Sales that require immediate attention.
		• This helps in highlighting Sum of Sales for a particular year, thus prioritizing them in the dashboard.


## Timelines:

Timelines are used to provide a visual representation of the project's schedule, offering an easy way to monitor progress over time:

	• Timeline in Years:

		• Displays the timeline in Years using a Gantt chart format.
		• The timeline is interactive, allowing users to filter by time period (e.g., yearly) to view Sum of Sales scheduled within a specific range. (i.e. Time Period in Years)


## Dashboard Design:

The dashboard integrates multiple elements like Pivot Tables, Charts, and Timelines to present a cohesive view of the project data. Key components include:

	• Interactive Charts:

		• Bar charts and pie charts are used to visually represent data from the Pivot Tables, such as the distribution of Sum of Sales by Employees and Categories.
		• These charts update dynamically when filters are applied, providing real-time insights.

	• Filters:

		• Dropdown filters offer additional flexibility in data selection, making the dashboard highly interactive and user-friendly.


## User Experience Enhancements:

	• Dynamic Updates:

		• The dashboard is designed to update automatically when new data is added to the "Compiled Data" sheet, using dynamic ranges and named ranges to ensure that Pivot Tables and Charts reflect the most current information.
		• This eliminates the need for manual updates, improving efficiency and reducing the risk of errors.

	• Interactive Elements:

		• Users can interact with the dashboard through clickable elements like buttons and dropdowns, enabling a customized view of the data based on specific requirements.
		• Timelines and charts are linked to these interactive controls, providing an intuitive and engaging user experience.

This project leverages advanced Excel functionalities to create an interactive, insightful and user-friendly dashboard that effectively communicates the Sum of Sales of the Project Year-wise.


# USAGE

**Viewing Sales Summary:**

	• Navigate to the “Compiled Data” sheet to view a summary of all project tasks.
	• Hover over the bar or pie charts for detailed data points, such as the exact number of Sales for each Region.


**Analyzing Sum of Sales:**

	• View the Pivot Table to get the Sum of Sales assigned to each Region, along with a visual representation in the accompanying chart.


**Exploring Data with Timelines:**

	• Use the Timeline control to filter data by specific year ranges, such as viewing the sum of sales completed within the last year.
	• Drag the Timeline slider to adjust the year range dynamically, and watch the Pivot Tables and charts update in real time.


**Utilizing VLOOKUP for Detailed Information:**

	• VLOOKUP is used in various sheets to pull in additional data. For example, in the “Compiled Data” sheet, employee details are pulled from a separate sheet: “Employees” to provide context on who is responsible for each Sum of Sales.
	• To see a detailed Sum of Sales information for a specific Employee, hover over or click on the relevant Employee Full Name entry in the Pivot Table.


**Updating the Dashboard:**

	• When new data is added to the “Compiled Data” sheet, refresh all Pivot Tables to incorporate the latest information.
	• Use the “Refresh All” button in the Data tab or a custom macro button if provided, to update all data connections, ensuring the dashboard reflects the most current status of the project.


**Exporting the Dashboard:**

	• To share the dashboard with stakeholders, export it as a PDF or print it directly from Excel.
	• Use the “File” menu to choose “Save As” and select PDF format, ensuring all visible sheets and charts are included in the export.


This usage guide provides a step-by-step walkthrough of how to effectively interact with and gain insights from the Excel Dashboard created using Pivot Tables, Timeline, and VLOOKUP in the S10-Project.xlsx file.


# CONCLUSION

In conclusion, the Excel Dashboard project developed using Pivot Tables, VLOOKUP and Timelines has successfully provided an advanced and interactive solution for monitoring and managing project data within the S10-Project.xlsx file. By transforming raw data into meaningful insights, the dashboard facilitates effective decision-making and project progress tracking.

The use of Pivot Tables enables dynamic data summarization and analysis, while VLOOKUP seamlessly integrates cross-sheet data to provide comprehensive insights. The Timelines feature offers a clear visual representation of project sales, enhancing the ability to monitor project timelines and identify potential delays.

This project has met its objectives by delivering a user-friendly, visually appealing, and highly functional dashboard that caters to the diverse analytical needs of project managers and stakeholders. As a robust tool for project management and reporting, it offers a clear, consolidated view of project performance, making it easier to identify trends, track progress, and optimize resource utilization.

With its flexible design and powerful analytical capabilities, the dashboard is well-positioned to evolve and adapt to future project needs, ensuring its continued relevance and value in facilitating data-driven decision-making and project oversight.


# PROJECT SCREENSHOTS

