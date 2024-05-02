### Production Plan Variations Analyzer

### Project Description:
This project aims to streamline the process of analyzing variations between two production plan roll-out files. It addresses the challenges faced by shift officers who traditionally spend significant time working with large datasets, performing VLOOKUP operations across multiple files to calculate differences between production week plans.

### Problem Statement:
The primary objective is to compare the week 2 production plans from two input files: File 1 (representing the current 13-week roll-out plan) and File 2 (representing the updated demand production plan for weeks 2 to 14). The focus is on identifying differences in week 2 plans between the two files, crucial for ordering the required raw materials (RM) and production materials (PM) in line with the updated production schedule.

### Key Features:
1. **Efficient Data Comparison**: The program efficiently compares the week 2 plans from File 1 and File 2, highlighting differences to facilitate decision-making.
2. **Visualization of Variations**: The output Excel file includes a visually appealing representation of differences, with positive values displayed in green and negative values in red.
3. **Handling Missing Products**: It addresses potential issues arising from missing products in either File 1 or File 2, providing separate sheets to display such data for further review.
4. **Automated Process**: Automation reduces manual effort and enhances line efficiency, empowering shift officers to focus on strategic tasks rather than data management.

### Implementation:
The project is implemented in Python, utilizing the Pandas library for efficient data processing and manipulation. It reads the input Excel files, extracts relevant data for comparison, calculates differences, and generates an output Excel file with three sheets:
- Sheet 1: Comparison of week 2 plans from File 1 and File 2, highlighting differences.
- Sheet 2: Products present in File 1 but not in File 2.
- Sheet 3: Products present in File 2 but not in File 1.

### Conclusion:
"Production Plan Variations Analyzer" simplifies and accelerates the process of analyzing production plan differences, enhancing operational efficiency and data management for shift officers. Its user-friendly interface and automated functionality make it a valuable tool for decision-making in production planning scenarios.

