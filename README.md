
# Data Analysis and Automated Report Generation

This project focuses on analyzing datasets and generating comprehensive reports in multiple formats, including HTML, PDF, and Word. The script automatically categorizes columns, detects and handles missing, duplicate, and constant values, and visualizes numeric and categorical data through charts.

## Features
- **Dataset Analysis**: 
  - Identifies missing, duplicate, and constant values.
  - Categorizes columns into numeric, categorical, boolean, and datetime types.
  - Cleans the dataset by removing duplicate and constant columns.
- **Visualizations**: 
  - Generates boxplots for numeric columns.
  - Creates distribution charts for up to six columns.
- **Automated Report Generation**:
  - Produces detailed reports in **HTML**, **PDF**, or **Word** formats.
  - Saves all visualizations to designated output folders.
  
## Folder Structure
```
project/
├── Data_Analysis2.py          # Main Python script
├── synthetic_health_data.csv  # Example dataset
├── report_output/             # Output folder for reports and visualizations
│   ├── report.pdf             # Generated report
│   ├── boxplots/              # Folder containing boxplots
│   └── distributions/         # Folder containing distribution charts
```

## Requirements
- Python 3.x
- Libraries:
  - `pandas`
  - `seaborn`
  - `matplotlib`
  - `fpdf2`
  - `os`

### Install Dependencies
Install the required Python libraries using the following command:
```bash
pip install pandas seaborn matplotlib fpdf2
```

## How to Run
1. Clone the repository:
   ```bash
   git clone https://github.com/your-username/data-analysis-report.git
   cd data-analysis-report
   ```
2. Place your dataset (CSV file) in the project directory.
3. Modify the script if needed, specifying the dataset file path in the function call:
   ```python
   generate_report("your_dataset.csv", output_format="pdf")
   ```
4. Run the script:
   ```bash
   python Data_Analysis2.py
   ```

## Output
- The analysis results are saved in the `report_output/` folder.
- Visualizations are stored in separate subfolders (`boxplots/` and `distributions/`).
- The final report is generated in the selected format (HTML, PDF, or Word).

## Example Usage
```python
# Generate a PDF report
generate_report("synthetic_health_data.csv", output_format="pdf")
```

## Notes
- Ensure that your dataset is in CSV format.
- Resolve any FPDF library conflicts by uninstalling older versions:
  ```bash
  pip uninstall --yes pypdf && pip install --upgrade fpdf2
  ```

## Screenshots
### Boxplots
![Example Boxplot](report_output/boxplots/example_boxplot.png)

### Distribution Charts
![Example Distribution Chart](report_output/distributions/example_distribution_chart.png)

## Author
**Habib Parvej**  
- [LinkedIn](https://www.linkedin.com/in/habibparvej)  
- [Email](mailto:habibparvej777@gmail.com)
