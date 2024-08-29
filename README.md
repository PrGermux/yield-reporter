# Yield Reporter

This Python project is a comprehensive data analysis and visualization tool built with Python, leveraging PyQt5 for the graphical user interface and Matplotlib for data plotting. The application is specifically designed to manage, aggregate, and visualize production data from various machines, offering both tabular and graphical representations of the data.

#### Key Features:
- **Multi-Tab Interface**: The application has multiple tabs for different views - weekly, monthly, and yearly data analysis.
- **Data Aggregation and Filtering**: It supports filtering data by year and month, aggregating the lengths of tapes and scraps produced by various machines.
- **Dynamic Plotting**: Utilizes Matplotlib to create dynamic, interactive bar charts with annotations showing the percentage contributions of each machine and scrap per month.
- **Customizable Plots**: The plot customization includes setting bar colors, grid configurations, and dynamic legend creation for clarity and better user understanding.
- **Data Integration**: Combines data from multiple Excel sheets, ensuring comprehensive analysis across different data sources.

#### Usage
This tool is highly useful for production managers, data analysts, and engineers who need to monitor, analyze, and report on machine performance and production efficiency. By providing both numerical summaries and graphical visualizations, the application helps in identifying trends, inefficiencies, and areas for improvement in production processes.

#### Python Branch and Complexity
- **Python Branch**: This project utilizes several advanced Python libraries, including PyQt5 for GUI development, Pandas for data manipulation, and Matplotlib for plotting. These libraries indicate a high-level proficiency in Python, especially in data analysis and visualization domains.
- **Complexity**: The project is moderately complex, involving GUI design, data processing, and dynamic plotting. The integration of multiple data sources and the need for interactive and annotated visualizations add to its complexity. The code demonstrates good practices in object-oriented programming and modular design.

#### Code Structure
- **Main Interface**: The main GUI is structured using QTabWidget to separate different views (weekly, monthly, yearly).
- **Data Handling**: Data is read from multiple Excel sheets, cleaned, and aggregated using Pandas.
- **Plotting**: Matplotlib is used extensively for creating bar charts with detailed annotations and custom legends.

#### Future Enhancements
- **Real-Time Data Integration**: Incorporate real-time data fetching and updating mechanisms.
- **Enhanced Customization**: Allow users to customize plots and reports further through the GUI.
- **Additional Data Sources**: Extend support for other data formats and sources, such as databases or APIs.

This repository is a valuable resource for professionals in manufacturing and production environments, providing powerful tools for data-driven decision-making and operational efficiency improvements.

#### Screenshots:

![grafik](https://github.com/user-attachments/assets/6609d257-4b27-4c74-b053-3d2ac4d9c28f)
![grafik](https://github.com/user-attachments/assets/fab917e0-483f-4d8b-8306-a69da0b3a5a5)
![grafik](https://github.com/user-attachments/assets/8d3bde84-9d53-46ae-b492-7c971948914e)
![grafik](https://github.com/user-attachments/assets/42678be6-20c7-4a7f-959d-7c61603f0227)
![grafik](https://github.com/user-attachments/assets/d2c5c225-0a74-4d89-aa24-fb3243bd0283)
![grafik](https://github.com/user-attachments/assets/e48a947b-dd2e-46a5-ac24-783220ada4f1)
![grafik](https://github.com/user-attachments/assets/17904771-0105-4c31-9650-d9d4eb1a3939)

## Installation
1. Clone the repository:
```sh
git clone https://github.com/yourusername/production-data-analysis.git
cd production-data-analysis
```
2. Install the required packages:
```sh
pip install -r requirements.txt
```

## Usage
**WARNING:** This programm works only with specific data files. The user must adjust each tab to his/her needs with the respect to the structure of Excel-file.

Run the main application:
```sh
python main.py
```
The source code of provided screenshots can be found in "Yield Reporter.zip".

## Freezing
Run the code in a command line:
```sh
pyinstaller --onefile --windowed --icon=icon.png --add-data "icon.png;." --hidden-import=scipy.special._cdflib --name "Yield Reporter" main.py
```

## Dependencies
- Python 3.x
- PyQt5
- Pandas
- Matplotlib

## License
This project is licensed under the MIT License for non-commercial use.
