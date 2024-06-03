import pandas as pd
from PyQt5.QtWidgets import QWidget, QVBoxLayout, QHBoxLayout, QComboBox, QPushButton, QTableWidget, QTableWidgetItem, QHeaderView, QTabWidget
from PyQt5.QtCore import Qt
from datetime import datetime, timedelta
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.figure import Figure

class PlotCanvas(FigureCanvas):
    def __init__(self, parent=None, width=5, height=4, dpi=100):
        fig = Figure(figsize=(width, height), dpi=dpi)
        self.ax = fig.add_subplot(111)
        super().__init__(fig)
        self.setParent(parent)

    def plot_histograms(self, data, title, xlabel):
        self.ax.clear()
        # Plot the bars
        data.plot(kind='bar', ax=self.ax, alpha=1, rot=0, zorder=3, color=['tab:cyan', 'tab:red'])  # Set a higher zorder for bars
        self.ax.set_title(title)
        self.ax.set_xlabel(xlabel)
        self.ax.set_ylabel('Length [m]')
        self.ax.tick_params(axis='x', direction='in', which='both')  # Ticks inside for x-axis
        self.ax.tick_params(axis='y', direction='in', which='both')  # Ticks inside for y-axis
        
        # Configure grid
        self.ax.grid(axis='y', color='gray', linestyle='-', linewidth=0.5, zorder=0)  # Ensure grid is behind bars with zorder=0
        
        # Setup legend with labels renamed
        legend_labels = ['Number', 'Scrap']  # Rename labels for clarity
        self.ax.legend(legend_labels)
        
        self.draw()

class Tab(QWidget):
    def __init__(self):
        super().__init__()
        self.initialize_ui()

    def initialize_ui(self):        
        self.Tab_tabs = QTabWidget()
        self.week_tab = QWidget()
        self.month_tab = QWidget()
        self.year_tab = QWidget()
        
        self.Tab_tabs.addTab(self.week_tab, "Week")
        self.Tab_tabs.addTab(self.month_tab, "Month")
        self.Tab_tabs.addTab(self.year_tab, "Year")
        
        Tab_layout = QVBoxLayout()
        Tab_layout.addWidget(self.Tab_tabs)
        self.setLayout(Tab_layout)

        self.setup_week_tab()
        self.setup_month_tab()
        self.setup_year_tab()
        self.load_initial_data()

    def setup_week_tab(self):
        layout = QVBoxLayout()
        self.year_combo = QComboBox()
        self.week_combo = QComboBox()
        self.report_button = QPushButton("Generate Report")
        self.report_button.clicked.connect(self.generate_report)
        
        self.table = QTableWidget()
        self.table.setColumnCount(6)
        self.table.setHorizontalHeaderLabels(["Number", "In [m]", "Out [m]", "Yield", "Error", "Test"]) # Allocate displayed table labels in program 
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)  # Adjust table width to match window width

        header = self.table.horizontalHeader()
        for i in range(self.table.columnCount()):
            header.setDefaultAlignment(Qt.AlignCenter)
        
        controls_layout = QHBoxLayout()
        controls_layout.addWidget(self.year_combo)
        controls_layout.addWidget(self.week_combo)
        controls_layout.addWidget(self.report_button)
        
        layout.addLayout(controls_layout)
        layout.addWidget(self.table)
        self.week_tab.setLayout(layout)

    def update_week_combo(self):
        if self.year_combo.count() == 0:
            return

        selected_year = int(self.year_combo.currentText())
        df_filtered = self.df[self.df['Year'] == selected_year]
        valid_weeks = sorted(df_filtered['Week'].dropna().unique(), reverse=True)

        self.week_combo.clear()
        self.week_combo.addItems([str(week) for week in valid_weeks])

    def setup_month_tab(self):
        layout = QVBoxLayout()
        controls_layout = QHBoxLayout()
        
        self.year_combo_month = QComboBox()
        self.month_combo = QComboBox()
        self.plot_button_month = QPushButton("Plot Monthly Data")
        self.plot_button_month.clicked.connect(self.plot_monthly_data)

        months = [datetime(2000, m, 1).strftime('%B') for m in range(1, 13)][::-1]  # Months in descending order
        self.month_combo.addItems(months)

        controls_layout.addWidget(self.year_combo_month)
        controls_layout.addWidget(self.month_combo)
        controls_layout.addWidget(self.plot_button_month)

        self.canvas_month = PlotCanvas(self, width=8, height=6)
        self.canvas_month.setVisible(False)

        layout.addLayout(controls_layout)
        layout.addWidget(self.canvas_month, 1)
        layout.addStretch()
        self.month_tab.setLayout(layout)

        self.year_combo_month.currentIndexChanged.connect(self.update_month_combo)

    def update_month_combo(self):
        selected_year = int(self.year_combo_month.currentText()) if self.year_combo_month.count() > 0 else None
        if selected_year is None:
            return

        df_filtered = self.df[self.df['Year'] == selected_year]
        available_months = sorted(df_filtered['Month'].dropna().unique(), reverse=True)

        self.month_combo.clear()
        self.month_combo.addItems([datetime(2000, int(m), 1).strftime('%B') for m in available_months])

    def setup_year_tab(self):
        layout = QVBoxLayout()
        controls_layout = QHBoxLayout()

        self.year_combo_year = QComboBox()
        self.plot_button_year = QPushButton("Plot Monthly Data")
        self.plot_button_year.clicked.connect(self.plot_yearly_data)

        controls_layout.addWidget(self.year_combo_year)
        controls_layout.addWidget(self.plot_button_year)

        self.canvas_year = PlotCanvas(self, width=8, height=6)
        self.canvas_year.setVisible(False)
        
        layout.addLayout(controls_layout)
        layout.addWidget(self.canvas_year, 1)
        layout.addStretch()
        self.year_tab.setLayout(layout)

    def center_table_items(self):
        for row in range(self.table.rowCount()):
            for col in range(self.table.columnCount()):
                item = self.table.item(row, col)
                if item:
                    item.setTextAlignment(Qt.AlignCenter)

    def load_initial_data(self):
        # Put here the Excel file path
        file_path = r""
        try:
            # Put in "sheets" name of tabs from Excel file like so sheets = ['Tab1', 'Tab2', 'Tab3']
            sheets = []
            dfs = []

            for sheet in sheets:
                # "skiprows" is how many rows to skip to find "Date" title of the column to indentify date
                df = pd.read_excel(file_path, sheet_name=sheet, skiprows=1,
                                   converters={'Date': lambda x: pd.to_datetime(x, format='%d.%m.%Y', errors='coerce')})
                df.dropna(subset=['Date'], inplace=True)
                # To check, if it was a test or not, can be omitted 
                df['Test'] = df['Test'].apply(lambda x: 'Yes' if str(x).strip().lower() == 'ja' else 'No')
                dfs.append(df)

            self.df = pd.concat([df for df in dfs if not df.empty], ignore_index=True)

            # Create the Year, Month, and Week columns
            self.df['Year'] = self.df['Date'].dt.year
            self.df['Month'] = self.df['Date'].dt.month
            self.df['Week'] = self.df['Date'].dt.isocalendar().week

            valid_years = self.df['Year'].dropna().unique()
            sorted_years = sorted(valid_years, reverse=True)

            for combo in [self.year_combo, self.year_combo_month, self.year_combo_year]:
                combo.clear()
                combo.addItems([str(int(year)) for year in sorted_years])

            self.year_combo.currentIndexChanged.connect(self.update_month_combo)
            self.update_month_combo()  # Call to fill month combo initially
            self.year_combo.currentIndexChanged.connect(self.update_week_combo)
            if self.year_combo.count() > 0:
                self.update_week_combo()  # Update immediately after loading

        except Exception as e:
            print(f"Failed to load data: {e}")

    def generate_report(self):
        try:
            selected_week = int(self.week_combo.currentText())
        except ValueError:
            print("No week selected.")
            return

        selected_year = int(self.year_combo.currentText())
        start_of_week = datetime.strptime(f'{selected_year} {selected_week-1} 1', "%Y %W %w")
        end_of_week = start_of_week + timedelta(days=6)

        # Put here the Excel file path
        file_path = r""
        try:
            # Put in "sheets" name of tabs from Excel file like so sheets = ['Tab1', 'Tab2', 'Tab3']
            sheets = []
            dfs = []

            for sheet in sheets:
                # "skiprows" is how many rows to skip to find "Date" title of the column to indentify date
                df = pd.read_excel(file_path, sheet_name=sheet, skiprows=1,
                                   converters={'Date': lambda x: pd.to_datetime(x, format='%d.%m.%Y', errors='coerce')})
                df_filtered = df[(df['Date'] >= start_of_week) & (df['Date'] <= end_of_week)]
                if not df_filtered.empty:
                    dfs.append(df_filtered)

            df_filtered = pd.concat(dfs, ignore_index=True)

            # Clear existing rows from the table
            self.table.setRowCount(0)

            # Populate the table with data from df_filtered from Excel file columns
            for index, row in df_filtered.iterrows():
                tape = str(row['Number']) if not pd.isnull(row['Number']) else ""
                in_val = int(row['Length']) if not pd.isnull(row['Length']) else 0
                out_val = in_val - int(row['Scrap']) if not pd.isnull(row['Scrap']) else 0
                yield_val = f"{(out_val / in_val * 100):.2f}%" if in_val != 0 else "N/A"
                error_val = str(row['Error']) if not pd.isnull(row['Error']) else ""
                test_val = 'Yes' if row['Test'] == 'Yes' else 'No'

                row_position = self.table.rowCount()
                self.table.insertRow(row_position)
                self.table.setItem(row_position, 0, QTableWidgetItem(tape))
                self.table.setItem(row_position, 1, QTableWidgetItem(str(in_val)))
                self.table.setItem(row_position, 2, QTableWidgetItem(str(out_val)))
                self.table.setItem(row_position, 3, QTableWidgetItem(yield_val))
                self.table.setItem(row_position, 4, QTableWidgetItem(error_val))
                self.table.setItem(row_position, 5, QTableWidgetItem(test_val))

                for col in range(self.table.columnCount()):
                    item = self.table.item(row_position, col)
                    if item:
                        item.setTextAlignment(Qt.AlignCenter)

        except Exception as e:
            print(f"Failed to generate report: {e}")

    def plot_monthly_data(self):
        if self.year_combo_month.count() == 0 or self.month_combo.count() == 0:
            print("No year or month selected.")
            return
        
        self.canvas_month.setVisible(True)
        selected_year = int(self.year_combo_month.currentText())
        selected_month = datetime.strptime(self.month_combo.currentText(), '%B').month

        try:
            # Filter for the selected year and month
            df_filtered = self.df[(self.df['Year'] == selected_year) & (self.df['Month'] == selected_month)]
            
            # Group by week and sum the relevant columns, f.i. "Length" and "Scrap"
            weekly_data = df_filtered.groupby('Week').agg({'Length': 'sum', 'Scrap': 'sum'})

            # Ensure the plot includes only weeks within the selected month
            weeks_in_month = df_filtered['Week'].unique()

            # Reindex, convert types, and fill missing values
            weekly_data = weekly_data.reindex(sorted(weeks_in_month)).infer_objects()
            weekly_data = weekly_data.fillna(0).astype(float)  # Fill missing weeks with zero

            self.canvas_month.plot_histograms(weekly_data, f"Weekly 'In' and 'Out' Data for {selected_year}, {self.month_combo.currentText()}", 'Week')

        except Exception as e:
            print(f"Failed to process or plot data: {e}")

    def plot_yearly_data(self):
        selected_year_text = self.year_combo_year.currentText()
        try:
            selected_year = int(float(selected_year_text))
        except ValueError as e:
            print(f"Error converting year: {e}")
            return
        
        self.canvas_year.setVisible(True)

        # Put here the Excel file path
        file_path = r""
        try:
            # Put in "sheets" name of tabs from Excel file like so sheets = ['Tab1', 'Tab2', 'Tab3']
            sheets = []
            dfs = []

            for sheet in sheets:
                # "skiprows" is how many rows to skip to find "Date" title of the column to indentify date
                df = pd.read_excel(file_path, sheet_name=sheet, skiprows=1,
                                   converters={'Date': lambda x: pd.to_datetime(x, format='%d.%m.%Y', errors='coerce')})
                if not df.empty:
                    dfs.append(df)

            df = pd.concat(dfs, ignore_index=True)

            # Handle non-numeric or corrupt data before processing
            df['Length'] = pd.to_numeric(df['Length'], errors='coerce').fillna(0)
            df['Scrap'] = pd.to_numeric(df['Scrap'], errors='coerce').fillna(0)

            df['Month'] = df['Date'].dt.month.dropna().astype(int)  # Safely convert month to integer
            df['Year'] = df['Date'].dt.year.dropna().astype(int)

            if df[df['Year'] == selected_year].empty:
                print(f"No data available for the year {selected_year}.")
                return

            # Aggregate data by month
            monthly_data = df[df['Year'] == selected_year].groupby('Month').agg({'Length': 'sum', 'Scrap': 'sum'})

            # Rename index to abbreviated month names
            monthly_data.index = [datetime(2000, month, 1).strftime('%b') for month in monthly_data.index.astype(int)]

            self.canvas_year.plot_histograms(monthly_data, f"Monthly 'In' and 'Out' Data for {selected_year}", 'Month')

        except Exception as e:
            print(f"Failed to process or plot data: {e}")
