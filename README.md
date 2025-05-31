# Bookkeeper

**Bookkeeper** is a Python tool designed to analyze community expenses by processing Excel spreadsheets. It helps homeowners associations or residential communities better understand their spending and make data-driven decisions.

## Features

* 📊 Parses and analyzes Excel files with expense data
* 📁 Supports common formats used by community treasurers
* 💡 Provides summaries, insights, and potential anomalies in spending
* 🧾 Helps with budget tracking and financial transparency

## Installation

```bash
git clone https://github.com/mariohhd/bookkeeper.git
cd bookkeeper
pip install -r requirements.txt
```

## Usage

1. Run the analyzer:

```bash
python main.py
```

2. View the summary output in the console or export to a report file.

## Example Output

* Total monthly expenses
* Spending breakdown by category (e.g. cleaning, water)
* Year-over-year comparison
* Detection of unusual entries

## Requirements

* Python 3.8+
* `pandas`
* `openpyxl` or `xlrd`
* `pyyaml`
* (Optional) `matplotlib` or `seaborn` for visualizations


## Contributing

Contributions are welcome! Please fork the repository and open a pull request.

## License

This project is licensed under the **GNU General Public License v3.0**.
See the [LICENSE](./LICENSE) file for more details.
