Aqui está o conteúdo do `README.md` para o seu código, formatado corretamente para você copiar e colar:

```markdown
# Opsgenie Alerts to Excel

This Python script retrieves alerts from Opsgenie that were created in the previous month and saves the data in an Excel file.

## Requirements

- Python 3.x
- `requests` library
- `openpyxl` library

To install the necessary libraries, run the following command:

```bash
pip install requests openpyxl
```

## How to Use

1. Clone or download this repository to your local machine.

2. Replace `"API_Token"` in the script with your actual Opsgenie API token.

```python
token = "API_Token"
```

3. Run the script. It will automatically fetch all the alerts created in the previous month and save them to an Excel file.

```bash
python opsgenie_alerts_to_excel.py
```

4. The script will save the data as an Excel file in the same directory with the following naming format: `alertas_opsgenie_<year>_<month>.xlsx`.

### Example Output Filename:
```
alertas_opsgenie_2024_11.xlsx
```

## Functionality

- **Fetch Alerts**: The script fetches all alerts from Opsgenie that were created in the previous month.
- **Save to Excel**: The script processes the data and saves it to an Excel file, flattening any nested dictionary structures into a single level of columns.
- **Pagination**: If there are more than 100 alerts, the script will handle pagination and fetch additional pages of alerts.

## License

This project is licensed under the MIT License.