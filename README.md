# Mutual Fund Portfolio Simulator

A Streamlit-based mutual fund planning and portfolio simulation tool for comparing SIP, investment cash flow, and SWP strategies in a single dashboard.

## Features

- Compare multiple analysis modes:
  - `SIP Only Mode`
  - `SIP + Investment CF Mode`
  - `SIP + SWP Mode`
  - `All Combination Mode`
- Switch between:
  - `Assumed Returns Basis`
  - `Actual NAV Basis`
- Build scenarios with:
  - SIP start date and monthly SIP
  - investment tenure
  - expected return and inflation
  - optional step-up
  - investment cash flows
  - SWP settings
- View:
  - simulation summary cards
  - scenario comparison table and chart
  - portfolio growth chart
  - cash flow schedule
- Export reports:
  - Excel report
  - PDF report with client details

## Project Structure

```text
main.py
portfolio_simulator/
  formatting.py
  models.py
  reporting.py
  simulation.py
assets/
  hl_logo.png
Funds Historical NAV/
  hdfc midcap fund NAV since 2007.xlsx
.streamlit/
  config.toml
requirements.txt
Procfile
```

## Requirements

- Python 3.11 or newer
- pip

Python packages are listed in [requirements.txt](requirements.txt).

## Local Setup

Install dependencies:

```bash
pip install -r requirements.txt
```

Run the app:

```bash
streamlit run main.py
```

If needed, you can also run:

```bash
python main.py
```

The app is configured to launch through Streamlit automatically.

## Data Files

The app expects these files to be present:

- [assets/hl_logo.png](assets/hl_logo.png)
- [Funds Historical NAV/hdfc midcap fund NAV since 2007.xlsx](Funds%20Historical%20NAV/hdfc%20midcap%20fund%20NAV%20since%202007.xlsx)

Do not remove or rename them unless you also update the paths in [main.py](main.py).

## Deployment

### Streamlit Community Cloud

1. Push this project to a GitHub repository.
2. Sign in to Streamlit Community Cloud.
3. Create a new app.
4. Select your GitHub repo and branch.
5. Set the main file path to `main.py`.
6. Deploy.

After deployment, future pushes to the deployed branch can trigger updates automatically.

### Procfile

This project also includes a [Procfile](Procfile):

```text
web: streamlit run main.py --server.port $PORT --server.address 0.0.0.0
```

## Notes

- Theme and UI behavior are configured through `.streamlit/config.toml` and custom CSS inside [main.py](main.py).
- Exported reports include client details entered during the export flow.
- Actual NAV analysis depends on the bundled historical NAV Excel file.

## License

This project currently does not define a license. Add one if you plan to share or publish it publicly.
