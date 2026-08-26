# Beginner Claude AI Agent

A Jupyter notebook that walks through building an Anthropic **Managed Agent** powered by Claude Sonnet 5 — from creating the agent and cloud sandbox to streaming its work and generating a final report.

The agent loads a CSV of company parental-leave policies, uses Anthropic's XLSX skill, writes and runs a Python analysis script, and produces a JSON summary and an Excel report.

## What it demonstrates

- Loading the Anthropic API key from a local `.env` file
- Creating a Claude Sonnet 5 managed agent with the `agent_toolset_20260401` tools and the Anthropic-managed XLSX skill
- Configuring a cloud sandbox environment (`environments.create`)
- Uploading a CSV dataset via the Files API and mounting it into the sandbox
- Creating a session that links the agent, environment, and mounted file
- Streaming `agent.message` / `agent.tool_use` events while the agent works
- Having the agent write and execute a Python analysis script inside the sandbox
- Generating a JSON summary and a formula-driven Excel workbook

## Files

| File | Purpose |
|---|---|
| [starter.ipynb](starter.ipynb) | Main tutorial notebook |
| [api.py](api.py) | Loads `ANTHROPIC_API_KEY` from `.env` |
| [parental_leave.csv](parental_leave.csv) | Dataset mounted into the sandbox (1,601 companies) |
| [MTA_Daily_Ridership.csv](MTA_Daily_Ridership.csv) | Extra sample dataset for experimentation |
| [outputs/](outputs/) | Generated files from a run; not committed to Git |

## Quickstart

```bash
git clone <this-repo-url>
cd Beginner-Claude-AI-Agent

python -m venv .venv
source .venv/bin/activate        # Windows: .venv\Scripts\activate

pip install anthropic python-dotenv jupyterlab

echo "ANTHROPIC_API_KEY=sk-ant-..." > .env

jupyter lab starter.ipynb
```

Requires Python 3.9+ and an Anthropic API key with Managed Agents beta access (`managed-agents-2026-04-01`).

## Generated outputs

Running the notebook end to end produces these files under `outputs/`:

| Output | Description |
|---|---|
| `analyze_data.py` | Agent-written and executed analysis script |
| `summary.json` | Aggregated leave statistics per industry |
| `parental_leave_report.xlsx` | Excel workbook built from the analysis |

Sample results from the included dataset (1,601 companies):

| Metric | Result |
|---|---|
| Average paid maternity leave | 10.91 weeks |
| Average unpaid maternity leave | 6.63 weeks |
| Average paid paternity leave | 7.33 weeks |
| Average unpaid paternity leave | 7.73 weeks |

## Notes

> [!WARNING]
> Running the notebook creates billable Anthropic resources (agent, environment, session). Clean up these resources after each run to avoid ongoing charges.

- Managed Agents uses the `managed-agents-2026-04-01` beta surface.
- Never commit your `.env` file or API keys — `.env` is already excluded via `.gitignore`.
- Generated artifacts are saved under `/mnt/session/outputs` inside the sandbox before being downloaded locally.

## License

MIT
