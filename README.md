# SOFR Monte Carlo Risk Simulator

> An Excel VBA-based risk simulation tool built for SOFR and U.S. Short-Term Interest Rate (STIR) futures traders — stress-testing trade PnL sequences across thousands of randomized scenarios to quantify risk of ruin, drawdown exposure, and expected return distributions.

---

## Why This Exists

In SOFR and Federal Funds Futures trading, a good monthly PnL number can hide a dangerous question: **were you skilled, or were you lucky with the order your wins arrived in?**

A trader who made $20,000 this month might have gone broke halfway through if their two largest losing trades had landed on Day 1 instead of Day 25. This tool answers that question systematically — by simulating 10,000+ randomized orderings of your actual trade history and measuring what would have happened to your account across all of them.

This is the same probabilistic thinking that underlies **Value at Risk (VaR)** frameworks used by institutional rates desks — applied to a STIR futures trading book through an accessible Excel interface.

---

## What It Does

You provide your historical trade PnL data (a column of numbers — each row is one trade's result). The tool then:

1. Takes your trade list and **randomly shuffles the order** thousands of times
2. Runs your starting account balance through each shuffled sequence
3. Tracks what happens — does the account survive? How far does it dip? Where does it end up?
4. Outputs a full distribution of outcomes across all simulations

This gives you a **probability distribution of outcomes** rather than a single historical number — so you can make position sizing and risk decisions based on what's *likely*, not just what *happened*.

---

## Key Output Metrics

| Metric | What It Tells You |
|---|---|
| **Risk of Ruin** | Probability your account hits zero across all simulated sequences |
| **Median Profit** | Expected PnL in a typical scenario |
| **Median Drawdown** | How far your account typically drops before recovering |
| **Equity Curves** | Visual range of possible account growth paths |
| **Scenario Breakdown** | Results segmented by starting equity and lot size |

---

## Relevance to SOFR / STIR Trading

This tool was adapted with a STIR futures context in mind. Specific use cases:

- **Position sizing validation** — test whether your current lot sizing survives volatile FOMC meeting windows across randomized trade sequences
- **Drawdown expectation setting** — before scaling volume (e.g. from 15K to 36K lots/month), simulate whether your account can absorb the increased sequence risk
- **Strategy robustness check** — a strategy that looks profitable historically may have high ruin probability under adverse sequencing. This surfaces that risk before it costs real capital
- **Basis trade analysis** — paste in PnL from SOFR vs. Fed Funds spread trades separately to compare risk profiles across instruments

---

## File Structure

```
sofr-montecarlo-risk-simulator/
│
├── monte_carlo_trade_simulator.xlsm   ← Main file — download and run this
│
├── clsSimulation.cls                  ← Core simulation engine (randomization + iteration logic)
├── clsEquityCurve.cls                 ← Tracks account balance through each simulated sequence
├── clsResult.cls                      ← Stores and computes output metrics (ruin %, drawdown, profit)
├── mdRun.bas                          ← Entry point — triggers when you press "Run"
├── mdFactory.bas                      ← Builds simulation objects and wires components together
├── INameProvider.cls                  ← Interface for consistent naming across modules
│
├── TestModule_All.bas                 ← Runs all unit tests (regression check after any code change)
├── TestModule_clsSimulation.bas       ← Unit tests for simulation logic
├── TestModule_clsEquityCurve.bas      ← Unit tests for equity curve tracking
├── TestModule_clsResult.bas           ← Unit tests for result computation
│
└── screenshots/                       ← Dashboard preview images
```

> The `.cls` and `.bas` files are the VBA source code exported for GitHub version control. They are already embedded inside the `.xlsm`. **You only need to download the `.xlsm` to use the tool.**

---

## Getting Started

![Screenshot](/screenshot.PNG)

### Step 1 — Download
Download `monte_carlo_trade_simulator.xlsm` from this repository.

### Step 2 — Unblock the file (Windows)
Because the file was downloaded from the internet, Windows blocks macros by default.

- Right-click the `.xlsm` file → **Properties**
- At the bottom, check **Unblock** → click **Apply** → **OK**

### Step 3 — Enable Macros in Excel
Open the file. When the yellow security bar appears at the top, click **Enable Content**.

If the bar doesn't appear, go to: `File → Options → Trust Center → Trust Center Settings → Macro Settings → Enable all macros`

### Step 4 — Input Your Trade Data
- Go to the **InputData** sheet
- Paste your trade PnL values — one trade per row (positive = win, negative = loss)
- For SOFR futures: input per-trade PnL in USD (e.g. +1250, -875, +3400...)

### Step 5 — Configure the Simulation
On the **Control** sheet, set:

| Parameter | Description |
|---|---|
| Starting Equity | Your account size in USD |
| Lot Size | Number of contracts per trade |
| Trades Per Year | Your average annual trade frequency |
| Number of Runs | How many simulations to run (1,000–10,000 recommended) |
| Margin | Margin per contract |

### Step 6 — Run
Press the **Run** button. Results populate automatically across the output sheets.

---

## Example Scenario

A SOFR futures trader with the following parameters:

- Starting equity: **$100,000**
- 60 trades over 3 months (1M SOFR, 3M SOFR, Fed Funds mix)
- Average win: **+$1,200** | Average loss: **-$850**
- Lot size: 10 contracts

Running 10,000 simulations might reveal:

- Risk of Ruin: **2.3%** — acceptable at current sizing
- Median Drawdown: **$14,500** (14.5% of capital)
- Median End Equity: **$118,400**
- Worst 5% scenario: account drops below **$72,000** at some point

This tells the trader: *at 10 lots, the strategy is viable. At 20 lots, re-run the simulation before scaling.*

---

## Tech Stack

- **Microsoft Excel** (.xlsm)
- **VBA** — object-oriented design using classes (`clsSimulation`, `clsEquityCurve`, `clsResult`)
- **No external dependencies** — runs entirely within Excel

---

## Background & Context

Monte Carlo simulation is a core technique in institutional fixed income risk management. It underpins:

- **VaR (Value at Risk)** calculations on rates books
- **CVA (Credit Valuation Adjustment)** on SOFR-linked swaps
- **Stress testing** for FOMC event risk and CPI surprise scenarios

This tool applies the same probabilistic logic at the individual trader level — making institutional-grade risk thinking accessible through Excel, without requiring a Bloomberg Terminal or Python environment.

---

## Limitations

- Past trade PnL is used as-is — the tool does not model autocorrelation between trades (i.e. it assumes each trade is independent)
- Does not account for changing market regimes (e.g. pre-FOMC vs. post-FOMC volatility differences)
- Margin calls and broker-specific liquidation rules are not modeled
- For live risk management, this should complement — not replace — formal VaR systems

---

## Author

**Punit Jain**
Market Analyst — Fixed Income & U.S. Rates | SOFR & Federal Funds Futures
[LinkedIn](https://linkedin.com/in/punit-jain-353197283) • [GitHub](https://github.com/punitjain18)

## Features

- Evaluate the performance of your trading strategies using Monte Carlo simulation
- Calculate the risk of ruin, median profit, median drawdown, and other key metrics for different starting equity amounts
- Analyze the impact of different lot sizes and number of trades per year
- Generate detailed equity curve data for further analysis

## Installation

To use the Excel VBA Monte Carlo Trade Simulator, simply download and open the Excel file `monte_carlo_trade_simulator.xlsm`. The other files in this repository are the VBA code files that already exist in the tool, but have been exported for versioning purposes on GitHub.

## Usage

1. Prepare a list of trade PNL (profit and loss) data for your trading strategy.
2. Paste the trades into the sheet labeled "InputData". 
3. On the "Control" sheet, set the required parameters for the simulation, such as lot size, number of trades per year, total number of runs, starting equity, and margin.
3. Press the button "Run" on the same sheet to run the Monte Carlo simulation and obtain the simulation results.
4. Analyze the results to evaluate the performance of your trading strategy.

## Contributing

Contributions are welcome, please feel free to open an issue or submit a pull request.

The module 'Test_Module_All" contains a subroutine called "RunAllTests" which can be used to test for any breaking changes after an amendment.

## Acknowledgements

This project is an original implementation of an idea proposed by Kevin Davey.
