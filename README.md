# RSI Backtesting Script

This script performs a backtest of a trading strategy based on the Relative Strength Index (RSI) indicator. It iterates through different parameter combinations to find the most profitable ones.

## How it works

The script reads historical stock data from an Excel file, calculates the RSI for different parameter settings, and generates buy and sell signals based on the RSI values. It then simulates the trades and calculates the performance of the strategy for each parameter combination.

The results are saved to an Excel file, with different sheets for different win rate ranges.

## Performance

This script uses the `multiprocessing` module to run the backtests for different parameter combinations in parallel. This can significantly speed up the execution time on multi-core machines.

## How to use

1.  **Install the dependencies:**

    ```bash
    pip install pandas openpyxl
    ```

2.  **Prepare your data:**

    *   Create an Excel file with your historical stock data. The file should have at least the following columns: `Date`, `Open`, `AdjClose`.
    *   Update the `data_fl_path` and `df_data_file` variables in the `main` function to point to your data file.
    *   Create an output Excel file with the following sheets: `Data_0`, `Data_1`, `Data_2`, `Data_3`, `Data_4`, `Data_5`, `Data_6`. Update the `stk_summ_fl_nm_2` variable to point to this file.

3.  **Run the script:**

    ```bash
    python backtesting.py
    ```

## Parameters

The script iterates through the following parameters to find the optimal combination:

*   `holddays`: The maximum number of days to hold the asset.
*   `MAwin`: The window for the moving average calculation.
*   `spn`: The span for the exponential moving average calculation.
*   `rmn`: The minimum RSI value for a buy signal.
*   `rmx`: The maximum RSI value for a sell signal.

The ranges for these parameters can be adjusted in the `main` function.
