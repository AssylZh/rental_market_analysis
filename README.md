# U.S. Rental Property Market Analysis

**Where is it cheaper to rent than to buy, what drives home prices, and where would a rental property make money?**

Tools: **SQL** (SQLite) · **Python** (pandas, NumPy, SciPy, matplotlib) · **Excel**
Data: **real Zillow Research data** for the 100 largest U.S. metros, September 2015 – September 2026

## Key findings

1. **Home prices outran rents.** Since 2015, typical U.S. home values rose **88%** while rents rose **64%**.
2. **Buying vs. renting varies widely.** The median price-to-rent ratio is **17.2**. Buying is most expensive
   relative to rent in **San Jose (33.3)**, Salt Lake City and San Francisco, and cheapest in **Jackson, MS (12.0)**,
   El Paso and Pittsburgh. **26 of 100** metros are above 20 (expensive to buy).
3. **Rent growth and affordability explain home-price growth.** A regression across 100 metros explains
   **72%** of the variation in 5-year home-value growth (R² = 0.72):
   - each extra 1 point of rent growth goes with **+0.85 points** of home-value growth (p < 0.001)
   - each extra 1 point of starting price-to-rent ratio goes with **−0.61 points** (p = 0.002): already-expensive
     markets grew more slowly
   - Western metros grew least (median **+10.9%** over 5 years vs. **+31.8%** in the Northeast)
4. **Rentals don't cash-flow at today's rates.** At a **7.18%** mortgage rate (Freddie Mac, Nov 2026) with 20% down,
   **0 of 100** metros produce positive monthly cash flow. The best, Jackson, MS, loses **$332/month**; the median
   loses **$1,239**. Jackson would break even at a **4.08%** rate; only **3 metros** cash-flow positive even at 3%.

![National trend](charts/1_national_trend.png)
![Price-to-rent](charts/2_price_to_rent.png)
![Regression](charts/3_rent_vs_home_growth.png)

## Files

| File | What it does |
|---|---|
| `analysis.ipynb` | **Start here.** SQL queries, regression, charts and cash-flow analysis, with outputs |
| `rental_market_analysis.xlsx` | Excel model: metro data, region summary with charts, cash-flow model for all 100 metros, 12-month forecast for any metro (dropdown). 2,000+ live formulas; change the mortgage rate and everything updates |
| `make_excel.py` | Builds the Excel file from the notebook's output |
| `data/` | Zillow ZHVI (home values) and ZORI (rents) metro files, plus `metro_summary.csv` |

## Method

1. **Reshape** Zillow's wide files (one column per month) into long tables and load them into SQLite.
2. **SQL:** join home values and rents by metro and month; pull values now, 1 year ago and 5 years ago;
   compute price-to-rent ratios and growth rates.
3. **Python regression (OLS):** 5-year home-value growth on 5-year rent growth, starting price-to-rent ratio
   and Census region.
4. **Cash-flow model:** mortgage payment (PMT), property tax, insurance, maintenance, vacancy and management
   for a typical home in each metro; break-even mortgage rate; 12-month forecast in Excel.

## Limitations

- ZHVI and ZORI are typical values for different housing mixes (rents include apartments).
- One national set of cost assumptions; taxes and insurance vary a lot by state.
- The regression shows association, not cause, on 100 metros.
- Cash flow ignores appreciation and tax benefits, which are a large part of real investor returns.

## Run it

```bash
pip install pandas numpy scipy matplotlib openpyxl jupyter
jupyter notebook analysis.ipynb     # run all cells
python make_excel.py
```
