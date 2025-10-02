# office
Contain office related scripts for automation

----------------------------------------------------------------------------------------------
git init #*
git status
git add FILE
git commit -m "TITLE"
git remote add origin URL #*
git push -u origin master
git pull origin master
git log

git config --global user.name "NAME"
git config --global user.email "EMAIL"

git checkout -b NEWBRANCH   #create branch
git checkout BRANCH
git merge NEWBRANCH
git branch -d NEWBRANCH   #delete branch



python -m venv my_env              #create venv
my_env\Scripts\activate            #activate venv
pip install numpy pandas            #install package
deactivate                                       #deactivate venv
rd /s /q my_env                              #remove venv
pip freeze > requirements.txt    #create requirements.txt
pip install -r requirements.txt   #install independency
----------------------------------------------------------------------------------------------

## SECTORAL ANALYSIS (Decide which market, sector, or stock universe you’re covering)
DATA FROM:
Financial statements: Income statement, balance sheet, cash flow (from company filings).
Indonesia → IDX website for annual/quarterly reports.
Global → SEC EDGAR (10-K, 10-Q), company websites.
Sector reports: Bloomberg, Reuters, Fitch, S&P, local banks’ research.
Economic indicators: Inflation, interest rates, GDP growth affecting sectors.

## MACROECONOMIC ANALYSIS (Check economic indicators, interest rates, inflation, and sector trends)
DATA FROM:
Besides Bank Indonesia (BI), you can rely on multiple sources for macro & sector data:
Statistics Indonesia (BPS – Badan Pusat Statistik) → GDP, inflation, unemployment, trade data.
Ministry of Finance (Kemenkeu) → Budget, fiscal policy, debt, government spending.
IDX (Indonesia Stock Exchange) → Sectoral indices, trading volumes, listed company data.
Financial news portals → Kontan, Bisnis Indonesia, Bloomberg Indonesia, Reuters, for real-time macro/sector trends.
International sources → IMF, World Bank, OECD, for global context impacting Indonesia.

## QUANTITATIVE ANALYSIS (Use quantitative filters: P/E, P/B, debt ratios, ROE, growth rates)
DATA FROM:
IDX (Bursa Efek Indonesia) → Annual & quarterly financial statements, sector data, market cap.
Company websites → Investor relations section for annual & quarterly reports.
Financial portals → RTI Business, Investing.com Indonesia, Kontan, Bloomberg Indonesia.
SEC EDGAR → 10-K, 10-Q filings (for US-listed companies).
Financial data providers → Yahoo Finance, Bloomberg, Reuters, Morningstar.

## QUALITATIVE ANALYSIS (Management quality, competitive moat, product pipeline, news)
DATA FROM:
Investor Relations websites → Annual reports, quarterly reports, presentations.
CEO/CFO letters → Insights on strategy, risk, and growth plans.
Business portals → Kontan, Bisnis Indonesia, Bloomberg Indonesia, Reuters.
Press releases → Product launches, partnerships, regulatory updates.
Local brokers → Danareksa, Mandiri Sekuritas, or regional equity research.
Global → Morningstar, Bloomberg, S&P, Fitch (for competitive analysis).
Trade associations, government reports, market studies → trends, pipelines, regulations.

## VALUATION ANALYSIS (Apply models DCF, relative valuation, or multiples)
DATA FROM:
Indonesia: IDX (Bursa Efek Indonesia), company websites (Laporan Tahunan / quarterly reports)
Global: SEC EDGAR (10-K, 10-Q), company IR websites
Indonesia: IDX, RTI Business, Investing.com Indonesia
Global: Yahoo Finance, Bloomberg, Reuters, Morningstar
Broker research, sell-side reports → provide comparable multiples, growth assumptions, risk rates

## SUMMARY (Executive summary -> macro → sector → stock analysis → valuation → recommendation)