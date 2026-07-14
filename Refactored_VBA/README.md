# VBA Refactoring

We have refactored the VBA code for the Portfolio Optimizer to improve modularity, reusability, and dashboard configuration.

## Key Changes
1. **Config.cls**: A new class to hold all constant parameters (Sheet names, target ratios, risk-free rows, half-life values). Currently defaults are set in `Class_Initialize()`, but they can easily be linked to read from cells on the Dashboard sheet in the future.
2. **PortfolioData.cls**: A skeleton class started for holding loaded and aligned arrays (returns, prices) to avoid having to re-fetch and align them repeatedly across different modules (like in `ComputeRiskFactors`).
3. **Module Refactoring**: Many hardcoded sheet names and parameter constants inside `allocationLogic.bas`, `Main.bas`, `solverUtils.bas`, `portfolioUtils.bas`, and `chartsUtils.bas` have been replaced to use the `config` object instead.

## How to Import back into Excel
1. Open the Developer Tab in Excel.
2. Go to Visual Basic Editor (VBE).
3. Remove the old versions of these `.bas` files.
4. Go to File -> Import File... and import the `.bas` and `.cls` files from this `Refactored_VBA` directory.
5. Save the workbook.

### Dashboard Automation
We have also included a new module `DashboardBuilder.bas`. This module contains subroutines that programmatically generate and format the Dashboard layout, including dynamically creating inputs, strategy tables, and settings blocks.
1. Import `DashboardBuilder.bas` into your VBE.
2. Run `BuildMainDashboard()` to instantly scaffold a clean layout and generate the required named ranges (like `InputTableStart`, `StrategyWeightsStart`, and `EqWeightingLimit`).
3. You can also use `AddCustomSection()` to dynamically add new tables to the Dashboard.
