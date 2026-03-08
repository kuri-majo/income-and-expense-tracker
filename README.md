# Project Overview

This is a project for an income and expenses tracker. It uses MS Excel on Windows for data entry and displays custom visualizations (i.e., visualizations that are not available in MS Excel) of the money stream in MS Excel itself.

Integration of Python and Excel is handled via the `xlwings` package.

This was a nice opportunity to think about a reasonable Python setup on Windows.

## Why integrating Excel and Python?

I had two constraints for my income and expense tracker:
1) Use of MS Excel for low-effort data entry and storage
2) Visualization of money flow also in MS Excel to avoid tool switching

I wanted to visualize my income and expenses with a Sankey plot, which is not available in MS Excel. I also did not want to pay extra money for an MS Excel add-In which would provide the Sankey plot functionality.

I therefore decided to look for the possibility of integrating Python into MS Excel.

## Why xlwings instead of the official "Python in Excel" or xlwings lite?

My MS Excel version neither supports the official "Python in Excel" functionality nor `xlwings lite`.
I therefore decided on the `xlwings` Python package and MS Excel add-in.
This setup requires that I install Python on Windows (because my MS Excel runs on a Windows machine).

# Getting Started

## Windows Python Setup

I use `uv` for managing Python versions, virtual environments, and Python packages in my project.

Here's how to install `uv` using PowerShell:
```
powershell -ExecutionPolicy ByPass -c "irm https://astral.sh/uv/install.ps1 | iex"
```

Then, use `uv` to install Python (I used Python 3.12)
```
uv python install 3.12
```

Clone the repository, cd into the repository and run the following to install your virtual environment.
```
uv venv  # creates a virtual environment in the workspace of the repository
uv sync  # installs the required packages specified in the uv.lock file
```

Activate the venv using the following command in the PowerShell terminal:
```
.venv\Scripts\activate
```

If you get a permission error here, set another PowerShell Execution Policy for your current terminal session:
```
Set-ExecutionPolicy Unrestricted -Scope Process
```

## Set up pre-commit hooks

Install pre-commit globally with uv:
```
uv tool install pre-commit
```

Install the pre-commit hooks:
```
pre-commit install --install-hooks
```

To run the hooks on all files, run:
```
pre-commit run -a
```

## File setup

Of course, my actual data is not committed to the repository. I have, however, included a template with the necessary data structure (yes, it's in German).

The Excel file that contains the data needs to have the same name as the Python module that contains the visualization code. Therefore, to try out the example, rename the file `income_and_expense_tracker_template.xlsm` to `income_and_expense_tracker.xlsm`.
