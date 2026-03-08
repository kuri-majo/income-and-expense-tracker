import xlwings as xw
from xlwings import func
import plotly.express as px
import pandas as pd



def main():
    wb = xw.Book.caller()
    sheet = wb.sheets["Sheet1"]

    data = sheet.used_range.value
    df = pd.DataFrame(data[1:], columns=data[0])
    # if sheet["A1"].value == "Hello xlwings!":
    #     sheet["A1"].value = "Bye xlwings!"
    # else:
    #     sheet["A1"].value = "Hello xlwings!"

    # Plotly chart
    df = px.data.iris()
    fig = px.scatter(df, x="sepal_width", y="sepal_length", color="species")

    sheet.pictures.add(fig, name='IrisScatterPlot', update=True)


@func
def hello(name):
    return f"Hello {name}!"


if __name__ == "__main__":
    xw.Book("income_and_expense_tracker.xlsm").set_mock_caller()
    main()
