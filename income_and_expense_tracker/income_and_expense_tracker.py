import itertools

import pandas as pd
import plotly.graph_objects as go
import xlwings as xw


def main():
    wb = xw.Book.caller()
    sheet = wb.sheets["Sheet1"]

    data = sheet.used_range.value
    df = pd.DataFrame(data[1:], columns=data[0])

    # put into long format
    # first, keep all transactions that surely come from a source
    source_transactions = df[[col for col in df.columns if col != "Kategorie"]]

    # then, identify transactions that went through an intermediate level and put into same format as source transactions
    target_transactions = df[df["Kategorie"].notnull()][[col for col in df.columns if col != "Quelle"]]
    target_transactions = target_transactions.rename(columns={"Ziel": "Quelle", "Kategorie": "Ziel"})

    # concatenate back together
    long_format_transactions = pd.concat([source_transactions, target_transactions]).sort_values("TransaktionsID")

    # aggregate data
    agg_df = long_format_transactions[["Quelle", "Ziel", "Wert"]].groupby(["Quelle", "Ziel"]).agg(sum).reset_index()

    # define labels, i.e., levels in the Sankey plot
    label_columns = ["Quelle", "Ziel"]
    label_links = long_format_transactions[label_columns].drop_duplicates()
    labels = list(set(itertools.chain.from_iterable([set(long_format_transactions[x]) for x in label_columns])))

    # get indices of labels
    agg_df["Quelle_idx"] = agg_df["Quelle"].apply(lambda x: labels.index(x))
    agg_df["Ziel_idx"] = agg_df["Ziel"].apply(lambda x: labels.index(x))

    fig = go.Figure(
        data=[
            go.Sankey(
                node=dict(
                    # pad = 15,
                    # thickness = 20,
                    # line = dict(color = "black", width = 0.5),
                    label=labels,
                    # color = "blue"
                ),
                link=dict(
                    source=agg_df[
                        "Quelle_idx"
                    ],  # [0, 1, 0, 2, 3, 3], # indices correspond to labels, eg A1, A2, A1, B1, ...
                    target=agg_df["Ziel_idx"],  # [2, 3, 3, 4, 4, 5],
                    value=agg_df["Wert"],  # [8, 4, 2, 8, 4, 2]
                ),
            )
        ]
    )

    fig.update_layout(title_text="Basic Sankey Diagram", font_size=10)

    sheet.pictures.add(fig, name="Basic Sankey Diagram", update=True)


if __name__ == "__main__":
    xw.Book("income_and_expense_tracker\\income_and_expense_tracker.xlsm").set_mock_caller()
    main()
