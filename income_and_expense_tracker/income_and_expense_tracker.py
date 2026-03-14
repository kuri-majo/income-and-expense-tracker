import itertools

import pandas as pd
import plotly.graph_objects as go
import xlwings as xw
from loguru import logger


def _get_data(wb):
    transactions_sheet = wb.sheets["Transaktionen"]

    data = transactions_sheet.used_range.value
    df = pd.DataFrame(data[1:], columns=data[0])

    return df


def _clean_data(df):
    # cast integer columns to integers
    int_columns = ["TransaktionsID", "Monat", "Jahr"]
    df[int_columns] = df[int_columns].apply(lambda x: x.astype("Int32"))

    # remove transactions with no value (i.e., placeholders)
    df = df.dropna(subset="Wert")

    return df


def _prepare_yearly_data(df, year):
    # keep only transactions for the requested year
    df = df[df["Jahr"] == year]

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

    return agg_df


def _create_sankey_plot(agg_df, year):
    # define labels, i.e., levels in the Sankey plot
    label_columns = ["Quelle", "Ziel"]
    labels = list(set(itertools.chain.from_iterable([set(agg_df[x]) for x in label_columns])))

    # get indices of labels
    agg_df["Quelle_idx"] = agg_df["Quelle"].apply(lambda x: labels.index(x))
    agg_df["Ziel_idx"] = agg_df["Ziel"].apply(lambda x: labels.index(x))

    fig = go.Figure(
        layout={"width": 1600, "height": 900, "font_size": 20, "title": f"Money Flow {year}"},
        data=[
            go.Sankey(
                node=dict(
                    align="justify",
                    pad=50,  # padding (in px) between the nodes when align="justify"
                    thickness=20,  # the thickness (in px) of the nodes
                    # groups parameter  # we can define groups of nodes
                    label=labels,
                ),
                link=dict(
                    source=agg_df["Quelle_idx"],
                    target=agg_df["Ziel_idx"],
                    value=agg_df["Wert"],
                ),
            )
        ],
    )

    return fig


def _add_sheet_if_not_exists(wb, sheet_name: str):
    try:
        sheet = wb.sheets(sheet_name)
    except Exception:
        logger.warning(f"Requested sheet does not seem to exist")
        logger.info(f"Adding requested sheet called '{sheet_name}'")
        sheet = wb.sheets.add(sheet_name)
    return sheet


def main():
    wb = xw.Book.caller()

    df = _get_data(wb)
    df = _clean_data(df)

    for year in df["Jahr"].unique():
        agg_df = _prepare_yearly_data(df, year)
        fig = _create_sankey_plot(agg_df, year)

        # add figure to the Plots sheet for the year
        plots_sheet = _add_sheet_if_not_exists(wb, f"Plots {year}")
        plots_sheet.pictures.add(
            fig,
            name=f"Money Flow {year}",
            update=True,
            format="svg",
        )

        logger.info(f"Added summary plot for year {year}")

    logger.info("Done")


if __name__ == "__main__":
    xw.Book("income_and_expense_tracker\\income_and_expense_tracker.xlsm").set_mock_caller()
    main()
