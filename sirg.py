#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# =======================================================================================================
# Created By: Haythm Alshehab - haythm@alshehab.org
# Version: 2.1
# Last Modified : Wed, May 3rd, 2023 at 4:59 PM
# Classification: Restricted
# =======================================================================================================
"""

This script allows the user to convert a Trello board with custom fields into an Excel report.

This tool accepts comma separated value files (.csv) as well as excel
(.xls, .xlsx) files and generates xlsx files.
"""
# =======================================================================================================
# Imports
# =======================================================================================================
# import numpy as np
import pandas as pd
import datetime as datetime
import plotly.graph_objects as go

# import plotly.express as px
import os

# import plotly.io as plt
from pretty_html_table import build_table
from termcolor import colored

# TODO: fix it!
pd.options.mode.chained_assignment = None  # default='warn'

# =======================================================================================================
# Global variables
# =======================================================================================================
COLORS = ["#E7C65B", "#225560", "#310D20", "#96031A"]
DEBUG = False


# ------------------------------------------------------------------------------
def about_script():
    """Asks the user to confirm his acknowledgment to the NDA before running the script."""

    print(
        colored(
            r"""
--------------------------------------------------------------------------------------
|   _____  _____        _____   ______  _____    ____   _____  _______        _____  ______  _   _  |
|  / ____||_   _|      |  __ \ |  ____||  __ \  / __ \ |  __ \|__   __|      / ____||  ____|| \ | | |
| | (___    | | ______ | |__) || |__   | |__) || |  | || |__) |  | | ______ | |  __ | |__   |  \| | |
|  \___ \   | ||______||  _  / |  __|  |  ___/ | |  | ||  _  /   | ||______|| | |_ ||  __|  | . ` | |
|  ____) | _| |_       | | \ \ | |____ | |     | |__| || | \ \   | |        | |__| || |____ | |\  | |
| |_____/ |_____|      |_|  \_\|______||_|      \____/ |_|  \_\  |_|         \_____||______||_| \_| |
|                                                                                                   |
--------------------------------------------------------------------------------------
    """,
            "red",
        )
    )
    print("---------------------")
    print("VERSION: 2.0 BETA" "\nAUTHOR: Haythm Alshehab - haythm@alshehab.org")
    print(
        "ABOUT: This script allows the user to clean and process tickets exported from Trello and "
        "generate charts and reports."
    )
    print("---------------------")


# ------------------------------------------------------------------------------
# TODO: Change this to PATH
def prepare_output():
    if not os.path.exists("OUTPUT/"):
        os.makedirs("OUTPUT/")

# ------------------------------------------------------------------------------
def initialisation():
    """Guides the user to use the correct naming for both files."""
    print(
        "INSTRUCTION: Before you continue, make sure you have the exported Trello board csv file (trello_board.csv) in "
        "the same directory as the "
        "script as in the following tree:"
    )
    print(
        r"""
    +--CURRENT_DIR/
    |  +--trello_board.csv
    |  +--sirg.py
        """
    )
    input("Press Enter to continue.")
    print("---------------------")
    prepare_output()  # ------------------------------------------------------------------------------


# ------------------------------------------------------------------------------
def load_trello_board():
    """Load the raw log file exported from LogRhythm.
    :return: raw logfile converted to pandas dataframe.
    """
    print(colored("[SUCCESS]", "green"), end=".....................")
    print("Loaded the Trello board csv file.")
    # log = Path("11_09_2019-CE_ALL_W44.csv")
    # if not trello_board.exists():
    #     print("The trello_board file doesn't exist!")
    # else:
    #     print("The trello_board file exists!")
    # Load dataset
    trello_board = pd.read_csv("./INPUT/j8wC07hR - sip-soc-shared.csv")

    trello_board.rename(
        columns={
            "Card Name": "T#",
            "Card Description": "DESC",
            "List Name": "STATUS",
            "CREATION_DATE": "TICKET_CREATION_TIMESTAMP",
            "Card ID": "TICKET_RESPONSE_TIMESTAMP",
            "RESOLUTION_DATE": "TICKET_RESOLUTION_TIMESTAMP",
            "CATEGORY": "CATEGORY",
            "LOG_SOURCE": "LOG_SOURCE",
            "PRIORITY": "PRIORITY",
            "OFFENSE_ID": "OFFENSE_ID",
            "RESOLUTION_CODE": "RESOLUTION_CODE",
        },
        inplace=True,
    )

    # Remove new line char since it is used EVERYWHERE
    trello_board["DESC"] = trello_board["DESC"].str.replace("\n", "")
    return trello_board


# ------------------------------------------------------------------------------
def calulate_default_start_and_end_dates():
    today = pd.to_datetime("today").normalize()
    idx = (today.weekday() + 4) % 7
    default_report_end_timestamp = (
        today - datetime.timedelta(idx) + pd.Timedelta(hours=16, seconds=-1)
    )
    default_report_start_timestamp = (
        default_report_end_timestamp
        - datetime.timedelta(weeks=1)
        + pd.Timedelta(seconds=1)
    )

    return default_report_start_timestamp, default_report_end_timestamp


# ------------------------------------------------------------------------------
def specify_report_time_range():
    (
        default_report_start_timestamp,
        default_report_end_timestamp,
    ) = calulate_default_start_and_end_dates()
    start_timestamp = input(
        "USER INPUT: Plese press Enter if you want to keep the default start date and time of the report ({}), "
        "otherwise, enter the desired start date and time of the report in the following format: yyyy-mm-dd "
        "hh:mm\n".format(colored(default_report_start_timestamp, "green"))
    )
    if not start_timestamp:
        start_timestamp = default_report_start_timestamp
    print("---------------------")
    end_timestamp = input(
        "USER INPUT: Plese press Enter if you want to keep the default end date and time of the report ({}), otherwise, enter the desired end date and time of the report in the following format: yyyy-mm-dd hh:mm\n".format(
            colored(default_report_end_timestamp, "green")
        )
    )
    if not end_timestamp:
        end_timestamp = default_report_end_timestamp
    print("---------------------")

    return start_timestamp, end_timestamp


# ------------------------------------------------------------------------------
def process_timestamps(trello_board, start_timestamp, end_timestamp):
    # Convert timestamps datetime objects
    trello_board["TICKET_CREATION_TIMESTAMP"] = pd.to_datetime(
        trello_board["TICKET_CREATION_TIMESTAMP"]
    )
    trello_board["TICKET_RESOLUTION_TIMESTAMP"] = pd.to_datetime(
        trello_board["TICKET_RESOLUTION_TIMESTAMP"]
    )

    # Generate start and end date of the trello table
    trello_board.start_timestamp = (
        trello_board["TICKET_CREATION_TIMESTAMP"].min().strftime("%d/%m/%Y|%H:%M:%S")
    )
    trello_board.end_timestamp = (
        trello_board["TICKET_CREATION_TIMESTAMP"].max().strftime("%d/%m/%Y|%H:%M:%S")
    )
    trello_board.no_days = len(
        trello_board["TICKET_CREATION_TIMESTAMP"].dt.normalize().unique()
    )
    trello_board = trello_board[
        (trello_board["TICKET_CREATION_TIMESTAMP"] >= start_timestamp)
    ]
    trello_board = trello_board[
        (trello_board["TICKET_CREATION_TIMESTAMP"] < end_timestamp)
    ]

    return trello_board


# ------------------------------------------------------------------------------
def filter_tickets(trello_board):
    # Only analyze resolved tickets
    trello_board = trello_board[trello_board["STATUS"] == "RESOLVED_AND_REVIEWED"]
    trello_board = trello_board[trello_board["CATEGORY"] == "VSOC_INVESTIGATION"]
    trello_board = trello_board.sort_values(by=["T#"], ignore_index=True)
    print(colored("[SUCCESS]", "green"), end=".....................")
    print("Filtered out non-resolved and non-security investigation cards.")

    if DEBUG:
        trello_board[trello_board.isna().any(axis=1)]
    return trello_board


# ------------------------------------------------------------------------------
# Deprecated
# def get_card_creation_date(trello_board):
#     trello_board['TICKET_RESPONSE_TIMESTAMP'] = [x[:8] for x in trello_board['TICKET_RESPONSE_TIMESTAMP']]
#     trello_board['TICKET_RESPONSE_TIMESTAMP'] = trello_board['TICKET_RESPONSE_TIMESTAMP'].apply(int, base=16)
#     trello_board['TICKET_RESPONSE_TIMESTAMP'] = pd.to_datetime(trello_board['TICKET_RESPONSE_TIMESTAMP'],unit='s')
#     trello_board['TICKET_RESPONSE_TIMESTAMP'] = trello_board['TICKET_RESPONSE_TIMESTAMP'].dt.tz_localize('GMT').dt.tz_convert('Asia/Riyadh').dt.tz_localize(None)

#     trello_board['TICKET_CREATION_TIMESTAMP_ADJUSTED'] = trello_board['TICKET_CREATION_TIMESTAMP']
#     trello_board['TICKET_RESPONSE_TIMESTAMP_ADJUSTED'] = trello_board['TICKET_RESPONSE_TIMESTAMP']
#     trello_board['TICKET_RESOLUTION_TIMESTAMP_ADJUSTED'] = trello_board['TICKET_RESOLUTION_TIMESTAMP']
#     if DEBUG:
#         display(trello_board[trello_board['TICKET_CREATION_TIMESTAMP'].isnull()])
#         display(trello_board[trello_board['TICKET_RESOLUTION_TIMESTAMP'].isnull()])    # Check if we have error in timestamps
#         trello_board['ERROR'] = (trello_board['TICKET_CREATION_TIMESTAMP'
#                                 ].dt.second
#                                 >= trello_board['TICKET_RESOLUTION_TIMESTAMP'
#                                 ].dt.second).astype(int)

#     display(trello_board[trello_board.duplicated('T#')])
#     trello_board.drop_duplicates(subset ="T#",keep = False, inplace = True)
#     return trello_board

# ------------------------------------------------------------------------------
# TODO: Fix it
# def offset_business_hours(trello_board):
#     sip_bh = pd.offsets.CustomBusinessHour(
#         start="08:15", end="15:30", weekmask="Sun Mon Tue Wed Thu"
#     )
#     trello_board["TICKET_RESPONSE_TIMESTAMP"] = [
#         x[:8] for x in trello_board["TICKET_RESPONSE_TIMESTAMP"]
#     ]
#     trello_board["TICKET_RESPONSE_TIMESTAMP"] = trello_board[
#         "TICKET_RESPONSE_TIMESTAMP"
#     ].apply(int, base=16)
#     trello_board["TICKET_RESPONSE_TIMESTAMP"] = pd.to_datetime(
#         trello_board["TICKET_RESPONSE_TIMESTAMP"], unit="s"
#     )
#     trello_board["TICKET_RESPONSE_TIMESTAMP"] = (
#         trello_board["TICKET_RESPONSE_TIMESTAMP"]
#         .dt.tz_localize("GMT")
#         .dt.tz_convert("Asia/Riyadh")
#         .dt.tz_localize(None)
#     )

#     trello_board["TICKET_CREATION_TIMESTAMP_ADJUSTED"] = trello_board[
#         "TICKET_CREATION_TIMESTAMP"
#     ]
#     trello_board["TICKET_RESPONSE_TIMESTAMP_ADJUSTED"] = trello_board[
#         "TICKET_RESPONSE_TIMESTAMP"
#     ]
#     trello_board["TICKET_RESOLUTION_TIMESTAMP_ADJUSTED"] = trello_board[
#         "TICKET_RESOLUTION_TIMESTAMP"
#     ]

#     for index, row in trello_board.iterrows():
#         roll1 = sip_bh.rollforward(
#             pd.Timestamp(trello_board["TICKET_CREATION_TIMESTAMP"][index])
#         )
#         trello_board["TICKET_CREATION_TIMESTAMP_ADJUSTED"][index] = roll1
#         roll2 = sip_bh.rollforward(
#             pd.Timestamp(trello_board["TICKET_RESPONSE_TIMESTAMP"][index])
#         )
#         trello_board["TICKET_RESPONSE_TIMESTAMP_ADJUSTED"][index] = roll2
#         roll3 = sip_bh.rollforward(
#             pd.Timestamp(trello_board["TICKET_RESOLUTION_TIMESTAMP"][index])
#         )
#         trello_board["TICKET_RESOLUTION_TIMESTAMP_ADJUSTED"][index] = roll3

#     trello_board["BUSINESS_HOURS_TO_ACKNOWLEDGE"] = trello_board.apply(
#         lambda x: len(
#             pd.date_range(
#                 start=x.TICKET_CREATION_TIMESTAMP,
#                 end=x.TICKET_RESPONSE_TIMESTAMP,
#                 freq=sip_bh,
#             )
#         ),
#         axis=1,
#     )
#     trello_board["BUSINESS_HOURS_TO_RESOLVE"] = trello_board.apply(
#         lambda x: len(
#             pd.date_range(
#                 start=x.TICKET_CREATION_TIMESTAMP,
#                 end=x.TICKET_RESOLUTION_TIMESTAMP,
#                 freq=sip_bh,
#             )
#         ),
#         axis=1,
#     )

#     # TICKETS/BUSINESS HOURS ratio
#     tickets_business_hours_ratio = len(trello_board) / len(
#         pd.bdate_range(
#             start=trello_board["TICKET_CREATION_TIMESTAMP"].min(),
#             end=trello_board["TICKET_CREATION_TIMESTAMP"].max(),
#             freq=sip_bh,
#         )
#     )

#     # print(
#     #     "Rate of resolved VSOC ticket per business hour = ",
#     #     tickets_business_hours_ratio,
#     #     "tickets per business hour",
#     # )
#     print(colored('[SUCCESS]', 'green'), end='.....................')
#     print('Offseted business hours.')
#     return trello_board


# ------------------------------------------------------------------------------
# Deprecated
# def extract_labels(trello_board):
#     # Handle labels
#     labels_dict = {'SC' : 'SEC_CONTROL', 'RC' : 'RESOLUTION_CODE', 'CO' : 'DEPENDENT_ON', 'TB' : 'ROOTICKET_CAUSE',
#     'PR' : 'PRIORITY'}
#     trello_board["LABELS"] = trello_board["LABELS"].str.replace('\([^)]*\)', '', regex=True)
#     trello_board["LABELS"] = trello_board["LABELS"].str.replace(',', ' ', regex=True)
#     #  Temp. solution :)
#     trello_board['SEC_CONTROL'] = trello_board['LABELS'].str.extract(r'(\bSC\d{2}\b)')
#     trello_board['RESOLUTION_CODE'] = trello_board['LABELS'].str.extract(r'(\bRC\d{2}\b)')
#     trello_board['DEPENDENT_ON'] = trello_board['LABELS'].str.extract(r'(\bCO\d{2}\b)')
#     trello_board['ROOT_CAUSE'] = trello_board['LABELS'].str.extract(r'(\bTB\d{2}\b)')
#     trello_board['PRIORITY'] = trello_board['LABELS'].str.extract(r'(\bPR\d{2}\b)')
#     # Load labels translation as a separate external file to keep the confidentiality of your log sources
#     labels_translation = pd.read_csv('./INPUT/labels_translation.csv', index_col=0, header=None, squeeze=True).
#     to_dict()
#     # print(labels_translation)
#     trello_board = trello_board.replace(labels_translation)
#     trello_board.drop('LABELS', axis=1, inplace=True)
#     return trello_board

# ------------------------------------------------------------------------------
def gen_barplot(trello_board):
    required_fields = ["LOG_SOURCE", "RESOLUTION_CODE"]

    trello_board.start_timestamp = (
        trello_board["TICKET_CREATION_TIMESTAMP"].min().strftime("%d/%m/%Y|%H:%M:%S")
    )
    trello_board.end_timestamp = (
        trello_board["TICKET_CREATION_TIMESTAMP"].max().strftime("%d/%m/%Y|%H:%M:%S")
    )
    trello_board.no_days = len(
        trello_board["TICKET_CREATION_TIMESTAMP"].dt.normalize().unique()
    )

    for required_field in required_fields:
        copy_of_trello_board = trello_board.copy()
        series = pd.value_counts(copy_of_trello_board[required_field])
        mask = (series / series.sum() * 100).lt(1.0)
        required_field_count = copy_of_trello_board[required_field].value_counts()
        required_field_count = required_field_count.rename_axis(
            required_field
        ).reset_index(name="{}_COUNT".format(required_field))
        required_field_count = required_field_count.reset_index(drop=True)
        required_field_count.index.rename("NO.", inplace=True)
        required_field_count.index += 1
        required_field_count["{}_PCT".format(required_field)] = (
            required_field_count["{}_COUNT".format(required_field)]
            / required_field_count["{}_COUNT".format(required_field)].sum()
        )
        report_start_date_abbreviated = (
            trello_board["TICKET_CREATION_TIMESTAMP"].min().strftime("%d%b%y")
        ).upper()
        report_end_date_abbreviated = (
            trello_board["TICKET_CREATION_TIMESTAMP"].max().strftime("%d%b%y")
        ).upper()
        required_field_count.to_csv(
            "./OUTPUT/[{}-{}]{}.csv".format(
                report_start_date_abbreviated,
                report_end_date_abbreviated,
                required_field,
            ),
            sep=",",
        )
        fig = go.Figure(
            data=[
                go.Pie(
                    labels=required_field_count[required_field],
                    values=required_field_count["{}_PCT".format(required_field)],
                    textinfo="label+percent",
                    insidetextorientation="radial",
                    showlegend=False,
                    marker=dict(colors=COLORS, line=dict(color="#000000", width=2)),
                )
            ]
        )
        fig.update_layout(
            title="Security Investigation Tickets distributed by<br> {}".format(
                required_field
            )
        )
        # fig.layout.images = [dict( source=logo_path, xref='paper', yref='paper', x=0.97,  y=0.97, sizex=0.50,
        # sizey=0.50, xanchor='center', yanchor='bottom')]
        report_start_date_abbreviated = (
            trello_board["TICKET_CREATION_TIMESTAMP"].min().strftime("%d%b%y")
        ).upper()
        report_end_date_abbreviated = (
            trello_board["TICKET_CREATION_TIMESTAMP"].max().strftime("%d%b%y")
        ).upper()
        fig.write_image(
            "./OUTPUT/[{}-{}]{}_COUNT.svg".format(
                report_start_date_abbreviated,
                report_end_date_abbreviated,
                required_field,
            )
        )
        print(colored("[SUCCESS]", "green"), end=".....................")
        print("Barplot chart generated and exported.")


# ------------------------------------------------------------------------------
def gen_summary_table(trello_board):
    # Generate a summary table of the main features
    duration_start = (
        trello_board["TICKET_CREATION_TIMESTAMP"].min().strftime("%Y-%m-%d")
    )
    # print(duration_start)
    summary_table_count = trello_board.groupby(["RESOLUTION_CODE"]).size()
    report_start_date_abbreviated = (
        trello_board["TICKET_CREATION_TIMESTAMP"].min().strftime("%d%b%y")
    ).upper()
    report_end_date_abbreviated = (
        trello_board["TICKET_CREATION_TIMESTAMP"].max().strftime("%d%b%y")
    ).upper()
    summary_table_count.to_excel(
        "./OUTPUT/[{}-{}]SUMMARY_TABLE_COUNT.xlsx".format(
            report_start_date_abbreviated, report_end_date_abbreviated, duration_start
        )
    )
    print(colored("[SUCCESS]", "green"), end=".....................")
    print("Sumamry table generated and exported.")


# ------------------------------------------------------------------------------
def gen_trendline(trello_board):
    trello_board.start_timestamp = (
        trello_board["TICKET_CREATION_TIMESTAMP"].min().strftime("%d/%m/%Y|%H:%M:%S")
    )
    trello_board.end_timestamp = (
        trello_board["TICKET_CREATION_TIMESTAMP"].max().strftime("%d/%m/%Y|%H:%M:%S")
    )
    trello_board.no_days = len(
        trello_board["TICKET_CREATION_TIMESTAMP"].dt.normalize().unique()
    )

    tickets_count = trello_board["T#"].value_counts()
    tickets_count = tickets_count.rename_axis("T#").reset_index(name="T#_COUNT")
    tickets_count = tickets_count.reset_index(drop=True)
    tickets_count.index.rename("NO.", inplace=True)
    tickets_count.index += 1
    tickets_count["T#_PCT"] = (
        tickets_count["T#_COUNT"] / tickets_count["T#_COUNT"].sum()
    )
    s = pd.to_datetime(trello_board["TICKET_CREATION_TIMESTAMP"])
    tickets_count = s.groupby(s.dt.floor("d")).size().reset_index(name="COUNT")

    # Plot ---------------------------------------------------------------------------------

    my_layout = go.Layout(font=dict(color="#7f7f7f"))
    my_data = go.Scatter(
        x=tickets_count["TICKET_CREATION_TIMESTAMP"],
        y=tickets_count["COUNT"],
        mode="lines+markers+text",
    )
    fig = go.Figure(data=my_data)

    fig.update_layout(
        shapes=[
            go.layout.Shape(  # Line Horizontal
                type="line",
                x0=tickets_count["TICKET_CREATION_TIMESTAMP"].min(),
                y0=tickets_count["COUNT"].mean(),
                x1=tickets_count["TICKET_CREATION_TIMESTAMP"].max(),
                y1=tickets_count["COUNT"].mean(),
                line=dict(color="black", width=1, dash="longdash"),
            )
        ]
    )

    fig.add_trace(
        go.Scatter(
            x=[
                tickets_count["TICKET_CREATION_TIMESTAMP"].max()
                - pd.Timedelta(days=trello_board.no_days) / 2.0
            ],
            y=[tickets_count["COUNT"].mean()],
            mode="markers+text",
            name="Markers and Text",
            hoverinfo="skip",
            textposition="top right",
        )
    )

    fig.update_traces(
        marker_color="rgb(231,198,91)",
        marker_line_color="black",
        marker_line_width=1,
        opacity=1.0,
    )
    fig.update_layout(
        title_text="VSOC tickets trendline grouped by the day<br>Dashed line represents the average no. VSOC tickets "
                   "({})".format(
            int(tickets_count["COUNT"].mean())
        )
    )

    fig.update_layout(showlegend=False)
    # fig.layout.images = [dict(
    #     source=logo_path,
    #     x=0.9,
    #     y=1.05,
    #     sizex=0.25,
    #     sizey=0.25,
    #     xanchor='center',
    #     yanchor='bottom',
    #     )]
    report_start_date_abbreviated = (
        trello_board["TICKET_CREATION_TIMESTAMP"].min().strftime("%d%b%y")
    ).upper()
    report_end_date_abbreviated = (
        trello_board["TICKET_CREATION_TIMESTAMP"].max().strftime("%d%b%y")
    ).upper()
    fig.write_image(
        "./OUTPUT/[{}-{}]TRENDLINE.svg".format(
            report_start_date_abbreviated, report_end_date_abbreviated
        )
    )
    print(colored("[SUCCESS]", "green"), end=".....................")
    print("Trendline chart generated and exported.")


# ------------------------------------------------------------------------------
def gen_customer_report(trello_board):
    # Save a customer report
    customer_report = trello_board[
        [
            "T#",
            "TICKET_CREATION_TIMESTAMP",
            "OFFENSE_ID",
            "LOG_SOURCE",
            "RESOLUTION_CODE",
            "DESC",
            "PRIORITY",
        ]
    ]
    customer_report.index.rename("NO.", inplace=True)
    customer_report = customer_report.sort_values(by=["T#"], ignore_index=True)
    customer_report.index.rename("NO.", inplace=True)
    customer_report.index += 1
    file_name = "EXTERNAL_REPORT.xlsx"
    report_start_date_abbreviated = (
        trello_board["TICKET_CREATION_TIMESTAMP"].min().strftime("%d%b%y")
    ).upper()
    report_end_date_abbreviated = (
        trello_board["TICKET_CREATION_TIMESTAMP"].max().strftime("%d%b%y")
    ).upper()
    customer_report.to_excel(
        "./OUTPUT/[{}-{}]{}".format(
            report_start_date_abbreviated, report_end_date_abbreviated, file_name
        )
    )
    print(colored("[SUCCESS]", "green"), end=".....................")
    print("External report generated and exported.")


def gen_internal_report(trello_board):
    report_start_date_abbreviated = (
        trello_board["TICKET_CREATION_TIMESTAMP"].min().strftime("%d%b%y")
    ).upper()
    report_end_date_abbreviated = (
        trello_board["TICKET_CREATION_TIMESTAMP"].max().strftime("%d%b%y")
    ).upper()
    trello_board = trello_board.reset_index(drop=True)
    trello_board.index += 1
    file_name = "[{}-{}]SOC_REPORT.html".format(
        report_start_date_abbreviated, report_end_date_abbreviated
    )
    trello_board.to_html("./OUTPUT/{}".format(file_name))
    trello_board.to_excel(
        "./OUTPUT/[{}-{}]INTERNAL_REPORT.xlsx".format(
            report_start_date_abbreviated, report_end_date_abbreviated
        )
    )
    html_table_blue_light = build_table(trello_board, "grey_light", index=True)
    with open("./OUTPUT/{}".format(file_name), "w") as f:
        f.write(html_table_blue_light)
    print(colored("[SUCCESS]", "green"), end=".....................")
    print("Internal report generated and exported.")


if __name__ == "__main__":
    about_script()
    initialisation()
    start_timestamp, end_timestamp = specify_report_time_range()

    trello_board = load_trello_board()
    trello_board = filter_tickets(trello_board)
    trello_board = process_timestamps(trello_board, start_timestamp, end_timestamp)
    # trello_board = offset_business_hours(trello_board)

    gen_internal_report(trello_board)
    gen_customer_report(trello_board)
    gen_summary_table(trello_board)
    gen_barplot(trello_board)
    gen_trendline(trello_board)
