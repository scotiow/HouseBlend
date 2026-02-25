import base64
import io
import os
import tempfile
from typing import Any

import dash
from dash import Dash, Input, Output, State, dash_table, dcc, html
import pandas as pd

from HouseBlend.HouseBlend import HansardRepository, HouseBlendSession


app = Dash(__name__, title="HouseBlend Scheduler")
server = app.server

app.layout = html.Div(
    [
        html.H2("HouseBlend Scheduler"),
        html.P("Upload a Hansard workbook, run optimisation, and download outputs."),
        dcc.Upload(
            id="hansard-upload",
            children=html.Div(["Drag and drop or ", html.A("select Hansard file")]),
            style={
                "width": "100%",
                "height": "90px",
                "lineHeight": "90px",
                "borderWidth": "1px",
                "borderStyle": "dashed",
                "borderRadius": "8px",
                "textAlign": "center",
                "marginBottom": "16px",
            },
            multiple=False,
        ),
        html.Div(id="upload-status", children="No file uploaded yet.", style={"marginBottom": "12px"}),
        html.Div(
            [
                html.Label("Periods to schedule"),
                dcc.Input(id="n-to-schedule", type="number", min=1, step=1, value=1),
                html.Label("Current period", style={"marginLeft": "16px"}),
                dcc.Input(id="current-period", type="number", min=1, step=1, value=1),
                html.Label("Mailmerge period", style={"marginLeft": "16px"}),
                dcc.Input(id="mailmerge-period", type="number", min=1, step=1, value=1),
            ],
            style={"display": "flex", "alignItems": "center", "gap": "8px", "flexWrap": "wrap"},
        ),
        html.Div(
            [
                html.Button("Run optimisation", id="run-btn", n_clicks=0),
                html.Button("Download updated Hansard", id="download-hansard-btn", n_clicks=0, style={"marginLeft": "8px"}),
                html.Button("Download mailmerge", id="download-mailmerge-btn", n_clicks=0, style={"marginLeft": "8px"}),
            ],
            style={"marginTop": "16px", "marginBottom": "16px"},
        ),
        html.Div(id="status", children="Waiting for input."),
        dcc.Store(id="result-store"),
        dcc.Download(id="download-hansard"),
        dcc.Download(id="download-mailmerge"),
        dcc.Tabs(
            [
                dcc.Tab(
                    label="Current Period Matchups",
                    children=[
                        dash_table.DataTable(
                            id="current-period-table",
                            data=[],
                            columns=[],
                            page_size=20,
                            style_table={"overflowX": "auto"},
                        )
                    ],
                ),
                dcc.Tab(
                    label="Full Schedule",
                    children=[
                        dash_table.DataTable(
                            id="full-schedule-table",
                            data=[],
                            columns=[],
                            page_size=20,
                            style_table={"overflowX": "auto"},
                        )
                    ],
                ),
            ],
            style={"marginTop": "16px"},
        ),
        html.Div(
            [
                html.Strong("Privacy statement"),
                html.P(
                    "Files are processed at your own risk. Uploaded data is handled only for the current run "
                    "and is not intentionally stored server-side after processing completes."
                ),
            ],
            style={
                "padding": "12px",
                "borderRadius": "8px",
                "backgroundColor": "#fff6e6",
                "border": "1px solid #efd9b5",
                "marginTop": "16px",
            },
        ),
    ],
    style={
        "maxWidth": "960px",
        "margin": "0 auto",
        "padding": "24px",
        "fontFamily": "'IBM Plex Sans', 'Avenir Next', sans-serif",
        "background": "linear-gradient(135deg, #f5f8ff 0%, #eef6f2 100%)",
        "minHeight": "100vh",
    },
)


def _decode_upload(contents: str) -> bytes:
    _, content_string = contents.split(",", 1)
    return base64.b64decode(content_string)


def _encode_bytes(content: bytes) -> str:
    return base64.b64encode(content).decode("utf-8")


def _mailmerge_to_bytes(participants_df: pd.DataFrame, assistants_df: pd.DataFrame) -> bytes:
    stream = io.BytesIO()
    with pd.ExcelWriter(stream, engine="openpyxl") as writer:
        participants_df.to_excel(writer, sheet_name="Participants", index=False)
        assistants_df.to_excel(writer, sheet_name="Assistants", index=False)
    stream.seek(0)
    return stream.read()


def _detect_current_period(upload_bytes: bytes) -> int:
    schedule = pd.read_excel(io.BytesIO(upload_bytes), sheet_name="Schedule", index_col=0)
    period_numbers: list[int] = []
    for column in schedule.columns:
        if not str(column).startswith("Period "):
            continue
        try:
            period_numbers.append(int(str(column).replace("Period ", "").strip()))
        except ValueError:
            continue

    if not period_numbers:
        return 1

    period_numbers.sort()
    last_filled = 0
    for number in period_numbers:
        column = f"Period {number}"
        values = schedule[column]
        filled_mask = values.notna() & values.astype(str).str.strip().ne("")
        if filled_mask.any():
            last_filled = number
    return last_filled + 1 if last_filled > 0 else 1


@app.callback(
    Output("upload-status", "children"),
    Output("current-period", "value"),
    Output("mailmerge-period", "value"),
    Input("hansard-upload", "contents"),
    State("hansard-upload", "filename"),
    prevent_initial_call=True,
)
def on_upload(upload_contents: str | None, upload_filename: str | None):
    if not upload_contents:
        return "No file uploaded yet.", 1, 1

    try:
        file_bytes = _decode_upload(upload_contents)
        inferred_current_period = _detect_current_period(file_bytes)
        name = upload_filename or "hansard.xlsx"
        return (
            f"Uploaded: {name}. Detected current period: {inferred_current_period}.",
            inferred_current_period,
            inferred_current_period,
        )
    except Exception as exc:
        return f"Upload received but could not read Schedule sheet: {exc}", 1, 1


@app.callback(
    Output("status", "children"),
    Output("result-store", "data"),
    Input("run-btn", "n_clicks"),
    State("hansard-upload", "contents"),
    State("hansard-upload", "filename"),
    State("n-to-schedule", "value"),
    State("current-period", "value"),
    State("mailmerge-period", "value"),
    prevent_initial_call=True,
)
def run_houseblend(
    n_clicks: int,
    upload_contents: str | None,
    upload_filename: str | None,
    n_to_schedule: Any,
    current_period: Any,
    mailmerge_period: Any,
):
    del n_clicks

    if not upload_contents:
        return "Please upload a Hansard .xlsx file.", None

    try:
        schedule_periods = int(n_to_schedule)
        current = int(current_period)
        mailmerge = int(mailmerge_period)
    except (TypeError, ValueError):
        return "Periods must be valid integers.", None

    if schedule_periods < 1 or current < 1 or mailmerge < 1:
        return "Periods must be >= 1.", None

    try:
        raw_file = _decode_upload(upload_contents)
        input_name = upload_filename or "hansard.xlsx"
        if not input_name.lower().endswith(".xlsx"):
            input_name = f"{input_name}.xlsx"
        input_name = os.path.basename(input_name)
        base_name, _ = os.path.splitext(input_name)

        with tempfile.TemporaryDirectory() as temp_dir:
            input_path = os.path.join(temp_dir, input_name)
            with open(input_path, "wb") as f:
                f.write(raw_file)

            session = HouseBlendSession.from_excel(
                folderpath=temp_dir,
                filename=input_name,
                parliament_name="uploaded",
            )
            session.update_participants().optimise(
                n_to_schedule=schedule_periods,
                current_period=current,
                save=False,
            )
            session.build_schedule(save=False)

            repo = HansardRepository()
            updated_hansard_name = input_name
            repo.save(
                session.contacts,
                session.dates,
                session.availability,
                session.schedule,
                parliament_name="uploaded",
                folderpath=temp_dir,
                filename=updated_hansard_name,
            )

            updated_hansard_path = os.path.join(temp_dir, updated_hansard_name)
            with open(updated_hansard_path, "rb") as f:
                updated_hansard_bytes = f.read()

            participants_df, assistants_df = session.export_mailmerge(period=mailmerge, save=False)
            mailmerge_bytes = _mailmerge_to_bytes(participants_df, assistants_df)
            mailmerge_name = f"{base_name}_period_{mailmerge}_mailmerge.xlsx"

            current_period_df = session.schedule_builder.period_meeting_list(
                session.contacts,
                session.bool_schedule,
                current,
            )
            full_schedule_df = session.schedule.copy().reset_index().rename(columns={"index": "Person"})

        payload = {
            "hansard_name": updated_hansard_name,
            "hansard_data": _encode_bytes(updated_hansard_bytes),
            "mailmerge_name": mailmerge_name,
            "mailmerge_data": _encode_bytes(mailmerge_bytes),
            "current_period_rows": current_period_df.to_dict("records"),
            "current_period_columns": [{"name": c, "id": c} for c in current_period_df.columns],
            "full_schedule_rows": full_schedule_df.to_dict("records"),
            "full_schedule_columns": [{"name": c, "id": c} for c in full_schedule_df.columns],
        }
        return "Optimisation complete. Use the download buttons.", payload
    except Exception as exc:
        return f"Error: {exc}", None


@app.callback(
    Output("current-period-table", "data"),
    Output("current-period-table", "columns"),
    Output("full-schedule-table", "data"),
    Output("full-schedule-table", "columns"),
    Input("result-store", "data"),
)
def update_results_tables(result_data: dict[str, Any] | None):
    if not result_data:
        return [], [], [], []
    return (
        result_data.get("current_period_rows", []),
        result_data.get("current_period_columns", []),
        result_data.get("full_schedule_rows", []),
        result_data.get("full_schedule_columns", []),
    )


@app.callback(
    Output("download-hansard", "data"),
    Input("download-hansard-btn", "n_clicks"),
    State("result-store", "data"),
    prevent_initial_call=True,
)
def download_hansard(n_clicks: int, result_data: dict[str, str] | None):
    del n_clicks
    if not result_data:
        return dash.no_update

    content = base64.b64decode(result_data["hansard_data"])
    return dcc.send_bytes(lambda b: b.write(content), filename=result_data["hansard_name"])


@app.callback(
    Output("download-mailmerge", "data"),
    Input("download-mailmerge-btn", "n_clicks"),
    State("result-store", "data"),
    prevent_initial_call=True,
)
def download_mailmerge(n_clicks: int, result_data: dict[str, str] | None):
    del n_clicks
    if not result_data:
        return dash.no_update

    content = base64.b64decode(result_data["mailmerge_data"])
    return dcc.send_bytes(lambda b: b.write(content), filename=result_data["mailmerge_name"])


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", "8050")), debug=False)
