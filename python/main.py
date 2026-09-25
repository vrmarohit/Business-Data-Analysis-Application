import streamlit as st
import pandas as pd
import numpy as np
from pathlib import Path
from datetime import datetime

from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment

# ============================================================
# PDF REPORT
# ============================================================

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER
from reportlab.lib.units import mm
from reportlab.platypus import (
    SimpleDocTemplate,
    Paragraph,
    Spacer,
    Table,
    TableStyle,
    PageBreak
)

# ============================================================
# PAGE SETUP
# ============================================================

st.set_page_config(
    page_title="Business Data Analysis Application",
    page_icon="📊",
    layout="wide"
)

# ============================================================
# PATHS
# ============================================================

BASE_DIR = Path(__file__).resolve().parent.parent
DATA_DIR = BASE_DIR / "data"
REPORT_DIR = BASE_DIR / "report"

DATA_DIR.mkdir(exist_ok=True)
REPORT_DIR.mkdir(exist_ok=True)

CLEANED_FILE = DATA_DIR / "cleaned_business_data.csv"
FORECAST_FILE = DATA_DIR / "business_forecast.csv"
INSIGHTS_FILE = DATA_DIR / "ai_business_insights.txt"

EXCEL_REPORT_FILE = REPORT_DIR / "Business_Analysis_Report.xlsx"
PDF_REPORT_FILE = REPORT_DIR / "Business_Analysis_Report.pdf"

# ============================================================
# COLUMN ALIASES
# ============================================================

COLUMN_ALIASES = {

    "Date": [
        "date",
        "order date",
        "transaction date",
        "sales date",
        "purchase date",
        "created date",
        "joining date",
        "hire date",
        "delivery date",
        "month"
    ],

    "ID": [
        "id",
        "order id",
        "order_id",
        "order number",
        "order no",
        "transaction id",
        "transaction_id",
        "customer id",
        "customer_id",
        "employee id",
        "employee_id",
        "invoice id",
        "invoice_id"
    ],

    "Revenue": [
        "sales",
        "sale",
        "revenue",
        "total sales",
        "total revenue",
        "amount",
        "total amount",
        "income",
        "turnover",
        "gmv",
        "net sales"
    ],

    "Cost": [
        "cost",
        "total cost",
        "expense",
        "expenses",
        "spending",
        "expenditure"
    ],

    "Profit": [
        "profit",
        "net profit",
        "profit amount",
        "gross profit",
        "net income",
        "earnings"
    ],

    "Quantity": [
        "quantity",
        "qty",
        "units",
        "units sold",
        "volume"
    ],

    "Product": [
        "product",
        "product name",
        "item",
        "item name",
        "service",
        "service name"
    ],

    "Category": [
        "category",
        "product category",
        "type",
        "segment",
        "business segment",
        "department",
        "division"
    ],

    "Region": [
        "region",
        "area",
        "zone",
        "location",
        "state",
        "city",
        "country",
        "territory"
    ],

    "Customer": [
        "customer",
        "customer name",
        "customer type",
        "customer category",
        "client",
        "client type"
    ],

    "Payment": [
        "payment",
        "payment method",
        "payment_method",
        "payment type",
        "payment mode"
    ],

    "Discount": [
        "discount",
        "discount percentage",
        "discount percent",
        "discount %",
        "discount rate"
    ]
}

# ============================================================
# HELPER FUNCTIONS
# ============================================================

def normalize_name(name):
    """Normalize column names for automatic detection."""

    name = str(name).strip().lower()

    for char in [
        "_", "-", "/", "\\",
        "(", ")", "%", ",", "."
    ]:
        name = name.replace(char, " ")

    return " ".join(name.split())


def detect_columns(df):
    """
    Automatically detects common business column meanings.
    Does NOT require all columns to exist.
    """

    existing = {
        normalize_name(column): column
        for column in df.columns
    }

    mapping = {}
    used_original_columns = set()

    for standard_name, aliases in COLUMN_ALIASES.items():

        for alias in aliases:

            normalized_alias = normalize_name(alias)

            if normalized_alias in existing:

                original = existing[normalized_alias]

                if original not in used_original_columns:

                    mapping[original] = standard_name
                    used_original_columns.add(original)

                break

    return mapping


def detect_data_types(df):
    """Detect dates, numeric and categorical columns."""

    date_columns = []
    numeric_columns = []
    categorical_columns = []

    for column in df.columns:

        series = df[column]

        if pd.api.types.is_numeric_dtype(series):

            numeric_columns.append(column)

            continue

        converted_numeric = pd.to_numeric(
            series,
            errors="coerce"
        )

        numeric_ratio = (
            converted_numeric.notna().mean()
            if len(series) > 0
            else 0
        )

        if numeric_ratio >= 0.80:

            numeric_columns.append(column)
            continue

        converted_date = pd.to_datetime(
            series,
            errors="coerce"
        )

        date_ratio = (
            converted_date.notna().mean()
            if len(series) > 0
            else 0
        )

        if date_ratio >= 0.80:

            date_columns.append(column)

        else:

            categorical_columns.append(column)

    return (
        date_columns,
        numeric_columns,
        categorical_columns
    )


def find_semantic_column(
    mapping,
    semantic_name
):
    """Find detected column by semantic name."""

    for original, detected in mapping.items():

        if detected == semantic_name:

            return detected

    return None


def get_column_by_semantic(
    df,
    semantic_name
):
    """
    Find a column using the standardized semantic names.
    """

    if semantic_name in df.columns:

        return semantic_name

    return None


def choose_primary_metric(
    df,
    mapping,
    numeric_columns
):
    """
    Select the most meaningful business metric.
    """

    priority = [
        "Revenue",
        "Sales",
        "Amount",
        "Income",
        "Profit",
        "Cost",
        "Quantity"
    ]

    for semantic in priority:

        if semantic in df.columns:

            if pd.api.types.is_numeric_dtype(
                df[semantic]
            ):

                return semantic

    # Fallback to first numeric column
    if numeric_columns:

        return numeric_columns[0]

    return None


def choose_secondary_metrics(
    df,
    primary_metric
):

    metrics = []

    for column in df.columns:

        if column == primary_metric:
            continue

        if pd.api.types.is_numeric_dtype(
            df[column]
        ):

            metrics.append(column)

    return metrics


# ============================================================
# DATA QUALITY
# ============================================================

def check_data_quality(df):

    quality = {}

    quality["total_rows"] = len(df)

    quality["total_columns"] = len(df.columns)

    quality["missing_values"] = int(
        df.isna().sum().sum()
    )

    quality["duplicate_rows"] = int(
        df.duplicated().sum()
    )

    invalid_dates = 0

    date_columns, _, _ = detect_data_types(df)

    for column in date_columns:

        converted = pd.to_datetime(
            df[column],
            errors="coerce"
        )

        invalid_dates += int(
            converted.isna().sum()
        )

    quality["invalid_dates"] = invalid_dates

    numeric_columns = [
        column
        for column in df.columns
        if pd.api.types.is_numeric_dtype(
            df[column]
        )
    ]

    negative_values = 0

    for column in numeric_columns:

        try:

            negative_values += int(
                (df[column] < 0).sum()
            )

        except Exception:

            pass

    quality["negative_numeric_values"] = (
        negative_values
    )

    total_issues = (
        quality["missing_values"]
        + quality["duplicate_rows"]
        + quality["invalid_dates"]
    )

    quality["total_issues"] = total_issues

    if total_issues == 0:

        quality["status"] = "Excellent"

    elif total_issues <= 10:

        quality["status"] = "Good"

    elif total_issues <= 50:

        quality["status"] = "Needs Attention"

    else:

        quality["status"] = "Poor"

    return quality


# ============================================================
# DATA CLEANING
# ============================================================

def clean_data(df):

    df = df.copy()

    original_rows = len(df)

    # Remove completely empty rows
    df = df.dropna(
        how="all"
    )

    # Remove duplicate rows
    duplicate_count = int(
        df.duplicated().sum()
    )

    df = df.drop_duplicates()

    # Detect types
    date_columns, numeric_columns, categorical_columns = (
        detect_data_types(df)
    )

    # Convert date columns
    for column in date_columns:

        df[column] = pd.to_datetime(
            df[column],
            errors="coerce"
        )

    # Convert numeric columns
    for column in numeric_columns:

        df[column] = pd.to_numeric(
            df[column],
            errors="coerce"
        )

        if df[column].notna().any():

            median_value = df[column].median()

            df[column] = df[column].fillna(
                median_value
            )

    # Fill categorical missing values
    for column in categorical_columns:

        df[column] = (
            df[column]
            .astype("object")
            .fillna("Unknown")
        )

    # Find primary metric
    mapping = detect_columns(df)

    # Rename recognized columns to standard names
    rename_mapping = {}

    for original, standard in mapping.items():

        if (
            original in df.columns
            and standard not in df.columns
        ):

            rename_mapping[original] = standard

    df = df.rename(
        columns=rename_mapping
    )

    # Remove rows where primary metric is missing
    primary_candidates = [
        "Revenue",
        "Sales",
        "Amount",
        "Profit",
        "Income"
    ]

    for column in primary_candidates:

        if column in df.columns:

            df = df.dropna(
                subset=[column]
            )

            break

    # Sort by date when available
    if "Date" in df.columns:

        df = df.sort_values(
            "Date"
        )

    df = df.reset_index(
        drop=True
    )

    return (
        df,
        original_rows,
        duplicate_count,
        len(df)
    )


# ============================================================
# FORECAST
# ============================================================

def create_forecast(
    df,
    date_column,
    metric_column
):

    if (
        date_column is None
        or metric_column is None
    ):

        return pd.DataFrame()

    temp = df[
        [date_column, metric_column]
    ].copy()

    temp = temp.dropna()

    if temp.empty:

        return pd.DataFrame()

    temp[date_column] = pd.to_datetime(
        temp[date_column],
        errors="coerce"
    )

    temp = temp.dropna(
        subset=[date_column]
    )

    if temp.empty:

        return pd.DataFrame()

    monthly = (
        temp
        .set_index(date_column)
        .resample("ME")[metric_column]
        .sum()
        .reset_index()
    )

    monthly.columns = [
        "Date",
        "Actual"
    ]

    if len(monthly) < 2:

        monthly["Forecast"] = np.nan

        return monthly

    x = np.arange(
        len(monthly)
    )

    y = monthly[
        "Actual"
    ].values

    try:

        slope, intercept = np.polyfit(
            x,
            y,
            1
        )

    except Exception:

        monthly["Forecast"] = np.nan

        return monthly

    monthly["Forecast"] = (
        slope * x + intercept
    )

    last_date = monthly[
        "Date"
    ].max()

    future_dates = pd.date_range(
        start=last_date
        + pd.offsets.MonthEnd(1),
        periods=6,
        freq="ME"
    )

    future_x = np.arange(
        len(monthly),
        len(monthly) + 6
    )

    future_values = (
        slope * future_x
        + intercept
    )

    future = pd.DataFrame({

        "Date": future_dates,

        "Actual": np.nan,

        "Forecast": future_values
    })

    return pd.concat(
        [
            monthly,
            future
        ],
        ignore_index=True
    )


# ============================================================
# AUTOMATIC BUSINESS INSIGHTS
# ============================================================

def generate_insights(
    df,
    primary_metric,
    numeric_columns,
    categorical_columns,
    date_columns
):

    insights = []
    recommendations = []

    if df.empty:

        return insights, recommendations

    # --------------------------------------------------------
    # PRIMARY METRIC
    # --------------------------------------------------------

    if primary_metric:

        values = pd.to_numeric(
            df[primary_metric],
            errors="coerce"
        ).dropna()

        if not values.empty:

            total = values.sum()

            average = values.mean()

            median = values.median()

            maximum = values.max()

            minimum = values.min()

            insights.append(
                f"The primary business metric "
                f"'{primary_metric}' has a total "
                f"value of {total:,.2f}."
            )

            insights.append(
                f"The average {primary_metric} "
                f"is {average:,.2f}, while the "
                f"median is {median:,.2f}."
            )

            insights.append(
                f"The highest recorded "
                f"{primary_metric} is "
                f"{maximum:,.2f}."
            )

            insights.append(
                f"The lowest recorded "
                f"{primary_metric} is "
                f"{minimum:,.2f}."
            )

    # --------------------------------------------------------
    # CATEGORICAL ANALYSIS
    # --------------------------------------------------------

    for column in categorical_columns[:5]:

        try:

            grouped = (
                df.groupby(column)[
                    primary_metric
                ]
                .sum()
                .sort_values(
                    ascending=False
                )
            )

            if len(grouped) > 0:

                top_category = grouped.index[0]

                insights.append(
                    f"'{top_category}' is the "
                    f"highest-performing "
                    f"{column} based on "
                    f"{primary_metric}."
                )

                recommendations.append(
                    f"Monitor and learn from "
                    f"the performance of "
                    f"'{top_category}' in "
                    f"{column}."
                )

        except Exception:

            pass

    # --------------------------------------------------------
    # NUMERIC CORRELATIONS
    # --------------------------------------------------------

    if len(numeric_columns) >= 2:

        try:

            correlation = (
                df[numeric_columns]
                .corr()
                .abs()
            )

            pairs = []

            for i in range(
                len(correlation.columns)
            ):

                for j in range(
                    i + 1,
                    len(correlation.columns)
                ):

                    value = correlation.iloc[
                        i, j
                    ]

                    if not pd.isna(value):

                        pairs.append(
                            (
                                value,
                                correlation.columns[i],
                                correlation.columns[j]
                            )
                        )

            if pairs:

                pairs.sort(
                    reverse=True
                )

                value, col_a, col_b = pairs[0]

                if value >= 0.70:

                    insights.append(
                        f"{col_a} and {col_b} "
                        f"show a strong statistical "
                        f"relationship "
                        f"(correlation ≈ {value:.2f})."
                    )

                    recommendations.append(
                        f"Investigate the relationship "
                        f"between {col_a} and {col_b} "
                        f"for business decision-making."
                    )

        except Exception:

            pass

    # --------------------------------------------------------
    # DATE TREND
    # --------------------------------------------------------

    if date_columns and primary_metric:

        date_column = date_columns[0]

        try:

            temp = df[
                [date_column, primary_metric]
            ].dropna()

            monthly = (
                temp
                .set_index(date_column)
                .resample("ME")[primary_metric]
                .sum()
            )

            if len(monthly) >= 2:

                first_value = monthly.iloc[0]

                last_value = monthly.iloc[-1]

                if first_value != 0:

                    change = (
                        (
                            last_value
                            - first_value
                        )
                        / abs(first_value)
                    ) * 100

                    direction = (
                        "increased"
                        if change >= 0
                        else "decreased"
                    )

                    insights.append(
                        f"The primary metric "
                        f"{direction} by approximately "
                        f"{abs(change):.1f}% between "
                        f"the first and last available "
                        f"months."
                    )

        except Exception:

            pass

    # --------------------------------------------------------
    # GENERAL RECOMMENDATIONS
    # --------------------------------------------------------

    recommendations.extend([

        "Monitor important business KPIs regularly.",

        "Investigate unusually high or low values.",

        "Compare performance across important "
        "business categories or segments.",

        "Use historical trends to support "
        "future planning.",

        "Improve data quality before making "
        "important business decisions."
    ])

    # Remove duplicate recommendations
    recommendations = list(
        dict.fromkeys(
            recommendations
        )
    )

    return (
        insights,
        recommendations
    )


# ============================================================
# MAIN ANALYSIS ENGINE
# ============================================================

def generate_analysis(df):

    results = {}

    # --------------------------------------------------------
    # DETECT DATA TYPES
    # --------------------------------------------------------

    date_columns, numeric_columns, categorical_columns = (
        detect_data_types(df)
    )

    results["date_columns"] = date_columns

    results["numeric_columns"] = numeric_columns

    results["categorical_columns"] = (
        categorical_columns
    )

    # --------------------------------------------------------
    # SEMANTIC COLUMN DETECTION
    # --------------------------------------------------------

    mapping = detect_columns(df)

    results["mapping"] = mapping

    # --------------------------------------------------------
    # PRIMARY METRIC
    # --------------------------------------------------------

    primary_metric = choose_primary_metric(
        df,
        mapping,
        numeric_columns
    )

    results["primary_metric"] = primary_metric

    # --------------------------------------------------------
    # KPI
    # --------------------------------------------------------

    results["total_records"] = len(df)

    results["total_columns"] = len(df.columns)

    if primary_metric:

        values = pd.to_numeric(
            df[primary_metric],
            errors="coerce"
        )

        results["metric_total"] = (
            values.sum()
        )

        results["metric_average"] = (
            values.mean()
        )

        results["metric_median"] = (
            values.median()
        )

        results["metric_std"] = (
            values.std()
        )

        results["metric_max"] = (
            values.max()
        )

        results["metric_min"] = (
            values.min()
        )

    else:

        results["metric_total"] = 0

        results["metric_average"] = 0

        results["metric_median"] = 0

        results["metric_std"] = 0

        results["metric_max"] = 0

        results["metric_min"] = 0

    # --------------------------------------------------------
    # REVENUE / COST / PROFIT
    # --------------------------------------------------------

    if (
        "Revenue" in df.columns
        and "Cost" in df.columns
        and "Profit" not in df.columns
    ):

        df["Profit"] = (
            pd.to_numeric(
                df["Revenue"],
                errors="coerce"
            )
            -
            pd.to_numeric(
                df["Cost"],
                errors="coerce"
            )
        )

        if "Profit" not in numeric_columns:

            numeric_columns.append(
                "Profit"
            )

    results["has_revenue"] = (
        "Revenue" in df.columns
    )

    results["has_cost"] = (
        "Cost" in df.columns
    )

    results["has_profit"] = (
        "Profit" in df.columns
    )

    # --------------------------------------------------------
    # PROFIT MARGIN
    # --------------------------------------------------------

    if (
        "Revenue" in df.columns
        and "Profit" in df.columns
    ):

        revenue = pd.to_numeric(
            df["Revenue"],
            errors="coerce"
        ).sum()

        profit = pd.to_numeric(
            df["Profit"],
            errors="coerce"
        ).sum()

        results["profit_margin"] = (
            profit / revenue
            if revenue != 0
            else 0
        )

    else:

        results["profit_margin"] = None

    # --------------------------------------------------------
    # CATEGORY ANALYSIS
    # --------------------------------------------------------

    category_tables = {}

    if primary_metric:

        for column in categorical_columns:

            if column not in df.columns:
                continue

            try:

                grouped = (
                    df.groupby(
                        column,
                        dropna=False
                    )[primary_metric]
                    .agg(
                        Total="sum",
                        Average="mean",
                        Count="count"
                    )
                    .reset_index()
                    .sort_values(
                        "Total",
                        ascending=False
                    )
                )

                category_tables[column] = grouped

            except Exception:

                pass

    results["category_tables"] = (
        category_tables
    )

    # --------------------------------------------------------
    # NUMERIC SUMMARY
    # --------------------------------------------------------

    if numeric_columns:

        numeric_summary = (
            df[numeric_columns]
            .describe()
            .T
            .reset_index()
            .rename(
                columns={
                    "index": "Column"
                }
            )
        )

    else:

        numeric_summary = pd.DataFrame()

    results["numeric_summary"] = (
        numeric_summary
    )

    # --------------------------------------------------------
    # CORRELATION
    # --------------------------------------------------------

    if len(numeric_columns) >= 2:

        correlation = (
            df[numeric_columns]
            .corr()
        )

    else:

        correlation = pd.DataFrame()

    results["correlation"] = correlation

    # --------------------------------------------------------
    # DATE ANALYSIS
    # --------------------------------------------------------

    monthly = pd.DataFrame()

    forecast = pd.DataFrame()

    if (
        date_columns
        and primary_metric
    ):

        date_column = date_columns[0]

        temp = df[
            [date_column, primary_metric]
        ].copy()

        temp[date_column] = pd.to_datetime(
            temp[date_column],
            errors="coerce"
        )

        temp = temp.dropna(
            subset=[
                date_column,
                primary_metric
            ]
        )

        if not temp.empty:

            monthly = (
                temp
                .set_index(date_column)
                .resample("ME")[primary_metric]
                .sum()
                .reset_index()
            )

            monthly.columns = [
                "Date",
                "Value"
            ]

            forecast = create_forecast(
                df,
                date_column,
                primary_metric
            )

    results["monthly"] = monthly

    results["forecast"] = forecast

    # --------------------------------------------------------
    # AI-ASSISTED INSIGHTS
    # --------------------------------------------------------

    insights, recommendations = (
        generate_insights(
            df,
            primary_metric,
            numeric_columns,
            categorical_columns,
            date_columns
        )
    )

    results["insights"] = insights

    results["recommendations"] = (
        recommendations
    )

    return results


# ============================================================
# SAVE OUTPUTS
# ============================================================

def save_outputs(
    results,
    cleaned_df
):

    cleaned_df.to_csv(
        CLEANED_FILE,
        index=False
    )

    if not results["forecast"].empty:

        results["forecast"].to_csv(
            FORECAST_FILE,
            index=False
        )

    with open(
        INSIGHTS_FILE,
        "w",
        encoding="utf-8"
    ) as file:

        file.write(
            "AI-ASSISTED BUSINESS INSIGHTS\n"
        )

        file.write(
            "=" * 50
            + "\n\n"
        )

        for number, insight in enumerate(
            results["insights"],
            1
        ):

            file.write(
                f"{number}. {insight}\n"
            )

        file.write(
            "\nBUSINESS RECOMMENDATIONS\n"
        )

        file.write(
            "=" * 50
            + "\n\n"
        )

        for number, recommendation in enumerate(
            results["recommendations"],
            1
        ):

            file.write(
                f"{number}. {recommendation}\n"
            )


# ============================================================
# EXCEL REPORT
# ============================================================

def generate_excel_report(
    results,
    cleaned_df
):

    with pd.ExcelWriter(
        EXCEL_REPORT_FILE,
        engine="openpyxl"
    ) as writer:

        # Executive Summary
        summary_rows = [

            [
                "Metric",
                "Value"
            ],

            [
                "Total Records",
                results["total_records"]
            ],

            [
                "Total Columns",
                results["total_columns"]
            ],

            [
                "Primary Metric",
                results["primary_metric"]
                or "Not detected"
            ],

            [
                "Primary Metric Total",
                results["metric_total"]
            ],

            [
                "Primary Metric Average",
                results["metric_average"]
            ],

            [
                "Primary Metric Median",
                results["metric_median"]
            ],

            [
                "Primary Metric Standard Deviation",
                results["metric_std"]
            ],

            [
                "Maximum Value",
                results["metric_max"]
            ],

            [
                "Minimum Value",
                results["metric_min"]
            ]
        ]

        if results["profit_margin"] is not None:

            summary_rows.append(
                [
                    "Profit Margin",
                    results["profit_margin"]
                ]
            )

        pd.DataFrame(
            summary_rows[1:],
            columns=summary_rows[0]
        ).to_excel(
            writer,
            sheet_name="Executive Summary",
            index=False
        )

        # Numeric summary
        if not results[
            "numeric_summary"
        ].empty:

            results[
                "numeric_summary"
            ].to_excel(
                writer,
                sheet_name="Numeric Analysis",
                index=False
            )

        # Correlation
        if not results[
            "correlation"
        ].empty:

            results[
                "correlation"
            ].to_excel(
                writer,
                sheet_name="Correlation",
                index=True
            )

        # Category tables
        for column, table in (
            results[
                "category_tables"
            ].items()
        ):

            sheet_name = (
                f"{column[:25]} Analysis"
            )

            table.to_excel(
                writer,
                sheet_name=sheet_name,
                index=False
            )

        # Monthly
        if not results[
            "monthly"
        ].empty:

            results[
                "monthly"
            ].to_excel(
                writer,
                sheet_name="Monthly Analysis",
                index=False
            )

        # Forecast
        if not results[
            "forecast"
        ].empty:

            results[
                "forecast"
            ].to_excel(
                writer,
                sheet_name="Forecast",
                index=False
            )

        # Insights
        pd.DataFrame({

            "No.": range(
                1,
                len(
                    results["insights"]
                ) + 1
            ),

            "Business Insight":
                results["insights"]

        }).to_excel(
            writer,
            sheet_name="Business Insights",
            index=False
        )

        # Recommendations
        pd.DataFrame({

            "No.": range(
                1,
                len(
                    results["recommendations"]
                ) + 1
            ),

            "Recommendation":
                results["recommendations"]

        }).to_excel(
            writer,
            sheet_name="Recommendations",
            index=False
        )

        # Cleaned Data
        cleaned_df.to_excel(
            writer,
            sheet_name="Cleaned Data",
            index=False
        )

    # --------------------------------------------------------
    # FORMAT EXCEL
    # --------------------------------------------------------

    workbook = load_workbook(
        EXCEL_REPORT_FILE
    )

    for worksheet in workbook.worksheets:

        for cell in worksheet[1]:

            cell.font = Font(
                bold=True
            )

            cell.alignment = Alignment(
                horizontal="center"
            )

        worksheet.freeze_panes = "A2"

        for column_cells in worksheet.columns:

            max_length = 0

            column_letter = (
                column_cells[0]
                .column_letter
            )

            for cell in column_cells:

                try:

                    max_length = max(
                        max_length,
                        len(
                            str(
                                cell.value
                            )
                        )
                    )

                except Exception:

                    pass

            worksheet.column_dimensions[
                column_letter
            ].width = min(
                max_length + 2,
                45
            )

    workbook.save(
        EXCEL_REPORT_FILE
    )

    return EXCEL_REPORT_FILE


# ============================================================
# PDF REPORT
# ============================================================

def pdf_table(
    data,
    header=True
):

    table = Table(
        data,
        repeatRows=1 if header else 0
    )

    style_commands = [

        (
            "GRID",
            (0, 0),
            (-1, -1),
            0.4,
            colors.grey
        ),

        (
            "VALIGN",
            (0, 0),
            (-1, -1),
            "MIDDLE"
        ),

        (
            "FONTSIZE",
            (0, 0),
            (-1, -1),
            7
        ),

        (
            "LEFTPADDING",
            (0, 0),
            (-1, -1),
            4
        ),

        (
            "RIGHTPADDING",
            (0, 0),
            (-1, -1),
            4
        )
    ]

    if header:

        style_commands.extend([

            (
                "BACKGROUND",
                (0, 0),
                (-1, 0),
                colors.lightgrey
            ),

            (
                "FONTNAME",
                (0, 0),
                (-1, 0),
                "Helvetica-Bold"
            )
        ])

    table.setStyle(
        TableStyle(
            style_commands
        )
    )

    return table


def generate_pdf_report(
    results,
    cleaned_df
):

    doc = SimpleDocTemplate(

        str(PDF_REPORT_FILE),

        pagesize=landscape(A4),

        rightMargin=12 * mm,

        leftMargin=12 * mm,

        topMargin=12 * mm,

        bottomMargin=12 * mm
    )

    styles = getSampleStyleSheet()

    title_style = ParagraphStyle(

        "ReportTitle",

        parent=styles["Title"],

        fontSize=22,

        leading=26,

        alignment=TA_CENTER,

        spaceAfter=12
    )

    heading_style = ParagraphStyle(

        "ReportHeading",

        parent=styles["Heading2"],

        fontSize=15,

        leading=18,

        spaceBefore=8,

        spaceAfter=8
    )

    normal_style = ParagraphStyle(

        "ReportNormal",

        parent=styles["BodyText"],

        fontSize=9,

        leading=12
    )

    story = []

    # --------------------------------------------------------
    # TITLE
    # --------------------------------------------------------

    story.append(
        Paragraph(
            "Business Data Analysis Report",
            title_style
        )
    )

    story.append(
        Paragraph(
            "Automatically generated by "
            "Business Data Analysis Application",
            normal_style
        )
    )

    story.append(
        Spacer(1, 10)
    )

    # --------------------------------------------------------
    # SUMMARY
    # --------------------------------------------------------

    story.append(
        Paragraph(
            "Executive Summary",
            heading_style
        )
    )

    summary_data = [

        ["Metric", "Value"],

        [
            "Total Records",
            f"{results['total_records']:,}"
        ],

        [
            "Total Columns",
            f"{results['total_columns']:,}"
        ],

        [
            "Primary Metric",
            results["primary_metric"]
            or "Not detected"
        ],

        [
            "Primary Metric Total",
            f"{results['metric_total']:,.2f}"
        ],

        [
            "Average",
            f"{results['metric_average']:,.2f}"
        ],

        [
            "Median",
            f"{results['metric_median']:,.2f}"
        ],

        [
            "Maximum",
            f"{results['metric_max']:,.2f}"
        ],

        [
            "Minimum",
            f"{results['metric_min']:,.2f}"
        ]
    ]

    if results["profit_margin"] is not None:

        summary_data.append(
            [
                "Profit Margin",
                f"{results['profit_margin'] * 100:.1f}%"
            ]
        )

    story.append(
        pdf_table(
            summary_data
        )
    )

    story.append(
        Spacer(1, 12)
    )

    # --------------------------------------------------------
    # DATA INFORMATION
    # --------------------------------------------------------

    story.append(
        Paragraph(
            "Detected Data Structure",
            heading_style
        )
    )

    structure_data = [

        ["Type", "Columns"],

        [
            "Date Columns",
            ", ".join(
                results["date_columns"]
            )
            or "None"
        ],

        [
            "Numeric Columns",
            ", ".join(
                results["numeric_columns"]
            )
            or "None"
        ],

        [
            "Categorical Columns",
            ", ".join(
                results["categorical_columns"]
            )
            or "None"
        ]
    ]

    story.append(
        pdf_table(
            structure_data
        )
    )

    story.append(
        PageBreak()
    )

    # --------------------------------------------------------
    # CATEGORY ANALYSIS
    # --------------------------------------------------------

    for column, table in (
        results[
            "category_tables"
        ].items()
    ):

        story.append(
            Paragraph(
                f"{column} Analysis",
                heading_style
            )
        )

        pdf_data = [
            list(table.columns)
        ]

        for _, row in table.head(20).iterrows():

            pdf_data.append(
                [
                    str(
                        value
                    )[:40]
                    for value in row.tolist()
                ]
            )

        story.append(
            pdf_table(
                pdf_data
            )
        )

        story.append(
            Spacer(1, 10)
        )

    story.append(
        PageBreak()
    )

    # --------------------------------------------------------
    # CORRELATION
    # --------------------------------------------------------

    if not results[
        "correlation"
    ].empty:

        story.append(
            Paragraph(
                "Numeric Correlation Analysis",
                heading_style
            )
        )

        corr = results[
            "correlation"
        ].round(2)

        corr_data = [
            [
                "Column"
            ]
            + list(
                corr.columns
            )
        ]

        for index, row in corr.iterrows():

            corr_data.append(
                [
                    index
                ]
                + [
                    f"{value:.2f}"
                    if not pd.isna(value)
                    else ""
                    for value in row
                ]
            )

        story.append(
            pdf_table(
                corr_data
            )
        )

        story.append(
            PageBreak()
        )

    # --------------------------------------------------------
    # FORECAST
    # --------------------------------------------------------

    if not results[
        "forecast"
    ].empty:

        story.append(
            Paragraph(
                "Forecast Analysis",
                heading_style
            )
        )

        forecast = results[
            "forecast"
        ]

        forecast_data = [

            [
                "Date",
                "Actual",
                "Forecast"
            ]
        ]

        for _, row in forecast.tail(12).iterrows():

            date_value = pd.to_datetime(
                row["Date"]
            )

            actual = (
                ""
                if pd.isna(
                    row["Actual"]
                )
                else f"{row['Actual']:,.2f}"
            )

            forecast_value = (
                ""
                if pd.isna(
                    row["Forecast"]
                )
                else f"{row['Forecast']:,.2f}"
            )

            forecast_data.append([

                date_value.strftime(
                    "%B %Y"
                ),

                actual,

                forecast_value
            ])

        story.append(
            pdf_table(
                forecast_data
            )
        )

        story.append(
            Spacer(1, 12)
        )

    # --------------------------------------------------------
    # INSIGHTS
    # --------------------------------------------------------

    story.append(
        Paragraph(
            "AI-Assisted Business Insights",
            heading_style
        )
    )

    for number, insight in enumerate(
        results["insights"],
        1
    ):

        story.append(
            Paragraph(
                f"{number}. {insight}",
                normal_style
            )
        )

        story.append(
            Spacer(1, 4)
        )

    story.append(
        Spacer(1, 10)
    )

    # --------------------------------------------------------
    # RECOMMENDATIONS
    # --------------------------------------------------------

    story.append(
        Paragraph(
            "Business Recommendations",
            heading_style
        )
    )

    for number, recommendation in enumerate(
        results["recommendations"],
        1
    ):

        story.append(
            Paragraph(
                f"{number}. {recommendation}",
                normal_style
            )
        )

        story.append(
            Spacer(1, 4)
        )

    story.append(
        Spacer(1, 15)
    )

    story.append(
        Paragraph(
            "Report generated automatically on "
            + datetime.now().strftime(
                "%d %B %Y, %I:%M %p"
            ),
            normal_style
        )
    )

    doc.build(
        story
    )

    return PDF_REPORT_FILE


# ============================================================
# MONEY FORMAT
# ============================================================

def money(value):

    if pd.isna(value):

        return "N/A"

    value = float(value)

    if abs(value) >= 10000000:

        return (
            f"₹{value / 10000000:.2f} Cr"
        )

    if abs(value) >= 100000:

        return (
            f"₹{value / 100000:.2f} L"
        )

    if abs(value) >= 1000:

        return (
            f"₹{value / 1000:.2f} K"
        )

    return f"₹{value:,.0f}"


# ============================================================
# SESSION STATE
# ============================================================

defaults = {

    "results": None,

    "cleaned_df": None,

    "mapping": {},

    "quality": None,

    "original_rows": 0,

    "duplicate_count": 0,

    "cleaned_rows": 0,

    "excel_report_ready": False,

    "pdf_report_ready": False
}

for key, value in defaults.items():

    if key not in st.session_state:

        st.session_state[key] = value


# ============================================================
# SIDEBAR
# ============================================================

st.sidebar.title(
    "📊 Business Analytics"
)

page = st.sidebar.radio(

    "Navigation",

    [

        "🏠 Dashboard",

        "📁 Dataset",

        "🔎 Data Quality",

        "🧹 Data Cleaning",

        "📊 Category Analysis",

        "🔢 Numeric Analysis",

        "📅 Time Analysis",

        "🔮 Forecast",

        "🤖 AI Insights",

        "💾 Downloads"
    ]
)

st.sidebar.markdown("---")

st.sidebar.info(

    "Upload almost any structured "
    "business CSV. The system automatically "
    "detects columns and chooses suitable "
    "analysis."
)


# ============================================================
# MAIN TITLE
# ============================================================

st.title(
    "📊 Business Data Analysis Application"
)

st.write(

    "An intelligent business analytics system "
    "that automatically detects, cleans, analyzes "
    "and reports structured business data."
)


# ============================================================
# FILE UPLOAD
# ============================================================

uploaded_file = st.file_uploader(

    "📁 Upload Business Dataset",

    type=["csv"]
)


# ============================================================
# PROCESS DATASET
# ============================================================

if uploaded_file is not None:

    try:

        df = pd.read_csv(
            uploaded_file
        )

        st.success(
            f"Dataset loaded: "
            f"{uploaded_file.name}"
        )

        # ----------------------------------------------------
        # ORIGINAL DATA TYPES
        # ----------------------------------------------------

        (
            original_dates,
            original_numeric,
            original_categories
        ) = detect_data_types(df)

        # ----------------------------------------------------
        # AUTOMATIC COLUMN DETECTION
        # ----------------------------------------------------

        mapping = detect_columns(
            df
        )

        with st.expander(
            "🔍 Automatic Column Detection"
        ):

            if mapping:

                mapping_df = pd.DataFrame([

                    {
                        "Original Column":
                            original,

                        "Detected Meaning":
                            detected
                    }

                    for original, detected
                    in mapping.items()
                ])

                st.dataframe(
                    mapping_df,
                    use_container_width=True,
                    hide_index=True
                )

            else:

                st.info(
                    "No standard business "
                    "column names were detected. "
                    "The system will analyze "
                    "the data based on its "
                    "actual data types."
                )

        # ----------------------------------------------------
        # QUALITY
        # ----------------------------------------------------

        quality = check_data_quality(
            df
        )

        st.session_state.quality = quality

        # ----------------------------------------------------
        # AUTOMATIC ANALYSIS
        # ----------------------------------------------------

        if st.button(
            "🚀 Analyze Dataset Automatically",
            type="primary"
        ):

            with st.spinner(
                "Cleaning data, detecting patterns, "
                "generating analysis and reports..."
            ):

                (
                    cleaned_df,
                    original_rows,
                    duplicate_count,
                    cleaned_rows
                ) = clean_data(df)

                results = generate_analysis(
                    cleaned_df
                )

                save_outputs(
                    results,
                    cleaned_df
                )

                generate_excel_report(
                    results,
                    cleaned_df
                )

                generate_pdf_report(
                    results,
                    cleaned_df
                )

                # Session
                st.session_state.results = (
                    results
                )

                st.session_state.cleaned_df = (
                    cleaned_df
                )

                st.session_state.mapping = (
                    mapping
                )

                st.session_state.original_rows = (
                    original_rows
                )

                st.session_state.duplicate_count = (
                    duplicate_count
                )

                st.session_state.cleaned_rows = (
                    cleaned_rows
                )

                st.session_state.excel_report_ready = (
                    True
                )

                st.session_state.pdf_report_ready = (
                    True
                )

            st.success(
                "✅ Dataset analyzed successfully! "
                "Reports generated automatically."
            )

    except Exception as error:

        st.error(
            "❌ Error while processing dataset."
        )

        st.exception(
            error
        )


# ============================================================
# LOAD SESSION
# ============================================================

results = st.session_state.results

cleaned_df = st.session_state.cleaned_df


# ============================================================
# DASHBOARD
# ============================================================

if page == "🏠 Dashboard":

    st.header(
        "🏠 Executive Business Dashboard"
    )

    if results is None:

        st.info(
            "Upload a CSV dataset and click "
            "'Analyze Dataset Automatically'."
        )

    else:

        st.subheader(
            "📌 Key Business Metrics"
        )

        primary_metric = (
            results["primary_metric"]
            or "Metric"
        )

        col1, col2, col3, col4, col5 = (
            st.columns(5)
        )

        with col1:

            st.metric(
                "Total Records",
                f"{results['total_records']:,}"
            )

        with col2:

            st.metric(
                f"Total {primary_metric}",
                money(
                    results["metric_total"]
                )
            )

        with col3:

            st.metric(
                f"Average {primary_metric}",
                money(
                    results["metric_average"]
                )
            )

        with col4:

            st.metric(
                "Maximum",
                money(
                    results["metric_max"]
                )
            )

        with col5:

            if results[
                "profit_margin"
            ] is not None:

                st.metric(
                    "Profit Margin",
                    f"{results['profit_margin'] * 100:.1f}%"
                )

            else:

                st.metric(
                    "Columns",
                    f"{results['total_columns']:,}"
                )

        st.markdown("---")

        # ----------------------------------------------------
        # CATEGORY VISUALS
        # ----------------------------------------------------

        tables = results[
            "category_tables"
        ]

        if tables:

            first_columns = list(
                tables.keys()
            )[:4]

            for start in range(
                0,
                len(first_columns),
                2
            ):

                cols = st.columns(2)

                for position, column in enumerate(
                    first_columns[start:start + 2]
                ):

                    table = tables[
                        column
                    ]

                    with cols[position]:

                        st.subheader(
                            f"{primary_metric} by {column}"
                        )

                        chart_data = (
                            table
                            .set_index(
                                column
                            )["Total"]
                        )

                        st.bar_chart(
                            chart_data
                        )

        # ----------------------------------------------------
        # MONTHLY
        # ----------------------------------------------------

        if not results[
            "monthly"
        ].empty:

            st.subheader(
                f"📈 {primary_metric} Trend"
            )

            monthly = results[
                "monthly"
            ].set_index(
                "Date"
            )["Value"]

            st.line_chart(
                monthly
            )

        st.success(
            "✅ Dashboard generated automatically "
            "from the uploaded business dataset."
        )


# ============================================================
# DATASET
# ============================================================

elif page == "📁 Dataset":

    st.header(
        "📁 Dataset Information"
    )

    if cleaned_df is None:

        st.info(
            "Analyze a dataset first."
        )

    else:

        col1, col2, col3 = st.columns(3)

        with col1:

            st.metric(
                "Records",
                f"{len(cleaned_df):,}"
            )

        with col2:

            st.metric(
                "Columns",
                f"{len(cleaned_df.columns):,}"
            )

        with col3:

            st.metric(
                "Numeric Columns",
                f"{len(results['numeric_columns']):,}"
            )

        st.markdown("---")

        st.subheader(
            "Dataset Preview"
        )

        st.dataframe(
            cleaned_df.head(100),
            use_container_width=True
        )

        st.subheader(
            "Detected Columns"
        )

        detected_data = []

        for column in cleaned_df.columns:

            if column in results[
                "date_columns"
            ]:

                data_type = "Date"

            elif column in results[
                "numeric_columns"
            ]:

                data_type = "Numeric"

            else:

                data_type = "Categorical/Text"

            detected_data.append({

                "Column":
                    column,

                "Detected Type":
                    data_type
            })

        st.dataframe(
            pd.DataFrame(
                detected_data
            ),
            use_container_width=True,
            hide_index=True
        )


# ============================================================
# DATA QUALITY
# ============================================================

elif page == "🔎 Data Quality":

    st.header(
        "🔎 Automatic Data Quality"
    )

    quality = st.session_state.quality

    if quality is None:

        st.info(
            "Upload a dataset first."
        )

    else:

        if quality[
            "status"
        ] == "Excellent":

            st.success(
                "✅ Data Quality: EXCELLENT"
            )

        elif quality[
            "status"
        ] == "Good":

            st.success(
                "🟢 Data Quality: GOOD"
            )

        elif quality[
            "status"
        ] == "Needs Attention":

            st.warning(
                "⚠️ Data Quality: NEEDS ATTENTION"
            )

        else:

            st.error(
                "❌ Data Quality: POOR"
            )

        quality_data = pd.DataFrame({

            "Quality Check": [

                "Missing Values",

                "Duplicate Rows",

                "Invalid Dates",

                "Negative Numeric Values"
            ],

            "Issues Found": [

                quality[
                    "missing_values"
                ],

                quality[
                    "duplicate_rows"
                ],

                quality[
                    "invalid_dates"
                ],

                quality[
                    "negative_numeric_values"
                ]
            ]
        })

        quality_data[
            "Status"
        ] = quality_data[
            "Issues Found"
        ].apply(

            lambda x:
            "✅ Pass"
            if x == 0
            else "⚠️ Review"
        )

        st.dataframe(
            quality_data,
            use_container_width=True,
            hide_index=True
        )

        st.metric(
            "Total Detected Issues",
            f"{quality['total_issues']:,}"
        )


# ============================================================
# DATA CLEANING
# ============================================================

elif page == "🧹 Data Cleaning":

    st.header(
        "🧹 Automatic Data Cleaning"
    )

    if cleaned_df is None:

        st.info(
            "Analyze a dataset first."
        )

    else:

        col1, col2, col3 = st.columns(3)

        with col1:

            st.metric(
                "Original Records",
                f"{st.session_state.original_rows:,}"
            )

        with col2:

            st.metric(
                "Duplicates Removed",
                f"{st.session_state.duplicate_count:,}"
            )

        with col3:

            st.metric(
                "Cleaned Records",
                f"{st.session_state.cleaned_rows:,}"
            )

        st.info(
            "The application automatically "
            "removes duplicate rows, converts "
            "numeric/date fields and handles "
            "missing values where appropriate."
        )

        st.dataframe(
            cleaned_df.head(100),
            use_container_width=True
        )


# ============================================================
# CATEGORY ANALYSIS
# ============================================================

elif page == "📊 Category Analysis":

    st.header(
        "📊 Automatic Category / Segment Analysis"
    )

    if results is None:

        st.info(
            "Analyze a dataset first."
        )

    else:

        tables = results[
            "category_tables"
        ]

        if not tables:

            st.info(
                "No categorical columns were "
                "available for category analysis."
            )

        else:

            for column, table in tables.items():

                st.subheader(
                    f"{column} Analysis"
                )

                st.dataframe(
                    table,
                    use_container_width=True,
                    hide_index=True
                )

                chart_data = (
                    table
                    .set_index(
                        column
                    )["Total"]
                )

                st.bar_chart(
                    chart_data
                )

                st.markdown("---")


# ============================================================
# NUMERIC ANALYSIS
# ============================================================

elif page == "🔢 Numeric Analysis":

    st.header(
        "🔢 Numeric Analysis"
    )

    if results is None:

        st.info(
            "Analyze a dataset first."
        )

    else:

        st.subheader(
            "Statistical Summary"
        )

        if not results[
            "numeric_summary"
        ].empty:

            st.dataframe(
                results[
                    "numeric_summary"
                ],
                use_container_width=True,
                hide_index=True
            )

        else:

            st.info(
                "No numeric columns detected."
            )

        if not results[
            "correlation"
        ].empty:

            st.subheader(
                "Correlation Analysis"
            )

            st.dataframe(
                results[
                    "correlation"
                ].round(2),
                use_container_width=True
            )

            st.info(
                "Correlation shows the strength "
                "and direction of a linear relationship "
                "between numeric variables. "
                "It does not prove causation."
            )


# ============================================================
# TIME ANALYSIS
# ============================================================

elif page == "📅 Time Analysis":

    st.header(
        "📅 Time-Based Business Analysis"
    )

    if results is None:

        st.info(
            "Analyze a dataset first."
        )

    elif results[
        "monthly"
    ].empty:

        st.info(
            "No usable date column was detected, "
            "so time-based analysis is unavailable "
            "for this dataset."
        )

    else:

        primary_metric = (
            results["primary_metric"]
        )

        st.subheader(
            f"{primary_metric} Trend"
        )

        monthly = results[
            "monthly"
        ].set_index(
            "Date"
        )["Value"]

        st.line_chart(
            monthly
        )

        st.subheader(
            "Monthly Data"
        )

        st.dataframe(
            results[
                "monthly"
            ],
            use_container_width=True,
            hide_index=True
        )


# ============================================================
# FORECAST
# ============================================================

elif page == "🔮 Forecast":

    st.header(
        "🔮 Automatic Business Forecast"
    )

    if results is None:

        st.info(
            "Analyze a dataset first."
        )

    elif results[
        "forecast"
    ].empty:

        st.info(
            "Forecasting requires a usable "
            "date column and a numeric business metric."
        )

    else:

        primary_metric = (
            results[
                "primary_metric"
            ]
        )

        st.subheader(
            f"{primary_metric} — Historical + Forecast"
        )

        forecast = results[
            "forecast"
        ]

        st.line_chart(

            forecast.set_index(
                "Date"
            )[
                [
                    "Actual",
                    "Forecast"
                ]
            ]
        )

        st.dataframe(
            forecast,
            use_container_width=True,
            hide_index=True
        )

        st.info(
            "The forecast uses a simple linear "
            "trend model based on historical monthly data. "
            "It is intended for analytical demonstration "
            "and planning support, not guaranteed prediction."
        )


# ============================================================
# AI INSIGHTS
# ============================================================

elif page == "🤖 AI Insights":

    st.header(
        "🤖 AI-Assisted Business Insights"
    )

    if results is None:

        st.info(
            "Analyze a dataset first."
        )

    else:

        st.info(
            "The system automatically generates "
            "business insights from detected "
            "patterns, statistics, categories "
            "and trends."
        )

        st.subheader(
            "📌 Key Insights"
        )

        for number, insight in enumerate(
            results["insights"],
            1
        ):

            st.write(
                f"**{number}.** {insight}"
            )

        st.markdown("---")

        st.subheader(
            "💡 Recommendations"
        )

        for number, recommendation in enumerate(
            results["recommendations"],
            1
        ):

            st.write(
                f"**{number}.** {recommendation}"
            )


# ============================================================
# DOWNLOADS
# ============================================================

elif page == "💾 Downloads":

    st.header(
        "💾 Download Reports & Data"
    )

    if results is None:

        st.info(
            "Analyze a dataset first."
        )

    else:

        # ----------------------------------------------------
        # CLEANED DATA
        # ----------------------------------------------------

        st.subheader(
            "📁 Cleaned Dataset"
        )

        st.download_button(

            "⬇️ Download Cleaned Dataset",

            cleaned_df.to_csv(
                index=False
            ),

            "cleaned_business_data.csv",

            "text/csv"
            )

        # ----------------------------------------------------
        # INSIGHTS
        # ----------------------------------------------------

        insights_text = (

            "AI-ASSISTED BUSINESS INSIGHTS\n\n"

            +

            "\n".join(

                [
                    f"{i}. {x}"
                    for i, x in enumerate(
                        results["insights"],
                        1
                    )
                ]
            )

            +

            "\n\nBUSINESS RECOMMENDATIONS\n\n"

            +

            "\n".join(

                [
                    f"{i}. {x}"
                    for i, x in enumerate(
                        results["recommendations"],
                        1
                    )
                ]
            )
        )

        st.download_button(

            "⬇️ Download Business Insights",

            insights_text,

            "ai_business_insights.txt",

            "text/plain"
        )

        # ----------------------------------------------------
        # EXCEL
        # ----------------------------------------------------

        st.markdown("---")

        st.subheader(
            "📊 Automatic Excel Report"
        )

        if EXCEL_REPORT_FILE.exists():

            with open(
                EXCEL_REPORT_FILE,
                "rb"
            ) as file:

                excel_data = file.read()

            st.success(
                "✅ Excel report ready."
            )

            st.download_button(

                "📥 Download Excel Report",

                excel_data,

                "Business_Analysis_Report.xlsx",

                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        # ----------------------------------------------------
        # PDF
        # ----------------------------------------------------

        st.markdown("---")

        st.subheader(
            "📄 Automatic PDF Report"
        )

        if PDF_REPORT_FILE.exists():

            with open(
                PDF_REPORT_FILE,
                "rb"
            ) as file:

                pdf_data = file.read()

            st.success(
                "✅ PDF report ready."
            )

            st.download_button(

                "📥 Download PDF Report",

                pdf_data,

                "Business_Analysis_Report.pdf",

                "application/pdf"
            )

        # ----------------------------------------------------
        # REPORT STATUS
        # ----------------------------------------------------

        st.markdown("---")

        st.subheader(
            "📋 Report Status"
        )

        col1, col2 = st.columns(2)

        with col1:

            if EXCEL_REPORT_FILE.exists():

                st.success(
                    "✅ Excel Report Ready"
                )

            else:

                st.warning(
                    "Excel report not found."
                )

        with col2:

            if PDF_REPORT_FILE.exists():

                st.success(
                    "✅ PDF Report Ready"
                )

            else:

                st.warning(
                    "PDF report not found."
                )


# ============================================================
# FOOTER
# ============================================================

st.markdown("---")

st.caption(

    "Business Data Analysis Application | "
    "Python + Pandas + NumPy + Streamlit + "
    "OpenPyXL + ReportLab | "
    "Automatic Data Cleaning + "
    "Dynamic Business Analysis + "
    "Forecasting + Excel/PDF Reporting + "
    "AI-Assisted Business Insights"
)