import pandas as pd
import numpy as np
import re
import ipaddress
from dateutil import parser
from collections import Counter


def clean_text(value):
    """移除前後空白，並將空字串轉為 NaN。"""
    if pd.isna(value):
        return np.nan

    text = str(value).strip()

    if text == "":
        return np.nan

    return text


def is_invalid_date(value):
    """檢查日期是否無效。PERMANENT 視為有效值。"""
    if pd.isna(value):
        return False

    text = str(value).strip()

    if text.lower() == "permanent":
        return False

    try:
        parser.parse(text)
        return False
    except Exception:
        return True


def is_invalid_phone(value):
    """檢查完整電話號碼是否有效。"""
    if pd.isna(value):
        return False

    digits = re.sub(r"\D", "", str(value))
    return not (7 <= len(digits) <= 15)


def is_valid_ip(value):
    """檢查 IP 是否為有效 IPv4 或 IPv6。"""
    try:
        ipaddress.ip_address(str(value))
        return True
    except Exception:
        return False


def clean_dataframe(df):
    """
    清洗 Excel 資料。

    回傳：
    - complete_df：完整資料
    - incomplete_df：不完整資料，含 Data_Issues 欄
    - summary_df：統計表
    - issue_map：每筆紀錄的問題明細
    """

    df = df.copy()

    # ----------------------------------------
    # 1. 基本文字清理
    # ----------------------------------------
    for col in df.columns:
        if df[col].dtype == object:
            df[col] = df[col].map(clean_text)

    issue_map = {idx: [] for idx in df.index}

    def add_issue(index, column, reason):
        issue_map[index].append(f"{column}: {reason}")

    # ----------------------------------------
    # 2. NULL、空值、N/A
    # ----------------------------------------
    na_tokens = {
        "null", "n/a", "na", "none",
        "nan", "-", "unknown"
    }

    for col in df.columns:
        for idx, value in df[col].items():

            if pd.isna(value):
                add_issue(idx, col, "NULL/空值")

            elif isinstance(value, str) and value.lower() in na_tokens:
                add_issue(idx, col, "NULL/N/A/未知值")

    # ----------------------------------------
    # 3. 日期格式檢查
    # ----------------------------------------
    date_columns = [
        "DOB",
        "Hire_Date",
        "Contract_End_Date",
        "Registration_Date",
        "Last_Login_Timestamp",
        "Last_Purchase_Date"
    ]

    for col in date_columns:
        if col not in df.columns:
            continue

        for idx, value in df[col].items():
            if is_invalid_date(value):
                add_issue(idx, col, "日期格式不一致或無效日期")

    # ----------------------------------------
    # 4. 電話格式檢查
    # --------------------------------------------
    # 只檢查完整電話號碼；
    # Mobile_Prefix 是區號，例如 +853。

    phone_columns = [
        "Phone_Number",
        "Emergency_Contact_Phone"
    ]

    for col in phone_columns:
        if col not in df.columns:
            continue

        for idx, value in df[col].items():
            if is_invalid_phone(value):
                add_issue(idx, col, "電話格式不一致")

    if "Mobile_Prefix" in df.columns:
        for idx, value in df["Mobile_Prefix"].items():
            if pd.notna(value) and not re.match(
                r"^\+?\d{1,5}$",
                str(value)
            ):
                add_issue(idx, "Mobile_Prefix", "電話區號格式不一致")

    # ----------------------------------------
    # 5. 電郵格式檢查
    # ----------------------------------------
    if "Email_Address" in df.columns:
        email_pattern = r"^[^\s@]+@[^\s@]+\.[^\s@]+$"

        for idx, value in df["Email_Address"].items():
            if pd.notna(value) and not re.match(email_pattern, str(value)):
                add_issue(idx, "Email_Address", "電郵格式不一致")

    # ----------------------------------------
    # 6. IP 格式檢查
    # ----------------------------------------
    if "IP_Address" in df.columns:
        for idx, value in df["IP_Address"].items():
            if pd.notna(value) and not is_valid_ip(value):
                add_issue(idx, "IP_Address", "IP 格式不一致")

    # ----------------------------------------
    # 7. 數值異常與格式檢查
    # ----------------------------------------
    numeric_rules = {
        "Age": (0, 120),
        "Salary_Monthly": (0, float("inf")),
        "Credit_Score": (0, 1000),
        "Total_Spend_USD": (0, float("inf")),
        "Discount_Rate": (0, 1),
        "Customer_Satisfaction_Score": (1, 10)
    }

    for col, (min_value, max_value) in numeric_rules.items():
        if col not in df.columns:
            continue

        numeric_values = pd.to_numeric(
            df[col].astype(str).str.replace(",", "", regex=False),
            errors="coerce"
        )

        for idx in df.index:
            original_value = df.loc[idx, col]

            if pd.isna(original_value):
                continue

            converted_value = numeric_values.loc[idx]

            if pd.isna(converted_value):
                add_issue(idx, col, "數值格式不一致")

            elif converted_value < min_value or converted_value > max_value:
                add_issue(idx, col, "數值異常")

    # ----------------------------------------
    # 8. 大小寫混亂檢查
    # --------------------------------------------
    # 只在同一個文字的大小寫變體內比較。
    # 例如：MOP / mop 會比較；
    # 但 MOP / USD 不會互相比較。

    case_check_columns = [
        "Gender", "Country", "State_Province", "City", "Timezone",
        "Department", "Job_Title", "Employment_Status", "Currency",
        "Education_Level", "Account_Status", "Device_Type", "Browser",
        "Operating_System", "Two_Factor_Auth", "VIP_Tier",
        "Payment_Method", "Subscription_Plan", "Renewal_Option",
        "Data_Quality_Label"
    ]

    for col in case_check_columns:
        if col not in df.columns:
            continue

        non_null_values = df[col].dropna().astype(str)

        lowercase_groups = non_null_values.groupby(
            non_null_values.str.lower()
        )

        for lower_value, group in lowercase_groups:

            variants = list(group.unique())

            if len(variants) <= 1:
                continue

            variant_counts = {
                variant: (group == variant).sum()
                for variant in variants
            }

            dominant_variant = max(
                variant_counts,
                key=variant_counts.get
            )

            for idx in df.index:
                value = df.loc[idx, col]

                if pd.isna(value):
                    continue

                value = str(value)

                if value.lower() == lower_value and value != dominant_variant:
                    add_issue(
                        idx,
                        col,
                        f"大小寫混亂：建議統一為 {dominant_variant}"
                    )

    # ----------------------------------------
    # 9. 分開完整與不完整資料
    # ----------------------------------------
    complete_indices = [
        idx for idx, issues in issue_map.items()
        if len(issues) == 0
    ]

    incomplete_indices = [
        idx for idx, issues in issue_map.items()
        if len(issues) > 0
    ]

    complete_df = df.loc[complete_indices].copy()
    incomplete_df = df.loc[incomplete_indices].copy()

    incomplete_df["Data_Issues"] = [
        "; ".join(issue_map[idx])
        for idx in incomplete_df.index
    ]

    # ----------------------------------------
    # 10. 建立統計表
    # --------------------------------------------
    # 統計「受影響的紀錄數」，
    # 而不是問題出現次數。

    total_records = len(df)
    complete_count = len(complete_df)
    incomplete_count = len(incomplete_df)

    summary_rows = [
        ["總紀錄數", total_records, "100.00%"],
        [
            "完整資料紀錄數",
            complete_count,
            f"{complete_count / total_records * 100:.2f}%"
        ],
        [
            "不完整資料紀錄數",
            incomplete_count,
            f"{incomplete_count / total_records * 100:.2f}%"
        ]
    ]

    issue_record_counts = Counter()

    for idx, issues in issue_map.items():
        unique_reasons = set()

        for issue in issues:
            reason = issue.split(": ", 1)[1]
            unique_reasons.add(reason)

        for reason in unique_reasons:
            issue_record_counts[reason] += 1

    for reason, affected_records in issue_record_counts.most_common():
        summary_rows.append([
            f"問題類型：{reason}",
            affected_records,
            f"{affected_records / total_records * 100:.2f}%"
        ])

    summary_df = pd.DataFrame(
        summary_rows,
        columns=[
            "統計項目",
            "受影響紀錄數",
            "佔總紀錄比例"
        ]
    )

    return complete_df, incomplete_df, summary_df, issue_map


def create_excel_file(complete_df, incomplete_df, summary_df):
    """
    建立包含三個 Sheet 的 Excel 檔案，回傳 bytes。
    已修正 pandas Arrow 字串與 NaN 造成的 TypeError。
    """
    import io
    from pandas import ExcelWriter

    output = io.BytesIO()

    with ExcelWriter(output, engine="xlsxwriter") as writer:

        complete_df.to_excel(
            writer,
            sheet_name="Complete_Data",
            index=False
        )

        incomplete_df.to_excel(
            writer,
            sheet_name="Incomplete_Data",
            index=False
        )

        summary_df.to_excel(
            writer,
            sheet_name="Summary_Statistics",
            index=False
        )

        for sheet_name, sheet_df in [
            ("Complete_Data", complete_df),
            ("Incomplete_Data", incomplete_df),
            ("Summary_Statistics", summary_df)
        ]:
            worksheet = writer.sheets[sheet_name]

            for column_num, column_name in enumerate(sheet_df.columns):

                # 預設欄寬：欄位名稱長度
                max_length = len(str(column_name))

                if len(sheet_df) > 0:

                    # 先把 NaN 轉為空字串，避免 .map(len) 出錯
                    text_series = (
                        sheet_df[column_name]
                        .fillna("")
                        .astype(str)
                    )

                    # 使用 .str.len()，比 .map(len) 更相容
                    text_lengths = text_series.str.len()

                    if not text_lengths.empty:
                        longest = text_lengths.max()

                        if pd.notna(longest):
                            max_length = max(
                                max_length,
                                int(longest)
                            )

                column_width = min(max_length + 2, 60)

                worksheet.set_column(
                    column_num,
                    column_num,
                    column_width
                )

    output.seek(0)
    return output.getvalue()
