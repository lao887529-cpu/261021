def create_excel_file(complete_df, incomplete_df, summary_df):
    """
    建立包含三個 Sheet 的 Excel 檔案，回傳 bytes。
    已修正空資料或空欄位造成的 TypeError。
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

                # 防止空 DataFrame 或空欄位造成 NaN
                if len(sheet_df) == 0:
                    max_length = len(str(column_name))
                else:
                    text_lengths = (
                        sheet_df[column_name]
                        .astype(str)
                        .map(len)
                    )

                    if text_lengths.empty:
                        max_length = len(str(column_name))
                    else:
                        max_length = int(text_lengths.max())

                column_width = max(
                    max_length,
                    len(str(column_name))
                ) + 2

                column_width = min(column_width, 60)

                worksheet.set_column(
                    column_num,
                    column_num,
                    column_width
                )

    output.seek(0)
    return output.getvalue()
