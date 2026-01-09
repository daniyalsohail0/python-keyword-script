import pandas as pd
from kw_processor import find_rows_with_keywords, append_df_to_excel


def test_find_rows_with_keywords_basic():
    df = pd.DataFrame({
        "Message": ["Error occurred", "All good", "Failed to start"],
        "Details": ["stack trace", "none", "exit code 1"]
    })
    kws = ["error", "fail"]
    res = find_rows_with_keywords(df, kws, columns=["Message"]) 
    assert len(res) == 2
    assert list(res["matched_keyword"]) == ["error", "fail"]
    assert list(res["matched_column"]) == ["Message", "Message"]


def test_append_df_to_excel(tmp_path):
    file = tmp_path / "out.xlsx"
    df1 = pd.DataFrame({"A": [1, 2], "B": ["x", "y"]})
    df2 = pd.DataFrame({"A": [3], "B": ["z"]})

    # create base file
    df1.to_excel(file, sheet_name="results", index=False)

    # append
    append_df_to_excel(str(file), df2, sheet_name="results")

    combined = pd.read_excel(file, sheet_name="results")
    assert len(combined) == 3
    assert combined.iloc[-1]["A"] == 3
    assert combined.iloc[-1]["B"] == "z"
