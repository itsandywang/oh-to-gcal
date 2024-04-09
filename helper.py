import re

# EFFECTS: takes in a 0-index numeric column, returns the corresponding spreadsheet column notation for it
# Example: 0->A, 26->AA, 27->AB, 
def col_num_to_letter_str(n: int) -> str:
    result = ""
    while n >= 0:
        n, remainder = divmod(n, 26)
        result = chr(65 + remainder) + result
        n -= 1  # Adjusting for zero indexing
    return result


# EFFECTS: takes in a string representing a Google Sheets column, returns the 0-indexed numeric 
def letter_str_to_col_num(letters: str) -> int:
    n = len(letters) - 1
    offset = ord(letters[-1]) - 65 
    return n * 26 + offset


# REQUIRES: startRowIndex < endRowIndex, startColumnIndex < endColumnIndex. Starts are inclusive; Ends are exclusive.
# EFFECTS: returns a A1 notation string that represents the region bounded by the indices
# See R1C1 documenation at https://developers.google.com/sheets/api/guides/concepts#expandable-1
def indices_to_A1_notation(startRowIndex:int, startColumnIndex:int, endColumnIndex:int, endRowIndex:int, **kwargs) -> str:
    return f"{col_num_to_letter_str(startColumnIndex)}{startRowIndex+1}:{col_num_to_letter_str(endColumnIndex-1)}{endRowIndex}"


def A1_notation_to_indices(range:str) -> dict[str:int]:
    cell_range = range.split('!')[-1] # cell range will be after <sheet_name>!, if any sheet_name is provided
    # start_cell, end_cell = cell_range.split(':')
    match = re.search(r"(?P<startColumnIndex>[A-Z]+)(?P<startRowIndex>[0-9]+):(?P<endColumnIndex>[A-Z]+)(?P<endRowIndex>[0-9]+)", cell_range)
    match = match.groupdict()
    match["startRowIndex"] = int(match["startRowIndex"]) - 1 # needs to be zero indexed
    match["endRowIndex"] = int(match["endRowIndex"]) # no change, 0-indexing and exclusive demands balance out
    match["startColumnIndex"] = letter_str_to_col_num(match["startColumnIndex"])
    match["endColumnIndex"] = letter_str_to_col_num(match["endColumnIndex"]) + 1 # column gets converted to 0-indexed, but now needs to be exclusive
    return match
     