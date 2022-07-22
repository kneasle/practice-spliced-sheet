#!/usr/bin/env python3

import re
from openpyxl.styles.fonts import Font
from openpyxl.workbook import Workbook
from openpyxl.styles import Alignment
from openpyxl.utils import get_column_letter
from typing import Dict, Iterable, List, Tuple


BELL_NAMES = "1234567890ETABCD"

STAGE = 8
TENOR = BELL_NAMES[STAGE - 1]
ROUNDS = BELL_NAMES[:STAGE]


def main():
    methods = load_methods()
    touches = read_touches("touches", methods)
    touches.sort(key = lambda touch: (touch.length, -touch.runs))

    for t in touches:
        print(f"{t.length}: {t.call_string} ({t.calling_position_string}, {t.runs} runs)")
    print(f"{len(touches)} touches")

    # Save sheet
    write_spreadsheet(methods, touches, "Lincoln.xlsx")


###########
# METHODS #
###########


def load_methods() -> Dict[str, "Method"]:
    return {
        # Core 7
        "C": Method("Cambridge", "-38-14-1258-36-14-58-16-78,12"),
        "Y": Method("Yorkshire", "-38-14-58-16-12-38-14-78,12"),
        "S": Method("Superlative", "-36-14-58-36-14-58-36-78,12"),
        "B": Method("Bristol", "-58-14.58-58.36.14-14.58-14-18,18"),
        "E": Method("Lessness", "-38-14-56-16-12-58-14-58,12"),
        "W": Method("Cornwall", "-56-14-56-38-14-58-14-58,18"),
        "L": Method("London", "38-38.14-12-38.14-14.58.16-16.58,12"),
        # Lincoln 11
        "N": Method("Double Norwich", "-14-36-58-18,18"),
        "V": Method("Deva", "-58-14.58-58.36-14-58-36-18,18"),
        "A": Method("Lancashire", "58-58.14-58-36-14-58.14-14.78,12"),
        "T": Method("Ytterbium", "-38-14-1256-16-12-58.16-12.78,12"),
        # Lincoln 15
        "D": Method("Double Coslany", "-14.58.36.14.58-18,18"),
        "M": Method("Mareham", "-58-14.58-12.38-12-18.36.12-18,18"),
        "G": Method("Glasgow", "36-56.14.58-58.36-14-38.16-16.38,18"),
        "R": Method("Carolina Reaper", "38-38.18-56-18-34-18.16-16.78,12"),
    }


class Method:
    def __init__(self, name: str, pn_string: str) -> None:
        self.name: str = name

        BOB_PLACES = [1, 4]
        SINGLE_PLACES = [1, 2, 3, 4]

        # Parse the place notation
        pns = parse_pn(pn_string)

        # Use the place notation to generate the first lead of the method
        current_row = ROUNDS
        lead_rows = []
        for places in pns:
            lead_rows.append(current_row)
            current_row = transpose_row_by_pn(current_row, places)

        self.lead_rows: List[str] = lead_rows
        self.lead_head_plain: str = current_row
        self.lead_head_bob: str = transpose_row_by_pn(lead_rows[-1], BOB_PLACES)
        self.lead_head_single: str = transpose_row_by_pn(lead_rows[-1], SINGLE_PLACES)


###################
# LOADING TOUCHES #
###################


def read_touches(path: str, methods: Dict[str, Method]) -> List["Touch"]:
    # I could do many things to make this regex more readable: like, for example, not using a regex.
    # Instead I will leave understanding this regex as a challenge for the reader.
    line_split_regex = re.compile(
        r"^\s*(?P<length>\d+)\s+(?P<calling>\S+)(\s+(?P<notes>\S.+?))?\s*$",
    )

    touches = []
    for line in open(path).read().splitlines(False):
        if line.lstrip().startswith("#"):
            continue

        re_match = line_split_regex.match(line)
        if re_match is None:
            print(f"Can't parse line {line.__repr__()}");
            exit(1)

        length = int(re_match.group("length"))
        call_string = re_match.group("calling")
        notes = re_match.group("notes")

        touches.append(Touch(length, call_string, notes, methods))

    return touches


class Touch:
    def __init__(self, length: int, call_string: str, notes: str, methods: Dict[str, Method]) -> None:
        self.length = length
        self.call_string = call_string
        self.notes = notes

        # Parse the call string into a sequence of leads
        lead_regex = re.compile(r"(?P<method>[a-zA-Z])(?P<call>[*.])?")
        leads = [
            (match.group("method"), match.group("call"))
            for match in lead_regex.finditer(call_string)
        ]

        # Determine which methods are rung
        self.method_counts = {}
        for shorthand, _ in leads:
            if shorthand in self.method_counts:
                self.method_counts[shorthand] += 1
            else:
                self.method_counts[shorthand] = 1

        # Convert the lead sequence into a sequence of rows and a calling string
        rows, self.calling_position_string = Touch.gen_rows_and_calling(call_string, leads, methods)
        # Check that the given length was correct
        if self.length != len(rows):
            raise ValueError(
                f"{self.call_string} is given len {self.length} but has {len(rows)} rows"
            )
        # Count runs
        run_regex_front = re.compile("^(1234|2345|3456|4567|5678|4321|5432|6543|7654|8765).*$")
        run_regex_back  = re.compile("^.*(1234|2345|3456|4567|5678|4321|5432|6543|7654|8765)$")
        self.runs = 0
        for row in rows:
            if run_regex_front.match(row):
                self.runs += 1
            if run_regex_back.match(row):
                self.runs += 1

        # TODO: Replace e.g. 'HHH' with '3H' in the call string

    @classmethod
    def gen_rows_and_calling(cls, call_string: str, leads, methods: Dict[str, Method]) -> Tuple[List[str], str]:
        calling_position_string = ""
        rows = []
        lead_head = ROUNDS
        for method_shorthand, call_shorthand in leads:
            method = methods[method_shorthand]
            # Add the rows for this lead
            rows += [transpose_row_by_row(lead_head, row) for row in method.lead_rows]
            # Decide which lead head to go to
            if call_shorthand is None:
                lead_head = transpose_row_by_row(lead_head, method.lead_head_plain)
            elif call_shorthand == ".":
                lead_head = transpose_row_by_row(lead_head, method.lead_head_bob)
                calling_position_string += calling_pos_at(lead_head, is_single=False)
            elif call_shorthand == "*":
                lead_head = transpose_row_by_row(lead_head, method.lead_head_single)
                calling_position_string += "s"
                calling_position_string += calling_pos_at(lead_head, is_single=True)
            else:
                raise ValueError(f"Invalid call {call_shorthand}")
        # Check for early rounds
        try:
            # Rounds always appears at the start, but snap finishes will make rounds appear again
            rounds_index = rows.index(ROUNDS, 1)
            # Trim anything from rounds onwards
            rows = rows[:rounds_index]
        except ValueError:
            # Rounds doesn't appear early, so check that the comp comes round
            if lead_head != ROUNDS:
                print(f"{call_string} doesn't come round")
                assert False

        return (rows, calling_position_string)



def calling_pos_at(row: str, is_single: bool = False) -> str:
    calling_positions = "LBTFVMWH" if is_single else "LIBFVMWH"
    return calling_positions[row.index(TENOR)]


#######################
# SPREADSHEET WRITING #
#######################


def write_spreadsheet(methods: Dict[str, Method], touches: List[Touch], path: str):
    # Header looks like:
    #       A/1    B/2    C/3   D/4    E/5     F/6      G/7     H/8        >=I/9
    #     +----+--------+-----+-----+------+--------+--------+-------+---------------
    #   1 |    |        |     |     |      |        |        |       |
    #     +----+--------+-----+-----+------+--------+--------+-------+---------------
    #     |    |                                                     |
    #   2 |    |                    <title>                          |
    #     |    |                                                     | <methods> ....
    #     +----+--------+-----+-----+------+--------+--------+-------+
    #   3 |    |    <made by me>    |      |        | back-  |       |
    #     +----+--------+-----+-----+ runs + queens + rounds + notes +---------------
    #   4 |    | length | <calling> |      |        |        |       | <method groups> ...
    #     +----+--------+-----+-----+------+--------+--------+-------+---------------
    #     |    |        |           |      |        |        |       |
    # >=5 |    |        |           |      |        |        |       |
    #                                ...
    #                             <touches>
    #                                ...
    #     |    |        |           |      |        |        |       |
    #     +----+--------+-----+-----+------+--------+--------+-------+---------------
    #     |                                   <note on calls>
    #     +----+--------+-----+-----+------+--------+--------+-------+---------------

    workbook = Workbook()
    sheet = workbook.active

    vertical_text = Alignment(text_rotation=90, horizontal="center")
    centre_text = Alignment(horizontal="center")

    font_family = "Times New Roman"
    font_size = 10

    # Set default font for all cells
    for col in range(2, 9 + len(methods) + 1):
        for row in range(2, 5 + len(touches) + 1):
            sheet.cell(row, col).font = Font(name=font_family, size=font_size)

    # === TOP-LEFT CORNER ===
    # title
    sheet.merge_cells("B2:H2")
    sheet["B2"] = "50 PPE Touches"
    sheet["B2"].font = Font(name=font_family, size=font_size * 4, bold=True)
    sheet["B2"].alignment = Alignment(horizontal="center", vertical="center")
    # made by me
    sheet.merge_cells("B3:D3")
    sheet["B3"] = "Compiled by Ben White-Horne"
    sheet["B3"].font = Font(name=font_family, size=font_size * 1.6, bold=True)
    sheet["B3"].alignment = Alignment(horizontal="center", vertical="center")
    # length/calling
    sheet["B4"] = "Length"
    sheet["B4"].alignment = centre_text
    sheet.merge_cells("C4:D4")
    sheet["C4"] = "Calling (* for single, . for bob, all near calls)"
    # music
    sheet.merge_cells("E3:E4")
    sheet.merge_cells("F3:F4")
    sheet.merge_cells("G3:G4")
    sheet["E3"].alignment = centre_text
    sheet["E3"] = "Runs"
    sheet["F3"].alignment = vertical_text
    sheet["F3"] = "Queens"
    sheet["G3"].alignment = vertical_text
    sheet["G3"] = "Backrounds"
    # notes
    sheet.merge_cells("H3:H4")
    sheet["H3"].alignment = centre_text
    sheet["H3"] = "Notes"

    # === METHODS ===
    for idx, shorthand in enumerate(methods):
        column = 9 + idx
        method = methods[shorthand]
        # Determine the name
        name = method.name
        if not method.name.startswith(shorthand):
            name += f" ({shorthand})"
        # Set the cell
        sheet.merge_cells(start_column=column, start_row=2, end_column=column, end_row=3)
        cell = sheet.cell(column = column, row = 2)
        cell.value = name
        cell.alignment = vertical_text

    # === TOUCHES ===
    for idx, touch in enumerate(touches):
        row = 5 + idx
        # Touch info
        for col in [2, 4, 5]:
            sheet.cell(row, col).alignment = centre_text
        for col in [3, 4]:
            sheet.cell(row, col).font = Font(name="Fira Code", size=font_size)
        sheet.cell(row, 2).value = touch.length
        sheet.cell(row, 3).value = touch.call_string
        sheet.cell(row, 4).value = touch.calling_position_string
        sheet.cell(row, 5).value = touch.runs
        # 6 is Queens
        # 7 is Backrounds
        sheet.cell(row, 8).value = touch.notes
        # Method matrix
        for meth_idx, shorthand in enumerate(methods):
            if shorthand in touch.method_counts:
                cell = sheet.cell(row, 9 + meth_idx)
                cell.value = touch.method_counts[shorthand]
                cell.alignment = centre_text

    # === ROW/COLUMN SIZES ===
    # Rows
    sheet.row_dimensions[2].height = font_size * 6 # Title
    sheet.row_dimensions[3].height = font_size * 4.5 # 'made by me'
    for row in range(4, len(touches) + 5 + 1):
        sheet.row_dimensions[row].height = font_size * 1.4
    # Columns
    vertical_text_column_width = 3
    sheet.column_dimensions["B"].width = 6.5 # 'Length'
    sheet.column_dimensions["C"].width = max_len((t.call_string for t in touches)) * 1.2 # 'Calling'
    sheet.column_dimensions["D"].width = max_len((t.calling_position_string for t in touches)) * 1.2
    sheet.column_dimensions["E"].width = 5 # 'Runs'
    sheet.column_dimensions["F"].width = vertical_text_column_width # 'Queens'
    sheet.column_dimensions["G"].width = vertical_text_column_width # 'Backrounds'
    sheet.column_dimensions["H"].width = max((len(t.notes or "") for t in touches)) * 0.9 # 'Notes'
    for method_idx in range(len(methods)):
        col_name = get_column_letter(9 + method_idx)
        sheet.column_dimensions[col_name].width = vertical_text_column_width

    workbook.save(path)


def max_len(strings: Iterable[str]) -> int:
    """Returns the maximum length of an `Iterable` of strings"""
    return max((len(s or "") for s in strings))

#######################
# PLACE NOTATION CODE #
#######################


# Place notation code is casually stolen from Wheatley:
# https://github.com/kneasle/wheatley/blob/9141bf8511dce737208731e55bfe138d48845319/wheatley/row_generation/helpers.py#L57


def transpose_row_by_row(lhs: str, rhs: str) -> str:
    return "".join((lhs[int(ch) - 1] for ch in rhs))


def transpose_row_by_pn(row: str, places: List[int]) -> str:
    new_row = ""
    index = 0
    while index < len(row):
        place = index + 1
        if place in places:
            # Don't do a swap
            new_row += row[index]
            index += 1
        else:
            assert place + 1 not in places
            # Swap two bells round
            new_row += row[index + 1]
            new_row += row[index]
            index += 2
    return new_row


def parse_pn(pn_str: str, expect_symmetric: bool = False) -> List[List[int]]:
    """Convert a place notation string into a list of places."""
    if "," in pn_str:
        pns = []
        for part in pn_str.split(","):
            pns += parse_pn(part, True)
        return pns

    if expect_symmetric:
        symmetric = not pn_str.startswith("+")
    else:
        symmetric = pn_str.startswith("&")

    # Assumes a valid place notation string is delimited by `.`
    # These can optionally be omitted around an `-` or `x`
    # We substitute to ensure `-` is surrounded by `.` and replace any `..` caused by `--` => `.-..-.
    dot_delimited_string = re.sub("[.]*[x-][.]*", ".-.", pn_str).strip(".&+ ")
    deduplicated_string = dot_delimited_string.replace("..", ".").split(".")

    # We suppress the type error here, because mypy will assign the list comprehension type 'List[object]',
    # not 'List[Places]'.
    converted: List[List[int]] = [
        [convert_bell_string(y) for y in place] if place != "-" else []  # type: ignore
        for place in deduplicated_string
    ]

    if symmetric:
        return converted + list(reversed(converted[:-1]))
    return converted


def convert_bell_string(bell: str) -> int:
    """Convert a single-char string representing a bell into an integer."""
    try:
        return BELL_NAMES.index(bell) + 1
    except ValueError as e:
        raise ValueError(f"'{bell}' is not known bell symbol") from e


if __name__ == "__main__":
    main()
