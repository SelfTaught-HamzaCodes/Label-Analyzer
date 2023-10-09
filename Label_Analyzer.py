"""
This is a Label Analyzing Software:
PL -> Packing List,
LB - > Labels,

After the extraction of Label (LB) and Packing List (PL),
and after all values are listed in each of the drop-down menus,

Using Regex Library for LB and Pandas for PL, we can go through each LB and tally its values against the value in PL

"""

# Imports for Front-End Development:
from tkinter import *
from tkinter import font
from tkinter import ttk
from tkinter import filedialog
from tkinter.scrolledtext import ScrolledText

# Imports for Back-End Development:
import re
import textract
import pandas as pd

# Setting up root for our Graphical User Interface (GUI).
root = Tk()
root.title("Label Checker")
root.geometry("1920x1080")
root.state("zoomed")

# Font:
main = font.Font(family="Open Sans", size=16, weight="bold")
heading = font.Font(family='Open Sans', size=12, weight='bold')
st = font.Font(family='Open Sans', size=10, weight='bold')

# Creating Horizontal & Vertical Scrollbar for our GUI.
canvas = Canvas(root)
frame = Frame(canvas)

v_scroll = Scrollbar(root, orient="vertical", command=canvas.yview)
h_scroll = Scrollbar(root, orient="horizontal", command=canvas.xview)

canvas.configure(yscrollcommand=v_scroll.set, xscrollcommand=h_scroll.set)

v_scroll.pack(side="right", fill="y")
h_scroll.pack(side="bottom", fill="x")
canvas.pack(side="left", fill="both", expand=True)

window = canvas.create_window((0, 0), window=frame, anchor="nw")

frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))

# ====================
# Global Variables:
# ====================
excel_path = ""
# excel_path: this will contain the path to the PL (Packing List) which will be extracted, using Tkinter filedialog.

pl = pd.DataFrame()
# pl: this will have the DataFrame that is extracted using Pandas and the excel_path above.

columns = []
# columns: contain extracted columns from packing_list.

active_combo_boxes = []
# will contain active combox_boxes, if in PL section N/A was not selected for a column.

text = ""
# text: this variable will hold the extracted text from LB (Labels) using textract library's function process!

labels = []
# labels: this will contain each label after extraction using regular expressions from the text variable.

variables = {"PL": {},
             "LB": {}}
# variables: this will contain information extracted for the label (LB) and packing list (PL).

equal_length_pl = False
equal_length_int_pl = 0
# equal_length_pl: this will remain true, if all Non-N/A columns have the same length in PL.
# equal_length_int_pl: this will contain / basically extracted values from each col of PL.

pl_values_extracted = False
lb_values_extracted = False
# pl/lb_values_extracted: this will become True once value are extracted from either PL / LB

format_in_pl = []
format_in_lb = []
# will remain False, until format is same, len(format_in_pl) must be equal to len(format_in_lb)

from_delimiter = False
# from_delimiter: if from_delimiter, then direct extract values can be extracted from dictionary.

pl_headings = ["COIL NUMBER",
               "COLOR/DESIGN",
               "N/W",
               "G/W",
               "THICKNESS",
               "WIDTH",
               "SIZE"]
# headings for our packing list and labels (+ FOUND).

# ========== #
# Styling
# ========== #
style = ttk.Style()
style.theme_use("clam")
#
style.configure("pl_head.TFrame", background="#1D6F42", foreground="WHITE")
style.configure("pl_body.TFrame", background="WHITE", foreground="WHITE")
#
style.configure("lb_head.TFrame", background="#41A5EE", foreground="WHITE")
style.configure("lb_body.TFrame", background="WHITE", foreground="WHITE")
#
style.configure("tl_head.TFrame", background="#D04423", foreground="WHITE")
style.configure("tl_body.TFrame", background="WHITE", foreground="WHITE")
#
style.configure("pl_main.TLabel",
                background="#1D6F42", foreground="WHITE",
                font=main)
style.configure("pl_head.TLabel",
                background="#1D6F42", foreground="WHITE",
                font=heading)
style.configure("pl_body.TLabel",
                background="WHITE", foreground="#1D6F42",
                font=heading)
#
style.configure("lb_main.TLabel",
                background="#103F91", foreground="WHITE",
                font=main)
style.configure("lb_head.TLabel",
                background="#103F91", foreground="WHITE",
                font=heading)
style.configure("lb_body.TLabel",
                background="WHITE", foreground="#103F91",
                font=heading)
#
style.configure("tl_main.TLabel",
                background="#D04423", foreground="WHITE",
                font=main)
style.configure("tl_head.TLabel",
                background="#D04423", foreground="WHITE",
                font=heading)
style.configure("tl_body.TLabel",
                background="WHITE", foreground="#D04423",
                font=heading)
#
style.configure("sr_main.TLabel",
                background="#af2031", foreground="WHITE",
                font=main)
#
style.configure('pl.TButton',
                focuscolor="WHITE",
                background="#1D6F42",
                foreground="WHITE",
                font=heading)

style.map('pl.TButton',
          background=[("pressed", "WHITE"),
                      ("disabled", "GREY")],
          foreground=[("pressed", "#1D6F42"),
                      ("disabled", "WHITE")])
#
style.configure('lb.TButton',
                focuscolor="WHITE",
                background="#103F91",
                foreground="WHITE",
                font=heading)

style.map('lb.TButton',
          background=[("pressed", "WHITE"),
                      ("disabled", "GREY")],
          foreground=[("pressed", "#103F91"),
                      ("disabled", "WHITE")])
#
style.configure('tl.TButton',
                focuscolor="WHITE",
                background="#D04423",
                foreground="WHITE",
                font=heading)

style.map('tl.TButton',
          background=[("pressed", "WHITE"),
                      ("disabled", "GREY")],
          foreground=[("pressed", "#D04423"),
                      ("disabled", "WHITE")])
#
style.configure("pl.TCombobox",
                selectbackground="#1D6F42",
                fieldbackground="#1D6F42",
                background="#1D6F42",
                selectforeground="WHITE",
                foreground="WHITE",
                arrowcolor="WHITE",
                font=heading)
#
style.map("pl.TCombobox",
          selectbackground=[("readonly", "#1D6F42")],
          fieldbackground=[("readonly", "#1D6F42")],
          background=[("readonly", "#1D6F42")])
#
style.configure("lb.TCombobox",
                selectbackground="#103F91",
                fieldbackground="#103F91",
                background="#103F91",
                selectforeground="WHITE",
                foreground="WHITE",
                arrowcolor="WHITE",
                font=heading)
#
style.map("lb.TCombobox",
          selectbackground=[("readonly", "#103F91")],
          fieldbackground=[("readonly", "#103F91")],
          background=[("readonly", "#103F91")])
#
style.configure("tl.TCombobox",
                selectbackground="#D04423",
                fieldbackground="#D04423",
                background="#D04423",
                selectforeground="WHITE",
                foreground="WHITE",
                arrowcolor="WHITE",
                font=heading)
#
style.map("tl.TCombobox",
          selectbackground=[("readonly", "#D04423")],
          fieldbackground=[("readonly", "#D04423")],
          background=[("readonly", "#D04423")])
#
style.configure("TRadiobutton",
                foreground="#103F91",
                background="WHITE",
                indicatorforeground="#103F91",
                font=heading)


# ====================
# Functions
# ====================
def check_conditions(values_of_pl=pl_values_extracted,
                     values_of_lb=lb_values_extracted):
    global active_combo_boxes
    global variables
    global equal_length_int_pl
    global format_in_pl

    # Initially Check if values in either variables["PL"] or variables["LB"] are empty:
    if not values_of_pl:
        pl_body_8_et.configure(foreground="RED")
        pl_body_8_et.delete(0, END)
        pl_body_8_et.insert(0, "Complete Packing-List Section To Analyze!")

    elif not values_of_lb:
        lb_body_1_bt["state"] = ACTIVE

        pl_body_8_et.configure(foreground="RED")
        pl_body_8_et.delete(0, END)
        pl_body_8_et.insert(0, "Complete Labels Section To Analyze!")

    else:
        format_in_pl = [format_ for format_ in active_combo_boxes[-3:] if format_ != "N/A"]

        if len(format_in_pl) > 0:

            if len(format_in_pl) == 3:

                lb_body_4_st.delete(1.0, END)
                lb_body_4_st.insert("insert", "Thickness, Width & Size were chosen!\n"
                                              "Select either:\n"
                                              "(Size) or (Thickness & Width)\n"
                                              "from Packing-List")

                # In-case all three options are selected, we DISABLE both Delimiter & Extract Labels!
                tool_body_1_bt["state"], lb_body_4_bt["state"] = DISABLED, DISABLED

            else:
                # LB: Thickness & Width | PL: Size
                if len(format_in_lb) > len(format_in_pl):
                    format_lb.set("TW_S")

                    tool_body_1_bt["state"], lb_body_4_bt["state"] = ACTIVE, DISABLED

                    tool_body_1_et.insert(0, "Specify DELIMITER to add 'THICKNESS' & 'WIDTH'")
                    tool_body_1_cb1.set(pl_body_6_cb7.get())

                    tool_body_1_cb2["values"] = ["*", ",", "x", "X"]
                    tool_body_1_cb2.set(["*", ",", "x", "X"][0])

                    lb_body_4_st.delete(1.0, END)
                    lb_body_4_st.insert("insert", "Formatting doesn't match, use Tool!\n"
                                                  "PL shows 'SIZE'\n"
                                                  "LB shows 'THICKNESS|WIDTH'")

                # LB: Size | PL: Thickness & Width
                elif len(format_in_lb) < len(format_in_pl):
                    format_lb.set("S_TW")

                    tool_body_1_bt["state"], lb_body_4_bt["state"] = ACTIVE, DISABLED

                    tool_body_1_et.insert(0, "Specify DELIMITER to add 'SIZE'")

                    tool_body_1_cb1.set(F"{pl_body_6_cb5.get()}, {pl_body_6_cb6.get()}")

                    tool_body_1_cb2["values"] = ["*", ",", "x", "X"]
                    tool_body_1_cb2.set(["*", ",", "x", "X"][0])

                    lb_body_4_st.delete(1.0, END)
                    lb_body_4_st.insert("insert", "Formatting doesn't match, use Tool!\n"
                                                  "PL shows 'THICKNESS|WIDTH'\n"
                                                  "LB shows 'SIZE'")

                else:
                    pass

        if len(format_in_pl) == 0 or len(format_in_pl) == (len(format_in_lb)):
            lb_body_4_st.delete(1.0, END)
            lb_body_4_bt["state"] = ACTIVE

            tool_body_1_et.delete(0, END)
            tool_body_1_bt["state"] = DISABLED

            lb_body_4_st.delete(1.0, END)
            lb_body_4_st.insert("insert", "Formatting Matches!"
                                          "\n")

            # display number of labels extracted, easier to tally with number of coils.
            lb_body_4_st.insert("insert", F"\n"
                                          F"Labels Extracted:{len(labels)} of {equal_length_int_pl}"
                                          F"\n\n")

            for values in variables["LB"]:
                value = "\n".join(str(v) for v in variables["LB"][values])
                lb_body_4_st.insert("insert", F"{values}:"
                                              F"\n"
                                              F"{value}"
                                              F"\n\n")

            pl_body_8_et.delete(0, END)
            pl_body_8_et.insert(0, "Press Analyze to proceed!")

            pl_body_8_bt["state"] = ACTIVE


def open_packing_list(prompt):
    global excel_path

    # To select our packing list file:
    excel_path = filedialog.askopenfilename(title="Select a file",
                                            filetypes=[("Microsoft Excel Spreadsheet", "*.xlsx"),
                                                       ("Microsoft Excel Spreadsheet", "*.xls")])

    # if excel_path refers to, if not equal to "" then proceed, "" means not file selected.
    if excel_path:

        # Getting File Name (only):
        file_name = excel_path.split("/")[-1]

        # Using pd.ExcelFile to read sheet names:
        excel_file = pd.ExcelFile(excel_path)
        excel_file_sheets = excel_file.sheet_names

        # Clear Scrolled Text:
        pl_body_2_st.delete(1.0, END)

        pl_body_2_st.insert("insert", F"File Name:\n{file_name}\n\nTotal Sheets: {len(excel_file_sheets)}\n")

        for index, sheet_name in enumerate(excel_file_sheets):
            pl_body_2_st.insert("insert", F"{index + 1}). {sheet_name}\n")

        for sheet_name in excel_file_sheets:
            pl_body_2_st.insert("insert", F"\n\nName: {sheet_name}\n")
            pl_body_2_st.insert("insert",
                                str(pd.read_excel(excel_file,
                                                  sheet_name=sheet_name).iloc[:, :3].head(3)))

        pl_body_2_st.insert("insert", F"\n\n{prompt}")

        # Inserting sheet names into our combox-box:
        pl_body_3_cb1["values"] = excel_file_sheets
        pl_body_3_cb1.set(excel_file_sheets[0])

        pl_body_3_cb2["values"] = [num for num in range(1, 11)]
        pl_body_3_cb2.set([num for num in range(1, 11)][0])

    else:
        pl_body_2_st.delete(1.0, END)
        pl_body_2_st.insert("insert", "No file was chosen!")


def get_col_names(excel_file, sheet_name, row_number):
    global pl
    global columns
    global from_delimiter

    if excel_file:

        # To reset from_delimiter.
        from_delimiter = False

        # To read Excel file using Pandas.
        pl = pd.read_excel(excel_file, sheet_name=sheet_name, header=(int(row_number) - 1))

        # Dropping any rows with in-complete information, done to remove total calculated in the end.
        pl.dropna(axis="index", how="any", inplace=True)

        # delete anything that is currently present in ScrolledTextWidget (extracted_columns)
        pl_body_4_st.delete(1.0, END)

        # Extracting each column's header:
        columns = list(pl.columns.values)

        # inserting columns names on each line:
        for index, value in enumerate(columns):
            pl_body_4_st.insert("insert", F'{index + 1}. {value}\n')

        # Insert names of columns into Combo-Box:
        columns.insert(0, "N/A")

        for combobox in pl_combo_boxes:
            combobox["values"] = columns
            combobox.set(columns[0])
            combobox.configure(justify=CENTER)

    else:
        pl_body_4_st.delete(1.0, END)
        pl_body_4_st.insert("insert", "No file was chosen!")


def extract_column_values(from_tool=from_delimiter):
    global variables
    global equal_length_pl
    global equal_length_int_pl
    global pl_values_extracted
    global active_combo_boxes

    # Set values extracted from Packing List to False, so we could update entry widget above analyze labels.
    pl_values_extracted = False

    # Check status of Combo-Boxes:
    active_combo_boxes = [cb.get() for cb in pl_combo_boxes]

    # Clear SText Widget:
    pl_body_7_st.delete(1.0, END)

    # If we re-take values from a new PL using get_col_names, we clear the old dictionary for PL values.
    if not from_delimiter:
        # To clear old values stored in our dictionary.
        variables["PL"].clear()

    # If not from_delimiter, we initialize the list to check for number of values for each column in PL | equal or not!
    equal_length_check = []

    # To get values from Packing List:
    for pos, state in enumerate(active_combo_boxes):

        # If heading is absent from variables ["PL"] we get them, else we just extract them from the dictionary.
        if pl_headings[pos] not in list(variables["PL"].keys()):

            # If Checkbox was enabled we continue the process:
            if state != "N/A":

                equal_length_check.append(len(list(pl[active_combo_boxes[pos]])))

                # If we are extracting Net-Weight / Gross-Weight from Packing List, we round it off.
                if pl_headings[pos] == "N/W" or pl_headings[pos] == "G/W":
                    variables["PL"][pl_headings[pos]] = list(pl[active_combo_boxes[pos]].round(3))

                elif pl_headings[pos] == "WIDTH":
                    variables["PL"][pl_headings[pos]] = list(pl[active_combo_boxes[pos]].astype("int64"))

                else:
                    variables["PL"][pl_headings[pos]] = list(pl[active_combo_boxes[pos]])

            else:
                pass

        else:
            pass

        if state != "N/A":
            equal_length_check.append(len(variables["PL"][pl_headings[pos]]))

            value = "\n".join(str(v) for v in variables["PL"][pl_headings[pos]])
            pl_body_7_st.insert("insert", F"{state}:"
                                          F"\n"
                                          F"{value}"
                                          F"\n\n")

    # To from_delimiter is true, we get number of values in a column using this:
    if from_delimiter:
        equal_length_int_pl = equal_length_check[0]

    # Setting Global Variables, if from_delimiter = False:
    equal_length_pl = False if len(set(equal_length_check)) != 1 else True
    equal_length_int_pl = equal_length_check[0] if equal_length_pl else 0

    # If values are successfully extracted:
    if equal_length_pl:
        pl_values_extracted = True

    # If delimiter is used, we call the function after extraction:
    if from_delimiter:
        check_conditions(values_of_pl=True, values_of_lb=True)

    else:
        check_conditions(pl_values_extracted)


def open_labels():
    global text

    # To select our label file:
    document = filedialog.askopenfilename(title="Select a file",
                                          filetypes=[("Microsoft Word Document", "*.docx")])

    # if documents means in Python an empty string is considered False | No file was selected above.
    if document:
        # To read text from our label file:
        text = textract.process(document)
        text = text.decode('utf-8')

        lines = text.split("\n")

        # Computer determines, start words / end words:
        start, end, end_point = "", "", 0

        # Within a Label, the first non-empty string will be considered the first word.
        for line in lines:
            if line != "":
                start = line
                break

        # Within a Label, the last non-empty string will be considered the last word.
        for line in reversed(lines):
            if line != "":
                end = line
                break

        # Within a Label, the first time the end-word appears must mean the end of the label.
        for index, line in enumerate(lines):
            if line == end:
                end_point = index
                break

        # Insert text (few lines) into Scrolled-Text Widget:
        lb_body_2_st.delete(1.0, END)
        lb_body_2_st.insert("insert", "Label Extracted\n\nComputer thinks, Label starts from:\n"
                                      F"{start}\n\nComputer thinks, Label ends at:\n"
                                      F"{end}\n\n"
                                      "*********************************"
                                      "\nIf correct, press Extract Label!\n"
                                      "*********************************"
                                      "\n\n"
                                      "Sample:\n")

        # Sample:
        lb_body_2_st.insert("insert", "\n".join(lines[0:end_point + 1]))

        lb_body_3_et1.delete(0, END), lb_body_3_et2.delete(0, END)
        lb_body_3_et1.insert(0, start), lb_body_3_et2.insert(0, end)

        format_in_lb.clear()

        if ("THICKNESS" or "WIDTH") in "\n".join(lines[0:end_point + 1]):
            format_in_lb.append("THICKNESS")
            format_in_lb.append("WIDTH")

        else:
            format_in_lb.append("SIZE")

        check_conditions(pl_values_extracted, values_of_lb=True)

    else:
        lb_body_2_st.delete(1.0, END)
        lb_body_2_st.insert("insert", "No file was chosen!")


def label_seperator():
    global labels
    global variables
    global equal_length_int_pl
    global active_combo_boxes
    global lb_values_extracted

    # Clearing dictionary:
    variables["LB"].clear()

    # To extract labels using start/end words from the label file.
    separate_labels = re.compile(Fr"\b{lb_body_3_et1.get()}.*?{lb_body_3_et2.get()}\b", re.DOTALL)
    labels = separate_labels.findall(text)

    for pos, state_cb in enumerate(active_combo_boxes):

        if state_cb != "N/A":

            variables["LB"][pl_headings[pos] + " FOUND"] = []

            for active_headings in variables["PL"][pl_headings[pos]]:
                for label_ in labels:
                    if str(active_headings) in label_:
                        variables["LB"][pl_headings[pos] + " FOUND"].append(str(active_headings))
                        break
                else:
                    variables["LB"][pl_headings[pos] + " FOUND"].append(0)

        else:
            continue

    check_conditions(values_of_pl=True, values_of_lb=True)


def temp_change_in_dict():
    global variables
    global columns
    global from_delimiter

    # LB: Thickness & Width | PL: Size
    if len(format_in_lb) > len(format_in_pl):

        variables["PL"]["THICKNESS"], variables["PL"]["WIDTH"] = [], []

        for size in variables["PL"]["SIZE"]:
            split = size.split(tool_body_1_cb2.get())

            variables["PL"]["THICKNESS"].append(float(split[0]))
            variables["PL"]["WIDTH"].append(int(split[1]))

        columns += ["THICKNESS", "WIDTH"]

        tool_body_1_et.delete(0, END)
        tool_body_1_et.insert(0, "THICKNESS | WIDTH added")

    # LB: Size | PL: Thickness | Width
    elif len(format_in_lb) < len(format_in_pl):

        variables["PL"]["SIZE"] = []

        for thk_wd in range(len(variables["PL"]["THICKNESS"])):
            variables["PL"]["SIZE"].append(
                str(variables["PL"]["THICKNESS"][thk_wd]) + tool_body_1_cb2.get() + str(
                    variables["PL"]["WIDTH"][thk_wd])
            )

        columns.append("SIZE")

        tool_body_1_et.delete(0, END)
        tool_body_1_et.insert(0, "SIZE added")

    # As Delimiter works for Thickness, Width & Size:
    for combobox in pl_combo_boxes[-3:]:
        combobox["values"] = columns
        combobox.set(columns[0])
        combobox.configure(justify=CENTER)

    from_delimiter = True


def get_result():

    # Summary_Results:
    total_errors = 0
    errors = {}
    compare_extractions = {}

    # Details_Results:
    mistake_in_coil = False
    value_in_packing_list = []
    value_in_labels = []
    details_body_st.delete(1.0, END)
    details_body_st.insert("insert", "Analyzed Details is as follows:"
                                     "\n\n")

    # Goes from zero to n (where n is the number of coils (values) extracted from the packing_list):
    for coils in range(equal_length_int_pl):

        # Adding each value extracted as a coil number.
        errors[coils + 1] = 0

        mistake_in_coil = False
        value_in_packing_list.clear()
        value_in_labels.clear()

        # Goes over each list in variables["PL"] like COIL NUMBER COIL 1, COLOR/DESIGN COIL 1.... COIL NUMBER COIL n...
        for pos, state in enumerate(active_combo_boxes):
            if state != "N/A":
                if str(variables["PL"][pl_headings[pos]][coils]) != str(variables["LB"][pl_headings[pos] + " FOUND"][coils]):
                    total_errors += 1
                    errors[coils + 1] += 1

                    value_in_packing_list.append((pl_headings[pos], variables["PL"][pl_headings[pos]][coils]))
                    value_in_labels.append(variables["LB"][pl_headings[pos] + " FOUND"][coils])

                    mistake_in_coil = True
            else:
                pass

        # Once above For-Loop finishes for a coil (after checking all details):
        # If a mistake was detected, COIL NUMBER along with COLUMN & VALUE is printed.
        if mistake_in_coil:
            details_body_st.insert("insert", F"COIL NO: {coils + 1}"
                                             "\n")

            for c in range(len(value_in_packing_list)):
                details_body_st.insert("insert", F"{value_in_packing_list[c][0]}: should be {value_in_packing_list[c][1]} "
                                                 F"instead of existing\n")

    # for heading__, values__ in variables["PL"].items():
        for pos_, state_ in enumerate(active_combo_boxes):
            if state_ != "N/A":

                if pl_headings[pos_] == "N/W" or pl_headings[pos_] == "G/W":
                    compare_extractions[pl_headings[pos_]] = {round(sum([float(vp) for vp in variables["PL"][pl_headings[pos_]]]), 3):
                                                              round(sum([float(vl) for vl in variables["LB"][pl_headings[pos_] + " FOUND"]]), 3)}

                else:
                    compare_extractions[pl_headings[pos_]] = \
                        {str(list(frozenset(variables["PL"][pl_headings[pos_]]))): str(list(frozenset(variables["LB"][pl_headings[pos_] + " FOUND"])))}
            else:
                pass

    # Displaying Title:
    summary_body_st.delete(1.0, END)
    summary_body_st.insert("insert",
                           "Analyzed Summary is as follows:"
                           "\n\n")

    # Displaying Overview of Extracted Values:
    summary_body_st.insert("insert",
                           "Overview of Extracted Values:"
                           "\n")

    # Displaying all common extracted values:
    for heading___, compare___ in compare_extractions.items():
        summary_body_st.insert("insert",
                               F"{heading___}: {list(compare___.items())[0][0]} in Packing-List."
                               F"\n"
                               F"{heading___}: {list(compare___.items())[0][1]} in Labels."
                               F"\n\n")

    # Displaying Total Errors:
    summary_body_st.insert("insert",
                           F"Total Errors: {total_errors}"
                           "\n\n")

    # Displaying Coils WITH errors:
    summary_body_st.insert("insert", "Errors in Coil Number:"
                                     "\n")
    for coil_no, error in errors.items():
        if error > 0:
            summary_body_st.insert("insert", F"COIL NO ({coil_no}): {error}"
                                             "\n")

    # Displaying Coils WITH NO errors:
    summary_body_st.insert("insert", "\n"
                                     "No Error in Coil Number:"
                                     "\n")
    for coil_no, error in errors.items():
        if error == 0:
            summary_body_st.insert("insert", F"COIL NO ({coil_no})"
                                             "\n")


# ====================
# Frames:
# ====================
packing_list = ttk.Frame(frame, relief=SUNKEN, borderwidth=1, style="pl_head.TFrame")
packing_list.grid(row=1, column=0, rowspan=2, padx=10, pady=10, ipadx=10, ipady=10, sticky=NSEW)

pl_heading_1_tt = ttk.Frame(frame, relief=SUNKEN, borderwidth=1, style="pl_head.TFrame")
pl_heading_1_tt.grid(row=1, column=1, padx=(10, 0), pady=(10, 0), ipadx=10, ipady=10, sticky=NSEW)
pl_body_1 = ttk.Frame(frame, relief=RAISED, borderwidth=1, style="pl_body.TFrame")
pl_body_1.grid(row=2, column=1, padx=(10, 0), pady=(0, 10), sticky=NSEW)
#
pl_heading_2 = ttk.Frame(frame, relief=SUNKEN, borderwidth=1, style="pl_head.TFrame")
pl_heading_2.grid(row=1, column=2, pady=(10, 0), padx=1, sticky=NSEW)
pl_body_2 = ttk.Frame(frame, relief=RAISED, borderwidth=2, style="pl_body.TFrame")
pl_body_2.grid(row=2, column=2, pady=(0, 10), padx=1, sticky=NSEW)
#
pl_heading_3 = ttk.Frame(frame, relief=SUNKEN, borderwidth=2, style="pl_head.TFrame")
pl_heading_3.grid(row=1, column=3, pady=(10, 0), padx=1, sticky=NSEW)
pl_body_3 = ttk.Frame(frame, relief=RAISED, borderwidth=2, style="pl_body.TFrame")
pl_body_3.grid(row=2, column=3, pady=(0, 10), padx=1, sticky=NSEW)
#
pl_heading_4 = ttk.Frame(frame, relief=SUNKEN, borderwidth=2, style="pl_head.TFrame")
pl_heading_4.grid(row=1, column=4, pady=(10, 0), padx=1, sticky=NSEW)
pl_body_4 = ttk.Frame(frame, relief=RAISED, borderwidth=2, style="pl_body.TFrame")
pl_body_4.grid(row=2, column=4, pady=(0, 10), padx=1, sticky=NSEW)
#
pl_heading_5 = ttk.Frame(frame, relief=SUNKEN, borderwidth=2, style="pl_head.TFrame")
pl_heading_5.grid(row=1, column=5, pady=(10, 0), padx=1, sticky=NSEW)
pl_body_5 = ttk.Frame(frame, relief=RAISED, borderwidth=2, style="pl_body.TFrame")
pl_body_5.grid(row=2, column=5, pady=(0, 10), padx=1, sticky=NSEW)
#
pl_heading_6 = ttk.Frame(frame, relief=SUNKEN, borderwidth=2, style="pl_head.TFrame")
pl_heading_6.grid(row=1, column=6, pady=(10, 0), padx=1, sticky=NSEW)
pl_body_6 = ttk.Frame(frame, relief=RAISED, borderwidth=2, style="pl_body.TFrame")
pl_body_6.grid(row=2, column=6, pady=(0, 10), padx=1, sticky=NSEW)
#
pl_heading_7 = ttk.Frame(frame, relief=SUNKEN, borderwidth=2, style="pl_head.TFrame")
pl_heading_7.grid(row=1, column=7, pady=(10, 0), padx=1, sticky=NSEW)
pl_body_7 = ttk.Frame(frame, relief=RAISED, borderwidth=2, style="pl_body.TFrame")
pl_body_7.grid(row=2, column=7, pady=(0, 10), padx=1, sticky=NSEW)
#
pl_heading_8 = ttk.Frame(frame, relief=SUNKEN, borderwidth=2, style="pl_head.TFrame")
pl_heading_8.grid(row=1, column=8, pady=(10, 0), padx=(1, 10), sticky=NSEW)
pl_body_8 = ttk.Frame(frame, relief=RAISED, borderwidth=2, style="pl_body.TFrame")
pl_body_8.grid(row=2, column=8, pady=(0, 10), padx=(1, 10), sticky=NSEW)
#
label = ttk.Frame(frame, relief=SUNKEN, borderwidth=1, style="lb_head.TFrame")
label.grid(row=3, column=0, rowspan=2, padx=10, pady=10, sticky=NSEW)
#
lb_heading_1 = ttk.Frame(frame, relief=SUNKEN, borderwidth=1, style="lb_head.TFrame")
lb_heading_1.grid(row=3, column=1, padx=(10, 0), pady=(10, 0), sticky=NSEW)
lb_body_1 = ttk.Frame(frame, relief=RAISED, borderwidth=1, style="lb_body.TFrame")
lb_body_1.grid(row=4, column=1, padx=(10, 0), pady=(0, 10), sticky=NSEW)
#
lb_heading_2 = ttk.Frame(frame, relief=SUNKEN, borderwidth=1, style="lb_head.TFrame")
lb_heading_2.grid(row=3, column=2, pady=(10, 0), padx=1, sticky=NSEW)
lb_body_2 = ttk.Frame(frame, relief=RAISED, borderwidth=1, style="lb_body.TFrame")
lb_body_2.grid(row=4, column=2, pady=(0, 10), padx=1, sticky=NSEW)
#
lb_heading_3 = ttk.Frame(frame, relief=SUNKEN, borderwidth=1, style="lb_head.TFrame")
lb_heading_3.grid(row=3, column=3, pady=(10, 0), padx=1, sticky=NSEW)
lb_body_3 = ttk.Frame(frame, relief=RAISED, borderwidth=1, style="lb_body.TFrame")
lb_body_3.grid(row=4, column=3, pady=(0, 10), padx=1, sticky=NSEW)
#
lb_heading_4 = ttk.Frame(frame, relief=SUNKEN, borderwidth=1, style="lb_head.TFrame")
lb_heading_4.grid(row=3, column=4, pady=(10, 0), padx=1, sticky=NSEW)
lb_body_4 = ttk.Frame(frame, relief=RAISED, borderwidth=1, style="lb_body.TFrame")
lb_body_4.grid(row=4, column=4, pady=(0, 10), padx=1, sticky=NSEW)
#
lb_heading_5 = ttk.Frame(frame, relief=SUNKEN, borderwidth=1, style="lb_head.TFrame")
lb_heading_5.grid(row=3, column=5, pady=(10, 0), padx=1, sticky=NSEW)
lb_body_5 = ttk.Frame(frame, relief=RAISED, borderwidth=1, style="lb_body.TFrame")
lb_body_5.grid(row=4, column=5, pady=(0, 10), padx=1, sticky=NSEW)
#
tool = ttk.Frame(frame, relief=SUNKEN, borderwidth=2, style="tl_head.TFrame")
tool.grid(row=3, column=6, rowspan=2, padx=10, pady=10, sticky=NSEW)
#
tool_heading_1 = ttk.Frame(frame, relief=SUNKEN, borderwidth=1, style="tl_head.TFrame")
tool_heading_1.grid(row=3, column=7, columnspan=2, pady=(10, 0), padx=(1, 10), sticky=NSEW)
tool_body_1 = ttk.Frame(frame, relief=RAISED, borderwidth=1, style="tl_body.TFrame")
tool_body_1.grid(row=4, column=7, columnspan=2, pady=(0, 10), padx=(1, 10), sticky=NSEW)
#
summary = ttk.Frame(frame, relief=GROOVE, borderwidth=1)
summary.grid(row=5, column=0, padx=10, pady=10, sticky=NSEW)
summary_body = ttk.Frame(frame, relief=GROOVE, borderwidth=1)
summary_body.grid(row=5, column=1, columnspan=3, padx=10, pady=10, sticky=NSEW)
#
details = ttk.Frame(frame, relief=GROOVE, borderwidth=2)
details.grid(row=5, column=5, padx=10, pady=10, sticky=NSEW)
details_body = ttk.Frame(frame, relief=GROOVE, borderwidth=2)
details_body.grid(row=5, column=6, columnspan=3, padx=(10, 0), pady=10, sticky=NSEW)

#
word_part = ttk.LabelFrame(frame,
                           text="Process for Labels",
                           labelanchor=NW,
                           width=960,
                           height=750,
                           relief=SUNKEN)

excel_part = ttk.LabelFrame(frame,
                            text="Process for Packing List",
                            labelanchor=NW,
                            width=960,
                            height=1450,
                            relief=SUNKEN)

results_part = ttk.LabelFrame(frame,
                              text="Results",
                              labelanchor=NW,
                              width=850,
                              height=2240,
                              relief=SUNKEN)

# ====================
# Widgets:
# ====================

title = ttk.Label(frame,
                  text="LABEL ANALYZER",
                  font=main,
                  background="#AF2031",
                  foreground="WHITE")

title.grid(row=0, column=0, columnspan=9, padx=10, pady=10, ipady=5, ipadx=10, sticky=NSEW)
#
pl_title = ttk.Label(packing_list,
                     text="PACKING LIST",
                     anchor=CENTER,
                     style="pl_main.TLabel")
#
pl_heading_1_tt = ttk.Label(pl_heading_1_tt,
                            text="STEP 1: GET PACKING LIST",
                            anchor=CENTER,
                            style="pl_head.TLabel")

pl_body_1_lb = ttk.Label(pl_body_1,
                         text="Choose Packing List",
                         anchor=CENTER,
                         style="pl_body.TLabel")

pl_body_1_bt = ttk.Button(pl_body_1,
                          text="Choose",
                          style="pl.TButton",
                          command=lambda: open_packing_list("Note: In-order to open file on window, right click on "
                                                            "the file of your choice and then click on 'Open' to "
                                                            "view contents "))

#
pl_heading_2_tt = ttk.Label(pl_heading_2,
                            text="STEP 2: PACKING LIST DETAILS",
                            anchor=CENTER,
                            style="pl_head.TLabel")

pl_body_2_st = ScrolledText(pl_body_2,
                            width=40,
                            height=9,
                            foreground="#1D6F42",
                            font=heading,
                            wrap="word")
pl_body_2_st.insert("insert", "Once a file is chosen,"
                              "\n"
                              "Details will be displayed here")
#
pl_heading_3_tt = ttk.Label(pl_heading_3,
                            text="STEP 3: SHEET NAME|ROW NUMBER",
                            anchor=CENTER,
                            style="pl_head.TLabel")

pl_body_3_lb1 = ttk.Label(pl_body_3,
                          text="Sheet Name",
                          anchor=CENTER,
                          style="pl_body.TLabel")

pl_body_3_cb1 = ttk.Combobox(pl_body_3,
                             style="pl.TCombobox",
                             justify=CENTER,
                             state="readonly")

pl_body_3_lb2 = ttk.Label(pl_body_3,
                          text="Row Number",
                          anchor=CENTER,
                          style="pl_body.TLabel")

pl_body_3_cb2 = ttk.Combobox(pl_body_3,
                             style="pl.TCombobox",
                             justify=CENTER)
#
pl_heading_4_tt = ttk.Label(pl_heading_4,
                            text="STEP 4: EXTRACTED COLUMNS",
                            anchor=CENTER,
                            style="pl_head.TLabel")

pl_body_4_bt = ttk.Button(pl_body_4,
                          text="Extract Columns",
                          style="pl.TButton",
                          command=lambda: get_col_names(excel_path,
                                                        pl_body_3_cb1.get(),
                                                        pl_body_3_cb2.get()))

pl_body_4_st = ScrolledText(pl_body_4,
                            width=32,
                            height=7,
                            foreground="#1D6F42",
                            font=heading,
                            wrap="word")
#
pl_heading_5_tt = ttk.Label(pl_heading_5,
                            text="STEP 5: SET COLUMNS A",
                            anchor=CENTER,
                            style="pl_head.TLabel")

pl_heading_6_tt = ttk.Label(pl_heading_6,
                            text="STEP 5: SET COLUMNS B",
                            anchor=CENTER,
                            style="pl_head.TLabel")
#
pl_body_5_lb1 = ttk.Label(pl_body_5,
                          text="Coil Number",
                          anchor=W,
                          style="pl_body.TLabel")

pl_body_5_cb1 = ttk.Combobox(pl_body_5,
                             style="pl.TCombobox",
                             state="readonly")

pl_body_5_lb2 = ttk.Label(pl_body_5,
                          text="Color|Design",
                          anchor=W,
                          style="pl_body.TLabel")

pl_body_5_cb2 = ttk.Combobox(pl_body_5,
                             style="pl.TCombobox",
                             state="readonly")

pl_body_5_lb3 = ttk.Label(pl_body_5,
                          text="Net Weight",
                          anchor=W,
                          style="pl_body.TLabel")

pl_body_5_cb3 = ttk.Combobox(pl_body_5,
                             style="pl.TCombobox",
                             state="readonly")

pl_body_5_lb4 = ttk.Label(pl_body_5,
                          text="Gross Weight",
                          anchor=W,
                          style="pl_body.TLabel")

pl_body_5_cb4 = ttk.Combobox(pl_body_5,
                             style="pl.TCombobox",
                             state="readonly")

pl_body_6_lb5 = ttk.Label(pl_body_6,
                          text="Thickness",
                          anchor=W,
                          style="pl_body.TLabel")

pl_body_6_cb5 = ttk.Combobox(pl_body_6,
                             style="pl.TCombobox",
                             state="readonly")

pl_body_6_lb6 = ttk.Label(pl_body_6,
                          text="Width",
                          anchor=W,
                          style="pl_body.TLabel")

pl_body_6_cb6 = ttk.Combobox(pl_body_6,
                             style="pl.TCombobox",
                             state="readonly")

pl_body_6_lb7 = ttk.Label(pl_body_6,
                          text="Size",
                          anchor=W,
                          style="pl_body.TLabel")

pl_body_6_cb7 = ttk.Combobox(pl_body_6,
                             style="pl.TCombobox",
                             state="readonly")

pl_body_6_bt = ttk.Button(pl_body_6,
                          text="Extract Values",
                          style="pl.TButton",
                          command=lambda: extract_column_values(from_tool=from_delimiter))
#
pl_combo_boxes = [
    pl_body_5_cb1,
    pl_body_5_cb2,
    pl_body_5_cb3,
    pl_body_5_cb4,
    pl_body_6_cb5,
    pl_body_6_cb6,
    pl_body_6_cb7]
#
pl_heading_7_tt = ttk.Label(pl_heading_7,
                            text="STEP 6: EXTRACTED VALUES",
                            anchor=CENTER,
                            style="pl_head.TLabel")

pl_body_7_st = ScrolledText(pl_body_7,
                            width=32,
                            height=9,
                            foreground="#1D6F42",
                            font=heading,
                            wrap="word")
#
pl_heading_8_tt = ttk.Label(pl_heading_8,
                            text="STEP 7: ANALYZE",
                            anchor=CENTER,
                            style="pl_head.TLabel")

pl_body_8_et = ttk.Entry(pl_body_8,
                         foreground="#1D6F42")

pl_body_8_bt = ttk.Button(pl_body_8,
                          text="Analyze Labels",
                          style="pl.TButton",
                          state=DISABLED,
                          command=get_result)
#
lb_title = ttk.Label(label,
                     text="LABELS",
                     anchor=CENTER,
                     style="lb_main.TLabel")
#
lb_heading_1_tt = ttk.Label(lb_heading_1,
                            text="STEP 1: GET LABELS",
                            anchor=CENTER,
                            style="lb_head.TLabel")

lb_body_1_lb = ttk.Label(lb_body_1,
                         text="Choose Labels",
                         anchor=CENTER,
                         style="lb_body.TLabel")

lb_body_1_bt = ttk.Button(lb_body_1,
                          text="Choose",
                          style="lb.TButton",
                          command=open_labels,
                          state=DISABLED)
#
lb_heading_2_tt = ttk.Label(lb_heading_2,
                            text="STEP 2: LABELS DETAILS",
                            anchor=CENTER,
                            style="lb_head.TLabel")

lb_body_2_st = ScrolledText(lb_body_2,
                            width=40,
                            height=9,
                            foreground="#103F91",
                            font=heading,
                            wrap="word")
#
lb_heading_3_tt = ttk.Label(lb_heading_3,
                            text="STEP 3: START WORD|END WORD",
                            anchor=CENTER,
                            style="lb_head.TLabel")

lb_body_3_lb1 = ttk.Label(lb_body_3,
                          text="Start Word",
                          anchor=CENTER,
                          style="lb_body.TLabel")

lb_body_3_et1 = ttk.Entry(lb_body_3,
                          style="lb.TEntry")

lb_body_3_lb2 = ttk.Label(lb_body_3,
                          text="End Word",
                          anchor=CENTER,
                          style="lb_body.TLabel")

lb_body_3_et2 = ttk.Entry(lb_body_3,
                          style="lb.TEntry")
#
lb_heading_4_tt = ttk.Label(lb_heading_4,
                            text="STEP 4: EXTRACTED LABELS",
                            anchor=CENTER,
                            style="lb_head.TLabel")

lb_body_4_bt = ttk.Button(lb_body_4,
                          text="Extract Label",
                          style="lb.TButton",
                          command=label_seperator,
                          state=DISABLED)

lb_body_4_st = ScrolledText(lb_body_4,
                            width=32,
                            height=7,
                            foreground="#103F91",
                            font=heading,
                            wrap="word")
#
lb_heading_5_tt = ttk.Label(lb_heading_5,
                            text="FORMATTING",
                            anchor=CENTER,
                            style="lb_head.TLabel")

lb_body_5_lb = ttk.Label(lb_body_5,
                         text="How are labels formatted ?",
                         anchor=CENTER,
                         style="lb_body.TLabel")

format_lb = StringVar()
format_lb.set("S_TW")
lb_body_5_rb1 = ttk.Radiobutton(lb_body_5,
                                text="Size (0.11 * 914)",
                                value="S_TW",
                                variable=format_lb,
                                takefocus=0)
lb_body_5_rb2 = ttk.Radiobutton(lb_body_5,
                                text="Thickness (0.11) | Width (914)",
                                value="TW_S",
                                variable=format_lb,
                                takefocus=0)
#
tool_title = ttk.Label(tool,
                       text="TOOL",
                       anchor=CENTER,
                       style="tl_main.TLabel")

tool_heading_1_tt = ttk.Label(tool_heading_1,
                              text="DELIMITATION",
                              anchor=CENTER,
                              style="tl_head.TLabel")

tool_body_1_lb1 = ttk.Label(tool_body_1,
                            text="Status",
                            anchor=CENTER,
                            style="tl_body.TLabel")

tool_body_1_et = ttk.Entry(tool_body_1,
                           foreground="#D04423")

tool_body_1_lb2 = ttk.Label(tool_body_1,
                            text="Column",
                            anchor=CENTER,
                            style="tl_body.TLabel")

tool_body_1_cb1 = ttk.Combobox(tool_body_1,
                               style="tl.TCombobox")

tool_body_1_lb3 = ttk.Label(tool_body_1,
                            text="Delimiter",
                            anchor=CENTER,
                            style="tl_body.TLabel")

tool_body_1_cb2 = ttk.Combobox(tool_body_1,
                               style="tl.TCombobox")

tool_body_1_bt = ttk.Button(tool_body_1,
                            text="Delimit",
                            style="tl.TButton",
                            command=temp_change_in_dict,
                            state=DISABLED)
#
summary_title = ttk.Label(summary,
                          text="ANALYZED"
                               "\n"
                               "SUMMARY",
                          anchor=CENTER,
                          style="sr_main.TLabel")

summary_body_st = ScrolledText(summary_body,
                               width=32,
                               height=18,
                               foreground="#af2031",
                               font=heading,
                               wrap="word")
#
details_title = ttk.Label(details,
                          text="ANALYZED"
                               "\n"
                               "   DETAILS",
                          anchor=CENTER,
                          style="sr_main.TLabel")

details_body_st = ScrolledText(details_body,
                               width=32,
                               height=18,
                               foreground="#af2031",
                               font=heading,
                               wrap="word")

#
pl_title.pack(fill=BOTH, expand=True)

pl_heading_1_tt.pack(fill=BOTH, expand=True)
pl_body_1_lb.grid(row=0, column=0, rowspan=2, ipadx=30, ipady=30, sticky=S)
pl_body_1_bt.grid(row=2, column=0, rowspan=2, padx=30, pady=(0, 30), sticky=N)

pl_heading_2_tt.pack(fill=BOTH, expand=True)
pl_body_2_st.grid(row=0, column=0, rowspan=4, sticky=NSEW)

pl_heading_3_tt.pack(fill=BOTH, expand=True, ipady=10, ipadx=10)
pl_body_3_lb1.grid(row=0, column=0, padx=(85, 0), ipady=15, sticky=N)
pl_body_3_cb1.grid(row=1, column=0, padx=(85, 0), sticky=NSEW)
pl_body_3_lb2.grid(row=2, column=0, padx=(85, 0), ipady=15, sticky=N)
pl_body_3_cb2.grid(row=3, column=0, padx=(85, 0), sticky=NSEW)

pl_heading_4_tt.pack(fill=BOTH, expand=True, ipady=10, ipadx=10)
pl_body_4_bt.grid(row=0, column=0, padx=5, pady=5, sticky=NSEW)
pl_body_4_st.grid(row=1, column=0, rowspan=3, sticky=NSEW)

pl_heading_5_tt.pack(fill=BOTH, expand=True, ipady=10, ipadx=10)
pl_body_5_lb1.grid(row=0, column=0, padx=5, pady=10, sticky=N)
pl_body_5_cb1.grid(row=0, column=1, padx=5, pady=10, sticky=NSEW)
pl_body_5_lb2.grid(row=1, column=0, padx=5, pady=10, sticky=N)
pl_body_5_cb2.grid(row=1, column=1, padx=5, pady=10, sticky=NSEW)
pl_body_5_lb3.grid(row=2, column=0, padx=5, pady=10, sticky=N)
pl_body_5_cb3.grid(row=2, column=1, padx=5, pady=10, sticky=NSEW)
pl_body_5_lb4.grid(row=3, column=0, padx=5, pady=10, sticky=N)
pl_body_5_cb4.grid(row=3, column=1, padx=5, pady=10, sticky=NSEW)

pl_heading_6_tt.pack(fill=BOTH, expand=True, ipady=10, ipadx=10)
pl_body_6_lb5.grid(row=0, column=0, padx=5, pady=10, sticky=N)
pl_body_6_cb5.grid(row=0, column=1, padx=5, pady=10, sticky=NSEW)
pl_body_6_lb6.grid(row=1, column=0, padx=5, pady=10, sticky=N)
pl_body_6_cb6.grid(row=1, column=1, padx=5, pady=10, sticky=NSEW)
pl_body_6_lb7.grid(row=2, column=0, padx=5, pady=10, sticky=N)
pl_body_6_cb7.grid(row=2, column=1, padx=5, pady=10, sticky=NSEW)
pl_body_6_bt.grid(row=3, column=0, columnspan=2, padx=5, pady=(10, 5), sticky=NSEW)

pl_heading_7_tt.pack(fill=BOTH, expand=True, ipady=10, ipadx=10)
pl_body_7_st.grid(row=0, column=0, sticky=NSEW)

pl_heading_8_tt.pack(fill=BOTH, expand=True, ipady=10, ipadx=10)
pl_body_8_et.grid(row=0, column=0, rowspan=2, padx=10, pady=(25, 0), sticky=NSEW)
pl_body_8_bt.grid(row=2, column=0, rowspan=2, padx=10, pady=(50, 0), sticky=NSEW)
#
lb_title.pack(fill=BOTH, expand=True)

lb_heading_1_tt.pack(fill=BOTH, expand=True, ipady=10, ipadx=10)
lb_body_1_lb.grid(row=0, column=0, rowspan=2, ipadx=60, ipady=30, sticky=S)
lb_body_1_bt.grid(row=2, column=0, rowspan=2, padx=30, pady=(0, 30), sticky=N)

lb_heading_2_tt.pack(fill=BOTH, expand=True, ipady=10, ipadx=10)
lb_body_2_st.grid(row=0, column=0, rowspan=4, sticky=NSEW)

lb_heading_3_tt.pack(fill=BOTH, expand=True, ipady=10, ipadx=10)
lb_body_3_lb1.grid(row=0, column=0, padx=(85, 0), ipady=15, sticky=N)
lb_body_3_et1.grid(row=1, column=0, padx=(85, 0), sticky=NSEW)
lb_body_3_lb2.grid(row=2, column=0, padx=(85, 0), ipady=15, sticky=N)
lb_body_3_et2.grid(row=3, column=0, padx=(85, 0), sticky=NSEW)

lb_heading_4_tt.pack(fill=BOTH, expand=True, ipady=10, ipadx=10)
lb_body_4_bt.grid(row=0, column=0, padx=5, pady=5, sticky=NSEW)
lb_body_4_st.grid(row=1, column=0, rowspan=3, sticky=NSEW)

lb_heading_5_tt.pack(fill=BOTH, expand=True, ipady=10, ipadx=10)
lb_body_5_lb.grid(row=0, column=0, rowspan=2, ipadx=25, ipady=25, sticky=NSEW)
lb_body_5_rb1.grid(row=2, column=0, padx=(10, 0), pady=5, sticky=NSEW)
lb_body_5_rb2.grid(row=3, column=0, padx=(10, 0), pady=5, sticky=NSEW)
#
tool_title.pack(fill=BOTH, expand=True)

tool_heading_1_tt.pack(fill=BOTH, expand=True, ipady=10, ipadx=10)
tool_body_1_lb1.grid(row=0, column=0, padx=(150, 5), pady=10, sticky=S)
tool_body_1_et.grid(row=0, column=1, padx=(5, 100), pady=10, sticky=S)
tool_body_1_lb2.grid(row=1, column=0, padx=(150, 5), pady=10, sticky=S)
tool_body_1_cb1.grid(row=1, column=1, padx=(5, 100), pady=10, sticky=S)
tool_body_1_lb3.grid(row=2, column=0, padx=(150, 5), pady=10, sticky=S)
tool_body_1_cb2.grid(row=2, column=1, padx=(5, 100), pady=10, sticky=S)
tool_body_1_bt.grid(row=3, column=0, columnspan=2, padx=(130, 0), pady=10, sticky=N)
#
summary_title.pack(fill=BOTH, expand=True)
summary_body_st.pack(fill=BOTH, expand=True)
#
details_title.pack(fill=BOTH, expand=True)
details_body_st.pack(fill=BOTH, expand=True)

check_conditions()

root.mainloop()
