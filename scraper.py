import yaml
import logging
import os.path
from copy import deepcopy
from datetime import time, datetime, timedelta, date, timezone
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from helper import *
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from logging.handlers import RotatingFileHandler


SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets.readonly",
    "https://www.googleapis.com/auth/calendar",
]
CONFIG_FILE = "config.yml"

with open(CONFIG_FILE, "r") as config_file:
    config = yaml.safe_load(config_file)

UNIQNAME = config["uniqname"]
OH_GOOGLE_SHEET_ID = config["oh_google_sheets_id"]
SEMESTER = config["semester"]
MONTH_TO_NUM = config["months_to_num"]
OH_WEEKDAYS = config["oh_weekdays"]
STORED_CALENDAR_FILE = config["gcal_file"]
GCAL_ID_FILE = config["gcal_id_file"]


class Shift(yaml.YAMLObject):
    """
    Used for a storage representation of an OH shift. 
    Derived class of yaml.YAMLObject to allow for easy serialization
    to .yml files, and easy deserialization from .yml files.
    """
    yaml_tag = "!Shift"  # sets class tag for yml serialization

    def __init__(self, month: str = None, day: int = None, weekday: str = None, 
                type: str = None, start_time: str = None, end_time: str = None,
                in_main_room: bool = False, gcal_event_id: str = None):
        self.month = month
        self.day = day
        self.weekday = weekday
        self.type = type
        self.start_time = start_time
        self.end_time = end_time
        self.in_main_room = in_main_room
        self.gcal_event_id = gcal_event_id

    
    def update(self, **kwargs):
        """
        Convenience method to mass update attributes and avoid multiple assignments
        """
        # Iterate over keyword arguments passed
        for key, value in kwargs.items():
            # Update attribute if value is not None
            if value is not None:
                setattr(self, key, value)


    def __repr__(self):
        """
        Overriden string representation of class
        """
        return "%s(start_time=%r, end_time=%r, in_main_room=%r, type=%r, month=%r)" % (
            self.__class__.__name__,
            self.start_time,
            self.end_time,
            self.in_main_room,
            self.type,
            self.month
        )

    def __eq__(self, other):
        """
        Overriden equality operator for class instances. Allows for comparison
        between self.stored_schedule (mappings for created Shifts in GCal
        with valid gcal_event_id attributes) and self.schedule (which contains all 
        the Shifts that should exist based on assignments/scheduling)
        """
        if isinstance(other, Shift):
            return (
                self.start_time == other.start_time
                and self.end_time == other.end_time
                and self.in_main_room == other.in_main_room
            )

        return NotImplemented


class Scraper:
    """
    Supports using the Google Sheets API to parse the sheets and construct an
    up to date schedule, before making any necessary updates to a
    secondary "EECS 280 OH" calendar that contains office hour shifts
    using the Google Calendar API.
    """
    def __init__(self):
        creds = None
        # The file token.json stores the user's access and refresh tokens, and is
        # created automatically when the authorization flow completes for the first
        # time.
        if os.path.exists("token.json"):
            creds = Credentials.from_authorized_user_file("token.json", SCOPES)
            # If there are no (valid) credentials available, let the user log in.
        if not creds or not creds.valid:
            if (
                creds and creds.expired and creds.refresh_token
            ):  # if possible, refresh token
                creds.refresh(Request())
            else:  # otherwise, use OAuth Client credentials to generate access/refresh token creds
                flow = InstalledAppFlow.from_client_secrets_file(
                    "credentials.json", SCOPES
                )
                creds = flow.run_local_server(port=0)
            # Save the credentials for the next run
            with open("token.json", "w") as token:
                token.write(creds.to_json())

        writeback_files = [STORED_CALENDAR_FILE, GCAL_ID_FILE]

        for file_path in writeback_files:  # create files that DNE
            if os.path.isfile(file_path) == False:
                with open(file_path, "w") as file:
                    pass

        with open(STORED_CALENDAR_FILE, "r") as gcal_file:
            stored_schedule = yaml.load(gcal_file, Loader=yaml.Loader)

        with open(GCAL_ID_FILE, "r") as gcal_file:
            gcal_id_dict = yaml.load(gcal_file, Loader=yaml.Loader)

        self.sheet_service = build("sheets", "v4", credentials=creds).spreadsheets()
        self.gcal_service = build("calendar", "v3", credentials=creds)
        self.assignments = {}
        self.schedule = {}
        self.stored_schedule = stored_schedule if stored_schedule else {}
        self.gcal_id = gcal_id_dict["gcal_id"] if gcal_id_dict else None
        self.year = date.today().year

    def __get_merged_cells_with_text(self, sheet_name: str) -> list:
        result = self.sheet_service.get(
            spreadsheetId=OH_GOOGLE_SHEET_ID, includeGridData=False, ranges=[sheet_name]
        ).execute()
        assignment_sheet = result.get("sheets", [])

        if not assignment_sheet:
            print("No data found.")
            return

        merges_without_text = assignment_sheet[0]["merges"]
        merges_with_text = []
        for merge in merges_without_text:
            my_range = f"{sheet_name}!{indices_to_A1_notation(**merge)}"
            merge_result = (
                self.sheet_service.values()
                .get(spreadsheetId=OH_GOOGLE_SHEET_ID, range=my_range)
                .execute()
            )
            merges_with_text.append(merge_result)
        return merges_with_text

    def __get_rows_of_data(self, sheet_name: str, startRowIndex: int,
        startColumnIndex: int, endRowIndex: int, endColumnIndex: int,
    ):
        day_schedule_range = indices_to_A1_notation(
            startRowIndex=startRowIndex,
            startColumnIndex=startColumnIndex,
            endRowIndex=endRowIndex,
            endColumnIndex=endColumnIndex,
        )
        result = self.sheet_service.get(
            spreadsheetId=OH_GOOGLE_SHEET_ID,
            includeGridData=True,
            ranges=f"{sheet_name}!{day_schedule_range}",
        ).execute()

        return result["sheets"][0]["data"][0]["rowData"]

    def __create_event(self, gcal_id: str, gcal_shift: Shift):
        start_dt = datetime.strptime(gcal_shift.start_time, "%I:%M %p")
        start_dt = (
            datetime(
                self.year,
                MONTH_TO_NUM[gcal_shift.month],
                gcal_shift.day,
                start_dt.hour,
                start_dt.minute,
            )
            .astimezone()
            .isoformat()
        )
        end_dt = datetime.strptime(gcal_shift.end_time, "%I:%M %p")
        end_dt = (
            datetime(
                self.year,
                MONTH_TO_NUM[gcal_shift.month],
                gcal_shift.day,
                end_dt.hour,
                end_dt.minute,
            )
            .astimezone()
            .isoformat()
        )
        body = {
            "summary": f"{'MAIN ROOM - ' if gcal_shift.in_main_room is True else ''}EECS 280 OH",
            "start": {"dateTime": start_dt, "timeZone": "America/Detroit"},
            "end": {"dateTime": end_dt, "timeZone": "America/Detroit"},
        }
        return (
            self.gcal_service.events().insert(calendarId=gcal_id, body=body).execute()
        )

    def sync_assignments(self) -> None:
        """
        Syncs most current version of "Assignments" sheet, updating 
        self.assigments to a nested dictionary that maps to a list of Shifts
        with correct start_time, end_time, and in_main_room attributes for 
        the UNIQNAME configured in config.yml
        
        E.g.: self.assignments["REGULAR"]["MONDAY"] = [Shift(start_time='6:00 PM', end_time='08:00 PM']
        """
        try:
            merges_with_text = self.__get_merged_cells_with_text("Assignments")
            # sort by starting row so weekdays are mapped/printed in order (for debugging purposes)
            merges_with_text.sort(key=lambda x: A1_notation_to_indices(x["range"])["startRowIndex"])

            start_col_to_oh_type = {} # <starting cell columnn #, OH type>, <key,value> pairs
            cell_info = {} # maps from an OH type and day to a specific cell 

            for cell in merges_with_text:
                cell.update(A1_notation_to_indices(cell["range"]))
                cell_text = cell.get("values", [[""]])[0][0].upper()
                if (match := re.match(r"(?P<oh_type>(LIGHT|REGULAR|HEAVY))", cell_text)) \
                        is not None:
                    oh_type = match["oh_type"]
                    start_col_to_oh_type[cell["startColumnIndex"]] = oh_type
                    cell_info[oh_type] = cell
                    self.assignments[oh_type] = {}

            day_pattern = r"(?P<weekday>(" + "|".join(OH_WEEKDAYS) + "))"

            for cell in merges_with_text:
                cell_text = cell.get("values", [[""]])[0][0].upper()
                if (match := re.match(day_pattern, cell_text)) is not None and \
                        (cell_start_col := cell.get("startColumnIndex", float("inf"))) \
                        in start_col_to_oh_type:  
                    oh_type = start_col_to_oh_type[cell_start_col]
                    weekday = match["weekday"]
                    cell_info[oh_type].setdefault("days", {})[weekday] = cell
                    self.assignments[oh_type][weekday] = []

            # iterate over every OH type, weekday combo, and use a combination of 
            # their cell range indices to search the cell space in which the 
            # configured uniqname might hold OH on the given type and day
            for oh_type, oh_type_info in cell_info.items():
                print(oh_type, oh_type_info)
                for weekday, oh_day_info in oh_type_info["days"].items():

                    rows = self.__get_rows_of_data(
                        "Assignments",
                        oh_day_info["startRowIndex"],
                        oh_day_info["startColumnIndex"] + 1,
                        oh_day_info["endRowIndex"],
                        endColumnIndex=oh_type_info["endColumnIndex"],
                    )
                    times = [
                        cell["formattedValue"]
                        for cell in rows[0]["values"]
                        if cell.get("effectiveValue") != None
                    ]  # grabs all half hour start timestamps for OH that day

                    num_rows, num_cols = len(rows), len(times)
                    cols = [
                        [rows[i]["values"][j] for i in range(1, num_rows)]
                        for j in range(num_cols)
                    ]  # reshape representation to list cells by columns instead of list by rows

                    shift_started, shift = False, Shift(type=oh_type, weekday=weekday)
                    shifts = self.assignments[oh_type][weekday]
                    for index, col in enumerate(cols):
                        uniqname_in_col = False
                        for cell in col:
                            if cell.get("formattedValue") == UNIQNAME:
                                uniqname_in_col = True
                                is_main = cell["effectiveFormat"]["textFormat"]["bold"]

                                if shift_started and is_main != shift.in_main_room:  
                                    shift.end_time, shift_started = times[index], False
                                    shifts.append(shift)

                                if not shift_started:
                                    shift = Shift(type=oh_type, weekday=weekday)  # start a new shift
                                    shift.start_time, shift_started = times[index], True
                                    shift.in_main_room = is_main

                                break

                        if shift_started and uniqname_in_col is False: 
                            shift.end_time, shift_started = times[index], False
                            shifts.append(shift)

                    if shift_started and shift.end_time == None:  #shift went til the end
                        end = datetime.strptime(times[-1], "%I:%M %p").time()  # time obj
                        # add 30 min to start time of last column to get OH end time
                        dt_end = datetime.combine(datetime.today(), end) + timedelta(minutes=30)  
                        end_time = dt_end.time() # extract time component
                        shift.end_time = end_time.strftime("%I:%M %p")
                        shifts.append(shift)

        except HttpError as err:
            print(err)

    def sync_schedule(self) -> None:
        year = self.year
        schedule_sheet_name = f"{SEMESTER} Dynamic"

        try:
            merges_with_text = self.__get_merged_cells_with_text(schedule_sheet_name)
            month_pattern = r"(?P<month>(" + "|".join(MONTH_TO_NUM.keys()) + "))"
            semester_schedule = {}
            for cell in merges_with_text:
                cell.update(A1_notation_to_indices(cell["range"]))
                cell_text = cell.get("values", [[""]])[0][0].upper()
                # if merged cell is the title cell for a month
                if (match := re.match(month_pattern, cell_text)) is not None:  
                    month = match["month"]
                    semester_schedule[month] = cell
                    self.schedule[month] = {}

            for month, month_cell in semester_schedule.items():
                rows = self.__get_rows_of_data(
                    schedule_sheet_name,
                    month_cell["startRowIndex"]+ 2,  # data starts 2 rows below Month cell
                    month_cell["startColumnIndex"],
                    month_cell["endRowIndex"] + 7,  # data ends 7 of rows below Month cell
                    month_cell["endColumnIndex"],
                )
                for row in rows:
                    for cell in row["values"]:
                        if (cell_day := cell.get("formattedValue", "")).isnumeric():
                            cell_day = int(cell_day)
                            background_colors = cell["effectiveFormat"]["backgroundColor"]
                            dt = datetime(year, MONTH_TO_NUM[month], cell_day)
                            weekday_str = dt.strftime("%A").upper()
                            try:
                                oh_type = get_schedule_from_background_colors(**background_colors)
                                self.schedule[month][cell_day] = []
                                for s in self.assignments[oh_type][weekday_str]:
                                    shift = deepcopy(s)
                                    shift.update(
                                        month=month,
                                        day=cell_day,
                                        weekday=weekday_str,
                                        type=oh_type,
                                    )
                                    self.schedule[month][cell_day].append(shift)
                            except ValueError as e:  # TODO:
                                print(e)
        except HttpError as err:
            print(err)

    def update_calendar(self):
        try:
            if self.gcal_id is None:  # no calendar yet, create one
                body = {"summary": "EECS 280 OH", "timeZone": "America/Detroit"}
                created_calendar = self.gcal_service.calendars().insert(body=body).execute()
                self.gcal_id = gcal_id_dict = {"gcal_id": created_calendar["id"]}
                with open(GCAL_ID_FILE, "w") as gcal_id_file:  # write back in case of crash
                    yaml.dump(gcal_id_dict, gcal_id_file, sort_keys=False)

            gcal_id = self.gcal_id
            cal = self.stored_schedule
            
            # updated schedule dont match those in the calendar
            if cal != self.schedule:
                for month in self.schedule:  # for every month
                    for day, shifts_for_day in self.schedule[month].items():
                        # if shifts in old cal and new cal dont match
                        if date(self.year, MONTH_TO_NUM[month], day) >= date.today() \
                                and (old_cal_day := cal.setdefault(month, {}).setdefault(day, [])) \
                                != (new_cal_day := self.schedule[month][day]
                                ):
                            print("mistmatch")
                            for old_shift in old_cal_day:  # get rid of all old events
                                self.gcal_service.events() \
                                .delete(calendarId=gcal_id, eventId=old_shift.gcal_event_id).execute()
                            old_cal_day.clear()

                            # look up shifts that should be on this day
                            oh_type = shifts_for_day[0].type
                            print(f"shifts_for_day: {shifts_for_day}. oh_type: {oh_type}")
                            weekday = shifts_for_day[0].weekday
                            # shifts = self.assignments[oh_type][weekday]
                            for shift in shifts_for_day:
                                gcal_shift = deepcopy(shift)
                                gcal_shift.update(month=month, day=day, weekday=weekday, type=oh_type)
                                cal.setdefault(month, {}).setdefault(day, []).append(gcal_shift)
                                cal.setdefault(oh_type, {}).setdefault(weekday, []).append(gcal_shift)
                                event = self.__create_event( gcal_id=gcal_id, gcal_shift=gcal_shift) 
                                gcal_shift.update(gcal_event_id=event["id"])
                                with open(STORED_CALENDAR_FILE, "w") as gcal_file:
                                     # write back updated gcal schedule with new event
                                    yaml.dump(cal, gcal_file, sort_keys=False) 

        except HttpError as err:
            print(err)


if __name__ == "__main__":
    s = Scraper()
    s.sync_assignments()
    s.sync_schedule()
    s.update_calendar()
