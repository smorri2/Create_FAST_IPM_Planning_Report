#!/usr/bin/env python3


# **********************************************************************************************************************
# **********************************************************************************************************************
# * Imports
# **********************************************************************************************************************
# **********************************************************************************************************************

# Standard library imports
from pathlib import Path
from datetime import datetime


# Third party imports
import xlsxwriter


# local application imports


# SGM Shared Module imports
from kclGetJiraSprintDates_2 import SprintDateData
from kclGetJiraSprintXlsxData_1 import JiraSprintData, JiraStoryRec


# ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
# ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
# * Class Declarations
# ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
# ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++


class SprintInfo:
    def __init__(self, name_in: str, start_date_in: datetime, end_date_in: datetime):
        self.name = name_in
        self.start_date: datetime = start_date_in
        self.end_date: datetime = end_date_in


class SprintData:
    def __init__(self):
        self.number: int = 0
        self.cur_sprint: SprintInfo | None
        self.prev_sprint: SprintInfo | None
        self.story_data: JiraSprintData | None


class AssigneesRec:
    def __init__(self, assignee_in, jira_story_in, story_points_in):
        self.assignee: str = assignee_in
        self.stories: list = [jira_story_in]
        self.total_points: int = story_points_in
        self.ws = None


class IpmPlanningSS:
    def __init__(self):
        self.workbook = None
        self.assignee_total_ws = None
        self.header_fmt = None
        self.left_fmt = None
        self.right_fmt = None
        self.center_fmt = None
        self.assignees: list = []


# ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
# ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
# * Functions
# ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
# ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

# ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
def create_ss_workbook_and_formats(sprint_to_plan: str) -> IpmPlanningSS:

    # create the IPM Planning spreadsheet data structure and then create spreadsheet workbook
    ipm_planning_ss = IpmPlanningSS()
    ipm_planning_ss.workbook = xlsxwriter.Workbook('Output files/' + sprint_to_plan + ' IPM Planning.xlsx')

    # add predefined formats to be used for formatting cells in the spreadsheet
    ipm_planning_ss.left_fmt = ipm_planning_ss.workbook.add_format({
        'font_size': 12,
        'align': 'left',
        'indent': 1
    })
    ipm_planning_ss.left_bold_fmt = ipm_planning_ss.workbook.add_format({
        'font_size': 12,
        'align': 'left',
        'bold': 1,
        'indent': 1
    })
    ipm_planning_ss.left_lv2_fmt = ipm_planning_ss.workbook.add_format({
        'font_size': 12,
        'align': 'left',
        'indent': 4
    })
    ipm_planning_ss.right_fmt = ipm_planning_ss.workbook.add_format({
        'font_size': 12,
        'align': 'right',
        'indent': 6
    })
    ipm_planning_ss.center_fmt = ipm_planning_ss.workbook.add_format({
        'font_size': 12,
        'align': 'center',
    })
    ipm_planning_ss.percent_fmt = ipm_planning_ss.workbook.add_format({
        'font_size': 12,
        'align': 'right',
        'indent': 6,
        'num_format': '0%'
    })
    ipm_planning_ss.header_fmt = ipm_planning_ss.workbook.add_format({
        'font_name': 'Calibri',
        'font_size': 13,
        'font_color': 'white',
        'text_wrap': 1,
        'align': 'center',
        'valign': 'vcenter',
        'bold': 1,
        'bg_color': '#4472C4',
        'pattern': 1,
        'border': 1
    })
    ipm_planning_ss.last_row_fmt = ipm_planning_ss.workbook.add_format({
        'font_size': 12,
        'align': 'center',
        'bottom': 6
    })
    ipm_planning_ss.totals_fmt = ipm_planning_ss.workbook.add_format({
        'font_size': 13,
        'align': 'center',
        'bold': 1,
    })

    return ipm_planning_ss


# ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
def create_assignee_worksheet(ipm_planning_ss: IpmPlanningSS) -> None:

    print('      ** Writing IPM Planning spreadsheet tab')
    ipm_planning_ss.data_ws = ipm_planning_ss.workbook.add_worksheet('IMP Planning')

    # Setup Details table layout
    ipm_planning_ss.data_ws.set_column('A:B', 14, ipm_planning_ss.center_fmt)
    ipm_planning_ss.data_ws.set_column('C:C', 80, ipm_planning_ss.center_fmt)
    ipm_planning_ss.data_ws.set_column('D:E', 15, ipm_planning_ss.center_fmt)
    ipm_planning_ss.data_ws.set_column('F:F', 20, ipm_planning_ss.center_fmt)
    ipm_planning_ss.data_ws.set_column('G:G', 18, ipm_planning_ss.center_fmt)
    ipm_planning_ss.data_ws.set_column('H:H', 20, ipm_planning_ss.center_fmt)

    # ******************************************************************
    # Set Sprint Data Table starting and ending cells.
    # params are (top_row, left_column, right_column, num_data_rows)
    # ******************************************************************
    ipm_planning_ss.detail_table = calc_table_starting_and_ending_cells(1, 'A', 'H', len(jira_sprint_data) - 1)

    return None


# ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
def update_planning_spreadsheet_assignees_stories(assignees_list: list[AssigneesRec],
                                                  jira_story_rec_in: JiraStoryRec,
                                                  jira_story_assignee_in) -> None:
    new_assignees_list = False
    if assignees_list:  # Check to see if the assignees list of assignees is empty
        new_assignee_in_list = True  # the assignees list is not empty, but this may be a new assignee to add to the list
        # iterate thru the assignees list checking to see if this assignee already has a record in the list
        for cur_assignees_rec in assignees_list:
            if cur_assignees_rec.assignee == jira_story_assignee_in:
                # this assignee already exists, so update the data for this assignee
                new_assignee_in_list = False  # assignee found, so not a new assignee for the list
                cur_assignees_rec.stories.append(jira_story_rec_in)  # add story to this assignees list of stories
                cur_assignees_rec.total_points += jira_story_rec_in.story_points
    else:
        # This is a new sprint for the sprints list
        new_assignee_in_list = True
    if new_assignees_list or new_assignee_in_list:
        # The assignees list is empty or this is a new assignee for the assignees list, so create a new assignees rec
        new_assignees_rec = AssigneesRec(jira_story_assignee_in, jira_story_rec_in, jira_story_rec_in.story_points)
        assignees_list.append(new_assignees_rec)

    return None


# ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
def calc_table_starting_and_ending_cells(top_row: int, left_col, right_col, num_data_rows) -> str:
    top_left_cell = left_col + str(top_row)
    bot_right_cell = right_col + str(top_row + num_data_rows + 1)
    table_coordinates = top_left_cell + ':' + bot_right_cell

    return table_coordinates


# ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
def create_sprint_report_spreadsheet(stories_by_assignee: list[AssigneesRec], sprint_to_plan: str) -> IpmPlanningSS:

    print('\n   Creating IPM Planning spreadsheet')
    # create the spreadsheet workbook and formats for the IPM Planning spreadsheet
    ipm_planning_ss = create_ss_workbook_and_formats(sprint_to_plan)

    ipm_planning_ss.assignees = stories_by_assignee

    # Setup the All Assignees worksheet tab to hold the totals by Assignee
    ipm_planning_ss.assignee_total_ws = ipm_planning_ss.workbook.add_worksheet('All Assignees')
    ipm_planning_ss.assignee_total_ws.set_column('A:A', 20)
    ipm_planning_ss.assignee_total_ws.write('A1', 'Assignee', ipm_planning_ss.header_fmt)
    ipm_planning_ss.assignee_total_ws.set_column('B:C', 14)
    ipm_planning_ss.assignee_total_ws.write('B1', 'Initial Story Points', ipm_planning_ss.header_fmt)
    ipm_planning_ss.assignee_total_ws.write('C1', 'Final Story Points', ipm_planning_ss.header_fmt)

    # create the worksheet and table layouts for the Metrics tab in the IPM Planning spreadsheet
    for cur_assignee in ipm_planning_ss.assignees:
        cur_assignee.ws = ipm_planning_ss.workbook.add_worksheet(cur_assignee.assignee)

        # Write out the worksheet header
        cur_assignee.ws.set_column('A:A', 12)
        cur_assignee.ws.write('A1', 'Key', ipm_planning_ss.header_fmt)

        cur_assignee.ws.set_column('B:B', 9, ipm_planning_ss.center_fmt)
        cur_assignee.ws.write('B1', 'Issue Type', ipm_planning_ss.header_fmt)

        cur_assignee.ws.set_column('C:C', 70, ipm_planning_ss.center_fmt)
        cur_assignee.ws.write('C1', 'Summary', ipm_planning_ss.header_fmt)

        cur_assignee.ws.set_column('D:D', 18, ipm_planning_ss.center_fmt)
        cur_assignee.ws.write('D1', 'Assignee', ipm_planning_ss.header_fmt)

        cur_assignee.ws.set_column('E:E', 14, ipm_planning_ss.center_fmt)
        cur_assignee.ws.write('E1', 'Status', ipm_planning_ss.header_fmt)
        cur_assignee.ws.set_column('F:F', 12, ipm_planning_ss.center_fmt)
        cur_assignee.ws.write('F1', 'Priority', ipm_planning_ss.header_fmt)

        cur_assignee.ws.set_column('G:J', 12, ipm_planning_ss.center_fmt)
        cur_assignee.ws.write('G1', 'Initial Story Points', ipm_planning_ss.header_fmt)
        cur_assignee.ws.write('H1', 'Carryover Story', ipm_planning_ss.header_fmt)
        cur_assignee.ws.write('I1', 'Remaining Story Points', ipm_planning_ss.header_fmt)
        cur_assignee.ws.write('J1', 'Final Story Points', ipm_planning_ss.header_fmt)

    return ipm_planning_ss


# ********************************************************************************************************************
def review_story_sprint_history(sprint_history: list[str], sprint_to_plan: int) -> str:

    prior_sprint = sprint_to_plan - 1
    if len(sprint_history) == 0:
        story_history_status = 'Moved To Backlog'
    else:
        prior_sprint_name = 'FASTR1i' + str(prior_sprint)
        if prior_sprint_name in sprint_history:
            story_history_status = 'Carryover Story'
        else:
            story_history_status = 'New Story'

    return story_history_status


# ********************************************************************************************************************
def write_ipm_planning_data_to_spreadsheet(ipm_planning_ss: IpmPlanningSS) -> None:

    for cur_assignees_rec in ipm_planning_ss.assignees:
        cur_assignees_rec.stories.sort(key=lambda jira_story_rec: jira_story_rec.carry_over_story, reverse=True)
        bottom_row = len(cur_assignees_rec.stories)
        ws_row = 1  # leave a empty row above the first row of data for easier manual insertion during IPM
        for cur_story in cur_assignees_rec.stories:
            ws_row += 1
            print(cur_story.key, cur_story.assignee)
            cur_assignees_rec.ws.write(ws_row, 0, cur_story.key, ipm_planning_ss.left_fmt)
            cur_assignees_rec.ws.write(ws_row, 1, cur_story.issue_type, ipm_planning_ss.left_fmt)
            cur_assignees_rec.ws.write(ws_row, 2, cur_story.summary, ipm_planning_ss.left_fmt)
            cur_assignees_rec.ws.write(ws_row, 3, cur_story.assignee, ipm_planning_ss.left_fmt)
            cur_assignees_rec.ws.write(ws_row, 4, cur_story.status, ipm_planning_ss.center_fmt)
            cur_assignees_rec.ws.write(ws_row, 5, cur_story.priority, ipm_planning_ss.center_fmt)
            cell_fmt = ipm_planning_ss.center_fmt
            cur_assignees_rec.ws.write(ws_row, 6, cur_story.story_points, cell_fmt)
            cur_assignees_rec.ws.write(ws_row, 7, cur_story.carry_over_story, cell_fmt)
            # formula to calculate remaining story points, if col H = 'Y' then it's a carryover story so return the
            # initial story points found in col G, if not then it is a new story so return an empty string
            remaining_points_fml = '=IF(H' + str(ws_row + 1) + '="Y", G' + str(ws_row + 1) + ', "")'
            cur_assignees_rec.ws.write(ws_row, 8, remaining_points_fml, cell_fmt)
            # formula to calculate fina story points, if col I is an empty string "" then it's a new story return the
            # initial story points, else it's a carryover story so return the remaining story points in col I
            final_points_fml = '=IF(I' + str(ws_row + 1) + '="", G' + str(ws_row + 1) + ', I' + str(ws_row + 1) + ')'
            cur_assignees_rec.ws.write(ws_row, 9, final_points_fml, cell_fmt)

        # leave an empty row between last story and totals row for easier insertion during IPM
        ws_row += 1
        test = (' ', ' ', ' ', ' ', ' ', ' ', ' ', ' ', ' ', ' ')
        cur_assignees_rec.ws.write_row(ws_row, 0, test, ipm_planning_ss.last_row_fmt)

        cur_assignees_rec.ws.write(ws_row + 1, 6, '=sum(G2:G' + str(ws_row + 1) + ')', ipm_planning_ss.totals_fmt)
        cur_assignees_rec.ws.write(ws_row + 1, 9, '=sum(J2:J' + str(ws_row + 1) + ')', ipm_planning_ss.totals_fmt)

    return None


# ********************************************************************************************************************
def write_ipm_planning_assignee_totals_to_spreadsheet(ipm_planning_ss: IpmPlanningSS) -> None:
    bottom_row = len(ipm_planning_ss.assignees)
    ws_row = 0
    for cur_assignees_rec in ipm_planning_ss.assignees:
        ws_row += 1
        ipm_planning_ss.assignee_total_ws.write(ws_row, 0, cur_assignees_rec.assignee, ipm_planning_ss.left_fmt)
        total_row = len(cur_assignees_rec.stories) + 4
        initial_points_total_loc = "='" + cur_assignees_rec.assignee + "'!G" + str(total_row)
        final_points_total_loc = "='" + cur_assignees_rec.assignee + "'!J" + str(total_row)
        if ws_row == bottom_row:
            cell_fmt = ipm_planning_ss.last_row_fmt
        else:
            cell_fmt = ipm_planning_ss.center_fmt
        ipm_planning_ss.assignee_total_ws.write(ws_row, 1, initial_points_total_loc, cell_fmt)
        ipm_planning_ss.assignee_total_ws.write(ws_row, 2, final_points_total_loc, cell_fmt)

    total_row = '=sum(G2:G' + str(ws_row + 1) + ')'
    ipm_planning_ss.assignee_total_ws.write(ws_row + 1, 1, '=sum(B2:B' + str(ws_row + 1) + ')', ipm_planning_ss.totals_fmt)
    ipm_planning_ss.assignee_total_ws.write(ws_row + 1, 2, '=sum(C2:C' + str(ws_row + 1) + ')', ipm_planning_ss.totals_fmt)

    return None


# ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
def get_sprint_num_to_plan() -> int:
    sprint_number: int = 0
    valid_option: bool = False

    while not valid_option:
        print('\n')
        print('*********************************************')
        print('***                                       ***')
        print('***    Enter the Sprint Number to Plan    ***')
        print('***                                       ***')
        print('*********************************************')
        user_input = input('\nEnter Sprint Number to plan (should be two digits only) ==> ')
        if user_input.isdecimal():  # Verify that the user input was a number
            sprint_number = int(user_input)
            # if sprint_number > 39 and sprint_number < 100:
            if 39 < sprint_number < 100:
                valid_option = True
            else:
                print('\n\n\nInvalid Sprint Number, valid Sprint Numbers are between 40 & 99 inclusive')
        else:
            print('\n\n\nInvalid option Selected, enter two digit sprint number only')

    return sprint_number


# ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
def get_sprint_name_and_dates(sprint_data: SprintData) -> None:
    # build the path to the Input folder where the Sprint Dates Spreadsheet and Sprint Data spreadsheets reside
    # FAST Sprint Start-End Dates.xlsx contains the name, start, and end dates for all FAST sprints in Jira
    jira_sprint_date_data = SprintDateData(Path.cwd() / 'Input files' / 'FAST Sprint Start-End Dates.xlsx')

    if jira_sprint_date_data:
        # Get the sprint date info for the sprint_number entered by the user
        jira_date_ss_rec = jira_sprint_date_data.get_sprint_data(sprint_data.number)
        sprint_data.cur_sprint = SprintInfo(jira_date_ss_rec.name,
                                            jira_date_ss_rec.start_date,
                                            jira_date_ss_rec.end_date)

        # Get the sprint date info for the previous sprint to the sprint_number entered by the user
        jira_date_ss_rec = jira_sprint_date_data.get_sprint_data(sprint_data.number - 1)
        sprint_data.prev_sprint = SprintInfo(jira_date_ss_rec.name,
                                             jira_date_ss_rec.start_date,
                                             jira_date_ss_rec.end_date)
    else:
        print('****** Error getting Sprint Name and Date Data ****** ')

    return None


# ++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
def get_jira_sprint_data_to_plan() -> SprintData:
    sprint_data = SprintData()

    # get the sprint number to process from the user via console input
    sprint_data.number = get_sprint_num_to_plan()
    # sprint_data.number = 41  # used for debugging chain number to sprint you want to use and comment out line above

    # Get the sprint date info for the sprint_number entered by the user
    get_sprint_name_and_dates(sprint_data)

    # get the jira sprint story data to process
    # build the path to the Input folder where the Sprint Dates Spreadsheet and Sprint Data spreadsheets reside
    # Jira.xlsx contains the Jira Story data for the stories in the sprint to process and report on
    jira_sprint_planning_data_path = Path.cwd() / 'Input files' / 'Jira Sprint Planning Data.xlsx'
    sprint_data.story_data = JiraSprintData(jira_sprint_planning_data_path)

    if sprint_data.cur_sprint is None or sprint_data.prev_sprint is None or sprint_data.story_data is None:
        sprint_data = None
        print('****** Bummer, didnt get either current sprint date, previous sprint date or story data  ******')

    return sprint_data


# **********************************************************************************************************************
# **********************************************************************************************************************
# * Main
# **********************************************************************************************************************
# **********************************************************************************************************************
def main():

    sprint_data = get_jira_sprint_data_to_plan()
    sprint_to_plan = sprint_data.cur_sprint.name

    print('\n\nStarting to Create IPM Planning Spreadsheet ' + sprint_to_plan)

    stories_by_assignee = []
    for cur_story_rec in sprint_data.story_data:
        # check if this is a carry over story and set carry_over_story field to 'Y' or 'N'
        if sprint_data.prev_sprint.name in cur_story_rec.sprints:
            cur_story_rec.carry_over_story = 'Y'
        else:
            cur_story_rec.carry_over_story = 'N'
        update_planning_spreadsheet_assignees_stories(stories_by_assignee, cur_story_rec, cur_story_rec.assignee)

    ipm_planning_ss = create_sprint_report_spreadsheet(stories_by_assignee, sprint_to_plan)

    write_ipm_planning_assignee_totals_to_spreadsheet(ipm_planning_ss)
    write_ipm_planning_data_to_spreadsheet(ipm_planning_ss)
    ipm_planning_ss.workbook.close()

    print('\nCompleted Create IPM Planning Spreadsheet')

    return None


if __name__ == "__main__":
    main()
