import os
from openpyxl import Workbook, load_workbook
from openpyxl.utils import exceptions
from openpyxl.styles import PatternFill, Alignment, Font, Border, Side
from datetime import datetime, timedelta
from openpyxl.formatting.rule import Rule

# import xlsxwriter

class Record():
    def __init__(self, current_date, passed_minutes):
        self.data = None
        self.current_date = current_date
        self.passed_minutes = passed_minutes
        # self.current_time = datetime.now().time()

        input_format = "%m/%d/%Y"
        output_format = "%d %b %Y"

        date_obj = datetime.strptime(self.current_date, input_format)
        today = date_obj.strftime(output_format)
        self.worksheet_name = today

    def check_timeline(self, worksheet):
        base = 6
        max_idx = (int(worksheet.max_column) - 5) // 48
        if max_idx < 0 : max_idx = 0
        for i in range(max_idx):
            date_value = str(worksheet.cell(row = 1, column = base + i * 48).value)
            if date_value == self.current_date:
                return (i, 1)
        return (max_idx, 0)
    
    def set_data(self, leagues):
        self.data = leagues

    def save_workbook(self, filename, workbook):
        success_flag = True
        try:
            workbook.save(filename)
        except PermissionError:
            print("Failed to save the workbook. The file is currently open.")
            success_flag = False
        except exceptions.ReadOnlyWorkbookException:
            print("Failed to save the workbook. The file is opened as read-only.")
            success_flag = False
        except Exception as e:
            print("An error occurred while saving the workbook:", str(e))
            success_flag = False

        if success_flag == False:
            print("Please check your file is currently open or opend as read-only")

    def get_match_count(self, worksheet):
        num_rows = int(worksheet.max_row)
        return (num_rows - 3) // 24
    
    def get_match_index(self, worksheet, match, league_name):
        base = 4
        # cnt = 0
        rows = (int(worksheet.max_row) - 3) // 24
        if rows < 0: rows = 0
        i = 0
        for i in range(rows):
            idx = base + i * 24
            cell = worksheet.cell(column = 2, row = idx)
            # print(str(cell.value))
            if "".join(str(cell.value).split()) == "": break
            if "".join(str(cell.value).split()) == "".join(str(match.match_name).split()):
                if "".join(str(worksheet.cell(column = 1, row = idx).value).split()) == "".join(str(league_name).split()):
                    # print("Checked: " + str(idx))
                    return i
                    # break
        return rows
            # cnt = cnt + 1
        # print("Checked: -1
    
    def check_match(self, worksheet, match, league_name):
        base = 4
        # cnt = 0
        rows = int(worksheet.max_row)
        for i in range((rows - 3) // 24):
            idx = base + i * 24
            cell = worksheet.cell(column = 2, row = idx)
            # print(str(cell.value))
            if "".join(str(cell.value).split()) == "": break
            if "".join(str(cell.value).split()) == "".join(str(match.match_name).split()):
                if "".join(str(worksheet.cell(column = 1, row = idx).value).split()) == "".join(str(league_name).split()):
                    # print("Checked: " + str(idx))
                    return i
            # cnt = cnt + 1
        # print("Checked: -1")
        return -1
            
    def create_match(self, worksheet, i, day_diff, match, league_name):
        
        base = 4 + i * 24
        font = Font(bold = True, name = 'Calibri', size = 11)
        alignment = Alignment(horizontal = 'center', vertical = 'center')
        pattern = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")

        border_style = "thin"
        border_color = "000000"

        border = Border(left=Side(style=border_style, color = border_color),
                        right=Side(style=border_style, color=border_color),
                        top=Side(style=border_style, color=border_color),
                        bottom=Side(style=border_style, color=border_color))

        index = self.passed_minutes // 30 + day_diff * 48

        str_check = str(worksheet.cell(row = base, column = 1).value)

        # print("CHECK: " + str_check)

        if "".join(str_check.split()) == "" or str_check == "None":
            print("NEW!")
            for j in range(24):
                worksheet.cell(row = base + j, column = 1).value = league_name
                worksheet.cell(row = base + j, column = 2).value = match.match_name
                worksheet.cell(row = base + j, column = 3).value = match.date
                worksheet.cell(row = base + j, column = 4).value = match.time

                cell = worksheet.cell(row = base + j, column = 5)
                
                if j % 3 == 0:
                    if j < 12:
                        cell.value = "H"
                    else:
                        cell.value = "O"
                elif j % 3 == 1:
                    if (j % 12) // 3 == 0:
                        cell.value = "MP"
                    else:
                        cell.value = str((j % 12) // 3) + "P"
                else:
                    if j < 12:
                        cell.value = "A"
                    else:
                        cell.value = "U"

                for k in range(5):
                    cell1 = worksheet.cell(row = base + j, column = k + 1)
                    cell1.font = font
                    cell1.alignment = alignment
                    cell1.border = border
                    if j < 12: continue
                    cell1.fill = pattern
            
                for k in range(48 * day_diff + 5, 48 * day_diff + 53):
                    cell1 = worksheet.cell(row = base + j, column = k + 1)
                    cell1.font = font
                    cell1.alignment = alignment
                    cell1.border = border
                    if j < 12: continue
                    cell1.fill = pattern
        
        worksheet.cell(row = base + 0, column = 6 + index).value = match.h1
        worksheet.cell(row = base + 1, column = 6 + index).value = match.hp1
        worksheet.cell(row = base + 2, column = 6 + index).value = match.a1
        worksheet.cell(row = base + 3, column = 6 + index).value = match.h2
        worksheet.cell(row = base + 4, column = 6 + index).value = match.hp2
        worksheet.cell(row = base + 5, column = 6 + index).value = match.a2
        worksheet.cell(row = base + 6, column = 6 + index).value = match.h3
        worksheet.cell(row = base + 7, column = 6 + index).value = match.hp3
        worksheet.cell(row = base + 8, column = 6 + index).value = match.a3
        worksheet.cell(row = base + 9, column = 6 + index).value = match.h4
        worksheet.cell(row = base + 10, column = 6 + index).value = match.hp4
        worksheet.cell(row = base + 11, column = 6 + index).value = match.a4

        worksheet.cell(row = base + 12, column = 6 + index).value = match.o1
        worksheet.cell(row = base + 13, column = 6 + index).value = match.op1
        worksheet.cell(row = base + 14, column = 6 + index).value = match.u1
        worksheet.cell(row = base + 15, column = 6 + index).value = match.o2
        worksheet.cell(row = base + 16, column = 6 + index).value = match.op2
        worksheet.cell(row = base + 17, column = 6 + index).value = match.u2
        worksheet.cell(row = base + 18, column = 6 + index).value = match.o3
        worksheet.cell(row = base + 19, column = 6 + index).value = match.op3
        worksheet.cell(row = base + 20, column = 6 + index).value = match.u3
        worksheet.cell(row = base + 21, column = 6 + index).value = match.o4
        worksheet.cell(row = base + 22, column = 6 + index).value = match.op4
        worksheet.cell(row = base + 23, column = 6 + index).value = match.u4

    def formatting_sheet_first_time(self, worksheet):
        # Formatting Worksheet For the First Time

        # merge sell
        start_cell = 'A1'
        end_cell = 'E2'

        merge_range = f"{start_cell}:{end_cell}"
        worksheet.merge_cells(merge_range)

        # set the value of merge cell
        merge_cell = worksheet[start_cell]
        merge_cell.value = "Timestamp of recording"

        # set the background color of the merge cell
        merge_cell.fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
        
        # set the font style, size, and alignment
        merge_font = Font(bold=True, name='Times New Roman', size=14)
        merge_alignment = Alignment(horizontal='center', vertical='center')
        merge_cell.font = merge_font
        merge_cell.alignment = merge_alignment

        # format title cells
        worksheet.cell(row = 3, column = 1).value = 'League'
        worksheet.cell(row = 3, column = 1).font = Font(bold=True, name='Calibri', size=11)
        worksheet.cell(row = 3, column = 2).value = 'Match'
        worksheet.cell(row = 3, column = 2).font = Font(bold=True, name='Calibri', size=11)
        worksheet.cell(row = 3, column = 3).value = 'Start Date'
        worksheet.cell(row = 3, column = 3).font = Font(bold=True, name='Calibri', size=11)
        worksheet.cell(row = 3, column = 4).value = 'Start Time'
        worksheet.cell(row = 3, column = 4).font = Font(bold=True, name='Calibri', size=11)

        # set column width
        worksheet.column_dimensions['A'].width = 60
        worksheet.column_dimensions['B'].width = 60
        worksheet.column_dimensions['C'].width = 15
        worksheet.column_dimensions['D'].width = 15

    def calculate_days_between_dates(self, date_str1, date_str2):
        # Define the input date format
        input_format = "%m/%d/%Y"

        # Convert the date strings to datetime objects
        date_obj1 = datetime.strptime(date_str1, input_format)
        date_obj2 = datetime.strptime(date_str2, input_format)

        # Calculate the difference between the two dates
        diff = date_obj2 - date_obj1

        # Extract the number of days from the difference
        days = diff.days

        return days
    
    def adding_timeline(self, worksheet, idx):
                
        # formatting timeline
        for i in range(48 * idx, 48 * idx + 48):
            worksheet.cell(row = 1, column = 6 + i).fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
            worksheet.cell(row = 2, column = 6 + i).fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")

        font = Font(bold=False, name='Calibri', size=11)

        for i in range(48):
            hour = i // 2
            if hour >= 12: hour = hour - 12
            if hour == 0: hour = 12
            minute = (i % 2) * 30
            ap = "AM"
            if i >= 24:
                ap = "PM"
            value = str(hour) + ":" + str(minute).zfill(2) + " " + ap
            
            cell = worksheet.cell(row = 2, column = 48 * idx + i + 6)
            cell.value = value
            cell.alignment = Alignment(text_rotation=90, horizontal='center', vertical='center')
            cell.font = font
            cell = worksheet.cell(row = 1, column = 48 * idx + i + 6)
            cell.value = self.current_date
            cell.alignment = Alignment(text_rotation=90, horizontal='center', vertical='center')
            cell.font = font




    def create_file_sheet(self, filename):

        create_flag = False

        if not os.path.exists(filename):
            workbook = Workbook()
            self.save_workbook(filename, workbook)
            create_flag = True

        workbook = load_workbook(filename)

        # Generate the name with today's date

        input_format = "%m/%d/%Y"
        output_format = "%d %b %Y"

        date_obj = datetime.strptime(self.current_date, input_format)
        today = date_obj.strftime(output_format)
        save_flag = False

        print(today)

        for i in range(6):
            after_days = date_obj + timedelta(days = i)

            current_date = after_days.strftime(output_format)
            isNew = False

            if current_date in workbook.sheetnames:
                print('Exists')
            else:
                print("No")
                workbook.create_sheet(title = current_date)
                isNew = True
                save_flag = True

            worksheet = workbook[current_date]
            if isNew:
                self.formatting_sheet_first_time(worksheet)

        if save_flag:
            self.save_workbook(filename, workbook)

        workbook = load_workbook(filename)

        # get data for processing
        leagues = self.data
        
        # update match record

        # initialize updating variables
        compare_date = leagues[0].matches[0].date
        update_flag = False
        cnt = 0
        date_idx = 0

        worksheet  = None

        for league in leagues:
            matches = league.matches
            for match in matches:
                # get match date's worksheet
                match_date = match.date
                
                if match_date != compare_date:
                    compare_date = match_date
                    update_flag = False
                    cnt = 0
                
                if update_flag == False:

                    date_obj = datetime.strptime(match_date, "%m/%d/%Y")
                    worksheet = workbook[date_obj.strftime("%d %b %Y")]

                    date_idx, isExist = self.check_timeline(worksheet)
                    
                    if isExist == False:
                        self.adding_timeline(worksheet, date_idx)

                    # get amount of current recoreded matches
                    cnt = self.get_match_count(worksheet)
                    print("current matches " + str(cnt))

                    update_flag = True

                    # idx = self.check_match(worksheet, match, league.league_name)
                    # update_flag = True
                    # if idx != -1:
                    #     cnt = idx

                idx = self.get_match_index(worksheet, match, league.league_name)
                
                # self.create_match(worksheet, cnt, date_idx, match, league.league_name)
                self.create_match(worksheet, idx, date_idx, match, league.league_name)
                cnt = cnt + 1

        self.save_workbook(filename, workbook)
