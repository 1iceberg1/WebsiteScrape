from bs4 import BeautifulSoup
import datetime
import os

class League():
    def __init__(self):
        self.league_name = ""
        self.matches = []

class Match():
    def __init__(self):
        self.match_name = ""
        self.time = ""
        self.date = ""
        self.time = ""
        self.h1 = ""
        self.hp1 = ""
        self.a1 = ""
        self.h2 = ""
        self.hp2 = ""
        self.a2 = ""
        self.h3 = ""
        self.hp3 = ""
        self.a3 = ""
        self.h4 = ""
        self.hp4 = ""
        self.a4 = ""
        self.o1 = ""
        self.op1 = ""
        self.u1 = ""
        self.o2 = ""
        self.op2 = ""
        self.u2 = ""
        self.o3 = ""
        self.op3 = ""
        self.u3 = ""
        self.o4 = ""
        self.op4 = ""
        self.u4 = ""

class Analyzer():
    def __init__(self, page, isToday, current_date, passed_minutes):
        
        # # read from external file
        # url = "saved_html.html"
        # page = open(url)
        # soup = BeautifulSoup(page.read(), 'html.parser')
        self.isToday = isToday
        self.current_date = current_date
        self.passed_minutes = passed_minutes

        self.leagues = []

        soup = BeautifulSoup(page, 'html.parser')

        # find all possible candidate tables by tbody
        candidate_tables = soup.find_all('tbody')

        # initialize football tables list
        football_tables = []

        # filter football tables
        for candidate_table in candidate_tables:
            if 'soclid' in candidate_table.attrs:
                if candidate_table['soclid'] == '0': continue
                football_tables.append(candidate_table)

        # print(football_tables[1])

        # process loop
        all_match_count = 0

        cnt_match = 0

        prev_match_time = ""

        for football_table in football_tables:
            # get league_name
            league_name = self.get_league_name(football_table) 
            
            league = League()
            league.league_name = league_name

            matches = []
            
            # filter out SABA ...
            # if self.isToday:
            if "SABA" in league_name:
                continue

            if "TEST" in league_name:
                continue

            if "CORNER" in league_name:
                continue

            if "BOOKING" in league_name:
                continue

            if "OFFSIDE" in league_name:
                continue

            if "HOME TEAM" in league_name:
                continue

            if "WINNER" in league_name:
                continue

            if "TOTAL GOALS" in league_name:
                continue

            if "WHICH" in league_name:
                continue
            
            if " - " in league_name:
                continue
            
            tr_tag = football_table.tr
            count = 0
            # count tr tag
            while tr_tag.next_sibling is not None:
                tr_tag = tr_tag.next_sibling
                count = count + 1
            print(count)
            if count < 1: continue
            # 
            tr_tag = football_table.tr

            match = Match()
            
            # check if it's pre-game
            tm_value = self.get_time_value(tr_tag.next_sibling.td, match)
            if self.isToday and self.check_time_value(tm_value) == False:
                continue

            match_flag = False
            index = 0

            for i in range(count):

                tr_tag = tr_tag.next_sibling
                td_tag = tr_tag.td
                
                # get time value
                time_value = self.get_time_value(td_tag, match)
                td_tag = td_tag.next_sibling

                if time_value != '':
                    time_flag = True
                    if match_flag:
                        matches.append(match)
                        cnt_match = cnt_match + 1
                    match_flag = True
                    match = Match()

                    if "/" in time_value:
                        date = time_value[:5]
                        input_format = "%d/%m"
                        output_format = "%m/%d"

                        date_obj = datetime.datetime.strptime(date, input_format)
                        date = date_obj.strftime(output_format)
                        days = self.calculate_days_between_dates(self.current_date[:5], date)
                        year = datetime.date.today().year
                        
                        if days < 0:
                            year = year + 1
                            days = days + 365

                        if days > 6:
                            match_flag = False
                            continue
                        match.date = date + "/" + str(year)
                    else :
                        if self.isToday:
                            match_time = time_value[-7:]
                            time_flag = self.check_match_time(prev_match_time, match_time)
                            if time_flag == False:
                                date_object = datetime.datetime.strptime(self.current_date, "%m/%d/%Y")
                                date_object = date_object + datetime.timedelta(days = 1)
                                self.current_date = datetime.datetime.strftime(date_object, "%m/%d/%Y")
                        match.date = self.current_date

                    match.time = time_value[-7:]
                    print(match.time)
                    print('')
                    index = 0

                    
                    # get team value
                    self.get_team_value(td_tag, match)

                    prev_match_time = time_value[-7:]

                print(cnt_match)

                
                td_tag = td_tag.next_sibling

                # if cnt_match > 15: print(td_tag.contents)

                # get FT HDP value
                self.get_ft_hdp_value(td_tag, match, index)
                td_tag = td_tag.next_sibling

                # get FT O/U value
                self.get_ft_ou_value(td_tag, match, index)
                td_tag = td_tag.next_sibling

                # # get FT 1x2 value
                # self.get_ft_1x2_value(td_tag, match, i)
                # td_tag = td_tag.next_sibling

                # # get 1H HDP value
                # self.get_1h_hdp_value(td_tag, match, i)
                # td_tag = td_tag.next_sibling

                # # get 1H O/U value
                # self.get_1h_ou_value(td_tag, match, i)
                # td_tag = td_tag.next_sibling

                # # get 1H 1x2 value
                # self.get_1h_1x2_value(td_tag, match, i)

                index = index + 1

                print('------------------------')

            if match_flag: 
                matches.append(match)
                cnt_match = cnt_match + 1

            all_match_count = all_match_count+ len(matches)

            if len(matches) > 0:
                league.matches = matches
                self.leagues.append(league)

                print(league_name)

            print('*************************************')
        
        print(len(self.leagues))       
        print(all_match_count)

    def check_match_time(self, prev_match_time, match_time):
        
        if prev_match_time == "":
            time = datetime.datetime.strptime(match_time, "%I:%M%p")
            if self.passed_minutes > time.hour * 60 + time.minute:
                return False
            else: return True

        time_format = "%I:%M%p"

        # Convert the time strings to datetime objects
        time1 = datetime.datetime.strptime(prev_match_time, time_format)
        time2 = datetime.datetime.strptime(match_time, time_format)

        # Cacluate the time difference in minutes
        time_diff = (time2 - time1).total_seconds() / 60

        return time_diff >= 0




    def calculate_days_between_dates(self, date_str1, date_str2):
        # Define the input date format
        input_format = "%m/%d"

        # Convert the date strings to datetime objects
        date_obj1 = datetime.datetime.strptime(date_str1, input_format)
        date_obj2 = datetime.datetime.strptime(date_str2, input_format)

        # Calculate the difference between the two dates
        diff = date_obj2 - date_obj1

        # Extract the number of days from the difference
        days = diff.days

        return days


    def check_time_value(self, time_value):
        if "AM" in time_value:
            return True
        if "PM" in time_value:
            return True
        return False

    def get_data(self):
        return self.leagues
    
    def get_league_name(self, tag):
        return tag.tr.tbody.tr.td.span.string

    def get_time_value(self, tag, match):
        # 
        # get time value
        time_flag = False
        if tag.find('span'):
            time_value = tag.span.text
            time_flag = True

        if time_flag == False:
            return ''

        time_value = "".join(time_value.split())
        
        if "LIVE" in time_value:
            time_value = time_value[4:]
        
        return time_value

    def get_team_value(self, tag, match):
        team_flag = False
        if tag.find('span', {'class' : 'Give'}) or tag.find('span', {'class' : 'Take'}):
            team1_value = tag.tbody.tr.td.span.text
            team2_value = tag.tbody.tr.td.span.next_sibling.next_sibling.text
            print("TEAM: " + team2_value)
            team_flag = True

        if team_flag == False:
            return

        # # print team value
        # print(team1_value + " vs " + team2_value)
        # print('')
        match.match_name = team1_value + " vs " + team2_value

    def get_ft_hdp_value(self, tag, match, i):
        # get pos and neg value
        ft_hdp_value_flag = False
        if tag.find('span', {'class' : 'PosOdds'}) or tag.find('span', {'class' : 'NegOdds'}):
            ft_hdp_pos_odds_value = tag.table.tbody.tr.td.next_sibling.span.string
            ft_hdp_value_flag = True

        if ft_hdp_value_flag == False: return
        
        ft_hdp_neg_odds_value = tag.table.tbody.tr.next_sibling.td.next_sibling.span.string

        # get first value
        ft_hdp_first_flag = True
        ft_hdp_first_value = tag.table.tbody.tr.td.string
        # check if first value exists
        if ft_hdp_first_value is None:
            ft_hdp_first_flag = False

        # get second value
        ft_hdp_second_flag = True
        ft_hdp_second_value = tag.table.tbody.tr.next_sibling.td.string
        if ft_hdp_second_value is None:
            ft_hdp_second_flag = False

        p_value = ""

        # print value
        if ft_hdp_first_flag == True:
            ft_hdp_first_value = "".join(ft_hdp_first_value.split())
            if not ft_hdp_first_value == "":
                p_value = ft_hdp_first_value
            # print(ft_hdp_first_value)
        if ft_hdp_second_flag == True:
            ft_hdp_second_value = "".join(ft_hdp_second_value.split())
            if ft_hdp_second_value != "":
                # print("Second:" + str(len(ft_hdp_second_value)) + ":" + ft_hdp_second_value)
                p_value = "-" + ft_hdp_second_value

        print("HDP: " + ft_hdp_pos_odds_value + '  ' + ft_hdp_neg_odds_value)
        # print('')

        if i == 0:
            match.h1 = ft_hdp_pos_odds_value
            match.hp1 = p_value
            match.a1 = ft_hdp_neg_odds_value
        if i == 1:
            match.h2 = ft_hdp_pos_odds_value
            match.hp2 = p_value
            match.a2 = ft_hdp_neg_odds_value
        if i == 2:
            match.h3 = ft_hdp_pos_odds_value
            match.hp3 = p_value
            match.a3 = ft_hdp_neg_odds_value
        if i == 3:
            match.h4 = ft_hdp_pos_odds_value
            match.hp4 = p_value
            match.a4 = ft_hdp_neg_odds_value

    def get_ft_ou_value(self, tag, match, i):
        ft_ou_flag = False
        # get pos and neg value
        if tag.find('span', {'class' : 'PosOdds'}) or tag.find('span', {'class' : 'NegOdds'}):
            ft_ou_pos_odds_value = tag.table.tbody.tr.find('span').string
            ft_ou_neg_odds_value = tag.table.tbody.tr.next_sibling.find('span').text
            ft_ou_flag = True

        if ft_ou_flag == False:
            return
        
        # get first value
        ft_ou_first_value = tag.table.tbody.tr.td.string

        if i == 0:
            match.o1 = ft_ou_pos_odds_value
            match.op1 = ft_ou_first_value
            match.u1 = ft_ou_neg_odds_value
        if i == 1:
            match.o2 = ft_ou_pos_odds_value
            match.op2 = ft_ou_first_value
            match.u2 = ft_ou_neg_odds_value
        if i == 2:
            match.o3 = ft_ou_pos_odds_value
            match.op3 = ft_ou_first_value
            match.u3 = ft_ou_neg_odds_value
        if i == 3:
            match.o4 = ft_ou_pos_odds_value
            match.op4 = ft_ou_first_value
            match.u4 = ft_ou_neg_odds_value


        # print value
        print(ft_ou_first_value)
        print(ft_ou_pos_odds_value + '  ' + ft_ou_neg_odds_value)

    # def get_ft_1x2_value(self, tag, match, i):
    #     # get 1x2 value
    #     ft_1x2_flag = False

    #     if tag.table.tbody.tr.td.find('span'):
    #         ft_1_value = tag.table.tbody.tr.td.span.string
    #         ft_x_value = tag.table.tbody.tr.next_sibling.td.span.string
    #         ft_2_value = tag.table.tbody.tr.next_sibling.next_sibling.td.span.string
    #         ft_1x2_flag = True

    #     if ft_1x2_flag == False:
    #         return
        
    #     print(ft_1_value + "  " + ft_x_value + "  "  + ft_2_value)

    # def get_1h_hdp_value(self, tag, match, i):
    #     h1_hdp_value_flag = False
    #     # get pos and neg value
    #     if tag.find('span', {'class' : 'PosOdds'}):
    #         h1_hdp_pos_odds_value = tag.find('span', {'class' : 'PosOdds'}).string
    #         h1_hdp_value_flag = True
            
    #     if h1_hdp_value_flag == False:
    #         return

    #     h1_hdp_neg_odds_value = tag.table.tbody.tr.next_sibling.td.next_sibling.span.string
        
    #     # get first value
    #     h1_hdp_first_flag = True
    #     h1_hdp_first_value = tag.table.tbody.tr.td.string
    #     # check if first value exists
    #     if h1_hdp_first_value is None:
    #         h1_hdp_first_flag = False

    #     # get second value
    #     h1_hdp_second_flag = True
    #     h1_hdp_second_value = tag.table.tbody.tr.next_sibling.td.string
    #     if h1_hdp_second_value is None:
    #         h1_hdp_second_flag = False

    #     # print value
    #     if h1_hdp_first_flag == True:
    #         print(h1_hdp_first_value)
    #     if h1_hdp_second_flag == True:
    #         print(h1_hdp_second_value)
    #     print(h1_hdp_pos_odds_value + '  ' + h1_hdp_neg_odds_value)
    #     print('')

    # def get_1h_ou_value(self, tag, match, i):
    #     h1_ou_flag = False
    #     # get pos and neg value
    #     if tag.find('span', {'class' : 'PosOdds'}):
    #         h1_ou_pos_odds_value = tag.table.tbody.tr.td.next_sibling.next_sibling.span.string
    #         h1_ou_neg_odds_value = tag.table.tbody.tr.next_sibling.td.next_sibling.next_sibling.span.string
    #         h1_ou_flag = True

    #     if h1_ou_flag == False: return

    #     # get first value
    #     h1_ou_first_value = tag.table.tbody.tr.td.string

    #     # print value
    #     print(h1_ou_pos_odds_value + '  ' + h1_ou_neg_odds_value)
    #     print('')

    # def get_1h_1x2_value(self, tag, match, i):
    #     # get 1x2 value
    #     h1_1x2_flag = False

    #     if tag.table.tbody.tr.td.find('span'):
    #         h1_1_value = tag.table.tbody.tr.td.span.string
    #         h1_x_value = tag.table.tbody.tr.next_sibling.td.span.string
    #         h1_2_value = tag.table.tbody.tr.next_sibling.next_sibling.td.span.string
    #         h1_1x2_flag = True

    #     if h1_1x2_flag == False:
    #         return
        
    #     print(h1_1_value + "  " + h1_x_value + "  "  + h1_2_value)