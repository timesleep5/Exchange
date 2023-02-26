from changeworkbook import ChangeWorkbook
from createworkbook import CreateWorkbook

#create a new, blank workbook
crwb = CreateWorkbook('FlowOpt.xlsx')

#change the workbook, fill in
wbname = crwb.get_workbook()
chwb = ChangeWorkbook(wbname)
chwb.update_schedule_dates()





#buwb.update_history()          #moves schedule to history
#chwb.clear_schedule()          #clears schedule
#chwb.update_schedule_dates()   #updates dates in schedule