from changeworkbook import ChangeWorkbook
from createworkbook import CreateWorkbook

wb = CreateWorkbook('Test.xlsx')
cw = ChangeWorkbook(wb)
cw.test()