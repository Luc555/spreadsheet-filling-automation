from datetime import datetime, timedelta
import calendar
from datetime import date


yesterdayName = date.today() - timedelta(1)
todayName = date.today() 
presentday = datetime.now() 
yesterday = presentday - timedelta(1) 
twoDaysAgo = presentday - timedelta(2) 
threeDaysAgo = presentday - timedelta(3) 
tomorrow = presentday + timedelta(1) 

threeDaysAgoNameShow = calendar.day_name[threeDaysAgo.weekday()]
twoDaysAgoNameShow = calendar.day_name[twoDaysAgo.weekday()]
todayNameShow = calendar.day_name[todayName.weekday()]


threeDaysAgo = threeDaysAgo.strftime('%d/%m/%Y')
twoDaysAgo = twoDaysAgo.strftime('%d/%m/%Y')
ontem = yesterday.strftime('%d/%m/%Y')
hoje = presentday.strftime('%d/%m/%Y')

print(threeDaysAgo)
print(threeDaysAgoNameShow)
print(twoDaysAgo)
print(twoDaysAgoNameShow)
print(threeDaysAgo)
print(ontem)
print(hoje)