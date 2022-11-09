from datetime import datetime
from datetime import timedelta

today = datetime.today().strftime('%d_%m_%y')
print(today)

print('Metabase_'+today)

yesterday = (datetime.now() - timedelta(1)).strftime('%d_%b_%y')
print('_' + yesterday)