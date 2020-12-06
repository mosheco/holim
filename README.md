# holim

Setting up an environment:
1. virtualenv env -q
2. source env/bin/activate
3. pip install google-api-python-client
4. pip install google_auth_oauthlib

## Running payment report.
(It assumes the balances sheet is named 'נוכחי')

python payment_report.py

## Stats

python pstats.py -h  (for help)

E.g. python pstats.py  -c -y 2020
