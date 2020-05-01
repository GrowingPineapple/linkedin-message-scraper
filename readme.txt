What does it do:

Requierments:
- Python 3
- openpyxl

Configure:

- login
- password
- Limit: Set the number of pages you ant to parse (e.g. if you want to get the last 140 conversations set the limit to 7 (140/20))


Error:
- in case you get IndexError: list index out of range, set the limit to the amount of recongnized messages (e.g 

1. message.py is for scraping messages a linkedin profile got

2. connections.py is for removing linkedin connections in bulk. Edit "input" by uploaing contacts in the proclaimed format