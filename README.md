Rotpot Trikotbestellung
===

Installation
---
1. install python 3
2. (optional) create and activate a virtual environment
3. install dependencies by typing `pip install -r requirements.txt`

Use
---
1. configure the correct starting date for the order  
e.g.:  
`ORDER_START_DAY = 1`  
`ORDER_START_MOTH = 9`  
`ORDER_START_YEAR = 2021`  
will drop all orders older than September 1 2021
2. run `python main.py`
3. open the file generated_orders/orderform.xlsx
4. copy the columns B-G from the tab "rotpot_order"
5. paste these columns (without formatting) into the tab "My order" (also B-G)
6. delete the tab "rotpot_order"
7. the orderform is now ready to be sent to force

For examples see completed_orders