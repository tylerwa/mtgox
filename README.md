MTGOX History for Excel 
=====
-----

Excel VBA macro that processes and intelligently combines the USD history file and BTC history text files into a single history.

To see how it works download the two history files and workbook below.
* [History (BTC) CSV file](http://tylerwames.com/misc/history_BTC.csv)
* [History (USD) CSV file](http://tylerwames.com/misc/history_BTC.csv)
* [Basick Workbook with macro](http://tylerwames.com/misc/Mtgox.xlsm)

Once downloaded, make sure they are in the same directory and open the workbook.  In column K you will see a button and after you click it the macro will import all the data from the two text files.

To illustrate what you can do with the history in the new format download the following workbook and place it in the same folder as the other files.  When you click the button (column AB) it will import the data as normal but now will show the transaction effects on your balance!

* [Workbook with extras](http://tylerwames.com/misc/Mtgox with extras.xlsm)

###Screenshots:
-------

**Raw History USD:**

![Raw USD History from MtGox](http://www.tylerwames.com/misc/History USD.png)


**Raw History BTC:**

![Raw BTC History from MtGox](http://www.tylerwames.com/misc/History BTC.png)


**Result (green) with example (yellow) of how you can use the combined history**

![Screenshot of result (with extras)](http://www.tylerwames.com/misc/Mtgox Screenshot.png)

### How it works ###
-----

1. It takes each CSV file and imports each line by line into an array. During this step:
*Each line is split up into a sub array based on the delimiter (comma).
*The elements of the line are parsed into the 8 or so components of each transaction.
* The trade ID, rate, and fee % are pulled from the description element
* The trade ID is compared to a running list of unique trade IDs and added if not present.
2. Once both of those are done (the BTC and USD history csv files) there are 3 arrays: BTC history, USD history, and unique transaction IDs.
3. Now the two history arrays are aggregated based on the unique transaction ids.
* I use the list of unique transaction ID, compiled while importing the CSV files, as the key for combining the line items together from each file.
* For instance, I would exclude the out from the BTC side but calculate the BTC effect based on the rate, and USD spent.
* The deposits and withdrawals do not have unique transaction ID's originally so I had to improvise and used the timestamp along with "deposit @" or "withdrawn @" as their unique ID.
4. So after this is finished, I have the list of all the orders and then I sort the list based on the date (descending was my preference).
5. Next, since the list is descending and therefore flows from bottom to top, I swap the transaction line with the transaction's fee line.
6. Now it finally prints it out on the spreadsheet "all".

Note: Not all types of transactions are handled because I haven't actually used them before so some may not show up correctly if I haven't encountered them.

Special handling example: Dwolla deposits now have a hard return in the middle of the description for some reason (no other transaction has this) so results in a line split up into two in the CSV file so that is handled with the following:
* The first line of each CSV file contains headings and establishes the number of columns.
* Each line is imported and split into an array like normal but then if the number of elements of that line is less than the total established by the first line it assumes this to be a split line. The next line is combined with the current.
