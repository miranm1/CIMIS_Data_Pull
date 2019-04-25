# CIMIS_Data_Pull
# How to use:
This program uses the API from the California Irrigation Mangement Information System to gather weather station data.  It specifically gathers daily precipitation and reference et values for a certain time period. To use the program just simply edit the start and end date located at "B2" and "B3". Now just press "Read CIMIS" button and the values will shortly post going down columns starting at row 8.  It will display the precipitation/reference et values and date right next to the corresponding value.

![alt_text](https://github.com/miranm1/CIMIS_Web_Scrapper/blob/master/cimis1.PNG)
![alt text](https://github.com/miranm1/CIMIS_Web_Scrapper/blob/master/cimis2.PNG)
![alt text](https://github.com/miranm1/CIMIS_Web_Scrapper/blob/master/cimis3.PNG)

# How it works:
The spreadsheets holds locations for data to be used in building the URL for the CIMIS API. This is the start data and end date and the weather station ID. These are located at "B2","B3", and going down row 6 and is incremented by 2 to skip the column used for date. When building the URL for the website API you will need a API that can be obtained after making an account at ( https://cimis.water.ca.gov/ ). Data being coming from the website will be in JSON format and will need to be parsed through. I used the JsonConverter that can be found here ( https://github.com/VBA-tools/VBA-JSON ). The data is stored in an array so you will need to use a nested for each statement to grab the data out. In this program I created two for each statements one is used to count the amount of items going to be read in. The second one is used to store the values in another array to be printed into the spreadsheet. When this column is done it will move to the next column. It will continue until there is no more data found in the row 6.
