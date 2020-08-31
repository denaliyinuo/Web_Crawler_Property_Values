# Anchorage Property Web Scraper ReadMe

# Overview

The purpose of this web scraper is to collect residential property information from the Municipality of Anchorage website.   The web scraper was written in Python and uses the libraries Selenium for web scrapping and Openpyxl for collecting the data in Excel.

# Links

[My Kaggle Notebook of Anchorage Single Family Home Anaylsis](https://www.kaggle.com/nathanoliver/anchorage-single-family-home-anaylsis)

[Anchorage Municipality Public Inquiry Website (site to be web scrapped)](https://www.muni.org/pw/public.html)

# Anchorage Municipality Public Inquiry Home Page

Below is an image of the home page to search for any property in Anchorage.  There are several different search options for the user to find the property in question, such as street address, name, and description.  For this web scraper, I took advantage of the parcel number option. Since each property has an associated number, the program can cycle through all of the properties in the Anchorage area, while ensuring that no properties have been missed.

![image 1](/images/image1.jpeg)

# Web Scraper Steps
# Step 1

The web scraper enters the parcel number, as shown in the red box.  In order to start from the lowest numbered property, it is recommended to input 000-00-000-000.  The search page results will show the lowest numbered property.  After the parcel number has been entered, the program will press the search button.

![image 2](/images/image2.jpeg)

# Step 2

The search results present different properties, showing the parcel number, address, and owner's name or names.  The program will select the link in the first row, shown in the red box.  The program will also copy the parcel number in the second row, shown in the blue box.  This parcel number is copied because it will be saved and used in the succeeding property search.   

![image 3](/images/image3.jpeg)

# Step 3

All of the information shown in blue boxes is copied and stored in an Excel file.  The information copied includes property prices, address, house and lot sizes, and other pertinent information.  The New Search link shown in the red box is then clicked, which brings the browser back to the home page shown in step 1, allowing the program to continue collecting information.

![image 4](/images/image4.jpeg)

