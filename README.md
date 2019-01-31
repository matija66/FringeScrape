# FringeScrape

An example of how to use Selenium to scrape a webpage and store the results in a structured Excel spreadsheet.

This code was used to scrape a website for Fringe Festival artists in Perth, WA and get contact information for artists and their shows. This was then collated into an Excel spreadsheet.

The code uses Selenium to loop through pages, then loops through each link in each page. This opens an artist page. The code then extracts the required information into an array for each page.

At the end of the code, this is stored in a dataframe and saved as an Excel spreadsheet.

Note that the Fringe URL needs a username and password. The URL is retained in the browser call, but "Username:Password" is used instead of the actual details. These details are not intended for public display. This code is for reference only.
