
# Web Scraping Bot for Tech Product Prices

## Project Description
This project is a web scraping bot that scrapes data from technology websites that sell tech products. Specifically, it scrapes price data of products and stores them in an Excel spreadsheet. Predefined formulas in the spreadsheet are used to calculate the profit margin.

## Installation
To install the necessary dependencies, run the following command:
```bash
pip3 install -r requirements.txt
```
The specific versions of the dependencies are listed in the requirements.txt file.

## Pre-requisites
The user needs to have a WhatsApp account for the messaging bot to function. The user should select the group where the bot will send messages. Please note that using APIs to access WhatsApp is not recommended as it can lead to account bans.

## Configuration
Before running the bot, the user needs to ensure that the correct XPaths are used in the code to locate the elements on the web pages. The XPaths are used in the Selenium WebDriver to interact with the elements on the web pages.

## Execution
To run the bot, simply execute the main Python script. The bot will start scraping the websites and send the scraped data excel and the format text to send to WhatsApp group.

## Troubleshooting
Common errors that may occur during the execution of the bot are related to XPaths. If the bot fails to locate an element on a web page, check the XPath used to locate that element. The structure of web pages can change over time, so the XPaths may need to be updated.

## Author
👤 Joao Melo

- Github: [@johnmelodev](https://github.com/johnmelodev)
- LinkedIn: [@joao-melo-dev](https://www.linkedin.com/in/joao-melo-dev)

## Show Your Support
Give a ⭐️ if this project helped you!
```
