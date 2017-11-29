# bed365_scraper

For scraping [bed365](https://www.bet365.com/)

## Instration

You need Python 3 (or 2) and a bunch of modules for scraping.

I highly recommend that you use an all-in-one Python distribution like Anaconda. This distribution comes with an excellent package manager named conda. It lets you install easily many modules on most platforms (Windows, Linux, Mac OS X), in 64-bit (recommended if you have a 64-bit OS) or 32-bit.

## Requirements
Install Selenium
```
pip install selenium
```

## Cloning the repository

You need git, a distributed versioning system, to download a local copy of this repository. Open a terminal and type:

git clone https://github.com/nkimoto/bed365_scraper
This will copy the repository in a local folder named bed365_scraper.

## Usage
```
python bed365_scraper.py -t 15 
```

- `-t`.`--time`  
You can specify waiting time between pagings. Default is 15.
