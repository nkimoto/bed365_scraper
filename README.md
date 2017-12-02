# bed365_scraper(dev)

For scraping [bed365](https://www.bet365.com/)

## Instration

### Python
You need Python 3 (or 2) and a bunch of modules for scraping.

I highly recommend that you use an all-in-one Python distribution like Anaconda. This distribution comes with an excellent package manager named conda. It lets you install easily many modules on most platforms (Windows, Linux, Mac OS X), in 64-bit (recommended if you have a 64-bit OS) or 32-bit.


### Selenium module
Install Selenium
```
pip install selenium
```


### Chrome
1. make ` /etc/yum.repos.d/google-chrome.repo` . You need sudo authorization.

        [google-chrome]
        name=google-chrome
        baseurl=http://dl.google.com/linux/chrome/rpm/stable/$basearch
        enabled=1
        gpgcheck=1
        gpgkey=https://dl-ssl.google.com/linux/linux_signing_key.pu

1. You can get chrome binary with under command.

        sudo yum install -y google-chrome-unstable libOSMesa google-noto-cjk-fonts

1. The newest version 2.33 can get from https://sites.google.com/a/chromium.org/chromedriver/downloads 

        wget https://chromedriver.storage.googleapis.com/2.33/chromedriver_linux64.zip
        unzip chromedriver_linux64.zip
        sudo mv chromedriver /usr/local/bin/

1. Change authorization

        sudo chown root:root /usr/local/bin/chromedriver

1. Install libraries

        sudo yum install -y libX11 GConf2 fontconfig




## Cloning the repository

You need git, a distributed versioning system, to download a local copy of this repository. Open a terminal and type:
```
git clone https://github.com/nkimoto/bed365_scraper
```
This will copy the repository in a local folder named bed365_scraper.

## Usage
example
```
python bed365_scraper.py -t 15 
```

- `-t`, `--time`  
You can specify waiting time between pagings. Default is 15.

- `-mn`, `--match_num`  
You can specify scraping match number. Default is 3.

- `-rc`, `--roop_count`  
You can scraping roop count. Default is 10.
