# o365getmail

Retrieve Emails from Office365 via MSGraph Interface

_This Python script may be used to retriev emails from Office 365
according to the new OAuth protocol. It's not the perfect solution right now._  
_This script tries to fix malformed Emails by boxing them into a forward mail._  
_Feel free to improve the script and in case you find some better solution, please let me know._

Requirements:

> Python 3  
> see imports at script header

## Getting Started

**Azure Setup**
before using the script you need to aquiee an appication id and a matching
secret and add them to config.py.

How this is done is described in https://github.com/O365/python-o365#authentication

**First run:**  
execute the script with `python o365getmail --auth` to get initial tokens.  
You will get a URL for past an copy to an Browser for Identification. After successful identification copy the returnd URL and past it back.  

**After successful token creation execute:**  
execute the script silent: `./o365getmail`  
or  
execute the script: `./o365getmail --verbose`  
After success use it with **cron**

## Available flags
Short | Long | Explanation
------------ | ------------ | ------------
 | | --version  | 'Print script version.'
-a | --auth     | 'Get initial or refresh token if authentification expired.'
-k | --keep     | 'Keep messages after pushing to MDA.'
-v | --verbose  | 'Output logger infromation to Screen.'
-m | --message  | 'Email message as ''*.eml'' to push to RT.'
-l | --limit    | 'Limit email pull to number of message to be pulled at once.'
