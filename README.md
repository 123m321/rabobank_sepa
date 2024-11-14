# rabobank_sepa
Tool to create a batch payment for Rabobank business accounts from an excel sheet

# Introduction
Sometimes you need to make a lot of payments with a business account. For examnple if you run a local football club, and you need to pay
a few dozen volunteers a fee. With this tool you don't need to do a lot of extra work, an excel sheet with the columns **naam**, **bedrag**, **iban** and **omschrijving** does the trick.

# Making it work
Download the files, go in the terminal (or command) to the folder.
First you need to import the python modules:
```
pip install -r requirements.txt
```

Launch the python file
```
python rabobank_sepa_maker.py
```

And add your own iban and name. Then you can **start** with importing an excel sheet with the information. The toprow needs to contain the four column headers.

# In the rabobank environment
Log in to the Rabobank account on your computer, import the xml file saved in the */data/** file and approve the payment. You can also verify first if the payments make sense to you.

# your own risk
This tool saved a lot of time for me. But as always, use it with your own risk.
