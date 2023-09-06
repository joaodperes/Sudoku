# Sudoku

This script creates a daily Sudoku puzzle and sends it (and it's resolution) to a list of e-mail adressess, configurable via Contacts_Sudoku.xlsx:

Name | E-mail Address
-----|--------------
Name1| Address1

The output is stored as .txt files, both the puzzle and the resolution, and are sent attached to the e-mail.

This uses the module 'yagmail' and requires a Google API.