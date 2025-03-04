# Use VBA scripting to analyze generated stock market data

Output for challenge module 2 (VBA) - UoT

Use VBA scripting to analyze generated stock market data

# This code includes 2 subs

# IMPORTANT : User must launch sub (Allsheets)

#   A) First one (Allsheets) will select every Worksheet in the Workbook, and will apply (call) the second sub
#        * the code can be used whether the workbook has 1 or N sheets
#    B) the second sub (vbachallenge) will perform all the needed transformation and calculation
#

# sub vbachallenge works as follows :
    # 1) all the processing is done in columns AA to AN
    # 2) the needed columns are copy/pasted  in columns I to Q at the end
#
# the processing works as follows

    # 1) on the basis of a distinct extraction of ticker value (using xls formula)
    # 2) we identify for each ticker the first and last date of trading
    # 3) we retrieve for each ticker the price of OPENING and CLOSING related to the dates identified earlier
    # 4) we calculate the price evolution during the period and the %
    # 5) we sum the column of stocks traded during the periode (using sumifs xls formula)
    # 6) we apply a conditional format on price evolution and %
#
# steps 2 to 6 are done by looping and use of XLS formulas
#
    # 7) we create the new table : Greatest % Increase, Greatest % Decrease, Greatest Total Volume
    # 8) we calculate the max and min of price % evolution
    # 9) we retrieve the ticker related using xls VLOOKUP formula
    #
    # 10) we copy past (value and format) the needed columns fo the challenge to columns I to Q
    # 11) all the columns used for processing are deleted at the end

