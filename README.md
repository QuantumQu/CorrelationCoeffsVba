# CorrelationCoeffsVba

Implementation of Spearman and Kendall correlation coefficient for MS Excel (VBA)

The BAS file contains three functions - getRanks, spearmanCorrel and kendallCorrel.

The first one uses bubble sort for sorting input data arrays according to observed values.
After that ranks are added to each value and finally the ranks are sorted according to indexes
of original observation. This leads to assigning of rank to each observation.

Meaning of spearmanCorrel and kendallCorrel functions is self-explanatory.
Once BAS file is added to VBA project in Excel (it is recommended to add them to PERSONAL workbook),
user can call functions spearmanCorrel and kendallCorrel from a sheet.
A user works with the functions in the same manner as with Excel-native function CORELL implementing
calcualtion of Pearson correlation coefficient.
