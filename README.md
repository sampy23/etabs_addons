# ETABS Applications
Aim is to develop python programs that interact with ETABS using CSI OAPI
## delta_ns add in
This application interacts with ETABS using python to calculate delta_ns. It assumes C_m conservatively as 1. This is to due to ambiguity in calculation of C_m in ETABS. This assumption avoids any chance of missing out on any columns. But this also means some amount of columns will be flagged wrongly.
## Future project
* Develop an application that selects column with PMM value greater than a user specified value
* Select members with a particular loading value
* Transfer load combination and load cases from one file to another
