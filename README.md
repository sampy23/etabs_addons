# ETABS Applications
Aim is to develop python programs that interact with ETABS using CSI OAPI
## delta_ns add in
This application interacts with ETABS using python to calculate delta_ns. 
# So far
1. Calculates del_ns only for critical PMM combos which is named as "fast"
2. Calculates del_ns for all combo which can be grouped by user by specifying first few letters to incluse and last few letters to be ignored. This is named as "slow"
3. Exports output to etabs root folder as excel sheet
# Issues
1. Calculation of Cm for combos with max & min is not always accurate as API for "Column design forces" table is not availabe
2. Beta_dns assumed as 1
3. Code duplication under class methods del_ns_fast & del_ns_slow
## Future project
* Develop an application that selects column with PMM value greater than a user specified value
* Select members with a particular loading value
* Transfer load combination and load cases from one file to another
