# ETABS Applications
Aim is to develop python programs that interact with ETABS using CSI OAPI
## delta_ns add in
This application interacts with ETABS using python to calculate delta_ns. 
### How to
 0. Running the code launches a window which allows user to choose the threshold value of del_ns below which will be ignored
 1. It also gives you two more options; Fast & Slow
 2. Fast option calculates delta_ns for only the combination for which PMM ratio is the maximum, hence faster
 3. Slow option calulates delta_ns for all the available combination, hence much slower. 
 4. However slow option allows the user to narrow down the targeting load combinations which starts with specific letters and also by ignoring combo ending with specific letters
 5. The output of the program can be inspected by user by checking DEL_NS.xlsx file created in root directory of ETABS.
### Issues
0. Calculation of Cm for combos with max & min is not always accurate as API for "Column design forces" table is not availabe
1. Beta_dns assumed as 1
2. Code duplication under class methods del_ns_fast & del_ns_slow
## Future project
* Develop an application that selects column with PMM value greater than a user specified value
* Select members with a particular loading value
* Transfer load combination and load cases from one file to another
