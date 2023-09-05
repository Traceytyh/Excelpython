# Excelpython
Using Openpyxl library in python to extract data from excel files

TestCode folder shows the various functions being tested separately. The functions are found in the table below.

The final code makes use of some of the functions, where users are able to type in the respective instructions to extract their required data.
Inputs:
1)	Directory of file for data extraction
2)	Parameters to extract
3)	If repeated parameters within a sheet represents a different set of data
4)	If matching of certain data is needed (arranging data by dates), to state which parameters are to be compared

The code has to identify:
1.	Identify if the data is written horizontally or vertically
2.	If there is a repetition of parameters within a sheet
3.	Function:
    1. Make data unique, by removing repeated values or obtaining an average
    2. Compare a reference parameter to get data for the same sampling point
    3. Compile average for a same date
    4. Extract other data for the same sampling point


*Usage of the function, `Special cases flagged to user, ~Assumptions made
|Function| 	What it does |
|--------|---------------|
|find_rc_no(workbook, sheet_name, search_value)	|Searches for a value and returns its coordinates as a string<br>*Used to find headers. Outputs a list in case the same header is repeated in the same sheet|
|Get_data(workbook, sheet_name, search_value)|	Takes a header value and returns the entire row/column of the data with its header and whether the data is horizontal or vertical<br>~Assumes that the headers are along the borders of the excel to determine if it is vertical or horizontal<br>`If header is found at A1, it asks the user if the data is horizontal or vertical<br>*Used to get data under a header. If header repeated in the same sheet, outputs the data as a list: <br>len(output) = number of times header is repeated, len(output[i]) = number of data under that header <br>*Hori_vert is used as a variable to denote whether to extract subsequent data in a vertical or horizontal manner (vertical means that the data from the same sampling point is in the same row)|
|Get_input(output,prompt)	|Asks user for input until the user typed ‘done’ and return user’s inputs as a string in a list<br>*Used in obtaining multiple inputs from the user|
|Get_date(date)	|Identify if the input date is a string or datetime and returns [day, month, year]<br>*Allows date data collected in different formats to be compared (DD/MM/YYYY, YYYY/MM/DD etc)|
|Datetime_to_str(date)	|Convert datetime into string (However, .strftime("%d/%m/%Y, %H:%M:%S" could be used instead)<br>*Used to type in date value as datetime types cant be appended into excel|
|Merge_same_index(list0,list1)| Merges the elements with same index of both lists, and returning that list. If the lists are of different lengths, None is being inserted. <br>*Used when merging different parameter values together as they need to be in the same sub-list to be appended into the same row in excel<br>~Assume that the list0[i] and list1[i] lengths are constant|
|Data_from_cell(cell_list)|	Extracts the values of the cells in the list and returns the values in a list|
|Data_with_cell(cell_list)|	Extracts the value in the cell and puts it in the list with the cell (eg. [<Cell 'Sheet1'.G1>, <Cell 'Sheet1'.G2>] -> [[<Cell 'Sheet1'.G1>, ‘G1 value’], [<Cell 'Sheet1'.G2>,'G2 value']]<br>*Used in “Make_unique_param” functions which needs to refer to the value and also retain the cell is taken from. |
|Make_unique_param_one_action(workbook, sheet, ref_param)|	Inputs parameter header (can be a string or a list of headers) and extracts all the data under it. Then asks the user if they want the former/latter of the repeated values are to be kept or for an average to be taken. This action will be executed for all the repeated data found under that header. Returns the list of cells corresponding to the header with unique values (includes the header cell), and another list containing the cells to be used for averaging.<br>*Used before “compare_val” functions. In order to obtain data for the same sample point, these sample points have to be distinct.| 
|Make_unique_param_detailed(workbook, sheet, ref_param)|	Similar to “Male_unique_param_one_action” but it asks the users every time there is a repeated value, so that if eg. Out of the 4 repeated data, the user can specify the 3rd repeated data to be kept. They can also choose to find average of specific repeated values. Returns the list of cells corresponding to the header with unique values (includes the header cell) and another list containing the cells to be used for averaging.<br>*Used before “compare_val” functions too|
|Compare_val_same_sheet(ref_list)|	Inputs lists of data to be compared [[<Cell 'Sheet1'.A1>, <Cell 'Sheet1'.A2>…], [<Cell 'Sheet1'.C1>, <Cell 'Sheet1'.C2>….], [<Cell 'Sheet1'.E1>, <Cell 'Sheet1'.E2>…]]. The first list is used as a reference. After A and C are compared, the resultant compared columns are used as the new reference and so on. Returns the compared list (includes the header cell). <br>len(output[i]) = number of common data<br>len(output) = number of lists compared<br>Output eg, [[<Cell 'Sheet1'.A1>, < Cell 'Sheet1'.C1>, <Cell 'Sheet1'.E1>], [<Cell 'Sheet1'.A2>, < Cell 'Sheet1'.C2>, <Cell 'Sheet1'.E4>]…]<br>`User will be asked if the reference data is a date, which will use the get_date function in case the formatting of the dates during comparison is different|
|Compare_val(ref_param, ref_workbook, ref_sheet, workbook, sheet)|	Takes a parameter header and searches for it in two workbooks to be cross-referenced. The data under this common header are compared and if they are the same, it will be kept. Returns the common cells in a list ([<Cell 'Sheet1'.A1>, <Cell 'Sheet2'.A1>], [<Cell 'Sheet1'.A2>, <Cell 'Sheet2'.A3>],…] where <Cell 'Sheet1'.A1>.value =  <Cell 'Sheet2'.A1>.value and <Cell 'Sheet1'.A2>.value = <Cell 'Sheet2'.A3>.value. <br>*Used to get data in two sheets with the same reference data under the given header, the reference data is usually a date<br>`User will be asked if the reference data is a date, which will use the get_date function in case the formatting of the dates during comparison is different|
|compare_val_multi(ref_param, wb_sheet_list)|	‘Compare_val’ only allows the comparison of two sheets, hence it was expanded to this, where wb_sheet_list = [[wb1, s1], [wb2, s2], ...]. Functions the same way as “compare_val”|
|Ext_for_same_pt(ref_list, workbook, sheet, param_name_list)|	Inputs a list of data excluding the header (eg. [<Cell 'Sheet1'.A2>, <Cell 'Sheet1'.A3>,..]), as well as a list of parameter headers. Having obtained the data under the parameter headers, according to Hori_vert (if the data is stored horizontally or vertically), the corresponding parameter at the same sampling point in ref_list is taken. Returns the ref value and parameter values in a list. (If it is a datetime, it will be read as "%d/%m/%Y, %H:%M:%S"). Eg [['02/02/2024, 00:00:00', 2, 20],['02/02/2024, 00:00:00', 3, 21]..]<br>Len(output) = number of sample points<br>Len(output[i]) = number of parameters in param_name_list + 1<br>*Used to append extracted values into an output file with data stored vertically|
|get_average(average_list, workbook, sheet, param_name_list)|	Inputs the average list from “make_unique” functions, uses ‘ext_for_same_pt’ to get the values. If the reference data is the same, the parameter value will be added, and averaged after the reference data changes. The reference data and the averaged data will be stored. Subsequent parameter averages will be appended accordingly. Returns the ref value and averaged parameter values in a list.<br>Len(output) = number of unique sampling points <br>Len(output[i]) = number of parameters + 1|
|write_ave(datasheet, date_col)|^Used for caustic extraction	*Instead of using ‘make_unique’, this function stores the data in the next column if the date is the same and calculates the average once there is a change in date. Returns a list containing the [[date1, average1], [date2, average2]…]<br>~Assumes that the dates are in a chronological manner. If the repeated dates are not sequential, they would be considered as different dates. <br>~Assumes no None data|
|write_ave_gen(datasheet, ref_col)|	Similar to write_ave, but is not specific to having date as the reference.  |
|Separate_list(input_list)|	Rearrange elements by their index. Eg <br>input_list = [[A1, B1, C2], [A2, B2, C3]...]<br>output = [[A1, A2..], [B1, B2..], [C2, C3..]]<br>*Used after compare_val so that the cells with same sheet are in a list together, for easy extraction|
|keep_if_specific_val (data_list, param_index_list, value_list)|	Inputs the index of the data_list reflecting the parameter that has a required value. Returns the data_list without those sample points that do not meet this required value. 
|save_file(values)	|Saving the data into a new file or an existing file|

These functions are primarily built using for loops that goes through the elements of a list. In cases like ‘compare_val’, ‘break’ function is used to minimise the computational power. 

