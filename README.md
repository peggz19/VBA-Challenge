# VBA-Challenge

Hello and Welcome to Peggy Tadi's VBA Challenge.

The VBA Script.txt contains the main macro, VBAChallenge, and, the Delete_and_Start_Over macro that I ran each time I needed to clear  all the sheets and rerun my main macro. However the VBA Script.vbs only contains the main macro, VBAChallenge, on which we are being tested.

Reference of the formulas in the script that we never learned in class:
ws.Range("M1:W1").Columns.AutoFit https://learn.microsoft.com/en-us/office/vba/api/excel.range.autofit

ws.Range("M1:W1").Font.Bold = True https://learn.microsoft.com/en-us/office/vba/api/excel.font.bold

 Set rngO = ws.Range("O:O") https://www.excelanytime.com/excel/index.php?option=com_content&view=article&id=105:find-smallest-and-largest-value-in-range-with-vba-excel&catid=79&Itemid=475
 
 Application.WorksheetFunction.Max(rngO) I knew how to do that from prior experience recording a Macro

Please note that because of the large size of Multiple Year Stock Data, my computer kept lagging. Therefore, if it looks like my file is full of mistakes, please refer to alphabetical_testing excel file where everything worked perfectly.

Note* I changed the orginal data of ABB in alphabetical_testing to what we have in Multiple Year Stock Data in order to compare my results with the provided images.

THANK YOU THANK YOU
