Stock Analysis with VBA
=======================

Overview of Project
-------------------

This project analyzed the historical stock returns of 12 companies
between 2017 and 2018. The data being used returned both trading volumes
and gains/losses for each stock.

### Results of Analysis

Comparing the two years of data (2017 and 2018), broad based it looks
like 2017 had almost universally positive year for each stock, versus a
complete opposite downward trend in 2018, except for one company. While
trading volume did fluctuate signficantly between years for an
individual stock, on a whole the amount of trading down year over year
did not seem to change much. There also did not seem to be much
correlation between trading volume and the increase/decrease in the
return for that stock in each year.

The big difference in execution times between my original script and the
refactored is I did not include the formatting and conditional
formatting in the same macro, I used a separate macro to perform this
function. Below is the code that probably made the biggest difference in
time.

    Worksheets("All Stock Analysis").Activate
        Range("A3:C3").Font.FontStyle = "Bold"
        Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
        Range("B4:B15").NumberFormat = "#,##0"
        Range("C4:C15").NumberFormat = "0.0%"
        Columns("B").AutoFit

        dataRowStart = 4
        dataRowEnd = 15

        For i = dataRowStart To dataRowEnd
            
            If Cells(i, 3) > 0 Then
                
                Cells(i, 3).Interior.Color = vbGreen
                
            Else
            
                Cells(i, 3).Interior.Color = vbRed
                
            End If
            
        Next i

### Summary

        - What are the advantages or disadvantages of refactoring code?

        The advantages of refactoring are fairly obvious, as starting with code allows you to have a framework in place and focus more on making the code better.  Starting from scratch can be a much more time intensive task.  However, working with unfamiliar variables or differences in coding styles can create confusion and unforseen road blocks.  Making sure variables were named properly or placed properly within the code became difficult during refactoring.

    - How do these pros and cons apply to refactoring the original VBA script?

        Personally I enjoyed and understood concepts more using original VBA script.  Refactoring code written by someone else became confusing to me, and small details became big issues in determining exactly the right order of operations.  In particular, I had issues with the totaling volumes and start and end prices for the subsequent 'i' after the first.  The main issue I had was making sure to go back and activate the worksheet, which I thought had already been done in previous code.  Keeping track of all while refactoring code became significantly more difficult for me than writing it from scratch.
