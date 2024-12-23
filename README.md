# VBA-challenge

Attribute VB_Name = "Module2_Challenge"
###

Sub stock_analysis()

  'Create a variable for the worksheet and others
    Dim ws As Worksheet
    Dim i As an Integer
    

    'Loop through Sheets 1 to 6 for ThisWorkBook
    
     ' Activate the worksheet
     
      ' Set title row

       ' Set initial values
           
       ' Get the row number of the last row with data

        'Loop through each row
            For i = 2 To rowCount
            
        ' If ticker changes then print results
            
        ' Stores results in variables
           
       ' Handle zero total volume
                   
       ' Print the results
                     
       ' Find first non-zero starting value                

       ' Calculate change
                       
       ' Start of the next stock ticker
                       
      ' Print the results
                      
      ' Colors positives green and negatives red                      

      ' Reset variables for new stock ticker
                 
      ' If ticker is still the same add results
               
       ' Take the max and min and place them in a separate part in the worksheet
          
        ' Returns one less because header row not a factor
            
         ' Final ticker symbol for total, greatest % of increase and decrease, and average
            
    Next ws

End Sub
