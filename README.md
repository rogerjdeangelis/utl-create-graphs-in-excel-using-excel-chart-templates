# utl-create-graphs-in-excel-using-excel-chart-templates
Create graphs in excel using excel chart templates
    Create graphs in excel using excel chart templates                                                                                    
                                                                                                                                          
    For output excel interactive pie chart see                                                                                            
    https://tinyurl.com/y5levzf8                                                                                                          
    https://github.com/rogerjdeangelis/utl-create-graphs-in-excel-using-excel-chart-templates/blob/master/chart_pie.xlsx                  
                                                                                                                                          
    github                                                                                                                                
    https://tinyurl.com/y4f8cteu                                                                                                          
    https://github.com/rogerjdeangelis/utl-create-graphs-in-excel-using-excel-chart-templates                                             
                                                                                                                                          
    macros (                                                                                                                              
    https://tinyurl.com/y9nfugth                                                                                                          
    https://github.com/rogerjdeangelis/utl-macros-used-in-many-of-rogerjdeangelis-repositories                                            
                                                                                                                                          
    Source utl_submit_py64_38.sas                                                                                                         
    https://tinyurl.com/y3ztaklb                                                                                                          
    https://www.tutorialspoint.com/python-plotting-different-types-of-style-charts-in-excel-sheet-using-xlsxwriter-module                 
                                                                                                                                          
    IMPORTANT: I updated the drop down to python macro, utl_submit_py64_38, to send standard error                                        
               to %sysfunc(pathname(work))/stderr.txt.                                                                                    
               You get a nasty error without this change.                                                                                 
               Updated macro on end and in github                                                                                         
                                                                                                                                          
    When you manually or programatically change the data in the excel sheet yhr                                                           
    graph will automatically be updated. At this point SAS is not needed.                                                                 
                                                                                                                                          
    Problem: Create an sheet1 with this data and associated excel pie chart                                                               
             Note it is possible to update the background table in an existing worksheet using SAS                                        
                                                                                                                                          
    /*                   _                                                                                                                
    (_)_ __  _ __  _   _| |_                                                                                                              
    | | `_ \| `_ \| | | | __|                                                                                                             
    | | | | | |_) | |_| | |_                                                                                                              
    |_|_| |_| .__/ \__,_|\__|                                                                                                             
            |_|                                                                                                                           
    */                                                                                                                                    
                                                                                                                                          
    %let xlsout=d:/xls/chart_pie.xlsx;                                                                                                    
                                                                                                                                          
    %let data=[['Apple','Cherry','Pecan'],[60,30,10],];                                                                                   
                                                                                                                                          
    /*           _               _                                                                                                        
      ___  _   _| |_ _ __  _   _| |_                                                                                                      
     / _ \| | | | __| `_ \| | | | __|                                                                                                     
    | (_) | |_| | |_| |_) | |_| | |_                                                                                                      
     \___/ \__,_|\__| .__/ \__,_|\__|                                                                                                     
                    |_|                                                                                                                   
    */                                                                                                                                    
                                                                                                                                          
    d:/xls/chart_pie.xlsx                                                                                                                 
                                                                                                                                          
    +---+--------------------------------------------------------------+                                                                  
    |   |     A      |    B       |     C      |    D       |    E     |                                                                  
    +---+--------------------------------------------------------------+                                                                  
    | 1 |   CATEGORY    VALUES                                         |                                                                  
    | 2 |                          Note if you change                  |                                                                  
    | 3 |    Apple        33       the values in the table             |                                                                  
    | 4 |    Cherry       33       the chart will auromatically        |                                                                  
    | 5 |    Pecan        33       be updated. At this point           |                                                                  
    | 6 |                          SAS is not needed.                  |                                                                  
    | 7 |                                                              |                                                                  
    | 8 |                                                              |                                                                  
    | 9 |                      *************                           |                                                                  
    |10 |               *******             ****                       |                                                                  
    |11 |             **                        ***                    |                                                                  
    |12 |           **                             **                  |                                                                  
    |13 |         ** ..                              **                |                                                                  
    |14 |        *     .       __ _ _ __ ___  ___ _ __**               |                                                                  
    |15 |      **        .    / _` | '__/ _ \/ _ \ '_ \ *              |                                                                  
    |16 |      *          .  | (_| | | |  __/  __/ | | | **            |                                                                  
    |17 |     *            .  \__, |_|  \___|\___|_| |_|  *            |                                                                  
    |18 |    **             . |___/                        *           |                                                                  
    |19 |    *                .                            **          |                                                                  
    |20 |   **                 ..                           *          |                                                                  
    |21 |   *               _    .                          **         |                                                                  
    |22 |   *  _ __ ___  __| |     .                         *         |                                                                  
    |23 |   * | '__/ _ \/ _` |      .                        *         |                                                                  
    |24 |   * | | |  __/ (_| |       + . . .. . .. . .. . .. *         |                                                                  
    |25 |   * |_|  \___|\__,_|      .                        *         |                                                                  
    |26 |   **                    ..                         *         |                                                                  
    |27 |    *                    .    _     _              **         |                                                                  
    |28 |    **                  .    | |__ | |_   _  ___    *         |                                                                  
    |29 |     *                 .     | '_ \| | | | |/ _ \  **         |                                                                  
    |30 |      *                .     | |_) | | |_| |  __/  *          |                                                                  
    |31 |       *            ..       |_.__/|_|\__,_|\___|  *          |                                                                  
    |32 |        *          .                              *           |                                                                  
    |33 |         **      .                              **            |                                                                  
    |34 |           **   .                             **              |                                                                  
    |35 |             **.                            **                |                                                                  
    |36 |               ***                      ***                   |                                                                  
    |37 |                   ****             ****                      |                                                                  
    |38 |                       *************                          |                                                                  
    |39 |                                                              |                                                                  
    |40 |                  +-------------------+                       |                                                                  
    |41 |                  |  LEGEND           |                       |                                                                  
    |42 |                  |                   |                       |                                                                  
    |43 |                  |  GREEN = APPLE    |                       |                                                                  
    |44 |                  |                   |                       |                                                                  
    |45 |                  |  RED   = PECAN    |                       |                                                                  
    |46 |                  |                   |                       |                                                                  
    |47 |                  |  BLUE  = Cherry   |                       |                                                                  
    |48 |                  +-------------------+                       |                                                                  
    |   |                                                              |                                                                  
    +---+--------------------------------------------------------------+                                                                  
                                                                                                                                          
    /*                                                                                                                                    
     _ __  _ __ ___   ___ ___  ___ ___                                                                                                    
    | `_ \| `__/ _ \ / __/ _ \/ __/ __|                                                                                                   
    | |_) | | | (_) | (_|  __/\__ \__ \                                                                                                   
    | .__/|_|  \___/ \___\___||___/___/                                                                                                   
    |_|                                                                                                                                   
    */                                                                                                                                    
                                                                                                                                          
    %let xlsout=d:/xls/chart_pie.xlsx;                                                                                                    
    %let data=[['Apple','Cherry','Pecan'],[60,30,10],];                                                                                   
                                                                                                                                          
    %utlfkil(d:/xls/chart_pie.xlsx);                                                                                                      
                                                                                                                                          
    %utl_submit_py64_38("                                                                                                                 
    # import xlsxwriter module                                                                              ;                             
    import xlsxwriter                                                                                       ;                             
    # Workbook() takes one, non-optional, argument which is the filename #that we want to create.           ;                             
    workbook = xlsxwriter.Workbook('&xlsout')                                                               ;                             
    # The workbook object is then used to add new worksheet via the #add_worksheet() method.                ;                             
    worksheet = workbook.add_worksheet()                                                                    ;                             
    # Create a new Format object to formats cells in worksheets using #add_format() method .                ;                             
    # here we create bold format object .                                                                   ;                             
    bold = workbook.add_format({'bold': 1})                                                                 ;                             
    # create a data list .                                                                                  ;                             
    headings = ['Category', 'Values']                                                                       ;                             
    data=&data                                                                                              ;                             
    # Write a row of data starting from 'A1' with bold format.                                              ;                             
    worksheet.write_row('A1', headings, bold)                                                               ;                             
    # Write a column of data starting from A2, B2, C2 respectively.                                         ;                             
    worksheet.write_column('A2', data[0])                                                                   ;                             
    worksheet.write_column('B2', data[1])                                                                   ;                             
    # Create a chart object that can be added to a worksheet using #add_chart() method.                     ;                             
    # here we create a pie chart object .                                                                   ;                             
    chart1 = workbook.add_chart({'type': 'pie'})                                                            ;                             
    # Add a data series to a chart using add_series method.                                                 ;                             
    # Configure the first series.                                                                           ;                             
    #[sheetname, first_row, first_col, last_row, last_col].                                                 ;                             
    chart1.add_series({'name':'Piesalesdata','categories':['Sheet1',1,0,3,0],'values':['Sheet1',1,1,3,1],}) ;                             
    # Add a chart title                                                                                     ;                             
    chart1.set_title({'name': 'Popular Pie Types'})                                                         ;                             
    # Set an Excel chart style. Colors with white outline and shadow.                                       ;                             
    chart1.set_style(10)                                                                                    ;                             
    # Insert the chart into the worksheet(with an offset).                                                  ;                             
    # the top-left corner of a chart is anchored to cell C2.                                                ;                             
    worksheet.insert_chart('C2', chart1, {'x_offset': 25, 'y_offset': 10})                                  ;                             
    # Finally, close the Excel file via the close() method.                                                 ;                             
    workbook.close()                                                                                        ;                             
    ");                                                                                                                                   
                                                                                                                                          
    /*                                                                                                                                    
     _ __ ___   __ _  ___ _ __ ___                                                                                                        
    | `_ ` _ \ / _` |/ __| `__/ _ \                                                                                                       
    | | | | | | (_| | (__| | | (_) |                                                                                                      
    |_| |_| |_|\__,_|\___|_|  \___/                                                                                                       
                                                                                                                                          
    */                                                                                                                                    
                                                                                                                                          
    Here is the new dropdown to Python                                                                                                    
                                                                                                                                          
    Also soon to be in                                                                                                                    
    macros                                                                                                                                
    https://tinyurl.com/y9nfugth                                                                                                          
    https://github.com/rogerjdeangelis/utl-macros-used-in-many-of-rogerjdeangelis-repositories                                            
                                                                                                                                          
    %macro utl_submit_py64_38(                                                                                                            
          pgm                                                                                                                             
         ,return=  /* name for the macro variable from Python */                                                                          
         )/des="Semi colon separated set of python commands - drop down to python";                                                       
                                                                                                                                          
      * write the program to a temporary file;                                                                                            
      filename py_pgm "%sysfunc(pathname(work))/py_pgm.py" lrecl=32766 recfm=v;                                                           
      data _null_;                                                                                                                        
        length pgm  $32755 cmd $1024;                                                                                                     
        file py_pgm ;                                                                                                                     
        pgm=&pgm;                                                                                                                         
        semi=countc(pgm,";");                                                                                                             
          do idx=1 to semi;                                                                                                               
            cmd=cats(scan(pgm,idx,";"));                                                                                                  
            if cmd=:". " then                                                                                                             
               cmd=trim(substr(cmd,2));                                                                                                   
             put cmd $char384.;                                                                                                           
             putlog cmd $char384.;                                                                                                        
          end;                                                                                                                            
      run;quit;                                                                                                                           
      %let _loc=%sysfunc(pathname(py_pgm));                                                                                               
      %let _stderr=%sysfunc(pathname(work))/stderr.txt;                                                                                   
      %let _stdout=%sysfunc(pathname(work))/stdout.txt;                                                                                   
      filename rut pipe  "c:\Python38\python.exe &_loc 2> &_stderr";                                                                      
      data _null_;                                                                                                                        
        file print;                                                                                                                       
        infile rut;                                                                                                                       
        input;                                                                                                                            
        put _infile_;                                                                                                                     
      run;                                                                                                                                
      filename rut clear;                                                                                                                 
      filename py_pgm clear;                                                                                                              
                                                                                                                                          
      * use the clipboard to create macro variable;                                                                                       
      %if "&return" ^= "" %then %do;                                                                                                      
        filename clp clipbrd ;                                                                                                            
        data _null_;                                                                                                                      
         length txt $200;                                                                                                                 
         infile clp;                                                                                                                      
         input;                                                                                                                           
         putlog "*******  " _infile_;                                                                                                     
         call symputx("&return",_infile_,"G");                                                                                            
        run;quit;                                                                                                                         
      %end;                                                                                                                               
                                                                                                                                          
    %mend utl_submit_py64_38;                                                                                                             
                                                                                                                                          
