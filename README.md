# utl_fix_excel_column_names_before_import
Use Microsoft SQL query with SAS passthru to Excel. Use SQL Query to rename columns before importing into SAS. Excel SQL dialect.

    ```  SAS-Forum: Programatic renaming of bad column names on the excel side using passthru  ```
    ```    ```
    ```    I fix the names on the excel side, using passthru and the MS SQL excel  ```
    ```    dialect. The excel SQL dialect is close to MS Access SQL dialect. I have not found  ```
    ```    good documentaion of the excel SQL dialect. (see access doc at end of post.  ```
    ```    ```
    ```    Best with the version of SAS that supports the classic editor.  ```
    ```    ```
    ```    "I would be greatefull if anyone could help me with ':' in variable name.  ```
    ```    When I import from excel files I ve got var names like 'Data var: id'n  ```
    ```    while I can manage(drop, keep, rename) with names without ':' exemple  'Data var name'n,  ```
    ```    I can`t do anything with first one."  ```
    ```    ```
    ```     WORKING CODE  ```
    ```     ============  ```
    ```    ```
    ```           * get bad names usin excel SQL and passthru  ```
    ```           from connection to Excel  (header =no)  ```
    ```           (  ```
    ```            Select  ```
    ```                top 1 *  ```
    ```            from  ```
    ```               [class$]  ** could be {sheet1$]  ```
    ```    ```
    ```            proc transpose data=names out=namxpo;  ```
    ```              var _all_;  ```
    ```    ```
    ```            _NAME_   COL1  ```
    ```    ```
    ```              F1     Name:-+@.$%^  ```
    ```              F2     -Sex-:@UHS+%^&*()  ```
    ```              F3     -Age-:@UHS-yrs  ```
    ```              F4     Height  ```
    ```              F5     Weight  ```
    ```    ```
    ```             * create rename statemnts  ```
    ```             select  ```
    ```                catx(" ",_name_,"as",compress(col1,".:$-@+%^*()+%^&*()"))  ```
    ```             into:  ```
    ```                rens separated by ","  ```
    ```             from  ```
    ```                namxpo  ```
    ```    ```
    ```              F1 as Name  ```
    ```             ,F2 as SexUHS  ```
    ```             ,F3 as AgeUHSyrs  ```
    ```             ,F4 as Height  ```
    ```             ,F5 as Weight  ```
    ```    ```
    ```           * do the rename on excel side then import to SAS;  ```
    ```    ```
    ```              from connection to Excel (header=no)  ```
    ```              (  ```
    ```               Select  ```
    ```                   &rens     ** f1 as Name  ```
    ```                             ** f2 as AgeUHSYrs  ```
    ```                             ...  ```
    ```               from  ```
    ```                  [class$A2:Z99] ** start with second row  ```
    ```    ```
    ```  see  ```
    ```  https://goo.gl/JfoKjj  ```
    ```  https://communities.sas.com/t5/Base-SAS-Programming/Variable-names-with-quot-quot/m-p/400579  ```
    ```    ```
    ```    ```
    ```  HAVE  (Excel sheet with bad names)  ```
    ```  =======================================  ```
    ```    ```
    ```     d:/xls/class.xlsx  ```
    ```    ```
    ```         +---------------------------------------------------------------------+  ```
    ```          |     A      |    B         |     C        |    D       |    E       |  ```
    ```          +-------------------------+----------------|-------------------------+  ```
    ```       1  |Name:-+@.$%^|-Sex-:@UHS+%^*|-Age-:@UHS-yrs|  HEIGHT    |  WEIGHT    |  ```
    ```          +------------+--------------+--------------+------------+------------+  ```
    ```       2  | ALFRED     |    M         |    14        |    69      |  112.5     |  ```
    ```          +------------+--------------+--------------+------------+------------+  ```
    ```           ...  ```
    ```          +------------+--------------+--------------+------------+------------+  ```
    ```       20 | WILLIAM    |    M         |    15        |   66.5     |  112       |  ```
    ```          +------------+--------------+--------------+------------+------------+  ```
    ```    ```
    ```     [CLASS$]  ```
    ```    ```
    ```    ```
    ```    ```
    ```  WANT   SAS dataset with nice names)  ```
    ```  ====================================  ```
    ```    ```
    ```     WORK.WANT total obs=19  ```
    ```    ```
    ```      Obs    NAME       SEXUHS    AGEUHSYRS    HEIGHT    WEIGHT  ```
    ```    ```
    ```        1    Alfred       M           14        69.0      112.5  ```
    ```        2    Alice        F           13        56.5       84.0  ```
    ```        3    Barbara      F           13        65.3       98.0  ```
    ```        4    Carol        F           14        62.8      102.5  ```
    ```        5    Henry        M           14        63.5      102.5  ```
    ```    ```
    ```    ```
    ```  *                _               _  ```
    ```   _ __ ___   __ _| | _____  __  _| |_____  __  ```
    ```  | '_ ` _ \ / _` | |/ / _ \ \ \/ / / __\ \/ /  ```
    ```  | | | | | | (_| |   <  __/  >  <| \__ \>  <  ```
    ```  |_| |_| |_|\__,_|_|\_\___| /_/\_\_|___/_/\_\  ```
    ```    ```
    ```  ;  ```
    ```    ```
    ```  options validvarname=any;  ```
    ```  %utlfkil(d:/xls/class.xlsx);  ```
    ```  libname xel "d:/xls/class.xlsx";  ```
    ```  data xel.class;  ```
    ```    set sashelp.class(rename=(  ```
    ```       name='Name:-+@#$%^'n  ```
    ```       sex='-Sex-:@UHS+%^&*()'n  ```
    ```       age='-Age-:@UHS-yrs'n  ```
    ```    ));  ```
    ```  run;quit;  ```
    ```  libname xel clear;  ```
    ```  options validvarname=upcase;  ```
    ```    ```
    ```  *          _       _   _  ```
    ```   ___  ___ | |_   _| |_(_) ___  _ __  ```
    ```  / __|/ _ \| | | | | __| |/ _ \| '_ \  ```
    ```  \__ \ (_) | | |_| | |_| | (_) | | | |  ```
    ```  |___/\___/|_|\__,_|\__|_|\___/|_| |_|  ```
    ```    ```
    ```  ;  ```
    ```    ```
    ```  proc datasets lib=work kill;  ```
    ```  run;quit;  ```
    ```    ```
    ```  %symdel rens / nowarn;  ```
    ```    ```
    ```  proc sql dquote=ansi;  ```
    ```   connect to excel (Path="d:/xls/class.xlsx" mixed=yes header=no);  ```
    ```      create  ```
    ```          table names as  ```
    ```      select  ```
    ```           *  ```
    ```          from connection to Excel  ```
    ```          (  ```
    ```           Select  ```
    ```               top 1 *  ```
    ```           from  ```
    ```              [class$]  ```
    ```          );  ```
    ```      disconnect from Excel;  ```
    ```  Quit;  ```
    ```    ```
    ```  proc transpose data=names out=namxpo;  ```
    ```    var _all_;  ```
    ```  run;quit;  ```
    ```    ```
    ```  proc sql;  ```
    ```     select  ```
    ```        catx(" ",_name_,"as",compress(col1,".:$-@+%^*()+%^&*()"))  ```
    ```     into:  ```
    ```        rens separated by ","  ```
    ```     from  ```
    ```        namxpo  ```
    ```  ;quit;  ```
    ```    ```
    ```   proc sql dquote=ansi;  ```
    ```    connect to excel (Path='d:/xls/class.xlsx' mixed=yes header=no);  ```
    ```       create  ```
    ```           table want  as  ```
    ```       select  ```
    ```            *  ```
    ```           from connection to Excel  ```
    ```           (  ```
    ```            Select  ```
    ```                &rens  ```
    ```            from  ```
    ```               [class$A2:Z99]  ```
    ```           );  ```
    ```       disconnect from Excel  ```
    ```   ;Quit;  ```
