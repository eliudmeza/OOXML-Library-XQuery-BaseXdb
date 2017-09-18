#Module for BaseX 8.4+ to handle OOXML Workbooks [ECMA-376]

This module help to read data and make simples updates from XML Workbooks files [ECMA-376] for BaseX 8.4+

## Installing this module

1. via command:
    ```REPO INSTALL OOXML-Module-for-BaseXdb.xqm```
    
2. via GUI:
 * Option
 * Packages ...
 * Instal ...
 * Select the file "OOXML-Module-for-BaseXdb.xqm"
##
##Use

Use the example below 

```xquery
import module namespace xlsx = 'http://basex.org/modules/ECMA-376/spreadsheetml';

(: Return the cell value of a worksheet :)
xlsx:get-cell-value('Libro1.xlsx','Hoja1','B1')

... 

(: Return the cells of a column :)
xlsx:get-col('Libro1.xlsx','Hoja1','B')

... 

(: Return the cells of a row :)
xlsx:get-col('Libro1.xlsx','Hoja1','13')

... 

(: Update the cell value of a worksheet :)
xlsx:set-cell-value('Libro1.xlsx','Hoja1','B1',23.45)

...

(: Export the content of a worksheet into simple table :)
xlsx:worksheet-to-table('Libro1.xlsx','Hoja1')
```


