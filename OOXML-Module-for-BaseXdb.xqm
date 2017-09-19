(:
 : --------------------------------
 : The Office Open XML File Formats [Office Open XML Workbook] XQuery Library for BaseX 8.4+
 : Standard ECMA-376
 : --------------------------------
 : Copyright (C) 2016 Eliúd Santiago Meza y Rivera 
 : email: eliud.meza@gmail.com
 :        eliud.meza@outlook.com
 : This library is free software; you can redistribute it and/or
 : modify it under the terms of the GNU Lesser General Public
 : License as published by the Free Software Foundation; either
 : version 2.1 of the License.
 : This library is distributed in the hope that it will be useful,
 : but WITHOUT ANY WARRANTY; without even the implied warranty of
 : MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
 : Lesser General Public License for more details.
 : You should have received a copy of the GNU Lesser General Public
 : License along with this library; if not, write to the Free Software
 : Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA  02110-1301  USA
 : For more information on the OOXML XQuery Library for BaseX, contact eliud.meza@gmail.com.
 : @version 1.0
 : @see     ...
 :) 

xquery version "3.1";
module namespace xlsx = 'http://basex.org/modules/ECMA-376/spreadsheetml';
(:OfficeOpenXML-Workbook:)
import module namespace file = "http://expath.org/ns/file";
(:import module namespace functx = "http://www.functx.com";:)

declare namespace xlsx-Content-Types = "http://schemas.openxmlformats.org/package/2006/content-types"; 
declare namespace xlsx-Core-Properties = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties"; 
declare namespace xlsx-Digital-Signatures = "http://schemas.openxmlformats.org/package/2006/digital-signature";
declare namespace xlsx-Relationships = "http://schemas.openxmlformats.org/package/2006/relationships";
declare namespace xlsx-Markup-Compatibility = "http://schemas.openxmlformats.org/markup-compatibility/2006";
declare namespace xlsx-spreadsheetml = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
declare namespace xlsx-sharedStrings = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings";
declare namespace xlsx-x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac";
declare namespace xlsx-mc="http://schemas.openxmlformats.org/markup-compatibility/2006";

(: ---------
Return a binary representation of the workbook file...
--------- :)
declare function xlsx:get-file(
   $file as xs:string
) as xs:base64Binary {
   try {
     let $f := file:read-binary($file)
     return $f    
   } catch * {
      xs:base64Binary(element error {
         element error_code {$err:code},
         element error_description {$err:description},
         element error_value{$err:value},
         element error_module{$err:module},
         element error_line_number{$err:line-number},
         element error_column_number{$err:column-number},
         element error_additional{$err:additional},
         element error_function_name { 'xlsx:get-file' }
      })
   }
};

(: ---------
Return a element containing the names of the worksheet of the workbook
--------- :)
declare function xlsx:get-sheets(
   $file as xs:base64Binary
) as element()? {
  try {
    element sheets {
      for $s in fn:parse-xml(
         archive:extract-text($file,"xl/workbook.xml")
      )/descendant::xlsx-spreadsheetml:sheet 
      return 
         element sheet {
            $s/@name
         }
    }
  } catch * {
      element error {
         element error_code {$err:code},
         element error_description {$err:description},
         element error_value{$err:value},
         element error_module{$err:module},
         element error_line_number{$err:line-number},
         element error_column_number{$err:column-number},
         element error_additional{$err:additional},
         element error_function_name { 'xlsx:get-sheets' }
      }
  }
};

(: ---------
Returns the Relationships elements contained in the workbook
--------- :)
declare function xlsx:get-Workbook-Relationships(
   $file as xs:base64Binary
) as item()*  {
   let $rs := fn:parse-xml(
      archive:extract-text(
         $file,
         "xl/_rels/workbook.xml.rels" )
      )
   return $rs
};

(: ---------
Return a string of the id of the worksheet
--------- :)
declare function xlsx:get-rId-worksheet(
   $file  as xs:base64Binary, 
   $sheet as xs:string
) as xs:string* {
  try {
    let $rs:= fn:parse-xml(
      archive:extract-text(
         $file,"xl/workbook.xml")
      )/descendant::xlsx-spreadsheetml:sheets
       /descendant::xlsx-spreadsheetml:sheet
          [@name = $sheet]/attribute::*[name(.) = 'r:id']
    return data($rs)   
  }  catch * {
    let $a:= ''
    return data($a)
  }
};

(: ---------
Returns the Shared-String elements contained in the workbook
--------- :)
declare function xlsx:get-sharedStrings(
   $file as xs:base64Binary
) as item()* {
  try {
    let $ss := fn:parse-xml(
      archive:extract-text(
         $file,"xl/sharedStrings.xml")
      )/descendant::xlsx-spreadsheetml:si
    return $ss
  } catch * {
      element error {
         element error_code {$err:code},
         element error_description {$err:description},
         element error_value{$err:value},
         element error_module{$err:module},
         element error_line_number{$err:line-number},
         element error_column_number{$err:column-number},
         element error_additional{$err:additional},
         element error_function_name { 'xlsx:get-sharedStrings' }
      }
  }
};

(: ---------
Returns the style elements contained in the workbook
--------- :)
declare function xlsx:get-style(
  $file as xs:base64Binary
) as item()* {
  try {
    let $ss := fn:parse-xml(
      archive:extract-text(
         $file,"xl/styles.xml")
      )/styleSheet/descendant::*
    return $ss
  } catch * {
      element error {
         element error_code {$err:code},
         element error_description {$err:description},
         element error_value{$err:value},
         element error_module{$err:module},
         element error_line_number{$err:line-number},
         element error_column_number{$err:column-number},
         element error_additional{$err:additional},
         element error_function_name { 'xlsx:get-style' }
      }
  }
};

(:Se necesita trabajar más ... // need more work ... :)
declare function xlsx:set-style(
  $file as xs:base64Binary,
  $new-style as item()*
) as item()* {
  try {
    element something{ "aa"}
  } catch * {
      element error {
         element error_code {$err:code},
         element error_description {$err:description},
         element error_value{$err:value},
         element error_module{$err:module},
         element error_line_number{$err:line-number},
         element error_column_number{$err:column-number},
         element error_additional{$err:additional},
         element error_function_name { 'xlsx:set-style' }
      }
    
  }
};

(: ---------
Returns the Calc-Chain contained in the workbook
--------- :)
declare function xlsx:get-calcChain(
  $file as xs:base64Binary
) as item()* {
  try {
    let $ss := fn:parse-xml(
      archive:extract-text(
         $file,"xl/calcChain.xml")
      )/descendant::xlsx-spreadsheetml:t
    return $ss
  } catch * {
      element error {
         element error_code {$err:code},
         element error_description {$err:description},
         element error_value{$err:value},
         element error_module{$err:module},
         element error_line_number{$err:line-number},
         element error_column_number{$err:column-number},
         element error_additional{$err:additional},
         element error_function_name { 'get-calcChain' }
      }
  }
};


(: ---------
Returns the xml path of the worksheet contained in the book
--------- :)
declare function xlsx:get-xml-path-worksheet(
   $file as xs:base64Binary, 
   $sheet as xs:string   
) as xs:string* {
   let $rsId := xlsx:get-rId-worksheet($file, $sheet)
   let $xml-path := xlsx:get-Workbook-Relationships($file)
      /descendant::xlsx-Relationships:Relationships
      /descendant::xlsx-Relationships:Relationship[@Id = data($rsId)]
   return data($xml-path/@Target) 
};

(: ---------
Returns the content of the worksheet 
--------- :)
declare function xlsx:get-worksheet-data (
   $file  as xs:string, 
   $sheet as xs:string
) as item()*{
  try {
    let $f:= xlsx:get-file($file)
    return (
      let $rs := fn:parse-xml(
            archive:extract-text(
               $f,
               "xl/" || xlsx:get-xml-path-worksheet($f,$sheet)
            )
         )/descendant::xlsx-spreadsheetml:sheetData
         return $rs
    )
  }catch * {
      element error {
         element error_code {$err:code},
         element error_description {$err:description},
         element error_value{$err:value},
         element error_module{$err:module},
         element error_line_number{$err:line-number},
         element error_column_number{$err:column-number},
         element error_additional{$err:additional},
         element error_function_name { 'xlsx:get-worksheet-data' }
      }      
   }
};

(: ---------
Returns the content of a specified row in the worksheet
--------- :)
declare function xlsx:get-row(
  $file as xs:string,
  $sheet as xs:string,
  $row_number as xs:string
) as item()* {
  try {
    let $sheet-data := xlsx:get-worksheet-data($file,$sheet)
    return $sheet-data/descendant::xlsx-spreadsheetml:row[@r=fn:upper-case($row_number)]
  } catch * {
      element error {
         element error_code {$err:code},
         element error_description {$err:description},
         element error_value{$err:value},
         element error_module{$err:module},
         element error_line_number{$err:line-number},
         element error_column_number{$err:column-number},
         element error_additional{$err:additional},
         element error_function_name { 'xlsx:get-row' }
      }      
   }
};

(: ---------
Returns the content of a specified column in the worksheet
--------- :)
declare function xlsx:get-col(
  $file as xs:string,
  $sheet as xs:string,
  $column as xs:string
) as item()* {
  try { 
    let $sheet-data := xlsx:get-worksheet-data($file,$sheet)
    let $pattern := '^('|| fn:upper-case($column) ||')+\d'
    return $sheet-data/descendant::xlsx-spreadsheetml:c[fn:matches(@r,$pattern)]
    } catch * {
      element error {
         element error_code {$err:code},
         element error_description {$err:description},
         element error_value{$err:value},
         element error_module{$err:module},
         element error_line_number{$err:line-number},
         element error_column_number{$err:column-number},
         element error_additional{$err:additional},
         element error_function_name { 'xlsx:get-col' }
      }      
   }
};

(: ---------
Returns the cell element specified in the worksheet
--------- :)
declare function xlsx:get-cell(
  $file as xs:string,
  $sheet as xs:string,
  $cell as xs:string
) as item()* {
  try { 
    let $sheet-data := xlsx:get-worksheet-data($file,$sheet)
    return $sheet-data/descendant::xlsx-spreadsheetml:c[@r=fn:upper-case($cell)]
  } catch * {
    element error {
       element error_code {$err:code},
       element error_description {$err:description},
       element error_value{$err:value},
       element error_module{$err:module},
       element error_line_number{$err:line-number},
       element error_column_number{$err:column-number},
       element error_additional{$err:additional},
       element error_function_name { 'xlsx:get-cell' }
    }      
 }  
};

(: ---------
Returns the cell value specified in the worksheet
--------- :)
declare function xlsx:get-cell-value(
   $file as xs:string,
   $sheet as xs:string,
   $cell as xs:string
) as item()* {
   try {
     let $c     := xlsx:get-cell($file,$sheet,$cell)
     let $f     := xlsx:get-file($file)
     return (
       if ( fn:empty($c/@t) )
       then ( data($c/descendant::xlsx-spreadsheetml:v) )
       else ( 
         let $ss := xlsx:get-sharedStrings($f)
         let $c-pos := xs:integer(data($c/descendant::xlsx-spreadsheetml:v))+1
         return data($ss[position() = $c-pos])
         ) 
       )
   } catch * {
      element error {
         element error_code {$err:code},
         element error_description {$err:description},
         element error_value{$err:value},
         element error_module{$err:module},
         element error_line_number{$err:line-number},
         element error_column_number{$err:column-number},
         element error_additional{$err:additional},
         element error_function_name { 'xlsx:get-cell-value' }
      }      
   }
};

declare %updating function 
   xlsx:upsert($e as element(), 
          $an as xs:QName, 
          $av as xs:anyAtomicType) 
   {
   let $ea := $e/attribute()[fn:node-name(.) = $an]
   return
      if (fn:empty($ea))
      then insert node attribute {$an} {$av} into $e
      else replace value of node $ea with $av
   };
   
(: ---------
Update the value of the cell --- original function
--------- :)
   declare updating function xlsx:set-cell-value-original(
   $file  as xs:string,
   $sheet as xs:string,
   $cell  as xs:string,
   $value as xs:anyAtomicType
) {
  (:
   let $file := 'Libro1.xlsx'
   let $sheet := 'Hoja1'
   let $cell := 'B1'
   let $value := 7890 :)
   let $f  := file:read-binary($file)
   let $xml-sheet := 'xl/' || xlsx:get-xml-path-worksheet($f,$sheet)
   let $entry := 
      copy $rs := fn:parse-xml(
                     archive:extract-text(
                        $f,
                        'xl/' || xlsx:get-xml-path-worksheet($f,$sheet)
                     )
                  )
      modify replace value of node $rs/descendant::xlsx-spreadsheetml:sheetData
                   /descendant::xlsx-spreadsheetml:c[@r=$cell]
                   /descendant::xlsx-spreadsheetml:v
       with $value
      return fn:serialize($rs)
   let $updated := archive:update($f,$xml-sheet,$entry)
   return file:write-binary($file,$updated)
};

(: ---------
Update the number value of the cell
--------- :)
declare %updating
function xlsx:update-number-value(
   $file  as xs:string,
   $sheet as xs:string,
   $cell  as xs:string,
   $value as xs:anyAtomicType
) {
  let $f  := xlsx:get-file($file)    
  let $xml-sheet := 'xl/' || xlsx:get-xml-path-worksheet($f,$sheet)
  let $row_number := tokenize(fn:upper-case($cell),'[A-Z]')
  let $row_number := $row_number[count($row_number)]
  let $new-cell-node := element c {
        attribute r {fn:upper-case($cell)},
        element v {
          $value
        }
      }
  let $new-row-node := element row {
        attribute r{$row_number},
        $new-cell-node
      }
  let $entry := 
    (:cell exists???:)
    if ( fn:empty(xlsx:get-cell($file,$sheet,fn:upper-case($cell))) ) 
    then ( 
      (:row exists???:)  
      if ( fn:empty(xlsx:get-row ($file,$sheet,$row_number)) )
      then (
        copy $rs := fn:parse-xml(
                       archive:extract-text(
                          $f,
                          $xml-sheet
                       )
                    )
        modify insert node $new-row-node
               after $rs/descendant::xlsx-spreadsheetml:sheetData
                        /descendant::xlsx-spreadsheetml:row
                        [xs:integer(@r) lt xs:integer($row_number)]
                        [last()]
        return fn:serialize($rs)
      )
      else(
        copy $rs := fn:parse-xml(
                       archive:extract-text(
                          $f,
                          $xml-sheet
                       )
                    )
        modify insert node $new-cell-node
               after $rs/descendant::xlsx-spreadsheetml:sheetData
                        /descendant::xlsx-spreadsheetml:row
                         [xs:integer(@r) eq xs:integer($row_number) ]
                        /descendant::xlsx-spreadsheetml:c[@r lt $cell][last()]
        return fn:serialize($rs)
      )
    ) 
    else (
      copy $rs := fn:parse-xml(
                     archive:extract-text(
                        $f,
                        $xml-sheet
                     )
                  )
      modify replace node $rs/descendant::xlsx-spreadsheetml:sheetData
                   /descendant::xlsx-spreadsheetml:c[@r=$cell]
              
       with $new-cell-node
      return fn:serialize($rs)
    ) 
  let $updated := archive:update($f,$xml-sheet,$entry)
  return file:write-binary($file,$updated)
};

(: ---------
Update the string value of the cell
--------- :)
declare %updating
function xlsx:update-string-value(
   $file  as xs:string,
   $sheet as xs:string,
   $cell  as xs:string,
   $value as xs:anyAtomicType
) { 
  let $f  := xlsx:get-file($file)    
  let $xml-sheet := 'xl/' || xlsx:get-xml-path-worksheet($f,$sheet)
  let $row_number := tokenize(fn:upper-case($cell),'[A-Z]')
  let $row_number := $row_number[count($row_number)]
  let $new-cell-node := element c {
    attribute r { fn:upper-case($cell) },
    attribute t {"inlineStr"},
    element is {
      element t { $value }
    }
  }
  let $new-row-node := element row {
        attribute r{$row_number},
        $new-cell-node
      }
  let $entry := 
    (:cell exists???:)
    if ( fn:empty(xlsx:get-cell($file,$sheet,fn:upper-case($cell))) ) 
    then ( 
      (:row exists???:)  
      if ( fn:empty(xlsx:get-row ($file,$sheet,$row_number)) )
      then (
        copy $rs := fn:parse-xml(
                       archive:extract-text(
                          $f,
                          $xml-sheet
                       )
                    )
        modify insert node $new-row-node
               after $rs/descendant::xlsx-spreadsheetml:sheetData
                        /descendant::xlsx-spreadsheetml:row
                        [xs:integer(@r) lt xs:integer($row_number)]
                        [last()]
        return fn:serialize($rs)
      )
      else(
        copy $rs := fn:parse-xml(
                       archive:extract-text(
                          $f,
                          $xml-sheet
                       )
                    )
        modify insert node $new-cell-node
               after $rs/descendant::xlsx-spreadsheetml:sheetData
                        /descendant::xlsx-spreadsheetml:row
                         [xs:integer(@r) eq xs:integer($row_number) ]
                        /descendant::xlsx-spreadsheetml:c[@r lt $cell][last()]
        return fn:serialize($rs)
      )
    ) 
    else (
      copy $rs := fn:parse-xml(
                     archive:extract-text(
                        $f,
                        $xml-sheet
                     )
                  )
      modify replace node $rs/descendant::xlsx-spreadsheetml:sheetData
                   /descendant::xlsx-spreadsheetml:c[@r=$cell]
              
       with $new-cell-node
      return fn:serialize($rs)
    ) 
  let $updated := archive:update($f,$xml-sheet,$entry)
  return file:write-binary($file,$updated)
};

(: ---------
Update the date value of the cell
--------- :)
declare updating
function xlsx:update-date-value(
   $file  as xs:string,
   $sheet as xs:string,
   $cell  as xs:string,
   $value as xs:anyAtomicType
) {
  let $f  := xlsx:get-file($file)    
  let $xml-sheet := 'xl/' || xlsx:get-xml-path-worksheet($f,$sheet)
  let $date_to_int:= ( ( xs:date($value) + xs:dayTimeDuration('P2D') ) -
                         xs:date('1900-01-01')) div xs:dayTimeDuration('P1D')  
  let $row_number := tokenize(fn:upper-case($cell),'[A-Z]')
  let $row_number := $row_number[count($row_number)]
  let $new-cell-node := element c {
      attribute r {$cell},
      attribute s {"3"},
      element v { $date_to_int }
    }
  let $new-row-node := element row {
        attribute r{$row_number},
        $new-cell-node
      }
  let $entry := 
    (:cell exists???:)
    if ( fn:empty(xlsx:get-cell($file,$sheet,fn:upper-case($cell))) ) 
    then ( 
      (:row exists???:)  
      if ( fn:empty(xlsx:get-row ($file,$sheet,$row_number)) )
      then (
        copy $rs := fn:parse-xml(
                       archive:extract-text(
                          $f,
                          $xml-sheet
                       )
                    )
        modify insert node $new-row-node
               after $rs/descendant::xlsx-spreadsheetml:sheetData
                        /descendant::xlsx-spreadsheetml:row
                        [xs:integer(@r) lt xs:integer($row_number)]
                        [last()]
        return fn:serialize($rs)
      )
      else(
        copy $rs := fn:parse-xml(
                       archive:extract-text(
                          $f,
                          $xml-sheet
                       )
                    )
        modify insert node $new-cell-node
               after $rs/descendant::xlsx-spreadsheetml:sheetData
                        /descendant::xlsx-spreadsheetml:row
                         [xs:integer(@r) eq xs:integer($row_number) ]
                        /descendant::xlsx-spreadsheetml:c[@r lt $cell][last()]
        return fn:serialize($rs)
      )
    ) 
    else (
      copy $rs := fn:parse-xml(
                     archive:extract-text(
                        $f,
                        $xml-sheet
                     )
                  )
      modify replace node $rs/descendant::xlsx-spreadsheetml:sheetData
                   /descendant::xlsx-spreadsheetml:c[@r=$cell]
              
       with $new-cell-node
      return fn:serialize($rs)
    ) 
  let $updated := archive:update($f,$xml-sheet,$entry)
  return file:write-binary($file,$updated)
};

(: ---------
Update the value of the cell
--------- :)
declare updating function xlsx:set-cell-value(
   $file  as xs:string,
   $sheet as xs:string,
   $cell  as xs:string,
   $value as xs:anyAtomicType
) {
 if (($value instance of xs:byte) or
     ($value instance of xs:short) or 
     ($value instance of xs:int) or
     ($value instance of xs:long) or 
     ($value instance of xs:unsignedByte) or 
     ($value instance of xs:unsignedShort) or 
     ($value instance of xs:unsignedInt) or 
     ($value instance of xs:unsignedLong) or 
     ($value instance of xs:positiveInteger) or 
     ($value instance of xs:nonNegativeInteger) or 
     ($value instance of xs:negativeInteger) or 
     ($value instance of xs:nonPositiveInteger) or 
     ($value instance of xs:integer) or 
     ($value instance of xs:decimal) or 
     ($value instance of xs:float)     ) 
 then ( 
   xlsx:update-number-value($file,$sheet,$cell,$value) )
 else 
   if (($value instance of xs:string) or  
       ($value instance of xs:normalizedString)  or
       ($value instance of xs:token)  or
       ($value instance of xs:language) or 
       ($value instance of xs:NMTOKEN) or 
       ($value instance of xs:Name) or 
       ($value instance of xs:NCName) or 
       ($value instance of xs:ID) or 
       ($value instance of xs:IDREF) or 
       ($value instance of xs:ENTITY)
       
      )
   then  ( 
     xlsx:update-string-value($file,$sheet,$cell,$value) )
   else 
      if ( $value instance of xs:date ) 
      then ( 
        xlsx:update-date-value($file,$sheet,$cell,$value) )
      else ()
};

(: ---------
Export the worksheet data to an html table ...
--------- :)
declare function xlsx:worksheet-to-table(
   $file  as xs:string, 
   $sheet as xs:string
) as item()*{
   try {
      let $sfn := $file
      let $ssn := $sheet
      let $f   := file:read-binary($sfn)
      let $fw  := parse-xml(archive:extract-text($f, "xl/workbook.xml"))
      let $fwrels  := parse-xml(archive:extract-text($f,"xl/_rels/workbook.xml.rels"))
      let $fss := parse-xml(archive:extract-text($f, "xl/sharedStrings.xml"))
         /descendant::xlsx-spreadsheetml:t
      let $fw-id := data($fw
         /descendant::xlsx-spreadsheetml:sheets
            /descendant::xlsx-spreadsheetml:sheet[(@name = $ssn)]
            /@*[(name(.) = "r:id")])
      let $fwrels-xml-path := $fwrels
         /descendant::xlsx-Relationships:Relationships
         /descendant::xlsx-Relationships:Relationship[@Id = data($fw-id)]
      let $fws := parse-xml(
            archive:extract-text($f, 'xl/' || data($fwrels-xml-path/@Target) )
         )/descendant::xlsx-spreadsheetml:sheetData
            /descendant::xlsx-spreadsheetml:row
      return 
         element table {
            attribute id {data($ssn)}
            ,
            attribute worksheet-id {$fw-id},
            attribute worksheet-xml-path {data($fwrels-xml-path/@Target)},
            for $r in $fws
            return (
               element tr {
                  attribute id {'row-' || $r/@r},
                  for $c in $r/descendant::xlsx-spreadsheetml:c
                  return (
                     element td {
                        attribute id {'cell-' || $c/@r},
                        if (empty($c/@t))
                        then data($c/descendant::xlsx-spreadsheetml:v)
                        else 
                           data (
                              $fss[position() = data($c/descendant::xlsx-spreadsheetml:v) + 1]
                           )
                     }
                  )               
               }
            )
         }
   } catch * {
       element error {
          element error_code {$err:code}, 
          element error_description {$err:description}, 
          element error_value { $err:value}, 
          element error_module {$err:module}, 
          element error_line_number {$err:line-number}, 
          element error_column_number {$err:column-number}, 
          element error_additional {$err:additional}
      } 
   }
};
