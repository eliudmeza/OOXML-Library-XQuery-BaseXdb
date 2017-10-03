(:
 : --------------------------------
 : Standard ECMA-376
 : The Office Open XML File Formats [Office Open XML Workbook] Library 
 : for BaseX 8.4+
 : By Eliúd Santiago Meza y Rivera eliud.meza@gmail.com
 : --------------------------------
 :BSD 3-Clause License
 :
 :Copyright (c) 2016 - 2017, Eliud Santiago Meza y Rivera
 :All rights reserved.
 :
 :Redistribution and use in source and binary forms, with or without
 :modification, are permitted provided that the following conditions are met:
 :
 :* Redistributions of source code must retain the above copyright notice, this
 :  list of conditions and the following disclaimer.
 :
 :* Redistributions in binary form must reproduce the above copyright notice,
 :  this list of conditions and the following disclaimer in the documentation
 :  and/or other materials provided with the distribution.
 :
 :* Neither the name of the copyright holder nor the names of its
 :  contributors may be used to endorse or promote products derived from
 :  this software without specific prior written permission.
 :
 :THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS"
 :AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE
 :IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE
 :DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT HOLDER OR CONTRIBUTORS BE LIABLE
 :FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL
 :DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR
 :SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER
 :CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY,
 :OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE
 :OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE. 
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

(:declare default element namespace "http://schemas.openxmlformats.org/spreadsheetml/2006/main";:)

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
2017-09-20: change param-type to string
--------- :)
declare function xlsx:get-sheets(
   $file as xs:string
) as element()? {
  try {
    element sheets {
      for $s in fn:parse-xml(
         archive:extract-text(xlsx:get-file($file),"xl/workbook.xml")
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
declare %private function xlsx:get-Workbook-Relationships(
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
declare %private function xlsx:get-rId-worksheet(
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
  $file as xs:string
) as item()* {
  try {
    let $ss := fn:parse-xml(
      archive:extract-text(
         xlsx:get-file($file),"xl/styles.xml")
      )(:/xlsx-spreadsheetml:styleSheet:)
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

(: ---------
Se necesita trabajar más ... // need more work ... 
--------- :)
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
Convert a date to int value for excel...
--------- :)
declare %private function xlsx:date-to-int(
   $value as xs:string
) as xs:date {
  ( ( xs:date($value) + xs:dayTimeDuration('P2D') ) -
                         xs:date('1900-01-01')
  ) div xs:dayTimeDuration('P1D')  
};

(: ---------
Convert a int value to date value for excel...
--------- :)
declare %private function xlsx:int-to-date(
   $value as xs:integer
) as xs:date {
   xs:date('1900-01-01') + 
   xs:dayTimeDuration('P' || ($value - 2) cast as xs:string || 'D')
   
};


declare %private function xlsx:decimal-to-fraction(
   $numero as xs:decimal 
) as xs:string {
   let $decimales := substring-after(string($numero - floor($numero)),'0.') 
   return (
      $decimales
      || '/' ||
      '1' || string-join((for $i in 1 to string-length($decimales) return '0'))
   )   
};

(: ---------
Display format of a value
--------- :)
declare function xlsx:format-value(
   $data as xs:string, 
   $excel-format-code as xs:integer) 
   as xs:string {
try {   
   switch ($excel-format-code)
      case 0 
         return $data
      case 1 
         return $data
      case 2 (: 0.00 :)
         return format-number($data cast as xs:double,'#.00')
      case 3 (: #,##0 :)
         return string(format-number($data cast as xs:double,'#,##0'))
      case 4 (: #,##0.00 :)
         return string(format-number($data cast as xs:double,'#,##0.00'))
      case 9 (: 0% :)
         return string(format-number($data cast as xs:double,'0%'))
      case 10 (: 0.00%:)
         return string(format-number($data cast as xs:double,'0.00%'))
      case 11 (: 0.00E+00  yet ... :)
         return $data
      case 12 (: # ?/? yet ... :)
         return xlsx:decimal-to-fraction ($data cast as xs:decimal)
      case 13 (: # ??/?? yet ... :)
         return xlsx:decimal-to-fraction ($data cast as xs:decimal)
      case 14 (: mm-dd-yy :)
         return (
            if (string(number($data)) != 'NaN' )
            then (
               format-date(xlsx:int-to-date($data cast as xs:integer), 
                  "[M01]-[D01]-[Y01]")
            )            
            else( format-date($data cast as xs:date, "[M01]-[D01]-[Y01]") )
         )
      case 15 (: d-mmm-yy :)
         return (
            if (string(number($data)) != 'NaN' )
            then (
               format-date(xlsx:int-to-date($data cast as xs:integer), 
               "[D]-[Mn,*-3]-[Y01]")
            )            
            else( format-date($data cast as xs:date, "[D]-[Mn,*-3]-[Y01]") )
         )
         
      case 16 (: d-mmm :)
         return (
            if (string(number($data)) != 'NaN' )
            then  (
               format-date(xlsx:int-to-date($data cast as xs:integer), 
               "[D01]-[Mn,*-3]")
            )
            else (
               format-date($data cast as xs:date, "[D01]-[Mn,*-3]")
            )
         )
      case 17 (: mmm-yy  :)
         return format-date($data cast as xs:date, "[Mn,*-3]-[Y01]")
      case 18 (: h:mm AM/PM :)
         return 
            format-time($data cast as xs:time, "[h]:[m01] [PN]", "en", (), ())
      case 19 (: h:mm:ss AM/PM :)
         return 
            format-time($data cast as xs:time, "[h]:[m01]:[s01] [PN]", "en", 
               (), ())
      case 20 (: h:mm :)
         return format-time($data cast as xs:time, "[h]:[m01]", "en", (), ())
      case 21 (: h:mm:ss :)
         return format-time($data cast as xs:time, "[h]:[m01]:[s01]", "en", 
            (), ())
      case 22 (: m/d/yy h:mm :)
         return format-date($data cast as xs:dateTime, "[m]-[d]-[y01]")
      case 37 (: #,##0 ;(#,##0) yet :)
         return $data
      case 38 (: #,##0 ;[Red](#,##0) yet :)
         return $data
      case 39 (: #,##0.00;(#,##0.00) yet :)
         return $data
      case 40 (: #,##0.00;[Red](#,##0.00) yet :)
         return $data
      case 45 (: mm:ss :)
         return 
            format-time($data cast as xs:time, "[m01]:[s01]", "en", (), ())
      case 46 (: [h]:mm:ss :)
         return 
            format-time($data cast as xs:time, "[h]:[m01]:[s01]", "en", (), ())
      case 47 (: mmss.0 :)
         return 
            format-time($data cast as xs:time, "[m01][s01].0", "en", (), ())
      case 48 (: ##0.0E+0 yet :)
         return $data
      case 49 (: @ :)
         return string($data)
      default 
         return $data
} catch * {
   element error {
      element error_code {$err:code},
      element error_description {$err:description},
      element error_value{$err:value},
      element error_module{$err:module},
      element error_line_number{$err:line-number},
      element error_column_number{$err:column-number},
      element error_additional{$err:additional},
      element error_variable_data {$data},
      element error_variable_excel_format_code {$excel-format-code},
      element error_function_name { 'xlsx:format-value' }
   }    
}         
};

(: ---------
Display format of a value
--------- :)
declare function xlsx:display-cell-value(
   $c as item()*,
   $style as item()*,
   $fss as item()*
) as item ()* {
   if (empty($c/@t)) 
      then (
         xlsx:format-value(
            string(data($c/descendant::xlsx-spreadsheetml:v)),
            $style/@numFmtId cast as xs:integer)
      ) 
      else (
         switch ( string(data($c/@t)) )  
            case "b" (: boolean type:)
               return $c/descendant::xlsx-spreadsheetml:v
            case "d" (: date-time type:)
               return
                  xlsx:format-value(
                     string(
                        data(
                           $c/descendant::xlsx-spreadsheetml:v
                        )
                     ),
                     $style/@numFmtId cast as xs:integer
                  )
            case "e" (: error type:)
               return "Error"
            case "inlineStr" (: In Line String type:)
               return 
                  data (
                     $c/descendant::xlsx-spreadsheetml:is/
                     descendant::xlsx-spreadsheetml:t
                  )
            case "n" (: number type:)
               return
                  xlsx:format-value(
                     string(
                        data(
                           $c/descendant::xlsx-spreadsheetml:v
                        )
                     ),
                     $style/@numFmtId cast as xs:integer
                  )
            case "s" 
               return 
                  data ($fss[position() = 
                     data(
                        $c/descendant::xlsx-spreadsheetml:v) 
                        + 1]
                     ) 
            case "str"
               return $c/descendant::xlsx-spreadsheetml:v
            default return $c/descendant::xlsx-spreadsheetml:v
         )   
};

(: ---------
Returns the Calc-Chain contained in the workbook
--------- :)
declare %private function xlsx:get-calcChain(
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
declare %private function xlsx:get-xml-path-worksheet(
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
    return $sheet-data/descendant::
       xlsx-spreadsheetml:row[@r=fn:upper-case($row_number)]
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
      let $fstyle:= xlsx:get-style($file)
      let $fss   := xlsx:get-sharedStrings(xlsx:get-file($file))
      let $style-Cell := $fstyle/descendant::xlsx-spreadsheetml:cellXfs/
         descendant::xlsx-spreadsheetml:xf
         [position() = (fn:number($c/@s) + 1)]      
      return (
         xlsx:display-cell-value($c,$style-Cell,$fss)
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
declare %private updating function xlsx:set-cell-value-original(
   $file  as xs:string,
   $sheet as xs:string,
   $cell  as xs:string,
   $value as xs:anyAtomicType
) {
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
         (:Solo se actualiza la celda... pero se debe actualizar el estilo :)
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
 typeswitch ($value)
    case $value as xs:byte |
       xs:short |
       xs:int |
       xs:long |
       xs:unsignedByte |
       xs:unsignedShort |
       xs:unsignedInt |
       xs:unsignedLong |
       xs:positiveInteger |
       xs:nonNegativeInteger |
       xs:negativeInteger | 
       xs:nonPositiveInteger |
       xs:integer |
       xs:decimal |
       xs:float
      return xlsx:update-number-value($file,$sheet,$cell,$value) 
   case $value as xs:string |
      xs:normalizedString |
      xs:token |
      xs:language |
      xs:NMTOKEN |
      xs:Name |
      xs:NCName |
      xs:ID |
      xs:IDREF |
      xs:ENTITY
      return xlsx:update-string-value($file,$sheet,$cell,$value) 
   case $value as xs:date
      return xlsx:update-date-value($file,$sheet,$cell,$value)           
   default return ()
};

(: ---------
Export the worksheet data to an html table ...
--------- :)
declare function xlsx:worksheet-to-table(
   $file  as xs:string, 
   $sheet as xs:string
) as item()*{
   try {
      (:new code ... I hope a better code ... :)
      let $wsd := xlsx:get-worksheet-data($file, $sheet)
      let $fss := xlsx:get-sharedStrings(xlsx:get-file($file))
      let $fstyle := xlsx:get-style($file)
      let $rows := $wsd/descendant::xlsx-spreadsheetml:row
      return element table{
         attribute id {data($sheet)},
         for $r in $rows
         return (
            element tr {
               attribute id {'row-' || $r/@r},
               for $c in $r/descendant::xlsx-spreadsheetml:c
               let $style-Cell := 
                     $fstyle/descendant::xlsx-spreadsheetml:cellXfs/
                     descendant::xlsx-spreadsheetml:xf
                     [position() = (fn:number($c/@s) + 1)]
               return (
                  element td {
                     attribute id {'cell-' || $c/@r},
                     attribute s {$c/@s || ' - ' || (fn:number($c/@s) + 1)},
                     xlsx:display-cell-value($c,$style-Cell, $fss)
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
