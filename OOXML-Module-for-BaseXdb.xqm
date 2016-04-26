(:
 : --------------------------------
 : Standard ECMA-376
 : The Office Open XML File Formats [Office Open XML Workbook] Library for BaseX 8.4+
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

 : For more information on the FunctX XQuery library, contact contrib@functx.com.

 : @version 1.0
 : @see     ...
 :) 

xquery version "3.1";
module namespace xlsx = 'http://basex.org/modules/ECMA-376/spreadsheetml';
(:OfficeOpenXML-Workbook:)
import module namespace file = "http://expath.org/ns/file";

declare namespace xlsx-Content-Types = "http://schemas.openxmlformats.org/package/2006/content-types"; 
declare namespace xlsx-Core-Properties = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties"; 
declare namespace xlsx-Digital-Signatures = "http://schemas.openxmlformats.org/package/2006/digital-signature";
declare namespace xlsx-Relationships = "http://schemas.openxmlformats.org/package/2006/relationships";
declare namespace xlsx-Markup-Compatibility = "http://schemas.openxmlformats.org/markup-compatibility/2006";
declare namespace xlsx-spreadsheetml = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
declare namespace xlsx-sharedStrings = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings";
declare namespace xlsx-x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac";
declare namespace xlsx-mc="http://schemas.openxmlformats.org/markup-compatibility/2006";


declare function xlsx:get-file(
   $file as xs:string
) as xs:base64Binary {
   try {
     let $f := file:read-binary($file)
     return $f    
   } catch * {
      element error {
         element error_code {$err:code},
         element error_description {$err:description},
         element error_value{$err:value},
         element error_module{$err:module},
         element error_line_number{$err:line-number},
         element error_column_number{$err:column-number},
         element error_additional{$err:additional}      
      }
   }
};

declare function xlsx:get-sheets(
   $file as xs:base64Binary
) as element()? {
   element sheets {
      for $s in fn:parse-xml(
         archive:extract-text($file,"xl/workbook.xml")
      )/descendant::xlsx-spreadsheetml:sheet 
      return 
         element sheet {
            $s/@name
         }
   }
};

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

declare function xlsx:get-rId-worksheet(
   $file  as xs:base64Binary, 
   $sheet as xs:string
) as xs:string* {
   let $rs:= fn:parse-xml(
      archive:extract-text(
         $file,"xl/workbook.xml")
      )/descendant::xlsx-spreadsheetml:sheets
       /descendant::xlsx-spreadsheetml:sheet
          [@name = $sheet]/attribute::*[name(.) = 'r:id']
   return data($rs)   
};

declare function xlsx:get-sharedStrings(
   $file as xs:base64Binary
) as item()* {
   let $ss := fn:parse-xml(
      archive:extract-text(
         $file,"xl/sharedStrings.xml")
      )/descendant::xlsx-spreadsheetml:t
   return $ss
};

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

declare function xlsx:get-worksheet-data(
   $file  as xs:string, 
   $sheet as xs:string
) as item()*{
   try {
      let $f := file:read-binary($file)
      return (
         let $rs := fn:parse-xml(
            archive:extract-text(
               $f,
               "xl/" || xlsx:get-xml-path-worksheet($f,$sheet)
            )
         )/descendant::xlsx-spreadsheetml:sheetData
         return $rs
      )
   } catch * {
      element error {
         element error_code {$err:code},
         element error_description {$err:description},
         element error_value{$err:value},
         element error_module{$err:module},
         element error_line_number{$err:line-number},
         element error_column_number{$err:column-number},
         element error_additional{$err:additional}      
      }      
   }
};

declare function xlsx:get-cell-value(
   $file as xs:string,
   $sheet as xs:string,
   $cell as xs:string
) as item()* {
   try {
      let $f := file:read-binary($file)
      return (
         let $rs := fn:parse-xml(
            archive:extract-text(
               $f,
               "xl/" || xlsx:get-xml-path-worksheet($f,$sheet)
            )
         )/descendant::xlsx-spreadsheetml:sheetData
          /descendant::xlsx-spreadsheetml:c[@r=$cell]
         return
            if ( fn:empty($rs/@t) )
            then (
               data($rs/descendant::xlsx-spreadsheetml:v)
            )
            else ( 
               data(xlsx:get-sharedStrings($f)[position() = data($rs/descendant::xlsx-spreadsheetml:v)+1 ])
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
         element error_additional{$err:additional}      
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
   
declare %updating function xlsx:set-cell-value(
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

(:
declare %updating 
function xlsx:set-cell-value(
   $file as xs:string,
   $sheet as xs:string,
   $cell as xs:string,
   $new-value as xs:string
) {
   try {
      let $f := file:read-binary($file)
      return (
         let $rs := fn:parse-xml(
            archive:extract-text(
               $f,
               "xl/" || xlsx:get-xml-path-worksheet($f,$sheet)
            )
         )/descendant::xlsx-spreadsheetml:sheetData
          /descendant::xlsx-spreadsheetml:c[@r=$cell]
         return(
            'valor cambiado',
            replace value of node $rs/descendant::xlsx-spreadsheetml:v with $new-value            
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
         element error_additional{$err:additional}      
      }      
   }
};
:)

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

            (:
            if ( fn:empty($rs/@t) )
            then (
               data($rs/descendant::xlsx-spreadsheetml:v)
            )
            else ( 
               data(xlsx:get-sharedStrings($f)[position() = data($rs/descendant::xlsx-spreadsheetml:v)+1 ])
            )
            :)
