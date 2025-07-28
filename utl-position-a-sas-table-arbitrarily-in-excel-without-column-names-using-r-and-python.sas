%let pgm=utl-position-a-sas-table-arbitrarily-in-excel-without-column-names-using-r-and-python;

%stop_submission;

Position a sas table arbitrarily in excel without column names using r and python;

Problem
  Position sas table starting at row 3 column 3

TWO SOLUTION
   1 r openxlsx
   2 python openxls

github
https://tinyurl.com/yc2bhtat
https://github.com/rogerjdeangelis/utl-position-a-sas-table-arbitrarily-in-excel-without-column-names-using-r-and-python

sas communites
https://tinyurl.com/39svheb4
https://communities.sas.com/t5/ODS-and-Base-Reporting/Move-Data-from-SAS-to-specific-cells-in-Excel/m-p/756109#M25150

/**************************************************************************************************************************/
/*  INPUT                        | PROCESS                                   |  OUTPUT                                    */
/*  =====                        | =======                                   |  ======                                    */
/*      KEY    VAL               | 1 R OPENXLSX                              |  d:/xls/routput.xlsx                       */
/*                               | ============                              |                                            */
/*  1 IP_SUM   10                |                                           |  --------------------+                     */
/*  2 CV_SUM   20                | %utlfkil(d:/xls/routput.xlsx);            |  | A1| fx     |      |                     */
/*  3 RS_SUM   30                |                                           |  -------------------------+                */
/*                               | %utl_rbeginx;                             |  [_] |  A | B |   C  |  D |                */
/*  options                      | parmcards4;                               |  -------------------------|                */
/*   validvarname=upcase;        | library(haven)                            |   1  |    |   |      |    |                */
/*  libname sd1 "d:/sd1";        | library(openxlsx)                         |   -- |----+---+------+----|                */
/*  data sd1.have;               | wb <- createWorkbook()                    |   2  |    |   |      |    |                */
/*   input key$ val;             | addWorksheet(wb, "want")                  |   -- |----+---+------+----|                */
/*  cards4;                      | have<-read_sas(                           |   3  |    |   |CV_SUM| 20 |                */
/*  IP_SUM 10                    |  "d:/sd1/have.sas7bdat")                  |   -- |----+---+------+----|                */
/*  CV_SUM 20                    | print(have)                               |   4  |    |   |RS_SUM| 30 |                */
/*  RS_SUM 30                    | writeData(wb,                             |   -- |----+---+------+----|                */
/*  ;;;;                         |   sheet = "want",                         |   5  |    |   |IP_SUM| 10 |                */
/*  run;quit;                    |   x = have,                               |   -- |----+---+------+----|                */
/*                               |   startCol = 3,                           |   6  |    |   |      |    |                */
/*                               |   startRow = 3,                           |   -- |----+---+------+----|                */
/*                               |   colNames = FALSE,                       |   7  |    |   |      |    |                */
/*                               |   rowNames = FALSE)                       |   -- |----+---+------+----|                */
/*                               | saveWorkbook(                             |  [WANT                                     */
/*                               |   wb                                      |                                            */
/*                               |  ,"d:/xls/routput.xlsx"                   |                                            */
/*                               |  ,overwrite = TRUE)                       |                                            */
/*                               | ;;;;                                      |                                            */
/*                               | %utl_rendx;                               |                                            */
/*                               |                                           |                                            */
/*                               |----------------------------------------------------------------------------------------*/
/*                               | 2 PYTHON OPENXL                           |                                            */
/*                               | ===============                           |                                            */
/*                               |                                           |                                            */
/*                               | %utlfkil(d:/xls/pyoutput.xlsx);           |  d:/xls/pyoutput.xlsx                      */
/*                               |                                           |                                            */
/*                               | %utl_pybeginx;                            |  --------------------+                     */
/*                               | parmcards4;                               |  | A1| fx     |      |                     */
/*                               | import pandas as pd                       |  -------------------------+                */
/*                               | import numpy as np                        |  [_] |  A | B |   C  |  D |                */
/*                               | import pyreadstat as ps                   |  -------------------------|                */
/*                               | from openpyxl import Workbook             |   1  |    |   |      |    |                */
/*                               | from openpyxl.utils.dataframe \           |   -- |----+---+------+----|                */
/*                               |  import dataframe_to_rows                 |   2  |    |   |      |    |                */
/*                               | have,meta = ps.read_sas7bdat( \           |   -- |----+---+------+----|                */
/*                               |   'd:/sd1/have.sas7bdat');                |   3  |    |   |CV_SUM| 20 |                */
/*                               | print(have)                               |   -- |----+---+------+----|                */
/*                               | wb = Workbook()                           |   4  |    |   |RS_SUM| 30 |                */
/*                               | ws = wb.create_sheet("want")              |   -- |----+---+------+----|                */
/*                               | for r_idx, row in enumerate(              |   5  |    |   |IP_SUM| 10 |                */
/*                               |    dataframe_to_rows(                     |   -- |----+---+------+----|                */
/*                               |       have                                |   6  |    |   |      |    |                */
/*                               |      ,index=False                         |   -- |----+---+------+----|                */
/*                               |      ,header=False)                       |   7  |    |   |      |    |                */
/*                               |      ,start=3):                           |   -- |----+---+------+----|                */
/*                               |     for c_idx \                           |  [WANT                                     */
/*                               |       ,value in enumerate(row, start=3):  |                                            */
/*                               |       ws.cell(                            |                                            */
/*                               |         row=r_idx                         |                                            */
/*                               |        ,column=c_idx                      |                                            */
/*                               |        ,value=value)                      |                                            */
/*                               | wb.save("d:/xls/pyoutput.xlsx")           |                                            */
/*                               | ;;;;                                      |                                            */
/*                               | %utl_pyendx;                              |                                            */
/**************************************************************************************************************************/

/*                   _
(_)_ __  _ __  _   _| |_
| | `_ \| `_ \| | | | __|
| | | | | |_) | |_| | |_
|_|_| |_| .__/ \__,_|\__|
        |_|
*/

options
 validvarname=upcase;
libname sd1 "d:/sd1";
data sd1.have;
 input key$ val;
cards4;
IP_SUM 10
CV_SUM 20
RS_SUM 30
;;;;
run;quit;

/************************************************************************************************************************/
/*  KEY      VAL                                                                                                        */
/* IP_SUM     10                                                                                                        */
/* CV_SUM     20                                                                                                        */
/* RS_SUM     30                                                                                                        */
/* IP_SUM     10                                                                                                        */
/* CV_SUM     20                                                                                                        */
/* RS_SUM     30                                                                                                        */
/************************************************************************************************************************/

/*                                      _
/ |  _ __    ___  _ __   ___ _ __ __  _| |_____  __
| | | `__|  / _ \| `_ \ / _ \ `_ \\ \/ / / __\ \/ /
| | | |    | (_) | |_) |  __/ | | |>  <| \__ \>  <
|_| |_|     \___/| .__/ \___|_| |_/_/\_\_|___/_/\_\
                 |_|
*/

%utlfkil(d:/xls/routput.xlsx);

%utl_rbeginx;
parmcards4;
library(haven)
library(openxlsx)
wb <- createWorkbook()
addWorksheet(wb, "want")
have<-read_sas(
 "d:/sd1/have.sas7bdat")
print(have)
writeData(wb,
  sheet = "want",
  x = have,
  startCol = 3,
  startRow = 3,
  colNames = FALSE,
  rowNames = FALSE)
saveWorkbook(
  wb
 ,"d:/xls/routput.xlsx"
 ,overwrite = TRUE)
;;;;
%utl_rendx;

/**************************************************************************************************************************/
/* --------------------+                                                                                                  */
/* | A1| fx     |      |                                                                                                  */
/* -------------------------+                                                                                             */
/* [_] |  A | B |   C  |  D |                                                                                             */
/* -------------------------|                                                                                             */
/*  1  |    |   |      |    |                                                                                             */
/*  -- |----+---+------+----|                                                                                             */
/*  2  |    |   |      |    |                                                                                             */
/*  -- |----+---+------+----|                                                                                             */
/*  3  |    |   |CV_SUM| 20 |                                                                                             */
/*  -- |----+---+------+----|                                                                                             */
/*  4  |    |   |RS_SUM| 30 |                                                                                             */
/*  -- |----+---+------+----|                                                                                             */
/*  5  |    |   |IP_SUM| 10 |                                                                                             */
/*  -- |----+---+------+ ---|                                                                                             */
/*  6  |    |   |      |    |                                                                                             */
/*  -- |----+---+------+----|                                                                                             */
/*  7  |    |   |      |    |                                                                                             */
/*  -- |----+---+------+----|                                                                                             */
/* [WANT]                                                                                                                 */
/**************************************************************************************************************************/

/*___                _   _                                               _
|___ \   _ __  _   _| |_| |__   ___  _ __     ___  _ __   ___ _ __ __  _| |
  __) | | `_ \| | | | __| `_ \ / _ \| `_ \   / _ \| `_ \ / _ \ `_ \\ \/ / |
 / __/  | |_) | |_| | |_| | | | (_) | | | | | (_) | |_) |  __/ | | |>  <| |
|_____| | .__/ \__, |\__|_| |_|\___/|_| |_|  \___/| .__/ \___|_| |_/_/\_\_|
        |_|    |___/                              |_|
*/

%utlfkil(d:/xls/pyoutput.xlsx);

%utl_pybeginx;
parmcards4;
import pandas as pd
import numpy as np
import pyreadstat as ps
from openpyxl import Workbook
from openpyxl.utils.dataframe \
 import dataframe_to_rows
have,meta = ps.read_sas7bdat( \
  'd:/sd1/have.sas7bdat');
print(have)
wb = Workbook()
ws = wb.create_sheet("want")
for r_idx, row in enumerate(
   dataframe_to_rows(
      have
     ,index=False
     ,header=False)
     ,start=3):
    for c_idx \
      ,value in enumerate(row, start=3):
      ws.cell(
        row=r_idx
       ,column=c_idx
       ,value=value)
wb.save("d:/xls/pyoutput.xlsx")
;;;;
%utl_pyendx;


/**************************************************************************************************************************/
/* --------------------+                                                                                                  */
/* | A1| fx     |      |                                                                                                  */
/* -------------------------+                                                                                             */
/* [_] |  A | B |   C  |  D |                                                                                             */
/* -------------------------|                                                                                             */
/*  1  |    |   |      |    |                                                                                             */
/*  -- |----+---+------+----|                                                                                             */
/*  2  |    |   |      |    |                                                                                             */
/*  -- |----+---+------+----|                                                                                             */
/*  3  |    |   |CV_SUM| 20 |                                                                                             */
/*  -- |----+---+------+----|                                                                                             */
/*  4  |    |   |RS_SUM| 30 |                                                                                             */
/*  -- |----+---+------+----|                                                                                             */
/*  5  |    |   |IP_SUM| 10 |                                                                                             */
/*  -- |----+---+------+ ---|                                                                                             */
/*  6  |    |   |      |    |                                                                                             */
/*  -- |----+---+------+----|                                                                                             */
/*  7  |    |   |      |    |                                                                                             */
/*  -- |----+---+------+----|                                                                                             */
/* [WANT]                                                                                                                 */
/**************************************************************************************************************************/

/*              _
  ___ _ __   __| |
 / _ \ `_ \ / _` |
|  __/ | | | (_| |
 \___|_| |_|\__,_|

*/
