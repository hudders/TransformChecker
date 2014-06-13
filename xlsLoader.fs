module xlsLoader

open System
open Microsoft.Office.Interop

open config

let xlsPath = @"C:\x_FSharpStuff\_dat\Panel Test PC Document V5.xlsx"

let xlApp = new Excel.ApplicationClass()
let xlWorkBookInput = xlApp.Workbooks.Open(xlsPath)
xlApp.Visible <- true

let getXLS(tab : string) =
    xlWorkBookInput.Worksheets.[tab] :?> Excel.Worksheet 

let dataExt = getXLS("Test Summary")

let cellValue (xlsFile : Excel.Worksheet, column : string, row : int) =
    xlsFile.Range(column + row.ToString()).Value2

let brand =
  let brandFile = getXLS("Test Summary")
  cellValue(brandFile, "C", 2).ToString()

// Brand group sequences
// ---------------------
let brandGroup(brand : string) =
    let cdlBrandFilterFreeSeq = 
        seq ["Hastings Direct";
             "Hastings Essentials";
             "igo4insurance";
             "One Quote";
             "Aquote";
             "Elite";
             "Virgin Money PC";
             "Esure Broker";
             "Sheilas' Wheels Broker";
             "Hastings Premier";
             "Hastings People's Choice";
             "Only Young Drivers";
             "Saga Select";
             "Hastings Direct SmartMiles";
             "Castle Cover";
             "Sure Thing!";
             "John Lewis Insurance"]

    let cdlBrandFilteredSeq =
        seq ["AutoNet";
             "Direct Choice"]

    let opengiBrandSeq =
        seq ["Express";
             "Octagon"]

    let sspBrandSeq =
        seq ["Drivology"]


    if contains brand cdlBrandFilterFreeSeq then
        "CDL-FilterFree"
    elif contains brand cdlBrandFilteredSeq then
        "AutoNet"
    elif contains brand opengiBrandSeq then
        "Open-GI"
    elif contains brand sspBrandSeq then
        "SSP Multi Quote"
    else
        brand
// ---------------------

let codeLookup(description : string, codeType : string) =
    let xlsFile, lowerBound, upperBound =
        if codeType = "<occCode>" then
            getXLS("occupation codes"), 2, 1962
        elif codeType = "<empCode>" then
            getXLS("business codes"), 2, 939
        elif codeType = "<conCode>" then
            getXLS("conviction codes"), 2, 88
        else
            getXLS("car codes"), 2, 3           
    let result = ref None in
        let rec loop n =
            if n <= upperBound then
                let theCode =
                    if cellValue(xlsFile, "D", n) = null || codeType = "<carCode>" then
                        cellValue(xlsFile, "A", n).ToString()
                    elif codeType = "<styleCode>" then
                        cellValue(xlsFile, "F", n).ToString()
                    else
                        cellValue(xlsFile, "D", n).ToString()
                let theDesc = (cellValue(xlsFile, "B", n).ToString())
                if description <> theDesc then
                    loop (n + 1)
                else
                    result := Some theCode
        loop lowerBound
    !result

let carDetailLookup(registration : string) =
    let xlsFile = getXLS("car codes")
    let rec loop n =
        let reg = cellValue(xlsFile, "B", n).ToString()
        let resultList =
            if reg = registration then
                [cellValue(xlsFile, "A", n).ToString();
                cellValue(xlsFile, "C", n).ToString();
                cellValue(xlsFile, "D", n).ToString();
                cellValue(xlsFile, "E", n).ToString();
                cellValue(xlsFile, "F", n).ToString();
                cellValue(xlsFile, "G", n).ToString();
                cellValue(xlsFile, "H", n).ToString();
                cellValue(xlsFile, "I", n).ToString()]
            else
                if cellValue(xlsFile, "A", n+1) <> null then
                    loop (n + 1)
                else
                    ["";"";"";"";"";"";"";""]
        resultList
    loop 2