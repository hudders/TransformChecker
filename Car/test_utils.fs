module test_utils

open System
open System.Text.RegularExpressions
open Microsoft.Office.Interop

open csvLoader
open config
open xlsLoader

let stripChars text (chars:string) =
    Array.fold (fun (s:string) c -> s.Replace(c.ToString(),"")) text (chars.ToCharArray())

let (|Prefix|_|) (p:string) (s:string) =
    if s.StartsWith(p) then
        Some(s.Substring(p.Length))
    else
        None

let todaysDate = (System.DateTime.Now.ToString("dd/MM/yyyy")).Split('/')
let todaysDay, todaysMonth, todaysYear = todaysDate.[0], todaysDate.[1], todaysDate.[2]
let passCell, failCell, resultCell = "G","H","J"

let mutable fillCell =
        (fun (xlFile : Excel.Worksheet, xlColumn : string, xlRow : int, value : string) -> xlFile.Range((xlColumn + xlRow.ToString()),(xlColumn + xlRow.ToString())).Value2 <- value)

let dateTester (dateFormat : string, day : string, month : string, year : string) =
    let date =
        if month <> null && month.Length > 0 then
            if day.Length > 0 then
                day + "/" + month + "/" + year
            else
                "01" + "/" + month + "/" + year
        else
            todaysDay.ToString() + "/" + todaysMonth.ToString() + "/" + year
    let parsedDate = DateTime.Parse date
    let dateFormat = Regex.Replace(dateFormat.ToLower(),"mm","MM")
    parsedDate.ToString(dateFormat)

let matchToExpected(dataSrc : Excel.Worksheet, acVal : string, exVal : string, exLoc : string, xlRow : int) =
    let acVal, exVal, exLoc, eCol = acVal.Replace("–","-"), exVal.Replace("–","-"), exLoc.Replace("–","-"), cellValue(dataSrc,"E",xlRow).ToString().Replace("–","-")
    let lineNumber =
        let rec loop n =
            if eCol.Split('\n').[n] = exLoc || eCol.Split('\n').[n] = exLoc.Split('[').[0] then
                n
            else
                loop (n + 1)
        loop 0
    let rec DataExtractionLoop row =
        if cellValue(dataExt, "B", row) = null then
            fillCell(dataExt, "C", row, lineNumber.ToString())
            let runNumber =
                if cellValue(dataExt, "C", row).ToString() = cellValue(dataExt, "C", row - 1).ToString() then
                    let x = cellValue(dataExt, "B", row - 1).ToString()
                    (System.Convert.ToInt32(x:string) + 1).ToString()
                else
                    "1"
            let matchResult =
                if acVal = exVal then
                    "Y"
                else
                    "N"
            fillCell(dataExt, "B", row, runNumber)
            fillCell(dataExt, "D", row, exVal + " in " + exLoc)
            fillCell(dataExt, "E", row, acVal)
            fillCell(dataExt, "F", row, matchResult)
            let rec DataMergeLoop n =
                if cellValue(dataExt, "H", n) = null || cellValue(dataExt, "H", n).ToString() = (cellValue(dataExt, "B", row).ToString()) then
                    fillCell(dataExt, "H", n, (cellValue(dataExt, "B", row).ToString()))
                    let dataFill(xCol : string, mCol : string) =
                        if cellValue(dataExt, mCol, n) = null then
                            cellValue(dataExt, xCol, row).ToString()
                        else
                            cellValue(dataExt, mCol, n).ToString() + "\n" + cellValue(dataExt, xCol, row).ToString()
                    fillCell(dataExt, "I", n, dataFill("E","I"))
                    fillCell(dataExt, "J", n, dataFill("F","J"))
                    fillCell(dataExt, "L", n, dataFill("D","L"))
                else
                    DataMergeLoop (n + 1)
                if contains "N" ((cellValue(dataExt, "J", n).ToString()).Split('\n')) then
                    fillCell(dataExt, "K", n, "Fail")
                else
                    fillCell(dataExt, "K", n, "Pass")
            DataMergeLoop 59
        else
            DataExtractionLoop (row + 1)
    DataExtractionLoop 59

let cleanUp() =
    let rec cleanMeLoop n =
        fillCell(dataExt, "B", n, null)
        fillCell(dataExt, "C", n, null)
        fillCell(dataExt, "D", n, null)
        fillCell(dataExt, "E", n, null)
        fillCell(dataExt, "F", n, null)
        fillCell(dataExt, "H", n, null)
        fillCell(dataExt, "I", n, null)
        fillCell(dataExt, "J", n, null)
        fillCell(dataExt, "K", n, null)
        fillCell(dataExt, "L", n, null)
        if cellValue(dataExt, "B", n + 1) <> null then
            cleanMeLoop (n + 1)
    cleanMeLoop 59

let deleteExt(path : string, extension : string) =
    for file in System.IO.Directory.EnumerateFiles(path + "/", "*." + extension) do
        let tempPath = System.IO.Path.Combine(dloadFolder, file)
        //printf "%s" tempPath
        System.IO.File.Delete(tempPath)

// howManySince
// Returns number of X between two dates
let howManySince(X : string, date2 : string) =
    let date1 = System.DateTime.Today
    let date2 = System.DateTime.Parse(date2)
    match X with
        | "Y"   -> (date1.Year - date2.Year).ToString()
        | "M"   -> (((date1.Year - date2.Year)*12) + (date1.Month - date2.Month)).ToString()
        | _     -> ""


//codeConvert
// Experimental function to process "<occCode>" and "<empCode>" values / "as input" values - eventually should replace dataLoader.fs
let codeConvert (dataType : string, fName : string, value : string, journeyNumber : string) =
    let j = toInt(journeyNumber)-1
    match value with
        | "<occCode>"                 -> match dataType with
                                             | "Proposer"   -> value.Replace(value, (codeLookup(personCollection.[j].OccupationTitleDescription, value).Value))
                                             | _            -> value.Replace(value, (codeLookup(additionalCollection.[j].OccupationTitleDescription, value).Value))
        | "<empCode>" | "<busCode>"   -> match dataType with
                                             | "Proposer"   -> value.Replace(value, (codeLookup(personCollection.[j].BusinessTypeDescription, "<empCode>").Value))
                                             | _            -> value.Replace(value, (codeLookup(additionalCollection.[j].BusinessTypeDescription, "<empCode>").Value))
        | "<MM>"                      -> match dataType, fName with
                                             | "Proposer", "Period licence held for?"    -> howManySince("M",dateTester("dd/mm/yyyy",personCollection.[j].LicenceDateDay,personCollection.[j].LicenceDateMonth,(System.DateTime.Today.Year - (toInt(personCollection.[j].LicenceHeldCode) - 1)).ToString()))
                                             | "Proposer", "Permanent UK Resident Since" -> howManySince("M",dateTester("dd/mm/yyyy",personCollection.[j].IsLivingInUkSinceDay,personCollection.[j].IsLivingInUkSinceMonthCode,personCollection.[j].IsLivingInUkSinceYear))
                                             | _                                         -> ""
        | "<YY>"                      -> match dataType, fName with
                                             | "Proposer", "Period licence held for?"    -> howManySince("Y",dateTester("dd/mm/yyyy",personCollection.[j].LicenceDateDay,personCollection.[j].LicenceDateMonth,(System.DateTime.Today.Year - (toInt(personCollection.[j].LicenceHeldCode) - 1)).ToString()))
                                             | "Proposer", "Permanent UK Resident Since" -> howManySince("Y",dateTester("dd/mm/yyyy",personCollection.[j].IsLivingInUkSinceDay,personCollection.[j].IsLivingInUkSinceMonthCode,personCollection.[j].IsLivingInUkSinceYear))
                                             | _                                         -> ""
        | "dd/mm/yyyy" | "01/mm/yyyy" -> match dataType, fName with
                                             | "Proposer", "Period licence held for?"    -> dateTester(value,personCollection.[j].LicenceDateDay,personCollection.[j].LicenceDateMonth,(System.DateTime.Today.Year - (toInt(personCollection.[j].LicenceHeldCode) - 1)).ToString())
                                             | _                                         -> ""
        | _ -> value


//"Proposer", "Permanent UK Resident Since" -> let lastOfMonth = (DateTime(personCollection.[j].IsLivingInUkSinceYear, personCollection.[j].IsLivingInUkSinceMonthCode, 1).AddMonths(1).AddDays(-1)).Day.ToString()
//howManySince("M",dateTester("dd/mm/yyyy","1",personCollection.[j].IsLivingInUkSinceMonthCode,personCollection.[j].IsLivingInUkSinceYear)
