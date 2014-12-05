module dataLoader

open config
open test_utils
open jsnLoader

let rowList = [3..200]

let loadData (product : string, dataType : string, xmlFile : string, a : int) =
    let dataSrc =
        match product, dataType with
            | "PC", "Proposer" | "HH", "Proposer" -> xlsLoader.getXLS("2.1 " + dataType)
            | "PC", "Additional Driver" | "HH", "Joint Proposer" -> xlsLoader.getXLS("2.2 " + dataType)
            | "PC", "Claim" | "HH", "Property" -> xlsLoader.getXLS("2.3 " + dataType)
            | "PC", "Conviction" | "HH", "Contents Cover" -> xlsLoader.getXLS("2.4 " + dataType)
            | "PC", "Vehicle" | "HH", "Buildings Cover" -> xlsLoader.getXLS("2.5 " + dataType)
            | "PC", "Policy" | "HH", "Locks and Security" -> xlsLoader.getXLS("2.6 " + dataType)
            | "HH", "Claims" -> xlsLoader.getXLS("2.7 " + dataType)
            | "HH", "Price Page" -> xlsLoader.getXLS("2.8 " + dataType)
            | _ -> xlsLoader.getXLS("")

    for xlRow in rowList do
        if xlsLoader.cellValue(dataSrc, "A", xlRow) <> null && xlsLoader.cellValue(dataSrc, "A", xlRow).ToString() = (risk.[a].testID).ToString() then     
            let xVal = xlsLoader.cellValue(dataSrc, "D", xlRow)
            let xLoc = xlsLoader.cellValue(dataSrc, "E", xlRow)
            if xLoc = null || xVal = null then
                dataSrc.Range(("I" + xlRow.ToString()), ("I" + xlRow.ToString())).Value2 <- "Not Applicable"
            else
                let xLoc, xVal = (xLoc.ToString()).Split('\n'), (xVal.ToString()).Split('\n')
                let rec loop n =
                    xmlLoader.checkXml (xVal.[n], xLoc.[n], xmlFile, dataSrc, xlRow, dataType, n)
                    if xLoc.Length <> n + 1 then
                        loop (n + 1)
                loop 0
                let rec ResultLoop n =
                    if xlsLoader.cellValue(xlsLoader.dataExt, "K", n) <> null then
                        if xlsLoader.cellValue(xlsLoader.dataExt, "K", n).ToString() = "Pass" then
                            fillCell(dataSrc, "G", xlRow, xlsLoader.cellValue(xlsLoader.dataExt, "K", n).ToString())
                            fillCell(dataSrc, "H", xlRow, "")
                            fillCell(dataSrc, "J", xlRow, xlsLoader.cellValue(xlsLoader.dataExt, "I", n).ToString())
                            if xlsLoader.cellValue(xlsLoader.dataExt, "L", n) <> null then
                                fillCell(dataSrc, "K", xlRow, xlsLoader.cellValue(xlsLoader.dataExt, "L", n).ToString())
                        else
                            if xlsLoader.cellValue(dataSrc, "G", xlRow) = null then
                                fillCell(dataSrc, "H", xlRow, xlsLoader.cellValue(xlsLoader.dataExt, "K", n).ToString())
                                fillCell(dataSrc, "J", xlRow, xlsLoader.cellValue(xlsLoader.dataExt, "I", n).ToString())
                                if xlsLoader.cellValue(xlsLoader.dataExt, "L", n) <> null then
                                    fillCell(dataSrc, "K", xlRow, xlsLoader.cellValue(xlsLoader.dataExt, "L", n).ToString())
                        if xlsLoader.cellValue(xlsLoader.dataExt, "K", n).ToString() <> "Pass" then
                            ResultLoop (n + 1)
                ResultLoop startCell
            let resultThing =
                if xlsLoader.cellValue(dataSrc, "J", xlRow) <> null then
                    if xlsLoader.cellValue(dataSrc, "G", xlRow) <> null then
                        "."
                    else
                        "F"
                else
                    "S"
            printf "%s" resultThing
            cleanUp()