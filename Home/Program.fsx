#r @"ref\FSharp.Data.dll"
#r @"ref\WebDriver.dll"
#r @"ref\Newtonsoft.Json.dll"
#r @"ref\SizSelCsZzz.dll"
#r @"ref\canopy.dll"

#r "System.Xml.dll"
#r "System.Xml.Linq.dll"
#r "Microsoft.Office.Interop.Excel"
#r "office"

open FSharp.Data
open System
open System.IO
open System.Collections.Generic

#load "config.fs"
#load "xlsLoader.fs"
#load "jsnLoader.fs"
#load "test_utils.fs"
#load "xmlLoader.fs"
#load "dataLoader.fs"

open config
open xlsLoader
open jsnLoader
open test_utils
open xmlLoader
open dataLoader

riskLoad()
itemCollectionLoad()
claimLoad()

let loadTests(testType : string, environment : string, journeyNumber : int) =
    let rec loop n =
        let pid = risk.[n].testID
        if journeyNumber = pid || journeyNumber = 0 then
            let startTime = System.DateTime.Now
            printfn ""
            printfn "---------------------------------------------------------"
            printfn "             %s TESTS FOR JOURNEY %s" testType (risk.[n].testID.ToString())
            printfn "                    %s" (startTime.ToString("HH:mm:ss"))
            printfn "---------------------------------------------------------"
            let xmlFile = (savedFolder + "/" + brand + "/xml/" + brandCode + "_" + startTime.ToString("ddMMyyyy") + "_" + (pid.ToString()) + ".xml").ToString()
            if not(File.Exists(xmlFile)) then
                getXML(environment, productX, brand, risk.[n].lastName, risk.[n].emailAddress, xmlFile)
            if File.Exists(xmlFile) then
                if testType = "PROPOSER" || testType = "ALL" then
                    loadData(productX, "Proposer", xmlFile, (pid - 1))
                if testType = "JOINT" || testType = "ALL" then
                    loadData(productX, "Joint Proposer", xmlFile, (pid - 1))
                if testType = "PROPERTY" || testType = "ALL" then
                    loadData(productX, "Property", xmlFile, (pid - 1))
                if testType = "CONTENTS" || testType = "ALL" then
                    loadData(productX, "Contents Cover", xmlFile, (pid - 1))
                if testType = "BUILDINGS" || testType = "ALL" then
                    loadData(productX, "Buildings Cover", xmlFile, (pid - 1))
                if testType = "LOCKS" || testType = "ALL" then
                    loadData(productX, "Locks and Security", xmlFile, (pid - 1))
                if testType = "CLAIMS" || testType = "ALL" then
                    loadData(productX, "Claims", xmlFile, (pid - 1))
                if testType = "PRICES" || testType = "ALL" then
                    loadData(productX, "Price Page", xmlFile, (pid - 1))
            let endTime = System.DateTime.Now
            let elapsed = endTime.Subtract(startTime)
            printfn ""
            printfn ""
            printfn "Time elapsed: %s.%s secs" ((elapsed.Seconds).ToString()) ((elapsed.Milliseconds).ToString())
            printfn "---------------------------------------------------------"
        if risk.Count > (n + 1) then
            loop (n + 1)
    loop 0

loadTests("ALL","UAT",0)