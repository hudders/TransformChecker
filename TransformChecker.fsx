﻿#r @"ref\FSharp.Data.dll"
#r @"ref\WebDriver.dll"
#r @"ref\Newtonsoft.Json.dll"
#r @"ref\SizSelCsZzz.dll"
#r @"ref\canopy.dll"

#r "System.Xml.dll"
#r "System.Xml.Linq.dll"
#r "Microsoft.Office.Interop.Excel"
#r "office"

#load "config.fs"
#load "xlsLoader.fs"
#load "test_utils.fs"
#load "csvLoader.fs"
#load "xmlLoader.fs"
#load "dataLoader.fs"

open System
open System.IO
open System.Text.RegularExpressions
open Microsoft.Office.Interop

open config

open csvLoader
open xlsLoader
open test_utils
open xmlLoader
open dataLoader

System.IO.Directory.SetCurrentDirectory("C:\x_FSharpStuff\FSharp Test Projects\TransformChecker")

personLoad(risk_csv)
additionalLoad(additional_drivers_csv)
claimLoad(claims_csv)
convictionLoad(convictions_csv)
vehicleLoad(vehicle_details_csv)
carUsageLoad(car_usage_csv)

let loadClaim (xmlFile : string, claimId : string, pid : int) =
    let rec loop id =
        if claimCollection.[id].ClaimsId = (stripChars claimId " ") then
            loadData("Claim", xmlFile, id, pid, pid)
        else
            if claimCollection.[id + 1].ClaimsId.Length > 0 then
                loop (id + 1)
    loop 0

let loadConviction (xmlFile : string, convictionId : string, pid : int) =
    let rec loop id =
        if convictionCollection.[id].ConvictionId = (stripChars convictionId " ") then
            loadData("Conviction", xmlFile, id, pid, pid)
        else
            if convictionCollection.[id + 1].ConvictionId.Length > 0 then
                loop (id + 1)
    loop 0

let loadAdditional (xmlFile : string, driverId : string, pid : int) =
    let rec loop id =
        if additionalCollection.[id].AdditionalDriverId = (stripChars driverId " ") then
            loadData("Additional", xmlFile, id, pid, pid)
        else
            if additionalCollection.[id + 1].AdditionalDriverId.Length > 0 then
                loop (id + 1)
    loop 0

let loadVehicle (xmlFile : string, vehicleId : string, vehicleUsage : string, pid : int) =
    let rec vehicleLoop vid =
        if vehicleCollection.[vid].VehicleId = (stripChars vehicleId " ") then
            let rec usageLoop uid =
                if carUsageCollection.[uid].CarUsageId = (stripChars vehicleUsage " ") then
                    loadData("Vehicle", xmlFile, vid, uid, pid)
                else
                    if carUsageCollection.[uid + 1].CarUsageId.Length > 0 then
                        usageLoop (uid + 1)
            usageLoop 0
        else
            if vehicleCollection.[vid + 1].VehicleId.Length > 0 then
                vehicleLoop (vid + 1)
    vehicleLoop 0

let loadProposer (xmlFile : string, propId : string) =
    let rec loop id =
        if personCollection.[id].TestId = (stripChars propId " ") then
            loadData("Proposer", xmlFile, id, id, id)
        else
            if personCollection.[id + 1].TestId.Length > 0 then
                loop (id + 1)
    loop 0

let loadTests (testType : string) =
    //System.Console.Clear()
    for proposer in personCollection do
        if proposer.Execution = "Yes" then
            let pid = personCollection.IndexOf(proposer)
            let linkID = getXML(proposer.Environment, brand, proposer.LastName, proposer.Email, proposer.VehicleId)
            let xmlFile = ("C:\Users\mhudson\Downloads\\" + "Request_" + brandGroup(brand) + "_After_" + linkID.[3] + "_2.xml").ToString()
            printfn "---------------------------------------------------------"
            printfn "            %s TESTS FOR JOURNEY %s" testType proposer.TestId
            printfn "---------------------------------------------------------"
            if testType = "PROPOSER" || testType = "ALL" then
                loadProposer(xmlFile, proposer.TestId)
            if testType = "ADDITIONAL" || testType = "ALL" then
                if proposer.AdditionalDriverIds.Trim().Length > 0 then
                    let list = proposer.AdditionalDriverIds.Split(',')
                    for id in list do
                        loadAdditional(xmlFile, id, pid)
            if testType = "CLAIMS" || testType = "PROPOSER" || testType = "ADDITIONAL" || testType = "ALL" then
                if testType <> "ADDITIONAL" then
                    if proposer.ClaimIds.Trim().Length > 0 then
                        let list = proposer.ClaimIds.Split(',')
                        for id in list do
                            loadClaim(xmlFile, id, pid)
                if testType <> "PROPOSER" then
                    if proposer.AdditionalDriverIds.Trim().Length > 0 then
                        let driverList = proposer.AdditionalDriverIds.Split(',')
                        for driverId in driverList do
                            let rec loop id =
                                if additionalCollection.[id].AdditionalDriverId = (stripChars driverId " ") then
                                    if additionalCollection.[id].ClaimIds.Trim().Length > 0 then
                                        let list = additionalCollection.[id].ClaimIds.Split(',')
                                        for claimId in list do
                                            loadClaim(xmlFile, claimId, pid)
                                else
                                    if additionalCollection.[id + 1].AdditionalDriverId.Length > 0 then
                                        loop (id + 1)
                            loop 0
            if testType = "CONVICTIONS" || testType = "PROPOSER" || testType = "ADDITIONAL" || testType = "ALL" then
                if testType <> "ADDITIONAL" then
                    if proposer.ConvictionIds.Trim().Length > 0 then
                        let list = proposer.ConvictionIds.Split(',')
                        for id in list do
                            loadConviction(xmlFile, id, pid)
                if testType <> "PROPOSER" then
                    if proposer.AdditionalDriverIds.Trim().Length > 0 then
                        let driverList = proposer.AdditionalDriverIds.Split(',')
                        for driverId in driverList do
                            let rec loop id =
                                if additionalCollection.[id].AdditionalDriverId = (stripChars driverId " ") then
                                    if additionalCollection.[id].ConvictionIds.Trim().Length > 0 then
                                        let list = additionalCollection.[id].ConvictionIds.Split(',')
                                        for convictionId in list do
                                            loadConviction(xmlFile, convictionId, pid)
                                else
                                    if additionalCollection.[id + 1].AdditionalDriverId.Length > 0 then
                                        loop (id + 1)
                            loop 0
            if testType = "VEHICLE" || testType = "ALL" then
                loadVehicle(xmlFile, proposer.VehicleId, proposer.VehicleUsage, pid)
            if testType = "POLICY" || testType = "ALL" then
                loadData("Policy", xmlFile, pid, pid, pid)
            deleteExt("xml")
            deleteExt("tmp")

loadTests ("ALL")