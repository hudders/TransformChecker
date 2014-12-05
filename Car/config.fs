module config

open System.Text.RegularExpressions

let nodeList = [3..200]

let dloadFolder = @"C:\Users\mhudson\Downloads\"
let savedFolder = @"C:\Dev\panel.testplans\Car\"
//let crrntFolder = @"C:\x_FSharpStuff\FSharp Test Projects\TransformChecker\Car\"
let crrntFolder = @"C:\Dev\TransformChecker\Car\"
let xlsRtFolder = @"C:\Dev\panel.testplans\"
let csvRtFolder = @"C:\Dev\motor_data_driven\motor.canopy.data_driven\"

let risk_csv = csvRtFolder + "risks.csv"
let additional_drivers_csv = csvRtFolder + "additional_drivers.csv"
let claims_csv = csvRtFolder + "claims.csv"
let convictions_csv = csvRtFolder + "convictions.csv"
let vehicle_details_csv = csvRtFolder + "vehicle_details.csv"
let car_usage_csv = csvRtFolder + "car_usage.csv"
let addresses_csv = csvRtFolder + "addresses.csv"

let contains lookFor inSeq = Seq.exists (fun elem -> elem = lookFor) inSeq

printf "Enter name of brand: "
let brandX = System.Console.ReadLine()
let productX = "PC"

let (|RegExParse|_|) pattern input =
    let m = Regex.Match(input, pattern) in
    if m.Success then Some (List.tail [ for g in m.Groups -> g.Value ]) else None

let toInt (str : string) =
    System.Convert.ToInt32(str)