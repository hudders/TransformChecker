module config

let nodeList = [3..200]

let dloadFolder = @"C:\Users\mhudson\Downloads\"
let savedFolder = @"C:\Dev\panel.testplans\Car\"
let crrntFolder = @"C:\x_FSharpStuff\FSharp Test Projects\TransformChecker\"
let xlsRtFolder = @"C:\Dev\panel.testplans\"
let csvRtFolder = @"C:\Dev\motor_data_driven\motor.canopy.data_driven\"

let risk_csv = csvRtFolder + "risks.csv"
let additional_drivers_csv = csvRtFolder + "additional_drivers.csv"
let claims_csv = csvRtFolder + "claims.csv"
let convictions_csv = csvRtFolder + "convictions.csv"
let vehicle_details_csv = csvRtFolder + "vehicle_details.csv"
let car_usage_csv = csvRtFolder + "car_usage.csv"

let contains lookFor inSeq = Seq.exists (fun elem -> elem = lookFor) inSeq