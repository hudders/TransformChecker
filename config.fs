module config

let nodeList = [3..200]

let dloadFolder = @"C:\Users\mhudson\Downloads"
let crrntFolder = @"C:\x_FSharpStuff\FSharp Test Projects\TransformChecker"

let risk_csv = @"..\..\..\..\x_ GIT repos\motor_data_driven\motor.canopy.data_driven\risks.csv"
let additional_drivers_csv = @"..\..\..\..\x_ GIT repos\motor_data_driven\motor.canopy.data_driven\additional_drivers.csv"
let claims_csv = @"..\..\..\..\x_ GIT repos\motor_data_driven\motor.canopy.data_driven\claims.csv"
let convictions_csv = @"..\..\..\..\x_ GIT repos\motor_data_driven\motor.canopy.data_driven\convictions.csv"
let vehicle_details_csv = @"..\..\..\..\x_ GIT repos\motor_data_driven\motor.canopy.data_driven\vehicle_details.csv"
let car_usage_csv = @"..\..\..\..\x_ GIT repos\motor_data_driven\motor.canopy.data_driven\car_usage.csv"

let contains lookFor inSeq = Seq.exists (fun elem -> elem = lookFor) inSeq