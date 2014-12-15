module config

open System.Text.RegularExpressions

let nodeList  = [3..200]
let startCell = 67

let dloadFolder = @"C:\Users\mhudson\Downloads\"
let savedFolder = @"C:\Dev\panel.testplans\Home\"
let crrntFolder = @"C:\Dev\TransformChecker\Home\"
let xlsRtFolder = @"C:\Dev\TransformChecker\TestPlans\"
let jsnRtFolder = crrntFolder + @"ref\json\"

let contains lookFor inSeq = Seq.exists (fun elem -> elem = lookFor) inSeq

let (|Prefix|_|) (p:string) (s:string) =
    if s.StartsWith(p) then
        Some(s.Substring(p.Length))
    else
        None

let (|RegExParse|_|) pattern input =
    let m = Regex.Match(input, pattern) in
    if m.Success then Some (List.tail [ for g in m.Groups -> g.Value ]) else None