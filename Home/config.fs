module config

let nodeList  = [3..200]
let startCell = 67

let dloadFolder = @"C:\Users\mhudson\Downloads\"
let savedFolder = @"C:\Dev\panel.testplans\Home\"
let crrntFolder = @"C:\x_FSharpStuff\FSharp Test Projects\TransformChecker\Home\"
let xlsRtFolder = @"C:\Dev\panel.testplans\"
let jsnRtFolder = @"C:\Dev\HomeTransformChecker\HomeTransformChecker\ref\json\"

let contains lookFor inSeq = Seq.exists (fun elem -> elem = lookFor) inSeq

let (|Prefix|_|) (p:string) (s:string) =
    if s.StartsWith(p) then
        Some(s.Substring(p.Length))
    else
        None