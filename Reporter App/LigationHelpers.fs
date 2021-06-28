module LigationHelpers

open System
open OfficeOpenXml
open DocumentFormat.OpenXml
open DocumentFormat.OpenXml.Packaging
open DocumentFormat.OpenXml.Wordprocessing
open System.Linq



let csId (body : Body) paragraphIndex runIndex =
    let paragraph = body.Elements<Paragraph>().ElementAt(paragraphIndex)
    let run = paragraph.Elements<Run>().ElementAt(runIndex)
    let runProperties = run.Elements<RunProperties>().First()
    let underline = runProperties.AppendChild<Underline>(new Underline(Val = EnumValue<UnderlineValues>DocumentFormat.OpenXml.Wordprocessing.UnderlineValues.Single))
    let position = runProperties.AppendChild<Position>(new Position(Val = StringValue("4")))
    run.Elements<RunProperties>().Equals(underline) |>ignore
    run.Elements<RunProperties>().Equals(position) |> ignore
    let text = run.AppendChild<Text>(new Text())
    text


let fillCells (body : Body) tableIndex tableRowIndex tableCellIndex paragraphIndex runIndex underlineIndex = 
    let table = body.Elements<Table>().ElementAt(tableIndex)
    let row = table.Elements<TableRow>().ElementAt(tableRowIndex)
    let cell = row.Elements<TableCell>().ElementAt(tableCellIndex)
    let paragraph = cell.Elements<Paragraph>().ElementAt(paragraphIndex)
    if not (paragraph.Elements<Run>().Any()) then
        let run = paragraph.AppendChild<Run>(new Run())
        let runProperties = run.AppendChild(new RunProperties())
        let fontSize = runProperties.AppendChild<FontSize>(new FontSize(Val = StringValue("18")))
        run.Elements<RunProperties>().Equals(fontSize) |> ignore
        let text = run.AppendChild<Text>(new Text())
        text
    else
        let run = paragraph.Elements<Run>().ElementAt(runIndex)
        let runProperties = run.AppendChild(new RunProperties())
        let fontSize = runProperties.AppendChild<FontSize>(new FontSize(Val = StringValue("18")))
        run.Elements<RunProperties>().Equals(fontSize) |> ignore
        if underlineIndex = 1 then
            let underline = runProperties.AppendChild<Underline>(new Underline(Val = EnumValue<UnderlineValues>DocumentFormat.OpenXml.Wordprocessing.UnderlineValues.Single))
            run.Elements<RunProperties>().Equals(underline) |>ignore 
            let text = run.AppendChild(new Text())
            text
        else    
            let text = run.AppendChild(new Text())
            text 

let footNoteSize (body : Body) tableIndex tableRowIndex tableCellIndex paragraphIndex runIndex = 
    let table = body.Elements<Table>().ElementAt(tableIndex)
    let row = table.Elements<TableRow>().ElementAt(tableRowIndex)
    let cell = row.Elements<TableCell>().ElementAt(tableCellIndex)
    let paragraph = cell.Elements<Paragraph>().ElementAt(paragraphIndex)
    let run = paragraph.Elements<Run>().ElementAt(runIndex)
    let runProperties = run.AppendChild(new RunProperties())
    let underline = runProperties.AppendChild<Underline>(new Underline(Val = EnumValue<UnderlineValues>DocumentFormat.OpenXml.Wordprocessing.UnderlineValues.Single))
    run.Elements<RunProperties>().Equals(underline) |>ignore
    let fontSize = runProperties.AppendChild<FontSize>(new FontSize(Val = StringValue("10")))
    run.Elements<RunProperties>().Equals(fontSize) |> ignore
    let text = run.AppendChild(new Text())
    text
    

 //Creates empty list - finds the Excel cell location for input and deconstructs the tuple into row and column numbers
let listFunction (item : string) (sheetName : ExcelWorksheet) columnIndex trimFirstCharacter =
    let list = List.init sheetName.Dimension.End.Row (fun i -> (i+1,1)) 
    if trimFirstCharacter then 
        let coordinates = List.find (fun (row,col) -> item.Equals(((string sheetName.Cells.[row,col].Value).Trim()).[1..item.Length], StringComparison.InvariantCultureIgnoreCase)) list
        let row, _colnum = coordinates
        let value = sheetName.Cells.[row,columnIndex].Value |> string
        value
    else    
        let coordinates = List.find (fun (row,col) -> item.Equals((string sheetName.Cells.[row,col].Value).Trim(), StringComparison.InvariantCultureIgnoreCase)) list
        let row, _colnum = coordinates
        let value = sheetName.Cells.[row,columnIndex].Value |> string
        value

let oligoStamp (scale : float) : (string * string* float * float * float * float * string) = 
    match scale with   
        |scale when scale <= 0.5 -> ("6", "3", 1.5, 1.7, 0.55, 0.25, "4")
        |1.0 -> ("9", "6", 3.0, 3.4, 1.1, 0.5, "8")
        |1.25 -> ("10.5", "7.5", 3.75, 4.25, 1.37, 0.63, "10")
        |1.5 -> ("12", "9", 4.5, 5.1, 1.65, 0.75, "12")
        |2.0 -> ("15", "12", 6.0, 6.8, 2.2, 1.0, "16")
        |2.5 -> ("18", "15", 7.5, 8.5, 2.75, 1.25, "20")
        |3.0 -> ("21", "18", 9.0, 10.2, 3.3, 1.5, "24")
        |3.5 -> ("24", "21", 10.5, 11.9, 3.85, 1.75, "28")
        |4.0 -> ("27", "24", 12.0, 13.6, 4.4, 2.0, "32")
        |4.5 -> ("30", "27", 13.5, 15.3, 4.95, 2.25, "36")
        |10.0 -> ("60", "60", 30.0, 34.0, 11.0, 5.0, "80")
        |_ -> failwith "Error..."


let calNote (body : Body) (inputParams : string list) (param : string) : unit =
    if inputParams.Length > 1 then 
        (footNoteSize body 1 6 1 0 4).Text <- "①"
        (footNoteSize body 1 6 1 2 3).Text <- "①"
        (footNoteSize body 1 6 1 4 4).Text <- "①"
        (footNoteSize body 1 6 1 6 4).Text <- "①"
        (footNoteSize body 1 6 1 8 2).Text <- "①"
        (footNoteSize body 1 6 1 11 2).Text <- "①"
        (footNoteSize body 1 5 1 7 3).Text <- "①"
        (footNoteSize body 1 5 1 7 9).Text <- "①"
        (footNoteSize body 1 5 1 10 2).Text <- "①"
        (footNoteSize body 1 5 1 10 7).Text <- "①"
        (footNoteSize body 1 5 1 10 11).Text <- "①"
        let lastLot = inputParams.Last()
        if param = lastLot then
            let restOfList = inputParams.[0..inputParams.Length - 2]  |> String.concat ", "
            let note = "① Calculations include " + restOfList + "."
            (fillCells body 1 14 0 2 0 0).Text <- note
       
        else
            let note = "① Calculations are on " + lastLot + "."
            (fillCells body 1 14 0 2 0 0).Text <- note


let digesting (float) : (string * string * string * string) = 
    let DEPC = (4.75 * float).ToString()
    let neBuffer = float.ToString()
    let cutterMix = (2.0 * float).ToString()
    let pviiEnzyme = (0.25 * float).ToString()
    DEPC, neBuffer, cutterMix, pviiEnzyme
    
let digestFootnote (body : Body) (inputParams : string list) = 
    if inputParams.Length > 1 then 
        (footNoteSize body 3 5 1 2 2).Text <- "①"
        (footNoteSize body 3 5 1 4 3).Text <- "①"
        (footNoteSize body 3 5 1 6 4).Text <- "①"
        (footNoteSize body 3 5 1 8 3).Text <- "①"
        (footNoteSize body 3 5 1 10 4).Text <- "①"
        (footNoteSize body 3 5 1 13 2).Text <- "①"


let oligoStampDateFinder (item : string) (sheetName : ExcelWorksheet) columnIndex =
    try 
        let list = List.init sheetName.Dimension.End.Row (fun i -> (i+1,2)) //initializes list of lenth of column 2 rows
        let coordinates = List.find (fun (row,col) -> item.Equals ((string sheetName.Cells.[row,col].Value).Trim(), StringComparison.InvariantCultureIgnoreCase)) list
        let row, _colnum = coordinates
        let value = sheetName.Cells.[row,columnIndex].Text |> string
        value
    with 
        |_ -> 
            let trimmedItem = item.[1..item.Length]
            let cReplacement = "C" + trimmedItem
            let list = List.init sheetName.Dimension.End.Row (fun i -> (i+1,2))
            let coordinates = List.find (fun (row,col) -> cReplacement.Equals ((string sheetName.Cells.[row,col].Value).Trim(), StringComparison.InvariantCultureIgnoreCase)) list
            let row, _colnum = coordinates
            let value = sheetName.Cells.[row,columnIndex].Text |> string
            value



let roundupbyfive(i) : float = 
    (System.Math.Ceiling(i / 5.0) * 5.0)
