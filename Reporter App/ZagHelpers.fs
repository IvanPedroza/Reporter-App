module ZagHelpers

open System
open OfficeOpenXml
open DocumentFormat.OpenXml
open DocumentFormat.OpenXml.Packaging
open DocumentFormat.OpenXml.Wordprocessing
open System.IO
open System.Linq


let getLotNumberCellText (body : Body) tableIndex rowIndex cellIndex paragraphIndex =
    let lot = body.Elements<Table>().ElementAt(tableIndex)
    let tableRow = lot.Elements<TableRow>().ElementAt(rowIndex)
    let tableCell = tableRow.Elements<TableCell>().ElementAt(cellIndex)
    let paragraph = tableCell.Elements<Paragraph>().ElementAt(paragraphIndex)
    let run = paragraph.AppendChild(new Run())
    let text = run.AppendChild(new Text())
    text

let getCalculations (body : Body) paragraphIndex runIndex footnoteId=
    let table = body.Elements<Table>().ElementAt(1)
    let tableRow = table.Elements<TableRow>().ElementAt(2)
    let tableCell = tableRow.Elements<TableCell>().ElementAt(1)
    let paragraph = tableCell.Elements<Paragraph>().ElementAt(paragraphIndex)
    let run = paragraph.Elements<Run>().ElementAt(runIndex)
    if footnoteId = 0 then    
        let runProperties = run.AppendChild(new RunProperties())
        let fontSize = runProperties.AppendChild<FontSize>(new FontSize(Val = StringValue("8")))
        run.Elements<RunProperties>().Equals(fontSize) |> ignore
        let text = run.AppendChild(new Text())
        text
    else
        let text = run.AppendChild(new Text())
        text


let getCsInfo (body : Body) paragraphIndex runIndex =
    let paragraph = body.Elements<Paragraph>().ElementAt(paragraphIndex)
    let run = paragraph.Elements<Run>().ElementAt(runIndex)
    let runProperties = run.Elements<RunProperties>().First()
    let underline = runProperties.AppendChild<Underline>(new Underline(Val = EnumValue<UnderlineValues>DocumentFormat.OpenXml.Wordprocessing.UnderlineValues.Single))
    let position = runProperties.AppendChild<Position>(new Position(Val = StringValue("3")))
    run.Elements<RunProperties>().Equals(underline) |>ignore
    run.Elements<RunProperties>().Equals(position) |> ignore
    let text = run.AppendChild(new Text())
    text


let calNote (body : Body) (inputParams : string list) (param : string) : unit =
    let firstLot = inputParams.[0]
    let table = body.Elements<Table>().ElementAt(2)
    let row = table.Elements<TableRow>().ElementAt(14)
    let cell = row.Elements<TableCell>().ElementAt(0)
    let paragraph = cell.Elements<Paragraph>().ElementAt(4)
    let run = paragraph.AppendChild (new Run())
    let runProperties = run.AppendChild(new RunProperties())
    let fontSize = runProperties.AppendChild<FontSize>(new FontSize(Val = StringValue("18")))
    run.Elements<RunProperties>().Equals(fontSize) |> ignore
    let commentText = run.AppendChild (new Text())

    if param = firstLot then
        let restOfList = inputParams.[1..inputParams.Length-1] |> String.concat ", "
        let theFirstNote = "① Calculations include " + restOfList + "."
        commentText.Text <- theFirstNote
    else

        let otherNote = "① Calculations are on " + firstLot + "."
        commentText.Text <- otherNote

//Creates empty list - finds the Excel cell location for input and deconstructs the tuple into row and column numbers
let zagListFunction (item : string) (sheetName : ExcelWorksheet) columnIndex =
    let list = List.init sheetName.Dimension.End.Row (fun i -> (i+1,1))
    let coordinates = List.find (fun (row,col) -> item.Equals ((string sheetName.Cells.[row,col].Value).Trim(), StringComparison.InvariantCultureIgnoreCase)) list
    let row, _colnum = coordinates
    let value = sheetName.Cells.[row,columnIndex].Value |> string
    value

