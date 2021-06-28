module NormalizationHelpers

open System
open OfficeOpenXml
open DocumentFormat.OpenXml
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

    

 //Creates empty list - finds the Excel cell location for input and deconstructs the tuple into row and column numbers
let listFunction (item : string) (sheetName : ExcelWorksheet) columnIndex =
    let list = List.init sheetName.Dimension.End.Row (fun i -> (i+1,1)) 
    let coordinates = List.find (fun (row,col) -> item.Equals((string sheetName.Cells.[row,col].Value).Trim(), StringComparison.InvariantCultureIgnoreCase)) list
    let row, _colnum = coordinates
    let value = sheetName.Cells.[row,columnIndex].Value |> string
    value


//Fills text in table cells
let fillCells (body : Body) tableIndex tableRowIndex tableCellIndex paragraphIndex runIndex = 
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
        let underline = runProperties.AppendChild<Underline>(new Underline(Val = EnumValue<UnderlineValues>DocumentFormat.OpenXml.Wordprocessing.UnderlineValues.Single))
        run.Elements<RunProperties>().Equals(underline) |>ignore 
        let fontSize = runProperties.AppendChild<FontSize>(new FontSize(Val = StringValue("18")))
        run.Elements<RunProperties>().Equals(fontSize) |> ignore
        let text = run.AppendChild(new Text())
        text

//Calculates theoretical volume of pooled probes
let thereticalVolumes (scale : float) (geneNumber : float) : float = 
    match scale with 
        | scale when scale <=  0.5 -> geneNumber * 14.0
        | 1.0 -> geneNumber * 31.0
        | 1.25 -> geneNumber * 39.5
        | 1.5 -> geneNumber * 48.0
        | 2.0 -> geneNumber * 65.0
        | 2.5 -> geneNumber * 82.0
        | 3.0 -> geneNumber * 99.0
        | 3.5 -> geneNumber * 116.0
        | 4.0 -> geneNumber * 133.0
        | 4.5 -> geneNumber * 150.0
        | 10.0 -> geneNumber * 337.0   
        | _ -> failwith "Error..."

//Calculates reagent qantities needed for pre-determined precipitation volumes 
let calculations (scaleMultiplier : float) (geneNumber : float) : (float * float * float) =
    let volume = System.Math.Ceiling(scaleMultiplier * geneNumber)
    let acetate = System.Math.Round((volume / 9.0),1)
    let alcohol = System.Math.Ceiling((volume + acetate) * 2.5)
    volume, acetate, alcohol