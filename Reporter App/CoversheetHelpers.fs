module CoversheetHelpers

open System
open OfficeOpenXml
open DocumentFormat.OpenXml
open DocumentFormat.OpenXml.Packaging
open DocumentFormat.OpenXml.Wordprocessing
open System.IO
open System.Linq
    
// assings form parameter to read cs formulation and populate cs type
let formulationToCodeSetType (formulation : string) : string =
    match formulation with
    | "XT" -> "RNA"
    | "TBD" -> "TBD"
    | "STD" | "miRNA" -> "Panel/CodeSet Plus (RNA)"
    | _ -> failwith "Error ..."

//defining a function with two parameters "csname" and "form"
//function looks for starting characters of CS name to determine formulation type
let determine (csname : string) (form : string) : string =
    match csname with 
        | csname when csname.StartsWith("CNV", StringComparison.CurrentCultureIgnoreCase)  -> "CNV (DNA)"
        | csname when csname.StartsWith("PLS", StringComparison.CurrentCultureIgnoreCase)  -> "Panel/CodeSet Plus (RNA)"
        | csname when csname.StartsWith("PLS_CNV", StringComparison.CurrentCultureIgnoreCase)  -> "Panel/CodeSet Plus (DNA)"
        | csname when csname.StartsWith("miR", StringComparison.CurrentCultureIgnoreCase)  -> "miRNA"
        | csname when csname.StartsWith("DNA", StringComparison.CurrentCultureIgnoreCase)  -> "DNA"
        | csname when csname.StartsWith("miX", StringComparison.CurrentCultureIgnoreCase)  -> "miRGE/miXED"
        | csname when csname.StartsWith("CHIP", StringComparison.CurrentCultureIgnoreCase)  -> "CHIP"
        | _ -> formulationToCodeSetType form


let fillTopCells (body : Body) paragraphIndex = 
    let paragraph = body.Elements<Paragraph>().ElementAt(paragraphIndex)
    let SdtRun = paragraph.Elements<SdtRun>().ElementAt(0)
    let contentRun = SdtRun.Elements<SdtContentRun>().ElementAt(0)
    let run = contentRun.Elements<Run>().ElementAt(0)
    let runproperties = run.AppendChild(new RunProperties())
    let fontSize = runproperties.AppendChild<FontSize>(FontSize(Val = StringValue("28")))
    run.Elements<RunProperties>().Equals(fontSize) |> ignore
    let text = run.Elements<Text>().ElementAt(0)
    text

let staticCells (body : Body) tableIndex tableRowIndex tableCellIndex paragraphIndex = 
    let table = body.Elements<Table>().ElementAt(tableIndex)
    let row = table.Elements<TableRow>().ElementAt(tableRowIndex)
    let cell = row.Elements<TableCell>().ElementAt(tableCellIndex)
    let paragraph = cell.Elements<Paragraph>().ElementAt(paragraphIndex)
    if (paragraph.Elements<Run>().Count() = 0) then
        let run = paragraph.AppendChild(new Run())
        let runproperties = run.AppendChild(new RunProperties())
        let fontSize = runproperties.AppendChild<FontSize>(FontSize(Val = StringValue("20")))
        run.Elements<RunProperties>().Equals(fontSize) |> ignore
        let text = run.AppendChild(new Text())
        text
    else
        let run = paragraph.Elements<Run>().ElementAt(0)
        let runproperties = run.AppendChild(new RunProperties())
        let fontSize = runproperties.AppendChild<FontSize>(FontSize(Val = StringValue("20")))
        run.Elements<RunProperties>().Equals(fontSize) |> ignore
        let text = run.Elements<Text>().First()
        text
        

let determiningConcentration (body : Body) tableIndex rowIndex = 
    let table = body.Elements<Table>().ElementAt(tableIndex)
    let tableRow = table.Elements<TableRow>().ElementAt(rowIndex)
    let tableCell = tableRow.Elements<TableCell>().ElementAt(0)
    let paragraph = tableCell.Elements<Paragraph>().ElementAt(0)
    let sdtRun = paragraph.Elements<SdtRun>().ElementAt(0)
    let sdtContentRun = sdtRun.Elements<SdtContentRun>().ElementAt(0)
    let run = sdtContentRun.Elements<Run>().ElementAt(0)
    let runproperties = run.AppendChild(new RunProperties())
    let fontSize = runproperties.AppendChild<FontSize>(FontSize(Val = StringValue("20")))
    run.Elements<RunProperties>().Equals(fontSize) |> ignore
    let text = run.Elements<Text>().ElementAt(0)
    text

let dropdownCells (body : Body) cellIndex = 
    let table = body.Elements<Table>().ElementAt(0)
    let tableRow = table.Elements<TableRow>().ElementAt(1)
    let sdtCell = tableRow.Elements<SdtCell>().ElementAt(cellIndex)
    let contentCell = sdtCell.Elements<SdtContentCell>().ElementAt(0)
    let tableCell = contentCell.Elements<TableCell>().ElementAt(0)
    let paragraph = tableCell.Elements<Paragraph>().ElementAt(0)
    let run = paragraph.Elements<Run>().ElementAt(0)
    let runproperties = run.AppendChild(new RunProperties())
    let fontSize = runproperties.AppendChild<FontSize>(FontSize(Val = StringValue("20")))
    run.Elements<RunProperties>().Equals(fontSize) |> ignore
    let text = run.Elements<Text>().ElementAt(0)
    text


//Pulling current year for date formating purposes    
let year = (DateTime.Now.Year.ToString())



