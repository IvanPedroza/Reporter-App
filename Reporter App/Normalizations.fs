module Normalizations


open System
open OfficeOpenXml
open DocumentFormat.OpenXml
open DocumentFormat.OpenXml.Packaging
open DocumentFormat.OpenXml.Wordprocessing
open System.Linq
open System.IO
open NormalizationHelpers


let normalizationStart (inputParams : string list) (normalizationForm : string) (normPrecipitationForm : string) (reporter : ExcelWorksheet) (myTools : ExcelWorksheet) =
    
    let user = Environment.UserName

    Console.WriteLine "Which bench space are you using?"
    let benchInput = Console.ReadLine ()


    //Reading in reagents excel book
    //Console.WriteLine "Which reagents are you using?"
    let reagentsInput = "normalizations" //Console.ReadLine ()


    for param in inputParams do 
        let pipetteCalibration = listFunction benchInput myTools 2
        let p1000 = listFunction benchInput myTools 3
        let p1000Id = listFunction benchInput myTools 4
        let p200 = listFunction benchInput myTools 5
        let p200Id = listFunction benchInput myTools 6
        let p100 = listFunction benchInput myTools 7
        let p100Id = listFunction benchInput myTools 8
        let p20 = listFunction benchInput myTools 9
        let p20Id = listFunction benchInput myTools 10
        let p10 = listFunction benchInput myTools 11
        let p10Id = listFunction benchInput myTools 12
        let p2 = listFunction benchInput myTools 13
        let p2Id = listFunction benchInput myTools 14
        let p2000 = listFunction benchInput myTools 15
        let p2000Id = listFunction benchInput myTools 16
        let mc8P20Id = listFunction benchInput myTools 18
        let mc12P20Id = listFunction benchInput myTools 22
        let mc8P200Id = listFunction benchInput myTools 20
        let mc12P200Id = listFunction benchInput myTools 24


        let waterLot = listFunction reagentsInput myTools 2
        let waterExp = listFunction reagentsInput myTools 3
        let waterUbd = listFunction reagentsInput myTools 4
        let sspeLot = listFunction reagentsInput myTools 5
        let sspeExp = listFunction reagentsInput myTools 6
        let sspeUbd = listFunction reagentsInput myTools 7
        let tweenLot = listFunction reagentsInput myTools 8
        let tweenExp = listFunction reagentsInput myTools 9
        let tweenUbd = listFunction reagentsInput myTools 10
        let nanodrop = listFunction reagentsInput myTools 11
        let nanodropCal = listFunction reagentsInput myTools 12
        let acetateLot = listFunction reagentsInput myTools 13
        let acetateExpiration = listFunction reagentsInput myTools 14
        let ethanolLot = listFunction reagentsInput myTools 15
        let ethanolExpiration = listFunction reagentsInput myTools 16
        let ethanolUBD = listFunction reagentsInput myTools 17


      
    
        let readExcelBytes = File.ReadAllBytes(normalizationForm)
        let excelStream = new MemoryStream(readExcelBytes)
        let formPackage = new ExcelPackage(excelStream)
        let normSheet = formPackage.Workbook.Worksheets.["Normalization"]
        let preIncSheet = formPackage.Workbook.Worksheets.["Pre-incubation"]
    

        //Takes value of each cell of the row in which the input lies and stores it for use in filling out Word Doc
        let lot = listFunction param reporter 1 
        let csName = listFunction param reporter 2 
        let geneNumber = listFunction param reporter 5  |> float
        let scale = listFunction param reporter 6  |> float
       


        normSheet.Cells.[7, 3].Value <- lot + " " + csName
        normSheet.Cells.[7, 6].Value <- geneNumber
        normSheet.Cells.[7, 8].Value <- scale

        normSheet.Cells.[10, 4].Value <- waterLot
        normSheet.Cells.[10, 7].Value <- waterExp
        normSheet.Cells.[11, 7].Value <- "Use by date: " + waterUbd
        normSheet.Cells.[12, 4].Value <- sspeLot
        normSheet.Cells.[12, 7].Value <- sspeExp
        normSheet.Cells.[13, 7].Value <- "Use by date: " + sspeUbd
        normSheet.Cells.[14, 4].Value <- tweenLot
        normSheet.Cells.[14, 7].Value <- tweenExp
        normSheet.Cells.[15, 7].Value <- "Use by date: " + tweenUbd
        normSheet.Cells.[17, 4].Value <- nanodrop
        normSheet.Cells.[17, 7].Value <- nanodropCal
        normSheet.Cells.[19, 4].Value <- p2Id
        normSheet.Cells.[19, 7].Value <- pipetteCalibration
        normSheet.Cells.[20, 4].Value <- p10Id
        normSheet.Cells.[20, 7].Value <- pipetteCalibration
        normSheet.Cells.[21, 4].Value <- p20Id
        normSheet.Cells.[21, 7].Value <- pipetteCalibration
        normSheet.Cells.[22, 4].Value <- p100Id
        normSheet.Cells.[22, 7].Value <- pipetteCalibration
        normSheet.Cells.[23, 4].Value <- p200Id
        normSheet.Cells.[23, 7].Value <- pipetteCalibration
        normSheet.Cells.[24, 4].Value <- p1000Id
        normSheet.Cells.[24, 7].Value <- pipetteCalibration


        preIncSheet.Cells.[6, 3].Value <- lot + " " + csName
        preIncSheet.Cells.[6, 6].Value <- geneNumber
        preIncSheet.Cells.[6, 8].Value <- scale

        preIncSheet.Cells.[9, 4].Value <- nanodrop
        preIncSheet.Cells.[9, 7].Value <- nanodropCal

        preIncSheet.Cells.[12, 4].Value <- p2Id
        preIncSheet.Cells.[12, 7].Value <- pipetteCalibration
        preIncSheet.Cells.[13, 4].Value <- p10Id
        preIncSheet.Cells.[13, 7].Value <- pipetteCalibration
        preIncSheet.Cells.[14, 4].Value <- p20Id
        preIncSheet.Cells.[14, 7].Value <- pipetteCalibration
        preIncSheet.Cells.[15, 4].Value <- p100Id
        preIncSheet.Cells.[15, 7].Value <- pipetteCalibration
        preIncSheet.Cells.[16, 4].Value <- p200Id
        preIncSheet.Cells.[16, 7].Value <- pipetteCalibration
        preIncSheet.Cells.[17, 4].Value <- p1000Id
        preIncSheet.Cells.[17, 7].Value <- pipetteCalibration

    
        Directory.CreateDirectory("W:/Production/Reporter requests/In Progress/" + param + " " + csName)

        let path = "W:/Production/Reporter requests/In Progress/" + param + " " + csName + "/" + param + " " + csName + " Normalization" + ".xlsx"

        let fileBytes = formPackage.GetAsByteArray()

        File.WriteAllBytes(path, fileBytes)

    
        //Reads in Word Doc and starts processing

        let memoryStream = new MemoryStream()
        use fileStream = new FileStream(normPrecipitationForm, FileMode.Open, FileAccess.Read)
        fileStream.CopyTo(memoryStream)
        let myDocument = WordprocessingDocument.Open(memoryStream, true)
        let body = myDocument.MainDocumentPart.Document.Body

        //Fills CS indentifying info for pages one and two
        (csId body 1 7).Text <- lot + " " + csName
        (csId body 1 14).Text <- geneNumber.ToString()
        (csId body 1 20).Text <- scale.ToString()
        (csId body 6 3).Text <- lot + " " + csName
        (csId body 6 7).Text <- geneNumber.ToString()
        (csId body 6 12).Text <- scale.ToString()  
        
        (csId body 10 3).Text <- lot + " " + csName
        (csId body 10 8).Text <- geneNumber.ToString()
        (csId body 10 13).Text <- scale.ToString()
   

        //filling out lot numbers
        (fillCells body 0 1 2 0 0).Text <- lot
        (fillCells body 0 1 3 0 0).Text <- "N/A"
        (fillCells body 0 2 2 0 0).Text <- acetateLot
        (fillCells body 0 2 3 0 0).Text <- acetateExpiration
        (fillCells body 0 3 2 0 0).Text <- ethanolLot
        (fillCells body 0 3 3 0 0).Text <- ethanolExpiration
        (fillCells body 0 3 3 1 2).Text <- ethanolUBD
        (fillCells body 0 5 0 0 0).Text <- p1000
        (fillCells body 0 5 1 0 0).Text <- p1000Id
        (fillCells body 0 5 2 0 0).Text <- pipetteCalibration
        (fillCells body 0 6 2 0 0).Text <- pipetteCalibration
        (fillCells body 0 7 2 0 0).Text <- pipetteCalibration

        //gets text in cells which value depends on the scale of the build and the bench where the process will be done
        let acetatePipetteSize = (fillCells body 0 6 0 0 0)
        let acetatePipetteId =  (fillCells body 0 6 1 0 0)
        let poolingPipetteSize = (fillCells body 0 7 0 0 0)
        let poolingPipetteId = (fillCells body 0 7 1 0 0)
        let precipitationVolume = (fillCells body 1 3 1 0 5)
        let acetateVolume = (fillCells body 1 4 1 0 5)
        let alcoholVolume = (fillCells body 1 6 1 0 9)

        //Fills out theoretical volume of pooled reporters
        let theoreticalVolume = thereticalVolumes (scale |> float) (geneNumber |> float)
        //(fillCells body 1 4 1 0 2).Text <- theoreticalVolume.ToString()

        //Calculates precipitation and reagent volumes and then assigns eqipment needed for working with those volumes
        if scale = 0.15 then 
            let volume, acetate, alcohol = calculations 5.1 (geneNumber |> float)

            precipitationVolume.Text <- volume.ToString()
            acetateVolume.Text <- acetate.ToString()
            alcoholVolume.Text <- alcohol.ToString()
            poolingPipetteSize.Text <- mc12P20Id
            poolingPipetteId.Text <- mc12P20Id

            match acetate with 
            | _ when acetate < 20.0 ->
                acetatePipetteSize.Text <- p20
                acetatePipetteId.Text <- p20Id
            | _ when acetate >= 20.0 ->
                acetatePipetteSize.Text <- p200
                acetatePipetteId.Text <- p200Id

        elif scale = 0.25 then 
            let volume, acetate, alcohol = calculations 8.5 (geneNumber |> float)

            precipitationVolume.Text <- volume.ToString()
            acetateVolume.Text <- acetate.ToString()
            alcoholVolume.Text <- alcohol.ToString()
            poolingPipetteSize.Text <- mc8P20Id
            poolingPipetteId.Text <- mc12P20Id

            match acetate with 
            | _ when acetate < 20.0 ->
                acetatePipetteSize.Text <- p20
                acetatePipetteId.Text <- p200Id
            | _ when acetate >= 20.0 ->
                acetatePipetteSize.Text <- p200
                acetatePipetteId.Text <- p200Id

        else
            let volume = System.Math.Ceiling(theoreticalVolume)
            let acetate = System.Math.Ceiling(volume / 9.0)
            let alcohol = System.Math.Ceiling((volume + acetate) * 2.5)

            //(fillCells body 2 2 1 0 2).Text <- volume.ToString()
            acetateVolume.Text <- acetate.ToString()
            alcoholVolume.Text <- alcohol.ToString()
            acetatePipetteSize.Text <- p200
            acetatePipetteId.Text <- p200Id
            poolingPipetteSize.Text <- mc12P200Id
            poolingPipetteId.Text <- mc12P200Id    

        let user = Environment.UserName
        let path = "C:\\users\\" + user + "\\AppData\\Local\\Temp\\ "+param + " Normalization Precipitation Form" + ".docx"
        myDocument.SaveAs(path).Close() |> ignore



