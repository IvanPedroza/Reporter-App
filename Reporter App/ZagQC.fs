module ZagQC

open System
open OfficeOpenXml
open DocumentFormat.OpenXml
open DocumentFormat.OpenXml.Packaging
open DocumentFormat.OpenXml.Wordprocessing
open System.IO
open System.Linq
open ZagHelpers



let zagStart (inputParams : string list) (zagForm : string) (reporter : ExcelWorksheet)(myTools : ExcelWorksheet) =

       Console.WriteLine "How many plates are you running?"
       let plateInput = Console.ReadLine() |> float
   
       //Console.WriteLine "Which reagents are you using?"
       let reagentBox = "rpzag" //Console.ReadLine ()

       Console.WriteLine "Are you running controls?"
       Console.WriteLine ("Yes or No")
       let controlsInput = Console.ReadLine ()
   

       let reagentsList = List.init myTools.Dimension.End.Row (fun i -> (i+1,1)) 
       let coordinates = List.find (fun (row,col) -> reagentBox.Equals ((string myTools.Cells.[row,col].Value).Trim(), StringComparison.InvariantCultureIgnoreCase)) reagentsList
       let reagentsRow, column = coordinates
   
       let gelVolume = ((plateInput - 1.0) * 10.0) + 43.0
   

       //Starting point for manipulating sheet info and storing it in "param"
       for param in inputParams do
       
       
           //Creates empty list and finds the Excel cell location for input and deconstructs the tuple into row and column numbers
           let csInfoList = List.init 100 (fun i -> (i+1,1)) 
           let coordinates = List.find (fun (row,col) -> param.Equals ((string reporter.Cells.[row,col].Value).Trim(), StringComparison.InvariantCultureIgnoreCase)) csInfoList
           let csRow, _colnum = coordinates

           //Takes value of each cell of the row in which the input lies and stores it for use in filling out Word Doc
           let lot = reporter.Cells.[csRow,1].Value |> string
           let csname = reporter.Cells.[csRow,2].Value |> string
           let geneNumber = reporter.Cells.[csRow,5].Value |> string
           let scale = reporter.Cells.[csRow,6].Value |> string

           //Pulls reagent info from the reagents workbook
           let gelLot = zagListFunction "rpzag" myTools 2 
           let idLot = zagListFunction "rpzag" myTools 3
           let ibLot = zagListFunction "rpzag" myTools 4
           let fmLot = zagListFunction "rpzag" myTools 5
           let ccLot = zagListFunction "rpzag" myTools 6
           let b1Lot = zagListFunction "rpzag" myTools 7
           let g1Lot = zagListFunction "rpzag" myTools 8
           let r1Lot = zagListFunction "rpzag" myTools 9
           let y1Lot = zagListFunction "rpzag" myTools 10
           let upperLot = zagListFunction "rpzag" myTools 11
           let g6Lot = zagListFunction "rpzag" myTools 12
           let firstLot = inputParams.[0]
 
       

           //Reads in Word Doc, creates memory stream and starts processing
           let rpZAG = "C:/Users/ipedroza/source/repos/FRM-10498-02_RP Ligation QC Using CE.docx"
           let docArray = File.ReadAllBytes(rpZAG)
           use docCopy = new MemoryStream(docArray)
           use myDocument = WordprocessingDocument.Open (docCopy, true)
           let body = myDocument.MainDocumentPart.Document.Body

           let asInt = geneNumber |> float
           let wholeNumber = asInt/96.0 
           let roundUp = Math.Ceiling(wholeNumber)
           let lotPlates = 
               if roundUp = 1.0 then 
                   1 |> string
               else
                   Console.WriteLine ("For " + param + " are you running the last plate?")
                   let answer = Console.ReadLine()
                   if answer.Equals("yes", StringComparison.InvariantCultureIgnoreCase) then
                       "1 - " + roundUp.ToString()
                   else          
                       "1 - " + (roundUp - 1.0).ToString()
       

           //CS lot identifier info and number of plates being run
           (getCsInfo body 3 3).Text <- lot + " " + csname
           (getCsInfo body 3 9).Text <- lotPlates
           (getCsInfo body 3 14).Text <- roundUp |> string
           (getCsInfo body 5 8).Text <- geneNumber
           (getCsInfo body 5 14).Text <- scale

           //finds cell of each reagent
           (getLotNumberCellText body 0 1 3 0).Text <- gelLot
           (getLotNumberCellText body 0 2 3 0).Text <- idLot 
           (getLotNumberCellText body 0 3 3 0).Text <- ibLot
           (getLotNumberCellText body 0 4 3 0).Text <- fmLot
           (getLotNumberCellText body 0 5 3 0).Text <- ccLot
           (getLotNumberCellText body 0 11 3 0).Text <- g6Lot
           let blue = (getLotNumberCellText body 0 6 3 0)
           let green = (getLotNumberCellText body 0 7 3 0)
           let red =(getLotNumberCellText body 0 8 3 0)
           let yellow = (getLotNumberCellText body 0 9 3 0)
           let upperStandard = (getLotNumberCellText body 0 10 3 0)

           if controlsInput.Equals("yes", StringComparison.InvariantCultureIgnoreCase) then
               if param = firstLot then 
                   blue.Text <- b1Lot
                   green.Text <- g1Lot
                   red.Text <- r1Lot
                   yellow.Text <- y1Lot
                   upperStandard.Text <- upperLot
               else
                   blue.Text <- "N/A"
                   green.Text <- "N/A"
                   red.Text <- "N/A"
                   yellow.Text <- "N/A"
                   upperStandard.Text <- "N/A"
           else
               blue.Text <- "N/A"
               green.Text <- "N/A"
               red.Text <- "N/A"
               yellow.Text <- "N/A"
               upperStandard.Text <- "N/A"

           //Calculations text and footnotes
           let gelCalculations = getCalculations body 2 5 1
           let gelFootNote = getCalculations body 2 7 0
           let dyeCalculations = getCalculations body 4 3 1
           let dyeFootNote = getCalculations body 4 5 0
           let ccCalculations = getCalculations body 6 2 1 
           let ccFootNote = getCalculations body 6 4 0

           //Adds footnote to calculations section and comment section
       
           if inputParams.Length > 1 then       
               if param = firstLot then
                   gelCalculations.Text <- gelVolume.ToString()
                   dyeCalculations.Text <- gelVolume.ToString()
                   ccCalculations.Text <- gelVolume.ToString()
                   gelFootNote.Text <- "①"
                   dyeFootNote.Text <- "①"
                   ccFootNote.Text <- "①"
               
               else
                   gelFootNote.Text <- "①"
                   dyeFootNote.Text <- "①"
                   ccFootNote.Text <- "①"
               
               calNote body inputParams param
           else
               gelCalculations.Text <- gelVolume.ToString()
               dyeCalculations.Text <- gelVolume.ToString()
               ccCalculations.Text <- gelVolume.ToString()
               gelFootNote.Text <- ""
               dyeFootNote.Text <- ""
               ccFootNote.Text <- ""


           //Saves filled out Doc
           myDocument.SaveAs("C:/Users/ipedroza/source/repos/"+param + " RP ZAG Form" + ".docx") |> ignore



