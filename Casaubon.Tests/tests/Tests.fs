namespace Casaubon.Tests

open FsUnit
open NUnit.Framework
open System.IO
open Casaubon.Core

[<TestFixture>]
type ReaderTests() =
  let filePath = Path.Combine(__SOURCE_DIRECTORY__, "..", "TestWB.xlsx")

  [<Test>]
  member this.TestMakeFromFilePath() =
    use package = ExcelFactory.make(filePath)
    package.Workbook.Worksheets.Count |> should equal 2

  [<Test>]
  member this.TestMakeFromStream() =
    use stream = new FileStream(filePath, FileMode.Open)
    use package = ExcelFactory.make(stream)
    package.Workbook.Worksheets.Count |> should equal 2

  [<Test>]
  member this.TestRowTraverse() =
    use stream = new FileStream(filePath, FileMode.Open)
    use package = ExcelFactory.make(stream)
    worksheetByName "TestWS1" package
    |> rows
    |> Seq.map Seq.length
    |> Seq.forall (fun l -> l = 3)
    |> should be True

  [<Test>]
  member this.TestColTraverse() = 
    use package = ExcelFactory.make(filePath)
    worksheetByNumber 2 package
    |> cols
    |> Seq.map Seq.length
    |> Seq.forall (fun l -> l = 4)
    |> should be True      