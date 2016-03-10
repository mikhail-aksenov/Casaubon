namespace Casaubon.Tests

open NUnit.Framework
open System.IO
open Casaubon.Core

[<TestFixture>]
type ReaderTests() =
  let filePath = Path.Combine(__SOURCE_DIRECTORY__, "TestWB.xlsx")

  [<Test>]
  member this.TestMakeFromFilePath() =
    use package = ExcelFactory.make(filePath)
    Assert.AreEqual(2, package.Workbook.Worksheets.Count)
    

  [<Test>]
  member this.TestMakeFromStream() =
    use stream = new FileStream(filePath, FileMode.Open)
    use package = ExcelFactory.make(stream)
    Assert.AreEqual(2, package.Workbook.Worksheets.Count)

  [<Test>]
  member this.TestRowTraverse() =
    use stream = new FileStream(filePath, FileMode.Open)
    use package = ExcelFactory.make(stream)
    let test = worksheetByName "TestWS1" package
               |> rows
               |> Seq.map Seq.length
               |> Seq.forall (fun l -> l = 3)
    Assert.IsTrue(test)

  [<Test>]
  member this.TestColTraverse() = 
    use package = ExcelFactory.make(filePath)
    let test = worksheetByNumber 2 package
               |> cols
               |> Seq.map Seq.length
               |> Seq.forall (fun l -> l = 4)
    Assert.IsTrue(test)