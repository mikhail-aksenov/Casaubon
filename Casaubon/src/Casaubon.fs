namespace Casaubon

module Core =
  open OfficeOpenXml
  open System
  open System.IO

  /// Factory wrapper for OfficeOpenXml.ExcelPackage class.
  /// Calling side must deal with error handling and resource disposal.
  type ExcelFactory = 
    static member make () = 
      new ExcelPackage()

    static member make (path: string) =
      let fileInfo = new FileInfo(path)
      new ExcelPackage(fileInfo)

    static member make (stream: Stream) = 
      new ExcelPackage(stream)
  
  /// Get worksheet by 1-based number.
  let worksheetByNumber (number: int) (package: ExcelPackage) =
    package.Workbook.Worksheets.[number]

  /// Get worksheet by name
  let worksheetByName (name: string) (package: ExcelPackage) =
    package.Workbook.Worksheets.[name]
  
  /// Retreive worksheet cells using row-first traverse.
  let cells (worksheet: ExcelWorksheet) = 
    let rowCount = worksheet.Dimension.End.Row
    let colCount = worksheet.Dimension.End.Column
    seq { for i in 1 .. rowCount do
            for j in 1 .. colCount do
              yield worksheet.Cells.[i, j] }

  /// Retreive sequence of rows.
  let rows (worksheet: ExcelWorksheet) =
    let rowCount = worksheet.Dimension.End.Row
    let colCount = worksheet.Dimension.End.Column
    seq { for i in 1 .. rowCount do 
            yield seq { for j in 1 .. colCount do 
                          yield worksheet.Cells.[i, j] }}
  
  /// Retreive sequence of columns
  let cols (worksheet: ExcelWorksheet) =
    let rowCount = worksheet.Dimension.End.Row
    let colCount = worksheet.Dimension.End.Column
    seq { for i in 1 .. colCount do 
            yield seq { for j in 1 .. rowCount do 
                          yield worksheet.Cells.[i, j] }}