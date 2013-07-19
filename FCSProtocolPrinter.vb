Imports System.Drawing

' dgp rev 4/12/2011
Public Class FCSProtocolPrinter

    Public Shared mHeaderText As String
    Public Shared mHeaderHeightPercent As Integer
    Public Shared mFooterHeightPercent As Integer
    Public Shared mInterSectionSpacingPercent As Integer
    Public Shared mHeaderPen As Pen
    Public Shared mFooterPen As Pen
    Public Shared mGridPen As Pen
    Public Shared mHeaderBrush As Brush
    Public Shared mEvenRowBrush As Brush
    Public Shared mColumnHeaderBrush As Brush
    Public Shared mOddRowBrush As Brush
    Public Shared mFooterBrush As Brush
    Public Shared mPagesAcross As Integer
    Public Shared mPrinterSettings As System.Drawing.Printing.PrinterSettings
    Public Shared mPrintDocument As System.Drawing.Printing.PrintDocument

    Public Shared Sub PrinterConfig()

        mHeaderText = ""

        mHeaderHeightPercent = New Decimal(New Integer() {5, 0, 0, 0})
        mFooterHeightPercent = CInt(New Decimal(New Integer() {5, 0, 0, 0}))
        mInterSectionSpacingPercent = CInt(New Decimal(New Integer() {1, 0, 0, 0}))
        mHeaderPen = New Pen(CType(Color.AliceBlue, System.Drawing.Color))
        mFooterPen = New Pen(CType(Color.AliceBlue, System.Drawing.Color))
        mGridPen = New Pen(CType(Color.AliceBlue, System.Drawing.Color))
        mHeaderBrush = Brushes.BlanchedAlmond
        mEvenRowBrush = Brushes.White
        mOddRowBrush = Brushes.White
        mFooterBrush = Brushes.White
        mColumnHeaderBrush = Brushes.White
        mPagesAcross = CInt(New Decimal(New Integer() {1, 0, 0, 0}))
        mPrintDocument = New Printing.PrintDocument
        mPrintDocument.PrinterSettings = New System.Drawing.Printing.PrinterSettings()

    End Sub


End Class
