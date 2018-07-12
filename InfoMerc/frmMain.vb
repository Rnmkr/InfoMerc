Imports Excel = Microsoft.Office.Interop.Excel
Imports System.IO
Public Class frmMain
    Private DisplayedScreen As Screen
    Private CurrentDisplay As Integer

    Private WithEvents mnuTaskbarMenu As ContextMenuStrip
    Private WithEvents mnuRefreshDisplay As ToolStripMenuItem
    Private WithEvents mnuSelectDisplay As ToolStripMenuItem
    Private WithEvents mnuSeparator As ToolStripSeparator
    Private WithEvents mnuExit As ToolStripMenuItem
    Private lastExcelRow As Integer = 0
    Private iCountOfParaArmar As Integer = 0
    Private iCountOfAControlar As Integer = 0
    Private iCountOfReingresado As Integer = 0
    Private iCountOfPendiente As Integer = 0
    Private iCountOfOK As Integer = 0
    Private icountTotalGridRows As Integer = 0
    Private ScrollRows As Integer = 0
    Private ScrollMod As Integer = 0
    Private ScrollRoundSteps As Integer = 0
    Private fontzise As Integer = 72 'rows font size
    Private OnScreenRows As Integer = 17 'number of rows that fit on screen
    Private xlWorkBook As Excel.Workbook
    Private xlWorkSheet As Excel.Worksheet
    Private xlApp = New Excel.Application
    Private appPath As String = Application.StartupPath()
    Private XLSMFileName As String = "PLANILLA DE INGRESOS PRODUCCION.xlsm"
    Private Position = 0

    Private Sub frmMain_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        mnuRefreshDisplay = New ToolStripMenuItem("Actualizar Display")
        mnuSelectDisplay = New ToolStripMenuItem("Cambiar Display")
        mnuSeparator = New ToolStripSeparator()
        mnuExit = New ToolStripMenuItem("Cerrar")
        mnuTaskbarMenu = New ContextMenuStrip
        mnuTaskbarMenu.Items.AddRange(New ToolStripItem() {mnuRefreshDisplay, mnuSelectDisplay, mnuSeparator, mnuExit})
        Try
            DisplayedScreen = Screen.AllScreens(1)
            CurrentDisplay = 1
        Catch ex As System.IndexOutOfRangeException
            DisplayedScreen = Screen.AllScreens(0)
            CurrentDisplay = 0
        End Try

        ' FORM PROPERTIES
        With Me
            .ShowInTaskbar = False
            .ControlBox = False
            .FormBorderStyle = 0
            .Icon = My.Resources.InfoMerc_Taskbar
            .StartPosition = FormStartPosition.Manual
            .Location = DisplayedScreen.Bounds.Location + New Point(100, 100)
            .WindowState = FormWindowState.Maximized
        End With

        'NOTIFY ICON PROPERTIES
        With NotifyIcon
            .ContextMenuStrip = mnuTaskbarMenu
            .Icon = Me.Icon
            .Text = "InfoMerc"
            .Visible = True
        End With

        'TIMER PROPERTIES
        With TimerScrollGrid
            .Interval = 15000
            .Enabled = False
        End With

        'FILESYSTEMWATCHER PROPERTIES
        With FileSystemWatcher
            .EnableRaisingEvents = True
            .Filter = (XLSMFileName)
            .IncludeSubdirectories = False
            .Path = (appPath)
            .NotifyFilter = NotifyFilters.Attributes
        End With

        'LABEL STATUS PROPERTIES
        With LabelStatus
            .Text = ""
            .BorderStyle = BorderStyle.FixedSingle
            .TextAlign = ContentAlignment.MiddleCenter
            .Font = New Font("Consolas", 48, FontStyle.Bold)
            .AutoSize = False
            .Width = Me.Width
            .Height = LabelStatus.Font.Size * 2
            .Location = New Point(Me.Width / 2 - LabelStatus.Width / 2, Me.Height - LabelStatus.Font.Size * 2)
        End With

        ' DataGridView PROPERTIES
        With DataGridView
            'data grid
            .EnableHeadersVisualStyles = False 'disables windows visual styles
            .Location = New Point(0, 0) 'grid location
            .Height = Me.Height - LabelStatus.Height 'grid height
            .Width = Me.Width 'grid width
            .ScrollBars = False 'removes scrollbars
            .AllowUserToAddRows = False 'removes ugly "blank" row at the end
            'columns and rows format
            .Columns.Add("colPedidos", "PEDIDOS") 'add new column "PEDIDOS"
            .Columns.Add("colCantidad", "CANTIDAD") 'add new column "CANTIDAD"
            .Columns.Add("colProducto", "PRODUCTO") 'add new column "PRODUCTO"
            .Columns.Add("colFecha", "FECHA") 'add new column "FECHA"
            .Columns(3).DefaultCellStyle.Format = "d" 'date format
            .AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill 'sets autosize mode for all columns
            .Columns(0).DefaultCellStyle.Font = New Font("Consolas", fontzise, FontStyle.Bold) 'cell font
            .Columns(1).DefaultCellStyle.Font = New Font("Consolas", fontzise, FontStyle.Bold) 'cell font
            .Columns(2).DefaultCellStyle.Font = New Font("Consolas", fontzise, FontStyle.Bold) 'cell font
            .Columns(3).DefaultCellStyle.Font = New Font("Consolas", fontzise, FontStyle.Bold) 'cell font
            .ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.Single 'headers border style
            .ColumnHeadersDefaultCellStyle.Font = New Font("Consolas", 48, FontStyle.Bold) 'headers font
            .ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter 'headers text alignment
            .ColumnHeadersHeight = 72 'headers height (font size * 2)
            .RowHeadersVisible = False 'removes left rows "header"
            .RowTemplate.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter 'cell text alignment
            .RowTemplate.Height = 117 'cell height
        End With
        ColectData()
    End Sub
    Private Sub mnuRefreshDisplay_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuRefreshDisplay.Click
        ColectData()
    End Sub
    Private Sub mnuSelectDisplay_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuSelectDisplay.Click
        If CurrentDisplay = 1 Then
            DisplayedScreen = Screen.AllScreens(0)
            CurrentDisplay = 0
        Else
            DisplayedScreen = Screen.AllScreens(1)
            CurrentDisplay = 1
        End If


        With Me
            .WindowState = FormWindowState.Normal
            .Location = DisplayedScreen.Bounds.Location
            .WindowState = FormWindowState.Maximized
        End With
    End Sub
    Private Sub mnuExit_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles mnuExit.Click
        Dim result As Integer = MessageBox.Show("Desea cerrar la aplicación?", "InfoMerc", MessageBoxButtons.YesNo, MessageBoxIcon.Information)
        If result = DialogResult.Yes Then
            releaseObject(xlApp)
            releaseObject(xlWorkBook)
            releaseObject(xlWorkSheet)
            Application.Exit()
        End If
    End Sub
    Public Sub FileSystemWatcher_Changed(ByVal sender As System.Object, ByVal e As System.IO.FileSystemEventArgs) Handles FileSystemWatcher.Changed
        ColectData()
    End Sub
    Public Sub FileSystemWatcher_Created(ByVal sender As Object, ByVal e As System.IO.FileSystemEventArgs) Handles FileSystemWatcher.Created
        ColectData()
    End Sub
    Public Sub FileSystemWatcher_Deleted(ByVal sender As Object, ByVal e As System.IO.FileSystemEventArgs) Handles FileSystemWatcher.Deleted
        ColectData()
    End Sub
    Public Sub FileSystemWatcher_Error(ByVal sender As Object, ByVal e As System.IO.ErrorEventArgs) Handles FileSystemWatcher.Error
        MsgBox(e.GetException.ToString)
    End Sub

    Private Sub releaseObject(ByVal obj As Object)
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
            obj = Nothing
        Catch ex As Exception
            obj = Nothing
        Finally
            GC.Collect()
        End Try
    End Sub

    Function isFileOpen(ByRef sName As String) As Boolean
        Dim blnRetVal As Boolean = False
        Dim fs As FileStream

        Try
            fs = File.Open(sName, FileMode.Open, FileAccess.Read, FileShare.None)
        Catch ex As Exception
            blnRetVal = True
        Finally
            If Not IsNothing(fs) Then : fs.Close() : End If
        End Try

        Return blnRetVal

    End Function

    Private Sub ColectData()
        TimerScrollGrid.Enabled = False
        DataGridView.Rows.Clear()
        Dim IsFileOpn = isFileOpen(appPath + "\" + XLSMFileName)

        'IF .XML IS ALREADY OPEN JUST READ IT, ELSE OPEN FILE (TO AVOID OPEN MULTIPLE INSTANCES OF THE XML)
        If IsFileOpn = False Then
            xlWorkBook = xlApp.Workbooks.Open(appPath + "\" + XLSMFileName)
            xlWorkSheet = xlWorkBook.Worksheets("PEDIDOS")
        Else
            xlApp = CType(System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application"), Excel.Application)
            xlWorkBook = GetObject(appPath + "\" + XLSMFileName)
            xlWorkSheet = xlWorkBook.Worksheets("PEDIDOS")
            xlWorkSheet.Activate()
        End If


        'GET LAST USED CELL
        lastExcelRow = xlWorkSheet.UsedRange.Rows.Count

        'On Error Resume Next
        'Dim pf_range As Range
        'pf_range = xlWorkSheet.Range(xlWorkSheet.Cells(2, 2), xlWorkSheet.Cells(lastExcelRow, 1))

        'SET COUNTERS System.Runtime.InteropServices.COMException
        Try
            iCountOfParaArmar = xlApp.WorksheetFunction.CountIf(xlWorkSheet.UsedRange, "PARA ARMAR")
            iCountOfAControlar = xlApp.WorksheetFunction.CountIf(xlWorkSheet.UsedRange, "A CONTROLAR")
            iCountOfOK = xlApp.WorksheetFunction.CountIf(xlWorkSheet.UsedRange, "OK")
            iCountOfReingresado = xlApp.WorksheetFunction.CountIf(xlWorkSheet.UsedRange, "REINGRESADO")
            iCountOfPendiente = xlApp.WorksheetFunction.CountIf(xlWorkSheet.UsedRange, "PENDIENTE")
            icountTotalGridRows = (iCountOfParaArmar + iCountOfAControlar)
            ScrollMod = (icountTotalGridRows Mod OnScreenRows) 'mod of the total rows to scroll
            ScrollRows = (icountTotalGridRows - ScrollMod) 'total rows to scroll w/o mod (multiple of OnScreenRows)
            ScrollRoundSteps = (ScrollRows / OnScreenRows)
        Catch ex As System.Runtime.InteropServices.COMException
            MsgBox("No se pudo refrescar la lista...")
        End Try

        'PRINT RESULTS IN "LABELSTATUS"
        LabelStatus.Text = ("RECIBIDOS:" & iCountOfAControlar & " | CONTROLADOS:" & iCountOfParaArmar & " | REINGRESADOS:" & iCountOfReingresado & " | PENDIENTES:" & iCountOfPendiente)


        For i = 8 To lastExcelRow
            DataGridView.ClearSelection()
            If xlWorkSheet.Cells(i, 13).value = "PARA ARMAR" Then
                DataGridView.RowTemplate.DefaultCellStyle.BackColor = Color.DarkGreen 'cell background color
                DataGridView.RowTemplate.DefaultCellStyle.ForeColor = Color.White 'cell font color
                DataGridView.Rows.Add({xlWorkSheet.Cells(i, 1).value, xlWorkSheet.Cells(i, 3).value, xlWorkSheet.Cells(i, 4).value, xlWorkSheet.Cells(i, 8).value})
            End If
        Next

        For i = 8 To lastExcelRow
            DataGridView.ClearSelection()
            If xlWorkSheet.Cells(i, 13).value = "A CONTROLAR" Then
                DataGridView.RowTemplate.DefaultCellStyle.BackColor = Color.Black 'cell background color
                DataGridView.RowTemplate.DefaultCellStyle.ForeColor = Color.White 'cell font color
                DataGridView.Rows.Add({xlWorkSheet.Cells(i, 1).value, xlWorkSheet.Cells(i, 3).value, xlWorkSheet.Cells(i, 4).value, xlWorkSheet.Cells(i, 8).value})
            End If
        Next

        For i = 8 To lastExcelRow
            DataGridView.ClearSelection()
            If xlWorkSheet.Cells(i, 13).value = "PENDIENTE" Then
                DataGridView.RowTemplate.DefaultCellStyle.BackColor = Color.LightGoldenrodYellow 'cell background color
                DataGridView.RowTemplate.DefaultCellStyle.ForeColor = Color.Black 'cell foreground color
                'DataGridView.DefaultCellStyle.ForeColor = Color.Black 'cell font color
                DataGridView.Rows.Add({xlWorkSheet.Cells(i, 1).value, xlWorkSheet.Cells(i, 3).value, xlWorkSheet.Cells(i, 4).value, xlWorkSheet.Cells(i, 8).value})
            End If
        Next

        'RELEASE XLAPP
        releaseObject(xlApp)
        releaseObject(xlWorkBook)
        releaseObject(xlWorkSheet)

        If icountTotalGridRows <= OnScreenRows Then
            TimerScrollGrid.Enabled = False
        Else
            TimerScrollGrid.Enabled = True
        End If

        NotifyIcon.ShowBalloonTip(5000, "InfoMerc", "Actualizando Display", ToolTipIcon.Info)
        DataGridView.FirstDisplayedScrollingRowIndex = 0
    End Sub

    Private Sub TimerScrollGrid_Tick(sender As Object, e As EventArgs) Handles TimerScrollGrid.Tick

        If ScrollRoundSteps = 0 Then
            DataGridView.FirstDisplayedScrollingRowIndex = 0
            ScrollRoundSteps = (ScrollRows / OnScreenRows)
        End If

        If ScrollRoundSteps > 0 Then
            DataGridView.FirstDisplayedScrollingRowIndex = DataGridView.FirstDisplayedScrollingRowIndex + OnScreenRows
            ScrollRoundSteps = (ScrollRoundSteps - 1)
        Else
            DataGridView.FirstDisplayedScrollingRowIndex = DataGridView.FirstDisplayedScrollingRowIndex + (ScrollMod - 3)
        End If

    End Sub
End Class





