unit MainForm;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants,
  System.Classes, Vcl.Graphics, Vcl.Controls, Vcl.Forms, Vcl.Dialogs,
  Vcl.ToolWin, Vcl.ComCtrls,
  Vcl.StdCtrls, Vcl.Menus, Vcl.Buttons, System.ImageList, Vcl.ImgList,
  Vcl.ExtCtrls,
  Vcl.TitleBarCtrls, Vcl.WinXCtrls, UDataModule,
  System.Generics.Collections, System.Net.URLClient, IniFiles, Vcl.Grids,
  FireDAC.Stan.Intf, FireDAC.Stan.Option, FireDAC.Stan.Param,
  FireDAC.Stan.Error, FireDAC.DatS, FireDAC.Phys.Intf, FireDAC.DApt.Intf,
  Data.DB, FireDAC.Comp.DataSet, FireDAC.Comp.Client,Excel2000, ComObj, DateUtils, JPEG, Clipbrd;


type
  TFormMain = class(TForm)
    clbToolButton: TCoolBar;
    panTitle: TPanel;
    labTitle: TLabel;
    panHeader: TPanel;
    tlbToolButton: TToolBar;
    mamMain: TMainMenu;
    File1: TMenuItem;
    Print1: TMenuItem;
    Exit1: TMenuItem;
    N1: TMenuItem;
    Edit1: TMenuItem;
    Copy1: TMenuItem;
    Display1: TMenuItem;
    Setting1: TMenuItem;
    btnClose: TSpeedButton;
    imlToolBar: TImageList;
    stbBase: TStatusBar;
    FDMemTable1: TFDMemTable;
    DateTimePickerEnd: TDateTimePicker;
    Label5: TLabel;
    SpeedButton1: TSpeedButton;
    ProgressBar1: TProgressBar;
    procedure Setting1Click(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure btnCloseClick(Sender: TObject);
    procedure btnPrintClick(Sender: TObject);
  private
    DataModuleCIMT: TDataModuleCIMT; // Instance of the data module
    FIsUseRESTAPIState: TToggleSwitchState;
    FDebugSQL: TToggleSwitchState;
    FUser: string;
    FPass: string;
    procedure ReadIniFile;
    function HasMonthlyData: Boolean;
    procedure WriteIniFile;
    procedure InitialProgram;
    procedure DebugSQLText(const SQLText: string);
    procedure InitConnection;
    procedure CreateMachineOKSheet(const Worksheet: OleVariant);
    procedure CreateNewMoldRepairSheet(const Worksheet: OleVariant);
    procedure CreateOverhaulSheet(const Worksheet: OleVariant);
    procedure CreateCleaningSheet(const Worksheet: OleVariant);
  public
  end;

var
  FormMain: TFormMain;

implementation

uses
  SettingForm;

{$R *.dfm}
{$R logo.res}

procedure LoadLogoFromResource(APicture: TPicture);
var
  ResStream: TResourceStream;
begin
  ResStream := TResourceStream.Create(HInstance, 'MYLOGO', RT_RCDATA);
  try
    APicture.LoadFromStream(ResStream);
  finally
    ResStream.Free;
  end;
end;

procedure InsertLogoToExcel(Worksheet: OleVariant);
var
  ResStream: TResourceStream;
  JpgImg: TJPEGImage;
  LogoPic: TPicture;
  LogoShape: OleVariant;
begin
  ResStream := TResourceStream.Create(HInstance, 'MYLOGO', RT_RCDATA);
  JpgImg := TJPEGImage.Create;
  LogoPic := TPicture.Create;
  try
    JpgImg.LoadFromStream(ResStream);
    LogoPic.Assign(JpgImg);
    Clipboard.Assign(LogoPic);

    Worksheet.Paste;
    LogoShape := Worksheet.Pictures(Worksheet.Pictures.Count);
    LogoShape.Left := Worksheet.Cells[1, 3].Left;
    LogoShape.Top := Worksheet.Cells[1, 2].Top;
    LogoShape.Width := 100;
    LogoShape.Height := 40;

  finally
    ResStream.Free;
    JpgImg.Free;
    LogoPic.Free;
  end;
end;


procedure TFormMain.Setting1Click(Sender: TObject);
begin
  if not Assigned(FormSetting) then
    FormSetting := TFormSetting.Create(Self);

  FormSetting.LoadSetting;

  if FormSetting.ShowModal = mrOk then
  begin
    ReadIniFile;
  end;
end;

procedure TFormMain.btnPrintClick(Sender: TObject);
var
  ExcelApp, Workbook: OleVariant;
begin
  if not HasMonthlyData then
  begin
    ShowMessage('No data found for this month.');
    Exit;
  end;

  // เริ่มต้น ProgressBar
  ProgressBar1.Position := 0;
  ProgressBar1.Max := 4;
  ProgressBar1.Visible := True;
  ProgressBar1.Refresh;

  ExcelApp := CreateOleObject('Excel.Application');
  ExcelApp.Visible := False;
  Workbook := ExcelApp.Workbooks.Add;

  Workbook.Worksheets[1].Name := 'MACHINE OK';
  Workbook.Worksheets[1].Tab.Color := $FFCC99;

  ProgressBar1.Position := 1;
  ProgressBar1.Refresh;
  CreateMachineOKSheet(Workbook.Worksheets['MACHINE OK']);

  Workbook.Worksheets.Add(After := Workbook.Sheets[Workbook.Sheets.Count]).Name := 'NEW MOLD,REPAIR';
  Workbook.Worksheets[2].Tab.Color := $FF99CC;

  ProgressBar1.Position := 2;
  ProgressBar1.Refresh;
  CreateNewMoldRepairSheet(Workbook.Worksheets['NEW MOLD,REPAIR']);

  Workbook.Worksheets.Add(After := Workbook.Sheets[Workbook.Sheets.Count]).Name := 'OVERHAUL';
  Workbook.Worksheets[3].Tab.Color := $99CCFF;

  ProgressBar1.Position := 3;
  ProgressBar1.Refresh;
  CreateOverhaulSheet(Workbook.Worksheets['OVERHAUL']);

  Workbook.Worksheets.Add(After := Workbook.Sheets[Workbook.Sheets.Count]).Name := 'CLEANING';
  Workbook.Worksheets[4].Tab.Color := $CCFFFF;

  ProgressBar1.Position := 4;
  ProgressBar1.Refresh;
  CreateCleaningSheet(Workbook.Worksheets['CLEANING']);

  ExcelApp.Visible := True;

  // ซ่อน ProgressBar เมื่อเสร็จ
  ProgressBar1.Visible := False;
end;




procedure TFormMain.CreateMachineOKSheet(const Worksheet: OleVariant);

  type
    TMoldKey = record
      MoldNo, ControlNo, JobType: string;
    end;

    TMachineDict = TDictionary<string, Double>;
    TMoldDict = TDictionary<string, TPair<TMoldKey, TMachineDict>>;

var
  SQLText: string;
  MemTable: TFDMemTable;
  DateBase: TDate;
  YearStr, MonthStr, StartDateStr, EndDateStr: string;
  ResultMessage: string;
  MoldNo, ControlNo, JobType, MachineNo, KeyStr, ColLetter: string;
  HourVal, TotalHour, GrandTotalHour, ColSum: Double;
  MoldKey: TMoldKey;
  MoldDict: TMoldDict;
  MachineDict: TMachineDict;
  Entry: TPair<TMoldKey, TMachineDict>;
  Pair: TPair<string, TPair<TMoldKey, TMachineDict>>;
  MachineList: TStringList;
  Row, I, MachineCol, LastCol, LastDataRow: Integer;
  ThaiMonthYear: string;
  LogoPic: TPicture;
  TempLogoPath: string;
  LogoShape: OleVariant;

begin
  DateBase := DateTimePickerEnd.Date;
  YearStr := FormatDateTime('yyyy', DateBase);
  MonthStr := FormatDateTime('mm', DateBase);
  StartDateStr := FormatDateTime('yyyy/mm/dd', EncodeDate(StrToInt(YearStr), StrToInt(MonthStr), 1));
  EndDateStr := FormatDateTime('yyyy/mm/dd', EndOfTheMonth(DateBase));

  SQLText :=
    'SELECT sm.katacd AS "MoldNo", sm.seizono AS "ControlNo", ' +
    'CASE ' +
    '  WHEN sm.kanryoymd IS NOT NULL AND sm.seizokbn = 11 THEN ''FG NEW MOLD'' ' +
    '  WHEN sm.kanryoymd IS NULL AND sm.seizokbn = 11 THEN ''FG NEW MOLD'' ' +
    '  WHEN sm.kanryoymd IS NOT NULL AND sm.seizokbn = 3 THEN ''FG REPAIR MOLD'' ' +
    '  WHEN sm.kanryoymd IS NULL AND sm.seizokbn = 3 THEN ''WIP REPAIR MOLD'' ' +
    'END AS "JobType", ' +
    'jd.kikaicd AS "MachineNo", TRUNC(SUM(jd.jh) / 60.00, 2) AS "Hour" ' +
    'FROM jisekidata jd ' +
    'INNER JOIN seizomst sm ON jd.seizono = sm.seizono AND jd.kikaicd LIKE ''M%'' ' +
    'INNER JOIN seizokbnmst sb ON sm.seizokbn = sb.seizokbn ' +
    'WHERE jd.ymds <= TO_DATE(''' + EndDateStr + ''', ''YYYY/MM/DD'') ' +
    '  AND jd.ymde >= TO_DATE(''' + StartDateStr + ''', ''YYYY/MM/DD'') ' +
    '  AND sm.seizono IS NOT NULL ' +
    'GROUP BY sm.katacd, sm.seizono, sm.kanryoymd, sm.seizokbn, jd.kikaicd ' +
    'ORDER BY "JobType", "MoldNo", "ControlNo", "MachineNo"';

  MemTable := TFDMemTable.Create(nil);
  try
    if FIsUseRESTAPIState = tssOff then
      DataModuleCIMT.FetchDataFromOracle(SQLText, MemTable)
    else
    begin
      ResultMessage := DataModuleCIMT.FetchDataFromREST(SQLText, MemTable);
      if ResultMessage <> 'Success' then
      begin
        ShowMessage(ResultMessage);
        Exit;
      end;
    end;

    if MemTable.IsEmpty then
    begin
      ShowMessage('No data found for this month.');
      Exit;
    end;

    MoldDict := TMoldDict.Create;
    MachineList := TStringList.Create;
    MachineList.Sorted := True;
    MachineList.Duplicates := dupIgnore;

    MemTable.First;
    while not MemTable.Eof do
    begin
      MoldNo := MemTable.FieldByName('MoldNo').AsString;
      ControlNo := MemTable.FieldByName('ControlNo').AsString;
      JobType := MemTable.FieldByName('JobType').AsString;
      MachineNo := MemTable.FieldByName('MachineNo').AsString;
      HourVal := MemTable.FieldByName('Hour').AsFloat;

      KeyStr := MoldNo + '|' + ControlNo + '|' + JobType;

      if not MoldDict.TryGetValue(KeyStr, Entry) then
      begin
        MoldKey.MoldNo := MoldNo;
        MoldKey.ControlNo := ControlNo;
        MoldKey.JobType := JobType;
        MachineDict := TMachineDict.Create;
        MoldDict.Add(KeyStr, TPair<TMoldKey, TMachineDict>.Create(MoldKey, MachineDict));
      end
      else
        MachineDict := Entry.Value;

      if MachineDict.ContainsKey(MachineNo) then
        MachineDict[MachineNo] := MachineDict[MachineNo] + HourVal
      else
        MachineDict.Add(MachineNo, HourVal);

      MachineList.Add(MachineNo);
      MemTable.Next;
    end;

    // === HEADERS ===

    InsertLogoToExcel(Worksheet);  // 👈 Add this line

    Worksheet.Cells[1, 1] := 'บริษัท นิฟโก้ (ไทยแลนด์) จำกัด';
    Worksheet.Range[Worksheet.Cells[1, 1], Worksheet.Cells[1, 3 + MachineList.Count + 2]].Merge;
    Worksheet.Range['A1'].HorizontalAlignment := -4108;
    Worksheet.Range['A1'].Font.Size := 24;
    Worksheet.Range['A1'].Font.Bold := True;

    ThaiMonthYear := 'REPORT SUMMARY MACHINE : ' + UpperCase(FormatDateTime('mmmm yyyy', DateBase));
    Worksheet.Cells[2, 1] := ThaiMonthYear;
    Worksheet.Range['A2'].HorizontalAlignment := -4131;
    Worksheet.Range['A2'].Font.Size := 16;
    Worksheet.Range['A2'].Font.Bold := True;

    Worksheet.Cells[4, 1] := 'MOLD NO.';
    Worksheet.Range['A4:A5'].Merge;
    Worksheet.Range['A4'].HorizontalAlignment := -4108;
    Worksheet.Range['A4'].VerticalAlignment := -4108;
    Worksheet.Range['A4'].Font.Bold := True;

    Worksheet.Cells[4, 2] := 'CONTROL NO.';
    Worksheet.Range['B4:B5'].Merge;
    Worksheet.Range['B4'].HorizontalAlignment := -4108;
    Worksheet.Range['B4'].VerticalAlignment := -4108;
    Worksheet.Range['B4'].Font.Bold := True;

    Worksheet.Cells[4, 3] := 'JOB TYPE';
    Worksheet.Range['C4:C5'].Merge;
    Worksheet.Range['C4'].HorizontalAlignment := -4108;
    Worksheet.Range['C4'].VerticalAlignment := -4108;
    Worksheet.Range['C4'].Font.Bold := True;

    Worksheet.Cells[5, 3] := 'MACHINE CODE';

    for I := 0 to MachineList.Count - 1 do
      Worksheet.Cells[5, 4 + I] := MachineList[I];

    Worksheet.Cells[4, 4] := 'MACHINE NO.';
    Worksheet.Range[Worksheet.Cells[4, 4], Worksheet.Cells[4, 3 + MachineList.Count]].Merge;
    Worksheet.Range[Worksheet.Cells[4, 4], Worksheet.Cells[4, 3 + MachineList.Count]].HorizontalAlignment := -4108;
    Worksheet.Range[Worksheet.Cells[4, 4], Worksheet.Cells[4, 3 + MachineList.Count]].Font.Bold := True;

    MachineCol := 4 + MachineList.Count;

    Worksheet.Cells[4, MachineCol] := 'GRAND TOTAL';
    Worksheet.Range[Worksheet.Cells[4, MachineCol], Worksheet.Cells[5, MachineCol]].Merge;
    Worksheet.Range[Worksheet.Cells[4, MachineCol], Worksheet.Cells[5, MachineCol]].HorizontalAlignment := -4108;
    Worksheet.Range[Worksheet.Cells[4, MachineCol], Worksheet.Cells[5, MachineCol]].VerticalAlignment := -4108;
    Worksheet.Range[Worksheet.Cells[4, MachineCol], Worksheet.Cells[5, MachineCol]].Font.Bold := True;
    Worksheet.Columns[MachineCol].ColumnWidth := 19;

    Inc(MachineCol);
    Worksheet.Cells[4, MachineCol] := 'AMOUNT COST';
    Worksheet.Range[Worksheet.Cells[4, MachineCol], Worksheet.Cells[5, MachineCol]].Merge;
    Worksheet.Range[Worksheet.Cells[4, MachineCol], Worksheet.Cells[5, MachineCol]].HorizontalAlignment := -4108;
    Worksheet.Range[Worksheet.Cells[4, MachineCol], Worksheet.Cells[5, MachineCol]].VerticalAlignment := -4108;
    Worksheet.Range[Worksheet.Cells[4, MachineCol], Worksheet.Cells[5, MachineCol]].Font.Bold := True;
    Worksheet.Columns[MachineCol].ColumnWidth := 19;

    LastCol := MachineCol;

    Worksheet.Rows[1].RowHeight := 50;
    Worksheet.Rows[2].RowHeight := 30;
    Worksheet.Rows[3].RowHeight := 9;
    Worksheet.Rows[4].RowHeight := 22;
    Worksheet.Rows[5].RowHeight := 22;

    Worksheet.Columns['A'].ColumnWidth := 12;
    Worksheet.Columns['B'].ColumnWidth := 14;
    Worksheet.Columns['C'].ColumnWidth := 18;

    Worksheet.Range[Worksheet.Cells[4, 1], Worksheet.Cells[5, LastCol]].Borders.LineStyle := 1;
    Worksheet.Range[Worksheet.Cells[4, 1], Worksheet.Cells[5, LastCol]].Borders.Weight := 2;

    // === DATA ===
    Row := 6;
    for Pair in MoldDict do
    begin
      Entry := Pair.Value;
      MoldKey := Entry.Key;
      MachineDict := Entry.Value;

      Worksheet.Cells[Row, 1] := MoldKey.MoldNo;
      Worksheet.Cells[Row, 2] := MoldKey.ControlNo;
      Worksheet.Cells[Row, 3] := MoldKey.JobType;

      TotalHour := 0;
      for I := 0 to MachineList.Count - 1 do
      begin
        MachineNo := MachineList[I];
        Worksheet.Cells[Row, 4 + I].NumberFormat := '#,##0.00';
        if MachineDict.TryGetValue(MachineNo, HourVal) then
          Worksheet.Cells[Row, 4 + I] := FormatFloat('#,##0.00', HourVal)
        else
          Worksheet.Cells[Row, 4 + I] := '0.00';

        TotalHour := TotalHour + HourVal;
      end;

      // Add left and right borders to the entire row
      for I := 1 to LastCol do
      begin
        Worksheet.Cells[Row, I].Borders[7].LineStyle := 1;  // xlEdgeLeft
        Worksheet.Cells[Row, I].Borders[10].LineStyle := 1; // xlEdgeRight
      end;

      Worksheet.Cells[Row, MachineCol - 1] := FormatFloat('#,##0.00', TotalHour);
      Worksheet.Cells[Row, MachineCol] := FormatFloat('#,##0.00', TotalHour *270 );
      Inc(Row);
    end;
     // === TOTAL ROW ===
    LastDataRow := Row - 1;
    GrandTotalHour := 0;

    Worksheet.Cells[Row, 1] := 'TOTAL';
    Worksheet.Range[Worksheet.Cells[Row, 1], Worksheet.Cells[Row, 2]].Merge;
    Worksheet.Range[Worksheet.Cells[Row, 1], Worksheet.Cells[Row, 2]].Font.Bold := True;

    for I := 0 to MachineList.Count - 1 do
    begin
      ColLetter := Chr(Ord('D') + I); // Column D onwards
      ColSum := Worksheet.Application.WorksheetFunction.Sum(
        Worksheet.Range[ColLetter + '6', ColLetter + IntToStr(LastDataRow)]
      );
      Worksheet.Cells[Row, 4 + I] := FormatFloat('#,##0.00', ColSum);
      GrandTotalHour := GrandTotalHour + ColSum;
    end;

    Worksheet.Cells[Row, MachineCol - 1] := FormatFloat('#,##0.00', GrandTotalHour);
    Worksheet.Cells[Row, MachineCol] := FormatFloat('#,##0.00', GrandTotalHour * 270);

    Worksheet.Range[Worksheet.Cells[Row, 1], Worksheet.Cells[Row, LastCol]].Font.Bold := True;
    Worksheet.Range[Worksheet.Cells[Row, 1], Worksheet.Cells[Row, LastCol]].Borders.LineStyle := 1;

    //APPROVED BY,CHECKED BY,ISSUED BY
    var  SignatureRow: Integer;

    inc(Row);
    inc(Row);
    // After TOTAL row is written (Row already incremented)
    SignatureRow := Row;

    // Set height of signature row
    Worksheet.Rows[SignatureRow+1].RowHeight := 116;

    // Merge and label the signature cells
    // Last machine column starts at column 4 → MachineList.Count columns used
    Worksheet.Range[Worksheet.Cells[SignatureRow, 4 + MachineList.Count - 1], Worksheet.Cells[SignatureRow, 4 + MachineList.Count - 1]].Merge;
    Worksheet.Range[Worksheet.Cells[SignatureRow, 4 + MachineList.Count - 1], Worksheet.Cells[SignatureRow, 4 + MachineList.Count - 1]].Value := 'ISSUED BY';

    Worksheet.Range[Worksheet.Cells[SignatureRow, MachineCol - 1], Worksheet.Cells[SignatureRow, MachineCol - 1]].Merge;
    Worksheet.Range[Worksheet.Cells[SignatureRow, MachineCol - 1], Worksheet.Cells[SignatureRow, MachineCol - 1]].Value := 'CHECKED BY';

    Worksheet.Range[Worksheet.Cells[SignatureRow, MachineCol], Worksheet.Cells[SignatureRow, MachineCol]].Merge;
    Worksheet.Range[Worksheet.Cells[SignatureRow, MachineCol], Worksheet.Cells[SignatureRow, MachineCol]].Value := 'APPROVED BY';

    // Apply center alignment and bold font
    Worksheet.Range[Worksheet.Cells[SignatureRow, 4 + MachineList.Count - 1], Worksheet.Cells[SignatureRow, MachineCol]].HorizontalAlignment := -4108;
    Worksheet.Range[Worksheet.Cells[SignatureRow, 4 + MachineList.Count - 1], Worksheet.Cells[SignatureRow, MachineCol]].Font.Bold := True;

    // Add border boxes
    Worksheet.Range[Worksheet.Cells[SignatureRow, 4 + MachineList.Count - 1], Worksheet.Cells[SignatureRow, 4 + MachineList.Count - 1]].Borders.LineStyle := 1;
    Worksheet.Range[Worksheet.Cells[SignatureRow, MachineCol - 1], Worksheet.Cells[SignatureRow, MachineCol - 1]].Borders.LineStyle := 1;
    Worksheet.Range[Worksheet.Cells[SignatureRow, MachineCol], Worksheet.Cells[SignatureRow, MachineCol]].Borders.LineStyle := 1;

    inc(SignatureRow);
    // Add border boxes
    Worksheet.Range[Worksheet.Cells[SignatureRow, 4 + MachineList.Count - 1], Worksheet.Cells[SignatureRow, 4 + MachineList.Count - 1]].Borders.LineStyle := 1;
    Worksheet.Columns[MachineCol - 2].ColumnWidth := 19;
    Worksheet.Range[Worksheet.Cells[SignatureRow, MachineCol - 1], Worksheet.Cells[SignatureRow, MachineCol - 1]].Borders.LineStyle := 1;
    Worksheet.Range[Worksheet.Cells[SignatureRow, MachineCol], Worksheet.Cells[SignatureRow, MachineCol]].Borders.LineStyle := 1;

    Worksheet.Activate;
    Worksheet.Cells[6, 1].Select;
    Worksheet.Application.ActiveWindow.FreezePanes := True;

    var
      DataRange, SheetRange: OleVariant;
    begin
      // กำหนดขนาดพื้นที่ทั้งแผ่น (เช่น 1 ถึง 1000 แถว และ 1 ถึง 100 คอลัมน์)
      SheetRange := Worksheet.Range[Worksheet.Cells[1, 1], Worksheet.Cells[1000, 100]];
      SheetRange.Interior.Color := $C0C0C0; // สีเทาทั้งพื้นหลัง

      // กำหนดขอบเขตที่มีข้อมูลให้เป็นสีขาว
      DataRange := Worksheet.Range[Worksheet.Cells[1, 1], Worksheet.Cells[Row+1, LastCol]];
      DataRange.Interior.Color := clWhite;
    end;

    // === Cleanup ===
    for Entry in MoldDict.Values do
      Entry.Value.Free;
    MoldDict.Free;
    MachineList.Free;

  finally
    MemTable.Free;
  end;
end;


procedure TFormMain.CreateNewMoldRepairSheet(const Worksheet: OleVariant);
type
  TMoldKey = record
    MoldNo, ControlNo, JobType: string;
  end;

  TWorkerDict = TDictionary<string, Double>;
  TMoldDict = TDictionary<string, TPair<TMoldKey, TWorkerDict>>;

var
  SQLText, WorkerSQL: string;
  MemTable, WorkerTable: TFDMemTable;
  DateBase: TDate;
  YearStr, MonthStr, StartDateStr, EndDateStr, ThaiMonthYear: string;
  ResultMessage: string;
  MoldNo, ControlNo, JobType, WorkerName, KeyStr: string;
  HourVal, TotalHour: Double;
  MoldKey: TMoldKey;
  MoldDict: TMoldDict;
  WorkerDict: TWorkerDict;
  Entry: TPair<TMoldKey, TWorkerDict>;
  Pair: TPair<string, TPair<TMoldKey, TWorkerDict>>;
  WorkerListDS, WorkerListMT: TStringList;
  AllWorkerList: TStringList;
  Row, I, ColIndex, LastCol: Integer;
begin
  DateBase := DateTimePickerEnd.Date;
  YearStr := FormatDateTime('yyyy', DateBase);
  MonthStr := FormatDateTime('mm', DateBase);
  StartDateStr := FormatDateTime('yyyy/mm/dd', EncodeDate(StrToInt(YearStr), StrToInt(MonthStr), 1));
  EndDateStr := FormatDateTime('yyyy/mm/dd', EndOfTheMonth(DateBase));

  SQLText :=
    'SELECT sm.katacd AS "MoldNo", sm.seizono AS "ControlNo", ' +
    '  CASE ' +
    '    WHEN sm.kanryoymd IS NOT NULL AND sm.seizokbn = 11 THEN ''FG NEW MOLD'' ' +
    '    WHEN sm.kanryoymd IS NULL AND sm.seizokbn = 11 THEN ''FG NEW MOLD'' ' +
    '    WHEN sm.kanryoymd IS NOT NULL AND sm.seizokbn = 3 THEN ''FG REPAIR MOLD'' ' +
    '    WHEN sm.kanryoymd IS NULL AND sm.seizokbn = 3 THEN ''WIP REPAIR MOLD'' ' +
    '  END AS "JobType", ' +
    '  tm.tantonm AS "WorkerName", ' +
    '  TRUNC(SUM(jd.jh) / 60.00, 2) AS "Hour"  ' +
    'FROM jisekidata jd ' +
    'INNER JOIN seizomst sm ON jd.seizono = sm.seizono ' +
    'INNER JOIN tantomst tm ON jd.tantocd = tm.tantocd ' +
    'WHERE tm.tantogrcd IN (''DS'', ''MT'') ' +
    '  AND jd.ymds <= TO_DATE(''' + EndDateStr + ''', ''YYYY/MM/DD'') ' +
    '  AND jd.ymde >= TO_DATE(''' + StartDateStr + ''', ''YYYY/MM/DD'') ' +
    '  AND sm.seizono IS NOT NULL ' +
    'GROUP BY sm.katacd, sm.seizono, sm.kanryoymd, sm.seizokbn, tm.tantonm, tm.tantogrcd ' +
    'ORDER BY CASE tm.tantogrcd WHEN ''DS'' THEN 0 WHEN ''MT'' THEN 1 ELSE 2 END, ' +
    '  "JobType", "MoldNo", "ControlNo", "WorkerName"';

  MemTable := TFDMemTable.Create(nil);
  WorkerTable := TFDMemTable.Create(nil);
  WorkerListDS := TStringList.Create;
  WorkerListMT := TStringList.Create;
  AllWorkerList := TStringList.Create;
  WorkerListDS.Sorted := False;
  WorkerListMT.Sorted := False;
  AllWorkerList.Sorted := False;
  WorkerListDS.Duplicates := dupIgnore;
  WorkerListMT.Duplicates := dupIgnore;

  try
    // Fetch usage data
    if FIsUseRESTAPIState = tssOff then
      DataModuleCIMT.FetchDataFromOracle(SQLText, MemTable)
    else
    begin
      ResultMessage := DataModuleCIMT.FetchDataFromREST(SQLText, MemTable);
      if ResultMessage <> 'Success' then
      begin
        ShowMessage(ResultMessage);
        Exit;
      end;
    end;

    // Fetch full list of workers
    WorkerSQL := 'SELECT tantonm, tantogrcd FROM tantomst WHERE tantogrcd IN (''DS'', ''MT'') ' +
                 'ORDER BY CASE tantogrcd WHEN ''DS'' THEN 0 WHEN ''MT'' THEN 1 ELSE 2 END, tantonm';
    if FIsUseRESTAPIState = tssOff then
      DataModuleCIMT.FetchDataFromOracle(WorkerSQL, WorkerTable)
    else
    begin
      ResultMessage := DataModuleCIMT.FetchDataFromREST(WorkerSQL, WorkerTable);
      if ResultMessage <> 'Success' then
      begin
        ShowMessage(ResultMessage);
        Exit;
      end;
    end;

    WorkerTable.First;
    while not WorkerTable.Eof do
    begin
      WorkerName := WorkerTable.FieldByName('tantonm').AsString;
      if WorkerTable.FieldByName('tantogrcd').AsString = 'DS' then
        WorkerListDS.Add(WorkerName)
      else
        WorkerListMT.Add(WorkerName);
      AllWorkerList.Add(WorkerName);
      WorkerTable.Next;
    end;

    if not MemTable.Active then
      MemTable.Open;

    if MemTable.IsEmpty then
    begin
      ShowMessage('No data found for NewMold.');
      Exit;
    end;

    // Map usage
    MoldDict := TMoldDict.Create;
    MemTable.First;
    while not MemTable.Eof do
    begin
      MoldNo := MemTable.FieldByName('MoldNo').AsString;
      ControlNo := MemTable.FieldByName('ControlNo').AsString;
      JobType := MemTable.FieldByName('JobType').AsString;
      WorkerName := MemTable.FieldByName('WorkerName').AsString;
      HourVal := MemTable.FieldByName('Hour').AsFloat;

      KeyStr := MoldNo + '|' + ControlNo + '|' + JobType;
      if not MoldDict.TryGetValue(KeyStr, Entry) then
      begin
        MoldKey.MoldNo := MoldNo;
        MoldKey.ControlNo := ControlNo;
        MoldKey.JobType := JobType;
        WorkerDict := TWorkerDict.Create;
        MoldDict.Add(KeyStr, TPair<TMoldKey, TWorkerDict>.Create(MoldKey, WorkerDict));
      end
      else
        WorkerDict := Entry.Value;

      if WorkerDict.ContainsKey(WorkerName) then
        WorkerDict[WorkerName] := WorkerDict[WorkerName] + HourVal
      else
        WorkerDict.Add(WorkerName, HourVal);

      MemTable.Next;
    end;

    // === Headers ===
    InsertLogoToExcel(Worksheet);  // 👈 Add this line


    Worksheet.Cells[1, 1] := 'บริษัท นิฟโก้ (ไทยแลนด์) จำกัด';
    Worksheet.Range[Worksheet.Cells[1, 1], Worksheet.Cells[1, 3 + AllWorkerList.Count + 2]].Merge;
    Worksheet.Range['A1'].HorizontalAlignment := -4108;
    Worksheet.Range['A1'].Font.Size := 24;
    Worksheet.Range['A1'].Font.Bold := True;

    ThaiMonthYear := 'REPORT SUMMARY WORKER : ' + UpperCase(FormatDateTime('mmmm yyyy', DateBase));
    Worksheet.Cells[2, 1] := ThaiMonthYear;
    Worksheet.Range['A2'].HorizontalAlignment := -4131;
    Worksheet.Range['A2'].Font.Size := 16;
    Worksheet.Range['A2'].Font.Bold := True;

    Worksheet.Cells[4, 1] := 'MOLD NO.';
    Worksheet.Range['A4:A5'].Merge;
    Worksheet.Range['A4'].HorizontalAlignment := -4108;
    Worksheet.Range['A4'].VerticalAlignment := -4108;
    Worksheet.Range['A4'].Font.Bold := True;

    Worksheet.Cells[4, 2] := 'CONTROL NO.';
    Worksheet.Range['B4:B5'].Merge;
    Worksheet.Range['B4'].HorizontalAlignment := -4108;
    Worksheet.Range['B4'].VerticalAlignment := -4108;
    Worksheet.Range['B4'].Font.Bold := True;

    Worksheet.Cells[4, 3] := 'JOB TYPE';
    Worksheet.Range['C4:C5'].Merge;
    Worksheet.Range['C4'].HorizontalAlignment := -4108;
    Worksheet.Range['C4'].VerticalAlignment := -4108;
    Worksheet.Range['C4'].Font.Bold := True;

    ColIndex := 4;

    // DESIGN
    if WorkerListDS.Count > 0 then
    begin
      Worksheet.Cells[4, ColIndex] := 'DESIGN';
      Worksheet.Range[Worksheet.Cells[4, ColIndex], Worksheet.Cells[4, ColIndex + WorkerListDS.Count - 1]].Merge;
      Worksheet.Range[Worksheet.Cells[4, ColIndex], Worksheet.Cells[4, ColIndex + WorkerListDS.Count - 1]].HorizontalAlignment := -4108;
      Worksheet.Range[Worksheet.Cells[4, ColIndex], Worksheet.Cells[4, ColIndex + WorkerListDS.Count - 1]].Font.Bold := True;
      for I := 0 to WorkerListDS.Count - 1 do
        Worksheet.Cells[5, ColIndex + I] := WorkerListDS[I];
      Inc(ColIndex, WorkerListDS.Count);
    end;


    // MOLD PRODUCTION AND MAINTENANCE
    if WorkerListMT.Count > 0 then
    begin
      Worksheet.Cells[4, ColIndex] := 'MOLD PRODUCTION AND MOLD MAINTENANCE';
      Worksheet.Range[Worksheet.Cells[4, ColIndex], Worksheet.Cells[4, ColIndex + WorkerListMT.Count - 1]].Merge;
      Worksheet.Range[Worksheet.Cells[4, ColIndex], Worksheet.Cells[4, ColIndex + WorkerListMT.Count - 1]].HorizontalAlignment := -4108;
      Worksheet.Range[Worksheet.Cells[4, ColIndex], Worksheet.Cells[4, ColIndex + WorkerListMT.Count - 1]].Font.Bold := True;
      for I := 0 to WorkerListMT.Count - 1 do
        Worksheet.Cells[5, ColIndex + I] := WorkerListMT[I];
      Inc(ColIndex, WorkerListMT.Count);
    end;


    // GRAND TOTAL + AMOUNT COST
    Worksheet.Cells[4, ColIndex] := 'GRAND TOTAL';
    Worksheet.Range[Worksheet.Cells[4, ColIndex], Worksheet.Cells[5, ColIndex]].HorizontalAlignment := -4108;
    Worksheet.Range[Worksheet.Cells[4, ColIndex], Worksheet.Cells[5, ColIndex]].VerticalAlignment := -4108;
    Worksheet.Range[Worksheet.Cells[4, ColIndex], Worksheet.Cells[5, ColIndex]].Font.Bold := True;
    Worksheet.Range[Worksheet.Cells[4, ColIndex], Worksheet.Cells[5, ColIndex]].Merge;
    Worksheet.Columns[ColIndex].ColumnWidth := 20;

    Inc(ColIndex);
    Worksheet.Cells[4, ColIndex] := 'AMOUNT COST';
    Worksheet.Range[Worksheet.Cells[4, ColIndex], Worksheet.Cells[5, ColIndex]].HorizontalAlignment := -4108;
    Worksheet.Range[Worksheet.Cells[4, ColIndex], Worksheet.Cells[5, ColIndex]].VerticalAlignment := -4108;
    Worksheet.Range[Worksheet.Cells[4, ColIndex], Worksheet.Cells[5, ColIndex]].Font.Bold := True;
    Worksheet.Range[Worksheet.Cells[4, ColIndex], Worksheet.Cells[5, ColIndex]].Merge;
    Worksheet.Columns[ColIndex].ColumnWidth := 20;
    LastCol := ColIndex;

    // Formatting
    Worksheet.Rows[1].RowHeight := 50;
    Worksheet.Rows[2].RowHeight := 30;
    Worksheet.Rows[3].RowHeight := 9;
    Worksheet.Rows[4].RowHeight := 22;
    Worksheet.Rows[5].RowHeight := 22;
    Worksheet.Columns['A'].ColumnWidth := 12;
    Worksheet.Columns['B'].ColumnWidth := 14;
    Worksheet.Columns['C'].ColumnWidth := 18;
    Worksheet.Range[Worksheet.Cells[4, 1], Worksheet.Cells[5, LastCol]].Borders.LineStyle := 1;

    // === Data rows ===
    Row := 6;
    for Pair in MoldDict do
    begin
      Entry := Pair.Value;
      MoldKey := Entry.Key;
      WorkerDict := Entry.Value;

      Worksheet.Cells[Row, 1] := MoldKey.MoldNo;
      Worksheet.Cells[Row, 2] := MoldKey.ControlNo;
      Worksheet.Cells[Row, 3] := MoldKey.JobType;

      TotalHour := 0;
      for I := 0 to AllWorkerList.Count - 1 do
      begin
        WorkerName := AllWorkerList[I];
        Worksheet.Cells[Row, 4 + I].NumberFormat := '#,##0.00';
        if WorkerDict.TryGetValue(WorkerName, HourVal) then
          Worksheet.Cells[Row, 4 + I] := FormatFloat('#,##0.00', HourVal)
        else
          Worksheet.Cells[Row, 4 + I] := '0.00';
        TotalHour := TotalHour + HourVal;
      end;

      // Add left and right borders to the entire row
      for I := 1 to LastCol do
      begin
        Worksheet.Cells[Row, I].Borders[7].LineStyle := 1;  // xlEdgeLeft
        Worksheet.Cells[Row, I].Borders[10].LineStyle := 1; // xlEdgeRight
      end;

      Worksheet.Cells[Row, LastCol - 1] := FormatFloat('#,##0.00', TotalHour);
      Worksheet.Cells[Row, LastCol] := FormatFloat('#,##0.00', TotalHour * 390);
      Inc(Row);


    end;

    var LastDataRow,GrandTotalHour,ColSum: Integer;
    var ColLetter: string;

    // === TOTAL ROW ===
    LastDataRow := Row - 1;
    GrandTotalHour := 0;

    Worksheet.Cells[Row, 1] := 'TOTAL';
    Worksheet.Range[Worksheet.Cells[Row, 1], Worksheet.Cells[Row, 2]].Merge;
    Worksheet.Range[Worksheet.Cells[Row, 1], Worksheet.Cells[Row, 2]].Font.Bold := True;

    for I := 0 to AllWorkerList.Count - 1 do
    begin
      ColLetter := Chr(Ord('D') + I); // Start from D
      ColSum := Worksheet.Application.WorksheetFunction.Sum(
        Worksheet.Range[ColLetter + '6', ColLetter + IntToStr(LastDataRow)]
      );
      Worksheet.Cells[Row, 4 + I] := FormatFloat('#,##0.00', ColSum);
      GrandTotalHour := GrandTotalHour + ColSum;
    end;

    Worksheet.Cells[Row, LastCol - 1] := FormatFloat('#,##0.00', GrandTotalHour);
    Worksheet.Cells[Row, LastCol] := FormatFloat('#,##0.00', GrandTotalHour * 390);
    Worksheet.Range[Worksheet.Cells[Row, 1], Worksheet.Cells[Row, LastCol]].Font.Bold := True;
    Worksheet.Range[Worksheet.Cells[Row, 1], Worksheet.Cells[Row, LastCol]].Borders.LineStyle := 1;

    // === Signature Block ===
    Inc(Row);
    Inc(Row);
    var SignatureRow: Integer := Row;
    Worksheet.Rows[SignatureRow + 1].RowHeight := 116;

    var IssuedCol := LastCol - 2;
    var CheckedCol := LastCol - 1;
    var ApprovedCol := LastCol;

    Worksheet.Columns[LastCol - 2].ColumnWidth := 19;

    Worksheet.Range[Worksheet.Cells[SignatureRow, IssuedCol], Worksheet.Cells[SignatureRow, IssuedCol]].Merge;
    Worksheet.Cells[SignatureRow, IssuedCol] := 'ISSUED BY';

    Worksheet.Range[Worksheet.Cells[SignatureRow, CheckedCol], Worksheet.Cells[SignatureRow, CheckedCol]].Merge;
    Worksheet.Cells[SignatureRow, CheckedCol] := 'CHECKED BY';

    Worksheet.Range[Worksheet.Cells[SignatureRow, ApprovedCol], Worksheet.Cells[SignatureRow, ApprovedCol]].Merge;
    Worksheet.Cells[SignatureRow, ApprovedCol] := 'APPROVED BY';

    Worksheet.Range[Worksheet.Cells[SignatureRow, IssuedCol], Worksheet.Cells[SignatureRow, ApprovedCol]].HorizontalAlignment := -4108;
    Worksheet.Range[Worksheet.Cells[SignatureRow, IssuedCol], Worksheet.Cells[SignatureRow, ApprovedCol]].Font.Bold := True;

    Worksheet.Range[Worksheet.Cells[SignatureRow, IssuedCol], Worksheet.Cells[SignatureRow, IssuedCol]].Borders.LineStyle := 1;
    Worksheet.Range[Worksheet.Cells[SignatureRow, CheckedCol], Worksheet.Cells[SignatureRow, CheckedCol]].Borders.LineStyle := 1;
    Worksheet.Range[Worksheet.Cells[SignatureRow, ApprovedCol], Worksheet.Cells[SignatureRow, ApprovedCol]].Borders.LineStyle := 1;

    // Add bottom border cells for spacing
    Inc(SignatureRow);
    Worksheet.Range[Worksheet.Cells[SignatureRow, IssuedCol], Worksheet.Cells[SignatureRow, IssuedCol]].Borders.LineStyle := 1;
    Worksheet.Range[Worksheet.Cells[SignatureRow, CheckedCol], Worksheet.Cells[SignatureRow, CheckedCol]].Borders.LineStyle := 1;
    Worksheet.Range[Worksheet.Cells[SignatureRow, ApprovedCol], Worksheet.Cells[SignatureRow, ApprovedCol]].Borders.LineStyle := 1;

    // === Freeze pane ===
    Worksheet.Activate;
    Worksheet.Cells[6, 1].Select;
    Worksheet.Application.ActiveWindow.FreezePanes := True;

    var DataRange, SheetRange: OleVariant;
    begin
      // กำหนดขนาดพื้นที่ทั้งแผ่น (เช่น 1 ถึง 1000 แถว และ 1 ถึง 100 คอลัมน์)
      SheetRange := Worksheet.Range[Worksheet.Cells[1, 1], Worksheet.Cells[1000, 100]];
      SheetRange.Interior.Color := $C0C0C0; // สีเทาทั้งพื้นหลัง

      // กำหนดขอบเขตที่มีข้อมูลให้เป็นสีขาว
      DataRange := Worksheet.Range[Worksheet.Cells[1, 1], Worksheet.Cells[Row+1, LastCol]];
      DataRange.Interior.Color := clWhite;
    end;

    // Cleanup
    for Entry in MoldDict.Values do
      Entry.Value.Free;
    MoldDict.Free;
  finally
    MemTable.Free;
    WorkerTable.Free;
    WorkerListDS.Free;
    WorkerListMT.Free;
    AllWorkerList.Free;
  end;
end;


procedure TFormMain.CreateOverhaulSheet(const Worksheet: OleVariant);
type
  TMoldKey = record
    MoldNo, ControlNo: string;
  end;

  TWorkerDict = TDictionary<string, Double>;
  TMoldDict = TDictionary<string, TPair<TMoldKey, TWorkerDict>>;

var
  SQLText, WorkerSQL, ResultMessage: string;
  MemTable, WorkerTable: TFDMemTable;
  DateBase: TDate;
  YearStr, MonthStr, StartDateStr, EndDateStr, ThaiMonthYear: string;
  MoldNo, ControlNo, WorkerName, KeyStr: string;
  HourVal, TotalHour: Double;
  MoldKey: TMoldKey;
  MoldDict: TMoldDict;
  WorkerDict: TWorkerDict;
  Entry: TPair<TMoldKey, TWorkerDict>;
  Pair: TPair<string, TPair<TMoldKey, TWorkerDict>>;
  WorkerListMT: TStringList;
  Row, I, ColIndex, LastCol, ItemNo: Integer;
begin
  DateBase := DateTimePickerEnd.Date;
  YearStr := FormatDateTime('yyyy', DateBase);
  MonthStr := FormatDateTime('mm', DateBase);
  StartDateStr := FormatDateTime('yyyy/mm/dd', EncodeDate(StrToInt(YearStr), StrToInt(MonthStr), 1));
  EndDateStr := FormatDateTime('yyyy/mm/dd', EndOfTheMonth(DateBase));

  SQLText :=
    'SELECT sm.katacd AS "MoldNo", sm.seizono AS "ControlNo", ' +
    '  tm.tantonm AS "WorkerName", TRUNC(SUM(jd.jh) / 60.00, 2) AS "Hour"  ' +
    'FROM jisekidata jd ' +
    'INNER JOIN seizomst sm ON jd.seizono = sm.seizono ' +
    'INNER JOIN tantomst tm ON jd.tantocd = tm.tantocd ' +
    'WHERE tm.tantogrcd = ''MT'' AND keikoteicd = ''OH'' ' +
    '  AND jd.ymds <= TO_DATE(''' + EndDateStr + ''', ''YYYY/MM/DD'') ' +
    '  AND jd.ymde >= TO_DATE(''' + StartDateStr + ''', ''YYYY/MM/DD'') ' +
    '  AND sm.seizono IS NOT NULL ' +
    'GROUP BY sm.katacd, sm.seizono, tm.tantonm ' +
    'ORDER BY "MoldNo", "ControlNo", "WorkerName"';

  WorkerSQL :=
    'SELECT tantonm FROM tantomst WHERE tantogrcd = ''MT'' ORDER BY tantonm';

  MemTable := TFDMemTable.Create(nil);
  WorkerTable := TFDMemTable.Create(nil);
  WorkerListMT := TStringList.Create;
  WorkerListMT.Duplicates := dupIgnore;

  try
    if FIsUseRESTAPIState = tssOff then
      DataModuleCIMT.FetchDataFromOracle(WorkerSQL, WorkerTable)
    else
    begin
      ResultMessage := DataModuleCIMT.FetchDataFromREST(WorkerSQL, WorkerTable);
      if ResultMessage <> 'Success' then Exit;
    end;

    WorkerTable.First;
    while not WorkerTable.Eof do
    begin
      WorkerListMT.Add(WorkerTable.FieldByName('tantonm').AsString);
      WorkerTable.Next;
    end;

    if FIsUseRESTAPIState = tssOff then
      DataModuleCIMT.FetchDataFromOracle(SQLText, MemTable)
    else
    begin
      ResultMessage := DataModuleCIMT.FetchDataFromREST(SQLText, MemTable);
      if ResultMessage <> 'Success' then
      begin
        ShowMessage('Data fetch error: ' + ResultMessage);
        Exit;
      end;
    end;

    if MemTable.IsEmpty then
    begin
      ShowMessage('No data found for OVERHAUL.');
      Exit;
    end;

    // ✅ ตอนนี้ปลอดภัยที่จะ MemTable.First ได้
    MemTable.First;

    MoldDict := TMoldDict.Create;
    while not MemTable.Eof do
    begin
      MoldNo := MemTable.FieldByName('MoldNo').AsString;
      ControlNo := MemTable.FieldByName('ControlNo').AsString;
      WorkerName := MemTable.FieldByName('WorkerName').AsString;
      HourVal := MemTable.FieldByName('Hour').AsFloat;

      KeyStr := MoldNo + '|' + ControlNo;
      if not MoldDict.TryGetValue(KeyStr, Entry) then
      begin
        MoldKey.MoldNo := MoldNo;
        MoldKey.ControlNo := ControlNo;
        WorkerDict := TWorkerDict.Create;
        MoldDict.Add(KeyStr, TPair<TMoldKey, TWorkerDict>.Create(MoldKey, WorkerDict));
      end
      else
        WorkerDict := Entry.Value;

      if WorkerDict.ContainsKey(WorkerName) then
        WorkerDict[WorkerName] := WorkerDict[WorkerName] + HourVal
      else
        WorkerDict.Add(WorkerName, HourVal);

      MemTable.Next;
    end;

    // === Headers ===

    InsertLogoToExcel(Worksheet);  // 👈 Add this line

    Worksheet.Cells[1, 1] := 'บริษัท นิฟโก้ (ไทยแลนด์) จำกัด';
    Worksheet.Range[Worksheet.Cells[1, 1], Worksheet.Cells[1, 3 + WorkerListMT.Count + 2]].Merge;
    Worksheet.Range['A1'].HorizontalAlignment := -4108;
    Worksheet.Range['A1'].Font.Size := 24;
    Worksheet.Range['A1'].Font.Bold := True;

    ThaiMonthYear := 'REPORT SUMMARY WORKER : ' + UpperCase(FormatDateTime('mmmm yyyy', DateBase));
    Worksheet.Cells[2, 1] := ThaiMonthYear;
    Worksheet.Range['A2'].HorizontalAlignment := -4131;
    Worksheet.Range['A2'].Font.Size := 16;
    Worksheet.Range['A2'].Font.Bold := True;

    Worksheet.Cells[3, 1] := 'OVERHAUL';
    Worksheet.Range['A3'].Font.Size := 16;
    Worksheet.Range['A3'].Font.Bold := True;

    Worksheet.Cells[4, 1] := 'ITEM';
    Worksheet.Range['A4:A5'].Merge;
    Worksheet.Range['A4'].HorizontalAlignment := -4108;
    Worksheet.Range['A4'].VerticalAlignment := -4108;
    Worksheet.Range['A4'].Font.Bold := True;

    Worksheet.Cells[4, 2] := 'MOLD NO.';
    Worksheet.Range['B4:B5'].Merge;
    Worksheet.Range['B4'].HorizontalAlignment := -4108;
    Worksheet.Range['B4'].VerticalAlignment := -4108;
    Worksheet.Range['B4'].Font.Bold := True;

    Worksheet.Cells[4, 3] := 'CONTROL NO.';
    Worksheet.Range['C4:C5'].Merge;
    Worksheet.Range['C4'].HorizontalAlignment := -4108;
    Worksheet.Range['C4'].VerticalAlignment := -4108;
    Worksheet.Range['C4'].Font.Bold := True;

  ColIndex := 4;
  if WorkerListMT.Count > 0 then
  begin
    for I := 0 to WorkerListMT.Count - 1 do
    begin
      Worksheet.Cells[4, ColIndex + I] := WorkerListMT[I];
      Worksheet.Cells[5, ColIndex + I] := 'TIME';
      Worksheet.Range[Worksheet.Cells[4, ColIndex + I], Worksheet.Cells[4, ColIndex + I]].HorizontalAlignment := -4108;
      Worksheet.Range[Worksheet.Cells[5, ColIndex + I], Worksheet.Cells[5, ColIndex + I]].HorizontalAlignment := -4108;
      Worksheet.Columns[ColIndex + I].ColumnWidth := 15;
    end;
    Inc(ColIndex, WorkerListMT.Count);
  end;


    Worksheet.Cells[4, ColIndex] := 'GRAND TOTAL';
    Worksheet.Range[Worksheet.Cells[4, ColIndex], Worksheet.Cells[5, ColIndex]].Merge;
    Worksheet.Range[Worksheet.Cells[4, ColIndex], Worksheet.Cells[5, ColIndex]].HorizontalAlignment := -4108;
    Worksheet.Range[Worksheet.Cells[4, ColIndex], Worksheet.Cells[5, ColIndex]].VerticalAlignment := -4108;
    Worksheet.Range[Worksheet.Cells[4, ColIndex], Worksheet.Cells[5, ColIndex]].Font.Bold := True;
    Worksheet.Range[Worksheet.Cells[4, ColIndex], Worksheet.Cells[5, ColIndex]].Merge;
    Worksheet.Columns[ColIndex].ColumnWidth := 20;
    Inc(ColIndex);
    Worksheet.Cells[4, ColIndex] := 'AMOUNT COST';
    Worksheet.Range[Worksheet.Cells[4, ColIndex], Worksheet.Cells[5, ColIndex]].Merge;
    Worksheet.Range[Worksheet.Cells[4, ColIndex], Worksheet.Cells[5, ColIndex]].HorizontalAlignment := -4108;
    Worksheet.Range[Worksheet.Cells[4, ColIndex], Worksheet.Cells[5, ColIndex]].VerticalAlignment := -4108;
    Worksheet.Range[Worksheet.Cells[4, ColIndex], Worksheet.Cells[5, ColIndex]].Font.Bold := True;
    Worksheet.Range[Worksheet.Cells[4, ColIndex], Worksheet.Cells[5, ColIndex]].Merge;
    Worksheet.Columns[ColIndex].ColumnWidth := 20;
    LastCol := ColIndex;

    // Formatting
    Worksheet.Rows[1].RowHeight := 50;
    Worksheet.Rows[2].RowHeight := 30;
    Worksheet.Rows[3].RowHeight := 20;
    Worksheet.Rows[4].RowHeight := 22;
    Worksheet.Rows[5].RowHeight := 22;
    Worksheet.Columns['A'].ColumnWidth := 10;
    Worksheet.Columns['B'].ColumnWidth := 14;
    Worksheet.Columns['C'].ColumnWidth := 14;
    Worksheet.Range[Worksheet.Cells[4, 1], Worksheet.Cells[5, LastCol]].Borders.LineStyle := 1;

    // === Data ===
    Row := 6;
    ItemNo := 1;
    for Pair in MoldDict do
    begin
      Entry := Pair.Value;
      MoldKey := Entry.Key;
      WorkerDict := Entry.Value;

      Worksheet.Cells[Row, 1] := ItemNo;
      Worksheet.Cells[Row, 2] := MoldKey.MoldNo;
      Worksheet.Cells[Row, 3] := MoldKey.ControlNo;

      TotalHour := 0;
      for I := 0 to WorkerListMT.Count - 1 do
      begin
        WorkerName := WorkerListMT[I];
        Worksheet.Cells[Row, 4 + I].NumberFormat := '#,##0.00';
        if WorkerDict.TryGetValue(WorkerName, HourVal) then
          Worksheet.Cells[Row, 4 + I] := FormatFloat('#,##0.00', HourVal)
        else
          Worksheet.Cells[Row, 4 + I] := '0.00';
        TotalHour := TotalHour + HourVal;
      end;

      // Add left and right borders to the entire row
      for I := 1 to LastCol do
      begin
        Worksheet.Cells[Row, I].Borders[7].LineStyle := 1;  // xlEdgeLeft
        Worksheet.Cells[Row, I].Borders[10].LineStyle := 1; // xlEdgeRight
      end;

      Worksheet.Cells[Row, LastCol - 1] := FormatFloat('#,##0.00', TotalHour);
      Worksheet.Cells[Row, LastCol] := FormatFloat('#,##0.00', TotalHour * 390);

      Inc(Row);
      Inc(ItemNo);
    end;

    // === TOTAL ROW ===
    var LastDataRow, SignatureRow: Integer;
    var GrandTotalHour, ColSum: Double;
    var ColLetter: string;
    LastDataRow := Row - 1;
    GrandTotalHour := 0;

    Worksheet.Cells[Row, 1] := 'TOTAL';
    Worksheet.Range[Worksheet.Cells[Row, 1], Worksheet.Cells[Row, 3]].Merge;
    Worksheet.Range[Worksheet.Cells[Row, 1], Worksheet.Cells[Row, 3]].Font.Bold := True;

    // Assuming WorkerList contains the list of MT workers and data starts at column 4
    for I := 0 to WorkerListMT.Count - 1 do
    begin
      ColLetter := Chr(Ord('D') + I); // column D onwards
      ColSum := Worksheet.Application.WorksheetFunction.Sum(
        Worksheet.Range[ColLetter + '6', ColLetter + IntToStr(LastDataRow)]
      );
      Worksheet.Cells[Row, 4 + I] := FormatFloat('#,##0.00', ColSum);
      GrandTotalHour := GrandTotalHour + ColSum;
    end;

    Worksheet.Cells[Row, LastCol - 1] := FormatFloat('#,##0.00', GrandTotalHour);
    Worksheet.Cells[Row, LastCol] := FormatFloat('#,##0.00', GrandTotalHour * 390);
    Worksheet.Range[Worksheet.Cells[Row, 1], Worksheet.Cells[Row, LastCol]].Font.Bold := True;
    Worksheet.Range[Worksheet.Cells[Row, 1], Worksheet.Cells[Row, LastCol]].Borders.LineStyle := 1;

    // === Signature Block ===
    Inc(Row); Inc(Row);
    SignatureRow := Row;

    Worksheet.Rows[SignatureRow + 1].RowHeight := 116;

    var IssuedCol := LastCol - 2;
    var CheckedCol := LastCol - 1;
    var ApprovedCol := LastCol;

    Worksheet.Columns[IssuedCol].ColumnWidth := 19;

    Worksheet.Cells[SignatureRow, IssuedCol] := 'ISSUED BY';
    Worksheet.Cells[SignatureRow, CheckedCol] := 'CHECKED BY';
    Worksheet.Cells[SignatureRow, ApprovedCol] := 'APPROVED BY';

    Worksheet.Range[Worksheet.Cells[SignatureRow, IssuedCol], Worksheet.Cells[SignatureRow, ApprovedCol]].Font.Bold := True;
    Worksheet.Range[Worksheet.Cells[SignatureRow, IssuedCol], Worksheet.Cells[SignatureRow, ApprovedCol]].HorizontalAlignment := -4108;

    Worksheet.Range[Worksheet.Cells[SignatureRow, IssuedCol], Worksheet.Cells[SignatureRow, IssuedCol]].Borders.LineStyle := 1;
    Worksheet.Range[Worksheet.Cells[SignatureRow, CheckedCol], Worksheet.Cells[SignatureRow, CheckedCol]].Borders.LineStyle := 1;
    Worksheet.Range[Worksheet.Cells[SignatureRow, ApprovedCol], Worksheet.Cells[SignatureRow, ApprovedCol]].Borders.LineStyle := 1;

    // Bottom row borders for spacing
    Inc(SignatureRow);
    Worksheet.Range[Worksheet.Cells[SignatureRow, IssuedCol], Worksheet.Cells[SignatureRow, IssuedCol]].Borders.LineStyle := 1;
    Worksheet.Range[Worksheet.Cells[SignatureRow, CheckedCol], Worksheet.Cells[SignatureRow, CheckedCol]].Borders.LineStyle := 1;
    Worksheet.Range[Worksheet.Cells[SignatureRow, ApprovedCol], Worksheet.Cells[SignatureRow, ApprovedCol]].Borders.LineStyle := 1;

    // === Freeze pane ===
    Worksheet.Activate;
    Worksheet.Cells[6, 1].Select;
    Worksheet.Application.ActiveWindow.FreezePanes := True;

    var DataRange, SheetRange: OleVariant;
    begin
      // กำหนดขนาดพื้นที่ทั้งแผ่น (เช่น 1 ถึง 1000 แถว และ 1 ถึง 100 คอลัมน์)
      SheetRange := Worksheet.Range[Worksheet.Cells[1, 1], Worksheet.Cells[1000, 100]];
      SheetRange.Interior.Color := $C0C0C0; // สีเทาทั้งพื้นหลัง

      // กำหนดขอบเขตที่มีข้อมูลให้เป็นสีขาว
      DataRange := Worksheet.Range[Worksheet.Cells[1, 1], Worksheet.Cells[Row+1, LastCol]];
      DataRange.Interior.Color := clWhite;
    end;



    for Entry in MoldDict.Values do
      Entry.Value.Free;
    MoldDict.Free;
  finally
    MemTable.Free;
    WorkerTable.Free;
    WorkerListMT.Free;
  end;
end;


procedure TFormMain.CreateCleaningSheet(const Worksheet: OleVariant);
type
  TMoldKey = record
    MoldNo: string;
  end;

  TWorkerDict = TDictionary<string, Double>;
  TMoldDict = TDictionary<string, TPair<TMoldKey, TWorkerDict>>;

var
  SQLText, WorkerSQL, ResultMessage: string;
  MemTable, WorkerTable: TFDMemTable;
  DateBase: TDate;
  YearStr, MonthStr, StartDateStr, EndDateStr, ThaiMonthYear: string;
  MoldNo, WorkerName, KeyStr: string;
  HourVal, TotalHour: Double;
  MoldKey: TMoldKey;
  MoldDict: TMoldDict;
  WorkerDict: TWorkerDict;
  Entry: TPair<TMoldKey, TWorkerDict>;
  Pair: TPair<string, TPair<TMoldKey, TWorkerDict>>;
  WorkerList: TStringList;
  Row, I, ColIndex, LastCol, ItemNo: Integer;
    LastDataRow: Integer;
  ColSum,GrandTotalHour: Double;
  ColLetter: string;
  SignatureRow, IssuedCol, CheckedCol, ApprovedCol: Integer;

begin
  DateBase := DateTimePickerEnd.Date;
  YearStr := FormatDateTime('yyyy', DateBase);
  MonthStr := FormatDateTime('mm', DateBase);
  StartDateStr := FormatDateTime('yyyy/mm/dd', EncodeDate(StrToInt(YearStr), StrToInt(MonthStr), 1));
  EndDateStr := FormatDateTime('yyyy/mm/dd', EndOfTheMonth(DateBase));

  SQLText :=
    'SELECT sm.katacd AS "MoldNo", tm.tantonm AS "WorkerName", ' +
    'TRUNC(SUM(jd.jh) / 60.00, 2) AS "Hour"  ' +
    'FROM jisekidata jd ' +
    'INNER JOIN seizomst sm ON jd.seizono = sm.seizono ' +
    'INNER JOIN tantomst tm ON jd.tantocd = tm.tantocd ' +
    'WHERE tm.tantogrcd = ''MT'' and keikoteicd = ''CL''' +
    'AND jd.ymds <= TO_DATE(''' + EndDateStr + ''', ''YYYY/MM/DD'') ' +
    'AND jd.ymde >= TO_DATE(''' + StartDateStr + ''', ''YYYY/MM/DD'') ' +
    'AND sm.seizono IS NOT NULL ' +
    'GROUP BY sm.katacd, tm.tantonm ' +
    'ORDER BY "MoldNo", "WorkerName"';

  WorkerSQL := 'SELECT tantonm FROM tantomst WHERE tantogrcd = ''MT'' ORDER BY tantonm';

  MemTable := TFDMemTable.Create(nil);
  WorkerTable := TFDMemTable.Create(nil);
  WorkerList := TStringList.Create;
  WorkerList.Sorted := False;
  WorkerList.Duplicates := dupIgnore;

  try
    if FIsUseRESTAPIState = tssOff then
      DataModuleCIMT.FetchDataFromOracle(WorkerSQL, WorkerTable)
    else
    begin
      ResultMessage := DataModuleCIMT.FetchDataFromREST(WorkerSQL, WorkerTable);
      if ResultMessage <> 'Success' then Exit;
    end;

    WorkerTable.First;
    while not WorkerTable.Eof do
    begin
      WorkerList.Add(WorkerTable.FieldByName('tantonm').AsString);
      WorkerTable.Next;
    end;

    if FIsUseRESTAPIState = tssOff then
      DataModuleCIMT.FetchDataFromOracle(SQLText, MemTable)
    else
    begin
      ResultMessage := DataModuleCIMT.FetchDataFromREST(SQLText, MemTable);
      if ResultMessage <> 'Success' then Exit;
    end;

    if FIsUseRESTAPIState = tssOff then
      DataModuleCIMT.FetchDataFromOracle(SQLText, MemTable)
    else
    begin
      ResultMessage := DataModuleCIMT.FetchDataFromREST(SQLText, MemTable);
      if ResultMessage <> 'Success' then
      begin
        ShowMessage('Data fetch error: ' + ResultMessage);
        Exit;
      end;
    end;

    if MemTable.IsEmpty then
    begin
      ShowMessage('No data found for Cleaning.');
      Exit;
    end;

    // ✅ ตอนนี้ปลอดภัยที่จะ MemTable.First ได้
    MemTable.First;

    MoldDict := TMoldDict.Create;
    while not MemTable.Eof do
    begin
      MoldNo := MemTable.FieldByName('MoldNo').AsString;
      WorkerName := MemTable.FieldByName('WorkerName').AsString;
      HourVal := MemTable.FieldByName('Hour').AsFloat;

      KeyStr := MoldNo;
      if not MoldDict.TryGetValue(KeyStr, Entry) then
      begin
        MoldKey.MoldNo := MoldNo;
        WorkerDict := TWorkerDict.Create;
        MoldDict.Add(KeyStr, TPair<TMoldKey, TWorkerDict>.Create(MoldKey, WorkerDict));
      end
      else
        WorkerDict := Entry.Value;

      if WorkerDict.ContainsKey(WorkerName) then
        WorkerDict[WorkerName] := WorkerDict[WorkerName] + HourVal
      else
        WorkerDict.Add(WorkerName, HourVal);

      MemTable.Next;
    end;

    // === Headers ===

    InsertLogoToExcel(Worksheet);  // 👈 Add this line

    Worksheet.Cells[1, 1] := 'บริษัท นิฟโก้ (ไทยแลนด์) จำกัด';
    Worksheet.Range[Worksheet.Cells[1, 1], Worksheet.Cells[1, 2 + WorkerList.Count + 2]].Merge;
    Worksheet.Range['A1'].HorizontalAlignment := -4108;
    Worksheet.Range['A1'].Font.Size := 24;
    Worksheet.Range['A1'].Font.Bold := True;

    ThaiMonthYear := 'REPORT SUMMARY WORKER : ' + UpperCase(FormatDateTime('mmmm yyyy', DateBase));
    Worksheet.Cells[2, 1] := ThaiMonthYear;
    Worksheet.Range['A2'].HorizontalAlignment := -4131;
    Worksheet.Range['A2'].Font.Size := 16;
    Worksheet.Range['A2'].Font.Bold := True;

    Worksheet.Cells[3, 1] := 'CLEANING';
    Worksheet.Range['A3'].Font.Size := 16;
    Worksheet.Range['A3'].Font.Bold := True;


    Worksheet.Cells[4, 1] := 'NO';
    Worksheet.Range['A4:A5'].Merge;
    Worksheet.Range['A4'].HorizontalAlignment := -4108;
    Worksheet.Range['A4'].VerticalAlignment := -4108;
    Worksheet.Range['A4'].Font.Bold := True;

    Worksheet.Cells[4, 2] := 'MOLD NO.';
    Worksheet.Range['B4:B5'].Merge;
    Worksheet.Range['B4'].HorizontalAlignment := -4108;
    Worksheet.Range['B4'].VerticalAlignment := -4108;
    Worksheet.Range['B4'].Font.Bold := True;

    ColIndex := 3;
    if WorkerList.Count > 0 then
    begin
      for I := 0 to WorkerList.Count - 1 do
      begin
      Worksheet.Cells[4, ColIndex + I] := WorkerList[I];
      Worksheet.Cells[5, ColIndex + I] := 'TIME';
      Worksheet.Range[Worksheet.Cells[4, ColIndex + I], Worksheet.Cells[4, ColIndex + I]].HorizontalAlignment := -4108;
      Worksheet.Range[Worksheet.Cells[5, ColIndex + I], Worksheet.Cells[5, ColIndex + I]].HorizontalAlignment := -4108;
      Worksheet.Columns[ColIndex + I].ColumnWidth := 15;

      end;

      Inc(ColIndex, WorkerList.Count);
    end;

    Worksheet.Cells[4, ColIndex] := 'GRAND TOTAL';
    Worksheet.Range[Worksheet.Cells[4, ColIndex], Worksheet.Cells[5, ColIndex]].Merge;
    Worksheet.Range[Worksheet.Cells[4, ColIndex], Worksheet.Cells[5, ColIndex]].Merge;
    Worksheet.Range[Worksheet.Cells[4, ColIndex], Worksheet.Cells[5, ColIndex]].HorizontalAlignment := -4108;
    Worksheet.Range[Worksheet.Cells[4, ColIndex], Worksheet.Cells[5, ColIndex]].VerticalAlignment := -4108;
    Worksheet.Range[Worksheet.Cells[4, ColIndex], Worksheet.Cells[5, ColIndex]].Font.Bold := True;
    Worksheet.Range[Worksheet.Cells[4, ColIndex], Worksheet.Cells[5, ColIndex]].Merge;
    Worksheet.Columns[ColIndex].ColumnWidth := 20;

    Inc(ColIndex);
    Worksheet.Cells[4, ColIndex] := 'AMOUNT COST';
    Worksheet.Range[Worksheet.Cells[4, ColIndex], Worksheet.Cells[5, ColIndex]].Merge;
    Worksheet.Range[Worksheet.Cells[4, ColIndex], Worksheet.Cells[5, ColIndex]].Merge;
    Worksheet.Range[Worksheet.Cells[4, ColIndex], Worksheet.Cells[5, ColIndex]].HorizontalAlignment := -4108;
    Worksheet.Range[Worksheet.Cells[4, ColIndex], Worksheet.Cells[5, ColIndex]].VerticalAlignment := -4108;
    Worksheet.Range[Worksheet.Cells[4, ColIndex], Worksheet.Cells[5, ColIndex]].Font.Bold := True;
    Worksheet.Range[Worksheet.Cells[4, ColIndex], Worksheet.Cells[5, ColIndex]].Merge;
    Worksheet.Columns[ColIndex].ColumnWidth := 20;

    LastCol := ColIndex;

    Worksheet.Rows[1].RowHeight := 50;
    Worksheet.Rows[2].RowHeight := 30;
    Worksheet.Rows[3].RowHeight := 20;
    Worksheet.Rows[4].RowHeight := 22;
    Worksheet.Rows[5].RowHeight := 22;
    Worksheet.Columns['A'].ColumnWidth := 10;
    Worksheet.Columns['B'].ColumnWidth := 14;
    Worksheet.Range[Worksheet.Cells[4, 1], Worksheet.Cells[5, LastCol]].Borders.LineStyle := 1;

    // === Data ===
    Row := 6;
    ItemNo := 1;
    for Pair in MoldDict do
    begin
      Entry := Pair.Value;
      MoldKey := Entry.Key;
      WorkerDict := Entry.Value;

      Worksheet.Cells[Row, 1] := ItemNo;
      Worksheet.Cells[Row, 2] := MoldKey.MoldNo;

      TotalHour := 0.00;
      for I := 0 to WorkerList.Count - 1 do
      begin
        WorkerName := WorkerList[I];
        Worksheet.Cells[Row, 3 + I].NumberFormat := '#,##0.00';
        if WorkerDict.TryGetValue(WorkerName, HourVal) then
          Worksheet.Cells[Row, 3 + I] := FormatFloat('#,##0.00', HourVal)
        else
          Worksheet.Cells[Row, 3 + I] := '0.00';
        TotalHour := TotalHour + HourVal;
      end;

        // Add left and right borders to the entire row
      for I := 1 to LastCol do
      begin
        Worksheet.Cells[Row, I].Borders[7].LineStyle := 1;  // xlEdgeLeft
        Worksheet.Cells[Row, I].Borders[10].LineStyle := 1; // xlEdgeRight
      end;

      Worksheet.Cells[Row, LastCol - 1] := FormatFloat('#,##0.00', TotalHour);
      Worksheet.Cells[Row, LastCol] := FormatFloat('#,##0.00', TotalHour * 390);
      Inc(Row);
      Inc(ItemNo);
    end;


  // === TOTAL ROW ===
  LastDataRow := Row - 1;
  GrandTotalHour := 0;

  Worksheet.Cells[Row, 1] := 'TOTAL';
  Worksheet.Range[Worksheet.Cells[Row, 1], Worksheet.Cells[Row, 2]].Merge;
  Worksheet.Range[Worksheet.Cells[Row, 1], Worksheet.Cells[Row, 2]].Font.Bold := True;

  for I := 0 to WorkerList.Count - 1 do
  begin
    ColLetter := Chr(Ord('C') + I); // Assuming worker data starts at column C (3)
    ColSum := Worksheet.Application.WorksheetFunction.Sum(
      Worksheet.Range[ColLetter + '6', ColLetter + IntToStr(LastDataRow)]
    );
    Worksheet.Cells[Row, 3 + I] := FormatFloat('#,##0.00', ColSum);
    GrandTotalHour := GrandTotalHour + ColSum;
  end;

  Worksheet.Cells[Row, LastCol - 1] := FormatFloat('#,##0.00', GrandTotalHour);
  Worksheet.Cells[Row, LastCol] := FormatFloat('#,##0.00', GrandTotalHour * 390);

  Worksheet.Range[Worksheet.Cells[Row, 1], Worksheet.Cells[Row, LastCol]].Font.Bold := True;
  Worksheet.Range[Worksheet.Cells[Row, 1], Worksheet.Cells[Row, LastCol]].Borders.LineStyle := 1;

  // === SIGNATURE BLOCK ===
  Inc(Row);
  Inc(Row);
  SignatureRow := Row;
  Worksheet.Rows[SignatureRow + 1].RowHeight := 116;

  IssuedCol := LastCol - 2;
  CheckedCol := LastCol - 1;
  ApprovedCol := LastCol;

  Worksheet.Columns[IssuedCol].ColumnWidth := 19;

  Worksheet.Cells[SignatureRow, IssuedCol] := 'ISSUED BY';
  Worksheet.Cells[SignatureRow, CheckedCol] := 'CHECKED BY';
  Worksheet.Cells[SignatureRow, ApprovedCol] := 'APPROVED BY';

  Worksheet.Range[Worksheet.Cells[SignatureRow, IssuedCol], Worksheet.Cells[SignatureRow, ApprovedCol]].HorizontalAlignment := -4108;
  Worksheet.Range[Worksheet.Cells[SignatureRow, IssuedCol], Worksheet.Cells[SignatureRow, ApprovedCol]].Font.Bold := True;

  Worksheet.Range[Worksheet.Cells[SignatureRow, IssuedCol], Worksheet.Cells[SignatureRow, IssuedCol]].Borders.LineStyle := 1;
  Worksheet.Range[Worksheet.Cells[SignatureRow, CheckedCol], Worksheet.Cells[SignatureRow, CheckedCol]].Borders.LineStyle := 1;
  Worksheet.Range[Worksheet.Cells[SignatureRow, ApprovedCol], Worksheet.Cells[SignatureRow, ApprovedCol]].Borders.LineStyle := 1;

  Inc(SignatureRow);
  Worksheet.Range[Worksheet.Cells[SignatureRow, IssuedCol], Worksheet.Cells[SignatureRow, IssuedCol]].Borders.LineStyle := 1;
  Worksheet.Range[Worksheet.Cells[SignatureRow, CheckedCol], Worksheet.Cells[SignatureRow, CheckedCol]].Borders.LineStyle := 1;
  Worksheet.Range[Worksheet.Cells[SignatureRow, ApprovedCol], Worksheet.Cells[SignatureRow, ApprovedCol]].Borders.LineStyle := 1;

    Worksheet.Activate;
    Worksheet.Cells[6, 1].Select;
    Worksheet.Application.ActiveWindow.FreezePanes := True;

    var
      DataRange, SheetRange: OleVariant;
    begin
      // กำหนดขนาดพื้นที่ทั้งแผ่น (เช่น 1 ถึง 1000 แถว และ 1 ถึง 100 คอลัมน์)
      SheetRange := Worksheet.Range[Worksheet.Cells[1, 1], Worksheet.Cells[1000, 100]];
      SheetRange.Interior.Color := $C0C0C0; // สีเทาทั้งพื้นหลัง

      // กำหนดขอบเขตที่มีข้อมูลให้เป็นสีขาว
      DataRange := Worksheet.Range[Worksheet.Cells[1, 1], Worksheet.Cells[Row+1, LastCol]];
      DataRange.Interior.Color := clWhite;
    end;


    for Entry in MoldDict.Values do
      Entry.Value.Free;
    MoldDict.Free;
  finally
    MemTable.Free;
    WorkerTable.Free;
    WorkerList.Free;
  end;
end;




procedure TFormMain.btnCloseClick(Sender: TObject);
begin
  WriteIniFile;
  Close;
end;

procedure TFormMain.WriteIniFile;
var
  IniFile: TIniFile;
  IniFileName, IniFileDir: string;
begin
  IniFileDir := ExtractFilePath(Application.ExeName) + 'GRD\';
  IniFileName := IniFileDir + ChangeFileExt
    (ExtractFileName(Application.ExeName), '.ini');

  // Ensure the directory exists
  if not DirectoryExists(IniFileDir) then
    ForceDirectories(IniFileDir);

  IniFile := TIniFile.Create(IniFileName);
  try

    IniFile.WriteInteger('Setting', 'IsUseRESTAPI',
    Integer(FIsUseRESTAPIState));
    IniFile.WriteInteger('Setting', 'DebugSQL', Integer(FDebugSQL));
    IniFile.WriteString('Setting', 'User', FUser);
    IniFile.WriteString('Setting', 'Pass', FPass);
  finally
    IniFile.Free;
  end;
end;

procedure TFormMain.ReadIniFile;
var
  IniFile: TIniFile;
  IniFileName: string;
begin
  IniFileName := ExtractFilePath(Application.ExeName) + 'GRD\' +
    ChangeFileExt(ExtractFileName(Application.ExeName), '.ini');
  IniFile := TIniFile.Create(IniFileName);
  try

    FIsUseRESTAPIState := TToggleSwitchState(IniFile.ReadInteger('Setting',
      'IsUseRESTAPI', Integer(tssOff)));
    FDebugSQL := TToggleSwitchState(IniFile.ReadInteger('Setting', 'DebugSQL',
      Integer(tssOff)));
    FUser := IniFile.ReadString('Setting', 'User', 'admin');
    FPass := IniFile.ReadString('Setting', 'Pass', 'admin');
  finally
    IniFile.Free;
  end;
end;

procedure TFormMain.FormCreate(Sender: TObject);
begin
  DateTimePickerEnd.Format := 'MM/yyyy';
  DateTimePickerEnd.Date := date; // Default to today  date
  InitialProgram;
end;

function TFormMain.HasMonthlyData: Boolean;
var
  SQLCheck, ResultMessage: string;
  MemTable: TFDMemTable;
  YearStr, MonthStr, StartDateStr, EndDateStr: string;
  DateBase: TDate;
begin
  Result := False;
  DateBase := DateTimePickerEnd.Date;
  YearStr := FormatDateTime('yyyy', DateBase);
  MonthStr := FormatDateTime('mm', DateBase);

  StartDateStr := FormatDateTime('yyyy/mm/dd', EncodeDate(StrToInt(YearStr), StrToInt(MonthStr), 1));
  EndDateStr := FormatDateTime('yyyy/mm/dd', EndOfTheMonth(DateBase));

  SQLCheck :=
    'SELECT COUNT(*) AS CNT FROM jisekidata jd ' +
    'INNER JOIN seizomst sm ON jd.seizono = sm.seizono ' +
    'WHERE jd.ymds <= TO_DATE(''' + EndDateStr + ''', ''YYYY/MM/DD'') ' +
    'AND jd.ymde >= TO_DATE(''' + StartDateStr + ''', ''YYYY/MM/DD'') ' +
    'AND sm.seizono IS NOT NULL';

  MemTable := TFDMemTable.Create(nil);
  try
    if FIsUseRESTAPIState = tssOff then
      DataModuleCIMT.FetchDataFromOracle(SQLCheck, MemTable)
    else
    begin
      ResultMessage := DataModuleCIMT.FetchDataFromREST(SQLCheck, MemTable);
      if ResultMessage <> 'Success' then
      begin
        ShowMessage('Data fetch error: ' + ResultMessage);
        Exit;
      end;
    end;

    if not MemTable.Active then
      MemTable.Open;
    var a :integer;

    if not MemTable.IsEmpty then
    begin
      if not MemTable.FieldByName('CNT').IsNull then
        Result := MemTable.FieldByName('CNT').AsInteger > 0;
    end;
    a := MemTable.FieldByName('CNT').AsInteger;
        if a >= 1  then
       begin
          Result := true ;
       end else
       begin
          Result := false ;
       end;


  finally
    MemTable.Free;
  end;
end;

{$REGION 'Template'}

procedure TFormMain.InitialProgram;
var
  ExeFileName, FileVersion: string;

  function GetFileVersion(const FileName: TFileName): string;
  var
    Size, Handle: DWORD;
    Buffer: array of Byte;
    FileInfo: PVSFixedFileInfo;
    FileInfoSize: UINT;
  begin
    Size := GetFileVersionInfoSize(PChar(FileName), Handle);
    if Size = 0 then
      RaiseLastOSError;

    SetLength(Buffer, Size);
    if not GetFileVersionInfo(PChar(FileName), Handle, Size, Buffer) then
      RaiseLastOSError;

    if not VerQueryValue(Buffer, '\', Pointer(FileInfo), FileInfoSize) then
      RaiseLastOSError;

    Result := Format('%d.%d.%d.%d', [HiWord(FileInfo.dwFileVersionMS),
      LoWord(FileInfo.dwFileVersionMS), HiWord(FileInfo.dwFileVersionLS),
      LoWord(FileInfo.dwFileVersionLS)]);
  end;

begin
  ExeFileName := ExtractFileName(Application.ExeName);
  FileVersion := GetFileVersion(Application.ExeName);

  stbBase.Panels.Add;
  stbBase.Panels[0].Width := 210;
  stbBase.Panels.Add;
  stbBase.Panels[1].Width := 160;
  stbBase.Panels.Add;
  stbBase.Panels[2].Width := 600;

  stbBase.Panels[0].Text := ' ' + ExeFileName;
  stbBase.Panels[1].Text := FileVersion;

  DataModuleCIMT := TDataModuleCIMT.Create(Self);
  // Create the data module instance
  ReadIniFile;
  InitConnection;
end;

procedure TFormMain.InitConnection;
var
  ConnectionStatus: string;
begin
  // Initialize connection before any data operations
  try
    if not DataModuleCIMT.UniConnection1.Connected then
    begin
      ConnectionStatus := DataModuleCIMT.InitializeConnection
        (FIsUseRESTAPIState = tssOn);
      stbBase.Panels[2].Text := ConnectionStatus;
    end;
  except
    on E: Exception do
      ShowMessage('Error establishing Oracle connection: ' + E.Message);
  end;
end;

procedure TFormMain.DebugSQLText(const SQLText: string);
var
  DialogForm: TForm;
  Memo: TMemo;
begin
  DialogForm := TForm.Create(nil);
  try
    DialogForm.Width := 800;
    DialogForm.Height := 600;
    DialogForm.Position := poScreenCenter;
    DialogForm.Caption := 'SQL Debug';

    Memo := TMemo.Create(DialogForm);
    Memo.Parent := DialogForm;
    Memo.Align := alClient;
    Memo.Lines.Text := SQLText;
    Memo.ReadOnly := True;
    Memo.ScrollBars := ssVertical;
    Memo.Font.Name := 'Courier New';
    Memo.Font.Size := 10;

    Memo.SelectAll;

    DialogForm.ShowModal;
  finally
    DialogForm.Free;
  end;
end;

{$ENDREGION}

end.
