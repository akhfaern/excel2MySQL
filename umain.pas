unit umain;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, ExtCtrls, StdCtrls, ComObj, DB, ZAbstractRODataset,
  ZAbstractDataset, ZDataset, ZConnection;

type
  TForm1 = class(TForm)
    Panel1: TPanel;
    GroupBox1: TGroupBox;
    edtExcelFile: TEdit;
    btnSelectExcelFile: TButton;
    Label1: TLabel;
    edtHeaderRowNumber: TEdit;
    GroupBox2: TGroupBox;
    Label2: TLabel;
    Label3: TLabel;
    Label4: TLabel;
    Label5: TLabel;
    edtMySQLHost: TEdit;
    edtMySQLPort: TEdit;
    edtMySQLUsername: TEdit;
    edtMySQLPassword: TEdit;
    Label6: TLabel;
    edtMySQLDatabase: TEdit;
    Label7: TLabel;
    edtMySQLTableName: TEdit;
    btnReadHeaders: TButton;
    btnReadTableHeaders: TButton;
    Panel2: TPanel;
    GroupBox3: TGroupBox;
    GroupBox4: TGroupBox;
    GroupBox5: TGroupBox;
    btnAddMatching: TButton;
    btnRemoveMatching: TButton;
    Panel3: TPanel;
    btnStartTransfer: TButton;
    Panel4: TPanel;
    memLogs: TMemo;
    lbExcelHeaders: TListBox;
    lbMySQLFields: TListBox;
    lbMatchings: TListBox;
    odSelectExcelFile: TOpenDialog;
    MySQLConnection: TZConnection;
    MySQLQuery: TZQuery;
    cbUnique: TCheckBox;
    GroupBox6: TGroupBox;
    Label8: TLabel;
    Label9: TLabel;
    Label10: TLabel;
    edtRelationMasterTableName: TEdit;
    edtRelationMasterFieldName: TEdit;
    edtRelationSlaveFieldName: TEdit;
    MySQLQuery2: TZQuery;
    Label12: TLabel;
    edtRelationMasterLookupFieldName: TEdit;
    procedure btnSelectExcelFileClick(Sender: TObject);
    procedure btnReadHeadersClick(Sender: TObject);
    procedure btnReadTableHeadersClick(Sender: TObject);
    procedure btnAddMatchingClick(Sender: TObject);
    procedure btnRemoveMatchingClick(Sender: TObject);
    procedure btnStartTransferClick(Sender: TObject);
  private
    { Private declarations }
  public
    procedure addLog(FLogMessage: string);
    function getFieldName(ParamStr: string): string;
    function getColIndex(ParamStr: string): integer;
  end;

var
  Form1: TForm1;

implementation

{$R *.dfm}

procedure TForm1.addLog(FLogMessage: string);
begin
  memLogs.Lines.Add(FLogMessage);
end;

procedure TForm1.btnSelectExcelFileClick(Sender: TObject);
begin
  if odSelectExcelFile.Execute then
    edtExcelFile.Text := odSelectExcelFile.FileName;
end;

procedure TForm1.btnReadHeadersClick(Sender: TObject);
var
  ExcelApp, Sheet,
  Range: OleVariant;
  i, ColCount,
  headerRowNumber: integer;
begin
  lbExcelHeaders.Items.Clear;
  if FileExists(edtExcelFile.Text) then
  begin
    try
      headerRowNumber := StrToInt(edtHeaderRowNumber.Text);
    except
      addLog('[ERROR] Please enter a number for Excel Header Row Number');
    end;
    addLog('[INFO] Initializing Excel...');
    try
      try
        ExcelApp := CreateOleObject('Excel.Application');
      except
        addLog('[ERROR] Excel failed!');
        Exit;
      end;
    finally
      addLog('[INFO] Reading Excel file...');
      ExcelApp.WorkBooks.Open(edtExcelFile.Text);
      Sheet := ExcelApp.WorkBooks[1].WorkSheets[1];
      Range := Sheet.UsedRange;
      ColCount := Range.Columns.Count;
      for i := 1 to ColCount do
        lbExcelHeaders.Items.Add(IntToStr(i) + ':' + string(Sheet.Cells[headerRowNumber, i]));
      if not VarIsEmpty(ExcelApp) then
      begin
        ExcelApp.DisplayAlerts := False;
        ExcelApp.Quit;
        ExcelApp := Unassigned;
      end;
      addLog('[INFO] Reading Excel Headers has finished!');
    end;
  end else
    addLog('[ERROR] Couldn''t locate selected Excel file!');
end;

procedure TForm1.btnReadTableHeadersClick(Sender: TObject);
var
  i: integer;
begin
  lbMySQLFields.Items.Clear;
  with MySQLConnection do
  begin
    if Connected then Connected := False;
    HostName := edtMySQLHost.Text;
    User := edtMySQLUsername.Text;
    Password := edtMySQLPassword.Text;
    Port := StrToInt(edtMySQLPort.Text);
    Database := edtMySQLDatabase.Text;
    try
      addLog('[INFO] Trying to connect to MySQL database...');
      Connect;
    finally
      if Connected then
      begin
        addLog('[INFO] Connected to MySQL database.');
        with MySQLQuery do
        begin
          if Active then Active := False;
          SQL.Text := 'SHOW FIELDS FROM ' + edtMySQLTableName.Text;
          try
            Active := True;
          finally
            First;
            for i := 0 to RecordCount - 1 do
            begin
              lbMySQLFields.Items.Add(Fields[0].AsString);
              Next;
            end;
            addLog('[INFO] MySQL table fields recieved.');
          end;
        end;
      end else
        addLog('[ERROR] Couldn''t connect to MySQL database.');
    end;
  end;
end;

procedure TForm1.btnAddMatchingClick(Sender: TObject);
var
  index1, index2: integer;
begin
  index1 := lbExcelHeaders.ItemIndex;
  index2 := lbMySQLFields.ItemIndex;
  if (index1 >= 0) and (index2 >= 0) then
    lbMatchings.Items.Add(lbExcelHeaders.Items[index1] + '=>' + lbMySQLFields.Items[index2]);
end;

procedure TForm1.btnRemoveMatchingClick(Sender: TObject);
begin
  if lbMatchings.ItemIndex >= 0 then
    lbMatchings.Items.Delete(lbMatchings.ItemIndex);
end;

procedure TForm1.btnStartTransferClick(Sender: TObject);
var
  ExcelApp, Sheet,
  Range: OleVariant;
  i, RowCount, j: integer;
  SQLString,
  params,
  SelectQueryText,
  FFieldName: string;
  addRecord: Boolean;
begin
  SQLString := 'Insert into ' + edtMySQLTableName.Text + ' (';
  SelectQueryText := 'SELECT * FROM ' + edtMySQLTableName.Text + ' WHERE ';
  params := '';
  for i := 0 to lbMatchings.Count - 1 do
  begin
    FFieldName := getFieldName(lbMatchings.Items[i]);
    SQLString := SQLString + FFieldName + ', ';
    params := params + ':' + FFieldName + ', ';
    SelectQueryText := SelectQueryText + FFieldName + '= :' + FFieldName + ' AND ';
  end;
  SQLString := Copy(SQLString, 1, Length(SQLString) - 2) + ') VALUES (' +
    Copy(params, 1, Length(params) - 2) + ');';
  SelectQueryText := Copy(SelectQueryText, 1, Length(SelectQueryText) - 5);
  if FileExists(edtExcelFile.Text) then
  begin
    addLog('[INFO] Initializing Excel...');
    try
      try
        ExcelApp := CreateOleObject('Excel.Application');
      except
        addLog('[ERROR] Excel failed!');
        Exit;
      end;
    finally
      ExcelApp.WorkBooks.Open(edtExcelFile.Text);
      Sheet := ExcelApp.WorkBooks[1].WorkSheets[1];
      Range := Sheet.UsedRange;
      RowCount := Range.Rows.Count;
      for i := (StrToInt(edtHeaderRowNumber.Text) + 1) to RowCount do
      begin
        Application.ProcessMessages;
        addRecord := True;
        if cbUnique.Checked then
        begin
          with MySQLQuery do
          begin
            SQL.Text := SelectQueryText;
            for j := 0 to lbMatchings.Items.Count - 1 do
            begin
              FFieldName := getFieldName(lbMatchings.Items[j]);
              addLog(FFieldName);
              if FFieldName = edtRelationSlaveFieldName.Text then
              begin
                MySQLQuery2.SQL.Text := 'SELECT ' + edtRelationMasterFieldName.Text +
                  ' FROM ' + edtRelationMasterTableName.Text + ' WHERE ' +
                  edtRelationMasterLookupFieldName.Text + ' = :' + edtRelationMasterLookupFieldName.Text;
                MySQLQuery2.ParamByName(edtRelationMasterLookupFieldName.Text).AsString := Sheet.Cells[i, getColIndex(lbMatchings.Items[j])];
                try
                  MySQLQuery2.Active := True;
                finally
                  if MySQLQuery2.RecordCount > 0 then
                    ParamByName(FFieldName).AsString := MySQLQuery2.FieldByName(edtRelationMasterFieldName.Text).AsString
                  else
                    ParamByName(FFieldName).AsString := '0';
                  MySQLQuery2.Active := False;
                end;
              end else
                ParamByName(FFieldName).AsString := Sheet.Cells[i, getColIndex(lbMatchings.Items[j])];
            end;
            try
              Active := True;
            finally
              if RecordCount > 0 then
                addRecord := False;
            end;
          end;
        end;
        if addRecord then
        begin
          with MySQLQuery do
          begin
            SQL.Text := SQLString;
            if Active then
              Active := False;
            for j := 0 to lbMatchings.Items.Count - 1 do
            begin
              FFieldName := getFieldName(lbMatchings.Items[j]);
              if FFieldName = edtRelationSlaveFieldName.Text then
              begin
                MySQLQuery2.SQL.Text := 'SELECT ' + edtRelationMasterFieldName.Text +
                  ' FROM ' + edtRelationMasterTableName.Text + ' WHERE ' +
                  edtRelationMasterLookupFieldName.Text + ' = :' + edtRelationMasterLookupFieldName.Text;
                MySQLQuery2.ParamByName(edtRelationMasterLookupFieldName.Text).AsString := Sheet.Cells[i, getColIndex(lbMatchings.Items[j])];
                try
                  MySQLQuery2.Active := True;
                finally
                  if MySQLQuery2.RecordCount > 0 then
                    ParamByName(FFieldName).AsString := MySQLQuery2.FieldByName(edtRelationMasterFieldName.Text).AsString
                  else
                    ParamByName(FFieldName).AsString := '0';
                  MySQLQuery2.Active := False;
                end;
              end else
                ParamByName(FFieldName).AsString := Sheet.Cells[i, getColIndex(lbMatchings.Items[j])];
            end;
            ExecSQL;
          end;
        end;
      end;
      if not VarIsEmpty(ExcelApp) then
      begin
        ExcelApp.DisplayAlerts := False;
        ExcelApp.Quit;
        ExcelApp := Unassigned;
      end;
      addLog('[INFO] Transfer completed!');
    end;
  end else
    addLog('[ERROR] Couldn''t locate selected Excel file!');
end;

function TForm1.getFieldName(ParamStr: string): string;
begin
  Result := Trim(Copy(ParamStr, Pos('=>', ParamStr) + 2, Length(ParamStr)));
end;

function TForm1.getColIndex(ParamStr: string): integer;
begin
  Result := StrToInt(Copy(ParamStr, 1, Pos(':', ParamStr) - 1));
end;

end.
