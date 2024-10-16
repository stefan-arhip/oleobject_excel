unit Unit1;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, Forms, Controls, Graphics, Dialogs, StdCtrls, Grids,
  EditBtn, Comobj;

type

  { TForm1 }

  TForm1 = class(TForm)
    Button1: TButton;
    Button2: TButton;
    FileNameEdit1: TFileNameEdit;
    StringGrid1: TStringGrid;
    procedure Button1Click(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    procedure FormCreate(Sender: TObject);
  private

  public

  end;

var
  Form1: TForm1;

implementation

{$R *.lfm}

{ TForm1 }

procedure TForm1.Button1Click(Sender: TObject);
var
  XLApp: olevariant;
  x, y: byte;
  FilePath: variant;
begin
  XLApp := CreateOleObject('Excel.Application');
  try
    XLApp.Visible := False;
    XLApp.DisplayAlerts := False;
    FilePath := FileNameEdit1.FileName;
    XLApp.Workbooks.Open(FilePath);
    for x := 1 to StringGrid1.ColCount - 1 do
    begin
      for y := 1 to stringgrid1.RowCount - 1 do
      begin
        StringGrid1.Cells[x, y] := XLApp.Cells[y, x].Value;
      end;
    end;
  finally
    XLApp.Quit;
    XLAPP := Unassigned;
  end;
end;

procedure TForm1.Button2Click(Sender: TObject);
var
  XLApp, Workbook: olevariant;
  x, y: byte;
  FilePath: variant;
begin
  XLApp := CreateOleObject('Excel.Application');
  try
    XLApp.Visible := False;
    XLApp.DisplayAlerts := False;
    FilePath := FileNameEdit1.FileName;
    if FileExists(FilePath) then
      Workbook := XLApp.Workbooks.Open(FilePath)
    else
      Workbook := XLApp.Workbooks.Add;

    for x := 1 to StringGrid1.ColCount - 1 do
    begin
      for y := 1 to StringGrid1.RowCount - 1 do
      begin
        XLApp.Cells[y, x].Value := StringGrid1.Cells[x, y];
      end;
    end;
  finally
    if FileExists(FilePath) then
      Workbook.Save
    else
      Workbook.SaveAs(FilePath);
    XLApp.Quit;
    XLAPP := Unassigned;
  end;
end;

procedure TForm1.FormCreate(Sender: TObject);
begin
  FileNameEdit1.FileName := IncludeTrailingPathDelimiter(ExtractFileDir(ParamStr(0))) +
    'Book1.xlsx';
end;

end.
