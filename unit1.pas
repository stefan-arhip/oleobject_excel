unit Unit1;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, Forms, Controls, Graphics, Dialogs, StdCtrls, Grids, Comobj;

type

  { TForm1 }

  TForm1 = class(TForm)
    Button1: TButton;
    Button2: TButton;
    Edit1: TEdit;
    StringGrid1: TStringGrid;
    procedure Button1Click(Sender: TObject);
    procedure Button2Click(Sender: TObject);
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
  path: variant;
begin
  XLApp := CreateOleObject('Excel.Application');
  try
    XLApp.Visible := False;
    XLApp.DisplayAlerts := False;
    path := edit1.Text;
    XLApp.Workbooks.Open(Path);
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
  XLApp: olevariant;
  x, y: byte;
  path: variant;
begin
  XLApp := CreateOleObject('Excel.Application');
  try
    XLApp.Visible := False;
    XLApp.DisplayAlerts := False;
    path := edit1.Text;
    XLApp.Workbooks.Open(Path);
    for x := 1 to StringGrid1.ColCount - 1 do
    begin
      for y := 1 to stringgrid1.RowCount - 1 do
      begin
        XLApp.Cells[y, x].Value := StringGrid1.Cells[x, y];
      end;
    end;
  finally
    XLApp.ActiveWorkBook.Save;
    XLApp.Quit;
    XLAPP := Unassigned;
  end;
end;

end.
