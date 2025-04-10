unit DetailForm;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.Grids;

type
  TFormDetail = class(TForm)
    stgDetail: TStringGrid;
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  FormDetail: TFormDetail;

implementation

{$R *.dfm}

end.
