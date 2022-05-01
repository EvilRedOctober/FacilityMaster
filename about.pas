unit about;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, Forms, Controls, Graphics, Dialogs, StdCtrls, ExtCtrls,
  Grids, ValEdit;

type

  { TAboutForm }

  TAboutForm = class(TForm)
    Button2: TButton;
    Image1: TImage;
    Memo1: TMemo;
    procedure Button2Click(Sender: TObject);

  private

  public

  end;

var
  AboutForm: TAboutForm;

implementation

{$R *.lfm}

{ TAboutForm }

procedure TAboutForm.Button2Click(Sender: TObject);
begin
  Aboutform.Visible:=false;
end;








end.

