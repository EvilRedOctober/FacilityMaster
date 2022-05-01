unit Test;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, Forms, Controls, Graphics, Dialogs, StdCtrls, ActnList, main, intrfcd, math;

type

  { TTestForm }

  TTestForm = class(TForm)
    Button1: TButton;
    Button2: TButton;
    ComboBox1: TComboBox;
    Label1: TLabel;
    procedure Button1Click(Sender: TObject);
    procedure Button2Click(Sender: TObject);
  private

  public

  end;

var
  TestForm: TTestForm;

implementation

{$R *.lfm}

procedure HowLongOpros(Nom:byte);
var Elapsed: cardinal;
  j:integer;
  X:array [1..3] of real;
begin
  Elapsed:=GetTickCount64;
  opros(Nom,x[3],wspom);
  opros(Nom,x[1],wspom);
  for j:=1 to 1000 do
  begin
    opros(Nom,x[2],wspom);
    if x[2]<x[1] then x[1]:=x[2];
    if x[2]>x[3] then x[3]:=x[2];
    sleep(random(5));
  end;
  Elapsed:=GetTickCount64-Elapsed;
  ShowMessage('Общее время выполнения: '+IntToStr(Elapsed)+' мс.'#13#10'Среднее время выполнения: '+FloatToStrF(Elapsed/1,ffFixed,10,3)+' мкс.'+
  #13#10'Канал'+IntToStr(nom)+' Макс:'+FloatToStrF(x[3],ffFixed,10,2)+' Мин:'+FloatToStrF(x[1],ffFixed,10,2));
end;

procedure HowLongMean;
var Elapsed: cardinal;
  j:integer;
  M:real;
  X:array [1..4] of real;
begin
  Elapsed:=GetTickCount64;
  X[1]:=10;X[2]:=450;X[3]:=1337;X[4]:=9999;
  for j:=1 to 1000000 do M:=Mean(X);
  Elapsed:=GetTickCount64-Elapsed;
  ShowMessage('Общее время выполнения: '+IntToStr(Elapsed)+' мс.'#13#10'Среднее время выполнения: '+FloatToStrF(Elapsed/1000,ffFixed,10,3)+' мкс.');
end;

procedure HowLongDisp;
var Elapsed: cardinal;
  j:integer;
  D:real;
  X:array [1..4] of real;
begin
  Elapsed:=GetTickCount64;
  X[1]:=10;X[2]:=450;X[3]:=1337;X[4]:=9999;
  for j:=1 to 1000000 do D:=StdDev(X);
  Elapsed:=GetTickCount64-Elapsed;
  ShowMessage('Общее время выполнения: '+IntToStr(Elapsed)+' мс.'#13#10'Среднее время выполнения: '+FloatToStrF(Elapsed/1000,ffFixed,10,3)+' мкс.');
end;

procedure HowLongPO;
var Elapsed: cardinal;
  j:integer;
  X:real;
begin
  Elapsed:=GetTickCount64;
  X:=10;
  for j:=1 to 1000000 do PO(X,30,7);
  Elapsed:=GetTickCount64-Elapsed;
  ShowMessage('Общее время выполнения: '+IntToStr(Elapsed)+' мс.'#13#10'Среднее время выполнения: '+FloatToStrF(Elapsed/1000,ffFixed,10,3)+' мкс.');
end;

procedure HowLongRF;
var Elapsed: cardinal;
  j:integer;
  X:real;
begin
  Elapsed:=GetTickCount64;
  X:=10;
  for j:=1 to 1000000 do X:=RF(X);
  Elapsed:=GetTickCount64-Elapsed;
  ShowMessage('Общее время выполнения: '+IntToStr(Elapsed)+' мс.'#13#10'Среднее время выполнения: '+FloatToStrF(Elapsed/1000,ffFixed,10,3)+' мкс.');
end;

procedure HowLongKS;
var Elapsed: cardinal;
  j:integer;
  flag:boolean;
  X:array [1..3] of real;
begin
  Elapsed:=GetTickCount64;
  X[1]:=439;X[2]:=450;X[3]:=451;
  for j:=1 to 1000000 do KS(X,flag);
  Elapsed:=GetTickCount64-Elapsed;
  ShowMessage('Общее время выполнения: '+IntToStr(Elapsed)+' мс.'#13#10'Среднее время выполнения: '+FloatToStrF(Elapsed/1000,ffFixed,10,3)+' мкс.');
end;

procedure HowLongWrite;
var Elapsed: cardinal;
  j:integer;
begin
  testform.label1.Visible:=true;
  Elapsed:=GetTickCount64;
  for j:=1 to 1000000 do testform.label1.Caption:=inttostr(i)+inttostr(i)+inttostr(i)+inttostr(i)+inttostr(i)+inttostr(i)+inttostr(i)+inttostr(i)+inttostr(i)+inttostr(i)+inttostr(i)+inttostr(i)+inttostr(i);
  Elapsed:=GetTickCount64-Elapsed;
  ShowMessage('Общее время выполнения: '+IntToStr(Elapsed)+' мс.'#13#10'Среднее время выполнения: '+FloatToStrF(Elapsed/1000,ffFixed,10,3)+' мкс.');
  testform.label1.Visible:=false;
end;

procedure HowLongSave;
var Elapsed: cardinal;
  j,i:integer;
  X:real;
  s:string;
begin
  Elapsed:=GetTickCount64;
  for j:=1 to 1000000 do
  begin
    for i:=1 to 14 do
      X:=i;
    S:='1234567890';
    s:=timetostr(now);
  end;
  Elapsed:=GetTickCount64-Elapsed;
  ShowMessage('Общее время выполнения: '+IntToStr(Elapsed)+' мс.'#13#10'Среднее время выполнения: '+FloatToStrF(Elapsed/1000,ffFixed,10,3)+' мкс.');
end;

{ TTestForm }

procedure TTestForm.Button2Click(Sender: TObject);
begin
  Testform.Visible:=false;
end;

procedure TTestForm.Button1Click(Sender: TObject);
begin
  case combobox1.ItemIndex of
    0: HowLongOpros(1);
    1: HowLongOpros(2);
    2: HowLongOpros(3);
    3: HowLongOpros(4);
    4: HowLongOpros(5);
    5: HowLongOpros(6);
    6: HowLongOpros(8);
    7: HowLongOpros(18);
    8: HowLongOpros(47);
    9: HowLongOpros(55);
    10: HowLongOpros(83);
    11: HowLongMean;
    12: HowLongDisp;
    13: HowLongKS;
    14: HowLongPO;
    15: HowLongRF;
    16: HowLongWrite;
    17: HowLongSave
    else ;
  end;
end;

end.

