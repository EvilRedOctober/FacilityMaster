unit main;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, Forms, Controls, Graphics, Dialogs, Menus, ExtCtrls, Spin,
  StdCtrls, ComCtrls, Buttons, DBGrids, DBCtrls, Grids, math, intrfcD, dbf, db,
  sqldb, pqconnection, odbcconn, mssqlconn, sqlite3conn, mysql57conn,
  mysql40conn, LConvEncoding, LazUTF8, TAGraph, TASeries, TASources,
  DateTimePicker,fileutil,fpspreadsheet, fpsTypes, xlsbiff8, fpsutils, fpsallformats;

type


   TFrame=record
     ch1,ch2,ch3,ch4,ch6,ch18,ch47,ch55: real;
     ch5: array[1..4] of real;
     ch8,ch83: array[1..3] of real;
     time:tdatetime;
     po1,po2,po3,po4: boolean;
     M5,D5,F55:extended;
   end;

  { TThreadPoll }

  TThreadPoll=class(TThread)
  public
    procedure Execute; override;
  end;

  { TMainForm }
  TMainForm = class(TForm)
    BtnExport: TBitBtn;
    BtnSetFilter: TBitBtn;
    BtnProtocol1: TBitBtn;
    BtnProtocol2: TBitBtn;
    BtnHisto: TBitBtn;
    BtnSetFilter1: TBitBtn;
    BtnStart: TBitBtn;
    BtnStat: TBitBtn;
    BtnStop: TBitBtn;
    BtnSave: TBitBtn;
    Chart1: TChart;
    Chart1BarSeries1: TBarSeries;
    ComboBox1: TComboBox;
    ComboBox2: TComboBox;
    ComboBox3: TComboBox;
    ComboBox4: TComboBox;
    DataSource1: TDataSource;
    Label8: TLabel;
    NTRDialog: TSaveDialog;
    ShowComment: TEdit;
    Label7: TLabel;
    ShowCount: TEdit;
    SpinInt: TSpinEdit;
    Label12: TLabel;
    ReportDialog: TSaveDialog;
    FMDataCH1: TFloatField;
    FMDataCH18: TFloatField;
    FMDataCH2: TFloatField;
    FMDataCH3: TFloatField;
    FMDataCH4: TLongintField;
    FMDataCH47: TFloatField;
    FMDataCH55: TFloatField;
    FMDataCH6: TLongintField;
    FMDataCH8: TFloatField;
    FMDataCH83: TFloatField;
    FMDataCOMMENT: TStringField;
    FMDataD5: TFloatField;
    FMDataDATETIME: TStringField;
    FMDataFUN55: TFloatField;
    FMDataM5: TFloatField;
    FMDataNAME: TStringField;
    FMDataPOL: TStringField;
    GroupBox4: TGroupBox;
    GroupBox5: TGroupBox;
    GroupBox6: TGroupBox;
    Label10: TLabel;
    Label11: TLabel;
    Label9: TLabel;
    Mini1: TFloatSpinEdit;
    Maxy1: TFloatSpinEdit;
    FMData: TDbf;
    DBGrid1: TDBGrid;
    EditName: TEdit;
    EditComment: TEdit;
    GroupBox1: TGroupBox;
    GroupBox2: TGroupBox;
    GroupBox3: TGroupBox;
    ImageList1: TImageList;
    Label1: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    Label4: TLabel;
    Label5: TLabel;
    Label6: TLabel;
    MainMenu1: TMainMenu;
    Memo1: TMemo;
    MenuItem2: TMenuItem;
    MenuAbout: TMenuItem;
    MenuExit: TMenuItem;
    PageControl1: TPageControl;
    FrameCount: TSpinEdit;
    ProgressBar1: TProgressBar;
    ExportDialog: TSaveDialog;
    SG1: TStringGrid;
    SGStats: TStringGrid;
    SQLQuery1: TSQLQuery;
    SQLTransaction1: TSQLTransaction;
    SGMeans: TStringGrid;
    StringGrid1: TStringGrid;
    TabSheet1: TTabSheet;
    TabSheet2: TTabSheet;
    TabSheet3: TTabSheet;
    procedure BtnExportClick(Sender: TObject);
    procedure BtnHistoClick(Sender: TObject);
    procedure BtnProtocol1Click(Sender: TObject);
    procedure BtnProtocol2Click(Sender: TObject);
    procedure BtnSetFilter1Click(Sender: TObject);
    procedure BtnSetFilterClick(Sender: TObject);
    procedure BtnStartClick(Sender: TObject);
    procedure BtnStatClick(Sender: TObject);
    procedure BtnStopClick(Sender: TObject);
    procedure BtnSaveClick(Sender: TObject);
    procedure ComboBox1Change(Sender: TObject);
    procedure ComboBox2Change(Sender: TObject);
    procedure DataSource1DataChange(Sender: TObject; Field: TField);
    procedure DBGrid1TitleClick(Column: TColumn);
    procedure FMDataFilterRecord(DataSet: TDataSet; var Accept: Boolean);
    function FMDataTranslate(Dbf: TDbf; Src, Dest: PChar; ToOem: Boolean
      ): Integer;
    procedure FormCreate(Sender: TObject);
    procedure Maxy1Change(Sender: TObject);
    procedure MenuAboutClick(Sender: TObject);
    procedure MenuExitClick(Sender: TObject);
    procedure MenuItem2Click(Sender: TObject);
    procedure Mini1Change(Sender: TObject);
    Procedure CountMeans(Sender: TObject);
    Procedure CountStats(Sender: TObject);
    Procedure AllText(Sender: TField; var AText: string;  DisplayText: Boolean);
  private

  public
    ThreadPoll:TThreadPoll;


  end;

  Function PO(X,Xmax,Xmin:real): boolean;
  Procedure KS(X:array of real; var flag:boolean);
  function RF(X:real):real;


var
  MainForm: TMainForm;
  datastart, datafinish: TDateTime;
  N,i,channel: integer;
  wspom:massiv;
  PollOn,CancelPoll:boolean;
  Data: array of TFrame;
  isStable, isPos: boolean;
  ChMax,ChMin,FMax,FMin: array [0..12] of real;
  FindMax,FindMin: real;
  FIO: string;


implementation

uses test,about;

{$R *.lfm}

{ TMainForm }

Function PO(X,Xmax,Xmin:real): boolean; //Контроль позиционных ограничений п1
begin
    if (X>Xmax)or(X<Xmin) then result:=false else result:= true;
end;

Procedure KS(X:array of real; var flag:boolean); //Контроль стабильности п1
var Xmax,Xmin:real;
begin
  Xmax:=maxvalue(X); Xmin:=minvalue(X);
  if Xmax<>Xmin then
  begin
    Xmax:=(Xmax-Xmin)/Xmin;
    if abs(Xmax)>0.1 then flag:=false;
  end;
end;

function RF(X:real):real; //Функция для 55 канала п1
begin
  result:=0.11*ln(X+20);
end;

Procedure TMainForm.CountMeans(Sender: TObject); //Рассчет средних п2
var Len:integer;
  M:array [0..12] of real;
  j:byte;
begin
  FMData.DisableControls;
  Len:=FMData.ExactRecordCount;
  for j:=0 to 12 do M[j]:=0;
  if Len>0 then
  begin
    FMData.First;
    while not (FMData.EOF) do
    begin
      M[0]+=FMData.FieldByName('Ch1').AsFloat;
      M[1]+=FMData.FieldByName('Ch2').AsFloat;
      M[2]+=FMData.FieldByName('Ch3').AsFloat;
      M[3]+=FMData.FieldByName('Ch4').AsFloat;
      M[4]+=FMData.FieldByName('M5').AsFloat;
      M[5]+=FMData.FieldByName('D5').AsFloat;
      M[6]+=FMData.FieldByName('Ch6').AsFloat;
      M[7]+=FMData.FieldByName('Ch8').AsFloat;
      M[8]+=FMData.FieldByName('Ch18').AsFloat;
      M[9]+=FMData.FieldByName('Ch47').AsFloat;
      M[10]+=FMData.FieldByName('Ch55').AsFloat;
      M[11]+=FMData.FieldByName('Fun55').AsFloat;
      M[12]+=FMData.FieldByName('Ch83').AsFloat;
      FMData.Next;
    end;
    for j:=0 to 12 do
    begin
      M[j]:=M[j]/Len;
      SGMeans.Cells[j+1,1]:=FloatToStrF(M[j],fffixed,10,2);
    end;
  end else
    for j:=0 to 12 do
      SGMeans.Cells[j+1,1]:='';

  FMData.EnableControls;
end;

Procedure TMainForm.CountStats(Sender: TObject); //Рассчет статистик п3
var Len:integer;
  m1,m2,m3,m4,skew,kurtosis: extended;
  SData: array [0..12] of array of real;
  j:byte;
  k,l:integer;
begin
  FMData.DisableControls;
  Len:=FMData.ExactRecordCount;
  for j:=0 to 12 do begin SetLength(SData[j],Len) end;
  if Len>0 then
  begin
    FMData.First;
    k:=0;
    while not (FMData.EOF) do
    begin
      SData[0][k]:=FMData.FieldByName('Ch1').AsFloat;
      SData[1][k]:=FMData.FieldByName('Ch2').AsFloat;
      SData[2][k]:=FMData.FieldByName('Ch3').AsFloat;
      SData[3][k]:=FMData.FieldByName('Ch4').AsFloat;
      SData[4][k]:=FMData.FieldByName('M5').AsFloat;
      SData[5][k]:=FMData.FieldByName('D5').AsFloat;
      SData[6][k]:=FMData.FieldByName('Ch6').AsFloat;
      SData[7][k]:=FMData.FieldByName('Ch8').AsFloat;
      SData[8][k]:=FMData.FieldByName('Ch18').AsFloat;
      SData[9][k]:=FMData.FieldByName('Ch47').AsFloat;
      SData[10][k]:=FMData.FieldByName('Ch55').AsFloat;
      SData[11][k]:=FMData.FieldByName('Fun55').AsFloat;
      SData[12][k]:=FMData.FieldByName('Ch83').AsFloat;
      FMData.Next;
      k+=1;
    end;
    for j:=0 to 12 do
    begin
      SGStats.Cells[j+1,1]:=FloatToStrF(Mean(SData[j]),fffixed,10,2);
      for k:=0 to len-2 do
        for l:=k+1 to len-1 do
        if SData[j][k]>SData[j][l] then
        begin
          m1:=SData[j][l];
          SData[j][l]:=SData[j][k];
          SData[j][k]:=m1;
        end;
      if (len mod 2)=1 then
      m1:=SData[j][len div 2]
      else  m1:=(SData[j][len div 2]+SData[j][(len div 2)-1])*0.5;
      SGStats.Cells[j+1,2]:=FloatToStrF(m1,fffixed,10,2); m1:=0;
      SGStats.Cells[j+1,3]:=FloatToStrF(StdDev(SData[j]),fffixed,10,2);
      SGStats.Cells[j+1,4]:=FloatToStrF(maxvalue(SData[j]),fffixed,10,2);
      SGStats.Cells[j+1,5]:=FloatToStrF(minvalue(SData[j]),fffixed,10,2);
      SGStats.Cells[j+1,6]:=FloatToStrF(maxvalue(SData[j])-minvalue(SData[j]),fffixed,10,2);
      momentskewkurtosis(SData[j],m1,m2,m3,m4,skew,kurtosis);
      SGStats.Cells[j+1,7]:=FloatToStrF(skew,fffixed,10,2);
      SGStats.Cells[j+1,8]:=FloatToStrF(kurtosis-3,fffixed,10,2);
      if isNan(skew) then SGStats.Cells[j+1,7]:='Не опр.';if isNan(kurtosis) then SGStats.Cells[j+1,8]:='Не опр.';

    end;
  end else
  for j:=1 to 13 do
    for l:=1 to 8 do
      SGStats.Cells[j,l]:='';
  FMData.EnableControls;
end;

{ ThreadPoll }

procedure TThreadPoll.Execute; //Поток регистрации п1
var Elapsed: cardinal;
  j:integer;
  S:string;
  DT:tdatetime;
begin
  Elapsed:=GetTickCount64;
  with MainForm do
  begin
    setlength(Data,N);
    BtnStart.Enabled:=false;
    BtnStop.Enabled:=true;
    PollOn:=true;
    ProgressBar1.Max:=N;
    ProgressBar1.Position:=0;
    DT:=now;
    for i:=0 to N-1 do
    begin
      repeat
        ProgressBar1.Position:=i+1;
        isStable:=true;
        Data[i].time:=DT;
        opros(1,Data[i].ch1,wspom);
        opros(2,Data[i].ch2,wspom);
        opros(3,Data[i].ch3,wspom);
        opros(4,Data[i].ch4,wspom);
        opros(6,Data[i].ch6,wspom);
        opros(18,Data[i].ch18,wspom);
        opros(47,Data[i].ch47,wspom);
        opros(55,Data[i].ch55,wspom);
        opros(5,Data[i].ch5[1],wspom);opros(5,Data[i].ch5[2],wspom);opros(5,Data[i].ch5[3],wspom);opros(5,Data[i].ch5[4],wspom);
        opros(8,Data[i].ch8[1],wspom);opros(8,Data[i].ch8[2],wspom);opros(8,Data[i].ch8[3],wspom);
        opros(83,Data[i].ch83[1],wspom);opros(83,Data[i].ch83[2],wspom);opros(83,Data[i].ch83[3],wspom);
        Data[i].f55:=RF(Data[i].ch55);
        MeanAndStdDev(Data[i].ch5,Data[i].M5,Data[i].D5);
        KS(Data[i].ch8,isStable);
        if not(isStable) then Memo1.Lines.Add('Кадр №'+inttostr(i+1)+':Нет стабильности канала 8!');
        KS(Data[i].ch83,isStable);
        if isStable then
           if not(isStable) then Memo1.Lines.Add('Кадр №'+inttostr(i+1)+':Нет стабильности канала 83!');
        Data[i].po1:=PO(Data[i].ch1,30,-35);
        Data[i].po2:=PO(Data[i].ch2,765,740);
        Data[i].po3:=PO(Data[i].ch3,100,50);
        Data[i].po4:=PO(Data[i].ch4,1,0);
        If not(Data[i].po1) then Memo1.Lines.Add('Кадр №'+inttostr(i+1)+': Нарушение позиционных ограничений по каналу 1.');
        If not(Data[i].po2) then Memo1.Lines.Add('Кадр №'+inttostr(i+1)+': Нарушение позиционных ограничений по каналу 2.');
        If not(Data[i].po3) then Memo1.Lines.Add('Кадр №'+inttostr(i+1)+': Нарушение позиционных ограничений по каналу 3.');
        If not(Data[i].po3) then Memo1.Lines.Add('Кадр №'+inttostr(i+1)+': Нарушение позиционных ограничений по каналу 4.');
      until isStable;
      if not(Data[i].po1) or not(Data[i].po2) or not(Data[i].po3) or not(Data[i].po4) then Memo1.Lines.Add('К1: '+FloatToStrF(Data[i].ch1,fffixed,10,2)+'; К2: '+FloatToStr(Data[i].ch2)+
        '; К3: '+FloatToStr(Data[i].ch3)+'; К4: '+FloatToStr(Data[i].ch4)+'; Среднее К5: '+FloatToStr(Data[i].M5)+
        '; Дисперсия К5: '+FloatToStrF(Data[i].D5,fffixed,10,2)+'; К6: '+FloatToStr(Data[i].ch6)+
        '; К8: '+FloatToStr(mean(Data[i].ch8))+'; К18: '+FloatToStrF(Data[i].ch18,fffixed,10,2)+
        '; К47: '+FloatToStrF(Data[i].ch47,fffixed,10,2)+'; К55: '+FloatToStrF(Data[i].ch55,fffixed,10,2)+
        '; Функци К55: '+FloatToStrF(Data[i].F55,fffixed,10,2)+'; К83: '+FloatToStr(mean(Data[i].ch83))+'.');
            //Memo1.Lines.Add(FloatToStr(Data[i].ch1));}
      if CancelPoll then break;
    end;
    N:=i;
    for j:=0 to i do
    begin
      S:='';
      If not(Data[j].po1) then s+='К1,';
      If not(Data[j].po2) then s+='К2,';
      If not(Data[j].po3) then s+='К3,';
      If not(Data[j].po3) then s+='К4,';
      delete(s,s.Length,1);
      If s='' then s:='Нет';
      SG1.RowCount:=2+i;
      //SG1.Cells[0,j+1]:=DateTimeToStr(Data[j].time);
      SG1.Cells[1,j+1]:=FloatToStrf(Data[j].ch1,fffixed,10,2);SG1.Cells[2,j+1]:=FloatToStrf(Data[j].ch2,fffixed,10,2);
      SG1.Cells[3,j+1]:=FloatToStrf(Data[j].ch3,fffixed,10,2);SG1.Cells[4,j+1]:=FloatToStrf(Data[j].ch4,fffixed,10,2);
      SG1.Cells[5,j+1]:=FloatToStrf(Data[j].M5,fffixed,10,2);SG1.Cells[6,j+1]:=FloatToStrf(Data[j].D5,fffixed,10,2);
      SG1.Cells[7,j+1]:=FloatToStrf(Data[j].ch6,fffixed,10,2);SG1.Cells[8,j+1]:=FloatToStrf(Data[j].ch8[1],fffixed,10,2);
      SG1.Cells[9,j+1]:=FloatToStrf(Data[j].ch18,fffixed,10,2);SG1.Cells[10,j+1]:=FloatToStrf(Data[j].ch47,fffixed,10,2);
      SG1.Cells[11,j+1]:=FloatToStrf(Data[j].ch55,fffixed,10,2);SG1.Cells[12,j+1]:=FloatToStrf(Data[j].F55,fffixed,10,2);
      SG1.Cells[13,j+1]:=FloatToStrf(Data[j].ch83[0],fffixed,10,2);SG1.Cells[14,j+1]:=S;
    end;
    Elapsed:=GetTickCount64-Elapsed;
    Memo1.Lines.Add('Дата и время: '+DateTimeToStr(dt));
    Memo1.Lines.Add('Конец регистрации. Время регистрации: '+IntToStr(Elapsed)+' мс');
    PollOn:=false;
    BtnStart.Enabled:=True;
    BtnStop.Enabled:=False;
    BtnSave.Enabled:=True;
  end;
end;

procedure TMainForm.MenuExitClick(Sender: TObject); //Выход
begin
    case QuestionDLG('Мастер Объекта','Желаете покинуть программу?',mtCustom,[mrYes,'Да', mrNo, 'Нет', 'IsDefault'],'') of
      mrYes: close;
      mrNo: ;
    end;
end;

procedure TMainForm.MenuItem2Click(Sender: TObject); //Проверка скорости выполнения
begin
  testform.show;
end;

procedure TMainForm.Mini1Change(Sender: TObject); //Контроль параметров поиска для мин. п2
var nom:byte;
begin
  nom:= ComboBox1.ItemIndex;
  FMin[nom]:=Mini1.Value;
end;


procedure TMainForm.BtnStartClick(Sender: TObject); //Начало регистрации п1
var i:integer;
  frame:tframe;
begin
  if not(PollOn) then
  begin
    CancelPoll:=false;
    N:=FrameCount.Value;
    randomize;
    memo1.Clear;
    ThreadPoll:=Tthreadpoll.Create(false);
    ThreadPoll.FreeOnTerminate:=True;
    ThreadPoll.Priority:=tpHigher;
  end;
  BtnSave.Enabled:=true;
end;

procedure TMainForm.BtnStatClick(Sender: TObject);
begin
  Countstats(Sender);
end;

procedure TMainForm.BtnHistoClick(Sender: TObject); //Гистограмма п3
  var Len:integer;
  HData,Y,X:array  of real;
  j,M,k:integer;
  HLeft,HRight,HTek,HStep: real;
begin
  FMData.DisableControls;
  Len:=FMData.ExactRecordCount;
  setlength(HData,len);
  if Len>0 then
  begin
    FMData.First;
    j:=0;
    while not (FMData.EOF) do
    begin
      HData[j]:=FMData.Fields.Fields[Combobox3.ItemIndex].AsFloat;
      FMData.Next;
      j+=1;
    end;
    Chart1BarSeries1.Clear;
    HLeft:=MinValue(HData);HRight:=MaxValue(HData);
    M:=1+round(ln(Len)/ln(2));
    M:=SpinInt.Value;
    setlength(Y,m);
    HTek:=HLeft;HStep:=(HRight-HLeft)/M;
    for k:=0 to M-2 do
    begin
      for j:=0 to len-1 do
        if (HData[j]>=HTek) and (HData[j]<(HTek+HStep))
          then Y[k]:=Y[k]+1;
      Y[k]:=Y[k];//len/HStep;
      //Chart1BarSeries1.add(Y[k], '['+floattostrf(HTek,fffixed,10,2)+';'+floattostrf(HTek+HStep,fffixed,10,2)+')');
      Chart1BarSeries1.AddXY((HTek+Hstep*0.5),Y[k]);
      HTek:=HTek+HStep;
    end;
    for j:=0 to len-1 do
      if (HData[j]>=HTek) and (HData[j]<=HRight)
        then Y[m-1]:=Y[m-1]+1;
    Y[m-1]:=Y[m-1];//len/HStep;
    //Chart1BarSeries1.add(Y[m-1], '['+floattostrf(HTek,fffixed,10,2)+';'+floattostrf(HRight,fffixed,10,2)+']');
    Chart1BarSeries1.AddXY((HTek+Hstep*0.5),Y[m-1]);
    Chart1.Title.text.text:=Combobox3.items[Combobox3.ItemIndex];

  end;
  FMData.EnableControls;
end;

procedure TMainForm.BtnProtocol1Click(Sender: TObject); //Отчет по данным
var MyWorkbook: TsWorkbook; //AVERAGE(num1 [, num2, ...] )
 MyWorksheet: TsWorksheet;
 j,k: integer;
 Symb: integer;
begin
  Symb:=ord('A');
  if FMData.ExactRecordCount<>0 then
  begin
    ReportDialog.InitialDir := GetCurrentDir;
    if ReportDialog.Execute then
    begin
      FMData.DisableControls;
      //FMData.Filtered:=true;
      FMData.Close;
      FMData.Open;
      FMData.First;
      k:=6;

      MyWorkbook := TsWorkbook.Create;//Создание рабочей книги
      MyWorksheet := MyWorkbook.AddWorksheet('Отчёт по данным');//Создание листа с данными
      MyWorksheet.PageLayout.Orientation:=spoLandscape;//Ориентация при печати
      MyWorkbook.SetDefaultFont('Times New Roman',10); //Стандартный шрифт
      MyWorkbook.Options := MyWorkbook.Options + [boReadFormulas, boCalcBeforeSaving, boAutoCalc]; //Установка опций для формул
      MyWorksheet.PageLayout.SetRepeatedCols(0,13); //Ссылка на столбцы, которые должны повторяться при печати
      MyWorksheet.PageLayout.SetRepeatedRows(6,6);  //Ссылка на строки, которые должны повторяться при печати
      MyWorksheet.WriteColWidth(0,8.5);//Установка ширины столбцов
      MyWorksheet.WriteColWidth(1,8.5);MyWorksheet.WriteColWidth(2,8);MyWorksheet.WriteColWidth(3,8);MyWorksheet.WriteColWidth(4,10);
      MyWorksheet.WriteColWidth(5,12.5);MyWorksheet.WriteColWidth(6,8);MyWorksheet.WriteColWidth(7,8);MyWorksheet.WriteColWidth(8,9);
      MyWorksheet.WriteColWidth(9,10);MyWorksheet.WriteColWidth(10,10);MyWorksheet.WriteColWidth(11,12);MyWorksheet.WriteColWidth(12,9);
      MyWorksheet.WriteColWidth(13,15.2);
      //Подготовка верхней части
      MyWorksheet.MergeCells('A1:N1'); //Объединение ячеек
      MyWorksheet.WriteFontStyle(0, 0, [fssBold]);  //Сделать полужирным
      MyWorksheet.WriteFontSize(0,0,12); //Смена размера шрифта
      MyWorksheet.WriteHorAlignment(0, 0, haCenter); //Центрировать по горизонтали
      MyWorksheet.WriteVertAlignment(0, 0, vaCenter); //Центрировать по вертикали
      MyWorksheet.WriteText(0,0,'Отчёт по данным'); //Вставить в ячейку текст
      MyWorksheet.MergeCells('A3:B3');MyWorksheet.MergeCells('C3:I3');MyWorksheet.MergeCells('J3:K3');MyWorksheet.MergeCells('L3:N3');
      MyWorksheet.MergeCells('A5:B5');MyWorksheet.MergeCells('C5:I5');MyWorksheet.MergeCells('J5:K5');MyWorksheet.MergeCells('L5:N5');
      MyWorksheet.WriteFontStyle(2, 0, [fssBold]);MyWorksheet.WriteFontStyle(2, 9, [fssBold]);MyWorksheet.WriteFontStyle(4, 0, [fssBold]);;MyWorksheet.WriteFontStyle(4, 9, [fssBold]);
      MyWorksheet.WriteText(2,0,'ФИО оператора:');MyWorksheet.WriteText(2,2,CP1251ToUTF8(FIO));
      MyWorksheet.WriteText(2,9,'Дата эксперимента:');MyWorksheet.WriteText(2,11,FMData.FieldByName('DateTime').AsString);
      MyWorksheet.WriteText(4,0,'Комментарий:');MyWorksheet.WriteText(4,2,CP1251ToUTF8(FMData.FieldByName('Comment').AsString));
      MyWorksheet.WriteText(4,9,'Число кадров:');MyWorksheet.WriteNumber(4,11,FMData.ExactRecordCount);MyWorksheet.WriteHorAlignment(4, 11, haLeft);
      MyWorksheet.PageLayout.Footers[HEADER_FOOTER_INDEX_ALL] := '&RСтраница &P из &N';

      //Пишет заголовки
      for j:=0 to 13 do
      begin
        MyWorksheet.WriteFontStyle(6, j, [fssBold]);
        MyWorksheet.WriteText(6,j,DBgrid1.Columns[j+2].Title.Caption);
        MyWorksheet.WriteBorders(6, j, [cbNorth, cbWest, cbEast, cbSouth]);//Рисует границы
      end;
      k+=1;

      while not (FMData.EOF) do //Вставка данных
      begin
        MyWorksheet.WriteText(k,13,CP1251ToUTF8(FMData.FieldByName('POL').AsString));
        MyWorksheet.WriteBorders(k, 13, [cbNorth, cbWest, cbEast, cbSouth]);
        for j:=0 to 12 do
        begin
          MyWorksheet.WriteNumber(k,j,FMData.Fields.Fields[j].AsFloat,nfFixed);
          MyWorksheet.WriteBorders(k, j, [cbNorth, cbWest, cbEast, cbSouth]);
        end;
        FMData.Next;
        k+=1;
      end;

      k+=1;
      MyWorksheet.MergeCells('A'+inttostr(k+1)+':N'+inttostr(k+1));
      MyWorksheet.WriteHorAlignment(k, 0, haCenter); MyWorksheet.WriteVertAlignment(k, 0, vaCenter);
      MyWorksheet.WriteFontStyle(k, 0, [fssBold]); MyWorksheet.WriteText(k,0,'Средние значения по данным');
      k+=1;

      for j:=0 to 12 do
      begin
        MyWorksheet.WriteBorders(k, j, [cbNorth, cbWest, cbEast, cbSouth]);
        MyWorksheet.WriteNumber(k,j,0,nfFixed);
        MyWorksheet.WriteFormula(k, j, '=AVERAGE('+chr(Symb+j)+'8:'+chr(Symb+j)+inttostr(k-2)+')');
      end;


      MyWorkbook.WriteToFile(ReportDialog.FileName, sfOOXML, True);
      FMData.EnableControls;
      ShowMessage('Успешное сохранение!');
    end else ShowMessage('Отмена сохранения.');

  end;
end;

procedure TMainForm.BtnProtocol2Click(Sender: TObject);
var MyWorkbook: TsWorkbook; //AVERAGE(num1 [, num2, ...] )
 MyWorksheet: TsWorksheet;
 j,k: integer;
begin
  Countstats(Sender);
  BtnHistoClick(Sender);
  if SGStats.Cells[1,1]<>'' then
  begin
    NTRDialog.InitialDir := GetCurrentDir;
    if NTRDialog.Execute then
    begin
      FMData.DisableControls;
      //FMData.Filtered:=true;
      FMData.Close;
      FMData.Open;
      FMData.First;
      k:=6;

      MyWorkbook := TsWorkbook.Create;//Создание рабочей книги
      MyWorksheet := MyWorkbook.AddWorksheet('Отчёт по НТР');//Создание листа с данными
      MyWorksheet.PageLayout.Orientation:=spoLandscape;//Ориентация при печати
      MyWorkbook.SetDefaultFont('Times New Roman',10); //Стандартный шрифт
      MyWorkbook.Options := MyWorkbook.Options + [boReadFormulas, boCalcBeforeSaving, boAutoCalc]; //Установка опций для формул
      MyWorksheet.WriteColWidth(1,8);//Установка ширины столбцов
      MyWorksheet.WriteColWidth(2,8);MyWorksheet.WriteColWidth(3,8);MyWorksheet.WriteColWidth(4,8);MyWorksheet.WriteColWidth(5,10);
      MyWorksheet.WriteColWidth(6,12.5);MyWorksheet.WriteColWidth(7,8);MyWorksheet.WriteColWidth(8,8);MyWorksheet.WriteColWidth(9,10.5);
      MyWorksheet.WriteColWidth(10,10);MyWorksheet.WriteColWidth(11,9);MyWorksheet.WriteColWidth(12,12);MyWorksheet.WriteColWidth(13,9.5);
      MyWorksheet.WriteColWidth(0,12);
      //Подготовка верхней части
      MyWorksheet.MergeCells('A1:N1'); //Объединение ячеек
      MyWorksheet.WriteFontStyle(0, 0, [fssBold]);  //Сделать полужирным
      MyWorksheet.WriteFontSize(0,0,12); //Смена размера шрифта
      MyWorksheet.WriteHorAlignment(0, 0, haCenter); //Центрировать по горизонтали
      MyWorksheet.WriteVertAlignment(0, 0, vaCenter); //Центрировать по вертикали
      MyWorksheet.WriteText(0,0,'Отчёт по научно-техническим расчетам'); //Вставить в ячейку текст
      MyWorksheet.MergeCells('A3:B3');MyWorksheet.MergeCells('C3:I3');MyWorksheet.MergeCells('J3:K3');MyWorksheet.MergeCells('L3:N3');
      MyWorksheet.MergeCells('A5:B5');MyWorksheet.MergeCells('C5:I5');MyWorksheet.MergeCells('J5:K5');MyWorksheet.MergeCells('L5:N5');
      MyWorksheet.WriteFontStyle(2, 0, [fssBold]);MyWorksheet.WriteFontStyle(2, 9, [fssBold]);MyWorksheet.WriteFontStyle(4, 0, [fssBold]);;MyWorksheet.WriteFontStyle(4, 9, [fssBold]);
      MyWorksheet.WriteText(2,0,'ФИО оператора:');MyWorksheet.WriteText(2,2,CP1251ToUTF8(FIO));
      MyWorksheet.WriteText(2,9,'Дата эксперимента:');MyWorksheet.WriteText(2,11,FMData.FieldByName('DateTime').AsString);
      MyWorksheet.WriteText(4,0,'Комментарий:');MyWorksheet.WriteText(4,2,CP1251ToUTF8(FMData.FieldByName('Comment').AsString));
      MyWorksheet.WriteText(4,9,'Число кадров:');MyWorksheet.WriteNumber(4,11,FMData.ExactRecordCount);MyWorksheet.WriteHorAlignment(4, 11, haLeft);
      MyWorksheet.PageLayout.Footers[HEADER_FOOTER_INDEX_ALL] := '&RСтраница &P из &N';

      //Вставка таблицы со статистикой
      for k:=1 to 8 do
      begin
        MyWorksheet.WriteText(6+k,0,SGStats.Cells[0,k]);
        MyWorksheet.WriteBorders(6+k, 0, [cbNorth, cbWest, cbEast, cbSouth]);
      end;
      for j:=0 to 13 do
      begin
        MyWorksheet.WriteFontStyle(6, j, [fssBold]);
        MyWorksheet.WriteText(6,j,SGStats.Cells[j,0]);
        MyWorksheet.WriteBorders(6, j, [cbNorth, cbWest, cbEast, cbSouth]);
      end;

      for j:=1 to 13 do
        for k:=1 to 8 do
        begin
          MyWorksheet.WriteBorders(6+k, j, [cbNorth, cbWest, cbEast, cbSouth]);//Рисует границы
          if (SGStats.Cells[j,k]<>'') and (SGStats.Cells[j,k]<>'Не опр.') then MyWorksheet.WriteNumber(k+6,j,StrToFloat(SGStats.Cells[j,k]),nfFixed)
          else MyWorksheet.WriteText(6+k,j, SGStats.Cells[j,k]);
        end;

      Chart1.SaveToBitmapFile('Chart1.bmp');
      MyWorksheet.MergeCells('C17:L31');
      MyWorksheet.WriteHorAlignment(16, 2, haCenter); //Центрировать по горизонтали
      MyWorksheet.WriteVertAlignment(16, 2, vaCenter); //Центрировать по вертикали
      MyWorksheet.WriteImage(16, 2, 'Chart1.bmp',0,0,0.028,0.0140);


      //MyWorksheet.GetImage();
      deletefile('Chart1.bmp');

      MyWorkbook.WriteToFile(NTRDialog.FileName, sfOOXML, True);
      ShowMessage('Успешное сохранение!');
    end else ShowMessage('Отмена сохранения.');

  end;
end;

procedure TMainForm.BtnExportClick(Sender: TObject);
var s: string;
begin
  if FMData.ExactRecordCount<>0 then
  begin

    ExportDialog.InitialDir := GetCurrentDir;
    if ExportDialog.Execute then
    begin
      FMData.Close;
      s:=GetCurrentDir+'\'+FMData.TableName;

      CopyFile(s,ExportDialog.FileName,false);

      FMData.TableName:=ExportDialog.FileName;
      FMData.Open;

      FMData.Filtered:=true;
      FMData.First;
      //
      while not (FMData.EOF) do
      begin
        FMData.Edit;
        FMData.FieldByName('Ch8').AsFloat:=1;
        FMData.Next;
      end;
      //FMData.Post;
      FMData.Filtered:=false;
      FMData.ShowDeleted:=true;
      FMData.First;
      while not (FMData.EOF) do
      begin
        if FMData.FieldByName('Ch8').AsFloat=0 then FMData.Delete
        else begin FMData.Edit; FMData.FieldByName('Ch8').AsFloat:=0; end;
        FMData.Next;
      end;
      FMData.PackTable;
      FMData.ShowDeleted:=false;
      FMData.close;
      FMData.TableName:=S;
      FMData.Open;
      FMData.Filtered:=true;
      ShowMessage('Успешный экспорт!');
    end else ShowMessage('Отмена экспорта.');

  end;
end;

procedure TMainForm.BtnSetFilter1Click(Sender: TObject);
begin
     FMData.Filtered:=false;
     FMData.Refresh;
     CountMeans(Sender);
end;

procedure TMainForm.BtnSetFilterClick(Sender: TObject); //Применение фильтра
begin
  if (Combobox4.ItemIndex>-1) and (Combobox2.ItemIndex>-1) then
  begin
    FMData.DisableControls;
    FMData.Filtered:=false;
    datastart:=StrToDateTime(Combobox4.Items[Combobox4.ItemIndex]);
    channel:=Combobox1.ItemIndex;
    FindMin:=Mini1.Value;FindMax:=Maxy1.Value;
    FIO:=UTF8ToCP1251(Combobox2.Items[Combobox2.ItemIndex]);

    FMData.Filtered:=true;

    CountMeans(Sender);
    MainForm.DBGrid1TitleClick(DBGrid1.Columns.Items[1]);
    FMData.EnableControls;
    FMData.Indexes.Update;
    FMData.Refresh;
    DBgrid1.Refresh;
    if FMDATA.ExactRecordCount>0 then ShowComment.Text:=CP1251ToUTF8(FMData.FieldByName('Comment').AsString)
    else ShowComment.Text:='';
    if FMDATA.ExactRecordCount<20 then SpinInt.MaxValue:=FMDATA.ExactRecordCount else SpinInt.MaxValue:=50;
    ShowCount.Text:=IntToStr(FMDATA.ExactRecordCount);


  end;
end;

procedure TMainForm.BtnStopClick(Sender: TObject); //Остановить регистрацию п1
begin
  CancelPoll:=True;
end;

procedure TMainForm.AllText(Sender: TField; var AText: string;  DisplayText: Boolean); //Перекодировка при чтении таблицы
begin
   AText:=CP1251ToUTF8(Sender.AsString);
end;

procedure TMainForm.BtnSaveClick(Sender: TObject); //Сохранение данных п1
var j:integer;
  S,s1,s2: string;
  var Elapsed: cardinal;
  flag:Boolean;
begin
  //HowLong(83);
  FMData.Filtered:=False;
  Elapsed:=GetTickCount64;
  Memo1.Lines.Add('Начало сохранения. Ожидайте...');
  ProgressBar1.Max:=i+1;
  ProgressBar1.Position:=0;
  s1:=EditName.Text;
  s2:= EditComment.Text;
  for j:=0 to i do
  begin
    S:='';
    ProgressBar1.Position:=j+1;
    If not(Data[j].po1) then s+='К1,';
    If not(Data[j].po2) then s+='К2,';
    If not(Data[j].po3) then s+='К3,';
    If not(Data[j].po3) then s+='К4,';
    delete(s,s.Length,1);
    If s='' then s:='Нет';
    FMData.Open;
    FMData.CodePage;
    FMData.append;
    FMData.FieldByName('Name').AsString:=UTF8ToCP1251(s1);
    FMData.FieldByName('DateTime').AsString:=(DateTimeToStr(Data[j].time));
    FMData.FieldByName('Comment').AsString:=UTF8ToCP1251(s2);
    FMData.FieldByName('Ch1').AsFloat:=Data[j].ch1;
    FMData.FieldByName('Ch2').AsFloat:=Data[j].ch2;
    FMData.FieldByName('Ch3').AsFloat:=Data[j].ch3;
    FMData.FieldByName('Ch4').AsFloat:=Data[j].ch4;
    FMData.FieldByName('M5').AsFloat:=Data[j].M5;
    FMData.FieldByName('D5').AsFloat:=Data[j].D5;
    FMData.FieldByName('Ch6').AsFloat:=Data[j].ch6;
    FMData.FieldByName('Ch8').AsFloat:=Data[j].ch8[1];
    FMData.FieldByName('Ch18').AsFloat:=Data[j].ch18;
    FMData.FieldByName('Ch47').AsFloat:=Data[j].ch47;
    FMData.FieldByName('Ch55').AsFloat:=Data[j].ch55;
    FMData.FieldByName('FUN55').AsFloat:=Data[j].f55;
    FMData.FieldByName('Ch83').AsFloat:=Data[j].ch83[1];
    FMData.FieldByName('Pol').AsString:=UTF8ToCP1251(S);
    FMData.Post;
    FMData.Close;
    {SQLQuery1.Close;
    SQLQuery1.SQL.Text := 'insert into FMData(Name,DATETIME,Ch1,Ch2,Ch3,Ch4,M5,D5,'+
      'Ch6,Ch8,Ch18,CH47,Ch55,Ch83,Pol,Comment) values(EditName.Text, Data[i].time, Data[i].ch1,'+
      'Data[i].ch2, Data[i].ch3, Data[i].ch4, Data[i].M5, Data[i].D5, Data[i].ch6, Data[i].ch8,'+
      'Data[i].ch18, Data[i].ch47, Data[i].ch55, Data[i].ch83, '', editcomment.Text);';
    SQLQuery1.ExecSQL;
    SQLTransaction1.Commit;}
  end;
  Elapsed:=GetTickCount64-Elapsed;
  Memo1.Lines.Add('Сохранение завершено. Время сохранения: '+IntToStr(Elapsed)+' мс');
  DBGrid1.Refresh;
  FMData.Close;
  FMData.Open;
  FMData.RegenerateIndexes;
  BtnSave.Enabled:=False;
  CountMeans(Sender);
  flag:=true;
  for j:=0 to ComboBox2.Items.Count-1 do
  begin
    If S1=ComboBox2.Items[j] then
    begin
      flag:=false;
      break;
    end;
  end;
  FMData.Next;
  if flag then ComboBox2.Items.Append(S1);
  ComboBox2Change(ComboBox2);
  FMData.Filtered:=True;
end;

procedure TMainForm.ComboBox1Change(Sender: TObject); //Контроль параметров поиска макс и мин. при смене канала п2
var max,min,step:real;
  nom:byte;
begin
  nom:= ComboBox1.ItemIndex;
  step:=(ChMax[nom]-ChMin[nom])/100;
  if (nom=3)or(nom=6) then step:=1;
  Maxy1.MaxValue:=ChMax[nom];Maxy1.MinValue:=ChMin[nom];Maxy1.Increment:=step;
  Mini1.MaxValue:=ChMax[nom];Mini1.MinValue:=ChMin[nom];Mini1.Increment:=step;
  Mini1.Value:=ChMin[nom];Maxy1.Value:=ChMax[nom];
end;

procedure TMainForm.ComboBox2Change(Sender: TObject);
var S1,s2: string;
  flag: boolean;
  len,j:integer;
begin
  ComboBox4.ItemIndex:=-1;
  ComboBox4.Clear;
  ComboBox4.items.Clear;
  if ComboBox2.ItemIndex>-1 then
  begin
  FMData.Filtered:=false;
  FMData.DisableControls;
  s2:=UTF8ToCP1251(Combobox2.Items[Combobox2.ItemIndex]);
  Len:=FMData.ExactRecordCount;
  if Len>0 then
  begin
    FMData.First;
    while not (FMData.EOF) do
    begin
      S1:=FMData.FieldByName('Datetime').AsString;
      flag:=true;
      for j:=0 to ComboBox4.Items.Count-1 do
      begin
        If (S1=ComboBox4.Items[j]) or ((FMData.FieldByName('Name').AsString)<>s2) then
        begin
          flag:=false;
          break;
        end;
      end;
      FMData.Next;
      if flag then ComboBox4.Items.Append(S1);
    end;
  end;
  if ComboBox4.Items.Count>0 then ComboBox4.ItemIndex:=0;
  FMData.Filtered:=true;
  end;
end;

procedure TMainForm.DataSource1DataChange(Sender: TObject; Field: TField);
begin

end;


procedure TMainForm.DBGrid1TitleClick(Column: TColumn);//Сортировка п2
var tag0:integer;
begin

  tag0:=column.tag;
  for i:=0 to 16 do dbgrid1.Columns.Items[i].Tag:=0;
  if Tag0=2 then
  begin
    column.tag:=1;
    FMData.AddIndex(Column.FieldName+'_A', Column.FieldName, []);
    FMData.IndexName:=Column.FieldName+'_A';
  end else
  begin
    column.tag:=2;

    FMData.AddIndex(Column.FieldName+'_D', Column.FieldName, [ixDescending]);
    FMData.IndexName:=Column.FieldName+'_D';
  end;
  for i:=0 to 16 do
    dbgrid1.Columns.Items[i].Title.ImageIndex:=dbgrid1.Columns.Items[i].Tag-1;
  FMData.Close;
  FMData.open;
  FMData.Indexes.Update;
  FMDATA.Refresh;
end;

procedure TMainForm.FMDataFilterRecord(DataSet: TDataSet; var Accept: Boolean); //Фильтрация таблицы
var k:string;
  ii:integer;
begin

  if (FMData.FieldByName('Datetime').AsDateTime = datastart) and
  (FMData.FieldByName('Name').AsString = FIO) and
  (FMData.Fields.Fields[channel].AsFloat <= FindMax) and (FMData.Fields.Fields[channel].AsFloat >= FindMin) then
  Accept:=true
  else
  Accept:=false;

end;


function TMainForm.FMDataTranslate(Dbf: TDbf; Src, Dest: PChar; ToOem: Boolean
  ): Integer;
begin

end;



procedure TMainForm.FormCreate(Sender: TObject); //Выполнение начальных функций и задание начальных параметров
var dbf: file of byte;
 CP: byte;
 len,j: integer;
 S: string;
 flag: boolean;
begin
  FMData.Open;
  FMData.close;
  AssignFile(dbf, 'FMData.dbf');
  Reset(dbf);
  Seek(dbf, 29);
  CP := $C9;
  Write(dbf, CP);
  CloseFile(dbf);
  CancelPoll:=False;
  PollOn:=False;
  FMData.Open;

  FMData.FieldByName('Comment').OnGetText:=@AllText;
  FMData.FieldByName('NAME').OnGetText:=@AllText;
  FMData.FieldByName('POL').OnGetText:=@AllText;

  FMData.RegenerateIndexes;
  FMData.IndexName:='FMData.mdx';

  ChMax[0]:=25;ChMin[0]:=-20;        //Кан1
  ChMax[1]:=770;ChMin[1]:=740;       //Кан2
  ChMax[2]:=100;ChMin[2]:=50;        //Кан3
  ChMax[3]:=1;ChMin[3]:=0;           //Кан4
  ChMax[4]:=3;ChMin[4]:=1;           //М5
  ChMax[5]:=1.2;ChMin[5]:=0;         //Д5
  ChMax[6]:=1;ChMin[6]:=0;           //Кан6
  ChMax[7]:=0;ChMin[7]:=0;           //Кан8
  ChMax[8]:=1000;ChMin[8]:=400;      //Кан18
  ChMax[9]:=80;ChMin[9]:=40;         //Кан47
  ChMax[10]:=110;ChMin[10]:=50;      //Кан55
  ChMax[11]:=0.6;ChMin[11]:=0.4;     //Фун55
  ChMax[12]:=0;ChMin[12]:=0;         //Кан83
  FMax[0]:=25;FMin[0]:=-20;        //Кан1
  FMax[1]:=770;FMin[1]:=740;       //Кан2
  FMax[2]:=100;FMin[2]:=50;        //Кан3
  FMax[3]:=1;FMin[3]:=0;           //Кан4
  FMax[4]:=3;FMin[4]:=1;           //М5
  FMax[5]:=1.2;FMin[5]:=0;         //Д5
  FMax[6]:=1;FMin[6]:=0;           //Кан6
  FMax[7]:=0;FMin[7]:=0;           //Кан8
  FMax[8]:=1000;FMin[8]:=400;      //Кан18
  FMax[9]:=80;FMin[9]:=40;         //Кан47
  FMax[10]:=110;FMin[10]:=50;      //Кан55
  FMax[11]:=0.6;FMin[11]:=0.4;     //Фун55
  FMax[12]:=0;FMin[12]:=0;         //Кан83



  ComboBox2.Clear;
  ComboBox2.items.Clear;
  FMData.DisableControls;
  Len:=FMData.ExactRecordCount;
  ComboBox2.ItemIndex:=-1;
  if Len>0 then
  begin
    FMData.First;
    while not (FMData.EOF) do
    begin
      S:=CP1251ToUTF8(FMData.FieldByName('Name').AsString);
      flag:=true;
      for j:=0 to ComboBox2.Items.Count-1 do
      begin
        If (S=ComboBox2.Items[j]) or (Trim(S)='') then
        begin
          flag:=false;
          break;
        end;
      end;
      FMData.Next;
      if flag then ComboBox2.Items.Append(S);
    end;
  end;
  FMData.Filtered:=true;
  if ComboBox2.Items.Count>0 then ComboBox2.ItemIndex:=0;
  ComboBox2Change(ComboBox2);
  CountMeans(Sender);
end;



procedure TMainForm.Maxy1Change(Sender: TObject); //Контроль параметров поиска для макс. п2
var nom:integer;
begin
  nom:= ComboBox1.ItemIndex;
  FMax[nom]:=Maxy1.Value;
end;

procedure TMainForm.MenuAboutClick(Sender: TObject); //О программе
begin
  aboutform.show;
end;

end.

