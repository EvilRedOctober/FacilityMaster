UNIT IntrfcD;
{‚ҐабЁп ¤«п Delphi}
INTERFACE
  USES Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls;
   var GetTime:String;
  TYPE
     MASSIV = ARRAY[1..52] OF REAL;

PROCEDURE OPROS(KANAL:INTEGER; VAR XIZM:REAL; VAR WSPOM : MASSIV);
PROCEDURE UPRAV(KANAL:INTEGER; UPR :REAL; VAR WSPOM : MASSIV);

IMPLEMENTATION

PROCEDURE OPROS(KANAL:INTEGER; VAR XIZM:REAL; VAR WSPOM : MASSIV);
(* Имитатор об"екта исследования : KANAL - номер канала опроса;  *)
(* XIZM - результат измерения по каналу; WSPOM - вспомогательный *)
(* массив типа REAL  с 52 элементами. *)

var
   i,j,j1,j2 : integer ;
   t,jtim,r,r1,r2,r3,r4,delt,tau,dx : real;
   h,m,s,s100 : word;


PROCEDURE GAUSS ( R,S : REAL; VAR R1:REAL);
(* Генерация псевдослучайных нормальных чисел *)

var
  i:integer;
  r2:real;


begin
   r1:=0; for i:=1 to 12 do
   begin
     r2:=0.001*random(1000); r1:=r1+r2
   end; r1:=(r1-6)*s+r
end; (* GAUSS *)

PROCEDURE IZMER7( J : INTEGER);
(* Генерация независимых переменных *)
var
   r,rr : real;



begin
   GetTime:=formatdatetime('tt',time);
   {DecodeTime(GetTime,h,m,s,s100); rr:=s; r:=m;} rr:=+60*r; r:=h;
   //rr:=strtofloat(copy(GetTime,7,2))+strtofloat(copy(GetTime,4,2))*60+3600*strtofloat(copy(GetTime,1,2));
   rr:=random(80000)+6400;
   wspom[11]:=rr;  r:=0.001*random(900);

   if j=1 then wspom[1]:=-20+50*r else
   if j=2 then wspom[2]:=740+30*r else
   if j=3 then wspom[3]:=50+50*r  else
   if j=4 then
   begin
      wspom[4]:=0; if r>0.5 then wspom[4]:=1
   end else
   if j=5 then
   begin
      wspom[5]:=2; if r<=0.33333 then wspom[5]:=1 else
      if r>=0.666667 then wspom[5]:=3
   end else
   if j=6 then
   begin
      wspom[6]:=0; if r>0.5 then wspom[6]:=1
   end
end; (* IZMER7 *)

begin
   (* Начало основной части процедуры OPROS *)

   if kanal>=1 then
   begin

      if kanal <=10 then
      begin
  (* Независимые переменные *)
         wspom[11]:=0; if kanal <= 6 then IZMER7(kanal);
         r:=wspom[kanal]
      end else if kanal <= 70 then
      begin
         r4:=0; r:=1; GAUSS(r4,r,r1);
         if kanal <= 25 then
         begin
   (* Лин. регресс. модель - пассивный эксперимент *)
            j1:=kanal-11; j2:=kanal-18;
            r:=30+j1*5-(5-j2*0.14)*wspom[1]+(0.6-j1*0.014)*wspom[2]+
               (350-j1*10)*wspom[4]-(3-j2*0.07)*wspom[1]*wspom[4]+
               (0.05+j1*0.001)*wspom[1]*wspom[3]+0.01*wspom[3]*wspom[3];
            r1:=5*r1
         end else if kanal <= 40 then
         begin
   (* Лин. регресс. модель - активно-пассивный эксперимент *)
            j1:=kanal-26; j2:=kanal-33;
            r:=100+j1*20+(7+j1*0.5)*wspom[1]+(0.6+j2*0.01)*wspom[2]-
               (15+j1*2)*wspom[5]+(4+j2*0.1)*wspom[7]+
               (40+j1*2)*wspom[8]+(0.01+j2*0.0005)*wspom[1]*wspom[2]-
               (0.5+j2*0.01)*wspom[7]*wspom[8]+(0.7-j1*0.02)*wspom[5]*
               wspom[7]-(0.15-j1*0.01)*wspom[1]*wspom[1]+(30-j1*1.4)*
               wspom[8]*wspom[8] ;
            r1:=5*r1
         end else if kanal <= 55 then
         begin
   (* Нелин. модель - пассивный эксперимент *)
            r:=30+(kanal-41)*3+(10+(kanal-48)*0.5)*exp((0.055-
               (kanal-41)*0.001)*wspom[1])-(2+(kanal-50)*0.1)*
               exp(0.07*wspom[1]-(0.0001+(kanal-46)*0.00002)*wspom[2]+
               (1+(kanal-48)*0.06)*wspom[4]);
            r1:=1.5*r1
         end else if kanal <= 70 then
         begin
   (* Нелинейная модель - активно-пассивный эксперимент *)
            j1:=kanal-56; j2:=kanal-63;
            r:=60+(kanal-65)*3+(-5.5+j1*0.1)*exp((0.05+j1*0.001)*
               wspom[1]+(-0.1+j2*0.006)*wspom[9])+(9-j1*0.2)*
               exp((-0.55+j1*0.01)*wspom[10]+(0.7+j2*0.03)*wspom[5]);
            r1:=1.5*r1
         end;
         r:=r+r1
      end else if kanal <= 85 then
      begin
   (* Динамическая модель - инерционное звено *)
       //GetTime:=formatdatetime('tt',time);
       //jtim:=strtofloat(copy(GetTime,7,2))+strtofloat(copy(GetTime,4,2))*60+3600*strtofloat(copy(GetTime,1,2));
       jtim:=random(80000)+6400;
         {decodetime(GetTime,h,m,s,s100); jtim:=s; r4:=m; jtim:=jtim+60*r4;
         r4:=h; jtim:=jtim+3600*r4;} tau:=15-(kanal-71)*0.4;
         r:=0; r3:=0; for i:=20 downto 1 do
         begin
            r2:=jtim-wspom[i+12]; r2:=r2/tau;
            if r2<=50 then r:=r+(wspom[i+32]-r3)*(1-exp(-r2));
            r3:=wspom[i+32]
         end; r:=(20+(kanal-78)*2)*r;

      end else
      begin
         j1:=kanal-100; r1:=j1*0.1; r4:=1; GAUSS(r1,r4,r)
      end
   end ; xizm:=r; wspom[12]:=0.001*random(1000)+0.001;
end;

PROCEDURE UPRAV(KANAL:INTEGER; UPR :REAL; VAR WSPOM : MASSIV);
(* Имитация подачи управл. воздействия на об'ект *)
type
   mm = array[1..4] of real;

var
   i,i1,j1 : integer;
   t,jtim,r: real;
   umin,umax : mm;
   h,m,s,s100 : word;

begin
   umin[1]:=-70; umax[1]:=-umin[1];
   umin[2]:=-3;  umax[2]:=-umin[2];
   umin[3]:=-5;  umax[3]:=-umin[3];
   umin[4]:=-2.5;umax[4]:=-umin[4];
   if ((kanal < 7) or (kanal > 10)) then
   writeln(' ОШИБКА UPRAV : канал N',kanal:7,' упр.возд.=',upr:10:3)
   else
   begin
      j1:=kanal-6; if upr>umax[j1] then upr:=umax[j1]
      else if upr<umin[j1] then upr:=umin[j1];
      wspom[kanal]:=upr;
      if kanal=10 then if abs(wspom[33]-upr)>=0.01 then
      begin
GetTime:=formatdatetime('tt',time);
       //jtim:=strtofloat(copy(GetTime,7,2))+strtofloat(copy(GetTime,4,2))*60+3600*strtofloat(copy(GetTime,1,2));
       jtim:=random(80000)+6400;
         {Decodetime(GetTime,h,m,s,s100); jtim:=s; r:=m; jtim:=jtim+60*r;
         r:=h; jtim:=jtim+3600*r;}
         for i:=2 to 20 do
         begin
            i1:=22-i; wspom[i1+32]:=wspom[i1+31];
            wspom[i1+12]:=wspom[i1+11]
         end;
         wspom[33]:=upr; wspom[13]:=jtim;

      end
   end
end;
END.
