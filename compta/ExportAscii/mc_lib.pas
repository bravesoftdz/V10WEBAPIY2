unit mc_lib;

interface

uses Hent1, Controls ;

//XMG 02/04/04 début
Const LP_CodeImprimanteExportFichier = 'EXP' ;
      TypeEtatTexte                  = 'T'   ;
      TypeEtatExportAscii            = 'F'   ;
//XMG 02/04/04 fin

//String
function sLeft(s:string;Count:integer):string;
function sRight(s:string;Count:integer):string;
Function InString ( Quoi : String ; Ou : Array of String ) : Boolean ;
Function NbCarInString(St,Car : String) : integer ;
Function gtfs (St:String;Car:string;n:integer):String;
function ptfs (St:String;Car:string;n:integer; SubSt:String) :string;
function nString(n:integer;car:string):string;
function  Postfs (St:String;Car:string;SubSt:String) :integer;
Function  FindEtReplace2 (St,StFind,StReplace : String ; Tous : Boolean) : string ;
function  RTFToString( RTF : String ) : String ;  //XMG 23/06/03
function StringToRTF ( St : String ) : String ; //01/07/03
function MC_WrapText ( Texte : String ; Longeur : integer ; sep : String ) : string ; //XMG 25/07/03 fiche S1 10308


//Booleans
Function TrueFalseSt ( TF : Boolean ) : String ;
Function TrueFalseLib ( TF : Boolean ) : String ;
Function TrueFalseBo ( TF : String ) : Boolean ;
Function ConvertTRUEFALSESt(TF : String ; FALSETRUEop : String='' ) : String ;


//Variants
function vInteger(v:variant):integer;
function vString(v:variant):String;
function vDate(v:variant):TDateTime;
function vDouble(v:variant):Double;

//Controle transaction
Function  _TRANSACTIONS ( FF : TProc ; NbMax : Integer) : TIOErr ;
procedure _BEGINTRANS ;
Procedure _ROLLBACK ;
Procedure _COMMITTRANS ;

//TToolWindow
Procedure CentreTw ( F : TWincontrol ; TW : TControl) ;

implementation

uses Sysutils, HDebug, Windows, HTB97, HRichOLE, Forms,Classes, Math; //XMG 25/07/03

//////////////////////////////////////////////////////////////////////////////
//XMG 23/06/03 début
function RTFToString( RTF : String ) : String ;
var F  : TForm ;
    Re : THRichEditOLE ;
    TS : TStringlist ;
Begin
Result:='' ;
F:=TForm.Create(Application);
TS:=Nil ;
try
   Re:=THRichEditOLE.Create(F) ;
   Re.WordWrap:=FALSE ; //XMG 23/06/03
   Re.Parent:=F ;
   TS:=TStringlist.Create ;
   TS.Text:=RTF ;
   StringstoRich(re,TS) ;
   result:=re.Lines.Text ;
  finally
   if assigned(TS) then TS.Free ;
   F.Free ;
  End ;
End ;
////////////////////////////////////////////////////////////////////////////////////
//XMG 23/06/03 fin
//XMG 01/07/03 début
function StringToRTF ( St : String ) : String ;
var F  : TForm ;
    Re : THRichEditOLE ;
Begin
Result:='' ;
F:=TForm.Create(Application);
try
   Re:=THRichEditOLE.Create(F) ;
   Re.WordWrap:=FALSE ;
   Re.Parent:=F ;
   Re.Lines.text:=St ;
   Result:=re.LinesRTF.Text ;
  finally
   F.Free ;
  End ;
End ;
//XMG 01/07/03 fin
//XMG 25/07/03 début fiche S1 10308
//////////////////////////////////////////////////////////////////////////////////
function MC_WrapText ( Texte : String ; Longeur : integer ; sep : String ) : string ;
var ll : integer ;
    St : String ;
begin
St:=trim(texte) ; Texte:='' ;
While Length(St)>0 do
  Begin
  ll:=minintvalue([pos(#13,St),pos(#10,St)]) ;
  if ll<=0 then ll:=length(St) ;
  if (ll<=0) or (ll>Longeur) then
     Begin
     Ll:=Longeur ;
     while (ll>0) and (St[ll]<>' ') do dec(ll) ;
     if LL<=0 then ll:=Longeur ;
     End else Dec(ll,ord(ll<length(St))) ;
  if trim(texte)<>'' then texte:=texte+Sep ;
  texte:=Texte+Copy(St,1,ll) ;
  Delete(St,1,ll) ;
  st:=trim(st) ;
  End ;
result:=Texte ;
end ;
//XMG 25/07/03 fin Fiche S1 10308
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
function  sRight(s:string;Count:integer):string;
var   L:integer;
begin
L:=length(s);
if (Count>L) or (s='') then
   begin
   result:=s;
   exit;
   end;
result:=Copy(s, L-Count+1, Count);
end;
//////////////////////////////////////////////////////////////////////////////
function sLeft(s:string;Count:integer):string;
var   L:integer;
begin
L:=length(s);
if (L<=Count) or (s='') then
   begin
   result:=s;
   exit;
   end;
result:=Copy(s, 1, Count);
end;
//////////////////////////////////////////////////////////////////////////////
Function InString ( Quoi : String ; Ou : Array of String ) : Boolean ;
Var st : String ;
    i  : integer ;
Begin
Quoi:=';'+Uppercase(trim(Quoi))+';' ;
St:='' ;
for i:=low(Ou) to high(ou) do
  Begin
  if trim(St)='' then St:=';' ;
  St:=St+Uppercase(Trim(ou[i]))+';' ;
  End ;
Result:=(pos(Quoi,St)>0) ;
End ;
//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Function  NbCarInString(St,Car : String) : integer ;
Var i,n : Integer ;
BEGIN
Result:=0 ; n:=0 ; i:=0 ;
repeat
   i:=i+n ;
   n:=pos(car,Copy(st,1+i,length(St))) ;
   if n>0 then inc(Result) else break ;
   until n=0 ;
END ;
///////////////////////////////////////////////////////////////////////////////
procedure TrouveDebEtLongueurtfs ( var Deb, Longueur : Integer ; St,car : String ; n : integer ) ;
var t,nn,i : integer ;
Begin
Deb:=0 ; Longueur:=0 ;
t:=0; nn:=0 ; i:=0 ;
while (t<n) and (n<>1) do
   begin
   i:=i+nn ;
   nn:=pos(car,Copy(st,1+i,length(St))) ;
   if nn=0 then break else inc(t) ;
   if t=n-1 then begin i:=i+nn ;break ;end ;
   end ;
if n=t+1 then
   begin
   Deb:=i+1+(Length(Car)-1)*ord(N<>1) ;
   nn:=pos(car,Copy(st,deb,length(St)));
   if nn>0 then Longueur:=nn-1
           else Longueur:=length(St)-Deb+1 ;
   end ;
End ;
//////////////////////////////////////////////////////////////////////////////////
Function AssureString(St,car : String ; n : integer ) : String ;
var t : integer ;
Begin
Result:=St ;
if sRight(St,length(Car))<>Car then St:=St+Car ;
t:=nbCarInString(St,Car) ; if t<n then St:=St+NString(n-t,Car) ;
End ;
//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Function  gtfs (St:String;Car:string;n:integer):String;
Var i,nn:integer;
begin;
St:=AssureString(St,Car,n) ;
Result:='' ;
TrouveDebEtLongueurtfs(i,nn,St,car,n) ;
if (i>0) and (nn>0) then Result:=Copy(St,i,nn) ;
end;
//////////////////////////////////////////////////////////////////////////////////
function  ptfs (St:String;Car:string;n:integer; SubSt:String) :string;
var d,l: integer ;
begin
St:=AssureString(St,Car,n) ;
TrouveDebEtLongueurtfs(d,l,St,car,n) ;
Result:=Copy(St,1,d-1)+Subst+copy(St,d+l,length(St)-(d+l)+1) ;
end ;

//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
function  Postfs (St:String;Car:string;SubSt:String) :integer;
var s,w:string ;
	 i,Cnt:integer ;
begin
result:=0;
s:='';
SubSt:=uppercase(SubSt) ;
Cnt:=NbCarInstring(St,Car);
for i:=1 to Cnt do
    begin
    w:=gtfs(st,car,i);
    if SubSt =uppercase(w) then begin result:=i; exit ; end;
    end ;
end ;
//////////////////////////////////////////////////////////////////////////////////
Function FindEtReplace2 (St,StFind,StReplace : String ; Tous : Boolean) : string ;
Var i,o : integer ;
BEGIN
Result:=St ;
if (StFind='') or (St='') then exit ;
if StFind=StReplace then exit ;
i:=Pos(StFind,St) ; o:=0 ;
While i>0 do
   BEGIN
   Delete(St,i+o,Length(StFind)) ;
   Insert(StReplace,St,i+o) ;
   o:=o+i+length(StReplace)-1 ;
   i:=Pos(StFind,Copy(St,o+1,Length(St))) ;
   if Not Tous then i:=0 ;
   END ;
Result:=St ;
END ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Const FALSETRUE     = '-X' ;
Const FALSETRUELib  = 'NonOui' ;
Const FALSETRUEopSt = 'NO' ;

Function TrueFalseSt ( TF : Boolean ) : String ;
Begin
Result:=Copy(FALSETRUE,1+ord(TF),1) ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Function TrueFalseLib ( TF : Boolean ) : String ;
Begin
Result:=TraduireMemoire(trim(Copy(FALSETRUELib,1+3*ord(TF),3))) ;
End ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Function TrueFalseBo ( TF : String ) : Boolean ;
Begin
Result:=(uppercase(Trim(TF))=copy(FALSETRUE,2,1)) ;
End ;
/////////////////////////////////////////////////////////////////////////////////////////////
Function ConvertTRUEFALSESt(TF : String ; FALSETRUEop : String='' ) : String ;
Begin
if (Trim(FALSETRUEop)='') or (Length(FALSETRUEop) mod 2<>0) then FALSETRUEop:=FALSETRUEopSt ;
Result:=Copy(FALSETRUEop,1+ord(TRUEFALSEbo(TF)),1) ;
End ;
//////////////////////////////////////////////////////////////////////////////////
function nString(n:integer;car:string):string;
var i : integer ;
begin
result:='' ;
for i:=1 to n do Result:=Result+Car ;
end;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
function vInteger(v:variant):integer;
Var St : String ;
begin
if VarType(v) in [varSmallint,varInteger,varSingle,varDouble,varCurrency]
   then result:=VarAsType(V,Varinteger)  else
    BEGIN
    St:=vstring(v) ;
    if IsNumeric(St) then result:=Valeuri(St) else result:=0 ;
    END ;
end;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
function vString(v:variant):String;
var st:string;
begin
if VarIsNull(v) then begin result:='' ; exit ; end ;
st:=VarAsType(v,VarOleStr);
if St=#0 then St:='';
Result:=St;
end;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
function vDate(v:variant):TDateTime;
var d : TDateTime;
begin
try
   if VarIsNull(v) then begin result:=iDate1900 ; exit ; end ;
   d:=VarAsType(V,VarDate);
   if (d<iDate1900) or (d>idate2099) then Result:=iDate1900
                                     else Result:=d;
except
   result:=iDate1900;
   end ;
end;

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
function vDouble(v:variant):Double;
var st:string ;
begin
if VarType(v) in [varSmallint,varInteger,varSingle,varDouble,varCurrency]
   then result:=VarAsType(V,VarDouble)  else
    BEGIN
    St:=vstring(v) ;
    if IsNumeric(St) then result:=Valeur(St) else result:=0 ;
    END ;
end;
//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Function  _TRANSACTIONS ( FF : TProc ; NbMax : Integer) : TIOErr ;
var s:string ;
begin
if V_PGI.Debug then s:= ('< Transaction lancée ('+intToStr(V_pgi.NbTransaction)+' transaction)' ) ;
result := TRANSACTIONS ( FF , NbMax) ;
if V_PGI.Debug then Debug(s+'. Transaction exécutée ('+intToStr(V_pgi.NbTransaction)+' transaction )' +' >') ;

end ;

//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PROCEDURE _BEGINTRANS ;
var s:string ;
begin
if V_PGI.Debug then s:=('< BeginTrans lancé ('+intToStr(V_pgi.NbTransaction)+' transaction)' ) ;
BEGINTRANS ;
if V_PGI.Debug then Debug (s+'. BeginTrans exécuté ('+intToStr(V_pgi.NbTransaction)+' transaction )' +' >' ) ;
end ;

//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PROCEDURE _ROLLBACK ;
var s:string ;
begin
if V_PGI.Debug then s:= ('< Rollback lancé ('+intToStr(V_pgi.NbTransaction)+' transaction )' ) ;
ROLLBACK ;
if V_PGI.Debug then Debug (s+'. Rollback exécuté ('+intToStr(V_pgi.NbTransaction)+' transaction )'  +' >') ;
end ;

//////////////////////////////////////////////////////////////////////////////////////////////////////////////////
PROCEDURE _COMMITTRANS ;
var s:string ;
begin
if V_PGI.Debug then s:= ('< CommitTrans lancé ('+intToStr(V_pgi.NbTransaction)+' transaction )' ) ;
COMMITTRANS ;
if V_PGI.Debug then Debug (s+'. CommitTrans exécuté ('+intToStr(V_pgi.NbTransaction)+' transaction )' +' >') ;
end ;
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
Procedure CentreTw ( F : TWincontrol ; TW : TControl) ;
var P : TPoint ;
Begin
P.Y:=(F.ClientHeight - TW.Height) div 2 ;
P.X:=(F.ClientWidth - TW.Width) div 2 ;
if TW is TToolWindow97 then p:=F.ClienttoScreen(P) ;
TW.Top:=P.Y ;
TW.Left:=P.X ;
End ;





end.
