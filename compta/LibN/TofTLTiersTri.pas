unit TofTLTiersTri;

interface

uses Classes, StdCtrls,
{$IFDEF EAGLCLIENT}
    maineAGL,
{$ELSE}
    FE_Main,
{$ENDIF}
     ComCtrls, UTof, HCtrls, Ent1, SysUtils, Vierge, HMsgBox;

type
    TOF_TLTIERSTRI = class(TOF)
  private
    ArgRetour: string;
    procedure TriOnChange(Sender : TObject);
  public
    procedure OnArgument(Arguments : string); override ;
    procedure OnUpdate; override ;
  end;

{JP 25/08/04 : FQ 13401, 13452, 13671}
function CPLanceFiche_TLTIERSTRI(Arg : string) : string;


implementation


{---------------------------------------------------------------------------------------}
function CPLanceFiche_TLTIERSTRI(Arg : string) : string;
{---------------------------------------------------------------------------------------}
begin
  Result := AglLanceFiche('CP', 'TLTIERSTRI', '', '', Arg);
end;

procedure TOF_TLTIERSTRI.OnUpdate ;
Var st,st1 : string;
    i : integer;
begin
  inherited ;
  St:=GetControlText('TRI');
  i:=0;
  While St<>'' do
    begin
    St1:=ReadTokenSt(St);
    if Pos(St1,St)>0 then
      i:=i+1 ;
    end;
  if i>0 then
    begin
    PgiBox('Des choix de tables sont dupliqués','Balance Agée');
    LastError:=1;
    Exit ;
    end;
  TFVierge(Ecran).retour := GetControlText('TRI');
end;

procedure TOF_TLTIERSTRI.OnArgument(Arguments : string);
Var i,j,ind : integer;
    St,St1,Q : string;
    Tri : THValComboBox;
begin
  inherited ;
  ArgRetour:=Arguments;
  TFVierge(Ecran).retour:=ArgRetour;
  St:=Arguments;
  SetControlText('TRI',St);
  ind:=0;
  While St<>'' do
    begin
    St1:=ReadTokenSt(St);
    ind:=ind+1;
    if Pos('AND',Q)>0
      then Q:=Q+'OR CC_CODE="'+St1+'" '
      else Q:=Q+'AND (CC_CODE="'+St1+'" ' ;
    end;
  Q:=Q+')';
  for i:=0 to 9 do
    begin
    Tri := THValComboBox(GetControl('TRI'+IntToStr(i)));
    Tri.OnChange := TRiOnChange;
    if (i<=ind-1) then
      Tri.plus := Q ;
    if (i>ind-1) then
      begin
      SetControlEnabled('TRI'+IntToStr(i),False);
      SetControlEnabled('TTRI'+IntToStr(i),False);
      end;
    end;
  for j:=0 to ind-1 do
    THValComboBox(GetControl('TRI'+IntToStr(j))).itemindex := j;
end;

procedure TOF_TLTIERSTRI.TriOnChange(Sender : TObject);
Var St : string;
    i: integer;
begin
  St := '' ;
  for i:=0 to 9 do
    if ( GetControlEnabled('TRI'+IntToStr(i)) and
         ( GetControlText('TRI'+IntToStr(i)) <> '' ) ) then
      St := St + GetControlText('TRI'+IntToStr(i)) + ';' ;
  SetControlText('TRI',St);
end;

initialization
RegisterClasses([TOF_TLTIERSTRI]) ;

end.

