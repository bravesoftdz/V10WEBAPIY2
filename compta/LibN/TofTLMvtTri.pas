unit TofTLMvtTri;

interface

uses Classes, StdCtrls,
{$IFDEF EAGLCLIENT}
        MaineAGL,
{$ELSE}
     {$IFNDEF DBXPRESS}dbtables{BDE},{$ELSE}uDbxDataSet,{$ENDIF}
     FE_Main,
     QRS1,
{$ENDIF}
     comctrls,UTof, HCtrls,
     Ent1, Graphics, HTB97, spin,
     SysUtils;

function CPLanceFiche_MvtTriTL(vStRange, vStLequel, vStArgs : string) : String ;

type
    TOF_TLMVTTRI = class(TOF)
  private
    ArgRetour: string;
    procedure TriOnChange(Sender : TObject);
    procedure ControleTri;
    Procedure BFermeOnClick(Sender : TObject);

  public
    procedure OnArgument(Arguments : string); override ;
    procedure OnUpdate; override ;
    procedure OnClose; override ;
    procedure OnCancel; override ;
  end;

implementation

uses HEnt1, LicUtil, Vierge,HMsgBox ;

// ===============================================================================

function CPLanceFiche_MvtTriTL(vStRange, vStLequel, vStArgs : string) : String ;
begin
  result := AglLanceFiche('CP','TLMVTTRI',vStRange, vStLequel, vStArgs) ;
end ;

// ===============================================================================

Procedure TOF_TLMVTTRI.BFermeOnClick(Sender : TObject);
BEGIN
TFVierge(Ecran).retour:=ArgRetour;
END;

procedure TOF_TLMVTTRI.OnCancel;
BEGIN
TFVierge(Ecran).retour:=ArgRetour;
END;

procedure TOF_TLMVTTRI.OnClose ;
BEGIN
ControleTri;
END;

procedure TOF_TLMVTTRI.OnUpdate ;
BEGIN
TFVierge(Ecran).retour:=GetControlText('TRI');
END;

Procedure TOF_TLMVTTRI.ControleTri;
Var st,st1 : string;
    i : integer;
BEGIN
St:=GetControlText('TRI');
i:=0;
While St<>'' do
begin
St1:=ReadTokenSt(St);
if Pos(St1,St)>0 then i:=i+1 ;
end;
if i>0 then
  begin
  PgiBox('Des choix de tables sont dupliqués','Grand-Livre Agée');
  LastError:=1;
  end;
END;

procedure TOF_TLMVTTRI.OnArgument(Arguments : string);
Var i,j,ind : integer;
    St,St1,Q : string;
    Tri : THValComboBox;
    BFerme : TToolBarButton97;
BEGIN
inherited ;
BFerme:=TToolBarButton97(GetControl('BFerme'));
if BFerme<>nil then BFerme.OnClick:=BFermeOnClick;
ArgRetour:=Arguments;
St:=Arguments;
SetControlText('TRI',St);
ind:=0;
While St<>'' do
begin
St1:=ReadTokenSt(St);
ind:=ind+1;
if Pos('AND',Q)>0 then Q:=Q+'OR CC_CODE="'+St1+'" ' else Q:=Q+'AND (CC_CODE="'+St1+'" ' ;
end;
Q:=Q+')';
for i:=0 to 3 do
begin
Tri:=THValComboBox(GetControl('TRI'+IntToStr(i)));
if Tri<>nil then Tri.OnChange:=TRiOnChange;
if (i<=ind-1) then Tri.plus:=Q;
if (i>ind-1) then begin SetControlEnabled('TRI'+IntToStr(i),False);SetControlEnabled('TTRI'+IntToStr(i),False);end;
end;
for j:=0 to ind-1 do THValComboBox(GetControl('TRI'+IntToStr(j))).itemindex:=j;
END;

procedure TOF_TLMVTTRI.TriOnChange(Sender : TObject);
Var St : string;
    i: integer;
BEGIN
for i:=0 to 3 do
begin
if (THValComboBox(GetControl('TRI'+IntToStr(i))).enabled=True) and (THValComboBox(GetControl('TRI'+IntToStr(i))).value<>'')
  then St:=St+GetControlText('TRI'+IntToStr(i))+';';
end;
SetControlText('TRI',St);
END;

initialization
RegisterClasses([TOF_TLMVTTRI]) ;

end.

