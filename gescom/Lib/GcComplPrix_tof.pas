{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 19/11/2002
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : GCCOMPLPRIX ()
Mots clefs ... : TOF;GCCOMPLPRIX
*****************************************************************}
Unit GcComplPrix_tof;

Interface

Uses StdCtrls,Controls,Classes,M3FP,
{$IFDEF EAGLCLIENT}
     MaineAgl,
{$ELSE}
     db,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}Fe_Main,
{$ENDIF}
     forms,sysutils,ComCtrls,HCtrls,HEnt1,HMsgBox,UTOF,UTOB,Vierge,AGLInit ;

Function  EntreeModifPrixAchat (Action : TActionFiche; TOBL : TOB; TarifTTC : Boolean=False) : boolean ;


Type
  TOF_GCCOMPLPRIX = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (Arguments : String ) ; override ;
    procedure OnClose                  ; override ;
    public
          Action   : TActionFiche ;
    procedure CalcMarge (Saisie : Boolean; TypePrix : string) ;
  end ;

var TOBLig: TOB;

Implementation

Function EntreeModifPrixAchat (Action : TActionFiche; TOBL : TOB; TarifTTC : Boolean=False) : boolean ;
var Arg,Retour : string;
    Nb : integer ;
BEGIN
Result:=False;
if (TOBL=Nil) then exit;
if TOBL.GetValue('GL_TYPELIGNE')<>'ART' then exit;
TOBLig:=TOB.Create('LIGNE', Nil,-1) ;
TOBLig.Dupliquer(TOBL,True,True) ;
Arg:=ActionToString(Action);
Retour:=AglLanceFiche('GC','GCCOMPLPRIXACHAT','','',Arg);
if Retour='VALIDE' then
    BEGIN
    if Action<>taconsult then
       Begin
       TOBL.Dupliquer(TOBLig,True,True) ;
       Result:=True;
       End else if TOBL.GetValue('GL_VIVANTE')='-' then
       BEGIN
       BeginTrans ;
       Nb:=ExecuteSQL('UPDATE LIGNE SET GL_DPR='+StrfPoint(TOBLig.GetValue('GL_DPR'))+ 'WHERE GL_NATUREPIECEG="'+TOBLig.GetValue('GL_NATUREPIECEG')+'" '
                     +'AND GL_SOUCHE="'+TOBLig.GetValue('GL_SOUCHE')+'" AND GL_NUMERO='+IntToStr(TOBLig.GetValue('GL_NUMERO'))+' '
                     +'AND GL_INDICEG='+IntToStr(TOBLig.GetValue('GL_INDICEG'))+' AND GL_NUMLIGNE='+IntToStr(TOBLig.GetValue('GL_NUMLIGNE'))) ;
       if Nb<>1 then RollBack else CommitTrans ; 
       END ;
    END;
TOBLig.Free;
END;

{==============================================================================================}
{================================= Evenement de la TOF ========================================}
{==============================================================================================}
procedure TOF_GCCOMPLPRIX.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_GCCOMPLPRIX.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_GCCOMPLPRIX.OnUpdate ;
begin
  Inherited ;
TOBLig.GetEcran(Ecran);
TFVierge(Ecran).Retour:='VALIDE' ;
end ;

procedure TOF_GCCOMPLPRIX.OnLoad ;
begin
  Inherited ;
TOBLig.PutEcran(Ecran);
CalcMarge(False,'DPA');
CalcMarge(False,'DPR');
CalcMarge(False,'PMAP');
CalcMarge(False,'PMRP');
end ;

procedure TOF_GCCOMPLPRIX.OnArgument (Arguments : String ) ;
var St : string;
    i_ind : integer ;
begin
  Inherited ;
St:=Arguments ;
i_ind:=Pos('ACTION=',St) ;
if i_ind>0 then
   BEGIN
   System.Delete(St,1,i_ind+6) ;
   St:=uppercase(ReadTokenSt(St)) ;
   if St='CREATION' then BEGIN Action:=taCreat ; END ;
   if St='MODIFICATION' then BEGIN Action:=taModif ; END ;
   if St='CONSULTATION' then BEGIN Action:=taConsult ; END ;
   END ;
TFVierge(Ecran).Retour:='' ;
end ;

procedure TOF_GCCOMPLPRIX.OnClose ;
begin
  Inherited ;
end ;

/////////////////////////////////////////////////////////////////////////////
procedure TOF_GCCOMPLPRIX.CalcMarge(Saisie : Boolean; TypePrix : string) ;
Var PUHT, PCI, MontMarge, PourcMarge : double;
    St : String;
begin
PUHT:=THNumEdit(GetControl('GL_PUHTNET')).Value ;
if Saisie then
   BEGIN
   St:=GetControlText('GL_'+TypePrix);
   if St='' then PCI:=0 else PCI:=StrToFloat(St) ;
   END else PCI:=THNumEdit(GetControl('GL_'+TypePrix)).Value ;
MontMarge:=Arrondi(PUHT-PCI,2);
if PUHT > 0 then PourcMarge:=(MontMarge/PUHT)* 100 else PourcMarge:=0;
PourcMarge:=Arrondi(PourcMarge,3);
SetControlText('MONTMARGE'+TypePrix,FloatToStr(MontMarge));
SetControlText('POURMARGE'+TypePrix,FloatToStr(PourcMarge));
end ;

/////////////////////////////////////////////////////////////////////////////

procedure TOF_GCComplPrix_CalcMarge(parms:array of variant; nb: integer ) ;
var  F : TForm ;
     TOTOF  : TOF;
begin
F:=TForm(Longint(Parms[0])) ;
if (F is TFVierge) then TOTOF:=TFVierge(F).LaTOF else exit;
if (TOTOF is TOF_GCCOMPLPRIX) then TOF_GCCOMPLPRIX(TOTOF).CalcMarge(True,Parms[1]) else exit;
end;

procedure InitTOFComplPrix ();
begin
RegisterAglProc('GCComplPrix_CalcMarge', True , 1, TOF_GCComplPrix_CalcMarge);
end;

Initialization
  registerclasses ( [ TOF_GCCOMPLPRIX ] ) ; 
  InitTOFComplPrix();
end.
