{***********UNITE*************************************************
Auteur  ...... : Michel RICHAUD
Créé le ...... : 02/08/2001
Modifié le ... : 03/08/2001
Description .. : Source TOF de la TABLE : CEGIDPCI ()
Suite ........ : Met à jour le DPR sur la ligne de vente pour changer la 
Suite ........ : marge réelle
Mots clefs ... : TOF;CEGID;PCI;
*****************************************************************}
Unit UTofCegidPCI ;

Interface

Uses StdCtrls, Controls, Classes, forms, sysutils, ComCtrls,
     HCtrls, HEnt1, HMsgBox,Vierge,M3FP,
{$IFDEF EAGLCLIENT}
      MaineAgl,
{$ELSE}
      db,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}FE_main,
{$ENDIF}
      UTOF, UTOB, AGLInit;

Function  ModifCegidPCI (Action : TActionFiche; TOBL : TOB; TarifTTC : Boolean=False) : boolean ;


Type
  TOF_CEGIDPCI = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (Arguments : String ) ; override ;
    procedure OnClose                  ; override ;
    public
          Action   : TActionFiche ;
    procedure CalcMarge (Saisie : Boolean) ;
  end ;

var TOBLig: TOB;

Implementation

{***********A.G.L.***********************************************
Auteur  ...... : 
Créé le ...... : 02/08/2001
Modifié le ... :   /  /    
Description .. : Point d'entrée pour l'appel spécifique à CEGID permettant la 
Suite ........ : modification du PCI
Mots clefs ... : PCI;CEGID
*****************************************************************}
Function ModifCegidPCI (Action : TActionFiche; TOBL : TOB; TarifTTC : Boolean=False) : boolean ;
var Arg,Retour : string;
    Nb : integer ;
BEGIN
Result:=False;
if (TOBL=Nil) then exit;
if TOBL.GetValue('GL_TYPELIGNE')<>'ART' then exit;
TOBLig:=TOB.Create('LIGNE', Nil,-1) ;
TOBLig.Dupliquer(TOBL,True,True) ;
Arg:=ActionToString(taModif);
Retour:=AglLanceFiche('GC','GCCEGIDPCI','','',Arg);
if Retour='VALIDE' then
    BEGIN
    if Action <> taconsult then
        begin
        TOBL.Dupliquer(TOBLig,True,True) ;
        Result := True;
        end else if (TobL.GetValue ('GL_VIVANTE') = '-') then
        begin
        TOBL.Dupliquer(TOBLig,True,True) ;
        BeginTrans ;
        Nb:=ExecuteSQL ('UPDATE LIGNE SET GL_DPR=' + StrfPoint(TOBLig.GetValue ('GL_DPR')) +
                        ', GL_REFCOLIS="' + TobLig.GetValue ('GL_REFCOLIS') + '"' +
                        ', GL_DATELIVRAISON="' + UsDateTime (StrToDate (TobLig.GetValue ('GL_DATELIVRAISON'))) + '"' +
                        ', GL_RESSOURCE="' + TobLig.GetValue ('GL_RESSOURCE') + '"' +
                        'WHERE GL_NATUREPIECEG="' + TOBLig.GetValue ('GL_NATUREPIECEG')+'" ' +
                        'AND GL_SOUCHE="' + TOBLig.GetValue ('GL_SOUCHE') +
                        '" AND GL_NUMERO=' + IntToStr(TOBLig.GetValue ('GL_NUMERO')) + ' ' +
                        'AND GL_INDICEG=' + IntToStr(TOBLig.GetValue ('GL_INDICEG')) +
                        ' AND GL_NUMLIGNE=' + IntToStr(TOBLig.GetValue ('GL_NUMLIGNE'))) ;
        if Nb<>1 then RollBack else CommitTrans ;
        end;
    END;
TOBLig.Free;
END;

{==============================================================================================}
{================================= Evenement de la TOF ========================================}
{==============================================================================================}
procedure TOF_CEGIDPCI.OnArgument (Arguments : String ) ;
var St : string;
    i : integer ;
begin
  Inherited ;
St:=Arguments ;
i:=Pos('ACTION=',St) ;
if i>0 then
   BEGIN
   System.Delete(St,1,i+6) ;
   St:=uppercase(ReadTokenSt(St)) ;
   if St='CREATION' then BEGIN Action:=taCreat ; END ;
   if St='MODIFICATION' then BEGIN Action:=taModif ; END ;
   if St='CONSULTATION' then BEGIN Action:=taConsult ; END ;
   END ;
TFVierge(Ecran).Retour:='' ;
end ;

procedure TOF_CEGIDPCI.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_CEGIDPCI.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_CEGIDPCI.OnUpdate ;
begin
  Inherited ;
TOBLig.GetEcran(Ecran);
TFVierge(Ecran).Retour:='VALIDE' ;
end ;

procedure TOF_CEGIDPCI.OnLoad ;
begin
  Inherited ;
TOBLig.PutEcran(Ecran);
CalcMarge(False);
end ;

procedure TOF_CEGIDPCI.OnClose ;
begin
  Inherited ;
end ;

/////////////////////////////////////////////////////////////////////////////
procedure TOF_CEGIDPCI.CalcMarge(Saisie : Boolean) ;
Var PUHT, PCI, MontMarge, PourcMarge : double;
    St : String;
begin
PUHT:=THNumEdit(GetControl('GL_PUHTNET')).Value ;
if Saisie then
   BEGIN
   St:=GetControlText('GL_DPR');
   if St='' then PCI:=0 else PCI:=StrToFloat(St) ;
   END else PCI:=THNumEdit(GetControl('GL_DPR')).Value ;
MontMarge:=PUHT-PCI;
if PUHT > 0 then PourcMarge:=(MontMarge/PUHT)* 100 else PourcMarge:=0;
PourcMarge:=Arrondi(PourcMarge,3);
SetControlText('MONTANTMARGE',FloatToStr(MontMarge));
SetControlText('POURCENTMARGE',FloatToStr(PourcMarge));
end ;

/////////////////////////////////////////////////////////////////////////////

procedure TOF_CegidPCI_CalcMarge(parms:array of variant; nb: integer ) ;
var  F : TForm ;
     TOTOF  : TOF;
begin
F:=TForm(Longint(Parms[0])) ;
if (F is TFVierge) then TOTOF:=TFVierge(F).LaTOF else exit;
if (TOTOF is TOF_CEGIDPCI) then TOF_CEGIDPCI(TOTOF).CalcMarge(True) else exit;
end;

procedure InitTOFCegidPCI ();
begin
RegisterAglProc('CegidPCI_CalcMarge', True , 0, TOF_CegidPCI_CalcMarge);
end;

Initialization
  registerclasses ( [ TOF_CEGIDPCI ] ) ;
  InitTOFCegidPCI();
end.
