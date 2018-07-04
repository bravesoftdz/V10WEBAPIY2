{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 24/04/2003
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : QRIMMOAMP ()
Mots clefs ... : TOF;QRIMMOAMP
*****************************************************************}
Unit QR_IMMOAMP_TOF ;

Interface

Uses
  Windows, Messages, Classes, Graphics, Controls, Dialogs,
  ExtCtrls, Hqry, Menus, HSysMenu,
  Grids, DBGrids, HDB,  StdCtrls, HTB97, Buttons,
  HPanel,UiUtil, Mask, ParamDat, ImEnt, LookUp,
  HPdfviewer,UTOB,ImPlan, ImPlanInfo, HStatus, Spin ,
  LicUtil,forms, sysutils,ComCtrls,HCtrls, HEnt1,
  HMsgBox,UTOF
{$IFDEF EAGLCLIENT}
  ,eQRS1
{$ELSE}
  ,db,dbtables, QRS1
{$ENDIF}
 ;

Type
  TOF_QRIMMOAMP = Class (TOF)
Private
    FEtat,table1,Table2,Table3,natimmo : THValComboBox;
    HmTrad : THSystemMenu ;
    PZlibre, PZLibreS1 : TTabSheet ;
    FTypeEdition : string;
    bArretDotation : boolean;
    procedure initialiseforme;
    procedure COMPTEREFElipsisClick(Sender: TObject);
    procedure UpdateZones;
    procedure IMMOElipsisClick(Sender: TObject);
    procedure HMTradBeforeTraduc(Sender: TObject);
    function  ControleZones : boolean;
    procedure AssignHelpContext;
    procedure FEtatChange(Sender: TObject);
    procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure TABLELIBRE1Change(Sender: TObject);
    procedure I_NATUREIMMOChange(Sender: TObject);
Public
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  end ;

Implementation
uses
{$IFDEF SERIE1}
{$ELSE}
ImOle,
{$ENDIF}
ImEdCalc,s1Util ;

procedure TOF_QRIMMOAMP.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_QRIMMOAMP.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_QRIMMOAMP.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_QRIMMOAMP.OnLoad ;
var DateCDeb,DateCFin,DateADeb,DateAFin: string ;
begin
Hent1.EnableControls(Ecran,false,true) ;
SetControlText('XX_WHERE','');
DateCDeb:=USDateTime(StrToDate(GetControlText('I_DATECESSION'))) ;
DateCFin:=USDateTime(StrToDate(GetControlText('I_DATECESSION'))) ;
DateADeb:=USDateTime(StrToDate(GetControlText('I_DATECESSION'))) ;
DateAFin:=USDateTime(StrToDate(GetControlText('I_DATECESSION'))) ;

SetControlText('I_DATEPIECEA',GetControlText('_I_DATEPIECEA'));
SetControlText('I_DATEPIECEA_',GetControlText('_I_DATEPIECEA_'));

if ControleZones then
  begin
  VHImmo^.PlanInfo.Free ; VHImmo^.PlanInfo:=nil ;
  end;
Hent1.EnableControls(Ecran,true,true) ;
VHImmo^.PlanInfo.Free ; VHImmo^.PlanInfo:=nil ;
  Inherited ;

end ;

procedure TOF_QRIMMOAMP.OnArgument (S : String ) ;
begin
  Inherited ;
{$IFNDEF IMP}
{$IFDEF SERIE1}
{$ELSE}
ProcZoomEdt:=ImZoomEdtEtatImmo ;
ProcCalcEdt:=CalcOLEEtatImmo ;
{$ENDIF}
{$ENDIF}

(*bEuro.Visible:=false ;   ??? YCP
bProg.Visible:=false ;*)

(*bEuro.visible:=false ;   ??? YCP
bProg.visible:=false ;*)
//Établie la fiche
updatecaption(Ecran) ;

AssignHelpContext;

TFQRS1(Ecran).ParamEtat:=InitParamEtat ;

UpdateZones;
InitialiseForme ;

end ;

procedure TOF_QRIMMOAMP.OnClose ;
begin
  Inherited ;
{$IFDEF SERIE1}
{$ELSE}
ProcZoomEdt:=ImZoomEdtEtat ;
ProcCalcEdt:=ImCalcOLEEtat ;
{$ENDIF}
end ;

procedure TOF_QRIMMOAMP.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_QRIMMOAMP.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_QRIMMOAMP.UpdateZones;
begin
  fTypeEdition := GetControlText('FEtat');
  bArretDotation :=(fTypeEdition='DOT') or (fTypeEdition='DFI') or (fTypeEdition='DDE') or (fTypeEdition='DEC') or (fTypeEdition='DCB');
  SetControlVisible('ARRETDOTATION',bArretDotation);
  SetControlVisible('LBLArretDotation',bArretDotation);
  if GetControlVisible('ARRETDOTATION') then SetControlText('ARRETDOTATION',DateToStr(VHImmo^.Encours.Fin)) ;
  {$IFDEF SERIE1}
//???  FListe.Visible:=true ; //ycp
  {$ELSE}
//???  FListe.Visible:=(fTypeEdition <> 'ETI') and (fTypeEdition <> 'FIM');
  {$ENDIF}
  if bArretDotation then SetControlText('_I_DATEPIECEA_',DateToStr(VHImmo^.Encours.Fin)) ;
  SetControlVisible('TI_DateCession',false);
  SetControlVisible('I_DATECESSION',false);
  SetControlVisible('TI_DateCession_',false);
  SetControlVisible('I_DATECESSION_',false);

end;

procedure TOF_QRIMMOAMP.IMMOElipsisClick(Sender: TObject);
begin
  inherited;
  LookUpList (TControl (Sender),'Immobilisation','IMMO','I_IMMO','I_LIBELLE','','I_IMMO',True,0) ;
end;


procedure TOF_QRIMMOAMP.COMPTEREFElipsisClick(Sender: TObject);
var CpteInf,CpteSup : string;
    sWhere : string;
begin
  inherited;
  if GetControlText('I_NATUREIMMO')='CB' then begin CpteSup := VHImmo^.CpteCBSup; CpteInf := VHImmo^.CpteCBInf;end
  else if GetControlText('I_NATUREIMMO')='LOC' then begin CpteSup := VHImmo^.CpteLocSup; CpteInf := VHImmo^.CpteLocInf; end
  else if GetControlText('I_NATUREIMMO')='PRO' then begin CpteSup := VHImmo^.CpteImmoSup;CpteInf := VHImmo^.CpteImmoInf;end
  else if GetControlText('I_NATUREIMMO')='FI'  then begin CpteSup := VHImmo^.CpteFinSup; CpteInf := VHImmo^.CpteFinInf; end
  else begin CpteSup := VHImmo^.CpteLocSup; CpteInf :=VHImmo^.CpteImmoInf;end;
  sWhere:='G_GENERAL<="'+CpteSup+'" AND G_GENERAL>="'+CpteInf+'"';
  LookUpList (TControl (Sender),'Généraux','GENERAUX','G_GENERAL','G_LIBELLE',sWhere,'G_GENERAL',True,0) ;
end;

procedure TOF_QRIMMOAMP.HMTradBeforeTraduc(Sender: TObject);
var Okok: boolean ;
begin
  inherited;
  {$IFDEF SERIE1}
  ImLibellesTableLibre(PzLibreS1,'TT_TABLELIBREIMMO','','I') ;
  Okok:=false ;
  {$ELSE}
  ImLibellesTableLibre(PzLibre,'TI_TABLE','I_TABLE','I') ;
  Okok:=true ;
  {$ENDIF}
  PzLibre.TabVisible:=Okok ;
  PzLibreS1.TabVisible:=not Okok ;
end;

function TOF_QRIMMOAMP.ControleZones : boolean;
begin
  result:=false ;
  if (not IsValidDate(GetControlText('_I_DATEPIECEA'))) or (not IsValidDate(GetControlText('_I_DATEPIECEA'))) then PgiBox('Veuillez saisir correctement les dates','Immobilisations')
  else if (bArretDotation) then
  begin
    if (not IsValidDate(GetControlText('ARRETDOTATION'))) then PgiBox('Veuillez saisir correctement les dates','Immobilisations')
    else if (StrToDate(GetControlText('ARRETDOTATION'))<VHImmo^.Encours.Deb) or (StrToDate(GetControlText('ARRETDOTATION'))>VHImmo^.Encours.Fin) then
    begin
      PgiBox('La date doit faire partie de l''exercice en cours','Immobilisations') ;
      SetControlText('ARRETDOTATION',DateToStr(VHImmo^.Encours.Fin));
    end
    else
      result:=true ;
  end
  else
    result:=true;
end;

procedure TOF_QRIMMOAMP.AssignHelpContext;
begin
{$IFDEF SERIE1}
  if (GetControlText('FEtat')='PRE') then Ecran.HelpContext:=553100
  else if (GetControlText('FEtat')='PRF') then Ecran.HelpContext:=553200;
{$ELSE}
  if (GetControlText('FEtat')='PRE') then Ecran.HelpContext := 2531000
  else if (GetControlText('FEtat')='PRF') then Ecran.HelpContext := 2532000;
{$ENDIF}

End ;

procedure TOF_QRIMMOAMP.FEtatChange(Sender: TObject);
begin
  AssignHelpContext;
  UpdateZones;
end;

Procedure TOF_QRIMMOAMP.InitialiseForme ;
var
   Immo,cpt      : THCritMaskEdit ;
begin
{$IFDEF SERIE1}
SetControlVisible('I_ETABLISSEMENT',False);
SetControlVisible('tI_ETABLISSEMENT',False);
{$ENDIF}
{$IFNDEF SERIE1}
// Paramétrage des états possible si entrée avec le mot de passe du jour
SetControlVisible('bParamEtat',True);
{$ENDIF}

SetControlText('I_compteref',VHImmo^.CpteImmoInf);
SetControlText('I_compteref_',VHImmo^.CpteLocSup);
SetControlText('I_DATEPIECEA',StDate1900);
SetControlText('I_DATEPIECEA_',StDate2099);
SetControlText('I_DATECESSION',StDate1900) ;
SetControlText('I_DATECESSION_',StDate2099) ;

SetControlText('I_QUALIFIMMO','R') ;
immo:=THCritMaskEdit(getcontrol('I_IMMO')) ;
if Assigned(immo) then
  Begin
  immo.OnElipsisClick:=ImmoElipsisClick ;
  End ;
immo:=THCritMaskEdit(getcontrol('I_IMMO_')) ;
if Assigned(immo) then
  Begin
  immo.OnElipsisClick:=ImmoElipsisClick ;
  End ;
cpt:=THCritMaskEdit(getcontrol('I_compteref')) ;
if Assigned(cpt) then
  Begin
  cpt.OnElipsisClick:=COMPTEREFElipsisClick ;
  End ;
cpt:=THCritMaskEdit(getcontrol('I_compteref_')) ;
if Assigned(cpt) then
  Begin
  cpt.OnElipsisClick:=COMPTEREFElipsisClick ;
  End ;
HmTrad:=THSystemMenu(GetControl('HmTrad')) ;
if (HmTrad<>nil) and (not assigned(HmTrad.OnBeforeTraduc)) then HmTrad.OnBeforeTraduc:=HMTradBeforeTraduc ;

PZlibre:=TTabSheet(GetControl('PZLibre')) ;
PZLibreS1:=TTabSheet(GetControl('PZLibreS1')) ;

FEtat:=THValComboBox(GetControl('Fetat')) ;
If (FEtat<>nil) and (not assigned(Fetat.OnChange)) then Fetat.onChange:=FetatChange ;

Table1:=THValComboBox(GetControl('TABLELIBRE1')) ;
If (Table1<>nil) and (not assigned(Table1.OnChange)) then Table1.onChange:=TableLibre1Change ;

Table2:=THValComboBox(GetControl('TABLELIBRE2')) ;
If (Table2<>nil) and (not assigned(Table2.OnChange)) then Table2.onChange:=TableLibre1Change ;

Table3:=THValComboBox(GetControl('TABLELIBRE3')) ;
If (Table3<>nil) and (not assigned(Table3.OnChange)) then Table3.onChange:=TableLibre1Change ;

NatImmo:=THValComboBox(GetControl('I_NATUREIMMO')) ;
If (NatImmo<>nil) and (not assigned(NatImmo.OnChange)) then NatImmo.onChange:=I_NATUREIMMOChange ;

if (Assigned(Ecran.OnKeyDown)) then Ecran.OnKeyDown:=FormKeyDown;

End ;

procedure TOF_QRIMMOAMP.FormKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
Case Key of
  VK_F9,VK_F10 : BEGIN  if not ControleZones then exit; END ;
  END ;
end;

procedure TOF_QRIMMOAMP.TABLELIBRE1Change(Sender: TObject);
begin
  inherited;
  if Sender=GetControl('TABLELIBRE1') then SetControlText('I_TABLE0',GetControlText('TABLELIBRE1'))
  else if Sender=GetControl('TABLELIBRE2') then SetControlText('I_TABLE1',GetControlText('TABLELIBRE2'))
  else if Sender=GetControl('TABLELIBRE3') then SetControlText('I_TABLE2',GetControlText('TABLELIBRE3'));
end;

procedure TOF_QRIMMOAMP.I_NATUREIMMOChange(Sender: TObject);
begin
  inherited;
  if GetControlText('I_NATUREIMMO')='CB' then begin SetControlText('I_compteref_',VHImmo^.CpteCBSup); SetControlText('I_compteref',VHImmo^.CpteCBInf); end
  else if GetControlText('I_NATUREIMMO')='LOC' then begin SetControlText('I_compteref_',VHImmo^.CpteLocSup); SetControlText('I_compteref',VHImmo^.CpteLocInf);  end
  else if GetControlText('I_NATUREIMMO')='PRO' then begin SetControlText('I_compteref_',VHImmo^.CpteImmoSup);SetControlText('I_compteref',VHImmo^.CpteImmoInf); end
  else if GetControlText('I_NATUREIMMO')='FI'  then begin SetControlText('I_compteref_',VHImmo^.CpteFinSup); SetControlText('I_compteref',VHImmo^.CpteFinInf);  end
  else begin SetControlText('I_compteref_',VHImmo^.CpteLocSup); SetControlText('I_compteref',VHImmo^.CpteImmoInf); end;
end;


end.
