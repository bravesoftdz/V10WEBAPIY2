{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 21/02/2007
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : PGMULPARAMHISTO ()
Mots clefs ... : TOF;PGMULPARAMHISTO
*****************************************************************
PT1  15/11/2007 FC V_80 : FQ 14934 Pb CWAS Qd on dblclick sur la liste, on arrive toujours sur le 1er élément
PT2  06/12/2007 FC V_81 : FQ 15006 pb multi
}
Unit UTofPGMulParamHisto ;

Interface

Uses StdCtrls, 
     Controls, 
     Classes,
{$IFNDEF EAGLCLIENT}
     db,
     FE_Main,
     HDB,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF} 
     mul, 
{$else}
     eMul, 
     MainEAGL,
{$ENDIF}
    uTob,
     forms, 
     sysutils, 
     ComCtrls,
     HCtrls, 
     HEnt1, 
     HMsgBox,
     HTB97,
     UTOF ; 

Type
  TOF_PGMULPARAMHISTO = Class (TOF)
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    private
    {$IFNDEF EAGLCLIENT}
    Grille : THDBGrid;
    {$ELSE}
    Grille : THGrid;
    {$ENDIF}
    TypeSaisie : String;
    Procedure GrilleDblClick(Sender : TObject);
    Procedure CreationZone(Sender : TObject);
  end ;

Implementation

procedure TOF_PGMULPARAMHISTO.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_PGMULPARAMHISTO.OnArgument (S : String ) ;
var BOuv : TToolBarButton97;
begin
  Inherited ;
  TypeSaisie := ReadTokenPipe(S,';');
  SetControltext('PPP_PGTYPEINFOLS',TypeSaisie);
  If TypeSaisie ='SAL' then
  begin
    SetControlVisible('TPPP_TYPENIVEAU',False);
    SetControlVisible('PPP_TYPENIVEAU',False);
    TFMul(Ecran).Caption := 'Paramétrage historique salarié';
  end
  else TFMul(Ecran).Caption := 'Paramétrage éléments dynamiques';
  UpdateCaption(TFMul(Ecran));
  BOuv  :=  TToolbarButton97 (getcontrol ('BInsert'));
  if Bouv  <>  NIL then  BOuv.OnClick  :=  CreationZone;
  {$IFNDEF EAGLCLIENT}
  Grille := THDBGrid (GetControl ('Fliste'));
  {$ELSE}
  Grille := THGrid (GetControl ('Fliste'));
  {$ENDIF}
  if Grille  <>  NIL then Grille.OnDblClick  :=  GrilleDblClick;
end ;

Procedure TOF_PGMULPARAMHISTO.GrilleDblClick(Sender : TObject);
var St : String;
    Bt  :  TToolbarButton97;
begin
//DEB PT1
{$IFDEF EAGLCLIENT}
TFMul (Ecran).Q.TQ.Seek (TFMul (Ecran).FListe.Row-1) ;
{$ENDIF}
//FIN PT1
  St := TFmul(Ecran).Q.FindField('PPP_PGINFOSMODIF').AsString;
  St := St + ';' + TFmul(Ecran).Q.FindField('PPP_PREDEFINI').AsString;    //PT2
  St := St + ';' + TFmul(Ecran).Q.FindField('PPP_NODOSSIER').AsString;    //PT2
  If TypeSaisie = 'SAL' then AGLLanceFiche('PAY', 'PARAMHISTO', '', St,'SAL')
  else AGLLanceFiche('PAY', 'PARAMZONELIBRE', '', St, 'ZLS');
  Bt  :=  TToolbarButton97 (GetControl('BCherche'));
  if Bt  <>  NIL then Bt.click;
end;

Procedure TOF_PGMULPARAMHISTO.CreationZone(Sender : TObject);
var Bt  :  TToolbarButton97;
begin
  If TypeSaisie = 'SAL' then AGLLanceFiche('PAY', 'PARAMHISTO', '', '','SAL;ACTION=CREATION')
  else AGLLanceFiche('PAY', 'PARAMZONELIBRE', '', '', 'ZLS;ACTION=CREATION');
  Bt  :=  TToolbarButton97 (GetControl('BCherche'));
  if Bt  <>  NIL then Bt.click;
end;


Initialization
  registerclasses ( [ TOF_PGMULPARAMHISTO ] ) ;
end.

