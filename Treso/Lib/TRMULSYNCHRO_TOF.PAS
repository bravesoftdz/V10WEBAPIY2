{ Unité : Source TOF de la FICHE : TRMULSYNCHRO
--------------------------------------------------------------------------------------
    Version  |  Date  | Qui | Commentaires
--------------------------------------------------------------------------------------
 0.91          04/09/03  JP   Création de l'unité
 1.50.000.000  16/04/04  JP   Suppression de la fonction de réparation des natures
 6.0X.xxx.xxx  04/08/04  JP   Gestion des PopupMenus
 7.05.001.001  09/08/06  JP   Mise en place du multi sociétés
 8.10.001.004  08/08/07  JP   Gestion des confidentialités
--------------------------------------------------------------------------------------}
Unit TRMULSYNCHRO_TOF ;

interface

uses {$IFDEF VER150} variants,{$ENDIF}
  StdCtrls, Controls, Classes, HQry,
  {$IFNDEF EAGLCLIENT}
  Mul, FE_Main,
  {$ELSE}
  eMul, MaineAGL,
  {$ENDIF}
  Forms, ComCtrls, HCtrls, HEnt1, HMsgBox,
  {$IFDEF TRCONF}
  uLibConfidentialite,
  {$ELSE}
  UTOF,
  {$ENDIF TRCONF}
  Commun, HTB97, Menus;


type
  {$IFDEF TRCONF}
  TOF_TRMULSYNCHRO = class (TOFCONF)
  {$ELSE}
  TOF_TRMULSYNCHRO = class (TOF)
  {$ENDIF TRCONF}
    procedure OnArgument(S : string); override;
  private
    bRepareDt : TToolBarButton97;
    bRepareFx : TToolBarButton97;
    PopupMenu : TPopUpMenu; {04/08/04}

    procedure TravailDates    (Sender : TObject);
    procedure TravailFlux     (Sender : TObject);
    procedure ListeOnDbleClick(Sender : TObject);
    procedure NoDossierChange (Sender : TObject);
  end;

procedure TRLanceFiche_VisuSynchro(Dom, Fiche, Range, Lequel, Arguments : string);

implementation

uses
  AccImportTreso, UtilPgi, UProcGen;

{---------------------------------------------------------------------------------------}
procedure TRLanceFiche_VisuSynchro(Dom, Fiche, Range, Lequel, Arguments : string);
{---------------------------------------------------------------------------------------}
begin
  AGLLanceFiche(Dom, Fiche, Range, Lequel, Arguments);
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRMULSYNCHRO.OnArgument (S : String ) ;
{---------------------------------------------------------------------------------------}
begin
  {$IFDEF TRCONF}
  TypeConfidentialite := tyc_Banque + ';';
  {$ENDIF TRCONF}
  inherited;
  Ecran.HelpContext := 50000131;
  {Ces deux boutons permettent d'éxecuter les manipulations indépendamment}
  bRepareDt := TToolBarButton97(GetControl('BREPARED'));
  bRepareFx := TToolBarButton97(GetControl('BREPAREF'));
  bRepareFx.Hint := TraduireMemoire('Réparer les codes rubriques des imports comptables');
  bRepareDt.Hint := TraduireMemoire('Réparer les dates de valeur des imports comptables');
  bRepareFx.OnClick := TravailFlux;
  bRepareDt.OnClick := TravailDates;

  PopupMenu := TPopUpMenu(GetControl('POPUPMENU')); {04/08/04}
  PopupMenu.Items[0].OnClick := TravailDates;
  PopupMenu.Items[1].OnClick := TravailFlux;
  AddMenuPop(PopupMenu, '', '');

  {08/08/06 : gestion du multi sociétés}
  SetControlVisible('TE_NODOSSIER' , IsTresoMultiSoc);
  SetControlVisible('TTE_NODOSSIER', IsTresoMultiSoc);

  Ecran.Refresh;
  Application.ProcessMessages;
  TFMul(Ecran).FListe.OnDblClick := ListeOnDbleClick;

  {Gestion des filtres multi sociétés sur banquecp et dossier}
  THEdit(GetControl('TE_GENERAL')).Plus := FiltreBanqueCp(THEdit(GetControl('TE_GENERAL')).DataType,'', '');
  THValComboBox(GetControl('TE_NODOSSIER')).Plus := 'DOS_NODOSSIER ' + FiltreNodossier;
  THValComboBox(GetControl('TE_NODOSSIER')).OnChange := NoDossierChange;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRMULSYNCHRO.NoDossierChange(Sender : TObject);
{---------------------------------------------------------------------------------------}
begin
  THEdit(GetControl('TE_GENERAL')).Plus := FiltreBanqueCp(THEdit(GetControl('TE_GENERAL')).DataType, '', GetControlText('TE_NODOSSIER'));
  SetControlText('TE_GENERAL', '');
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRMULSYNCHRO.TravailDates(Sender : TObject);
{---------------------------------------------------------------------------------------}
begin
  TravaillerDates(RecupWhereCritere(TPageControl(GetControl('PAGES'))));
  {JP 04/04/04 : FQ 10017}
  HShowMessage('1;' + Ecran.Caption + ';La réparation des dates de valeur a été effectuée.;I;O;O;O;', '', '');
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRMULSYNCHRO.TravailFlux(Sender : TObject);
{---------------------------------------------------------------------------------------}
begin
  TravaillerFlux;
  {JP 04/04/04 : FQ 10017}
  HShowMessage('2;' + Ecran.Caption + ';La réparation des codes rubriques a été effectuée.;I;O;O;O;', '', '');
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRMULSYNCHRO.ListeOnDbleClick(Sender : Tobject);
{---------------------------------------------------------------------------------------}
var
  s : string;
begin
  s := GetField('TE_NODOSSIER') + ';' + GetField('TE_NUMTRANSAC') + ';' +
       VarToStr(GetField('TE_NUMEROPIECE')) + ';' + VarToStr(GetField('TE_NUMLIGNE'));
  AGLLanceFiche('TR', 'TRFICECRITURE', '', s, GetField('TE_NATURE') + ';');
end;


initialization
  RegisterClasses([TOF_TRMULSYNCHRO]);
  
end.
