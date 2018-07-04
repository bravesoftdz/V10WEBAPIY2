{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 05/02/2002
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : OUTIVERIFDOSSIER ()
Mots clefs ... : TOF;OUTIVERIFDOSSIER
*****************************************************************}
unit uTOFVerifDossier;

interface

uses StdCtrls,
  Controls,
  Classes,
{$IFNDEF EAGLCLIENT}
  db,
  dbtables,
{$ENDIF}
  forms,
  sysutils,
  ComCtrls,
  HCtrls,
  HEnt1,
  HMsgBox,
  UTOF,
  Menus,
  Filtre,
  Htb97,
  Vierge,
  Windows,
  uLibWindows;

type
  TOF_OUTIVERIFDOSSIER = class(TOF)
    procedure OnNew; override;
    procedure OnDelete; override;
    procedure OnUpdate; override;
    procedure OnLoad; override;
    procedure OnArgument(S: string); override;
    procedure OnKeyDownEcran(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure OnClose; override;

    procedure OnClickBValider(Sender : TObject);
    procedure OnClickBSelect (Sender : TObject);
    function OkSelection : Boolean;

  private
    FPages: TPageControl;
    FFiltres: THValComboBox;
    FBoAllSelected : Boolean;

    function SuppressionJournal : Boolean;

  end;

implementation

uses uControlCP;

procedure TOF_OUTIVERIFDOSSIER.OnNew;
begin
  inherited;
end;

procedure TOF_OUTIVERIFDOSSIER.OnDelete;
begin
  inherited;
end;

procedure TOF_OUTIVERIFDOSSIER.OnUpdate;
begin
  inherited;
end;

procedure TOF_OUTIVERIFDOSSIER.OnLoad;
begin
  inherited;
end;

procedure TOF_OUTIVERIFDOSSIER.OnArgument(S: string);
begin
  inherited;
  TToolbarButton97(GetControl('BValider')).OnClick := OnClickBValider;
  TToolBarButton97(GetControl('BSelect')).OnClick := OnClickBSelect;

  FPages := TPageControl(GetControl('FPages'));
  //FPages.ActivePage := TTabSheet(GetControl('FTabSheet1'));
  //TTabSheet(GetControl('FTabSheet1')).TabVisible := False;

  // Suppression menu analytique
  TGroupBox(GetControl('FGBAnalytique')).Visible := False;
  Ecran.Height := 295;

  Ecran.OnKeyDown := OnKeyDownEcran;

  FFiltres := THValComboBox(GetControl('FFILTRES'));
  FFiltres.ItemIndex := 0;
end;

procedure TOF_OUTIVERIFDOSSIER.OnClose;
begin
  inherited;
end;

procedure TOF_OUTIVERIFDOSSIER.OnKeyDownEcran(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
  case Key of

    VK_F10 : TToolbarButton97(GetControl('BValider')).Click;

    VK_Escape : TToolbarButton97(GetControl('BFerme')).Click;

    65 : if Shift=[ssCtrl] then TToolbarButton97(GetControl('BSelect')).Click;

  else
  end;

end;

{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 04/04/2002
Modifié le ... :   /  /
Description .. : Test si au moins un controle a été sélectionné
Mots clefs ... :
*****************************************************************}
function TOF_OUTIVERIFDOSSIER.OkSelection: Boolean;
var lIt : integer;
begin
  Result := False;

  for lIt := 0 to Ecran.ComponentCount - 1 do
  begin
    if Ecran.Components[lIt] is TCheckBox then
    begin
      if (not Result) and (TCheckBox( Ecran.Components[lIt] ).Checked) then
        Result := True;
    end;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 04/04/2002
Modifié le ... :   /  /    
Description .. : Click sur le Bouton Valider
Mots clefs ... : 
*****************************************************************}
procedure TOF_OUTIVERIFDOSSIER.OnClickBValider(Sender: TObject);
begin
  if not OkSelection then
  begin
    PgiInfo('Vous devez sélectionner les contrôles à effectuer', Ecran.Caption);
    Exit;
  end;

  if PgiAsk('Confirmez vous la vérification des dossiers ?', Ecran.Caption) = mrNo then
  begin
    TFVierge(Ecran).Retour := 'NON';
    Exit;
  end
  else
  begin
    if not SuppressionJournal then Exit;

    // Enregistrement du filtre à utiliser pour le traitement
    SaveFiltre(cFI_TableVerif, FFiltres, FPages);
    TFVierge(Ecran).Retour := 'OUI; /CPVERIF=' + FFiltres.Text;
    Ecran.Close;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 04/04/2002
Modifié le ... :   /  /    
Description .. : Click sur le Bouton de Selection
Mots clefs ... : 
*****************************************************************}
procedure TOF_OUTIVERIFDOSSIER.OnClickBSelect(Sender: TObject);
var i : integer;
begin
  FBoAllSelected := not FBoAllSelected;

  for i := 0 to Ecran.ComponentCount - 1 do
  begin
    if Ecran.Components[i] is TCheckBox then
    begin
      TCheckBox( Ecran.Components[i] ).Checked := FBoAllSelected;
    end;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 15/04/2002
Modifié le ... :   /  /
Description .. : Supprime les anciens journaux d'erreurs de vérification
Mots clefs ... :
*****************************************************************}
function TOF_OUTIVERIFDOSSIER.SuppressionJournal : Boolean;
begin
  Result := True;

  if not FileExistsDelete(GetWindowsTempPath + '\' + CNomRapportCpt) then
  begin
    PgiError('Traitement annulé : impossible de supprimer l''ancien rapport sur les comptes.', Ecran.Caption);
    Result := False;
  end;

  if not FileExistsDelete(GetWindowsTempPath + '\' + CNomRapportMvt) then
  begin
    PgiError('Traitement annulé : impossible de supprimer l''ancien rapport sur les mouvements.', Ecran.Caption);
    Result := False;
  end;

  if not FileExistsDelete(GetWindowsTempPath + '\' + CNomRapportLet) then
  begin
    PgiError('Traitement annulé : impossible de supprimer l''ancien rapport sur le lettrage.', Ecran.Caption);
    Result := False;
  end;

end;


initialization
  registerclasses([TOF_OUTIVERIFDOSSIER]);
end.

