{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 06/02/2002
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : OUTICORREDOSSIER ()
Mots clefs ... : TOF;OUTICORREDOSSIER
*****************************************************************}
unit uTOFCorrectDossier;

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
  TOF_OUTICORREDOSSIER = class(TOF)
    procedure OnNew; override;
    procedure OnDelete; override;
    procedure OnUpdate; override;
    procedure OnLoad; override;
    procedure OnArgument(S: string); override;
    procedure OnKeyDownEcran(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure OnClose; override;

    procedure OnClickBValider(Sender : TObject);
    procedure OnClickBselect (Sender : TObject);

  private
    FPages: TPageControl;
    FFiltres: THValComboBox;
    FBoAllSelected : Boolean;
  end;

implementation

uses UControlCP;

procedure TOF_OUTICORREDOSSIER.OnNew;
begin
  inherited;
end;

procedure TOF_OUTICORREDOSSIER.OnDelete;
begin
  inherited;
end;

procedure TOF_OUTICORREDOSSIER.OnUpdate;
begin
  inherited;
end;

procedure TOF_OUTICORREDOSSIER.OnLoad;
begin
  inherited;
end;

procedure TOF_OUTICORREDOSSIER.OnArgument(S: string);
begin
  inherited;
  TToolbarButton97(GetControl('BValider')).OnClick := OnClickBValider;
  TToolBarButton97(GetControl('BSelect')).OnClick := OnClickBSelect;

  FPages := TPageControl(GetControl('FPages'));
  TTabSheet(GetControl('FTabSheet')).TabVisible := False;
  FPages.ActivePage := TTabSheet(GetControl('FTabSheet'));

  Ecran.OnKeyDown := OnKeyDownEcran;

  FFiltres := THValComboBox(GetControl('FFILTRES'));
  FFiltres.ItemIndex := 0;
end;

procedure TOF_OUTICORREDOSSIER.OnClose;
begin
  inherited;
end;

procedure TOF_OUTICORREDOSSIER.OnKeyDownEcran(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
  case Key of

    VK_F10 : TToolbarButton97(GetControl('BValider')).Click;

    VK_Escape : TToolbarButton97(GetControl('BFerme')).Click;

    65 : if Shift=[ssCtrl] then TToolbarButton97(GetControl('BSelect')).Click;

  else
  end;

end;

procedure TOF_OUTICORREDOSSIER.OnClickBValider(Sender: TObject);
begin
  if not (// Fiches
          TCheckBox(GetControl('FCBGeneraux')).Checked       or
          TCheckBox(GetControl('FCBAuxiliaires')).Checked    or
          TCheckBox(GetControl('FCBSections')).Checked       or
          TCheckBox(GetControl('FCBJournaux')).Checked       or
          TCheckBox(GetControl('FCBSoldeCompte')).Checked    or
          // Mouvements comptables
          TCheckBox(GetControl('FCBEquilibreMvt')).Checked   or
          TCheckBox(GetControl('FCBDateLettre')).Checked     or
          TCheckBox(GetControl('FCBContrepartie')).Checked   or
          TCheckBox(GetControl('FCBPeriodeSemaine')).Checked or
          TCheckBox(GetControl('FCBJournal')).Checked) then
  begin
    PgiInfo('Vous devez sélectionner les corrections à effectuer', Ecran.Caption);
    Exit;
  end;

  if PgiAsk('Confirmez vous la correction des dossiers ?', Ecran.Caption) = mrNo then
  begin
    TFVierge(Ecran).Retour := 'NON';
    Exit;
  end
  else
  begin
    // Suppression du rapport de Correction des dossiers
    if not FileExistsDelete(GetWindowsTempPath + '\' + CNomRapportCor) then
    begin
      PgiError('Impossible de supprimer l''ancien rapport sur les erreurs de correction.', Ecran.Caption);
      Exit;
    end;

    SaveFiltre(cFI_TableCorrect, FFiltres, FPages);
    TFVierge(Ecran).Retour := 'OUI; /CPCORRECT=' + FFiltres.Text;
    Ecran.Close;
  end;
end;

procedure TOF_OUTICORREDOSSIER.OnClickBselect(Sender: TObject);
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

initialization
  registerclasses([TOF_OUTICORREDOSSIER]);
end.

