{***********UNITE*************************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 21/11/2002
Modifié le ... : 22/11/2002
Description .. : Source TOF de la FICHE : MBO/AFFCDELANCE
Mots clefs ... : TOF;MBOAFFCDELANCE;AFFCDE
*****************************************************************}
unit MBOAFFCDELANCE_TOF;

interface

uses
  Classes, sysutils, Forms, controls, stdctrls, graphics,
  {$IFDEF EAGLCLIENT}
  Maineagl,
  {$ELSE}
  FE_Main,
  {$ENDIF}
  UTOF, UTOB, HCtrls, HEnt1, HTB97, UtilAffCde;

type
  TOF_MBOAFFCDELANCE = class(TOF)
    procedure OnNew; override;
    procedure OnDelete; override;
    procedure OnUpdate; override;
    procedure OnLoad; override;
    procedure OnArgument(S: string); override;
    procedure OnClose; override;
  private
    Etape: TTypeEtapeAffCde;
    LanceAuto: boolean;
    LanceFin: boolean;
    CodeAff: string;
    TOBAffCde: TOB;
    procedure MessTraitement(Indice: integer; Libelle: string; Fini: boolean);
    procedure DebutTraitement(CodeStatut: string);
    procedure FinTraitement(CodeStatut: string);
    procedure Traitement;
    function LanceAffectation: integer;
    function LancePreparation: integer;
    procedure BPARAMClick(Sender: TObject);
    function GetRadioButtonState(NomChamp: string): boolean;
    procedure ActivePanels;
    procedure GEA_CODEAFFChange(Sender: TObject);
  end;

implementation

uses
  EntGC;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 26/11/2002
Modifié le ... : 29/11/2002
Description .. : Affichage d'un message d'état du traitement avec la
Suite ........ : mention 'En cours...' ou 'ok'
Mots clefs ... :
*****************************************************************}

procedure TOF_MBOAFFCDELANCE.MessTraitement(Indice: integer; Libelle: string; Fini: boolean);
var
  Stg, NomChamp: string;
begin
  NomChamp := 'PHASE' + IntToStr(Indice);
  SetControlText(NomChamp, TraduireMemoire(Libelle));
  SetControlVisible(NomChamp, True);
  NomChamp := 'ETAT' + IntToStr(Indice);
  if Fini then Stg := 'Ok' else Stg := 'En cours...';
  SetControlText(NomChamp, TraduireMemoire(Stg));
  SetControlVisible(NomChamp, True);
  Application.ProcessMessages;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 26/11/2002
Modifié le ... : 29/11/2002
Description .. : Début du traitement
Mots clefs ... :
*****************************************************************}

procedure TOF_MBOAFFCDELANCE.DebutTraitement(CodeStatut: string);
begin
  SetControlVisible('ENCOURS', True);
  SetControlEnabled('BValider', False);
  SourisSablier;
  // Mise à jour du statut
  MajStatutAFFCDEENTETE(CodeAff, CodeStatut, False);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 26/11/2002
Modifié le ... : 29/11/2002
Description .. : Fin du traitement
Mots clefs ... :
*****************************************************************}

procedure TOF_MBOAFFCDELANCE.FinTraitement(CodeStatut: string);
var
  Ctrl: TControl;
begin
  // Mise à jour du statut
  MajStatutAFFCDEENTETE(CodeAff, CodeStatut, True);
  // Message écran
  Ctrl := GetControl('ENCOURS');
  if (Ctrl <> nil) and (Ctrl is TLabel) then
  begin
    TLabel(Ctrl).Caption := 'Traitement terminé.';
    TLabel(Ctrl).Font.Style := [fsBold];
  end;
  // Ecriture en base de registres du code affectation utilisé
  SaveSynRegKey('CodeAffCde', CodeAff, True);
  SourisNormale;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 26/11/2002
Modifié le ... : 26/11/2002
Description .. : Traitement
Mots clefs ... :
*****************************************************************}

procedure TOF_MBOAFFCDELANCE.Traitement;
var
  Stg, CodeStatut: string;
  Nbl: integer;
begin
  CodeAff := GetControlText('GEA_CODEAFF');
  CodeStatut := DonneCodeStatutAFFCDE(Etape);
  if CodeAff <> '' then
  begin
    // vérifie si l'étape précédente a été faite
    if not VerifStatutAFFCDEENTETE(CodeAff, Etape) then Exit;
    // Lancement du traitement
    DebutTraitement(CodeStatut);
    Stg := '';
    case Etape of
      afcReservation:
        begin
          // Réservation des commandes
          Stg := 'Suppression des stocks de l''affectation';
          MessTraitement(1, Stg, False);
          Nbl := DeleteAFFCDEDISPO(CodeAff);
          Stg := Stg + ' : ' + IntToStr(Nbl) + TraduireMemoire(' ligne(s) supprimée(s)');
          MessTraitement(1, Stg, True);

          Stg := 'Suppression des pièces de l''affectation';
          MessTraitement(2, Stg, False);
          Nbl := DeleteAFFCDEPIECE(CodeAff);
          Stg := Stg + ' : ' + IntToStr(Nbl) + TraduireMemoire(' pièce(s) supprimé(s)');
          MessTraitement(2, Stg, True);

          Stg := 'Suppression des lignes génériques';
          MessTraitement(3, Stg, False);
          Nbl := DeleteAFFCDELIGNE(CodeAff, 'GEN');
          Stg := Stg + ' : ' + IntToStr(Nbl) + TraduireMemoire(' ligne(s) supprimée(s)');
          MessTraitement(3, Stg, True);

          Stg := 'Création des stocks de l''affectation';
          MessTraitement(4, Stg, False);
          Nbl := InsertAFFCDEDISPO(CodeAff);
          Stg := Stg + ' : ' + IntToStr(Nbl) + TraduireMemoire(' stock(s) créé(s)');
          MessTraitement(4, Stg, True);

          Stg := 'Réservation des lignes de commandes';
          MessTraitement(5, Stg, False);
          Nbl := LanceAffectation;
          Stg := Stg + ' - ' + IntToStr(Nbl) + TraduireMemoire(' ligne(s) réservée(s)');
          MessTraitement(5, Stg, True);

          Stg := 'Création des lignes génériques';
          MessTraitement(6, Stg, False);
          Nbl := InsertGeneriqueAFFCDELIGNE(CodeAff);
          Stg := Stg + ' : ' + IntToStr(Nbl) + TraduireMemoire(' ligne(s) créée(s)');
          MessTraitement(6, Stg, True);

          Stg := 'Création des pièces';
          MessTraitement(7, Stg, False);
          Nbl := InsertPieceAFFCDELIGNE(CodeAff);
          Stg := Stg + ' : ' + IntToStr(Nbl) + TraduireMemoire(' pièce(s) créée(s)');
          MessTraitement(7, Stg, True);
          CodeStatut := '020';
        end;
      afcAffectation:
        begin
          // Affectation des commandes
          Stg := 'Affectation des lignes de commandes';
          MessTraitement(1, Stg, False);
          Nbl := LanceAffectation;
          Stg := Stg + ' : ' + IntToStr(Nbl) + TraduireMemoire(' ligne(s) affectée(s)');
          MessTraitement(1, Stg, True);
          CodeStatut := '030';
        end;
      afcPreparation:
        begin
          // Préparation des commandes
          Stg := 'Préparation des lignes de commandes';
          MessTraitement(1, Stg, False);
          Nbl := LancePreparation;
          Stg := Stg + ' : ' + IntToStr(Nbl) + TraduireMemoire(' ligne(s) préparée(s)');
          MessTraitement(1, Stg, True);
          CodeStatut := '040';
        end;
      afcFin:
        begin
          // Fin de l'affectation des commandes : suppression des lignes existantes
          Stg := 'Suppression des lignes d''affectation';
          MessTraitement(1, Stg, False);
          Nbl := DeleteAFFCDELIGNE(CodeAff, '');
          Stg := Stg + ' : ' + IntToStr(Nbl) + TraduireMemoire(' ligne(s) supprimée(s)');
          MessTraitement(1, Stg, True);

          Stg := 'Suppression des stocks de l''affectation';
          MessTraitement(2, Stg, False);
          Nbl := DeleteAFFCDEDISPO(CodeAff);
          Stg := Stg + ' : ' + IntToStr(Nbl) + TraduireMemoire(' stock(s) supprimé(s)');
          MessTraitement(2, Stg, True);

          Stg := 'Suppression des pièces de l''affectation';
          MessTraitement(3, Stg, False);
          Nbl := DeleteAFFCDEPIECE(CodeAff);
          Stg := Stg + ' : ' + IntToStr(Nbl) + TraduireMemoire(' pièce(s) supprimé(s)');
          MessTraitement(3, Stg, True);
          CodeStatut := '050';
        end;
    end;
    FinTraitement(CodeStatut);
    if LanceFin then
      AGLLanceFiche('MBO', 'AFFCDELANCE', '', '', 'FIN;LANCEAUTO;CODEAFF=' + CodeAff)
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 26/11/2002
Modifié le ... : 26/11/2002
Description .. : Lancement du traitement d'affectation
Mots clefs ... :
*****************************************************************}

function TOF_MBOAFFCDELANCE.LanceAffectation: integer;
var
  MaxNbRow: integer;
  UseDataMem: boolean;
  VerifArtFerme: boolean;
  VerifCliFerme: boolean;
  Ctrl: TControl;
begin
  Ctrl := GetControl('SPMAXNBROW');
  if (Ctrl <> nil) and (Ctrl is THSpinEdit) then
    MaxNbRow := THSpinEdit(Ctrl).Value
  else
    MaxNbRow := 1000;
  UseDataMem := (GetCheckBoxState('CBUSEDATAMEM') = cbChecked);
  VerifArtFerme := (GetCheckBoxState('CBARTFERME') = cbChecked);
  VerifCliFerme := (GetCheckBoxState('CBCLIFERME') = cbChecked);
  Result := TraitementAffectationCde(CodeAff, Etape, MaxNbRow, UseDataMem, VerifArtFerme, VerifCliFerme, '');
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 26/11/2002
Modifié le ... : 26/11/2002
Description .. : Lancement du traitement de préparation
Mots clefs ... :
*****************************************************************}

function TOF_MBOAFFCDELANCE.LancePreparation: integer;
var
  Eclate: boolean;
  DeGroupeRemise: boolean;
  AvecComment: boolean;
  AvecCompteRendu: boolean;
  ContinuOnError: boolean;
  MajToutStock: boolean;
  TraitFinTravaux: integer;
  Ctrl: TControl;
begin
  Eclate := (GetCheckBoxState('CBECLATE') = cbChecked);
  DeGroupeRemise := (GetCheckBoxState('CBDEGROUPEREMISE') = cbChecked);
  AvecComment := (GetCheckBoxState('CBAVECCOMMENT') = cbChecked);
  AvecCompteRendu := (GetCheckBoxState('CBAVECCOMPTERENDU') = cbChecked);
  ContinuOnError := (GetCheckBoxState('CBCONTINUONERROR') = cbChecked);
  if GetRadioButtonState('RBFIN') then
  begin
    TraitFinTravaux := 2;
    MajToutStock := False;
    LanceFin := True;
  end
  else
    if GetRadioButtonState('RBRELIQUAT') then
  begin
    TraitFinTravaux := 1;
    MajToutStock := False;
    LanceFin := False;
  end
  else
    if GetRadioButtonState('RBRELIQSTOCK') then
  begin
    TraitFinTravaux := 1;
    MajToutStock := True;
    LanceFin := False;
  end
  else
  begin
    TraitFinTravaux := 0;
    MajToutStock := False;
    LanceFin := False;
  end;
  Ctrl := GetControl('COMPTERENDU');
  if (Ctrl <> nil) and not (Ctrl is TMemo) then Ctrl := nil ;
  Result := TraitementPreparationCde(CodeAff, Eclate, DeGroupeRemise, AvecComment,
    AvecCompteRendu, ContinuOnError, MajToutStock, TraitFinTravaux, TMemo(Ctrl));
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 13/02/2003
Modifié le ... : 13/02/2003
Description .. : Traitement du bouton "Paramètres avancés"
Mots clefs ... :
*****************************************************************}

procedure TOF_MBOAFFCDELANCE.BPARAMClick(Sender: TObject);
var
  Ctrl: TControl;
  Nom: string;
  Visu: boolean;
begin
  case Etape of
    afcReservation,
      afcAffectation:
      Nom := 'POPTIONS';
    afcPreparation:
      Nom := 'POPTPREPA';
  else
    Nom := '';
  end;
  if Nom <> '' then
  begin
    Ctrl := GetControl('BPARAM');
    if (Ctrl <> nil) and (Ctrl is TToolbarButton97) then
      Visu := TToolbarButton97(Ctrl).Down
    else
      Visu := False;
    Ctrl := GetControl(Nom);
    if Ctrl <> nil then Ctrl.Visible := Visu;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 13/02/2003
Modifié le ... : 13/02/2003
Description .. : Lit le statut d'un bouton radio
Mots clefs ... :
*****************************************************************}

function TOF_MBOAFFCDELANCE.GetRadioButtonState(NomChamp: string): boolean;
var
  Ctrl: TControl;
begin
  Result := False;
  Ctrl := GetControl(NomChamp);
  if (Ctrl <> nil) and (Ctrl is TRadioButton) then
    Result := TRadioButton(Ctrl).Checked;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 26/11/2002
Modifié le ... : 26/11/2002
Description .. : Rende visible un panel
Mots clefs ... :
*****************************************************************}

procedure TOF_MBOAFFCDELANCE.ActivePanels;
var
  Ctrl: TControl;
  Nom: string;
begin
  case Etape of
    afcReservation,
      afcAffectation:
      Nom := 'POPTIONS';
    afcPreparation:
      Nom := 'POPTPREPA';
  else
    Nom := '';
  end;
  if Nom <> '' then
  begin
    Ctrl := GetControl(Nom);
    if Ctrl <> nil then
    begin
      Ctrl.Visible := False;
      Ctrl.Align := alBottom;
    end;
    Ctrl := GetControl('BPARAM');
    if (Ctrl <> nil) and (Ctrl is TToolbarButton97) then
    begin
      Ctrl.Visible := V_PGI.SAV;
      TToolbarButton97(Ctrl).OnClick := BPARAMClick;
    end;
  end;
  Ctrl := GetControl('PRESULTAT');
  if Ctrl <> nil then
  begin
    Ctrl.Visible := True;
    Ctrl.Align := alClient;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 26/11/2002
Modifié le ... : 26/11/2002
Description .. : Changement du code affectation
Mots clefs ... :
*****************************************************************}

procedure TOF_MBOAFFCDELANCE.GEA_CODEAFFChange(Sender: TObject);
var
  Ok : boolean;
begin
  CodeAff := GetControlText('GEA_CODEAFF');
  if TOBAffCde <> nil then FreeAndNil(TOBAffCde);
  TOBAffCde := TOB.Create('AFFCDEENTETE', nil, -1);
  TOBAffCde.SelectDB('"'+ CodeAff +'"', nil);

  Ok := (TOBAffCde.GetValue('GEA_VERIFARTFERME') = 'X');
  SetControlChecked('CBARTFERME', Ok);
  Ok := (TOBAffCde.GetValue('GEA_VERIFCLIFERME') = 'X');
  SetControlChecked('CBCLIFERME', Ok);
  Ok := (TOBAffCde.GetValue('GEA_ECLATEPARCDE') = 'X');
  SetControlChecked('CBECLATE', Ok);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 26/11/2002
Modifié le ... : 26/11/2002
Description .. : OnNew
Mots clefs ... :
*****************************************************************}

procedure TOF_MBOAFFCDELANCE.OnNew;
begin
  inherited;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 26/11/2002
Modifié le ... : 26/11/2002
Description .. : OnDelete
Mots clefs ... :
*****************************************************************}

procedure TOF_MBOAFFCDELANCE.OnDelete;
begin
  inherited;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 26/11/2002
Modifié le ... : 26/11/2002
Description .. : OnUpdate
Mots clefs ... :
*****************************************************************}

procedure TOF_MBOAFFCDELANCE.OnUpdate;
begin
  inherited;
  if not LanceAuto then Traitement;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 26/11/2002
Modifié le ... : 26/11/2002
Description .. : OnLoad
Mots clefs ... :
*****************************************************************}

procedure TOF_MBOAFFCDELANCE.OnLoad;
begin
  inherited;
  if LanceAuto then Traitement;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 26/11/2002
Modifié le ... : 26/11/2002
Description .. : OnArgument
Mots clefs ... :
*****************************************************************}

procedure TOF_MBOAFFCDELANCE.OnArgument(S: string);
var
  Titre, Stg, sArg, sCode: string;
  Ctrl: TControl;
begin
  inherited;
  // Lecture en base de registres du dernier code affectation utilisé
  sCode := GetSynRegKey('CodeAffCde', '', True);
  LanceAuto := False;
  LanceFin := False;
  Etape := afcFin;
  Titre := 'Fin de l''affectation des commandes';
  Stg := S;
  while Stg <> '' do
  begin
    sArg := ReadTokenSt(Stg);
    if sArg = 'RESERVATION' then
    begin
      Etape := afcReservation;
      Titre := 'Réservation des commandes';
    end else
      if sArg = 'AFFECTATION' then
    begin
      Etape := afcAffectation;
      Titre := 'Affectation des commandes';
    end else
      if sArg = 'PREPARATION' then
    begin
      Etape := afcPreparation;
      Titre := 'Préparation des commandes';
    end else
      if sArg = 'LANCEAUTO' then
    begin
      LanceAuto := True;
    end else
      if Copy(sArg, 1, 8) = 'CODEAFF=' then
    begin
      sCode := Copy(sArg, 9, MaxInt);
    end;
  end;
  ActivePanels;
  InitCaption(Ecran, '', TraduireMemoire(Titre));

  TOBAffCde := nil;
  Ctrl := GetControl('GEA_CODEAFF');
  if (Ctrl <> nil) and (Ctrl is THValComboBox) then
    THValComboBox(Ctrl).OnChange := GEA_CODEAFFChange;
  if sCode <> '' then SetControlText('GEA_CODEAFF', sCode);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 26/11/2002
Modifié le ... : 26/11/2002
Description .. : OnClose
Mots clefs ... :
*****************************************************************}

procedure TOF_MBOAFFCDELANCE.OnClose;
begin
  inherited;
  if TOBAffCde <> nil then FreeAndNil(TOBAffCde);
end;

initialization
  registerclasses([TOF_MBOAFFCDELANCE]);
end.
