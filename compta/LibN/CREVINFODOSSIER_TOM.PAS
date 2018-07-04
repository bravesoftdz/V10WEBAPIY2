{***********UNITE*************************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 15/03/2007
Modifié le ... :   /  /
Description .. : Source TOM de la TABLE : CREVINFODOSSIER (CREVINFODOSSIER)
Mots clefs ... : TOM;CREVINFODOSSIER
*****************************************************************}
Unit CREVINFODOSSIER_TOM ;

Interface

Uses StdCtrls, 
     Controls, 
     Classes,
     Graphics,      // ClBtnFace 
{$IFDEF EAGLCLIENT}
     eMul,
     eFiche,        // TFFiche
     MainEagl,      // AGLLanceFiche
     UtileAGL,      // LanceEtat
{$ELSE}
     db,
     HDB,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul,
     Fiche,         // TFFiche
     FE_main,       // AGLLanceFiche
     EdtREtat,      // LanceEtat
{$ENDIF}
     forms,
     sysutils, 
     ComCtrls,
     HCtrls, 
     HEnt1, 
     HMsgBox, 
     UTOM,
     HRichOle,     // THRichEditOle
     Htb97,        // TToolBarButton97
     UTob ;

Type
  TOM_CREVINFODOSSIER = Class (TOM)
    procedure OnNewRecord                ; override ;
    procedure OnDeleteRecord             ; override ;
    procedure OnUpdateRecord             ; override ;
    procedure OnAfterUpdateRecord        ; override ;
    procedure OnAfterDeleteRecord        ; override ;
    procedure OnLoadRecord               ; override ;
    procedure OnChangeField ( F: TField) ; override ;
    procedure OnArgument ( S: String )   ; override ;
    procedure OnClose                    ; override ;
    procedure OnCancelRecord             ; override ;

  private
    FStPlanRevision    : string;

    FStSupervise       : string;
    FStPretSupervi     : string;


    FInActivePage      : integer;
    FInLeMode          : integer;

    FBoLoading         : Boolean;
  public
    Pages              : THPageControl2;
    PAPG               : TTabSheet;
    PSYN               : TTabSheet;
    PEXP               : TTabSheet;

    BHistoAPG          : TToolBarButton97;
    BSynthese          : TToolBarButton97;
    BEdition           : TToolBarButton97;

  {$IFDEF EAGLCLIENT}
    CIR_PLANREVISION   : THEdit;
    CIR_BLOQUECREATION : THCheckBox;
    CIR_PRETREVISER    : TCheckBox;
    CIR_PRETSUPERVI    : TCheckBox;
    CIR_SUPERVISE      : TCheckBox;
    CIR_BLOCNOTE       : THRichEditOle;
    CIR_ATTESTATION    : THValComboBox;
  {$ELSE}
    CIR_PLANREVISION   : THDBEdit;
    CIR_BLOQUECREATION : THDBCheckBox;
    CIR_PRETREVISER    : THDBCheckBox;
    CIR_PRETSUPERVI    : THDBCheckBox;
    CIR_SUPERVISE      : THDBCheckBox;
    CIR_BLOCNOTE       : THDBRichEditOle;
    CIR_ATTESTATION    : THDBValComboBox;
  {$ENDIF}
  
    MemoAPG            : THRichEditOle;
    MemoSYN            : THRichEditOle;

    CBECRANACCUEILN2   : THValComboBox;
    CBECRANACCUEILN3   : THValComboBox;
    CBECRANACCUEILN4   : THValComboBox;
    
    TCIR_NBATRAITER    : THLabel;
    TCIR_NBENSUSPENS   : THLabel;
    TCIR_NBSUPERVISE   : THLabel;
    TCIR_NBVALIDE      : THLabel;

    //CIR_PRETREVISERPAR : THValComboBox
    //CIR_PRETREVISERLE  : THEdit;

   procedure OnClickPretReviser     ( Sender : TObject );
   procedure OnClickPretSupervise   ( Sender : TObject );
   procedure OnClickSupervise       ( Sender : TObject );
   procedure OnClickBSynthese       ( Sender : TObject );
   procedure OnClickBHistoAPG       ( Sender : TObject );
   procedure OnClickBEdition        ( Sender : TObject );
   procedure PasseEnModeEdition     ( Sender : TObject );

   procedure ActiveLignePretReviser;
   procedure ActiveLignePretSuperviser;
   procedure ActiveLigneSupervise;

   procedure ChargeMemoAPG;
   procedure ChargeMemoSYN;

   procedure SauveMemoAPG;
   procedure SauveMemoSYN;

  end ;

////////////////////////////////////////////////////////////////////////////////
function CPLanceFiche_CPREVINFODOSSIER( vStNoDossier, vStExercice : string;
         vInLeMode : Integer; vStActionFiche : TActionFiche;
         vInActivePage : integer = -1 ) : string;
////////////////////////////////////////////////////////////////////////////////

Implementation

uses
  {$IFDEF MODENT1}
  CPTypeCons,
  {$ENDIF MODENT1}
  uLibWindows,         // IIF
  uTofHistoCpte,       // CC_LanceFicheHistoCpte( '', 'APG' );
  uFImgListe,          // FImgListe
  ParamSoc,            // SetParamSoc,
  uLibRevision,        // TAligneRevision
  AGLInit,             // ActionToStr
  Ed_Tools,            // InitMoveProgressForm
  Ent1;                // JaiLeRoleCompta

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 15/03/2007
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
//function CPLanceFiche_CREVNOTESUPERVISION( vStActionFiche : TActionFiche ) : string;
//begin
//  Result := AGLLanceFiche('CP', 'CPREVINFODOSSIER', V_PGI.NoDossier + ';' + VH^.Encours.Code, '' , ActionToString(vStActionFiche) + ';X');
//end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 15/03/2007
Modifié le ... :   /  /
Description .. : vInLeMode : 0 -> normal
Suite ........ :             1 -> Note de supervision
Suite ........ :             2 -> Révision du dossier
Mots clefs ... :
*****************************************************************}
function CPLanceFiche_CPREVINFODOSSIER( vStNoDossier, vStExercice : string;
         vInLeMode : Integer; vStActionFiche : TActionFiche;
         vInActivePage : integer = -1 ) : string;
var lTob : Tob;
begin
  if vStExercice = '' then
  begin
    vStExercice := GetColonneSQL('DB' + VStNoDossier + '.dbo.EXERCICE',
                   'EX_EXERCICE', 'EX_ETATCPTA = "OUV" ORDER BY EX_DATEDEBUT ASC');
  end;

  if not ExisteSQL('SELECT CIR_NODOSSIER FROM CREVINFODOSSIER WHERE ' +
                   'CIR_NODOSSIER = "' + vStNoDossier + '" AND ' +
                   'CIR_EXERCICE = "' + vStExercice + '"') then
  begin
    lTob := Tob.Create('CREVINFODOSSIER', nil, -1 );
    lTob.PutValue('CIR_NODOSSIER', vStNoDossier);
    lTob.PutValue('CIR_EXERCICE', vStExercice);
    lTob.PutValue('CIR_PREDEFINI', 'CEG');
    lTob.PutValue('CIR_ATTESTATION', cEnAttente);
    lTob.PutValue('CIR_PLANREVISION', VH^.Revision.Plan);
    lTob.PutValue('CIR_BLOQUECREATION', BoolToStr_(GetParamSocSecur('SO_CPREVISBLOQUEGENE', False)));
    lTob.InsertDB(nil);
    FreeAndNil(lTob);
  end
  else
    ReCalculInfoRevision;

  Result := AGLLanceFiche('CP', 'CPREVINFODOSSIER', '', vStNoDossier + ';' +
              vStExercice, ActionToString(vStActionFiche) + ';' +
              IntToStr( vInLeMode ) + ';' +
              IntToStr(vInActivePage));
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 15/03/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOM_CREVINFODOSSIER.OnArgument ( S: String ) ;
begin
  Inherited ;
  FBoLoading := True;

  Pages := THPageControl2(GetControl('PAGES', True));
  PAPG  := TTabSheet(GetControl('PAPG', True));
  PSYN  := TTabSheet(GetControl('PSYN', True));
  PEXP  := TTabSheet(GetControl('PEXP', True));

{$IFDEF EAGLCLIENT}
  CIR_PLANREVISION   := THEdit(GetControl('CIR_PLANREVISION', True));
  CIR_BLOQUECREATION := THCheckBox(GetControl('CIR_BLOQUECREATION', True));
  CIR_PRETREVISER    := TCheckBox(GetControl('CIR_PRETREVISER', True));
  CIR_PRETSUPERVI    := TCheckBox(GetControl('CIR_PRETSUPERVI', True));
  CIR_SUPERVISE      := TCheckBox(GetControl('CIR_SUPERVISE', True));
  CIR_BLOCNOTE       := THRichEditOle(GetControl('CIR_BLOCNOTE', True));
  CIR_ATTESTATION    := THValComboBox(Getcontrol('CIR_ATTESTATION', True));
{$ELSE}
  CIR_PLANREVISION   := THDBEdit(GetControl('CIR_PLANREVISION', True));
  CIR_BLOQUECREATION := THDBCheckBox(GetControl('CIR_BLOQUECREATION', True));
  CIR_PRETREVISER    := THDBCheckBox(GetControl('CIR_PRETREVISER', True));
  CIR_PRETSUPERVI    := THDBCheckBox(GetControl('CIR_PRETSUPERVI', True));
  CIR_SUPERVISE      := THDBCheckBox(GetControl('CIR_SUPERVISE', True));
  CIR_BLOCNOTE       := THDBRichEditOle(GetControl('CIR_BLOCNOTE', True));
  CIR_ATTESTATION    := THDBValComboBox(Getcontrol('CIR_ATTESTATION', True));
{$ENDIF}

  MemoAPG          := THRichEditOle(GetControl('MEMOAPG', True));
  MemoSYN          := THRichEditOle(GetControl('MEMOSYN', True));

  CBEcranAccueilN2 := THValComboBox(GetControl('CBECRANACCUEILN2', True));
  CBEcranAccueilN3 := THValComboBox(GetControl('CBECRANACCUEILN3', True));
  CBEcranAccueilN4 := THValComboBox(GetControl('CBECRANACCUEILN4', True));

  TCIR_NBATRAITER  := THLabel(GetControl('TCIR_NBATRAITER', True));
  TCIR_NBENSUSPENS := THLabel(GetControl('TCIR_NBENSUSPENS', True));
  TCIR_NBSUPERVISE := THLabel(GetControl('TCIR_NBSUPERVISE', True));
  TCIR_NBVALIDE    := THLabel(GetControl('TCIR_NBVALIDE', True));

  BHistoAPG        := TToolBarButton97(GetControl('BHISTOAPG', True));
  BSynthese        := TToolBarButton97(GetControl('BSYNTHESE', True));
  BEdition         := TToolBarButton97(GetControl('BEDITION', True));

  // Initialisation
  if (not V_Pgi.ModeTSE) then
  begin
    if FImgListe = nil then
      Application.CreateForm(TFImgListe, FImgListe);
    Pages.Images := FImgListe.BiblioActes;
  end;

  ReadTokenSt(S); // ACTION=
  FInLeMode          := StrToInt(ReadTokenSt(S));
  FInActivePage      := StrToInt(ReadTokenSt(S));

  // Branchement des événements
  MemoAPG.OnChange          := PasseEnModeEdition;
  MemoSyn.Onchange          := PasseEnModeEdition;
  CBEcranAccueilN2.OnChange := PasseEnModeEdition;
  CBEcranAccueilN3.OnChange := PasseEnModeEdition;
  CBEcranAccueilN4.OnChange := PasseEnModeEdition;

  CIR_PRETREVISER.OnClick := OnClickPretReviser;
  CIR_PRETSUPERVI.OnClick := OnClickPretSupervise;
  CIR_SUPERVISE.OnClick   := OnClickSupervise;

  BHistoAPG.OnClick := OnClickBHistoAPG;
  BSynthese.OnClick := OnClickBSynthese;
  BEdition.OnClick  := OnClickBEdition;
end ;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 15/03/2007
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOM_CREVINFODOSSIER.OnClose ;
begin
  // Plus de plan de révision, on supprime dans CREVINFODOSSIER
  if VH^.Revision.Plan = '' then
  begin
    ExecuteSQL('DELETE FROM CREVINFODOSSIER WHERE ' +
               'CIR_NODOSSIER = "' + V_Pgi.NoDossier + '" AND ' +
               'CIR_EXERCICE = "' + VH^.EnCours.Code + '"');

    SetParamSoc('SO_CPECRANACCUEILN2', '');
    SetParamSoc('SO_CPECRANACCUEILN3', '');
    SetParamSoc('SO_CPECRANACCUEILN4', '');
    AvertirCacheServer('PARAMSOC');
  end;

  FreeAndNil( FImgListe );
  Inherited ;
end ;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 15/03/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOM_CREVINFODOSSIER.OnLoadRecord ;
begin
  Inherited ;
  FStPlanRevision := GetField('CIR_PLANREVISION');
  FStSupervise    := GetField('CIR_SUPERVISE');
  FStPretSupervi  := GetField('CIR_PRETSUPERVI');

  TTabSheet(GetControl('PGENERAL', True)).TabVisible := FInLeMode <> 1;
  TTabSheet(GetControl('PGENERAL', True)).Visible    := FInLeMode <> 1;

  SetControlVisible('TLACCUEIL',          CtxPcl in V_Pgi.PGIContexte );
  SetControlVisible('TLCBECRANACCUEILN2', CtxPcl in V_Pgi.PGIContexte );
  SetControlVisible('CBECRANACCUEILN2',   CtxPcl in V_Pgi.PGIContexte );
  SetControlVisible('TLCBECRANACCUEILN3', CtxPcl in V_Pgi.PGIContexte );
  SetControlVisible('CBECRANACCUEILN3',   CtxPcl in V_Pgi.PGIContexte );
  SetControlVisible('TLCBECRANACCUEILN4', CtxPcl in V_Pgi.PGIContexte );
  SetControlVisible('CBECRANACCUEILN4',   CtxPcl in V_Pgi.PGIContexte );

  // GCO - 09/07/2007 - FQ 20852 - Grisé pour les collaborateurs
  if (CtxPcl in V_Pgi.PGIContexte) and (not (JaiLeRoleCompta( rcReviseur ))) then
  begin
    SetControlEnabled('CIR_PLANREVISION', False);
    SetControlEnabled('CIR_BLOQUECREATION', False);
    SetControlEnabled('CBECRANACCUEILN2', False);
    SetControlEnabled('CBECRANACCUEILN3', False);
    SetControlEnabled('CBECRANACCUEILN4', False);
  end;

  PAPG.TabVisible   := FInLeMode = 1;
  PSYN.TabVisible   := FInLeMode = 1;
  PEXP.TabVisible   := FInLeMode = 1;

  BHistoAPG.Visible := FInLeMode = 1;
  BSynthese.Visible := FInLeMode = 1;

  if FInLeMode = 2 then
  begin
    CIR_PLANREVISION.Enabled := False;
    CIR_BLOQUECREATION.Enabled := False;
    CBECRANACCUEILN2.Enabled := False;
    CBECRANACCUEILN3.Enabled := False;
    CBECRANACCUEILN4.Enabled := False;
  end;

  PEXP.ImageIndex := -1;
  if CIR_BLOCNOTE.Text <> '' then
  begin
    if V_Pgi.ModeTSE then
      PEXP.Caption := PEXP.Caption + ' (***)'
    else
      PEXP.ImageIndex := cImgBlocNote;
  end;

  if FInLeMode = 1 then
  begin
    TFFiche(Ecran).Caption := TraduireMemoire('Supervision des travaux');
    UpdateCaption( Ecran );
    ChargeMemoAPG; // Appréciation générale
    ChargeMemoSYN; // Synthèse des cycles

    if VH^.Revision.DossierSupervise then
    begin
      MemoAPG.ReadOnly      := True;
      MemoSyn.ReadOnly      := True;
      CIR_BLOCNOTE.ReadOnly := True;

      MemoAPG.Color      := ClBtnFace;
      MemoSyn.Color      := ClBtnFace;
      CIR_BLOCNOTE.Color := ClBtnFace;

      SetControlEnabled('CIR_ATTESTATION', False);
      SetControlEnabled('CIR_ATTESTATIONPAR', False);
      SetControlEnabled('CIR_ATTESTATIONLE', False);
      SetControlEnabled('CIR_CHOIXLE', False);
      BSynthese.Enabled := False;
    end;
  end
  else
  begin
    // GCO - 09/05/2007 - FQ 20275
    ActiveLignePretReviser;
    ActiveLignePretSuperviser;
    ActiveLigneSupervise;
   end;

  if FInActivePage >= 0 then
  begin
    Pages.ActivePageIndex := FInActivePage;
    case FInActivePage of
      1 : MemoAPG.SetFocus;
      2 : MemoSYN.SetFocus;
      3 : SetFocusControl('CIR_BLOCNOTE');
    end;
  end;

  // GCO - 31/05/2007 - FQ 20481 : FROMDISK à TRUE
  CBEcranAccueilN2.Value   := GetParamSocSecur('SO_CPECRANACCUEILN2', '', True);
  CBEcranAccueilN3.Value   := GetParamSocSecur('SO_CPECRANACCUEILN3', '', True);
  CBEcranAccueilN4.Value   := GetParamSocSecur('SO_CPECRANACCUEILN4', '', True);

  TCIR_NBATRAITER.Caption  := IntToStr(GetField('CIR_NBATRAITER')) + ' ' + TCIR_NBATRAITER.Caption;
  TCIR_NBENSUSPENS.Caption := IntToStr(GetField('CIR_NBENSUSPENS')) + ' ' + TCIR_NBENSUSPENS.Caption;

  TCIR_NBSUPERVISE.Caption := IntToStr(GetField('CIR_NBSUPERVISE')) + ' ' + TCIR_NBSUPERVISE.Caption;
  TCIR_NBVALIDE.Caption    := IntToStr(GetField('CIR_NBVALIDE')) + ' ' + TCIR_NBVALIDE.Caption;

  if GetField('CIR_ATTESTATION') <> cEnAttente then
  begin
    SetControlEnabled('CIR_ATTESTATIONLE', False );
    SetControlEnabled('CIR_CHOIXLE', False );
  end;

  FBoLoading := False;
end ;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 15/03/2007
Modifié le ... : 15/03/2007
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOM_CREVINFODOSSIER.OnNewRecord ;
begin
  Inherited ;
  SetField('CIR_NODOSSIER', V_Pgi.NoDossier);
  SetField('CIR_EXERCICE', VH^.EnCours.Code);
  SetField('CIR_PREDEFINI', 'CEG');
  SetField('CIR_ATTESTATION', 'ATT');
end ;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 15/03/2007
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOM_CREVINFODOSSIER.OnUpdateRecord ;
var lStPlanRevision : string;
begin
  Inherited ;

  if FInLeMode = 1 then
  begin
    if GetField('CIR_ATTESTATION') <> cEnAttente then
    begin
      if GetField('CIR_ATTESTATIONLE') <= VH^.EnCours.Deb then
      begin
        PgiInfo('La date d''attestation doit être supérieure au ' + DateToStr(VH^.EnCours.Deb) + '.', 'Enregistrement impossible'); 
        LastError := -1;
        Exit;
      end;

      if GetField('CIR_CHOIXLE') <= VH^.EnCours.Deb then
      begin
        PgiInfo('La date du choix effectué doit être supérieure au ' + DateToStr(VH^.EnCours.Deb) + '.', 'Enregistrement impossible');
        LastError := -1;
        Exit;
      end;

      if GetField('CIR_ATTESTATIONPAR') <> '' then
      begin
        if GetRoleComptaUtilisateur( GetField('CIR_ATTESTATIONPAR') ) <> RcExpert then
        begin
          PgiInfo('L''attestation ne peut être délivrée que par un expert comptable (Niveau N4).', 'Enregistrement impossible');
          LastError := -1;
          Exit;
        end;
      end
      else
      begin
        PgiInfo('Vous devez renqeigner l''utilisateur qui effectue l''attestation expert.', 'Enregistrement impossible');
        LastError := -1;
        Exit;
      end;
    end;

    SauveMemoAPG;
    SauveMemoSYN;
  end
  else
  begin
    lStPlanRevision := GetField('CIR_PLANREVISION');

    if lStPlanRevision <> FStPlanRevision then
    begin
      // Alignement du dossier sur le nouveau plan saisi.
      if lStPlanRevision <> '' then
      begin
        if PgiAsk('Attention ! Le plan de révision a été changé.' + #13#10 +
                  'Etes-vous sur de vouloir aligner le dossier sur ce plan de révision ?',
                  'Mise à jour de la révision') = MrNo then
        begin
          LastError := 1;
          Exit;
        end
        else
        begin
          if FStPlanRevision = '' then
            // Pas de plan au départ, on en affecte un
            AlignementPlanRevision(lStPlanRevision, True, False)
          else
            // Plan dejà associé, on en affecte un autre, on doit faire un supprime avant
            AlignementPlanRevision(lStPlanRevision, True, True);
          FStPlanRevision := GetField('CIR_PLANREVISION');
        end;
      end
      else
      begin
        if PgiAsk('Attention ! Le plan de révision a été supprimé.' + #13#10 +
                  'Voulez-vous effacer les informations de révision ?',
                  'Suppression de la révision') = MrNo then
        begin
          LastError := -1;
          Exit;
        end
        else
        begin
          SupprimeRevision(True);
          FStPlanRevision := '';
        end;
      end;
    end;

    // GCO - 01/08/2007 - FQ 21205
    try
      InitMoveProgressForm(nil, 'Alignement du plan de révision ' + FStPlanRevision,
      'Mise à jour des informations de révision en cours...', 1 , True, True);

      MoveCurProgressForm('Enregistrement des paramètres société...');

      if (FStPretSupervi <> GetField('CIR_PRETSUPERVI')) then
      begin
        if (GetField('CIR_PRETSUPERVI') = 'X') then
          TraiteToutLeDossier( cValide );
      end;

      // GCO - 26/07/2007 - FQ
      Sleep(1000);

      if (FStSupervise <> GetField('CIR_SUPERVISE')) then
      begin
        if (GetField('CIR_SUPERVISE') = 'X') then
          TraiteToutLeDossier( cSupervise );
      end;

      FStSupervise   := GetField('CIR_SUPERVISE');
      FStPretSupervi := GetField('CIR_PRETSUPERVI');

      // Mise à jour du paramsoc de blocage des fonctions de révision, etc...
      VH^.Revision.DossierSupervise     := (GetField('CIR_SUPERVISE') = 'X');
      VH^.REvision.DossierPretSupervise := (GetField('CIR_PRETSUPERVI') = 'X');

      SetParamSoc('SO_CPREVISBLOQUEGENE', GetField('CIR_BLOQUECREATION'));
      SetParamSoc('SO_CPECRANACCUEILN2', THValComboBox(GetControl('CBECRANACCUEILN2', True)).Value);
      SetParamSoc('SO_CPECRANACCUEILN3', THValComboBox(GetControl('CBECRANACCUEILN3', True)).Value);
      SetParamSoc('SO_CPECRANACCUEILN4', THValComboBox(GetControl('CBECRANACCUEILN4', True)).Value);
    finally
      FiniMoveProgressForm;
    end;  
  end;
end ;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 15/03/2007
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOM_CREVINFODOSSIER.OnAfterUpdateRecord ;
begin
  Inherited ;
end ;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 15/03/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOM_CREVINFODOSSIER.OnDeleteRecord ;
begin
  Inherited ;
end ;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 15/03/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOM_CREVINFODOSSIER.OnAfterDeleteRecord ;
begin
  Inherited ;
end ;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 15/03/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOM_CREVINFODOSSIER.OnChangeField ( F: TField ) ;
begin
  Inherited ;

  if F.FieldName = 'CIR_PLANREVISION' then
  begin
    if F.Value <> '' then
      SetControlText('LIBPLANREVISION', GetColonneSQL('CREVPLAN', 'CPR_LIBELLE', 'CPR_PLANREVISION = "' + F.Value + '"'))
    else
      SetControlText('LIBPLANREVISION', '');
  end;

  if F.FieldName = 'CIR_PRETREVISER' then
  begin
    if FBoLoading then Exit;
    if F.AsString = 'X' then
    begin
      if GetField('CIR_PRETREVISERPAR') = '' then SetField('CIR_PRETREVISERPAR', V_Pgi.User);
      if GetField('CIR_PRETREVISERLE') = idate1900 then SetField('CIR_PRETREVISERLE', Date);
     end
    else
    begin
      SetField('CIR_PRETREVISERPAR', '');
      SetField('CIR_PRETREVISERLE',  iDate1900);
    end;
  end;

  if F.FieldName = 'CIR_PRETSUPERVI' then
  begin
    if FBoLoading then Exit;
    if F.AsString = 'X' then
    begin
      if GetField('CIR_PRETSUPERVIPAR') = '' then SetField('CIR_PRETSUPERVIPAR', V_Pgi.User);
      if GetField('CIR_PRETSUPERVILE') = idate1900 then SetField('CIR_PRETSUPERVILE', Date);
      if GetField('CIR_PRETREVISER') = '-' then SetField('CIR_PRETREVISER'  , 'X');
     end
    else
    begin
      SetField('CIR_PRETSUPERVIPAR', '');
      SetField('CIR_PRETSUPERVILE',  iDate1900);
    end;
  end;

  if F.FieldName = 'CIR_SUPERVISE' then
  begin
    if FBoLoading then Exit;
    if F.AsString = 'X' then
    begin
      if GetField('CIR_SUPERVISEPAR') = '' then SetField( 'CIR_SUPERVISEPAR', V_Pgi.User);
      if GetField('CIR_SUPERVISELE') = idate1900 then SetField( 'CIR_SUPERVISELE', Date);
      if GetField('CIR_PRETREVISER') = '-' then SetField('CIR_PRETREVISER'  , 'X');
      if GetField('CIR_PRETSUPERVI') = '-' then SetField('CIR_PRETSUPERVI'   , 'X');
    end
    else
    begin
      SetField( 'CIR_SUPERVISEPAR', '');
      SetField( 'CIR_SUPERVISELE',  iDate1900);
    end;
  end;

  if F.FieldName = 'CIR_ATTESTATION' then
  begin
    if F.AsString = cEnAttente then
    begin
      SetControlEnabled('CIR_ATTESTATIONLE', True );
      SetControlEnabled('CIR_CHOIXLE', True );
      if not FBoLoading then
      begin
        SetField('CIR_ATTESTATIONLE', DateToStr(idate1900));
        SetField('CIR_CHOIXLE', DateToStr(idate1900));
      end;
    end
    else
    begin
      if not FBoLoading then
      begin
        SetField('CIR_ATTESTATIONLE', DateToStr(Date));
        SetField('CIR_CHOIXLE', DateToStr(Date));
      end;
    end;
  end;

end ;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 15/03/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOM_CREVINFODOSSIER.OnCancelRecord ;
begin
  Inherited ;
end ;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 15/03/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... :
*****************************************************************}
procedure TOM_CREVINFODOSSIER.OnClickPretReviser(Sender: TObject);
begin
  ActiveLignePretReviser;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 15/03/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOM_CREVINFODOSSIER.ActiveLignePretReviser;
begin
  // GCO - 21/05/2007 - FQ 20329
  SetControlEnabled('CIR_PRETREVISER', JaiLeRoleCompta( rcCollaborateur ) and (not CIR_PRETSUPERVI.Checked));
  if (CIR_PRETREVISER.Enabled) then
  begin
    SetControlEnabled('CIR_PRETREVISERPAR', (not FBoLoading) and (CIR_PRETREVISER.Checked));
    SetControlEnabled('CIR_PRETREVISERLE',  (not FBoLoading) and (CIR_PRETREVISER.Checked));

    if not FBoLoading then
      SetField('CIR_PRETREVISER', IIF(CIR_PRETREVISER.Checked, 'X', '-'));
  end
  else
  begin
    SetControlEnabled('CIR_PRETREVISERPAR', False);
    SetControlEnabled('CIR_PRETREVISERLE',  False);
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 15/03/2007
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOM_CREVINFODOSSIER.OnClickPretSupervise(Sender: TObject);
begin
  ActiveLignePretSuperviser;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 15/03/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOM_CREVINFODOSSIER.ActiveLignePretSuperviser;
begin
  SetControlEnabled('CIR_PRETSUPERVI', JaiLeRoleCompta( rcReviseur ) and (not CIR_SUPERVISE.Checked));

  if (CIR_PRETSUPERVI.Enabled) then
  begin
    SetControlEnabled('CIR_PRETSUPERVIPAR', (not FBoLoading) and (CIR_PRETSUPERVI.Checked));
    SetControlEnabled('CIR_PRETSUPERVILE',  (not FBoLoading) and (CIR_PRETSUPERVI.Checked));

    if not FBoLoading then
      SetField('CIR_PRETSUPERVI', IIF(CIR_PRETSUPERVI.Checked, 'X', '-'));

    // GCO - 15/10/2007 - FQ 21649
    if CIR_PRETSUPERVI.Checked then
    begin
      SetControlEnabled('CIR_PLANREVISION',   False);
      SetControlEnabled('CIR_BLOQUECREATION', False);
    end
    else
    begin
      SetControlEnabled('CIR_PLANREVISION',  FInLeMode <> 2); // GCO - 28/05/2007
      SetControlEnabled('CIR_BLOQUECREATION',FInLeMode <> 2); // GCO - 28/05/2007
    end;
  end
  else
  begin
    SetControlEnabled('CIR_PRETSUPERVIPAR', False);
    SetControlEnabled('CIR_PRETSUPERVILE',  False);
  end;
  ActiveLignePretReviser;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 15/03/2007
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOM_CREVINFODOSSIER.OnClickSupervise(Sender: TObject);
begin
  ActiveLigneSupervise;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 15/03/2007
Modifié le ... :   /  /
Description .. :
Mots clefs ... : 
*****************************************************************}
procedure TOM_CREVINFODOSSIER.ActiveLigneSupervise;
begin
  SetControlEnabled('CIR_SUPERVISE', JaiLeRoleCompta( rcSuperviseur ));

  if CIR_SUPERVISE.Enabled then
  begin
    SetControlEnabled('CIR_SUPERVISEPAR', (not FBoLoading) and (CIR_SUPERVISE.Checked));
    SetControlEnabled('CIR_SUPERVISELE',  (not FBoLoading) and (CIR_SUPERVISE.Checked));

    if not FBoLoading then
      SetField('CIR_SUPERVISE', IIF(CIR_SUPERVISE.Checked, 'X', '-'));

    if CIR_SUPERVISE.Checked then
    begin
      SetControlEnabled('CIR_PLANREVISION',   False);
      SetControlEnabled('CIR_BLOQUECREATION', False);
    end
    else
    begin
      SetControlEnabled('CIR_SUPERVISEPAR',  False);
      SetControlEnabled('CIR_SUPERVISELE',   False);
      SetControlEnabled('CIR_PLANREVISION',  FInLeMode <> 2); // GCO - 28/05/2007
      SetControlEnabled('CIR_BLOQUECREATION',FInLeMode <> 2); // GCO - 28/05/2007 
    end;
  end
  else
  begin
    SetControlEnabled('CIR_SUPERVISEPAR', False);
    SetControlEnabled('CIR_SUPERVISELE',  False);
  end;

  ActiveLignePretSuperviser;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 20/03/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOM_CREVINFODOSSIER.OnClickBSynthese(Sender: TObject);
var i : integer;
    lTob : Tob;
    lSt  : string;
    Memo : THRichEditOle;
begin
  Memo := THRichEditOle(GetControl('MEMOTEMP', True));

  lTob := Tob.Create('', nil, -1);
  lSt := 'SELECT CBN_CODE, CPC_LIBELLECYCLE, CBN_BLOCNOTE FROM CREVBLOCNOTE ' +
  'LEFT JOIN CREVCYCLE ON CCY_CODECYCLE = CBN_CODE AND CCY_EXERCICE = CBN_EXERCICE ' +
  'LEFT JOIN CREVPARAMCYCLE ON CPC_CODECYCLE = CBN_CODE WHERE ' +
  'CBN_NATURE = "SCY" AND CCY_EXERCICE = "' + VH^.EnCours.Code + '" AND ' +
  'CCY_ACTIVECYCLE = "X" ' +
  // GCO -13/07/2007 - FQ 21074 et
  'AND CPC_PLANASSOCIE LIKE "%' + VH^.Revision.Plan + ';%" ' +
  'ORDER BY CBN_CODE';

  lTob.LoadDetailFromSQL(lSt , False, True);

  MemoSyn.Clear;
  for i := 0 to lTob.Detail.Count -1 do
  begin
    MemoSyn.lines.Add(lTob.Detail[i].GetString('CBN_CODE') + ' - ' +
                      lTob.Detail[i].GetString('CPC_LIBELLECYCLE'));

    StringToRich(Memo, lTob.Detail[i].GetString('CBN_BLOCNOTE'));
    MemoSyn.lines.Add('----------');
    // GCO - 31/08/2007 - FQ 21235
    // MemoSyn.Lines.Add(Memo.Lines.Text);
    MemoSyn.Lines.Add(Memo.Text);
    MemoSyn.LinesRTF.Add(Memo.LinesRTF.Text);
    MemoSyn.lines.Add('');
  end;

  FreeAndNil( lTob );
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 29/05/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOM_CREVINFODOSSIER.OnClickBEdition(Sender: TObject);
begin
  // GCO - 25/09/2007 - FQ 21196
  if DS.State <> dsBrowse then
  begin
    case PgiAskCancel('Voulez-vous enregistrer les modifications ?', Ecran.Caption) of
      mrNo                : ;
      mrCancel            : Exit;
      mrYes               : TToolBarbutton97(GetControl('BVALIDER', True)).Click;
    end;
  end;  
  TRic.CPLanceEtat_NoteSupervision( 0 );
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 21/03/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOM_CREVINFODOSSIER.OnClickBHistoAPG(Sender: TObject);
begin
  CC_LanceFicheHistoCpte( '', 'APG' );
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 20/03/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOM_CREVINFODOSSIER.ChargeMemoAPG;
begin
  MemoAPG.Clear;
  StringToRich(MemoAPG, ChargeMemoCREVBLOCNOTE('APG', '', VH^.EnCours.Code));
  // Obligation de remettre à FALSE après chargement
  MemoAPG.Modified := False;

  PAPG.ImageIndex := -1;
  if MemoAPG.Text <> '' then
  begin
    if V_Pgi.ModeTSE then
      PAPG.Caption := PAPG.Caption + ' (***)'
    else
      PAPG.ImageIndex := cImgBlocnote;
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 28/03/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOM_CREVINFODOSSIER.PasseEnModeEdition(Sender: TObject);
begin
  if fBoLoading then Exit;
  if DS.State = dsBrowse then DS.Edit;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 20/03/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOM_CREVINFODOSSIER.ChargeMemoSYN;
begin
  MemoSyn.Clear;
  StringToRich(MemoSyn, ChargeMemoCREVBLOCNOTE('SYN', '', VH^.EnCours.Code));
  // Obligation de remettre à FALSE après chargement
  MemoSyn.Modified := False;

  PSYN.ImageIndex := -1;
  if MemoSyn.Text <> '' then
  begin
    if V_Pgi.ModeTSE then
      PSYN.Caption := PSYN.Caption + ' (***)'
    else
      PSYN.ImageIndex := cImgBlocNote;
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 20/03/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOM_CREVINFODOSSIER.SauveMemoAPG;
begin
  if Trim(MemoAPG.Text) = '' then         
    ExecuteSQL('DELETE FROM CREVBLOCNOTE WHERE ' +
               'CBN_NATURE = "APG" AND ' +
               'CBN_CODE = "" AND ' +
               'CBN_EXERCICE = "' + VH^.EnCours.Code + '"')
  else
    SauveMemoCREVBLOCNOTE('APG', '', VH^.EnCours.Code, RichToString(MemoAPG) );
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 20/03/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOM_CREVINFODOSSIER.SauveMemoSYN;
begin
  if Trim(MemoSYN.Text) = '' then
    ExecuteSQL('DELETE FROM CREVBLOCNOTE WHERE ' +
               'CBN_NATURE = "SYN" AND ' +
               'CBN_CODE = "" AND ' +
               'CBN_EXERCICE = "' + VH^.EnCours.Code + '"')
  else
    SauveMemoCREVBLOCNOTE('SYN', '', VH^.EnCours.Code, RichToString(MemoSYN) );
end;

///////////////////////////////////////////////////////////////////////////////


Initialization
  registerclasses ( [ TOM_CREVINFODOSSIER ] ) ;
end.
