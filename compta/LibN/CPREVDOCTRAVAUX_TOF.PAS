{***********UNITE*************************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 19/12/2006
Modifié le ... : 24/01/2007
Description .. : Source TOF de la FICHE : CPREVDOCTRAVAUX ()
Mots clefs ... : TOF;CPREVDOCTRAVAUX
*****************************************************************}
Unit CPREVDOCTRAVAUX_TOF;

Interface

Uses StdCtrls, 
     Controls,
     Classes, 
{$IFDEF EAGLCLIENT}
     eMul,
     MainEagl,      // AGLLanceFiche
{$ELSE}
     db,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul,
     FE_main,       // AGLLanceFiche
{$ENDIF}
     forms,
     sysutils, 
     ComCtrls,
     HCtrls,
     HEnt1, 
     HMsgBox, 
     uTob,
     Windows,      // VK
     Menus,        // TPopUpMenu
     Ent1,         // VH , JaiLeRoleCompta
     HPanel,       // THPanel
     Htb97,        // TToolbarButton97
     HRichOle,     // THRichEditOle
     Graphics,     // ClBtnFace
     UTOF ;

Type
  TOF_CPREVDOCTRAVAUX = Class (TOF)

    PanelCycle       : THPanel;
    PanelGeneral     : THPanel;

    Pages            : THPageControl2;

    TabMemocycle     : TTabSheet;
    TabMemoObjectif  : TTabSheet;
    TabMemoSynthese  : TTabSheet;
    TabMemoMillesime : TTabSheet;
    TabMemoCompte    : TTabSheet;

    TLPOURCENTVISE   : THLabel;
    TLPARLE          : THLabel;
    CYCLE            : THLabel;
    GENERAL          : THLabel;
    CCY_ETATCYCLEPAR : THLabel;
    DATEREVISION     : THLabel;

    CCY_ETATCYCLE    : THValComboBox;

    CCY_ETATCYCLELE  : THEdit;

    VISAREVISION     : TCheckBox;

    MemoCycle        : THRichEditOle;
    MemoObjectif     : THRichEditOle;
    MemoSynthese     : THRichEditOle;
    MemoMillesime    : THRichEditOle;
    MemoCompte       : THRichEditOle;

    BHistoMemo       : TToolBarButton97;
    BHistoCycle      : TToolBarButton97;

    BCycle           : TToolBarButton97;
    BEdition         : TToolBarButton97;
    BSupervisTravaux : TToolBarButton97;
    BValider         : TToolBarButton97;

    PopEdition           : TPopUpMenu;
    PopUpSupervisTravaux : TPopUpMenu;
    PopF11               : TPopUpMenu;

    procedure OnArgument (S : String ) ; override ;
    procedure OnClose                  ; override ;
    procedure OnLoad                   ; override ;
    procedure OnDisplay                ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnCancel                 ; override ;

    procedure OnChangeCCY_Etatcycle  (Sender : TObject);
    procedure OnClickBHistoMemo      (Sender : TObject);
    procedure OnClickBHistoCycle     (Sender : TObject);
    procedure OnClickBCycle          (Sender : TObject);
    procedure OnClickNoteCtrlCompte  (Sender : TObject);
    procedure OnClickNoteCtrlCycle   (Sender : TObject);
    procedure OnClickNoteSupervisionDossier (Sender : TObject); 
    procedure OnClickVisaRevision    (Sender : TObject);
    procedure OnClickMenuVisaRevision(Sender : TObject);

    procedure OnClickAPG             ( Sender : TObject ); // Appréciation générale
    procedure OnClickSCY             ( Sender : TObject ); // Synthèse des cycles
    procedure OnClickEXP             ( Sender : TObject ); // Attestation Expert

    procedure OnPopUpPopF11          (Sender : TObject);
    procedure OnPopUpPopEdition      (Sender : TObject);
    procedure OnKeyDownEcran         (Sender : TObject; var Key: Word; Shift: TShiftState);

  private
    FStGeneral       : string;
    FStExercice      : string;
    FStCycle         : string;
    FStPredefini     : string; // GCO - 18/09/2007 - FQ 21433
    FStEtatCycle     : string;
    FStEtatCycleLe   : string; 
    FInActivePage    : Integer;

    FBoLoading       : Boolean;

    fOnSaveKeyDownEcran : procedure(Sender: TObject; var Key: Word; Shift: TShiftState) of object;

    procedure ChargeInfoGeneral;
    procedure ChargeInfoCycle;
    procedure ChargeMemoCommentaireCycle;
    procedure ChargeMemoSyntheseCycle;
    procedure ChargeMemoOBjectif;
    procedure ChargeMemoMillesime;

    procedure SauveInfoCycle;
    procedure SauveMemoGeneral;
    procedure SauveMemoCommentaireCycle;
    procedure SauveMemoSyntheseCycle;
    procedure SauveMemoMillesime;
    procedure InitContexte;

    function  SauveAvantEdition : integer;

  end ;

function CPLanceFiche_CPREVDOCTRAVAUX( vStGeneral, vStCycle, vStExercice : string; vInActivePage : integer = -1 ) : string;

Implementation

uses
{$IFDEF MODENT1}
  CPTypeCons,
{$ENDIF MODENT1}
{$IFDEF COMPTA}
{$IFNDEF CCMP}
     uLibRevision,        // ChangeEtatCycle
     CRevParamCycle_TOM,  // CPLanceFiche_CREVPARAMCYCLE
     CRevInfoDossier_TOM, // CPLanceFiche_CREVINFODOSSIER
     CPRevHistoCycle_TOF, // CPLanceFiche_CPREVHISTOCYCLE
{$ENDIF}
{$ENDIF}
     uFImgListe,          // FImgListe
     uLibWindows,         // IIF
     uTOFHistoCpte;       // ChargeMemo, SauveMemo

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 02/01/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
function CPLanceFiche_CPREVDOCTRAVAUX( vStGeneral, vStCycle, vStExercice : string; vInActivePage : integer = -1 ) : string;
begin
  Result := AGLLanceFiche('CP', 'CPREVDOCTRAVAUX', '', '', vStGeneral + ';' + vStCycle + ';' +
            vStExercice + ';' + IntToStr(vInActivePage));
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 19/12/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPREVDOCTRAVAUX.OnArgument (S : String ) ;
begin
  Inherited ;
  FBoLoading       := True;

  PanelCycle       := THPanel(GetControl('PANELCYCLE', True));
  PanelGeneral     := THPanel(Getcontrol('PANELGENERAL', True));

  CYCLE            := THLabel(GetControl('CYCLE', True));
  DATEREVISION     := THLabel(GetControl('DATEREVISION', True));
  GENERAL          := THLabel(GetControl('GENERAL', True));
  CCY_ETATCYCLEPAR := THLabel(GetControl('CCY_ETATCYCLEPAR', True));

  CCY_ETATCYCLE    := THValComboBox(GetControl('CCY_ETATCYCLE', True));
  CCY_ETATCYCLELE  := THEdit(GetControl('CCY_ETATCYCLELE', True));

  VISAREVISION     := TCheckBox(GetControl('VISAREVISION', True));

  TLPOURCENTVISE   := THLabel(GetControl('TLPOURCENTVISE', True));
  TLPARLE          := THLabel(GetControl('TLPARLE', True));

  Pages            := THPageControl2(GetControl('PAGES', True));

  TabMemoCycle     := TTabSheet(GetControl('TABMEMOCYCLE', True));
  TabMemoObjectif  := TTabSheet(GetControl('TABMEMOOBJECTIF', True));
  TabMemoSynthese  := TTabSheet(GetControl('TABMEMOSYNTHESE', True));
  TabMemoMillesime := TTabSheet(GetControl('TABMEMOMILLESIME', True));;
  TabMemoCompte    := TTabSheet(GetControl('TABMEMOCOMPTE', True));

  MemoCycle        := THRichEditOle(GetControl('MEMOCYCLE', True));
  MemoObjectif     := THRichEditOle(GetControl('MEMOOBJECTIF', True));
  MemoSynthese     := THRichEditOle(GetControl('MEMOSYNTHESE', True));
  MemoMillesime    := THRichEditOle(GetControl('MEMOMILLESIME', True));
  MemoCompte       := THRichEditOle(GetControl('MEMOCOMPTE', True));


  BHistoMemo       := TToolBarButton97(GetControl('BHISTOMEMO', True));
  BHistoCycle      := TToolBarButton97(GetControl('BHISTOCYCLE', True));
  BCycle           := TToolBarButton97(GetControl('BCYCLE', True));
  BEdition         := TToolBarButton97(GetControl('BEDITION', True));
  BSupervisTravaux := TToolBarButton97(GetControl('BSUPERVISTRAVAUX', True));
  BValider         := TToolBarButton97(GetControl('BVALIDER', True));

  PopUpSupervisTravaux := TPopUpMenu(GetControl('POPUPSUPERVISTRAVAUX', True));
  PopEdition           := TPopUpMenu(GetControl('POPEDITION', True));
  PopF11               := TPopUpMenu(GetControl('POPF11', True));

  // Evénements des composants
  CCY_EtatCycle.OnChange  := OnChangeCCY_EtatCycle;

  BHistoMemo.OnClick      := OnClickBHistoMemo;
  BHistoCycle.OnClick     := OnClickBHistoCycle;
  BCycle.OnClick          := OnClickBCycle;
  VisaRevision.OnClick    := OnClickVisaRevision;

  fOnSaveKeyDownEcran     := Ecran.OnKeyDown;
  Ecran.OnKeyDown         := OnKeyDownEcran;

  PopEdition.Items[0].OnClick := OnClickNoteCtrlCompte;
  PopEdition.Items[1].OnClick := OnClickNoteCtrlCycle;
  PopEdition.Items[3].OnClick := OnClickNoteSupervisionDossier;
  PopEdition.OnPopUp          := OnPopUpPopEdition;

  PopUpSupervisTravaux.Items[0].OnClick := OnClickAPG;
  PopUpSupervisTravaux.Items[1].OnClick := OnClickSCY;
  PopUpSupervisTravaux.Items[2].OnClick := OnClickEXP;

  PopF11.OnPopUp          := OnPopUpPopF11;
  PopF11.Items[0].OnClick := OnClickMenuVisaRevision;
  PopF11.Items[1].OnClick := OnClickBHistoMemo;
  PopF11.Items[2].OnClick := OnClickBHistoCycle;
  PopF11.Items[3].OnClick := OnClickBCycle;
  PopF11.Items[4].Items[0].OnCLick := OnClickAPG;
  PopF11.Items[4].Items[1].OnCLick := OnClickSCY;
  PopF11.Items[4].Items[2].OnCLick := OnClickEXP;

  PopF11.Items[6].Items[0].OnClick := OnClickNoteCtrlCompte;
  PopF11.Items[6].Items[1].OnClick := OnClickNoteCtrlCycle;
  PopF11.Items[6].Items[3].OnClick := OnClickNoteSupervisionDossier;

  // Initialisation
  if (not V_Pgi.ModeTSE) then
  begin
    if FImgListe = nil then
      Application.CreateForm(TFImgListe, FImgListe);
    Pages.Images := FImgListe.BiblioActes;
  end;  

  FStGeneral    := ReadTokenSt(S);
  FStCycle      := ReadTokenSt(S);
  FStExercice   := ReadTokenSt(S);
  FInActivePage := StrToInt(ReadTokenSt(S));

  DateRevision.Caption := DateToStr(VH^.EnCours.Fin);

  // GCO - 11/06/2007 - FQ 20631 
  if FStGeneral <> '' then
    BHistoMemo.Hint := TraduireMemoire('Historique du compte')
  else
    BHistoMemo.Hint := TraduireMemoire('Historique des commentaires');

  // GCO - 04/10/2007
  ActivateXpPopUp( PopF11 );
end ;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 19/12/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPREVDOCTRAVAUX.OnClose ;
begin
  FreeAndNil( FImgListe );
  Inherited ;
end ;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 19/12/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPREVDOCTRAVAUX.OnLoad ;
begin
  Inherited ;
  if FStGeneral <> '' then
  begin
    ChargeInfoGeneral;
    ChargeMemoMillesime;
  end;

  if FStCycle <> '' then
  begin
    // GCO - 18/09/2007 - FQ 21433
    FStPredefini := GetColonneSQL('CREVCYCLE', 'CCY_PREDEFINI', 'CCY_CODECYCLE = "' + FStCycle + '"');
    ChargeInfoCycle;
    ChargeMemoCommentaireCycle;
    ChargeMemoSyntheseCycle;
    ChargeMemoOBjectif;
  end
  else
    FStPredefini := '';

  InitContexte;
  FBoLoading := False;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 26/02/2007
Modifié le ... :   /  /
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPREVDOCTRAVAUX.InitContexte;
begin
  PanelCycle.Visible          := (FStCycle <> '');
  PanelGeneral.Visible        := (FStGeneral <> '');

  TabMemoCycle.TabVisible     := (FStCycle <> '') and (VH^.OkModRIC) and JaiLeRoleCompta( RcReviseur );

  // GCO - 05/09/2007 - FQ 21369
  TabMemoObjectif.TabVisible  := (CtxPcl In V_Pgi.PgiContexte) and (FStCycle <> '') and (VH^.OkModRIC) and JaiLeRoleCompta( RcReviseur );
  TabMemoSynthese.TabVisible  := (FStCycle <> '') and (VH^.OkModRIC) and JaiLeRoleCompta( RcReviseur );
  TabMemoMillesime.TabVisible := (FStGeneral <> '');
  TabMemoCompte.TabVisible    := (FStGeneral <> '');

  VisaRevision.Enabled        := (FStGeneral <> '') and (AutoriseSuppresionVisaRevision(FStGeneral));

  BHistoMemo.Visible          := (FStCycle <> '') and (VH^.OkModRIC);
  BHistoCycle.Visible         := (FStCycle <> '') and (VH^.OkModRIC) and JaiLeRoleCompta( RcReviseur );
  BCycle.Visible              := (FStCycle <> '') and JaiLeRoleCompta( RcReviseur );
  BSupervisTravaux.Visible    := (CtxPcl In V_Pgi.PgiContexte) and (FStCycle <> '') and (VH^.OkModRIC) and (JaileRoleCompta( RcSuperviseur ));
  BEdition.Visible            := (FStCycle <> '') and (VH^.OkModRIC) and JaiLeRoleCompta( RcReviseur );

  if (not VH^.OkModRIC) then
  begin
    SetControlVisible('TLETAT', False );
    SetControlVisible('CCY_ETATCYCLE', False );
    SetControlVisible('TLPAR', False );
    SetControlVisible('CCY_ETATCYCLEPAR', False );
    SetControlVisible('TLLE', False );
    SetControlVisible('CCY_ETATCYCLELE', False );
  end;

  if (FStCycle = '') or
     (not JaiLeRoleCompta( rcReviseur )) or
     (VH^.Revision.DossierSupervise) or
     // GCO - 03/05/2007 - FQ 20171
     // GCO - 02/07/2007 - FQ 20891
     (not JaileRoleCompta( RcSuperviseur ) and ((VH^.Revision.DossierPretSupervise) or (CCY_EtatCycle.Value = cSupervise)))
     then // Blocage totale des fonctions de révision
  begin
    CCY_EtatCycle.Enabled     := False;
    CCY_EtatCycleLe.Enabled   := False;
  end
  else
  begin
    CCY_EtatCycle.Enabled     := True;
    CCY_EtatCycle.Enabled     := True;

    // Suppresion systématique du CODE non Actif, car il n'est dispo que pour
    // la visualisation de l'historique
    CCY_EtatCycle.Plus := ' AND CO_CODE <> "' + cNonActif + '"';

    // Changement d'état d'un cycle enabled mais un réviseur ne doit pas pouvoir
    // supervisé le cycle
    if (not JaiLeRoleCompta(rcSuperviseur)) and (not VH^.Revision.DossierSupervise) then
      CCY_EtatCycle.Plus := CCY_EtatCycle.Plus + ' AND CO_CODE <> "' + CSupervise + '"';

    // GCO - 11/06/2007- FQ 20633
    if (not JaiLeRoleCompta(rcSuperviseur) and (CCY_EtatCycle.Value = cValide)) then
      CCY_EtatCycleLe.Enabled   := False;
  end;

  // GCO - 09/05/2007 - FQ 20277 + Blocage totale de la révison car le dossier est supervisé
  // GCO - 09/05/2007 - FQ 20274
  MemoCycle.ReadOnly     := (VH^.Revision.DossierSupervise) or
    ((FStCycle <> '') and ((CCY_EtatCycle.Value = cSupervise) or
    ((CCY_EtatCycle.Value = cValide) and (not JaiLeRoleCompta(rcSuperviseur)))));

  MemoSynthese.ReadOnly  := (VH^.Revision.DossierSupervise) or
    ((FStCycle <> '') and ((CCY_EtatCycle.Value = cSupervise) or
    ((CCY_EtatCycle.Value = cValide) and (not JaiLeRoleCompta(rcSuperviseur)))));

  MemoMillesime.ReadOnly := (VH^.Revision.DossierSupervise) or
    ((FStCycle <> '') and ((CCY_EtatCycle.Value = cSupervise) or
    ((CCY_EtatCycle.Value = cValide) and (not JaiLeRoleCompta(rcSuperviseur)))));

  MemoCompte.ReadOnly    := (VH^.Revision.DossierSupervise) or
    ((FStCycle <> '') and ((CCY_EtatCycle.Value = cSupervise) or
    ((CCY_EtatCycle.Value = cValide) and (not JaiLeRoleCompta(rcSuperviseur)))));

  if MemoCycle.ReadOnly     then
    MemoCycle.Color := ClBtnFace;

  if MemoSynthese.ReadOnly  then
    MemoSynthese.Color := ClBtnFace;

  if MemoMillesime.ReadOnly then
    MemoMillesime.Color := ClBtnFace;

  if MemoCompte.ReadOnly    then
    MemoCompte.Color := ClBtnFace;

  // Activation des boutons
  if VisaRevision.Enabled then
  begin
    if VisaRevision.Checked then
      VisaRevision.Hint := TraduireMemoire('Enlever le visa du compte')
    else
      VisaRevision.Hint := TraduireMemoire('Viser le compte');
  end;

  if FInActivePage >= 0 then
  begin
    Pages.ActivePageIndex := FInActivePage;
    case FInActivePage of
      0 : MemoCycle.SetFocus;
      1 : MemoObjectif.SetFocus;
      2 : MemoSynthese.SetFocus;
      3 : MemoMillesime.SetFocus;
      4 : MemoCompte.SetFocus;
    end;
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 19/12/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPREVDOCTRAVAUX.OnDisplay () ;
begin
  Inherited ;
end ;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 12/02/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPREVDOCTRAVAUX.OnKeyDownEcran(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
  case key of

    VK_ESCAPE : Ecran.Close;

    VK_F11 : PopF11.Popup(Mouse.CursorPos.x, Mouse.CursorPos.y);

    Ord('1') : if (Shift = [ssAlt]) and (TabMemoCycle.TabVisible) then
                 Pages.ActivePage := TabMemoCycle;

    Ord('2') : if (Shift = [ssAlt]) and (TabMemoObjectif.TabVisible) then
                 Pages.ActivePage := TabMemoObjectif;

    Ord('3') : if (Shift = [ssAlt]) and (TabMemoSynthese.TabVisible) then
                 Pages.ActivePage := TabMemoSynthese;

    Ord('4') : if (Shift = [ssAlt]) and (TabMemoMillesime.TabVisible) then
                 Pages.ActivePage := TabMemoMillesime;

    Ord('5') : if (Shift = [ssAlt]) and (TabMemoCompte.TabVisible) then
                 Pages.ActivePage := TabMemoCompte;

    86 : begin // ALT + V -> Viser le compte
           if (ssAlt in Shift) and (VisaRevision.Enabled) then
             VisaRevision.Checked := not VisaRevision.Checked;
         end;

  else
    fOnSaveKeyDownEcran(Sender, Key, Shift);
  end;
end;


////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 19/12/2006
Modifié le ... : 09/02/2007    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPREVDOCTRAVAUX.OnUpdate ;
begin
  Inherited ;

  if FStCycle <> '' then
  begin
    if CCY_EtatCycle.Value <> 'ATR' then
    begin
      if not IsValidDate( CCY_EtatCycleLe.Text ) then
        LastError := 1;
    end;
  end;

  SauveMemoMillesime;
  SauveMemoGeneral;
  SauveInfoCycle;
  SauveMemoCommentaireCycle;
  SauveMemoSyntheseCycle;

  // GCO - 11/10/2007 - FQ 21626
  MemoCycle.Modified     := False;
  MemoSynthese.Modified  := False;
  MemoMillesime.Modified := False;
  MemoCompte.Modified    := False;

end ;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 19/12/2006
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPREVDOCTRAVAUX.OnCancel () ;
begin
  Inherited ;
end ;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 08/01/2007
Modifié le ... : 15/01/2007
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPREVDOCTRAVAUX.ChargeInfoCycle;
var lQuery : TQuery;
    lNbCpt  : integer;
    lNbVise : integer;
begin
  if FStCycle = '' then
  begin
    Cycle.Caption  := TraduireMemoire('<<Aucun>>');
    FStEtatCycle   := '';
    FStEtatCycleLe := '';
    Exit;
  end;

  MemoCycle.Clear;
  lQuery := nil;
  try
    lQuery := OpenSql('SELECT CCY_LIBELLECYCLE, CCY_ETATCYCLE, CCY_ETATCYCLEPAR, ' +
                      'CCY_ETATCYCLELE FROM CREVCYCLE WHERE ' +
                      'CCY_CODECYCLE = "' + FStCycle + '" AND ' +
                      'CCY_EXERCICE = "' + FStExercice + '"', True);

    if not lQuery.Eof then
    begin
      Cycle.Caption            := FStCycle + ' - ' + lQuery.FindField('CCY_LIBELLECYCLE').AsString;
      CCY_EtatCycle.Value      := lQuery.FindField('CCY_ETATCYCLE').AsString;
      CCY_EtatCyclePar.Caption := '';
      if lQuery.FindField('CCY_ETATCYCLEPAR').AsString <> '' then
        CCY_EtatCyclePar.Caption := RechDom('TTUTILISATEUR', lQuery.FindField('CCY_ETATCYCLEPAR').AsString, False);
       CCY_EtatCycleLe.Text    := DateTosTr(lQuery.FindField('CCY_ETATCYCLELE').AsDateTime);
      FStEtatCycle             := CCY_EtatCycle.Value;
      FStEtatCycleLe           := CCY_EtatCycleLe.Text;
    end;
  finally
    Ferme(lQuery);
  end;

  // GCO - 10/05/2007 - FQ 20270
  lQuery := OpenSql('SELECT COUNT(*) TOTAL FROM GENERAUX WHERE ' +
                    'G_CYCLEREVISION = "' + FStCycle +  '" AND ' +
                    '(G_TOTDEBE <> 0 OR G_TOTCREE <> 0)', True);
  lNbCpt := lQuery.FindField('TOTAL').AsInteger;
  Ferme(lQuery);

  if lNbCpt <> 0 then
  begin
    lQuery := OpenSql('SELECT COUNT(*) TOTAL FROM GENERAUX WHERE ' +
                      'G_CYCLEREVISION = "' + FStCycle +  '" AND ' +
                      'G_VISAREVISION = "X" AND (G_TOTDEBE <> 0 OR G_TOTCREE <> 0)', True);
    lNbVise := lQuery.FindField('TOTAL').AsInteger;
    TLPourcentVise.Caption := StrFMontant((lNbVise * 100) / lNbCpt , 15, V_PGI.OkdecV, '', True ) + ' %';
    Ferme(lQuery);
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 19/12/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPREVDOCTRAVAUX.ChargeInfoGeneral;
var lQuery : TQuery;
begin
  MemoCompte.Clear;
  try
    lQuery := OpenSQL('SELECT G_GENERAL, G_LIBELLE, G_CYCLEREVISION, G_BLOCNOTE ' +
                      'FROM GENERAUX ' +
                      'WHERE G_GENERAL = "' + FStGeneral + '"', True);

    General.Caption := lQuery.FindField('G_GENERAL').AsString + ' - ' +
                       lQuery.FindField('G_LIBELLE').AsString;

    // GCO - 06/02/2007 - Chargement du cycle du compte général si aucun passé en paramètre.
    if FStCycle = '' then
      FStCycle := lQuery.FindField('G_CYCLEREVISION').AsString;

    StringToRich(MemoCompte, lQuery.FindField('G_BLOCNOTE').AsString);

    TabMemoCompte.ImageIndex := -1;
    if MemoCompte.Text <> '' then
    begin
      if V_Pgi.ModeTSE then
        TabMemoCompte.Caption := TabMemoCompte.Caption + ' (***)'
      else
        TabMemoCompte.ImageIndex := cImgBlocNote;
    end;

    // Obligation de remettre à FALSE après chargement
    MemoCompte.Modified := False;

  finally
    Ferme(lQuery);
  end;

  try
    lQuery := OpenSQL('SELECT CRG_VISAREVISION, CRG_VISAPAR, CRG_VISALE, CRG_VISAAUTO ' +
                      'FROM CREVGENERAUX WHERE ' +
                      'CRG_GENERAL = "' + FStGeneral + '" AND ' +
                      'CRG_EXERCICE = "' + FStExercice + '"', True);

    if not lQuery.Eof then
    begin
      VISAREVISION.Checked := (lQuery.FindField('CRG_VISAREVISION').AsString = 'X');

      // GCO - 09/05/2007 - FQ 20249
      if lQuery.FindField('CRG_VISAAUTO').AsString = 'X' then
        TLPARLE.Caption := TraduireMemoire('Visa automatique')
      else
      begin
        if lQuery.FindField('CRG_VISAPAR').AsString <> '' then
          TLPARLE.Caption := TraduireMemoire('Par') + ' ' +
                             RechDom('TTUTILISATEUR', lQuery.FindField('CRG_VISAPAR').AsString, False) +
                             ' ' + TraduireMemoire('le') + ' ' +
                             lQuery.FindField('CRG_VISALE').AsString;
      end;                       
    end;
  finally
    Ferme(lQuery);
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 17/01/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPREVDOCTRAVAUX.ChargeMemoCommentaireCycle;
begin
  MemoCycle.Clear;
  if FStCycle = '' then Exit;
  StringToRich(MemoCycle, ChargeMemoCREVBLOCNOTE('CCY', FStCycle, FStExercice));
  // Obligation de remettre à FALSE après chargement
  MemoCycle.Modified:= False;

  TabMemoCycle.ImageIndex := -1;
  if MemoCycle.Text <> '' then
  begin
    if V_Pgi.ModeTse then
      TabMemoCycle.Caption := TabMemoCycle.Caption +  ' (***)'
    else
      TabMemoCycle.ImageIndex := cImgBlocNote;
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 19/12/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPREVDOCTRAVAUX.ChargeMemoObjectif;
var lQuery : TQuery;
begin
  MemoObjectif.ReadOnly := True;
  MemoObjectif.Color    := ClBtnFace;
  MemoObjectif.Clear;
  try
    lQuery := OpenSQl('SELECT CPC_BLOCNOTE FROM CREVPARAMCYCLE WHERE ' +
                      'CPC_CODECYCLE = "' + FStCycle + '" AND ' +
                      // GCO - 18/09/2007 - FQ 21433
                      'CPC_PREDEFINI = "' + FStPredefini + '"', True);

    if not lQuery.Eof then
      StringToRich(MemoObjectif, lQuery.FindField('CPC_BLOCNOTE').AsString);

    TabMemoObjectif.ImageIndex := -1;
    if MemoObjectif.Text <> '' then
    begin
      if V_Pgi.ModeTSE then
        TabMemoObjectif.Caption := TabMemoObjectif.Caption + ' (***)'
      else
        TabMemoObjectif.ImageIndex := cImgBlocNote;
    end;

  finally
    Ferme(lQuery);
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 19/12/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPREVDOCTRAVAUX.ChargeMemoSyntheseCycle;
begin
  MemoSynthese.Clear;
  StringToRich(MemoSynthese, ChargeMemoCREVBLOCNOTE('SCY', FStCycle, FStExercice));
  // Obligation de remettre à FALSE après chargement
  MemoSynthese.Modified := False;

  TabMemoSynthese.ImageIndex := -1;
  if MemoSynthese.Text <> '' then
  begin
    if V_Pgi.ModeTSE then
      TabMemoSynthese.Caption := TabMemoSynthese.Caption + ' (***)'
    else
      TabMemoSynthese.ImageIndex := cImgBlocNote;
  end;

  if not JaiLeRoleCompta( rcReviseur ) then
  begin
    MemoSynthese.ReadOnly := True;
    MemoSynthese.Color    := ClBtnFace;
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 17/01/2007
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPREVDOCTRAVAUX.ChargeMemoMillesime;
begin
  MemoMillesime.Clear;
  StringToRich(MemoMillesime, ChargeMemoCREVBLOCNOTE('GEN', FStGeneral, FStExercice));
  // Obligation de remettre à FALSE après chargement
  MemoMillesime.Modified := False;

  TabMemoMillesime.ImageIndex := -1;
  if MemoMillesime.Text <> '' then
  begin
    if V_Pgi.ModeTSE then
      TabMemoMillesime.Caption := TabMemoMillesime.Caption + ' (***)'
    else
      TabMemoMillesime.ImageIndex := cImgBlocNote;
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 19/12/2006
Modifié le ... : 29/12/2006
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPREVDOCTRAVAUX.SauveMemoGeneral;
var lTob     : Tob;
    lTobTemp : Tob;
begin
  if not MemoCompte.Modified then Exit;
  try
    try
      lTob := Tob.Create('', nil, -1);
      lTob.LoadDetailDBFromSQL('GENERAUX',
                               'SELECT G_GENERAL, G_BLOCNOTE FROM GENERAUX WHERE ' +
                               'G_GENERAL = "' + FStGeneral + '"');

      if lTob.Detail.Count > 0 then
      begin
        lTobTemp := lTob.Detail[0];
        lTobTemp.PutValue('G_BLOCNOTE', RichToString(MemoCompte));
        lTob.UpdateDB;
      end;

    except
      on E: Exception do
        PgiError('Erreur de requête SQL : ' + E.Message, 'SauveMemoGeneral');
    end;

  finally
    FreeAndNil(lTob);
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 19/12/2006
Modifié le ... : 17/01/2007
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPREVDOCTRAVAUX.SauveInfoCycle;
var lRic : TRic;
begin
  if FStCycle = '' then Exit;

  if (CCY_EtatCycle.Value <> FStEtatCycle) or
     (CCY_EtatCycleLe.Text <> FStEtatCycleLe) then
  begin
    lRic := TRic.Create;
    try
      lRic.ChangeEtatCycle( FStCycle, FStExercice, CCY_EtatCycle.Value, V_Pgi.User, StrToDate(CCY_EtatCycleLe.Text));
    finally
      lRic.Free;
      // Mise à jour des variables initialisés à l'ouverture de la fiche
      FStEtatCycle   := CCY_EtatCycle.Value;
      FStEtatCycleLe := CCY_EtatCycleLe.Text;
    end;  
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 17/01/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPREVDOCTRAVAUX.SauveMemoCommentaireCycle;
begin
  if not MemoCycle.Modified then Exit;
  if Trim(MemoCycle.Text) = '' then
    ExecuteSQL('DELETE FROM CREVBLOCNOTE WHERE ' +
               'CBN_NATURE = "CCY" AND ' +
               'CBN_CODE = "' + FStCycle + '" AND ' +
               'CBN_EXERCICE = "' + VH^.EnCours.Code + '"')
  else
    SauveMemoCREVBLOCNOTE('CCY', FStCycle, FStExercice, RichToString(MemoCycle));
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 19/12/2006
Modifié le ... : 17/01/2007
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPREVDOCTRAVAUX.SauveMemoSyntheseCycle;
begin
  if not MemoSynthese.Modified then Exit;
  if Trim(MemoSynthese.Text) = '' then
    ExecuteSQL('DELETE FROM CREVBLOCNOTE WHERE ' +
               'CBN_NATURE = "SCY" AND ' +
               'CBN_CODE = "' + FStCycle + '" AND ' +
               'CBN_EXERCICE = "' + VH^.EnCours.Code + '"')
  else
    SauveMemoCREVBLOCNOTE('SCY', FStCycle, FStExercice, RichToString(MemoSynthese));
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 19/12/2006
Modifié le ... : 17/01/2007
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPREVDOCTRAVAUX.SauveMemoMillesime;
begin
  if not MemoMillesime.Modified then Exit;
  if Trim(MemoMillesime.Text) = '' then
    ExecuteSQL('DELETE FROM CREVBLOCNOTE WHERE ' +
               'CBN_NATURE = "GEN" AND ' +
               'CBN_CODE = "' + FStGeneral + '" AND ' +
               'CBN_EXERCICE = "' + VH^.EnCours.Code + '"')
  else
    SauveMemoCREVBLOCNOTE('GEN', FStGeneral, FStExercice, RichToString(MemoMillesime) );
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 17/01/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPREVDOCTRAVAUX.OnChangeCCY_EtatCycle(Sender: TObject);
begin
  CCY_EtatCyclePar.Enabled := (CCY_EtatCycle.Value <> cATraiter) and (JaiLeRoleCompta( rcReviseur ));
  CCY_EtatCycleLe.Enabled  := (CCY_EtatCycle.Value <> cATraiter) and (JaiLeRoleCompta( rcReviseur ));

  if CCY_EtatCycle.Value = cATraiter then
  begin
    CCY_EtatCyclePar.Caption := '';
    CCY_EtatCycleLe.Text     := DateToStr(iDate1900);
  end
  else
  begin
    CCY_EtatCyclePar.Caption := RechDom('TTUTILISATEUR', V_PGI.User, False);
    if CCY_EtatCycleLe.Text = DateToStr(iDate1900) then
      CCY_EtatCycleLe.Text := DateToStr(Date);
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 12/04/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPREVDOCTRAVAUX.OnClickBHistoMemo(Sender: TObject);
begin
  // GCO - 11/06/2007 - FQ 20631
  if FStGeneral <> '' then
    CC_LanceFicheHistoCpte( FStGeneral )
  else
    CC_LanceFicheHistoCpte( FStCycle, 'CCY');
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 12/04/2007
Modifié le ... :   /  /
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPREVDOCTRAVAUX.OnClickBHistoCycle(Sender: TObject);
begin
  CPLanceFiche_CPREVHISTOCYCLE( FStCycle, VH^.EnCours.Code );
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 22/01/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPREVDOCTRAVAUX.OnClickVisaRevision(Sender: TObject);
begin
  if FBoLoading then Exit;

  if MiseAJourG_VisaRevision( FStGeneral, VisaRevision.Checked ) then
  begin
    if VisaRevision.Checked then
    begin
      TLPARLE.Caption := TraduireMemoire('Par') + ' ' +
                         RechDom('TTUTILISATEUR', V_Pgi.User, False) + ' ' +
                         TraduireMemoire('le') + ' ' + DateTimeToStr(Now);
      VisaRevision.Hint := TraduireMemoire('Enlever le visa du compte');
    end
    else
    begin
      TLPARLE.Caption := TraduireMemoire('par') +
                         ' .............................................. ' +
                         TraduireMemoire('le') + ' __/__/____';
      VisaRevision.Hint := TraduireMemoire('Viser le compte');
    end;
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 05/03/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPREVDOCTRAVAUX.OnClickMenuVisaRevision(Sender: TObject);
begin
  VisaRevision.Checked := not VisaRevision.Checked;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 22/01/2007
Modifié le ... :   /  /    
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPREVDOCTRAVAUX.OnClickBCycle(Sender: TObject);
begin
  // GCO - 18/09/2007 - FQ 21433
  CPLanceFiche_CREVPARAMCYCLE( FStCycle + ';' + FStPredefini, TaConsult );
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 26/02/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPREVDOCTRAVAUX.OnClickNoteCtrlCompte(Sender: TObject);
begin
  if SauveAvantEdition <> MrCancel then
    TRic.CPLanceEtat_NoteCtrlCompte( FStGeneral );
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 26/02/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPREVDOCTRAVAUX.OnClickNoteCtrlCycle(Sender: TObject);
begin
  // GCO - 01/08/2007 - FQ 21207
  if SauveAvantEdition <> MrCancel then
    TRic.CPLanceEtat_NoteCtrlCycle( FStCycle, 0 );
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 10/05/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPREVDOCTRAVAUX.OnClickNoteSupervisionDossier(Sender: TObject);
begin
  if SauveAvantEdition <> MrCancel then
    TRic.CPLanceEtat_NoteSupervision( 0 );
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 23/02/2007
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPREVDOCTRAVAUX.OnPopUpPopF11(Sender: TObject);
begin
  PopF11.Items[0].Visible := VisaRevision.Enabled;     // Viser le compte
  PopF11.Items[0].Caption := VisaRevision.Hint;
  PopF11.Items[1].Visible := BHistoMemo.Visible;       // Historique des commentaires
  PopF11.Items[2].Visible := BHistoCycle.Visible;      // Historique du cyle
  PopF11.Items[3].Visible := BCycle.Visible;           // Paramétrage du cycle
  PopF11.Items[4].Visible := BSupervisTravaux.Visible; // Supervision des travaux
  // Séparateur
  PopF11.Items[5].Visible := BEdition.Visible;
  PopF11.Items[6].Visible := BEdition.Visible;
  PopF11.Items[6].Items[0].Visible := (FStGeneral <> ''); // Note de contrôle Compte
  PopF11.Items[6].Items[1].Visible := (FStCycle <> '');   // Note de contrôle Cycle
  PopF11.Items[6].Items[2].Visible := False;              // Note de Supervision du dossier
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 12/04/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPREVDOCTRAVAUX.OnPopUpPopEdition(Sender: TObject);
begin
  PopEdition.Items[0].Visible := (FStGeneral <> '');
  PopEdition.Items[1].Visible := (FStCycle <> '');
  PopEdition.Items[2].Visible := False;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 12/04/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPREVDOCTRAVAUX.OnClickAPG(Sender: TObject);
begin
  CPLanceFiche_CPREVINFODOSSIER( V_Pgi.NoDossier, VH^.EnCours.Code, 1, TaModif, 1);
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 12/04/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPREVDOCTRAVAUX.OnClickSCY(Sender: TObject);
begin
  CPLanceFiche_CPREVINFODOSSIER( V_Pgi.NoDossier, VH^.EnCours.Code, 1, TaModif, 2);
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 12/04/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPREVDOCTRAVAUX.OnClickEXP(Sender: TObject);
begin
  CPLanceFiche_CPREVINFODOSSIER( V_Pgi.NoDossier, VH^.EnCours.Code, 1, TaModif, 3);
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 03/08/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
function TOF_CPREVDOCTRAVAUX.SauveAvantEdition : integer;
begin
  Result := MrNo;

  if (TabMemoCycle.TabVisible and MemoCycle.Modified) or
     (TabMemoSynthese.TabVisible and MemoSynthese.Modified) or
     (TabMemoMillesime.TabVisible and MemoMillesime.Modified) or
     (TabMemoCompte.TabVisible and MemoCompte.Modified) then
  begin
    Result := PgiAskCancel('Voulez-vous enregistrer les modifications ?', Ecran.Caption);
    case Result of
       mrNo                : ;
       mrCancel            : Exit;
       mrYes               : BValider.Click;
    end;
  end;
end;

////////////////////////////////////////////////////////////////////////////////

Initialization
  registerclasses ( [ TOF_CPREVDOCTRAVAUX ] ) ;
end.

