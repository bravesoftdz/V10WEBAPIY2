{***********UNITE*************************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 23/02/2007
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : CPREVPROGTRAVAIL ()
Mots clefs ... : TOF;CPREVPROGTRAVAIL
*****************************************************************}
Unit CPREVPROGTRAVAIL_TOF ;

Interface

Uses StdCtrls, 
     Controls, 
     Classes, 
{$IFDEF EAGLCLIENT}
     eMul,
     MainEagl,      // AGLLanceFiche
     UtileAGL,      // LanceEtat
{$ELSE}
     db,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul,
     FE_main,       // AGLLanceFiche
     Hdb,
     EdtREtat,      // LanceEtat
{$ENDIF}
     Forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,

     Windows,       // VK_F11
     Menus,         // TPopUpMenu
     Htb97,         // TToolBarButton97
     uTob,          // TOB
     {$IFDEF MODENT1}
     CPTypeCons,
     {$ENDIF MODENT1}
     Ent1,
     UTOF ;

Type
  TOF_CPREVPROGTRAVAIL = Class (TOF)
    procedure OnArgument (S : String ) ; override ;
    procedure OnClose                  ; override ;
    procedure OnLoad                   ; override ;
    procedure OnDisplay                ; override ;

    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnCancel                 ; override ;

    procedure OnKeyDownEcran         ( Sender : TObject; var Key: Word; Shift: TShiftState);
    procedure OnClickBSelectAll      ( Sender : TObject );      

  private
    FRoleCompta : TCPRoleCompta;

    fOnSaveKeyDownEcran : procedure(Sender: TObject; var Key: Word; Shift: TShiftState) of object;

    procedure _ModifEtatCycleEnSerie( const vStEtatCycle : string );
    procedure _ModifActivationCycle ( const vStActivation : string );
    procedure _ActivationBouton;

  public

{$IFDEF EAGLCLIENT}
    FListe : THGrid;
{$ELSE}
    FListe : THDBGrid;
{$ENDIF}

    CCY_ETATCYCLE        : THMultiValComboBox;
    CCY_ACTIVECYCLE      : THMultiValComboBox;
    CCY_CODECYCLE        : THEdit;

    BDocTravaux          : TToolBarButton97;
    BEtatCycle           : TToolBarButton97;
    BParam               : TToolBarButton97;
    BActivation          : TToolBarButton97;
    BHistoMemo           : TToolBarButton97;
    BHistoCycle          : TToolBarButton97;
    BSupervisTravaux     : TToolBarButton97;
    BOuvrir              : TToolBarButton97;
    BSelectAll           : TToolBarButton97;

    PopUpEtatCycle       : TPopUpMenu;
    PopUpActivation      : TPopUpMenu;
    PopUpDocTravaux      : TPopUpMenu;
    PopUpSupervisTravaux : TPopUpMenu;
    PopUpEdition         : TPopUpMenu;
    PopUpF11             : TPopUpMenu;

    procedure OnFlipSelectionFListe  ( Sender : TObject );

    procedure OnClickMemoCycle       ( Sender : TObject ); // Mémo cycle de révision de la documentation des travaux
    procedure OnClickMemoObjectif    ( Sender : TObject ); // Mémo Objectif de révision de la documentation des travaux
    procedure OnClickMemoSynthese    ( Sender : TObject ); // Mémo synthèse

    procedure OnClickBParam          ( Sender : TObject );
    procedure OnClickBHistoMemo      ( Sender : TObject );
    procedure OnClickBHistoCycle     ( Sender : TObject );
    procedure OnClickAPG             ( Sender : TObject ); // Appréciation générale
    procedure OnClickSCY             ( Sender : TObject ); // Synthèse des cycles
    procedure OnClickEXP             ( Sender : TObject ); // Attestation Expert

    procedure OnPopUpEtatCycle       ( Sender : TObject );
    procedure OnPopUpF11             ( Sender : TObject );

    procedure OnClickEtatATraiter    ( Sender : TObject );
    procedure OnClickEtatEnCours     ( Sender : TObject );
    procedure OnClickEtatEnSuspens   ( Sender : TObject );
    procedure OnClickEtatValide      ( Sender : TObject );
    procedure OnClickEtatSupervise   ( Sender : TObject );
    procedure OnClickRendreActif     ( Sender : TObject );
    procedure OnClickRendreNonActif  ( Sender : TObject );
    procedure OnClickEtatInitial     ( Sender : TObject );
    procedure OnClickDetailCycle     ( Sender : TObject );

    procedure OnClickProgrammeTravail       ( Sender : TObject );
    procedure OnClickNoteCtrlCycle          ( Sender : TObject );
    procedure OnClickNoteSupervisionDossier ( Sender : TObject );

  end ;

////////////////////////////////////////////////////////////////////////////////
procedure CPLanceFiche_CPREVProgTravail( vStCodeExo : HString; vStRacine : HString = '');
////////////////////////////////////////////////////////////////////////////////

Implementation

uses AGLInit,             // TheMulQ
     ParamSoc,            // GetParamSocSecur
     uLibRevision,        // ChangeEtatCycle
     uLibWindows,         // HGBeginUpdate
     uTOFConsGene,        // CPLanceFiche_CPCONSGENE
     uTofHistoCpte,       // //CC_LanceFicheHistoCpte
     CRevInfoDossier_Tom, // CPLanceFiche_CREVNOTESUPERVISION
     CRevParamCycle_Tom,  // CPLanceFiche_CRevParamCycle
     CPRevHistoCycle_TOF, // CPLanceFiche_CPREVHISTOCYCLE
     CPRevDocTravaux_TOF; // CPLanceFiche_CPREVDOCTRAVAUX

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 23/02/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure CPLanceFiche_CPREVProgTravail( vStCodeExo : HString; vStRacine : HString = '');
var lRic : TRic;
begin
  lRic := nil;

  // GCO - 02/08/2007 - FQ 20877
  if Pos('dy cycle', GetColonneSQL('DECHAMPS', 'DH_LIBELLE', 'DH_NOMCHAMP = "CCY_LIBELLECYCLE"')) > 0 then
  begin
    ExecuteSQL('UPDATE DECHAMPS SET ' +
               'DH_LIBELLE = "' + TraduireMemoire('Libellé du cycle') + '" WHERE ' +
               'DH_NOMCHAMP = "CCY_LIBELLECYCLE"');
  end;             

  // GCO - 12/10/2007 - FQ 21633 - Par Sécurité, on conserve ce test de présénce
  // des cycles pour le millésimé de l'exerice.
  if not ExisteSQL('SELECT CCY_CODECYCLE FROM CREVCYCLE WHERE ' +
                   'CCY_EXERCICE = "' + VH^.EnCours.Code + '"') then
  begin
    try
      lRic := TRic.Create;
      lRic.ProgressForm := True;
      if not lRic.Affecte(VH^.Revision.Plan) then
        PgiInfo( 'Traitement annulé. ' + lRic.LastErrorMsg, 'Mise à jour de la révision');
    finally
      lRic.Free;
    end;
  end;

  AGLLanceFiche('CP', 'CPREVPROGTRAVAIL', '', '', vStCodeExo + ';' + vStRacine);
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 23/02/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPREVPROGTRAVAIL.OnArgument(S : String ) ;
begin
  Inherited ;

{$IFDEF EAGLCLIENT}
  FListe := THGrid(GetControl('FLISTE', True));
{$ELSE}
  FListe := THDBGrid(GetControl('FLISTE', True));
{$ENDIF}

  if S <> '' then
  begin
    THEdit(GetControl('XX_WHERE', True)).Text := 'CCY_EXERCICE = "' + ReadTokenSt(S) + '"';

    THEdit(GetControl('CCY_CODECYCLE', True)).Text := ReadTokenSt(S);
  end;

  CCY_ETATCYCLE    := THMultiValComboBox(GetControl('CCY_ETATCYCLE', True));
  CCY_ACTIVECYCLE  := THMultiValComboBox(GetControl('CCY_ACTIVECYCLE', True));
  CCY_CODECYCLE    := THEdit(GetControl('CCY_CODECYCLE', True));

  BDocTravaux      := TToolBarButton97(GetControl('BDOCTRAVAUX', True));
  BEtatCycle       := TToolBarButton97(GetControl('BETATCYCLE', True));
  BParam           := TToolBarButton97(GetControl('BPARAM', True));
  BActivation      := TToolBarButton97(GetControl('BACTIVATION', True));
  BHistoMemo       := TToolBarButton97(GetControl('BHISTOMEMO', True));
  BHistoCycle      := TToolBarButton97(GetControl('BHISTOCYCLE', True));
  BSupervisTravaux := TToolBarButton97(GetControl('BSUPERVISTRAVAUX', True));
  BOuvrir          := TToolBarButton97(GetControl('BOUVRIR', True));
  BSelectAll       := TToolBarButton97(GetControl('BSELECTALL', True));

  PopUpEtatCycle       := TPopUpMenu(GetControl('POPUPETATCYCLE', True));
  PopUpActivation      := TPopUpMenu(GetControl('POPUPACTIVATION', True));
  PopUpDocTravaux      := TPopUpMenu(GetControl('POPUPDOCTRAVAUX', True));
  PopUpSupervisTravaux := TPopUpMenu(GetControl('POPUPSUPERVISTRAVAUX', True));
  PopUpEdition         := TPopUpMenu(GetControl('POPUPEDITION', True));
  PopUpF11             := TPopUpMenu(GetControl('POPUPF11', True));

  // Branchement des événements
  BHistoMemo.OnClick       := OnClickBHistoMemo;
  BHistoCycle.OnClick      := OnClickBHistoCycle;
  BParam.OnClick           := OnClickBParam;
  BSelectAll.OnClick       := OnClickBSelectAll;

  PopUpEtatCycle.OnPopUp := OnPopUpEtatCycle;
  PopUpEtatCycle.Items[0].OnClick := OnClickEtatATraiter;
  PopUpEtatCycle.Items[1].OnClick := OnClickEtatEnCours;
  PopUpEtatCycle.Items[2].OnClick := OnClickEtatEnSuspens;
  PopUpEtatCycle.Items[3].OnClick := OnClickEtatValide;
  PopUpEtatCycle.Items[4].OnClick := OnClickEtatSupervise;

  // GCO - 13/09/2007 - FQ 21368 - Pas d'objectif de révision en PGE
  PopUpDocTravaux.Items[1].Visible := CtxPcl in V_Pgi.PgiContexte;

  PopUpDocTravaux.Items[0].OnClick := OnClickMemoCycle;
  PopUpDocTravaux.Items[1].OnClick := OnClickMemoObjectif;
  PopUpDocTravaux.Items[2].OnClick := OnClickMemoSynthese;

  PopUpSupervisTravaux.Items[0].OnClick := OnClickAPG;
  PopUpSupervisTravaux.Items[1].OnClick := OnClickSCY;
  PopUpSupervisTravaux.Items[2].OnClick := OnClickEXP;

  PopUpActivation.Items[0].OnClick := OnClickRendreActif;
  PopUpActivation.Items[1].OnClick := OnClickRendreNonActif;
  PopUpActivation.Items[2].OnClick := OnClickEtatInitial;

  PopupEdition.Items[0].OnClick    := OnClickProgrammeTravail;
  PopupEdition.Items[1].OnClick    := OnClickNoteCtrlCycle;
  PopupEdition.Items[3].OnClick    := OnClickNoteSupervisionDossier;

  PopUpF11.OnPopUp := OnPopUpF11;
  PopUPF11.Items[0].OnClick           := OnClickBHistoMemo;
  PopUPF11.Items[1].OnClick           := OnClickBHistoCycle;
  PopUPF11.Items[2].OnClick           := OnClickBParam;
  PopUPF11.Items[3].Items[0].OnClick  := OnClickMemoCycle;
  PopUPF11.Items[3].Items[1].OnClick  := OnClickMemoObjectif;
  PopUPF11.Items[3].Items[2].OnClick  := OnClickMemoSynthese;
  PopUPF11.Items[4].Items[0].OnClick  := OnClickAPG;
  PopUPF11.Items[4].Items[1].OnClick  := OnClickSCY;
  PopUPF11.Items[4].Items[2].OnClick  := OnClickEXP;
  PopUPF11.Items[5].Items[0].OnClick  := OnClickEtatATraiter;
  PopUPF11.Items[5].Items[1].OnClick  := OnClickEtatEnCours;
  PopUPF11.Items[5].Items[2].OnClick  := OnClickEtatEnSuspens;
  PopUPF11.Items[5].Items[3].OnClick  := OnClickEtatValide;
  PopUPF11.Items[5].Items[4].OnClick  := OnClickEtatSupervise;
  PopUPF11.Items[6].Items[0].OnClick  := OnClickRendreActif;
  PopUPF11.Items[6].Items[1].OnClick  := OnClickRendreNonActif;
  PopUPF11.Items[6].Items[2].OnClick  := OnClickEtatInitial;
  PopUpF11.Items[8].Items[0].OnClick  := OnClickProgrammeTravail;
  PopUpF11.Items[8].Items[1].OnClick  := OnClickNoteCtrlCycle;
  PopUpF11.Items[8].Items[3].OnClick  := OnClickNoteSupervisionDossier;

  fOnSaveKeyDownEcran     := Ecran.OnKeyDown;
  Ecran.OnKeyDown         := OnKeyDownEcran;

  FListe.OnDblClick       := OnClickDetailCycle;
  FListe.OnFlipSelection  := OnFlipSelectionFListe;


  // Initialisation
  if JaiLeRoleCompta( rcSuperviseur ) then
  begin
    FRoleCompta := RcSuperviseur;
    Ecran.Caption := Ecran.Caption + ' (' + TraduireMemoire('Superviseur') + ')';
  end
  else
  begin
    FRoleCompta := RcReviseur;
    Ecran.Caption := Ecran.Caption + ' (' + TraduireMemoire('Réviseur') + ')';
  end;

  CCY_ETATCYCLE.Text    := '<<' + TraduireMemoire('Tous') + '>>';
  CCY_ETATCYCLE.Plus    := ' AND CO_CODE <> "' + cNonActif + '"'; // GCO - 11/06/2007 - FQ 20661

  CCY_CODECYCLE.Plus    := ' AND CCY_EXERCICE = "' + VH^.EnCours.Code + '"';
  CCY_ACTIVECYCLE.Value := 'X;';
  
  Ecran.Caption := Ecran.Caption + ' : ' + DateToStr(VH^.EnCours.Fin);
  UpdateCaption(Ecran);

  // GCO - 04/10/2007
  ActivateXpPopUp( PopUPF11 );

end ;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 23/02/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPREVPROGTRAVAIL.OnClose ;
begin
  Inherited ;
end ;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 23/02/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPREVPROGTRAVAIL.OnLoad ;
begin
  Inherited ;
  if FListe.CanFocus then FListe.SetFocus;
end ;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 23/02/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPREVPROGTRAVAIL.OnDisplay () ;
begin
  Inherited ;
  _ActivationBouton;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 23/03/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPREVPROGTRAVAIL.OnFlipSelectionFListe(Sender: TObject);
begin
  inherited ;
  _ActivationBouton;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 23/03/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPREVPROGTRAVAIL._ActivationBouton;
var lBocycleActif : Boolean;
begin
  lBoCycleActif := (GetField('CCY_ACTIVECYCLE') = 'X');

  BDocTravaux.Enabled := (not FListe.AllSelected) and (FListe.NbSelected = 0) and lBoCycleActif;
  BParam.Enabled      := (not FListe.AllSelected) and (FListe.NbSelected = 0);
  BHistoMemo.Enabled  := (not FListe.AllSelected) and (FListe.NbSelected = 0) and lBoCycleActif;
  BHistoCycle.Enabled := (not FListe.AllSelected) and (FListe.NbSelected = 0) and lBoCycleActif;

  // GCO - 13/09/2007 - FQ 21368
  BSupervisTravaux.Visible := (CtxPCL in V_PGI.PgiContexte) and (FRoleCompta = rCSuperviseur);
  // GCO - 09/05/2007 - FQ 20276
  BSupervisTravaux.Enabled := (not FListe.AllSelected) and (FListe.NbSelected = 0) and lBoCycleActif;

  // GCO - 15/06/2007 - FQ 20598
  if ((not FListe.AllSelected) and (FListe.NbSelected = 0)) or
     (VH^.Revision.DossierSupervise) or
     ((VH^.Revision.DossierPretSupervise) and (not JaileRoleCompta( RcSuperviseur))) then
    BEtatCycle.Enabled := False
  else
    BEtatCycle.Enabled := True;

  BActivation.Enabled  := BEtatCycle.Enabled;

  BOuvrir.Enabled      := ((not FListe.AllSelected) and (FListe.NbSelected = 0) and lBoCycleActif) or
                          ((FListe.AllSelected) or (FListe.NbSelected > 0));
end;

////////////////////////////////////////////////////////////////////////////////

procedure TOF_CPREVPROGTRAVAIL.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_CPREVPROGTRAVAIL.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_CPREVPROGTRAVAIL.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_CPREVPROGTRAVAIL.OnCancel () ;
begin
  Inherited ;
end ;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 30/03/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPREVPROGTRAVAIL.OnClickMemoCycle(Sender: TObject);
begin
  if GetDataSet.Bof and GetDataSet.Eof then Exit;
  CPLanceFiche_CPRevDocTravaux( '', GetField('CCY_CODECYCLE'), VH^.EnCours.Code, 0 );
  TToolBarButton97(GetControl('BCHERCHE', True)).Click;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 30/03/2007
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPREVPROGTRAVAIL.OnClickMemoObjectif(Sender: TObject);
begin
  if GetDataSet.Bof and GetDataSet.Eof then Exit;
  CPLanceFiche_CPRevDocTravaux( '', GetField('CCY_CODECYCLE'), VH^.EnCours.Code, 1 );
  TToolBarButton97(GetControl('BCHERCHE', True)).Click;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 30/03/2007
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPREVPROGTRAVAIL.OnClickMemoSynthese(Sender: TObject);
begin
  if GetDataSet.Bof and GetDataSet.Eof then Exit;
  CPLanceFiche_CPRevDocTravaux( '', GetField('CCY_CODECYCLE'), VH^.EnCours.Code, 2 );
  TToolBarButton97(GetControl('BCHERCHE', True)).Click;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 07/03/2007
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPREVPROGTRAVAIL.OnClickBParam( Sender : TObject );
begin
  if GetDataSet.Bof and GetDataSet.Eof then Exit;
  CPLanceFiche_CRevParamCycle( GetField('CCY_CODECYCLE') + ';' + GetField('CCY_PREDEFINI'), TaConsult );
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 07/03/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPREVPROGTRAVAIL.OnPopUpEtatCycle( Sender : TObject );
begin
  PopUpEtatCycle.Items[3].Visible := (FRoleCompta >= rcReviseur);
  PopUpEtatCycle.Items[4].Visible := (FRoleCompta >= rcSuperviseur);
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 07/03/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPREVPROGTRAVAIL.OnPopUpF11( Sender : TObject );
begin
  PopUPF11.Items[0].Visible          := BHistoMemo.Enabled;
  PopUPF11.Items[1].Visible          := BHistoCycle.Enabled;
  PopUPF11.Items[2].Visible          := BParam.Enabled;

  // Docuùmentation des travaux
  PopUPF11.Items[3].Visible          := BDocTravaux.Enabled;

  // GCO - 13/09/2007 - FQ 21368
  PopUPF11.Items[3].Items[1].Visible      := CtxPcl in V_Pgi.PgiContexte;

  // Supervision des travaux pour (N3) ou (N4)
  PopUPF11.Items[4].Visible          := (FRoleCompta >= rcSuperviseur) and BSupervisTravaux.Enabled;
  // Changement d'état du cycle
  PopUPF11.Items[5].Visible          := BEtatCycle.Enabled;
  PopUPF11.Items[5].Items[0].Visible := BEtatCycle.Enabled; // A traiter
  PopUPF11.Items[5].Items[1].Visible := BEtatCycle.Enabled; // En cours
  PopUPF11.Items[5].Items[2].Visible := BEtatCycle.Enabled; // En suspens
  PopUPF11.Items[5].Items[3].Visible := BEtatCycle.Enabled and (FRoleCompta >= rcReviseur);
  PopUPF11.Items[5].Items[4].Visible := BEtatCycle.Enabled and (FRoleCompta >= rcSuperviseur);

  // Changement d'activation du cycle
  PopUPF11.Items[6].Visible          := BActivation.Enabled;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 13/04/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPREVPROGTRAVAIL.OnClickBHistoMemo(Sender: TObject);
begin
  if GetDataSet.Bof and GetDataSet.Eof then Exit;
  CC_LanceFicheHistoCpte( GetField('CCY_CODECYCLE'), 'CCY');
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 07/03/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPREVPROGTRAVAIL.OnClickBHistoCycle( Sender : TObject );
begin
  if GetDataSet.Bof and GetDataSet.Eof then Exit;
  CPLanceFiche_CPREVHISTOCYCLE( GetField('CCY_CODECYCLE'), VH^.EnCours.Code );
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 30/03/2007
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPREVPROGTRAVAIL.OnClickAPG(Sender: TObject);
begin
  CPLanceFiche_CPREVINFODOSSIER( V_Pgi.NoDossier, VH^.EnCours.Code, 1, TaModif, 1);
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 02/04/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPREVPROGTRAVAIL.OnClickSCY(Sender: TObject);
begin
  CPLanceFiche_CPREVINFODOSSIER( V_Pgi.NoDossier, VH^.EnCours.Code, 1, TaModif, 2);
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 02/04/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPREVPROGTRAVAIL.OnClickEXP(Sender: TObject);
begin
  CPLanceFiche_CPREVINFODOSSIER( V_Pgi.NoDossier, VH^.EnCours.Code, 1, TaModif, 3);
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 07/03/2007
Modifié le ... :   /  /    
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPREVPROGTRAVAIL.OnClickEtatATraiter(Sender: TObject);
begin
  _ModifEtatCycleEnSerie( cATraiter );
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 23/04/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPREVPROGTRAVAIL.OnClickEtatEnCours( Sender : TObject );
begin
  _ModifEtatCycleEnSerie( cEncours );
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 19/03/2007
Modifié le ... :   /  /
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPREVPROGTRAVAIL.OnClickEtatEnSuspens(Sender: TObject);
begin
  _ModifEtatCycleEnSerie( cEnSuspens );
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 19/03/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPREVPROGTRAVAIL.OnClickEtatSupervise(Sender: TObject);
begin
  _ModifEtatCycleEnSerie( cSupervise );
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 19/03/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPREVPROGTRAVAIL.OnClickEtatValide(Sender: TObject);
begin
  _ModifEtatCycleEnSerie( cValide );
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 08/03/2007
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPREVPROGTRAVAIL.OnClickRendreActif(Sender: TObject);
begin
  _ModifActivationCycle( cActif );
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 08/03/2007
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPREVPROGTRAVAIL.OnClickRendreNonActif(Sender: TObject);
begin
  _ModifActivationCycle( cNonActif );
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 08/03/2007
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPREVPROGTRAVAIL.OnClickEtatInitial(Sender: TObject);
begin
  _ModifActivationCycle( cEtatInitial );
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 28/03/2007
Modifié le ... : 114/09/2007
Description .. :
Suite ........ : On envoit plus qu'un cycle lors de l'appel de CPCONSGENE
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPREVPROGTRAVAIL.OnClickDetailCycle(Sender: TObject);
//var lSt : string;
//    i : integer;
begin
  if (not BOuvrir.Enabled) then Exit;
  if GetDataSet.Bof and GetDataSet.Eof then Exit;

  (*
  if (FListe.AllSelected) or (FListe.nbSelected <> 0) then
  begin
    if FListe.AllSelected then
    begin
      lSt := '';
      TheMulQ.First;
      while not TheMulQ.Eof do
      begin
        lSt := lSt + TheMulQ.FindField('CCY_CODECYCLE').AsString + '||';
        TheMulQ.Next;
      end;
    end
    else
    begin
      for i := 0 to FListe.nbSelected - 1 do
      begin
        FListe.GotoLeBookMark(i);
        lSt := lSt + GetField('CCY_CODECYCLE') + '||';
      end;
    end;
  end
  else *)
  //  lSt := GetField('CCY_CODECYCLE');

  CPLanceFiche_CPCONSGENE( ';' + GetField('CCY_CODECYCLE'));
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 07/03/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPREVPROGTRAVAIL.OnKeyDownEcran( Sender : TObject; var Key: Word;  Shift: TShiftState);
begin
  case Key of

    Ord('A') : if (FListe.Focused) and (Shift = [ssShift]) then // Shift + A
                 _ModifActivationCycle( cActif )
               else
                 fOnSaveKeyDownEcran(Sender, Key, Shift);

    Ord('C') : if (FListe.Focused) and (Shift = [ssAlt]) then // Alt + C
                 _ModifEtatCycleEnSerie( cEnCours )
               else
                 fOnSaveKeyDownEcran(Sender, Key, Shift);

    Ord('E') : if (FListe.Focused) and (Shift = [ssAlt]) then // Alt + E
                 _ModifEtatCycleEnSerie( cEnSuspens )
               else
                 fOnSaveKeyDownEcran(Sender, Key, Shift);

    Ord('I') : if (FListe.Focused) and (Shift = [ssShift]) then // Shift + I
                 _ModifActivationCycle( cEtatInitial )
               else
                 fOnSaveKeyDownEcran(Sender, Key, Shift);

    Ord('N') : if (FListe.Focused) and (Shift = [ssShift]) then // Shift + N
                 _ModifActivationCycle( cNonActif )
               else
                 fOnSaveKeyDownEcran(Sender, Key, Shift);

    Ord('S') : if (FListe.Focused) and (Shift = [ssAlt]) then // ALt + S
               begin
                 // GCO - 30/05/2007 - FQ 20471
                 if JaiLeRoleCompta( RcSuperviseur ) then
                  _ModifEtatCycleEnSerie( cSupervise )
               end
               else
                 fOnSaveKeyDownEcran(Sender, Key, Shift);

    Ord('V') : if (FListe.Focused) and (Shift = [ssAlt]) then // Alt + V
               _ModifEtatCycleEnSerie( cValide )
              else
               fOnSaveKeyDownEcran(Sender, Key, Shift);

    Ord('T') : if (FListe.Focused) and (Shift = [ssAlt]) then // Alt + T
               _ModifEtatCycleEnSerie( cATraiter )
             else
               fOnSaveKeyDownEcran(Sender, Key, Shift);

    VK_F11 : PopUpF11.Popup(Mouse.CursorPos.x, Mouse.CursorPos.y);

  else
    fOnSaveKeyDownEcran(Sender, Key, Shift);
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 17/04/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPREVPROGTRAVAIL.OnClickBSelectAll(Sender: TObject);
begin
{$IFDEF EAGLCLIENT}
  if not FListe.AllSelected then
  begin
    if not TFMul(Ecran).FetchLesTous then
    begin
      PGIInfo('Sélection annulée. Impossible de récupérer tous les enregistrements.');
      Exit;
    end;
  end;
{$ENDIF}

  TFMUl(Ecran).bSelectAllClick( nil );
  _ActivationBouton;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 23/03/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPREVPROGTRAVAIL._ModifEtatCycleEnSerie( const vStEtatCycle: string );
var i : integer;
    lStEtatCycle : string;
    lRic : TRic;
begin
  // Blocage de la fonction d'activation d'un cycle si Dossier Supervise
  if (VH^.Revision.DossierSupervise) or
     (VH^.Revision.DossierPretSupervise and (not JaiLeRoleCompta( rcSuperviseur ))) then Exit;

{$IFDEF EAGLCLIENT}
  TheMulQ := TFMul(Ecran).Q.TQ;
{$ELSE}
  TheMulQ := TFMul(Ecran).Q;
{$ENDIF}
  lRic := nil;
  try
    if (FListe.AllSelected) or (FListe.nbSelected <> 0) then
    begin
      lRic := TRic.Create;
      if FListe.AllSelected then
      begin
        TheMulQ.First;
        while not TheMulQ.Eof do
        begin
          // GCO - 23/05/2007 - FQ 20331
          lStEtatCycle := GetColonneSql('CREVCYCLE', 'CCY_ETATCYCLE',
                          'CCY_CODECYCLE = "' + TheMulQ.FindField('CCY_CODECYCLE').AsString + '" AND ' +
                          'CCY_EXERCICE = "' + VH^.EnCours.Code + '"');

          if (TheMulQ.FindField('CCY_ACTIVECYCLE').AsString = 'X') then
          begin // Uniquement sur les cycles actifs
            if (lStEtatCycle = cSupervise) and (JaiLeRoleCompta(rcSuperviseur)) or
               (lStEtatCycle <> cSupervise) and (JaiLeRoleCompta(rcReviseur)) then
              lRic.ChangeEtatCycle( TheMulQ.FindField('CCY_CODECYCLE').AsString,
                               VH^.EnCours.Code, vStEtatCycle, V_Pgi.User, Date );
          end;
          TheMulQ.Next;
        end;
      end
      else
      begin
        for i := 0 to FListe.nbSelected - 1 do
        begin
          FListe.GotoLeBookMark(i);
          // GCO - 23/05/2007 - FQ 20331
          lStEtatCycle := GetColonneSql('CREVCYCLE', 'CCY_ETATCYCLE',
                          'CCY_CODECYCLE = "' + GetField('CCY_CODECYCLE') + '" AND ' +
                          'CCY_EXERCICE = "' + VH^.EnCours.Code + '"');

          if GetField('CCY_ACTIVECYCLE') = 'X' then
          begin // Uniquement sur les filtres actifs
            if (lStEtatCycle = cSupervise) and (JaiLeRoleCompta(rcSuperviseur)) or
               (lStEtatCycle <> cSupervise) and (JaiLeRoleCompta(rcReviseur)) then
              lRic.ChangeEtatCycle( GetField('CCY_CODECYCLE'), VH^.EnCours.Code,
                               vStEtatCycle, V_Pgi.User, Date );
          end;
        end;
      end;
    end;
  finally
    TheMulQ := nil; // GCO - 26/06/2007 - FQ 20844
    lRic.Free;
    TToolBarButton97(GetControl('BCHERCHE', True)).Click;
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 23/03/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPREVPROGTRAVAIL._ModifActivationCycle( const vStActivation : string );
var i : integer;
    lRic : TRic;
begin
  // Blocage de la fonction d'activation d'un cycle si Dossier Supervise
  if (VH^.Revision.DossierSupervise) or
     (VH^.Revision.DossierPretSupervise and (not JaiLeRoleCompta( rcSuperviseur ))) then Exit;

{$IFDEF EAGLCLIENT}
  TheMulQ := TFMul(Ecran).Q.TQ;
{$ELSE}
  TheMulQ := TFMul(Ecran).Q;
{$ENDIF}
  lRic := nil;
  try
    if (FListe.AllSelected) or (FListe.nbSelected <> 0) then
    begin
      lRic := TRic.Create;
      if FListe.AllSelected then
      begin
        TheMulQ.First;
        while not TheMulQ.Eof do
        begin
          lRic.ChangeActivationCycle( TheMulQ.FindField('CCY_CODECYCLE').AsString, VH^.EnCours.Code, vStActivation );
          TheMulQ.Next;
        end;
      end
      else
      begin
        for i := 0 to FListe.nbSelected - 1 do
        begin
          FListe.GotoLeBookMark(i);
          lRic.ChangeActivationCycle( GetField('CCY_CODECYCLE'), VH^.EnCours.Code, vStActivation );
        end;
      end;
    end;
  finally
    TheMulQ := nil; // GCO - 26/06/2007 - FQ 20844
    lRic.Free;
    TToolBarButton97(GetControl('BCHERCHE', True)).Click;
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 11/05/2007
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPREVPROGTRAVAIL.OnClickProgrammeTravail(Sender: TObject);
begin
  TRic.CPLanceEtat_ProgrammeTravail(0);
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 11/05/2007
Modifié le ... :   /  /
Description .. :
Mots clefs ... : // GCO - 01/08/2007 - FQ 21207
*****************************************************************}
procedure TOF_CPREVPROGTRAVAIL.OnClickNoteCtrlCycle( Sender : TObject );
var i : integer;
begin
{$IFDEF EAGLCLIENT}
  TheMulQ := TFMul(Ecran).Q.TQ;
{$ELSE}
  TheMulQ := TFMul(Ecran).Q;
{$ENDIF}

  if (FListe.AllSelected) or (FListe.nbSelected <> 0) then
  begin
    if FListe.AllSelected then
    begin
      TheMulQ.First;
      while not TheMulQ.Eof do
      begin
        TRic.CPLanceEtat_NoteCtrlCycle( GetField('CCY_CODECYCLE'), 0);
        TheMulQ.Next;
      end;
    end
    else
    begin
      for i := 0 to FListe.nbSelected - 1 do
      begin
        FListe.GotoLeBookMark(i);
        TRic.CPLanceEtat_NoteCtrlCycle( GetField('CCY_CODECYCLE'), 0);
      end;
    end;
  end
  else
    TRic.CPLanceEtat_NoteCtrlCycle( GetField('CCY_CODECYCLE'), 0);
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 11/05/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPREVPROGTRAVAIL.OnClickNoteSupervisionDossier( Sender : TObject );
begin
  TRic.CPLanceEtat_NoteSupervision( 0 );
end;

////////////////////////////////////////////////////////////////////////////////

Initialization
  registerclasses ( [ TOF_CPREVPROGTRAVAIL ] ) ;
end.
