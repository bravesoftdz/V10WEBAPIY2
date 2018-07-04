{***********UNITE*************************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 06/02/2006
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : RECHERCHEECR ()
Mots clefs ... : TOF;RECHERCHEECR
*****************************************************************}
Unit uTofRechercheEcr ;

Interface

Uses StdCtrls, 
     Controls, 
     Classes,
{$IFDEF EAGLCLIENT}
     eMul,
     uTob,
     MainEagl,      // AGLLanceFiche
{$ELSE}
     db,
     Hdb,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     Mul,
     FE_main,       // AGLLanceFiche
{$ENDIF}
     forms,
     sysutils,
     ComCtrls,
     Ent1,          // VH^.
     HCtrls,        // THMultiValComboBox
     Htb97,         // TToolBarButton97
     HQry,          // RecupWhereCritere
     Filtre,        // faPublic
     ParamSoc,      // GetParamSocSecur
     HEnt1,
     HMsgBox,
     UTOF ;

Type
  TOF_RECHERCHEECR = Class (TOF)

    E_DATECOMPTABLE   : THEdit;
    E_DATECOMPTABLE_  : THEdit;

    E_GENERAL         : THEdit;
    E_GENERAL_        : THEdit;

    E_AUXILIAIRE      : THEdit;
    E_AUXILIAIRE_     : THEdit;

    E_CONTREPARTIEGEN : THEdit;
    E_CONTREPARTIEGEN_: THEdit;

    E_CONTREPARTIEAUX : THEdit;
    E_CONTREPARTIEAUX_: THEdit;

    ELIBELLE          : THEdit;
    EREFINTERNE       : THEdit;
    XX_WHERE          : THEdit;

    PTABLELIBRE       : TTabSheet;

    BTOBV             : TToolBarButton97;

    EMONTANT          : THEdit;

    CBEtatPointage    : TCheckBox;
    CBAnalytique      : TCheckBox;

    E_NaturePiece   : THMultiValComboBox;
    E_QualifPiece   : THMultiValComboBox;
    E_Journal       : THMultiValComboBox;
    J_NatureJal     : THMultiValComboBox;
    E_EtatLettrage  : THMultiValComBoBox;
    E_Etablissement : THMultiValComboBox;

    PAGES           : TPageControl;
    FFiltres        : THValComboBox ;

    ComboMontant      : THValComboBox;
    ComboOPMontant    : THValComboBox;
    ComboOpLibelle    : THValComboBox;
    ComboOpRefInterne : THValComboBox;

    G_NatureGene    : THValComboBox;
    T_NatureAuxi    : THValComboBox;
    EExercice       : THValComboBox;

    PageControl     : TPageControl;

  {$IFDEF EAGLCLIENT}
    FListe : THGrid ;
  {$ELSE}
    FListe : THDBGrid ;
  {$ENDIF}

    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnLoad                   ; override ;
    procedure OnUpdate                 ; override ;

    procedure OnChangeEExercice      ( Sender : TObject );
    procedure OnChangeG_NatureGene   ( Sender : TObject );
    procedure OnChangeT_NatureAuxi   ( Sender : TObject );

    procedure OnClickBTobV           ( Sender : TObject );

    procedure OnExitE_General        ( Sender : TObject );
    procedure OnExitE_Auxiliaire     ( Sender : TObject );
    procedure OnDblClickFListe       ( Sender : TObject );

    procedure OnKeyDownEcran         ( Sender : TObject; var Key: Word; Shift: TShiftState);

    //procedure OnNew                    ; override ;
    //procedure OnDelete                 ; override ;
    //procedure OnCancel                 ; override ;
    procedure OnClose                  ; override ;
    procedure AuxiElipsisClick(Sender : TObject);

  private
    fBoConsultation : Boolean;
    fBoSaisieParam  : Boolean; // Test saisie paramétrable
    fOnSaveKeyDownEcran : procedure(Sender: TObject; var Key: Word; Shift: TShiftState) of object;

    procedure InitDefaut;
    procedure RecupMesCriteres;
    procedure OnMyItemNouveau( Sender : TObject );
  end;


function CPLanceFiche_CPRechercheEcr( vBoConsultation : Boolean; vBoSaisieParam : Boolean = False ) : string ;

Implementation

uses
  {$IFDEF MODENT1}
  CPTypeCons,
  CPProcMetier,
  {$ENDIF MODENT1}

{$IFDEF COMPTA}
  Saisie,	     // Pour Saisie eAGL
  SaisBor,           // TrouveEtLanceSaisie
  CPSaisiePiece_Tof, // TrouveEtLanceSaisieParam
  ConsultRevis_Tof,  // CPLanceFiche_ConsultRevis
{$ENDIF}
  AGLInit,           // TheMulQ
  uLibWindows,       // IIF
  uLibEcriture,      // CEstSaisieOuverte
  UTofMulParamGen,   {26/04/07 YMO F5 sur Auxiliaire }
  uLibExercice;      // CExerciceVersRelatif

const cOngletMouvement  = 0;
      cOngletJournal    = 1;
      cOngletCompte     = 2;
      cOngletComplement = 3;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 06/02/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
function CPLanceFiche_CPRechercheEcr( vBoConsultation : Boolean; vBoSaisieParam : Boolean = False ) : string ;
var lStTemp : string;
begin
  // Mode Consultation ou Modification
  if vBoConsultation then
    lStTemp :='X;'
  else
    lStTemp :='-;';

  // Saisie paramétrable
  if vBoSaisieParam then
    lStTemp := lStTemp + 'X;'
  else
    lStTemp := lStTemp + '-;';

  AGLLanceFiche('CP', 'CPRECHERCHEECR', '', '', lStTemp);
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 06/02/2006
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_RECHERCHEECR.OnArgument (S : String ) ;
begin
  Inherited ;
  fBoConsultation := (ReadTokenSt(S) = 'X');

  // Blocage en consultation de la recherche si un verrour de saisie est présent
  // dans la table Courrier
  if not FboConsultation then
  begin
    if CEstSaisieOuverte(False) then
      fBoConsultation := True;
  end;

  fBoSaisieParam  := (ReadTokenSt(S) = 'X');

  Ecran.Caption := IIF(fBoConsultation,
                       'Recherche de mouvements comptables (Consultation seulement)',
                       'Recherche de mouvements comptables (Mode modification)');
  UpdateCaption(Ecran);

  fOnSaveKeyDownEcran := Ecran.OnKeyDown;
  Ecran.OnKeyDown := OnKeyDownEcran;

  E_NaturePiece   := THMultiValComboBox(GetControl('E_NATUREPIECE', True));
  E_QualifPiece   := THMultiValComboBox(GetControl('E_QUALIFPIECE', True));
  E_Journal       := THMultiValComboBox(GetControl('E_JOURNAL', True));
  J_NatureJal     := THMultiValComboBox(GetControl('J_NATUREJAL', True));
  E_EtatLettrage  := THMultiValComBoBox(GetControl('E_ETATLETTRAGE', True));
  E_Etablissement := THMultiValComboBox(GetControl('E_ETABLISSEMENT', True));

  BTOBV              := TToolBarButton97(GetControl('BTOBV', True));

  E_DATECOMPTABLE    := THEdit(GetControl('E_DATECOMPTABLE', True));
  E_DATECOMPTABLE_   := THEdit(GetControl('E_DATECOMPTABLE_', True));

  E_GENERAL          := THEdit(GetControl('E_GENERAL', True));
  E_GENERAL_         := THEdit(GetControl('E_GENERAL_', True));

  E_AUXILIAIRE       := THEdit(GetControl('E_AUXILIAIRE', True));
  E_AUXILIAIRE_      := THEdit(GetControl('E_AUXILIAIRE_', True));

  E_CONTREPARTIEGEN  := THEdit(GetControl('E_CONTREPARTIEGEN', True));
  E_CONTREPARTIEGEN_ := THEdit(GetControl('E_CONTREPARTIEGEN_', True));

  E_CONTREPARTIEAUX  := THEdit(GetControl('E_CONTREPARTIEAUX', True));
  E_CONTREPARTIEAUX_ := THEdit(GetControl('E_CONTREPARTIEAUX_', True));

  EMONTANT           := THEdit(GetControl('EMONTANT', True));
  ELIBELLE           := THEdit(GetControl('ELIBELLE', True));
  EREFINTERNE        := THEdit(GetControl('EREFINTERNE', True));
  XX_WHERE           := THEdit(GetControl('XX_WHERE', True));

  PTABLELIBRE        := TTabSheet(GetControl('PTABLELIBRE', True));

  CBEtatPointage    := TCheckBox(GetControl('CBETATPOINTAGE', True));
  CBAnalytique      := TCheckBox(GetControl('CBANALYTIQUE', True));

  EExercice         := THValComboBox(GetControl('EEXERCICE', True));

  PAGES             := TPageControl(GetControl('PAGES', True));
  FFiltres          := THValComboBox(GetControl('FFILTRES'));

  ComboMontant      := THValComboBox(GetControl('CBMONTANT', True));
  ComboOpMontant    := THValComboBox(GetControl('CBOPMONTANT', True));
  ComboOpLibelle    := THValComboBox(GetControl('CBOPLIBELLE', True));
  ComboOpRefInterne := THValComboBox(GetControl('CBOPREFINTERNE', True));

  G_NatureGene    := THValComboBox(GetControl('G_NATUREGENE', True));
  T_NatureAuxi    := THValComboBox(GetControl('T_NATUREAUXI', True));

  PageControl     := TPageControl(GetControl('PAGES', True));

{$IFDEF EAGLCLIENT}
  FListe  := THGrid( GetControl('FLISTE',True));
  TheMulQ := TFMul(Ecran).Q.TQ;
{$ELSE}
  FListe  := THDBGrid( GetControl('FLISTE',True));
  TheMulQ := TFMul(Ecran).Q;
{$ENDIF}

  // Branchement des événements
  CInitComboExercice(EExercice);
  EExercice.OnChange := OnChangeEExercice;

  E_General.OnExit          := OnExitE_General;
  E_General_.OnExit         := OnExitE_General;

  E_ContrepartieGen.OnExit  := OnExitE_General;
  E_ContrepartieGen_.OnExit := OnExitE_General;

  E_Auxiliaire.OnExit       := OnExitE_Auxiliaire;
  E_Auxiliaire_.OnExit      := OnExitE_Auxiliaire;

  E_ContrepartieAux.OnExit  := OnExitE_Auxiliaire;
  E_ContrepartieAux_.OnExit := OnExitE_Auxiliaire;

  G_NatureGene.OnChange    := OnChangeG_NatureGene;
  T_NatureAuxi.OnChange    := OnchangeT_NatureAuxi;

  BTobV.OnClick            := OnClickBTobV;

  FListe.OnDblClick        := OnDblClickFListe ;

  InitDefaut;

  // GCO - 29/08/2006 - FQ 18456
  THEdit(GetControl('E_TABLE0', True)).DataType := 'TZNATECR0';
  THEdit(GetControl('E_TABLE1', True)).DataType := 'TZNATECR1';
  THEdit(GetControl('E_TABLE2', True)).DataType := 'TZNATECR2';
  THEdit(GetControl('E_TABLE3', True)).DataType := 'TZNATECR3';
  // FIN GCO

  if GetParamSocSecur('SO_CPMULTIERS', false) then
  begin
    E_AUXILIAIRE.OnElipsisClick:=AuxiElipsisClick;
    E_AUXILIAIRE_.OnElipsisClick:=AuxiElipsisClick;
  end;

end ;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 06/02/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_RECHERCHEECR.InitDefaut;
var lStTous : string;
begin
  lStTous := TraduireMemoire('<<Tous>>');

  E_Journal.SelectAll;
  E_Journal.Text := lStTous;

  E_NaturePiece.SelectAll;
  E_NaturePiece.Text := lStTous;

  E_QualifPiece.SelectAll;
  E_QualifPiece.Text := lStTous;

  E_Journal.SelectAll;
  E_Journal.Text := lStTous;

  J_NatureJal.SelectAll;
  J_NatureJal.Text := lStTous;

  E_EtatLettrage.SelectAll;
  E_EtatLettrage.Text := lStTous;

  E_Etablissement.SelectAll;
  E_Etablissement.Text := lStTous;

  G_NatureGene.ItemIndex := 0;
  T_NatureAuxi.ItemIndex := 0;

  if CtxPcl in V_Pgi.PgiContexte then
  begin
    if vH^.CPExoRef.Code <> '' then
      EExercice.Value := CExerciceVersRelatif( VH^.CPExoRef.Code )
    else
      EExercice.Value := CExerciceVersRelatif( VH^.Encours.Code );
  end
  else
    EExercice.Value := CExerciceVersRelatif( VH^.Entree.Code );

  ComboMontant.Value   := '(E_DEBIT OR E_CREDIT)';

  ComboOpMontant.Plus  := ' AND CO_CODE <> "C" AND CO_CODE <> "D"' +
                       ' AND CO_CODE <> "L" AND CO_CODE <> "M"' +
                       ' AND CO_CODE <> "V" ANd CO_CODE <> "VV"';

  ComboOpMontant.Value := '=';

  // Opérateur Libellé
  ComboOpLibelle.Plus  := ' AND CO_CODE <> "<" AND CO_CODE <> "<="' +
                          ' AND CO_CODE <> ">" AND CO_CODE <> ">="' +
                          ' AND CO_CODE <> "E" AND CO_CODE <> "G"' +
                          ' AND CO_CODE <> "I" AND CO_CODE <> "J"';
                          
  ComboOpLibelle.Value := 'L';

  // Opérateur Référence Interne
  ComboOpRefInterne.Plus  := ComboOpLibelle.Plus;
  ComboOpRefInterne.Value := 'L';

  E_General.MaxLength           := VH^.Cpta[fbGene].Lg;
  E_General_.MaxLength          := E_General.MaxLength;
  E_ContrepartieGen.MaxLength   := E_General.MaxLength;
  E_ContrepartieGen_.MaxLength  := E_General.MaxLength;

  E_Auxiliaire.MaxLength        := VH^.Cpta[fbAux].Lg;
  E_Auxiliaire_.MaxLength       := E_Auxiliaire.MaxLength;
  E_ContrepartieAux.MaxLength   := E_Auxiliaire.MaxLength;
  E_ContrepartieAux_.MaxLength  := E_Auxiliaire.MaxLength;
end;


////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 06/02/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_RECHERCHEECR.OnDisplay () ;
begin
  Inherited ;
end ;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 14/02/2006
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_RECHERCHEECR.OnKeyDownEcran(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
  case key of

   71 : begin // CTRL + G -> Onglet Comptes
          if ssCtrl in Shift then
          begin
            PageControl.ActivePage := PageControl.Pages[cOngletCompte];
            if E_General.CanFocus then E_General.SetFocus;
          end;
        end;

   74 : begin // CTRL + J -> Onglet Journal des données
          if ssCtrl in Shift then
          begin
            PageControl.ActivePage := PageControl.Pages[cOngletJournal];
            if E_Journal.CanFocus then E_Journal.SetFocus;
          end;
        end;

   77 : begin // CTRL + M -> Onglet Mouvement
           if ssCtrl in Shift then
           begin
             PageControl.ActivePage := PageControl.Pages[cOngletMouvement];
             if EMontant.CanFocus then EMontant.SetFocus;
           end;
        end;

   84 : begin // Ctrl + T -> Onglet Complément
           if ssCtrl in Shift then
           begin
             PageControl.ActivePage := PageControl.Pages[cOngletComplement];
             if E_EtatLettrage.CanFocus then E_EtatLettrage.SetFocus;
           end;
        end;

  else
    fOnSaveKeyDownEcran(Sender, Key, Shift);
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 15/02/2006
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_RECHERCHEECR.OnDblClickFListe(Sender: TObject);
{$IFDEF COMPTA}
var AA       : TActionFiche;
    lDossier : String ;
    lBoLocal : Boolean ;
{$ENDIF}
begin
{$IFDEF COMPTA}
  if GetDataSet.Bof and GetDataSet.Eof then Exit ;

  if FBoConsultation then
    AA := taConsult
  else
    AA := taModif;

  // Réaffectation systématique de TheMulQ car vaut nil au 2ème passage
{$IFDEF EAGLCLIENT}
  TheMulQ := TFMul(Ecran).Q.TQ;
  TheMulQ.Seek( FListe.Row - 1 );
{$ELSE}
  TheMulQ := TFMul(Ecran).Q;
{$ENDIF}
  lDossier := V_PGI.SchemaName ;
  lBoLocal := True ;

  if (GetControlText('MULTIDOSSIER') <> '') then // EstMultiSoc and ...
  begin
    lDossier := GetField('SYSDOSSIER') ;
    // Consultation multi-Soc
    if lDossier <> V_PGI.SchemaName then
    begin
      lBoLocal := False ;
      if not V_PGI.Sav then
        AA := taConsult ;
    end ;
  end ;

  (*
  if ((sMode <> '-') and (sMode <> '')) then
  begin
    if lBoLocal then
      LanceSaisieFolio(TheMulQ, AA) ;
  end
  else
  begin
    if FBoSaisieParam or ( not lBoLocal ) then // Saisie param ou consultation multi-Soc
      TrouveEtLanceSaisieParam(TheMulQ, AA, GetControlText('E_QUALIFPIECE'), FALSE, lDossier)
    else
      TrouveEtLanceSaisie(TheMulQ, AA, GetControlText('E_QUALIFPIECE'));
  end;*)

  if FBoSaisieParam or ( not lBoLocal ) then // Saisie param ou consultation multi-Soc
    TrouveEtLanceSaisieParam(TheMulQ, AA, TheMulQ.FindField('E_QUALIFPIECE').AsString, False, lDossier)
  else
    TrouveEtLanceSaisie(TheMulQ, AA, TheMulQ.FindField('E_QUALIFPIECE').AsString);

  // GCO - 30/05/2006 - FQ 10641
  if CtxPcl in V_Pgi.PgiContexte then
    TFMUL(ECRAN).BChercheClick(nil)
  else
  begin
    if GetParamSocSecur('SO_BDEFGCPTE', False) then
      TFMUL(ECRAN).BChercheClick(nil);
  end;

{$ENDIF}
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 07/02/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_RECHERCHEECR.OnChangeEExercice(Sender: TObject);
begin
  if EExercice.ItemIndex = 0 then
  begin
    if VH^.ExoV8.Code <> '' then
      E_DateComptable.Text := DateToStr(VH^.ExoV8.Deb)
    else
      // Date de début du premier Exercice
      E_DateComptable.Text := DateToStr(VH^.Exercices[1].Deb);
    E_DateComptable_.Text  := DateToStr(iDate2099-1);
  end
  else
  begin
    //ExoToDates(EExercice.Value, E_DATECOMPTABLE, E_DATECOMPTABLE_);
    CExoRelatifToDates(EExercice.Value, E_DATECOMPTABLE, E_DATECOMPTABLE_);
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 07/02/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_RECHERCHEECR.OnExitE_General(Sender: TObject);
begin
  if (Trim(THEdit(Sender).Text) = '') or (TestJoker(THEdit(Sender).Text)) then Exit;

  if Length(THEdit(Sender).Text) < VH^.Cpta[fbGene].Lg then
    THEdit(Sender).Text := BourreEtLess( THEdit(Sender).Text, fbGene);
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 07/02/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_RECHERCHEECR.OnExitE_Auxiliaire(Sender: TObject);
begin
  if (Trim(THEdit(Sender).Text) = '') or (TestJoker(THEdit(Sender).Text)) then Exit;

  if Length(THEdit(Sender).Text) < VH^.Cpta[fbAux].Lg then
    THEdit(Sender).Text := BourreEtLess( THEdit(Sender).Text, fbAux);
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 13/02/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_RECHERCHEECR.RecupMesCriteres;
var lStTemp : string;
begin
  // GCO - 04/05/2007 - FQ 19540 
  XX_WHERE.Text := '(E_QUALIFPIECE = "N" OR ' +
                    'E_QUALIFPIECE = "S" OR ' +
                    'E_QUALIFPIECE = "U")';

  if (ComboMontant.Value <> '') and
     (ComboOpMontant.Value <> '') and
     (EMontant.Text <> '') then
  begin
    lStTemp := TraduitOperateur(ComboOpMontant.Value, EMontant.Text, False);
    if lStTemp <> '' then
    begin
      if (ComboMontant.Value = 'E_DEBIT') then
        XX_Where.Text := XX_Where.Text + ' AND ' + ComboMontant.Value + ' ' + lStTemp + ' AND E_CREDIT = 0'
      else
        if(ComboMontant.Value = 'E_CREDIT') then
          XX_Where.Text := XX_Where.Text + ' AND ' + ComboMontant.Value + ' ' + lStTemp + ' AND E_DEBIT = 0'
        else
          XX_Where.Text := XX_Where.Text + ' AND ((E_DEBIT ' + lStTemp + ' AND E_CREDIT = 0) OR (E_CREDIT ' + lStTemp +' AND E_DEBIT = 0))';
    end;
  end;

  // GCO - 02/05/2006 - FQ 17790
  if EExercice.ItemIndex <> 0 then
  begin
    XX_Where.Text := XX_Where.Text + ' AND E_EXERCICE = "' + CRelatifVersExercice(EEXERCICE.Value) + '"';
  end;

  if (ELibelle.Text <> '') or (ComboOpLibelle.Value = 'V') or (ComboOpLibelle.Value = 'VV') then
  begin
    lStTemp := TraduitOperateur(ComboOpLibelle.Value, ELibelle.Text, True);
    if lStTemp <> '' then
    begin
      if XX_WHere.Text <> '' then XX_WHERE.Text :=  XX_Where.Text + ' AND ';
      XX_WHere.Text := XX_WHere.TEXT + ' E_LIBELLE ' + lStTemp;
    end;
  end;

  if (ERefInterne.Text <> '') or (ComboOpRefInterne.Value = 'V') or (ComboOpRefInterne.Value = 'VV') then
  begin
    lStTemp := TraduitOperateur(ComboOpRefInterne.Value, ERefinterne.Text, True);
    if lStTemp <> '' then
    begin
      if XX_Where.Text <> '' then XX_Where.Text :=  XX_Where.Text + ' AND ';
      XX_Where.Text := XX_WHere.Text + ' E_REFINTERNE ' + lStTemp;
    end;
  end;

  if CBEtatPointage.State <> CBGrayed then
  begin
    XX_Where.Text := XX_Where.Text + ' AND E_REFPOINTAGE ' + IIF(CBEtatPointage.Checked, ' <> ""', '= ""');
  end;

  if CBAnalytique.State <> CBGrayed then
  begin
    XX_Where.Text := XX_Where.Text + ' AND E_ANA = ' + IIF(CBAnalytique.Checked, '"X"', '"-"');
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 06/02/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_RECHERCHEECR.OnLoad ;
begin
  Inherited ;
  RecupMesCriteres;

  LibellesTableLibre(PTableLibre, 'TE_TABLE', 'E_TABLE', 'E');

  //TheMulQ.SQL.Text := CMajRequeteExercice (E_EXERCICE.Value, TheMulQ.SQL.Text);

  TFMUl(Ecran).ListeFiltre.OnItemNouveau := OnMyItemNouveau;

{$IFDEF CCSTD}
  TFMul(Ecran).ListeFiltre.ForceAccessibilite := FaPublic;
{$ENDIF}
end ;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 01/09/2006
Modifié le ... :   /  /    
Description .. : FQ 18534
Mots clefs ... : 
*****************************************************************}
procedure TOF_RECHERCHEECR.OnMyItemNouveau(Sender: TObject);
begin
  videFiltre( FFiltres, Pages ) ;
  TFMUL(Ecran).ListeFiltre.new ;

  InitDefaut;

  ComboMontant.Value   := '(E_DEBIT OR E_CREDIT)';
  ComboOPMontant.Value := '=';
  EMontant.Text     := '';
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 07/02/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_RECHERCHEECR.OnUpdate ;
begin
  Inherited ;
end ;

////////////////////////////////////////////////////////////////////////////////

(*
procedure TOF_RECHERCHEECR.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_RECHERCHEECR.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_RECHERCHEECR.OnCancel () ;
begin
  Inherited ;
end ;*)

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 06/02/2006
Modifié le ... : 18/09/2006
Description .. : LG - 18/09/2006 - TheNUlQ = nil sion les fenetre AGL 
Suite ........ : plante apres l'appel de cette fenetre
Mots clefs ... : 
*****************************************************************}
procedure TOF_RECHERCHEECR.OnClose ;
begin
 TheMulQ := nil ;
 Inherited ;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : YMO
Créé le ...... : 26/04/2007
Modifié le ... :   /  /
Description .. : Branchement de la fiche auxiliaire
Mots clefs ... :
*****************************************************************}
procedure TOF_RECHERCHEECR.AuxiElipsisClick( Sender : TObject );
begin
     THEdit(Sender).text:= CPLanceFiche_MULTiers('M;' +THEdit(Sender).text + ';' +THEdit(Sender).Plus + ';');
end;


////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 13/02/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_RECHERCHEECR.OnChangeG_NatureGene(Sender: TObject);
begin
  if G_NatureGene.ItemIndex = 0 then
    E_General.Plus := ''
  else
    E_General.Plus := ' AND G_NATUREGENE = "' + G_NatureGene.Value + '"';

  E_ContrePartieGen.Plus := E_General.Plus;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 13/02/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... : 
*****************************************************************}
procedure TOF_RECHERCHEECR.OnChangeT_NatureAuxi(Sender: TObject);
begin
  if T_NatureAuxi.ItemIndex = 0 then
    E_Auxiliaire.Plus := ''
  else
    E_Auxiliaire.Plus := ' AND T_NATUREAUXI = "' + T_NatureAuxi.Value + '"';

  E_Auxiliaire.Plus := E_Auxiliaire.Plus + ' AND T_NATUREAUXI <> "NCP"' +
                       ' AND T_NATUREAUXI <> "CON" AND T_NATUREAUXI <> "PRO"' +
                       ' AND T_NATUREAUXI <> "SUS"';

  E_ContrePartieAux.Plus := E_Auxiliaire.Plus;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 21/03/2006
Modifié le ... :   /  /    
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_RECHERCHEECR.OnClickBTobV(Sender: TObject);
{$IFDEF COMPTA}
var WhereSQL : String ;
{$ENDIF}
begin
{$IFDEF COMPTA}
  if ((E_Journal.Text = '') or (E_Journal.Tous)) and
     (E_General.Text = '') and
     (E_Auxiliaire.Text = '') then
  begin
    if PGIAsk('Vous n''avez pas renseigné de journal ni de comptes. ' +
              'La sélection peut être importante. Confirmez-vous l''analyse ?',
              Ecran.Caption) <> mrYes then Exit ;
  end;

  WhereSQL := RecupWhereCritere(TPageControl(GetControl('Pages', True))) ;
  CPLanceFiche_ConsultRevis(WhereSql);
{$ENDIF}  
end;

////////////////////////////////////////////////////////////////////////////////

Initialization
  registerclasses ( [ TOF_RECHERCHEECR ] ) ;
end.
