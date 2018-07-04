{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 03/06/2015
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BTFAMILLEMATERIEL ()
Mots clefs ... : TOF;BTFAMILLEMATERIEL
*****************************************************************}
Unit BTMATERIEL_TOF;

Interface

Uses StdCtrls, 
     Controls, 
     Classes, 
     db,
     {$IFNDEF DBXPRESS}
     dbtables,
     {$ELSE}
     uDbxDataSet,
     {$ENDIF}
     fe_main,
     mul,
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HPanel,
     HTB97,
     HEnt1,
     HMsgBox,
     uTOB,
     Paramsoc,
     LookUp,
     UtilsGrille,
     UtilsEtat,
     Vierge,
     HSysMenu,
     HRichEdt,
     HRichOLE,
     UTOF;

Type
  TOF_BTMATERIEL = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;

  private
    //Variable nécessaire pour la gestion de l'état
    OptionEdition : TOptionEdition;
    fEtat         : THValComboBox;
    TheType       : String;
    TheNature     : String;
    TheTitre      : String;
    TheModele     : String;
    //
    ValConsoPresta: Boolean;
    //
    Action        : TActionFiche;
    //
    PageCTRL      : TPageControl;
    //Panel #1
    Panel1        : THPanel;
    CodeMateriel  : ThEdit;
    LibMateriel   : THEdit;
    CodeRessource : THEdit;
    LibRessource  : THLabel;
    Fermer        : TCheckBox;
    //Panel #2
    Panel2        : THPanel;
    CodeFamille   : THEdit;
    LibFamille    : THLabel;
    //Panel #3
    Panel3        : THPanel;
    CodeArticle   : THEdit;
    LibArticle    : THLabel;
    //Panel #4
    LastAction    : THLabel;
    //Panel #6
    Panel6        : THGroupBox;
    LibAlpha1     : THLabel;
    LibAlpha2     : THLabel;
    LibAlpha3     : THLabel;
    LibAlpha4     : THLabel;
    LibAlpha5     : THLabel;
    LibAlpha6     : THLabel;
    LibAlpha7     : THLabel;
    LibAlpha8     : THLabel;
    LibAlpha9     : THLabel;
    LibAlpha10    : THLabel;
    ValAlpha1     : THedit;
    ValAlpha2     : THedit;
    ValAlpha3     : THedit;
    ValAlpha4     : THedit;
    ValAlpha5     : THedit;
    ValAlpha6     : THedit;
    ValAlpha7     : THedit;
    ValAlpha8     : THedit;
    ValAlpha9     : THedit;
    ValAlpha10    : THedit;
    //
    Panel7        : THPanel;
    Panel8        : THGroupBox;
    //Panel #9
    DateCreation  : THEdit;
    DateModif     : THEdit;
    TBMA_DATESUPP : TLabel;
    DateSupp      : THEdit;
    UserCreate    : THEdit;
    UserModif     : THEdit;
    //Panel #10
    Panel10       : THGroupBox;
    Libdate1      : THLabel;
    Libdate2      : THLabel;
    Libdate3      : THLabel;
    Libdate4      : THLabel;
    Valdate1      : THedit;
    Valdate2      : THedit;
    Valdate3      : THedit;
    Valdate4      : THedit;
    //Panel #11
    Panel11       : THGroupBox;
    ValBoolean1   : TCheckBox;
    ValBoolean2   : TCheckBox;
    ValBoolean3   : TCheckBox;
    ValBoolean4   : TCheckBox;
    //
    Panel12       : THPanel;
    //Panel #13
    Panel13       : THPanel;
    ValPa         : THNumEdit;
    ValPr         : THNumEdit;
    ValPv         : THNumEdit;
    ValAmort      : THNumEdit;
    BtCout        : TToolBarButton97;
    CoutRevient   : THNumEdit;
    Unite         : THValComboBox;
    LibUnite      : THLabel;
    BTAffect      : TToolBarButton97;
    BtEvent       : TToolbarButton97;
    //Panel #15
    Panel15       : THPanel;
    Borne1        : THEdit;
    Borne2        : THEdit;
    PeriodeHeure  : THNumedit;
    PeriodeJour   : THNumedit;
    UtilHeure     : THNumedit;
    Utiljour      : THNumedit;
    UtilPercent   : THNumedit;
    ImmoHeure     : THNumedit;
    ImmoJour      : THNumedit;
    ImmoPercent   : THNumedit;
    EntretienHeure: THNumedit;
    entretienJour : THNumedit;
    entretienPercent: THNumedit;
    NotUseHeure   : THNumedit;
    NotUseJour    : THNumedit;
    NotUsePercent : THNumedit;
    Amortissement : THNumedit;
    CoutEntretien : THNumedit;
    TotalCoutPeriode: THNumedit;
    CtRevientJour : THNumEdit;
    CtRevientHeure: THNumEdit;
    //
    Panel16       : THPanel;
    Panel17       : THPanel;

    BlocNote      : THRichEditOLE;

    //Panel #19
    Panel19       : THPanel;
    GrilleAffect  : THGrid;
    GGrilleAffect : TGestionGS;
    //
    GGrilleEvent  : TGestionGS;
    GrilleEvent   : THGrid;
    //
    PEvent        : TTabSheet;
    Paffect       : TTabSheet;
    //
    TobFamMat     : Tob;
    TobMateriel   : Tob;
    TobRessource  : TOB;
    TobArticle    : TOB;
    TobAffect     : Tob;
    TobEvent      : Tob;
    TobEdition    : TOB;
    //
    StSQL         : String;
    OldCodeR      : String;
    SCodeFamille  : string;
    //
    QQ            : TQuery;
    //
    BtDelete      : TToolbarButton97;
    Btdelete1     : TToolbarButton97;
    BtParamList   : TToolbarButton97;
    BtNouveau     : TToolbarButton97;
    BTImprimer    : TToolbarButton97;
    BTDuplication : TToolbarButton97;
    //
    FF            : String;
    //
    procedure AffichageInitEcran(OkAff: Boolean);
    //
    procedure ArticleOnElipsisClick(Sender: Tobject);
    procedure ArticleOnExit  (Sender: Tobject);
    //
    procedure BTAffectOnClick(Sender : TObject);
    procedure BTCoutOnclick  (Sender : TObject);
    Procedure BTEventOnClick (Sender : TObject);
    procedure BtParamListOnClick(Sender: Tobject);
    //
    procedure ChargeInfoAffect;
    procedure ChargeInfoArticle;
    procedure ChargeInfoEvenement;
    Procedure ChargeInfoFamMat;
    procedure ChargeInfoRessource;
    procedure ChargeTobEdition;
    procedure ChargeZoneEcran;
    procedure ChargeZoneTOB;
    procedure Controlechamp(Champ, Valeur: String);
    procedure CreateEdition;
    procedure CreateTOB;

    procedure DateSuppOnExit(Sender: Tobject);
    procedure DblClickGrille(Sender : TObject);
    procedure DestroyTOB;

    Procedure FamMatOnExit(Sender: Tobject);
    procedure FamMatOnElipsisClick(Sender: Tobject);
    procedure FermerOnClick(Sender: Tobject);

    procedure GetObjects;

    procedure InitGrilleAffect;
    procedure InitGrilleEvent;
    procedure InitTobEdition;

    procedure MaterielOnExit(Sender: tobject);

    procedure OnDelEnregGrille(Sender: Tobject);
    procedure OnDupliqueFiche(Sender: TObject);
    procedure OngletOnenter(Sender: Tobject);
    procedure OnImprimeFiche(Sender: TObject);
    procedure OnNewRecord(Sender: TObject);

    procedure RAZZoneEcran;
    procedure RazZoneFamMat;
    procedure RessourceOnExit(Sender: Tobject);
    procedure RessourceOnElipsisClick(Sender: Tobject);
    procedure SetScreenEvents;


  end ;

const
	// libellés des messages
  TexteMessage: array[1..19] of string  = (
          {1}   'Le Code Materiel est obligatoire'
          {2},  'Suppression impossible ce Matériel se trouve sur un évènement'
          {3},  'Suppression impossible ce Matériel se trouve sur une Affectation'
          {4},  'La suppression a échoué'
          {5},  'Ce code matériel existe déjà. Veuillez le modifier'
          {6},  'Suppression de l''événement impossible'
          {7},  'Suppression de l''Affectation impossible'
          {8},  'L''identifiant est obligatoire dans la liste, veuillez le paramétrer si vous voulez faire une suppression'
          {9},  'L''identifiant est obligatoire dans la liste, veuillez le paramétrer si vous voulez faire une Modification'
          {10}, 'La Ressource est obligatoire'
          {11}, 'Code Materiel inexistant'
          {12}, 'Code Famille  Inexistant'
          {13}, 'Code Ressource Inexistant'
          {14}, 'Code Prestation Inexistant'
          {15}, 'Attention, cette ressource est déjà présente sur un matériel'
          {16}, 'Le calcul ne peut pas être à cheval sur deux années'
          {17}, 'La famille materiel est obligatoire'
          {18}, 'Les valeurs de Prestation ont changé voulez-vous les récupérer ?'
          {19}, 'Les valeurs de la Ressource ont changé voulez-vous les récupérer ?'
                                         );

Implementation

Uses  TiersUtil,
      AGLInitGC,
      ParamDBG,
      UtilsParc,
      UtilRessource,
      BTPUtil,
      TntGrids,
      EntGC;

procedure TOF_BTMATERIEL.OnArgument (S : String ) ;
var Critere : string;
    Champ   : string;
    Valeur  : string;
    i       : Integer;
    x       : Integer;
begin
  Inherited ;
  //
  TFvierge(Ecran).FormResize := false;
  //Chargement des zones ecran dans des zones programme
  GetObjects;
  //
  CreateTOB;
  //
  CreateEdition;
  //
  InitGrilleAffect;
  //
  InitGrilleEvent;
  //
  Critere := uppercase(Trim(ReadTokenSt(S)));
  while Critere <> '' do
  begin
     x := pos('=', Critere);
     if x <> 0 then
        begin
        Champ  := copy(Critere, 1, x - 1);
        Valeur := copy(Critere, x + 1, length(Critere));
        end
     else
        Champ  := Critere;
     ControleChamp(Champ, Valeur);
     Critere:= uppercase(Trim(ReadTokenSt(S)));
  end;

  FF:='#';
  if V_PGI.OkDecV>0 then
  begin
    FF:='# ##0.';
    for i:=1 to V_PGI.OkDecV-1 do
    begin
      FF:=FF+'0';
    end;
    FF:=FF+'0';
  end;

  ValConsoPresta := GetParamSocSecur('SO_BTVALOSCONSO', False);

  //Gestion des évènement des zones écran
  SetScreenEvents;

  //Remise à blanc des zone ecran
  RAZZoneEcran;

  //Affichage de l'écran initial
  AffichageInitEcran(False);
  //
  BtParamList.Visible   := False;
  BtDelete1.Visible     := False;
  BtNouveau.Visible     := False;
  Btimprimer.Visible    := True;
  BTDuplication.Visible := True;

  if action = tacreat then
  begin
    BtDelete.Visible      := False;
    Btimprimer.Visible    := False;
    BTDuplication.Visible := False;
    //
    DateCreation.Text := DateToStr(Now);
    UserCreate.text   := V_PGI.User;
  end;
  //
  DateModif.text      := DateToStr(Now);
  UserModif.text      := V_PGI.User;
  //
  AppliqueFontDefaut(BlocNote);
  //
  THSystemMenu(getControl('HMtrad')).ResizeGridColumns(GrilleAffect);
  THSystemMenu(getControl('HMtrad')).ResizeGridColumns(GrilleEvent);


end ;

procedure TOF_BTMATERIEL.OnNew ;
begin
  Inherited ;
  //Initialisation des zones écran
  RAZZoneEcran;
  //
end ;

procedure TOF_BTMATERIEL.OnDelete ;
begin
  Inherited ;

  if PGIAsk('Confirmez-vous la suppression de ce matériel : ' + CodeMateriel.Text + ' ?', 'Matériel') = Mryes then
  begin
    //contrôle si Matériel présent sur évènements
    StSQl := 'SELECT BEM_CODEMATERIEL FROM BTEVENTMAT WHERE BEM_CODEMATERIEL = "' + CodeMateriel.text + '" ';
    if existeSQL(STSQL) then
      PGIError(TexteMessage[2], 'Matériel')
    else
    Begin
      //contrôle si Matériel présent sur affectations
      StSQL := 'SELECT BFF_CODEMATERIEL FROM BTAFFECTATION WHERE BFF_CODEMATERIEL="' + CodeMateriel.Text + '"';
      if ExisteSQL(StSQL) then
        PGIError(TexteMessage[3], 'Matériel')
      else
      begin
        //suppression pure et simple de l'enregistrement avec confirmation
        StSQL := 'DELETE BTMATERIEL WHERE BMA_CODEMATERIEL="' + CodeMateriel.Text + '"';
        if ExecuteSQL(StSQL)=0 then PGIError(TexteMessage[4], ' Matériel');
      end;
    end;
  end;

  Ecran.ModalResult := 1;

end ;

procedure TOF_BTMATERIEL.OnUpdate ;
begin
  Inherited ;

  Ecran.ModalResult := 0;
  PageCTRL.ActivePageIndex := 0;

  if  CodeMateriel.Text = '' then
  begin
    PGIError(TexteMessage[11], 'Erreur Saisie');
    CodeMateriel.SetFocus;
    Exit;
  end;

  if CodeFamille.Text = '' then
  begin
    PGIError(TexteMessage[17], 'Erreur Saisie');
    CodeFamille.SetFocus;
    Exit;
  end;

  if not FamilleMatExist(CodeFamille.text, '')     then
  begin
    PGIError(TexteMessage[12], 'Erreur Saisie');
    CodeFamille.SetFocus;
    Exit;
  end;

  if not PrestationExist(CodeArticle.text, 'AND GA_TYPEARTICLE= "PRE" AND BNP_TYPERESSOURCE IN ("MAT","LOC", "OUT")') then
  begin
    PGIError(TexteMessage[14], 'Erreur Saisie');
    CodeArticle.SetFocus;
    Exit;
  end;

  if not RessourceExist(CodeRessource.text, ' AND ' + CodeRessource.Plus) then
  begin
    PGIError(TexteMessage[13], 'Erreur Saisie');
    codeRessource.SetFocus;
    Exit;
  end;

  Ecran.ModalResult := 1;

  ChargeZoneTOB;                                  

  //mise à jour de la table Famille Matériel
  TobMateriel.InsertOrUpdateDB(true);


end ;

procedure TOF_BTMATERIEL.OnLoad ;
begin
  Inherited ;

  //chargement du Matériel
  StSQL := 'SELECT * FROM BTMATERIEL WHERE BMA_CODEMATERIEL="' + CodeMateriel.Text + '"';
  QQ := OpenSQL(StSQL, False);

  If not QQ.Eof then
  begin
    TobMateriel.SelectDB('BTMATERIEL', QQ);

    //chargement des zones à l'écran
    ChargeZoneEcran;

    if Fermer.checked then
    begin
      TBMA_DATESUPP.visible := True;
      DateSupp.Visible := True;
    end;

    //chargement des évèenements
    ChargeInfoEvenement;

    ChargeInfoAffect;
    //
    DateModif.text := DateToStr(Now);
    UserModif.text := V_PGI.User;
    //
    //BTCoutOnclick(Self);
    //Initialisation de la tob d'édition nécessaire pour imprimer la fiche matériel
    InitTobEdition;
    //
    ChargeTobEdition;
  end;

  Ferme(QQ);

end ;

procedure TOF_BTMATERIEL.OnClose ;
begin
  Inherited ;

  DestroyTOB;

end ;

procedure TOF_BTMATERIEL.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_BTMATERIEL.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_BTMATERIEL.GetObjects;
begin

  PageCTRL      := TPageControl(GetControl('ONGLETMATERIEL'));
  FEtat         := ThValComboBox(ecran.FindComponent('FEtat'));
  //Panel #1
  CodeMateriel  := THEdit(Getcontrol('BMA_CODEMATERIEL'));
  LibMateriel   := THEdit(Getcontrol('BMA_LIBELLE'));
  CodeRessource := THEdit(Getcontrol('BMA_RESSOURCE'));
  LibRessource  := THLabel(Getcontrol('LIBRESSOURCE'));
  Fermer        := TCheckBox(Getcontrol('BMA_FERME'));
  //Panel #2
  CodeFamille   := THEdit(Getcontrol('BMA_CODEFAMILLE'));
  LibFamille    := THLabel(Getcontrol('LIBFAMILLEMAT'));
  //Panel #3
  CodeArticle   := THEdit(Getcontrol('BMA_CODEARTICLE'));
  LibArticle    := THLabel(Getcontrol('LIBARTICLE'));
  //Panel #4
  LastAction    := THLabel(Getcontrol('LASTACTION'));
  //Panel #5     =
  //Panel #6
  LibAlpha1     := THLabel(Getcontrol('ALPHA1'));
  LibAlpha2     := THLabel(Getcontrol('ALPHA2'));
  LibAlpha3     := THLabel(Getcontrol('ALPHA3'));
  LibAlpha4     := THLabel(Getcontrol('ALPHA4'));
  LibAlpha5     := THLabel(Getcontrol('ALPHA5'));
  LibAlpha6     := THLabel(Getcontrol('ALPHA6'));
  LibAlpha7     := THLabel(Getcontrol('ALPHA7'));
  LibAlpha8     := THLabel(Getcontrol('ALPHA8'));
  LibAlpha9     := THLabel(Getcontrol('ALPHA9'));
  LibAlpha10    := THLabel(Getcontrol('ALPHA10'));
  ValAlpha1     := THedit(Getcontrol('BMA_VALALPHA1'));
  ValAlpha2     := THedit(Getcontrol('BMA_VALALPHA2'));
  ValAlpha3     := THedit(Getcontrol('BMA_VALALPHA3'));
  ValAlpha4     := THedit(Getcontrol('BMA_VALALPHA4'));
  ValAlpha5     := THedit(Getcontrol('BMA_VALALPHA5'));
  ValAlpha6     := THedit(Getcontrol('BMA_VALALPHA6'));
  ValAlpha7     := THedit(Getcontrol('BMA_VALALPHA7'));
  ValAlpha8     := THedit(Getcontrol('BMA_VALALPHA8'));
  ValAlpha9     := THedit(Getcontrol('BMA_VALALPHA9'));
  ValAlpha10    := THedit(Getcontrol('BMA_VALALPHA10'));
  //Panel #10    =
  Libdate1      := THLabel(Getcontrol('DATE1'));
  Libdate2      := THLabel(Getcontrol('DATE2'));
  Libdate3      := THLabel(Getcontrol('DATE3'));
  Libdate4      := THLabel(Getcontrol('DATE4'));
  Valdate1      := THedit(Getcontrol('BMA_VALDATE1'));
  Valdate2      := THedit(Getcontrol('BMA_VALDATE2'));
  Valdate3      := THedit(Getcontrol('BMA_VALDATE3'));
  Valdate4      := THedit(Getcontrol('BMA_VALDATE4'));
  //Panel #11    =
  ValBoolean1   := TCheckBox(GetControl('BMA_VALBOOLEAN1'));
  ValBoolean2   := TCheckBox(GetControl('BMA_VALBOOLEAN2'));
  ValBoolean3   := TCheckBox(GetControl('BMA_VALBOOLEAN3'));
  ValBoolean4   := TCheckBox(GetControl('BMA_VALBOOLEAN4'));
  //Panel #9     =
  DateCreation  := THEdit(GetCONTROL('BMA_DATECREATION'));
  DateModif     := THEdit(GetCONTROL('BMA_DATEMODIF'));
  TBMA_DATESUPP := TLabel(GetCONTROL('TBMA_DATESUPP'));
  DateSupp      := THEdit(GetCONTROL('BMA_DATESUPP'));
  UserCreate    := THEdit(GetCONTROL('BMA_CREATEUR'));
  UserModif     := THEdit(GetCONTROL('BMA_UTILISATEUR'));
  //Panel #13    =
  ValPa         := THNumEdit(GetCONTROL('BMA_PA'));
  ValPr         := THNumEdit(GetCONTROL('BMA_PR'));
  ValPv         := THNumEdit(GetCONTROL('BMA_PV'));
  ValAmort      := THNumEdit(GetCONTROL('BMA_VALAMORT'));

  CoutRevient   := THNumEdit(GetCONTROL('COUTREVIENT'));
  Unite         := THValComboBox(GetCONTROL('UNITE'));
  LibUnite      := THLabel(GetCONTROL('TLAB'));
  BTAffect      := TToolBarButton97(GetCONTROL('BTAFFECT'));
  BTEvent       := TToolbarButton97(GetControl('BtEvent'));
  BtDelete      := TToolBarButton97(GetCONTROL('BDELETE'));
  Btdelete1     := TToolBarButton97(GetCONTROL('BDELETE1'));
  BtNouveau     := TToolBarButton97(GetCONTROL('BNOUVEAU'));
  BtParamList   := TToolbarButton97(GetControl('BPARAMLISTE'));
  Btimprimer    := TToolbarButton97(GetControl('BIMPRIMER'));
  BTDuplication := TToolbarButton97(GetControl('BDUPLICATION'));
  //Panel #15    =
  BtCout        := TToolbarButton97(GetControl('BTCOUT'));
  Borne1        := THEdit(GetCONTROL('BORNE1'));
  Borne2        := THEdit(GetCONTROL('BORNE1_'));
  PeriodeHeure  := THNumedit(GetCONTROL('PERIODEHEURE'));
  PeriodeJour   := THNumedit(GetCONTROL('PERIODEJOUR'));
  UtilHeure     := THNumedit(GetCONTROL('UTILHEURE'));
  Utiljour      := THNumedit(GetCONTROL('UTILJOUR'));
  UtilPercent   := THNumedit(GetCONTROL('UTILPERCENT'));
  ImmoHeure     := THNumedit(GetCONTROL('IMMOHEURE'));
  ImmoJour      := THNumedit(GetCONTROL('IMMOJOUR'));
  ImmoPercent   := THNumedit(GetCONTROL('IMMOPERCENT'));
  EntretienHeure:= THNumedit(GetCONTROL('ENTRETIENHEURE'));
  entretienJour := THNumedit(GetCONTROL('ENTRETIENJOUR'));
  entretienPercent:= THNumedit(GetCONTROL('ENTRETIENPERCENT'));
  NotUseHeure   := THNumedit(GetCONTROL('NOTUSEHEURE'));
  NotUseJour    := THNumedit(GetCONTROL('NOTUSEJOUR'));
  NotUsePercent := THNumedit(GetCONTROL('NOTUSEPERCENT'));
  Amortissement := THNumedit(GetCONTROL('AMORTISSEMENT'));
  CoutEntretien := THNumedit(GetCONTROL('COUTENTRETIEN'));
  TotalCoutPeriode := THNumedit(GetCONTROL('TOTALCOUTPERIODE'));
  CtRevientJour := THNumEdit(GetCONTROL('CTREVIENTJOUR'));
  CtRevientHeure:= THNumEdit(GetCONTROL('CTREVIENTHEURE'));

  GrilleAffect  := THGrid(GetControl('GRILLEAFFECT'));
  GrilleEvent   := THGrid(GetControl('GRILLEEVENT'));

  PEvent        := TTabSheet(GetControl('PEVENEMENT'));
  Paffect       := TTabSheet(GetControl('PAFFECTATION'));

  BlocNote      := THRichEditOLE(Getcontrol('BMA_BLOCNOTE'));

  Panel1        := THPanel(GetCONTROL('PANEL1'));
  Panel2        := THPanel(GetCONTROL('PANEL2'));
  Panel3        := THPanel(GetCONTROL('PANEL3'));
  Panel6        := THGroupBox(GetCONTROL('GROUPALPHA'));
  Panel7        := THPanel(GetCONTROL('PANEL7'));
  Panel8        := THGroupBox(GetCONTROL('PANEL8'));
  Panel10       := THGroupBox(GetCONTROL('GROUPDATE'));
  Panel11       := THGroupBox(GetCONTROL('GROUPBOOLEAN'));
  Panel12       := THPanel(GetCONTROL('PANEL12'));
  Panel13       := THPanel(GetCONTROL('PANEL13'));
  Panel15       := THPanel(GetCONTROL('PANEL15'));
  Panel16       := THPanel(GetCONTROL('PANEL16'));
  Panel17       := THPanel(GetCONTROL('PANEL17'));
  Panel19       := THPanel(GetCONTROL('PANEL19'));
  //

end;

procedure TOF_BTMATERIEL.SetScreenEvents;
begin

  //on va déjà gérer les boutons...
  BTEvent.OnClick               := BTEventOnClick;
  BTAffect.OnClick              := BTAffectOnClick;
  BtParamList.OnClick           := BtParamListOnClick;
  BtDelete1.OnClick             := OnDelEnregGrille;
  BtNouveau.OnClick             := OnNewRecord;
  BtImprimer.OnClick            := OnImprimeFiche;
  BtDuplication.OnClick         := OnDupliqueFiche;
  //
  GrilleEvent.OnDblClick        := DblClickGrille;
  GrilleAffect.OnDblClick       := DblClickGrille;
  //
  CodeMateriel.OnExit           := MaterielOnExit;
  CodeRessource.OnExit          := RessourceOnExit;
  CodeFamille.OnExit            := FamMatOnExit;
  CodeArticle.OnExit            := ArticleOnExit;
  DateSupp.OnExit               := DateSuppOnExit;

  Fermer.OnClick                := FermerOnClick;
  ValAmort.OnExit               := BTCoutOnclick;
  //
  if Assigned(Borne1) then   Borne1.OnExit  := BTCoutOnclick;
  if Assigned(Borne2) then   Borne2.OnExit  := BTCoutOnclick;
  //
  if Assigned(BtCout) then   BtCout.OnClick := BtCoutOnClick;
  //
  CodeRessource.OnElipsisClick  := RessourceOnElipsisClick;
  CodeRessource.Plus            := 'ARS_TYPERESSOURCE IN ("OUT","MAT","LOC","AUT")';
  CodeFamille.OnElipsisClick    := FamMatOnElipsisClick;
  CodeArticle.OnElipsisClick    := ArticleOnElipsisClick;

  PageCTRL.OnChange             := OngletOnenter;

end;

Procedure TOF_BTMATERIEL.DateSuppOnExit(Sender : Tobject);
Var DateDel : TDateTime;
begin

  DateDel := StrToDate(DateSupp.text);

  if ((DateDel <> iDate1900) and (Datedel <> iDate2099))  And (not Fermer.Checked) then Fermer.checked := True;

end;


Procedure TOF_BTMATERIEL.FermerOnClick(Sender : Tobject);
Var DateDel : TDateTime;
begin

  if fermer.checked then
    DateDel := now
  else
    DateDel := idate2099;

  DateSupp.text := DateToStr(DateDel);

end;

Procedure TOF_BTMATERIEL.Controlechamp(Champ, Valeur : String);
begin

  if Champ='ACTION' then
  begin
    if      Valeur='CREATION'     then Action:=taCreat
    else if Valeur='MODIFICATION' then Action:=taModif
    else if Valeur='CONSULTATION' then Action:=taConsult;
  end;

  if Champ='CODEMATERIEL'         then CodeMateriel.text  := Valeur
  else if Champ = 'CODERESSOURCE' then CodeRessource.Text := Valeur;

end;

Procedure TOF_BTMATERIEL.CreateTOB;
begin

  TobMateriel   := Tob.Create('BTMATERIEL', nil, -1);
  TobRessource  := TOB.Create('RESSOURCE' , nil, -1);
  TobArticle    := TOB.Create('ARTICLE'   , nil, -1);
  TobFamMat     := TOB.Create('FAMILLEMAT', nil, -1);
  //****************************
  // Affectation
  //****************************
  TobAffect     := TOB.Create('BTAFFECTATION',nil, -1);
  //****************************
  // Evènement
  //****************************
  TobEvent      := TOB.Create('BTEVENTMAT'  , nil, -1);
  //****************************
  // Edition
  //****************************
  TobEdition    := TOB.Create(' UN MATERIEL', nil, -1);
  //
end;

procedure TOF_BTMATERIEL.CreateEdition;
begin
  //
  TheType       := 'E';
  TheNature     := 'PMA';
  TheTitre      := 'Fiche Matériel';
  TheModele     := 'EFM';
  //
  OptionEdition := TOptionEdition.Create(TheType,TheNature,TheModele, TheTitre, '', True, True, True, False, False, PageCTRL, fEtat);
  //
  OptionEdition.Apercu    := True;
  OptionEdition.DeuxPages := False;
  OptionEdition.Spages    := PageCTRL;

end;

procedure TOF_BTMATERIEL.DestroyTOB;
begin

  FreeAndNil(TobMateriel);
  FreeAndNil(TobRessource);
  FreeAndNil(TobArticle);
  FreeAndNil(TobFamMat);

  //Destruction des objets... c'était plus simple de les mettre ici !!!
  FreeAndNil(GGrilleAffect);
  FreeAndNil(GGrilleEvent);

  FreeAndNil(TobEdition);
  FreeAndNil(OptionEdition);

end;

{***********UNITE*************************************************
Auteur  ...... : FV1
Créé le ...... : 03/06/2015
Modifié le ... :   /  /
Description .. : Gestion des différents chargement de TOB ou d'écran
Mots clefs ... : Chargement, TOB, Table, Ecran
*****************************************************************}
Procedure TOF_BTMATERIEL.RAZZoneEcran;
Var Annee   : Word;
    Mois    : Word;
    Jour    : Word;
begin
  //Panel #1
  //CodeMateriel.Text   := '';
  LibMateriel.Text      := '';
  CodeRessource.Text    := '';
  LibRessource.Caption  := '';
  Fermer.checked        := False;
  //Panel #2
  CodeFamille.text      := '';
  LibFamille.Caption    := '';
  //Panel #3
  CodeArticle.text      := '';
  LibArticle.caption    := '';
  //Panel #4
  LastAction.caption    := '';
  //Panel #6
  LibAlpha1.caption     := '';
  LibAlpha2.caption     := '';
  LibAlpha3.caption     := '';
  LibAlpha4.caption     := '';
  LibAlpha5.caption     := '';
  LibAlpha6.caption     := '';
  LibAlpha7.caption     := '';
  LibAlpha8.caption     := '';
  LibAlpha9.caption     := '';
  LibAlpha10.caption    := '';
  //
  ValAlpha1.text        := '';
  ValAlpha2.text        := '';
  ValAlpha3.text        := '';
  ValAlpha4.text        := '';
  ValAlpha5.text        := '';
  ValAlpha6.text        := '';
  ValAlpha7.text        := '';
  ValAlpha8.text        := '';
  ValAlpha9.text        := '';
  ValAlpha10.text       := '';
  //Panel #10
  Libdate1.caption      := '';
  Libdate2.caption      := '';
  Libdate3.caption      := '';
  Libdate4.caption      := '';
  Valdate1.Text         := DateTimeToStr(idate2099);
  Valdate2.Text         := DateTimeToStr(idate2099);
  Valdate3.Text         := DateTimeToStr(idate2099);
  Valdate4.Text         := DateTimeToStr(idate2099);
  //Panel #11
  ValBoolean1.checked   := False;
  ValBoolean2.checked   := False;
  ValBoolean3.checked   := False;
  ValBoolean4.checked   := False;
  //Panel #9
  DateCreation.Text     := DateTimeToStr(idate1900);
  DateModif.text        := DateTimeToStr(Idate1900);
  DateSupp.text         := DateTimeToStr(idate2099);
  UserCreate.text       := '';
  UserModif.text        := '';
  //Panel #13
  Unite.Value           := '003';
  //Panel #14
  //Panel #15
  DecodeDate(now,Annee,Mois,Jour);
  Borne1.text := '01/01/' + IntToStr(Annee); //==> début d'année
  Borne2.text := '31/12/' + IntToStr(Annee); //==> fin d'année
  //
  ValPa.Text            := FormatFloat(FF, 0);
  ValPr.Text            := FormatFloat(FF, 0);
  ValPv.Text            := FormatFloat(FF, 0);
  ValAmort.Text         := FormatFloat(FF, 0);
  CoutRevient.Text      := FormatFloat(FF, 0);

  PeriodeHeure.Text     := FormatFloat(FF, 0);
  PeriodeJour.Text      := FormatFloat(FF, 0);
  UtilHeure.Text        := FormatFloat(FF, 0);
  Utiljour.Text         := FormatFloat(FF, 0);
  UtilPercent.Text      := FormatFloat(FF, 0);
  ImmoHeure.Text        := FormatFloat(FF, 0);
  ImmoJour.Text         := FormatFloat(FF, 0);
  ImmoPercent.Text      := FormatFloat(FF, 0);
  EntretienHeure.Text   := FormatFloat(FF, 0);
  entretienJour.Text    := FormatFloat(FF, 0);
  entretienPercent.Text := FormatFloat(FF, 0);
  NotUseHeure.Text      := FormatFloat(FF, 0);
  NotUseJour.Text       := FormatFloat(FF, 0);
  NotUsePercent.Text    := FormatFloat(FF, 0);
  Amortissement.Text    := FormatFloat(FF, 0);
  CoutEntretien.Text    := FormatFloat(FF, 0);
  TotalCoutPeriode.Text := FormatFloat(FF, 0);
  CtRevientJour.Text    := FormatFloat(FF, 0);
  CtRevientHeure.Text   := FormatFloat(FF, 0);

  BlocNote.Clear;

  SCodeFamille           := '';

end;

Procedure ToF_BTMATERIEL.RazZoneFamMat;
begin

  //la zone contenant les zones alpha est rendue invisible au chargement
  Panel6.Visible  := false;

  //la zone contenant les zones date est rendue invisible au chargement
  Panel10.Visible  := false;

  //la zone contenant les zones choix est rendue invisible au chargement
  Panel11.Visible  := false;

  //Panel #2
  LibFamille.Caption    := '';

  //Panel #6
  ValAlpha1.text        := '';
  ValAlpha2.text        := '';
  ValAlpha3.text        := '';
  ValAlpha4.text        := '';
  ValAlpha5.text        := '';
  ValAlpha6.text        := '';
  ValAlpha7.text        := '';
  ValAlpha8.text        := '';
  ValAlpha9.text        := '';
  ValAlpha10.text       := '';

  //Panel #10
  Libdate1.caption      := '';
  Libdate2.caption      := '';
  Libdate3.caption      := '';
  Libdate4.caption      := '';

  //Panel #11
  ValBoolean1.Caption   := 'Valeur O/N';
  ValBoolean2.Caption   := 'Valeur O/N';
  ValBoolean3.caption   := 'Valeur O/N';
  ValBoolean4.Caption   := 'Valeur O/N';

end;

Procedure TOF_BTMATERIEL.ChargeZoneEcran;
Var Annee   : Word;
    Mois    : Word;
    Jour    : Word;
begin

    //Panel #1
    CodeMateriel.Text     := TobMateriel.GetString('BMA_CODEMATERIEL');
    if Action = TaModif then CodeMateriel.Enabled := False;

    LibMateriel.Text      := TobMateriel.GetString('BMA_LIBELLE');
    //
    CodeRessource.Text    := TobMateriel.GetString('BMA_RESSOURCE');
    //
    if TobMateriel.GetString('BMA_FERME')='X' then
      Fermer.checked      := True
    else
      Fermer.checked      := False;

    //Panel #2
    CodeFamille.text      := TobMateriel.GetString('BMA_CODEFAMILLE');
    SCodeFamille          := CodeFamille.text;

    //Panel #3
    CodeArticle.text      := TobMateriel.GetString('BMA_CODEARTICLE');

    //Panel #4
    LastAction.caption    := 'Penser à charger l''évènement quand la saisie fonctionnera';

    //Panel #6
    ValAlpha1.text        := TobMateriel.GetString('BMA_VALALPHA1');
    ValAlpha2.text        := TobMateriel.GetString('BMA_VALALPHA2');
    ValAlpha3.text        := TobMateriel.GetString('BMA_VALALPHA3');
    ValAlpha4.text        := TobMateriel.GetString('BMA_VALALPHA4');
    ValAlpha5.text        := TobMateriel.GetString('BMA_VALALPHA5');
    ValAlpha6.text        := TobMateriel.GetString('BMA_VALALPHA6');
    ValAlpha7.text        := TobMateriel.GetString('BMA_VALALPHA7');
    ValAlpha8.text        := TobMateriel.GetString('BMA_VALALPHA8');
    ValAlpha9.text        := TobMateriel.GetString('BMA_VALALPHA9');
    ValAlpha10.text       := TobMateriel.GetString('BMA_VALALPHA10');
    //Panel #10
    Valdate1.Text         := TobMateriel.GetValue('BMA_VALDATE1');;
    Valdate2.Text         := TobMateriel.GetValue('BMA_VALDATE2');;
    Valdate3.Text         := TobMateriel.GetValue('BMA_VALDATE3');;
    Valdate4.Text         := TobMateriel.GetValue('BMA_VALDATE4');;
    //Panel #11
    if TobMateriel.GetString('BMA_VALBOOLEAN1')='X' then
      ValBoolean1.checked   := True
    else
      ValBoolean1.checked   := False;
    //
    if TobMateriel.GetString('BMA_VALBOOLEAN2')='X' then
      ValBoolean2.checked   := True
    else
      ValBoolean2.checked   := False;

    if TobMateriel.GetString('BMA_VALBOOLEAN3')='X' then
      ValBoolean3.checked   := True
    else
      ValBoolean3.checked   := False;

    if TobMateriel.GetString('BMA_VALBOOLEAN4')='X' then
      ValBoolean4.checked   := True
    else
      ValBoolean4.checked   := False;

    //Panel #9
    DateCreation.Text     := TobMateriel.GetString('BMA_DATECREATION');
    DateModif.text        := TobMateriel.GetString('BMA_DATEMODIF');
    DateSupp.text         := TobMateriel.GetString('BMA_DATESUPP');
    UserCreate.text       := TobMateriel.GetString('BMA_CREATEUR');
    UserModif.text        := TobMateriel.GetString('BMA_UTILISATEUR');
    //Panel #13
    ValPa.Text            := FormatFloat(FF, TobMateriel.GetValue('BMA_PA'));
    ValPr.Text            := FormatFloat(FF, TobMateriel.GetValue('BMA_PR'));
    ValPv.Text            := FormatFloat(FF, TobMateriel.GetValue('BMA_PV'));
    ValAmort.Text         := FormatFloat(FF, TobMateriel.GetValue('BMA_VALAMORT'));
    //
    CoutRevient.Text      := FormatFloat(FF, TobMateriel.GetValue('BMA_COUT'));
    Unite.Value           := TobMateriel.GetValue('BMA_UNITE');
    //Panel #14
    //Panel #15
    DecodeDate(now,Annee,Mois,Jour);
    //
    Borne1.text := '01/01/' + IntToStr(Annee); //==> début d'année
    Borne2.text := '31/12/' + IntToStr(Annee); //==> fin d'année
    //
    PeriodeHeure.Text     := FormatFloat(FF, 0);
    PeriodeJour.Text      := FormatFloat(FF, 0);
    UtilHeure.Text        := FormatFloat(FF, 0);
    Utiljour.Text         := FormatFloat(FF, 0);
    UtilPercent.Text      := FormatFloat(FF, 0);
    ImmoHeure.Text        := FormatFloat(FF, 0);
    ImmoJour.Text         := FormatFloat(FF, 0);
    ImmoPercent.Text      := FormatFloat(FF, 0);
    EntretienHeure.Text   := FormatFloat(FF, 0);
    entretienJour.Text    := FormatFloat(FF, 0);
    entretienPercent.Text := FormatFloat(FF, 0);
    NotUseHeure.Text      := FormatFloat(FF, 0);
    NotUseJour.Text       := FormatFloat(FF, 0);
    NotUsePercent.Text    := FormatFloat(FF, 0);
    Amortissement.Text    := FormatFloat(FF, 0);
    CoutEntretien.Text    := FormatFloat(FF, 0);
    TotalCoutPeriode.Text := FormatFloat(FF, 0);
    CtRevientJour.Text    := FormatFloat(FF, 0);
    CtRevientHeure.Text   := FormatFloat(FF, 0);
    //

    BTCoutOnclick(self);

    BlocNote.text          := TobMateriel.GetValue('BMA_BLOCNOTE');

    ChargeInfoRessource;

    //ChargeInfoArticle;

    ChargeInfoFamMat;

end;

Procedure TOF_BTMATERIEL.ChargeZoneTOB;
begin

    //Panel #1
    TobMateriel.PutValue('BMA_CODEMATERIEL', CodeMateriel.Text);
    TobMateriel.PutValue('BMA_LIBELLE', LibMateriel.text);
    //
    TobMateriel.PutValue('BMA_RESSOURCE', CodeRessource.Text);
    //
    if Fermer.Checked then
      TobMateriel.PutValue('BMA_FERME','X')
    else
      TobMateriel.PutValue('BMA_FERME','-');

    //Panel #2
    TobMateriel.Putvalue('BMA_CODEFAMILLE',CodeFamille.text);
    //Panel #3
    TobMateriel.PutValue('BMA_CODEARTICLE',CodeArticle.text);
    //Panel #6
    //
    TobMateriel.Putvalue('BMA_VALALPHA1',ValAlpha1.text);
    TobMateriel.Putvalue('BMA_VALALPHA2',ValAlpha2.text);
    TobMateriel.Putvalue('BMA_VALALPHA3',ValAlpha3.text);
    TobMateriel.Putvalue('BMA_VALALPHA4',ValAlpha4.text);
    TobMateriel.Putvalue('BMA_VALALPHA5',ValAlpha5.text);
    TobMateriel.Putvalue('BMA_VALALPHA6',ValAlpha6.text);
    TobMateriel.Putvalue('BMA_VALALPHA7',ValAlpha7.text);
    TobMateriel.Putvalue('BMA_VALALPHA8',ValAlpha8.text);
    TobMateriel.Putvalue('BMA_VALALPHA9',ValAlpha9.text);
    TobMateriel.Putvalue('BMA_VALALPHA10',ValAlpha10.text);
    //Panel #10
    TobMateriel.Putvalue('BMA_VALDATE1',StrtoDate(Valdate1.Text));
    TobMateriel.Putvalue('BMA_VALDATE2',StrtoDate(Valdate2.Text));
    TobMateriel.Putvalue('BMA_VALDATE3',StrtoDate(Valdate3.Text));
    TobMateriel.Putvalue('BMA_VALDATE4',StrtoDate(Valdate4.Text));

    //Panel #11
    if ValBoolean1.checked then
      TobMateriel.Putvalue('BMA_VALBOOLEAN1','X')
    else
      TobMateriel.Putvalue('BMA_VALBOOLEAN1','-');
    //
    if ValBoolean2.checked then
      TobMateriel.Putvalue('BMA_VALBOOLEAN2','X')
    else
      TobMateriel.Putvalue('BMA_VALBOOLEAN2','-');
    //
    if ValBoolean3.checked then
      TobMateriel.Putvalue('BMA_VALBOOLEAN3','X')
    else
      TobMateriel.Putvalue('BMA_VALBOOLEAN3','-');
    //
    if ValBoolean4.checked then
      TobMateriel.Putvalue('BMA_VALBOOLEAN4','X')
    else
      TobMateriel.Putvalue('BMA_VALBOOLEAN4','-');

    //Panel #9
    TobMateriel.Putvalue('BMA_DATECREATION',  StrToDate(DateCreation.Text));
    TobMateriel.Putvalue('BMA_DATEMODIF',     StrToDate(DateModif.text));
    TobMateriel.Putvalue('BMA_DATESUPP',      StrToDate(DateSupp.text));
    TobMateriel.Putvalue('BMA_CREATEUR',      UserCreate.text);
    TobMateriel.Putvalue('BMA_UTILISATEUR',   UserModif.text);

    //Panel #13
    TobMateriel.Putvalue('BMA_PA',      ValPa.Text);
    TobMateriel.Putvalue('BMA_PR',      ValPr.Text);
    TobMateriel.Putvalue('BMA_PV',      ValPv.Text);
    //
    TobMateriel.Putvalue('BMA_VALAMORT',ValAmort.Text);
    TobMateriel.Putvalue('BMA_COUT',    CoutRevient.Text);
    TobMateriel.Putvalue('BMA_UNITE',   Unite.Value);

    TobMateriel.PutValue('BMA_BLOCNOTE',BlocNote.Text);

end;

procedure TOF_BTMATERIEL.ChargeInfoArticle;
begin

  if CodeArticle.Text = '' then exit;

  If not Assigned(TobArticle) then Exit;

  StSQl := 'SELECT * FROM ARTICLE WHERE GA_CODEARTICLE = "' + CodeArticle.text;
  StSQL := StSQL + '"  OR GA_ARTICLE= "' + CodeArticle.Text;
  StSQL := StSQL + '" AND GA_TYPEARTICLE="PRE"';

  QQ := OpenSQL(StSQL, False);

  If Not QQ.eof then
  begin
    TobArticle.SelectDB('', QQ);
    //
    LibArticle.Caption    := TobArticle.GetString('GA_LIBELLE');
    if ValConsoPresta then
    begin
      if Action = TaModif then
      begin
        if (TobMateriel.GetValue('BMA_PA') <> TobArticle.GetValue('GA_PAHT')) Or
           (TobMateriel.GetValue('BMA_PR') <> TobArticle.GetValue('GA_DPR'))  Or
           (TobMateriel.GetValue('BMA_PV') <> TobArticle.GetValue('GA_PVHT')) Then
        begin
          if PGIAsk(TexteMessage[18], 'Lecture Prestation') = MrYes then
          begin
            ValPa.Text         := FormatFloat(FF, TobArticle.GetValue('GA_PAHT'));
            ValPR.Text         := FormatFloat(FF, TobArticle.GetValue('GA_DPR'));
            ValPV.Text         := FormatFloat(FF, TobArticle.GetValue('GA_PVHT'));
            CoutRevient.Text   := FormatFloat(FF, StrTofloat(ValPr.Text));
          end;
        end;
      end
      else
      begin
        ValPa.Text         := FormatFloat(FF, TobArticle.GetValue('GA_PAHT'));
        ValPR.Text         := FormatFloat(FF, TobArticle.GetValue('GA_DPR'));
        ValPV.Text         := FormatFloat(FF, TobArticle.GetValue('GA_PVHT'));
        CoutRevient.Text   := FormatFloat(FF, StrTofloat(ValPr.Text));
      end;
    end;
  end;

  Ferme(QQ);

end;

Procedure TOF_BTMATERIEL.ChargeInfoRessource;
begin

  //Je n'enlève pas ce petit bout de code on sait jamais ça peut revenir aussi vite que c'est partit
  {if CodeRessource.text = '' then
  begin
     PGIError(TexteMessage[10], 'Matériel');
     CodeRessource.SetFocus;
     Exit;
  end;}

  If not Assigned(TobRessource)    then Exit;

  if OldCodeR = CodeRessource.text then Exit;

  if CodeRessource.text = '' then
  begin
    LibRessource.Caption := '';
    if not ValConsoPresta then
    begin
      ValPa.Text         := FormatFloat(FF, 0);
      ValPR.Text         := FormatFloat(FF, 0);
      ValPV.Text         := FormatFloat(FF, 0);
      CoutRevient.Text   := FormatFloat(FF, ValPr.Value);
    end;
    Exit;
  end;

  //contrôle si cette ressource n'est pas déjà sur un matériel
  StSQL := 'SELECT 1 FROM BTMATERIEL WHERE BMA_RESSOURCE="' + CodeRessource.text + '"';
  StSQL := StSQL + ' AND BMA_CODEMATERIEL <> "' + CodeMateriel.text + '"';
  if ExisteSQL(StSQL) then
  begin
     PGIError(TexteMessage[15], 'Matériel');
     CodeRessource.text := '';
     CodeRessource.SetFocus;
     Exit;
  end;

  StSQl := 'SELECT * FROM RESSOURCE WHERE ARS_RESSOURCE="' + Coderessource.text;
  StSQL := StSQL + '" AND ARS_TYPERESSOURCE IN ("OUT","MAT","LOC","AUT")';

  QQ := OpenSQL(StSQL, False);

  If Not QQ.eof then
  begin
    TobRessource.SelectDB('', QQ);
    //
    If CodeMateriel.text  = '' then CodeMateriel.text  := TobRessource.GetString('ARS_RESSOURCE');
    if LibMateriel.Text   = '' then LibMateriel.Text   := TobRessource.GetString('ARS_LIBELLE');
    if CodeArticle.text   = '' then
    begin
      CodeArticle.Text   := TobRessource.GetString('ARS_ARTICLE');
      ChargeInfoArticle;
    end;
    LibRessource.Caption := TobRessource.GetString('ARS_LIBELLE');
    BlocNote.Text        := TobRessource.GetString('ARS_BLOCNOTE');
    if not ValConsoPresta then
    begin
      if Action = TaModif then
      begin
        if (TobMateriel.GetValue('BMA_PA') <> TobRessource.GetValue('ARS_TAUXUNIT'))       Or
           (TobMateriel.GetValue('BMA_PR') <> TobRessource.GetValue('ARS_TAUXREVIENTUN'))  Or
           (TobMateriel.GetValue('BMA_PV') <> TobRessource.GetValue('ARS_PVHT')) Then
        begin
          if PGIAsk(TexteMessage[19], 'Lecture Ressource') = MrYes then
          begin
            ValPa.Text         := FormatFloat(FF, TobRessource.GetValue('ARS_TAUXUNIT'));
            ValPR.Text         := FormatFloat(FF, TobRessource.GetValue('ARS_TAUXREVIENTUN'));
            ValPV.Text         := FormatFloat(FF, TobRessource.GetValue('ARS_PVHT'));
            CoutRevient.Text   := FormatFloat(FF, StrTofloat(ValPr.Text));
          end;
        end;
      end
      else
      begin
        ValPa.Text         := FormatFloat(FF, TobRessource.GetValue('ARS_TAUXUNIT'));
        ValPR.Text         := FormatFloat(FF, TobRessource.GetValue('ARS_TAUXREVIENTUN'));
        ValPV.Text         := FormatFloat(FF, TobRessource.GetValue('ARS_PVHT'));
        CoutRevient.Text   := FormatFloat(FF, StrTofloat(ValPr.Text));
      end;
    end;
  end;

  OldCodeR := CodeRessource.Text;

  Ferme(QQ);

end;

Procedure TOF_BTMATERIEL.ChargeInfoFamMat;
begin

  StSQl := 'SELECT * FROM BTFAMILLEMAT WHERE BFM_CODEFAMILLE="' + CodeFamille.text + '"';
  QQ := OpenSQL(StSQL, False);

  If Not QQ.eof then
  begin
    TobFamMat.SelectDB('', QQ);
    //
    If Assigned(TobfamMat) then
    begin
      Panel11.Visible  := True;
      //
      LibFamille.Caption    := TobFamMat.GetString('BFM_LIBELLE');
      //
      LibAlpha1.caption     := TobFamMat.GetString('BFM_LIBALPHA1');
      if LibAlpha1.Caption = '' then ValAlpha1.Visible := False else ValAlpha1.Visible := True;

      LibAlpha2.caption     := TobFamMat.GetString('BFM_LIBALPHA2');
      if LibAlpha2.Caption = '' then ValAlpha2.Visible := False else ValAlpha2.Visible := True;

      LibAlpha3.caption     := TobFamMat.GetString('BFM_LIBALPHA3');
      if LibAlpha3.Caption = '' then ValAlpha3.Visible := False else ValAlpha3.Visible := True;

      LibAlpha4.caption     := TobFamMat.GetString('BFM_LIBALPHA4');
      if LibAlpha4.Caption = '' then ValAlpha4.Visible := False else ValAlpha4.Visible := True;

      LibAlpha5.caption     := TobFamMat.GetString('BFM_LIBALPHA5');
      if LibAlpha5.Caption = '' then ValAlpha5.Visible := False else ValAlpha5.Visible := True;

      LibAlpha6.caption     := TobFamMat.GetString('BFM_LIBALPHA6');
      if LibAlpha6.Caption = '' then ValAlpha6.Visible := False else ValAlpha6.Visible := True;

      LibAlpha7.caption     := TobFamMat.GetString('BFM_LIBALPHA7');
      if LibAlpha7.Caption = '' then ValAlpha7.Visible := False else ValAlpha7.Visible := True;

      LibAlpha8.caption     := TobFamMat.GetString('BFM_LIBALPHA8');
      if LibAlpha8.Caption = '' then ValAlpha8.Visible := False else ValAlpha8.Visible := True;

      LibAlpha9.caption     := TobFamMat.GetString('BFM_LIBALPHA9');
      if LibAlpha9.Caption = '' then ValAlpha9.Visible := False else ValAlpha9.Visible := True;

      LibAlpha10.caption    := TobFamMat.GetString('BFM_LIBALPHA10');
      if LibAlpha10.Caption= '' then ValAlpha10.Visible:= False else ValAlpha10.Visible:= True;
      //
      if (LibAlpha1.Caption = '') and (LibAlpha6.Caption = '')  and
         (LibAlpha2.Caption = '') and (LibAlpha7.Caption = '')  and
         (LibAlpha3.Caption = '') and (LibAlpha8.Caption = '')  and
         (LibAlpha4.Caption = '') and (LibAlpha9.Caption = '')  and
         (LibAlpha5.Caption = '') and (LibAlpha10.Caption = '') then
        Panel6.Visible := false
      else
        Panel6.Visible := True;
      //
      Libdate1.caption      := TobFamMat.GetString('BFM_LIBDATE1');
      If Libdate1.Caption = '' then ValDate1.Visible := False else ValDate1.Visible := True;
      Libdate2.caption      := TobFamMat.GetString('BFM_LIBDATE2');
      If Libdate2.Caption = '' then ValDate2.Visible := False else ValDate2.Visible := True;
      Libdate3.caption      := TobFamMat.GetString('BFM_LIBDATE3');
      If Libdate3.Caption = '' then ValDate3.Visible := False else ValDate3.Visible := True;
      Libdate4.caption      := TobFamMat.GetString('BFM_LIBDATE4');
      If Libdate4.Caption = '' then ValDate4.Visible := False else ValDate4.Visible := True;
      //
      if (Libdate1.Caption = '') And
         (Libdate2.Caption = '') And
         (Libdate3.Caption = '') And
         (Libdate4.Caption = '') then
        Panel10.Visible := False
      else
        Panel10.Visible := True;
      //
      ValBoolean1.caption   := TobFamMat.GetString('BFM_LIBBOOLEAN1');
      If ValBoolean1.Caption = '' then ValBoolean1.Visible := False else ValBoolean1.Visible := True;
      ValBoolean2.caption   := TobFamMat.GetString('BFM_LIBBOOLEAN2');
      If ValBoolean2.Caption = '' then ValBoolean2.Visible := False else ValBoolean2.Visible := True;
      ValBoolean3.caption   := TobFamMat.GetString('BFM_LIBBOOLEAN3');
      If ValBoolean3.Caption = '' then ValBoolean3.Visible := False else ValBoolean3.Visible := True;
      ValBoolean4.caption   := TobFamMat.GetString('BFM_LIBBOOLEAN4');
      If ValBoolean4.Caption = '' then ValBoolean4.Visible := False else ValBoolean4.Visible := True;
      //
      if (ValBoolean1.Caption = '') And
         (ValBoolean2.Caption = '') And
         (ValBoolean3.Caption = '') And
         (ValBoolean4.Caption = '') then
        Panel11.Visible := False
      else
        Panel11.Visible := True;
      //
    end;
  end;

  Ferme(QQ);

end;

Procedure TOF_BTMATERIEL.ChargeInfoEvenement;
begin

  TobEvent.ClearDetail;

  StSQl := 'SELECT * FROM BTEVENTMAT_VUE WHERE BEM_CODEMATERIEL = "' + CodeMateriel.text + '" ORDER BY BEM_DATEDEB DESC';
  QQ := OpenSQL(StSQL, False);

  If Not QQ.eof then
  begin
    Tobevent.LoadDetailDB('BTEVENTMAT','','',QQ, False);
  end;

  Ferme(QQ);

  if Tobevent.Detail.count >0 then
  begin
    PEvent.tabvisible := True;
    BtEvent.visible   := False;
    GGrilleEvent.ChargementGrille;
  end
  else
  begin
    PEvent.tabvisible := False;
    BtEvent.visible   := True;
  end;

end;

Procedure TOF_BTMATERIEL.ChargeInfoAffect;
begin

  TobAffect.ClearDetail;

  StSQl := 'SELECT * FROM BTAFFECTATION_VUE WHERE BFF_CODEMATERIEL = "' + CodeMateriel.text + '" ORDER BY BFF_DATEFIN DESC';

  QQ := OpenSQL(StSQL, False);

  If Not QQ.eof then
  begin
    TobAffect.LoadDetailDB('BTAFFECTATION','','',QQ, False);
  end;

  if TobAffect.Detail.count >0 then
  begin
    Paffect.tabvisible  := True;
    BtAffect.visible    := False;
    GGrilleAffect.ChargementGrille;
  end
  else
  begin
    Paffect.tabvisible := false;
    BtAffect.Visible := true;
  end;

  Ferme(QQ);

end;

{***********UNITE*************************************************
Auteur  ...... : FV1
Créé le ...... : 03/06/2015
Modifié le ... :   /  /
Description .. : Gestion des évènements sur les boutons
Mots clefs ... : Evenements, boutons
*****************************************************************}
Procedure TOF_BTMATERIEL.BTCoutOnclick(Sender : TObject);
Var TOBCout : Tob;
    YearD   : Word;
    YearF   : Word;
    DayD    : Word;
    MonthD  : Word;
    DayF    : Word;
    MonthF  : Word;
begin


  //contrôle que le bornage de date ne dépasse pas les un an
  DecodeDate(StrtoDate(Borne1.text), YearD, MonthD, DayD);
  DecodeDate(StrtoDate(Borne2.text), YearF, MonthF, DayF);
  if YearD <> YearF then
  begin
    PGIError(TexteMessage[16], 'Matériel');
    Borne2.SetFocus;
    exit;
  end;

  TobCout := tob.create('COUT', Nil, -1);

  TOBCout.AddChampSupValeur('RESSOURCE',      CodeRessource.text);
  TOBCout.AddChampSupValeur('STANDCALEN',     TobRessource.GetString('ARS_STANDCALEN'));
  TOBCout.AddChampSupValeur('CALENSPECIF',    TobRessource.GetString('ARS_CALENSPECIF'));
  TOBCout.AddChampSupValeur('BORNE1',         StrToDateTime(Borne1.text));
  TOBCout.AddChampSupValeur('BORNE2',         StrToDateTime(Borne2.text));
  TOBCout.AddChampSupValeur('CODEMATERIEL',   CodeMateriel.text);
  TOBCout.AddChampSupValeur('AMORTISSEMENT',  ValAmort.text);


  //On Lance isi les calculs de coûts en fonction des deux dates rentrées
  CalculCout(TobCout);

  //Chargement de TobCout dans les Zones Ecran
  PeriodeHeure.text     := FormatFloat(FF, TobCout.GetValue('PeriodeHeure'));
  PeriodeJour.text      := FormatFloat(FF, TobCout.GetValue('PeriodeJour'));
  UtilHeure.text        := FormatFloat(FF, TobCout.GetValue('UtilHeure'));
  Utiljour.text         := FormatFloat(FF, TobCout.GetValue('Utiljour'));
  UtilPercent.text      := FormatFloat(FF, TobCout.GetValue('UtilPercent'));
  ImmoHeure.text        := FormatFloat(FF, TobCout.GetValue('ImmoHeure'));
  ImmoJour.text         := FormatFloat(FF, TobCout.GetValue('ImmoJour'));
  ImmoPercent.text      := FormatFloat(FF, TobCout.GetValue('ImmoPercent'));
  EntretienHeure.text   := FormatFloat(FF, TobCout.GetValue('EntretienHeure'));
  entretienJour.text    := FormatFloat(FF, TobCout.GetValue('entretienJour'));
  entretienPercent.text := FormatFloat(FF, TobCout.GetValue('entretienPercent'));
  NotUseHeure.text      := FormatFloat(FF, TobCout.GetValue('NotUseHeure'));
  NotUseJour.text       := FormatFloat(FF, TobCout.GetValue('NotUseJour'));
  NotUsePercent.text    := FormatFloat(FF, TobCout.GetValue('NotUsePercent'));
  Amortissement.Text    := FormatFloat(FF, TobCout.GetValue('Amortissement'));
  CoutEntretien.text    := FormatFloat(FF, TobCout.GetValue('CoutEntretien'));
  TotalCoutPeriode.text := FormatFloat(FF, TobCout.GetValue('TotalCoutPeriode'));
  CtRevientJour.text    := FormatFloat(FF, TobCout.GetValue('CtRevientJour'));
  CtRevientHeure.text   := FormatFloat(FF, TobCout.GetValue('CtRevientHeure'));

  FreeAndNil(TobCout);

  if PageCTRL.ActivePageIndex = 1 then CoutRevient.SetFocus;

end;

procedure TOF_BTMATERIEL.OngletOnenter(Sender : Tobject);
begin

  if (PageCtrl.ActivePageIndex = 3) Or  (PageCtrl.ActivePageIndex = 4) then
  begin
    BtParamList.Visible := True;
    BtDelete1.Visible   := True;
    BtNouveau.Visible   := True;
    
    if (PageCtrl.ActivePageIndex = 3) then
    begin
      Btdelete1.Hint      := 'Suppression Evénement';
      BtNouveau.Hint      := 'Nouvel Evénement';
    end
    else
    begin
      Btdelete1.Hint      := 'Suppression affectation';
      BtNouveau.Hint      := 'Nouvelle affectation';
    end;
  end
  else
  begin
    BtParamList.Visible := false;
    BtDelete1.Visible   := false;
    BtNouveau.Visible   := False
  end;

    
end;

procedure TOF_BTMATERIEL.BtParamListOnClick(Sender : Tobject);
begin

  if PageCtrl.ActivePage.Name = 'PAFFECTATION' then
  begin
    ParamListe(GGrilleAffect.NomListe, nil, nil, '');
    //on doit recharger les grilles !!!
    FreeAndNil(GGrilleAffect);
    InitGrilleAffect;
    GGrilleAffect.ChargementGrille;
  end
  else if PageCtrl.ActivePage.Name = 'PEVENEMENT'  then
  begin
    ParamListe(GGrilleEvent.NomListe, nil, nil, '');
    //on doit recharger les grilles !!!
    FreeAndNil(GGrilleEvent);
    InitGrilleEvent;
    GGrilleEvent.ChargementGrille;
  end;  

end;

procedure TOF_BTMATERIEL.BTEventOnClick(Sender: Tobject);
begin

  AGLLanceFiche('BTP','BTEVENTMAT','','','MATERIEL='+CodeMateriel.text+';ACTION=CREATION');

  ChargeInfoEvenement;

end;

procedure TOF_BTMATERIEL.BTAffectOnClick(Sender: TObject);
begin

  AGLLanceFiche('BTP','BTAFFECTPARC','','','MATERIEL='+CodeMateriel.text+';ACTION=CREATION');

  ChargeInfoAffect;

end;


procedure TOF_BTMATERIEL.RessourceOnExit(Sender: Tobject);
begin

  ChargeInfoRessource;

end;

Procedure TOF_BTMATERIEL.MaterielOnExit(Sender : tobject);
begin

  StSQL := 'SELECT * FROM BTMATERIEL WHERE BMA_CODEMATERIEL="' + CodeMateriel.Text + '"';

  //contrôle si le code matériel n'est pas déjà crée
  if ExisteSQL(StSQL) then
  begin
    PGIError(TexteMessage[5], 'Matériel');
    CodeMateriel.Text := '';
    //CodeMateriel.SetFocus;
  end;    

end;

procedure TOF_BTMATERIEL.RessourceOnElipsisClick(Sender: Tobject);
Var Title   : String;
    StWhere : String;
    StChamp : String;
begin

  title := 'Recherche Ressource Matériel';

  //stWhere := 'AND ARS_TYPERESSOURCE IN ("OUT","MAT","LOC","AUT")';
  if (VH_GC.AFRechResAv) then
    stWhere := 'TYPERESSOURCE : OUT,MAT,LOC,AUT'
  else
    stWhere := 'AND ARS_TYPERESSOURCE IN ("OUT","MAT","LOC","AUT")';

  StChamp  := CodeRessource.Text;

  DispatchRecherche(CodeRessource, 3, StWhere,'ARS_RESSOURCE=' + Trim(CodeRessource.Text), '');

  ChargeInfoRessource;

end;

procedure TOF_BTMATERIEL.FamMatOnExit(Sender: Tobject);
begin

  if CodeFamille.Text <> SCodeFamille then
  begin
    ScodeFamille := CodeFamille.Text;
    RazZoneFamMat;
    ChargeInfoFamMat;
  end;

end;

procedure TOF_BTMATERIEL.FamMatOnElipsisClick(Sender: Tobject);
Var Title : string;
begin

  title := 'Recherche Famille Matériel';

  if not LookupList(CodeFamille,Title,'BTFAMILLEMAT','BFM_CODEFAMILLE','BFM_LIBELLE','','BFM_LIBELLE',True,60) then Exit;

  if CodeFamille.Text <> SCodeFamille then
  begin
    ScodeFamille := CodeFamille.Text;
    RazZoneFamMat;
    ChargeInfoFamMat;
  end;

end;


procedure TOF_BTMATERIEL.ArticleOnExit(Sender: Tobject);
begin

  ChargeInfoArticle;

end;

procedure TOF_BTMATERIEL.ArticleOnElipsisClick(Sender: Tobject);
Var StWhere : String;
    StChamps: String;
    title   : string;
begin

  title := 'Recherche Prestation Associée';

  stWhere := 'GA_TYPEARTICLE="PRE"';

  stchamps := 'TYPERESSOURCE=MAT,LOC,OUT';

  if stWhere <> '' then StChamps := StChamps + ';' + stWhere;

  CodeArticle.text := AGLLanceFiche('BTP', 'BTPREST_RECH', 'GA_CODEARTICLE='+Trim(Copy(CodeArticle.text, 1, 18)), '', stchamps);

  if CodeArticle.text <> '' then ChargeInfoArticle;

end;

Procedure TOF_BTMATERIEL.AffichageInitEcran(OkAff : Boolean);
begin

  //la zone contenant les zones alpha est rendue invisible au chargement
  Panel6.Visible  := OkAff;

  //la zone contenant les zones date est rendue invisible au chargement
  Panel10.Visible  := OkAff;

  //la zone contenant les zones choix est rendue invisible au chargement
  Panel11.Visible  := OkAff;

  //les onglets évènements et action/Affectations ne seront affichés que s'il y a quelque chose
  //à charger dedans.
  PageCtrl.Pages[2].TabVisible := OkAff;
  PageCtrl.pages[3].TabVisible := OkAff;
  PageCtrl.Pages[4].TabVisible := OkAff;

  //positionnement sur l'onglet 1 - Généralités
  PageCTRL.ActivePageIndex := 0;

end;

Procedure TOF_BTMATERIEL.InitGrilleAffect;
begin

  //Une recherche de la grille au niveau de la table des liste serait bien venu !!!
  GGrilleAffect           := TGestionGS.Create;

  GGrilleAffect.Ecran     := TFVierge(Ecran);
  GGrilleAffect.GS        := GrilleAffect;
  GGrilleAffect.TOBG      := TobAffect;

  GGrilleAffect.NomListe  := 'BTAFFECTATION';

  GGrilleAffect.ChargeInfoListe;

  if GGrilleAffect.NomListe <> '' then
    GGrilleAffect.DessineGrille
  else
    PageCtrl.pages[4].TabVisible := False;

end;

Procedure TOF_BTMATERIEL.InitGrilleEvent;
begin

  //Une recherche de la grille au niveau de la table des liste serait bien venu !!!
  GGrilleEvent           := TGestionGS.Create;

  GGrilleEvent.Ecran     := TFVierge(Ecran);
  GGrilleEvent.GS        := GrilleEvent;
  GGrilleEvent.TOBG      := TobEvent;

  GGrilleEvent.NomListe  := 'BTEVENTMAT';

  GGrilleEvent.ChargeInfoListe;

  if GGrilleEvent.NomListe <> '' then
    GGrilleEvent.DessineGrille
  else
    PageCtrl.pages[3].TabVisible := False;

end;

procedure TOF_BTMATERIEL.DblClickGrille(Sender : TObject);
Var Code      : Integer;
    Argument  : string;
begin

  if PageCTRL.ActivePageIndex = 3 then
  begin
    code := GGrilleEvent.TOBG.detail[GrilleEvent.row-1].GetValue('BEM_IDEVENTMAT');
    if code = 0 then
    begin
      PGIError(TexteMessage[8], ' Matériel');
      Exit;
    end
    else
    begin
      Argument := 'IDEVENEMENT=' + InttoStr(Code) + ';ACTION=MODIFICATION';
      AGLLanceFiche('BTP','BTEVENTMAT','','',Argument);
    end;
    ChargeInfoEvenement;
  end
  else if PageCTRL.ActivePageIndex = 4 then
  begin
    code := GGrilleAffect.TOBG.detail[GrilleAffect.row-1].GetValue('BFF_IDAFFECTATION');
    if code = 0 then
    begin
      PGIError(TexteMessage[8], ' Matériel');
      Exit;
    end
    else
    begin
      Argument := 'IDAFFECTATION=' + IntToStr(Code) + ';ACTION=MODIFICATION';
      AGLLanceFiche('BTP','BTAFFECTPARC','','',Argument);
    end;
    ChargeInfoAffect;
  end;

end;
Procedure TOF_BTMATERIEL.OnDupliqueFiche(Sender : TObject);
begin

  if action = TaCreat then Exit;

  Action                := TaCreat;

  BtDelete.Visible      := False;
  Btimprimer.Visible    := False;
  BTDuplication.Visible := False;
  //
  DateCreation.Text     := DateToStr(Now);
  UserCreate.text       := V_PGI.User;
  //
  CodeMateriel.Text     := '';
  LibMateriel.Text      := '';
  CodeMateriel.Enabled  := True;
  //
  CodeRessource.Text    := '';
  LibRessource.Caption  := '';
  //
  PeriodeHeure.text     := FormatFloat(FF, 0);
  PeriodeJour.text      := FormatFloat(FF, 0);
  UtilHeure.text        := FormatFloat(FF, 0);
  Utiljour.text         := FormatFloat(FF, 0);
  ImmoHeure.text        := FormatFloat(FF, 0);
  ImmoJour.text         := FormatFloat(FF, 0);
  EntretienHeure.text   := FormatFloat(FF, 0);
  EntretienJour.text    := FormatFloat(FF, 0);
  NotUseHeure.Text      := FormatFloat(FF, 0);
  NotUseJour.Text       := FormatFloat(FF, 0);
  Amortissement.text    := FormatFloat(FF, 0);
  CoutEntretien.text    := FormatFloat(FF, 0);
  TotalCoutPeriode.text := FormatFloat(FF, 0);
  CtRevientHeure.Text   := FormatFloat(FF, 0);
  CtRevientJour.Text    := FormatFloat(FF, 0);

  //chargement des évèenements
  ChargeInfoEvenement;
  //
  ChargeInfoAffect;
  //

end;

Procedure TOF_BTMATERIEL.OnImprimeFiche(Sender : TObject);
Var TOBL : Tob;
begin

  TobL := TobEdition.detail[0];
  //
  TOBL.AddChampSupValeur('BORNE1',        Borne1.text);
  TOBL.AddChampSupValeur('BORNE2',        Borne2.Text);
  TOBL.AddChampSupValeur('PERIODEH',      PeriodeHeure.text);
  TOBL.AddChampSupValeur('PERIODEJ',      PeriodeJour.text);
  TOBL.AddChampSupValeur('UTILISATIONH',  UtilHeure.text);
  TOBL.AddChampSupValeur('UTILISATIONJ',  Utiljour.text);
  TOBL.AddChampSupValeur('IMMOH',         ImmoHeure.text);
  TOBL.AddChampSupValeur('IMMOJ',         ImmoJour.text);
  TOBL.AddChampSupValeur('ENTRETIENH',    EntretienHeure.text);
  TOBL.AddChampSupValeur('ENTETIENJ',     EntretienJour.text);
  TOBL.AddChampSupValeur('NOUSEH',        NotUseHeure.Text);
  TOBL.AddChampSupValeur('NOUSEJ',        NotUseJour.Text);
  TOBL.AddChampSupValeur('AMORTISSEMENT', Amortissement.text);
  TOBL.AddChampSupValeur('COUTE',         CoutEntretien.text);
  TOBL.AddChampSupValeur('TOTALCOUT',     TotalCoutPeriode.text);
  TOBL.AddChampSupValeur('COUTH',         CtRevientHeure.Text);
  TOBL.AddChampSupValeur('COUTJ',         CtRevientJour.Text);
  //
  if OptionEdition.LanceImpression('', TobEdition) < 0 then V_PGI.IoError:=oeUnknown;

end;


procedure TOF_BTMATERIEL.OnNewRecord(Sender : TObject);
Var Argument  : string;
begin

  Argument := 'ACTION=CREATION';

  if PageCTRL.ActivePageIndex = 3 then
    BTEventOnClick(self)
  else if PageCTRL.ActivePageIndex = 4 then
    BtAffectOnclick(Self);

end;

procedure TOF_BTMATERIEL.OnDelEnregGrille(Sender : Tobject);
Var Code : Integer;
begin

  if PageCTRL.ActivePageIndex = 3 then
  begin
    code := GGrilleEvent.TOBG.detail[GrilleEvent.row-1].GetValue('BEM_IDEVENTMAT');
    if code = 0 then
    begin
      PGIError(TexteMessage[8], ' Matériel');
      Exit;
    end
    else
    begin
      if PGIAsk('Confirmez-vous la suppression de cet Evénement : ' + IntToStr(Code) + ' ?', 'Evénement Matériel') = Mryes then
      begin
        //contrôle si Matériel présent sur évènements
        StSQl := 'DELETE BTEVENTMAT WHERE BEM_IDEVENTMAT=' + IntToStr(code) + ' AND BEM_CODEMATERIEL = "' + CodeMateriel.text + '" ';
        if ExecuteSQL(StSQL)=0 then PGIError(TexteMessage[6], ' Matériel');
      end;
    end;
    ChargeInfoEvenement;
  end
  else if PageCTRL.ActivePageIndex = 4 then
  begin
    code := GGrilleAffect.TOBG.detail[GrilleEvent.row-1].GetValue('BFF_IDAFFECTATION');
    if code = 0 then
    begin
      PGIError(TexteMessage[8], ' Matériel');
      Exit;
    end
    else
    begin
      if PGIAsk('Confirmez-vous la suppression de cette Affectation : ' + IntToStr(Code) + ' ?', 'Affectation Matériel') = Mryes then
      begin
        //contrôle si Matériel présent sur évènements
        StSQl := 'DELETE BTAFFECTATION WHERE BFF_IDAFFECTATION=' + IntToStr(code) + ' AND BFF_CODEMATERIEL = "' + CodeMateriel.text + '" ';
        if ExecuteSQL(StSQL)=0 then PGIError(TexteMessage[7], ' Matériel');
      end;
    end;
    ChargeInfoAffect;
  end;

end;

procedure Tof_BTMATERIEL.ChargeTobEdition;
Var TOBL : TOB;
begin

  TOBL := TobEdition.detail[0];

  TOBL.PutValue('LIBFAMILLE',             LibFamille.Caption);

  TOBL.PutValue('LIBRESSOURCE',           LibRessource.Caption);

  TOBL.PutValue('LIBPRESTATION',          LibArticle.caption);
  //
  TOBL.PutValue('LIBALPHA1',              LibAlpha1.caption);
  TOBL.PutValue('LIBALPHA2',              LibAlpha2.Caption);
  TOBL.PutValue('LIBALPHA3',              LibAlpha3.Caption);
  TOBL.PutValue('LIBALPHA4',              LibAlpha4.Caption);
  TOBL.PutValue('LIBALPHA5',              LibAlpha5.Caption);
  TOBL.PutValue('LIBALPHA6',              LibAlpha6.Caption);
  TOBL.PutValue('LIBALPHA7',              LibAlpha7.Caption);
  TOBL.PutValue('LIBALPHA8',              LibAlpha8.Caption);
  TOBL.PutValue('LIBALPHA9',              LibAlpha9.Caption);
  TOBL.PutValue('LIBALPHA10',             LibAlpha10.Caption);
  //
  TOBL.PutValue('LIBDATE1',               LibDate1.Caption);
  TOBL.PutValue('LIBDATE2',               LibDate2.Caption);
  TOBL.PutValue('LIBDATE3',               LibDate3.Caption);
  TOBL.PutValue('LIBDATE4',               LibDate4.Caption);
  //
  TOBL.PutValue('LIBBOOLEAN1',            ValBoolean1.Caption);
  TOBL.PutValue('LIBBOOLEAN2',            ValBoolean2.Caption);
  TOBL.PutValue('LIBBOOLEAN3',            ValBoolean3.Caption);
  TOBL.PutValue('LIBBOOLEAN4',            ValBoolean4.Caption);
  //
  TOBL.AddChampSupValeur('BLOCNOTE',      BlocNote.text);
  //
end;

Procedure TOF_BTMATERIEL.InitTobEdition;
var TobL : TOB;
begin

  StSQL := 'SELECT * FROM BTMATERIEL WHERE BMA_CODEMATERIEL = "' + CodeMateriel.text + '"';
  QQ := OpenSQL(StSQL, False);

  TobEdition.LoadDetailDB('BTMATERIEL','','',QQ, False);

  TobL := TobEdition.detail[0];

  TobL.AddChampSupValeur('LIBFAMILLE',        '');
  TobL.AddChampSupValeur('LIBRESSOURCE',      '');
  TobL.AddChampSupValeur('LIBPRESTATION',     '');
  //
  TobL.AddChampSupValeur('LIBALPHA1',         '');
  TobL.AddChampSupValeur('LIBALPHA2',         '');
  TobL.AddChampSupValeur('LIBALPHA3',         '');
  TobL.AddChampSupValeur('LIBALPHA4',         '');
  TobL.AddChampSupValeur('LIBALPHA5',         '');
  TobL.AddChampSupValeur('LIBALPHA6',         '');
  TobL.AddChampSupValeur('LIBALPHA7',         '');
  TobL.AddChampSupValeur('LIBALPHA8',         '');
  TobL.AddChampSupValeur('LIBALPHA9',         '');
  TobL.AddChampSupValeur('LIBALPHA10',        '');
  //
  TobL.AddChampSupValeur('LIBDATE1',          '');
  TobL.AddChampSupValeur('LIBDATE2',          '');
  TobL.AddChampSupValeur('LIBDATE3',          '');
  TobL.AddChampSupValeur('LIBDATE4',          '');
  //
  TobL.AddChampSupValeur('LIBBOOLEAN1',       '');
  TobL.AddChampSupValeur('LIBBOOLEAN2',       '');
  TobL.AddChampSupValeur('LIBBOOLEAN3',       '');
  TobL.AddChampSupValeur('LIBBOOLEAN4',       '');
  //
  TobL.AddChampSupValeur('BORNE1',            '');
  TobL.AddChampSupValeur('BORNE2',            '');
  TobL.AddChampSupValeur('PERIODEH',          '');
  TobL.AddChampSupValeur('PERIODEJ',          '');
  TobL.AddChampSupValeur('UTILISATIONH',      '');
  TobL.AddChampSupValeur('UTILISATIONJ',      '');
  TobL.AddChampSupValeur('IMMOH',             '');
  TobL.AddChampSupValeur('IMMOJ',             '');
  TobL.AddChampSupValeur('ENTRETIENH',        '');
  TobL.AddChampSupValeur('ENTETIENJ',         '');
  TobL.AddChampSupValeur('NOUSEH',            '');
  TobL.AddChampSupValeur('NOUSEJ',            '');
  TobL.AddChampSupValeur('AMORTISSEMENT',     '');
  TobL.AddChampSupValeur('COUTE',             '');
  TobL.AddChampSupValeur('TOTALCOUT',         '');
  TobL.AddChampSupValeur('COUTH',             '');
  TobL.AddChampSupValeur('COUTJ',             '');
  //
  TobL.AddChampSupValeur('BLOCNOTE',          '');
  //
end;

Initialization
  registerclasses ( [ TOF_BTMATERIEL ] ) ;
end.

