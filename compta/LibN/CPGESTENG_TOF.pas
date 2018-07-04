{***********UNITE*************************************************
Auteur  ...... : Thong Hor LIM
Créé le ...... : 17/03/2006
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : CPGESTENG ()
Mots clefs ... : TOF;CPGESTENG
*****************************************************************}
Unit CPGESTENG_TOF ;

Interface

Uses StdCtrls,
     Controls,
     Classes,                                
     Graphics,
     forms,
     Windows,         // VK_
     vierge,
     Paramsoc,
     AGLInit,         // TheData
     HEnt1,
     Ent1,            // VH^, TFichierBase
     utilPGI,         // OpenSelect
     Ed_Tools,        // Pour le videListe
     Htb97,           // TToolBarButton97, TToolWindow97
     Grids,
     Hdb,
     HQry,
     HPanel,        //THpanel
     Dialogs,

{$IFDEF eAGLCLIENT}
     MenuOLX,
{$ELSE}
     MenuOLG,
{$ENDIF eAGLCLIENT}

{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul,
     Fe_Main,     // AGLLanceFiche
     printdbg,   {printdbgrid}
{$else}
     eMul,
     MaineAGL,
     utileagl,   {printdbgrid}
{$ENDIF}
     sysutils,
     ComCtrls,
     Menus,           // TMenuItem, PopUpMenu
     LookUp,          // LookUpList
     HCtrls,
     HMsgBox,
     UTof,
     SAISUTIL,		    // Pour RMVT
     UObjFiltres,         //gestion des filtres
     ULibPieceCompta,     // TPieceCompta    {FP 02/05/2006}
     CPSAISIEPIECE_TOF,   // RmvtToPiece     {FP 02/05/2006}
     CPOBJENGAGE,
     UTob;

Type

  TOF_CPGESTENG = Class (TOF)
//    FEcran    : TFMul ;

    // Eléments interface
    PlanCompta         :THValCombobox ;
    LeCompte           :THEdit;
    SoldeEng           :THEdit;
    SoldeCpta          :THEdit;
    CLIBELLE           :THLabel;
    BGUP               :TToolBarButton97;
    BGDOWN             :TToolBarButton97;

    //onglet engagement
    E_JournalEng       :THMultiValComboBox;
    CEN_StatutEng      :THMultiValComboBox;
    E_DATEREFEXTERNE   :THEdit;
    E_DATEREFEXTERNE_  :THEdit;
    E_DATECREATION     :THEdit;
    E_DATECREATION_    :THEdit;
    E_DATEECHEANCE     :THEdit;
    E_DATEECHEANCE_    :THEdit;
    E_REFINTERNE       :THEdit;
    E_REFINTERNE_      :THEdit;
    //onglet Cpta
    E_Journal          :THMultiValComboBox ;
    E_EXERCICE         :THValComboBox ;
    E_NATUREPIECE      :THMultiValComboBox ;
    E_DATECOMPTABLE    :THEdit;
    E_DATECOMPTABLE_   :THEdit;
    E_ETABLISSEMENT    :THMultiValComboBox ;   // FQ 20517 : SBO 08/08/2007 : Ajout critère Etablissement

    BChercher          :TButton ;
    BValider           :TButton ;
    BRechercher        :TToolBarButton97;
    BImprimer          :TToolBarButton97;
    BVoirEngLies       :TToolBarButton97;
    BAgrandir          :TToolBarButton97;
    Z_C1, Z_C2, Z_C3   :THValCombobox ;


    ONGLETSHAUT        : TPageControl;
    {
    ONGLETENGAGE,
    ONGLETCOMPTA,
    ONGLETTABLIBRES,
    ONGLETAVANCEs      : TTabSheet;
    }
    ObjFiltre          : TObjFiltre;
    FFindDialog        : TFindDialog; // Fenêtre de recherche sur les listes

    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;

  private
    FFiltreEcrRapproAuto     : string;
    FFiltreEngRapproAuto     : string;
    FAgrandirEcran           : boolean;
    FFindFirst               : Boolean;
    FSelectedGrid            : THGrid ;

    LMsg                     : THMsgBox ;
    PCpta, PEnga , PGene     : THPanel;

    GEnga,
    GCpta                    : THGrid ;
    FTobComptes              : TOB;         // Tob qui contient les comptes Tiers ou Généraux
    PopupGestionEng          : TPopUpMenu;  // pour menu de gestion enga.
    PopupGEnga,                             // pour factures liées
    PopupGCpta               : TPopUpMenu;  // pour engagements liés
    FBoConfidentielCpt       : Boolean; // Gestion Confidentiel du cpte pour blocage des fonctions

    TobEcrEng,
    TobEcrCpta               : TOB;

    IndMess                  : Integer;
    LstPiecesEngOuFac        : TList ;  //sert pour l'affichage.
    LstPiecesEng             : TList ;  //sert pour le dé-rapprochement.

    TotSoldeEng,
    TotSoldeCpta             : Double;


    EngagementHT             : boolean; // pour test

    procedure OnKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState) ;
    procedure OnFormResize(Sender:  TObject) ;
    procedure GMouseUp(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer) ;

    procedure InitMsg ;
    function AfficheMsg(num : integer;Av,Ap : string ) : Word ;

    procedure OnClickBChercher                 (Sender: TObject);
    procedure OnClickBImprimer                 (Sender: TObject);
    procedure OnClickBVoirEngLies              (Sender: TObject);
    procedure OnRechercherClick                (Sender: TObject);     // Recherche sur les liste
    procedure OnFindAFindDialog                (Sender: TObject);    // Recherche sur les liste
    procedure OnGrilleDblClick                 (Sender: TObject);
    procedure PlanComptaChange                 (Sender: TObject);

    procedure E_EXERCICEChange(Sender: TObject);
    procedure OnExitLeCompte( Sender : TObject );
    function ChargeLeCompteGene(Force : boolean): Boolean;
    function ChargeLeCompteTier(Force : boolean): Boolean;
    function ChargeLeCompte(Force : boolean): Boolean;
    procedure GeneSuivant (vBoSuiv : Boolean; vBoREsteSurZone : Boolean);
    procedure TiersSuivant (vBoSuiv : Boolean; vBoREsteSurZone : Boolean);
    procedure LeCompteSuivant (vBoSuiv : Boolean; vBoREsteSurZone : Boolean);
    procedure OnElipsisClickLeCompte( Sender : TObject );
    procedure OnClickBGUP( Sender : TObject );
    procedure OnClickBGDOWN( Sender : TObject );
    procedure ChangeAffichage(Plus: boolean) ;

    {b FP 02/05/2006}
    function  OnBeforeSuppLigne(Tof: TOF_CPSAISIEPIECE; Piece: TPieceCompta; NumLigne: Integer): Boolean;
    function  OnBeforeValidePieceComptaSimul(Tof: TOF_CPSAISIEPIECE; Ecran: TForm; Piece: TPieceCompta): Boolean;
    {e FP 02/05/2006}

    procedure InitPopUP(vActivation: Boolean; vNumPopUp : integer =  0);
    procedure OnPopupPopUpGestionEng(Sender : TObject);
    procedure OnPopupPOPUPGEnga(Sender : TObject);
    procedure OnPopupPOPUPGCpta(Sender : TObject);
    procedure OnAgrandirClick(Sender: TObject);
    procedure OnValiderRapproAuto(Sender: TObject);
    procedure OnChercherRapproAuto(Sender: TObject);
    procedure OnGridEnter(Sender: TObject);

//    procedure InitCriteres ;
    procedure InitGrid(GS: THGrid) ;
    procedure AfficheSoldes;

    function  isEngagementHT: boolean;
    function  IsCompteOk: boolean;
    function  IsCompteGene: boolean;
    function  IsRapproManuelAutorise: boolean;
    function  IsTerminerUnEngagementAutorise: boolean;
    function  IsTransFactAutorise: boolean;    {FP 02/05/2006}
    function  IsExistFacturesLiees : boolean;
    function  IsExistEngagementsLies: boolean;
    function  IsDerapproAutorise: Boolean;

    function  ClauseSelectEngRapproche(Jou, Pie, QualifPiece: string; DateComptable: TDateTime; Exe: string): string;
    function  ClauseSelectFacRapprochee(Jou, Pie, QualifPiece: string; DateComptable: TDateTime; Exe: string): string;
    procedure ChargeLstPiecesEngOuFac(Q1 : TQuery);
    procedure ChargeLstPiecesEngPourDRapp(Q1 : TQuery);

    function  GetStatutEng(pos: integer; LstEcr: TOB): char;
    function  GetSoldeFact(pos: integer; LstEcr: TOB): double;
    function  GetRowSelect(G: THGrid): Integer;
    function  GridRowToTobIndex(G: THGrid; Row: Integer): Integer;         {Retoune l'indice de la TOB à partir de l'indice de la grille}
    function  GetClauseRechEngOuFac(pos: integer; LstEcr: TOB): string;
    procedure MarqueTob(LstEcr: TOB; Marque: Boolean);     {FP 02/05/2006}
    function  RapprochementManuel(EngReste: TPieceEngagement): Boolean;


    function  InitControles: boolean ;

    function  ClauseJoinEng(Fact: boolean): string;
    function  ClauseStatutEng :string;
//    function  ClauseJournaux(JNX: THMultiValComboBox):string;
    function  ClauseOngletCompte :string;
    function  ClauseOngletEngagement:string;
    function  ClauseOngletEcrCompta:string;
    function  ClauseOngletAvance :string;
    function  ClauseOngleTabLibres :string;

    function  GetClauseSelectEcrCompta: string;
    function  GetClauseSelectEcrEngagement: string;

    procedure CalculMtHTEcritCpta;
    procedure FillGridsEngaEtCpta;
    procedure FillTobsEngaEtCpta ;
    procedure FillGrid(G: THGrid; TobEcr: Tob);

    {Fonctions Popup}
    procedure AffListeEngagementsLies(Sender : TObject);

    {Les outils de traitement}
    procedure TransFactClick(Sender : TObject);
    procedure RapManuelClick(Sender : TObject);
    procedure RapAutoClick(Sender : TObject);
    procedure DeRapproClick(Sender : TObject);
    procedure TerminerUnEngClick(Sender : TObject);

    protected

    public
  end ;

procedure CPLanceFiche_CPGESTENG(Argument: String='');

procedure  CEngRemplirCBPlan ( vCombo : THValcomboBox ) ;

Implementation

Uses
  {$IFDEF MODENT1}
  CPTypeCons,
  CPProcGen,
  CPProcMetier,
  ULibExercice,
  {$ENDIF MODENT1}
  SaisComm,
  uLibWindows,      // AfficheDBCR  (testjoker)
  DelVisuE,         // Pour affichage des pieces rapprochées
  uLibEngage,
  ULibEcriture;


const
  //communes pour les 2 grids
  ColLigneTOB      = 0;
  ColDateCpta      = ColLigneTOB+1;
  ColJournal       = ColDateCpta + 1;
  ColPiece         = ColJournal + 1;
  ColRefCde        = ColPiece + 1;
  ColLibelle       = ColRefCde + 1 ;
  ColDebit         = ColLibelle + 1;
  ColCredit        = ColDebit +1 ;
  //Eng
  ColMtEngInit     = ColCredit + 1;     // =7
  ColStatutEng     = ColMtEngInit + 1;  // =8
  //Cpta
  ColMtHTCpta      = ColCredit + 1;     // =7
  ColSoldeCpta     = ColMtHTCpta + 1;   // =8

procedure CPLanceFiche_CPGESTENG(Argument: String='');
begin
  AGLLanceFiche('CP', 'CPGESTENG_VIE', '', '', Argument) ;
end;

procedure CEngRemplirCBPlan ( vCombo : THValcomboBox ) ;
begin
  vCombo.Items.Clear ;
  vCombo.Values.Clear ;
  if GetParamSocSecur('SO_CPGESTENGAGE', '') = '2' then
    begin
    vCombo.Items.Add('Généraux') ;
    vCombo.Values.Add('G') ;
    vCombo.ItemIndex := 0 ;
    vCombo.Enabled := False ;
    end
  else
    begin
    vCombo.Items.Add('Fournisseurs') ;
    vCombo.Values.Add('F') ;
    vCombo.Items.Add('Clients') ;
    vCombo.Values.Add('C') ;
    vCombo.Enabled := True ;
    vCombo.ItemIndex := 0 ;
    end ;

end ;

procedure TOF_CPGESTENG.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_CPGESTENG.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_CPGESTENG.OnUpdate ;
begin
  Inherited ;
  
  BChercher.Click;
end ;

procedure TOF_CPGESTENG.OnLoad ;
begin
  Inherited ;
  TFVierge(Ecran).HelpContext := 0 ;
  InitControles;
  InitMsg ;
  InitPopUp(False);

  if (FFiltreEcrRapproAuto <> '') and (FFiltreEngRapproAuto <> '') then
    begin
    BAgrandir.Click;
    BAgrandir.Visible := False;
    BChercher.Click;
    end
  else
    ObjFiltre.Charger;

  if GetControlText('LeCompte') <> '' then
    BChercher.Click;

end ;

procedure TOF_CPGESTENG.OnArgument (S : String ) ;
var
  Composants : TControlFiltre;
  Argument   : String;
begin
  Inherited ;

  EngagementHT := GetParamSocSecur('SO_CPGESTENGAGE', '') = '2' ;


  FFindDialog := TFindDialog.Create(Ecran);
  FFindDialog.OnFind := OnFindAFindDialog;

  //Elements de la  TOF.
  PlanCompta         := THValComboBox(GetControl('PlanCompta', True));

  Z_C1 := THValComboBox(GetControl('Z_C1', True));
  Z_C2 := THValComboBox(GetControl('Z_C2', True));
  Z_C3 := THValComboBox(GetControl('Z_C3', True));

  FTobComptes       := Tob.Create('FTobComptes', nil, -1);
  FSelectedGrid     := nil;

  LeCompte          := THEdit(GetControl('LeCompte'));
  SoldeEng          := THEdit(GetControl('SoldeEng', True))  ;
  SoldeCpta         := THEdit(GetControl('SoldeCpta', True))  ;

  CLIBELLE          := THLabel(GetControl('CLIBELLE'));

  BGUP              := TToolBarButton97(GetControl('BGUP', True));
  BGDOWN            := TToolBarButton97(GetControl('BGDOWN', True));

  PopupGestionEng   := TPopUpMenu(GetControl('POPUPGESTIONENG', True));
  PopupGEnga        := TPopUpMenu(Getcontrol('PopupGEnga', True));
  PopupGCpta        := TPopUpMenu(Getcontrol('PopupGCpta', True));

  E_JOURNAL          := THMultiValComboBox(GetControl('E_JOURNAL', True)) ;
  E_NATUREPIECE      := THMultiValComboBox(GetControl('E_NATUREPIECE', True)) ;
  // FQ 20517 : SBO 08/08/2007 : Ajout critère Etablissement
  E_ETABLISSEMENT    := THMultiValComboBox(GetControl('E_ETABLISSEMENT', True)) ;
  E_EXERCICE         := THValComboBox(GetControl('E_EXERCICE', True)) ;
  E_DATECOMPTABLE    := THEdit(GetControl('E_DATECOMPTABLE', True)) ;
  E_DATECOMPTABLE_   := THEdit(GetControl('E_DATECOMPTABLE_', True)) ;
  E_DATECREATION     := THEdit(GetControl('E_DATECREATION', True))  ;
  E_DATECREATION_    := THEdit(GetControl('E_DATECREATION_', True)) ;
  E_DATEECHEANCE     := THEdit(GetControl('E_DATEECHEANCE', True)) ;
  E_DATEECHEANCE_    := THEdit(GetControl('E_DATEECHEANCE_', True)) ;

  E_JournalEng       := THMultiValComboBox(GetControl('E_JournalEng', True)) ;
  CEN_StatutEng      := THMultiValComboBox(GetControl('CEN_StatutEng', True));

  E_DATEREFEXTERNE   := THEdit(GetControl('E_DATEREFEXTERNE', True)) ;
  E_DATEREFEXTERNE_  := THEdit(GetControl('E_DATEREFEXTERNE_', True)) ;
  E_REFINTERNE       := THEdit(GetControl('E_REFINTERNE', True)) ;
  E_REFINTERNE_      := THEdit(GetControl('E_REFINTERNE_', True));
  BChercher          := TButton(GetControl('BCHERCHER', True));
  BValider           := TButton(GetControl('BVALIDER', True)) ;
  BImprimer          := TToolBarbutton97(GetControl('BIMPRIMER')) ;
  BRechercher        := TToolBarbutton97(GetControl('BRECHERCHER')) ;
  BVoirEngLies       := TToolBarbutton97(GetControl('BVOIRENGLIES')) ;
  BAgrandir          := TToolBarButton97(GetControl('BAGRANDIR1'));


  GEnga              := THGrid(GetControl('GENGAGE', True));
  GCpta              := THGrid(GetControl('GCOMPTA', True));

  GEnga.OnDblClick   := OnGrilleDblClick ;
  GCpta.OnDblClick   := OnGrilleDblClick ;

  ONGLETSHAUT        := TPageControl(GetControl('ONGLETSHAUT', True)) ;
  {
  ONGLETENGAGE       := TTabSheet(GetControl('ONGLETENGAGE', True)) ;
  ONGLETCOMPTA       := TTabSheet(GetControl('ONGLETCOMPTA', True)) ;
  ONGLETAVANCES      := TTabSheet(GetControl('ONGLETAVANCES', True)) ;
  ONGLETTABLIBRES    := TTabSheet(GetControl('ONGLETTABLIBRES', True)) ;
  }

  PCpta := THPanel(GetControl('PCPTA'));
  PEnga := THPanel(GetControl('PEnga'));
  PGene := THPanel(GetControl('PGENE'));


  CEngRemplirCBPlan(PlanCompta) ;
  
  SetControlText('Devise', V_PGI.DevisePivot);

  E_JOURNAL.plus       := ' J_FERME="-" AND J_MODESAISIE="-" AND J_NATUREJAL IN ("ACH", "VTE") ' ;
  E_JOURNALENG.plus    := ' J_FERME="-" AND J_MODESAISIE="-" AND J_NATUREJAL IN ("ACH", "VTE") ' ;
  E_JOURNALENG.Text    := TraduireMemoire('<<Tous>>');
  E_JOURNAL.Text       := TraduireMemoire('<<Tous>>');
  E_ETABLISSEMENT.Text := TraduireMemoire('<<Tous>>');
  E_NATUREPIECE.plus   := ' AND CO_CODE IN ("AC","AF","FC","FF")' ;
  E_NATUREPIECE.Text   := 'AC;AF;FC;FF;' ;
  E_Exercice.ItemIndex := 0;
  BImprimer.Hint       := TraduireMemoire('Imprimer la grille sélectionnée');
  {$IFDEF EAGLCLIENT}
  BImprimer.Visible    := True;
  {$ENDIF}


  CEN_STATUTENG.text  := TraduireMemoire('<<Tous>>');

  GEnga.PopupMenu      := PopupGEnga;
  GCpta.PopupMenu      := PopupGCpta;
  GEnga.OnEnter        := OnGridEnter;
  GCpta.OnEnter        := OnGridEnter;

  LeCompte.OnExit           := OnExitLeCompte;
  LeCompte.OnElipsisClick   := OnElipsisClickLeCompte;
  LeCompte.OnDblClick       := OnElipsisClickLeCompte;
  BGUP.OnClick              := OnClickBGUP;
  BGDOWN.OnClick            := OnClickBGDOWN;
  PopupGestionEng.OnPopUp   := OnPopupPopupGestionEng;
  PopupGEnga.OnPopup        := OnPopupPOPUPGEnga;
  PopupGCpta.OnPopup        := OnPopupPOPUPGCpta;

  PlanCompta.OnChange       := PlanComptaChange ; 
  E_EXERCICE.OnChange 	    := E_EXERCICEChange ;

  BChercher.OnClick         := OnClickBChercher ;
  BImprimer.OnClick         := OnClickBImprimer ;
  BVoirEngLies.OnClick      := OnClickBVoirEngLies ;
  BAgrandir.OnClick         := OnAgrandirClick;
  BRechercher.OnClick       := OnRechercherClick;

  Composants.Filtres  := THValComboBox   (Getcontrol('FFILTRES'));
  Composants.Filtre   := TToolBarButton97(Getcontrol('BFILTRES'));
  Composants.PageCtrl := TPageControl    (Getcontrol('ONGLETSHAUT'));
  ObjFiltre           := TObjFiltre.Create(Composants, 'CPGESTENG_TOF');
  LstPiecesEngOuFac   := TList.Create;
  LstPiecesEng        := TList.Create;

  FAgrandirEcran       := False;
  FFiltreEcrRapproAuto := '';
  FFiltreEngRapproAuto := '';
  if Length(S) <> 0 then
    begin
    Argument := S;
    if ReadTokenSt(Argument) = 'RAPPROAUTO' then
      begin
      LeCompte.Text        := ReadTokenSt(Argument);
      FFiltreEcrRapproAuto := ReadTokenPipe(Argument, '|');
      FFiltreEngRapproAuto := WhereEcriture(tsGene, DecodeLC(ReadTokenPipe(Argument, '|')), True);
      BValider.Hint        := TraduireMemoire('Rapprocher');
      BValider.OnClick     := OnValiderRapproAuto;
      BChercher.OnClick    := OnChercherRapproAuto;
      SetControlVisible('BOUTILSGEST', False);
      SetControlVisible('FILTRETOOLWINDOW971', False);
      Ecran.Caption := TraduireMemoire('Rapprochement d''un engagement');
      UpdateCaption(Ecran);
      end;
    end;
end ;

procedure TOF_CPGESTENG.OnClose ;
begin
  Inherited ;
  if Assigned(ObjFiltre) then FreeAndNil(ObjFiltre);
  if Assigned(FTobComptes) then FreeAndNil(FTobComptes);
  if Assigned(TobEcrCpta) then FreeAndNil(TobEcrCpta);
  if Assigned(TobEcrEng) then FreeAndNil(TobEcrEng);
  LMsg.Free;
  VideListe(LstPiecesEngOuFac) ;
  LstPiecesEngOuFac.Free;
  VideListe(LstPiecesEng) ;
  LstPiecesEng.Free;
end;

procedure TOF_CPGESTENG.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_CPGESTENG.OnCancel () ;
begin
  Inherited ;
end ;

function TOF_CPGESTENG.ChargeLeCompteGene(Force: boolean): Boolean;
var
  Text    : string;
  lQuery : TQuery;
begin
  Result := False;
  if LeCompte.Text = '' then begin
    FBoConfidentielCpt := False;
    FTobComptes.InitValeurs;
    CLIBELLE.Caption := '';
    LeCompte.DataType := 'TZGENERAL';
    Exit;
    end;
  LeCompte.Text := BourreEtLess(LeCompte.Text, fbgene);
  Text := 'SELECT * FROM GENERAUX WHERE G_GENERAL = "' + LeCompte.Text + '"';
  try
    lQuery := OpenSQL(Text, True);
    if not lQuery.Eof then begin
      FTobComptes.SelectDB('', lQuery);
      // Test de la confidentialité des comptes
      FBoConfidentielCpt := EstConfidentiel(FTobComptes.GetValue('G_CONFIDENTIEL'));
      if not FBoConfidentielCpt then begin
        CLibelle.Caption := FTobComptes.GetValue('G_LIBELLE')
        //VerifieLesDates;
        end;
     end
    else begin
      LeCompte.DataType := 'TZGENERAL';
      LeCompte.ElipsisClick(nil);
      end;
  finally
    Ferme( lQuery );
    end;
end;

function TOF_CPGESTENG.ChargeLeCompteTier(Force: boolean): Boolean;

  procedure InitDataTypeCpte;
  begin
    if PlanCompta.value = 'F' then begin
      LeCompte.DataType := 'TZTFOURN';
      LeCompte.Plus     := ' AND T_NATUREAUXI="FOU" ';
      end
     else if PlanCompta.value = 'C' then begin
      LeCompte.DataType := 'TZTCLIENT';
      LeCompte.Plus     := ' AND T_NATUREAUXI="CLI" ';
      end
     else if PlanCompta.value = 'G' then
       LeCompte.DataType := 'TZGENERAL';
  end;

var
  Text    : string;
  lQuery : TQuery;
begin {ChargeLeCompteTier}
  Result := False;
  if LeCompte.Text = '' then begin
    FBoConfidentielCpt := False;
    FTobComptes.InitValeurs;
    CLIBELLE.Caption := '';
    InitDataTypeCpte;
    Exit;
    end;
  LeCompte.Text := BourreEtLess(LeCompte.Text, fbAux);
  Text := 'SELECT * FROM TIERS WHERE T_AUXILIAIRE = "' + LeCompte.Text + '"';
  try
    lQuery := OpenSQL(Text, True);
    if not lQuery.Eof then begin
      FTobComptes.SelectDB('', lQuery);
      // Test de la confidentialité des comptes
      FBoConfidentielCpt := EstConfidentiel(FTobComptes.GetValue('T_CONFIDENTIEL'));
      if not FBoConfidentielCpt then begin
        CLibelle.Caption := FTobComptes.GetValue('T_LIBELLE')
        //VerifieLesDates;
        end;
      end
    else begin
      InitDataTypeCpte; // selon le type de compte Four/Clie
      LeCompte.ElipsisClick(nil);
      end;
  finally
    Ferme( lQuery );
    end;
end;

function TOF_CPGESTENG.ChargeLeCompte(Force: boolean): Boolean;
begin
  Result := False;
  if ((FTobComptes = nil) or (csDestroying in LeCompte.ComponentState)) then
    Exit;
  if PlanCompta.value = 'G'
    then ChargeLeCompteGene(Force)
    else ChargeLeCompteTier(Force) ; //  Fourn/Client
end;

procedure TOF_CPGESTENG.E_EXERCICEChange(Sender: TObject);
begin
  ExoToDates( GetControlText('E_EXERCICE'), TEdit(GetControl('E_DATECOMPTABLE', True)), TEdit(GetControl('E_DATECOMPTABLE_', True)) ) ;
  if ((GetControlText('E_EXERCICE')='')) then begin
    SetControlText('E_DATECOMPTABLE',     stDate1900) ;
    SetControlText('E_DATECOMPTABLE_',	stDate2099) ;
    end;
end;

procedure TOF_CPGESTENG.GeneSuivant(vBoSuiv, vBoREsteSurZone: Boolean);
var
  lQuery: TQuery;
  lSQL, lCpt, lNewCpt, lCptLibelle: string;
begin
  lCpt     := LeCompte.text;
  lNewCpt  := lCpt;

  if vBOSuiv then
    lSQL := 'SELECT G_GENERAL ,G_LIBELLE FROM GENERAUX WHERE G_GENERAL >"' + lCpt + '"'
  else
    lSQL := 'SELECT G_GENERAL, G_LIBELLE FROM GENERAUX WHERE G_GENERAL <"' + lCpt + '"';
  // Gestion du V_Pgi.Confidentiel
  lSql := lSql + ' AND ' + CGenereSQLConfidentiel('G');
  // Order By
  if vBoSuiv then
    lSQL := lSQL + ' ORDER BY G_GENERAL '
  else
    lSQL := lSQL + ' ORDER BY G_GENERAL DESC ';

  lQuery := OpenSQL(lSQL, True);
  if not lQuery.EOF then begin
    lNewCpt     := lQuery.FindField('G_GENERAL').AsString;
    lCptLibelle := lQuery.FindField('G_LIBELLE').AsString;
    end;
  Ferme(lQuery);

  if lNewCpt <> lCpt then begin
    LeCompte.Text := lNewCpt;
    CLIBELLE.Caption := lCptLibelle;
    ChargeLeCompteGene(False);
    //RefreshPclPge;
    if vBoResteSurZone then
      SetFocusControl('YSECTION');
    end;
end;

procedure TOF_CPGESTENG.TiersSuivant(vBoSuiv, vBoREsteSurZone: Boolean);
var
  lQuery:                TQuery;
  sNat, lSQL, lCpt,
  lNewCpt, lCptLibelle:  string;
begin
  lCpt     := LeCompte.text;
  lNewCpt  := lCpt;

  if PlanCompta.value = 'C' then
    sNat := 'CLI'
  else
    sNat := 'FOU';

  if vBOSuiv then
    lSQL := 'SELECT T_AUXILIAIRE ,T_LIBELLE FROM TIERS WHERE T_AUXILIAIRE >"' + lCpt + '"'
  else
    lSQL := 'SELECT T_AUXILIAIRE, T_LIBELLE FROM TIERS WHERE T_AUXILIAIRE <"' + lCpt + '"';
  // Gestion du V_Pgi.Confidentiel
  lSql := lSql + ' AND ' + CGenereSQLConfidentiel('T') + ' And T_NATUREAUXI = "' + sNat + '"';
  // Order By
  if vBoSuiv then
    lSQL := lSQL + ' ORDER BY T_AUXILIAIRE '
  else
    lSQL := lSQL + ' ORDER BY T_AUXILIAIRE DESC ';

  lQuery := OpenSQL(lSQL, True);
  if not lQuery.EOF then begin
    lNewCpt     := lQuery.FindField('T_AUXILIAIRE').AsString;
    lCptLibelle := lQuery.FindField('T_LIBELLE').AsString;
    end;
  Ferme(lQuery);

  if lNewCpt <> lCpt then begin
    LeCompte.Text := lNewCpt;
    CLIBELLE.Caption := lCptLibelle;
    ChargeLeCompteTier(False);
    //RefreshPclPge;
    if vBoResteSurZone then
      SetFocusControl('YSECTION');
    end;
end;

procedure TOF_CPGESTENG.LeCompteSuivant(vBoSuiv, vBoREsteSurZone: Boolean);
begin
  if PlanCompta.value = 'G'
    then GeneSuivant(vBoSuiv, vBoREsteSurZone)
    else TiersSuivant(vBoSuiv, vBoREsteSurZone);  //  Fourn/Client

  if GetControlText('LeCompte') <> '' then
    BChercher.Click;

end;

procedure TOF_CPGESTENG.OnClickBGDOWN(Sender: TObject);
begin
  LeCompteSuivant(True, True);
end;

procedure TOF_CPGESTENG.OnClickBGUP(Sender: TObject);
begin
  LeCompteSuivant(False, True);
end;

procedure TOF_CPGESTENG.OnElipsisClickLeCompte(Sender: TObject);
var sType, sNat : string;
begin
  sType := PlanCompta.Items[PlanCompta.ItemIndex];
  if PlanCompta.value = 'G' then // compte généraux
    LookUpList(THEdit(Sender),
                'Comptes généraux',
                'GENERAUX',
                'G_GENERAL' ,
                'G_LIBELLE',
                CGenereSQLConfidentiel('G'),
                'G_GENERAL' ,
                True,
                0 )  {0=enlève l'accés au cpte, 1=géné, 2=tiers}
  else begin //compte Fourn ou Client
    if PlanCompta.value = 'C' then
      sNat := 'CLI'
    else
      sNat := 'FOU';
    LookUpList(THEdit(Sender),
                'Comptes ' + sType,
                'TIERS',
                'T_AUXILIAIRE' ,
                'T_LIBELLE',
                CGenereSQLConfidentiel('T')  + ' And T_NATUREAUXI = "' + sNat + '"' ,
                'T_AUXILIAIRE',
                True,
                0 );
    end;
end;

procedure TOF_CPGESTENG.OnExitLeCompte(Sender: TObject);
begin
  ChargeLeCompte(False);
end;

procedure TOF_CPGESTENG.OnPopupPopupGestionEng(Sender: TObject);
begin
  InitPopUp(True, 1);
end;

procedure TOF_CPGESTENG.OnPopupPOPUPGCpta(Sender: TObject);
begin
  InitPopUp(True, 2);
end;

procedure TOF_CPGESTENG.OnPopupPOPUPGEnga(Sender: TObject);
begin
  InitPopUp(True, 3);
end;

procedure TOF_CPGESTENG.InitPopUP(vActivation: Boolean; vNumPopUp : integer =  0);
var
  i : integer;
begin
  if vNumPopUp= 1 then
  for i := 0 to PopupGestionEng.Items.Count -1 do begin
    if PopupGestionEng.Items[i].Name = 'TRANSFACT' then begin
      PopupGestionEng.Items[i].OnClick := TransFactClick; // transformer en fact
      if vActivation then
        if isEngagementHT then
          begin
          PopupGestionEng.Items[i].visible := False;
          PopupGestionEng.Items[i].Enabled := False;
          end
        else PopupGestionEng.Items[i].Enabled := IsTransFactAutorise;    {FP 02/05/2006}
      end;

    if PopupGestionEng.Items[i].Name = 'RAPMANU' then begin
      PopupGestionEng.Items[i].OnClick := RapManuelClick;
      if vActivation then
        PopupGestionEng.Items[i].Enabled := IsRapproManuelAutorise;
      end;

    if PopupGestionEng.Items[i].Name = 'RAPAUTO' then begin
      PopupGestionEng.Items[i].OnClick := RapAutoClick;
      if vActivation then
        PopupGestionEng.Items[i].Visible := False; // aà supp car l'accès se fera via le menu.
      end;

    if PopupGestionEng.Items[i].Name = 'DERAP' then begin
      PopupGestionEng.Items[i].OnClick := DeRapproClick;
      if vActivation then
        PopupGestionEng.Items[i].Enabled := IsDerapproAutorise;
      end;
    if PopupGestionEng.Items[i].Name = 'TERMENG' then begin // terminer un engagement.
      PopupGestionEng.Items[i].OnClick := TerminerUnEngClick;
      if vActivation then
        PopupGestionEng.Items[i].Enabled := IsTerminerUnEngagementAutorise;
      end;
    end;

  if vNumPopUp= 2 then
  for i := 0 to PopupGCpta.Items.Count -1 do begin
    if PopupGCpta.Items[i].Name = 'ENGALIES' then begin
      PopupGCpta.Items[i].OnClick := AffListeEngagementsLies;
      if vActivation then
        PopupGCpta.Items[i].Enabled := IsExistEngagementsLies;
      end;
    end;
  if vNumPopUp= 3 then
  for i := 0 to PopupGEnga.Items.Count -1 do begin
    if PopupGEnga.Items[i].Name = 'FACTULIEES' then begin
      PopupGEnga.Items[i].OnClick := AffListeEngagementsLies;
      if vActivation then
        PopupGEnga.Items[i].Enabled := IsExistFacturesLiees;
      end;
    end;
end;

procedure TOF_CPGESTENG.DeRapproClick(Sender: TObject);
  function CleEngCourant(M : RMVT): string;
  begin
    Result := M.JAL + '/' + IntToStr(M.Num) + '/' + M.EXO;
  end;

Var
  M           : RMVT ;
  i           : integer ;
  S           : string ;
  sEng_prec   : string;
  lEng        : TPieceEngagement;
  Ok          : Boolean;
  P           : P_MV;
begin
  Ok := False;
  if FSelectedGrid = GEnga then
    begin
    if not IsExistFacturesLiees then   {C'est un engagement terminé. Il n'y a donc pas de facture rattachée}
      begin
      P   := P_MV.Create;
      P.R := TOBToIdent(TobEcrEng.Detail[GridRowToTobIndex(GEnga, GEnga.Row)], False);
      LstPiecesEng.Add(P); // liste des Eng pour dé-rapprocher
      end;
    Ok := True;
    end
  else if FSelectedGrid = GCpta then
    Ok := IsExistEngagementsLies;
  Ok := Ok and (LstPiecesEng.Count>0);

  if Ok and (AfficheMsg(03, '', '') = mrYes) then begin  //confirmer le dé-rappro
    sEng_prec := '';
    for i := 0 to LstPiecesEng.Count-1 do begin
      M := P_MV(LstPiecesEng[i]).R ;
      S := WhereEcriture(tsGene, M, False);
      if (sEng_prec <> CleEngCourant(M)) then begin //il pourait avoir 2 lignes tiers ds la mm pièce si multi échéance.
        sEng_prec := CleEngCourant(M);
        lEng := TPieceEngagement.CreerPieceEng;
        try
          lEng.LoadFromSQL(S);
          if not DerapprocherEngagement(lEng) then begin
            AfficheMsg(04,'',''); // echec dé-rappro
            Break;
            end;
        finally
          lEng.Free;
          end;
        end;
      end;{for}
    BChercher.Click;            // réaff les grilles.
    end;
end;

procedure TOF_CPGESTENG.RapManuelClick(Sender: TObject);
begin
  RapprochementManuel(nil);
  BChercher.Click;            // réaff les grilles.
end;

procedure TOF_CPGESTENG.TerminerUnEngClick(Sender: TObject);
var
  lEng :  TPieceEngagement;
  lDate:  TDateTime;
  cleEng:  string;
begin
  if (GetStatutEng(GetRowSelect(GEnga), TobEcrEng) = 'T') then Exit;
  cleEng := GetClauseRechEngOuFac(GetRowSelect(GEnga), TobEcrEng);
  if  (cleEng = '') then  Exit;
  if (AfficheMsg(02, '', '') <> mrYes) then  Exit; //confirmation.

  lDate := Date;
  lEng := TPieceEngagement.CreerPieceEng;
  BeginTrans;
  try
    lEng.LoadFromSQL (cleEng);
    if ExtourneEng (lEng, nil, lDate) then
      CommitTrans
    else rollback;
  except rollback;
  end;
  lEng.Free;

  BChercher.Click;            // réaff les grilles.
end;

{b FP 02/05/2006}
type
  TInfoParamSaisieFacture = class
  private
    FTof: TOF_CPSAISIEPIECE;
    FEcran: TForm;
    FPiece: TPieceCompta;

    FJOURNALExit:       TNotifyEvent;
    FNATUREPIECEExit:   TNotifyEvent;
    FDATECOMPTABLEExit: TNotifyEvent;
    FDEVISEExit:        TNotifyEvent;
    FETABLISSEMENTExit: TNotifyEvent;

    procedure InitComposant;
  public
    constructor Create;
    destructor  Destroy; override;

    function  OnBeforeValidePieceCompta(Tof: TOF_CPSAISIEPIECE; Ecran: TForm; Piece: TPieceCompta): Boolean;
    procedure OnBeforeChargePiece(Tof: TOF_CPSAISIEPIECE; Ecran: TForm; Piece: TPieceCompta; var Action: TActionFiche);
    procedure OnAfterChargePiece(Tof: TOF_CPSAISIEPIECE; Ecran: TForm; Piece: TPieceCompta);

    procedure OnJOURNALExit(Sender: TObject);
    procedure OnNATUREPIECEExit(Sender: TObject);
    procedure OnDATECOMPTABLEExit(Sender: TObject);
    procedure OnETABLISSEMENTExit(Sender: TObject);
  public

  end;
{e FP 02/05/2006}

procedure TOF_CPGESTENG.TransFactClick(Sender: TObject);

  procedure InitMontantResteRappro(PieceCompta: TPieceEngagement);
  var
    i:          Integer;
    iLigne:     Integer;
    Champ:      String;
  begin
    for i:=0 to PieceCompta.Count-1 do
      begin
      iLigne := i+1;
      if PieceCompta.GetValue(iLigne, 'E_DEBITDEV') = 0 then
        Champ := 'E_CREDITDEV'
      else
        Champ := 'E_DEBITDEV';
      PieceCompta.PutValue(iLigne, Champ,  Abs(PieceCompta.GetSoldeLigne(iLigne)));
      end;
  end;

  function CreationPieceSimul(TobEcrEng: TOB; PieceCompta: TPieceEngagement; LstPieceEng: TStrings): Boolean;
  var
    i:           Integer;
    NbPiece:     Integer;
    NbLigne:     Integer;
    M:           RMVT;
    TobEng:      TOB;
  begin
    FillChar(M, sizeof(M), #0);
    NbPiece := 0;
    NbLigne := 0;
      for i:=0 to TobEcrEng.Detail.Count-1 do
        begin
        TobEng := TobEcrEng.Detail[i];
        if StrToBool_(TobEng.GetString('MARQUE')) then
          begin
          if NbPiece = 0 then
            begin
            {Transforme la première pièce d'engagement en pièce comptable de type simulation}
            PieceCompta.LoadFromSQL( WhereEcritureTOB(tsGene, TobEng, False) );
            PieceCompta.Action := taCreat;
            PieceCompta.PutEntete('E_QUALIFPIECE',   'S') ;
            PieceCompta.PutEntete('E_DATECOMPTABLE', V_PGI.DateEntree);
            PieceCompta.PutEntete('E_MODESAISIE',    '-');
            PieceCompta.AttribNumeroTemp;
            InitMontantResteRappro(PieceCompta);
            if PieceCompta.Save then
              begin
              M       := PieceCompta.GetRMVT;
              NbLigne := PieceCompta.Count;
              LstPieceEng.AddObject(IntToStr(PieceCompta.Count), TobEng);
              end;
            end
          else if M.Jal <> '' then
            begin
            {Ajoute les autres pièces d'engagement à la nouvelle pièce comptable}
            PieceCompta.LoadFromSQL( WhereEcritureTOB(tsGene, TobEng, False) );
            PieceCompta.Action := taModif;
            PieceCompta.InitPiece(M);
            PieceCompta.PutEntete('E_NUMEROPIECE', M.Num);
            PieceCompta.RenumeroteLignes(NbLigne);  {Renumérote les lignes pour ne pas avoir de doubons}
            PieceCompta.TobOrigine.ClearDetail;     {Evite de supprimer la pièce d'origine}
            InitMontantResteRappro(PieceCompta);
            if PieceCompta.Save then
              begin
              Inc(NbLigne, PieceCompta.Count);
              LstPieceEng.AddObject(IntToStr(PieceCompta.Count), TobEng);
              end
            else
              M.Jal := '';
            end;
          Inc(NbPiece);
          end;
        end;
    Result := (NbPiece <> 0) and (NbLigne <> 0) and (M.Jal <> '');
  end;

  function LanceSaisieSimul(PieceCompta: TPieceCompta; LstPieceEng: TStrings): Boolean;
  var
    i, j:       Integer;
    NbLigne:    Integer;
    TobEcr:     TOB;
    PBeforeVal: PBeforeValidePieceCompta;
    PSuppLigne: PBeforeSuppLigne;
  begin
    {charge toutes les lignes de la pièce}
    PieceCompta.LoadFromSQL;

    {Pour chaque ligne indique la pièce d'engagement d'origine}
    NbLigne := 0;
    for i:=0 to LstPieceEng.Count-1 do
      begin
      for j:=NbLigne to NbLigne+StrToInt(LstPieceEng[i])-1 do
        begin
        TobEcr := PieceCompta.Detail[j];
        TobEcr.AddChampSup('ENGNUMPIECEORIG', True);
        TobEcr.PutValue('ENGNUMPIECEORIG', i+1);
        end;
      Inc(NbLigne, StrToInt(LstPieceEng[i]));
      end;

    System.New(PBeforeVal);
    System.New(PSuppLigne);
    try
      PBeforeVal^ := OnBeforeValidePieceComptaSimul;
      PieceCompta.AddChampSup('BeforeValidePieceCompta', False);
      PieceCompta.PutValue('BeforeValidePieceCompta', Integer(PBeforeVal));

      PSuppLigne^ := OnBeforeSuppLigne;
      PieceCompta.AddChampSup('BeforeSuppLigne', False);
      PieceCompta.PutValue('BeforeSuppLigne', Integer(PSuppLigne));

      PieceCompta.ModifEnCours := True;       {Permet de valider la pièce sans avoir à la modifier}
      Result := ModifiePieceCompta(PieceCompta, taModif);
      PieceCompta.OnError := nil;
      PieceCompta.PutValue('BeforeValidePieceCompta', 0);
      PieceCompta.PutValue('AfterFormShow', 0);
    finally
      System.Dispose(PBeforeVal);
      System.Dispose(PSuppLigne);
      end;
  end;

  procedure FusionnePieceSimul(PieceAllSimul, PieceFacture: TPieceCompta);
  var
    LstCpt:   TStrings;
    i, j:     Integer;
    s:        String;
    sGene:    String;
    sAuxi:    String;
    TobEcr:   TOB;
    NumLigneSrc: Integer;
    NumLigneDst: Integer;
  begin
    PieceFacture.InitPiece(PieceAllSimul.GetRMVT);              
    LstCpt := TStringList.Create;
    try
      {Liste de tous les comptes différents}
      for i:=0 to PieceAllSimul.Detail.Count-1 do
        begin
        TobEcr := PieceAllSimul.Detail[i];
        s      := TobEcr.GetValue('E_GENERAL')+';'+TobEcr.GetValue('E_AUXILIAIRE');
        if LstCpt.IndexOf(s) = -1 then
          LstCpt.Add(s);
        end;

      {Ajoute une écriture par compte dans PieceFacture}
      for i:=0 to LstCpt.Count-1 do                        
        begin
        for j:=0 to PieceAllSimul.Detail.Count-1 do
          begin
          NumLigneSrc := j+1;
          s := LstCpt[i];
          sGene := ReadTokenSt(s);
          sAuxi := ReadTokenSt(s);
          if (sGene = PieceAllSimul.GetValue(NumLigneSrc,'E_GENERAL')) and
             (sAuxi = PieceAllSimul.GetValue(NumLigneSrc,'E_AUXILIAIRE')) then
            begin
            NumLigneDst := PieceFacture.TrouveIndiceLigneCompte(sGene, sAuxi);
            if NumLigneDst = -1 then   
              PieceFacture.DupliquerLigne(PieceFacture.Detail.Count+1, PieceAllSimul, NumLigneSrc)
            else        
              PieceFacture.CumulerMontant(NumLigneDst, PieceAllSimul, NumLigneSrc);
            end;         
          end;
        end;
      PieceFacture.RenumeroteLignes;                 
      if not PieceFacture.EstPieceEquilibree then
        PieceFacture.AttribSolde(PieceFacture.Detail.Count);
    finally
      LstCpt.Free;
      end;
  end;

  function LanceSaisieFacture(PieceCompta: TPieceCompta): Boolean;
  var
    InfoSaisie:       TInfoParamSaisieFacture;
    PBeforeChrgPiece: PBeforeChargePiece;
    PAfterChrgPiece:  PAfterChargePiece;
    PBeforeVal:       PBeforeValidePieceCompta;
  begin
    InfoSaisie := TInfoParamSaisieFacture.Create;
    System.New(PBeforeChrgPiece);
    System.New(PAfterChrgPiece);
    System.New(PBeforeVal);
    try
      PBeforeChrgPiece^ := InfoSaisie.OnBeforeChargePiece;
      PieceCompta.AddChampSup('BeforeChargePiece', False);
      PieceCompta.PutValue('BeforeChargePiece', Integer(PBeforeChrgPiece));

      PAfterChrgPiece^ := InfoSaisie.OnAfterChargePiece;
      PieceCompta.AddChampSup('AfterChargePiece', False);
      PieceCompta.PutValue('AfterChargePiece', Integer(PAfterChrgPiece));

      PBeforeVal^ := InfoSaisie.OnBeforeValidePieceCompta;
      PieceCompta.AddChampSup('BeforeValidePieceCompta', False);
      PieceCompta.PutValue('BeforeValidePieceCompta', Integer(PBeforeVal));

      PieceCompta.PutEntete('E_QUALIFPIECE', 'N');
      PieceCompta.Action := taCreat;
      PieceCompta.AttribNumeroTemp;

      PieceCompta.ModifEnCours := True;       {Permet de valider la pièce sans avoir à la modifier}
      Result := ModifiePieceCompta(PieceCompta, taCreat);
      PieceCompta.OnError := nil;
      if Result then
        Result := PieceCompta.Save;
    finally
      System.Dispose(PBeforeChrgPiece);
      System.Dispose(PAfterChrgPiece);
      System.Dispose(PBeforeVal);
      FreeAndNil(InfoSaisie);
      end;
  end;

  function ExtractPieceEng(PieceAllSimul: TPieceEngagement; PieceEngSimul: TPieceEngagement; NumPiece: Integer): Boolean;
  var
    i    : Integer;
    iLig : integer ;
    lTob : Tob ;
  begin
    for i:=0 to PieceAllSimul.Count-1 do
      if PieceAllSimul.GetValue(i+1, 'ENGNUMPIECEORIG') = NumPiece then
        begin
        iLig := PieceEngSimul.Count+1 ;
        PieceEngSimul.DupliquerLigne( iLig, PieceAllSimul, i+1);
        lTob := CGetTOBEngage( PieceAllSimul.Detail[i] ) ;
        if Assigned( lTob ) then
          CPutTobEngage( PieceEngSimul.CurTob, lTob  );
        end ;
    PieceEngSimul.RenumeroteLignes;
    Result := PieceEngSimul.Count > 0;
  end;

  procedure RapprochementEngFac(PieceAllSimul: TPieceEngagement; PieceFacture: TPieceEngagement; LstPieceEng: TStrings);
  var
    TobEng:         TOB;
    PieceEng:       TPieceEngagement;
    PieceEngSimul:  TPieceEngagement;
    iPieceEng:      Integer;
    Ok:             Boolean;
  begin
    Ok        := True;
    PieceEng  := TPieceEngagement.CreerPieceEng;
    try
      BEGINTRANS;
      try
        for iPieceEng:=0 to LstPieceEng.Count-1 do
          begin
          TobEng := (LstPieceEng.Objects[iPieceEng] as TOB);
          PieceEng.LoadFromSQL( WhereEcritureTOB(tsGene, TobEng, False) );
          PieceEngSimul := TPieceEngagement.CreerPieceEng;
          try
            Ok := Ok and ExtractPieceEng(PieceAllSimul, PieceEngSimul, iPieceEng+1);
            Ok := Ok and ExtourneEng(PieceEng, nil, V_PGI.DateEntree);
            Ok := Ok and RapprocherEngFac(PieceEng, PieceEngSimul, PieceFacture,  V_PGI.DateEntree);
            Ok := Ok and (PieceEng.IsEngagementEstTermine or ResteEng(PieceEng, nil));
          finally
            FreeAndNil(PieceEngSimul);
            end;
          if not Ok then
            break;
          end;
        if Ok then
          COMMITTRANS
        else
          ROLLBACK;
      except on E: Exception do
        begin
        ROLLBACK;
        PGIError(E.message);
        end;
      end;
    finally
      FreeAndNil(PieceEng);   
      end;
  end;
var
  PieceAllSimul: TPieceEngagement;     {Pièce regroupant toutes les pièces d'engagements. C'est une pièce de simulation}
  PieceFacture:  TPieceEngagement;
  LstPieceEng:   TStrings;
begin               {TransFactClick}
  if (AfficheMsg(07, '', '') <> mrYes) then  Exit; //confimer la transformation en fact

  {b FP 02/05/2006}
  PieceAllSimul := TPieceEngagement.CreerPieceEng(nil);
  PieceFacture  := TPieceEngagement.CreerPieceEng(nil);
  LstPieceEng   := TStringList.Create;
  try
    if CreationPieceSimul(TobEcrEng, PieceAllSimul, LstPieceEng) then
      begin
       if LanceSaisieSimul(PieceAllSimul, LstPieceEng) then
         begin
         FusionnePieceSimul(PieceAllSimul, PieceFacture);
         if LanceSaisieFacture(PieceFacture) then
           RapprochementEngFac(PieceAllSimul, PieceFacture, LstPieceEng);
         PieceAllSimul.DetruitPiece;
         end;
      end;
  finally
    FreeAndNil(PieceAllSimul);
    FreeAndNil(PieceFacture);
    FreeAndNil(LstPieceEng);
    end;
  {e FP 02/05/2006}

  BChercher.Click;            // réaff les grilles.
end;

procedure TOF_CPGESTENG.RapAutoClick(Sender: TObject);
begin
// ce sera ds le menu mdispS5
end;

{procedure TOF_CPGESTENG.InitCriteres;
begin
  if VH^.Precedent.Code<>'' then
    E_DATECOMPTABLE.Text := DateToStr(VH^.Precedent.Deb)
  else
    E_DATECOMPTABLE.Text := DateToStr(VH^.Encours.Deb) ;
  E_DATECOMPTABLE_.Text := DateToStr(V_PGI.DateEntree) ;
end;
}

function TOF_CPGESTENG.InitControles: boolean;
begin
  Try
    if Ecran<>nil then begin
       TFORM(Ecran).OnKeyDown:=OnKeyDown ;
       if V_PGI.LaSerie = S7 then
            TFORM(Ecran).OnResize := OnFormResize ;
    end;
    ONGLETSHAUT.ActivePage := ONGLETSHAUT.Pages[0];
    // Paramètrage tables libres
    LibellesTableLibre(TTabSheet(GetControl('ONGLETTABLIBRES', True)),'TE_TABLE','E_TABLE','E') ;

//    InitCriteres ;

    if VH^.CPExoRef.Code<>'' then begin
      E_EXERCICE.Value      := VH^.CPExoRef.Code ;
      E_DATECOMPTABLE.Text  := DateToStr(VH^.CPExoRef.Deb) ;
      E_DATECOMPTABLE_.Text := DateToStr(VH^.CPExoRef.Fin) ;
      end
    else begin
      E_EXERCICE.Value      := VH^.Entree.Code ;
      E_DATECOMPTABLE.Text  := DateToStr(V_PGI.DateEntree) ;
      E_DATECOMPTABLE_.Text := DateToStr(iDate2099-1);
      end ;

    E_DATECREATION.Text   := DateToStr(iDate1900);         {FP 02/05/2006}
    E_DATECREATION_.Text  := DateToStr(iDate2099);
    E_DATEECHEANCE.Text   := DateToStr(iDate1900);         {FP 02/05/2006}
    E_DATEECHEANCE_.Text  := DateToStr(iDate2099);

    E_DATEREFEXTERNE.Text  := DateToStr(iDate1900);        {FP 02/05/2006}
    E_DATEREFEXTERNE_.Text := DateToStr(iDate2099);

    PCpta.Height := pgene.ClientHeight div 2;

    InitGrid(GEnga) ;
    InitGrid(GCpta) ;
    BVoirEngLies.Enabled := False;
  except
    Result := False ;
    PgiBox(TraduireMemoire('Une erreur est survenue lors de l''initialisation de la fiche, vérifiez la version de votre base'),'Erreur affichage fiche') ;
    exit;
  end;
  Result:=true ;

end;

procedure TOF_CPGESTENG.AfficheSoldes;
begin
  SoldeEng.Text    := StrS(TotSoldeEng,   V_PGI.OkDecV);
  SoldeCpta.Text   := StrS(TotSoldeCpta,  V_PGI.OkDecV);
end;

function TOF_CPGESTENG.ClauseJoinEng(Fact: boolean): string;
begin
  if Fact then
    Result := ' LEFT OUTER JOIN CENGAGEMENT '
  else
    Result := ' INNER JOIN CENGAGEMENT ';
  Result := Result +
    ' ON E_EXERCICE=CEN_EXERCICE AND E_JOURNAL=CEN_JOURNAL ' +
    ' AND E_DATECOMPTABLE=CEN_DATECOMPTABLE AND E_NUMEROPIECE=CEN_NUMEROPIECE ' +
    ' AND E_NUMLIGNE=CEN_NUMLIGNE AND E_QUALIFPIECE=CEN_QUALIFPIECE ';
 //   ' AND E_NUMECHE=CEN_NUMECHE '; faut-il le mettre ?
end;

function TOF_CPGESTENG.ClauseStatutEng :string;
var Cod, St, sWhere:  string;
begin
  sWhere := '' ;
  if (CEN_STATUTENG.Text <> '<<Tous>>') and  (CEN_STATUTENG.Text <> '') then  begin
    Cod := CEN_STATUTENG.Text ;
    St := ReadTokenPipe(Cod, ';') ;
    sWhere := ' AND (CEN_STATUTENG ="' + st + '" ' ;
    St := ReadTokenpipe(Cod, ';') ;
    while St <> '' do begin
      sWhere:= sWhere + ' OR CEN_STATUTENG="' + st + '" ' ;
      St := ReadTokenPipe(Cod, ';') ;
      end;
    sWhere := sWhere + ') AND ' ;
    end
  else
  if (CEN_STATUTENG.Text = '<<Tous>>') then
    sWhere := ' AND (CEN_STATUTENG <> "X") AND ' ;
  Result := sWhere;
end;
{
function TOF_CPGESTENG.ClauseJournaux(JNX: THMultiValComboBox): string;
var Cod, St, sWhere:  string;
begin
  sWhere := '' ;
  if (JNX.Text <> '<<Tous>>') and  (JNX.Text <> '') then  begin
    Cod := JNX.Text ;
    St := ReadTokenPipe(Cod, ';') ;
    sWhere := ' AND (E_JOURNAL ="' + st + '" ' ;
    St := ReadTokenpipe(Cod, ';') ;
    while St <> '' do begin
      sWhere:= sWhere + ' OR E_JOURNAL="' + st + '" ' ;
      St := ReadTokenPipe(Cod, ';') ;
      end;
    sWhere := sWhere + ')' ;
    end;
  Result := sWhere;
end;
}
function TOF_CPGESTENG.ClauseOngletAvance: string;

  function MoinsSymboleInt(St : String) : string;
  Var i : Integer ;
      Alors : String ;                             { Déformate les montants }
  begin
    Alors := '' ;
    if St='' then St := '0' ;
    for i:=1 to Length(St) do
      if St[i] in ['0'..'9'] then Alors := Alors + St[i] ;
    Result := Alors ;
  end; 

  function  GetCompare(Z_C:THValComboBox; sZO, sZV: string) :string;
  var ZO:   THValComboBox ;
  begin
    Result := '';
    if (GetControl(sZO) <> nil) and (GetControl(sZV) <> nil)  then begin
      ZO := THValComboBox(GetControl(sZO, True));
      if (ZO.Value <> '') and (TEdit(GetControl(sZV)).Text <> '') then begin
        case Z_C.ItemIndex of
          0:begin
             if IsValidDate(TEdit(GetControl(sZV)).Text) then
               Result := 'E_DATECOMPTABLE ' + ZO.Value + '"' + UsDateTime(StrToDate(TEdit(GetControl(sZV)).Text)) + '"';
             end;
          1:Result := 'E_JOURNAL '       + ZO.Value + '"' + TEdit(GetControl(sZV)).Text + '"';
          2:Result := 'E_NUMEROPIECE '   + ZO.Value + MoinsSymboleInt(TEdit(GetControl(sZV)).Text);
          3:Result := 'E_REFINTERNE '    + ZO.Value + '"' + TEdit(GetControl(sZV)).Text + '"';
          4:Result := 'E_LIBELLE '       + ZO.Value + '"' + TEdit(GetControl(sZV)).Text + '"';
          5:Result := 'E_DEBIT '         + ZO.Value + FloatToStr(MoinsSymbole(TEdit(GetControl(sZV)).Text));
          6:Result := 'E_CREDIT '        + ZO.Value + FloatToStr(MoinsSymbole(TEdit(GetControl(sZV)).Text));
          end;
        end;
      end;  
  end;

var sW:   string;
begin {ClauseOngletAvance}
//  Z_C1 .. Z_C3  pour les champs,
//  ZO1 .. ZO3,   les opérateurs =, >, > ...
//  ZV1 .. ZV3,   les valeurs saisies
//  ZG1..ZG2   (opérateur Et/Ou
  sW := '';
  if Z_C1.ItemIndex<>-1 then
    sW := GetCompare(Z_C1, 'ZO1', 'ZV1');
  if (sW <> '') then begin  // si le 1er critère est renseigné 
    if Z_C2.ItemIndex<>-1 then begin
      if (GetControl('ZG1') <> nil) then
        if (TComboBox(GetControl('ZG1')).ItemIndex = 1) then
          sW := sW + ' Or '
        else
          sW := sW + ' And ';
      sW := sW + GetCompare(Z_C2, 'ZO2', 'ZV2');
      end;
    if Z_C3.ItemIndex<>-1 then begin
      if (GetControl('ZG2') <> nil) then
        if (TComboBox(GetControl('ZG2')).ItemIndex = 1) then
          sW := sW + ' Or '
        else
          sW := sW + ' And ';
      sW := sW + GetCompare(Z_C3, 'ZO3', 'ZV3');
      end;
    end;  
  Result := sW;
  if Result <> '' then
     Result := ' AND ' + Result;
end;

function TOF_CPGESTENG.ClauseOngleTabLibres: string;
  var E_TABLE:  THEdit;
begin
  Result := '';
  E_TABLE := THEdit(GetControl('E_TABLE0', True));
  if E_TABLE.Enabled and (E_TABLE.Text<>'') then
    Result := ' E_TABLE0 LIKE ' + E_TABLE.Text + '% ';
  E_TABLE := THEdit(GetControl('E_TABLE1', True));
  if E_TABLE.Enabled and (E_TABLE.Text<>'') then begin
    if Result <> '' then
      Result := Result + ' AND ';
    Result := 'E_TABLE1 LIKE ' + E_TABLE.Text + '% ';
    end;
  E_TABLE := THEdit(GetControl('E_TABLE2', True));
  if E_TABLE.Enabled and (E_TABLE.Text<>'') then begin
    if Result <> '' then
      Result := Result + ' AND ';
    Result := 'E_TABLE2 LIKE ' + E_TABLE.Text + '% ';
    end;
  E_TABLE := THEdit(GetControl('E_TABLE3', True));
  if E_TABLE.Enabled and (E_TABLE.Text<>'') then begin
    if Result <> '' then
      Result := Result + ' AND ';
    Result := 'E_TABLE3 LIKE ' + E_TABLE.Text + '% ';
    end;
  if Result <> '' then
     Result := ' AND ' + Result;
end;

function TOF_CPGESTENG.ClauseOngletCompte: string;
var lStSQL : string ;   // FQ 20517 : SBO 08/08/2007 : Ajout critère Etablissement
    lStLib : string ;
begin
  if (PlanCompta.value='G') then // Géné
    Result := ' E_GENERAL ="' + GetControlText('LeCompte') + '" AND E_TYPEMVT = "HT"'
  else
    Result := ' E_AUXILIAIRE ="' + GetControlText('LeCompte') + '" AND E_TYPEMVT = "TTC"' ;
  Result := Result + ' AND E_DEVISE = "' + GetControlText('Devise') + '" ';
  // FQ 20517 : SBO 08/08/2007 : Ajout critère Etablissement
  TraductionTHMultiValComboBox( E_ETABLISSEMENT, lStSQL, lStLib, 'E_ETABLISSEMENT') ;
  Result := Result + ' AND ' + lStSQL + ' ' ;
end;

function TOF_CPGESTENG.ClauseOngletEcrCompta: string;
var lStSQL : string ;
    lStLib : string ;
begin
//  Result := ClauseJournaux(E_Journal);
  TraductionTHMultiValComboBox( E_JOURNAL, result, lStLib, 'E_JOURNAL', False) ;
  if trim( result ) <> '' then // FQ
    result := ' AND ' + result ;
  if E_EXERCICE.Value <> '' then
    Result := Result + ' AND E_EXERCICE = "' + E_EXERCICE.Value +'"';
  TraductionTHMultiValComboBox( E_NATUREPIECE, lStSQL, lStLib, 'E_NATUREPIECE', False) ;
  Result := Result + ' AND ' + lStSQL +
      ' AND E_DATECOMPTABLE >= "'+ UsDateTime(StrToDate(E_DATECOMPTABLE.Text)) + '"' +
      ' AND E_DATECOMPTABLE <= "'+ UsDateTime(StrToDate(E_DATECOMPTABLE_.Text)) + '"' +
      ' AND E_QUALIFPIECE="N"  AND E_ECRANOUVEAU="N" ';

  if GetControlText('DEJASOLDE')<>'X' then
    Result := Result + ' AND ( CEN_MONTANTRAP IS NULL OR (CEN_MONTANTRAP<>(E_DEBIT-E_CREDIT)) ) ' ;

end;

function TOF_CPGESTENG.ClauseOngletEngagement: string;
var lStLib : string ;
begin
//  Result := ClauseJournaux(E_JournalEng);
  TraductionTHMultiValComboBox( E_JournalEng, result, lStLib, 'E_JOURNAL', False) ;
  if Result <> '' then
    Result := Result + ' AND ';
  Result := Result +  ClauseStatutEng +
    ' E_DATEREFEXTERNE >= "'+ UsDateTime(StrToDate(E_DATEREFEXTERNE.Text)) + '"  AND ' +
    ' E_DATEREFEXTERNE <= "'+ UsDateTime(StrToDate(E_DATEREFEXTERNE_.Text)) + '" AND ' +
    ' E_DATECREATION >= "'+   UsDateTime(StrToDate(E_DATECREATION.Text)) + '" AND '  +
    ' E_DATECREATION <= "'+   UsDateTime(StrToDate(E_DATECREATION_.Text)) + '" AND ' +
    ' E_QUALIFPIECE="P" ';
  if not IsCompteGene then
    Result := Result + ' AND ' +
    ' E_DATEECHEANCE >= "'+   UsDateTime(StrToDate(E_DATEECHEANCE.Text)) + '" AND '  +
    ' E_DATEECHEANCE <= "'+   UsDateTime(StrToDate(E_DATEECHEANCE_.Text)) + '" ';
  if (GetControlText('E_REFINTERNE_') <> '') then
    Result := Result + ' AND ' +
    ' E_REFINTERNE >="' + GetControlText('E_REFINTERNE') + '" AND ' +
    ' E_REFINTERNE <="' + GetControlText('E_REFINTERNE_') + '" AND E_QUALIFPIECE="P" ';
  Result := ' AND ' + Result;
end;

function TOF_CPGESTENG.GetClauseSelectEcrCompta: string;
{pour recuperer les écritures cptable (=lesfactures)}
var
  sSelect, sWhere, sOrder: String;
begin
  sSelect := ' SELECT * FROM ECRITURE ';
  sOrder := ' ORDER BY E_DATECOMPTABLE ' ;

  if FFiltreEcrRapproAuto = '' then
    begin
    sWhere :=  ' WHERE ' + ClauseOngletCompte;
    sWhere := sWhere + ClauseOngletEcrCompta;
    sWhere := sWhere + ClauseOngletAvance;
    sWhere := sWhere + ClauseOngleTabLibres;
    end
  else
    sWhere := ' where ' + FFiltreEcrRapproAuto ;

  Result := sSelect + ClauseJoinEng(True) + sWhere + sOrder
end;

function TOF_CPGESTENG.GetClauseSelectEcrEngagement: string;{pour recuperer les écritures d'engagements}
var
  sSelect, sWhere,  sOrder: String;
begin
  sSelect := ' SELECT * FROM ECRITURE ';
  sOrder := ' ORDER BY E_DATECOMPTABLE' ;

  if FFiltreEngRapproAuto = '' then
    begin
    sWhere :=  ' WHERE ' + ClauseOngletCompte;
    sWhere := sWhere + ClauseOngletEngagement;
    sWhere := sWhere + ClauseOngletAvance;
    sWhere := sWhere + ClauseOngleTabLibres;
    end
  else
    sWhere := ' where ' + FFiltreEngRapproAuto;

  Result := sSelect + ClauseJoinEng(False) + sWhere + sOrder
end;

procedure TOF_CPGESTENG.FillTobsEngaEtCpta;
var
  Q:       TQuery;
  Sql:     String ;
begin
  if TobEcrEng <> nil then
    FreeAndNil(TobEcrEng);
  TobEcrEng := TOB.create('ECRITURE', nil, -1) ;                   {FP 02/05/2006}
  Sql := GetClauseSelectEcrEngagement;
  Q := OpenSQL(Sql, true) ;
  Try
    TobEcrEng.LoadDetailDB('ECRITURE', '', '', Q, False, True) ;   {FP 02/05/2006}
  finally
    ferme(Q);                                                       
  end;
                                      
  if TobEcrCpta <> nil then
    FreeAndNil(TobEcrCpta);
  TobEcrCpta := TOB.create('ECRITURE', nil, -1) ;
  Sql := GetClauseSelectEcrCompta;
  Q := OpenSQL(Sql, True) ;
  Try
    TobEcrCpta.LoadDetailDB('ECRITURE', '', '', Q, False, True) ;
    if TobEcrCpta.Detail.Count > 0 then begin
      CalculMtHTEcritCpta;
      end;
  finally
    ferme(Q);
  end;
end;

procedure TOF_CPGESTENG.InitGrid(GS: THGrid);
var i: integer ;
begin
  GS.RowCount  := 2;
  GS.FixedRows := 1;
  GS.ColCount := GS.ColCount+1;
  if not Assigned(GS.OnKeyDown)  then
    GS.OnkeyDown := FormKeyDown;
  if not Assigned(GS.OnMouseUp) then
    GS.OnMouseUp := GMouseUp;
                                        
  for i:=GS.ColCount-1 downto 1 do        {Décale les intitulés des colonnes à cause de la colonne ColLigneTOB}
    begin
    GS.ColWidths[i] := GS.ColWidths[i-1];
    GS.Cells[i, 0]  := GS.Cells[i-1, 0];
    end;
  GS.ColWidths[ColLigneTOB]   := -1;
  GS.ColAligns[ColDebit]     := taRightJustify;
  GS.ColAligns[ColCredit]    := taRightJustify;
  if ((GS as THGrid).Name = GCpta.Name) then begin
    GS.ColAligns[ColMtHTCpta]  := taRightJustify;
    GS.ColAligns[ColSoldeCpta] := taRightJustify;
    if (not isEngagementHT) then
      GS.ColWidths[ColMtHTCpta] := -1
    else
      GS.ColWidths[ColMtHTCpta] := GS.ColWidths[ColSoldeCpta];
    end
  else begin
    GS.ColAligns[ColMtEngInit] := taRightJustify;
    GS.ColAligns[ColStatutEng] := taLeftJustify;
    end;

  {D: Colonne de type date
   R: Colonne de type montant}
  GS.ColTypes[ColDateCpta]  := 'D';
  GS.ColTypes[ColPiece]     := 'R';
  GS.ColTypes[ColDebit]     := 'R';
  GS.ColTypes[ColCredit]    := 'R';
  if GS = GEnga then
    begin
    GS.ColTypes[ColMtEngInit] := 'R';
    end
  else if GS = GCpta then
    begin
    GS.ColTypes[ColMtHTCpta]  := 'R';
    GS.ColTypes[ColSoldeCpta] := 'R';
    end;

  GS.ColNames[ColDateCpta]  := 'E_DATECOMPTABLE';
  GS.ColNames[ColJournal]   := 'E_JOURNAL';
  GS.ColNames[ColPiece]     := 'E_NUMEROPIECE';
  GS.ColNames[ColRefCde]    := 'E_REFINTERNE';
  GS.ColNames[ColLibelle]   := 'E_LIBELLE';
  GS.ColNames[ColDebit]     := 'E_DEBIT';
  GS.ColNames[ColCredit]    := 'E_CREDIT';
  if GS = GEnga then
    begin
    GS.ColNames[ColMtEngInit] := 'CEN_MONTANTENGINI';
    GS.ColNames[ColStatutEng] := 'CEN_STATUTENG';
    end
  else if GS = GCpta then
    begin
    GS.ColNames[ColMtHTCpta]  := 'MONTANTHT';
    GS.ColNames[ColSoldeCpta] := 'SOLDECPTA';
    end;
end;

procedure TOF_CPGESTENG.OnKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
  case Key of
    VK_F9 : BChercher.Click; // F9
  end ;
end;

procedure TOF_CPGESTENG.FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
var G:THGrid; OkG,Vide : Boolean ;
begin
  G    := THGrid(Sender);
  OkG  := G.Focused ;
  Vide := (Shift=[]) ;
  Case Key of
    VK_SPACE  : if ((OkG) and (Vide)) then ;
    //Ctrl + H
    70: if Shift = [ssCtrl] then begin
          end ;
    end;
end;

procedure TOF_CPGESTENG.OnFormResize(Sender: TObject);
begin
  pcpta.Height := pgene.ClientHeight div 2;
end;

procedure TOF_CPGESTENG.GMouseUp(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
var G: THGrid; C,R : Longint ;
begin
  G:=THGrid(Sender);
  if ((ssCtrl in Shift) and (Button=mbLeft)) then begin
     G.MouseToCell(X,Y,C,R) ;
     if R>0 then ;
     end;
end;

function TOF_CPGESTENG.AfficheMsg(num: integer; Av, Ap: string): Word;
begin
  Result := mrNone ;
  if LMsg=nil then exit ;
  if (Num<0) or (Num>LMsg.Mess.Count-1) then Num:=0 ; //erreur inconnue
  Result := LMsg.Execute(num,Av,Ap) ;
end;

procedure TOF_CPGESTENG.InitMsg;
begin
LMsg := THMsgBox.create(FMenuG) ;
{00}LMsg.Mess.Add(traduirememoire('0;Attention;Vous devez renseigner le compte;W;O;O;O'));
{01}LMsg.Mess.Add(traduirememoire('1;Cet engagement est déjà terminé.;E;O;O;O'));

{03}LMsg.Mess.Add(traduirememoire('2;Terminer un engagement;Confirmez-vous ce traitement?;Q;YN;N;N'));
{03}LMsg.Mess.Add(traduirememoire('3;Dé-rapprochement;Confirmez-vous ce traitement?;Q;YN;N;N'));
{04}LMsg.Mess.Add(traduirememoire('4;Echec dans le traitement dé-rapprochement.;W;O;O;O'));
{05}LMsg.Mess.Add(traduirememoire('5;Il ne faut pas sélectionner un engagement terminé.;E;O;O;O')); //dispos
{06}LMsg.Mess.Add(traduirememoire('6;Rapprochement manuel;Confirmez-vous ce traitement?;Q;YN;N;N'));
{07}LMsg.Mess.Add(traduirememoire('7;Transformation en facture;Confirmez-vous ce traitement?;Q;YN;N;N'));

{08}LMsg.Mess.Add(traduirememoire('8; .. ;Confirmez-vous ce traitement?;Q;YN;N;N'));//dispos
{09}LMsg.Mess.Add(traduirememoire('9; .. ;Confirmez-vous ce traitement?;Q;YN;N;N'));//dispos
{10}LMsg.Mess.Add(traduirememoire('10; .. ;Confirmez-vous ce traitement?;Q;YN;N;N'));//dispos
{11}LMsg.Mess.Add(traduirememoire('11; .. ;Confirmez-vous ce traitement?;Q;YN;N;N'));//dispos
{12}LMsg.Mess.Add(traduirememoire('12; .. ;Confirmez-vous ce traitement?;Q;YN;N;N'));//dispos

{13}LMsg.Mess.Add(traduirememoire('03;Gestion des engagements;Pour imprimer les écritures engagements : (Oui)'#13 +
                                                             ' Pour imprimer les écritures comptables  : (Non)'#13 +
                                                             ' Abandonner :(Annuler) ;Q;YNC;Y;C'));

end;

procedure TOF_CPGESTENG.OnClickBChercher(Sender: TObject);
begin
  if Not IsCompteOk then begin
    AfficheMsg(00,'','');
    Exit;
    end;
  FillGridsEngaEtCpta;
end;

procedure TOF_CPGESTENG.FillGrid(G: THGrid; TobEcr: Tob);
  function GetTobValue(TE: TOB; Col: Integer): Variant;
  begin
    Result := TE.GetValue(G.ColNames[Col]);
  end;
var
  i, j:             Integer ;
  Deci:             Integer;
  TE:               TOB ;
  mtD, mtC :        Double;
  SoldeCpta:        Double;
begin
  Deci := V_PGI.OkDecV;
  SoldeCpta := 0;
  G.RowCount  := 2;
  G.ClearSelected; //Désactive les lignes selected. (Gs.AllSelected)


  for i:=0 to G.ColCount-1 do
    G.Cells[i, 1] := '';

  J := 1;
  if TobEcr.Detail.Count<>0 then
    for i:=0 to TobEcr.Detail.Count-1 do begin
      TE := TobEcr.Detail[i] ;
      G.RowCount := J + 1 ;
      G.Cells[ColLigneTOB, J] := IntToStr(i);
      G.Cells[ColDateCpta, J] := GetTobValue(TE, ColDateCpta);
      G.Cells[ColJournal, J]  := GetTobValue(TE, ColJournal);
      G.Cells[ColPiece, J]    := GetTobValue(TE, ColPiece);
      G.Cells[ColRefCde, J]   := GetTobValue(TE, ColRefCde);
      G.Cells[ColLibelle, J]  := GetTobValue(TE, ColLibelle);
      mtD := GetTobValue(TE, ColDebit);
      mtC := GetTobValue(TE, ColCredit);
      G.Cells[ColDebit, J]    := StrS(mtD, Deci);
      G.Cells[ColCredit, J]   := StrS(mtC, Deci);
      FormatMontant(G, ColDebit, J, Deci);
      FormatMontant(G, ColCredit, J, Deci);

      if (G = GEnga) then begin // Ecrit Engagement
        G.Cells[ColMtEngInit, J]   := StrS(GetTobValue(TE, ColMtEngInit), Deci);
        FormatMontant(G, ColMtEngInit, J, Deci);
        G.Cells[ColStatutEng, J]   := GetStatutEng(i, TobEcr);
        end
      else begin // Ecrit Compta.
        if isEngagementHT then
          begin
          G.Cells[ColMtHTCpta, J]  := StrS(GetTobValue(TE, ColMtHTCpta), Deci);
          FormatMontant(G, ColMtHTCpta, J, Deci);
          end;
        SoldeCpta := GetTobValue(TE, ColSoldeCpta);
        G.Cells[ColSoldeCpta, J] := StrS(SoldeCpta, Deci);
        FormatMontant(G, ColSoldeCpta, J, Deci);
        end;

      if (G = GEnga) and (GetStatutEng(i, TobEcr)  in ['E', 'P']) then
        TotSoldeEng  := TotSoldeEng  +  (mtD - mtC)
      else
      if (G = GCpta) then
        TotSoldeCpta := TotSoldeCpta + SoldeCpta;
      inc(J) ;
      end;
end;

function TOF_CPGESTENG.isEngagementHT: boolean;
begin
//  Result := False;
  Result := EngagementHT;
  // Result := (PlanCompta.ItemIndex in [0, 1]) and VH^.EngementHT;   si C ou F et HT
end;

function TOF_CPGESTENG.IsCompteOk: boolean;
begin
  Result := GetControlText('LeCompte') <> '';
end;

function TOF_CPGESTENG.IsCompteGene: boolean;
begin
  Result := (PlanCompta.value = 'G');
end;

function TOF_CPGESTENG.IsRapproManuelAutorise: boolean;
var Statut : char;  SoldeF: double;
begin
  if (not Assigned(TobEcrEng)) or (not Assigned(TobEcrCpta)) then
    begin
    Result := False;
    Exit;
    end;

  Result := (*not IsCompteGene and *) (TobEcrEng.Detail.Count>0) and (TobEcrCpta.Detail.Count>0) and
           (GEnga.nbSelected = 1) and (GCpta.nbSelected = 1);
  if Result then begin
    Statut := GetStatutEng(GetRowSelect(GEnga), TobEcrEng);
    SoldeF := GetSoldeFact(GetRowSelect(GCpta), TobEcrCpta);
    Result := Result and  (Statut in ['E', 'P']) and (SoldeF > 0);
    end;
end;

function TOF_CPGESTENG.IsTerminerUnEngagementAutorise: boolean;
begin
  if (not Assigned(TobEcrEng)) or (not Assigned(TobEcrCpta)) then
    begin
    Result := False;
    Exit;
    end;


  Result := (TobEcrEng.Detail.Count>0) and (GEnga.nbSelected = 1) and (GCpta.nbSelected = 0);
  if Result then
    Result := (GetStatutEng(GetRowSelect(GEnga), TobEcrEng) in ['E', 'P']);
//  if not Result then AfficheMsg(01,'','');
end;

(*FP 02/05/2006
function TOF_CPGESTENG.IsTransformerEnFactAutorise: boolean;
var
  Statut:          char;
  i, J:            Integer;
  ExistEngTermine: Boolean;
begin
  Result := not IsCompteGene and (TobEcrEng.Detail.Count>0) and (GEnga.nbSelected >= 1)
            and (GCpta.nbSelected = 0);
  J := 0; ExistEngTermine := False;
  if Result then
  for i := 0 to (GEnga.nbSelected-1) do begin
    Statut := GetStatutEng(GetRowSelect(GEnga), TobEcrEng);
    if (Statut in ['E', 'P']) then
      inc(J)
    else
      ExistEngTermine := True;
    end;
  Result := (J > 0) and not ExistEngTermine;
  if ExistEngTermine then
    AfficheMsg(05,'',''); // ne pas selectionner un eng terminé
end;*)
function TOF_CPGESTENG.IsTransFactAutorise: boolean;
var
  Statut: char;
  i, j:   Integer;
begin
  {bv FP 02/05/2006}
  if EngagementHT or (not Assigned(TobEcrEng)) or (not Assigned(TobEcrCpta)) then
    begin
    Result := False;
    Exit;
    end;

  Result := not IsCompteGene and (TobEcrEng.Detail.Count>0) and (GEnga.nbSelected >= 1);
  if not Result then
    Exit;

  Result := False;
  if GEnga.AllSelected then
    begin
    for i:=0 to TobEcrEng.Detail.Count-1 do
      begin
      Statut := GetStatutEng(i, TobEcrEng);
      if (Statut in ['E', 'P']) then
        begin
        MarqueTob(TobEcrEng.Detail[i], True);
        Result := True;
        end
      else
        begin
        MarqueTob(TobEcrEng.Detail[i], False);
        end;
      end;
    end
  else
    begin
    for j:=0 to GEnga.NbSelected-1 do
      begin
      GEnga.GotoLeBookMark(j);
      i := GridRowToTobIndex(GEnga, GEnga.Row);
      Statut := GetStatutEng(i, TobEcrEng);
      if (Statut in ['E', 'P']) then
        begin
        MarqueTob(TobEcrEng.Detail[i], True);
        Result := True;
        end
      else
        begin
        MarqueTob(TobEcrEng.Detail[i], False);
        end;
      end ;
    end;
  {ev FP 02/05/2006}
end;

procedure TOF_CPGESTENG.MarqueTob(LstEcr: TOB; Marque: Boolean);
begin
  {bv FP 02/05/2006}
  if not LstEcr.FieldExists('MARQUE') then      
    LstEcr.AddChampSup('MARQUE', True);
  LstEcr.PutValue('MARQUE', BoolToStr_(Marque));
  {ev FP 02/05/2006}
end;

procedure TOF_CPGESTENG.CalculMtHTEcritCpta;

  function MtHTdeLaPiece(TobE: Tob): Double;
  var
    Q:        TQuery ;
    St:       String ;
  begin
    //St := ' SELECT SUM(CASE WHEN E_DEBIT <>0 THEN E_DEBIT ELSE E_CREDIT END) MONTANTHT ' +
    St := ' SELECT SUM(E_DEBIT) DEBIT_HT , SUM(E_CREDIT) CREDIT_HT' +
          ' FROM ECRITURE ' +
          ' WHERE '+WhereEcritureTOB(tsGene, TobE, False) +
          ' AND E_TYPEMVT="HT"';
    try
      Q := OpenSql(St, True);
      if not Q.Eof then begin
        Result := Q.FindField('DEBIT_HT').AsFloat;
        if Result=0 then
        Result := Q.FindField('CREDIT_HT').AsFloat
        end
      else
         Result := 0;
    finally
      Ferme(Q);
      end;
  end;

var
  i:      Integer ;
  TE:     TOB ;
  MtHt:   double;
  SoldeCpta: double;
begin{CalculMtHTEcritCpta}
  MtHt := 0;
  if TobEcrCpta.Detail.Count<>0 then
    for i:=0 to TobEcrCpta.Detail.Count-1 do begin
      TE := TobEcrCpta.Detail[i];
      if isEngagementHT then
        begin
        MtHt := MtHTdeLaPiece(TE);
        SoldeCpta := MtHt - TE.GetDouble('CEN_MONTANTRAP');
        end
      else
        begin
        SoldeCpta := TE.GetValue('E_DEBIT') - TE.GetValue('E_CREDIT') - TE.GetValue('CEN_MONTANTRAP');
        end;
      TE.AddChampSupValeur('MONTANTHT', MtHt) ;
      TE.AddChampSupValeur('SOLDECPTA', SoldeCpta);
      end;
end;

procedure TOF_CPGESTENG.AffListeEngagementsLies(Sender: TObject);
begin
  if LstPiecesEngOuFac.Count>0 then begin
    VisuPiecesGenere(LstPiecesEngOuFac, EcrGen, IndMess);
    end ;
end;

{Recherche les engagements "liés" à la Facture selectionnée}
function TOF_CPGESTENG.IsExistEngagementsLies: boolean;  {clic droit dans GCpta}
var
  lEcr:     TOB ;
  Text:     string;
  lQ:       TQuery;
begin
  VideListe(LstPiecesEngOuFac);
  VideListe(LstPiecesEng);
  BVoirEngLies.Enabled := False;
  IndMess := 16;       // utilisé dans VisuPiecesGenere

  lEcr  := TobEcrCpta.Detail[GridRowToTobIndex(GCpta, GCpta.Row)];  // la facture courante
  Result := (lEcr <> Nil) and (*(not IsCompteGene) and*)
           (TobEcrEng.Detail.Count>0) and (TobEcrCpta.Detail.Count>0) and (GCpta.Row >= 1);
  if not Result then  Exit;
  //récuprère les Eng. liées (=rappro.) pour affichage et le dé-rapprochement.
  Text := ClauseSelectEngRapproche(lEcr.GetValue('E_JOURNAL'),
                                   IntToStr(lEcr.GetValue('E_NUMEROPIECE')),
                                   lEcr.GetValue('E_QUALIFPIECE'),
                                   lEcr.GetValue('E_DATECOMPTABLE'),
                                   lEcr.GetValue('E_EXERCICE'));
  try
    lQ := OpenSQL(Text, True);
    Result := not lQ.Eof;
    if Result then begin
      ChargeLstPiecesEngOuFac(lQ);
      lQ.First;
      ChargeLstPiecesEngPourDRapp(lQ);
      {GCpta.FlipSelection(GCpta.Row); permet de sélectionner la ligne mais il faut le déselect. à la main}
      end;
  finally
    Ferme(lQ);
    end;
end;

{Recherche le(s) engagement(s) rapproché(s) avec la facture }
function TOF_CPGESTENG.IsExistFacturesLiees: boolean;   {clic droit dans GEnga}
var
  lEcr:     TOB ;
  Text:     String;
  lQ:       TQuery;
begin
  VideListe(LstPiecesEngOuFac);
  VideListe(LstPiecesEng);
  BVoirEngLies.Enabled := False;
  IndMess := 17;      // utilisé dans VisuPiecesGenere

  lEcr  := TobEcrEng.Detail[GridRowToTobIndex(GEnga, GEnga.Row)];
  Result := (lEcr <> Nil) and (*(not IsCompteGene) and*)
           (TobEcrEng.Detail.Count>0) and (TobEcrCpta.Detail.Count>0) and (GEnga.Row >= 1) and
           (lEcr.GetValue('CEN_FJOURNAL') <> '');
  if not Result then  Exit;

  //récuprère les Fact liées (=rappro.) pour affichage.
  Text := ClauseSelectFacRapprochee(lEcr.GetValue('CEN_FJOURNAL'),
                                   IntToStr(lEcr.GetValue('CEN_FNUMEROPIECE')),
                                   lEcr.GetValue('CEN_FQUALIFPIECE'),
                                   lEcr.GetValue('CEN_FDATECOMPTA'),
                                   lEcr.GetValue('CEN_FEXERCICE'));
  try
    lQ := OpenSQL(Text, True);
    Result := not lQ.Eof;
    if Result then begin
      ChargeLstPiecesEngOuFac(lQ);
      {GEnga.FlipSelection(GEnga.Row);}
      end;
  finally
    Ferme(lQ);
    end;

  //récuprère les Eng. rapprochés pour le dé-rapprochement.
  Text := ClauseSelectEngRapproche(lEcr.GetValue('CEN_FJOURNAL'),
                                   IntToStr(lEcr.GetValue('CEN_FNUMEROPIECE')),
                                   lEcr.GetValue('CEN_FQUALIFPIECE'),
                                   lEcr.GetValue('CEN_FDATECOMPTA'),
                                   lEcr.GetValue('CEN_FEXERCICE'));
  try
    lQ := OpenSQL(Text, True);
    Result := not lQ.Eof;
    if Result then ChargeLstPiecesEngPourDRapp(lQ);
  finally
    Ferme(lQ);
    end;
end;

function TOF_CPGESTENG.ClauseSelectEngRapproche(Jou, Pie, QualifPiece: string; DateComptable: TDateTime; Exe: string): string;
//on prend la ligne où le compte (Tiers) est présent.
begin
  Result := ' SELECT * FROM ECRITURE ' + ClauseJoinEng(False);
  Result := Result +
                ' Where ' + ClauseOngletCompte + 'AND' +
                ' CEN_FJOURNAL ="'   + Jou + '" AND ' +
                ' CEN_FNUMEROPIECE=' + Pie + ' AND ' +
                ' CEN_FQUALIFPIECE="'   + QualifPiece + '" AND ' +
                ' CEN_FDATECOMPTA="' + UsDateTime(DateComptable) + '" AND ' +
                ' CEN_FEXERCICE="'   + Exe + '" ';
  Result := Result +
          ' Order by E_EXERCICE, E_JOURNAL, E_NUMEROPIECE ';
end;

function TOF_CPGESTENG.ClauseSelectFacRapprochee(Jou, Pie, QualifPiece: string; DateComptable: TDateTime; Exe: string): string;
//on prend la ligne où le compte (Tiers) est présent.
begin
  Result := ' SELECT * FROM ECRITURE ' + ClauseJoinEng(False);
  Result := Result +
                ' Where ' + ClauseOngletCompte + 'AND' +
                ' E_JOURNAL ="'   + Jou + '" AND ' +
                ' E_NUMEROPIECE=' + Pie + ' AND ' +
                ' E_DATECOMPTABLE="' + UsDateTime(DateComptable) + '" AND ' +
                ' E_QUALIFPIECE="'   + QualifPiece + '" AND ' +
                ' E_EXERCICE="'   + Exe + '" ';
  Result := Result +
          ' Order by E_EXERCICE, E_JOURNAL, E_NUMEROPIECE ';
end;

procedure TOF_CPGESTENG.ChargeLstPiecesEngOuFac(Q1: TQuery);
var
  O:     TOBM ;
begin
  while not Q1.EOF do begin
    O := TOBM.Create(EcrGen,'',True) ;
    O.ChargeMvt(Q1) ;
    if (O<>Nil) then
      LstPiecesEngOuFac.Add(O);  // liste des Eng ou Fact pour affichage.
    Q1.Next;
    end;
  BVoirEngLies.Enabled := (LstPiecesEngOuFac.Count>0);
end;

procedure TOF_CPGESTENG.ChargeLstPiecesEngPourDRapp(Q1: TQuery);
var
  M:   RMVT ;
  P:   P_MV ;
begin
  while not Q1.EOF do begin
    M := MvtToIdent(Q1,fbGene,False) ;
    P := P_MV.Create ;
    P.R := M ;
    LstPiecesEng.Add(P); // liste des Eng pour dé-rapprocher
    Q1.Next;
    end;
end;

function TOF_CPGESTENG.GetClauseRechEngOuFac(pos: integer; LstEcr: TOB): string;
var
  lEcr:    TOB ;
begin
  Result := '';
  if (pos < 0) then
    Exit;
  lEcr  := LstEcr.Detail[pos];
  if lEcr = Nil then Exit;
  Result := ' E_JOURNAL="'    + lEcr.GetValue('E_JOURNAL') + '" AND ' +
            ' E_NUMEROPIECE=' + IntToStr(lEcr.GetValue('E_NUMEROPIECE')) + ' AND ' +
            ' E_EXERCICE="'   + lEcr.GetValue('E_EXERCICE') + '" AND ' +
            ' E_QUALIFPIECE="'   + lEcr.GetValue('E_QUALIFPIECE') + '" AND ' +
            ' E_DATECOMPTABLE="' + UsDateTime(lEcr.GetValue('E_DATECOMPTABLE')) + '"';
end;

function TOF_CPGESTENG.GetRowSelect(G: THGrid): Integer;
//attention une seule ligne est sélectionnée
begin
  Result := -1;
  if (G.nbSelected = 1) then begin
    G.GotoLeBookMark(0);
    Result := GridRowToTobIndex(G, G.Row);
    end;
end;

procedure TOF_CPGESTENG.OnClickBImprimer(Sender: TObject);
var
  TitreEnga:  String;
  TitreCpta:  String;
begin
  TitreEnga := TraduireMemoire('Gestion des engagements - Ecritures engagements');
  TitreCpta := TraduireMemoire('Gestion des engagements - Ecritures comptables');

  if FSelectedGrid = nil then
    PGIInfo(TraduireMemoire('Vous devez sélectionner la grille à imprimer'))
  {$IFDEF EAGLCLIENT}
  else if (FSelectedGrid = GEnga) then
    LanceEtatTob('E','CEN','CE1',TobEcrEng, True,false,false,nil,'',TitreEnga,False,0,'',0)
  else if (FSelectedGrid = GCpta ) then
    LanceEtatTob('E','CEN','CE2',TobEcrCpta,True,false,false,nil,'',TitreCpta,False,0,'',0);
  {$ELSE}
  else if (FSelectedGrid = GEnga) then
    PrintDBGrid(GEnga, nil, TitreEnga,'')
  else if (FSelectedGrid = GCpta ) then
    PrintDBGrid(GCpta, nil, TitreCpta, '');
  {$ENDIF}
end;

function TOF_CPGESTENG.GetStatutEng(pos: integer; LstEcr: TOB): char;
var
  Statut : string;
  lEcr:    TOB ;
begin
  Result := ' ';
  if (pos < 0) then
    Exit;
  lEcr  := LstEcr.Detail[pos];
  Statut := lEcr.GetValue('CEN_STATUTENG');
  if Length(Statut)> 0 then
    Result := Statut[1]
end;

function TOF_CPGESTENG.GetSoldeFact(pos: integer; LstEcr: TOB): double;
var
  lEcr:    TOB ;
begin
  Result := 0;
  if (pos < 0) then
    Exit;
  lEcr  := LstEcr.Detail[pos];
  if isEngagementHT then
    Result := lEcr.GetValue('MONTANTHT') - lEcr.GetValue('CEN_MONTANTRAP')
  else
    Result := lEcr.GetValue('E_DEBIT') + lEcr.GetValue('E_CREDIT') - lEcr.GetValue('CEN_MONTANTRAP');
end;

procedure TOF_CPGESTENG.OnClickBVoirEngLies(Sender: TObject);
begin
  AffListeEngagementsLies(Nil);
end;

{b FP 02/05/2006}
function TOF_CPGESTENG.OnBeforeValidePieceComptaSimul(Tof: TOF_CPSAISIEPIECE; Ecran: TForm; Piece: TPieceCompta): Boolean;
var
  i:           Integer;
  TobEcr:      TOB;
  NbPieceEng:  Integer;
  Solde:       Double;
  TotalDebit:  Double;
  TotalCredit: Double;
begin
  Result     := True;
  NbPieceEng := 1;

  {Après modifications, chaque piece d'engagements doit être équilibrée}
  for i:=0 to Piece.Detail.Count-1 do
    begin
    TobEcr := Piece.Detail[i];
    if TobEcr.FieldExists('ENGNUMPIECEORIG') then
      NbPieceEng := TobEcr.GetValue('ENGNUMPIECEORIG')
    else
      TobEcr.AddChampSupValeur('ENGNUMPIECEORIG', NbPieceEng, False);
    end;

  for i:=1 to NbPieceEng do
    begin
    Piece.GetTotauxPourChamps('ENGNUMPIECEORIG', IntToStr(i), TotalDebit, TotalCredit);
    Solde :=  Arrondi(TotalDebit-TotalCredit, Piece.Devise.Decimale);
    if Solde <> 0 then
      begin
      Ecran.ModalResult := mrNone;
      Result := False;
      break;
      end;
    end;
end;

function TOF_CPGESTENG.OnBeforeSuppLigne(Tof: TOF_CPSAISIEPIECE; Piece: TPieceCompta; NumLigne: Integer): Boolean;
begin
  {il faut vérifier:
      Interdiction de saisir un nouveau tiers
      Interdiction de modifier le collectif et le tiers d'une ligne TTC
      Vérifier le sens des montants dans la fusion des écritures}
  Result := Trim(Piece.GetValue(NumLigne, 'E_AUXILIAIRE')) = '';  {Impossible de supprimer une écriture de tiers}
end;
{e FP 02/05/2006}


{ TInfoParamSaisieFacture }
{b FP 02/05/2006}
constructor TInfoParamSaisieFacture.Create;
begin
  FTof               := nil;
  FEcran             := nil;
  FPiece             := nil;
  FJOURNALExit       := nil;
  FJOURNALExit       := nil;
  FNATUREPIECEExit   := nil;
  FDATECOMPTABLEExit := nil;
  FDEVISEExit        := nil;
  FETABLISSEMENTExit := nil;
end;

destructor TInfoParamSaisieFacture.Destroy;
begin
  FTof    := nil;
  FEcran  := nil;
  FPiece  := nil;
  inherited;
end;

procedure TInfoParamSaisieFacture.InitComposant;
begin
  {Boutons}
  FTof.BParamListe.Enabled := False;
  FTof.BSolde.Enabled      := False;
  FTof.BEche.Enabled       := False;
  FTof.BVentil.Enabled     := False;
  FTof.BGenereTVA.Enabled  := False;
  FTof.BControlTVA.Enabled := False;
  FTof.BModifTVA.Enabled   := False;
  FTof.BComplement.Enabled := False;
  FTof.BRechercher.Enabled := False;
  FTof.BProrata.Enabled    := False;
  FTof.BMenuTva.Enabled    := False;
  FTof.FListe.Enabled      := False;
  FTof.SetControlEnabled('BMENUZOOM',      False);
  FTof.SetControlEnabled('BMENUMODIFS',    False);
  FTof.SetControlEnabled('BMENULIGNES',    False);

  {Entête}
  FTof.E_JOURNAL.Enabled       := True;
  FTof.E_DEVISE.Enabled        := False;
  FTof.E_DATECOMPTABLE.Enabled := True;
  FTof.E_NATUREPIECE.Enabled   := True;
  FTof.E_ETABLISSEMENT.Enabled := True;
  FTof.BMenuTva.Visible        := False;
end;

procedure TInfoParamSaisieFacture.OnAfterChargePiece(Tof: TOF_CPSAISIEPIECE; Ecran: TForm; Piece: TPieceCompta);
begin
  InitComposant;

  FJOURNALExit       := FTof.E_JOURNAL.OnExit;
  FNATUREPIECEExit   := FTof.E_NATUREPIECE.OnExit;
  FDATECOMPTABLEExit := FTof.E_DATECOMPTABLE.OnExit;
  FETABLISSEMENTExit := FTof.E_ETABLISSEMENT.OnExit;

  FTof.E_JOURNAL.OnExit        := OnJOURNALExit;
  FTof.E_NATUREPIECE.OnExit    := OnNATUREPIECEExit;
  FTof.E_DATECOMPTABLE.OnExit  := OnDATECOMPTABLEExit;
  FTof.E_ETABLISSEMENT.OnExit  := OnETABLISSEMENTExit;
  FTof.SaisiePiece.OnUserCellEnter := nil;
  FTof.SaisiePiece.OnUserRowEnter  := nil;

  if FTof.E_JOURNAL.CanFocus then
    FTof.E_JOURNAL.SetFocus;
end;

procedure TInfoParamSaisieFacture.OnBeforeChargePiece(Tof: TOF_CPSAISIEPIECE; Ecran: TForm; Piece: TPieceCompta; var Action: TActionFiche);
begin
  FTof   := Tof;
  FEcran := Ecran;
  FPiece := Piece;

  Action := taModif;
end;

function TInfoParamSaisieFacture.OnBeforeValidePieceCompta(Tof: TOF_CPSAISIEPIECE; Ecran: TForm; Piece: TPieceCompta): Boolean;
begin
  Result := true;
end;

procedure TInfoParamSaisieFacture.OnDATECOMPTABLEExit(Sender: TObject);
begin
  if Assigned(FDATECOMPTABLEExit) then
    FDATECOMPTABLEExit(Sender);
  InitComposant;
end;

procedure TInfoParamSaisieFacture.OnETABLISSEMENTExit(Sender: TObject);
begin
  if Assigned(FETABLISSEMENTExit) then
    FETABLISSEMENTExit(Sender);
  InitComposant;
end;

procedure TInfoParamSaisieFacture.OnJOURNALExit(Sender: TObject);
begin
  if Assigned(FJOURNALExit) then
    FJOURNALExit(Sender);
  InitComposant;
end;

procedure TInfoParamSaisieFacture.OnNATUREPIECEExit(Sender: TObject);
begin
  if Assigned(FNATUREPIECEExit) then
    FNATUREPIECEExit(Sender);
  InitComposant;
end;
{e FP 02/05/2006}

procedure TOF_CPGESTENG.FillGridsEngaEtCpta;
begin
  FillTobsEngaEtCpta;
  TotSoldeEng  := 0;
  TotSoldeCpta := 0;
  FillGrid(GCpta, TobEcrCpta);  //Remplir les grilles
  FillGrid(GEnga, TobEcrEng);
  AfficheSoldes;
end;

procedure TOF_CPGESTENG.ChangeAffichage(Plus: boolean);
begin
if plus then OngletsHaut.Align:=AlNone else OngletsHaut.Align:=AlTop ;
OngletsHaut.Visible:=not Plus ;
FAgrandirEcran := Plus;
end;

procedure TOF_CPGESTENG.OnAgrandirClick(Sender: TObject);
begin
  ChangeAffichage(not FAgrandirEcran);
end;

procedure TOF_CPGESTENG.OnValiderRapproAuto(Sender: TObject);
var
  EngReste:   TPieceEngagement;
begin
  if not IsRapproManuelAutorise then begin
    PGIError(TraduireMemoire('Le rapprochement n''est pas autorisé.'));
    Exit;
    end;

  EngReste := TPieceEngagement.CreerPieceEng;
  try
    if RapprochementManuel(EngReste) then
      FFiltreEngRapproAuto := EngReste.GetWhereSQL+' AND E_AUXILIAIRE="'+LeCompte.Text+'"';
    BChercher.Click;
    if (TobEcrEng.Detail.Count = 0) or (TobEcrCpta.Detail.Count = 0) then
      Ecran.Close;
  finally
    FreeAndNil(EngReste);
    end;
end;

function TOF_CPGESTENG.RapprochementManuel(EngReste: TPieceEngagement): Boolean;
var
  lEng :        TPieceEngagement;
  lFac :        TPieceEngagement;
  lDate:        TDateTime;
  cleEng, cleFac:  string;
begin
  Result := False;
  cleEng := GetClauseRechEngOuFac(GetRowSelect(GEnga), TobEcrEng);
  cleFac := GetClauseRechEngOuFac(GetRowSelect(GCpta), TobEcrCpta);
  if  (cleEng = '') or (cleFac = '') then
    Exit;

  if (AfficheMsg(06, '', '') <> mrYes) then  Exit; //confirmation rapp manuel.
  lDate := Date;
  lEng := TPieceEngagement.CreerPieceEng;
  lFac := TPieceEngagement.CreerPieceEng;
  BeginTrans;
  try
    lEng.LoadFromSQL(cleEng);
    lFac.LoadFromSQL(cleFac);
    Result := ExtourneEng(lEng, Nil, lDate);
    Result := Result and RapprocherEngFac(lEng, lFac, lFac, lDate);
    Result := Result and (lEng.IsEngagementEstTermine  or { si pas de Eng Reste à générer}
                         ResteEng(lEng, EngReste) );      { ou si l'on arrive à créer le Reste Eng}
    if Result then
      CommitTrans
    else Rollback;
  except
    rollback;
    end;
  lEng.Free;
  lFac.Free;
end;

procedure TOF_CPGESTENG.OnChercherRapproAuto(Sender: TObject);
begin
  FillGridsEngaEtCpta;
  if GEnga.RowCount = 2 then       {Sélectionne l'engagement}
    GEnga.FlipSelection(1);
  if GCpta.RowCount = 2 then       {Si une seule facture}
    GCpta.FlipSelection(1);
end;

function TOF_CPGESTENG.IsDerapproAutorise: Boolean;
begin
  if (not Assigned(TobEcrEng)) or (not Assigned(TobEcrCpta)) then
    begin
    Result := False;
    Exit;
    end;

  Result := False;
  if (FSelectedGrid = GEnga) and (FSelectedGrid.Row>0) then
    Result := GetStatutEng(GridRowToTobIndex(FSelectedGrid, FSelectedGrid.Row), TobEcrEng) in ['P', 'T', 'X']
  else if (FSelectedGrid = GCpta) and (FSelectedGrid.Row>0) then
    Result := TobEcrCpta.Detail[GridRowToTobIndex(FSelectedGrid, FSelectedGrid.Row)].GetValue('CEN_MONTANTRAP') <> 0;
end;

procedure TOF_CPGESTENG.OnGridEnter(Sender: TObject);
begin
  FSelectedGrid := nil;
  if Sender is THGrid then
    FSelectedGrid := (Sender as THGrid);
end;

function TOF_CPGESTENG.GridRowToTobIndex(G: THGrid; Row: Integer): Integer;
begin
  Result := StrToInt(G.Cells[ColLigneTOB, Row]);
end;

procedure TOF_CPGESTENG.OnFindAFindDialog(Sender: TObject);
begin
  if FSelectedGrid <> nil then
    Rechercher(FSelectedGrid, FFindDialog, FFindFirst);
end;

procedure TOF_CPGESTENG.OnRechercherClick(Sender: TObject);
begin
  FFindFirst := True;
  if FSelectedGrid = nil then
    PGIInfo(TraduireMemoire('Vous devez sélectionner la grille des engagements ou des écritures avant de faire une recherche'))
  else
    FFindDialog.Execute;
end;

procedure TOF_CPGESTENG.OnGrilleDblClick(Sender: TObject);
var laGrille : THGrid ;
    lInIdx   : integer ;
    lTobEcr  : Tob ;
    lTobMere : Tob ;
begin

  if Sender = nil then Exit ;
  if not (Sender is THGrid) then Exit ;

  laGrille := THGrid(Sender) ;
  lInIdx  := GridRowToTobIndex( laGrille, laGrille.Row ) ;

  if THGrid(Sender).Name = GEnga.Name
    then lTobMere := TobEcrEng
    else lTobMere := TobEcrCpta ;

  if (lInIdx < 0) or (lInIdx >= lTobMere.Detail.count) then Exit ;

  lTobEcr := lTobMere.Detail[lInIdx] ;
  SaisiePieceCompta( lTobEcr.GetString('E_QUALIFPIECE'),
                     taConsult,
                     lTobEcr.GetString('E_JOURNAL'),
                     lTobEcr.GetDateTime('E_DATECOMPTABLE'),
                     lTobEcr.GetInteger('E_NUMEROPIECE') ) ;

end;

procedure TOF_CPGESTENG.PlanComptaChange(Sender: TObject);
begin

  if (csDestroying in LeCompte.ComponentState) then Exit;
  if ObjFiltre.InChargement then Exit ;

  LeCompte.Text := '' ;
  ChargeLeCompte(False);

end;

Initialization
  registerclasses ( [ TOF_CPGESTENG ] ) ;
end.

