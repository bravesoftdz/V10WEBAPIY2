unit TofRPoin;

interface
 
uses
    Windows,
    Controls,
    StdCtrls,
    ExtCtrls,
    Graphics,
    forms,
    Grids,
    Classes,
    ComCtrls,
    sysutils,
    Spin,
    menus,
    vierge,
    Paramsoc,
{$IFDEF EAGLCLIENT}
    MaineAGL,
    UtileAgl, {JP 17/08/05 : FQ 15433 : LanceEtatTob}
  {$IFNDEF TRESO}
    CPCODEAFB_TOF,
  {$ENDIF}
{$ELSE}
    FE_Main,
    EdtREtat, {JP 17/08/05 : FQ 15433 : LanceEtatTob}
    DB,
    Hdb,
    {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
    PrintDBG,
  {$IFNDEF TRESO}
    CODEAFB,
  {$ENDIF}
    Extraibq,
{$ENDIF}
    Dialogs,
    HCtrls,
    HEnt1,
    LettUtil,
    HMsgBox,
    LettAuto,
    UTOF,
    UTOB,
    UtilDiv,
    ent1,
    Saisutil,
  {$IFDEF TRESO}
   TomCIB,
   Constantes,
  {$ELSE}
    Saisie,
    SaisBor,
  {$ENDIF}
    SaisComm,
    HTB97,
    CPTESAV      // RecalculTotPointeNew1
    {$IFNDEF TRESO}
    ,CPGeneraux_TOM
    {$ENDIF}
    , UObjFiltres {JP 13/08/04 : gestion des filtres}
    ;

procedure CP_LancePointageAuto;

Type
    TOF_PointeReleve = class (TOF)
    private
        // Controles de la fiche AGL :
        OngletHaut   :TPageControl ;
        GC           :THGrid ;
        GB           :THGrid ;
        CBanque      :THValCombobox ;
        cRefReleve   :THValCombobox ;
        cDevReleve   :THValCombobox ;
        cDevise      :THValCombobox ;
        cAffDev      :THValCombobox ;
        MultiAFB     :ThMultiValComboBox ;
        MultiModepaie:ThMultiValComboBox ;
        CDate        :TCheckBox ;
        CNumCheque   :TCheckBox ;
        CCombinatoire:TCheckBox ;
        CModeReg     :TCheckBox ;
        ModeTiers    :TCheckBox ;
        BAgrandir    :TButton ;
        BReduire     :TButton ;
        BChercher    :TButton ;
        BPointageAuto:TButton ;
        BValider     :TButton ;
        BParam       :TButton ;
        BZoom        :TButton ;
        BTnAFB       :TToolBarButton97 ;
        BRechercher  :TToolBarButton97 ;
        BImprimer    :TToolBarButton97 ;
        PopZoom      :TPopupMenu;
        vpiece, vcompte   :TMenuItem;
        TTotalC      :THEdit ;
        TTotalB      :THEdit ;
        TEcart       :THEdit ;
        TDate        :TSpinEdit ;
        DuCarNumCheque   :TSpinEdit ;
        AuCarNumCheque   :TSpinEdit ;
        TNiveau      :TSpinEdit ;
        DateDe       :THEdit;
        DateJusqua   :THEdit;
        comboRef       :THValCombobox;
        BHelp: TToolbarButton97;

        pcpta, pbanq, pgene : TPanel;
        {$IFDEF TRESO}
        CRegAccro      : TCheckBox ;
        ReferenceEcrit : string; {Colonne de la compta servant de référence pièce}
        ToleranceDate  : Integer; {Nombre de jours de tolérance choisi pour les dates}
        {JP 28/11/05 : FQ TRESO 10315
         TblRegles : array [1..99] of string; {Tableau des numéros de règles pour chaque code AFB/CIB, 0 = pas de règle}
        TobRegles : TOB;
        TobCompta : TOB;
        {$ENDIF}

        AFindDialog   : TFindDialog; // Fenêtre de recherche sur les listes (Ajout SBO : fiche 12033)


        //Variables de l'objet
        A_DEBIT,
        A_CREDIT,
        C_DEBIT,
        C_CREDIT,
        RefReleve : string ;
        TobReleve,
        TobEcriture : TOB;
        TobCodeAFB : TOB ;
        TobDevise : Tob ;
        MttC,
        MttB : Double;
        NbDecDev,NbCoche, NbCocheC, NbCocheB : integer ;
        PointageAuto,
        PointageManu,
        ChangeBqeOk,
        PointeAvecListeNonComplete,
        CDevReleveOk : boolean;
        LMsg : THMsgBox ;
        nk,DecCompta,DecReleve : Integer ;
        PasDePointagePossible : Boolean ;
        LDF, LDP, DeviseCompta, DeviseReleve : String ;
        gdDateReleve : TDateTime ;        // Date utilisée pour maj Date pointage
        FFindFirst: Boolean;  // Utilisé pour la recherche sur les listes  (Ajout SBO : fiche 12033)
        listeRech : String ;  // Nom de la liste sélectionnée pour la recherche (Ajout SBO fiche 12033)
        //Méthodes de l'objet
        //Méthodes liées a la fiche et aux controles de l'objet
        function RecupeControles: boolean ;
        function AssignControles: boolean ;
        function InitiaControles: boolean ;
        {$IFDEF TRESO}
        procedure CRegAccroOnClick(Sender: TObject);
        procedure CMODEREGOnClick (Sender: TObject);
        {JP 10/03/04 : Génération des écritures de copta et de tréso}
        procedure GenererEcriture (Sender: TObject);
        {$ENDIF}
        procedure OnChangeCBanque (Sender: TObject) ;
        procedure OnDblClickGCompta(Sender: TObject) ;
        procedure OnDblClickGBanque(Sender: TObject) ;
        procedure OnClickBPointageAuto(Sender: TObject) ;
        procedure OnClickBChercher(Sender: TObject) ;
        procedure OnClickCDate(Sender: TObject) ;
        procedure OnClickCNumCheque(Sender:  TObject) ;
        procedure OnClickCCombinatoire(Sender:  TObject) ;
        procedure OnClickBAgrandir(Sender:  TObject) ;
        procedure OnClickBReduire(Sender:  TObject) ;
        procedure OnClickBParam(Sender:  TObject) ;
        procedure OnDateExit(Sender:  TObject) ;
        procedure OnKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
        procedure OnPopUpZoom(Sender: TObject);
        procedure OnClickVpiece(Sender:  TObject) ;
        procedure OnClickVcompte(Sender:  TObject) ;
        procedure OnFormResize(Sender:  TObject) ;
        procedure OnClickBImprimer(Sender: TObject);
        procedure BRechercherClick(Sender: TObject);     // Recherche sur les liste  (Modif SBO : fiche 12033)
        procedure OnFindAFindDialog  (Sender : TObject); // Recherche sur les liste  (Ajout SBO : fiche 12033)

        procedure GetCellCanvasB(Acol,ARow : LongInt ; Canvas : TCanvas; AState: TGridDrawState) ;
        procedure GetCellCanvasC(Acol,ARow : LongInt ; Canvas : TCanvas; AState: TGridDrawState) ;
        procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState) ;
        procedure GMouseUp(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer) ;
        procedure InitGrid(GS : THGrid ; Okel : Boolean = FALSE) ;
        procedure InitVariables ;
        procedure RemetToutAZero ;
        procedure RemetRefPointageABlanc;
        function  ExisteMoyenPaiement: boolean;

        //Méthodes liées aux messages de la fiche
        procedure InitMsg ;
        function AfficheMsg(num : integer;Av,Ap : string ) : Word ;

        //Méthodes d'affichages des valeurs dans la fiche AGL
        procedure RemplirValRelev(Banque: string) ;
        procedure RemplirValDevis(Banque: string) ;
        function  PutMontant(Montant : Double; Dec : Integer) : string ;
        procedure AfficheMontant(TTotal : THEdit;Mtt : Double) ;
        procedure LesVisibles(Visible : Boolean) ;
        procedure ChangeAffichage(Plus: boolean) ;
        procedure ChangeAffichageGrid(ARow : LongInt ; GS : THGrid ) ;
        procedure RempliCompta ;
        procedure RempliReleve ;

        //Méthodes liées aux pointages
        //Methodes liées aux traitements du pointage
        procedure LanceLePointageAuto ;
        function  ConstruitListe(var LM : T_D;var LP,LB : T_I; MM:Double) : integer;
        procedure RempliLesTobs ;
        {$IFDEF TRESO}
        procedure ChargerTobs;
        {$ELSE}
        Procedure ChargeTobCodeAFB ;
        {$ENDIF}
        procedure PointeManu(Sender: TObject) ;
        //Methodes liées a l'affichage du pointage
        procedure CocheDecocheAuto(Ligne : integer) ;
        procedure CocheDecoche(G:THGrid) ;
        procedure CocheDecocheGB ;
        procedure CocheDecocheGC ;

        procedure BHelpClick(Sender: TObject);
        procedure BtnAFBOnClick(Sender: TObject);

        //Diverses fonctions
        function  EcritureSelectionnee : boolean ;
        procedure PositionneDevise(Tout: boolean) ;
        function  GetLeModeRglt(CodeAfb: string): string;
        function  RechercheDevise(Dev: String): String;
        procedure ChargeTobDevise;
        function  RechercheDecimale(Dev: String): Integer;
        procedure OuvrirRecherche ; // Recherche sur les listes (Ajout SBO : fiche 12033)
        {JP 13/08/04 : FQ 13167, problème de la gestion du confidentiel}
        function GetPlus : string;
        {JP 17/08/05 : FQ COMPTA 15433 : Impression en eAgl}
        procedure ImprimerGrille(aGrille : THGrid; aTitre : string);
        {GP 16/03/2008 : Rappro troc de lille : choix de la ref eexbqlig}
        Function  OkCBRefCEL : Boolean ;
        procedure CBPARAM2OnClick(Sender:  TObject) ;
        procedure CBOKPOSBQOnClick(Sender:  TObject) ;
        procedure ActiveStop(Ok : Boolean) ;
        procedure bStopClick(Sender : Tobject) ;
        procedure AfficheCR(St,St1 : String) ;
        procedure Reinit(St : String) ;
        procedure Affecte(Nom,St : String) ;
        procedure BVISULIBClick(Sender: TObject); //Gv 10/06/02
        procedure AffecteZone(NomHG : String ; Ou : Integer = -1) ;
        Function  OkTTB : Boolean ;
        procedure GCRowEnter(Sender: TObject; Ou: Longint; var Cancel: Boolean; Chg: Boolean);
        procedure GBRowEnter(Sender: TObject; Ou: Longint; var Cancel: Boolean; Chg: Boolean);
        procedure BRENVOIECRClick(Sender: TObject); //Gv 10/06/02
        procedure BRENVOIBQClick(Sender: TObject); //Gv 10/06/02
    public
        ObjFiltre  : TObjFiltre;
        procedure OnArgument(S: string); override;
        procedure OnLoad ; override ;
        procedure OnCancel ; override ;
        procedure OnUpdate ; override ;
        procedure OnClose ; override ;
    end;

const
    SP_DATE        = 0 ;
    SP_REFPOINTAGE = SP_DATE+1 ;
    SP_REFERENCE   = SP_REFPOINTAGE+1 ;
    SP_SOLDE       = SP_REFERENCE+1 ;
    SP_RGLT        = SP_SOLDE+1 ;
    SP_DEVAFF      = SP_RGLT+1 ;
    SP_LIBELLE     = SP_DEVAFF+1 ;
    SP_DATEVAL     = SP_LIBELLE+1 ;
    SP_ROW         = SP_DATEVAL+1 ;
    SP_POSTOB      = SP_ROW+1 ;
    SP_DEVISE      = SP_POSTOB+1 ;
    SP_POINTER     = SP_DEVISE+1 ;
    SP_NUMPOINTER  = SP_POINTER+1 ;
    SP_CLASSER     = SP_NUMPOINTER+1 ;
    {$IFDEF TRESO}
    {27/04/05 : FQ 10246 : ajout du libellé enrichi, en version trésorerie. Valable à partir de la 6.50.001.001}
    SP_LIBENRICHI  = SP_CLASSER + 1;
    {$ENDIF TRESO}

    CARPOINTER     = '+' ;
    ISO_EURO       = 'EUR' ;
    ISO_FRANC      = 'FRF' ;

implementation

uses
    {$IFDEF eAGLCLIENT}
    MenuOLX
    {$ELSE}
    MenuOLG
    {$ENDIF eAGLCLIENT}
    {$IFDEF TRESO}
    , TRSAISIEFLUX_TOF, Commun, HQry, UProcEcriture, UProcCommission
    {$ELSE}
    ,UTofAncetre
    {$ENDIF}
    ,Math
    ,hPanel ;

procedure CP_LancePointageAuto;
begin
  if GetParamSocSecur('SO_POINTAGEJAL', False) then
  begin // 14532
    PGIInfo('Fonction non disponible en mode pointage sur journal.#10#13Utilisez la fonction du changement de mode de pointage.', 'Pointage automatique');
    exit;
  end;

  // GCO - 14/10/2004 - FQ 14804
  if CEstPointageEnConsultationSurDossier then
  begin
    PgiInfo('Vous avez indiqué une liaison avec une comptabilité ' +
            RechDom('CPLIENCOMPTABILITE', GetParamSocSecur('SO_CPLIENGAMME', ''), False) +
            ' et la gestion du pointage ' + #10 +
            'est effectuée ' + RechDom('CPPOINTAGESX', GetParamSocSecur('SO_CPPOINTAGESX', 'CLI'), False) + '. ' +
            'Vous n''avez pas accès à cette commande.', 'Pointage Automatique');
    Exit;
  end;

  {$IFDEF TRESO}
    AGLLanceFiche('TR', 'RLVPOINTAGE', '', '', '') ;
  {$ELSE}
    AGLLanceFiche('CP', 'RLVPOINTAGE', '', '', '') ;
  {$ENDIF}
end;

(*******************************************************************************



    Méthodes publiques de la TOF



*******************************************************************************)

procedure TOF_PointeReleve.OnLoad;
begin
    inherited ;
    TFVierge(Ecran).HelpContext := 7774000 ;
    ChargeTobdevise ;
    {$IFNDEF TRESO}
    ChargeTobCodeAFB ;
    {$ENDIF}
    LDF := V_PGI.DeviseFongible ;
    LDP := V_PGI.DevisePivot ;
    // Recupération des controles de la fiche
    if not RecupeControles then exit ;
    if not AssignControles then exit ;
    if not InitiaControles then exit ;
    OnChangeCBanque(nil) ;
    InitMsg ;
    {JP 13/08/2004 : Nouvelle gestion des filtres}
    ObjFiltre.Charger;
  {$IFDEF TRESO}
  CRegAccroOnClick(CRegAccro);
  CMODEREGOnClick(CMODEREG);
  {On charge par défaut les codes cib de référence}
  MultiAfb.Plus := 'TCI_BANQUE = "' + CODECIBREF + '"';
  {$ENDIF}
end;

procedure TOF_PointeReleve.BRechercherClick (Sender : TObject) ;
begin
  case
    // Sur quel liste doit porter la recherche ?
    PGIAskCancel('Voulez-vous effectuer la recherche sur la liste des écritures comptables ?',
                 'Choix de la liste' ) of

    // Liste des écritures
    mrYes : listeRech := GC.Name ;

    // liste des relevés
    mrNo :  listeRech := GB.Name ;

    // Annulation
    mrCancel : Exit ;

  end ;

  OuvrirRecherche ;

end;

procedure TOF_PointeReleve.OnClickVPiece(Sender: TObject); //Gv 10/06/02
var LigneTob : Integer ;
    TobE : Tob ;
    Q : TQuery ;
    St : String ;
begin
    LigneTob := StrToInt(Gc.cells[SP_POSTOB,GC.Row]);
    if (ligneTob < 0) then exit;

    TobE := TobEcriture.Detail[LigneTob];

  {$IFNDEF TRESO}
//  St := 'SELECT E_JOURNAL,E_EXERCICE,E_DATECOMPTABLE,E_NUMEROPIECE,E_QUALIFPIECE,E_MODESAISIE FROM ECRITURE WHERE E_JOURNAL="' + TobE.GetValue('E_JOURNAL') + '" AND E_EXERCICE="' + TobE.GetValue('E_EXERCICE') + '" AND E_NUMEROPIECE='
  // CA - 10/08/2004 - FQ 13782
  St := 'SELECT * FROM ECRITURE WHERE E_JOURNAL="' + TobE.GetValue('E_JOURNAL') + '" AND E_EXERCICE="' + TobE.GetValue('E_EXERCICE') + '" AND E_NUMEROPIECE='
        + IntToStr(TobE.GetValue('E_NUMEROPIECE')) + ' AND E_DATECOMPTABLE="' + UsDateTime(TObE.GetValue('E_DATECOMPTABLE')) + '" AND E_QUALIFPIECE="' + TobE.GetValue('E_QUALIFPIECE') + '" AND E_MODESAISIE="' + TobE.GetValue('E_MODESAISIE') + '"';
  Q := OpenSql(St,true);

  if (TobE.GetValue('E_MODESAISIE') = '-') then TrouveEtLanceSaisie(Q,TaConsult,'')
  else LanceSaisieFolio(Q,TaConsult);
  Ferme(Q);
  {$ENDIF}
end;

procedure TOF_PointeReleve.OnCancel ;
begin
    inherited;

    if (ChangeBqeOk) then
    begin
        AfficheMsg(14, '', '');
        exit;
    end;
    if (AfficheMsg(4, '', '') = mrYes) then
    begin
        RemetRefPointageABlanc;
        RemetToutAZero;
    end;
end;

procedure TOF_PointeReleve.OnUpdate ;
var
    i: integer;
    cpt : string;
begin
    inherited;

    if (ChangeBqeOk) then
    begin
        AfficheMsg(14, '', '');
        exit;
    end;

    cpt := GetControlText('CBANQUE');//Cbanque.Value; {JP 13/08/04 : FQ 13167}

    if (TobReleve<>nil) and (TobEcriture<>nil) then
    begin
        // BPY le 10/12/2003 => Fiche 12993 : equilibre du pointage ....
        if (not (arrondi(MttB - MttC,4) = 0)) then AfficheMsg(11, '', '')
        // Fin BPY
        else
        begin
            if (AfficheMsg(5, '', '') = mrYes) then
            begin
                for i := 0 to TobReleve.Detail.Count-1 do
                begin
                    if (TobReleve.Detail[i].GetValue('POINTE') = 'X') then TobReleve.Detail[i].UpdateDB;
                end;
                for i := 0 to TobEcriture.Detail.Count-1 do
                begin
                    if (TobEcriture.Detail[i].GetValue('POINTE') = 'X') then begin
                      {$IFDEF TRSYNCHRO}
                      {Cela mettra à jour les écritures de tréso et permettra de réaliser les écritures importées}
                      TobEcriture.Detail[i].PutValue('E_TRESOSYNCHRO', 'MOD');
                      {$ENDIF}
                      TobEcriture.Detail[i].UpdateDB;
                    end;
                end;

                // BPY le 03/11/2003 : Fiche 12678 : mettre a jours les comul de pointage du compte !
                // recalcul des cumul pointé !
                RecalculTotPointeNew1(GetControlText('CBANQUE'));{JP 13/08/04 : FQ 13167}
                // fin BPY

                ChangeBqeOk := true;
                PointageAuto := false;
                PointageManu := false;
                OnClickBChercher(nil);
            end;
        end;
    end;
end;

procedure TOF_PointeReleve.OnClose ;
begin
  inherited;
  if (NbCoche <> 0) and (AfficheMsg(6, '', '') = mrNo) then
    Lasterror := 1
  else
  begin   // BPY le 23/09/2003 : fiche 12312 : on efface les variable que si on quitte REELEMENT la fiche !!
    {$IFDEF TRESO}
    if Assigned(TobRegles) then FreeAndNil(TobRegles);
    if Assigned(TobCompta) then FreeAndNil(TobCompta);
    {$ENDIF}

    //InitVariables;
    LMsg.Free ;
    FreeAndNil(AFindDialog);  // Ajout SBO fiche 12033 : Bouton rechercher
    {JP 13/08/2004 : Nouvelle gestion des filtres}
    if Assigned(ObjFiltre)    then FreeAndNil(ObjFiltre);
    {JP 01/03/06 : Correction des fuites mémoires}
    if Assigned(TobCodeAfb ) then FreeAndNil(TobCodeAfb );
    if Assigned(TobDevise  ) then FreeAndNil(TobDevise  );
    if Assigned(Tobecriture) then FreeAndNil(Tobecriture);
    if Assigned(TobReleve  ) then FreeAndNil(TobReleve  );
  end;
end;

(*******************************************************************************



     Diverses fonctions



*******************************************************************************)

function TOF_PointeReleve.GetLeModeRglt(CodeAfb: string): string ;
var
    TobMP : TOB;
begin
    result := CodeAfb;
    {$IFDEF TRESO}
    if Length(CodeAfb)= 2 then TobMP := TobCodeAFB.FindFirst(['TCI_CODECIB'],[CodeAfb],true)
                          else TobMP := TobCodeAFB.FindFirst(['TCI_CODECIB'],[Copy(CodeAfb,2,1)],true) ;
    {$ELSE}
    if length(codeafb)=2 then TobMP := TobCodeAFB.FindFirst(['AF_CODEAFB'],[CodeAfb],true)
    else TobMP := TobCodeAFB.FindFirst(['AF_CODEAFB'],[Copy(CodeAfb,2,1)],true) ;
    {$ENDIF}
    if TobMP=Nil then begin result := CodeAFB+ '/...' ; exit ; end;
    if Length(Codeafb)=1 then CodeAfb := '0' + Codeafb ;

    {$IFDEF TRESO}
    Result := CodeAFB + '/'+ TobMp.GetValue('TCI_MODEPAIE') ;
    {$ELSE}
    result := CodeAFB + '/'+TobMp.GetValue('AF_MODEPAIEMENT') ;
    {$ENDIF}
end;

Procedure TOF_PointeReleve.ChargeTobDevise ;
var
    Q : TQuery;
begin
    if TobDevise=nil then
    begin
        TobDevise := Tob.Create('Devise',nil,-1) ;
        Q := OpenSql('Select * from Devise', true) ;
        TobDevise.LoadDetailDb('Devise', '', '', Q, true) ;
        Ferme(Q) ;
    end;
end;

function TOF_PointeReleve.EcritureSelectionnee : boolean;
var
  Tolerance,i,NbeCar,PosCar : integer;
  DateTmp,DateEcriture : TDateTime;
  Okok1,Okok2 : Boolean;
  Reference,ModePaieEcr,ModePaieRel,sRef : string;
  VerifieLeCheque : Boolean;
  NomChampRefBQ : String ;
  NomChampRefECR : String ;
  SPPosBQ1,SPPosBQ2 : TSpinEdit ;
  SPPosECR1,SPPosECR2 : TSpinEdit ;
  RespectCasse : Boolean ;
  {$IFDEF TRESO}
  TobL,
  TobR    : TOB;
  Regle,
  sRefEcr,
  sRefRel : string;

    {Teste l'égalité approximative des deux dates de la colonne fournie}
    {------------------------------------------------------------------------}
    function EgalTolerant(Col : Integer) : Boolean;
    {------------------------------------------------------------------------}
    begin
      {JP 04/03/04 : sans commentaires
      Result := Trunc(Abs(StrToDate(GC.Cells[SP_DATEVAL, GC.Row]) - StrToDate(GB.Cells[SP_DATE, GB.Row]))) <= ToleranceDate;}
      Result := Trunc(Abs(StrToDate(GC.Cells[Col, GC.Row]) - StrToDate(GB.Cells[Col, GB.Row]))) <= ToleranceDate;
    end;
  {$ENDIF}

  {JP 11/04/05 : FQ 10241 : application de ce traitement à la TRESO
  {------------------------------------------------------------------------}
  function VerifCheque : Boolean;
  {------------------------------------------------------------------------}
  begin
    {$IFDEF TRESO}
    {TobMP := TobCodeAFB.FindFirst(['TCI_CODECIB'],['01'], True);
    if Assigned(TobMP) then}
    VerifieLeCheque := True;
    {$ELSE}
    VerifieLeCheque := ((CNumCheque.Checked) and (TobCodeAFB.detail[0].GetString('AF_MODEPAIEMENT') = TobEcriture.Detail[GC.Row-1].GetString('E_MODEPAIE')));
    {$ENDIF TRESO}

    if (VerifieLeCheque) then begin
      {En comptabilité, il sera toujours à True si on arrive ici ...}
      if CNumCheque.Checked then begin
        PosCar := DuCarNumCheque.value;
        NbeCar := AuCarNumCheque.value;
        NbeCar  :=  NbeCar - PosCar + 1;
// GP       Si comboref.itemindex=0 Alors ref=Idem avec CELLibelle
// GP        Reference := Copy(TobReleve.Detail[GB.Row-1].GetString('CEL_REFPIECE'), PosCar, NbeCar);

        Reference := Copy(TobReleve.Detail[GB.Row-1].GetString(NomChampRefBQ), PosCar, NbeCar);
      end
      {... ce cas ne peut arriver qu'en Tréso}
      else
        Reference := TobReleve.Detail[GB.Row-1].GetString(NomChampRefBQ);

      case comboRef.ItemIndex of
          0 : sRef := TobEcriture.Detail[GC.Row-1].GetString('E_LIBELLE');
          1 : sRef := TobEcriture.Detail[GC.Row-1].GetString('E_REFEXTERNE');
          2 : sRef := TobEcriture.Detail[GC.Row-1].GetString('E_REFINTERNE');
          3 : sRef := TobEcriture.Detail[GC.Row-1].GetString('E_REFLIBRE');
          4 : sRef := TobEcriture.Detail[GC.Row-1].GetString('E_NUMTRAITECHQ');
      end;
      Result := (Pos(Reference, sRef) > 0);
    end
    else
      Result := True;
  end;


  {GP 30/03/2008 : FQ 10241 : Correctif TDI Post Edition 11 (cf fiche onglet paramètres avancés)
  {------------------------------------------------------------------------}
  function VerifNew : Boolean;
  {------------------------------------------------------------------------}
  Var llECR,llBQ,ll : Integer ;
  begin

  SPPosBQ1:=TSpinEdit(GetControl('POSBQ1')) ; SPPosBQ2:=TSpinEdit(GetControl('POSBQ2')) ;
  SPPosECR1:=TSpinEdit(GetControl('POSECR1')) ; SPPosECR2:=TSpinEdit(GetControl('POSECR2')) ;
  sRef := TobEcriture.Detail[GC.Row-1].GetString(NomChampRefECR);
  RespectCasse:=TRUE ;
  if OkCBRefCEL And (SPPosBQ1<>NIL) And (SPPosBQ2<>NIL) And (SPPosECR1<>NIL) And (SPPosECR2<>NIL) then
    begin
    If GetControlText('CBCASSE')='-' Then RespectCasse:=FALSE ;

    If GetControlText('CBOKPOSBQ')='X' Then
      BEGIN
      PosCar := SPPosBQ1.value;
      NbeCar := SPPosBQ2.value;
      NbeCar  :=  NbeCar - PosCar + 1;
      Reference := Copy(TobReleve.Detail[GB.Row-1].GetString(NomChampRefBQ), PosCar, NbeCar);
      END Else
      BEGIN
      Reference := TobReleve.Detail[GB.Row-1].GetString(NomChampRefBQ);
      END ;
    If GetControlText('CBOKPOSECR')='X' Then
      BEGIN
      PosCar := SPPosECR1.value;
      NbeCar := SPPosECR2.value;
      NbeCar  :=  NbeCar - PosCar + 1;
      SRef := Copy(SRef, PosCar, NbeCar);
      END ; // Par défaut sref initialisé plus haut
    end else
    BEGIN
    Reference := TobReleve.Detail[GB.Row-1].GetString(NomChampRefBQ);
    END ;
  If Not RespectCasse then
    BEGIN
    Reference:=AnsiUpperCase(Reference) ;
    sRef:=AnsiUpperCase(sRef) ;
    END ;
  llECR:=Length(sRef) ; llBQ:=Length(Reference) ; ll:=Min(llECR,llBQ) ;
  Result := Copy(Reference,1,ll)=Copy(sRef,1,ll) ;
  end;

begin
NomChampRefBQ:='CEL_REFPIECE' ;
NomChampRefECR:='E_LIBELLE' ;
If OkCBRefCEL Then
  BEGIN
  NomChampRefBQ:=GetControlText('COMBOREFBQ') ;
  NomChampRefECR:=GetControlText('COMBOREFECR') ;
  END ;
{$IFDEF TRESO}
  Result := False;
  {Devise identique}
  if (GC.Cells[SP_DEVAFF, GC.Row] <> GB.Cells[SP_DEVAFF, GB.Row]) then Exit;

  if CModeReg.Checked then begin
    ModePaieEcr := GC.Cells[SP_RGLT, GC.Row] ;
    ModePaieRel := GB.Cells[SP_RGLT, GB.Row] ;
    if ((CModeReg.Checked) and (ModePaieRel <> '') and (Copy(ModePaieRel,4,3) <> ModePaieEcr)) then Exit;

    if CRegAccro.Checked then begin
      // Recherche de la règle correspondante
      TobR := TobReleve.Detail[StrToInt(GB.Cells[SP_POSTOB, GB.Row])]; // Ligne de Tob Relevé correspondante
      {JP 28/11/05 : FQ TRESO 10315 : on revient à un système plus simple
          i := StrToInt(TobR.GetValue('CEL_CODEAFB')); // Code AFB de la Tob Relevé
          Regle := '';
          Assert((i >= 1) and (i <= 99));
          Regle := TblRegles[i]; {Numéro de la règle}
      {Récupération de la ligne de paramétrage correspondant au CIB en cours}
      TobL := nil;
      if Assigned(TobR) then
        TobL := TobCodeAFB.FindFirst(['TCI_CODECIB'], [TobR.GetString('CEL_CODEAFB')], True);
      {Récupération de la règle d'accrochage}
      if Assigned(TobL) then
        Regle := TobL.GetString('TCI_REGLEACCRO');
      {Pas de règle associée}
      if Regle = '' then Exit;

      {Critères de la règle}
      TobL := TobRegles.FindFirst(['TRG_REGLEACCRO'], [Regle], True);
      if not Assigned(TobL) then Exit;
      {Date d'opération}
      if (TobL.GetValue('TRG_BDATEOPE') = 'X') and not EgalTolerant(SP_DATE) then Exit;
      {Date de valeur}
      if (TobL.GetValue('TRG_BDATEVAL') = 'X') and not EgalTolerant(SP_DATEVAL) then Exit;
      if TobL.GetValue('TRG_BPIECE') = 'X' then begin
         {JP 11/04/05 : FQ 10241 : Par contre il faudra créer une nouvelle règle d'accrochage}
        //sRefEcr := TobEcriture.Detail[StrToInt(GC.Cells[SP_POSTOB, GC.Row])].GetValue(sRef);
        //sRefRel := TobR.GetValue('CEL_REFPIECE');
        //  for i:=0 to 9 do    Autre référence ?
        //    if pos(sRefRel,TobEcriture.Detail[].GetValue('E_LIBRETEXTE'+IntToStr(i)))>0 then Ok;

        if not VerifCheque then Exit;
      end;
    end;
  end;

  Result := True;
{$ELSE}
  if (GC.Cells[SP_DEVAFF, GC.Row] <> GB.Cells[SP_DEVAFF, GB.Row]) then
  begin
      result := false;
      exit;
  end;

  Okok1 := false;
  ModePaieEcr := GC.Cells[SP_RGLT,GC.Row];
  ModePaieRel := GB.Cells[SP_RGLT,GB.Row];

  if ((CModeReg.Checked) and (ModePaieRel <> '') and (Copy(ModePaieRel,4,3) <> ModePaieEcr)) then
  begin
      result := false;
      exit;
  end;

  {Verification des dates}
  if (CDate.Checked) then
  begin
      DateEcriture := StrToDate(GC.Cells[SP_DATE,GC.Row]);
      Tolerance := TDate.Value;
      for i := Tolerance*-1 to Tolerance do
      begin
          DateTmp := StrToDate(GB.Cells[SP_DATE,GB.Row]) + i;
          if (DateTmp = DateEcriture) then
          begin
              Okok1 := true;
              break;
          end;
      end;
  end
  else Okok1 := true;

  {Verification des cheques}
  (* JP 11/04/05 : FQ 10241 : le traitement étant appliqué à la Tréso, on en fait une fonction "VerifCheque"
  VerifieLeCheque := ((CNumCheque.Checked) and (TobCodeAFB.detail[0].GetValue('AF_MODEPAIEMENT') = TobEcriture.Detail[GC.Row-1].GetValue('E_MODEPAIE')));
  if (VerifieLeCheque) then
  begin
      PosCar := DuCarNumCheque.value;
      NbeCar := AuCarNumCheque.value;
      NbeCar  :=  NbeCar - PosCar + 1;
      Reference := copy(TobReleve.Detail[GB.Row-1].GetValue('CEL_REFPIECE'),PosCar,NbeCar);
      case comboRef.ItemIndex of
          0: sRef := TobEcriture.Detail[GC.Row-1].GetValue('E_LIBELLE');
          1: sRef := TobEcriture.Detail[GC.Row-1].GetValue('E_REFEXTERNE');
          2: sRef := TobEcriture.Detail[GC.Row-1].GetValue('E_REFINTERNE');
          3: sRef := TobEcriture.Detail[GC.Row-1].GetValue('E_REFLIBRE');
          4: sRef := TobEcriture.Detail[GC.Row-1].GetValue('E_NUMTRAITECHQ');
      end;
      if (pos(Reference, sRef) > 0) then Okok2 := true;
  end
  else
    Okok2 := true;
  *)
  If OkCBRefCEL Then OkOk2:=VerifNew
                Else Okok2:=VerifCheque;
  result := (Okok1 and Okok2);
{$ENDIF}
end;

{$IFDEF TRESO}
{JP 16/12/03 : chargement des règles d'accrochage, des codecib et chargement du tableau de
               correspondance entre le code cib et la règle d'accrochage
{---------------------------------------------------------------------------------------}
procedure TOF_PointeReleve.ChargerTobs;
{---------------------------------------------------------------------------------------}
var
  Q : TQuery;
  T : TOB;
  w : string;
  i : Integer;
begin
  if not Assigned(TobCodeAfb) then
    TobCodeAfb := Tob.Create('***', nil ,-1)
  else
    {JP 23/11/04 : Dans GetLeModeRglt, on fait un FindFirst sur la Tob uniquement sur le CIB. Il faut donc
                   vider la Tob pour être sur de récupérer le bon mode de règlement}
    TobCodeAfb.ClearDetail;

  if GetControlText('CBANQUE') <> '' then {JP 13/08/04 : FQ 13167}
    w := 'WHERE TCI_BANQUE = (SELECT BQ_BANQUE FROM BANQUECP WHERE BQ_GENERAL = "' + GetControlText('CBANQUE') + '")';
  {Charge le tableau de correspondance CIB-Règle}
  Q := OpenSQL('SELECT TCI_REGLEACCRO, TCI_MODEPAIE, TCI_CODECIB FROM CIB ' + w + ' ORDER BY TCI_CODECIB', True);
  i := 1;
  try
    while not Q.EOF do begin
      {JP 28/11/05 : FQ TRESO 10315 : la gestion des équivalences par l'intermédiaire d'un tableau
                     était une fausse bonne idée : en effets, si un CIB a été supprimé, tout le tableau
                     se trouve décalé et il devient impossible d'établir une correspondance entre l'indice
                     du tableau et le CIB
       TblRegles[i] := Q.Fields[0].AsString; // Règle 0 = pas de règle}
      T := TOB.Create('***', TobCodeAfb, -1);
      T.AddChampSupValeur('TCI_CODECIB'   , Q.Fields[2].AsString);
      T.AddChampSupValeur('TCI_MODEPAIE'  , Q.Fields[1].AsString);
      T.AddChampSupValeur('TCI_REGLEACCRO', Q.Fields[0].AsString); {JP 28/11/05 : FQ TRESO 10315}
      Q.Next;
      Inc(i);
    end;
  finally
    Ferme(Q);
  end ;

  {Charge les critères dans une Tob}
  if TobRegles = nil then begin
    TobRegles := TOB.Create('_REGLEACCRO', nil, -1);
    TobRegles.LoadDetailDB('REGLEACCRO', '', '', nil, False, True);
  end;
end;

{$ELSE}

Procedure TOF_PointeReleve.ChargeTobCodeAFB ;
var
    Q : TQuery ;
begin
    if (TobCodeAfb = nil) then
    begin
        TobCodeAfb := Tob.Create('AFB',nil,-1);
        Q := OpenSql('Select * from CodeAfb',true);
        TobCodeAfb.LoadDetailDb('CodeAFB','','',Q,true);
        Ferme(Q);
    end;
end;
{$ENDIF}

Function TOF_PointeReleve.RechercheDevise(Dev: String) : String ;
var TobD : Tob ;
begin
result := '' ;
PasDePointagePossible := false ;
TobD := TobDevise.Findfirst(['D_DEVISE'],[Dev],true) ;
if TobD <> Nil then result := TobD.GetValue('D_CODEISO') else
  begin
  PasDePointagePossible := true ;
  AfficheMsg(22, '', '') ;
  end;
end;

Function TOF_PointeReleve.RechercheDecimale(Dev: String) : Integer ;
var TobD : Tob ;
begin
result := 4 ;
TobD := TobDevise.Findfirst(['D_DEVISE'],[Dev],true) ;
if TobD <> Nil then result := TobD.GetValue('D_DECIMALE') ;
end;


procedure TOF_PointeReleve.PositionneDevise(Tout: boolean) ;
var
  sCompta,
  sReleve : string ;
begin
    if CDevise.ItemIndex=-1 then SetControlText('Devise',LDP);
    sCompta := cDevise.Value ;
    if Cdevreleve.ItemIndex=-1 then SetControlText('CdevReleve',LDP) ;
    sReleve := cDevReleve.Value ;

    if (sCompta<>LDF) and (sCompta<>LDP) then
    begin
        C_DEBIT := 'DEBITDEV' ;
        C_CREDIT := 'CREDITDEV';
    end
    else
    begin
        C_DEBIT := 'DEBIT' ;
        C_CREDIT := 'CREDIT';
    end;
//       if sCompta=LDP then
//               begin C_DEBIT := 'DEBIT' ;     C_CREDIT := 'CREDIT'      ; end
//          else begin C_DEBIT := 'DEBITEURO' ; C_CREDIT := 'CREDITEURO'  ; end;
    DeviseCompta := RechercheDevise(SCompta);
    DecCompta := RechercheDecimale(Scompta) ;

    if (sReleve<>LDF) and (sReleve<>LDP) then
    begin
        A_DEBIT := 'DEBITDEV' ;
        A_CREDIT := 'CREDITDEV';
    end
    else
    begin
        A_DEBIT := 'DEBITEURO' ;
        A_CREDIT := 'CREDITEURO';
    end;
//         if sReleve=LDP then
//                 begin A_DEBIT := 'DEBITEURO' ;     A_CREDIT := 'CREDITEURO' ;    end
//            else begin A_DEBIT := 'DEBIT' ; A_CREDIT := 'CREDIT'; end;
    DeviseReleve := RechercheDevise(sReleve) ;
    DecReleve := RechercheDecimale(SReleve) ;

end;

(*******************************************************************************

    Méthodes liées a la fiche AGL et aux controles de la fiche AGL

*******************************************************************************)
function TOF_PointeReleve.RecupeControles: boolean ;
begin
try
  OngletHaut   :=TPageControl(GetControl('ONGLETHAUT')) ;
  GC           :=THGrid(GetControl('GCOMPTA'));
  GB           :=THGrid(GetControl('GBANQUE'));
  BChercher    :=TButton(GetControl('BCHERCHER')) ;
  BPointageAuto:=TButton(GetControl('BPOINTAGEAUTO')) ;
  BValider     :=TButton(GetControl('BVALIDER')) ;
  BParam       :=TButton(GetControl('BPARAM')) ;
  BAgrandir    :=TButton(GetControl('BAGRANDIR')) ;
  BReduire     :=TButton(GetControl('BREDUIRE')) ;
  BtnAFB       :=TToolBarbutton97(GetControl('BTNAFB')) ;
  BRechercher  :=TToolBarbutton97(GetControl('BRECHERCHER')) ;
  BImprimer    :=TToolBarbutton97(GetControl('BIMPRIMER')) ;

  BImprimer.Visible := True; // Rendre visible 03/02/2003 car pb SocRef

  CBanque      :=THValCombobox(GetControl('CBANQUE')) ;
  {$IFDEF TRESO}
  {JP 13/08/04 : Remplacement de TTBANQUECP par TZGBANQUE, pour la gestion des confidentiels,
                 cf FQ 13167. Par contre en Tréso, c'est la tablettes TTBANQUECP qui est
                 toujours utilisée}
  CBANQUE.Plus := ' BQ_GENERAL IN (SELECT G_GENERAL FROM GENERAUX WHERE G_POINTABLE="X")';
  {$ELSE}
  CBANQUE.Plus := GetPlus;
  {$ENDIF}
  CRefReleve   :=THValCombobox(GetControl('CREFRELEVE')) ;
  CDevReleve   :=THValCombobox(GetControl('CDEVRELEVE')) ;
  CDevise      :=THValCombobox(GetControl('DEVISE')) ;
  CAffDev      :=THValCombobox(GetControl('AFFDEV')) ;
  CDate        :=TCheckBox(GetControl('CDATE')) ;
  CNumCheque   :=TCheckBox(GetControl('CNUMCHEQUE')) ;
  CCombinatoire:=TCheckBox(GetControl('CCOMBINATOIRE')) ;
  CModeReg     :=TCheckBox(GetControl('CMODEREG')) ;
  ModeTiers    :=TCheckBox(GetControl('MODETIERS')) ;
  TTotalC      :=THEdit(GetControl('TTOTALC')) ;
  TTotalB      :=THEdit(GetControl('TTOTALB')) ;
  TEcart       :=THEdit(GetControl('TECART')) ;
  TDate        :=TSpinEdit(GetControl('TDATE')) ;
  DuCarNumCheque   :=TSpinEdit(GetControl('TNUMCHEQUE')) ;
  AuCarNumCheque   :=TSpinEdit(GetControl('TPOSCHEQUE')) ;
  TNiveau      :=TSpinEdit(GetControl('TNIVEAU')) ;
  DateDe       := THEdit(GetControl('DATEDE'));
  DateJusqua   := THEdit(GetControl('DATEJUSQUA'));
  comboRef       := THValCombobox(GetControl('COMBOREF'));

  MultiAFB     :=ThMultiValComboBox(GetControl('MULTIAFB')) ;
  MultiModepaie:=ThMultiValComboBox(GetControl('MULTIMODEPAIE')) ;

  BZoom        := TButton(GetControl('BZOOM'));
  PopZoom      := TPopupMenu(GetControl('POPZOOM'));
  vpiece           := PopZoom.Items[0];
  vcompte          := PopZoom.Items[1];

  BHelp:=TToolbarButton97(GetControl('BAIDE')) ;

  pcpta := TPanel(GetControl('PCPTA'));
  pbanq := TPanel(GetControl('PBANQ'));
  pgene := TPanel(GetControl('PGENE'));
  {$IFDEF TRESO}
  CRegAccro    :=TCheckBox(GetControl('CREGACCRO')) ;
  {$ENDIF}

  // Suppression de la date de pointage de la fiche (Fiche 10695) 31/01/2001
  if GetControl('DATEPOINT') <> nil then
    THEdit(GetControl('DATEPOINT')).Visible := false ;
  if GetControl('LBDATEPOINT') <> nil then
    THEdit(GetControl('LBDATEPOINT')).Visible := false ;

except
  result:=false ;
  PgiBox(TraduireMemoire('Une erreur est survenue sur la fiche, vérifiez la version de votre base'),'Erreur affichage fiche') ;
  exit;
end;

result:=true ;

end;

function TOF_PointeReleve.AssignControles: boolean ;
begin
try
  if Ecran<>nil then begin
     TFORM(Ecran).OnKeyDown:=OnKeyDown ;
     if V_PGI.LaSerie = S7 then
          TFORM(Ecran).OnResize :=OnFormResize ;
  end;
  if GC<>nil then begin
    if not Assigned(GC.OnDblClick)    then GC.OnDblClick:=OnDblClickGCompta;
    if not Assigned(GC.GetCellCanvas) then GC.GetCellCanvas:=GetCellCanvasC ;
    end;
  if GB<>nil then begin
    if not Assigned(GB.OnDblClick)    then GB.OnDblClick:=OnDblClickGBanque;
    if not Assigned(GB.GetCellCanvas) then GB.GetCellCanvas:=GetCellCanvasB ;
    end;
  if (CDate<>nil)         and (not Assigned(CDate.OnClick))         then CDate.OnClick         :=OnClickCDate ;
  if (CNumCheque<>nil)    and (not Assigned(CNumCheque.OnClick))    then CNumCheque.OnClick    :=OnClickCNumCheque ;
  if (CCombinatoire<>nil) and (not Assigned(CCombinatoire.OnClick)) then CCombinatoire.OnClick :=OnClickCCombinatoire ;
  if (BChercher<>nil)     and (not Assigned(BChercher.OnClick))     then BChercher.OnClick     :=OnClickBChercher ;
  if (BImprimer<>nil)     and (not Assigned(BImprimer.OnClick))     then BImprimer.OnClick     :=OnClickBImprimer ;
  if (BPointageAuto<>nil) and (not Assigned(BPointageAuto.OnClick)) then BPointageAuto.OnClick :=OnClickBPointageAuto ;
  if (BtnAFB<>nil)        and (not Assigned(BTNAFB.OnClick))        then BtnAFB.OnClick        :=BtnAFBOnClick ;
  if (CBanque<>nil)       and (not Assigned(CBanque.OnChange))      then CBanque.OnChange      :=OnChangeCBanque ;
  if (BAgrandir<>nil)     and (not Assigned(BAgrandir.OnClick))     then BAgrandir.OnClick     :=OnClickBAgrandir ;
  if (BReduire<>nil)      and (not Assigned(BReduire.OnClick))      then BReduire.OnClick      :=OnClickBReduire ;
  if (DateDe<>nil)        and (not Assigned(DateDe.OnExit))         then DateDe.OnExit         :=OnDateExit ;
  if (DateJusqua<>nil)    and (not Assigned(DateJusqua.OnExit))     then DateJusqua.OnExit     :=OnDateExit ;
  if (POPZOOM<>nil)       and (not Assigned(POPZOOM.OnPopup))       then PopZoom.OnPopup       :=OnPopUpZoom ;
  if (Vpiece<>nil)        and (not Assigned(Vpiece.OnClick))        then Vpiece.OnClick        :=OnClickVpiece ;
  if (Vcompte<>nil)       and (not Assigned(Vcompte.OnClick))       then Vcompte.OnClick       :=OnClickVcompte ;
  if (BParam<>nil)        and (not Assigned(BParam.OnClick))        then BParam.OnClick        :=OnClickBParam;
  if (BHelp <> nil)       and (not Assigned(BHelp.OnClick))         then BHelp.OnClick         :=BHelpClick ;
  if (Brechercher <> nil) and (not Assigned(BRechercher.OnClick))   then BRechercher.OnClick   :=BRechercherClick  ;
  {$IFDEF TRESO}
  if (CRegAccro <> nil)   and (not Assigned(CRegAccro.OnClick))     then CRegAccro.OnClick     := CRegAccroOnClick;
  if (CModeReg  <> nil)   and (not Assigned(CModeReg .OnClick))     then CModeReg .OnClick     := CModeRegOnClick;
  {$ENDIF}
except
  result:=false ;
  PgiBox(TraduireMemoire('Une erreur est survenue sur la fiche, vérifiez la version de votre base'),'Erreur affichage fiche') ;
  exit;
end;
result:=true ;
end;

function TOF_PointeReleve.InitiaControles: boolean ;
begin
try
  OngletHaut.ActivePage:=OngletHaut.Pages[0] ;
  InitGrid(GC) ;
  InitGrid(GB,TRUE) ;
  {$IFDEF TRESO}
  CBanque.ItemIndex:=0 ;
  {$ENDIF}
  RemplirValRelev('') ;
  RemplirValDevis('') ;
  CDate.Checked:=false ;
  CNumCheque.Checked:=false ;
  CCombinatoire.Checked:=false ;
  CModeReg.Visible:=true ;
  CModeReg.Checked:=false;
  AfficheMontant(TTotalC,MttC);
  AfficheMontant(TTotalB,MttB);
  cDevReleve.ItemIndex:=0 ;
  ChangeBqeOk:=true ;
  InitVariables ;
  OnClickCDate(nil);
  DateDe.Text := DateToStr(VH^.Encours.Deb);
  DateJusqua.Text := DateToStr(V_PGI.DateEntree);
  comboRef.ItemIndex := 0;
  comboRef.enabled := false;
    {$IFDEF TRESO}
    {JP 28/11/05 : FQ TRESO 10315 : TblRegles[1] := '-1';}
    {$ENDIF}

    // BPY le 08/12/2003 => fiche 13075 : vas de paire avec une bob !
    pcpta.Height := pgene.ClientHeight div 2;
    // fin BPY

except
  result:=false ;
  PgiBox(TraduireMemoire('Une erreur est survenue lors de l''initialisation de la fiche, vérifiez la version de votre base'),'Erreur affichage fiche') ;
  exit;
end;
result:=true ;
end;

procedure TOF_PointeReleve.OnChangeCBanque(Sender: TObject) ;
var
  s : string; {JP 13/08/04 : FQ 13167}
begin
  s := GetControlText('CBANQUE');
  CDevReleveOK:=true ;
  ChangeBqeOk:=true ;
  RemplirValRelev(s);
  if not ObjFiltre.InChargement then
    RemplirValDevis(s);
  {$IFDEF TRESO}
  MultiAfb.Plus := 'TCI_BANQUE = (SELECT BQ_BANQUE FROM BANQUECP WHERE BQ_GENERAL = "' + s + '")';
  {$ENDIF}
end;

procedure TOF_PointeReleve.OnClickCDate(Sender: TObject) ;
begin
  {$IFDEF TRESO}
  TDate.Enabled := CDate.Checked and CDate.Enabled;
  {$ELSE}
  if CDate.Checked then TDate.Enabled:=true else TDate.Enabled:=false ;
  {$ENDIF}
end;

procedure TOF_PointeReleve.OnClickCNumCheque(Sender:  TObject) ;
begin
  {$IFDEF TRESO}
  DuCarNumCheque.Enabled := CNumCheque.Checked and CNumCheque.Enabled;
  AuCarNumCheque.Enabled := DuCarNumCheque.Enabled;
  ComboRef      .Enabled := DuCarNumCheque.Enabled;
  SetControlEnabled('LBCARACT', DuCarNumCheque.Enabled); {JP 26/04/05 : FQ TRESO 10241}
  SetControlEnabled('LBSUR', DuCarNumCheque.Enabled); {JP 26/04/05 : FQ TRESO 10241}
  {$ELSE}
  DuCarNumCheque.Enabled := not DuCarNumCheque.Enabled;
  AuCarNumCheque.Enabled := not AuCarNumCheque.Enabled;
  comboRef      .Enabled := CNumCheque.Checked;
  {$ENDIF}
end;

procedure TOF_PointeReleve.OnClickCCombinatoire(Sender:  TObject) ;
begin
TNiveau.Enabled:=not TNiveau.Enabled ;
end;

procedure TOF_PointeReleve.OnClickBChercher(Sender : TObject) ;
var Okok : boolean ;
begin
Okok:=true ;
if PointageAuto or PointageManu then if AfficheMsg(4,'','')=mrNo then Okok:=false ;
if Okok then begin
  {$IFDEF TRESO}
  ChargerTobs; {Chargement des règles et des codes cibs}
  {$ENDIF}

  RemetRefPointageABlanc;
  InitVariables ;
  PositionneDevise(true) ;
  RempliLesTobs ;
  RempliReleve ;
  RempliCompta ;
  LesVisibles(true) ;
  ChangeBqeOk:=false ;
  end;
end;

procedure TOF_PointeReleve.OnDblClickGBanque(Sender: TObject) ;
begin
if GB.Cells[SP_DATE,GB.Row]<>'' then PointeManu(GB);
end;

procedure TOF_PointeReleve.OnDblClickGCompta(Sender: TObject) ;
begin
if GC.Cells[SP_DATE,GC.Row]<>'' then PointeManu(GC) ;
end;

procedure TOF_PointeReleve.OnClickBAgrandir(Sender:  TObject) ;
begin
ChangeAffichage(true) ;
end;

procedure TOF_PointeReleve.OnClickBReduire(Sender:  TObject) ;
begin
ChangeAffichage(false) ;
end;

procedure TOF_PointeReleve.OnClickBPointageAuto(Sender : TObject) ;
begin
if ChangeBqeOk then  begin AfficheMsg(14,'','') ; exit ; end;
if PointageManu then begin AfficheMsg(7,'','') ; exit ; end;
if CNUMCHEQUE.Checked and (DuCarNumCheque.value > AuCarNumCheque.value) then
begin
   AfficheMsg(16,'','');
   exit;
end;
if CModeReg.Checked and (not ExisteMoyenPaiement) then exit ;

If GB.RowCount>500 Then If PgiAsk('Confirmez-vous le Traitement ?','')<>mrYes Then Exit ;

RemetRefPointageABlanc;
RemetToutAZero ;
LanceLePointageAuto;
end;


procedure TOF_PointeReleve.GetCellCanvasB(ACol,ARow : LongInt ; Canvas : TCanvas; AState: TGridDrawState) ;
var TLibB,TDateVal,TNumLigB : THLabel ;
begin
ChangeAffichageGrid(ARow,GB) ;
TLibB:=THLabel(GetControl('TLIBB')) ;
TDateVal:=THLabel(GetControl('TDATEVAL')) ;
TNumLigB:=THLabel(GetControl('TNUMLIGB')) ;
TLibB.Caption:=GB.Cells[SP_LIBELLE,GB.Row] ;
TDateVal.Caption:=GB.Cells[SP_DATEVAL,GB.Row] ;
TNumLigB.Caption:=GB.Cells[SP_CLASSER,GB.Row] ;
end;

procedure TOF_PointeReleve.GetCellCanvasC(ACol,ARow : LongInt ; Canvas : TCanvas; AState: TGridDrawState) ;
var TLibC,TdateEch,TNumLigC : THLabel ;
begin
ChangeAffichageGrid(ARow,GC) ;
TLibC:=THLabel(GetControl('TLIBC')) ;
TdateEch:=THLabel(GetControl('TDATEECH')) ;
TNumLigC:=THLabel(GetControl('TNUMLIGC')) ;
TLibC.Caption:=GC.Cells[SP_LIBELLE,GC.Row] ;
TdateEch.Caption:=GC.Cells[SP_DATEVAL,GC.Row] ;
TNumLigC.Caption:=GC.Cells[SP_CLASSER,GC.Row] ;
end;

procedure TOF_PointeReleve.FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
var G:THGrid; OkG,Vide : Boolean ;
begin
  G    := THGrid(Sender);
  OkG  := G.Focused ;
  Vide := (Shift=[]) ;
  Case Key of
    // Espace
    VK_SPACE  : if ((OkG) and (Vide)) then PointeManu(G);
    //Ctrl + H
    70: if Shift = [ssCtrl] then
          begin
          if G.Name = GB.Name
            then listeRech := GB.Name   // recherche sur la liste des relevés
            else listeRech := GC.Name ; // recherche sur la liste des écritures
          OuvrirRecherche ;
          end ;
    end;
end;

procedure TOF_PointeReleve.GMouseUp(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
var G:THGrid;C,R : Longint ;
begin
G:=THGrid(Sender);
if ((ssCtrl in Shift) and (Button=mbLeft)) then begin
   G.MouseToCell(X,Y,C,R) ;
   if R>0 then PointeManu(G) ;
   end;
end;

procedure TOF_PointeReleve.InitGrid(GS : THGrid ; Okel : Boolean = FALSE) ;
var i,k : integer ;
    CB : thValComboBox ;
    OkRepositionneTitre : Boolean ;
begin
GS.RowCount := 2; GS.FixedRows := 1;
GS.FColAligns[SP_SOLDE]:=taRightJustify;
if not Assigned(GS.OnKeyDown)  then GS.OnkeyDown:=FormKeyDown;
if not Assigned(GS.OnMouseUp)  then GS.OnMouseUp:=GMouseUp;
for i:=0 to GS.ColCount-1 do GS.Cells[i,1]:= '';
GS.ColWidths[SP_POINTER]:=-1;
GS.ColWidths[SP_ROW]:=-1;
GS.ColWidths[SP_CLASSER]:=-1;
GS.ColWidths[SP_DEVISE]:=-1;
GS.ColWidths[SP_POSTOB]:=-1;
OkRepositionneTitre:=(GetControlText('CBPARAM2')='X') OR (GetControlText('CBPARAM2')='-') ;
If OkRepositionneTitre Then
  BEGIN
  If OkEl Then
    BEGIN
    CB:=thValComboBox(GetControl('COMBOREFBQ')) ;
    If CB<>NIL THen
      BEGIN
      If Cb.ItemIndex>=0 Then
        BEGIN
        k:=Cb.ItemIndex ;
        If (CB.Values[k]='CEL_REFPIECE') Or (CB.Values[k]='CEL_LIBELLE')
          Then GS.Cells[SP_REFERENCE,0]:=traduireMemoire('Référence')
          Else GS.Cells[SP_REFERENCE,0]:=cb.items[Cb.ItemIndex] ;
        END ;
      END ;
    END Else
    BEGIN
    CB:=thValComboBox(GetControl('COMBOREFECR')) ;
    If CB<>NIL THen
      BEGIN
      If Cb.ItemIndex>=0 Then
        BEGIN
        k:=Cb.ItemIndex ;
        If (CB.Values[k]='E_REFINTERNE') Or (CB.Values[k]='E_LIBELLE')
          Then GS.Cells[SP_REFERENCE,0]:=traduireMemoire('Référence')
          Else GS.Cells[SP_REFERENCE,0]:=cb.items[Cb.ItemIndex] ;
        END ;
      END ;
    END ;
  END ;
end;

procedure TOF_PointeReleve.RemetToutAZero ;
begin
InitVariables ;
RempliReleve ;
RempliCompta ;
LesVisibles(true) ;
AfficheMontant(TTotalC,MttC);
AfficheMontant(TTotalB,MttB);
end;

procedure TOF_PointeReleve.InitVariables ;
begin
gdDateReleve := V_PGI.DateEntree ; // Init date relevé
MttB:=0;
MttC:=0;
PositionneDevise(true) ;
PointageAuto:=false;
PointageManu:=false;
NbCoche:=0 ;
NbCocheC:=0 ;
NbCocheB:=0 ;
RefReleve:='' ;
BValider.Enabled := false;
end;

procedure TOF_PointeReleve.OnKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
  case Key of
    // F9
    VK_F9 : OnClickBChercher(nil) ;
  end ;
end;

(*******************************************************************************



    Méthodes liées aux messages de l'objet



*******************************************************************************)
procedure TOF_PointeReleve.InitMsg ;
begin
LMsg:=THMsgBox.create(FMenuG) ;
{00}LMsg.Mess.Add(traduirememoire('0;Pointage des Relevés bancaires;Erreur Inconnue.;W;O;O;O'));
{01}LMsg.Mess.Add(traduirememoire('1;Pointage des Relevés bancaires;Vous ne pouvez pas faire de pointage manuel pendant un pointage automatique.;W;O;O;O'));
{02}LMsg.Mess.Add(traduirememoire('2;Pointage des Relevés bancaires;Le relevé existe déjà, voulez-vous l''integrer ?;Q;YN;N;N'));
{03}LMsg.Mess.Add(traduirememoire('3;Pointage des Relevés bancaires;Aucune banque pour ce relevé.;W;O;O;O'));
{04}LMsg.Mess.Add(traduirememoire('4;Pointage des Relevés bancaires;Voulez-vous annuler le pointage en cours ?;Q;YN;N;N'));
{05}LMsg.Mess.Add(traduirememoire('5;Pointage des Relevés bancaires;Voulez-vous enregistrer le pointage ?;Q;YN;N;N'));
{06}LMsg.Mess.Add(traduirememoire('6;Pointage des Relevés bancaires;Voulez-vous abandonnez le pointage en cours ?;Q;YN;N;N'));
{07}LMsg.Mess.Add(traduirememoire('7;Pointage des Relevés bancaires;Vous ne pouvez pas lancer de pointage automatique pendant un pointage manuel.;W;O;O;O'));
{08}LMsg.Mess.Add(traduirememoire('8;Pointage des Relevés bancaires;Voulez-vous effacer le relevé ?;W;YN;N;N'));
{09}LMsg.Mess.Add(traduirememoire('9;Pointage des Relevés bancaires;Vous devez sélectionner une ligne de relevé.;W;O;O;O'));
{10}LMsg.Mess.Add(traduirememoire('10;Pointage des Relevés bancaires;Vous ne pouvez sélectionner qu''un seul relevé à la fois.;W;O;O;O'));
{11}LMsg.Mess.Add(traduirememoire('11;Pointage des Relevés bancaires;Le pointage n''est pas équilibré, impossible de l''enregistrer.;W;O;O;O'));
{12}LMsg.Mess.Add(traduirememoire('12;Pointage des Relevés bancaires;  %% ligne(s) pointée(s)                     ;E;O;O;O'));
{13}LMsg.Mess.Add(traduirememoire('13;Pointage des Relevés bancaires;Devises incompatibles.;E;O;O;O'));
{14}LMsg.Mess.Add(traduirememoire('14;Pointage des Relevés bancaires;Vous devez lancer une recherche avant d''effectuer un traitement.;E;O;O;O'));
{15}LMsg.Mess.Add(traduirememoire('15;Pointage des Relevés bancaires;Aucun mouvement ne peut être pointé.;E;O;O;O'));
{16}LMsg.Mess.Add(traduirememoire('16;Pointage des Relevés bancaires;Paramètre de la vérification des numéros de chèque incorrect.;E;O;O;O'));
{17}LMsg.Mess.Add(traduirememoire('17;ATTENTION;Vous devez paramétrer un code moyen de paiement associé à vos codes AFB %%#13#10afin de pouvoir faire un pointage automatique avec vérification des modes de règlements.;E;O;O;O'));
{18}LMsg.Mess.Add(traduirememoire('18;ATTENTION;La fourchette de date est incorrecte.;E;O;O;O'));
{19}LMsg.Mess.Add(traduirememoire('19;ATTENTION;La date de pointage est antérieure à la date de début d''exercice.;E;O;O;O'));
{20}LMsg.Mess.Add(traduirememoire('20;ATTENTION;La date de pointage est postérieure à la date de fin d''exercice.;E;O;O;O'));
{21}LMsg.Mess.Add(traduirememoire('21;ATTENTION;Certains montants n''ont pu être pointés correctement.;E;O;O;O'));
{22}LMsg.Mess.Add(traduirememoire('22;Pointage des relevés bancaires;le pointage ne peut s''effectuer. Verifiez les codes ISO des devises.;E;O;O;O'));
{23}LMsg.Mess.Add(traduirememoire('23;ATTENTION;Vous ne pourrez pas lettrer des lignes de relevés bancaires. Voulez-vous continuer ?;Q;YN;O;O'));
{JP 17/08/05 : FQ 15433 : modification du message pour le rendre plus clair}
{24}LMsg.Mess.Add(traduirememoire('24;Pointage des Relevés bancaires;Souhaitez-vous : Imprimer les lignes de relevés bancaires (Oui)?'#13 +
                                                         '                            Imprimer les écritures comptables (Non) ?'#13 +
                                                         '                            Abandonner (Annuler) ?;Q;YNC;Y;C'));
{JP 12/01/05 : FQ 15078 : on n'interdit le traitement si pas de lignes de relevé}
{25}LMsg.Mess.Add(traduirememoire('25;Pointage des Relevés bancaires;Il n''est pas possible de pointer manuellement s''il n''y pas de ligne de relevé.;E;O;O;O'));
end;

function TOF_PointeReleve.AfficheMsg(num : integer;Av,Ap : string ) : Word ;
begin
result:=mrNone ;  if LMsg=nil then exit ;
if (Num<0) or (Num>LMsg.Mess.Count-1) then Num:=0 ; //erreur inconnue
result:=LMsg.Execute(num,Av,Ap) ;
end;

function TOF_PointeReleve.PutMontant(Montant : Double ; Dec: integer) : string ;
begin
result:='' ;
NbDecDev := Dec ;
if Montant<>0 then result:=StrFMontant(Montant, 15, NbDecDev,'',true)
end;

procedure TOF_PointeReleve.RemplirValRelev(Banque: string) ;
var Q: TQuery ;
    i: integer;
begin
if (CRefReleve=nil) or not (CRefReleve is THValCombobox) then exit ;
CRefReleve.Items.clear ; CRefReleve.Values.clear ;CRefReleve.items.Add(traduireMemoire('<<Tous>>')) ;
Q:=OpenSQL('SELECT CEL_REFPOINTAGE FROM EEXBQLIG WHERE CEL_GENERAL="'+Banque
          +'" AND CEL_DATEPOINTAGE="'+UsDateTime(idate1900)+'" GROUP BY CEL_REFPOINTAGE',true) ;

try
  while Not Q.EOF do begin
    CRefReleve.items.Add(Q.FindField('CEL_REFPOINTAGE').AsString) ;
    Q.Next ;
    end;
  for i := 0 to CRefReleve.items.Count - 1 do
      CRefReleve.Values.Add(inttostr(i));
finally
  ferme(Q) ;
  CRefReleve.ItemIndex:=0 ;
  CRefReleve.Enabled:=(CRefReleve.Items.Count>1) ;
  end;
end;

procedure TOF_PointeReleve.RemplirValDevis(Banque: string) ;
begin
if (CAffDev=nil) or not (CAffDev is THValCombobox) then exit ;
setcontroltext('devise',LDP) ;
SetControlText('Cdevreleve',LDP) ;
SetControlText('AffDev',LDP) ;
end;

procedure TOF_PointeReleve.AfficheMontant(TTotal : THEdit;Mtt : Double) ;
var Decim : Integer ;
    Devi : String ;
begin
if TTotal.Name='TTOTALC' then
  begin
  Decim:=DecCompta ;
  devi:=DeviseCompta ;
  end else
  begin
  Decim:=DecReleve ;
  Devi:=DeviseReleve ;
  end;
if ( Mtt < 0 ) then TTotal.Text:=PutMontant(0-Mtt, Decim)+' D' else
if ( Mtt > 0 ) then TTotal.Text:=PutMontant(Mtt, Decim)+' C' else
                    TTotal.Text:= '0';
TTotal.Text:=TTotal.Text+' '+Devi  ;
//if (arrondi(abs(MttB) - abs(MttC),DecCompta) = 0) then TEcart.Text := '0'
  //                                                else TEcart.Text := PutMontant(arrondi(abs(MttB) - abs(MttC),DecCompta),DecCompta);
  {JP 12/08/04 : FQ 13403 : si -150 et 150, avec Abs cela donne 0, alors que l'écart est de 300.
                            Par ailleurs MttC était faux (cf. CocheDecocheGC) car on faisait Cred - Deb}
  if (Arrondi(MttB - MttC, DecCompta) = 0) then TEcart.Text := '0'
                                           else TEcart.Text := PutMontant(Arrondi(MttB - MttC, DecCompta), DecCompta);
end;

procedure TOF_PointeReleve.LesVisibles(Visible : Boolean) ;
begin
if CBanque <> nil  then CBanque.Enabled    :=Visible ;
if cRefReleve<>nil then cRefReleve.Enabled :=Visible ;
if cDevReleve<>nil then cDevReleve.Enabled :=(Visible and CDevReleveOk) ;
if cDevise<>nil    then cDevise.Enabled    :=Visible ;
if cAffDev<>nil    then cAffDev.Enabled    :=Visible ;
end;

procedure TOF_PointeReleve.ChangeAffichage(Plus: boolean) ;
begin
if not (BAgrandir is TButton) or not (BReduire is TButton) or not (OngletHaut is TPageControl) then exit ;
BAgrandir.Visible:=Not Plus ; BReduire.Visible:=Plus ;
if plus then OngletHaut.Align:=AlNone else OngletHaut.Align:=AlTop ;
OngletHaut.Visible:=not Plus ;
end;

procedure TOF_PointeReleve.ChangeAffichageGrid(ARow : LongInt ; GS : THGrid ) ;
begin
if ARow<=0 then exit ;
end;

procedure TOF_PointeReleve.RempliReleve;
var  i,J :integer;TR : TOB;
     NomChampRefBQ : String ;
//     NomChampRefECR : String ;
begin
InitGrid(GB,TRUE) ;
J:=1 ;
NomChampRefBQ:='CEL_REFPIECE' ;
//NomChampRefECR:='E_LIBELLE' ;
If OkCBRefCEL Then
  BEGIN
  NomChampRefBQ:=GetControlText('COMBOREFBQ') ;
//  NomChampRefECR:=GetControlText('COMBOREFECR') ;
  END ;
if TobReleve.Detail.Count<>0 then begin
  for i:=0 to TobReleve.Detail.Count-1 do begin
    TR:=TobReleve.Detail[i] ;
    if (CRefReleve.ItemIndex=0) or (CRefReleve.Items.strings[CRefReleve.ItemIndex]=TR.GetValue('CEL_REFPOINTAGE')) then begin
      GB.RowCount:=J+1 ;
      GB.Cells[SP_DATE,J]     :=TR.GetValue('CEL_DATEOPERATION');
      GB.Cells[SP_REFPOINTAGE,J]:=TR.GetValue('CEL_REFPOINTAGE');
      If (NomChampRefBQ='CEL_REFPIECE') Or (NomChampRefBQ='CEL_LIBELLE')
        Then GB.Cells[SP_REFERENCE,J]:=TR.GetValue('CEL_REFPIECE')
        Else GB.Cells[SP_REFERENCE,J]:=TR.GetValue(NomChampRefBQ);
      GB.Cells[SP_SOLDE,J]    :=PutMontant(TR.GetValue('CEL_'+A_CREDIT)-TR.GetValue('CEL_'+A_DEBIT),DecReleve) ;
      GB.Cells[SP_RGLT,J]     :=GetLeModeRglt(TR.GetValue('CEL_CODEAFB'));
      GB.Cells[SP_DEVISE,J]   :=TR.GetValue('D_CODEISO') ;
      GB.Cells[SP_LIBELLE,J]  :=TR.GetValue('CEL_LIBELLE');
      GB.Cells[SP_DATEVAL,J]  :=TR.GetValue('CEL_DATEVALEUR');
      GB.Cells[SP_POINTER,J]  :='' ;
      GB.Cells[SP_NUMPOINTER,J]  :='' ;
      GB.Cells[SP_DEVAFF,J]   :=TR.GetValue('D_CODEISO') ;
      GB.Cells[SP_POSTOB,J]   :=IntToStr(i) ;
      GB.Cells[SP_CLASSER,J]  :=IntToStr(i+1) ;
      {$IFDEF TRESO}
      {27/04/05 : FQ TRESO 10246 : ajout du libellé enrichi, en version trésorerie
       19/12/05 : FQ TRESO 10319 : LIBELLE1, c'est mieux que LIBELLE !!!}
      GB.Cells[SP_LIBENRICHI, J] := TR.GetValue('CEL_LIBELLE1');
      {$ENDIF TRESO}
      if Arrondi(Valeur(GB.Cells[SP_SOLDE, J]), NbDecDev)=0 then GB.RowHeights[j]:=15 else GB.RowHeights[J]:=GB.DefaultRowHeight;
      inc(J) ;
      end;
    end;
  end;
MttB:=0 ;
AfficheMontant(TTotalB,MttB);
end;

procedure TOF_PointeReleve.RempliCompta ;
var i,J : integer ; TR : TOB ;
    NomChampRefECR : String ;
begin
InitGrid(GC) ;
J:=1 ;
NomChampRefECR:='E_REFINTERNE' ;
//NomChampRefECR:='E_LIBELLE' ;
If OkCBRefCEL Then
  BEGIN
  NomChampRefECR:=GetControlText('COMBOREFECR') ;
  If NomChampRefECR='E_LIBELLE' Then NomChampRefECR:='E_REFINTERNE' ;
  END Else
  BEGIN
  if comboRef.enabled then
    case comboRef.ItemIndex of
      0: NomChampRefECR := 'E_LIBELLE';
      1: NomChampRefECR := 'E_REFEXTERNE';
      2: NomChampRefECR := 'E_REFINTERNE';
      3: NomChampRefECR := 'E_REFLIBRE';
      4: NomChampRefECR := 'E_NUMTRAITECHQ';
    end
  else NomChampRefECR := 'E_REFINTERNE';
  END ;

if TobEcriture.Detail.Count<>0 then
  begin
  for i:=0 to TobEcriture.Detail.Count-1 do
      begin
      TR:=TobEcriture.Detail[i] ;
      GC.Cells[SP_DATE,J]     :=TR.GetValue('E_DATECOMPTABLE');
      GC.Cells[SP_REFPOINTAGE,J]:=TR.GetValue('E_REFPOINTAGE');
      GC.Cells[SP_REFERENCE,J]:=TR.GetValue(NomChampRefECR);
      GC.Cells[SP_RGLT,J]     :=TR.GetValue('E_MODEPAIE');
      GC.Cells[SP_LIBELLE,J]  :=TR.GetValue('E_LIBELLE');
      GC.Cells[SP_DATEVAL,J]  :=TR.GetValue('E_DATEVALEUR');
      GC.Cells[SP_POSTOB,J]   :=IntToStr(i) ;
      GC.Cells[SP_POINTER,J]  :='' ;
      GC.Cells[SP_NUMPOINTER,J]  :='' ;
      GC.Cells[SP_CLASSER,J]  :=IntToStr(i+1) ;
      GC.Cells[SP_DEVAFF,J]:=TR.GETVALUE('E_DEVISE') ;
      GC.Cells[SP_DEVISE,J]:=TR.GETVALUE('E_DEVISE') ;
      GC.Cells[SP_SOLDE,J]:=PutMontant(TR.GetValue('E_'+C_DEBIT) -TR.GetValue('E_'+C_CREDIT),DecCompta) ;
      inc(J) ;
      end;
  end;
if J=1 then inc(J) ;
GC.RowCount:=J ;
MttC:=0 ;
AfficheMontant(TTotalC,0);
end;
(*******************************************************************************



        Methodes liées aux traitements du pointage



*******************************************************************************)
procedure TOF_PointeReleve.PointeManu(Sender: TObject) ;
begin
nk:=1 ;
if PointageAuto then begin
  if THGrid(Sender).Cells[SP_POINTER,THGrid(Sender).Row]=CARPOINTER then begin
      CocheDecocheAuto(StrToInt(THGrid(Sender).Cells[SP_NUMPOINTER,THGrid(Sender).Row]))
    end
    else begin
    AfficheMsg(1,'','');
    end
  end
  else
    CocheDecoche(THGrid(Sender))
end;

procedure TOF_PointeReleve.LanceLePointageAuto ;
var
  LM    : T_D ;
  LP,LB : T_I ;
  Solde : double ;
  Infos : REC_AUTO ;
  i, j  : integer ;
  St    : String ;
  pos   : Integer ;
  {$IFDEF TRESO}
  TobR  : TOB;
  {$ENDIF}
begin
  if GB.Cells[SP_DATE,1]='' then exit ;
  if PasDePointagePossible then exit ;
  PointageAuto:=true ;
  PointeAvecListeNonComplete :=false ;
  Infos.Decim:=NbDecDev ;
  GC.Visible:=false ;

{$IFDEF TRESO}
  if CDate.Checked then	ToleranceDate := TDate.Value
                   else	ToleranceDate := 0;

  Infos.Decim:=NbDecDev ;
  Infos.Temps:=0;

  if CCombinatoire.Checked then begin
    Infos.Nival := TNiveau.Value - 1;
    Infos.Temps := 30;
    Infos.Unique:= False;
  end
  else begin
    Infos.Nival := -1;
    Infos.Temps := MaxTempo;
    Infos.Unique:= True;
  end ;

  St := TraduireMemoire('Traitement Ecriture relevé n° ');
  nk:=0 ; {n° de pointage}
  for i:=1 to GB.RowCount-1 do begin
    SetControlCaption('TRAITEMENT', St + IntToStr(i) + ' s/ ' + IntToStr(GB.RowCount-1));
    Application.ProcessMessages ;
    GB.Row := i ;
    TobR := TobReleve.Detail[StrToInt(GB.Cells[SP_POSTOB, i])];
    Solde := Arrondi(TobR.GetValue('CEL_' + A_CREDIT) - TobR.GetValue('CEL_' + A_DEBIT), DecReleve) ;
    Infos.NbD := ConstruitListe(LM, LP, LB, Solde);
    if LettrageAuto(Solde,LM,LP,Infos) = 1 then begin
      {JP 08/03/04 : FQ Tréso 10020 : Cela évitera que toutes les lignes soient à zéro}
      Inc(nk);
      GB.Row:=i;
      CocheDecoche(GB) ;
      for j:=1 to Infos.NbD do
        if LP[j-1]<>0 then begin
          GC.Row := LB[j-1] ;
          CocheDecoche(GC);
          GB.Cells[SP_ROW,GB.Row] := GB.Cells[SP_ROW,GB.Row]+IntToStr(GC.Row)+';';
        end;
    end;
  end;

{$ELSE}
  ActiveStop(TRUE) ;
  ActiveProgessImmo(TRUE,Ecran,Self,TRUE) ;
  SetControlVisible('_LB1',FALSE) ;    
  SetControlVisible('_PASSAGE',FALSE) ;
  SetControlVisible('_LB3',FALSE) ;
  SetControlVisible('_LB4',FALSE) ;
  AfficheCR(traduireMemoire('Pointées'),'0') ;

  if CCombinatoire.Checked then begin
    Infos.Nival:=TNiveau.Value-1 ; Infos.Temps:=30 ; Infos.Unique:=false ;
    end
    else begin
    Infos.Nival:=-1 ; Infos.Temps:=MaxTempo  ; Infos.Unique:=true ;
    end;
  St:=TraduireMemoire('Traitement Ecriture relevé n° ') ;
  nk:=0 ; // n° de pointage
  for i:=1 to GB.RowCount-1 do begin
    SetControlCaption('TRAITEMENT',St+IntToStr(i)+' s/ '+IntToStr(GB.RowCount-1)) ;
    Application.ProcessMessages ;
    GB.Row:=i ;
    Solde:=Arrondi(TobReleve.Detail[i-1].GetValue('CEL_'+A_CREDIT) - TobReleve.Detail[i-1].GetValue('CEL_'+A_DEBIT), DecReleve) ;
    Infos.NbD:=ConstruitListe(LM,LP,LB,Solde) ;
    If HalteAuFeu Then Break ;
    if LettrageAuto(Solde,LM,LP,Infos)=1 then begin
      nk:=nk+1 ;
      GB.Row:=i; CocheDecoche(GB) ;
      // Récupération de la date de relevé
      pos := StrToInt(GB.Cells[SP_POSTOB,GB.Row]) ;
      gdDateReleve := TobReleve.Detail[pos].GetValue('EE_DATESOLDE') ;
      for j:=1 to Infos.NbD do
        if LP[j-1]<>0 then begin
          GC.Row:=LB[j-1] ;
          CocheDecoche(GC);
          GB.Cells[SP_ROW,GB.Row]:=GB.Cells[SP_ROW,GB.Row]+IntToStr(GC.Row)+';';
        end;
      AfficheCR(traduireMemoire('Pointées'),IntToStr(NbCoche)) ;
      end;
    end;
  ActiveStop(FALSE) ;
  If HalteAuFeu Then
    BEGIN
    PgiError('Traitement interrompu') ;
    END ;

  ActiveProgessImmo(FALSE,Ecran,Self) ;

{$ENDIF}
  GC.Visible :=true ;

  if PointeAvecListeNonComplete then AfficheMsg(21,'','');

  if NbCoche<>0 then begin
                     BValider.Enabled := true;
                     AfficheMsg(12,IntToStr(NbCoche),'');
                end
                else begin
                     PointageAuto:=false ;
                     BValider.Enabled := false;
                     AfficheMsg(15,'','');
                end;
end;


function  TOF_PointeReleve.ConstruitListe(var LM : T_D;var LP,LB: T_I; MM : Double) : integer ;
var i,j,k : integer ; Mtt : double ; LaTob : Tob ;
begin
result:=0;i:=0 ;
if GC.Cells[SP_DATE,1]='' then exit;
FillChar(LM,Sizeof(LM),#0) ; FillChar(LP,Sizeof(LP),#0) ;
for j:=1 to GC.RowCount-1 do
  begin
  GC.Row:=j ;
  LaTob:=TobEcriture.Detail[j-1] ;
  if (LaTob.GetValue('E_REFPOINTAGE')='') and EcritureSelectionnee then
    begin
    Mtt:=LaTob.GetValue('E_'+C_DEBIT) - LaTob.GetValue('E_'+C_CREDIT) ;
    if (((MM<=0) and (Mtt>=0)) or ((MM>=0) and (Mtt<=0))) then continue ; //gv débit comptable avec crédit bancaire et inversement.
    k:=i ;
    if k=MaxDroite-1 then begin PointeAvecListeNonComplete:=true ; exit ; end;
    while k<>0 do // classement des montants par ordre décroissant
      begin
      if Mtt<LM[k-1] then begin LM[k]:=LM[k-1] ; LP[k]:=LP[k-1] ; LB[k]:=LB[k-1] ; end
                     else break;
      dec(k) ;
      end;
    LM[k]:=Mtt ; LP[k]:=0 ; LB[k]:=j ;
    inc(i) ;
    end;
  end;
result:=i ;
end;


procedure TOF_PointeReleve.RempliLesTobs ;
var Q:TQuery ;
    Sql : string ;
    Cod, st, LeWhere,LeWheredevise,Journal,NomChampRefBQ : String ;
begin
if TobReleve<>nil then begin TobReleve.free; TobReleve:=nil; end;
TobReleve:=TOB.create('_RELEVE',nil,-1) ;
LeWhereDevise:='EE_DEVISE="'+DeviseReleve+'" AND ' ;
NomChampRefBQ:='CEL_REFPIECE' ; If OkCBRefCEL Then NomChampRefBQ:=GetControlText('COMBOREFBQ') ;
If (NomChampRefBQ='CEL_LIBELLE') Then NomChampRefBQ:='CEL_REFPIECE' ;

Sql := 'SELECT CEL_CODEAFB,CEL_CODERAPPRO,CEL_CREDITDEV,CEL_CREDITEURO,CEL_DATEOPERATION, '+
       'CEL_DATEPOINTAGE,CEL_DATEVALEUR,CEL_DEBITDEV,CEL_DEBITEURO,CEL_DISPONIBLE,CEL_EXONERE, '+
       'CEL_GENERAL,CEL_IMO,CEL_LIBELLE,CEL_LIBELLE1,CEL_LIBELLE2,CEL_LIBELLE3,CEL_NATUREINTERNE, '+
//GP       'CEL_NUMLIGNE,CEL_NUMRELEVE,CEL_REFORIGINE,CEL_REFPIECE,CEL_REFPOINTAGE,CEL_RIB,CEL_VALIDE, '+
       'CEL_NUMLIGNE,CEL_NUMRELEVE,CEL_REFORIGINE,'+NomChampRefBQ+', CEL_REFPOINTAGE,CEL_RIB,CEL_VALIDE, '+
       'EE_GENERAL, EE_NUMRELEVE, EE_DEVISE, EE_DATEPOINTAGE, EE_REFPOINTAGE, D_DEVISE, D_CODEISO '+
       ', EE_DATESOLDE '+ // Ajout de la date de relevé pour traitement
       'FROM EEXBQ LEFT JOIN EEXBQLIG ON CEL_GENERAL=EE_GENERAL AND CEL_NUMRELEVE=EE_NUMRELEVE '+
       'LEFT JOIN DEVISE ON EE_DEVISE=D_CODEISO WHERE '+
       LeWhereDevise +
       'CEL_GENERAL="'+GetControlText('CBANQUE')+'" AND CEL_DATEPOINTAGE = "'+UsDateTime(idate1900)+'"' ; {JP 13/08/04 : FQ 13167}

if StrtoInt(CRefReleve.value)>0 then Sql :=Sql + ' AND CEL_REFPOINTAGE="'+CRefReleve.items[StrToInt(CRefReleve.Value)]+'"' ;
LeWhere:='' ;
if (MultiAFB.Text <> '<<Tous>>') and  (MultiAFB.Text <> '') then
  begin
  LeWhere:='' ;
  Cod := MultiAFB.Text ;
  St := ReadTokenpipe(Cod,';') ;
  LeWhere:=' AND (CEL_CODEAFB="'+st+'" ' ;
  St := ReadTokenpipe(Cod,';') ;
  while st <> '' do
    begin
    if length(st)=1 then st:='0'+st ;
    LeWhere:= LeWhere + ' OR CEL_CODEAFB="'+st+'" ' ;
    St := ReadTokenpipe(Cod,';') ;
    end;
  LeWhere:=LeWhere+')' ;
  end;
Sql:=Sql+LeWhere+' ORDER BY CEL_DATEOPERATION' ;
Q:=OpenSQL(SQL,true) ;
try
  TobReleve.LoadDetailDB('EEXBQLIG','','',Q,false,true) ;
  if TobReleve.Detail.Count > 0 then TobReleve.Detail[0].AddChampSup('POINTE',true) ;
finally
  ferme(Q);
  end;

if TobEcriture<>nil then begin TobEcriture.free; TobEcriture:=nil; end;
TobEcriture:=TOB.create('_ECRITURE',nil,-1) ;

  {JP 20/12/04 : DeviseCompta contient D_CODEISO : cf. RechercheDevise(
  LeWhereDevise:='D_DEVISE="'+DeviseCompta+'" AND ' ;}
  LeWhereDevise:='D_CODEISO = "' + DeviseCompta + '" AND ';

// gv ModeTiers
if ModeTiers.Checked then
  begin
  Q:=OpenSql('SELECT J_JOURNAL FROM JOURNAL WHERE J_CONTREPARTIE="' + GetControlText('CBANQUE') + '"', True);{JP 13/08/04 : FQ 13167}
  if Not Q.EOF then journal := Q.Findfield('J_JOURNAL').AsString ;
  ferme(Q) ;
  end;
// Gv ModeTiers //Rajout 06/05/06 e_modesasie pour consult de piece
Sql:='SELECT E_JOURNAL, E_EXERCICE, E_DATECOMPTABLE, E_NUMEROPIECE, E_NUMLIGNE, E_NUMECHE, E_QUALIFPIECE, ' +
     'E_LIBELLE, E_REFEXTERNE, E_REFINTERNE, E_REFLIBRE, E_NUMTRAITECHQ, E_MODEPAIE, E_MODESAISIE, '+
     'E_DEVISE, E_REFPOINTAGE, E_DATEVALEUR, E_GENERAL, E_DATEPOINTAGE, '+
     'E_DEBIT,E_CREDIT,E_DEBITDEV,E_CREDITDEV, D_CODEISO, D_DEVISE '+
     'FROM ECRITURE LEFT JOIN DEVISE ON E_DEVISE=D_DEVISE ' ;
if ModeTiers.Checked then
     Sql:=Sql+'WHERE E_GENERAL <>"'+ GetControlText('CBANQUE') + '" AND E_JOURNAL ="'+Journal+'"' else // GV ModeTiers
     Sql:=Sql+'WHERE E_GENERAL ="' + GetControlText('CBANQUE') + '" AND ' ; {JP 13/08/04 : FQ 13167}
     Sql:=Sql+LeWhereDevise +
     'E_DATEPOINTAGE = "'+UsDateTime(idate1900)+'" AND ' +
     'E_DATECOMPTABLE >= "'+ UsDateTime(StrToDate(DateDe.Text)) +'" AND ' +
     'E_DATECOMPTABLE <= "'+ UsDateTime(StrToDate(DateJusqua.Text)) +'" ';

LeWhere:='' ;

{JP 09/02/04 : je pense qu'il s'agissait d'un copier/coller mal modifier}
//if (MultiModePaie.Text <> '<<Tous>>') and  (MultiAFB.Text <> '') then
if (MultiModePaie.Text <> '<<Tous>>') and  (MultiModePaie.Text <> '') then
  begin
  LeWhere := '' ;
  Cod := MultiModePaie.Text ;
  St := ReadTokenpipe(Cod,';') ;
  LeWhere:=' AND (E_MODEPAIE="'+st+'" ' ;
  St := ReadTokenpipe(Cod,';') ;
  while st <> '' do
    begin
    LeWhere:= LeWhere + ' OR E_MODEPAIE="'+st+'" ' ;
    St := ReadTokenpipe(Cod,';') ;
    end;
  LeWhere:=LeWhere+')' ;
  end;

  LeWhere:=LeWhere+' AND E_QUALIFPIECE="N" AND (E_ECRANOUVEAU="H" Or E_ECRANOUVEAU="N") ' ; // Ignorance des écritures cloturéés
  Sql:=Sql+LeWhere+' ORDER BY E_DATECOMPTABLE' ;
  Q:=OpenSQL(Sql,true) ;
  Try
    TobEcriture.LoadDetailDB('ECRITURE','','',Q,false,true) ;
    if Tobecriture.detail.Count > 0 then TobEcriture.Detail[0].AddChampSup('POINTE',true) ;
  finally
    ferme(Q);
  end;
end;

(*******************************************************************************



         Methodes liées a l'affichage du pointage



*******************************************************************************)

procedure TOF_PointeReleve.CocheDecocheAuto(Ligne : integer) ;
var
  i: integer ;
begin
for i:= 0 to GB.Rowcount-1 do
  begin
  if GB.cells[SP_NUMPOINTER,i]=intToStr(Ligne) then begin GB.row:=i ; CocheDecoche(GB) ; end;
  end;
for i:= 0 to GC.Rowcount-1 do
  begin
  if GC.cells[SP_NUMPOINTER,i]=IntToStr(Ligne) then begin GC.row:=i ; CocheDecoche(GC) ; end;
  end;
end;

procedure TOF_PointeReleve.CocheDecoche(G : THGrid) ;
begin
if G=GB then CocheDecocheGB else CocheDecocheGC ;
  BValider.Enabled := ((NbCocheC>0) or (NbCocheB>0)) and ((abs(MttB) - abs(MttC)) = 0);  // FQ 13077
if NbCoche=0 then begin
  LesVisibles(true);
  RefReleve:='';
  PointageManu:=false ;
  PointageAuto:=false ;
  end
  else begin
  LesVisibles(false) ;
  if Not PointageAuto then PointageManu:=true ;
  end;
  if PointageManu and (G=GB) and (NbCocheB=0) then begin
     CocheDecocheGB;
     BValider.Enabled := (NbCocheC<>0) and (NbCocheB<>0);
     OnCancel;
  end;
end;

{JP 03/03/05 : FQ TRESO 10213 : le code de VL pour la fiche COMPTA 13077 (pointage manuel d'écritures comptables
               entre elles) n'était pas sans posée de problèmes sur le rapprochement automatique multi-relevés :
               les références de pointages des différents relevés se retrouvaient idnetiques et la table eexbqlig
               était mise à jour avec la nouvelle référence : il devenait impossible de dépointer les écritures.
               Cependant, je ne suis pas sûr du tout que ma modification n'entrainera pas d'effets de bord non plus !!!!}
procedure TOF_PointeReleve.CocheDecocheGB ;
var
  Mtt : double ;
  TR : TOB ;
  pos: integer ;
begin
pos := StrToInt(GB.Cells[SP_POSTOB,GB.Row]) ;
TR  := TobReleve.Detail[pos] ;
Mtt := TR.GetValue('CEL_'+A_CREDIT) - TR.GetValue('CEL_'+A_DEBIT);
if RefReleve = '' then RefReleve := TR.GetValue('CEL_REFPOINTAGE') ;
{JP 03/03/05 : FQ TRESO 10213 : réactivation de la condition, sinon le code de VL (FQ 13077) est
               exécuté systématiquement et si l'on a plusieurs références de pointage elles sont
               écrasées avec la dernière ...}
if (TR.GetValue('CEL_REFPOINTAGE') = RefReleve) or (PointageAuto) then begin
  if GB.Cells[SP_POINTER,GB.Row]=CARPOINTER then begin
    // Décoche la ligne
    GB.Cells[SP_POINTER,GB.Row]:=' ';
    GB.Cells[SP_NUMPOINTER,GB.Row]:='' ;
    GB.Cells[SP_ROW,GB.Row]:='';
    TR.PutValue('POINTE','') ;
    TR.PutValue('CEL_DATEPOINTAGE',idate1900) ;
    MttB:=arrondi(MttB-Mtt,4) ;
    NbCoche:=NbCoche-1 ;
    NbCocheB:=NbCocheB-1 ;
    if NbCocheB=0 then gdDateReleve := V_PGI.DateEntree ;
    end
  else begin
    // Coche la ligne
    GB.Cells[SP_POINTER,GB.Row]:=CARPOINTER;
    GB.Cells[SP_NUMPOINTER,GB.Row]:=IntToStr(nk);
    // Date de pointage = Date de relevé
    gdDateReleve := TR.GetValue('EE_DATESOLDE') ;
    TR.PutValue('CEL_DATEPOINTAGE', gdDateReleve ) ;
    TR.PutValue('POINTE','X') ;
    MttB:=arrondi(MttB+Mtt,4) ;
    NbCoche:=NbCoche+1 ;
    NbCocheB:=NbCocheB+1 ;
  end
end;                                                                                    
(* JP 11/04/05 : FQ TRESO 10236  : Je mets tout ce code en commentaire car à mon avis il n'y a
                 aucune raison de modifier la référence de pointage ou alors il faut aller plus
                 loin dans la réflexion
{JP 03/03/05 : FQ TRESO 10213 : ... suite, si pointage manuel uniquement}
else
  // FQ 13077
  // Si c'est un relevé différent
  if (TR.GetValue('CEL_REFPOINTAGE')<>RefReleve) then begin
    // On parcours la tob des relevés
    for i := 0 to TobReleve.Detail.count-1 do begin
      // Sauf celle sur laquelle on se trouve en ce moment
      if (i <> pos) then begin
        // et on regarde si elle est sélectionnée
        if (TobReleve.Detail[i].GetValue('POINTE') = 'X') then begin
          // On compare alors les dates d'opération
          if (TobReleve.Detail[i].GetDateTime('CEL_DATEOPERATION') >= TR.GetDateTime('CEL_DATEOPERATION')) then begin
            // et on y met la référence de pointage la plus récente
            TR.PutValue('CEL_REFPOINTAGE', TobReleve.Detail[i].GetValue('CEL_REFPOINTAGE'));
            // sans oublier la grille
            GB.Cells[SP_REFPOINTAGE,GB.Row]:= TobReleve.Detail[i].GetValue('CEL_REFPOINTAGE');
          end
          else begin
            ToBReleve.Detail[i].PutValue('CEL_REFPOINTAGE', TR.GetValue('CEL_REFPOINTAGE'));
            {JP 03/03/05 : FQ TRESO 10213 : i + 1, sinon le titre de la colonne se véra affecté la référence de pointage}
            GB.Cells[SP_REFPOINTAGE,i+1]:= TR.GetValue('CEL_REFPOINTAGE');
          end;
        end;
      end;
    end;
  end;
*)
//end
//else AfficheMsg(10,'','');
AfficheMontant(TTotalB,MttB) ;
GB.Invalidate ;
end;

procedure TOF_PointeReleve.CocheDecocheGC ;
var Mtt : double ;TR : TOB ; pos: integer ;
begin
 {JP 07/07/05 : en attendant l'analyse, on bloque} 
 if NbCoche=0 then begin AfficheMsg(9,'',''); exit; end; // gv 110402
pos:=StrToInt(GC.Cells[SP_POSTOB,GC.Row]) ;
TR:=TobEcriture.Detail[Pos];

{JP 20/12/04 : FQ 15078 : Pour bien faire il faudrait peut-être modifier le message qui
               n'est pas très explicite si on pointe les écritures entre elles et que l'on
               n'a pas de relevé : un message du genre "Pas de référence de pointage" serait
               plus explicite que "Devises incompatible" !! à voir ...}
if (GC.Cells[SP_DEVAFF, GC.Row]<>GB.Cells[SP_DEVAFF, GB.Row]) then
  begin
  {JP 12/01/04 : FQ 15078 : on peut penser que si la devise et la référence de pointage sont vides
                 c'est que l'on n'a pas de lignes de relevé. Quoiqu'il en soit, sans réfénce de pointage
                 on ne peut pas lancer le traitement}
  if (GB.Cells[SP_DEVAFF, GB.Row] = '') and
     (GB.Cells[SP_REFPOINTAGE, GB.Row] = '') then AfficheMsg(25,'','') {Message plus explicite !!!}
                                             else AfficheMsg(13,'','');
  exit;
  end;

{JP 13/08/2004 FQ 13403 Deb - Cred et non l'inverse cf RempliCompta}
Mtt := TR.GetValue('E_' + C_DEBIT) - TR.GetValue('E_' + C_CREDIT) ;

if GC.Cells[SP_POINTER,GC.Row]=CARPOINTER then
  begin
  GC.Cells[SP_NUMPOINTER,GC.Row]:=' ';
  GC.Cells[SP_POINTER,GC.Row]:=' ';
  GC.Cells[SP_REFPOINTAGE,GC.Row]:=' ';
  GC.Cells[SP_ROW,GC.Row]:='';
  TR.PutValue('E_REFPOINTAGE','') ;
  TR.PutValue('POINTE','') ;
  TR.PutValue('E_DATEPOINTAGE',idate1900) ;
  MttC := Arrondi(MttC - Mtt, 4);
  NbCoche:=NbCoche-1 ;
  NbCocheC:=NbCocheC-1 ;
  end else
  begin
  GC.Cells[SP_POINTER,GC.Row]:=CARPOINTER;
  GC.Cells[SP_NUMPOINTER,GC.Row]:=IntToStr(nk);
  GC.Cells[SP_REFPOINTAGE,GC.Row]:=GB.Cells[SP_REFPOINTAGE,GB.Row];
  GC.Cells[SP_ROW,GC.Row]:=IntToStr(nk);
  TR.PutValue('E_REFPOINTAGE',GB.Cells[SP_REFPOINTAGE,GB.Row]) ;
  TR.PutValue('POINTE','X') ;
  // Date de pointage = Date du relevé
  TR.PutValue('E_DATEPOINTAGE', gdDateReleve) ;
  MttC := Arrondi(MttC + Mtt, 4); 
  NbCoche:=NbCoche+1 ;
  NbCocheC:=NbCocheC+1 ;
  end;
AfficheMontant(TTotalC,Mttc) ;
GC.Invalidate ;
end;

//Ajout ----------------CTW le 17/10/2000
{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 30/04/2004
Modifié le ... : 30/04/2004
Description .. : Modification ....
Suite ........ : avant la fonction ne fesait que rester la reference de 
Suite ........ : pointage des lignes d'ecriture ....
Suite ........ : ce qui laissait les ecritures a moitier pointe et les ligne de 
Suite ........ : releve pointé !
Suite ........ : 
Suite ........ : Fiche qualité n° 13492
Mots clefs ... : 
*****************************************************************}
procedure TOF_PointeReleve.RemetRefPointageABlanc;
var i: integer;
begin
    if ((TobEcriture <> nil) and (TobEcriture.Detail.Count <> 0)) then
    begin
        for i:=0 to TobEcriture.Detail.Count-1 do
        begin
            if (TobEcriture.Detail[i].GetValue('E_REFPOINTAGE') <> '')         then TobEcriture.Detail[i].PutValue('E_REFPOINTAGE','');
            if (TobEcriture.Detail[i].GetValue('E_DATEPOINTAGE') <> idate1900) then TobEcriture.Detail[i].PutValue('E_DATEPOINTAGE',idate1900);
            if (TobEcriture.Detail[i].GetValue('POINTE') <> '')                then TobEcriture.Detail[i].PutValue('POINTE','');
        end;
    end;
    if ((TobReleve <> nil) and (TobReleve.Detail.Count <> 0)) then
    begin
        for i:=0 to TobReleve.Detail.Count-1 do
        begin
            if (TobReleve.Detail[i].GetValue('CEL_DATEPOINTAGE') <> idate1900) then TobReleve.Detail[i].PutValue('CEL_DATEPOINTAGE',idate1900);
            if (TobReleve.Detail[i].GetValue('POINTE') <> '')                  then TobReleve.Detail[i].PutValue('POINTE','');
        end;
    end;
end;

function TOF_PointeReleve.ExisteMoyenPaiement: boolean;
var
   i: integer;
   s: string;
begin
  result := true;
  for i := 0 to TobReleve.Detail.Count-1 do begin
    s := TobReleve.Detail[i].GetValue('CEL_CODEAFB');
    if GetLeModeRglt(s) = '' then begin
      result := false;
      AfficheMsg(17,'"'+s+'"','');
      break;
    end;
  end;
end;

procedure TOF_PointeReleve.OnDateExit(Sender: TObject);
begin
  if THEDIT(Sender).Name = 'DATEDE' then
  begin
    if StrToDate(DateDe.Text) > StrToDate(DateJusqua.Text) then begin
      DateJusqua.Text := DateDe.Text;
      exit;
    end;
  end
  else if THEDIT(Sender).Name = 'DATEJUSQUA' then begin
    if (Strtodate(DateJusqua.Text) < VH^.Encours.Deb) then begin
      AfficheMsg(19,'','');
      SetFocusControl('DATEJUSQUA');
      exit;
    end;
    if StrToDate(DateDe.Text) > StrToDate(DateJusqua.Text) then begin
      AfficheMsg(18,'','');
      SetFocusControl('DATEJUSQUA');
    end;
  end;
end;

procedure TOF_PointeReleve.OnPopUpZoom(Sender: TObject);
begin
  vpiece.Enabled := GC.Cells[0,1] <> '';
end;

procedure TOF_PointeReleve.OnClickVcompte(Sender: TObject);
begin
  {$IFNDEF TRESO}
  FicheGene(nil, '', GetControlText('CBANQUE'), taConsult, 0);{JP 13/08/04 : FQ 13167}
  {$ENDIF}
end;

procedure TOF_PointeReleve.BHelpClick(Sender: TObject);
begin
  CallHelpTopic(Ecran) ;
end;

procedure TOF_PointeReleve.OnFormResize(Sender: TObject);
begin
    // BPY le 08/12/2003 => fiche 13075 : vas de paire avec une bob !
    pcpta.Height := pgene.ClientHeight div 2;
    // fin BPY
end;

procedure TOF_PointeReleve.OnClickBParam(Sender: TObject);
begin
{$IFDEF EAGLCLIENT}
    // TODO
{$ELSE}
     ExtraitBanquaire;
{$ENDIF}
end;

procedure TOF_PointeReleve.OnClickBImprimer(Sender: TObject);
var
  Mr : TModalResult;
begin
  {JP 17/08/05 : FQ COMPTA 15433 : Impression en eAgl}
  Mr := AfficheMsg(24, '', '');
       if (Mr = mrYes) then ImprimerGrille(GB, 'Pointage automatique des relevés bancaires - Relevés bancaires')
  else if (Mr = mrNo ) then ImprimerGrille(GC, 'Pointage automatique des relevés bancaires - Ecritures comptables');
end;

Procedure TOF_PointeReleve.BtnAFBOnClick (Sender : TObject) ;
begin
  {$IFDEF TRESO}
  TRLanceFiche_CIB('TR','TRCIB','','','ACTION=MODIFICATION;' + tc_CIB);
  {$ELSE}
  ParamCodeAFB;
  {$ENDIF}
end;

{***********A.G.L.***********************************************
Auteur  ...... : Stéphane BOUSSERT
Créé le ...... : 17/07/2003
Modifié le ... : 18/07/2003
Description .. : Effectue la recherche sur la liste sélectionnée
Suite ........ : (Fiche 12033)
Mots clefs ... : 
*****************************************************************}
procedure TOF_PointeReleve.OnFindAFindDialog(Sender: TObject);
begin
  if listeRech = GB.Name
    then Rechercher(GB, AFindDialog, FFindFirst)  // liste des relevés
    else Rechercher(GC, AFindDialog, FFindFirst); // Liste des écritures
end;

{GP 16/03/2008 : Rappro troc de lille : choix de la ref eexbqlig}
Function TOF_PointeReleve.OkCBRefCEL : Boolean ;
BEGIN
Result:=(GetControl('COMBOREFBQ')<>NIL) And (GetControl('COMBOREFECR')<>NIL) And (GetControlText('CBPARAM2')='X') ;
END ;

{***********A.G.L.***********************************************
Auteur  ...... : Stéphane BOUSSERT
Créé le ...... : 18/07/2003
Modifié le ... : 18/07/2003
Description .. : Initialisation de la fenêtre de recherche sur les listes
Suite ........ : (Fiche 12033)
Mots clefs ... :
*****************************************************************}
procedure TOF_PointeReleve.OnArgument(S: string);
var
  Composants : TControlFiltre; {JP 13/08/2004}
  CB : tCheckBox ;
  Bt : TToolBarButton97 ;
  HG : thGrid ;
  i : Integer ;
begin
  // Ajout SBO pour fiche 12033 : Bouton rechercher
  AFindDialog := TFindDialog.Create(Ecran);
  AFindDialog.OnFind := OnFindAFindDialog;
  SetControlVisible('BRECHERCHER', True) ;
  {$IFDEF TRESO}
  {JP 11/03/04 : mise en place de la création d'écriture à partir des extraits bancaires}
  if Assigned (TPopupMenu(GetControl('PPECRITURE')).Items[0]) then
    TPopupMenu(GetControl('PPECRITURE')).Items[0].OnClick := GenererEcriture;

  {JP 12/03/04 : Integration en compta des éventuelles écritures crées}
  TobCompta := TOB.Create('***', Nil, -1);
  {$ENDIF}

  {JP 13/08/2004 : Nouvelle gestion des filtres}
  Composants.PopupF   := TPopUpMenu      (Getcontrol('POPF'));
  Composants.Filtres  := THValComboBox   (Getcontrol('FFILTRES'));
  Composants.Filtre   := TToolBarButton97(Getcontrol('BFILTRES'));
  Composants.PageCtrl := TPageControl    (Getcontrol('ONGLETHAUT'));
  {$IFDEF TRESO}
  ObjFiltre := TObjFiltre.Create(Composants, 'TRTOFRPOIN');
  SetControlVisible('BZOOM', False); {FQ TRESO 10190 : 23/11/04, on cache pour ne tirer la compta}
  {$ELSE}
  ObjFiltre := TObjFiltre.Create(Composants, 'TOFRPOIN');
  {$ENDIF}

{GP 16/03/2008 : Rappro troc de lille : choix de la ref eexbqlig}
If OkCBRefCEL Then
  BEGIN
  If GetControlTExt('COMBOREFBQ')=''  Then SetControlText('COMBOREFBQ','CEL_REFPIECE') ;
  If GetControlTExt('COMBOREFECR')=''  Then SetControlText('COMBOREFBQ','E_LIBELLE') ;
  END ;
CB:=tCheckBox(GetControl('CBPARAM2')) ; If CB<>NIL Then CB.OnClick:=CBPARAM2OnClick ;
CB:=tCheckBox(GetControl('CBOKPOSBQ')) ; If CB<>NIL Then CB.OnClick:=CBOKPOSBQOnClick ;
CB:=tCheckBox(GetControl('CBOKPOSECR')) ; If CB<>NIL Then CB.OnClick:=CBOKPOSBQOnClick ;
CBPARAM2OnClick(NIL) ;
Bt:=TToolBarButton97(GetControl('BSTOP')) ; If Bt<>Nil Then Bt.OnClick := BStopClick ;
Bt:=TToolBarButton97(GetControl('BVISULIB')) ; If Bt<>Nil Then Bt.OnClick := BVISULIBClick ;
Bt:=TToolBarButton97(GetControl('BRENVOIECR')) ; If Bt<>Nil Then Bt.OnClick := BRENVOIECRClick ;
Bt:=TToolBarButton97(GetControl('BRENVOIBQ')) ; If Bt<>Nil Then Bt.OnClick := BRENVOIBQClick ;
HG:=tHgrid(GetControl('GCOMPTA')) ; If HG<>NIL Then HG.OnRowEnter:=GCRowEnter ;
HG:=tHgrid(GetControl('GBANQUE')) ; If HG<>NIL Then HG.OnRowEnter:=GBRowEnter ;
HG:=THGrid(GetControl('GLIBNUM')) ; If HG<>NIL Then For i:=0 To HG.ColCount-1 Do HG.Cells[i,0]:=IntToStr((i+1) Mod 10) ;
end;


{***********A.G.L.***********************************************
Auteur  ...... : Stéphane BOUSSERT
Créé le ...... : 18/07/2003
Modifié le ... : 18/07/2003
Description .. : Ouvre la fenêtre de recherche sur les liste
Suite ........ : (Fiche 12033)
Mots clefs ... :
*****************************************************************}
procedure TOF_PointeReleve.OuvrirRecherche;
begin
  FFindFirst := True;
  AFindDialog.Execute;
end;

{JP 13/08/04 : FQ 13167, problème de la gestion du confidentiel
{---------------------------------------------------------------------------------------}
function TOF_PointeReleve.GetPlus : string;
{---------------------------------------------------------------------------------------}
// GP le 27/05/2008 : rajout de Result:=Result+...
begin
  Result := ' G_POINTABLE="X" ';
  if V_PGI.Confidentiel = '0' then
    Result := Result+' AND (G_CONFIDENTIEL = "0" OR G_CONFIDENTIEL = "-")'
  else
    Result := Result+' AND (G_CONFIDENTIEL <= "' + V_PGI.Confidentiel +
              '" OR G_CONFIDENTIEL = "X" OR G_CONFIDENTIEL = "-")';
end;

{$IFDEF TRESO}
{---------------------------------------------------------------------------------------}
procedure TOF_PointeReleve.CRegAccroOnClick(Sender: TObject) ;
{---------------------------------------------------------------------------------------}
begin
  CDate.Enabled := CRegAccro.Checked and CRegAccro.Enabled;
  OnClickCDate(Nil);
  CNumCheque.Enabled := CDate.Enabled;
  OnClickCNumCheque(Nil);
end;

{On peut appliquer les règles d'accro que si le mode de paiement correspond
{---------------------------------------------------------------------------------------}
procedure TOF_PointeReleve.CMODEREGOnClick(Sender: TObject);
{---------------------------------------------------------------------------------------}
begin
  CRegAccro.Enabled := CModeReg.Checked;
  CRegAccroOnClick(Nil);
end;

{JP 10/03/04 : Génération des écritures de compta et de tréso}
{---------------------------------------------------------------------------------------}
procedure TOF_PointeReleve.GenererEcriture(Sender: TObject);
{---------------------------------------------------------------------------------------}
var
  Cle : string;
  str : string;
  Q   : THQuery;
  T   : TQuery;
  n   : Integer;
  d   : Double;
begin
  {Constitution de la chaine de paramètres : Compte, "RAPPRO", Date Opération, Libellé, Montant, Date valeur}
  n := GB.Row;
  Cle := GetControlText('CBANQUE');
  if Pos('<<', Cle) > 0 then Cle := '';
  {JP 22/03/04 : C'est le sens du flux saisi qui détermine le signe de l'écriture de trésorerie}
  d := Abs(Valeur(GB.CellValues[SP_SOLDE, n]));
  str := Cle + ';RAPPRO;' + GB.CellValues[SP_DATE, n] + ';' + GB.CellValues[SP_LIBELLE, n] +
                      ';' + FloatToStr(d) + ';' + GB.CellValues[SP_DATEVAL, n] + ';';

  {JP 08/08/05 : FQ 10088 : On reprend le code CIB}
  Str := Str + TobReleve.Detail[StrToInt(GB.Cells[SP_POSTOB, GB.Row])].GetString('CEL_CODEAFB') + ';';

  {Le retour est la clé en valeur de l'écriture de trésorerie créée}
  Cle := TRLanceFiche_SaisieFlux('TR', 'TRSAISIEFLUX', '', '', str);

  {Si il n'y a pas eu d'abandon et que l'utilisateur a les droits}
  if (Cle <> '') and V_PGI.Superviseur then begin
    {1/ Rélisation de l'écriture}
    UpdatePieceStr('', '', Cle, 'TE_NATURE', na_Realise);

    {2/ éventuelle intégration en compta}
    if not GetParamSocSecur('SO_TRINTEGAUTO', False) then Exit;

    {JP 23/04/04 : Si l'utilisateur n'a pas accès à l'écran d'intégration automatique, on ne l'autorise
                   pas à intégrer automatiquement les écritures en compta}
    if not AutoriseFonction(dac_Integration) then Exit;

    if HShowMessage('2;' + Ecran.Caption + ';Voulez-vous intégrer en comptabilité l''écriture créée ?;Q;YNC;N;N;', '', '') = mrNo then Exit;

    TobCompta.ClearDetail;
    T := OpenSQL('SELECT * FROM TRECRITURE WHERE TE_NUMTRANSAC IN (SELECT TE_NUMTRANSAC FROM TRECRITURE WHERE TE_CLEOPERATION = "' + Cle + '")', True);

    {$IFDEF EAGLCLIENT}
    Q := THQuery.Create(Ecran);
    Q.TQ := Tob.Create('', nil, -1);
    Q.TQ.LoadDetailDB('', '', '', T, False);
    {$ELSE}
    Q := THQuery(T);
    {$ENDIF}
    try
      {Générations des écritures de comptabilité à partir du record}
      PrepareEcritCompta(TobCompta, Q);

      {Maintenant que l'on a les écritures au format de la compta dans la tobCompta,
       on lance le processus d'intégration en comptabilité proprement dit}
      if TobCompta.Detail.Count > 0 then begin
        {22/07/04 : Lancement du processus proprement dit d'intégration des écritures}
        LanceIntegration(TobCompta, False, Ecran.Caption);
       end
       else
         HShowMessage('3;' + Ecran.Caption + ';L''intégration en compta n''a pas pu être réalisée;W;O;O;O', '', '');
    finally
      Ferme(T);
      {$IFDEF EAGLCLIENT}
      if Assigned(Q) then FreeAndNil(Q);
      {$ENDIF}
    end;
  end;
end;
{$ENDIF}

{JP 17/08/05 : FQ COMPTA 15433 : Impression en eAgl
{---------------------------------------------------------------------------------------}
procedure TOF_PointeReleve.ImprimerGrille(aGrille : THGrid; aTitre : string);
{---------------------------------------------------------------------------------------}
var
  M : TOB;
  F : TOB;
  n : Integer;
begin
  M := TOB.Create('$RELEVE', nil, -1);
  try
    for n := 1 to aGrille.RowCount - 1 do begin
      F := TOB.Create('$RELEVE', M, -1);
      F.AddChampSupValeur('DATE'       , aGrille.Cells[SP_DATE       , n]);
      F.AddChampSupValeur('REFPOINTAGE', aGrille.Cells[SP_REFPOINTAGE, n]);
      F.AddChampSupValeur('REFERENCE'  , aGrille.Cells[SP_REFERENCE  , n]);
      F.AddChampSupValeur('SOLDE'      , aGrille.Cells[SP_SOLDE      , n]);
      F.AddChampSupValeur('RGLT'       , aGrille.Cells[SP_RGLT       , n]);
      F.AddChampSupValeur('ADEVISE'    , aGrille.Cells[SP_DEVISE     , n]);
      F.AddChampSupValeur('LIBELLE'    , aGrille.Cells[SP_LIBELLE    , n]);
      F.AddChampSupValeur('DATEVAL'    , aGrille.Cells[SP_DATEVAL    , n]);
      F.AddChampSupValeur('POINTER'    , aGrille.Cells[SP_POINTER    , n]);
      {$IFDEF TRESO}                                
      {15/12/05 : FQ 10319 : ajout du libellé enrichi, en version trésorerie, si relevé}
      if aGrille = GB then
        F.AddChampSupValeur('LIBELLEENR', aGrille.Cells[SP_LIBENRICHI , n]);
      {$ENDIF TRESO}
    end;

    if aGrille = GB then
      LanceEtatTob('E', 'CPE', 'RA7', M, True, False, False, nil, '', aTitre, False)
    else
      LanceEtatTob('E', 'CPE', 'RP7', M, True, False, False, nil, '', aTitre, False);
  finally
    FreeAndNil(M);
  end;
end;

procedure TOF_PointeReleve.CBPARAM2OnClick(Sender:  TObject) ;
Var OkParam : Boolean ;
BEGIN
OkParam:=GetControlText('CBPARAM2')='X' ;
If Not OkParam Then
  BEGIN
  SetControlText('COMBOREFBQ','CEL_REFPIECE') ;
  SetControlText('CBOKPOSBQ','-') ;
  SetControlText('COMBOREFECR','E_LIBELLE') ;
  SetControlText('CBOKPOSECR','-') ;
  SetControlText('CBCASSE','X') ;
  END ;
SetControlEnabled('LIBREFBQ',OkParam) ;
SetControlEnabled('COMBOREFBQ',OkParam) ;
SetControlEnabled('CBOKPOSBQ',OkParam) ;
SetControlEnabled('POSBQ1',OkParam) ;
SetControlEnabled('FPOSBQ2',OkParam) ;
SetControlEnabled('POSBQ2',OkParam) ;
SetControlEnabled('LIBREFECR',OkParam) ;
SetControlEnabled('COMBOREFECR',OkParam) ;
SetControlEnabled('CBOKPOSECR',OkParam) ;
SetControlEnabled('POSECR1',OkParam) ;
SetControlEnabled('FPOSECR2',OkParam) ;
SetControlEnabled('POSECR2',OkParam) ;
SetControlEnabled('CBCASSE',OkParam) ;
//SetControlEnabled('GBBQ',OkParam) ; SetControlEnabled('GBECR',OkParam) ;
END ;

procedure TOF_PointeReleve.CBOKPOSBQOnClick(Sender:  TObject) ;
Var OkParam : Boolean ;
    St,St1 : String ;
    OkEcr,OkBQ : Boolean ;
BEGIN
If Sender=Nil Then Exit ;
St:=tcontrol(Sender).Name ; St1:='POSBQ' ;
If St='CBOKPOSECR' Then St1:='POSECR' ;
OkEcr:=GetControlText('CBOKPOSECR')='X' ;
OkBq:=GetControlText('CBOKPOSBQ')='X' ;
OkParam:=GetControlText(St)='X' ;
SetControlEnabled(St1+'1',OkParam) ;
SetControlEnabled('F'+St1+'2',OkParam) ;
SetControlEnabled(St1+'2',OkParam) ;
SetControlVisible('BVISULIB',OkEcr Or OkBq) ;
SetControlVisible('BRENVOIECR',OkEcr) ;
SetControlVisible('BRENVOIBQ',OkBq) ;
END ;

procedure TOF_PointeReleve.ActiveStop(Ok : Boolean) ;
BEGIN
VH^.STOPRSP:=FALSE ;
SetControlProperty('BSTOP','VISIBLE',Ok) ;
Application.ProcessMessages ;
END ;

procedure TOF_PointeReleve.bStopClick(Sender : Tobject) ;
BEGIN
DemandeStop('') ;
END ;

procedure TOF_PointeReleve.AfficheCR(St,St1 : String) ;
BEGIN
If St='KILL' Then
  BEGIN
  SetControlText('_TR1','') ;
  SetControlText('_TRAITEMENT','') ;
  END Else
  BEGIN
  If St<>'' Then SetControlText('_TR1',St) ;
  If St1<>'' Then SetControlText('_TRAITEMENT',St1) ;
  END ;
END ;

procedure TOF_PointeReleve.Reinit(St : String) ;
Var HG : THGrid ;
    i : Integer ;
BEGIN
HG:=THGrid(GetControl(St)) ; If HG=NIL Then Exit ;
For i:=0 To HG.ColCount-1 Do HG.Cells[i,0]:='' ;
END ;

procedure TOF_PointeReleve.Affecte(Nom,St : String) ;
Var HG : THGrid ;
    i,ll : Integer ;
BEGIN
HG:=THGrid(GetControl(Nom)) ; If HG=NIL Then Exit ;
If St='' Then Exit ;
ll:=Min(HG.ColCount,Length(St)) ;
For i:=0 To ll-1 Do HG.Cells[i,0]:=St[i+1] ;
END ;

procedure TOF_PointeReleve.AffecteZone(NomHG : String ; Ou : Integer = -1) ;
var LigneTob : Integer ;
    TobG,TobE : Tob ;
    St : String ;
    NomChampRef : String ;
    LaRow : Integer ;
    GHG : thGrid ;
BEGIN
ReInit(NomHG) ;
If NomHg='GLIBECR' Then
  BEGIN
  NomChampRef:=GetControlText('COMBOREFECR') ; GHG:=GC ; TobG:=TobEcriture ;
  END Else If NomHg='GLIBBQ' Then
  BEGIN
  NomChampRef:=GetControlText('COMBOREFBQ') ; GHG:=GB ; TobG:=TobReleve ;
  END Else Exit ;
LaRow:=GHG.Row ; If Ou>-1 Then LaRow:=Ou ;
If GHG.RowCount<=1 Then Exit ;
If GHG.cells[SP_POSTOB,LaRow]='' Then Exit ;
LigneTob := StrToInt(GHG.cells[SP_POSTOB,LaRow]);
if (ligneTob >= 0) And (LigneTob<=TobG.Detail.Count-1) then
  BEGIN
  TobE := TobG.Detail[LigneTob];
  If TobE<>NIL Then
    BEGIN
    St:=TobE.GetValue(NomChampRef) ;
    Affecte(NomHG,St)
    END ;
  END ;
END ;

procedure TOF_PointeReleve.BVISULIBClick(Sender: TObject); //Gv 10/06/02
var BT : TToolbarButton97 ;
(* GP : a Affiner :
    Delta,i,j,ll,llInit,llIdeal : Integer ;
    G  : THGrid  ;
    PP : thPanel ;
*)
begin
BT:=TToolbarButton97(GetControl('BVISULIB')) ; If BT=NIL Then Exit ;
SetControlVisible('TBVISULIB',BT.Down) ;
If Not Bt.Down Then Exit ;
AffecteZone('GLIBECR') ; AffecteZone('GLIBBQ') ;

(* GP : a Affiner :
G:=THGrid(getControl('GBANQUE')) ; If G=NIL Then Exit ;
llInit:=G.Height ;
ll:=G.VisibleRowCount+1 ;
i :=(G.Height Div ll) ;
j:=(G.Height Mod ll) ;
//If J / G.DefaultRowHeight>.1 Then Exit ;
//G.DefaultRowHeight:=i ;
If (j<=1)  Then Exit ;
llIdeal:=i*(G.VisibleRowCount+1) ;
PP:=thPanel(GetControl('PINFOBANQ')) ;
Delta:=llInit-llIdeal ;
If PP<>NIL Then
  BEGIN
  If (Delta>3) And (Delta<G.DefaultRowHeight-3) Then PP.height:=PP.height+Delta ;
  END ;
*)
end;

Function TOF_PointeReleve.OkTTB : Boolean ;
Var BT : TToolbarButton97 ;
BEGIN
Result:=FALSE ;
BT:=TToolbarButton97(GetControl('BVISULIB')) ; If BT=NIL Then Exit ;
If Not Bt.Down Then Exit ;
If Not GetControlVisible('TBVISULIB') Then Exit ;
Result:=TRUE ;
END ;

procedure TOF_PointeReleve.GCRowEnter(Sender: TObject; Ou: Longint; var Cancel: Boolean; Chg: Boolean);
begin
If Not OkTTB Then Exit ;
AffecteZone('GLIBECR',ou) ;
End ;

procedure TOF_PointeReleve.GBRowEnter(Sender: TObject; Ou: Longint; var Cancel: Boolean; Chg: Boolean);
begin
If Not OkTTB Then Exit ;
AffecteZone('GLIBBQ',ou) ;
End ;

procedure TOF_PointeReleve.BRENVOIECRClick(Sender: TObject); //Gv 10/06/02
Var SPPosECR1,SPPosECR2 : TSpinEdit ;
    HG : THGrid ;
BEGIN
HG:=THGrid(GetControl('GLIBECR')) ; If HG=NIL Then Exit ;
SPPosECR1:=TSpinEdit(GetControl('POSECR1')) ; SPPosECR2:=TSpinEdit(GetControl('POSECR2')) ;
SPPosECR1.Value:=HG.Selection.Left+1 ;
SPPosECR2.Value:=HG.Selection.Right+1 ;
END ;

procedure TOF_PointeReleve.BRENVOIBQClick(Sender: TObject); //Gv 10/06/02
Var SPPosECR1,SPPosECR2 : TSpinEdit ;
    HG : THGrid ;
BEGIN
HG:=THGrid(GetControl('GLIBBQ')) ; If HG=NIL Then Exit ;
SPPosECR1:=TSpinEdit(GetControl('POSBQ1')) ; SPPosECR2:=TSpinEdit(GetControl('POSBQ2')) ;
SPPosECR1.Value:=HG.Selection.Left+1 ;
SPPosECR2.Value:=HG.Selection.Right+1 ;
END ;
initialization
    registerclasses([TOF_PointeReleve]);
end.

