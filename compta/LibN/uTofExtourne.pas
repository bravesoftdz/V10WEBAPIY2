{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 19/02/2002
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : EXTOURNE ()
Mots clefs ... : TOF;EXTOURNE
*****************************************************************}

unit UTOFEXTOURNE;

//================================================================================
// Interface
//================================================================================
interface
                                       
uses
    StdCtrls,
    Windows, // Pour VK_SPACE
    Controls,
    Classes,
{$IFDEF EAGLCLIENT}
    MaineAGL,
{$ELSE}
    db,
    {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
    fe_main,
{$ENDIF}
    Saisie,
    SaisBor,
    forms,
    sysutils,
    Graphics,
    ComCtrls,
    Menus,
    HCtrls,
    HEnt1,
    HMsgBox,
    UTOF,
    HTB97,
    HDB,
    HQry,
    uTOB,
    HStatus,
    HSysMenu,
    ed_tools,
    Vierge,
    UObjFiltres,  //SG6 12/01/05 gestion Filtres V6 FQ 15255
    Ent1,
    HPanel,
    SaisUtil,
    UtilSais,
    DelVisuE,
    AGLInit,
    UtilSoc,
    DateUtils,
    paramsoc,
    uLibEcriture,
   // Utob,         // InsertOrUpdate
    uObjEtats,    //  TObjEtats.GenereEtatGrille
    extctrls      //FP 10/11/2005 FQ16598 pour TImage
    ;

//==================================================
// Definition des Constante
//==================================================
const
    NOM_FILTRE = 'EXTOURNE';
    // L'accès au paramétrage est interdit. Ces colonnes sont donc fixes.
    COL_NATURE =      1;
    COL_DATE =        2;
    COL_PIECE =       3;
    COL_GENERAL =     4;
    COL_AUXILIAIRE =  5;
    COL_REFINTERNE =  6;
    COL_LIBELLE =     7;
    COL_DEBIT =       8;
    COL_CREDIT =      9;
    COL_PERIODE =     10;
    COL_LIGNE =       11;
    COL_SELECTION=    12;
    COL_NOMBRE =      13;

procedure CPLanceFiche_Extourne;

//==================================================
// Definition de class
//==================================================
type
    TOF_EXTOURNE = class(TOF)
        procedure OnNew                  ; override ;
        procedure OnDelete               ; override ;
        procedure OnUpdate               ; override ;
        procedure OnLoad                 ; override ;
    procedure OnArgument(S: string); override;
    procedure OnClose; override;
    procedure OnClickBImprimer(Sender: Tobject);
    procedure AuxiElipsisClick(Sender : TObject);
  private
    FCritModified: Boolean;                           {FP 10/11/2005 FQ16598}
    ObjFiltre : TObjFiltre; //SG6 12/01/05 Gestion Filtres V6 FQ 15255
    FHSystemMenu    : THSystemMenu;
    FGrille         : THGrid;
    FFiltres        : THValComboBox;
    FPages          : TPageControl;
    FListeEcr       : TOB;
    FListePiece     : TOB;
    FDateGene       : TDateTime;
    FRefOrigine     : Boolean;
    FTypeGene       : string;
    FNegatifGene    : string;
    FLibelleGene    : string;
    FRefInterneGene : string;
    FRefExterneGene : string;
    FJournalGene    : string;
    FJournal        : string;
    FModeSaisie     : string;
    FRapport        : TList;
    FTotalLigne     : integer;
    FTotalDebit     : double;
    FTotalCredit    : double;
    FSelectionLigne : integer;
    FSelectionDebit : double;
    FSelectionCredit: double;

    {b FP 10/11/2005 FQ16598}
    iCritGlyph      : TImage;
    iCritGlyphModified: TImage;
    {e FP 10/11/2005 FQ16598}
    procedure SetCritModified(const Value: Boolean);  {FP 10/11/2005 FQ16598}
    procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure OnEcranResize ( Sender : TObject);
    function  RechercheNatureJal : String;
    procedure OnLanceExtourneClick(Sender: TOBject);
    procedure OnZoomEcritureClick(Sender: TOBject);
    procedure OnChercheClick(Sender: TOBject);
    procedure OnListePiecesClick(Sender: TOBject);
    procedure OnToutSelectionnerClick(Sender: TOBject);
    procedure OnSelectionFinClick(Sender: TOBject);
    procedure FListeDblClick(Sender: TObject);
    procedure OnE_EXERCICEChange(Sender: TObject);
    procedure OnSELECTION_PIECEClick(Sender: TObject);
    procedure OnE_JOURNALChange(Sender: TObject);
    procedure OnGridFlipSelection(Sender: TObject);
    procedure OnChangeLargeurColonneGrille ( Sender : TObject);
    procedure OnPieceSuivanteClick ( Sender : TObject );
    procedure OnPiecePrecedenteClick ( Sender : TObject );
    procedure OnDebutPieceClick( Sender : TObject );
    procedure ChargeLesEcritures;
    procedure GenereLesExtournes;
    procedure ExtourneLigneEcriture(T: TOB; TInfoPiece: TOB);
    procedure FlipSelectionExtourne(i: integer);
    procedure ChargeAnalytique(T: TOB);
    procedure FaireRapport;
    function JePeuxValider: boolean;
    function TrouvePieceDejaCalculee(T: TOB): TOB;
    function MemePiece(iref, i: integer; ForceBordereau : boolean = False): boolean;
    procedure InitMultiCriteres;
    procedure RazSelection;
    procedure MajZoneCumul;
    procedure MajCumulSelection;
    procedure CalculNumGroupeEcrEquivalent;
    procedure PieceSuivante ;
    procedure PiecePrecedente;
    procedure ToutSelectionner(DepuisLigne: integer);
    procedure DebutPiece;
    function  ConstruitRequete : String ;
    function  GetWhereSQL : String ;
    {b FP 10/11/2005 FQ16598}
    procedure InitCritereChange(Parent: TControl);
    procedure CritereChange(Sender: TObject);
    {e FP 10/11/2005 FQ16598}
    procedure MajEcrCompl (TobE, TobNewE : Tob );
    property CritModified: Boolean read FCritModified write SetCritModified; {FP 10/11/2005 FQ16598}
  end;

implementation

uses
  {$IFDEF MODENT1}
  CPTypeCons,
  CPProcGen,
  ULibExercice,
  {$ENDIF MODENT1}
  uLibAnalytique , // Pour AlloueAxe
  UTofMulParamGen; {26/04/07 YMO F5 sur Auxiliaire }

{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 04/01/2005
Modifié le ... :   /  /
Description .. : - FB 15169 - LG 04/01/2004 - Fonction accessible même
Suite ........ : avec le concept applicatif interdissant la saisie
Mots clefs ... :
*****************************************************************}
procedure CPLanceFiche_Extourne;
begin

 if not ExJaiLeDroitConcept(TConcept(ccSaisEcritures), true) then
  exit ;

 AGLLanceFiche('CP','EXTOURNE','','','');

end;

procedure TOF_EXTOURNE.OnNew;
begin
  inherited;
end;

procedure TOF_EXTOURNE.OnDelete;
begin
  inherited;
end;

procedure TOF_EXTOURNE.OnUpdate;
begin
  inherited;
end;

procedure TOF_EXTOURNE.OnLoad;
begin
  inherited;
  ObjFiltre.Charger;
end;

procedure TOF_EXTOURNE.OnArgument(S: string);
var
  PopF, PopF11: TPopupMenu;
  Composants : TControlFiltre;
  BFILTRE:TToolBarButton97;
begin
  inherited;

  {b FP 10/11/2005 FQ16598}
  iCritGlyph := nil;
  iCritGlyphModified := nil;
  {e FP 10/11/2005 FQ16598}

  FTotalLigne := 0;
  FTotalDebit := 0;
  FTotalCredit := 0;
  FSelectionLigne := 0;
  FSelectionDebit := 0;
  FSelectionCredit := 0;
  Ecran.HelpContext := 7677000;
  Ecran.OnKeyDown := FormKeyDown;
  Ecran.OnResize := OnEcranResize;
  FHSystemMenu := THSystemMenu(TFVierge(ECRAN).HMTrad);

  TCheckBox ( GetControl('SELECTION_PIECE') ).OnClick := OnSELECTION_PIECEClick;
  THValComboBox ( GetControl ('E_JOURNAL') ).OnChange := OnE_JOURNALChange;

  FFiltres := THValComboBox(GetControl('FFILTRES'));

  FPages := TPageControl(GetControl('PAGES'));
  PopF := TPopupMenu(GetControl('POPF'));

  //SG6 12/01/05 Gestion Filtres V56 FQ 15255
  BFILTRE:=TToolBarButton97(GetControl('BFILTRE')); if BFILTRE=nil then Exit;

  Composants.PopupF   := POPF;
  Composants.Filtres  := FFILTRES;
  Composants.Filtre   := BFILTRE;
  Composants.PageCtrl := FPages;
  ObjFiltre := TObjFiltre.Create(Composants,NOM_FILTRE);

  PopF11 := TPopupMenu(GetControl('POPF11'));
  PopF11.Items[0].OnClick := OnZoomEcritureClick;
  PopF11.Items[1].OnClick := OnToutSelectionnerClick;
  PopF11.Items[2].OnClick := OnPiecePrecedenteClick;
  PopF11.Items[3].OnClick := OnPieceSuivanteClick;
  PopF11.Items[4].OnClick := OnDebutPieceClick;
  PopF11.Items[5].OnClick := OnLanceExtourneClick;

  FGrille := THGrid(GetControl('FLISTE'));
  FGrille.SynEnabled := False;
  FGrille.OnFlipSelection := OnGridFlipSelection;
  FGrille.OnColumnWidthsChanged := OnChangeLargeurColonneGrille;
  FGrille.ColCount := COL_NOMBRE;
  FGrille.ColWidths[COL_PERIODE] := -1;
  FGrille.ColLengths[COL_PERIODE] := -1;
  FGrille.ColWidths[COL_LIGNE] := -1;
  FGrille.ColLengths[COL_LIGNE] := -1;
  FGrille.ColWidths[COL_SELECTION] := -1;
  FGrille.ColLengths[COL_SELECTION] := -1;

//  FGrille.ColWidths[0] := 20; {Colonne du dbIndicator}
  FGrille.ColWidths[COL_NATURE] := 50;
  FGrille.ColWidths[COL_DATE] := 80;
  FGrille.ColWidths[COL_PIECE] := 40;
  FGrille.ColWidths[COL_GENERAL] := 100;
  FGrille.ColWidths[COL_AUXILIAIRE] := 100;
  FGrille.ColWidths[COL_REFINTERNE] := 100;
  FGrille.ColWidths[COL_LIBELLE] := 160;
  FGrille.ColWidths[COL_DEBIT] := 100;
  FGrille.ColWidths[COL_CREDIT] := 100;
  FGrille.ColAligns[COL_NATURE] := taCenter;
  FGrille.ColAligns[COL_PIECE] := taCenter;
  FGrille.ColAligns[COL_DATE] := taCenter;
  FGrille.ColAligns[COL_LIBELLE] := taLeftJustify;
  FGrille.ColAligns[COL_GENERAL] := taLeftJustify;
  FGrille.ColAligns[COL_AUXILIAIRE] := taLeftJustify;
  FGrille.ColAligns[COL_DEBIT] := taRightJustify;
  FGrille.ColAligns[COL_CREDIT] := taRightJustify;

  FGrille.ColFormats[COL_PERIODE] := '#.##';
  FGrille.ColFormats[COL_PIECE] := '#.##';
  FGrille.ColFormats[COL_DEBIT] := '#,##0.00;-#,##0.00; ;';
  FGrille.ColFormats[COL_CREDIT] := '#,##0.00;-#,##0.00; ;';

  FGrille.Cells[COL_NATURE, 0] := TraduireMemoire('Nature');
  FGrille.Cells[COL_PIECE, 0] := TraduireMemoire('N°');
  FGrille.Cells[COL_DATE, 0] := TraduireMemoire('Date');
  FGrille.Cells[COL_GENERAL, 0] := TraduireMemoire('Général');
  FGrille.Cells[COL_AUXILIAIRE, 0] := TraduireMemoire('Auxiliaire');
  FGrille.Cells[COL_LIBELLE, 0] := TraduireMemoire('Libellé');
  FGrille.Cells[COL_REFINTERNE, 0] := TraduireMemoire('Référence');
  FGrille.Cells[COL_DEBIT, 0] := TraduireMemoire('Débit');
  FGrille.Cells[COL_CREDIT, 0] := TraduireMemoire('Crédit');

  FListeEcr := TOB.Create('', nil, -1);
  FListePiece := TOB.Create('', nil, -1);
  FRapport := TList.Create;
  TToolBarButton97(GetControl('BVALIDER')).OnClick := OnLanceExtourneClick;
  TToolBarButton97(GetControl('BZOOM')).OnClick := OnZoomEcritureClick;
  TToolBarButton97(GetControl('BTOUTSEL')).OnClick := OnToutSelectionnerClick;
  TToolBarButton97(GetControl('BCHERCHE')).OnClick := OnChercheClick;
  TToolBarButton97(GetControl('BLISTEPIECES')).OnClick := OnListePiecesClick;
  TToolBarButton97(GetControl('BSELECTIONFIN')).OnClick := OnSelectionFinClick;
  TToolBarButton97(GetControl('BIMPRIMER')).OnClick := OnClickBImprimer;
  THGrid(GetControl('FLISTE')).OnDblClick := FListeDblClick;
  THValComboBox(GetControl('E_EXERCICE')).OnChange := OnE_EXERCICEChange;

  //YMO 30/03/2006 (Tâche 3747) Sélection multi-établissement
  THValComboBox(GetControl('E_ETABLISSEMENT')).Vide := True;
  THValComboBox(GetControl('E_ETABLISSEMENT')).VideString := Traduirememoire('<<Tous>>');
  THValComboBox(GetControl('E_ETABLISSEMENT')).Reload;


  InitMultiCriteres;
  OnE_JOURNALChange(nil);
  FGrille.SynEnabled := True;

  {b FP 10/11/2005 FQ16598}
  iCritGlyph := TImage(GetControl('ICRITGLYPH', True));
  iCritGlyphModified := TImage(GetControl('ICRITGLYPHMODIFIED', True));
  InitCritereChange(Ecran);    {Surcharge l'évènement OnChange ou OnClick des composants pour savoir si un critère a été modifié}
  CritModified := false;
  {e FP 10/11/2005 FQ16598}

  if GetParamSocSecur('SO_CPMULTIERS', false) then
    THEdit(GetControl('E_AUXILIAIRE', true)).OnElipsisClick:=AuxiElipsisClick;

end;

procedure TOF_EXTOURNE.OnClose;
begin
  FreeAndNil(ObjFiltre); // SG6 15/11/04 Gestion des Filtres FQ 14826
  VideListe(FRapport);
  FRapport.Free;
  FListePiece.Free;
  FListeEcr.Free;
  inherited;
end;

{***********A.G.L.***********************************************
Auteur  ...... : YMO
Créé le ...... : 08/03/2007
Modifié le ... :   /  /
Description .. : FQ15070 Rajout du bouton imprimer
Mots clefs ... : 15070 imprimer
*****************************************************************}
procedure TOF_EXTOURNE.OnClickBImprimer(Sender: Tobject);
begin
  if FListeEcr.Detail.Count > 0 then
     TObjEtats.GenereEtatGrille (FGrille,Ecran.Caption,False);
end;

{***********A.G.L.***********************************************
Auteur  ...... : YMO
Créé le ...... : 26/04/2007
Modifié le ... :   /  /
Description .. : Branchement de la fiche auxiliaire
Mots clefs ... :
*****************************************************************}
procedure TOF_EXTOURNE.AuxiElipsisClick( Sender : TObject );
begin
     THEdit(Sender).text:= CPLanceFiche_MULTiers('M;' +THEdit(Sender).text + ';' +THEdit(Sender).Plus + ';');
end;


procedure TOF_EXTOURNE.OnListePiecesClick(Sender: TOBject);
begin
  if FRapport.Count > 0 then
    VisuPiecesGenere(FRapport, EcrGen, 10);
end;

procedure TOF_EXTOURNE.OnZoomEcritureClick(Sender: TOBject);
begin
  FListeDblClick(nil);
end;

procedure TOF_EXTOURNE.FListeDblClick(Sender: TObject);
var
  P: RParFolio;
  z, zz: integer;
  Periode, Piece: integer;
  DateComptable: TDateTime;
  Q: TQuery;
  St: string;
begin
  if FListeEcr.Detail.Count = 0 then exit;
  Periode := StrToInt(FGrille.Cells[COL_PERIODE, FGrille.Row]);
  Piece := StrToInt(FGrille.Cells[COL_PIECE, FGrille.Row]);
  DateComptable := StrToDate(FGrille.Cells[COL_DATE, FGrille.Row]);
  if FModeSaisie = '-' then // mode pièce
  begin
    St := 'SELECT E_DATECOMPTABLE, E_JOURNAL, E_NUMEROPIECE, E_NUMLIGNE FROM ECRITURE '
      + ' WHERE E_DATECOMPTABLE="' + USDateTime(DateComptable) + '"'
      + ' AND E_JOURNAL="' + FJournal + '" AND E_NUMEROPIECE=' + IntToStr(Piece)
      + ' AND E_NUMLIGNE=1 AND E_NUMECHE<=1'
      + ' AND E_QUALIFPIECE="' + GetControlText('E_QUALIFPIECE') + '"'
      + ' ORDER BY E_JOURNAL, E_EXERCICE, E_DATECOMPTABLE, E_NUMEROPIECE, E_NUMLIGNE, '
      + 'E_NUMECHE, E_QUALIFPIECE';
    ;
    Q := OpenSQL(St, True);
    TrouveEtLanceSaisie(Q, taConsult, GetControlText('E_QUALIFPIECE'));
    Ferme(Q);
  end
  else // mode bordereau ou libre
  begin
    FillChar(P, Sizeof(P), #0);
    z := Trunc(Periode / 100);
    zz := Periode - (z * 100);
    P.ParPeriode := DateToStr(DebutDeMois(EncodeDate(z, zz, 1)));
    P.ParCodeJal := FJournal;
    P.ParNumFolio := IntToStr(Piece);
    P.ParNumLigne := StrToInt(FGrille.Cells[COL_LIGNE, FGrille.Row]);
    ChargeSaisieFolio(P, taConsult);
  end;
end;

procedure TOF_EXTOURNE.OnE_EXERCICEChange(Sender: TObject);
begin
  inherited;
  ExoToDates(GetControlText('E_EXERCICE'), GetControl('E_DATECOMPTABLE'),
    GetControl('E_DATECOMPTABLE_'));
end;

function TOF_EXTOURNE.JePeuxValider: boolean;
var
  i: integer;
  D, C: double;
begin
  Result := False;
  // Si aucun écriture sélectionnée ???? .... et bien on sort !
  if (Arrondi(FSelectionDebit,V_PGI.OkDecV)=0) and (Arrondi(FSelectionCredit,V_PGI.OkDecV)=0) then
  begin
    MessageAlerte('Aucune écriture sélectionnée.');
    exit;
  end;
  // Vérification systématique de l'équilibre de la pièce.
//  if GetCheckBoxState('SELECTION_PIECE') = cbUnChecked then
//  begin
  D := 0;
  C := 0;
  for i := 0 to FListeEcr.Detail.Count - 1 do
  begin
    if FListeEcr.Detail[i].GetValue('SELECTION') = '+' then
    begin
      D := D + FListeEcr.Detail[i].GetValue('E_DEBIT');
      C := C + FListeEcr.Detail[i].GetValue('E_CREDIT');
    end;
  end;
  if (Arrondi(D - C, V_PGI.OkDecV) <> 0) then
  begin
    MessageAlerte('Extourne impossible.#10#13Le lot d''écritures sélectionnées n''est pas équilibré.');
    Exit;
  end;
//  end;
  Result := True;
end;

procedure TOF_EXTOURNE.GenereLesExtournes;
var
  i, iRef, NewPiece: integer;
  TEcr, TInsert, TNewEcr, TInfoPiece: TOB;
  bSelectionPiece, bPieceUnique: boolean;
begin
  TInfoPiece := nil;
  FListePiece.ClearDetail;
  bSelectionPiece := (GetCheckBoxState('SELECTION_PIECE') = cbChecked);
  bPieceUnique := False;

  // Suppression des enregistrements non sélectionnés
  i := 0;
  TInsert := TOB.Create('', nil, -1);
  while (i < FListeEcr.Detail.Count) do
  begin
    if FListeEcr.Detail[i].GetValue('SELECTION') <> '+' then
      FListeEcr.Detail[i].Free
    else
      Inc(i, 1);
  end;

  // Instanciation de la totalité des renseignements des écritures
  for i := 0 to FListeEcr.Detail.count - 1 do
    begin
    FListeEcr.Detail[i].loadDB ;
    end ;

  // Calcul du numéro de pièce pour le cas de la sélection libre

  // Traitements des enregistrements sélectionnés
  iRef := -1;
  for i := 0 to FListeEcr.Detail.Count - 1 do
  begin
    TEcr := FListeEcr.Detail[i];
    TEcr.PutValue('E_QUALIFORIGINE', 'URN');
  
    // Détection de changement de pièce
    if (not (bPieceUnique)) and (not (MemePiece(iRef, i))) then
    begin
      iRef := i;
      TInfoPiece := TrouvePieceDejaCalculee(TEcr);
      // Calcul du nouveau numéro de pièce
      if TInfoPiece = nil then
      begin
        // Si nouvelle pièce, on enregistre la dernière pièce pour ne pas
        // perturber le calcul du numéro de pièce
        if ((i <> 0) and (TInsert <> nil)) then
        begin
          TInsert.SetAllModifie(True);
          TInsert.InsertDB(nil, True);
          MajSoldesEcritureTOB(TInsert, True);
          TInsert.ClearDetail;
        end;
        // Calcul des infos de la nouvelle pièce
        // FQ 16868 SBO 14/10/2005 : Calcul du nouveau numéro de pièce suivant type de pièce paramétré, et non plus filtre de recherche
        {YMO 27/03/2007 Journal de génération et non journal d'origine} {FQ19894 YMO 30/03/2007 Compteur du journal de génération et non du journal d'origine}
        NewPiece := GetNewNumJal( FJournalGene, (FTypeGene = 'N') or (FTypeGene = 'I'), FDateGene, '', FModeSaisie);
        TInfoPiece := TOB.Create('', FListePiece, -1);
        TInfoPiece.AddChampSupValeur('OLDPIECE',
          TEcr.GetValue('E_NUMEROPIECE'));
        TInfoPiece.AddChampSupValeur('NEWPIECE', NewPiece);
        TInfoPiece.AddChampSupValeur('OLDPERIODE', TEcr.GetValue('E_PERIODE'));
        TInfoPiece.AddChampSupValeur('OLDDATE',
          TEcr.GetValue('E_DATECOMPTABLE'));
        TInfoPiece.AddChampSupValeur('NUMLIGNE', 1);
        TInfoPiece.AddChampSupValeur('NUMGROUPEECR', 1);
      end
      else if FModeSaisie <> '-' then
        TInfoPiece.PutValue('NUMGROUPEECR', TInfoPiece.GetValue('NUMGROUPEECR') + 1);
      {YMO FQ18415 Une seule pièce que l'on coche ou pas
      if not bSelectionPiece then
          bPieceUnique := True;}

    end;
    TNewEcr := TOB.Create('ECRITURE', TInsert, -1);
    // Création de la nouvelle écritureisteEcr,-1);
    TNewEcr.Dupliquer(TEcr, False, True);
    // Chargement de l'analytique sous l'écriture
    if (TNewEcr.GetValue('E_ANA') = 'X') then
      ChargeAnalytique(TNewEcr);
    ExtourneLigneEcriture(TNewEcr, TInfoPiece);
    {YMO 29/03/07 utilisation champ ECRCOMPL pour la requête edit GL analytique}
    {L'écriture d'extourne à la date de l'écriture d'origine et inversement}
    MajEcrCompl(TEcr, TNewEcr);
    MajEcrCompl(TNewEcr, TEcr);

  end;
  if (TInsert <> nil) then
  begin
    TInsert.SetAllModifie(True);
    TInsert.InsertDB(nil, True);

    MajSoldesEcritureTOB(TInsert, True);
  end;
  FListeEcr.UpdateDB(True);
// TVA sur encaissement non prise en compte en mode PCL
  {#TVAENC}
  {if VH^.OuiTvaEnc then
     BEGIN
     InitMove(TPiece.Count,'') ;
     for i:=0 to TPiece.Count-1 do
         BEGIN
         MoveCur(False) ;
         O:=TOBM(TPiece[i]) ; RR:=OBMToIdent(O,False) ;
         ElementsTvaEnc(RR,True) ;
         END ;
     FiniMove ;
     END ;
   }
    // Constitution du rapport
end;

{***********A.G.L.***********************************************
Auteur  ...... : YMO
Créé le ...... : 09/03/2007
Modifié le ... :   /  /
Description .. : Choix du Jal de génération : on va n'offrir le choix que dans
Suite ........ : la même nature de journal que celui d'origine
Mots clefs ... : 13822 NATUREJAL
*****************************************************************}
function TOF_EXTOURNE.RechercheNatureJal : String;
var Q : TQuery;
begin

  Q:=OpenSQL('SELECT J_NATUREJAL FROM JOURNAL WHERE J_JOURNAL="'+FJournal+'"',True) ;
  if not Q.EOF then
  begin
   Result:=Q.Fields[0].AsString;
   Ferme(Q) ;
  end;

end;

{***********A.G.L.***********************************************
Auteur  ...... : YMO
Créé le ...... : 02/04/2007
Modifié le ... : 05/12/2007
Description .. : Utilisation champ ECRCOMPL pour la requête edit GL
Suite ........ : analytique ; l'écriture d'origine a la date de l'écriture
Suite ........ : d'extourne et inversement
Mots clefs ... : ECRCOMPL EC_DATEEXT
*****************************************************************}
procedure TOF_EXTOURNE.MajEcrCompl (TobE, TobNewE : Tob );
var
  lTobComplE, lTobComplNewE : Tob ;
begin

  lTobComplE    := CSelectDBTOBCompl( TobE, nil ) ;
  lTobComplNewE := CSelectDBTOBCompl( TobNewE, nil ) ;

  CMAJTobCompl( TobNewE ) ; // clé ecriture

  CPutValueTOBCompl( TOBNewE, 'EC_DATEEXT', TobE.GetDateTime('E_DATECOMPTABLE') ) ;
  {FQ21881  05/12/2007  YMO Recopie de la période (le 2e passage en sens inverse ne gêne pas)}
  CPutValueTOBCompl( TOBNewE, 'EC_CUTOFFDEB', lTobComplE.GetDateTime('EC_CUTOFFDEB') ) ;
  CPutValueTOBCompl( TOBNewE, 'EC_CUTOFFFIN', lTobComplE.GetDateTime('EC_CUTOFFFIN') ) ;

  lTobComplNewE.InsertorUpdateDb ;

end;

procedure TOF_EXTOURNE.OnLanceExtourneClick(Sender: TOBject);
var TInfo : TOB;
begin
  if not JePeuxValider then
    exit;
  TInfo := TOB.Create('', nil, -1);
  TInfo.AddChampSupValeur('JOURNAL',FJournal);
  TInfo.AddChampSupValeur('NATUREJAL',RechercheNatureJal);
  TInfo.AddChampSupValeur('MODESAISIE',FModeSaisie);
  TInfo.AddChampSupValeur('LIGNES',FSelectionLigne);
  TInfo.AddChampSupValeur('QUALIFPIECE',GetControlText('E_QUALIFPIECE'));
  TInfo.AddChampSupValeur('MONTANT',Abs(FSelectionDebit));
  TheTOB := TInfo;
  AGLLanceFiche('CP','EXTOURNE_PARAM','','','');
  TInfo.Free;
  if TheTOB = nil then exit;
  FRefOrigine := (TheTOB.GetString('CONTRE_REFORIGINE')='X');
  FDateGene := TheTOB.GetDateTime('CONTRE_DATE');
  FTypeGene := TheTOB.GetString('CONTRE_TYPE');
  FNegatifGene := TheTOB.GetString('CONTRE_NEGATIF');
  FLibelleGene := TheTOB.GetString('CONTRE_LIBELLE');
  FRefInterneGene := TheTOB.GetString('CONTRE_REFINTERNE');
  FRefExterneGene := TheTOB.GetString('CONTRE_REFEXTERNE');
  FJournalGene := TheTOB.GetString('CONTRE_JOURNAL'); {YMO 09/03/07 FQ13822 Choix du Jal de génération}
  if (PGIAsk('Confirmez-vous la génération de l''écriture d''extourne ?',
    ECRAN.Caption) <> mrYes) then
    exit;
  if Transactions(GenereLesExtournes, 3) <> oeOK then
    MessageAlerte('ATTENTION : Problème lors de l''extourne!')
  else
  begin
    MarquerPublifi(True) ;
    CPStatutDossier ;
    FaireRapport;
    RazSelection;
    if FRapport.Count > 0 then
      if (PGIAsk('Voulez-vous visualiser les écritures générées ?',Ecran.Caption)=mrYes) then
        VisuPiecesGenere(FRapport, EcrGen, 10);
    ChargeLesEcritures;        
  end;
end;

procedure TOF_EXTOURNE.ChargeLesEcritures;
var
  St: string;
  i: integer;
begin
    FGrille.SynEnabled := False;
    FListeEcr.ClearDetail;
    FGrille.VidePile(False);

    // SBO 14/10/2005 : Modif de la requête pour optimisation
    St := ConstruitRequete ;

    // chargement et affichage des ecritures
    FListeEcr.LoadDetailDBFromSQl('ECRITURE',St,false);

    FListeEcr.PutGridDetail(THGrid(GetControl('FLISTE')),false,false,'E_NATUREPIECE;E_DATECOMPTABLE;E_NUMEROPIECE;E_GENERAL;E_AUXILIAIRE;E_REFINTERNE;E_LIBELLE;E_DEBIT;E_CREDIT;E_PERIODE;E_NUMLIGNE');

    // ajout d'une zone de selection
    if FListeEcr.Detail.Count > 0 then
      FListeEcr.Detail[0].AddChampSupValeur('SELECTION', '-', True);

    // selection
    FGrille.OnFlipSelection := nil;
    for i := 1 to FGrille.RowCount - 1 do
      if FGrille.IsSelected(i) then FGrille.FlipSelection(i);
    FGrille.OnFlipSelection := OnGridFlipSelection;

    if FModeSaisie='LIB' then
      CalculNumGroupeEcrEquivalent;

    // Redimensionnement des colonnes
    TFVierge(Ecran).FormResize := True;
    FHSystemMenu.ResizeGridColumns(FGrille);

    // Recalcul des totaux
    FTotalLigne := FListeEcr.Detail.Count;
    FTotalDebit := FListeEcr.Somme('E_DEBIT',[''],[''],False);
    FTotalCredit := FListeEcr.Somme('E_CREDIT',[''],[''],False);

    // affichage des cummuls
    MajCumulSelection;
    MajZoneCumul;
    FGrille.SynEnabled := True;

end;

procedure TOF_EXTOURNE.OnChercheClick(Sender: TOBject);
begin
  CritModified := False;         {FP 10/11/2005 FQ16598}
  ChargeLesEcritures;
end;

procedure TOF_EXTOURNE.OnGridFlipSelection(Sender: TObject);
var
  i, iRef: integer;
begin
  inherited;
  // Mémorisation de la ligne de référence
  iRef := FGrille.Row;
  // Mise à jour de la ligne de référence dans la TOB
  if FListeEcr.Detail[iRef - 1].GetValue('SELECTION') = '+' then
  begin
    FListeEcr.Detail[iRef - 1].PutValue('SELECTION', '-');
    FGrille.Cells[COL_SELECTION,iRef]:='-';
  end
  else
  begin
    FListeEcr.Detail[iRef - 1].PutValue('SELECTION', '+');
    FGrille.Cells[COL_SELECTION,iRef]:='+';
  end;
  // Détection des écritures de la pièce
  // On ne fait pas ce traitement si suite click bouton tout sélectionner
  if (GetCheckBoxState('SELECTION_PIECE') = cbchecked) then
  begin
    // Détection des écritures de la pièce en aval
    for i := iRef + 1 to FGrille.RowCount - 1 do
    begin
      if MemePiece(iRef - 1, i - 1) then
        FlipSelectionExtourne(i)
      else
        break;
    end;
    // Détection des écritures de la pièce en amont
    for i := iRef - 1 downto 1 do
    begin
      if MemePiece(iRef - 1, i - 1) then
        FlipSelectionExtourne(i)
      else
        break;
    end;
  end;
  // Rafraichissement de la grille
  FGrille.Invalidate;
  MajCumulSelection;
  // Positionnement sur la pièce suivante
  if GetCheckBoxState('SELECTION_PIECE') = cbChecked then
    PieceSuivante else
  if FGrille.Row < (FGrille.RowCount - 1) then FGrille.Row := FGrille.Row+1;
end;

procedure TOF_EXTOURNE.ExtourneLigneEcriture(T: TOB; TInfoPiece: TOB);
var
  EL: string;
  StRef: string;
  X: double;
  i, j: integer;
  TAxe, TAna: TOB;
begin
  if T = nil then
    exit;

  if FLibelleGene <> '' then T.PutValue('E_LIBELLE', FLibelleGene);

  {YMO 09/03/07 FQ12700 Choix de la référence interne}
  if not FRefOrigine then
  begin
    if FRefInterneGene = '' then
    begin
        StRef := Copy(TraduireMemoire('Extourne') + ' ' +
        Inttostr(T.GetValue('E_NUMEROPIECE')) + '  ' +
        DateToStr(T.GetValue('E_DATECOMPTABLE')), 1, 35);

        T.PutValue('E_REFINTERNE', StRef);
    end
    else
        T.PutValue('E_REFINTERNE', FRefInterneGene);
  end;


  if FRefExterneGene <> '' then T.PutValue('E_REFEXTERNE', FRefExterneGene);

  T.PutValue('E_JOURNAL', FJournalGene); {YMO 09/03/07 FQ13822 Choix du Jal de génération}
  T.PutValue('E_DATECOMPTABLE', FDateGene);
  T.PutValue('E_EXERCICE',QuelExoDt(FDateGene));
  T.PutValue('E_DATEECHEANCE', FDateGene);
  if FModeSaisie <> '-' then
  begin
    T.PutValue('E_NUMGROUPEECR', TInfoPiece.GetValue('NUMGROUPEECR'));
    T.PutValue('E_NUMLIGNE', TInfoPiece.GetValue('NUMLIGNE'));
    TInfoPiece.PutValue('NUMLIGNE', TInfoPiece.GetValue('NUMLIGNE') + 1);
  end;
{$IFNDEF SPEC302}
  T.PutValue('E_PERIODE', GetPeriode(FDateGene));
  T.PutValue('E_SEMAINE', NumSemaine(FDateGene));
{$ENDIF}
  T.PutValue('E_EXERCICE', QuelExoDT(FDateGene));
  T.PutValue('E_QUALIFPIECE', FTypeGene);
  T.PutValue('E_NUMEROPIECE', TInfoPiece.GetValue('NEWPIECE'));
  T.PutValue('E_VALIDE', '-');
  T.PutValue('E_LETTRAGE', '');
  T.PutValue('E_TRACE', '');
  T.PutValue('E_QUALIFORIGINE', 'EXT');
  EL := T.GetValue('E_ETATLETTRAGE');
  if ((EL = 'TL') or (EL = 'PL') or (EL = 'AL')) then
    T.PutValue('E_ETATLETTRAGE', 'AL')
  else
    T.PutValue('E_ETATLETTRAGE', 'RI');
  T.PutValue('E_LETTRAGEDEV', '-');
  T.PutValue('E_COUVERTURE', 0);
  T.PutValue('E_COUVERTUREDEV', 0);
  T.PutValue('E_DATEPAQUETMIN', FDateGene);
  T.PutValue('E_DATEPAQUETMAX', FDateGene);
  T.PutValue('E_REFPOINTAGE', '');
  T.PutValue('E_DATEPOINTAGE', IDate1900);
  T.PutValue('E_REFRELEVE', '');
  T.PutValue('E_FLAGECR', '');
  T.PutValue('E_ETAT', '0000000000');
  T.PutValue('E_NIVEAURELANCE', 0);
  T.PutValue('E_DATERELANCE', IDate1900);
  T.PutValue('E_DATECREATION', Date);
  T.PutValue('E_DATEMODIF', NowH);
  T.PutValue('E_SUIVDEC', '');
  T.PutValue('E_NOMLOT', '');
  T.PutValue('E_EDITEETATTVA', '-');
  if ( (GetParamSocSecur('SO_CPLIENGAMME','')<>'') and (GetParamSocSecur('SO_CPLIENGAMME','')<>'AUC') and (GetParamSocSecur('SO_CPLIENGAMME','')<>'SI') ) then
  begin
    T.PutValue('E_IO', 'X');
    { FQ 14984 - CA - 01/12/2004 - Ne pas reporter les informations de synchro sur les écritures extournées }
    T.PutValue('E_REFREVISION',0);
    T.PutValue('E_LIBRETEXTE9','');
  end;
  if FNegatifGene = 'X' then
  begin
    T.PutValue('E_DEBIT', -T.GetValue('E_DEBIT'));
    T.PutValue('E_CREDIT', -T.GetValue('E_CREDIT'));
    T.PutValue('E_DEBITDEV', -T.GetValue('E_DEBITDEV'));
    T.PutValue('E_CREDITDEV', -T.GetValue('E_CREDITDEV'));
    T.PutValue('E_QTE1', -T.GetValue('E_QTE1'));
    T.PutValue('E_QTE2', -T.GetValue('E_QTE2'));
  end
  else
  begin
    X := T.GetValue('E_DEBIT');
    T.PutValue('E_DEBIT', T.GetValue('E_CREDIT'));
    T.PutValue('E_CREDIT', X);
    X := T.GetValue('E_DEBITDEV');
    T.PutValue('E_DEBITDEV', T.GetValue('E_CREDITDEV'));
    T.PutValue('E_CREDITDEV', X);
  end;
  // Traitement de l'analytique
  for i := 0 to T.Detail.Count - 1 do
  begin
    TAxe := T.Detail[i];
    for j := 0 to TAxe.Detail.Count - 1 do
    begin
      TAna := TAxe.Detail[j];
      TAna.PutValue('Y_JOURNAL', FJournalGene); {FQ22340 18.02.08 YMO Jal de génération reporté sur l'analytique}
      TAna.PutValue('Y_DATECOMPTABLE', FDateGene);
      TAna.PutValue('Y_NUMLIGNE', T.GetValue('E_NUMLIGNE'));
{$IFNDEF SPEC302}
      TAna.PutValue('Y_PERIODE', GetPeriode(FDateGene));
      TAna.PutValue('Y_SEMAINE', NumSemaine(FDateGene));
{$ENDIF}
      TAna.PutValue('Y_EXERCICE', QuelExoDT(FDateGene));
      TAna.PutValue('Y_QUALIFPIECE', FTypeGene);
      TAna.PutValue('Y_NUMEROPIECE', TInfoPiece.GetValue('NEWPIECE'));
      TAna.PutValue('Y_TRACE', '');
      TAna.PutValue('Y_VALIDE', '-');
      TAna.PutValue('Y_DATECREATION', Date);
      TAna.PutValue('Y_DATEMODIF', NowH);
      if FNegatifGene='X' then
      begin
        TAna.PutValue('Y_DEBIT', -TAna.GetValue('Y_DEBIT'));
        TAna.PutValue('Y_CREDIT', -TAna.GetValue('Y_CREDIT'));
        TAna.PutValue('Y_DEBITDEV', -TAna.GetValue('Y_DEBITDEV'));
        TAna.PutValue('Y_CREDITDEV', -TAna.GetValue('Y_CREDITDEV'));
        TAna.PutValue('Y_QTE1', -TAna.GetValue('Y_QTE1'));
        TAna.PutValue('Y_QTE2', -TAna.GetValue('Y_QTE2'));
        TAna.PutValue('Y_TOTALQTE1', -TAna.GetValue('Y_TOTALQTE1'));
        TAna.PutValue('Y_TOTALQTE2', -TAna.GetValue('Y_TOTALQTE2'));
        TAna.PutValue('Y_TOTALECRITURE', -TAna.GetValue('Y_TOTALECRITURE'));
        TAna.PutValue('Y_TOTALDEVISE', -TAna.GetValue('Y_TOTALDEVISE'));
      end
      else
      begin
        X := TAna.GetValue('Y_DEBIT');
        TAna.PutValue('Y_DEBIT', TAna.GetValue('Y_CREDIT'));
        TAna.PutValue('Y_CREDIT', X);
        X := TAna.GetValue('Y_DEBITDEV');
        TAna.PutValue('Y_DEBITDEV', TAna.GetValue('Y_CREDITDEV'));
        TAna.PutValue('Y_CREDITDEV', X);
      end;
    end;
  end;
end;

function TOF_EXTOURNE.MemePiece(iref, i: integer; ForceBordereau : boolean = False): boolean;
begin
  Result := False;
  if FListeEcr.Detail.Count = 0 then
    exit;
  if i >= FListeEcr.Detail.Count then
    exit;
  if iref = -1 then
    exit;

  if FModeSaisie = '-' then
    Result := FListeEcr.Detail[iref].GetValue('E_NUMEROPIECE') =
      FListeEcr.Detail[i].GetValue('E_NUMEROPIECE')
  else if ((FModeSaisie = 'BOR') or (ForceBordereau)) then
    Result := (FListeEcr.Detail[iref].GetValue('E_NUMGROUPEECR') =
      FListeEcr.Detail[i].GetValue('E_NUMGROUPEECR'))
      and (FListeEcr.Detail[iref].GetValue('E_NUMEROPIECE') =
        FListeEcr.Detail[i].GetValue('E_NUMEROPIECE'))
      and (FListeEcr.Detail[iRef].GetValue('E_PERIODE') =
        FListeEcr.Detail[i].GetValue('E_PERIODE'))
      {FQ18415 YMO 22/03/2007 Folio différent par établissement}
      and (FListeEcr.Detail[iRef].GetValue('E_ETABLISSEMENT') =
        FListeEcr.Detail[i].GetValue('E_ETABLISSEMENT'))
  else if FModeSaisie = 'LIB' then
  begin
    Result := (FListeEcr.Detail[iref].GetValue('E_NUMEROPIECE') =
        FListeEcr.Detail[i].GetValue('E_NUMEROPIECE'))
      and (FListeEcr.Detail[iRef].GetValue('E_PERIODE') =
        FListeEcr.Detail[i].GetValue('E_PERIODE'))
     {FQ18415 YMO 22/03/2007 Folio différent par établissement}
      and (FListeEcr.Detail[iRef].GetValue('E_ETABLISSEMENT') =
        FListeEcr.Detail[i].GetValue('E_ETABLISSEMENT'));
    {FQ19187 YMO 12/03/07 Sélection en mode libre par période on prend ttes les lignes}
    {Result := FListeEcr.Detail[iref].GetValue('NUMGROUPEECRLIB') =
      FListeEcr.Detail[i].GetValue('NUMGROUPEECRLIB');}
  end;
end;

function TOF_EXTOURNE.TrouvePieceDejaCalculee(T: TOB): TOB;
var
  i: integer;
begin
  Result := nil;
  for i := 0 to FListePiece.Detail.Count - 1 do
  begin
    if (FListePiece.Detail[i].GetValue('OLDPIECE') = T.GetValue('E_NUMEROPIECE'))
      and (FListePiece.Detail[i].GetValue('OLDPERIODE') =
        T.GetValue('E_PERIODE')) then
    begin
      Result := FListePiece.Detail[i];
      break;
    end;
  end;
end;

procedure TOF_EXTOURNE.FlipSelectionExtourne(i: integer);
begin
  if i <= 0 then
    exit;
  FGrille.OnFlipSelection := nil;
  FGrille.FlipSelection(i);
  FGrille.OnFlipSelection := OnGridFlipSelection;
  if FListeEcr.Detail[i - 1].GetValue('SELECTION') = '+' then
  begin
    FListeEcr.Detail[i - 1].PutValue('SELECTION', '-');
    FGrille.Cells[COL_SELECTION,i] := '-';
  end
  else
  begin
    FListeEcr.Detail[i - 1].PutValue('SELECTION', '+');
    FGrille.Cells[COL_SELECTION,i] := '+';    
  end;
end;

procedure TOF_EXTOURNE.ChargeAnalytique(T: TOB);
var
  i: integer;
  TAxe: TOB;
  Q: TQuery;
  St: string;
begin
  AlloueAxe( T ) ;
  for i := 1 to MAXAXE do
  begin
    TAxe :=  T.Detail[i-1] ; // SBO 25/01/2006
    St := 'SELECT * FROM ANALYTIQ WHERE Y_EXERCICE="' + T.GetValue('E_EXERCICE')
      + '"'
      + ' AND Y_JOURNAL="' + T.GetValue('E_JOURNAL') + '"'
      + ' AND Y_DATECOMPTABLE="' +
        USDateTime(StrToDate(T.GetValue('E_DATECOMPTABLE'))) + '"'
      + ' AND Y_NUMEROPIECE=' + IntToStr(T.GetValue('E_NUMEROPIECE'))
      + ' AND Y_NUMLIGNE=' + IntToStr(T.GetValue('E_NUMLIGNE'))
      + ' AND Y_QUALIFPIECE="' + T.GetValue('E_QUALIFPIECE') + '"'
      + ' AND Y_AXE="' + 'A' + IntToStr(i) + '"'
      + ' ORDER BY Y_NUMVENTIL';
    Q := OpenSQL(St, True);
    TAxe.LoadDetailDB('ANALYTIQ', '', '', Q, True);
    Ferme(Q);
  end;
end;

procedure TOF_EXTOURNE.FormKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if (csDestroying in Ecran.ComponentState) then
    Exit;
  if (key =  VK_cherche)  then OnChercheClick(nil)
  else if (key =  VK_F5) then FListeDblClick(nil) {YMO 08/03/07 FQ18703 F11 par défaut de PGI court-circuité (VK_choixmul)}
  else if (key =  vk_liste) then
  begin
    If FGrille.Focused then
    begin
      TPageControl ( GetControl ('PAGES') ).SetFocus ;
      Hent1.NextControl(Ecran) ;
    end else FGrille.SetFocus ;
  end
  else if (key =  vk_applique) then
  begin
    Key:=0 ;
    OnChercheClick(nil);
    if FGrille.CanFocus then FGrille.SetFocus ;
  end
  else if (key =  vk_valide) then
  begin
    Key:=0 ;
    OnLanceExtourneClick(nil);
  end
  else if FGrille.Focused then
  begin
    if ((key =  Ord('A')) and (ssCtrl in Shift)) then
    begin
      Key := 0;
      TToolBarButton97(GetControl('BTOUTSEL')).Click;
    // Pièce précédente
    end else if (key =  Ord('P'))  then
    begin
      Key := 0;
      PiecePrecedente;
    // Pièce suivante
    end else if (key =  Ord('S'))  then
    begin
      Key := 0;
      PieceSuivante;
    end else if (key =  Ord('D')) then
    begin
      Key := 0;
      DebutPiece;
    end else if ((key =  Ord('N')) and (ssCtrl in Shift)) then
    begin
      Key := 0;
      TToolBarButton97(GetControl('BSELECTIONFIN')).Click;
    end
    else if (key =  VK_F11)  then
    begin
      Key := 0;
      TPopupMenu(GetControl('POPF11')).Popup (Mouse.CursorPos.x,Mouse.CursorPos.y);
    end;
  end;
end;

procedure TOF_EXTOURNE.FaireRapport;
var
  i: integer;
  OHisto: TOBM;
  Q: TQuery;
begin
  VideListe(FRapport);
  for i := 0 to FListePiece.Detail.Count - 1 do
  begin
    Q := OpenSQL('SELECT * FROM ECRITURE WHERE E_EXERCICE="' +
      QuelExoDt(FDateGene) + '"'
      + ' AND E_JOURNAL="' + FJournalGene + '"'  {FQ19892 YMO 30/03/07 Choix du journal de génération}
      + ' AND E_QUALIFPIECE="' + FTypeGene + '"'
      + ' AND E_NUMEROPIECE=' +
        IntToStr(FListePiece.Detail[i].GetValue('NEWPIECE'))
      + ' AND E_PERIODE=' + IntToStr(GetPeriode(FDateGene))
      + ' AND E_NUMLIGNE=1', True);
    OHisto := TOBM.Create(EcrGen, '', False);
    OHisto.ChargeMvt(Q);
    FRapport.Add(OHisto);
    Ferme(Q);
  end;
end;


procedure TOF_EXTOURNE.OnSELECTION_PIECEClick(Sender: TObject);
begin
  inherited;
  if GetCheckBoxState('SELECTION_PIECE')=cbChecked then
    RazSelection;
end;

procedure TOF_EXTOURNE.InitMultiCriteres;
var Jal : string;
begin
  // Initialisation du multicritères
  SetControlText('E_DEVISE', V_PGI.DevisePivot);
  if VH^.CPExoRef.Code <> '' then
  begin
    SetControlText('E_EXERCICE', VH^.CPExoRef.Code);
    // Préselection dernier mois exercice
    SetControlText('E_DATECOMPTABLE', DateToStr(VH^.CPExoRef.Deb));
    SetControlText('E_DATECOMPTABLE_', DateToStr(VH^.CPExoRef.Fin));
  end
  else
  begin
    SetControlText('E_EXERCICE', VH^.Encours.Code);
    SetControlText('E_DATECOMPTABLE', DateToStr(VH^.Encours.Deb));
    SetControlText('E_DATECOMPTABLE_', DateToStr(VH^.Encours.Fin));
  end;
  SetControlText('E_QUALIFPIECE', 'N');
  SetControlText('E_DATEECHEANCE', StDate1900);
  SetControlText('E_DATEECHEANCE_', StDate2099);
  // On se positionne sur le premier journal de type OD
  Jal := GetColonneSQL ('JOURNAL','J_JOURNAL','J_NATUREJAL="OD"');
  if Jal <> '' then THValComboBox(GetControl('E_JOURNAL')).Value := Jal
    else THValComboBox(GetControl('E_JOURNAL')).ItemIndex := 0;
  if VH^.EtablisDefaut <> '' then
    SetControlText('E_ETABLISSEMENT', VH^.EtablisDefaut)
  else
    THValComboBox(GetControl('E_ETABLISSEMENT')).ItemIndex := 0;
  PositionneEtabUser(GetControl('E_ETABLISSEMENT'), False); // 15090
end;

procedure TOF_EXTOURNE.RazSelection;
var i : integer;
begin
  inherited;
  for i := 1 to FGrille.RowCount - 1 do
  begin
    if i <= FListeEcr.Detail.Count then
      FListeEcr.Detail[i-1].AddChampSupValeur('SELECTION','-');
    FGrille.Cells[COL_SELECTION,i] := '-';
    FGrille.OnFlipSelection := nil;
    if FGrille.IsSelected(i) then
      FGrille.FlipSelection(i);
    FGrille.OnFlipSelection := OnGridFlipSelection;
  end;
  MajCumulSelection;
end;

procedure TOF_EXTOURNE.MajZoneCumul;
var HN : THNumEdit ;
    HC : THCritMaskEdit;
    PTotaux, PSelection : THPanel;
    i , LeftDebit, LeftCredit, WidthDebit, WidthCredit : integer;
    R : TRect;
begin
  // calcul des positions des zones de cumul
  R := FGrille.CellRect(COL_DEBIT,0) ;
  LeftDebit := R.Left ;
  WidthDebit := R.Right-R.Left ;
  R := FGrille.CellRect(COL_CREDIT,0) ;
  LeftCredit := R.Left ;
  WidthCredit := R.Right-R.Left ;

  // Mise à jour des totaux
  PTotaux := THPanel ( GetControl ('PTOTAUX') );
  if PTotaux <> nil then
  begin
    for i := 0 to PTotaux.ControlCount - 1 do
    begin
      if PTotaux.Controls[i] is THNumEdit then
      begin
        HN := THNumEdit ( PTotaux.Controls[i] );
        if ExtractSuffixe ( HN.Name ) = 'DEBIT' then
        begin
          HN.Left := LeftDebit;
          HN.Width := WidthDebit;
          HN.Value := FTotalDebit
        end else
        if ExtractSuffixe ( HN.Name ) = 'CREDIT' then
        begin
          HN.Left := LeftCredit;
          HN.Width := WidthCredit;
          HN.Value := FTotalCredit
        end;
      end  else if PTotaux.Controls[i] is THCritMaskEdit then
      begin
        HC := THCritMaskEdit ( PTotaux.Controls[i] ) ;
        HC.Text := ' '+TraduireMemoire('Totaux')+' ( '+IntToStr(FTotalLigne)+' lignes )';
      end;
    end;
  end;
  // Mise à jour des cumuls de la sélection
  PSelection := THPanel ( GetControl ('PSELECTION') );
  if PSelection <> nil then
  begin
    for i := 0 to PSelection.ControlCount - 1 do
    begin
      if PSelection.Controls[i] is THNumEdit then
      begin
        HN := THNumEdit ( PSelection.Controls[i] );
        if ExtractSuffixe ( HN.Name ) = 'DEBIT' then
        begin
          if Arrondi(FSelectionDebit-FSelectionCredit,6) <> 0 then
            HN.Font.Color := clRed
          else HN.Font.Color := clGreen;
          HN.Left := LeftDebit;
          HN.Width := WidthDebit;
          HN.Value := FSelectionDebit
        end else
        if ExtractSuffixe ( HN.Name ) = 'CREDIT' then
        begin
          if Arrondi(FSelectionDebit-FSelectionCredit,6) <> 0 then
            HN.Font.Color := clRed
          else HN.Font.Color := clGreen;
          HN.Left := LeftCredit;
          HN.Width := WidthCredit;
          HN.Value := FSelectionCredit;
        end;
      end  else if PSelection.Controls[i] is THCritMaskEdit then
      begin
        HC := THCritMaskEdit ( PSelection.Controls[i] ) ;
        HC.Text := ' '+TraduireMemoire('Selection')+' ( '+IntToStr(FSelectionLigne)+' lignes )';
      end;
    end;
  end;
end;

procedure TOF_EXTOURNE.OnEcranResize(Sender: TObject);
begin
  inherited;
  MajZoneCumul;
end;

procedure TOF_EXTOURNE.OnChangeLargeurColonneGrille(Sender: TObject);
begin
  inherited;
  MajZoneCumul;
end;

procedure TOF_EXTOURNE.OnToutSelectionnerClick(Sender: TOBject);
begin
  inherited;
  ToutSelectionner (0);
end;

procedure TOF_EXTOURNE.MajCumulSelection;
begin
  // calcul des cumuls
  FSelectionLigne := Trunc (FListeEcr.Somme('SELECTION',['SELECTION'],['+'],False,True));
  FSelectionDebit := FListeEcr.Somme('E_DEBIT',['SELECTION'],['+'],False);
  FSelectionCredit := FListeEcr.Somme('E_CREDIT',['SELECTION'],['+'],False);
  MajZoneCumul;
end;

procedure TOF_EXTOURNE.CalculNumGroupeEcrEquivalent;
var i, iNumGroupeEcr, iRef : integer;
    CumulDebit, CumulCredit : double;
    bNewPiece : boolean;
begin
  // Calcul de NUMGROUPEECR équivalent pour le mode libre
  iNumGroupeEcr := 0; iRef := -1;
  CumulDebit := 0; CumulCredit := 0;
  for i:=0 to FListeEcr.Detail.Count - 1 do
  begin
    bNewPiece :=  not MemePiece (iRef,i, True);
    if (i>0) and (not bNewPiece) and ( i < FListeEcr.Detail.Count - 1 ) then
    begin
      bNewPiece :=  (Arrondi(CumulDebit - CumulCredit,6)=0) and
        ((FListeEcr.Detail[i+1].GetValue('E_DEBIT')<>0) or (FListeEcr.Detail[i+1].GetValue('E_CREDIT')<>0));
    end;
    if bNewPiece then
    begin
      Inc (iNumGroupeEcr,1);
      iRef := i;
      CumulDebit := FListeEcr.Detail[i].GetValue('E_DEBIT');
      Cumulcredit := FListeEcr.Detail[i].GetValue('E_CREDIT');
    end else
    begin
      CumulDebit := CumulDebit + FListeEcr.Detail[i].GetValue('E_DEBIT');
      CumulCredit := CumulCredit + FListeEcr.Detail[i].GetValue('E_CREDIT');
    end;
     {FQ19187 YMO 12/03/07 Sélection en mode libre par période on prend ttes les lignes}
     {Si l'on veut revenir dans l'ancien mode (par groupe soldé), il faut dégriser ces comment}
    //FListeEcr.Detail[i].AddChampSupValeur('NUMGROUPEECRLIB',iNumGroupeEcr);
  end;
end;

procedure TOF_EXTOURNE.OnE_JOURNALChange(Sender: TObject);
begin
  inherited;
  FJournal := GetControlText('E_JOURNAL');
  FModeSaisie := GetColonneSQL('JOURNAL', 'J_MODESAISIE', 'J_JOURNAL="' + FJournal + '"');
end;

procedure TOF_EXTOURNE.PieceSuivante;
var i : integer;
begin
  // Recherche de la pièce suivante
  for i:=FGrille.Row+1 to FGrille.RowCount - 1 do
  begin
    if not MemePiece( FGrille.Row-1,i-1) then
    begin
      FGrille.Row := i;
      break;
    end;
  end;
end;

procedure TOF_EXTOURNE.PiecePrecedente;
var i : integer;
begin
  DebutPiece;
  // Recherche de la pièce précédente
  for i:=FGrille.Row-1 downto 1 do
  begin
    if not MemePiece( FGrille.Row-1,i-1) then
    begin
      FGrille.Row := i;
      break;
    end;
  end;
end;

procedure TOF_EXTOURNE.DebutPiece;
var i : integer;
begin
  // Recherche de la première ligne de la pièce
  for i:=FGrille.Row-1 downto 1 do
  begin
    if not MemePiece( FGrille.Row-1,i-1) then
    begin
      FGrille.Row := i+1;
      break;
    end;
  end;
  if i=0 then FGrille.Row := 1;
end;


procedure TOF_EXTOURNE.ToutSelectionner(DepuisLigne : integer);
var i : integer;
    cNew, cOld : Char;
begin
  inherited;
  if DepuisLigne = 0 then DepuisLigne := 1;
  FGrille.OnFlipSelection := nil;
  if FGrille.Row > (FListeEcr.Detail.Count) then exit; {YMO FQ18703 30/03/07 Dernière ligne active}
  cOld := '-'; cNew := '-';
  if FListeEcr.Detail[FGrille.Row-1].GetValue('SELECTION')='+' then cOld:='+'
  else cNew := '+';
  for i:=DepuisLigne to FGrille.RowCount - 1 do
  begin
    if FListeEcr.Detail[i-1].GetValue('SELECTION')=cOld then
    begin
      FListeEcr.Detail[i - 1].PutValue('SELECTION', cNew);
      FGrille.Cells[COL_SELECTION,i] := cNew;
      FGrille.FlipSelection (i);
    end;
  end;
  FGrille.OnFlipSelection := OnGridFlipSelection;
  MajCumulSelection;
end;

procedure TOF_EXTOURNE.OnSelectionFinClick(Sender: TOBject);
begin
  inherited;
  if GetCheckBoxState('SELECTION_PIECE') = cbUnChecked then
    ToutSelectionner (FGrille.Row)
  else
  begin
    DebutPiece;
    ToutSelectionner(FGrille.Row);
  end;
end;

procedure TOF_EXTOURNE.OnDebutPieceClick(Sender: TObject);
begin
  DebutPiece;
end;

procedure TOF_EXTOURNE.OnPiecePrecedenteClick(Sender: TObject);
begin
  PiecePrecedente;
end;

procedure TOF_EXTOURNE.OnPieceSuivanteClick(Sender: TObject);
begin
  PieceSuivante;
end;


function TOF_EXTOURNE.ConstruitRequete: String;
var St  : String ;
    Pf  : String ;
begin

  if (GetCheckBoxState('SELECTION_PIECE') = cbChecked)
    then Pf  := 'E1.'
    else Pf  := '' ;

  // **** Les Champs ****
  St := '@NOCONF@SELECT DISTINCT ' + Pf + 'E_NATUREPIECE, ' + Pf + 'E_DATECOMPTABLE, ' + Pf + 'E_NUMEROPIECE, '
                           + Pf + 'E_GENERAL, '     + Pf + 'E_AUXILIAIRE, '    + Pf + 'E_REFINTERNE, '
                           + Pf + 'E_LIBELLE, '     + Pf + 'E_DEBIT, '         + Pf + 'E_CREDIT, '
                           + Pf + 'E_PERIODE, '     + Pf + 'E_NUMLIGNE,'       + Pf + 'E_JOURNAL, '
                           + Pf + 'E_QUALIFPIECE, ' + Pf + 'E_EXERCICE,'       + Pf + 'E_NUMECHE, '
                           + Pf + 'E_NUMGROUPEECR ';  //FQ19187 YMO 12/03/07 Sélection en mode bordereau par groupe et non pas folio

  // *** Les tables ***
  St :=   St + 'FROM ECRITURE' ;
  if (GetCheckBoxState('SELECTION_PIECE') = cbChecked) then
    St := St  + ' E1 LEFT JOIN ECRITURE E2 ON E1.E_JOURNAL=E2.E_JOURNAL AND E1.E_EXERCICE=E2.E_EXERCICE'
                   + ' AND E1.E_DATECOMPTABLE=E2.E_DATECOMPTABLE AND E1.E_NUMEROPIECE=E2.E_NUMEROPIECE'
                   + ' AND E1.E_QUALIFPIECE=E2.E_QUALIFPIECE' ;

  // **** Les Conditions ****
  St := St  + ' WHERE ' + GetWhereSQL ;

  // *** Tri ***
  St := St + ' ORDER BY ' + Pf + 'E_PERIODE, ' + Pf + 'E_NUMEROPIECE, ' + Pf + 'E_NUMLIGNE';

  result := St ;

 {Ancienne requête

    // creation de la requete !
    St := 'SELECT * FROM ECRITURE WHERE ';
    // BPY le 21/04/04 => Fiche FFF n° 125 : recup de pieces !
    if (GetCheckBoxState('SELECTION_PIECE') = cbChecked) then
    begin
      // JP 06/07/05 : FQ 16119 : Oracle 7 ne connait pas le cast
      if V_PGI.Driver = dbORACLE7 then begin
        St := St + 'E_JOURNAL || E_EXERCICE || TO_CHAR(E_DATECOMPTABLE, "DD/MM/YYYY") || TO_CHAR(E_NUMEROPIECE) || E_QUALIFPIECE';
        St := St + ' IN (SELECT E_JOURNAL || E_EXERCICE || TO_CHAR(E_DATECOMPTABLE, "DD/MM/YYYY") || TO_CHAR(E_NUMEROPIECE) || E_QUALIFPIECE FROM ECRITURE WHERE ';
      end
      else begin
        St := St + 'E_JOURNAL || E_EXERCICE || CAST(E_DATECOMPTABLE as varchar(30)) || CAST(E_NUMEROPIECE as varchar(30)) || E_QUALIFPIECE';
        St := St + ' IN (SELECT E_JOURNAL || E_EXERCICE || CAST(E_DATECOMPTABLE as varchar(30)) || CAST(E_NUMEROPIECE as varchar(30)) || E_QUALIFPIECE FROM ECRITURE WHERE ';
      end;
    end;
    // Critère exercice
    St := St + ' E_EXERCICE="' + GetControlText('E_EXERCICE') + '" ';
    // Critère journal
    St := St + 'AND E_JOURNAL="' + GetControlText('E_JOURNAL') + '" ';
    // Critère nature de pièce
    if GetControlText('E_NATUREPIECE') <> '' then St := St + 'AND E_NATUREPIECE="' + GetControlText('E_NATUREPIECE') + '" ';
    // Critère date comptable
    St := St + 'AND E_DATECOMPTABLE>="' + USDate(THCritMaskEdit(GetControl('E_DATECOMPTABLE'))) + '" ';
    St := St + 'AND E_DATECOMPTABLE<="' + USDate(THCritMaskEdit(GetControl('E_DATECOMPTABLE_'))) + '" ';
    // Critère date échéance
    St := St + 'AND E_DATEECHEANCE>="' + USDate(THCritMaskEdit(GetControl('E_DATEECHEANCE'))) + '" ';
    St := St + 'AND E_DATEECHEANCE<="' + USDate(THCritMaskEdit(GetControl('E_DATEECHEANCE_'))) + '" ';
    // Critère numéro de pièce
    if THCritMaskEdit(GetControl('E_NUMEROPIECE')).Text <> '' then St := St + 'AND E_NUMEROPIECE>=' + GetControlText('E_NUMEROPIECE');
    if THCritMaskEdit(GetControl('E_NUMEROPIECE_')).Text <> '' then St := St + 'AND E_NUMEROPIECE<=' + GetControlText('E_NUMEROPIECE_');
    // Critère référence interne
    if GetControlText('E_REFINTERNE') <> '' then St := St + 'AND E_REFINTERNE="' + GetControlText('E_REFINTERNE') + '" ';
    // Critère compte général
    if GetControlText('E_GENERAL') <> '' then St := St + 'AND E_GENERAL="' + GetControlText('E_GENERAL') + '" ';
    // Critère compte auxiliaire
    if GetControlText('E_AUXILIAIRE') <> '' then St := St + 'AND E_AUXILIAIRE="' + GetControlText('E_AUXILIAIRE') + '" ';
    // Critère validée
    if GetCheckBoxState('E_VALIDE') = cbChecked then St := St + 'AND E_VALIDE="X" '
    else if GetCheckBoxState('E_VALIDE') = cbUnChecked then St := St + 'AND E_VALIDE="-" ';
    // Critère déjà extournée
    if GetCheckBoxState('DEJAEXT') = cbChecked then St := St + 'AND E_QUALIFORIGINE="URN"'
    else if GetCheckBoxState('DEJAEXT') = cbUnChecked then St := St + 'AND ((E_QUALIFORIGINE<>"TP" AND E_QUALIFORIGINE<>"URN") OR E_QUALIFORIGINE="") '
    else if GetCheckBoxState('DEJAEXT') = cbGrayed then St := St + 'AND (E_QUALIFORIGINE<>"TP" OR E_QUALIFORIGINE="") ';
    // Critère établissement
    if GetControlText('E_ETABLISSEMENT') <> '' then St := St + 'AND E_ETABLISSEMENT="' + GetControlText('E_ETABLISSEMENT') + '" ';
    // Critère type d'écriture
    if GetControlText('E_QUALIFPIECE') <> '' then St := St + 'AND E_QUALIFPIECE="' + GetControlText('E_QUALIFPIECE') + '" ';
    // Critère code devise
    if GetControlText('E_DEVISE') <> '' then St := St + 'AND E_DEVISE="' + GetControlText('E_DEVISE') + '" ';
    // Critères fixes
    St := St + ' AND E_NATUREPIECE<>"ECC" AND E_ECRANOUVEAU="N" AND E_NUMECHE<=1 AND E_TRESOLETTRE<>"X" AND E_CREERPAR<>"DET"';
    // BPY le 21/04/04 => Fiche FFF n° 125 : recup de pieces !
    if (GetCheckBoxState('SELECTION_PIECE') = cbChecked) then St := St + ') ';

    // Tri des écritures
    St := St + ' ORDER BY E_PERIODE,E_NUMEROPIECE,E_NUMLIGNE';
}

end;


{***********A.G.L.***********************************************
Auteur  ...... : Christophe Ayel
Créé le ...... : 22/12/2006
Modifié le ... : 22/12/2006
Description .. : Construction du where de la requête des écritures à afficher
Suite ........ : dans le multicritères
Suite ........ : FQ 19371 - CA - 22/12/2006 - Ajout d'un espace avant
Suite ........ : chaque AND pour problème de requête sur DB2
Mots clefs ... :
*****************************************************************}
function TOF_EXTOURNE.GetWhereSQL: String;
var St : String ;
    Pf : String ;
begin

  if (GetCheckBoxState('SELECTION_PIECE') = cbChecked)
    then Pf := 'E2.'
    else Pf := '' ;

  // Critère exercice
  St := St + Pf + 'E_EXERCICE="' + GetControlText('E_EXERCICE') + '" ';

  // Critère journal
  St := St + ' AND ' + Pf + 'E_JOURNAL="' + GetControlText('E_JOURNAL') + '" ';

  // Critère nature de pièce
  if GetControlText('E_NATUREPIECE') <> '' then
    St := St + ' AND '+ Pf + 'E_NATUREPIECE="' + GetControlText('E_NATUREPIECE') + '" ';

  // Critère date comptable
  St := St + ' AND ' + Pf + 'E_DATECOMPTABLE>="' + USDate(THCritMaskEdit(GetControl('E_DATECOMPTABLE'))) + '" ';
  St := St + ' AND ' + Pf + 'E_DATECOMPTABLE<="' + USDate(THCritMaskEdit(GetControl('E_DATECOMPTABLE_'))) + '" ';

  // Critère date échéance
  St := St + ' AND ' + Pf + 'E_DATEECHEANCE>="' + USDate(THCritMaskEdit(GetControl('E_DATEECHEANCE'))) + '" ';
  St := St + ' AND ' + Pf + 'E_DATEECHEANCE<="' + USDate(THCritMaskEdit(GetControl('E_DATEECHEANCE_'))) + '" ';

  // Critère numéro de pièce
  if THCritMaskEdit(GetControl('E_NUMEROPIECE')).Text <> '' then
    St := St + ' AND ' + Pf + 'E_NUMEROPIECE>=' + GetControlText('E_NUMEROPIECE');
  if THCritMaskEdit(GetControl('E_NUMEROPIECE_')).Text <> '' then
    St := St + ' AND ' + Pf + 'E_NUMEROPIECE<=' + GetControlText('E_NUMEROPIECE_');

  // Critère référence interne
  if GetControlText('E_REFINTERNE') <> '' then
    St := St + ' AND ' + Pf + 'E_REFINTERNE="' + GetControlText('E_REFINTERNE') + '" ';

  // Critère compte général
  if GetControlText('E_GENERAL') <> '' then
    St := St + ' AND ' + Pf + 'E_GENERAL="' + GetControlText('E_GENERAL') + '" ';

  // Critère compte auxiliaire
  if GetControlText('E_AUXILIAIRE') <> '' then
    St := St + ' AND ' + Pf + 'E_AUXILIAIRE="' + GetControlText('E_AUXILIAIRE') + '" ';  {FP 10/11/2005 FQ16598: Il faut remplacer la référence à E2 par pf}

  // Critère validée
  if GetCheckBoxState('E_VALIDE') = cbChecked then
    St := St + ' AND ' + Pf + 'E_VALIDE="X" '
  else if GetCheckBoxState('E_VALIDE') = cbUnChecked then
    St := St + ' AND ' + Pf + 'E_VALIDE="-" ';

  // Critère déjà extournée
  if GetCheckBoxState('DEJAEXT') = cbChecked then
    St := St + ' AND ' + Pf + 'E_QUALIFORIGINE="URN"'
  else if GetCheckBoxState('DEJAEXT') = cbUnChecked then
    St := St + ' AND ((' + Pf + 'E_QUALIFORIGINE<>"TP" AND ' + Pf + 'E_QUALIFORIGINE<>"URN") OR ' + Pf + 'E_QUALIFORIGINE="") '
  else if GetCheckBoxState('DEJAEXT') = cbGrayed then
    St := St + ' AND (' + Pf + 'E_QUALIFORIGINE<>"TP" OR ' + Pf + 'E_QUALIFORIGINE="") ';

  // Critère établissement
  if GetControlText('E_ETABLISSEMENT') <> '' then
    St := St + ' AND ' + Pf + 'E_ETABLISSEMENT="' + GetControlText('E_ETABLISSEMENT') + '" ';

  // Critère type d'écriture
  if GetControlText('E_QUALIFPIECE') <> '' then
    St := St + ' AND ' + Pf + 'E_QUALIFPIECE="' + GetControlText('E_QUALIFPIECE') + '" ';

  // Critère code devise
  if GetControlText('E_DEVISE') <> '' then
    St := St + ' AND ' + Pf + 'E_DEVISE="' + GetControlText('E_DEVISE') + '" ';

  // Critères fixes
  St := St + ' AND ' + Pf + 'E_NATUREPIECE<>"ECC" AND '
                     + Pf + 'E_ECRANOUVEAU="N" AND '
                     + Pf + 'E_NUMECHE<=1 AND '
                     + Pf + 'E_TRESOLETTRE<>"X" AND '
                     + Pf + 'E_CREERPAR<>"DET"' ;

  {JP 06/06/07 : FQ 20307 : on débranche la gestion des confidentiels, pour la gérer à la main pour
                 éviter sous Oracle un message "Champ ambigu" du fait de l'absence de préfixe sur E_CONFIDENTIEL}
  if V_PGI.Confidentiel <> '1' then
    St := St + ' AND ' + Pf + 'E_CONFIDENTIEL <= "' + IntToStr(V_PGI.NiveauAccesConf) + '"';

  result := St ;

end;

{b FP 10/11/2005 FQ16598}
procedure TOF_EXTOURNE.SetCritModified(const Value: Boolean);
var
  BCherche: TToolBarButton97;
begin
  if (FCritModified<>Value) and (iCritGlyph<>nil) and (iCritGlyphModified<>nil) then
  begin
    BCherche := TToolBarButton97(GetControl('BCHERCHE'));
    if Value then
      BCherche.Glyph := iCritGlyphModified.Picture.BitMap
    else
      BCherche.Glyph := iCritGlyph.Picture.BitMap;
  end;
  FCritModified := Value;
end;

procedure TOF_EXTOURNE.CritereChange(Sender: TObject);
begin
  CritModified := True; {Indique qu'un critère à changer}
end;

procedure TOF_EXTOURNE.InitCritereChange(Parent: TControl);
  // Pour critères avec groubbox ou panel
  function MyParentIsTTabSheet(C: TControl): Boolean;
  begin
    Result := False;
    while C.Parent <> nil do
    begin
      if C.Parent is TTabsheet then
      begin
        Result := True;
        Break;
      end;
      C := TControl(C.Parent);
    end;
  end;
var
  i:     integer;
  Name:  String;
begin
  {Identique à la méthode InitAutoSearch de l'unité Mul.pas}
  for i := 0 to Parent.ComponentCount - 1 do
  begin
    if (Parent.Components[i] is TControl) and
       MyParentIsTTabSheet((Parent.Components[i] as TControl)) then
    begin
      Name := (Parent.Components[i] as TControl).Name;
      if (Parent.Components[i] is THEdit) and (not Assigned(THEdit(Parent.Components[i]).OnChange)) then
      begin
        THEdit(Parent.Components[i]).OnChange := CritereChange;  {Remplace l'évènement}
      end;
      if (Parent.Components[i] is THValComboBox) and (not Assigned(THValComboBox(Parent.Components[i]).OnClick)) then
      begin
        THValComboBox(Parent.Components[i]).OnClick := CritereChange;
      end;
      if (Parent.Components[i] is TCheckBox) and (not Assigned(TCheckBox(Parent.Components[i]).OnClick)) then
      begin
        TCheckBox(Parent.Components[i]).OnClick := CritereChange;
      end;
    end;
  end;
end;
{e FP 10/11/2005 FQ16598}
initialization
  registerclasses([TOF_EXTOURNE]);
end.

