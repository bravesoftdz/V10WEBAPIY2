unit CopierCollerUtil;

interface
uses {$IFDEF VER150} variants,{$ENDIF} sysutils,classes,windows,messages,controls,forms,hmsgbox,stdCtrls,clipbrd,nomenUtil,
     HCtrls,SaisUtil,HEnt1,Ent1,EntGC,UtilPGI,UTOB,HTB97,FactUtil,FactComm,Menus,
     FactTOB, FactPiece, FactArticle,
{$IFDEF EAGLCLIENT}
  maineagl,
{$ELSE}
  fe_main,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ENDIF}
  FactVariante,FactEmplacementLivr,
{$IFDEF BTP}
     BTPUtil,FactOuvrage, UTobDebug,Utilfichiers, UtilsMetresXLS,
{$ENDIF}
     HRichOLE,UTOFBTMEMORISATION,CBPPath,UtilNumParag,UEntCommun;

const MAXITEMS = 7;

type TCopieColle = class(TObject)
private
    fform : TForm;
    fgrid : THGrid;
    TOBimp : TOB;
    ftransfert : TToolWindow97;
    fmemo : TMemo;
    Mcopy : TstringList;
    RowSelected : integer;
    procedure SetGrid(const Value: THGrid); virtual;
protected
		TempFile : string;
    function    RecupTob(Atob: TOB): Integer;
    procedure   ExportTob(TOBLoc : TOB; ChampBlocNote : string; Newgestion : boolean=false);
public
    constructor create (Parent : TForm);
    destructor  destroy; override;
    property    Grid : THGrid read fgrid write SetGrid;
    Property    TOBTmp : TOB read TOBIMP write TOBImp;
end;

type TCopieColleDoc = class (TcopieColle)
private
		FF : TForm;
    TOBLR : TOB; // ligne a raffraichier
    TOBPiece,TOBArticles,TOBAffaire,TOBCatalogu,TOBCpta,TOBTiers,TOBAnaS,TOBAnaP,TOBAdresses : TOB;
    TOBSerie,TOBDesLots,TobNomen,TOBOuvrage : TOB;
    TOBBases,TOBBasesL : TOB;
    SaContexte  : TModeAff;
    GereLot,GereSerie,GereStock : boolean ;
    DEV         : RDevise;
    fTypeInfo   : string;
    RowCollage  : integer;
    ColCollage  : integer;
    POPGS       : TPopupMenu;
    fcreatedPop : boolean;
    fMaxItems   : integer;
    MesMenuItem : array[0..MAXITEMS] of TMenuItem;
    venteAchat  : string;
    TTNUMP      : TNumParag;
    TobVarDoc   : TOB;
    TobMetres   : TOB;
    TTObPiece   : TOB;
    TheMetredoc : TMetreArt;
    procedure AjouteLigneSel (TOBSel : TOB; Indice: Integer; WithMemo : boolean=false);
    procedure SetGrid(const Value: THGrid); override;
    function AjouteLigneImport(TOBLoc: TOB; var Niveau, Arow: integer; var SauteAFinPar, ArticleNonOk: boolean): boolean;
    procedure InsertDebParagrapheImp(TOBLOC: TOB; niveau, Arow: integer);
    procedure InsertFinParagrapheImp(TOBLOC: TOB; niveau, Arow: integer);
    function InsertLigneArticleImp(TOBLOC: TOB; var Arow: integer;Niveau: integer; var ArticleNonOk: boolean): boolean;
    procedure InsertLigneCommentaireImp(TOBLoc: TOB; Arow,Niveau: Integer);
    procedure InsertLigneTotalImp(TOBLoc: TOB; Arow, Niveau: integer);
    function IsArticleOk(Arow: integer): boolean;
    procedure Collagedonnee (NewGestion : boolean=false);
    function IsTypeTOB: boolean;
    procedure RappellerMemo;
    procedure EnregistreDonne;
    procedure CopieDonnee;
    procedure CouperDonnee;
    procedure DefiniMenuPop(Parent : Tform);
    procedure GSContextPopup(Sender: TObject; MousePos: TPoint; var Handled: Boolean);
    procedure GScopierClick(Sender: TObject);
    procedure GSCouperClick (Sender: TObject);
    procedure GSEnregistrerClick(Sender: TObject);
    procedure GSCollerClick(Sender: TObject);
    procedure GSRappelerClick(Sender: TObject);
    procedure GereImprimable (Sender : TObject);
    procedure MenuEnabled (State : boolean);
    procedure VerifChampsSupsDetail(TOBOUV: TOB);
		procedure VerifChampsSupsLigne(TOBL : TOB);
    procedure HideMenu;
    procedure CopieChampsLigne(Table : string; TOBLOC, Tobl: Tob);
    procedure CopieDonneeOuvrage (TOBDest,TOBOUV, TOBL : TOB; Ok_MetreDoc, Ok_MetreBib, Ok_MetreExcel : Boolean);
    procedure CopieDetailOuvrage (TOBDest, TOBOUV, TOBL: TOB; Ok_MetreDoc, Ok_MetreBib, Ok_MetreExcel : Boolean);
    procedure AjouteUneLigneAraffraichir (Arow: integer);
    function FindLigne(NumOrdre: integer): TOB;
    procedure ReinitLigneFac (TOBL : TOB);
    function CopiesImportable(TOBAIMP: TOB): boolean;
    function AjouteLigneImportOuv(TOBLigneDoc,TOBL,TOBAINS: TOB; var ArticleNonOk : boolean) : boolean;
    function AjouteLigneImportDetailOuv(TOBOUV, TOBLigneDoc, TOBI: TOB; Indice : integer ;var ArticleNonOk: boolean): boolean;
    procedure CalculeLaLigneDoc(TOBL: TOB; WithRecalc: boolean=true);
    function TraiteLigneCopierCollerOuv(TOBD, TOBpiece, TOBL,TOBArticle: TOB; DEV: RDevise; memedoc: boolean): boolean;
    function FindTOBPere(TOBXX: TOB): TOB;

public
    constructor create (Parent : TForm);
    destructor  destroy; override;
    property    LastSelected : integer read RowSelected;
    property    Devise : RDevise read DEV write DEV;
    property    TypeDonnee : string read FtypeInfo write ftypeInfo;
    property    TheNumParag : TNumParag read TTNUMP write TTNUMP;
    procedure   SetTobs (ThePiece,TheSerie,TheDesLots,TheAdresses,TheArticles,TheAffaire,
                         TheCatalogu,TheNomen,TheOuvrage,TheCpta,TheTiers,TheAnaS,TheAnaP,
                         TheBases,TheBasesL, TheVarDoc, TheMetres : TOB; Metredoc : TmetreArt);
    procedure   SetInfoDocument (IfGereLot,IfGereSerie,IfGereStock : boolean; ContexteSaisie : TModeAff; PDEV:RDevise);
    procedure   PositionneCell(Acol,Arow: integer);
    procedure   FlipSelection(Sender: TObject);
    procedure   deselectionneRows;
    procedure   BeforeFlip(Sender: TObject; ARow: Integer;var Cancel: Boolean);
    procedure   Copier;
    procedure   Couper;
    procedure   Enregistrer;
    procedure   Coller;
    procedure   rapeller;
    procedure   Activate;
    procedure SetInfoLigne (TOBl : TOB);
end;

procedure InitChampsSupNULL (TOBAssoc: TOB);

implementation
uses facture
    ,FactCalc
    ,PiecesRecalculs
    ,FactDOmaines
    ,UtilArticle
    ,FactureBtp
    ,factDetOuvrage
    ,UtilsCopierCollerMetresXLS
    ,USpecifPOC
    ;
{ TCopieColle }


function IsSamePiece(TOBDOC,TOBPiece : TOB) : Boolean;
begin
  result := (TOBDOC.GetValue('GL_NATUREPIECEG')=TOBPiece.GetValue('GP_NATUREPIECEG')) and
            (TOBDOC.GetValue('GL_SOUCHE')=TOBPiece.GetValue('GP_SOUCHE')) and
            (TOBDOC.GetValue('GL_NUMERO')=TOBPiece.GetValue('GP_NUMERO')) and
            (TOBDOC.GetValue('GL_INDICEG')=TOBPiece.GetValue('GP_INDICEG'));
end;

procedure ReinitMontantAch (TOBL : TOB);
begin
  TOBL.PutValue('GL_MONTANTFC',0); TOBL.PutValue('GL_MONTANTFR',0); TOBL.PutValue('GL_MONTANTPR',0);
end;

procedure InitCopierCollerDetail (TOBA,TOBD,TOBpiece : TOB);
begin
  // Correction FQ 12072
  TOBD.PutValue('BLO_COEFFC',0);
  TOBD.PutValue('BLO_COEFFR',0);
  TOBD.PutValue('BLO_MONTANTFC',0);
  TOBD.PutValue('BLO_MONTANTFR',0);
  if (TOBPIece.GetValue('GP_COEFFC') <> 0) and (TOBD.GetValue('BLO_NONAPPLICFC')<>'X' ) THEN TOBD.PutValue('BLO_COEFFC',TOBPIece.GetValue('GP_COEFFC'));
  if (TOBPIece.GetValue('GP_COEFFR') <> 0) and (TOBD.GetValue('BLO_NONAPPLICFRAIS')<>'X' ) THEN TOBD.PutValue('BLO_COEFFR',TOBPIece.GetValue('GP_COEFFR'));

  //Why this Mother fucker area not initialize ????
  TOBD.PutValue('BLO_COEFMARG',0);

  if TOBD.GetValue('BLO_NATURETRAVAIL') <> '' then
  begin
    TOBD.PutValue('BLO_NATURETRAVAIL','');
    TOBD.PutValue('BLO_FOURNISSEUR',TOBA.GetValue('GA_FOURNPRINC'));
  end;

end;



constructor TCopieColle.create(Parent: TForm);
begin
  inherited create;
  fform := Parent;
  ftransfert := TToolWindow97.Create (Parent);
  ftransfert.parent := Parent;
  ftransfert.visible := false;
  fmemo := Tmemo.create (ftransfert);
  fmemo.parent := ftransfert;
  fmemo.visible := false;
  Mcopy := TStringList.Create;
  TempFile  := IncludeTrailingBackslash (TCBPPath.GetLocalAppData)+ComputerName+'.TMP';
end;

destructor TCopieColle.destroy;
begin
  fmemo.Free;
  ftransfert.Free;
  Mcopy.Clear;
  Mcopy.free;
  inherited;
end;

procedure TCopieColle.ExportTob(TOBLoc : TOB; ChampBlocNote : string; Newgestion : boolean=false);
var TheText : string;
begin
	if not Newgestion then
  begin
    thetext := TobLoc.SaveToBuffer (true,true,ChampBlocNote);
    MCopy.Add(theText);
  end else
  begin
  	TOBLoc.SaveToBinFile (TempFile,false,true,true,true);
  end;
end;

procedure TCopieColle.SetGrid(const Value: THGrid);
begin
  fgrid := Value;
end;


function TCopieColle.RecupTob(Atob:TOB): Integer;
begin
  Result := 0 ;
  TOBImp := ATob ;
end;

{ TCopieColleDoc }

constructor TCopieColleDoc.create(Parent: TForm);
var ThePop : Tcomponent;
begin
  inherited;
  FF := Parent;
  TOBLR := TOB.Create ('LES LIGNES A RAFFRAICHIR',nil,-1);
  GereLot := false;
  GereSerie := false;
  saContexte := [tModeGridStd];
  RowCollage := -1;
  ColCollage :=-1;
  ThePop := Parent.Findcomponent  ('POPBTP');
  if ThePop = nil then
  BEGIN
    // pas de menu BTP trouve ..on le cree
    POPGS := TPopupMenu.Create(Parent);
    POPGS.Name := 'POPBTP';
    fCreatedPop := true;
  END else
  BEGIN
    fCreatedPop := false;
    POPGS := TPopupMenu(thePop);
  END;
  DefiniMenuPop(Parent);
end;

procedure TCopieColleDoc.DefiniMenuPop (Parent : Tform);
var Indice : integer;

begin
  fMaxItems := 0;
  if not fcreatedPop then
  begin
    MesMenuItem[fMaxItems] := TmenuItem.Create (parent);
    with MesMenuItem[fMaxItems] do
      begin
      Caption := '-';
      end;
    inc (fMaxItems);
  end;
  // Passage en non imprimable ..ou pas
  MesMenuItem[fMaxItems] := TmenuItem.Create (parent);
  with MesMenuItem[fMaxItems] do
    begin
    Name := 'TImprimable';
    Caption := TraduireMemoire ('Ne pas Imprimer');
    OnClick := GereImprimable;
    end;
  MesMenuItem[fMaxItems].ShortCut := ShortCut( Word('I'), [ssCtrl,ssAlt]);
  inc (fMaxItems);
  //
  MesMenuItem[fMaxItems] := TmenuItem.Create (parent);
  with MesMenuItem[fMaxItems] do
    begin
    Name := 'SepImprim';
    Caption := '-';
    end;
  inc (fMaxItems);

  // CTRL+C = Copier
  MesMenuItem[fMaxItems] := TmenuItem.Create (parent);
  with MesMenuItem[fMaxItems] do
    begin
    Name := 'BCOPIER';
    Caption := TraduireMemoire ('Copier');
    OnClick := GSCopierClick;
    end;
  MesMenuItem[fMaxItems].ShortCut := ShortCut( Word('C'), [ssCtrl]);
  inc (fMaxItems);

  // CTRL+X = Couper
{ ---- ON VA ATTENDRE ENCORE UN PEU -------------
  MesMenuItem[fMaxItems] := TmenuItem.Create (parent);
  with MesMenuItem[fMaxItems] do
    begin
    Name := 'BCOUPER';
    Caption := TraduireMemoire ('Couper');
    OnClick := GSCouperClick;
    end;
  MesMenuItem[fMaxItems].ShortCut := ShortCut( Word('X'), [ssCtrl]);
  inc (fMaxItems);
}
  // CTRL+V = Coller
  MesMenuItem[fMaxItems] := TmenuItem.Create (parent);
  with MesMenuItem[fMaxItems] do
    begin
    Name := 'BCOLLER';
    Caption := TraduireMemoire('Coller');
    OnClick := GSCollerClick;
    end;
  MesMenuItem[fMaxItems].ShortCut := ShortCut(Word('V'), [ssCtrl]);
  inc (fMaxItems);

  // CTRL+ALT+S = Enregistrer
  MesMenuItem[fMaxItems] := TmenuItem.Create (parent);
  with MesMenuItem[fMaxItems] do
    begin
    Name := 'BENREGISTRER';
    Caption := TraduireMemoire('Enregistrer');
    OnClick :=GSEnregistrerClick;
    end;
  MesMenuItem[fMaxItems].ShortCut := ShortCut(Word('S'), [ssCtrl,ssAlt]);
  inc (fMaxItems);

  // CTRL+ALT+P = Rappeler
  MesMenuItem[fMaxItems] := TmenuItem.Create (parent);
  with MesMenuItem[fMaxItems] do
    begin
    Name := 'BRAPELLER';
    Caption := TraduireMemoire('Rappeler');
    OnClick := GSRappelerClick;
    end;
  MesMenuItem[fMaxItems].ShortCut := ShortCut(Word('P'), [ssCtrl,ssAlt]);
  inc (fMaxItems);

  for Indice := 0 to fMaxItems -1 do
    begin
      if MesMenuItem [Indice] <> nil then POPGS.Items.Add (MesMenuItem[Indice]);
    end;
end;

procedure TCopieColleDoc.SetInfoDocument (IfGereLot,IfGereSerie,IfGereStock : boolean; ContexteSaisie : TModeAff; PDEV:RDevise);
begin
  GereLot := IfGereLot;
  GereSerie := IfGereSerie;
  GereStock := IfGereStock;
  SaContexte := ContexteSaisie;
  DEV := PDEV;
end;

procedure TCopieColleDoc.SetTobs(ThePiece,TheSerie,TheDesLots,TheAdresses,TheArticles,TheAffaire,
                                 TheCatalogu,TheNomen,TheOuvrage,TheCpta,TheTiers,TheAnaS,TheAnaP,
                                 TheBases,TheBasesL, TheVardoc, TheMetres : TOB; Metredoc : TMetreArt);
begin
  TOBPiece    := ThePiece;
  TOBSerie    := TheSerie;
  TOBDesLots  := TheDesLots;
  TOBAdresses := TheAdresses;
  TOBArticles := TheArticles;
  TOBAffaire  := TheAffaire;
  TOBCatalogu := TheCatalogu;
  TOBOuvrage  := TheOuvrage;
  TOBNomen    := TheNomen;
  TOBCpta     := TheCpta;
  TOBTiers    := TheTiers;
  TOBAnas     := TheAnas;
  TOBAnaP     := TheAnaP;
  TOBBases    := TheBases;
  TOBBasesL   := TheBasesL;
  //Métrés
  TobVarDoc   := TheVardoc;
  TobMetres   := TheMetres;
  TheMetredoc := Metredoc;
// definition par defaut (récup)
  VenteAchat  :=  TOBPiece.getValue('GP_VENTEACHAT');
end;


procedure TCopieColleDoc.Coller;
begin
  TOBLR.clearDetail;
  if fgrid = nil then Exit;
  if Clipboard.HasFormat(CF_TEXT) then BEGIN SendMessage (fgrid.InplaceEditor.Handle,WM_PASTE,0,0); Exit; END;
  if (not IsTypeTOB)  then
     begin
     SendMessage (fgrid.InplaceEditor.Handle,WM_PASTE,0,0);
     exit;
     end;
  {$IFDEF BTP}
//  fmemo.PasteFromClipboard;
  CollageDonnee (true);
  RowCollage := -1;
  ColCollage := -1;
  {$ENDIF}
end;

function TCopieColleDoc.CopiesImportable (TOBAIMP : TOB) : boolean;
var II : integer;
begin
  result := true;
  For II := 0 to TOBIMP.detail.count -1 do
  begin
    if (IsVariante(TOBIMP.detail[II])) or
       (IsSousTotal(TOBIMP.detail[II])) or
       (IsCommentaire(TOBIMP.detail[II])) or
       (IsParagraphe(TOBIMP.detail[II])) then
    begin
      result := false;
      exit;
    end;
  end;
end;

procedure TCopieColleDoc.CalculeLaLigneDoc(TOBL: TOB; WithRecalc : boolean);
var indiceNomen : integer;
		TOBOP : TOB;
    valeurs : T_Valeurs;
    Qte   : Double;
    Dpa   : Double;
    Dpr   : Double;
    PUHT  : Double;
    PUTTC : Double;
begin
  IndiceNomen := TOBL.getValue('GL_INDICENOMEN'); if IndiceNomen = 0 then exit;
  TOBOP := TOBOuvrage.Detail[IndiceNomen-1];
  if Withrecalc then
  begin
    InitTableau (Valeurs);
    CalculeOuvrageDoc (TOBOP,1,1,true,DEV,valeurs,(TOBPIECE.GetValue('GP_FACTUREHT')='X'));
    GetValoDetail (TOBOP); // pour le cas des Article en prix posés
    Qte := TOBL.Getvalue('GL_QTEFACT');
    Dpa   := TOBL.GetValue('GL_DPA');
    Dpr   := TOBL.GetValue('GL_DPR');
    TOBL.Putvalue('GL_MONTANTPAFG',valeurs[10]*Qte);
    TOBL.Putvalue('GL_MONTANTPAFR',valeurs[11]*Qte);
    TOBL.Putvalue('GL_MONTANTPAFC',valeurs[12]*Qte);
    TOBL.Putvalue('GL_MONTANTFG',valeurs[13]*Qte);
    TOBL.Putvalue('GL_MONTANTFR',valeurs[14]*Qte);
    TOBL.Putvalue('GL_MONTANTFC',valeurs[15]*Qte);
    TOBL.Putvalue('GL_MONTANTPA',   Arrondi((Qte * Dpa),V_PGI.okdecV));
    TOBL.Putvalue('GL_MONTANTPR',   Arrondi((Qte * Dpr),V_PGI.okdecV));

    TOBL.Putvalue('GL_DPA',valeurs[16]);
    TOBL.Putvalue('GL_DPR',valeurs[17]);
    //
    TOBL.Putvalue('GL_PUHTDEV',valeurs[2]);
    TOBL.Putvalue('GL_PUTTCDEV',valeurs[3]);
    //
    PUHT  := TOBL.GetValue('GL_PUHTDEV');
    PUTTC := TOBL.GetValue('GL_PUTTCDEV');
    PUHT  := DeviseToPivotEx(PUHT,  DEV.Taux,DEV.quotite,V_PGI.OkdecP);
    PUTTC := DevisetoPivotEx(PUTTC, DEV.taux,DEV.quotite,V_PGI.okdecP);
    //
    TOBL.Putvalue('GL_PUHT',        PUHT);
    TOBL.Putvalue('GL_PUTTC',       PUTTC);
    TOBL.Putvalue('GL_PUHTBASE',    PUHT);
    TOBL.Putvalue('GL_PUTTCBASE',   PUTTC);
    //
    TOBL.Putvalue('GL_DPA',valeurs[0]);
    TOBL.Putvalue('GL_DPR',valeurs[1]);
    TOBL.Putvalue('GL_PMAP',valeurs[6]);
    TOBL.Putvalue('GL_PMRP',valeurs[7]);
    TOBL.putvalue('GL_TPSUNITAIRE',valeurs[9]);
    //
    StockeInfoTypeLigne (TOBL,valeurs)
    //
  end;
  if (TOBL.GetString('GL_PIECEPRECEDENTE')='') and (TOBL.FieldExists('BLF_QTESITUATION')) then
  begin
    Qte   := TOBL.Getvalue('GL_QTEFACT');
    PUHT  := TOBL.GetValue('GL_PUHTDEV');
    //
    TOBL.PutValue('BLF_QTESITUATION', Qte); { NEWPIECE }
    TOBL.PutValue('BLF_MTSITUATION',Arrondi(TOBL.GetValue('BLF_QTESITUATION')*PUHT/TOBL.GEtValue('GL_PRIXPOURQTE'),DEV.decimale));
    TOBL.PutValue('BLF_POURCENTAVANC',100);
  end;
  TOBL.PutValue('GL_RECALCULER', 'X');
  ZeroligneMontant (TOBL);
  TFFacture(FF).CalculheureTot (TOBL);
  TOBPiece.PutValue('GP_RECALCULER', 'X');
  TFFacture(FF).CalculeLaSaisie(-1, -1, True);
end;


function TCopieColleDoc.FindTOBPere (TOBXX : TOB) : TOB;
var II : integer;
begin
  result := nil;
  for II := TOBXX.getIndex downto 0 do
  begin
    if isOuvrage(TOBPiece.detail[II]) and (not IsSousDetail(TOBPiece.detail[II])) then
    begin
      result := TOBPiece.detail[II];
      break;
    end;
  end;
end;


procedure TCopieColleDoc.Collagedonnee (NewGestion : boolean=false);
var TOBLOC,TOBL,TOBligneDoc,TOBXX: TOB;
    Indice,Niveau,Localisation,II : Integer;
    SauteAfinPar,ArticleNonOk : boolean;
    Arow,SavRow,SavCol : integer;
    OldEna,bc,cancel : boolean;
    TOBDLR : TOB;
    DepartIns,FinIns : integer;
    FirstPass : boolean;
    ColleinSousDetail : boolean;
begin
  ColleinSousDetail := false;
  DepartIns := 0 ; FinIns := 0;
  FirstPass := true;
  if RowCollage = -1 then
  begin
    RowCollage := fgrid.row;
    ColCollage := fgrid.col;
  end;
  SavRow := RowCollage;
  SavCol := ColCollage;
  TFFacture(FF).GoToLigne(SavRow,SavCol);
  TOBL:=GetTOBLigne(TOBPiece,Savrow) ;
//  if IsLigneDetail (nil,TOBL) then exit;
  OldEna:=fgrid.SynEnabled ;
  fgrid.SynEnabled:=False ;
  SauteAFinPar := false; ArticleNonOk := false;
  if TOBL = nil then
  begin
    Niveau := 0;
  end
  else
  begin
    Niveau := TOBL.GetValue('GL_NIVEAUIMBRIC');
    // Sur une ligne de début de §, on insère au-dessus donc au niveau précédent
    // VARIANTE
    (* if copy(TOBL.getValue('GL_TYPELIGNE'),1,2)='DP' then Niveau:=Niveau-1; *)
    if IsDebutParagraphe(TOBL) then Niveau:=Niveau-1;
  end;
  if not newGestion then
  begin
    TOBLoadFromBuffer (fmemo.Text,RecupTob);
  end;
  Arow := SavRow;
  if TOBIMP = nil then exit;
  
  TRY
    if (TOBImp.getValue('TYPE') = fTypeInfo) then
    BEGIN
      if IsSousDetail(TOBL) then
      begin
        // Phase de controle
        if CopiesImportable (TOBIMP) then
        begin
          TOBligneDoc := FindTOBPere (TOBL);
          if TOBligneDoc = nil then exit;
          DepartIns := TOBligneDoc.GetIndex +1;
          FinIns := DepartIns;
          SavRow := DepartIns;
          AjouteLigneImportOuv (TOBligneDoc,TOBL,TOBIMP,ArticleNonOk);
          TOBXX := TOB.Create ('UNE LIGNE', TOBLR,-1);
          TOBXX.AddChampSupValeur('NUMORDRE',TOBligneDoc.getInteger('GL_NUMORDRE'));
          ColleinSousDetail := true;
          TOBLigneDoc.PutValue('GL_RECALCULER','X');
          CalculeLaLigneDoc (TOBLigneDoc);
        end else
        begin
          PgiError ('Les paragraphes, sous totaux, commentaires et variantes #13#10 ne peuvent être importées dans un sous détail');
          exit;
        end;
      end else
      begin
        For Indice :=0 to TOBImp.detail.count -1 do
        begin
          TOBLoc := TOBImp.detail[Indice];
          if FIrstPass then
          begin
            DepartIns := Arow;
            FirstPass := false;
          end;
          AjouteLigneImport (TOBLoc,Niveau,Arow,SauteAFinPar,ArticleNonOk);
          FinIns := Arow;
        end;
      end;
    end;
    if ArticleNonOk then TFFacture(fform).HPiece.Execute(44,TFFacture(fform).Caption,'') ;
    Arow := SavRow;
  //NumeroteLignesGC(nil,TOBpiece);
    if ColleinSousDetail then
    begin
      TFFacture(fform).SuppressionDetailOuvrage(TOBLignedoc.getIndex+1,true);
    end;
    if TOBLR.detail.count > 0 then
    begin
      Indice := 0;
      repeat
        TOBDLR := TOBLR.detail[Indice];
        TOBL := FindLigne(TOBDLR.getValue('NUMORDRE'));
        if TOBL <> nil then
        begin
          Localisation := TOBL.getIndex;
            LoadLesLibDetOuvLig (TOBPIece,TOBOuvrage,TOBTiers,TOBAffaire,TOBL,nil,Localisation,DEV, TheMetredoc,TFFacture(fform).AffSousDetailUnitaire);
          ZeroLigneMontant (TOBL);
        end;
        TOBDLR.free;
      until TOBLR.detail.count = 0;
      //
    end;
    JustNumerote (TObpiece,DepartIns-1);
    for II := Arow to FinIns do
    begin
      TTNUMP.SetInfoLigne (TOBPiece,II);
    end;
  //
    if TOBPiece.Detail.Count>=fgrid.RowCount-1 then
    begin
      fgrid.RowCount:=TOBPiece.Detail.Count+2 ;
    end;
    TOBPiece.putValue('GP_RECALCULER','X');
    TFFacture(FForm).CalculeLaSaisie (-1,-1,false,true,DepartIns,FinIns);
    // --
    TFFacture(FForm).GS.BeginUpdate;
    for indice := DepartIns to TOBPiece.Detail.Count - 1 do TFFacture(FForm).AfficheLaLigne(Indice + 1);
    TFFacture(FF).GoToLigne(SavRow,SavCol);
    fGrid.row := SavRow;
    fGRid.col := SavCol;
    fgrid.refresh;
    deselectionneRows;
    TFFacture(FForm).GS.EndUpdate;
    TFFacture(FForm).AfficheLaLigne(fgrid.row);
    TFFacture(fform).GoToLigne (SavRow,SavCol);
    TFFacture(fform).PosValueCell (fgrid.Cells[SavCol,SavRow]) ;
  FINALLY
    TOBIMP.free;
    TOBLR.ClearDetail;
  END;
end;

procedure TCopieColleDoc.SetGrid(const Value: THGrid);
begin
inherited;
{$IFDEF BTP}
fgrid.PopupMenu := POPGS;
fgrid.OnContextPopup := GSContextPopup;
fgrid.OnFlipSelection := FlipSelection;
fgrid.OnBeforeFlip := BeforeFlip;
{$ENDIF}
end;

Function TCopieColleDoc.IsTypeTOB : boolean;
begin
result := false;
{$IFDEF LAST}
  fmemo.Clear;
  fmemo.PasteFromClipboard;
  if copy(fmemo.Text,1,4)='$|0|' then Result := true;
  fmemo.clear;
{$ELSE}
  TOBimp := Tob.Create ( 'TOB RECUP', Nil, -1 ) ;
	TOBLoadFromBinFile (tempfile,RecupTob,nil);
  if TOBimp.detail.count > 0 then
  BEGIN
  	result := true;
  END else
  begin
  	freeAndNil (TOBimp);
    SysUtils.DeleteFile(Tempfile);
  end;
{$ENDIF}
end;

procedure TCopieColleDoc.BeforeFlip(Sender: TObject; ARow: Integer; var Cancel: Boolean);
var      TOBL : TOB;
begin
TOBL:=GetTOBLigne(TOBPiece,Arow) ; if TOBL=Nil then BEGIN cancel := true;Exit; END;
// VARIANTE
(* if (copy(TOBL.getValue('GL_TYPELIGNE'),1,2)='TP') then BEGIN cancel := true; Exit; END; *)
if IsFinParagraphe(TOBL) then BEGIN cancel := true; Exit; END;
if not ISLigneGerableCC (TOBPiece,TOBArticles,GereLot,GereSerie,Arow) then BEGIN Cancel := true; exit; END;
if IsLigneDetail (nil,TOBL) then BEGIN cancel := true; Exit; END;
if IsLigneDetailMode (TOBL) then BEGIN cancel := true; Exit; END;      // Rajout pour Mode
if not cancel then RowSelected := Arow;
end;

procedure TCopieColleDoc.FlipSelection(Sender: TObject);
var Arow,Niveau : integer;
    TOBL,TOBL1 : TOB;
    retrieve : boolean;
    RefSais : string;
begin
  if fgrid = nil then Exit;
  if Assigned (fgrid.OnFlipSelection) then
  begin
    Retrieve := true;
    fgrid.OnFlipSelection := nil;
    fgrid.OnBeforeFlip := nil;
  end else Retrieve := false;

  TOBL:=GetTOBLigne(TOBPiece,RowSelected) ; if TOBL=Nil then Exit ;
  // ARIANTE
  (* if (copy(TOBL.getValue('GL_TYPELIGNE'),1,2)='DP') then *)
  if ISDebutParagraphe (TOBL) then
  begin
    Niveau := TOBL.GetValue('GL_NIVEAUIMBRIC');
    for Arow := RowSelected+1 to fgrid.rowcount do
    begin
      TOBL1:=GetTOBLigne(TOBPiece,ARow) ; if TOBL1=Nil then break ;
      fgrid.FlipSelection (Arow);
      // VARIANTE
      (* if (copy(TOBL1.getValue('GL_TYPELIGNE'),1,2)='TP') and *)
      if IsFinParagraphe (TOBL1)  and (Niveau = TOBL1.GetValue('GL_NIVEAUIMBRIC')) then break;
    end;
    RowSelected := Arow;
  end;
  if (TOBL.GetValue ('GL_TYPEARTICLE')='OUV') and (not IsSousDetail (TOBL)) then
  begin
    Niveau := TOBL.GetValue('GL_INDICENOMEN');
    for Arow := RowSelected+1 to fgrid.rowcount do
    begin
      TOBL1:=GetTOBLigne(TOBPiece,ARow) ; if TOBL1=Nil then break ;
      // VARIANTE
      (* if (TOBL1.getValue ('GL_TYPELIGNE') = 'COM') and *)
      if (IsCommentaire(TOBL1) or (IsSousDetail(TOBL1))) and
      (TOBL1.GetVAlue('GL_INDICENOMEN')=Niveau) then fgrid.flipSelection(Arow)
                                                else break;
    end;
    RowSelected := Arow-1;
  end;
  // Rajout pour Mode
  // VARIANTE
  (* if (TOBL.GetValue ('GL_TYPELIGNE')='COM') and (TOBL.GetValue ('GL_TYPEDIM')='GEN') then*)
  if (IsCommentaire(TOBL) or IsSousDetail(TOBL)) and (TOBL.GetValue ('GL_TYPEDIM')='GEN') then
  begin
    RefSais := TOBL.GetValue ('GL_CODESDIM');
    for Arow := RowSelected+1 to fgrid.rowcount do
    begin
      TOBL1:=GetTOBLigne(TOBPiece,ARow) ; if TOBL1=Nil then break ;
      // VARIANTE
      (* if (TOBL1.getValue ('GL_TYPELIGNE') = 'ART') and *)
      if (ISArticle(TOBL1)) and
      (TOBL1.GetVAlue('GL_TYPEDIM')='DIM') and (TOBL1.GetVAlue('GL_CODEARTICLE')=RefSais) then
      BEGIN
        fgrid.flipSelection(Arow);
      END else break;
    end;
    RowSelected := Arow-1;
  end;

  if Retrieve then
  begin
    fgrid.OnFlipSelection := FlipSelection;
    fgrid.OnBeforeFlip := BeforeFlip;
  end;
  //if fgrid.nbselected = 0 then RowSelected := 0;
end;

procedure TCopieColleDoc.Couper;
begin
  Clipboard.clear;
  if fgrid = nil then Exit;
  if (fgrid.nbSelected = 0) and (TFFacture(fform).Action <> taConsult) then
  begin
     SysUtils.DeleteFile (Tempfile);
     SendMessage (fgrid.InplaceEditor.Handle,WM_CUT,0,0);
     exit;
  end;
  {$IFDEF BTP}
  CouperDonnee;
  {$ENDIF}
end;

procedure TCopieColleDoc.Copier;
begin
  Clipboard.clear;
  if fgrid = nil then Exit;
  if (fgrid.nbSelected = 0) and (TFFacture(fform).Action <> taConsult) then
  begin
     SysUtils.DeleteFile (Tempfile);
     SendMessage (fgrid.InplaceEditor.Handle,WM_COPY,0,0);
     exit;
  end;
  {$IFDEF BTP}
  CopieDonnee;
  {$ENDIF}
end;

procedure TCopieColleDoc.CopieDonnee;
var Indice : Integer;
    TOBSEL: TOB;
begin
  if fgrid.nbselected = 0 then exit;
  fgrid.OnFlipSelection := nil;
  fgrid.OnBeforeFlip := nil;
  TOBSEL := TOB.create ('DONNEAINS',nil,-1);
  TRY
    TOBSEL.AddChampSupValeur ('TYPE',ftypeInfo);
    for Indice := 1 to fgrid.RowCount do
    begin
    if (fgrid.IsSelected (Indice) ) then
       begin
       ajouteLigneSel (TOBSel,Indice,true);
       end;
    end;
    fform.Enabled := false;
    ExportTob(TOBSel,'GL_BLOCNOTE',True);
    fform.Enabled := true;
    Clipboard.clear;
    fMemo.Lines :=Mcopy;
    fMemo.SelectAll;
    fMemo.CopyToClipboard;
    Mcopy.Clear ;
  FINALLY
    TOBSel.free;
    fgrid.OnFlipSelection := FlipSelection ;
    fgrid.OnBeforeFlip := BeforeFlip;
  END;
end;


procedure TCopieColleDoc.CouperDonnee;
var Indice : Integer;
    TOBSEL: TOB;
    TOBTOCUT,TOBTC : TOB;
begin
  if fgrid.nbselected = 0 then exit;
  fgrid.OnFlipSelection := nil;
  fgrid.OnBeforeFlip := nil;
  TOBTOCUT := TOB.Create ('LES LIGNES A COUPER',nil,-1);
  TOBSEL := TOB.create ('DONNEAINS',nil,-1);
  TRY
    TOBSEL.AddChampSupValeur ('TYPE',ftypeInfo);
    for Indice := 1 to fgrid.RowCount do
    begin
      if (fgrid.IsSelected (Indice) ) then
      begin
      	if TFFacture(FF).ControleLigneToCut(Indice) then
        begin
        	TOBTC := TOB.Create ('UNE LIGNE',TOBTOCUT,-1);
          TOBTC.AddChampSupValeur('LA LIGNE',Indice);
        	ajouteLigneSel (TOBSel,Indice,true);
        end;
      end;
    end;
    fform.Enabled := false;
    ExportTob(TOBSel,'GL_BLOCNOTE',True);
    fform.Enabled := true;
    Clipboard.clear;
    fMemo.Lines :=Mcopy;
    fMemo.SelectAll;
    fMemo.CopyToClipboard;
    Mcopy.Clear ;
	FINALLY
    TOBSel.free;
  	TOBTOCUT.free;
  	deselectionneRows;
    fgrid.OnFlipSelection := FlipSelection ;
    fgrid.OnBeforeFlip := BeforeFlip;
	END;
end;

procedure TCopieColleDoc.Enregistrer;
begin
if fgrid = nil then Exit;
if fgrid.nbSelected = 0 then exit;
RowCollage := -1;
ColCollage := -1;
EnregistreDonne;
end;

procedure TCopieColleDoc.EnregistreDonne;
var CodeEnreg,LibEnreg : string;
    Indice : Integer;
    TOBSEL,TOBENREG : TOB;
begin
TOBEnreg := nil;
TOBSel := nil;
if fgrid.nbSelected = 0 then exit;
if FtypeInfo='' then BEGIN (* prout *) exit; END;
fgrid.OnFlipSelection := nil;
fgrid.OnBeforeFlip := nil;
TRY
if DemandeNomEnreg (fTypeInfo,CodeEnreg,LibEnreg) then
   begin
   TOBSEL := TOB.create ('DONNEAINS',nil,-1);
   TOBSEL.AddChampSupValeur ('TYPE',FtypeInfo);
   for Indice := 1 to fgrid.RowCount do
       begin
       if (fgrid.IsSelected (Indice) ) then ajouteLigneSel (TOBSel,Indice);
       end;
   ExportTob(TOBSel,'GL_BLOCNOTE');
   fMemo.Lines :=Mcopy;
   fMemo.SelectAll;
   //
   TOBENREG := TOB.create ('BMEMORISATION',nil,-1);
   TOBENreg.PutValue('BMO_TYPEMEMO',fTypeInfo);
   TOBENreg.PutValue('BMO_CODEMEMO',CodeEnreg);
   TOBENreg.PutValue('BMO_LIBMEMO',LibEnreg);
   TOBENREG.PutValue('BMO_MEMO',fmemo.Text);
   TOBENreg.InsertOrUpdateDB (true);
   end;
FINALLY
if TOBEnreg <> nil then TOBEnreg.free;
if TOBSel <> nil then TOBSel.free;
fgrid.OnFlipSelection := FlipSelection;
fgrid.OnBeforeFlip := BeforeFlip;
END;
end;

procedure TCopieColleDoc.AjouteLigneSel (TOBSel : TOB; Indice: Integer; WithMemo : boolean=false);

  function GetDetailouvrage (TOBOUV : TOB; UniquebLO : integer) : TOB;
  var II : integer;
  begin
    result := nil;
    for II := 0 to TOBOUV.detail.count -1 do
    begin
      if TOBOUV.detail[II].getInteger('BLO_UNIQUEBLO') = UniquebLo then
      begin
        result := TOBOUV.detail[II];
        break;
      end;
    end;
  end;

var IndiceNomen,II: Integer;
    TOBLoc,TOBL,TOBOuv,TOBLIen,TOBOO,TOBP : TOB;
    TypeArticle   : string;
    OK_Metredoc   : Boolean;
    Ok_MetreExcel : Boolean;
    Ok_MetreBib   : Boolean;
    QteFact       : Double;
    Qte           : Double;
begin

  OK_Metredoc   := _GetParamSocSecur('SO_BTMETREDOC', False);
  Ok_MetreExcel := _GetParamSocSecur('SO_BTMETRESEXCEL', False);
  Ok_MetreBib   := _GetParamSocSecur('SO_BTMETREBIB', False);

  TOBL:=GetTOBLigne(TOBPiece,Indice) ;

  if TOBL = nil then exit;

  if IsSousDetail(TOBL) then
  begin
    TOBP := FindTOBPere (TOBL);
    if TOBP = nil then exit;
    // si la ligne pere est sélectionnée on ne prend pas le detail
    if fgrid.IsSelected (TOBP.GetIndex +1) then exit;
  end;

  TTObPiece := TOBPIece;

  // VARIANTE
  //if ((not IsCommentaire(TOBL)) and (not IsSousDetail(TOBL))) or (TOBL.GetValue ('GL_INDICENOMEN')=0) then
  if (not IsCommentaire(TOBL)) then
  begin
  TOBLOC:=NewTobLigne(TOBSEL,0);
  CopieChampsLigne ('LIGNE',TOBLOC,TOBL);
  CopieChampsSup (TOBLOC,TOBL);
    //
    IndiceNomen := TOBL.GetValue('GL_INDICENOMEN');
    if IndiceNomen <> 0 then
    begin
      TobOuv      := TOBOUvrage.detail[IndiceNomen -1];
      TOBLien     := TOB.create ('LIGNOMEN',TOBLOC,-1);
    end;
    //
    if IsSousDetail(TOBLOC) then
    begin
      // transforme un sous detail en ligne
      TOBLOC.SetString('GL_TYPELIGNE','ART');
      QteFact := TOBP.GetDouble('GL_QTEFACT');
      Qte     := TOBLOC.GetDouble('GL_QTEFACT');

      IF QteFact <> 0 then
        Qte := Arrondi(Qte/QteFact,V_PGI.okdecQ)
      else
        Qte := QteFact;
      //
      TOBLOC.SetDouble('GL_QTERESTE', Qte);
      TOBLOC.SetDouble('GL_QTESTOCK', Qte);
      //FV1 - 19/01/2018 - FS#2877 - SCHREIBER - Copier/coller Récupère la valeur de la quantité de ligne au lieu de 1
      TOBLOC.SetDouble('GL_QTEFACT', Qte);
      // --- GUINIER ---
      TOBLOC.SetDouble('GL_MTRESTE', TOBLOC.GetDouble('GL_MONTANTHTDEV'));
    end;
  TOBLOC.ClearDetail; // on enleve les lignes d'analytique eventuelle
  //
	if (not WithMemo) then TOBLoc.putValue('GL_BLOCNOTE','');
  TypeArticle := TOBL.GetValue ('GL_TYPEARTICLE');
  if (TypeArticle = 'OUV') or (TypeArticle = 'ARP') then
     begin
     if (TOBOUvrage.detail.count > 0) and (IndiceNomen > 0) then
        begin
        TobOuv := TOBOUvrage.detail[IndiceNomen -1];
        TOBLien := TOB.create ('',TOBLOC,-1);
        //
        if not IsSousDetail(TOBL) then
        begin
          CopieDonneeOuvrage (TOBLien,TOBOUV,TOBL, Ok_MetreDoc, Ok_MetreBib, Ok_MetreExcel);
        end else
        begin
          TOBOO := GetDetailouvrage (TOBOUV,TOBL.GetInteger('UNIQUEBLO'));
          if TOBOO <> nil then CopieDonneeOuvrage (TOBLien,TOBOO, TOBL, Ok_MetreDoc, Ok_MetreBib, Ok_MetreExcel);
        end;
     end;
    end;
  if (TypeArticle = 'NOM') then
     begin
     IndiceNomen := TOBL.GetValue('GL_INDICENOMEN');
     TobOuv := TOBnOMEN.detail[IndiceNomen -1];
     TOBLien := TOB.create ('LIGNOMEN',TOBLOC,-1);
      TOBLIEN.AddChampSup('UTILISE',False) ;
      TOBLIEN.PutValue('UTILISE','-') ;
     TOBLIen.dupliquer (TOBOuv,true,true);
     end;
    if IsSousDetail(TOBL) then
    begin
      TOBOO := GetDetailouvrage (TOBOUV, TOBL.GetInteger('UNIQUEBLO'));
      if TOBOO <> nil then
      begin
        TOBLOC.PutValue('UNIQUEBLO', TOBOO.GETVALUE('BLO_UNIQUEBLO'));
        //FS#2387 - SCHREIBER - Copier/Coller Métrés sur sous-détail au niveau de l'écran facture
        //Ne faire cela que si les metrés Excel sur document sont paramétrés
        if Ok_MetreBib then CopieTobVarDoc (TTObPiece, TOBLOC, TOBP, TobvarDoc);
        if OK_Metredoc then
        begin
          if Ok_MetreExcel then CopieTobMetres (TTObPiece, TOBLOC, TOBP, TobMetres);
        end;
      end;
      //copie d'une ligne de sous-détail du document
      //CopieTobVarDoc (TTObPiece, TOBLOC, TOBP, TobvarDoc);
      //CopieTobMetres (TTObPiece, TOBLOC, TOBP, TobMetres);
  end
  else
  begin
      //Ne faire cela que si les metrés Excel sur document sont paramétrés
      if Ok_MetreBib then CopieTobVarDoc (TTOBPiece, TOBLOC, TOBL, TobvarDoc);
      if OK_Metredoc then
      begin
        if Ok_MetreExcel then CopieTobMetres (TTOBPiece, TOBLOC, TOBL, TobMetres);
      end;
    end;
  end
  else
  begin
    TOBLOC:=NewTobLigne(TOBSEL,0);
    CopieChampsLigne ('LIGNE',TOBLOC,TOBL);
    CopieChampsSup (TOBLOC,TOBL);
    TOBLOC.ClearDetail; // on enleve les lignes d'analytique eventuelle
    //
    if (not WithMemo) then TOBLoc.putValue('GL_BLOCNOTE','');
  end;

end;    

function TCopieColleDoc.AjouteLigneImportDetailOuv (TOBOUV,TOBLigneDoc,TOBI : TOB; Indice : integer; var ArticleNonOk : boolean) : boolean;
var TOBOO,TOBART,TOBOR,TOBID : TOB;
    QQ : TQuery;
    NatureTravail : double;
    LibelleFou,Mess,Libelle : string;
    X : double;
    MemeDoc : boolean;
    II : integer;
begin
  MemeDoc := IsSamePiece(TOBI,TOBPiece);
  Libelle := TOBI.GetString('GL_LIBELLE');
  TOBOR := TOBOuv.detail[0];

  TOBART := TOBArticles.findFirst(['GA_ARTICLE'],[TOBI.GetValue('GL_ARTICLE')],true);
  if TOBART = nil then
  begin
    TOBART:=CreerTOBArt(TOBArticles) ;
    QQ := OpenSql ('SELECT A.*,AC.*,N.BNP_TYPERESSOURCE,N.BNP_LIBELLE FROM ARTICLE A '+
                  'LEFT JOIN NATUREPREST N ON N.BNP_NATUREPRES=A.GA_NATUREPRES '+
                  'LEFT JOIN ARTICLECOMPL AC ON AC.GA2_ARTICLE=A.GA_ARTICLE WHERE A.GA_ARTICLE="'+
                  TOBI.GetValue('GL_ARTICLE')+'"',true);
    if QQ.EOF then
    begin
      Mess := 'Le composant '+ TOBI.Getvalue('GL_CODEARTICLE')+' - '+TOBI.Getvalue('GL_LIBELLE')+
              ' ne peut être récupéré #13#10 car il est inexistant.';
      PgiBox (TraduireMemoire(Mess),Traduirememoire('Copier-coller de lignes'));
      ArticleNonOk := true;
      Result := False;
      Exit;
    end;
    //
    TOBART.SelectDB('',QQ) ;
    Ferme (QQ);
    InitChampsSupArticle (TOBART);
  end;

  NatureTravail := valeur(TOBLigneDoc.getValue('GLC_NATURETRAVAIL'));
  LibelleFou := TOBLigneDoc.getValue('LIBELLEFOU');
  //
  TOBOO := TOB.Create ('LIGNEOUV',TOBOUV,Indice);
  InsertionChampSupOuv (TOBOO,false);
  TOBPiece.putValue('GP_UNIQUEBLO',TOBPiece.getValue('GP_UNIQUEBLO')+1);
  TOBOO.putValue('BLO_UNIQUEBLO',TOBPiece.getValue('GP_UNIQUEBLO'));
  if (natureTravail > 0) and (NatureTravail <= 10) then
  begin
  	TOBOO.putValue('BLO_NATURETRAVAIL',TOBLigneDoc.GetValue('GLC_NATURETRAVAIL'));
  	TOBOO.putValue('BLO_FOURNISSEUR',TOBLigneDoc.GetValue('GL_FOURNISSEUR'));
    if TOBOO.FieldExists('BLF_NATURETRAVAIL') then
    begin
      TOBOO.putvalue('BLF_NATURETRAVAIL',TOBLigneDoc.GetValue('GLC_NATURETRAVAIL'));
      TOBOO.putvalue('BLF_FOURNISSEUR',TOBLigneDoc.GetValue('GL_FOURNISSEUR'));
    end;
  end;
  TOBOO.putvalue('LIBELLEFOU',LibelleFou);
  //
  TOBOO.SetDouble ('BLO_QTESAIS',TOBI.GetDouble('GL_QTESAIS'));
  TOBOO.SetDouble ('BLO_RENDEMENT',TOBI.GetDouble('GL_RENDEMENT'));
  TOBOO.SetDouble ('BLO_PERTE',TOBI.GetDouble('GL_PERTE'));
  TOBOO.SetSTring ('BLO_QUALIFHEURE',TOBI.getString('GL_QUALIFHEURE'));
  //
  TOBOO.Putvalue ('BLO_QTEDUDETAIL',TOBOR.getInteger('BLO_QTEDUDETAIL'));
  TOBOO.PutValue('BLO_NATUREPIECEG',TOBLigneDoc.getValue('GL_NATUREPIECEG'));
  TOBOO.PutValue('BLO_SOUCHE',TOBLigneDoc.getValue('GL_SOUCHE'));
  TOBOO.PutValue('BLO_NUMERO',TOBLigneDoc.getValue('GL_NUMERO'));
  TOBOO.PutValue('BLO_INDICEG',TOBLigneDoc.getValue('GL_INDICEG'));
  TOBOO.PutValue('BLO_DOMAINE',TOBLigneDoc.getValue('GL_DOMAINE'));
  TOBOO.PutValue('BLO_CODEARTICLE', TOBArt.getValue('GA_CODEARTICLE'));
  TOBOO.PutValue('BLO_ARTICLE', TOBArt.getValue('GA_ARTICLE'));
  TOBOO.PutValue('BLO_ORDRECOMPO',1);
  RenseigneTOBOuv(TOBpiece,TOBOO,TOBLigneDoc,TOBART,Libelle,DEV);
  //
  //FV1 : 22/01/2018 - FS#2886 - TREUIL - Unité de facturation non conservée suite à copie d'une ligne dans sous-détail d'ouvrage
  TOBOO.PutValue('BLO_QUALIFQTEVTE',TOBI.GetValue('GL_QUALIFQTEVTE'));
  TOBOO.PutValue('BLO_QUALIFQTESTO',TOBI.GetValue('GL_QUALIFQTESTO'));
  //
  TOBOO.putValue('BLO_REFARTSAISIE',TOBLigneDoc.getValue('GL_REFARTSAISIE'));
  TOBOO.PutValue('BLO_COMPOSE',TOBOR.getValue('BLO_COMPOSE')) ;
  TOBOO.Putvalue('BLO_NOMENCLATURE',TOBOR.getValue('BLO_NOMENCLATURE')) ;
  TOBOO.Putvalue('BLO_QTEFACT',TOBI.getValue('GL_QTEFACT')) ;
  TOBOO.Putvalue('BLO_INDICENOMEN',TOBOR.getInteger('BLO_QTEDUDETAIL')) ;
  TOBOO.PutValue('BLO_DPA',TOBI.getValue('GL_DPA'));
  TOBOO.PutValue('BLO_COEFFG',TOBI.getValue('GL_COEFFG'));
  TOBOO.PutValue('BLO_COEFFC',TOBI.getValue('GL_COEFFC'));
  TOBOO.PutValue('BLO_COEFFR',TOBI.getValue('GL_COEFFR'));
  TOBOO.PutValue('BLO_COEFMARG',TOBI.getValue('GL_COEFMARG'));
  TOBOO.PutValue('BLO_DPR',TOBI.getValue('GL_DPR'));
  TOBOO.PutValue('BLO_PUHTDEV',TOBI.getValue('GL_PUHTDEV'));
  TOBOO.PutValue('BLO_PUHT',TOBI.getValue('GL_PUHT'));
  TOBOO.PutValue('BLO_PUHT',TOBI.getValue('GL_PUHT'));
  //FV1 - 25/01/2018 : FS#2892 - TREUIL - Le copier-coller ne récupère pas le coef. de marge
  //TOBOO.PutValue('BLO_COEFMARG',    0);
  //
  if not MemeDoc then
  begin
    //
    TOBOO.PutValue('BLO_COEFFC',0);
    TOBOO.PutValue('BLO_COEFFR',0);
    TOBOO.PutValue('BLO_MONTANTFC',0);
    TOBOO.PutValue('BLO_MONTANTFR',0);
    //
    if (TOBLigneDoc.getValue('GL_DOMAINE')<>'') then
    begin
      TOBOO.putValue('BLO_DOMAINE',TOBLigneDoc.getValue('GL_DOMAINE'));
    end;
    if TOBOO.getValue('BLO_DOMAINE') <> '' then
    begin
      if (TOBART.GetValue('GA_PRIXPASMODIF')<>'X') and (TOBLigneDoc.GetValue('GLC_FROMBORDEREAU')<>'X') then
      begin
        //FV1 : 12/09/2013 - FS#33 - POUCHAIN : En facturation de consos, les coefficients de l'UO ne sont pas repris
        //AppliqueCoefDomaineActOuv (TOBD);
        AppliqueCoefDomaineActOuv (TOBOO, copy(TOBOO.GetValue('BLO_AFFAIRE'),0,1));
        if VH_GC.BTCODESPECIF = '001' then ApliqueCoeflignePOC (TOBOO,DEV);
      end;
      CalculeLigneAcOuv (TOBOO,DEV,true,TOBART);
    end;
    //
    if TOBPiece.GetValue('GP_FACTUREHT')='X' then X := TOBLigneDoc.GetValue('GL_PUHTDEV')
                                             else X := TOBLigneDoc.GetValue('GL_PUTTCDEV');
    // Retablissement du pu dans la monnaie du client
    if TOBLigneDoc.GetValue('GL_DEVISE') <> TOBPiece.getValue('GP_DEVISE') then
    begin
      X:= DEVISETODEVISE (X,TOBLigneDoc.GetValue('GL_TAUXDEV'),TOBLigneDoc.GetValue('GL_COTATION'),
                          TOBPiece.GetValue('GP_TAUXDEV'),TOBPiece.GetValue('GP_COTATION'),V_PGI.okdecV);
    end else
    BEGIN
      if TOBPiece.GetValue('GP_FACTUREHT')='X' then X := TOBLigneDoc.GEtValue('GL_PUHTDEV')
                                               else X := TOBLigneDoc.GetValue('GL_PUTTCDEV');
    END;

    if TOBPiece.GetValue('GP_FACTUREHT')='X' then TOBOO.PutValue('BLO_PUHTDEV',X)
                                             else TOBOO.PutValue('BLO_PUTTCDEV',X);
    TOBOO.PutValue('GL_PUHT',DeviseToPIvotEx(TOBOO.getValue('BLO_PUHTDEV'),DEV.taux,DEV.quotite,V_PGI.okdecP));
    if TOBOO.getValue('BLO_DPR')<>0 then TOBOO.SetDouble('BLO_COEFMARG',Arrondi(TOBOO.getValue('BLO_PUHT')/TOBOO.getValue('BLO_DPR'),4))
                                    else TOBOO.SetDouble('BLO_COEFMARG',0);

  end;
  TOBOO.PutValue('ANCPA',TOBOO.GetValue('BLO_DPA'));
  TOBOO.PutValue('ANCPR',TOBOO.GetValue('BLO_DPR'));
  GetValoDetail (TOBOO); // pour le cas des Article en prix posés
  TOBOO.PutValue('GA_PVHT',TOBOO.GetValue('BLO_PUHTDEV'));
  if TOBLigneDoc.GetValue('GL_FACTUREHT')='X' then TOBOO.putvalue ('ANCPV',TOBOO.Getvalue ('BLO_PUHTDEV'))
                                              else TOBOO.putvalue ('ANCPV',TOBOO.Getvalue ('BLO_PUTTCDEV'));
  ReinitLigneFac (TOBOO);

  if TheMetreDoc <> nil then
  begin
    If TheMetredoc.OkMetreBiblio then ColleTobVarDoc(TOBOO, TOBI, TobVarDoc, TOBLigneDoc);
    if TheMetredoc.OkMetreDoc then
    begin
      if TheMetredoc.OkExcel then ColleTobMetres(TOBOO, TOBI, TOBLigneDoc, TheMetredoc);
    end;
  end;
  //
  if TOBI.detail.count > 0 then
  begin
    TOBID := TOBI.detail[0];
    II := 0;
    repeat
      ArticleNonOk := Not TraiteLigneCopierCollerOuv (TOBID.detail[0],TOBPiece,TOBLigneDoc,TOBArticles,Dev,MemeDoc);
      //
      TOBID.detail[II].ChangeParent(TOBOO,-1);
    until TOBID.detail.count = 0;
  end;
  
end;

function TCopieColleDoc.AjouteLigneImportOuv (TOBLigneDoc,TOBL,TOBAINS: TOB; var ArticleNonOk : boolean) : boolean;
var Indice,IndiceNomen,II : integer;
    TOBOUV : TOB;
begin
  IndiceNomen := TOBLigneDoc.GetInteger('GL_INDICENOMEN') ;
  if IndiceNomen = 0 then exit;
  TOBOUV := TOBOuvrage.detail[IndiceNomen-1];
  if TOBOUV = nil then exit;
  // recherche de l'element ou l'on va inserer les lignes
  for Indice := 0 to TOBOUV.detail.count -1 do
  begin
    if TOBOUV.detail[Indice].GetInteger('BLO_UNIQUEBLO')=TOBL.GetInteger('UNIQUEBLO') then break;
  end;
  if Indice >= TOBOUV.detail.count then indice := -1;
  for II := 0 to TOBIMP.detail.count -1 do
  begin
    AjouteLigneImportDetailOuv (TOBOUV,TOBLigneDoc,TOBAINS.detail[II],Indice,ArticleNonOk);
    if Indice <> -1 then inc(Indice);
  end;
  NumeroteLigneOuv (TOBOUV,TOBLigneDoc,1,1,0,0,0);
end;

function TCopieColleDoc.AjouteLigneImport (TOBLoc : TOB;Var Niveau : integer;var Arow : integer; Var SauteAFinPar : boolean; Var ArticleNonOk : boolean) : boolean;
begin
  Result := false;
  // VARIANTE
  (* if copy(TOBLOc.getValue('GL_TYPELIGNE'),1,2)='DP' then *)

  if IsDebutParagraphe(TOBLoc) then
  begin
    // Debut de paragraphe
    if Niveau = 9 then BEGIN SauteAfinPar := true; ArticleNonOk := true; END;
    if not SauteAFinPar then
    begin
      inc (niveau);
      InsertDebParagrapheImp (TOBLOC,niveau,Arow); inc(Arow);
      result := true;
    end;
    //  VARIANTE
    (* end else if copy(TOBLOC.getValue('GL_TYPELIGNE'),1,2)='TP' then *)
  end else if IsFinParagraphe(TOBLoc) then
  begin
    if not SauteAFinPar then
    begin
      // Fin de paragraphe
      InsertFinParagrapheImp (TOBLOC,niveau,Arow); inc(Arow);
      Dec (niveau);
      result := true;
    end;
    if SauteAFinPar then SauteAFinPar := false;
  end else
  begin
    if not SauteAFinPar then
    begin
      // VARIANTE
      (* if TOBLoc.GetValue('GL_TYPELIGNE') = 'ART' then *)
      if IsArticle(TOBLoc) then
      begin
        result := InsertLigneArticleImp (TOBLOC,Arow,Niveau,ArticleNonOk);
        // VARIANTE
        (* end else if (TOBLoc.getValue('GL_TYPELIGNE') = 'COM') and *)
      end else if (IsCommentaire(TOBLoc)) and
                  (TOBLoc.GetValue('GL_TYPEARTICLE') <> 'EPO') then
      begin
        InsertLigneCommentaireImp (TOBLoc,Arow,Niveau);  inc(Arow);
        result := true;
      end else if (TOBLoc.getValue('GL_TYPELIGNE') = 'TOT') then
      begin
        InsertLigneTotalImp (TOBLoc,Arow,Niveau);  inc(Arow);
        result := true;
      end;
    end;
  end;
end;


procedure TCopieColleDoc.InsertDebParagrapheImp (TOBLOC : TOB; niveau : integer;Arow : integer);
var TOBL : TOB;
    InVariante : boolean;
begin
TOBL := GetTOBLigne (TobPiece,Arow);
InVariante := IsVariante (TOBL);

// Code insertion nouveau paragraphe
InsertTobLigne (TOBPiece,Arow);
TOBL := GetTOBLIgne (TOBPiece,Arow);
AddLesSupLigne(TOBL,False) ;  InitLesSupLigne (TOBL);

PieceVersLigne (TOBPIece,TOBL);
AffaireVersLigne (TOBPiece,TOBL,TOBAffaire);
TOBL.PutValue('GL_LIBELLE',TOBLOC.GetValue('GL_LIBELLE'));
TOBL.PutValue('GL_NIVEAUIMBRIC',Niveau);
TOBL.PutValue('GL_PIECEPRECEDENTE','');
TOBL.PutValue('GL_PIECEORIGINE','');
TOBL.PutValue('GLC_DOCUMENTLIE','');
TOBL.PutValue('GLC_NUMEROTATION','');
TOBL.SetBoolean('GLC_NUMFORCED',false);
// VARIANTE
(* TOBL.PutValue('GL_TYPELIGNE','DP'+ intToStr(Niveau)); *)
if VariantePieceAutorisee (TOBPiece) then
begin
  if InVariante then TOBL.PutValue('GL_TYPELIGNE','DV'+ intToStr(Niveau))
                else TOBL.PutValue('GL_TYPELIGNE',copy(TOBLOc.GetValue('GL_TYPELIGNE'),1,2)+ intToStr(Niveau));
end else TOBL.PutValue('GL_TYPELIGNE','DP'+ intToStr(Niveau));
TOBL.PutValue('GL_NUMORDRE',GetMaxNumOrdre(TOBpiece));
end;

procedure TCopieColleDoc.InsertFinParagrapheImp (TOBLOC : TOB;niveau,Arow : integer);
var TOBL : TOB;
    InVariante : boolean;
begin
TOBL := GetTOBLigne (TobPiece,Arow);
InVariante := IsVariante (TOBL);
// Code insertion fin de paragraphe
InsertTobLigne (TOBPiece,Arow);
TOBL := GetTOBLIgne (TOBPiece,Arow);
AddLesSupLigne(TOBL,False) ;  InitLesSupLigne (TOBL);
PieceVersLigne (TOBPIece,TOBL);
AffaireVersLigne (TOBPiece,TOBL,TOBAffaire);
TOBL.PutValue('GL_LIBELLE',TOBLOC.GetValue('GL_LIBELLE'));
TOBL.PutValue('GL_NIVEAUIMBRIC',Niveau);
TOBL.PutValue('GL_PIECEPRECEDENTE','');
TOBL.PutValue('GL_PIECEORIGINE','');
TOBL.PutValue('GLC_DOCUMENTLIE','');
TOBL.PutValue('GLC_NUMEROTATION','');
TOBL.SetBoolean('GLC_NUMFORCED',false);
// VARIANTE
(* TOBL.PutValue('GL_TYPELIGNE','TP'+ intToStr(Niveau)); *)
if VariantePieceAutorisee (TOBPiece) then
begin
  if InVariante then TOBL.PutValue('GL_TYPELIGNE','TV'+ intToStr(Niveau))
                else TOBL.PutValue('GL_TYPELIGNE',copy(TOBLOc.GetValue('GL_TYPELIGNE'),1,2)+ intToStr(Niveau));
end else TOBL.PutValue('GL_TYPELIGNE','TP'+ intToStr(Niveau));
TOBL.PutValue('GL_NUMORDRE',GetMaxNumOrdre(TOBpiece));
end;

procedure TCopieColleDoc.VerifChampsSupsDetail(TOBOuv : TOB);
var Indice : integer;
begin
  if TOBOuv.fieldExists ('__COEFMARG') then
    if VarType (TOBOuv.GetValue('__COEFMARG')) = VarString then
      TOBOuv.PutValue ('__COEFMARG',Valeur(TOBOuv.GetValue('__COEFMARG')));
  if TOBOuv.fieldExists ('GCA_PRIXBASE') then
    if VarType (TOBOUV.GetValue('GCA_PRIXBASE')) = VarString then
      TOBOuv.PutValue ('GCA_PRIXBASE',Valeur(TOBOUV.GetValue('GCA_PRIXBASE')));
  if TOBOuv.fieldExists ('GA_PAHT') then
    if VarType (TOBOUV.GetValue('GA_PAHT')) = VarString then
      TOBOuv.PutValue ('GA_PAHT',Valeur(TOBOUV.GetValue('GA_PAHT')));
  if TOBOuv.fieldExists ('GA_PVHT') then
    if VarType (TOBOUV.GetValue('GA_PVHT')) = VarString then
      TOBOuv.PutValue ('GA_PVHT',Valeur(TOBOUV.GetValue('GA_PVHT')));
    (* OPTIMIZATION *)
  if TOBOuv.fieldExists ('BLO_PRXACHBASE') then
    if VarType (TOBOUV.GetValue('BLO_PRXACHBASE')) = VarString then
      TOBOuv.PutValue ('BLO_PRXACHBASE',Valeur(TOBOUV.GetValue('BLO_PRXACHBASE')));
  if TOBOuv.fieldExists ('GA_HEURE') then
    if VarType (TOBOUV.GetValue('GA_HEURE')) = VarString then
      TOBOuv.PutValue ('GA_HEURE',Valeur(TOBOUV.GetValue('GA_HEURE')));
  if TOBOuv.fieldExists ('GF_PRIXUNITAIRE') then
    if VarType (TOBOUV.GetValue('GF_PRIXUNITAIRE')) = VarString then
      TOBOuv.PutValue ('GF_PRIXUNITAIRE',Valeur(TOBOUV.GetValue('GF_PRIXUNITAIRE')));
  if TOBOuv.fieldExists ('ANCPA') then
    if VarType (TOBOUV.GetValue('ANCPA')) = VarString then
      TOBOuv.PutValue ('ANCPA',Valeur(TOBOUV.GetValue('ANCPA')));
  if TOBOuv.fieldExists ('ANCPR') then
    if VarType (TOBOUV.GetValue('ANCPR')) = VarString then
      TOBOuv.PutValue ('ANCPR',Valeur(TOBOUV.GetValue('ANCPR')));
(*
  if TOBOuv.fieldExists ('TPSUNITAIRE') then
    if VarType (TOBOUV.GetValue('TPSUNITAIRE')) = VarString then
      TOBOuv.PutValue ('TPSUNITAIRE',Valeur(TOBOUV.GetValue('TPSUNITAIRE')));
*)
  if TOBOuv.fieldExists ('INDICEMETRE') then
    if VarType (TOBOUV.GetValue('INDICEMETRE')) = VarString then
      TOBOuv.PutValue ('INDICEMETRE',Valeur(TOBOUV.GetValue('INDICEMETRE')));
  if TOBOuv.fieldExists ('INDICELIENDEVCHA') then
    if VarType (TOBOUV.GetValue('INDICELIENDEVCHA')) = VarString then
      TOBOuv.PutValue ('INDICELIENDEVCHA',Valeur(TOBOUV.GetValue('INDICELIENDEVCHA')));
  if TOBOUV.detail.count > 0 then
  begin
  	For Indice := 0 to TOBOUV.detail.count -1 do
    begin
    	VerifChampsSupsDetail (TOBOUV.detail[Indice]);
    end;
  end;
end;

Function TCopieColleDoc.TraiteLigneCopierCollerOuv (TOBD,TOBpiece,TOBL,TOBArticle: TOB;DEV : RDevise; memedoc : boolean) : boolean;
var Indice : integer;
  		TOBA : TOB;
      QQ : TQuery;
      Mess : String;
begin
    Result := True;

    TOBA := TOBArticles.findFirst(['GA_ARTICLE'],[TOBD.GetValue('BLO_ARTICLE')],true);
    if TOBA = nil then
    begin
      TOBA:=CreerTOBArt(TOBArticles) ;
      QQ := OpenSql ('SELECT A.*,AC.*,N.BNP_TYPERESSOURCE,N.BNP_LIBELLE FROM ARTICLE A '+
                    'LEFT JOIN NATUREPREST N ON N.BNP_NATUREPRES=A.GA_NATUREPRES '+
                    'LEFT JOIN ARTICLECOMPL AC ON AC.GA2_ARTICLE=A.GA_ARTICLE WHERE A.GA_ARTICLE="'+
                    TOBD.GetValue('BLO_ARTICLE')+'"',true);
      if QQ.EOF then
      begin
        Mess := 'Le composant '+ TOBD.Getvalue('BLO_CODEARTICLE')+' - '+TOBD.Getvalue('BLO_LIBELLE')+
                ' dans l''ouvrage ' + TOBL.Getvalue('GL_CODEARTICLE')+' - '+TOBL.Getvalue('GL_LIBELLE') +
                ' ne peut être récupéré #13#10 car il est inexistant.';
        PgiBox (TraduireMemoire(Mess),Traduirememoire('Copier-coller de lignes'));
        TOBD.free;
        Result := False;
        Exit;
      end;

      TOBA.SelectDB('',QQ) ;
      Ferme (QQ);
      InitChampsSupArticle (TOBA);
    end;
    //
    if not memedoc then
    begin
      if (TOBL.getValue('GL_DOMAINE')<>'') then
      begin
        TOBD.putValue('BLO_DOMAINE',TOBL.getValue('GL_DOMAINE'));
      end;
      if TOBD.getValue('BLO_DOMAINE') <> '' then
      begin

        if (TOBA.GetValue('GA_PRIXPASMODIF')<>'X') and (TOBL.GetValue('GLC_FROMBORDEREAU')<>'X') then
        begin
          //FV1 : 12/09/2013 - FS#33 - POUCHAIN : En facturation de consos, les coefficients de l'UO ne sont pas repris
          //AppliqueCoefDomaineActOuv (TOBD);
          AppliqueCoefDomaineActOuv (TOBD, copy(TOBD.GetValue('BLO_AFFAIRE'),0,1));
        if VH_GC.BTCODESPECIF = '001' then ApliqueCoeflignePOC (TOBD,DEV);
        end;
    end else
    begin
      if (TOBA.GetValue('GA_PRIXPASMODIF')<>'X') and (TOBL.GetValue('GLC_FROMBORDEREAU')<>'X') then
      begin
        if VH_GC.BTCODESPECIF = '001' then ApliqueCoeflignePOC (TOBD,DEV);
      end;
    end;
  end;
  // iNCRMETATION DU NUMERO UNIQUE DE LIGNE OUVRAGE
  TOBPiece.putValue('GP_UNIQUEBLO',TOBPiece.getValue('GP_UNIQUEBLO')+1);
  TOBD.putValue('BLO_UNIQUEBLO',   TOBPiece.getValue('GP_UNIQUEBLO'));
  //
  if TheMetreDoc <> nil then
  begin
    If TheMetredoc.OkMetreBiblio then ColleTobVarDoc(TOBD, TOBD, TobVarDoc, TOBL);
    if TheMetredoc.OkMetreDoc then
    begin
      if TheMetredoc.OkExcel then ColleTobMetres(TOBD, TOBD, TOBL, TheMetredoc);
    end;
  end;
  //
  TOBD.DelChampSup('VARMETRELIGNE', False);
  TOBD.DelChampSup('VARFILELIGNE', False);
  //
		ReinitLigneFac (TOBD);
    if TOBD.detail.count > 0 then
    begin
    	Indice := 0;
      while Indice < TOBD.detail.count do
      begin
        result := TraiteLigneCopierCollerOuv (TOBD.Detail[Indice],Tobpiece,TOBL,TOBArticles,Dev,MemeDoc);
        Inc(Indice);
      end;
    end;
    //

    VerifChampsSupsDetail(TOBD);
    InitCopierCollerDetail (TOBA,TOBD,TOBpiece);

    if not MemeDoc then
    begin
      if (TOBD.getValue('BLO_DOMAINE')<>'') then
      begin
        CalculeLigneAcOuv (TOBD,DEV,true,TOBA);
      end;

      ReajusteLigneParDoc (TOBL,TOBD,DEV,1);
    end;
      GetValoDetail (TOBD); // pour le cas des Article en prix posés
      TOBD.PutValue('GA_PVHT',TOBD.GetValue('BLO_PUHTDEV'));
      if TOBL.GetValue('GL_FACTUREHT')='X' then TOBD.putvalue ('ANCPV',TOBD.Getvalue ('BLO_PUHTDEV'))
                                           else TOBD.putvalue ('ANCPV',TOBD.Getvalue ('BLO_PUTTCDEV'));
end;


function TCopieColleDoc.InsertLigneArticleImp (TOBLOC : TOB;var Arow : integer ;Niveau : integer;var ArticleNonOk:boolean):boolean;
var TOBL,TOBA,TOBOuv : TOB;
    X : double;
    DetailATraiter : boolean;
    TOBLN : TOB;
{$IFDEF BTP}
    Indice : integer;
    valeurs : T_Valeurs;
    PuFix,Qte : double;
{$ENDIF}
    InVariante : boolean;
    QQ : TQuery;
    Mess : String;
    MemeDoc : Boolean;
    QteFact : double;
begin
  MemeDoc := IsSamePiece(TOBLOC,TOBPiece);
  result := false;
  TOBOuv := nil;
  DetailATraiter := false;
  TOBL := GetTOBLigne (TobPiece,Arow);
  InVariante := IsVariante (TOBL);
  // Code Insertion ligne de document
  InsertTobLigne (TOBPiece,Arow);
  TOBL := GetTOBLIgne (TOBPiece,Arow);
  if (TOBLoc.detail.count > 0) and ((TOBLoc.GetValue('GL_TYPEARTICLE')='OUV') or (TOBLoc.GetValue('GL_TYPEARTICLE')='ARP')) then
  begin
    // gestion du detail de l'ouvrage
    TOBOuv := TOB.create ('',nil,-1);
    TOBOUV.dupliquer (TOBLoc.detail[0],true,true);
    TobLoc.ClearDetail;
  end;
  AddLesSupLigne(TOBL,False) ;  InitLesSupLigne (TOBL);
  CopieChampsLigne ('LIGNE',TOBL,TOBLOC);
  CopieChampsSup (TOBL,TOBLOC);
  VerifChampsSupsLigne(TOBL);
  TOBA := TOBArticles.findFirst(['GA_ARTICLE'],[TOBL.GetValue('GL_ARTICLE')],true);
  if TOBA = nil then
  begin
    TOBA:=CreerTOBArt(TOBArticles) ;
    QQ := OpenSql ('SELECT A.*,AC.*,N.BNP_TYPERESSOURCE,N.BNP_LIBELLE FROM ARTICLE A '+
                  'LEFT JOIN NATUREPREST N ON N.BNP_NATUREPRES=A.GA_NATUREPRES '+
                  'LEFT JOIN ARTICLECOMPL AC ON AC.GA2_ARTICLE=A.GA_ARTICLE WHERE A.GA_ARTICLE="'+
                  TOBL.GetValue('GL_ARTICLE')+'"',true);
    if QQ.EOF then
    begin
      Mess := 'Le composant '+ TOBL.Getvalue('GL_CODEARTICLE')+' - '+TOBL.Getvalue('GL_LIBELLE')+
              ' ne peut être récupéré #13#10 car il est inexistant.';
      PgiBox (TraduireMemoire(Mess),Traduirememoire('Copier-coller de lignes'));
      TOBL.free;
      Result := False;
      Exit;
    end;

    TOBA.SelectDB('',QQ) ;
    Ferme (QQ);
    InitChampsSupArticle (TOBA);
  end;

  //
  QteFact := TOBL.GetDouble('GL_QTEFACT');
  //
  TOBL.PutValue('BLP_NUMMOUV',0);
  TOBL.PutValue('GL_QTESTOCK', QteFact);
  TOBL.PutValue('GL_QTERESTE', QteFact);
  // --- GUINIER ---
  TOBL.PutValue('GL_MTRESTE', TOBL.GetValue('GL_MONTANTHTDEV'));
  TOBL.PutValue('GL_PIECEPRECEDENTE','');
  TOBL.PutValue('GL_PIECEORIGINE','');
  TOBL.PutValue('GLC_DOCUMENTLIE','');
  TOBL.PutValue('GLC_NUMEROTATION','');
  TOBL.SetBoolean('GLC_NUMFORCED',false);
  // Correction FQ 12072
  TOBL.PutValue('GL_COEFFC',0);
  TOBL.PutValue('GL_COEFFR',0);
  // --
  if TOBL.GetValue('GLC_NATURETRAVAIL') <> '' then
  begin
    TOBL.PutValue('GLC_NATURETRAVAIL','');
    TOBL.PutValue('GL_FOURNISSEUR',TOBA.GetValue('GA_FOURNPRINC'));
  end;
  //
  ReinitLigneFac (TOBL);
  //
  TOBL.PutValue('GL_BLOQUETARIF','-');
  if not MemeDoc then
  begin
    if VenteAchat = 'VEN' then
    begin
      if IsprestationST (TOBL) then
      begin
        if TOBPIece.getValue('GP_APPLICFGST')='-' then TOBL.putValue('GLC_NONAPPLICFRAIS','X') else TOBL.putValue('GLC_NONAPPLICFRAIS','-');
        if TOBPIece.getValue('GP_APPLICFCST')='-' then TOBL.putValue('GLC_NONAPPLICFC','X') else TOBL.putValue('GLC_NONAPPLICFC','-');
      end;
    end;
  end;
  //NewTOBLigneFille (TOBL);
  PieceVersLigne (TOBPIece,TOBL);
  AffaireVersLigne (TOBPiece,TOBL,TOBAffaire);
  TOBL.PutValue('GL_NIVEAUIMBRIC',Niveau);
  if not MemeDoc then
  begin
    if TOBL.GetValue('GL_TYPEARTICLE') <> 'POU' then
    begin
      if TOBPiece.GetValue('GP_FACTUREHT')='X' then X := TOBL.GetValue('GL_PUHTDEV')
                                               else X := TOBL.GetValue('GL_PUTTCDEV');
      // Retablissement du pu dans la monnaie du client
      if TOBL.GetValue('GL_DEVISE') <> TOBPiece.getValue('GP_DEVISE') then
      begin
        X:= DEVISETODEVISE (X,TOBL.GetValue('GL_TAUXDEV'),TOBL.GetValue('GL_COTATION'),
        TOBPiece.GetValue('GP_TAUXDEV'),TOBPiece.GetValue('GP_COTATION'),V_PGI.okdecV);
      end else
      BEGIN
        if TOBPiece.GetValue('GP_FACTUREHT')='X' then X := TOBL.GEtValue('GL_PUHTDEV')
                                                 else X := TOBL.GetValue('GL_PUTTCDEV');
      END;

      if TOBPiece.GetValue('GP_FACTUREHT')='X' then TOBL.PutValue('GL_PUHTDEV',X)
                                               else TOBL.PutValue('GL_PUTTCDEV',X);
      TOBL.PutValue('GL_PUHT',DeviseToPIvotEx(TOBL.getValue('GL_PUHTDEV'),DEV.taux,DEV.quotite,V_PGI.okdecP));
      if TOBL.getValue('GL_DPR')<>0 then TOBL.PutValue('GL_COEFMARG',Arrondi(TOBL.getValue('GL_PUHT')/TOBL.getValue('GL_DPR'),4))
                                    else TOBL.SetDouble('GL_COEFMARG',0);
    end;
  end;

  if Not IsArticleOk (Arow) then
  begin
    ArticleNonOk := true;
    PgiBox (TraduireMemoire('L''article '+ TOBL.Getvalue('GL_CODEARTICLE')+' - '+TOBL.Getvalue('GL_LIBELLE')+'#13#10 ne peut être récupéré car inexistant.'),Traduirememoire('Copier-coller de lignes'));
    TOBL.free;
  end else
  begin
    TOBA := FindTOBArtRow (TOBPiece,TOBarticles,Arow);

(*
    TOBL.PutValue('GL_DOMAINE', TOBA.GetValue('GA_DOMAINE')); // application du domaine d'activite
    if TOBPiece.getValue('GP_DOMAINE')<>'' then TOBL.putValue('GL_DOMAINE',TOBPIece.getValue('GP_DOMAINE'));
    if (TOBL.getValue('GL_DOMAINE')<>'') then
    begin
      if (TOBA.GetValue('GA_PRIXPASMODIF')<>'X') and (TOBL.GetValue('GLC_FROMBORDEREAU')<>'X') then
      begin
        AppliqueCoefDomaineLig(TOBL);
      end;
    end;
*)

    if VariantePieceAutorisee (TOBPiece) then
    begin
      if InVariante then SetTypeLigne (TOBL,InVariante)
                    else TOBL.Putvalue('GL_TYPELIGNE',TOBLOc.getValue('GL_TYPELIGNE'));
    end;
    TFFacture(fform).TraiteLaCompta(ARow);
    TOBL.PutValue('GL_NIVEAUIMBRIC',Niveau);
    {$IFDEF BTP}
    if not MemeDoc then
    begin
      if TOBL.GetValue('GLC_FROMBORDEREAU')='X' then
      begin
        TOBL.PutValue('GL_BLOQUETARIF','X');
      end else
      begin
        if (TOBPiece.GetValue('_BLOQUETARIF')='X') and (TOBL.getValue('GL_DEVISE')=TOBPiece.getValue('GP_DEVISE')) Then
        begin
          TOBL.PutValue('GL_BLOQUETARIF','X')
        end else
        begin
          TOBL.PutValue('GL_BLOQUETARIF','-')
        end;
      end;
    end;
    if TOBL.getValue('GL_PRIXPOURQTE')=0 then TOBL.putValue('GL_PRIXPOURQTE',1);
    if (TobOuv <> nil) and (TOBOUV.detail.count > 0) and (TOBOuv.detail[0].NomTable = 'LIGNEOUV') then
    begin
      TOBLN := TOB.create ('',TOBOUvrage,-1);
      InsertionChampSupOuv (TOBLN,false);
      TOBLN.dupliquer (TOBOuv,true,true);
      if (pos(TOBL.GetValue('GL_TYPEARTICLE'),'ARP;ARV')>0) and (not NaturepieceOKPourOuvrage(TOBL)) then
      begin
        if TOBLN.detail[0].detail.count > 1 then
        begin
          TOBLN.detail[0].detail[1].free;
        end;
      end;
      // --
      //FV1 : 30/11/20016 - Déplacement pour cohérence Numordre sur ouvrage et sous-détail
      TOBL.PutValue('GL_NUMORDRE',GetMaxNumOrdre(TOBpiece));
      // --
      if TOBLN.detail.count > 0 then  // On saute le groupe d'ouvrage
      begin
        DetailATraiter := true;
//        RepriseDonneeArticle (TOBLN,nil,True,false);
    	  Indice := 0;
        while Indice < TOBLN.detail.count do
        begin
          //
          ArticleNonOk := Not TraiteLigneCopierCollerOuv (TOBLN.detail[Indice],TOBPiece,TOBL,TOBArticles,Dev,MemeDoc);
          Inc(Indice);
          // --
        end;

        // Correction FQ 12072
        if not MemeDoc then
        begin
          TOBL.PutValue('GL_COEFFC',0);
          TOBL.PutValue('GL_COEFFR',0);
          TOBL.PutValue('GL_MONTANTFC',0);
          TOBL.PutValue('GL_MONTANTFR',0);
          //
          if (ArticleNonOk = True) or (TOBL.getValue('GL_DOMAINE')<>'') then
          begin
            InitTableau (Valeurs);
            CalculeOuvrageDoc (TOBLN,1,1,true,DEV,valeurs,(TOBPiece.getValue('GP_FACTUREHT')='X'),true,true);
            Qte := TOBL.Getvalue('GL_QTEFACT');
            TOBL.Putvalue('GL_MONTANTPAFG',valeurs[10]*Qte);
            TOBL.Putvalue('GL_MONTANTPAFR',valeurs[11]*Qte);
            TOBL.Putvalue('GL_MONTANTPAFC',valeurs[12]*Qte);
            TOBL.Putvalue('GL_MONTANTFG',valeurs[13]*Qte);
            TOBL.Putvalue('GL_MONTANTFR',valeurs[14]*Qte);
            TOBL.Putvalue('GL_MONTANTFC',valeurs[15]*Qte);
            TOBL.Putvalue('GL_MONTANTPA',Arrondi((Qte * TOBL.GetValue('GL_DPA')),V_PGI.okdecV));
            TOBL.Putvalue('GL_MONTANTPR',Arrondi((Qte * TOBL.GetValue('GL_DPR')),V_PGI.okdecV));

            TOBL.Putvalue('GL_DPA',valeurs[16]);
            TOBL.Putvalue('GL_DPR',valeurs[17]);
            //
            TOBL.Putvalue('GL_PUHTDEV',valeurs[2]);
            TOBL.Putvalue('GL_PUTTCDEV',valeurs[3]);
            TOBL.Putvalue('GL_PUHT',DeviseToPivotEx(TOBL.GetValue('GL_PUHTDEV'),DEV.Taux,DEV.quotite,V_PGI.OkdecP));
            TOBL.Putvalue('GL_PUTTC',DevisetoPivotEx(TOBL.GetValue('GL_PUTTCDEV'),DEV.taux,DEV.quotite,V_PGI.okdecP));
            TOBL.Putvalue('GL_PUHTBASE',TOBL.GetValue('GL_PUHT'));
            TOBL.Putvalue('GL_PUTTCBASE',TOBL.GetValue('GL_PUTTC'));
            TOBL.Putvalue('GL_DPA',valeurs[0]);
            TOBL.Putvalue('GL_DPR',valeurs[1]);
            TOBL.Putvalue('GL_PMAP',valeurs[6]);
            TOBL.Putvalue('GL_PMRP',valeurs[7]);
            TOBL.putvalue('GL_TPSUNITAIRE',valeurs[9]);
            TOBL.putvalue('GL_HEURE',Arrondi(valeurs[9]*Qte,V_PGI.okdecQ));
            TOBL.putvalue('GL_TOTALHEURE',TOBL.getDouble('GL_HEURE'));  
            StockeInfoTypeLigne (TOBL,valeurs)
          end;
        end;
      end;
      TOBL.Putvalue('GL_INDICENOMEN',TobOuvrage.detail.count);
      //
      NumeroteLigneOuv (TOBLN,TOBL,1,1,0,0,0);
      // Correction FQ 12072
      if not MemeDoc then
      begin
        TOBL.PutValue('GL_COEFFC',0);
        TOBL.PutValue('GL_COEFFR',0);
        TOBL.PutValue('GL_MONTANTFC',0);
        TOBL.PutValue('GL_MONTANTFR',0);
        TOBL.PutValue('GL_COEFMARG',0);
        // --
        {$IFDEF BTP}
        PuFix := TOBL.GetValue('GL_PUHTDEV');
        InitTableau (Valeurs);
        CalculeOuvrageDoc (TOBLN,1,1,true,DEV,valeurs,(TOBPiece.getValue('GP_FACTUREHT')='X'),true,true);
        if TOBL.GetValue('GLC_FROMBORDEREAU')='X' then
        begin
          // forcement en prix bloqué
          TraitePrixOuvrage(TOBPiece,TOBL,TOBBases,TOBBasesL ,TOBOuvrage,(TOBPiece.getValue('GP_FACTUREHT')='X'),PuFix,Valeurs[2],DEV,false,True);
        end else
        begin
          TraitePrixOuvrage(TOBPiece,TOBL,TOBBases,TOBBasesL ,TOBOuvrage,(TOBL.GetValue('GL_FACTUREHT')='X'),PuFix,0,DEV,false);
        end;
        StockeInfoTypeLigne (TOBL,valeurs)
      end;
      // -----------------
//      positionneCoefMarge (TOBL);
      //
      if (pos(TOBL.GetValue('GL_TYPEARTICLE'),'ARP;ARV')>0) and (not NaturepieceOKPourOuvrage(TOBL)) then
      begin
        TOBLN.free;
        TOBL.Putvalue('GL_INDICENOMEN',0);
      end;
      {$ENDIF}
      //
    end
    else
    begin
      if not MemeDoc then
      begin
        TOBL.putValue('GL_DPA',TOBLOc.getValue('GL_DPA'));
        TOBL.putValue('GL_DPR',TOBLOc.getValue('GL_DPA'));
        // Correction FQ 12072
        TOBL.PutValue('GL_COEFFC',0);
        TOBL.PutValue('GL_COEFFR',0);
        // --
        ReinitMontantAch (TOBL);
        //
        if (TOBPIece.GetValue('GP_COEFFC') <> 0) and (TOBL.GetValue('GLC_NONAPPLICFC')<>'X' ) THEN TOBL.PutValue('GL_COEFFC',TOBPIece.GetValue('GP_COEFFC'));
        if (TOBPIece.GetValue('GP_COEFFR') <> 0) and (TOBL.GetValue('GLC_NONAPPLICFRAIS')<>'X' ) THEN TOBL.PutValue('GL_COEFFR',TOBPIece.GetValue('GP_COEFFR'));
        // --
        if TOBL.getVAlue('GL_DOMAINE')<>'' then
        begin
          CalculeLigneAc (TOBL,DEV,true);
        end else
        begin
          CalculeLigneAc (TOBL,DEV,false);
        end;
        CalculFraisFromLigne (TOBpiece,TOBL);
  //      CalculeBaseFraisLigne (TOBL);
        TOBL.PutValue('GL_COEFMARG',0);
      end;
      TOBL.PutValue('GL_NUMORDRE',GetMaxNumOrdre(TOBpiece));
      TOBL.Putvalue('GL_INDICENOMEN',0);
    end;
    StringToRich( TFFacture(fform).Descriptif1, TOBL.GetValue ('GL_BLOCNOTE') );
    if (Length(TFFacture(fform).Descriptif1.text)<>0) and (TFFacture(fform).Descriptif1.text<>#$D#$A) then
    BEGIN
      TOBL.PutValue('GL_BLOCNOTE', ExRichToString (TFFacture(fform).Descriptif1) );
    END else
    BEGIN
      TOBL.PutValue ('GL_BLOCNOTE','');
    END;
    {$ENDIF}
    result := true;
    QteFact := TOBL.GetDouble('GL_QTEFACT');
    TOBL.PutValue('GL_RECALCULER','X');
    TOBL.PUTVALUE('GL_LIBELLE',TOBLOC.GetValue('GL_LIBELLE'));
    TOBL.PutValue('GL_POURCENTAVANC',0)      ;
    TOBL.PutValue('GL_QTEPREVAVANC',0)      ;
    TOBL.PutValue('GL_QTESIT',       0);
    TOBL.PutValue('GL_QTESTOCK',  QteFact);
    TOBL.PutValue('GL_QTERESTE',  QteFact);
    // --- GUINIER ---
    TOBL.PutValue('GL_MTRESTE', TOBL.GetValue('GL_MONTANTHTDEV'));

    //FV1 : 30/11/20016 - Déplacement pour cohérence Numordre sur ouvrage et sous-détail
    //TOBL.PutValue('GL_NUMORDRE',GetMaxNumOrdre(TOBpiece));
    //   TOBL.PutValue('GL_NUMORDRE',0)      ;

    ZeroLigneMontant (TOBL);  // reinit des montant de la ligne pour eviter de deduire les montants a ajouter

    {$IFDEF BTP}
    // Traitement des détail ouvrages
    if DetailATraiter then
    BEGIN
      //dec(Arow);
      AjouteUneLigneAraffraichir (TOBL.GetValue('GL_NUMORDRE'));
      //LoadLesLibDetOuvLig (TOBPIece,TOBOuvrage,TOBTiers,TOBAffaire,TOBL,Arow,DEV);
    END;
    //
    //C'est ici qu'il faut se placer pour gérer la copie des métrés...
    //Ne gérer que si les paramètres sont renseignés...
    if TheMetreDoc <> nil then
    begin
      If TheMetredoc.OkMetreBiblio then  ColleTobVarDoc(TOBL, TOBLOc, TobVarDoc, TOBL);
      if TheMetredoc.OkMetreDoc then
      begin
        if TheMetredoc.OkExcel then ColleTobMetres(TOBL, TOBLOc, TOBL, TheMetredoc);
      end;
    end;
    //

    {$ENDIF}
    inc(Arow);
  end;

  if TOBOuv <> nil then TOBOuv.free;
end;

procedure TCopieColleDoc.InsertLigneCommentaireImp (TOBLoc : TOB;Arow,Niveau : Integer);
var TOBL : TOB;
  InVariante : boolean;
begin
TOBL := GetTOBLigne (TobPiece,Arow);
InVariante := IsVariante (TOBL);
InsertTobLigne (TOBPiece,Arow);
TOBL := GetTOBLIgne (TOBPiece,Arow);
AddLesSupLigne(TOBL,False) ;  InitLesSupLigne (TOBL);
CopieChampsLigne ('LIGNE',TOBL,TOBLOC);
CopieChampsSup (TOBL,TOBLOC);
VerifChampsSupsLigne(TOBL);
//NewTOBLigneFille (TOBL);
PieceVersLigne (TOBPIece,TOBL);
AffaireVersLigne (TOBPiece,TOBL,TOBAffaire);
TOBL.PutValue('GL_NIVEAUIMBRIC',Niveau);
TOBL.PutValue('GL_PIECEPRECEDENTE','');
TOBL.PutValue('GL_PIECEORIGINE','');
TOBL.PutValue('GLC_DOCUMENTLIE','');
TOBL.PutValue('GLC_NUMEROTATION','');
TOBL.SetBoolean('GLC_NUMFORCED',false);
if VariantePieceAutorisee (TOBPiece) then
  begin
  if InVariante Then TOBL.Putvalue('GL_TYPELIGNE','COV')
                else TOBL.Putvalue('GL_TYPELIGNE',TOBLOc.getValue('GL_TYPELIGNE'))
  end else TOBL.Putvalue('GL_TYPELIGNE','COM');
{$IFDEF BTP}
if (not IsBLobVide(FF,TOBL,'GL_BLOCNOTE')) and (not IsLigneReferenceLivr(TOBL)) then
BEGIN
  TOBL.PutValue('GL_LIBELLE', 'BLOC NOTE ASSOCIE');
  TOBL.PutValue('BLOBONE', 'X');
END ELSE
BEGIN
	TOBL.PutValue ('GL_BLOCNOTE','');
  TOBL.PutValue('BLOBONE', '-');
END;
{$ELSE}
TOBL.PutValue ('GL_LIBELLE',TOBLoc.getValue('GL_LIBELLE'));
{$ENDIF}
TOBL.PutValue('GL_NUMORDRE',GetMaxNumOrdre(TOBpiece));
end;

procedure TCopieColleDoc.InsertLigneTotalImp (TOBLoc : TOB;Arow,Niveau: integer);
var TOBl : TOB;
begin
InsertTobLigne (TOBPiece,Arow);
TOBL := GetTOBLIgne (TOBPiece,Arow);
AddLesSupLigne(TOBL,False) ; InitLesSupLigne (TOBL);
CopieChampsLigne ('LIGNE',TOBL,TOBLOC);
CopieChampsSup (TOBL,TOBLOC);
VerifChampsSupsLigne(TOBL);
//NewTOBLigneFille (TOBL);
PieceVersLigne (TOBPIece,TOBL);
AffaireVersLigne (TOBPiece,TOBL,TOBAffaire);
TOBL.PutValue('GL_NIVEAUIMBRIC',Niveau);
TOBL.PutValue('GL_PIECEPRECEDENTE','');
TOBL.PutValue('GL_PIECEORIGINE','');
TOBL.PutValue('GLC_DOCUMENTLIE','');
TOBL.PutValue('GLC_NUMEROTATION','');
TOBL.SetBoolean('GLC_NUMFORCED',false);
TOBL.PutValue('GL_NUMORDRE',GetMaxNumOrdre(TOBpiece));
end;


Function TCopieColleDoc.IsArticleOk (Arow : integer):boolean;
var TOBArt,TOBL : TOB;
    rechart : T_RechArt;
    refsais : string;
begin
result := false;
TOBArt:=FindTOBArtRow (TOBPiece,TobArticles,Arow);
TOBL:=GetTOBLigne(TOBPiece,Arow) ; if TOBL = nil then exit;
if TOBArt<>Nil then
   BEGIN
   Result := true;
   END else
   BEGIN
   TOBArt:=CreerTOBArt(TOBArticles) ;
   RefSais := TOBL.GetValue('GL_CODEARTICLE');
   RechArt:=TrouverArticleSQL(TFFacture(fform).CleDoc.NaturePiece,RefSais,TFFacture(fform).GP_DOMAINE.Value,
                              '',TFFacture(fform).CleDoc.DatePiece,TOBArt,TOBTiers,TOBL.GetValue('GL_ARTICLE')) ;
   if (RechArt = traOk) or (RechArt = TraOkGene) then
      begin
      if Not ArticleAutorise(TOBPiece,TOBArticles,TFFacture(fform).CleDoc.NaturePiece,ARow) then
         BEGIN
         result := false;
         END else
         BEGIN
         result := true;
         END ;
      end;
   END ;
end;

procedure TCopieColleDoc.PositionneCell(Acol, Arow: integer);
begin
rowCollage := Arow;
colCollage := Acol;
end;

procedure TCopieColleDoc.rapeller;
begin
if fgrid = nil then Exit;
RappellerMemo;
RowCollage := -1;
ColCollage := -1;
end;

procedure TCopieColleDoc.deselectionneRows;
begin
{$IFDEF BTP}
fgrid.OnFlipSelection := nil;
fgrid.OnBeforeFlip := nil;
fgrid.AllSelected := false;
fgrid.OnFlipSelection := FlipSelection;
fgrid.OnBeforeFlip := BeforeFlip;
{$ELSE}
fgrid.ClearSelected ;
{$ENDIF}
end;

procedure TCopieColleDoc.RappellerMemo;
var Retour : string;
    TOBMemo : TOB;
    QQ :TQuery;
begin
deselectionneRows;
Retour := AGLLanceFiche ('BTP','BTMEMORIS_MUL','BMO_TYPEMEMO='+fTypeInfo,'','');
if Retour <> '' then
   begin
   TOBMemo := TOB.create ('BMEMORISATION',nil,-1);
   QQ := OpenSQL('Select BMO_MEMO from BMEMORISATION Where BMO_TYPEMEMO="'+fTypeInfo+'" AND' +
             ' BMO_CODEMEMO="'+Retour+'"',True,-1,'',true);
   TOBMemo.SelectDB ('',QQ);
   ferme(QQ);
   fmemo.Clear;
   fmemo.text := TOBMemo.getValue('BMO_MEMO');
   TobMemo.Free;
   Collagedonnee;
   deselectionneRows;
   end;
end;


procedure TCopieColleDoc.GSContextPopup(Sender: TObject; MousePos: TPoint;var Handled: Boolean);
var Acol,Arow : Integer;
    TOBL : TOB;
    RowOri,ColOri : integer;
begin
  if (MousePos.X = -1) or (MousePos.y = -1) then exit;
  fgrid.MouseToCell (MousePos.X,MousePos.Y,Acol,ARow);
  TOBL:=GetTOBLigne(TOBPiece,Arow) ;
  if TOBL = nil then
     begin
     Handled := true;
     exit;
     end;
  if Acol < fgrid.FixedCols then Acol := fgrid.fixedCols;
//  PositionneCell(Acol,Arow);
//  TFfacture(FF).GoToLigne(RowOri,ColOri);
  if TFFacture(fform).Action = TaConsult then MenuEnabled (false)
                                         else MenuEnabled (true);
end;

procedure TCopieColleDoc.GSCollerClick(Sender: TObject);
begin
Coller;
end;


procedure TCopieColleDoc.GSCouperClick(Sender: TObject);
begin
Couper;
end;

procedure TCopieColleDoc.GScopierClick(Sender: TObject);
begin
Copier;
end;

procedure TCopieColleDoc.GSEnregistrerClick(Sender: TObject);
begin
Enregistrer;
end;

procedure TCopieColleDoc.GSRappelerClick(Sender: TObject);
begin
rapeller;
end;

destructor TCopieColleDoc.destroy;
var indice : integer;
begin
  inherited;
  TOBLR.free;
  for Indice := 0 to fMaxItems -1 do
    begin
    MesMenuItem[Indice].Free;
    end;
if fcreatedPop then POPGS.free;
end;

procedure TCopieColleDoc.MenuEnabled(State: boolean);
var Indice : integer;
begin
  for Indice := 0 to fMaxItems -1 do
  begin
    if (MesMenuItem[Indice].name <> 'BCOPIER') AND
       (MesMenuItem[Indice].name <> 'BCOUPER') AND
       (MesMenuItem[Indice].name <> 'BENREGISTRER') then MesMenuItem[Indice].Enabled := State;
  end;
end;


procedure InitChampsSupNULL (TOBAssoc: TOB);
var NomChamp : string;
    Ind2 : integer;
    Result : string;
begin
   ind2 := 1000;
   NomChamp := TobAssoc.GetNomChamp(ind2);
   while NomChamp <> '' do
   begin
      Result := VarAsType(TobAssoc.GetValeur(ind2), varString);
      if Result = #0 then TobAssoc.PutValeur(ind2,'');
      Inc(ind2);
      NomChamp := TobAssoc.GetNomChamp(ind2);
   end;
end;


procedure TCopieColleDoc.VerifChampsSupsLigne(TOBL: TOB);
var Indice : integer;
		NomChamp : string;
begin
//
  if TOBL.fieldExists ('TOTALACHAT') then
    if VarType (TOBL.GetValue('TOTALACHAT')) = VarString then
      TOBL.PutValue ('TOTALACHAT',Valeur(TOBL.GetValue('TOTALACHAT')));
  if TOBL.fieldExists ('TOTALREVIENT') then
    if VarType (TOBL.GetValue('TOTALREVIENT')) = VarString then
      TOBL.PutValue ('TOTALREVIENT',Valeur(TOBL.GetValue('TOTALREVIENT')));
(*
  if TOBL.fieldExists ('TPSUNITAIRE') then
    if VarType (TOBL.GetValue('TPSUNITAIRE')) = VarString then
      TOBL.PutValue ('TPSUNITAIRE',Valeur(TOBL.GetValue('TPSUNITAIRE')));
*)
  if TOBL.fieldExists ('__COEFMARG') then
    if VarType (TOBL.GetValue('__COEFMARG')) = VarString then
      TOBL.PutValue ('__COEFMARG',Valeur(TOBL.GetValue('__COEFMARG')));
  if TOBL.fieldExists ('GLC_NATUREPIECEG') then
    if VarType (TOBL.GetValue('GLC_NATUREPIECEG')) = VarString then
      TOBL.PutValue ('GLC_NATUREPIECEG','');
  if TOBL.fieldExists ('GLC_SOUCHE') then
    if VarType (TOBL.GetValue('GLC_SOUCHE')) = VarString then
      TOBL.PutValue ('GLC_SOUCHE','');
  if TOBL.fieldExists ('GLC_NUMERO') then
    if VarType (TOBL.GetValue('GLC_NUMERO')) = VarString then
      TOBL.PutValue ('GLC_NUMERO',TOBL.GetValue('GL_NUMERO'));
  if TOBL.fieldExists ('GLC_INDICEG') then
    if VarType (TOBL.GetValue('GLC_INDICEG')) = VarString then
      TOBL.PutValue ('GLC_INDICEG',TOBL.GetValue('GL_INDICEG'));
  if TOBL.fieldExists ('GLC_NUMORDRE') then
    if VarType (TOBL.GetValue('GLC_NUMORDRE')) = VarString then
      TOBL.PutValue ('GLC_NUMORDRE',TOBL.GetValue('GL_NUMORDRE'));
  if TOBL.fieldExists ('GLC_QTACCSAIS') then
    if VarType (TOBL.GetValue('GLC_QTACCSAIS')) = VarString then
      TOBL.PutValue ('GLC_QTACCSAIS',0);
  if TOBL.fieldExists ('GLC_DATEACC') then
    if VarType (TOBL.GetValue('GLC_DATEACC')) = VarString then
      TOBL.PutValue ('GLC_DATEACC',iDate1900);
  if TOBL.fieldExists ('GLC_IDENTIFIANTWNT') then
    if VarType (TOBL.GetValue('GLC_IDENTIFIANTWNT')) = VarString then
      TOBL.PutValue ('GLC_IDENTIFIANTWNT',0);
  if TOBL.fieldExists ('GLC_LIGNEORDRE') then
    if VarType (TOBL.GetValue('GLC_LIGNEORDRE')) = VarString then
      TOBL.PutValue ('GLC_LIGNEORDRE',0);
  if TOBL.fieldExists ('GLC_DATEDEB') then
    if VarType (TOBL.GetValue('GLC_DATEDEB')) = VarString then
      TOBL.PutValue ('GLC_DATEDEB',iDate1900);
  if TOBL.fieldExists ('GLC_DATEFIN') then
    if VarType (TOBL.GetValue('GLC_DATEFIN')) = VarString then
      TOBL.PutValue ('GLC_DATEFIN',iDate1900);
//
  if TOBL.fieldExists ('GCA_PRIXBASE') then
    if VarType (TOBL.GetValue('GCA_PRIXBASE')) = VarString then
      TOBL.PutValue ('GCA_PRIXBASE',Valeur(TOBL.GetValue('GCA_PRIXBASE')));
  if TOBL.fieldExists ('GA_PAHT') then
    if VarType (TOBL.GetValue('GA_PAHT')) = VarString then
      TOBL.PutValue ('GA_PAHT',Valeur(TOBL.GetValue('GA_PAHT')));
  if TOBL.fieldExists ('GA_PVHT') then
    if VarType (TOBL.GetValue('GA_PVHT')) = VarString then
      TOBL.PutValue ('GA_PVHT',Valeur(TOBL.GetValue('GA_PVHT')));
  if TOBL.fieldExists ('BLO_PRXACHBASE') then
    if VarType (TOBL.GetValue('BLO_PRXACHBASE')) = VarString then
      TOBL.PutValue ('BLO_PRXACHBASE',Valeur(TOBL.GetValue('BLO_PRXACHBASE')));
//
for Indice := 0 to TOBL.ChampsSup.Count -1 do
begin
	NomChamp := TOBL.GetNomChamp(1000+indice);
	if VarIsNull(TOBL.GetValue(NomChamp)) or (VarAsType(TOBL.GetValue(NomChamp), varString) = #0 ) then
  begin
  	PGIInfo ('ATTENTION : le champs : '+NomChamp+' est NULL');
  end;
end;

end;

procedure TCopieColleDoc.HideMenu;
var Indice : integer;
begin
  for Indice := 0 to fMaxItems -1 do
  begin
    if (MesMenuItem[Indice].name = 'BCOPIER') OR
    	 (MesMenuItem[Indice].name = 'BCOUPER') OR
    	 (MesMenuItem[Indice].name = 'BCOLLER') OR
       (MesMenuItem[Indice].name = 'BENREGISTRER') or
       (MesMenuItem[Indice].name = 'BRAPELLER') then MesMenuItem[Indice].Visible := false;
  end;
end;

procedure TCopieColleDoc.Activate;
begin
  if (tModeSaisieBordereau in TFFacture(FF).SaContexte) then HideMenu;
end;

procedure TCopieColleDoc.CopieChampsLigne(Table : string; TOBLOC,Tobl: Tob);
var
  TobLigne: Tob;
  i: Integer;
  NomChamp: String;
begin
  if Assigned(Tobl) then
  begin
    TobLigne := Tob.Create(Table, nil, -1);
    try
      for i := 1 to TobLigne.NbChamps do
      begin
        NomChamp := TobLigne.GetNomChamp(i);
        if TOBLoc.fieldExists (NomChamp) then
        begin
          TOBLoc.putValue(NomChamp,TOBL.GetValue(NomChamp));
        end;
      end;
    finally
      TobLigne.Free;
    end;
  end;
end;

procedure TCopieColleDoc.CopieDonneeOuvrage(TOBDest, TOBOUV, TOBL: TOB; Ok_MetreDoc, Ok_MetreBib, Ok_MetreExcel : Boolean);
begin
	// 1er niveau
  InsertionChampSupOuv (TOBDest,false);
  CopieChampsSup (TOBDest,TOBOUV);
  // niveau suivant
  CopieDetailOuvrage (TOBDest,TOBOUV, TOBL, Ok_MetreDoc, Ok_MetreBib, Ok_MetreExcel);
end;

procedure TCopieColleDoc.CopieDetailOuvrage (TOBDest,TOBOUV, TOBL : TOB; Ok_MetreDoc, Ok_MetreBib, Ok_MetreExcel : Boolean);
var Indice : integer;
		TOBSuite,TOBDETOUV : TOB;
begin
  for Indice := 0 to TOBOUV.detail.count -1 do
  begin
    TOBDETOUV := TOBOUV.detail[Indice];
    TOBSUite := TOB.Create ('LIGNEOUV',TOBDest,-1);
//    TOBSUite := TOBDest.detail[tobdest.detail.count-1];
    InsertionChampSupOuv (TOBSuite,false);
  	CopieChampsLigne ('LIGNEOUV',TOBSuite,TOBDETOUV);
    CopieChampsSup (TOBSuite,TOBDETOUV);
    VerifChampsSupsDetail (TobSuite);
  	if isBlobVide (FF,TOBSuite,'BLO_BLOCNOTE') then TOBSuite.putValue('BLO_BLOCNOTE','');
    //
    //Ne faire cela que si les metrés Excel sur document sont paramétrés
    if Ok_MetreBib then CopieTobVarDoc (TTObPiece, TOBSUite, TOBL, TobVardoc);
    if OK_Metredoc then
    begin
      if Ok_MetreExcel then CopieTobMetres (TTObPiece, TOBSUite, TOBL, TobMetres);
    end;
    //  	InitChampsSupNULL (TOBSuite);
    TOBPiece.putValue('GP_UNIQUEBLO', TOBPiece.getValue('GP_UNIQUEBLO')+1);
    TOBSUITE.putValue('BLO_UNIQUEBLO',TOBPiece.getValue('GP_UNIQUEBLO'));
    //
    if TOBDETOUV.detail.count > 0 then CopieDetailOuvrage (TOBSuite,TOBDetOUV, TOBL, Ok_MetreDoc, Ok_MetreBib, Ok_MetreExcel);
  end;
end;


procedure TCopieColleDoc.AjouteUneLigneAraffraichir(Arow : integer);
var TOBALR : TOB;
begin
  TOBALR := TOB.Create('UNE LIGNE',TOBLR,-1);
  TOBALR.AddChampSupValeur ('NUMORDRE',Arow);
end;

function TCopieColleDoc.FindLigne (NumOrdre : integer) : TOB;
var Indice : integer;
    TOBL : TOB;
begin
  result := nil;
  Indice := TOBPiece.detail.count -1;
  repeat
    TOBL := TOBPiece.detail[Indice];
    if TOBL.getValue('GL_NUMORDRE')=NumOrdre then
    begin
      result := TOBL;
      break;
    end else dec(Indice);
  until Indice = 0;
end;

procedure TCopieColleDoc.ReinitLigneFac(TOBL: TOB);
var NomChamp : string;
    Indice : integer;
begin
  if not TOBL.FieldExists('BLF_NATUREPIECEG') then exit;
  for Indice := 0 to TOBL.ChampsSup.Count -1 do
  begin
		NomChamp := TOBL.GetNomChamp(1000+indice);
    if (pos(NomChamp,'BLF_NATUREPIECEG;BLF_SOUCHE;BLF_FOURNISSEUR')> 0) then
    begin
      TOBL.PutValeur(1000+Indice,'');
    end else if (copy(NomChamp,1,6)='BLF_MT') or (copy(NomChamp,1,6)='BLF_QTE') or (pos(NomChamp,'BLF_TOTALTTCDEV')> 0) then
    begin
      TOBL.PutValeur(1000+Indice,0);
    end else if (pos(NomChamp,'BLF_UNIQUEBLO')> 0) then
    begin
      TOBL.PutValeur(1000+Indice,0);
    end;
  end;
end;

procedure TCopieColleDoc.SetInfoLigne(TOBl: TOB);
var TheMenuItem : TMenuItem;
  Indice : Integer;
begin
  TheMenuItem := nil;
  for Indice := 0 to fMaxItems -1 do
  begin
    if (MesMenuItem[Indice].name = 'TImprimable') then
    begin
      TheMenuItem := MesMenuItem[Indice];
      break;
    end;
  end;
  if (TheMenuItem <> nil) and (TOBL <> nil)  then
  begin
    if TOBL.GetBoolean('GL_NONIMPRIMABLE') then
    begin
      TheMenuItem.Caption := TraduireMemoire ('A Imprimer');
    end else
    begin
      TheMenuItem.Caption := TraduireMemoire ('Ne pas Imprimer');
    end;
  end;
end;

procedure TCopieColleDoc.GereImprimable(Sender: TObject);
var TOBL : TOB;
begin
  if fgrid = nil then Exit;
  if (TFFacture(fform).Action = taConsult) then Exit;
  TOBL := GetTOBLigne(TOBPiece,fgrid.row);
  if TOBL.GetBoolean('GL_NONIMPRIMABLE') then
  begin
    TOBL.SetBoolean('GL_NONIMPRIMABLE',false);
  end else
  begin
    TOBL.SetBoolean('GL_NONIMPRIMABLE',true);
  end;
  SetInfoLigne (TOBL);
  fgrid.InvalidateRow(fgrid.row); 
end;

end.
