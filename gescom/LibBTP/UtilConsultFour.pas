unit UtilConsultFour;

interface
uses Classes,UTOB,FactComm,Hctrls,HEnt1,SaisUtil,UtilPGi,forms,SysUtils,
{$IFDEF EAGLCLIENT}
  maineagl,UtileAGL,Ed_Tools
{$ELSE}
  Doc_Parser,DBCtrls, Db, {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} fe_main,uPDFBatch,EdtREtat
{$ENDIF}
	,HPdfPrev
  ,Menus
	,HmsgBox
  ,HTB97
  ,Hpanel
  ,Graphics
  ,Grids
  ,Vierge
  ,Windows
  ,DecisionAchatUtil
  ,AglInit
  ,hPDFViewer
  ,Messages
  ,ParamSoc
  ,Controls
;


type
TGestConsultFou = class
  private
    Usable : boolean;
    stprev : string;
    fCodeEtat : string;
    XX : TForm;
    GS : THgrid;
    BDELCONSFOU : TToolbarButton97;
    BSORTIECONSFOU : TToolbarButton97;
    PCONSULTFOU : THPanel;
    BCONSULTFOU : TToolbarButton97;
    fConsultWin : TToolWindow97;
    CHOIXFOURC,POPM2  : TPopupMenu;
    MnCatalogueC,MnFournisseurC : TmenuItem;
    Mndelete,MnSelection : TmenuItem;
    MaxIndice : integer;
    //
    TOBConsultFou : TOB;
    TOBDecisionnel : TOB;
    TheCurrentDecisionnel : TOB;
    TheTOBCons : TOB; // la consultation courante
    TOBFourn : tob;
    //
    FnumDecisionnel : string;
    FNomTable, FLien, FSortBy, stCols, FTitre,stColListe : string;
    FLargeur, FAlignement, FParams, title, NC, FPerso: string;
    OkTri, OkNumCol : boolean;
    nbcolsInListe : integer;
    FFQte : string;
    //
    G_NUMORDRE,G_FOURNISSEUR,G_LIBELLE,G_PA,G_DELAI,G_SELECT,G_REFARTTIERS : integer;
    //
    procedure AddChampSupFille(TOBL: TOB);
    function  AddFille (TOBI : TOB) : TOB;
    function  AddFilleFromDecisionnel(TOBD: TOB): TOB;
    procedure AfficheLagrille(TOBC: TOB);
    procedure AfficheLigne(Grid: THgrid; TOBL: TOB; indice: integer);
    function  AjouteLignecons : TOB;
    procedure ConstitueLaTOBConsult(TOBInterm: TOB);
    procedure ConsultFouClick(Sender: TOBject);
    function  ControleLigne (Arow : integer ; var Acol : integer) : boolean;
    procedure DefiniRowCounts;
    procedure DeleteClick (Sender : TObject);
    procedure DeleteLigne;
    function  DeleteConsulFouDetail (TOBDecisionL : TOB) : boolean;
    function  DeleteDownDecision(TOBDECISIONL: TOB): boolean;
    function  DeleteUpDecision (TOBDECISIONL : TOB) : boolean;
    procedure DelConsfouClick (Sender : TObject);
    //
    procedure DefiniGrille;
    procedure DefiniToolWindow;
    function  ExistConsulFouDetail (TOBDecisionL : TOB) : boolean;
    procedure fConsultWinClose(Sender: TObject);
    procedure FermeConsFou (Sender : TObject);
    procedure GetComponents;
    function  GetConsulFou(TOBD: TOB): TOB; overload;
    function  GetConsulFou(Indice : integer): TOB; overload;
    //
    // evenement sur grid et form
    procedure GSCellEnter(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
    procedure GSCellExit(Sender: TObject; var ACol, ARow: Integer;  var Cancel: Boolean);
    procedure GSElipsisClick(Sender: TOBject);
    procedure GSRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
    procedure GSRowExit(Sender: TObject; Ou: Integer;  var Cancel: Boolean; Chg: Boolean);
		procedure GSPostDrawCell(ACol, ARow: Longint; Canvas: TCanvas; AState: TGridDrawState);
    procedure GSKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure GetCellCanvas(ACol, ARow: Longint; Canvas: TCanvas; AState: TGridDrawState);
    //
    function  FindFournisseur (TOBL : TOB; Valeur : string) : boolean;
    procedure IndiceLaTOBCOnsult ;
    function  IsFournisseurExiste (Arow,Acol : integer) : boolean;
    function  IsLigneVide(Arow: integer): boolean; overload;
    function  IsLigneVide(TOBL : TOB): boolean; overload;
    procedure libereTout;
    procedure MnCatalogueClick(Sender: TObject);
    procedure MnFournisseurClick(Sender: TObject);
    procedure NettoieConsult(TheTobCons: TOB);
    procedure PositionneDansGrid(Arow, Acol: integer);
    procedure RemplitlaGrille ;
    procedure ReindiceGrid;
    procedure ResizeToolWin (Sender : Tobject);
    //
    procedure SetBranche;
    procedure SetComponents;
    procedure SetEventsGrid (Actif : boolean);
    procedure SetParamgrille;
    procedure SetPrimalEvents;
    procedure SetNumDecisionnel(const Value: string);
    function  VerifDownDecision(TOBDECISIONL: TOB): boolean;
    function  VerifUpDecision(TOBDECISIONL: TOB): boolean;
    function  ZoneAccessible(Grille: THgrid; var ACol,ARow: integer): boolean;
    procedure ZoneSuivanteOuOk(Grille: THGrid; var ACol, ARow: integer;var Cancel: boolean);
    procedure SetConsultPere(TOBDep: TOB; Code: string);
    procedure SetConsultFille(TOBDep: TOB; Code: string);
    procedure PositionneIndiceConsult(TOBC: TOB; Code: string;Indice: integer; TraitementSurGrille : boolean = true);
    procedure SelectionneLeFourn(Sender: TObject);
    procedure EnleveAutreSelection;
    procedure RenumeroteFille(TOBD, TOBC: TOB);
    procedure BeforeSortie;
    procedure RecupLesCataLogSelect(TheTOBparam: TOB);
    procedure GSEnter(Sender: TObject);
    procedure AjouteLesCatalogues(TheTOBCons: TOB);
    procedure SetAdresseFou(TOBL : TOB;Fournisseur: string);
    procedure SetUnusable(const Value: boolean);
    procedure AffecteDecisionnel(TOBD,TOBF: TOB);
  public
    constructor create (FF : TForm);
    destructor destroy ; override;
    property NumDecisionnel : string write SetNumDecisionnel;
    property NonUtilisable : boolean write SetUnusable;
    procedure RowChange;
    procedure Renumerote(TOBD: TOB);
    procedure Supprime;
    procedure ecrit;
    procedure PrepareEdition ;
    function IsExisteConsult : boolean;
    procedure SetEtat (LeCodeEtat : string);
    procedure MultipleAffecte (TOBFournisseurs : TOB); // provenance de la multi-selection
    procedure PositionneTheConsult(TOBC: TOB);
end;

implementation
uses DECISIONACHAT_TOF,FactOuvrage,FactGrpBtp,FactAdresseBtp,UTofListeInv;

{ TGestConsultFou }



procedure TGestConsultFou.AddChampSupFille (TOBL : TOB);
begin
  TOBL.AddChampSupValeur ('BDF_NUMERO',0);
  TOBL.AddChampSupValeur ('BDF_N1',0);
  TOBL.AddChampSupValeur ('BDF_N2',0);
  TOBL.AddChampSupValeur ('BDF_N3',0);
  TOBL.AddChampSupValeur ('BDF_N4',0);
  TOBL.AddChampSupValeur ('INDICECONSULT',0);
end;

Function TGestConsultFou.AddFilleFromDecisionnel (TOBD : TOB) : TOB;
begin
  Result := TOB.Create ('DECISIONACH L FOU',TOBConsultFou,-1);
  AddChampSupFille (result);
  // NEW LS
  result.putValue('BDF_NUMERO',TOBD.GetValue('BAD_NUMERO'));
  result.putValue('BDF_N1',TOBD.GetValue('BAD_NUMN1'));
  result.putValue('BDF_N2',TOBD.GetValue('BAD_NUMN2'));
  result.putValue('BDF_N3',TOBD.GetValue('BAD_NUMN3'));
  result.putValue('BDF_N4',TOBD.GetValue('BAD_NUMN4'));
end;

Function TGestConsultFou.AddFille(TOBI: TOB) : TOB;
begin
  Result := TOB.Create ('DECISIONACH L FOU',TOBConsultFou,-1);
  AddChampSupFille (result);
  // NEW LS
  result.putValue('BDF_NUMERO',TOBI.GetValue('BDF_NUMERO'));
  result.putValue('BDF_N1',TOBI.GetValue('BDF_N1'));
  result.putValue('BDF_N2',TOBI.GetValue('BDF_N2'));
  result.putValue('BDF_N3',TOBI.GetValue('BDF_N3'));
  result.putValue('BDF_N4',TOBI.GetValue('BDF_N4'));
end;

procedure TGestConsultFou.AfficheLagrille (TOBC : TOB);
var Indice : integer;
begin
  SetEventsGrid (False);
  GS.VidePile (false);
  DefiniRowCounts;
  for Indice := 0 to TOBC.detail.count -1 do
  begin
    AfficheLigne(GS,TOBC.detail[Indice],Indice);
  end;
  TFVierge(XX).HMTrad.ResizeGridColumns (GS);
  GS.Visible := True;
  // entree dans la grille
  GS.SetFocus;
  GSEnter(self);
  // et hop on active les evenements sur la grille
  SetEventsGrid (True);
end;

procedure TGestConsultFou.AfficheLigne(Grid : THgrid;TOBL : TOB;indice : integer);
begin
  TOBL.PutLigneGrid (Grid,Indice+1,false,false,stColListe);
end;

function TGestConsultFou.AjouteLignecons : TOB;
begin
   result := TOB.Create ('DECISIONACHLFOU',TheTOBCons,-1);
   result.AddChampSupValeur  ('T_LIBELLE','');
   result.PutValue ('BDF_NUMORDRE',TheTOBCons.detail.count);
end;

procedure TGestConsultFou.ConstitueLaTOBConsult (TOBInterm : TOB);
var Indice : Integer;
		TOBI,TOBN : TOB;
    N1,N2,N3,N4 : integer;
begin
	Indice := 0;
	repeat
  	TOBI := TOBInterm.detail[Indice];
    N1 := TOBI.GetValue('BDF_N1');
    N2 := TOBI.GetValue('BDF_N2');
    N3 := TOBI.GetValue('BDF_N3');
    N4 := TOBI.GetValue('BDF_N4');
    TOBN := TOBconsultFou.findFirst (['BDF_N1','BDF_N2','BDF_N3','BDF_N4'],[N1,N2,N3,N4],true);
    if TOBN = nil then
    begin
      TOBN := AddFille (TOBI);
    end;
    if TOBN <> nil then TOBI.ChangeParent (TOBN,-1) else inc(Indice);
  until (TobInterm.Detail.count = 0) or (Indice >= TobInterm.Detail.count) ;
end;

procedure TGestConsultFou.ConsultFouClick (Sender : TOBject);
begin
  SetBranche;
  if TheCurrentDecisionnel = nil then BEGIN BCONSULTFOU.down := false; exit; END;
  if GetConsulFou(TheCurrentDecisionnel) = nil then
  begin
//
    if ExistConsulFouDetail(TheCurrentDecisionnel) then
    begin
      if fConsultWin.Visible then
      begin
        if PgiAsk('Une consultation fournisseur existe à un niveau différent. Supprimer ?') = MrYes then
        begin
          DeleteConsulFouDetail (TheCurrentDecisionnel);
          fConsultWin.Visible := BCONSULTFOU.Down;
        end else
        begin
          fConsultWin.Visible := false;
        end;
      end;
    end else
    begin
      fConsultWin.Visible := BCONSULTFOU.Down;
    end;
  end else
  begin
    fConsultWin.Visible := BCONSULTFOU.Down;
  end;
  if fConsultWin.Visible then
  begin
    RemplitlaGrille ;
  end else
  begin
    beforeSortie;
  end;
end;

function TGestConsultFou.ControleLigne(Arow: integer; var Acol : integer): boolean;
begin
  result := true;
end;

constructor TGestConsultFou.create(FF: TForm);
var Indice : integer;
begin
  usable := true;
  XX := FF;
  MaxIndice := 0;
  //
  FFQTE := '###';
  if V_PGI.OkDecQ > 0 then
  begin
    FFQTE := FFQTE+'0.';
    for indice := 1 to V_PGI.OkDecP do
    begin
      FFQTE := FFQTE + '0';
    end;
  end;
  //
  TOBConsultFou := TOB.create ('LES CONSULTFOURN',nil,-1);
  TOBFourn := TOB.Create ('LES FOURNISSEURS',nil,-1);
  //
  GetComponents;
  if not Usable then exit;
  DefiniToolWindow;
  SetParamgrille;
  DefiniGrille;
  SetComponents;
  SetPrimalEvents;
  //
end;


procedure TGestConsultFou.DefiniGrille;
var st,lestitres,lesalignements,FF,alignement,Nam,leslargeurs,lalargeur,letitre : string;
    Obli,OkLib,OkVisu,OkNulle,OkCumul,Sep : boolean;
    dec : integer;
    indice : integer;
    FFLQTE : string;
begin

  GS.ColCount := NbColsInListe;

  st := stColListe;
  lesalignements := Falignement;
  lestitres := FTitre;
  leslargeurs := flargeur;

  for indice := 0 to nbcolsInListe -1 do
  begin
    Nam := ReadTokenSt (St); // nom
    alignement := ReadTokenSt(lesalignements);
    lalargeur := readtokenst(leslargeurs);
    letitre := readtokenst(lestitres);
    TransAlign(alignement,FF,Dec,Sep,Obli,OkLib,OkVisu,OkNulle,OkCumul) ;

    GS.cells[Indice,0] := leTitre;

    if copy(Alignement,1,1)='G' then GS.ColAligns[indice] := taLeftJustify
    else if copy(Alignement,1,1)='D' then GS.ColAligns[indice] := taRightJustify
    else if copy(Alignement,1,1)='C' then GS.ColAligns[indice] := taCenter;

    GS.ColWidths[indice] := strtoint(lalargeur);

    if OkLib then GS.ColFormats[indice] := 'CB=' + Get_Join(Nam)
    else if (Dec<>0) or (Sep) then GS.ColFormats[indice] := FF ;

    if (nam = 'BDF_PAHT') then
    begin
      if OkNulle then FFLQte := FFQTE+';-'+FFQTE+'; ;';
      GS.ColFormats[indice] := FFLQTE ;
    end else if Nam='BDF_NBJOUR' then
    begin
      if OkNulle then FFLQte := FF+';-'+FF+'; ;';
      GS.ColFormats[indice] := FFLQTE ;
    end else if Nam='BDF_SELECTION' then
    begin
      GS.ColTypes [indice] := 'B' ;
      GS.colaligns[indice]:= tacenter;
      GS.colformats[indice]:= inttostr(Integer(csCoche));
    end;
  end ;
end;

procedure TGestConsultFou.DefiniRowCounts;
begin
  if TheTOBCOns.detail.count = 0 then AjouteLignecons;
  if TheTOBCOns.detail.count > 0 then GS.RowCount := TheTOBCOns.detail.count +2;
  if GS.RowCount < 2 then GS.rowcount := 3;
end;

procedure TGestConsultFou.DefiniToolWindow;
begin
  if not Usable then exit;
  fConsultWin := TToolWindow97.create (XX);
  fConsultWin.Parent := XX;
  fConsultWin.Caption := 'Consultations fournisseurs';
  fConsultWin.BorderStyle := bsSingle;
  fConsultWin.clientAreaHeight := PCONSULTFOU.Height ;
  fConsultWin.clientAreaWidth := PCONSULTFOU.Width;
  fConsultWin.ClientHeight := fConsultWin.clientAreaHeight;
  fConsultWin.clientWidth := fConsultWin.clientAreaWidth;
  fConsultWin.DragHandleStyle := dhDouble;
  fConsultWin.HideWhenInactive := True;
  fConsultWin.fullsize := false;
  fConsultWin.Resizable := true;
  fConsultWin.Height := PCONSULTFOU.height + 50;
  fConsultWin.Width := PCONSULTFOU.width + 10;
  fConsultWin.top := PCONSULTFOU.top + ((XX.height - fConsultWin.height) div 2);
  fConsultWin.left := PCONSULTFOU.left + ((XX.width - fConsultWin.width) div 2);
  PCONSULTFOU.parent := fConsultWin;
  PCONSULTFOU.align := alclient;
  fConsultWin.Visible := false;
  fconsultWin.OnResize := ResizeToolWin;
end;

procedure TGestConsultFou.DeleteLigne;
var LastRow : integer;
begin
  lastrow := GS.row;
  GS.DeleteRow (GS.row);
  if GS.row <= TheTOBCons.detail.count then
  begin
    TheTOBCons.detail[GS.row-1].free;
  end;
  ReindiceGrid;
  RemplitlaGrille;
  if lastrow > TheTOBCons.detail.count then lastRow := TheTOBCons.detail.count;
  PositionneDansGrid (LastRow,1);
end;

function TGestConsultFou.DeleteConsulFouDetail(TOBDecisionL: TOB) : boolean;
var Niv : string;
    TOBD : TOB;
begin
  result := false;
  Niv := TOBDecisionL.getValue('BAD_TYPEL');
  // en remontant
  TOBDecisionL.putValue('POSITIONCONSULT','AUC');
  TOBD := GetConsulFou (TOBDecisionL.GetValue('INDICECONSULT')); if TOBD <> nil then TOBD.Free;
  TOBDecisionL.putValue('INDICECONSULT',0);
  if (Niv='ART') or (Copy(Niv,3,1) > '1') then
  begin
    result := DeleteUpDecision (TOBDECISIONL);
  end;
  if result then exit;
  // En descendant
  if TOBDecisionL.getValue('BAD_TYPEL') <> 'ART' then
  begin
    result := DeleteDownDecision (TOBDECISIONL);
  end;
end;

function TGestConsultFou.DeleteDownDecision (TOBDECISIONL : TOB) : boolean;
var fiston : TOB;
    Indice : integer;
    TOBD : TOB;
begin
  result := false;
  if TOBDECISIONL.detail.count = 0 then exit;
  for Indice := 0 to TOBDECISIONL.detail.count -1 do
  begin
    fiston := TOBDECISIONL.detail[Indice];
    TOBD := GetConsulFou (Fiston);
    if TOBD <> nil  then
    BEGIN
      Fiston.putValue('INDICECONSULT','0');
      Fiston.putValue('POSITIONCONSULT','AUC');
      TOBD.free;
      result := true;
      break ;
    END;
    Fiston.putValue('INDICECONSULT','0');
    Fiston.putValue('POSITIONCONSULT','AUC');
    if DeleteDownDecision (fiston) then BEGIN result := true; break ; END;
  end;
end;

function TGestConsultFou.DeleteUpDecision(TOBDECISIONL: TOB): boolean;
var Papounet,TOBD : TOB;
begin
  result := false;
  Papounet := TOBDECISIONL.Parent;
  if Papounet = nil then exit;
  TOBD := GetConsulFou (Papounet);
  if TOBD <> nil then
  BEGIN
    TOBD.free;
    result := true;
    Papounet.putValue('INDICECONSULT','0');
    Papounet.putValue('POSITIONCONSULT','AUC');
    Exit;
  END;
  Papounet.putValue('INDICECONSULT','0');
  Papounet.putValue('POSITIONCONSULT','AUC');
  if DeleteUpDecision (papounet) then BEGIN result := true; Exit; END;
end;

destructor TGestConsultFou.destroy;
begin
  libereTout;
  inherited;
end;

function TGestConsultFou.ExistConsulFouDetail(TOBDecisionL: TOB): boolean;
var Niv : string;
begin
  result := false;
  Niv := TOBDecisionL.getValue('BAD_TYPEL');
  // en remontant
  if (Niv='ART') or (Copy(Niv,3,1) > '1') then
  begin
    result := VerifUpDecision (TOBDECISIONL);
  end;
  if result then exit;
  // En descendant
  if TOBDecisionL.getValue('BAD_TYPEL') <> 'ART' then
  begin
    result := VerifDownDecision (TOBDECISIONL);
  end;
end;

procedure TGestConsultFou.BeforeSortie;
begin
  if TheTOBCons <> nil then
  begin
    // le précédent est renseigné on verifie le contenu
    NettoieConsult (TheTobCons);
    if TheTOBCons.detail.Count = 0 then
    begin
      PositionneIndiceConsult (TheTOBCons,'AUC',0);
      FreeAndNil (TheTOBCOns);
    end;
    THgrid(XX.FindComponent('GS')).invalidate;
  end;
end;

procedure TGestConsultFou.fConsultWinClose (Sender : TObject);
begin
  BeforeSortie;
  BCONSULTFOU.Down := false;
end;

procedure TGestConsultFou.FermeConsFou(Sender: TObject);
begin
 fConsultWin.Visible := false;
 fConsultWinClose(self);
end;

function TGestConsultFou.FindFournisseur(TOBL: TOB; Valeur: string): boolean;
begin
  result := true;
  TOBL.putValue('BDF_TIERS',Valeur);
end;

procedure TGestConsultFou.GetCellCanvas(ACol, ARow: Integer; Canvas: TCanvas; AState: TGridDrawState);
begin

end;

procedure TGestConsultFou.GetComponents;
begin
  // Grille base
  GS := THgrid(XX.FindComponent('GSFOU'));
  if GS = nil then BEGIN Usable := false; Exit; END; 
  BDELCONSFOU := TToolbarButton97(XX.FindComponent ('BDELCONSFOU'));
  BSORTIECONSFOU := TToolbarButton97(XX.FindComponent ('BSORTIECONSFOU'));
  PCONSULTFOU := THPanel(XX.FindComponent ('PCONSULTFOU'));
  BCONSULTFOU := TToolbarButton97(XX.FindComponent ('BCONSULTFOU'));
  CHOIXFOURC := TpopupMenu(XX.FindComponent ('CHOIXFOURC'));
  MnCatalogueC := TmenuItem(XX.FindComponent ('MnCatalogueC'));
  MnFournisseurC := TmenuItem(XX.FindComponent ('MnFournisseurC'));
  POPM2 := TpopupMenu(XX.FindComponent ('POPM2'));
  Mndelete := TmenuItem(XX.FindComponent ('Mndelete'));
  MnSelection := TmenuItem(XX.FindComponent ('MnSelection'));

end;

function TGestConsultFou.GetConsulFou (TOBD : TOB) : TOB;
var Pos : integer;
begin
  result := nil;
  if (TOBD <> nil) and (TOBD.FieldExists('INDICECONSULT')) then
  begin
    Pos := TOBD.GetValue('INDICECONSULT');
    result := TOBConsultFou.findFirst(['INDICECONSULT'],[pos],true);
  end;
end;

procedure TGestConsultFou.GSCellEnter(Sender: TObject; var ACol,ARow: Integer; var Cancel: Boolean);
var TOBL : TOB;
begin
  ZoneSuivanteOuOk(GS,ACol, ARow, Cancel);
  if not Cancel then
  begin
    TOBL := TheTOBCons.detail[Arow-1];
    if TOBL <> nil then
    begin
      if (TOBL.GetValue('BDF_TIERS')='') and (Acol > G_FOURNISSEUR) then BEGIN Acol := G_FOURNISSEUR; PositionneDansGrid (Arow,Acol); END;
    end;
    GS.ElipsisButton := (Acol = G_FOURNISSEUR);
    stprev := GS.Cells [Acol,Arow];
  end;
end;

procedure TGestConsultFou.GSCellExit(Sender: TObject; var ACol,ARow: Integer; var Cancel: Boolean);
Var TOBLigne : TOB;
begin
  // pas de controle si remonté et ligne vide
  if (GS.row < Arow) and (IsLigneVide (Arow))  then exit;

  if (stPrev <> GS.cells[Acol,Arow]) or (GS.cells[Acol,Arow]= '') then
  begin
    TOBLigne := TheTOBCons.detail[Arow-1];
    if ACol = G_FOURNISSEUR then
    Begin
      if not FindFournisseur (TOBLigne,GS.cells[Acol,Arow]) then BEGIN cancel := true; exit; END;
    end else if Acol = G_REFARTTIERS then
    begin
      TOBLIgne.putValue('BDF_REFARTTIERS',GS.cells[Acol,Arow]);
    end else if Acol = G_PA then
    begin
      if GS.cells[Acol,Arow]='' then GS.cells[Acol,Arow] := '0';
      TOBLIgne.putValue('BDF_PAHT',Valeur(GS.cells[Acol,Arow]));
    end else if Acol = G_DELAI then
    begin
      if GS.cells[Acol,Arow]='' then GS.cells[Acol,Arow] := '0';
      TOBLIgne.putValue('BDF_NBJOUR',StrToInt(GS.cells[Acol,Arow]));
    end;

    if IsFournisseurExiste (Arow,Acol) Then
    BEGIN
      PgiBox('Ce fournisseur est déjà présent dans la saisie');
      cancel := true;
    END;
    AfficheLigne (GS,TOBLIGNE,Arow-1);

    if not cancel then
    begin
      stPrev := GS.cells[Acol,Arow];
    end;
  end;
end;

procedure TGestConsultFou.GSElipsisClick(Sender: TOBject);
var Coords : Trect;
    PointD,PointF : Tpoint;

begin
  if GS.col = G_FOURNISSEUR then
  begin
   Coords := GS.CellRect (GS.col,GS.row);
   PointD := GS.ClientToScreen ( Coords.Topleft)  ;
   PointF := GS.ClientToScreen ( Coords.BottomRight )  ;

   CHOIXFOURC.Popup (pointF.X ,pointD.y+10);
  end;
end;

procedure TGestConsultFou.GSEnter(Sender: TObject);
var Acol,Arow : integer;
    cancel : boolean;
begin
  cancel := false;
  GS.row := 1;
  GS.col := 1;
  Arow := 1;
  Acol := 1;
  ZoneSuivanteOuOk(GS,ACol, ARow, Cancel);
  cancel := false;
  GSRowEnter (GS,Arow,cancel,false);
  GSCellEnter (GS,Acol,Arow,cancel);
  GS.row := Arow;
  GS.col := Acol;
  GS.ShowEditor;
  GS.ElipsisButton := (Acol = G_FOURNISSEUR);
  stprev := GS.Cells [Acol,Arow];
end;

procedure TGestConsultFou.GSPostDrawCell(ACol, ARow: Integer; Canvas: TCanvas; AState: TGridDrawState);
begin

end;


procedure TGestConsultFou.GSKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
    case Key of
    VK_DELETE: if (Shift = [ssShift]) then
    begin
      DeleteLigne;
      Key := 0;
    end;
    VK_F5: if (Shift = [ssCtrl]) then
    begin
      if TFVierge(XX).ActiveControl = GS then
      begin
        GSElipsisClick(Sender);
        Key := 0;
      end;
    end;
    VK_RETURN : if (Shift = []) then
    begin
    	Key := 0;
      SendMessage(THedit(Sender).Handle, WM_KeyDown, VK_TAB, 0);
    end;
  end;
end;

procedure TGestConsultFou.GSRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
Var TobLigne : Tob;
begin
  if (Ou > 1) and (IsLigneVide (Ou-1)) then
  BEGIN
    cancel := true;
    Exit;
  END;
  if (Ou >= GS.rowCount -1) then GS.rowCount := GS.rowCount +1;
  if Ou > TheTOBCons.Detail.count then
  begin
    TOBLigne := AjouteLignecons;
    AfficheLigne (GS,TOBLigne,Ou-1);
  end;
end;

procedure TGestConsultFou.GSRowExit(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
var Acol : integer;
Begin
  if IsLigneVide (Ou) then exit;
  if not ControleLigne (Ou,Acol) then
  BEGIN
  	cancel := true;
    PositionneDansGrid (Ou,Acol);
  END;
end;

function TGestConsultFou.IsFournisseurExiste(Arow, Acol: integer): boolean;
begin
  //
  result := false;
end;

function TGestConsultFou.IsLigneVide (Arow : integer) : boolean;
var TOBL : TOB;
begin
  result := true;
  TOBL := TheTOBCons.detail[Arow-1];
  if TOBL.GetValue('BDF_TIERS') <> '' then result:= false;
end;

function TGestConsultFou.IsLigneVide (TOBL : TOB) : boolean;
begin
  result := true;
  if TOBL.GetValue('BDF_TIERS') <> '' then result:= false;
end;

procedure TGestConsultFou.libereTout;
begin
  fConsultWin.free;
  TOBConsultFou.free;
  TOBFourn.free;
end;


procedure TGestConsultFou.MnFournisseurClick (Sender : TObject);
var TOBRecupFou,LAlIG : TOB;
    Indice : integer;
    SQL : String;
    QQ : Tquery;
begin
  TOBRecupFou := TOB.Create ('LES FOURNISSEURS',nil,-1);
  TRY
    TheTOB := TOBrecupFou;
    AGLLanceFiche('BTP', 'BTFOURNMULTI_MUL', 'T_NATUREAUXI=FOU', '', 'SELECTION');
    if TheTOB = nil then exit;
    if TheTOB.detail.count = 0 then exit;
    for Indice := 0 to TheTOB.detail.count -1 do
    begin
      if TheTOBCons.findFirst(['BDF_TIERS'],[TheTOB.detail[Indice].getValue('T_TIERS')],True) = nil then
      begin
        // ajoute les lignes founisseurs
        if Indice = 0 Then
        begin
          LAlIG := TheTOBCons.detail[GS.row-1];
        end else
        begin
          LAlIG := AjouteLignecons;
        end;
        Sql := 'SELECT T_TIERS,T_LIBELLE FROM TIERS WHERE T_TIERS="'+TheTOB.detail[Indice].getValue('T_TIERS')+'" AND T_NATUREAUXI="FOU"';
        QQ := OPenSql (Sql,True);
        LAlIG.putValue('BDF_TIERS',QQ.findField('T_TIERS').AsString);
        LAlIG.putValue('T_LIBELLE',QQ.findField('T_LIBELLE').AsString);
        ferme (QQ);
      end;
    end;
  FINALLY
    TheTOB := nil;
    TOBRecupFou.free;
  END;
  AfficheLaGrille (TheTOBCons);
end;


procedure TGestConsultFou.RecupLesCataLogSelect (TheTOBparam : TOB);
var Indice : integer;
    TOBP,TOBC : TOB;
    Sql : String;
    QQ : Tquery;
begin
  for Indice := 0 to TheTObParam.detail.count-1 do
  begin
    TOBP := TheTOBparam.detail[Indice];
    if TheTOBCons.findFirst(['BDF_TIERS'],[TOBP.getValue('FOURNISSEUR')],True) = nil then
    begin
      // ajoute les lignes founisseurs
      if Indice = 0 Then
      begin
        TOBC := TheTOBCons.detail[GS.row-1];
      end else
      begin
        TOBC := AjouteLignecons;
      end;
      Sql := 'SELECT T_TIERS,T_LIBELLE FROM TIERS WHERE T_TIERS="'+TOBP.getValue('FOURNISSEUR')+'" AND T_NATUREAUXI="FOU"';
      QQ := OPenSql (Sql,True);
      TOBC.putValue('BDF_TIERS',QQ.findField('T_TIERS').AsString);
      TOBC.putValue('T_LIBELLE',QQ.findField('T_LIBELLE').AsString);
      TOBC.putValue('BDF_REFARTTIERS',TOBP.GetValue('REFERENCE'));
      TOBC.putValue('BDF_PAHT',TOBP.GetValue('PRIXACH'));
      TOBC.putValue('BDF_NBJOUR',TOBP.GetValue('DELAI'));
      ferme (QQ);
    end;
  end;
end;

procedure TGestConsultFou.MnCatalogueClick (Sender : TObject);
var TheTOBParam : TOB;
    TOBL : TOB;
begin
  TOBL := TheTOBCons.detail[GS.row-1];

  TheTOBparam := TOB.Create ('LES PARAMETRES',nil,-1);
  TheTOBParam.AddChampSupValeur ('FOURNISSEUR',TOBL.GEtValue('BDF_TIERS'));
  TheTOBParam.AddChampSupValeur ('ARTICLE',TheCurrentDecisionnel.GEtValue('BAD_ARTICLE'));
  TheTOBParam.AddChampSupValeur ('LIBELLEART',TheCurrentDecisionnel.GEtValue('BAD_LIBELLE'));
  TRY
  //
    MultipleLookUpFournisseur (TheTOBparam);
    if TheTOBparam.detail.count > 0 then
    begin
      RecupLesCataLogSelect (TheTOBparam);
      AfficheLaGrille (TheTOBCons);
    end;
  FINALLY
    TheTOBparam.free;
  END;
end;

procedure TGestConsultFou.PositionneDansGrid (Arow,Acol : integer);
var LastMode : boolean;
begin
  LastMode := assigned(GS.OnCellEnter);
  GS.CacheEdit;
  SetEventsGrid (false);
  GS.row := Arow;
  GS.col := Acol;
  stprev := GS.Cells [Acol,Arow];
  SetEventsGrid (LastMode);
  GS.row := Arow;
  GS.col := Acol;
  GS.ShowEditor;
  GS.ElipsisButton := (Acol = G_FOURNISSEUR);
end;

procedure TGestConsultFou.AjouteLesCatalogues (TheTOBCons : TOB);
var TOBcata,TOBLigne : TOB;
    Req : String;
    Indice : integer;
begin
  TOBcata := TOB.Create ('LES CATALOG',nil,-1);
  Req := 'SELECT *,T_LIBELLE FROM CATALOGU LEFT JOIN TIERS ON T_TIERS=GCA_TIERS AND T_NATUREAUXI="FOU" WHERE GCA_ARTICLE="'+
          TheCurrentDecisionnel.getValue('BAD_ARTICLE')+'"';
  TOBCata.LoadDetailDBFromSQL ('CATALOGU',req,false);
  for Indice := 0 to TOBCata.detail.count -1  do
  begin
    TOBLigne := AjouteLignecons;
    TOBLigne.putValue('BDF_TIERS',TOBCata.detail[Indice].GetValue('GCA_TIERS'));
    TOBLigne.putValue('T_LIBELLE',TOBCata.detail[Indice].GetValue('T_LIBELLE'));
    TOBLigne.putValue('BDF_REFARTTIERS',TOBCata.detail[Indice].GetValue('GCA_REFERENCE'));
    TOBLigne.putValue('BDF_COEFUAUS',TOBCata.detail[Indice].GetValue('GCA_COEFCONVQTEACH'));
    if TOBCata.detail[Indice].GetValue('GCA_DPA') <> 0 then
    begin
      TOBLigne.putValue('BDF_PAHT',TOBCata.detail[Indice].GetValue('GCA_DPA'));
    end else if TOBCata.detail[Indice].GetValue('GCA_PRIXVENTE') <> 0 then
    begin
      TOBLigne.putValue('BDF_PAHT',TOBCata.detail[Indice].GetValue('GCA_PRIXVENTE'));
    end;
    TOBLigne.putValue('BDF_NBJOUR',TOBCata.detail[Indice].GetValue('GCA_DELAILIVRAISON'));
  end;
  TOBcata.free;
end;

procedure TGestConsultFou.RemplitlaGrille;
var QQ : TQuery;
    Libelle : string;
    TOBLIGNE : TOB;
    FUA,FUS : double;
begin
  TheTOBCons := GetConsulFou (TheCurrentDecisionnel);
  if TheTOBCons = nil then
  begin
    FUS := RatioMesure('PIE', TheCurrentDecisionnel.GetValue('BAD_QUALIFQTESTO')); if FUS = 0 then FUS := 1;
    FUA := RatioMesure('PIE', TheCurrentDecisionnel.GetValue('BAD_QUALIFQTEACH')); if FUA = 0 then FUA := 1;
    TheTOBCons := AddFilleFromDecisionnel (TheCurrentDecisionnel);
    Inc(MaxIndice);
    PositionneIndiceConsult (TheTOBCONS,'ICI',MaxIndice);
    if TheCurrentDecisionnel.getValue('BAD_FOURNISSEUR')<> '' then
    begin
      // on recupere le fournisseur indiqué
      TOBLigne := AjouteLignecons;
      QQ := OpenSql ('SELECT T_LIBELLE FROM TIERS WHERE T_TIERS="'+ TheCurrentDecisionnel.getValue('BAD_FOURNISSEUR')+'" AND T_NATUREAUXI="FOU"',True);
      if not QQ.eof then Libelle := QQ.findField('T_LIBELLE').AsString
                    else Libelle := 'Inconnu';
      ferme (QQ);
      TOBLIGNE.putValue('BDF_TIERS',TheCurrentDecisionnel.getValue('BAD_FOURNISSEUR'));
      TOBLIGNE.putValue('BDF_COEFUAUS',TheCurrentDecisionnel.getValue('BAD_COEFUAUS'));
      TOBLIGNE.putValue('T_LIBELLE',Libelle);
      if TheCurrentDecisionnel.getValue('BAD_COEFUAUS') <> 0 then
      begin
      	TOBLIGNE.putValue('BDF_PAHT',Arrondi(TheCurrentDecisionnel.getValue('BAD_PRIXACHNET')*TheCurrentDecisionnel.getValue('BAD_COEFUAUS'),V_PGI.okdecP));
      end else
      begin
      	TOBLIGNE.putValue('BDF_PAHT',Arrondi(TheCurrentDecisionnel.getValue('BAD_PRIXACHNET')*FUS/FUA,V_PGI.okdecP));
      end;
    end;
    AjouteLesCatalogues (TheTOBCons);
  end;
  //
  AfficheLagrille (TheTOBCons);
  //
end;

procedure TGestConsultFou.ReindiceGrid;
var Indice : integer;
begin
  for Indice := 0 to TheTOBCons.detail.Count -1 do
  begin
    TheTOBCons.detail[Indice].putValue('BDF_NUMORDRE',Indice+1);
  end;
end;

procedure TGestConsultFou.SetBranche;
var Arow : integer;
begin
  TheCurrentDecisionnel := nil;
  Arow := TOF_DECISIONACHAT(TFVierge(XX).LaTOF).LaLigne;
  if  Arow < 1 then exit;
  TheCurrentDecisionnel := GetTOBDecisAch (TOBDecisionnel,TOF_DECISIONACHAT(TFVierge(XX).LaTOF).LaLigne);
end;

procedure TGestConsultFou.SetComponents;
begin
end;

procedure TGestConsultFou.SetEventsGrid(Actif: boolean);
begin
  if Actif then
  begin
    GS.OnEnter  := GSEnter;
    GS.OnRowEnter := GSRowEnter;
    GS.OnRowExit := GSRowExit;
    GS.OnCellEnter  := GSCellEnter;
    GS.OnCellExit  := GSCellExit;
    GS.OnElipsisClick := GSElipsisClick;
    GS.PostDrawCell := GSPostDrawCell;
    GS.GetCellCanvas := GetCellCanvas;
    GS.OnKeyDown := GSKeyDOwn;
    GS.PopupMenu  := POPM2;
  end
  else
  begin
    GS.OnEnter  := nil;
    GS.OnRowEnter := nil;
    GS.OnRowExit := nil;
    GS.OnCellEnter  := nil;
    GS.OnCellExit  := nil;
    GS.OnElipsisClick := nil;
    GS.PostDrawCell := nil;
    GS.GetCellCanvas := nil;
    GS.OnKeyDown := nil;
  end;
end;

procedure TGestConsultFou.SetNumDecisionnel(const Value: string);
var Req : string;
    TOBInterm : TOB;
begin
  if not Usable then exit; 
  TOBDecisionnel := TOF_DECISIONACHAT(TFVierge(XX).LaTOF).TheTOBdecisionnel;
  FnumDecisionnel := Value;
	TOBInterm := TOB.Create ('LE DETAIL',nil,-1);
  TRY
    //
    Req := 'SELECT DEC.*,T_LIBELLE FROM DECISIONACHLFOU DEC LEFT JOIN TIERS ON T_TIERS=BDF_TIERS AND T_NATUREAUXI="FOU" WHERE DEC.BDF_NUMERO='+
           FnumDecisionnel+
           ' ORDER BY DEC.BDF_NUMERO,DEC.BDF_N1,DEC.BDF_N2,DEC.BDF_N3,DEC.BDF_N4,DEC.BDF_NUMORDRE';
    //
    TOBINTERM.LoadDetailDBFromSQL ('DECISIONACHLFOU',Req,false);
    if TOBInterm.detail.count > 0 then
    begin
      ConstitueLaTOBConsult (TOBInterm);
      IndiceLaTOBCOnsult;
    end;
  FINALLY
  	TOBInterm.free;
  END;
end;

procedure TGestConsultFou.SetParamgrille;
var LaListe,lelement : string;
begin
  // récupération du paramétrage général des grilles
  FNomTable := 'DECISIONACHLFOU';
  FLien := '';
  FSortBy := '';
  Fparams := '';
  stcols := 'BDF_NUMORDRE;BDF_TIERS;T_LIBELLE;BDF_REFARTTIERS;BDF_PAHT;BDF_NBJOUR;BDF_SELECTION;';
  FTitre := 'N°;Fournisseur;Désignation;Référence F.;Prix;Délai;Retenu;';
  FLargeur := '2;10;16;16;8;5;5;';
  Falignement := 'G.0  ---;G.0  ---;G.0  ---;G.0  ---;D.2  -X-;D.0  -X-;C.0  ---;';
  title := 'Consultation fournisseur';
  NC := '1;1;1;1;1;1;1;';
  Fperso := '---';
  OkTri := True;
  OkNumCOl := false;
  //
  stColListe := stCols;
  nbcolsInListe := 0;
  LaListe := stColListe;
  repeat
    lelement := READTOKENST (laListe);

    if lelement <> '' then
    begin
      if lelement = 'BDF_NUMORDRE' then G_NUMORDRE := nbcolsinListe else
      if lelement = 'BDF_TIERS' then G_FOURNISSEUR := nbcolsinListe else
      if lelement = 'T_LIBELLE' then G_LIBELLE := nbcolsinListe else
      if lelement = 'BDF_REFARTTIERS' then G_REFARTTIERS := nbcolsinListe else
      if lelement = 'BDF_PAHT' then  G_PA := nbcolsinListe else
      if lelement = 'BDF_NBJOUR' then  G_DELAI := nbcolsinListe else
      if lelement = 'BDF_SELECTION' then  G_SELECT := nbcolsinListe ;
      inc(nbcolsInListe);
    end;
  until lelement = '';
end;

procedure TGestConsultFou.SetPrimalEvents;
begin
  BCONSULTFOU.OnClick := ConsultFouClick;
  fConsultWin.OnClose := fConsultWinClose;
  BSORTIECONSFOU.onclick := FermeConsFou;
  MnCatalogueC.onClick := MnCatalogueClick;
  MnFournisseurC.OnClick := MnFournisseurClick;
  MnSelection.onClick := SelectionneLeFourn;
  Mndelete.onClick := DeleteClick;
  BDELCONSFOU.OnClick := DelConsfouClick;
end;

function TGestConsultFou.VerifDownDecision (TOBDECISIONL : TOB) : boolean;
var fiston : TOB;
    Indice : integer;
begin
  result := false;
  if TOBDECISIONL.detail.count = 0 then exit;
  for Indice := 0 to TOBDECISIONL.detail.count -1 do
  begin
    fiston := TOBDECISIONL.detail[Indice];
    if GetConsulFou (fiston)<> nil then BEGIN result := true; break ; END;
    if VerifDownDecision (fiston) then BEGIN result := true; break ; END;
  end;
end;

function TGestConsultFou.VerifUpDecision (TOBDECISIONL : TOB) : boolean;
var Papounet : TOB;
begin
  result := false;
  Papounet := TOBDECISIONL.Parent;
  if Papounet = nil then exit;
  if GetConsulFou (papounet)<> nil then BEGIN result := true; Exit; END;
  if VerifUpDecision (papounet) then BEGIN result := true; Exit; END;
end;

function TGestConsultFou.ZoneAccessible( Grille : THgrid; var ACol, ARow : integer) : boolean;
var TOBL : TOB;
begin
  TOBL := nil;
  if Arow < TheTOBCons.detail.count then TOBL := TheTOBCons.detail[Arow-1];
  result := true;
  if Grille.ColWidths[acol] = 0 then BEGIN result := false; exit; END;
  if (Acol=G_FOURNISSEUR) then
  BEGIN
    if (TOBL <> nil) and (TOBL.GetValue('BDF_TIERS')<>'') then
    begin
      result := false;
      exit;
    end;
  END;
  if (Acol=G_SELECT) OR (Acol = G_LIBELLE) then
  BEGIN
    result := false;
    exit;
  END;
end;

procedure TGestConsultFou.ZoneSuivanteOuOk(Grille : THGrid;var ACol, ARow : integer; var Cancel : boolean);
var Sens, ii, Lim: integer;
  OldEna, ChgLig, ChgSens: boolean;
begin
  OldEna := Grille.SynEnabled;
  Grille.SynEnabled := False;
  Sens := -1;
  ChgLig := (Grille.Row <> ARow);
  ChgSens := False;
  if Grille.Row > ARow then Sens := 1 else if ((Grille.Row = ARow) and (ACol <= Grille.Col)) then Sens := 1;
  ACol := Grille.Col;
  ARow := Grille.Row;
  ii := 0;
  while not ZoneAccessible(Grille,ACol, ARow) do
  begin
    Cancel := True;
    inc(ii);
    if ii > 500 then Break;
    if Sens = 1 then
    begin
      // Modif BTP
      Lim := Grille.RowCount ;
      // ---
      if ((ACol = Grille.ColCount - 1) and (ARow >= Lim)) then
      begin
        if ChgSens then Break else
        begin
          // Ajout d'une ligne
          break;
        end;
      end;
      if ChgLig then
      begin
        ACol := Grille.FixedCols - 1;
        ChgLig := False;
      end;
      if ACol < Grille.ColCount - 1 then
      begin
        Inc(ACol);
      end else
      begin
        Inc(ARow);
        ACol := Grille.FixedCols;
      end;
    end else
    begin
      if ((ACol = Grille.FixedCols) and (ARow = 1)) then
      begin
        if ChgSens then Break else
        begin
          Sens := 1;
          Continue;
        end;
      end;
      if ChgLig then
      begin
        ACol := Grille.ColCount;
        ChgLig := False;
      end;
      if ACol > Grille.FixedCols then Dec(ACol) else
      begin
        Dec(ARow);
        ACol := Grille.ColCount - 1;
      end;
    end;
  end;
  Grille.SynEnabled := OldEna;
end;

procedure TGestConsultFou.RowChange;
begin
  if not Usable then exit;
  beforeSortie;
  ConsultFouClick (self);
end;

procedure TGestConsultFou.DeleteClick(Sender: TObject);
begin
  deleteLigne;
end;

procedure TGestConsultFou.ResizeToolWin(Sender: Tobject);
begin
  TFVierge(XX).HMTrad.ResizeGridColumns (GS);
end;

procedure TGestConsultFou.NettoieConsult (TheTobCons : TOB);
var Indice : integer;
    TOBL : TOB;
begin
  Indice := 0;
  repeat
    TOBL := TheTOBCOns.detail[Indice];
    if IsLigneVide(TOBL) then
    begin
      TOBL.Free
    end else Inc(Indice);
  until indice >= TheTOBCons.detail.count;
end;

procedure TGestConsultFou.PositionneIndiceConsult (TOBC : TOB; Code : string; Indice : integer; TraitementSurGrille : boolean = true);
var N1,N2,N3,N4 : integer;
    TOBD : TOB;
begin
  N1 := TOBC.GetValue('BDF_N1');
  N2 := TOBC.GetValue('BDF_N2');
  N3 := TOBC.GetValue('BDF_N3');
  N4 := TOBC.GetValue('BDF_N4');
  TOBD := TOBDecisionnel.findFirst(['BAD_NUMN1','BAD_NUMN2','BAD_NUMN3','BAD_NUMN4'],[N1,N2,N3,N4],true);
  if TOBD <> Nil then
  begin
    TOBC.putValue('INDICECONSULT',Indice);
    TOBD.putValue('INDICECONSULT',Indice);
    if Code = 'ICI' then
    begin
      TOBD.putValue('POSITIONCONSULT','ICI');
      SetConsultPere (TOBD,'APR');
      SetConsultFille (TOBD,'AVA');
    end else if Code = 'AUC' then
    begin
      TOBD.putValue('POSITIONCONSULT','AUC');
      SetConsultPere (TOBD,'AUC');
      SetConsultFille (TOBD,'AUC');
    end;
  end;
  if TraitementSurGrille then THgrid(XX.FindComponent('GS')).invalidate;
end;

procedure TGestConsultFou.PositionneTheConsult (TOBC : TOB);
var indice : integer;
begin
  if (TOBC.getValue('POSITIONCONSULT')='ICI') then
  BEGIN
    SetConsultFille (TOBC,'AVA');
    SetConsultPere (TOBC,'APR');
  END ELSE
  begin
    for Indice := 0 to TOBC.detail.count -1 do
    begin
      PositionneTheConsult (TOBC.detail[Indice]);
    end;
  end;
end;

procedure TGestConsultFou.IndiceLaTOBCOnsult;
var TOBC : TOB;
    Indice : integer;
begin
  MaxIndice := 0;
  for Indice := 0 to TOBconsultFou.detail.count -1 do
  begin
    TOBC := TOBconsultFou.detail[Indice];
    Inc(MaxIndice);
    PositionneIndiceConsult (TOBC,'ICI',MaxIndice);
  end;
end;

procedure TGestConsultFou.SetConsultPere (TOBDep : TOB;Code : string);
var Papounet : TOB;
begin
  papounet := TOBDep.Parent;
  if (Papounet <> nil) and (Papounet.FieldExists ('POSITIONCONSULT')) then
  begin
    papounet.putValue('POSITIONCONSULT',Code);
    SetConsultPere(papounet,Code);
  end;
end;

procedure TGestConsultFou.SetConsultFille (TOBDep : TOB;Code : string);
var fille : TOB;
    Indice : integer;
begin
  for Indice := 0 to TOBDep.detail.count -1 do
  begin
    fille :=  TOBDep.detail[Indice];
    fille.putValue('POSITIONCONSULT',Code);
    SetConsultFille(fille,Code);
  end;
end;


procedure TGestConsultFou.EnleveAutreSelection;
var Indice : integer;
begin
  for Indice := 0 to TheTOBCons.detail.count -1 do
  begin
    if Indice <> gs.row -1 then
    begin
      TheTOBCOns.detail[Indice].putValue('BDF_SELECTION','-');
    end;
  end;
end;

procedure TGestConsultFou.SelectionneLeFourn (Sender : TObject);
var TOBL : TOB;
    lastRow,Lastcol : integer;
begin
  lastRow := gs.row;
  LastCol := GS.col;
  TOBL := TheTOBCons.detail[Gs.row-1];
  if (TOBL.GetValue('BDF_TIERS')<>'') and (TOBL.GetValue('BDF_PAHT')<>0)  then
  begin
    if TOBL.GetValue('BDF_SELECTION') = '-' then
    begin
      if PgiAsk ('Confirmez-vous la sélection de ce fournisseur ?') <> MrYes then exit;
      EnleveAutreSelection;
      TOBL.putValue('BDF_SELECTION','X');
      AfficheLagrille (TheTObCOns);
      PositionneDansGrid (LastRow,LastCol);
      //
      // Positionne maintenant le fournisseur ensuite le prix net dans la grille du decisionnel
      //
      TOF_DECISIONACHAT(TFvierge(XX).LaTOF).TheSelectFournisseur := TOBL.GetValue('BDF_TIERS');
      TOF_DECISIONACHAT(TFvierge(XX).LaTOF).TheSelectPrix  := TOBL.GetValue('BDF_PAHT');
    end else
    begin
      if TOBL.GetValue('BDF_SELECTION') = '-' then TOBL.putValue('BDF_SELECTION','X')
                                              else TOBL.putValue('BDF_SELECTION','-');
      AfficheLagrille (TheTObCOns);
      PositionneDansGrid (LastRow,LastCol);
    end;
  end;
end;

procedure TGestConsultFou.DelConsfouClick(Sender: TObject);
begin
  PositionneIndiceConsult (TheTOBCons,'AUC',0);
  FreeAndNil (TheTOBCOns);
  BCONSULTFOU.down := false;
  fConsultWin.Visible := false;
end;

procedure TGestConsultFou.Renumerote(TOBD : TOB);
var TOBC : TOB;
begin
  if not Usable then exit;
  if not TOBD.FieldExists('INDICECONSULT') then exit;
  if TOBD.GetValue('INDICECONSULT') > 0 then
  begin
    TOBC := TOBConsultFou.FindFirst (['INDICECONSULT'],[TOBD.getValue('INDICECONSULT')],True);
    if TOBC <> nil then
    begin
      TOBC.putValue('BDF_NUMERO',TOBD.GetValue('BAD_NUMERO'));
      TOBC.putValue('BDF_N1',TOBD.GetValue('BAD_NUMN1'));
      TOBC.putValue('BDF_N2',TOBD.GetValue('BAD_NUMN2'));
      TOBC.putValue('BDF_N3',TOBD.GetValue('BAD_NUMN3'));
      TOBC.putValue('BDF_N4',TOBD.GetValue('BAD_NUMN4'));
      if TOBC.detail.count > 0 then
      begin
        RenumeroteFille (TOBD,TOBC);
      end;
    end;
  end;
end;

procedure TGestConsultFou.RenumeroteFille (TOBD,TOBC : TOB);
var Indice : integer;
    TOBF : TOB;
begin
  for Indice := 0 to TOBC.detail.count -1 do
  begin
    TOBF := TOBC.detail[Indice];
    TOBF.putValue('BDF_NUMERO',TOBD.GetValue('BAD_NUMERO'));
    TOBF.putValue('BDF_N1',TOBD.GetValue('BAD_NUMN1'));
    TOBF.putValue('BDF_N2',TOBD.GetValue('BAD_NUMN2'));
    TOBF.putValue('BDF_N3',TOBD.GetValue('BAD_NUMN3'));
    TOBF.putValue('BDF_N4',TOBD.GetValue('BAD_NUMN4'));
  end;
end;

procedure TGestConsultFou.Supprime;
var Sql : string;
begin
  if not USable then exit;
  Sql := 'DELETE FROM DECISIONACHLFOU WHERE BDF_NUMERO="'+FnumDecisionnel+'"';
  ExecuteSql (Sql);
end;

procedure TGestConsultFou.ecrit;
begin
  if not Usable then exit;
	TOBConsultFou.SetAllModifie (true);
  if not TOBConsultFou.InsertDBByNivel (true) then V_PGI.IOError := OeUnknown;
end;

function TGestConsultFou.IsExisteConsult: boolean;
begin
  if not Usable then BEGIN result := false; exit; END;
  result := (TOBConsultFou.detail.count > 0);
end;

procedure TGestConsultFou.PrepareEdition;

  procedure DefinilaListe (TOBliste,TOBD : TOB; CurrentIndice : integer);
  var TOBL,TOBC : TOB;
      Indice : integer;
      Niv : string;
      Fournisseur,Destination : string;
  begin
    //
    Niv := TOBD.getValue('BAD_TYPEL');
    if Niv = 'ART' then
    begin
      if (TOBD.getValue('BAD_LIVCHANTIER')='X') and (TOBD.getValue('BAD_AFFAIRE')<>'') then
      begin
        Destination := 'CHA;'+TOBD.getValue('BAD_AFFAIRE')+';';
      end else
      begin
        if TOBD.getValue('BAD_DEPOT') <> GetParamsocSecur('SO_GCDEPOTDEFAUT','') then
        begin
          Destination := 'DEP;'+TOBD.getValue('BAD_DEPOT')+';';
        end else
        begin
          Destination := '';
        end;
      end;
    end else
    begin
      Destination := '';
    end;
    //
    TOBC := GetConsulFou (CurrentIndice); if TOBC = nil then exit;
    for Indice := 0 TO TOBC.detail.count -1 do
    begin
      TOBL := TOB.create ('DECISIONACHLIG',TOBListe,-1);
      TOBL.Dupliquer (TOBD,false,true);
      Fournisseur := TOBC.detail[Indice].GetValue('BDF_TIERS');
      SetAdresseFou (TOBL,Fournisseur);
      //
      TOBL.PutValue ('BAD_FOURNISSEUR',fournisseur);
      TOBL.AddChampSupValeur ('FOURNISSEUR',fournisseur);
      TOBL.AddChampSupValeur ('DESTINATION',destination);
    end;
  end;


  procedure RecupConsultFou (TOBD,TOBListe : TOB; MemoIndice : integer; ApplicAllLvl : boolean);
  var CurrentIndice : integer;
      Niv : string;
      Indice : integer;
  begin
    Niv := TOBD.getValue('BAD_TYPEL');
    CurrentIndice := 0;
    if (not ApplicAllLvl) then
    begin
      CurrentIndice := TOBD.GetValue('INDICECONSULT');
      if CurrentIndice > 0 then
      begin
        DefinilaListe (TOBliste,TOBD,CurrentIndice);
      end else
      begin
        if Niv= 'ART' Then exit;
        for Indice:=0 to TOBD.detail.count -1 do
        begin
          RecupConsultFou (TOBD.detail[Indice],TOBListe,0,ApplicAllLvl);
        end;
      end;
    end;

    if (ApplicAllLvl)  then
    begin
      if memoIndice > 0 then
      begin
        CurrentIndice := MemoIndice
      end else
      begin
        if (TOBD.GetValue('INDICECONSULT') > 0) then
        begin
          CurrentIndice := TOBD.GetValue('INDICECONSULT');
        end;
      end;
      //
      if Niv = 'ART' then
      begin
        if CurrentIndice > 0 then DefinilaListe (TOBliste,TOBD,CurrentIndice);
      end else
      begin
        for Indice:=0 to  TOBD.detail.count -1 do
        begin
          RecupConsultFou (TOBD.detail[Indice],TOBListe,CurrentIndice,ApplicAllLvl);
        end;
      end;
    end;
  end;

  procedure SetAdresseLivChantier (TOBAdresses : TOB; MemoTemp : TStringList);
  begin
    if GetParamSoc('SO_GCPIECEADRESSE') then
    BEGIN
      if (TobAdresses.detail[0].getvalue('GPA_libelle') = '') and
         (TobAdresses.detail[0].getvalue('GPA_libelle2') = '') then exit;
      MemoTemp.add (TraduireMemoire('Livraison à l''adresse suivante'));
      if TobAdresses.detail[0].getvalue('GPA_LIBELLE') <> '' then
        MemoTemp.add (TobAdresses.detail[0].getvalue('GPA_LIBELLE'));
      if TobAdresses.detail[0].getvalue('GPA_LIBELLE2') <> '' then
        MemoTemp.add (TobAdresses.detail[0].getvalue('GPA_LIBELLE2'));
      if TobAdresses.detail[0].getvalue('GPA_ADRESSE1') <> '' then
        MemoTemp.add (TobAdresses.detail[0].getvalue('GPA_ADRESSE1'));
      if TobAdresses.detail[0].getvalue('GPA_ADRESSE2') <> '' then
        MemoTemp.add (TobAdresses.detail[0].getvalue('GPA_ADRESSE2'));
      if TobAdresses.detail[0].getvalue('GPA_ADRESSE3') <> '' then
        MemoTemp.add (TobAdresses.detail[0].getvalue('GPA_ADRESSE3'));
      if TobAdresses.detail[0].getvalue('GPA_CODEPOSTAL') <> '' then
        MemoTemp.add (TobAdresses.detail[0].getvalue('GPA_CODEPOSTAL')+ ' '+
                      TobAdresses.detail[0].getvalue('GPA_VILLE'));

    END ELSE
    BEGIN
      if (TobAdresses.detail[0].getvalue('ADR_libelle') = '') and
          (TobAdresses.detail[0].getvalue('ADR_libelle2') = '') then exit;
       MemoTemp.add (TraduireMemoire('Livraison à l''adresse suivante'));
      if TobAdresses.detail[0].getvalue('ADR_libelle') <> '' then
        MemoTemp.add (TobAdresses.detail[0].getvalue('ADR_libelle'));
      if TobAdresses.detail[0].getvalue('ADR_libelle2') <> '' then
        MemoTemp.add (TobAdresses.detail[0].getvalue('ADR_libelle2'));
      if TobAdresses.detail[0].getvalue('ADR_ADRESSE1') <> '' then
        MemoTemp.add (TobAdresses.detail[0].getvalue('ADR_ADRESSE1'));
      if TobAdresses.detail[0].getvalue('ADR_ADRESSE2') <> '' then
        MemoTemp.add (TobAdresses.detail[0].getvalue('ADR_ADRESSE2'));
      if TobAdresses.detail[0].getvalue('ADR_ADRESSE3') <> '' then
        MemoTemp.add (TobAdresses.detail[0].getvalue('ADR_ADRESSE3'));
      if TobAdresses.detail[0].getvalue('ADR_CODEPOSTAL') <> '' then
        MemoTemp.add (TobAdresses.detail[0].getvalue('ADR_CODEPOSTAL')+' '+
                      TobAdresses.detail[0].getvalue('ADR_VILLE'));
    END;
  end;

  procedure AddDestination (TheTOBEdition,TOBP : TOB);
  var TOBL,TOBS,TOBPiece,TOBAdresses : TOB;
      First,Second,Chaine : string;
      MemoTemp : TStringList;
      Indice : integer;
  begin
    if TOBP.GetValue('DESTINATION')<>'' then
    begin
      Chaine := TOBP.GetValue('DESTINATION');
      First := READTOKENST (Chaine);
      Second := READTOKENST (Chaine);
      if (First = 'DEP') then
      begin
        MemoTemp := TStringList.create;
        MemoTemp.text := '';

        TOBS := TOB.Create ('LIGNE',nil,-1);
        TOBL.dupliquer (TOBP,false,true);
        TOBS.putvalue('GL_DEPOT',Second);
        getAdresseSoc (TOBS,MemoTemp);
        for Indice := 0 to MemoTemp.count -1 do
        begin
          TOBL := TOB.Create ('DECISIONACHLIG',TheTOBEdition,-1);
          TOBL.PutValue('BAD_TYPEL','COM');
          TOBL.PutValue('BAD_LIBELLE',MemoTemp.Strings [Indice]);
        end;
        //
        TOBS.Free;
        MemoTemp.Free;
      end else
      if first = 'CHA' then
      begin
        MemoTemp := TStringList.create;
        MemoTemp.text := '';
        TOBPiece := TOB.Create ('PIECE',nil,-1);
        TOBAdresses := TOB.Create ('LES ADRESSES',nil,-1);
        if GetParamSoc('SO_GCPIECEADRESSE') then
        begin
          TOB.Create('PIECEADRESSE', TOBAdresses, -1); {Livraison}
          TOB.Create('PIECEADRESSE', TOBAdresses, -1); {Facturation}
        end else
        begin
          TOB.Create('ADRESSES', TOBAdresses, -1); {Livraison}
          TOB.Create('ADRESSES', TOBAdresses, -1); {Facturation}
        end;
        TOBPiece.putValue('GP_AFFAIRE',TOBP.GetValue('BAD_AFFAIRE'));
        LivAffaireVersAdresses (nil,TOBAdresses,TOBPiece,false);
        SetAdresseLivChantier (TOBAdresses,MemoTemp);
        for Indice := 0 to MemoTemp.count -1 do
        begin
          TOBL := TOB.Create ('DECISIONACHLIG',TheTOBEdition,-1);
          TOBL.dupliquer (TOBP,false,true);
          TOBL.PutValue('BAD_TYPEL','COM');
          TOBL.PutValue('BAD_LIBELLE',MemoTemp.Strings [Indice]);
        end;
        //
        Memotemp.free;
        TOBPiece.free;
        TOBAdresses.free;
      end;
    end;
  end;

var Indice,IndC : integer;
    TOBD,TOBL,TOBparam,TOBListe,TheTobEdition : TOB;
    TOBprepare,TOBEdition : TOB;
    ApplicAllLvl,SautPageDest,ImpGlobale : boolean;
    Fournisseur,Destination : string;
    Error : boolean;
begin
  if not Usable then exit;
  // initialisation
  TOBEdition := TOB.Create ('LES EDITIONS',nil,-1);
  TOBPrepare := TOB.Create ('LA LISTE ORDONNEE',nil,-1);
  TOBListe := TOB.Create ('LA LISTE ',nil,-1);
  TOBParam := TOB.create ('LES PARAMS',nil,-1);
  TOBParam.AddChampSupValeur ('ALLLEVEL','-');
  TOBParam.AddChampSupValeur ('SAUTPAGEDEST','-');
  TOBParam.AddChampSupValeur ('IMPGLOBALE','X');
  // demande d'infos
  TheTOB := TOBParam;
  AGLLanceFiche ('BTP','BTPAREDCONS','','','');
  if TheTOB <> nil then
  begin
    TheTOB := nil;
    ApplicAllLvl := (TOBParam.getValue('ALLLEVEL')='X');
    SautPageDest := (TOBParam.getValue('SAUTPAGEDEST')='X');
    Impglobale := (TOBParam.getValue('IMPGLOBALE')='X');
    // 1ere partie --
    // on va trier la tob decisionnel par fournisseur / destination marchandise
    for Indice := 0 to TOBDecisionnel.detail.count -1 do
    begin
      TOBD := TOBdecisionnel.detail[Indice];
      TOBListe.ClearDetail;
      RecupConsultFou (TOBD,TOBListe,0,ApplicAllLvl);
      if TOBListe.detail.count > 0 then
      begin
        IndC := 0;
        repeat
          TOBL := TOBListe.detail[Indc];
          TOBL.changeParent (TOBPrepare,-1);
        until Indc >= TOBListe.detail.Count;
      end;
    end;
    if TOBPrepare.detail.count > 0 then
    begin
      TOBPrepare.detail.sort ('FOURNISSEUR;DESTINATION;BAD_ARTICLE');
      // 2eme partie  Constitution des editions en fonctions du paramétrage
      Destination := '';
      Fournisseur := '';
      TheTobEdition := nil;
      Indice := 0;
      repeat
        if Fournisseur <> TOBPrepare.detail[Indice].getValue ('FOURNISSEUR') then
        begin
          TheTOBEdition := TOB.Create ('UNE NOUVELLE EDITION',TOBEdition,-1);
          Fournisseur := TOBPrepare.detail[Indice].getValue ('FOURNISSEUR');
          if SautPageDest then
          begin
            Destination := TOBPrepare.detail[Indice].getValue ('DESTINATION'); // pour eviter un saut de page suplémentaire
            AddDestination (TheTOBEdition,TOBPrepare.detail[Indice]);
          end else
          begin
            Destination := '';
          end;
        end;
        if (Destination <> TOBPrepare.detail[Indice].getValue ('DESTINATION')) and (SautPageDest) then
        begin
          Destination := TOBPrepare.detail[Indice].getValue ('DESTINATION');
          TheTOBEdition := TOB.Create ('UNE NOUVELLE EDITION',TOBEdition,-1);
          AddDestination (TheTOBEdition,TOBPrepare.detail[Indice]);
        end else
        if (Destination <> TOBPrepare.detail[Indice].getValue ('DESTINATION')) and (not SautPageDest) then
        begin
          AddDestination (TheTOBEdition,TOBPrepare.detail[Indice]);
          Destination := TOBPrepare.detail[Indice].getValue ('DESTINATION');
        end;
        TOBPrepare.detail[Indice].ChangeParent (TheTOBEdition,-1);
      Until Indice >= TOBPrepare.detail.count;
      // Editions
      if ImpGlobale then StartPdfBatch;
      Error := false;
      for Indice := 0 to TOBEdition.detail.count -1 do
      begin
        if LanceEtatTob ('E','GPJ',fCodeEtat,TobEdition.detail[Indice],true,false,false,nil,'','Edition demande de prix',false) < 0 then
        begin
          exit;
        end;
      end;
      if (Impglobale) and (not Error) then
      begin
        CancelPDFBatch ;
        PreviewPDFFile('',GetMultiPDFPath,True);
      end;
    end;
  end;
  // libération
  TOBPrepare.free;
  TOBListe.free;
  TOBParam.free;
  TOBEdition.free;
end;

function TGestConsultFou.GetConsulFou(Indice: integer): TOB;
begin
  result := nil;
  if Indice = 0 then exit;
  result := TOBConsultFou.findFirst(['INDICECONSULT'],[Indice],true);
end;

procedure TGestConsultFou.SetEtat(LeCodeEtat: string);
begin
  if not usable then exit;
  fCodeEtat := LeCodeEtat;
end;

procedure TGestConsultFou.SetAdresseFou (TOBL:TOB;Fournisseur : string);
var TOBF : TOB;
    QQ : TQuery;
begin
  TOBF := TOBFourn.FindFirst(['T_TIERS'],[Fournisseur],true);
  if TOBF = nil then
  begin
    TOBF := TOB.Create ('TIERS',TOBFourn,-1);
    QQ := OPenSql ('SELECT T_TIERS,T_FORMEJURIDIQUE,T_LIBELLE,T_ADRESSE1,T_ADRESSE2,T_ADRESSE2,T_CODEPOSTAL,T_VILLE,T_FAX FROM TIERS WHERE T_TIERS="'+Fournisseur+'" AND T_NATUREAUXI="FOU"',True);
    if not QQ.eof then
    begin
      TOBF.SelectDB ('',QQ);
    end;
  end;
  TOBL.AddChampSupValeur ('T_FORMEJURIDIQUE',TOBF.GetValue('T_FORMEJURIDIQUE'));
  TOBL.AddChampSupValeur ('T_LIBELLE',TOBF.GetValue('T_LIBELLE'));
  TOBL.AddChampSupValeur ('T_ADRESSE1',TOBF.GetValue('T_ADRESSE1'));
  TOBL.AddChampSupValeur ('T_ADRESSE2',TOBF.GetValue('T_ADRESSE2'));
  TOBL.AddChampSupValeur ('T_ADRESSE3',TOBF.GetValue('T_ADRESSE3'));
  TOBL.AddChampSupValeur ('T_CODEPOSTAL',TOBF.GetValue('T_CODEPOSTAL'));
  TOBL.AddChampSupValeur ('T_VILLE',TOBF.GetValue('T_VILLE'));
  TOBL.AddChampSupValeur ('T_FAX',TOBF.GetValue('T_FAX'));
end;

procedure TGestConsultFou.SetUnusable(const Value: boolean);
begin
  Usable := false;
end;

procedure TGestConsultFou.AffecteDecisionnel (TOBD,TOBF : TOB);
var TOBDD,TOBCC : TOB;
    Indice : integer;
begin
  TOBCC := GetConsulFou ( TOBD); if TOBCC <> nil then DeleteConsulFouDetail (TOBD); // sur le niveau courant
  if ExistConsulFouDetail (TOBD) then DeleteConsulFouDetail (TOBD);
  TheTOBCons := AddFilleFromDecisionnel (TOBD);  // positionne la branche
  for Indice := 0 to TOBF.detail.count - 1 do
  begin
    TOBDD := AjouteLignecons;
    TOBDD.putValue('BDF_TIERS',TOBF.detail[Indice].GetValue('FOURNISSEUR'));
    TOBDD.putValue('T_LIBELLE',TOBF.detail[Indice].GetValue('LIBELLE'));
  end;
  Inc(MaxIndice);
  PositionneIndiceConsult (TheTOBCONS,'ICI',MaxIndice,false);
end;

procedure TGestConsultFou.MultipleAffecte(TOBFournisseurs: TOB);
var Indice : integer;
    TOBD : TOB;
begin
  // provenance de la multi selection
  for Indice := 0 to TOBdecisionnel.detail.count -1 do
  begin
    TOBD:= TOF_DECISIONACHAT(TFVierge(XX)).FindSelected (TOBDecisionnel.detail[Indice]);
    if TOBD <> nil then
    begin
      AffecteDecisionnel (TOBD,TOBFournisseurs);
    end;
  end;

end;

end.
