unit UTransferts;

interface
uses sysutils,classes,windows,messages,controls,forms,hmsgbox,stdCtrls,clipbrd,nomenUtil,
     SaisUtil,Ent1,EntGC,UtilPGI,UTOB,HTB97,FactUtil,FactComm,Menus,ParamSoc,ExtCtrls,
     ComCtrls,
     HCtrls,
     Hpanel,
     HEnt1,
     AglInit,FactTob,FactVariante,vierge,UtilNumParag,uEntCommun,
{$IFDEF EAGLCLIENT}
     maineagl,
{$ELSE}
     fe_main,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ENDIF}
     HRichOLE,UTOF;

const MAXITEMS = 6;

type

  TGestTransfert = class (TObject)
  private
    fCreatedPop : boolean;
    FF : TForm;
    TTWTransfert : TToolWindow97;
    TVTRansfert : TTreeView;
    TOBTransferts : TOB;
    fusable : boolean;
    POPGS : TPopupMenu;
    MesMenuItem: array[0..MAXITEMS] of TMenuItem;
    fMaxItems : integer;
    procedure DefiniMenuPop(Parent: Tform);
    procedure AjoutTransfert (Sender : TObject);
    procedure DetailTransfert (Sender : TObject);
    procedure ModifTransfert (Sender : TObject);
    function IsTransfert (TOBL : TOB) : boolean;
    procedure ActiveMenu (Etat : boolean);
    procedure POPYTSClick (Sender : TObject);
    function ConstitueTOBDetailTransfert(TOBL,TOBTRFPOC: TOB): boolean;
    procedure ConstitueTreeViewTrf;
    function IsExistsTransfert(TOBL: TOB): boolean;
    procedure TreeViewClick (Sender : TObject);
  public
    property CurrentSaisie : TForm read FF;
    destructor  destroy ; override;
    constructor create(TT: TForm);
    procedure DefiniMenu (TOBL : TOB);
    procedure SetPiece (NaturePiece : string);
  end;

var CurrentTransfert : TGestTransfert;

function FindTransfert (TOBTRFPOC,TOBL : TOB) : TOB;
procedure LoadLesTOBTRF(CleDoc : r_cledoc;TOBPiece,TOBTRFPOC,TOBOuvrage : TOB);
procedure ValideLesTOBTRF (TOBPiece,TOBTRFPOC : TOB);
procedure DeleteLesTOBTRF (Cledoc :r_cledoc);
function LastTransfert (TOBTRFPOC,TOBT : TOB) : boolean;
function FindLigneDocFromDetail(TOBPiece,TOBD: TOB): TOB;

implementation
uses Facture,UtilTOBPiece, Variants,FactOuvrage,FactureBTP;


function ExisteTransfertSuiv (TOBTRFPOC,TT : TOB) : boolean;
var TS : TOB;
begin
  result := false;
  TS := TOBTRFPOC.findfirst(['BT3_NUMORDRE','BT3_UNIQUEBLO'],[TT.GetInteger('BT3_NUMORDRE'),TT.GetInteger('BT3_UNIQUEBLO')],true);
  repeat
    if TS = nil then break;
    if TS.GetInteger('BT3_UNIQUE') > TT.GetInteger('BT3_UNIQUE') then
    begin
      result := True;
      Exit;
    end;
    TS := TOBTRFPOC.findnext(['BT3_NUMORDRE','BT3_UNIQUEBLO'],[TT.GetInteger('BT3_NUMORDRE'),TT.GetInteger('BT3_UNIQUEBLO')],true);
  until TS = nil;
end;

function LastTransfert (TOBTRFPOC,TOBT : TOB) : boolean;
var II : integer;
    TT : TOB;
begin
  result := true;
  for II := 0 to TOBT.detail.count -1 do
  begin
    TT := TOBT.detail[II];
    if ExisteTransfertSuiv(TOBTRFPOC,TT) then
    begin
      result := false;
      exit;
    end;
  end;
end;


procedure DeleteLesTOBTRF (Cledoc :r_cledoc);
begin
  ExecuteSQL('DELETE FROM BTRFENTETE WHERE '+WherePiece(Cledoc,ttdTRFEntPOC,true ));
  ExecuteSQL('DELETE FROM BTRFDETAIL WHERE '+WherePiece(CleDoc,ttdTRFDetPOC,true));
end;


function FindTransfert (TOBTRFPOC,TOBL : TOB) : TOB;
begin
  Result := TOBTRFPOC.findfirst(['BT2_UNIQUE'],[TOBL.GetInteger('NUMTRANSFERT')],True);
end;

procedure LoadLesTOBTRF(CleDoc : r_cledoc;TOBPiece,TOBTRFPOC,TOBOUvrage : TOB);

  function FindLigneInDoc (TT : TOB) : TOB;
  var II : Integer;
  begin
    Result := nil;
    if TT.GetInteger('BT3_NUMORDRE')<> 0 then
    begin
      for II := 0 to TOBPiece.detail.count -1 do
      begin
        if TOBPiece.detail[II].GetInteger('GL_NUMORDRE')= TT.GetInteger('BT3_NUMORDRE') then
        begin
          result := TOBPiece.detail[II];
          break;
        end;
      end;
    end;
  end;

var QQ : TQuery;
    TOBT,TT,ThePere,TOBL : TOB;
    LastUnique : integer;
    II : Integer;
begin
  LastUnique := -1;
  ThePere := nil;
  TOBT := TOB.Create ('LES LIGNES TRF',nil,-1);
  TRY
    QQ := OpenSQL('SELECT * FROM BTRFENTETE WHERE '+WherePiece(CleDoc,ttdTRFEntPoc,true)+' ORDER BY BT2_UNIQUE',True,-1,'',True);
    if not QQ.Eof then
    begin
      TOBTRFPOC.LoadDetailDB('BTRFENTETE','','',QQ,False);
    end;
    ferme (QQ);
    //
    QQ := OpenSQL('SELECT * FROM BTRFDETAIL WHERE '+WherePiece(CleDoc,ttdTRFDetPOC,true)+' ORDER BY BT3_UNIQUE,BT3_TYPELIGNETRF,BT3_NUMORDRE',True,-1,'',True);
    if not QQ.Eof then
    begin
      TOBT.LoadDetailDB('BTRFDETAIL','','',QQ,False);
      II := 0;
      repeat
        TT := TOBT.detail[II];
        if LastUnique <> TT.GetInteger('BT3_UNIQUE') then
        begin
          ThePere := TOBTRFPOC.FindFirst(['BT2_UNIQUE'],[TT.GetInteger('BT3_UNIQUE')],true);
          if ThePere = nil then
          begin
            TT.free;
            continue;
          end;
          LastUnique := TT.GetInteger('BT3_UNIQUE');
        end;
        if ThePere <> nil then
        begin
          TT.ChangeParent(ThePere,-1);
          TOBL := FindLigneInDoc (TT);
          if TOBL <> nil then TT.Data := TOBL;
        end;
      until II >= TOBT.detail.Count;
    end;
    ferme (QQ);
  finally
    TOBT.free;
  end;
end;


procedure ValideLesTOBTRF (TOBPiece,TOBTRFPOC : TOB);
var II,JJ : Integer;
    TOBT,TT : TOB;
begin
  for II := 0 to TOBTRFPOC.detail.count -1 do
  begin
    TOBT := TOBTRFPOC.detail[II];
    TOBT.SetString('BT2_NATUREPIECEG',TOBPiece.GetString('GP_NATUREPIECEG') );
    TOBT.SetString('BT2_SOUCHE',TOBPiece.GetString('GP_SOUCHE') );
    TOBT.SetInteger('BT2_NUMERO',TOBPiece.GetInteger('GP_NUMERO') );
    TOBT.SetInteger('BT2_INDICEG',TOBPiece.GetInteger('GP_INDICEG') );
    TOBT.SetAllModifie(true);
    for JJ := 0 to TOBT.detail.count -1 do
    begin
      TT := TOBT.detail[JJ];
      TT.SetString('BT3_NATUREPIECEG',TOBPiece.GetString('GP_NATUREPIECEG') );
      TT.SetString('BT3_SOUCHE',TOBPiece.GetString('GP_SOUCHE') );
      TT.SetInteger('BT3_NUMERO',TOBPiece.GetInteger('GP_NUMERO') );
      TT.SetInteger('BT3_INDICEG',TOBPiece.GetInteger('GP_INDICEG') );
      TT.SetInteger('BT3_UNIQUE',TOBT.GetInteger('BT2_UNIQUE'));
      TT.SetAllModifie(true);
    end;
  end;
  TOBTRFPOC.InsertDBByNivel(false); 
end;


{ TGestTransfert }

function FindLigneDocFromDetail(TOBPiece,TOBD: TOB): TOB;
var Inddep,Indice : integer;
		indOuvrage : integer;
begin
	result := nil;
  Inddep := TOBD.GetIndex ;
  IndOuvrage := TOBD.GetValue('GL_INDICENOMEN');
  for Indice := Inddep-1 downto 0 do
  begin
    if (ISOuvrage(TOBpiece.detail[Indice])) and
    	 (TOBpiece.detail[Indice].getValue('GL_INDICENOMEN')=IndOuvrage) and (IsArticle(TOBpiece.detail[Indice])) then
    begin
    	result := TOBpiece.detail[Indice];
      break;
    end;
  end;
end;

procedure TGestTransfert.ActiveMenu(Etat: boolean);
var II : integer;
begin
  for II := 0 to fMaxItems -1 do
  begin
    MesMenuItem[II].visible := Etat;
  end;
end;

procedure TGestTransfert.AjoutTransfert(Sender: TObject);
var TOBT,LastTOBL : TOB;
    TOBPiece : TOB;
    TOBTRFPOC : TOB;
    Arow,Acol : Integer;

  function TransfertOnlyOuvrage (TOBL : TOB) : boolean;
  var TOBOuvrages,TOBOUV : TOB;
      IndiceNomen,II : integer;
  begin
    Result := false;
    TOBOuvrages := TFFACture(FF).TheTOBOuvrage;
    IndiceNomen := TOBL.GetInteger('GL_INDICENOMEN'); if IndiceNomen = 0 then exit;
    TOBOUV := TOBOuvrages.detail[IndiceNomen-1]; if TOBOUV = nil then exit;
    for II := 0 to TOBOUV.Detail.count -1 do
    begin
      if TOBOUV.detail[II].GetDouble('MTTRANSFERT') <> 0 then Exit;
    end;
    Result := True;
  end;

  function BeforeTransfert (TOBPiece,TOBTRFPOC,TOBT : TOB) : boolean;
  var II : Integer;
      GS : Thgrid;
      TOBL,TL,TOBLIgne,TOBOuvrages,TOBO  : TOB;
      MontantPa : Double;
  begin
    GS := TFFActure(FF).GS;
    TOBOuvrages := TFFACture(FF).TheTOBOuvrage;
    Acol := GS.Col;
    Result := false;
    if GS.nbSelected = 0 then BEGIN PGIInfo('Aucune ligne séléctionné'); Exit; END;
    for II := 1 to GS.RowCount do
    begin
      TOBL := GetTOBLigne(TOBPiece,II);
      if not GS.IsSelected(II) then Continue;
      if (TOBL.GetString('GL_TYPELIGNE')<>'ART') and (TOBL.GetString('GL_TYPELIGNE')<>'SD')  then continue;
      //
      if (TOBL.GetString('GL_TYPELIGNE')='ART') then
      begin
        if (Arrondi(TOBL.GetDouble('GL_MONTANTPA')-TOBL.GetDouble('MTTRANSFERT'),TFFActure(FF).DEV.Decimale)= 0) then
        begin
          PGIInfo('La ligne '+TOBL.GetString('GL_NUMLIGNE')+' est déjà transférée en totalité');
          Exit;
        end;
        if IsOuvrage(TOBL) then
        begin
          if not TransfertOnlyOuvrage(TOBL) then
          begin
            PGIInfo('Transfert de l''ouvrage '+TOBL.GetString('GL_CODEARTICLE')+'impossible#13#10 '+'Des sous détails sont déjà transférés.');
            Exit;
          end;
        end;
      end else if (TOBL.GetString('GL_TYPELIGNE')='SD') then
      begin
        TOBO := FindLigneOuvFromUnique (TOBL,TOBOUvrages,TOBL.GetInteger('UNIQUEBLO'));
        if TOBO <> nil then
        begin
          TOBLigne := FindLigneDocFromDetail (TOBPiece,TOBL);
          if TOBLigne <> nil then
          begin
            if TOBT.findfirst(['BT3_NUMORDRE','BT3_UNIQUEBLO'],[TOBLigne.GetInteger('GL_NUMORDRE'),0],True)<> nil then Continue; // l'ouvrage est sélectionné --> on ne prend pas le sous détail
            MontantPa := TOBO.GetDouble('BLO_MONTANTPA') * TOBLigne.GetDouble('GL_QTEFACT');
            if TOBLigne.getDOuble('MTTRANSFERT')<>0 then
            begin
              PGIInfo('Impossible : L''ouvrage '+TOBLigne.GetString('GL_CODEARTICLE')+'#13#10 '+' est déjà partiellement ou totalement transféré.');
              Exit;
            end;
          end;
        end else Exit;
        //
        if (Arrondi(Montantpa-TOBL.GetDouble('MTTRANSFERT'),TFFActure(FF).DEV.Decimale)= 0) then
        begin
          PGIInfo('La ligne '+TOBL.GetString('GL_NUMLIGNE')+' est déjà transférée en totalité');
          Exit;
        end;
        if (Arrondi(TOBLigne.GetDouble('GL_MONTANTPA')-TOBLigne.GetDouble('MTTRANSFERT'),TFFActure(FF).DEV.Decimale)= 0) then
        begin
          PGIInfo('L''ouvrage '+TOBLigne.GetString('GL_NUMLIGNE')+' est déjà transféré en totalité');
          Exit;
        end;
      end;
      //
      LastTOBL := TOBL;
      TL := TOB.Create('BTRFDETAIL',TOBT,-1);
      TL.SetInteger('BT3_UNIQUE',TOBT.GetInteger('BT2_UNIQUE'));
      TL.SetString('BT3_TYPELIGNETRF','000');
      if (TOBL.GetString('GL_TYPELIGNE')='ART') then
      begin
        TL.SetInteger('BT3_NUMORDRE',TOBL.GetInteger('GL_NUMORDRE'));
      end else
      begin
        TL.SetInteger('BT3_UNIQUEBLO',TOBL.GetInteger('UNIQUEBLO'));
      end;
      TL.Data := TOBL;
    end;
    Result := True;
  end;

begin
  TOBPIece := TFFacture (FF).LaPieceCourante;
  TOBTRFPOC := TFFActure(FF).XTOBTRFPOC;
  LastTOBL := nil;
  //
  TOBT := TOB.create ('BTRFENTETE',nil,-1);
  TOBT.SetInteger('BT2_UNIQUE',TOBPiece.GetInteger('MAXUNIQUE')+1);
  TOBT.SetDateTime('BT2_DATECREATION',Now);
  TOBT.SetDateTime('BT2_DATEMODIF',NowH);
  TOBT.SetSTring('BT2_CREATEUR',V_PGI.User);
  TOBT.SetSTring('BT2_UTILISATEUR',V_PGI.User);
  TOBT.AddChampSupValeur('MODE','CREATION');
  TOBT.AddChampSupValeur('OKOK','-');
  if not BeforeTransfert (TOBPiece,TOBTRFPOC,TOBT) then exit;
  Arow := LastTOBL.GetIndex+1;

  //
  if TOBT.detail.count > 0 then
  begin
    TheTOB := TOBT;
    AGLLanceFiche('BTP','BTSAISTRFPOC','','','');
    TheTOB := nil;
    if TOBT.GetString('OKOK')='X' then
    begin
      TOBT.ChangeParent(TOBTRFPOC,-1);
      TFFacture(FF).SupLesLibDetail (TOBpiece);
      TFFacture(FF).CopierColleObj.deselectionneRows;
      TOBPiece.SetInteger('MAXUNIQUE',TOBPiece.GetInteger('MAXUNIQUE')+1);
      TFFacture(FF).AffichageDesDetailOuvrages;  // en TOB
      DefiniRowCount (FF,TOBPiece);

      TFFacture(FF).RefreshGrid (Acol,ARow);
      TFFacture(FF).PieceCoTrait.SetSaisie;
    end else
    begin
      TOBT.free;
    end;
  end else
  begin
    TOBT.Free;
  end;
end;

constructor TGestTransfert.create(TT: TForm);
var ThePop : Tcomponent;
begin
  TOBTransferts := TOB.Create ('LES TRASNFERTS',nil,-1);
  fCreatedPop := false;
  fusable := false;
  FF := TT;
  ThePop := TT.Findcomponent  ('POPBTP');
  if ThePop = nil then
  BEGIN
    // pas de menu BTP trouve ..on le cree
    POPGS := TPopupMenu.Create(TT);
    POPGS.Name := 'POPBTP';
    fCreatedPop := true;
  END else
  BEGIN
    POPGS := TPopupMenu(thePop);
  END;
  DefiniMenuPop(TT);
  // TToolWindow
  TTWTransfert := TToolWindow97.Create(FF);
  TTWTransfert.Name := 'XTRFS';
  TTWTransfert.caption := 'Informations sur transferts';
  TTWTransfert.Parent := FF;
  TTWTransfert.Visible := false;
  TTWTransfert.top := TFFacture(FF).TTVParag.top;
  TTWTransfert.left := TFFacture(FF).TTVParag.left - 10;
  TTWTransfert.Width := 600;
  TTWTransfert.height := TFFacture(FF).TTVParag.height;
  TTWTransfert.CloseButton := True;
  // TreeView
  TVTRansfert := TTreeView.Create(TTWTransfert);
  TVTRansfert.Parent := TTWTransfert;
  TVTRansfert.Align := alClient;
  TVTRansfert.MultiSelect := false;
  TVTRansfert.OnClick := TreeViewClick;
  //
  CurrentTransfert := Self;
end;

procedure TGestTransfert.DefiniMenu(TOBL: TOB);
var II : Integer;
begin
  if not fusable then Exit;
  for II := 0 to fMaxItems -1 do
  begin
    if MesMenuItem[II].Name = 'mListTransfert' then
    begin
      if IsExistsTransfert (TOBL) then MesMenuItem[II].Enabled := true
                                  else MesMenuItem[II].Enabled := false;
    end;
    if MesMenuItem[II].Name = 'mModifTransfert' then
    begin
      if ISTransfert (TOBL) then MesMenuItem[II].Enabled := true
                            else MesMenuItem[II].Enabled := false;
    end;
    if MesMenuItem[II].Name = 'mTransfert' then
    begin
      if ISTransfert (TOBL) then MesMenuItem[II].Enabled := false
                            else MesMenuItem[II].Enabled := true;
    end;
    if MesMenuItem[II].Name = 'POPYTS' then
    begin
      MesMenuItem[II].Enabled := (TOBL.GetString ('GL_TYPELIGNE')='ART');
    end;
  end;
end;

procedure TGestTransfert.DefiniMenuPop (Parent : Tform);
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
  MesMenuItem[fMaxItems] := TmenuItem.Create (parent);
  with MesMenuItem[fMaxItems] do
  begin
    Name := 'POPYTS';
    Caption := TraduireMemoire ('Avenants / TS');
    OnClick := POPYTSClick;
  end;
  inc (fMaxItems);
  MesMenuItem[fMaxItems] := TmenuItem.Create (parent);
  with MesMenuItem[fMaxItems] do
  begin
    Name := 'SEPNN';
    Caption := TraduireMemoire ('-');
  end;
  inc (fMaxItems);

  MesMenuItem[fMaxItems] := TmenuItem.Create (parent);
  with MesMenuItem[fMaxItems] do
  begin
    Name := 'mTransfert';
    Caption := TraduireMemoire ('Créer transfert');
    OnClick := AjoutTransfert;
  end;
  inc (fMaxItems);
  MesMenuItem[fMaxItems] := TmenuItem.Create (parent);
  with MesMenuItem[fMaxItems] do
  begin
    Name := 'mModifTransfert';
    Caption := TraduireMemoire ('Modifier transfert');
    OnClick := ModifTransfert;
    Enabled := false;
  end;
  inc (fMaxItems);
  MesMenuItem[fMaxItems] := TmenuItem.Create (parent);
  with MesMenuItem[fMaxItems] do
  begin
    Name := 'mListTransfert';
    Caption := TraduireMemoire ('Détail transferts');
    OnClick := DetailTransfert;
    Enabled := false;
  end;
  inc (fMaxItems);

  for Indice := 0 to fMaxItems -1 do
  begin
    if MesMenuItem [Indice] <> nil then POPGS.Items.Add (MesMenuItem[Indice]);
  end;
end;

destructor TGestTransfert.destroy;
var indice : integer;
begin
  for Indice := 0 to fMaxItems -1 do
  begin
    MesMenuItem[Indice].Free;
  end;
  if fcreatedPop then POPGS.free;
  CurrentTransfert := nil;
  TOBTransferts.free;
  TVTRansfert.Free;
  TTWTransfert.Free;
  inherited;
end;

procedure TGestTransfert.ConstitueTreeViewTrf;

  procedure AddDetails(TT : TOB; RootTN : TTreeNode);
  var Tn : TTreeNode;
     II : Integer;
  begin
    TN := TVTRansfert.Items.AddChild (RootTN,TT.GetValue('LIBELLE'));
    if TT.Data <> nil then Tn.data := TT.Data;
    if TT.Detail.count > 0 then
    begin
      for II := 0 to TT.Detail.count -1 do
      begin
        AddDetails(TT.detail[II],TN);
      end;
    end;
  end;

var Tn,RootTN: TTreeNode;
    II : Integer;
    TT : TOB;
begin
  TVTRansfert.items.clear;

  RootTN := TVTRansfert.items.add(nil,'Détail des transferts');
  for II := 0 to TOBTransferts.Detail.count -1 do
  begin
    TT := TOBTransferts.detail[II];
    AddDetails(TT,RootTN);
  end;
end;


function TGestTransfert.ConstitueTOBDetailTransfert (TOBL,TOBTRFPOC : TOB) : boolean;
var TT,TOBT,TOBTS,ONELig,ONET,TDETT,TTT : TOB;
    NumOrdre,UniqueBLo : Integer;
    II,AA : Integer;
begin
  Result := false;
  TOBTransferts.ClearDetail;
  //
  if TOBL.getString('GL_TYPELIGNE')='SD' then
  begin
    NumOrdre := 0;
    Uniqueblo := TOBL.getInteger('UNIQUEBLO');
  end else
  begin
    NumOrdre := TOBL.GetInteger ('GL_NUMORDRE');
    UniqueBLo := 0;
  end;
  //
  TT := TOBTRFPOC.findfirst(['BT3_NUMORDRE','BT3_UNIQUEBLO'],[NumOrdre,UniqueBlo],True);
  repeat
    if TT = nil then Exit;
    //
    ONET := TT.Parent;
    TOBT := TOB.Create('UN TRANSFERT',TOBTransferts,-1);
    TOBT.AddChampSupValeur ('LIBELLE','Transfert N° '+ONET.GetString('BT2_UNIQUE'));
    //
    TTT:= TOB.Create('UN TRANSFERT',TOBT,-1);
    TTT.AddChampSupValeur('LIBELLE','Provenance');
    for II := 0 to ONET.detail.count -1 do
    begin
      TDETT := ONET.detail[II];
      ONELig := TOB(TDETT.data);
      if (TDETT.GetString('BT3_TYPELIGNETRF')<> '000') then break;
      //
      TOBTS := TOB.Create('UN DEPART',TTT,-1);
      TOBTS.AddChampSupValeur ('LIBELLE','Transfert de '+TDETT.GetString('GL_CODEARTICLE')+' '+OneLig.GetString('GL_LIBELLE')+'  pour un montant de '+STRF00(TDETT.GetDouble('BT3_MTTRANSFERT'),2) );
      TOBTS.Data := ONELig;
    end;

    TTT:= TOB.Create('UN TRANSFERT',TOBT,-1);
    TTT.AddChampSupValeur('LIBELLE','Destination');
    for II := 0 to ONET.detail.count -1 do
    begin
      TDETT := ONET.detail[II];
      ONELig := TOB(TDETT.data);
      if (TDETT.GetString('BT3_TYPELIGNETRF')= '000') then continue;
      if (TDETT.GetString('BT3_TYPELIGNETRF')= '001') and (TDETT.GeTString('BT3_CONTREP')='X') then continue;
      //
      TOBTS := TOB.Create('UNE DESTINATION',TTT,-1);
      TOBTS.AddChampSupValeur ('LIBELLE','Destination '+TDETT.GetString('GL_CODEARTICLE')+' '+OneLig.GetString('GL_LIBELLE')+'  pour un montant de '+STRF00(OneLig.GetDouble('GL_MONTANTPA'),2) );
      TOBTS.Data := ONELig;
    end;
    //
    TT := TOBTRFPOC.findnext(['BT3_NUMORDRE','BT3_UNIQUEBLO'],[NumOrdre,UniqueBlo],True);
  until TT = nil;
  Result := true;
end;

procedure TGestTransfert.DetailTransfert(Sender: TObject);
var TOBPiece,TOBL,TOBTRFPOC : TOB;
begin
  TOBPIece := TFFacture (FF).LaPieceCourante;
  TOBTRFPOC := TFFActure(FF).XTOBTRFPOC;
  //
  TOBL := TOBPiece.Detail[TFFacture(FF).GS.Row -1];
  if ConstitueTOBDetailTransfert (TOBL,TOBTRFPOC) then
  begin
    ConstitueTreeViewTrf;
    TTWTransfert.Visible := True;
    TVTRansfert.FullExpand;
  end;
end;

function TGestTransfert.IsTransfert(TOBL: TOB): boolean;
begin
  Result := false;
  if TOBL = nil then exit;
  if not TOBL.FieldExists('NUMTRANSFERT') or VarIsNull(TOBL.getValue('NUMTRANSFERT')) or (VarAsType(TOBL.getValue('NUMTRANSFERT'), varString) = #0) then exit;
  Result := (TOBL.GetString('NUMTRANSFERT')<>'') and (Valeur(TOBL.GetString('MTTRANSFERT'))<>0);
end;

function TGestTransfert.IsExistsTransfert(TOBL: TOB): boolean;
begin
  Result := false;
  if TOBL = nil then exit;
  Result := (TOBL.GetString('NUMTRANSFERT')<>'') or (Valeur(TOBL.GetString('MTTRANSFERT'))<>0);
end;


procedure TGestTransfert.ModifTransfert(Sender: TObject);
var TOBT,TOBL : TOB;
    TOBPiece : TOB;
    TOBTRFPOC : TOB;
    Arow,Acol : Integer;
begin
  Acol := TFFacture (FF).GS.Col;
  Arow := TFFacture (FF).GS.row;
  TOBPIece := TFFacture (FF).LaPieceCourante;
  TOBTRFPOC := TFFActure(FF).XTOBTRFPOC;
  TOBL := TOBPiece.Detail[TFFacture(FF).GS.Row -1];
  //
  TOBT := FindTransfert(TOBTRFPOC,TOBL);
  if TOBT <> nil then
  begin
    if not LastTransfert(TOBTRFPOC,TOBT) then
    begin
      PgiInfo('Vous ne pouvez modifier que le dernier transfert');
      exit;
    end;
    TOBT.AddChampSupValeur('MODE','MODIFICATION');
    TOBT.AddChampSupValeur('OKOK','-');
    TheTOB := TOBT;
    AGLLanceFiche('BTP','BTSAISTRFPOC','','','');
    TheTOB := nil;
    if TOBT.GetString('OKOK')='X' then
    begin
      TFFacture(FF).SupLesLibDetail (TOBpiece);
      TFFacture(FF).CopierColleObj.deselectionneRows;
      TFFacture(FF).AffichageDesDetailOuvrages;  // en TOB
//      Arow := TOBL.GetIndex+1;
      TFFacture(FF).RefreshGrid (Acol,ARow);
      TFFacture(FF).PieceCoTrait.SetSaisie;
    end else if TOBT.GetString('OKOK')='D' then
    begin
      TFFacture(FF).SupLesLibDetail (TOBpiece);
      TFFacture(FF).CopierColleObj.deselectionneRows;
      TFFacture(FF).AffichageDesDetailOuvrages;  // en TOB
      if TFFacture (FF).GS.row < TOBpiece.detail.count then
      begin
        TOBL := TOBPiece.Detail[TFFacture(FF).GS.Row -1]
      end else
      begin
        TOBL := TOBPiece.detail[TOBPiece.detail.count -1];
      end;
      Arow := TOBL.GetIndex+1;
      TFFacture(FF).RefreshGrid (Acol,ARow);
      TFFacture(FF).PieceCoTrait.SetSaisie;
      TOBT.free;
    end;
    TFFACture(FF).AfficheLaGrille; 
  end;
end;

procedure TGestTransfert.POPYTSClick(Sender: TObject);
begin
  TFFacture(FF).POPYTSClick(Self); 
end;

procedure TGestTransfert.SetPiece(NaturePiece: string);
begin
  if (VH_GC.BTCODESPECIF <> '001') or (NaturePiece <>'BCE') then
  begin
    ActiveMenu (false);
    fusable := false;
  end else
  begin
    ActiveMenu (true);
    fusable := true;
  end;
end;

procedure TGestTransfert.TreeViewClick(Sender: TObject);
var Tn: TTreeNode;
  TOBL: TOB;
  Acol, Arow: integer;
begin
  Tn := TVTRansfert.selected;
  TOBL := TOB(Tn.data);
  if (TOBL <> nil) then
  begin
    ARow := TOBL.GetValue('GL_NUMLIGNE');
    ACol := SG_Lib;
    TFFacture(FF).GS.SetFocus;
    TFFacture(FF).GSEnter(TFFacture(FF).GS);
    TFFacture(FF).GS.Row := Arow;
    TFFacture(FF).GS.col := Acol;
    TFFacture(FF).StCellCur := TFFacture(FF).GS.Cells[TFFacture(FF).GS.Col, TFFacture(FF).GS.Row];
  end;
end;

end.
