unit Factrgpbesoin;

interface

uses Hctrls,classes,UTOB,EntGC,forms,
  {$IFDEF EAGLCLIENT}
  maineagl,
  {$ELSE}
  DBCtrls, Db, {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} fe_main, UserChg, AglIsoflex,
  {$ENDIF}
  Menus,HEnt1,HmsgBox,AglInit,Windows,Graphics,SysUtils,SaisUtil;
const MAXITEMS = 2;

type
  TFactRgpBesoin = class
  	private
    	fCreatedPop : boolean;
    	POPGS : TPopupMenu;
    	FF : TForm;
      MesMenuItem: array[0..MAXITEMS] of TMenuItem;
      fMaxItems : integer;
      fUsable : boolean;
      fTobPiece,fTOBL : TOB;
    	fGS : THGrid;
      TheItem : TMenuItem;
    	procedure SetGrid(const Value: THGrid);
    	procedure SetTOBPiece(const Value: TOB);
    	procedure DefiniMenuPop(Parent: Tform);
    	procedure GSDecomposition (Sender : Tobject);
      procedure SetLigne(const Value: TOB);
    	procedure InitZeroDetail(LaTOBInterm: TOB);
      procedure SetFullInterm (laTOBinterm : TOB);

    public
      NbLignes : integer;
    	constructor create (TheFiche : TForm);
      destructor destroy; override;
      property Ligne : TOB write SetLigne;
      property Usable : boolean read fUsable;
      property TobPiece : TOB read fTOBPiece write SetTOBPiece;
      property GS : THGrid read fGs Write SetGrid;
      procedure ReajusteGrid;
      procedure DeleteDecomposition (Arow : integer);
    	procedure ModifQteCentralisateur (Arow : integer);
  end;

function IsDetailBesoin (TOBL : TOB; NumOrdre : integer=0) : boolean;
function IsCentralisateurBesoin (TOBL : TOB) : boolean;
procedure InformeCentralisation (FF : TForm ; TOBL : TOB; Arect : Trect);
procedure RepercutePrixCentralisation (TOBpiece,TOBL : TOB;EnHt : boolean; DEV : RDevise);

implementation

uses facture,FactTOB,FactureBtp,FactComm,FactVariante, BTPUtil,UtilPGI,factCalc;

procedure InformeCentralisation (FF : TForm ; TOBL : TOB; Arect : Trect);
var FFact : TFFacture;
  LastBrush, LastPen: Tcolor;
  TheChaine: string;
begin
  FFAct := TFFacture (FF);
  with FFact do
  begin
    TheChaine := '';
    if TOBL.NomTable = 'LIGNE' then
    begin
      if TOBL.GetDouble('GL_QTEFACT') <> 0 then TheChaine := trim(strs(TOBL.GetValue('GL_QTEFACT'), V_PGI.okDecQ));
    end;

    LastBrush := GS.Canvas.Brush.Color;
    LastPen := GS.Canvas.Pen.Color;
    GS.Canvas.FillRect(ARect);
    GS.Canvas.Brush.Style := bsSolid;
    GS.Canvas.Brush.Color := clBlue;
    GS.Canvas.Pen.Color := clBlue;
    GS.Canvas.Polygon([Point(Arect.left, Arect.top), point(Arect.left + 4, Arect.top), Point(Arect.left, Arect.top + 4)]);
    GS.Canvas.Brush.Color := LastBrush;
    GS.Canvas.Pen.Color := LastPen;
    GS.Canvas.TextOut(Arect.Right - canvas.TextWidth(TheChaine) - 3, Arect.Top + 2, TheChaine);
  end;
end;

function IsCentralisateurBesoin (TOBL : TOB) : boolean;
var prefixe : string;
begin
	result := false;
  Prefixe := GetPrefixeTable (TOBL);
  if Prefixe = '' then exit;

	if (TOBL.GetValue(prefixe+'_TYPELIGNE')='CEN') then result := true;
end;

function IsDetailBesoin (TOBL : TOB; NumOrdre : integer=0) : boolean;
var prefixe : string;
begin
	result := false;
  Prefixe := GetPrefixeTable (TOBL);
	if (TOBL.GetValue(prefixe+'_TYPELIGNE')='ART') then
  begin
  	if NUmOrdre = 0 then
    begin
    	if (TOBL.GetValue(prefixe+'_LIGNELIEE') >0) then result := true
    end else
    begin
    	if (TOBL.GetValue(prefixe+'_LIGNELIEE')=NumOrdre) then result := true;
    end;
  end;
end;

{ TFactRgpBesoin }

constructor TFactRgpBesoin.create (TheFiche : TForm);
var ThePop : Tcomponent;
begin
  inherited create;
  TheItem := nil;
  fTobPiece := nil;
  fusable := true; // par defaut
  if not (TheFiche is TFFacture) then
  begin
  	fusable := false;
    exit;
  end else FF := TheFiche;

  ThePop := TFFacture(FF).Findcomponent  ('POPBTP');
  if ThePop = nil then
  BEGIN
    // pas de menu BTP trouve ..on le cree
    POPGS := TPopupMenu.Create(FF);
    POPGS.Name := 'POPBTP';
    fCreatedPop := true;
  END else
  BEGIN
    fCreatedPop := false;
    POPGS := TPopupMenu(thePop);
  END;
  DefiniMenuPop(FF);
end;

procedure TFactRgpBesoin.DefiniMenuPop(Parent: Tform);
var Indice : integer;
begin
  fmaxitems := 0;
  if not fcreatedPop then
  begin
    MesMenuItem[fmaxitems] := TmenuItem.Create (parent);
    with MesMenuItem[fmaxitems] do
      begin
      Caption := '-';
      end;
    inc (fmaxitems);
  end;
  // CTRL+ALT+G = Voir la decomposition
  MesMenuItem[fmaxitems] := TmenuItem.Create (parent);
  with MesMenuItem[fmaxitems] do
    begin
    Caption := TraduireMemoire ('Voir/modifier le détail');  // par défaut
    Name := 'BDECOMPOSITION';
    OnClick := GSDecomposition;
    enabled := false;
    end;
  MesMenuItem[fmaxitems].ShortCut := ShortCut( Word('G'), [ssCtrl,ssAlt]);
  inc (fmaxitems);

  for Indice := 0 to fmaxitems -1 do
    begin
      if MesMenuItem [Indice] <> nil then POPGS.Items.Add (MesMenuItem[Indice]);
    end;
end;

destructor TFactRgpBesoin.destroy;
var indice : integer;
begin
  for Indice := 0 to fmaxitems -1 do
  begin
    MesMenuItem[Indice].Free;
  end;
  if fcreatedPop then POPGS.free;
  inherited;
end;

procedure TFactRgpBesoin.InitZeroDetail (LaTOBInterm : TOB);
var Indice  : integer;
		TOBL    : TOB;
    OldQTe  : double;
    OldMt   : Double;
begin

  for Indice := 0 to LaTOBinterm.detail.count -1 do
  begin
  	TOBL := LaTOBinterm.detail[Indice];
    if TOBL.getValue('GL_TYPELIGNE')='ART' then
    begin
      OldQte := TOBL.GetValue('GL_QTEFACT');
      TOBL.putValue('GL_QTEFACT',0);
      TOBL.PutValue('GL_QTERESTE', TOBL.GetValue('GL_QTERESTE')- OldQte + TOBL.GetValue('GL_QTEFACT'));  //=> I don't understand ???
      TOBL.PutValue('GL_QTESTOCK', TOBL.GetValue('GL_QTEFACT'));
      PutQteReliquat(Tobl, OldQte);
      //--- GUINIER ---
      if CtrlOkReliquat(TOBL, 'GL') then
      Begin
        OldMt  := TOBL.GetValue('GL_MONTANTHTDEV');
        TOBL.PutValue('GL_MTRESTE',  TOBL.GetValue('GL_MTRESTE') - OldMt);
        PutMTReliquat(Tobl, OldMt);
      end;
    end;
  end;

end;

procedure TFactRgpBesoin.SetFullInterm (laTOBinterm : TOB);
var Indice : integer;
		TOBL : TOB;
    OldQTe : double;
begin
  for Indice := 0 to LaTOBinterm.detail.count -1 do
  begin
  	TOBL := LaTOBinterm.detail[Indice];
    if TOBL.getValue('GL_TYPELIGNE')='ART' then
    begin
      OldQte := TOBL.GetValue('GL_QTEFACT');
      TOBL.putValue('GL_QTEFACT', OldQte + TOBL.GetValue('GL_QTERELIQUAT'));
      TOBL.PutValue('GL_QTERESTE',0);
      TOBL.PutValue('GL_QTESTOCK',TOBL.GetValue('GL_QTEFACT'));
      PutQteReliquat(Tobl, OldQte);
      // --- GUINIER ---
      if CtrlOkReliquat(TOBL, 'GL') then
      begin
        OldQte := TOBL.GetValue('GL_MONTANTHTDEV');
        TOBL.putValue('GL_MONTANTHTDEV', OldQte + TOBL.GetValue('GL_MTRELIQUAT'));
        TOBL.PutValue('GL_MTRESTE',0);
        PutMTReliquat(Tobl, OldQte);
      end;
    end;
  end;
end;

procedure TFactRgpBesoin.GSDecomposition(Sender: Tobject);
var LaTOBInterm,TOBL,TOBSearch,TOBInsere,TOBArech : TOB;
    Lindice     : integer;
    Traitement  : string;
    OldQte      : Double;
    OldMt       : Double;
begin
	if ((not TFFacture(FF).TransfoPiece) or (TFFacture(FF).DuplicPiece)) and (TFFacture(FF).Action <> TaModif) then exit;
  LaTOBInterm := TOB.Create ('DECOMPOSITION',nil,-1);
  TRY
  	Traitement := 'VISU';
    TOBL := GetTOBLigne (fTOBPiece,Fgs.row);
    LaTOBInterm.AddChampSupValeur ('ARTICLE',   TOBL.GetValue('GL_CODEARTICLE'));
    LaTOBInterm.AddChampSupValeur ('LIGNE',     TOBL.GetValue('GL_NUMORDRE'));
    LaTOBInterm.AddChampSupValeur ('LIBELLE',   TOBL.GetValue('GL_LIBELLE'));
    LaTOBInterm.AddChampSupValeur ('QTEAREPART',TOBL.GetValue('GL_QTEFACT'));
    LaTOBInterm.AddChampSupValeur ('QTEREPART', TOBL.GetValue('GL_QTEFACT'));
    if CtrlOkReliquat(TOBL, 'GL') then
    begin
      LaTOBInterm.AddChampSupValeur ('MTAREPART',TOBL.GetValue('GL_MONTANTHTDEV'));
      LaTOBInterm.AddChampSupValeur ('MTREPART', TOBL.GetValue('GL_MONTANTHTDEV'));
    end;

    LaTOBInterm.AddChampSupValeur ('RETOUR',-1);
    Lindice := TOBL.GetValue('GL_LIGNELIEE');
    TOBSearch := fTOBPiece.FindFirst (['GL_TYPELIGNE','GL_LIGNELIEE'],['ART',Lindice],true);
    repeat
      if TOBSearch <> nil then
      begin
        TOBInsere := TOB.Create ('LIGNE',LaTOBinterm,-1);
        TOBInsere.Dupliquer (TOBSearch,false,true);
    		TOBSearch := fTOBPiece.Findnext (['GL_TYPELIGNE','GL_LIGNELIEE'],['ART',Lindice],true);
      end;
    until TOBSearch = nil;

    if LaTOBInterm.Detail.Count > 0 then
    begin
      TheTOB := LaTOBInterm;

      //    if (TFFacture(FF).TransfoPiece) and (not TFFacture(FF).DuplicPiece) and (not TFFacture(FF).IsReliquatGere ) then Traitement := 'REPARTITION' else Traitement := 'RECEPTIONNE' ;
  		if (TFFacture(FF).TransfoPiece) and (TFFacture(FF).IsReliquatGere)  then Traitement := 'RECEPTIONNE' else Traitement := 'MODIFQTE';

      TFFacture(FF).SauveColList;
      FGS.synEnabled := false;
      AGLLanceFiche ('BTP','BTREPARTBESOIN','','','TRAITEMENT='+Traitement);
      TFFacture(FF).RestoreColList;
      Fgs.Invalidate ;
      FGS.synEnabled := True;
      //
      if (TheTOB <> nil) and (TheTOB.getValue('RETOUR')=0) then
      begin
        Exit;
      end;
      //
    	if TheTOB.getValue('QTEREPART')<>TheTOB.GetValue('QTEAREPART') then
      begin
        if Traitement = 'VISU' then
        begin
          OldQte := TOBL.GetValue('GL_QTEFACT');
          TOBL.PutValue('GL_QTEFACT',TheTOB.GetValue('QTEREPART'));
          TOBL.PutValue('GL_QTESTOCK',TheTOB.GetValue('QTEREPART'));
          TOBL.PutValue('GL_QTERESTE',TOBL.GetValue('GL_QTERESTE') - OldQte + TheTOB.GetValue('QTEREPART'));
          PutQteReliquat (TOBL,OldQte);
          // --- GUINIER ---
          if CtrlOkReliquat(TOBL, 'GL') then
          begin
            OldMt := TOBL.GetValue('GL_MONTANTHTDEV');
            TOBL.PutValue('GL_MONTANTHTDEV', TheTOB.GetValue('MTREPART'));
            TOBL.PutValue('GL_MTRESTE',TOBL.GetValue('GL_MTRESTE') - OldMt + TheTOB.GetValue('MTREPART'));
            PutMTReliquat (TOBL,OldMt);
          end;
          //
          TOBL.putValue('GL_RECALCULER','X');
          TFFacture(FF).AfficheLaLigne (fgs.row);
          //
          if fgs.col = SG_QF then TFFacture(FF).StCellCur := fGS.cells[fgs.col,fgs.row];
          //
          for Lindice := 0 to TheTOB.detail.count -1 do
          begin
            TOBARech := TheTOB.detail[Lindice];
            TOBInsere := fTOBPiece.findFirst(['GL_NUMORDRE'],[TOBARech.GetValue('GL_NUMORDRE')],true);
            TOBInsere.PutValue('GL_QTEFACT', TOBARech.GetValue('GL_QTEFACT'));
            TOBInsere.PutValue('GL_QTESTOCK',TOBARech.GetValue('GL_QTEFACT'));
            TOBInsere.PutValue('GL_QTERESTE',TOBARech.GetValue('GL_QTERESTE'));
            PutQteReliquat (TOBInsere,OldQte);
            // --- GUINIER ---
            if CtrlOkReliquat(TOBARech, 'GL') then
            begin
              TOBInsere.PutValue('GL_MONTANTHTDEV', TOBARech.GetValue('GL_MONTANTHTDEV'));
              TOBInsere.PutValue('GL_MTRESTE',      TOBARech.GetValue('GL_MTRESTE'));
              PutMTReliquat (TOBL,OldMt);
            end;              //
            TOBInsere.putValue('GL_RECALCULER','X');
          end;
        end
        else
          begin
          	OldQte := TOBL.GetValue('GL_QTEFACT');
          	TOBL.PutValue('GL_QTEFACT',TheTOB.GetValue('QTEREPART'));
          	TOBL.PutValue('GL_QTESTOCK',TheTOB.GetValue('QTEREPART'));
            TOBL.PutValue('GL_QTERESTE',TOBL.GetValue('GL_QTERESTE')-OldQte+ TheTOB.GetValue('QTEREPART'));
            PutQteReliquat (TOBL,OldQte);
            if CtrlOkReliquat(TOBL, 'GL') then
            begin
              OldMt := TOBL.GetValue('GL_MONTANTHTDEV');
            	TOBL.PutValue('GL_MONTANTHTDEV', TheTOB.GetValue('MTREPART'));
              TOBL.PutValue('GL_MTRESTE',TOBL.GetValue('GL_MTRESTE')- OldMt + TheTOB.GetValue('MTREPART'));
              PutMTReliquat (TOBL,OldMt);
            end;
            TOBL.putValue('GL_RECALCULER','X');
            TFFacture(FF).AfficheLaLigne (fgs.row);
            if fgs.col = SG_QF then TFFacture(FF).StCellCur := fGS.cells[fgs.col,fgs.row];
            for Lindice := 0 to TheTOB.detail.count -1 do
            begin
              TOBARech  := TheTOB.detail[Lindice];
              TOBInsere := fTOBPiece.findFirst(['GL_NUMORDRE'],[TOBARech.GetValue('GL_NUMORDRE')],true);
              TOBInsere.PutValue('GL_QTEFACT', TOBARech.GetValue('GL_QTEFACT'));
              TOBInsere.PutValue('GL_QTESTOCK',TOBARech.GetValue('GL_QTEFACT'));
              TOBInsere.PutValue('GL_QTERESTE',TOBARech.GetValue('GL_QTERESTE'));
              PutQteReliquat (TOBInsere,OldQte);
              // --- GUINIER ---
              if CtrlOkReliquat(TOBArech, 'GL') then
              begin
                TOBInsere.PutValue('GL_MONTANTHTDEV',TOBARech.GetValue('GL_MONTANTHTDEV'));
                TOBInsere.PutValue('GL_MTRESTE',TOBARech.GetValue('GL_MTRESTE'));
                PutMTReliquat (TOBInsere,OldMt);
              end;

            	TOBInsere.putValue('GL_RECALCULER','X');
            end;
          end;
          fTOBPiece.PutValue('GP_RECALCULER','X');
      end;
      TheTOB := nil;
    end;
  FINALLY
    TheTOB := nil;
  	LaTOBInterm.free;
    if fTOBPiece.GetValue('GP_RECALCULER')='X' then TFFacture(FF).CalculeLaSaisie (-1,-1,True);
  END;
end;

procedure TFactRgpBesoin.ModifQteCentralisateur (Arow : integer);
var LaTOBInterm,TOBL,TOBSearch,TOBInsere,TOBArech : TOB;
    Lindice     : integer;
    Traitement  : string;
    OldQte      : double;
    OldMt       : Double;
    Ok_Reliquat : Boolean;
begin
	if ((not TFFacture(FF).TransfoPiece) or (TFFacture(FF).DuplicPiece)) and (TFFacture(FF).Action <> TaModif) then exit;
  LaTOBInterm := TOB.Create ('DECOMPOSITION',nil,-1);
  TRY
  	if (TFFacture(FF).TransfoPiece) or (TFFacture(FF).IsReliquatGere)  then Traitement := 'RECEPTIONNE' else Traitement := 'MODIFQTE';
    TOBL := GetTOBLigne (fTOBPiece,Fgs.row);
    Ok_Reliquat := CtrlOkReliquat(TOBL, 'GL');
    LaTOBInterm.AddChampSupValeur ('ARTICLE',TOBL.GetValue('GL_CODEARTICLE'));
    LaTOBInterm.AddChampSupValeur ('LIGNE',TOBL.GetValue('GL_NUMORDRE'));
    LaTOBInterm.AddChampSupValeur ('LIBELLE',TOBL.GetValue('GL_LIBELLE'));
    LaTOBInterm.AddChampSupValeur ('QTEAREPART',Valeur(FGs.Cells[SG_QF,Arow]));
    LaTOBInterm.AddChampSupValeur ('QTEREPART',TOBL.GetValue('GL_QTEFACT'));
    LaTOBInterm.AddChampSupValeur ('QTETOTALE',TOBL.GetValue('GL_QTEFACT')      + TOBL.GetValue('GL_QTERELIQUAT'));
    if Ok_Reliquat then LaTOBInterm.AddChampSupValeur ('MTTOTAL',  TOBL.GetValue('GL_MONTANTHTDEV') + TOBL.GetValue('GL_MTRELIQUAT'));

    LaTOBInterm.AddChampSupValeur ('RETOUR',1);
    Lindice := TOBL.GetValue('GL_LIGNELIEE');
    TOBSearch := fTOBPiece.FindFirst (['GL_TYPELIGNE','GL_LIGNELIEE'],['ART',Lindice],true);
    repeat
      if TOBSearch <> nil then
      begin
        TOBInsere := TOB.Create ('LIGNE',LaTOBinterm,-1);
        TOBInsere.Dupliquer (TOBSearch,false,true);
    		TOBSearch := fTOBPiece.Findnext (['GL_TYPELIGNE','GL_LIGNELIEE'],['ART',Lindice],true);
      end;
    until TOBSearch = nil;
    if LaTOBInterm.Detail.Count > 0 then
    begin
      TheTOB := LaTOBInterm;
      if (LaTOBInterm.GetValue ('QTEAREPART') = 0) and (Traitement = 'RECEPTIONNE') Then
      begin
        InitZeroDetail (LaTOBInterm);
    		LaTOBInterm.AddChampSupValeur ('QTEREPART',0);
    		LaTOBInterm.PutValue ('RETOUR',0);
      end else
      if (LaTOBInterm.GetValue ('QTEAREPART') = LaTOBInterm.GetValue ('QTETOTALE')) and (Traitement = 'RECEPTIONNE') Then
      begin
      	SetFullInterm (laTOBinterm);
    		LaTOBInterm.AddChampSupValeur ('QTEREPART',LaTOBInterm.GetValue ('QTEAREPART'));
    		LaTOBInterm.PutValue ('RETOUR',0);
      end
      else
      begin
        TFFacture(FF).SauveColList;
        FGS.synEnabled := false;
        AGLLanceFiche ('BTP','BTREPARTBESOIN','','','TRAITEMENT='+Traitement);
        TFFacture(FF).RestoreColList;
        Fgs.Invalidate ;
        FGS.synEnabled := True;
      end;
      if (TheTOB <> nil) and (TheTOB.getValue('RETOUR')=0) then
      begin
        OldQte := TOBL.GetValue('GL_QTEFACT');
        TOBL.PutValue('GL_QTEFACT',TheTOB.GetValue('QTEREPART'));
        TOBL.PutValue('GL_QTESTOCK',TheTOB.GetValue('QTEREPART'));
        TOBL.PutValue('GL_QTERESTE',TOBL.GetValue('GL_QTERESTE')-OldQte+ TheTOB.GetValue('QTEREPART'));
        PutQteReliquat (TOBL,OldQte);
        if Ok_Reliquat then
        begin
          OldMT := TOBL.GetValue('GL_MONTANTHTDEV');
          TOBL.PutValue('GL_MONTANTHTDEV', TheTOB.GetValue('MTREPART'));
          TOBL.PutValue('GL_MTRESTE', TOBL.GetValue('GL_MTRESTE')- OldMt + TheTOB.GetValue('MTREPART'));
          PutMTReliquat (TOBL,OldMt);
        end;

        TOBL.putValue('GL_RECALCULER','X');

        TFFacture(FF).AfficheLaLigne (fgs.row);
        if fgs.col = SG_QF then TFFacture(FF).StCellCur := fGS.cells[fgs.col,fgs.row];
        for Lindice := 0 to TheTOB.detail.count -1 do
        begin
          TOBARech := TheTOB.detail[Lindice];
          TOBInsere := fTOBPiece.findFirst(['GL_NUMORDRE'],[TOBARech.GetValue('GL_NUMORDRE')],true);
          OldQte := TOBInsere.GetValue('GL_QTEFACT');
          TOBInsere.PutValue('GL_QTEFACT',TOBARech.GetValue('GL_QTEFACT'));
          TOBInsere.PutValue('GL_QTESTOCK',TOBARech.GetValue('GL_QTEFACT'));
          TOBInsere.PutValue('GL_QTERESTE',TOBARech.GetValue('GL_QTERESTE'));
          PutQteReliquat (TOBInsere,OldQte);
          // --- GUINIER ---
          IF CtrlOkReliquat(TOBARech, 'GL') then
          begin
            OldMT := TOBInsere.GetValue('GL_MONTANTHTDEV');
            TOBInsere.PutValue('GL_MONTANTHTDEV',TOBARech.GetValue('GL_MONTANTHTDEV'));
            TOBInsere.PutValue('GL_MTRESTE',TOBARech.GetValue('GL_MTRESTE'));
            PutMTReliquat (TOBInsere,OldMt);
          end;

          TOBInsere.putValue('GL_RECALCULER','X');
        end;
        fTOBPiece.PutValue('GP_RECALCULER','X');
      end;
      TheTOB := nil;
    end;
  FINALLY
  	LaTOBInterm.free;
  END;
end;

procedure TFactRgpBesoin.ReajusteGrid;
var Indice : integer;
		TOBL : TOB;
begin
  if not fusable then exit; // si rien a traiter pourquoi s'embetter
  NbLignes := 0;

  if fGS=nil then exit;

  fGS.BeginUpdate ;
  for Indice := 1 to fGS.RowCount-1 do
  begin
  	fGS.RowHeights [Indice] := fGS.DefaultRowHeight; // on positionne la ligne a la taille par defaut
  	if Indice <= fTobPiece.detail.count then
    begin
      TOBL := GetTOBLigne (fTOBPiece,Indice);
      if IsDetailBesoin (TOBL) then
      begin
        Inc(NbLignes);
        fGS.RowHeights [Indice] := -1; // on cache la ligne
      end;
    end;
  end;
	fGS.EndUpdate;
end;

procedure TFactRgpBesoin.SetLigne(const Value: TOB);
var CurItem : TmenuItem;
    indice : integer;
    TypeLigne : string;
    perfixe : string;
begin
  if not fusable then exit;
  if TheItem <> nil then TheItem.enabled := false;
  fTOBL := value;
  if fTOBL = nil then exit;
  if (POPGS <> nil) and (TheItem = nil) then
  begin
    for indice := 0 to POPGS.Items.Count do
    begin
      CurItem := POPGS.Items [indice];
      if CurItem.Name = 'BDECOMPOSITION' then BEGIN TheItem := CurItem; Break; END;
    end;
  end;
  if TheItem = nil then exit;
  TheItem.enabled := false;
  if not IsCentralisateurBesoin  (fTOBL) then BEGIN exit; END;
  TheItem.enabled := true;
end;

procedure TFactRgpBesoin.SetGrid(const Value: THGrid);
begin
	if not fUsable then exit;
  fGs := Value;
  if fCreatedPop then if not assigned(fGS.popupmenu) then fGS.PopupMenu := POPGS;
end;

procedure TFactRgpBesoin.SetTOBPiece(const Value: TOB);
begin
  fTOBPiece := Value;
  if GetInfoParPiece(fTOBPiece.GetValue('GP_NATUREPIECEG'), 'GPP_VENTEACHAT') <> 'ACH' then
  begin
  	fUsable := false;
    fTOBPiece := nil;
  	exit;
  end;
end;

procedure TFactRgpBesoin.DeleteDecomposition (Arow : integer);
var TOBLig : TOB;
		Indice : integer;
begin
	if not fusable then exit;
  For Indice := Arow+1 to fGS.rowcount -1 do
  begin
    TOBLig := GetTOBLigne (fTobPiece,Indice);
    if TOBLig <> nil then
    begin
    	if IsCentralisateurBesoin (TOBLIG) then exit;
    	if IsDetailBesoin (TOBLIG,TOBLIg.GetValue('GL_LIGNELIEE')) then TFfacture(FF).TheLigADeleted := Indice; 
    end;
  end;
end;

procedure RepercutePrixCentralisation (TOBpiece,TOBL : TOB;EnHt : boolean; DEV : RDevise);
var TOBR : TOB;
		Centralisateur : integer;
begin
  Centralisateur := TOBL.getValue('GL_LIGNELIEE');
  TOBL.PutValue('GL_RECALCULER','X');
  TOBR := TOBPiece.findfirst(['GL_TYPELIGNE','GL_LIGNELIEE'],['ART',Centralisateur],true);
  if TOBR <> nil then
  begin
    repeat
      if TOBR <> nil then
      begin
        TOBR.SetDouble('GL_DPA',TOBL.getDouble('GL_DPA'));
        TOBR.SetDouble('GL_COEFFG',TOBL.getDouble('GL_COEFFG'));
        TOBR.SetDouble('GL_COEFFC',TOBL.getDouble('GL_COEFFC'));
        TOBR.SetDouble('GL_COEFFR',TOBL.getDouble('GL_COEFFR'));
        TOBR.SetDouble('GL_COEFMARG',TOBL.getDouble('GL_COEFMARG'));
        TOBR.SetDouble('GL_DPR',TOBL.getDouble('GL_DPR'));
        TOBR.PutValue('GL_PUHTDEV',TOBL.GetValue('GL_PUHTDEV'));
        TOBR.PutValue('GL_PUTTCDEV',TOBL.GetValue('GL_PUTTCDEV'));
        TOBR.SetDouble('GL_PUHT',TOBL.getDouble('GL_PUHT'));
        TOBR.SetDouble('GL_PUHTNET',TOBL.getDouble('GL_PUHTNET'));
        TOBR.SetDouble('GL_PUHTNETDEV',TOBL.getDouble('GL_PUHTNETDEV'));
        TOBR.PutValue('GL_RECALCULER','X');
  			TOBR := TOBPiece.findnext(['GL_TYPELIGNE','GL_LIGNELIEE'],['ART',Centralisateur],true);
      end;
    until TOBR = nil;
  end;
end;

end.
