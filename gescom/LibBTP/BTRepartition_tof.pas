{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 27/12/2004
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BTREPARTITION ()
Mots clefs ... : TOF;BTREPARTITION
*****************************************************************}
Unit BTRepartition_tof ;

Interface

Uses StdCtrls, 
     Controls, 
     Classes, 
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} 
     mul, 
{$else}
     eMul,
     uTob,
{$ENDIF}
		 Hpanel,
     HCtrls,
     forms,
     sysutils,
     ComCtrls,
     HEnt1,
     HMsgBox,
     HsysMenu,
     CalcOlegenericBTP,
     EntGc,
     UTOF,
     Grids,
     AglInit,
     HTB97,
     Vierge,
     UTOB,uEntCommun,UtilTOBPiece;

Type
	TModeTraitement = (TrqVisu,TrqModifQCen,TrqRepart);

  TOF_BTREPARTITION = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  private
  	SG_SEL,SG_AFFAIRE,SG_NATUREPIECE,SG_NUMERO,SG_DATELIVPREVU,SG_GL_QTEFACT,SG_GL_QTERELIQUAT, SG_GL_MTRELIQUAT : integer;
  	LastSaisie : string;
  	PHaut,PBas : THpanel;
    LARTICLE : THLabel;
    QTEAREPART,QTEREPART : THNumEdit;
    fMode : TModeTraitement;
    LaGrid : THgrid;
    LesChamps : string;
    NbrCols : integer;
    HmTrad : THSystemMenu;
    Bannul,BValider : TToolBarButton97;
    procedure AfficheGrid;
    procedure GetComponents;
    procedure PrepareEcran;
    procedure PrepareGrid;
    procedure LaGridEnter(Sender: TObject);
    procedure LaGridCellEnter(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
    procedure LaGridCellExit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
    procedure LaGridRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
    procedure LaGridRowExit(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
    procedure ZoneSuivanteOuOk(var ACol, ARow: Longint; var Cancel: boolean);
    function ZoneAccessible(ACol, ARow: Longint): boolean;
    procedure GridEvt(Etat: boolean);
    procedure TraiteQte(Acol, Arow: integer);
		procedure LaGridDblClick (Sender : Tobject);
    procedure DefiniCols(LesChamps: string);
    procedure BannulCkick (Sender : Tobject);
    procedure BValiderClick (Sender : Tobject);
    procedure CalcQterepart;
  public
   	MyOwnTOB,MySavTOB : TOB;
  end ;

Implementation

uses factcomm,Facture,FactTob, BTPUtil;

procedure TOF_BTREPARTITION.AfficheGrid;
var Indice : integer;
begin
	For Indice := 0 to MyOwnTOB.detail.count -1 do
  begin
  	MyOwnTOB.detail[Indice].PutLigneGrid (LaGrid,Indice+1,false,false,lesChamps);
  end;
  HmTrad.ResizeGridColumns (LaGrid);
  LaGridEnter (Self);
end;

procedure TOF_BTREPARTITION.GridEvt (Etat : boolean);
begin
	if Etat then
  begin
    LAGrid.OnEnter := LAGridEnter;
    LAGrid.OnRowEnter := LAGridRowEnter;
    LAGrid.OnRowExit := LAGridRowEXit;
    LAGrid.OnCellEnter := LAGridCellEnter;
    LAGrid.OnCellExit := LAGridCellExit;
    LAGrid.OnDblClick  := LAGridDblClick;
  end else
  begin
    LAGrid.OnEnter := nil;
    LAGrid.OnRowEnter := nil;
    LAGrid.OnRowExit := nil;
    LAGrid.OnCellEnter := nil;
    LAGrid.OnCellExit := nil;
    LAGrid.OnDblClick  := nil;
  end
end;

procedure TOF_BTREPARTITION.GetComponents;
begin
	LAGrid := THgrid(getControl('GS'));
  LARTICLE := THLabel (GetControl('LARTICLE'));
  PHaut := THPanel (GetControl('PHAUT'));
  PBas := THPanel (GetControl('PBAS'));
  QTEAREPART := THNumEdit (GetControl('QTEAREPART'));
  QTEREPART := THNumEdit (GetControl('QTEREPART'));
  Bannul := TToolbarButton97 (GetControl('BFERME'));
  BValider := TToolbarButton97 (GetControl('BVALIDER'));
end;

procedure TOF_BTREPARTITION.DefiniCols (LesChamps : string);
var TheChamps,LeChamps : string;
		Indice : integer;
begin
	Indice := 0;
  TheChamps := lesChamps;
  leChamps := READTOKENST (TheChamps);
  repeat
  	if LeChamps <> '' then
    begin
    	if LeChamps = 'SEL'             then SG_SEL             := Indice else
    	if LeChamps = 'AFFAIRE'         then SG_AFFAIRE         := Indice else
  		if LeChamps = 'NATUREPIECE'     then SG_NATUREPIECE     := indice else
  		if LeChamps = 'NUMERO'          then SG_NUMERO          := indice else
  		if LeChamps = 'DATELIVPREVU'    then SG_DATELIVPREVU    := indice else
  		if LeChamps = 'GL_QTEFACT'      then SG_GL_QTEFACT      := indice else
  		if LeChamps = 'GL_QTERELIQUAT'  then SG_GL_QTERELIQUAT  := indice else
      if LeChamps = 'GL_MTRELIQUAT'   then SG_GL_MTRELIQUAT   := indice;
      inc (indice);
  		leChamps := READTOKENST (TheChamps);
    end;
  until LeChamps = '';
end;

procedure TOF_BTREPARTITION.PrepareEcran;
begin
	SG_SEL            := -1;
  SG_AFFAIRE        := -1;
  SG_NATUREPIECE    := -1;
  SG_NUMERO         := -1;
  SG_DATELIVPREVU   := -1;
  SG_GL_QTEFACT     := -1;
  SG_GL_QTERELIQUAT := -1;
  SG_GL_MTRELIQUAT  := -1;


	if (fMode = TrqVisu) or (fMode = TrqModifQCen) then
  begin
  	PBas.Visible := false;
    LesChamps := 'SEL;AFFAIRE;NATUREPIECE;NUMERO;GL_QTEFACT;';
    NbrCols := 5;
    Ecran.caption := 'Détail de la ligne '+IntToStr(MyOwnTOB.getValue('LIGNE'));
  end else
  begin
    LesChamps := 'SEL;AFFAIRE;NATUREPIECE;NUMERO;DATELIVPREVU;GL_QTEFACT;GL_QTERELIQUAT;GL_MTRELIQUAT';
    NbrCols := 8;
    Ecran.caption := 'Détail de la ligne '+IntToStr(MyOwnTOB.getValue('LIGNE'));
  end;
  DefiniCols (LesChamps);
  LARTICLE.Caption := TraduireMemoire('Article')+ ': ' + MyOwnTOB.GetValue('ARTICLE')+' '+MyOwnTOB.GetValue('LIBELLE');
  QTEAREPART.Value := MyOwnTOB.GetValue('QTEAREPART');
  QTEREPART.Value := MyOwnTOB.GetValue('QTEREPART');
  if (fMode = TrqModifQCen) then
  begin
  	PBas.Visible := True;
  end;
  Bannul.onclick := BannulCkick;
  BValider.onclick := BValiderClick;
  PrepareGrid;
end;

procedure TOF_BTREPARTITION.PrepareGrid;

	function GetDateLivPrevu (Cledoc : r_cledoc) : TDateTime;
  var SQl : string;
  		QQ : TQuery;
  begin
  	 result := iDate2099;
     SQL := 'SELECT GL_DATELIVRAISON FROM LIGNE WHERE '+WherePiece(CleDoc, ttdLigne,True,true);
     QQ := OpenSql (SQL,true,-1,'',true);
     if not QQ.eof then result := QQ.findfield('GL_DATELIVRAISON').AsDateTime;
     ferme (QQ);
  end;

var indice,Piece : integer;
		NaturePiece,Affaire,PieceOrigine : string;
    Cledoc : R_cledoc;
    DateLivPrevu : TDateTime;
    FF : string;
begin
  if V_PGI.OkDecQ > 0 then
  begin
    FF := '0.';
    for indice := 1 to V_PGI.OkDecQ - 1 do
    begin
      FF := FF + '#';
    end;
    FF := FF + '0';
  end;

	LAGrid.VidePile (false);
  LAGrid.RowCount := MyOwnTOB.detail.count +1;
  LAGrid.ColCount := NbrCols;

	// Entete
  LAGrid.Cells [SG_SEL,0] := '';
  LAGrid.ColWidths [SG_SEL] := 10;
  LAGrid.ColLengths [SG_SEL] := -1;   // non saisissable
  LAGrid.ColAligns[SG_SEL] := taCenter;
  //
  LAGrid.cells [SG_AFFAIRE,0] := TraduireMemoire('Affaire');
  LAGrid.ColWidths [SG_AFFAIRE] := 70;
  LAGrid.ColLengths [SG_AFFAIRE] := -1;   // non saisissable
  LAGrid.ColAligns[SG_AFFAIRE] := taCenter;
  //
  LAGrid.cells [SG_NATUREPIECE,0] := TraduireMemoire('Pièce');
  LAGrid.ColWidths [SG_NATUREPIECE] := 70;
  LAGrid.ColLengths [SG_NATUREPIECE] := -1;   // non saisissable
  LAGrid.ColAligns[SG_NATUREPIECE] := taCenter ;
  //
  LAGrid.cells [SG_NUMERO,0] := TraduireMemoire('Numéro');
  LAGrid.ColWidths [SG_NUMERO] := 24;
  LAGrid.ColLengths [SG_NUMERO] := -1;   // non saisissable
  LAGrid.ColAligns[SG_NUMERO] := taRightJustify ;
  //
  if SG_DATELIVPREVU <> -1 then
  begin
    LAGrid.cells [SG_DATELIVPREVU,0] := TraduireMemoire('Liv. Prévue');
    LAGrid.ColWidths [SG_DATELIVPREVU] := 40;
    LAGrid.ColLengths [SG_DATELIVPREVU] := -1;   // non saisissable
    LAGrid.ColAligns[SG_DATELIVPREVU] := taRightJustify ;
  end;
  //
  LAGrid.cells [SG_GL_QTEFACT,0] := TraduireMemoire('Quantité');
  LAGrid.ColWidths [SG_GL_QTEFACT] := 40;
  LAGrid.ColAligns[SG_GL_QTEFACT] := taRightJustify ;
  LAGrid.ColFormats [SG_GL_QTEFACT] := FF;
  //
  if SG_GL_QTERELIQUAT <> -1 then
  begin
    LAGrid.cells [SG_GL_QTERELIQUAT,0] := TraduireMemoire('Reste');
    LAGrid.ColWidths [SG_GL_QTERELIQUAT] := 40;
    LAGrid.ColAligns[SG_GL_QTERELIQUAT] := taRightJustify ;
  	LAGrid.ColLengths [SG_GL_QTERELIQUAT] := -1;   // non saisissable
    LAGrid.ColFormats [SG_GL_QTERELIQUAT] := FF;
  end;
  //
  if SG_GL_MTRELIQUAT <> -1 then
  begin
    LAGrid.cells [SG_GL_MTRELIQUAT,0]     := TraduireMemoire('Mt Reste');
    LAGrid.ColWidths [SG_GL_MTRELIQUAT]   := 50;
    LAGrid.ColAligns[SG_GL_MTRELIQUAT]    := taRightJustify ;
  	LAGrid.ColLengths [SG_GL_MTRELIQUAT]  := -1;   // non saisissable
    LAGrid.ColFormats [SG_GL_MTRELIQUAT]  := FF;
  end;

  // détail
  for indice := 0 to MyOwnTOB.detail.count -1 do
  begin
    Piece := 0;
    NaturePiece := '';
    DateLivPrevu := iDate2099;
  	PieceOrigine := MyOwnTOB.detail[Indice].getValue('GL_PIECEORIGINE');
    Affaire := MyOwnTOB.detail[Indice].getValue('GL_AFFAIRE');
    if PieceOrigine <> '' then
    begin
  		DecodeRefPiece (PieceOrigine ,cledoc);
			Naturepiece := Cledoc.NaturePiece;
      Piece := Cledoc.NumeroPiece;
      DateLivPrevu := GetDateLivPrevu (Cledoc);
    end;
  	MyOwnTOB.detail[Indice].AddChampSupValeur ('SEL',Indice+1);
  	MyOwnTOB.detail[Indice].AddChampSupValeur ('NATUREPIECE',GetInfoParPiece(Naturepiece, 'GPP_LIBELLE'));
  	MyOwnTOB.detail[Indice].AddChampSupValeur ('DATELIVPREVU',DateLivPrevu);
  	MyOwnTOB.detail[Indice].AddChampSupValeur ('NUMERO',Piece);
  	MyOwnTOB.detail[Indice].AddChampSupValeur ('AFFAIRE',BTPCodeAffaireAffiche(Affaire));
  end;
end;

procedure TOF_BTREPARTITION.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_BTREPARTITION.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_BTREPARTITION.OnUpdate ;
var Sok : boolean;
begin
  Inherited ;
end ;

procedure TOF_BTREPARTITION.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_BTREPARTITION.OnArgument (S : String ) ;
var St1 : String;
begin
  Inherited ;
  MyOwnTOB := TOB.Create ('DECOMPOSITION',nil,-1);
  MySavTOB := TOB.Create ('DECOMPOSITION',nil,-1);
  HmTrad := THSystemMenu.Create (ecran);
  MyOwnTOB.Dupliquer (LaTOB,true,true);
  MySavTOB.Dupliquer (LaTOB,true,true);
  st1 := Copy(S, 0, Pos('=',S) - 1);
  if st1 = 'TRAITEMENT' then
  begin
    st1 := Copy(S, Pos('=',S) + 1, 255);
    if st1 = 'VISU' then fMode := TrqVisu else if st1 = 'MODIFQTE' Then fmode := TrqModifQCen else fMode := TrqRepart;
  end;
  if fmode = TrqRepart then TFVierge(Ecran).Width := TFVierge(Ecran).Width + 100;
  GetComponents;
  PrepareEcran;
  AfficheGrid;
  GridEvt (true);
end ;

procedure TOF_BTREPARTITION.OnClose ;
begin
  Inherited ;
  HmTrad.free;
  MyOwnTOB.free;
  MySavTOB.free;
end ;

procedure TOF_BTREPARTITION.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_BTREPARTITION.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_BTREPARTITION.LAGridEnter(Sender: TObject);
var Acol,Arow : integer;
		Cancel : boolean;
begin
	LAGrid.CacheEdit;
  LAGrid.SynEnabled := false;
  ACol := LAGrid.col;
  ARow := LAGrid.Row;
  ZoneSuivanteOuOk(ACol, ARow, Cancel);
  LAGridRowEnter(LAGrid, ARow, cancel, False);
  LAGridCellEnter(LAGrid, ACol, ARow, Cancel);
  LAGrid.col := Acol;
  LAGrid.Row := ARow;
  LAGrid.SynEnabled := True;
  LAGrid.ShowEditor;
  LastSaisie := LAGrid.Cells[Acol,Arow];
end;

procedure TOF_BTREPARTITION.LAGridCellEnter(Sender: TObject; var ACol,ARow: Integer; var Cancel: Boolean);
begin
  ZoneSuivanteOuOk(ACol, ARow, Cancel);
  if not Cancel then
  begin
  	LastSaisie := LAGrid.Cells[Acol,Arow];
  end;
end;

procedure TOF_BTREPARTITION.LAGridCellExit(Sender: TObject; var ACol,ARow: Integer; var Cancel: Boolean);
begin

	if LAGrid.cells[Acol,Arow] = LastSaisie then exit;

  if Acol = SG_GL_QTEFACT then TraiteQte (Acol,Arow);

  MyOwnTOB.detail[Arow-1].PutLigneGrid (LAGrid,Arow,false,false,LesChamps);
end;

procedure TOF_BTREPARTITION.LAGridRowEnter(Sender: TObject; Ou: Integer;var Cancel: Boolean; Chg: Boolean);
begin
//
end;

procedure TOF_BTREPARTITION.LAGridRowExit(Sender: TObject; Ou: Integer;var Cancel: Boolean; Chg: Boolean);
begin

end;

procedure TOF_BTREPARTITION.CalcQterepart;
var indice : integer;
		TOBL : TOB;
begin
	MyOwnTOB.putValue('QTEREPART',0);
	for Indice := 0 To MyOwnTOB.detail.count -1 do
  begin
  	TOBL := MyOwnTOB.Detail[Indice];
		MyOwnTOB.putValue('QTEREPART',MyOwnTOB.GetValue('QTEREPART') + TOBL.GetValue('GL_QTEFACT'));
  end;
  QTEREPART.Value := MyOwnTOB.GetValue('QTEREPART');
end;

procedure TOF_BTREPARTITION.TraiteQte (Acol,Arow : integer);
var TOBL    : TOB;
		OldQte  : double;
    OldMt   : Double;
begin
  TOBL := MyOwnTOB.Detail[Arow-1];
  if TOBL = nil then exit;
	if (fMode = TrqVisu) or (fmode = TrqModifQCen) or (fmode = TrqRepart) then
  begin
  	OldQte := TOBL.GetValue('GL_QTEFACT');
    //
    TOBL.putValue('GL_QTEFACT',valeur(LAGrid.cells[Acol,Arow]));
    TOBL.PutValue('GL_QTERESTE', TOBL.GetValue('GL_QTERESTE')- OldQte + TOBL.GetValue('GL_QTEFACT'));
    TOBL.PutValue('GL_QTESTOCK',TOBL.GetValue('GL_QTEFACT'));
    //
    PutQteReliquat(Tobl, OldQte);
    //
    if CtrlOkReliquat(TOBL, 'GL') then
    begin
      OldMt  := TOBL.GetValue('GL_MONTANTHTDEV');
      TOBL.putValue('GL_MONTANTHTDEV', valeur(TOBL.GetValue('GL_QTEFACT') * TOBL.GetValue('GL_PUHTDEV')));
      TOBL.PutValue('GL_MTRESTE', TOBL.GetValue('GL_MTRESTE')- OldMt + TOBL.GetValue('GL_MONTANTHTDEV'));
      PutMTReliquat(Tobl, OldMT);
    end;
    CalcQterepart;
  end else
  begin
  end;
end;

function TOF_BTREPARTITION.ZoneAccessible(ACol, ARow: Integer): boolean;
VAR tobl : TOB;
begin
  Result := True;
  TOBL := MyOwnTOB.detail[ARow-1];
  if TOBL = nil then Exit;
  if LAGrid.ColLengths[Acol]=-1 then begin result := false; exit; end;
end;

procedure TOF_BTREPARTITION.ZoneSuivanteOuOk(var ACol, ARow: Integer;var Cancel: boolean);
var Sens, ii, Lim: integer;
  OldEna, ChgLig, ChgSens: boolean;
  RowFirst : integer;
begin
	RowFirst := ARow;
  OldEna := LAGrid.SynEnabled;
  LAGrid.SynEnabled := False;
  Sens := -1;
  ChgLig := (LAGrid.Row <> ARow);
  ChgSens := False;
	Lim := LaGrid.RowCount - 1;
  if LAGrid.Row > ARow then Sens := 1 else if ((LAGrid.Row = ARow) and (ACol <= LAGrid.Col)) then Sens := 1;
  ACol := LAGrid.Col;
  ARow := LAGrid.Row;
  ii := 0;
  while not ZoneAccessible(ACol, ARow) do
  begin
    Cancel := True;
    inc(ii);
    if ii > 500 then Break;
    if Sens = 1 then
    begin
      if ((ACol = LAGrid.ColCount - 1) and (ARow >= Lim)) then
      begin
        if ChgSens then Break else
        begin
          Sens := -1;
          Continue;
          ChgSens := True;
        end;
      end;
      if ChgLig then
      begin
        ACol := LAGrid.FixedCols - 1;
        ChgLig := False;
      end;
      if ACol < LAGrid.ColCount - 1 then Inc(ACol) else
      begin
        Inc(ARow);
        ACol := LAGrid.FixedCols;
      end;
    end else
    begin
      if ((ACol = LAGrid.FixedCols) and (ARow = 1)) then
      begin
        if ChgSens then Break else
        begin
          Sens := 1;
          Continue;
        end;
      end;
      if ChgLig then
      begin
        ACol := LAGrid.ColCount;
        ChgLig := False;
      end;
      if ACol > LAGrid.FixedCols then Dec(ACol) else
      begin
        Dec(ARow);
        ACol := LAGrid.ColCount - 1;
      end;
    end;
  end;
  LAGrid.SynEnabled := OldEna;
  if (Arow <> RowFirst) then
  BEGIN
  	if Cancel then LAGridRowEnter (self,RowFirst,cancel,false)
    					else LAGridRowEnter (self,Arow,cancel,false);
  END;
end;

procedure TOF_BTREPARTITION.LAGridDblClick(Sender: Tobject);
var TOBL : TOB;
		Pieceorigine : String;
    Cledoc : r_cledoc;
begin
  TOBL := MyOwnTOB.Detail[LAGrid.row-1];
  if MyOwnTOB = nil then exit;
  PieceOrigine := TOBL.GetValue('GL_PIECEORIGINE');
  if PieceOrigine = '' then exit;
  DecodeRefPiece(PieceOrigine, Cledoc);
  SaisiePiece(Cledoc, taConsult);
end;

procedure TOF_BTREPARTITION.BannulCkick(Sender: Tobject);
begin
  LaTOB.dupliquer (MySavTOB,true,true);
  TheTOB := LaTOB;
  TheTOB.putValue('RETOUR',0);
end;

procedure TOF_BTREPARTITION.BValiderClick(Sender: Tobject);
var Sok : boolean;
		force : boolean;
begin
  Sok := true;
  if MyOwnTOB.getValue('QTEREPART') <> MyOwnTOB.getValue('QTEAREPART') then
  begin
    PgiError ('Vous devez saisir la répartition de la réception.');
    TFVierge(ecran).ModalResult := 0;
    exit;
  end;
  LaTOB.dupliquer (MyOwnTOB,true,true);
  TheTOB := LaTOB;
  TheTOB.putValue('RETOUR',0);
end;

Initialization
  registerclasses ( [ TOF_BTREPARTITION ] ) ; 
end.

