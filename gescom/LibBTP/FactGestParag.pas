unit FactGestParag;

interface
uses sysutils,classes,windows,messages,controls,forms,hmsgbox,stdCtrls,clipbrd,nomenUtil,
     HCtrls,SaisUtil,HEnt1,Ent1,EntGC,UtilPGI,UTOB,HTB97,FactUtil,FactComm,Menus,ParamSoc,
     AglInit,FactTob,FactVariante,vierge,UtilNumParag,
{$IFDEF EAGLCLIENT}
  maineagl,
{$ELSE}
  fe_main,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ENDIF}
     HRichOLE,UTOF;

const MAXITEMS = 2;

type
TActionOnLig = (TaoReplace,TaoInsert);
TModeAction = (TmaVerif,TmaTraitement);

TOF_BINSPARAG = class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
end;

TGestParagraphe = class
  private
    fCreatedPop : boolean;
    FF : TForm;
    fusable : boolean;
    fgrid : THGrid;
    POPGS : TPopupMenu;
    TOBPiece : TOB;
    TOBTiers : TOB;
    TOBAffaire : TOB;
    TOBStructParag : TOB;
    MesMenuItem: array[0..MAXITEMS] of TMenuItem;
    TheMenuItem : TmenuItem;
    fMaxItems : integer;
    NiveauAAppliquer : integer;
    TexteParag : string;
    TTNUMP : TNumParag;

    procedure SetGrid(const Value: THGrid); virtual;
    procedure DefiniMenuPop(Parent: Tform);
    procedure Setdocument(const Value: TOB);
    procedure GSInsertParagraphe (Sender: TObject);
    function InsParagPieceAutorisee : boolean;
    procedure InsertionParagraphe(Mode: integer);
    function AjouteParagraphe(Ligne: integer; ActionSurLigne: TActionOnLig;
      TypeParagraphe: TTypeParagraphe; ModeAction: TModeAction): boolean;
    procedure SetTiers(const Value: TOB);
    procedure SetAffaire(const Value: TOB);
    procedure PositionneLigneInparag(First, Last: integer);
    function AddDebutParag (Indice : integer) : boolean;
    function ControleExistDebutParag (Indice : integer) : boolean;
    function OKStructureParag : boolean;
    Procedure SetUsable (Statut : boolean);
  public
    constructor create (TT : TForm);
    destructor  destroy ; override;
    property    TheNumP : TNumParag read TTNUMP write TTNUMP;
    property    Grid : THGrid write setGrid ;
    property    Piece : TOB read TobPiece write Setdocument;
    property    Tiers : TOB read TobTiers write SetTiers;
    property    Affaire : TOB read Tobaffaire write SetAffaire;
end;

implementation
uses facture,FactPiece,FactureBTP;
{ TGestParagraphe }

function TGestParagraphe.AjouteParagraphe(Ligne : integer;ActionSurLigne : TActionOnLig;
																					TypeParagraphe : TTypeParagraphe; ModeAction : TModeAction) : boolean;
var TOBL : TOB;
		NiveauCourant : integer;
    variante,VarianteCourante : boolean;
    Position : integer;
begin
	result := false;
  TOBL := GEtTOBLigne (TOBPiece,Ligne);
  if (not IsCommentaire (TOBL)) and (ActionSurLigne=TaoReplace) then
  begin
  	PgiBox('On ne peut remplacer une ligne que si elle est de type commentaire');
		result := true;
    exit;
  end;

  if (TypeParagraphe=Ttpdebut) and (ModeAction = TmaVerif) then
  begin
  	NiveauAAppliquer := TOBL.GetValue('GL_NIVEAUIMBRIC');
    if not IsParagraphe(TOBL) then inc(NiveauAAppliquer);
    if NiveauAAppliquer = 0 then NiveauAAppliquer := 1;
    if ActionSurLigne = TaoReplace then TexteParag := TOBL.GEtValue('GL_LIBELLE')
    															 else texteParag := 'Paragraphe';
  end;

  //
  if ModeAction = TmaTraitement then
  begin
  	VarianteCourante := IsVariante (TOBL);
    if (TypeParagraphe = TtpFin) and (ActionSurLigne=TaoInsert) then
    begin
    	Position := Ligne+1;
    end else
    begin
    	Position := Ligne;
    end;

    if (ActionSurLigne = TaoInsert) then
    begin
      InsertTOBLigne (TobPiece,Position);
      TOBL := GEtTOBLigne (TOBPiece,Position);
  		InitialiseTOBLigne(TOBPiece, TOBTiers, TOBAffaire, Position);
    end;

    SetParagraphe (TOBPiece,TOBL,Position,NiveauAAppliquer,variante,TypeParagraphe);

    if (ActionSurLigne = TaoInsert) then
    begin
      if isDebutParagraphe (TOBL) then
      begin
        TOBL.PutValue('GL_LIBELLE',TexteParag);
      end else
      begin
        TOBL.PutValue('GL_LIBELLE','Total '+TexteParag);
      end;
    end;

//    SetTypeLigne (TOBL,VarianteCourante);
  end;

end;

constructor TGestParagraphe.create(TT: TForm);
var ThePop : Tcomponent;
begin
  FF := TT;
  TheMenuItem := nil;
  SetUsable(True);
  TOBStructParag := TOB.Create ('STRUCTURE SELECTION',nil,-1);
{$IFDEF BTP}
  if FF is TFFacture then ThePop := TFFacture(TT).Findcomponent  ('POPBTP') else thePop := nil;
  if ThePop = nil then
  BEGIN
    // pas de menu BTP trouve ..on le cree
    POPGS := TPopupMenu.Create(TT);
    POPGS.Name := 'POPBTP';
    fCreatedPop := true;
  END else
  BEGIN
    fCreatedPop := false;
    POPGS := TPopupMenu(thePop);
  END;
  DefiniMenuPop(TT);
{$ENDIF}
  if TT is TFFActure then  grid := TFFacture(TT).GS;
end;

procedure TGestParagraphe.DefiniMenuPop(Parent: Tform);
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
  // CTRL+ALT+A = inserer dans paragraphe
  MesMenuItem[fmaxitems] := TmenuItem.Create (parent);
  TheMenuItem :=MesMenuItem[fmaxitems];
  with MesMenuItem[fmaxitems] do
    begin
    Caption := TraduireMemoire ('Insérer dans paragraphe');  // par défaut
    Name := 'BINSPARAG';
    OnClick := GSInsertParagraphe;
    enabled := true;
    end;
  MesMenuItem[fmaxitems].ShortCut := ShortCut( Word('A'), [ssCtrl,ssAlt]);
  inc (fmaxitems);

  for Indice := 0 to fmaxitems -1 do
    begin
      if MesMenuItem [Indice] <> nil then POPGS.Items.Add (MesMenuItem[Indice]);
    end;
end;

destructor TGestParagraphe.destroy;
var indice : integer;
begin
  TOBStructParag.free;
  for Indice := 0 to fmaxitems -1 do
  begin
    MesMenuItem[Indice].Free;
  end;
  if fcreatedPop then POPGS.free;
  inherited;
end;

procedure TGestParagraphe.GSInsertParagraphe(Sender: TObject);
var TheParam : TOB;
		Indice : integer;
    lastcol,lastrow : integer;
begin
	LastCol := fgrid.col;
  LastRow := fgrid.row;

	if THgrid(fgrid).nbselected = 0 then
  begin
  	PgiBox('Vous devez faire une sélection de lignes consécutives',FF.caption);
    exit;
  end;

  TOBStructParag.clearDetail;

  TheParam := TOB.Create ('_PARAM_',nil,-1);
  TRY
  	THEPARAM.AddChampSupValeur ('MODE',-1);
  	THEPARAM.AddChampSupValeur ('FROMEXCEL','-');
		if TFfacture(FF).OrigineEXCEL then THEPARAM.putValue ('FROMEXCEL','X');
    TheTOB := THEPARAM;
    AglLanceFiche ('BTP','BTINSPARAG','','','ACTION=MODIFICATION');
    if theTOB <> nil then
    begin
    	if TheTOB.GetValue('MODE') >= 0 then
      begin
      	InsertionParagraphe (TheTOB.GetValue('MODE'));
      end;
    end;
  FINALLY
  	TheParam.free;
    TheTOB := nil;
  END;
  TOBStructParag.clearDetail;
  NumeroteLignesGC (fgrid,TOBPiece,false,false);
	TFFacture(FF).ApresInsertionParagraphe;
  TFFacture(FF).GoToLigne (LastRow,lastcol);
	TFFacture(FF).StCellCur := TFFacture(FF).GS.Cells[LastCol,LastRow];
end;

function TGestParagraphe.InsParagPieceAutorisee: boolean;
begin
	result := true;
//	result := not TFfacture(FF).OrigineEXCEL;
end;

procedure TGestParagraphe.InsertionParagraphe (Mode : integer);
var Indice : integer;
		First,Last : integer;
    erreur : boolean;
    Nbdecal : integer;
begin
	first := -1;
  last := -1;
  Indice := 1;
  erreur := false;
  Nbdecal := 0;
  //
  repeat
    if (fgrid.IsSelected (Indice) ) then
    begin
    	if first = -1 then
      begin
        if (Mode and 1) = 1 then
      	begin
        	// remplacement de la ligne début par le début de paragraphe
          erreur := AjouteParagraphe(Indice,TaoReplace,Ttpdebut,TmaVerif);
        end else
        begin
        	// Insertion d'une debut de paragraphe
          erreur := AjouteParagraphe(Indice,TaoInsert,Ttpdebut,TmaVerif);
        end;
        first := Indice;
      end;
      if IsDebutParagraphe (GetTOBligne(TOBPiece,indice)) then
      begin
      	Erreur := AddDebutParag (Indice);
      end;
      if (IsFinParagraphe (GetTOBligne(TOBPiece,indice))) then
      begin
      	Erreur := not (ControleExistDebutParag (Indice));
      end;
      if erreur then exit; // pas la peine d'aller plus loin
      last := Indice;
    end;
    inc(indice);
  until indice > fgrid.rowCount;

  if not OKStructureParag then exit;
  //
  if Last <> -1 then
  begin
    if (Mode and 2) = 2 then
    begin
      // remplacement de la ligne fin par le fin de paragraphe
      erreur := AjouteParagraphe(last,TaoReplace,Ttpfin,TmaVerif);
    end else
    begin
      // Insertion d'une debut de paragraphe
      erreur := AjouteParagraphe(last,TaoInsert,Ttpfin,TmaVerif);
    end;
    if erreur then exit; // pas la peine d'aller plus loin
  end;

  for indice := First to last do
  begin
    if not (fgrid.IsSelected (Indice) ) then
    begin
    	PgiBox ('les lignes sélectionnées ne sont pas consécutives');
      exit;
    end;
  end;

  // fin de la verif..si c bon ..on continu
  for indice := First to last do
  begin
  	if fgrid.IsSelected(indice) then TFFacture(FF).Deselectionne(indice);
  end;

  // on ajoute dabord la fin de paragraphe pour eviter de decaler les indices
  if (Mode and 2) = 2 then
  begin
    // remplacement de la ligne fin par le fin de paragraphe
    AjouteParagraphe(last,TaoReplace,Ttpfin,TmaTraitement);
  end else
  begin
    // Insertion d'une debut de paragraphe
    AjouteParagraphe(last,TaoInsert,Ttpfin,TmaTraitement);
    inc(Nbdecal);
  end;

  if (Mode and 1) = 1 then
  begin
    // remplacement de la ligne début par le début de paragraphe
    AjouteParagraphe(first,TaoReplace,Ttpdebut,TmaTraitement);
  end else
  begin
    // Insertion d'une debut de paragraphe
    AjouteParagraphe(first,TaoInsert,Ttpdebut,TmaTraitement);
    inc(Nbdecal);
  end;
  TOBPIece.putValue('GP_RECALCULER','X');
  PositionneLigneInparag (First,Last+Nbdecal);
  for Indice := First to Last+NbDecal do
  begin
    TTNUMP.SetInfoLigne (TOBPiece,Indice);
  end;
  fgrid.rowcount := fgrid.rowcount +Nbdecal;
  DefiniRowCount (TFfacture(FF),TOBPiece);
end;


procedure TGestParagraphe.PositionneLigneInparag (First,Last : integer);
var indice,Niveau : integer;
		DateLiv : TdateTime;
    variante : boolean;
    TOBL : TOB;
begin
	TOBL := GetTobLigne(TOBPiece,First);
  DateLiv := TOBL.getValue('GL_DATELIVRAISON');
  variante := isVariante(TOBL);
  for Indice := First to Last do
  begin
    TOBL := GetTOBLigne (TOBpiece,indice);
  	if (IsParagraphe (TOBL)) and ((Indice=First) or (Indice=last)) then
    begin
  		Niveau := TOBL.GetValue('GL_NIVEAUIMBRIC');
    	continue;
    end else
    begin if IsParagraphe (TOBL) then
  		Niveau := TOBL.GetValue('GL_NIVEAUIMBRIC')+1;
    end;

    if IsDebutParagraphe(TOBL) then
    begin
    	SetParagraphe (TOBPiece,TOBL,Indice,niveau,variante,TtpDebut);
    end else if IsFinParagraphe(TOBL) then
    begin
    	SetParagraphe (TOBPiece,TOBL,Indice,niveau,variante,Ttpfin);
    end else
    begin
    	TOBL.putValue('GL_NIVEAUIMBRIC',niveau);
    end;

    TOBL.putValue('GL_DATELIVRAISON',DateLiv);
//    SetTypeLigne (TOBL,Variante);
  end;
end;

procedure TGestParagraphe.Setdocument(const Value: TOB);
var ThisTOb : TOB;
begin
  ThisTOB := Value;

  if ThisTOb = nil then BEGIN SetUsable(False); Exit; END;
  TOBPiece := ThisTOB;
  if not InsParagPieceAutorisee then
  BEGIN
  	SetUsable(false);
    Exit;
  END;
end;

procedure TGestParagraphe.SetGrid(const Value: THGrid);
begin
	fgrid := value;
end;

procedure TGestParagraphe.SetTiers(const Value: TOB);
begin
  TOBTiers := value;
end;

procedure TGestParagraphe.SetAffaire(const Value: TOB);
begin
    TobAffaire := Value;
end;

function TGestParagraphe.AddDebutParag(Indice: integer) : boolean;
var TOBS,TOBL : TOB;
begin
	result := false;
	TOBL := GetTOBLigne (TOBPiece,Indice);
  if TOBL.GetValue('GL_NIVEAUIMBRIC') > 7 then
  begin
  	PgiBox ('Nombre de sous paragraphe atteint..');
    result := true;
  end;
	TOBS := TOB.Create ('STRUCT PARAG',TOBStructParag,-1);
  TOBS.AddChampSupValeur ('NIVEAU',TOBL.GetValue('GL_NIVEAUIMBRIC'));
  TOBS.AddChampSupValeur ('STROK','-');
end;

function TGestParagraphe.ControleExistDebutParag(Indice: integer): boolean;
var TOBL,TOBS,TOBRef : TOB;
begin
	result := true;
	TOBRef := nil;
	TOBL := GetTOBLigne (TOBPiece,Indice);
  TOBS := TOBStructParag.FindFirst (['NIVEAU'],[TOBL.GetValue('GL_NIVEAUIMBRIC')],false);

  repeat
    if TOBS <> nil then
    begin
    	if TOBS.getValue('STROK')='-' then
      begin
        TOBRef := TOBS;
        break;
      end;
      TOBS := TOBStructParag.FindNext (['NIVEAU'],[TOBL.GetValue('GL_NIVEAUIMBRIC')],false);
    end;
  until TOBS = nil;

  if TOBRef = nil then
  begin
  	PgiBox('la sélection est incorrecte');
    result := false;
  end else
  begin
  	TOBref.putValue('STROK','X')
  end;
end;

function TGestParagraphe.OKStructureParag: boolean;
var TOBS : TOB;
		indice : integer;
begin
	Indice := 0;
  result := true;
  if TOBStructParag.detail.count = 0 then exit;
  for indice := 0 to TOBStructParag.detail.count -1 do
  begin
  	TOBS := TOBStructParag.detail[Indice];
    if TOBS.getValue('STROK')='-' then
    begin
      PgiBox('la sélection est incorrecte');
      result := false;
      break;
    end;
  end;
end;

procedure TGestParagraphe.SetUsable(Statut: boolean);
begin
  fUsable := Statut;
  if TheMenuItem <> nil then TheMenuItem.Enabled := Statut;
end;

{ TOF_BINSPARAG }

procedure TOF_BINSPARAG.OnArgument(S: String);
begin
  inherited;
end;

procedure TOF_BINSPARAG.OnCancel;
begin
  inherited;
end;

procedure TOF_BINSPARAG.OnClose;
begin
  inherited;

end;

procedure TOF_BINSPARAG.OnDelete;
begin
  inherited;

end;

procedure TOF_BINSPARAG.OnDisplay;
begin
  inherited;

end;

procedure TOF_BINSPARAG.OnLoad;
begin
  inherited;

end;

procedure TOF_BINSPARAG.OnNew;
begin
  inherited;

end;

procedure TOF_BINSPARAG.OnUpdate;
begin
  inherited;
  TheTOB := laTOB;
  TheTOB.putValue('MODE',0);
  if TcheckBox(getControl('C_FIRST')).Checked then theTOB.putValue('MODE',theTOB.GetValue('MODE')+1);
  if TcheckBox(getControl('C_LAST')).Checked then theTOB.putValue('MODE',theTOB.GetValue('MODE')+2);
  (*
  if (TheTOB.GetValue('FROMEXCEL')='X') and(TheTOB.GetValue('MODE')<>3) then
  begin
  	PgiInfo ('Les lignes de début et de fin de paragraphe doivent être présentes dans le fichier Excel');
    TFVierge(Ecran).ModalResult := 0;
    TcheckBox(getControl('C_FIRST')).SetFocus;
  end;
  *)
end;

Initialization
  registerclasses ( [ TOF_BINSPARAG ] ) ;
end.
