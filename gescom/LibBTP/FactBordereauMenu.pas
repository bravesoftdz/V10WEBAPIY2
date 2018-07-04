unit FactBordereauMenu;

interface

uses sysutils,classes,windows,messages,controls,forms,hmsgbox,stdCtrls,clipbrd,nomenUtil,
     HCtrls,SaisUtil,HEnt1,Ent1,EntGC,UtilPGI,UTOB,HTB97,FactUtil,FactComm,Menus,ParamSoc,
{$IFDEF EAGLCLIENT}
  maineagl,
{$ELSE}
  fe_main,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ENDIF}
     HRichOLE;

const MAXITEMS = 2;

type

	TFactBordereauMenu = class
    fCreatedPop : boolean;
    FF : TForm;
    fusable : boolean;
    fgrid : THGrid;
    POPGS : TPopupMenu;
    MesMenuItem: array[0..MAXITEMS] of TMenuItem;
    fMaxItems : integer;
    fTOBL,fTOBPiece : TOB;
  private
    procedure SetGrid(const Value: THGrid); virtual;
    procedure GSPrendCeDetail (Sender : Tobject);
    procedure DefiniMenuPop(Parent: Tform);
    procedure SetTOBPiece(const Value: TOB);
    function IsAvecSousDetail(TypeArticle: string): boolean;
    function IsCeDetailFromBordereau: boolean;
    procedure SetLigne(TOBLig: TOB);
  public
    constructor create (TT : TForm);
    destructor  destroy ; override;
    procedure Activate;
    property 		Grid : THGrid read Fgrid write SetGrid;
    property    Piece : TOB read fTOBPiece write SetTOBPiece;
    property    Ligne : TOB read fTOBL write SetLIgne;
  end;

implementation

uses facture,FactTOB
{$IFDEF BTP}
     ,BTPUtil,FactureBtp
{$ENDIF}
;

{ TFactBordereauMenu }

constructor TFactBordereauMenu.create(TT: TForm);
begin
  FF := TT;
end;

procedure TFactBordereauMenu.DefiniMenuPop(Parent: Tform);
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
  // CTRL+ALT+B = Prends ce sous détail dans pièces
  MesMenuItem[fmaxitems] := TmenuItem.Create (parent);
  with MesMenuItem[fmaxitems] do
  begin
    Caption := TraduireMemoire ('Prendre ce sous détail');  // par défaut
    Name := 'BGETCEDETAIL';
    OnClick := GSPrendCeDetail;
    enabled := false;
  end;
  MesMenuItem[fmaxitems].ShortCut := ShortCut( Word('B'), [ssCtrl,ssAlt]);
  inc (fmaxitems);

  for Indice := 0 to fmaxitems -1 do
  begin
  	if MesMenuItem [Indice] <> nil then POPGS.Items.Add (MesMenuItem[Indice]);
  end;
end;

destructor TFactBordereauMenu.destroy;
var indice : integer;
begin
  for Indice := 0 to fmaxitems -1 do
  begin
    MesMenuItem[Indice].Free;
  end;
{$IFDEF BTP}
  if fcreatedPop then POPGS.free;
{$ENDIF}
  inherited;
end;

procedure TFactBordereauMenu.GSPrendCeDetail(Sender: Tobject);
var TOBL : TOB;
		Arow : integer;
begin
  if not fusable then exit;
  Arow := fgrid.row;
  TOBL := GetTOBLigne (fTOBPiece,Arow);
  if (TOBL.GetValue('GLC_GETCEDETAIL')<> 'X') then TOBL.PutValue('GLC_GETCEDETAIL','X')
  																						 else TOBL.PutValue('GLC_GETCEDETAIL','-');
  fgrid.InvalidateRow (fgrid.row);
  TFFacture(FF).AfficheLaLigne (Fgrid.row);
  SetLigne (TOBL);
end;

procedure TFactBordereauMenu.SetGrid(const Value: THGrid);
begin
	fgrid := Value;
{$IFDEF BTP}
  if fCreatedPop then if not assigned(fgrid.popupmenu) then fgrid.PopupMenu := POPGS;
{$ENDIF}
end;

function TFactBordereauMenu.IsAvecSousDetail (TypeArticle : string): boolean;
begin
	Result :=  (TypeArticle='OUV') or (TypeArticle = 'ARP');
end;

function TFactBordereauMenu.IsCeDetailFromBordereau : boolean;
begin
	result :=  fTOBL.GetValue('GLC_GETCEDETAIL')='X';
end;

procedure TFactBordereauMenu.SetLigne(TOBLig : TOB);
var CurItem,TheItem : TmenuItem;
    indice : integer;
    TypeArticle : string;
begin
	if not fusable then exit;
  TheItem := nil;
	fTOBL := TOBLig; if fTOBL = nil then exit;
	TypeArticle := fTOBL.getValue('GL_TYPEARTICLE');

  {$IFDEF BTP}
  if POPGS <> nil then
  begin
    for indice := 0 to POPGS.Items.Count do
    begin
      CurItem := POPGS.Items [indice];
      if CurItem.Name = 'BGETCEDETAIL' then BEGIN TheItem := CurItem; Break; END;
    end;
  end;
  {$ENDIF}
  if TheItem = nil then exit;
  if not IsAvecSousDetail (TypeArticle) or not fusable then TheItem.Enabled := false
  																							 			 else TheItem.Enabled := true;
  if not fusable then exit;
  if TypeArticle = 'OUV' then
  begin
    if IsCeDetailFromBordereau then TheItem.caption := TraduireMemoire ('Prendre le sous détail de la bibliothèque')
    															else TheItem.caption := TraduireMemoire ('Prendre ce sous détail');
  end else
  begin
    if IsCeDetailFromBordereau then TheItem.caption := TraduireMemoire ('Prendre cet article dans la bibliothèque')
                                  else TheItem.caption := TraduireMemoire ('Prendre cet article dans le bordereau');
  end;
end;

procedure TFactBordereauMenu.SetTOBPiece(const Value: TOB);
begin
  fTOBPiece := Value;
  if fTOBPiece = nil then BEGIN Fusable := False; Exit; END;
end;

procedure TFactBordereauMenu.Activate;
var ThePop : Tcomponent;
begin
  ThePop := nil;
  if not (tModeSaisieBordereau in TFFacture(FF).SaContexte) then BEGIN Fusable := false; exit; END;
  fusable := true;
{$IFDEF BTP}
  if FF is TFFacture then ThePop := TFFacture(FF).Findcomponent  ('POPBTP');
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
{$ENDIF}
  if FF is TFFActure then  grid := TFFacture(FF).GS;
end;

end.
