{***********UNITE*************************************************
Auteur  ...... : LS
Créé le ...... : 27/07/2005
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BTTRARTHIER ()
Mots clefs ... : TOF;BTTRARTHIER
*****************************************************************}
Unit BTTRARTHIER_TOF ;

Interface

Uses StdCtrls, 
     Controls,
     Classes,
     grids,
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} 
     mul,
     FE_Main,
{$else}
		 Maineagl,
     eMul, 
     uTob, 
{$ENDIF}
		 AglInit,
     forms,
     sysutils, 
     ComCtrls,
     HCtrls, 
     HEnt1,
     HMsgBox,
     UTOB,
     HTB97,
     UTOF ;

Type
  TOF_BTTRARTHIER = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  private
  	CodeTarifArt : THValComboBox;
    GS : THgrid;
    TOBDetail : TOB;
    LesCols : string;
    CurrentTarif : string;
    CellCur : string;
    EltSuprime : boolean;
    procedure AlloueTobs;
    procedure CodeTarifArtChange (Sender : TObject);
  	procedure ChargeTob;
    procedure GetControls;
    procedure LibereTobs;
    procedure SetEvents;
    function  DemandeEcriture: boolean;
    procedure Ecriture;
    procedure Definigrille;
    procedure ChargeGrille;
    procedure EnabledEventsGs (Active : boolean);
		procedure GSCellEnter (Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
		procedure GSCellExit  (Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
		PROCEDURE GSRowEnter (Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
		procedure GSRowExit  (Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
		procedure GSDblClick  (Sender: TObject);
		procedure SetControls;
    procedure PositionneCell(ACol, Arow: integer);
		procedure BdeleteClick (Sender :TObject);
    procedure AddNewDetail;
		function ExistCode(Code : string;Ligne : integer) : boolean;
    function TOBVide (TOBL : TOB) : boolean;
    function EltModifie: boolean;
    procedure BinsertClick (Sender : TObject);
  end ;

procedure FamilleTarifHierarchique;

Implementation

procedure FamilleTarifHierarchique;
begin
	AGLLanceFiche ('BTP','BTTARARTHIER','','','ACTION=MODIFICATION');
end;

procedure TOF_BTTRARTHIER.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_BTTRARTHIER.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_BTTRARTHIER.OnUpdate ;
begin
  Inherited ;
	Ecriture;
end ;

procedure TOF_BTTRARTHIER.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_BTTRARTHIER.OnArgument (S : String ) ;
begin
  Inherited ;
  EltSuprime := false;
  CurrentTarif := '';
  AlloueTobs;
  GetControls;
	LesCols:='FIXED;CC_TYPE;CC_CODE;CC_LIBELLE;CC_ABREGE;CC_LIBRE' ;
  Definigrille;
  SetEvents;
  SetControls;
end ;

procedure TOF_BTTRARTHIER.OnClose ;
begin
  Inherited ;
  LibereTobs;
  AvertirTable ('BTSOUSFAMTARART');
end ;

procedure TOF_BTTRARTHIER.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_BTTRARTHIER.OnCancel () ;
begin
  Inherited ;
	if ((TOBdetail.Detail.count > 0) and (TOBdetail.IsOneModifie )) or (EltSuprime) then
  begin
  	if DemandeEcriture then Ecriture;
  end;
end ;

function TOF_BTTRARTHIER.EltModifie : boolean;
var Indice : integer;
begin
	result := false;
	if (TOBdetail.Detail.count > 0) then
  begin
  	for indice := 0 to TOBdetail.detail.count -1 do
    begin
    	if (not TOBVide(TOBDetail.detail[Indice])) and (TOBDetail.detail[Indice].IsOneModifie) then
      begin
      	result := true;
        break;
      end;
    end;
  end;
end;

procedure TOF_BTTRARTHIER.CodeTarifArtChange(Sender: TObject);
begin
	if (EltModifie) or (EltSuprime) then
  begin
  	if DemandeEcriture then Ecriture;
  end;
  CurrentTarif := CodeTarifArt.Value;
  ChargeTob;
  EltSuprime := false;
  ChargeGrille;
  SetControls;
end;

procedure TOF_BTTRARTHIER.GetControls;
begin
	CodeTarifArt := THValComboBox(getCOntrol('FAMILLETARIFART'));
  GS := THGrid(getControl('GS'));
end;

procedure TOF_BTTRARTHIER.SetEvents;
begin
  CodeTarifArt.OnChange := CodeTarifArtChange;
  TToolbarButton97 (getControl('Binsert')).onclick := BinsertClick;
end;

procedure TOF_BTTRARTHIER.AlloueTobs;
begin
	TOBDetail := TOB.Create ('LES SOUS FAMILLES',nil,-1);
end;

procedure TOF_BTTRARTHIER.LibereTobs;
begin
	TOBDetail.free;
end;

function  TOF_BTTRARTHIER.DemandeEcriture : boolean;
begin
	result := (PGiAsk ('Désirez-vous sauvegarder les modifications',ecran.caption)=MrYes);
end;

procedure TOF_BTTRARTHIER.Ecriture;
var Indice : integer;
begin
	TOBdetail.SetAllModifie (true);
	//if executeSql ('DELETE FROM CHOIXCOD WHERE CC_TYPE="BFT" AND CC_LIBRE="'+CurrentTarif+'"') >= 0 then
	if executeSql ('DELETE FROM BTSOUSFAMILLETARIF WHERE BSF_FAMILLETARIF="'+CurrentTarif+'"') >= 0 then
  begin
  	for Indice := 0 to TOBdetail.detail.count -1 do
    begin
    	if not TOBVide(TOBdetail.detail[Indice]) then
      begin
      	TOBDetail.detail[Indice].InsertOrUpdateDB (true);
      	TOBDetail.detail[Indice].PutValue('EXISTS','X');
      	TOBDetail.detail[Indice].SetAllModifie (false);
      end;
    end;
  end;
end;


procedure TOF_BTTRARTHIER.ChargeTob;
var QQ : TQuery;
begin
	TOBdetail.ClearDetail;
  //QQ := OpenSql ('SELECT * FROM CHOIXCOD WHERE CC_TYPE="BFT" AND CC_LIBRE="'+CurrentTarif+'"',true,-1,'',true);
  QQ := OpenSql ('SELECT * FROM BTSOUSFAMILLETARIF WHERE BSF_FAMILLETARIF="'+CurrentTarif+'"',true,-1,'',true);
  TRY
    if not QQ.eof then
    begin
      TOBDetail.LoadDetailDB ('CHOIXCOD','','',QQ,false);
      TOBDetail.detail[0].AddChampSupValeur ('EXISTS','X',true);
    end else
    begin
      AddNewDetail;
    end;
  FINALLY
  	ferme (QQ);
  END;
end;

procedure TOF_BTTRARTHIER.Definigrille;
begin
  GS.ColCount := 6;
  GS.ColWidths[0]:=8;
  // CC_TYPE
  GS.ColWidths[1]:=0;
  GS.ColLengths[1]:=-1;
  // CC_CODE
  GS.Cells[2,0] := Traduirememoire('Code');
  GS.ColWidths[2]:=50;
  GS.ColFormats[2]:='UPPER';
  GS.ColAligns[2]:=taLeftJustify ;
  GS.ColLengths[2]:=3;
  // CC_LIBELLE
  GS.Cells[3,0] := Traduirememoire('Désignation');
  GS.ColWidths[3]:=150;
  GS.ColAligns[3]:=taLeftJustify ;
  GS.ColLengths[3]:=35;
  // CC_ABREGE
  GS.Cells[4,0] := Traduirememoire('Abrégé');
  GS.ColWidths[4]:=80;
  GS.ColAligns[4]:=taLeftJustify ;
  GS.ColLengths[4]:=17;
  // CC_LIBRE
  GS.ColWidths[5]:=0;
  GS.ColLengths[5]:=-1;
end;

procedure TOF_BTTRARTHIER.ChargeGrille;
var cancel : boolean;
		Acol,Arow : integer;
begin
	EnabledEventsGs(False);
	GS.VidePile (false);
	if TOBdetail.detail.count > 0 then GS.RowCount := TOBdetail.Detail.Count + 1
  															else GS.RowCount := 2;

	TOBDetail.PutGridDetail (GS,false,false,LesCols);
  Arow := 1; Acol := 1;
  GSRowEnter (self,Arow,cancel,false);
	PositionneCell (Acol,Arow);
  GSCellEnter (self,Acol,Arow,cancel);
  GS.row := Arow;
  GS.col := Acol;
	EnabledEventsGs(True);
  GS.ShowEditor;
end;

procedure TOF_BTTRARTHIER.EnabledEventsGs(Active: boolean);
begin
	if Active then
  begin
    GS.OnCellEnter := GSCellEnter;
    GS.OnCellExit := GSCellExit;
    GS.OnRowEnter := GSRowEnter;
    GS.OnRowExit  := GSRowExit;
    GS.OnDblClick := GSDblClick;
  end else
  begin
    GS.OnCellEnter := nil;
    GS.OnCellExit := nil;
    GS.OnRowEnter := nil;
    GS.OnRowExit  := nil;
    GS.OnDblClick := nil;
  end;
end;

procedure TOF_BTTRARTHIER.GSCellEnter(Sender: TObject; var ACol,ARow: Integer; var Cancel: Boolean);
begin
  CellCur := GS.cells[GS.col,GS.row];
end;

procedure TOF_BTTRARTHIER.GSCellExit(Sender: TObject; var ACol,ARow: Integer; var Cancel: Boolean);
var TOBL : TOB;
begin
	if Cellcur = GS.cells[Acol,Arow] then exit;
  TOBL := TOBDetail.detail[Arow-1];
  if Acol = 2 then // Code
  begin
  	if TOBL.GetValue('EXISTS')='X' then
    begin
    	GS.Cells[Acol,Arow] := TOBL.getValue('CC_CODE');
    end else
    begin
    	if not ExistCode(Gs.cells[Acol,Arow],Arow) then
      begin
    		TOBL.putValue('CC_CODE',GS.cells[Acol,Arow]);
      end else
      begin
      	PgiBox('Ce code existe déjà',ecran.caption);
        GS.cells[Acol,Arow] := '';
        Cancel := true;
      end;
    end;
  end else if Acol = 3 then
  begin
  	TOBL.putValue('CC_LIBELLE',GS.cells[Acol,Arow]);
  end else if Acol = 4 then
  begin
  	TOBL.putValue('CC_ABREGE',GS.cells[Acol,Arow]);
  end;
end;

procedure TOF_BTTRARTHIER.GSRowEnter(Sender: TObject; Ou: Integer;var Cancel: Boolean; Chg: Boolean);
begin
end;

procedure TOF_BTTRARTHIER.GSRowExit(Sender: TObject; Ou: Integer;var Cancel: Boolean; Chg: Boolean);
begin
end;

procedure TOF_BTTRARTHIER.SetControls;
begin
  TToolbarButton97(GetControl('BDelete')).Enabled := (TOBDetail.detail.count > 0);
  TToolbarButton97(GetControl('BDelete')).Onclick := BdeleteClick;
end;

procedure TOF_BTTRARTHIER.GSDblClick(Sender: TObject);
begin
end;

procedure TOF_BTTRARTHIER.PositionneCell(ACol,Arow : integer);
var LastMode : boolean;
begin
	if GoRowSelect in GS.options then LastMode := false else LastMode := true;
	GS.Col := 1;
  GS.row := Arow;
end;

procedure TOF_BTTRARTHIER.BdeleteClick(Sender: TObject);
var TOBL : TOB;
		cancel : boolean;
    Acol,Arow : integer;
begin
  if PGIAsk ('Désirez-vous supprimer cette sous famille ?',ecran.caption) = MrYes then
  begin
    TOBDetail.detail[GS.row-1].free;
    EltSuprime := true;
    EnabledEventsGs (false);
    if TOBdetail.detail.count = 0 then AddNewDetail;
    gs.rowcount := TOBdetail.Detail.count+1;
    TOBDetail.PutGridDetail (GS,false,false,LesCols);
    GSRowEnter(self,GS.row,cancel,false);
    Arow := GS.row;
    Acol := 1;
    GSCellEnter (self,Acol,Arow,cancel);
    EnabledEventsGs (true);
  end;
end;

procedure TOF_BTTRARTHIER.AddNewDetail;
var TOBL : TOB;
begin
  TOBL := TOB.create ('CHOIXCOD',Tobdetail,-1);
  TOBL.PutValue('CC_TYPE','BFT');
  TOBL.PutValue('CC_LIBRE',CurrentTarif);
  TOBL.AddChampSupValeur ('EXISTS','-');
end;

function TOF_BTTRARTHIER.ExistCode(Code: string; Ligne: integer): boolean;
var TOBL,TOBCMP : TOB;
		QQ : TQuery;
begin
	result := True;
  if Code = '' then exit;
  TOBL := TOBDetail.detail[Ligne-1];
  // contrainte 1
  TOBCMP := TOBDetail.findfirst(['CC_CODE'],[Code],true);
  if (TOBCMP <> nil) and (TOBL = TOBCMP) then exit;  // il existe deja et pas sur la mm ligne
  //QQ := OpenSql ('SELECT CC_CODE FROM CHOIXCOD WHERE CC_TYPE="BFT" AND CC_CODE="'+Code+'" AND CC_LIBRE<>"'+CurrentTarif+'"',true,-1,'',true);
  QQ := OpenSql ('SELECT BSF_SOUSFAMTARART FROM BTSOUSFAMILLETARIF WHERE BSF_SOUSFAMTARART="'+Code+'" AND BSF_FAMILLETARIF<>"'+CurrentTarif+'"',true,-1,'',true);

  result := (not QQ.eof);
  ferme (QQ);
end;

function TOF_BTTRARTHIER.TOBVide(TOBL: TOB): boolean;
begin
  result := (TOBL.GetValue('CC_CODE')='');
end;

procedure TOF_BTTRARTHIER.BinsertClick(Sender: TObject);
var cancel : boolean;
		Acol,Arow : integer;
begin
  if TOBVide(TOBDetail.detail[TObdetail.detail.count-1]) then exit;
  AddNewDetail;
  gs.rowcount := TOBdetail.Detail.count+1;
  TOBDetail.PutGridDetail (GS,false,false,LesCols);
  Gs.row := Gs.rowCount-1;
  GSRowEnter(self,GS.row,cancel,false);
  Arow := GS.row;
  Acol := 1;
  GSCellEnter (self,Acol,Arow,cancel);
  EnabledEventsGs (true);
end;

Initialization
  registerclasses ( [ TOF_BTTRARTHIER ] ) ;
end.

