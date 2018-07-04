{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 04/04/2011
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BTAFFAIREFRSGEST ()
Mots clefs ... : TOF;BTAFFAIREFRSGEST
*****************************************************************}
Unit BTAFFAIREFRSGEST_TOF ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul,
     fe_main,
{$else}
     eMul,
     MainEagl,
{$ENDIF}
     uTob,
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     vierge,
     windows,
     LookUp,
     UTOF,
     HTB97 ;

Type

  TOF_BTAFFAIREFRSGEST = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  private
  	faction : TActionFiche;
  	faffaire : string;
  	fTOBFrsEnt,fTOBFrsDet : TOB;
    FTOBEnt,fTOBL : TOB;
  	GE,GS : THgrid;
    BInsertDet,BdeleteDet,Binsert,Bdelete : TToolBarButton97;
    //
    fColNamesGE,fColNamesGS : string;
    Ftitre : Hstring;
    FLargeur, FAlignement: string;
    //
    GSSEL,GSTIERS,GSLIB,GSPVHT,GSPVTTC,GSCOEFF : integer;
    stprev : string;
    //
    procedure GetComponents;
    procedure DefiniGrille(Grid: Thgrid; ColNames, Aligments, Titles, Widths: string);
    procedure DefinigridSaisieGS(fColNames: string);
    procedure AffichelaGrille(UneTOB: TOB; Grid: Thgrid; ColNames: string;
      Entete: boolean=false);
    procedure AfficheLaLigne(UneTOB: TOB; Indice: integer; grid: THgrid;
      ColNames: string; Entete: boolean=false);
    procedure ChargeLaTOBEnt;
  	procedure SetEventsgrid;
    procedure GEEnter(Sender: TObject);
    procedure GERowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean;
      Chg: Boolean);
    procedure GERowExit(Sender: TObject; Ou: Integer; var Cancel: Boolean;
      Chg: Boolean);
    procedure GSCellEnter(Sender: TObject; var ACol, ARow: Integer;
      var Cancel: Boolean);
    procedure GSCellExit(Sender: TObject; var ACol, ARow: Integer;
      var Cancel: Boolean);
    procedure GSRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean;
      Chg: Boolean);
    procedure GSRowExit(Sender: TObject; Ou: Integer; var Cancel: Boolean;
      Chg: Boolean);
    procedure GSEnter(Sender: TObject);
    procedure AfficheEnt;
    procedure DefiniGrilleGE;
    procedure DefiniGrilleGS;
    procedure SetEventsGS(Active: boolean);
    procedure PositionneVisibleGS;
    function GetTOBEnt(Ou: integer): TOB;
    function GetTOBLig(Ou: integer): TOB;
    procedure ChargelesLignes;
    procedure AffichelesLignes;
    procedure DecodeArgs(S: string);
    procedure SetEvents;
		procedure BInsertClick (Sender : Tobject);
		procedure BInsertDetClick (Sender : Tobject);
  	procedure BDeleteClick (Sender : Tobject);
  	procedure BDeleteDetClick (Sender : Tobject);
    procedure AjouteFraisgestion (TheCode : string);
    function FindCodePort (TheCode : string) : boolean;
    procedure ZoneSuivanteOuOk(Grille: THgrid; var ACol, ARow: integer;var Cancel: boolean);
    function  ZoneAccessible(Grille: THgrid; var ACol,ARow: integer): boolean;
    function FindInterv(CodeInter: string): boolean;
    function AjouteFrsInterv(CodeInter: string): TOB;
    procedure FormKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);

  end ;

Implementation

uses Grids;

procedure TOF_BTAFFAIREFRSGEST.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_BTAFFAIREFRSGEST.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_BTAFFAIREFRSGEST.OnUpdate ;
begin
  Inherited ;
  fTOBFrsDet.InsertOrUpdateDB(false);
  fTOBFrsDet.PutValue('MODIFIED','-');
  FTOBEnt.PutValue('NEW','-');
end ;

procedure TOF_BTAFFAIREFRSGEST.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_BTAFFAIREFRSGEST.OnArgument (S : String ) ;
var cancel : boolean;
begin
  Inherited ;
	DecodeArgs (S);
  fTOBFrsEnt := TOB.create ('LES ENTETES',nil,-1);
  fTOBFrsDet := TOB.create ('LES DETAILS',nil,-1); fTOBFrsDet.AddChampSupValeur('MODIFIED','-');
  GetComponents;
  SetEvents;
  ChargeLaTOBEnt;
  AfficheEnt;
  SetEventsgrid;
  GEENter(self);
  GERowEnter(self,1,cancel,false);
  TFvierge(ecran).OnKeyDown := FormKeyDown; 
end ;

procedure TOF_BTAFFAIREFRSGEST.OnClose ;
begin
	if (ftobent <> nil) and ((FTOBEnt.getValue('NEW')='X') or (fTOBFrsDet.getValue('MODIFIED')='X')) then
  begin
  	if PgiAsk ('Désirez-vous sauvegarder les modifications ?')=mryes then fTOBFrsDet.InsertOrUpdateDB(false); 
  end;
	freeAndNil(fTOBFrsEnt);
  FreeAndnil(fTOBFrsDet);
  Inherited ;
end ;

procedure TOF_BTAFFAIREFRSGEST.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_BTAFFAIREFRSGEST.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_BTAFFAIREFRSGEST.GetComponents;
begin
	GE := Thgrid(GetControl('GE'));
	GS := Thgrid(GetControl('GS'));
  BInsertDet := TToolbarBUtton97(getControl('BINSERTDET'));
  BDeleteDet := TToolbarBUtton97(getControl('BDELETEDET'));
  BInsert := TToolbarBUtton97(getControl('BINSERT'));
  BDelete := TToolbarBUtton97(getControl('BDELETE'));
end;

procedure TOF_BTAFFAIREFRSGEST.DefiniGrilleGE;
begin
  // Définition de la list de saisie pour la grille Entete
  fColNamesGE := 'SEL;BAF_CODEPORT;GPO_LIBELLE;GPO_TYPEPORT;';
  Falignement := 'D.0  ---;G.0  ---;G.0  ---;C.0 $---;';
  Ftitre := 'Num;Code;Désignation;Type;';
  fLargeur := '5;18;60;35;';
  //
  DefiniGrille (GE,fColNamesGE,Falignement,Ftitre,fLargeur);
  //
end;

procedure TOF_BTAFFAIREFRSGEST.DefiniGrilleGS;
begin
  // Définition de la liste de saisie pour la grille Détail
  fColNamesGS := 'SEL;BAF_TIERS;LIBTIERS;BAF_PVTTC;BAF_COEFF';
  Falignement := 'D.0  ---;G.0  ---;G.0  ---;D.2  ---;D.0  ---;';
  Ftitre := 'Num;Cotraitant;Raison sociale;Montant TTC; % TTC;';
  fLargeur := '5;18;60;12;8;';
  //
  DefiniGrille (GS,fColNamesGS,Falignement,Ftitre,fLargeur);
  DefinigridSaisieGS(fColNamesGS);
  //
  PositionneVisibleGS;
end;

procedure TOF_BTAFFAIREFRSGEST.DefinigridSaisieGS(fColNames : string);
var ColNames,ColName : string;
		Ind : integer;
begin
  ColNames := fColNames;
  Ind := 0;
  repeat
  	ColName := READTOKENST(ColNames);
    if colName <> '' then
    begin
      if colname = 'SEL' then GSSEL := Ind
      else if colname = 'BAF_TIERS' then GSTIERS := ind
      else if colname = 'LIBTIERS' then GSLIB := ind
      else if colname = 'BAF_PVHT' then GSPVHT := ind
      else if colname = 'BAF_PVTTC' then GSPVTTC := ind
      else if colname = 'BAF_COEFF' then GSCOEFF := ind;
      inc(Ind);
    end;
	until ColName = '';
end;

procedure TOF_BTAFFAIREFRSGEST.DefiniGrille (Grid : Thgrid; ColNames,Aligments, Titles, Widths: string);
var st,lestitres,lesalignements,FF,alignement,Nam,leslargeurs,lalargeur,letitre,lelement : string;
    Obli,OkLib,OkVisu,OkNulle,OkCumul,Sep,Okimg : boolean;
    dec,NbCols : integer;
    indice,ind : integer;
    FFQTE,FFDEC : string;
begin
  //
  st := ColNames;
  NbCols := 0;
  repeat
    lelement := READTOKENST (st);
    if lelement <> '' then
    begin
      inc(NbCols);
    end;
  until lelement = '';
  //
  Grid.ColCount := Nbcols;
  //
  st := ColNames;
  lesalignements := Aligments;
  lestitres := Titles;
  leslargeurs := Widths;

  for indice := 0 to Nbcols -1 do
  begin
    Nam := ReadTokenSt (St); // nom
    alignement := ReadTokenSt(lesalignements);
    lalargeur := readtokenst(leslargeurs);
    letitre := readtokenst(lestitres);
    TransAlign(alignement,FF,Dec,Sep,Obli,OkLib,OkVisu,OkNulle,OkCumul) ;
    grid.cells[Indice,0] := leTitre;
    grid.ColNames [Indice] := Nam;
    if copy(Alignement,1,1)='G' then grid.ColAligns[indice] := taLeftJustify
    else if copy(Alignement,1,1)='D' then grid.ColAligns[indice] := taRightJustify
    else if copy(Alignement,1,1)='C' then grid.ColAligns[indice] := taCenter;
    okImg := (copy(Alignement,8,1)='X');
    if OkVisu then
    begin
  		grid.ColWidths[indice] := strtoint(lalargeur)*GS.Canvas.TextWidth('W');
    end else
    begin
    	grid.ColWidths[indice] := -1;
    end;

    if (OkLib) or (okImg) then
    begin
    	grid.ColFormats[indice] := 'CB=' + Get_Join(Nam);
      if OkImg then
      begin
      	grid.ColDrawingModes[Indice]:= 'IMAGE';
      end;
    end else if (Dec<>0) or (Sep) then
    begin
      FFQTE := '#';
      if Dec > 0 then
      begin
        FFQTE := FFQTE+'0.';
        for ind := 1 to Dec do
        begin
          FFQTE := FFQTE + '0';
        end;
      end else FFQTE := FFQTE + '0';
    	if OkNulle then
      begin
        //
				FFDEC := FFQTE+';-'+FFQTE+'; ;';
      	GS.ColFormats[indice] := FFDEC ;
      end else
      begin
      	GS.ColFormats[indice] := FFQTE ;
      end;
    end;
  end ;
end;

procedure TOF_BTAFFAIREFRSGEST.AfficheEnt;
begin
  GE.VidePile(false);
  DefiniGrilleGE;
  if fTOBFrsEnt.detail.count = 0 then GE.rowCount := 2 else GE.rowCount := fTOBFrsEnt.detail.count+1;
  AffichelaGrille (fTOBFrsEnt,GE,fColNamesGE ,false);
end;

procedure TOF_BTAFFAIREFRSGEST.AffichelesLignes;
begin
  GS.VidePile(false);
  DefiniGrilleGS;
  if fTOBFrsDet.detail.count = 0 then GS.rowCount := 2 else GS.rowCount := fTOBFrsDet.detail.count+1;
  AffichelaGrille (fTOBFrsDet,GS,fColNamesGS ,false);
end;

procedure TOF_BTAFFAIREFRSGEST.AffichelaGrille (UneTOB : TOB; Grid : Thgrid; ColNames: string; Entete : boolean=false);
var Indice : integer;
begin
	for Indice := 0 to UneTOB.detail.Count -1 do
  begin
  	AfficheLaLigne (UneTOB.detail[Indice],Indice,grid,ColNames, Entete);
  end;
  TFVierge(ecran).HMTrad.ResizeGridColumns (grid);
end;

procedure TOF_BTAFFAIREFRSGEST.AfficheLaLigne (UneTOB: TOB; Indice : integer; grid : THgrid; ColNames : string; Entete : boolean=false);
begin
	UneTOB.PutLigneGrid(grid,Indice+1,Entete,Entete,ColNames);
end;

procedure TOF_BTAFFAIREFRSGEST.ChargeLaTOBEnt;
var QQ : Tquery;
		Sql : string;
    indice : integer;
begin
  TRY
    Sql := 'SELECT DISTINCT " " AS SEL,BAF_AFFAIRE,BAF_CODEPORT,GPO_TYPEPORT,GPO_LIBELLE,GPO_PVHT,GPO_PVTTC,GPO_COEFF '+
           'FROM AFFAIREFRSGEST '+
           'LEFT JOIN PORT ON BAF_CODEPORT=GPO_CODEPORT '+
           'WHERE BAF_AFFAIRE="'+fAffaire+'" ';
    //
    QQ := OpenSql (Sql,true,-1,'',true);
    if not QQ.eof then
    begin
      fTOBFrsEnt.LoadDetailDB('ONE TOB','','',QQ,false);
      for Indice := 0 to fTOBFrsEnt.detail.count -1 do
      begin
      	fTOBFrsEnt.detail[Indice].AddChampSupValeur('NEW','-');
      end;
    end;
  FINALLY
    ferme (QQ);
  END;
end;

procedure TOF_BTAFFAIREFRSGEST.SetEvents;
begin
	Binsert.OnClick := BInsertClick;
	BInsertDet.OnClick := BInsertDetClick;
  Bdelete.OnClick := BDeleteClick;
  BdeleteDet.OnClick := BDeleteDetClick;
end;

procedure TOF_BTAFFAIREFRSGEST.SetEventsgrid;
begin
	GE.OnEnter := GEEnter;
	GE.OnRowEnter := GERowEnter;
	GE.OnRowExit := GERowExit;
  GS.OnEnter := GSEnter;
end;

procedure TOF_BTAFFAIREFRSGEST.SetEventsGS(Active : boolean);
begin
  if active then
  begin
    GS.OnRowEnter := GSRowEnter;
    GS.OnRowExit := GSRowExit;
    GS.OnCellEnter := GSCellEnter;
    GS.OnCellExit := GSCellExit;
  end else
  begin
		GS.OnRowEnter := nil;
    GS.OnRowExit := nil;
    GS.OnCellEnter := nil;
    GS.OnCellExit := nil;
  end;
end;

procedure TOF_BTAFFAIREFRSGEST.GEEnter (Sender : TObject);
begin
	SetEventsGS(False);
  BInsertDet.Visible := false;
  BDeleteDet.Visible := false;
  BInsert.Visible := true;
  BDelete.Visible := true;
end;

procedure TOF_BTAFFAIREFRSGEST.GERowEnter (Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
begin
  FTOBEnt := GetTOBEnt(Ou);
  ChargelesLignes;
  AffichelesLignes;
  GS.row := 1;
  GS.Col := 1;
  GSEnter  (self);
  GE.setfocus;
  GEEnter  (self);
end;


procedure TOF_BTAFFAIREFRSGEST.GERowExit (Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
begin
	// verif si changement dans les lignes et proposition de l'enregistrement si c'est le cas
  if fTOBFrsDet.getValue('MODIFIED')='X' then
  begin
  	if PGIAsk('Désirez-vous sauvegarder les modifications ?',ecran.caption)=mryes then
    begin
    	fTOBFrsDet.InsertOrUpdateDB(false);
      FTOBEnt.putValue('NEW','-');
    end else
    begin
    	if fTOBENt.getValue('NEW')='X' then
      begin
      	fTOBfrsDet.ClearDetail; fTOBfrsDet.putvalue('MODIFIED','-');
        fTOBent.free;
        if GE.rowCount = 1 then
        begin
        	AfficheEnt;
        end else
        begin
        	GE.DeleteRow(GE.rowCount -1);
        end;
      end;
    end;
  end;
end;

procedure TOF_BTAFFAIREFRSGEST.GSEnter (Sender : TObject);
var cancel : boolean;
		Acol,Arow : integer;
begin
	SetEventsGS(true);
  BInsertDet.Visible := true;
  BDeleteDet.Visible := true;
  BInsert.Visible := false;
  BDelete.Visible := false;
  GSRowEnter(Self,GS.row,cancel,false);
  Acol := GS.col;
  Arow := GS.row;
  GSCellEnter(self,Acol,Arow,cancel);
  GS.row := Arow;
  GS.col := Acol;
end;

procedure TOF_BTAFFAIREFRSGEST.GSRowEnter (Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
begin
	fTOBL := GetTOBLig (ou);
end;

procedure TOF_BTAFFAIREFRSGEST.GSRowExit (Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
begin
end;

procedure TOF_BTAFFAIREFRSGEST.GSCellEnter(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
begin
  if fAction = taConsult then Exit;
  ZoneSuivanteOuOk(GS,ACol, ARow, Cancel);
  if not Cancel then
  begin
    stprev := GS.Cells [Acol,Arow];
  end;
end;

procedure TOF_BTAFFAIREFRSGEST.GSCellExit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
var TOBL : TOB;
		mode : string;
begin
  if fAction = taConsult then Exit;
  if (stPrev <> GS.cells[Acol,Arow]) then
  begin
    TOBL := GetTOBLig (Arow);
    //
    Mode := fTOBENT.getValue('GPO_TYPEPORT');
    //
    if (ACol = GSPVHT) and (Mode = 'MT')  then
    Begin
    	TOBL.putValue('BAF_PVHT',VALEUR(GS.cells[Acol,Arow]));
    	TOBL.putValue('BAF_PVTTC',0);
    	TOBL.putValue('BAF_COEFF',0);
      fTOBFrsDet.putValue('MODIFIED','X');
    end else if (ACol = GSPVTTC) and (Mode = 'MT') then
    begin
    	TOBL.putValue('BAF_PVTTC',VALEUR(GS.cells[Acol,Arow]));
    	TOBL.putValue('BAF_PVHT',0);
    	TOBL.putValue('BAF_COEFF',0);
      fTOBFrsDet.putValue('MODIFIED','X');
    end else if (ACol = GSCOEFF) and (Mode = 'PT') then
    begin
    	TOBL.putValue('BAF_COEFF',VALEUR(GS.cells[Acol,Arow]));
    	TOBL.putValue('BAF_PVHT',0);
    	TOBL.putValue('BAF_PVTTC',0);
      fTOBFrsDet.putValue('MODIFIED','X');
    end;
    if not cancel then
    begin
			AfficheLaLigne (TOBL,ARow-1,GS,fColNamesGS ,false);
      stPrev := GS.cells[Acol,Arow];
    end;
  end;

end;


procedure TOF_BTAFFAIREFRSGEST.PositionneVisibleGS;
var TypePort : string;
begin
  if fTOBEnt = nil then TypePort := 'MT'
  								 else TypePort := fTOBENT.getValue('GPO_TYPEPORT');
  GS.ColEditables [GSTIERS] := false;
  GS.ColEditables [GSLIB] := false;
  if  TypePort = 'PT' then
  begin
  	// pourcentage du TTC
		GS.ColWidths [GSPVHT] := -1; GS.ColEditables [GSPVHT] := false;
    GS.ColWidths [GSPVTTC] := -1; GS.ColEditables [GSPVTTC] := false;
    GS.ColEditables [GSCOEFF] := true;
  end else
  begin
  	// Montant HT ou TTC
		GS.ColWidths [GSCOEFF] := -1; GS.ColEditables [GSCOEFF] := false;
		GS.ColEditables [GSPVHT] := true;
    GS.ColEditables [GSPVTTC] := true;
  end;
end;

function TOF_BTAFFAIREFRSGEST.GetTOBEnt(Ou : integer) : TOB;
begin
	result := nil;
	if Ou > fTOBFrsEnt.detail.count then exit;
	result := fTOBFrsEnt.detail[Ou-1];
end;

function TOF_BTAFFAIREFRSGEST.GetTOBLig(Ou : integer) : TOB;
begin
	result := nil;
	if Ou > fTOBFrsDet.detail.count then exit;
	result := fTOBFrsDet.detail[Ou-1];
end;

procedure  TOF_BTAFFAIREFRSGEST.ChargelesLignes ;
var QQ: TQuery;
		Sql : String;
begin
	fTOBFrsDet.clearDetail; fTOBFrsDet.putValue('MODIFIED','-');
  if fTOBEnt = nil then exit;
  Sql := 'SELECT AFFAIREFRSGEST.*,T_LIBELLE AS LIBTIERS,"" AS SEL FROM AFFAIREFRSGEST '+
  			 'LEFT JOIN TIERS ON BAF_TIERS=T_TIERS AND T_NATUREAUXI="FOU" '+
  			 'WHERE BAF_AFFAIRE="'+fAffaire+'" '+
  			 'AND BAF_CODEPORT="'+FTOBEnt.getValue('BAF_CODEPORT')+'"';
  QQ := OpenSql (Sql,true,-1,'',true);
  if not QQ.eof then
  begin
    fTOBFrsDet.LoadDetailDB('AFFAIREFRSGEST','','',QQ,false);
  end;
  ferme (QQ);
end;

procedure TOF_BTAFFAIREFRSGEST.DecodeArgs(S: string);
var LesArgs,UnArg : string;
		UnCode,Unevaleur : string;
    IposEgal : integer;
begin
	lesArgs := S;
  Repeat
  	UnArg := READTOKENST(LesArgs);
    if UnArg <> '' then
    begin
      IposEgal := Pos('=',UnArg);
      if IposEgal > 0 then
      begin
        Uncode := copy(unArg,1,IPosEgal-1);
        UneValeur := copy(unArg,IPosEgal+1,255);
        if UnCode='AFFAIRE' then
        begin
        	fAffaire := UneValeur;
        end;
        if UnCode='ACTION' then
        begin
        	if UneValeur = 'MODIFICATION' then faction := taModif else faction := taConsult;
        end;
      end;
    end;
  Until UnArg='';
  //
end;

procedure TOF_BTAFFAIREFRSGEST.BDeleteClick(Sender: Tobject);
var cancel : boolean;
begin
  if PGIask ('Confirmez-vous la suppression ?',ecran.caption)=mryes then
  begin
  	if FTOBEnt <> nil then
    begin
    	ExecuteSql('DELETE FROM AFFAIREFRSGEST WHERE BAF_AFFAIRE="'+fAffaire+'" AND BAF_CODEPORT="'+fTOBENT.getValue('BAF_CODEPORT')+'"');
      fTOBEnt.free;
      GE.deleteRow(GE.row);
      if GE.rowCount < 2 then AfficheEnt;
      GERowEnter(self,GE.row,cancel,false);
    end;
  end;
end;

procedure TOF_BTAFFAIREFRSGEST.BDeleteDetClick(Sender: Tobject);
var cancel : boolean;
begin
  if PGIask ('Confirmez-vous la suppression de cet affectation ?',ecran.caption)=mryes then
  begin
  	if fTOBL  <> nil then
    begin
    	ExecuteSql('DELETE FROM AFFAIREFRSGEST WHERE BAF_AFFAIRE="'+fAffaire+'" '+
      					 'AND BAF_CODEPORT="'+fTOBENT.getValue('BAF_CODEPORT')+'"'+
                 'AND BAF_TIERS="'+fTOBL.getValue('BAF_TIERS')+'"');
      fTOBL.free;
      GSEnter(GS);
      GS.deleteRow(GS.row);
      if fTOBFrsDet.detail.count = 0 then
      begin
      	GS.rowCount := 2 ;
				AffichelaGrille (fTOBFrsDet,GS,fColNamesGS ,false);
        GS.row := 1;
      end else
      begin
      	GS.rowCount := fTOBFrsDet.detail.count+1;
      end;
      SetEventsGS (false);
      GSRowEnter(GS,GS.row,cancel,false);
    	stprev := GS.Cells [GS.col,GS.row];
      SetEventsGS (true);
    end;
  end;
end;

procedure TOF_BTAFFAIREFRSGEST.BInsertClick(Sender: Tobject);
var CODE: THCritMaskEdit;
	  Coord: TRect;
    TheCode : string;
begin
  Coord := GE.CellRect(GE.Col, GE.Row);
  CODE := THCritMaskEdit.Create(ECRAN);
  CODE.Parent := GE;
  CODE.Top := Coord.Top;
  CODE.Left := Coord.Left;
  CODE.Width := 3;
  CODE.Visible := False;
  CODE.Text := '';
  CODE.DataType := 'GCPORT';
  CODE.Plus := ' AND (GPO_TYPEFRAIS = "B00")';
  if LookupCombo(CODE) then
  begin
    if CODE.Text <> '' then TheCode := CODE.Text;
    CODE.free;
    if FindCodePort (TheCode) then
    begin
      PgiInfo ('IMPOSSIBLE : Frais de gestion déjà présent',ecran.name);
      exit;
    end;
    AjouteFraisgestion (TheCode);
  end;
end;

function TOF_BTAFFAIREFRSGEST.FindInterv (CodeInter : string) : boolean;
begin
	result := (fTOBFrsDet.findFirst(['BAF_TIERS'],[CodeInter],true)<>nil);
end;

function TOF_BTAFFAIREFRSGEST.AjouteFrsInterv (CodeInter : string) : TOB;
var TOBDF : TOB;
		TOBT : TOB;
    QQ : Tquery;
begin
	TOBT := TOB.create ('TIERS',nil,-1);
  QQ := OpenSql ('SELECT T_LIBELLE AS LIBTIERS FROM TIERS WHERE T_TIERS="'+CodeInter+'" AND T_NATUREAUXI="FOU"',True,1,'',true);
  TOBT.SelectDB('',QQ);
  ferme (QQ);
  TOBDF := TOB.Create ('AFFAIREFRSGEST',fTOBFrsDet,-1);
  TOBDF.putvalue('BAF_AFFAIRE',faffaire);
  TOBDF.putvalue('BAF_CODEPORT',fTOBEnt.getValue('BAF_CODEPORT'));
  TOBDF.putvalue('BAF_TIERS',CodeInter);
  TOBDF.putvalue('BAF_TYPEPORT',fTOBEnt.getValue('GPO_TYPEPORT'));
  TOBDF.putvalue('BAF_COEFF',fTOBEnt.getValue('GPO_COEFF'));
  TOBDF.putvalue('BAF_PVHT',fTOBEnt.getValue('GPO_PVHT'));
  TOBDF.putvalue('BAF_PVTTC',fTOBEnt.getValue('GPO_PVTTC'));
  TOBDF.AddChampSupValeur ('LIBTIERS',TOBT.getValue('LIBTIERS'));
  TOBDF.AddChampSupValeur ('SEL','');
  fTOBFrsDet.putValue('MODIFIED','X');
  result:= TOBDF;
  TOBT.free;
end;

procedure TOF_BTAFFAIREFRSGEST.BInsertDetClick(Sender: Tobject);
var retour : string;
		TOBL : TOB;
    Acol,Arow : integer;
    cancel : boolean;
begin
  retour := AGLLanceFiche('BTP','BTMULAFFAIREINTER','BAI_AFFAIRE='+faffaire,'','');
  if retour <> '' then
  begin
  	if not FindInterv (retour) then
    begin
			TOBL := AjouteFrsInterv (retour);
      GS.InsertRow(GS.RowCount);
      if fTOBFrsDet.detail.count = 0 then GS.rowCount := 2 else GS.rowCount := fTOBFrsDet.detail.count+1;
      AffichelaGrille (fTOBFrsDet,GS,fColNamesGS ,false);
      GSRowEnter(self,GS.rowCount-1,cancel,false);
      Acol := 1;
      Arow := GS.rowCount-1;
      GSCellEnter(self,Acol,Arow,cancel);
      GS.row := Arow;
      GS.col := Acol;
    end else
    begin
    	PgiInfo('IMPOSSIBLE : cet intervenant est déjà présent',ecran.Caption );
    end;
  end;
end;

procedure TOF_BTAFFAIREFRSGEST.AjouteFraisgestion(TheCode: string);
var OneTOB,TOBTT,TOBT,TOBDF : TOB;
		QQ : Tquery;
    cancel : boolean;
    indice : integer;
begin
	TOBTT := TOB.Create ('LES COTRAITANTS',nil,-1);
	GERowexit(self,GE.row,cancel,false);
	if fTOBFrsEnt.detail.count > 0 then
  begin
  	GE.rowCount := GE.RowCount +1;
    GE.row := GE.RowCount -1;
  end else
  begin
  	GE.row := 1;
  end;
  OneTOB := TOB.create ('ONE TOB',fTOBFrsEnt,-1);
  OneTOB.AddChampSupValeur('NEW','X');
  ONETOB.AddChampSupValeur('SEL','');
  ONETOB.AddChampSupValeur('BAF_AFFAIRE',faffaire);
  ONETOB.AddChampSupValeur('BAF_CODEPORT',theCode);
  ONETOB.AddChampSupValeur('GPO_TYPEPORT','');
  ONETOB.AddChampSupValeur('GPO_LIBELLE','');
  ONETOB.AddChampSupValeur('GPO_PVHT',0.0);
  ONETOB.AddChampSupValeur('GPO_PVTTC',0.0);
  ONETOB.AddChampSupValeur('GPO_COEFF',0.0);
  QQ := openSql ('SELECT GPO_TYPEPORT,GPO_LIBELLE,GPO_PVHT,GPO_PVTTC,GPO_COEFF '+
  							 'FROM PORT WHERE GPO_CODEPORT="'+TheCode+'"',true,1,'',true);
  if not QQ.eof then
  begin
    ONETOB.AddChampSupValeur('GPO_TYPEPORT',QQ.findField('GPO_TYPEPORT').AsString );
    ONETOB.AddChampSupValeur('GPO_LIBELLE',QQ.findField('GPO_LIBELLE').AsString);
    ONETOB.AddChampSupValeur('GPO_PVHT',QQ.findField('GPO_PVHT').AsFloat );
    ONETOB.AddChampSupValeur('GPO_PVTTC',QQ.findField('GPO_PVTTC').AsFloat);
    ONETOB.AddChampSupValeur('GPO_COEFF',QQ.findField('GPO_COEFF').AsFloat);
  end;
  ferme (QQ);
  QQ := OpenSql ('SELECT BAI_TIERSFOU,T_LIBELLE AS LIBTIERS FROM AFFAIREINTERV '+
  							 'LEFT JOIN TIERS ON BAI_TIERSFOU=T_TIERS AND T_NATUREAUXI="FOU" '+
  							 'WHERE BAI_AFFAIRE="'+fAffaire+'"',True,-1,'',true);
  if not QQ.eof then
  begin
    TOBTT.LoadDetailDB('LES COTRAITANTS','','',QQ,false);
  end;
  ferme (QQ);
  fTOBFrsDet.ClearDetail; fTOBFrsDet.putValue('MODIFIED','X');
  for Indice := 0 to TOBTT.detail.count -1 do
  begin
    TOBT := TOBTT.detail[indice];
    TOBDF := TOB.Create ('AFFAIREFRSGEST',fTOBFrsDet,-1);
    TOBDF.putvalue('BAF_AFFAIRE',faffaire);
    TOBDF.putvalue('BAF_CODEPORT',TheCode);
    TOBDF.putvalue('BAF_TIERS',TOBT.getValue('BAI_TIERSFOU'));
    TOBDF.putvalue('BAF_TYPEPORT',OneTOB.getValue('GPO_TYPEPORT'));
    TOBDF.putvalue('BAF_COEFF',OneTOB.getValue('GPO_COEFF'));
    TOBDF.putvalue('BAF_PVHT',OneTOB.getValue('GPO_PVHT'));
    TOBDF.putvalue('BAF_PVTTC',OneTOB.getValue('GPO_PVTTC'));
    TOBDF.AddChampSupValeur ('LIBTIERS',TOBT.getValue('LIBTIERS'));
    TOBDF.AddChampSupValeur ('SEL','');
  end;
  TOBTT.free;
  fTOBEnt := ONETOB;
  AfficheLaLigne (ONETOB,GE.row-1,GE,fColNamesGE,false);
  AffichelesLignes;
end;

function TOF_BTAFFAIREFRSGEST.FindCodePort(TheCode: string): boolean;
begin
  result :=  (fTOBFrsEnt.findFirst(['BAF_CODEPORT'],[TheCode],true)<>nil);
end;


procedure TOF_BTAFFAIREFRSGEST.ZoneSuivanteOuOk(Grille : THgrid;var ACol, ARow : integer;var  Cancel : boolean);
var Sens, ii, Lim: integer;
  OldEna, ChgLig, ChgSens: boolean;
begin
  OldEna := Grille.SynEnabled;
  // Modif BTP
  Lim := Grille.RowCount ;
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
      // ---
      if ((ACol >= Grille.ColCount-1) and (ARow >= Lim -1)) then
      begin
        Arow := 1;
        Acol := Grille.FixedCols;
      end;
      if ACol < Grille.ColCount-1 then
      begin
      	Inc(ACol)
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
        ACol := 1;
      end;
    end;
  end;
  Grille.SynEnabled := OldEna;
end;

function TOF_BTAFFAIREFRSGEST.ZoneAccessible( Grille : THgrid; var ACol, ARow : integer) : boolean;
begin
  result := true;
  if (Grille.ColWidths[acol] = 0) or (not Grille.ColEditables[acol] ) then BEGIN result := false; exit; END;
  if Arow < GS.FixedRows then exit;
  if Acol < GS.fixedCols then exit;
  if GS.RowHeights[ARow] <= 0 then begin result := false;Exit;  end;
end;

procedure TOF_BTAFFAIREFRSGEST.FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
var FocusGridGS,FocusGridGE: Boolean;
begin
  if fAction = taConsult then Exit;
  FocusGridGE := false;
  FocusGridGS := false;
  if (Screen.ActiveControl = GS) then
  begin
    FocusGridGS := True;
  end;
  if (Screen.ActiveControl = GE) then
  begin
    FocusGridGS := true;
  end;
  case Key of
    VK_RETURN: Key := VK_TAB;
    VK_INSERT:
      begin
      	Key := 0;
        if FocusGridGS then BInsertDetClick (self);
        if FocusGridGE then BInsertClick (self);
      end;
    VK_DELETE:
      begin
        if (Shift = [ssCtrl]) then
        begin
          Key := 0;
          if FocusgridGS then BDeleteDetClick(self);
          if FocusgridGE then BDeleteClick(self);
        end;
      end;
  end;
end;

Initialization
  registerclasses ( [ TOF_BTAFFAIREFRSGEST ] ) ;
end.
