{***********UNITE*************************************************
Auteur  ...... : LS
Créé le ...... : 08/12/2011
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BTPIECERGCOT ()
Mots clefs ... : TOF;BTPIECERGCOT
*****************************************************************}
Unit BTPIECERGCOT_TOF ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
     types,
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul,
     fe_main,
{$else}
     eMul,
     MainEagl,
{$ENDIF}
		 AglInit,
		 vierge,
     uTob,
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     Graphics,
     Grids,
     HEnt1,
     HMsgBox,
     HTB97,
     menus,
     Saisutil,
     UTOF ;

Type
  TOF_BTPIECERGCOT = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  private
    fColNames : string;
    Falignement : string;
    fTitre : string;
		fLargeur : string;
    fNC : string;
    fDEV : Rdevise;
    // ---
    TOBPieceTrait : TOB;
  	TOBPiece : TOB;
    TOBPieceRG : TOB;
    TOBBases : TOB;
    TOBBasesRG : TOB;
    TOBPorcs : TOB;
    // ---
    GS : THGrid;
    TYPERG : THValComboBox;
    TAUXRG : THNumEdit;
    MnCaution,MnApplication : TMenuItem;
    // -- Champs
    GSBASEHT,GSBASETTC,GSVALEURHT,GSVALEURTTC : integer;
    LARGBASEHT, LARGBASETTC, LARGVALEURHT, LARGVALEURTTC : integer;
    //
    // evenement sur grid et form
    procedure GSRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
		procedure GSPostDrawCell(ACol, ARow: Longint; Canvas: TCanvas; AState: TGridDrawState);
    procedure GSEnter(Sender: TObject);
    //
		procedure AffichelaGrille;
    procedure AfficheLigne(TOBL: TOB; Ligne: integer);
    procedure DefinieGrille;
    procedure DefinieTOBSaisie;
    procedure NettoieTOBSaisie;
		//
    procedure GetControls;
    procedure SetValues;
		procedure SetEvents;
    procedure SetEventsGrid (State : boolean);
    procedure PositionneMenus (ligne : integer);

    //
    function GetTypeInterv(Intervenant : string): string;
    function ModeCaution(TOBT: TOB): string;
		procedure TypeRGChange (Sender : TObject);
  	procedure TauxRgChange (Sender : Tobject);
  	procedure ChangeCaution (Sender :Tobject) ;
  	procedure MnApplicationClick (Sender : Tobject);
    procedure SetTOBValues(TypeRG: string; TauxRG: double);

    //
  end ;

Implementation

uses UCotraitance,FactRG;

procedure TOF_BTPIECERGCOT.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_BTPIECERGCOT.OnDelete ;
begin
  Inherited ;
  TOBPieceRG.ClearDetail;
  TOBBasesRG.clearDetail;
  TheTOB := TOBPieceRG;
  ecran.Close;
end ;

procedure TOF_BTPIECERGCOT.OnUpdate ;
begin
  Inherited ;
  TOBPieceRG.data := TOBBasesRG;
  TheTOB := TOBPieceRG;
end ;

procedure TOF_BTPIECERGCOT.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_BTPIECERGCOT.OnArgument (S : String ) ;
begin
  Inherited ;
  TOBPieceTrait := laTOB;
  TOBPieceRG := TOB(TOBPieceTrait.data);
  TOBPiece := TOB(TOBpieceRG.data);
  TOBBasesRG := TOB(TOBPiece.data);
  TOBBases := TOB(TOBBasesRG.data);
  TOBPorcs := TOB(TOBBases.data);
  fDEV.Code  := TOBPiece.GetString('GP_DEVISE');
  GetInfosDevise(fDEV);

  //
  fColNames := 'SEL;TYPEINTERV;LIBELLEFOU;PRG_APPLICABLE;BASEHT;BASETTC;PRG_MTHTRGDEV;PRG_MTTTCRGDEV;CAUTION';
  Falignement := 'C.0  ---;C.0O --X;G.0  ---;C.0  X--;D/2O -X-;D/2O -X-;D/2O -X-;D/2O -X-;C.0  ---;';
  Ftitre := ' ;Type;Raison sociale;Application;Marché;Marché;Montant RG;Montant RG;Caution';
  fLargeur := '5;5;70;18;20;20;20;20;18;';
  fNC := '1;1;1;0;1;1;1;1;1;';
  //
  GetControls;
  //
  DefinieTOBSaisie;
  DefinieGrille;
  SetValues;
  GS.row := 1;
  AffichelaGrille;
  SetEvents;
  TypeRGChange(self);
end ;

procedure TOF_BTPIECERGCOT.OnClose ;
begin
  NettoieTOBSaisie;
  Inherited ;
end ;

procedure TOF_BTPIECERGCOT.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_BTPIECERGCOT.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_BTPIECERGCOT.GetControls;
begin
	GS := THGrid(GetCOntrol('GS'));
  TYPERG := THValComboBox (GetCOntrol('TYPERG'));
  TAUXRG := THNumEdit(GeTControl('TAUXRG'));
  MnCaution := TMenuItem(GetControl('MnCaution'));
  MnApplication := TMenuItem(getControl('MnApplication'));
end;

procedure TOF_BTPIECERGCOT.DefinieGrille;
var st,lestitres,lesalignements,FF,alignement,Nam,leslargeurs,lalargeur,letitre,lelement : string;
    Obli,OkLib,OkVisu,OkNulle,OkCumul,Sep,Okimg : boolean;
    dec,NbCols : integer;
    indice,ind : integer;
    FFQTE,FFDEC : string;
begin
  //
  st := fColNames;
  NbCols := 0;
  repeat
    lelement := READTOKENST (st);
    if lelement <> '' then
    begin
      inc(NbCols);
    end;
  until lelement = '';
  //
  GS.ColCount := Nbcols;
  //
  st := fColNames;
  lesalignements := Falignement ;
  lestitres := fTitre ;
  leslargeurs := fLargeur;

  for indice := 0 to Nbcols -1 do
  begin
    Nam := ReadTokenSt (St); // nom
    if Nam = 'BASEHT' then GSBASEHT := Indice
    else if Nam = 'BASETTC' then GSBASETTC := Indice
    else if Nam = 'PRG_MTHTRGDEV' then GSVALEURHT := Indice
    else if Nam = 'PRG_MTTTCRGDEV' then GSVALEURTTC := Indice;
    alignement := ReadTokenSt(lesalignements);
    lalargeur := readtokenst(leslargeurs);
    letitre := readtokenst(lestitres);
    TransAlign(alignement,FF,Dec,Sep,Obli,OkLib,OkVisu,OkNulle,OkCumul) ;
    GS.cells[Indice,0] := leTitre;
    GS.ColNames [Indice] := Nam;
    if copy(Alignement,1,1)='G' then GS.ColAligns[indice] := taLeftJustify
    else if copy(Alignement,1,1)='D' then GS.ColAligns[indice] := taRightJustify
    else if copy(Alignement,1,1)='C' then GS.ColAligns[indice] := taCenter;
    okImg := (copy(Alignement,8,1)='X');
    if OkVisu then
    begin
  		GS.ColWidths[indice] := strtoint(lalargeur)*GS.Canvas.TextWidth('W');
    end else
    begin
    	GS.ColWidths[indice] := -1;
    end;

    if Nam = 'BASEHT' then LARGBASEHT := GS.ColWidths[indice]
    else if Nam = 'BASETTC' then LARGBASETTC := GS.ColWidths[indice]
    else if Nam = 'PRG_MTHTRGDEV' then LARGVALEURHT := GS.ColWidths[indice]
    else if Nam = 'PRG_MTTTCRGDEV' then LARGVALEURTTC := GS.ColWidths[indice];

    if (ChampToType(nam) = 'BOOLEAN') or (Nam = 'CAUTION')  then
    begin
      GS.ColTypes [indice] := 'B' ;
      GS.colaligns[indice]:= tacenter;
      GS.colformats[indice]:= inttostr(Integer(csCoche));
    end else if (OkLib) or (okImg) then
    begin
    	GS.ColFormats[indice] := 'CB=' + Get_Join(Nam);
      if OkImg then
      begin
      	GS.ColDrawingModes[Indice]:= 'IMAGE';
      end;
    end else if (Dec<>0) or (Sep) then
    begin
      FFQTE := '# ##';
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


procedure TOF_BTPIECERGCOT.DefinieTOBSaisie;
var Indice : integer;
		SumHt,SumTTC,Ecart : double;
    TOBPRG : TOB;
    XD,XT : double;
    Intervenant : string;
begin
	For Indice := 0 to TOBPieceRG.detail.count - 1 do
  begin
		TOBPieceRG.detail[Indice].AddChampSupValeur('SEL','');
		TOBPieceRG.detail[Indice].AddChampSupValeur('TYPEINTERV',GetTypeInterv(TOBPieceRG.detail[Indice].GetString('PRG_FOURN')));
		TOBPieceRG.detail[Indice].AddChampSupValeur('LIBELLEFOU',GetLibIntervenant(TOBPieceRG.detail[Indice].GetString('PRG_FOURN')));
		TOBPieceRG.detail[Indice].AddChampSupValeur('BASEHT',0);
		TOBPieceRG.detail[Indice].AddChampSupValeur('BASETTC',0);
		TOBPieceRG.detail[Indice].AddChampSupValeur('CAUTION',ModeCaution(TOBPieceRG.detail[Indice]));
  end;
  SumHt := 0; SumTTC := 0;
  for Indice := 0 to TOBPieceTrait.detail.count -1 do
  begin
		Intervenant := TOBpieceTrait.detail[Indice].GetString('BPE_FOURNISSEUR');
//    if TOBpieceTrait.detail[Indice].GetString('TYPEPAIE')='002' then Intervenant := '';
    TOBPRG := TOBPieceRG.findFirst(['PRG_FOURN'],[Intervenant],true);
    if TOBPRG <> nil then
    begin
      XD := TOBpieceTrait.detail[Indice].GetDouble ('BPE_TOTALHTDEV');
      XT := TOBpieceTrait.detail[Indice].GetDouble ('BPE_TOTALTTCDEV');
      TOBPRG.SetDouble ('BASEHT',TOBPRG.GetDouble('BASEHT')+XD);
      TOBPRG.SetDouble ('BASETTC',TOBPRG.GetDouble('BASETTC')+XT);
      SumHt := SumHt + XD;
      SumTTC := SumTTC + XT;
    end;
  end;
  (*
  Ecart := TOBPiece.GetDouble ('GP_TOTALHTDEV') - SumHt;
  if Ecart <>0 then
  begin
		TOBPRG := TOBPieceRG.findfirst(['PRG_FOURN'],[''],true);
    if TOBPRG <> nil then
    begin
      TOBPRG.setDouble ('BASEHT',TOBPRG.GetDouble ('BASEHT')+Ecart);
    end;
  end;
  Ecart := TOBPiece.GetDouble ('GP_TOTALTTCDEV') - SumTTC;
  if Ecart <>0 then
  begin
		TOBPRG := TOBPieceRG.findfirst(['PRG_FOURN'],[''],true);
    if TOBPRG <> nil then
    begin
      TOBPRG.setDouble ('BASETTC',TOBPRG.GetDouble ('BASETTC')+Ecart);
    end;
  end;
  *)
end;

function TOF_BTPIECERGCOT.GetTypeInterv ( Intervenant : string) : string;
var TOBT : TOB;
begin
	result := '';
  TOBT := TOBPieceTrait.FindFirst(['BPE_FOURNISSEUR'],[Intervenant],true);
  if TOBT = nil then exit;
  result := TOBT.GetString('BPE_TYPEINTERV');
end;

procedure TOF_BTPIECERGCOT.SetValues;
begin
  TYPERG.Value := TOBPieceRG.detail[0].getString('PRG_TYPERG');
  TAUXRG.Value := TOBPieceRG.detail[0].GetDouble ('PRG_TAUXRG');
  if TYPERG.Value = '' then
  begin
    TypeRG.Value := 'HT';
    TAUXRG.Value := 5;
    SetTOBValues(TypeRG.Value,TauxRG.Value);
    TypeRGChange(self);
  end;
  RecalculeRG (TOBPorcs ,TOBPiece ,TOBPieceRG ,TOBBases,TOBBasesRG ,TOBPieceTrait,fdeV);
end;

function TOF_BTPIECERGCOT.ModeCaution(TOBT : TOB) : string;
begin
	result := '-';
  if TOBT.GetString('PRG_NUMCAUTION')='' then exit; // pas de numéro de caution définie
  if TOBT.GetDouble('PRG_CAUTIONMTDEV') < TOBT.GetDouble('PRG_MTTTCRGDEV') then
  begin
    exit;
  end;
  result := 'X';
end;

procedure TOF_BTPIECERGCOT.AfficheLagrille;
var Indice : integer;
begin
  SetEventsGrid (False);
  GS.VidePile (false);
  GS.RowCount := TOBPieceRG.detail.count +1;
  for Indice := 0 to TOBPieceRG.detail.count -1 do
  begin
    AfficheLigne(TOBPieceRG.detail[Indice],Indice+1);
  end;
  TFVierge(self).HMTrad.ResizeGridColumns(GS); 
  // entree dans la grille
  GS.SetFocus;
  GSEnter(self);
  // et hop on active les evenements sur la grille
  SetEventsGrid (True);
end;

procedure TOF_BTPIECERGCOT.AfficheLigne(TOBL : TOB;Ligne : integer);
begin
  TOBL.PutLigneGrid (GS,Ligne,false,false,fColNames);
end;

procedure TOF_BTPIECERGCOT.GSPostDrawCell(ACol, ARow: Integer;Canvas: TCanvas; AState: TGridDrawState);
var ARect: TRect;
		Triangle: array[0..2] of TPoint;
begin
  if Arow = 0 then exit;
  if (ACol = 0) then
  begin
    Arect := GS.CellRect(Acol, Arow);
    Canvas.Brush.Color := GS.FixedColor;
    Canvas.FillRect(ARect);
    if (ARow = GS.row) then
    begin
      Canvas.Brush.Color := clBlack;
      Canvas.Pen.Color := clBlack;
      Triangle[1].X := ARect.Right - 2;
      Triangle[1].Y := ((ARect.Top + ARect.Bottom) div 2);
      Triangle[0].X := Triangle[1].X - 5;
      Triangle[0].Y := Triangle[1].Y - 5;
      Triangle[2].X := Triangle[1].X - 5;
      Triangle[2].Y := Triangle[1].Y + 5;
      if false then Canvas.PolyLine(Triangle) else Canvas.Polygon(Triangle);
    end;
  end;
end;

procedure TOF_BTPIECERGCOT.GSRowEnter(Sender: TObject; Ou: Integer;var Cancel: Boolean; Chg: Boolean);
begin
	PositionneMenus (ou);
  GS.Invalidate;
end;

procedure TOF_BTPIECERGCOT.SetEvents;
begin
  TYPERG.OnChange := TypeRGChange;
  TAUXRG.OnExit := TauxRgChange;
  MnCaution.OnClick := ChangeCaution;
  MnApplication.OnClick := MnApplicationClick;
end;

procedure TOF_BTPIECERGCOT.SetEventsGrid(State: boolean);
begin
  if State then
  begin
    GS.OnEnter  := GSEnter;
    GS.OnRowEnter := GSRowEnter;
    GS.PostDrawCell := GSPostDrawCell;
  end
  else
  begin
    GS.OnEnter  := nil;
    GS.OnRowEnter := nil;
    GS.PostDrawCell := nil;
  end;
end;

procedure TOF_BTPIECERGCOT.GSEnter(Sender: TObject);
var cancel : boolean;
		Arow : integer;
begin
  cancel := false;
  Arow := 1;
  cancel := false;
  GSRowEnter (GS,Arow,cancel,false);
  GS.row := Arow;
end;

procedure TOF_BTPIECERGCOT.MnApplicationClick(Sender: Tobject);
var TOBT : TOB;
begin
	TOBT := TOBPieceRG.Detail[GS.row-1];
  if TOBT.GetBoolean ('PRG_APPLICABLE') then TOBT.SetBoolean('PRG_APPLICABLE',false)
  																			else TOBT.SetBoolean('PRG_APPLICABLE',True);
	AfficheLigne(TOBT,GS.row);  
end;

procedure TOF_BTPIECERGCOT.ChangeCaution(Sender: Tobject);
begin
  TOBPieceRG.detail[GS.row-1].AddChampSupValeur('DEV',fDEV.Code);
	TheTOB := TOBPieceRG.detail[GS.row-1];
  AGLLanceFiche('BTP','BTPIECERGCAU','','',ActionToString(TFVierge(ecran).TypeAction) );
  TOBPieceRG.detail[GS.row-1].DelChampSup('DEV',false);
  TheTOB := nil;
  TOBPieceRG.detail[GS.row-1].AddChampSupValeur('CAUTION',ModeCaution(TOBPieceRG.detail[GS.row-1]));
  AfficheLigne (TOBPieceRG.detail[GS.row-1],GS.row);
end;

procedure TOF_BTPIECERGCOT.TauxRgChange(Sender: Tobject);
begin
  SetTOBValues(TypeRG.Value,TauxRG.Value);
  RecalculeRG (TOBPorcs ,TOBPiece ,TOBPieceRG ,TOBBases,TOBBasesRG ,TOBPieceTrait,fdeV);
	AffichelaGrille;
end;

procedure TOF_BTPIECERGCOT.TypeRGChange(Sender: TObject);
var Indice : integer;
begin
	if TYPERG.Value = 'HT' then
  begin
    GS.ColWidths [GSBASETTC] := -1;
    GS.ColWidths [GSVALEURTTC] := -1;
    GS.ColWidths [GSBASEHT] := LARGBASEHT;
    GS.ColWidths [GSVALEURHT] := LARGVALEURHT;
  end else
  begin
    GS.ColWidths [GSBASETTC] := LARGBASETTC;
    GS.ColWidths [GSVALEURTTC] := LARGVALEURTTC;
    GS.ColWidths [GSBASEHT] := -1;
    GS.ColWidths [GSVALEURHT] := -1;
  end;
  For Indice := 0 to TOBPieceRG.detail.count -1 do
  begin
    TOBPIeceRG.detail[Indice].setString('PRG_TYPERG',TYPERG.Value);
  end;
  AffichelaGrille;
end;

procedure TOF_BTPIECERGCOT.PositionneMenus(ligne: integer);
var TOBT : TOB;
begin
  MnCaution.Enabled := false;
  MnApplication.visible := false;
	TOBT := TOBPieceRG.Detail[ligne-1];
  if TOBT.GetBoolean ('PRG_APPLICABLE') then
  begin
		MnApplication.Caption := 'Ne pas appliquer la R.G';
  end else
  begin
		MnApplication.Caption := 'Appliquer la R.G';
  end;
	if (TYPERG.Value = 'TTC') AND (TOBT.GetBoolean ('PRG_APPLICABLE')) then
  begin
    MnCaution.Enabled := true;
  end;
end;

procedure TOF_BTPIECERGCOT.NettoieTOBSaisie;
var Indice : integer;
begin
	For Indice := 0 to TOBPieceRG.detail.count - 1 do
  begin
		TOBPieceRG.detail[Indice].DelChampSup('SEL',false);
		TOBPieceRG.detail[Indice].DelChampSup('TYPEINTERV',false);
		TOBPieceRG.detail[Indice].DelChampSup('LIBELLEFOU',false);
		TOBPieceRG.detail[Indice].DelChampSup('BASEHT',false);
		TOBPieceRG.detail[Indice].DelChampSup('BASETTC',false);
		TOBPieceRG.detail[Indice].DelChampSup('CAUTION',false);
  end;
end;

procedure TOF_BTPIECERGCOT.SetTOBValues(TypeRG: string; TauxRG: double);
var Indice : integer;
begin
	For Indice := 0 to TOBPieceRG.Detail.count -1 do
  begin
    TOBPieceRG.detail[Indice].SetString('PRG_TYPERG',TypeRG);
    TOBPieceRG.detail[Indice].SetDouble('PRG_TAUXRG',TauxRG);
  end;
end;

Initialization
  registerclasses ( [ TOF_BTPIECERGCOT ] ) ;
end.
