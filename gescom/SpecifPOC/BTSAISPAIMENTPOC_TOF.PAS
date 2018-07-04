{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 07/11/2017
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BTSAISPAIMENTPOC ()
Mots clefs ... : TOF;BTSAISPAIMENTPOC
*****************************************************************}
Unit BTSAISPAIMENTPOC_TOF ;

Interface

Uses StdCtrls, 
     Controls, 
     Classes, 
{$IFNDEF EAGLCLIENT}
     db, 
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF} 
     mul,Fe_Main,
{$else}
     eMul, 
{$ENDIF}
     uTob,
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     Vierge,
     Graphics,
     SAISUTIL,
     Windows,
     Messages,
     AGLInit,
     HTB97,
     UTOF ;

Type
  TOF_BTSAISPAIMENTPOC = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  private
    TOBSAISIE : TOB;
    TOBParam : TOB;
    TOBPieceTrait : TOb;
    TOBSOUSTrait : TOB;
    GS : THGrid;
    stcell : string;
    fColNames,Falignement,Ftitre,fLargeur : string;
    ColSaisie : Integer;
    DEV : Rdevise;
    procedure BfermeClick (Sender : TObject);
    procedure DefiniGrid;
    procedure DefiniGrille(Grid: ThGrid; fColNames, Falignement, Ftitre,fLargeur: string);
    procedure ChargeLaGrille;
    procedure GSCellEnter(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
    procedure GSCellExit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
    procedure GSKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure ZoneSuivanteOuOk(Grille: THgrid; var ACol, ARow: integer; var Cancel: boolean);
    function ZoneAccessible(Grille: THgrid; var ACol,ARow: integer): boolean;
  end ;

function DemandeReglementSousTrait (TOBPiece,TOBTiers,TOBEches,TOBAcomptes,TObpieceTrait,TOBPieceRG,TOBSousTrait,TOBPorcs: TOB; MontantPaiement : double;  DEV : RDevise) : boolean;

Implementation
uses FactComm;

function DemandeReglementSousTrait (TOBPiece,TOBTiers,TOBEches,TOBAcomptes,TObpieceTrait,TOBPieceRG,TOBSousTrait,TOBPorcs: TOB; MontantPaiement : double;  DEV : RDevise) : boolean;
var TOBParam : TOB;
begin
  Result := True;
  TOBParam := TOB.Create('LES PARAMS',nil,-1);
  TRY
    TOBParam.AddChampSupValeur('OKOK','X');
    TOBParam.AddChampSupValeur('MODIF','-');
    TOBParam.AddChampSupValeur('NETAPAYER',MontantPaiement);
    TOBParam.AddChampSupValeur('NBDEC',DEV.decimale);
    TheTOB := TOBParam;
    TOBParam.Data := TObpieceTrait;
    TOBpieceTrait.data := TOBSousTrait;
    AGLLanceFiche('BTP','BTSAISPAIMENTPOC','','','ACTION=MODIFICATION');
    TheTOB := nil;
    TOBpieceTrait.data := nil;
    if TOBParam.getstring('MODIF')='X' then
    begin
      TOBEches.clearDetail;
      GereEcheancesGC(TOBPiece, TOBTiers, TOBEches, TOBAcomptes, TOBPieceRG,TOBPieceTrait,TOBPorcs, taCreat, DEV, false);
    end;
    if TOBParam.GetString('OKOK') <> 'X' then Result := false;
  FINALLY
    TOBParam.free;
  end;
end;

procedure TOF_BTSAISPAIMENTPOC.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_BTSAISPAIMENTPOC.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_BTSAISPAIMENTPOC.OnUpdate ;
var II : Integer;
    TOBS , TOBL, TOBI  : TOB;
begin
  Inherited ;
  if ARRONDI(THNumEdit(getCOntrol('NETAPAYER')).Value,DEV.Decimale) <> ARRONDI(THNumEdit(getCOntrol('ENPAIEMENT')).Value,DEV.decimale) then
  begin
    PGIError('Le montant en règlement ne corrrespond pas au montant à payer par le client');
    ecran.ModalResult := 0;
    exit;
  end;
  //
  for II := 0 to TOBSAISIE.Detail.count -1 do
  begin
    TOBS := TOBSAISIE.detail[II];
    TOBL := TOBPieceTrait.FindFirst(['BPE_FOURNISSEUR'],[TOBS.GetString('BPE_FOURNISSEUR')],true);
    if TOBL = nil then continue;
    if TOBS.GetDouble('BPE_MONTANTREGL')<>TOBL.GetDouble('BPE_MONTANTREGL') then
    begin
      TOBL.SetDouble('BPE_MONTANTREGL',TOBS.GetDouble('BPE_MONTANTREGL'));
      TOBL.SetBoolean ('BPE_REGLSAISIE',true);
      if TOBS.GetString('BPE_FOURNISSEUR') <> '' then
      begin
        TOBI := TOBSOUSTrait.FindFirst(['BPI_TIERSFOU'],[TOBS.GetString('BPE_FOURNISSEUR')],true);
        if TOBI <> nil then
        begin

          if TOBS.GetDouble('BPE_MONTANTREGL') <> 0 then
          begin
            TOBI.SetString('BPI_TYPEPAIE','001');
            TOBL.SetString('TYPEPAIE','001');
          end else
          begin
            TOBI.SetString('BPI_TYPEPAIE','002');
            TOBL.SetString('TYPEPAIE','002');
          end;
        end;
      end;
    end;
  end;
  TOBParam.SetString('MODIF','X');
end ;

procedure TOF_BTSAISPAIMENTPOC.OnLoad ;
begin
  Inherited ;
end ;



procedure TOF_BTSAISPAIMENTPOC.OnArgument (S : String ) ;
var cancel :boolean;
    Arow,Acol : Integer;
begin
  Inherited ;
  TOBSAISIE := TOB.Create('LES LIGNES',nil,-1);
  //
  TOBParam := LaTOB;
  DEV.decimale := TOBParam.Getinteger('NBDEC');
  TOBPieceTrait := TOB(TOBParam.Data);
  TOBSOUSTRAIT := TOB(TOBPieceTrait.Data);
  TOBPieceTrait.Data := nil;
  TOBParam.Data := nil;
  TOBSAISIE.Dupliquer(TOBPieceTrait,True,true);
  GS := THGrid(GetControl('GS'));
  THNumEdit(getCOntrol('NETAPAYER')).Value := Arrondi(TOBParam.GetDouble('NETAPAYER'),DEV.Decimale);
  THNumEdit(getCOntrol('ENPAIEMENT')).Value := Arrondi(TOBParam.GetDouble('NETAPAYER'),DEV.Decimale);
  TToolbarButton97 (GetControl('Bferme')).onclick := BfermeClick;
  DefiniGrid;
  ChargeLaGrille;
  Acol := ColSaisie;
  Arow := 1;
  SetFocusControl('GS'); 
  GS.Row := 1;
  GS.Col := ColSaisie;
  GS.OnCellEnter := GSCellEnter;
  GS.OnCellExit := GSCellExit;
  GS.OnKeyDown := GSKeyDown;
  GSCellEnter(self,Acol,Arow,cancel);
end ;

procedure TOF_BTSAISPAIMENTPOC.OnClose ;
begin
  TOBSAISIE.free;
  Inherited ;
end ;

procedure TOF_BTSAISPAIMENTPOC.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_BTSAISPAIMENTPOC.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_BTSAISPAIMENTPOC.DefiniGrid;
begin
  fColNames := 'SEL;BPE_FOURNISSEUR;LIBELLE;BPE_MONTANTREGL;';
  Falignement := 'D.0  ---;G.0  ---;G.0  ---;D/2O -X-;';
  Ftitre := 'Num;Sous-Traitant;Raison sociale;Règlement client;';
  fLargeur := '5;18;60;20;';
  DefiniGrille (GS,fColNames,Falignement,Ftitre,fLargeur);
end;

procedure TOF_BTSAISPAIMENTPOC.DefiniGrille (Grid: ThGrid; fColNames,Falignement,Ftitre,fLargeur : string);
var st,lestitres,lesalignements,FF,alignement,Nam,leslargeurs,lalargeur,letitre,lelement : string;
    Obli,OkLib,OkVisu,OkNulle,OkCumul,Sep,Okimg : boolean;
    dec,NbCols : integer;
    indice,ind : integer;
    FFQTE : string;
begin
  //
  if Grid.rowcount < 2 then Grid.rowCount := 2;
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
  Grid.ColCount := Nbcols;
  //
  st := fColNames;
  lesalignements := Falignement;
  lestitres := Ftitre;
  leslargeurs := fLargeur;

  for indice := 0 to Nbcols -1 do
  begin
    Nam := ReadTokenSt (St); // nom
    alignement := ReadTokenSt(lesalignements);
    lalargeur := readtokenst(leslargeurs);
    letitre := readtokenst(lestitres);
    OkVisu  := True;
    TransAlign(alignement,FF,Dec,Sep,Obli,OkLib,OkVisu,OkNulle,OkCumul) ;
    grid.cells[Indice,0] := leTitre;
    grid.ColNames [Indice] := Nam;
    if copy(Alignement,1,1)='G'      then grid.ColAligns[indice] := taLeftJustify
    else if copy(Alignement,1,1)='D' then grid.ColAligns[indice] := taRightJustify
    else if copy(Alignement,1,1)='C' then grid.ColAligns[indice] := taCenter;
    okImg := (copy(Alignement,8,1)='X');
    if (Nam = 'BPE_MONTANTREGL') then
    begin
      grid.ColEditables[indice] := True;
      ColSaisie := indice;
    end else
    begin
      grid.ColEditables[indice] := false;
      grid.ColLengths [indice] := 0;
    end;
    //
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
      if Sep then FFQte := FFQTE+' ##';
      if Dec > 0 then
      begin
        FFQTE := FFQTE+'0.';
        for ind := 1 to Dec do
        begin
          FFQTE := FFQTE + '0';
        end;
      end else FFQTE := FFQTE + '0';
      GS.ColFormats[indice] := FFQTE ;
    end;

  end ;
end;


procedure TOF_BTSAISPAIMENTPOC.ChargeLaGrille;
begin
	GS.RowCount := TOBSAISIE.Detail.count+1;
  if GS.rowCount < 2 then GS.rowCount := 2;
	TOBSAISIE.PutGridDetail(GS,false,false,fColNames);
  TFvierge(ecran).HMTrad.ResizeGridColumns(GS);
end;

procedure TOF_BTSAISPAIMENTPOC.GSCellEnter(Sender: TObject; var ACol,ARow: Integer; var Cancel: Boolean);
begin
  ZoneSuivanteOuOk  (GS,Acol,Arow,Cancel);
  if not cancel then
  begin
    stcell := GS.Cells[Acol,Arow];  
  end;
end;

procedure TOF_BTSAISPAIMENTPOC.GSCellExit(Sender: TObject; var ACol,ARow: Integer; var Cancel: Boolean);
var TOBL : TOB;
begin
  if GS.cells[Acol,Arow] = stcell then Exit;
  if Acol = ColSaisie then
  begin
    TOBL := TOBSAISIE.detail[Arow-1];

    THNumEdit(getCOntrol('ENPAIEMENT')).Value := arrondi(THNumEdit(getCOntrol('ENPAIEMENT')).Value - TOBL.GetDouble('BPE_MONTANTREGL'),DEV.decimale);

    TOBL.SetDouble('BPE_MONTANTREGL',VALEUR(GS.cells[Acol,Arow]));
    
    THNumEdit(getCOntrol('ENPAIEMENT')).Value := arrondi(THNumEdit(getCOntrol('ENPAIEMENT')).Value + TOBL.GetDouble('BPE_MONTANTREGL'),DEV.Decimale);

    if THNumEdit(getCOntrol('ENPAIEMENT')).Value <> THNumEdit(getCOntrol('NETAPAYER')).Value then
    begin
      THNumEdit(getCOntrol('ENPAIEMENT')).Font.Color := clred;
    end else
    begin
      THNumEdit(getCOntrol('ENPAIEMENT')).Font.Color := clblack;
    end;
    TOBL.PutLigneGrid(GS,ARow,False,False,fColNames);
    StCell := GS.cells[Acol,Arow];
  end;
end;



function TOF_BTSAISPAIMENTPOC.ZoneAccessible( Grille : THgrid; var ACol, ARow : integer) : boolean;
begin
  result := true;
  if (Grille.ColWidths[acol] = 0) or (not Grille.ColEditables[acol] ) then BEGIN result := false; exit; END;
  if Arow < GS.FixedRows then exit;
  if Acol < GS.fixedCols then exit;
end;

procedure TOF_BTSAISPAIMENTPOC.ZoneSuivanteOuOk(Grille : THgrid;var ACol, ARow : integer;var  Cancel : boolean);
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
        Sens := -1;
        continue;
      end;
      if ChgLig then
      begin
        ACol := Grille.FixedCols - 1;
        ChgLig := False;
      end;
      if ACol < Grille.ColCount - 1 then Inc(ACol) else
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


procedure TOF_BTSAISPAIMENTPOC.GSKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
var OkG, Vide: Boolean;
  TOBL : TOB;
  Diff : double;
begin
  OkG := (Screen.ActiveControl = GS);
  Vide := (Shift = []);
  case Key of
    VK_RETURN: if ((OkG) and (Vide)) then
      begin
        Key := 0;
        SendMessage(GS.Handle, WM_KeyDown, VK_TAB, 0);
      end;
    VK_F11: if ((OkG) and (Vide)) then
      begin
        Key := 0;
        TOBL := TOBSAISIE.detail[GS.row-1];
        Diff := TOBL.GetDouble('BPE_MONTANTREGL') + THNumEdit(getCOntrol('NETAPAYER')).Value - THNumEdit(getCOntrol('ENPAIEMENT')).Value;
        GS.cells[GS.Col,GS.row] := STRFPOINT(diff);
        SendMessage(GS.Handle, WM_KeyDown, VK_TAB, 0);
      end;
  end;
end;

procedure TOF_BTSAISPAIMENTPOC.BfermeClick(Sender: TObject);
begin
  TOBParam.SetString('OKOK','-');
  Ecran.Close;
end;

Initialization
  registerclasses ( [ TOF_BTSAISPAIMENTPOC ] ) ;
end.
