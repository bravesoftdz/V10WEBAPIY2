{***********UNITE*************************************************
Auteur  ...... : Vautrain Franck
Créé le ...... : 06/01/2004
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : BTSAISIEVAR ()
Mots clefs ... : TOF;BTSAISIEVAR
*****************************************************************}
Unit UTofBtSaisieVar ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
     Vierge,
     HTB97,
     AglInit,
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul,
     fe_main,
{$else}
     eMul,
     maineagl,
{$ENDIF}
     uTob,
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     Grids,
     Graphics,
     Types,
     UtilsGrille,
     StrUtils,
     HSysMenu,
     UTOF ;

Type

  TOF_BTSAISIEVAR = Class (TOF)
  private
    //
    BValider      : TToolbarButton97;
    BValExcel     : TToolbarButton97;
    ToutVoir      : TCheckBox;
    //
    Sens          : Integer;
    //
    TOBVariables  : TOB;
    TobPartielle  : TOB;
    // Objets
    GrilleVar     : THGrid;
    GGrilleVar    : TGestionGS;
    //
    TOBResult     : TOB;
    //
    Procedure DestroyTOB;
    procedure GetObjects;
    procedure InitGrillevar;
    procedure ChargementEcran;
    procedure ChargementTobVariablePartielle;
    procedure ControleChamp(Champ, Valeur: String);
    //
    procedure GSCellEnter(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
    procedure GSCellExit (Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
    procedure GSGetCellCanvas(ACol, ARow: Integer; Canvas: TCanvas; AState: TGridDrawState);
    procedure GSPostDrawCell (ACol, ARow: Integer; Canvas: TCanvas;  AState: TGridDrawState);
    procedure GSRowEnter(Sender: TObject; ARow: Integer; var Cancel: Boolean; Chg : Boolean);
    procedure GSRowExit(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
    //
    procedure MAJTobVariables;
    procedure SetGrilleEvents(Etat: Boolean);
    procedure SetScreenEvents;
    procedure ToutVoirClick(Sender: Tobject);
    procedure ValiderClick(Sender: TObject);
    procedure ValiderXLClick(Sender: TObject);
    function  ZoneAccessible(ACol, ARow: Integer): boolean;
    procedure ZoneSuivanteOuOk(var ACol, ARow: Integer; var Cancel: boolean);


  Public
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ; // c'est la où tout commence
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  end;

Implementation

const SG_CODE     = 1;
      SG_LIBELLE  = 2;
      SG_VALEUR   = 3;

{Evenement sur la TOF }
procedure TOF_BTSAISIEVAR.OnArgument(S: String);
var x       : Integer;
    Critere : string;
    Champ   : string;
    Valeur  : string;
    ACol    : Integer;
    Arow    : Integer;
    Cancel  : Boolean;
begin
  Inherited ;
  //
  Cancel := false;
  //
  //Chargement des zones ecran dans des zones programme
  GetObjects;
  //
  Critere := uppercase(Trim(ReadTokenSt(S)));
  while Critere <> '' do
  begin
     x := pos('=', Critere);
     if x <> 0 then
        begin
        Champ  := copy(Critere, 1, x - 1);
        Valeur := copy(Critere, x + 1, length(Critere));
        end
     else
        Champ := Critere;
     ControleChamp(Champ, Valeur);
     Critere:= uppercase(Trim(ReadTokenSt(S)));
  end;

  TOBVariables := LaTob;
  //
  TobPartielle := TOB.Create('LES VARIABLES',nil,-1);
  TobPartielle.AddChampSupValeur ('RESULT',TOBVariables.GetValue('RESULT'));
  TobPartielle.AddChampSupValeur ('VALIDE',TOBVariables.GetValue('VALIDE'));
  //
  ChargementTobVariablePartielle;
  //
  //Gestion des évènement des zones écran
  SetScreenEvents;

  SetGrilleEvents(True);

  //Initialisation des grilles
  InitGrilleVar;

  //chargement des zones écran avec les zones tob
  ChargementEcran;

  if TOBVariables.GetValue('VALIDE') = '-' then
  Begin
    BValExcel.visible := False;
    ToutVoir.visible  := false;
  end
  Else
  Begin
    BValExcel.visible := True;
    ToutVoir.visible  := True;
  end;

  GSCellEnter(Self, ACol, Arow, Cancel);

end;

procedure TOF_BTSAISIEVAR.ChargementTobVariablePartielle;
Var TOBLPartielle : TOB;
    TOBL          : TOB;
    Ind           : Integer;
begin

  If TobPartielle.Detail.Count <> 0 then TobPartielle.ClearDetail;

  For ind := 0 to TobVariables.detail.count - 1 do
  begin
    TOBL := TOBVariables.Detail[Ind];
    if ToutVoir.Checked then
    Begin
      TobLPartielle := TOB.Create('BVARDOC', TobPartielle, -1);
      TobLPartielle.AddChampSupValeur('AFFICHABLE', TOBL.GetValue('AFFICHABLE'));
      TobLPartielle.Dupliquer(TOBL, true, true);
    end
    else
    begin
     if TOBL.GetValue('AFFICHABLE')= 'X' then
      begin
        TobLPartielle := TOB.Create('BVARDOC', TobPartielle, -1);
        TobLPartielle.AddChampSupValeur('AFFICHABLE', TOBL.GetValue('AFFICHABLE'));
        TobLPartielle.Dupliquer(TOBL, true, true);
      end;
    end;
  end;

  if GGrilleVar <> nil then GGrilleVar.TOBG := TobPartielle;

end;

procedure TOF_BTSAISIEVAR.OnCancel;
begin
  inherited;

end;

procedure TOF_BTSAISIEVAR.OnClose;
begin
  inherited;

end;

procedure TOF_BTSAISIEVAR.OnDelete;
begin
  inherited;

end;

procedure TOF_BTSAISIEVAR.OnDisplay;
begin
  inherited;

end;

procedure TOF_BTSAISIEVAR.OnLoad;
begin
  inherited;

end;

procedure TOF_BTSAISIEVAR.OnNew;
begin
  inherited;

end;

procedure TOF_BTSAISIEVAR.OnUpdate;
begin
  inherited;

  TheTOB := TOBVariables;

  DestroyTOB;
  
end;

procedure TOF_BTSAISIEVAR.GetObjects ;
begin
  //
  GrilleVar := THGrid(GetControl('GRID'));
  ToutVoir  := TCheckBox (GetControl('TOUTVOIR'));
  //
  BValider  := TToolbarButton97(GetControl('bvalider'));
  BValExcel := TToolbarButton97(GetControl('BAPPELEXCELL'));

end;

Procedure TOF_BTSAISIEVAR.DestroyTOB;
begin

  //Destruction des objets... c'était plus simple de les mettre ici !!!
  FreeAndNil(GGrilleVar);

  FreeAndNil(TOBPartielle);

end;

Procedure TOF_BTSAISIEVAR.SetScreenEvents;
begin

  ToutVoir.OnClick        := ToutVoirClick;
  BValider.OnClick        := ValiderClick;
  BValExcel.OnClick       := ValiderXLClick;

end;

Procedure TOF_BTSAISIEVAR.SetGrilleEvents(Etat : Boolean);
begin

  if Etat then
  begin
    GrilleVar.OnCellEnter     := GSCellEnter;
    GrilleVar.OnCellExit      := GSCellExit;
    GrilleVar.OnRowEnter      := GSRowEnter;
    GrilleVar.OnRowExit       := GSRowExit;
    GrilleVar.GetCellCanvas   := GSGetCellCanvas;
    GrilleVar.PostDrawCell    := GsPostDrawCell;
  end
  else
  begin
    GrilleVar.OnCellEnter     := Nil;
    GrilleVar.OnCellExit      := Nil;
    GrilleVar.OnRowEnter      := Nil;
    GrilleVar.OnRowExit       := Nil;
    GrilleVar.GetCellCanvas   := Nil;
    GrilleVar.PostDrawCell    := Nil;
  end;

end;
//
//--- Actions/Evènements sur la grille
//
procedure TOF_BTSAISIEVAR.GSCellEnter (Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
begin

  ZoneSuivanteOuOk(Acol, ARow, Cancel);

end;

procedure TOF_BTSAISIEVAR.GSRowEnter (Sender: TObject; ARow: Integer; var Cancel: Boolean; Chg : Boolean);
begin

end;

procedure TOF_BTSAISIEVAR.GSRowExit(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
begin

  GGrilleVar.GS.SynEnabled := False;

  if GGrilleVar.TOBG.detail[ou-1] <> nil then
  begin
    GGrilleVar.TOBG.detail[Ou-1].PutValue('BVD_VALEUR', Trim(GrilleVar.cells[SG_VALEUR,Ou]));
    TobVariables.Detail[Ou-1].PutValue('BVD_VALEUR',    Trim(GrilleVar.cells[SG_VALEUR,Ou]));
  end;

  GGrilleVar.GS.SynEnabled := true;

end;

procedure TOF_BTSAISIEVAR.GSCellExit (Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
begin

  if GGrilleVar.TOBG.detail[Arow-1] <> nil then
  begin
    GGrilleVar.TOBG.detail[Arow-1].PutValue('BVD_VALEUR', Trim(GrilleVar.cells[SG_VALEUR,Arow]));
    TobVariables.Detail[Arow-1].PutValue('BVD_VALEUR',    Trim(GrilleVar.cells[SG_VALEUR,Arow]));
  end;

end;

procedure TOF_BTSAISIEVAR.GSGetCellCanvas(ACol, ARow: Integer; Canvas: TCanvas; AState: TGridDrawState);

begin

  if GrilleVar.RowHeights[ARow] <= 0 then Exit;

  if (ACol < GrilleVar.FixedCols) or (Arow < GrilleVar.Fixedrows) then Exit;
  //
  // gestion des titres
  if (Acol = SG_LIBELLE) then
  begin
    if GrilleVar.Cells[SG_CODE,Arow]= '' then
    begin
      Canvas.Font.Style := Canvas.Font.Style + [fsBold];
      GrilleVar.Canvas.Brush.Color := $f8dfdf;
    end;
  end;

end;

procedure TOF_BTSAISIEVAR.GSPostDrawCell(ACol, ARow: Integer; Canvas: TCanvas; AState: TGridDrawState);
var Arect     : Trect;
    Decalage  : integer;
    TheText   : string;
begin

  if GrilleVar.RowHeights[ARow] <= 0 then Exit;

  if (ACol < GrilleVar.FixedCols) or (Arow < GrilleVar.Fixedrows) then Exit;

  ARect := GrilleVar.CellRect(ACol, ARow);

  if (ACol < GrilleVar.FixedCols) or (ARow < GrilleVar.FixedRows) then exit;

  GrilleVar.Canvas.Pen.Style   := psSolid;
  GrilleVar.Canvas.Brush.Style := BsSolid;

  Decalage := canvas.TextWidth('w') * 2;

  if (Acol = SG_LIBELLE) and (GrilleVar.Cells[SG_CODE, Arow]= '')then
  begin
    TheText := GrilleVar.Cells[SG_LIBELLE, Arow];
    GrilleVar.ColAligns[SG_LIBELLE] := taLeftJustify;
    Canvas.FillRect(ARect);
    GrilleVar.Canvas.Brush.Style := bsSolid;
    GrilleVar.Canvas.TextOut (Arect.left+Decalage+1,Arect.Top +2 ,Thetext);
  end;

end;

procedure TOF_BTSAISIEVAR.ZoneSuivanteOuOk(var ACol, ARow: Longint; var Cancel: boolean);
var Lim       : integer;
    OldEna    : Boolean;
    RowFirst  : Integer;
begin

	RowFirst := 1;

  OldEna  := GGrilleVar.GS.SynEnabled;
  GGrilleVar.GS.SynEnabled := False;

	Lim     := GGrilleVar.GS.RowCount - 1;

  ACol := GGrilleVar.GS.Col;
  ARow := GGrilleVar.GS.Row;

  while not ZoneAccessible(ACol, ARow) do
  begin
    Cancel := True;
    if Sens = -1 then
    begin
      if (ARow <= 1) then
      begin
        ARow    := 1;
        Sens    := 1;
      end
      else Dec(ARow);
    end
    else if Sens = 1 then
        begin
      if (ARow >= Lim) then
      begin
          ARow := Lim;
        Sens    := -1;
    end
      else Inc(ARow);
        end;
    GGrilleVar.GS.Row := Arow;
    GGrilleVar.GS.Col := Acol;
      end;

  GGrilleVar.GS.SynEnabled := OldEna;

  if (Arow <> RowFirst) then
  Begin
  	GSRowExit (self,RowFirst,cancel,false);
  	if Cancel then
    begin
    	GSRowEnter (self,RowFirst,cancel,false);
    end else
    begin
    	GSRowEnter (self,Arow,cancel,false);
    end;
  End;

end;

function TOF_BTSAISIEVAR.ZoneAccessible(ACol, ARow: Integer): boolean;
VAR tobl : TOB;
begin

  Result := false;

  TOBL := GGrilleVar.TOBG.detail[ARow-1];

  if TOBL = nil then Exit;

  if Acol <> SG_VALEUR then GGrilleVar.GS.col := SG_VALEUR;

  if tobl.GetValue('BVD_CODEVARIABLE')= '' then exit;

  if tobl.GetValue('SAISISSABLE') = '-' then  exit;

  if GrilleVar.Row = 0 then  exit;

  Result := True;

end;

//
//
//
Procedure TOF_BTSAISIEVAR.Controlechamp(Champ, Valeur : String);
begin

end;

Procedure TOF_BTSAISIEVAR.InitGrillevar;
begin

  //Une recherche de la grille au niveau de la table des liste serait bien venu !!!
  GGrilleVar                := TGestionGS.Create;

  GGrilleVar.Ecran          := TFVierge(Ecran);
  GGrilleVar.GS             := GrilleVar;
  GGrilleVar.GS.DBIndicator := True;
  GGrilleVar.TOBG           := TOBPartielle;

  // Définition de la liste de saisie pour la grille Détail
  GGrilleVar.ColNamesGS     := 'BVD_CODEVARIABLE;BVD_LIBELLE;BVD_VALEUR';
  GGrilleVar.alignementGS   := 'G.0  ---;C.0  ---;G.0  ---; ';
  GGrilleVar.LibColNameGS   := 'Code;Désignation;Valeur; ';
  GGrilleVar.LargeurGS      := '50;100;50; ';
  //
  GGrilleVar.RowHeight := 18;

  GGrilleVar.DessineGrille;

  THSystemMenu(getControl('HMtrad')).ResizeGridColumns(Grillevar);

  GrilleVar.FixedRows := 1;

end;


procedure TOF_BTSAISIEVAR.ChargementEcran;
var indice  : integer;
    LibChamp: String;
    NumLig  : Integer;
    Cancel  : Boolean;
    Arow, Acol : Integer;
begin

  Cancel := False;

  With GGrilleVar do
  begin
    if TOBG.Detail.count <> 0 then
      GS.RowCount := TOBG.detail.count + 1
    else
      Exit;
    //
    GS.DoubleBuffered := true;
    GS.BeginUpdate;
    //
    TRY
      GS.SynEnabled := false;
      for Indice := 0 to TOBG.detail.count -1 do
      begin
        GS.row := Indice+1;
        TOBG.PutGridDetail(GS,False,True,ColNamesGS);
      end;
    FINALLY
      GS.SynEnabled := true;
      GS.EndUpdate;
    END;
    TFVierge(Tform).HMTrad.ResizeGridColumns(GS);
    ARow := 1;
    ACol := SG_VALEUR;
    GS.Row := Arow;
    GS.Col := Acol;
    Sens := 1;
    ZoneSuivanteOuOk(Acol, Arow, Cancel);
  end;



end;


procedure TOF_BTSAISIEVAR.ValiderClick (Sender: TObject);
begin

  MAJTobVariables;

  TobVariables.putValue('RESULT','X');
  TobVariables.putValue('VALIDE','-');

  OnUpdate;

end;

procedure TOF_BTSAISIEVAR.ValiderXLClick (Sender: TObject);
Var Ind : Integer;
    TOBL: TOB;
    TOBLPartielle : TOB;
begin

  MAJTobVariables;

  TobVariables.putValue('RESULT','X');
  TobVariables.putValue('VALIDE','X');

  OnUpdate;

end;

Procedure TOF_BTSAISIEVAR.MAJTobVariables;
Var Ind           : Integer;
    TOBL          : TOB;
    TOBLPartielle : TOB;
    NumArtLig     : String;
    StSQL         : String;
begin

  //On Recharge TobVariable avec TobPartielle
  For Ind := 0 To TobPartielle.detail.count -1  do
  begin
    TOBLPartielle := TobPartielle.detail[Ind];
    TOBL := TOBVariables.findfirst(['BVD_NATUREPIECE', 'BVD_SOUCHE', 'BVD_NUMERO', 'BVD_INDICE',
                                    'BVD_NUMORDRE','BVD_UNIQUEBLO','BVD_ARTICLE','BVD_CODEVARIABLE'],
                                   [TOBLPartielle.GetValue('BVD_NATUREPIECE'), TOBLPartielle.GetValue('BVD_SOUCHE'),
                                   TOBLPartielle.GetValue('BVD_NUMERO'), TOBLPartielle.GetValue('BVD_INDICE'),
                                   TOBLPartielle.GetValue('BVD_NUMORDRE'),TOBLPartielle.GetValue('BVD_UNIQUEBLO'),
                                   TOBLPartielle.GetValue('BVD_ARTICLE'), TOBLPartielle.GetValue('BVD_CODEVARIABLE')],False);
    if TOBL <> nil then
    begin
      TOBL.PutValue('BVD_VALEUR', GGrilleVar.GS.Cells[SG_VALEUR, Ind+1]);
      NumArtLig := TOBLPartielle.GetValue('BVD_NATUREPIECE') + '-' + TOBLPartielle.GetValue('BVD_SOUCHE') + '-' + IntToStr(TOBLPartielle.GetValue('BVD_NUMERO')) + '-' + IntToStr(TOBLPartielle.GetValue('BVD_INDICE')) + '-' + IntToStr(TOBLPartielle.GetValue('BVD_NUMORDRE')) + '-' + IntToStr(TOBLPartielle.GetValue('BVD_UNIQUEBLO'));
      If (TOBL.GetValue('BVD_ARTICLE') = NumArtLig) Then
      begin
        StSQL := 'UPDATE BVARIABLES SET BVA_VALEUR="' + TOBL.GetValue('BVD_VALEUR') + '"'
               + ' WHERE BVA_TYPE="L" '
               + '   AND BVA_ARTICLE="' + NumArtLig
               + '"  AND BVA_CODEVARIABLE="' + TOBL.GetValue('BVD_CODEVARIABLE') + '"';
        ///
        ExecuteSQL(STSQL); 
      end;
    end;
  end;

end;


procedure TOF_BTSAISIEVAR.ToutVoirClick(Sender: Tobject);
begin

  //ToutVoir permet d'afficher toutes les lignes même les ligne non affichable....
  ChargementTobVariablePartielle;

  ChargementEcran;

end;


Initialization
  registerclasses ( [ TOF_BTSAISIEVAR ] ) ;
end.

