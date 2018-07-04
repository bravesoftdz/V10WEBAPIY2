unit EtudesStruct;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,HEnt1,Dicobtp,
{$IFDEF EAGLCLIENT}
  maineagl,
{$ELSE}
  Doc_Parser,DBCtrls, Db, {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} fe_main,
{$ENDIF}
  SaisUtil,Menus, Grids, Hctrls, HTB97, ExtCtrls,UTOB, StdCtrls,EtudesUtil,HmsgBox,EtudesExt,
  TntGrids;

procedure DefinitionStructXls (TOBEtude:TOB;var TOBStruct : TOB; Arow : integer; Action:TActionFiche);

type
  TFEtudesStruct = class(TForm)
    Panel1: TPanel;
    DockBottom: TDock97;
    Valide97: TToolbar97;
    BValider: TToolbarButton97;
    BAbandon: TToolbarButton97;
    BAide: TToolbarButton97;
    GS: THGrid;
    Panel2: TPanel;
    GroupBox1: TGroupBox;
    ApplicAffaire: TRadioButton;
    ApplicClient: TRadioButton;
    Outils97: TToolbar97;
    BDelete: TToolbarButton97;
    BNewligne: TToolbarButton97;
    PopupDel: TPopupMenu;
    LigCurDel: TMenuItem;
    StructDel: TMenuItem;
    procedure FormCreate(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure GSCellEnter(Sender: TObject; var ACol, ARow: Integer;
      var Cancel: Boolean);
    procedure GSCellExit(Sender: TObject; var ACol, ARow: Integer;
      var Cancel: Boolean);
    procedure BValiderClick(Sender: TObject);
    procedure BAbandonClick(Sender: TObject);
    procedure StructDelClick(Sender: TObject);
    procedure BNewligneClick(Sender: TObject);
    procedure GSEnter(Sender: TObject);
    procedure LigCurDelClick(Sender: TObject);
    procedure GSDblClick(Sender: TObject);
  private
    { Déclarations privées }
    EtudeExt : TEtudesExt;
    TOBEtude,TOBStrDoc,TOBSTruct: TOB;
    RowEtude : integer;
    GS_TYPE,GS_FEUILLE,GS_VALEUR,GS_OBLIG : integer;
    Action:TActionFiche;
    cellcurr : string;
    AppAffaireInit : boolean;
    procedure AlloueTobs;
    procedure ChargeTobs;
    procedure definiGrille;
    procedure GSPostDrawCell(ACol, ARow: Integer; Canvas: TCanvas;
      AState: TGridDrawState);
    procedure LibereTobs;
    function remplitBlob: string;
    function ControlePlage: boolean;
    function ControleSaisie: boolean;
    function controleTExcel(Valeur : string): boolean;
    procedure PositionneLigne(Arow:integer; Acol: integer=-1);
    procedure ZoneSuivanteOuOk(var ACol, ARow: Integer; var Cancel: boolean);
    function ZoneAccessible(ACol, ARow: Integer): boolean;
    function TrouveLigneCtrl(TOBCtrl: TOB; feuille: string;mini: integer): integer;
  public
    { Déclarations publiques }
  end;

const
	TexteMsg: array[1..9] of string 	= (
          {1}        'Les plages de données à traiter sont incorrecte',
          {2}        'La colonne N° de prix est incorrecte',
          {3}        'La colonne Libellé est incorrecte' ,
          {4}        'La colonne quantité est incorrecte' ,
          {5}        'La colonne unité est incorrecte' ,
          {6}        'La colonne Prix unitaire est incorrecte',
          {7}        'La colonne montant est incorrecte',
          {8}        'Le nom de la feuille excel n''est pas définie',
          {9}        'La colonne Code Article/Prestation est incorrecte'
          );

var
  FEtudesStruct: TFEtudesStruct;

implementation

{$R *.DFM}

procedure DefinitionStructXls (TOBEtude:TOB;var TOBStruct : TOB; Arow : integer; Action:TActionFiche);
var X: TFEtudesStruct;
begin
X := TFEtudesStruct.Create (application);
X.TOBEtude := TOBEtude;
X.TOBSTruct := TOBStruct;
X.RowEtude := Arow;
X.action := Action;
TRY
   X.ShowModal;
FINALLY
   TOBStruct := X.TOBStruct;
   X.free;
END;
end;

procedure  TFEtudesStruct.AlloueTobs;
begin

  TOBStrDoc := ajouteTobstructure (nil,-1);

end;

procedure  TFEtudesStruct.ChargeTobs;
var QQ : TQuery;
    req : string;
    TOBL : TOB;
begin
  TOBL := TOBEtude.Detail[RowEtude-1];

  if TOBStruct <> nil then
  begin
    TOBStrDOc.dupliquer (TOBStruct,true,true);
    AppAffaireInit := true;
    exit;
  end;

  Req := 'SELECT * FROM BSTRDOC WHERE BSD_TIERS="'+TOBEtude.getvalue('AFF_TIERS')+'" AND ';
  REq := REq + 'BSD_TYPE="'+TOBL.getValue('BDE_TYPE')+'" AND ';
  REq := REq + 'BSD_NATUREDOC="'+TOBL.getValue('BDE_NATUREDOC')+'"';

  QQ := Opensql (req,true,-1, '', True);

  if not QQ.eof then
  BEGIN
    TOBStrDoc.SelectDB ('',QQ);
    ApplicClient.Checked := true;
    AppAffaireInit := false;
  END;

   ferme(QQ);

end;


procedure TFEtudesStruct.definiGrille;
begin
  GS.cells[GS_TYPE,0]       := 'Type';
  GS.ColWidths [GS_TYPE]    := 199;
  GS.ColLengths [GS_TYPE]   := -1;

  GS.cells [GS_FEUILLE,0]   := 'Feuille excel';
  GS.ColWidths [GS_FEUILLE] := 100;

  GS.Cells[GS_VALEUR,0]     := 'Valeur';
  GS.ColWidths [GS_VALEUR]  := 398;

  GS.Cells[GS_OBLIG,0]      := 'Trait.';
  GS.ColWidths [GS_OBLIG]   := 30;
  GS.ColLengths [GS_OBLIG]  := -1;
  GS.ColTypes [GS_OBLIG]    :='B' ;
  GS.colaligns[GS_OBLIG]    := tacenter;
  GS.colformats[GS_OBLIG]   := inttostr(Integer(csCoche));

end;


procedure TFEtudesStruct.FormCreate(Sender: TObject);
begin
  GS_TYPE     := 0;
  GS_FEUILLE  := 1;
  GS_VALEUR   := 2;
  GS_OBLIG    := 3;
  //
  AlloueTobs;

  definiGrille;

  GS.PostDrawCell := GSPostDrawCell;

  EtudeExt := TEtudesExt.Create;

end;

procedure TFEtudesStruct.FormShow(Sender: TObject);
var Index : integer;
begin
  ChargeTobs;

  caption := caption + ' '+rechdom ('BTNATUREDOC',TOBEtude.detail[RowEtude-1].getValue('BDE_NATUREDOC'),false);

  EtudeExt.createStruct (TOBStrDoc,TOBEtude,RowEtude);
  EtudeExt.RempliStruct (TOBStrDoc);

  GS.rowCount := TOBStrDoc.detail.count+1;

  for Index := 0 to TOBStrDoc.detail.count -1 do TOBStrDoc.detail[Index].PutLigneGrid (GS,Index+1,false,false,'NATURE;FEUILLE;VALEUR;OBLIG');

  AffecteGrid (GS,Action) ;
end;

procedure TFEtudesStruct.LibereTobs;
begin
TOBStrDoc.free;
end;

procedure TFEtudesStruct.GSPostDrawCell(ACol, ARow: Integer; Canvas: TCanvas;
  AState: TGridDrawState);
Var ARect : TRect ;
    chaine : string;
    PosG : integer;
begin
if GS.RowHeights[ARow]<=0 then Exit ;
//if ARow>GS.TopRow+GS.VisibleRowCount-1 then Exit ;
ARect:=GS.CellRect(ACol,ARow) ;
GS.Canvas.Pen.Style:=psSolid ; GS.Canvas.Pen.Color:=clgray ;
GS.Canvas.Brush.Style:=BsSolid ;
if ((Acol=GS_Type) and (Arow >= GS.fixedRows)) then
   begin
   if GS.cells[Acol,Arow]= '000' then chaine := 'Plages des lignes à traiter';
   if GS.cells[Acol,Arow]= '001' then chaine := 'Colonne du N° prix';
   if GS.cells[Acol,Arow]= '002' then chaine := 'Colonne des désignations';
   if GS.cells[Acol,Arow]= '003' then chaine := 'Colonne des Quantités';
   if GS.cells[Acol,Arow]= '004' then chaine := 'Colonne des Unités';
   if GS.cells[Acol,Arow]= '005' then chaine := 'Colonne des Prix unitaire';
   if GS.cells[Acol,Arow]= '006' then chaine := 'Colonne des Montants';
   // Modified by f.vautrain 24/10/2017 15:30:16 - FS#2753 - NCN : développement selon devis ML-171509-3A
   if GS.cells[Acol,Arow]= '007' then chaine := 'Colonne des Code Art/Presta.';
//   if GS.cells[Acol,Arow]= '007' then chaine := 'Nom de la feuille excel';

   PosG := Arect.Left + ((ARect.Right - ARect.left - canvas.textwidth(chaine)) div 2);
   canvas.FillRect (Arect);
   canvas.TextOut (PosG,ARect.top + 1,Chaine);
   end;

end;

procedure TFEtudesStruct.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
LibereTobs;
EtudeExt.Free;
end;

procedure TFEtudesStruct.GSCellEnter(Sender: TObject; var ACol,
  ARow: Integer; var Cancel: Boolean);
begin
if action = taconsult then exit;
ZoneSuivanteOuOk(ACol,ARow,Cancel) ;
if cancel then exit;
cellcurr := GS.cells[GS.col,GS.row];
end;

procedure TFEtudesStruct.GSCellExit(Sender: TObject; var ACol,
  ARow: Integer; var Cancel: Boolean);
var TOBL : TOB;
begin
  if action = taconsult then exit;
  //
  if GS.Cells[ACol,ARow]=CellCurr then exit;

  if CellCurr <> GS.Cells [Acol,Arow] then
  begin
  if Acol = GS_VALEUR then
  begin
    TOBL := TOBStrDoc.detail[Arow-1];
    if (TOBL <> nil) then
    begin
      TOBL.PutValue('VALEUR',Uppercase(GS.cells[Acol,Arow]));
      GS.cells[Acol,Arow] := Uppercase(GS.cells[Acol,Arow]);
      // Modified by f.vautrain 24/10/2017 15:30:16 - FS#2753 - NCN : développement selon devis ML-171509-3A
      if TOBL.GetValue('NATURE') = '007' then
      begin
        if GS.cells[Acol,Arow] = '' then
          TOBL.PutValue('OBLIG','-')
        else
          TOBL.PutValue('OBLIG','X');
        PositionneLigne(ARow, ACol);
        TOBL.PutLigneGrid (GS,GS.row,false,false,'NATURE;FEUILLE;VALEUR;OBLIG');
      end;
    end;
  end;
  if Acol = GS_FEUILLE then
    begin
    TOBL := TOBStrDoc.detail[Arow-1];
    if (TOBL <> nil) then TOBL.PutValue('FEUILLE',GS.cells[Acol,Arow]);
    end;
  end;
end;

Function TFEtudesStruct.remplitBlob:string;
var TOBSav,TOBS : tob;
    Indice,Index : integer;
begin
Index := 0;
TOBSav := TOB.create ('SAUVEGARDE',nil,-1);
for Indice := 0 to TOBStrDoc.detail.count -1 do
    begin
    if TOBSTRDOC.detail[Indice].getvalue('NATURE') = '000' then
       begin
       TOBSTRDOC.detail[Indice].putvalue('INDICE',Index);
       Inc(Index);
       end;
    TOBS:=TOB.Create ('STRUCTLIG',tobsav,-1);
    TOBS.dupliquer (TOBStrDoc.detail[Indice],true,true);
    end;
TOBStrDoc.ClearDetail;
result :=Tobsav.SaveToBuffer (true,true,'');
TOBsav.free;
end;

function TFEtudesStruct.TrouveLigneCtrl (TOBCtrl : TOB;feuille : string; mini : integer): integer;
var TOBL : TOB;
    Indice : integer;
begin
result := -1;
if TOBCtrl.detail.count = 0 then exit;
for Indice := 0 to TOBCtrl.detail.count -1 do
    begin
    TOBL := TOBCtrl.detail[indice];
    if feuille < TOBL.GetValue('FEUILLE') then
       begin
       result := Indice;
       break;
       end;
    if (feuille = TOBL.GetValue('FEUILLE')) and (mini < TOBL.GetValue('MINI')) then
       begin
       result := Indice;
       break;
       end;
    end;
end;

function TFEtudesStruct.ControlePlage : boolean;
var TOBCtrl,TOBC,TOBRef : TOB;
    chaine,schaine : string;
    mini,maxi : integer;
    indice,reference,Index : integer;
begin
  Indice := GS.row -1;
  result := true;
  //TOBL := TOBEtude.Detail[RowEtude-1];
  //if TOBL.GetValue('BDE_QUALIFNAT')='PRINC1' then exit;
  TOBCtrl := TOB.create ('TOBMERE',nil,-1);
  for Index := 0 to TOBSTRDOC.detail.count -1 do
  begin
    if TOBSTRDOC.Detail[Index].getValue('NATURE') <> '000' then break;
    if TOBStrDOC.detail[Index].getValue('FEUILLE') = '' then
       begin
       // verification des nom de feuilles
       PGIBoxAF (TexteMsg[8],caption);
       result := false;
       GS.row := Index +1;
       break;
       end;
    Chaine := TOBStrDoc.Detail[Index].getValue('VALEUR');
    if chaine = '' then BEGIN result := false; GS.row := Index +1;break; END;
    repeat
      schaine := ReadTokenSt(Chaine);
      if schaine <> '' then
         begin
         determineMiniMaxi (schaine,mini,maxi);
         if (mini < 0) or (maxi < 0) then BEGIN result := false;GS.row := Index +1; break; END;
         if mini > maxi then BEGIN Result := false;GS.row := Index +1; break; END;
         // tri de la tob pour ordonner en feuille+mini
         Indice := TrouveLigneCtrl (TOBCtrl,TOBStrDOC.detail[Index].getValue('FEUILLE'),mini);
         TOBC:= TOB.create ('DONNEE',TOBCtrl,Indice);
         TOBC.addChampSupValeur('FEUILLE',TOBStrDOC.detail[Index].getValue('FEUILLE'));
         TOBC.addChampSupValeur('MINI',mini);
         TOBC.addChampSupValeur('MAXI',maxi);
         TOBC.addChampSupValeur('LIGNE',Index+1);
         end;
    until schaine = '';
  end;

  if not result then
  BEGIN
    if TOBCtrl<>nil then TOBCtrl.free;
    GS.row := Indice+1;
    Exit;
  END;

  //if TOBCtrl.detail.count = 0 then BEGIN result := false; TOBCtrl.Free; Exit; END;
  for reference := 0 To TOBCtrl.detail.count -1 do
  begin
    TOBRef := TOBCtrl.detail[reference];
    if Reference = TOBCtrl.detail.count -1 then break;
    for Indice := reference+1 to TOBCtrl.detail.count -1 do
    begin
      TOBC := TOBCtrl.detail[Indice];
      if TOBC.GetValue('FEUILLE') <> TOBRef.GetValue('FEUILLE') then break;
      if TOBC.GetValue ('MINI') < TOBRef.getValue('MAXI') then
      BEGIN
        REsult:= false;
        GS.row := TOBC.GetValue('LIGNE');
        break;
      END;
    end;
  end;
  //
  TOBCtrl.free;
  //
end;

function TFEtudesStruct.controleTExcel (Valeur : string) : boolean;
var indice,Arow : integer;
    Control,TypeA,TypeS: string;
    TOBL : TOB;
    videautorise : boolean;
begin
Arow := -1;
TOBL:=Nil;
for Indice := 0 to TOBStrDoc.detail.count -1 do
    begin
    TOBL := TOBStrDoc.detail[Indice];
    if TOBL.GetValue('NATURE')=Valeur then
       begin
       Arow := Indice +1;
       Control := TOBL.GetValue('VALEUR');
       break;
       end;
    end;

  if (TOBL = nil) or (Arow < 0 ) then
  BEGIN
    result := false;
    exit;
  END;

  typea := '';
  result := true;
  VideAutorise := (TOBL.GetValue('OBLIG')<>'X') ;

  if (Control = '') and (not videautorise) then BEGIN result:= false; GS.row := Arow; Exit; END;

  for Indice:= 1 to length (Control) do
  begin
    typeS := '';
    if (copy(Control,Indice,1)>= 'A') and (copy(Control,Indice,1) <= 'Z') then TypeS := 'A';
    //     if (copy(Control,Indice,1)>= '0') and (copy(Control,Indice,1) <= '9') then TypeS := 'N';
    if (TypeS <> 'A') and (TypeS <> 'N') then BEGIN result :=false; GS.row := Arow; break; END;
    if TypeA = '' then BEGIN TypeA := TypeS; COntinue; END;
    if TypeS <> TypeA then BEGIN result := false; GS.row := Arow; break; END; // on melange pas les genres....
  end;

end;

function TFEtudesStruct.ControleSaisie:boolean;
begin
result := false;
if not ControlePlage then
   begin
   PGIBoxAF (TexteMsg[1],caption);
   exit;
   end;
if not controleTExcel ('001') then
   begin
   // zone de N° Prix
   PGIBoxAF (TexteMsg[2],caption);
   exit;
   end;
if not controleTExcel ('002') then
   begin
   // libellé
   PGIBoxAF (TexteMsg[3],caption);
   exit;
   end;
if not controleTExcel ('003') then
   begin
   // quantite
   PGIBoxAF (TexteMsg[4],caption);
   exit;
   end;
if not controleTExcel ('004')then
   begin
   // unite
   PGIBoxAF (TexteMsg[5],caption);
   exit;
   end;
if not controleTExcel ('005') then
   begin
   // prix unitaire
   PGIBoxAF (TexteMsg[6],caption);
   exit;
   end;
if not controleTExcel ('006')then
   begin
   // montant
   PGIBoxAF (TexteMsg[7],caption);
   exit;
   end;
  // Modified by f.vautrain 24/10/2017 15:30:16 - FS#2753 - NCN : développement selon devis ML-171509-3A
  if not controleTExcel ('007')then
  begin
    // Code Article/Prestation
    PGIBoxAF (TexteMsg[9],caption);
    exit;
  end;

   result := true;
end;

procedure TFEtudesStruct.BValiderClick(Sender: TObject);
var Acol, Arow : integer;
    Cancel : boolean;
    TOBL : TOB;
begin
Acol := GS.col;
Arow := GS.row;
Cancel := false;
GSCellexit (self,Acol,Arow,cancel);
if not ControleSaisie then exit;
TOBL := TOBEtude.detail[RowEtude-1];
if ApplicAffaire.Checked then
   BEGIN
   TOBSTRDOC.putValue('BSD_TIERS',TOBEtude.getValue('AFF_TIERS'));
   TOBStrDOC.PutValue('BSD_AFFAIRE',TOBEtude.getValue('AFF_AFFAIRE'));
   TOBStrDOC.PutValue('BSD_ORDRE',TOBL.getValue('BDE_ORDRE'));
   TOBStrDOC.PutValue('BSD_TYPE',TOBL.getValue('BDE_TYPE'));
   TOBStrDOC.PutValue('BSD_NATUREDOC',TOBL.getValue('BDE_NATUREDOC'));
   TOBStrDOC.PutValue('BSD_INDICE',TOBL.getValue('BDE_INDICE'));
   END ELSE
   BEGIN
   TOBSTRDOC.putValue('BSD_TIERS',TOBETUDE.getValue('AFF_TIERS'));
   TOBStrDOC.PutValue('BSD_AFFAIRE','');
   TOBStrDOC.PutValue('BSD_ORDRE',0);
   TOBStrDOC.PutValue('BSD_TYPE',TOBL.getValue('BDE_TYPE'));
   TOBStrDOC.PutValue('BSD_NATUREDOC',TOBL.getValue('BDE_NATUREDOC'));
   TOBStrDOC.PutValue('BSD_INDICE',0);
   END;
TOBSTrDOC.putvalue('BSD_DESCRIPT',remplitBlob);
if ApplicAffaire.Checked  then
   begin
   if TobStruct = nil then
      begin
      TOBStruct := ajouteTobstructure (nil,-1)
      end;
      TOBStruct.Dupliquer (TOBSTRDOC,true,true);
   end else
   begin
   if TobStruct <> nil then BEGIN TOBStruct.free; TOBSTruct := nil; END;
   TOBStrDOC.InsertOrUpdateDB(true);
   end;
close;
end;

procedure TFEtudesStruct.BAbandonClick(Sender: TObject);
begin
close;
end;

procedure TFEtudesStruct.StructDelClick(Sender: TObject);
begin
if PGIAsk ('Désirez-vous réellement supprimer cette structure de document ?',caption) = mrYes then
   begin
   if TobStruct <> nil then BEGIN TOBStruct.free; TOBSTruct := nil; END;
   if not (ApplicAffaire.Checked ) then
   begin
   TOBStrDOC.DELETEDB;
   end;
   close;
   end;
end;

procedure TFEtudesStruct.BNewligneClick(Sender: TObject);
var TOBL : TOB;
    Indice,Index,Reference : integer;
begin
Index := 0;
reference := -1;
For Indice := 0 to TOBStrDoc.detail.count -1 do
    begin
    TOBL := TOBStrDoc.Detail[Indice];
    if TOBL = nil then continue;
    if TOBL.GetValue('NATURE') <> '000' then break;
    if TOBL.GetValue('NATURE') = '000' then
       begin
       reference := Indice+1;
       Index := TOBL.GetValue('INDICE')+1;
       end;
    end;
if reference = -1 then
   begin
   Index := 0;
   reference := 0;
   end;
TOBL := TOB.Create ('STRUCTLIG',TOBSTRDOC,reference);
TOBL.addchampsupValeur('NATURE','000');
TOBL.addchampsupValeur('FEUILLE','');
TOBL.addchampsupValeur('INDICE',Index);
TOBL.addchampsupValeur('VALEUR','');
TOBL.addchampsupValeur('OBLIG','X');

GS.InsertRow (reference+1);
TOBL.PutLigneGrid (GS,reference+1,false,false,'NATURE;FEUILLE;VALEUR;OBLIG');
PositionneLigne (reference+1);
end;

procedure TFEtudesStruct.PositionneLigne (Arow:integer;Acol : integer=-1);
var GRow,GCol : integer;
    Cancel : boolean;
    Synchro : boolean;
begin

  Cancel := false;

  Gcol := GS.col;
  Grow := GS.row;
  GSCellExit (self,Gcol,Grow,cancel);
  Synchro := GS.synenabled;
  GS.synEnabled := false;
  Grow := Arow;
  if Acol <> -1 then GCol := Acol else Gcol := GS.Fixedcols ;
  GS.col := Gcol;
  GS.row:= Grow;
  cancel := false;
  GS.CacheEdit;
  GSCellenter (Self,Gcol,Grow,Cancel);
  GS.col := Gcol;
  GS.row:= Grow;
  cellcurr := GS.cells[Gcol,Grow];
  GS.SynEnabled := synchro;
  GS.MontreEdit;

end;


procedure TFEtudesStruct.ZoneSuivanteOuOk ( Var ACol,ARow : Longint ; Var Cancel : boolean ) ;
Var Sens,ii,Lim : integer ;
    OldEna,ChgLig,ChgSens  : boolean ;
BEGIN
  OldEna:=GS.SynEnabled ;
  GS.SynEnabled:=False ;
  Sens:=-1 ;
  ChgLig:=(GS.Row<>ARow) ;
  ChgSens:=False ;
  if GS.Row>ARow then Sens:=1 else if ((GS.Row=ARow) and (ACol<=GS.Col)) then Sens:=1 ;
  ACol:=GS.Col ; ARow:=GS.Row ; ii:=0 ;

  While Not ZoneAccessible(ACol,ARow)  do
  BEGIN
   Cancel:=True ; inc(ii) ; if ii>500 then Break ;
   if Sens=1 then
      BEGIN
      Lim:=TOBEtude.detail.count ;
      if ((ACol=GS.ColCount-1) and (ARow>=Lim)) then
         BEGIN
         if ChgSens then Break else BEGIN Sens:=-1 ; Continue ; ChgSens:=True ; END ;
         END ;
      if ChgLig then BEGIN ACol:=GS.FixedCols-1 ; ChgLig:=False ; END ;
      if ACol<GS.ColCount-1 then Inc(ACol) else BEGIN Inc(ARow) ; ACol:=GS.FixedCols ; END ;
      END else
      BEGIN
      if ((ACol=GS.FixedCols) and (ARow=1)) then
         BEGIN
         if ChgSens then Break else BEGIN Sens:=1 ; Continue ; END ;
         END ;
      if ChgLig then BEGIN ACol:=GS.ColCount ; ChgLig:=False ; END ;
      if ACol>GS.FixedCols then Dec(ACol) else BEGIN Dec(ARow) ; ACol:=GS.ColCount-1 ; END ;
      END ;
  END ;

  GS.SynEnabled:=OldEna ;
END ;

Function TFEtudesStruct.ZoneAccessible ( ACol,ARow : Longint) : boolean ;
var TOBL : TOB;
BEGIN
Result:=True ;
if GS.ColLengths [Acol] = -1 then BEGIN Result := false;exit;END;
TOBL := TOBStrDoc.detail[Arow-1];
if (Acol = GS_FEUILLE) and (TOBL.GetValue('NATURE') <> '000') then result:= false;
END ;

procedure TFEtudesStruct.GSEnter(Sender: TObject);
var Arow, Acol : Integer;
    cancel : boolean;
begin
Arow := 1;
Acol := GS_FEUILLE;
Cancel := false;
GSCellenter (Self,Acol,Arow,Cancel);
GS.row := Arow;;
GS.col := Acol;
//PositionneLigne (Arow,Acol);
end;

procedure TFEtudesStruct.LigCurDelClick(Sender: TObject);
var TOBl : TOB;
begin
TOBL := TOBStrDoc.Detail[GS.row-1];
if TOBL.GetValue('NATURE') <> '000' then exit;
if PGIAsk ('Désirez-vous réellement supprimer cette feuille ?',caption) = mrYes then
   begin
   TOBL.free;
   GS.DeleteRow (GS.row);
//   if TOBStrDoc.findfirst (['NATURE'],['000'],true) = nil then BEGIN BNewligneClick (self); exit; END;
   PositionneLigne (GS.row);
   end;
end;

procedure TFEtudesStruct.GSDblClick(Sender: TObject);
var TOBL : TOB;
begin
TOBL := TOBStrDoc.Detail[GS.row-1];
if TOBL.GetValue('NATURE') = '001' then exit;
if TOBL.GetValue('OBLIG')='X' then TOBL.PutValue('OBLIG','-')
                              else TOBL.PutValue('OBLIG','X');
TOBL.PutLigneGrid (GS,GS.row,false,false,'NATURE;FEUILLE;VALEUR;OBLIG');
end;

end.
