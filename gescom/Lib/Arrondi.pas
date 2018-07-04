unit arrondi;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
{$IFNDEF EAGLCLIENT}
  FichList, HSysMenu, hmsgbox, Db, {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} Hqry, StdCtrls, HTB97,
  DBCtrls, ExtCtrls, HPanel, Grids, DBGrids, HDB, Hent1, UIUtil, Mask,
{$IFDEF V530}
  EdtEtat,
{$ELSE}
  EdtREtat,
  {$ENDIF}
{$ELSE}
	UtilEagl,eFichList,
{$ENDIF}
  Hctrls, Math, SaisUtil, UTOB, LookUp, ImgList,Comctrls, Buttons,TarifUtil,
  ADODB;

procedure EntreeArrondi (Action : TActionFiche) ;

Const
     ARR_Type    : integer  = 0;
     ARR_Curs    : integer  = 1;
     ARR_Mont    : integer  = 2;
     ARR_Poid    : integer  = 3;
     ARR_Libe    : integer  = 4;
     ARR_Cons    : integer  = 5;
     NbRowsInit  : integer  = 25;
     NbRowsPlus  : integer  = 20 ;
     Curseur     : string   = 'Ø';

type
  TFArrondi = class(TFFicheListe)
    PEntete: TPanel;
    TGAR_CODEARRONDI: TLabel;
    GAR_CODEARRONDI: THDBEdit;
    TGAR_LIBELLE: THLabel;
    G_GAR: THGrid;
    MsgBox: THMsgBox;
    InitLibelle: TEdit;
    TTest: TToolWindow97;
    NValeur: THNumEdit;
    TNVal: THLabel;
    TNResult: THLabel;
    NResult: THNumEdit;
    btest: TToolbarButton97;
    procedure FormShow(Sender: TObject);
    procedure ChargeEnreg ; Override ;
    procedure FormCreate(Sender: TObject);
    procedure G_GAREnter(Sender: TObject);
    procedure G_GARRowEnter(Sender: TObject; Ou: Integer;
      var Cancel: Boolean; Chg: Boolean);
    procedure G_GARRowExit(Sender: TObject; Ou: Integer;
      var Cancel: Boolean; Chg: Boolean);
    procedure G_GARCellEnter(Sender: TObject; var ACol, ARow: Integer;
      var Cancel: Boolean);
    procedure G_GARCellExit(Sender: TObject; var ACol, ARow: Integer;
      var Cancel: Boolean);
    procedure G_GARElipsisClick(Sender: TObject);
    procedure FormKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure BValiderClick(Sender: TObject) ; override ;
    procedure STaDataChange(Sender: TObject; Field: TField);
    procedure STaUpdateData(Sender: TObject);
    procedure bDefaireClick(Sender: TObject);
    procedure BNextClick(Sender: TObject);
    procedure BPrevClick(Sender: TObject);
    procedure BLastClick(Sender: TObject);
    procedure BFirstClick(Sender: TObject);
    procedure FListeEnter(Sender: TObject);
    procedure GAR_CODEARRONDIExit(Sender: TObject);
    procedure BinsertClick(Sender: TObject) ; override ;
    procedure BDeleteClick(Sender: TObject);
    procedure BImprimerClick(Sender: TObject);
    procedure btestClick(Sender: TObject);
    procedure TTestClose(Sender: TObject);
    procedure NValeurKeyUp(Sender: TObject; var Key: Word;
      Shift: TShiftState);
  private
    { Déclarations privées }
    CodeArrPrec,LesColFliste : String;
    DEV          : RDEVISE ;
    TOBArr       : TOB ;
    BValide      : Boolean;
    LgPrec       : Integer;
  public
    { Déclarations publiques }
    Action      : TActionFiche ;
    iTableLigne : Integer;
// Initialisation
    procedure InitialiseGrille;
    procedure InitLesCols;
    procedure InitialiseEntete;
// Actions liées au Grid
    procedure EtudieColsListe ;
    procedure FormateZoneSaisie (ACol,ARow : Longint) ;
    procedure DessineCell (ACol,ARow : Longint; Canvas : TCanvas;
                                     AState: TGridDrawState);
// Lignes
    procedure InitialiseLigne(Arow : integer);
    procedure GetGArrondiRecherche(GAR_GArrondi : THCritMaskEdit);
    procedure AfficheLigne(Arow : integer);
    function  ControlLigne(Valider: boolean): boolean;
    function  ControlPentete(Valider : boolean) : boolean;
    procedure SupprimeLigne(Arow: integer);
    function  LigneVide(Arow: integer; var ACol: integer): boolean;
    procedure CreerTobLigne(Arow : integer);
    function  GetTobLigne(Arow :integer):TOB;
    procedure DepileTobLigne;
// Champs lignes
    procedure TraiterMontant(Acol :integer; Arow :integer);
    procedure TraiterPoids(Acol :integer; Arow :integer);
    procedure TraiterMethode(Acol :integer; Arow :integer);
    procedure TraiterConstante(Acol :integer; Arow :integer);
//  Actions liées au Boutons
    procedure BoutonsNavigator(Sender: TObject;Button: TNavigateBtn);
    function  TobModifier : Boolean;
    function  Sauvegarder : boolean;
 end;
var
  FArrondi: TFArrondi;

implementation

{$R *.DFM}

procedure EntreeArrondi (Action : TActionFiche) ;
var FF : TFArrondi ;
    PPANEL  : THPanel ;
begin
SourisSablier;
FF := TFArrondi.Create(Application) ;
FF.Action:=Action ;
PPANEL := FindInsidePanel ;
if PPANEL = Nil then
   BEGIN
    try
      FF.ShowModal ;
    finally
      FF.Free ;
    end ;
   SourisNormale ;
   END else
   BEGIN
   InitInside (FF, PPANEL) ;
   FF.Show ;
   END ;
end;

{==============================================================================================}
{=============================== Evènements de la Form ========================================}
{==============================================================================================}
procedure TFArrondi.FormShow(Sender: TObject);
begin
TableName:='ARRONDI' ;
UniqueName:= 'GAR_RANG;GAR_CODEARRONDI' ;
CodeName:='GAR_CODEARRONDI' ;
LibelleName:='GAR_LIBELLE' ;
NumCle:=2 ;
FRange:='1' ;
TypeAction:=Action ;
G_Gar.ListeParam := 'GCARRONDI';
EtudieColsListe;
HMTrad.ResizeGridColumns(G_Gar);
AffecteGrid (G_Gar, Action);
//Récupérer la devise pour avoir le nombre décimale
DEV.Code := V_PGI.DevisePivot ;
GetInfosDevise (DEV) ;
InitialiseEntete;
G_GAR.ElipsisButton:= False;
BValide := False;
G_Gar.GetCellCanvas:= DessineCell;
  inherited;
end;

procedure TFArrondi.ChargeEnreg;
Var i_ind,nbrCol : integer;
    Q : TQuery;
    Tobl : Tob;
BEGIN
inherited;
InitialiseGrille ;
For i_ind := TOBArr.Detail.Count - 1 downto 0 do TOBArr.Detail[i_ind].Free;
Q := OpenSQL('SELECT * FROM ARRONDI WHERE GAR_CODEARRONDI="'+GAR_CODEARRONDI.Text+'"' ,True,-1,'',true) ;
TOBArr.LoadDetailDB('ARRONDI','','',Q,False) ;
Ferme(Q) ;
//Trier montantSeuil dans ordre croissant
TobArr.detail.sort('GAR_VALEURSEUIL');
If TOBArr.Detail.Count<>0 Then
     InitLibelle.text := TobArr.Detail[0].GetValue('GAR_LIBELLE');
For i_ind := 1 To TOBArr.Detail.Count Do
   BEGIN
   Tobl := GetTobLigne(i_ind);
   If Tobl = NIL Then Exit;
   If i_ind = 1 Then
      BEGIN
      TObl.AddChampSup('(curseur)',False);
      TObl.putvalue('(curseur)',Curseur);
      END;
   Tobl.PutligneGrid(G_GAR,i_ind,False,False,LesColFliste);
   If Tobl.Getvalue('GAR_METHODE')<>'' Then
   G_GAR.Cells[Arr_Libe,i_ind]:= Rechdom('YYMETHODEARRONDI',Tobl.Getvalue('GAR_METHODE'),False);
   For nbrCol := 1 To G_GAR.colcount-1 Do
      FormateZoneSaisie(nbrCol,i_ind);
   END;
END;

procedure TFArrondi.FormCreate(Sender: TObject);
begin
  inherited;
iTableLigne := PrefixeToNum('GAR') ;
TOBArr:=TOB.Create('',Nil,-1);
InitLesCols();
end;

procedure TFArrondi.FormKeyDown(Sender: TObject; var Key: Word;
                                    Shift: TShiftState);
var Arow : integer;
    Cancel, Chg : Boolean;
begin
  inherited;
If(Screen.ActiveControl = G_Gar) Then
    BEGIN
    ARow := FListe.Row;
    Case Key Of
        VK_F5     : G_GARElipsisClick(Sender);
        VK_RETURN : Key:=VK_TAB ;
        VK_DELETE : BEGIN
                    If Shift=[ssCtrl] Then
                        BEGIN
                        Key := 0 ;
                        SupprimeLigne (ARow);
                        Cancel := False; Chg := False;
                        G_GarRowEnter (Sender, G_Gar.Row, Cancel, Chg);
                        END;
                    END;
        END;
    END;
end;

{==============================================================================================}
{=========================      Initialisation    =============================================}
{==============================================================================================}
procedure TFArrondi.InitLesCols;
begin
ARR_Type:= -1;
ARR_Mont:= -1;
ARR_Poid:= -1;
ARR_Cons:= -1;
end;

procedure TFArrondi.InitialiseEntete;
begin
GAR_CODEARRONDI.Text := '' ;
InitLibelle.Text := '';
NValeur.Decimals := DEV.Decimale;
NResult.Decimals := DEV.Decimale;
InitialiseGrille;
end;

procedure TFArrondi.InitialiseGrille;
begin
G_Gar.Enabled := True;
G_Gar.VidePile(True) ;
G_Gar.RowCount:= NbRowsInit;
end;

{==============================================================================================}
{========================= Evenement de l'Entete  =============================================}
{==============================================================================================}
procedure TFArrondi.GAR_CODEARRONDIExit(Sender: TObject);
Var CC : TComponent ;
begin
inherited;
If existeSQL('SELECT * FROM ARRONDI WHERE ' +
             'GAR_CODEARRONDI="'+GAR_CODEARRONDI.Text+'"') Then
   BEGIN
   HM2.execute(0,Caption,'');
   CC:=FindComponent('GAR_CODEARRONDI') ;
   FocusControle(CC) ;
   END;
inherited;
end;

{==============================================================================================}
{=============================== Actions liées au Grid ========================================}
{==============================================================================================}
procedure TFArrondi.EtudieColsListe;
Var NomCol,LesCols : String ;
    icol,ichamp    : Integer ;
BEGIN
G_Gar.ColWidths[0]:=0 ;
LesCols:=G_Gar.Titres[0] ;
LesColFliste := LesCols ;
icol:=0 ;
Repeat
    NomCol:=uppercase(Trim(ReadTokenSt(LesCols))) ;
    If NomCol<>'' Then
        BEGIN
        ichamp:=ChampToNum(NomCol) ;
        If ichamp>=0 Then
           BEGIN
           If NomCol=''                 Then ARR_Curs:=icol  Else
           If NomCol='GAR_VALEURSEUIL'  Then ARR_Mont:=icol  Else
           If NomCol='GAR_POIDSARRONDI' Then ARR_Poid:=icol  Else
           If NomCol='Méthode'          Then ARR_Libe:=icol  Else
           If NomCol='GAR_CONSTANTE'    Then ARR_Cons:=icol;
           END ;
        END ;
    Inc(icol) ;
    Until ((LesCols='') OR (NomCol='')) ;
    // Rq nomcol=méthode et nomcol='' non affecté dc en dur
end;

procedure TFArrondi.FormateZoneSaisie(ACol, ARow: Integer);
Var St,StC : String ;
BEGIN
St:=G_Gar.Cells[ACol,ARow];
StC:=St ;
If ACol=Arr_Mont Then StC:=StrF00(Valeur(St),0) Else
If ACol=Arr_Poid Then StC:=StrF00(Valeur(St),DEV.Decimale) Else
If ACol=Arr_Cons Then StC:=StrF00(Valeur(St),DEV.Decimale) ;
G_Gar.Cells[ACol,ARow]:=StC ;
end;

procedure TFArrondi.DessineCell(ACol, ARow: Integer; Canvas: TCanvas;
  AState: TGridDrawState);
Var Coord : TRect;
    ST : string;
begin
if (ACol = 1) AND (ARow > 0) then
    BEGIN
    Coord := G_Gar.CellRect (ACol, ARow);
    //Canvas.Font.Name  := 'Wingdings 3';
    Canvas.Font.Name  := 'Wingdings';
    Canvas.Font.Size  := 10;
    Canvas.Font.Color := TColor ($000000);
    Canvas.Font.Style := [fsBold];
    st := G_Gar.Cells [ACol, Arow];
    Canvas.TextOut ( (Coord.Left+Coord.Right) div 2 - Canvas.TextWidth(st) div 2,
                     (Coord.Top+ Coord.Bottom) div 2 - Canvas.TextHeight(st) div 2, st);
    END;
end;

{==============================================================================================}
{=============================== Evènements de la Grid ========================================}
{==============================================================================================}

procedure TFArrondi.G_GAREnter(Sender: TObject);
var Cancel, Chg : Boolean;
    ACol, ARow  : integer;
begin
  inherited;
Cancel := False;
Chg := False;
G_GARRowEnter (Sender, G_Gar.Row, Cancel, Chg);
ACol := G_GAR.Col ;
ARow := G_GAR.Row ;
G_GARCellEnter (Sender, ACol, ARow, Cancel);
end;

procedure TFArrondi.G_GARRowEnter(Sender: TObject; Ou: Integer;
                                      var Cancel: Boolean; Chg: Boolean);
var ARow, ACol : Integer;
begin
  inherited;
If Ou >= G_GAR.RowCount - 1 Then
   G_GAR.RowCount := G_GAR.RowCount + NbRowsPlus ;
ARow := Min (Ou, TOBArr.detail.count + 1);
If ((ARow = TOBArr.detail.count + 1) AND (Not LigneVide (ARow - 1, ACol))) Then
   CreerTOBLigne(ARow);
AfficheLigne(Arow);
If Ou > TOBArr.detail.count Then
   G_GAR.Row := TOBArr.detail.count;
end;

procedure TFArrondi.G_GARRowExit(Sender: TObject; Ou: Integer;
                                     var Cancel: Boolean; Chg: Boolean);
var ACol : integer;
begin
If csDestroying in ComponentState Then Exit ;
  inherited;
If LigneVide (Ou, ACol) Then G_GAR.Row := Min (G_GAR.Row, Ou);
//enlever le curseur précedent
G_Gar.cells[ARR_Curs,Ou]:='';
end;

procedure TFArrondi.G_GARCellEnter(Sender: TObject; var ACol,
                                       ARow: Integer; var Cancel: Boolean);
begin
If Action = taconsult Then Exit;
 inherited;
 If Not Cancel Then
   BEGIN
   If (G_GAR.col <> Arr_Mont) Then
      BEGIN
      If (G_GAR.cells[Arr_Mont, G_GAR.row]='') Then
         G_GAR.col := Arr_Mont;
      END;
   END;
G_GAR.ElipsisButton := (G_GAR.Col = Arr_libe);
end;

procedure TFArrondi.G_GARCellExit(Sender: TObject; var ACol,
                                      ARow: Integer; var Cancel: Boolean);
var ForceCol : integer;
begin
If csDestroying in ComponentState Then Exit ;
 inherited;
FormateZoneSaisie (ACol, ARow);
If Acol = Arr_Mont        Then TraiterMontant(Acol,Arow) Else
   If Acol = Arr_Poid     Then TraiterPoids(Acol,Arow)   Else
     If Acol = Arr_libe   Then TraiterMethode(Acol,Arow) Else
       If Acol = Arr_Cons Then TraiterConstante(Acol,Arow);
ForceCol := -1;
If Not LigneVide (Arow, ForceCol) then
    BEGIN
    If ForceCol <> -1 Then
        BEGIN
        ACol := ForceCol;
        G_GAR.Col := ACol;
        G_GAR.Row := ARow;
        G_GARCellEnter (Sender, ACol, ARow, Cancel);
        END;
    END;
end;

procedure TFArrondi.G_GARElipsisClick(Sender: TObject);
Var GAR_Type : THCritMaskEdit;
    Coord    : TRect;
begin
  inherited;
If G_GAR.Col = Arr_libe Then
    BEGIN
    Coord := G_GAR.CellRect (G_GAR.Col, G_GAR.Row);
    GAR_Type := THCritMaskEdit.Create (Self);
    GAR_Type.Parent := G_GAR;
    GAR_Type.Top := Coord.Top;
    GAR_Type.Left := Coord.Left;
    GAR_Type.Width := 3;
    GAR_Type.Visible := False;
    GAR_TYPE.DataType := 'YYMETHODEARRONDI';
    GetGArrondiRecherche(GAR_TYPE);
    If GAR_TYPE.Text <> '' Then
       BEGIN
       G_GAR.Cells[G_GAR.Col-4,G_GAR.Row]:= GAR_TYPE.Text;
       G_GAR.Cells[G_GAR.Col,G_GAR.Row]:=Rechdom('YYMETHODEARRONDI',GAR_TYPE.Text,False);
       END;
    GAR_TYPE.Destroy;
    END;
end;

{==============================================================================================}
{========================  Manipulation liées au Grid   =======================================}
{==============================================================================================}

procedure TFArrondi.InitialiseLigne(Arow: integer);
Var Tobl : TOB;
begin
Tobl := GetTobLigne(Arow);
If Tobl<>Nil Then Tobl.initvaleurs;
G_GAR.Rows[Arow].clear;
G_GAR.Cells [Arr_Mont, ARow]:='99999999';
Tobl.putvalue('GAR_VALEURSEUIL',Valeur(G_GAR.cells[Arr_Mont,Arow]));
end;

procedure TFArrondi.GetGArrondiRecherche(GAR_GArrondi : THCritMaskEdit);
begin
GAR_GArrondi.text := '';
LookupCombo(GAR_GArrondi) ;
end;

procedure TFArrondi.AfficheLigne(Arow: integer);
Var Tobl  : TOB;
    i_ind : integer;
begin
Tobl := GetTobLigne(Arow);
If Tobl = NIL Then Exit;
TObl.AddChampSup('(curseur)',False);
TObl.putvalue('(curseur)',Curseur);
Tobl.PutligneGrid(G_GAR,Arow,False,False,LesColFliste);
If Tobl.Getvalue('GAR_METHODE')<>'' Then
G_GAR.Cells[Arr_Libe,Arow]:= Rechdom('YYMETHODEARRONDI',Tobl.Getvalue('GAR_METHODE'),False);
For i_ind := 1 To G_GAR.colcount-1 Do
   FormateZoneSaisie(i_ind,Arow);
end;

function TFArrondi.ControlLigne(Valider: boolean): boolean;
Var i_ind, Acol  : integer;
begin
Result:= True;
i_ind := 0;
ACol  := -1;
If TobArr.Detail.Count > 0 Then
   BEGIN
   Repeat
       If LigneVide(i_ind+1,ACol) Then
          BEGIN
          If(TobArr.Detail.Count = 1)Or(Acol=-1) Then
            BEGIN
            MsgBox.Execute(4,Caption,'');
            Result := False;
            Exit;
            END
          END;
       INC(i_ind);
   Until (i_ind>=TobArr.detail.count) OR (Acol <> -1);
   END;
   If (Not Valider) AND (Acol <> -1 ) Then
   BEGIN
   G_GAR.row := i_ind;
   G_GAR.col := Acol;
   G_GAR.setfocus;
   END;
end;

function TFArrondi.ControlPentete(Valider: boolean): boolean;
begin
Result := False;
If GAR_CODEARRONDI.text = '' Then
BEGIN
   If Valider Then MsgBox.Execute(2,Caption,'');
   GAR_CODEARRONDI.setfocus;
   Exit;
END;
If InitLibelle.text = '' Then
BEGIN
   If Valider Then MsgBox.Execute(1,Caption,'');
   InitLibelle.setfocus;
   Exit;
END;
Result := True;
end;

procedure TFArrondi.SupprimeLigne(Arow: integer);
begin
If Action = TaConsult Then Exit;
If Arow < 1 Then Exit;
If (Arow > TobArr.detail.count) Then Exit;
G_GAR.cacheEdit;
G_GAR.SynEnabled := False;
G_GAR.deleteRow(Arow);
TobArr.detail[Arow-1].free;
If G_GAR.rowcount < NbRowsInit Then G_GAR.rowcount := NbRowsInit;
G_GAR.MontreEdit;
G_GAR.synEnabled := True;
end;

function TFArrondi.LigneVide(Arow: integer; var ACol: integer): boolean;
Var Suiv,MontNeg,PoidNeg : string;
begin
Suiv := Trim(G_GAR.Cells [Arr_Mont, ARow+1]);
MontNeg := Copy(G_GAR.Cells [Arr_Mont,ARow],1,1);
PoidNeg := Copy(G_GAR.Cells [Arr_Poid,ARow],1,1);
Result  := True;
If (G_GAR.Cells [Arr_Mont, ARow] <> '0')AND
   (G_GAR.Cells [Arr_Poid, Arow] <> '0')AND
   (MontNeg<>'-') AND (PoidNeg <>'-')   AND
   (Trim(G_GAR.Cells [Arr_Libe, ARow])<>'')Then
   BEGIN
   If(Trim(G_GAR.Cells [Arr_Mont,ARow])<> '') AND
     (Trim(G_GAR.Cells [Arr_Poid,ARow])<> '') AND
     (G_GAR.Cells [Arr_libe,ARow] <> 'Error')Then
     Result := False;
   END;
//Si derniere ligne vide?vide Acol<>-1 dc éviter afficher message'lg incorrecte
If ((Trim(G_GAR.Cells [Arr_Mont,ARow])= '') OR
   (G_GAR.Cells [Arr_Mont,ARow]= '0'))AND
   (Trim(G_GAR.Cells [Arr_Poid,ARow])= '') AND
   (Trim(G_GAR.Cells [Arr_Cons,ARow])= '')AND(Suiv='')Then
     ACol := Arr_Mont;
end;

procedure TFArrondi.CreerTOBLigne(Arow: integer);
begin
If Arow <> TobArr.detail.count+1 Then Exit;
TOB.Create('ARRONDI',TOBArr,Arow-1);
Initialiseligne(Arow);
end;

function TFArrondi.GetTobLigne(Arow: integer): TOB;
begin
Result := Nil;
If ((Arow<=0) OR (Arow>TobArr.detail.Count)) Then Exit;
Result:= TobArr.detail[Arow-1];
end;

procedure TFArrondi.DepileTobLigne;
Var i_ind : integer;
begin
For i_ind := TobArr.Detail.Count -1 DownTo 0 Do
 TobArr.detail[i_ind].free;
end;

{==============================================================================================}
{=====================        Manipulation des LIGNES             =============================}
{==============================================================================================}
procedure TFArrondi.TraiterMontant(Acol, Arow: integer);
Var Tobl  : TOB;
    ValMont: Extended;
begin
Tobl := GetTobLigne(Arow);
If Tobl = NIL Then Exit;
ValMont   := Valeur(G_GAR.cells[Acol,Arow]);
If (ValMont > 0) Then
   Tobl.putvalue('GAR_VALEURSEUIL',ValMont);
end;

procedure TFArrondi.TraiterPoids(Acol, Arow: integer);
Var Tobl   : TOB;
    ValPoid : Extended;
begin
Tobl := GetTobLigne(Arow);
If Tobl = NIL Then Exit;
ValPoid := Valeur(G_GAR.cells[Acol,Arow]);
If (ValPoid > 0)Then
   Tobl.PutValue('GAR_POIDSARRONDI',ValPoid);
end;

procedure TFArrondi.TraiterMethode(Acol, Arow: integer);
Var Tobl     : TOB;
    TypMethod,Libel : String;
begin
Tobl := GetTobLigne(Arow);
If Tobl = NIL Then Exit;
Libel:=Rechdom('YYMETHODEARRONDI',Tobl.Getvalue('GAR_METHODE'),False);
If G_GAR.cells[Acol,Arow]<>''Then
   BEGIN
   TypMethod:=Trim(G_GAR.cells[Acol,Arow]);
   If (TypMethod='Supérieur')OR(TypMethod='Inférieur')
      OR(TypMethod='Au plus près') Then
      BEGIN
      Tobl.PutValue('GAR_METHODE',G_GAR.cells[Acol-4,Arow]);
      END
      Else
      BEGIN
      If (TypMethod='P')OR(TypMethod='S')OR(TypMethod='I')Then
         BEGIN
         G_GAR.Cells[Acol-4, Arow] := TypMethod;
         Tobl.putvalue('GAR_METHODE',TypMethod);
         G_GAR.Cells[Acol, Arow]:= Rechdom('YYMETHODEARRONDI',TypMethod,False);
         END else
         G_GAR.Cells[Acol, Arow]:= Libel;
      END;
   END else
   BEGIN
   G_GAR.Cells[Acol, Arow]:=Libel;
   MsgBox.Execute(6,Caption,'');
   END;
end;

procedure TFArrondi.TraiterConstante(Acol, Arow: integer);
Var Tobl   : TOB;
    ValCons : Extended;
begin
Tobl := GetTobLigne(Arow);
If Tobl = NIL Then Exit;
ValCons := Valeur(G_GAR.Cells[Acol,Arow]);
{If ValCons < 0 Then
   BEGIN
   G_GAR.Cells[Acol,Arow]:= Tobl.Getvalue('GAR_CONSTANTE');
   MsgBox.Execute(5,Caption,'');
   END else }
   Tobl.PutValue('GAR_CONSTANTE',ValCons);
end;

{==============================================================================================}
{============================ Evenement lié aux Boutons =======================================}
{==============================================================================================}
procedure TFArrondi.BImprimerClick(Sender: TObject);
Var SQL : String ;
    Pages : TpageControl;
begin
//inherited;
Pages := TPageControl.Create(Application);
SQL:='SELECT * FROM ARRONDI ORDER BY GAR_CODEARRONDI,GAR_VALEURSEUIL';
LanceEtat('E','GCE','MAR',True,False,False,Pages,SQL,'',False) ;
Pages.Free;
end;

procedure TFArrondi.BDeleteClick(Sender: TObject);
Var Requete,CodeArr : String;
    Rang : Integer;
    Q : Tquery;
begin
CodeArr := GAR_CODEARRONDI.Text;
inherited;
//suppression le reste des données à partir le rang 2
Requete := 'SELECT GAR_RANG FROM ARRONDI WHERE GAR_CODEARRONDI="'+CodeArr+'"';
Q := OpenSql(Requete,True,-1,'',true);
If Not Q.Eof Then
    BEGIN
    Rang := Q.Fields[0].Asinteger;
    If Rang <> 1 Then
        ExecuteSQL('DELETE FROM ARRONDI WHERE GAR_CODEARRONDI="'+CodeArr+'"');
    END;
Ferme(Q);
end;

procedure TFArrondi.BValiderClick(Sender: TObject);
Var Col,Row : Integer;
    FinLg   : Boolean;
    Newli   : Variant;
begin
Col := -1 ; FinLg := False;
If TobArr.detail.Count > 0 Then
BEGIN
If ControlLigne(true) And ControlPentete(True) Then
    BEGIN
    For Row:=0 To TobArr.detail.count-1 do
        BEGIN
        If not LigneVide(Row+1,Col)Then
            BEGIN
            TobArr.Detail[Row].PutValue('GAR_CODEARRONDI',Trim(GAR_CODEARRONDI.text));
            TobArr.Detail[Row].PutValue('GAR_LIBELLE',Trim(InitLibelle.text));
            If(G_Gar.Cells[Arr_Cons,Row+1]<>'')and(Valeur(G_Gar.Cells[Arr_Cons,Row+1])<>0)Then
            TobArr.Detail[Row].Putvalue('GAR_CONSTANTE',Valeur(G_Gar.Cells[Arr_Cons,Row+1]))
            Else TobArr.Detail[Row].Putvalue('GAR_CONSTANTE','0');
            TobArr.Detail[Row].PutValue('GAR_RANG',Row+1);
            END Else FinLg := True; //curseur fin de ligne
        END;
    TobArr.InsertOrUpdateDB(False);
    If FinLg Then ExecuteSQL('DELETE FROM ARRONDI WHERE GAR_CODEARRONDI=" "');
    BValide := true;
    //Cas création et Bvalider;si changement lgne de liste=>éviter message affichée
    TobArr.SetAllModifie(False);
    Ta.Edit;
    NewLi:=InitLibelle.text;
    Ta.FieldByName('GAR_LIBELLE').Value := NewLi;
    STaDataChange(Sender,Ta.FieldByName('GAR_LIBELLE'));
    //Eviter le messag info 'enregister modif' si déjà actionner valider
    Ta.Refresh;
    END;
END
Else
   BEGIN
   MsgBox.Execute(7,Caption,'');
   Exit;
   END;
//inherited;
end;

procedure TFArrondi.BinsertClick(Sender: TObject);
begin
inherited;
InitLibelle.text := '';
end;

procedure TFArrondi.bDefaireClick(Sender: TObject);
begin
ChargeEnreg;
inherited;
end;

procedure TFArrondi.STaUpdateData(Sender: TObject);
begin
If Modifier = false Then
    BEGIN
    BValiderClick(Sender);
    BValide := false;
    END;
//Modifier := False;
//inherited;
Ta.cancel;
Ta.Refresh;
end;

procedure TFArrondi.STaDataChange(Sender: TObject; Field: TField);
var LgTotal,i_ind,LgSelect : Integer;
    Requete : String;
    Q : TQuery;
begin
Requete := 'SELECT COUNT(DISTINCT(GAR_CODEARRONDI)) FROM ARRONDI';
Q := OpenSql(Requete,True,-1,'',true);
If Q.Eof Then LgTotal := 1
Else LgTotal := Q.Fields[0].Asinteger;
Ferme(Q);
LgSelect:= Fliste.Row;
If CodeArrPrec <> '' Then
   BEGIN
   //cas annulation de la validation
   GAR_CODEARRONDI.text := CodeArrPrec;
   Fliste.row := LgPrec;
   CodeArrPrec := '';
   inherited;
   //forcer le curseur de revenir prec/suiv
   If Lgprec = 1 Then BFirstClick(Sender);
   If Lgprec = LgTotal Then BLastClick(Sender);
   If LgSelect < Lgprec Then
        For i_ind:=LgSelect To Lgprec-1 do  BNextClick(Sender)
   Else
        For i_ind:=LgSelect Downto Lgprec+1 do BPrevClick(Sender);
   LgPrec := 0; Exit;
   END
Else inherited;
end;

/////////       NAVIGATOR         /////////////
//////////////////////////////////////////////
procedure TFArrondi.BNextClick(Sender: TObject);
begin
BoutonsNavigator(Sender,nbnext);
//inherited;
end;

procedure TFArrondi.BPrevClick(Sender: TObject);
begin
BoutonsNavigator(Sender,nbPrior);
//inherited;
end;

procedure TFArrondi.BLastClick(Sender: TObject);
begin
BoutonsNavigator(Sender,nbLast);
//inherited;
end;

procedure TFArrondi.BFirstClick(Sender: TObject);
begin
BoutonsNavigator(Sender,nbFirst);
//inherited;
end;

procedure TFArrondi.BoutonsNavigator(Sender: TObject;Button: TNavigateBtn);
begin
If (ControlLigne(True))AND(TobArr.Detail.Count<>0) Then
    BEGIN
    If (InitLibelle.text<>TobArr.Detail[0].GetValue('GAR_LIBELLE'))
        OR TobModifier Then
        BEGIN
        If Sauvegarder Then
        If not TransacNav(DBNav.BtnClick,Button,25) Then MessageAlerte(HM.Mess[5]);
        END
        Else
        If not TransacNav(DBNav.BtnClick,Button,25) Then MessageAlerte(HM.Mess[5]);
    END
    ELse If not TransacNav(DBNav.BtnClick,Button,25) Then MessageAlerte(HM.Mess[5]);
end;

{===========================================================================}
{========      Evenement lié avec TFFICHELISTE            ==================}
{===========================================================================}
procedure TFArrondi.FListeEnter(Sender: TObject);
begin
If (ControlLigne(True))AND(TobArr.Detail.Count<>0) Then
    BEGIN
    If (InitLibelle.text<>TobArr.Detail[0].GetValue('GAR_LIBELLE'))
        OR TobModifier Then
        BEGIN
        BValide := False;
        If Not Sauvegarder Then
            BEGIN
            //cas annulation de la validation
            CodeArrPrec := GAR_CODEARRONDI.text;
            LgPrec:=Fliste.row;
            Fliste.ClearSelected;
            END;
        END;
    END;
//inherited;
end;

function TFArrondi.TobModifier: Boolean;
Var i_ind : Integer;
begin
Result := False;
For i_ind:=0 To TobArr.Detail.Count-1 Do
    BEGIN
    If TobArr.Detail[i_ind].IsFieldModified('GAR_VALEURSEUIL') OR
       TobArr.Detail[i_ind].IsFieldModified('GAR_POIDSARRONDI')OR
       TobArr.Detail[i_ind].IsFieldModified('GAR_METHODE')     OR
       TobArr.Detail[i_ind].IsFieldModified('GAR_CONSTANTE')   Then
       Result := True;
    END;
end;

function TFArrondi.Sauvegarder ;
Var Rep : Integer;
begin
Result := False;
If BValide Then BValide := False
Else
    BEGIN
    Rep:=HM.execute(0,Caption,'') ;
    Case Rep of
     mrYes : BEGIN
             BValiderClick(Self);
             BValide := False;
             Result  := True;
             Exit;
             END;
     mrNo  : BEGIN Result := True; Exit; END;
     mrCancel : Exit ;
     End ;
    END;
end;

////////////////////////////////////////////////////////////////////////////////
///************************** Test de l'arrondi *****************************///
////////////////////////////////////////////////////////////////////////////////

procedure TFArrondi.btestClick(Sender: TObject);
begin
  inherited;
TTest.Visible := btest.Down;
if btest.Down then
   begin
   NValeur.Text := StrS(0,DEV.Decimale);
   NResult.Text := StrS(0,DEV.Decimale);
   NValeur.SetFocus;
   NValeur.SelectAll;
   end;
end;

procedure TFArrondi.TTestClose(Sender: TObject);
begin
  inherited;
btest.Down := false;
end;

procedure TFArrondi.NValeurKeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
Var TobDup,TOBA : TOB;
    i_ind : integer ;
    Prix : double;
begin
  inherited;
if not isNumeric(NValeur.Text) then
   begin
   NResult.Text := StrS(0,DEV.Decimale);
   exit;
   end
   else if NValeur.Text = '-' then exit;

TOBA := nil ;
Prix := StrToFloat(NValeur.Text);
TobDup := TOB.Create('',nil,-1);
TobDup.Dupliquer(TobArr,true,true);
TobDup.Detail.Sort('GAR_VALEURSEUIL');
for i_ind := TOBArr.detail.count - 1 downto 0 do
    BEGIN
    if Prix <= TobDup.Detail[i_ind].GetValue ('GAR_VALEURSEUIL') then
        TOBA:=TobDup.Detail[i_ind] else Break;
    END;
if TOBA = nil then
   begin
   TobDup.Free;
   exit;
   end;
Prix:= CalculArrondi (TOBA,Prix) ;  //AC Focntion générique dans TarifUtil   
NResult.Text := FloatToStr(Prix) ;
TobDup.Free;
end;

end.
