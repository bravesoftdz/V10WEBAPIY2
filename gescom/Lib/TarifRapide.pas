unit TarifRapide;

// cette saisie rapide n'existe que pour les tarifs quantitatifs : GF_QUANTITATIF = 'X'

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  HTB97, Grids, Hctrls, ComCtrls, ExtCtrls, HPanel, StdCtrls, Mask, UIUtil,
  Hent1, TarifUtil,SaisUtil, HSysMenu,
{$IFNDEF CCS3}
  TarifCond,
{$ENDIF}
{$IFDEF EAGLCLIENT}
{$ELSE}
  {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}Fe_main,
{$ENDIF}
  UTOB, math,
  hmsgbox, AglInit,
  HDimension, Menus, Buttons, EntGc;

procedure EntreeTarifRapide (Action : TActionFiche; TOBArticle : TOB;
                             StCodeCatTiers, StCodeTiers, StCodeDevise : string; ttdTableTarif : T_TableTarif;
                             var TOBTarifInit : TOB; tstTarif : T_SaisieTarif);
                                          
type
  TFTarifRapide = class(TForm)
    PENTETE: THPanel;
    GF_ARTICLE: THCritMaskEdit;
    TGF_ARTICLE: THLabel;
    TGF_DEPOT: THLabel;
    GF_DATEDEBUT: THCritMaskEdit;
    TGF_DATEDEBUT: THLabel;
    GF_DATEFIN: THCritMaskEdit;
    TGF_DATEFIN: THLabel;
    GF_LIBELLE: THCritMaskEdit;
    TGF_LIBELLE: THLabel;
    TGF_LIBELLEARTICLE: THLabel;
    TGF_TARIFTIERS: THLabel;
    G_SaisRap: THGrid;
    PBOUTONS: TDock97;
    PBARREOUTIL: TToolWindow97;
    BAbandon: TToolbarButton97;
    BAide: TToolbarButton97;
    HMTrad: THSystemMenu;
    MsgBox: THMsgBox;
    PPIED: THPanel;
    GF_REMISE: THCritMaskEdit;
    TGF_REMISE: THLabel;
    GF_DEPOT: THValComboBox;
    GF_TARIFTIERS: THValComboBox;
    BValider: TToolbarButton97;
    BCondAplli: TToolbarButton97;
    procedure BCondAplliClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormCreate(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure GF_DEPOTExit(Sender: TObject);
    procedure GF_DATEDEBUTExit(Sender: TObject);
    procedure GF_DATEFINExit(Sender: TObject);
    procedure GF_DATEDEBUTElipsisClick(Sender: TObject);
    procedure GF_DATEFINElipsisClick(Sender: TObject);
    procedure GF_LIBELLEExit(Sender: TObject);
    procedure G_SaisRapCellExit(Sender: TObject; var ACol, ARow: Integer;
                                var Cancel: Boolean);
    procedure G_SaisRapRowEnter(Sender: TObject; Ou: Integer;
                                var Cancel: Boolean; Chg: Boolean);
    procedure G_SaisRapCellEnter(Sender: TObject; var ACol, ARow: Integer;
                                 var Cancel: Boolean);
    procedure G_SaisRapMouseWheelDown(Sender: TObject; Shift: TShiftState;
                                      MousePos: TPoint; var Handled: Boolean);
    procedure G_SaisRapMouseWheelUp(Sender: TObject; Shift: TShiftState;
                                    MousePos: TPoint; var Handled: Boolean);
    procedure G_SaisRapEnter(Sender: TObject);
    procedure G_SaisRapRowExit(Sender: TObject; Ou: Integer;
                               var Cancel: Boolean; Chg: Boolean);
    procedure FormKeyDown(Sender: TObject; var Key: Word;
                          Shift: TShiftState);
    procedure GF_REMISEChange(Sender: TObject);
    procedure BValiderClick(Sender: TObject);
    procedure GF_TARIFTIERSExit(Sender: TObject);
    procedure G_SaisRapExit(Sender: TObject);
    procedure BAideClick(Sender: TObject);
  private
    { Déclarations privées }
    // Objets mémoire
    LesColSaisRap : String ;
    CodeCatTiers, CodeTiers : String;
    DEV       : RDEVISE ;
    TOBTarif, TOBTarifOrig, TOBArt, TobCondition : TOB;
    // ENTETE
    function ControleEntete (Mess : Boolean) : boolean;
    function ControleLigne (Validation : boolean) : boolean;
    Procedure InitialiseEntete;
    // LIGNES
    procedure AfficheLaLigne (ARow : integer) ;
    procedure CreerTOBLigne (ARow : integer);
    Function  GetTOBLigne (ARow : integer) : TOB ;
    procedure InitialiseLigne (ARow : integer);
    function LigneVide (Arow : integer; var ACol : integer) : boolean;
    procedure SupprimeLigne (ARow : Longint);
    function TraiterBorneInf (ACol, ARow : integer) : boolean;
    procedure TraiterPrix (ACol, ARow : integer);
    procedure TraiterRemise (ACol, ARow : integer);
    // PIED
    procedure AffichePied (Arow : integer);
    // Actions liées au Grid
    procedure EtudieColsListe ;
    Procedure InitLesCols ;
    procedure FormateZoneSaisie (ACol,ARow : Longint) ;
    // Initialisations
    procedure InitialiseGrille;
    // Validations
    procedure RenseigneTOBTarif;
  public
    { Déclarations publiques }
    iTableLigne : integer;
    ttd : T_TableTarif;
    tst : T_SaisieTarif;
    Action : TActionFiche ;
  end;

var
  FTarifRapide: TFTarifRapide;

Const SR_QInf     : integer = 0;
      SR_Px       : integer = 0;
      SR_Rem      : integer = 0;


implementation

procedure EntreeTarifRapide (Action : TActionFiche; TOBArticle : TOB;
                             StCodeCatTiers, StCodeTiers, StCodeDevise : String; ttdTableTarif : T_TableTarif;
                             var TOBTarifInit : TOB; tstTarif : T_SaisieTarif);
var FF : TFTarifRapide;
    PPANEL  : THPanel ;
BEGIN
SourisSablier;
FF := TFTarifRapide.Create(Application) ;
FF.Action:=Action ;
FF.DEV.Code := StCodeDevise ;
GetInfosDevise (FF.DEV) ;

FF.ttd := ttdTableTarif;
FF.tst := tstTarif;

FF.TOBArt := TOB.Create ('ARTICLE', Nil, -1);
FF.TOBart.Dupliquer (TOBArticle, True, True);

FF.CodeCatTiers := StCodeCatTiers;
FF.CodeTiers := StCodeTiers;

FF.TOBTarifOrig := TOBTarifInit;

PPANEL := FindInsidePanel ; // permet de savoir si la forme dépend d'un PANEL
if PPANEL = Nil then        // Le PANEL est le premier ecran affiché
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
END ;

{==============================================================================================}
{======================================= Initialisations ======================================}
{==============================================================================================}
procedure TFTarifRapide.InitialiseGrille ;
BEGIN
G_SaisRap.Enabled := True;
G_SaisRap.VidePile(True) ;
G_SaisRap.RowCount:= NbRowsInit ;
END;

{==============================================================================================}
{========================= Evenement de l'Entete  =============================================}
{==============================================================================================}

procedure TFTarifRapide.GF_DEPOTExit(Sender: TObject);
BEGIN
if csDestroying in ComponentState then Exit ;  //pour eviter erreur si sortie de l'application
end;

procedure TFTarifRapide.GF_DATEDEBUTExit(Sender: TObject);
BEGIN
if csDestroying in ComponentState then Exit ;  //pour eviter erreur si sortie de l'application
end;

procedure TFTarifRapide.GF_DATEFINExit(Sender: TObject);
BEGIN
if csDestroying in ComponentState then Exit ;  //pour eviter erreur si sortie de l'application
end;

procedure TFTarifRapide.GF_LIBELLEExit(Sender: TObject);
begin
if csDestroying in ComponentState then Exit ;  //pour eviter erreur si sortie de l'application
end;

procedure TFTarifRapide.GF_TARIFTIERSExit(Sender: TObject);
begin
if csDestroying in ComponentState then Exit ;  //pour eviter erreur si sortie de l'application
end;

{==============================================================================================}
{=============================== Initialisation de la Grid ====================================}
{==============================================================================================}

procedure TFTarifRapide.EtudieColsListe ;
Var NomCol,LesCols : String ;
    icol,ichamp : integer ;
BEGIN
// Dans la liste associée, le champ GF_ARTICLE est utilisé pour bien etaler
// les colonnes dans la GRID
G_SaisRap.ColWidths[0]:=0 ;
LesCols:=G_SaisRap.Titres[0] ; LesColSaisRap := LesCols ; icol:=0 ;
Repeat
    NomCol:=uppercase(Trim(ReadTokenSt(LesCols))) ;
    if NomCol<>'' then
        BEGIN
        ichamp:=ChampToNum(NomCol) ;
        if ichamp>=0 then
           BEGIN
           if NomCol='GF_BORNEINF'     then SR_QInf:=icol else
           if NomCol='GF_PRIXUNITAIRE' then SR_Px:=icol else
           if NomCol='GF_CALCULREMISE' then SR_Rem:=icol else
           END ;
        END ;
    Inc(icol) ;
    Until ((LesCols='') or (NomCol='')) ;
END ;

Procedure TFTarifRapide.InitLesCols ;
BEGIN
SR_QInf := -1; SR_Px := -1; SR_Rem := -1;
END ;

{==============================================================================================}
{=============================== Actions liées au Grid ========================================}
{==============================================================================================}

procedure TFTarifRapide.G_SaisRapCellEnter(Sender: TObject;
                                                        var ACol, ARow: Integer;
                                                        var Cancel: Boolean);
begin
if Action=taConsult then Exit ;
if Not Cancel then
    BEGIN
    if (G_SaisRap.Col <> SR_Qinf) then
        BEGIN
        if (G_SaisRap.Cells [SR_Qinf,G_SaisRap.Row] = '') then G_SaisRap.Col := SR_Qinf;
        end;
    END;
end;

procedure TFTarifRapide.G_SaisRapEnter(Sender: TObject);
var Cancel, Chg : Boolean;
    ACol, ARow : integer;
begin
Cancel := False; Chg := False;
G_SaisRapRowEnter (Sender, G_SaisRap.Row, Cancel, Chg);
ACol := G_SaisRap.Col ; ARow := G_SaisRap.Row ;
G_SaisRapCellEnter (Sender, ACol, ARow, Cancel);
G_SaisRap.MontreEdit;
end;

procedure TFTarifRapide.G_SaisRapRowExit(Sender: TObject;
                                                      Ou: Integer; var Cancel: Boolean;
                                                      Chg: Boolean);
var ACol : integer;
begin
if csDestroying in ComponentState then Exit ;  //pour eviter erreur si sortie de l'application
if G_SaisRap.RowCount > G_SaisRap.Row + 1 then
    begin
        if LigneVide (G_SaisRap.Row + 1, Acol) then
        begin
        SupprimeLigne (G_SaisRap.Row + 1);
        end;
    end;
if LigneVide (Ou, ACol) Then G_SaisRap.Row := Min (G_SaisRap.Row, Ou);
end;

procedure TFTarifRapide.G_SaisRapMouseWheelDown(Sender: TObject;
                                                Shift: TShiftState;
                                                MousePos: TPoint;
                                                var Handled: Boolean);
begin
Handled := True;
end;

procedure TFTarifRapide.G_SaisRapMouseWheelUp(Sender: TObject;
                                              Shift: TShiftState; MousePos: TPoint;
                                              var Handled: Boolean);
begin
Handled := True;
end;

{==============================================================================================}
{========================= Manipulation de l'entête ===========================================}
{==============================================================================================}

function TFTarifRapide.ControleEntete (Mess : Boolean) : Boolean ;
var tdtDateDebut, tdtDateFin : TDateTime;
BEGIN
Result := False;
if GF_DATEDEBUT.Text = '' then
    BEGIN
    if Mess then MsgBox.Execute (5,Caption,'') ;
    GF_DATEDEBUT.SetFocus;
    exit;
    end;
if GF_DATEFIN.Text = '' then
    BEGIN
    if Mess then MsgBox.Execute (6,Caption,'') ;
    GF_DATEDEBUT.SetFocus;
    exit;
    end;

tdtDateDebut := strtodate (GF_DATEDEBUT.Text);
tdtDateFin := strtodate (GF_DATEFIN.Text);
if (tdtDateDebut > tdtDateFin) then
    BEGIN
    MsgBox.Execute (1,Caption,'') ;
    GF_DATEDEBUT.SetFocus;
    exit;
    end;

if GF_LIBELLE.Text = '' then
    BEGIN
    if Mess then MsgBox.Execute (0,Caption,'');
    GF_LIBELLE.SetFocus;
    exit;
    end;
    
if ttd = ttdArtQCa then
    begin
    if GF_TARIFTIERS.Text = '' then
        BEGIN
        if Mess then MsgBox.Execute (4,Caption,'') ;
        GF_TARIFTIERS.SetFocus;
        exit;
        end;
    end;
Result := True;
end;

function TFTarifRapide.ControleLigne (Validation : boolean) : boolean;
var Index, ACol : integer;
begin
Result := True;
Index := 0;
ACol := -1;

if not LigneVide (G_SaisRap.Row, ACol) then
    begin
    if G_SaisRap.Col = SR_QInf then Result := TraiterBorneInf (G_SaisRap.Col, G_SaisRap.Row)
    else if G_SaisRap.Col = SR_Px then TraiterPrix (G_SaisRap.Col, G_SaisRap.Row)
    else TraiterRemise (G_SaisRap.Col, G_SaisRap.Row);
    end;

if not Result then exit;

if TobTarif.Detail.Count > 0 then
    begin
    repeat
        if ligneVide (Index + 1, ACol) then
            begin
            if (Validation) or (Index + 1 <> TobTarif.Detail.Count) then
                begin
                TobTarif.Detail [Index].Free;
                G_SaisRap.DeleteRow (Index + 1);
                end;
            end else
            begin
            if ACol <> -1 then Result := False;
            end;
            Index := Index + 1;
    until (Index >= TobTarif.Detail.Count) or (ACol <> -1);
    end;
if (not Validation) and (ACol <> -1) then
    begin
    G_SaisRap.Row := Index;
    G_SaisRap.Col := ACol;
    G_SaisRap.SetFocus;
    end;
end;

Procedure TFTarifRapide.InitialiseEntete ; // a modifier pour Tiers
BEGIN
if ttd <> ttdArtQCa then
    begin
    TGF_TARIFTIERS.visible := False;
    GF_TARIFTIERS.Visible := False;
    TGF_TARIFTIERS.Enabled := True;
    GF_TARIFTIERS.Enabled := True;
    end;
GF_ARTICLE.Text := TOBart.GetValue ('GA_CODEARTICLE');
TGF_LIBELLEARTICLE.Caption := TOBArt.GetValue ('GA_LIBELLE');
InitialiseGrille;
GF_DATEDEBUT.Text := DateToStr (V_PGI.DateEntree);
END;

{==============================================================================================}
{=============================== Manipulation des lignes ======================================}
{==============================================================================================}

procedure TFTarifRapide.AfficheLaLigne (ARow : integer) ;
Var TOBL : TOB ;
    i_ind : integer ;
BEGIN
TOBL:=GetTOBLigne(ARow) ; if TOBL = Nil then exit;
TOBL.PutLigneGrid(G_SaisRap,ARow,False,False,LesColSaisRap) ;
for i_ind := 1 to G_SaisRap.ColCount-1 do
    BEGIN
    FormateZoneSaisie(i_ind, ARow) ;
    END;
END ;

Procedure TFTarifRapide.CreerTOBLigne (ARow : integer);
BEGIN
if ARow <> TOBTarif.Detail.Count + 1 then exit;
TOB.Create ('TARIF', TOBTarif, ARow-1) ;
InitialiseLigne (ARow) ;
END;

procedure TFTarifRapide.FormateZoneSaisie (ACol,ARow : Longint) ;
Var St,StC : String ;
BEGIN
St:=G_SaisRap.Cells[ACol,ARow];
StC:=St ;
if ACol=SR_Px then StC:=StrF00(Valeur(St),DEV.Decimale) else
    if ACol=SR_QInf then StC:=StrF00(Valeur(St),0);
G_SaisRap.Cells[ACol,ARow]:=StC ;
END ;

Function TFTarifRapide.GetTOBLigne (ARow : integer) : TOB ;
BEGIN
Result:=Nil ;
if ((ARow<=0) or (ARow>TOBTarif.Detail.Count)) then Exit ;
Result:=TOBTarif.Detail[ARow-1] ;
END ;

procedure TFTarifRapide.InitialiseLigne (ARow : integer) ;
Var TOBL : TOB ;
BEGIN
TOBL:=GetTOBLigne(ARow) ;
if TOBL<>Nil then TOBL.InitValeurs ;
G_SaisRap.Rows[Arow].Clear;
END;

Function TFTarifRapide.LigneVide (ARow : integer; var Acol : integer) : Boolean;
BEGIN
Result := True;
if (G_SaisRap.Cells [SR_QInf, ARow] <> '') or (G_SaisRap.Cells [SR_Px, Arow] <> '') or
   (G_SaisRap.Cells [SR_Rem, ARow] <> '') then
   BEGIN
   Result := False;
   if (G_SaisRap.Cells [SR_QInf, ARow] = '') or (G_SaisRap.Cells [SR_QInf, ARow] = '0') then
        ACol := SR_QInf;
   END;
END;

procedure TFTarifRapide.SupprimeLigne (ARow : Longint) ;
BEGIN
if Action=taConsult then Exit ;
if ARow < 1 then Exit ;
if (ARow > TOBTarif.Detail.Count) then Exit;

G_SaisRap.CacheEdit; G_SaisRap.SynEnabled := False;
G_SaisRap.DeleteRow (ARow);
TOBTarif.Detail[ARow-1].Free;
if G_SaisRap.RowCount < NbRowsInit then G_SaisRap.RowCount := NbRowsInit;
G_SaisRap.MontreEdit; G_SaisRap.SynEnabled := True;
END;

function TFTarifRapide.TraiterBorneInf (ACol, ARow : integer) : boolean;
var TOBL : TOB;
    f_QteInf : Extended;
BEGIN
Result := True;
if G_SaisRap.Cells [Acol, Arow] = '' then
    begin
    G_SaisRap.Col := Acol;
    G_SaisRap.SetFocus;
    Result := False;
    end else
    begin
    TOBL := GetTOBLigne (ARow);
    f_QteInf := Valeur (G_SaisRap.Cells [ACol, ARow]);
    if (f_QteInf > 9999999) or (f_QteInf <= 0) then
        BEGIN
        MsgBox.Execute (3,Caption,'') ;
        G_SaisRap.Cells [ACol, ARow] := inttostr (TOBL.GetValue ('GF_BORNEINF'));
        G_SaisRap.SetFocus;
        Result := False;
        END else TOBL.PutValue ('GF_BORNEINF', f_QteInf);
    end;
END;

procedure TFTarifRapide.TraiterPrix (ACol, ARow : integer);
var TOBL : TOB;
BEGIN
TOBL := GetTOBLigne (ARow);
TOBL.PutValue ('GF_PRIXUNITAIRE', Valeur (G_SaisRap.Cells [ACol, ARow]));
END;

procedure TFTarifRapide.TraiterRemise(ACol, ARow : integer);
var TOBL : TOB;
    St : string;
BEGIN
TOBL := GetTOBLigne (ARow);
St := G_SaisRap.Cells [ACol, ARow];
TOBL.PutValue ('GF_CALCULREMISE', St);
TOBL.PutValue ('GF_REMISE', RemiseResultante (St));
AffichePied (ARow);
END;

{==============================================================================================}
{============================= Manipulation du pied ===========================================}
{==============================================================================================}

Procedure TFTarifRapide.AffichePied (ARow : Longint) ;
var TOBL : TOB;
    FF   : TForm;
    i_ind, Num : integer;
    CHPS : TControl;
    Nom : string;
BEGIN
TOBL:=GetTOBLigne(ARow);
if TOBL=Nil then Exit ;

FF := TForm (PPIED.Owner);
for i_ind := 0 to FF.ComponentCount - 1 do
    BEGIN
    CHPS := TControl (FF.Components[i_ind]);
    if (CHPS is THCritMaskEdit) AND (CHPS.Parent.Name = 'PPIED') then
        BEGIN
        Nom := Uppercase (CHPS.Name);
        Num := TOBL.GetNumChamp (Nom);
        if Num > 0 then
            BEGIN
            if THCritMaskEdit(CHPS).OpeType = otReel then
                THCritMaskEdit(CHPS).Text := FloatToStr (TOBL.Valeurs[Num])
                else THCritMaskEdit(CHPS).Text := TOBL.Valeurs[Num] ;
            END;
        END;
    if (CHPS is THValComboBox) AND (CHPS.Parent.Name = 'PPIED') then
        BEGIN
        Nom := Uppercase (CHPS.Name);
        Num := TOBL.GetNumChamp (Nom);
        if Num > 0 then
            BEGIN
            if ((THVaLComboBox(CHPS).Vide) and (TOBL.Valeurs[Num]='')) then
                THVaLComboBox(CHPS).ItemIndex := 0
                else THVaLComboBox(CHPS).Value := TOBL.Valeurs[Num] ;
            END;
        END;
    END;
END;

{==============================================================================================}
{=============================== Evenements du pied ===========================================}
{==============================================================================================}

procedure TFTarifRapide.GF_REMISEChange(Sender: TObject);
var St : string;
begin
St := GF_REMISE.Text;
St:=StrF00(Valeur(St),ADecimP);
GF_REMISE.Text := St;
end;

{==============================================================================================}
{=============================== Evènements de la Form ========================================}
{==============================================================================================}

procedure TFTarifRapide.BValiderClick(Sender: TObject);
begin
if (ControleEntete (True)) and (ControleLigne (True)) then RenseigneTOBTarif
else ModalResult := 0;
end;

procedure TFTarifRapide.FormClose(Sender: TObject;
  var Action: TCloseAction);
BEGIN
G_SaisRap.VidePile(True) ;
TOBTarif.Free ; TOBTarif:=Nil ;
TOBArt.Free ; TOBArt:=Nil ;
TobCondition.Free; TobCondition := Nil;
if IsInside(Self) then Action:=caFree ;
end;

procedure TFTarifRapide.FormCreate(Sender: TObject);
BEGIN
G_SaisRap.RowCount := NbRowsInit ;
iTableLigne := PrefixeToNum('GF') ;
TOBTarif := TOB.Create ('TARIF', Nil, -1) ;
InitLesCols;
TobCondition := TOB.Create ('TARIF', nil, -1) ;
end;

procedure TFTarifRapide.FormShow(Sender: TObject);
BEGIN
G_SaisRap.ListeParam := 'GCTARIFSR';
EtudieColsListe ;
HMTrad.ResizeGridColumns (G_SaisRap) ;
AffecteGrid (G_SaisRap, Action) ;
InitialiseEntete ;
{$IFDEF CCS3}
BCondAplli.Visible:=False ;
{$ENDIF}
if not VH_GC.GCMultiDepots then
   begin    //mcd 25/08/03
   TGF_DEPOT.enabled := false;
   GF_DEPOT.enabled := false;
   TGF_DEPOT.Visible := false;
   GF_DEPOT.Visible := false;
   end;
end;

procedure TFTarifRapide.GF_DATEDEBUTElipsisClick(Sender: TObject);
var DATE : THCritMaskEdit;
BEGIN
DATE := THCritMaskEdit.Create (Self);
DATE.Parent := GF_DATEDEBUT.Parent;
DATE.Top := GF_DATEDEBUT.Top;
DATE.left := GF_DATEDEBUT.Left;
DATE.Width := 3; DATE.Visible := False;
DATE.OpeType:=otDate;
GetDateRecherche (TForm(DATE.Owner), DATE) ;
if DATE.Text <> '' then GF_DATEDEBUT.Text := DATE.Text;
DATE.Destroy;
end;

procedure TFTarifRapide.GF_DATEFINElipsisClick(Sender: TObject);
var DATE : THCritMaskEdit;
BEGIN
DATE := THCritMaskEdit.Create (Self);
DATE.Parent := GF_DATEFIN.Parent;
DATE.Top := GF_DATEFIN.Top;
DATE.left := GF_DATEFIN.Left;
DATE.Width := 3; DATE.Visible := False;
DATE.OpeType:=otDate;
GetDateRecherche (TForm(DATE.Owner), DATE) ;
if DATE.Text <> '' then GF_DATEFIN.Text := DATE.Text;
DATE.Destroy;
end;

procedure TFTarifRapide.G_SaisRapCellExit(Sender: TObject; var ACol,
                                          ARow: Integer; var Cancel: Boolean);
var ForceCol : integer;
BEGIN
if csDestroying in ComponentState then Exit ;  //pour eviter erreur si sortie de l'application
FormateZoneSaisie (ACol, ARow);
if ACol = SR_QInf then TraiterBorneInf (ACol, ARow) else
    if ACol = SR_Px then TraiterPrix (ACol, ARow) else
        if ACol = SR_Rem then TraiterRemise (ACol, ARow);
ForceCol := -1;
if not LigneVide (Arow, ForceCol) then
    begin
    if ForceCol <> -1 then
        begin
        ACol := ForceCol;
        G_SaisRap.Col := ACol;
        G_SaisRap.Row := ARow;
        G_SaisRapCellEnter (Sender, ACol, ARow, Cancel);
        end;
    end;
    
END;

procedure TFTarifRapide.G_SaisRapExit(Sender: TObject);
begin
ControleLigne (False);
end;

procedure TFTarifRapide.G_SaisRapRowEnter(Sender: TObject; Ou: Integer;
                                                       var Cancel: Boolean; Chg: Boolean);
var ARow, ACol : Integer;
BEGIN
if Ou >= G_SaisRap.RowCount - 1 then G_SaisRap.RowCount := G_SaisRap.RowCount + NbRowsPlus ;
ARow := Min (Ou, TOBTarif.detail.count + 1);
if (ARow = TOBTarif.detail.count + 1) AND (not LigneVide (ARow - 1, ACol)) then
    BEGIN
    CreerTOBligne (ARow);
    END;
AfficheLaLigne (Arow);
if Ou > TOBTarif.detail.count then
    BEGIN
    G_SaisRap.Row := TOBTarif.detail.count;
    END;
AffichePied (ARow);
END;

{==============================================================================================}
{================================= Validation =================================================}
{==============================================================================================}

procedure TFTarifRapide.RenseigneTOBTarif;
var i_nb_ligne, Index, ACol : integer;
BEGIN
ACol := -1;
for Index := 0 to TOBtarif.Detail.count - 1 do
    begin
    if not lignevide (Index + 1, ACol) then
        begin
        if ACol <> SR_QInf then
            begin
            TobTarif.Detail[Index].PutValue ('GF_BORNEINF', Valeur (G_SaisRap.Cells [SR_QInf, Index + 1]));
            i_nb_ligne := TOBTarifOrig.Detail.count;
            if (i_nb_ligne > 0) and (TobTarifOrig.Detail [i_nb_ligne - 1].GetValue ('GF_LIBELLE') = '') then
                begin
                i_nb_ligne := i_nb_ligne - 1;
                end else
                begin
                TOB.create ('TARIF', TOBTarifOrig, i_nb_ligne);
                end;

            TOBTarifOrig.Detail [i_nb_ligne].dupliquer (TOBTarif.Detail [Index], true, true);
            // champs de l'entete communs à toutes les saisies
            TOBTarifOrig.Detail [i_nb_ligne].PutValue ('GF_DATEDEBUT', strtodate (GF_DATEDEBUT.Text));
            TOBTarifOrig.Detail [i_nb_ligne].PutValue ('GF_DATEFIN', strtodate (GF_DATEFIN.Text));
            TOBTarifOrig.Detail [i_nb_ligne].PutValue ('GF_DEPOT', GF_DEPOT.Value);
            TOBTarifOrig.Detail [i_nb_ligne].PutValue ('GF_LIBELLE', GF_LIBELLE.Text);
            // champs par défaut communs à toutes les saisies
            TOBTarifOrig.Detail [i_nb_ligne].PutValue ('GF_MODECREATION', 'MAN');
            TOBTarifOrig.Detail [i_nb_ligne].PutValue ('GF_CASCADEREMISE', 'MIE');
            TOBTarifOrig.Detail [i_nb_ligne].PutValue ('GF_DEVISE', DEV.Code);
            TOBTarifOrig.Detail [i_nb_ligne].PutValue ('GF_QUALIFPRIX', 'GRP');
            TOBTarifOrig.Detail [i_nb_ligne].PutValue ('GF_FERME', '-');
            TOBTarifOrig.Detail [i_nb_ligne].PutValue ('GF_SOCIETE', V_PGI.CodeSociete) ;
            TOBTarifOrig.Detail [i_nb_ligne].PutValue ('GF_TARIF', 0);
            TOBTarifOrig.Detail [i_nb_ligne].PutValue ('GF_QUANTITATIF', 'X');
            TOBTarifOrig.Detail [i_nb_ligne].PutValue ('GF_ARTICLE', TOBArt.GetValue ('GA_ARTICLE'));

            if tst = tstArticle then
                begin
                if ttd = ttdArtQca then
                    begin
                    TOBTarifOrig.Detail [i_nb_ligne].PutValue ('GF_TARIFTIERS', GF_TARIFTIERS.Value);
                    end;
                end else
                begin
                if tst = tstCliArt then
                    begin
                    TOBTarifOrig.Detail [i_nb_ligne].PutValue ('GF_TIERS', CodeTiers);
                    end else
                    begin // tst = tstTiers
                    TOBTarifOrig.Detail [i_nb_ligne].PutValue ('GF_TARIFTIERS', CodeCatTiers);
                    end;
                end;

            // champs de la grid communs à toutes les saisies
            if Index < TOBtarif.Detail.count - 1 then
                begin
                if TOBTarif.Detail [Index + 1].GetValue ('GF_BORNEINF') - 1 >
                    TOBTarif.Detail [Index].GetValue ('GF_BORNEINF') then
                        TOBTarifOrig.Detail [i_nb_ligne].PutValue ('GF_BORNESUP',
                                                TOBTarif.Detail [Index + 1].GetValue ('GF_BORNEINF') - 1)
                    else TOBTarifOrig.Detail [i_nb_ligne].PutValue ('GF_BORNESUP',
                                TOBTarif.Detail [Index].GetValue ('GF_BORNEINF') + 1);
                end else
                begin
                TOBTarifOrig.Detail [i_nb_ligne].PutValue ('GF_BORNESUP', 999999);
                end;
            TOBTarifOrig.Detail [i_nb_ligne].PutValue ('GF_PRIXUNITAIRE',
                                                       TOBTarif.Detail [Index].GetValue ('GF_PRIXUNITAIRE'));
            TOBTarifOrig.Detail [i_nb_ligne].PutValue ('GF_REMISE',
                                                       RemiseResultante (
                                                       TOBTarif.Detail [Index].GetValue ('GF_CALCULREMISE')));
            TOBTarifOrig.Detail [i_nb_ligne].PutValue ('GF_CALCULREMISE',
                                                       TOBTarif.Detail [Index].GetValue ('GF_CALCULREMISE'));

            TOBTarifOrig.Detail [i_nb_ligne].PutValue ('GF_CONDAPPLIC',
                                                       TobCondition.GetValue ('GF_CONDAPPLIC'));
            end;
        end;
    end;
END;

{$R *.DFM}

procedure TFTarifRapide.FormKeyDown(Sender: TObject;
                                                 var Key: Word; Shift: TShiftState);
var Arow : integer;
    Cancel, Chg : Boolean;
begin
if(Screen.ActiveControl = G_SaisRap) then
    BEGIN
    ARow := G_SaisRap.Row;
    Case Key of
        VK_RETURN : Key:=VK_TAB ;
        VK_DELETE : BEGIN
                    if Shift=[ssCtrl] then
                        BEGIN
                        Key := 0 ;
                        SupprimeLigne (ARow);
                        Cancel := False; Chg := False;
                        G_SaisRapRowEnter (Sender, G_SaisRap.Row, Cancel, Chg);
                        end;
                    END;
        END;
    END;
end;

procedure TFTarifRapide.BCondAplliClick(Sender: TObject);
begin
{$IFNDEF CCS3}
EntreeTarifCond (Action, TobCondition);
{$ENDIF}
end;

procedure TFTarifRapide.BAideClick(Sender: TObject);
begin
CallHelpTopic(Self);
end;

END.
