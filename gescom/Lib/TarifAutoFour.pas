unit TarifAutoFour;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  HTB97, Grids, Hctrls, ComCtrls, ExtCtrls, HPanel, StdCtrls, Mask, UIUtil,
  Hent1, TarifUtil, SaisUtil, HSysMenu, UTOB, math, hmsgbox,ParamSoc,
  AglInit,UtilArticle,
{$IFDEF EAGLCLIENT}
    MaineAgl,
{$ELSE} // EAGLCLIENT
  Fe_main, {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ENDIF} // EAGLCLIENT
  HDimension, Menus, Buttons, Hqry, ImgList, HImgList, TntStdCtrls,
  TntGrids, TntExtCtrls;

procedure EntreeTarifAutoFour (Action : TActionFiche; StParFou : string);
function GetValChamp (Table, Champ, Cle : string) : string;
procedure LoadLesTobParFou (StParFou : string; TobEntete, TobLigne : TOB);


type

  TSChamp = class
    nom : string;
    TypeD : string;
  end;

  TSStruct = class
     Nom     : string;
     Libelle : string;
     ColType : char;
     ColAlign : TAlignment;
     ColWidth : integer;
     ColEditable : boolean;
     ColFormat : string;
     ColLength : integer;
  end;

  TParamImportTarFou = class(TForm)
    PBOUTON: TDock97;
    ToolWindow971: TToolWindow97;
    BValider: TToolbarButton97;
    BAbandon: TToolbarButton97;
    BAide: TToolbarButton97;
    BNouveau: TToolbarButton97;
    bSupprimer: TToolbarButton97;
    TGRF_PARFOU: THLabel;
    GRF_PARFOU: THCritMaskEdit;
    TGRF_TIERS: THLabel;
    GRF_TIERS: THCritMaskEdit;
    PageControl1: TPageControl;
    OpenDialogButton: TOpenDialog;
    TabSheet2: TTabSheet;
    TGRF_FICHIER: THLabel;
    GRF_FICHIER: THCritMaskEdit;
    Panel1: TPanel;
    PCorps: THPanel;
    GRF_TYPEENREG: TCheckBox;
    TGRF_LONGENREG: THLabel;
    GRF_LONGENREG: THNumEdit;
    TGRF_SEPARATEUR: THLabel;
    GRF_SEPARATEUR: THValComboBox;
    TGRF_SEPTEXTE: THLabel;
    GRF_SEPTEXTE: THCritMaskEdit;
    TabSheet1: TTabSheet;
    Panel2: TPanel;
    PPARTICLE: TPanel;
    GRF_TENUESTOCK: TCheckBox;
    TGRF_FAMILLETAXE1: THLabel;
    GRF_FAMILLETAXE1: THValComboBox;
    GRF_REMISEPIED: TCheckBox;
    TGRF_COMPTAARTICLE: THLabel;
    GRF_COMPTAARTICLE: THValComboBox;
    GRF_ESCOMPTABLE: TCheckBox;
    Panel4: TPanel;
    TGRF_PREFIXE: THLabel;
    GRF_PREFIXE: THCritMaskEdit;
    TGRF_SUFFIXE: THLabel;
    GRF_SUFFIXE: THCritMaskEdit;
    TGRF_PROFILARTICLE: THLabel;
    GRF_PROFILARTICLE: THCritMaskEdit;
    GRF_FOURPRINC: TCheckBox;
    HMTrad: THSystemMenu;
    MsgBox: THMsgBox;
    LGRF_TIERS: THLabel;
    TGRF_TYPEARTICLE: THLabel;
    GRF_TYPEARTICLE: THValComboBox;
    CheckProfil: TCheckBox;
    TTFICHIERASSOC: TToolWindow97;
    GFL_FICHIER: THCritMaskEdit;
    TNomFic: THLabel;
    BvisionFic: TToolbarButton97;
    ODFichier: TOpenDialog;
    IList: THImageList;
    PRESTATION: THCritMaskEdit;
    GRF_PRESTATION: THCritMaskEdit;
    TPRESTATION: TLabel;
    TGFL_CHAMP: TLabel;
    PGauche: TPanel;
    PCentre: TPanel;
    BFlecheDroite: TToolbarButton97;
    BFlecheGauche: TToolbarButton97;
    Panel6: TPanel;
    GChamp: THGrid;
    GListeChamps: THGrid;
    GRF_MULTIFOU: THCheckbox;
    procedure GRF_FICHIERChange(Sender: TObject);
    procedure GRF_FICHIERElipsisClick(Sender: TObject);
    procedure GRF_FICHIERExit(Sender: TObject);
    procedure GRF_TIERSExit(Sender: TObject);
    procedure GRF_PARFOUExit(Sender: TObject);
    procedure GRF_TYPEENREGClick(Sender: TObject);
    procedure GRF_LONGENREGChange(Sender: TObject);
    procedure GRF_SEPARATEURChange(Sender: TObject);
    procedure GRF_SEPARATEURExit(Sender: TObject);
    procedure GRF_TYPEARTICLEExit(Sender: TObject);
    procedure GRF_FAMILLETAXE1Exit(Sender: TObject);
    procedure GRF_COMPTAARTICLEExit(Sender: TObject);
    procedure GRF_PREFIXEExit(Sender: TObject);
    procedure GRF_SUFFIXEExit(Sender: TObject);
    procedure GRF_SEPTEXTEExit(Sender: TObject);
//    procedure TLChampColumnClick(Sender: TObject; Column: TListColumn);
//    procedure TLChampCompare(Sender: TObject; Item1, Item2: TListItem;
//      Data: Integer; var Compare: Integer);
    procedure GChampCellEnter(Sender: TObject; var ACol, ARow: Integer;
      var Cancel: Boolean);
    procedure GChampCellExit(Sender: TObject; var ACol, ARow: Integer;
      var Cancel: Boolean);
    procedure GChampEnter(Sender: TObject);
    procedure GChampRowEnter(Sender: TObject; Ou: Integer;
      var Cancel: Boolean; Chg: Boolean);
    procedure GChampRowExit(Sender: TObject; Ou: Integer;
      var Cancel: Boolean; Chg: Boolean);
    procedure BFlecheDroiteClick(Sender: TObject);
    procedure BFlecheGaucheClick(Sender: TObject);
    procedure bSupprimerClick(Sender: TObject);
    procedure BNouveauClick(Sender: TObject);
    procedure BValiderClick(Sender: TObject);
    procedure BAideClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
    procedure FormCreate(Sender: TObject);
    procedure FormKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure GRF_PROFILARTICLEExit(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure FormResize(Sender: TObject);
    procedure CheckProfilClick(Sender: TObject);
    procedure BvisionFicClick(Sender: TObject);
    procedure TTFICHIERASSOCClose(Sender: TObject);
    procedure GFL_FICHIERExit(Sender: TObject);
    procedure GFL_FICHIERElipsisClick(Sender: TObject);
    procedure GRF_TYPEARTICLEChange(Sender: TObject);
    procedure PRESTATIONChange(Sender: TObject);
    procedure GRF_PARFOUEnter(Sender: TObject);
    procedure GChampKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure GRF_MULTIFOUClick(Sender: TObject);
  private
  	LastPar : string;
    NomList, FNomTable, FLien, FSortBy,  FTitre,stColListe : string;
    FLargeur, FAlignement, FParams, title, NC, FPerso: string;
    OkTri, OkNumCol : boolean;
    { Déclarations privées }
    stCellCur : string;
    LStructure,LSChamp : Tlist;
    LesColChamp : string ;
    iTableLigne, ColumnToSort,NBCols : integer;
    bModif : boolean;
    iWidthsLongueur, iWidthsDebut : integer;
    ListeChampCat, ListeChampArt  : HTStringList;
    ListeLibChampArtC, ListeLibChampArt,ListTypeChampArtC,ListeChampArtC, ListeLibChampCat : HTStringList;

    //FV1 : 17/12/2015 - FS#1838 - En import catalogue, ajouter la récupération du code barre de l'article.
    ListeChampCab, ListeLibChampCab,ListTypeChampCab : HTStringList;
    //
    // Modif BTP
    ListTypeChampCat,ListTypeChampArt : HTStringList;
    TOBListeChamps : TOB;
    // --
    TobLigDelete : TOB;
    SAF_SELECT      : integer;
    SAF_Champ       : integer ;
    SAF_Longueur    : integer ;
    SAF_Offset      : integer ;
    SAF_Libelle     : integer ;
    // Ajout BTP
    SAF_SepDecimale : integer;
    SAF_NbrDecimale : integer;
    SAF_FichierAssocie : integer;
    // --
    // Gestion des Données
    procedure ChargeParametrage;
    function  ControleModif : boolean;
    procedure DeleteParFouLig (ARow : integer);
    // Initialisations
    Procedure InitialiseEntete;
    procedure InitialiseGrilles;
    procedure InitialiseLigne (ARow : integer);
    procedure InitialiseListBox;
    // Initialisation de la Grid
    procedure EtudieColsListe ;
    Procedure InitLesCols;
    // Manipulation de l'entete
    procedure AffecteEntete;
    procedure AffecteLibelleLignes;
    procedure AffecteLibelleTiers;
    function  ControleEntete (Mess : boolean) : boolean ;
    function  ControleLigne (Validation : boolean) : boolean ;
    // Manipulation des lignes
    procedure AfficheLaLigne (ARow : integer) ;
    procedure CreerTOBLigne (ARow : integer);
    Function  GetTOBLigne (ARow : integer) : TOB ;
    procedure InsereListeChamp (stChamp, stLib : string);
    function  LigneVide (Arow : integer; var ACol : integer) : boolean;
    function  RecupereLibelleChamp (stChamp : string) : string;
    procedure RenseigneLigneGrid (StChamp,Libelle : string; var IndiceRow : integer);
    procedure SupprimeLigneParfou (ARow : Longint) ;
    procedure TraiterChamp (ACol, ARow, Provenance : integer);
    procedure TraiterNombre (ACol, ARow : integer; stCol : string);
    // Validation
    procedure EnregistreTarifAutoFour;
    procedure RenseigneDetail;
    procedure RenseigneEntete;
    procedure ValideTarifAutoFour;
    function ControleProfileArticle: boolean;
    procedure TraiterSepDecimale(ARow: integer);
    procedure AjouteLSChamp(champ, TypeD: string);
    function ChampSuivantOk(var Acol : integer ; var Arow: integer): boolean;
    function TypeChamps(CodeChamp: string): string;
    function ChampOk(Acol, Arow: integer): boolean;
    Procedure PostDrawCell ( ACol,ARow : Longint; Canvas : TCanvas ; AState: TGridDrawState) ;
    function FamilleMereRenseigne(Niveau: integer): boolean;
    procedure AddLesChampsSup(TOBL: TOB);
    procedure AfficheGridListChamps;
    procedure AfficheLigneListChamps(Ligne : integer; TOBL: TOB);
    procedure DefiniGChamp;
    procedure AddChampsSupLig(TOBL: TOB);
    procedure AffecteLibelleLigne(TOBL: TOB);
    procedure ActiveEvents(status: boolean);
    function ControleChampOk(TOBL: TOB): boolean;
    procedure refreshGridListChamps;
    procedure DropUsed;

  public
    // Déclarations publiques
    Action : TActionFiche ;
    // Objets mémoire
    TobParFou, TobParFouLig : TOB;
  end;

var
  ParamImportTarFou: TParamImportTarFou;

implementation

{$R *.DFM}

procedure EntreeTarifAutoFour (Action : TActionFiche; StParFou : string);
var FF : TParamImportTarfou;
    PPANEL  : THPanel ;
begin
  SourisSablier;
  FF := TParamImportTarfou.Create(Application) ;
  FF.Action := Action ;
  if FF.Action = TaModif then
  begin
    FF.GRF_PARFOU.Text := StParFou;
  end;                                                                

  PPANEL := FindInsidePanel ; // permet de savoir si la forme dépend d'un PANEL
  if PPANEL = Nil then        // Le PANEL est le premier ecran affiché
  begin
    try
      FF.ShowModal ;
    finally
      FF.Free ;
    end ;
    SourisNormale ;
  end else
  begin
   InitInside (FF, PPANEL);
   FF.Show ;
  end ;
end ;

function GetValChamp (Table, Champ, Cle : string) : string;
var Q : TQuery;
begin
  Q:=OpenSQL('SELECT '+Champ+' from '+Table+' where '+Cle,TRUE,-1,'',true) ;
  result:='';
  if Not Q.EOF then result:=Q.FindField(Champ).AsString;
  Ferme(Q) ;
end;

procedure LoadLesTobParFou (StParFou : string; TobEntete, TobLigne : TOB);
var Q : TQuery;
begin
  TobEntete.SelectDB ('"' + StParFou + '"', Nil) ;

//  Q := OpenSQL ('SELECT * FROM PARFOULIG WHERE GFL_PARFOU="' + StParFou + '" ORDER BY GFL_OFFSET', True) ;
  Q := OpenSQL ('SELECT * FROM PARFOULIG WHERE GFL_PARFOU="' + StParFou + '"', True,-1,'',true) ;
  if not Q.Eof then TobLigne.LoadDetailDB ('PARFOULIG', '', '', Q, False);
  Ferme (Q);
end;

procedure TParamImportTarFou.GRF_FICHIERChange(Sender: TObject);
begin
  bModif := True;
end;

procedure TParamImportTarFou.GRF_FICHIERElipsisClick(Sender: TObject);
begin
  if OpenDialogButton.Execute then
      if OpenDialogButton.FileName <> '' then
          GRF_FICHIER.Text := OpenDialogButton.Filename;
end;

procedure TParamImportTarFou.GRF_FICHIERExit(Sender: TObject);
begin
  if csDestroying in ComponentState then Exit ;  //pour eviter erreur si sortie de l'application
end;

procedure TParamImportTarFou.GRF_TIERSExit(Sender: TObject);
begin
  if csDestroying in ComponentState then Exit ;  //pour eviter erreur si sortie de l'application
  AffecteLibelleTiers;
end;

procedure TParamImportTarFou.GRF_PARFOUExit(Sender: TObject);
begin
  if csDestroying in ComponentState then Exit ;  //pour eviter erreur si sortie de l'application

  if not GRF_TYPEENREG.Checked  then
  begin
    GRF_SEPARATEUR.Enabled := True;
    GRF_SEPARATEUR.Color := clWindow;
    GRF_SEPARATEUR.Value := 'AUT';
    GRF_SEPARATEUR.Text := 'Autre';
    GRF_SEPTEXTE.Enabled := True;
    GRF_SEPTEXTE.Color := clWindow;
    GRF_SEPTEXTE.Text := '';
    GRF_LONGENREG.Enabled := False;
    GRF_LONGENREG.Color := clActiveBorder;
//    GChamp.Options := GChamp.Options - [goEditing] + [goRowSelect];
		//
    GChamp.ColLengths [SAF_LONGUEUR] := -1;
    GChamp.ColLengths [SAF_OFFSET] := -1;
    GChamp.ColWidths  [SAF_LONGUEUR] := 0;
    GChamp.ColWidths [SAF_OFFSET] := 0;
	  //
    HMTrad.ResizeGridColumns (GChamp);
    InsereListeChamp ('PASS', 'Passer un champ');
  end;

  if (GRF_PARFOU.Text <> '') and (LastPar <> GRF_PARFOU.Text) then
  begin
    GRF_PARFOU.Text := Trim(UpperCase(GRF_PARFOU.Text));
		InitialiseListBox;
	  ChargeParametrage;
  	DropUsed;
  	//
  	AfficheGridListChamps;
  end;
  if GRF_MULTIFOU.Enabled then GRF_MULTIFOU.OnClick := GRF_MULTIFOUClick;
end;

procedure TParamImportTarFou.GRF_TYPEENREGClick(Sender: TObject);
var ChampListe : TOB;
    iRow : integer;
begin
  if csDestroying in ComponentState then Exit ;  //pour eviter erreur si sortie de l'application
  if GRF_TYPEENREG.Checked = True then
  begin
    GRF_SEPARATEUR.Enabled := False;
    GRF_SEPARATEUR.Color := clActiveBorder;
    GRF_SEPTEXTE.Enabled := False;
    GRF_SEPTEXTE.Color := clActiveBorder;
    GRF_LONGENREG.Enabled := True;
    GRF_LONGENREG.Color := clWindow;
    //GChamp.Options := GChamp.Options + [goEditing] - [goRowSelect];
    //
    GChamp.ColLengths [SAF_LONGUEUR] := 50;
    GChamp.ColLengths [SAF_OFFSET] := 50;
    GChamp.ColWidths  [SAF_LONGUEUR] := 50;
    GChamp.ColWidths [SAF_OFFSET] := 50;
    //
    HMTrad.ResizeGridColumns (GChamp);
    ChampListe := TOBListeChamps.findFirst (['CHAMPS'],['PASS'],True);
    if ChampListe <> Nil then ChampListe.free;
    //
		refreshGridListChamps;
    //
    iRow := 1;
    while iRow < GChamp.RowCount do
    begin
      if GChamp.Cells [SAF_Libelle, iRow] = 'PASS' then
      begin
        GChamp.Row := iRow;
        SupprimeLigneParfou (iRow);
        iRow := iRow - 1;
      end;
      iRow := iRow + 1;
    end;
    GChamp.Row := 1;
  end else
  begin
    GRF_SEPARATEUR.Enabled := True;
    GRF_SEPARATEUR.Color := clWindow;
    GRF_SEPARATEUR.Value := 'AUT';
    GRF_SEPARATEUR.Text := 'Autre';
    GRF_SEPTEXTE.Enabled := True;
    GRF_SEPTEXTE.Color := clWindow;
    GRF_SEPTEXTE.Text := '';
    GRF_LONGENREG.Enabled := False;
    GRF_LONGENREG.Color := clActiveBorder;
//    GChamp.Options := GChamp.Options - [goEditing] + [goRowSelect];
		//
    GChamp.ColLengths [SAF_LONGUEUR] := -1;
    GChamp.ColLengths [SAF_OFFSET] := -1;
    GChamp.ColWidths  [SAF_LONGUEUR] := 0;
    GChamp.ColWidths [SAF_OFFSET] := 0;
	  //
    HMTrad.ResizeGridColumns (GChamp);
    InsereListeChamp ('PASS', 'Passer un champ');
  end;
  bModif := True;
end;

procedure TParamImportTarFou.GRF_LONGENREGChange(Sender: TObject);
begin
  bModif := True;
end;

procedure TParamImportTarFou.GRF_SEPARATEURChange(Sender: TObject);
begin
  bModif := True;
end;

procedure TParamImportTarFou.GRF_SEPARATEURExit(Sender: TObject);
begin
  if csDestroying in ComponentState then Exit ;  //pour eviter erreur si sortie de l'application
  if (GRF_SEPARATEUR.Value <> 'AUT') and (GRF_SEPARATEUR.Value <> '') then
  begin
    GRF_SEPTEXTE.Enabled := False;
    GRF_SEPTEXTE.Color := clActiveBorder;
    if GRF_SEPARATEUR.Value = 'TAB' then GRF_SEPTEXTE.Text := GRF_SEPARATEUR.Value
    else if GRF_SEPARATEUR.Value = 'PIP' then GRF_SEPTEXTE.Text := '|'
    else if GRF_SEPARATEUR.Value = 'PTV' then GRF_SEPTEXTE.Text := ';';
  end else
  begin
    GRF_SEPARATEUR.Value := 'AUT';
    GRF_SEPARATEUR.Text := 'Autre';
    GRF_SEPTEXTE.Enabled := True;
    GRF_SEPTEXTE.Color := clWindow;
    GRF_SEPTEXTE.Text := '';
  end;
end;

procedure TParamImportTarFou.GRF_TYPEARTICLEExit(Sender: TObject);
begin
  if csDestroying in ComponentState then Exit ;  //pour eviter erreur si sortie de l'application
end;

procedure TParamImportTarFou.GRF_FAMILLETAXE1Exit(Sender: TObject);
begin
  if csDestroying in ComponentState then Exit ;  //pour eviter erreur si sortie de l'application
end;

procedure TParamImportTarFou.GRF_COMPTAARTICLEExit(Sender: TObject);
begin
  if csDestroying in ComponentState then Exit ;  //pour eviter erreur si sortie de l'application
end;

procedure TParamImportTarFou.GRF_PREFIXEExit(Sender: TObject);
begin
  if GRF_PREFIXE.Text <> '' then GRF_PREFIXE.Text := Trim(UpperCase(GRF_PREFIXE.Text));
end;

procedure TParamImportTarFou.GRF_SUFFIXEExit(Sender: TObject);
begin
  if GRF_SUFFIXE.Text <> '' then GRF_SUFFIXE.Text := Trim(UpperCase(GRF_SUFFIXE.Text));
end;

procedure TParamImportTarFou.GRF_SEPTEXTEExit(Sender: TObject);
begin
  if csDestroying in ComponentState then Exit ;  //pour eviter erreur si sortie de l'application
end;

(*
procedure TParamImportTarFou.TLChampColumnClick(Sender: TObject;Column: TListColumn);
begin
  ColumnToSort := Column.Index;
  (Sender as TCustomListView).AlphaSort;
end;

procedure TParamImportTarFou.TLChampCompare(Sender: TObject; Item1,Item2: TListItem; Data: Integer; var Compare: Integer);
var iInd: Integer;
begin
  if ColumnToSort = 0 then
  begin
    Compare := CompareText (Item1.Caption, Item2.Caption);
  end else
  begin
    iInd := ColumnToSort - 1;
    Compare := CompareText (Item1.SubItems[iInd], Item2.SubItems[iInd]);
  end;
end;
*)

function TParamImportTarFou.TypeChamps (CodeChamp : string) : string;
var Indice : Integer;
    TheChamp : TSChamp;
begin
  result := '';
  for Indice := 0 to LSChamp.Count -1 do
  begin
    TheChamp := TSChamp(LSChamp.Items [Indice]);
    if TheChamp.nom = CodeChamp then
    begin
      result := TheChamp.TypeD;
      break;
    end;
  end;
end;

function TParamImportTarFou.ChampOk (Acol,Arow : integer) : boolean;
var TOBL : TOB;
begin
  result := false;
  if (Arow < 0) or (Acol < 0) then exit; 
  TOBL := GetTOBLigne (Arow);  if TOBL = nil then exit;
  if (GChamp.ColLengths [Acol] = -1) then exit;
  if ((Acol = SAF_SepDecimale) or (Acol = SAF_NBRDecimale)) and
      (TypeChamps (TOBL.GetValue('GFL_CHAMP')) <> 'DOUBLE') then Exit;
  if (Acol = SAF_NBRDECIMALE) and (GChamp.CellValues [SAF_SEPDECIMALE,Arow] <> '') then Exit;
  result := True;
end;

function TParamImportTarFou.ChampSuivantOk (var Acol : integer ; var Arow : integer) : boolean;
var EnAvant : boolean;
begin
  result := true;  if TobParFouLig.Detail.count = 0 then exit;
  if ((GChamp.row = Arow) and (GChamp.col > Acol)) or (GChamp.row > Arow) then EnAvant := true else EnAvant := false;
  Acol := GChamp.col;
  Arow := GChamp.row;
//  GChamp.CacheEdit;
  GChamp.SynEnabled := false;
  while not ChampOk (Acol,Arow) do
  begin
    result := false;
    if EnAvant then inc (Acol) else Dec(Acol);
    if (Acol < GChamp.FixedCols) then
    begin
      Acol := GChamp.colcount -1;
      dec(Arow);
    end;
    if (ARow < 1) then
    begin
      Arow := GChamp.RowCount -1;
      Acol := GChamp.colcount -1;
    end;
    if (Acol > GChamp.ColCount -1 ) then
    Begin
      Inc(Arow);
      Acol := GChamp.fixedCols;
    END;
    if (Arow > GChamp.RowCount -1 ) then
    begin
      Arow := 1;
      Acol := GChamp.fixedCols;
    end;
  end;
  GChamp.SynEnabled := True;
//  GChamp.MontreEdit;
end;

procedure TParamImportTarFou.GChampCellEnter(Sender: TObject; var ACol,ARow: Integer; var Cancel: Boolean);
begin
  if Action = taConsult then Exit;
  if not ChampSuivantOk (Acol,Arow) then BEGIN Cancel := true; Exit; END;
  stCellCur := GChamp.Cells [Acol,Arow];
end;

procedure TParamImportTarFou.TraiterSepDecimale (ARow : integer);
Var TOBL : TOB ;
begin
TOBL := GetTOBLigne (ARow) ;
if TOBL = Nil then exit;
TOBL.PutValue ('GFL_SEPDECIMALE', GChamp.CellValues [SAF_SEPDECIMALE,Arow]);
end;

procedure TParamImportTarFou.GChampCellExit(Sender: TObject; var ACol,ARow: Integer; var Cancel: Boolean);
begin
  if csDestroying in ComponentState then Exit ;  //pour eviter erreur si sortie de l'application
  if GChamp.Cells [Acol,Arow] = stCellCur then exit;
  //
  if Acol = SAF_LIBELLE     then GChamp.cells [Acol,Arow] := stcellcur;

  if ACol = SAF_Champ       then TraiterChamp (ACol, ARow, 0) else
  if ACol = SAF_Longueur    then TraiterNombre (ACol, ARow, 'GFL_LONGUEUR') else
  if ACol = SAF_Offset      then TraiterNombre (ACol, ARow, 'GFL_OFFSET') else
  // Modif BTP
  if ACol = SAF_SepDecimale then TraiterSepDecimale (ARow) else
  if ACol = SAF_NbrDecimale then TraiterNombre (ACol, ARow, 'GFL_NBRDECIMALE') else
  ;
  // --
end;

procedure TParamImportTarFou.GChampEnter(Sender: TObject);
var Cancel, Chg : boolean;
    ACol, ARow : integer;
begin
  Cancel := False; Chg := False;
  Arow := GChamp.row;
  Acol := GChamp.row;
  GChampRowEnter (Sender, ARow , Cancel, Chg);
  GChampCellEnter (Sender, ACol, ARow, Cancel);
  GChamp.row := Arow; GChamp.col := Acol;
  GChamp.SetFocus;
end;

procedure TParamImportTarFou.GChampRowEnter(Sender: TObject; Ou: Integer;var Cancel: Boolean; Chg: Boolean);
VAR    TOBL : TOB;
begin
  TGFL_CHAMP.visible := false;
  TOBL := GetTOBLigne (Ou);
  if TOBL = nil then exit;
  GFL_FICHIER.text := TOBL.GetValue('GFL_FICHIER');
  TGFL_CHAMP.Caption := TOBL.getValue('GFL_CHAMP');
  TGFL_CHAMP.Visible := true;
  if (TypeChamps(TOBL.GetValue('GFL_CHAMP'))='COMBO') then
  begin
    BVisionFic.enabled := true;
    if (BvisionFic.down) and (not TTFICHIERASSOC.Visible) then TTFICHIERASSOC.Visible := true;
  end else
  begin
    TTFICHIERASSOC.Visible := false;
    BVisionFic.enabled := false;
  end;

end;

procedure TParamImportTarFou.GChampRowExit(Sender: TObject; Ou: Integer;var Cancel: Boolean; Chg: Boolean);
var TOBL : TOB;
begin
  if csDestroying in ComponentState then Exit ;  //pour eviter erreur si sortie de l'application
  TOBL := GetTOBLigne (Ou);  if TOBL = nil then exit;
  TOBL.putValue('GFL_FICHIER',GFL_FICHIER.Text);
  GChamp.InvalidateCell (SAF_FichierAssocie,Ou);
end;

function TParamImportTarFou.FamilleMereRenseigne (Niveau : integer) : boolean;
var QQ : TQuery;
{$IFNDEF BTP}
    TOBP : TOB;
    Req : string;
{$ENDIF}
begin
  result:= false;
{$IFDEF BTP}
	// pourquoi controler la famille ???
  // A partir du moment ou l'on a un profil c'est suffisant
  QQ := OpenSql ('SELECT GPF_PROFILARTICLE FROM PROFILART WHERE GPF_PROFILARTICLE="'+GRF_PROFILARTICLE.Text+'"',True,-1,'',true);
  if not QQ.eof Then result := true;
  Ferme (QQ);
{$ELSE}
  TOBP := TOB.Create ('PROFILART',nil,-1);
  TRY
    TOBP.putValue('GPF_PROFILARTICLE',GRF_PROFILARTICLE.Text);
    TOBP.loadDb (true);
    if TOBP.GetValue('GPF_FAMILLENIV'+inttostr(Niveau)) <> '' then result := true;
  FINALLY
    TOBP.Free;
  END;
{$ENDIF}
end;

function TParamImportTarFou.ControleChampOk (TOBL : TOB): boolean;
var Niveau : integer;
begin
  Result := true;
  if (copy(TOBL.getValue('CHAMP') ,1,13) = 'GA_FAMILLENIV') then
  begin
    Niveau := strtoint(copy(TOBL.getValue('CHAMP') ,14,1));
    if (Niveau > 1) and (GetParamSoc('SO_GCFAMHIERARCHIQUE')=True) then
    begin
      if (not CheckProfil.Checked) then BEGIN result := false; PgiBox('Veuillez renseigner un profil d''article'); Exit; END;
      if (CheckProfil.Checked) then
      begin
        if not FamilleMereRenseigne (Niveau-1) then
        BEGIN
          result := false;
          PgiBox('La famille de niveau '+inttostr(Niveau-1)+' n''est pas renseignée dans le profil');
          Exit;
        END;
      end;
    end;
  end;
end;

procedure TParamImportTarFou.BFlecheDroiteClick(Sender: TObject);
var ARow : integer;
		TOBLC : TOB;
begin
  if GListeChamps.RowCount = 0 then exit;
  TOBLC := TOBListeChamps.detail[GListeChamps.row-1];
  if not ControleChampOk(TOBLC) Then Exit;
  if (TOBLC.getValue('CHAMPS') <> 'PASS') or (GRF_TYPEENREG.Checked <> True) then
  begin
    RenseigneLigneGrid (TOBLC.getValue('CHAMPS'), TOBLC.getValue('LIBELLE'),ARow);
    AfficheLaLigne (ARow) ;
    TraiterChamp (SAF_Champ, GListeChamps.Row, 1);
    GChamp.row := Arow; GChamp.col := SAF_Libelle;
		GChampEnter (Self);
//    GChamp.enabled := true;
  end;
end;

procedure TParamImportTarFou.BFlecheGaucheClick(Sender: TObject);
var TOBL : TOB;
begin
  TOBL := GetTOBLigne (GChamp.row);
  if TOBL = nil then exit;
  if GChamp.Row < 1 then Exit;
  if GChamp.Row > TobParFouLig.Detail.Count then Exit;
  TraiterChamp (SAF_Champ, GChamp.Row, 2);
  GChamp.col := SAF_LIBELLE;
  GChampEnter (Self);
//  if TobParFouLig.detail.count = 0 then GChamp.Enabled := false;
end;

procedure TParamImportTarFou.bSupprimerClick(Sender: TObject);
var Select : string;
    mrResultat : integer;
begin
  if GRF_PARFOU.Text = '' then Exit;
  mrResultat := MsgBox.Execute (18, Caption, '');
  if (mrResultat = mrCancel) or (mrResultat = mrNo) then Exit;
  BeginTrans;
  Select := 'Delete from PARFOU where GRF_PARFOU="' + GRF_PARFOU.Text + '"';
  ExecuteSQL(Select);
  Select := 'Delete from PARFOULIG where GFL_PARFOU="' + GRF_PARFOU.Text + '"';
  ExecuteSQL(Select);
  try
    CommitTrans;
  except
    RollBack;
  end;
  Action := taCreat;
  GRF_PARFOU.Text := '';
  GRF_PARFOU.Enabled := True;
  TobParFouLig.ClearDetail;
  TobLigDelete.ClearDetail;
  InitialiseEntete;
end;

procedure TParamImportTarFou.BNouveauClick(Sender: TObject);
var bNouveau : boolean;
begin
  bNouveau := ControleModif;
  if bNouveau = True then
  begin
    Action := TaCreat;
    GRF_PARFOU.Text := '';
    GRF_PARFOU.Enabled := True;
    TobParFouLig.ClearDetail;
    TobLigDelete.ClearDetail;
    InitialiseEntete;
    bModif := False;
  end;
end;

procedure TParamImportTarFou.BValiderClick(Sender: TObject);
begin
  if (ControleEntete (True)) and (ControleLigne (True)) then
  begin
    ValideTarifAutoFour;
    TobParFou.SetAllModifie (False);
    bModif := False;
  end else ModalResult := 0;
end;

procedure TParamImportTarFou.BAideClick(Sender: TObject);
begin
CallHelpTopic(Self) ;
end;


{==============================================================================================}
{======================================= Gestion des Données ==================================}
{==============================================================================================}

procedure TParamImportTarfou.ChargeParametrage;
var iIndex : integer;
begin
  LoadLesTobParFou (GRF_PARFOU.Text, TobParFou, TobParFouLig);
  if not GRF_TYPEENREG.checked then TOBParFouLig.detail.Sort('GFL_OFFSET');
  if TOBParFouLig.detail.count > 0 then
  begin
    GChamp.RowCount := TobParFouLig.detail.count+1;
//    GChamp.enabled := true;
  end else
  begin
    GChamp.RowCount := 2;
//  	GChamp.enabled := false;
  end;

  AffecteLibelleLignes;
  AffecteEntete;
  DefiniGChamp;
  if GRF_TIERS.Text <> '' then
  begin
  	GRF_MULTIFOU.Checked := false;
  	GRF_MULTIFOU.enabled := false;
  end;
  if GRF_MULTIFOU.Enabled then
  begin
    if GRF_MULTIFOU.Checked then GRF_TIERS.Enabled := false
                            else GRF_TIERS.Enabled := True;
  end;
  for iIndex := 0 to TobParFouLig.Detail.Count - 1 do
  begin
  	AfficheLaLigne (iIndex + 1);
  end;
  HMTrad.ResizeGridColumns (GChamp) ;
end;

function TParamImportTarfou.ControleModif : boolean;
var mrResultat : integer;
begin
Result := True;
if (bModif = True) or (TobParFouLig.IsOneModifie) or (TobParFou.IsOneModifie) then
    begin
    mrResultat := MsgBox.Execute (15, Caption, '');
    if mrResultat = mrCancel then Result := False
    else
        begin
        Result := True;
        if mrResultat = mrYes then
            begin
            if (ControleEntete (True)) and (ControleLigne (True)) then ValideTarifAutoFour
            else Result := False;
            end;
        end;
    end;
end;

procedure TParamImportTarfou.DeleteParFouLig (ARow : integer);
var Index : integer;
begin
Index := TobLigDelete.Detail.Count;
TOB.Create ('PARFOULIG', TobLigDelete, Index);
TobLigDelete.Detail[Index].Dupliquer (TobParFouLig.Detail[ARow-1], False, True);
end;

{==============================================================================================}
{======================================= Initialisations ======================================}
{==============================================================================================}

Procedure TParamImportTarfou.InitialiseEntete;
begin
TobParFou.InitValeurs;
TobParFouLig.InitValeurs;
GRF_LONGENREG.Enabled   := False;
GRF_LONGENREG.Color     := clActiveBorder;
GRF_SEPTEXTE.Enabled    := False;
GRF_SEPTEXTE.Color      := clActiveBorder;

GRF_TIERS.Text          := '';
LGRF_TIERS.Caption      := '';
GRF_FICHIER.Text        := '';
GRF_TYPEENREG.Checked   := False;
GRF_SEPARATEUR.Value    := '';
GRF_SEPTEXTE.Text       := '';
GRF_LONGENREG.Value     := 0;
GRF_TYPEARTICLE.Value   := '';
GRF_FAMILLETAXE1.Value  := '';
GRF_COMPTAARTICLE.Value := '';
GRF_TENUESTOCK.Checked  := False;
GRF_REMISEPIED.Checked  := False;
GRF_ESCOMPTABLE.Checked := False;
GRF_SUFFIXE.Text        := '';
GRF_PREFIXE.Text        := '';
// MODIF BTP
GRF_PROFILARTICLE.text  := '';
// --
ExtractFields ('ARTICLECOMPL', 'X', ListeLibChampArtC, ListeChampArtC,  ListTypeChampArtC,False);
ExtractFields ('CATALOGU',     'L', ListeLibChampCat,  ListeChampCat,    ListTypeChampCat, False);
ExtractFields ('ARTICLE',      'L', ListeLibChampArt,  ListeChampArt,    ListTypeChampArt, False);
//FV1 : 17/12/2015 - FS#1838 - En import catalogue, ajouter la récupération du code barre de l'article.
ExtractFields ('BTCODEBARRE',  'L', ListeLibChampCab,  ListeChampCab,    ListTypeChampCab, False);
//
InitialiseGrilles;
InitialiseListBox;

if Action = taCreat then
    begin
    GRF_PARFOU.Enabled := True;
    GRF_TIERS.Enabled := True;
    GRF_PARFOU.SetFocus
    end else
    begin
    GRF_PARFOU.Enabled := False;
    GRF_TIERS.Enabled := False;
    GRF_FICHIER.SetFocus;
    end;
end;

procedure TParamImportTarfou.InitialiseGrilles ;
begin
  // GChamps
  GChamp.VidePile(False);
  GChamp.RowCount:= 2;
  DefiniGChamp;
  if Action = TaModif then begin end //ChargeParametrage
  else
  begin
    GChamp.ColWidths [SAF_Longueur] := 0;
    GChamp.ColWidths [SAF_Offset] := 0;
    GChamp.ColLengths [SAF_LONGUEUR] := -1;
    GChamp.ColLengths [SAF_OFFSET] := -1;
    HMTrad.ResizeGridColumns (GChamp);
  end;
  // GlisteChamps
  GListeChamps.VidePile(false);
  //
  GlisteChamps.ColCount := 2;
  GListeChamps.Cells [0,0] := 'Champs';
  GListeChamps.Cells [1,0] := 'Désignation';
  //
end;

procedure TParamImportTarfou.InitialiseLigne (ARow : integer) ;
Var TOBL : TOB ;
begin
TOBL := GetTOBLigne (ARow) ;
if TOBL <> Nil then TOBL.InitValeurs;
TOBL.PutValue ('LIBELLE', '');
GChamp.Rows [ARow].Clear;
end;

procedure TParamImportTarfou.AjouteLSChamp (champ,TypeD : string);
var TheSChamp : TSChamp;
begin
TheSChamp := TSCHamp.Create;
TheSChamp.nom := Champ;
TheSChamp.TypeD := TypeD;
LSChamp.Add (TheSChamp);
end;

procedure TParamImportTarfou.InitialiseListBox ;
var TobPar : TOB;
    iChamp : integer;
begin
TOBListeChamps.clearDetail;

  for iChamp := 0 to ListeChampArtC.Count - 1 do
  begin
    TobPar := TobParFouLig.FindFirst (['GFL_CHAMP'], [ListeChampArtC.Strings[iChamp]], False);
    AjouteLSChamp (ListeChampArtC.Strings [iChamp], ListTypeChampArtC [ichamp]);
    InsereListeChamp (ListeChampArtC.Strings [iChamp],ListeLibChampArtC.Strings [iChamp]);
  end;

  for iChamp := 0 to ListeChampCat.Count - 1 do
  begin
    TobPar := TobParFouLig.FindFirst (['GFL_CHAMP'], [ListeChampCat.Strings[iChamp]], False);
    AjouteLSChamp (ListeChampCat.Strings [iChamp],ListTypeChampCat [ichamp]);
    if ((ListeChampCat.Strings[iChamp] <> 'GCA_TIERS')    OR (GRF_MULTIFOU.checked)) and
       ( ListeChampCat.Strings[iChamp] <> 'GCA_DATESUP') and (TobPar = Nil) then
       InsereListeChamp (ListeChampCat.Strings [iChamp], ListeLibChampCat.Strings [iChamp]);
  end;

  for iChamp := 0 to ListeChampArt.Count - 1 do
  begin
    TobPar := TobParFouLig.FindFirst (['GFL_CHAMP'], [ListeChampArt.Strings[iChamp]], False);
    AjouteLSChamp (ListeChampArt.Strings [iChamp], ListTypeChampArt [ichamp]);
    if (ListeChampArt.Strings[iChamp] <> 'GA_ARTICLE') and (TobPar = Nil) then
        InsereListeChamp (ListeChampArt.Strings [iChamp], ListeLibChampArt.Strings [iChamp]);
  End;

  //FV1 : 17/12/2015 - FS#1838 - En import catalogue, ajouter la récupération du code barre de l'article.
  for iChamp := 0 to ListeChampCab.Count - 1 do
  begin
    TobPar := TobParFouLig.FindFirst (['GFL_CHAMP'], [ListeChampCab.Strings[iChamp]], False);
    AjouteLSChamp (ListeChampCab.Strings [iChamp], ListTypeChampCab [ichamp]);
    InsereListeChamp (ListeChampCab.Strings [iChamp], ListeLibChampCab.Strings [iChamp]);
  End;

  if (GRF_TYPEENREG.Checked = False) then InsereListeChamp ('PASS', 'Passer un champ');

end;

{==============================================================================================}
{=============================== Initialisation de la Grid ====================================}
{==============================================================================================}

procedure TParamImportTarfou.DefiniGChamp;
var Nam,st: string;
		iCol : integer;
begin
  GChamp.colcount := Nbcols;
  if TOBParFouLig.detail.count > 0 then
  begin
    GChamp.RowCount := TobParFouLig.detail.count+1;
//    GChamp.enabled := true;
  end else
  begin
    GChamp.RowCount := 2;
//  	GChamp.enabled := false;
  end;

  Gchamp.Cells [SAF_SELECT,0] := '';
  GChamp.ColEditables [SAF_SELECT] := false;
  Gchamp.ColWidths [SAF_SELECT]  := 10;
  GChamp.ColLengths [SAF_SELECT] := -1;
  GChamp.ColLengths [SAF_NbrDecimale] := 25;

  Gchamp.Cells [SAF_LIBELLE,0] := 'Désignation';
  Gchamp.ColAligns [SAF_LIBELLE]  := taLeftJustify;
  Gchamp.ColWidths [SAF_LIBELLE]  := 80;
  GChamp.ColLengths [SAF_LIBELLE] := 80;
  Gchamp.ColEditables  [SAF_LIBELLE]  := true;

  Gchamp.ColAligns [SAF_Champ]  := taLeftJustify;
  Gchamp.ColWidths [SAF_CHAMP]  := 0;
  GChamp.ColLengths [SAF_CHAMP] := -1;

  Gchamp.Cells [SAF_LONGUEUR,0] := 'Longueur';
  Gchamp.ColAligns [SAF_Longueur]  := taRightJustify;
  GChamp.ColLengths [SAF_LONGUEUR] := 50;
  GChamp.ColWidths [SAF_Longueur] := 50;
  GChamp.ColFormats [SAF_LONGUEUR] := '####0;-####0; ;';

  Gchamp.Cells [SAF_OFFSET,0] := 'Début';
  Gchamp.ColAligns [SAF_OFFSET]  := taRightJustify;
  GChamp.ColLengths [SAF_OFFSET] := 50;
  GChamp.ColWidths [SAF_Offset] := 50;
  GChamp.ColFormats [SAF_LONGUEUR] := '####0;-####0; ;';

  Gchamp.Cells [SAF_SepDecimale,0] := 'Séparateur Dec.';
  GChamp.ColLengths [SAF_SepDecimale] := 45;
  GChamp.ColWidths [SAF_SepDecimale] := 45;
  GChamp.ColFormats  [SAF_SepDecimale] := 'CB=GCSEPDECIMALE||<<Aucun>>';

  Gchamp.Cells [SAF_NbrDecimale,0] := 'Nbr décimale';
  GChamp.ColLengths [SAF_NbrDecimale] := 25;
  GChamp.ColWidths [SAF_NbrDecimale] := 25;
  GChamp.ColFormats  [SAF_NbrDecimale] := '####0;-####0; ;';

  GChamp.Cells[SAF_FichierAssocie,0] := 'Attach.';
  GCHamp.ColWidths [SAF_FichierAssocie] := 30 ;
  GCHamp.Collengths [SAF_FichierAssocie] := -1 ;

  if not GRF_TYPEENREG.Checked then
  begin
    GChamp.ColLengths [SAF_LONGUEUR] := -1;
    GChamp.ColLengths [SAF_OFFSET] := -1;
    GChamp.ColWidths  [SAF_LONGUEUR] := 0;
    GChamp.ColWidths [SAF_OFFSET] := 0;
  end;
end;

procedure TParamImportTarfou.EtudieColsListe ;
var Lescols,NomCol : string;
    Icol : integer;
begin
  LesCols:=LesColChamp ; icol:=0 ;
  Repeat
    NomCol:=uppercase(Trim(ReadTokenSt(LesCols))) ;
    if NomCol<>'' then
    BEGIN
      if NomCol='SELECT'   then SAF_SELECT:=icol else
      if NomCol='LIBELLE'   then SAF_LIBELLE:=icol else
      if NomCol='GFL_CHAMP' then SAF_Champ:=icol else
      if NomCol='GFL_LONGUEUR' then SAF_Longueur:=icol else
      if NomCol='GFL_OFFSET' then SAF_OffSet:=icol else
      if NomCol='GFL_SEPDECIMALE' then SAF_SepDecimale:=icol else
      if NomCol='GFL_FICHIER' then SAF_FichierAssocie := Icol else
      if NomCol='GFL_NBRDECIMALE' then SAF_NbrDecimale:=icol else
      ;
    END;
    Inc(icol) ;
  Until ((LesCols='') or (NomCol='')) ;
  nbCols := icol;
end ;

Procedure TParamImportTarfou.InitLesCols;
BEGIN
SAF_Champ           := -1;
SAF_Longueur        := -1;
SAF_Offset          := -1;
SAF_Libelle         := -1;
SAF_SEPDecimale     := -1;
SAF_NbrDecimale     := -1;
SAF_FichierAssocie  := -1;
SAF_SELECT          := -1;
END ;

{==============================================================================================}
{================================ Manipulation de l'entete ====================================}
{==============================================================================================}

procedure TParamImportTarfou.AffecteEntete;
begin
if (not TobParFou.GetBoolean('GRF_MULTIFOU')) and (TOBPARFOU.GetValue('GRF_TIERS') = '') then exit; // evite de reinitialiser
GRF_TIERS.Text := TobParFou.GetValue ('GRF_TIERS');
AffecteLibelleTiers;
GRF_FICHIER.Text := TobParFou.GetValue ('GRF_FICHIER');
if TobParFou.GetValue ('GRF_TYPEENREG') = 'X' then
    begin
    GRF_TYPEENREG.Checked := True;
    GRF_SEPARATEUR.enabled := False;
    GRF_SEPARATEUR.Color := clActiveBorder;
    GRF_SEPTEXTE.enabled := False;
    GRF_SEPTEXTE.Color := clActiveBorder;
    GRF_LONGENREG.Value := TobParFou.GetValue ('GRF_LONGENREG');
    end else
    begin
    GRF_TYPEENREG.Checked := False;
    GRF_SEPARATEUR.enabled := True;
    GRF_SEPARATEUR.Color := clWindow;
    GRF_SEPARATEUR.Value := TobParFou.GetValue ('GRF_SEPARATEUR');
    if GRF_SEPARATEUR.Value = 'AUT' then
        begin
        GRF_SEPTEXTE.Text := TobParFou.GetValue ('GRF_SEPTEXTE');
        GRF_SEPTEXTE.enabled := True;
        GRF_SEPTEXTE.Color := clWindow;
        end else
        begin
        GRF_SEPTEXTE.enabled := False;
        GRF_SEPTEXTE.Color := clActiveBorder;
        end;
    GRF_SEPTEXTE.Text := TobParFou.GetValue ('GRF_SEPTEXTE');
    GChamp.ColLengths [SAF_LONGUEUR] := -1;
    GChamp.ColLengths [SAF_OFFSET] := -1;
    GChamp.ColWidths [SAF_Longueur] := 0;
    GChamp.ColWidths [SAF_Offset] := 0;
    HMTrad.ResizeGridColumns (GChamp);
    end;
GRF_Prestation.text := TobParFou.GetValue ('GRF_PRESTATION');
PRESTATION.text := copy (GRF_PRESTATIOn.Text,1,18);
GRF_TYPEARTICLE.Value := TobParFou.GetValue ('GRF_TYPEARTICLE');
GRF_FAMILLETAXE1.Value := TobParFou.GetValue ('GRF_FAMILLETAXE1');
GRF_COMPTAARTICLE.Value := TobParFou.GetValue ('GRF_COMPTAARTICLE');
if TobParFou.GetValue ('GRF_TENUESTOCK') = 'X' then GRF_TENUESTOCK.Checked := True
else GRF_TENUESTOCK.Checked := False;
if TobParFou.GetValue ('GRF_REMISEPIED') = 'X' then GRF_REMISEPIED.Checked := True
else GRF_REMISEPIED.Checked := False;
if TobParFou.GetValue ('GRF_ESCOMPTABLE') = 'X' then GRF_ESCOMPTABLE.Checked := True
else GRF_ESCOMPTABLE.Checked := False;
GRF_SUFFIXE.Text := TobParFou.GetValue ('GRF_SUFFIXE');
GRF_PREFIXE.Text := TobParFou.GetValue ('GRF_PREFIXE');
GRF_PROFILARTICLE.Text := TobParFou.GetValue ('GRF_PROFILARTICLE');
if GRF_PROFILARTICLE.Text <> '' then CheckProfil.Checked := true
                                else CheckProfil.Checked := false;
GRF_FOURPRINC.Checked := (TobParFou.GetValue ('GRF_FOURPRINC')='X');
GRF_MULTIFOU.Checked := (TobParFou.GetValue ('GRF_MULTIFOU')='X');
end;


procedure TParamImportTarfou.AddChampsSupLig (TOBL : TOB);
begin
	TOBL.AddChampSupValeur  ('LIBELLE', '');
	TOBL.AddChampSupValeur  ('SELECT', '');
end;

procedure TParamImportTarfou.AffecteLibelleLigne (TOBL : TOB);
var LibelleChamp : string;
begin
  if TobL.GetValue ('GFL_CHAMP') <> 'PASS' then
  begin
  	LibelleChamp := RecupereLibelleChamp (TobL.GetValue ('GFL_CHAMP'));
  end else
  begin
  	LibelleChamp := 'Passer un champ';
  end;
  AddChampsSupLig (TobL);
  TobL.putValue  ('LIBELLE', LibelleChamp);
end;

procedure TParamImportTarfou.AffecteLibelleLignes;
var Index : integer;
begin
  for Index := TobParFouLig.Detail.Count - 1 downto 0 do
  begin
    AffecteLibelleLigne (TobParFouLig.Detail[Index]);
  end;
end;

procedure TParamImportTarfou.AffecteLibelleTiers;
var LibelleTiers : string;
begin
  if GRF_TIERS.Text <> '' then
  begin
    LibelleTiers := GetValChamp ('TIERS', 'T_LIBELLE',
                                 'T_TIERS="' + GRF_TIERS.Text + '" AND T_NATUREAUXI = "FOU"');
    LGRF_TIERS.Caption := LibelleTiers;
  end;
end;

function TParamImportTarfou.ControleEntete (Mess : boolean) : boolean ;
var LibelleTiers : string;
    Index, LgDet : integer;
begin
  Result := False;
  if (GRF_PARFOU.Text = '')then
  begin
    if Mess then MsgBox.Execute (0, Caption, '') ;
    if PageControl1.ActivePageIndex <> 0 then PageControl1.activepageIndex := 0;
    GRF_PARFOU.SetFocus;
    exit;
  end;
  if (GRF_TIERS.Text = '')  and (not GRF_MULTIFOU.checked)  then
  begin
    if Mess then MsgBox.Execute (1, Caption, '') ;
    if PageControl1.ActivePageIndex <> 0 then PageControl1.activepageIndex := 0;
    GRF_TIERS.SetFocus;
    exit;
  end else
  begin
    if (not GRF_MULTIFOU.checked) then
    begin
      LibelleTiers := GetValChamp ('TIERS', 'T_LIBELLE',
                                   'T_TIERS="' + GRF_TIERS.Text + '" AND T_NATUREAUXI = "FOU"');
      if LibelleTiers = '' then
      begin
        if Mess then MsgBox.Execute (2, Caption, '') ;
        exit;
      end;
    end;
  end;

  if GRF_TYPEENREG.Checked = True then
  begin
    if GRF_LONGENREG.Value = 0 then
    begin
      if Mess then MsgBox.Execute (4, Caption, '') ;
      if PageControl1.ActivePageIndex <> 0 then PageControl1.activepageIndex := 0;
      GRF_LONGENREG.SetFocus;
      exit;
    end
//  verif cohérence longueur saisie - detail des champs...
    else
    begin
    {
      LgDet := 0;
      for Index := 0 to TobParFouLig.Detail.Count - 1 do
          LgDet := LgDet + TobParFouLig.Detail[Index].GetValue('GFL_LONGUEUR');
      if LgDet > GRF_LONGENREG.Value then
      begin
        if Mess then MsgBox.Execute (17, Caption, '') ;
        if PageControl1.ActivePageIndex <> 0 then PageControl1.activepageIndex := 0;
        GRF_LONGENREG.SetFocus;
        exit;
      end
    }
    end;
  end else
  begin
    if GRF_SEPARATEUR.Text = '' then
    begin
      if Mess then MsgBox.Execute (6, Caption, '');
      if PageControl1.ActivePageIndex <> 0 then PageControl1.activepageIndex := 0;
      GRF_SEPARATEUR.SetFocus;
      exit;
    end else
    begin
      if (GRF_SEPARATEUR.Value = 'AUT') then
      begin
        if GRF_SEPTEXTE.Text = '' then
        begin
          if Mess then MsgBox.Execute (7, Caption, '');
          if PageControl1.ActivePageIndex <> 0 then PageControl1.activepageIndex := 0;
          GRF_SEPTEXTE.SetFocus;
          exit;
        end else
      end else
      begin
        if GRF_SEPARATEUR.Value = 'TAB' then GRF_SEPTEXTE.Text := GRF_SEPARATEUR.Value
        else if GRF_SEPARATEUR.Value = 'PIP' then GRF_SEPTEXTE.Text := '|'
        else if GRF_SEPARATEUR.Value = 'PTV' then GRF_SEPTEXTE.Text := ';';
      end;
    end;
  end;

  if (GRF_TYPEARTICLE.Text = '') and (not CheckProfil.Checked)  then
  begin
    if Mess then MsgBox.Execute (8, Caption, '');
    if PageControl1.ActivePageIndex <> 1 then PageControl1.activepageIndex := 1;
    GRF_TYPEARTICLE.SetFocus;
    exit;
  end;

  if (GRF_FAMILLETAXE1.Text = '') and (not CheckProfil.Checked) then
  begin
    if Mess then MsgBox.Execute (9, Caption, '');
    if PageControl1.ActivePageIndex <> 1 then PageControl1.activepageIndex := 1;
    GRF_FAMILLETAXE1.SetFocus;
    exit;
  end;

  if (GRF_COMPTAARTICLE.Text = '') and (not CheckProfil.Checked) then
  begin
    if Mess then MsgBox.Execute (13, Caption, '');
    if PageControl1.ActivePageIndex <> 1 then PageControl1.activepageIndex := 1;
    GRF_COMPTAARTICLE.SetFocus;
    exit;
  end;

  if (CheckProfil.Checked) and (GRF_PROFILARTICLE.text= '') then
  begin
    if Mess Then MsgBox.Execute (19, Caption, '');
    if PageControl1.ActivePageIndex <> 1 then PageControl1.activepageIndex := 1;
    GRF_PROFILARTICLE.SetFocus;
    exit;
  end;

  Result := True;
end;

function TParamImportTarfou.ControleLigne (Validation : boolean) : boolean;
var Index, ACol : integer;
    CodeReference,CodeTiers : boolean;
begin
  CodeReference := False;
  Result := True;
  Index := 0;
  if not GRF_MULTIFOU.Checked then CodeTiers := True else CodeTiers := false;
  if TobParFouLig.Detail.Count > 0 then
  begin
    repeat
      ACol := -1;
      if ligneVide (Index + 1, ACol) then
      begin
        if ACol <> -1 then
        begin
          Result := False;
          CodeReference := True;
          break;
        end;
      end else
      begin
        if (Validation) and (GRF_MULTIFOU.checked) and
           (TobParFouLig.Detail[Index].GetValue ('GFL_CHAMP') = 'GCA_TIERS') then CodeTiers := True;
    		if (Validation) and
           (TobParFouLig.Detail[Index].GetValue ('GFL_CHAMP') = 'GCA_REFERENCE') then CodeReference := True;
      	Index := Index + 1;
      end;
    until (Index >= TobParFoulig.Detail.Count) or (ACol <> -1);
  end;

  if (validation) and (CodeReference = False) then
  begin
    MsgBox.Execute (12, Caption, '');
    Result := False;
  end else if (validation) and (not Codetiers) then
  begin
    MsgBox.Execute (20, Caption, '');
    Result := False;
  end else
  begin
    if ACol <> -1 then
    begin
      MsgBox.Execute (14, Caption, '');
      if PageControl1.ActivePageIndex <> 0 then PageControl1.activepageIndex := 0;
      GChamp.Row := Index+1;
      GChamp.Col := ACol;
      stCellCur := GChamp.cells[GChamp.row,Gchamp.row];
      GChamp.SetFocus;
    end;
  end;
end;

{==============================================================================================}
{=============================== Manipulation des lignes ======================================}
{==============================================================================================}

procedure TParamImportTarfou.AfficheLaLigne (ARow : integer) ;
Var TOBL : TOB ;
//    LibelleChamp : string;
begin
TOBL := GetTOBLigne (ARow) ;
if TOBL = Nil then exit;
TOBL.PutLigneGrid (GChamp, ARow, False, False, LesColChamp);
(*
// Inverse le Code et le libelle a l'affichage
LibelleChamp := GChamp.Cells [SAF_Libelle, ARow];
GChamp.Cells [SAF_Libelle, ARow] := GChamp.Cells [SAF_Champ, ARow];
GChamp.Cells [SAF_Champ, ARow] := LibelleChamp;
*)
if GRF_TYPEENREG.Checked = False then
    begin
    GChamp.Cells [SAF_Longueur, ARow] := '';
    GChamp.Cells [SAF_Offset, ARow] := '';
    end;
end ;

Procedure TParamImportTarfou.CreerTOBLigne (ARow : integer);
var TOBL : TOB;
begin
if ARow <> TobParFouLig.Detail.Count + 1 then exit;
TOBL := TOB.Create ('PARFOULIG', TobParFouLig, ARow-1) ;
//
AddChampsSupLig (TOBL);
//
InitialiseLigne (ARow);
end;

Function TParamImportTarfou.GetTOBLigne (ARow : integer) : TOB ;
begin
Result:=Nil ;
if ((ARow <= 0) or (ARow > TobParFouLig.Detail.Count)) then Exit ;
Result := TobParFouLig.Detail [ARow-1] ;
end ;

procedure TParamImportTarfou.AddLesChampsSup (TOBL : TOB);
begin
	TOBL.AddChampSupValeur ('CHAMPS','');
	TOBL.AddChampSupValeur ('LIBELLE','');
end;

procedure TParamImportTarfou.refreshGridListChamps;
begin
  GListeChamps.VidePile(false);
  //
  GlisteChamps.ColCount := 2;
  GListeChamps.Cells [0,0] := 'Champs';
  GListeChamps.Cells [1,0] := 'Désignation';
  AfficheGridListChamps;
end;

procedure TParamImportTarfou.InsereListeChamp (stChamp, stLib : string);
var LibelleChamp : string;
    NouveauChamp : TListItem;
    TOBL : TOB;
begin
  if stChamp <> '' then
  begin
    if stLib <> '' then LibelleChamp := stLib
    else LibelleChamp := RecupereLibelleChamp (stChamp);

    if LibelleChamp <> '' then
    begin
      TOBL := TOBListeChamps.findFirst(['CHAMPS'],[stChamp],true);
      if TOBL = nil then
      begin
        TOBL := TOB.Create ('UN CHAMPS',TOBListeChamps,-1);
        AddLesChampsSup (TOBL);
        TOBL.PutValue('CHAMPS',stChamp);
        TOBL.PutValue('LIBELLE',LibelleChamp);
        TOBListeChamps.detail.sort('CHAMPS');
      end;
      refreshGridListChamps;
    end;
  end;
end;

Function TParamImportTarfou.LigneVide (ARow : integer; var ACol : integer) : boolean;
var TOBL : TOB;
begin
  Result := false;
  TOBL := GetTOBLigne (Arow);
  if (GRF_TYPEENREG.Checked = true) and ((TOBL.GetValue('GFL_LONGUEUR') = 0) or (TOBL.GetValue('GFL_OFFSET')=0)) then
  begin
    Acol := SAF_Offset;
    result := true;
    exit
  end;
end;

function TParamImportTarfou.RecupereLibelleChamp (StChamp : string) : string;
var iInd : integer;
    ListeChamp, ListeLib : HTStringList;
begin
  if pos ('GCA', stChamp) > 0 then
  begin
    ListeChamp := ListeChampCat;
    ListeLib := ListeLibChampCat;
  end
  else if Pos ('GA2',StChamp) > 0 then
  BEGIN
    ListeChamp := ListeChampArtc;
    ListeLib := ListeLibChampArtc;
  end
  else if Pos ('BCB',StChamp) > 0 then
  BEGIN
    ListeChamp := ListeChampCab;
    ListeLib := ListeLibChampCab;
  end
  else
  begin
    ListeChamp := ListeChampArt;
    ListeLib := ListeLibChampArt;
  end;

  iInd := 0;
  while (iInd < ListeChamp.Count) and (ListeChamp.Strings [iInd] <> stChamp) do
  iInd := iInd + 1;

  if iInd < ListeChamp.Count then
    Result := ListeLib.Strings [iInd]
  else
    Result := 'Champ inconnu';

end;

procedure TParamImportTarfou.RenseigneLigneGrid (StChamp ,Libelle: string; var IndiceRow : integer);
var LibelleChamp : string;
begin

  IndiceRow := TobParFouLig.Detail.count + 1;
  if IndiceRow >= GChamp.RowCount then GChamp.RowCount := GChamp.RowCount + 1;
  CreerTobLigne (IndiceRow);
  //
  TobParFouLig.Detail[IndiceRow-1].PutValue ('LIBELLE', Libelle);
  TobParFouLig.Detail[IndiceRow-1].PutValue ('GFL_CHAMP', StChamp);
end;

procedure TParamImportTarfou.SupprimeLigneParfou (ARow : Longint) ;
var TOBL : TOB;
begin
if Action=taConsult then Exit ;
if ARow < 1 then Exit ;
TOBL := GETTOBLigne (Arow);
if (TobL.GetValue ('GFL_CHAMP') <> 'PASS') then InsereListeChamp (TobL.GetValue ('GFL_CHAMP'), TobL.GetValue ('LIBELLE'));

GChamp.CacheEdit;
GChamp.SynEnabled := False;
GChamp.DeleteRow (ARow);

if (TobL.GetValue ('GFL_PARFOU') <> '') and (Action = taModif) then
    begin
    DeleteParFouLig (ARow);
    end;

TobL.free;
//if GChamp.RowCount < NbRowsInit then GChamp.RowCount := NbRowsInit;
GChamp.MontreEdit;
GChamp.SynEnabled := True;
bModif := True;
end;

procedure TParamImportTarfou.TraiterChamp (ACol, ARow, Provenance : integer);
var TOBL,TOBP : TOB;
    ChampListe : TListItem;
begin
    case Provenance of
    0 : begin
        end;
    1 : begin
  				TOBL := TOBListeChamps.detail[Arow-1];
          if TOBL <> nil then
          begin
            if TOBL.GetValue('CHAMPS') <> 'PASS' then
            begin
            	TOBL.free;
            	GListeChamps.DeleteRow (Arow);
            end;
          end;
        end;
    2 : begin
        	SupprimeLigneParfou (ARow) ;
        end;
    end;
end;

procedure TParamImportTarfou.TraiterNombre (ACol, ARow : integer; stCol : String);
var TOBL : TOB;
begin
if ((GChamp.Cells [ACol, ARow] <> '') and (GChamp.Cells [ACol, ARow] <> '0')) or
   (GRF_TYPEENREG.Checked = False) then
    begin
    TOBL := GetTOBLigne (ARow);
    TOBL.PutValue (stCol, StrToInt (GChamp.Cells [ACol, ARow]));
    end else
    begin
      GChamp.Cells [ACol, ARow] := stCellCur;
    end;
end;

{====================================================================== ========================}
{================================= Validation =================================================}
{==============================================================================================}

procedure TParamImportTarfou.EnregistreTarifAutoFour;
begin
TobParFou.SetAllModifie (True);
TobParFou.InsertOrUpdateDB;
if Action = taModif then TobLigDelete.DeleteDB;
TobParFouLig.SetAllModifie (True);
TobParFouLig.DeleteDB;
TobParFouLig.InsertOrUpdateDB;
TobParFouLig.SetAllModifie (False);
end;

procedure TParamImportTarfou.RenseigneDetail;
var iIndex, ACol : integer;
begin
for iIndex := 0 to TobParFouLig.Detail.Count - 1 do
    begin
    if not LigneVide (iIndex + 1, ACol) then
        begin
        TobParFouLig.Detail [iIndex].PutValue ('GFL_PARFOU', GRF_PARFOU.Text);
        if GChamp.Cells[SAF_Longueur, iIndex + 1] = '' then
            TobParFouLig.Detail [iIndex].PutValue ('GFL_LONGUEUR', 0);
        if GRF_TYPEENREG.Checked = False then
            TobParFouLig.Detail [iIndex].PutValue ('GFL_OFFSET', iIndex)
        else if GChamp.Cells[SAF_Offset, iIndex + 1] = '' then
            TobParFouLig.Detail [iIndex].PutValue ('GFL_OFFSET', 0);
        end;
    end;
end;

procedure TParamImportTarfou.RenseigneEntete;
begin
TobParFou.PutValue ('GRF_PARFOU', GRF_PARFOU.Text);
TobParFou.PutValue ('GRF_TIERS', GRF_TIERS.Text);
TobParFou.PutValue ('GRF_FICHIER', GRF_FICHIER.Text);
if GRF_TYPEENREG.Checked = True then
    TobParFou.PutValue ('GRF_TYPEENREG', 'X')
else
    begin
    TobParFou.PutValue ('GRF_TYPEENREG', '-');
    GRF_LONGENREG.Value := 0;
    end;
TobParFou.PutValue ('GRF_SEPARATEUR', GRF_SEPARATEUR.Value);
TobParFou.PutValue ('GRF_SEPTEXTE', GRF_SEPTEXTE.Text);
TobParFou.PutValue ('GRF_TYPEARTICLE', GRF_TYPEARTICLE.Value);
TobParFou.PutValue ('GRF_PRESTATION', GRF_PRESTATION.Text);

TobParFou.PutValue ('GRF_LONGENREG', StrToInt (GRF_LONGENREG.Text));
if GRF_TENUESTOCK.Checked = True then
    TobParFou.PutValue ('GRF_TENUESTOCK', 'X')
else TobParFou.PutValue ('GRF_TENUESTOCK', '-');
TobParFou.PutValue ('GRF_FAMILLETAXE1', GRF_FAMILLETAXE1.Value);
TobParFou.PutValue ('GRF_COMPTAARTICLE', GRF_COMPTAARTICLE.Value);
if GRF_REMISEPIED.Checked = True then
    TobParFou.PutValue ('GRF_REMISEPIED', 'X')
else TobParFou.PutValue ('GRF_REMISEPIED', '-');
if GRF_ESCOMPTABLE.Checked = True then
    TobParFou.PutValue ('GRF_ESCOMPTABLE', 'X')
else TobParFou.PutValue ('GRF_ESCOMPTABLE', '-');
TobParFou.PutValue ('GRF_PREFIXE', GRF_PREFIXE.Text);
TobParFou.PutValue ('GRF_SUFFIXE', GRF_SUFFIXE.Text);
//
if GRF_FOURPRINC.Checked then TobParFou.PutValue ('GRF_FOURPRINC','X' )
                         else TobParFou.PutValue ('GRF_FOURPRINC','-' );
TobParFou.PutValue ('GRF_PROFILARTICLE', GRF_PROFILARTICLE.Text);
TobParFou.SetBoolean  ('GRF_MULTIFOU', GRF_MULTIFOU.Checked);
//
end;

procedure TParamImportTarfou.ValideTarifAutoFour;
begin
RenseigneDetail;
RenseigneEntete;
EnregistreTarifAutoFour;
end;


procedure TParamImportTarFou.FormClose(Sender: TObject;var Action: TCloseAction);
var Indice : integer;
begin

for indice:= 0 to LStructure.Count -1 do
begin
  TSStruct(LStructure.Items [Indice]).free;
end;

for indice:= 0 to LSChamp.Count -1 do
begin
  TSChamp(LSCHamp.Items [Indice]).free;
end;

  Lstructure.free;

  GChamp.VidePile(True) ;

  FreeAndNil(TobParFou);
  FreeAndNil(TobParFouLig);
  FreeAndNil(TobLigDelete);

  ListeChampCat.Free;
  ListeChampArt.Free;
  ListeChampCab.Free;
  ListeChampArtC.Free;

  ListeLibChampCat.Free;
  ListeLibChampArt.Free;
  ListeLibChampArtC.Free;
  ListeLibChampCab.Free;

  ListTypeChampCat.free;
  ListTypeChampCab.free;
  ListTypeChampArt.free;
  ListTypeChampArtC.free;

  TOBListeChamps.free;
  TOBListeChamps := nil;

  if IsInside(Self) then Action := caFree ;

end;

procedure TParamImportTarFou.FormCloseQuery(Sender: TObject;var CanClose: Boolean);
begin
CanClose := ControleModif;
end;

procedure TParamImportTarFou.FormCreate(Sender: TObject);
begin
  LStructure          := Tlist.create;
	TOBListeChamps      := TOB.Create ('LES CHAMPS DISPO',nil,-1);
  GChamp.RowCount     := NbRowsInit ;
{$IFDEF BTP}
  BVisionFic.Visible  := true;
{$ENDIF}
  iTableLigne         := PrefixeToNum ('GFL') ;
  //
  TobParFou           := TOB.Create ('PARFOU', Nil, -1) ;
  TobParFouLig        := TOB.Create ('', Nil, -1) ;
  ToBLigDelete        := TOB.Create ('', Nil, -1) ;
  //
  ListeChampCat       := HTStringList.Create;
  ListeChampArt       := HTStringList.Create;
  ListeChampArtC      := HTStringList.Create;
  ListeChampCab       := HTStringList.Create;

  ListeLibChampCat    := HTStringList.Create;
  ListeLibChampArt    := HTStringList.Create;
  ListeLibChampArtC   := HTStringList.Create;
  ListeLibChampCab    := HTStringList.Create;

  ListTypeChampCat    := HTStringList.create;
  ListTypeChampArt    := HTStringList.create;
  ListTypeChampArtC   := HTStringList.create;
  ListTypeChampCab    := HTStringList.create;

  LSChamp             := TList.create;
  //
  InitLesCols;
  //
end;

procedure TParamImportTarFou.FormKeyDown(Sender: TObject; var Key: Word;Shift: TShiftState);
var Arow,Acol : integer;
    Cancel, Chg : boolean;
begin
  if(Screen.ActiveControl = GChamp) then
  begin
    ARow := GChamp.Row;
    Case Key of
      VK_RETURN : Key:=VK_TAB ;
      VK_DELETE : begin
                    if Shift=[ssCtrl] then
                    begin
                        Key := 0 ;
                    end ;
                  end;
    end;
  end;
end;

function TParamImportTarFou.ControleProfileArticle : boolean;
begin
  Result := ExisteSQL ('SELECT GPF_PROFILARTICLE FROM PROFILART WHERE GPF_PROFILARTICLE="'+GRF_PROFILARTICLE.Text+'"');
end;

procedure TParamImportTarFou.GRF_PROFILARTICLEExit(Sender: TObject);
begin
  if (GRF_PROFILARTICLE.text <> '') and (ControleProfileArticle) then PPARTICLE.Enabled := false
                                                                 else PPARTICLE.Enabled := true;
end;

procedure TParamImportTarFou.AfficheLigneListChamps (Ligne: integer; TOBL : TOB);
begin
	TOBL.PutLigneGrid  (GListeChamps,ligne,false,false,'CHAMPS;LIBELLE');
end;

procedure TParamImportTarFou.AfficheGridListChamps;
var Indice : integer;
begin
  GListeChamps.RowCount := TOBListeChamps.detail.count+1;
  for Indice := 0 to TOBListeChamps.detail.count -1 do
  begin
  	AfficheLigneListChamps (Indice+1,TOBListeChamps.detail[indice]);
  end;
  HMTrad.ResizeGridColumns (GListeChamps);
end;

procedure TParamImportTarFou.DropUsed;
var TOBL,TOBLC : TOB;
		indice : integer;
begin
	for Indice :=0 to TOBParFouLig.detail.count -1 do
  begin
  	TOBL := TOBParFouLig.detail[indice];
    TOBLC := TOBListeChamps.findfirst(['CHAMPS'],[TOBL.getValue('GFL_CHAMP')],true);
    if TOBLC <> nil then TOBLC.free;
  end;
end;

procedure TParamImportTarFou.FormShow(Sender: TObject);
var NomList,Before,After : string;
		cancel,chg : boolean;
    Acol,Arow : integer;
begin
  TGFL_CHAMP.visible := false;
  lesColChamp := 'SELECT;GFL_CHAMP;LIBELLE;GFL_OFFSET;GFL_LONGUEUR;GFL_SEPDECIMALE;GFL_NBRDECIMALE;GFL_FICHIER;';
  //
  EtudieColsListe ;
  iWidthsDebut      := GChamp.DefaultColWidth;
  iWidthsLongueur   := GChamp.DefaultColWidth;
  AffecteGrid (GChamp, Action) ;
  InitialiseEntete;
  //
  ChargeParametrage;
  //
  DropUsed;
  //
  AfficheGridListChamps;
  ActiveEvents (false);
  bModif := False;
  //
  GChamp.CacheEdit;
  ActiveEvents (true);
  GChamp.row := 1; GChamp.col := SAF_Libelle;
	GChampEnter (Self);
  GChamp.ShowEditor;
end;

procedure TParamImportTarFou.ActiveEvents (status : boolean);
begin
	if status then
  begin
  	GChamp.PostDrawCell := PostDrawCell;
  	GChamp.OnCellEnter := GChampCellEnter ;
  	GChamp.OnCellExit := GChampCellExit ;
  	GChamp.OnRowEnter := GChampRowEnter;
  	GChamp.OnRowExit  := GChampRowExit;
  end else
  begin
  	GChamp.PostDrawCell := nil;
  	GChamp.OnCellEnter := nil ;
  	GChamp.OnCellExit := nil ;
  	GChamp.OnRowEnter := nil;
  	GChamp.OnRowExit  := nil;
  end;
end;

procedure TParamImportTarFou.FormResize(Sender: TObject);
begin
HMTrad.ResizeGridColumns (GChamp);
end;

procedure TParamImportTarFou.CheckProfilClick(Sender: TObject);
begin
  if CheckProfil.Checked then
  begin
    TGRF_PROFILARTICLE.Enabled := true;
    GRF_PROFILARTICLE.Enabled := true;
    GRF_FAMILLETAXE1.Text := '';
    GRF_COMPTAARTICLE.Text := '';
    GRF_FAMILLETAXE1.Value  := '';
    GRF_COMPTAARTICLE.value := '';
    GRF_TENUESTOCK.Checked  := false;
    GRF_ESCOMPTABLE.Checked  := false;
    PPARTICLE.Enabled := false;
  end else
  begin
    PPARTICLE.Enabled := true;
    TGRF_PROFILARTICLE.Enabled := false;
    GRF_PROFILARTICLE.Enabled := false;
  end;
end;

procedure TParamImportTarFou.BvisionFicClick(Sender: TObject);
var TOBL : TOB;
begin
  if BVisionFic.Down then TTFICHIERASSOC.Visible := true
                     else TTFICHIERASSOC.Visible := false;
  TOBL := GetTOBLigne (GChamp.row);  if TOBL = nil then exit;
  TOBL.putValue('GFL_FICHIER',extractfileName(trim(GFL_FICHIER.Text)));
  GChamp.InvalidateCell (SAF_FichierAssocie,GChamp.row);
end;

procedure TParamImportTarFou.TTFICHIERASSOCClose(Sender: TObject);
var TOBL : TOB;
begin
  TOBL := GetTOBLigne (GChamp.row);  if TOBL = nil then exit;
  TOBL.putValue('GFL_FICHIER',extractfileName(trim(GFL_FICHIER.Text)));
  GChamp.InvalidateCell (SAF_FichierAssocie,GChamp.row);
  BVisionFic.Down := false;
end;

procedure TParamImportTarFou.GFL_FICHIERExit(Sender: TObject);
var TOBL : TOB;
begin
  TOBL := GetTOBLigne (GChamp.row);  if TOBL = nil then exit;
  TOBL.putValue('GFL_FICHIER',extractfileName(trim(GFL_FICHIER.Text)));
  GChamp.InvalidateCell (SAF_FichierAssocie,GChamp.row);
end;

procedure TParamImportTarFou.GFL_FICHIERElipsisClick(Sender: TObject);
var TOBL : TOB;
begin
  if ODFichier.Execute then
  begin
    GFL_FICHIER.text := extractfileName(ODFichier.FileName);
    TOBL := GetTOBLigne (GChamp.row);  if TOBL = nil then exit;
    TOBL.putValue('GFL_FICHIER',GFL_FICHIER.Text);
    GChamp.InvalidateCell (SAF_FichierAssocie,GChamp.row);
  end;
end;

procedure TParamImportTarFou.PostDrawCell(ACol, ARow: Integer;Canvas: TCanvas; AState: TGridDrawState);
Var ARect : TRect ;
    TOBL : TOB;
begin
  if (Arow = 0) or (Acol=0) then exit;
  TOBL := GetTOBLigne (Arow); if TOBL = nil then exit;
{$IFDEF BTP}
  if (ACol = SAF_FichierAssocie) and (GChamp.ColWidths [SAF_FichierAssocie] > 0) then
  begin
    ARect:=GChamp.CellRect(ACol,ARow) ;
    Canvas.FillRect (ARect);
    if (TypeChamps(TOBL.GetValue('GFL_CHAMP'))='COMBO') then
    begin
      IList.DrawingStyle := dsTransparent;
      if TOBL.GetValue('GFL_FICHIER') <> '' then IList.Draw (CanVas,ARect.left,ARect.top,1)
                                            else IList.Draw (CanVas,ARect.left,ARect.top,0);
    end;
  end;
{$ENDIF}
end;

procedure TParamImportTarFou.GRF_TYPEARTICLEChange(Sender: TObject);
begin
  bModif := True;
  if GRF_TYPEARTICLE.Value = 'ARP' then
  begin
    PRESTATION.Visible := True;
    TPrestation.Visible := true;
  end else
  begin
    TPrestation.Visible := false;
    PRESTATION.Visible := false;
  end;
end;

procedure TParamImportTarFou.PRESTATIONChange(Sender: TObject);
begin
  if Prestation.text = '' then BEGIN GRF_PRESTATION.text := ''; Exit END;
  GRF_PRESTATION.Text :=  CodeArticleUnique2 (PRESTATION.text,'');
end;

procedure TParamImportTarFou.GRF_PARFOUEnter(Sender: TObject);
begin
  LastPar := GRF_PARFOU.Text;
end;

procedure TParamImportTarFou.GChampKeyDown(Sender: TObject; var Key: Word;  Shift: TShiftState);
var OkG, Vide: Boolean;
begin
  OkG := (Screen.ActiveControl = GChamp);
  Vide := (Shift = []);
	case key of
    VK_DELETE: if ((OkG) and (Shift = [ssCtrl])) then
      begin
        Key := 0;
        BFlecheGaucheClick(Sender);
      end;
  end;
end;

procedure TParamImportTarFou.GRF_MULTIFOUClick(Sender: TObject);
begin
  if GRF_MULTIFOU.Checked then GRF_TIERS.Enabled := false
  												else GRF_TIERS.Enabled := True;
  InitialiseListBox;
  ChargeParametrage;
  DropUsed;
  //
  AfficheGridListChamps;
end;

end.
