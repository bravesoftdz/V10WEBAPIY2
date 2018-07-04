{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... :   /  /
Modifié le ... :   /  /    
Description .. : Gestion d'une grille dont la première colonne est organisée
Suite ........ : en arbre hierarchique (treeview) avec la possibilité de
Suite ........ : développer ou de réduire des branches de l'arbre.
Mots clefs ... : TREEVIEW;GRID
*****************************************************************}
{
PT1 : 19/10/2007 GGU V_80 Fonction d'impression
PT2 : 19/12/2007 GGU V_80 Compatibilité Delphi 5 / Version 7 pour reprise de l'outils en V7
}
unit PGTreeTobFrame;

interface

uses
{$IFNDEF VER130}   //PT2
  Variants,
{$ENDIF}
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms,
  Dialogs, Vierge, uTob, HSysMenu, HTB97, Menus, Grids, Hctrls, ExtCtrls,
  HPanel, uTom, ImgList, ComCtrls, StdCtrls, wCommuns;

type
  { Types de donnée }
{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... :   /  /
Modifié le ... :   /  /
Description .. : Types d'images possibles pour l'affichage de l'arbre
Suite ........ : tttiCBG  : Fin de branche                '-
Suite ........ : tttiTEG  : Elément sur branche branche   |-
Suite ........ : tttiVER  : Branche sans élément          |
Suite ........ : tttiNone : Pas de branche à ce niveau
Mots clefs ... :
*****************************************************************}
  TTreeTobTvImg = (tttiCBG, tttiTEG, tttiVER, tttiNone);
  TArrayFieldAndValue = Array of Array [0..1] of Variant;

  { Types classes }
  TTreeGridCheckBox = class(TCustomCheckBox)
  private
    FGrid: THGrid;
    procedure WMKeyDown(var Message: TWMKeyDown); message CN_KEYDOWN;
  end;

  TTreeTom = class(Tom); //@@GGU
  TTreeManager = class;

{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... :   /  /
Modifié le ... :   /  /
Description .. : Objet premettant de lier les actions prédéfinies aux bouttons
Mots clefs ... :
*****************************************************************}
  TTreeTobLinkedCtrls = class
  public
    BtAddRow, BtDeleteRow, BtUndo,
    BtAddChildRow, BtParamListe,
    BtExpand, BtSearch, BtSelectAll,
    BtSelectNode, BtPrintGrid,
    BtLoupe: TToolBarButton97;
    PmLoupe: THPopupMenu;
  end;

{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... :   /  /
Modifié le ... :   /  /
Description .. : Objet de gestion d'un Item (une ligne)
Mots clefs ... :
*****************************************************************}
  TTreeManagerItem = class
  private
    FManager: TTreeManager;
    FExpanded, FVisible: Boolean;
    FTreeImages: Array of TTreeTobTvImg;
    FTob: Tob;
    FNodeCoord: TPoint;
    FIndex, FRowIndex: Integer;
    function GetImagesCount: Integer;
    function GetTreeImage(Index: Integer): TTreeTobTvImg;
    procedure SetTreeImage(Index: Integer; const Value: TTreeTobTvImg);
    procedure SetNodeCoord(const Value: TPoint);
    procedure SetVisible(const Value: Boolean);
    function GetRowIndex: Integer;
    function GetParentIndex: Integer;
    function GetLevel: Integer;
    function GetIsParent: Boolean;
    function GetParent: TTreeManagerItem;
    function GetFirstParentCollapsed: TTreeManagerItem;
  public
    procedure InvertExpandState;
    procedure ClearImages;
    procedure AddImage(Image: TTreeTobTvImg);
    function NextItem: TTreeManagerItem;
    function NextItemVisible: TTreeManagerItem;
    function PrevItem: TTreeManagerItem;
    function PrevItemVisible: TTreeManagerItem;
    property Expanded: Boolean read FExpanded write FExpanded;
    property RowIndex: Integer read GetRowIndex;
    property NodeCoord: TPoint read FNodeCoord write SetNodeCoord;
    property TreeImages[Index: Integer]: TTreeTobTvImg read GetTreeImage write SetTreeImage;
    property TobLine: Tob read FTob write FTob;
    property CountImages: Integer read GetImagesCount;
    property Visible: Boolean read FVisible write SetVisible;
    property Index: Integer read FIndex;
    property ParentIndex: Integer read GetParentIndex;
    property Parent: TTreeManagerItem read GetParent;
    property Level: Integer read GetLevel;
    property IsParent: Boolean read GetIsParent;
    property FirstParentCollapsed: TTreeManagerItem read GetFirstParentCollapsed;
    constructor Create(Parent: TTreeManager; const ATob: Tob; const ItemIndex: Integer = -1);
    destructor Destroy; override;
  end;

{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... :   /  /
Modifié le ... :   /  /
Description .. : Objet de gestion de l'arbre
Mots clefs ... :
*****************************************************************}
  TTreeManager = class
  private
    FItems: TList;
    FTob: Tob;
    FRowOffSet, FLastColFound, FLastIndex: Integer;
    FShowRoot, FFirstFind: Boolean;
    function GetTreeManagerItem(Index: Integer): TTreeManagerItem;
    function GetCountItems: Integer;
  public
    function IndexOf(T: Tob): TTreeManagerItem;
    function Find(const St: String; const Options: TFindOptions;
                  const WithAlert: Boolean; out FoundFieldName: String): TTreeManagerItem;
    procedure InitTreeManager;
    procedure Clear;
    procedure Add(const T: Tob);
    procedure Insert(const ItemIndex: Integer; const T: Tob);
    procedure Delete(const ItemIndex: Integer);
    property TobData: Tob read FTob write FTob;
    property Items[Index: Integer]: TTreeManagerItem read GetTreeManagerItem; default;
    property ShowRoot: Boolean read FShowRoot write FShowRoot;
    property CountItems: Integer read GetCountItems;
    constructor Create(const ATob: Tob; const RowOffSet: Integer; const AShowRoot: Boolean);
    destructor Destroy; override;
  end;

  { Type événements }
  TTreeGridEvent = procedure of object;
  TOnMnuGetCtx = procedure(var Ctx: String) of object;
  TOnElipsisGetPlus = procedure(const FieldName: String; var Plus: String) of object;
  TOnControlField = procedure(const FieldName: String) of object;
  TOnControlRecord = procedure(var LastErrorField: String) of object;
  TTreeGridCanDoEvent = procedure(var Cancel: Boolean) of object;

{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... :   /  /
Modifié le ... :   /  /
Description .. : Fiche d'affichage de l'arborescence sous forme de grille
suite ........ : avec un treeView dans la première colonne
Mots clefs ... :
*****************************************************************}
  TFFrameTreeTob = class(TFrame)
    PmLoupe: THPopupMenu;
    MnSavShowTob: THMenuItem;
    FImageList: TImageList;
    TreeGrid: THGrid;
    HMTrad: THSystemMenu;
    MnSavSepare: THMenuItem;
    PCumul: THPanel;
    PnCbBooleanField: THPanel;
    MnSavShowTob2: THMenuItem;
    { Evénements Grille }
    procedure TreeGridExit(Sender: TObject);
    procedure TreeGridMouseMove(Sender: TObject; Shift: TShiftState; X, Y: Integer);
    procedure TreeGridRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
    procedure TreeGridRowExit(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
    procedure TreeGridCellEnter(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
    procedure TreeGridCellExit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
    procedure TreeGridMouseDown(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
    procedure TreeGridElipsisClick(Sender: TObject);
    procedure TreeGridKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure TreeGridFlipSelection(Sender: TObject);
    { Evénements divers }
    procedure PmLoupePopup(Sender: TObject);
    procedure MnSavShowTobClick(Sender: TObject);
    procedure FindDialogFind(Sender: TObject);
    procedure CbBooleanFieldClick(Sender: TObject);
    procedure FrameResize(Sender: TObject);
    procedure MnSavShowTob2Click(Sender: TObject);
  private
    FForm                 : TForm;
    FTreeManager          : TTreeManager;
    FindDialog            : TFindDialog;
    FCurrentRow, FCurrentCol, FDoControl, FDoRefresh, FDoUpdate,
    FLastError : Integer;
    FShowRoot, FInitialized, FTomEnabled, FConsultation : Boolean;
    FDBListe              : String;
    FTreeTob, FTobFields, FTreeTobSaved, FTob2Delete, FTreeTobBeforeModif,
    FCurrentProcessedTob, FSysTob, FTobCaptions : Tob;
    FTom                  : TTreeTom;
    FListe                : TWListe;
    FPoint, FCellFound    : TPoint;
    FRatedFields, FEditableFields : TStrings;
    FOnNodeMouseEnter, FOnNodeMouseClick, FOnExpandNode, FOnAfterDeleteRecord,
    FOnNewRecord, FOnBeforeRefresh, FOnCollapseNode : TTreeGridEvent;
    FOnMnuLoupeGetCtx     : TOnMnuGetCtx;
    FOnElipsisGetPlus     : TOnElipsisGetPlus;
    FOnCalculField,
    FOnControlField       : TOnControlField;
    FOnControlRecord      : TOnControlRecord;
    FBtParamListe, FBtAddChildRow, FBtDeleteRow, FBtUndo, FBtSearch, FBtAddRow,
    FBtSelectAll, FBtPrintGrid, FBtSelectNode, FBtExpand, FBtLoupe : TToolBarButton97;
    FLinkedCtrls          : TTreeTobLinkedCtrls;
    FUneditableConditions : TArrayFieldAndValue;
    FOnBeforeDeleteRecord : TTreeGridCanDoEvent;
    FCbBooleanField       : TTreeGridCheckBox;
    FNotifyEvent          : TNotifyEvent;
    FDisplayTree          : Boolean;
    { Déclarations privées }
    procedure Form_Resize(Sender: TObject);
    function Form_CountFrameTreeTob: Integer;
    { Relatifs grille }
    procedure FillTreeGrid;
    procedure Grid_OnDrawCell(ACol, ARow: Integer; Canvas: TCanvas; AState: TGridDrawState);
    procedure Grid_OnColumnWidthsChanged(Sender: TObject);
    function Grid_GetLineWidth: Integer;
    function Grid_GetLineHeight: Integer;
    procedure Grid_InvertNodeState(const iRow: Integer; const Recurse: Boolean; const Visibility: Boolean);
    function Grid_FixedCols: Integer;
    function Grid_GetXOffSet: Integer;
    function Grid_CanClickNode(const X, Y: Integer): Boolean;
    function Grid_SetFocus(const iCol, iRow: Integer): Boolean;
    procedure Grid_DeleteRow;
    procedure Grid_InsertRow;
    procedure InitCollumns;
    procedure RefreshBtExpand;
    procedure InitPCumul;
    procedure FitCheckBoxToCell;
    function GridAddRow: Integer;
    function NodeIsExpanded(IM: TTreeManagerItem): Boolean;
    procedure BtExpandClick(Sender: TObject);
    procedure BtParamListeClick(Sender: TObject);
    procedure BtSearchClick(Sender: TObject);
    procedure BtPrintGridClick(Sender: TObject);
    procedure BtSelectAllClick(Sender: TObject);
    procedure BtSelectNodeClick(Sender: TObject);
    procedure BtDeleteRowClick(Sender: TObject);
    procedure BtUndoClick(Sender: TObject);
    procedure BtAddRowClick(Sender: TObject);
    procedure BtAddChildRowClick(Sender: TObject);
    function RecordIsValid(const ShowMessages: Boolean = True): Boolean;
    procedure UnassignNotifyEvent(Event: TNotifyEvent);
    procedure ReassignNotifyEvent(Event: TNotifyEvent);
    { Relatif TreeTob }
    procedure GetTobFields;
    function GetFirstLevel: Integer;
    procedure CreateTobsForListe;
    procedure FreeTobsForListe;
    function TobGetString(T: Tob; const FieldName: String): String;
    function TobGetInteger(T: Tob; const FieldName: String): Integer;
    function TobGetDouble(T: Tob; const FieldName: String): Double;
    function TobGetDateTime(T: Tob; const FieldName: String): TDateTime;
    function TobGetBoolean(T: Tob; const FieldName: String): Boolean;
    procedure TobDoUpdate(T: Tob; const FieldName: String; const Value: Variant);
    procedure TobSetString(T: Tob; const FieldName, Value: String);
    procedure TobSetInteger(T: Tob; const FieldName: String; const Value: Integer);
    procedure TobSetDouble(T: Tob; const FieldName: String; const Value: Double);
    procedure TobSetDateTime(T: Tob; const FieldName: String; const Value: TDateTime);
    procedure TobSetBoolean(T: Tob; const FieldName: String; const Value: Boolean);
    function TobGetNiveau(T: Tob): Integer;
    procedure DoControlField(T: Tob; const FieldName: String; const ShowMessages: Boolean = True);
    procedure DoControlRecord(var LastErrorFieldName: String; const ShowMessages: Boolean = True);
    procedure SaveCell2Tob(const ACol, ARow: Integer; T: Tob; const ControlEgalite: Boolean = True);
    procedure DeleteTob(T: Tob);
    procedure AddIkc(T: Tob);
    procedure InitCellFound;
    procedure FillTreeGridLine(IM: TTreeManagerItem; const iRow: Integer);
    { Getter / Setter }
    { Getter }
    function GetExpanded: Boolean;
    function GetIsParentNode: Boolean;
    function GetCurrentTob: Tob;
    function GetCurrentItemManager: TTreeManagerItem;
    function GetCurrentFieldName: String;
    function GetTableName: String;
    function GetCountDetails(iRow: Integer): Integer;
    function GetRowEditable(iRow: Integer): Boolean;
    function GetTobRow(iRow: Integer): Tob;
    function GetRowTob(ATob: Tob): Integer;
    function GetItemManagerRow(iRow: Integer): TTreeManagerItem;
    function GetFieldCol(iCol: Integer): String;
    function GetColField(FieldName: String): Integer;
    function GetAbsoluteParentIndex(T: Tob): Integer;
    function GetExpandedRow(iRow: Integer): Boolean;
    function GetIkc: Char;
    function GetWhere: String;
    function GetMultiSelected: Boolean;
    function GetFirstRow: Integer;
    function GetLastRow: Integer;
    function GetRowSelected(iRow: Integer): Boolean;
    function GetEditableCol(iCol: Integer): Boolean;
    function GetSelectionEnabled: Boolean;
    function GetColVisible(iCol: Integer): Boolean;
    { Setter }
    procedure SetConsultation(const Value: Boolean);
    procedure SetTreeTob(const Value: Tob);
    procedure SetShowRoot(const Value: Boolean);
    procedure SetDBListe(const Value: String);
    procedure SetBtAddRow(const Value: TToolBarButton97);
    procedure SetBtAddChildRow(const Value: TToolBarButton97);
    procedure SetBtDeleteRow(const Value: TToolBarButton97);
    procedure SetBtUndo(const Value: TToolBarButton97);
    procedure SetBtExpand(const Value: TToolBarButton97);
    procedure SetPmLoupe(const Value: THPopupMenu);
    procedure SetBtLoupe(const Value: TToolBarButton97);
    procedure SetBtParamListe(const Value: TToolBarButton97);
    procedure SetBtSearch(const Value: TToolBarButton97);
    procedure SetBtPrintGrid(const Value: TToolBarButton97);
    procedure SetBtSelectAll(const Value: TToolBarButton97);
    procedure SetBtSelectNode(const Value: TToolBarButton97);
    procedure SetLinkedCtrls(const Value: TTreeTobLinkedCtrls);
    procedure SetColVisible(iCol: Integer; const Value: Boolean);
  public
    { Déclarations publiques }
    { Affichage }
    procedure UpdateCumulField(const FieldName: String);
    procedure UpdateCumuls(const iRow: Integer; const ForceCalcul: Boolean = False);
    procedure UpdateTotal;
    procedure DisplayTreeTob;
    procedure RefreshDBListe;
    procedure RefreshTreeTob;
    procedure ResizeGrid;
    procedure RefreshTreeTobLigne(const iRow: Integer);
    function GridSetFocusCol(const FieldName: String): Boolean;
    procedure CollapseNode(const Recurse: Boolean = False);
    procedure ExpandNode(const Recurse: Boolean = False);
    procedure AddCol(const ColName, ColTitle: String;
                     const ColWidth: Integer; const ColType: Char; const ColFormat: String;
                     const ColAlign: TAlignment);
    { Divers }
    procedure DisableTom;
    procedure EnableTom;
    function IsEditableField(const FieldName: String): Boolean;
    procedure DeleteRow;
    procedure AddRow(const Child: Boolean);
    procedure Undo;
    procedure BeginUpdate;
    procedure EndUpdate;
    procedure DisableControl;
    procedure EnableControl;
    procedure DisableRefresh;
    procedure EnableRefresh;
    procedure Unselect(const iRow: Integer = -1);
    function isEmpty: Boolean;
    procedure Debug;
    procedure PrintGrid; //PT1
    { Donnée }
    procedure Cancel;
    function Valid(const DoTransactions: Boolean): Boolean;
    { Get }
    function GetString(const FieldName: String): String;
    function GetInteger(const FieldName: String): Integer;
    function GetDouble(const FieldName: String): Double;
    function GetDateTime(const FieldName: String): TDateTime;
    function GetBoolean(const FieldName: String): Boolean;
    function GetOldString(const FieldName: String): String;
    function GetOldInteger(const FieldName: String): Integer;
    function GetOldDouble(const FieldName: String): Double;
    function GetOldDateTime(const FieldName: String): TDateTime;
    function GetOldBoolean(const FieldName: String): Boolean;
    function GetParentString(const FieldName: String): String;
    function GetParentInteger(const FieldName: String): Integer;
    function GetParentDouble(const FieldName: String): Double;
    function GetParentDateTime(const FieldName: String): TDateTime;
    function GetParentBoolean(const FieldName: String): Boolean;
    function GetChildString(const Index: Integer; const FieldName: String): String;
    function GetChildInteger(const Index: Integer; const FieldName: String): Integer;
    function GetChildDouble(const Index: Integer; const FieldName: String): Double;
    function GetChildDateTime(const Index: Integer; const FieldName: String): TDateTime;
    function GetChildBoolean(const Index: Integer; const FieldName: String): Boolean;
    { Set }
    procedure SetString(const FieldName, Value: String);
    procedure SetInteger(const FieldName: String; Value: Integer);
    procedure SetDouble(const FieldName: String; Value: Double);
    procedure SetDateTime(const FieldName: String; Value: TDateTime);
    procedure SetBoolean(const FieldName: String; Value: Boolean);
    procedure SetParentString(const FieldName, Value: String);
    procedure SetParentInteger(const FieldName: String; Value: Integer);
    procedure SetParentDouble(const FieldName: String; Value: Double);
    procedure SetParentDateTime(const FieldName: String; Value: TDateTime);
    procedure SetParentBoolean(const FieldName: String; Value: Boolean);
    procedure SetChildString(const Index: Integer; const FieldName, Value: String);
    procedure SetChildInteger(const Index: Integer; const FieldName: String; Value: Integer);
    procedure SetChildDouble(const Index: Integer; const FieldName: String; Value: Double);
    procedure SetChildDateTime(const Index: Integer; const FieldName: String; Value: TDateTime);
    procedure SetChildBoolean(const Index: Integer; const FieldName: String; Value: Boolean);
    procedure SetChildrenString(const FieldName, Value: String);
    procedure SetChildrenInteger(const FieldName: String; Value: Integer);
    procedure SetChildrenDouble(const FieldName: String; Value: Double);
    procedure SetChildrenDateTime(const FieldName: String; Value: TDateTime);
    procedure SetChildrenBoolean(const FieldName: String; Value: Boolean);
    procedure SetDisplayTree(const Value: Boolean);
    { Propriétés }
    property TreeTob: Tob read FTreeTob write SetTreeTob;
    property NodeExpanded: Boolean read GetExpanded;
    property Expanded[iRow: Integer]: Boolean read GetExpandedRow;
    property IsParentNode: Boolean read GetIsParentNode;
    property CurrentTob: Tob read GetCurrentTob;
    property CurrentItemManager: TTreeManagerItem read GetCurrentItemManager;
    property CurrentRow: Integer read FCurrentRow;
    property CurrentCol: Integer read FCurrentCol;
    property CurrentField: String read GetCurrentFieldName;
    property EditableFields: TStrings read FEditableFields;
    property TableName: String read GetTableName;
    property LastError: Integer read FLastError write FLastError;
    property Initialized: Boolean read FInitialized;
    property CountDetails[iRow: Integer]: Integer read GetCountDetails;
    property UneditableConditions: TArrayFieldAndValue read FUneditableConditions write FUneditableConditions;
    property RowEditable[iRow: Integer]: Boolean read GetRowEditable;
    property TobRow[iRow: Integer]: Tob read GetTobRow;
    property RowTob[ATob: Tob]: Integer read GetRowTob;
    property ItemManagerRow[iRow: Integer]: TTreeManagerItem read GetItemManagerRow;
    property FieldCol[iCol: Integer]: String read GetFieldCol;
    property ColField[FieldName: String]: Integer read GetColField;
    property AbsoluteParentIndex[T: Tob]: Integer read GetAbsoluteParentIndex;
    property TreeTom: TTreeTom read FTom;
    property Ikc: Char read GetIkc;
    property Where: String read GetWhere;
    property MultiSelected: Boolean read GetMultiSelected;
    property RatedFields: TStrings read FRatedFields;
    property FirstRow: Integer read GetFirstRow;
    property LastRow: Integer read GetLastRow;
    property Selected[iRow: Integer]: Boolean read GetRowSelected;
    property ColEditable[iCol: Integer]: Boolean read GetEditableCol;
    property ColVisible[iCol: Integer]: Boolean read GetColVisible write SetColVisible;
    property SelectionEnabled: Boolean read GetSelectionEnabled;
    property DisplayTree: Boolean read FDisplayTree write SetDisplayTree;
//    property GestionCache => TODO !
    { constructeur / destructeur }
    constructor Create(F: TForm; C: TWinControl; ATreeTob: Tob; const ADBListe: String;
                       ALinkedCtrls: TTreeTobLinkedCtrls; const AConsultation: Boolean = True;
                       const AEditableFields: TStrings = nil; const AShowRoot: Boolean = True;
                       const ADisplayTree: Boolean = True); reintroduce;
    destructor Destroy; override;
  published
    property Consultation: Boolean read FConsultation write SetConsultation default True;
    property DBListe: String read FDBListe write SetDBListe;
    property ShowRoot: Boolean read FShowRoot write SetShowRoot default True;
    property OnNodeMouseEnter: TTreeGridEvent read FOnNodeMouseEnter write FOnNodeMouseEnter;
    property OnNodeMouseClick: TTreeGridEvent read FOnNodeMouseClick write FOnNodeMouseClick;
    property OnCollapseNode: TTreeGridEvent read FOnCollapseNode write FOnCollapseNode;
    property OnExpandNode: TTreeGridEvent read FOnExpandNode write FOnExpandNode;
    property OnBeforeRefresh: TTreeGridEvent read FOnBeforeRefresh write FOnBeforeRefresh;
    property OnMnuLoupeGetCtx: TOnMnuGetCtx read FOnMnuLoupeGetCtx write FOnMnuLoupeGetCtx;
    property OnElipsisGetPlus: TOnElipsisGetPlus read FOnElipsisGetPlus write FOnElipsisGetPlus;
    property OnControlField: TOnControlField read FOnControlField write FOnControlField;
    property OnCalculField: TOnControlField read FOnCalculField write FOnCalculField;
    property OnControlRecord: TOnControlRecord read FOnControlRecord write FOnControlRecord;
    property OnBeforeDeleteRecord: TTreeGridCanDoEvent read FOnBeforeDeleteRecord write FOnBeforeDeleteRecord;
    property OnAfterDeleteRecord: TTreeGridEvent read FOnAfterDeleteRecord write FOnAfterDeleteRecord;
    property OnNewRecord: TTreeGridEvent read FOnNewRecord write FOnNewRecord;
    property BtExpand: TToolBarButton97 read FBtExpand write SetBtExpand;
    property BtSearch: TToolBarButton97 read FBtSearch write SetBtSearch;
    property BtPrintGrid: TToolBarButton97 read FBtPrintGrid write SetBtPrintGrid;
    property BtSelectAll: TToolBarButton97 read FBtSelectAll write SetBtSelectAll;
    property BtSelectNode: TToolBarButton97 read FBtSelectNode write SetBtSelectNode;
    property BtLoupe: TToolBarButton97 read FBtLoupe write SetBtLoupe;
    property BtAddRow: TToolBarButton97 read FBtAddRow write SetBtAddRow;
    property BtAddChildRow: TToolBarButton97 read FBtAddChildRow write SetBtAddChildRow;
    property BtDeleteRow: TToolBarButton97 read FBtDeleteRow write SetBtDeleteRow;
    property BtUndo: TToolBarButton97 read FBtUndo write SetBtUndo;
    property BtParamListe: TToolBarButton97 read FBtParamListe write SetBtParamListe;
    property PopupLoupe: THPopupMenu read PmLoupe write SetPmLoupe;
    property LinkedCtrls: TTreeTobLinkedCtrls read FLinkedCtrls write SetLinkedCtrls;
  end;

  TTreeTransac = class
  private
    FTT: TFFrameTreeTob;
    procedure DoTreeTrans(const T: Tob);
  public
    constructor Create(const AFrameTreeTob: TFFrameTreeTob);
    procedure InsertOrUpdate;
  end;

function FindTextInTob(FromFieldIndex: Integer; const S: String; T: Tob; const Options: TFindOptions; out FoundFieldName: String): Boolean;

var
  FFrameTreeTob: TFFrameTreeTob;

const
  ErrTomControlField = 22222;
  ChampsSupList = '';

implementation

{$R *.DFM}

uses
  hMsgBox
  , UIUtil
  , HEnt1
  , M3FP
  , Math
  , ParamDbg
  , wMnu
  , uTobDebug
  , ParamDat
  , LookUp
  , HDrawXP
  , wRapport
{$IFNDEF EAGLCLIENT}
  , PrintDBG //GGU PT1
{$else}
  , UtileAgl //GGU PT1
{$ENDIF}
{$IFNDEF VER130} //PT2
  , StrUtils
{$ENDIF}
  ;

{$IFDEF VER130} //PT2
function RightStr(const AText: AnsiString; const ACount: Integer): AnsiString;
var
  Index : Integer;
begin
  Index := Length(AText) - ACount;
  result := Copy(AText, Index, ACount);
end;
{$ENDIF}


{ Erreur Mode SAV uniquement --------------------------------------------------- }
procedure SavError(const ErrorMsg, ErrorTitle: String);
begin
  if IsDayPass then
    PgiError(ErrorMsg, ErrorTitle)
end;

{ TFTreeTob }

{ Passage en mode consultation ou edition -------------------------------------- }
procedure TFFrameTreeTob.SetConsultation(const Value: Boolean);
begin
  FConsultation := Value;
  { Phase d'init permettant la modification : }
  { Affichage de la grille : Edition, Sélection à la cellule }
  if FConsultation then
    TreeGrid.Options := TreeGrid.Options + [goRowSelect] - [goEditing, goAlwaysShowEditor]
  else
    TreeGrid.Options := TreeGrid.Options - [goRowSelect] + [goEditing, goAlwaysShowEditor];
  { Visibilité des controls d'édition }
  if Assigned(FBtAddChildRow) then
    FBtAddChildRow.Visible := not FConsultation;
  if Assigned(FBtAddRow) then
    FBtAddRow.Visible := not FConsultation;
  if Assigned(FBtDeleteRow) then
    FBtDeleteRow.Visible := True;
  if Assigned(FBtUndo) then
    FBtUndo.Visible := not FConsultation;
  { Visibilité des controls de sélection }
  if Assigned(FBtSelectAll) then
    FBtSelectAll.Visible := FConsultation;
  if Assigned(FBtSelectNode) then
    FBtSelectNode.Visible := FConsultation;
  if FInitialized then
    TreeGrid.Refresh;
  { Ikc }
  AddIkc(FTreeTob);
end;

procedure TFFrameTreeTob.FillTreeGridLine(IM: TTreeManagerItem; const iRow: Integer);
var
  i: Integer;
  TT, sField: String;
  wField: TWFieldListe;
  T: Tob;
  function GetFieldsList: String;
  var
    i: Integer;
  begin
    Result := '';
    for i := Grid_FixedCols to Pred(TreeGrid.ColCount) do
      Result := Result + TreeGrid.ColNames[i] + ';'
  end;
begin
  T := IM.TobLine;
  if iRow >= TreeGrid.FixedRows then
  begin
    TreeGrid.Objects[0, iRow] := IM;
    { Balayage des colonnes }
    for i := Grid_FixedCols to Pred(TreeGrid.ColCount) do
    begin
      wField := FListe.ChampsByName[FTobFields.GetString('FIELD_' + IntToStr(i))];
      if Assigned(wField) then
      begin
        sField := wField.Name;
        if (Pos('(', sField) = 1) and (Pos(')', sField) = Length(sField)) then
        begin
          Delete(sField, 1, 1);
          Delete(sField, Length(sField), 1);
        end;
        if T.FieldExists(wField.Name) then
        begin
          { Affichage de l'image }
          if (wGetSimpleTypeField(sField) = 'C') and wField.Image then
          begin
            TT := ChampToTT(sField);
            if TT <> '' then
              TreeGrid.CellValues[i, iRow] := '#ICO#' + IntToStr(GetNumImage(TT, TTToNum(TT), T.GetString(sField)))
          end
          { Affichage du libellé complet }
          else if not IsEditableField(sField) and (wGetSimpleTypeField(sField) = 'C') and wField.LibelleComplet then
          begin
            TT := ChampToTT(sField);
            if TT <> '' then
            begin
              if T.GetString(sField) <> '' then
                TreeGrid.CellValues[i, iRow] := RechDom(TT, T.GetString(sField), False)
              else
                TreeGrid.CellValues[i, iRow] := T.GetString(sField);
            end;
          end
          { Blanc si null }
          else if wField.EmptyIfNull and (T.GetValue(sField) = wGetInitValue(sField)) then
            TreeGrid.CellValues[i, iRow] := ''
          else
          begin
            case wGetSimpleTypeField(sField) of
              { Formattage des décimaux }
              'N': TreeGrid.CellValues[i, iRow] := FormatFloat(StringReplace(wField.Format, ' ', ',', [rfIgnoreCase, rfReplaceAll]),
                                                               T.GetDouble(wField.Name))
            else
              TreeGrid.CellValues[i, iRow] := T.GetString(wField.Name);
            end;
          end;
        end;
      end;
    end;
  end;
end;

{ Rempli la grille de la TreeTob récursivement sur cette dernière... -----------
  ...on se base sur la liste des champs affichés contenu dans la FTobFields ---- }
procedure TFFrameTreeTob.FillTreeGrid;
var
  i: Integer;
begin
  if FTreeManager.CountItems > 0 then
    FillTreeGridLine(FTreeManager[0], TreeGrid.FixedRows);
  for i := 1 to Pred(FTreeManager.CountItems) do
    if FTreeManager[i].Visible then
      FillTreeGridLine(FTreeManager[i], GridAddRow);
  { Init. de l'affichage du bouton Etendre / Réduire }
  if FDoUpdate = 0 then
    RefreshBtExpand;
  UpdateTotal
end;

{ Affichage de la tob en grille ------------------------------------------------ }
procedure TFFrameTreeTob.DisplayTreeTob;
var
  i, j, iPosRow: Integer;
begin
  if not FInitialized then
    SavError('non initialisée !', 'FrameTreeTob :')
  else
  begin
    if not Assigned(FTreeTob) then
      SavError('non assignée !', 'Propriété "TreeTob" :')
    else
    begin
      iPosRow := TreeGrid.Row;
      TreeGrid.SquizzInvisibleCells := True;
      GetTobFields;
      BeginUpdate;
      try
        { RAZ FTobFields + Grille }
        TreeGrid.VidePile(False);
        { Init. lignes }
        if TreeGrid.FixedRows = 0 then
          TreeGrid.RowHeights[0] := TreeGrid.DefaultRowHeight;
        { Init. colonnes }
        TreeGrid.ColCount := FTobFields.ChampsSup.Count + TreeGrid.FixedCols;
        { RAZ Captions }
        for i := 0 to Pred(TreeGrid.FixedRows) do
          for j := 0 to Pred(TreeGrid.ColCount - Grid_FixedCols) do
            TreeGrid.Cells[j, i] := '';
        for i := 0 to Pred(TreeGrid.RowCount - TreeGrid.FixedRows) do
          for j := 0 to Pred(Grid_FixedCols) do
            TreeGrid.Cells[j, i] := '';
        { Affichage libellé des colonnes }
        if TreeGrid.FixedRows > 0 then
        begin
          TreeGrid.Objects[Pred(TreeGrid.ColCount), Pred(TreeGrid.FixedRows)] := FTobFields;
          for i := Grid_FixedCols to (TreeGrid.ColCount - Grid_FixedCols) do
            TreeGrid.Cells[i, Pred(TreeGrid.FixedRows)] := VarToStr(FTobCaptions.GetValeur(1000 + i - Grid_FixedCols));
        end;
        { Nom & Format des colonnes }
        InitCollumns;
        { Affichage des données }
        FillTreeGrid
      finally
        EndUpdate
      end;
      { Initialisation des cumuls }
      InitPCumul;
      { Focus sur la bonne ligne }
      if wBetween(iPosRow, TreeGrid.FixedRows, Pred(TreeGrid.RowCount)) then
        TreeGrid.Row := iPosRow
      else
        TreeGrid.Row := TreeGrid.FixedRows;
      { Resize des colonnes de la grille }
      ResizeGrid();
      { En fonction des champs éditables, donne le focus à la première colonne éditable }
      if not FConsultation and RowEditable[TreeGrid.Row] then
      begin
        i := TreeGrid.FixedCols;
        while not TreeGrid.ColEditables[i] and (i < Pred(TreeGrid.ColCount)) do
          Inc(i);
        TreeGrid.Col := i;
        FCurrentCol := TreeGrid.Col;
      end;
      if not FConsultation and Assigned(TreeGrid.ValCombo) then
        TreeGrid.ValCombo.Name := 'GRIDCOMP_ValComboBox';
      { Init. de la focalisation de la grille }
      Grid_SetFocus(TreeGrid.Col, TreeGrid.Row)
    end
  end
end;

function TFFrameTreeTob.GridAddRow: Integer;
begin
  TreeGrid.RowCount := TreeGrid.RowCount + 1;
  Result := Pred(TreeGrid.RowCount)
end;

{ Permet de donner le focus à une colonne spécifique --------------------------- }
function TFFrameTreeTob.GridSetFocusCol(const FieldName: String): Boolean;
var
  i: Integer;
begin
  i := Grid_FixedCols;
  Result := False;
  while not Result and (i < TreeGrid.RowCount) do
  begin
    Result := FieldCol[i] = FieldName;
    if not Result then
      Inc(i)
  end;
  if Result then
    TreeGrid.Col := i;
end;

{ Initialise la structure du TreeView (Codes Images dans Tob) ------------------ }
procedure TFFrameTreeTob.SetTreeTob(const Value: Tob);
begin
  FTreeTob := Value;
  FTreeManager.TobData := FTreeTob;
  FTreeManager.InitTreeManager;
  if FInitialized then
    RefreshTreeTob
end;

{ Mise à jour de la DBListe au niveau de la grille ----------------------------- }
procedure TFFrameTreeTob.SetDBListe(const Value: String);
begin
  FDBListe := Value;
  if UpperCase(FDBListe) <> UpperCase(TreeGrid.ListeParam) then
    TreeGrid.ListeParam := FDBListe;
end;

{ Initialise les tob contenant la structure de la liste ------------------------ }
procedure TFFrameTreeTob.GetTobFields;
var
  i, CptFields: Integer;
begin
  { RAZ des ChampsSup  }
  FreeTobsForListe;
  CreateTobsForListe;
  if FDBListe = '' then
    SavError('Propriété "DBListe" vide !', 'TFFrameTreeTob.GetTobFields');
  { Extraction des éléments de la liste dans les tobs FTobFields & FTobCaptions }
  CptFields := Grid_FixedCols;
  for i := 0 to Pred(FListe.ChampsCount) do
  begin
    if FListe.Champs[i].Visible then
    begin
      FTobFields.AddChampSupValeur  ('FIELD_' + IntToStr(CptFields), FListe.Champs[i].Name);
      FTobCaptions.AddChampSupValeur(FListe.Champs[i].Name, FListe.Champs[i].Caption);
      Inc(CptFields)
    end
  end
end;

{ Largeur des traits de inter-lignes ------------------------------------------- }
function TFFrameTreeTob.Grid_GetLineWidth: Integer;
begin
  if goVertLine in TreeGrid.Options then
    Result := TreeGrid.GridLineWidth
  else if V_Pgi.Draw2003 then
    Result := 0
  else
    Result := TreeGrid.GridLineWidth
end;

{ Largeur des traits de inter-colonnes ----------------------------------------- }
function TFFrameTreeTob.Grid_GetLineHeight: Integer;
begin
  if goHorzLine in TreeGrid.Options then
    Result := TreeGrid.GridLineWidth
  else if V_Pgi.Draw2003 then
    Result := 0
  else
    Result := TreeGrid.GridLineWidth
end;

{ Dessine le treeview dans la première colonne champ --------------------------- }
procedure TFFrameTreeTob.Grid_OnDrawCell(ACol, ARow: Integer; Canvas: TCanvas; AState: TGridDrawState);
var
  T: Tob;
  ARect: TRect;
  i, XOffSet, xTvImg, yTvImg, iTmp: Integer;
  PenStyle: TPenStyle;
  BrushStyle: TBrushStyle;
  BrushColor, PenColor: TColor;
  IM: TTreeManagerItem;
  procedure DrawTvImg(const X, Y: Integer; Ttti: TTreeTobTvImg);
  begin
    { Type de pinceau }
    if V_Pgi.Draw2003 then
      Canvas.Pen.Style := psSolid
    else
      Canvas.Pen.Style := psDot;
    { Dessin }
    with Canvas do
    begin
      case Ttti of
        tttiCBG:
          begin
            MoveTo(X, Y);
            LineTo(X, Y + TreeGrid.DefaultRowHeight div 2);
            LineTo(X + TreeGrid.DefaultRowHeight div 2, Y + TreeGrid.DefaultRowHeight div 2);
          end;
        tttiTEG:
          begin
            DrawTvImg(X, Y, tttiVER);
            MoveTo(X, Y + TreeGrid.DefaultRowHeight div 2);
            LineTo(X + TreeGrid.DefaultRowHeight div 2, Y + TreeGrid.DefaultRowHeight div 2);
          end;
        tttiVER:
          begin
            MoveTo(X, Y);
            LineTo(X, Y + TreeGrid.DefaultRowHeight);
          end;
      end
    end
  end;
  { Ramène la position X en fonction du niveau de la tob }
  function GetX(Level: Integer): Integer;
  begin
    if FShowRoot then
      Inc(Level);

    Result := Pred(Level) * (TreeGrid.DefaultRowHeight div 2)
            + Level * (FImageList.Width div 2)
  end;
  { Ramène la position Y }
  function GetY: Integer;
  var
    i: Integer;
  begin
    Result := 0;
    for i := 0 to Pred(TreeGrid.FixedRows) do
      if TreeGrid.RowHeights[i] > -1 then
        Result := Result + TreeGrid.RowHeights[i] + iif(TreeGrid.RowHeights[i] > -1, Grid_GetLineHeight, 0);
    for i := TreeGrid.TopRow to Pred(ARow) do
      if TreeGrid.RowHeights[i] > -1 then
        Result := Result + TreeGrid.RowHeights[i] + iif(TreeGrid.RowHeights[i] > -1, Grid_GetLineHeight, 0)
  end;
  { + ou - / Std ou XP }
  function GetIndexImg: Integer;
  begin
    if IM.Expanded then
      Result := iif(V_Pgi.Draw2003, 3, 1)
    else
      Result := iif(V_Pgi.Draw2003, 2, 0)
  end;
  { Rétablit les couleurs par défaut du pinceau et de la brosse }
  procedure ResetColors;
  begin
    Canvas.Brush.Color := BrushColor;
    Canvas.Pen.Color := PenColor;
  end;
  { Permet de savoir si la ligne courante est en couleur ou non }
  function IsColoredRow(const iRow: Integer): Boolean;
  begin
    if TreeGrid.RowHeights[iRow] > 0 then
    begin
      if iRow = TreeGrid.FixedRows then
        Result := False
      else
        Result := not IsColoredRow(Pred(iRow))
    end
    else
      Result := IsColoredRow(Pred(iRow))
  end;
  { Identifie la Colonne de l'arbre }
  function IsColTree: Boolean;
  begin
    Result := FDisplayTree and (ARow >= TreeGrid.FixedRows) and (ACol = Grid_FixedCols)
  end;
  { Cellule à redessiner ? }
  function CanRedrawCell: Boolean;
  begin
    Result := TreeGrid.TwoColors and (TreeGrid.RowHeights[ARow] > 0) and (ACol >= Grid_FixedCols)
  end;
  { Ramène la position X du text à dessiner en fonction de son alignement définit dans la liste }
  function GetLeft(R: TRect): Integer;
  var
    wField: TWFieldListe;
  begin
    wField := FListe.ChampsByName[TreeGrid.ColNames[ACol]];
    if Assigned(wField) then
    begin
      case wField.Alignment of
        taCenter      : Result := R.Left + (R.Right - R.Left - Canvas.TextWidth(TreeGrid.Cells[ACol, ARow])) div 2;
        taRightJustify: Result := R.Right - Canvas.TextWidth(TreeGrid.Cells[ACol, ARow]) - 2;
      else
        Result := R.Left + 2
      end
    end
    else
      Result := R.Left + 2
  end;
  { Permet de gérer la valeur à afficher dans la cellule }
  function GetTextValue: String;
  begin
    { Booléen }
    if TreeGrid.ColTypes[ACol] = 'B' then
      Result := HCheckBoxGridSymbol(TCheckBoxStyle(ValeurI(TreeGrid.ColFormats[ACol])), TreeGrid.Cells[ACol, ARow])
    { Autres cas }
    else
      Result := TreeGrid.Cells[ACol, ARow]
  end;
  function IsRowSelected: Boolean;
  begin
    Result := (gdSelected in AState) or TreeGrid.IsSelected(ARow)
  end;
  function IsRowVisible: Boolean;
  begin
    Result := TreeGrid.RowHeights[ARow] > 0
  end;
  const
    COL_GRADXP1STATE1 = $001595EE;
    COL_GRADXP2STATE1 = $007FD9F8;
    COL_GRADXP1STATE2 = $007ED1F9;
    COL_GRADXP2STATE2 = $00DBFEFE;
begin
  if not IsRowVisible() or (FDoRefresh > 0) then
    Exit;
  BrushColor := Canvas.Brush.Color;
  PenColor := Canvas.Pen.Color;
  { S.A.V. }
  if IsDayPass and (TreeGrid.FixedRows > 0) and (TreeGrid.FixedCols > 0) and (ACol = 0) and (ARow = 1) then
  begin
    ARect := TreeGrid.CellRect(ACol, ARow - 1);
    XOffSet := (ARect.Bottom - ARect.Top) div 4;
    ARect.Top := ARect.Top + XOffSet;
    ARect.Bottom := ARect.Bottom - XOffSet;
    XOffSet := (ARect.Right - ARect.Left) div 4;
    ARect.Left := ARect.Left + XOffSet;
    ARect.Right := ARect.Right - XOffSet;
    Canvas.Brush.Color := COL_GRADXP2STATE1;
    Canvas.Pen.Color := COL_GRADXP1STATE1;
    Canvas.Rectangle(ARect);
    ResetColors
  end;
  if FConsultation and (FCellFound.x = ACol) and (FCellFound.y = ARow) then
    Canvas.Font.Style:= [fsBold, fsItalic];
  { Gestion TwoColors }
  if CanRedrawCell then
  begin
    { Cellule sélectionnée }
    if IsRowSelected then
    begin
      Canvas.Brush.Color := clHighlight;
      Canvas.Font.Color := clHighlightText;
    end
    else
    begin
      { Couleur normale ou alternée }
      if Pos('#COL#', TreeGrid.Cells[ACol, ARow]) > 0 then
        Canvas.Brush.Color := TColor(ValeurI(Copy(TreeGrid.Cells[ACol, ARow], 6, Length(TreeGrid.Cells[ACol, ARow]) - 5)))
      else if IsColoredRow(ARow) then
        Canvas.Brush.Color := TreeGrid.AlternateColor
      else
        Canvas.Brush.Color := TreeGrid.Color;
    end;
    if not IsColTree then
    begin
      ARect := TreeGrid.CellRect(ACol, ARow);
      if V_Pgi.Draw2003 then
        ARect.Right := ARect.Right - 1;
      Canvas.FillRect(ARect);
      { Icône OU Texte }
      if Pos('#ICO#', TreeGrid.Cells[ACol, ARow]) > 0 then
        V_Pgi.GraphList.Draw(Canvas, Max(ARect.Left, ARect.Left + (ARect.Right - ARect.Left - V_Pgi.GraphList.Width) div 2),
                                     Max(ARect.Top , ARect.Top  + (ARect.Bottom - ARect.Top - V_Pgi.GraphList.Height) div 2),
                             ValeurI(Copy(TreeGrid.Cells[ACol, ARow], 6, Length(TreeGrid.Cells[ACol, ARow]) - 5)) - 1)
      else
        Canvas.TextRect(ARect, GetLeft(ARect), ARect.Top + 2, GetTextValue)
    end
  end;
  if IsColTree then
  begin
    { Tob associée à la ligne de grille }
    T := TobRow[ARow];
    { Gestionnaire de données relatives Affichage/Tob }
    IM := ItemManagerRow[ARow];
    if Assigned(T) and IsRowVisible() then
    begin
      PenStyle := Canvas.Pen.Style;
      try
        { Rectangle de la cellule }
        ARect := TreeGrid.CellRect(ACol, ARow);
        if V_Pgi.Draw2003 then
          ARect.Right := ARect.Right - 1;
        Canvas.FillRect(ARect);
        XOffSet := Grid_GetXOffSet;
        { Dessin des images par niveau }
        for i := 0 to Pred(IM.CountImages) do
          if TreeGrid.GridLineWidth > 0 then
            DrawTvImg(XOffSet + GetX(i + 1), GetY, IM.TreeImages[i]);
        { Dessin du + ou - }
        xTvImg := XOffSet + iif(TobGetNiveau(T) > GetFirstLevel, GetX(TobGetNiveau(T)) - (FImageList.Width div 2), 0) + 1;
        yTvImg := GetY + ((TreeGrid.DefaultRowHeight - FImageList.Height) div 2) + 1;
        IM.NodeCoord := Point(xTvImg, yTvImg);
        if (T.Detail.Count > 0) then
          FImageList.Draw(Canvas, xTvImg, yTvImg, GetIndexImg);
        { Texte }
        ARect.Left := ARect.Left + xTvImg + FImageList.Width div 2;
        Canvas.TextRect(ARect, ARect.Left,
                        GetY + ((TreeGrid.DefaultRowHeight - Canvas.TextHeight('W')) div 2),
                        TreeGrid.CellValues[ACol, ARow])
      finally
        Canvas.Pen.Style := PenStyle
      end
    end
  end
  { Dessin des champs de progression : Type REEL & ligne visible }
  else if (FRatedFields.IndexOf(TreeGrid.ColNames[ACol]) >= 0)
  and (TreeGrid.ColTypes[ACol] = 'R') and IsRowVisible() then
  begin
    { Progression }
    ARect := TreeGrid.CellRect(ACol, ARow);
    Canvas.FillRect(ARect);
    //ARect.Left + Trunc(Random * 100) - 1;
    ARect.Right := Min(ARect.Right, ARect.Left + Trunc(TreeGrid.ColWidths[ACol] * Min(100, Valeur(TreeGrid.Cells[ACol, ARow])) / 100));
    if (ARect.Right - ARect.Left) > 0 then
    begin
      PenStyle := Canvas.Pen.Style;
      try
        if V_Pgi.Draw2003 then
        begin
          iTmp := ARect.Bottom;
          ARect.Bottom := ARect.Top + TreeGrid.RowHeights[ARow] div 2;
          DrawVGradiant(Canvas, ARect, iif(IsRowSelected, COL_GRADXP2STATE1, COL_GRADXP2STATE2),
                                       iif(IsRowSelected, COL_GRADXP1STATE1, COL_GRADXP1STATE2));
          ARect.Top := ARect.Bottom;// + 1;
          ARect.Bottom := iTmp - iif(TreeGrid.RowHeights[ARow] mod 2 = 0, 0, 1);
          DrawVGradiant(Canvas, ARect, iif(IsRowSelected, COL_GRADXP1STATE1, COL_GRADXP1STATE2),
                                       iif(IsRowSelected, COL_GRADXP2STATE1, COL_GRADXP2STATE2));
        end
        else
        begin
          Canvas.Pen.Style := psClear;
          Canvas.Brush.Color := Luminosite(TreeGrid.AlternateColor, iif(IsRowSelected, -3, -1));
          Canvas.FillRect(ARect);
        end;
      finally
        Canvas.Pen.Style := PenStyle
      end;
    end;
    { Texte }
    BrushStyle := Canvas.Brush.Style;
    try
      Canvas.Brush.Style := bsClear;
      ARect := TreeGrid.CellRect(ACol, ARow);
      Canvas.TextRect(ARect, GetLeft(ARect), ARect.Top + 2, TreeGrid.Cells[ACol, ARow] + ' %');
    finally
      Canvas.Brush.Style := BrushStyle
    end
  end
end;

{ Refresh de l'affichage pour affichage du noeud Root -------------------------- }
procedure TFFrameTreeTob.SetShowRoot(const Value: Boolean);
begin
  FShowRoot := Value;
  DisableRefresh;
  try
    RefreshTreeTob
  finally
    EnableRefresh
  end
end;

{ Premier niveau affiché selon si l'on montre le noeud root ou non ------------- }
function TFFrameTreeTob.GetFirstLevel: Integer;
begin
  Result := iif(FShowRoot, 0, 1)
end;

{ Recharge la DBListe dans la FListe ------------------------------------------- }
procedure TFFrameTreeTob.RefreshDBListe;
begin
  { Libère la liste }
  if Assigned(FListe) then
    FreeAndNil(FListe);
  { Recrée la liste }
  FListe := TWListe.Create(FDBListe, Self);
end;

{ Recharge et réaffiche la tob  ------------------------------------------------ }
procedure TFFrameTreeTob.RefreshTreeTob;
var
  Cancel: Boolean;
begin
  Cancel := False;
  if FInitialized then
  begin
    BeginUpdate;
    try
      if Assigned(FOnBeforeRefresh) and (FDoRefresh = 0) then
      begin
        EnableControls(FForm, False, True);
        try
          FOnBeforeRefresh();
          Cancel := (FTreeTob.Detail.Count = 0) and not FTreeTob.ExistDB
        finally
          EnableControls(FForm, True, True);
        end;
      end;
      { Mise à jour des données de structure }
      FTreeManager.TobData := FTreeTob;
      FTreeManager.InitTreeManager;
      { Affichage }
      DisplayTreeTob;
      { Désélectionne }
      if SelectionEnabled then
        Unselect
    finally
      EndUpdate
    end
  end;
  { Si la tob est vide, on sort de la fiche }
  if Cancel and (Form_CountFrameTreeTob = 1) then
  begin
    PgiInfo('Aucune donnée à afficher.'#13#10' La fiche va être fermée.');
    FForm.Close
  end
end;

{ Permet de rafraîchir l'affichage d'une ligne de la grille -------------------- }
procedure TFFrameTreeTob.RefreshTreeTobLigne(const iRow: Integer);
begin
  BeginUpdate;
  try
    FillTreeGridLine(ItemManagerRow[iRow], iRow);
    if SelectionEnabled then
      Unselect(iRow)
  finally
    EndUpdate
  end
end;

{ Création des objets servant à la structure ----------------------------------- }
procedure TFFrameTreeTob.CreateTobsForListe;
begin
  FTobFields := Tob.Create('_FieldsList_', nil, -1);
  FTobCaptions := Tob.Create('_CaptionsList_', nil, -1);
  FTob2Delete := Tob.Create('_Records2Delete_', nil, -1);
  FTreeTobBeforeModif := Tob.Create(TableName, nil, -1);
  FSysTob := Tob.Create('_Sys_', nil, -1);
end;

{ Libération des objets servant à la structure --------------------------------- }
procedure TFFrameTreeTob.FreeTobsForListe;
begin
  if Assigned(FSysTob) then
    FSysTob.Free;
  if Assigned(FTreeTobBeforeModif) then
    FTreeTobBeforeModif.Free;
  if Assigned(FTob2Delete) then
    FTob2Delete.Free;
  if Assigned(FTobCaptions) then
    FTobCaptions.Free;
  if Assigned(FTobFields) then
    FTobFields.Free;
end;

{ OffSet de début de dessin du treegrid à cause des FixedCols ------------------ }
function TFFrameTreeTob.Grid_GetXOffSet: Integer;
var
  i: Integer;
begin
  Result := 0;
  for i := 0 to Pred(Grid_FixedCols) do
    Result := Result + (TreeGrid.ColWidths[i] + Grid_GetLineWidth);
end;

{ Identification de la cellule concernant un noeud ----------------------------- }
procedure TFFrameTreeTob.TreeGridMouseMove(Sender: TObject; Shift: TShiftState; X, Y: Integer);
var
  iCol, iRow, i: Integer;
  T: Tob;
  IM: TTreeManagerItem;
begin
  inherited;
  if TreeGrid.Enabled then
  begin
    TreeGrid.MouseToCell(X, Y, iCol, iRow);
    TreeGrid.Cursor := crDefault;
    if (Grid_FixedCols = TreeGrid.LeftCol) and (iCol = Grid_FixedCols) and (iRow >= TreeGrid.FixedRows) then
    begin
      T := TobRow[iRow];
      if Assigned(T) and (T.Detail.Count > 0) then
      begin
        IM := ItemManagerRow[iRow];
        if IM.NodeCoord.x > 0 then // si non, on est en ouverture de la fiche et elle est en train de se dessiner
        begin
          FPoint.x := IM.NodeCoord.x;
          FPoint.y := 0;
          for i := 0 to Pred(TreeGrid.FixedRows) do
            if TreeGrid.RowHeights[i] > -1 then
              FPoint.y := FPoint.y + TreeGrid.RowHeights[i] + iif(TreeGrid.RowHeights[i] > -1, Grid_GetLineHeight, 0);
          for i := TreeGrid.TopRow to Pred(iRow) do
            if TreeGrid.RowHeights[i] > -1 then
              FPoint.y := FPoint.y + TreeGrid.RowHeights[i] + iif(TreeGrid.RowHeights[i] > -1, Grid_GetLineHeight, 0);
          FPoint.y := FPoint.y - (TreeGrid.RowHeights[iRow] - FImageList.Height) div 2;
          if Grid_CanClickNode(X, Y) then
          begin
            TreeGrid.Cursor := crHandPoint;
            if Assigned(FOnNodeMouseEnter) then
              FOnNodeMouseEnter
          end
        end
      end
    end
  end
end;

{ Ramène le nombre de colonnes fixes ------------------------------------------- }
function TFFrameTreeTob.Grid_FixedCols: Integer;
begin
  Result := TreeGrid.FixedCols;
end;

procedure TFFrameTreeTob.Grid_DeleteRow;
begin
  TreeGrid.RowCount := Pred(TreeGrid.RowCount)
end;

procedure TFFrameTreeTob.Grid_InsertRow;
begin
  TreeGrid.RowCount := Succ(TreeGrid.RowCount)
end;

{ Permet de réduire ou étendre le noeud courant -------------------------------- }
procedure TFFrameTreeTob.Grid_InvertNodeState(const iRow: Integer; const Recurse: Boolean; const Visibility: Boolean);
var
  iLevel, i: Integer;
  SubState: Boolean;
  IM: TTreeManagerItem;
  function ParentExpanded: Boolean;
    procedure _ParentExpanded(AIM: TTreeManagerItem; var AResult: Boolean);
    var
      AIMP: TTreeManagerItem;
    begin
      if AResult then
      begin
        AIMP := AIM.Parent;
        if Assigned(AIMP) and (AIMP.Level > iLevel) then
        begin
          AResult := AIMP.Expanded;
          _ParentExpanded(AIMP, AResult)
        end
      end
    end;
  begin
    Result := True;
    _ParentExpanded(IM, Result)
  end;
begin
  SourisSablier;
  try
    IM := ItemManagerRow[iRow];
    { Affiche ou Masque }
    if Assigned(IM) then
    begin
      BeginUpdate;
      try
        IM.InvertExpandState;
        SubState := IM.Expanded;
        iLevel := IM.Level;
        i := iRow + 1;
        IM := IM.NextItem;
        while Assigned(IM) and (IM.Index < FTreeManager.CountItems) and (IM.Level > iLevel) do
        begin
          if IM.IsParent and Recurse then
            IM.Expanded := SubState;
          if not Visibility then
          begin
            if IM.Visible then
            begin
              Grid_DeleteRow();
              IM.Visible := False;
              Dec(IM.FRowIndex)
            end
          end
          else
          begin
            if ParentExpanded() then
            begin
              Grid_InsertRow();
              IM.Visible := not Assigned(IM.Parent) or IM.Parent.Expanded;
              Inc(IM.FRowIndex);
              FillTreeGridLine(IM, i);
              Inc(i)
            end
          end;
          IM := IM.NextItem
        end;
        { Mise à jour de l'affichage }
        IM := ItemManagerRow[iRow];
        for i := Succ(iRow) to Pred(TreeGrid.RowCount) do
        begin
          IM := IM.NextItemVisible;
          FillTreeGridLine(IM, i)
        end
      finally
        EndUpdate
      end
    end;
    if Visibility and Assigned(FOnExpandNode) then
      FOnExpandNode
    else if not Visibility and Assigned(FOnCollapseNode) then
      FOnCollapseNode;
    { Gestion du bouton FullExpandNode / FullCollapseNode }
    RefreshBtExpand;
  finally
    SourisNormale;
    TreeGrid.SetFocus;
  end
end;

{ Permet de réduire le noeud courant ------------------------------------------- }
procedure TFFrameTreeTob.CollapseNode(const Recurse: Boolean);
begin
  Grid_InvertNodeState(FCurrentRow, Recurse, False)
end;

{ Permet de d'étendre le noeud courant ----------------------------------------- }
procedure TFFrameTreeTob.ExpandNode(const Recurse: Boolean);
begin
  Grid_InvertNodeState(FCurrentRow, Recurse, True)
end;

{ Ajout d'une colonne. Renvoi l'index de celle-ci }
procedure TFFrameTreeTob.AddCol(const ColName, ColTitle: String;
  const ColWidth: Integer; const ColType: Char; const ColFormat: String;
  const ColAlign: TAlignment);
begin
  if ColField[ColName] = -1 then
    FListe.AddField(ColName, ColTitle, ColWidth, ColType, ColFormat, ColAlign)
end;

{ Permet de savoir si le champ "FieldName" est éditable dans la grille --------- }
function TFFrameTreeTob.IsEditableField(const FieldName: String): Boolean;
begin
  Result := not FConsultation and Assigned(FEditableFields) and ((FEditableFields.Count = 0) or (FEditableFields.IndexOf(FieldName) > -1))
end;

function TFFrameTreeTob.NodeIsExpanded(IM: TTreeManagerItem): Boolean;
begin
  if Assigned(IM) then
  begin
    if IM.IsParent then
      Result := IM.Expanded
    else if Assigned(IM.Parent) then
      Result := IM.Parent.Expanded
    else
      Result := False
  end
  else
    Result := False
end;

{ Permet de savoir si le noeud courante est étendue ---------------------------- }
function TFFrameTreeTob.GetExpanded: Boolean;
begin
  Result := NodeIsExpanded(CurrentItemManager)
end;

{ Permet de savoir si le de la ligne "iRow" courante est étendue --------------- }
function TFFrameTreeTob.GetExpandedRow(iRow: Integer): Boolean;
begin
  Result := NodeIsExpanded(ItemManagerRow[iRow])
end;

procedure TFFrameTreeTob.TreeGridRowEnter(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
begin
 FCurrentRow := TreeGrid.Row;
  { Mode editable ou non }
  if not FConsultation then
  begin
    TobRow[FCurrentRow].Modifie := False;
    if FTomEnabled then
    begin
      { Mise en phase de la tob data de la tom et de la tob en cours }
      FTom.FTob := TobRow[FCurrentRow];
      FTom.LoadBufferAvantModif(FTom.FTob);
    end;
    { Mémorisation de la tob avant toute modif. }
    FTreeTobBeforeModif.Dupliquer(TobRow[FCurrentRow], False, True);
    if not RowEditable[TreeGrid.Row] then
      TreeGrid.Options := TreeGrid.Options + [goRowSelect] - [goEditing, goAlwaysShowEditor]
    else
      TreeGrid.Options := TreeGrid.Options - [goRowSelect] + [goEditing, goAlwaysShowEditor];
    TreeGrid.InvalidateRow(TreeGrid.Row)
  end;
  RefreshBtExpand
end;

{ Envoi l'événement OnControlRecord et bloque si non valide -------------------- }
procedure TFFrameTreeTob.TreeGridRowExit(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
var
  LastErrorFieldName: String;
begin
  InitCellFound();
  if not FConsultation then
  begin
    DoControlRecord(LastErrorFieldName);
    Cancel := FLastError > 0;
    if Cancel and (LastErrorFieldName <> '') then
      GridSetFocusCol(LastErrorFieldName);
  end
end;

{ Focus de la cellule : Mode edit - Elipsis ------------------------------------ }
procedure TFFrameTreeTob.TreeGridCellEnter(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
begin
  inherited;
  if not FConsultation then
  begin
    FCurrentCol := TreeGrid.Col;
    { Montre l'elipsis dans le cas d'une cellule éditable de type "Combo" ou de type "Date" }
    TreeGrid.ElipsisButton := TreeGrid.ColEditables[TreeGrid.Col] and
                              ((wGetSimpleTypeField(CurrentField) = 'D') or
                                ((ChampToTT(CurrentField) <> '') and (Pos('CB=', TreeGrid.ColFormats[ACol]) = 0)));
    { Gère la taille du ValComboBox d'édition de la cellule }
    if TreeGrid.ColEditables[TreeGrid.Col] and (Pos('CB=', TreeGrid.ColFormats[ACol]) > 0) then
    begin
      TreeGrid.ValCombo.Height := TreeGrid.RowHeights[TreeGrid.Row] + iif(V_Pgi.Draw2003, 1, 0);
      TreeGrid.ValCombo.Width  := TreeGrid.ColWidths [TreeGrid.Col] + iif(V_Pgi.Draw2003, 1, 0);
    end;
    { Gère la "CheckBox" d'1 cellule éditable de type booléen }
    if TreeGrid.ColEditables[TreeGrid.Col] and (TreeGrid.ColTypes[TreeGrid.Col] = 'B') then
    begin
      TreeGrid.ColEditables[TreeGrid.Col] := False;
      FCbBooleanField.Checked := StrToBool_(TreeGrid.CellValues[TreeGrid.Col, TreeGrid.Row]);
      PnCbBooleanField.Visible := True;
      FitCheckBoxToCell;
      TreeGrid.CacheEdit;
      if FCbBooleanField.CanFocus then
        FCbBooleanField.SetFocus;
    end;
    if Cancel then
      FCurrentCol := ACol
  end
end;

{ Sauvegarde la cellule en cours dans la tob en cours -------------------------- }
procedure TFFrameTreeTob.SaveCell2Tob(const ACol, ARow: Integer; T: Tob; const ControlEgalite: Boolean = True);
var
  FieldName: String;
begin
  if ColEditable[FCurrentCol] then
  begin
    FieldName := FieldCol[ACol];  
    case wGetSimpleTypeField(FieldName) of
      'I': if not ControlEgalite or (ValeurI(TreeGrid.CellValues[ACol, ARow]) <> T.GetInteger(FieldName)) then
             T.SetInteger(FieldName, ValeurI(TreeGrid.CellValues[ACol, ARow]));
      'N': if not ControlEgalite or (Valeur(TreeGrid.CellValues[ACol, ARow]) <> T.GetDouble(FieldName)) then
             T.SetDouble(FieldName, Valeur(TreeGrid.CellValues[ACol, ARow]));
      'D': if not ControlEgalite or (StrToDate(TreeGrid.CellValues[ACol, ARow]) <> T.GetDateTime(FieldName)) then
             T.SetDateTime(FieldName, StrToDate(TreeGrid.CellValues[ACol, ARow]));
      'B': if not ControlEgalite or (StrToBool_(TreeGrid.CellValues[ACol, ARow]) <> T.GetBoolean(FieldName)) then
             T.SetBoolean(FieldName, StrToBool_(TreeGrid.CellValues[ACol, ARow]));
    else
      if TreeGrid.CellValues[ACol, ARow] <> T.GetString(FieldName) then
        T.SetString(FieldName, TreeGrid.CellValues[ACol, ARow]);
    end;
  end
end;

{ Exit de la cellule : Contrôle de la valeur ----------------------------------- }
procedure TFFrameTreeTob.TreeGridCellExit(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
var
  CurFieldName: String;
  T           : Tob;
begin
  inherited;
  InitCellFound();
  if not FConsultation then
  begin
    if ColEditable[ACol] then
    begin
      CurFieldName := FieldCol[ACol];
      T := TobRow[ARow];
      { Erreur ou pas on renseigne la tob }
      SaveCell2Tob(ACol, ARow, CurrentTob);
      { Mise à jour des données : DoControlField  passe par OnControlField de la Tom métier,
        puis dans le OnControlField local, puis dans le OnCalculField de la Tom métier,
        puis dans le OnCalculField local }
      DoControlField(T, CurFieldName);
      { Mise à jour des cumuls }
      UpdateCumuls(ARow);
      { Réactualise la tob servant à mémoriser les anciennes valeurs }
      if FLastError = 0 then
        SaveCell2Tob(ACol, ARow, FTreeTobBeforeModif, False);
      { En cas d'erreur on empêche la sortie de cellule }
      Cancel := FLastError > 0
    end;
    if not Cancel then
    begin
      { Si l'on sort d'une cellule de type boolean éditable, on masque la case à cocher
        puis on restitue le status d'édition de la cellule }
      if PnCbBooleanField.Visible and ((ACol <> TreeGrid.Col) or (ARow <> TreeGrid.Row)) then
      begin
        PnCbBooleanField.Visible := False;
        TreeGrid.ColEditables[ACol] := IsEditableField(FieldCol[ACol])
      end
    end
  end
end;

{ Initialisation des colonnes : Nom du champ associé + Format ------------------ }
procedure TFFrameTreeTob.InitCollumns;
var
  i: Integer;
  procedure SetColFormat(const FieldName: String);
  var
    FieldParams: TWFieldListe;
    TT, sField: String;
  begin
    { Gestion du mode d'édition de la cellule }
    TreeGrid.ColEditables[i] := IsEditableField(FieldName);
    { Gestion des Type, Largeur, Alignement et Format de la cellule }
    FieldParams := FListe.ChampsByName[FieldName];
    if Assigned(FieldParams) then
    begin
      TreeGrid.ColAligns[i] := FieldParams.Alignment;
      TreeGrid.ColWidths[i] := FieldParams.Width;
      sField := FieldName;
      if (Pos('(', sField) = 1) and (Pos(')', sField) = Length(sField)) then
      begin
        Delete(sField, 1, 1);
        Delete(sField, Length(sField), 1);
      end;
      case wGetSimpleTypeField(sField) of
        'B':
          begin
            TreeGrid.ColTypes[i] := 'B';
            TreeGrid.ColFormats[i] := IntToStr(Byte(csCoche));
          end;
        'D':
          begin
            TreeGrid.ColTypes[i] := 'D';
            TreeGrid.ColFormats[i] := ShortDateFormat;
          end;
        'N', 'I':
          begin
            TreeGrid.ColTypes[i] := 'R';
            TreeGrid.ColFormats[i] := FieldParams.Format;
          end;
        'M':
          begin
            TreeGrid.ColTypes[i] := 'M';
            TreeGrid.ColFormats[i] := IntToStr(Byte(msBook));
          end;
      else
        TreeGrid.ColTypes[i] := 'C';
        if TreeGrid.ColEditables[i] then
        begin
          TT := ChampToTT(FieldName);
          if (FieldParams.LibelleComplet) and (TT <> '') then
            TreeGrid.ColFormats[i] := 'CB=' + TT
          else
            TreeGrid.ColFormats[i] := ''
        end
      end
    end;
  end; 
begin
  PCumul.Visible := False;
  try
    { Formats d'affichage }
    for i := Grid_FixedCols to (TreeGrid.ColCount - Grid_FixedCols) do
    begin
      TreeGrid.ColNames[i] := FTobFields.GetString('FIELD_' + IntToStr(i));
      SetColFormat(FieldCol[i]);
    end;
  finally
    PCumul.Visible := True;
  end;
end;

{ Permet de savoir si le noeud courant est un noeud parent --------------------- }
function TFFrameTreeTob.GetIsParentNode: Boolean;
begin
  Result := Assigned(CurrentTob) and (CurrentTob.Detail.Count > 0)
end;

function TFFrameTreeTob.TobGetBoolean(T: Tob; const FieldName: String): Boolean;
begin
  Result := T.GetBoolean(FieldName)
end;

function TFFrameTreeTob.TobGetDateTime(T: Tob; const FieldName: String): TDateTime;
begin
  Result := T.GetDateTime(FieldName)
end;

function TFFrameTreeTob.TobGetDouble(T: Tob; const FieldName: String): Double;
begin
  Result := T.GetDouble(FieldName)
end;

function TFFrameTreeTob.TobGetInteger(T: Tob; const FieldName: String): Integer;
begin
  Result := T.GetInteger(FieldName)
end;

function TFFrameTreeTob.TobGetString(T: Tob; const FieldName: String): String;
begin
  Result := T.GetString(FieldName)
end;

procedure TFFrameTreeTob.TobDoUpdate(T: Tob; const FieldName: String; const Value: Variant);
var
  SynDataSources: Boolean;
  TobTom,
  FTreeTobBeforeModifSaved,
  FCurrentProcessedTobSaved: Tob;
  { Synchronise la source de donnée à traiter }
  procedure BeginSyn;
  begin
    if SynDataSources then
    begin
      { Synchro. de la source de donnée avec la Tom }
      if FTomEnabled then
      begin
        TobTom := FTom.FTob;
        FTom.FTob := T;
        FTom.LoadBufferAvantModif(FTom.FTob);
      end
      else
        TobTom := nil;
      { Synchro. de la source de donnée avec la TreeTob }
      FCurrentProcessedTobSaved := FCurrentProcessedTob;
      FCurrentProcessedTob := T;
      FTreeTobBeforeModifSaved := Tob.Create(TableName, nil, -1);
      FTreeTobBeforeModifSaved.Dupliquer(FTreeTobBeforeModif, False, True, True);
      FTreeTobBeforeModif.Dupliquer(T, False, True, False)
    end
  end;
  { Resynchronise la source de donnée traitée }
  procedure EndSyn;
  begin
    if SynDataSources then
    begin
      { recopie la dernière valeur saisie }
      if FLastError = 0 then
        FTreeTobBeforeModifSaved.PutValue(FieldName, Value);
      { Resynchro. de la source de donnée avec la TreeTob }
      FTreeTobBeforeModif.Dupliquer(FTreeTobBeforeModifSaved, False, True, True);
      FTreeTobBeforeModifSaved.Free;
      FCurrentProcessedTob := FCurrentProcessedTobSaved;
      { Resynchro. de la source de donnée avec la Tom }
      if FTomEnabled then
        FTom.FTob := TobTom
    end
  end;
begin
  SynDataSources := not FConsultation and (T <> TobRow[FCurrentRow]);  
  BeginUpdate;
  BeginSyn;
  try
    T.PutValue(FieldName, Value);
    if not FConsultation then
      DoControlField(T, FieldName);
//    T.Modifie := False;
  finally
    EndSyn;
    EndUpdate
  end
end;

procedure TFFrameTreeTob.TobSetString(T: Tob; const FieldName, Value: String);
begin
  TobDoUpdate(T, FieldName, Value)
end;

procedure TFFrameTreeTob.TobSetInteger(T: Tob; const FieldName: String; const Value: Integer);
begin
  TobDoUpdate(T, FieldName, Value)
end;

procedure TFFrameTreeTob.TobSetDouble(T: Tob; const FieldName: String; const Value: Double);
begin
  TobDoUpdate(T, FieldName, Value)
end;

procedure TFFrameTreeTob.TobSetDateTime(T: Tob; const FieldName: String; const Value: TDateTime);
begin
  TobDoUpdate(T, FieldName, Value)
end;

procedure TFFrameTreeTob.TobSetBoolean(T: Tob; const FieldName: String; const Value: Boolean);
begin
  TobDoUpdate(T, FieldName, Value)
end;

{ GETTER de donnée de la tob en cours --- Début -------------------------------- }
function TFFrameTreeTob.GetBoolean(const FieldName: String): Boolean;
begin
  Result := TobGetBoolean(CurrentTob, FieldName)
end;

function TFFrameTreeTob.GetDateTime(const FieldName: String): TDateTime;
begin
  Result := TobGetDateTime(CurrentTob, FieldName)
end;

function TFFrameTreeTob.GetDouble(const FieldName: String): Double;
begin
  Result := TobGetDouble(CurrentTob, FieldName)
end;

function TFFrameTreeTob.GetInteger(const FieldName: String): Integer;
begin
  Result := TobGetInteger(CurrentTob, FieldName)
end;

function TFFrameTreeTob.GetString(const FieldName: String): String;
begin
  Result := TobGetString(CurrentTob, FieldName)
end;

function TFFrameTreeTob.GetOldString(const FieldName: String): String;
begin
  Result := TobGetString(FTreeTobBeforeModif, FieldName)
end;

function TFFrameTreeTob.GetOldInteger(const FieldName: String): Integer;
begin
  Result := TobGetInteger(FTreeTobBeforeModif, FieldName)
end;

function TFFrameTreeTob.GetOldDouble(const FieldName: String): Double;
begin
  Result := TobGetDouble(FTreeTobBeforeModif, FieldName)
end;

function TFFrameTreeTob.GetOldDateTime(const FieldName: String): TDateTime;
begin
  Result := TobGetDateTime(FTreeTobBeforeModif, FieldName)
end;

function TFFrameTreeTob.GetOldBoolean(const FieldName: String): Boolean;
begin
  Result := TobGetBoolean(FTreeTobBeforeModif, FieldName)
end;
{ --- Fin GETTER de donnée tob en cours ---------------------------------------- }

{ SETTER de donnée de la tob en cours --- Début -------------------------------- }
procedure TFFrameTreeTob.SetString(const FieldName, Value: String);
begin
  TobSetString(CurrentTob, FieldName, Value);
end;

procedure TFFrameTreeTob.SetInteger(const FieldName: String; Value: Integer);
begin
  TobSetInteger(CurrentTob, FieldName, Value);
end;

procedure TFFrameTreeTob.SetDouble(const FieldName: String; Value: Double);
begin
  TobSetDouble(CurrentTob, FieldName, Value);
end;

procedure TFFrameTreeTob.SetDateTime(const FieldName: String; Value: TDateTime);
begin
  TobSetDateTime(CurrentTob, FieldName, Value);
end;

procedure TFFrameTreeTob.SetBoolean(const FieldName: String; Value: Boolean);
begin
  TobSetBoolean(CurrentTob, FieldName, Value);
end;
{ --- Fin SETTER de donnée tob en cours ---------------------------------------- }

{ SETTER de donnée de la tob en cours --- Début -------------------------------- }
procedure TFFrameTreeTob.SetParentString(const FieldName, Value: String);
begin
  TobSetString(CurrentTob.Parent, FieldName, Value)
end;

procedure TFFrameTreeTob.SetParentInteger(const FieldName: String; Value: Integer);
begin
  TobSetInteger(CurrentTob.Parent, FieldName, Value)
end;

procedure TFFrameTreeTob.SetParentDouble(const FieldName: String; Value: Double);
begin
  TobSetDouble(CurrentTob.Parent, FieldName, Value)
end;

procedure TFFrameTreeTob.SetParentDateTime(const FieldName: String; Value: TDateTime);
begin
  TobSetDateTime(CurrentTob.Parent, FieldName, Value)
end;

procedure TFFrameTreeTob.SetParentBoolean(const FieldName: String; Value: Boolean);
begin
  TobSetBoolean(CurrentTob.Parent, FieldName, Value)
end;
{ --- Fin SETTER de donnée tob en cours ---------------------------------------- }

{ SETTER de donnée de la tob fille ---- Début ---------------------------------- }
procedure TFFrameTreeTob.SetChildString(const Index: Integer; const FieldName, Value: String);
begin
  TobSetString(CurrentTob.Detail[Index], FieldName, Value)
end;

procedure TFFrameTreeTob.SetChildInteger(const Index: Integer; const FieldName: String; Value: Integer);
begin
  TobSetInteger(CurrentTob.Detail[Index], FieldName, Value)
end;

procedure TFFrameTreeTob.SetChildDouble(const Index: Integer; const FieldName: String; Value: Double);
begin
  TobSetDouble(CurrentTob.Detail[Index], FieldName, Value)
end;

procedure TFFrameTreeTob.SetChildDateTime(const Index: Integer; const FieldName: String; Value: TDateTime);
begin
  TobSetDateTime(CurrentTob.Detail[Index], FieldName, Value)
end;

procedure TFFrameTreeTob.SetChildBoolean(const Index: Integer; const FieldName: String; Value: Boolean);
begin
  TobSetBoolean(CurrentTob.Detail[Index], FieldName, Value)
end;

procedure TFFrameTreeTob.SetChildrenString(const FieldName, Value: String);
var
  i: Integer;
begin
  for i := 0 to Pred(CurrentTob.Detail.Count) do
    SetChildString(i, FieldName, Value)
end;

procedure TFFrameTreeTob.SetChildrenInteger(const FieldName: String; Value: Integer);
var
  i: Integer;
begin
  for i := 0 to Pred(CurrentTob.Detail.Count) do
    SetChildInteger(i, FieldName, Value)
end;

procedure TFFrameTreeTob.SetChildrenDouble(const FieldName: String; Value: Double);
var
  i: Integer;
begin
  for i := 0 to Pred(CurrentTob.Detail.Count) do
    SetChildDouble(i, FieldName, Value)
end;

procedure TFFrameTreeTob.SetChildrenDateTime(const FieldName: String; Value: TDateTime);
var
  i: Integer;
begin
  for i := 0 to Pred(CurrentTob.Detail.Count) do
    SetChildDateTime(i, FieldName, Value)
end;

procedure TFFrameTreeTob.SetChildrenBoolean(const FieldName: String; Value: Boolean);
var
  i: Integer;
begin
  for i := 0 to Pred(CurrentTob.Detail.Count) do
    SetChildBoolean(i, FieldName, Value)
end;
{ --- Fin SETTER de donnée tob fille ------------------------------------------- }

{ --- Début autre SETTER --- }
procedure TFFrameTreeTob.SetDisplayTree(const Value: Boolean);
begin
  if Value <> FDisplayTree then
  begin
    FDisplayTree := Value;
    TreeGrid.Repaint
  end
end;
{ ---- Fin autre SETTER ---- }

{ GETTER de donnée de la tob parent --- Début ---------------------------------- }
function TFFrameTreeTob.GetParentBoolean(const FieldName: String): Boolean;
begin
  Result := TobGetBoolean(CurrentTob.Parent, FieldName)
end;

function TFFrameTreeTob.GetParentDateTime(const FieldName: String): TDateTime;
begin
  Result := TobGetDateTime(CurrentTob.Parent, FieldName)
end;

function TFFrameTreeTob.GetParentDouble(const FieldName: String): Double;
begin
  Result := TobGetDouble(CurrentTob.Parent, FieldName)
end;

function TFFrameTreeTob.GetParentInteger(const FieldName: String): Integer;
begin
  Result := TobGetInteger(CurrentTob.Parent, FieldName)
end;

function TFFrameTreeTob.GetParentString(const FieldName: String): String;
begin
  Result := TobGetString(CurrentTob.Parent, FieldName)
end;
{ --- Fin GETTER de donnée tob parent ------------------------------------------ }

{ GETTER de donnée de la tob fille ---- Début ---------------------------------- }
function TFFrameTreeTob.GetChildString(const Index: Integer; const FieldName: String): String;
begin
  Result := TobGetString(CurrentTob.Detail[Index], FieldName)
end;

function TFFrameTreeTob.GetChildInteger(const Index: Integer; const FieldName: String): Integer;
begin
  Result := TobGetInteger(CurrentTob.Detail[Index], FieldName)
end;

function TFFrameTreeTob.GetChildDouble(const Index: Integer; const FieldName: String): Double;
begin
  Result := TobGetDouble(CurrentTob.Detail[Index], FieldName)
end;

function TFFrameTreeTob.GetChildDateTime(const Index: Integer; const FieldName: String): TDateTime;
begin
  Result := TobGetDateTime(CurrentTob.Detail[Index], FieldName)
end;

function TFFrameTreeTob.GetChildBoolean(const Index: Integer; const FieldName: String): Boolean;
begin
  Result := TobGetBoolean(CurrentTob.Detail[Index], FieldName)
end;
{ --- Fin GETTER de donnée tob fille ------------------------------------------- }

{ Ramène la tob en cours i.e. la tob associée à la ligne de grille sélectionnée  }
function TFFrameTreeTob.GetCurrentTob: Tob;
begin
  if FCurrentProcessedTob = nil then
    Result := TobRow[FCurrentRow]
  else
    Result := FCurrentProcessedTob
end;

function TFFrameTreeTob.GetCurrentItemManager: TTreeManagerItem;
begin
  Result := ItemManagerRow[FCurrentRow]
end;

{ Permet de connaître le nom du champ associé à la colonne en cours ------------ }
function TFFrameTreeTob.GetCurrentFieldName: String;
begin
  Result := FieldCol[FCurrentCol]
end;

{ Permet de connaître le nom de la table sur laquelle s'appuie la TreeTob ------ }
function TFFrameTreeTob.GetTableName: String;
begin
  if TableToNum(FTreeTob.NomTable) > -1 then
    Result := FTreeTob.NomTable
  else
    Result := ''
end;

{ Compte le nombre de tobs multi-niveaux --------------------------------------- }
function TFFrameTreeTob.GetCountDetails(iRow: Integer): Integer;
  procedure _CountDetails(T: Tob; var Nb: Integer);
  var
    i: Integer;
  begin
    for i := 0 to Pred(T.Detail.Count) do
    begin
      Inc(Nb);
      _CountDetails(T.Detail[i], Nb)
    end
  end;   
begin
  Result := 0;
  _CountDetails(TobRow[iRow], Result)
end;

{ Traite l'affichage en fonction de la recherche ------------------------------- }
procedure TFFrameTreeTob.FindDialogFind(Sender: TObject);
var
  Coords : TGridCoord;
  TM     : TTreeManager;
  IM, IMP: TTreeManagerItem;
  sField : String;
begin
  SourisSablier;
  try
    TM := FTreeManager;
    Coords.X := TreeGrid.Col;
    Coords.Y := TreeGrid.Row;
    FTreeManager.FLastIndex := ItemManagerRow[TreeGrid.Row].Index;
    if FTreeManager.FFirstFind then
    begin
      if FConsultation then
        FTreeManager.FLastColFound := TobRow[TreeGrid.Row].GetNumChamp(FieldCol[iif(frDown in FindDialog.Options, TreeGrid.FixedCols, Pred(TreeGrid.ColCount))])
      else if FieldCol[TreeGrid.Col] <> '' then
        FTreeManager.FLastColFound := TobRow[TreeGrid.Row].GetNumChamp(FieldCol[TreeGrid.Col])
      else
        FTreeManager.FLastColFound := -1
    end;
    { Recherche de la chaîne... }
    with FindDialog do
    begin
      IM := TM.Find(FindText, Options, True, sField);
      while Assigned(IM) and (ColField[sField] = -1) do
        IM := TM.Find(FindText, Options, True, sField)
    end;
    { Recherche & focus dans la grille }
    if Assigned(IM) then
    begin
      { Visibilité du noeud recherché }
      IMP := IM.FirstParentCollapsed;
      while Assigned(IMP) do
      begin
        Grid_InvertNodeState(IMP.RowIndex, False, True);
        IMP := IM.FirstParentCollapsed;
      end;
      { Positionnement X,Y }
      TreeGrid.Row := IM.RowIndex;
      TreeGrid.Col := ColField[sField];
      { Pour mise en valeur de la cellule trouvé }
      FCellFound := Point(ColField[sField], TreeGrid.Row);
      { Focus }
      if FConsultation or (Coords.X <> FCellFound.X) or (Coords.Y <> FCellFound.Y) then
        Grid_SetFocus(FCellFound.X, FCellFound.Y)
    end
  finally
    SourisNormale
  end
end;

{ Ramène le niveau de la tob en fonction du Niveau du premier noeud de la FTreeTob }
function TFFrameTreeTob.TobGetNiveau(T: Tob): Integer;
begin
  Result := T.Niveau - FTreeTob.Niveau
end;

{ Contexte Loupe --------------------------------------------------------------- }
procedure TFFrameTreeTob.PmLoupePopup(Sender: TObject);
var
  Ctx: String;
begin
  inherited;
  Ctx := '';
  if Assigned(FOnMnuLoupeGetCtx) then
    FOnMnuLoupeGetCtx(Ctx);
  wSetMnuLoupe(False, Ctx, PmLoupe);
  ActivateXPPopUp(PmLoupe);
end;

{ Etendre / Réduire un noeud --------------------------------------------------- }
procedure TFFrameTreeTob.TreeGridMouseDown(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
var
  Coords: TGridCoord;
  RowSaved: Integer;
begin
  inherited;
  if Button = mbLeft then
  begin
    if Grid_CanClickNode(X, Y) then
    begin
      Coords := TreeGrid.MouseCoord(X, Y);
      if RecordIsValid(False) then
      begin
        RowSaved := FCurrentRow;
        if FDisplayTree then
          FCurrentRow := Coords.Y;
        try
          with CurrentTob do
          begin
            if (Detail.Count > 0) then
            begin
              if ItemManagerRow[FCurrentRow].Expanded then
                CollapseNode
              else
                ExpandNode
            end
          end;
        finally
          FCurrentRow := RowSaved
        end;
        if Assigned(FOnNodeMouseClick) then
          FOnNodeMouseClick
      end
    end;
    if  IsDayPass and (TreeGrid.FixedRows > 0) and (TreeGrid.FixedCols > 0)
    and wBetween(X, 1, TreeGrid.ColWidths[0] - 1) and wBetween(Y, 1, TreeGrid.RowHeights[0] - 1) then
      TobDebug(FTreeTob)
  end
end;

function TFFrameTreeTob.Grid_CanClickNode(const X, Y: Integer): Boolean;
begin
  Result := (FPoint.x > 0) and (FPoint.y > 0)
            and wBetween(X, FPoint.x, FPoint.x + FImageList.Width)
            and wBetween(Y, FPoint.y, FPoint.y + FImageList.Height)
end;

{ En Mode S.A.V. uniquement, ShowTob de la Tob en cours ------------------------ }
procedure TFFrameTreeTob.MnSavShowTobClick(Sender: TObject);
begin
  TobDebug(CurrentTob)
end;

procedure TFFrameTreeTob.MnSavShowTob2Click(Sender: TObject);
begin
  TobDebug(FTreeTobBeforeModif)
end;

{ Gestion de l'elipsis --------------------------------------------------------- }
procedure TFFrameTreeTob.TreeGridElipsisClick(Sender: TObject);
var
  KeyDate : Char;
  TT, Plus: String;
  C       : THEdit;
  R       : TRect;
begin
  inherited;
  if TreeGrid.ColEditables[TreeGrid.Col] then
  begin
    { Choix de la date... }
    if wGetSimpleTypeField(CurrentField) = 'D' then
    begin
      KeyDate := '*';
      ParamDate(TForm(Self), TreeGrid, KeyDate)
    end
    else
    { LookUp basé sur une tablette...d'où la création d'un THEdit masqué }
    begin
      TT := ChampToTT(CurrentField);
      if TT <> '' then
      begin
        C := THEdit.Create(TreeGrid);
        try
          C.Visible := False;
          C.Parent := TreeGrid;
          C.DataType := TT;
          Plus := '';
          if Assigned(FOnElipsisGetPlus) then
            FOnElipsisGetPlus(CurrentField, Plus);
          C.Plus := Plus;
          R := TreeGrid.CellRect(TreeGrid.Col, TreeGrid.Row);
          C.Top := R.Top;
          C.Left := R.Left;
          C.Width := R.Right - R.Left;
          C.Height := R.Bottom - R.Top;
          if LookUpCombo(C) then
            TreeGrid.CellValues[TreeGrid.Col, TreeGrid.Row] := C.Text
        finally
          C.Free
        end
      end
    end
  end
end;

{ Gestion du click du bouton d'exploration de noeuds --------------------------- }
procedure TFFrameTreeTob.BtExpandClick(Sender: TObject);
begin
  inherited;
  if NodeExpanded then
    CollapseNode(True)
  else
    ExpandNode(True)
end;

{ Gestion du click du bouton de paramétrage de la DBListe ---------------------- }
procedure TFFrameTreeTob.BtParamListeClick(Sender: TObject);
begin
  if RecordIsValid then
  begin
{$IFDEF VER130}    //PT2
  {$IFDEF EAGLCLIENT}
      ParamListe(FDBListe, nil, '');
  {$ELSE  EAGLCLIENT}
      ParamListe(FDBListe, nil, nil, '');
  {$ENDIF EAGLCLIENT}
{$ELSE  VER130}
  {$IFDEF EAGLCLIENT}
      if ParamListe(FDBListe, nil, '') then
  {$ELSE  EAGLCLIENT}
      if ParamListe(FDBListe, nil, nil, '') then
  {$ENDIF EAGLCLIENT}
{$ENDIF VER130}
    begin
      DisableRefresh;
      try
        RefreshDBListe();
        RefreshTreeTob();
      finally
        EnableRefresh
      end
    end
  end
end;

{ Gestion du click du bouton d'ajout de ligne ---------------------------------- }
procedure TFFrameTreeTob.BtAddRowClick(Sender: TObject);
begin
  AddRow(False);
end;

procedure TFFrameTreeTob.BtAddChildRowClick(Sender: TObject);
begin
  AddRow(True);
end;

{ Gestion du click du bouton de suppression de ligne --------------------------- }
procedure TFFrameTreeTob.BtDeleteRowClick(Sender: TObject);
begin
  if PgiAsk('Confirmez-vous la suppression de la branche sélectionnée ?') = mrYes then
    DeleteRow;
end;

{ Gestion du click du bouton d'annulation de ligne ----------------------------- }
procedure TFFrameTreeTob.BtUndoClick(Sender: TObject);
begin
  if PgiAsk('Confirmez-vous l''annulation complète de la saisie ?') = mrYes then
    Undo;
end;

{ Réduit / Etend le noeud courant ---------------------------------------------- }
{ Gère l'affichage du bouton Réduire/Etendre le noeud courant ------------------ }
procedure TFFrameTreeTob.RefreshBtExpand;
const
  Img_FullExpand = 4;
  Img_FullCollapse = 5;
begin
  if Assigned(FBtExpand) then
  begin
    FBtExpand.Visible := IsParentNode;
    FBtExpand.ImageIndex := iif(NodeExpanded, Img_FullCollapse, Img_FullExpand);
    FBtExpand.Hint := iif(NodeExpanded, TraduireMemoire('Réduire toute la branche'), TraduireMemoire('Etendre toute la branche'))
  end
end;

{ Permet de gérer nativement un bouton d'exploration de noeuds ----------------- }
procedure TFFrameTreeTob.SetBtExpand(const Value: TToolBarButton97);
begin
  FBtExpand := Value;
  if Assigned(FBtExpand) then
  begin
    FBtExpand.Visible := True;
    FBtExpand.ShowHint := True;
    FBtExpand.Hint := TraduireMemoire('Réduire toute la branche');
    FBtExpand.Images := FImageList;
    FBtExpand.OnClick := BtExpandClick
  end
end;

procedure TFFrameTreeTob.SetBtLoupe(const Value: TToolBarButton97);
begin
  FBtLoupe := Value;
  if Assigned(FBtLoupe) then
  begin
    FBtLoupe.Visible := True;
    FBtLoupe.ShowHint := True;
    FBtLoupe.Hint := TraduireMemoire('Menu zoom');
    FBtLoupe.DropdownMenu := PmLoupe
  end
end;

procedure TFFrameTreeTob.SetBtParamListe(const Value: TToolBarButton97);
begin
  FBtParamListe := Value;
  if Assigned(FBtParamListe) then
  begin
    FBtParamListe.Visible := True;
    FBtParamListe.ShowHint := True;
    FBtParamListe.Hint := TraduireMemoire('Paramétrer la présentation');
    FBtParamListe.OnClick := BtParamListeClick;
  end
end;

procedure TFFrameTreeTob.SetBtAddRow(const Value: TToolBarButton97);
begin
  FBtAddRow := Value;
  if Assigned(FBtAddRow) then
  begin
    FBtAddRow.Visible := not FConsultation;
    FBtAddRow.ShowHint := True;
    FBtAddRow.Hint := TraduireMemoire('Nouveau');
    FBtAddRow.OnClick := BtAddRowClick;
  end
end;

procedure TFFrameTreeTob.SetBtAddChildRow(const Value: TToolBarButton97);
begin
  FBtAddChildRow := Value;
  if Assigned(FBtAddChildRow) then
  begin
    FBtAddChildRow.Visible := not FConsultation;
    FBtAddChildRow.ShowHint := True;
    FBtAddChildRow.Hint := TraduireMemoire('Nouveau sous-niveau');
    FBtAddChildRow.OnClick := BtAddChildRowClick;
  end
end;

procedure TFFrameTreeTob.SetBtDeleteRow(const Value: TToolBarButton97);
begin
  FBtDeleteRow := Value;
  if Assigned(FBtDeleteRow) then
  begin
    FBtDeleteRow.Visible := True;
    FBtDeleteRow.ShowHint := True;
    FBtDeleteRow.Hint := TraduireMemoire('Supprimer');
    FBtDeleteRow.OnClick := BtDeleteRowClick;
  end
end;

procedure TFFrameTreeTob.SetBtUndo(const Value: TToolBarButton97);
begin
  FBtUndo := Value;
  if Assigned(FBtUndo) then
  begin
    FBtUndo.Visible := not FConsultation;
    FBtUndo.ShowHint := True;
    FBtUndo.Hint := TraduireMemoire('Annuler');
    FBtUndo.OnClick := BtUndoClick;
  end
end;

{ Gestion du click du bouton de recherche -------------------------------------- }
procedure TFFrameTreeTob.BtSearchClick(Sender: TObject);
begin
  if RecordIsValid then
  begin
    FTreeManager.FFirstFind := True;
    FindDialog.Execute
  end
end;

procedure TFFrameTreeTob.BtPrintGridClick(Sender: TObject);
begin
  PrintGrid;
end;

procedure TFFrameTreeTob.SetBtSearch(const Value: TToolBarButton97);
begin
  FBtSearch := Value;
  if Assigned(FBtSearch) then
  begin
    FBtSearch.Visible := True;
    FBtSearch.ShowHint := True;
    FBtSearch.Hint := TraduireMemoire('Rechercher dans la liste');
    FBtSearch.OnClick := BtSearchClick;
  end
end;

procedure TFFrameTreeTob.SetBtPrintGrid(const Value: TToolBarButton97);
begin
  FBtPrintGrid := Value;
  if Assigned(FBtPrintGrid) then
  begin
    FBtPrintGrid.Visible := True;
    FBtPrintGrid.ShowHint := True;
    FBtPrintGrid.Hint := TraduireMemoire('Imprimer l''arbre');
    FBtPrintGrid.OnClick := BtPrintGridClick;
  end
end;

{ Gestion du click du bouton de sélection complète ----------------------------- }
procedure TFFrameTreeTob.BtSelectAllClick(Sender: TObject);
begin
  UnassignNotifyEvent(TreeGrid.OnFlipSelection);
  try
    if FConsultation then
      TreeGrid.AllSelected := TToolBarButton97(Sender).Down
    else
      SavError('Contexte exclusif en consultation !', 'TFFrameTreeTob.BtSelectAllClick')
  finally
    ReassignNotifyEvent(TreeGrid.OnFlipSelection)
  end
end;

procedure TFFrameTreeTob.SetBtSelectAll(const Value: TToolBarButton97);
begin
  FBtSelectAll := Value;
  TreeGrid.MultiSelect := Assigned(FBtSelectAll) or Assigned(FBtSelectNode);
  if Assigned(FBtSelectAll) then
  begin
    FBtSelectAll.Visible := FConsultation;
    FBtSelectAll.ShowHint := True;
    FBtSelectAll.Hint := TraduireMemoire('Tout sélectionner');
    FBtSelectAll.OnClick := BtSelectAllClick;
    FBtSelectAll.GroupIndex := 22;
    FBtSelectAll.AllowAllUp := True;
  end
end;

{ Gestion du click du bouton de sélection du noeud ----------------------------- }
procedure TFFrameTreeTob.BtSelectNodeClick(Sender: TObject);
var
  i: Integer;
begin
  if FConsultation then
  begin
    TreeGrid.AllSelected := False;
    if FBtSelectNode.Down then
    begin
      if not NodeExpanded then
        ExpandNode(True);
      TreeGrid.FlipSelection(FCurrentRow);
      for i := 0 to Pred(CountDetails[FCurrentRow]) do
        TreeGrid.FlipSelection(FCurrentRow + i + 1)
    end
  end
  else
    SavError('Contexte exclusif en consultation !', 'TFFrameTreeTob.BtSelectNodeClick')
end;

procedure TFFrameTreeTob.SetBtSelectNode(const Value: TToolBarButton97);
begin
  FBtSelectNode := Value;
  TreeGrid.MultiSelect := Assigned(FBtSelectAll) or Assigned(FBtSelectNode);
  if Assigned(FBtSelectNode) then
  begin
    FBtSelectNode.Visible := FConsultation;
    FBtSelectNode.ShowHint := True;
    FBtSelectNode.Hint := TraduireMemoire('Sélectionner la branche');
    FBtSelectNode.OnClick := BtSelectNodeClick;
    FBtSelectNode.GroupIndex := 22;
    FBtSelectNode.AllowAllUp := True;
  end
end;

procedure TFFrameTreeTob.SetLinkedCtrls(const Value: TTreeTobLinkedCtrls);
begin
  FLinkedCtrls := Value;
  if Assigned(FLinkedCtrls) then
  begin
    BtExpand      := FLinkedCtrls.BtExpand;
    BtLoupe       := FLinkedCtrls.BtLoupe;
    BtParamListe  := FLinkedCtrls.BtParamListe;
    BtSearch      := FLinkedCtrls.BtSearch;
    BtPrintGrid   := FLinkedCtrls.BtPrintGrid;
    PopupLoupe    := FLinkedCtrls.PmLoupe;
    BtAddRow      := FLinkedCtrls.BtAddRow;
    BtAddChildRow := FLinkedCtrls.BtAddChildRow;
    BtDeleteRow   := FLinkedCtrls.BtDeleteRow;
    BtUndo        := FLinkedCtrls.BtUndo;
    BtSelectAll   := FLinkedCtrls.BtSelectAll;
    BtSelectNode  := FLinkedCtrls.BtSelectNode;
  end
end;

{ Permet de gérer le menu zoom ------------------------------------------------- }
procedure TFFrameTreeTob.SetPmLoupe(const Value: THPopupMenu);
var
  TabMI: Array of TMenuItem;
begin
  if Assigned(Value) and (Value.Items.Count > 0) then
  begin
    while Value.Items.Count > 0 do
    begin
      SetLength(TabMI, Length(TabMI) + 1);
      TabMI[Pred(Length(TabMI))] := Value.Items[0];
      Value.Items.Delete(0)
    end;
    PmLoupe.Items.Add(TabMI)
  end
end;

{ Combinaisons de touches ------------------------------------------------------ }
procedure TFFrameTreeTob.TreeGridKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
  { Echap }
  if (Key = Vk_Escape) and Assigned(FForm) then
    SendMessage(FForm.Handle, WM_KEYDOWN, Key, 0)
  { Recherche }
  else if (ssCtrl in Shift) and (Key = Vk_Cherche) and Assigned(FBtSearch)
  and (FBtSearch.Visible) and (FBtSearch.Enabled) then
    BtSearch.Click
  { Recherche }
  else if (ssCtrl in Shift) and (Key = vk_imprime) and Assigned(FBtPrintGrid)
  and (FBtPrintGrid.Visible) and (FBtPrintGrid.Enabled) then
    BtPrintGrid.Click
  { Sélection }
  else if (ssCtrl in Shift) and (Key = Vk_SelectAll) and Assigned(FBtSelectAll)
  and (FBtSelectAll.Visible) and (FBtSelectAll.Enabled) then
    BtSelectAll.Click
  { Ouvrir ligne }
  else if FConsultation and (Key = vk_Recherche) then
    TreeGrid.OnDblClick(Sender)
  { Elipsis }
  else if (Key = Vk_Recherche) and TreeGrid.ElipsisButton then
    TreeGridElipsisClick(Sender)
  { Ajout }
  else if (ssCtrl in Shift) and (Key = Vk_InsMul) then
  begin
    { Ajout même niveau }
    if (ssShift in Shift) and Assigned(FBtAddChildRow) and FBtAddChildRow.Visible and FBtAddChildRow.Enabled then
      FBtAddChildRow.Click
    { Ajout Fils }
    else if Assigned(FBtAddRow) and FBtAddRow.Visible and FBtAddRow.Enabled then
      FBtAddRow.Click
  end
  { Suppression }
  else if (ssCtrl in Shift) and (Key = Vk_Delete) and Assigned(FBtDeleteRow)
  and FBtDeleteRow.Visible and FBtDeleteRow.Enabled then
  begin
    FBtDeleteRow.Click;
    Key := 0
  end
  { Annulation }
  else if (ssCtrl in Shift) and (Key = Vk_Defaire) and Assigned(FBtUndo)
  and FBtUndo.Visible and FBtUndo.Enabled then
  begin
    FBtUndo.Click;
    Key := 0
  end
end;

{ Gestion de la sélection de lignes -------------------------------------------- }
procedure TFFrameTreeTob.TreeGridFlipSelection(Sender: TObject);
begin
  UnassignNotifyEvent(FBtSelectAll.OnClick);
  try
    FBtSelectAll.Down := TreeGrid.NbSelected = TreeGrid.RowCount - TreeGrid.FixedRows
  finally
    ReassignNotifyEvent(FBtSelectAll.OnClick)
  end
end;

{ Crée la FrameTreeTob dans une Form ------------------------------------------- }
constructor TFFrameTreeTob.Create(F: TForm; C: TWinControl; ATreeTob: Tob; const ADBListe: String;
                                  ALinkedCtrls: TTreeTobLinkedCtrls; const AConsultation: Boolean = True;
                                  const AEditableFields: TStrings = nil; const AShowRoot: Boolean = True;
                                  const ADisplayTree: Boolean = True);
var
  i: Integer;
  sClef: String;
  iTable: Integer;
begin
  inherited Create(F);
  FForm                   := F;
  Parent                  := C;
  FTreeManager            := TTreeManager.Create(ATreeTob, TreeGrid.FixedRows, AShowRoot);
  TreeTob                 := ATreeTob;
  DBListe                 := ADBListe;
  LinkedCtrls             := ALinkedCtrls;
  FEditableFields         := TStringList.Create;
  FShowRoot               := AShowRoot;
  Consultation            := AConsultation;
  FRatedFields            := TStringList.Create;
  InitCellFound();
  { Initialisation de la visu des champs de type RATE }
  iTable := TableToNum(TableName);
  if iTable > 0 then
  begin
    {$IFDEF EAGLCLIENT}
      if High(V_Pgi.DEChamps[iTable]) <= 0 then
        ChargeDeChamps(iTable, TableToPrefixe(TableName));
    {$ENDIF EAGLCLIENT}
    for i := Low(V_Pgi.DEChamps[iTable]) to High(V_Pgi.DEChamps[iTable]) do
      if V_Pgi.DEChamps[iTable, i].Tipe = 'RATE' then
        FRatedFields.Add(V_Pgi.DEChamps[iTable, i].Nom)
  end;
  FindDialog              := TFindDialog.Create(FForm);
  FindDialog.OnFind       := FindDialogFind;
  FCbBooleanField         := TTreeGridCheckBox.Create(PnCbBooleanField);
  FCbBooleanField.FGrid   := TreeGrid;
  FCbBooleanField.Name    := 'GRIDCOMP_CbBooleanField';
  FCbBooleanField.Parent  := PnCbBooleanField;
  FCbBooleanField.OnClick := CbBooleanFieldClick;
  if Assigned(FForm) then
    FForm.OnResize := Form_Resize;
  { Renseigne la liste des champs éditables
    Liste vide (mais assignée) signifie "Tous les champs de la DBListe" }
  if Assigned(AEditableFields) then
  begin
    AEditableFields.Add('');
    FEditableFields.Assign(AEditableFields);
    { Suppression des champs de la clef }
    sClef := wMakeFieldString(TableName, ';') + ';';
    for i := Pred(FEditableFields.Count) downto 0 do
      if Pos(FEditableFields[i] + ';', sClef) > 0 then
        FEditableFields.Delete(i);
  end;
  if not FConsultation then
  begin
    EnableTom;
    FTreeTobSaved := Tob.Create(TableName, nil, -1);
    FTreeTobSaved.Dupliquer(FTreeTob, True, True, True)
  end;
  { S.A.V. }
  MnSavSepare.Visible := V_Pgi.Sav;
  MnSavShowTob.Visible := V_Pgi.Sav;
  MnSavShowTob2.Visible := not FConsultation and V_Pgi.Sav;
  if FConsultation then
    MnSavShowTob.Caption := 'S.A.V.';
  { Création de la liste }
  FListe := TWListe.Create(FDBListe, Self);
  { Tobs sur la DbListe }
  CreateTobsForListe;
  { Evénements grille }
  TreeGrid.PostDrawCell := Grid_OnDrawCell;
  TreeGrid.OnColumnWidthsChanged := Grid_OnColumnWidthsChanged;
  FCurrentRow := TreeGrid.FixedRows;
  FCurrentCol := TreeGrid.FixedCols;
  FInitialized := True;
  FDisplayTree := ADisplayTree;
  { Affichage de la tob en grille }
  DisplayTreeTob;
  ChargeImageList;
end;

{ Libère la FrameTreeTob dans une Form ----------------------------------------- }
destructor TFFrameTreeTob.Destroy;
begin
  inherited;
  if Assigned(FindDialog) then
    FindDialog.Free;
  if Assigned(FRatedFields) then
    FRatedFields.Free;
  { Libère les tobs de structure }
  FreeTobsForListe;
  { Libère la liste }
  if Assigned(FListe) then
    FListe.Free;
  if Assigned(FEditableFields) then
    FEditableFields.Free;
  if Assigned(FTreeTobSaved) then
    FTreeTobSaved.Free;
  { Libère la tom associée }
  DisableTom;
  { Libère l'objet de gestion écran }
  if Assigned(FTreeManager) then
    FTreeManager.Free;
end;

{ Permet de gérer le Panel des cumuls ------------------------------------------ }
procedure TFFrameTreeTob.InitPCumul;
var
  i: Integer;
  procedure CreateNewCumul(Field: TWFieldListe);
  begin
    with THPanel.Create(PCumul) do
    begin
      Parent := PCumul;
      Name := 'Pn' + wStringToAlphaNumString(Field.Name, '_');
      Align := alNone;
      Alignment := Field.Alignment;
      BevelOuter := bvNone;
      BorderStyle := bsSingle;
      UpdateCumulField(Field.Name);
      Color := clWindow;
      Ctl3D := False;
      Enabled := False;
      ParentFont := True;
      ParentCtl3D := False;
      Visible := True
    end
  end;
begin
  PCumul.Visible := False;
  try
    while PCumul.ControlCount > 0 do
      PCumul.Controls[0].Free;
    for i := 0 to Pred(FListe.ChampsCount) do
      if (FTreeTob.Detail.Count > 0) and FTreeTob.FieldExists(FListe.Champs[i].Name) and FListe.Champs[i].Cumul then
        CreateNewCumul(FListe.Champs[i]);
  finally
    PCumul.Visible := True
  end
end;

{ Gère la position des Cumuls sur le Panel "PCumul" ---------------------------- }
procedure TFFrameTreeTob.Grid_OnColumnWidthsChanged(Sender: TObject);
var
  FieldName: String;
  PnC      : THPanel;
  i, iCol  : Integer;
  Rect     : TRect;
begin
  { Gestion de la position des cumuls }
  if PCumul.ControlCount > 0 then
  begin
    for i := 0 to Pred(PCumul.ControlCount) do
    begin
      if PCumul.Controls[i] is THPanel then
      begin
        PnC := THPanel(PCumul.Controls[i]);
        FieldName := PnC.Name;
        Delete(FieldName, 1, 2);
        iCol := ColField[FieldName];
        if iCol = -1 then
          iCol := ColField['(' + FieldName + ')']; //+> '()' pour gérer les formules "dures"
        if iCol <> -1 then
        begin
          Rect := TreeGrid.CellRect(iCol, TreeGrid.FixedRows);
          PnC.Visible := False;
          try
            PnC.Top := -1;
            PnC.Left := Rect.Left;
            PnC.Width := Rect.Right - Rect.Left + TreeGrid.GridLineWidth;
            PnC.Height := PnC.Parent.Height;
          finally
            PnC.Visible := True;
          end
        end
      end
    end
  end;
  { Gestion de la "CheckBox" }
  if PnCbBooleanField.Visible then
    FitCheckBoxToCell
end;

{ Mise à jour du total des lignes ---------------------------------------------- }
procedure TFFrameTreeTob.UpdateTotal;
begin
  PCumul.Caption := TraduireMemoire(' Totaux') + ' (' + IntToStr(TreeGrid.RowCount - TreeGrid.FixedRows)
                  + ' ' + TraduireMemoire('lignes') + ')'
end;

{ Mise à jour du cumul pour les champs de type numérique ----------------------- }
procedure TFFrameTreeTob.UpdateCumuls(const iRow: Integer; const ForceCalcul: Boolean = False);
var
  wField: TWFieldListe;
  i     : Integer;
begin
  for i := TreeGrid.FixedCols to Pred(TreeGrid.ColCount) do
  begin
    wField := FListe.ChampsByName[TreeGrid.ColNames[i]];
    if Assigned(wField) and (TreeGrid.ColTypes[i] = 'R') 
    and (ForceCalcul or TobRow[iRow].IsFieldModified(wField.Name))
    and wField.Visible and wField.Cumul then
      UpdateCumulField(wField.Name)
  end;
end;

{ Mise à jour du caption du cumul concerné par le champ ------------------------ }
procedure TFFrameTreeTob.UpdateCumulField(const FieldName: String);
var
  i: Integer;
  Found: Boolean;
  wField: TWFieldListe;
  Pn: THPanel;
begin
  Found := False;
  i := 0;
  while not Found and (i < PCumul.ControlCount) do
  begin
    Found := UpperCase(PCumul.Controls[i].Name) = 'PN' + wStringToAlphaNumString(FieldName, '_');
    if not Found then
      Inc(i)
  end;
  if Found then
  begin
    wField := FListe.ChampsByName[FieldName];
    if Assigned(wField) then
    begin
      Pn := THPanel(PCumul.Controls[i]);
      Pn.Visible := False;
      try
        Pn.Caption := FormatFloat(StringReplace(wField.Format, ' ', ',', [rfIgnoreCase, rfReplaceAll]),
                                  FTreeTob.Somme(FieldName, [''], [''], True)
                                  + iif(FShowRoot, FTreeTob.GetDouble(FieldName), 0));
      finally
        Pn.Visible := True
      end
    end
  end
end;

{ Permet de savoir si la ligne "iRow" est éditable ou non ----------------------
  selon la propriété "UneditableConditions" ------------------------------------ }
function TFFrameTreeTob.GetRowEditable(iRow: Integer): Boolean;
var
  T: Tob;
  i: Integer;
begin
  T := TobRow[iRow];
  Result := Assigned(T);
  i := 0;
  while Result and (i < Length(FUneditableConditions)) do
  begin
    Result := T.GetValue(FUneditableConditions[i][0]) = FUneditableConditions[i][1];
    Inc(i)
  end
end;

{ Reporte la valeur du "CheckBox" dans la cellule en cours --------------------- }
procedure TFFrameTreeTob.CbBooleanFieldClick(Sender: TObject);
begin
  TreeGrid.CellValues[TreeGrid.Col, TreeGrid.Row] := BoolToStr_(TCheckBox(Sender).Checked);
  SaveCell2Tob(FCurrentCol, FCurrentRow, CurrentTob);
  DoControlField(CurrentTob, CurrentField);
  if FLastError = 0 then
    SaveCell2Tob(FCurrentCol, FCurrentRow, FTreeTobBeforeModif, False);
end;

{ Permet d'ajuster position et taille de la "CheckBox" en fonction de la cellule }
procedure TFFrameTreeTob.FitCheckBoxToCell;
var
  R: TRect;
  wField: TWFieldListe;
  function GetLeft(AAlign: TAlignment): Integer;
  begin
    case AAlign of
      taLeftJustify : Result := R.Left + Grid_GetLineWidth + 2;
      taRightJustify: Result := R.Right - Grid_GetLineWidth - FCbBooleanField.Width - 2;
      taCenter      : Result := GetLeft(taLeftJustify)
                              + ((TreeGrid.ColWidths [TreeGrid.Col] - PnCbBooleanField.Width ) div 2)
    else
      Result := 0
    end
  end;
const
  MaxCbSize = 14;
begin
  if PnCbBooleanField.Visible then
  begin
    R := TreeGrid.CellRect(TreeGrid.Col, TreeGrid.Row);
    FCbBooleanField.Width  := Min(13, R.Right - R.Left);
    FCbBooleanField.Height := FCbBooleanField.Width;
    PnCbBooleanField.Top   := R.Top  + Grid_GetLineHeight + ((TreeGrid.RowHeights[TreeGrid.Row] - PnCbBooleanField.Height) div 2) + 1;
    wField := FListe.ChampsByName[CurrentField];
    if Assigned(wField) then
      PnCbBooleanField.Left := GetLeft(wField.Alignment)
    else
      PnCbBooleanField.Left := GetLeft(taCenter);
  end
end;

{ Gère le form resize ---------------------------------------------------------- }
procedure TFFrameTreeTob.Form_Resize(Sender: TObject);
begin
  UnassignNotifyEvent(FForm.OnResize);
  try
    { Resize de la Form }
    if Assigned(FNotifyEvent) then
      FNotifyEvent(Sender);
    { Resize de la Frame }
    FrameResize(Self)
  finally
    ReassignNotifyEvent(FForm.OnResize)
  end
end;

{ Compte le nombre d'instances de TFFrameTreeTob sur la forme parente ---------- }
function TFFrameTreeTob.Form_CountFrameTreeTob: Integer;
var
  i: Integer;
begin
  Result := 0;
  for i := 0 to Pred(FForm.ComponentCount) do
    if FForm.Components[i] is TFFrameTreeTob then
      Inc(Result)
end;

{ Réajustements graphique lors du retaillage de la TFrame ---------------------- }
procedure TFFrameTreeTob.FrameResize(Sender: TObject);
begin
  ResizeGrid();
end;

{ Supression du noeud courant -------------------------------------------------- }
procedure TFFrameTreeTob.DeleteRow;
var
  Cancel  : Boolean;
  iRow, i : Integer;
  T       : Tob;
  wRapport: TWRapport;
begin
  iRow := TreeGrid.Row;
  if not FConsultation then
  begin
    if not FShowRoot or (iRow > TreeGrid.FixedRows) then
    begin
      Cancel := False;
      if Assigned(FOnBeforeDeleteRecord) then
        FOnBeforeDeleteRecord(Cancel);
      if not Cancel then
      begin
        { Suppression lignes }
        BeginUpdate;
        try
          T := CurrentTob;
          { Suppression lignes }
          for i := iRow + wTobCountDetails(T) downto iRow do
          begin
            TreeGrid.DeleteRow(i);
            FTreeManager.Delete(i);
          end;
          { Suppression data }
          DeleteTob(T);
          { Recalcul le iRow = Ligne qui prendra le focus }
          if iRow >= TreeGrid.RowCount then
            iRow := Pred(TreeGrid.RowCount);
          { Cumul }
          UpdateCumuls(iRow, True);
          { Totaux }
          UpdateTotal;
          FCurrentRow  := iRow;
        finally
          EndUpdate
        end;
        { Focus Cellule }
        Grid_SetFocus(TreeGrid.Col, iRow);
        if PnCbBooleanField.Visible then
          FitCheckBoxToCell;
        if Assigned(FOnAfterDeleteRecord) then
          FOnAfterDeleteRecord()
      end
    end
  end else begin
    if Assigned(FOnBeforeDeleteRecord) then
      FOnBeforeDeleteRecord(Cancel);
    wRapport := TWRapport.Create('Rapport de suppression');
    try
      wDeleteTable(TableName, Where, True, wRapport);
      wRapport.Display
    finally
      wRapport.Free
    end;
    if Assigned(FOnAfterDeleteRecord) then
      FOnAfterDeleteRecord();
    RefreshTreeTob
  end
end;

{ Ajout d'une ligne (Child = en sous niveau) ----------------------------------- }
procedure TFFrameTreeTob.AddRow(const Child: Boolean);
var
  T, TobMere: Tob;
  iRow, iRowMere: Integer;
begin
  if not FConsultation then
  begin
    if RecordIsValid then
    begin
      if Child then
        TobMere := CurrentTob
      else
        TobMere := CurrentTob.Parent;  
      if Assigned(TobMere) then
      begin
        { *** Gestion côté DONNEE *** }
        { Ajout de la tob }
        T := Tob.Create(TableName, TobMere, -1);

        { On passe dans le OnNewRecord de la Tom métier }
        if FTomEnabled then
          FTom.InitTOB(T);
        T.AddChampSupValeur('IKC', 'C');
        { *** Gestion côté AFFICHAGE *** }
        BeginUpdate;
        try
          { Si le noeud courant est fermé, on l'étend }
          if Child and not NodeExpanded then
            ExpandNode(False);
          { Insertion de la ligne }
          if Child then
            iRowMere := TreeGrid.Row
          else
            iRowMere := TreeGrid.Row - AbsoluteParentIndex[TobRow[TreeGrid.Row]];
          iRow := CountDetails[iRowMere] + iRowMere;
          if iRow > Pred(TreeGrid.RowCount) then
          begin
            TreeGrid.RowCount := TreeGrid.RowCount + 1;
            iRow := Pred(TreeGrid.RowCount);
            FTreeManager.Add(T)
          end
          else
          begin
            TreeGrid.InsertRow(iRow);
            FTreeManager.Insert(iRow, T)
          end;
          { Affichage des données dans la ligne }
          FillTreeGridLine(ItemManagerRow[iRow], iRow);
        finally
          EndUpdate
        end;
        TreeGrid.Row := iRow;
        TreeGrid.Col := TreeGrid.Col;
        Grid_SetFocus(TreeGrid.Col, TreeGrid.Row);
        { MàJ écran }
        if PnCbBooleanField.Visible then
          FitCheckBoxToCell;
        if Assigned(FOnNewRecord) then
          FOnNewRecord();
        { Cumul }
        UpdateCumuls(iRow, True);
        { Totaux }
        UpdateTotal;
      end
    end
  end
  else
    SavError('Appel en mode consultation impossible !', 'TFFrameTreeTob.AddRow');
end;

{ Annule complètement la saisie, on revient à la tob d'init -------------------- }
procedure TFFrameTreeTob.Undo;
begin
  if not FConsultation then
  begin
    BeginUpdate;
    try
      FTreeTob.ClearDetail;
      FTreeTob.InitValeurs;
      FTreeTob.Dupliquer(FTreeTobSaved, True, True, True);
      DisableRefresh;
      try
        RefreshTreeTob
      finally
        EnableRefresh
      end
    finally
      EndUpdate
    end
  end
  else
    SavError('Appel en mode consultation impossible !', 'TFFrameTreeTob.Undo')
end;

{ On ne passe pas par les méthodes de la Tom ----------------------------------- }
procedure TFFrameTreeTob.DisableTom;
begin
  FTomEnabled := False;
  if Assigned(FTom) then
    FTom.Free
end;

{ On passe par les méthodes de la Tom ------------------------------------------ }
procedure TFFrameTreeTob.EnableTom;
begin
  FTomEnabled := True;
  FTom := TTreeTom(CreateTom(TableName, nil, False, True));
end;

{ Passe par le OnControlRecord et ramène la validité --------------------------- }
function TFFrameTreeTob.RecordIsValid(const ShowMessages: Boolean): Boolean;
var
  LastErrorFieldName: String;
begin
  if not FConsultation then
  begin
    if IsEditableField(FieldCol[FCurrentCol]) then
      SaveCell2Tob(FCurrentCol, FCurrentRow, CurrentTob);
    DoControlRecord(LastErrorFieldName, ShowMessages);
    Result := FLastError = 0;
    if not Result and (LastErrorFieldName <> '') then
      GridSetFocusCol(LastErrorFieldName)
    else
      SaveCell2Tob(FCurrentCol, FCurrentRow, FTreeTobBeforeModif, False);
  end
  else
    Result := True
end;

{ Ramène la tob associée à la ligne -------------------------------------------- }
function TFFrameTreeTob.GetTobRow(iRow: Integer): Tob;
begin
  if wBetween(iRow, TreeGrid.FixedRows, Pred(TreeGrid.RowCount)) and Assigned(ItemManagerRow[iRow]) then
    Result := ItemManagerRow[iRow].TobLine
  else
    Result := nil
end;

function TFFrameTreeTob.GetRowTob(ATob: Tob): Integer;
var
  i: Integer;
begin
  Result := -1;
  i := FirstRow;
  while (Result = -1) and (i <= LastRow) do
  begin
    if Assigned(ItemManagerRow[i]) and (ItemManagerRow[i].TobLine = ATob) then
      Result := i;
    Inc(i)
  end
end;

{ Ramène l'objet permettant de gérer une ligne---------------------------------- }
function TFFrameTreeTob.GetItemManagerRow(iRow: Integer): TTreeManagerItem;
begin
  if wBetween(iRow, FirstRow, LastRow) then
    Result := TTreeManagerItem(TreeGrid.Objects[0, iRow])
  else
    Result := nil
end;

{ Ramène le nom du champ associé à la colone ----------------------------------- }
function TFFrameTreeTob.GetFieldCol(iCol: Integer): String;
begin
  if wBetween(iCol, TreeGrid.FixedCols, Pred(TreeGrid.ColCount)) then
    Result := TreeGrid.ColNames[iCol]
  else
    Result := '';
end;

{ Ramène le numéro de la colonne correspondant au nom de champ indiqué --------- }
function TFFrameTreeTob.GetColField(FieldName: String): Integer;
var
  i: Integer;
begin
  FieldName := UpperCase(FieldName);
  Result := -1;
  i := 0;
  while (Result = -1) and (i < TreeGrid.ColCount) do
  begin
    if FieldName = UpperCase(TreeGrid.ColNames[i]) then
      Result := i
    else
      Inc(i)
  end
end;

{ Désactive la gestion des contrôles de données -------------------------------- }
procedure TFFrameTreeTob.DisableControl;
begin
  Inc(FDoControl)
end;

{ Réactive la gestion des contrôles de données --------------------------------- }
procedure TFFrameTreeTob.EnableControl;
begin
  FDoControl := Max(0, FDoControl - 1)
end;

{ Désactive le rafraîchissement ------------------------------------------------ }
procedure TFFrameTreeTob.DisableRefresh;
begin
  Inc(FDoRefresh)
end;

{ Réactive le rafraîchissement ------------------------------------------------- }
procedure TFFrameTreeTob.EnableRefresh;
begin
  FDoRefresh := Max(0, FDoRefresh - 1)
end;

{ Prévient la mise à jour de la grille ----------------------------------------- }
procedure TFFrameTreeTob.BeginUpdate;
begin
  if FDoUpdate = 0 then
  begin
    TreeGrid.BeginUpdate;
    TreeGrid.Enabled := False;
    TreeGrid.SynEnabled := False
  end;
  Inc(FDoUpdate)
end;

{ Actualise la mise à jour de la grille ---------------------------------------- }
procedure TFFrameTreeTob.EndUpdate;
begin
  FDoUpdate := Max(0, FDoUpdate - 1);

  if FDoUpdate = 0 then
  begin
    TreeGrid.Enabled := True;
    TreeGrid.SynEnabled := True;
    TreeGrid.EndUpdate;
    RefreshBtExpand
  end
end;

{ Ramène l'index absolu d'un tob par rapport à son parent direct --------------- }
function TFFrameTreeTob.GetAbsoluteParentIndex(T: Tob): Integer;
  function _GetAbsoluteParentIndex(T: Tob): Integer;
  var
    i: Integer;
  begin
    Result := 0;
    for i := 0 to Pred(T.Detail.Count) do
      Result := Result + 1 + _GetAbsoluteParentIndex(T.Detail[i])
  end;
var
  i: Integer;
begin
  if Assigned(T.Parent) then
  begin
    Result := 0;
    for i := 0 to T.GetIndex do
    begin
      Inc(Result);
      if i < T.GetIndex then
        Result := Result + _GetAbsoluteParentIndex(T.Parent.Detail[i])
    end
  end
  else
    Result := 0
end;

{ Met d'un champ --------------------------------------------------------------- }
procedure TFFrameTreeTob.DoControlField(T: Tob; const FieldName: String; const ShowMessages: Boolean);
var
  IM: TTreeManagerItem;
begin
  { Init. de l'id error }
  FLastError := 0;
  if FTomEnabled then
  begin
    FTom.LastError := 0;
    FTom.LastErrorMsg := '';
  end;    
  if (FDoControl = 0) and (T.GetValue(FieldName) <> FTreeTobBeforeModif.GetValue(FieldName)) then
  begin
    { OnControlField métier }
    if FTomEnabled then
    begin
 //@@GGU     FTom.ControlField(FieldName);  //Remplacé par :
      if (FTom.LastError <> 0) and Assigned(FTom.Ecran) then
        FTom.SetFocusControl(FieldName);
      if FTom.LastError > 0 then
        FLastError := ErrTomControlField;
    end;

    { OnControlField local }
    if (FLastError = 0) and Assigned(FOnControlField) then
      FOnControlField(FieldName);
           
    { Contrôles des données passés...  }
    if FLastError = 0 then
    begin
      { OnCalculField métier }
//@@GGU      if FTomEnabled then
//@@GGU        FTom.CalculField(FieldName);

      { OnCalculField local }
      if Assigned(FOnCalculField) then
        FOnCalculField(FieldName);

      IM := FTreeManager.IndexOf(T);
      FillTreeGridLine(IM, IM.RowIndex);
    end
    else if ShowMessages and FTomEnabled and (FLastError = ErrTomControlField) and (FTom.LastErrorMsg <> '') then
      PgiError(FTom.LastErrorMsg)
  end;
end;

{ Effectue un control sur toute un enregistrement ------------------------------ }
procedure TFFrameTreeTob.DoControlRecord(var LastErrorFieldName: String; const ShowMessages: Boolean);
begin
  FLastError := 0;

  if ColEditable[FCurrentCol] then
    DoControlField(TobRow[FCurrentRow], FieldCol[FCurrentCol], ShowMessages);
  if (FDoControl = 0) and (FLastError = 0) then
  begin
    if TobRow[FCurrentRow].IsOneModifie(False) then
    begin
      { RecordIsValid }
      if FTomEnabled then
      begin
        FTom.FTob := TobRow[FCurrentRow];
//@@GGU        FTom.RecordIsvalid;  //Remplacé par
//debut
//  { Init des champs }
//  DisableControl;
//  try
//    wSetInitFields(True);
//  finally
//    EnableControl;
//  end;
//  { Vérification des champs obligatoires }
//  if not ( VerifChampsObligatoires and VerifInfosLibres) and
  if FTom.LastError > 0 then
  begin
    if not Assigned(FTom.Ecran) then
      FTom.FTob.AddChampSupValeur('Error', FTom.LastErrorMsg, False);
  end;
//Fin

        if FTom.LastError > 0 then
        begin
          FLastError := ErrTomControlField;
          if ShowMessages and (FTom.LastErrorMsg <> '') then
            PgiError(FTom.LastErrorMsg)
        end
      end;

      if (FLastError = 0) and Assigned(FOnControlRecord) then
        FOnControlRecord(LastErrorFieldName)
    end
  end
end;

{ Permet d'effectuer tous les contrôles si l'on sort de la grille -------------- }
procedure TFFrameTreeTob.TreeGridExit(Sender: TObject);
var
  LastErrorFieldName: String;
begin
  if not FConsultation and Assigned(FForm.ActiveControl)
  and (Pos('GRIDCOMP_', FForm.ActiveControl.Name) = 0) then
  begin
    SaveCell2Tob(FCurrentCol, FCurrentRow, CurrentTob);
    DoControlRecord(LastErrorFieldName);
    if FLastError > 0 then
    begin
      if LastErrorFieldName = '' then
        LastErrorFieldName := CurrentField;

      if TreeGrid.CanFocus then
        TreeGrid.SetFocus;
      GridSetFocusCol(LastErrorFieldName)
    end
    else
      SaveCell2Tob(FCurrentCol, FCurrentRow, FTreeTobBeforeModif, False)
  end
end;

{ Init. à la focalisation ------------------------------------------------------ }
function TFFrameTreeTob.Grid_SetFocus(const iCol, iRow: Integer): Boolean;
var
  Cancel: Boolean;
  varCol, varRow: Integer;
begin
  if TreeGrid.CanFocus then
    TreeGrid.SetFocus;
    
  Cancel := False;
  varRow := iRow;
  TreeGridRowEnter(TreeGrid, varRow, Cancel, False);
  Result := not Cancel;

  if Result and not FConsultation then
  begin
    varCol := iCol;
    if not Cancel then
      TreeGridCellEnter(TreeGrid, varCol, varRow, Cancel);
    if TreeGrid.InplaceEditor.CanFocus then
      TreeGrid.InplaceEditor.SetFocus;

    Result := not Cancel
  end
  else
    Result := True
end;

{ Annule la saisie en cours ( Ferme la fiche )----------------------------------- }
procedure TFFrameTreeTob.Cancel;
begin
  if not FConsultation then
  begin
    FTreeTob.ClearDetail;
    FTreeTob.Dupliquer(FTreeTobSaved, False, True, True);
  end
end;

{ Valide la saisie en cours. Si DoTransactions, on gère la base de donnée ------ }
function TFFrameTreeTob.Valid(const DoTransactions: Boolean): Boolean;
var
  Transac: TTreeTransac;
begin
  Result := FTreeTob.Detail.Count > 0;
  if Result and not FConsultation and DoTransactions then
  begin
    TreeGridExit(TreeGrid);
    Transac := TTreeTransac.Create(Self);
    try
      Transactions(Transac.InsertOrUpdate, 0);
      Result := V_Pgi.IOError = oeOk
    finally
      Transac.Free
    end
  end
end;

{ Supprime une tob : appelé depuis la suppression de ligne de grille ----------- }
procedure TFFrameTreeTob.DeleteTob(T: Tob);
var
  i: Integer;
  ArrFields: MyArrayField;
  ArrValues: MyArrayValue;
begin
  if Ikc <> 'C' then
  begin
    SetLength(ArrFields, 0);
    ArrFields := wStringToArray(wMakeFieldString(TableName, ';'), ';');
    SetLength(ArrValues, Length(ArrFields));
    for i := Low(ArrFields) to High(ArrFields) do
      ArrValues[i] := T.GetValue(ArrFields[i]);
    if not Assigned(FTob2Delete.FindFirst(ArrFields, ArrValues, True)) then
      T.ChangeParent(FTob2Delete, -1)
  end
  else
    T.Free
end;

{ Initialise un "Ikc" pour toutes les tobs ------------------------------------- }
procedure TFFrameTreeTob.AddIkc(T: Tob);
var
  i: Integer;
begin
  T.AddChampSupValeur('IKC', iif(FConsultation, 'M', 'C'));
  for i := 0 to Pred(T.Detail.Count) do
    AddIkc(T.Detail[i])
end;

{ Ramène le "Ikc" de la saisie, on considère que c'est celui de la tob mère ---- }
function TFFrameTreeTob.GetIkc: Char;
begin
  Result := FTreeTob.GetString('IKC')[1]
end;

procedure TFFrameTreeTob.InitCellFound;
begin
  FCellFound := Point(-1, -1)
end;

{ TTreeGridCheckBox } { - } { Classe héritant de TCheckBox afin de gérer l'événement KeyDown }

procedure TTreeGridCheckBox.WMKeyDown(var Message: TWMKeyDown);
begin
  FGrid.SetFocus;
  SendMessage(FGrid.Handle, WM_KEYDOWN, Message.CharCode, Message.KeyData);
end;

{ TTreeTransac } { - } { Classe d'exécution des transactions }

constructor TTreeTransac.Create(const AFrameTreeTob: TFFrameTreeTob);
begin
  FTT := AFrameTreeTob;
end;

{ Recurse sur la TreeTob afin d'enregistrer chaque Tob et d'effectuer
  les contrôles de la Tom ------------------------------------------------------ }
procedure TTreeTransac.DoTreeTrans(const T: Tob);
var
  i: Integer;
  Cancel: Boolean;

  { Insert }
  function InsertUniqueTob: Boolean;
  var
    TobToInsert: Tob;
  begin
    TobToInsert := Tob.Create(FTT.TableName, nil, -1);
    try
      TobToInsert.Dupliquer(T, False, True, True);
      Result := TobToInsert.InsertDB(nil)
    finally
      TobToInsert.Free
    end
  end;

  { Update }
  function UpdateUniqueTob: Boolean;
  var
    TobToUpdate: Tob;
  begin
    TobToUpdate := Tob.Create(FTT.TableName, nil, -1);
    try
      TobToUpdate.Dupliquer(T, False, True);
      TobToUpdate.SetAllModifie(True);
      Result := TobToUpdate.UpdateDB
    finally
      TobToUpdate.Free
    end
  end;

begin
  wMoveProgressForm();

  { BeforeUpdate }
  if FTT.FTom.VerifTOB(T) then
  begin
    { Insert ou Update }
    if T.GetString('IKC') = 'C' then
      Cancel := not InsertUniqueTob
    else
      Cancel := not UpdateUniqueTob;

    { AfterUpdate }
    if not Cancel then
      FTT.FTom.AfterVerifTOB(T)
  end;

  { Récurse }
  for i := 0 to Pred(T.Detail.Count) do
    DoTreeTrans(T.Detail[i])
end;

{ Gestion DB ------------------------------------------------------------------- }
procedure TTreeTransac.InsertOrUpdate;
var
  i: Integer;
begin
  wInitProgressForm(nil, TraduireMemoire('Enregistrement en cours...'), '',
                    wTobCountDetails(FTT.FTreeTob) + iif(FTT.ShowRoot, 1, 0), False, True);
  try
    { Suppressions }
    for i := 0 to Pred(FTT.FTob2Delete.Detail.Count) do
      if FTT.FTom.DeleteTOB(FTT.FTob2Delete.Detail[i]) then
        FTT.FTob2Delete.Detail[i].DeleteDB;

    { Insert ou Update }
    if FTT.ShowRoot then
    begin
      { Test du premier enregistrement }
      if FTT.FTreeTob.ExistDB then
        FTT.FTreeTob.SetString('IKC', 'M');
      { Enregistrement }
      DoTreeTrans(FTT.FTreeTob);
    end
    else
    begin
      for i := 0 to Pred(FTT.FTreeTob.Detail.Count) do
        DoTreeTrans(FTT.FTreeTob.Detail[i])
    end
  finally
    wFiniProgressForm
  end;
end;

{ Ramène un where en fonction de la sélection de la grille --------------------- }
function TFFrameTreeTob.GetWhere: String;
var
  i: Integer;

  function AddToWhere(const T: Tob): String;
  var
    sClef1: String;

    function GetFieldsValues(sClef: String): String;
    begin
      Result := '';
      while sClef <> '' do
        Result := Result + iif(Result = '', '', ';') + T.GetString(ReadTokenSt(sClef));
    end;

  begin
    if Assigned(T) then
    begin
      sClef1 := wMakeFieldString(T.NomTable, ';');
      Result := wMakeWhereSQL(sClef1, GetFieldsValues(sClef1));
    end
  end;

begin
  Result := '';
  if TreeGrid.MultiSelect and (TreeGrid.nbSelected > 0) then
  begin
    for i := TreeGrid.FixedRows to Pred(TreeGrid.RowCount) do
      if TreeGrid.IsSelected(i) then
        Result := Result + iif(Result <> '', ' OR ', '') + '(' + AddToWhere(TobRow[i]) + ')'
  end
  else
    Result := '(' + AddToWhere(CurrentTob) + ')'
end;

{ Grille en multisélection ? --------------------------------------------------- }
function TFFrameTreeTob.GetMultiSelected: Boolean;
begin
  Result := (TreeGrid.NbSelected > 0) or (TreeGrid.AllSelected)
end;

{ Première ligne de la grille -------------------------------------------------- }
function TFFrameTreeTob.GetFirstRow: Integer;
begin
  Result := TreeGrid.FixedRows
end;

{ Dernière ligne de la grille -------------------------------------------------- }
function TFFrameTreeTob.GetLastRow: Integer;
begin
  Result := Pred(TreeGrid.RowCount)
end;

{ Ligne sélectionnée ? --------------------------------------------------------- }
function TFFrameTreeTob.GetRowSelected(iRow: Integer): Boolean;
begin
  Result := TreeGrid.IsSelected(iRow);
end;

{ Colonne éditable ? ----------------------------------------------------------- }
function TFFrameTreeTob.GetEditableCol(iCol: Integer): Boolean;
begin
  Result := TreeGrid.ColEditables[iCol] and IsEditableField(TreeGrid.ColNames[iCol])
end;

function TFFrameTreeTob.GetSelectionEnabled: Boolean;
begin
  Result := Assigned(FBtSelectAll) or Assigned(FBtSelectNode)
end;

function TFFrameTreeTob.GetColVisible(iCol: Integer): Boolean;
begin
  if iCol > -1 then
    Result := TreeGrid.ColWidths[iCol] > -1
  else
    Result := False
end;

procedure TFFrameTreeTob.SetColVisible(iCol: Integer; const Value: Boolean);
begin
  if iCol > -1 then
  begin
    if Value then
      TreeGrid.ColWidths[iCol] := FListe.ChampsByName[FieldCol[iCol]].Width
    else
      TreeGrid.ColWidths[iCol] := -1
  end
end;

{ Désélectionne les lignes ----------------------------------------------------- }
procedure TFFrameTreeTob.Unselect(const iRow: Integer);
begin
  if Assigned(FBtSelectAll) then
  begin
    UnassignNotifyEvent(FBtSelectAll.OnClick);
    try
      if TreeGrid.MultiSelect and FConsultation then
      begin
        if iRow = -1 then
          TreeGrid.AllSelected := False
        else
          TreeGrid.FlipSelection(iRow);

        if FBtSelectAll.Down then
          FBtSelectAll.Down := False
      end
    finally
      ReassignNotifyEvent(FBtSelectAll.OnClick)
    end
  end;

  if Assigned(FBtSelectNode) then
  begin
    UnassignNotifyEvent(FBtSelectNode.OnClick);
    try
      if FBtSelectNode.Down then
        FBtSelectNode.Down := False
    finally
      ReassignNotifyEvent(FBtSelectNode.OnClick);
    end
  end
end;

function TFFrameTreeTob.isEmpty: Boolean;
begin
  Result := not ShowRoot and (FTreeTob.Detail.Count = 0);
end;

procedure TFFrameTreeTob.ReassignNotifyEvent(Event: TNotifyEvent);
begin
  FNotifyEvent := Event;
  Event := nil;
end;

procedure TFFrameTreeTob.UnassignNotifyEvent(Event: TNotifyEvent);
begin
  Event := FNotifyEvent;
  FNotifyEvent := nil;
end;

procedure TFFrameTreeTob.ResizeGrid;
begin
  PCumul.Visible := False;
  try
    { Resize des colonnes de la grille }
    if HMTrad.ActiveResize then
      HMTrad.ResizeGridColumns(TreeGrid);
  finally
    PCumul.Visible := True;
  end;

  { Ajustement de la "CheckBox" }
  FitCheckBoxToCell;
end;

procedure TFFrameTreeTob.Debug;
begin
  if V_Pgi.Sav then
    TobDebug(FTreeTob)
end;

{ TTreeManagerItem } { - } { Classe de définition des items parallèles aux tobs permettant de... }
                           { ...gérer l'affichage ("+", Coordonnées du "+", Ligne associée à la tob) }

{ Initialisation des données : La tob associée, L'état du noeud, L'indice ligne dans la grille,
  les numéros d'images permettant de dessiner l'arborescence }
constructor TTreeManagerItem.Create(Parent: TTreeManager; const ATob: Tob; const ItemIndex: Integer = -1);
var
  i: Integer;
  iLevel: Integer;
  PrevIM: TTreeManagerItem;

  function GetFirstLevel: Integer;
  begin
    Result := iif(FManager.ShowRoot, 0, 0)
  end;

begin
  FManager := Parent;
  FTob := ATob;
  FExpanded := IsParent;
  FVisible := True;
  if ItemIndex = -1 then
    FIndex := FManager.CountItems
  else
  begin
    FIndex := ItemIndex;
    for i := ItemIndex to Pred(FManager.CountItems) do
    begin
      Inc(FManager[i].FIndex);
      Inc(FManager[i].FRowIndex)
    end
  end;
  ClearImages();
  FRowIndex := FIndex + FManager.FRowOffSet;

  { Images par niveau }
  iLevel := Pred(Level);
  if iLevel > Pred(GetFirstLevel()) then
  begin
    PrevIM := FManager[Pred(FIndex)];
    for i := GetFirstLevel() to Pred(iLevel) do
    begin
      if (PrevIM.TreeImages[i] = tttiTEG) or (PrevIM.TreeImages[i] = tttiVER) then
        AddImage(tttiVER)
      else
        AddImage(tttiNone)
    end;

    { Image rattachée aux noeuds direct }
    if (FTob.GetIndex = Pred(FTob.Parent.Detail.Count)) then
      AddImage(tttiCBG)
    else
      AddImage(tttiTEG)
  end
end;

destructor TTreeManagerItem.Destroy;
var
  i: Integer;
begin
  inherited;

  ClearImages();

  { Décalage des index }
  for i := Succ(Index) to Pred(FManager.CountItems) do
  begin
    Dec(FManager[i].FIndex);
    Dec(FManager[i].FRowIndex)
  end
end;

{ Inverse l'état du noeud ------------------------------------------------------ }
procedure TTreeManagerItem.InvertExpandState;
begin
  FExpanded := not FExpanded;
end;

{ Nettoie les numéros d'images du tree view ------------------------------------ }
procedure TTreeManagerItem.ClearImages;
begin
  SetLength(FTreeImages, 0)
end;

{ Ajoute une image ------------------------------------------------------------- }
procedure TTreeManagerItem.AddImage(Image: TTreeTobTvImg);
begin
  SetLength(FTreeImages, Length(FTreeImages) + 1);
  FTreeImages[Pred(Length(FTreeImages))] := Image
end;

function TTreeManagerItem.NextItem: TTreeManagerItem;
begin
  Result := FManager[Succ(Index)]
end;

function TTreeManagerItem.NextItemVisible: TTreeManagerItem;
begin
  Result := NextItem;
  if Assigned(Result) and not Result.Visible then
    Result := Result.NextItemVisible
end;

function TTreeManagerItem.PrevItem: TTreeManagerItem;
begin
  Result := FManager[Pred(Index)]
end;

function TTreeManagerItem.PrevItemVisible: TTreeManagerItem;
begin
  Result := PrevItem;
  if Assigned(Result) and not Result.Visible then
    Result := Result.PrevItemVisible
end;

{ Compte le nombre d'images ---------------------------------------------------- }
function TTreeManagerItem.GetImagesCount: Integer;
begin
  Result := Length(FTreeImages)
end;

{ Ramène le numéro de ligne associé au TreeManagerItem ------------------------- }
function TTreeManagerItem.GetRowIndex: Integer;
begin
  Result := FRowIndex
end;

function TTreeManagerItem.GetLevel: Integer;
begin
  Result := FTob.Niveau - FManager.FTob.Niveau
end;

function TTreeManagerItem.GetParentIndex: Integer;
const
  LastIndex: Integer = -1;
  LastParentIndex: Integer = -1;
var
  iLevel: Integer;

  function _GetParentIndex(IM: TTreeManagerItem): Integer;
  begin
    if IM.Level = Pred(iLevel) then
      Result := IM.Index
    else if IM.Index > 0 then
      Result := _GetParentIndex(IM.FManager[Pred(IM.Index)])
    else
      Result := -1
  end;

begin
  if Index > 0 then
  begin
    if Index <> LastIndex then
    begin
      LastIndex := Index;
      iLevel := Level;
      Result := _GetParentIndex(Self);
      LastParentIndex := Result
    end
    else
      Result := LastParentIndex
  end
  else
    Result := -1
end;

function TTreeManagerItem.GetParent: TTreeManagerItem;
begin
  if ParentIndex > -1 then
    Result := FManager[ParentIndex]
  else
    Result := nil
end;

{ Ramène le premier noeud parent "fermé" }
function TTreeManagerItem.GetFirstParentCollapsed: TTreeManagerItem;
begin
  Result := Parent;
  if Assigned(Result) and (Result.Expanded or not Result.Visible) then
    Result := Result.FirstParentCollapsed
end;

function TTreeManagerItem.GetIsParent: Boolean;
begin
  Result := FTob.Detail.Count > 0
end;

{ Ramène la nème image --------------------------------------------------------- }
function TTreeManagerItem.GetTreeImage(Index: Integer): TTreeTobTvImg;
begin
  Result := FTreeImages[Index];
end;

{ Affecte la nème image -------------------------------------------------------- }
procedure TTreeManagerItem.SetTreeImage(Index: Integer; const Value: TTreeTobTvImg);
begin
  FTreeImages[Index] := Value
end;

{ Affecte les coordonnées du "+" du noeud -------------------------------------- }
procedure TTreeManagerItem.SetNodeCoord(const Value: TPoint);
begin
  FNodeCoord := Value;
end;

{ Change l'état Visible de l'item avec impactage sur les RowIndex -------------- }
procedure TTreeManagerItem.SetVisible(const Value: Boolean);

  procedure UpdateRowIndex(IM: TTreeManagerItem);
  begin
    if Value then
      Inc(IM.FRowIndex)
    else
      Dec(IM.FRowIndex)
  end;

var
  IM: TTreeManagerItem;
begin
  if FVisible <> Value then
  begin
    FVisible := Value;
    UpdateRowIndex(Self);
    IM := NextItem;
    while Assigned(IM) do
    begin
      UpdateRowIndex(IM);
      IM := IM.NextItem;
    end
  end
end;

{ TTreeManager } { - } { Classe de regroupement des items de gestion de l'affichage de tob }

constructor TTreeManager.Create(const ATob: Tob; const RowOffSet: Integer; const AShowRoot: Boolean);
begin
  FItems := TList.Create();
  ShowRoot := AShowRoot;
  FTob := ATob;
  FRowOffSet := RowOffSet;
end;

{ Ramène un IM en fonction d'une tob ------------------------------------------- }
function TTreeManager.IndexOf(T: Tob): TTreeManagerItem;
var
  i: Integer;
begin
  Result := nil;
  i := 0;
  while not Assigned(Result) and (i < CountItems) do
  begin
    if T = Items[i].TobLine then
      Result := Items[i]; 
    Inc(i)
  end
end;

function FindTextInTob(FromFieldIndex: Integer; const S: String; T: Tob; const Options: TFindOptions; out FoundFieldName: String): Boolean;

  function CompareFieldValue(SToFind, SToCompare: String): Boolean;
  begin
    if not (frMatchCase in Options) then
    begin
      SToCompare := UpperCase(SToCompare);
      SToFind := UpperCase(SToFind);
    end;

    Result := (    (frWholeWord in Options) and (SToFind = SToCompare))
           or (not (frWholeWord in Options) and (Pos(SToFind, SToCompare) > 0))
  end;

var
  j: Integer;
begin
  Result := False;

  if frDown in Options then
  begin
    if FromFieldIndex < 1 then
      FromFieldIndex := 1;

    j := FromFieldIndex;
    while not Result and (j <= T.NbChamps) do
    begin
      Result := CompareFieldValue(S, VarToStr(T.GetValeur(j)));
      Inc(j)
    end;

    if not Result then
    begin
      if j < 1000 then
        j := 1000;
      while not Result and (j < 1000 + T.ChampsSup.Count) do
      begin
        Result := CompareFieldValue(S, VarToStr(T.GetValeur(j)));
        Inc(j)
      end
    end;
  end
  else
  begin
    if FromFieldIndex > Pred(1000 + T.ChampsSup.Count) then
      FromFieldIndex := Pred(1000 + T.ChampsSup.Count);

    j := FromFieldIndex;
    while not Result and (j >= 1000) do
    begin
      Result := CompareFieldValue(S, VarToStr(T.GetValeur(j)));
      Dec(j)
    end;

    if not Result then
    begin
      if j > Pred(T.NbChamps) then
        j := T.NbChamps;
      while not Result and (j > 0) do
      begin
        Result := CompareFieldValue(S, VarToStr(T.GetValeur(j)));
        Dec(j)
      end
    end
  end;

  if Result then
    FoundFieldName := T.GetNomChamp(iif(frDown in Options, Pred(j), Succ(j)))
end;

{ Recherche d'une chaîne dans les tob de chaque TreeManagerItem ---------------- }
function TTreeManager.Find(const St: String; const Options: TFindOptions;
                           const WithAlert: Boolean; out FoundFieldName: String): TTreeManagerItem;
var
  IM: TTreeManagerItem;
begin
  if FLastIndex >= CountItems then
    FLastIndex := Pred(CountItems)
  else if FLastIndex < 0 then
    FLastIndex := 0;
  FoundFieldName := '';
  Result := nil;
  FFirstFind := False;
  while not Assigned(Result) and (((frDown in Options) and (FLastIndex < CountItems)) or (not (frDown in Options) and (FLastIndex >= 0))) do
  begin
    if FLastColFound = -1 then
      FLastColFound := iif(frDown in Options, 1, Pred(1000 + Items[FLastIndex].TobLine.ChampsSup.Count));
    IM := Items[FLastIndex];
    if FindTextInTob(FLastColFound, St, IM.TobLine, Options, FoundFieldName) then
      Result := IM
    else if frDown in Options then
      Inc(FLastIndex)
    else
      Dec(FLastIndex);
    if not Assigned(Result) then
      FLastColFound := -1
    else
      FLastColFound := IM.TobLine.GetNumChamp(FoundFieldName) + iif(frDown in Options, 1, -1)
  end;
  if not Assigned(Result) and WithAlert then
    PgiInfo(TraduireMemoire('Impossible de trouver') + ' "' + St + '".', TraduireMemoire('Recherche'))
end;

{ Init. de la structure -------------------------------------------------------- }
procedure TTreeManager.InitTreeManager;

  procedure InitTreeManager(T: Tob);
  var
    i: Integer;
  begin
    FItems.Add(TTreeManagerItem.Create(Self, T));
    for i := 0 to Pred(T.Detail.Count) do
      InitTreeManager(T.Detail[i])
  end;

var
  i: Integer;
begin
  Clear;
  if ShowRoot then
    InitTreeManager(FTob)
  else
    for i := 0 to Pred(FTob.Detail.Count) do
      InitTreeManager(FTob.Detail[i]);
end;

{ RAZ de la structure ---------------------------------------------------------- }
procedure TTreeManager.Clear;
begin
  while CountItems > 0 do
  begin
    TTreeManagerItem(FItems[Pred(CountItems)]).Free;
    FItems.Delete(Pred(CountItems))
  end
end;

{ Supprime le nème élément de la structure ------------------------------------- }
procedure TTreeManager.Delete(const ItemIndex: Integer);
begin
  TTreeManagerItem(FItems[ItemIndex]).Free;
  FItems.Delete(ItemIndex);
end;

{ Ajoute un élément dans la structure ------------------------------------------ }
procedure TTreeManager.Add(const T: Tob);
begin
  FItems.Add(TTreeManagerItem.Create(Self, T, -1))
end;

{ Insert un élément dans la structure ------------------------------------------ }
procedure TTreeManager.Insert(const ItemIndex: Integer; const T: Tob);
begin
  FItems.Insert(ItemIndex, TTreeManagerItem.Create(Self, T, ItemIndex))
end;

destructor TTreeManager.Destroy;
begin
  inherited;
  Clear;    
  if Assigned(FItems) then
    FItems.Free;
end;

{ Ramène le n_ième Item de la structure ------------------------------------------ }
function TTreeManager.GetTreeManagerItem(Index: Integer): TTreeManagerItem;
begin
  if wBetween(Index, 0, Pred(CountItems)) then
    Result := TTreeManagerItem(FItems[Index])
  else
    Result := nil
end;

function TTreeManager.GetCountItems: Integer;
begin
  Result := FItems.Count
end;

//PT1
procedure TFFrameTreeTob.PrintGrid;
var
  indexRow, indexNiveau : Integer;
  TabDecalages : Array of Integer;
  StDecalage : String;
  IM: TTreeManagerItem;
  Function DrawTvImgASCII(Ttti: TTreeTobTvImg; Var NbStDecal : Integer) : String;
  begin
    case Ttti of
      tttiCBG: begin
        result := '   ''-- ';
        NbStDecal := NbStDecal + 7;
      end;
      tttiTEG: begin
        result := '   ¦-- ';
        NbStDecal := NbStDecal + 7;
      end;
      tttiVER: begin
        result := '   ¦';
        NbStDecal := NbStDecal + 4;
      end;
      tttiNone: begin
        result := '    ';
        NbStDecal := NbStDecal + 4;
      end;
    end
  end;
begin
  SetLength(TabDecalages, TreeGrid.RowCount);
  BeginUpdate;
  try
    for indexRow := 0 to TreeGrid.RowCount -1 do
    begin
      StDecalage := '';
      if Assigned(TobRow[indexRow]) then
      begin
        IM := ItemManagerRow[indexRow];
        for indexNiveau := 0 to Pred(IM.CountImages) do
          StDecalage := StDecalage + DrawTvImgASCII(IM.TreeImages[indexNiveau], TabDecalages[indexRow]);
        TreeGrid.Cells[1, indexRow] := StDecalage + TreeGrid.Cells[1, indexRow];
      end;
    end;
{$IFNDEF EAGLCLIENT}
    PrintDBGrid(TreeGrid,Nil, Caption,'');
{$ELSE}
    PrintDBGrid('TreeGrid', '', Caption,'');  //GGU : Problème AGL : FQ 12726  non réglée au 19/10/2007
{$ENDIF}
    for indexRow := 0 to TreeGrid.RowCount -1 do
    begin
      TreeGrid.Cells[1, indexRow] := RightStr((TreeGrid.Cells[1, indexRow]), Length(TreeGrid.Cells[1, indexRow]) - TabDecalages[indexRow]);
    end;
  finally
    EndUpdate;
  end;                
  DisableRefresh;
  try
    RefreshTreeTob;
  finally
    EnableRefresh;
  end
end;

end.
