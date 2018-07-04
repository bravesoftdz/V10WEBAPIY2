unit UParametre;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Grids, ExtCtrls, ComCtrls, UTOB, Db,
  HPanel, StdCtrls,
  Mask,
{$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet, Variants, ADODB, {$ENDIF}
{$IFDEF EAGLCLIENT}
     UtileAGL,
{$ELSE}
     PrintDbG,
{$ENDIF}
  DBCtrls, HTB97, DBGrids, HEnt1, HSysMenu, hmsgbox, licUtil,
  Hctrls, Buttons, ImgList, 
{$IFDEF CISXPGI}
  uYFILESTD,
{$ENDIF}
  Spin;

type
  TFParametre = class(TForm)
    TV: TTreeView;
    Splitter1: TSplitter;
    Panel1: TPanel;
    DS1: TDataSource;
    ImageList1: TImageList;
    Panel2: TPanel;
    Panel_Champs: TPanel;
    Binsert: TToolbarButton97;
    BDelete: TToolbarButton97;
    BAnnuler: TToolbarButton97;
    BAide: TToolbarButton97;
    BOuvrir: TToolbarButton97;
    BRechercher: TToolbarButton97;
    BImprimer: TToolbarButton97;
    Splitter2: TSplitter;
    PCPied: TPageControl;
    Page1: TTabSheet;
    Page2: TTabSheet;
    MEContenu: TMaskEdit;
    TMEContenu: TLabel;
    TMELongueur: TLabel;
    TMEMajusMin: TLabel;
    TMEType: TLabel;
    TMEContenuoblig: TLabel;
    TMENomPGI: TLabel;
    TMECommentaire: TLabel;
    MEContenuoblig: TMaskEdit;
    MENomPGI: TMaskEdit;
    MECommentaire: TMaskEdit;
    BFirst: TToolbarButton97;
    BPrev: TToolbarButton97;
    BNext: TToolbarButton97;
    BLast: TToolbarButton97;
    bDefaire: TToolbarButton97;
    bPost: TToolbarButton97;
    bDupliquer: TToolbarButton97;
    MENumero: TMaskEdit;
    MELibre1: TMaskEdit;
    TMELibre1: TLabel;
    Label1: TLabel;
    MELibre2: TMaskEdit;
    Label2: TLabel;
    MELibre3: TMaskEdit;
    DBG: TDBGrid;
    FindDialog: TFindDialog;
    MELibre4: TMaskEdit;
    Label3: TLabel;
    MELibre5: TMaskEdit;
    Label4: TLabel;
    Label5: TLabel;
    Label6: TLabel;
    Label7: TLabel;
    MELibre6: TMaskEdit;
    MELibre7: TMaskEdit;
    MELibre8: TMaskEdit;
    MELibre9: TMaskEdit;
    Label8: TLabel;
    MELibre10: TMaskEdit;
    Label9: TLabel;
    MEObligatoire: TCheckBox;
    METypeAlpha: TComboBox;
    MEMajusMin: TComboBox;
    MEFige: TCheckBox;
    DS3: TDataSource;
    HMsgBox: THMsgBox;
    HMTrad: THSystemMenu;
    TMEDomaine: TLabel;
    TMETable: TLabel;
    TMENom: TLabel;
    METableName: TMaskEdit;
    MENom: TMaskEdit;
    MEModifiable: TCheckBox;
    SG1: THGrid;
    MEAlignement: TComboBox;
    Label10: TLabel;
    Label11: TLabel;
    MESeparateur: TEdit;
    Label12: TLabel;
    MENumVersion: TMaskEdit;
    TWCalcul: TToolWindow97;
    MECalcul: TMemo;
    bCalcul: TBitBtn;
    TEOffset: TEdit;
    LTEOffset: TLabel;
    Page3: TTabSheet;
    LTETri: TLabel;
    Label13: TLabel;
    MEFamCorresp: TMaskEdit;
    MEOrdreTri: TMaskEdit;
    MELongueur: TSpinEdit;
    MENbdecimal: TSpinEdit;
    MEDomaine: THValComboBox;
    procedure FormShow(Sender: TObject);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure OnFieldEnter(Sender: TObject);
    procedure OnFieldExit(Sender: TObject);
    procedure TVChange(Sender: TObject; Node: TTreeNode);
    procedure SG1Click(Sender: TObject);
    procedure BAnnulerClick(Sender: TObject);
    procedure BOuvrirClick(Sender: TObject);
    procedure BInsertClick(Sender: TObject);
    procedure BPostClick(Sender: TObject);
    procedure BDefaireClick(Sender: TObject);
    procedure BFirstClick(Sender: TObject);
    procedure BPrevClick(Sender: TObject);
    procedure BNextClick(Sender: TObject);
    procedure BLastClick(Sender: TObject);
    procedure BDeleteClick(Sender: TObject);
    procedure bDupliquerClick(Sender: TObject);
    procedure BImprimerClick(Sender: TObject);
    procedure BRechercherClick(Sender: TObject);
    procedure bCalculClick(Sender: TObject);
    procedure FindDialogFind(Sender: TObject);
(*    procedure UpDownLgClick(Sender: TObject; Button: TUDBtnType);
    procedure UpDownNbClick(Sender: TObject; Button: TUDBtnType);
*)
    procedure SG1Enter(Sender: TObject);
    procedure SG1Exit(Sender: TObject);
    procedure PostDrawCell(ACol,ARow : Longint; Canvas : TCanvas ; AState: TGridDrawState);
    procedure TWCalculVisibleChanged(Sender: TObject);
    procedure tvDragOver(Sender, Source: TObject; X, Y: Integer;
			State: TDragState; var Accept: Boolean);
    procedure tvDragDrop(Sender, Source: TObject; X, Y: Integer);
    procedure FormKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure TVClick(Sender: TObject);
    procedure METableNameClick(Sender: TObject);
    procedure MENomClick(Sender: TObject);

  private
    { Déclarations privées }
    TobPar          : TOB;
    TobTV           : TOB;
    TobDomaine      : TOB;
    TobCourante     : TOB;
    TobNewDomaine   : TOB;
    TobNewTable     : TOB;
    TobNewChamp     : TOB;
    TobNewTVDomaine : TOB;
    TobNewTVTable   : TOB;
    OneModified     : boolean;
    ModeInsert      : boolean;
    DontDo          : boolean;
    LesCols         : string;
    LeDomaine       : string;
    LTitre          : string;
    LePays          : string;
    LastObject      : TObject;
    {$IFNDEF CISXPGI}
    TPar            : TADOTABLE;
    {$ENDIF}
    procedure ChargeLesTobs;
    procedure InitTobDomaine;
    procedure InitTob(TobL : TOB);
{$IFNDEF CISXPGI}
    procedure InsereDansTobPar;
    procedure ChargementComboDomaine;
{$ELSE}
    procedure InsereDansTobPar (TobTemp : TOB);
{$ENDIF}
    procedure AnalyseTV;
    procedure EffaceChamps;
    procedure AfficheChamps(Tous : boolean = True);
    procedure EnregistreTob(TobM : TOB);
    function  SearchNodeFromTob(TobL : TOB) : TTreeNode;
    function  SearchGridFromTob(TobL : TOB) : integer;
    procedure PositionneBoutons(Sender : TObject; TN1 : TTreeNode);
    procedure PutGrid(TobC : TOB; SG : TStringGrid; LCols : string);
    function  ControleContenu(TMELoc : TMaskEdit) : boolean;
  public
    { Déclarations publiques }
  end;

var
  FParametre: TFParametre;
  
procedure ParamDictionnaire(PP : THPanel; Dom : string=''; Titre : string='Les paramètres'; Pay : string='');

implementation

uses UDMIMP, UIUtil,
{$IFDEF EAGLCLIENT}
UScriptTob,
{$ELSE}
UScript,
{$ENDIF}
UPDomaine(*, VoirTob*);

{$R *.DFM}

procedure ParamDictionnaire(PP : THPanel; Dom : string=''; Titre : string='Les paramètres'; Pay : string='');
var XX : TFParametre ;
BEGIN
XX:=TFParametre.Create(Application);
XX.LeDomaine := Dom;
XX.LTitre    := Titre;
XX.LePays    := Pay;
if PP=Nil then
   BEGIN
    Try
     XX.ShowModal ;
    Finally
     XX.Free ;
end ;
   END else
   BEGIN
   InitInside(XX,PP) ;
   XX.Show ;
   END ;
end;

//==============================================================================
procedure TFParametre.FormShow(Sender: TObject);
var ind1, NbCols : integer;
    stSep : string;
    Max   : integer;
begin
LesCols := '';
stSep  := '';
NbCols := 0;
Max    :=0;
{$IFNDEF CISXPGI}
if TPar = nil then TPar := TADOTable.Create(Application);
TPar.Connection := DMImport.DBGLOBAL;
TPar := DMImport.GzImpPar;
TPar.Open;

for ind1 := 3 to 9 do         // A voir pour resize
begin
    if Pos('Libre', TPar.Fields.Fields[ind1].FieldName) <> 0 then Continue;
    if Pos('Modifiable', TPar.Fields.Fields[ind1].FieldName) <> 0 then Continue;
    LesCols := LesCols + stSep + TPar.Fields.Fields[ind1].FieldName;
    if (SG1.ColCount - 1) < NbCols then
        SG1.ColCount := NbCols + 1;
    if TPar.Fields.Fields[ind1].FieldName = 'Contenu' then
        SG1.Cells[NbCols, 0] := 'Valeur'
    else if TPar.Fields.Fields[ind1].FieldName = 'Contenuoblig' then
        SG1.Cells[NbCols, 0] := 'Val. pos.'
        else
        SG1.Cells[NbCols, 0] := TPar.Fields.Fields[ind1].FieldName;
    if Max < SG1.Canvas.TextWidth(SG1.Cells[NbCols, 0]) then
    Max :=  SG1.Canvas.TextWidth(SG1.Cells[NbCols, 0]);
    SG1.ColWidths[NbCols] := Max+10;
    stSep := ';';
    Inc(NbCols);
end;
{$ELSE}
       LesCols := 'Nom;Contenu;Longueur;Obligatoire;NomPGI;Commentaire;Contenuoblig';
       for ind1 := 1 to 6 do
       begin
             if (SG1.ColCount - 1) < NbCols then
              SG1.ColCount := NbCols + 1;
             case NbCols of
                    0 : SG1.Cells[NbCols, 0] := 'Nom';
                    1 : SG1.Cells[NbCols, 0] := 'Valeur';
                    2 : SG1.Cells[NbCols, 0] := 'Longueur';
                    3 : SG1.Cells[NbCols, 0] := 'Obligatoire';
                    4 : SG1.Cells[NbCols, 0] := 'NomPGI';
                    5 : SG1.Cells[NbCols, 0] := 'Commentaire';
                    6 : SG1.Cells[NbCols, 0] := 'Val.Pos.';
                 end;
                if Max < SG1.Canvas.TextWidth(SG1.Cells[NbCols, 0]) then
                Max :=  SG1.Canvas.TextWidth(SG1.Cells[NbCols, 0]);
                SG1.ColWidths[NbCols] := Max+10;
                Inc(NbCols);
       end;
{$ENDIF}
    SG1.ColWidths[0] := 45;

SG1.ColCount := NbCols;
{$IFNDEF CISXPGI}
ChargementComboDomaine;
{$ELSE}
if LeDomaine = '' then MEDomaine.dataType  := 'CPZIMPDOMAINE'
else
begin
     MEDomaine.dataType  := 'TTPAYS';
     TMEDomaine.Caption  := 'Pays';
     MEDomaine.ItemIndex := MEDomaine.Values.IndexOf(LePays);
end;
{$ENDIF}
InitTobDomaine;
ChargeLesTobs;
TobTV.PutTreeView(TV, nil, 'NomTable');
AnalyseTV;
TV.TopItem.Expand(False);
OneModified := False;
{$IFNDEF CISXPGI}
MEModifiable.Checked := not (V_PGI.PassWord = DayPass(Date));
{$ELSE}
MEModifiable.Checked := TRUE;
{$ENDIF}
BDelete.Enabled := False;
BOuvrir.Enabled := False;
BDupliquer.Enabled := False;
SG1.PostDrawCell := PostDrawCell;

  Page2.Visible   := False;
  Page2.TabVisible := False;
{$IFNDEF CISXPGI}
  if not (V_PGI.PassWord = DayPass(Date)) then
  begin
    Binsert.Enabled  := FALSE;
    bPost.Enabled    := FALSE;
    BOuvrir.Enabled  := FALSE;
    Page3.Visible    := False;
    Page3.TabVisible := False;
  end
  else
{$ENDIF}
  begin
    Binsert.Enabled  := TRUE;
    bPost.Enabled    := TRUE;
    BOuvrir.Enabled  := TRUE;
    //Page3.Visible    := True;
    //Page3.TabVisible := True;
  end;

  HMTrad.ResizeGridColumns(SG1);

  PCPied.ActivePage     := Page1;
end;


procedure TFParametre.FormCloseQuery(Sender: TObject; var CanClose: Boolean);
var Rep : word;
begin
if OneModified then
begin
    Rep := HMsgBox.Execute(4, Caption, '');
    case Rep of
        mrYes    : BOuvrirClick(Sender);
        mrCancel : CanClose := False
    end;
end;
end;

procedure TFParametre.FormClose(Sender: TObject; var Action: TCloseAction);
begin
{$IFNDEF CISXPGI}
TPar.Close;
{$ENDIF}

if TobPar <> nil then
    TobPar.Free;
if TobTV <> nil then
    TobTV.Free;
if TobDomaine <> nil then
    TobDomaine.Free;
DontDo := True;
end;

procedure TFParametre.OnFieldEnter(Sender: TObject);
begin
        LastObject := Sender;
end;


procedure TFParametre.OnFieldExit(Sender: TObject);
var TMELoc   : TMaskEdit;
    TMECombo : TComboBox;
    TMECheck : TCheckBox;
    TMEMemo  : TMemo;
    TMEspin   : TSpinEdit;
    NomChamp, ValeurChamp : string;
    ARow                  : integer;
    Domaine               : string;
begin
if DontDo then Exit;
if (not (Sender is TMaskEdit)) and
   (not (Sender is TCheckBox)) and
   (not (Sender is TMemo)) and
   (not (Sender is TUpDown)) and
   (not (Sender is TComboBox)) and
   (not (Sender is TSpinEdit)) then Exit;
if (Sender is TMaskEdit) then
begin
    TMELoc := TMaskEdit(Sender);
    NomChamp := Copy(TMELoc.Name, 3, 255);
//    ValeurChamp := Trim(TMELoc.Text);  modif 3/6/3
    ValeurChamp := TMELoc.Text;
    if not ControleContenu(TMELoc) then
        Exit;
end;
if (Sender is TSpinEdit) then
begin
    TMEspin := TSpinEdit(Sender);
    NomChamp := Copy(TMEspin.Name, 3, 255);
    ValeurChamp := TMEspin.Text;
end;

if Sender is TComboBox then
begin
    TMECombo := TComboBox(Sender);
//    if Trim(TMECombo.Text) = '' then Exit;
    NomChamp := Copy(TMECombo.Name, 3, 255);
    if NomChamp = 'TypeAlpha' then
        ValeurChamp := Trim(Copy(TMECombo.Text, 1, 1))
    else if NomChamp = 'MajusMin' then
        ValeurChamp := UpperCase(Trim(Copy(TMECombo.Text, 1, 3)))
    else if NomChamp = 'Alignement' then
    begin
        ValeurChamp := Trim(Copy(TMECombo.Text, 1, 1));
        if ValeurChamp = '<' then
            ValeurChamp := '';
    end;
end;
if Sender is TCheckBox then
begin
    TMECheck := TCheckBox(Sender);
    NomChamp := Copy(TMECheck.Name, 3, 255);
    ValeurChamp := 'N';
    if TMECheck.Checked then ValeurChamp := 'O';
end;
if Sender is TMemo then
begin
    TMEMemo := TMemo(Sender);
    NomChamp := Copy(TMEMemo.Name, 3, 255);
    ValeurChamp := Trim(TMEMemo.Text);
end;

if (not ModeInsert)  then
begin
     // ajout me 16-01-2004
    if (TobCourante <> nil) and (Uppercase(TobCourante.GetValue(NomChamp)) <> UpperCase(ValeurChamp)) then
    begin
        if not TobCourante.FieldExists('Modifie') then
            TobCourante.AddChampSup('Modifie', False);
        TobCourante.PutValue(NomChamp, ValeurChamp);
        OneModified := True;
        AnalyseTV;
        bImprimer.Enabled := False;
        ARow := SG1.Row;
        SG1.RowCount := 2;
        TobCourante.Parent.PutGridDetail(SG1, False, False, LesCols, True);
        SG1.Row := ARow;
    end;
end
    else
begin
    IF (MENom.Text = '') and  (TV <> nil) and (TV.Selected.Level = 2) then
    begin
         PGIInfo ('Renseignez le nom du champ');
         exit;
    end;
    IF (METableName.Text = '') and  (TV <> nil) and (TV.Selected.Level = 1)  then
    begin
       PGIInfo ('Renseignez le nom de la table');
       exit;
    end;

    Domaine := MEDomaine.value;
    if Trim(MEDomaine.Text) <> '' then
    begin
        TobNewDomaine := TobPar.FindFirst(['NomTable'], [Domaine], False);
        TobNewTVDomaine := TobTV.FindFirst(['NomTable'], [Domaine], False);
    end;
    if TobNewDomaine = nil then
    begin
        TobNewDomaine := TOB.Create(Domaine, TobPar, -1);
        InitTob(TobNewDomaine);
        TobNewDomaine.PutValue('NomTable', Domaine);    // ajout me 16-01-2004
        TobNewDomaine.PutValue('Domaine', (*MEDomaine.Text[1]*)Domaine);
        TobNewDomaine.AddChampSup('Nouveau', False);
        OneModified := True;                          // ajout me 16-01-2004
        TobNewTVDomaine := TOB.Create((*MEDomaine.Text[1]*)Domaine, TobTV, -1);
        TobNewTVDomaine.Dupliquer(TobNewDomaine, False, True);
        TobTV.PutTreeView(TV, nil, 'NomTable');
        AnalyseTV;
        TV.TopItem.Expand(False);
        TV.Selected := SearchNodeFromTob(TobNewDomaine);
        SG1.RowCount := 2;
    end;
    METableName.Enabled := True;
    if NomChamp = 'Domaine' then Exit;
    if Trim(METableName.Text) <> '' then
    begin
        if TobNewDomaine = nil then Exit;
        TobNewTable := TobNewDomaine.FindFirst(['NomTable'], [METableName.Text], False);
        TobNewTVTable := TobNewTVDomaine.FindFirst(['NomTable'], [METableName.Text], False);
    end;
    if TobNewTable = nil then
    begin
        TobNewTable := TOB.Create(METableName.Text, TobNewDomaine, -1);
        InitTob(TobNewTable);
        TobNewTable.PutValue('NomTable', METableName.Text);
        TobNewTable.PutValue('Domaine', TobNewDomaine.GetValue('NomTable'));
        TobNewTable.PutValue('TableName', METableName.Text);
        TobNewTable.AddChampSup('Nouveau', False);
        OneModified := True;
        TobNewTVTable := TOB.Create(METableName.Text, TobNewTVDomaine, -1);
        TobNewTVTable.Dupliquer(TobNewTable, False, True);
        TobTV.PutTreeView(TV, nil, 'NomTable');
        AnalyseTV;
        TV.TopItem.Expand(False);
        TV.Selected := SearchNodeFromTob(TobNewTable);
        SG1.RowCount := 2;
    end;
    if NomChamp = 'TableName' then Exit;
    if NomChamp = 'Nom' then
    begin
        if Trim(MENom.Text) = '' then Exit;
        if TobNewDomaine = nil then Exit;
        if TobNewTable = nil then Exit;
        if TobNewChamp = nil then
            TobNewChamp := TobNewTable.FindFirst(['NomTable'], [MENom.Text], False)
            else
        begin
            TobNewChamp.PutValue('NomTable',        MENom.Text);
            TobNewChamp.PutValue('Nom', MENom.Text);
        end;
        if TobNewChamp = nil then
        begin
            TobNewChamp := TOB.Create(MENom.Text, TobNewTable, -1);
            InitTob(TobNewChamp);
            TobNewChamp.PutValue('NomTable'            , MENom.Text);
            TobNewChamp.PutValue('Domaine' , TobNewDomaine.GetValue('NomTable'));
            TobNewChamp.PutValue('TableName'   , TobNewTable.GetValue('NomTable'));
            TobNewChamp.PutValue('Nom'     , MENom.Text);
            TobNewChamp.AddChampSup('Nouveau', False);
            TobNewChamp.AddChampSup('Champ', False);
            TobNewChamp.AddChampSupValeur('ImageIndex', 4, False);
            OneModified := True; 
        end;
        BPost.Enabled := True;
    end
        else
    begin
        if UpperCase(TobNewChamp.GetValue(NomChamp)) <> UpperCase(ValeurChamp) then
            OneModified := True;
        TobNewChamp.PutValue(NomChamp, ValeurChamp);
    end;
    TobNewChamp.PutLigneGrid(SG1, SG1.Row, False, False, LesCols);
end;
end;


procedure TFParametre.tvDragOver(Sender, Source: TObject; X,
	Y: Integer; State: TDragState; var Accept: Boolean);
var
	TN, TN1 : TTreeNode;
begin
	Accept := False;
	if not ( Source is TTreeView) then Exit;
	TN := tv.Selected;
	TN1 := tv.GetNodeAt(X, Y);
	if TN1 = nil then Accept := TN.Parent.GetNextSibling = nil
	else if TN1.Level = TN.Level then
	begin if TN.Parent = TN1.Parent  then Accept := True; end
	else if TN1 = TN.Parent.GetNextSibling then Accept := True;
end;
procedure TFParametre.tvDragDrop(Sender, Source: TObject; X, Y: Integer);
var
	TN, TN1, TN3          : TTreeNode;
        N,I                   : integer;
        TT,TOBL               : TOB;
        Ligne                 : integer;
begin
	TN1 := tv.GetNodeAt(X, Y);
	TN := tv.Selected;
        TT := TOB(TN.Data);
        tv.Items.BeginUpdate;
	if TN1 = nil then TN.MoveTo(TN.Parent, naAddChild)
	else if TN <> TN1 then
		if TN.Level = TN1.Level then TN.MoveTo(TN1, naInsert)
		else TN.MoveTo(TN.Parent, naAddChild);
	tv.Items.EndUpdate;

       Ligne := 1;
       TN3 := tv.selected.Parent.GetFirstChild;
       for I := 0 to tv.selected.Parent.Count-1 do
       begin
          TobL := TobPar.FindFirst(['Domaine', 'TableName'],
                                  [TT.GetValue('Domaine'),
                                   TN3.Text], True);
          TOBL.PutValue('Numero', Ligne);
          for N :=0 to TobL.detail.Count - 1 do
          begin
                 TOBL.detail[N].PutValue('Numero', Ligne);
                 if not TOBL.FieldExists('Modifie') then
                 TOBL.detail[N].AddChampSup('Modifie', False);
                 inc(Ligne);
          end;

         TN3 := tv.selected.Parent.GetNextChild(TN3);
       end;

//debug       TobPar.SaveToFile('XXXX.tmp',True,True,True);

end;

procedure TFParametre.TVChange(Sender: TObject; Node: TTreeNode);
begin
if Node.Data = nil then Exit;
if TV.Selected.Level = 0 then exit;
TobCourante := TobPar.FindFirst(['Domaine', 'TableName', 'Nom'],
                                [TOB(Node.Data).GetValue('Domaine'),
                                 TOB(Node.Data).GetValue('TableName'),
                                 TOB(Node.Data).GetValue('Nom')], True);
if TobCourante = nil then Exit;
if TobCourante.NomTable = 'Les paramètres' then Exit;
if not TobCourante.FieldExists('Champ') then
begin
    EffaceChamps;
    AfficheChamps(False);
    MEDomaine.Enabled := False;
    METableName.Enabled := False;
    SG1.VidePile(False);
    SG1.Enabled := True;
    SG1.RowCount := 2;
    if Node.Level = 2 then
    begin
//        PutGrid(TobCourante, SG1, LesCols);
        TobCourante.PutGridDetail(SG1, False, False, LesCols, True);
        SG1.Row := 1;
        SG1Click(Sender)
    end;
// ajout me
        tv.OnDragOver := tvDragOver;
        tv.OnDragDrop := tvDragDrop;

end;
end;

procedure TFParametre.SG1Click(Sender: TObject);
begin
if SG1.Objects[0, SG1.Row] = nil then Exit;
DontDo := False;
TobCourante := TOB(SG1.Objects[0, SG1.Row]);
// ajout me 17/11/2005
IF (MENom.Text = '') and (SG1.Row > 1) then
   TobCourante := TOB(SG1.Objects[0, SG1.Row-1]);

AfficheChamps;
end;

procedure TFParametre.BAnnulerClick(Sender: TObject);
begin
//Close;
end;

procedure TFParametre.BOuvrirClick(Sender: TObject);
begin
EnregistreTob(TobPar);
OneModified := False;
bImprimer.Enabled := True;
bOuvrir.Enabled := False;
TobPar.Free;
TobTV.Free;
SG1.VidePile(False);
SG1.Enabled := True;
SG1.RowCount := 2;
ChargeLesTobs;
TobTV.PutTreeView(TV, nil, 'NomTable');
AnalyseTV;
TV.TopItem.Expand(False);
DontDo := True;
end;

procedure TFParametre.BinsertClick(Sender: TObject);
begin
if ModeInsert and (MEDomaine.Text = '') then
     exit;
if ModeInsert then
    BPostClick(Sender);

    // ajout me
if (TEOffset.Text <> '') and (MeLongueur.value <> 0) then
     TEOffset.Text := IntToStr(StrToInt(TEOffset.Text) + MeLongueur.value);

MEDomaine.Text := '';
METableName.Text := '';
MENom.Text := '';
EffaceChamps;
ModeInsert := True;
TobNewDomaine := nil;
TobNewTable := nil;
TobNewChamp := nil;
{$IFNDEF CISXPGI}
MEModifiable.Checked := not (V_PGI.PassWord = DayPass(Date));
{$ELSE}
MEModifiable.Checked := TRUE;
{$ENDIF}
MEDomaine.Enabled := True;
METableName.Enabled := True;
MENom.Enabled := True;
BPost.Enabled := False;
DontDo := True;
if (TV.Selected <> nil) then
begin
      if  (TV.Selected.Level = 0) then
          MEDomaine.SetFocus
      else if TV.Selected.Level = 1  then
      begin
          if TOB(TV.Selected.Data).FieldExists('Domaine') then
          MEDomaine.ItemIndex := MEDomaine.Values.IndexOf(TOB(TV.Selected.Data).GetValue('Domaine'));
          MEDomaine.Enabled := False;
          METableName.SetFocus;
      end
      else
      begin
          if TOB(TV.Selected.Data).FieldExists('Domaine') then
          MEDomaine.ItemIndex :=  MEDomaine.Values.IndexOf(TOB(TV.Selected.Data).GetValue('Domaine'));
          MEDomaine.Enabled := False;
          METableName.Text := TOB(TV.Selected.Data).GetValue('TableName');
          METableName.Enabled := False;
          MENom.Enabled := True;
          MENom.SetFocus;
      end;
end;
MEContenu.Enabled     := True;
MELongueur.Enabled    := True;
MEObligatoire.Enabled := True;
MEMajusMin.Enabled    := True;
METypeAlpha.Enabled   := True;
MEContenuoblig.Enabled:= True;
MENomPGI.Enabled      := True;
MECommentaire.Enabled := True;
MEModifiable.Enabled  := True;
MEFige.Enabled        := True;
MECalcul.Enabled      := True;
// ajout me
MEOrdreTri.Enabled         := TRUE;
MEFamCorresp.Enabled      := TRUE;
SG1.RowCount := SG1.RowCount + 1;
SG1.Row := SG1.RowCount - 1;
DontDo := False;
end;

procedure TFParametre.BPostClick(Sender: TObject);
begin
OnFieldExit(LastObject);
if (ModeInsert) and (TobNewChamp <> nil) then
    TobCourante := TobNewChamp;
ModeInsert := False;
BOuvrir.Enabled := True;
MEDomaine.Enabled := False;
METableName.Enabled := False;
MENom.Enabled := False;
PutGrid(TobCourante.Parent, SG1, LesCols);
AnalyseTV;
end;

procedure TFParametre.BDefaireClick(Sender: TObject);
var TobT1 : TOB;
begin
if ModeInsert then
begin
    TobNewChamp.Free;
    TobNewChamp := nil;
    if TobNewTable.Detail.Count = 0 then
    begin
        TobNewTable.Free;
        TobNewTable := nil;
    end;
    if TobNewDomaine.Detail.Count = 0 then
    begin
        TobNewDomaine.Free;
        TobNewDomaine := nil;
    end;
    MEDomaine.Text := '';
    METableName.Text := '';
    MENom.Text := '';
    EffaceChamps;
    MEDomaine.Enabled := True;
    METableName.Enabled := True;
    MEDomaine.SetFocus;
end
    else
begin
    TobT1 := TOB(SG1.Objects[0, SG1.Row]);
    if TobT1.FieldExists('A Supprimer') then
    begin
        TobT1.DelChampSup('A Supprimer', False);
        TobT1.PutLigneGrid(SG1, SG1.Row, False, False, LesCols);
    end;
end;
end;

procedure TFParametre.BFirstClick(Sender: TObject);
var TN1 : TTreeNode;
begin
TN1 := TV.Items.GetFirstNode;
TN1.Selected := True;
end;

procedure TFParametre.BPrevClick(Sender: TObject);
var TN1 : TTreeNode;
begin
TN1 := TV.Selected;
if TN1.GetPrev = nil then Exit;
PositionneBoutons(Sender, TN1);
TN1 := TN1.GetPrev;
TN1.Selected := True;
end;

procedure TFParametre.BNextClick(Sender: TObject);
var TN1 : TTreeNode;
begin
TN1 := TV.Selected;
if TN1.GetNext = nil then Exit;
PositionneBoutons(Sender, TN1);
TN1 := TN1.GetNext;
TN1.Selected := True;
end;

procedure TFParametre.BLastClick(Sender: TObject);
var TN1 : TTreeNode;
begin
TN1 := TV.Items[TV.Items.Count - 1];
TN1.Selected := True;
end;

procedure TFParametre.BDeleteClick(Sender: TObject);
var
TobT1,TobTable        : TOB;
i,j                   : integer;
begin
    if (TV.Selected.Level = 1) then // niveau domaine
    begin
               if TobPar = nil then exit;
               TobTable := TobPar.FindFirst(['Domaine'],
                                  [TOB(TV.Selected.Data).GetValue('Domaine')], True);
               TV.Selected := SearchNodeFromTob(TobTable);
               for j := 0 to TobTable.detail.count-1 do
               begin
                    TobCourante := TobTable.detail[j];
                    for i := 0 to TobCourante.detail.count-1 do
                       begin
                           TobT1 := TobCourante.detail[i];
                           TobT1.AddChampSup('A Supprimer', False);
                           OneModified := True;
                           SG1.RePaint;
                      end;
               end;
               TV.Selected.Delete;
    end
    else
    begin
          if ((TV.Selected.Level = 2) and (not SG1.Focused)) then // niveau des enregistrements
          begin

                 TobCourante := TobPar.FindFirst(['Domaine', 'TableName', 'Nom'],
                                      [TOB(TV.Selected.Data).GetValue('Domaine'),
                                       TOB(TV.Selected.Data).GetValue('TableName'),
                                       TOB(TV.Selected.Data).GetValue('Nom')], True);
                 if TobCourante = nil then Exit;
                 for i := 0 to TobCourante.detail.count-1 do
                 begin
                     TobT1 := TobCourante.detail[i];
                     TobT1.AddChampSup('A Supprimer', False);
                     OneModified := True;
      //               SG1.RePaint;
                end;
                TV.Selected.Delete;
          end
          else
          begin
                if not SG1.Focused then Exit;
                TobT1 := TOB(SG1.Objects[0, SG1.Row]);
                if TobT1.Detail.Count > 0 then Exit;
                TobT1.AddChampSup('A Supprimer', False);
                OneModified := True;
                SG1.RePaint;
          end;
    end;
end;

procedure TFParametre.bDupliquerClick(Sender: TObject);
var TobT1 : TOB;
begin
if not SG1.Focused then Exit;
TobT1 := TOB(SG1.Objects[0, SG1.Row]);
if not TobT1.FieldExists('Champ') then Exit;
TobNewTable := TobT1.Parent;
TobNewDomaine := TobNewTable.Parent;
TobCourante := TOB.Create('', TobNewTable, -1);
InitTob(TobCourante);
TobCourante.Dupliquer(TobT1, False, True, True);
TobCourante.AddChampSupValeur('Nouveau', '', False);
TobCourante.AddChampSup('Champ', False);
TobCourante.AddChampSupValeur('ImageIndex', 4, False);
MEDomaine.Enabled := False;
METableName.Enabled := False;
MENom.SetFocus;
OneModified := True;
end;

procedure TFParametre.BImprimerClick(Sender: TObject);
begin
DBG.DataSource := DS1;
{$IFDEF EAGLCLIENT}
PrintDBGrid(Caption, SG1.ListeParam, '', '');
{$ELSE}
PrintDBGrid(DBG, nil, Caption, '');
{$ENDIF}
DBG.DataSource := nil;
end;

procedure TFParametre.BRechercherClick(Sender: TObject);
begin
DBG.DataSource := nil;
DBG.DataSource := DS1;
FindDialog.Execute;
end;

procedure TFParametre.bCalculClick(Sender: TObject);
begin
{TWCalcul.Height := Panel_Champs.Height;
TWCalcul.Width  := Panel_Champs.Width;
TWCalcul.Left   := Panel_Champs.Left;
TWCalcul.Top    := Panel_Champs.Top; }
TWCalcul.Visible := True;
end;

procedure TFParametre.FindDialogFind(Sender: TObject);
var TobL : TOB;
    ind1 : integer;
    FirstFind : boolean;
    TN1 : TTreeNode;
begin
FirstFind := True;
Rechercher(DBG, FindDialog, FirstFind);
TobL := TobPar.FindFirst(['Numero', 'Domaine'],
                         [TobCourante.GetValue('Numero'), TobCourante.GetValue('Domaine')], True);
ind1 := SearchGridFromTob(TobL);
if ind1 < 0 then
begin
    TN1 := SearchNodeFromTob(TobL.Parent);
    if TN1 <> nil then
    begin
        TV.Selected := TN1;
        ind1 := SearchGridFromTob(TobL);
    end;
end;
if ind1 > 0 then
    SG1.Row := ind1;
//    SearchNodeFromGrid(ind1);
end;

procedure TFParametre.SG1Enter(Sender: TObject);
begin
BDelete.Enabled := True;
BDupliquer.Enabled := True;
end;

procedure TFParametre.SG1Exit(Sender: TObject);
begin
BDelete.Enabled := False;
BDupliquer.Enabled := False;
end;

procedure TFParametre.PostDrawCell(ACol,ARow : Longint; Canvas : TCanvas ; AState: TGridDrawState);
var TobT1 : TOB;
    TRTmp : TRect;
begin
if ARow < 1 then Exit;
if TOB(SG1.Objects[0, ARow]) = nil then Exit;
TobT1 := TOB(SG1.Objects[0, ARow]);
if not TobT1.FieldExists('A Supprimer') then Exit;
With SG1.Canvas do
begin
    TRTmp := SG1.CellRect(ACol, ARow);
    TRTmp.Top  := TRTmp.Top + Trunc((TRTmp.Bottom - TRTmp.Top + 1) / 2);
    TRTmp.Bottom := TRTmp.Top + 1;
    MoveTo(TRTmp.Left, TRTmp.Top);
    LineTo(TRTmp.Right, TRTmp.Top);
end;
end;

//==============================================================================
procedure TFParametre.ChargeLesTobs;
{$IFDEF CISXPGI}
var
    TobTemp                : TOB;
    CodeRetour             : integer;
    i                      : integer;
{$ENDIF}
begin
{$IFNDEF CISXPGI}
TobPar := TOB.Create('Les paramètres', nil, -1);
TobPar.AddChampSupValeur('NomTable', 'Les paramètres', False);
InitTob(TobPar);
TobTV := TOB.Create('Les paramètres', nil, -1);
TobTV.AddChampSupValeur('NomTable', 'Les paramètres', False);
InitTob(TobTV);
TPar := DMImport.GzImpPar;
if not TPar.Active then TPar.Open;
with TPar do
begin
    First;
    while not Eof do
    begin
        InsereDansTobPar;
        Next;
    end;
end;
{$ELSE}
        TobTemp := nil;
        CodeRetour := RendTobparametre (LeDomaine, TobTemp, 'COMPTA', 'PARAM', LePays);
        TobPar := TOB.Create(LTitre, nil, -1);
        TobPar.AddChampSupValeur('NomTable', LTitre, False);
        InitTob(TobPar);
        TobTV := TOB.Create(LTitre, nil, -1);
        TobTV.AddChampSupValeur('NomTable', LTitre, False);
        InitTob(TobTV);
        if CodeRetour = -1 then  // existe dans la base
        begin
              //relevé
              For i:= 0 to TobTemp.detail.count-1 do
              begin
                       InsereDansTobPar(TobTemp.detail[i]);
              end;
        end;
{$ENDIF}
end;

procedure TFParametre.InitTobDomaine;
{$IFNDEF CISXPGI}
  var TobL : TOB;
  Q        : TQuery;
{$ELSE}
{$IFNDEF EAGLCLIENT}
  var TobL : TOB;
  Q        : TQuery;
{$ELSE}
{$ENDIF}
{$ENDIF}
begin
{$IFNDEF CISXPGI}
Q := OpenSQLADO ('Select * from '+DMImport.GzImpDomaine.TableName, DMImport.DBGlobal.ConnectionString );
TobDomaine := TOB.Create('', nil, -1);
While not Q.EOF do
begin
        TobL := TOB.Create('', TobDomaine, -1);
        TobL.AddChampSupValeur('Domaine', Q.FindField('Domaine').asstring, False);
        TobL.AddChampSupValeur('Libelle', Q.FindField('Libelle').asstring, False);
        Q.next;
end;
Ferme (Q);
{$ELSE}
{$IFNDEF EAGLCLIENT}
Q := OpenSQL ('select * from PAYS', TRUE);
TobDomaine := TOB.Create('', nil, -1);
While not Q.EOF do
begin
        TobL := TOB.Create('', TobDomaine, -1);
        TobL.AddChampSupValeur('Domaine', Q.FindField('PY_PAYS').asstring, False);
        TobL.AddChampSupValeur('Libelle', Q.FindField('PY_LIBELLE').asstring, False);
        Q.next;
end;
Ferme (Q);
{$ELSE}
ChargementPays (TobDomaine);
{$ENDIF}

{$ENDIF}
end;


procedure TFParametre.InitTob(TobL : TOB);
begin
if TObL.FieldExists ('NomTable') then exit;
TobL.AddChampSupValeur('NomTable'                ,  '', False);
TobL.AddChampSupValeur('Numero'      ,  -1, False);
TobL.AddChampSupValeur('Domaine'     ,  '', False);
TobL.AddChampSupValeur('TableName'   ,  '', False);
TobL.AddChampSupValeur('Nom'         ,  '', False);
TobL.AddChampSupValeur('Contenu'     ,  '', False);
TobL.AddChampSupValeur('Longueur'    ,   0, False);
TobL.AddChampSupValeur('Obligatoire' ,  'N', False);
TobL.AddChampSupValeur('MajusMin'    ,  '', False);
TobL.AddChampSupValeur('TypeAlpha'   ,  '', False);
TobL.AddChampSupValeur('Contenuoblig',  '', False);
TobL.AddChampSupValeur('NomPGI'      ,  '', False);
TobL.AddChampSupValeur('Commentaire' ,  '', False);
TobL.AddChampSupValeur('Libre1'      ,  '', False);
TobL.AddChampSupValeur('Libre2'      ,  '', False);
TobL.AddChampSupValeur('Libre3'      ,  '', False);
TobL.AddChampSupValeur('Libre4'      ,  '', False);
TobL.AddChampSupValeur('Libre5'      ,  '', False);
TobL.AddChampSupValeur('Libre6'      ,  '', False);
TobL.AddChampSupValeur('Libre7'      ,  '', False);
TobL.AddChampSupValeur('Libre8'      ,  '', False);
TobL.AddChampSupValeur('Libre9'      ,  '', False);
TobL.AddChampSupValeur('Libre10'     ,  '', False);
{$IFNDEF CISXPGI}
if not (V_PGI.PassWord = DayPass(Date)) then
    TobL.AddChampSupValeur('Modifiable'  , 'O', False)
    else
    TobL.AddChampSupValeur('Modifiable'  , 'N', False);
{$ELSE}
TobL.AddChampSupValeur('Modifiable'  , 'O', False);
{$ENDIF}
TobL.AddChampSupValeur('Fige'        , 'N', False);
TobL.AddChampSupValeur('Alignement'  ,  '', False);
TobL.AddChampSupValeur('Separateur'  ,  '', False);
TobL.AddChampSupValeur('Nbdecimal'  ,   2, False);
TobL.AddChampSupValeur('NumVersion'  ,   1, False);
TobL.AddChampSupValeur('Calcul'      ,  '', False);
TobL.AddChampSupValeur('Famcorresp'      ,  '', False);
TobL.AddChampSupValeur('Ordretri'      ,  '', False);
end;

{$IFNDEF CISXPGI}
procedure TFParametre.InsereDansTobPar;
var TobDom, TobTable, TobChamp : TOB;
    TobTVDomaine, TobTVTable : TOB;
    cTemp : string;
begin
TobDom := TobPar.FindFirst(['NomTable'], [TPar.FindField('Domaine').AsString], False);
TobTVDomaine := TobTV.FindFirst(['NomTable'], [TPar.FindField('Domaine').AsString], False);
if TobDom = nil then
begin
    TobDom := TOB.Create(TPar.FindField('Domaine').AsString, TobPar, -1);
    InitTob(TobDom);
    TobDom.PutValue('NomTable', TPar.FindField('Domaine').AsString);
    TobDom.PutValue('Numero'      , TPar.FindField('Numero').AsInteger);
    TobDom.PutValue('Domaine'     , TPar.FindField('Domaine').AsString);
    TobDom.PutValue('TableName'       , '');
    TobDom.PutValue('Nom'         , '');
    TobDom.AddChampSupValeur('ImageIndex', 7, False);
    TobTVDomaine := TOB.Create(TPar.FindField('Domaine').AsString, TobTV, -1);
    TobTVDomaine.Dupliquer(TobDom, False, True);
end;
TobTable := TobDom.FindFirst(['NomTable'], [TPar.FindField('TableName').AsString], False);
//TobTVTable := TobTVDomaine.FindFirst(['NomTable'], [TParTable.AsString], False);
if TobTable = nil then
begin
    TobTable := TOB.Create(TPar.FindField('TableName').AsString, TobDom, -1);
    InitTob(TobTable);
    TobTable.PutValue('NomTable', TPar.FindField('TableName').AsString);
    TobTable.PutValue('Numero'      , TPar.FindField('Numero').AsInteger);
    TobTable.PutValue('Domaine'     , TPar.FindField('Domaine').AsString);
    TobTable.PutValue('TableName'       , TPar.FindField('TableName').AsString);
    TobTable.PutValue('Nom'         , '');
    TobTable.AddChampSupValeur('ImageIndex', 1, False);
    TobTVTable := TOB.Create(TPar.FindField('TableName').AsString, TobTVDomaine, -1);
    TobTVTable.Dupliquer(TobTable, False, True);
end;
TobChamp := TOB.Create(TPar.FindField('Nom').AsString, TobTable, -1);
InitTob(TobChamp);
TobChamp.PutValue('Numero'     , TPar.FindField('Numero').AsInteger);
TobChamp.PutValue('Domaine'     , TPar.FindField('Domaine').AsString);
TobChamp.PutValue('TableName'   , TPar.FindField('TableName').AsString);
TobChamp.PutValue('NomTable'                , TPar.FindField('Nom').AsString);
TobChamp.PutValue('Nom'         , TPar.FindField('Nom').AsString);
TobChamp.PutValue('Contenu'     , TPar.FindField('Contenu').AsString);
TobChamp.PutValue('Longueur'    , TPar.FindField('Longueur').AsString);
TobChamp.PutValue('Obligatoire' , TPar.FindField('Obligatoire').AsString);
TobChamp.PutValue('MajusMin'    , TPar.FindField('MajusMin').AsString);
TobChamp.PutValue('TypeAlpha'   , TPar.FindField('TypeAlpha').AsString);
TobChamp.PutValue('Contenuoblig', TPar.FindField('Contenuoblig').AsString);
TobChamp.PutValue('NomPGI'      , TPar.FindField('NomPGI').AsString);
TobChamp.PutValue('Commentaire'     , TPar.FindField('Commentaire').AsString);
TobChamp.PutValue('Libre1'      , TPar.FindField('Libre1').AsString);
TobChamp.PutValue('Libre2'      , TPar.FindField('Libre2').AsString);
TobChamp.PutValue('Libre3'      , TPar.FindField('Libre3').AsString);
TobChamp.PutValue('Libre4'      , TPar.FindField('Libre4').AsString);
TobChamp.PutValue('Libre5'      , TPar.FindField('Libre5').AsString);
TobChamp.PutValue('Libre6'      , TPar.FindField('Libre6').AsString);
TobChamp.PutValue('Libre7'      , TPar.FindField('Libre7').AsString);
TobChamp.PutValue('Libre8'      , TPar.FindField('Libre8').AsString);
TobChamp.PutValue('Libre9'      , TPar.FindField('Libre9').AsString);
TobChamp.PutValue('Libre10'     , TPar.FindField('Libre10').AsString);
cTemp := Copy(TPar.FindField('Modifiable').AsString, 1, 1);
if Trim(cTemp) = '' then
    cTemp := '1';
TobChamp.PutValue('Modifiable' , Chr(StrToInt(cTemp[1]) + Ord('N')));
cTemp := Copy(TPar.FindField('Fige').AsString, 1, 1);
if Trim(cTemp) = '' then
    cTemp := '0';
TobChamp.PutValue('Fige'        , Chr(StrToInt(cTemp[1]) + Ord('N')));
TobChamp.PutValue('Alignement'  , TPar.FindField('Alignement').AsString);
TobChamp.PutValue('Separateur'  , TPar.FindField('Separateur').AsString);
TobChamp.PutValue('Nbdecimal'   , TPar.FindField('Nbdecimal').AsInteger);
TobChamp.PutValue('NumVersion'  , TPar.FindField('NumVersion').AsInteger);
TobChamp.PutValue('Calcul'      , TPar.FindField('Calcul').AsString);
// ajout me
TobChamp.PutValue('Famcorresp'   , TPar.FindField('Famcorresp').Asstring);
TobChamp.PutValue('Ordretri'    ,  TPar.FindField('Ordretri').Asstring);

TobChamp.AddChampSupValeur('ImageIndex', 4, False);
TobChamp.AddChampSup('Champ', False);
end;

{$ELSE}

procedure TFParametre.InsereDansTobPar (TobTemp : TOB);
var TobDom, TobTable, TobChamp : TOB;
    TobTVDomaine, TobTVTable : TOB;
    cTemp : string;

begin
TobDom := TobPar.FindFirst(['NomTable'], [TobTemp.Getvalue('Domaine')], False);
TobTVDomaine := TobTV.FindFirst(['NomTable'], [TobTemp.Getvalue('Domaine')], False);
if TobDom = nil then
begin
    TobDom := TOB.Create(TobTemp.Getvalue('Domaine'), TobPar, -1);
    InitTob(TobDom);
    TobDom.PutValue('NomTable', TobTemp.Getvalue('Domaine'));
    TobDom.PutValue('Numero'      , TobTemp.Getvalue('Numero'));
    TobDom.PutValue('Domaine'     , TobTemp.Getvalue('Domaine'));
    TobDom.PutValue('TableName'       , '');
    TobDom.PutValue('Nom'         , '');
    TobDom.AddChampSupValeur('ImageIndex', 7, False);
    TobTVDomaine := TOB.Create(TobTemp.Getvalue('Domaine'), TobTV, -1);
    TobTVDomaine.Dupliquer(TobDom, False, True);
end;
TobTable := TobDom.FindFirst(['NomTable'], [TobTemp.Getvalue('TableName')], False);
if TobTable = nil then
begin
    TobTable := TOB.Create(TobTemp.Getvalue('TableName'), TobDom, -1);
    InitTob(TobTable);
    TobTable.PutValue('NomTable', TobTemp.Getvalue('TableName'));
    TobTable.PutValue('Numero'      ,  TobTemp.Getvalue('Numero'));
    TobTable.PutValue('Domaine'     , TobTemp.Getvalue('Domaine'));
    TobTable.PutValue('TableName'       , TobTemp.Getvalue('TableName'));
    TobTable.PutValue('Nom'         , '');
    TobTable.AddChampSupValeur('ImageIndex', 1, False);
    TobTVTable := TOB.Create(TobTemp.Getvalue('TableName'), TobTVDomaine, -1);
    TobTVTable.Dupliquer(TobTable, False, True);
end;
TobChamp := TOB.Create(TobTemp.Getvalue('Nom'), TobTable, -1);
InitTob(TobChamp);
TobChamp.PutValue('Numero'      , TobTemp.Getvalue('Numero'));
TobChamp.PutValue('Domaine'     , TobTemp.Getvalue('Domaine'));
TobChamp.PutValue('TableName'   , TobTemp.Getvalue('TableName'));
TobChamp.PutValue('NomTable'    , TobTemp.Getvalue('Nom'));
TobChamp.PutValue('Nom'         , TobTemp.Getvalue('Nom'));
TobChamp.PutValue('Contenu'     , TobTemp.Getvalue('Contenu'));
TobChamp.PutValue('Longueur'    , TobTemp.Getvalue('Longueur'));
TobChamp.PutValue('Obligatoire' , TobTemp.Getvalue('Obligatoire'));
TobChamp.PutValue('MajusMin'    , TobTemp.Getvalue('MajusMin'));
TobChamp.PutValue('TypeAlpha'   , TobTemp.Getvalue('TypeAlpha'));
TobChamp.PutValue('Contenuoblig', TobTemp.Getvalue('Contenuoblig'));
TobChamp.PutValue('NomPGI'      , TobTemp.Getvalue('NomPGI'));
TobChamp.PutValue('Commentaire' , TobTemp.Getvalue('Commentaire'));
TobChamp.PutValue('Libre1'      , TobTemp.Getvalue('Libre1'));
TobChamp.PutValue('Libre2'      , TobTemp.Getvalue('Libre2'));
TobChamp.PutValue('Libre3'      , TobTemp.Getvalue('Libre3'));
TobChamp.PutValue('Libre4'      , TobTemp.Getvalue('Libre4'));
TobChamp.PutValue('Libre5'      , TobTemp.Getvalue('Libre5'));
TobChamp.PutValue('Libre6'      , TobTemp.Getvalue('Libre6'));
TobChamp.PutValue('Libre7'      , TobTemp.Getvalue('Libre7'));
TobChamp.PutValue('Libre8'      , TobTemp.Getvalue('Libre8'));
TobChamp.PutValue('Libre9'      , TobTemp.Getvalue('Libre9'));
TobChamp.PutValue('Libre10'     , TobTemp.Getvalue('Libre10'));
cTemp := Copy(TobTemp.Getvalue('Modifiable'), 1, 1);
if Trim(cTemp) = '' then
    cTemp := '1';
if cTemp = 'O' then cTemp := '0'
else cTemp := '1';
TobChamp.PutValue('Modifiable'  , Chr(StrToInt(cTemp[1]) + Ord('N')));
cTemp := Copy(TobTemp.Getvalue('Fige'), 1, 1);
if Trim(cTemp) = '' then
    cTemp := '0';
if cTemp = 'O' then cTemp := '0'
else cTemp := '1';
TobChamp.PutValue('Fige'        , Chr(StrToInt(cTemp[1]) + Ord('N')));
TobChamp.PutValue('Alignement'  , TobTemp.Getvalue('Alignement'));
TobChamp.PutValue('Separateur'  , TobTemp.Getvalue('Separateur'));
TobChamp.PutValue('Nbdecimal'   , TobTemp.Getvalue('Nbdecimal'));
TobChamp.PutValue('NumVersion'  , TobTemp.Getvalue('NumVersion'));
TobChamp.PutValue('Calcul'      , TobTemp.Getvalue('Calcul'));
TobChamp.PutValue('Famcorresp'   , TobTemp.Getvalue('Famcorresp'));
TobChamp.PutValue('Ordretri'    ,  TobTemp.Getvalue('Ordretri'));

TobChamp.AddChampSupValeur('ImageIndex', 4, False);
TobChamp.AddChampSup('Champ', False);
end;
{$ENDIF}

procedure TFParametre.AnalyseTV;
var TN1 : TTreeNode;
    VInt : Variant;
    TobL : TOB;
begin
TN1 := TV.TopItem;
TN1.ImageIndex := 11;
TN1.SelectedIndex := 11;
TN1 := TN1.GetNext;
while TN1 <> nil do
begin
    if TN1.Level = 1 then
    begin
        TobL := TobDomaine.FindFirst(['Domaine'], [TN1.Text], True);
        if TobL <> nil then
            TN1.Text := TOB(TN1.Data).GetValue('NomTable') + ' (' +
                        TobL.GetValue('Libelle') + ')';
    end;
    if TOB(TN1.Data).FieldExists('ImageIndex') then
        VInt := TOB(TN1.Data).GetValue('ImageIndex')
        else
        VInt := 4;
    TN1.ImageIndex := VarAsType(VInt,VarInteger);
    if TOB(TN1.Data).FieldExists('Modifie') then TN1.ImageIndex := TN1.ImageIndex - 1;
    TN1.SelectedIndex := TN1.ImageIndex;
    if TOB(TN1.Data).FieldExists('A Supprimer') then
    begin
        TN1.ImageIndex := 10;
        TN1.SelectedIndex := 10;
    end;
    if TOB(TN1.Data).FieldExists('Nouveau') then
    begin
        TN1.ImageIndex := 9;
        TN1.SelectedIndex := 9;
    end;
    TN1 := TN1.GetNext;
end;
end;

procedure TFParametre.EffaceChamps;
begin
MEContenu.Text         := '';
MELongueur.value       := 1;
MEObligatoire.Checked  := False;
MEMajusMin.ItemIndex   := 0;
METypeAlpha.ItemIndex  := 0;
MEContenuoblig.Text    := '';
MENomPGI.Text          := '';
MECommentaire.Text     := '';
MELibre1.Text          := '';
MELibre2.Text          := '';
MELibre3.Text          := '';
MELibre4.Text          := '';
MELibre5.Text          := '';
MELibre6.Text          := '';
MELibre7.Text          := '';
MELibre8.Text          := '';
MELibre9.Text          := '';
MELibre10.Text         := '';
MEFige.Checked         := False;
MEAlignement.ItemIndex := 0;
MESeparateur.Text      := '';
MENbdecimal.value      := 0;
MENumVersion.Text      := '1';
MECalcul.Text          := '';
// ajout me
MEOrdreTri.Text             := '';
MEFamCorresp.Text       := '';
end;

procedure TFParametre.AfficheChamps(Tous : boolean = True);
var Pass     : string;
    Autorise : boolean;
    ind1, Offset : integer;
begin
Pass := DayPass(Date);
{$IFDEF CISXPGI}
Autorise := TRUE;
{$ELSE}
Autorise := ((V_PGI.PassWord = Pass) or (TobCourante.GetValue(TPar.FindField('Modifiable').FieldName) = 'O'));
{$ENDIF}
if TobCourante.FieldExists('Numero') then
    MENumero.Text       := TobCourante.GetValue('Numero');

MEDomaine.ItemIndex :=  MEDomaine.Values.IndexOf(TobCourante.GetValue('Domaine'));

if TobCourante.FieldExists('TableName') then
    METableName.Text    := TobCourante.GetValue('TableName');
if TobCourante.FieldExists('Nom') then
    MENom.Text          := TobCourante.GetValue('Nom');
if not Tous then Exit;
MENom.Enabled         := Autorise;
MEContenu.Text        := TobCourante.GetValue('Contenu');
MEContenu.Enabled     := Autorise;
MELongueur.value      := TobCourante.GetValue('Longueur');
MELongueur.Enabled    := Autorise;
MEObligatoire.Checked := (TobCourante.GetValue('Obligatoire') = 'O');
MEObligatoire.Enabled := Autorise;
if TobCourante.GetValue('MajusMin') = 'MAJ' then
    MEMajusMin.ItemIndex  := MEMajusMin.Items.IndexOf('Majuscule')
else if TobCourante.GetValue('MajusMin') = 'MIN' then
    MEMajusMin.ItemIndex  := MEMajusMin.Items.IndexOf('Minuscule')
    else
    MEMajusMin.ItemIndex  := 0;
MEMajusMin.Enabled    := Autorise;
if TobCourante.GetValue('TypeAlpha') = 'A' then
    METypeAlpha.ItemIndex      := METypeAlpha.Items.IndexOf('Alphanumérique')
else if TobCourante.GetValue('TypeAlpha') = 'N' then
    METypeAlpha.ItemIndex      := METypeAlpha.Items.IndexOf('Numérique')
else if TobCourante.GetValue('TypeAlpha') = 'D' then
    METypeAlpha.ItemIndex      := METypeAlpha.Items.IndexOf('Date')
else if TobCourante.GetValue('TypeAlpha') = 'H' then
    METypeAlpha.ItemIndex      := METypeAlpha.Items.IndexOf('Heure')
else
    METypeAlpha.ItemIndex      := 0;
METypeAlpha.Enabled        := Autorise;
MEContenuoblig.Text   := TobCourante.GetValue('ContenuOblig');
MEContenuoblig.Enabled:= Autorise;
MENomPGI.Text         := TobCourante.GetValue('NomPGI');
MENomPGI.Enabled      := Autorise;
MECommentaire.Text    := TobCourante.GetValue('Commentaire');
MECommentaire.Enabled := Autorise;
MELibre1.Text         := TobCourante.GetValue('Libre1');
MELibre1.Enabled      := Autorise;
MELibre2.Text         := TobCourante.GetValue('Libre2');
MELibre2.Enabled      := Autorise;
MELibre3.Text         := TobCourante.GetValue('Libre3');
MELibre3.Enabled      := Autorise;
MELibre4.Text         := TobCourante.GetValue('Libre4');
MELibre4.Enabled      := Autorise;
MELibre5.Text         := TobCourante.GetValue('Libre5');
MELibre5.Enabled      := Autorise;
MELibre6.Text         := TobCourante.GetValue('Libre6');
MELibre6.Enabled      := Autorise;
MELibre7.Text         := TobCourante.GetValue('Libre7');
MELibre7.Enabled      := Autorise;
MELibre8.Text         := TobCourante.GetValue('Libre8');
MELibre8.Enabled      := Autorise;
MELibre9.Text         := TobCourante.GetValue('Libre9');
MELibre9.Enabled      := Autorise;
MELibre10.Text        := TobCourante.GetValue('Libre10');
MELibre10.Enabled     := Autorise;
MEModifiable.Checked  := (TobCourante.GetValue('Modifiable') = 'O');
MEModifiable.Enabled  := Autorise;
MEFige.Checked        := (TobCourante.GetValue('Fige') = 'O');
MEFige.Enabled        := Autorise;
// ajout me
MEFamCorresp.Text         := TobCourante.GetValue('Famcorresp');
MEFamCorresp.Enabled      := Autorise;
MEOrdreTri.Text               := TobCourante.GetValue('Ordretri');
MEOrdreTri.Enabled            := Autorise;

if TobCourante.GetValue('Alignement') = 'G' then
    MEAlignement.ItemIndex      := MEAlignement.Items.IndexOf('Gauche')
else if TobCourante.GetValue('Alignement') = 'D' then
    MEAlignement.ItemIndex      := MEAlignement.Items.IndexOf('Droite')
    else
    MEAlignement.ItemIndex      := 0;
MEAlignement.Enabled  := Autorise;
MESeparateur.Text     := TobCourante.GetValue('Separateur');
MESeparateur.Enabled  := Autorise;
MECalcul.Text         := TobCourante.GetValue('Calcul');
MECalcul.Enabled      := Autorise;
// ajout me
MENbdecimal.value       := TobCourante.GetValue('Nbdecimal');
MENbdecimal.Enabled    := Autorise;
Offset := 1;
for ind1 := 0 to TobCourante.Parent.Detail.Count - 1 do
begin
    if TobCourante.Parent.Detail[ind1] = TobCourante then Break;
    Offset := Offset + TobCourante.Parent.Detail[ind1].GetValue('Longueur');
end;
TEOffset.Text := IntToStr(Offset);
end;

{$IFNDEF CISXPGI}
procedure TFParametre.EnregistreTob(TobM : TOB);
var ind1, ind2 : integer;
    TobF : TOB;
    cTemp,Filename : string;

    procedure InsereNewEnreg;
begin
    cTemp := Copy(TobF.GetValue('Modifiable'), 1, 1);
{$IFNDEF CISXPGI}
    if (cTemp[1] = 'N') or (cTemp[1] = 'O') then
    begin
        if not (V_PGI.PassWord = DayPass(Date)) then
            TobF.PutValue('Modifiable', Ord(cTemp[1]) - Ord('N'))
            else
            TobF.PutValue('Modifiable', 0);
    end
        else
{$ENDIF}
            TobF.PutValue('Modifiable', 0);

    cTemp := Copy(TobF.GetValue('Fige'), 1, 1);
    if (cTemp[1] = 'N') or (cTemp[1] = 'O') then
        TobF.putValue('Fige', Ord(cTemp[1]) - Ord('N') )
        else
        TobF.putValue('Fige', 0);
    TPar.InsertRecord([ TobF.GetValue('Numero'),
                        TobF.GetValue('Domaine'),
                        TobF.GetValue('TableName'),
                        TobF.GetValue('Nom'),
                        TobF.GetValue('Contenu'),
                        TobF.GetValue('Longueur'),
                        TobF.GetValue('Obligatoire'),
                        TobF.GetValue('NomPGI'),
                        TobF.GetValue('Commentaire'),
                        TobF.GetValue('Contenuoblig'),
                        TobF.GetValue('MajusMin'),
                        TobF.GetValue('TypeAlpha'),
                        TobF.GetValue('Libre1'),
                        TobF.GetValue('Libre2'),
                        TobF.GetValue('Libre3'),
                        TobF.GetValue('Libre4'),
                        TobF.GetValue('Libre5'),
                        TobF.GetValue('Libre6'),
                        TobF.GetValue('Libre7'),
                        TobF.GetValue('Libre8'),
                        TobF.GetValue('Libre9'),
                        TobF.GetValue('Libre10'),
                        TobF.GetValue('Modifiable'),
                        TobF.GetValue('Fige'),
                        TobF.GetValue('Alignement'),
                        TobF.GetValue('Separateur'),
                        TobF.GetValue('Nbdecimal'),
                        TobF.GetValue('NumVersion'),
                        TobF.GetValue('Calcul'),
                        // ajout me
                        TobF.GetValue('Famcorresp'),
                        TobF.GetValue('Ordretri')
                      ]);
end;

    procedure UpdateEnreg;
begin
//gcvoirtob(TobF);
    TPar.Edit;
    TPar.FindField('TableName').asstring    := TobF.GetValue('TableName');
    TPar.FindField('Nom.').asstring         := TobF.GetValue('Nom');
    TPar.FindField('Contenu').asstring      := TobF.GetValue('Contenu');
    TPar.FindField('Longueur').asinteger    := TobF.GetValue('Longueur');
    TPar.FindField('Obligatoire').asstring  := TobF.GetValue('Obligatoire');
    TPar.FindField('MajusMin').asstring     := TobF.GetValue('MajusMin');
    TPar.FindField('TypeAlpha').asstring    := TobF.GetValue('TypeAlpha');
    TPar.FindField('Contenuoblig').asstring := TobF.GetValue('Contenuoblig');
    TPar.FindField('NomPGI').asstring       := TobF.GetValue('NomPGI');
    TPar.FindField('Commentaire').asstring      := TobF.GetValue('Commentaire');
    TPar.FindField('Libre1').asstring       := TobF.GetValue('Libre1');
    TPar.FindField('Libre2').asstring       := TobF.GetValue('Libre2');
    TPar.FindField('Libre3').asstring       := TobF.GetValue('Libre3');
    TPar.FindField('Libre4').asstring       := TobF.GetValue('Libre4');
    TPar.FindField('Libre5').asstring       := TobF.GetValue('Libre5');
    TPar.FindField('Libre6').asstring       := TobF.GetValue('Libre6');
    TPar.FindField('Libre7').asstring       := TobF.GetValue('Libre7');
    TPar.FindField('Libre8').asstring       := TobF.GetValue('Libre8');
    TPar.FindField('Libre9').asstring       := TobF.GetValue('Libre9');
    TPar.FindField('Libre10').asstring      := TobF.GetValue('Libre10');
    cTemp := Copy(TobF.GetValue('Modifiable'), 1, 1);
{$IFNDEF CISXPGI}
    if (cTemp[1] = 'N') or (cTemp[1] = 'O') then
    begin
        if not (V_PGI.PassWord = DayPass(Date)) then
            TPar.FindField('Modifiable').asinteger := Ord(cTemp[1]) - Ord('N')
            else
            TPar.FindField('Modifiable').asinteger := 0;
    end
        else
{$ENDIF}
        TPar.FindField('Modifiable').asinteger := 0;

    cTemp := Copy(TobF.GetValue('Fige'), 1, 1);
    if (cTemp[1] = 'N') or (cTemp[1] = 'O') then
        TPar.FindField('Fige').asinteger := Ord(cTemp[1]) - Ord('N')
        else
        TPar.FindField('Fige').asinteger := 0;
    TPar.FindField('Alignement').asstring   := TobF.GetValue('Alignement');
    TPar.FindField('Separateur').asstring   := TobF.GetValue('Separateur');
    TPar.FindField('Nbdecimal').asstring    := TobF.GetValue('Nbdecimal');
    TPar.FindField('NumVersion').asstring   := TobF.GetValue('NumVersion');
    TPar.FindField('Calcul').asstring       := TobF.GetValue('Calcul');
// ajout me
    TPar.FindField('Famcorresp').asstring   := TobF.GetValue('Famcorresp');
    TPar.FindField('Ordretri').asstring     := TobF.GetValue('Ordretri');

    TPar.Post;
end;

begin
for ind1 := 0 to TobM.Detail.Count - 1 do
begin
    TobF := TobM.Detail[ind1];
    if TobF.Detail.Count > 0 then
        EnregistreTob(TobF)
    else
    begin
        if TobF.GetValue('Alignement') = '<' then
            TobF.PutValue('Alignement', '');
        if (TobF.FieldExists('Nouveau')) and (not TobF.FieldExists('A Supprimer')) then
        begin
//   Nouvel enregistrement : creation complete.
            cTemp := Copy(TobF.GetValue('Modifiable'), 1, 1);
            if (cTemp[1] = 'N') or (cTemp[1] = 'O') then
                TobF.PutValue('Modifiable', Ord(cTemp[1]) - Ord('N'))
                else
                TobF.PutValue('Modifiable', 0);
            cTemp := Copy(TobF.GetValue('Fige'), 1, 1);
            if (cTemp[1] = 'N') or (cTemp[1] = 'O') then
                TobF.PutValue('Fige', Ord(cTemp[1]) - Ord('N'))
                else
                TobF.PutValue('Fige', 0);
            ind2 := TPar.RecordCount + 1;
            while TPar.Locate('Numero', ind2, [loCaseInsensitive]) do
                Inc(ind2);
            TobF.PutValue('Numero', ind2);
            InsereNewEnreg;
        end
        else if (TobF.FieldExists('A Supprimer')) and (not TobF.FieldExists('Nouveau')) then
        begin
                if TPar.Locate('Numero;Domaine', VarArrayOf([TobF.GetValue('Numero'),
                             TobF.GetValue('Domaine')]), [loCaseInsensitive]) then
                TPar.Delete;
        end
        else
        begin
//   enregistrement modifié : on ne renseigne que les champs modifiés.
            if not TobF.FieldExists('Modifie') then Continue;
            if not TPar.Locate('Numero;Domaine', VarArrayOf([TobF.GetValue('Numero'),
                             TobF.GetValue('Domaine')]), [loCaseInsensitive]) then

                                 InsereNewEnreg;

            UpdateEnreg;
        end;
        TobF.DelChampSup('Nouveau', False);
        TobF.DelChampSup('Modifie', False);
    end;
end;
end;
{$ELSE}
procedure TFParametre.EnregistreTob(TobM : TOB);
var ind, ind1, ind2 : integer;
    TobF, TobR, TobL : TOB;
    cTemp,Filename : string;
    TobDico, TD        : TOB;
begin
         TobDico := TOB.Create ('Dictionnaire', nil, -1);
         for ind := 0 to TobM.Detail.Count - 1 do
         begin
                   TobR := TobM.Detail[ind];
                   for ind1 := 0 to TobR.Detail.Count - 1 do
                   begin
                      TobL := TobR.Detail[ind1];
                      for ind2 := 0 to TobL.Detail.Count - 1 do
                      begin
                                TobF := TOBL.detail[ind2];
                                if (TobF.FieldExists('A Supprimer'))then
                                        continue;
                                if TobF.GetValue('Alignement') = '<' then
                                    TobF.PutValue('Alignement', '');
                                if (TobF.FieldExists('Nouveau')) and (not TobF.FieldExists('A Supprimer')) then
                                begin
                                   //   Nouvel enregistrement : creation complete.
                                    cTemp := Copy(TobF.GetValue('Modifiable'), 1, 1);
                                    if (cTemp[1] = 'N') or (cTemp[1] = 'O') then
                                        TobF.PutValue('Modifiable', Ord(cTemp[1]) - Ord('N'))
                                        else
                                        TobF.PutValue('Modifiable', 0);
                                    cTemp := Copy(TobF.GetValue('Fige'), 1, 1);
                                    if (cTemp[1] = 'N') or (cTemp[1] = 'O') then
                                        TobF.PutValue('Fige', Ord(cTemp[1]) - Ord('N'))
                                        else
                                        TobF.PutValue('Fige', 0);
                                    TobF.PutValue('Numero', ind2+1);
                                end;
                                TobF.DelChampSup('Nouveau', False);
                                TobF.DelChampSup('Modifie', False);

                                TD := TOB.Create ('', TobDico, -1);
                                TD.Dupliquer(TobF, False, True);
                          end;
                  end;
                  if V_PGI.DosPath[length (V_PGI.DosPath)] <> '\' then
                  Filename :=  V_PGI.DosPath +'\'+LeDomaine+'Compta.Cix'
                  else Filename :=  V_PGI.DosPath +LeDomaine+'Compta.Cix';
                  TobDico.SaveToXMLFile(Filename, True, True);

                  AGL_YFILESTD_IMPORT(Filename,'CISX', LeDomaine+'Compta.Cix', '.CIX',
                  LeDomaine, 'COMPTA', 'PARAM', TobR.Getvalue ('Domaine'), '','-','-','-','-','-',
                  V_PGI.LanguePrinc, 'CEG', LTitre, V_PGI.NoDossier);
         end;

end;
{$ENDIF}

function TFParametre.SearchNodeFromTob(TobL : TOB) : TTreeNode;
begin
Result := nil;
if TV.Items.Count <= 0 then Exit;
Result := TV.Items.GetFirstNode;
while Result <> nil do
begin
    if (TOB(Result.Data).GetValue('Domaine') = TobL.GetValue('Domaine')) and
       (TOB(Result.Data).GetValue('TableName') = TobL.GetValue('TableName')) and
       (TOB(Result.Data).GetValue('Nom') = TobL.GetValue('Nom')) then
    begin
        Result.Selected := True;
        Exit;
    end;
    Result := Result.GetNext;
end;
end;

function TFParametre.SearchGridFromTob(TobL : TOB) : integer;
var ind1 : integer;
begin
Result := -1;
if TobL = nil then Exit;
for ind1 := 1 to SG1.RowCount - 1 do
    if SG1.Objects[0, ind1] = TobL then
    begin
        Result := ind1;
        Exit;
    end;
end;

procedure TFParametre.PositionneBoutons(Sender : TObject; TN1 : TTreeNode);
begin
if not (Sender is TToolbarButton97) then Exit;
if TToolbarButton97(Sender).Name = 'BPrev' then
begin
    if TN1.GetPrev.GetPrev = nil then
    begin
        bFirst.Enabled := False;
        bPrev.Enabled := False;
        bLast.Enabled := True;
        bNext.Enabled := True;
    end
        else
    begin
        bFirst.Enabled := True;
        bPrev.Enabled := True;
    end;
    if TN1.GetPrev.GetNext <> nil then
    begin
        bLast.Enabled := True;
        bNext.Enabled := True;
    end;
end
else if TToolbarButton97(Sender).Name = 'BNext' then
begin
    if TN1.GetNext.GetNext = nil then
    begin
        bLast.Enabled := False;
        bNext.Enabled := False;
        bFirst.Enabled := True;
        bPrev.Enabled := True;
    end
        else
    begin
        bLast.Enabled := True;
        bNext.Enabled := True;
    end;
    if TN1.GetNext.GetPrev <> nil then
    begin
        bFirst.Enabled := True;
        bPrev.Enabled := True;
    end;
end;
end;

procedure TFParametre.PutGrid(TobC : TOB; SG : TStringGrid; LCols : string);
var ind1, ARow, ACol : integer;
    TobL : TOB;
    ColsLoc, NomCol : string;
begin
if SG.RowCount <= TobC.Detail.Count then
    SG.RowCount := TobC.Detail.Count + 1;
ARow := 1;
for ind1 := 0 to TobC.Detail.Count - 1 do
begin
    TobL := TobC.Detail[ind1];
    SG.Objects[0, ARow] := TobL;
    ColsLoc := LCols;
    NomCol := ReadTokenSt(ColsLoc);
    while NomCol <> '' do
    begin
        for ACol := 0 to SG.ColCount - 1 do
            if UpperCase(SG.Cells[ACol, 0]) = UpperCase(NomCol) then Break;
        if ACol < SG.ColCount then
            SG.Cells[ACol, ARow] := VarAsType(TobL.GetValue(NomCol), varString);
        NomCol := ReadTokenSt(ColsLoc);
    end;
    Inc(ARow);
end;
end;

function TFParametre.ControleContenu(TMELoc : TMaskEdit) : boolean;
begin
Result := True;
if TMELoc.Name = 'MEContenu' then
begin
    if (MEFige.Checked) and (Trim(TMELoc.Text) = '') then
    begin
        HMsgBox.Execute(0, Caption, '');
        Result := False;
        Exit;
    end;
    if StrLen(PChar(TMELoc.Text)) > MELongueur.value then
    begin
        HMsgBox.Execute(1, Caption, '');
        Result := False;
        Exit;
    end;
    if Trim(METypeAlpha.Text) = '' then
        METypeAlpha.ItemIndex := METypeAlpha.Items.IndexOf('Alphanumérique');
    if UpperCase(METypeAlpha.Text[1]) = 'A' then
    begin
        if MEMajusMin.Text = 'MAJ' then
            TMELoc.Text := Uppercase(TMELoc.Text)
        else if MEMajusMin.Text = 'MIN' then
            TMELoc.Text := Lowercase(TMELoc.Text);
    end
        else
    begin
        if Trim(TMELoc.Text) = '' then Exit;
        if not IsNumeric(TMELoc.Text) then
        begin
            HMsgBox.Execute(2, Caption, '');
            Result := False;
            Exit;
        end;
        if StrLen(PChar(TMELoc.Text)) > 7 then
        begin
            HMsgBox.Execute(3, Caption, '');
            Result := False;
            Exit;
        end
    end;
end
else if TMELoc.Name = 'MEContenuOblig' then
begin
end;
end;

procedure TFParametre.TWCalculVisibleChanged(Sender: TObject);
begin
  if TWCalcul.Visible then
    OnFieldEnter(MECalcul)
  else
    OnFieldExit(MECalcul);
end;


procedure TFParametre.FormKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
     Case Key of
          VK_F6 : BEGIN
          if (TV.Selected.Level = 0) or  (TV.Selected.Level = 1) then exit;
          BinsertClick(Sender);
          if TobCourante.FieldExists('TableName') then
              METableName.Text    := TobCourante.GetValue('TableName');
          if TobCourante.FieldExists('Nom') then
              MENom.Text          := TobCourante.GetValue('Nom');
          if TobCourante.FieldExists('Contenu') then
          MEContenu.Text          := TobCourante.GetValue('Contenu');
          if TobCourante.FieldExists('Longueur')then
          MELongueur.value        := TobCourante.GetValue('Longueur');
          Key:=0 ;  MENom.SetFocus ;
          END ;
     end;
end;

// ajout me
procedure TFParametre.TVClick(Sender: TObject);
begin
     BDelete.Enabled := TRUE;
end;

{$IFNDEF CISXPGI}
procedure TFParametre.ChargementComboDomaine;
var
  Imptable: TADOTable;
begin
  if not Assigned(dmImport) then
    dmImport := TdmImport.Create(Application);
  ImpTable := DMImport.GzImpDomaine;
  with ImpTable do
  begin
    if Active then
      Close;
    Open;
    First;
    while not Eof do
    begin
        MEDomaine.ITems.add(ImpTable.FieldByName('Libelle').asstring);
        MEDomaine.Values.add(ImpTable.FieldByName('Domaine').asstring);
      Next;
    end;
    Close;
  end;
end;
{$ENDIF}


procedure TFParametre.METableNameClick(Sender: TObject);
begin
 if (METableName.text = '') and  (not ModeInsert)  then
     BInsertClick(Sender);
end;

procedure TFParametre.MENomClick(Sender: TObject);
begin
 if (MENom.text = '') and  (not ModeInsert)  then
     BInsertClick(Sender);
end;

end.
