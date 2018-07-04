unit ModTlEcr;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  hmsgbox, HSysMenu, DB,
  {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
  Grids, Hctrls, StdCtrls, Buttons,
  ExtCtrls, Hcompte, Mask, ComCtrls, Ent1, HEnt1, {Filtre,} Menus, ParamDat,
  Cpteutil, HTB97
  ,HPanel, UIUtil // MODIF PACK AVANCE pour gestion mode inside
  , UObjFiltres {JP 21/01/05 : FQ 15255}
  {$IFDEF MODENT1}
  , CPTypeCons
  {$ENDIF MODENT1}
   ;

procedure ModifSerieTableLibreEcr(QuelTab : String ; FSetTyp : SetttTypePiece) ;

type
  TFModTlEcr = class(TForm)
    Pages: TPageControl;
    PCritere: TTabSheet;
    Bevel1: TBevel;
    TE_JOURNAL: THLabel;
    TE_EXERCICE: THLabel;
    TE_DATECOMPTABLE: THLabel;
    TE_DATECOMPTABLE2: THLabel;
    TTL: TLabel;
    DATECOMPTABLE_: THCritMaskEdit;
    DATECOMPTABLE: THCritMaskEdit;
    JOURNAL: THValComboBox;
    EXERCICE: THValComboBox;
    TABLE: THValComboBox;
    PComplement: TTabSheet;
    Bevel2: TBevel;
    TE_GENERAL: THLabel;
    TAUXILIAIRE: THLabel;
    GENERAL: THCpteEdit;
    AUXILIAIRE: THCpteEdit;
    FListe: THGrid;
    Q: TQuery;
    HMTrad: THSystemMenu;
    HM: THMsgBox;
    TNATUREPIECE: THLabel;
    NATUREPIECE: THValComboBox;
    TAncVal: TLabel;
    AncVal: THCpteEdit;
    SaveDialog: TSaveDialog;
    Axe: THValComboBox;
    TE_NUMEROPIECE: THLabel;
    NUMEROPIECE: THCritMaskEdit;
    HLabel1: THLabel;
    NUMEROPIECE_: THCritMaskEdit;
    TNewVal: TLabel;
    NewVal: THCpteEdit;
    TE_QUALIFPIECE: THLabel;
    QUALIFPIECE: THValComboBox;
    POPF: TPopupMenu;
    BCreerFiltre: TMenuItem;
    BSaveFiltre: TMenuItem;
    BDelFiltre: TMenuItem;
    BRenFiltre: TMenuItem;
    BNouvRech: TMenuItem;
    Nb1: TLabel;
    Tex1: TLabel;
    Dock: TDock97;
    PanelBouton: TToolWindow97;
    BTag: TToolbarButton97;
    BValider: TToolbarButton97;
    BFerme: TToolbarButton97;
    BAide: TToolbarButton97;
    FSearchTimer: TTimer;
    PanelTxt: TPanel;
    BDetag: TToolbarButton97;
    Dock971: TDock97;
    PanelFiltres: TToolWindow97;
    BFiltre: TToolbarButton97;
    FFiltres: THValComboBox;
    BCherche_: TToolbarButton97;
    BCherche: TToolbarButton97;
    TITREE: THComboBox;
    TITREA: THComboBox;
    TITREU: THComboBox;


    procedure EXERCICEChange(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure TABLEChange(Sender: TObject);
    procedure BFermeClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure AxeChange(Sender: TObject);
    procedure DATECOMPTABLEKeyPress(Sender: TObject; var Key: Char);
    procedure BChercheClick(Sender: TObject);
    procedure FListeKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure FListeMouseDown(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
    procedure FListeDblClick(Sender: TObject);
    procedure BdetagClick(Sender: TObject);
    procedure BTagClick(Sender: TObject);
    procedure BValiderClick(Sender: TObject);
    procedure BAideClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);


  private
    ObjFiltre : TObjFiltre; {JP 21/01/05 : Gestion des filtres FQ 15255}
    QuelTab : String ;
    FNomFiltre : String ;
    WMinX,WMinY : Integer ;
    TotalSelec : Integer ;
    RunMaj : Boolean ;
    Champ : String ;
    SetTyp : SetttTypePiece ;
    FAutoSearch     : TAutoSearch;
    FCritModified   : Boolean;        // Indicateur de modifiaction dans le PAGECONTROL
//    FSearchTimer    : TTimer;
    FLoading        : Boolean;        // Indicateur pour en cours de chargement
    procedure WMGetMinMaxInfo(var MSG: Tmessage); message WM_GetMinMaxInfo;
    Procedure InitLeCritere ;
    Procedure FabriqueLaRequete ;
    Procedure RempliLaListe ;
    Procedure CompteElemSelectionner ;
    Procedure TagDetag(Avec : Boolean) ;
    Function  ControleCpteOk : Boolean ;
    Procedure GetCellCanvas(Acol,ARow : LongInt ; Canvas : TCanvas; AState: TGridDrawState) ;

    // YMO 04/10/2005 Ajout Autosearch + bouton Recherche standard (FQ16657)
    procedure InitAutoSearch;
    procedure CritereChange    (Sender: TObject);
    procedure SetCritModified  (Value: Boolean);
    procedure SearchTimerTimer (Sender: TObject);
    procedure ResetTimer       (Sender: TObject);
    property CritModified         : Boolean      read FCritModified        write SetCritModified;
    property AutoSearch           : TAutoSearch  read FAutoSearch          write FAutoSearch;

  public
    { Déclarations publiques }
    procedure ApresChargementFiltre;{JP 24/01/05 : FQ 15255}
  end;

implementation

{$R *.DFM}

uses
  {$IFDEF MODENT1}
  ULibExercice,
  CPProcMetier,
  {$ENDIF MODENT1}
  HStatus;

var FModTlEcr : TFModTlEcr ; // YM0 04/10/2005 remontée de la déclaration
                             // de la proc 'ModifSerieTableLibreEcr' (FQ16657)


{***********A.G.L.***********************************************
Auteur  ...... :
Créé le ...... : 09/09/2003
Modifié le ... : 09/09/2003
Description .. :
Suite ........ : 09/09/2003, SBO : MODIF PACK AVANCE pour gestion
Suite ........ : mode inside
Mots clefs ... :
*****************************************************************}
procedure ModifSerieTableLibreEcr(QuelTab : String ; FSetTyp : SetttTypePiece) ;
var
    PP : THPanel ;
begin
  FModTlEcr:=TFModTlEcr.Create(Application) ;
  FModTlEcr.FNomFiltre:='MODTLECR'+QuelTab[1] ;
  FModTlEcr.QuelTab:=QuelTab ;
  FModTlEcr.SetTyp:=FSetTyp ;

  PP:=FindInsidePanel ;
  if PP=Nil then
    begin
    try
      FModTlEcr.ShowModal ;
      finally
      FModTlEcr.Free ;
      end;
    end
  else
    begin
    InitInside(FModTlEcr,PP) ;
    FModTlEcr.Show ;
    end ;

  SourisNormale ;
end;

procedure TFModTlEcr.FormCreate(Sender: TObject);
var
  Composants : TControlFiltre;
begin

  {JP 21/01/05 : Gestion des filtres FQ 15255}
  Composants.Filtres  := FFILTRES;
  Composants.Filtre   := BFILTRE;
  Composants.PageCtrl := Pages;
  ObjFiltre := TObjFiltre.Create(Composants, '');
  ObjFiltre.ApresChangementFiltre := ApresChargementFiltre;

  WMinX:=Width ;
  WMinY:=Height ;
  // YMO 04/10/2005 Composant timer (FQ16657)
  FSearchTimer.Enabled := False;
  FSearchTimer.OnTimer := SearchTimerTimer;
end;

procedure TFModTlEcr.WMGetMinMaxInfo(var MSG: Tmessage);
begin
with PMinMaxInfo(MSG.lparam)^.ptMinTrackSize do begin X := WMinX ; Y := WMinY ; end;
end;

procedure TFModTlEcr.BFermeClick(Sender: TObject);
begin Close ; end;

procedure TFModTlEcr.EXERCICEChange(Sender: TObject);
begin
ExoToDates(EXERCICE.Value,DATECOMPTABLE,DATECOMPTABLE_) ;
end;

Procedure TFModTlEcr.InitLeCritere ;
BEGIN
if QuelTab='' then Exit ;
ChargeComboTableLibre(QuelTab[1],TABLE.Values,TABLE.Items) ;
Case QuelTab[1] of
     'A' : BEGIN
           JourNal.DataType:='ttJournal' ; Tauxiliaire.Caption:=HM.Mess[2] ; Auxiliaire.Visible:=True ;
           Axe.Visible:=True ; Axe.Value:=Axe.Values[0] ; NATUREPIECE.Visible:=False ; TNaturePiece.Caption:=HM.Mess[1] ;
           FListe.ColCount:=8 ; FListe.Titres:=TitreA.Items ;
           FListe.ColWidths[0]:=40 ; FListe.ColWidths[1]:=75 ;
           FListe.ColWidths[2]:=50 ; FListe.ColWidths[3]:=92 ; FListe.ColWidths[4]:=92 ;
           FListe.ColWidths[5]:=92 ; FListe.ColWidths[6]:=92 ; FListe.ColWidths[7]:=0 ;
           Caption:=HM.Mess[11] ; HelpContext:=7370500 ;
           END ;
     'E' : BEGIN
           JourNal.DataType:='ttJournal' ; Tauxiliaire.Caption:=HM.Mess[0] ; Auxiliaire.Visible:=True ;
           Axe.Visible:=False ; NATUREPIECE.Visible:=True ; NATUREPIECE.DataType:='ttNaturePiece' ;
           FListe.ColCount:=8 ; FListe.Titres:=TitreE.Items ;
           FListe.ColWidths[0]:=40 ; FListe.ColWidths[1]:=75 ;
           FListe.ColWidths[2]:=50 ; FListe.ColWidths[3]:=92 ; FListe.ColWidths[4]:=92 ;
           FListe.ColWidths[5]:=92 ; FListe.ColWidths[6]:=92 ; FListe.ColWidths[7]:=0 ;
           Caption:=HM.Mess[10] ;
           END ;
     'U' : BEGIN
           JourNal.DataType:='ttBudJal' ; Tauxiliaire.Caption:=HM.Mess[1] ; Auxiliaire.Visible:=False ;
           General.ZoomTable:=tzBudgen ; Axe.Visible:=True ; Axe.Top:=10 ; If Axe.Values.Count>0 Then Axe.Value:=Axe.Values[0] ;
           NATUREPIECE.Visible:=True ; NATUREPIECE.DataType:='ttNatEcrBud' ;
           FListe.ColCount:=8 ; FListe.Titres:=TitreU.Items ;
           FListe.ColWidths[0]:=50 ; FListe.ColWidths[1]:=76 ;
           FListe.ColWidths[2]:=61 ; FListe.ColWidths[3]:=61 ; FListe.ColWidths[4]:=95 ;
           FListe.ColWidths[5]:=95 ; FListe.ColWidths[6]:=95 ; FListe.ColWidths[7]:=0 ;
           Caption:=HM.Mess[12] ;
           END ;
  End ;
END ;

procedure TFModTlEcr.FormShow(Sender: TObject);
begin
FListe.GetCellCanvas:=GetCellCanvas ;
InitLeCritere ; Pages.ActivePage:=Pages.Pages[0] ;
  {JP 21/01/05 : Gestion des filtres FQ 15255}
  ObjFiltre.FFI_TABLE := FNomFiltre;
  ObjFiltre.Charger;
//ChargeFiltre(FNomFiltre,FFiltres,Pages) ;
if Journal.Items.Count>0 then JOURNAL.ItemIndex:=0 ;
if TABLE.Values.Count>0 then TABLE.Value:=TABLE.Values[0] ;
NATUREPIECE.ItemIndex:=0 ; RunMaj:=False ;
if SetTyp<>[] then
   BEGIN
   QUALIFPIECE.Enabled:=False ;
   if tpReel in SetTyp then BEGIN HelpContext:=7259500 ; QUALIFPIECE.Value:='N' ; END else
   if tpSim  in SetTyp then BEGIN HelpContext:=7280500 ; QUALIFPIECE.Value:='S' ; END else
   if tpSitu in SetTyp then BEGIN HelpContext:=7301500 ; QUALIFPIECE.Value:='U' ; END else
   if tpPrev in SetTyp then BEGIN HelpContext:=7316500 ; QUALIFPIECE.Value:='P'; END else
      QUALIFPIECE.Enabled:=True ;
   END else
   BEGIN
   QUALIFPIECE.Enabled:=True ;
   END ;
if VH^.CPExoRef.Code<>'' then
   BEGIN
   EXERCICE.Value:=VH^.CPExoRef.Code ;
   DATECOMPTABLE.Text:=DateToStr(VH^.CPExoRef.Deb) ;
   DATECOMPTABLE_.Text:=DateToStr(VH^.CPExoRef.Fin) ;
   END else
   BEGIN
   EXERCICE.Value:=VH^.Entree.Code ;
   DATECOMPTABLE.Text:=DateToStr(V_PGI.DateEntree) ;
   DATECOMPTABLE_.Text:=DateToStr(V_PGI.DateEntree) ;
   END ;
//  if FFiltres.Text = 'DEFAUT' then BEGIN FFiltresChange(Nil) ; BChercheClick(Nil) ; END ;
  {JP 21/01/05 : Gestion des filtres FQ 15255}
  ApresChargementFiltre ;
  InitAutoSearch ;
end;

procedure TFModTlEcr.TABLEChange(Sender: TObject);
begin
  {JP 21/01/05 : Gestion des filtres FQ 15255}
  if TABLE.Value <> '' then begin
    AncVal.ZoomTable:=NatureToTz(TABLE.Value) ;
    NewVal.ZoomTable:=AncVal.ZoomTable ;
  end;
end;

procedure TFModTlEcr.AxeChange(Sender: TObject);
begin
  {JP 21/01/05 : Gestion des filtres FQ 15255}
  if (Axe.Value <> '') and (Length(Axe.Value) > 1) then begin
    if (QuelTab='E') or (QuelTab='U') then Exit ;
    Auxiliaire.ZoomTable:=AxeToTz(Axe.Value) ;
  end;
end;

procedure TFModTlEcr.DATECOMPTABLEKeyPress(Sender: TObject; var Key: Char);
begin ParamDate(Self,Sender,Key) ; end;

procedure TFModTlEcr.BChercheClick(Sender: TObject);
begin
FLoading := true;
FListe.VidePile(False) ;
RunMaj:=True ;
FabriqueLaRequete ;
RempliLaListe ;
CritModified := False;
FLoading := false;
end;

Procedure TFModTlEcr.FabriqueLaRequete ;
Var Sql : String ;
BEGIN
if QuelTab='' then Exit ;
if Length(Table.Value)<3 then Exit ;
Case QuelTab[1] of
     'A' : BEGIN
           Champ:='Y_TABLE'+TABLE.Value[3]+' ' ;
           Sql:='Select Y_JOURNAL,Y_DATECOMPTABLE,Y_NUMEROPIECE,Y_GENERAL,Y_SECTION,Y_DEBIT,Y_CREDIT,'+Champ+
                'From ANALYTIQ Where '+
                'Y_DATECOMPTABLE>="'+USDate(DATECOMPTABLE)+'" And Y_DATECOMPTABLE<="'+USDate(DATECOMPTABLE_)+'" And '+
                'Y_TABLE'+TABLE.Value[3]+'="'+AncVal.Text+'" And '+
                'Y_AXE="'+Axe.Value+'" ' ;
           if Journal.Value<>'' then Sql:=Sql+'And Y_JOURNAL="'+Journal.Value+'" ' ;
           if Exercice.Value<>'' then Sql:=Sql+'And Y_EXERCICE="'+Exercice.Value+'" ' ;
           if General.Text<>'' then Sql:=Sql+'And Y_GENERAL Like "'+General.Text+'%" ' ;
           if Auxiliaire.Text<>'' then Sql:=Sql+'And Y_SECTION Like "'+Auxiliaire.Text+'%" ' ;
           if Numeropiece.Text<>'' then Sql:=Sql+'And Y_NUMEROPIECE>='+Numeropiece.Text+' ' ;
           if Numeropiece_.Text<>'' then Sql:=Sql+'And Y_NUMEROPIECE<='+Numeropiece_.Text+' ' ;
           END ;
     'E' : BEGIN
           Champ:='E_TABLE'+TABLE.Value[3]+' ' ;
           Sql:='Select E_JOURNAL,E_DATECOMPTABLE,E_NUMEROPIECE,E_GENERAL,E_AUXILIAIRE,E_DEBIT,E_CREDIT,'+Champ+
                'From ECRITURE Where '+
                'E_DATECOMPTABLE>="'+USDate(DATECOMPTABLE)+'" And E_DATECOMPTABLE<="'+USDate(DATECOMPTABLE_)+'" And '+
                'E_TABLE'+TABLE.Value[3]+'="'+AncVal.Text+'" And '+
                'E_ECRANOUVEAU="N" ';
           if NaturePiece.Value<>'' then Sql:=Sql+'And E_NATUREPIECE="'+NaturePiece.Value+'" ';
           if Journal.Value<>'' then Sql:=Sql+'And E_JOURNAL="'+Journal.Value+'" ' ;
           if Exercice.Value<>'' then Sql:=Sql+'And E_EXERCICE="'+Exercice.Value+'" ' ;
           if General.Text<>'' then Sql:=Sql+'And E_GENERAL Like "'+General.Text+'%" ' ;
           if Auxiliaire.Text<>'' then Sql:=Sql+'And E_AUXILIAIRE Like "'+Auxiliaire.Text+'%" ' ;
           if Numeropiece.Text<>'' then Sql:=Sql+'And E_NUMEROPIECE>='+Numeropiece.Text+' ' ;
           if Numeropiece_.Text<>'' then Sql:=Sql+'And E_NUMEROPIECE<='+Numeropiece_.Text+' ' ;
           If SetTyp<>[] Then Sql:=Sql+WhereSupp('E_',SetTyp) ;
           END ;
     'U' : BEGIN
           Champ:='BE_TABLE'+TABLE.Value[3]+' ' ;
           Sql:='Select BE_BUDJAL,BE_DATECOMPTABLE,BE_NUMEROPIECE,BE_NATUREBUD,BE_BUDGENE,BE_BUDSECT,BE_ETABLISSEMENT,'+Champ+
                'From BUDECR Where '+
                'BE_DATECOMPTABLE>="'+USDate(DATECOMPTABLE)+'" And BE_DATECOMPTABLE<="'+USDate(DATECOMPTABLE_)+'" And '+
                'BE_TABLE'+TABLE.Value[3]+'="'+AncVal.Text+'" ' ;
           if NaturePiece.Value<>'' then Sql:=Sql+'And BE_NATUREBUD="'+NaturePiece.Value+'" ';
           if Journal.Value<>'' then Sql:=Sql+'And BE_BUDJAL="'+Journal.Value+'" ' ;
           if Exercice.Value<>'' then Sql:=Sql+'And BE_EXERCICE="'+Exercice.Value+'" ' ;
           if General.Text<>'' then Sql:=Sql+'And BE_BUDGENE Like"'+General.Text+'%" ' ;
           if Numeropiece.Text<>'' then Sql:=Sql+'And BE_NUMEROPIECE>='+Numeropiece.Text+' ' ;
           if Numeropiece_.Text<>'' then Sql:=Sql+'And BE_NUMEROPIECE<='+Numeropiece_.Text+' ' ;
           END ;
  End ;
Q.Sql.Clear ; Q.Sql.Add(Sql) ; ChangeSql(Q) ;
END ;

Procedure TFModTlEcr.RempliLaListe ;
Var i : Integer ;
BEGIN
Q.Open ; i:=2 ;
While Not Q.Eof do
  BEGIN
  Fliste.RowCount:=i ;
  FListe.Cells[0,Fliste.RowCount-1]:=Q.Fields[0].AsString ;
  FListe.Cells[1,Fliste.RowCount-1]:=Q.Fields[1].AsString ;
  FListe.Cells[2,Fliste.RowCount-1]:=Q.Fields[2].AsString ;
  FListe.Cells[3,Fliste.RowCount-1]:=Q.Fields[3].AsString ;
  FListe.Cells[4,Fliste.RowCount-1]:=Q.Fields[4].AsString ;
  FListe.Cells[5,Fliste.RowCount-1]:=Q.Fields[5].AsString ;
  FListe.Cells[6,Fliste.RowCount-1]:=Q.Fields[6].AsString ;
  FListe.Cells[FListe.ColCount-1,Fliste.RowCount-1]:='*' ;
  Inc(i) ; Q.Next ;
  END ;
Q.Close ; FListe.Invalidate ; CompteElemSelectionner ;
END ;

Procedure TFModTlEcr.GetCellCanvas(Acol,ARow : LongInt ; Canvas : TCanvas; AState: TGridDrawState) ;
BEGIN
if FListe.Cells[FListe.ColCount-1,ARow]='*' then FListe.Canvas.Font.Style:=FListe.Canvas.Font.Style+[fsItalic]
                                            else FListe.Canvas.Font.Style:=FListe.Canvas.Font.Style-[fsItalic] ;
END ;

procedure TFModTlEcr.FListeKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
if ((Shift=[]) And (Key=VK_SPACE)) or ((ssShift in Shift) And (Key=VK_DOWN)) or
   ((ssShift in Shift) And (Key=VK_UP))then FlisteMouseDown(Nil,mbLeft,[ssCtrl],0,0) ;
{if (Shift=[]) And (Key=VK_SPACE) then // 13908
    BEGIN
    FlisteMouseDown(Nil,mbLeft,[ssCtrl],0,0) ;}
    if ((FListe.Row<FListe.RowCount-1) and (Key=VK_SPACE)) then FListe.Row:=FListe.Row+1 ;
//    END ;
end;

procedure TFModTlEcr.FListeMouseDown(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
if Not(ssCtrl in Shift) then Exit ;
if Button<>mbLeft then Exit ;
if Fliste.Cells[FListe.ColCount-1,FListe.Row]='*'
   then Fliste.Cells[FListe.ColCount-1,FListe.Row]:=''
   else Fliste.Cells[FListe.ColCount-1,FListe.Row]:='*' ;
FListe.Invalidate ; CompteElemSelectionner ;
end;

procedure TFModTlEcr.FListeDblClick(Sender: TObject); // 13908
var
  Key : Word;
begin
  Key := VK_DOWN;
  FListeKeyDown(Sender, Key, [ssShift]);
end;

Procedure TFModTlEcr.CompteElemSelectionner ;
Var i : Integer ;
BEGIN
TotalSelec:=0 ;
for i:=1 to FListe.RowCount-1 do
    if FListe.Cells[FListe.ColCount-1,i]='*' then Inc(TotalSelec) ;
Nb1.Caption:=IntToStr(TotalSelec) ;
if TotalSelec>1 then Tex1.Caption:=HM.Mess[4] else Tex1.Caption:=HM.Mess[3] ;
END ;

procedure TFModTlEcr.BdetagClick(Sender: TObject);
begin TagDetag(False) ; end;

procedure TFModTlEcr.BTagClick(Sender: TObject);
begin TagDetag(True) ; end;

Procedure TFModTlEcr.TagDetag(Avec : Boolean) ;
Var  i : Integer ;
begin
if Fliste.Cells[0,1]='' then Exit ;
for i:=1 to FListe.RowCount-1 do
    if Avec then FListe.Cells[FListe.ColCount-1,i]:='*'
            else FListe.Cells[FListe.ColCount-1,i]:='' ;
FListe.Invalidate ; Bdetag.Visible:=Avec ; BTag.Visible:=Not Avec ; CompteElemSelectionner ;
end;

Function TFModTlEcr.ControleCpteOk : Boolean ;
BEGIN
Result:=False ;
if NewVal.Text<>'' then
  if NewVal.ExisteH<=0 then
     BEGIN
     Pages.ActivePage:=PComplement ; NewVal.SetFocus ; HM.Execute(5,'','') ; Exit ;
     END ;
Result:=True ;
END ;

procedure TFModTlEcr.BValiderClick(Sender: TObject);
Var i : Integer ;
begin
if Not RunMaj then Exit ;
if FListe.Cells[0,1]='' then BEGIN HM.Execute(6,'','') ; Exit ; END ;
if TotalSelec=0 then BEGIN HM.Execute(7,'','') ; Exit ; END ;
if Not ControleCpteOk then Exit ;
if NewVal.Text=AncVal.Text then BEGIN HM.Execute(9,'','') ; Exit ; END ;
if HM.Execute(8,'','')<>mrYes then Exit ;
Q.Open ; Q.First ;
InitMove(FListe.RowCount-1,'') ;
for i:=1 to FListe.RowCount-1 do
   BEGIN
   MoveCur(False) ;
   if FListe.Cells[FListe.ColCount-1,i]='*' then
      BEGIN
      Q.Edit ;
      Q.FindField(Trim(Champ)).AsString:=NewVal.Text ;
      Q.Post ; Q.Next ;
      END else Q.Next ;
   END ;
FiniMove ; BChercheClick(Nil) ;
end;

procedure TFModTlEcr.BAideClick(Sender: TObject);
begin
CallHelptopic(Self) ;
end;

{---------------------------------------------------------------------------------------}
procedure TFModTlEcr.FormClose(Sender: TObject; var Action: TCloseAction);
{---------------------------------------------------------------------------------------}
begin
  {JP 21/01/05 : Gestion des filtres FQ 15255}
  FreeAndNil(ObjFiltre);
  FreeAndNil(FSearchTimer);
end;

{---------------------------------------------------------------------------------------}
procedure TFModTlEcr.ApresChargementFiltre;
{---------------------------------------------------------------------------------------}
begin
  {JP 24/01/05 : On lance la recherche après chargement du filtre}
  if FFiltres.Text <> '' then BChercheClick(BCherche);
end;


{***********A.G.L.***********************************************
Auteur  ...... : Yann MORENO
Créé le ...... : 04/10/2005
Modifié le ... :   /  /
Description .. : Ajout Autosearch + bouton Recherche standard (FQ16657)
Suite ........ : Copie à partir de uTofViergeMul
Mots clefs ... : YMO
*****************************************************************}
procedure TFModTlEcr.InitAutoSearch;
var
  i: integer;
begin

  if (V_PGI.AutoSearch) and (AutoSearch <> AsMouetteForce) then
    AutoSearch := asTimer;

  for i := 0 to FModTlEcr.ComponentCount - 1 do
  begin
    if (Components[i] is TControl) and
      (TControl(Components[i]).Parent is TTabSheet) then
      if (Components[i] is TControl) and
        TControl(Components[i]).Visible and
        (TControl(Components[i]).Enabled) then
      begin
        if (Components[i] is THCritMaskEdit)
        and (not Assigned(THCritMaskEdit(Components[i]).OnEnter)) then
          THCritMaskEdit(Components[i]).OnEnter := V_PGI.EgaliseOnEnter;

        case AutoSearch of
          asChange:
            begin
              if (Components[i] is TEdit) and not
                assigned(TEdit(Components[i]).OnChange) then
                TEdit(Components[i]).OnChange := SearchTimerTimer;
              if (Components[i] is THValComboBox) and not
                assigned(THValComboBox(Components[i]).OnClick) then
                THValComboBox(Components[i]).OnClick := SearchTimerTimer;
              if (Components[i] is TCheckBox) and not
                assigned(TCheckBox(Components[i]).OnClick) then
                TCheckBox(Components[i]).OnClick := SearchTimerTimer;
              if (Components[i] is THCritMaskEdit) and not
                assigned(THCritMaskEdit(Components[i]).OnChange) then
                THCritMaskEdit(Components[i]).OnChange :=
                  SearchTimerTimer;
              if (Components[i] is THRadioGroup) and not
                assigned(THRadioGroup(Components[i]).OnClick) then
                THRadioGroup(Components[i]).OnClick := SearchTimerTimer;
            end;
          asExit:
            begin
              if (Components[i] is TEdit) and not
                assigned(TEdit(Components[i]).onExit) then
                TEdit(Components[i]).onExit := SearchTimerTimer;
              if (Components[i] is THValComboBox) and not
                assigned(THValComboBox(Components[i]).onExit) then
                THValComboBox(Components[i]).onExit := SearchTimerTimer;
              if (Components[i] is TCheckBox) and not
                assigned(TCheckBox(Components[i]).onExit) then
                TCheckBox(Components[i]).onExit := SearchTimerTimer;
              if (Components[i] is THCritMaskEdit) and not
                assigned(THCritMaskEdit(Components[i]).onExit) then
                THCritMaskEdit(Components[i]).onExit := SearchTimerTimer;
              if (Components[i] is THRadioGroup) and not
                assigned(THRadioGroup(Components[i]).onExit) then
                THRadioGroup(Components[i]).onExit := SearchTimerTimer;
            end;
          asTimer:
            begin
              if (Components[i] is TEdit) and not
                assigned(TEdit(Components[i]).OnChange) then
                TEdit(Components[i]).OnChange := ResetTimer;
              if (Components[i] is THValComboBox) and not
                assigned(THValComboBox(Components[i]).OnClick) then
                THValComboBox(Components[i]).OnClick := ResetTimer;
              if (Components[i] is TCheckBox) and not
                assigned(TCheckBox(Components[i]).OnClick) then
                TCheckBox(Components[i]).OnClick := ResetTimer;
              if (Components[i] is THCritMaskEdit) and not
                assigned(THCritMaskEdit(Components[i]).OnChange) then
                THCritMaskEdit(Components[i]).OnChange := ResetTimer;
              if (Components[i] is THRadioGroup) and not
                assigned(THRadioGroup(Components[i]).OnClick) then
                THRadioGroup(Components[i]).OnClick := ResetTimer;
            end;
        else
          begin
            if (Components[i] is TEdit) and not
              assigned(TEdit(Components[i]).OnChange) then
              TEdit(Components[i]).OnChange := CritereChange;
            if (Components[i] is THValComboBox) and not
              assigned(THValComboBox(Components[i]).OnClick) then
              THValComboBox(Components[i]).OnClick := CritereChange;
            if (Components[i] is TCheckBox) and not
              assigned(TCheckBox(Components[i]).OnClick) then
              TCheckBox(Components[i]).OnClick := CritereChange;
            if (Components[i] is THCritMaskEdit) and not
              assigned(THCritMaskEdit(Components[i]).OnChange) then
              THCritMaskEdit(Components[i]).OnChange := CritereChange;
            if (Components[i] is THRadioGroup) and not
              assigned(THRadioGroup(Components[i]).OnClick) then
              THRadioGroup(Components[i]).OnClick := CritereChange;
          end;

        end;
      end;
  end;

end;

{***********A.G.L.***********************************************
Auteur  ...... : Yann MORENO
Créé le ...... : 04/10/2005
Modifié le ... :   /  /
Description .. : Copie à partir de uTofViergeMul, pour importer le
Suite ........ : comportement du bouton de recherche (avec l'entonnoir) (FQ16657)
Mots clefs ... : YMO
*****************************************************************}
procedure TFModTlEcr.CritereChange(Sender: TObject);
begin
  CritModified := True;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Yann MORENO
Créé le ...... : 04/10/2005
Modifié le ... :   /  /
Description .. : Copie à partir de uTofViergeMul, pour importer le
Suite ........ : comportement du bouton de recherche (avec l'entonnoir) (FQ16657)
Mots clefs ... : YMO
*****************************************************************}
procedure TFModTlEcr.SetCritModified(Value: Boolean);
begin
  FCritModified     := Value;
  BCHERCHE.Visible  := not Value;
  BCHERCHE_.Visible := Value;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Yann MORENO
Créé le ...... : 04/10/2005
Modifié le ... :   /  /
Description .. : Ajout fonction Autosearch  (FQ16657)
Suite ........ : Copie à partir de uTofViergeMul
Mots clefs ... : YMO
*****************************************************************}
procedure TFModTlEcr.SearchTimerTimer(Sender: TObject);
begin
  FSearchTimer.Enabled := False;
  CritModified := True;
  BChercheClick(nil);
end;

////////////////////////////////////////////////////////////////////////////////

{***********A.G.L.***********************************************
Auteur  ...... : Yann MORENO
Créé le ...... : 04/10/2005
Modifié le ... :   /  /
Description .. : Ajout fonction Autosearch  (FQ16657)
Suite ........ : Copie à partir de uTofViergeMul
Mots clefs ... : YMO
*****************************************************************}
procedure TFModTlEcr.ResetTimer(Sender: TObject);
begin
  if FLoading then Exit;

  CritModified := True;
  FSearchTimer.Enabled := False;
  FSearchTimer.Enabled := True;
end;

end.
