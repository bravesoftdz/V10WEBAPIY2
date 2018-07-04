{***********UNITE*************************************************
Auteur  ...... : Vincent Laroche
Créé le ...... : 26/01/2005
Modifié le ... :   /  /
Description .. : Passage en eAGL
Mots clefs ... :
*****************************************************************}
unit ModiTali;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs, Ent1,
  Hent1, hmsgbox, Hcompte, StdCtrls, Hctrls, Buttons, ExtCtrls,
{$IFDEF EAGLCLIENT}
  UTob,
{$ELSE}
  DB,
  {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ENDIF}
  Grids, ComCtrls, HSysMenu, Menus, HTB97,
  HPanel,
  {$IFDEF MODENT1}
  CPTypeCons,
  {$ENDIF MODENT1}
  UIUtil, // MODIF PACK AVANCE pour gestion mode inside
  UObjFiltres {JP 21/01/05 : FQ 15255}
  ;

Procedure ModifSerieTableLibre(Unfb : TFichierBase) ;

type
  TFModiTali = class(TForm)
    MsgBox: THMsgBox;
    Pages: TPageControl;
    PStandards: TTabSheet;
    TCpteDe: TLabel;
    CpteDe: THCpteEdit;
    TCpteA: TLabel;
    CpteA: THCpteEdit;
    TAxe: TLabel;
    Axe: THValComboBox;
    TTL: TLabel;
    TL: THValComboBox;
    Bevel1: TBevel;
    TAncVal: TLabel;
    AncVal: THCpteEdit;
    CpteDeJ: TEdit;
    FListe: THGrid;
    HMTrad: THSystemMenu;
    PCompGen: TTabSheet;
    NatGen: THValComboBox;
    TNatGen: THLabel;
    Ventil: THValComboBox;
    TVentil: THLabel;
    Bevel2: TBevel;
    G_COLLECTIF: TCheckBox;
    G_FERME: TCheckBox;
    G_LETTRABLE: TCheckBox;
    G_POINTABLE: TCheckBox;
    PCompTiers: TTabSheet;
    HLabel1: THLabel;
    T_NATUREAUXI: THValComboBox;
    TT_COLLECTIF: THLabel;
    T_COLLECTIF: THCpteEdit;
    TT_Secteur: THLabel;
    T_SECTEUR: THValComboBox;
    TT_REGIMETVA: THLabel;
    T_REGIMETVA: THValComboBox;
    T_LETTRABLE: TCheckBox;
    T_MULTIDEVISE: TCheckBox;
    Bevel3: TBevel;
    POPF: TPopupMenu;
    BCreerFiltre: TMenuItem;
    BSaveFiltre: TMenuItem;
    BDelFiltre: TMenuItem;
    BRenFiltre: TMenuItem;
    BNouvRech: TMenuItem;
    Panel2: TPanel;
    TNewVal: TLabel;
    NewVal: THCpteEdit;
    Dock971: TDock97;
    ToolWindow971: TToolWindow97;
    Tex1: TLabel;
    Nb1: TLabel;
    BTag: TToolbarButton97;
    Bdetag: TToolbarButton97;
    BValider: TToolbarButton97;
    BFerme: TToolbarButton97;
    BAide: TToolbarButton97;
    Dock972: TDock97;
    ToolWindow972: TToolWindow97;
    BChercher: TToolbarButton97;
    FFiltres: THValComboBox;
    BFiltre: TToolbarButton97;
    procedure FormShow(Sender: TObject);
    procedure BFerme1Click(Sender: TObject);
    procedure AxeChange(Sender: TObject);
    procedure TLChange(Sender: TObject);
    procedure CpteDeChange(Sender: TObject);
    procedure BChercher1Click(Sender: TObject);
    procedure FListeKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure Bdetag1Click(Sender: TObject);
    procedure BTag1Click(Sender: TObject);
    procedure FListeMouseDown(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
    procedure BValider1Click(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure BAide1Click(Sender: TObject);
    procedure BValiderClick(Sender: TObject);
    procedure BFermeClick(Sender: TObject);
    procedure BAideClick(Sender: TObject);
    procedure BdetagClick(Sender: TObject);
    procedure BTagClick(Sender: TObject);
    procedure BChercherClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
  private
    ObjFiltre  : TObjFiltre; {JP 21/01/05 : Gestion des filtres FQ 15255}
    FNomFiltre : string; {JP 24/01/05 : FQ 15255}
    UnFb : TFichierBase ;
    Pref : String ;
    Code,Lib,Champ : String ;
    TotalSelec : Integer ;
    RunMaj : Boolean ;
    WMinX,WMinY : Integer ;
    QTL : TQuery;  // VL 26/01/2005
    procedure WMGetMinMaxInfo(var MSG: Tmessage); message WM_GetMinMaxInfo;
    Function  ControleCpteOk : Boolean ;
    Function  FabriqueRequete : String ;
    Procedure FaitMajTableLibre(Sql : String) ;
    Procedure CompteElemSelectionner ;
    Procedure TagDetag(Avec : Boolean) ;
    Procedure GetCellCanvas(Acol,ARow : LongInt ; Canvas : TCanvas; AState: TGridDrawState) ;
    Procedure InitZone ;
    Function  UnAndGen : String ;
    Function  UnAndTiers : String ;
  public
    { Déclarations publiques }
    procedure ApresChargementFiltre;{JP 24/01/05 : FQ 15255}
  end;


implementation

{$R *.DFM}

Uses
  {$IFDEF MODENT1}
  CPProcMetier,
  {$ENDIF MODENT1}
  HStatus;

{***********A.G.L.***********************************************
Auteur  ...... : 
Créé le ...... : 09/09/2003
Modifié le ... : 09/09/2003
Description .. : 
Suite ........ : 09/09/2003, SBO : MODIF PACK AVANCE pour gestion
Suite ........ : mode inside
Mots clefs ... :
*****************************************************************}
Procedure ModifSerieTableLibre(UnFb : TFichierBase) ;
var FModiTali : TFModiTali ;
    PP : THPanel ;
begin
  FModiTali:=TFModiTali.Create(Application) ;
  FModiTali.UnFb:=UnFb ;

  PP:=FindInsidePanel ;
  if PP=Nil then
    begin
    Try
      FModiTali.ShowModal ;
      Finally
      FModiTali.Free ;
      End ;
    end
  else
    begin
    InitInside(FModiTali,PP) ;
    FModiTali.Show ;
    end ;

  SourisNormale ;
end ;

procedure TFModiTali.FormCreate(Sender: TObject);
var
  Composants : TControlFiltre;
begin
  {JP 21/01/05 : Gestion des filtres FQ 15255}
  Composants.PopupF   := POPF;
  Composants.Filtres  := FFILTRES;
  Composants.Filtre   := BFILTRE;
  Composants.PageCtrl := Pages;
  ObjFiltre := TObjFiltre.Create(Composants, '');
  ObjFiltre.ApresChangementFiltre := ApresChargementFiltre;
  WMinX:=Width;
  WMinY:=Height;
end;

procedure TFModiTali.WMGetMinMaxInfo(var MSG: Tmessage);
begin
with PMinMaxInfo(MSG.lparam)^.ptMinTrackSize do begin X := WMinX ; Y := WMinY ; end;
end;

procedure TFModiTali.BFerme1Click(Sender: TObject);
begin Close ; end;

Procedure TFModiTali.InitZone ;
BEGIN
If Axe.Values.Count>0 Then Axe.Value:=Axe.Values[0] ; NatGen.ItemIndex:=0 ; Ventil.ItemIndex:=0 ;
T_SECTEUR.ItemIndex:=0 ; T_NATUREAUXI.ItemIndex:=0 ; T_REGIMETVA.ItemIndex:=0 ;
END ;

procedure TFModiTali.FormShow(Sender: TObject);
begin
FListe.GetCellCanvas:=GetCellCanvas ;
//Pages.ActivePage:=Pages.Pages[0] ;
PCompGen.TabVisible:=(UnFb in [fbGene]) ;
PCompTiers.TabVisible:=(UnFb in [fbAux]) ;
Pages.ActivePage:=PStandards ;
Axe.Visible:=((UnFb in [fbSect]) or (UnFb in [fbBudSec1..fbBudSec5])) ;
TAxe.Visible:=((UnFb in [fbSect]) or (UnFb in [fbBudSec1..fbBudSec5])) ;
InitZone ;
CpteDeJ.Visible:=False ; RunMaj:=False ;
Case Unfb of
     fbSect :    BEGIN
                 Pref:='S' ; TCpteDe.Caption:=MsgBox.Mess[0] ;
                 HelpContext:=7182000 ;
                 FNomFiltre := 'MODITALISECT';
                 END ;
     fbGene :    BEGIN
                 Pref:='G' ; TCpteDe.Caption:=MsgBox.Mess[1] ;
                 CpteDe.ZoomTable:=tzGeneral ; CpteA.ZoomTable:=CpteDe.ZoomTable ;
                 HelpContext:=7116000 ;
                 FNomFiltre := 'MODITALIGENE';
                 END ;
     fbAux :     BEGIN
                 Pref:='T' ; TCpteDe.Caption:=MsgBox.Mess[2] ;
                 CpteDe.ZoomTable:=tzTiers ; CpteA.ZoomTable:=CpteDe.ZoomTable ;
                 HelpContext:=7149000 ;
                 FNomFiltre := 'MODITALIAUX';
                 END ;
     fbBudGen :  BEGIN
                 Pref:='B' ; TCpteDe.Caption:=MsgBox.Mess[3] ;
                 CpteDe.ZoomTable:=tzBudgen ; CpteA.ZoomTable:=CpteDe.ZoomTable ;
                 HelpContext:=15118000 ;
                 FNomFiltre := 'MODITALIBUDGEN';
                 END ;
     fbBudSec1 : BEGIN
                 Pref:='D' ; TCpteDe.Caption:=MsgBox.Mess[4] ;
                 CpteDe.ZoomTable:=tzBudSec1 ; CpteA.ZoomTable:=CpteDe.ZoomTable ;
                 HelpContext:=15138000 ;
                 FNomFiltre := 'MODITALIBUDSEC1';
                 END ;
     fbBudSec2 : BEGIN
                 Pref:='D' ; TCpteDe.Caption:=MsgBox.Mess[4] ;
                 CpteDe.ZoomTable:=tzBudSec2 ; CpteA.ZoomTable:=CpteDe.ZoomTable ;
                 HelpContext:=15138000 ;
                 FNomFiltre := 'MODITALIBUDSEC2';
                 END ;
     fbBudSec3 : BEGIN
                 Pref:='D' ; TCpteDe.Caption:=MsgBox.Mess[4] ;
                 CpteDe.ZoomTable:=tzBudSec3 ; CpteA.ZoomTable:=CpteDe.ZoomTable ;
                 HelpContext:=15138000 ;
                 FNomFiltre := 'MODITALIBUDSEC3';
                 END ;
     fbBudSec4 : BEGIN
                 Pref:='D' ; TCpteDe.Caption:=MsgBox.Mess[4] ;
                 CpteDe.ZoomTable:=tzBudSec4 ; CpteA.ZoomTable:=CpteDe.ZoomTable ;
                 HelpContext:=15138000 ;
                 FNomFiltre := 'MODITALIBUDSEC4';
                 END ;
     fbBudSec5 : BEGIN
                 Pref:='D' ; TCpteDe.Caption:=MsgBox.Mess[4] ;
                 CpteDe.ZoomTable:=tzBudSec5 ; CpteA.ZoomTable:=CpteDe.ZoomTable ;
                 HelpContext:=15138000 ;
                 FNomFiltre := 'MODITALIBUDSEC5';
                 END ;
     fbImmo :    BEGIN
                 Pref:='I' ; TCpteDe.Caption:=MsgBox.Mess[13] ;
                 CpteDe.ZoomTable:=tzImmo ; CpteA.ZoomTable:=CpteDe.ZoomTable ;
                 HelpContext:=0 ;
                 FNomFiltre := 'MODITALIIMMO';
                 END ;
   End ;
ChargeComboTableLibre(Pref,TL.Values,TL.Items) ;
if TL.Values.Count>0 then TL.Value:=TL.Values[0] ;
  {JP 21/01/05 : Gestion des filtres FQ 15255}
  ObjFiltre.FFI_TABLE := FNomFiltre;
  ObjFiltre.Charger;
end;

Procedure TFModiTali.GetCellCanvas(Acol,ARow : LongInt ; Canvas : TCanvas; AState: TGridDrawState) ;
BEGIN
if FListe.Cells[FListe.ColCount-1,ARow]='*' then FListe.Canvas.Font.Style:=FListe.Canvas.Font.Style+[fsItalic]
                                            else FListe.Canvas.Font.Style:=FListe.Canvas.Font.Style-[fsItalic] ;
END ;

procedure TFModiTali.AxeChange(Sender: TObject);
Var ii : TZoomTable ;
begin
  {JP 21/01/05 : Gestion des filtres FQ 15255}
  if not ObjFiltre.InChargement then begin
    if (UnFb in [fbGene,fbAux,fbBudgen,fbImmo]) or
       {JP 24/01/05 : Par précaution, je me suis retrouvé dans le cas}
       (Axe.Value = '') or (Length(Axe.Value) < 2) then Exit ;
    if UnFb in [fbBudSec1..fbBudSec5] then
       BEGIN
       ii:=AxeToTz(Axe.Value) ;
       Case ii of
          tzSection  : BEGIN CpteDe.ZoomTable:=tzBudSec1 ; CpteA.ZoomTable:=tzBudSec1 ; END ;
          tzSection2 : BEGIN CpteDe.ZoomTable:=tzBudSec2 ; CpteA.ZoomTable:=tzBudSec2 ; END ;
          tzSection3 : BEGIN CpteDe.ZoomTable:=tzBudSec3 ; CpteA.ZoomTable:=tzBudSec3 ; END ;
          tzSection4 : BEGIN CpteDe.ZoomTable:=tzBudSec4 ; CpteA.ZoomTable:=tzBudSec4 ; END ;
          tzSection5 : BEGIN CpteDe.ZoomTable:=tzBudSec5 ; CpteA.ZoomTable:=tzBudSec5 ; END ;
          END ;
       END else
       BEGIN
       CpteDe.ZoomTable:=AxeToTz(Axe.Value) ;
       CpteA.ZoomTable:=CpteDe.ZoomTable ;
       END ;
  end;
end;

procedure TFModiTali.TLChange(Sender: TObject);
begin
  {JP 21/01/05 : Gestion des filtres FQ 15255}
  if TL.Value <> '' then begin
    AncVal.ZoomTable:=NatureToTz(TL.Value) ;
    NewVal.ZoomTable:=AncVal.ZoomTable ;
  end;
end;

procedure TFModiTali.CpteDeChange(Sender: TObject);
Var AvecJoker : Boolean ;
begin
AvecJoker:=Joker(CpteDe, CpteA, CpteDeJ) ;
TCpteA.Visible:=Not AvecJoker ;
end;

procedure TFModiTali.BChercher1Click(Sender: TObject);
Var Sql : String ;
begin
RunMaj:=True ; Sql:=FabriqueRequete ; FaitMajTableLibre(Sql) ;
end;

Function TFModiTali.ControleCpteOk : Boolean ;
BEGIN
Result:=False ;
if NewVal.Text<>'' then
  if NewVal.ExisteH<=0 then BEGIN MsgBox.Execute(5,'','') ; NewVal.SetFocus ; Exit ; END ;
Result:=True ;
END ;

Function TFModiTali.FabriqueRequete : String ;
Var StD,StA : String ;
    Table,Pre,WhereAx,LeAnd,AndSup : String ;
    LeFb : TFichierBase ;
BEGIN
Result:='' ; LeFb:=UnFb ; WhereAx:='' ;
Case UnFb of
     fbSect :               BEGIN
                            Table:='SECTION'  ; Pre:='S_'  ; Code:=Pre+'SECTION'    ;
                            Case Axe.Value[2] of
                               '1':Lefb:=fbAxe1 ; '2':Lefb:=fbAxe2 ; '3':Lefb:=fbAxe3 ; '4':Lefb:=fbAxe4 ; '5':Lefb:=fbAxe5 ;
                               End ;
                            END ;
     fbGene :               BEGIN Table:='GENERAUX' ; Pre:='G_'  ; Code:=Pre+'GENERAL'    ; END ;
     fbImmo :               BEGIN Table:='IMMO'     ; Pre:='I_'  ; Code:=Pre+'IMMO'    ; END ;
     fbAux :                BEGIN Table:='TIERS'    ; Pre:='T_'  ; Code:=Pre+'AUXILIAIRE' ; END ;
     fbBudGen :             BEGIN Table:='BUDGENE'  ; Pre:='BG_' ; Code:=Pre+'BUDGENE'    ; END ;
     fbBudSec1..fbBudSec5 : BEGIN Table:='BUDSECT'  ; Pre:='BS_' ; Code:=Pre+'BUDSECT'    ; END ;
   End ;
if (Axe.Visible) And ((Table='BUDSECT') or (Table='SECTION'))then
   WhereAx:=' And '+Pre+'AXE="'+Axe.Value+'" ' ;
Champ:=Pre+'TABLE'+IntToStr(StrToInt(Copy(TL.Value,2,2))) ; Lib:=Pre+'LIBELLE' ;
LeAnd:='' ; StD:='' ;StA:='' ;
if Not CpteA.Visible then
   BEGIN
   LeAnd:=' And '+Code+' Like "'+TraduitJoker(CpteDeJ.Text)+'"' ;
   END else
   BEGIN
   StD:=CpteDe.Text ; StA:=CpteA.Text ;
   if (StD='') And (StA='') then LeAnd:='' ;
   if (StD<>'') And (StA='') then LeAnd:=' And '+Code+' >= "'+BourreLaDonc(StD,LeFb)+'"' ;
   if (StD='') And (StA<>'') then LeAnd:=' And '+Code+' <= "'+BourreLaDonc(StA,LeFb)+'"' ;
   if (StD<>'') And (StA<>'') then LeAnd:=' And '+Code+' >="'+BourreLaDonc(StD,LeFb)+'" And '+Code+ '<="'+BourreLaDonc(StA,LeFb)+'"' ;
   END ;
If Unfb=fbGene Then AndSup:=UnAndGen Else If Unfb=fbAux Then AndSup:=UnAndTiers Else AndSup:='' ;
{ FQ 21950 BVE 08.10.07 }
if UnFb = fbSect then
   Result:='Select '+Pre+'AXE,'+Code+','+Lib+','+Champ+' From '+Table+' Where '+Champ+'="'+AncVal.Text+'" '+LeAnd+WhereAx+AndSup+' Order by '+Code
else
   Result:='Select '+Code+','+Lib+','+Champ+' From '+Table+' Where '+Champ+'="'+AncVal.Text+'" '+LeAnd+WhereAx+AndSup+' Order by '+Code ;
{END FQ 21950 }
END ;

Procedure TFModiTali.FaitMajTableLibre(Sql : String) ;
BEGIN
QTL := OpenSQL(Sql, False); // // VL 26/01/2005

FListe.VidePile(False) ; InitMove(RecordsCount(QTL),'') ;
While Not QTL.Eof do
  BEGIN
  MoveCur(False) ;
  FListe.Cells[0,FListe.RowCount-1]:=QTL.FindField(Code).AsString ;
  FListe.Cells[1,FListe.RowCount-1]:=QTL.FindField(Lib).AsString ;
  FListe.Cells[2,FListe.RowCount-1]:=QTL.FindField(Champ).AsString ;
  FListe.Cells[FListe.ColCount-1,FListe.RowCount-1]:='*' ;
  FListe.RowCount:=FListe.RowCount+1 ;
  QTL.Next ;
  END ;
if FListe.RowCount>2 then FListe.RowCount:=FListe.RowCount-1 ;
FListe.Invalidate ; CompteElemSelectionner ; QTl.Close ; FiniMove ;
END ;

procedure TFModiTali.FListeKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
if ((Shift=[]) And (Key=VK_SPACE)) or ((ssShift in Shift) And (Key=VK_DOWN))then
   FlisteMouseDown(Nil,mbLeft,[ssCtrl],0,0) ;
end;

Procedure TFModiTali.CompteElemSelectionner ;
Var i : Integer ;
BEGIN
TotalSelec:=0 ;
for i:=1 to FListe.RowCount-1 do
    if FListe.Cells[FListe.ColCount-1,i]='*' then Inc(TotalSelec) ;
Nb1.Caption:=IntToStr(TotalSelec) ;
if TotalSelec>1 then Tex1.Caption:=MsgBox.Mess[7] else Tex1.Caption:=MsgBox.Mess[6] ;
END ;

procedure TFModiTali.Bdetag1Click(Sender: TObject);
begin TagDetag(False) ; CompteElemSelectionner ; end;

procedure TFModiTali.BTag1Click(Sender: TObject);
begin TagDetag(True) ; CompteElemSelectionner ; end;

Procedure TFModiTali.TagDetag(Avec : Boolean) ;
Var  i : Integer ;
begin
if Fliste.Cells[0,1]='' then Exit ;
for i:=1 to FListe.RowCount-1 do
    if Avec then FListe.Cells[FListe.ColCount-1,i]:='*'
            else FListe.Cells[FListe.ColCount-1,i]:='' ;
FListe.Invalidate ; Bdetag.Visible:=Avec ; BTag.Visible:=Not Avec ;
end;

procedure TFModiTali.FListeMouseDown(Sender: TObject; Button: TMouseButton;
  Shift: TShiftState; X, Y: Integer);
begin
if Not(ssCtrl in Shift) then Exit ;
if Button<>mbLeft then Exit ;
if Fliste.Cells[FListe.ColCount-1,FListe.Row]='*'
   then Fliste.Cells[FListe.ColCount-1,FListe.Row]:=''
   else Fliste.Cells[FListe.ColCount-1,FListe.Row]:='*' ;
FListe.Invalidate ; CompteElemSelectionner ;
end;

procedure TFModiTali.BValider1Click(Sender: TObject);
Var i : Integer ;
    Probleme : Boolean ;
begin
if Not RunMaj then Exit ;
if FListe.Cells[0,1]='' then BEGIN MsgBox.Execute(8,'','') ; Exit ; END ;
if TotalSelec=0 then BEGIN MsgBox.Execute(9,'','') ; Exit ; END ;
if Not ControleCpteOk then Exit ;
if NewVal.Text=AncVal.Text then BEGIN MsgBox.Execute(11,'','') ; Exit ; END ;
if MsgBox.Execute(10,'','')<>mrYes then Exit ;
QTL.Open ; QTL.First ; Probleme:=False ;
InitMove(FListe.RowCount-1,'') ;
for i:=1 to FListe.RowCount-1 do
   BEGIN
   MoveCur(False) ;
   if FListe.Cells[FListe.ColCount-1,i]='*' then
      BEGIN
      if QTL.Locate(Code,FListe.Cells[0,i],[]) then
         BEGIN
         QTL.Edit ;
         QTL.FindField(Champ).AsString:=NewVal.Text ;
         QTL.Post ;
         END else Probleme:=True ;
      END ;
   END ;
FiniMove ; BChercherClick(Nil) ;
if Probleme then MsgBox.Execute(12,'','') ;
end;

Function TFModiTali.UnAndGen : String ;
Var AndSup : string ;
BEGIN
Result:='' ; AndSup:='' ;
if UnFb<>fbGene then Exit ;
if NatGen.Value<>'' then AndSup:=AndSup+' And G_NATUREGENE="'+NatGen.Value+'" ' ;
if Ventil.Value<>'' then
   BEGIN
   Case Ventil.Value[2] of
        '1' : AndSup:=AndSup+' And G_VENTILABLE1="'+Ventil.Value+'" ' ;
        '2' : AndSup:=AndSup+' And G_VENTILABLE2="'+Ventil.Value+'" ' ;
        '3' : AndSup:=AndSup+' And G_VENTILABLE3="'+Ventil.Value+'" ' ;
        '4' : AndSup:=AndSup+' And G_VENTILABLE4="'+Ventil.Value+'" ' ;
        '5' : AndSup:=AndSup+' And G_VENTILABLE5="'+Ventil.Value+'" ' ;
      End ;
   END ;
Case G_COLLECTIF.State of
     cbChecked : AndSup:=AndSup+' And G_COLLECTIF="X" ' ;
     cbUnchecked : AndSup:=AndSup+' And G_COLLECTIF="-" ' ;
     End ;
Case G_FERME.State of
     cbChecked : AndSup:=AndSup+' And G_FERME="X" ' ;
     cbUnchecked : AndSup:=AndSup+' And G_FERME="-" ' ;
     End ;
Case G_LETTRABLE.State of
     cbChecked : AndSup:=AndSup+' And G_LETTRABLE="X" ' ;
     cbUnchecked : AndSup:=AndSup+' And G_LETTRABLE="-" ' ;
     End ;
Case G_POINTABLE.State of
     cbChecked : AndSup:=AndSup+' And G_POINTABLE="X" ' ;
     cbUnchecked : AndSup:=AndSup+' And G_POINTABLE="-" ' ;
     End ;
Result:=AndSup ;
END ;

Function TFModiTali.UnAndTiers : String ;
Var AndSup : String ;
BEGIN
Result:='' ; AndSup:='' ;
if UnFb<>fbAux then Exit ;
if T_SECTEUR.Value<>'' then AndSup:=AndSup+' And T_SECTEUR="'+T_SECTEUR.Value+'" ' ;
if T_NATUREAUXI.Value<>'' then AndSup:=AndSup+' And T_NATUREAUXI="'+T_NATUREAUXI.Value+'" ' ;
if T_REGIMETVA.Value<>'' then AndSup:=AndSup+' And T_REGIMETVA="'+T_REGIMETVA.Value+'" ' ;
Case T_LETTRABLE.State of
     cbChecked : AndSup:=AndSup+' And T_LETTRABLE="X" ' ;
     cbUnchecked : AndSup:=AndSup+' And T_LETTRABLE="-" ' ;
     End ;
Case T_MULTIDEVISE.State of
     cbChecked : AndSup:=AndSup+' And T_MULTIDEVISE="X" ' ;
     cbUnchecked : AndSup:=AndSup+' And T_MULTIDEVISE="-" ' ;
     End ;
if T_COLLECTIF.Text<>'' then AndSup:=AndSup+' And T_COLLECTIF Like "'+T_COLLECTIF.Text+'%" ' ;
Result:=AndSup ;
END ;

procedure TFModiTali.BAide1Click(Sender: TObject);
begin
CallHelpTopic(Self) ;
end;

procedure TFModiTali.BValiderClick(Sender: TObject);
Var i : Integer ;
    Probleme : Boolean ;
begin
if Not RunMaj then Exit ;
if FListe.Cells[0,1]='' then BEGIN MsgBox.Execute(8,'','') ; Exit ; END ;
if TotalSelec=0 then BEGIN MsgBox.Execute(9,'','') ; Exit ; END ;
if Not ControleCpteOk then Exit ;
if NewVal.Text=AncVal.Text then BEGIN MsgBox.Execute(11,'','') ; Exit ; END ;
if MsgBox.Execute(10,'','')<>mrYes then Exit ;
QTL.Open ; QTL.First ; Probleme:=False ;
InitMove(FListe.RowCount-1,'') ;
for i:=1 to FListe.RowCount-1 do
   BEGIN
   MoveCur(False) ;
   if FListe.Cells[FListe.ColCount-1,i]='*' then
      BEGIN
      if QTL.Locate(Code,FListe.Cells[0,i],[]) then
         BEGIN
         QTL.Edit ;                                 
         QTL.FindField(Champ).AsString:=NewVal.Text ;
         QTL.Post ;
         END else Probleme:=True ;
      END ;
   END ;
FiniMove ; BChercherClick(Nil) ;
if Probleme then MsgBox.Execute(12,'','') ;
end;

procedure TFModiTali.BFermeClick(Sender: TObject);
begin
  Close;
end;

procedure TFModiTali.BAideClick(Sender: TObject);
begin
  CallHelpTopic(Self);
end;

procedure TFModiTali.BdetagClick(Sender: TObject);
begin
  TagDetag(False);
  CompteElemSelectionner;
end;

procedure TFModiTali.BTagClick(Sender: TObject);
begin
  TagDetag(True);
  CompteElemSelectionner; 
end;

procedure TFModiTali.BChercherClick(Sender: TObject);
Var Sql : String ;
begin
RunMaj:=True ; Sql:=FabriqueRequete ; FaitMajTableLibre(Sql) ;
end;

{---------------------------------------------------------------------------------------}
procedure TFModiTali.FormClose(Sender: TObject; var Action: TCloseAction);
{---------------------------------------------------------------------------------------}
begin
  {JP 21/01/05 : Gestion des filtres FQ 15255}
  FreeAndNil(ObjFiltre);
end;

{---------------------------------------------------------------------------------------}
procedure TFModiTali.ApresChargementFiltre;
{---------------------------------------------------------------------------------------}
begin
  {JP 24/01/05 : On lance la recherche après chargement du filtre}
  if FFiltres.Text <> '' then BChercherClick(BChercher);
end;

end.
