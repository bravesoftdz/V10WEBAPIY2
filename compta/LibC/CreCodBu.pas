{***********UNITE*************************************************
Auteur  ...... : Vincent Laroche
Créé le ...... : 27/01/2005
Modifié le ... :   /  /    
Description .. : Passage en eAGL
Mots clefs ... : 
*****************************************************************}
unit CreCodBu;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs, Ent1,
  Hent1, hmsgbox, StdCtrls, Buttons, ExtCtrls, Hctrls, Grids, ComCtrls,
  HSysMenu, HCompte, HStatus, MajCodBu, HTB97, ed_tools,
{$IFDEF EAGLCLIENT}

{$ELSE}
  DB,
  {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ENDIF}
{$IFDEF MODENT1}
  CPTypeCons,
{$ENDIF MODENT1}
  UTob, HPanel, UIUtil // MODIF PACK AVANCE pour gestion mode inside
  ;

Type InfoCpte = Class
     Lib : String ;
     Fait : Boolean ;
     End ;

Procedure CreationDesCodesBudget(Lefb : TFichierBase) ;

type
  TFCreCodBu = class(TForm)
    HM: THMsgBox;
    Pages: TPageControl;
    Pparam: TTabSheet;
    FListe: THGrid;
    HPB: TPanel;
    Tex1: TLabel;
    Nb1: TLabel;
    Panel1: TPanel;
    Bdetag: THBitBtn;
    BTag: THBitBtn;
    BAide: THBitBtn;
    BFerme: THBitBtn;
    BValider: THBitBtn;
    TSens: THLabel;
    Sens: THValComboBox;
    TSigne: THLabel;
    Signe: THValComboBox;
    EChoix: TEdit;
    LTri: TLabel;
    Bevel1: TBevel;
    Buse: TToolbarButton97;
    HMTrad: THSystemMenu;
    CbTabli: THValComboBox;
    ListCpte: THValComboBox;
    BCherche: THBitBtn;
    TAxe: THLabel;
    Axe: THValComboBox;
    CbRac: TRadioGroup;
    procedure BFermeClick(Sender: TObject);
    procedure BuseClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure BChercheClick(Sender: TObject);
    procedure BTagClick(Sender: TObject);
    procedure BdetagClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FListeMouseDown(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure BValiderClick(Sender: TObject);
    procedure AxeChange(Sender: TObject);
    procedure FListeKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure BAideClick(Sender: TObject);
  private
    Lefb : TFichierBase ;
    Tri : String ;
    WMinX,WMinY    : Integer ;
    Pref,Table,TabLib,Cod,Lib,WhereAx,Ax,Cle : String ;
    procedure WMGetMinMaxInfo(var MSG: Tmessage); message WM_GetMinMaxInfo;
    Procedure DefautUseTabLib ;
    Function  VerifieEChoix : Boolean ;
    Procedure TagDetag(Avec : Boolean) ;
    Procedure CompteElemSelectionner ;
    Function  ListeVide : Boolean ;
    Procedure DetruitCombos ;
    Procedure ChercheEnreg ;
    Function  FaitCodeLibBud(Var L : String; Q : TQuery) : String ;
    Procedure RempliLeGrille ;
    Function  QuelVraifb : TFichierBase ;
    Procedure GetCellCanvas(Acol,ARow : LongInt ; Canvas : TCanvas; AState: TGridDrawState) ;
    Procedure InverseSelection ;
    Procedure ExtraitRacine(Var Rubri : String) ;
    Function  FaitLaRubrique(St,St1 : String) : String ;
    Function  FaitUnObjet(UnRow : HTStrings ; QuelTab : String) : TCodBud ;
  public
    { Déclarations publiques }
  end;


implementation

{$R *.DFM}

Uses
  {$IFDEF MODENT1}
  CPProcMetier,
  {$ENDIF MODENT1}
  TriTabLi;  // ChoixTriTableLibre

Const Code = 0 ; Libel = 1 ; Rubri = 2 ; Selection = 3 ;

{***********A.G.L.***********************************************
Auteur  ...... : 
Créé le ...... : 09/09/2003
Modifié le ... :   /  /
Description .. : 
Suite ........ : 09/09/2003, SBO : MODIF PACK AVANCE pour gestion 
Suite ........ : mode inside
Mots clefs ... : 
*****************************************************************}
Procedure CreationDesCodesBudget(Lefb : TFichierBase) ;
var FCreCodBu : TFCreCodBu ;
    PP : THPanel ;
BEGIN
  FCreCodBu:=TFCreCodBu.Create(Application) ;
  Case Lefb of
    fbBudGen             : FCreCodBu.Caption:=FCreCodBu.HM.Mess[0] ;
    fbBudSec1..fbBudSec5 : FCreCodBu.Caption:=FCreCodBu.HM.Mess[1] ;
    End ;
  FCreCodBu.Lefb:=Lefb ;

  PP:=FindInsidePanel ;
  if PP=Nil then
    begin
    Try
      FCreCodBu.ShowModal ;
      Finally
      FCreCodBu.Free ;
      End ;
    end
  else
    begin
    InitInside(FCreCodBu,PP) ;
    FCreCodBu.Show ;
    end ;

  SourisNormale ;
END ;

procedure TFCreCodBu.FormCreate(Sender: TObject);
begin
PopUpMenu:=ADDMenuPop(PopUpMenu,'','') ;
WMinX:=Width ; WMinY:=Height ; Tri:='' ;
end;

procedure TFCreCodBu.FormClose(Sender: TObject; var Action: TCloseAction);
begin DetruitCombos ; FListe.VidePile(True) ; end;

procedure TFCreCodBu.WMGetMinMaxInfo(var MSG: Tmessage);
BEGIN
with PMinMaxInfo(MSG.lparam)^.ptMinTrackSize do begin X := WMinX ; Y := WMinY ; end;
END ;

procedure TFCreCodBu.BFermeClick(Sender: TObject);
begin
  Close ;
  if IsInside(Self) then
    CloseInsidePanel(Self) ;
end;

procedure TFCreCodBu.BuseClick(Sender: TObject);
begin
Tri:=EChoix.Text ;
ChoixTriTableLibre(QuelVraifb,Tri,False,'') ;
EChoix.Text:=Tri ;
end;

Procedure TFCreCodBu.DetruitCombos ;
Var i : Integer ;
BEGIN
(*St:=EChoix.Text ;
While St<>'' do
   BEGIN
   St1:=ReadTokenSt(St) ;
   if St1<>'' then THValComboBox(FindComponent(St1)).Free ;
   END ;*)
for i:=0 to ListCpte.Items.Count-1 do
   if ListCpte.Items.Objects[i]<>Nil then InfoCpte(ListCpte.Items.Objects[i]).Free ;
END ;

Procedure TFCreCodBu.TagDetag(Avec : Boolean) ;
Var  i : Integer ;
begin
for i:=1 to FListe.RowCount-1 do
    if Avec then FListe.Cells[FListe.ColCount-1,i]:='*'
            else FListe.Cells[FListe.ColCount-1,i]:='' ;
FListe.Invalidate ; FListe.SetFocus ;
Bdetag.Visible:=Avec ; BTag.Visible:=Not Avec ; CompteElemSelectionner ;
end;

procedure TFCreCodBu.BdetagClick(Sender: TObject);
begin TagDetag(False) ; end;

procedure TFCreCodBu.BTagClick(Sender: TObject);
begin TagDetag(True) ; end;

Procedure TFCreCodBu.CompteElemSelectionner ;
Var i,j : Integer ;
BEGIN
j:=0 ;
if Not ListeVide then
   BEGIN
   for i:=1 to FListe.RowCount-1 do
       if FListe.Cells[FListe.ColCount-1,i]='*' then Inc(j) ;
   END ;
Case j of
     0,1: BEGIN
          Nb1.Caption:=IntTostr(j) ; Tex1.Caption:=HM.Mess[4] ;
          END ;
     else BEGIN
          Nb1.Caption:=IntTostr(j) ; Tex1.Caption:=HM.Mess[5] ;
          END ;
   End ;
END ;

Function TFCreCodBu.ListeVide : Boolean ;
BEGIN Result:=FListe.Cells[0,1]='' ; END ;

procedure TFCreCodBu.FListeMouseDown(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin if (ssCtrl in Shift) And (Button=mbLeft)then InverseSelection ; end;

procedure TFCreCodBu.AxeChange(Sender: TObject);
begin Lefb:=AxeToFbBud(Axe.Value) ; end;

procedure TFCreCodBu.FormShow(Sender: TObject);
begin
FListe.GetCellCanvas:=GetCellCanvas ;
Axe.ItemIndex:=0 ;
Case Lefb of
     fbBudGen             : BEGIN Axe.Enabled:=False ; TAxe.Enabled:=False ; Pref:='G' ; Table:='GENERAUX' ; Cle:='_GENERAL' ; HelpContext:=7577320 ; END ;
     fbBudSec1..fbBudSec5 : BEGIN Axe.Enabled:=True  ; TAxe.Enabled:=True ; Pref:='S' ; Table:='SECTION'  ; Cle:='_SECTION' ; HelpContext:=7577410 ; END ;
   End ;
TabLib:=Pref+'_TABLE' ; Cod:=Pref+Cle ;  Lib:=Pref+'_LIBELLE' ;
DefautUseTabLib ; Signe.Value:='POS' ; Sens.Value:='M' ;
end;

Procedure TFCreCodBu.DefautUseTabLib ;
Var QLoc : TQuery ;
BEGIN
CbTabli.Values.Clear ; CbTabLi.Items.Clear ;
QLoc:=OpenSql('Select NT_TYPECPTE,NT_LIBELLE From NATCPTE Where NT_TYPECPTE Like"'+Pref+'%" Order By NT_TYPECPTE',True) ;
While Not QLoc.Eof do
  BEGIN
  CbTabli.Values.Add(QLoc.Fields[0].AsString) ;
  CbTabli.Items.Add(QLoc.Fields[1].AsString) ;
  QLoc.Next ;
  END ;
Ferme(QLoc) ;
END ;

Function TFCreCodBu.VerifieEChoix : Boolean ;
Var St,St1 : String ;
BEGIN
Result:=False ; St:=EChoix.Text ;
if EChoix.Text='' then BEGIN HM.Execute(3,Caption,'') ; EChoix.SetFocus ; Exit ; END ;
While St<>'' do
  BEGIN
  St1:=ReadTokenSt(St) ;
  if (CbTabli.Values.IndexOf(St1)<=-1) Or (Pos(St1,St)>0)then
      BEGIN HM.Execute(2,Caption,'') ; EChoix.SetFocus ; Exit ; END ;
  END ;
Result:=True ;
END ;

procedure TFCreCodBu.BChercheClick(Sender: TObject);
begin
DetruitCombos ; FListe.VidePile(True) ;
if VerifieEChoix then
   BEGIN
   Ax:=Axe.Value ; ChercheEnreg ; RempliLeGrille ;
   END ;
end;

Procedure TFCreCodBu.ChercheEnreg ;
Var Sql : String ;
    St,StTabli : String ;
    CodB,L : String ;
    X : InfoCpte ;
    Q : TQuery;
BEGIN
St:=EChoix.Text ; StTabli:='' ; ListCpte.Values.Clear ; ListCpte.Items.Clear ;
While St<>'' do StTabli:=StTabli+TabLib+Copy(ReadTokenSt(St),3,1)+'<>"" And ' ;
Delete(StTabli,Length(StTabli)-4,4) ;
Case Lefb of
     fbBudGen             : WhereAx:='' ;
     fbBudSec1..fbBudSec5 : WhereAx:=' And S_AXE="'+Ax+'"' ;
   End ;
Sql:='Select * From '+Table+' Where '+StTabli+WhereAx ;
Q := OpenSQL(Sql, True);
While Not Q.Eof do
   BEGIN
   L:='' ; CodB:=FaitCodeLibBud(L, Q);
   X:=InfoCpte.Create ;
   X.Lib:=L ; X.Fait:=False ;
   ListCpte.Values.Add(CodB) ; ListCpte.Items.AddObject(Q.FindField(Cod).AsString,X) ;
   Q.Next ;
   END ;
Ferme(Q);
END ;

Function TFCreCodBu.FaitCodeLibBud(Var L : String; Q : TQuery) : String ;
Var
  St,St1,CPte, Nat : String ;
  QLib : TQuery;
BEGIN
St:=EChoix.Text ; Cpte:='' ;
While St<>'' do
  BEGIN
  St1:=ReadTokenSt(St) ;
  Nat := Q.FindField(TabLib+Copy(St1,3,1)).AsString;
  QLib := OpenSQL('SELECT NT_LIBELLE FROM NATCPTE WHERE NT_TYPECPTE="'+St1+'" AND NT_NATURE="'+Nat+'"', True);
  L := L+ QLib.FindField('NT_LIBELLE').AsString;
  Cpte := Cpte + Nat;
  Ferme(QLib);
  END ;
Result:=Cpte ;
END ;

Function TFCreCodBu.QuelVraifb : TFichierBase ;
BEGIN
Result:=fbGene ;
Case Lefb of
     fbBudGen  : Result:=fbGene ;
     fbBudSec1 : Result:=fbAxe1 ;
     fbBudSec2 : Result:=fbAxe2 ;
     fbBudSec3 : Result:=fbAxe3 ;
     fbBudSec4 : Result:=fbAxe4 ;
     fbBudSec5 : Result:=fbAxe5 ;
   End ;
END ;

Procedure TFCreCodBu.RempliLeGrille ;
Var i,j : Integer ;
    Rub : String ;
BEGIN
InitMove(ListCpte.Values.Count,'') ;
for i:=0 to ListCpte.Values.Count-1 do
  BEGIN
  Rub:='' ; MoveCur(False) ;
  if InfoCpte(ListCpte.Items.Objects[i]).Fait then Continue else
     BEGIN
     InfoCpte(ListCpte.Items.Objects[i]).Fait:=True ;
     for j:=i to ListCpte.Items.Count-1 do
        BEGIN
        if (ListCpte.Values[j]=ListCpte.Values[i]) And (Not InfoCpte(ListCpte.Items.Objects[j]).Fait) then
            BEGIN
            if CbRac.ItemIndex=0 then Rub:=Rub+BourreLess(ListCpte.Items[j],QuelVraifb)+';'
                                 else Rub:=Rub+ListCpte.Items[j]+';' ;
            InfoCpte(ListCpte.Items.Objects[j]).Fait:=True ;
            END ;
        END ;
     if CbRac.ItemIndex=0 then Rub:=BourreLess(ListCpte.Items[i],QuelVraifb)+';'+Rub+';'
                          else Rub:=ListCpte.Items[i]+';'+Rub+';' ;
     if CbRac.ItemIndex=0 then ExtraitRacine(Rub) ;
     FListe.Cells[Code,FListe.RowCount-1]:=ListCpte.Values[i] ;
     FListe.Cells[Libel,FListe.RowCount-1]:=InfoCpte(ListCpte.Items.Objects[i]).Lib ;
     FListe.Cells[Rubri,FListe.RowCount-1]:=Rub ;
     FListe.Cells[Selection,FListe.RowCount-1]:='*' ;
     FListe.RowCount:=FListe.RowCount+1 ;
     END ;
  END ;
if (FListe.Cells[Code,FListe.RowCount-1]='') And (FListe.RowCount>2) then FListe.RowCount:=FListe.RowCount-1 ;
FListe.Invalidate ; CompteElemSelectionner ; FiniMove ; FListe.SetFocus ;
END ;

procedure TFCreCodBu.BValiderClick(Sender: TObject);
Var QuelTab,QuelChamp,Prefi,CompteRub,QuelCodRub,St : String ;
    i : Integer ;
    X : TCodBud ;
    Li : TList ;
    DoublonCod : Boolean ;
    T : Tob;
begin
if ListeVide then Exit ;
Case Lefb of
     fbBudGen : BEGIN
       QuelTab:='BUDGENE' ; QuelChamp:='BG_BUDGENE' ; Prefi:='BG_' ; CompteRub:='BG_COMPTERUB' ;
     END ;
     fbBudSec1..fbBudSec5 : BEGIN
       QuelTab:='BUDSECT' ; QuelChamp:='BS_BUDSECT' ; Prefi:='BS_' ; CompteRub:='BS_SECTIONRUB' ;
     END ;
   End ;
QuelCodRub:=Prefi+'RUB' ;

InitMove(FListe.RowCount-1,'') ; Li:=TList.Create ; DoublonCod:=False ;
T := Tob.Create(QuelTab, nil, -1);
for i:=1 to FListe.RowCount-1 do
   BEGIN
   MoveCur(False) ;
   if FListe.Cells[Selection,i]='*' then
      BEGIN
      if(Presence(QuelTab,QuelChamp,Copy(FListe.Cells[Code,i],1,17))) or
        (Presence(QuelTab,QuelCodRub,Copy(FListe.Cells[Code,i],1,5))) then
         BEGIN
         DoublonCod:=True ; X:=FaitUnObjet(FListe.Rows[i],QuelTab) ;
         Li.Add(X) ; Continue ;
         END else
         BEGIN
         T.InitValeurs(False);
         T.SetString(QuelChamp, Copy(FListe.Cells[Code,i],1,17));
         T.SetString(Prefi+'LIBELLE', Copy(FListe.Cells[Libel,i],1,35));
         T.SetString(Prefi+'ABREGE', Copy(FListe.Cells[Libel,i],1,17));
         St:=FListe.Cells[Rubri,i] ;
         if Length(St)>250 then St:=FaitLaRubrique(St,FListe.Cells[Code,i]) ;
         T.SetString(CompteRub, St);
         T.SetString(Prefi+'SIGNE', Signe.Value);
         T.SetString(Prefi+'SENS', Sens.Value);
         T.SetString(QuelCodRub, Copy(FListe.Cells[Code,i],1,5));
         if QuelTab='BUDSECT' then T.SetString(Prefi+'AXE', Ax);
         T.InsertOrUpdateDB;
         END ;
      END ;
   END ;
FiniMove ;
if DoublonCod then if HM.Execute(7,Caption,'')=mrYes then MajCode(Li,QuelTab) ;
VideListe(Li) ; Li.Free ;
T.Free;
end;

Function TFCreCodBu.FaitUnObjet(UnRow : HTStrings ; QuelTab : String) : TCodBud ;
Var X : TCodBud ;
    i : Integer ;
    St : String ;
BEGIN
X:=TCodBud.Create ;
for i:=0 to UnRow.Count-1 do
   BEGIN
   Case i of
      0 : BEGIN X.UnCod:=UnRow.Strings[i] ; X.UnRub:=Copy(UnRow.Strings[i],1,5) ; END ;
      1 : BEGIN X.UnLib:=Copy(UnRow.Strings[i],1,35) ; X.UnAbr:=Copy(UnRow.Strings[i],1,17) ; END ;
      2 : BEGIN
          St:=UnRow.Strings[i] ;
          if Length(St)>250 then St:=FaitLaRubrique(St,UnRow.Strings[0]) ;
          X.UnCpR:=St ;
          END ;

    End ;
   END ;
X.UnSig:=Signe.Value ; X.UnSen:=Sens.Value ;
if QuelTab='BUDSECT' then X.UnAxe:=Axe.Value else X.UnAxe:='' ;
Result:=X ;
END ;

Function TFCreCodBu.FaitLaRubrique(St,St1 : String) : String ;
Var j,k : Integer ;
BEGIN
HM.Execute(6,Caption,' : '+St1) ;
St:=Copy(St,1,248) ; k:=0 ;
for j:=Length(St) downto 1 do if St[j]=';' then BEGIN k:=j ; Break ; END ;
St:=Copy(St,1,k) ;
if St[Length(St)]<>';' then St:=St+';;' else St:=St+';' ;
Result:=St ;
END ;

Procedure TFCreCodBu.ExtraitRacine(Var Rubri : String) ;
Var St1,St2,St3,StRub,StTemp : String ;
    Trouver : Boolean ;
BEGIN
StTemp:=Rubri ; StRub:='' ;
While StTemp<>'' do
  BEGIN
  St1:=ReadTokenSt(StTemp) ; Trouver:=False ;
  if (Pos(St1,StRub)>0) or (St1='') then Continue ;
  St2:=StTemp ;
  While St2<>'' do
    BEGIN
    St3:=ReadTokenSt(St2) ;
    if St3='' then Continue ;
    if Length(St3) > Length(St1) then
       BEGIN
       if Pos(St1,St3)=1 then Delete(StTemp,Pos(St3,StTemp),Length(St3)) ;
       END else
       BEGIN
       if Pos(St3,St1)=1 then if Pos(St3,StRub)<=0 then BEGIN StRub:=StRub+St3+';' ; Trouver:=True ; END ;
       END ;
    END ;
  if Not Trouver then if Pos(St1,StRub)<=0 then StRub:=StRub+St1+';' ;
  END ;
Rubri:=StRub+';' ;
END ;

procedure TFCreCodBu.FListeKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
if ((ssShift in Shift) And (Key=VK_DOWN)) or ((ssShift in Shift) And (Key=VK_UP)) then InverseSelection else
   if (Shift=[]) And (Key=VK_SPACE) then
      BEGIN
      InverseSelection ;
      if ((FListe.Row<FListe.RowCount-1) and (Key<>VK_SPACE)) then FListe.Row:=FListe.Row+1 ;
      END ;
end;

Procedure TFCreCodBu.InverseSelection ;
BEGIN
if ListeVide then Exit ;
if FListe.Cells[FListe.ColCount-1,FListe.Row]='*' then FListe.Cells[FListe.ColCount-1,FListe.Row]:=''
                                                  else FListe.Cells[FListe.ColCount-1,FListe.Row]:='*' ;
CompteElemSelectionner ; FListe.Invalidate ;
END ;

Procedure TFCreCodBu.GetCellCanvas(Acol,ARow : LongInt ; Canvas : TCanvas; AState: TGridDrawState) ;
BEGIN
if FListe.Cells[FListe.ColCount-1,ARow]='*' then FListe.Canvas.Font.Style:=FListe.Canvas.Font.Style+[fsItalic]
                                            else FListe.Canvas.Font.Style:=FListe.Canvas.Font.Style-[fsItalic] ;
END ;

procedure TFCreCodBu.BAideClick(Sender: TObject);
begin
CallHelpTopic(Self) ;
end;

end.
