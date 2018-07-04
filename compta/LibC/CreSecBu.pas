{***********UNITE*************************************************
Auteur  ...... : Vincent Laroche
Créé le ...... : 27/01/2005
Modifié le ... :   /  /    
Description .. : Passage en eAGL
Mots clefs ... : 
*****************************************************************}
unit CreSecBu;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs, Ent1,
  HEnt1, hmsgbox, Grids, Hctrls, StdCtrls, Buttons, ExtCtrls,
{$IFDEF EAGLCLIENT}

{$ELSE}
  DB,
  {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF} 
{$ENDIF}
  ComCtrls, Menus, Filtre, HStatus, HSysMenu, MajCodBu, ed_tools, HPanel, UiUtil, UTob,
  HTB97;

Procedure GenerationSectionBudgetaire ;

Type TSectInfo = Class
     Cod : String ;
     Lib : String ;
     Abr : String ;
     Lax : String ;
     Conf : String ;
     CodRub : String ;
     End ;

type
  TFCreSecBu = class(TForm)
    Pages: TPageControl;
    Param: TTabSheet;
    Sens: THValComboBox;
    TSigne: THLabel;
    Signe: THValComboBox;
    TSens: THLabel;
    HPB: TToolWindow97;
    FListe: THGrid;
    HM: THMsgBox;
    Bevel1: TBevel;
    Axe: THValComboBox;
    TAxe: THLabel;
    HMTrad: THSystemMenu;
    BCherche: TToolbarButton97;
    Nb1: TLabel;
    Tex1: TLabel;
    BTag: TToolbarButton97;
    BCodeAbr: TToolbarButton97;
    BValider: TToolbarButton97;
    BFerme: TToolbarButton97;
    BAide: TToolbarButton97;
    Bdetag: TToolbarButton97;
    Dock: TDock97;
    procedure BFermeClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure BChercheClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure BdetagClick(Sender: TObject);
    procedure BTagClick(Sender: TObject);
    procedure FListeMouseDown(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure BValiderClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FListeKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure AxeChange(Sender: TObject);
    procedure BCodeAbrClick(Sender: TObject);
    procedure BAideClick(Sender: TObject);
  private
    UnAxe : String ;
    WMinX,WMinY    : Integer ;
    CodeCpteur : Boolean ;
    procedure WMGetMinMaxInfo(var MSG: Tmessage); message WM_GetMinMaxInfo;
    Procedure TagDetag(Avec : Boolean) ;
    Procedure CompteElemSelectionner ;
    Function  ListeVide : Boolean ;
    Procedure DetruitSectEtRub ;
    Procedure InverseSelection ;
    Procedure GetCellCanvas(Acol,ARow : LongInt ; Canvas : TCanvas; AState: TGridDrawState) ;
    Function  FaitUnObjet(UnObj : TSectInfo) : TCodBud ;
    Procedure CodeRubIncrement ;
  public
    { Déclarations publiques }
  end;


implementation

Uses SaisUtil,
     CabrSsec; // ParamCodeAbrege
{$R *.DFM}

Procedure GenerationSectionBudgetaire ;
var FCreSecBu : TFCreSecBu ;
    PP : THPanel ;
BEGIN
If VH^.DupSectBud then begin   {FP FQ15648}
  FCreSecBu:=TFCreSecBu.Create(Application) ;
  PP:=FindInsidePanel ;
  if PP=Nil then
     BEGIN
      Try
       FCreSecBu.ShowModal ;
      Finally
       FCreSecBu.Free ;
      End ;
     SourisNormale ;
     END else
     BEGIN
     InitInside(FCreSecBu,PP) ;
     FCreSecBu.Show ;
     END ;
end
{b FP FQ15648}
else
  HShowMessage('0;Génération des sections budgétaires;'+
    'Pour utiliser la génération des sections budgétaires#13#10Il faut cocher "La duplication section analytique section budgétaire";E;O;O;O;','','') ;
{e FP FQ15648}
END ;

procedure TFCreSecBu.FormCreate(Sender: TObject);
begin
PopUpMenu:=ADDMenuPop(PopUpMenu,'','') ;
WMinX:=Width ; WMinY:=Height ; CodeCpteur:=True ;
end;

procedure TFCreSecBu.WMGetMinMaxInfo(var MSG: Tmessage);
BEGIN
with PMinMaxInfo(MSG.lparam)^.ptMinTrackSize do begin X := WMinX ; Y := WMinY ; end;
END ;

procedure TFCreSecBu.BFermeClick(Sender: TObject);
begin
  Close ;
  if IsInside(Self) then
    CloseInsidePanel(Self) ;
end;

procedure TFCreSecBu.AxeChange(Sender: TObject);
begin CodeRubIncrement ; end;

procedure TFCreSecBu.FormShow(Sender: TObject);
begin
Axe.ItemIndex:=0 ; Signe.Value:='POS' ; Sens.Value:='M' ;
FListe.GetCellCanvas:=GetCellCanvas ; AxeChange(Nil) ; BCodeAbr.Enabled:=False ;
end;

Procedure TFCreSecBu.GetCellCanvas(Acol,ARow : LongInt ; Canvas : TCanvas; AState: TGridDrawState) ;
BEGIN
if FListe.Cells[FListe.ColCount-1,ARow]='*' then FListe.Canvas.Font.Style:=FListe.Canvas.Font.Style+[fsItalic]
                                            else FListe.Canvas.Font.Style:=FListe.Canvas.Font.Style-[fsItalic] ;
END ;

Procedure TFCreSecBu.TagDetag(Avec : Boolean) ;
Var  i : Integer ;
begin
for i:=1 to FListe.RowCount-1 do
    if Avec then FListe.Cells[FListe.ColCount-1,i]:='*'
            else FListe.Cells[FListe.ColCount-1,i]:='' ;
FListe.Invalidate ; FListe.SetFocus ;
Bdetag.Visible:=Avec ; BTag.Visible:=Not Avec ; CompteElemSelectionner ;
end;

procedure TFCreSecBu.BdetagClick(Sender: TObject);
begin TagDetag(False) ; end;

procedure TFCreSecBu.BTagClick(Sender: TObject);
begin TagDetag(True) ; end;

procedure TFCreSecBu.BChercheClick(Sender: TObject);
Var X : TSectInfo ;
    QSec : TQuery ;
Const UnCode = '00000' ;
begin
FListe.VidePile(True) ;
UnAxe:=Axe.Value ;
QSec:=OpenSql('Select * From SECTION Where S_AXE="'+UnAxe+'" '+
              'And S_SECTION Not In (Select BE_BUDSECT From BUDECR Where BE_AXE="'+UnAxe+'" ) ',True) ;
InitMove(RecordsCount(QSec),'') ;
While Not QSec.Eof do
   BEGIN
   MoveCur(False) ;
   X:=TSectInfo.Create ;
   X.Cod:=QSec.FindField('S_SECTION').AsString ;
   X.Lib:=QSec.FindField('S_LIBELLE').AsString ;
   X.Abr:=QSec.FindField('S_ABREGE').AsString ;
   X.Conf:=QSec.FindField('S_CONFIDENTIEL').AsString ;
   X.Lax:=UnAxe ;
   X.CodRub:=Copy(UnCode,1,5-Length(IntToStr(FListe.RowCount-1)))+IntToStr(FListe.RowCount-1) ;
   FListe.Cells[0,FListe.RowCount-1]:=X.Cod ; FListe.Cells[1,FListe.RowCount-1]:=X.Lib ;
   FListe.Cells[2,FListe.RowCount-1]:=X.CodRub ;
   FListe.Objects[0,FListe.RowCount-1]:=X ;
   FListe.Cells[FListe.ColCount-1,FListe.RowCount-1]:='*' ;
   FListe.RowCount:=FListe.RowCount+1 ;
   QSec.Next ;
   END ;
if FListe.RowCount>2 then FListe.RowCount:=FListe.RowCount-1 ;
FListe.Invalidate ; FiniMove ; Ferme(QSec) ;
CompteElemSelectionner ;
BCodeAbr.Enabled:=Not CodeCpteur ;
end;

Function TFCreSecBu.ListeVide : Boolean ;
BEGIN Result:=FListe.Cells[0,1]='' ; END ;

Procedure TFCreSecBu.CompteElemSelectionner ;
Var i,j : Integer ;
BEGIN
j:=0 ;
if Not ListeVide then for i:=1 to FListe.RowCount-1 do if FListe.Cells[FListe.ColCount-1,i]='*' then Inc(j) ;
Case j of
     0,1: BEGIN Nb1.Caption:=IntTostr(j) ; Tex1.Caption:=HM.Mess[0] ; END ;
     else BEGIN Nb1.Caption:=IntTostr(j) ; Tex1.Caption:=HM.Mess[1] ; END ;
   End ;
END ;

procedure TFCreSecBu.FormClose(Sender: TObject; var Action: TCloseAction);
begin
FListe.VidePile(True) ;
if Parent is THPanel then Action:=caFree ;
end;

procedure TFCreSecBu.FListeKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
if ((ssShift in Shift) And (Key=VK_DOWN)) or ((ssShift in Shift) And (Key=VK_UP)) then InverseSelection else
   if (Shift=[]) And (Key=VK_SPACE) then
      BEGIN
      InverseSelection ;
      if ((FListe.Row<FListe.RowCount-1) and (Key<>VK_SPACE)) then FListe.Row:=FListe.Row+1 ;
      END ;
end;

procedure TFCreSecBu.FListeMouseDown(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin if (ssCtrl in Shift) And (Button=mbLeft)then InverseSelection ; end;

Procedure TFCreSecBu.InverseSelection ;
BEGIN
if ListeVide then Exit ;
if FListe.Cells[FListe.ColCount-1,FListe.Row]='*' then FListe.Cells[FListe.ColCount-1,FListe.Row]:=''
                                                  else FListe.Cells[FListe.ColCount-1,FListe.Row]:='*' ;
CompteElemSelectionner ; FListe.Invalidate ;
END ;

Procedure TFCreSecBu.DetruitSectEtRub ;
Var
  QLoc,QLoc1 : TQuery ;
BEGIN

QLoc:=OpenSql('Select BS_RUB,BS_BUDSECT from BUDSECT Where BS_AXE="'+UnAxe+'" And BS_ATTENTE="-" And '+
              'BS_BUDSECT Not In (Select BE_BUDSECT From BUDECR Where BE_AXE="'+UnAxe+'" )',True) ;

QLoc1:=OpenSql('Select RB_RUBRIQUE,RB_FAMILLES From RUBRIQUE Where RB_NATRUB="BUD" And '+
               'RB_AXE="'+UnAxe+'" And (RB_FAMILLES="CBS;" Or RB_FAMILLES="G/S;" Or RB_FAMILLES="S/G;")',True) ;
InitMove(RecordsCount(QLoc),HM.Mess[3]) ;
While Not QLoc.Eof do
  BEGIN
  MoveCur(False) ; QLoc1.First ;
  While Not QLoc1.Eof do
     BEGIN
     if Pos(QLoc.FindField('BS_RUB').AsString,QLoc1.FindField('RB_RUBRIQUE').AsString)>0 then
        BEGIN
        ExecuteSQL('DELETE FROM RUBRIQUE WHERE RB_RUBRIQUE="'+QLoc1.FindField('RB_RUBRIQUE').AsString+'" AND RB_NATRUB="BUD" AND '+
                   'RB_AXE="'+UnAxe+'" AND RB_FAMILLES="'+QLoc1.FindField('RB_FAMILLES').AsString+'"');
        END ;
     QLoc1.Next ;
     END ;
  QLoc.Next ;
  END ;

ExecuteSQL('DELETE FROM BUDSECT WHERE BS_AXE="'+UnAxe+'" AND BS_ATTENTE="-" AND '+
           'BS_BUDSECT NOT IN (SELECT BE_BUDSECT FROM BUDECR WHERE BE_AXE="'+UnAxe+'" )');
Ferme(QLoc) ; Ferme(QLoc1) ; FiniMove ;
END ;

procedure TFCreSecBu.BValiderClick(Sender: TObject);
Var i : Integer ;
    Trouver : Boolean ;
    DoublonCod : Boolean ;
    X : TCodBud ;
    Li : TList ;
    T : Tob;
    Sect : TSectInfo;
begin
if ListeVide then Exit ; Trouver:=False ;
for i:=1 to FListe.RowCount-1 do
    if FListe.Cells[FListe.ColCount-1,i]='*' then BEGIN Trouver:=True ; Break ; END ;
if Not Trouver then Exit ;
if HM.Execute(2,'','')<>mrYes then Exit ;
BeginTrans ;
DetruitSectEtRub ;
InitMove(FListe.RowCount-1,HM.Mess[4]) ; DoublonCod:=False ;
Li:=TList.Create ;
T := Tob.Create('BUDSECT', nil, -1);
for i:=1 to FListe.RowCount-1 do
    BEGIN
    MoveCur(False) ;
    if FListe.Cells[FListe.ColCount-1,i]='*' then
       BEGIN
       Sect := TSectInfo(FListe.Objects[0,i]);
       if (PresenceComplexe('BUDSECT',['BS_BUDSECT','BS_AXE'],['=','='],[Sect.Cod,Sect.Lax],['S','S'])) or
          (PresenceComplexe('BUDSECT',['BS_RUB','BS_AXE'],['=','='],[Copy(Sect.CodRub,1,5),Sect.Lax],['S','S'])) or
          (Copy(Sect.CodRub,1,5)='') then
          BEGIN
          DoublonCod:=True ; X:=FaitUnObjet(Sect) ;
          Li.Add(X) ; Continue ;
          END ;
       T.InitValeurs;
       T.SetString('BS_BUDSECT', Sect.Cod);
       T.SetString('BS_LIBELLE', Sect.Lib);
       T.SetString('BS_ABREGE', Sect.Abr);
       T.SetString('BS_SECTIONRUB', Sect.Cod+';;');
       T.SetString('BS_SIGNE', Signe.Value);
       T.SetString('BS_SENS', Sens.Value);
       T.SetString('BS_CONFIDENTIEL', Sect.Conf);
       T.SetString('BS_AXE', Sect.Lax);
       T.SetString('BS_RUB', Copy(Sect.CodRub,1,5));
       T.InsertOrUpdateDB;
       END ;
    END ;
CommitTrans ; FiniMove ;
if DoublonCod then if HM.Execute(6,'','')=mrYes then MajCode(Li,'BUDSECT') ;
VideListe(Li) ; Li.Free ; T.Free; SourisNormale ;
end;

Function TFCreSecBu.FaitUnObjet(UnObj : TSectInfo) : TCodBud ;
Var X : TCodBud ;
BEGIN
X:=TCodBud.Create ;
X.UnCod:=UnObj.Cod ; X.UnLib:=Copy(UnObj.Lib,1,35) ; ; X.UnAbr:=Copy(X.UnLib,1,17) ;
X.UnCpR:=X.UnCod ; X.UnRub:=Copy(UnObj.CodRub,1,5) ; X.UnSig:=Signe.Value ;
X.UnSen:=Sens.Value ; X.UnAxe:=Axe.Value ;
Result:=X ;
END ;

Procedure TFCreSecBu.CodeRubIncrement ;
Var QLoc : TQuery ;
BEGIN
CodeCpteur:=False ;
QLoc:=OpenSql('Select * From STRUCRSE Where SS_AXE="'+Axe.Value+'"',True) ;
if QLoc.Eof then CodeCpteur:=True else
   if RecordsCount(QLoc)>4 then CodeCpteur:=True ;
Ferme(Qloc) ;
END ;

procedure TFCreSecBu.BCodeAbrClick(Sender: TObject);
Var LaListe : HTStrings ;
    i,j : Integer ;
    Cod,CodAbr,St : String ;
    Deb,Lon : Integer ;
begin
LaListe := HTStringList.Create ;
if ParamCodeAbrege(Axe.Value,LaListe,'') then
   BEGIN
   for i:=1 to FListe.RowCount-1 do Fliste.Cells[2,i]:='' ;
   for i:=0 to LaListe.Count-1 do
       BEGIN
       St:=LaListe.Strings[i] ;
       Deb:=StrToInt(ReadTokenSt(St)) ;
       Lon:=StrToInt(ReadTokenSt(St)) ;
       Cod:=ReadTokenSt(St) ;
       CodAbr:=ReadTokenSt(St) ;
       for j:=1 to FListe.RowCount-1 do
           if Copy(FListe.Cells[0,j],Deb,Lon)=Cod then
              BEGIN
              FListe.Cells[2,j]:=FListe.Cells[2,j]+CodAbr ;
              TSectInfo(FListe.Objects[0,j]).CodRub:=FListe.Cells[2,j] ;
              END ;
       END ;
   END ;
LaListe.Free ;
end;

procedure TFCreSecBu.BAideClick(Sender: TObject);
begin
CallHelpTopic(Self) ; 
end;

end.
