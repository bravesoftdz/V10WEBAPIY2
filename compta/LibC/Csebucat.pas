{***********UNITE*************************************************
Auteur  ...... : Vincent Laroche
Créé le ...... : 27/01/2005
Modifié le ... :   /  /    
Description .. : Passage en eAGL
Mots clefs ... : 
*****************************************************************}
unit Csebucat;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Grids, Hctrls, ExtCtrls, StdCtrls, Buttons, ComCtrls, Ent1, HEnt1,
{$IFDEF EAGLCLIENT}

{$ELSE}
  DB,
  {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}
{$ENDIF}
{$IFDEF MODENT1}
  CPTypeCons,
{$ENDIF MODENT1}
  hmsgbox, HSysMenu, MajCodBu, HStatus, HPanel, UTob,
  UIUtil; // MODIF PACK AVANCE pour gestion mode inside

Procedure GenSecBudCategorie ;

Type TLib = Class
     UnLib : String ;
     END ;

Type TSectInfoCat = Class
     Cod : String ;
     Lib : String ;
     Abr : String ;
     Lax : String ;
     Categorie : String ;
     CodRub : String ;
     End ;
type
  TFCsebucat = class(TForm)
    Pages: TPageControl;
    Param: TTabSheet;
    Bevel1: TBevel;
    TSigne: THLabel;
    TSens: THLabel;
    TAxe: THLabel;
    Sens: THValComboBox;
    Signe: THValComboBox;
    Axe: THValComboBox;
    BCherche: THBitBtn;
    HPB: TPanel;
    Tex1: TLabel;
    Nb1: TLabel;
    Panel1: TPanel;
    Bdetag: THBitBtn;
    BTag: THBitBtn;
    BAide: THBitBtn;
    BFerme: THBitBtn;
    BValider: THBitBtn;
    BCodeAbr: THBitBtn;
    TcbCatjal: THLabel;
    cbCatjal: THValComboBox;
    FListe: THGrid;
    HMTrad: THSystemMenu;
    HM: THMsgBox;
    procedure FormShow(Sender: TObject);
    procedure AxeChange(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure BFermeClick(Sender: TObject);
    procedure BTagClick(Sender: TObject);
    procedure BdetagClick(Sender: TObject);
    procedure BChercheClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FListeKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure FListeMouseDown(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
    procedure BValiderClick(Sender: TObject);
    procedure BAideClick(Sender: TObject);
    procedure BCodeAbrClick(Sender: TObject);
  private
    WMinX,WMinY    : Integer ;
    UnFb : TFichierBase ;
    TabSousPlan : TSousPlanCat ;
    LiC : HTStringList ;
    CodeCpteur : Boolean ;
    procedure WMGetMinMaxInfo(var MSG: Tmessage); message WM_GetMinMaxInfo;
    Procedure TagDetag(Avec : Boolean) ;
    Procedure CompteElemSelectionner ;
    Function  ListeVide : Boolean ;
    Function  InitCodSect : String ;
    Function  CreerUnObj(Lib : String) : TLib ;
    Procedure GenereListe(StCod : String ; Var Lib : String ; i : Integer) ;
    Procedure ClearLaListe ;
    Procedure GetCellCanvas(Acol,ARow : LongInt ; Canvas : TCanvas; AState: TGridDrawState) ;
    Procedure InverseSelection ;
    Procedure DetruitSectEtRub ;
    Function  FaitUnObjet(i : Integer) : TCodBud ;
    Function  FaitCodeRub(StC : String) : String ;
    Procedure CodeRubIncrement ;
    Function  TransformeCode(StC : String) : String ;
    Function  FaitCodeRubAbr(StC : String) : String ;
  public
    { Déclarations publiques }
  end;

implementation

{$R *.DFM}

Uses
  {$IFDEF MODENT1}
  CPProcMetier,
  {$ENDIF MODENT1}
  SaisUtil, CabrSsec, ed_Tools ;

{***********A.G.L.***********************************************
Auteur  ...... : 
Créé le ...... : 09/09/2003
Modifié le ... :   /  /
Description .. : 
Suite ........ : 09/09/2003, SBO : MODIF PACK AVANCE pour gestion 
Suite ........ : mode inside
Mots clefs ... : 
*****************************************************************}
Procedure GenSecBudCategorie ;
var FCsebucat : TFCsebucat ;
    PP : THPanel ;
BEGIN
  FCsebucat:=TFCsebucat.Create(Application) ;

  PP:=FindInsidePanel ;
  if PP=Nil then
    begin
    Try
      FCsebucat.ShowModal ;
      Finally
      FCsebucat.Free ;
      End ;
    end
  else
    begin
    InitInside(FCsebucat,PP) ;
    FCsebucat.Show ;
    end ;

  SourisNormale ;
END ;

procedure TFCsebucat.WMGetMinMaxInfo(var MSG: Tmessage);
BEGIN with PMinMaxInfo(MSG.lparam)^.ptMinTrackSize do begin X := WMinX ; Y := WMinY ; end; END ;

procedure TFCsebucat.BAideClick(Sender: TObject);
begin CallHelpTopic(Self) ; end;

procedure TFCsebucat.FormCreate(Sender: TObject);
begin
PopUpMenu:=ADDMenuPop(PopUpMenu,'','') ; WMinX:=Width ; WMinY:=Height ;
FillChar(TabSousPlan,SizeOf(TabSousPlan),#0) ;
LiC:=HTStringList.Create ;
end;

Procedure TFCsebucat.ClearLaListe ;
Var i : Integer ;
BEGIN
for i:=0 to LiC.Count-1 do
   BEGIN
   if Lic.Objects[i]<>Nil then BEGIN TObject(Lic.Objects[i]).Free ; Lic.Objects[i]:=Nil ; END ;
   END ;
LiC.Clear ;
END ;

procedure TFCsebucat.FormClose(Sender: TObject; var Action: TCloseAction);
begin ClearLaListe ; LiC.Free ; end;

procedure TFCsebucat.FormShow(Sender: TObject);
begin
If Axe.Values.Count>0 Then Axe.Value:=Axe.Values[0] ; Signe.Value:='POS' ; Sens.Value:='M' ;
FListe.GetCellCanvas:=GetCellCanvas ; CodeRubIncrement ; BCodeAbr.Enabled:=False ;
end;

procedure TFCsebucat.BFermeClick(Sender: TObject);
begin
  Close ;
  if IsInside(Self) then
    CloseInsidePanel(Self);
end;

Procedure TFCsebucat.GetCellCanvas(Acol,ARow : LongInt ; Canvas : TCanvas; AState: TGridDrawState) ;
BEGIN
if FListe.Cells[FListe.ColCount-1,ARow]='*' then FListe.Canvas.Font.Style:=FListe.Canvas.Font.Style+[fsItalic]
                                            else FListe.Canvas.Font.Style:=FListe.Canvas.Font.Style-[fsItalic] ;
END ;

procedure TFCsebucat.BTagClick(Sender: TObject);
begin TagDetag(True) ; end;

procedure TFCsebucat.BdetagClick(Sender: TObject);
begin TagDetag(False) ; end;

Procedure TFCsebucat.TagDetag(Avec : Boolean) ;
Var  i : Integer ;
begin
for i:=1 to FListe.RowCount-1 do
    if Avec then FListe.Cells[FListe.ColCount-1,i]:='*'
            else FListe.Cells[FListe.ColCount-1,i]:='' ;
FListe.Invalidate ; FListe.SetFocus ;
Bdetag.Visible:=Avec ; BTag.Visible:=Not Avec ; CompteElemSelectionner ;
end;

Function TFCsebucat.ListeVide : Boolean ;
BEGIN Result:=FListe.Cells[0,1]='' ; END ;

Procedure TFCsebucat.InverseSelection ;
BEGIN
if ListeVide then Exit ;
if FListe.Cells[FListe.ColCount-1,FListe.Row]='*' then FListe.Cells[FListe.ColCount-1,FListe.Row]:=''
                                                  else FListe.Cells[FListe.ColCount-1,FListe.Row]:='*' ;
CompteElemSelectionner ; FListe.Invalidate ;
END ;

Procedure TFCsebucat.CompteElemSelectionner ;
Var i,j : Integer ;
BEGIN
j:=0 ;
if Not ListeVide then for i:=1 to FListe.RowCount-1 do if FListe.Cells[FListe.ColCount-1,i]='*' then Inc(j) ;
Case j of
     0,1: BEGIN Nb1.Caption:=IntTostr(j) ; Tex1.Caption:=HM.Mess[0] ; END ;
     else BEGIN Nb1.Caption:=IntTostr(j) ; Tex1.Caption:=HM.Mess[1] ; END ;
   End ;
END ;

procedure TFCsebucat.AxeChange(Sender: TObject);
Var QLoc : TQuery ;
begin
cbCatjal.Values.Clear ; cbCatjal.Items.Clear ;
QLoc:=OpenSql('Select CC_CODE,CC_LIBELLE From CHOIXCOD Where CC_TYPE="CJB" And '+
              'CC_ABREGE="'+Axe.Value+'"',True) ;
While Not QLoc.Eof do
   BEGIN
   cbCatjal.Values.Add(QLoc.Fields[0].AsString) ; cbCatjal.Items.Add(QLoc.Fields[1].AsString) ;
   QLoc.Next ;
   END ;
if cbCatjal.Values.Count>0 then cbCatjal.Value:=cbCatjal.Values[0] ;
Ferme(QLoc) ;
end;

Function TFCsebucat.InitCodSect : String ;
Var i : Integer ;
    St : String ;
BEGIN for i:=1 to VH^.Cpta[UnFb].Lg do St:=St+'.' ; Result:=St ; END ;

Function TFCsebucat.CreerUnObj(Lib : String) : TLib ;
Var X : TLib ;
BEGIN X:=TLib.Create ; X.UnLib:=Lib ; Result:=X ; END ;

Procedure TFCsebucat.GenereListe(StCod : String ; Var Lib : String ; i : Integer) ;
Var j : Integer ;
    X : TLib ;
    St : String ;
    Deb,Lon : Integer ;
BEGIN
for j:=0 to TabSousPlan[i].ListeSP.Count-1 do
   BEGIN
   St:=TabSousPlan[i].ListeSP.Strings[j] ;
   Deb:=TabSousPlan[i].Debut ; Lon:=TabSousPlan[i].Longueur ;
   Insert(ReadTokenSt(St),StCod,Deb) ; Delete(StCod,Deb+Lon,Lon) ;
   if (i+1<=MaxSousPlan) And (TabSousPlan[i+1].Code<>'') then
      BEGIN Lib:=Lib+ReadTokenSt(St) ; GenereListe(StCod,Lib,i+1) END else
      BEGIN X:=CreerUnObj(Lib+ReadTokenSt(St)) ; Lic.AddObject(StCod,X) ; END ;
   END ;
Lib:='' ;
END ;

Procedure TFCsebucat.CodeRubIncrement ;
Var QLoc : TQuery ;
BEGIN
CodeCpteur:=False ;
QLoc:=OpenSql('Select Count(*) From STRUCRSE Where SS_AXE="'+Axe.Value+'"',True) ;
CodeCpteur:=(QLoc.Fields[0].AsInteger>=4) ;
Ferme(Qloc) ;
END ;

procedure TFCsebucat.BChercheClick(Sender: TObject);
Var StCod,Lib,St : String ;
    i : Integer ;
Const UnCode = '00000' ;
begin
BCodeAbr.Enabled:=False ;
if cbCatjal.Values.Count<=0 then Exit ;
UnFb:=AxeToFb(Axe.Value) ;
TabSousPlan:=SousPlanCat(cbCatjal.Value,False) ;
StCod:=InitCodSect ; ClearLaListe ; i:=1 ; Lib:='' ; FListe.VidePile(False) ;
GenereListe(StCod,Lib,i) ;
for i:=0 to Lic.Count-1 do
    BEGIN
    FListe.Cells[0,FListe.RowCount-1]:=FaitCodeRub(Lic.Strings[i]) ;
    FListe.Cells[1,FListe.RowCount-1]:=TLib(Lic.Objects[i]).UnLib ;
    if CodeCpteur then FListe.Cells[2,FListe.RowCount-1]:=Copy(UnCode,1,5-Length(IntToStr(FListe.RowCount-1)))+IntToStr(FListe.RowCount-1) else
       BEGIN
       St:=FaitCodeRubAbr(Lic.Strings[i]) ;
       if Length(St)<=5 then FListe.Cells[2,FListe.RowCount-1]:=St
                        else FListe.Cells[2,FListe.RowCount-1]:=Copy(UnCode,1,5-Length(IntToStr(FListe.RowCount-1)))+IntToStr(FListe.RowCount-1) ;
       END ;
    FListe.Cells[FListe.ColCount-1,FListe.RowCount-1]:='*' ;
    FListe.RowCount:=FListe.RowCount+1 ;
    END ;
if FListe.RowCount>2 then FListe.RowCount:=FListe.RowCount-1 ;
CompteElemSelectionner ; BCodeAbr.Enabled:=True ;
end;

procedure TFCsebucat.FListeMouseDown(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin if (ssCtrl in Shift) And (Button=mbLeft)then InverseSelection ; end;

procedure TFCsebucat.FListeKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
if ((ssShift in Shift) And (Key=VK_DOWN)) or ((ssShift in Shift) And (Key=VK_UP)) then InverseSelection else
   if (Shift=[]) And (Key=VK_SPACE) then
      BEGIN
      InverseSelection ;
      if ((FListe.Row<FListe.RowCount-1) and (Key<>VK_SPACE)) then FListe.Row:=FListe.Row+1 ;
      END ;
end;

Procedure TFCsebucat.DetruitSectEtRub ;
Var QLoc,QLoc1 : TQuery ;
    UnAxe : String ;
BEGIN
UnAxe:=Axe.Value ;

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
           'BS_CATEGORIE="'+cbCatjal.Value+'" AND BS_BUDSECT NOT IN (SELECT BE_BUDSECT FROM BUDECR WHERE BE_AXE="'+UnAxe+'" )');
Ferme(QLoc) ; Ferme(QLoc1) ; FiniMove ;
END ;

procedure TFCsebucat.BValiderClick(Sender: TObject);
Var i : Integer ;
    Trouver : Boolean ;
    DoublonCod : Boolean ;
    X : TCodBud ;
    Li : TList ;
    UnAxe : String ;
    T : Tob;
begin
if ListeVide then Exit ; Trouver:=False ;
for i:=1 to FListe.RowCount-1 do if FListe.Cells[FListe.ColCount-1,i]='*' then BEGIN Trouver:=True ; Break ; END ;
if Not Trouver then Exit ;
if HM.Execute(2,'','')<>mrYes then Exit ;
UnAxe:=Axe.Value ;
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
       if (PresenceComplexe('BUDSECT',['BS_BUDSECT','BS_AXE'],['=','='],[FListe.Cells[0,i],UnAxe],['S','S'])) or
          (PresenceComplexe('BUDSECT',['BS_BUDSECT','BS_RUB','BS_AXE'],['=','=','='],[FListe.Cells[0,i],Copy(FListe.Cells[2,i],1,5),UnAxe],['S','S','S'])) or
          (Copy(FListe.Cells[2,i],1,5)='') then
          BEGIN
          DoublonCod:=True ; X:=FaitUnObjet(i) ;
          Li.Add(X) ; Continue ;
          END ;
       T.InitValeurs;
       T.SetString('BS_BUDSECT', FListe.Cells[0,i]);
       T.SetString('BS_LIBELLE', FListe.Cells[1,i]);
       T.SetString('BS_ABREGE', Copy(FListe.Cells[1,i],1,17));
       T.SetString('BS_SECTIONRUB', TransFormeCode(FListe.Cells[0,i])+';');
       T.SetString('BS_SIGNE', Signe.Value);
       T.SetString('BS_SENS', Sens.Value);
       T.SetString('BS_CONFIDENTIEL', '0');
       T.SetString('BS_AXE', Axe.Value);
       T.SetString('BS_RUB', Copy(FListe.Cells[2,i],1,5));
       T.SetString('BS_CATEGORIE', cbCatjal.Value);
       T.InsertOrUpdateDB;
       END ;
    END ;
CommitTrans ; FiniMove ;
if DoublonCod then if HM.Execute(6,'','')=mrYes then MajCode(Li,'BUDSECT') ;
VideListe(Li) ; Li.Free ; T.Free; SourisNormale ;
end ;

Function TFCsebucat.FaitCodeRub(StC : String) : String ;
Var i,j : Integer ;
BEGIN
j:=Length(Stc) ;
for i:=j Downto 1 do if Stc[i]<>'.' then Break else Stc:=Copy(StC,1,i-1) ;
Result:=StC ;
END ;

Function TFCsebucat.FaitCodeRubAbr(StC : String) : String ;
Var i : Integer ;
    St : String ;
BEGIN
St:='' ;
for i:=1 to Length(Stc) do if StC[i]<>'.' then St:=St+StC[i] ;
Result:=St ;
END ;

Function TFCsebucat.TransformeCode(StC : String) : String ;
BEGIN
StC:=FaitCodeRub(StC) ;
While Pos('.',Stc)>0 do Stc[Pos('.',StC)]:='?' ;
Result:=StC ;
END ;

Function TFCsebucat.FaitUnObjet(i : Integer) : TCodBud ;
Var X : TCodBud ;
BEGIN
X:=TCodBud.Create ;
X.UnCod:=FListe.Cells[0,i] ; X.UnLib:=Copy(FListe.Cells[1,i],1,35) ; ; X.UnAbr:=Copy(FListe.Cells[1,i],1,17) ;
X.UnCpR:=FaitCodeRub(FListe.Cells[0,i]) ;
X.UnRub:=Copy(FListe.Cells[2,i],1,5) ; X.UnSig:=Signe.Value ;
X.UnSen:=Sens.Value ; X.UnAxe:=Axe.Value ; X.UnCat:=cbCatjal.Value ;
Result:=X ;
END ;

procedure TFCsebucat.BCodeAbrClick(Sender: TObject);
Var LaListe : HTStrings ;
    i,j : Integer ;
    Cod,CodAbr,St : String ;
    Deb,Lon : Integer ;
    QLoc : TQuery ;
begin
QLoc:=OpenSql('Select CC_LIBRE From CHOIXCOD Where CC_TYPE="CJB" '+
              'And CC_CODE="'+cbCatjal.Value+'" And CC_ABREGE="'+Axe.Value+'"',True) ;
St:=Copy(QLoc.Fields[0].AsString,1,Pos('/',QLoc.Fields[0].AsString)-1) ;
Ferme(QLoc) ;
LaListe:=HTstringList.Create ;
if ParamCodeAbrege(Axe.Value,LaListe,St) then
   BEGIN
   if LaListe.Count>0 then for i:=1 to FListe.RowCount-1 do FListe.Cells[2,i]:='' ;
   for i:=0 to LaListe.Count-1 do
       BEGIN
       St:=LaListe.Strings[i] ;
       Deb:=StrToInt(ReadTokenSt(St)) ;
       Lon:=StrToInt(ReadTokenSt(St)) ;
       Cod:=ReadTokenSt(St) ;
       CodAbr:=ReadTokenSt(St) ;
       for j:=1 to FListe.RowCount-1 do if Copy(FListe.Cells[0,j],Deb,Lon)=Cod then FListe.Cells[2,j]:=FListe.Cells[2,j]+CodAbr ;
       END ;
   END ;
LaListe.Free ;
end;

end.
