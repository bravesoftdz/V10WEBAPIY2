{***********UNITE*************************************************
Auteur  ...... : Vincent Laroche
Créé le ...... : 27/01/2005
Modifié le ... :   /  /    
Description .. : Passage en eAGL
Mots clefs ... : 
*****************************************************************}
unit CodBuRup;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  hmsgbox, HSysMenu, ExtCtrls, Grids, Hctrls, StdCtrls, Buttons, ComCtrls, Ent1, HEnt1,
{$IFDEF EAGLCLIENT}
  UTob,
{$ELSE}
  DB,
  {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}
{$ENDIF}
{$IFDEF MODENT1}
  CPTypeCons,
{$ENDIF MODENT1}
  HStatus, HCompte, MajCodBu, ed_tools, HPanel, UiUtil, HTB97 ;

Procedure CreationCodBudSurRupture(Unfb : TFichierBase) ;

type
  TFCodBuRup = class(TForm)
    Pages: TPageControl;
    Pparam: TTabSheet;
    Bevel1: TBevel;
    TSens: THLabel;
    TAxe: THLabel;
    Sens: THValComboBox;
    Signe: THValComboBox;
    BCherche: TToolbarButton97;
    Axe: THValComboBox;
    FListe: THGrid;
    HPB: TToolWindow97;
    HMTrad: THSystemMenu;
    TCbRupt: TLabel;
    CbRupt: THValComboBox;
    HM: THMsgBox;
    Nb1: TLabel;
    Tex1: TLabel;
    BTag: TToolbarButton97;
    BImprimer: TToolbarButton97;
    BValider: TToolbarButton97;
    BFerme: TToolbarButton97;
    BAide: TToolbarButton97;
    Bdetag: TToolbarButton97;
    Dock: TDock97;
    procedure BFermeClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure AxeChange(Sender: TObject);
    procedure BChercheClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure BTagClick(Sender: TObject);
    procedure BdetagClick(Sender: TObject);
    procedure FListeKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure FListeMouseDown(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
    procedure BValiderClick(Sender: TObject);
    procedure BAideClick(Sender: TObject);
    procedure BImprimerClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
  private
    Unfb : TFichierBase ;
    WMinX,WMinY    : Integer ;
    NatRupt,PlanRupt,Ordre : String ;
    Deb : TabByte ;
    Lon : TabByte ;
    Cod : TabSt3 ;
    Lib : TabSt35 ;
    procedure WMGetMinMaxInfo(var MSG: Tmessage); message WM_GetMinMaxInfo;
    Procedure PositionneVariables ;
    Procedure TagDetag(Avec : Boolean) ;
    Procedure CompteElemSelectionner ;
    Procedure GetCellCanvas(Acol,ARow : LongInt ; Canvas : TCanvas; AState: TGridDrawState) ;
    Function  ListeVide : Boolean ;
    Procedure InverseSelection ;
    Procedure RempliLeGrille ;
    Procedure RunLaCreation ;
    Procedure RempliCbCodSect ;
    Function  FaitRubriqueSect(St : String) : String ;
    Function  FaitUnObjet(UnRow : HTStrings ; QuelTab : String) : TCodBud ;
  public
    { Déclarations publiques }
  end;


implementation

Uses
  {$IFDEF MODENT1}
  CPProcMetier,
  {$ENDIF MODENT1}
  SaisUtil,
{$IFDEF EAGLCLIENT}
  UtileAGL;
{$ELSE}
  PrintDBG;
{$ENDIF}

{$R *.DFM}

Procedure CreationCodBudSurRupture(Unfb : TFichierBase) ;
var FCodBuRup : TFCodBuRup ;
    PP : THPanel ;
BEGIN
FCodBuRup:=TFCodBuRup.Create(Application) ;
FCodBuRup.UnFb:=UnFb ;
PP:=FindInsidePanel ;
if PP=Nil then
   BEGIN
    Try
     FCodBuRup.ShowModal ;
    Finally
     FCodBuRup.Free ;
    End ;
   SourisNormale ;
   END else
   BEGIN
   InitInside(FCodBuRup,PP) ;
   FCodBuRup.Show ;
   END ;
END ;

procedure TFCodBuRup.WMGetMinMaxInfo(var MSG: Tmessage);
BEGIN with PMinMaxInfo(MSG.lparam)^.ptMinTrackSize do begin X := WMinX ; Y := WMinY ; end; END ;

procedure TFCodBuRup.FormCreate(Sender: TObject);
begin WMinX:=Width ; WMinY:=Height ; end;

procedure TFCodBuRup.BFermeClick(Sender: TObject);
begin
  Close ;
  if IsInside(Self) then
    CloseInsidePanel(Self) ;
end;

procedure TFCodBuRup.FormShow(Sender: TObject);
Var Delta : integer ; 
begin
FListe.GetCellCanvas:=GetCellCanvas ; Ordre:='' ;
Axe.Enabled:=UnFb<>fbGene ; TAxe.Enabled:=UnFb<>fbGene ;
if UnFb=fbGene then
   BEGIN
   CbRupt.ItemIndex:=0 ; FListe.Cells[2,0]:=HM.Mess[0] ;
   Caption:=HM.Mess[4] ;
   if Not V_PGI.OutLook then Pages.Height:=75 else
      BEGIN
      Axe.Visible:=False ; TAxe.Visible:=False ;
      Delta:=Tex1.Left-Nb1.Left ;
      Nb1.Left:=TAxe.Left ; Tex1.Left:=Nb1.Left+Delta ;  
      END ;
   HelpContext:=7577310 ;
   END else
   BEGIN
   FListe.Cells[2,0]:=HM.Mess[1] ;
   Caption:=HM.Mess[5] ; Pages.Height:=115 ;
   HelpContext:=7577400 ;
   END ;
If Axe.Values.Count>0 Then Axe.Value:=Axe.Values[0] ; Sens.ItemIndex:=0 ; Signe.ItemIndex:=0 ;
end;

Procedure TFCodBuRup.TagDetag(Avec : Boolean) ;
Var  i : Integer ;
begin
for i:=1 to FListe.RowCount-1 do
    BEGIN
    if Avec then
       BEGIN
       if FListe.Cells[0,FListe.RowCount-1]<>'' then FListe.Cells[FListe.ColCount-1,i]:='*'
                                                else FListe.Cells[FListe.ColCount-1,i]:='' ;
       END else FListe.Cells[FListe.ColCount-1,i]:='' ;
    END ;
FListe.Invalidate ; Bdetag.Visible:=Avec ; BTag.Visible:=Not Avec ;
CompteElemSelectionner ; FListe.SetFocus ;
end;

Function TFCodBuRup.ListeVide : Boolean ;
BEGIN Result:=FListe.Cells[0,1]='' ; END ;

Procedure TFCodBuRup.CompteElemSelectionner ;
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
          Nb1.Caption:=IntTostr(j) ; Tex1.Caption:=HM.Mess[2] ;
          END ;
     else BEGIN
          Nb1.Caption:=IntTostr(j) ; Tex1.Caption:=HM.Mess[3] ;
          END ;
   End ;
END ;

procedure TFCodBuRup.BTagClick(Sender: TObject);
begin TagDetag(True) ; end;

procedure TFCodBuRup.BdetagClick(Sender: TObject);
begin TagDetag(False) ; end;

Procedure TFCodBuRup.GetCellCanvas(Acol,ARow : LongInt ; Canvas : TCanvas; AState: TGridDrawState) ;
BEGIN
if FListe.Cells[FListe.ColCount-1,ARow]='*' then FListe.Canvas.Font.Style:=FListe.Canvas.Font.Style+[fsItalic]
                                            else FListe.Canvas.Font.Style:=FListe.Canvas.Font.Style-[fsItalic] ;
END ;

procedure TFCodBuRup.FListeKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
if (ssShift in Shift) And (Key=VK_DOWN) or (ssShift in Shift) And (Key=VK_UP)then InverseSelection else
   if (Shift=[]) And (Key=VK_SPACE) then
      BEGIN
      InverseSelection ;
      if ((FListe.Row<FListe.RowCount-1) and (Key<>VK_SPACE)) then FListe.Row:=FListe.Row+1 ;
      END ;
end;

procedure TFCodBuRup.FListeMouseDown(Sender: TObject; Button: TMouseButton;
  Shift: TShiftState; X, Y: Integer);
begin if (ssCtrl in Shift) And (Button=mbLeft)then InverseSelection ; end;

Procedure TFCodBuRup.InverseSelection ;
BEGIN
if ListeVide then Exit ;
if FListe.Cells[FListe.ColCount-1,FListe.Row]='*' then FListe.Cells[FListe.ColCount-1,FListe.Row]:=''
                                                  else FListe.Cells[FListe.ColCount-1,FListe.Row]:='*' ;
CompteElemSelectionner ; FListe.Invalidate ;
END ;

procedure TFCodBuRup.AxeChange(Sender: TObject);
begin
if UnFb=fbGene then Exit ;
Case Axe.Value[2] of
     '1' : CbRupt.DataType:='ttRuptSect1' ;
     '2' : CbRupt.DataType:='ttRuptSect2' ;
     '3' : CbRupt.DataType:='ttRuptSect3' ;
     '4' : CbRupt.DataType:='ttRuptSect4' ;
     '5' : CbRupt.DataType:='ttRuptSect5' ;
   End ;
if CbRupt.Values.Count>0 then CbRupt.Value:=CbRupt.Values[0] ;
end;

Procedure TFCodBuRup.PositionneVariables ;
BEGIN
Case UnFb of
     fbGene : NatRupt:='RUG' ;
     else     NatRupt:='RU'+Axe.Value[2] ;
   End ;
PlanRupt:=CbRupt.Value ;
END ;

procedure TFCodBuRup.BChercheClick(Sender: TObject);
begin
PositionneVariables ;
if UnFb<>fbGene then BEGIN RempliCbCodSect ; ChargeTablo(Axe.Value,Deb,Lon,Cod,Lib) ; END ;
RempliLeGrille ;
end;

Procedure TFCodBuRup.RempliCbCodSect ;
Var QLoc : TQuery ;
BEGIN
QLoc:=OpenSql('Select CC_LIBRE From CHOIXCOD Where CC_TYPE="RU'+Axe.Value[2]+'" And CC_CODE="'+CbRupt.Value+'"' ,True) ;
Ordre:=QLoc.Fields[0].AsString ; Ferme(QLoc) ;
END ;

Function TFCodBuRup.FaitRubriqueSect(St : String) : String ;
Var i,j : Integer ;
    StOrd,StTemp,St1,St2 : String ;
BEGIN
Result:='' ; if Ordre='' then Exit ;
StOrd:=Ordre ; St1:='' ;
for i:=1 to VH^.Cpta[AxeToFb(Axe.Value)].Lg do St1:=St1+'?' ;
While StOrd<>'' do
  BEGIN
  StTemp:=ReadTokenSt(StOrd) ;
  if StTemp='' then Continue ; j:=0 ;
  for i:=1 to 17 do if Cod[i]=StTemp then BEGIN j:=i ; Break ; END ;
  St2:=Copy(St,1,Lon[j]) ; Delete(St,1,Lon[j]) ;
  Delete(St1,Deb[j],Lon[j]) ; Insert(St2,St1,Deb[j]) ;
  END ;
j:=Length(St1) ;
for i:=j downto 1 do if St1[i]<>'?' then Break else St1:=Copy(St1,1,i-1) ;
Result:=St1 ;
END ;

Procedure TFCodBuRup.RempliLeGrille ;
Var QLoc : TQuery ;
    St : String ;
BEGIN
FListe.VidePile(False) ;
QLoc:=OpenSql('Select RU_CLASSE,RU_LIBELLECLASSE From RUPTURE Where RU_NATURERUPT="'+NatRupt+'" '+
              'And RU_PLANRUPT="'+PlanRupt+'"',True) ;
While Not QLoc.Eof do
   BEGIN
   St:=QLoc.Fields[0].AsString ;
   if St[Length(St)]='x' then St:=Copy(St,1,Length(St)-1) ;
   FListe.Cells[0,FListe.RowCount-1]:=St ;
   FListe.Cells[1,FListe.RowCount-1]:=QLoc.Fields[1].AsString ;
   if UnFb=fbGene then FListe.Cells[2,FListe.RowCount-1]:=St
                  else FListe.Cells[2,FListe.RowCount-1]:=FaitRubriqueSect(St) ;
   FListe.Cells[3,FListe.RowCount-1]:=Copy(St,1,5) ;
   FListe.RowCount:=FListe.RowCount+1 ; QLoc.Next ;
   END ;
if FListe.RowCount>2 then FListe.RowCount:=FListe.RowCount-1 ;
Ferme(QLoc) ; BTagClick(Nil) ;
END ;

procedure TFCodBuRup.BValiderClick(Sender: TObject);
Var i : Integer ;
    Trouver : Boolean ;
    io : TIoErr ;
begin
Trouver:=False ;
for i:=1 to FListe.RowCount-1 do
    if FListe.Cells[FListe.ColCount-1,i]='*' then BEGIN Trouver:=True ; Break ; END ;
if Not Trouver then BEGIN HM.Execute(6,'','') ; Exit ; END ;
if HM.Execute(7,'','')<>mrYes then Exit ;
io:=Transactions(RunLaCreation,2) ;
if io<>oeOk then MessageAlerte(HM.Mess[9]) else HM.Execute(10,'','') ;
SourisNormale ;
end;

Function TFCodBuRup.FaitUnObjet(UnRow : HTStrings ; QuelTab : String) : TCodBud ;
Var X : TCodBud ;
    i : Integer ;
BEGIN
X:=TCodBud.Create ;
for i:=0 to UnRow.Count-1 do
   BEGIN
   Case i of
      0 : X.UnCod:=UnRow.Strings[i] ;
      1 : BEGIN X.UnLib:=Copy(UnRow.Strings[i],1,35) ; X.UnAbr:=Copy(UnRow.Strings[i],1,17) ; END ;
      2 : X.UnCpR:=UnRow.Strings[i] ;
      3 : X.UnRub:=UnRow.Strings[i] ;
    End ;
   END ;
X.UnSig:=Signe.Value ; X.UnSen:=Sens.Value ;
if QuelTab='BUDSECT' then X.UnAxe:=Axe.Value else X.UnAxe:='' ;
Result:=X ;
END ;

Procedure TFCodBuRup.RunLaCreation ;
Var QLoc : TQuery ;
    Sql : String ;
    Cle,Table,CRub,CompteRub,Pref : String ;
    i : Integer ;
    DoublonCod : Boolean ;
    Sig,Sen : String ;
    X : TCodBud ;
    Li : TList ;
BEGIN
if UnFb=fbGene then
   BEGIN
   Sql:='Select * from BUDGENE Where BG_BUDGENE="'+W_W+'"' ;
   Cle:='BG_BUDGENE' ; Table:='BUDGENE' ; CompteRub:='BG_COMPTERUB' ;
   Pref:='BG_' ;
   END else
   BEGIN
   Sql:='Select * from BUDSECT Where BS_BUDSECT="'+W_W+'"' ;
   Cle:='BS_BUDSECT' ; Table:='BUDSECT' ; CompteRub:='BS_SECTIONRUB' ;
   Pref:='BS_' ;
   END ;
CRub:=Pref+'RUB' ; DoublonCod:=False ;
QLoc:=OpenSql(Sql,False) ;
Sig:=Signe.Value ; Sen:=Sens.Value ;
InitMove(FListe.RowCount-1,'') ;
Li:=TList.Create ;
for i:=1 to Fliste.RowCount-1 do
   BEGIN
   MoveCur(False) ;
   if FListe.Cells[FListe.ColCount-1,i]='*' then
      BEGIN
      if Presence(Table,Cle,FListe.Cells[0,i]) or
         Presence(Table,CRub,Copy(FListe.Cells[0,i],1,5))then
         BEGIN
         DoublonCod:=True ; X:=FaitunObjet(FListe.Rows[i],Table) ;
         Li.Add(X) ; Continue ;
         END ;
      QLoc.Insert ; InitNew(QLoc) ;
      QLoc.FindField(Cle).AsString:=FListe.Cells[0,i] ;
      QLoc.FindField(Pref+'LIBELLE').AsString:=Copy(FListe.Cells[1,i],1,35) ;
      QLoc.FindField(Pref+'ABREGE').AsString:=Copy(FListe.Cells[1,i],1,17) ;
      QLoc.FindField(Pref+'SIGNE').AsString:=Sig ;
      QLoc.FindField(Pref+'SENS').AsString:=Sen ;
      QLoc.FindField(CRub).AsString:=Copy(FListe.Cells[0,i],1,5) ;
      QLoc.FindField(CompteRub).AsString:=FListe.Cells[2,i]+';;' ;
      if Table='BUDSECT' then QLoc.FindField(Pref+'AXE').AsString:=Axe.Value ;
      QLoc.Post ;
      END ;
   END ;
Ferme(QLoc) ; FiniMove ;
if DoublonCod then if HM.Execute(8,'','')=mrYes then MajCode(Li,Table) ;
VideListe(Li) ; Li.Free ;
END ;

procedure TFCodBuRup.BAideClick(Sender: TObject);
begin CallHelpTopic(Self) ; end;

procedure TFCodBuRup.BImprimerClick(Sender: TObject);
{$IFDEF EAGLCLIENT}
var
  T, F : Tob;
  i : Integer;
{$ENDIF}
begin
{$IFDEF EAGLCLIENT}
  SourisSablier;
  T := TOB.Create('non', nil, -1);
  for i := 1 to FListe.RowCount-1 do begin
    F := TOB.Create('non', T, -1);
    F.AddChampSup('C1', False);
    F.AddChampSup('L1', False);
    F.AddChampSup('RUB', False);
    F.AddChampSup('CRUB', False);
    if i=1 then begin
      F.AddChampSup('TITRE', False);
      F.SetString('TITRE', Caption);
    end;
    F.SetString('C1', FListe.Cells[0, i]);
    F.SetString('L1', FListe.Cells[1, i]);
    F.SetString('RUB', FListe.Cells[2, i]);
    F.SetString('CRUB', FListe.Cells[3, i]);
  end;
  SourisNormale;
  LanceEtatTob('E','CST','CBS',T, True, False, False, nil, '', Caption, False, 0, '', 0, '');
  T.Free;
{$ELSE}
  PrintDBGrid(FListe,Nil,Caption,'') ;
{$ENDIF}
end;

procedure TFCodBuRup.FormClose(Sender: TObject; var Action: TCloseAction);
begin
FListe.VidePile(True) ;
if Parent is THPanel then Action:=caFree ;
end;

end.
