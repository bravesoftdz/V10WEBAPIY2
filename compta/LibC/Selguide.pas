unit SelGuide;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ExtCtrls, StdCtrls, Hctrls, DB, hmsgbox, DBTables, Buttons, Grids,
  DBGrids,Ent1,Hent1, HSysMenu, HTB97, HPanel, UiUtil, SaisUtil ;

type
  TFSelGuide = class(TForm)
    FListe: TDBGrid;
    PCriteres: TPanel;
    QGUIDES: TQuery;
    SGuides: TDataSource;
    GU_JOURNAL: THValComboBox;
    TGU_JOURNAL: THLabel;
    GU_DEVISE: THValComboBox;
    TGU_DEVISE: THLabel;
    GU_NATUREPIECE: THValComboBox;
    TGU_NATUREPIECE: THLabel;
    GU_ETABLISSEMENT: THValComboBox;
    TGU_ETABLISSEMENT: THLabel;
    GU_TYPE: TEdit;
    Bevel1: TBevel;
    HM: THMsgBox;
    HMTrad: THSystemMenu;
    Appelf: TToolbarButton97;
    OKBtn: TToolbarButton97;
    CancelBtn: TToolbarButton97;
    HelpBtn: TToolbarButton97;
    Dock: TDock97;
    HPB: TToolWindow97;
    TGU_LIBELLE: THLabel;
    GU_LIBELLE: TEdit;
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
    procedure AppelfClick(Sender: TObject);
    procedure FListeDblClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure GU_JOURNALClick(Sender: TObject);
    procedure HelpBtnClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure GU_LIBELLEChange(Sender: TObject);
    procedure CancelBtnClick(Sender: TObject);
  private
    Guide,Jal,NatP,Dev,Etab,LeCodeGuide : String3 ;
    AvecEuro                : Boolean ;
    procedure AppliqueCriteres  ;
  public
  end;

function SelectGuide(Jal,NatP,Dev,Etab : String3 ; AvecEuro : boolean ; LeCodeGuide : String = '' ) : String3 ;

implementation

uses CPGUIDE_TOM, Saisie ;

{$R *.DFM}

function SelectGuide(Jal,NatP,Dev,Etab : String3 ; AvecEuro : boolean ; LeCodeGuide : String = '' ) : String3 ;
var FSelGuide: TFSelGuide ;
    PP : THPanel ;
BEGIN
FSelGuide:=TFSelGuide.Create(Application) ;
FSelGuide.Jal:=Jal ; FSelGuide.NatP:=NatP ;
FSelGuide.Dev:=Dev ; FSelGuide.Etab:=Etab ;
FSelGuide.AvecEuro:=AvecEuro ;
FSelGuide.LeCodeGuide:=LeCodeGuide ;
PP:=FindInsidePanel ;
if PP=Nil then
   BEGIN
    try
     FSelGuide.ShowModal ;
    finally
     Result:=FSelGuide.Guide ;
     FSelGuide.Free ;
    end ;
   Screen.Cursor:=SyncrDefault ;
   END else
   BEGIN
   InitInside(FSelGuide,PP) ;
   FSelGuide.Show ;
   END ;
END ;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent Gendreau
Créé le ...... : 27/05/2002
Modifié le ... :   /  /    
Description .. : on n'ouvre pas les guides de saisie de tresorerie
Mots clefs ... : 
*****************************************************************}
procedure TFSelGuide.AppliqueCriteres ;
Var C     : TControl ;
    First : boolean ;
    i     : integer ;
    St    : String ;
BEGIN
QGUIDES.Close ;
QGUIDES.SQL.Clear ;
QGUIDES.SQL.Add('SELECT GU_GUIDE, GU_LIBELLE, GU_JOURNAL FROM GUIDE ') ;
QGUIDES.SQL.Add('LEFT JOIN JOURNAL ON GU_JOURNAL=J_JOURNAL ') ;
QGUIDES.SQL.Add('WHERE (J_MODESAISIE="" OR J_MODESAISIE="-") AND GU_TRESORERIE<>"X" AND ') ; //LG* 27/05/2002
First:=True ;
for i:=0 to PCriteres.ControlCount-1 do
  BEGIN
  St:='' ; C:=PCriteres.Controls[i] ;
  if C is THValComboBox then
    if (THValComboBox(C).Value<>'') then St:=St+THValComboBox(C).Name+'="'+THValComboBox(C).Value+'" ' ;
  if C is TEdit then
    if (TEdit(C).Text<>'') then
      BEGIN
      If C.Tag=1 Then
        BEGIN
        St:=St+TEdit(C).Name+' LIKE "'+TraduitJoker(TEdit(C).Text)+'" ' ;
        END Else St:=St+TEdit(C).Name+'="'+TEdit(C).Text+'" ' ;
      END ;
  if First then
    BEGIN
    if St<>'' then First:=False ;
    END else if St<>'' then St:='AND '+St ;
  if St<>'' then QGUIDES.SQL.Add(St) ;
  END ;
QGUIDES.SQL.ADD('ORDER BY GU_LIBELLE') ;
ChangeSQL(QGUIDES) ; //QGUIDES.Prepare ;
PrepareSQLODBC(QGUIDES) ;
QGUIDES.Open ;
FListe.Columns.Items[0].Title.Caption:=HM.Mess[0] ;
FListe.Columns.Items[0].Width:=70 ;
FListe.Columns.Items[1].Title.Caption:=HM.Mess[1] ;
FListe.Columns.Items[1].Width:=350 ;
FListe.Columns.Items[2].Title.Caption:=HM.Mess[2] ;
FListe.Columns.Items[2].Width:=70 ;
END ;

procedure TFSelGuide.FormCloseQuery(Sender: TObject; var CanClose: Boolean);
begin
if ((Not isInside(Self)) and (ModalResult=mrOk)) then Guide:=QGUIDES.FindField('GU_GUIDE').AsString ;
end;

procedure TFSelGuide.AppelfClick(Sender: TObject);
begin if not QGUIDES.Eof then ParamGuide (QGUIDES.FindField('GU_GUIDE').AsString,'NOR',taModif) ; end;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent Gendreau
Créé le ...... : 27/05/2002
Modifié le ... :   /  /    
Description .. : Suppression de warning
Mots clefs ... :
*****************************************************************}
procedure TFSelGuide.FListeDblClick(Sender: TObject);
Var R : RMVT ;
begin
if not (QGUIDES.Eof And QGUIDES.Bof) then
   BEGIN
   if isInside(Self) then
      BEGIN
      FillChar(R,Sizeof(R),#0) ;
      R.LeGuide:=QGUIDES.FindField('GU_GUIDE').AsString ;
      R.DateC:=V_PGI.DateEntree ; R.Exo:=QuelExoDT(R.DateC) ; R.Simul:='N' ;
      R.TypeGuide:='NOR' ; R.ANouveau:=False ; R.SaisieGuidee:=True ;
      LanceSaisie(Nil,taCreat,R) ;
      END else
      BEGIN
      ModalResult:=mrOk ;
      END ;
   END ;
end ;

procedure TFSelGuide.FormCreate(Sender: TObject);
begin
PopUpMenu:=ADDMenuPop(PopUpMenu,'','') ;
Guide:='' ;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent Gendreau
Créé le ...... : 27/05/2002
Modifié le ... :   /  /    
Description .. : suppression warning
Mots clefs ... : 
*****************************************************************}
procedure TFSelGuide.FormShow(Sender: TObject);
Var Okok : boolean ;
begin
if Jal<>'' then GU_JOURNAL.Value:=Jal else GU_JOURNAL.ItemIndex:=0 ;
if NatP<>'' then GU_NATUREPIECE.Value:=NatP else GU_NATUREPIECE.ItemIndex:=0 ;
if Dev<>'' then GU_DEVISE.Value:=Dev else GU_DEVISE.ItemIndex:=0 ;
if Etab<>'' then GU_ETABLISSEMENT.Value:=Etab else
   BEGIN
   GU_ETABLISSEMENT.ItemIndex:=0 ;
   PositionneEtabUser(GU_ETABLISSEMENT) ;
   END ;
AppliqueCriteres ;
if ((LeCodeGuide<>'') and (QGuides.Active)) then
   BEGIN
   QGuides.First ; Okok:=False ;
   Repeat
    if QGuides.FindField('GU_GUIDE').AsString=LeCodeGuide then BEGIN Okok:=True ; Break ; END ;
    QGuides.Next ;
   Until ((Okok) or (QGuides.EOF)) ;
   if Not Okok then QGuides.First ; 
   END ;
end;

procedure TFSelGuide.GU_JOURNALClick(Sender: TObject);
begin
AppliqueCriteres ;
end;

procedure TFSelGuide.HelpBtnClick(Sender: TObject);
begin
CallHelpTopic(Self) ; 
end;

procedure TFSelGuide.FormClose(Sender: TObject; var Action: TCloseAction);
begin
if Parent is THPanel then Action:=caFree ;
end;

procedure TFSelGuide.GU_LIBELLEChange(Sender: TObject);
begin
AppliqueCriteres ;
end;

procedure TFSelGuide.CancelBtnClick(Sender: TObject);
begin
  //SG6 06/12/2004
  Close;
  if IsInside(Self) then
    CloseInsidePanel(Self);
end;

end.
