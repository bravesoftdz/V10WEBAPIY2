unit GuidTool;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, Buttons, Spin, Hctrls, Ent1, HEnt1, ExtCtrls, HSysMenu
  {$IFDEF EAGLCLIENT}
  ,uTob
  {$ELSE}
  ,DB
  {$IFNDEF DBXPRESS},dbtables{$ELSE},uDbxDataSet{$ENDIF}
  {$ENDIF}
   ;

function ChoixChampZone ( CurLig : integer ; Nature : String ) : String ;

type
  TFGuidTool = class(TForm)
    FTable: THValComboBox;
    FListe: TListBox;
    H_Table: THLabel;
    FChamp: TEdit;
    FNumL: TSpinEdit;
    H_NumL: THLabel;
    H_Champ: THLabel;
    FCache: THValComboBox;
    GDate: TGroupBox;
    FFormat: TComboBox;
    H_Format: THLabel;
    H_ExempleD: THLabel;
    FExempleD: TEdit;
    GString: TGroupBox;
    H_Debut: THLabel;
    FDebut: TSpinEdit;
    H_Longueur: THLabel;
    FLongueur: TSpinEdit;
    H_exempleS: THLabel;
    FExempleS: TEdit;
    PBouton: TPanel;
    BValider: THBitBtn;
    BFermer: THBitBtn;
    BAide: THBitBtn;
    HMTrad: THSystemMenu;
    CLigCur: TRadioButton;
    CLigPrec: TRadioButton;
    CLigFixe: TRadioButton;
    procedure FormShow(Sender: TObject);
    procedure FTableChange(Sender: TObject);
    procedure BFermerClick(Sender: TObject);
    procedure FListeClick(Sender: TObject);
    procedure FFormatChange(Sender: TObject);
    procedure BValiderClick(Sender: TObject);
    procedure FListeDblClick(Sender: TObject);
    procedure FDebutChange(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure BAideClick(Sender: TObject);
    procedure CLigFixeClick(Sender: TObject);
  private
    procedure ToutEnabled ( G : TGroupBox ; Ena : boolean ) ;
    procedure RenseigneLG ( FF : String ) ;
    procedure ChargeListe ;
    procedure RempliFichier ;
    function  QuelS : String ;
  public
    CurLig : integer ;
    Renvoi : String ;
  end;

implementation

{$R *.DFM}

function ChoixChampZone ( CurLig : integer ; Nature : String ) : String ;
Var X : TFGuidTool ;
BEGIN
Result:='' ;
X:=TFGuidTool.Create(Application) ;
 Try
  X.CurLig:=CurLig ;
  if Nature='GUI' then X.HelpContext := 1430200 ;
  if Nature='LIB' then X.HelpContext := 1425100 ;
  X.ShowModal ;
 Finally
  Result:=X.Renvoi ;
  X.Free ;
 End ;
Screen.Cursor:=SyncrDefault ;
END ;

procedure TFGuidTool.FDebutChange(Sender: TObject);
begin FExempleS.Text:=QuelS ; end;

procedure TFGuidTool.BFermerClick(Sender: TObject);
begin Close ; end;

procedure TFGuidTool.FListeDblClick(Sender: TObject);
begin BValiderClick(Nil) ; end;

procedure TFGuidTool.FFormatChange(Sender: TObject);
begin FExempleD.Text:=FormatDateTime(FFormat.Text,V_PGI.DateEntree) ; end;

procedure TFGuidTool.FTableChange(Sender: TObject);
begin if FTable.Value<>'' then ChargeListe ; end;

procedure TFGuidTool.ChargeListe ;
Var St,pf : String ;
   {$IFDEF CCS3}
   Ctrl : String ;
   {$ENDIF}
    QC : TQuery ;
BEGIN
FCache.Items.Clear ; FCache.Values.Clear ;
FListe.Clear ; if FTable.Value='' then Exit ; if FTable.ItemIndex<0 then Exit ;
PF:=FTable.Values[FTable.ItemIndex] ;
QC:=OpenSQL('Select DH_NOMCHAMP, DH_TYPECHAMP, DH_LIBELLE, DH_CONTROLE, DH_PREFIXE '+
            ' From DECHAMPS Where DH_PREFIXE="'+Pf+'" and DH_CONTROLE like "%G%"',TRUE) ;
While Not QC.EOF do
   BEGIN
   St:=QC.Fields[2].AsString ; if St='' then St:=QC.Fields[0].AsString ;
   {$IFDEF CCS3}
   Ctrl:=QC.Fields[3].AsString ;
   if Pos('#',Ctrl)>0 then BEGIN QC.Next ; Continue ; END ;
   {$ENDIF}
   FListe.Items.Add(St) ;
   FCache.Items.Add(QC.Fields[0].AsString) ;
   FCache.Values.Add(QC.Fields[1].AsString) ;
   QC.Next ;
   END ;
Ferme(QC) ;
FListe.ItemIndex:=0 ; FListe.SetFocus ; FListeClick(Nil) ;
END ;

procedure TFGuidTool.RempliFichier ;
Var QF : TQuery ;
BEGIN
FTable.Items.Clear ; FTable.Values.Clear ;
QF:=OpenSQL('Select DT_NOMTABLE, DT_LIBELLE, DT_PREFIXE from DETABLES '
           +'Where DT_PREFIXE="E" or DT_PREFIXE="G" or DT_PREFIXE="J" or DT_PREFIXE="T" '
           +'order By DT_PREFIXE',True) ;
While Not QF.EOF do
   BEGIN
   FTable.Items.Add(QF.Fields[0].AsString) ; FTable.Values.Add(QF.Fields[2].AsString) ;
   QF.Next ;
   END ;
Ferme(QF) ;
FTable.ItemIndex:=0 ; FTableChange(Nil) ;
END ;

procedure TFGuidTool.FormShow(Sender: TObject);
begin
RempliFichier ; FNumL.Value:=CurLig ;
end;

procedure TFGuidTool.ToutEnabled ( G : TGroupBox ; Ena : boolean ) ;
Var i : integer ;
BEGIN
G.Enabled:=Ena ;
for i:=0 to G.ControlCount-1 do G.Controls[i].Enabled:=Ena ; 
END ;

function TFGuidTool.QuelS : String ;
Var St : String ;
    LMax,Lg,Deb : integer ;
BEGIN
LMax:=FDebut.MaxValue ; Deb:=FDebut.Value ; Lg:=FLongueur.Value ;
if Deb+Lg-1>LMax then Lg:=LMax-Deb+1 ;
St:=Format_String('',LMax) ;
if Deb>1 then FillChar(St[1],Deb-1,'-') ;
FillChar(St[Deb],Lg,'x') ;
if Deb+Lg-1<LMax then FillChar(St[Deb+Lg],LMax-(Deb+Lg)+1,'-') ;
Result:=St ; 
END ;

procedure TFGuidTool.RenseigneLG ( FF : String ) ;
Var L,Pos1 : integer ;
    St     : String ;
BEGIN
if FF='COMBO' then L:=3 else
   BEGIN
   Pos1:=Pos('VARCHAR(',FF) ; if Pos1<=0 then Exit ;
   St:=Trim(Copy(FF,Pos1+8,5)) ; Delete(St,Length(St),1) ; L:=StrToInt(St) ;
   END ;
FDebut.MaxValue:=L ; FDebut.Value:=1 ;
FLongueur.MaxValue:=L ; FLongueur.Value:=L ;
FExempleS.Text:=QuelS ;
END ;

procedure TFGuidTool.FListeClick(Sender: TObject);
Var FF : String ;
    OkD,OkS : Boolean ;
begin
if Fliste.ItemIndex<0 then Exit ;
FChamp.Text:=FCache.Items[FListe.ItemIndex] ;
FF:=FCache.Values[FListe.ItemIndex] ;
OkD:=(FF='DATE') ; OkS:=((FF='COMBO') or (Pos('VARCHAR',FF)>0)) ;
if OkD then
   BEGIN
   ToutEnabled(GDate,True) ; GDate.Visible:=True ; GString.Visible:=False ;
   FFormatChange(Nil) ;
   END else ToutEnabled(GDate,False) ;
if OkS then
   BEGIN
   GDate.Visible:=False ; ToutEnabled(GString,True) ;
   RenseigneLG(FF) ;
   GString.Visible:=True ;
   END else ToutEnabled(GString,False) ;
end;


procedure TFGuidTool.BValiderClick(Sender: TObject);
Var StF : String ;
    NumL : integer ;
begin
StF:='' ;
if GDate.Visible then StF:=FFormat.Text ;
if GString.Visible then
   BEGIN
   if ((FDebut.Value<>1) or (FLongueur.Value<>FLongueur.MaxValue)) then StF:=IntToStr(FDebut.Value)+','+IntToStr(FLongueur.Value) ;
   END ;
if CLigCur.Checked then NumL:=CurLig else
 if CLigPrec.Checked then NumL:=-1 else NumL:=FNumL.Value ;
Renvoi:='[' ;
if StF<>'' then Renvoi:=Renvoi+'"'+StF+'"' ;
Renvoi:=Renvoi+FChamp.Text ;
if ((NumL<>CurLig) and (NumL<>0)) then Renvoi:=Renvoi+':L'+IntToStr(NumL) ;
Renvoi:=Renvoi+']' ;
Close ;
end;

procedure TFGuidTool.FormCreate(Sender: TObject);
begin
PopUpMenu:=ADDMenuPop(PopUpMenu,'','') ;
end;

procedure TFGuidTool.BAideClick(Sender: TObject);
begin
CallHelpTopic(Self) ;
end;

procedure TFGuidTool.CLigFixeClick(Sender: TObject);
begin
{b FP FQ16096 + Les évènements OnClick des composants CLigCur et CLigPrec sont remplacés par CLigFixeClick}
FNumL.Enabled:=CLigFixe.Checked;
H_NumL.Enabled:=CLigFixe.Checked;
{e FP FQ16096}
end;

end.
