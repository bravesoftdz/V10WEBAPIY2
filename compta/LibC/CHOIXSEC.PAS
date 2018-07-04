unit ChoixSec;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, Buttons, Hctrls, ExtCtrls, {$IFNDEF DBXPRESS}dbtables, hmsgbox,
  HSysMenu, Db, Hqry, ComCtrls{$ELSE}uDbxDataSet{$ENDIF}, hmsgbox, HSysMenu, DB, Ent1,Hent1,HCompte,
  ComCtrls, Hqry ;

Type TChoixSect=(osSaisie,osEdition) ;

procedure ChoisirSection(Axe : String3 ; Section : THCpteEdit ; OS : TChoixSect) ;

type
  TFChoixSec = class(TForm)
    PBouton: TPanel;
    Panel1: TPanel;
    BAide: TBitBtn;
    BFerme: TBitBtn;
    BValider: TBitBtn;
    TCPlan: THTable;
    SPlan: TDataSource;
    TCPlanPS_AXE: TStringField;
    TCPlanPS_SOUSSECTION: TStringField;
    TCPlanPS_CODE: TStringField;
    TCPlanPS_LIBELLE: TStringField;
    TCPlanPS_ABREGE: TStringField;
    TCPlanPS_SOCIETE: TStringField;
    TPlan: THTable;
    HMTrad: THSystemMenu;
    MsgBox: THMsgBox;
    TPlanSS_AXE: TStringField;
    TPlanSS_SOUSSECTION: TStringField;
    TPlanSS_LIBELLE: TStringField;
    TPlanSS_CONTROLE: TStringField;
    TPlanSS_DEBUT: TIntegerField;
    TPlanSS_LONGUEUR: TIntegerField;
    TPlanSS_SOCIETE: TStringField;
    ChoixPControl: TPageControl;
    TS1: TTabSheet;
    LC1: TLabel;
    C1: THValComboBox;
    C3: THValComboBox;
    LC3: TLabel;
    C7: THValComboBox;
    LC7: TLabel;
    LC5: TLabel;
    C5: THValComboBox;
    TS2: TTabSheet;
    C11: THValComboBox;
    LC11: TLabel;
    C15: THValComboBox;
    LC15: TLabel;
    LC13: TLabel;
    C13: THValComboBox;
    C12: THValComboBox;
    C14: THValComboBox;
    C16: THValComboBox;
    LC17: TLabel;
    C17: THValComboBox;
    LC12: TLabel;
    LC16: TLabel;
    LC14: TLabel;
    C2: THValComboBox;
    C4: THValComboBox;
    C6: THValComboBox;
    C8: THValComboBox;
    LC2: TLabel;
    LC4: TLabel;
    LC8: TLabel;
    LC6: TLabel;
    C10: THValComboBox;
    LC10: TLabel;
    C9: THValComboBox;
    LC9: TLabel;
    procedure FormShow(Sender: TObject);
    procedure BValiderClick(Sender: TObject);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
  private
    FAxe     : String3 ;
    FSection : String17 ;
    NbPlans,LgSect   : byte ;
    ListeLg  : TStrings ;
    DouViensTu  : TChoixSect ;
    procedure AfficheLesCombos ;
    public
    { Déclarations publiques }
  end;

implementation

const ValeurVide ='?' ;
      JokerVide  ='*' ;
{$R *.DFM}

procedure ChoisirSection(Axe : String3 ; Section : THCpteEdit ; OS : TChoixSect) ;
var
  FChoixSec: TFChoixSec ;
BEGIN
FChoixSec:=TFChoixSec.Create(Application) ;
try
  FChoixSec.FAxe:=Axe ;
  FChoixSec.FSection:=Section.Text ;
  FChoixSec.DouViensTu:=OS ;
  FChoixSec.ShowModal ;
  finally
  if FChoixSec.ModalResult=mrOk then Section.Text:=FChoixSec.FSection ;
  FChoixSec.free ;
  end ;
SourisNormale ;
END ;


procedure TFChoixSec.FormShow(Sender: TObject);
begin
PopUpMenu:=ADDMenuPop(PopUpMenu,'','') ;
TPlan.Open ;
TPlan.SetRange([FAxe],[FAxe]) ;
TCPlan.Open ;
LgSect:=V_PGI.Cpta[AxeToFb(FAxe)].Lg ;
AfficheLesCombos ;
end;

Procedure TFChoixSec.AfficheLesCombos ;
var Y,Deb,Lg : Integer ;
    L     : TLabel ;
    T     : THValComboBox ;
    Code  : String ;
BEGIN
TPlan.First ; NbPlans:=0 ; Y:=100 ;
ListeLg:=TStringList.Create ;
While not TPlan.Eof do
  BEGIN
  Inc(NbPlans) ;
  Deb:=TPlan.FindField('SS_DEBUT').AsInteger ;
  Lg:=TPlan.FindField('SS_LONGUEUR').AsInteger ;
  ListeLg.Add(IntToStr(Lg)) ;
  T:=THValComboBox(FindComponent('C'+IntToStr(NbPlans))) ;
  T.Values.Add('') ;
  if DouViensTu=osSaisie then T.Items.Add(MsgBox.Mess[0])
                         else T.Items.Add(MsgBox.Mess[1]) ;
  T.ItemIndex:=0 ;
  TCPlan.First ;
  while not TCPlan.Eof do
    BEGIN
    Code:=TCPlan.FindField('PS_CODE').AsString ;
    T.Values.Add(Code) ; T.Items.Add(TCPlan.FindField('PS_LIBELLE').AsString) ;
    if (FSection<>'') and (Copy(FSection,Deb,Lg)=Code) then T.Value:=Code ;
    TCPlan.Next ;
    END ;
  T.Visible:=True ;
  L:=TLabel(FindComponent('LC'+IntToStr(NbPlans))) ;
  L.Caption:=TPlan.FindField('SS_LIBELLE').AsString ;
  L.Visible:=True ;
  if (NbPlans<10) then Inc(Y,25) ;
  TPlan.Next ;
  END ;
TS2.TabVisible:=(NbPlans>=10) ;
Height:=Y ;
END ;

procedure TFChoixSec.BValiderClick(Sender: TObject);
var T : THValComboBox ;
    Lg,i,j: integer ;
    S : String ;
    AuMoinsUn : boolean ;
begin
if ListeLg.Count=0 then Exit ;
FSection:='' ;
for i:=1 to NbPlans do
  BEGIN
  T:=THValComboBox(FindComponent('C'+IntToStr(i))) ;
  if T.Visible then
    if T.Value<>'' then FSection:=FSection+T.Value
                   else
                   BEGIN
                   S:='' ;
                   Lg:=StrToInt(ListeLg[i-1]) ;
                   for j:=1 to Lg do S:=S+ValeurVide ;
                   FSection:=FSection+S ;
                   END ;
  END ;
if (DouViensTu=osEdition) and (Copy(FSection,Length(FSection),1)=ValeurVide) then
  BEGIN
  AuMoinsUn:=False ;
  for i:=Length(FSection)-1 downto 1 do
    if (FSection[i]<>ValeurVide) then
      BEGIN
      if i=1 then FSection:=JokerVide
             else FSection:=Copy(FSection,1,i)+JokerVide ;
      AuMoinsUn:=True ;
      Break ;
      END ;
  if not AuMoinsUn then FSection:=JokerVide ;
  END ;
end;

procedure TFChoixSec.FormCloseQuery(Sender: TObject;
  var CanClose: Boolean);
begin ListeLg.Clear ; ListeLg.Free ; end;

end.
