unit RepImmo;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, Buttons, ExtCtrls, ImEnt, HEnt1, Hctrls, hmsgbox, HSysMenu, SoldeCpt, HPanel, UiUtil,
  HTB97 ;

type
  TFRepImmo = class(TForm)
    Panel1: TPanel;
    FControleMvt: TGroupBox;
    HPB: TToolWindow97;
    BAide: TToolbarButton97;
    BValider: TToolbarButton97;
    BFerme: TToolbarButton97;
    FOp: TCheckBox;
    MsgRien: THMsgBox;
    FcontroleCpt: TGroupBox;
    FImmos: TCheckBox;
    FCptes: TCheckBox;
    Mes: THMsgBox;
    HMTrad: THSystemMenu;
    MsgBar: THMsgBox;
    FEche: TCheckBox;
    FAmort: TCheckBox;
    TTravail: TLabel;
    TTravailLeq: TLabel;
    Dock: TDock97;
    procedure FormShow(Sender: TObject);
    procedure BValiderClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure BAideClick(Sender: TObject);
  private
    { Déclarations privées }
    Immos, Associes, Plans, Operations, Echeanciers : boolean ;
    Procedure LanceControl ;
    Procedure AvecQui ;
    Function  Selection : Boolean ;
    Procedure MajCaption( M : byte ) ;
  public
    { Déclarations publiques }
  end;

procedure RepareFichiersImmo ;

implementation

uses VerImmo, UtilPgi ;

{$R *.DFM}

procedure RepareFichiersImmo ;
var FRepImmo : TFRepImmo ; PP : THPanel ;
BEGIN
if Not _BlocageMonoPoste(True) then Exit ;
FRepImmo:=TFRepImmo.Create(Application) ;
PP:=FindInsidePanel ;
if PP=nil then
   BEGIN
    try
     FRepImmo.ShowModal ;
    finally
     FRepImmo.Free ;
     _DeblocageMonoPoste(True) ;
    end ;
   SourisNormale ;
   END else
   BEGIN
   InitInside(FRepImmo,PP) ;
   FRepImmo.Show ;
   END ;
END ;

procedure TFRepImmo.FormShow(Sender: TObject);
begin
PopUpMenu:=ADDMenuPop(PopUpMenu,'','') ;
TTravail.Caption:='' ; TTravailLeq.Caption:='' ;
end;

procedure TFRepImmo.BValiderClick(Sender: TObject);
begin
if not Selection then begin MsgRien.Execute(0,'','') ; Exit ; end ;
if MsgRien.Execute(2,'','')<>mryes then exit ;
TTravail.Caption:=Mes.Mess[4] ;
AvecQui ;
LanceControl ;
TTravail.Caption:='' ;
TTravailLeq.Caption:='' ;
end;

Procedure TFRepImmo.AvecQui ;
BEGIN
Immos:=FImmos.Checked ;
Associes:=FCptes.Checked ;
Plans:=FAmort.Checked ;
Operations:=FOp.Checked ;
Echeanciers:=FEche.Checked ;
END ;

Procedure TFRepImmo.LanceControl ;
BEGIN
EnableControls(Self,False) ;
if Immos       then begin MajCaption(1) ; ExecuteVerifImmos(tvImmos, TRUE) ; end ;
if Associes    then begin MajCaption(2) ; ExecuteVerifImmos(tvCptes, TRUE) ; end ;
if Plans       then begin MajCaption(3) ; ExecuteVerifImmos(tvAmort, TRUE) ; end ;
if Operations  then begin MajCaption(4) ; ExecuteVerifImmos(tvOp,    TRUE) ; end ;
if Echeanciers then begin MajCaption(5) ; ExecuteVerifImmos(tvEch,   TRUE) ; end ;
EnableControls(Self,True) ;
Screen.Cursor:=SyncrDefault ;
MsgRien.Execute(1,'','') ;
END ;

Function TFRepImmo.Selection : Boolean ;
BEGIN
Result:=FImmos.Checked or FCptes.Checked or FAmort.Checked or FOp.Checked or FEche.Checked ;
END ;

Procedure TFRepImmo.MajCaption( M : byte ) ;
BEGIN
TTravail.Caption:=MsgBar.Mess[0]+' ' ;
TTravailLeq.Caption:=MsgBar.Mess[M] ;
Application.ProcessMessages ;
END ;

procedure TFRepImmo.FormClose(Sender: TObject; var Action: TCloseAction);
begin
if Parent is THPanel then
   BEGIN
   _DeblocageMonoPoste(True) ;
   Action:=caFree ;
   END ;
end;

procedure TFRepImmo.BAideClick(Sender: TObject);
begin
CallHelpTopic(Self) ;
end;

end.
