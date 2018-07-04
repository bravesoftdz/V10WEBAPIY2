unit SaisEnc;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, ExtCtrls, Hcompte, Hctrls, Buttons, HEnt1, Ent1, SaisUtil,
  HSysMenu,uTob ;


Function InfosTvaEnc ( O : TOB ; Action : TActionFiche) : boolean ;

type
  TFSaisEnc = class(TForm)
    POutils: TPanel;
    BValide: THBitBtn;
    BAbandon: THBitBtn;
    BAide: THBitBtn;
    PEntete: TPanel;
    H_REGIMETVA: THLabel;
    E_REGIMETVA: THValComboBox;
    H_GENERAL: THLabel;
    E_GENERAL: THCpteEdit;
    G_LIBELLE: THLabel;
    PTva: TPanel;
    Label1: TLabel;
    RModeTVA: TRadioGroup;
    E_TVA: THValComboBox;
    H_TVA: THLabel;
    Label2: TLabel;
    Panel1: TPanel;
    CPTTVA: THCpteEdit;
    Label3: TLabel;
    H_CPTTVA: THLabel;
    H_TXTVA: THLabel;
    TXTVA: THNumEdit;
    HMTrad: THSystemMenu;
    procedure FormShow(Sender: TObject);
    procedure DeduitTva(Sender: TObject);
    procedure BAideClick(Sender: TObject);
    procedure BAbandonClick(Sender: TObject);
    procedure BValideClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
  private
    Achat : boolean ;
  public
    O   : TOB ;
    Action : TActionFiche ;
  end;


implementation

uses
  {$IFDEF MODENT1}
  CPTypeCons,
  CPProcMetier,
  {$ENDIF MODENT1}
  utilPGI; //XMG 31/07/03

{$R *.DFM}

Function InfosTvaEnc ( O : TOB ; Action : TActionFiche ) : boolean ;
Var X : TFSaisEnc ;
BEGIN
Result:=False ;
X:=TFSaisEnc.Create(Application) ;
 Try
  X.O:=O ; X.Action:=Action ;
  if X.ShowModal=mrOk then Result:=True ; 
 Finally
  X.Free ;
 End ;
END ;


procedure TFSaisEnc.FormShow(Sender: TObject);
Var Exi,NatP : String3 ;
begin
if Action=taConsult then PTva.Enabled:=False ; 
E_REGIMETVA.Value:=O.GetString('E_REGIMETVA') ;
E_GENERAL.Text:=O.GetString('E_GENERAL') ; E_GENERAL.ExisteH ;
E_TVA.Value:=O.GetString('E_TVA') ;
RmodeTVA.Visible:=(VH^.PaysLocalisation<>CodeISOES) ;
Exi:=O.GetString('E_TVAENCAISSEMENT') ;
if (Exi='X') and (VH^.PaysLocalisation<>CodeISOES) then
   RModeTva.ItemIndex:=1
 else
  RModeTva.ItemIndex:=0 ; //XVI 24/02/2005
NatP:=O.GetString('E_NATUREPIECE') ; Achat:=((NatP='FF') or (NatP='AF')) ;
DeduitTva(Nil) ;
end;

procedure TFSaisEnc.DeduitTva(Sender: TObject);
begin
CptTva.Text:='' ; TxTva.Value:=0 ;
if E_TVA.Value='' then Exit ;
if RModeTva.ItemIndex=0 then CptTva.Text:=Tva2Cpte(E_REGIMETVA.Value,E_TVA.Value,Achat)
                        else CptTva.Text:=Tva2Encais(E_REGIMETVA.Value,E_TVA.Value,Achat) ;
TxTva.Value:=Tva2Taux(E_REGIMETVA.Value,E_TVA.Value,Achat) ;
end;

procedure TFSaisEnc.BAideClick(Sender: TObject);
begin
CallHelpTopic(Self) ;
end;

procedure TFSaisEnc.BAbandonClick(Sender: TObject);
begin
Close ;
end;

procedure TFSaisEnc.BValideClick(Sender: TObject);
begin
if Action=taConsult then ModalResult:=mrCancel else
   BEGIN
   O.PutValue('E_TVA',E_TVA.Value) ;
   if RModeTva.ItemIndex=0
     then O.PutValue('E_TVAENCAISSEMENT','-')
     else O.PutValue('E_TVAENCAISSEMENT','X') ;
   ModalResult:=mrOk ;
   END ;
end;

procedure TFSaisEnc.FormCreate(Sender: TObject);
begin
PopUpMenu:=ADDMenuPop(PopUpMenu,'','') ; 
end;

end.
