unit LettRegu;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Hctrls, StdCtrls, Hcompte, ExtCtrls, Buttons, LettUtil, Ent1, SaisUtil,
  HSysMenu ;

type
  TFLettreRegu = class(TForm)
    BValide: THBitBtn;
    BAbandon: THBitBtn;
    Panel2: TPanel;
    GroupBox1: TGroupBox;
    C1: THCpteEdit;
    C2: THCpteEdit;
    C3: THCpteEdit;
    H3: THLabel;
    H2: THLabel;
    H1: THLabel;
    E_SOLDE: THNumEdit;
    H_Solde: TLabel;
    R: TRadioGroup;
    Label1: TLabel;
    HMTrad: THSystemMenu;
    procedure FormShow(Sender: TObject);
    procedure BValideClick(Sender: TObject);
  private
  public
    RJ    : REGJAL ;
    CpteR,Symbole : String ;
    Debit : boolean ;
    Solde : double ;
    Decim,IndiceRegul : byte ;
  end;

function ChoixCpteRegul ( J : REGJAL ; Debit : boolean ; Solde : double ; Decim : byte ; Symbole : String ; Var Ind : integer ) : String ;

implementation

{$R *.DFM}

uses
  {$IFDEF MODENT1}
  CPProcGen,
  CPProcMetier,
  {$ENDIF MODENT1}
  HEnt1; 


function ChoixCpteRegul ( J : REGJAL ; Debit : boolean ; Solde : double ; Decim : byte ; Symbole : String ; Var Ind : integer ) : String ;
Var X    : TFLettreRegu ;
    i,k  : integer ;
    JJ   : TCR ;
    Diff : boolean ; 
BEGIN
{vérifier solution triviale avant de lancer une fenêtre de choix}
Result:='' ; Ind:=0 ;
if Debit then JJ:=J.D else JJ:=J.C ;
if ((JJ[1]='') and (JJ[2]='') and (JJ[3]='')) then Exit ;
Diff:=False ; for i:=1 to 3 do for k:=i+1 to 3 do if ((JJ[i]<>'') and (JJ[k]<>'') and (JJ[i]<>JJ[k])) then Diff:=True ;
if Not Diff then for i:=1 to 3 do if JJ[i]<>'' then BEGIN Result:=JJ[i] ; Ind:=i ; Exit ; END ;
{fenêtre de choix des comptes de régul}
X:=TFLettreRegu.Create(Application) ;
 Try
  X.RJ:=J ; X.CpteR:='' ; X.Debit:=Debit ; X.Solde:=Solde ; X.Decim:=Decim ; X.Symbole:=Symbole ;
  if X.ShowModal=mrOk then BEGIN Result:=X.CpteR ; Ind:=X.IndiceRegul ; END ;
 Finally
  X.Free ;
 End ;
Screen.Cursor:=SyncrDefault ;
END ;

procedure TFLettreRegu.FormShow(Sender: TObject);
Var i    : integer ;
    Cpte : String17 ;
begin
PopUpMenu:=ADDMenuPop(PopUpMenu,'','') ;
IndiceRegul:=1 ;
for i:=1 to 3 do
    BEGIN
    if Debit then Cpte:=RJ.D[i] else Cpte:=RJ.C[i] ;
    THCpteEdit(FindComponent('C'+Inttostr(i))).Text:=Cpte ;
    TLabel(FindComponent('H'+Inttostr(i))).Caption:='' ;
    END ;
C1.ExisteH ; C2.ExisteH ; C3.ExisteH ;
ChangeMask(E_SOLDE,Decim,Symbole) ;
if Debit then AfficheLeSolde(E_SOLDE,Solde,0) else AfficheLeSolde(E_SOLDE,0,Solde) ;
end;

procedure TFLettreRegu.BValideClick(Sender: TObject);
begin
if R.ItemIndex>=0 then
   BEGIN
   CpteR:=THCpteEdit(FindComponent('C'+Inttostr(R.ItemIndex+1))).Text ;
   IndiceRegul:=R.ItemIndex+1 ;
   END ;
end;

end.
