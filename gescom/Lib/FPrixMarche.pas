unit FPrixMarche;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, Hctrls, HTB97, ExtCtrls, HPanel,HEnt1;

function GestionPrixMarche (var MontantHt: double;var MontantHtPar : double; var ModeCalcul : integer; EnHt : boolean; var EtenduApplication: integer) : boolean;

type
  TPrixMarche = class(TForm)
    PPanBase: THPanel;
    Dock971: TDock97;
    ToolWindow971: TToolWindow97;
    BAide: TToolbarButton97;
    BAbandon: TToolbarButton97;
    BValider: TToolbarButton97;
    LMontantHt: THLabel;
    TMontantHt: THNumEdit;
    HNumEdit1: THNumEdit;
    GroupBox1: TGroupBox;
    RMontant: TRadioButton;
    RMarge: TRadioButton;
    HPanel1: THPanel;
    G_APPLICCALC: TGroupBox;
    RDOC: TRadioButton;
    RPARAG: TRadioButton;
    procedure BValiderClick(Sender: TObject);
    procedure BAbandonClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure RPARAGClick(Sender: TObject);
    procedure RDOCClick(Sender: TObject);
  private
    EnHt : boolean;
    PriseEncompteModif : boolean;
    MontanthtDoc : double;
    MontantHtPar : double;
    Etendu : integer;
    { Déclarations privées }
  public
    { Déclarations publiques }
  end;

var
  PrixMarche: TPrixMarche;

implementation

{$R *.DFM}

function GestionPrixMarche (var MontantHt: double;var MontantHtPar : double; var ModeCalcul : integer; EnHt : boolean; var EtenduApplication : integer) : boolean;
var PrixMarche: TPrixMarche;
begin
PrixMarche := TprixMarche.Create (application);
PrixMarche.MontantHtDoc := MontantHt;
PrixMarche.MontantHtPar := MontantHtPar;
PrixMarche.Etendu := EtenduApplication;
PrixMarche.EnHt := EnHt;
PrixMarche.PriseEncompteModif := true;
Try
PrixMarche.showModal;
result := PrixMarche.PriseEncompteModif;
if result then
   begin
   EtenduApplication := PrixMarche.Etendu;
   if EtenduApplication = 0 then MontantHt := PrixMarche.TMontantHt.Value
      else MontantHtPar := PrixMarche.TMontantHt.Value;
   if PrixMarche.RMontant.Checked then ModeCalcul := 1 else ModeCalcul := 2;
   end;
finally
PrixMarche.free;
end;
end;


procedure TPrixMarche.BValiderClick(Sender: TObject);
Begin
nextcontrol(self,true);
priseencompteModif := true;
close;
end;

procedure TPrixMarche.BAbandonClick(Sender: TObject);
begin
priseencompteModif := false;
close;
end;

procedure TPrixMarche.FormShow(Sender: TObject);
begin
if (Etendu = 0) or (MontantHtPar = 0) then
   begin
   RDOC.Checked := true;
   TMontantHt.Value := MontantHtDoc;
   if EnHt then LMontantHt.Caption := 'Montant H.T du document'
      else LMontantHt.Caption := 'Montant T.T.C du document';
   if MontantHtPar = 0 then G_APPLICCALC.Enabled := false ;
   end else
   begin
   RPARAG.Checked := true;
   TMontantHt.Value := MontantHtPar;
   if EnHt then LMontantHt.Caption := 'Montant H.T du paragraphe'
      else LMontantHt.Caption := 'Montant T.T.C du paragraphe';
   end;
end;

procedure TPrixMarche.RPARAGClick(Sender: TObject);
begin
TMontantHt.Value := MontantHtPar;
Etendu := 1;
if EnHt then LMontantHt.Caption := 'Montant H.T du paragraphe'
   else LMontantHt.Caption := 'Montant T.T.C du paragraphe';
end;

procedure TPrixMarche.RDOCClick(Sender: TObject);
begin
TMontantHt.Value := MontantHtDoc;
Etendu := 0;
if EnHt then LMontantHt.Caption := 'Montant H.T du document'
   else LMontantHt.Caption := 'Montant T.T.C du document';
end;

end.
