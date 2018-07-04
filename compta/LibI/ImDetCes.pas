unit ImDetCes;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, Hctrls, ComCtrls, HSysMenu, hmsgbox, HTB97, ExtCtrls, HPanel,
  dbtables,Ent1, HEnt1;

type
  TFDetailCession = class(TForm)
    HPanel3: THPanel;
    bValider: TToolbarButton97;
    bAnnuler: TToolbarButton97;
    bAide: TToolbarButton97;
    HM: THMsgBox;
    HMTrad: THSystemMenu;
    PageControl1: TPageControl;
    DEconomique: TTabSheet;
    DFiscal: TTabSheet;
    HPanel1: THPanel;
    rbRemplacerEco: TRadioButton;
    rbAjouterEco: TRadioButton;
    HLabel4: THLabel;
    DotationEco: THNumEdit;
    DotPartieCedeeEco: THNumEdit;
    ExcepEco: THNumEdit;
    HLabel5: THLabel;
    HLabel2: THLabel;
    HPanel2: THPanel;
    HLabel1: THLabel;
    VNCEco: THNumEdit;
    HPanel4: THPanel;
    HLabel3: THLabel;
    VNCFisc: THNumEdit;
    HPanel5: THPanel;
    HLabel6: THLabel;
    HLabel7: THLabel;
    HLabel8: THLabel;
    rbRemplacerFisc: TRadioButton;
    rbAjouterFisc: TRadioButton;
    DotationFisc: THNumEdit;
    DotPartieCedeeFisc: THNumEdit;
    ExcepFisc: THNumEdit;
    procedure bValiderClick(Sender: TObject);
    procedure bAnnulerClick(Sender: TObject);
    procedure RecalculDotationFisc(Sender: TObject);
    procedure RecalculDotationEco(Sender: TObject);
    procedure FormShow(Sender: TObject);
  private
    { Déclarations privées }
    fDotationEcoTheo : double;
    fDotationFiscTheo : double;
    fCodeImmo : string;
    fDateOpe : TDateTime;
  public
    { Déclarations publiques }
  end;

function ExecuteDetailCession (CodeImmo : string; DateOpe : TDateTime): boolean;

implementation

uses PlanAmor, Outils;

{$R *.DFM}

function ExecuteDetailCession (CodeImmo : string ; DateOpe : TDateTime) : boolean;
var  FDetailCession: TFDetailCession;
begin
  FDetailCession := TFDetailCession.Create(Application) ;
  FDetailCession.fCodeImmo := CodeImmo;
  FDetailCession.fDateOpe := DateOpe;  
  try
    if FDetailCession.ShowModal = mrYes then Result := True
    else Result := False;
  finally
    FDetailCession.Free ;
  end ;
end ;

procedure TFDetailCession.bValiderClick(Sender: TObject);
begin
  ModalResult := mrYes;
end;

procedure TFDetailCession.bAnnulerClick(Sender: TObject);
begin
  if HM.Execute(0,Caption,'')=mrYes then ModalResult := mrNo
  else ModalResult := mrYes;
end;

procedure TFDetailCession.RecalculDotationFisc(Sender: TObject);
begin
  if rbAjouterFisc.Checked  then DotationFisc.Value := fDotationFiscTheo + DotPartieCedeeFisc.Value
  else DotationFisc.Value := DotPartieCedeeFisc.Value;
  ExcepFisc.Value := DotationFisc.Value - fDotationFiscTheo;
end;

procedure TFDetailCession.RecalculDotationEco(Sender: TObject);
begin
  if rbAjouterEco.Checked  then DotationEco.Value := fDotationEcoTheo + DotPartieCedeeEco.Value
  else DotationEco.Value := DotPartieCedeeEco.Value;
  ExcepEco.Value := DotationEco.Value - fDotationEcoTheo;
end;

procedure TFDetailCession.FormShow(Sender: TObject);
var QPlan : TQuery;
    Plan : TPlanAmort;
    MontantDotation: double;
    ProRata : double;
    NbMoisPeriode : integer;
begin
  QPlan:=OpenSQL('SELECT * FROM IMMO WHERE I_IMMO="'+fCodeImmo+'"', FALSE) ;
  Plan := TPlanAmort.Create;
  RecuperePlanAmortissement(Plan,QPlan);
  DFiscal.TabVisible := Plan.Fiscal;
  MontantDotation := RecupereDotationExercice(VH^.Encours.Fin,Plan.NbCotations,Plan.TableauDate,Plan.AmortEco);
  CalculProrataLin(VH^.Encours.Deb,VH^.Encours.Fin,FDateOpe,VH^.Encours.Deb,Prorata,NbMoisPeriode);
  fDotationEcoTheo  := Arrondi (MontantDotation*Prorata,V^.OkDecV);
  DotPartieCedeeEco.Value := fDotationEcoTheo;
  VNCEco.Value := RecupereVNCEco(QPlan.FindField('I_MONTANTHT').AsFloat,Plan, fDateOpe);
  DotationEco.Value := fDotationEcoTheo;
  ExcepEco.Value := 0.0;
  if Plan.Fiscal then
  begin
    fDotationFiscTheo := RecupereDotationExercice(fDateOpe,Plan.NbCotations,Plan.TableauDate,Plan.AmortFisc);
    DotPartieCedeeFisc.Value := fDotationFiscTheo;
    VNCFisc.Value := RecupereVNCFisc(QPlan.FindField('I_MONTANTHT').AsFloat,Plan, fDateOpe);
    DotationFisc.Value := fDotationFiscTheo;
    ExcepFisc.Value := 0.0;
  end;
  DetruitPlanAmortissement(Plan);
  Ferme(QPlan);
end;

end.
