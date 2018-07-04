unit BTASSIST_MAJ_TARIF;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  assist, ComCtrls, HSysMenu, hmsgbox, StdCtrls, ExtCtrls, HPanel, HTB97,
  Hctrls,UTOB,HENT1;

type
  TBTASSIST_MAJ_TARIF = class(TFAssist)
    TSDEPART: TTabSheet;
    RBMAJAPA: TRadioButton;
    RBMAJPV: TRadioButton;
    TSACHAT: TTabSheet;
    TSVENTE: TTabSheet;
    POURCENTACH: THNumEdit;
    HLabel1: THLabel;
    AUGMENTACH: TRadioButton;
    DIMINUACH: TRadioButton;
    SURPVOK: TCheckBox;
    Label1: TLabel;
    Label2: TLabel;
    HLabel2: THLabel;
    POURCENTVTE: THNumEdit;
    AUGMENTVTE: TRadioButton;
    DIMINUVTE: TRadioButton;
    TSRECAP: TTabSheet;
    PanelFin: TPanel;
    TTextFin1: THLabel;
    TTextFin2: THLabel;
    TRecap: THLabel;
    ListRecap: TListBox;
    procedure bSuivantClick(Sender: TObject);
    procedure bPrecedentClick(Sender: TObject);
    procedure RBMAJPVClick(Sender: TObject);
    procedure RBMAJAPAClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure bFinClick(Sender: TObject);
  private
    { Déclarations privées }
    TOBART : TOB;
    procedure definiRecap;
    procedure ChangeTarifs;
    procedure AppliqueChangeTarif (TOBA : TOB);
    procedure MajPv (TOBA : TOB;Pourcent : double; Augment : boolean);
    procedure MajPA (TOBA : TOB;Pourcent : double; Augment,SurPV : boolean);
    function ControleSaisie : boolean;
  public
    { Déclarations publiques }
  end;

procedure BTAssist_MajTarifArt (TOBArticle : TOB);

implementation

{$R *.DFM}
procedure BTAssist_MajTarifArt (TOBArticle : TOB);
var
   Fo_Assist : TBTASSIST_MAJ_TARIF;
Begin
  Fo_Assist := TBTASSIST_MAJ_TARIF.Create (Application);
  Fo_Assist.TOBArt:= TOBArticle;
  Try
     Fo_Assist.ShowModal;
  Finally
     Fo_Assist.free;
  End;
end;

procedure TBTASSIST_MAJ_TARIF.bSuivantClick(Sender: TObject);
begin
  inherited;
  if P.ActivePageIndex = P.PageCount-1 then
  begin
  	bfin.Enabled := true;
  	bSuivant.Enabled := false;
  end else if P.Pages[P.ActivePageIndex+1].tag = -1  then
  begin
  	bfin.Enabled := true;
  	bSuivant.Enabled := false;
  end else
  begin
  	bfin.Enabled := false;
  end;
  definiRecap;
end;

procedure TBTASSIST_MAJ_TARIF.bPrecedentClick(Sender: TObject);
begin
  inherited;
  if P.ActivePageIndex < P.PageCount-1 then
  begin
  	bfin.Enabled := false;
  end;
end;

procedure TBTASSIST_MAJ_TARIF.RBMAJPVClick(Sender: TObject);
begin
  inherited;
  TSACHAT.PageIndex := P.PageCount-1;
  TSACHAT.Tag := -1;
  TSACHAT.TabVisible := false;
  TSVENTE.PageIndex := 1;
  TSVENTE.Tag := 0;
  TSVENTE.TabVisible := True;
end;

procedure TBTASSIST_MAJ_TARIF.RBMAJAPAClick(Sender: TObject);
begin
  inherited;
  TSVENTE.PageIndex := P.PageCount-1;
  TSVENTE.Tag := -1;
  TSVENTE.TabVisible := false;
  TSACHAT.PageIndex := 1;
  TSACHAT.Tag := 0;
  TSACHAT.TabVisible := True;
end;


procedure TBTASSIST_MAJ_TARIF.FormShow(Sender: TObject);
begin
  inherited;
  RBMAJPVClick (self);
end;

procedure TBTASSIST_MAJ_TARIF.bFinClick(Sender: TObject);
begin
  inherited;
  if not ControleSaisie then exit;
  ChangeTarifs;
  close;
end;

procedure TBTASSIST_MAJ_TARIF.definiRecap;
begin
  ListRecap.Clear;

  if RBMAJPV.checked then
  begin
    ListRecap.Items.add ('Mise à jour des prix de ventes');
    if AUGMENTVTE.Checked then
    begin
      ListRecap.Items.add ('Augmentation de '+FloatToStr(POURCENTVTE.Value)+' %');
    end else
    begin
      ListRecap.Items.add ('Diminution de '+FloatToStr(POURCENTVTE.Value)+' %');
    end;
  end else
  begin
    ListRecap.Items.add ('Mise à jour des prix d''achats');
    if AUGMENTACH.Checked then
    begin
      ListRecap.Items.add ('Augmentation de '+FloatToStr(POURCENTACH.Value)+' %');
    end else
    begin
      ListRecap.Items.add ('Diminution de '+FloatToStr(POURCENTACH.Value)+' %');
    end;
    if SURPVOK.Checked then  ListRecap.Items.add ('Avec répercution sur le Prix de vente')
                       else ListRecap.Items.add ('Sans répercution sur le Prix de vente');
  end;
end;

procedure TBTASSIST_MAJ_TARIF.ChangeTarifs;
var indice : integer;
    TOBA : TOB;
begin
  for Indice := 0 to TOBART.detail.count -1 do
  begin
    TOBA := TOBART.detail[Indice];
    AppliqueChangeTarif (TOBA);
  end;
end;

procedure TBTASSIST_MAJ_TARIF.AppliqueChangeTarif(TOBA : TOB);
begin
  if RBMAJPV.checked then
  begin
    MajPv (TOBA,POURCENTVTE.Value,AUGMENTVTE.Checked)
  end else
  begin
    MajPA (TOBA,POURCENTACH.Value,AUGMENTACH.Checked,SURPVOK.Checked);
  end;
end;


procedure TBTASSIST_MAJ_TARIF.MajPA(TOBA: TOB; Pourcent: double; Augment, SurPV: boolean);
var valeur : double;
begin
  Valeur := Arrondi(TOBA.GetValue('GA_PAHT')*(POurcent/100),V_PGI.OKdecP);
  if Augment then
  begin
    TOBA.putValue('GA_PAHT',TOBA.GetValue('GA_PAHT')+Valeur);
  end else
  begin
    TOBA.putValue('GA_PAHT',TOBA.GetValue('GA_PAHT')-Valeur);
  end;
  if TOBA.GetValue('GA_COEFFG')=0 then TOBA.PutValue('GA_COEFFG',1);
  TOBA.PutValue ('GA_DPR', Arrondi(TOBA.GetValue('GA_PAHT')*TOBA.GetValue('GA_COEFFG'),V_PGI.okDECV));
  if SurPV then
  begin
    if TOBA.GetValue('GA_COEFCALCHT')=0 then TOBA.PutValue('GA_COEFCALCHT',1);
    TOBA.PutValue ('GA_PVHT', Arrondi(TOBA.GetValue('GA_DPR')*TOBA.GetValue('GA_COEFCALCHT'),V_PGI.okDECV));
  end else
  begin
    if TOBA.GetValue('GA_DPR') <> 0 then
    begin
      TOBA.PutValue ('GA_COEFCALCHT', Arrondi(TOBA.GetValue('GA_PVHT')/TOBA.GetValue('GA_DPR'),4));
    end;
  end;
  TOBA.UpdateDB;
end;

procedure TBTASSIST_MAJ_TARIF.MajPv(TOBA: TOB; Pourcent: double; Augment: boolean);
var Valeur : double;
begin
  Valeur := Arrondi(TOBA.GetValue('GA_PVHT')*(POurcent/100),V_PGI.OKdecP);
  if Augment then
  begin
    TOBA.putValue('GA_PVHT',TOBA.GetValue('GA_PVHT')+Valeur);
  end else
  begin
    TOBA.putValue('GA_PVHT',TOBA.GetValue('GA_PVHT')-Valeur);
  end;
  if TOBA.GetValue('GA_DPR') <> 0 then
  begin
    TOBA.PutValue ('GA_COEFCALCHT', Arrondi(TOBA.GetValue('GA_PVHT')/TOBA.GetValue('GA_DPR'),4));
  end;
  TOBA.UpdateDB;
end;

function TBTASSIST_MAJ_TARIF.ControleSaisie: boolean;
begin
  result := true;
  if RBMAJPV.checked then
  begin
    if POURCENTVTE.Value <= 0 then
    begin
      PgiBox ('Vous devez indiquer un pourcentage');
      Result:=false;
    end;
  end else
  begin
    if POURCENTACH.Value <= 0 then
    begin
      PgiBox ('Vous devez indiquer un pourcentage');
      Result:=false;
    end;
  end;
end;

end.

