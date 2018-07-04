unit EchgImp_Parametres;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  vierge, HSysMenu, HTB97, Mask, Hctrls, StdCtrls,

HmsgBox,          // PgiInfo, PgiBox


  Ent1,             // VH : ^LaVariableHalley;
  Echg_Code, Grids, ucalendar;

type
  TfEchgImp_Parametres = class(TFVierge)
    GroupBox2: TGroupBox;
    labDate: THLabel;
    labLibelle: THLabel;
    libJournal: THLabel;
    libCompte: THLabel;
    HValCB_Libelle: TEdit;
    CB_AvecDateImport: TCheckBox;
    Hcme_Journal: THCritMaskEdit;
    Hcme_Compte: THCritMaskEdit;
    GroupBox3: TGroupBox;
    cbRAZPrealable: TCheckBox;
    cmeDate: THCritMaskEdit;
    LibLeurre: THLabel;
    HLabel1: THLabel;
    Hcme_Etablissement: THCritMaskEdit;
    procedure BValiderClick(Sender: TObject);
    procedure BFermeClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
  private
    { Déclarations privées }
  public
    { Déclarations publiques }
    Focuzer: TWinControl;
  end;

var
  fEchgImp_Parametres: TfEchgImp_Parametres;

implementation

{$R *.DFM}

procedure TfEchgImp_Parametres.BValiderClick(Sender: TObject);
begin
  if NOT ParametresEnregistre then
    ModalResult := mrNone;
end;

procedure TfEchgImp_Parametres.FormShow(Sender: TObject);
begin
  inherited;
  if Focuzer <> nil then
  begin
    Focuzer.SetFocus;
    Focuzer := nil;
  end
  else
  begin
    if HValCB_Libelle.Enabled then HValCB_Libelle.SetFocus
    else if Hcme_Journal.Enabled then Hcme_Journal .SetFocus
    else if Hcme_Compte.Enabled then Hcme_Compte.SetFocus
    else if cbRAZPrealable.Enabled then cbRAZPrealable.SetFocus
    else if Hcme_Etablissement.Enabled then Hcme_Etablissement.SetFocus
    ;
  end;
end;

procedure TfEchgImp_Parametres.FormCreate(Sender: TObject);
begin
  inherited;
  Focuzer := nil;
end;

procedure TfEchgImp_Parametres.BFermeClick(Sender: TObject);
begin
  if NOT PamametresAnnulation then
    ModalResult := mrNone;
end;

procedure TfEchgImp_Parametres.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
  inherited;
  BFermeClick(Sender);
end;

end.
