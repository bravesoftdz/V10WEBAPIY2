unit EchgImp_Assistant;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Vierge, StdCtrls, HFolders, Menus, Grids, Hctrls, Mask, ExtCtrls, HPanel,

  UIUtil,Ent1,             // VH : ^LaVariableHalley;

  HSysMenu, HTB97, hent1, hmsgbox, EchgImp_Parametres, CheckLst;

type
  TFEchgImp_Assistant = class(TFVierge)
    btnParametres: TToolbarButton97;
    btnTables: TToolbarButton97;
    PopTables: TPopupMenu;
    mConvCptes: TMenuItem;
    mConvJnx: TMenuItem;
    bfKit: TBrowseFolder;
    BRapport: TToolbarButton97;
    BKit: TToolbarButton97;
    PANELREGROUPEMENT: THPanel;
    HPanel1: THPanel;
    hgVariables: THGrid;
    panAuto: TPanel;
    Panel3: TPanel;
    panDateBal: TPanel;
    HLabDate: THLabel;
    cmeDate: THCritMaskEdit;
    HLABEL2: THLabel;
    Hcme_Solutions: THCritMaskEdit;
    HLabel1: THLabel;
    cmeNomFich: THCritMaskEdit;
    btnSelectScriptsReco: TToolbarButton97;
    btnModeOp: TToolbarButton97;
    btnVisuFciDonnees: TToolbarButton97;
    HLabel3: THLabel;
    Hcme_Corbeille: THCritMaskEdit;
    HLabel4: THLabel;
    Hcme_Masque: THCritMaskEdit;
    bValider2: TToolbarButton97;
    procedure FormShow(Sender: TObject);
    procedure BValiderClick(Sender: TObject);
    procedure btnParametresClick(Sender: TObject);
    procedure btnModeOpClick(Sender: TObject);
    procedure btnVisuFciDonneesClick(Sender: TObject);
    procedure Hcme_SolutionsElipsisClick(Sender: TObject);
    procedure mTablesConversionClick(Sender: TObject);
    procedure FormConstrainedResize(Sender: TObject; var MinWidth,
      MinHeight, MaxWidth, MaxHeight: Integer);
    procedure Hcme_SolutionsChange(Sender: TObject);
    procedure cmeNomFichChange(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure BRapportClick(Sender: TObject);
    procedure BKitClick(Sender: TObject);
  private
    procedure AttrapeMessage(var Msg: TMsg; var Handled: Boolean);
    { Déclarations privées }
  public
    { Déclarations publiques }
  end;

var
  fEchgImp_Assistant: TFEchgImp_Assistant;

implementation

uses Echg_code;

{$R *.DFM}

procedure TFEchgImp_Assistant.AttrapeMessage (var Msg: TMsg; var Handled: Boolean);
begin
  TraiteMessage (Msg, Handled);
end;

procedure TFEchgImp_Assistant.FormShow(Sender: TObject);
begin
  inherited;
  if not InitialiseLeBiniou then
  begin
    Close;
    if IsInside (Self) then CloseInsidePanel (Self);
  end;
  Application.OnMessage := AttrapeMessage;
end;

procedure TFEchgImp_Assistant.BValiderClick(Sender: TObject);
begin
  inherited;
  BoutonValiderClick;
end;

procedure TFEchgImp_Assistant.btnParametresClick(Sender: TObject);
begin
  inherited;
  ParametresAffichage;
end;

procedure TFEchgImp_Assistant.btnModeOpClick(Sender: TObject);
begin
  inherited;
  VisuModeOperatoire (Hcme_Solutions.Plus);
end;

procedure TFEchgImp_Assistant.btnVisuFciDonneesClick(Sender: TObject);
begin
  inherited;
  VisuFicDonnees (cmeNomFich.Text);
end;

procedure TFEchgImp_Assistant.Hcme_SolutionsElipsisClick(Sender: TObject);
begin
  inherited;
  LookUpA (Hcme_Solutions);
end;

procedure TFEchgImp_Assistant.mTablesConversionClick(Sender: TObject);
begin
  inherited;
  TraitePopUp (Sender, Hcme_Solutions.Plus, cmeNomFich.Text);
end;

procedure TFEchgImp_Assistant.FormConstrainedResize(Sender: TObject;
  var MinWidth, MinHeight, MaxWidth, MaxHeight: Integer);
begin
  inherited;
  FormateGrille (hgVariables);
end;

procedure TFEchgImp_Assistant.Hcme_SolutionsChange(Sender: TObject);
begin
  inherited;
  TrtValueChange (Sender);
end;

procedure TFEchgImp_Assistant.cmeNomFichChange(Sender: TObject);
begin
  inherited;
  TrtValueChange (Sender);
end;

procedure TFEchgImp_Assistant.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
  inherited;
  if fEchgImp_Parametres <> nil then
    fEchgImp_Parametres.Free;
  if Parent is THPanel then action := caFree;
end;

procedure TFEchgImp_Assistant.BRapportClick(Sender: TObject);
begin
  inherited;
  CompteRenduComAffichage;
end;

procedure TFEchgImp_Assistant.BKitClick(Sender: TObject);
begin
  inherited;
  Kit;
end;

end.


