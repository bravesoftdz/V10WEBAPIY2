unit Expecc;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs, Filtre,
  StdCtrls, Hcompte, HRegCpte, Hctrls, ComCtrls, hmsgbox, Menus, HSysMenu,
  HTB97, Ent1, HEnt1, Mask,HPanel, UiUtil, dbTables, ExtCtrls ;

Procedure ExportEnCours ;

type
  TFExpECC1 = class(TForm)
    Dock971: TDock97;
    PFiltres: TToolWindow97;
    FFiltres: THValComboBox;
    HPB: TToolWindow97;
    BValider: TToolbarButton97;
    BFerme: TToolbarButton97;
    BAide: TToolbarButton97;
    HMTrad: THSystemMenu;
    POPF: TPopupMenu;
    BCreerFiltre: TMenuItem;
    BSaveFiltre: TMenuItem;
    BDelFiltre: TMenuItem;
    BRenFiltre: TMenuItem;
    BNouvRech: TMenuItem;
    Msg: THMsgBox;
    Sauve: TSaveDialog;
    Pages: TPageControl;
    TabSheet1: TTabSheet;
    BFiltre: TToolbarButton97;
    PLibres: TTabSheet;
    TabSheet2: TTabSheet;
    Panel1: TPanel;
    TCPTEDEBUT: THLabel;
    TCPTEFIN: THLabel;
    LFile: THLabel;
    RechFile: TToolbarButton97;
    LFDateButoir: THLabel;
    Label1: TLabel;
    HLabel3: THLabel;
    HDateNeg1: THLabel;
    HDateNeg2: THLabel;
    CPTEDEBUT: THCpteEdit;
    CPTEFIN: THCpteEdit;
    FileName: TEdit;
    FDateButoir: THCritMaskEdit;
    Format: THValComboBox;
    Masque: TEdit;
    DateNeg1: THCritMaskEdit;
    DateNeg2: THCritMaskEdit;
    CBExportDetail: TCheckBox;
    Panel2: TPanel;
    HLabel2: THLabel;
    ST413EAR: THLabel;
    ST413EEP: THLabel;
    T411: TEdit;
    T413EAR: TEdit;
    T413EEP: TEdit;
    FMethodeECTNE: TRadioGroup;
    FMethodeECTENR: TRadioGroup;
    Panel3: TPanel;
    TT_TABLE0: THLabel;
    TT_TABLE5: THLabel;
    TT_TABLE1: THLabel;
    TT_TABLE6: THLabel;
    TT_TABLE2: THLabel;
    TT_TABLE7: THLabel;
    TT_TABLE3: THLabel;
    TT_TABLE8: THLabel;
    TT_TABLE4: THLabel;
    TT_TABLE9: THLabel;
    T_TABLE0: THCpteEdit;
    T_TABLE5: THCpteEdit;
    T_TABLE1: THCpteEdit;
    T_TABLE6: THCpteEdit;
    T_TABLE2: THCpteEdit;
    T_TABLE7: THCpteEdit;
    T_TABLE3: THCpteEdit;
    T_TABLE8: THCpteEdit;
    T_TABLE4: THCpteEdit;
    T_TABLE9: THCpteEdit;
    OKPL: TCheckBox;
    OKTL: TCheckBox;
    SFDateEcr1: THLabel;
    FDateEcr1: THCritMaskEdit;
    SFDateEcr2: THLabel;
    FDateEcr2: THCritMaskEdit;
    SFDateModif1: THLabel;
    FDateModif1: THCritMaskEdit;
    SFDateModif2: THLabel;
    FDateModif2: THCritMaskEdit;
    TTCPTSOLDE: THLabel;
    TCPTSOLDE: TEdit;
    OkAutreMP: TCheckBox;
    procedure FormShow(Sender: TObject);
    procedure BCreerFiltreClick(Sender: TObject);
    procedure BSaveFiltreClick(Sender: TObject);
    procedure BDelFiltreClick(Sender: TObject);
    procedure BRenFiltreClick(Sender: TObject);
    procedure BNouvRechClick(Sender: TObject);
    procedure RechFileClick(Sender: TObject);
    procedure BValiderClick(Sender: TObject);
    procedure FFiltresChange(Sender: TObject);
    procedure POPFPopup(Sender: TObject);
    procedure FormatChange(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure CBExportDetailClick(Sender: TObject);
  private
    { Déclarations privées }
    NomFiltre : String ;
  public
    { Déclarations publiques }
  end;

implementation

{$R *.DFM}

Procedure ExportEnCours ;
var FExpEcc:TFExpEcc1 ;
    PP : THPanel ;
BEGIN
FExpEcc:=TFExpEcc1.Create(Application) ;
PP:=FindInsidePanel ;
if PP=Nil then
   BEGIN
   try
     FExpEcc.ShowModal ;
     finally
     FExpEcc.Free ;
     END ;
   SourisNormale ;
   END else
   BEGIN
   InitInside(FExpEcc,PP) ;
   FExpEcc.Show ;
   END ;
END ;

procedure DirDefault(Sauve : TSaveDialog ; FileName : String) ;
var j,i : integer ;
BEGIN
j:=Length(FileName);
for i:=Length(FileName) downto 1 do if FileName[i]='\' then BEGIN j:=i ; Break ; END ;
Sauve.InitialDir:=Copy(FileName,1,j) ;
END ;

procedure TFExpECC1.FormShow(Sender: TObject);
begin
Pages.ActivePage:=Pages.Pages[0] ;
NomFiltre:='TEST' ;
FDateButoir.Text:=DateToStr(V_PGI.DateEntree) ;
DateNeg1.Text:=DateToStr(VH^.Entree.Deb) ; DateNeg2.Text:=DateToStr(VH^.Entree.Fin) ;
Format.ItemIndex:=0 ;
ChargeFiltre(NomFiltre,FFiltres,Pages) ;
end;

procedure TFExpECC1.BCreerFiltreClick(Sender: TObject);
begin
NewFiltre(NomFiltre,FFiltres,Pages) ;
end;

procedure TFExpECC1.BSaveFiltreClick(Sender: TObject);
begin
SaveFiltre(NomFiltre,FFiltres,Pages) ;
end;

procedure TFExpECC1.BDelFiltreClick(Sender: TObject);
begin
DeleteFiltre(NomFiltre,FFiltres) ;
end;

procedure TFExpECC1.BRenFiltreClick(Sender: TObject);
begin
RenameFiltre(NomFiltre,FFiltres) ;
end;

procedure TFExpECC1.BNouvRechClick(Sender: TObject);
begin
VideFiltre(FFiltres,Pages) ;
end;

procedure TFExpECC1.RechFileClick(Sender: TObject);
begin
DirDefault(Sauve,FileName.Text) ;
if Sauve.Execute then FileName.Text:=Sauve.FileName ;
end;

procedure TFExpECC1.BValiderClick(Sender: TObject);
Var FEC : tFicEnCours ;
    Crit : tCritExpECC ;
    i : Integer ;
    LL : THCpteEdit ;
begin
if Msg.Execute(0,Caption,'')<>mrYes then Exit ;
If FileName.Text='' Then BEGIN Msg.Execute(2,Caption,'') ; Pages.ActivePage:=Pages.Pages[0] ; FileName.SetFocus ; Exit ; END ;
end;

procedure TFExpECC1.FFiltresChange(Sender: TObject);
begin
LoadFiltre(NomFiltre,FFiltres,Pages) ;
end;

procedure TFExpECC1.POPFPopup(Sender: TObject);
begin
UpdatePopFiltre(BSaveFiltre,BDelFiltre,BRenFiltre,FFiltres) ;
end;

procedure TFExpECC1.FormatChange(Sender: TObject);
begin
DateNeg1.Enabled:=Format.Value='NEG' ; DateNeg2.Enabled:=Format.Value='NEG' ;
HDateNeg1.Enabled:=Format.Value='NEG' ; HDateNeg2.Enabled:=Format.Value='NEG' ;
FMethodeECTENR.Enabled:=Format.Value<>'NEG' ;
CBExportDetail.enabled:=Format.Value='PRO' ;
OkAutreMP.Enabled:=Format.Value='ORL' ;
If Not CBExportDetail.enabled Then CBExportDetail.Checked:=FALSE ;
If Not OkAutreMP.enabled Then OkAutreMP.Checked:=FALSE ;
end;

procedure TFExpECC1.FormClose(Sender: TObject; var Action: TCloseAction);
begin
if Parent is THPanel then
  BEGIN
  Action:=caFree ;
  END ;
end;

procedure TFExpECC1.CBExportDetailClick(Sender: TObject);
begin
FMethodeECTNE.Visible:=Not CBExportDetail.Checked ; FMethodeECTENR.Visible:=Not CBExportDetail.Checked ;
FDateButoir.Enabled:=Not CBExportDetail.Checked ; LFDateButoir.Enabled:=Not CBExportDetail.Checked ;
T413EAR.Enabled:=Not CBExportDetail.Checked ; ST413EAR.Enabled:=Not CBExportDetail.Checked ;
T413EEP.Enabled:=Not CBExportDetail.Checked ; ST413EEP.Enabled:=Not CBExportDetail.Checked ;
TCPTSOLDE.Enabled:=Not CBExportDetail.Checked ;
OkPL.Visible:=CBExportDetail.Checked ; OkTL.Visible:=CBExportDetail.Checked ;
FDateEcr1.Visible:=CBExportDetail.Checked ; SFDateEcr1.Visible:=CBExportDetail.Checked ;
FDateEcr2.Visible:=CBExportDetail.Checked ; SFDateEcr2.Visible:=CBExportDetail.Checked ;
FDateModif1.Visible:=CBExportDetail.Checked ; SFDateModif1.Visible:=CBExportDetail.Checked ;
FDateModif2.Visible:=CBExportDetail.Checked ; SFDateModif2.Visible:=CBExportDetail.Checked ;
end;

end.
