unit Zoomana;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Mul, Mask, Hctrls, Hcompte, StdCtrls, Menus, DB, Hqry, Grids,
  {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}
  DBGrids, ExtCtrls, ComCtrls, Buttons, Hent1, Ent1, SaisComm, SaisUtil,SaisODA,
  HRichEdt, HSysMenu, CritEdt, HDB, HTB97, ColMemo, HMsgBox,
  HRichOLE, ADODB, HPanel,AGLInit ;

Procedure VisuZoomAna(Jal,Cpte,Section,QuelAxe,Lexo : String) ;

type
  TFZoomAna = class(TFMul)
    TY_EXERCICE: TLabel;
    TY_JOURNAL: TLabel;
    Y_JOURNAL: THValComboBox;
    Y_EXERCICE: THValComboBox;
    TY_NUMEROPIECE: THLabel;
    TY_DATECOMPTABLE: TLabel;
    Y_DATECOMPTABLE: THCritMaskEdit;
    Y_NUMEROPIECE: THCritMaskEdit;
    TY_NUMEROPIECE_: THLabel;
    Y_NUMEROPIECE_: THCritMaskEdit;
    TY_DATECOMPTABLE_: TLabel;
    Y_DATECOMPTABLE_: THCritMaskEdit;
    TY_SECTION: TLabel;
    TY_GENERAL: TLabel;
    Y_GENERAL: THCpteEdit;
    Y_SECTION: THCpteEdit;
    TY_NUMEROLIGNE: THLabel;
    Y_NUMLIGNE: THCritMaskEdit;
    TY_NUMEROLIGNE_: THLabel;
    Y_NUMLIGNE_: THCritMaskEdit;
    TY_REFINTERNE: THLabel;
    TY_AXE: TLabel;
    Y_AXE: THValComboBox;
    TY_NATUREPIECE: THLabel;
    Y_NATUREPIECE: THValComboBox;
    TY_QUALIFPIECE: THLabel;
    Y_QUALIFPIECE: THValComboBox;
    Y_REFINTERNE: TEdit;
    BGL: TToolbarButton97;
    BMenuZoom: TToolbarButton97;
    PopZ: TPopupMenu;
    BVisuP: TToolbarButton97;
    Pzlibre: TTabSheet;
    Bevel5: TBevel;
    TY_TABLE0: TLabel;
    Y_TABLE0: THCpteEdit;
    TY_TABLE2: TLabel;
    Y_TABLE2: THCpteEdit;
    TY_TABLE3: TLabel;
    Y_TABLE3: THCpteEdit;
    Y_TABLE1: THCpteEdit;
    TY_TABLE1: TLabel;
    TY_ETABLISSSEMENT: THLabel;
    Y_ETABLISSEMENT: THValComboBox;
    procedure FormShow(Sender: TObject);
    procedure FListeDblClick(Sender: TObject); override;
    procedure BNouvRechClick(Sender: TObject);
    procedure Y_EXERCICEChange(Sender: TObject);
    procedure BOuvrirClick(Sender: TObject); override;
    procedure Y_AXEChange(Sender: TObject);
{$IFNDEF CCMP}
    procedure BGLClick(Sender: TObject);
{$ENDIF}
    procedure BVisuPClick(Sender: TObject);
    procedure HMTradBeforeTraduc(Sender: TObject);
    procedure FormCreate(Sender: TObject); override;
    procedure BMenuZoomMouseEnter(Sender: TObject);
  private    { Déclarations privées }
    procedure InitCriteres;

  public
    { Déclarations publiques }
  end;

implementation

{$R *.DFM}

Uses
  {$IFDEF MODENT1}
  CPProcMetier,
  CPProcGen,
  ULibExercice,
  {$ENDIF MODENT1}
  UtilEdt,
{$IFNDEF CCMP}
     uTofCPGLAna,
{$ENDIF}
     FILTRE,UtilPGI ;

Procedure VisuZoomAna(Jal,Cpte,Section,QuelAxe,Lexo : String) ;
var FZoomAna: TFZoomAna ;
begin
if _Blocage(['nrCloture'],False,'nrAucun') then Exit ;
FZoomAna:=TFZoomAna.Create(Application) ;
 try
  FZoomAna.FNomFiltre:='CPZOOMANA' ;
  FZoomAna.Q.Manuel:=TRUE ;
  FZoomAna.Q.Liste:='CPZOOMANAL' ;
  If Jal<>'' Then FZoomAna.Y_JOURNAL.Value:=Jal
             else FZoomAna.Y_JOURNAL.ItemIndex:=0 ;
  FZoomAna.BGL.Enabled:=Not (Jal<>'') ;
  If Cpte<>'' Then FZoomAna.Y_GENERAL.Text:=Cpte ;
  If Section<>'' Then FZoomAna.Y_SECTION.Text:=Section ;
  FZoomAna.Y_EXERCICE.Value:=Lexo ;
  FZoomAna.Y_AXE.Value:=QuelAxe ;
  FZoomAna.Y_SECTION.ZoomTable:=AxeToTz(QuelAxe) ;
  FZoomAna.Y_QUALIFPIECE.ItemIndex:=0 ;
  FZoomAna.Y_NATUREPIECE.ItemIndex:=0 ;
  FZoomAna.ShowModal ;
 finally
  FZoomAna.Free ;
 end ;
Screen.Cursor:=SyncrDefault ;
END ;


procedure TFZoomAna.FormShow(Sender: TObject);
begin
Y_AXE.Enabled:=False ;
PositionneEtabUser(Y_ETABLISSEMENT) ;
  inherited;
Q.Manuel:=FALSE ; ChercheClick ;
Caption:=Caption+Y_SECTION.Text ;
UpdateCaption(Self) ;
{$IFDEF CCS3}
Y_AXE.Visible:=False ; TY_AXE.Visible:=False ; 
{$ENDIF}
end;

procedure TFZoomAna.FListeDblClick(Sender: TObject);
begin
  inherited;
TrouveEtLanceSaisieODA(Q,taConsult) ;
end;

procedure TFZoomAna.BVisuPClick(Sender: TObject);
begin
  inherited;
TrouveEtLanceSaisieODA(Q,taConsult) ;
end;

procedure TFZoomAna.InitCriteres;
begin
if VH^.Precedent.Code<>'' then Y_DATECOMPTABLE.Text:=DateToStr(VH^.Precedent.Deb)
                          else Y_DATECOMPTABLE.Text:=DateToStr(VH^.Encours.Deb) ;
Y_DATECOMPTABLE_.Text:=DateToStr(V_PGI.DateEntree) ;
end;


procedure TFZoomAna.BNouvRechClick(Sender: TObject);
begin
  inherited;
InitCriteres ;
end;

procedure TFZoomAna.Y_EXERCICEChange(Sender: TObject);
begin
  inherited;
ExoToDates(Y_EXERCICE.Value,Y_DATECOMPTABLE,Y_DATECOMPTABLE_) ;
end;

procedure TFZoomAna.BOuvrirClick(Sender: TObject);
begin
  inherited;
FListeDblClick(Nil) ;
end;

procedure TFZoomAna.Y_AXEChange(Sender: TObject);
begin
  inherited;
if Y_AXE.Value='' then Exit ;
Y_SECTION.ZoomTable:=AxeToTz(Y_AXE.Value) ;
end;

{$IFNDEF CCMP}
procedure TFZoomAna.BGLClick(Sender: TObject);
Var
  ACritEdt : ClassCritEdt;
  D1,D2 : TdateTime ;
begin
  inherited;
  ACritEdt := ClassCritEdt.Create;
  Fillchar(ACritEdt.CritEdt, SizeOf(ACritEdt.CritEdt), #0);
  ACritEdt.CritEdt.GL.Axe := Y_AXE.Value ;
  D1:=StrToDate(Y_DATECOMPTABLE.text) ; D2:=StrToDate(Y_DATECOMPTABLE_.text) ;
  ACritEdt.CritEdt.Date1:=D1 ; ACritEdt.CritEdt.Date2:=D2 ;
  ACritEdt.CritEdt.DateDeb:=ACritEdt.CritEdt.Date1 ;
  ACritEdt.CritEdt.DateFin:=ACritEdt.CritEdt.Date2 ;
  ACritEdt.CritEdt.Cpt1:=Y_SECTION.text ; ACritEdt.CritEdt.Cpt2:=ACritEdt.CritEdt.Cpt1 ;
  TheData := ACritEdt;
  CPLanceFiche_CPGLAna;
  TheData := nil;
  ACritEdt.Free;
end;
{$ENDIF}

procedure TFZoomAna.HMTradBeforeTraduc(Sender: TObject);
begin
  inherited;
LibellesTableLibre(PzLibre,'TY_TABLE','Y_TABLE','A') ;
end;

procedure TFZoomAna.FormCreate(Sender: TObject);
begin
  inherited;
MemoStyle:=msBook ; 
end;

procedure TFZoomAna.BMenuZoomMouseEnter(Sender: TObject);
begin
  inherited;
PopZoom97(BMenuZoom,POPZ) ;
end;

end.
