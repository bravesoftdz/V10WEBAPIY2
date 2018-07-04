unit ZoomMvt;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Mul, Hcompte, Mask, Hctrls, StdCtrls, Menus, DB, Hqry, Grids,
  {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}
  DBGrids, ExtCtrls, ComCtrls, Buttons, Ent1, HEnt1, SaisUtil, Saisie, SaisComm,
  HRichEdt, HSysMenu, HDB, HTB97, ColMemo, HMsgBox, HRichOLE, ADODB,
  {$IFDEF MODENT1}
  CPTypeCons,
  {$ENDIF MODENT1}
  HPanel ;

Function VisuZoomMvt(NatureCpte : TFichierBase ; Cpte : String ; Lexo : String ;
                     MDeb,MFin : String ; TypEcritur : String ; Action : TActionFiche) : byte ;

type
  TFZoomMvt = class(TFMul)
    TE_EXERCICE: TLabel;
    E_EXERCICE: THValComboBox;
    TE_JOURNAL: TLabel;
    E_JOURNAL: THValComboBox;
    TE_NATUREPIECE: THLabel;
    E_NATUREPIECE: THValComboBox;
    TE_NUMEROPIECE: THLabel;
    TE_REFINTERNE: THLabel;
    TE_DATECOMPTABLE: TLabel;
    E_DATECOMPTABLE: THCritMaskEdit;
    E_REFINTERNE: THCritMaskEdit;
    E_NUMEROPIECE: THCritMaskEdit;
    HLabel2: THLabel;
    E_NUMEROPIECE_: THCritMaskEdit;
    E_DATECOMPTABLE_: THCritMaskEdit;
    TE_DATECOMPTABLE_: TLabel;
    TE_DEBIT: TLabel;
    E_DEBIT: THCritMaskEdit;
    TE_DEBIT_: TLabel;
    E_DEBIT_: THCritMaskEdit;
    E_CREDIT_: THCritMaskEdit;
    TE_CREDIT_: TLabel;
    E_CREDIT: THCritMaskEdit;
    TE_CREDIT: TLabel;
    TE_GENERAL: TLabel;
    E_GENERAL: THCpteEdit;
    E_AUXILIAIRE: THCpteEdit;
    TE_AUXILIAIRE: TLabel;
    TE_ETABLISSEMENT: THLabel;
    E_ETABLISSEMENT: THValComboBox;
    procedure BNouvRechClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure FListeDblClick(Sender: TObject); override;
    procedure FormCreate(Sender: TObject); override;
  private    { Déclarations privées }
    NatureCpte : TFichierBase ;
    TypEcritur,Cpte : String ;
    ActionHyperZoom : TActionFiche ;
    AFaitQuoi : Byte ;
    procedure InitCriteres ;
  public
    { Déclarations publiques }
  end;

implementation

{$R *.DFM}

Uses UtilPGI ; 

Function VisuZoomMvt(NatureCpte : TFichierBase ; Cpte : String ; Lexo : String ;
                     MDeb,MFin : String ; TypEcritur : String ; Action : TActionFiche) : Byte ;
var FZoomMvt: TFZoomMvt ;
begin
Result:=0 ;
if _Blocage(['nrCloture'],True,'nrAucun') then Exit ;
FZoomMvt:=TFZoomMvt.Create(Application) ;
 try
  FZoomMvt.FNomFiltre:='CPZOOMMVT' ;
  FZoomMvt.Q.Manuel:=TRUE ;
  FZoomMVT.Q.Liste:='CPZOOMMVT' ;
  FZoomMvt.E_EXERCICE.Value:=Lexo ;
  FZoomMvt.NatureCpte:=NatureCpte ;
  FZoomMvt.Cpte:=Cpte ;
  FZoomMvt.E_DATECOMPTABLE.Text:=MDeb ;
  FZoomMvt.E_DATECOMPTABLE_.Text:=MFin ;
  FZoomMvt.TypEcritur:=TypEcritur ;
  FZoomMvt.ActionHyperZoom:=Action ;
  FZoomMvt.AFaitQuoi:=0 ;
  FZoomMvt.ShowModal ;
  If FZoomMvt.AFaitQuoi=1 Then Result:=1 ;
 finally
  FZoomMvt.Free ;
 end ;
Screen.Cursor:=SyncrDefault ;
END ;

procedure TFZoomMvt.BNouvRechClick(Sender: TObject);
begin
  inherited;
InitCriteres ;
end;

procedure TFZoomMvt.FormShow(Sender: TObject);
begin
Case NatureCpte of
   fbGene : BEGIN E_GENERAL.Text:=Cpte ; END ;
   fbAux   : BEGIN E_AUXILIAIRE.Text:=Cpte ; END ;
   fbJal : BEGIN E_JOURNAL.Value:=Cpte ; END ;
   end ;
PositionneEtabUser(E_ETABLISSEMENT) ;
  inherited;
Q.Manuel:=FALSE ;
end;

procedure TFZoomMvt.FListeDblClick(Sender: TObject);
begin
  inherited;
if TrouveEtLanceSaisie(Q,ActionHyperZoom,TypEcritur) then AFaitQuoi:=1 ;
end;

procedure TFZoomMvt.InitCriteres ;
BEGIN
if VH^.Precedent.Code<>'' then E_DATECOMPTABLE.Text:=DateToStr(VH^.Precedent.Deb)
                          else E_DATECOMPTABLE.Text:=DateToStr(VH^.Encours.Deb) ;
E_DATECOMPTABLE_.Text:=DateToStr(V_PGI.DateEntree) ;
END ;

procedure TFZoomMvt.FormCreate(Sender: TObject);
begin
  inherited;
MemoStyle:=msBook ; 
end;

end.
