unit AnaLibre;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Mul, hmsgbox, HSysMenu, Menus, DB, {$IFNDEF DBXPRESS}dbtables, Hcompte,
  Mask, Hctrls, StdCtrls, ExtCtrls, Hqry, HRichOLE, ComCtrls, HRichEdt,
  Grids, DBGrids, HDB, HTB97, ColMemo{$ELSE}uDbxDataSet{$ENDIF}, Hqry, StdCtrls, Grids,
  DBGrids, HDB, ComCtrls, HRichEdt, Hctrls, ExtCtrls, Buttons, Hcompte,
  Mask, Ent1, HEnt1, HTB97, ColMemo, HPanel, UiUtil, HRichOLE ;

type
  TFAnaLibre = class(TFMul)
    HM: THMsgBox;
    Y_JOURNAL: THValComboBox;
    Y_NUMVENTIL_: THCritMaskEdit;
    Y_NUMVENTIL: THCritMaskEdit;
    Y_EXERCICE: THValComboBox;
    Y_AXE: THValComboBox;
    TY_EXERCICE: THLabel;
    TY_JOURNAL: TLabel;
    TY_AXE: TLabel;
    TY_NUMEROPIECE: TLabel;
    TY_DATECOMPTABLE: THLabel;
    Y_TYPEANALYTIQUE: TCheckBox;
    Y_DATECOMPTABLE: THCritMaskEdit;
    Y_NUMEROPIECE: THCritMaskEdit;
    TY_DATECOMPTABLE_: THLabel;
    HLabel1: THLabel;
    Y_NUMEROPIECE_: THCritMaskEdit;
    Y_DATECOMPTABLE_: THCritMaskEdit;
    Y_QUALIFPIECE: THValComboBox;
    Y_DATECREATION: THCritMaskEdit;
    Y_DEVISE: THValComboBox;
    Y_DATECREATION_: THCritMaskEdit;
    TY_DATECREATION_: THLabel;
    TY_DATECREATION: THLabel;
    TY_DEVISE: THLabel;
    TY_QUALIFPIECE: THLabel;
    TY_GENERAL: TLabel;
    TY_QUALIFQTE1: THLabel;
    RUne: TCheckBox;
    Y_QUALIFQTE1: THValComboBox;
    Y_GENERAL: THCpteEdit;
    TY_SECTION: TLabel;
    TY_QUALIFQTE2: THLabel;
    Y_QUALIFQTE2: THValComboBox;
    Y_SECTION: THCpteEdit;
    Y_VALIDE: TCheckBox;
    procedure Y_EXERCICEChange(Sender: TObject);
    procedure Y_AXEChange(Sender: TObject);
    procedure Y_QUALIFPIECEChange(Sender: TObject);
    procedure RUneClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure BChercheClick(Sender: TObject);
    procedure BOuvrirClick(Sender: TObject);
  private
    procedure InitCriteres ;
  public
  end;

procedure MultiCritereAnaLibre(Lequel : String);

implementation

{$R *.DFM}

procedure MultiCritereAnaLibre(Lequel : String);
var FAnaLibre: TFAnaLibre;
    PP : THPanel ;
begin
FAnaLibre:=TFAnaLibre.Create(Application) ;
FAnaLibre.FNomFiltre:='ANALIBRE'+Lequel ;
FAnaLibre.Caption:=FAnaLibre.HM.Mess[0] +' '+Lequel  ;
FAnaLibre.Q.Liste:='ANALIBRE'+Lequel ;
FAnaLibre.Q.Manuel:=TRUE ;
PP:=FindInsidePanel ;
if PP=Nil then
   BEGIN
    try
     FAnaLibre.ShowModal ;
    finally
     FAnaLibre.Free ;
    end;
   Screen.Cursor:=SyncrDefault ;
   END else
   BEGIN
   InitInside(FAnalibre,PP) ;
   FAnalibre.Show ;
   END ;
end;

procedure TFAnaLibre.Y_EXERCICEChange(Sender: TObject);
begin
  inherited;
ExoToDates(Y_EXERCICE.Value,Y_DATECOMPTABLE,Y_DATECOMPTABLE_) ;
end;

procedure TFAnaLibre.Y_AXEChange(Sender: TObject);
Var tt : TZoomTable ;
    i : Byte ;
begin
  inherited;
if Y_AXE.Value='' then Exit ;
tt:=tzSection ; Dec(tt) ; i:=StrToInt(Copy(Y_AXE.Value,2,1)) ; Byte(tt):=Byte(tt)+i ;
Y_SECTION.ZoomTable:=tt ;
end;

procedure TFAnaLibre.Y_QUALIFPIECEChange(Sender: TObject);
begin
  inherited;
if Y_QUALIFPIECE.Value<>'R' then Exit ;
if Not V_PGI.Controleur then Y_QUALIFPIECE.Value:='N' ;
end;

procedure TFAnaLibre.RUneClick(Sender: TObject);
begin
  inherited;
If RUne.Checked Then BEGIN Y_NUMVENTIL.Text:='1' ; Y_NUMVENTIL_.text:='1' ; END
                Else BEGIN Y_NUMVENTIL.Text:='0' ; Y_NUMVENTIL_.text:='9999' ; END ;
end;

procedure TFAnaLibre.FormShow(Sender: TObject);
Var Cpt1 : String17 ;
    Date1,Date2 : TDateTime ;
begin
if VH^.CPExoRef.Code<>'' then
   BEGIN
   Y_EXERCICE.Value:=VH^.CPExoRef.Code ;
   Y_DATECOMPTABLE.Text:=DateToStr(VH^.CPExoRef.Deb) ;
   Y_DATECOMPTABLE_.Text:=DateToStr(VH^.CPExoRef.Fin) ;
   END else
   BEGIN
   Y_EXERCICE.Value:=VH^.Entree.Code ;
   Y_DATECOMPTABLE.Text:=DateToStr(V_PGI.DateEntree) ;
   Y_DATECOMPTABLE_.Text:=DateToStr(V_PGI.DateEntree) ;
   END ;
Y_DATECREATION.Text:=StDate1900 ; Y_DATECREATION_.Text:=StDate2099 ;
InitCriteres ;
Cpt1:=SQLPremierDernier(fbJal,TRUE) ;
Y_Journal.Value:=Cpt1 ;
Y_AXE.Value:='A1' ;
  inherited;
Q.Manuel:=FALSE ; Q.UpdateCriteres ;
CentreDBGrid(FListe) ;
If (V_PGI.OutLook) And (IsInside(Self)) Then If HMTrad.ResizeDBGrid Then HMTrad.ResizeDBGridColumns(Fliste) ;
end;

procedure TFAnaLibre.InitCriteres ;
BEGIN
if VH^.Precedent.Code<>'' then Y_DATECOMPTABLE.Text:=DateToStr(VH^.Precedent.Deb)
                         else Y_DATECOMPTABLE.Text:=DateToStr(VH^.Encours.Deb) ;
Y_DATECOMPTABLE_.Text:=DateToStr(V_PGI.DateEntree) ;
END ;


procedure TFAnaLibre.BChercheClick(Sender: TObject);
begin
If (Y_GENERAL.Text<>'') Or (Y_SECTION.Text<>'') then RUne.Checked:=FALSE ;
  inherited;
end;

procedure TFAnaLibre.BOuvrirClick(Sender: TObject);
begin bExportClick(nil) ; end;

end.
