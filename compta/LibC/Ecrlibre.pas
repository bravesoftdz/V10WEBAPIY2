unit EcrLibre;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Mul, HSysMenu, Menus, DB, {$IFNDEF DBXPRESS}dbtables, hmsgbox, StdCtrls,
  Hcompte, Mask, Hctrls, ExtCtrls, Hqry, HRichOLE, ComCtrls, HRichEdt,
  Grids, DBGrids, HDB, HTB97, ColMemo{$ELSE}uDbxDataSet{$ENDIF}, Hqry, StdCtrls, Grids, DBGrids, HDB,
  ComCtrls, HRichEdt, Hctrls, ExtCtrls, Buttons, Ent1, HEnt1, hmsgbox,
  Hcompte, Mask, HTB97, ColMemo, HPanel, UiUtil, HRichOLE ;

type
  TFEcrLibre = class(TFMul)
    E_JOURNAL: THValComboBox;
    E_NUMLIGNE: THCritMaskEdit;
    TE_JOURNAL: THLabel;
    TE_EXERCICE: THLabel;
    TE_NATUREPIECE: THLabel;
    E_EXERCICE: THValComboBox;
    E_NATUREPIECE: THValComboBox;
    E_NUMLIGNE_: THCritMaskEdit;
    E_NUMECHE: THCritMaskEdit;
    TE_NUMEROPIECE: THLabel;
    TE_DATECOMPTABLE: THLabel;
    TE_REFINTERNE: THLabel;
    E_DATECOMPTABLE: THCritMaskEdit;
    E_NUMEROPIECE: THCritMaskEdit;
    E_NUMEROPIECE_: THCritMaskEdit;
    HLabel1: THLabel;
    TE_DATECOMPTABLE2: THLabel;
    E_DATECOMPTABLE_: THCritMaskEdit;
    E_REFINTERNE: TEdit;
    Bevel5: TBevel;
    E_QUALIFPIECE: THValComboBox;
    TE_QUALIFPIECE: THLabel;
    TE_DATECREATION: THLabel;
    TE_DATEECHEANCE: THLabel;
    E_DATEECHEANCE: THCritMaskEdit;
    E_DATECREATION: THCritMaskEdit;
    TE_DATECREATION_: THLabel;
    TE_DATEECHEANCE2: THLabel;
    E_DATECREATION_: THCritMaskEdit;
    E_DATEECHEANCE_: THCritMaskEdit;
    TE_GENERAL: THLabel;
    TE_UTILISATEUR: THLabel;
    TE_AFFAIRE: THLabel;
    E_GENERAL: THCpteEdit;
    E_UTILISATEUR: THValComboBox;
    E_AFFAIRE: TEdit;
    TE_AUXILIAIRE: THLabel;
    TE_DEVISE: THLabel;
    E_VALIDE: TCheckBox;
    E_AUXILIAIRE: THCpteEdit;
    E_DEVISE: THValComboBox;
    RUne: TCheckBox;
    HM: THMsgBox;
    procedure FormShow(Sender: TObject);
    procedure E_EXERCICEChange(Sender: TObject);
    procedure RUneClick(Sender: TObject);
    procedure BChercheClick(Sender: TObject);
    procedure BOuvrirClick(Sender: TObject);
  private
    procedure InitCriteres ;
  public
  end;

procedure MultiCritereEcrLibre(Lequel : String);

implementation

{$R *.DFM}


procedure MultiCritereEcrLibre(Lequel : String);
var FEcrLibre: TFEcrLibre;
    PP : THPanel ;
begin
FEcrLibre:=TFEcrLibre.Create(Application) ;
FEcrLibre.FNomFiltre:='ECRLIBRE'+Lequel ;
FEcrLibre.Caption:=FEcrLibre.HM.Mess[0] +' '+Lequel  ;
FEcrLibre.Q.Liste:='ECRLIBRE'+Lequel ;
FEcrLibre.Q.Manuel:=TRUE ;
PP:=FindInsidePanel ;
if PP=Nil then
   BEGIN
    try
     FEcrLibre.ShowModal ;
    finally
     FEcrLibre.Free ;
    end;
   Screen.Cursor:=SyncrDefault ;
   END else
   BEGIN
   InitInside(FEcrLibre,PP) ;
   FEcrLibre.Show ;
   END ;
end;

procedure TFEcrLibre.FormShow(Sender: TObject);
begin
if VH^.CPExoRef.Code<>'' then
   BEGIN
   E_EXERCICE.Value:=VH^.CPExoRef.Code ;
   E_DATECOMPTABLE.Text:=DateToStr(VH^.CPExoRef.Deb) ;
   E_DATECOMPTABLE_.Text:=DateToStr(VH^.CPExoRef.Fin) ;
   END else
   BEGIN
   E_EXERCICE.Value:=VH^.Entree.Code ;
   E_DATECOMPTABLE.Text:=DateToStr(V_PGI.DateEntree) ;
   E_DATECOMPTABLE_.Text:=DateToStr(V_PGI.DateEntree) ;
   END ;
E_DATECREATION.Text:=StDate1900 ; E_DATECREATION_.Text:=StDate2099 ;
E_DATEECHEANCE.Text:=StDate1900 ; E_DATEECHEANCE_.Text:=StDate2099 ;
  inherited;
InitCriteres ;
{
if ANouveau then
   BEGIN
   E_JOURNAL.DataType:=ttJalANouveau ; XX_WHERE.Text:='E_ECRANOUVEAU="H" OR E_ECRANOUVEAU="OAN"' ;
   END ;
}
if E_JOURNAL.Items.Count>0 then
   BEGIN
   if E_JOURNAL.Vide then E_JOURNAL.ItemIndex:=1 else E_JOURNAL.ItemIndex:=0 ;
   END ;
  inherited;
Q.Manuel:=FALSE ; Q.UpdateCriteres ;
If (V_PGI.OutLook) And (IsInside(Self)) Then If HMTrad.ResizeDBGrid Then HMTrad.ResizeDBGridColumns(Fliste) ;
end;

procedure TFEcrLibre.InitCriteres ;
BEGIN
if VH^.Precedent.Code<>'' then E_DATECOMPTABLE.Text:=DateToStr(VH^.Precedent.Deb)
                         else E_DATECOMPTABLE.Text:=DateToStr(VH^.Encours.Deb) ;
E_DATECOMPTABLE_.Text:=DateToStr(V_PGI.DateEntree) ;
END ;

procedure TFEcrLibre.E_EXERCICEChange(Sender: TObject);
begin
  inherited;
ExoToDates(E_EXERCICE.Value,E_DATECOMPTABLE,E_DATECOMPTABLE_) ;
end;

procedure TFEcrLibre.RUneClick(Sender: TObject);
begin
  inherited;
If RUne.Checked Then BEGIN E_NUMECHE.text:='1' ; E_NUMLIGNE.text:='1' ; E_NUMLIGNE_.text:='1' ; END
                Else BEGIN E_NUMECHE.text:='9999' ; E_NUMLIGNE.text:='0' ; E_NUMLIGNE_.text:='9999' ; END ;
end;

procedure TFEcrLibre.BChercheClick(Sender: TObject);
begin
If (E_GENERAL.text<>'') Or (E_AUXILIAIRE.Text<>'') Then RUne.Checked:=FALSE ;
  inherited;
end;

procedure TFEcrLibre.BOuvrirClick(Sender: TObject);
begin bExportClick(nil) ; end;

end.
