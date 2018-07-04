unit Modepaie;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  FichList, {$IFNDEF DBXPRESS}dbtables, Db, StdCtrls, Hcompte, HRegCpte,
  Hctrls, DBCtrls, Mask, HSysMenu, hmsgbox, Hqry, HTB97, ExtCtrls, HPanel,
  Grids, DBGrids, HDB{$ELSE}uDbxDataSet{$ENDIF}, DB, StdCtrls, Hcompte, HDB, Spin, Hctrls, DBCtrls,
  Mask, hmsgbox, Buttons, ExtCtrls, Grids, DBGrids, Hent1, Ent1, HSysMenu,
  Hqry, HRegCpte, HTB97, UiUtil, HPanel ;

Procedure FicheModePaie(Quel : String) ;
Function FicheModePaieZoom(Quel : String) : String ;

type
  TFModePaie = class(TFFicheListe)
    TMP_MODEPAIE: THLabel;
    MP_MODEPAIE: TDBEdit;
    MP_SOCIETE: TDBEdit;
    TMP_ABREGE: THLabel;
    MP_ABREGE: TDBEdit;
    TMP_LIBELLE: THLabel;
    MP_LIBELLE: TDBEdit;
    MP_EDITABLE: TDBCheckBox;
    MP_POINTABLE: TDBCheckBox;
    MP_CONDITION: TDBCheckBox;
    HLabel1: THLabel;
    MP_CATEGORIE: THDBValComboBox;
    TMP_CODEACCEPT: THLabel;
    MP_CODEACCEPT: THDBValComboBox;
    HLabel2: THLabel;
    MP_ENCAISSEMENT: THDBValComboBox;
    TMP_GENERAL: THLabel;
    MP_GENERAL: THDBCpteEdit;
    TaMP_MODEPAIE: TStringField;
    TaMP_LIBELLE: TStringField;
    TaMP_ABREGE: TStringField;
    TaMP_FORMATCFONB: TStringField;
    TaMP_CODEAFB: TStringField;
    TaMP_CONDITION: TStringField;
    TaMP_MONTANTMAX: TFloatField;
    TaMP_REMPLACEMAX: TStringField;
    TaMP_EDITABLE: TStringField;
    TaMP_POINTABLE: TStringField;
    TaMP_GENERAL: TStringField;
    TaMP_ENCAISSEMENT: TStringField;
    TaMP_CATEGORIE: TStringField;
    TaMP_CODEACCEPT: TStringField;
    TTMP_GENERAL: THLabel;
    TaMP_DELAIRETIMPAYE: TIntegerField;
    PCondMt: TPanel;
    MP_MONTANTMAX: TDBEdit;
    TMP_MONTANTMAX: THLabel;
    MP_REMPLACEMAX: THDBValComboBox;
    TMP_REMPLACEMAX: THLabel;
    TaMP_LETTRECHEQUE: TStringField;
    MP_LETTRECHEQUE: TDBCheckBox;
    MP_LETTRETRAITE: TDBCheckBox;
    TaMP_LETTRETRAITE: TStringField;
    procedure MP_CATEGORIEChange(Sender: TObject);
    procedure MP_GENERALExit(Sender: TObject);
    procedure BDeleteClick(Sender: TObject);
    procedure MP_CONDITIONClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure BValiderClick(Sender: TObject);
    procedure FListeDblClick(Sender: TObject);
  private    { Déclarations privées }
    MpRetour : String ;
    OkZoom : Boolean ;
    Function DansModRegl : Boolean ;
    Function EnregOK : boolean ; Override ;
    Procedure NewEnreg ; Override ;
    Procedure ChargeEnreg ; Override ;
  public    { Déclarations publiques }
  end;

implementation

{$R *.DFM}

Procedure FicheModePaie(Quel : String) ;
var FModePaie: TFModePaie;
    PP : THPanel ;
begin
if Blocage(['nrCloture','nrBatch'],True,'nrAucun') then Exit ;
FModePaie:=TFModePaie.Create(Application) ;
FModePaie.InitFL('MP','PRT_MODEPAIE',Quel,'',taModif,TRUE,FModePaie.TaMP_MODEPAIE,
                 FModePaie.TaMP_LIBELLE,FModePaie.TaMP_MODEPAIE,['ttModePaie']) ;
FModePaie.TTMP_GENERAL.Caption:='' ; FModePaie.OkZoom:=FALSE ;
PP:=FindInsidePanel ;
if PP=Nil then
   BEGIN
    try
     FModePaie.ShowModal ;
    finally
     FModePaie.Free ;
    end ;
   Screen.Cursor:=SyncrDefault ;
   END else
   BEGIN
   InitInside(FModePaie,PP) ;
   FModePaie.Show ;
   END ;
end ;

Function FicheModePaieZoom(Quel : String) : String ;
var FModePaie: TFModePaie;
begin
Result:='' ;
FModePaie:=TFModePaie.Create(Application) ;
FModePaie.InitFL('MP','PRT_MODEPAIE',Quel,'',taModif,TRUE,FModePaie.TaMP_MODEPAIE,
                 FModePaie.TaMP_LIBELLE,FModePaie.TaMP_MODEPAIE,['ttModePaie']) ;
FModePaie.TTMP_GENERAL.Caption:='' ; FModePaie.OkZoom:=TRUE ;
try
 FModePaie.ShowModal ;
finally
 If FModePaie.MPRetour<>'' Then Result:=FModePaie.MPRetour ;
 FModePaie.Free ;
end ;
Screen.Cursor:=SyncrDefault ;
end ;


Procedure TFModePaie.ChargeEnreg ;
BEGIN
Inherited ;
MP_GENERAL.ExisteH ;
MP_CATEGORIEChange(Nil) ;
PCondMt.Enabled:=MP_CONDITION.checked ;
END ;

Procedure TFModePaie.NewEnreg ;
BEGIN
Inherited ;
TMP_CODEACCEPT.Enabled:=False ;
MP_CODEACCEPT.Enabled:=False ;
END ;

Function TFModePaie.EnregOK : boolean ;
BEGIN
result:=Inherited EnregOK  ; if Not Result then Exit ;
Modifier:=True ;
if ((result) and (Ta.state in [dsEdit,dsInsert])) then
   BEGIN
   Result:=FALSE ;
   //CP 23/01/97...
   //if TaMP_GENERAL.AsString='' then BEGIN HM2.Execute(0,'','') ; MP_GENERAL.SetFocus ; Exit ; END ;
   if ((MP_GENERAL.Text<>'') and (MP_GENERAL.ExisteH<=0)) then BEGIN HM2.Execute(4,'','') ; MP_GENERAL.SetFocus ; END ;
   if MP_CATEGORIE.Value='LCR' then
      BEGIN
      if TaMP_CODEACCEPT.AsString='' then
         BEGIN HM2.Execute(2,'','') ; MP_CODEACCEPT.SetFocus ; Exit ; END ;
      END ;
   if ((Trim(MP_Montantmax.Text)<>'') and (StrToFloat(MP_Montantmax.Text)<0)) then
          BEGIN HM2.Execute(1,'','') ; MP_MONTANTMAX.SetFocus ; exit ; END ;
   if TaMP_REMPLACEMAX.AsString=TaMP_MODEPAIE.AsString then BEGIN HM2.Execute(3,'','') ; MP_REMPLACEMAX.SetFocus ; Exit ; END ;
   if ((TaMP_REMPLACEMAX.AsString='') and (TaMP_MONTANTMAX.AsFloat<>0)) then BEGIN HM2.Execute(6,'','') ; MP_REMPLACEMAX.SetFocus ; Exit ; END ;
   END ;
Result:=TRUE  ; Modifier:=False ;
if TaMP_CATEGORIE.AsString<>'LCR' then TaMP_CODEACCEPT.AsString:='' ;
END ;

Function TFModePaie.DansModRegl : Boolean ;
Var SQl : String ;
BEGIN
SQL:='SELECT MR_MP1, MR_MP2, MR_MP3, MR_MP4, MR_MP5, MR_MP6, '
                   + 'MR_MP7, MR_MP8, MR_MP9, MR_MP10, MR_MP11, MR_MP12 '
                   + 'FROM MODEREGL ' ;
SQL:=SQL+' WHERE MR_MP1="'+MP_MODEPAIE.Text+'" or MR_MP2="'+MP_MODEPAIE.Text+'" ';
SQL:=SQL+' or MR_MP3="'+MP_MODEPAIE.Text+'" or MR_MP4="'+MP_MODEPAIE.Text+'" ';
SQL:=SQL+' or MR_MP5="'+MP_MODEPAIE.Text+'" or MR_MP6="'+MP_MODEPAIE.Text+'" ';
SQL:=SQL+' or MR_MP7="'+MP_MODEPAIE.Text+'" or MR_MP8="'+MP_MODEPAIE.Text+'" ';
SQL:=SQL+' or MR_MP9="'+MP_MODEPAIE.Text+'" or MR_MP10="'+MP_MODEPAIE.Text+'" ';
SQL:=SQL+' or MR_MP11="'+MP_MODEPAIE.Text+'" or MR_MP12="'+MP_MODEPAIE.Text+'" ';
Result:=ExisteSQL(SQL) ; Screen.Cursor:=SyncrDefault ;
END ;



procedure TFModePaie.MP_CATEGORIEChange(Sender: TObject);
begin
  inherited;
TMP_CODEACCEPT.Enabled:=(MP_CATEGORIE.Value='LCR') ;
MP_CODEACCEPT.Enabled:=(MP_CATEGORIE.Value='LCR') ;
end;

procedure TFModePaie.MP_GENERALExit(Sender: TObject);
begin
  inherited;
MP_GENERAL.ExisteH ;
end;

procedure TFModePaie.BDeleteClick(Sender: TObject);
begin
if DansModRegl then BEGIN HM2.Execute(5,'','') ; Exit ; END ;
inherited;
end;

procedure TFModePaie.MP_CONDITIONClick(Sender: TObject);
begin
  inherited;
PCondMt.Enabled:=MP_CONDITION.Checked ;
end;

procedure TFModePaie.FormShow(Sender: TObject);
begin
  inherited;
if (Ta.Eof) And (Ta.Bof) And (FTypeAction<>taConsult) then
   BEGIN
   if ta.State=dsInsert then NewEnreg else BinsertClick(Nil) ;
   END ;
end;

procedure TFModePaie.BValiderClick(Sender: TObject);
begin
  inherited;
If OkZoom Then // ne peut être envoyé en inside
  BEGIN
  MPRetour:=MP_ModePaie.Text ;
  Close ;
  END ;
end;

procedure TFModePaie.FListeDblClick(Sender: TObject);
begin
  inherited;
If OkZoom Then // ne peut être envoyé en inside
  BEGIN
  MPRetour:=MP_ModePaie.Text ;
  Close ;
  END ;
end;

end.
