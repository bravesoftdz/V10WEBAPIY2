unit ParamRel;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  FichList, DBTables, DB, StdCtrls, Hctrls, DBCtrls, Mask, hmsgbox,
  Buttons, ExtCtrls, Grids, DBGrids, HDB, Ent1, HEnt1, HSysMenu, Hqry,
  HTB97, HPanel, UiUtil ;

Procedure ParamRelance(Tipe : String3 ; LeQuel : String ; Action : TActionFiche) ;

type
  TFParamRel = class(TFFicheListe)
    TEX_EXERCICE: THLabel;
    RR_FAMILLERELANCE: TDBEdit;
    TEX_LIBELLE: THLabel;
    RR_LIBELLE: TDBEdit;
    GroupBox1: TGroupBox;
    RR_NONECHU: TDBCheckBox;
    RR_SCOORING: TDBCheckBox;
    RR_GROUPELETTRE: TDBCheckBox;
    GrpBoxEtatExo: TGroupBox;
    HRR_DELAI1: THLabel;
    HRR_DELAI2: THLabel;
    HRR_DELAI3: THLabel;
    HRR_DELAI4: THLabel;
    HRR_DELAI5: THLabel;
    HRR_DELAI6: THLabel;
    HRR_DELAI7: THLabel;
    Label2: TLabel;
    HRR_MODELE1: THLabel;
    HRR_MODELE2: THLabel;
    HRR_MODELE3: THLabel;
    HRR_MODELE4: THLabel;
    HRR_MODELE5: THLabel;
    HRR_MODELE6: THLabel;
    HRR_MODELE7: THLabel;
    Label1: TLabel;
    RR_DELAI1: TDBEdit;
    RR_MODELE1: THDBValComboBox;
    RR_MODELE2: THDBValComboBox;
    RR_DELAI2: TDBEdit;
    RR_MODELE3: THDBValComboBox;
    RR_DELAI3: TDBEdit;
    RR_MODELE4: THDBValComboBox;
    RR_DELAI4: TDBEdit;
    RR_MODELE5: THDBValComboBox;
    RR_DELAI5: TDBEdit;
    RR_MODELE6: THDBValComboBox;
    RR_DELAI6: TDBEdit;
    RR_MODELE7: THDBValComboBox;
    RR_DELAI7: TDBEdit;
    XX_WHERE: TPanel;
    TaRR_TYPERELANCE: TStringField;
    TaRR_FAMILLERELANCE: TStringField;
    TaRR_LIBELLE: TStringField;
    TaRR_GROUPELETTRE: TStringField;
    TaRR_NONECHU: TStringField;
    TaRR_SCOORING: TStringField;
    TaRR_DELAI1: TIntegerField;
    TaRR_MODELE1: TStringField;
    TaRR_DELAI2: TIntegerField;
    TaRR_MODELE2: TStringField;
    TaRR_DELAI3: TIntegerField;
    TaRR_MODELE3: TStringField;
    TaRR_DELAI4: TIntegerField;
    TaRR_MODELE4: TStringField;
    TaRR_DELAI5: TIntegerField;
    TaRR_MODELE5: TStringField;
    TaRR_DELAI6: TIntegerField;
    TaRR_MODELE6: TStringField;
    TaRR_DELAI7: TIntegerField;
    TaRR_MODELE7: TStringField;
    RR_TYPERELANCE: TDBEdit;
    procedure FormShow(Sender: TObject);
    procedure RR_DELAI1KeyPress(Sender: TObject; var Key: Char);
    procedure STaDataChange(Sender: TObject; Field: TField);
    procedure RR_SCOORINGClick(Sender: TObject);
    procedure RR_GROUPELETTREClick(Sender: TObject);
    procedure BImprimerClick(Sender: TObject);
  private    { Déclarations privées }
    Tipe     : String3;
    Function  EnregOK : boolean ; Override ;
    Procedure NewEnreg ; Override ;
    function  VerifJour : Boolean ;
    Function  ContinuiteDelaiModele : Boolean ;
    Procedure FaitCaption ;
  public    { Déclarations publiques }
  end;

implementation

uses PrintDBG, UtilPgi ;

{$R *.DFM}
//XX_Where.Hint:=' WHERE RR_TYPERELANCE="'+Tipe+'"' ;
//PrintDBGrid (Nil,XX_Where,Caption,'PRT_PARAMREL');

Procedure ParamRelance(Tipe : String3 ; LeQuel : String ; Action : TActionFiche) ;
var FParamRel: TFParamRel;
    PP : THPanel ;
begin
if _Blocage(['nrCloture'],False,'nrAucun') then Exit ;
FParamRel:=TFParamRel.Create(Application) ;
FParamRel.Tipe:=Tipe ;
if Tipe='RTR' then FParamRel.HelpContext := 7574000 else FParamRel.HelpContext := 7577000 ;
FParamRel.InitFL('RR','PRT_PARAMREL',LeQuel,Tipe,Action,TRUE,FParamRel.TaRR_FAMILLERELANCE,
                 FParamRel.TaRR_LIBELLE,NIL,['ttRelanceTraite','ttRelanceRegle']) ;
PP:=FindInsidePanel ;
if PP=Nil then
   BEGIN
    try
     FParamRel.ShowModal ;
    finally
     FParamRel.Free ;
    end ;
   Screen.Cursor:=SyncrDefault ;
   END else
   BEGIN
   InitInside(FParamRel,PP) ;
   FParamRel.Show ;
   END ;
end ;

Procedure TFParamRel.NewEnreg ;
BEGIN
Inherited ;
TaRR_TYPERELANCE.AsString:=Tipe ;
END ;

Procedure TFParamRel.FaitCaption ;
BEGIN
if Tipe='RRG' then Caption:=HM2.Mess[0]+' : '+TaRR_FAMILLERELANCE.AsString+' '+TaRR_LIBELLE.AsString
              else Caption:=HM2.Mess[1]+' : '+TaRR_FAMILLERELANCE.AsString+' '+TaRR_LIBELLE.AsString ;
UpdateCaption(Self) ;
END ;

procedure TFParamRel.FormShow(Sender: TObject);
var Reste, Index, Titre : string ;
    i : Integer ;
begin
  inherited;
if (Ta.Eof) And (Ta.Bof) And (FTypeAction<>taConsult) then
   BEGIN
   if ta.State=dsInsert then NewEnreg else BinsertClick(Nil) ;
   END;
FaitCaption ;
Reste:=HM2.Mess[2] ; Index:=ReadTokenSt(Reste) ; Titre:=ReadTokenSt(Reste) ;
if Tipe='RRG' then HM2.Mess[2]:=Index+';'+HM2.Mess[0]+';'+Reste
              else HM2.Mess[2]:=Index+';'+HM2.Mess[1]+';'+Reste ;
for i:=0 to HM.Mess.Count-1 do
  BEGIN
  Reste:=HM.Mess[i] ; Index:=ReadTokenSt(Reste) ; Titre:=ReadTokenSt(Reste) ;
  if (Titre<>'') then
    if Tipe='RRG' then HM.Mess[i]:=Index+';'+HM2.Mess[0]+';'+Reste
                  else HM.Mess[i]:=Index+';'+HM2.Mess[1]+';'+Reste ;

  END ;
// Seulement 3 niveaux de relance en S3
if V_PGI.LaSerie = S3 then
	begin
  for i:=4 to 7 do
      BEGIN
      TWinControl(FindComponent('HRR_DELAI'+IntToStr(i))).Visible:=False  ; TWinControl(FindComponent('RR_DELAI'+IntToStr(i))).Visible:=False ;
      TWinControl(FindComponent('HRR_MODELE'+IntToStr(i))).Visible:=False ; TWinControl(FindComponent('RR_MODELE'+IntToStr(i))).Visible:=False ;
      END ;
  RR_SCOORING.Visible:=False ;
  end;
end;

Function TFParamRel.ContinuiteDelaiModele : Boolean ;
Var i : Byte ;
    Trou : Boolean ;
BEGIN
Result:=False ; Trou:=True ;
for i:=1 to 7 do
   if Ta.FindField('RR_DELAI'+IntToStr(i)).AsInteger<>0 then BEGIN Trou:=False ; Break ; END ;
if Trou then BEGIN HM2.Execute(7,caption,'') ; RR_DELAI1.SetFocus ; Exit ; END ;
for i:=7 Downto 2 do
   BEGIN
   if(Ta.FindField('RR_DELAI'+IntToStr(i)).AsInteger<>0) And
     (Ta.FindField('RR_MODELE'+IntToStr(i)).AsString<>'') then
     if(Ta.FindField('RR_DELAI'+IntToStr(i-1)).AsInteger=0) And
       (Ta.FindField('RR_MODELE'+IntToStr(i-1)).AsString='') then
       BEGIN
       HM2.Execute(8,caption,'') ; TDBEdit(FindComponent('RR_DELAI'+IntToStr(i-1))).SetFocus ; Exit ;
       END ;
   END ;
Result:=True ;
END ;

function TFParamRel.VerifJour : Boolean ;
Var Model : String ;
    Delai,i : Integer ;
BEGIN
Result:=False ;
for i:=1 to 7 do
  BEGIN
  Model:=Ta.FindField('RR_MODELE'+IntToStr(i)).AsString ;
  Delai:=Ta.FindField('RR_DELAI'+IntToStr(i)).AsInteger ;
  if(Delai=0) And (Model<>'') then
    BEGIN
    If i<>1 Then
      BEGIN
      HM2.Execute(5,caption,'') ; TDBEdit(FindComponent('RR_DELAI'+IntToStr(i))).SetFocus ; Exit ;
      END ;
    END else
    if(Delai<>0) And (Model='') then
       BEGIN
       HM2.Execute(6,caption,'') ; THDBValComboBox(FindComponent('RR_MODELE'+IntToStr(i))).SetFocus ; Exit ;
       END else
       if i>1 then
          BEGIN
          if(Delai<>0) And (Model<>'') then
             BEGIN
             if Delai<=Ta.FindField('RR_DELAI'+IntToStr(i-1)).AsInteger then
                BEGIN
                HM2.Execute(2,HM2.Mess[4]+' '+IntToStr(i)+' ','') ;
                TDBEdit(FindComponent('RR_DELAI'+IntToStr(i))).SetFocus ; Exit ;
                END ;
             END ;
          END ;
  END ;
Result:=True ;
END ;

function TFParamRel.EnregOk : boolean ;
BEGIN
Result:=InHerited EnregOK ; if Not Result then Exit ;
Modifier:=True ;
if ((Result) and (Ta.state in [dsEdit,dsInsert])) then
  BEGIN
  Result:=False ;
  if Not ContinuiteDelaiModele then Exit ;
  if Not VerifJour then Exit ;
  END ;
Result:=True ; Modifier:=False ;
END ;

procedure TFParamRel.RR_DELAI1KeyPress(Sender: TObject; var Key: Char);
begin
  inherited;
if (Key='-') then BEGIN beep ; Key:=Chr(0) ; END ;
end;

procedure TFParamRel.STaDataChange(Sender: TObject; Field: TField);
begin
  inherited;
if Field=Nil then FaitCaption ;
end;

procedure TFParamRel.RR_SCOORINGClick(Sender: TObject);
begin if RR_SCOORING.Checked then RR_GROUPELETTRE.Checked:=False ; end;

procedure TFParamRel.RR_GROUPELETTREClick(Sender: TObject);
begin if RR_GROUPELETTRE.Checked then RR_SCOORING.Checked:=False ; end;

procedure TFParamRel.BImprimerClick(Sender: TObject);
var MyBookmark: TBookmark ;
begin
XX_Where.Hint:=' WHERE RR_TYPERELANCE="'+Tipe+'"' ;
MyBookmark :=Ta.GetBookmark ;
PrintDBGrid (Nil,XX_WHERE,Copy(Caption,1,Pos(':',Caption)-1),'PRT_PARAMREL') ;
Ta.GotoBookmark(MyBookmark) ; Ta.FreeBookmark(MyBookmark) ;
end;

end.
