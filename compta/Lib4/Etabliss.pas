unit Etabliss;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  FichList, DB, {$IFNDEF DBXPRESS}dbtables, StdCtrls, Mask, DBCtrls,
  Hctrls, ExtCtrls, HSysMenu, hmsgbox, Hqry, HTB97, HPanel, Grids, DBGrids,
  HDB{$ELSE}uDbxDataSet{$ENDIF}, StdCtrls, Hctrls, Mask, DBCtrls, hmsgbox, Region,
  Buttons, ExtCtrls, Grids, DBGrids, HDB, HEnt1, Ent1, CodePost, HSysMenu,
  Hqry, HTB97, HPanel, UiUtil ;

Procedure FicheEtablissement(Mode : TActionFiche) ;

type
  TFEtablis = class(TFFicheListe)
    TET_ETABLISSEMENT: THLabel;
    ET_ETABLISSEMENT: TDBEdit;
    TET_ABREGE: THLabel;
    ET_LIBELLE: TDBEdit;
    TET_LIBELLE: THLabel;
    TET_ADRESSE1: THLabel;
    TET_ADRESSE2: THLabel;
    TET_ADRESSE3: THLabel;
    TET_CODEPOSTAL: THLabel;
    TET_DIVTERRIT: THLabel;
    TET_TELEPHONE: THLabel;
    TET_FAX: THLabel;
    TET_TELEX: THLabel;
    ET_TELEX: TDBEdit;
    ET_FAX: TDBEdit;
    ET_TELEPHONE: TDBEdit;
    ET_DIVTERRIT: TDBEdit;
    ET_CODEPOSTAL: TDBEdit;
    ET_ADRESSE3: TDBEdit;
    ET_ADRESSE2: TDBEdit;
    ET_ADRESSE1: TDBEdit;
    ET_VILLE: TDBEdit;
    ET_PAYS: THDBValComboBox;
    TET_ETABLIE: THLabel;
    TET_PAYS: THLabel;
    TET_VILLE: THLabel;
    TET_LANGUE: THLabel;
    TET_JURIDIQUE: THLabel;
    ET_JURIDIQUE: THDBValComboBox;
    ET_LANGUE: THDBValComboBox;
    ET_ETABLIE: THDBValComboBox;
    TaET_ETABLISSEMENT: TStringField;
    TaET_LIBELLE: TStringField;
    TaET_ABREGE: TStringField;
    TaET_ADRESSE1: TStringField;
    TaET_ADRESSE2: TStringField;
    TaET_ADRESSE3: TStringField;
    TaET_CODEPOSTAL: TStringField;
    TaET_VILLE: TStringField;
    TaET_PAYS: TStringField;
    TaET_LANGUE: TStringField;
    TaET_TELEPHONE: TStringField;
    TaET_FAX: TStringField;
    TaET_TELEX: TStringField;
    TaET_SOCIETE: TStringField;
    TaET_JURIDIQUE: TStringField;
    TaET_ETABLIE: TStringField;
    ET_ABREGE: TDBEdit;
    Bevel1: TBevel;
    Bevel2: TBevel;
    TaET_DIVTERRIT: TStringField;
    procedure ET_CODEPOSTALDblClick(Sender: TObject);
    procedure ET_VILLEDblClick(Sender: TObject);
    procedure ET_DIVTERRITDblClick(Sender: TObject);
    procedure ET_PAYSChange(Sender: TObject);
    procedure BDeleteClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
  private    { Déclarations privées }
    Function IsMouvementer : Boolean ;
  public    { Déclarations publiques }
  end;

implementation

uses Devise;

{$R *.DFM}

Procedure FicheEtablissement(Mode : TActionFiche) ;
var FEtablis : TFEtablis ;
    PP       : THPanel ; 
BEGIN
if Blocage(['nrCloture','nrBatch'],True,'nrBatch') then Exit ;
FEtablis:=TFEtablis.Create(Application) ;
FEtablis.InitFL('ET','PRT_ETABLIS','','',Mode,TRUE,
                FEtablis.TaET_ETABLISSEMENT,FEtablis.TaET_LIBELLE,FEtablis.TaET_ETABLISSEMENT,['ttEtablissement']) ;
PP:=FindInsidePanel ;
if PP=Nil then
   BEGIN
   Try
    FEtablis.ShowModal ;
   Finally
    FEtablis.Free ;
    Bloqueur('nrBatch',False) ;
   End ;
   Screen.Cursor:=crDefault ;
   END else
   BEGIN
   InitInside(FEtablis,PP) ;
   FEtablis.Show ;
   END ;
END ;

procedure TFEtablis.ET_CODEPOSTALDblClick(Sender: TObject);
begin
  inherited;
VerifCodePostal(Ta,ET_CODEPOSTAL,ET_VILLE,True) ;
end;

procedure TFEtablis.ET_VILLEDblClick(Sender: TObject);
begin
  inherited;
VerifCodePostal(Ta,ET_CODEPOSTAL,ET_VILLE,False) ;
end;

procedure TFEtablis.ET_DIVTERRITDblClick(Sender: TObject);
begin
  inherited;
PaysRegion(ET_PAYS,ET_DIVTERRIT,True) ;
end;

procedure TFEtablis.ET_PAYSChange(Sender: TObject);
begin
  inherited;
PaysRegion(ET_PAYS,ET_DIVTERRIT,False) ;
end;

procedure TFEtablis.BDeleteClick(Sender: TObject);
begin
if TaET_ETABLISSEMENT.AsString=VH^.EtablisDefaut then BEGIN HM.Execute(6,'','') ; Exit ; END ;
if IsMouvementer then Exit ;
  inherited;
end;

Function TFEtablis.IsMouvementer : Boolean ;
Var QLoc : TQuery ;
    Sql : String ;
BEGIN
Result:=True ;
QLoc:=TQuery.Create(Application) ; QLoc.DataBaseName:='SOC' ; QLoc.Sql.Clear ;
Sql:='Select ET_ETABLISSEMENT From ETABLISS Where ET_ETABLISSEMENT="'+TaET_ETABLISSEMENT.AsString+'" And '+
      '(Exists (Select E_ETABLISSEMENT From ECRITURE Where E_ETABLISSEMENT="'+TaET_ETABLISSEMENT.AsString+'"))' ;
QLoc.Close ; QLoc.Sql.Clear ; QLoc.Sql.Add(Sql) ; ChangeSql(QLoc) ; QLoc.Open ;
if Not QLoc.Eof then BEGIN HM.Execute(7,'','') ; Ferme(QLoc) ; Exit ; END ;

Sql:='Select ET_ETABLISSEMENT From ETABLISS Where ET_ETABLISSEMENT="'+TaET_ETABLISSEMENT.AsString+'" And '+
      '(Exists (Select Y_ETABLISSEMENT From ANALYTIQ Where Y_ETABLISSEMENT="'+TaET_ETABLISSEMENT.AsString+'"))' ;
QLoc.Close ; QLoc.Sql.Clear ; QLoc.Sql.Add(Sql) ; ChangeSql(QLoc) ; QLoc.Open ;
if Not QLoc.Eof then BEGIN HM.Execute(8,'','') ; Ferme(QLoc) ; Exit ; END ;

Sql:='Select ET_ETABLISSEMENT From ETABLISS Where ET_ETABLISSEMENT="'+TaET_ETABLISSEMENT.AsString+'" And '+
      '(Exists (Select BE_ETABLISSEMENT From BUDECR Where BE_ETABLISSEMENT="'+TaET_ETABLISSEMENT.AsString+'"))' ;
QLoc.Close ; QLoc.Sql.Clear ; QLoc.Sql.Add(Sql) ; ChangeSql(QLoc) ; QLoc.Open ;
if Not QLoc.Eof then BEGIN HM.Execute(9,'','') ; Ferme(QLoc) ; Exit ; END ;
Ferme(QLoc) ; Result:=False ;
END ;

procedure TFEtablis.FormShow(Sender: TObject);
begin
  inherited;
if (Ta.Eof) And (Ta.Bof) And (FTypeAction<>taConsult) then
   BEGIN
   if ta.State=dsInsert then NewEnreg else BinsertClick(Nil) ;
   END ;
if ((EstSerie(S3)) or (EstSerie(S5))) then BEGIN ET_LANGUE.Visible:=False ; TET_LANGUE.Visible:=False ; END ;
end;

procedure TFEtablis.FormClose(Sender: TObject; var Action: TCloseAction);
begin
if isInside(Self) then Bloqueur('nrBatch',False) ;
  inherited;
end;

end.
