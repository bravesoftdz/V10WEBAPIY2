unit ScenaCoE;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  FichList, HSysMenu, hmsgbox, DB, DBTables, DBCtrls, StdCtrls, Buttons,
  ExtCtrls, Grids, DBGrids, HDB, Mask, Hctrls, ComCtrls, Ent1, HEnt1, Spin,
  Hqry, HTB97, HPanel, Hcompte, HRegCpte, ADODB, udbxDataset ;

Procedure ParamSupExport(Import,Nature,Format,Code : String ; Comment : TActionFiche ; QuellePage : Integer ;
                         FromImp : Boolean=FALSE);

type
  TFParamExpSup = class(TFFicheListe)
    TNT_NATURE: THLabel;
    TNT_LIBELLE: THLabel;
    FS_LIBELLE: TDBEdit;
    FS_CODE: TDBEdit;
    Pages: TPageControl;
    Tinfos1: TTabSheet;
    TaFS_IMPORT: TStringField;
    TaFS_NATURE: TStringField;
    TaFS_FORMAT: TStringField;
    TaFS_CODE: TStringField;
    TaFS_TRESO: TStringField;
    TaFS_GENATTEND: TStringField;
    TaFS_CLIATTEND: TStringField;
    TaFS_FOUATTEND: TStringField;
    TaFS_SALATTEND: TStringField;
    TaFS_DIVATTEND: TStringField;
    TaFS_SECTION1: TStringField;
    TaFS_SECTION2: TStringField;
    TaFS_SECTION3: TStringField;
    TaFS_SECTION4: TStringField;
    TaFS_SECTION5: TStringField;
    TaFS_COLLCLI: TStringField;
    TaFS_CPTCOLLCLI: TStringField;
    TaFS_COLFOU: TStringField;
    TaFS_CPTCOLLFOU: TStringField;
    TaFS_CORRIMPG: TStringField;
    TaFS_CORRIMPT: TStringField;
    TaFS_CORRIMP1: TStringField;
    TaFS_CORRIMP2: TStringField;
    TaFS_CORRIMP3: TStringField;
    TaFS_CORRIMP4: TStringField;
    TaFS_CORRIMP5: TStringField;
    TaFS_MPDEFAUT: TStringField;
    TaFS_MRDEFAUT: TStringField;
    FS_FILTREGENE: TDBCheckBox;
    Label12: TLabel;
    FS_FILTREAUX: TDBCheckBox;
    Label13: TLabel;
    TaFS_CHEMIN: TStringField;
    TaFS_PREFIXE: TStringField;
    TaFS_SUFFIXE: TStringField;
    TaFS_NOMINTER: TStringField;
    TaFS_NOMREJET: TStringField;
    TaFS_TREGDEFAUT: TStringField;
    TaFS_TMRDEFAUT: TStringField;
    TaFS_FILTREGENE: TStringField;
    TaFS_NATUREGENE: TStringField;
    TaFS_FILTREAUX: TStringField;
    TaFS_NATUREAUX: TStringField;
    FS_NATUREGENE: THDBValComboBox;
    FS_NATUREAUX: THDBValComboBox;
    Label1: TLabel;
    FS_NATMVT: THDBValComboBox;
    TaFS_NATMVT: TStringField;
    FS_FORMAT: THDBValComboBox;
    HLabel1: THLabel;
    TaFS_LIBELLE: TStringField;
    TaFS_TRAITETVA: TStringField;
    TaFS_TRAITECTR: TStringField;
    TaFS_LIBMP: TStringField;
    TaFS_RIBCLIENT: TStringField;
    TaFS_DOUBLON: TStringField;
    TaFS_FORCEPIECE: TStringField;
    TaFS_ANOUVEAUD: TStringField;
    FS_FORMAT2: THValComboBox;
    procedure FormShow(Sender: TObject);
    procedure BCAbandonClick(Sender: TObject);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
    procedure FS_NATMVTChange(Sender: TObject);
  private
    { Déclarations privées }
    Import,Nature,Format,Code     : String ;
    FermeCa,FromImp : Boolean ;
    Procedure Sors ;
  public
    Function  EnregOK : boolean ; Override ;
    Procedure NewEnreg ; Override ;
  public
    { Déclarations publiques }
  end;

implementation

{$R *.DFM}

Uses FmtChoix ;


Procedure ParamSupExport(Import,Nature,Format,Code : String ; Comment : TActionFiche ; QuellePage : Integer ;
                         FromImp : Boolean=FALSE);
Var FParamExpSup : TFParamExpSup ;
    NomPrint : string ;
BEGIN
NomPrint:='PRT_SCEXP' ;
FParamExpSup:=TFParamExpSup.Create(Application) ;
 try
   FParamExpSup.Import:=Import ;
   FParamExpSup.Nature:=Nature ;
   FParamExpSup.Format:=Format ;
   FParamExpSup.Code:=Code ;
   FParamExpSup.FromImp:=FromImp ;
   FParamExpSup.InitFL('FS',NomPrint,'','-',Comment,TRUE,FParamExpSup.TaFS_CODE,FParamExpSup.TaFS_LIBELLE,Nil,['']) ;
   FParamExpSup.ShowModal ;
 Finally
   FParamExpSup.free ;
 End ;
Screen.Cursor:=SyncrDefault ;
END;

procedure TFParamExpSup.FormShow(Sender: TObject);
Var i : Integer ;
begin
ChangeDataTypeFmt(Import,Nature,FS_Format2) ;
FS_FORMAT.Items.Clear ; FS_FORMAT.Values.Clear ;
For i:=0 To FS_FORMAT2.Items.Count-1 Do FS_FORMAT.Items.Add(FS_FORMAT2.Items[i]) ;
For i:=0 To FS_FORMAT2.Values.Count-1 Do FS_FORMAT.Values.Add(FS_FORMAT2.Values[i]) ;
FS_FORMAT.Refresh ; //FS_FORMAT.Reload ;
  inherited;
If Import='X' Then Caption:=HM2.Mess[0] Else Caption:=HM2.Mess[1] ;
Pages.ActivePage:=Pages.Pages[0] ;
FermeCa:=False ;
if (Ta.Eof) And (Ta.Bof) And (FTypeAction<>taConsult) then
   BEGIN
   if ta.State=dsInsert then NewEnreg else BinsertClick(Nil) ;
   END ;
//   if FTypeAction in [taCreat,taCreatEnSerie,taCreatOne] then BinsertClick(Nil) ;
//FS_NATMVT.SetFocus ;
end;

Function TFParamExpSup.EnregOK : boolean ;
BEGIN
result:=Inherited EnregOK  ; if Not Result then Exit ;
Modifier:=True ;
if FromImp then
   BEGIN
   if TA.State=dsBrowse then
      begin
      //TRIB.Locate('R_NUMEROCOMPTE',NumCompte,[]) ;
      while not TA.EOF do
        if (TA.FindField('FS_FORMAT').AsString=Format) And
           ((TA.FindField('FS_CODE').AsString=Code) Or (Code='')) And
           (TA.FindField('FS_IMPORT').AsString=Import) And
           (TA.FindField('FS_NATURE').AsString=Nature) then Break else TA.Next ;
      end ;
   END Else
   BEGIN
   if ((result) and (Ta.state in [dsInsert])) then
      BEGIN
      Result:=False ;
      if PresenceComplexe('FMTSUP',['FS_IMPORT','FS_NATURE','FS_FORMAT','FS_CODE'],['=','=','=','='],[Import,Nature,Format,FS_CODE.Text],['S','S','S','S']) then
         BEGIN HM.Execute(4,'','') ; if FS_CODE.CanFocus then FS_CODE.SetFocus ; Exit ; END ;
      END ;
   END ;
Result:=TRUE  ; Modifier:=False ;
END ;

Procedure TFParamExpSup.NewEnreg ;
BEGIN
Inherited ;
TaFS_IMPORT.AsString:=Import ;
TaFS_NATURE.AsString:=Nature ;
TaFS_FORMAT.AsString:=Format ;
FS_FORMAT.Value:=Format ; FS_FORMAT.Refresh ;
FS_NATMVT.Value:='PIE' ; FS_NATMVT.Refresh ;
END ;

procedure TFParamExpSup.BCAbandonClick(Sender: TObject);
begin
  inherited;
Sors ;
end;

Procedure TFParamExpSup.Sors ;
BEGIN
FermeCa:=True ; Close ;
END ;

procedure TFParamExpSup.FormCloseQuery(Sender: TObject; var CanClose: Boolean);
begin
if FermeCa then CanClose:=True Else inherited;
end;

procedure TFParamExpSup.FS_NATMVTChange(Sender: TObject);
begin
  inherited;
If FS_NATMVT.Value='PIE' Then
  BEGIN
  FS_NATUREGENE.ItemIndex:=0 ; FS_NATUREAUX.ItemIndex:=0 ;
  FS_FILTREGENE.Checked:=FALSE ; FS_FILTREAUX.Checked:=FALSE ;
  END ;
end;

end.
