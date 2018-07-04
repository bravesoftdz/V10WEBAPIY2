{***********UNITE*************************************************
Auteur  ...... : Vincent Laroche
Créé le ...... : 28/01/2005
Modifié le ... :   /  /    
Description .. : Remplacé en eAGL par MULVALBU_TOF.PAS
Mots clefs ... : 
*****************************************************************}
unit MulValBu;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Mul, HSysMenu, Menus, DB, DBTables, Hqry, StdCtrls, Grids, DBGrids, HDB,
  ComCtrls, HRichEdt, Hctrls, ExtCtrls, Buttons, Hcompte, Mask, hmsgbox,
  Ent1, Hent1, SaisBud, SaisUtil, HTB97, ColMemo, HPanel, UiUtil, HRichOLE ;

Procedure MultiCritereValBud(AValider : Boolean) ;

type
  TFMulValBu = class(TFMul)
    HM: THMsgBox;
    BE_BUDJAL: THValComboBox;
    TBE_BUDJAL: TLabel;
    BE_DATECOMPTABLE: THCritMaskEdit;
    BE_DATECOMPTABLE_: THCritMaskEdit;
    TBE_DATECOMPTABLE2: THLabel;
    TBE_DATECOMPTABLE: THLabel;
    TBE_EXERCICE: THLabel;
    BE_EXERCICE: THValComboBox;
    BE_AXE: THValComboBox;
    TBE_AXE: TLabel;
    TBE_NATUREBUD: TLabel;
    BE_NATUREBUD: THValComboBox;
    TBE_QUALIFPIECE: THLabel;
    BE_QUALIFPIECE: THValComboBox;
    TBE_UTILISATEUR: THLabel;
    BE_UTILISATEUR: THValComboBox;
    TBE_ETABLISSEMENT: TLabel;
    BE_ETABLISSEMENT: THValComboBox;
    BE_REFINTERNE: TEdit;
    TBE_REFINTERNE: THLabel;
    TBE_GENERAL: THLabel;
    BE_BUDGENE: THCpteEdit;
    BE_VALIDE: TCheckBox;
    CbJal: TComboBox;
    TBE_NUMEROPIECE: TLabel;
    BE_NUMEROPIECE: THCritMaskEdit;
    TBE_NUMEROPIECE_: TLabel;
    BE_NUMEROPIECE_: THCritMaskEdit;
    procedure FormShow(Sender: TObject);
    procedure FListeDblClick(Sender: TObject); override;
    procedure BOuvrirClick(Sender: TObject); override;
    procedure QBeforeOpen(DataSet: TDataSet);
  private
    AValider : Boolean ;
    NbLig : Integer ;
    Procedure RempliCbJal ;
    Procedure LanceTraitement ;
  public
  end;


implementation

{$R *.DFM}

Procedure MultiCritereValBud(AValider : Boolean) ;
var FMulValBu : TFMulValBu ;
    PP : THPanel ;
BEGIN
FMulValBu:=TFMulValBu.Create(Application) ;
if Avalider then
   BEGIN
   FMulValBu.Caption:=FMulValBu.HM.Mess[0] ;
   FMulValBu.BOuvrir.Hint:=FMulValBu.HM.Mess[2] ;
   FMulValBu.BE_VALIDE.State:=cbUnChecked ;
   FMulValBu.HelpContext:=15250000 ;
   END else
   BEGIN
   FMulValBu.Caption:=FMulValBu.HM.Mess[1] ;
   FMulValBu.BOuvrir.Hint:=FMulValBu.HM.Mess[3] ;
   FMulValBu.BE_VALIDE.State:=cbChecked ;
   FMulValBu.HelpContext:=15260000 ;
   END ;
FMulValBu.FNomFiltre:='MULVALBUD' ; FMulValBu.Q.Liste:='MULVALBUD' ;
FMulValBu.Q.Manuel:=TRUE ; FMulValBu.AValider:=Avalider ;
PP:=FindInsidePanel ;
if PP=Nil then
   BEGIN
    try
     FMulValBu.ShowModal ;
    finally
     FMulValBu.Free ;
    end;
   Screen.Cursor:=SyncrDefault ;
   END else
   BEGIN
   InitInside(FMulValBu,PP) ;
   FMulValBu.Show ;
   END ;
END ;

procedure TFMulValBu.FormShow(Sender: TObject);
begin
BE_DATECOMPTABLE.text:=StDate1900 ; BE_DATECOMPTABLE_.text:=StDate2099 ;
BE_BUDJAL.ItemIndex:=0 ; BE_NATUREBUD.ItemIndex:=0 ; RempliCbJal ;
PositionneEtabUser(BE_ETABLISSEMENT) ;
  inherited;
end;

procedure TFMulValBu.FListeDblClick(Sender: TObject);
begin
  inherited;
if (Q.Eof) And (Q.Bof) then Exit ;
TrouveEtLanceSaisBud(Q,taConsult) ;
end;

procedure TFMulValBu.BOuvrirClick(Sender: TObject);
begin
  inherited;
NbLig:=Fliste.NbSelected ;
if NbLig<=0 then BEGIN HM.Execute(Ord(AValider)+9,'','') ; Exit ; END ;
if HM.Execute(Ord(AValider)+4,'','')=mrYes then
   BEGIN LanceTraitement ; BChercheClick(Nil) ; END ;
end;

Procedure TFMulValBu.LanceTraitement ;
Var Flag : Char ;
    i : Integer ;
    QLoc : TQuery ;
    Sql : String ;
BEGIN
if AValider then Flag:='X' else Flag:='-' ;
QLoc:=TQuery.Create(Self) ; QLoc.DataBaseName:='SOC' ;
for i:=0 to NbLig-1 do
    BEGIN
    Fliste.GotoLeBookMark(i) ;
    if CbJal.Items.IndexOf(Q.FindField('BE_BUDJAL').AsString)<>-1 then Continue ;
    Sql:='UPDATE BUDECR SET BE_VALIDE="'+Flag+'" '+
         'Where BE_BUDJAL="'+Q.FindField('BE_BUDJAL').AsString+'" '+
         'And BE_NATUREBUD="'+Q.FindField('BE_NATUREBUD').AsString+'" '+
         'And BE_NUMEROPIECE='+IntToStr(Q.FindField('BE_NUMEROPIECE').AsInteger)+' '+
//         'And BE_DATECOMPTABLE="'+UsDateTime(Q.FindField('BE_DATECOMPTABLE').AsDateTime)+'" '+
         'And BE_QUALIFPIECE="'+Q.FindField('BE_QUALIFPIECE').AsString+'" ' ;
    QLoc.Close ; QLoc.Sql.Clear ; QLoc.Sql.Add(Sql) ; ChangeSql(QLoc) ; QLoc.ExecSql ;
    END ;
Ferme(QLoc) ;
END ;

Procedure TFMulValBu.RempliCbJal ;
Var QLoc : TQuery ;
BEGIN
CbJal.Items.Clear ;
QLoc:=OpenSql('Select BJ_BUDJAL From BUDJAL Where BJ_FERME="X"',True) ;
While Not QLoc.Eof do
    BEGIN
    CbJal.Items.Add(QLoc.Fields[0].AsString) ; QLoc.Next ;
    END ;
Ferme(QLoc) ;
END ;

procedure TFMulValBu.QBeforeOpen(DataSet: TDataSet);
begin
  inherited;
  Q.SQL.Text := ModifieRequeteGrille(Q.SQL.Text); {FP FQ16057}
end;
end.
