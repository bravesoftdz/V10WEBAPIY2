// Remplacé en eAGL par REFAUTO_TOM
unit Paramlib;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  FichList, DB, DBTables, StdCtrls, Hctrls, Mask, DBCtrls, hmsgbox, HCompte,
  Buttons, ExtCtrls, Grids, DBGrids, HDB, Ent1, HEnt1, SaisUtil, HSysMenu,
  Hqry, HTB97, HPanel, UiUtil ;

Procedure ParamLibelle ;

type
  TFParamLib = class(TFFicheListe)
    TRA_CODE: TLabel;
    RA_CODE: TDBEdit;
    TRA_LIBELLE: TLabel;
    RA_LIBELLE: TDBEdit;
    TRA_JOURNAL: TLabel;
    RA_JOURNAL: THDBValComboBox;
    TRA_NATUREPIECE: TLabel;
    RA_NATUREPIECE: THDBValComboBox;
    TRA_FORMULELIB: TLabel;
    RA_FORMULELIB: TDBEdit;
    TRA_FORMULEREF: TLabel;
    RA_FORMULEREF: TDBEdit;
    TaRA_CODE: TStringField;
    TaRA_LIBELLE: TStringField;
    TaRA_JOURNAL: TStringField;
    TaRA_NATUREPIECE: TStringField;
    TaRA_FORMULEREF: TStringField;
    TaRA_FORMULELIB: TStringField;
    BAssist: TToolbarButton97;
    procedure RA_JOURNALChange(Sender: TObject);
    procedure BAssistClick(Sender: TObject);
    procedure RA_FORMULELIBEnter(Sender: TObject);
    procedure RA_CODEEnter(Sender: TObject);
    procedure RA_FORMULEREFEnter(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    Function  EnregOK : boolean ; override ;
    Procedure NewEnreg ; Override ;
  private    { Déclarations privées }
    LibFocu,RefFocu : Boolean ;
  public    { Déclarations publiques }
  end;

implementation

Uses GuidTool, UtilPgi ;

{$R *.DFM}

Procedure ParamLibelle ;
var FParamLib : TFParamLib ;
    PP : THPanel ;
BEGIN
if _Blocage(['nrCloture'],False,'nrAucun') then Exit ;
PP:=FindInsidePanel ;
FParamLib:=TFParamLib.Create(Application) ;
FParamLib.InitFL('RA','PRT_LIBAUTO','','',taModif,TRUE,FParamLib.TaRA_CODE,
               FParamLib.TaRA_LIBELLE,FParamLib.TaRA_CODE,['']) ;
FParamLib.LibFocu:=False ;
FParamLib.RefFocu:=False ;
if PP=Nil then
   BEGIN
    Try
     FParamLib.ShowModal ;
    Finally
     FParamLib.Free ;
    End ;
   Screen.Cursor:=crDefault ;
   END else
   BEGIN
   InitInside(FParamLib,PP) ;
   FParamLib.Show ;
   END ;
END ;

Procedure TFParamLib.NewEnreg ;
BEGIN
Inherited ;
RA_NATUREPIECE.Enabled:=TRUE ;
END ;

procedure TFParamLib.RA_JOURNALChange(Sender: TObject);
Var Q : TQuery ;
begin
  inherited;
if Ta.state in [dsEdit,dsInsert] then
   BEGIN
    Q:=OpenSQL('Select J_NATUREJAL FROM JOURNAL WHERE J_JOURNAL="'+RA_JOURNAL.Value+'"',TRUE) ;
    Case CaseNatJal(Q.Fields[0].AsString) of
       tzJVente  : BEGIN RA_NATUREPIECE.DataType:='ttNatPieceVente'  ; RA_NATUREPIECE.Value:='FC' ; END ;
       tzJAchat  : BEGIN RA_NATUREPIECE.DataType:='ttNatPieceAchat'  ; RA_NATUREPIECE.Value:='FF' ; END ;
       tzJBanque : BEGIN RA_NATUREPIECE.DataType:='ttNatPieceBanque' ; RA_NATUREPIECE.Value:='RC' ; END ;
       tzJOD     : BEGIN RA_NATUREPIECE.DataType:='ttNaturePiece'    ; RA_NATUREPIECE.Value:='OD' ; END ;
       END ;
    Ferme(Q) ;  RA_NATUREPIECE.Enabled:=True ;
   END ;
end;

procedure TFParamLib.BAssistClick(Sender: TObject);
Var St,StC : String ;
    DBE : TComponent ;
begin
  inherited;
if (Not LibFocu) And (Not RefFocu) then BEGIN HM.Execute(7,'','') ; Exit ; END ;
St:=ChoixChampZone(0,'LIB') ; if St='' then Exit ;
if Ta.State=dsBrowse then Ta.Edit ;
if LibFocu then DBE:=FindComponent('RA_FORMULELIB')
           else DBE:=FindComponent('RA_FORMULEREF') ;
Stc:=TDBEdit(DBE).Text ;
if Length(Stc+' '+St)>100 then HM.Execute(6,'','') ;
if TDBEdit(DBE).SelLength>0 then Delete(StC,TDBEdit(DBE).SelStart+1,TDBEdit(DBE).SelLength) ;
if StC='' then StC:=St else StC:=StC+' '+St ;
if LibFocu then TaRA_FORMULELIB.AsString:=StC
           else TaRA_FORMULEREF.AsString:=StC ;
end;

procedure TFParamLib.RA_FORMULELIBEnter(Sender: TObject);
begin
  inherited;
LibFocu:=True ; RefFocu:=False ;
end;

procedure TFParamLib.RA_FORMULEREFEnter(Sender: TObject);
begin
  inherited;
LibFocu:=False ; RefFocu:=True ;
end;

procedure TFParamLib.RA_CODEEnter(Sender: TObject);
begin
  inherited;
LibFocu:=False ; RefFocu:=False ;
end;

procedure TFParamLib.FormShow(Sender: TObject);
begin
  inherited;
if (Ta.Eof) And (Ta.Bof) And (FTypeAction<>taConsult) then
   BEGIN
   if ta.State=dsInsert then NewEnreg else BinsertClick(Nil) ;
   END ;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent Gendreau
Créé le ...... : 21/02/2002
Modifié le ... :   /  /    
Description .. : suppression du contrôle d'un libelle unique pour un journal
Mots clefs ... :
*****************************************************************}
Function TFParamLib.EnregOK : boolean ;
//Var QQ : TQuery ;
//    Ex : Boolean ;
BEGIN
Result:=Inherited EnregOK ; if Not Result then Exit ;
{
Ex:=False ;
if Ta.State in [dsInsert,dsEdit] then
   BEGIN
   QQ:=OpenSQL('SELECT * FROM REFAUTO WHERE RA_JOURNAL="'+RA_JOURNAL.Value+'" AND RA_NATUREPIECE="'+RA_NATUREPIECE.Value+'" AND RA_CODE<>"'+RA_CODE.Text+'"',True) ;
   Ex:=(Not QQ.EOF) ;
   Ferme(QQ) ;
   if Ex then BEGIN Result:=False ; HM.Execute(8,'','') ; END ;
   END ;}
END ;

procedure TFParamLib.FormCreate(Sender: TObject);
begin
  inherited;
RA_JOURNAL.Style:=csDropDownList ;
end;

end.
