{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... :
Modifié le ... : 03/03/2004
Description .. :
Suite ........ : GCO - 03/03/2004
Suite ........ : -> Uniformisation de l'appel à FicheJournal en 2/3 et CWAS
Mots clefs ... : 
*****************************************************************}
unit MulJal;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Mul, Menus, DB,
  {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
  Hqry, Grids, DBGrids, StdCtrls, Hctrls, Ent1,
  ExtCtrls, ComCtrls, Buttons, Mask, HEnt1, hmsgbox,
  CPJournal_TOM,
  HRichEdt,
  HSysMenu, HDB, HTB97, ColMemo, HPanel, UiUtil, HRichOLE, HPop97,
  SOUCHE_TOM, ADODB
  ;

Procedure MulticritereJournal(Comment : TActionFiche);

type
  TFMulJal = class(TFMul)
    TJ_JOURNAL: THLabel;
    J_JOURNAL: TEdit;
    TJ_LIBELLE: THLabel;
    J_LIBELLE: TEdit;
    NePasVirer: TComboBox;
    J_NATUREJAL: THValComboBox;
    TJ_NATUREJAL: THLabel;
    HM: THMsgBox;
    TJ_DATEMODIFICATION: THLabel;
    J_DATEMODIF: THCritMaskEdit;
    TJ_DATEMODIFICATION2: THLabel;
    J_DATEMODIF_: THCritMaskEdit;
    TJ_DATEDERNMVT: THLabel;
    J_DATEDERNMVT: THCritMaskEdit;
    TJ_DATEDERNMVT2: THLabel;
    J_DATEDERNMVT_: THCritMaskEdit;
    J_CENTRALISABLE: TCheckBox;
    TJ_ABREGE: THLabel;
    J_ABREGE: TEdit;
    PopMenu1: TPopupMenu;
    popCompteurs: TMenuItem;
    procedure FormShow(Sender: TObject);
    procedure FListeDblClick(Sender: TObject); override;
    procedure BOuvrirClick(Sender: TObject); override;
    procedure BinsertClick(Sender: TObject);
    procedure popCompteursClick(Sender: TObject);
  private
    { Déclarations privées }
  public
    { Déclarations publiques }
  end;

implementation

{$R *.DFM}

Uses UtilPGI ,UFonctionsCBP;

Procedure MulticritereJournal(Comment : TActionFiche);
var FMulJal: TFMulJal;
    PP       : THPanel ;
begin
if Comment<>taConsult then if _Blocage(['nrCloture'],False,'nrAucun') then Exit ;
FMulJal:=TFMulJal.Create(Application) ;
FMulJal.TypeAction:=Comment ;
Case Comment Of
  taConsult : begin
              FMulJal.Caption:=FMulJal.HM.Mess[0] ;
              FMulJal.FNomFiltre:='MULVJAL' ;
              FMulJal.Q.Liste:='MULVJOURNAL' ;
              FMulJal.HelpContext:=7205000 ;
              end ;
    taModif : begin
              FMulJal.Caption:=FMulJal.HM.Mess[1] ;
              FMulJal.FNomFiltre:='MULMJAL' ;
              FMulJal.Q.Liste:='MULMJOURNAL' ;
              FMulJal.HelpContext:=7211000 ;
              end ;
  end ;
if ((EstSerie(S5)) or (EstSerie(S3))) then FMulJal.Caption:=FMulJal.HM.Mess[4] ;
PP:=FindInsidePanel ;
if PP=Nil then
   BEGIN
    try
     FMulJal.ShowModal ;
    finally
     FMulJal.Free ;
    end ;
   Screen.Cursor:=SyncrDefault ;
   END else
   BEGIN
   InitInside(FMulJal,PP) ;
   FMulJal.Show ;
   END ;
end;


procedure TFMulJal.FormShow(Sender: TObject);
begin
Pages.Pages[1].TabVisible:=FALSE ;
if FCompte<>'' then J_JOURNAL.text:=FCompte ;
if FLibelle<>'' then J_Libelle.text:=FLibelle ;
J_DATEMODIF.text:=StDate1900 ; J_DATEMODIF_.text:=StDate2099 ;
J_DATEDERNMVT.text:=StDate1900 ; J_DATEDERNMVT_.text:=StDate2099 ;
inherited;
if ((V_PGI.OutLook) and (ExJaiLeDroitConcept(TConcept(ccJalCreat),False)) and (TypeAction<>taConsult)) then BInsert.Visible:=True ;


end;

procedure TFMulJal.FListeDblClick(Sender: TObject);
var Q1 : TQuery ;
    AA : TActionFiche ;
begin
if(Q.Eof)And(Q.Bof) then Exit ;
  inherited;
AA:=TypeAction ;
if Zoom then
   BEGIN
   if (TypeAction=taModif) then
     BEGIN
     Q1:=OpenSQL('SELECT J_FERME FROM JOURNAL WHERE J_JOURNAL="'+Q.FindField('J_JOURNAL').AsString+'"',True) ;
     if Q1.Fields[0].AsString='X'then HM.Execute(3,'','') ;
     Ferme(Q1) ;
     END ;
   if ((TypeAction=taModif) and (Not ExJaiLeDroitConcept(TConcept(ccGenModif),False))) then AA:=taConsult ;
   if V_PGI.MonoFiche then Fichejournal(nil,'',Q.FindField('J_JOURNAL').AsString,AA,0)
                   else Fichejournal(Q,'',Q.FindField('J_JOURNAL').AsString,AA,0) ;
   if typeaction<>taConsult then BChercheClick(Nil) ;
   END else FCompte:=Q.FindField('J_JOURNAL').AsString ;
Screen.Cursor:=SyncrDefault ;
end;

procedure TFMulJal.BOuvrirClick(Sender: TObject);
begin
  inherited;
FListeDblClick(Sender) ;
end;

procedure TFMulJal.BinsertClick(Sender: TObject);
begin
  inherited;
if V_PGI.OutLook then FicheJournal(Nil,'','',taCreatEnSerie,0) ;
end;

procedure TFMulJal.popCompteursClick(Sender: TObject);
begin
  inherited;
  YYLanceFiche_Souche('CPT','','ACTION=MODIFICATION;CPT'); //Compteurs
end;

end.
