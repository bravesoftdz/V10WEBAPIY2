unit MulJalSI;

interface

uses
    Windows,
    Messages,
    SysUtils,
    Classes,
    Graphics,
    Controls,
    Forms,
    Dialogs,
    Mul,
    Menus,
    DB,
    Hqry,
    Grids,
    DBGrids,
    StdCtrls,
    Hctrls,
    Ent1,
    ExtCtrls,
    ComCtrls,
    Buttons,
    Mask,
    HEnt1,
    hmsgbox,
    CPJOURNAL_TOM,
    HRichEdt,
    HSysMenu,
    HDB,
    HTB97,
    ColMemo,
    HPanel,
    UiUtil,
    HRichOLE, ADODB, udbxDataset
    ;

Procedure MulticritereJournalSISCO(Comment : TActionFiche);
Function VerifJalSISCO(AvecMess : Boolean) : Boolean ;

type
  TFMulJalSI = class(TFMul)
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
    XX_WHERE: TEdit;
    procedure FormShow(Sender: TObject);
    procedure FListeDblClick(Sender: TObject);
    procedure BOuvrirClick(Sender: TObject);
    procedure BinsertClick(Sender: TObject);
  private
    { Déclarations privées }
  public
    { Déclarations publiques }
  end;

implementation

{$R *.DFM}

Procedure MulticritereJournalSISCO(Comment : TActionFiche);
var FMulJal: TFMulJalSI;
    PP       : THPanel ;
begin
if Comment<>taConsult then if Blocage(['nrCloture'],False,'nrAucun') then Exit ;
FMulJal:=TFMulJalSI.Create(Application) ;
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


procedure TFMulJalSI.FormShow(Sender: TObject);
begin
Pages.Pages[1].TabVisible:=FALSE ;
if FCompte<>'' then J_JOURNAL.text:=FCompte ;
if FLibelle<>'' then J_Libelle.text:=FLibelle ;
J_DATEMODIF.text:=StDate1900 ; J_DATEMODIF_.text:=StDate2099 ;
J_DATEDERNMVT.text:=StDate1900 ; J_DATEDERNMVT_.text:=StDate2099 ;
inherited;
if ((V_PGI.OutLook) and (ExJaiLeDroitConcept(TConcept(ccJalCreat),False)) and (TypeAction<>taConsult)) then BInsert.Visible:=True ;
end;

procedure TFMulJalSI.FListeDblClick(Sender: TObject);
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

procedure TFMulJalSI.BOuvrirClick(Sender: TObject);
begin
  inherited;
FListeDblClick(Sender) ;
end;

procedure TFMulJalSI.BinsertClick(Sender: TObject);
begin
  inherited;
if V_PGI.OutLook then FicheJournal(Nil,'','',taCreatEnSerie,0) ;
end;

Function VerifJalSISCO(AvecMess : Boolean) : Boolean ;
Var Q : TQuery ;
    OkOk : Boolean ;
    OkAch,OkVte,OkBqe,OkCai,OkAno : Boolean ;
    St : String ;
BEGIN
Result:=TRUE ; OkOk:=TRUE ; OkAch:=FALSE ; OkVte:=FALSE ; OkBqe:=FALSE ; OkCai:=FALSE ; OkAno:=FALSE ; 
Q:=OpenSQL('SELECT * FROM JOURNAL WHERE J_CREERPAR="RCS"',FALSE) ;
While Not Q.Eof Do
  BEGIN
  If (Q.FindField('J_COMPTEURNORMAL').AsString='') Then OkOk:=FALSE ;
  If Q.FindField('J_NATUREJAL').AsString='VTE' Then OkVte:=TRUE ;
  If Q.FindField('J_NATUREJAL').AsString='ACH' Then OkAch:=TRUE ;
  If Q.FindField('J_NATUREJAL').AsString='BQE' Then OkBqe:=TRUE ;
  If Q.FindField('J_NATUREJAL').AsString='CAI' Then OkCai:=TRUE ;
  If Q.FindField('J_NATUREJAL').AsString='ANO' Then OkAno:=TRUE ;
  Q.Next ;
  END ;
Ferme(Q) ;
St:='' ;
If Not OkVte Then St:=St+'Vente / ' ;
If Not OkAch Then St:=St+'Achat / ' ;
If Not OkBqe Then St:=St+'Banque / ' ;
If Not OkCai Then St:=St+'Caisse / ' ;
If Not OkAno Then St:=St+'A-Nouveaux / ' ;
If St<>'' Then Delete(St,length(St)-2,2) ;
If AvecMess Then
  BEGIN
  If St<>'' Then HShowMessage('0;Journaux importés de SISCO II;Attention : il n''y a pas de journal de nature '+St+';E;O;O;O;','','') ;
  END ;
END ;

end.
