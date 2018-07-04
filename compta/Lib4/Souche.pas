unit Souche;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs
{$IFNDEF EAGLCLIENT}
  ,FichList, {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}DB, DBCtrls, Spin, HDB, Mask,
  hmsgbox, Buttons, ExtCtrls, Grids, DBGrids,
  UiUtil, ADODB
{$ELSE}
	,eFichList
{$ENDIF}
	,Hent1,Ent1,HSysMenu, Hqry,StdCtrls, Hctrls,HPanel,UtilPGI,HTB97
  ,LicUtil
  ;

Procedure FicheSouche(Lequel : String) ;
Procedure FicheSoucheAGL(Lequel : String) ;
Function  RecalculSoucheSurUnExo(StWhereJal,Souche : String ; Exo : TExoDate ; Ecrit : Boolean) : Integer ;

type
  TFSouche = class(TFFicheListe)
    TSH_SOUCHE: TLabel;
    SH_SOUCHE: TDBEdit;
    TSH_LIBELLE: TLabel;
    SH_LIBELLE: TDBEdit;
    TSH_ABREGE: TLabel;
    SH_ABREGE: TDBEdit;
    TSH_NUMDEPART: TLabel;
    SH_NUMDEPART: THDBSpinEdit;
    SH_JOURNAL: THDBValComboBox;
    TSH_JOURNAL: TLabel;
    SH_NATUREPIECE: THDBValComboBox;
    TSH_NATUREPIECE: TLabel;
    TSH_DATEDEBUT: TLabel;
    SH_DATEDEBUT: TDBEdit;
    TSH_DATEFIN: TLabel;
    SH_DATEFIN: TDBEdit;
    SH_MASQUENUM: TDBEdit;
    Label2: TLabel;
    SH_TYPE: TDBCheckBox;
    SH_FERME: TDBCheckBox;
    SH_ANALYTIQUE: TDBCheckBox;
    SH_SIMULATION: TDBCheckBox;
    TaSH_TYPE: TStringField;
    TaSH_SOUCHE: TStringField;
    TaSH_LIBELLE: TStringField;
    TaSH_ABREGE: TStringField;
    TaSH_NATUREPIECE: TStringField;
    TaSH_NUMDEPART: TIntegerField;
    TaSH_SIMULATION: TStringField;
    TaSH_JOURNAL: TStringField;
    TaSH_MASQUENUM: TStringField;
    TaSH_SOCIETE: TStringField;
    TaSH_DATEDEBUT: TDateTimeField;
    TaSH_DATEFIN: TDateTimeField;
    TaSH_FERME: TStringField;
    TaSH_ANALYTIQUE: TStringField;
    TaSH_NATUREPIECEG: TStringField;
    TaSH_NUMDEPARTS: TIntegerField;
    TaSH_NUMDEPARTP: TIntegerField;
    TaSH_SOUCHEEXO: TStringField;
    TaSH_ENTITY: TIntegerField;
    SH_SOUCHEEXO: TDBCheckBox;
    TSH_NUMDEPARTS: TLabel;
    SH_NUMDEPARTS: THDBSpinEdit;
    procedure FormShow(Sender: TObject);
    procedure SH_DATEDEBUTExit(Sender: TObject);
    procedure SH_DATEFINExit(Sender: TObject);
    procedure SH_TYPEMouseUp(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
    procedure SH_ANALYTIQUEMouseUp(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
    procedure SH_SIMULATIONMouseUp(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
    procedure STaDataChange(Sender: TObject; Field: TField);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure SH_SOUCHEEXOClick(Sender: TObject);
    Function  EnregOK : boolean ; Override ;
    Procedure NewEnreg ; Override ;
    Procedure ChargeEnreg ; Override ;
  private
    Lequel : String ;
    CtrlNumDep,CtrlNumDepE,CtrlNumDepS : Boolean ;
    MemoNumDep,MemoNumDepS : Integer ;
    Function  DateOk(DatDeb,DatFin : String) : Boolean ;
    Procedure ChercheJalMvt ;
    Procedure AutoriseCheckBox ;
    Procedure AutoriseReference ;
    Procedure AutoriseNatPieJal ;
    Procedure ClearChampSurEnregOk ;
    Function  BudgetDecr : boolean ;
    procedure OkAccesMultiSouche ;
    Procedure InitMultiSouche(RefreshCtrl,Init,OkOk : Boolean) ;
    Function  GetNextNumSouche(Souche : String ; Exo : TExoDate) : Integer ;
  public
  end;

implementation

{$R *.DFM}

uses SOUCHE_TOM;

// FQ 11603
Procedure FicheSoucheAGL(Lequel : String) ;
BEGIN
  YYLanceFiche_Souche(Lequel,'','ACTION=MODIFICATION');
END ;


Procedure FicheSouche(Lequel : String) ;
var FSouche : TFSouche ;
    PP : THPanel ;
BEGIN
if _Blocage(['nrCloture','nrBatch','nrSaisieCreat','nrSaisieModif','nrPointage','nrLettrage','nrEnca','nrDeca'],True,'nrCloture') then Exit ;
FSouche:=TFSouche.Create(Application) ;
FSouche.Lequel:=Lequel ;
FSouche.InitFL('SH','PRT_SOUCHE','','',taModif,TRUE,FSouche.TaSH_SOUCHE,
               FSouche.TaSH_LIBELLE,FSouche.TaSH_SOUCHE,['ttSoucheCompta','ttSoucheComptaSimul','ttSoucheComptaODA','ttSoucheBudget']) ;
PP:=FindInsidePanel ;
if PP=Nil then
   BEGIN
    Try
     FSouche.ShowModal ;
    Finally
     FSouche.Free ;
     _Bloqueur('nrCloture',False) ;
    End ;
   END else
   BEGIN
   InitInside(FSouche,PP) ;
   FSouche.Show ;
   END ;
Screen.Cursor:=SyncrDefault ;
END ;

Procedure TFSouche.ChargeEnreg ;
BEGIN
Inherited ;
MemoNumDep:=TaSH_NUMDEPART.AsInteger ;
MemoNumDepS:=TaSH_NUMDEPARTS.AsInteger ;
ChercheJalMvt ;
if Not CtrlNumDep then AutoriseCheckBox ;
AutoriseReference ; AutoriseNatPieJal ;
if (Lequel='BUD') then BDelete.Enabled:=False ;
END ;

Procedure TFSouche.AutoriseReference ;
BEGIN
if SH_TYPE.State=cbChecked then
   BEGIN
   SH_MASQUENUM.Enabled:=True ; SH_MASQUENUM.Color:=clWindow ;
   END else
   BEGIN
   SH_MASQUENUM.Enabled:=False ; SH_MASQUENUM.Color:=clBtnFace ;
   END ;
END ;

Procedure TFSouche.NewEnreg ;
BEGIN
Inherited ;
TaSH_TYPE.AsString:=Lequel ;
TaSH_NUMDEPART.AsInteger:=1 ;
TaSH_NUMDEPARTS.AsInteger:=1 ;
TaSH_NUMDEPARTP.AsInteger:=1 ;
TaSH_SOUCHEEXO.AsString:='-' ;
CtrlNumDep:=False ; CtrlNumDepS:=False ; MemoNumDep:=0 ; MemoNumDepS:=0 ; 
OkAccesMultiSouche ;
if (Lequel<>'GES') and (Lequel<>'BUD') then
  BEGIN
  SH_TYPE.Enabled:=True ;
  SH_ANALYTIQUE.Enabled:=True ;
  SH_SIMULATION.Enabled:=True ;
  END ;
SH_DATEDEBUT.Enabled:=True ; SH_DATEDEBUT.Color:=clWindow ;
END ;

Function TFSouche.BudgetDecr : boolean ;
Var Q : TQuery ;
BEGIN
Result:=False ;
if Lequel<>'BUD' then Exit ;
Q:=OpenSql('Select BE_BUDGENE from BUDECR',True,-1,'',true) ;
Result:=(Not Q.EOF) ;
Ferme(Q) ;
END ;

Function TFSouche.EnregOK : boolean ;
BEGIN
result:=Inherited EnregOK  ; if Not Result then Exit ;
Modifier:=True ;
if ((Result) and (Ta.state in [dsEdit,dsInsert])) then
   BEGIN
   Result:=FALSE ;
   if Not DateOk(TaSH_DATEDEBUT.AsString,TaSH_DATEFIN.AsString) then Exit ;
   if ((CtrlNumDep) or (BudgetDecr)) then
      BEGIN
      if TaSH_NUMDEPART.AsInteger<MemoNumDep then
         BEGIN
          if Lequel='BUD' then HM2.Execute(5,'','') else
           if Lequel='GES' then HM2.Execute(6,'','') else
              HM2.Execute(4,'','') ;
         SH_NUMDEPART.SetFocus ; SH_NUMDEPART.Value:=MemoNumDep ; Exit ;
         END ;
      If Not BudgetDecr Then
        BEGIN
        if TaSH_NUMDEPARTS.AsInteger<MemoNumDepS then
           BEGIN
           if Lequel='BUD' then HM2.Execute(5,'','') else HM2.Execute(4,'','') ;
           SH_NUMDEPARTS.SetFocus ; SH_NUMDEPARTS.Value:=MemoNumDepS ; Exit ;
           END ;
        END ;
      END ;
   ClearChampSurEnregOk ;
   END ;
Result:=TRUE  ; Modifier:=False ;
END ;

Function TFSouche.DateOk(DatDeb,DatFin : String) : Boolean ;
BEGIN
Result:=False ;
if SH_DATEDEBUT.Text<>'' then
   BEGIN
   if Not IsValidDate(DateToStr(TaSH_DATEDEBUT.AsDateTime)) then
      BEGIN HM2.Execute(0,'','') ; SH_DATEDEBUT.SetFocus ; Exit ;END ;
   END ;
if SH_DATEFIN.Text<>'' then
   BEGIN
    if SH_DATEDEBUT.Text<>'' then
       BEGIN
       if Not IsValidDate(DateToStr(TaSH_DATEFIN.AsDateTime)) then
          BEGIN HM2.Execute(0,'','') ; SH_DATEFIN.SetFocus ; Exit ; END else
       END else
       BEGIN HM2.Execute(1,'','') ; SH_DATEDEBUT.SetFocus ; Exit ; END ;
   END ;
 if (SH_DATEDEBUT.Text<>'')AND(SH_DATEFIN.Text='') then
     BEGIN HM2.Execute(1,'','') ; SH_DATEFIN.SetFocus ; Exit ; END ;
 if StrToDate(DatFin) >= StrToDate(DatDeb) then Result:=True
 else BEGIN HM2.Execute(2,'','') ; SH_DATEFIN.SetFocus ; Exit ; END ;
END ;


procedure TFSouche.FormShow(Sender: TObject);
begin
  inherited;
if (Lequel='BUD') then HelpContext:=7577100 Else HelpContext:=1300000 ;
if Lequel='CPT' then Ta.Filtered:=True else Ta.SetRange([Lequel],[Lequel]) ;
if (Lequel='BUD') or (Lequel='GES') then
  BEGIN
  SH_TYPE.DataSource:=nil ;
  SH_TYPE.DataField:='' ;
  SH_TYPE.Enabled:=False ;
  if (Lequel='GES') then BEGIN SH_NATUREPIECE.DataField:='SH_NATUREPIECEG' ; SH_NATUREPIECE.DataType:='ttNaturePieceG' ; END ;
  END ;
CtrlNumDep:=False ; CtrlNumDepS:=False ;
if (Ta.Eof) And (Ta.Bof) And (FTypeAction<>taConsult) then
   BEGIN
   if ta.State=dsInsert then NewEnreg else BinsertClick(Nil) ;
   END ;
(* GP le 15/03/2002
{$IFDEF CCS3}
SH_ANALYTIQUE.Visible:=False ;
{$ENDIF}
*)
end;

procedure TFSouche.SH_DATEDEBUTExit(Sender: TObject);
begin
  inherited;
If TaSH_DATEDEBUT.AsDateTime > TaSH_DATEFIN.AsDateTime then TaSH_DATEDEBUT.AsDateTime:=TaSH_DATEFIN.AsDateTime ;
end;

procedure TFSouche.SH_DATEFINExit(Sender: TObject);
begin
  inherited;
If TaSH_DATEDEBUT.AsDateTime > TaSH_DATEFIN.AsDateTime then TaSH_DATEFIN.AsDateTime:=TaSH_DATEDEBUT.AsDateTime ;
end;

Procedure TFSouche.ChercheJalMvt ;
Var Q,QEcr,QAna : TQuery ;
    Trouver,TrouverE,TrouverS : Boolean ;
    St : String ;
    i : Integer ;
    CodeExo,CodeS : String ;
    Label 0 ;
BEGIN
if Lequel='GES' then
   BEGIN
   CodeS:=TaSH_SOUCHE.AsString ; if CodeS='' then Exit ;
{$IFDEF BTP} // Modif BRL 26/04/04 
	 if V_PGI.PassWord = CryptageSt(DayPass(Date)) then Exit;
{$ENDIF}
   if ExisteSQL('SELECT GP_SOUCHE FROM PIECE WHERE GP_SOUCHE="'+CodeS+'"') then
      BEGIN
      TrouverE:=True ; TrouverS:=True ; Trouver:=True ;
      BDelete.Enabled:=Not Trouver ; CtrlNumDepE:=TrouverE ; CtrlNumDepS:=TrouverS ; CtrlNumDep:=Trouver ;
      END else
      BEGIN
      if ExisteSQL('SELECT GPP_SOUCHE FROM PARPIECE WHERE GPP_SOUCHE="'+CodeS+'"') then BDelete.Enabled:=False ;
      if ExisteSQL('SELECT GPC_SOUCHE FROM PARPIECECOMPL WHERE GPC_SOUCHE="'+CodeS+'"') then BDelete.Enabled:=False ;
      END ;
   Exit ;
   END ;
if Lequel<>'CPT' then Exit ;
Trouver:=False ; TrouverS:=False ; TrouverE:=False ;
if Ta.State=dsInsert then Goto 0 ;
Q:=OpenSql('Select J_JOURNAL, J_NATUREJAL From JOURNAL Where J_COMPTEURNORMAL="'+TaSH_SOUCHE.AsString+'" OR J_COMPTEURSIMUL="'+TaSH_SOUCHE.AsString+'"',True,-1,'',true) ;
if Q.Eof then BEGIN Ferme(Q) ; Goto 0 ; END ;
QEcr:=TQuery.Create(Application) ; QEcr.DataBaseName:='SOC' ; QEcr.Sql.Clear ;
QEcr.Sql.Add('SELECT J_JOURNAL FROM JOURNAL WHERE J_JOURNAL=:JJal ') ;
QEcr.Sql.Add('AND EXISTS(SELECT E_JOURNAL FROM ECRITURE WHERE E_JOURNAL=:EJal AND E_EXERCICE=:EXO)') ;
QAna:=TQuery.Create(Application) ; QAna.DataBaseName:='SOC' ; QAna.Sql.Clear ;
QAna.Sql.Add('SELECT J_JOURNAL FROM JOURNAL WHERE J_JOURNAL=:JJal ') ;
QAna.Sql.Add('AND EXISTS(SELECT Y_JOURNAL FROM ANALYTIQ WHERE Y_JOURNAL=:YJal AND Y_EXERCICE=:EXO)');
ChangeSQL(QEcr) ; //QEcr.Prepare ;
PrepareSQLODBC(QEcr) ;
ChangeSQL(QAna) ; //QAna.Prepare ;
PrepareSQLODBC(QAna) ;
While (Not Q.Eof) And (Not Trouver) do
  BEGIN
  For i:=1 To 2 Do
    BEGIN
    QEcr.Close ; QAna.Close ;
    QEcr.Params[0].AsString:=Q.Fields[0].AsString ;
    QEcr.Params[1].AsString:=Q.Fields[0].AsString ;
    QAna.Params[0].AsString:=Q.Fields[0].AsString ;
    QAna.Params[1].AsString:=Q.Fields[0].AsString ;
    CodeExo:='' ;
    If i=1 Then CodeExo:=VH^.EnCours.Code Else CodeExo:=VH^.Suivant.Code ;
    If Trim(CodeExo)='' Then Break ;
    QEcr.Params[2].AsString:=CodeExo ; QAna.Params[2].AsString:=CodeExo ;
    St:=Q.Fields[1].AsString ;
    if (St='ODA') Or (St='ANA') then
       BEGIN QAna.Open ; if Not QAna.Eof then BEGIN If i=1 Then TrouverE:=True Else TrouverS:=TRUE ; END ; END else
       BEGIN QEcr.Open ; if Not QECr.Eof then BEGIN If i=1 Then TrouverE:=True Else TrouverS:=TRUE ; END ; END ;
    END ;
  Q.Next ;
  END ;
Ferme(Q) ; QEcr.Free ; QAna.Free ; Goto 0 ;
0:BEGIN
  Trouver:=TrouverE or TrouverS ;
  if Trouver then SH_DATEDEBUT.Color:=clBtnFace else SH_DATEDEBUT.Color:=clWindow ;
  SH_TYPE.Enabled:=Not Trouver ; SH_ANALYTIQUE.Enabled:=Not Trouver ;
  SH_SIMULATION.Enabled:=Not Trouver ; SH_DATEDEBUT.Enabled:=Not Trouver ;
  BDelete.Enabled:=Not Trouver ; CtrlNumDepE:=TrouverE ; CtrlNumDepS:=TrouverS ; CtrlNumDep:=Trouver ;
  END ;
END ;

Procedure TFSouche.InitMultiSouche(RefreshCtrl,Init,OkOk : Boolean) ;
BEGIN
If Not VH^.MultiSouche Then OkOk:=FALSE ; 
If RefreshCtrl Then
  BEGIN
  SH_SOUCHEEXO.Enabled:=OkOk ; SH_NUMDEPARTS.Enabled:=OkOk ; TSH_NUMDEPARTS.Enabled:=OkOk ;
  END ;
If Init Then
  BEGIN
  taSH_SOUCHEEXO.AsString:='-' ; taSH_NUMDEPARTS.AsInteger:=1 ; taSH_NUMDEPARTP.AsInteger:=1 ;
  END ;
END ;

Procedure TFSouche.AutoriseCheckBox ;
BEGIN
if (Lequel='BUD') or (Lequel='GES') then
  BEGIN
  SH_TYPE.Checked:=False ; SH_TYPE.Enabled:=False ;
  SH_ANALYTIQUE.Checked:=False ; SH_ANALYTIQUE.Enabled:=False ;
  SH_SIMULATION.Checked:=False ; SH_SIMULATION.Enabled:=False ;
  InitMultiSouche(TRUE,FALSE,FALSE) ;
  Exit ;
  END ;
if SH_TYPE.Checked then
   BEGIN
   SH_ANALYTIQUE.Checked:=False ; SH_ANALYTIQUE.Enabled:=False ;
   SH_SIMULATION.Checked:=False ; SH_SIMULATION.Enabled:=False ;
   InitMultiSouche(TRUE,FALSE,FALSE) ;
   END else
   if SH_ANALYTIQUE.Checked then
      BEGIN
      SH_TYPE.Checked:=False ; SH_TYPE.Enabled:=False ;
      SH_SIMULATION.Checked:=False ; SH_SIMULATION.Enabled:=False ;
      InitMultiSouche(TRUE,FALSE,FALSE) ;
      END else
      if SH_SIMULATION.Checked then
         BEGIN
         SH_ANALYTIQUE.Checked:=False ; SH_ANALYTIQUE.Enabled:=False ;
         SH_TYPE.Checked:=False ; SH_TYPE.Enabled:=False ;
         InitMultiSouche(TRUE,FALSE,FALSE) ;
         END Else
         BEGIN
         InitMultiSouche(TRUE,FALSE,TRUE) ;
         END ;
END ;

procedure TFSouche.SH_TYPEMouseUp(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
  inherited;
if (SH_Type.Checked) then
  if (PresenceComplexe('SOUCHE',['SH_TYPE','SH_SOUCHE'],['=','<>'],['REL',SH_SOUCHE.Text],['S','S'])) then BEGIN HM2.Execute(3,'',''); SH_Type.Checked:=False ; END ;
if SH_Type.Checked then
  BEGIN
  //TSoucheSH_TYPE.AsString:='REL' ;
  SH_MASQUENUM.Enabled:=True ; SH_MASQUENUM.Color:=clWindow ;
  SH_ANALYTIQUE.Checked:=False ; SH_ANALYTIQUE.Enabled:=False ;
  SH_SIMULATION.Checked:=False ; SH_SIMULATION.Enabled:=False ;
  TSH_NATUREPIECE.Enabled:=True ; SH_NATUREPIECE.Enabled:=True ;
  TSH_JOURNAL.Enabled:=True ; SH_JOURNAL.Enabled:=True ;
  END else
  BEGIN
  //TSoucheSH_TYPE.AsString:='CPT' ;
  SH_MASQUENUM.Text:='' ;
  SH_MASQUENUM.Enabled:=False ; SH_MASQUENUM.Color:=clBtnFace ;
  SH_ANALYTIQUE.Enabled:=True ; SH_SIMULATION.Enabled:=True ;
  TSH_NATUREPIECE.Enabled:=False ; SH_NATUREPIECE.Enabled:=False ;
  TSH_JOURNAL.Enabled:=False ; SH_JOURNAL.Enabled:=False ;
  END ;
OkAccesMultiSouche ;
end;

procedure TFSouche.SH_ANALYTIQUEMouseUp(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
  inherited;
if SH_ANALYTIQUE.Checked then
   BEGIN
//   SH_MASQUENUM.Text:='' ; SH_MASQUENUM.Enabled:=False ; SH_MASQUENUM.Color:=clBtnFace ;
   SH_Type.Checked:=False ; SH_Type.Enabled:=False ;
   SH_SIMULATION.Checked:=False ; SH_SIMULATION.Enabled:=False ;
   END else
   BEGIN
   SH_Type.Enabled:=True ; SH_SIMULATION.Enabled:=True ;
//   SH_MASQUENUM.Enabled:=True ; SH_MASQUENUM.Color:=clWindow ;
   END ;
OkAccesMultiSouche ;
end;

procedure TFSouche.SH_SIMULATIONMouseUp(Sender: TObject; Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
  inherited;
if SH_SIMULATION.Checked then
   BEGIN
   SH_TYPE.Checked:=False ; SH_TYPE.Enabled:=False ;
   SH_ANALYTIQUE.Checked:=False ; SH_ANALYTIQUE.Enabled:=False ;
   END else
   BEGIN
   SH_TYPE.Enabled:=True ; SH_ANALYTIQUE.Enabled:=True ;
   END ;
OkAccesMultiSouche ;
end;

procedure TFSouche.OkAccesMultiSouche ;
BEGIN
if (Lequel='BUD') or (Lequel='GES') then BEGIN InitMultiSouche(TRUE,FALSE,FALSE) ; Exit ; END ;
if SH_TYPE.Checked then InitMultiSouche(TRUE,FALSE,FALSE) else
  if SH_ANALYTIQUE.Checked then InitMultiSouche(TRUE,FALSE,FALSE) else
      if SH_SIMULATION.Checked then InitMultiSouche(TRUE,FALSE,FALSE) Else
         InitMultiSouche(TRUE,FALSE,TRUE) ;
END ;

procedure TFSouche.STaDataChange(Sender: TObject; Field: TField);
var Ok : boolean ;
begin
  inherited;
if Field=nil then
  BEGIN
  Ok:=BInsert.enabled ;
  if (Lequel<>'CPT') and (Lequel<>'GES') and (not ((Ta.Bof) and (Ta.Eof))) then BInsert.Enabled:=False else BInsert.enabled:=Ok ;
  END ;
end;

Procedure TFSouche.ClearChampSurEnregOk ;
BEGIN
if (TaSH_TYPE.AsString<>'REL') then
   BEGIN TaSH_NATUREPIECE.AsString:='' ;  TaSH_JOURNAL.AsString:='' ; END ;
END ;

Procedure TFSouche.AutoriseNatPieJal ;
BEGIN
if (Lequel='GES') then
  BEGIN
//  TSH_NATUREPIECE.Enabled:=True ; SH_NATUREPIECE.Enabled:=True ;
  TSH_JOURNAL.Enabled:=False ; SH_JOURNAL.Enabled:=False ;
  Exit ;
  END ;
if SH_Type.Checked then
  BEGIN
  TSH_NATUREPIECE.Enabled:=True ; SH_NATUREPIECE.Enabled:=True ;
  TSH_JOURNAL.Enabled:=True ; SH_JOURNAL.Enabled:=True ;
  END else
  BEGIN
  TSH_NATUREPIECE.Enabled:=False ; SH_NATUREPIECE.Enabled:=False ;
  TSH_JOURNAL.Enabled:=False ; SH_JOURNAL.Enabled:=False ;
  END ;
END ;

procedure TFSouche.FormClose(Sender: TObject; var Action: TCloseAction);
begin
if Parent is THPanel then _Bloqueur('nrCloture',False) ;
  inherited;
end;

procedure TFSouche.SH_SOUCHEEXOClick(Sender: TObject);
begin
  inherited;
SH_NUMDEPARTS.Enabled:=SH_SOUCHEEXO.Checked ;
TSH_NUMDEPARTS.Enabled:=SH_SOUCHEEXO.Checked ;
If SH_SOUCHEEXO.Checked And (Ta.state in [dsEdit]) Then SH_NUMDEPARTS.Value:=GetNextNumSouche(SH_SOUCHE.Text,VH^.Suivant) ;
end;

Function TFSouche.GetNextNumSouche(Souche : String ; Exo : TExoDate) : Integer ;
Var Sql,Jal,StWhereJal : String ;
    Q1 : TQuery ;
BEGIN
Result:=1 ;
If Not VH^.MultiSouche Then Exit ;
SQL:='Select J_JOURNAL from JOURNAL Where J_COMPTEURNORMAL="'+Souche+'" ' ;
Q1:=OpenSql(Sql,True,-1,'',true) ;
StWhereJal:='' ;
While Not Q1.Eof Do
  BEGIN
  Jal:=Q1.FindField('J_JOURNAL').AsString ;
  If StWhereJal='' Then StWhereJal:='E_JOURNAL="'+Jal+'" '
                   Else StWhereJal:=StWhereJal+' OR E_JOURNAL="'+Jal+'" ' ;
  Q1.Next ;
  END ;
Ferme(Q1) ;
If StWhereJal<>'' Then
   BEGIN
   StWhereJal:='('+StWhereJal+') ' ;
   Result:=RecalculSoucheSurUnExo(StWhereJal,Souche,Exo,FALSE) ;
   END ;
END ;

Function RecalculSoucheSurUnExo(StWhereJal,Souche : String ; Exo : TExoDate ; Ecrit : Boolean) : Integer ;
Var Num : Integer ;
    SQL : String ;
    Q1 : TQuery ;
BEGIN
Num:=0 ;
SQL:='Select Max(E_NUMEROPIECE) From ECRITURE Where '+StWhereJal+' AND E_QUALIFPIECE="N"' ;
If Exo.Code<>'' Then SQL:=SQL+' AND E_EXERCICE="'+Exo.Code+'" ' ;
Q1:=OpenSql(Sql,True,-1,'',true) ;
FetchSQLODBC(Q1) ;
If Not Q1.Eof Then Num:=Q1.Fields[0].AsInteger ;
Ferme(Q1) ;
Inc(Num) ;
If Ecrit Then
  BEGIN
  SQL:='' ;
  If Exo.Code='' Then SQL:='Update SOUCHE Set SH_NUMDEPART='+IntToStr(Num)+' WHERE SH_SOUCHE="'+Souche+'"' Else
   If Exo.Code=VH^.Suivant.Code Then SQL:='Update SOUCHE Set SH_NUMDEPARTS='+IntToStr(Num)+' WHERE SH_SOUCHE="'+Souche+'"' Else
    If Exo.Code=VH^.EnCours.Code Then SQL:='Update SOUCHE Set SH_NUMDEPART='+IntToStr(Num)+' WHERE SH_SOUCHE="'+Souche+'"' Else
     If Exo.Code=VH^.Precedent.Code Then SQL:='Update SOUCHE Set SH_NUMDEPARTP='+IntToStr(Num)+' WHERE SH_SOUCHE="'+Souche+'"' ;
  If SQL<>'' Then ExecuteSQL(SQL) ;
  END ;
Result:=Num ;
END ;

end.
