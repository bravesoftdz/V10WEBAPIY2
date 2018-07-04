unit AssistSS ;

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
    assist,
    ComCtrls,
    hmsgbox,
    StdCtrls,
    Hctrls,
    ExtCtrls,
    Mask,
    DBCtrls,
    DB,
    DBTables,
    Spin,
    HDB,
    Grids,
    DBGrids,
    HEnt1,
    SocSISCO,
    Ent1,
    FichComm,
    Devise_tom,
    ExoSISCO,
    Tablette,
    hCompte,
    Menus,
    CPSTRUCTURE_TOF,
    CPAxe_TOM,
    HStatus,
    CPSECTION_TOM,
    CPGENERAUX_TOM,
    CPJOURNAL_TOM,
    CPTIERS_TOM,
    Tva,
{$IFDEF V530}
    EdtDoc,
{$ELSE}
    EdtRDoc,
{$ENDIF}
    CodePost,
    SOUCHE_TOM,
    HSysMenu,
    HTB97,
    Hqry,
    FouCpt,
    ParamSoc,
    UtilSoc, HPanel, ADODB, udbxDataset
    ;

procedure LanceAssistRSERVANT ;

type
  TFAssistRSERVANT = class(TFAssist)
    Welcome: TTabSheet;
    HLabel1: THLabel;
    SSociete: TDataSource;
    TSO_LIBELLE: THLabel;
    SO_LIBELLE: TEdit;
    HLabel2: THLabel;
    SO_CODE: TEdit;
    HLabel3: THLabel;
    SPlanRef: TDataSource;
    QPlanRef: TQuery;
    ParamCPT: TTabSheet;
    Users: TTabSheet;
    HLabel5: THLabel;
    bUsers: TToolbarButton97;
    bUserGroups: TToolbarButton97;
    EtabExo: TTabSheet;
    HLabel7: THLabel;
    bEtab: TToolbarButton97;
    bExos: TToolbarButton97;
    bDevises: TToolbarButton97;
    bSoc1: TToolbarButton97;
    bSoc3: TToolbarButton97;
    bSoc4: TToolbarButton97;
    bSoc5: TToolbarButton97;
    HLabel8: THLabel;
    GroupBox2: TGroupBox;
    GroupBox4: TGroupBox;
    GroupBox5: TGroupBox;
    GroupBox6: TGroupBox;
    StepMenu: TPopupMenu;
    ip1: TToolbarButton97;
    ip2: TToolbarButton97;
    FinEtape: TMenuItem;
    N1: TMenuItem;
    Reglements: TTabSheet;
    HLabel9: THLabel;
    GroupBox7: TGroupBox;
    bModePaiement: TToolbarButton97;
    bModeRegl: TToolbarButton97;
    Analytique: TTabSheet;
    HLabel10: THLabel;
    GroupBox8: TGroupBox;
    bAna1: TToolbarButton97;
    bAna2: TToolbarButton97;
    PopAxes: TPopupMenu;
    Axe1: TMenuItem;
    Axe2: TMenuItem;
    Axe3: TMenuItem;
    Axe4: TMenuItem;
    Axe5: TMenuItem;
    bStep: TToolbarButton97;
    PlanCpt: TTabSheet;
    HLabel11: THLabel;
    GroupBox9: TGroupBox;
    bCpt2: TToolbarButton97;
    TVA: TTabSheet;
    HLabel12: THLabel;
    GroupBox10: TGroupBox;
    bTaxes1: TToolbarButton97;
    bTaxes2: TToolbarButton97;
    bTaxes3: TToolbarButton97;
    Tiers: TTabSheet;
    HLabel13: THLabel;
    GroupBox11: TGroupBox;
    bTiers9: TToolbarButton97;
    Divers: TTabSheet;
    HLabel14: THLabel;
    GroupBox12: TGroupBox;
    bDivers1: TToolbarButton97;
    bDivers2: TToolbarButton97;
    bSoc7: TToolbarButton97;
    Label1: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    Label4: TLabel;
    Label5: TLabel;
    Label6: TLabel;
    Label7: TLabel;
    Label8: TLabel;
    Label9: TLabel;
    Label10: TLabel;
    Label11: TLabel;
    Label12: TLabel;
    Label13: TLabel;
    Label14: TLabel;
    Label15: TLabel;
    Label16: TLabel;
    Label17: TLabel;
    Label18: TLabel;
    Label21: TLabel;
    Label22: TLabel;
    Label23: TLabel;
    Label24: TLabel;
    Label25: TLabel;
    Label26: TLabel;
    Label27: TLabel;
    Label28: TLabel;
    Label29: TLabel;
    Label30: TLabel;
    Label31: TLabel;
    Label32: TLabel;
    Label33: TLabel;
    Label34: TLabel;
    Label35: TLabel;
    Label36: TLabel;
    Label37: TLabel;
    Label38: TLabel;
    Label19: TLabel;
    Label20: TLabel;
    Label39: TLabel;
    Label40: TLabel;
    Sauve: TSaveDialog;
    HM: THMsgBox;
    procedure bSoc1Click(Sender: TObject);
    procedure bSoc3Click(Sender: TObject);
    procedure bSoc4Click(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure bUsersClick(Sender: TObject);
    procedure bUserGroupsClick(Sender: TObject);
    procedure bEtabClick(Sender: TObject);
    procedure bExosClick(Sender: TObject);
    procedure bDevisesClick(Sender: TObject);
    procedure StepMenuClick(Sender: TObject);
    procedure PChange(Sender: TObject); Override ;
    procedure FinEtapeClick(Sender: TObject);
    procedure bModePaiementClick(Sender: TObject);
    procedure bModeReglClick(Sender: TObject);
    procedure bAna2Click(Sender: TObject);
    procedure bCpt2Click(Sender: TObject);
    procedure bTaxes1Click(Sender: TObject);
    procedure bTaxes2Click(Sender: TObject);
    procedure bTaxes3Click(Sender: TObject);
    procedure bTiers9Click(Sender: TObject);
    procedure bDivers1Click(Sender: TObject);
    procedure bDivers2Click(Sender: TObject);
    procedure bFinClick(Sender: TObject);
    function  FirstPage : TTabSheet ; Override ;
    procedure bAna1Click(Sender: TObject);
    procedure bSoc5Click(Sender: TObject);
    procedure bSoc7Click(Sender: TObject);
  private
    { Déclarations privées }
    Steps  : Array[1..20] of Boolean ;
    procedure ChangeStepBitmap ;
    procedure ParamLaSoc(p : integer) ;
    procedure SauverLaSoc ( AvecEtapes : boolean ) ;
    procedure RestoreButton(b : TToolbarButton97) ;
  public
    { Déclarations publiques }
  end;


implementation

uses FmtChoix, JalSISCO, InfFicSI, SISCOCOR, ImpFicU, TImpFic, UTOMUTILISAT,
  USERGRP_TOM ;

{$R *.DFM}

procedure LanceAssistRSERVANT ;
Var X : TFAssistRSERVANT ;
BEGIN
// GP le 02/09/98 if Not BlocageMonoPoste(True) then Exit ;
SourisSablier ;
X:=TFAssistRSERVANT.Create(Application) ;
 Try
  X.ShowModal ;
 Finally
  X.Free ;
// GP le 02/09/98   DeblocageMonoPoste(True) ;
 End ;
SourisNormale ;
END ;

Procedure RecupSoc(StFichier : String) ;
Var Fichier : TextFile ;
    St,St1 : String ;
    Ok00,Ok01,Ok02,Ok03,Ok50 : Boolean ;
    i : Integer ;
{$IFDEF SPEC302}
    Q : TQuery ;
{$ENDIF}
    Q1 : TQuery ;
begin
{$IFDEF SPEC302}
Q:=OpenSQL('Select * FROM SOCIETE',FALSE) ;
If Q.Eof Then BEGIN Ferme(Q) ; Exit ; END ;
{$ENDIF}
AssignFile(Fichier,StFichier) ; {$i-} Reset(Fichier) ; {$i+}
InitMove(1000,'') ;
Ok00:=FALSE ; Ok01:=FALSE ; Ok02:=FALSE ; Ok03:=FALSE ; Ok50:=FALSE ;
//SetParamSoc('SO_SOCIETE',SO_CODE.Text) ; SetParamSoc('SO_LIBELLE',SO_LIBELLE.Text) ;
{$IFDEF SPEC302}
Q.Edit ;
{$ENDIF}
While Not Eof(Fichier) Do
  BEGIN
  MoveCur(FALSE) ;
  ReadLn(Fichier,St) ;
  If (Copy(St,1,2)='00') Then
    BEGIN
    Ok00:=TRUE ;
{$IFDEF SPEC302}
    Q.FindField('SO_LGCPTEGEN').AsInteger:=StrToInt(Copy(St,21,2)) ;
    Q.FindField('SO_LGCPTEAUX').AsInteger:=StrToInt(Copy(St,21,2)) ;
{$ELSE}
    i:=StrToInt(Copy(St,21,2)) ; SetParamSoc('SO_LGCPTEGEN',i) ;
    i:=StrToInt(Copy(St,21,2)) ; SetParamSoc('SO_LGCPTEAUX',i) ;
{$ENDIF}
    END ;
  If (Copy(St,1,2)='01') Then
    BEGIN
    Ok01:=TRUE ;
{$IFDEF SPEC302}
    Q.FindField('SO_LIBELLE').ASString:=Copy(St,3,30) ;
{$ELSE}
    St1:=Copy(St,3,30) ; SetParamSoc('SO_LIBELLE',St1) ;
{$ENDIF}
    END ;
  If (Copy(St,1,2)='02') Then
    BEGIN
    Ok02:=TRUE ;
{$IFDEF SPEC302}
    Q.FindField('SO_ADRESSE1').ASString:=Trim(Copy(St,3,32)) ;
    Q.FindField('SO_ADRESSE2').ASString:=Trim(Copy(St,35,32)) ;
    Q.FindField('SO_ADRESSE3').ASString:=Trim(Copy(St,67,32)) ;
{$ELSE}
    St1:=Trim(Copy(St,3,32))  ; SetParamSoc('SO_ADRESSE1',St1) ;
    St1:=Trim(Copy(St,35,32)) ; SetParamSoc('SO_ADRESSE2',St1) ;
    St1:=Trim(Copy(St,67,32)) ; SetParamSoc('SO_ADRESSE3',St1) ;
{$ENDIF}
    END ;
  If (Copy(St,1,2)='03') Then
    BEGIN
    Ok03:=TRUE ;
{$IFDEF SPEC302}
    Q.FindField('SO_SIRET').ASString:=Trim(Copy(St,3,14)) ;
    Q.FindField('SO_APE').ASString:=Trim(Copy(St,17,4)) ;
    Q.FindField('SO_TELEPHONE').ASString:=Trim(Copy(St,21,12)) ;
{$ELSE}
    St1:=Trim(Copy(St,3,14))  ; SetParamSoc('SO_SIRET',St1) ;
    St1:=Trim(Copy(St,17,4))  ; SetParamSoc('SO_APE',St1) ;
    St1:=Trim(Copy(St,21,12)) ; SetParamSoc('SO_TELEPHONE',St1) ;
{$ENDIF}
    If Copy(St,54,1)='N' Then Ok50:=TRUE ;
    END ;
  If (Copy(St,1,2)='50') Then
    BEGIN
    Ok50:=TRUE ;
    Q1:=OpenSQL('SELECT * FROM AXE WHERE X_AXE="A1"',FALSE) ;
    If Not Q1.Eof Then
      BEGIN
      Q1.Edit ;
      Q1.FindField('X_LONGSECTION').AsInteger:=StrToInt(Copy(St,44,2)) ;
      Q1.Post ;
      END ;
    Ferme(Q1) ;
    END ;
  If Ok00 And Ok01 And Ok02 And Ok03 And Ok50 Then Break ;
  END ;
{$IFDEF SPEC302}
Q.Post ;
Ferme(Q) ;
{$ENDIF}
System.Close(Fichier) ;
FiniMove ;
end;

Procedure InsereFourSISCO(Q1 : TQuery ; Var i : Integer ; Cpt,Code : String) ;
BEGIN
Inc(i) ;
If i<=5 Then
  BEGIN
  Q1.Insert ; InitNew(Q1) ;
  Q1.FindField('CC_TYPE').AsString:='SIS' ;
  Q1.FindField('CC_CODE').AsString:=Code+IntToStr(i) ;
  Q1.FindField('CC_LIBELLE').AsString:=Cpt ;
  Q1.FindField('CC_ABREGE').AsString:=Cpt ;
  Q1.Post ;
  END ;
END ;

Procedure FourchetteDefautSISCO ;
Var Q,Q1 : TQuery ;
    Cpt : String ;
    icli,ifou : Integer ;
BEGIN
Q:=OpenSQL('Select * FROM SOCIETE',FALSE) ;
If Not Q.Eof Then
  BEGIN
  Q.Edit ;
  Cpt:='60000000000000000' ; Cpt:=BourreEtLess(Cpt,fbGene) ; Q.FindField('SO_CHADEB1').AsString:=Cpt ;
  Cpt:='69999999999999999' ; Cpt:=BourreEtLess(Cpt,fbGene) ; Q.FindField('SO_CHAFIN1').AsString:=Cpt ;
  Cpt:='70000000000000000' ; Cpt:=BourreEtLess(Cpt,fbGene) ; Q.FindField('SO_PRODEB1').AsString:=Cpt ;
  Cpt:='79999999999999999' ; Cpt:=BourreEtLess(Cpt,fbGene) ; Q.FindField('SO_PROFIN1').AsString:=Cpt ;
  Q.Post ;
  END ;
Ferme(Q) ;
ExecuteSQL('DELETE FROM CHOIXCOD WHERE CC_TYPE="SIS" AND CC_CODE<>"NOM"') ;
Q:=OpenSQL('SELECT * FROM CORRESP WHERE CR_TYPE="SIS"',TRUE) ;
Q1:=OpenSQL('Select * from CHOIXCOD WHERE CC_TYPE="SIS"',FALSE) ;  Icli:=0 ; IFou:=0 ;
While Not Q.Eof Do
  BEGIN
  Cpt:=Q.FindField('CR_CORRESP').AsString ; Cpt:=BourreEtLess(Cpt,fbGene) ;
  If Copy(Cpt,1,2)='40' Then InsereFourSISCO(Q1,iFou,Cpt,'FO') Else
    If Copy(Cpt,1,2)='41' Then InsereFourSISCO(Q1,iCli,Cpt,'CL') ;
  Q.Next ;
  END ;
Ferme(Q) ;
Q1.Insert ; InitNew(Q1) ;
Q1.FindField('CC_TYPE').AsString:='SIS' ;
Q1.FindField('CC_CODE').AsString:='BA1' ;
Cpt:='51200000000000000' ; Cpt:=BourreEtLess(Cpt,fbGene) ;
Q1.FindField('CC_LIBELLE').AsString:=Cpt ;
Cpt:='51499999999999999' ; Cpt:=BourreEtLess(Cpt,fbGene) ;
Q1.FindField('CC_ABREGE').AsString:=Cpt ;
Q1.Post ;
Q1.Insert ; InitNew(Q1) ;
Q1.FindField('CC_TYPE').AsString:='SIS' ;
Q1.FindField('CC_CODE').AsString:='CA1' ;
Cpt:='53000000000000000' ; Cpt:=BourreEtLess(Cpt,fbGene) ;
Q1.FindField('CC_LIBELLE').AsString:=Cpt ;
Cpt:='53999999999999999' ; Cpt:=BourreEtLess(Cpt,fbGene) ;
Q1.FindField('CC_ABREGE').AsString:=Cpt ;
Q1.Post ;
Ferme(Q1) ;
END ;

Procedure AlimParDefautSISCO(StFichier : String) ;
Var St : String ;
BEGIN
If ExisteSQL('SELECT * FROM GENERAUX') Or ExisteSQL('SELECT * FROM TIERS') Or ExisteSQL('SELECT * FROM SECTION') Or ExisteSQL('SELECT * FROM JOURNAL') Then
  BEGIN
  St:='0;Opération impossible;Des Comptes existent dans la base;E;O;O;O;' ; HShowMessage(St,'','') ; Exit ;
  END ;
If ExisteSQL('SELECT * FROM ECRITURE') Then
  BEGIN
  St:='0;Opération impossible;Des mouvements existent dans la base;E;O;O;O;' ; HShowMessage(St,'','') ; Exit ;
  END ;
If HShowMessage('0;Initialisation par défaut;Confirmez-vous le traitement?;Q;YNC;N;C;','','')<>mrYes Then Exit ;
If Not EstFichierSISCO(StFichier,TRUE) Then Exit ;
RecupSOC(StFichier) ;
ChargeSocieteHALLEY ;
MajSISCOCOR(StFichier) ;
FourchetteDefautSISCO ;
ChargeSocieteHALLEY ;
END ;

procedure TFAssistRSERVANT.RestoreButton(b : TToolbarButton97) ;
BEGIN
//b.Restore ;
END ;

function TFAssistRSERVANT.FirstPage : TTabSheet  ;
var i : Integer ;
BEGIN
//Result := Inherited FirstPage ;
result := Welcome ; // positionnement sur la première étape non terminée
i := 1 ; while Steps[i] do Inc(i) ;
if i>1 then result := TTabSheet(P.Pages[i-1]) ;
END ;

procedure TFAssistRSERVANT.ParamLaSoc(p : integer) ;
Var b : Boolean ;
    sSeule,sVirer : String ; 
BEGIN
SourisSablier ;
SauverLaSoc(False) ;
b := (p>=5) ;
{$IFDEF SPEC302}
FicheSocieteSISCO(taModif,p,b,'') ;
{$ELSE}
sSeule:='' ; sVirer:='' ;
Case p of
   0 : sseule:='SCO_COORDONNEES' ;
   1 : sseule:='SCO_COMPTABLES' ;
   2 : sseule:='SCO_FOURCHETTES' ;
   3 : sseule:='SCO_COMPTESSPECIAUX' ;
   4 : sseule:='SCO_DATESDIVERS;SCO_DIVERS' ;
   5 : sseule:='SCO_REGLEMENTSTVA' ;
   6 : sseule:='SCO_EURO' ;
   END ;
if sSeule='' then sSeule:='SCO_PARAMETRESGENERAUX' ;
ParamSociete(False,sVirer,sSeule,'',ChargeSocieteHalley,ChargePageSoc,SauvePageSocSansVerif,InterfaceSoc,1105000) ;
{$ENDIF}
SourisNormale ;
END ;

procedure TFAssistRSERVANT.bSoc1Click(Sender: TObject);
begin
ParamLaSoc(0) ;
RestoreButton(bSoc1) ;
end;

procedure TFAssistRSERVANT.bSoc5Click(Sender: TObject);
begin
ParamLaSoc(1) ;
RestoreButton(bSoc5) ;
end;

procedure TFAssistRSERVANT.bSoc3Click(Sender: TObject);
begin
ParamFourchetteImportCpt ; // fourchettes
RestoreButton(bSoc3) ;
end;

procedure TFAssistRSERVANT.bSoc4Click(Sender: TObject);
begin
ParamLaSoc(3) ; // comptes spéciaux
RestoreButton(bSoc4) ;
end;


procedure TFAssistRSERVANT.FormShow(Sender: TObject);
var i  : Integer ;
    mi : TMenuItem ;
    s,sCode,sLibelle,sRecup  : String ;
    Q : TQuery ;
begin
for i:=low(Steps) to high(Steps) do Steps[i] := False ;
sCode:='' ; sLibelle:='' ; sRecup:='' ;
{$IFDEF SPEC302}
Q:=OpenSQL('Select SO_SOCIETE, SO_LIBELLE, SO_RECUPCPTA from SOCIETE',True) ;
if Not Q.EOF then
   BEGIN
   sCode:=Q.Fields[0].AsString ; sLibelle:=Q.Fields[1].AsString ;
   srecup:=Q.Fields[2].AsString ;
   END ;
Ferme(Q) ;
{$ELSE}
Q:=OpenSQL('Select SO_SOCIETE from SOCIETE',True) ;
if Not Q.EOF then sCode:=Q.Fields[0].AsString ;
Ferme(Q) ;
sLibelle:=GetParamSocSecur('SO_LIBELLE','') ; sRecup:=GetParamSocSecur('SO_RECUPCPTA','') ;
{$ENDIF}
SO_CODE.Text:=sCode ; SO_LIBELLE.Text:=sLibelle ;
s:=sRecup ;
for i:=1 to Length(s) do BEGIN if Copy(s,i,1)='1' then Steps[i]:=True else Steps[i]:=False ; END ;
for i:=0 to P.ControlCount-1 do
   BEGIN
   If TTabSheet(P.Controls[i]).TabVisible Then
     BEGIN
     mi := TMenuItem.Create(Self);
     mi.Caption := TTabSheet(P.Controls[i]).Hint ;
     mi.OnClick := StepMenuClick ;
     StepMenu.Items.Add(mi);
     END ;
   END ;
inherited ;
If RecupSISCO Then Caption:=HM.Mess[0] Else
  If RecupSERVANT Then
    BEGIN
    Caption:=HM.Mess[1] ;  P.ActivePage:=Welcome ;
    END ;
Caption:=NomHalley+' - '+Caption ; UpdateCaption(Self) ;
SourisNormale ;
end;

procedure TFAssistRSERVANT.bUsersClick(Sender: TObject);
begin
FicheUser('') ;
RestoreButton(bUsers) ;
end;

procedure TFAssistRSERVANT.bUserGroupsClick(Sender: TObject);
begin
FicheUserGrp ;
RestoreButton(bUserGroups) ;
end;

procedure TFAssistRSERVANT.bEtabClick(Sender: TObject);
begin
FicheEtablissement_AGL(taCreat) ;
RestoreButton(bEtab) ;
end;

procedure TFAssistRSERVANT.bExosClick(Sender: TObject);
begin
FicheExerciceSISCO('',taCreat,1) ;
RestoreButton(bExos) ;
end;

procedure TFAssistRSERVANT.bDevisesClick(Sender: TObject);
begin
FicheDevise('',taModif,True) ;
RestoreButton(bDevises) ;
end;

procedure TFAssistRSERVANT.StepMenuClick(Sender: TObject);
var i : Integer ;
begin
RestorePage ;
i := TMenuItem(Sender).MenuIndex-2 ;
If P.Pages[i].TabVisible Then P.ActivePage := P.Pages[i] ;
PChange(nil) ;
end;

procedure TFAssistRSERVANT.PChange(Sender: TObject);
begin
inherited;
FinEtape.Checked := Steps[P.ActivePage.PageIndex+1] ;
ChangeStepBitmap ;
end ;

procedure TFAssistRSERVANT.ChangeStepBitmap ;
begin
if Steps[P.ActivePage.PageIndex + 1] then bStep.Glyph := ip2.Glyph
                                     else bStep.Glyph := ip1.Glyph ;
end ;

procedure TFAssistRSERVANT.FinEtapeClick(Sender: TObject);
begin
FinEtape.Checked := not FinEtape.Checked ;
Steps[P.ActivePage.PageIndex + 1] := FinEtape.Checked ;
ChangeStepBitmap ;
end;

procedure TFAssistRSERVANT.bModePaiementClick(Sender: TObject);
begin
FicheModePaie_AGL('');
RestoreButton(bModePaiement) ;
end;

procedure TFAssistRSERVANT.bModeReglClick(Sender: TObject);
begin
FicheRegle_AGL('',FALSE,taModif) ;
RestoreButton(bModeRegl) ;
end;

procedure TFAssistRSERVANT.bAna2Click(Sender: TObject);
begin
ParamPlanAnal('') ;
RestoreButton(bAna2) ;
end;

procedure TFAssistRSERVANT.bCpt2Click(Sender: TObject);
begin
FicheGene(Nil,'','',taCreatEnSerie,0) ;
RestoreButton(bCpt2) ;
end;

procedure TFAssistRSERVANT.bTaxes1Click(Sender: TObject);
begin
ParamTable('TTREGIMETVA',taCreat,1165000,Nil) ;
RestoreButton(bTaxes1) ;
end;

procedure TFAssistRSERVANT.bTaxes2Click(Sender: TObject);
begin
ParamTvaTpf(True) ;
RestoreButton(bTaxes2) ;
end;

procedure TFAssistRSERVANT.bTaxes3Click(Sender: TObject);
begin
ParamTvaTpf(False) ;
RestoreButton(bTaxes3) ;
end;

procedure TFAssistRSERVANT.bTiers9Click(Sender: TObject);
begin
FicheTiers(Nil,'','',taCreatEnSerie,1) ;
RestoreButton(bTiers9) ;
end;

procedure TFAssistRSERVANT.bDivers1Click(Sender: TObject);
begin
FicheJournal(Nil,'','',taCreatEnSerie,0) ;
RestoreButton(bDivers1) ;
end;

procedure TFAssistRSERVANT.bDivers2Click(Sender: TObject);
begin
YYLanceFiche_Souche('CPT','','ACTION=MODIFICATION;CPT'); //Compteurs
RestoreButton(bDivers2) ;
end;

procedure TFAssistRSERVANT.SauverLaSoc ( AvecEtapes : boolean ) ;
var s : String ;
    i : Integer ;
BEGIN
s:='' ;
if AvecEtapes then
   BEGIN
   for i:=1 to high(Steps) do BEGIN if Steps[i] then s := s + '1' else s := s + '0' ; END ;
   END ;
{$IFDEF SPEC302}
if AvecEtapes then ExecuteSQL('UPDATE SOCIETE SET SO_SOCIETE="'+SO_CODE.Text+'", SO_LIBELLE="'+SO_LIBELLE.Text+'", SO_RECUPCPTA="'+s+'"')
              else ExecuteSQL('UPDATE SOCIETE SET SO_SOCIETE="'+SO_CODE.Text+'", SO_LIBELLE="'+SO_LIBELLE.Text+'"') ;
{$ELSE}
ExecuteSQL('UPDATE SOCIETE SET SO_SOCIETE="'+SO_CODE.Text+'"') ;
SetParamSoc('SO_SOCIETE',SO_CODE.Text) ; SetParamSoc('SO_LIBELLE',SO_LIBELLE.Text) ;
if AvecEtapes then SetParamSoc('SO_RECUPCPTA',s) ;
{$ENDIF}
END ;

procedure TFAssistRSERVANT.bFinClick(Sender: TObject);
var s : String ;
    i : Integer ;
begin
inherited;
SauverLaSoc(True) ;
// Sauvegarde de la chaîne des étapes
Close ;
end;

procedure TFAssistRSERVANT.bAna1Click(Sender: TObject);
begin
FicheAxeAna ('') ;
RestoreButton(bAna1) ;
end;

procedure TFAssistRSERVANT.bSoc7Click(Sender: TObject);
begin
{$IFDEF SPEC302}
ParamLaSoc(4) ;
{$ELSE}
ParamLaSoc(1) ;
{$ENDIF}
RestoreButton(bSoc7) ;
end;

end.
