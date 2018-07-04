unit RegulInv;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, Hcompte, hmsgbox, Grids, DBGrids, Buttons, Hctrls, Mask, HEnt1,
  ExtCtrls, ComCtrls, LettUtil, DB, DBTables, Hqry, SaisUtil, SaisComm,
  Ent1, PrintDBG, Journal, DelVisuE, Letbatch, HStatus, HDB, Filtre, Menus,
  HSysMenu, HTB97, ed_tools, HPanel, UiUtil, UtilPGI, ADODB, udbxDataset,
  TntDBGrids, TntStdCtrls ;

type
  TFRegulInv = class(TForm)
    Pages: TPageControl;
    Princ: TTabSheet;
    HPB: TToolWindow97;
    BImprimer: TToolbarButton97;
    BValider: TToolbarButton97;
    BFerme: TToolbarButton97;
    BAide: TToolbarButton97;
    BRecherche: TToolbarButton97;
    GR: THDBGrid;
    HMRegul: THMsgBox;
    Cache: THCpteEdit;
    QR: THQuery;
    SR: TDataSource;
    FindRegul: TFindDialog;
    TabSheet1: TTabSheet;
    BGenere: TToolbarButton97;
    PFiltres: TToolWindow97;
    FFiltres: THValComboBox;
    BChercher: TToolbarButton97;
    BAgrandir: TToolbarButton97;
    BReduire: TToolbarButton97;
    HLabel4: THLabel;
    HLabel1: THLabel;
    HLabel3: THLabel;
    E_DATECOMPTABLE: THCritMaskEdit;
    E_AUXILIAIRE: THCpteEdit;
    E_GENERAL: THCpteEdit;
    HLabel5: THLabel;
    HLabel2: THLabel;
    HLabel6: THLabel;
    E_DATECOMPTABLE_: THCritMaskEdit;
    E_AUXILIAIRE_: THCpteEdit;
    E_GENERAL_: THCpteEdit;
    HLabel8: THLabel;
    Label14: THLabel;
    T_NATUREAUXI: THValComboBox;
    E_DEVISE: THValComboBox;
    Bevel1: TBevel;
    H_JOURNALR: THLabel;
    H_JOURNALE: THLabel;
    HLabel9: THLabel;
    HLabel7: THLabel;
    REFREGUL: TEdit;
    DATEGENERATION: THCritMaskEdit;
    JOURNAL: THValComboBox;
    HLabel12: THLabel;
    HLabel13: THLabel;
    HLabel10: THLabel;
    LIBREGUL: TEdit;
    CPTCREDIT: TComboBox;
    CPTDEBIT: TComboBox;
    H_SD: THLabel;
    H_SC: THLabel;
    MAXCREDIT: THNumEdit;
    MAXDEBIT: THNumEdit;
    Bevel2: TBevel;
    HMTrad: THSystemMenu;
    PLibres: TTabSheet;
    Bevel3: TBevel;
    TT_TABLE0: THLabel;
    TT_TABLE1: THLabel;
    TT_TABLE2: THLabel;
    TT_TABLE3: THLabel;
    TT_TABLE4: THLabel;
    TT_TABLE5: THLabel;
    TT_TABLE6: THLabel;
    TT_TABLE7: THLabel;
    TT_TABLE8: THLabel;
    TT_TABLE9: THLabel;
    T_TABLE4: THCpteEdit;
    T_TABLE3: THCpteEdit;
    T_TABLE2: THCpteEdit;
    T_TABLE1: THCpteEdit;
    T_TABLE0: THCpteEdit;
    T_TABLE5: THCpteEdit;
    T_TABLE6: THCpteEdit;
    T_TABLE7: THCpteEdit;
    T_TABLE8: THCpteEdit;
    T_TABLE9: THCpteEdit;
    POPF: TPopupMenu;
    BCreerFiltre: TMenuItem;
    BSaveFiltre: TMenuItem;
    BDelFiltre: TMenuItem;
    BRenFiltre: TMenuItem;
    BNouvRech: TMenuItem;
    BFiltre: TToolbarButton97;
    H_JOURNALC: THLabel;
    ModeOppose: TCheckBox;
    ExpliEuro: TLabel;
    Dock: TDock97;
    Dock971: TDock97;
    procedure E_AUXILIAIREDblClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure E_DEVISEChange(Sender: TObject);
    procedure BChercherClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure BSaveFiltreClick(Sender: TObject);
    procedure FFiltresChange(Sender: TObject);
    procedure BRechercheClick(Sender: TObject);
    procedure FindRegulFind(Sender: TObject);
    procedure BFermeClick(Sender: TObject);
    procedure BValiderClick(Sender: TObject);
    procedure BImprimerClick(Sender: TObject);
    procedure JOURNALChange(Sender: TObject);
    procedure H_JOURNALRDblClick(Sender: TObject);
    procedure DATEGENERATIONExit(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure BGenereClick(Sender: TObject);
    procedure E_DATECOMPTABLEKeyPress(Sender: TObject; var Key: Char);
    procedure BDelFiltreClick(Sender: TObject);
    procedure BAgrandirClick(Sender: TObject);
    procedure BReduireClick(Sender: TObject);
    procedure BNouvRechClick(Sender: TObject);
    procedure BRenFiltreClick(Sender: TObject);
    procedure BCreerFiltreClick(Sender: TObject);
    procedure BAideClick(Sender: TObject);
    procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure POPFPopup(Sender: TObject);
  private
    DEV : RDEVISE ;
    J   : REGJAL ;
    DateGene,NowFutur : TDateTime ;
    FindFirst,GeneCharge : boolean ;
    TPIECES   : TList ;
    TitreF    : String ;
    WMinX,WMinY : Integer ;
    Qui : tProfilTraitement ;
    procedure WMGetMinMaxInfo(var MSG: Tmessage); message WM_GetMinMaxInfo;
    procedure AttribCollectif ;
    procedure RechercheEcritures ;
    procedure ChargeAVide ;
    procedure GenereRegul ;
    function  WhereCrit : String ;
    function  WhereDetail ( RL : RLETTR ) : String ;
    procedure RemplirLeM ( Var M : RMVT ; RL : RLETTR ) ;
    procedure ValideLaRegul ;
    procedure InitCriteres ;
    procedure CouvreRegul ( O : TOBM ; RL : RLETTR ) ;
    procedure TripoteTitres ;
  public
    Regul : PG_LETTRAGE ;
  end;

procedure RegulLettrageInverse ( Regul,Convert : Boolean ) ;
{$IFDEF CCMP}
procedure RegulLettrageMPInverse ( Regul,Convert : Boolean ; Qui : tProfilTraitement) ;
{$ENDIF}

implementation

uses Paramdat;

{$R *.DFM}

procedure RegulLettrageInverse ( Regul,Convert : Boolean ) ;
Var X  : TFLetRegulInv ;
    PP : THPanel ;
BEGIN
if PasCreerDate(V_PGI.DateEntree) then Exit ;
if _Blocage(['nrCloture','nrBatch','nrSaisieModif','nrLettrage','nrEnca','nrDeca'],True,'nrBatch') then Exit ;
X:=TFLetRegulInv.Create(Application) ;
if Convert then BEGIN X.Regul:=pglConvert ; X.TitreF:='REGULCONVERT' ; END else
   if Regul then BEGIN X.Regul:=pglRegul ; X.TitreF:='REGULLETTRAGE'; END else
      BEGIN X.Regul:=pglEcart ; X.TitreF:='ECARTCHANGE' ; END ;
PP:=FindInsidePanel ;
if PP=Nil then
   BEGIN
    Try
     X.Qui:=prAucun ;
     X.ShowModal ;
    Finally
     X.Free ;
     _Bloqueur('nrBatch',False) ;
    End ;
   Screen.Cursor:=SyncrDefault ;
   END else
   BEGIN
   InitInside(X,PP) ;
   X.Show ;
   END ;
END ;

{$IFDEF CCMP}
procedure RegulLettrageMPInverse ( Regul,Convert : Boolean ; Qui : tProfilTraitement) ;
Var X  : TFLetRegulInv ;
    PP : THPanel ;
BEGIN
if PasCreerDate(V_PGI.DateEntree) then Exit ;
If Not MonProfilOk(Qui) Then
  BEGIN
  if _Blocage(['nrCloture','nrBatch','nrSaisieModif','nrLettrage','nrEnca','nrDeca'],True,'nrBatch') then Exit ;
  END ;
X:=TFLetRegulInv.Create(Application) ;
if Convert then BEGIN X.Regul:=pglConvert ; X.TitreF:='REGULCONVERT' ; END else
   if Regul then BEGIN X.Regul:=pglRegul ; X.TitreF:='REGULLETTRAGE'; END else
      BEGIN X.Regul:=pglEcart ; X.TitreF:='ECARTCHANGE' ; END ;
PP:=FindInsidePanel ;
if PP=Nil then
   BEGIN
    Try
     X.Qui:=Qui ;
     X.ShowModal ;
    Finally
     X.Free ;
     If Not MonProfilOk(Qui) Then _Bloqueur('nrBatch',False) ;
    End ;
   Screen.Cursor:=SyncrDefault ;
   END else
   BEGIN
   InitInside(X,PP) ;
   X.Show ;
   END ;
END ;
{$ENDIF}

procedure TFRegulInv.BCreerFiltreClick(Sender: TObject);
begin NewFiltre(TitreF,FFiltres,Pages) ; end;

procedure TFRegulInv.BSaveFiltreClick(Sender: TObject);
begin SaveFiltre(TitreF,FFiltres,Pages) ; end;

procedure TFRegulInv.BDelFiltreClick(Sender: TObject);
begin DeleteFiltre(TitreF,FFiltres) ; end;

procedure TFRegulInv.BRenFiltreClick(Sender: TObject);
begin RenameFiltre(TitreF,FFiltres) ; end;

procedure TFRegulInv.FFiltresChange(Sender: TObject);
begin LoadFiltre(TitreF,FFiltres,Pages) ; end;

procedure TFRegulInv.BAgrandirClick(Sender: TObject);
begin ChangeListeCrit(Self,True) ; end;

procedure TFRegulInv.BReduireClick(Sender: TObject);
begin ChangeListeCrit(Self,False) ; end;

procedure TFRegulInv.E_DATECOMPTABLEKeyPress(Sender: TObject; var Key: Char);
begin ParamDate(Self,Sender,Key) ; end;

procedure TFRegulInv.H_JOURNALRDblClick(Sender: TObject);
Var A : TActionFiche ;
begin
if Not ExJaiLeDroitConcept(TConcept(ccJalModif),False) then A:=taConsult else A:=taModif ;
if JOURNAL.Value<>'' then
   BEGIN
   FicheJournal(Nil,'',JOURNAL.Value,A,0) ; if A=taModif then JournalChange(Nil) ;
   END ;
end;

procedure TFRegulInv.E_AUXILIAIREDblClick(Sender: TObject);
begin
AttribCollectif ;
end;

procedure TFRegulInv.BChercherClick(Sender: TObject);
begin
if ((VH^.ExoV8.Code<>'') and (StrToDate(E_DATECOMPTABLE.Text)<VH^.ExoV8.Deb)) then E_DATECOMPTABLE.Text:=DateToStr(VH^.ExoV8.Deb) ;
RechercheEcritures ;
end;

procedure TFRegulInv.BGenereClick(Sender: TObject);
Var Indice : integer ;
begin
if Regul=pglRegul then Indice:=3 else
   if Regul=pglEcart then Indice:=4 else Indice:=14 ;
if TPieces.Count>0 then VisuPiecesGenere(TPieces,EcrGen,Indice) ;
end;

procedure TFRegulInv.FormCreate(Sender: TObject);
begin
QR.Manuel:=True ; TPIECES:=TList.Create ;
PopUpMenu:=ADDMenuPop(PopUpMenu,'','') ;
WMinX:=Width ; WMinY:=252 ;
end;

procedure TFRegulInv.BRechercheClick(Sender: TObject);
begin FindFirst:=True ; FindRegul.Execute ; end;

procedure TFRegulInv.FindRegulFind(Sender: TObject);
begin Rechercher(GR,FindRegul,FindFirst) ; end;

procedure TFRegulInv.BFermeClick(Sender: TObject);
begin Close ; end;

procedure TFRegulInv.BImprimerClick(Sender: TObject);
begin PrintDBGrid(GR,Nil,Caption,'') ; end;

{================================ ENTETE =========================================}
procedure TFRegulInv.AttribCollectif ;
Var Q : TQuery ;
BEGIN
if ((E_AUXILIAIRE_.Text<>'') or (E_GENERAL.Text<>'') or
    (E_GENERAL_.Text<>'') or (E_AUXILIAIRE.Text='')) then Exit ;
Q:=OpenSQL('Select T_COLLECTIF from Tiers Where T_AUXILIAIRE="'+E_AUXILIAIRE.Text+'"',True) ;
if Not Q.EOF then E_GENERAL.Text:=Q.Fields[0].AsString ; Ferme(Q) ;
END ;

procedure TFRegulInv.E_DEVISEChange(Sender: TObject);
Var DEVF : RDEVISE ;
begin
if E_DEVISE.Value=DEV.Code then Exit ; DEV.Code:=E_DEVISE.Value ;
GetInfosDevise(DEV) ; DEV.Taux:=GetTaux(DEV.Code,DEV.DateTaux,V_PGI.DateEntree) ;
ChangeMask(MAXDEBIT,DEV.Decimale,DEV.Symbole) ;
ChangeMask(MAXCREDIT ,DEV.Decimale,DEV.Symbole) ;
if ((VH^.TenueEuro) and (DEV.Code=V_PGI.DevisePivot) and (ModeOppose.Checked)) then
   BEGIN
   DEVF.Code:=V_PGI.DeviseFongible ; GetInfosDevise(DEVF) ;
   END else
   BEGIN
   DEVF:=DEV ;
   END ;
MAXDEBIT.Value:=DEVF.MaxDebit ; MAXCREDIT.Value:=DEVF.MaxCredit ;
if Regul=pglEcart then
   BEGIN
   CptDebit.Items.Clear ; CptCredit.Items.Clear ;
   CptDebit.Items.Add(DEV.CptDebit) ; CptCredit.Items.Add(DEV.CptCredit) ;
   if CptDebit.Items.Count>0 then CptDebit.ItemIndex:=0 ;
   if CptCredit.Items.Count>0 then CptCredit.ItemIndex:=0 ;
   CPTDEBIT.Enabled:=False ; CPTCREDIT.Enabled:=False ;
   END ;
end;

procedure TFRegulInv.JOURNALChange(Sender: TObject);
Var i : integer ;
begin
if JOURNAL.Value=J.Journal then Exit ;
RemplirInfosRegul(J,JOURNAL.Value) ;
if ((Journal.Value<>'') and (J.Facturier='') and (Regul<>pglConvert)) then
   BEGIN
   if Not GeneCharge then HMRegul.Execute(16,'','') ;
   Journal.Value:='' ;
   Exit ;
   END ;
if Regul=pglRegul then
   BEGIN
   CptDebit.Items.Clear ; CptCredit.Items.Clear ;
   for i:=1 to 3 do if J.D[i]<>'' then CptDebit.Items.Add(J.D[i]) ;
   for i:=1 to 3 do if J.C[i]<>'' then CptCredit.Items.Add(J.C[i]) ;
   if CptDebit.Items.Count>0 then CptDebit.ItemIndex:=0 ;
   if CptCredit.Items.Count>0 then CptCredit.ItemIndex:=0 ;
   END else if Regul=pglConvert then
   BEGIN
   CptDebit.Items.Clear ; CptCredit.Items.Clear ;
   CptDebit.Items.Add(VH^.EccEuroDebit) ;
   CptCredit.Items.Add(VH^.EccEuroCredit) ;
   if CptDebit.Items.Count>0 then CptDebit.ItemIndex:=0 ;
   if CptCredit.Items.Count>0 then CptCredit.ItemIndex:=0 ;
   END ;
end;

procedure TFRegulInv.DATEGENERATIONExit(Sender: TObject);
Var DD : TDateTime ;
    Err : integer ;
BEGIN
if Not IsValidDate(DATEGENERATION.Text) then
   BEGIN
   HMRegul.Execute(5,caption,'') ;
   DATEGENERATION.Text:=DateToStr(V_PGI.DateEntree) ; DateGene:=V_PGI.DateEntree ;
   END else
   BEGIN
   DD:=StrToDate(DATEGENERATION.Text) ; Err:=DateCorrecte(DD) ;
   if Err>0 then
      BEGIN
      HMRegul.Execute(5+Err,caption,'') ;
      DATEGENERATION.Text:=DateToStr(V_PGI.DateEntree) ;
      DateGene:=V_PGI.DateEntree ;
      END else
      BEGIN
      if RevisionActive(DD) then
         BEGIN
         DATEGENERATION.Text:=DateToStr(V_PGI.DateEntree) ;
         DateGene:=V_PGI.DateEntree ;
         END else
         BEGIN
         DateGene:=DD ;
         END ;
      END ;
   END ;
end;

{================================ REQUETE =========================================}
function TFRegulInv.WhereDetail ( RL : RLETTR ) : String ;
Var EL : String ;
BEGIN
if Regul=pglRegul then EL:='PL' else EL:='TL' ;
Result:=' E_AUXILIAIRE="'+RL.Auxiliaire+'"'+' AND E_GENERAL="'+RL.General+'"'
       +' AND E_ETATLETTRAGE="'+EL+'" AND E_LETTRAGE="'+RL.CodeLettre+'"' ;
if RL.LettrageOppose then Result:=Result+' AND E_LETTRAGEEURO="X"' ;
END ;

function TFRegulInv.WhereCrit : String ;
Var St : String ;
    i  : integer ;
    CC : THCpteEdit ;
BEGIN
St:=' AND E_DATEPAQUETMIN>="'+USDATETIME(StrToDate(E_DATECOMPTABLE.Text))+'"'
   +' AND E_DATEPAQUETMAX<="'+USDATETIME(StrToDate(E_DATECOMPTABLE_.Text))+'"' ;
if T_NATUREAUXI.Value<>'' then St:=St+' AND T_NATUREAUXI="'+T_NATUREAUXI.Value+'"' ;
Case Regul of
   pglRegul : if E_DEVISE.Value=V_PGI.DevisePivot then St:=St+' AND E_LETTRAGEDEV="-"' else
               if E_DEVISE.Value<>'' then St:=St+' AND E_LETTRAGEDEV="X"' ;
   pglEcart : BEGIN
              St:=St+' AND E_LETTRAGEDEV="X"' ;
              St:=St+' AND (E_DATEPAQUETMIN>="'+USDATETIME(V_PGI.DateDebutEuro)+'"'
                 +' OR E_DATEPAQUETMAX<"'+USDATETIME(V_PGI.DateDebutEuro)+'")' ;
              END ;
 pglConvert : ;
   END ;
if Regul<>pglConvert then
   BEGIN
   if ModeOppose.Checked then St:=St+' AND E_LETTRAGEEURO="X"' else St:=St+' AND E_LETTRAGEEURO="-"' ;
   END ;
if Plibres.TabVisible then for i:=0 to 9 do
   BEGIN
   CC:=THCpteEdit(FindComponent('T_TABLE'+IntToStr(i))) ; if CC=Nil then Continue ;
   if CC.Text<>'' then St:=St+' AND T_TABLE'+IntToStr(i)+'="'+CC.Text+'"' ;
   END ;
WhereCrit:=St ;
END ;

procedure TFRegulInv.TripoteTitres ;
Var DV : RDEVISE ;
BEGIN
if Regul=pglEcart then
   BEGIN
   GR.Columns.Items[4].Title.Caption:=HMRegul.Mess[22] ;
   GR.Columns.Items[6].Title.Caption:=HMRegul.Mess[23] ;
   GR.Columns.Items[7].Title.Caption:=HMRegul.Mess[24] ;
   END else if Regul=pglConvert then
   BEGIN
   if VH^.TenueEuro then
      BEGIN
      DV.Code:=V_PGI.DeviseFongible ; GetInfosDevise(DV) ;
      if DV.Symbole='' then DV.Symbole:=Copy(DV.Libelle,1,2) ;
      GR.Columns.Items[6].Title.Caption:=HMRegul.Mess[25]+' '+DV.Symbole ;
      GR.Columns.Items[9].Title.Caption:=HMRegul.Mess[26]+' '+DV.Symbole ;
      GR.Columns.Items[11].Title.Caption:=HMRegul.Mess[27]+' '+DV.Symbole ;
      END ;
   END ;
END ;

procedure TFRegulInv.RechercheEcritures ;
Var StXP,StXN,StXP2,StXN2,StXP3,StXN3 : String ;
    Ind                   : integer ;
BEGIN
GR.ClearSelected ;
QR.Close ; QR.SQL.Clear ;
if Regul=pglRegul then Ind:=1 else if Regul=pglEcart then Ind:=0 else Ind:=3 ;
QR.SQL.Add(SelectQDL(Ind,ModeOppose.Checked)+' Where ') ;
QR.SQL.Add(LWhereBase(Regul=pglRegul,True,False,False)) ;
if Regul=pglRegul then QR.SQL.Add(' AND E_ETATLETTRAGE="PL"') else QR.SQL.Add(' AND E_ETATLETTRAGE="TL"') ;
QR.SQL.Add(LWhereComptes(Self)) ;
QR.SQL.Add(WhereCrit) ;
QR.SQL.Add('Group by E_GENERAL, E_AUXILIAIRE, E_DEVISE, E_LETTRAGE') ;
if Regul=pglRegul then
   BEGIN
   if E_DEVISE.Value=V_PGI.DevisePivot then
      BEGIN
      if ModeOppose.Checked then
         BEGIN
         QR.SQL.Add('Having (SUM(E_DEBITEURO-E_CREDITEURO)>0 AND SUM(E_DEBITEURO-E_CREDITEURO)<'+StrFPoint(MAXDEBIT.Value)+')') ;
         QR.SQL.Add('OR (SUM(E_DEBITEURO-E_CREDITEURO)<0 AND SUM(E_DEBITEURO-E_CREDITEURO)>'+StrFPoint(-MAXCREDIT.Value)+')') ;
         END else
         BEGIN
         QR.SQL.Add('Having (SUM(E_DEBIT-E_CREDIT)>0 AND SUM(E_DEBIT-E_CREDIT)<'+StrFPoint(MAXDEBIT.Value)+')') ;
         QR.SQL.Add('OR (SUM(E_DEBIT-E_CREDIT)<0 AND SUM(E_DEBIT-E_CREDIT)>'+StrFPoint(-MAXCREDIT.Value)+')') ;
         END ;
      END else
      BEGIN
      QR.SQL.Add('Having (SUM(E_DEBITDEV-E_CREDITDEV)>0 AND SUM(E_DEBITDEV-E_CREDITDEV)<'+StrFPoint(MAXDEBIT.Value)+')') ;
      QR.SQL.Add('OR (SUM(E_DEBITDEV-E_CREDITDEV)<0 AND SUM(E_DEBITDEV-E_CREDITDEV)>'+StrFPoint(-MAXCREDIT.Value)+')') ;
      END ;
   END else if Regul=pglConvert then
   BEGIN
   StXP:=StrFPoint(9*Resolution(V_PGI.OkDecV+1))  ; StXN:=StrFPoint(-9*Resolution(V_PGI.OkDecV+1)) ;
   StXP2:=StrFPoint(9*Resolution(V_PGI.OkDecE+1)) ; StXN2:=StrFPoint(-9*Resolution(V_PGI.OkDecE+1)) ;
   StXP3:=StrFPoint(9*Resolution(DEV.Decimale+1)) ; StXN3:=StrFPoint(-9*Resolution(DEV.Decimale+1)) ;
   if EstMonnaieIN(DEV.Code) then
      BEGIN
      QR.SQL.Add('Having ((') ;
      QR.SQL.Add('Sum(E_COUVERTUREEURO*'+SoldeSurDC+') Not Between '+StXN2+' AND '+StXP2) ;
      QR.SQL.Add(') OR (') ;
      QR.SQL.Add('Sum(E_COUVERTURE*'+SoldeSurDC+') Not Between '+StXN+' AND '+StXP) ;
      QR.SQL.Add(')) AND (') ;
      QR.SQL.Add('Sum(E_COUVERTUREDEV*'+SoldeSurDC+') Between '+StXN3+' AND '+StXP3) ;
      QR.SQL.Add(')') ;
      END else
      BEGIN
      if VH^.TenueEuro then
         BEGIN
         QR.SQL.Add('Having (') ;
         QR.SQL.Add('Sum(E_COUVERTURE*'+SoldeSurDC+') Not Between '+StXN2+' AND '+StXP2) ;
         QR.SQL.Add(') AND (') ;
         QR.SQL.Add('Sum(E_COUVERTUREEURO*'+SoldeSurDC+') Between '+StXN+' AND '+StXP) ;
         QR.SQL.Add(')') ;
         END else
         BEGIN
         QR.SQL.Add('Having (') ;
         QR.SQL.Add('Sum(E_COUVERTUREEURO*'+SoldeSurDC+') Not Between '+StXN2+' AND '+StXP2) ;
         QR.SQL.Add(') AND (') ;
         QR.SQL.Add('Sum(E_COUVERTURE*'+SoldeSurDC+') Between '+StXN+' AND '+StXP) ;
         QR.SQL.Add(')') ;
         END ;
      END ;
   END else {écart change}
   BEGIN
   if VH^.TenueEuro then
      BEGIN
      StXP:=StrFPoint(9*Resolution(V_PGI.OkDecV+1)) ; StXN:=StrFPoint(-9*Resolution(V_PGI.OkDecV+1)) ;
      QR.SQL.Add('Having Sum(E_COUVERTURE*'+SoldeSurDC+') Not Between '+StXN+' AND '+StXP) ;
      END else
      BEGIN
      StXP:=StrFPoint(9*Resolution(V_PGI.OkDecE+1)) ; StXN:=StrFPoint(-9*Resolution(V_PGI.OkDecE+1)) ;
      QR.SQL.Add('Having Sum(E_COUVERTUREEURO*'+SoldeSurDC+') Not Between '+StXN+' AND '+StXP) ;
      END ;
   END ;
ChangeSQL(QR) ; QR.Open ; ChangeAspectLettre(GR) ;
GereSelectionsGrid(GR,QR) ;
TripoteTitres ;
END ;

{============================== AFFICHAGES ========================================}
procedure TFRegulInv.FormShow(Sender: TObject);
begin
Pages.ActivePage:=Princ ;
InitTablesLibresTiers(PLibres) ;
GeneCharge:=True ;
if VH^.TenueEuro then ModeOppose.Caption:=RechDom('TTDEVISETOUTES',V_PGI.DeviseFongible,False) ;
if Regul=pglEcart then BEGIN ModeOppose.Visible:=False ; ExpliEuro.Visible:=False ; END ;
if Regul=pglConvert then ExpliEuro.Caption:=HMRegul.Mess[20] ;
FillChar(DEV,Sizeof(DEV),#0) ; FillChar(J,Sizeof(J),#0) ;
if Regul=pglRegul then
   BEGIN
   Caption:=HMRegul.Mess[0] ; Journal.DataType:='ttJalRegul' ;
   H_JOURNALR.Visible:=True ; HelpContext:=7520000 ;
   RefRegul.Text:=HMRegul.Mess[28] ; LibRegul.Text:=HMRegul.Mess[28] ;
   END else if Regul=pglEcart then
   BEGIN
   Caption:=HMRegul.Mess[1] ; Journal.DataType:='ttJalEcart' ;
   H_JOURNALE.Visible:=True ; HelpContext:=7523000 ;
   MaxDebit.Visible:=False  ; H_SD.Visible:=False ;
   MaxCredit.Visible:=False ; H_SC.Visible:=False ;
   RefRegul.Text:=HMRegul.Mess[29] ; LibRegul.Text:=HMRegul.Mess[29] ;
   END else
   BEGIN
   Caption:=HMRegul.Mess[18] ; Journal.DataType:='ttJournal' ; Journal.Enabled:=False ;
   H_JOURNALC.Visible:=True ; Journal.Value:=VH^.JalEcartEuro ;
   MaxDebit.Visible:=False  ; H_SD.Visible:=False ;
   MaxCredit.Visible:=False ; H_SC.Visible:=False ;
   BValider.Hint:=HMRegul.Mess[21] ; HelpContext:=7520500 ;
   ModeOppose.Visible:=False ; ExpliEuro.Visible:=False ;
   RefRegul.Text:=HMRegul.Mess[30] ; LibRegul.Text:=HMRegul.Mess[30] ;
   END ;
if Regul=pglRegul then
   BEGIN
   E_DEVISE.DataType:='ttDevise' ;
   E_DEVISE.Value:=V_PGI.DevisePivot ;
   QR.Liste:='REGULLETEURO' ;
   END else if Regul=pglEcart then
   BEGIN
   E_DEVISE.DataType:='ttDeviseEtat' ; E_DEVISE.Value:='' ;
   if E_DEVISE.Items.Count>0 then BEGIN E_DEVISE.ItemIndex:=0 ; E_DEVISEChange(Nil) ; END ;
   QR.Liste:='REGULLETEURO' ;
   END else
   BEGIN
   E_DEVISE.DataType:='ttDevise' ;
   E_DEVISE.Value:=V_PGI.DevisePivot ;
   QR.Liste:='ECARTCONVERTEURO' ;
   END ;
DateGeneration.Text:=DateToStr(V_PGI.DateEntree) ; DateGene:=V_PGI.DateEntree ;
ChargeFiltre(TitreF,FFiltres,Pages) ; InitCriteres ;
if FFiltres.Text='DEFAUT' then BChercherClick(Nil) else
   BEGIN
   ChargeAVide ; E_AUXILIAIRE.SetFocus ;
   if Regul<>pglConvert then
      BEGIN
      if JOURNAL.Items.Count>0 then BEGIN JOURNAL.ItemIndex:=0 ; JournalChange(Nil) ; END ;
      END else
      BEGIN
      Journal.Value:=VH^.JalEcartEuro ;
      END ;
   END ;
GeneCharge:=False ;
UpdateCaption(Self) ;
Tripotetitres ;
end;

procedure TFRegulInv.ChargeAVide ;
Var ind : integer ;
    St : String ;
BEGIN
QR.Close ; QR.SQL.Clear ;
St:='' ;
If Qui<>prAucun Then St:=' Left Outer Join GENERAUX On E_GENERAL=G_GENERAL ' ;
if Regul=pglRegul then Ind:=0 else if Regul=pglEcart then Ind:=1 else Ind:=3 ;
QR.SQL.Add(SelectQDL(Ind,ModeOppose.Checked)+ 'Where '+LWhereVide) ;
If MonProfilOk(Qui) Then
  BEGIN
  St:=WhereProfilUser(QR,Qui) ; QR.SQL.Add(' AND '+St) ;
  END ;
QR.SQL.Add('Group by E_GENERAL, E_AUXILIAIRE, E_DEVISE, E_LETTRAGE') ;
ChangeSQL(QR) ; QR.Open ; ChangeAspectLettre(GR) ;
END ;

procedure TFRegulInv.FormClose(Sender: TObject; var Action: TCloseAction);
Var Ind : integer ;
begin
if TPieces.Count>0 then
   BEGIN
   if Regul=pglRegul then Ind:=3 else if Regul=pglEcart then Ind:=4 else Ind:=14 ;
   if HMRegul.Execute(13,caption,'')=mrYes then VisuPiecesGenere(TPieces,EcrGen,Ind) ;
   END ;
VideListe(TPieces) ; TPieces.Free ;
if Parent is THPanel then
   BEGIN
   If Not MonProfilOk(Qui) Then _Bloqueur('nrBatch',False) ;
   Action:=caFree ;
   END ;
end;

{============================= TRAITEMENT DE REGUL ======================================}
procedure TFRegulInv.RemplirLeM ( Var M : RMVT ; RL : RLETTR ) ;
Var Facturier : String3 ;
BEGIN
FillChar(M,Sizeof(M),#0) ;
M.Jal:=JOURNAL.Value ; Facturier:=J.Facturier ;
M.ModeSaisieJal:=J.ModeSaisie ; 
//SetIncNum(EcrGen,Facturier,M.Num) ;
M.Axe:='' ; M.Etabl:=VH^.EtablisDefaut ; M.ModeOppose:=ModeOppose.Checked ;
if ((Regul=pglRegul) and (RL.LettrageDevise)) or (Regul=pglConvert)
    then BEGIN M.CodeD:=DEV.Code ; M.TauxD:=DEV.Taux ; END
    else BEGIN M.CodeD:=V_PGI.DevisePivot ; M.TauxD:=1 ; END ;
M.DateTaux:=V_PGI.DateEntree ; M.Valide:=False ; M.Simul:='N' ;
if Regul in [pglRegul,pglConvert] then M.Nature:='OD' else M.Nature:='ECC' ;
M.DateC:=DateGene ; M.Exo:=QuelExo(DateToStr(M.DateC)) ; M.General:='' ;
SetIncNum(EcrGen,Facturier,M.Num,M.DateC) ;
END ;

procedure TFRegulInv.CouvreRegul ( O : TOBM ; RL : RLETTR ) ;
Var TotD,TotF,TotE,Taux : double ;
BEGIN
TotF:=O.GetMvt('E_DEBIT')+O.GetMvt('E_CREDIT') ;
TotE:=O.GetMvt('E_DEBITEURO')+O.GetMvt('E_CREDITEURO') ;
TotD:=O.GetMvt('E_DEBITDEV')+O.GetMvt('E_CREDITDEV') ;
Case Regul of
   pglRegul : BEGIN
              if RL.LettrageDevise then
                 BEGIN
                 O.PutMvt('E_COUVERTUREDEV',TotD) ; Taux:=O.GetMvt('E_TAUXDEV') ;
                 if VH^.TenueEuro then
                    BEGIN
                    if Arrondi(O.GetMvt('E_COUVERTURE')-TotE,V_PGI.OkDecV)<>0 then O.PutMvt('E_COUVERTURE',DeviseToEuro(TotD,Taux,DEV.Quotite)) ;
                    if Arrondi(O.GetMvt('E_COUVERTUREEURO')-TotF,V_PGI.OkDecE)<>0 then O.PutMvt('E_COUVERTUREEURO',DeviseToPivot(TotD,Taux,DEV.Quotite)) ;
                    END else
                    BEGIN
                    if Arrondi(O.GetMvt('E_COUVERTURE')-TotF,V_PGI.OkDecV)<>0 then O.PutMvt('E_COUVERTURE',DeviseToPivot(TotD,Taux,DEV.Quotite)) ;
                    if Arrondi(O.GetMvt('E_COUVERTUREEURO')-TotE,V_PGI.OkDecE)<>0 then O.PutMvt('E_COUVERTUREEURO',DeviseToEuro(TotD,Taux,DEV.Quotite)) ;
                    END ;
                 END else
                 BEGIN
                 if Not RL.LettrageOppose then
                    BEGIN
                    O.PutMvt('E_COUVERTURE',TotF) ; ConvertCouverture(O,tsmPivot) ;
                    END else
                    BEGIN
                    O.PutMvt('E_COUVERTUREEURO',TotE) ; ConvertCouverture(O,tsmEuro) ;
                    END ;
                 END ;
              END ;
   pglEcart : BEGIN
              {Ecart de change couvre tout}
              O.PutMvt('E_COUVERTUREDEV',TotD) ;
              O.PutMvt('E_COUVERTURE',TotF) ;
              O.PutMvt('E_COUVERTUREEURO',TotE) ;
              END ;
 pglConvert : BEGIN
              {Ecart de conversion couvre tout}
              O.PutMvt('E_COUVERTUREDEV',TotD) ;
              O.PutMvt('E_COUVERTURE',TotF) ;
              O.PutMvt('E_COUVERTUREEURO',TotE) ;
              END ;
   END ;
 END ;

procedure TFRegulInv.GenereRegul ;
Var RL : RLETTR ;
    Q  : TQuery ;
    X,LimiteP,LimiteE,Montant,MenOppose,TotF,TotD,TotE,Taux : Double ;
    St,Cpte,StInd : String ;
    O,OOO     : TOBM ;
    Premier   : boolean ;
    M         : RMVT ;
    XX        : T_ECARTCHANG ;
    DMin,DMax : TDateTime ;
BEGIN
if Regul<>pglConvert then
   BEGIN
   Montant:=Arrondi(GR.Fields[7].AsFloat,V_PGI.OkDecV) ;
   MenOppose:=0 ;
   END else
   BEGIN
   Montant:=Arrondi(GR.Fields[9].AsFloat,V_PGI.OkDecV) ;
   MenOppose:=Arrondi(GR.Fields[10].AsFloat,V_PGI.OkDecE) ;
   END ;
DMax:=0 ; DMin:=0 ;
LimiteP:=Resolution(V_PGI.OkDecV) ; LimiteE:=Resolution(V_PGI.OkDecE) ;
if Regul<>pglConvert then
   BEGIN
   if Abs(Montant)<LimiteP then Exit ;
   END else
   BEGIN
   if ((Abs(Montant)<LimiteP) and (Abs(MenOppose)<LimiteE)) then Exit ;
   END ;
Premier:=True ; FillChar(RL,Sizeof(RL),#0) ;
RL.General:=GR.Fields[0].AsString ;
RL.Auxiliaire:=GR.Fields[1].AsString ; RL.Appel:=tlMenu ; RL.GL:=Nil ;
if Regul<>pglConvert then
   BEGIN
   RL.CodeLettre:=GR.Fields[8].AsString ;
   RL.DeviseMvt:=GR.Fields[9].AsString  ;
   RL.LettrageDevise:=(RL.DeviseMvt<>V_PGI.DevisePivot) ;
   RL.LettrageOppose:=(GR.Fields[11].AsString='X') ;
   END else
   BEGIN
   RL.CodeLettre:=GR.Fields[11].AsString ;
   RL.DeviseMvt:=GR.Fields[12].AsString  ;
   RL.LettrageDevise:=(RL.DeviseMvt<>V_PGI.DevisePivot) ;
   RL.LettrageOppose:=(GR.Fields[11].AsString='X') ;
   END ;
if ((Regul=pglRegul) and (RL.LettrageDevise)) then
   BEGIN
   {Récupérer l'écart en devise}
   Q:=OpenSQL('Select SUM(E_DEBITDEV-E_CREDITDEV) from Ecriture Where '+WhereDetail(RL),True) ;
   if Not Q.EOF then Montant:=Q.Fields[0].AsFloat ;
   Ferme(Q) ;
   END ;
if Regul=pglEcart then
   BEGIN
   {Récupérer l'écart opposé aussi}
   if VH^.TenueEuro then
      BEGIN
      Q:=OpenSQL('Select SUM(E_DEBITEURO-E_CREDITEURO) from Ecriture Where '+WhereDetail(RL),True) ;
      if Not Q.EOF then MenOppose:=Arrondi(Q.Fields[0].AsFloat,V_PGI.OkDecE) ;
      Ferme(Q) ;
      END else
      BEGIN
      Q:=OpenSQL('Select SUM(E_DEBIT-E_CREDIT) from Ecriture Where '+WhereDetail(RL),True) ;
      if Not Q.EOF then MenOppose:=Arrondi(Q.Fields[0].AsFloat,V_PGI.OkDecE) ;
      Ferme(Q) ;
      END ;
   END ;
if Regul<>pglConvert then BEGIN if Montant=0 then Exit ; END
                     else BEGIN if ((Montant=0) and (MenOppose=0)) then Exit ; END ;
if Regul=pglConvert then
   BEGIN
   if Not VH^.TenueEuro then BEGIN X:=MenOppose ; MenOppose:=Montant ; Montant:=X ; END ;
   END ;
Q:=OpenSQL('Select * from Ecriture Where '+WhereDetail(RL),True) ;
While Not Q.EOF do
   BEGIN
   O:=TOBM.Create(EcrGen,'',False) ; O.ChargeMvt(Q) ;
   O.PutMvt('E_ETATLETTRAGE','TL') ; O.PutMvt('E_LETTRAGE',uppercase(O.GetMvt('E_LETTRAGE'))) ;
   {Etat}
   St:=O.GetMvt('E_ETAT') ;
   if ((Montant>0) or ((Montant=0) and (MenOppose>0)))
      then StInd:=IntToStr(CPTDEBIT.ItemIndex+1)
      else StInd:=IntToStr(CPTCREDIT.ItemIndex+1) ;
   if Length(St)>=4 then
      BEGIN
      if Regul=pglRegul then St[1]:='X' else
         BEGIN
         if Regul=pglEcart then St[2]:='X' else St[2]:='#' ;
         St[3]:='-' ;
         END ;
      St[4]:=StInd[1] ;
      END ;
   O.PutMvt('E_ETAT',St) ;
   CouvreRegul(O,RL) ;
   {Dates}
   if Premier then
      BEGIN
      DMin:=O.GetMvt('E_DATEPAQUETMIN') ; if DateGene<DMin then DMin:=DateGene ;
      DMax:=O.GetMvt('E_DATEPAQUETMAX') ; if DateGene>DMax then DMax:=DateGene ;
      END ;
   O.PutMvt('E_DATEPAQUETMIN',DMin) ; O.PutMvt('E_DATEPAQUETMAX',DMax) ;
   if Not GoReqMajLet(O,RL.General,RL.Auxiliaire,NowFutur,True) then V_PGI.IoError:=oeUnknown ;
   if Premier then
      BEGIN
      RemplirLeM(M,RL) ;
      if Regul<>pglConvert then
         BEGIN
         if Montant>0 then Cpte:=CPTDEBIT.Text else Cpte:=CptCredit.Text ;
         END else
         BEGIN
         if (Montant>0) or ((Montant=0) and (MenOppose>0)) then Cpte:=CPTDEBIT.Text else Cpte:=CptCredit.Text ;
         END ;
      XX.Cpte:=Cpte ; XX.Regul:=Regul ; XX.Ref:=REFREGUL.Text ;
      XX.Lib:=LIBREGUL.Text ; XX.Quotite:=DEV.Quotite ; XX.Cpte:=Cpte ;
      XX.DPMin:=DMin ; XX.DPMax:=DMax ; XX.Decimale:=DEV.Decimale ;
      Case Regul of
         pglRegul   : BEGIN XX.Montant1:=Montant ; XX.Montant2:=0 ; END ;
         pglEcart   : BEGIN XX.Montant1:=Montant ; XX.Montant2:=MenOppose ; END ;
         pglConvert : BEGIN XX.Montant1:=Montant ; XX.Montant2:=MenOppose ; END ;
         END ;
      OOO:=CreerPartieDoubleLett(M,RL,XX,O) ;
      TPieces.Add(OOO) ;
      END ;
   Q.Next ; Premier:=False ; O.Free ; O:=Nil ;
   END ;
Ferme(Q) ;
END ;

procedure TFRegulInv.ValideLaRegul ;
Var i : integer ;
BEGIN
InitMove(GR.NbSelected,'') ;
for i:=0 to GR.NbSelected-1 do
    BEGIN
    GR.GotoLeBookMark(i) ; MoveCur(FALSE) ;
    if V_PGI.IoError=oeOK then GenereRegul else Break ;
    END ;
FiniMove ;
END ;

procedure TFRegulInv.BValiderClick(Sender: TObject);
Var Num : integer ;
begin
NowFutur:=NowH ;
if JOURNAL.Value='' then BEGIN HMRegul.Execute(3,caption,'') ; Exit ; END ;
if ((CPTDEBIT.Text='') or (CPTCREDIT.Text='')) then BEGIN HMRegul.Execute(4,caption,'') ; Exit ; END ;
if Not Presence('GENERAUX','G_GENERAL',CPTDEBIT.Text) then BEGIN HMRegul.Execute(10,caption,'') ; Exit ; END ;
if Not Presence('GENERAUX','G_GENERAL',CPTCREDIT.Text) then BEGIN HMRegul.Execute(11,caption,'') ; Exit ; END ;
if RevisionActive(DateGene) then Exit ;
if GR.NbSelected<=0 then BEGIN HMRegul.Execute(2,caption,'') ; Exit ; END ;
if Regul=pglRegul then Num:=12 else if Regul=pglEcart then Num:=14 else Num:=19 ;
if HMRegul.Execute(Num,caption,'')<>mrYes then Exit ;
if Transactions(ValideLaRegul,5)<>oeOK then MessageAlerte(HMRegul.Mess[15]) ;
BChercherClick(Nil) ;
end;

procedure TFRegulInv.BNouvRechClick(Sender: TObject);
Var CodeD : String ;
begin
CodeD:=E_DEVISE.Value ;
VideFiltre(FFiltres,Pages) ;
E_DEVISE.Value:=CodeD ;
InitCriteres ;
end;

procedure TFRegulInv.InitCriteres ;
Var DD : TDateTime ;
BEGIN
if VH^.Precedent.Code<>'' then E_DATECOMPTABLE.Text:=DateToStr(VH^.Precedent.Deb)
                          else E_DATECOMPTABLE.Text:=DateToStr(VH^.Encours.Deb) ;
E_DATECOMPTABLE_.Text:=DateToStr(V_PGI.DateEntree) ;
if ((VH^.ExoV8.Code<>'') and (StrToDate(E_DATECOMPTABLE.Text)<VH^.ExoV8.Deb)) then E_DATECOMPTABLE.Text:=DateToStr(VH^.ExoV8.Deb) ;
DD:=StrToDate(DATEGENERATION.Text) ;
if ModeRevisionActive(DD) then
   BEGIN
   DATEGENERATION.Text:=DateToText(V_PGI.DateEntree) ;
   DateGene:=V_PGI.DateEntree ;
   END ;
END ;

procedure TFRegulInv.WMGetMinMaxInfo(var MSG: Tmessage);
begin
with PMinMaxInfo(MSG.lparam)^.ptMinTrackSize do begin X := WMinX ; Y := WMinY ; end;
end;

procedure TFRegulInv.BAideClick(Sender: TObject);
begin
CallHelpTopic(Self) ;
end;

procedure TFRegulInv.FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
if ((Shift=[]) and (Key=VK_F9)) then BEGIN Key:=0 ; BChercherClick(Nil) ; END ;
end;

procedure TFRegulInv.POPFPopup(Sender: TObject);
begin
UpdatePopFiltre(BSaveFiltre,BDelFiltre,BRenFiltre,FFiltres) ;
end;

end.
