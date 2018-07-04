unit LetRegul;

interface

uses   
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  {$IFNDEF EAGLCLIENT}
   DBGrids,
   DB,
   {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}
   PrintDBG,
  {$ELSE}

   UTOBDEbug,
  {$ENDIF}
  paramsoc,
  StdCtrls, Hcompte, hmsgbox, Grids, Buttons, Hctrls, Mask, HEnt1,
  ExtCtrls, ComCtrls, LettUtil,  Hqry, SaisUtil, SaisComm,
  Ent1, CPJournal_TOM, Menus, HSysMenu, HDB, HTB97,DelVisuE,  UTOB,
  Letbatch, HStatus, UObjFiltres {SG6 12/01/05 Gestion Filtres V6 FQ 15255}, ed_tools, HPanel, UiUtil, UtilPGI,
  TntGrids, TntStdCtrls ;

type
  TFLetRegul = class(TForm)
    Pages: TPageControl;
    Princ: TTabSheet;
    HPB: TToolWindow97;
    BImprimer: TToolbarButton97;
    BValider: TToolbarButton97;
    BFerme: TToolbarButton97;
    BAide: TToolbarButton97;
    BRecherche: TToolbarButton97;
    HMRegul: THMsgBox;
    Cache: THCpteEdit;
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
    ExpliEuro: TLabel;
    Dock: TDock97;
    Dock971: TDock97;
    XX_WHERESEL: TEdit;
    bSelectAll: TToolbarButton97;
    GR: THGrid;
    procedure E_AUXILIAIREDblClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure E_DEVISEChange(Sender: TObject);
    procedure BChercherClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
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
    procedure BAgrandirClick(Sender: TObject);
    procedure BReduireClick(Sender: TObject);
    procedure BAideClick(Sender: TObject);
    procedure bSelectAllClick(Sender: TObject);
    procedure FormKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);
  private
    //SG6 12/01/05 Gestion filtres V6 FQ 15255
    OBjFiltre : TObjFiltre;
    //QR : TQuery ;
    FTOB : TOB ;
    DEV : RDEVISE ;
    JalReG : REGJAL ;
    DateGene,NowFutur : TDateTime ;
    FindFirst,GeneCharge : boolean ;
    TPIECES   : TList ;
    TitreF    : String ;
    WMinX,WMinY : Integer ;
    Qui : tProfilTraitement ;
    Auto : Boolean ;
    FBoRegulAnc : boolean ;
    procedure WMGetMinMaxInfo(var MSG: Tmessage); message WM_GetMinMaxInfo;
    procedure AttribCollectif ;
    procedure RechercheEcritures ;
    procedure GenereRegul (vTOB : TOB ) ;
    function  WhereCrit : String ;
    function  WhereDetail ( RL : RLETTR ) : String ;
    procedure RemplirLeM ( Var M : RMVT ; RL : RLETTR ) ;
    procedure ValideLaRegul ;
    procedure AlimXX_WHERESEL ;
    procedure InitCriteres ;
    procedure CouvreRegul ( O : TOBM ; RL : RLETTR ) ;
    procedure TripoteTitres ;
  public
    Regul : PG_LETTRAGE ;
  end;

procedure RegulLettrage ( Regul,Convert : Boolean ; RegulAnc : boolean = false ) ;
procedure RegulInverse ;
procedure RegulLettrageMP ( Regul,Convert : Boolean ; Qui : tProfilTraitement) ;
procedure RegulLettrageMPAuto ( Regul,Convert : Boolean ) ;

implementation

uses
  {$IFDEF MODENT1}
  CPProcGen,
  CPVersion,
  {$ENDIF MODENT1}
  Paramdat;

{$R *.DFM}

procedure RegulLettrage ( Regul,Convert : Boolean ; RegulAnc : boolean = false ) ;
Var X  : TFLetRegul ;
    PP : THPanel ;
BEGIN
if PasCreerDate(V_PGI.DateEntree) then Exit ;
if _Blocage(['nrCloture','nrBatch','nrSaisieModif','nrLettrage','nrEnca','nrDeca'],True,'nrBatch') then Exit ;
X:=TFLetRegul.Create(Application) ;
if RegulAnc then
 begin
  X.FBoRegulAnc  := true ;
  X.Regul        := pglRegul ;
  X.TitreF       :='REGULLETTRAGE';
 end
  else
   if Convert then BEGIN X.Regul:=pglConvert ; X.TitreF:='REGULCONVERT' ; END else
    if Regul then BEGIN X.Regul:=pglRegul ; X.TitreF:='REGULLETTRAGE'; END else
      BEGIN X.Regul:=pglEcart ; X.TitreF:='ECARTCHANGE' ; END ;
PP:=FindInsidePanel ;
if PP=Nil then
   BEGIN
    Try
     X.Auto:=FALSE ;
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

procedure RegulInverse ;
Var X  : TFLetRegul ;
    PP : THPanel ;
BEGIN
if PasCreerDate(V_PGI.DateEntree) then Exit ;
if _Blocage(['nrCloture','nrBatch','nrSaisieModif','nrLettrage','nrEnca','nrDeca'],True,'nrBatch') then Exit ;
X:=TFLetRegul.Create(Application) ;
X.Regul:=pglInverse ; X.TitreF:='REGULINVERSE' ;
PP:=FindInsidePanel ;
if PP=Nil then
   BEGIN
    Try
     X.Qui:=prAucun ;
     X.Auto:=FALSE ;
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

procedure RegulLettrageMP ( Regul,Convert : Boolean ; Qui : tProfilTraitement) ;
Var X  : TFLetRegul ;
    PP : THPanel ;
BEGIN
if PasCreerDate(V_PGI.DateEntree) then Exit ;
If Not MonProfilOk(Qui) Then
  BEGIN
  if _Blocage(['nrCloture','nrBatch','nrSaisieModif','nrLettrage','nrEnca','nrDeca'],True,'nrBatch') then Exit ;
  END ;
X:=TFLetRegul.Create(Application) ;
if Convert then BEGIN X.Regul:=pglConvert ; X.TitreF:='REGULCONVERT' ; END else
   if Regul then BEGIN X.Regul:=pglRegul ; X.TitreF:='REGULLETTRAGE'; END else
      BEGIN X.Regul:=pglEcart ; X.TitreF:='ECARTCHANGE' ; END ;
PP:=FindInsidePanel ;
if PP=Nil then
   BEGIN
    Try
     X.Auto:=FALSE ;
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

procedure RegulLettrageMPAuto ( Regul,Convert : Boolean ) ;
Var X  : TFLetRegul ;
    PP : THPanel ;
BEGIN
if PasCreerDate(V_PGI.DateEntree) then Exit ;
X:=TFLetRegul.Create(Application) ;
if Convert then BEGIN X.Regul:=pglConvert ; X.TitreF:='REGULCONVERT' ; END else
   if Regul then BEGIN X.Regul:=pglRegul ; X.TitreF:='REGULLETTRAGE'; END else
      BEGIN X.Regul:=pglEcart ; X.TitreF:='ECARTCHANGE' ; END ;
PP:=FindInsidePanel ;
if PP=Nil then
   BEGIN
    Try
     X.Auto:=TRUE ;
     X.Qui:=prAucun ;
     X.ShowModal ;
    Finally
     X.Free ;
    End ;
   Screen.Cursor:=SyncrDefault ;
   END else
   BEGIN
   InitInside(X,PP) ;
   X.Show ;
   END ;
END ;



procedure TFLetRegul.BAgrandirClick(Sender: TObject);
begin ChangeListeCrit(Self,True) ; end;

procedure TFLetRegul.BReduireClick(Sender: TObject);
begin ChangeListeCrit(Self,False) ; end;

procedure TFLetRegul.E_DATECOMPTABLEKeyPress(Sender: TObject; var Key: Char);
begin ParamDate(Self,Sender,Key) ; end;

procedure TFLetRegul.H_JOURNALRDblClick(Sender: TObject);
Var A : TActionFiche ;
begin
if Not ExJaiLeDroitConcept(TConcept(ccJalModif),False) then A:=taConsult else A:=taModif ;
if JOURNAL.Value<>'' then
   BEGIN
   FicheJournal(Nil,'',JOURNAL.Value,A,0) ; if A=taModif then JournalChange(Nil) ;
   END ;
end;

procedure TFLetRegul.E_AUXILIAIREDblClick(Sender: TObject);
begin
AttribCollectif ;
end;

procedure TFLetRegul.BChercherClick(Sender: TObject);
begin
if ((VH^.ExoV8.Code<>'') and (StrToDate(E_DATECOMPTABLE.Text)<VH^.ExoV8.Deb)) then E_DATECOMPTABLE.Text:=DateToStr(VH^.ExoV8.Deb) ;
RechercheEcritures ;
end;

procedure TFLetRegul.BGenereClick(Sender: TObject);
Var Indice : integer ;
begin
if (Regul=pglRegul) then Indice:=3 else
   if Regul=pglEcart then Indice:=4 else Indice:=14 ;
  // FQ 12785 (SBO) suite au passage full euro, plus de visu de pièces générer Si on est en ecart de change
  if Regul<>pglEcart then
    if TPieces.Count>0 then
      VisuPiecesGenere(TPieces,EcrGen,Indice) ;
  // Fin FQ 12785
end;

procedure TFLetRegul.FormCreate(Sender: TObject);
var
  Composants : TControlFiltre;
begin
Composants.PopupF   := POPF;
Composants.Filtres  := FFILTRES;
Composants.Filtre   := BFILTRE;
Composants.PageCtrl := Pages;
ObjFiltre := TObjFiltre.Create(Composants,TitreF);

TPIECES:=TList.Create ; FTOB := TOB.Create('',nil,-1) ;
PopUpMenu:=ADDMenuPop(PopUpMenu,'','') ;
WMinX:=Width ; WMinY:=252 ;
end;

procedure TFLetRegul.BRechercheClick(Sender: TObject);
begin FindFirst:=True ; FindRegul.Execute ; end;

procedure TFLetRegul.FindRegulFind(Sender: TObject);
begin Rechercher(GR,FindRegul,FindFirst) ; end;

procedure TFLetRegul.BFermeClick(Sender: TObject);
begin
  Close ;
  //SG6 08.02.05 FQ 15328
  if IsInside(Self) then
  begin
    CloseInsidePanel(Self) ;
  end;
end;

procedure TFLetRegul.BImprimerClick(Sender: TObject);
begin
{$IFNDEF EAGLCLIENT}
 PrintDBGrid(GR,Nil,Caption,'') ;
{$ENDIF}
end;

{================================ ENTETE =========================================}
procedure TFLetRegul.AttribCollectif ;
Var Q : TQuery ;
BEGIN
if ((E_AUXILIAIRE_.Text<>'') or (E_GENERAL.Text<>'') or
    (E_GENERAL_.Text<>'') or (E_AUXILIAIRE.Text='')) then Exit ;
Q:=OpenSQL('Select T_COLLECTIF from Tiers Where T_AUXILIAIRE="'+E_AUXILIAIRE.Text+'"',True) ;
if Not Q.EOF then E_GENERAL.Text:=Q.Fields[0].AsString ; Ferme(Q) ;
END ;

procedure TFLetRegul.E_DEVISEChange(Sender: TObject);
Var DEVF : RDEVISE ;
begin
if E_DEVISE.Value=DEV.Code then Exit ; DEV.Code:=E_DEVISE.Value ;
GetInfosDevise(DEV) ; DEV.Taux:=GetTaux(DEV.Code,DEV.DateTaux,V_PGI.DateEntree) ;
ChangeMask(MAXDEBIT,DEV.Decimale,DEV.Symbole) ;
ChangeMask(MAXCREDIT ,DEV.Decimale,DEV.Symbole) ;
DEVF:=DEV ;
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

procedure TFLetRegul.JOURNALChange(Sender: TObject);
Var i : integer ;
begin
if JOURNAL.Value=JalReG.Journal then Exit ;
RemplirInfosRegul(JalReG,JOURNAL.Value) ;
if ((Journal.Value<>'') and (JalReG.Facturier='') and (Regul<>pglConvert) and (Regul<>pglInverse)) then
   BEGIN
   if Not GeneCharge then HMRegul.Execute(16,'','') ;
   Journal.Value:='' ;
   Exit ;
   END ;
if (Regul=pglRegul) then
   BEGIN
   CptDebit.Items.Clear ; CptCredit.Items.Clear ;
   for i:=1 to 3 do if JalReG.D[i]<>'' then CptDebit.Items.Add(JalReG.D[i]) ;
   for i:=1 to 3 do if JalReG.C[i]<>'' then CptCredit.Items.Add(JalReG.C[i]) ;
   if CptDebit.Items.Count>0 then CptDebit.ItemIndex:=0 ;
   if CptCredit.Items.Count>0 then CptCredit.ItemIndex:=0 ;
   END else if Regul in [pglConvert,pglInverse] then
   BEGIN
   CptDebit.Items.Clear ; CptCredit.Items.Clear ;
 //  CptDebit.Items.Add(VH^.EccEuroDebit) ;
//   CptCredit.Items.Add(VH^.EccEuroCredit) ;
   if CptDebit.Items.Count>0 then CptDebit.ItemIndex:=0 ;
   if CptCredit.Items.Count>0 then CptCredit.ItemIndex:=0 ;
   END ;
end;

procedure TFLetRegul.DATEGENERATIONExit(Sender: TObject);
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
function TFLetRegul.WhereDetail ( RL : RLETTR ) : String ;
Var EL : String ;
BEGIN
if (Regul=pglRegul) then EL:='PL' else EL:='TL' ;
Result:=' E_AUXILIAIRE="'+RL.Auxiliaire+'"'+' AND E_GENERAL="'+RL.General+'"'
       +' AND E_ETATLETTRAGE="'+EL+'" AND E_LETTRAGE="'+RL.CodeLettre+'"' ;
END ;

function TFLetRegul.WhereCrit : String ;
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
if Plibres.TabVisible then for i:=0 to 9 do
   BEGIN
   CC:=THCpteEdit(FindComponent('T_TABLE'+IntToStr(i))) ; if CC=Nil then Continue ;
   if CC.Text<>'' then St:=St+' AND T_TABLE'+IntToStr(i)+'="'+CC.Text+'"' ;
   END ;
WhereCrit:=St ;
END ;

procedure TFLetRegul.TripoteTitres ;
Var DV : RDEVISE ;
BEGIN
if Regul=pglEcart then
   BEGIN
   GR.Cells[0,4]:=HMRegul.Mess[22] ;
   GR.Cells[0,6]:=HMRegul.Mess[23] ;
   GR.Cells[0,7]:=HMRegul.Mess[24] ;
   END else if Regul in [pglConvert,pglInverse] then
   BEGIN
   if VH^.TenueEuro then
      BEGIN
      DV.Code:=V_PGI.DeviseFongible ; GetInfosDevise(DV) ;
      if DV.Symbole='' then DV.Symbole:=Copy(DV.Libelle,1,2) ;
      GR.Cells[0,6]:=HMRegul.Mess[25]+' '+DV.Symbole ;
      GR.Cells[0,9]:=HMRegul.Mess[26]+' '+DV.Symbole ;
      GR.Cells[0,11]:=HMRegul.Mess[27]+' '+DV.Symbole ;
      END ;
   END ;
HMTrad.ResizeGridColumns(GR) ;
END ;

procedure TFLetRegul.RechercheEcritures ;
Var StXP,StXN,StXP2,StXN2,StXP3,StXN3 : String ;
    Ind                   : integer ;
    lStSQL : string ;
    lQR : TQuery ;
    i : integer ;
BEGIN
GR.ClearSelected ;
if (Regul=pglRegul) then Ind:=1 else if Regul=pglEcart then Ind:=0 else Ind:=3 ;
lStSQL := SelectQDL(Ind)+' Where ' ;
lStSQL := lStSQL + LWhereBase(Regul=pglRegul,True,False,False) ;
//QR.SQL.Add(SelectQDL(Ind)+' Where ') ;
//QR.SQL.Add(LWhereBase(Regul=pglRegul,True,False,False)) ;
if Regul=pglRegul then lStSQL := lStSQL +' AND E_ETATLETTRAGE="PL"' else lStSQL := lStSQL +' AND E_ETATLETTRAGE="TL"' ;
if (EstSpecif('51188')) And (Regul=pglConvert) And (Not (ctxPCL in V_PGI.PGIContexte)) then
  BEGIN
  lStSQL := lStSQL + ' AND E_DATEPAQUETMIN<"'+UsDateTime(EncodeDate(2002,01,01))+'"' ;
  END ;
lStSQL := lStSQL + LWhereComptes(Self) + WhereCrit ;
If XX_WHERESEL.Text<>'' Then lStSQL := lStSQL + ' AND '+XX_WHERESEL.Text ;
lStSQL := lStSQL + ' Group by E_GENERAL, E_AUXILIAIRE, E_DEVISE, E_LETTRAGE' ;
if Regul=pglRegul then
   BEGIN
   if E_DEVISE.Value=V_PGI.DevisePivot then
      BEGIN
         lStSQL := lStSQL + ' Having (SUM(E_DEBIT-E_CREDIT)>0 AND SUM(E_DEBIT-E_CREDIT)<'+StrFPoint(MAXDEBIT.Value)+')' ;
         lStSQL := lStSQL + ' OR (SUM(E_DEBIT-E_CREDIT)<0 AND SUM(E_DEBIT-E_CREDIT)>'+StrFPoint(-MAXCREDIT.Value)+')' ;
      END else
      BEGIN
      lStSQL := lStSQL + ' Having (SUM(E_DEBITDEV-E_CREDITDEV)>0 AND SUM(E_DEBITDEV-E_CREDITDEV)<'+StrFPoint(MAXDEBIT.Value)+')' ;
      lStSQL := lStSQL + ' OR (SUM(E_DEBITDEV-E_CREDITDEV)<0 AND SUM(E_DEBITDEV-E_CREDITDEV)>'+StrFPoint(-MAXCREDIT.Value)+')' ;
      END ;
   END else if Regul in [pglConvert,pglInverse] then
   BEGIN
   StXP:=StrFPoint(9*Resolution(V_PGI.OkDecV+1))  ; StXN:=StrFPoint(-9*Resolution(V_PGI.OkDecV+1)) ;
   StXP2:=StrFPoint(9*Resolution(V_PGI.OkDecE+1)) ; StXN2:=StrFPoint(-9*Resolution(V_PGI.OkDecE+1)) ;
   StXP3:=StrFPoint(9*Resolution(DEV.Decimale+1)) ; StXN3:=StrFPoint(-9*Resolution(DEV.Decimale+1)) ;

   {if EstMonnaieIN(DEV.Code) then
      BEGIN
      lStSQL := lStSQL +' Having ((' ;
   //   QR.SQL.Add('Sum(E_COUVERTUREEURO*'+SoldeSurDC+') Not Between '+StXN2+' AND '+StXP2) ;
   //   QR.SQL.Add(') OR (') ;
      lStSQL := lStSQL + ' Sum(E_COUVERTURE*'+SoldeSurDC+') Not Between '+StXN+' AND '+StXP ;
      lStSQL := lStSQL + ')) AND (' ;
      lStSQL := lStSQL + ' Sum(E_COUVERTUREDEV*'+SoldeSurDC+') Between '+StXN3+' AND '+StXP3 ;
      lStSQL := lStSQL + ')' ;
      END else}
      BEGIN
      if Regul=pglConvert then
         BEGIN
         if VH^.TenueEuro then
            BEGIN
            lStSQL := lStSQL + ' Having (' ;
        //    QR.SQL.Add('Sum(E_COUVERTUREEURO*'+SoldeSurDC+') Not Between '+StXN2+' AND '+StXP2) ;
       //     QR.SQL.Add(') AND (') ;
            lStSQL := lStSQL + ' Sum(E_COUVERTURE*'+SoldeSurDC+') Between '+StXN+' AND '+StXP ;
            lStSQL := lStSQL +')' ;
            END else
            BEGIN
            lStSQL := lStSQL +' Having (' ;
            lStSQL := lStSQL +' Sum(E_COUVERTURE*'+SoldeSurDC+') Not Between '+StXN2+' AND '+StXP2 ;
         //   QR.SQL.Add(') AND (') ;
         //   QR.SQL.Add('Sum(E_COUVERTUREEURO*'+SoldeSurDC+') Between '+StXN+' AND '+StXP) ;
            lStSQL := lStSQL +')' ;
            END ;
         END else if Regul=pglInverse then
         BEGIN
            lStSQL := lStSQL +' Having (' ;
            lStSQL := lStSQL +' Sum(E_COUVERTURE*'+SoldeSurDC+') Not Between '+StXN+' AND '+StXP ;
//            QR.SQL.Add(') AND (') ;
//            QR.SQL.Add('Sum(E_COUVERTUREEURO*'+SoldeSurDC+') Between '+StXN2+' AND '+StXP2) ;
            lStSQL := lStSQL +')' ;
         END ;
      END ;
   END else {écart change}
   BEGIN
      StXP:=StrFPoint(9*Resolution(V_PGI.OkDecV+1)) ; StXN:=StrFPoint(-9*Resolution(V_PGI.OkDecV+1)) ;
      lStSQL := lStSQL +' Having Sum(E_COUVERTURE*'+SoldeSurDC+') Not Between '+StXN+' AND '+StXP ;

   END ;
 lQR:=OpenSQL(lStSQL,true) ;
 FTOB.clearDetail ;
 FTOB.LoadDetailDB('','','',lQR,true) ;
 FTOB.PutGridDetail(GR,false,false,'',true);
 for i:=0 to GR.ColCount-1 do BEGIN
    if i in [3,4] then GR.ColColors[i] := clNavy;
    if i in [5,6] then GR.ColColors[i] := clMaroon;
    if i in [7]   then GR.ColColors[i] := clGreen;
  END;
 TripoteTitres ;
 Ferme(lQR) ;
END ;

{============================== AFFICHAGES ========================================}
procedure TFLetRegul.FormShow(Sender: TObject);
begin
Pages.ActivePage:=Princ ;
InitTablesLibresTiers(PLibres) ;
GeneCharge:=True ;
if Regul in [pglConvert,pglInverse] then begin ExpliEuro.Caption:=HMRegul.Mess[20] ; ExpliEuro.Visible:=False ; end ;
FillChar(DEV,Sizeof(DEV),#0) ; FillChar(JalReG,Sizeof(JalReG),#0) ;
if (Regul=pglRegul) then
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
   END else {Convert,Inverse}
   BEGIN
   Caption:=HMRegul.Mess[18] ; Journal.DataType:='ttJournal' ; Journal.Enabled:=False ;
   H_JOURNALC.Visible:=True ; //Journal.Value:=VH^.JalEcartEuro ;
   MaxDebit.Visible:=False  ; H_SD.Visible:=False ;
   MaxCredit.Visible:=False ; H_SC.Visible:=False ;
   BValider.Hint:=HMRegul.Mess[21] ; HelpContext:=7520500 ;
   RefRegul.Text:=HMRegul.Mess[30] ; LibRegul.Text:=HMRegul.Mess[30] ;
   ExpliEuro.Visible:=False ;
   END ;
if (Regul=pglRegul) then
   BEGIN
   E_DEVISE.DataType:='ttDevise' ;
   E_DEVISE.Value:=V_PGI.DevisePivot ;
   GR.ListeParam:='REGULLETEURO' ;
   END else if Regul=pglEcart then
   BEGIN
   E_DEVISE.DataType:='ttDeviseEtat' ; E_DEVISE.Value:='' ;
   if E_DEVISE.Items.Count>0 then BEGIN E_DEVISE.ItemIndex:=0 ; E_DEVISEChange(Nil) ; END ;
   GR.ListeParam:='REGULLETEURO' ;
   END else {Convert,Inverse}
   BEGIN
   E_DEVISE.DataType:='ttDevise' ;
   E_DEVISE.Value:=V_PGI.DevisePivot ;
   GR.ListeParam:='ECARTCONVERTEURO' ;
   If Regul=pglInverse Then
     BEGIN
    // If VH^.TenueEuro Then QR.Titres:='Général;Auxiliaire;Intitulé auxiliaire;Nb D;Débit Euro;Débit Frcs;'+
    //                                  'Nb C;Crédit Euro;Crédit Frcs;Solde Euro;Solde Frcs;Let;Dev;LD;LE;' ;
     END ;
   END ;
DateGeneration.Text:=DateToStr(V_PGI.DateEntree) ; DateGene:=V_PGI.DateEntree ;
//SG6 12/01/05 Gestion Filtre V6 FQ 15255
ObjFiltre.Charger;
InitCriteres;
GeneCharge:=False ;
UpdateCaption(Self) ;
Tripotetitres ;
If Auto Then
  BEGIN
  BChercherClick(Nil) ;
  if GR.Cells[1,4]<>'' then BSelectAllClick(Nil) ;
  END ;
HMTrad.ResizeGridColumns(GR) ;
if FBoRegulAnc then
 begin
  BChercherClick(nil) ;
  if GR.Cells[1,1]<>'' then 
   begin
    bSelectAllClick(nil) ;
    BValiderClick(nil) ;
   end ;
  PostMessage(self.Handle, WM_CLOSE, 0, 0);
 end ;
end;


procedure TFLetRegul.FormClose(Sender: TObject; var Action: TCloseAction);
Var Ind : integer ;
begin
  //SG6 12/01/05 Gestion Filtres V6 FQ 15255
  FreeAndNil(ObjFiltre);
  // FQ 12785 (SBO) suite au passage full euro, plus de visu de pièces générées si on est en ecart de change
  if Regul<>pglEcart then
    if not FBoRegulAnc and (TPieces.Count>0) then
      BEGIN
      if Regul=pglRegul then Ind:=3 else if Regul=pglEcart then Ind:=4 else Ind:=14 ;
      if HMRegul.Execute(13,caption,'')=mrYes then VisuPiecesGenere(TPieces,EcrGen,Ind) ;
      END ;
  // Fin FQ 12785
VideListe(TPieces) ; TPieces.Free ; FTOB.Free ;
if Parent is THPanel then
   BEGIN
   If (Not MonProfilOk(Qui)) And (Not Auto) Then _Bloqueur('nrBatch',False) ;
   Action:=caFree ;
   END ;
end;

{============================= TRAITEMENT DE REGUL ======================================}
procedure TFLetRegul.RemplirLeM ( Var M : RMVT ; RL : RLETTR ) ;
Var Facturier : String3 ;
BEGIN
FillChar(M,Sizeof(M),#0) ;
M.Jal:=JOURNAL.Value ; Facturier:=JalReG.Facturier ;
M.ModeSaisieJal:=JalReG.ModeSaisie ; 
//SetIncNum(EcrGen,Facturier,M.Num) ;
M.Axe:='' ; M.Etabl:=VH^.EtablisDefaut ; 
if ((Regul=pglRegul) and (RL.LettrageDevise)) or (Regul in [pglConvert,pglInverse])
    then BEGIN M.CodeD:=DEV.Code ; M.TauxD:=DEV.Taux ; END
    else BEGIN M.CodeD:=V_PGI.DevisePivot ; M.TauxD:=1 ; END ;
M.DateTaux:=V_PGI.DateEntree ; M.Valide:=False ; M.Simul:='N' ;
if Regul in [pglRegul,pglConvert,pglInverse] then M.Nature:='OD' else M.Nature:='ECC' ;
M.DateC:=DateGene ; M.Exo:=QuelExo(DateToStr(M.DateC)) ; M.General:='' ;
SetIncNum(EcrGen,Facturier,M.Num,M.DateC) ;
END ;

procedure TFLetRegul.CouvreRegul ( O : TOBM ; RL : RLETTR ) ;
Var TotD,TotF(*,Taux*) : double ;
BEGIN
TotF:=O.GetMvt('E_DEBIT')+O.GetMvt('E_CREDIT') ;
//TotE:=O.GetMvt('E_DEBITEURO')+O.GetMvt('E_CREDITEURO') ;
TotD:=O.GetMvt('E_DEBITDEV')+O.GetMvt('E_CREDITDEV') ;
Case Regul of
   pglRegul : BEGIN
              if RL.LettrageDevise then
                 BEGIN
                 O.PutMvt('E_COUVERTUREDEV',TotD) ; //Taux:=O.GetMvt('E_TAUXDEV') ;
                 

                 END else
                 BEGIN
                    O.PutMvt('E_COUVERTURE',TotF) ; ConvertCouverture(O,tsmPivot) ;
                 END ;
              END ;
   pglEcart : BEGIN
              {Ecart de change couvre tout}
              O.PutMvt('E_COUVERTUREDEV',TotD) ;
              O.PutMvt('E_COUVERTURE',TotF) ;

              END ;
 pglConvert : BEGIN
              {Ecart de conversion couvre tout}
              O.PutMvt('E_COUVERTUREDEV',TotD) ;
              O.PutMvt('E_COUVERTURE',TotF) ;

              END ;
 pglInverse : BEGIN
              {Ecart de conversion couvre tout}
              O.PutMvt('E_COUVERTUREDEV',TotD) ;
              O.PutMvt('E_COUVERTURE',TotF) ;

              END ;
   END ;
 END ;

{***********A.G.L.***********************************************
Auteur  ...... : Laurent gendreau
Créé le ...... : 18/06/2002
Modifié le ... :   /  /
Description .. : Suppression de warning
Mots clefs ... :
*****************************************************************}
procedure TFLetRegul.GenereRegul (vTOB : TOB ) ;
Var RL : RLETTR ;
    Q  : TQuery ;
    X,LimiteP,LimiteE,Montant,MenOppose : Double ;
    St,Cpte,StInd : String ;
    O,OOO     : TOBM ;
    Premier   : boolean ;
    M         : RMVT ;
    XX        : T_ECARTCHANG ;
    DMin,DMax : TDateTime ;
    ForceSaisieEuro : String ;
BEGIN
if vTOB = nil then exit ;
ForceSaisieEuro:='' ;
if ((Regul<>pglConvert) and (Regul<>pglInverse)) then
   BEGIN
   Montant:=Arrondi(vTOB.GetValue('SOLDE'),V_PGI.OkDecV) ;
   MenOppose:=0 ;
   END else
   BEGIN
   Montant:=Arrondi(vTOB.GetValue('COLUMN'),V_PGI.OkDecV) ;
   MenOppose:=Arrondi(vTOB.GetValue('COLUMN10'),V_PGI.OkDecE) ;
   END ;
DMax:=0 ; DMin:=0 ;
LimiteP:=Resolution(V_PGI.OkDecV) ; LimiteE:=Resolution(V_PGI.OkDecE) ;
if ((Regul<>pglConvert) and (Regul<>pglInverse)) then
   BEGIN
   if Abs(Montant)<LimiteP then Exit ;
   END else
   BEGIN
   if ((Abs(Montant)<LimiteP) and (Abs(MenOppose)<LimiteE)) then Exit ;
   END ;
Premier:=True ; FillChar(RL,Sizeof(RL),#0) ;
RL.General:=vTOB.GetValue('E_GENERAL') ;
RL.Auxiliaire:=vTOB.GetValue('E_AUXILIAIRE') ; RL.Appel:=tlMenu ; RL.GL:=Nil ;
RL.CodeLettre:=vTOB.GetValue('E_LETTRAGE') ;
RL.DeviseMvt:=vTOB.GetValue('E_DEVISE') ;
RL.LettrageDevise:=(RL.DeviseMvt<>V_PGI.DevisePivot) ;
if ((Regul=pglRegul) and (RL.LettrageDevise)) then
   BEGIN
   {Récupérer l'écart en devise}
   Q:=OpenSQL('Select SUM(E_DEBITDEV-E_CREDITDEV) S from Ecriture Where '+WhereDetail(RL),True) ;
   if Not Q.EOF then Montant:=Q.FindField('N').AsFloat ;
   Ferme(Q) ;
   END ;
if Regul=pglEcart then
   BEGIN
   {Récupérer l'écart opposé aussi}
      Q:=OpenSQL('Select SUM(E_DEBIT-E_CREDIT) N from Ecriture Where '+WhereDetail(RL),True) ;
      if Not Q.EOF then MenOppose:=Arrondi(Q.FindField('N').AsFloat,V_PGI.OkDecE) ;
      Ferme(Q) ;
//      END ;
   END ;
if ((Regul<>pglConvert) and (Regul<>pglInverse))
   then BEGIN if Montant=0 then Exit ; END
   else BEGIN if ((Montant=0) and (MenOppose=0)) then Exit ; END ;
if Regul in [pglConvert,pglInverse] then
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
         if Regul=pglEcart then St[2]:='X' else
            if Regul=pglConvert then St[2]:='#' else St[2]:='&' ;
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
      if ((Regul<>pglConvert) and (Regul<>pglInverse)) then
         BEGIN
         if Montant>0 then Cpte:=CPTDEBIT.Text else Cpte:=CptCredit.Text ;
         END else {Convert,Inverse}
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
         pglInverse : BEGIN XX.Montant1:=Montant ; XX.Montant2:=MenOppose ; ForceSaisieEuro:='-' ; END ;
         END ;
      OOO:=CreerPartieDoubleLett(M,RL,XX,O,ForceSaisieEuro) ;
      TPieces.Add(OOO) ;
      END ;
   Q.Next ; Premier:=False ; O.Free ;
   END ;
Ferme(Q) ;
END ;

procedure TFLetRegul.ValideLaRegul ;
Var i : integer ;
lTOB : TOB ;
BEGIN
InitMove(GR.NbSelected,'') ;
 if ((GR.AllSelected) or (GR.nbSelected <> 0)) then
        begin
            if (GR.AllSelected) then
            begin
                for i := 1 to GR.RowCount-1 do
                begin
                    GR.Row := i; lTOB:=GetO(GR,i);
                    MoveCur(FALSE) ;
                    if V_PGI.IoError=oeOK then GenereRegul(lTOB) else exit ;
                end;
            end
            else // Traitement <> si AllSelected car sinon ça merde, Si AllSelected alors NbSelected est faux, il vaut 0
            begin
                for i := 0 to GR.nbSelected - 1 do
                begin
                    GR.GotoLeBookMark(i); //GR.Row := i;
                    MoveCur(FALSE) ;
                    lTOB:=GetO(GR,GR.Row);
                    if V_PGI.IoError=oeOK then GenereRegul(lTOB) else exit ;
                end;
            end;
        end; //if
FiniMove ;
END ;

procedure TFLetRegul.BValiderClick(Sender: TObject);
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
if not FBoRegulAnc and  ( HMRegul.Execute(Num,caption,'')<>mrYes ) then Exit ;
if Transactions(ValideLaRegul,1)<>oeOK then MessageAlerte(HMRegul.Mess[15]+#10#13+#10#13+V_PGI.LastSQLError) ;
if not FBoRegulAnc then BChercherClick(Nil) ;
end;


procedure TFLetRegul.AlimXX_WHERESEL ;
Var QS : TQuery ;
    Gene,Aux,Lett,St,St1 : String ;
BEGIN
St:='' ; St1:='' ;
QS:=OpenSQL('SELECT * FROM CPMPTEMPOR WHERE CTT_USER="'+V_PGI.User+'" AND CTT_TYPETRAITEMENT="REG" ',TRUE) ;
While Not QS.Eof Do
  BEGIN
  Gene:=QS.FindField('CTT_GENERAL').AsString ;
  Aux:=QS.FindField('CTT_AUXILIAIRE').AsString ;
  Lett:=QS.FindField('CTT_ZONELIBRE1').AsString ;
  St1:='(E_GENERAL="'+Gene+'" AND E_AUXILIAIRE="'+Aux+'" AND E_LETTRAGE="'+Lett+'") ' ;
  If St='' Then St:=St1 Else St:=St+' OR '+St1 ;
  QS.Next ;
  END ;
Ferme(QS) ;
If St='' Then Exit ;
St:='E_JOURNAL<>"'+W_W+'" AND ('+St+')' ;
XX_WHERESEL.Text:=St ;
END ;

procedure TFLetRegul.InitCriteres ;
Var DD : TDateTime ;
BEGIN
If Auto Then
  BEGIN
  BSelectAll.Visible:=TRUE ;
  If XX_WHERESEL.TExt='' Then
    BEGIN
    AlimXX_WHERESEL ;
    E_GENERAL.Text:='' ; E_GENERAL_.Text:='' ;
    E_AUXILIAIRE.Text:='' ; E_AUXILIAIRE_.Text:='' ;
    E_DATECOMPTABLE.Text:=StDate1900 ; E_DATECOMPTABLE_.Text:=StDate2099 ;
  //  E_DEVISE.Value:=V_PGI.DevisePivot ; T_NATUREAUXI.ItemIndex:=0 ; ModeOppose.Checked:=FALSE ;
    T_TABLE0.Text:='' ; T_TABLE1.Text:='' ; T_TABLE2.Text:='' ; T_TABLE3.Text:='' ; T_TABLE4.Text:='' ;
    T_TABLE5.Text:='' ; T_TABLE6.Text:='' ; T_TABLE7.Text:='' ; T_TABLE8.Text:='' ; T_TABLE9.Text:='' ;
    END ;
  END ELse
  BEGIN
  if VH^.Precedent.Code<>'' then E_DATECOMPTABLE.Text:=DateToStr(VH^.Precedent.Deb)
                            else E_DATECOMPTABLE.Text:=DateToStr(VH^.Encours.Deb) ;
  if FBoRegulAnc then
   begin
    Journal.Value := GetParamSocSecur('SO_LETCHOIXJAL','') ;
    JOURNALChange(nil) ;
    E_DATECOMPTABLE_.Text:=StDate2099 ;
   end
    else
     E_DATECOMPTABLE_.Text:=DateToStr(V_PGI.DateEntree) ;
  if ((VH^.ExoV8.Code<>'') and (StrToDate(E_DATECOMPTABLE.Text)<VH^.ExoV8.Deb)) then E_DATECOMPTABLE.Text:=DateToStr(VH^.ExoV8.Deb) ;
  END ;
DD:=StrToDate(DATEGENERATION.Text) ;
if ModeRevisionActive(DD) then
   BEGIN
   DATEGENERATION.Text:=DateToStr(V_PGI.DateEntree) ;
   DateGene:=V_PGI.DateEntree ;
   END ;
END ;

procedure TFLetRegul.WMGetMinMaxInfo(var MSG: Tmessage);
begin
with PMinMaxInfo(MSG.lparam)^.ptMinTrackSize do begin X := WMinX ; Y := WMinY ; end;
end;

procedure TFLetRegul.BAideClick(Sender: TObject);
begin
CallHelpTopic(Self) ;
end;

procedure TFLetRegul.bSelectAllClick(Sender: TObject);
begin
 GR.AllSelected := not  GR.AllSelected ; 
end;

procedure TFLetRegul.FormKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
if ((Shift=[]) and (Key=VK_F9)) then BEGIN Key:=0 ; BChercherClick(Nil) ; END ;
end;

end.
