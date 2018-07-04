unit QRTvaAcc;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  QR, HSysMenu, Menus, HMsgbox, HQuickrp, DB,
  {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}
  StdCtrls, Buttons,
  Hctrls, HQry, ExtCtrls, Mask, Hcompte, ComCtrls, Ent1, Hent1,CritEdt,UtilEdt,QRRupt,
  CpteUtil,
(*{$IFNDEF ESP}
  TVA,
{$ELSE}
{$ENDIF ESP} //XMG 19/01/04*)
  SaisComm,SaisUtil, HTB97, HPanel, UiUtil ;

// A mettre soit en ligne de menu
// soit en liste (MUL) préparatoire
// et positionner sur l'onglet de mise en page

procedure EditionTvaAcc(Crit : TCritEdt) ;
procedure EditionTvaAccZoom(O : TOBM) ;

type
  TFQRTvaAcc = class(TFQR)
    QRLabel1: TQRLabel;
    TE_DATECOMPTABLE: TQRLabel;
    TE_JOURNAL: TQRLabel;
    TE_PIECE: TQRLabel;
    TE_REFINTERNE: TQRLabel;
    TE_LIBELLE: TQRLabel;
    TE_TTC: TQRLabel;
    TE_COUV: TQRLabel;
    MsgBox: THMsgBox;
    REPORT1TTC: TQRLabel;
    REPORT1COUV: TQRLabel;
    REPORT2TTC: TQRLabel;
    REPORT2COUV: TQRLabel;
    TITRE1REP: TQRLabel;
    TITRE2REP: TQRLabel;
    TE_MONTANT: TQRLabel;
    REPORT1MONTANT: TQRLabel;
    TE_MODEPAIE: TQRLabel;
    TE_LETTRAGE: TQRLabel;
    FSautPage: TCheckBox;
    TypeTVA: TRadioGroup;
    TLE_AUXILIAIRE: TQRLabel;
    E_DATECOMPTABLE: TQRDBText;
    E_JOURNAL: TQRDBText;
    EE_PIECELIGECH: TQRLabel;
    E_LETTRAGE: TQRDBText;
    E_REFINTERNE: TQRDBText;
    E_LIBELLE: TQRDBText;
    E_AUXILIAIRE: TQRDBText;
    LE_PAYE: TQRLabel;
    BFinFact: TQRBand;
    BFinAux: TQRBand;
    TotalTiers: TQRLabel;
    LE_Base1: TQRLabel;
    LE_Base2: TQRLabel;
    LE_Base3: TQRLabel;
    LE_Base4: TQRLabel;
    LE_Base5: TQRLabel;
    LE_BaseFact1: TQRLabel;
    LE_BaseFact2: TQRLabel;
    LE_BaseFact3: TQRLabel;
    LE_BaseFact4: TQRLabel;
    LE_BaseFact5: TQRLabel;
    LE_BaseAcc1: TQRLabel;
    LE_BaseAcc2: TQRLabel;
    LE_BaseAcc3: TQRLabel;
    LE_BaseAcc4: TQRLabel;
    LE_BaseAcc5: TQRLabel;
    LE_BaseCalc1: TQRLabel;
    LE_BaseCalc2: TQRLabel;
    LE_BaseCalc3: TQRLabel;
    LE_BaseCalc4: TQRLabel;
    LE_BaseCalc5: TQRLabel;
    LE_Ecart1: TQRLabel;
    LE_Ecart2: TQRLabel;
    LE_Ecart3: TQRLabel;
    LE_Ecart4: TQRLabel;
    LE_Ecart5: TQRLabel;
    Q2: TQuery;
    QRLabel16: TQRLabel;
    QRLabel19: TQRLabel;
    QRLabel20: TQRLabel;
    QRLabel21: TQRLabel;
    QRLabel22: TQRLabel;
    QRLabel17: TQRLabel;
    QRLabel23: TQRLabel;
    QRLabel25: TQRLabel;
    CouvAcc: TQRLabel;
    CouvFact: TQRLabel;
    RatioCalc: TQRLabel;
    QRBand1: TQRBand;
    QRLigne2: TQRLigne;
    QRLigne1: TQRLigne;
    Trait4: TQRLigne;
    Trait3: TQRLigne;
    Trait2: TQRLigne;
    Trait1: TQRLigne;
    Trait0: TQRLigne;
    Ligne1: TQRLigne;
    LE_COUVERTURE: TQRLabel;
    TE_RIEN: TQRLabel;
    REPORT2MONTANT: TQRLabel;
    LE_MODEPAIE: TQRLabel;
    HM: THMsgBox;
    FDetail: TCheckBox;
    LE_Base1Aux: TQRLabel;
    LE_Base2Aux: TQRLabel;
    LE_Base3Aux: TQRLabel;
    LE_Base4Aux: TQRLabel;
    LE_Base5Aux: TQRLabel;
    LE_BaseFact1Aux: TQRLabel;
    LE_BaseFact2Aux: TQRLabel;
    LE_BaseFact3Aux: TQRLabel;
    LE_BaseFact4Aux: TQRLabel;
    LE_BaseFact5Aux: TQRLabel;
    LE_BaseAcc1Aux: TQRLabel;
    LE_BaseAcc2Aux: TQRLabel;
    LE_BaseAcc3Aux: TQRLabel;
    LE_BaseAcc4Aux: TQRLabel;
    LE_BaseAcc5Aux: TQRLabel;
    LE_BaseCalc1Aux: TQRLabel;
    LE_BaseCalc2Aux: TQRLabel;
    LE_BaseCalc3Aux: TQRLabel;
    LE_BaseCalc4Aux: TQRLabel;
    LE_BaseCalc5Aux: TQRLabel;
    LE_Ecart1Aux: TQRLabel;
    LE_Ecart2Aux: TQRLabel;
    LE_Ecart3Aux: TQRLabel;
    LE_Ecart4Aux: TQRLabel;
    LE_Ecart5Aux: TQRLabel;
    QRLabel44: TQRLabel;
    QRLabel45: TQRLabel;
    QRLabel46: TQRLabel;
    QRLabel47: TQRLabel;
    QRLabel48: TQRLabel;
    QRLabel3: TQRLabel;
    QRLabel6: TQRLabel;
    CouvAccAux: TQRLabel;
    QRLabel9: TQRLabel;
    CouvFactAux: TQRLabel;
    RatioCalcAux: TQRLabel;
    BRecapG: TQRBand;
    QRLabel13: TQRLabel;
    LE_Base1G: TQRLabel;
    LE_Base2G: TQRLabel;
    LE_Base3G: TQRLabel;
    LE_Base4G: TQRLabel;
    LE_Base5G: TQRLabel;
    LE_BaseFact1G: TQRLabel;
    LE_BaseFact2G: TQRLabel;
    LE_BaseFact3G: TQRLabel;
    LE_BaseFact4G: TQRLabel;
    LE_BaseFact5G: TQRLabel;
    LE_BaseAcc1G: TQRLabel;
    LE_BaseAcc2G: TQRLabel;
    LE_BaseAcc3G: TQRLabel;
    LE_BaseAcc4G: TQRLabel;
    LE_BaseAcc5G: TQRLabel;
    LE_BaseCalc1G: TQRLabel;
    LE_BaseCalc2G: TQRLabel;
    LE_BaseCalc3G: TQRLabel;
    LE_BaseCalc4G: TQRLabel;
    LE_BaseCalc5G: TQRLabel;
    LE_Ecart1G: TQRLabel;
    LE_Ecart2G: TQRLabel;
    LE_Ecart3G: TQRLabel;
    LE_Ecart4G: TQRLabel;
    LE_Ecart5G: TQRLabel;
    QRLabel56: TQRLabel;
    QRLabel57: TQRLabel;
    QRLabel58: TQRLabel;
    QRLabel59: TQRLabel;
    QRLabel60: TQRLabel;
    QRLabel61: TQRLabel;
    QRLabel62: TQRLabel;
    CouvAccG: TQRLabel;
    QRLabel64: TQRLabel;
    CouvFactG: TQRLabel;
    RatioCalcG: TQRLabel;
    FRecapG: TCheckBox;
    FRecapCpt: TCheckBox;
    RTva: TQRLabel;
    RTypeTva: TQRLabel;
    QRDLMvtFact: TQRDetailLink;
    S2: TDataSource;
    QRDLMulti: TQRDetailLink;
    QRDLRecapG: TQRDetailLink;
    procedure FormShow(Sender: TObject);
    procedure QAfterOpen(DataSet: TDataSet);
    procedure BDetailBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure BFinEtatBeforePrint(var PrintBand: Boolean;  var Quoi: string);
    procedure Q2BeforeOpen(DataSet: TDataSet);
    procedure BFinFactBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure Q2AfterOpen(DataSet: TDataSet);
    procedure BFinAuxBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure QRDLMultiNeedData(var MoreData: Boolean);
    procedure QRDLRecapGNeedData(var MoreData: Boolean);
    procedure BRecapGAfterPrint(BandPrinted: Boolean);
    procedure EntetePageBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure BRecapGBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure GenereSQL ; Override ;
    procedure FinirPrint ; Override ;
    procedure RecupCritEdt ; Override ;
    procedure RenseigneCritere ; Override ;
    function  CritOk : Boolean ; Override ;
    procedure InitDivers ; Override ;
    procedure ChoixEdition ; Override ;
  private
    Qr1Debit,Qr1Modepaie,Qr1Credit,Qr2Debit,Qr2Credit,Qr1Couv,Qr2Couv : TFloatField ;
    Qr2Regime,Qr1E_AUXILIAIRE,QR1T_LIBELLE : TStringField ;
    TotBaseFact : Array[0..12] of Double ;
    TabTva : Array[1..5] of Double ;
    O : TOBM ;
    (* Tableau des couvertures et bases HT :
       0 : règlement acompte
       1 : couverture acompte
       2 : couvertures factures
       3 : bases facturées
       4 : bases acomptes
    *)
    LTiers,LRecapG : TStringList ;
    Function  QuoiMvt(i : integer) : String ;
    Function  QuoiAux(i : Integer) : String ;
    Procedure TvaAccZoom(Quoi : String) ;
    procedure GenereSQLSub ;
    procedure AfficheBases(Suffixe : String ; Tot : Array of Double) ;
  public
    { Déclarations publiques }
  end;

implementation

{$R *.DFM}

Function TFQRTvaAcc.QuoiAux(i : Integer) : String ;
BEGIN

Case i Of
  1 : Result:=QR1E_AUXILIAIRE.AsString+' '+'#'+QR1T_LIBELLE.AsString+'@'+'2' ;
  END ;
END ;

Function TFQRTvaAcc.QuoiMvt( i : integer) : String ;
var St : String ;
BEGIN
case i of
  // zoom acompte
  1 : St:=Qr1E_AUXILIAIRE.AsString+' '+Qr1T_LIBELLE.AsString+' '+{Le_Solde.Caption+}        '#'+Q.FindField('E_JOURNAL').AsString+' N° '+IntToStr(Q.FindField('E_NUMEROPIECE').AsInteger)+' '+
      DateToStr(Q.FindField('E_DateComptable').AsDateTime)+' . . . '+
      PrintSolde(Qr1DEBIT.AsFloat,Qr1CREDIT.AsFloat,CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole)+
      '@'+'5;'+Q.FindField('E_JOURNAL').AsString+';'+UsDateTime(Q.FindField('E_DATECOMPTABLE').AsDateTime)+';'+Q.FindField('E_NUMEROPIECE').AsString+';'+Q.FindField('E_EXERCICE').asString+';'+
        IntToStr(Q.FindField('E_NumLigne').AsInteger)+';' ;
  // zoom Lettrage
  2 : St:=Qr1E_AUXILIAIRE.AsString+' '+Qr1T_LIBELLE.AsString+' '+{Le_Solde.Caption+}        '#'+Q.FindField('E_JOURNAL').AsString+' N° '+IntToStr(Q.FindField('E_NUMEROPIECE').AsInteger)+' '+
      DateToStr(Q.FindField('E_DateComptable').AsDAteTime)+' . . . '+
      PrintSolde(Qr1DEBIT.AsFloat,Qr1CREDIT.AsFloat,CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole)+
      '@'+'6;'+Q.FindField('E_JOURNAL').AsString+';'+UsDateTime(Q.FindField('E_DATECOMPTABLE').AsDateTime)+';'+Q.FindField('E_NUMEROPIECE').AsString+';'+Q.FindField('E_EXERCICE').asString+';'+
        IntToStr(Q.FindField('E_NumLigne').AsInteger)+';' ;
  END ;
Result:=St ;
END ;

Procedure TFQRTvaAcc.TvaAccZoom(Quoi : String) ;
Var Lp,i: Integer ;
    St : String ;
BEGIN
Lp:=Pos('@',Quoi) ; If Lp=0 Then Exit ;
St:=Copy(Quoi,Lp+1,1) ;
If St='!' Then i:=100 Else i:=StrToInt(St) ;
If (i=5) or (i=6) or (i=100) Then
   BEGIN
   Quoi:=Copy(Quoi,Lp+3,Length(Quoi)-lp-2) ;
   If QRP.QrPrinter.FSynShiftDblClick Then i:=6 ;
   If QRP.QRPrinter.FSynCtrlDblClick Then i:=11 ;
   END ;
ZoomEdt(i,Quoi) ;
END ;

procedure EditionTvaAcc(Crit : TCritEdt) ;
var QR : TFQRTvaAcc ;
    Edition : TEdition ;
    PP : THPanel ;
BEGIN
PP:=FindInsidePanel ;
QR:= TFQRTvaAcc.Create(Application) ;
Edition.Etat:=etTvaAcc ;
QR.O:=nil ;
QR.QRP.QRPrinter.OnSynZoom:=QR.TvaAccZoom ;
QR.CritEdt:=Crit ;
if Crit.GL.Deductible then QR.TypeTva.ItemIndex:=1 ;
QR.InitType (nbGen,neGL,msGenEcr,'QRTvaAcc','',FALSE,FALSE,TRUE,Edition) ;
if PP=Nil then
   BEGIN
   try
    QR.ShowModal ;
    finally
    QR.Free ;
    end ;
   Screen.Cursor:=SyncrDefault ;
   END Else
   BEGIN
   InitInside(QR,PP) ;
   QR.Show ;
   END ;
END;

procedure EditionTvaAccZoom(O : TOBM) ;
var QR : TFQRTvaAcc ;
    Edition : TEdition ;
    Crit : TCritEdt ;
    D1,D2 : TDateTime ;
BEGIN
QR:= TFQRTvaAcc.Create(Application) ;
Edition.Etat:=etTvaAcc ;
try
 QR.O:=O ;
 Fillchar(Crit,SizeOf(Crit),#0) ;
 Crit.NatureEtat:=neGL ;
 InitCritEdt(Crit) ;
 D1:=VH^.Encours.Deb ; D2:=VH^.Encours.Fin ;
 If VH^.Entree.Code=VH^.Suivant.Code Then BEGIN D1:=VH^.Suivant.Deb ; D2:=VH^.Suivant.Fin ; END ;
 Crit.Date1:=D1 ; Crit.Date2:=D2 ;
 Crit.DateDeb:=Crit.Date1 ; Crit.DateFin:=Crit.Date2 ;
 Crit.GL.Deductible:=(O.GetMvt('E_NATUREPIECE')='OF') ;
 QR.CritEdt:=Crit ;
 if Crit.GL.Deductible then QR.TypeTva.ItemIndex:=1 ;
 QR.QRP.QRPrinter.OnSynZoom:=QR.TvaAccZoom ;
 QR.InitType (nbGen,neGL,msGenEcr,'QRTvaAcc','',FALSE,TRUE,TRUE,Edition) ;
 finally
 QR.Free ;
 end ;
SourisNormale ;
END;

procedure TFQRTvaAcc.GenereSQL ;
{ Construction de la requête SQL en fonction du multicritère }
BEGIN
Inherited ;
Q.Close ; Q.SQL.Clear ;
Q.SQL.Add('Select E_GENERAL,E_AUXILIAIRE,E_EXERCICE,E_DATECOMPTABLE,E_NUMEROPIECE,E_NUMLIGNE,') ;
Q.SQL.Add(       'E_REFINTERNE,E_LIBELLE,E_MODEPAIE,E_DEVISE,E_LETTRAGE,E_JOURNAL,E_NATUREPIECE,E_ETATLETTRAGE,') ;
Q.SQL.Add(       'E_CONTREPARTIEGEN,E_ECHEENC1,E_ECHEENC2,E_ECHEENC3,E_ECHEENC4,E_ECHEDEBIT,T_LIBELLE AS LIBAUX,') ;
Case CritEdt.Monnaie of
  0 : BEGIN Q.SQL.Add('E_DEBIT as DEBIT,E_CREDIT as CREDIT, E_COUVERTURE as COUVERTURE') ; END ;
  1 : BEGIN Q.SQL.Add('E_DEBITDEV as DEBIT,E_CREDITDEV as CREDIT, E_COUVERTUREDEV as COUVERTURE') ; END ;
end ;
Q.SQL.Add(' From ECRITURE ') ;
Q.SQL.Add(' Left Join TIERS on E_AUXILIAIRE=T_AUXILIAIRE ') ;
if (O<>nil) then
   BEGIN
   Q.SQL.Add('WHERE E_JOURNAL="'+O.GetMvt('E_JOURNAL')+'"') ;
   Q.SQL.Add(' AND E_DATECOMPTABLE="'+UsDateTime(O.GetMvt('E_DATECOMPTABLE'))+'"') ;
   Q.SQL.Add(' And E_EXERCICE="'+O.GetMvt('E_EXERCICE')+'"') ;
   Q.SQL.Add(' And E_NUMEROPIECE='+IntToStr(O.GetMvt('E_NUMEROPIECE'))) ;
   Q.SQL.Add(' And E_QUALIFPIECE="'+O.GetMvt('E_QUALIFPIECE')+'"') ;
   END else
   BEGIN
   Q.SQL.Add('WHERE E_DATECOMPTABLE>="'+UsDateTime(CritEdt.GL.Date21)+'" And E_DATECOMPTABLE<="'+usdatetime(CritEdt.Date2)+'"') ;
   if CritEdt.Exo.Code<>'' then Q.SQL.Add(' And E_EXERCICE="'+CritEdt.Exo.Code+'" ') ;
   If CritEdt.GL.Deductible Then
                            BEGIN
                            Q.SQL.Add(' AND (T_NATUREAUXI="FOU" Or T_NATUREAUXI="AUC") ')
                            END Else
                            BEGIN
                            Q.SQL.Add(' AND (T_NATUREAUXI="CLI" Or T_NATUREAUXI="AUD") ') ;
                            Q.SQL.Add(SQLCollCliEncTVA('')) ;
                            END ;
   If CritEdt.GL.Deductible Then
      BEGIN
      Q.SQL.Add(' AND (E_NATUREPIECE="OF")') ;
      END Else
      BEGIN
      Q.SQL.Add(' AND (E_NATUREPIECE="OC")') ;
      END ;
   // Acomptes édités puis lettrés avec les factures associées (bases HT non remises à zéro)
   Q.SQL.Add(' AND E_EDITEETATTVA="X"') ;
   Q.SQL.Add(' AND (E_ETATLETTRAGE="TL" OR E_ETATLETTRAGE="PL")') ;
   END ;
Q.SQL.Add(' AND (E_ECHEENC1<>0 OR E_ECHEENC2<>0 OR E_ECHEENC3<>0 OR E_ECHEENC4<>0 OR E_ECHEDEBIT<>0)') ;
{ Construction de la clause Order By de la SQL }
Q.SQL.Add(' Order By E_AUXILIAIRE,E_GENERAL,E_ETATLETTRAGE,E_LETTRAGE,E_DATECOMPTABLE') ;
ChangeSQL(Q) ; //Q.Prepare ;
PrepareSQLODBC(Q) ;
Q.Open ;
GenereSQLSub ;
END ;

procedure TFQRTvaAcc.GenereSQLSub ;
{ Construction de la requête SQL en fonction du multicritère }
BEGIN
Q2.Close ; Q2.SQL.Clear ;
Q2.SQL.Add('Select E_GENERAL,E_AUXILIAIRE,E_EXERCICE,E_DATECOMPTABLE,E_NUMEROPIECE,E_NUMLIGNE,E_REGIMETVA,') ;
Q2.SQL.Add(       'E_REFINTERNE,E_LIBELLE,E_DEVISE,E_LETTRAGE,E_JOURNAL,E_NATUREPIECE,E_CONTREPARTIEGEN,') ;
Q2.SQL.Add(       'E_ECHEENC1,E_ECHEENC2,E_ECHEENC3,E_ECHEENC4,E_ECHEDEBIT,') ;
Case CritEdt.Monnaie of
  0 : BEGIN Q2.SQL.Add('E_DEBIT as DEBIT2,E_CREDIT as CREDIT2, E_COUVERTURE as COUVERTURE2') ; END ;
  1 : BEGIN Q2.SQL.Add('E_DEBITDEV as DEBIT2,E_CREDITDEV as CREDIT2, E_COUVERTUREDEV as COUVERTURE2') ; END ;
end ;
Q2.SQL.Add(' From ECRITURE ') ;
Q2.SQL.Add(' Where E_AUXILIAIRE=:AUX And E_GENERAL=:GEN And E_ETATLETTRAGE=:ETATL And E_LETTRAGE=:LETT') ;
If CritEdt.GL.Deductible Then Q2.SQL.Add(' AND (E_NATUREPIECE="FF" OR E_NATUREPIECE="AF") ')
                         else Q2.SQL.Add(' AND (E_NATUREPIECE="FC" OR E_NATUREPIECE="AC") ') ;
{ Construction de la clause Order By de la SQL }
Q2.SQL.Add(' Order By E_AUXILIAIRE,E_GENERAL,E_ETATLETTRAGE,E_LETTRAGE,E_DATECOMPTABLE') ;
ChangeSQL(Q2) ; //Q2.Prepare ;
PrepareSQLODBC(Q2) ;
Q2.Open ;
END ;

procedure TFQRTvaAcc.FormShow(Sender: TObject);
begin
  inherited;
FEXERCICE.value:=QuelExoDT(CritEdt.Date1) ;
FDateCompta1.Text:=DateToStr(CritEdt.Date1) ; FDateCompta2.Text:=DateToStr(CritEdt.Date2) ;
Pages.ActivePage:=Option ;
TabSup.TabVisible:=False ;
FGroupChoixRupt.Visible:=False ; FGroupChoixRupt.Enabled:=False ;
AvecRevision.Visible:=False ;
FReport.Checked:=False ; FReport.Enabled:=False ;
HelpContext:=7652000 ;
Standards.HelpContext:=7652100 ;
Avances.HelpContext:=7652200 ;
Mise.HelpContext:=7652300 ;
Option.HelpContext:=7652400 ;
end;

procedure TFQRTvaAcc.InitDivers ;
BEGIN
Inherited ;
// saut de page pour récap global avec (détail et/ou récapt tiers)
if (CritEdt.GL.FormatPrint.PrSepCompte[10]) and (CritEdt.GL.FormatPrint.PrSepCompte[2] or CritEdt.GL.FormatPrint.PrSepCompte[8]) then BRecapG.ForceNewPage:=TRUE ; 
if CritEdt.GL.FormatPrint.PrSepCompte[10]
   and not CritEdt.GL.FormatPrint.PrSepCompte[2]
   and not CritEdt.GL.FormatPrint.PrSepCompte[8] then TLE_AUXILIAIRE.Caption:='' else TLE_AUXILIAIRE.Caption:=HM.Mess[8] ;
if CritEdt.GL.FormatPrint.PrSepCompte[2] then
  BEGIN
  TE_DATECOMPTABLE.Caption:=HM.Mess[9] ; TE_JOURNAL.Caption:=HM.Mess[10] ;
  TE_PIECE.Caption:=HM.Mess[11] ; TE_LETTRAGE.Caption:=HM.Mess[12] ; TE_REFINTERNE.Caption:=HM.Mess[13] ;
  TE_LIBELLE.Caption:=HM.Mess[14] ; TE_MODEPAIE.Caption:=HM.Mess[15] ;
  TE_TTC.Caption:=HM.Mess[16] ; TE_COUV.Caption:=HM.Mess[17] ;
  END ;
END ;

procedure TFQRTvaAcc.RenseigneCritere ;
BEGIN
Inherited ;
RCpte.Visible:=False ; TRa.Visible:=False ; RCpte2.Visible:=False ;
RTypeTva.Caption:=Copy(TypeTva.Items[TypeTva.ItemIndex],2,Length(TypeTva.Items[TypeTva.ItemIndex])-1) ;
END ;

procedure TFQRTvaAcc.RecupCritEdt ;
BEGIN
Inherited ;
With CritEdt Do
  BEGIN
  if FSautPage.State=cbGrayed then SautPage:=0 ;
  if FSautPage.State=cbChecked then SautPage:=1 ;
  if FSautPage.State=cbUnChecked then SautPage:=2 ;
  GL.Deductible:=TypeTva.ItemIndex=1 ;
  With GL.FormatPrint Do
    BEGIN
    PrSepCompte[2]:=FDetail.Checked ;
    PrSepCompte[8]:=FRecapCpt.Checked ;
    PrSepCompte[10]:=FRecapG.Checked ;
    END ;
  END ;
END ;

function TFQRTvaAcc.CritOk : Boolean ;
BEGIN
Result:=Inherited CritOK ;
if not (CritEdt.GL.FormatPrint.PrSepCompte[2])
   and not (CritEdt.GL.FormatPrint.PrSepCompte[8])
   and not (CritEdt.GL.FormatPrint.PrSepCompte[10]) then BEGIN NumErreurCrit:=0 ; Result:=False ; END ;
END ;

procedure TFQRTvaAcc.ChoixEdition ;
{ Initialisation des options d'édition }
BEGIN
Inherited ;
ChargeGroup(LTiers,[MsgBox.Mess[7]]) ; ChargeRecap(LRecapG) ;
ChgMaskChamp(Qr1DEBIT , CritEDT.Decimale, CritEDT.AfficheSymbole, CritEDT.Symbole, False) ;
ChgMaskChamp(Qr1CREDIT, CritEDT.Decimale, CritEDT.AfficheSymbole, CritEDT.Symbole, False) ;
ChgMaskChamp(Qr2DEBIT , CritEDT.Decimale, CritEDT.AfficheSymbole, CritEDT.Symbole, False) ;
ChgMaskChamp(Qr2CREDIT, CritEDT.Decimale, CritEDT.AfficheSymbole, CritEDT.Symbole, False) ;
END ;

procedure TFQRTvaAcc.FinirPrint ;
begin
Inherited ;
VideGroup(LTiers) ; VideRecap(LRecapG) ;
end;


procedure TFQRTvaAcc.BFinEtatBeforePrint(var PrintBand: Boolean;  var Quoi: string);
begin
  inherited;
PrintBand:=False ;
BOTTOMREPORT.enabled:=FALSE ;
end;

procedure TFQRTvaAcc.BDetailBeforePrint(var PrintBand: Boolean; var Quoi: string);
var TotPaye : Double ;
begin
  inherited;
if not CritEdt.GL.FormatPrint.PrSepCompte[2] then PrintBand:=False ;
EE_PIECELIGECH.Caption:=Q.FindField('E_NUMEROPIECE').AsString+'/'+Q.FindField('E_NUMLIGNE').AsString ;
If CritEdt.GL.Deductible Then TotPaye:=Arrondi(Qr1Debit.AsFloat-Qr1Credit.AsFloat,CritEdt.Decimale)
                         Else TotPaye:=Arrondi(Qr1Credit.AsFloat-Qr1Debit.AsFloat,CritEdt.Decimale) ;
Fillchar(TotBaseFact,SizeOf(TotBaseFact),#0) ;
TotBaseFact[0]:=TotPaye ; TotBaseFact[1]:=Qr1Couv.AsFloat ;
AddGroup(LTiers, [Qr1E_AUXILIAIRE], [TotBaseFact[0],TotBaseFact[1],0,0,0,0,0,0,0,0,0,0,0]) ;
Quoi:=QuoiMvt(1) ;
if not PrintBand then Exit ;
LE_MODEPAIE.Caption:=' '+RechDom('ttModePaie',Qr1Modepaie.AsString,False) ;
LE_PAYE.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, TotPaye, CritEdt.AfficheSymbole ) ;
LE_COUVERTURE.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, Qr1Couv.AsFloat, CritEdt.AfficheSymbole ) ;
end;

procedure TFQRTvaAcc.AfficheBases(Suffixe : String ; Tot : Array of Double) ;
var Ratio : Double ;
    i : Integer ;
    Pourcent : String ;
BEGIN
Ratio:=1 ;
if Tot[2]<>0 then Ratio:=Tot[1]/Tot[2] ;
for i:=1 to 5 do
  BEGIN
  if i<=4 then
     BEGIN
     Pourcent:=StrS0(100.0*Tva2Taux(QR2Regime.AsString,VH^.NumCodeBase[i],CritEdt.GL.Deductible)) ;
     if (Trim(Pourcent)<>'') then Pourcent:=Trim(Pourcent)+' %' ;
     TQRLabel(FindComponent('LE_Base'+IntToStr(i)+Suffixe)).Caption:=' '+HM.Mess[i-1]+' '+VH^.NumCodeBase[i]+' '+Pourcent ;
     END else
     BEGIN
     Pourcent:='' ;
     TQRLabel(FindComponent('LE_Base'+IntToStr(i)+Suffixe)).Caption:=' '+HM.Mess[i-1] ;
     END ;
  TQRLabel(FindComponent('LE_BaseAcc'+IntToStr(i)+Suffixe)).Caption:=StrS0(Tot[i+2]) ;
  TQRLabel(FindComponent('LE_BaseFact'+IntToStr(i)+Suffixe)).Caption:=StrS0(Tot[i+7]) ;
  TQRLabel(FindComponent('LE_BaseCalc'+IntToStr(i)+Suffixe)).Caption:=StrS0(Tot[i+7]*Ratio) ;
  TQRLabel(FindComponent('LE_Ecart'+IntToStr(i)+Suffixe)).Caption:=StrS0(Valeur(TQRLabel(FindComponent('LE_BaseCalc'+IntToStr(i)+Suffixe)).Caption)-Valeur(TQRLabel(FindComponent('LE_BaseAcc'+IntToStr(i)+Suffixe)).Caption)) ;
  END ;
TQRLabel(FindComponent('CouvAcc'+Suffixe)).Caption:=StrS0(Arrondi(Tot[0],V_PGI.OkDecV))+' /' ;
TQRLabel(FindComponent('CouvFact'+Suffixe)).Caption:=StrS0(Tot[2]) ;
if Ratio<>0 then Pourcent:=' %' else Pourcent:='' ;
TQRLabel(FindComponent('RatioCalc'+Suffixe)).Caption:=StrS0(Arrondi((Ratio*100.0),V_PGI.OkDecV))+Pourcent ;
END ;

procedure TFQRTvaAcc.BFinFactBeforePrint(var PrintBand: Boolean; var Quoi: string);
Var j : integer ;
    XX,Signe,TotFact,TotCouv,Couv : Double ;
begin
  inherited;
if not CritEdt.GL.FormatPrint.PrSepCompte[2] then PrintBand:=False ;
TotFact:=0 ; TotCouv:=0 ; FillChar(TabTva,SizeOf(TabTva),#0) ;
While not Q2.Eof do
    BEGIN
    Signe:=1.0 ;
    if not CritEdt.GL.Deductible then
       BEGIN
       XX:=Qr2Debit.AsFloat-Qr2Credit.AsFloat ;
       if Qr2Debit.AsFloat=0 then Signe:=-1.0 ;
       END else
       BEGIN
       XX:=Qr2Credit.AsFloat-Qr2Debit.AsFloat ;
       if Qr2Credit.AsFloat=0 then Signe:=-1.0 ;
       END ;
    for j:=1 to 4 do TabTva[j]:=TabTva[j]+Q2.FindField('E_ECHEENC'+IntToStr(j)).AsFloat(**Signe*) ;
    TabTva[5]:=TabTva[5]+Q2.FindField('E_ECHEDEBIT').AsFloat ;
    Couv:=Arrondi(Qr2Couv.AsFloat,CritEdt.Decimale) ;
    TotFact:=TotFact+XX ;
    TotCouv:=TotCouv+(Couv*Signe) ;
    Q2.Next ;
    END ;

// Alim. pour affichage et recaps
TotBaseFact[2]:=TotCouv ; TotBaseFact[3]:=Q.FindField('E_ECHEENC1').AsFloat ;
TotBaseFact[4]:=Q.FindField('E_ECHEENC2').AsFloat ; TotBaseFact[5]:=Q.FindField('E_ECHEENC3').AsFloat ;
TotBaseFact[6]:=Q.FindField('E_ECHEENC4').AsFloat ; TotBaseFact[7]:=Q.FindField('E_ECHEDEBIT').AsFloat ;
TotBaseFact[8]:=TabTva[1] ; TotBaseFact[9]:=TabTva[2] ;
TotBaseFact[10]:=TabTva[3] ; TotBaseFact[11]:=TabTva[4] ;
TotBaseFact[12]:=TabTva[5] ;
AddGroup(LTiers, [Qr1E_AUXILIAIRE],[0,0,TotBaseFact[2],TotBaseFact[3],TotBaseFact[4],TotBaseFact[5],TotBaseFact[6],TotBaseFact[7],TabTva[1],TabTva[2],TabTva[3],TabTva[4],TabTva[5]]) ;
Quoi:=QuoiMvt(2) ;
if not PrintBand then Exit ;
AfficheBases('',TotBaseFact) ;
end;

procedure TFQRTvaAcc.Q2AfterOpen(DataSet: TDataSet);
begin
  inherited;
Qr2Debit:=TFloatField(Q2.FindField('DEBIT2')) ;
Qr2Credit:=TFloatField(Q2.FindField('CREDIT2')) ;
Qr2Couv:=TFloatField(Q2.FindField('COUVERTURE2')) ;
Qr2Regime:=TStringField(Q2.FindField('E_REGIMETVA')) ;
end;

procedure TFQRTvaAcc.QAfterOpen(DataSet: TDataSet);
begin
  inherited;
Qr1Modepaie:=TFloatField(Q.FindField('E_MODEPAIE')) ;
Qr1Debit:=TFloatField(Q.FindField('DEBIT')) ;
Qr1Credit:=TFloatField(Q.FindField('CREDIT')) ;
Qr1Couv:=TFloatField(Q.FindField('COUVERTURE')) ;
Qr1E_AUXILIAIRE:=TStringField(Q.FindField('E_AUXILIAIRE')) ;
Qr1T_LIBELLE:=TStringField(Q.FindField('LIBAUX')) ;
end;

procedure TFQRTvaAcc.Q2BeforeOpen(DataSet: TDataSet);
begin
  inherited;
Q2.Params[0].AsString:=Q.FindField('E_AUXILIAIRE').AsString ;
Q2.Params[1].AsString:=Q.FindField('E_GENERAL').AsString ;
Q2.Params[2].AsString:=Q.FindField('E_ETATLETTRAGE').AsString ;
Q2.Params[3].AsString:=Q.FindField('E_LETTRAGE').AsString ;
end;

procedure TFQRTvaAcc.BFinAuxBeforePrint(var PrintBand: Boolean;
  var Quoi: string);
begin
  inherited;
if not CritEdt.GL.FormatPrint.PrSepCompte[8] then BEGIN PrintBand:=False ; Exit ; END ;
Quoi:=QuoiAux(1) ;
end;

procedure TFQRTvaAcc.QRDLMultiNeedData(var MoreData: Boolean);
Var CodRupt, LibRupt : String ;
    Tot : Array[0..12] of Double ;
    QuelleRupt : integer ;
begin
  inherited;
MoreData:=PrintGroup(LTiers, Q, [Qr1E_AUXILIAIRE], CodRupt, LibRupt, Tot,QuelleRupt) ;
If MoreData Then
   BEGIN
   TotalTiers.Caption:=HM.Mess[7]+' '+CodRupt+' '+QR1T_LIBELLE.AsString ;
   AfficheBases('Aux' ,Tot) ;
   AddRecap(LRecapG,['Total'],[''],Tot) ;
   END ;
end;

procedure TFQRTvaAcc.QRDLRecapGNeedData(var MoreData: Boolean);
Var Cod,Lib : String ;
    TotRecapG : Array[0..12] Of Double ;
begin
  inherited;
MoreData:=PrintRecap(LRecapG, Cod, Lib, TotRecapG) ;
if MoreData then
  BEGIN
  AfficheBases('G',TotRecapG) ;
  END ;
end;

procedure TFQRTvaAcc.BRecapGAfterPrint(BandPrinted: Boolean);
begin
  inherited;
if BRecapG.ForceNewPage then BRecapG.ForceNewPage:=FALSE ;
end;

procedure TFQRTvaAcc.EntetePageBeforePrint(var PrintBand: Boolean;
  var Quoi: string);
begin
  inherited;
if (CritEdt.GL.FormatPrint.PrSepCompte[10]) and (Q.Eof) then
   BEGIN
   TLE_AUXILIAIRE.Caption:='' ; TE_DATECOMPTABLE.Caption:='' ; TE_JOURNAL.Caption:='' ;
   TE_PIECE.Caption:='' ; TE_LETTRAGE.Caption:='' ; TE_REFINTERNE.Caption:='' ;
   TE_LIBELLE.Caption:='' ; TE_MODEPAIE.Caption:='' ;
   TE_TTC.Caption:='' ; TE_COUV.Caption:='' ;
   END ;
end;

procedure TFQRTvaAcc.BRecapGBeforePrint(var PrintBand: Boolean;
  var Quoi: string);
begin
  inherited;
if not CritEdt.GL.FormatPrint.PrSepCompte[10] then PrintBand:=False ;
end;

end.
