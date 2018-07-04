unit QRGLAuxL;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  QR, hmsgbox, HQuickrp, DB,
  {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}
  StdCtrls, Buttons, Hctrls, ExtCtrls,
  Mask, Hcompte, ComCtrls, UtilEdt, Ent1, HEnt1, CpteUtil, QRRupt, EdtLegal,
  HQry, CritEdt, HSysMenu, Menus, HTB97, HPanel, UiUtil, tCalcCum ;

procedure GLAuxiliaireL ;
// GC - 20/12/2001
procedure GLAuxiliaireLChaine( vCritEdtChaine : TCritEdtChaine );
// GC - 20/12/2001 - FIN
procedure GLAuxiliaireLZoom(Crit : TCritEDT) ;

type
  TFGdLivAuxL = class(TFQR)
    FValide: TCheckBox;
    HLabel9: THLabel;
    FNumPiece1: TMaskEdit;
    Label12: TLabel;
    FNumPiece2: TMaskEdit;
    FRefInterne: TEdit;
    HLabel10: THLabel;
    FSautPage: TCheckBox;
    FLigneCptPied: TCheckBox;
    QEcr: TQuery;
    SEcr: TDataSource;
    QRDLCpt: TQRDetailLink;
    QRDLMulti: TQRDetailLink;
    DLRupt: TQRDetailLink;
    MsgBox: THMsgBox;
    RValide: TQRLabel;
    QRLabel17: TQRLabel;
    QRLabel10: TQRLabel;
    RRefInterne: TQRLabel;
    QRLabel13: TQRLabel;
    RNumPiece1: TQRLabel;
    QRLabel14: TQRLabel;
    RNumPiece2: TQRLabel;
    TRLegende: TQRLabel;
    RLegende: TQRLabel;
    TLE_VALIDE: TQRLabel;
    TLE_DATECOMPTABLE: TQRLabel;
    TE_PIECE: TQRLabel;
    TE_REFINTERNE: TQRLabel;
    TE_LIBELLE: TQRLabel;
    TLE_DEBIT: TQRLabel;
    TLE_CREDIT: TQRLabel;
    TLE_SOLDE: TQRLabel;
    TLE_LETTRAGE: TQRLabel;
    TITRE1REP: TQRLabel;
    REPORT1DEBIT: TQRLabel;
    REPORT1CREDIT: TQRLabel;
    REPORT1SOLDE: TQRLabel;
    T_AUXILIAIRE: TQRLabel;
    TCumul: TQRLabel;
    CumDebit: TQRLabel;
    CumSolde: TQRLabel;
    CumCredit: TQRLabel;
    BDetailEcr: TQRBand;
    LE_DATECOMPTABLE: TQRDBText;
    LE_CREDIT: TQRLabel;
    LE_DEBIT: TQRLabel;
    E_LIBELLE: TQRDBText;
    E_REFINTERNE: TQRDBText;
    LE_LETTRAGE: TQRDBText;
    LE_SOLDE: TQRLabel;
    LE_VALIDE: TQRLabel;
    E_PIECELIGECH: TQRLabel;
    BSDMulti: TQRBand;
    LibTotalMulti: TQRLabel;
    TotMultiDebit: TQRLabel;
    TotMultiCredit: TQRLabel;
    TotSoldeMulti: TQRLabel;
    BFCompteAux: TQRBand;
    T_AUXILIAIRE_: TQRLabel;
    TotSoldeCptAux: TQRLabel;
    TotCptCredit: TQRLabel;
    TotCptDebit: TQRLabel;
    BRupt: TQRBand;
    DebitRupt: TQRLabel;
    TCodRupt: TQRLabel;
    CreditRupt: TQRLabel;
    SoldeRupt: TQRLabel;
    QRLabel33: TQRLabel;
    TotGenDebit: TQRLabel;
    TotGenCredit: TQRLabel;
    TotalSoldes: TQRLabel;
    QRBand2: TQRBand;
    Trait5: TQRLigne;
    Trait4: TQRLigne;
    Trait3: TQRLigne;
    Trait2: TQRLigne;
    Trait1: TQRLigne;
    Trait0: TQRLigne;
    Ligne1: TQRLigne;
    TITRE2REP: TQRLabel;
    REPORT2DEBIT: TQRLabel;
    REPORT2CREDIT: TQRLabel;
    REPORT2SOLDE: TQRLabel;
    TE_JOURNAL: TQRLabel;
    LE_JOURNAL: TQRDBText;
    FSoldeProg: TCheckBox;
    FTotMens: TCheckBox;
    FTotEche: TCheckBox;
    Label1: TLabel;
    QRLabel8: TQRLabel;
    TotDesMvtsDebit: TQRLabel;
    TotDesMvtsCredit: TQRLabel;
    TotDesMvtsSolde: TQRLabel;
    FLigneRupt: TCheckBox;
    BRupTete: TQRBand;
    CodeRuptAu: TQRLabel;
    DLRuptH: TQRDetailLink;
    AnoRupt: TQRLabel;
    CumRupt: TQRLabel;
    AnoDebitRupt: TQRLabel;
    CumDebitRupt: TQRLabel;
    AnoCreditRupt: TQRLabel;
    CumCreditRupt: TQRLabel;
    AnoSoldeRupt: TQRLabel;
    CumSoldeRupt: TQRLabel;
    FLettrable: TCheckBox;
    QRLabel1: TQRLabel;
    RLettrable: TQRLabel;
    TotRupt: TQRLabel;
    procedure FormShow(Sender: TObject);
    procedure QEcrCalcFields(DataSet: TDataSet);
    procedure BDetailBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure BDetailEcrAfterPrint(BandPrinted: Boolean);
    procedure BFCompteAuxBeforePrint(var PrintBand: Boolean;var Quoi: string);
    procedure BFinEtatBeforePrint(var PrintBand: Boolean;var Quoi: string);
    procedure QRDLMultiNeedData(var MoreData: Boolean);
    procedure BSDMultiBeforePrint(var PrintBand: Boolean;var Quoi: string);
    procedure DLRuptNeedData(var MoreData: Boolean);
    procedure BRuptBeforePrint(var PrintBand: Boolean;var Quoi: string);
    procedure BRuptAfterPrint(BandPrinted: Boolean);
    procedure TOPREPORTBeforePrint(var PrintBand: Boolean;var Quoi: string);
    procedure BOTTOMREPORTBeforePrint(var PrintBand: Boolean;var Quoi: string);
    procedure BDetailEcrBeforePrint(var PrintBand: Boolean;  var Quoi: string);
    procedure QEcrAfterOpen(DataSet: TDataSet);
    procedure QAfterOpen(DataSet: TDataSet);
    procedure BFCompteAuxAfterPrint(BandPrinted: Boolean);
    procedure DLRuptHNeedData(var MoreData: Boolean);
    procedure BRupTeteBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure FSansRuptClick(Sender: TObject);
    procedure FRupturesClick(Sender: TObject);
    procedure BDetailAfterPrint(BandPrinted: Boolean);
    procedure BDetailCheckForSpace;
    procedure Timer1Timer(Sender: TObject);
    procedure GenereSQL ; Override ;
    procedure FinirPrint ; Override ;
    procedure InitDivers ; Override ;
    procedure RenseigneCritere ; Override ;
    procedure ChoixEdition ; Override ;
    procedure RecupCritEdt ; Override ;
    function  CritOk : Boolean ; Override ;
  private    { Déclarations privées }
    DansTotGen, IsECC             : Boolean ;
    LibRupt, CodRupt, StReportAux : string ;
    Imprime, Affiche,SautPageRupture : Boolean ;
    AncienEche                    : Integer ;
//    QBal                          : TQuery ;
    TotEdt, TotCptAux,TotAnCRupt,
    TotDesMvts                    : TabTot ;
    LMulti, LRupt                 : TStringList ;
    Tot                           : Array[0..12] of Double ;
    Qr1T_AUXILIAIRE               : TStringField ;
    Qr1T_LIBELLE                  : TStringField ;
    Qr1T_SAUTPAGE                 : TStringField ;
    Qr1T_SOLDEPROGRESSIF          : TStringField ;
    Qr1T_TOTAUXMENSUELS           : TStringField ;
    QR2E_AUXILIAIRE,Qr2E_JOURNAL,
    Qr2E_EXERCICE,Qr2E_VALIDE,
    Qr2E_QUALIFPIECE,Qr2E_LETTRAGE,
    Qr2E_ETATLETTRAGE              : TStringField ;
    Qr2E_NUMEROPIECE,Qr2E_NUMLIGNE,Qr2E_NUMECHE,QR2MULTIECHE : TIntegerFIeld ;
    Qr2E_DateComptable,Qr2E_DATEPAQUETMIN,Qr2E_DATEPAQUETMAX   : TDateTimeField ;
    Qr2DEBIT,Qr2CREDIT,QR2COUV   : TFloatField ;
    FLoading             : Boolean ;
    IsLegal, OkEnteteRup : Boolean ;
    Qr1T_CORRESP1, Qr1T_CORRESP2 : TStringField ;
    QQ : TQuery ;
    procedure GenereSQLBaseL ;
    procedure GenereSubSQL ;
    Function  WhatExisteL : String ;
    Function QuoiAux(i : Integer) : String ;
    Function QuoiMvt : String ;
    Procedure GLAuxZoom(Quoi : String) ;
    Function  QuelDC(Var D,C : Double) : Boolean ;
    Procedure CalculCumulAu(Aux : String ; var CumulAu : TabTot) ;
  public     { Déclarations publiques }
  end;

implementation

{$R *.DFM}

Function TFGdLivAuxL.QuoiAux(i : Integer) : String ;
BEGIN
Inherited ;
Case i Of
  1 : Result:=Qr1T_AUXILIAIRE.AsString+' '+'#'+Qr1T_LIBELLE.AsString+'@'+'2' ;
  2 : Result:=Qr1T_AUXILIAIRE.AsString+' '+'#'+Qr1T_LIBELLE.AsString+' '+
              TotSoldecptAux.Caption+'@'+'2' ;
  END ;
END ;

Function TFGdLivAuxL.QuoiMvt : String ;
BEGIN
Inherited ;
Result:=Qr2E_AUXILIAIRE.AsString+' '+Qr1T_LIBELLE.AsString+' '+Le_Solde.Caption+
        '#'+Qr2E_JOURNAL.AsString+' N° '+IntToStr(Qr2E_NUMEROPIECE.AsInteger)+' '+DateToStr(Qr2E_DateComptable.AsDAteTime)+'-'+
        PrintSolde(Qr2DEBIT.AsFloat,Qr2Credit.AsFloat,CritEDT.Decimale,CritEDT.Symbole,CritEDT.AfficheSymbole)+
        '@'+'5;'+Qr2E_JOURNAL.AsString+';'+UsDateTime(Qr2E_DATECOMPTABLE.AsDateTime)+';'+Qr2E_NUMEROPIECE.AsString+';'+Qr2E_EXERCICE.asString+';'+
        IntToStr(Qr2E_NumLigne.AsInteger)+';' ;
END ;

Procedure TFGdLivAuxL.GLAuxZoom(Quoi : String) ;
Var Lp,i: Integer ;
BEGIN
Inherited ;
Lp:=Pos('@',Quoi) ; If Lp=0 Then Exit ;
i:=StrToInt(Copy(Quoi,Lp+1,1)) ;
If (i=5) Then
   BEGIN
   Quoi:=Copy(Quoi,Lp+3,Length(Quoi)-lp-2) ;
   If QRP.QrPrinter.FSynShiftDblClick Then i:=6 ;
   If QRP.QRPrinter.FSynCtrlDblClick Then i:=11 ;
   END ;
ZoomEdt(i,Quoi) ;
END ;

procedure GLAuxiliaireL ;
var QR : TFGdLivAuxL ;
    Edition : TEdition ;
    PP : THPanel ;
BEGIN
PP:=FindInsidePanel ;
QR:=TFGdLivAuxL.Create(Application) ;
Edition.Etat:=etGlAux ;
QR.QRP.QRPrinter.OnSynZoom:=QR.GLAuxZoom ;
QR.InitType (nbAux,neGL,msAuxEcr,'QRGLAUXL','QR_GLAUXL',TRUE,FALSE,FALSE,Edition) ;
QR.IsLegal:=FALSE ;
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
END ;

// GC - 20/12/2001
procedure GLAuxiliaireLChaine( vCritEdtChaine : TCritEdtChaine );
var QR : TFGdLivAuxL ;
    Edition : TEdition ;
    PP : THPanel ;
BEGIN
PP:=FindInsidePanel ;
QR:=TFGdLivAuxL.Create(Application) ;
Edition.Etat:=etGlAux ;
QR.QRP.QRPrinter.OnSynZoom:=QR.GLAuxZoom ;
QR.CritEdtChaine := vCritEdtChaine;
QR.InitType (nbAux,neGL,msAuxEcr,'QRGLAUXL','QR_GLAUXL',TRUE,FALSE,FALSE,Edition) ;
QR.IsLegal:=FALSE ;
if PP=Nil then
   BEGIN
   try
    QR.Timer1.Enabled := True;
    QR.QRP.QRPrinter.Copies := QR.CritEdtChaine.NombreExemplaire;
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
END ;
// GC - 20/12/2001 - FIN

procedure GLAuxiliaireLZoom(Crit : TCritEDT) ;
var QR : TFGdLivAuxL ;
    Edition : TEdition ;
BEGIN
QR:=TFGdLivAuxL.Create(Application) ;
Edition.Etat:=etGlAux ;
Try
  QR.QRP.QRPrinter.OnSynZoom:=QR.GLAuxZoom ;
  QR.IsLegal:=FALSE ;
  QR.CritEDT:=Crit ;
  QR.InitType (nbAux,neGL,msAuxEcr,'QRGLAUXL','QR_GLAUXL',FALSE,TRUE,FALSE,Edition) ;
finally
  QR.Free ;
  End ;
Screen.Cursor:=SyncrDefault ;
END ;

procedure TFGdLivAuxL.InitDivers ;
BEGIN
InHerited ;
{ Gestion multiéchéances }
AncienEche:= 1 ; TLE_VALIDE.Caption:='' ;
{ Les tableaux de montants }
Fillchar(TotEdt,SizeOf(TotEdt),#0) ;
Fillchar(TotAncRupt,SizeOf(TotAncRupt),#0) ;
BFCompteAux.Frame.DrawTop:=CritEDT.GL.FormatPrint.PrSepCompte[2] ;
BRupt.Frame.DrawTop:=CritEdt.GL.FormatPrint.PrSepCompte[3] ;
//BRupt.Frame.DrawBottom:=BRupt.Frame.DrawTop ;
If CritEdt.GA Then CritEdt.SQLGA:=GetParamListe(NomListeGA,Self) ;
//OkEnteteRup:=(CritEdt.Rupture=rCorresp) ;
If CritEdt.SautPageRupt And (CritEdt.GL.RuptOnly=Avec) And (CritEdt.SautPage<>1) Then BDetail.ForceNewPage:=FALSE ;
SautPageRupture:=FALSE ;
END ;

Procedure TFGdLivAuxL.FinirPrint ;
BEGIN
InHerited ;
VideGroup(LMulti) ; QEcr.Close ;
//If GCalc<>NIL Then BEGIN GCalc.Free ; GCalc:=NIL ; END ;
if CritEdt.Rupture<>rRien then VideRupt(LRupt) ;
Ferme(QQ) ;
END ;

Function TFGdLivAuxL.WhatExisteL : String ;
Var St,St2 : String ;
    SetType : SetttTypePiece ;
begin
St:='(exists (Select E_AUXILIAIRE From ECRITURE ' ;
St:=St+' Where E_AUXILIAIRE=Q.T_AUXILIAIRE ' ;
St:=St+' And E_DATECOMPTABLE<="'+usdatetime(CritEDT.Date2)+'" ' ;
if CritEDT.DeviseSelect<>'' then St:=St+' And E_DEVISE="'+CritEDT.DeviseSelect+'"' ;
St:=St+' And (E_ECRANOUVEAU="N" OR E_ECRANOUVEAU="H") and E_QUALIFPIECE<>"C" ' ;
(*
Case CritEdt.GL.Lettrable of
  2 : St:=St+' And (E_LETTRAGE=""  OR (E_LETTRAGE<>"" AND E_DATEPAQUETMAX<="'+usdatetime(CritEDT.Date2)+'" ))' ;
  1 : St:=St+' And E_LETTRAGE<>"" ' ;
 End ;
*)
SetType:=WhatTypeEcr(CritEdt.QualifPiece,V_PGI.Controleur,AvecRevision.State) ;
St2:=WhereSupp('E_',SetType) ;
If St2<>'' Then St:=St+St2 ;
St:=St+'))' ;
Result:=St ;
end ;

Procedure TFGdLivAuxL.genereSQLBaseL ;
Var LeSauf,LeSQLSauf,NatCpt,St : String ;
    StSelect1,stSelect2,stSelect3,StFrom,StCpt,StOrder,StWhere1,StWhere2 : String ;
    NatureRupt : String3 ;
    P          : String ;
    EOnlyCompteAssocie : Boolean ;
    EPlansCorresp : Byte ;
BEGIN
Q.Close ; Q.SQL.Clear ;
StSelect1:='' ; StSelect2:='' ; StFrom:='' ; StCpt:='' ; StOrder:='' ;
StWhere1:='' ; StWhere2:='' ; stSelect3:='' ;
NatCpt:=CritEDT.NatureCpt ;
LeSauf:=CritEDT.GL.Sauf ; LeSQLSauf:=CritEdt.GL.SQLSauf ;
EOnlyCompteAssocie:=CritEdt.GL.OnlyCptAssocie ;
EPlansCorresp:=CritEDT.GL.PlansCorresp ;
stSelect2:=' ,T_SAUTPAGE, T_SOLDEPROGRESSIF, T_TOTAUXMENSUELS ' ;
StSelect1:='Select T_AUXILIAIRE,T_LIBELLE,T_COLLECTIF, T_TOTDEBANO, T_TOTCREANO, T_TOTDEBE, T_TOTCREE, T_TOTDEBANON1, T_TOTCREANON1 ' ;
Case CritEdt.Rupture of
  rLibres  : BEGIN
             StSelect1:='Select T_AUXILIAIRE,T_LIBELLE,T_COLLECTIF, T_TOTDEBANO, T_TOTCREANO, T_TOTDEBE, T_TOTCREE, T_TOTDEBANON1, T_TOTCREANON1 ' ;
             StSelect3:=', T_TABLE0, T_TABLE1, T_TABLE2, T_TABLE3, T_TABLE4, T_TABLE5, T_TABLE6, T_TABLE7, T_TABLE8, T_TABLE9 ' ;
             END ;
  rCorresp : StSelect1:='Select T_AUXILIAIRE,T_LIBELLE,T_COLLECTIF, T_TOTDEBANO, T_TOTCREANO, T_TOTDEBE, T_TOTCREE, T_TOTDEBANON1, T_TOTCREANON1, T_CORRESP1, T_CORRESP2 ' ;
  End ;
StFrom:=' From TIERS Q Where ' ;
if CritEdt.NatureCpt<>'' then StWhere1:='AND T_NATUREAUXI="'+CritEdt.NatureCpt+'" ' ;
if (OkV2 and (V_PGI.Confidentiel<>'1')) then StWhere1:=StWhere1+'AND T_CONFIDENTIEL<>"1" ' ;
StOrder:=' Order By T_AUXILIAIRE' ;
Case CritEdt.Rupture of
  rLibres  : BEGIN
             IF EOnlyCompteAssocie then StWhere1:=StWhere1+WhereLibre(CritEdt.LibreCodes1,CritEdt.LibreCodes2,fbAux,EOnlyCompteAssocie) ;
             StOrder:='Order By '+OrderLibre(CritEdt.LibreTrie)+'T_AUXILIAIRE' ;
             END ;
  rCorresp : Case EPlansCorresp Of
               1 : StOrder:=' Order By T_CORRESP1, T_AUXILIAIRE' ;
               2 : StOrder:=' Order By T_CORRESP2, T_AUXILIAIRE' ;
               Else StOrder:=' Order By T_AUXILIAIRE' ;
               END ;
  End ;
If StSelect1<>'' Then Q.SQL.Add(StSelect1) ;
If StSelect2<>'' Then Q.SQL.Add(StSelect2) ;
If StSelect3<>'' Then Q.SQL.Add(StSelect3) ;
If StFrom<>'' Then Q.SQL.Add(StFrom) ;
St:=WhatExisteL ; If (St<>'') Then Q.SQL.Add(' '+St) ;
St:=SQLQuelCpt ; If St<>'' Then Q.SQL.Add(St) ;
If StWhere1<>'' Then Q.SQL.Add(StWhere1) ;
If StWhere2<>'' Then Q.SQL.Add(StWhere2) ;
if LeSauf<>'' then Q.SQL.Add(LeSQLSauf) ;
St:=CritEDT.SQLPlusBase ; If St<>'' Then Q.SQL.Add(St) ;
Case CritEdt.Rupture of
  rCorresp : BEGIN
             If EOnlyCompteAssocie Then
                BEGIN
                P:='T' ;NatureRupt:='AU'+IntToStr(EPlansCorresp) ;
                Q.SQL.Add(' AND Exists(SELECT CR_CORRESP FROM CORRESP WHERE CR_TYPE="'+NatureRupt+'"') ;
                Q.SQL.Add(' AND CR_CORRESP=Q.'+P+'_CORRESP'+IntToStr(EPlansCorresp)+')') ;
                END ;
              END ;
  End ;
If StOrder<>'' Then Q.SQL.Add(StOrder) ;
ChangeSQL(Q) ; Q.Open ;
END ;

procedure TFGdLivAuxL.GenereSubSQL ;
{ Construction de la requête SQL en fonction du multicritère }
//Var St : String ;
BEGIN
{ Construction de la clause Select de la SQL }
QEcr.Close ;
QEcr.SQL.Clear ;
QEcr.SQL.Add('Select E_AUXILIAIRE,E_EXERCICE,E_DATECOMPTABLE,E_NUMEROPIECE,E_NUMLIGNE,') ;
QEcr.SQL.Add(       'E_REFINTERNE,E_ETABLISSEMENT,E_LIBELLE,E_NUMECHE,E_DEVISE,E_VALIDE,') ;
// GP : bug edition avec filtre sur N° pièce trop grand (à cause du *10)
//QEcr.SQL.Add(       'E_NUMEROPIECE*10+E_NUMLIGNE as MULTIECHE,E_LETTRAGE,E_GENERAL,') ;
QEcr.SQL.Add(       'E_NUMLIGNE as MULTIECHE,E_LETTRAGE,E_GENERAL,') ;
QEcr.SQL.Add(       'E_JOURNAL, E_NATUREPIECE, E_QUALIFPIECE, E_DATEPAQUETMIN, E_DATEPAQUETMAX, E_LETTRAGE, E_ETATLETTRAGE, ') ;
If CritEdt.SQLGA<>'' Then QECR.SQL.Add(CritEdt.SQLGA) ;
Case CritEDT.Monnaie of
  0 : BEGIN
      QEcr.SQL.Add('E_DEBIT DEBIT, E_CREDIT CREDIT, E_COUVERTURE COUV ') ;
      END ;
  1 : BEGIN
      QEcr.SQL.Add('E_DEBITDEV DEBIT, E_CREDITDEV CREDIT, E_COUVERTUREDEV COUV ') ;
      END ;
end ;

{ Tables explorées par la SQL }
QEcr.SQL.Add(' From ECRITURE ') ;
{ Construction de la clause Where de la SQL }
QEcr.SQL.Add(' Where E_AUXILIAIRE=:T_AUXILIAIRE ') ;
QEcr.SQL.Add(' And E_DATECOMPTABLE>="'+usdatetime(CritEDT.GL.Date21)+'" And E_DATECOMPTABLE<="'+usdatetime(CritEDT.Date2)+'" ') ;
if FExercice.ItemIndex>0 then QEcr.SQL.Add(' And E_EXERCICE="'+CritEDT.Exo.Code+'" ') ;
QEcr.SQL.Add(   ' And E_NUMEROPIECE>='+IntToStr(CritEDT.GL.NumPiece1)+' ') ;
QEcr.SQL.Add(   ' And E_NUMEROPIECE<='+IntToStr(CritEDT.GL.NumPiece2)+' ') ;
QEcr.SQL.Add(TraduitNatureEcr(CritEDT.Qualifpiece, 'E_QUALIFPIECE', true, CritEdt.ModeRevision)) ;
if CritEDT.RefInterne<>'' then QEcr.SQL.Add(' And UPPER(E_REFINTERNE) like "'+TraduitJoker(CritEDT.RefInterne)+'" ' );
if CritEDT.Etab<>'' then QEcr.SQL.Add(' And E_ETABLISSEMENT="'+CritEDT.Etab+'"') ;
if CritEDT.Valide<>'g' then QEcr.SQL.Add(' And E_VALIDE="'+CritEDT.Valide+'" ') ;
if CritEDT.DeviseSelect<>'' then QEcr.SQL.Add(' And E_DEVISE="'+CritEDT.DeviseSelect+'"') ;
QEcr.SQL.Add(' And E_ECRANOUVEAU="N" and E_QUALIFPIECE<>"C" ') ;
(*
Case CritEdt.GL.Lettrable of
  2 : QEcr.SQL.Add(' And (E_LETTRAGE=""  OR (E_LETTRAGE<>"" AND E_DATEPAQUETMAX<="'+usdatetime(CritEDT.Date2)+'" ))') ;
  1 : QEcr.SQL.Add(' And E_LETTRAGE<>"" ') ;
 End ;
*)
{ Construction de la clause Order By de la SQL }
If CritEDT.SQLPlus<>'' Then QEcr.SQL.Add(CritEDT.SQLPlus) ;
QEcr.SQL.Add(' Order By E_AUXILIAIRE,E_EXERCICE,E_DATECOMPTABLE,E_NUMEROPIECE,E_NUMLIGNE,E_NUMECHE') ;
ChangeSQL(QEcr) ; QEcr.Open ;
END ;

procedure TFGdLivAuxL.GenereSQL ;
BEGIN
InHerited ;
GenereSQLBaseL ;
GenereSubSQL ;
END ;

procedure TFGdLivAuxL.RenseigneCritere ;
Var St : String ;
{ Récupération des champs du multicritère dans l'entête d'état }
BEGIN
Inherited ;
If EstSerie(S5) Then RLegende.Caption:=MsgRien.Mess[7] ;
RRefInterne.Caption:=FRefInterne.Text ;
RNumPiece1.Caption:=IntToStr(CritEdt.GL.NumPiece1)   ; RNumPiece2.Caption:=IntToStr(CritEdt.GL.NumPiece2) ;
(*
if FSansRupt.Checked then BEGIN RSansRupt.Caption:='þ' ; RAvecRupt.Caption:='o' ; END ;
if FAvecRupt.Checked then BEGIN RSansRupt.Caption:='o' ; RAvecRupt.Caption:='þ' ; END ;
if FPlanRupt.Enabled then RPlanRupt.Caption:=FPlanRupt.Text ;
*)
CaseACocher(FValide,RValide) ; CaseACocher(FLettrable,RLettrable) ;
(*
CaseACocher(FSoldeProg,RSoldeProg) ;
CaseACocher(FSautPage,RSautPage)        ; CaseACocher(FTotMens,RTotMens) ;
CaseACocher(FLigneCptPied,RLigneCptPied);
CaseACocher(FTotEche,RTotEche)          ;
*)
If CritEdt.GL.Lettrable=2 Then St:=' '+MsgBox.Mess[17] Else St:=' '+MsgBox.Mess[18] ;
DateCumulAuGL(TCumul,CritEdt,MsgBox.Mess[2]) ; TCumul.Caption:=TCumul.Caption+St ;
DateCumulAuGL(AnoRupt,CritEdt,MsgBox.Mess[2]) ; AnoRupt.Caption:=AnoRupt.Caption+St ;
END ;

procedure TFGdLivAuxL.ChoixEdition ;
Var St,StNum,stDen,StF1,StF2,StV8 : String ;
{ Initialisation des options d'édition }
BEGIN
Inherited ;
DLRupt.PrintBefore:=TRUE ;
if CritEdt.Rupture=rLibres then
   BEGIN
   DLRupt.PrintBefore:=FALSE ;
   ChargeGroup(LRupt,['T00','T01','T02','T03','T04','T05','T06','T07','T08','T09']) ;
   END else
if CritEdt.Rupture=rRuptures then
   BEGIN
   ChargeRupt(LRupt, 'RUT', CritEdt.GL.PlanRupt, '', '') ;
   NiveauRupt(LRupt) ;
   END Else
if CritEdt.Rupture=rCorresp then
   BEGIN
   ChargeRuptCorresp(LRupt, CritEdt.GL.PlanRupt, '', '', False) ;
   NiveauRupt(LRupt) ;
   END ;
ChargeGroup(LMulti,[MsgBox.Mess[5],MsgBox.Mess[6]]) ;
ChgMaskChamp(Qr2DEBIT , CritEDT.Decimale, CritEDT.AfficheSymbole, CritEDT.Symbole, False) ;
ChgMaskChamp(Qr2CREDIT, CritEDT.Decimale, CritEDT.AfficheSymbole, CritEDT.Symbole, False) ;
//If CritEdt.GA then BDetailEcr.Height:=Sup1.Top+Sup1.Height else BDetailEcr.Height:=Sup1.Top ;

StNum:='(E_DEBIT+E_DEBITDEV)' ;
StDen:='(E_DEBIT+E_DEBITDEV+E_CREDIT+E_CREDITDEV)' ;
StF1:='('+StNum+'/'+StDen+')' ;
StNum:='(E_CREDIT+E_CREDITDEV)' ;
StF2:='('+StNum+'/'+StDen+')' ;
St:='SELECT ' ;
Case CritEDT.Monnaie of
  0 : BEGIN
      St:=St+' SUM(E_COUVERTURE*'+StF1+'), SUM(E_COUVERTURE*'+StF2+'), SUM(E_DEBIT), SUM(E_CREDIT) '  ;
      END ;
  1 : BEGIN
      St:=St+' SUM(E_COUVERTUREDEV*'+StF1+'), SUM(E_COUVERTUREDEV*'+StF2+'), SUM(E_DEBITDEV), SUM(E_CREDITDEV) '  ;
      END ;
end ;
St:=St+', E_DATEPAQUETMAX ' ;
{ Tables explorées par la SQL }
St:=St+' From ECRITURE ' ;
{ Construction de la clause Where de la SQL }
St:=St+' Where E_AUXILIAIRE=:AUX ' ;
St:=St+' And E_DATECOMPTABLE<"'+usdatetime(CritEDT.GL.Date21)+'" ' ;
StV8:=LWhereV8 ; if StV8<>'' then St:=St+' AND '+StV8+' ' ;
St:=St+TraduitNatureEcr(CritEDT.Qualifpiece, 'E_QUALIFPIECE', true, CritEdt.ModeRevision) ;
if CritEDT.RefInterne<>'' then St:=St+' And UPPER(E_REFINTERNE) like "'+TraduitJoker(CritEDT.RefInterne)+'" ' ;
if CritEDT.Etab<>'' then St:=St+' And E_ETABLISSEMENT="'+CritEDT.Etab+'"' ;
if CritEDT.Valide<>'g' then St:=St+' And E_VALIDE="'+CritEDT.Valide+'" ' ;
if CritEDT.DeviseSelect<>'' then St:=St+' And E_DEVISE="'+CritEDT.DeviseSelect+'"' ;
St:=St+' And (E_ECRANOUVEAU="N" OR E_ECRANOUVEAU="H") and E_QUALIFPIECE<>"C" ' ;
St:=St+' GROUP BY E_DATEPAQUETMAX ' ;
(*
Case CritEdt.GL.Lettrable of
  2 : St:=St+' And (E_ETATLETTRAGE="AL" OR E_ETATLETTRAGE="PL")' ;
  1 : St:=St+' And (E_ETATLETTRAGE="PL" OR E_ETATLETTRAGE="TL")' ;
 End ;
Case CritEdt.GL.Lettrable of
  2 : St:=St+' And (E_LETTRAGE=""  OR (E_LETTRAGE<>"" AND E_DATEPAQUETMAX<="'+usdatetime(CritEDT.Date2)+'" ))' ;
  1 : St:=St+' And E_LETTRAGE<>"" ' ;
 End ;
*)
QQ:=PrepareSQL(St,TRUE) ;
END ;

{***************** EVENEMENTS LIES A L'AFFICHAGE DES BANDES ********************}

Function TFGdLivAuxL.QuelDC(Var D,C : Double) : Boolean ;
BEGIN
Result:=TRUE ;
D:=0 ; C:=0 ;
If QR2E_ETATLETTRAGE.AsString='AL' Then BEGIN D:=QR2DEBIT.AsFloat ; C:=QR2CREDIT.AsFloat ; END Else
  BEGIN
  If QR2E_DATEPAQUETMAX.AsDAteTime>CritEdt.Date2 Then BEGIN D:=QR2DEBIT.AsFloat ; C:=QR2CREDIT.AsFloat ; END Else
    BEGIN
    If Arrondi(QR2DEBIT.AsFloat,CritEdt.Decimale)=0 Then C:=Arrondi(QR2CREDIT.AsFloat-QR2Couv.AsFloat,CritEdt.Decimale)
                                                    Else D:=Arrondi(QR2DEBIT.AsFloat-QR2Couv.AsFloat,CritEdt.Decimale) ;
    If (Arrondi(D,CritEdt.Decimale)=0) And (Arrondi(C,CritEdt.Decimale)=0) Then Result:=FALSE ;
    END ;
  END ;
END ;

procedure TFGdLivAuxL.BDetailEcrBeforePrint(var PrintBand: Boolean; var Quoi: string);
Var i : Integer ;
    CptRupt : String ;
    D,C : Double ;
begin
Inherited ;
//If QEcr.Eof Then BEGIN PrintBand:=FALSE ; Exit ; END ;
PrintBand:=Affiche And (Not QR2E_AUXILIAIRE.IsNull) ;
if PrintBand then
   BEGIN
   If Not QuelDC(D,C) Then BEGIN PrintBand:=FALSE ; Exit ; END ;
   Le_Debit.Caption:= AfficheMontant(CritEdt.FormatMontant,CritEdt.Symbole,D,CritEdt.AfficheSymbole ) ;
   Le_Credit.Caption:=AfficheMontant(CritEdt.FormatMontant,CritEdt.Symbole,C,CritEdt.AfficheSymbole ) ;
   If IsECC then
      BEGIN
      LE_DEBIT.Caption:='0' ;
      LE_CREDIT.Caption:='0' ;
      END ;
   If CritEdt.GA Then For i:=1 To 7 Do AlimListe(QEcr,i,Self) ;
   E_PIECELIGECH.Caption:=Qr2E_NUMEROPIECE.AsString+' / '+Qr2E_NUMLIGNE.AsString+'   '+Qr2E_NUMECHE.AsString ;
   AddGroup(LMulti, [Qr2E_DATECOMPTABLE, Qr2MULTIECHE], [D, C]) ;
   LE_VALIDE.Caption:=ValiQuali(QR2E_VALIDE.AsString,QR2E_QUALIFPIECE.AsString) ;
   { Affectation Du Calcul du Solde Progressif O/N Sinon, d'aprés l'info sur le compte }
   Case CritEdt.SoldeProg of
     0 : if Qr1T_SOLDEPROGRESSIF.AsString='X' then LE_SOLDE.Caption:=PrintSolde(ProgressDebit+D,ProgressCredit+Qr2CREDIT.asFloat,CritEDT.Decimale,CritEDT.Symbole,CritEDT.AfficheSymbole)
                                              Else LE_SOLDE.Caption:=PrintSolde(D,C,CritEDT.Decimale,CritEDT.Symbole,CritEDT.AfficheSymbole) ;
     1 : LE_SOLDE.Caption:=PrintSolde(ProgressDebit+D,ProgressCredit+C,CritEDT.Decimale,CritEDT.Symbole,CritEDT.AfficheSymbole) ;
     2 : LE_SOLDE.Caption:=PrintSolde(D,C,CritEDT.Decimale,CritEDT.Symbole,CritEDT.AfficheSymbole) ;
    end ;
   TotDesMvts[1].TotDebit :=Arrondi(TotDesMvts[1].TotDebit + D,CritEdt.Decimale) ;
   TotDesMvts[1].TotCredit:=Arrondi(TotDesMvts[1].TotCredit+C,CritEdt.Decimale) ;
   TotCptAux[1].TotDebit :=Arrondi(TotCptAux[1].TotDebit + D,CritEDT.Decimale) ;
   TotCptAux[1].TotCredit:=Arrondi(TotCptAux[1].TotCredit+C,CritEDT.Decimale) ;
   AddReport([1,2],CritEDT.GL.FormatPrint.Report,D,C,CritEDT.Decimale) ;

   Case CritEdt.Rupture of
     rLibres   : AddGroupLibre(LRupt,Q,fbAux,CritEdt.LibreTrie,[1,D,C,0,0,0,0]) ;
     rRuptures : AddRupt(LRupt,Qr2E_AUXILIAIRE.AsString,[1,D,C,0,0,0,0]) ;
     rCorresp  : BEGIN
                 Case CritEDT.GL.PlansCorresp Of
                   1 : If Qr1T_CORRESP1.AsString<>'' Then CptRupt:=Qr1T_CORRESP1.AsString+Qr2E_AUXILIAIRE.AsString
                                                     Else CptRupt:='.'+Qr2E_AUXILIAIRE.AsString ;
                   2 : If Qr1T_CORRESP2.AsString<>'' Then CptRupt:=Qr1T_CORRESP2.AsString+Qr2E_AUXILIAIRE.AsString
                                                     Else CptRupt:='.'+Qr2E_AUXILIAIRE.AsString ;
                   Else CptRupt:=Qr2E_AUXILIAIRE.AsString ;
                  End ;
                 AddRuptCorres(LRupt,CptRupt,[1,D,C,0,0,0,0])
                 END ;
     End ;
   END ;
Quoi:=QuoiMvt ;
end;

procedure TFGdLivAuxL.RecupCritEDT ;
Var NonLibres : Boolean ;
BEGIN
Inherited ;
With CritEDT Do
  BEGIN
  If FNumPiece1.Text<>'' then GL.NumPiece1:=StrToInt(FNumPiece1.Text) else GL.NumPiece1:=0 ;
  If FNumPiece2.Text<>'' then GL.NumPiece2:=StrToInt(FNumPiece2.Text) else GL.NumPiece2:=999999999 ;
  RefInterne:=FRefInterne.Text ;
  if FValide.State=cbGrayed then Valide:='g' ;
  if FValide.State=cbChecked then Valide:='X' ;
  if FValide.State=cbUnChecked then Valide:='-' ;
  if FSautPage.State=cbGrayed then SautPage:=0 ;
  if FSautPage.State=cbChecked then SautPage:=1 ;
  if FSautPage.State=cbUnChecked then SautPage:=2 ;
  if FTotMens.State=cbGrayed then TotMens:=0 ;
  if FTotMens.State=cbChecked then TotMens:=1 ;
  if FTotMens.State=cbUnChecked then TotMens:=2 ;
  if FSoldeProg.Checked then SoldeProg:=1 Else SoldeProg:=2 ;
  if FLettrable.State=cbUnChecked then GL.Lettrable:=2 ;
  if FLettrable.State=cbChecked then GL.Lettrable:=1 ;
  if FLettrable.State=cbGrayed then GL.Lettrable:=0 ;
  GL.TotEche:=FTotEche.Checked ;
  QualifPiece:=FNatureEcr.Value ;
  GL.RuptOnly:=QuelleTypeRupt(1,FSAnsRupt.Checked,FAvecRupt.Checked,FALSE) ;

  NonLibres:=((Rupture=rRuptures) or (Rupture=rCorresp)) ;
  if NonLibres then GL.PlanRupt:=FPlanRuptures.Value ;
  If (CritEdt.Rupture=rCorresp) Then GL.PlansCorresp:=FPlanRuptures.ItemIndex+1 ;
  GL.OnlyCptAssocie:=(Rupture<>rRien) and FOnlyCptAssocie.Checked ;
  With GL.FormatPrint Do
     BEGIN
     PrSepCompte[2]:=FLigneCptPied.Checked ;
     PrSepCompte[3]:=FLigneRupt.Checked ;
     END ;
 END ;
END ;

function  TFGdLivAuxL.CritOk : Boolean ;
BEGIN
Result:=Inherited CritOK ;
if Result then
   BEGIN
//  QBal:=PrepareTotCpt(fbAux,[],Dev,Etab,Exo,DevEnP) ;
(*
   Gcalc:=TGCalculCum.create(Un,fbAux,fbaux,QuelTypeEcr,Dev,Etab,Exo,DevEnP,CritEdt.Monnaie=2,CritEdt.Decimale,V_PGI.OkDecE) ;
   GCalc.initCalcul('','','','',CritEdt.DeviseSelect,CritEdt.Etab,CritEdt.Exo.Code,
                    CritEdt.Date1,CritEdt.GL.Date11,TRUE) ;
*)
   END ;
END ;

procedure TFGdLivAuxL.FormShow(Sender: TObject);
begin
HelpContext:=7560000 ;
//Standards.HelpContext:=7418010 ;
//Avances.HelpContext:=7418020 ;
//Mise.HelpContext:=7418030 ;
//Option.HelpContext:=7418040 ;
//TabRuptures.HelpContext:=7418050 ;
inherited ;
Floading:=FALSE ;
TabSup.TabVisible:=False;
TFGenJoker.Visible:=False;
If FPlanRuptures.Values.Count>0 Then FPlanRuptures.Value:=FPlanRuptures.Values[0] ;
FRefInterne.Text:='' ; FLigneRupt.Enabled:=False ;
FValide.Checked:=True ; FValide.State:=cbGrayed ;
FJoker.MaxLength:=VH^.Cpta[fbAux].lg ;
FSurRupt.Visible:=False ;
FCodeRupt1.Visible:=False  ; FCodeRupt2.Visible:=False ;
TFCodeRupt1.Visible:=False ; TFCodeRupt2.Visible:=False ;
FOnlyCptAssocie.Enabled:=False ;
if CritEdtChaine.Utiliser then InitEtatchaine('QRGLAUXL');
{$IFDEF CCMP}
FNatureCpt.Vide := False;
if (VH^.CCMP.LotCli) then begin FNatureCpt.Plus := ' AND (CO_CODE="AUD" OR CO_CODE="CLI" OR CO_CODE="DIV")'; FNatureCpt.Value:='CLI'; end
                     else begin FNatureCpt.Plus := ' AND (CO_CODE="AUC" OR CO_CODE="DIV" OR CO_CODE="FOU" OR CO_CODE="SAL")'; FNatureCpt.Value:='FOU'; end;
{$ENDIF}
end;

procedure TFGdLivAuxL.QEcrCalcFields(DataSet: TDataSet);
{ Champs calculés dans le cas de ruptures mensuelles ou échéances }
begin
  inherited;
//Qr2MOIS.AsString:=FormatDateTime('mmmm yyyy',Qr2E_DATECOMPTABLE.AsDateTime) ;
//Qr2MULTIECHE.AsString:=IntToStr(Qr2E_NUMEROPIECE.AsInteger+Qr2E_NUMLIGNE.AsInteger) ;
//AncienEche:=Qr2E_NUMECHE.AsInteger ;
end;

procedure TFGdLivAuxL.BDetailEcrAfterPrint(BandPrinted: Boolean);
Var D,C : Double ;
begin
  inherited;
If Not QuelDC(D,C) Then Exit ;
Case CritEdt.SoldeProg of
  1 :  Progressif(True,D,C) ;
  0 :  if Qr1T_SOLDEPROGRESSIF.AsString='X' then Progressif(True,D,C) ;
  End ;
end;

procedure TFGdLivAuxL.BFCompteAuxBeforePrint(var PrintBand: Boolean;  var Quoi: string);
{ Affichage D,C,S pour le compte et incrément sur le total général }
begin
  inherited;
T_AUXILIAIRE_.Caption:=MsgBox.Mess[11]+' '+Qr1T_AUXILIAIRE.AsString+' '+Qr1T_LIBELLE.AsString ;
PrintBand:=Affiche ;
If Not PrintBand then Exit ;
if PrintBand then
   BEGIN
   TotCptDebit.Caption:= AfficheMontant(CritEDT.FormatMontant, CritEDT.Symbole,TotCptAux[1].TotDebit, CritEDT.AfficheSymbole) ;
   TotCptCredit.Caption:=AfficheMontant(CritEDT.FormatMontant, CritEDT.Symbole,TotCptAux[1].TotCredit, CritEDT.AfficheSymbole) ;
   TotSoldeCptAux.Caption:=PrintSolde(TotCptAux[1].TotDebit,TotCptAux[1].TotCredit,CritEDT.Decimale,CritEDT.Symbole,CritEDT.AfficheSymbole) ;
   TotDesMvtsDebit.Caption:= AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole,TotDesMvts[1].TotDebit, CritEdt.AfficheSymbole) ;
   TotDesMvtsCredit.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole,TotDesMvts[1].TotCredit, CritEdt.AfficheSymbole) ;
   TotDesMvtsSolde.Caption:=PrintSolde(TotDesMvts[1].TotDebit,TotDesMvts[1].TotCredit,CritEdt.Decimale,CritEdt.Symbole,CritEdt.AfficheSymbole) ;
   TotEdt[1].TotDebit :=Arrondi(TotEdt[1].TotDebit+TotCptAux[1].TotDebit,CritEDT.Decimale) ;
   TotEdt[1].TotCredit:=Arrondi(TotEdt[1].TotCredit+TotCptAux[1].TotCredit,CritEDT.Decimale) ;
   END ;
Quoi:=QuoiAux(2) ;
end;

procedure TFGdLivAuxL.BFinEtatBeforePrint(var PrintBand: Boolean;  var Quoi: string);
{ Affichage D,C,S pour le total général }
begin
  inherited;
TotGenDebit.Caption:=AfficheMontant(CritEDT.FormatMontant, CritEDT.Symbole,TotEdt[1].TotDebit, CritEDT.AfficheSymbole) ;
TotGenCredit.Caption:=AfficheMontant(CritEDT.FormatMontant, CritEDT.Symbole,TotEdt[1].TotCredit, CritEDT.AfficheSymbole) ;
TotalSoldes.Caption:=PrintSolde(TotEdt[1].TotDebit,TotEdt[1].TotCredit,CritEDT.Decimale,CritEDT.Symbole,CritEDT.AfficheSymbole) ;
BOTTOMREPORT.enabled:=FALSE ;
end;

procedure TFGdLivAuxL.QRDLMultiNeedData(var MoreData: Boolean);
Var QuelleRupt : integer ;
begin
  inherited;
//If QEcr.Eof Then BEGIN MoreData:=FALSE ; Exit ; END ;
MoreData:=PrintGroup(LMulti, QEcr, [Qr2E_DATECOMPTABLE , Qr2MULTIECHE], CodRupt, LibRupt, Tot,Quellerupt) ;
TotMultiDebit.Caption:=AfficheMontant(CritEDT.FormatMontant, CritEDT.Symbole,Tot[0], CritEDT.AfficheSymbole) ;
TotMultiCredit.Caption:=AfficheMontant(CritEDT.FormatMontant, CritEDT.Symbole,Tot[1],CritEDT.AfficheSymbole) ;
If MoreData Then
   BEGIN
   if LibRupt=MsgBox.Mess[6] then
      BEGIN
      LibTotalMulti.Caption:=MsgBox.Mess[6]+' '+IntToStr(Qr2E_NUMEROPIECE.AsInteger)+' / '+Qr2E_NUMLIGNE.AsString ;
      BSDMulti.Font.Color:=clPurple ;
      TotSoldeMulti.Caption:='' ;
      if Qr2E_NUMECHE.AsInteger =1 then Imprime:=False
                                   else Imprime:=CritEDT.GL.TotEche ;
      END ;
   if LibRupt=MsgBox.Mess[5] then
      BEGIN
      LibTotalMulti.Caption:=MsgBox.Mess[7]+' '+CodRupt ;
      BSDMulti.Font.Color:=clRed ;
      TotSoldeMulti.Caption:=PrintSolde(Tot[0],Tot[1], CritEDT.Decimale,CritEDT.Symbole,CritEDT.AfficheSymbole) ;
      (* Rony 10/04/97 -- Pas de Progerssif
      Case CritEdt.SoldeProg of
        0 : if Qr1T_SOLDEPROGRESSIF.AsString='X' then TotSoldeMulti.Caption:=PrintSolde(ProgressDebit,ProgressCredit,CritEDT.Decimale,CritEDT.Symbole,CritEDT.AfficheSymbole)
                                                 Else TotSoldeMulti.Caption:=PrintSolde(Tot[0],Tot[1], CritEDT.Decimale,CritEDT.Symbole,CritEDT.AfficheSymbole) ;
        1 : TotSoldeMulti.Caption:=PrintSolde(ProgressDebit,ProgressCredit,CritEDT.Decimale,CritEDT.Symbole,CritEDT.AfficheSymbole) ;
        2 : TotSoldeMulti.Caption:=PrintSolde(Tot[0],Tot[1], CritEDT.Decimale,CritEDT.Symbole,CritEDT.AfficheSymbole) ;
       end ;
       *)
      Case CritEdt.TotMens of             { Total Mensuel O/N Sinon, d'aprés l'info sur le compte }
       0 : Imprime:=(Qr1T_TOTAUXMENSUELS.AsString='X') ;
       1 : Imprime:=True  ;
       2 : Imprime:=False  ;
       end;
      END ;
   END ;

//ShowMessage( FloatToStr(Tot[0])+'  '+FloatToStr(Tot[1])+ #13+FloatToStr(Tot[2])+'  '+FloatToStr(Tot[0]-Tot[1]) ) ;
end;

procedure TFGdLivAuxL.BSDMultiBeforePrint(var PrintBand: Boolean;  var Quoi: string);
begin
  inherited;
PrintBand:=Imprime And Affiche ;
end;

(*
Procedure CalculEtAfficheMontantRupture ;

AnoRupt,AnoDebitRupt,AnoCreditRupt,AnoSoldeRupt,
CumRupt,CumDebitRupt,CumCreditRupt,CumSoldeRupt,
TotRupt,DebitRupt,CreditRupt,SoldeRupt : TQrLabel ;
TotRuptM : TTTotRupt ;
BRupt : TQRBand ;
BEGIN

AnoRupt.Visible:=(CritEdt.Rupture in [rCorresp,rLibres]) ;
AnoDebitRupt.Visible:=(CritEdt.Rupture in [rCorresp,rLibres]) ;
AnoCreditRupt.Visible:=(CritEdt.Rupture in [rCorresp,rLibres]) ;
AnoSoldeRupt.Visible:=(CritEdt.Rupture in [rCorresp,rLibres]) ;

CumDebitRupt.Visible:=(CritEdt.Rupture in [rCorresp,rLibres]) ;
CumCreditRupt.Visible:=(CritEdt.Rupture in [rCorresp,rLibres]) ;
CumSoldeRupt.Visible:=(CritEdt.Rupture in [rCorresp,rLibres]) ;
CumRupt.Visible:=(CritEdt.Rupture in [rCorresp,rLibres]) ;
Case CritEdt.Rupture Of
  rCorresp :  BEGIN
              OkEnteteRup:=True ;
              AnoDebitRupt.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, TotRupt[3], CritEdt.AfficheSymbole) ;
              AnoCreditRupt.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, TotRupt[4], CritEdt.AfficheSymbole) ;
              AnoSoldeRupt.Caption:=PrintSolde(TotRupt[3], TotRupt[4], CritEdt.Decimale, CritEdt.Symbole, CritEdt.AfficheSymbole) ;

              CumDebitRupt.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, TotRupt[5], CritEdt.AfficheSymbole) ;
              CumCreditRupt.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, TotRupt[6], CritEdt.AfficheSymbole) ;
              CumSoldeRupt.Caption:=PrintSolde(TotRupt[5], TotRupt[6], CritEdt.Decimale, CritEdt.Symbole, CritEdt.AfficheSymbole) ;

              SumD:=Arrondi(TotRupt[5]+ TotRupt[1],CritEdt.Decimale) ;
              SumC:=Arrondi(TotRupt[6]+TotRupt[2],CritEdt.Decimale) ;
              TCodRupt.Top:=37 ; DebitRupt.Top:=37 ; CreditRupt.Top:=37 ; SoldeRupt.Top:=37 ;
              BRupt.Height:=54 ;
              END ;
  rlibres  :  BEGIN
              OkEnteteRup:=True ;
              AnoDebitRupt.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, TotRupt[3], CritEdt.AfficheSymbole) ;
              AnoCreditRupt.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, TotRupt[4], CritEdt.AfficheSymbole) ;
              AnoSoldeRupt.Caption:=PrintSolde(TotRupt[3], TotRupt[4], CritEdt.Decimale, CritEdt.Symbole, CritEdt.AfficheSymbole) ;

              CumDebitRupt.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, TotRupt[1], CritEdt.AfficheSymbole) ;
              CumCreditRupt.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, TotRupt[2], CritEdt.AfficheSymbole) ;
              CumSoldeRupt.Caption:=PrintSolde(TotRupt[1], TotRupt[2], CritEdt.Decimale, CritEdt.Symbole, CritEdt.AfficheSymbole) ;

              SumD:=Arrondi(TotRupt[3]+ TotRupt[1],CritEdt.Decimale) ;
              SumC:=Arrondi(TotRupt[4]+TotRupt[2],CritEdt.Decimale) ;
              END ;
   Else BEGIN
        TCodRupt.Top:=2 ; DebitRupt.Top:=2 ; CreditRupt.Top:=2 ; SoldeRupt.Top:=2;
        If (CritEdt.Rupture<>rlibres) Then BRupt.Height:=20 ;
        END;
  END ;
If (CritEdt.GL.QuelAN=SansAN) Then
   BEGIN
   AnoDebitRupt.Caption:=MsgBox.Mess[9] ; AnoCreditRupt.Caption:='' ; AnoSoldeRupt.Caption:='' ;
   END ;
DebitRupt.Caption:= AfficheMontant(CritEDT.FormatMontant, CritEDT.Symbole, SumD, CritEDT.AfficheSymbole) ;
CreditRupt.Caption:=AfficheMontant(CritEDT.FormatMontant, CritEDT.Symbole, SumC, CritEDT.AfficheSymbole) ;
SoldeRupt.Caption:=PrintSolde(SumD, SumC, CritEDT.Decimale, CritEDT.Symbole, CritEDT.AfficheSymbole) ;
END ;
*)
procedure TFGdLivAuxL.DLRuptNeedData(var MoreData: Boolean);
Var TotRupt     : TTotRupt ;
    SumD, SumC  : Double ;
    Quellerupt  : Integer ;
    OkOk        : Boolean ;
    CptRupt,Lib1, Stcode  : String ;
    Col         : TColor ;
    LibRuptInf : Array[1..10] Of TRuptInf ;
begin
  inherited;
MoreData:=False ;
//If QEcr.Eof Then BEGIN MoreData:=FALSE ; Exit ; END ;
Case CritEdt.Rupture of
  rLibres   : BEGIN
              MoreData:=PrintGroupLibre(LRupt,Q,fbAux,CritEdt.LibreTrie,CodRupt,LibRupt,Lib1,TotRupt,Quellerupt,Col,LibRuptInf) ;
              If MoreData then
                 BEGIN
                 StCode:=CodRupt ;
                 Delete(StCode,1,Quellerupt+2) ;
                 MoreData:=DansChoixCodeLibre(StCode,Q,fbAux,CritEdt.LibreCodes1,CritEdt.LibreCodes2, CritEdt.LibreTrie) ;
                 BRupt.Font.Color:=Col ;
                 END ;
              END ;
  rRuptures : MoreData:=PrintRupt(LRupt,Qr2E_AUXILIAIRE.AsString,CodRupt,LibRupt,DansTotGen,QRP.EnRupture,TotRupt) ;
  rCorresp  : BEGIN
              OkOk:=TRUE ;
              Case CritEDT.GL.PlansCorresp  Of
                1 : If Qr1T_CORRESP1.AsString<>'' Then CptRupt:=Qr1T_CORRESP1.AsString+Qr2E_AUXILIAIRE.AsString
                                                  Else CptRupt:='.'+Qr2E_AUXILIAIRE.AsString ;
                2 : If Qr1T_CORRESP2.AsString<>'' Then CptRupt:=Qr1T_CORRESP2.AsString+Qr2E_AUXILIAIRE.AsString
                                                  Else CptRupt:='.'+Qr2E_AUXILIAIRE.AsString ;
                Else OkOk:=FALSE ;
                END ;
              If OkOk Then MoreData:=PrintRupt(LRupt,CptRupt,CodRupt,LibRupt,DansTotGen,QRP.EnRupture,TotRupt) Else MoreData:=FALSE ;
              END ;
  End ;
If MoreData Then
   BEGIN
   TCodRupt.Width:=TLE_DATECOMPTABLE.Width+TE_JOURNAL.Width+TE_PIECE.Width+TE_REFINTERNE.Width+1 ;
   TCodRupt.Caption:='' ;
   BRupt.Height:=HauteurBandeRuptIni ;
   Case CritEdt.Rupture of
     rLibres   : BEGIN
                 insert(MsgBox.Mess[15]+' ',CodRupt,Quellerupt+2) ;
                 TCodRupt.Caption:=CodRupt+' '+Lib1 ;
                 AlimRuptSup(HauteurBandeRuptIni,QuelleRupt,TCodRupt.Width,BRupt,LibRuptInf,Self) ;
                 END ;
     rRuptures, rCorresp : TCodRupt.Caption:=CodRupt+'  '+LibRupt ;
     End ;
   SumD:=Arrondi(TotAnCRupt[1].TotDebit+ TotRupt[1],CritEDT.Decimale) ;
   SumC:=Arrondi(TotAnCRupt[1].TotCredit+TotRupt[2],CritEDT.Decimale) ;
   (*
   AnoRupt.Visible:=(CritEdt.Rupture=rCorresp) ;
   AnoDebitRupt.Visible:=(CritEdt.Rupture=rCorresp) ;
   AnoCreditRupt.Visible:=(CritEdt.Rupture=rCorresp) ;
   AnoSoldeRupt.Visible:=(CritEdt.Rupture=rCorresp) ;

   CumDebitRupt.Visible:=(CritEdt.Rupture=rCorresp) ;
   CumCreditRupt.Visible:=(CritEdt.Rupture=rCorresp) ;
   CumSoldeRupt.Visible:=(CritEdt.Rupture=rCorresp) ;
   CumRupt.Visible:=(CritEdt.Rupture=rCorresp) ;
   If (CritEdt.Rupture=rCorresp) then
      BEGIN
      OkEnteteRup:=True ;
      AnoDebitRupt.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, TotRupt[3], CritEdt.AfficheSymbole) ;
      AnoCreditRupt.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, TotRupt[4], CritEdt.AfficheSymbole) ;
      AnoSoldeRupt.Caption:=PrintSolde(TotRupt[3], TotRupt[4], CritEdt.Decimale, CritEdt.Symbole, CritEdt.AfficheSymbole) ;

      CumDebitRupt.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, TotRupt[5], CritEdt.AfficheSymbole) ;
      CumCreditRupt.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, TotRupt[6], CritEdt.AfficheSymbole) ;
      CumSoldeRupt.Caption:=PrintSolde(TotRupt[5], TotRupt[6], CritEdt.Decimale, CritEdt.Symbole, CritEdt.AfficheSymbole) ;

      SumD:=Arrondi(TotRupt[5]+ TotRupt[1],CritEdt.Decimale) ;
      SumC:=Arrondi(TotRupt[6]+TotRupt[2],CritEdt.Decimale) ;
      TCodRupt.Top:=37 ; DebitRupt.Top:=37 ; CreditRupt.Top:=37 ; SoldeRupt.Top:=37 ;
      BRupt.Height:=54 ;
      END Else
      BEGIN
      TCodRupt.Top:=2 ; DebitRupt.Top:=2 ; CreditRupt.Top:=2 ; SoldeRupt.Top:=2;
      If (CritEdt.Rupture<>rlibres) Then BRupt.Height:=20 ;
      END;
   *)

   AnoRupt.Visible:=(CritEdt.Rupture in [rCorresp,rLibres]) ;
   AnoDebitRupt.Visible:=(CritEdt.Rupture in [rCorresp,rLibres]) ;
   AnoCreditRupt.Visible:=(CritEdt.Rupture in [rCorresp,rLibres]) ;
   AnoSoldeRupt.Visible:=(CritEdt.Rupture in [rCorresp,rLibres]) ;

   CumDebitRupt.Visible:=(CritEdt.Rupture in [rCorresp,rLibres]) ;
   CumCreditRupt.Visible:=(CritEdt.Rupture in [rCorresp,rLibres]) ;
   CumSoldeRupt.Visible:=(CritEdt.Rupture in [rCorresp,rLibres]) ;
   CumRupt.Visible:=(CritEdt.Rupture in [rCorresp,rLibres]) ;
   Case CritEdt.Rupture Of
     rCorresp :  BEGIN
                 OkEnteteRup:=True ;
                 AnoDebitRupt.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, TotRupt[3], CritEdt.AfficheSymbole) ;
                 AnoCreditRupt.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, TotRupt[4], CritEdt.AfficheSymbole) ;
                 AnoSoldeRupt.Caption:=PrintSolde(TotRupt[3], TotRupt[4], CritEdt.Decimale, CritEdt.Symbole, CritEdt.AfficheSymbole) ;

                 CumDebitRupt.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, TotRupt[5], CritEdt.AfficheSymbole) ;
                 CumCreditRupt.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, TotRupt[6], CritEdt.AfficheSymbole) ;
                 CumSoldeRupt.Caption:=PrintSolde(TotRupt[5], TotRupt[6], CritEdt.Decimale, CritEdt.Symbole, CritEdt.AfficheSymbole) ;

                 SumD:=Arrondi(TotRupt[5]+ TotRupt[1],CritEdt.Decimale) ;
                 SumC:=Arrondi(TotRupt[6]+TotRupt[2],CritEdt.Decimale) ;
                 TCodRupt.Top:=37 ; DebitRupt.Top:=37 ; CreditRupt.Top:=37 ; SoldeRupt.Top:=37 ;
                 BRupt.Height:=54 ;
                 END ;
     rlibres  :  BEGIN
                 OkEnteteRup:=True ;
                 AnoDebitRupt.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, TotRupt[3], CritEdt.AfficheSymbole) ;
                 AnoCreditRupt.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, TotRupt[4], CritEdt.AfficheSymbole) ;
                 AnoSoldeRupt.Caption:=PrintSolde(TotRupt[3], TotRupt[4], CritEdt.Decimale, CritEdt.Symbole, CritEdt.AfficheSymbole) ;

                 CumDebitRupt.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, TotRupt[1], CritEdt.AfficheSymbole) ;
                 CumCreditRupt.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, TotRupt[2], CritEdt.AfficheSymbole) ;
                 CumSoldeRupt.Caption:=PrintSolde(TotRupt[1], TotRupt[2], CritEdt.Decimale, CritEdt.Symbole, CritEdt.AfficheSymbole) ;

                 SumD:=Arrondi(TotRupt[3]+ TotRupt[1],CritEdt.Decimale) ;
                 SumC:=Arrondi(TotRupt[4]+TotRupt[2],CritEdt.Decimale) ;
                 (*
                 TCodRupt.Top:=2 ; DebitRupt.Top:=2 ; CreditRupt.Top:=2 ; SoldeRupt.Top:=2;
                 *)
                 (*
                 TCodRupt.Top:=37 ; DebitRupt.Top:=37 ; CreditRupt.Top:=37 ; SoldeRupt.Top:=37 ;
                 BRupt.Height:=54 ;
                 *)
                 END ;
      Else BEGIN
           TCodRupt.Top:=2 ; DebitRupt.Top:=2 ; CreditRupt.Top:=2 ; SoldeRupt.Top:=2;
           If (CritEdt.Rupture<>rlibres) Then BRupt.Height:=20 ;
           END;
     END ;
   DebitRupt.Caption:= AfficheMontant(CritEDT.FormatMontant, CritEDT.Symbole, SumD, CritEDT.AfficheSymbole) ;
   CreditRupt.Caption:=AfficheMontant(CritEDT.FormatMontant, CritEDT.Symbole, SumC, CritEDT.AfficheSymbole) ;
   SoldeRupt.Caption:=PrintSolde(SumD, SumC, CritEDT.Decimale, CritEDT.Symbole, CritEDT.AfficheSymbole) ;
   SautPageRupture:=CritEdt.SautPageRupt And (CritEdt.GL.RuptOnly=Avec) And (CritEdt.SautPage<>1) And SautPageRuptAFaire(CritEdt,BDetail,QuelleRupt) ;
   END ;
end;

procedure TFGdLivAuxL.BRuptBeforePrint(var PrintBand: Boolean;  var Quoi: string);
begin
  inherited;
PrintBand:=(CritEdt.Rupture<>rRien) ;
end;

procedure TFGdLivAuxL.BRuptAfterPrint(BandPrinted: Boolean);
begin
  inherited;
Fillchar(TotAncRupt,SizeOf(TotAncRupt),#0) ;
end;

procedure TFGdLivAuxL.TOPREPORTBeforePrint(var PrintBand: Boolean;  var Quoi: string);
begin
  inherited;
Titre1Rep.Caption:=Titre2Rep.Caption ;
TITREREPORTH.Caption:=TITREREPORTB.Caption ;
Report1Debit.Caption:=Report2Debit.Caption ;
Report1Credit.Caption:=Report2Credit.Caption ;
Report1Solde.Caption:=Report2Solde.Caption ;
end;

procedure TFGdLivAuxL.BOTTOMREPORTBeforePrint(var PrintBand: Boolean;  var Quoi: string);
Var D,C : Double ;
begin
  inherited;
Case QuelReport(CritEDT.GL.FormatPrint.Report,D,C) Of
  1 : BEGIN TITREREPORTB.Caption:=MsgBox.Mess[3] ; Titre2Rep.Caption:='' ; END ;
  2 : BEGIN TITREREPORTB.Caption:=MsgBox.Mess[4] ; Titre2Rep.Caption:=StReportAux ; END ;
  END ;
Report2Debit.Caption:=AfficheMontant(CritEDT.FormatMontant, CritEDT.Symbole, D, CritEDT.AfficheSymbole ) ;
Report2Credit.Caption:=AfficheMontant(CritEDT.FormatMontant, CritEDT.Symbole, C, CritEDT.AfficheSymbole ) ;
Report2Solde.Caption:=PrintSolde(D,C,CritEDT.Decimale,CritEDT.Symbole,CritEDT.AfficheSymbole) ;
end;

Procedure TFGdLivAuxL.CalculCumulAu(Aux : String ; var CumulAu : TabTot) ;
BEGIN
Fillchar(CumulAu,SizeOf(CumulAu),#0) ;
QQ.Close ; QQ.Params[0].AsString:=Aux ; QQ.Open ;
While Not QQ.Eof Do
  BEGIN
  (*
  CumulAu[1].TotDebit:=Arrondi(QQ.Fields[2].AsFloat-QQ.Fields[0].AsFloat,CritEdt.Decimale) ;
  CumulAu[1].TotCredit:=Arrondi(QQ.Fields[3].AsFloat-QQ.Fields[1].AsFloat,CritEdt.Decimale) ;
  *)
  If QQ.FindField('E_DATEPAQUETMAX').AsDAteTime>CritEdt.Date2 Then
    BEGIN
    CumulAu[1].TotDebit:=Arrondi(CumulAu[1].TotDebit+QQ.Fields[2].AsFloat,CritEdt.Decimale) ;
    CumulAu[1].TotCredit:=Arrondi(CumulAu[1].TotCredit+QQ.Fields[3].AsFloat,CritEdt.Decimale) ;
    END Else
    BEGIN
    CumulAu[1].TotDebit:=Arrondi(CumulAu[1].TotDebit+QQ.Fields[2].AsFloat-QQ.Fields[0].AsFloat,CritEdt.Decimale) ;
    CumulAu[1].TotCredit:=Arrondi(CumulAu[1].TotCredit+QQ.Fields[3].AsFloat-QQ.Fields[1].AsFloat,CritEdt.Decimale) ;
    END ;
  QQ.Next ;
  END ;
QQ.Close ;
END ;

procedure TFGdLivAuxL.BDetailBeforePrint(var PrintBand: Boolean;  var Quoi: string);
var CumulAu : TabTot ;
    D,C   : Double ; CptRupt : String ;
begin
  inherited;
T_AUXILIAIRE.Caption:=MsgBox.Mess[10]+' '+Qr1T_AUXILIAIRE.AsString+' '+Qr1T_LIBELLE.AsString ;
Fillchar(TotDesMvts,SizeOf(TotDesMvts),#0) ;
Fillchar(TotCptAux,SizeOf(TotCptAux),#0) ;
Fillchar(CumulAu,SizeOf(CumulAu),#0) ;
Case CritEdt.Rupture of
  rLibres    : if CritEdt.GL.OnlyCptAssocie then PrintBand:=DansRuptLibre(Q,fbAux,CritEdt.LibreCodes1, CritEdt.LibreCodes2,CritEdt.LibreTrie) ;
  rRuptures  : if CritEdt.GL.OnlyCptAssocie then PrintBand:=DansRupt(LRupt,Qr1T_AUXILIAIRE.AsString) ;
  rCorresp   : if CritEdt.GL.OnlyCptAssocie then
                 if CritEDT.GL.PlansCorresp=1 then PrintBand:=(Qr1T_CORRESP1.AsString<>'') Else
                 if CritEDT.GL.PlansCorresp=2 then PrintBand:=(Qr1T_CORRESP2.AsString<>'') ;
  End;
Affiche:=PrintBand ;
If Not PrintBand then Exit ;
if PrintBand then
   BEGIN
   StReportAux:=Qr1T_AUXILIAIRE.AsString ;
   InitReport([2],CritEDT.GL.FormatPrint.Report) ;
   Case CritEdt.SautPage of
     0 : BDetail.forceNewPage:=(Qr1T_SAUTPAGE.AsString='X') ;
     1 : BDetail.forceNewPage:=True ;
     2 : BDetail.forceNewPage:=False ;
    end ;
   CalculCumulAu(Qr1T_AUXILIAIRE.AsString,CumulAu) ;
   CumulAu[0].TotDebit:=0 ; CumulAu[0].TotCredit:=0 ;
   D:=CumulAu[0].TotDebit+CumulAu[1].TotDebit ;
   C:=CumulAu[0].TotCredit+CumulAu[1].TotCredit ;

   { A Nouveau }
//   AnvDebit.Caption:=MsgBox.Mess[9] ; AnvCredit.Caption:='' ; AnvSolde.Caption:='' ;
   CumDebit.Caption:=MsgBox.Mess[9] ; CumCredit.Caption:='' ;  CumSolde.Caption:='' ;
   { A Nouveau + Cumul }
   CumDebit.Caption:= AfficheMontant(CritEDT.FormatMontant, CritEDT.Symbole,D, CritEDT.AfficheSymbole) ;
   CumCredit.Caption:=AfficheMontant(CritEDT.FormatMontant, CritEDT.Symbole,C,CritEDT.AfficheSymbole) ;
   CumSolde.Caption:= PrintSolde(D,C,CritEDT.Decimale,CritEDT.Symbole,CritEDT.AfficheSymbole) ;
   { Totaux tiers et rupture }
   TotCptAux[1].TotDebit:= Arrondi(TotCptAux[1].TotDebit+D,CritEDT.Decimale) ;
   TotCptAux[1].TotCredit:=Arrondi(TotCptAux[1].TotCredit+C,CritEDT.Decimale) ;
   TotAnCRupt[1].TotDebit:= Arrondi(TotAnCRupt[1].TotDebit+D,CritEDT.Decimale) ;
   TotAnCRupt[1].TotCredit:=Arrondi(TotAnCRupt[1].TotCredit+C,CritEDT.Decimale) ;


   { Ajout (ANouveau + Cuml) dans le solde progr. }
   { Init du Cumul Progressif }
   Case CritEdt.SoldeProg of
     1 : BEGIN
                 Progressif(False,0,0) ; { Si Progressif  obligatoire alors initialise Solde Progressif }
                 Progressif(True,D,C) ;
                 END ;
     0 : if Qr1T_SOLDEPROGRESSIF.AsString='X' then
                    BEGIN
                    Progressif(False,0,0) ; { Si Progressif ou non alors seulement si demandé par le compte }
                    Progressif(True,D,C) ;
                    END ;
    end ;
   AddReport([1,2],CritEDT.GL.FormatPrint.Report,D,C,CritEDT.Decimale) ;
   Case CritEdt.Rupture of
     rLibres   : AddGroupLibre(LRupt,Q,fbAux,CritEdt.LibreTrie,[1,0,0,CumulAu[0].TotDebit,CumulAu[0].TotCredit,D,C]) ;
     rRuptures : AddRupt(Lrupt,Qr1T_AUXILIAIRE.AsString,[1,0,0,CumulAu[0].TotDebit,CumulAu[0].TotCredit,D,C]) ;
     rCorresp  : BEGIN
                 Case CritEDT.GL.PlansCorresp Of
                   1 : If Qr1T_CORRESP1.AsString<>'' Then CptRupt:=Qr1T_CORRESP1.AsString+Qr2E_AUXILIAIRE.AsString
                                                     Else CptRupt:='.'+Qr2E_AUXILIAIRE.AsString ;
                   2 : If Qr1T_CORRESP2.AsString<>'' Then CptRupt:=Qr1T_CORRESP2.AsString+Qr2E_AUXILIAIRE.AsString
                                                     Else CptRupt:='.'+Qr2E_AUXILIAIRE.AsString ;
                  End ;
                 AddRuptCorres(LRupt,CptRupt,[1,0,0,CumulAu[0].TotDebit,CumulAu[0].TotCredit,D,C]) ;
                 END ;
     End ;
   END ;
Quoi:=QuoiAux(1) ;
end;

procedure TFGdLivAuxL.QEcrAfterOpen(DataSet: TDataSet);
begin
  inherited;
//
QR2E_AUXILIAIRE:=TStringField(QEcr.FindField('E_AUXILIAIRE')) ;
Qr2E_JOURNAL:=TStringField(QEcr.FindField('E_JOURNAL')) ;
Qr2E_VALIDE:=TStringField(QEcr.FindField('E_VALIDE')) ;
Qr2E_QUALIFPIECE:=TStringField(QEcr.FindField('E_QUALIFPIECE')) ;
Qr2E_EXERCICE:=TStringField(QEcr.FindField('E_EXERCICE')) ;
Qr2E_NUMEROPIECE:=TIntegerField(QEcr.FindField('E_NUMEROPIECE')) ;
Qr2E_NUMLIGNE:=TIntegerField(QEcr.FindField('E_NUMLIGNE')) ;
Qr2E_NUMECHE:=TIntegerField(QEcr.FindField('E_NUMECHE')) ;
Qr2E_DateComptable:=TDateTimeField(QEcr.FindField('E_DATECOMPTABLE')) ;
Qr2E_DATEPAQUETMIN:=TDateTimeField(QEcr.FindField('E_DATEPAQUETMIN')) ;
Qr2E_DATEPAQUETMAX:=TDateTimeField(QEcr.FindField('E_DATEPAQUETMAX')) ;
QR2E_LETTRAGE:=TStringField(QEcr.FindField('E_LETTRAGE')) ;
QR2E_ETATLETTRAGE:=TStringField(QEcr.FindField('E_ETATLETTRAGE')) ;
Qr2E_DATECOMPTABLE.Tag:=1 ;
Qr2DEBIT:=TFloatField(QEcr.FindField('DEBIT')) ;
Qr2CREDIT:=TFloatField(QEcr.FindField('CREDIT')) ;
Qr2COUV:=TFloatField(QEcr.FindField('COUV')) ;
QR2MULTIECHE:=TIntegerField(QEcr.FindField('MULTIECHE')) ;
IsECC:=(FDevises.Value<>V_PGI.DevisePivot)and(FMontant.ITemIndex=1)and(QEcr.FindField('E_NATUREPIECE').AsString='ECC') ;
ChgMaskChamp(Qr2DEBIT , CritEDT.Decimale, CritEDT.AfficheSymbole, CritEDT.Symbole, False) ;
ChgMaskChamp(Qr2CREDIT, CritEDT.Decimale, CritEDT.AfficheSymbole, CritEDT.Symbole, False) ;
end;


(*
FPlanRupt.visible:=(Not FSansRupt.Checked) And (Not FJoker.Visible) ;
FPlanRupt.Enabled:=FPlanRupt.visible ;
TFPlanRupt.Visible:=FPlanRupt.visible ;

*)
procedure TFGdLivAuxL.QAfterOpen(DataSet: TDataSet);
begin
  inherited;
Qr1T_AUXILIAIRE     :=TStringField(Q.FindField('T_AUXILIAIRE'));
Qr1T_LIBELLE        :=TStringField(Q.FindField('T_LIBELLE'));
Qr1T_SAUTPAGE       :=TStringField(Q.FindField('T_SAUTPAGE')) ;
Qr1T_SOLDEPROGRESSIF:=TStringField(Q.FindField('T_SOLDEPROGRESSIF')) ;
Qr1T_TOTAUXMENSUELS :=TStringField(Q.FindField('T_TOTAUXMENSUELS')) ;
If CritEDT.Rupture=rCorresp then
   BEGIN
   Qr1T_CORRESP1         :=TStringField(Q.FindField('T_CORRESP1'));
   Qr1T_CORRESP2         :=TStringField(Q.FindField('T_CORRESP2'));
   END ;
end;

procedure TFGdLivAuxL.BFCompteAuxAfterPrint(BandPrinted: Boolean);
begin
  inherited;
InitReport([2],CritEDT.GL.FormatPrint.Report) ;
end;

procedure TFGdLivAuxL.DLRuptHNeedData(var MoreData: Boolean);
begin
  inherited;
(* En-tête de rupture/corresp
MoreData:=False ;
if OkEnteteRup then
   If CritEdt.Rupture=rCorresp then
      BEGIN
      Case CritEdt.Bal.PlansCorresp of
        1 : BEGIN MoreData:=Qr1T_CORRESP1.AsString<>'' ;  END ;
        2 : BEGIN MoreData:=Qr1T_CORRESP2.AsString<>'' ;  END ;
        End ;
      END ;
OkEnteteRup:=False ;
*)
end;


procedure TFGdLivAuxL.BRupTeteBeforePrint(var PrintBand: Boolean; var Quoi: string);
begin
  inherited;
Case CritEdt.GL.PlansCorresp of
  1 : CodeRuptAu.Caption:=Qr1T_CORRESP1.AsString+'x' ;
  2 : CodeRuptAu.Caption:=Qr1T_CORRESP2.AsString+'x' ;
  End ;
end;

procedure TFGdLivAuxL.FSansRuptClick(Sender: TObject);
begin
  inherited;
FLigneRupt.Enabled:=Not FSansRupt.Checked ;
FLigneRupt.checked:=Not FSansRupt.Checked ;
FOnlyCptAssocie.Enabled:=Not FSansRupt.Checked ;
FOnlyCptAssocie.Checked:=Not FSansRupt.Checked ;
FRupturesClick(Nil) ;
end;

procedure TFGdLivAuxL.FRupturesClick(Sender: TObject);
begin
  inherited;
if FPlansCo.Checked then FGroupRuptures.Caption:=' '+MsgBox.Mess[14] ;
if FRuptures.Checked then If FPlanRuptures.Values.Count>0 Then FPlanRuptures.Value:=FPlanRuptures.Values[0] ;
end;

procedure TFGdLivAuxL.BDetailAfterPrint(BandPrinted: Boolean);
begin
  inherited;
If CritEdt.SautPageRupt And (CritEdt.GL.RuptOnly=Avec) And (CritEdt.SautPage<>1) Then
  BEGIN
  BDetail.ForceNewPage:=FALSE ; SautPageRupture:=FALSE ;
  END ;
end;

procedure TFGdLivAuxL.BDetailCheckForSpace;
begin
  inherited;
If SautPageRupture Then BDetail.ForceNewPage:=TRUE ;
end;

procedure TFGdLivAuxL.Timer1Timer(Sender: TObject);
begin
// GC - 20/12/2001
  inherited;
// GC - 20/12/2001 - FIN
end;

end.
