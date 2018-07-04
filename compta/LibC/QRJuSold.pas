unit QRJuSold;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  QR, hmsgbox, HQuickrp, DB,
  {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
  StdCtrls, Buttons, Hctrls, ExtCtrls,
  Mask, Hcompte, ComCtrls, UtilEdt, Ent1, HEnt1, HQry, Saisutil, CpteUtil,
  QRRupt,
{$IFNDEF CCMP}
  EdtLegal,
{$ENDIF}
  CritEDT, HSysMenu, Menus, HTB97, HPanel, UiUtil, tCalcCum ;

procedure JustSolde(Origine : TSuiviMP = smpAucun) ;
// GC - 20/12/2001
procedure JustSoldeChaine( vCritEdtChaine : TCritEdtChaine );
// GC - 20/12/2001 - FIN
procedure JustSoldeZoom(Crit : TCritEdt) ;

type
  TFJustSold = class(TFQR)
    TFCpt2: THLabel;
    FGen1: THCpteEdit;
    FGen2: THCpteEdit;
    TFJokerCpt2: THLabel;
    FGenJoker: TEdit;
    FLettrage: THValComboBox;
    TFLettrage: THLabel;
    FSituation: TCheckBox;
    HLabel11: THLabel;
    FRefInterne: TEdit;
    FSens: THValComboBox;
    TFSens: THLabel;
    FSautPage: TCheckBox;
    FLigneGenEntete: TCheckBox;
    FLigneGenPied: TCheckBox;
    FLigneTiePied: TCheckBox;
    FTriPar: TRadioGroup;
    RGen: TQRLabel;
    RGen1: TQRLabel;
    TRaG: TQRLabel;
    RGen2: TQRLabel;
    TRLettrage: TQRLabel;
    RLettrage: TQRLabel;
    TRLegende: TQRLabel;
    RLegende: TQRLabel;
    QRLabel10: TQRLabel;
    RRefInterne: TQRLabel;
    RSens: TQRLabel;
    TRSens: TQRLabel;
    TRSituation: TQRLabel;
    RSituation: TQRLabel;
    TE_VALIDE: TQRLabel;
    TLE_DATECOMPTABLE: TQRLabel;
    TE_JAL: TQRLabel;
    TE_PIECELIGNE: TQRLabel;
    TE_REFINTERNE: TQRLabel;
    TE_LIBELLE: TQRLabel;
    TE_MODEPAIE: TQRLabel;
    TE_DATEECHEANCE: TQRLabel;
    TLE_DEBIT: TQRLabel;
    TLE_CREDIT: TQRLabel;
    TLE_SOLDE: TQRLabel;
    TLE_LETTRAGE: TQRLabel;
    QRLabel17: TQRLabel;
    TITRE1REP: TQRLabel;
    REPORT1DEBIT: TQRLabel;
    REPORT1CREDIT: TQRLabel;
    REPORT1SOLDE: TQRLabel;
    TLT_AUXILIAIRE: TQRLabel;
    BSDHGene: TQRBand;
    TE_GENERAL: TQRLabel;
    BDetailEcr: TQRBand;
    LE_DATECOMPTABLE: TQRDBText;
    LE_CREDIT: TQRDBText;
    LE_DEBIT: TQRDBText;
    LE_LIBELLE: TQRDBText;
    LE_LETTRAGE: TQRDBText;
    LE_SOLDE: TQRLabel;
    LEVALIDE: TQRLabel;
    LE_JOURNAL: TQRDBText;
    LE_DATEECHEANCE: TQRDBText;
    LE_MODEPAIE: TQRDBText;
    LE_REFINTERNE: TQRDBText;
    LE_LETTRAGEDEV: TQRDBText;
    LE_PIECELIG: TQRLabel;
    BSDFGene: TQRBand;
    TOTE_GENERAL: TQRLabel;
    TotGenDebit: TQRLabel;
    TotGenCredit: TQRLabel;
    TotGenSolde: TQRLabel;
    BFCompteAux: TQRBand;
    TTOTE_AUXILIAIRE: TQRLabel;
    TotAuxSolde: TQRLabel;
    TotAuxCredit: TQRLabel;
    TotAuxDebit: TQRLabel;
    QRLabel33: TQRLabel;
    TotalDebit: TQRLabel;
    TotalCredit: TQRLabel;
    TotalSolde: TQRLabel;
    QRBand2: TQRBand;
    Trait0: TQRLigne;
    Trait1: TQRLigne;
    Trait2: TQRLigne;
    Trait3: TQRLigne;
    Trait4: TQRLigne;
    Trait5: TQRLigne;
    Ligne1: TQRLigne;
    TITRE2REP: TQRLabel;
    REPORT2DEBIT: TQRLabel;
    REPORT2CREDIT: TQRLabel;
    REPORT2SOLDE: TQRLabel;
    QEcr: TQuery;
    SEcr: TDataSource;
    QRDLFGene: TQRDetailLink;
    MsgLibel: THMsgBox;
    QRDLHGene: TQRDetailLink;
    QRDLCpt: TQRDetailLink;
    FNE: THValComboBox;
    HLabel10: THLabel;
    MsgBox: THMsgBox;
    TFaS: TLabel;
    BRupt: TQRBand;
    DebitLibre: TQRLabel;
    TCodRupt: TQRLabel;
    CreditLibre: TQRLabel;
    DLRupt: TQRDetailLink;
    SoldeLibre: TQRLabel;
    FLigneRupt: TCheckBox;
    FAfficheCol: TCheckBox;
    FSaufCptSolde: TCheckBox;
    FContrepartieSit: TCheckBox;
    procedure FormShow(Sender: TObject);
    procedure BDetailBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure BDetailEcrBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure BSDFGeneBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure BFCompteAuxBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure BFinEtatBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure QRDLFGeneNeedData(var MoreData: Boolean);
    procedure BSDHGeneBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure BFCompteAuxAfterPrint(BandPrinted: Boolean);
    procedure QRDLHGeneNeedData(var MoreData: Boolean);
    procedure TOPREPORTBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure BOTTOMREPORTBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure FNEChange(Sender: TObject);
    procedure QEcrAfterOpen(DataSet: TDataSet);
    procedure FSituationClick(Sender: TObject);
    procedure FNatureEcrChange(Sender: TObject);
    procedure BSDFGeneAfterPrint(BandPrinted: Boolean);
    procedure FGen1Change(Sender: TObject);
    procedure BNouvRechClick(Sender: TObject);
    procedure DLRuptNeedData(var MoreData: Boolean);
    procedure FRupTabLibreClick(Sender: TObject);
    procedure BRuptBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure QAfterOpen(DataSet: TDataSet);
    procedure FSansRuptClick(Sender: TObject);
    procedure FRupturesClick(Sender: TObject);
    procedure FOnlyCptAssocieClick(Sender: TObject);
    procedure FPlanRupturesChange(Sender: TObject);
    procedure BDetailAfterPrint(BandPrinted: Boolean);
    procedure BDetailCheckForSpace;
    procedure Timer1Timer(Sender: TObject);
    procedure GenereSQL ; Override ;
    procedure FinirPrint ; Override ;
    procedure InitDivers ; Override ;
    procedure RenseigneCritere ; Override ;
    procedure ChoixEdition ; Override ;
    function  CritOk : Boolean ; Override ;
    procedure RecupCritEdt ; Override ;
  private
    { Déclarations privées }
    LGen, LRupt                            : TStringList ;
    LibGen, CodGen,StReportAux,StReportGen : string ;
    Tot                                    : Array[0..12] of Double ;
    TotEdt, TotAux,TotGen              : TabTot ;
    OKEnteteGen, Affiche,SautPageRupture   : Boolean ;
    Origine : TSuiviMP ;
    Qr1AUXILIAIRE,Qr1LIBELLE,
    Qr1CORRESP1, Qr1CORRESP2               : TStringField ;
    Qr2G_LIBELLE                           : TStringField ;
    QR2E_AUXILIAIRE                        : TStringField ;
    QR2E_GENERAL                           : TStringField ;
    QR2E_JOURNAL                           : TStringField ;
    QR2E_NUMEROPIECE                       : TIntegerField ;
    QR2E_NUMLIGNE                          : TIntegerField ;
    Qr2DEBIT,Qr2CREDIT,Qr2COUVERTURE       : TFloatField ;
    QR2E_DATECOMPTABLE                     : TDateTimeField ;
    QR2E_DATEPAQUETMAX                     : TDateTimeField ;
    QR2E_LETTRAGE,QR2E_LETTRAGEDEV         : TStringField ;
    QR2E_EXERCICE                          : TStringField ;
    QR2E_VALIDE                            : TStringField ;
    QR2E_ETATLETTRAGE                      : TStringField ;
    Procedure Calculs ;
    procedure ContrePartie ;
    Function  QuoiGen(i : Integer) : String ;
    Function  QuoiAux(i : Integer) : String ;
    Function  QuoiMvt : String ;
    Procedure JustifZoom(Quoi : String) ;
    procedure SQLEcr ;
    procedure SqlCpt ;
    Procedure ActiverImprimeCollectif ;

  public
    { Déclarations publiques }
  end;

implementation

{$R *.DFM}

Uses HDEBUG ;

procedure JustSolde(Origine : TSuiviMP = smpAucun) ;
var QR : TFJustSold ;
    Edition : TEdition ;
    PP : THPanel ;
    NomFiltre : String ;
BEGIN
PP:=FindInsidePanel ;
QR:=TFJustSold.Create(Application) ;
Edition.Etat:=etJuSold ;
QR.QRP.QRPrinter.OnSynZoom:=QR.JustifZoom ;
QR.Origine:=Origine ;
{$IFDEF CCMP}
If Origine=smpEncTous Then NomFiltre:='QRJUSOLDETIEREC' Else NomFiltre:='QRJUSOLDETIERDE' ;
{$ELSE}
NomFiltre:='QRJUSOLDETIER' ;
{$ENDIF}
QR.InitType (nbAux,neJU,msRien,NomFiltre,'',TRUE,FALSE,FALSE,Edition) ;
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
procedure JustSoldeChaine( vCritEdtChaine : TCritEdtChaine );
var QR : TFJustSold ;
    Edition : TEdition ;
    PP : THPanel ;
    NomFiltre : String ;
BEGIN
PP:=FindInsidePanel ;
QR:=TFJustSold.Create(Application) ;
Edition.Etat:=etJuSold ;
QR.QRP.QRPrinter.OnSynZoom:=QR.JustifZoom ;
//QR.Origine:=Origine ;
NomFiltre:='QRJUSOLDETIER' ;
QR.CritEdtChaine := vCritEdtChaine;
QR.InitType (nbAux,neJU,msRien,NomFiltre,'',TRUE,FALSE,FALSE,Edition) ;
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

procedure JustSoldeZoom(Crit : TCritEdt) ;
var QR : TFJustSold ;
    Edition : TEdition ;
BEGIN
QR:=TFJustSold.Create(Application) ;
Edition.Etat:=etJuSold ;
try
 QR.QRP.QRPrinter.OnSynZoom:=QR.JustifZoom ;
 QR.CritEdt:=Crit ;
 QR.InitType (nbAux,neJU,msRien,'QRJUSOLDE','',FALSE,TRUE,FALSE,Edition) ;
 finally
 QR.Free ;
 end ;
Screen.Cursor:=SyncrDefault ;
END ;

Function TFJustSold.QuoiGen(i : Integer) : String ;
BEGIN
Inherited ;
Case i Of
  1 : Result:=QR2E_GENERAL.AsString+'#'+Qr2G_LIBELLE.AsString+'@'+'1' ;
  2 : Result:=QR2E_GENERAL.AsString+'#'+Qr2G_LIBELLE.AsString+'@'+'1' ;
  END ;
END ;

Function TFJustSold.QuoiAux(i : Integer) : String ;
Var si : String ;
BEGIN
Inherited ;
If CritEDT.JU.OnTiers Then si:='2' Else si:='1' ;
Case i Of
  1 : Result:=Qr1AUXILIAIRE.AsString+'#'+Qr1LIBELLE.AsString+'@'+si ;
  2 : Result:=Qr1AUXILIAIRE.AsString+'#'+Qr1LIBELLE.AsString+' '+TotAuxSolde.Caption+'@'+si ;
  END ;
END ;

Function TFJustSold.QuoiMvt : String ;
BEGIN
Inherited ;
If CritEDT.JU.OnTiers
Then Result:=QR2E_AUXILIAIRE.AsString+' '+Qr1LIBELLE.AsString+' '+Le_Solde.Caption+
             '#'+QR2E_JOURNAL.AsString+' N° '+IntToStr(QR2E_NUMEROPIECE.AsInteger)+' '+DateToStr(QR2E_DateComptable.AsDAteTime)+'-'+
              PrintSolde(Qr2DEBIT.AsFloat,Qr2Credit.AsFloat,CritEDT.Decimale,CritEDT.Symbole,CritEDT.AfficheSymbole)+' '+
              QR2E_LETTRAGE.AsString+
              '@'+'5;'+QR2E_JOURNAL.AsString+';'+UsDateTime(QR2E_DATECOMPTABLE.AsDateTime)+';'+QR2E_NUMEROPIECE.AsString+';'+QR2E_EXERCICE.asString+';'+
              IntToStr(QR2E_NumLigne.AsInteger)+';'
Else Result:=QR2E_GENERAL.AsString+' '+Qr1LIBELLE.AsString+' '+Le_Solde.Caption+
             '#'+QR2E_JOURNAL.AsString+' N° '+IntToStr(QR2E_NUMEROPIECE.AsInteger)+' '+DateToStr(QR2E_DateComptable.AsDAteTime)+'-'+
              PrintSolde(Qr2DEBIT.AsFloat,Qr2Credit.AsFloat,CritEDT.Decimale,CritEDT.Symbole,CritEDT.AfficheSymbole)+' '+
              QR2E_LETTRAGE.AsString+
              '@'+'5;'+QR2E_JOURNAL.AsString+';'+UsDateTime(QR2E_DATECOMPTABLE.AsDateTime)+';'+QR2E_NUMEROPIECE.AsString+';'+QR2E_EXERCICE.asString+';'+
              IntToStr(QR2E_NumLigne.AsInteger)+';' ;

END ;

Procedure TFJustSold.JustifZoom(Quoi : String) ;
Var Lp,i: Integer ;
BEGIN
Inherited ;
LP:=Pos('@',Quoi) ; If LP=0 Then Exit ;
i:=StrToInt(Copy(Quoi,LP+1,1)) ;
If (i=5) Then
   BEGIN
   Quoi:=Copy(Quoi,Lp+3,Length(Quoi)-lp-2) ;
   If QRP.QRPrinter.FSynShiftDblClick Then i:=6 ;
   If QRP.QRPrinter.FSynCtrlDblClick Then i:=11 ;
   END ;
ZoomEdt(i,Quoi) ;
END ;

procedure TFJustSold.finirPrint ;
{$IFNDEF CCMP}
var Solde : Double ;
{$ENDIF}
begin
Inherited ;
if CritEDT.JU.OnTiers Then
   BEGIN
   VideGroup(LGen) ;
   END ;
if (CritEdt.Rupture<>rRien) then VideRupt(LRupt) ;
QEcr.Close ;
{$IFNDEF CCMP}
if OkMajEdt Then
   BEGIN
   Solde:=TotEdt[1].TotDebit-TotEdt[1].TotCredit ;
   if Solde<0
      then MajEdition('JFS', '', DateToStr(CritEDT.Date1), DateToStr(CritEDT.Date1),
                      '', TotEdt[6].TotDebit, TotEdt[6].TotCredit, Solde, 0)
      else MajEdition('JFS', '', DateToStr(CritEDT.Date1), DateToStr(CritEDT.Date1),
                      '', TotEdt[6].TotDebit, TotEdt[6].TotCredit, 0, Solde) ;
   END ;
{$ENDIF}
end ;

procedure TFJustSold.InitDivers ;
BEGIN
Inherited ;
 { Gestion de l'apparition de l'entête de compte général }
OKEnteteGen:=True ;
BSDHGene.Frame.DrawTop:=CritEDT.JU.FormatPrint.PrSepCompte[2] ;
BSDFGene.Frame.DrawTop:=CritEDT.JU.FormatPrint.PrSepCompte[3] ;
BFCompteAux.Frame.DrawTop:=CritEDT.JU.FormatPrint.PrSepCompte[4] ;
BRupt.Frame.DrawTop:=CritEdt.JU.FormatPrint.PrSepCompte[5] ;
BRupt.Frame.DrawBottom:=BRupt.Frame.DrawTop ;
BDetail.ForceNewPage:=FSautPage.checked ;
If CritEdt.SautPageRupt And (CritEdt.JU.RuptOnly=Avec) And (CritEdt.SautPage<>1) Then BDetail.ForceNewPage:=FALSE ;
SautPageRupture:=FALSE ;
END ;

procedure TFJustSold.SqlCpt ;
Var Lequel,Pre,StPlus2,St,NatureRupt,P,St1 : String ;
{$IFDEF CCMP}
  QTiers : TQuery;
  szNature : String;
{$ENDIF}

// DGP : Clause Exist à valider (cf QDCJUSOLD)
begin
Q.Close ;
Q.SQL.Clear ;
If CritEDT.JU.OnTiers Then
   BEGIN
   Q.SQL.Add('Select T_AUXILIAIRE as CODE,T_LIBELLE as LIB, T_NATUREAUXI, T_TELEPHONE as TEL, T_ISPAYEUR as ISPAYEUR, T_PAYEUR AS PAYEUR ') ;
   Case CritEdt.Rupture of
     rLibres  : Q.SQL.Add(', T_TABLE0 , T_TABLE1, T_TABLE2, T_TABLE3, T_TABLE4, T_TABLE5, T_TABLE6, T_TABLE7, T_TABLE8, T_TABLE9 ') ;
     rCorresp : Q.SQL.Add(', T_CORRESP1 as CORRESPONDANCE1, T_CORRESP2 as CORRESPONDANCE2 ');
     End ;
{ Tables explorées par la SQL }
   Q.SQL.Add(' From TIERS T Where T_LETTRABLE="X" AND ') ;
   Lequel:='AUXILIAIRE' ; Pre:='T_' ;
   //if (OkV2 and (V_PGI.Confidentiel<>'1')) then Q.SQL.Add(' '+Pre+'CONFIDENTIEL<>"1" AND ') ;
   END Else
   BEGIN
   Q.SQL.Add('Select G_GENERAL as CODE, G_LIBELLE as LIB, G_NATUREGENE, G_TELEPHONE as TEL, G_TVASURENCAISS as ISPAYEUR, G_TVA AS PAYEUR ') ;
   Case CritEdt.Rupture of
     rLibres  : Q.SQL.Add(', G_TABLE0, G_TABLE1, G_TABLE2, G_TABLE3, G_TABLE4 as TABLE4, G_TABLE5, G_TABLE6, G_TABLE7, G_TABLE8, G_TABLE9 ') ;
     rCorresp : Q.SQL.Add(', G_CORRESP1 as CORRESPONDANCE1, G_CORRESP2 as CORRESPONDANCE2 ');
     End ;
   { Tables explorées par la SQL }
   Q.SQL.Add(' From GENERAUX T Where G_LETTRABLE="X" AND ') ;
   Lequel:='GENERAL' ; Pre:='G_' ;
   END ;
{ Construction de la clause Where de la SQL }
if (OkV2 and (V_PGI.Confidentiel<>'1')) then Q.SQL.Add(' '+Pre+'CONFIDENTIEL<>"1" AND ') ;
if CritEDT.DeviseSelect<>'' then
   BEGIN
   Q.SQL.Add(' (Exists (Select E_'+Lequel+',E_DATECOMPTABLE From ECRITURE Where E_'+Lequel+'=T.'+Pre+Lequel) ;
   Q.SQL.Add(' And E_DEVISE="'+CritEDT.DeviseSelect+'")) ') ;
   END else
   BEGIN
   (* GP le 19/06/97 pour synapses
   Q.SQL.Add(' '+Pre+Lequel+' <> "*" ') ;
   *)
   StPlus2:='' ;
   If CritEDT.JU.OnTiers Then
      BEGIN
      if CritEDT.SJoker then
         BEGIN
         StPlus2:=' And (E_GENERAL like "'+TraduitJoker(CritEDT.SCpt1)+'") ' ;
         END Else
         BEGIN
         if CritEDT.SCpt1<>'' then StPlus2:=StPlus2+' And (E_GENERAL>="'+CritEDT.SCpt1+'") ' ;
         if CritEDT.SCpt2<>'' then StPlus2:=StPlus2+' And (E_GENERAL<="'+CritEDT.SCpt2+'") ' ;
         END ;
      END ;
   St1:=WhereSuppEdtTiers(CritEdt.QualifPiece,V_PGI.Controleur,AvecRevision.State) ;
   Case CritEDT.JU.Lettrage of
      1 : BEGIN { Que les lettrés }
          if VH^.ExoV8.Code<>'' then
             BEGIN
             Q.SQL.Add(' (Exists (Select E_'+Lequel+',E_ETATLETTRAGE From ECRITURE Where E_'+Lequel+'=T.'+Pre+Lequel) ;
             If StPlus2<>'' Then Q.SQL.Add(StPlus2) ;
             Q.SQL.Add(' And (E_ETATLETTRAGE="TL" OR E_ETATLETTRAGE="PL") And (E_DATECOMPTABLE>="'+UsDateTime(VH^.ExoV8.Deb)+'"))) ') ;
             END Else
             BEGIN
             Q.SQL.Add(' (Exists (Select E_'+Lequel+',E_ETATLETTRAGE From ECRITURE Where E_'+Lequel+'=T.'+Pre+Lequel) ;
             If StPlus2<>'' Then Q.SQL.Add(StPlus2) ;
             Q.SQL.Add(' And (E_ETATLETTRAGE="TL" OR E_ETATLETTRAGE="PL"))) ') ;
             END ;
          END ;
      2 : BEGIN { Que le non lettrés }
          if VH^.ExoV8.Code<>'' then
             BEGIN
             Q.SQL.Add(' (Exists (Select E_'+Lequel+',E_ETATLETTRAGE From ECRITURE Where E_'+Lequel+'=T.'+Pre+Lequel) ;
             If StPlus2<>'' Then Q.SQL.Add(StPlus2) ;
             If St1<>'' Then Q.SQL.Add(St1) ;
             Q.SQL.Add(' And (E_ETATLETTRAGE="AL" OR E_ETATLETTRAGE="PL") And (E_DATECOMPTABLE>="'+UsDateTime(VH^.ExoV8.Deb)+'") AND (E_QUALIFPIECE<>"C" AND E_ECRANOUVEAU<>"OAN"))) ') ;
             END Else
             BEGIN
             Q.SQL.Add(' (Exists (Select E_'+Lequel+',E_ETATLETTRAGE From ECRITURE Where E_'+Lequel+'=T.'+Pre+Lequel) ;
             If StPlus2<>'' Then Q.SQL.Add(StPlus2) ;
             If St1<>'' Then Q.SQL.Add(St1) ;
             Q.SQL.Add(' And (E_ETATLETTRAGE="AL" OR E_ETATLETTRAGE="PL") And (E_QUALIFPIECE<>"C" AND E_ECRANOUVEAU<>"OAN"))) ') ;
             END ;
          END ;
      Else BEGIN
           if VH^.ExoV8.Code<>'' then
              BEGIN
              Q.SQL.Add(' (Exists (Select E_'+Lequel+',E_DATECOMPTABLE From ECRITURE Where E_'+Lequel+'=T.'+Pre+Lequel) ;
              If StPlus2<>'' Then Q.SQL.Add(StPlus2) ;
              If St1<>'' Then Q.SQL.Add(St1) ;
              Q.SQL.Add(' And E_DATECOMPTABLE>="'+UsDateTime(VH^.ExoV8.Deb)+'" )) ') ;
              END Else Q.SQL.Add(' '+Pre+Lequel+' <> "*" ') ;
           END ;
     End ;
   END ;
if CritEDT.Joker then Q.SQL.Add(' AND '+Pre+Lequel+' like "'+TraduitJoker(CritEDT.Cpt1)+'" ') Else
   BEGIN
   if CritEDT.Cpt1<>'' then Q.SQL.Add(' AND '+Pre+Lequel+'>="'+CritEDT.Cpt1+'" ') ;
   if CritEDT.Cpt2<>'' then Q.SQL.Add(' AND '+Pre+Lequel+'<="'+CritEDT.Cpt2+'" ') ;
   END ;
if CritEDT.NatureCpt<>'' then
   BEGIN
   If CritEDT.JU.OnTiers Then Q.SQL.Add(' AND T_NATUREAUXI="'+CritEDT.NatureCpt+'" ')
                         Else Q.SQL.Add(' AND G_NATUREGENE="'+CritEDT.NatureCpt+'" ') ;
   END Else
   BEGIN
{$IFDEF CCMP}
  QTiers := OpenSQL('SELECT T_NATUREAUXI FROM TIERS WHERE T_AUXILIAIRE="'+CritEDT.Cpt1+'"',True);
  szNature := QTiers.Fields[0].AsString;
  Ferme(QTiers);

  If (szNature='TID') Then Q.SQL.Add(' AND (G_NATUREGENE="TID") ');
  If (szNature='CLI') Then Q.SQL.Add(' AND ((T_NATUREAUXI="CLI") OR (T_NATUREAUXI="AUD")) ') ;

  If (szNature='TIC') Then Q.SQL.Add(' AND (G_NATUREGENE="TIC") ');
  If (szNature='FOU') Then Q.SQL.Add(' AND ((T_NATUREAUXI="FOU") OR (T_NATUREAUXI="AUC")) ') ;
{$ELSE}
   If Not CritEDT.JU.OnTiers Then Q.SQL.Add(' AND (G_NATUREGENE="TID" OR G_NATUREGENE="TIC")')
                             Else Q.SQL.Add(' AND T_NATUREAUXI<>"NCP" ') ;
{$ENDIF}
   END ;

If CritEDT.JU.OnTiers Then
   BEGIN
//   Q.SQL.Add(' AND (T_TOTALDEBIT<>0 or T_TOTALCREDIT<>0) ') ;
   If CritEdt.JU.SaufCptSolde Then Q.SQL.Add(' AND (T_TOTALDEBIT<>T_TOTALCREDIT) ') ;
   END Else
   BEGIN
//   Q.SQL.Add(' AND (G_TOTALDEBIT<>0 or G_TOTALCREDIT<>0) ') ;
   If CritEdt.JU.SaufCptSolde Then Q.SQL.Add(' AND (G_TOTALDEBIT<>G_TOTALCREDIT) ') ;
   END ;
If CritEdt.Rupture=rLibres then
   BEGIN
   St:='' ;
   If CritEDT.JU.OnTiers then St:=WhereLibre(CritEdt.LibreCodes1,CritEdt.LibreCodes2,fbAux,CritEdt.JU.OnlyCptAssocie)
                         Else St:=WhereLibre(CritEdt.LibreCodes1,CritEdt.LibreCodes2,fbGene,CritEdt.JU.OnlyCptAssocie) ;
   If St<>'' Then Q.SQL.Add(St) ;
   END ;
Case CritEdt.Rupture of
  rCorresp : BEGIN
//             if FOnlyCptAssocie.Checked then
             If CritEdt.JU.OnlyCptAssocie Then
                BEGIN
                If CritEdt.JU.OnTiers Then
                  BEGIN
                  NatureRupt:='AU'+IntToStr(CritEDT.JU.PlansCorresp) ;P:='T' ;
                  END Else
                  BEGIN
                  NatureRupt:='GE'+IntToStr(CritEDT.JU.PlansCorresp) ; P:='G' ;
                  END ;
                Q.SQL.Add(' AND Exists(SELECT CR_CORRESP FROM CORRESP WHERE CR_TYPE="'+NatureRupt+'"') ;
                Q.SQL.Add(' AND CR_CORRESP=T.'+P+'_CORRESP'+IntToStr(CritEDT.JU.PlansCorresp)+')') ;
                END ;
              END ;
  End ;

{ Construction de la clause Order By de la SQL }
If CritEdt.Rupture=rLibres then
   BEGIN
   Q.SQL.Add('Order By '+OrderLibre(CritEdt.LibreTrie)+' '+Pre+LEquel+' ') ;
   END Else
If CritEdt.Rupture=rCorresp then
   BEGIN
   Case CritEDT.JU.PlansCorresp Of
     1 : Q.Sql.Add('Order By '+Pre+'CORRESP1, '+Pre+Lequel+' ')  ;
     2 : Q.Sql.Add('Order By '+Pre+'CORRESP2, '+Pre+Lequel+' ') ;
      Else Q.Sql.Add('Order By '+Pre+Lequel+' ') ;
     End ;
   END Else
   BEGIN
   Case CritEDT.JU.TriePar of
     0 : BEGIN
         If CritEDT.JU.OnTiers Then Q.SQL.Add(' Order By T_NATUREAUXI, T_AUXILIAIRE')
                               Else Q.SQL.Add(' Order By G_NATUREGENE, G_GENERAL') ;
         END ;
     1 : BEGIN
         If CritEDT.JU.OnTiers Then Q.SQL.Add(' Order By T_NATUREAUXI, T_LIBELLE')
                               Else Q.SQL.Add(' Order By G_NATUREGENE, G_LIBELLE') ;
         END ;
     End ;
   END ;
ChangeSQL(Q) ; Q.Open ;
end ;

procedure TFJustSold.GenereSQL ;
{ Construction de la requête SQL en fonction du multicritère }
BEGIN
{ Construction de la clause Select de la SQL }
Inherited ;
SQLCpt ;
SQLECR ;
END ;

procedure TFJustSold.SQLEcr ;
{ Construction de la requête SQL en fonction du multicritère }
BEGIN
Inherited ;
{ Construction de la clause Select de la SQL }
QEcr.Close ;
QEcr.SQL.Clear ;
QEcr.SQL.Add('Select E_AUXILIAIRE,E_GENERAL, E_DATECOMPTABLE, E_NUMEROPIECE, E_NUMLIGNE,E_VALIDE, ') ;
QEcr.SQL.Add(       'E_REFINTERNE, E_ETABLISSEMENT, E_LIBELLE, E_ETATLETTRAGE, E_JOURNAL, ') ;
QEcr.SQL.Add(       'E_LETTRAGEDEV, E_LETTRAGE, E_DATEECHEANCE, E_MODEPAIE, E_DATEPAQUETMAX, E_EXERCICE, E_NUMECHE, ') ;
If CritEDT.JU.OnTiers Then QEcr.SQL.Add('G_LIBELLE as GLIB, ') ;
Case CritEDT.Monnaie of
  0 : BEGIN QEcr.SQL.Add(' E_DEBIT DEBIT, E_CREDIT CREDIT, E_COUVERTURE COUVERTURE ')         ; END ;
  1 : BEGIN QEcr.SQL.Add(' E_DEBITDEV DEBIT, E_CREDITDEV CREDIT, E_COUVERTUREDEV COUVERTURE '); END ;
end ;
{ Tables explorées par la SQL }
QEcr.SQL.Add(' From ECRITURE') ;

{ Construction de la clause Left Join de la SQL }
If CritEDT.JU.OnTiers Then QEcr.SQL.Add(' Left Join GENERAUX on G_GENERAL=E_GENERAL ') ;
{ Construction de la clause Where de la SQL }
If CritEDT.JU.OnTiers Then QEcr.SQL.Add('Where E_AUXILIAIRE=:CODE ')
                  Else QEcr.SQL.Add('Where E_GENERAL=:CODE ') ;
If CritEDT.JU.OnTiers Then
   BEGIN
   if CritEDT.SJoker then
      BEGIN
      QEcr.SQL.Add(' And E_GENERAL like "'+TraduitJoker(CritEDT.SCpt1)+'" ' )
      END Else
      BEGIN
      if CritEDT.SCpt1<>'' then QEcr.SQL.Add(' And E_GENERAL>="'+CritEDT.SCpt1+'" ') ;
      if CritEDT.SCpt2<>'' then QEcr.SQL.Add(' And E_GENERAL<="'+CritEDT.SCpt2+'" ') ;
      END ;
   END ;
QEcr.SQL.Add(' And (E_DATECOMPTABLE<="'+USDateTime(CritEDT.Date1)+'" Or E_DATEPAQUETMIN<="'+USDateTime(CritEDT.Date1)+'") ') ;
if VH^.ExoV8.Code<>'' then QEcr.SQL.Add(' And E_DATECOMPTABLE>="'+UsDateTime(VH^.ExoV8.Deb)+'" ') ;
QEcr.SQL.Add(TraduitNatureEcr(CritEDT.Qualifpiece, 'E_QUALIFPIECE', TRUE, CritEdt.ModeRevision)) ;
if CritEDT.RefInterne<>'' then QEcr.SQL.Add(' And UPPER(E_REFINTERNE) like "'+TraduitJoker(CritEDT.RefInterne)+'" ' );
Case CritEDT.JU.Lettrage of
   1 : QEcr.SQL.Add(' And E_ETATLETTRAGE="TL" ') ;
   2 : QEcr.SQL.Add(' And E_ETATLETTRAGE<>"TL" ') ;
  End ;
QEcr.SQL.Add(' And E_ETATLETTRAGE<>"RI" ') ;
if CritEDT.Etab<>'' then QEcr.SQL.Add(' And E_ETABLISSEMENT="'+CritEDT.Etab+'"') ;
if CritEDT.DeviseSelect<>'' then QEcr.SQL.Add(' And E_DEVISE="'+CritEDT.DeviseSelect+'"') ;
QEcr.SQL.Add(' And E_ECRANOUVEAU<>"CLO" And E_ECRANOUVEAU<>"OAN" and E_QUALIFPIECE<>"C" ') ;
If CritEDT.SQLPlus<>'' Then QEcr.SQL.Add(CritEDT.SQLPlus) ;
Case CritEDT.JU.Sens of
  0 : BEGIN
      Case CritEDT.Monnaie of
        0 : BEGIN QEcr.SQL.Add(' And E_CREDIT<>0 ' ) ;    END ;
        1 : BEGIN QEcr.SQL.Add(' And E_CREDITDEV<>0 ') ;  END ;
      end ;
      END ;
  1 : BEGIN
      Case CritEDT.Monnaie of
        0 : BEGIN QEcr.SQL.Add(' And E_DEBIT<>0 ' ) ;    END ;
        1 : BEGIN QEcr.SQL.Add(' And E_DEBITDEV<>0 ') ;  END ;
      end ;
      END ;
 end ;
{ Construction de la clause Order By de la SQL }
QEcr.SQL.Add(' Order By E_AUXILIAIRE, E_GENERAL, E_ETATLETTRAGE DESC, E_LETTRAGE, E_DATECOMPTABLE, E_DATEECHEANCE, E_NUMECHE, E_NUMEROPIECE ') ;
ChangeSQL(QEcr) ; QEcr.Open ;
END ;

procedure TFJustSold.RenseigneCritere ;
BEGIN
Inherited ;
TE_VALIDE.Caption:='' ;
RGen.Visible:=CritEDT.JU.OnTiers ;
RGen1.Visible:=CritEDT.JU.OnTiers ; RGen2.Visible:=CritEDT.JU.OnTiers ;
TRaG.Visible:=CritEDT.JU.OnTiers ;
If CritEDT.JU.OnTiers Then
   BEGIN
   RCpte.Caption:=MsgLibel.Mess[11] ; RJoker.Caption:=MsgLibel.Mess[12] ;
   if CritEdt.SJoker then
      BEGIN
      RGen.Caption:=MsgLibel.Mess[7] ;
      RGen1.Caption:=FGenJoker.Text ;
      END Else
      BEGIN
      RGen.Caption:=MsgLibel.Mess[6] ;
      RGen1.Caption:=FGen1.Text ; RGen2.Caption:=FGen2.Text ;
      END ;
      RGen2.Visible:=Not CritEdt.SJoker ; TRaG.Visible:=Not CritEdt.SJoker ;
   END ELSE
   BEGIN
   RCpte.Caption:=MsgLibel.Mess[6] ; RJoker.Caption:=MsgLibel.Mess[7] ;
   END ;
RSens.Caption:=FSens.Text ;
RRefInterne.Caption:=FRefInterne.Text ;
RLettrage.Caption:=FLettrage.Text ;
CaseACocher(FSituation, RSituation) ;
END ;

procedure TFJustSold.ChoixEdition ;
{ Initialisation des options d'édition }
Var Pref : char ;
    StType : String ;
BEGIN
Inherited ;
If CritEDT.JU.OnTiers Then BEGIN ChargeGroup(LGen,['']) ; Pref:='T' ; StType:='RUT' ; END
                      Else BEGIN Pref:='G' ; StType:='RUG' ; END ;
DLRupt.PrintBefore:=TRUE ;
Case CritEdt.Rupture of
  rLibres   : BEGIN
              DLRupt.PrintBefore:=FALSE ;
              ChargeGroup(LRupt,[Pref+'00',Pref+'01',Pref+'02',Pref+'03',Pref+'04',Pref+'05',Pref+'06',Pref+'07',Pref+'08',Pref+'09']) ;
              END ;
  rRuptures : BEGIN
              ChargeRupt(LRupt, StType, CritEdt.JU.PlanRupt, CritEdt.JU.CodeRupt1,CritEdt.JU.CodeRupt2) ;
              NiveauRupt(LRupt);
              END ;
  rCorresp  : BEGIN
              ChargeRuptCorresp(LRupt, CritEdt.JU.PlanRupt, CritEdt.JU.CodeRupt1, CritEdt.JU.CodeRupt2, False) ;
              NiveauRupt(LRupt);
              END ;
  End ;

TRSituation.Visible:=FSituation.checked ;  RSituation.Visible:=FSituation.checked ;
ChgMaskChamp(Qr2DEBIT,CritEDT.Decimale,CritEDT.AfficheSymbole,CritEDT.Symbole,False) ;
ChgMaskChamp(Qr2CREDIT,CritEDT.Decimale,CritEDT.AfficheSymbole,CritEDT.Symbole,False) ;
END ;

procedure TFJustSold.RecupCritEdt ;
Var NonLibres : Boolean ;
BEGIN
Inherited ;
With CritEDT Do
  BEGIN
  JU.CodeRupt1:='' ; JU.CodeRupt2:='' ;
  SautPage:=2 ;
  If FSautPage.Checked Then SautPage:=1 ;
  RefInterne:=FRefInterne.Text ;
  JU.TriePar:=FTriPar.ItemIndex ;
  JU.Lettrage:=FLettrage.ItemIndex ;
  JU.Sens:=FSens.ItemIndex ;
  JU.EnSituation:=FSituation.Checked ;
  If JU.EnSituation Then JU.SansContrepartie:=Not FContrepartieSit.Checked Else JU.SansContrepartie:=FALSE ;
  JU.OnTiers:=(FNE.Value = 'TIE');
  If JU.OnTiers Then
     BEGIN
     SJoker:=FGenJoker.Visible ;
     if SJoker Then BEGIN SCpt1:=FGenJoker.Text ; SCpt2:=FGenJoker.Text ; END
               Else BEGIN SCpt1:=FGen1.Text     ; SCpt2:=FGen2.Text ;     END ;
     END ;
  JU.RuptOnly:=QuelleTypeRupt(0,FSAnsRupt.Checked,FAvecRupt.Checked,FSurRupt.Checked) ;
  NonLibres:=((Rupture=rRuptures) or (Rupture=rCorresp)) ;
  If NonLibres Then JU.PlanRupt:=FPlanRuptures.Value ;
  If NonLibres Then BEGIN JU.CodeRupt1:=FCodeRupt1.Text ; JU.CodeRupt2:=FCodeRupt2.Text ; END ;
  JU.OnlyCptAssocie:=(Rupture<>rRien) and FOnlyCptAssocie.Checked ;
  If (CritEdt.Rupture=rCorresp) Then JU.PlansCorresp:=FPlanRuptures.ItemIndex+1 ;
  Ju.PasDeCollectifAAfficher:=FALSE ;
  If (CritEdt.Rupture=rLibres) And (JU.OnlyCptAssocie) And ((JU.RuptOnly=Avec) Or (JU.RuptOnly=Sur)) Then
     Ju.PasDeCollectifAAfficher:= Not FAfficheCol.Checked ;
  JU.SaufCptSolde:=FSaufCptSolde.Checked ;
  With JU.FormatPrint Do
       BEGIN
       If CritEDT.JU.OnTiers Then
          BEGIN
          PrSepCompte[2]:=FLigneGenEntete.Checked ;
          PrSepCompte[3]:=FLigneGenPied.Checked ;
          END ;
       PrSepCompte[4]:=FLigneTiePied.Checked ;
       PrSepCompte[5]:=FLigneRupt.Checked ;
       END ;
  END ;
END ;

function  TFJustSold.CritOK : Boolean ;
BEGIN
Result:=Inherited CritOk ;
If Result Then Fillchar(TotEdt,SizeOf(TotEdt),#0) ;
END ;

procedure TFJustSold.Calculs ;
{ Lignes de mouvement standards et incrémentation des variables de calculs }
Var MReport : TabTrep ;
    LeFb    : TfichierBase ;
    CptRupt : String ;
BEGIN
Inherited ;
Fillchar(MReport,SizeOf(MReport),#0) ;
TotGen[1].TotDebit:=     Arrondi(TotGen[1].TotDebit+Qr2DEBIT.AsFloat,CritEDT.Decimale) ;
TotGen[1].TotCredit:=    Arrondi(TotGen[1].TotCredit+Qr2CREDIT.AsFloat,CritEDT.Decimale) ;
TotAux[1].TotDebit:=     Arrondi(TotAux[1].TotDebit+Qr2DEBIT.AsFloat,CritEDT.Decimale) ;
TotAux[1].TotCredit:=    Arrondi(TotAux[1].TotCredit+Qr2CREDIT.AsFloat,CritEDT.Decimale) ;
TotEdt[1].TotDebit:= Arrondi(TotEdt[1].TotDebit+Qr2DEBIT.AsFloat,CritEDT.Decimale) ;
TotEdt[1].TotCredit:=Arrondi(TotEdt[1].TotCredit+Qr2CREDIT.AsFloat,CritEDT.Decimale) ;
if Qr2DEBIT.AsFloat=0 then
   BEGIN
   LE_SOLDE.Caption:=PrintSolde(Qr2COUVERTURE.AsFloat,Qr2CREDIT.AsFloat, CritEDT.Decimale,CritEDT.Symbole,CritEDT.AfficheSymbole) ;
   MReport[2].TotDebit:=Qr2COUVERTURE.AsFloat ;
   MReport[2].TotCredit:=0 ;
   END Else
   BEGIN
   LE_SOLDE.Caption:=PrintSolde(Qr2DEBIT.AsFloat,Qr2COUVERTURE.AsFloat, CritEDT.Decimale,CritEDT.Symbole,CritEDT.AfficheSymbole) ;
   MReport[2].TotCredit:=Qr2COUVERTURE.AsFloat ;
   MReport[2].TotDebit:=0 ;
   END ;
LE_LIBELLE.Font.color:=clBlack ; LE_LIBELLE.Font.Style:=[] ;
{ Pour le calcul des Montant en C et D pour le Report }
MReport[1].TotDebit:= Arrondi(MReport[1].TotDebit+Qr2DEBIT.AsFloat,CritEDT.Decimale) ;
MReport[1].TotCredit:=Arrondi(MReport[1].TotCredit+Qr2CREDIT.AsFloat,CritEDT.Decimale) ;
If CritEDT.JU.OnTiers Then AddReportBAL([1,2,3],CritEDT.JU.FormatPrint.Report,MReport,CritEDT.Decimale)
                      Else AddReportBAL([1,2],CritEDT.JU.FormatPrint.Report,MReport,CritEDT.Decimale) ;


if CritEDT.JU.OnTiers then LeFb:=fbAux else LeFb:=fbGene ;
Case CritEdt.Rupture of
  rLibres   : AddGroupLibre(LRupt,Q,LeFb,CritEdt.LibreTrie,[1,Qr2DEBIT.AsFloat,Qr2CREDIT.AsFloat,0,0,0]) ;
  rRuptures : AddRupt(LRupt,Qr1AUXILIAIRE.AsString,[1,Qr2DEBIT.AsFloat,Qr2CREDIT.AsFloat,0,0,0]) ;
  rCorresp  : BEGIN
              Case CritEDT.JU.PlansCorresp Of
                1 : If Qr1CORRESP1.AsString<>'' Then CptRupt:=Qr1CORRESP1.AsString+Qr1AUXILIAIRE.AsString
                                                  Else CptRupt:='.'+Qr1AUXILIAIRE.AsString ;
                2 : If Qr1CORRESP2.AsString<>'' Then CptRupt:=Qr1CORRESP2.AsString+Qr1AUXILIAIRE.AsString
                                                  Else CptRupt:='.'+Qr1AUXILIAIRE.AsString ;
                Else CptRupt:=Qr1AUXILIAIRE.AsString ;
                End ;
              AddRuptCorres(LRupt,CptRupt,[1,Qr2DEBIT.AsFloat,Qr2CREDIT.AsFloat,0,0,0]) ;
              END ;
  End ;


END ;

procedure TFJustSold.ContrePartie ;
{ Lignes de mouvements spéciales pour les contreparties en situation }
Var MReport : TabTrep ;
BEGIN
Inherited ;
Fillchar(MReport,SizeOf(MReport),#0) ;
LE_DEBIT.Caption:='' ; LE_CREDIT.Caption:='' ; LE_LETTRAGE.Caption:='' ; LE_LETTRAGEDEV.Caption:='' ;
LE_DATECOMPTABLE.Caption:='' ; LE_JOURNAL.Caption:='' ; LE_PIECELIG.Caption:='' ;
LE_REFINTERNE.Caption:='' ; LE_MODEPAIE.Caption:='' ;
LE_LIBELLE.Caption:=MsgLibel.Mess[2]+' '+ DateToStr(CritEDT.Date1) ;
LE_LIBELLE.Font.color:=clRed ; LE_LIBELLE.Font.Style:=[fsItalic] ; LEVALIDE.Caption:='' ;
LE_SOLDE.Caption:=PrintSolde(Qr2CREDIT.AsFloat,Qr2DEBIT.AsFloat, CritEDT.Decimale,CritEDT.Symbole,CritEDT.AfficheSymbole) ;
{ Calcul des Montants en D et C pour le Report... }
MReport[1].TotDebit:= Arrondi(MReport[1].TotDebit+0,CritEDT.Decimale) ;
MReport[1].TotCredit:=Arrondi(MReport[1].TotCredit+0,CritEDT.Decimale) ;
{ Calcul de la Contrepartie pour le Report... }
MReport[3].TotDebit:= Arrondi(MReport[3].TotDebit+Qr2DEBIT.AsFloat,CritEDT.Decimale) ;
MReport[3].TotCredit:=Arrondi(MReport[3].TotCredit+Qr2CREDIT.AsFloat,CritEDT.Decimale) ;
If CritEDT.JU.OnTiers Then AddReportBal([1,2,3],CritEDT.JU.FormatPrint.Report,MReport,CritEDT.Decimale)
                      Else AddReportBal([1,2],CritEDT.JU.FormatPrint.Report,MReport,CritEDT.Decimale) ;
END ;

procedure TFJustSold.FormShow(Sender: TObject);
begin
HelpContext:=7535000 ;
//Standards.HelpContext:=7535010 ;
//Avances.HelpContext:=7535020 ;
//Mise.HelpContext:=7535030 ;
//Option.HelpContext:=7535040 ;
//TabRuptures.HelpContext:=7535050 ;

FGen1.Text:='' ; FGen2.Text:='' ;
FLettrage.Value:='NL' ;
Fsens.ItemIndex:=2 ;
FDateCompta1.Text:=DateToStr(V_PGI.DateEntree) ;
FNE.ItemIndex:=1 ;
  inherited;
TabSup.TabVisible:=False;
If FFiltres.Text='' then FDateCompta1.Text:=DateToStr(V_PGI.DateEntree) ;
AvecRevision.Visible:=False ;
(*
FSurRupt.Visible:=False ;
FCodeRupt1.Visible:=False  ; FCodeRupt2.Visible:=False ;
TFCodeRupt1.Visible:=False ; TFCodeRupt2.Visible:=False ;
FOnlyCptAssocie.Enabled:=False ;
*)
//FLigneRupt.Enabled:=False ;

{$IFDEF CCMP}
// BPY le 10/09/2004 : Fiche n° 13956 => Edition des TIC TID dans CCMP !
//FNE.Enabled:=FALSE ; FNE.Value:='TIE' ;
////FnatureCpt.Enabled:=FALSE ;
FNatureCpt.Vide := False;
FNEChange(FNE);
//if (VH^.CCMP.LotCli) then begin FNatureCpt.Plus := ' AND (CO_CODE="AUD" OR CO_CODE="CLI" OR CO_CODE="DIV")'; FNatureCpt.Value:='CLI'; end
//                     else begin FNatureCpt.Plus := ' AND (CO_CODE="AUC" OR CO_CODE="DIV" OR CO_CODE="FOU" OR CO_CODE="SAL")'; FNatureCpt.Value:='FOU'; end;
// Fin BPY
{$ENDIF}

// GC - 20/12/2001
if CritEdtChaine.Utiliser then InitEtatchaine('QRJUSOLDETIER');
// GC - 20/12/2001 - FIN
end;

procedure TFJustSold.BDetailBeforePrint(var PrintBand: Boolean;  var Quoi: string);
Var St,StPlus,StLib : String ;
    QAux : TQuery ;
begin
  inherited;
(* Rony 2/04/97
StReportAux:=LT_AUXILIAIRE.Caption ;
*)
OKEnteteGen:=True ; StPlus:='' ;
If CritEdt.Ju.OnTiers And VH^.OuiTP Then
  BEGIN
  If Q.FindField('ISPAYEUR').AsString='X' then StPlus:=StPlus+'   (Tiers payeur)' Else
    BEGIN
    St:=Q.FindField('PAYEUR').AsString ; StLib:='' ;
    If St<>'' then
      BEGIN
      QAux:=OpenSQL('SELECT T_ABREGE FROM TIERS WHERE T_AUXILIAIRE="'+St+'" ',TRUE) ;
      If Not QAux.Eof Then StLib:=QAux.Fields[0].AsString Else StLib:=TraduireMemoire('Inexistant') ;
      Ferme(QAux) ;
      StPlus:=StPlus+'   '+'(Tiers payeur : '+StLib+')' ;
      END ;
    END ;
  END ;
TLT_AUXILIAIRE.Caption:=MsgLibel.Mess[8]+'   '+Q.FindField('CODE').AsString+'   '+Q.FindField('LIB').AsString +'   '+Q.FindField('TEL').AsString+StPlus;
InitReport([2],CritEDT.JU.FormatPrint.Report) ;
StReportAux:=Q.FindField('CODE').AsString ;
Fillchar(TotAux,SizeOf(TotAux),#0) ;  { Init du total Aux }
Case CritEdt.Rupture of
  rLibres    : if CritEdt.JU.OnlyCptAssocie then PrintBand:=DansRuptLibre(Q,fbAux,CritEdt.LibreCodes1, CritEdt.LibreCodes2,CritEdt.LibreTrie) ;
  rRuptures  : if CritEdt.JU.OnlyCptAssocie then PrintBand:=DansRupt(LRupt,Qr1AUXILIAIRE.AsString) ;
  rCorresp   : if CritEdt.JU.OnlyCptAssocie then
                 BEGIN
                 if CritEDT.JU.PlansCorresp=1 then PrintBand:=DansRuptCorresp(LRupt,Qr1CORRESP1.AsString) Else
                 if CritEDT.JU.PlansCorresp=2 then PrintBand:=DansRuptCorresp(LRupt,Qr1CORRESP2.AsString) ;
                 END ;
  (*
                 if CritEDT.JU.PlansCorresp=1 then PrintBand:=(Qr1CORRESP1.AsString<>'') Else
                 if CritEDT.JU.PlansCorresp=2 then PrintBand:=(Qr1CORRESP2.AsString<>'') ;
  *)
  End;
Affiche:=PrintBand ;
if PrintBand then Quoi:=QuoiAux(1) ;
end;

procedure TFJustSold.BDetailEcrBeforePrint(var PrintBand: Boolean;  var Quoi: string);
begin
  inherited;
PrintBand:=Affiche And (Not QR2E_GENERAL.IsNull) ;
If Not PrintBand then Exit ; //Rony 30/10/97
LE_PIECELIG.Caption:=QR2E_NUMEROPIECE.AsString+'/'+QR2E_NUMLIGNE.AsString ;
if QR2E_VALIDE.AsString='X' then LEVALIDE.Caption:='V' else LEVALIDE.Caption:=' ' ;
if QR2E_LETTRAGEDEV.AsString='-' then LE_LETTRAGEDEV.Caption:='' ;
if CritEDT.JU.EnSituation then   { En situation }
   BEGIN
   if (QR2E_ETATLETTRAGE.asString='TL') then  { Lettrage total }
      BEGIN
      if QR2E_DATEPAQUETMAX.AsDateTime<=CritEDT.Date1
         then BEGIN PrintBand:=False ; END
         else BEGIN
              If QR2E_DATECOMPTABLE.AsDateTime>CritEDT.Date1
                 then BEGIN ContrePartie ; If CritEdt.JU.SansContrepartie Then PrintBand:=FALSE ; END
                 else BEGIN Calculs ; END ;
              END ;
      END else
      BEGIN        { Lettrage partiel ou inexistant }
      if QR2E_DATECOMPTABLE.AsDateTime>CritEDT.Date1
         then BEGIN ContrePartie ; If CritEdt.JU.SansContrepartie Then PrintBand:=FALSE ; END
         else BEGIN Calculs ; END ;
      END ;
   END else   { On est en instantané }
   BEGIN
   Calculs ;
   END ;
If CritEDT.JU.OnTiers Then AddGroup(LGen,[QR2E_GENERAL],[1]) ;
Quoi:=QuoiMvt ;
If CritEdt.Ju.RuptOnly=Sur then PrintBand:=False ;
end;

procedure TFJustSold.BSDFGeneBeforePrint(var PrintBand: Boolean;  var Quoi: string);
begin
  inherited;
If Not CritEDT.JU.OnTiers Then BEGIN PrintBand:=FALSE ; Exit ; END ;
PrintBand:=Affiche And (Not QR2E_GENERAL.IsNull) ;
TOTE_GENERAL.Caption:=MsgLibel.Mess[10]+'   '+QR2E_GENERAL.AsString+'   '+Qr2G_LIBELLE.AsString ;
TOTE_GENERAL.Width:=TE_LIBELLE.Width+TE_MODEPAIE.Width+TE_DATEECHEANCE.Width+2 ;
If PrintBand then
   BEGIN
   TotGenDebit.Caption:= AfficheMontant(CritEDT.FormatMontant, CritEDT.Symbole,TotGen[1].TotDebit, CritEDT.AfficheSymbole) ;
   TotGenCredit.Caption:=AfficheMontant(CritEDT.FormatMontant, CritEDT.Symbole,TotGen[1].TotCredit,CritEDT.AfficheSymbole) ;
   TotGenSolde.Caption:=PrintSolde(TotGen[1].TotDebit,TotGen[1].TotCredit, CritEDT.Decimale,CritEDT.Symbole,CritEDT.AfficheSymbole) ;
   Quoi:=QuoiGen(2) ;
   END ;
If CritEdt.JU.PasDeCollectifAAfficher Then PrintBand:=FALSE ;
end;

procedure TFJustSold.BFCompteAuxBeforePrint(var PrintBand: Boolean;  var Quoi: string);
begin
  inherited;
PrintBand:=Affiche ;
If PrintBand then
   BEGIN
   TTOTE_AUXILIAIRE.Caption:=MsgLibel.Mess[10]+'   '+
                             Q.FindField('CODE').AsString+'   '+Q.FindField('LIB').AsString;
   TTOTE_AUXILIAIRE.Left:=TE_REFINTERNE.Left ;
   TTOTE_AUXILIAIRE.Width:=TE_REFINTERNE.Width+TE_LIBELLE.Width+TE_MODEPAIE.Width+TE_DATEECHEANCE.Width+3 ;
   TotAuxDebit.Caption:=AfficheMontant(CritEDT.FormatMontant, CritEDT.Symbole, TotAux[1].TotDebit, CritEDT.AfficheSymbole ) ;
   TotAuxCredit.Caption:=AfficheMontant(CritEDT.FormatMontant, CritEDT.Symbole, TotAux[1].TotCredit, CritEDT.AfficheSymbole ) ;
   TotAuxSolde.Caption:=PrintSolde(TotAux[1].TotDebit, TotAux[1].TotCredit ,CritEDT.Decimale,CritEDT.Symbole,CritEDT.AfficheSymbole) ;
   Quoi:=QuoiAux(2) ;
   END ;
end;

procedure TFJustSold.BFinEtatBeforePrint(var PrintBand: Boolean;  var Quoi: string);
begin
  inherited;
TotalDebit.Caption:=AfficheMontant(CritEDT.FormatMontant, CritEDT.Symbole, TotEdt[1].TotDebit, CritEDT.AfficheSymbole ) ;
TotalCredit.Caption:=AfficheMontant(CritEDT.FormatMontant, CritEDT.Symbole, TotEdt[1].TotCredit, CritEDT.AfficheSymbole ) ;
TotalSolde.Caption:=PrintSolde(TotEdt[1].TotDebit,TotEdt[1].TotCredit,CritEDT.Decimale,CritEDT.Symbole,CritEDT.AfficheSymbole) ;
BOTTOMREPORT.enabled:=FALSE ;
end;

procedure TFJustSold.QRDLFGeneNeedData(var MoreData: Boolean);
Var QuelleRupt: integer ;
begin
  inherited;
If Not CritEDT.JU.OnTiers Then Exit ;
MoreData:=PrintGroup(LGen, QEcr, [QR2E_GENERAL], CodGen, LibGen, Tot,Quellerupt) ;
if ((MoreData) and (Not QEcr.EOF)) then OKEnteteGen:=True
end;

procedure TFJustSold.BSDHGeneBeforePrint(var PrintBand: Boolean;  var Quoi: string);
begin
  inherited;
Fillchar(TotGen,SizeOf(TotGen),#0) ;  { Init du total Gen }
If Not CritEDT.JU.OnTiers Then BEGIN PrintBand:=FALSE ; Exit ; END ;
PrintBand:=Affiche And (Not QR2E_GENERAL.IsNull) ;
TE_GENERAL.Caption:=MsgLibel.Mess[9]+'   '+QR2E_GENERAL.AsString+'   '+Qr2G_LIBELLE.AsString ;
InitReport([3],CritEDT.JU.FormatPrint.Report) ;
StReportGen:=QR2E_GENERAL.AsString ;
if PrintBand then Quoi:=QuoiGen(1) ;
If CritEdt.JU.PasDeCollectifAAfficher Then PrintBand:=FALSE ;
end;

procedure TFJustSold.BFCompteAuxAfterPrint(BandPrinted: Boolean);
begin
  inherited;
//OKEnteteGen:=True ;
InitReport([2],CritEDT.JU.FormatPrint.Report) ;
end;

procedure TFJustSold.QRDLHGeneNeedData(var MoreData: Boolean);
begin
  inherited;
MoreData:=OKEnteteGen ;
OKEnteteGen:=False ;
end;

procedure TFJustSold.TOPREPORTBeforePrint(var PrintBand: Boolean;  var Quoi: string);
begin
  inherited;
TITREREPORTH.Caption:=TITREREPORTB.Caption ;
TITRE1REP.Caption:=TITRE2REP.Caption ;
Report1Debit.Caption:=Report2Debit.Caption ;
Report1Credit.Caption:=Report2Credit.Caption ;
Report1Solde.Caption:=Report2Solde.Caption ;

end;

procedure TFJustSold.BOTTOMREPORTBeforePrint(var PrintBand: Boolean;  var Quoi: string);
Var MReport        : TabTrep ;
    SoldeD, SoldeC : Double ;
begin
  inherited;
Fillchar(MReport,SizeOf(MReport),#0) ;
Case QuelReportBAL(CritEDT.JU.FormatPrint.Report,MReport) Of
  1 : BEGIN TITREREPORTB.Caption:=MsgLibel.Mess[3] ; Titre2Rep.Caption:='' ;END ;
  2 : BEGIN TITREREPORTB.Caption:=MsgLibel.Mess[4] ; Titre2Rep.Caption:=StReportAux ; END ;
  3 : BEGIN TITREREPORTB.Caption:=MsgLibel.Mess[5] ; Titre2Rep.Caption:=StReportGen ; END ;
  END ;
SoldeD:=MReport[1].TotDebit+MReport[2].TotDebit+MReport[3].TotCredit ;
SoldeC:=MReport[1].TotCredit+MReport[2].TotCredit+MReport[3].TotDebit ;
Report2Debit.Caption:=AfficheMontant(CritEDT.FormatMontant, CritEDT.Symbole, MReport[1].TotDebit, CritEDT.AfficheSymbole ) ;
Report2Credit.Caption:=AfficheMontant(CritEDT.FormatMontant, CritEDT.Symbole,MReport[1].TotCredit, CritEDT.AfficheSymbole ) ;
Report2Solde.Caption:=PrintSolde(SoldeD,SoldeC,CritEDT.Decimale,CritEDT.Symbole,CritEDT.AfficheSymbole) ;
end;

procedure TFJustSold.FNEChange(Sender: TObject);
Var
    OnTiers : Boolean ;
    Edition : TEdition ;
    NomFiltre : String ;
begin
    inherited;

    //BPY le 10/09/2004 => sinon en cas de trie des tablette .... ca devien n'importe quoi !
    OnTiers:=(FNE.Value = 'TIE');

    TFJokerCpt2.Visible := OnTiers;
    TFCpt2.Visible := OnTiers;
    FGen1.Visible := OnTiers;
    FGen2.Visible := OnTiers;
    FGenJoker.Visible := OnTiers;
    TFaS.Visible := OnTiers;

    FLigneGenEntete.Visible := OnTiers;
    FLigneGenPied.Visible := OnTiers;

    Edition.Etat := etJuSold;

{$IFDEF CCMP}
    If Origine=smpEncTous Then NomFiltre:='QRJUSOLDETIEREC' Else NomFiltre:='QRJUSOLDETIERDE' ;
{$ELSE}
    NomFiltre:='QRJUSOLDETIER' ;
{$ENDIF}

    if (OnTiers) then
    begin
        TFGenJoker.Caption := MsgBox.Mess[2];
        TFGen.Caption := MsgBox.Mess[0];
        initType(nbAux,neJU,msRien,NomFiltre,'',TRUE,FALSE,FALSE,Edition);
        If FGenJoker.text = '' Then FGenJoker.Visible := FALSE;
        FPlanRuptures.DataType := 'ttRuptTiers';
    end
    else
    begin
        TFGenJoker.Caption := MsgBox.Mess[3];
        TFGen.Caption := MsgBox.Mess[1];
        initType(nbGenT,neJU,msRien,NomFiltre,'',TRUE,FALSE,FALSE,Edition) ;
        FGen1.Text := '';
        FGen2.Text := '';
        FGenJoker.text := '';
        FPlanRuptures.DataType := 'ttRuptGene';
    end;
    GeneOuTiers;

// BPY le 10/09/2004 : Fiche n° 13956 : edition des TIC TID dans CCMP
{$IFDEF CCMP}
    FNatureCpt.Vide := false;

    if (OnTiers) then           // sur les tiers ...
    begin
        FNatureCpt.enabled := true;
        if (origine = smpEncTous) then FNatureCpt.Datatype := 'TTNATTIERSCPTAENC'   // encaissement
        else FNatureCpt.Datatype := 'TTNATTIERSCPTADEC';                            // decaissement
        FNatureCptChange(FNatureCpt);
    end
    else                        // sur les generaux ...
    begin
        if (origine = smpEncTous) then FNatureCpt.Value := 'TID'    // encaissement
        else FNatureCpt.Value := 'TIC';                             // decaissement
        FNatureCptChange(FNatureCpt);
        FNatureCpt.enabled := false;
    end;
{$ENDIF}
// Fin BPY

    if (Not QRLoading) Then
    BEGIN
        FJoker.Text := '';
        FCpte1.Text := '';
        FCpte2.Text := '';
        if (FNatureCpt.enabled) then FNatureCpt.ItemIndex := 0;
    end;

    if FTablesLibres.Checked then FLibTriPar.Text := '';

    FSansRuptClick(Nil);
end;

procedure TFJustSold.QEcrAfterOpen(DataSet: TDataSet);
begin
  inherited;
If CritEDT.JU.OnTiers Then Qr2G_LIBELLE :=TStringField(QEcr.FindField('GLIB')) ;
QR2E_AUXILIAIRE    :=TStringField(QEcr.FindField('E_AUXILIAIRE')) ;
QR2E_GENERAL       :=TStringField(QEcr.FindField('E_GENERAL')) ;
QR2E_JOURNAL       :=TStringField(QEcr.FindField('E_JOURNAL')) ;
QR2E_NUMEROPIECE   :=TIntegerField(QEcr.FindField('E_NUMEROPIECE')) ;
QR2E_NUMLIGNE      :=TIntegerField(QEcr.FindField('E_NUMLIGNE')) ;
Qr2DEBIT           :=TFloatField(QEcr.FindField('DEBIT')) ;
Qr2CREDIT          :=TFloatField(QEcr.FindField('CREDIT')) ;
Qr2COUVERTURE      :=TFloatField(QEcr.FindField('COUVERTURE')) ;
QR2E_DATECOMPTABLE :=TDAteTimeField(Qecr.FindField('E_DATECOMPTABLE')) ;
QR2E_DATEPAQUETMAX :=TDAteTimeField(Qecr.FindField('E_DATEPAQUETMAX')) ;
QR2E_LETTRAGE      :=TStringField(QEcr.FindField('E_LETTRAGE')) ;
QR2E_LETTRAGEDEV   :=TStringField(QEcr.FindField('E_LETTRAGEDEV')) ;
QR2E_EXERCICE      :=TStringField(QEcr.FindField('E_EXERCICE')) ;
QR2E_VALIDE        :=TStringField(QEcr.FindField('E_VALIDE')) ;
QR2E_ETATLETTRAGE  :=TStringField(QEcr.FindField('E_ETATLETTRAGE')) ;
ChgMaskChamp(Qr2DEBIT,CritEDT.Decimale,CritEDT.AfficheSymbole,CritEDT.Symbole,False) ;
ChgMaskChamp(Qr2CREDIT,CritEDT.Decimale,CritEDT.AfficheSymbole,CritEDT.Symbole,False) ;
end;

procedure TFJustSold.FSituationClick(Sender: TObject);
begin
  inherited;
If FSituation.Checked then
   BEGIN
   FLettrage.ItemIndex:=0 ;
   FSaufCptSolde.Checked:=FALSE ; FSaufCptSolde.Enabled:=FALSE ;
   END Else BEGIN FSaufCptSolde.Enabled:=TRUE ; END ;
FLettrage.Enabled:=Not FSituation.Checked ;
FContrepartieSit.Enabled:=FSituation.Checked ;
If Not QRLoading Then FContrepartieSit.Checked:=TRUE ;
end;

procedure TFJustSold.FNatureEcrChange(Sender: TObject);
begin
  inherited;
if (FNatureEcr.Value='PRE')or(FNatureEcr.Value='SSI') then
   BEGIN
   FSituation.State:=cbUnchecked ; FSituation.Enabled:=False ;
   FLettrage.Value:='NL' ; FLettrage.Enabled:=False ;
   END else
   BEGIN
   FLettrage.Enabled:=True ; FSituation.Enabled:=True ;
   END ;
end;

procedure TFJustSold.BSDFGeneAfterPrint(BandPrinted: Boolean);
begin
  inherited;
InitReport([3],CritEDT.JU.FormatPrint.Report) ;
end;

procedure TFJustSold.FGen1Change(Sender: TObject);
Var AvecJokerG : Boolean ;
begin
  inherited;
If FNE.ItemIndex=0 Then
   BEGIN
   //TFaS.Visible:=FALSE ; rony 06/06/97
   // ??? Rony 27/05/97 ???
   //TFGen.Visible:=FALSE ;
   //TFGenJoker.Visible:=FALSE ;
   Exit ;
   END ;
AvecJokerG:=Joker(FGen1, FGen2, FGenJoker) ;
TFaS.Visible:=Not AvecJokerG ;
TFCpt2.Visible:=Not AvecJokerG ;
TFJokerCpt2.Visible:=AvecJokerG ;
end;

procedure TFJustSold.BNouvRechClick(Sender: TObject);
begin
  inherited;
If FGenJoker.Visible then FGenJoker.Text:='' ;

end;

procedure TFJustSold.DLRuptNeedData(var MoreData: Boolean);
var TotRupt              : Array[0..12] of Double ;
    Librupt, CodRupt, Lib1, CptRupt, Stcode : String ;
    Quellerupt           : Integer ;
    LeFb                 : TfichierBase ;
    Col                  : TColor ;
    OkOk, DansTotal      : Boolean ;
    LibRuptInf : Array[1..10] Of TRuptInf ;
begin
  inherited;
//SautPageRupture:=FALSE ;
MoreData:=false ; QuelleRupt:=0 ;
if CritEDT.JU.OnTiers then LeFb:=fbAux else LeFb:=fbGene ;
Case CritEdt.Rupture of
  rLibres   : BEGIN
              MoreData:=PrintGroupLibre(LRupt,Q,LeFb,CritEdt.LibreTrie,CodRupt,LibRupt,Lib1,TotRupt,Quellerupt,Col,LibRuptInf) ;
              If MoreData then
                 BEGIN
                 StCode:=CodRupt ;
                 Delete(StCode,1,Quellerupt+2) ;
                 MoreData:=DansChoixCodeLibre(StCode,Q,fbAux,CritEdt.LibreCodes1,CritEdt.LibreCodes2, CritEdt.LibreTrie) ;
                 BRupt.Font.Color:=Col ;
                 END ;
              END ;
  rRuptures : MoreData:=PrintRupt(LRupt,Qr1AUXILIAIRE.AsString,CodRupt,LibRupt,DansTotal,QRP.EnRupture,TotRupt) ;
  rCorresp  : BEGIN
              OkOk:=TRUE ;
              Case CritEDT.JU.PlansCorresp  Of
                1 : If Qr1CORRESP1.AsString<>'' Then CptRupt:=Qr1CORRESP1.AsString+Qr1AUXILIAIRE.AsString
                                                Else CptRupt:='.'+Qr1AUXILIAIRE.AsString ;
                2 : If Qr1CORRESP2.AsString<>'' Then CptRupt:=Qr1CORRESP2.AsString+Qr1AUXILIAIRE.AsString
                                                Else CptRupt:='.'+Qr1AUXILIAIRE.AsString ;
                Else OkOk:=FALSE ;
                END ;
              If OkOk Then MoreData:=PrintRupt(LRupt,CptRupt,CodRupt,LibRupt,DansTotal,QRP.EnRupture,TotRupt) Else MoreData:=FALSE ;
              END ;
  End ;
if MoreData then
   BEGIN
   TCodRupt.Width:=TE_REFINTERNE.Width+TE_LIBELLE.Width+1 ;
   TCodRupt.Caption:='' ;
   BRupt.Height:=HauteurBandeRuptIni ;
   if CritEdt.Rupture=rLibres then
      BEGIN
      insert(MsgLibel.Mess[14]+' ',CodRupt,Quellerupt+2) ;
      TCodRupt.Caption:=CodRupt+' '+Lib1 ;
      AlimRuptSup(HauteurBandeRuptIni,QuelleRupt,TCodRupt.Width,BRupt,LibRuptInf,Self) ;
      END Else TCodRupt.Caption:=CodRupt+'   '+LibRupt ;
   DebitLibre.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, TotRupt[1], CritEdt.AfficheSymbole) ;
   CreditLibre.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, TotRupt[2], CritEdt.AfficheSymbole) ;
   SoldeLibre.Caption:=PrintSolde(TotRupt[1],TotRupt[2],CritEDT.Decimale,CritEDT.Symbole,CritEDT.AfficheSymbole) ;
   SautPageRupture:=CritEdt.SautPageRupt And (CritEdt.JU.RuptOnly=Avec) And (CritEdt.SautPage<>1) And SautPageRuptAFaire(CritEdt,BDetail,QuelleRupt) ;
   END ;
end;

procedure TFJustSold.FRupTabLibreClick(Sender: TObject);
begin
  inherited;
// CCF
//FTriPar.Enabled:=Not FTablesLibres.Checked ;
end;

procedure TFJustSold.BRuptBeforePrint(var PrintBand: Boolean; var Quoi: string);
begin
  inherited;
PrintBand:=(CritEdt.Rupture<>rRien) ;
end;

procedure TFJustSold.QAfterOpen(DataSet: TDataSet);
begin
  inherited;
Qr1AUXILIAIRE:=TStringField(Q.FindField('CODE')) ;
Qr1LIBELLE:=TStringField(Q.FindField('LIB')) ;
If CritEDT.Rupture=rCorresp then
   BEGIN
   Qr1CORRESP1         :=TStringField(Q.FindField('CORRESPONDANCE1'));
   Qr1CORRESP2         :=TStringField(Q.FindField('CORRESPONDANCE2'));
   END ;
end;

procedure TFJustSold.FSansRuptClick(Sender: TObject);
begin
  inherited;
FLigneRupt.Enabled:=Not FSansRupt.Checked ;
FLigneRupt.checked:=Not FSansRupt.Checked ;
FOnlyCptAssocie.Enabled:=Not FSansRupt.Checked ;
FOnlyCptAssocie.Checked:=Not FSansRupt.Checked ;
FRupturesClick(Nil) ;
ActiverImprimeCollectif ;
end;

procedure TFJustSold.FRupturesClick(Sender: TObject);
begin
  inherited;
If FPlansCo.Checked then
   BEGIN
   FGroupRuptures.Caption:=' '+MsgLibel.Mess[17] ;
   CorrespToCombo(FPlanRuptures,fbAux) ;
   If FPlanRuptures.Values.Count>0 Then FPlanRuptures.Value:=FPlanRuptures.Values[0] ;
   END Else
If FRuptures.Checked then
   BEGIN
   FPlanRuptures.Reload    ;
   FGroupRuptures.Caption:=' '+MsgLibel.Mess[16] ;
   If FPlanRuptures.Values.Count>0 Then FPlanRuptures.Value:=FPlanRuptures.Values[0] ;
   END ;
ActiverImprimeCollectif ;
end;

Procedure TFJustSold.ActiverImprimeCollectif ;
BEGIN
If (FAvecRupt.Checked Or FSurRupt.Checked) And FTablesLibres.Checked And FOnlyCptAssocie.Checked Then FAfficheCol.Visible:=TRUE Else
   BEGIN
   FAfficheCol.Visible:=FALSE ; FAfficheCol.Checked:=TRUE ;
   END ;
END ;

procedure TFJustSold.FOnlyCptAssocieClick(Sender: TObject);
begin
  inherited;
ActiverImprimeCollectif ;
end;

procedure TFJustSold.FPlanRupturesChange(Sender: TObject);
begin
  inherited;
If QRLoading Then Exit ;
if FPlansCo.Checked then CorrespToCodes(FPlanRuptures,FCodeRupt1,FCodeRupt2) Else
if FRuptures.Checked then RuptureToCodes(FPlanRuptures,FCodeRupt1,FCodeRupt2,FbAux) ;
FCodeRupt1.ItemIndex:=0 ; FCodeRupt2.ItemIndex:=FCodeRupt2.Items.Count-1 ;
end;

procedure TFJustSold.BDetailAfterPrint(BandPrinted: Boolean);
begin
  inherited;
If CritEdt.SautPageRupt And (CritEdt.JU.RuptOnly=Avec) And (CritEdt.SautPage<>1) Then
  BEGIN
  BDetail.ForceNewPage:=FALSE ; SautPageRupture:=FALSE ;
  END ;
end;

procedure TFJustSold.BDetailCheckForSpace;
begin
  inherited;
If SautPageRupture Then BDetail.ForceNewPage:=TRUE ;
end;

procedure TFJustSold.Timer1Timer(Sender: TObject);
begin
// GC - 20/12/2001
  inherited;
// GC - 20/12/2001 - FIN
end;

end.

