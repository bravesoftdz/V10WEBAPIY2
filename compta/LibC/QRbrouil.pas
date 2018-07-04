unit QRBrouil;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  HQuickRP, StdCtrls, ExtCtrls, DB,
  {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}
  Buttons, Mask, Hctrls,
  ComCtrls, HQry, HMsgBox, Ent1, HCompte, HEnt1, ParamDat,
  ULibExercice, QRRupt, SaisUtil, UtilEdt, Grids, Printers, Filtre, CritEDT,
  Menus, HSysMenu, HTB97, HPanel, UiUtil, tCalcCum, UObjFiltres;

procedure Brouillard(Code : char) ;

Const OkV2 : Boolean = FALSE ;

type
  TQRBrouillard = class(TForm)
    SEcr                  : TDataSource;
    QRP                   : TQuickReport;
    GPiece                : TQRGroup;
    Panel1                : TPanel;
    Brouillard            : TQRBand;
    BDetail               : TQRBand;
    E_NUMLIGNE            : TQRDBText;
    E_DEBIT               : TQRDBText;
    QAna : TQuery;
    SAna                : TDataSource;
    QDAna                 : TQRDetailLink;
    Pages                 : TPageControl;
    Standards             : TTabSheet;
    Avances               : TTabSheet;
    FDateCompta2          : TMaskEdit;
    Label7                : TLabel;
    FDateCompta1          : TMaskEdit;
    TFaG: TLabel;
    FRefInterne: TEdit;
    FDevises              : THValComboBox;
    TTitre                : TQRSysData;
    BEntetePiece          : TQRBand;
    TE_JOURNAL            : TQRLabel;
    TE_NUMEROPIECE        : TQRLabel;
    E_CREDIT              : TQRDBText;
    G_ABREGE              : TQRDBText;
    T_ABREGE              : TQRDBText;
    E_LIBELLE             : TQRDBText;
    E_MODEPAIE            : TQRDBText;
    E_DATEECHEANCE        : TQRDBText;
    BAnalytique           : TQRBand;
    Y_DEBIT               : TQRDBText;
    Y_CREDIT              : TQRDBText;
    Y_LIBELLE             : TQRDBText;
    QRLabel3              : TQRLabel;
    RCpte: TQRLabel;
    QRLabel6              : TQRLabel;
    QRLabel8              : TQRLabel;
    QRLabel10             : TQRLabel;
    QRLabel11             : TQRLabel;
    QRLabel15             : TQRLabel;
    QRLabel17             : TQRLabel;
    RDateCompta1          : TQRLabel;
    RDateCompta2          : TQRLabel;
    RCpte1: TQRLabel;
    RCpte2: TQRLabel;
    RTri                  : TQRLabel;
    RRefInterne: TQRLabel;
    REtab: TQRLabel;
    RValide               : TQRLabel;
    QRLabel18             : TQRLabel;
    RDevises              : TQRLabel;
    S_ABREGE              : TQRDBText;
    BPiedPiece            : TQRBand;
    QEcr: TQuery;
    BFinEtat              : TQRBand;
    TE_NUMEROPIECE_: TQRLabel;
    QRLabel33             : TQRLabel;
    E_REFINTERNE          : TQRDBText;
    Y_VALIDE: TQRLabel;
    MsgRien               : THMsgBox;
    Mise: TTabSheet;
    E_VALIDE: TQRLabel;
    FMontant: TRadioGroup;
    MsgListe              : THMsgBox;
    GroupBox1             : TGroupBox;
    FAfficheAnal          : TCheckBox;
    FAfficheQte           : TCheckBox;
    TRLegende      : TQRLabel;
    RLegende       : TQRLabel;
    FMonetaire     : TCheckBox;
    TFGen: THLabel;
    HLabel6        : THLabel;
    HLabel1        : THLabel;
    HLabel2        : THLabel;
    HLabel9        : THLabel;
    HLabel10       : THLabel;
    FExercice      : THValComboBox;
    FCpte1: THCpteEdit;
    FCpte2: THCpteEdit;
    FEtab          : THValComboBox;
    Formateur: THNumEdit;
    MsgBox: THMsgBox;
    QRBand1: TQRBand;
    Trait1: TQRLigne;
    Trait2: TQRLigne;
    Trait3: TQRLigne;
    Trait0: TQRLigne;
    Ligne1: TQRLigne;
    EnteteAutrePage: TQRBand;
    TitreEntete: TQRSysData;
    TitreBarre: TQRShape;
    EntetePage: TQRBand;
    TNumLigne: TQRLabel;
    TE_REFINTERNE: TQRLabel;
    TDebit: TQRLabel;
    TCredit: TQRLabel;
    TModePaiement: TQRLabel;
    TDateEcheance: TQRLabel;
    TE_LIBELLE: TQRLabel;
    TCptGen: TQRLabel;
    TLibAux: TQRLabel;
    TRien: TQRLabel;
    QRLabel20: TQRLabel;
    RExercice: TQRLabel;
    Bevel1: TBevel;
    TotDebitPiece: TQRLabel;
    TotCreditPiece: TQRLabel;
    TotDebitGen: TQRLabel;
    TotCreditGen: TQRLabel;
    MsgNatEcr: THMsgBox;
    TOPREPORT: TQRBand;
    REPORT1CREDIT: TQRLabel;
    REPORT1DEBIT: TQRLabel;
    TITRE1REP: TQRLabel;
    BOTTOMREPORT: TQRBand;
    REPORT2CREDIT: TQRLabel;
    REPORT2DEBIT: TQRLabel;
    TITRE2REP: TQRLabel;
    piedpage: TQRBand;
    RCopyright: TQRLabel;
    QRSysData1: TQRSysData;
    QRSysData2: TQRSysData;
    RSociete: TQRLabel;
    RUtilisateur: TQRLabel;
    FReport: TCheckBox;
    Option: TTabSheet;
    GroupBox2: TGroupBox;
    Bevel2: TBevel;
    FTrait: TCheckBox;
    Reduire: TCheckBox;
    FLignePieceEntete: TCheckBox;
    FLignePiecePied: TCheckBox;
    tabSup: TTabSheet;
    Apercu: TCheckBox;
    FListe: TCheckBox;
    FCouleur: TCheckBox;
    Y_NumLV: TQRLabel;
    QRLabel12: TQRLabel;
    RRefExt: TQRLabel;
    QRLabel21: TQRLabel;
    RRefLib: TQRLabel;
    QRLabel23: TQRLabel;
    RAffaire: TQRLabel;
    QRLabel26: TQRLabel;
    RDateRefExt: TQRLabel;
    Y_SECTION: TQRLabel;
    Y_REFINTERNE: TQRDBText;
    Y_POURCENTAGE: TQRLabel;
    TE_ETABLISSEMENT: TQRLabel;
    FValide: TCheckBox;
    FTri: TCheckBox;
    HLabel8: THLabel;
    FNumPiece1: TEdit;
    Label12: TLabel;
    FNumPiece2: TEdit;
    QRLabel13: TQRLabel;
    RNumPiece1: TQRLabel;
    TRa: TQRLabel;
    RNumPiece2: TQRLabel;
    LeCompte: TQRLabel;
    FJoker: TEdit;
    TFGenJoker: THLabel;
    RJoker: TQRLabel;
    E_LIBELLE2: TQRLabel;
    QRDLEche: TQRDetailLink;
    BEche: TQRBand;
    TOTAlEche: TQRLabel;
    TotDebitEche: TQRLabel;
    TotCreditEche: TQRLabel;
    TE_DEVISE: TQRLabel;
    FRefExt: TCheckBox;
    FRefLib: TCheckBox;
    FAffaire: TCheckBox;
    FDateRefExt: TCheckBox;
    E_REFINTERNE2: TQRLabel;
    TE_TAUXDEV: TQRLabel;
    QRLabel1: TQRLabel;
    QRLabel2: TQRLabel;
    QRLabel4: TQRLabel;
    QRLabel5: TQRLabel;
    FAvecJalAN: TCheckBox;
    FRappelCrit: TCheckBox;
    HMTrad: THSystemMenu;
    POPF: TPopupMenu;
    BCreerFiltre: TMenuItem;
    BSaveFiltre: TMenuItem;
    BDelFiltre: TMenuItem;
    BRenFiltre: TMenuItem;
    BNouvRech: TMenuItem;
    Dock971: TDock97;
    PanelFiltre: TToolWindow97;
    HPB: TToolWindow97;
    BFiltre: TToolbarButton97;
    FFiltres: THValComboBox;
    BParamListe: TToolbarButton97;
    BValider: TToolbarButton97;
    BFerme: TToolbarButton97;
    BAide: TToolbarButton97;
    FTriJal: TCheckBox;
    QRLabel7: TQRLabel;
    RTriJal: TQRLabel;
    QEcrCALCRUPT: TStringField;
    QEcrE_EXERCICE: TStringField;
    QEcrE_JOURNAL: TStringField;
    QEcrE_DATECOMPTABLE: TDateTimeField;
    QEcrE_NUMEROPIECE: TIntegerField;
    QEcrE_NUMLIGNE: TIntegerField;
    QEcrE_GENERAL: TStringField;
    QEcrE_AUXILIAIRE: TStringField;
    QEcrE_REFINTERNE: TStringField;
    QEcrDEBIT: TFloatField;
    QEcrCREDIT: TFloatField;
    QEcrE_LIBELLE: TStringField;
    QEcrE_NATUREPIECE: TStringField;
    QEcrE_QUALIFPIECE: TStringField;
    QEcrE_TYPEMVT: TStringField;
    QEcrE_UTILISATEUR: TStringField;
    QEcrE_DATECREATION: TDateTimeField;
    QEcrE_ETABLISSEMENT: TStringField;
    QEcrE_REGIMETVA: TStringField;
    QEcrE_TVA: TStringField;
    QEcrE_TPF: TStringField;
    QEcrE_MODEPAIE: TStringField;
    QEcrE_DATEECHEANCE: TDateTimeField;
    QEcrE_DEVISE: TStringField;
    QEcrE_TAUXDEV: TFloatField;
    QEcrE_DATETAUXDEV: TDateTimeField;
    QEcrE_QTE1: TFloatField;
    QEcrE_QTE2: TFloatField;
    QEcrE_QUALIFQTE1: TStringField;
    QEcrE_ECHE: TStringField;
    QEcrE_QUALIFQTE2: TStringField;
    QEcrE_VALIDE: TStringField;
    QEcrT_ABREGE: TStringField;
    QEcrG_ABREGE: TStringField;
    QEcrE_ETATLETTRAGE: TStringField;
    QEcrE_COTATION: TFloatField;
    QEcrT_LIBELLE: TStringField;
    QEcrG_LIBELLE: TStringField;
    QEcrE_NUMECHE: TIntegerField;
    QEcrE_REFEXTERNE: TStringField;
    QEcrE_REFLIBRE: TStringField;
    QEcrE_AFFAIRE: TStringField;
    QEcrE_DATEREFEXTERNE: TDateTimeField;
    RNumVersion: TQRLabel;
    QEcrMULTIECHE: TIntegerField;
    BDupFiltre: TMenuItem;
    procedure FormShow(Sender: TObject);
    procedure BDetailBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure BAnalytiqueBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure BEntetePieceAfterPrint(BandPrinted: Boolean);
    procedure BPiedPieceBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure BrouillardAfterPrint(BandPrinted: Boolean);
    procedure BValiderClick(Sender: TObject);
    procedure piedpageAfterPrint(BandPrinted: Boolean);
    procedure FDateCreat1KeyPress(Sender: TObject; var Key: Char);
    procedure FExerciceChange(Sender: TObject);
    procedure FDevisesChange(Sender: TObject);
    procedure BFinEtatBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure TOPREPORTBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure BOTTOMREPORTBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure EnteteAutrePageAfterPrint(BandPrinted: Boolean);
    procedure BEntetePieceBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure QEcrAfterOpen(DataSet: TDataSet);
    procedure QAnaAfterOpen(DataSet: TDataSet);
    procedure QAnaBeforeOpen(DataSet: TDataSet);
    procedure FCpte1Change(Sender: TObject);
    procedure QRDLEcheNeedData(var MoreData: Boolean);
    procedure BEcheBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure BPiedPieceAfterPrint(BandPrinted: Boolean);
    procedure FormActivate(Sender: TObject);
    procedure BAideClick(Sender: TObject);
    procedure POPFPopup(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormCreate(Sender: TObject);
    procedure FTriClick(Sender: TObject);
    procedure FTriJalClick(Sender: TObject);
    procedure QEcrCalcFields(DataSet: TDataSet);
    procedure BFermeClick(Sender: TObject);
  private { Déclarations privées }
    HighBand : Boolean ;
    StReportPiece, FNomFiltre           : String ;
    Last, AncienEche     : integer ;
    CritEdt                             : TCritEdt ;
    PremTabCol                          : TTabCollAffEdt ;
    TotGen, TotPiece                    : TabTot ;
    LEche                               : TStringList ;
    CodeNatureEcr                       : String3 ;
    QRCritMemoire                       : HTStrings ;
    OkChargeCritMemoire                 : Boolean ;
    QR2E_EXERCICE, QR2E_JOURNAL,
    QR2E_GENERAL,  QR2E_AUXILIAIRE,
    QR2E_REFINTERNE, QR2E_LIBELLE,
    QR2E_VALIDE, QR2E_NATUREPIECE,
    QR2E_QUALIFPIECE, QR2E_TYPEMVT,
    QR2E_UTILISATEUR, QR2E_ETABLISSEMENT,
    QR2E_REGIMETVA, QR2E_MODEPAIE,
    QR2E_DEVISE, QR2E_QUALIFQTE1,
    QR2E_QUALIFQTE2, QR2T_LIBELLE,
    QR2G_LIBELLE, QR2E_ETATLETTRAGE,
    QR2E_REFEXTERNE,
    QR2E_REFLIBRE, QR2E_AFFAIRE               : TStringField ;
    QR2E_DATEECHEANCE, QR2E_DATECOMPTABLE,
    QR2E_DATECREATION, QR2E_DATETAUXDEV,
    QR2E_DATEREFEXTERNE                       : TDateTimeField ;
    QR2E_NUMEROPIECE, QR2E_NUMLIGNE,
    QR2E_ECHE, QR2E_NUMECHE, QR2MULTIECHE     : TIntegerField ;
    QR2E_TVA, QR2E_TPF, QR2E_TAUXDEV, Qr2E_COTATION,
    QR2E_QTE1, QR2E_QTE2, QR2DEBIT, QR2CREDIT : TFloatField ;
    { Analytique }
    QR3Y_EXERCICE, QR3Y_AFFAIRE,
    QR3Y_JOURNAL, QR3Y_REFEXTERNE             : TStringField ;
    QR3Y_DATECOMPTABLE, QR3Y_DATEREFEXTERNE   : TDateTimeField ;
    QR3Y_NUMEROPIECE, QR3Y_NUMLIGNE,
    QR3Y_NUMVENTIL                            : TIntegerField ;
    QR3Y_GENERAL, QR3Y_VALIDE, QR3Y_AXE,
    QR3Y_QUALIFQTE1, QR3Y_QUALIFQTE2,
    QR3Y_REFINTERNE, QR3Y_REFLIBRE            : TStringField ;
    QR3Y_LIBELLE        : TStringField ;
    QR3Y_SECTION        : TStringField ;
    QR3Y_POURCENTQTE1   : TFloatField ;
    QR3Y_POURCENTQTE2   : TFloatField ;
    QR3Y_POURCENTAGE    : TFloatField ;
    QR3Y_QTE1           : TFloatField ;
    QR3Y_QTE2           : TFloatField ;
    QR3S_ABREGE         : TStringField ;
    QR3DEBIT            : TFloatField ;
    QR3CREDIT           : TFloatField ;
    ObjFiltre : TObjFiltre; // 14676
    procedure InitDivers ;
    procedure SQLEcr ;
    procedure SQLAna ;
    function  AfficheMontant( Formatage, LeSymbole : String ; LeMontant : Double ; OkSymbole : Boolean ) : String ;
    procedure RenseigneCritere ;
    procedure ChoixEdition ;
    procedure RecupCritEdt ;
    function  CritEdtOk : Boolean ;
    Procedure BrouilZoom(Quoi : String) ;
    Function  QuoiMvt(i : Integer) : String ;
    Procedure FinirPrint ;
    Function  RetourneCouple(A, B : String) : String ;
  public

  end;

implementation

{$R *.DFM}

uses QRE,QR ;

Procedure TQRBrouillard.BrouilZoom(Quoi : String) ;
Var Lp,i : Integer ;
BEGIN
Lp:=Pos('@',Quoi) ; If Lp=0 Then Exit ;
i:=StrToInt(Copy(Quoi,Lp+1,1)) ;
If (i=5) Then
   BEGIN
   Quoi:=Copy(Quoi,Lp+3,Length(Quoi)-lp-2) ;
   If QRP.QrPrinter.FSynShiftDblClick Then i:=6 ;
   If QRP.QRPrinter.FSynCtrlDblClick Then i:=11 ;
   END ;
If (i=7) Then Quoi:=Copy(Quoi,Lp+3,Length(Quoi)-lp-2) ;
ZoomEdt(i,Quoi) ;

(* Rony 26/05/97
Lp:=Pos('@',Quoi) ; If Lp=0 Then Exit ;
Quoi:=Copy(Quoi,Lp+3,Length(Quoi)-lp-2) ;
ZoomEdt(5,Quoi) ;
*)
END ;


Function TQRBrouillard.QuoiMvt(i : Integer) : String ;
BEGIN
if i=0 then
   BEGIN
   Result:=QR2E_GENERAL.AsString+'   '+Qr2E_AUXILIAIRE.AsString+
           '#'+Qr2E_JOURNAL.AsString+' N° '+IntToStr(Qr2E_NUMEROPIECE.AsInteger)+' '+DateToStr(Qr2E_DateComptable.AsDAteTime)+'-'+
           PrintSolde(Qr2DEBIT.AsFloat,Qr2Credit.AsFloat,CritEDT.Decimale,CritEDT.Symbole,CritEDT.AfficheSymbole)+
           '@'+'5;'+Qr2E_JOURNAL.AsString+';'+UsDateTime(Qr2E_DATECOMPTABLE.AsDateTime)+';'+IntToStr(Qr2E_NUMEROPIECE.AsInteger)+';'+Qr2E_EXERCICE.asString+';'+
           IntToStr(Qr2E_NumLigne.AsInteger)+';' ;
   END Else
if i=1 then
   BEGIN
   //Result:=QR3Y_AXE.AsString+'   '+QR3Y_SECTION.AsString+'   '+Qr3Y_GENERAL.AsString+
   Result:=QR3Y_SECTION.AsString+'   '+QR3S_ABREGE.AsString+
           '#'+Qr3Y_JOURNAL.AsString+' N° '+IntToStr(Qr3Y_NUMEROPIECE.AsInteger)+' '+DateToStr(Qr3Y_DateComptable.AsDAteTime)+'-'+
           PrintSolde(Qr3DEBIT.AsFloat,Qr3Credit.AsFloat,CritEDT.Decimale,CritEDT.Symbole,CritEDT.AfficheSymbole)+
           '@'+'7;'+Qr3Y_JOURNAL.AsString+';'+UsDateTime(Qr3Y_DATECOMPTABLE.AsDateTime)+';'+Qr3Y_NUMEROPIECE.AsString+';'+Qr3Y_EXERCICE.asString+';'+
           IntToStr(Qr3Y_NumLigne.AsInteger)+';'+QR3Y_AXE.AsString+';' ;
   END ;
END ;

procedure Brouillard(Code : char) ;
var QR : TQRBrouillard ;
    Libelle : String ;
    PP : THPanel ;
BEGIN
SourisSablier ;
PP:=FindInsidePanel ;
QR:=TQRBrouillard.Create(Application) ;
QR.CodeNatureEcr:=Code ;
Case Code of
 'N' : BEGIN
       Libelle:=QR.MsgBox.Mess[11] ; QR.FNomFiltre:='BROUILNOR' ;
       QR.HelpContext:=7268000 ;
       //QR.Standards.HelpContext:=7268100 ;
       //QR.Avances.HelpContext:=7268200 ;
       //QR.Mise.HelpContext:=7268300 ;
       //QR.Option.HelpContext:=7268400 ;
       END ;
 'P' : BEGIN Libelle:=QR.MsgBox.Mess[12] ; QR.FNomFiltre:='BROUILPRE' ; END ;
 'R' :
       BEGIN
       Libelle:=QR.MsgBox.Mess[13] ; QR.FNomFiltre:='BROUILREV' ;
       QR.HelpContext:=7670000 ;
       //QR.Standards.HelpContext:=7670010 ;
       //QR.Avances.HelpContext:=7670020 ;
       //QR.Mise.HelpContext:=7670030 ;
       //QR.Option.HelpContext:=7670040 ;
       END ;
 'S' : BEGIN
       Libelle:=QR.MsgBox.Mess[14] ; QR.FNomFiltre:='BROUILSIM' ;
       QR.HelpContext:=7289000 ;
       //QR.Standards.HelpContext:=7289100 ;
       //QR.Avances.HelpContext:=7289200 ;
       //QR.Mise.HelpContext:=7289300 ;
       //QR.Option.HelpContext:=7289400 ;
       END ;
 'U' : BEGIN Libelle:=QR.MsgBox.Mess[15] ; QR.FNomFiltre:='BROUILSIT' ; END ;
 End ;
QR.Caption:=QR.MsgBox.Mess[9]+' '+Libelle ;
QR.QRP.QRPrinter.OnSynZoom:=QR.BrouilZoom ;
if PP=Nil then
   BEGIN
   try
    QR.ShowModal ;
    finally
    QR.Free ;
    SourisNormale ;
    end ;
   Screen.Cursor:=SyncrDefault ;
   END Else
   BEGIN
   InitInside(QR,PP) ;
   QR.Show ;
   END ;
END ;

procedure TQRBrouillard.BValiderClick(Sender: TObject);
  Label 0 ;
begin
SourisSablier ;
EnableControls(Self,False) ;
if CritEdtOK then
   BEGIN
   If CritEdt.RappelCrit=QueRappel Then Goto 0 ;
   InitDivers ;
   SQLEcr ;
   if CritEdt.Bro.AfficheAnal then SQLAna ;
   RenseigneCritere ;
   ChoixEdition ;
   QRP.QRPrinter.SynOnGrid:=(FListe.Checked) ;
   if Reduire.Checked then QRP.QRPrinter.Thumbs:=2 else QRP.QRPrinter.Thumbs:=1 ;
   if QEcr.Eof
      then BEGIN MsgRien.Execute(0,'','')END
      else BEGIN if Apercu.Checked then QRP.Preview else QRP.Print ; END ;
   FinirPrint ;
   END Else MsgRien.Execute(1,'','') ;
0:
If Apercu.Checked Then
   BEGIN
   If CritEdt.RappelCrit=QueRappel Then
      BEGIN
      PrintPageDeGarde(Pages,FCouleur.Checked,Apercu.Checked,TRUE,Caption,0) ;
      END ;
   END Else
   BEGIN
   If CritEdt.RappelCrit<>SansRappel Then
      BEGIN
      PrintPageDeGarde(Pages,FCouleur.Checked,Apercu.Checked,TRUE,Caption,0) ;
      END ;
   END ;
EnableControls(Self,True) ;
SourisNormale ;
end;

Procedure TQRBrouillard.FinirPrint ;
BEGIN
VideGroup(LEche) ;
QEcr.Close ; if CritEdt.Bro.AfficheAnal then QAna.Close ;
END ;

procedure TQRBrouillard.SQLEcr ;
{ Construction de la requête SQL en fonction du multicritère }
BEGIN
{ Construction de la clause Select de la SQL }
QEcr.Close ;
QEcr.SQL.Clear ;
QEcr.SQL.Add('Select  E_EXERCICE, E_JOURNAL, E_DATECOMPTABLE, E_NUMEROPIECE, E_NUMLIGNE, E_ECHE,') ;
QEcr.SQL.Add(       ' E_GENERAL, E_AUXILIAIRE, E_QUALIFQTE2 , E_REFINTERNE, E_LIBELLE, E_VALIDE,') ;
QEcr.SQL.Add(       ' E_NATUREPIECE, E_QUALIFPIECE, E_TYPEMVT, E_UTILISATEUR, E_DATECREATION,') ;
QEcr.SQL.Add(       ' E_ETABLISSEMENT, E_REGIMETVA, E_TVA, E_TPF, E_MODEPAIE, E_DATEECHEANCE,') ;
QEcr.SQL.Add(       ' E_DEVISE, E_TAUXDEV, E_DATETAUXDEV, E_COTATION, ') ;
QEcr.SQL.Add(       ' E_QTE1, E_QTE2, E_QUALIFQTE1, T_LIBELLE, T_ABREGE, G_LIBELLE, G_ABREGE, E_ETATLETTRAGE, E_NUMECHE, ') ;
QEcr.SQL.Add(       ' E_REFEXTERNE, E_REFLIBRE, E_AFFAIRE, E_DATEREFEXTERNE, ') ;
// GP : bug edition avec filtre sur N° pièce trop grand (à cause du *10)
//QEcr.SQL.Add(       ' E_NUMEROPIECE*10.0+E_NUMLIGNE as MULTIECHE, ') ;
QEcr.SQL.Add(       ' E_NUMLIGNE as MULTIECHE, ') ;
Case CritEdt.Monnaie of
  0 : BEGIN QEcr.SQL.Add(' E_DEBIT DEBIT, E_CREDIT CREDIT ')         ; END ;
  1 : BEGIN QEcr.SQL.Add(' E_DEBITDEV DEBIT, E_CREDITDEV CREDIT ')   ; END ;
end ;
{ Tables explorées par la SQL }
QEcr.SQL.Add(' From Ecriture ') ;
QEcr.SQL.Add(       ' LEFT JOIN TIERS On T_AUXILIAIRE=E_AUXILIAIRE ') ;
QEcr.SQL.Add(       ' LEFT JOIN GENERAUX On G_GENERAL=E_GENERAL ') ;
{ Construction de la clause Where de la SQL }
QEcr.SQL.Add(' Where E_QUALIFPIECE="'+CodeNatureEcr+'" ') ;
QEcr.SQL.Add(' And E_DATECOMPTABLE>="'+USDateTime(CritEdt.Date1)+'" And E_DATECOMPTABLE<="'+USDateTime(CritEdt.Date2)+'" ') ;
if CritEdt.Joker then QEcr.SQL.Add(' AND E_JOURNAL like "'+TraduitJoker(CritEdt.Cpt1)+'" ') Else
   BEGIN
   if CritEdt.Cpt1<>'' then QEcr.SQL.Add(' AND E_JOURNAL>="'+CritEdt.Cpt1+'" ') ;
   if CritEdt.Cpt2<>'' then QEcr.SQL.Add(' AND E_JOURNAL<="'+CritEdt.Cpt2+'" ') ;
   END ;
QEcr.SQL.Add(' And E_NUMEROPIECE>='+IntToStr(CritEdt.Bro.NumPiece1)+' ') ;
QEcr.SQL.Add(' And E_NUMEROPIECE<='+IntToStr(CritEdt.Bro.NumPiece2)+' ') ;
if CritEdt.Exo.Code<>'' then QEcr.SQL.Add(' And E_EXERCICE="'+CritEdt.Exo.Code+'" ') ;
if CritEdt.Bro.RefInterne<>'' then QEcr.SQL.Add('and UPPER(E_REFINTERNE) like "'+TraduitJoker(CritEdt.Bro.RefInterne)+'" ' );
if CritEdt.Etab<>'' then QEcr.SQL.Add(' And E_ETABLISSEMENT="'+CritEdt.Etab+'" ') ;
if CritEdt.Valide<>'g' then QEcr.SQL.Add(' And E_VALIDE="'+CritEdt.Valide+'" ') ;
if CritEdt.DeviseSelect<>'' then QEcr.SQL.Add(' And E_DEVISE="'+CritEdt.DeviseSelect+'" ') ;
If Not CritEdt.Bro.AvecJalAN Then QEcr.SQL.Add(' And E_ECRANOUVEAU="N" ')
                             Else QEcr.SQL.Add(' And (E_ECRANOUVEAU="N" or E_ECRANOUVEAU="H" or E_ECRANOUVEAU="OAN") ') ;

if (OkV2 and (V_PGI.Confidentiel<>'1')) then QEcr.SQL.Add('AND E_CONFIDENTIEL<>"1" ') ;

{ Construction de la clause Order By de la SQL }
(*
if CritEdt.Bro.Tri then QEcr.SQL.Add(' Order By E_DATECOMPTABLE, E_NUMEROPIECE, E_NUMLIGNE ') Else
  if CritEdt.Bro.TriJal then QEcr.SQL.Add(' Order By E_JOURNAL, E_NUMEROPIECE, E_NUMLIGNE ') Else
                                QEcr.SQL.Add(' Order By E_NUMEROPIECE, E_NUMLIGNE, E_JOURNAL ') ;
*)
if CritEdt.Bro.Tri then QEcr.SQL.Add(' Order By E_DATECOMPTABLE, E_JOURNAL, E_NUMEROPIECE, E_NUMLIGNE ') Else
  if CritEdt.Bro.TriJal then QEcr.SQL.Add(' Order By E_JOURNAL, E_NUMEROPIECE, E_NUMLIGNE ') Else
                                QEcr.SQL.Add(' Order By E_NUMEROPIECE, E_NUMLIGNE, E_JOURNAL, E_DATECOMPTABLE  ') ;
                   //else QEcr.SQL.Add(' Order By E_NUMEROPIECE, E_NUMLIGNE ') ;
ChangeSQL(QEcr) ; QEcr.Open ;
END ;

procedure TQRBrouillard.SQLAna ;
{ Construction de la requête SQL en fonction du multicritère }
BEGIN
{ Construction de la clause Select de la SQL }
QAna.Close ;
QAna.SQL.Clear ;
QAna.SQL.Add('Select  Y_EXERCICE, Y_JOURNAL, Y_DATECOMPTABLE, Y_NUMEROPIECE, Y_NUMLIGNE,  ') ;
QAna.SQL.Add(       ' Y_GENERAL,Y_AXE, Y_QUALIFQTE2 , Y_REFINTERNE, Y_LIBELLE, Y_SECTION, ') ;
QAna.SQL.Add(       ' Y_POURCENTAGE,Y_POURCENTQTE1,Y_POURCENTQTE2, Y_REFLIBRE, ') ;
QAna.SQL.Add(       ' Y_QTE1, Y_QTE2, Y_QUALIFQTE1, S_ABREGE, Y_NUMVENTIL, ') ;
QAna.SQL.Add(       ' Y_AFFAIRE, Y_REFEXTERNE, Y_DATEREFEXTERNE, Y_VALIDE, ') ;
Case CritEdt.Monnaie of
  0 : BEGIN QAna.SQL.Add(' Y_DEBIT YDEBIT, Y_CREDIT YCREDIT ')         ; END ;
  1 : BEGIN QAna.SQL.Add(' Y_DEBITDEV YDEBIT, Y_CREDITDEV YCREDIT ')   ; END ;
end ;
{ Tables explorées par la SQL }
QAna.SQL.Add(' From ANALYTIQ ') ;
QAna.SQL.Add(' LEFT JOIN SECTION on S_AXE=Y_AXE and S_SECTION=Y_SECTION ') ;
{ Construction de la clause Where de la SQL }
QAna.SQL.Add(' Where Y_JOURNAL=:E_JOURNAL ') ;
QAna.SQL.Add(' And Y_EXERCICE=:E_EXERCICE  ') ;
QAna.SQL.Add(' And Y_DATECOMPTABLE=:E_DATECOMPTABLE ') ;
QAna.SQL.Add(' And Y_NUMEROPIECE=:E_NUMEROPIECE ') ;
QAna.SQL.Add(' And Y_NUMLIGNE=:E_NUMLIGNE ') ;
QAna.SQL.Add(' And Y_QUALIFPIECE="'+CodeNatureEcr+'" ') ;
{ Construction de la clause Order By de la SQL }
QAna.SQL.Add(' Order By Y_JOURNAL, Y_EXERCICE, Y_DATECOMPTABLE, Y_NUMEROPIECE, Y_NUMLIGNE, Y_AXE, Y_NUMVENTIL,Y_QUALIFPIECE ') ;
(*if CritEdt.Bro.Tri then QAna.SQL.Add(' Order By Y_DATECOMPTABLE, Y_NUMEROPIECE, Y_AXE  ')
                   Else QAna.SQL.Add(' Order By Y_NUMEROPIECE, Y_AXE ') ;
*)
ChangeSQL(QAna) ; QAna.Open ;
END ;

procedure TQRBrouillard.InitDivers ;
Var ll : Integer ;
BEGIN
{ Labels sur les bandes } //
AncienEche:=1 ;
TRien.Caption:='' ;
If (Not FRefExt.Checked) and (Not FDateRefExt.Checked) then
   BEGIN
   TE_REFINTERNE.Caption:=MsgBox.Mess[19]  ;
   E_REFINTERNE.DataField:='E_REFINTERNE' ; Y_REFINTERNE.DataField:='Y_REFINTERNE' ;
   END Else
If (FRefExt.Checked) and (Not FDateRefExt.Checked) then
   BEGIN
   TE_REFINTERNE.Caption:=MsgBox.Mess[22] ;
   E_REFINTERNE.DataField:='E_REFEXTERNE' ; Y_REFINTERNE.DataField:='Y_REFEXTERNE' ;
   END Else
If (Not FRefExt.Checked) and (FDateRefExt.Checked) then
   BEGIN
   TE_REFINTERNE.Caption:=MsgBox.Mess[21] ;
   E_REFINTERNE.DataField:='E_DATEREFEXTERNE' ; Y_REFINTERNE.DataField:='Y_DATEREFEXTERNE' ;
   END Else
If (FRefExt.Checked) and (FDateRefExt.Checked) then
   BEGIN
   TE_REFINTERNE.Caption:=MsgBox.Mess[23] ;
   E_REFINTERNE.DataField:='' ; Y_REFINTERNE.DataField:='' ;
   END ;


If (Not FRefLib.Checked) and (Not FAffaire.Checked) then
   BEGIN
   TE_LIBELLE.Caption:=MsgBox.Mess[20] ;
   E_LIBELLE.DataField:='E_LIBELLE' ; Y_LIBELLE.DataField:='Y_LIBELLE' ;
   END Else
If (FRefLib.Checked) and (Not FAffaire.Checked) then
   BEGIN
   TE_LIBELLE.Caption:=MsgBox.Mess[24] ;
   E_LIBELLE.DataField:='E_REFLIBRE' ; Y_LIBELLE.DataField:='Y_REFLIBRE' ;
   END Else
If (Not FRefLib.Checked) and (FAffaire.Checked) then
   BEGIN
   TE_LIBELLE.Caption:=MsgBox.Mess[25] ;
   E_LIBELLE.DataField:='E_AFFAIRE' ; Y_LIBELLE.DataField:='Y_AFFAIRE' ;
   END Else
If (FRefLib.Checked) and (FAffaire.Checked) then
   BEGIN
   TE_LIBELLE.Caption:=MsgBox.Mess[26] ;
   E_LIBELLE.DataField:='' ; Y_LIBELLE.DataField:='' ;
   END ;
RSociete.Caption:=MsgBox.Mess[0]+' '+V_PGI.CodeSociete+' '+V_PGI.NomSociete ;
RUtilisateur.Caption:=MsgBox.Mess[1]+' '+V_PGI.User+' '+V_PGI.UserName ;
RCopyright.Caption := Copyright + ' - ' + TitreHalley ;
RNumversion.Caption:=MsgBox.Mess[31]+' '+V_PGI.NumVersion+' '+MsgBox.Mess[32]+' '+DateToStr(V_PGI.DateVersion);

EnteteAutrePage.Enabled:=False ;
if CritEdt.Bro.FormatPrint.Report.OkAff then
   BEGIN
   TopReport.Enabled:=FALSE ; BottomReport.Enabled:=TRUE ;
   END else
   BEGIN
   TopReport.Enabled:=FALSE ; BottomReport.Enabled:=FALSE ;
   END ;

HighBand:=((FRefExt.Checked)or(FRefLib.Checked)or(FAffaire.Checked)or(FDateRefExt.Checked)) ;
if HighBand then
   BEGIN
   BEntetePiece.Height:=(TE_NUMEROPIECE.Top+TE_NUMEROPIECE.Height)+E_REFINTERNE2.Height+3 ;
   BDetail.Font.Size:=7 ; BAnalytique.Font.Size:=7 ;
   END else
   BEGIN
   BEntetePiece.Height:=(TE_NUMEROPIECE.Top+TE_NUMEROPIECE.Height)+3 ;
   BDetail.Font.Size:=8 ; BAnalytique.Font.Size:=8 ;
   END ;



BEntetePiece.Frame.DrawTop:=CritEdt.Bro.FormatPrint.PrSepCompte[1] ;
BEntetePiece.Frame.DrawBottom:=CritEdt.Bro.FormatPrint.PrSepCompte[1] ;
BPiedPiece.Frame.DrawTop:=CritEdt.Bro.FormatPrint.PrSepCompte[2] ;
BPiedPiece.Frame.DrawBottom:=CritEdt.Bro.FormatPrint.PrSepCompte[2] ;

ll:=ChangeFormatEdt(Self,PremTabCol,CritEdt.Bro.FormatPrint.TabColl) ;
CalculPourMiseEnPageEdt(ll,QRP,Self,CritEdt.Bro.FormatPrint) ;

END ;

procedure TQRBrouillard.RenseigneCritere ;
{ Récupération des champs du multicritère dans l'entête d'état }
Var St1,St2 : String ;
BEGIN
ReInitEdit(EntetePage,FCouleur.Checked) ;
With CritEdt do
  BEGIN
  if FJoker.Visible then
     BEGIN
     RCpte.Visible:=False ; TRa.visible:=False ;
     RCpte2.Visible:=False ; RJoker.Visible:=True;
     RCpte1.Caption:=FJoker.Text ; RCpte2.Caption:=FJoker.Text ;
     END else
     BEGIN
     RCpte.Visible:=True ; TRa.visible:=True ;
     RCpte2.Visible:=True ; RJoker.Visible:=False;
     PositionneFourchetteST(FCpte1,FCpte2,St1,St2) ;
     RCpte1.Caption:=St1 ;
     RCpte2.Caption:=St2 ;
     END ;
  RDateCompta1.Caption:=DatetoStr(Date1)           ; RDateCompta2.Caption:=DatetoStr(Date2) ;
 // PositionneFourchette(FJal1,FJal2,RJal1,RJal2)    ;
  RRefInterne.Caption:=CritEdt.Bro.RefInterne       ;
  RNumPiece1.Caption:=IntToStr(Bro.NumPiece1)     ; RNumPiece2.Caption:=IntToStr(Bro.NumPiece2) ;
  REtab.Caption:=CritEdt.Etab                     ; RDevises.Caption:=FDevises.text ;
  RExercice.Caption:=FExercice.Text ;
  CaseACocher(FTri,RTri)         ; CaseACocher(FValide,RValide) ; CaseACocher(FTriJal,RTriJal) ;
  CaseACocher(FRefExt,RRefExt)   ; CaseACocher(FRefLib,RRefLib) ;
  CaseACocher(FAffaire,RAffaire) ; CaseACocher(FDateRefExt,RDateRefExt) ;
  END ;
Case CritEdt.Monnaie Of
  0 : RDevises.Caption:=RDevises.Caption+' / Aff. '+VH^.LibDevisePivot ;
  1 : RDevises.Caption:=RDevises.Caption+TraduireMemoire(' / Aff. Devise') ;
END ;
InitQrPrinter(QRP,FListe.Checked,Reduire.Checked,FCouleur.Checked) ;
END ;

procedure TQRBrouillard.ChoixEdition ;
{ Initialisation des options d'édition }
Var RDev : RDevise ;
BEGIN
ChargeGroup(LEche,['ECHE']) ;
if (CritEdt.DeviseSelect='') or (CritEdt.Monnaie=0) Or (CritEdt.DeviseSelect=V_PGI.DevisePivot) then
   BEGIN
   CritEdt.DeviseAffichee:=V_PGI.DevisePivot ;
   CritEdt.Decimale:=V_PGI.OkDecV ;
   CritEdt.Symbole:=V_PGI.SymbolePivot ;
   END Else
   BEGIN
   CritEdt.DeviseAffichee:=CritEdt.DeviseSelect ;
   RDev.Code:=CritEdt.DeviseAffichee ;
   GetInfosDevise(RDev) ;
   CritEdt.Decimale:=RDev.Decimale ;
   CritEdt.Symbole:=RDev.Symbole ;
   If CritEdt.Monnaie=2 Then
     BEGIN
     CritEdt.Decimale:=V_PGI.OkDecE ;
     //If Not VH^.TenueEuro Then CritEdt.Symbole:='' ;
     END ;
   END ;

If CritEdt.Monnaie=2 Then CritEdt.AfficheSymbole:=FALSE ;
ChangeMask(Formateur,CritEdt.Decimale,CritEdt.Symbole);
CritEdt.FormatMontant:=Formateur.Masks.PositiveMask;

ChgMaskChamp(QR2DEBIT,CritEdt.Decimale,CritEdt.AfficheSymbole,CritEdt.Symbole,False) ;
ChgMaskChamp(QR2CREDIT,CritEdt.Decimale,CritEdt.AfficheSymbole,CritEdt.Symbole,False) ;
// Rony 4/04/97 ChgMaskChamp(QR3DEBIT,CritEdt.Decimale,CritEdt.AfficheSymbole,CritEdt.Symbole,False) ;
//ChgMaskChamp(QR3CREDIT,CritEdt.Decimale,CritEdt.AfficheSymbole,CritEdt.Symbole,False) ;
END ;

procedure TQRBrouillard.FormShow(Sender: TObject);
begin
QRCritMemoire:=HTStringList.Create ;
OkChargeCritMemoire:=TRUE ;
PopUpMenu:=ADDMenuPop(PopUpMenu,'','') ;
TabSup.TabVisible:=False;
QRP.ReportTitle:=Caption ;
Pages.ActivePage:=Standards ;
HorzScrollBar.Visible:=FALSE ; VertScrollBar.Visible:=FALSE ;
FCpte1.Text:=''; FCpte2.Text:=''; // Rony 16/04/97 FUtil1.Text:=''; FUtil2.Text:='';
FRefInterne.Text:='';
FDevises.ItemIndex:=0 ; FExercice.Value:=VH^.Entree.Code ;
FEtab.ItemIndex:=0 ; PositionneEtabUser(FEtab) ;
FMontant.ItemIndex:=0 ;
FValide.State:=cbGrayed ;
FValide.Visible:=(CodeNatureEcr<>'S')and(CodeNatureEcr<>'R') ;
FCouleur.Checked:=V_PGI.QRCouleur ;

InitFormatEdt(Self,PremTabCol,CritEdt.Bro.FormatPrint.TabColl) ;
InitEdit(EntetePage,QRP) ;
SourisNormale ;

ObjFiltre.FFI_TABLE := FNomFiltre;
ObjFiltre.Charger;
end;

procedure TQRBrouillard.BEntetePieceAfterPrint(BandPrinted: Boolean);
begin
Fillchar(TotPiece, SizeOf(TotPiece),#0) ;
EntetePage.Enabled:=True ;
end;

procedure TQRBrouillard.BDetailBeforePrint(var PrintBand: Boolean; var Quoi: string);
var MReport : TabTRep ;
    St      : String ;
begin
AncienEche:=QR2E_NUMECHE.AsInteger ; Fillchar(MReport,SizeOf(MReport),#0) ;
AddGroup(LEche, [Qr2MULTIECHE], [Qr2DEBIT.AsFloat, Qr2CREDIT.AsFloat]) ;
LeCompte.Caption:=QR2E_GENERAL.ASString+'  '+QR2E_AUXILIAIRE.ASString ;
if QR2E_VALIDE.AsString='X' then E_VALIDE.Caption:='V' else E_VALIDE.Caption:=' ' ;
If QR2E_NUMECHE.AsInteger=0 then E_DATEECHEANCE.Caption:='' ;
If FRefExt.Checked and FDateRefExt.Checked then
   BEGIN
   If (QR2E_DATEREFEXTERNE.AsDateTime<>IDate1900) then St:=DateToStr(QR2E_DATEREFEXTERNE.AsDateTime)
                                                  Else St:='' ;
   E_REFINTERNE.Caption:=RetourneCouple(St, QR2E_REFEXTERNE.AsString) ;
   END ;
If FRefLib.Checked and FAffaire.Checked then
   BEGIN
   E_LIBELLE.Caption:=RetourneCouple(QR2E_REFLIBRE.AsString, QR2E_AFFAIRE.AsString) ;
   END ;
If (Not FRefExt.Checked) and (FDateRefExt.Checked) and (QR2E_DATEREFEXTERNE.AsDateTime=IDate1900) then E_REFINTERNE.Caption:='' ;
{quoi:=E_GENERAL.Caption ;}
MReport[1].TotDebit:=QR2DEBIT.AsFloat ;
MReport[1].TotCredit:=QR2CREDIT.AsFloat ;
TotGen[1].TotDebit:= Arrondi(TotGen[1].TotDebit+QR2DEBIT.AsFloat,CritEdt.Decimale) ;
TotGen[1].TotCredit:=Arrondi(TotGen[1].TotCredit+QR2CREDIT.AsFloat,CritEdt.Decimale) ;
TotPiece[1].TotDebit:= Arrondi(TotPiece[1].TotDebit+QR2DEBIT.AsFloat,CritEdt.Decimale) ;
TotPiece[1].TotCredit:=Arrondi(TotPiece[1].TotCredit+QR2CREDIT.AsFloat,CritEdt.Decimale) ;
if (QR2E_ECHE.asString='X') And (QR2E_ETATLETTRAGE.asString<>'RI') then
   BEGIN
   E_NUMLIGNE.Caption:=IntToStr(QR2E_NUMLIGNE.AsInteger)+'/'+IntToStr(QR2E_NUMECHE.AsInteger) ;
   END else
   BEGIN
   END ;
T_ABREGE.Visible:=(QR2E_AUXILIAIRE.asString<>'') ;
G_ABREGE.Visible:=Not T_ABREGE.Visible ;
AddReportBAL([1,2], CritEdt.Bro.FormatPrint.Report, MReport, CritEdt.Decimale) ;
Quoi:=QuoiMvt(0) ;
end;

procedure TQRBrouillard.BAnalytiqueBeforePrint(var PrintBand: Boolean; var Quoi: string);
Var St : String ;
begin
if (Not CritEdt.Bro.AfficheAnal) or (QAna.Active=FALSE) or (QAna.Eof) then PrintBand:=FALSE ;
if PrintBand then
   BEGIN
   if QR3Y_VALIDE.AsString='X' then Y_VALIDE.Caption:='V' else Y_VALIDE.Caption:=' ' ;
   Y_NumLV.Caption:=QR3Y_NUMLIGNE.AsString+' / '+QR3Y_NUMVENTIL.AsString ;
   Y_SECTION.Caption:=QR3Y_SECTION.AsString+' - '+QR3Y_AXE.AsString ;
   Y_POURCENTAGE.Caption:=FloatToStr(QR3Y_POURCENTAGE.AsFloat)+' '+MsgBox.Mess[28] ;
   If FRefExt.Checked and FDateRefExt.Checked then
      BEGIN
      If (QR3Y_DATEREFEXTERNE.AsDateTime<>IDate1900) then St:=DateToStr(QR3Y_DATEREFEXTERNE.AsDateTime)
                                                                   Else St:='' ;
      Y_REFINTERNE.Caption:=RetourneCouple(St, QR3Y_REFEXTERNE.AsString) ;
      END ;
   If FRefLib.Checked and FAffaire.Checked then
      BEGIN
      Y_LIBELLE.Caption:=RetourneCouple(QR3Y_REFLIBRE.AsString, QR3Y_AFFAIRE.AsString) ;
      END ;
   If (Not FRefExt.Checked) and (FDateRefExt.Checked) and (QR3Y_DATEREFEXTERNE.AsDateTime=IDate1900) then Y_REFINTERNE.Caption:='' ;
   Quoi:=QuoiMvt(1) ;
   END ;
end;

procedure TQRBrouillard.BPiedPieceBeforePrint(var PrintBand: Boolean; var Quoi: string);
begin
TE_NUMEROPIECE_.Caption:=MsgBox.Mess[3]+'   '+IntToStr(QR2E_NUMEROPIECE.AsInteger) ;
TotDebitPiece.Caption:= AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, TotPiece[1].TotDebit,  CritEdt.AfficheSymbole) ;
TotCreditPiece.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, TotPiece[1].TotCredit, CritEdt.AfficheSymbole) ;
end;

procedure TQRBrouillard.BrouillardAfterPrint(BandPrinted: Boolean);
begin
TitreEntete.Visible:=False ;
TitreBarre.Visible:=False ;
end;

procedure TQRBrouillard.piedpageAfterPrint(BandPrinted: Boolean);
begin
TitreEntete.Visible:=True ; TitreBarre.Visible:=True ;
EnteteAutrePage.Enabled:=True ;
end;

procedure TQRBrouillard.FDateCreat1KeyPress(Sender: TObject; var Key: Char);
begin
ParamDate(Self, Sender, Key) ;
end;

procedure TQRBrouillard.FExerciceChange(Sender: TObject);
begin
//SG6 047/03/05 FQ 15466
if ObjFiltre.InChargement then Exit;
ExoToDates(FExercice.Value,FDateCompta1,FDatecompta2) ;

(* rony 27/05/97
if FExercice.ItemIndex<>0 then
   BEGIN
   // Rony 15/04/97 ExoToDates(FExercice.Value,FDateCreat1,FDateCreat2) ;
   ExoToDates(FExercice.Value,FDateCompta1,FDatecompta2) ;
   END ;*)
end;

procedure TQRBrouillard.FDevisesChange(Sender: TObject);
begin
if (FDevises.ItemIndex=0) or (FDevises.Value=V_PGI.DevisePivot) then
   BEGIN
   if FMontant.ItemIndex=1 then FMontant.ItemIndex:=Last ;
   Last:= FMontant.ItemIndex ;
   END ;
end;

procedure TQRBrouillard.RecupCritEdt ;
BEGIN
Fillchar(CritEdt,SizeOf(CritEdt),#0) ;
With CritEdt Do
  BEGIN
  Decimale:=V_PGI.OkDecV ; Symbole:=V_PGI.SymbolePivot ;
  DeviseSelect:='' ; if FDevises.ItemIndex<>0 then DeviseSelect:=FDevises.Value ;
  DeviseAffichee:=V_PGI.DevisePivot ;
  Monnaie:=FMontant.ItemIndex ;
  if Monnaie=1 Then DeviseAffichee:=DeviseSelect ;
  RappelCrit:=SansRappel ;
  Case FRappelCrit.State Of cbChecked : RappelCrit:=QueRappel ; cbGrayed : RappelCrit:=AvecRappel ; Else RappelCrit:=SansRappel ; END ;
  Bro.AvecJalAN:=FAvecJalAN.Checked ;
  Joker:=FJoker.Visible ;
  if Joker then
     BEGIN
     Cpt1:=FJoker.Text ; Cpt2:=FJoker.Text ; LCpt1:=Cpt1 ; LCpt2:=Cpt2 ;
     END else
     BEGIN
     Cpt1:=FCpte1.Text ; Cpt2:=FCpte2.Text ;
     PositionneFourchetteSt(FCpte1,FCpte2,CritEdt.LCpt1,CritEdt.LCpt2) ;
     END ;
  Date1:=StrToDate(FDateCompta1.Text) ; Date2:=StrToDate(FDateCompta2.Text) ;
  DateDeb:=DAte1 ; DateFin:=DAte2 ;
  Bro.RefInterne:=FRefInterne.Text ;
  If FNumPiece1.Text<>'' then Bro.NumPiece1:=StrToInt(FNumPiece1.Text) else Bro.NumPiece1:=0 ;
  If FNumPiece2.Text<>'' then Bro.NumPiece2:=StrToInt(FNumPiece2.Text) else Bro.NumPiece2:=999999999 ;
  if FEtab.ItemIndex<>0 then Etab:=FEtab.Value ;
  AfficheSymbole:=FMonetaire.Checked ;
  Bro.AfficheAnal:=FAfficheAnal.Checked ;
  Bro.Tri:=FTri.Checked ; Bro.TriJal:=FTriJal.Checked ;
  Bro.AfficheTauxEuro:=True ; // Modif SBO FQ 12878
  if FValide.State=cbGrayed then Valide:='g' ;
  if FValide.State=cbChecked then Valide:='X' ;
  if FValide.State=cbUnChecked then Valide:='-' ;
  MonoExo:=FExercice.ItemIndex>0 ;
  if MonoExo Then Exo.Code:=FExercice.Value ;
  QuelExoDate(Date1,Date2,MonoExo,Exo) ;
(* Rony 13/05/1997
  if Not MonoExo Then
     BEGIN
     if DansExo(V_PGI.Precedent,Date1,Date2) Then BEGIN MonoExo:=TRUE ; Exo:=V_PGI.Precedent ; END Else
        if DansExo(V_PGI.EnCours,Date1,Date2) Then BEGIN MonoExo:=TRUE ; Exo:=V_PGI.EnCours ; END Else
            if DansExo(V_PGI.Suivant,Date1,Date2) Then BEGIN MonoExo:=TRUE ; Exo:=V_PGI.Suivant ; END ;
     END ;
 *)
  With Bro.FormatPrint Do
     BEGIN
     PrSepMontant:=FTrait.Checked ;
     PrSepCompte[1]:=FLignePieceEntete.Checked ;
     PrSepCompte[2]:=FLignePiecePied.Checked ;
     Report.OkAff:=FReport.Checked ;
     END ;
  END ;
END ;

function  TQRBrouillard.CritEdtOk : Boolean ;
BEGIN
RecupCritEdt ;
Result:=CtrlPerExo(CritEDT.DateDeb,CritEDT.DateFin) ;
If Result Then
   BEGIN
   Result:=CritEdt.MonoExo ;
   Fillchar(TotGen, SizeOf(TotGen),#0) ;
   END ;
END ;

function  TQRBrouillard.AfficheMontant( Formatage, LeSymbole : String ; LeMontant : Double ; OkSymbole : Boolean ) : String ;
BEGIN
IF OkSymbole then AfficheMontant:=FormatFloat(Formatage+' '+LeSymbole,LeMontant)
             Else AfficheMontant:=FormatFloat(Formatage,LeMontant) ;
END ;

procedure TQRBrouillard.BFinEtatBeforePrint(var PrintBand: Boolean; var Quoi: string);
begin
TotDebitGen.Caption:= AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, TotGen[1].TotDebit,  CritEdt.AfficheSymbole) ;
TotCreditGen.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, TotGen[1].TotCredit, CritEdt.AfficheSymbole) ;
BOTTOMREPORT.enabled:=FALSE ;
end;

procedure TQRBrouillard.TOPREPORTBeforePrint(var PrintBand: Boolean; var Quoi: string);
begin
Titre1Rep.Caption:=Titre2Rep.Caption ;
Report1Debit.Caption:=Report2Debit.Caption ;
Report1Credit.Caption:=Report2Credit.Caption ;
end;

procedure TQRBrouillard.BOTTOMREPORTBeforePrint(var PrintBand: Boolean; var Quoi: string);
var MReport : TabTRep ;
begin
Case QuelReportBAL(CritEdt.Bro.FormatPrint.Report,MReport) of
  1 : Titre2Rep.Caption:=MsgBox.Mess[29] ;
  2 : Titre2Rep.Caption:=MsgBox.Mess[30]+' '+StReportPiece ;
 end ;
Report2Debit.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, MReport[1].TotDebit,  CritEdt.AfficheSymbole ) ;
Report2Credit.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, MReport[1].TotCredit, CritEdt.AfficheSymbole ) ;
end;


procedure TQRBrouillard.EnteteAutrePageAfterPrint(BandPrinted: Boolean);
begin
TOPREPORT.Enabled:=TRUE ;
end;

procedure TQRBrouillard.BEntetePieceBeforePrint(var PrintBand: Boolean; var Quoi: string);
Var St1, St2, St3 : String ;
    StTaux,StCotation : String ;
    MoIn : Boolean ;
begin
St1:='' ; St2:='' ; St3:='' ;
TE_NUMEROPIECE.Left:=TNumLigne.Left ;
TE_NUMEROPIECE.Width:=TNumLigne.Width+TCptGen.Width+1 ;
E_REFINTERNE2.Left:=TLibAux.Left ;
E_REFINTERNE2.Width:=TLibAux.Width+TE_REFINTERNE.Width+1 ;
E_LIBELLE2.Left:=TDateEcheance.Left ;
E_LIBELLE2.Width:=TDateEcheance.Width+TE_LIBELLE.Width+1 ;
TE_TAUXDEV.Left:=TModePaiement.Left ;
TE_TAUXDEV.Width:=TModePaiement.Width+TDateEcheance.Width+1 ;
MoIn:=FALSE ; If EstMonnaieIn(QR2E_DEVISE.AsString) Then MoIn:=TRUE ;
//Quoi:=Quoicpt(0) ;
TE_NUMEROPIECE.Caption:=MsgBox.Mess[2]+' '+IntToStr(QR2E_NUMEROPIECE.AsInteger)+'   '+QR2E_NATUREPIECE.AsString+'   '+DateToStr(QR2E_DATECOMPTABLE.AsDateTime) ;
St2:=QR2E_JOURNAL.AsString+'   '+RechDom('ttjournaux',QR2E_JOURNAL.AsString,False) ;
if QR2E_DEVISE.AsString<>V_PGI.DevisePivot then
   BEGIN
   StTaux:=formatFloat('##0.#####',QR2E_TAUXDEV.AsFloat) ;
   StCotation:=formatFloat('##0.#####',QR2E_COTATION.AsFloat) ;
   If Qr2E_DATECOMPTABLE.AsDateTime<V_PGI.DateDebutEuro Then St3:=StTaux+'   '+DateToStr(QR2E_DATETAUXDEV.AsDateTime) Else
     BEGIN
     If MoIn Then
       BEGIN
       If CritEdt.Bro.AfficheTauxEuro Then St3:=StCotation+'(Eur)' Else St3:=StTaux+'   ' ;
       END Else
       BEGIN
       If CritEdt.Bro.AfficheTauxEuro Then St3:=StCotation+'(Eur)'+DateToStr(QR2E_DATETAUXDEV.AsDateTime)
                                      Else St3:=StTaux+'   '+DateToStr(QR2E_DATETAUXDEV.AsDateTime) ;
       END ;
     END ;
   END Else
   BEGIN
   St3:='';
   END ;
TE_JOURNAL.Caption:=St2 ;
TE_DEVISE.Caption:=QR2E_DEVISE.AsString ;
TE_TAUXDEV.Caption:=St3 ;
TE_ETABLISSEMENT.Caption:=RechDom('ttEtablissement',QR2E_ETABLISSEMENT.AsString,False) ;

TE_DEVISE.Left:=TE_TAUXDEV.left-TE_DEVISE.Width-5 ;

TE_JOURNAL.Left:=TLibAux.Left ;
TE_JOURNAL.Width:=TE_DEVISE.Left-2 ;

E_REFINTERNE2.Caption:=MsgBox.Mess[16]+' : '+QR2E_REFINTERNE.AsString ;
E_LIBELLE2.Caption:=MsgBox.Mess[17]+' : '+QR2E_LIBELLE.AsString ;
E_REFINTERNE2.Visible:=HighBand ;
E_LIBELLE2.Visible:=HighBand ;
StReportPiece:=IntToStr(QR2E_NUMEROPIECE.AsInteger) ;
InitReport([2],CritEdt.Bro.FormatPrint.Report) ;
end;

procedure TQRBrouillard.QEcrAfterOpen(DataSet: TDataSet);
begin
QR2E_EXERCICE        :=TStringField(QEcr.FindField('E_EXERCICE')) ;
QR2E_JOURNAL         :=TStringField(QEcr.FindField('E_JOURNAL')) ;
QR2E_DATECOMPTABLE   :=TDateTimeField(QEcr.FindField('E_DATECOMPTABLE')) ;
QR2E_NUMEROPIECE     :=TIntegerField(QEcr.FindField('E_NUMEROPIECE')) ;
QR2E_NUMLIGNE        :=TIntegerField(QEcr.FindField('E_NUMLIGNE')) ;
QR2E_NUMECHE         :=TIntegerField(QEcr.FindField('E_NUMECHE')) ;
QR2E_ECHE            :=TIntegerField(QEcr.FindField('E_ECHE')) ;
QR2E_GENERAL         :=TStringField(QEcr.FindField('E_GENERAL')) ;
QR2E_AUXILIAIRE      :=TStringField(QEcr.FindField('E_AUXILIAIRE')) ;
QR2E_REFINTERNE      :=TStringField(QEcr.FindField('E_REFINTERNE')) ;
QR2E_LIBELLE         :=TStringField(QEcr.FindField('E_LIBELLE')) ;
QR2E_VALIDE          :=TStringField(QEcr.FindField('E_VALIDE')) ;
QR2E_NATUREPIECE     :=TStringField(QEcr.FindField('E_NATUREPIECE')) ;
QR2E_QUALIFPIECE     :=TStringField(QEcr.FindField('E_QUALIFPIECE')) ;
QR2E_TYPEMVT         :=TStringField(QEcr.FindField('E_TYPEMVT')) ;
QR2E_UTILISATEUR     :=TStringField(QEcr.FindField('E_UTILISATEUR')) ;
QR2E_DATECREATION    :=TDateTimeField(QEcr.FindField('E_DATECREATION')) ;
QR2E_ETABLISSEMENT   :=TStringField(QEcr.FindField('E_ETABLISSEMENT')) ;
QR2E_REGIMETVA       :=TStringField(QEcr.FindField('E_REGIMETVA')) ;
QR2E_TVA             :=TFloatField(QEcr.FindField('E_TVA')) ;
QR2E_TPF             :=TFloatField(QEcr.FindField('E_TPF')) ;
QR2E_MODEPAIE        :=TStringField(QEcr.FindField('E_MODEPAIE')) ;
QR2E_DATEECHEANCE    :=TDateTimeField(QEcr.FindField('E_DATEECHEANCE')) ;
QR2E_DEVISE          :=TStringField(QEcr.FindField('E_DEVISE')) ;
QR2E_TAUXDEV         :=TFloatField(QEcr.FindField('E_TAUXDEV')) ;
QR2E_COTATION        :=TFloatField(QEcr.FindField('E_COTATION')) ;
QR2E_DATETAUXDEV     :=TDateTimeField(QEcr.FindField('E_DATETAUXDEV')) ;
QR2E_QTE1            :=TFloatField(QEcr.FindField('E_QTE1')) ;
QR2E_QTE2            :=TFloatField(QEcr.FindField('E_QTE2')) ;
QR2E_QUALIFQTE1      :=TStringField(QEcr.FindField('E_QUALIFQTE1')) ;
QR2E_QUALIFQTE2      :=TStringField(QEcr.FindField('E_QUALIFQTE2')) ;
QR2T_LIBELLE         :=TStringField(QEcr.FindField('T_LIBELLE')) ;
QR2G_LIBELLE         :=TStringField(QEcr.FindField('G_LIBELLE')) ;
QR2E_ETATLETTRAGE    :=TStringField(QEcr.FindField('E_ETATLETTRAGE')) ;
QR2E_REFEXTERNE      :=TStringField(QEcr.FindField('E_REFEXTERNE')) ;
QR2E_REFLIBRE        :=TStringField(QEcr.FindField('E_REFLIBRE')) ;
QR2E_AFFAIRE         :=TStringField(QEcr.FindField('E_AFFAIRE')) ;
QR2E_DATEREFEXTERNE  :=TDateTimeField(QEcr.FindField('E_DATEREFEXTERNE')) ;
QR2MULTIECHE         :=TIntegerField(QEcr.FindField('MULTIECHE')) ;
QR2DEBIT             :=TFloatField(QEcr.FindField('DEBIT')) ;
QR2CREDIT            :=TFloatField(QEcr.FindField('CREDIT')) ;
ChgMaskChamp(Qr2DEBIT,CritEdt.Decimale,CritEdt.AfficheSymbole,CritEdt.Symbole,False) ;
ChgMaskChamp(Qr2CREDIT,CritEdt.Decimale,CritEdt.AfficheSymbole,CritEdt.Symbole,False) ;
end;

procedure TQRBrouillard.QAnaAfterOpen(DataSet: TDataSet);
begin
QR3Y_VALIDE         :=TStringField(QAna.FindField('Y_VALIDE')) ;
QR3Y_EXERCICE       :=TStringField(QAna.FindField('Y_EXERCICE')) ;
QR3Y_JOURNAL        :=TStringField(QAna.FindField('Y_JOURNAL')) ;
QR3Y_DATECOMPTABLE  :=TDateTimeField(QAna.FindField('Y_DATECOMPTABLE')) ;
QR3Y_NUMEROPIECE    :=TIntegerField(QAna.FindField('Y_NUMEROPIECE')) ;
QR3Y_NUMLIGNE       :=TIntegerField(QAna.FindField('Y_NUMLIGNE')) ;
QR3Y_NUMVENTIL      :=TIntegerField(QAna.FindField('Y_NUMVENTIL')) ;
QR3Y_GENERAL        :=TStringField(QAna.FindField('Y_GENERAL')) ;
QR3Y_AXE            :=TStringField(QAna.FindField('Y_AXE')) ;
QR3Y_QUALIFQTE1     :=TStringField(QAna.FindField('Y_QUALIFQTE1')) ;
QR3Y_QUALIFQTE2     :=TStringField(QAna.FindField('Y_QUALIFQTE2')) ;
QR3Y_REFINTERNE     :=TStringField(QAna.FindField('Y_REFINTERNE')) ;
QR3Y_LIBELLE        :=TStringField(QAna.FindField('Y_LIBELLE')) ;
QR3Y_SECTION        :=TStringField(QAna.FindField('Y_SECTION')) ;
QR3Y_POURCENTQTE1   :=TFloatField(QAna.FindField('Y_POURCENTQTE1')) ;
QR3Y_POURCENTQTE2   :=TFloatField(QAna.FindField('Y_POURCENTQTE2')) ;
QR3Y_POURCENTAGE    :=TFloatField(QAna.FindField('Y_POURCENTAGE')) ;
QR3Y_QTE1           :=TFloatField(QAna.FindField('Y_QTE1')) ;
QR3Y_QTE2           :=TFloatField(QAna.FindField('Y_QTE2')) ;
QR3S_ABREGE         :=TStringField(QAna.FindField('S_ABREGE')) ;
QR3Y_AFFAIRE        :=TStringField(QAna.FindField('Y_AFFAIRE')) ;
QR3Y_REFEXTERNE     :=TStringField(QAna.FindField('Y_REFEXTERNE')) ;
QR3Y_DATEREFEXTERNE :=TDateTimeField(QAna.FindField('Y_DATEREFEXTERNE')) ;
QR3Y_REFLIBRE       :=TStringField(QAna.FindField('Y_REFLIBRE')) ;
QR3DEBIT            :=TFloatField(QAna.FindField('YDEBIT')) ;
QR3CREDIT           :=TFloatField(QAna.FindField('YCREDIT')) ;
ChgMaskChamp(QR3DEBIT,CritEdt.Decimale,CritEdt.AfficheSymbole,CritEdt.Symbole,False) ;
ChgMaskChamp(QR3CREDIT,CritEdt.Decimale,CritEdt.AfficheSymbole,CritEdt.Symbole,False) ;
end;

procedure TQRBrouillard.QAnaBeforeOpen(DataSet: TDataSet);
begin
QAna.PArams[0].AsString:=QEcr.FindField('E_JOURNAL').AsString ;
QAna.PArams[1].AsString:=QEcr.FindField('E_EXERCICE').AsString ;
QAna.PArams[2].AsDateTime:=QEcr.FindField('E_DATECOMPTABLE').AsDateTime ;
QAna.PArams[3].AsInteger:=QEcr.FindField('E_NUMEROPIECE').AsInteger ;
QAna.PArams[4].AsInteger:=QEcr.FindField('E_NUMLIGNE').AsInteger ;
end;


procedure TQRBrouillard.FCpte1Change(Sender: TObject);
Var AvecJoker : Boolean ;
begin
AvecJoker:=Joker(FCpte1, FCpte2, FJoker) ;
TFaG.Visible:=Not AvecJoker ;
TFGen.Visible:=Not AvecJoker ;
TFGenJoker.Visible:=AvecJoker ;
end;

Function TQRBrouillard.RetourneCouple(A, B : String) : String ;
BEGIN
Result:='' ;
if (A<>'')and(B<>'') then Result:=A+' / '+B else
if (A<>'')and(B='') then Result:=A else
if (A='')and(B<>'') then Result:=B else Exit ;
END ;

procedure TQRBrouillard.QRDLEcheNeedData(var MoreData: Boolean);
Var Cod,Lib : String ;
    Tot : Array[0..12] of Double ;
    Quellerupt : integer ;
begin
MoreData:=PrintGroup(LEche, QEcr, [Qr2MULTIECHE], Cod, Lib, Tot,Quellerupt) ;
if MoreData then
   BEGIN
   if Tot[0]<>0 then TotDebitEche.Caption:=AfficheMontant(CritEDT.FormatMontant, CritEDT.Symbole,Tot[0], CritEDT.AfficheSymbole)
                Else TotDebitEche.Caption:='' ;
   if Tot[1]<>0 then TotCreditEche.Caption:=AfficheMontant(CritEDT.FormatMontant, CritEDT.Symbole,Tot[1],CritEDT.AfficheSymbole)
                Else TotCreditEche.Caption:='' ;
   END ;
end;

procedure TQRBrouillard.BEcheBeforePrint(var PrintBand: Boolean; var Quoi: string);
begin
if (AncienEche<=1) then PrintBand:=False else PrintBand:=True ;
end;

procedure TQRBrouillard.BPiedPieceAfterPrint(BandPrinted: Boolean);
begin
InitReport([2],CritEdt.Bro.FormatPrint.Report) ;
end;

procedure TQRBrouillard.FormActivate(Sender: TObject);
begin
If OkChargeCritMemoire Then SauveCritMemoire (QRCritMemoire,Pages);
OkChargeCritMemoire:=FALSE ;
end;

procedure TQRBrouillard.BAideClick(Sender: TObject);
begin
CallHelpTopic(Self) ;
end;

procedure TQRBrouillard.POPFPopup(Sender: TObject);
begin
UpdatePopFiltre(BSaveFiltre,BDelFiltre,BRenFiltre,FFiltres) ;
end;

procedure TQRBrouillard.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
if Parent is THPanel then Action:=caFree ;
// 14676
if Assigned(ObjFiltre) then FreeAndNil(ObjFiltre);
end;

procedure TQRBrouillard.FormCreate(Sender: TObject);
var
  Composants : TControlFiltre; // 14676
begin
HorzScrollBar.Range:=0 ; HorzScrollBar.Visible:=FALSE ;
VertScrollBar.Range:=0 ; VertScrollBar.Visible:=FALSE ;
Pages.Top:=0 ; Pages.Left:=0 ;
ClientHeight:=Pages.Top+Pages.Height+HPB.Height+PanelFiltre.Height-1 ;
ClientWidth:=Pages.Left+Pages.Width ;
if V_PGI.OutLook then Pages.Align:=alClient ;

// 14676
Composants.PopupF   := POPF;
Composants.Filtres  := FFILTRES;
Composants.Filtre   := BFILTRE;
Composants.PageCtrl := PAGES;
ObjFiltre := TObjFiltre.Create(Composants, '');
end;

procedure TQRBrouillard.FTriClick(Sender: TObject);
begin
If FTri.Checked Then FTriJal.Checked:=FALSE ;
end;

procedure TQRBrouillard.FTriJalClick(Sender: TObject);
begin
If FTriJal.Checked Then FTri.Checked:=FALSE ;
end;

procedure TQRBrouillard.QEcrCalcFields(DataSet: TDataSet);
begin
if CritEdt.Bro.Tri
  then QEcrCalcRupt.AsString:=DateToStr(QEcrE_DATECOMPTABLE.AsDateTime)+Format_String(QEcrE_JOURNAL.AsString,3)+FormatFloat('00000000',QEcrE_NUMEROPIECE.AsInteger)
  Else
if CritEdt.Bro.TriJal
  then QEcrCalcRupt.AsString:=Format_String(QEcrE_JOURNAL.AsString,3)+FormatFloat('00000000',QEcrE_NUMEROPIECE.AsInteger)+DateToStr(QEcrE_DATECOMPTABLE.AsDateTime)
  Else
       QEcrCalcRupt.AsString:=FormatFloat('00000000',QEcrE_NUMEROPIECE.AsInteger)+Format_String(QEcrE_JOURNAL.AsString,3)+DateToStr(QEcrE_DATECOMPTABLE.AsDateTime) ;
end;


procedure TQRBrouillard.BFermeClick(Sender: TObject);
begin
  {b FP FQ15554}
  Close;
  if IsInside(Self) then
  begin
    CloseInsidePanel(Self) ;
  end;
  {e FP FQ15554}
end;

end.

 
