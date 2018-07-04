unit QRSynPoi;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  QR, HSysMenu, Menus, hmsgbox, HQuickrp, DB, DBTables, StdCtrls, Buttons,
  Hctrls, ExtCtrls, Mask, Hcompte, ComCtrls, Ent1, Hent1, CritEdt, UtilEdt, HQry,
  CpteUtil,Choix, HTB97, HPanel, UiUtil ;

procedure EtatPointageSyn ;

{
  Fonction à revoir dès que la purge fonctionne : Les totaux pointé sur N-XX peuvent être
  détruits.
}

type
  TFSynEtatPointage = class(TFQR)
    HLabel3: THLabel;
    FRefP1: TEdit;
    HLabel9: THLabel;
    FRefP2: TEdit;
    FQueLesBanques: TCheckBox;
    FBanque: THValComboBox;
    TFBanque: THLabel;
    FLigneGenPied: TCheckBox;
    FRuptBanque: TCheckBox;
    FLigneBanque: TCheckBox;
    TRBanque: TQRLabel;
    RBanque: TQRLabel;
    QRLabel1: TQRLabel;
    RRefP1: TQRLabel;
    QRLabel13: TQRLabel;
    RRefP2: TQRLabel;
    TDateCompta: TQRLabel;
    TPiece: TQRLabel;
    TRefInterne: TQRLabel;
    TLibelle: TQRLabel;
    QRLabel2: TQRLabel;
    QRLabel19: TQRLabel;
    QRLabel20: TQRLabel;
    QRLabel6: TQRLabel;
    TDebit: TQRLabel;
    TCredit: TQRLabel;
    HeadBanque: TQRBand;
    QRLabel14: TQRLabel;
    TBQ_BANQUE: TQRLabel;
    TBQ_NOMBANQUE: TQRLabel;
    HeadGene: TQRBand;
    E_GENERAL: TQRDBText;
    QRLabel23: TQRLabel;
    G_LIBELLE: TQRDBText;
    RRib: TQRLabel;
    HeadPointage: TQRBand;
    QRLabel3: TQRLabel;
    E_REFPOINTAGE: TQRDBText;
    QRLabel8: TQRLabel;
    QRLabel9: TQRLabel;
    FSoldeD: TQRLabel;
    FSoldeC: TQRLabel;
    QRLabel10: TQRLabel;
    E_DATECOMPTABLE: TQRDBText;
    E_NUMEROPIECE: TQRDBText;
    E_REFINTERNE: TQRDBText;
    E_LIBELLE: TQRDBText;
    E_DATEECHEANCE: TQRDBText;
    E_MODEPAIE: TQRDBText;
    E_CONTREPARTIEGENE: TQRDBText;
    E_CONTREPARTIEAUXI: TQRDBText;
    E_DEBIT: TQRLabel;
    E_CREDIT: TQRLabel;
    FootPointage: TQRBand;
    TCREDIT_REFP: TQRLabel;
    TDEBIT_REFP: TQRLabel;
    TCOUNT_REFP: TQRLabel;
    TTotalPointage: TQRLabel;
    E_REFPOINTAGE2: TQRDBText;
    LigneTot1: TQRShape;
    FootGene: TQRBand;
    LigneTot2: TQRShape;
    TTotalPointageCpt: TQRLabel;
    E_GENERAL2: TQRDBText;
    TCOUNT_CPTE: TQRLabel;
    TDEBIT_CPTE: TQRLabel;
    TCREDIT_CPTE: TQRLabel;
    QBOverlay: TQRBand;
    Trait0: TQRLigne;
    FCol4: TQRLigne;
    Trait2: TQRLigne;
    Trait1: TQRLigne;
    REPORT2DEBIT: TQRLabel;
    REPORT2CREDIT: TQRLabel;
    REPORT1DEBIT: TQRLabel;
    REPORT1CREDIT: TQRLabel;
    GPointage: TQRGroup;
    GGene: TQRGroup;
    GBanque: TQRGroup;
    MsgBox: THMsgBox;
    Ligne1: TQRLigne;
    Trait3: TQRLigne;
    TRQueLesBanques: TQRLabel;
    RQueLesBanques: TQRLabel;
    FSautPage: TCheckBox;
    HSynTotPointeAvant: TQRLabel;
    FSynTotDebP: TQRLabel;
    FSynTotCreP: TQRLabel;
    HSynRefTotPointeAvant: TQRLabel;
    HSynDateTotPointeAvant: TQRLabel;
    QRLabel16: TQRLabel;
    QRLabel17: TQRLabel;
    QRLabel21: TQRLabel;
    FSynTotGDPointe: TQRLabel;
    FSynTotGCPointe: TQRLabel;
    FSynTotGSDPointe: TQRLabel;
    FSynTotGSCPointe: TQRLabel;
    FSynSoldeGD: TQRLabel;
    FSynSoldeGC: TQRLabel;
    GEcritures: TRadioGroup;
    procedure FRuptBanqueClick(Sender: TObject);
    procedure QAfterOpen(DataSet: TDataSet);
    procedure FormShow(Sender: TObject);
    procedure HeadGeneBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure HeadBanqueBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure FRefP1DblClick(Sender: TObject);
    procedure HeadPointageBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure FootGeneBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure FQueLesBanquesClick(Sender: TObject);
    procedure BDetailBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure FootPointageBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure TOPREPORTBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure BOTTOMREPORTBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure FootGeneAfterPrint(BandPrinted: Boolean);
    procedure FootPointageAfterPrint(BandPrinted: Boolean);
  private
    { Déclarations privées }
    QR2E_BANQUE, QR2E_DOMICILIATION, QR2E_ETABBQ, QR2E_REFPOINTAGE,
    QR2E_GUICHET, QR2E_NUMEROCOMPTE, QR2E_CLERIB, QR2E_MODEPAIE,
    QR2E_CODEBIC, QR2E_CONTREPARTIEAUXI, QR2E_CONTREPARTIEGENE,
    QR2E_JOURNAL,QR2E_EXERCICE,
    QR2E_GENERAL, QR2E_REFINTERNE, QR2E_LIBELLE, QR2G_LIBELLE      : TStringField ;
    QR2E_DATECOMPTABLE, QR2E_DATEECHEANCE, QR2E_DATEPOINTAGE       : TDateTimeField ;
    QR2E_NUMEROPIECE,QR2E_NUMLIGNE                                 : TIntegerField ;
    QR2DEBIT, QR2CREDIT                                            : TFloatField ;
    COUNT_REFP, COUNT_CPTE                                         : Integer ;
    DEBIT_REFP, CREDIT_REFP, DEBIT_CPTE, CREDIT_CPTE               : Double ;
    StReportGen,StReportRef                                        : string ;
    DateRefP,DateRefPMoins                                         : TDateTime ;
    RefP,RefPMoins                                                 : String ;
    TotDebitPMoins,TotCreditPMOins                                 : Double ;
    TotBqDebitMoins,TotBqCreditMoins,TotBqDebit,TotBqCredit      : Double ;
    procedure FinirPrint ; Override ;
    procedure GenereSQL ; Override ;
    procedure RenseigneCritere ; Override ;
    procedure ChoixEdition ; Override ;
    procedure InitDivers ; Override ;
    procedure RecupCritEdt ; Override ;
    function  CritOk : Boolean ; Override ;
    procedure InitSoldes(Gene, DateP, RefP: string);
    procedure InitSoldesSynapses ;
    Function QuoiGen : String ;
    Function QuoiMvt : String ;
    Procedure PointageZoom(Quoi : String) ;
  public
    { Déclarations publiques }
  end;


implementation

{$R *.DFM}

Uses SAISUTIL ;

Function TFSynEtatPointage.QuoiGen : String ;
BEGIN
Inherited ;
Result:=QR2E_GENERAL.AsString+'#'+Qr2G_LIBELLE.AsString+'@'+'1' ;
END ;

Function TFSynEtatPointage.QuoiMvt : String ;
BEGIN
Inherited ;
Result:=QR2E_GENERAL.AsString+' '+Qr2G_LIBELLE.AsString+' '+{Le_Solde.Caption+}
        '#'+QR2E_JOURNAL.AsString+' N° '+IntToStr(QR2E_NUMEROPIECE.AsInteger)+' '+DateToStr(QR2E_DateComptable.AsDAteTime)+'-'+
        PrintSolde(Qr2DEBIT.AsFloat,Qr2Credit.AsFloat,CritEDT.Decimale,CritEDT.Symbole,CritEDT.AfficheSymbole)+
       '@'+'5;'+QR2E_JOURNAL.AsString+';'+UsDateTime(QR2E_DATECOMPTABLE.AsDateTime)+';'+QR2E_NUMEROPIECE.AsString+';'+QR2E_EXERCICE.asString+';'+
        IntToStr(QR2E_NumLigne.AsInteger)+';' ;
END ;

Procedure TFSynEtatPointage.PointageZoom(Quoi : String) ;
Var Lp,i: Integer ;
BEGIN
Inherited ;
LP:=Pos('@',Quoi) ; If LP=0 Then Exit ;
i:=StrToInt(Copy(Quoi,LP+1,1)) ;
If (i=5) Then
   BEGIN
   Quoi:=Copy(Quoi,Lp+3,Length(Quoi)-lp-2) ;
   If QRP.QRPrinter.FSynShiftDblClick Then i:=6 ;
   END ;
ZoomEdt(i,Quoi) ;
END ;

procedure EtatPointageSyn ;
var QR: TFSynEtatPointage;
    Edition : TEdition ;
    PP : THPanel ;
BEGIN
QR:=TFSynEtatPointage.Create(Application) ;
Edition.Etat:=etSynPointage ;
QR.QRP.QRPrinter.OnSynZoom:=QR.PointageZoom ;
QR.InitType (nbGen,nePoi,msRien,'QRPOINT','',TRUE,FALSE,FALSE,Edition) ;
PP:=FindInsidePanel ;
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


procedure TFSynEtatPointage.GenereSQL ;
Var St,StV8 : String ;
BEGIN
InitSoldesSynapses ;
Q.Close; Q.SQL.Clear;
Q.SQL.Add(' Select G_LIBELLE, ') ;
if CritEdt.Poi.QueLesBque then
   BEGIN
   Q.SQL.Add(' BQ_BANQUE, BQ_DEVISE, BQ_DOMICILIATION,');
   Q.SQL.Add(' BQ_ETABBQ, BQ_GUICHET, BQ_NUMEROCOMPTE, BQ_CLERIB, BQ_CODEBIC,');
   END ;
Q.SQL.Add(' E_GENERAL, E_DATECOMPTABLE, E_NUMEROPIECE, E_LIBELLE, E_DATEECHEANCE, E_REFINTERNE, E_JOURNAL, E_EXERCICE,E_NUMLIGNE, ');
if CritEdt.DeviseSelect=V_PGI.DevisePivot then Q.SQL.Add(' E_DEBIT DEBIT, E_CREDIT CREDIT, ')
                                       else Q.SQL.Add(' E_DEBITDEV DEBIT, E_CREDITDEV CREDIT, ');
Q.SQL.Add(' E_REFPOINTAGE, E_DATEPOINTAGE, E_MODEPAIE, E_CONTREPARTIEAUX, E_CONTREPARTIEGEN');
if CritEdt.Poi.QueLesBque then
   BEGIN
   Q.SQL.Add(' From GENERAUX G, BANQUECP B, ECRITURE E');
   Q.SQL.Add(' Where G_NATUREGENE="BQE" and G_POINTABLE="X" and BQ_GENERAL=G.G_GENERAL and E_GENERAL=G.G_GENERAL');
   if CritEdt.Poi.Banque<>'' then Q.SQL.Add(' and BQ_BANQUE="'+CritEdt.Poi.Banque+'"');
   Q.SQL.Add(' and BQ_DEVISE="'+CritEdt.DeviseSelect+'"');
   END else
   BEGIN
   Q.SQL.Add(' From GENERAUX G, ECRITURE E');
   Q.SQL.Add(' Where G_NATUREGENE<>"BQE" and G_POINTABLE="X" and E_GENERAL=G.G_GENERAL');
   END ;
(*
if CritEdt.Joker then
   BEGIN
   Q.SQL.Add(' And E_GENERAL like "'+TraduitJoker(CritEdt.Cpt1)+'" ' )
   END Else
   BEGIN
   if CritEdt.Cpt1<>'' then Q.SQL.Add(' And E_GENERAL>="'+CritEdt.Cpt1+'" ') ;
   if CritEdt.Cpt2<>'' then Q.SQL.Add(' And E_GENERAL<="'+CritEdt.Cpt2+'" ') ;
   END ;
*)
if CritEdt.Cpt1<>'' then Q.SQL.Add(' And E_GENERAL="'+CritEdt.Cpt1+'" ') ;
Q.SQL.Add(' and E_QUALIFPIECE="N" and E_ECRANOUVEAU<>"OAN" and E_ECRANOUVEAU<>"C" ');
St:='' ;
Case CritEdt.Poi.TypSelectPointe Of
  0,2 : BEGIN
       St:=' and ((E_REFPOINTAGE<>"" and E_REFPOINTAGE="'+CritEdt.Poi.RefP1+'") Or ';
       St:=St+'(E_REFPOINTAGE="" AND E_DATECOMPTABLE<="'+UsDateTime(DateRefP)+'"' ;
       StV8:=LWhereV8 ;
       if StV8<>'' then St:=St+' AND '+StV8+')) ' Else St:=St+')) ' ;
       END ;
  1 : BEGIN
      St:=' and E_REFPOINTAGE<>"" and E_REFPOINTAGE="'+CritEdt.Poi.RefP1+'"';
      END ;
  END ;
Q.SQL.Add(St) ;
Q.SQL.Add(' Order by ');
if CritEdt.Poi.QueLesBque then  Q.SQL.Add(' BQ_BANQUE, ');
Q.SQL.Add(' E_GENERAL, E_DATEPOINTAGE, E_DATECOMPTABLE, E_NUMEROPIECE');
ChangeSQL(Q) ; Q.Open;
END;

function  TFSynEtatPointage.CritOk : Boolean ;
BEGIN
Result:=Inherited CritOK ;
If Result Then If FCpte1.text='' Then BEGIN Result:=FALSE ; NumErreurCrit:=7 ; END ;
If Result Then If FRefP1.text='' Then BEGIN Result:=FALSE ; NumErreurCrit:=8 ; END ;
END ;

Procedure TFSynEtatPointage.FinirPrint ;
BEGIN
InHerited ;
// ff
END ;

procedure TFSynEtatPointage.RecupCritEdt ;
BEGIN
Inherited ;
With CritEdt Do
  BEGIN
  Poi.Banque:='' ; if FBanque.ItemIndex<>0 then Poi.Banque:=FBanque.Value ;
  Poi.RefP1:=FRefP1.Text ; Poi.RefP2:=FRefP2.Text ;
  Poi.QueLesBque:=FQueLesBanques.Checked ;
  if FSautPage.State=cbChecked then SautPage:=1 ;
  if FSautPage.State=cbUnChecked then SautPage:=2 ;
  Poi.SurUnCompte:=TRUE ;
  Poi.TypSelectPointe:=GEcritures.ItemIndex ;
  With Poi.FormatPrint Do
    BEGIN
    PrSepCompte[1]:=FLigneBanque.Checked ;
    PrSepCompte[2]:=FLigneCpt.Checked ;
    PrSepCompte[3]:=FLigneGenPied.Checked ;
    END ;
  END ;
END;


procedure TFSynEtatPointage.FRuptBanqueClick(Sender: TObject);
begin
  inherited;
FLigneBanque.Checked:=false ;
FLigneBanque.Enabled:=(FRuptBanque.Checked);
end;

procedure TFSynEtatPointage.QAfterOpen(DataSet: TDataSet);
begin
  inherited;
if CritEdt.Poi.QueLesBque then
   BEGIN
   QR2E_BANQUE := TStringField(Q.FindField('BQ_BANQUE')) ;
   QR2E_DOMICILIATION := TStringField(Q.FindField('BQ_DOMICILIATION')) ;
   QR2E_ETABBQ := TStringField(Q.FindField('BQ_ETABBQ')) ;
   QR2E_GUICHET := TStringField(Q.FindField('BQ_GUICHET')) ;
   QR2E_NUMEROCOMPTE := TStringField(Q.FindField('BQ_NUMEROCOMPTE')) ;
   QR2E_CLERIB := TStringField(Q.FindField('BQ_CLERIB')) ;
   QR2E_CODEBIC := TStringField(Q.FindField('BQ_CODEBIC')) ;
   END ;
QR2E_CONTREPARTIEAUXI :=TStringField(Q.FindField('E_CONTREPARTIEAUX')) ;
QR2E_CONTREPARTIEGENE :=TStringField(Q.FindField('E_CONTREPARTIEGEN')) ;
QR2E_GENERAL          :=TStringField(Q.FindField('E_GENERAL')) ;
QR2E_JOURNAL          :=TStringField(Q.FindField('E_JOURNAL')) ;
QR2E_EXERCICE         :=TStringField(Q.FindField('E_EXERCICE')) ;
QR2E_DATECOMPTABLE    :=TDateTimeField(Q.FindField('E_DATECOMPTABLE')) ;
QR2E_NUMEROPIECE      :=TIntegerField(Q.FindField('E_NUMEROPIECE')) ;
QR2E_NUMLIGNE         :=TIntegerField(Q.FindField('E_NUMLIGNE')) ;
QR2E_REFINTERNE       :=TStringField(Q.FindField('E_REFINTERNE')) ;
QR2E_LIBELLE          :=TStringField(Q.FindField('E_LIBELLE')) ;
QR2G_LIBELLE          :=TStringField(Q.FindField('G_LIBELLE')) ;
QR2DEBIT              :=TFloatField(Q.FindField('DEBIT')) ;
QR2CREDIT             :=TFloatField(Q.FindField('CREDIT')) ;
QR2E_DATEECHEANCE     :=TDateTimeField(Q.FindField('E_DATEECHEANCE')) ;
QR2E_MODEPAIE         :=TStringField(Q.FindField('E_MODEPAIE')) ;
QR2E_DATEPOINTAGE     :=TDateTimeField(Q.FindField('E_DATEPOINTAGE')) ;
QR2E_REFPOINTAGE      :=TStringField(Q.FindField('E_REFPOINTAGE')) ;


//ChgMaskChamp(Qr2DEBIT,CritEdt.Decimale,CritEdt.AfficheSymbole,CritEdt.Symbole,False) ;
//ChgMaskChamp(Qr2CREDIT,CritEdt.Decimale,CritEdt.AfficheSymbole,CritEdt.Symbole,False) ;
end;

procedure TFSynEtatPointage.FormShow(Sender: TObject);
begin
// A cause du filtre par défaut
FBanque.ItemIndex:=0;
FDevises.Value:=V_PGI.DevisePivot ;
FDevises.Parent:=Standards ;
HLabel8.Parent:=Standards ;
GEcritures.ItemIndex:=1;
  inherited;
end;

procedure TFSynEtatPointage.RenseigneCritere;
BEGIN
Inherited ;
TRBanque.visible:=CritEdt.Poi.QueLesBque ; RBanque.visible:=CritEdt.Poi.QueLesBque ;
if CritEdt.Poi.QueLesBque then
   BEGIN
   RBanque.Caption:=CritEdt.Poi.Banque ;
   QRLabel18.Top:=57 ; RDevises.Top:=57 ;
   END Else
   BEGIN
   QRLabel18.Top:=TRBanque.Top ; RDevises.Top:=RBanque.Top ;
   END ;
RRefP1.Caption:=CritEdt.Poi.RefP1 ; RRefP2.Caption:=CritEdt.Poi.RefP2 ;
RBanque.Caption:=FBanque.Text ;
CaseACocher(FQueLesBanques,RQueLesBanques) ;
END;

procedure TFSynEtatPointage.InitDivers ;
BEGIN
Inherited ;
BFinEtat.Enabled:=False ; //Rony -- Pas de Bande de Total Edtition
{ Séparateurs de bandes }
HeadBanque.ForceNewPage:=(CritEdt.SautPage=1) ;
HeadBanque.Frame.DrawBottom:=CritEdt.Poi.FormatPrint.PrSepCompte[1] ;
HeadGene.Frame.DrawBottom:=CritEdt.Poi.FormatPrint.PrSepCompte[2] ;
FootGene.Frame.DrawBottom:=CritEdt.Poi.FormatPrint.PrSepCompte[3] ;
if CritEdt.Poi.QueLesBque then GBanque.DataField:='BQ_BANQUE' else GBanque.DataField:='';
END;


procedure TFSynEtatPointage.ChoixEdition ;
{ Initialisation des options d'édition }
BEGIN
Inherited ;
ChgMaskChamp(QR2DEBIT , CritEdt.Decimale, CritEdt.AfficheSymbole, CritEdt.Symbole, False) ;
ChgMaskChamp(QR2CREDIT, CritEdt.Decimale, CritEdt.AfficheSymbole, CritEdt.Symbole, False) ;
END ;

procedure TFSynEtatPointage.InitSoldes(Gene, DateP, RefP: string);
var S  : double;
    St : String;
    Q2  : TQuery;
BEGIN
FSoldeD.caption:='' ; FSoldeC.caption:='' ;
St:='Select EE_NEWSOLDECRE, EE_NEWSOLDEDEB From EEXBQ ' ;
St:=St+' Where EE_GENERAL="'+Gene+'" and EE_DATEPOINTAGE="'+DateP+'" and EE_REFPOINTAGE="'+RefP+'"';
Q2:=OpenSQL(St,true) ;
if Q2.EOF then BEGIN Ferme(Q2) ; exit ; END ;
S:=Q2.FindField('EE_NEWSOLDECRE').AsFloat - Q2.FindField('EE_NEWSOLDEDEB').AsFloat ;
Ferme(Q2) ;
St:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, ABS(S), CritEdt.AfficheSymbole ) ;
// affichage du solde bancaire dans lesens bancaire (D=déc et C=enc)
if S<0 then FSoldeD.caption:=St else FSoldeC.caption:=St ;
END;

procedure TFSynEtatPointage.InitSoldesSynapses ;
var St : String;
    Q2  : TQuery;
    OkP,OkPMoins,Inutile : Boolean ;
    CodeExo,CodeExoMoins : String ;
    ExoDateRefP : TExoDate ;
    SurUnExercice : Boolean ;
BEGIN
DAteRefP:=VH^.EnCours.Deb ; DateRefPMoins:=DateRefP ;
RefP:=CritEdt.Poi.RefP1 ; RefPMoins:='' ;
OkP:=FALSE ; OkPMoins:=FALSE ; SurUnExercice:=TRUE ;
TotDebitPMoins:=0 ; TotCreditPMoins:=0 ;
TotBqDebitMoins:=0 ; TotBqCreditMoins:=0 ; TotBqDebit:=0 ; TotBqCredit:=0 ;
{ 1 : Récupération de la référence de pointage }
St:='Select * From EEXBQ ' ;
St:=St+' Where EE_GENERAL="'+CritEdt.Cpt1+'" and EE_REFPOINTAGE="'+RefP+'" ';
//St:=St+' and EE_DATEPOINTAGE>="'+UsDateTime(V_PGI.EnCours.Deb)+'" ORDER BY EE_DATEPOINTAGE DESC';
St:=St+' ORDER BY EE_DATEPOINTAGE DESC';
Q2:=OpenSQL(St,true) ;
If Not Q2.Eof Then
   BEGIN
   While (Not Q2.Eof) And (Not OkP) Do
     BEGIN
     If (Not OkP) And (Q2.FindField('EE_REFPOINTAGE').AsString=CritEdt.Poi.RefP1) Then
        BEGIN
        DateRefP:=Q2.FindField('EE_DATEPOINTAGE').ASDateTime ; OkP:=TRUE ;
        TotBqDebit:=Q2.FindField('EE_NEWSOLDEDEB').AsFloat ;
        TotBqCredit:=Q2.FindField('EE_NEWSOLDECRE').AsFloat ;
        END ;
     END ;
   END ;
Ferme(Q2) ;

QuelExoDate(DateRefP,DateRefP,Inutile,ExoDateRefP) ;
If ExoDateRefP.Code=VH^.Suivant.Code Then ExoDateRefP:=VH^.EnCours ;

{ 2 : Récupération de la référence de pointage -1 }
St:='Select * From EEXBQ ' ;
St:=St+' Where EE_GENERAL="'+CritEdt.Cpt1+'" ' {+'" and EE_REFPOINTAGE="'+RefP+'" '};
//St:=St+' and EE_DATEPOINTAGE>="'+UsDateTime(ExoDateRefP.Deb)+'" AND EE_DATEPOINTAGE<"'+UsDateTime(DateRefP)+'" ' ;
St:=St+' AND EE_DATEPOINTAGE<"'+UsDateTime(DateRefP)+'" ' ;
St:=St+' ORDER BY EE_DATEPOINTAGE DESC';
Q2:=OpenSQL(St,true) ;
If Not Q2.Eof Then
   BEGIN
   While (Not Q2.Eof) And (Not OkPMoins) Do
     BEGIN
     DateRefPMoins:=Q2.FindField('EE_DATEPOINTAGE').ASDateTime ; OkPMoins:=TRUE ;
     RefPMoins:=Q2.FindField('EE_REFPOINTAGE').AsString ;
     TotBqDebitMoins:=Q2.FindField('EE_NEWSOLDEDEB').AsFloat ;
     TotBqCreditMoins:=Q2.FindField('EE_NEWSOLDECRE').AsFloat ;
     END ;
   END ;
Ferme(Q2) ;

{ 3 : Calcul du pointé depuis le début de l'exercice }

CodeExoMoins:=QuelExoDT(DateRefPMoins) ; CodeExo:=QuelExoDT(DateRefP) ;
If CodeExoMoins=VH^.Suivant.Code Then SurUnExercice:=FALSE ;
If SurUnExercice Then St:='SELECT SUM(E_DEBIT),SUM(E_CREDIT) FROM ECRITURE Where E_EXERCICE="'+VH^.EnCours.Code+'" AND E_GENERAL="'+CritEdt.Cpt1+'" AND '
                 Else St:='SELECT SUM(E_DEBIT),SUM(E_CREDIT) FROM ECRITURE Where (E_EXERCICE="'+VH^.EnCours.Code+'" OR E_EXERCICE="'+VH^.Suivant.Code+'") AND E_GENERAL="'+CritEdt.Cpt1+'" AND ' ;
St:=St+'E_REFPOINTAGE<>"" AND E_DATEPOINTAGE<="'+UsDateTime(DateRefPMoins)+'" AND E_QUALIFPIECE="N" AND  ' ;
St:=St+'E_DATECOMPTABLE>="'+UsDateTime(ExoDateRefP.Deb)+'" AND E_ECRANOUVEAU<>"C"' ;
Q2:=OpenSQL(St,true) ;
If Not Q2.Eof Then
   BEGIN
   TotDebitPMoins:=Q2.Fields[0].AsFloat ; TotCreditPMoins:=Q2.Fields[1].AsFloat ;
   END ;
Ferme(Q2) ;
(*
St:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, ABS(S), CritEdt.AfficheSymbole ) ;
// affichage du solde bancaire dans lesens bancaire (D=déc et C=enc)
if S<0 then FSoldeD.caption:=St else FSoldeC.caption:=St ;
*)
END;

procedure TFSynEtatPointage.HeadGeneBeforePrint(var PrintBand: Boolean;  var Quoi: string);
var St :string;
begin
  inherited;
RRIB.Caption:='';
Quoi:=QuoiGen ;
If CritEdt.Poi.QueLesBque then
   BEGIN
   St:=QR2E_ETABBQ.AsString+' '+ QR2E_GUICHET.AsString+' '+ QR2E_NUMEROCOMPTE.AsString+' '+QR2E_CLERIB.AsString;
   if St='' then St:=QR2E_CODEBIC.AsString;
   RRIB.Caption:=QR2E_DOMICILIATION.AsString+' - '+St;
   END ;

COUNT_CPTE:=0 ; DEBIT_CPTE:=0 ; CREDIT_CPTE:=0 ;
StReportGen:= Qr2E_GENERAL.AsString ;
InitReport([2],CritEdt.Poi.FormatPrint.Report) ;
end;


procedure TFSynEtatPointage.HeadBanqueBeforePrint(var PrintBand: Boolean;  var Quoi: string);
begin
  inherited;
if Not CritEdt.Poi.QueLesBque then BEGIN PrintBand:=False ; Exit ; END ;
if Not FRuptBanque.Checked then BEGIN PrintBand:=False ; Exit ; END ;
TBQ_BANQUE.caption:=QR2E_BANQUE.AsString;
TBQ_NOMBANQUE.caption:=RechDom('ttBanque',QR2E_BANQUE.AsString,false);
end;

procedure TFSynEtatPointage.FRefP1DblClick(Sender: TObject);
var St : string;
begin
  inherited;
St:='';
if FCpte1.Text<>'' then St:='EE_GENERAL="'+FCpte1.Text+'"';
St:=Choisir(MsgBox.Mess[1],'EEXBQ','EE_REFPOINTAGE','EE_REFPOINTAGE',St,'EE_REFPOINTAGE');
Screen.Cursor:=SyncrDefault ;
if St='' then exit;
TEdit(Sender).Text:=St;
end;

procedure TFSynEtatPointage.HeadPointageBeforePrint(var PrintBand: Boolean;  var Quoi: string);
var St : string;
    S : Double ;
begin
  inherited;
//InitSoldesSynapses ;

HSynRefTotPointeAvant.Caption:=RefPMoins ;
HSynDateTotPointeAvant.Caption:=DateToStr(DateRefPMoins) ;
FSoldeD.caption:='' ; FSoldeC.caption:='' ; FSynTotDebP.caption:='' ; FSynTotCreP.caption:='' ;

If Qr2E_REFPOINTAGE.AsString='' Then BEGIN PrintBand:=FALSE ; Exit ; END ;

// Affichage Total en banque sur Ref-1
S:=Arrondi(TotBqDebitMoins-TotBqCreditMoins,CritEdt.Decimale) ;
St:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, ABS(S), CritEdt.AfficheSymbole ) ;
if S>0 then FSoldeD.caption:=St else FSoldeC.caption:=St ;

// Affichage Total Pointé jusqu'à Ref-1
S:=Arrondi(TotDebitPMoins-TotCreditPMoins,CritEdt.Decimale) ;
St:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, ABS(S), CritEdt.AfficheSymbole ) ;
if S>0 then FSynTotDebP.caption:=St else FSynTotCreP.caption:=St ;

COUNT_REFP:=0 ; DEBIT_REFP:=0 ; CREDIT_REFP:=0 ;
StReportRef:= E_REFPOINTAGE.caption ;
InitReport([3],CritEdt.Poi.FormatPrint.Report) ;
end;

procedure TFSynEtatPointage.FootGeneBeforePrint(var PrintBand: Boolean;  var Quoi: string);
begin
  inherited;
Quoi:=QuoiGen ;
LigneTot2.Left:=TTotalPointageCpt.Left ;
LigneTot2.Width:=(FootGene.Width-TTotalPointageCpt.Left) ;
TCOUNT_CPTE.Caption:= IntToStr(COUNT_CPTE) ;
TCREDIT_CPTE.Caption:= AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, CREDIT_CPTE, CritEdt.AfficheSymbole ) ;
TDEBIT_CPTE.Caption:= AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, DEBIT_CPTE, CritEdt.AfficheSymbole ) ;
end;


procedure TFSynEtatPointage.FQueLesBanquesClick(Sender: TObject);
begin
  inherited;
TFBanque.Visible:=(FQueLesBanques.Checked) ;
FBanque.Visible:=(FQueLesBanques.Checked) ;
if FQueLesBanques.Checked then begin FRuptBanque.Checked:=False ; FRuptBanqueClick(Nil); end ;
FRuptBanque.Enabled:=(FQueLesBanques.Checked) ;
end;

procedure TFSynEtatPointage.BDetailBeforePrint(var PrintBand: Boolean;  var Quoi: string);
Var OkRep : Boolean ;
begin
  inherited;
Quoi:=QuoiMvt ;
OkRep:=TRUE ;
if QR2DEBIT.AsFloat=0 then E_DEBIT.Caption:=''
                      else E_DEBIT.Caption:= AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, QR2DEBIT.AsFloat, CritEdt.AfficheSymbole ) ;
if QR2CREDIT.AsFloat=0 then E_CREDIT.Caption:=''
                       else E_CREDIT.Caption:= AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, QR2CREDIT.AsFloat, CritEdt.AfficheSymbole ) ;
If QR2E_REFPOINTAGE.AsString<>'' Then
   BEGIN
   E_DEBIT.Font.style:=E_DEBIT.Font.style-[fsItalic] ;
   E_CREDIT.Font.style:=E_DEBIT.Font.style ;
   inc(COUNT_REFP)                           ; inc(COUNT_CPTE) ;
   DEBIT_REFP:=DEBIT_REFP + QR2DEBIT.AsFloat ; CREDIT_REFP:=CREDIT_REFP + QR2CREDIT.AsFloat ;
   DEBIT_CPTE:=DEBIT_CPTE + QR2DEBIT.AsFloat ; CREDIT_CPTE:=CREDIT_CPTE + QR2CREDIT.AsFloat ;
   END Else
   BEGIN
   E_DEBIT.Font.style:=E_DEBIT.Font.style+[fsItalic] ;
   E_CREDIT.Font.style:=E_DEBIT.Font.style ;
   END ;
If CritEdt.Poi.TypSelectPointe=2 Then If QR2E_REFPOINTAGE.AsString<>'' Then PrintBand:=FALSE ;
If QR2E_REFPOINTAGE.AsString='' Then OkRep:=FALSE ;
If OkRep Then
   AddReport([1,2,3],CritEdt.Poi.FormatPrint.Report,QR2DEBIT.AsFloat,QR2CREDIT.AsFloat,CritEdt.Decimale) ;
end;

procedure TFSynEtatPointage.FootPointageBeforePrint(var PrintBand: Boolean;  var Quoi: string);
Var D,C,M : Double ;
    St : String ;
begin
  inherited;
If Qr2E_REFPOINTAGE.AsString='' Then BEGIN PrintBand:=FALSE ; Exit ; END ;
LigneTot1.Left:=TTotalPointage.Left ;
LigneTot1.Width:=(FootPointage.Width-TTotalPointage.Left) ;
TCOUNT_REFP.Caption:= IntToStr(COUNT_REFP) ;
TDEBIT_REFP.Caption:= AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, DEBIT_REFP, CritEdt.AfficheSymbole ) ;
TCREDIT_REFP.Caption:= AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, CREDIT_REFP, CritEdt.AfficheSymbole ) ;
FSynTotGDPointe.Caption:=''  ; FSynTotGCPointe.Caption:='' ;
FSynTotGSDPointe.Caption:='' ; FSynTotGSCPointe.Caption:='' ;
FSynSoldeGD.Caption:=''      ; FSynSoldeGC.Caption:='' ;
D:=Arrondi(DEBIT_REFP+TotDebitPMoins,CritEdt.Decimale) ; C:=Arrondi(CREDIT_REFP+TotCreditPMoins,CritEdt.Decimale) ;
FSynTotGDPointe.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, D , CritEdt.AfficheSymbole ) ;
FSynTotGCPointe.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, C , CritEdt.AfficheSymbole ) ;
M:=Arrondi(D-C,CritEdt.Decimale) ;
St:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, ABS(M), CritEdt.AfficheSymbole ) ;
if M>0 then FSynTotGSDPointe.caption:=St else FSynTotGSCPointe.Caption:=St ;
FSynSoldeGD.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, TotBqDebit , CritEdt.AfficheSymbole ) ;
FSynSoldeGC.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, TotBqCredit , CritEdt.AfficheSymbole ) ;
end;

procedure TFSynEtatPointage.TOPREPORTBeforePrint(var PrintBand: Boolean;  var Quoi: string);
begin
  inherited;
TitreReportH.Caption:=TitreReportB.Caption ;
Report1Debit.Caption:=Report2Debit.Caption ;
Report1Credit.Caption:=Report2Credit.Caption ;
end;

procedure TFSynEtatPointage.BOTTOMREPORTBeforePrint(var PrintBand: Boolean;  var Quoi: string);
Var D,C : Double ;
    R   : integer ;
begin
  inherited;
R:=QuelReport(CritEdt.Poi.FormatPrint.Report,D,C) ;
Case R Of
  1 : TitreReportB.Caption:='' ;
  2 : TitreReportB.Caption:='REPORT COMPTE '+StReportGen ;
  3 : TitreReportB.Caption:='REPORT REFERENCE '+StReportRef ;
  END ;
if R<>1 then
   BEGIN
   Report2Debit.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, D, CritEdt.AfficheSymbole ) ;
   Report2Credit.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, C, CritEdt.AfficheSymbole ) ;
   END else
   BEGIN
   Report2Debit.Caption:='' ; Report2Credit.Caption:='' ;
   END ;
end;

procedure TFSynEtatPointage.FootGeneAfterPrint(BandPrinted: Boolean);
begin
  inherited;
InitReport([2],CritEdt.Poi.FormatPrint.Report) ;
end;

procedure TFSynEtatPointage.FootPointageAfterPrint(BandPrinted: Boolean);
begin
  inherited;
InitReport([3],CritEdt.Poi.FormatPrint.Report) ;
end;

end.
