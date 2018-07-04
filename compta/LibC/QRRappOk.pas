unit QRRappOk;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  QR, HSysMenu, Menus, hmsgbox, HQuickrp, DB, DBTables, StdCtrls, Buttons,
  Hctrls, ExtCtrls, Mask, Hcompte, ComCtrls, Ent1, Hent1, CritEdt, UtilEdt, HQry,
  CpteUtil, DBCtrls, HTB97, HPanel, UiUtil ;

procedure EtatRappro ;
procedure EtatRapproZoom(Crit : TCritEdt) ;

type
  TFEtatRappro = class(TFQR)
    Bevel3: TBevel;
    HLabel3: THLabel;
    GSens: TRadioGroup;
    TRIB: TQRLabel;
    RRib: TQRLabel;
    QRLabel1: TQRLabel;
    RReference: TQRLabel;
    FRefPointage: TLabel;
    TJournal: TQRLabel;
    TDateCompta: TQRLabel;
    TPiece: TQRLabel;
    TRefInterne: TQRLabel;
    TDebit: TQRLabel;
    TCredit: TQRLabel;
    REPORT1DEBIT: TQRLabel;
    REPORT1CREDIT: TQRLabel;
    HeadGene: TQRBand;
    FSolde1: TQRLabel;
    FSolde1D: TQRLabel;
    FSolde1C: TQRLabel;
    E_JOURNAL: TQRDBText;
    E_DATECOMPTABLE: TQRDBText;
    E_NUMEROPIECE: TQRDBText;
    E_REFINTERNE: TQRDBText;
    E_DEBIT: TQRLabel;
    E_CREDIT: TQRLabel;
    FootGene: TQRBand;
    FSolde2: TQRLabel;
    FSolde2D: TQRLabel;
    FSolde2C: TQRLabel;
    FSolde2R: TQRLabel;
    FSolde2RD: TQRLabel;
    FSolde2RC: TQRLabel;
    REPORT2DEBIT: TQRLabel;
    REPORT2CREDIT: TQRLabel;
    QBOverlay: TQRBand;
    Trait3: TQRLigne;
    Trait2: TQRLigne;
    Trait1: TQRLigne;
    Trait0: TQRLigne;
    QRef: TQuery;
    SRef: TDataSource;
    MsgBox: THMsgBox;
    GGene: TQRGroup;
    QRLabel2: TQRLabel;
    RSens: TQRLabel;
    FDateP: TDBLookupComboBox;
    Ligne1: TQRLigne;
    HLabel9: THLabel;
    QRLabel3: TQRLabel;
    E_DEVISE: TQRDBText;
    procedure FormShow(Sender: TObject);
    procedure FootGeneBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure FCpte1Exit(Sender: TObject);
    procedure FDatePClick(Sender: TObject);
    procedure BDetailBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure TOPREPORTBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure BOTTOMREPORTBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure HeadGeneBeforePrint(var PrintBand: Boolean; var Quoi: string);
    procedure QAfterOpen(DataSet: TDataSet);
    procedure FCpte1Enter(Sender: TObject);
    procedure FCpte1Change(Sender: TObject);
    procedure FFiltresChange(Sender: TObject);
    procedure BNouvRechClick(Sender: TObject);
  private
    { Déclarations privées }
    QR2E_GENERAL, QR2E_JOURNAL, QR2E_REFINTERNE, QR2E_EXERCICE : TStringField ;
    QR2E_DATECOMPTABLE           : TDateTimeField ;
    QR2E_NUMEROPIECE,QR2E_NUMLIGNE              : TIntegerField ;
    QR2DEBIT, QR2CREDIT                         : TFloatField ;
    SoldeD, SoldeC                              : Double;
    OldCpt : String ;
    GCalc1 : TGCalculCum ;
    SurExoSuivant : Boolean ;
    procedure FinirPrint ; Override ;
    procedure GenereSQL ; Override ;
    procedure RenseigneCritere ; Override ;
    procedure ChoixEdition ; Override ;
    procedure InitDivers ; Override ;
    procedure RecupCritEdt ; Override ;
    function  CritOk : Boolean ; Override ;
    procedure InitSoldes;
    procedure CalcSoldes(LeSens : string) ;
    Function  DevDuCpte(Gene : string) : Boolean ;
    Function  QuoiMvt : String ;
    Procedure RapproZoom(Quoi : String) ;
  public
    { Déclarations publiques }
  end;

implementation

{$R *.DFM}

Function TFEtatRappro.QuoiMvt : String ;
BEGIN
Inherited ;
Result:=QR2E_GENERAL.AsString+' '+{Qr2G_LIBELLE.AsString+' '+Le_Solde.Caption+}
        '#'+QR2E_JOURNAL.AsString+' N° '+IntToStr(QR2E_NUMEROPIECE.AsInteger)+' '+DateToStr(QR2E_DateComptable.AsDAteTime)+'-'+
        PrintSolde(Qr2DEBIT.AsFloat,Qr2Credit.AsFloat,CritEDT.Decimale,CritEDT.Symbole,CritEDT.AfficheSymbole)+
       '@'+'5;'+QR2E_JOURNAL.AsString+';'+UsDateTime(QR2E_DATECOMPTABLE.AsDateTime)+';'+QR2E_NUMEROPIECE.AsString+';'+QR2E_EXERCICE.asString+';'+
        IntToStr(QR2E_NumLigne.AsInteger)+';' ;
END ;

Procedure TFEtatRappro.RapproZoom(Quoi : String) ;
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


procedure EtatRappro ;
var QR : TFEtatRappro;
    Edition : TEdition ;
    PP : THPanel ;
BEGIN
PP:=FindInsidePanel ;
QR:=TFEtatRappro.Create(Application) ;
Edition.Etat:=etRappro ;
QR.QRP.QRPrinter.OnSynZoom:=QR.RapproZoom ;
QR.InitType (nbGen,neRap,msRien,'QRRAPPRO','',TRUE,FALSE,FALSE,Edition) ;
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

procedure EtatRapproZoom(Crit : TCritEdt) ;
var QR : TFEtatRappro;
    Edition : TEdition ;
BEGIN
QR:=TFEtatRappro.Create(Application) ;
Edition.Etat:=etRappro ;
try
 QR.QRP.QRPrinter.OnSynZoom:=QR.RapproZoom ;
 QR.CritEDT:=Crit ;
 QR.InitType (nbGen,neRap,msRien,'QRRAPPRO','',FALSE,TRUE,FALSE,Edition) ;
 finally
 QR.Free ;
 end ;
Screen.Cursor:=SyncrDefault ;
END ;


procedure TFEtatRappro.FormShow(Sender: TObject);
begin
HelpContext:=7619010 ;
Standards.HelpContext:=7619010 ;
//Avances.HelpContext:=7619020 ;
Mise.HelpContext:=7619030 ;
Option.HelpContext:=7619020 ;

GSens.ItemIndex:=0;
FRefPointage.caption:='' ; OldCpt:='' ; FDateP.KeyValue:='' ;
ChangeSQL(QRef) ; //QRef.Prepare ;
PrepareSQLODBC(QRef) ;
  inherited;
if Not V_PGI.OutLook then
   BEGIN
   FDevises.Parent:=Standards ;
   HLabel8.Parent:=Standards ;
   END ; 
If FFiltres.Text<>'' then begin FCpte1Exit(Nil) ; FDatePClick(Nil) ; End ;
end;

procedure TFEtatRappro.RenseigneCritere;
var St,StSens : string;
    Q2 : TQuery; Ou : Integer ;
BEGIN
Inherited ;
St:='Select BQ_DOMICILIATION, BQ_ETABBQ, BQ_GUICHET, BQ_NUMEROCOMPTE, BQ_CLERIB, BQ_CODEBIC';
St:=St+' From BANQUECP  Where BQ_GENERAL="'+CritEdt.Cpt1+'" AND BQ_NODOSSIER="'+V_PGI.NoDossier+'"'; // 24/10/2006 YMO Multisociétés
Q2:=OpenSQL(St,true);
if Not Q2.EOF then
   BEGIN
   St:=Q2.Fields[0].AsString+' - '+Q2.Fields[1].AsString+' '+Q2.Fields[2].AsString+' ';
   St:=St+Q2.Fields[3].AsString+' '+Q2.Fields[4].AsString;
   if St='' then St:=Q2.Fields[5].AsString;
   RRib.Caption:=St;
   END else
   BEGIN
   RRIB.Visible:=false ; TRIB.Visible:=false ;
   END;
Ferme(Q2);
RReference.Caption:=CritEdt.Rap.RefP ;
// Rony 23/05/97 RSens.caption:=GSens.Items[CritEdt.Rap.Sens];
StSens:=GSens.Items[CritEdt.Rap.Sens] ;
Ou:=Pos('&',StSens) ; If Ou>0 Then StSens[Ou]:=' ' ;
RSens.caption:=StSens;
END;

procedure TFEtatRappro.InitDivers;
BEGIN
Inherited ;
BFinEtat.Enabled:=FALSE ;
END;

function  TFEtatRappro.CritOk : Boolean ;
Var SetTyp : SetttTypePiece  ;
BEGIN
Result:=Inherited CritOK ;
If Result Then
   BEGIN
   Result:=ExisteCpte(CritEdt.Cpt1,fbGene) ;
   if Not Result Then BEGIN MsgBox.Execute(1,'',''); NumErreurCrit:=1000 ; END ;
   If Result Then
      BEGIN
      Result:=CritEdt.Rap.RefP<>'';
      If Not Result then BEGIN MsgBox.Execute(7,'',''); NumErreurCrit:=1000 ; END ;
      END ;
   END ;
SurExoSuivant:=FALSE ; If CritEdt.Exo.Code=VH^.Suivant.Code Then SurExoSuivant:=TRUE ;
If Result Then
   BEGIN
   { GP 13/10/97
     Bug : Ne fonctionne pas si N et N+1 mouvementé ou so Date de référence sur N-X
   }
   SetTyp:=[tpReel] ;
   Gcalc:=TGCalculCum.create(Un,fbGene,fbGene,SetTyp,Dev,Etab,Exo,DevEnP,CritEdt.Monnaie=2,CritEdt.Decimale,V_PGI.OkDecE) ;
   GCalc.initCalcul(CritEdt.Cpt1,CritEdt.Cpt1,'','',CritEdt.DeviseSelect,'',CritEdt.Exo.Code,
                    CritEdt.DateDeb,CritEdt.DateFin,TRUE) ;
   If SurExoSuivant Then
      BEGIN
      Gcalc1:=TGCalculCum.create(Un,fbGene,fbGene,SetTyp,Dev,Etab,TRUE,DevEnP,CritEdt.Monnaie=2,CritEdt.Decimale,V_PGI.OkDecE) ;
      GCalc1.initCalcul(CritEdt.Cpt1,CritEdt.Cpt1,'','',CritEdt.DeviseSelect,'',VH^.EnCours.Code,
                        VH^.EnCours.Deb,VH^.EnCours.Fin,TRUE) ;
      END ;
   //Rony 15/07/97, Init des décimales dans choixediton  --- InitSoldes ;
   END ;
END ;

Procedure TFEtatRappro.FinirPrint ;
BEGIN
InHerited ;
//GCalc.Free ; GCalc1.Free ;
If GCalc<>NIL Then BEGIN GCalc.Free ; GCalc:=NIL ; END ;
If GCalc1<>NIL Then BEGIN GCalc1.Free ; GCalc1:=NIL ; END ;
END ;

procedure TFEtatRappro.GenereSQL;
Var StV8 : String ;
BEGIN
Inherited ;
Q.Close ; Q.SQL.Clear ;
Q.SQL.Add('Select E_GENERAL, E_JOURNAL, E_DATECOMPTABLE, E_NUMEROPIECE, E_REFINTERNE, E_EXERCICE, E_NUMLIGNE, E_DEVISE,');
if (CritEdt.DeviseSelect=V_PGI.DevisePivot) Or (CritEdt.DeviseSelect='') then
  BEGIN
  If CritEdt.Monnaie=0 Then Q.SQL.Add(' E_DEBIT as DEBIT, E_CREDIT as CREDIT ')
                       Else Q.SQL.Add(' E_DEBITEURO as DEBIT, E_CREDITEURO as CREDIT ') ;
  END else Q.SQL.Add(' E_DEBITDEV as DEBIT, E_CREDITDEV as CREDIT ');
Q.SQL.Add(' From GENERAUX G, ECRITURE E');
Q.SQL.Add(' Where G_POINTABLE="X" and E_GENERAL=G.G_GENERAL');
Q.SQL.Add(' and E_GENERAL="'+CritEdt.Cpt1+'" and E_EXERCICE<="'+CritEdt.Exo.Code+'" and ');
StV8:=LWhereV8 ;
if StV8<>'' then Q.SQL.Add(StV8+' AND ') ;
Q.SQL.Add(' E_DATECOMPTABLE<="'+USDATETIME(CritEdt.Date2)+'" and ');
Q.SQL.Add(' E_QUALIFPIECE="N" and E_ECRANOUVEAU<>"OAN" and E_ECRANOUVEAU<>"C" and ');
Q.SQL.Add(' ((E_REFPOINTAGE<>"" and E_DATEPOINTAGE>"'+USDATETIME(CritEdt.Date2)+'") or (E_REFPOINTAGE=""))');
If (CritEdt.DeviseSelect<>'') And (CritEdt.DeviseSelect<>V_PGI.DevisePivot) Then
  BEGIN
  Q.SQL.Add(' AND E_DEVISE="'+CritEdt.DeviseSelect+'"');
  END ;
Q.SQL.Add(' Order by E_DATECOMPTABLE, E_NUMEROPIECE');
ChangeSQL(Q) ; Q.Open;
END;

procedure TFEtatRappro.InitSoldes;
BEGIN
if CritEdt.Rap.Sens=0 then
   BEGIN
   FSolde1.Caption:=MsgBox.Mess[3];
   FSolde2.Caption:=MsgBox.Mess[4];
   FSolde2R.Caption:=MsgBox.Mess[8];
   CalcSoldes('B') ;
   END else
   BEGIN
   FSolde1.Caption:=MsgBox.Mess[5];
   FSolde2.Caption:=MsgBox.Mess[6];
   FSolde2R.Caption:=MsgBox.Mess[9];
   CalcSoldes('C') ;
   END;
END;


procedure TFEtatRappro.FootGeneBeforePrint(var PrintBand: Boolean;  var Quoi: string);
var SC,SD : string ;
begin
  inherited;
SC:='' ; SD:='' ; AfficheLeSolde(Formateur, SoldeD, SoldeC);
if Formateur.Debit then SD:= AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, Formateur.Value, CritEdt.AfficheSymbole )
                   else SC:= AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, Formateur.Value, CritEdt.AfficheSymbole ) ;
if GSens.ItemIndex=0 then  // affichage du solde à la banque théorique
   BEGIN
   FSolde2C.Caption:=SD ; FSolde2D.Caption:=SC ;
   END else                // affichage du solde comptable théorique
   BEGIN
   FSolde2C.Caption:=SC ; FSolde2D.Caption:=SD ;
   END ;
BOTTOMREPORT.Enabled:=false ;
end;


procedure TFEtatRappro.CalcSoldes(LeSens : string) ;
var Cb,Cc,Db,Dc     : string ;
    SCb,SCc,SDb,SDc,DCpt,CCpt,DCpt1,CCpt1 : double ;
    TotCpte,TotCpte1 : TabTot ;
BEGIN
(*
Function TFQR.AfficheMontant( Formatage, LeSymbole : String ; LeMontant : Double ; OkSymbole : Boolean ) : String ;
BEGIN
if OkSymbole then
   BEGIN
   if LeMontant=0 then AfficheMontant:=''
                  else AfficheMontant:=FormatFloat(Formatage+' '+LeSymbole,LeMontant) ;
   END else
   BEGIN
   if LeMontant=0 then AfficheMontant:=''
                  else AfficheMontant:=FormatFloat(Formatage,LeMontant) ;
   END ;
END ;

*)
FSolde1C.caption:='' ; FSolde1D.caption:='' ; FSolde2RC.caption:='' ; FSolde2RD.caption:='' ;
SoldeC:=0 ; SoldeD:=0 ; SCc:=0 ; SDc:=0 ;SCb:=0 ; SDb:=0 ;
// calcul du solde bancaire à la date d'arréte
if Not QRef.Active then BEGIN QRef.ParamByName('Gene').AsString:=CritEdt.Cpt1 ; QRef.Open ; END;
While Not QRef.EOF do
   BEGIN
   if QRef.FindField('EE_DATEPOINTAGE').AsDateTime=CritEdt.Date2 then
      BEGIN
      If ((VH^.TenueEuro) And (CritEdt.Monnaie=0)) Or ((Not VH^.TenueEuro) And (CritEdt.Monnaie=2)) Then
        BEGIN
        SCb:=QRef.FindField('EE_NEWSOLDECREEURO').AsFloat ;
        SDb:=QRef.FindField('EE_NEWSOLDEDEBEURO').AsFloat ;
        END Else
        BEGIN
        SCb:=QRef.FindField('EE_NEWSOLDECRE').AsFloat ;
        SDb:=QRef.FindField('EE_NEWSOLDEDEB').AsFloat ;
        END ;
      Cb:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, SCb, CritEdt.AfficheSymbole ) ;
      Db:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, SDb, CritEdt.AfficheSymbole ) ;
      Break ;
      END;
   QRef.Next;
   END;

{ GP 13/10/97
  Bug : Ne fonctionne pas si N et N+1 mouvementé ou so Date de référence sur N-X
}

DCpt:=0 ; CCpt:=0 ; DCpt1:=0 ; CCpt1:=0 ;
If SurExoSuivant Then { Calcul Cumul sur N/N+1 Et déduction systématique ANO provisoire N+1 }
   BEGIN
   GCalc.Calcul ; TotCpte:=GCalc.ExecCalc.TotCpt ; { N+1 }
   DCpt:=TotCpte[1].TotDebit ; CCpt:=TotCpte[1].TotCredit ;
   GCalc1.Calcul ; TotCpte1:=GCalc1.ExecCalc.TotCpt ; { N }
   DCpt1:=TotCpte1[0].TotDebit+TotCpte1[1].TotDebit ;
   CCpt1:=TotCpte1[0].TotCredit+TotCpte1[1].TotCredit ;
   END Else
   BEGIN
   GCalc.Calcul ; TotCpte:=GCalc.ExecCalc.TotCpt ;
   DCpt:=TotCpte[0].TotDebit+TotCpte[1].TotDebit ;
   CCpt:=TotCpte[0].TotCredit+TotCpte[1].TotCredit ;
   END ;
DCpt:=DCpt+DCpt1 ; CCpt:=CCpt+CCpt1 ;
AfficheLeSolde(Formateur,DCpt,CCpt);
if Formateur.Debit then SDc:=Formateur.Value else SCc:=Formateur.Value ;
Dc:= AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, SDc, CritEdt.AfficheSymbole ) ;
Cc:= AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, SCc, CritEdt.AfficheSymbole ) ;
// affichage des soldes
if LeSens='B' then
   BEGIN
   SoldeC:=SDb ; SoldeD:=SCb ;       // affichage du solde dans le sens bancaire, mais mémorisation pour calcul du solde cpta théorique dans le sens comptable !!!!
   //if SCb>=0 then FSolde1C.caption:=Cb else FSolde1D.caption:=Db ; //Rony 19/06/97
   if (SCb>=0)and(SCb>SDb) then FSolde1C.caption:=Cb else FSolde1D.caption:=Db ;
   if SCc>0 then FSolde2RC.caption:=Cc else FSolde2RD.caption:=Dc ;
   END else
   BEGIN
   SoldeC:=SCc ; SoldeD:=SDc ;
   if SCc>0 then FSolde1C.caption:=Cc else FSolde1D.caption:=Dc ;
   if SCb>0 then FSolde2RC.caption:=Cb else FSolde2RD.caption:=Db ;
   END ;
END;


procedure TFEtatRappro.BDetailBeforePrint(var PrintBand: Boolean;  var Quoi: string);
Var OnDebit : Boolean ;
begin
  inherited;
Quoi:=QuoiMvt ; OnDebit:=TRUE ;
if QR2DEBIT.AsFloat=0 then BEGIN E_DEBIT.Caption:='' ; OnDebit:=FALSE ; END
                      else E_DEBIT.Caption:= AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, QR2DEBIT.AsFloat, CritEdt.AfficheSymbole ) ;
if QR2CREDIT.AsFloat=0 then E_CREDIT.Caption:=''
                       else BEGIN E_CREDIT.Caption:= AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, QR2CREDIT.AsFloat, CritEdt.AfficheSymbole ) ; OnDebit:=FALSE ; END ;
if CritEdt.Rap.Sens=0 then
   BEGIN
   SoldeD:=SoldeD+QR2DEBIT.AsFloat ;
   SoldeC:=SoldeC+QR2CREDIT.AsFloat ;
   END else
   BEGIN
   If OnDebit Then
      BEGIN
      If SoldeD-QR2DEBIT.AsFloat>=0 Then SoldeD:=SoldeD-QR2DEBIT.AsFloat Else SoldeC:=SoldeC+QR2DEBIT.AsFloat ;
      END Else
      BEGIN
      If SoldeC-QR2CREDIT.AsFloat>=0 Then SoldeC:=SoldeC-QR2CREDIT.AsFloat Else SoldeD:=SoldeD+QR2CREDIT.AsFloat ;
      END ;
   END;
AddReport([1],CritEdt.Poi.FormatPrint.Report,QR2DEBIT.AsFloat,QR2CREDIT.AsFloat,CritEdt.Decimale) ;
end;

procedure TFEtatRappro.RecupCritEdt ;
BEGIN
Inherited ;
With CritEdt Do
  BEGIN
//  Rap.DateP:=StrToDate(FDateCompta2.Text) ;
//  Exo.Code:=QuelExo(FDateCompta2.Text) ;
  Rap.RefP:=FRefPointage.Caption ;
  Rap.Sens:=GSens.ItemIndex ;
  END ;
END ;

Function TFEtatRappro.DevDuCpte(Gene : string) : Boolean ;
var Q2 : TQuery ;
BEGIN
Result:=FALSE ;
FDevises.Text:=V_PGI.DevisePivot ;
Q2:=OpenSQL('Select BQ_DEVISE From BANQUECP Where BQ_GENERAL="'+Gene+ '" AND BQ_NODOSSIER="'+V_PGI.NoDossier+'"', true); // 24/10/2006 YMO Multisociétés
if Not Q2.EOF then
  BEGIN
  FDevises.Value:=Q2.Fields[0].AsString ;
  If (FDevises.Value=V_PGI.DevisePivot) Or (FDevises.Value=V_PGI.DeviseFongible) Then FDevises.ItemIndex:=0 ;  
  Result:=TRUE ;
  END ;
Ferme(Q2);
END ;

procedure TFEtatRappro.ChoixEdition ;
{ Initialisation des options d'édition }
BEGIN
Inherited ;
ChgMaskChamp(QR2DEBIT , CritEdt.Decimale, CritEdt.AfficheSymbole, CritEdt.Symbole, False) ;
ChgMaskChamp(QR2CREDIT, CritEdt.Decimale, CritEdt.AfficheSymbole, CritEdt.Symbole, False) ;
InitSoldes ;
END ;

procedure TFEtatRappro.TOPREPORTBeforePrint(var PrintBand: Boolean;  var Quoi: string);
begin
  inherited;
TitreReportH.Caption:=TitreReportB.Caption ;
Report1Debit.Caption:=Report2Debit.Caption ;
Report1Credit.Caption:=Report2Credit.Caption ;
end;

procedure TFEtatRappro.BOTTOMREPORTBeforePrint(var PrintBand: Boolean;  var Quoi: string);
Var D,C : Double ;
begin
  inherited;
QuelReport(CritEdt.Rap.FormatPrint.Report,D,C) ;
Report2Debit.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, D, CritEdt.AfficheSymbole ) ;
Report2Credit.Caption:=AfficheMontant(CritEdt.FormatMontant, CritEdt.Symbole, C, CritEdt.AfficheSymbole ) ;
end;


procedure TFEtatRappro.HeadGeneBeforePrint(var PrintBand: Boolean;  var Quoi: string);
begin
  inherited;
InitReport([1],CritEdt.Poi.FormatPrint.Report) ;
end;

procedure TFEtatRappro.QAfterOpen(DataSet: TDataSet);
begin
  inherited;
QR2E_GENERAL          :=TStringField(Q.FindField('E_GENERAL')) ;
QR2E_JOURNAL          :=TStringField(Q.FindField('E_JOURNAL')) ;
QR2E_EXERCICE         :=TStringField(Q.FindField('E_EXERCICE')) ;
QR2E_DATECOMPTABLE    :=TDateTimeField(Q.FindField('E_DATECOMPTABLE')) ;
QR2E_NUMEROPIECE      :=TIntegerField(Q.FindField('E_NUMEROPIECE')) ;
QR2E_NUMLIGNE         :=TIntegerField(Q.FindField('E_NUMLIGNE')) ;
QR2E_REFINTERNE       :=TStringField(Q.FindField('E_REFINTERNE')) ;
QR2DEBIT              :=TFloatField(Q.FindField('DEBIT')) ;
QR2CREDIT             :=TFloatField(Q.FindField('CREDIT')) ;
//ChgMaskChamp(Qr2DEBIT,CritEdt.Decimale,CritEdt.AfficheSymbole,CritEdt.Symbole,False) ;
//ChgMaskChamp(Qr2CREDIT,CritEdt.Decimale,CritEdt.AfficheSymbole,CritEdt.Symbole,False) ;
end;

procedure TFEtatRappro.FCpte1Enter(Sender: TObject);
begin
If OldCpt<>FCpte1.Text Then BEGIN FDateP.Keyvalue:='' ; FRefPointage.Caption:='' ; END ;
  inherited;
// Roro 2/6/97 If OldCpt<>FCpte1.Text Then BEGIN FDateP.Keyvalue:='' ; FRefPointage.Caption:='' ; END ;
end;

procedure TFEtatRappro.FCpte1Exit(Sender: TObject);
begin
If OldCpt<>FCpte1.Text Then
   BEGIN
   QRef.Close ; QRef.ParamByName('Gene').AsString := FCpte1.Text ; QRef.Open ;
   if Not QRef.EOF then FDateP.DropDownRows:=8 Else begin FDateP.DropDownRows:=0; FRefPointage.caption:='' end ;
(* GP le 15/06/99
   DevDuCpte(FCpte1.Text) ;
*)
   END ;
DevDuCpte(FCpte1.Text) ;
OldCpt:=FCpte1.Text ;
  inherited;
(*
If OldCpt<>FCpte1.Text Then
   BEGIN
   QRef.Close ; QRef.ParamByName('Gene').AsString := FCpte1.Text ; QRef.Open ;
   if Not QRef.EOF then FDateP.DropDownRows:=8 else begin FDateP.DropDownRows:=0; FRefPointage.caption:='' end ;
   DevDuCpte(FCpte1.Text) ;
   END ;
OldCpt:=FCpte1.Text ;
*)
end;

procedure TFEtatRappro.FDatePClick(Sender: TObject);
begin
If FDateP.Text<>'' then
   BEGIN
   FRefPointage.caption:=FDateP.KeyValue;
   FDateCompta2.Text:=FDateP.Text;
   END ;
  inherited;
(*
If FDateP.Text<>'' then
   BEGIN
   FRefPointage.caption:=FDateP.KeyValue;
   FDateCompta2.Text:=FDateP.Text;
   END ;
*)
end;

procedure TFEtatRappro.FCpte1Change(Sender: TObject);
begin
  inherited;
FDateP.Keyvalue:='' ; FRefPointage.Caption:='' ;
end;

procedure TFEtatRappro.FFiltresChange(Sender: TObject);
begin
  inherited;
If FFiltres.Text<>'' then begin FCpte1Exit(Nil) ; FDatePClick(Nil) ; End ;

end;

procedure TFEtatRappro.BNouvRechClick(Sender: TObject);
begin
  inherited;
FRefPointage.Caption:='' ;
end;

end.
