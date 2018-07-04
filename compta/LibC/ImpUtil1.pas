unit ImpUtil;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, Hctrls, Buttons, ExtCtrls, Ent1, HEnt1, Menus, DB, {$IFNDEF DBXPRESS}dbtables{$ELSE}uDbxDataSet{$ENDIF}, ed_tools,
  ComCtrls, FmtChoix,CPTESAV,MajTable, HStatus, HSysMenu, hmsgbox, IniFiles,ImpFic ;

procedure RecupDevise(Code : String3 ; var DecDev : Integer ; var Quotite : Double ; TDev : TList) ;
function ImporteLesEcritures(NomFic : String ; Fmt,QualifP : String3 ; Leq : String ; Pos,lgCptes : boolean ; Msg : THMsgBox ;
                             Var InfoImp : TInfoImport) : boolean ;
procedure VerifDoublons ;
Procedure LanceLettrage ;
procedure MajImpErr(ListeErreurs : TStringList ; LaTotale : boolean) ;
procedure IntegreEcr(FAss : TForm ; Fmt,Leq : String ; Var InfoImp : TInfoImport) ;
procedure MajSelection(OkRejetes : boolean) ;
function  MajValCombo(ttCorr: string ; var Where : String ; Val,Lib,Abr : String) : byte ;
function  RecupJoin(NomChp : String ; MultiTT : string ) : String ;
procedure ChargeDevEtSect ;
procedure MajIMPECR(ListeComptes : Array of TStringList) ;
//function  MajLettrageEtVentileGen(Compte : String17) : boolean ;
function  MajEtatLettrageTiers(Compte : String17) : boolean ;
Procedure LanceVerifCpta(EnBatch : boolean) ;
function  FirstModePaie(ModeRegle : String3) : String3 ;
procedure RempliListesNewCptes(Leq,Fmt,FileN,Prefixe : String ; var ListeComptes : Array of TStringList ; Msg : THMsgBox ; LAide : THLabel ) ;
function  MajIEGeneral(Compte : String17 ; Lettrable,Ventilable : boolean) : boolean ;
function  MajIETiers(Compte : String17 ; Lettrable : boolean) : boolean ;

//procedure AjouteReference(Gene : String17 ; StRef : String ; DD : TDateTime) ;

// Objets d'importation

type
  FMvtImport = ^TFMvtImport ;
  TFMvtImport = RECORD
//String
  IE_AFFAIRE,IE_ETATLETTRAGE,IE_LETTRAGE,IE_LETTREPOINTLCR,IE_LIBELLE,IE_REFEXTERNE,
  IE_REFINTERNE,IE_REFLIBRE,IE_REFPOINTAGE,IE_REFRELEVE,IE_RIB,IE_SECTION,
  IE_NUMPIECEINTERNE : String ;
  IE_AXE : String[2] ;
// Combos
  IE_JOURNAL,IE_ENCAISSEMENT,IE_ECRANOUVEAU,IE_ETABLISSEMENT,IE_FLAGECR,IE_MODEPAIE,IE_NATUREPIECE,
  IE_REGIMETVA,IE_QUALIFPIECE,IE_QUALIFQTE1,IE_QUALIFQTE2,IE_SOCIETE,IE_TPF,IE_TVA,
  IE_TYPEANOUVEAU,IE_TYPEECR,IE_TYPEMVT,IE_DEVISE : String3 ;
// String17
  IE_AUXILIAIRE,IE_GENERAL,IE_CONTREPARTIEAUX,IE_CONTREPARTIEGEN : String17 ;
// entiers
  IE_CHRONO : Integer ;
  IE_NUMECHE,IE_NUMLIGNE,IE_NUMPIECE,
  IE_NUMVENTIL : Integer ;
// Dates
  IE_DATECOMPTABLE,IE_DATEECHEANCE,IE_DATEPAQUETMAX,IE_DATEPAQUETMIN,
  IE_DATEPOINTAGE,IE_DATEREFEXTERNE,IE_DATERELANCE,IE_DATETAUXDEV,
  IE_DATEVALEUR,IE_ORIGINEPAIEMENT : TDateTime ;
// booleens
  IE_ECHE,IE_CONTROLE,IE_LETTRAGEDEV,IE_OKCONTROLE,
  IE_SELECTED,IE_TVAENCAISSEMENT,IE_TYPEANALYTIQUE,IE_VALIDE,IE_ANA,IE_INTEGRE  : String1 ;
// doubles
  IE_POURCENTAGE,IE_POURCENTQTE1,IE_POURCENTQTE2,IE_QTE1,IE_QTE2,
  IE_QUOTITE : Double ;
  IE_RELIQUATTVAENC,IE_TAUXDEV,IE_TOTALTVAENC : Double ;
  IE_DEBIT,IE_CREDIT,IE_CREDITDEV,IE_CREDITEURO,IE_COUVERTURE,
  IE_COUVERTUREDEV,IE_DEBITDEV,IE_DEBITEURO : Double ;
  END ;

type String5 = String[5] ;

type TFDevise = Class
  Code : String3 ;
  Decimale : Integer ;
  Quotite : Double ;
  TauxDev : Double ;
  END ;

type TAnal = RECORD
   Journal,NatPiece,QualP : String3 ;
   General : String17 ;
   NumP,Numl,Numv : Integer ;
   RefInt,Lib : String ;
   Debit,Credit  : Double ;
   END ;

// Objets d'intégration

type TTotLigne = RECORD
     TotEcr:Double ;
     TotDev:Double ;
     TotEuro:Double ;
     TotQte1:Double ;
     TotQte2:Double ;
     QualQte1:String3 ;
     QualQte2:String3 ;
     END ;

type Piece = RECORD
     Jal,Nat,QualP : String3 ;
     DateP : TDateTime ;
     NumP : Integer ;
     END ;

type TFDoublon = Class
     IMPORT : boolean ;
     JOURNAL : String3 ;
     DATEC : TDateTime ;
     NATUREPIECE : String3 ;
     NUMPIECE : Integer ;
     QUALIFPIECE : String3 ;
     END ;

type TFJal  = Class
     Nature : String3 ;
     Normal,Simul : String3 ;
     LastDate : TDateTime ;
     LastNum : Integer ;
     LesMontants : Array[1..4,1..2] of Double ;
     END ;

type TFTotCpte = Class
     Table : String[20] ;
     Nature,RegimeTva : String3 ;
     Pointable,Lettrable : String1 ;
     Axe : String[5] ;
     LastDate : TDateTime ;
     LastNum,LastNumL : Integer ;
     LesMontants : Array[1..5,1..2] of Double ;
     END ;

Type TFSection = Class
     Axe : String3 ;
     Section : String17 ;
     LastDate : TDateTime ;
     LastNum,LastNumL : Integer ;
     LesMontants : Array[1..5,1..2] of Double ;
     END ;

type TReference = Class
     Ref : String ;
     Date : TDateTime ;
     END ;

type TFttListe = Class
     TypeTable : String ;
     Code,Lib    : String ;
     END ;

var
    AnnuleImport : boolean ;

const SepLigneIE=Chr(165) ;
Const Imp_Methode1 : Boolean = TRUE ;
Const QAJParam : BOOLEAN = FALSE ;
implementation

uses SaisUtil,HCompte,VERCPTA,LettUtil,LetBatch,RapSuppr,NewCpte,FmtImpor,ImporFmt,UtilTrans ;

Const MaxCptesEnMemoire = 5000 ;

var

    // Variables pour l'importation
    MsgBox            : THMsgBox ;
//    QAjoute  : TQuery ;
    Prefixe,Format,QualifPiece: String3 ;
    Table,Lequel,FileName   : String ;
    Positifs,LgComptes,
    MvtAZero,OkRupt   : boolean ;
    NbCptes,NbLig,(*NbLigIntegre,*)NbLigEnCours,NumVentil : LongInt ;

    NumLigneGeneA0,NumLigneVentilA0,NumLigneEcheA0,
    NumLignePiece,
    NumEche(*,NbPiece*)   : Integer ;
    Pointable,OldQualP: String1 ;
    OldJal,OldNatP,
    OldModePaie       : String3 ;
    OldNumP,NumP,NumL : LongInt ;
    OldDateC,OldDateEche,DateDeb,DateFin : TDateTime ;
    OldAuxiliaire     : String17 ;
(*    TotDeb,TotCred    : Double ; *)
    Qte1,QualifQte1   : String ;
    Fichier   : TextFile ;
    SectAttente : Array[1..5] of String ;
    TDev      : TList ;
    TypeMvt   : String ;
    Jal,QualP,NatP : String3 ;
    DateC     : TDateTime ;
    NumPiece  : LongInt ;
    MvtImport,InitMvtImp : FMvtImport ;
    Erreur : boolean ;

        // Variables pour l'intégration
    FAssImp : TForm ;
    QFiches,QGen,QTiers,
    QE,QY,QImpEcr           : TQuery ;
    DebitMvt,CreditMvt      : Double ;

    OldNat,JalChange,
    RegimeTva,NatureCpte,
    NatureJal,Facturier     : String3 ;

    CpteEnCours,OldCpteEnCours : String17 ;
    TotLigne                : TTotLigne ;
    TTotalCpte,TListeJal,ListePointe    : TStringList ;
    TErreur,TLett           : TList ;
//    EnregMvt,
    TListeSection           : TList ;

Function NomFichierCompteRendu(var StName : String) : String ;
var IniFile       : TIniFile ;
    StPath : String ;
BEGIN
StPath:=ExtractFilePath(Application.ExeName) ;
IniFile:=TIniFile.Create(StPath+'IMPORT.INI') ;
StName:=IniFile.ReadString('IMPORT','Fichier','') ;
IniFile.Free ;
Result:=StPath+ChangeFileExt(ExtractFileName(StName),'.ERR') ;
END ;

procedure CompteRenduBatch(QuelleErreur : Integer ; Var InfoImp : TInfoImport) ;
var F   : TextFile ;
    Fic,StName : String ;
BEGIN
StName:='' ;
Fic:=NomFichierCompteRendu(StName) ;
AssignFile(F,Fic) ;
{$I-} Append (F) ; {$I+}
if IOResult<>0 then
  BEGIN
  {$I-} ReWrite (F) ; {$I+}
  if IOResult<>0 then Exit ;
  Writeln(F,'Compte rendu d''importation du '+FormatDateTime('dddd dd mmmm yyyy "à" hh:nn',Now)) ;
  Writeln(F,'Fichier : '+StName) ;
  Writeln(F,'') ;
  END ;
WriteLn(F,'') ; WriteLn(F,'') ;
Case QuelleErreur of
  0 : WriteLn(F,'Erreur. Veuillez vérifier la première ligne de votre fichier !') ;
  1 : BEGIN
      WriteLn(F,'Lignes importées    : '+IntToStr(InfoImp.NbLigIntegre)) ;
      WriteLn(F,'Ecritures importées : '+IntToStr(InfoImp.NbPiece)) ;
      if InfoImp.TotDeb=0  then WriteLn(F,'Total Débit         : '+FormatFloat('#,##0.00 '+V_PGI.SymbolePivot,InfoImp.TotDeb))
                   else WriteLn(F,'Total Débit         : '+AfficheMontant('#,##0.00',V_PGI.SymbolePivot,InfoImp.TotDeb,True)) ;
      if InfoImp.TotCred=0 then WriteLn(F,'Total Crédit        : '+FormatFloat('#,##0.00 '+V_PGI.SymbolePivot,InfoImp.TotCred))
                   else WriteLn(F,'Total Crédit        : '+AfficheMontant('#,##0.00',V_PGI.SymbolePivot,InfoImp.TotCred,True)) ;
      END ;
  2 : WriteLn(F,'Fichier non trouvé ou vide !') ;
  END ;
CloseFile(F) ;
END ;

function GoodSoc : byte ;
var St, Lib : String ;
BEGIN
if not FileExists(FileName) then BEGIN Result:=2 ; Exit ; END ;
AssignFile(Fichier,FileName) ;
{$I-} Reset (Fichier) ; {$I+}
if EOF(Fichier) then BEGIN Result:=2 ; CloseFile(Fichier); Exit ; END ;
Readln(Fichier,St) ;
Lib:=Trim(Copy(St,1,70)) ;
NbLig:=0 ;
While not EOF(Fichier) do
    BEGIN
    if (Format='EDI') and (Lequel='FBA')  then
      if (Copy(St,1,5)='01480') then NbCptes:=Round(Valeur(Trim(Copy(St,9,15)))) ;
    Inc(NbLig) ;
    Readln(Fichier,St) ;
    END ;
CloseFile(Fichier) ;
Result:=1 ;
END ;

(*
Function Format_Date_HAL(Dat : String) : TDateTime ;
var An : String ;
    Y,M,J : Word ;
BEGIN
Result:=iDate1900 ;
if Trim(Dat)='' then Exit ;
An:=Copy(Dat,5,4) ;
Y:=Round(Valeur(An)) ;
M:=Round(Valeur(Copy(Dat,3,2))) ;
J:=Round(Valeur(Copy(Dat,1,2))) ;
Result:=EncodeDate(Y,M,J) ;
END ;
*)
Function Format_Date(Dat : String) : TDateTime ;
var An : String ;
    Y,M,J : Word ;
BEGIN
Result:=iDate1900 ;
if Trim(Dat)='' then Exit ;
An:=Copy(Dat,5,2) ;
DecodeDate(Date,Y,M,J) ;
if ((Round(Valeur(An))>=0) and (Round(Valeur(An))<=80)) or (Y>=2000) then An:='20'+An
                                                           else An:='19'+An ;
Y:=Round(Valeur(An)) ;
M:=Round(Valeur(Copy(Dat,3,2))) ;
J:=Round(Valeur(Copy(Dat,1,2))) ;
Result:=EncodeDate(Y,M,J) ;
END ;

function Format_DateEDI(St : String) : TDateTime ;
var Y,M,J : Word ;
BEGIN
if St='' then St:='19000101' ;
Y:=Round(Valeur(Copy(St,1,4))) ;
M:=Round(Valeur(Copy(St,5,2))) ;
J:=Round(Valeur(Copy(St,7,2))) ;
Result:=EncodeDate(Y,M,J) ;
END ;

function FirstModePaie(ModeRegle : String3) : String3 ;
var Q : TQuery ;
BEGIN
Result:='' ; if (ModeRegle='') then Exit ;
Q:=OpenSQL('SELECT MP_MODEPAIE FROM MODEPAIE ORDER BY MP_MODEPAIE',True) ;
if not Q.Eof then Result:=Q.Fields[0].AsString ;
Ferme(Q) ;
Q:=OpenSQL('SELECT MR_MP1 FROM MODEREGL WHERE MR_MODEREGLE="'+ModeRegle+'"',True) ;
if not Q.Eof then Result:=Q.Fields[0].asString ;
Ferme(Q) ;
END ;

function TrouveModePaie(CategP : String3 ; Montant : Double ; OnRemplace : Boolean) : String3 ;
var Q : TQuery ;
BEGIN
Result:='' ;
Q:=OpenSQL('Select MP_MODEPAIE,MP_MONTANTMAX,MP_REMPLACEMAX,MP_CONDITION FROM MODEPAIE WHERE MP_CATEGORIE="'+Trim(CategP)+'" ORDER BY MP_MODEPAIE',True) ;
if Q.Eof then
  BEGIN
  Ferme(Q) ;
  Q:=OpenSQL('Select MP_MODEPAIE,MP_MONTANTMAX,MP_REMPLACEMAX,MP_CONDITION FROM MODEPAIE',True) ;
  END ;
While not Q.Eof do
  BEGIN
  if (Result='') then Result:=Q.Fields[0].AsString ;
  if OnRemplace And (Q.Fields[3].AsString='X') and (Q.Fields[1].AsFloat>0) and (Montant>Q.Fields[1].AsFloat) then BEGIN Result:=Q.Fields[2].AsString ; Break ; END ;
  Q.Next ;
  END ;
Ferme(Q) ;
END ;

function CHERCHEUNMODE( Mode : String3 ; OkTreso,OkV7 : boolean ) : String3 ;
BEGIN
// Modes de paiements connus;
if Mode='O' then CHERCHEUNMODE:='ESP' else
if Mode='C' then CHERCHEUNMODE:='CHQ' else
if ((not OkTreso) and (Not OkV7)) then if (Mode='C') or (Mode='U') then CHERCHEUNMODE:='CB' else
if Mode='V' then CHERCHEUNMODE:='VIR' else
if Mode='P' then CHERCHEUNMODE:='PRE' else
if Mode='T' then CHERCHEUNMODE:='TRD' else
if OkV7 then BEGIN if Mode='A' then CHERCHEUNMODE:='TRA' END else
if Mode='L' then CHERCHEUNMODE:='LCR' else
if Mode='B' then CHERCHEUNMODE:='BOR' else
if (Mode='S') Or (Mode='') then CHERCHEUNMODE:='CHQ' else // à revoir pour ces deux valeurs possibles !!
// Autres type de LCR...
if Mode='T' then CHERCHEUNMODE:='LCR' else CHERCHEUNMODE:='' ; ///type 'S'...
END ;

function EnRuptAscii(Var infoImp : TInfoImport) : Boolean ;
BEGIN
EnRuptAscii:=False ;
If (Trim(Jal)<>Trim(OldJal)) or (DateC<>OldDateC) or (NumPiece<>OldNumP)
   or (Trim(QualP)<>Trim(OldQualP)) or (Trim(NatP)<>Trim(OldNatP)) then
    BEGIN
    OldJal:=Jal ; OldDateC:=DateC ; OldNumP:=NumPiece ; OldQualP:=QualP ; OldNatP:=NatP ; MvtAZero:=False ;
    if (TypeMvt<>'L') then Inc(InfoImp.NbPiece) ;
    EnRuptAscii:=True ;
    END ;
END;

procedure InitRuptEche ;
BEGIN
NumEche:=0 ;
OldModePaie:='' ;
OldDateEche:=iDate1900 ;
OldAuxiliaire:='' ;
END ;

function TestBreak(Var infoImp : TInfoImport) : Boolean ;
var Ok:boolean ;
BEGIN
Result:=False ;
If Imp_Methode1 Then Application.ProcessMessages ;
if (Lequel='FBA') and (Format='EDI') then
  BEGIN
  if AnnuleImport and (MsgBox<>nil) then
    BEGIN
    if (MsgBox.Execute(39,'','')<>mryes) then AnnuleImport:=False ;
    Result:=AnnuleImport ;
    END ;
  Exit ;
  END ;
if (Format='RAP') or (Format='CRA') or (Trim(Format)='MP') then Ok:=True else Ok:=EnRuptAscii(InfoImp) ;
if Ok then
  BEGIN
  NumLignePiece:=0 ; NumVentil:=0 ; InitRuptEche ;
  if AnnuleImport and (MsgBox<>nil) then if (MsgBox.Execute(39,'','')<>mryes) then AnnuleImport:=False ;
  Result:=AnnuleImport ;
  END ;
END ;

procedure ChargeDevEtSect ;
var Q,Q1       : TQuery ;
    TDevise : TFDevise ;
    i : integer ;
BEGIN
i:=0 ;
Q:=OpenSQL('SELECT X_SECTIONATTENTE FROM AXE ORDER BY X_AXE',True) ;
While not Q.Eof do BEGIN inc(i) ; SectAttente[i]:=Q.Fields[0].AsString ; Q.Next ; END ;
Ferme(Q) ;
Q:=OpenSQL('SELECT D_DEVISE,D_QUOTITE,D_DECIMALE FROM DEVISE ORDER BY D_DEVISE',True) ;
Q1:=TQuery.Create(Application) ; Q1.DatabaseName:='SOC' ;
Q1.SQL.Add('SELECT H_TAUXREEL FROM CHANCELL WHERE H_DEVISE=:DEV'
          +' AND H_DATECOURS=(SELECT MAX(H_DATECOURS) FROM CHANCELL WHERE H_DEVISE=:DEV)') ;
ChangeSQL(Q1) ;Q1.Prepare ;
TDev:=TList.Create ;
While not Q.Eof do
  BEGIN
  TDevise:=TFDevise.Create ;
  TDevise.Code:=Q.Fields[0].AsString ;
  TDevise.Quotite:=Q.Fields[1].AsFloat ;
  TDevise.Decimale:=Q.Fields[2].AsInteger ;
  Q1.Params[0].asString:=Q.Fields[0].AsString ; ;
  Q1.Open ; TDevise.TauxDev:=Q1.Fields[0].AsFloat ; Q1.Close ;
  TDev.Add(TDevise) ;
  Q.Next ;
  END ;
Ferme(Q) ; Q1.Free ;
END ;

procedure RecupDevise(Code : String3 ; var DecDev : Integer ; var Quotite : Double ; TDev : TList) ;
var TDevise : TFDevise ;
    i : integer ;
BEGIN
if (Code='') or (Code=V_PGI.DevisePivot) then
  BEGIN
  DecDev:=V_PGI.OkDecV ; Quotite:=1 ;
  END else
  BEGIN
  for i:=0 to TDev.Count-1 do
    BEGIN
    TDevise:=TDev[i] ;
    if (TDevise.Code=Code) then
      BEGIN
      DecDev:=TDevise.Decimale ; Quotite:=TDevise.Quotite ;
      Break ;
      END ;
    END ;
  END ;
END ;

// RECORD des mouvements à importer ;
procedure InitMvtImport ;
BEGIN
With MvtImport^ do
  BEGIN
  IE_AFFAIRE:='';IE_ETATLETTRAGE:='RI'; IE_LETTRAGE:=''; IE_LETTREPOINTLCR:='';
  IE_LIBELLE:='';IE_REFEXTERNE:='';IE_REFINTERNE:='';IE_REFLIBRE:='';IE_REFPOINTAGE:='';
  IE_REFRELEVE:='';IE_RIB:='';IE_SECTION:='';IE_AXE:='';
  IE_JOURNAL:='';IE_ECRANOUVEAU:='N';IE_ETABLISSEMENT:=VH^.EtablisDefaut;IE_FLAGECR:='';IE_MODEPAIE:='';
  IE_NATUREPIECE:='';IE_REGIMETVA:='';IE_QUALIFPIECE:='N';IE_QUALIFQTE1:='';IE_QUALIFQTE2:='';
  IE_SOCIETE:=V_PGI.CodeSociete;IE_TPF:='';IE_TVA:='';IE_TYPEANOUVEAU:='';IE_TYPEECR:='';IE_TYPEMVT:='';
  IE_DEVISE:=V_PGI.DevisePivot;IE_AUXILIAIRE:='';IE_GENERAL:='';IE_CONTREPARTIEAUX:='';
  IE_CONTREPARTIEGEN:='';IE_NUMPIECEINTERNE:='' ; IE_NUMPIECE:=1 ;
  {IE_CHRONO : Integer ;
  IE_NUMLIGNE,IE_NUMPIECE,
  IE_DATECOMPTABLE,
  }
  IE_DATEECHEANCE:=iDate1900 ;IE_DATEPAQUETMAX:=iDate1900 ; IE_DATEPAQUETMIN:=iDate1900 ; IE_DATEPOINTAGE:=iDate1900 ;
  IE_DATEREFEXTERNE:=iDate1900 ; IE_DATERELANCE:=iDate1900 ; IE_DATETAUXDEV:=iDate1900 ; IE_DATEVALEUR:=iDate1900 ;
  IE_ORIGINEPAIEMENT:=iDate1900 ;

  IE_ECHE:='-' ;IE_ENCAISSEMENT:='RIE';IE_CONTROLE:='-' ;IE_LETTRAGEDEV:='-' ;
  IE_OKCONTROLE:='-';IE_SELECTED:='-' ; IE_TVAENCAISSEMENT:='-' ; IE_TYPEANALYTIQUE:='-' ;
  IE_VALIDE:='-' ; IE_ANA:='-' ;IE_INTEGRE:='-' ;

  IE_NUMECHE:=0 ;IE_NUMVENTIL:=0 ;
  IE_POURCENTAGE:=0 ; IE_POURCENTQTE1:=0 ; IE_POURCENTQTE2:=0 ; IE_QTE1:=0 ; IE_QTE2:=0 ;
  IE_QUOTITE:=1 ; IE_RELIQUATTVAENC:=0 ; IE_TAUXDEV:=1 ; IE_TOTALTVAENC:=0 ;
  IE_DEBIT:=0 ; IE_CREDIT:=0 ; IE_CREDITDEV:=0 ; IE_CREDITEURO:=0 ; IE_COUVERTURE:=0 ;
  IE_COUVERTUREDEV:=0 ; IE_DEBITDEV:=0 ; IE_DEBITEURO:=0 ;
  END ;
New(InitMvtImp) ; InitMvtImp^:=MvtImport^ ;
END ;

procedure PrepareQuery(QAjoute : TQuery) ;
BEGIN
if (Not Imp_Methode1) Or (QAJParam)  Then Exit ;
QAjoute.Close ; QAjoute.SQL.Clear ;
QAjoute.SQL.Add('INSERT INTO IMPECR (') ;

QAjoute.SQL.Add('IE_CHRONO,IE_AFFAIRE,IE_ETATLETTRAGE,IE_LETTRAGE,IE_LETTREPOINTLCR,IE_LIBELLE,') ;
QAjoute.SQL.Add('IE_REFEXTERNE,IE_REFINTERNE,IE_REFLIBRE,IE_REFPOINTAGE,IE_REFRELEVE,') ;
QAjoute.SQL.Add('IE_RIB,IE_SECTION,IE_AXE,IE_JOURNAL,IE_ECRANOUVEAU,IE_ETABLISSEMENT,') ;

QAjoute.SQL.Add('IE_FLAGECR,IE_MODEPAIE,IE_NATUREPIECE,IE_REGIMETVA,IE_QUALIFPIECE,');
QAjoute.SQL.Add('IE_QUALIFQTE1,IE_QUALIFQTE2,IE_SOCIETE,IE_TPF,IE_TVA,IE_TYPEANOUVEAU,');
QAjoute.SQL.Add('IE_TYPEECR,IE_TYPEMVT,IE_DEVISE,IE_AUXILIAIRE,IE_GENERAL,IE_CONTREPARTIEAUX,') ;

QAjoute.SQL.Add('IE_CONTREPARTIEGEN,IE_NUMECHE,IE_NUMPIECEINTERNE,IE_NUMLIGNE,');
QAjoute.SQL.Add('IE_NUMPIECE,IE_NUMVENTIL,IE_DATECOMPTABLE,IE_DATEECHEANCE,IE_DATEPAQUETMAX,');
QAjoute.SQL.Add('IE_DATEPAQUETMIN,IE_DATEPOINTAGE,IE_DATEREFEXTERNE,IE_DATERELANCE,') ;

QAjoute.SQL.Add('IE_DATETAUXDEV,IE_DATEVALEUR,IE_ORIGINEPAIEMENT,IE_ECHE,IE_ENCAISSEMENT,');
QAjoute.SQL.Add('IE_CONTROLE,IE_LETTRAGEDEV,IE_OKCONTROLE,IE_SELECTED,IE_INTEGRE,IE_TVAENCAISSEMENT,');
QAjoute.SQL.Add('IE_TYPEANALYTIQUE,IE_VALIDE,IE_ANA,IE_POURCENTAGE,IE_POURCENTQTE1,');

QAjoute.SQL.Add('IE_POURCENTQTE2,IE_QTE1,IE_QTE2,IE_QUOTITE,IE_RELIQUATTVAENC,IE_TAUXDEV,');
QAjoute.SQL.Add('IE_TOTALTVAENC,IE_DEBIT,IE_CREDIT,IE_CREDITDEV,IE_CREDITEURO,');
QAjoute.SQL.Add('IE_COUVERTURE,IE_COUVERTUREDEV,IE_DEBITDEV,IE_DEBITEURO') ;
END ;

procedure RemplirJalEtSect ;
var TJal    : TFJal ;
    i,j : integer ;
    TSection : TFSection ;
BEGIN
QFiches:=TQuery.Create(Application) ; QFiches.DataBaseName:='SOC' ;
QFiches.SQL.Clear ;
if (Prefixe='BE') then
  QFiches.SQL.Add('SELECT BJ_BUDJAL,BJ_NATJAL,BJ_COMPTEURNORMAL,BJ_COMPTEURSIMUL '+
                  'FROM BUDJAL ORDER BY BJ_BUDJAL')
  else
  QFiches.SQL.Add('SELECT J_JOURNAL,J_NATUREJAL,J_COMPTEURNORMAL,J_COMPTEURSIMUL,J_DATEDERNMVT,'+
                  'J_NUMDERNMVT,J_DEBITDERNMVT,J_CREDITDERNMVT,J_TOTALDEBIT,J_TOTALCREDIT,'+
                  'J_TOTDEBE,J_TOTCREE,J_TOTDEBS,J_TOTCRES FROM JOURNAL ORDER BY J_JOURNAL') ;
ChangeSQL(QFiches) ; QFiches.Prepare ; QFiches.Open ;
TListeJal:=TStringList.Create ;
While not QFiches.Eof do
  BEGIN
  TJal:=TFJal.Create ;
  TJal.Nature:=QFiches.Fields[1].AsString ;
  TJal.Normal:=QFiches.Fields[2].AsString ;
  TJal.Simul:=QFiches.Fields[3].AsString ;
  if Prefixe<>'BE' then
    BEGIN
    TJal.LastDate:=QFiches.Fields[4].AsDateTime ;
    TJal.LastNum:=QFiches.Fields[5].AsInteger ;
    j:=5 ;
    for i:=1 to 4 do
      BEGIN
      Inc(j) ; TJal.LesMontants[i,1]:=QFiches.Fields[j].AsFloat ;
      Inc(j) ; TJal.LesMontants[i,2]:=QFiches.Fields[j].AsFloat ;
      END ;
    END ;
  TListeJal.AddObject(Trim(QFiches.Fields[0].AsString),TJal) ;
  QFiches.Next ;
  END ;
QFiches.Close ;

QFiches.SQL.Clear ;
if Prefixe<>'BE' then
  QFiches.SQL.Add('SELECT S_SECTION,S_AXE,S_DATEDERNMVT,'+
                  'S_NUMDERNMVT,S_LIGNEDERNMVT,S_DEBITDERNMVT,S_CREDITDERNMVT,S_TOTALDEBIT,S_TOTALCREDIT,'+
                  'S_TOTDEBE,S_TOTCREE,S_TOTDEBS,S_TOTCRES,S_TOTDEBANO,S_TOTCREANO FROM SECTION ORDER BY S_SECTION,S_AXE')
  else
  QFiches.SQL.Add('SELECT BS_BUDSECT,BS_AXE FROM BUDSECT ORDER BY BS_BUDSECT,BS_AXE') ;

ChangeSQL(QFiches) ; QFiches.Prepare ; QFiches.Open ;
TListeSection:=TList.Create ;
While not QFiches.Eof do
  BEGIN
  TSection:=TFSection.Create ;
  TSection.Section:=QFiches.Fields[0].AsString ;
  TSection.Axe:=QFiches.Fields[1].AsString ;
  if Prefixe<>'BE' then
    BEGIN
    TSection.LastDate:=QFiches.Fields[2].AsDateTime ;
    TSection.LastNum:=QFiches.Fields[3].AsInteger ;
    TSection.LastNumL:=QFiches.Fields[4].AsInteger ;
    j:=4 ;
    for i:=1 to 5 do
      BEGIN
      Inc(j) ; TSection.LesMontants[i,1]:=QFiches.Fields[j].AsFloat ;
      Inc(j) ; TSection.LesMontants[i,2]:=QFiches.Fields[j].AsFloat ;
      END ;
    END ;
  TListeSection.Add(TSection) ;
  QFiches.Next ;
  END ;
QFiches.Close ; QFiches.SQL.Clear ;
END ;

procedure RemplirLesListes ;
BEGIN
RemplirJalEtSect ;
if Lequel='FBE' then Exit ;
TTotalCpte:=TStringList.Create ;
QGen:=TQuery.Create(Application) ; QGen.DataBaseName:='SOC' ;
QGen.SQL.Add('SELECT G_NATUREGENE,G_REGIMETVA,G_DATEDERNMVT,G_NUMDERNMVT,'+
             'G_LIGNEDERNMVT,G_DEBITDERNMVT,G_CREDITDERNMVT,G_TOTALDEBIT,G_TOTALCREDIT,'+
             'G_TOTDEBE,G_TOTCREE,G_TOTDEBS,G_TOTCRES,G_TOTDEBANO,G_TOTCREANO,'+
             'G_VENTILABLE1,G_VENTILABLE2,G_VENTILABLE3,G_VENTILABLE4,G_VENTILABLE5,G_LETTRABLE,G_POINTABLE FROM GENERAUX') ;
QGen.SQL.Add('WHERE G_GENERAL=:CPTE') ;
ChangeSQL(QGen) ; QGen.Prepare ;
QTiers:=TQuery.Create(Application) ; QTiers.DataBaseName:='SOC' ;
QTiers.SQL.Add('SELECT T_NATUREAUXI,T_REGIMETVA,T_DATEDERNMVT,T_NUMDERNMVT,'+
               'T_LIGNEDERNMVT,T_DEBITDERNMVT,T_CREDITDERNMVT,T_TOTALDEBIT,T_TOTALCREDIT,'+
               'T_TOTDEBE,T_TOTCREE,T_TOTDEBS,T_TOTCRES,T_TOTDEBANO,T_TOTCREANO,T_LETTRABLE FROM TIERS') ;
QTiers.SQL.Add('WHERE T_AUXILIAIRE=:CPTE') ;
ChangeSQL(QTiers) ; QTiers.Prepare ;
END ;

procedure UpdateTotTables(Lequel : String) ;
var i,Nb          : integer ;
    TJal       : TFJal ;
    TTotCpte   : TFTotCpte ;
    TSection   : TFSection ;
    Nom,Pref : String ;
BEGIN
if (Format='RAP') or (Format='CRA') or (Format='MP') then Exit ;
if (Lequel='') or (lequel='Journaux') then
  BEGIN
  if FAssImp<>nil then THLabel(FAssImp.FindComponent('LAide')).Caption:=THMsgBox(FAssImp.FindComponent('HM')).Mess[20] ;
  If Imp_Methode1 Then Application.ProcessMessages ;
  Nb:=TListeJal.Count ;
  InitMove(Nb,'') ; MoveCur(False) ;
  for i:=0 to TListeJal.Count-1 do
    BEGIN
    TJal:=TFJal(TListeJal.Objects[i]) ;
    ExecuteSQL('UPDATE JOURNAL SET J_DATEDERNMVT="'+USDateTime(TJal.LastDate)+'",'
              +'J_NUMDERNMVT='+IntToStr(TJal.LastNum)+',J_DEBITDERNMVT='
              +StrFPoint(TJal.LesMontants[1,1])+',J_CREDITDERNMVT='
              +StrFPoint(TJal.LesMontants[1,2])+',J_TOTALDEBIT='+StrFPoint(TJal.LesMontants[2,1])
              +',J_TOTALCREDIT='+StrFPoint(TJal.LesMontants[2,2])+',J_TOTDEBE='+StrFPoint(TJal.LesMontants[3,1])
              +',J_TOTCREE='+StrFPoint(TJal.LesMontants[3,2])+',J_TOTDEBS='+StrFPoint(TJal.LesMontants[4,1])
              +',J_TOTCRES='+StrFPoint(TJal.LesMontants[4,2])+' WHERE J_JOURNAL="'+Trim(TListeJal[i])+'"') ;
    MoveCur(False) ;
    END ;
  FiniMove ; //FAssImp.LAide.Caption:='' ; Application.ProcessMessages ; //ActiveAttente(20,False) ;
  END ;
if (Lequel='') or (lequel='Comptes') then
  BEGIN
  //ActiveAttente(18,True) ;
  if FAssImp<>nil then THLabel(FAssImp.FindComponent('LAide')).Caption:=THMsgBox(FAssImp.FindComponent('HM')).Mess[18] ;
  If Imp_Methode1 Then Application.ProcessMessages ;
  Nb:=TTotalCpte.Count ;
  InitMove(Nb,'') ; MoveCur(False) ;
  for i:=0 to TTotalCpte.Count-1 do
    BEGIN
    TTotCpte:=TFTotCpte(TTotalCpte.Objects[i]) ;
    if (Trim(TTotCpte.Table)='GENERAUX') then
      BEGIN
      Nom:='GENERAL' ; Pref:='G_'
      END else
    if (Trim(TTotCpte.Table)='TIERS') then
      BEGIN
      Nom:='AUXILIAIRE' ; Pref:='T_' ;
      END else break ;

    ExecuteSQL('UPDATE '+Trim(TTotCpte.Table)+' SET '+Pref+'DATEDERNMVT="'+USDateTime(TTotCpte.LastDate)+'",'
              +Pref+'NUMDERNMVT='+IntToStr(TTotCpte.LastNum)+','+Pref+'LIGNEDERNMVT='
              +IntToStr(TTotCpte.LastNumL)+','+Pref+'DEBITDERNMVT='
              +StrFPoint(TTotCpte.LesMontants[1,1])+','+Pref+'CREDITDERNMVT='
              +StrFPoint(TTotCpte.LesMontants[1,2])+','+Pref+'TOTALDEBIT='+StrFPoint(TTotCpte.LesMontants[2,1])
              +','+Pref+'TOTALCREDIT='+StrFPoint(TTotCpte.LesMontants[2,2])+','+Pref+'TOTDEBE='+StrFPoint(TTotCpte.LesMontants[3,1])
              +','+Pref+'TOTCREE='+StrFPoint(TTotCpte.LesMontants[3,2])+','+Pref+'TOTDEBS='+StrFPoint(TTotCpte.LesMontants[4,1])+','
              +Pref+'TOTCRES='+StrFPoint(TTotCpte.LesMontants[4,2])+','
              +Pref+'TOTDEBANO='+StrFPoint(TTotCpte.LesMontants[5,1])+','
              +Pref+'TOTCREANO='+StrFPoint(TTotCpte.LesMontants[5,2])+' WHERE '+Pref+Nom+'="'+Trim(TTotalCpte[i])+'"') ;
    MoveCur(False) ;
    END ;
  FiniMove ; //FAssImp.LAide.Caption:='' ; Application.ProcessMessages ; //ActiveAttente(18,False) ;
  END ;
if (Lequel='') or (Lequel='Sections') then
  BEGIN
  //ActiveAttente(19,True) ;
  if FAssImp<>nil then THLabel(FAssImp.FindComponent('LAide')).Caption:=THMsgBox(FAssImp.FindComponent('HM')).Mess[19] ;
  If Imp_Methode1 Then Application.ProcessMessages ;
  Nb:=TListeSection.Count ;
  InitMove(Nb,'') ; MoveCur(False) ;
  for i:=0 to TListeSection.Count-1 do
    BEGIN
    TSection:=TFSection(TListeSection[i]) ;
    ExecuteSQL('UPDATE SECTION SET S_DATEDERNMVT="'+USDateTime(TSection.LastDate)+'",'
              +'S_NUMDERNMVT='+IntToStr(TSection.LastNum)+',S_LIGNEDERNMVT='
              +IntToStr(TSection.LastNumL)+',S_DEBITDERNMVT='
              +StrFPoint(TSection.LesMontants[1,1])+',S_CREDITDERNMVT='
              +StrFPoint(TSection.LesMontants[1,2])+',S_TOTALDEBIT='+StrFPoint(TSection.LesMontants[2,1])
              +',S_TOTALCREDIT='+StrFPoint(TSection.LesMontants[2,2])+',S_TOTDEBE='+StrFPoint(TSection.LesMontants[3,1])
              +',S_TOTCREE='+StrFPoint(TSection.LesMontants[3,2])+',S_TOTDEBS='+StrFPoint(TSection.LesMontants[4,1])+','
              +'S_TOTCRES='+StrFPoint(TSection.LesMontants[4,2])+','
              +'S_TOTDEBANO='+StrFPoint(TSection.LesMontants[5,1])+','
              +'S_TOTCREANO='+StrFPoint(TSection.LesMontants[5,2])
              +' WHERE S_AXE="'+Trim(TSection.Axe)+'" AND S_SECTION="'+Trim(TSection.Section)+'"') ;
    MoveCur(False) ;
    END ;
  FiniMove ;
  if FAssImp<>nil then THLabel(FAssImp.FindComponent('LAide')).Caption:='' ;
  If Imp_Methode1 Then Application.ProcessMessages ; //ActiveAttente(19,False) ;
  END ;
END ;

procedure RecupInfosCpte(CpteGene,CpteAuxi : String17) ;
var TTotCpte : TFTotCpte ;
    i,j : integer ;
    Tabl : String[20] ;
BEGIN
if Lequel='FBE' then Exit ;
if (CpteAuxi='') then BEGIN Tabl:='GENERAUX' ; CpteEnCours:=CpteGene ; QFiches:=QGen ; END
                 else BEGIN RecupInfosCpte(CpteGene,'') ; Tabl:='TIERS' ; CpteEnCours:=CpteAuxi ; QFiches:=QTiers ; END ;
if (CpteEnCours=OldCpteEnCours) then Exit else OldCpteEnCours:=CpteEnCours ;
Pointable:='-' ;
if TTotalCpte<>nil then
  BEGIN
  if (TTotalCpte.Count>MaxCptesEnMemoire) then
    BEGIN
    UpdateTotTables('Comptes') ;
    TTotalCpte.Clear ;
    END else
    BEGIN
    i:=TTotalCpte.IndexOf(Trim(CpteEnCours)) ;
    if i>-1 then
      BEGIN
      TTotCpte:=TFTotCpte(TTotalCpte.Objects[i]) ;
      NatureCpte:=TTotCpte.Nature ;
      RegimeTva:=TTotCpte.RegimeTva ;
      if QFiches=QGen then Pointable:=TTotCpte.Pointable ;
      Exit ;
      END ;
    END ;
  END ;

QFiches.Params[0].AsString:=CpteEnCours ;
QFiches.Open ;
if QFiches.Eof then BEGIN QFiches.Close ; Exit ; END ;
NatureCpte:=QFiches.Fields[0].AsString ;
RegimeTva:=QFiches.Fields[1].AsString ;

TTotCpte:=TFTotCpte.Create ;
TTotCpte.Table:=Tabl ;
TTotCpte.Nature:=NatureCpte ;
TTotCpte.RegimeTva:=RegimeTva ;
TTotCpte.LastDate:=QFiches.Fields[2].AsDatetime ;
TTotCpte.LastNum:=QFiches.Fields[3].AsInteger ;
TTotCpte.LastNumL:=QFiches.Fields[4].AsInteger ;
i:=4 ;
for j:=1 to 5 do
  BEGIN
  Inc(i) ; TTotCpte.LesMontants[j,1]:=QFiches.Fields[i].AsFloat ;
  Inc(i) ; TTotCpte.LesMontants[j,2]:=QFiches.Fields[i].AsFloat ;
  END ;

if QFiches=QGen then
  BEGIN
  TTotCpte.Axe:='-----' ;
  if QFiches.Fields[15].AsString='X' then TTotCpte.Axe[1]:='X' else
  if QFiches.Fields[16].AsString='X' then TTotCpte.Axe[2]:='X' else
  if QFiches.Fields[17].AsString='X' then TTotCpte.Axe[3]:='X' else
  if QFiches.Fields[18].AsString='X' then TTotCpte.Axe[4]:='X' else
  if QFiches.Fields[19].AsString='X' then TTotCpte.Axe[5]:='X' ;
  TTotCpte.Lettrable:=QFiches.Fields[20].AsString ;
  TTotCpte.Pointable:=QFiches.Fields[21].AsString ;
  Pointable:=TTotCpte.Pointable ;
  END else TTotCpte.Lettrable:=QFiches.Fields[15].AsString ;
TTotalCpte.AddObject(Trim(CpteEnCours),TTotCpte) ;
QFiches.Close ;
END ;

function NatureCompte(Compte : String17 ; OkGen : boolean ; var let : boolean) : String3 ;
var TCpte : TFTotCpte ;
    i : integer ;
BEGIN
Result:='' ; Let:=False ;
i:=TTotalCpte.IndexOf(Trim(Compte)) ;
if i >-1 then
  BEGIN
  TCpte:=TFTotCpte(TTotalCpte.Objects[i]) ;
  if TCpte<>nil then
    BEGIN
    Result:=TCpte.Nature ;
    Let:=(TCpte.Lettrable='X') ;
    END ;
  END ;
END ;

function  IsVentilable(CpteGene : String17 ; var Axe : String5 ; TypCpte : String1) : Boolean ;
var i : Integer ;
    TCpte : TFTotCpte ;
BEGIN
Result:=False ; Axe:='-----' ;
i:=TTotalCpte.IndexOf(Trim(CpteGene)) ;
if i>-1 then
  BEGIN
  TCpte:=TFTotCpte(TTotalCpte.Objects[i]) ;
  Axe:=TCpte.Axe ;
  if Pos('X',Axe)<>0 then Result:=True ;
  END else if (TypCpte='A') then BEGIN Result:=True ; Axe:='X----' ; END ;
END ;

procedure VideLettrage(Let : boolean) ;
BEGIN
With MvtImport^ do
  BEGIN
  IE_LETTRAGE:='';IE_LETTRAGEDEV:='-';IE_COUVERTURE:=0; IE_COUVERTUREDEV:=0;
  IE_DATEPAQUETMIN:=IDate1900 ; IE_DATEPAQUETMAX:=IDate1900 ;
  IE_ETATLETTRAGE:='RI' ;
  if Let then
    BEGIN
    IE_ETATLETTRAGE:='AL' ;
    IE_DATEPAQUETMIN:=IE_DATECOMPTABLE ; IE_DATEPAQUETMAX:=IE_DATECOMPTABLE ;
    if IE_DATEECHEANCE<IE_DATECOMPTABLE then IE_DATEECHEANCE:=IE_DATECOMPTABLE ;
    END ;
  END ;
END ;

procedure AjouteDevise(Code : String3) ;
var i : integer ;
    TDevise : TFDevise ;
BEGIN
With MvtImport^ do
  BEGIN
  if Code=V_PGI.DevisePivot then
    BEGIN
    IE_DEVISE:=V_PGI.DevisePivot;IE_TAUXDEV:=1;IE_DEBITDEV:=IE_DEBIT ;
    IE_CREDITDEV:=IE_CREDIT;IE_QUOTITE:=1 ;
    IE_DEBITEURO:=DeviseToEuro(IE_DEBITDEV,IE_TAUXDEV,IE_QUOTITE) ;
    IE_CREDITEURO:=DeviseToEuro(IE_CREDITDEV,IE_TAUXDEV,IE_QUOTITE) ;
    END else
    BEGIN
    for i:=0 to TDev.Count-1 do
      BEGIN
      TDevise:=TDev[i] ;
      if (TDevise.Code=Code) then
        BEGIN
        IE_DEVISE:=Code ;
        IE_QUOTITE:=TDevise.Quotite ;
        IE_TAUXDEV:=TDevise.TauxDev ;
        IE_DEBITDEV:=PIVOTTODEVISE(IE_DEBIT,IE_TAUXDEV,IE_QUOTITE,TDevise.Decimale) ;
        IE_CREDITDEV:=PIVOTTODEVISE(IE_CREDIT,IE_TAUXDEV,IE_QUOTITE,TDevise.Decimale) ;
        IE_DEBITEURO:=DeviseToEuro(IE_DEBITDEV,IE_TAUXDEV,IE_QUOTITE) ;
        IE_CREDITEURO:=DeviseToEuro(IE_CREDITDEV,IE_TAUXDEV,IE_QUOTITE) ;
        Break ;
        END ;
      END ;
    END ;
  END ;
END ;

function NatJal(OldJal : String3) : String3 ;
var QNatCpte: TQuery ;
    JalLu : TCptLu ;
BEGIN
Result:='' ;
QNatCpte:=OpenSQL('SELECT J_NATUREJAL FROM JOURNAL WHERE J_JOURNAL="'+Trim(OldJal)+'"',True) ;
if not QNatCpte.Eof then Result:=QNatCpte.Fields[0].AsString ;
Ferme(QNatCpte) ;
END ;

function NatJal2(OldJal : String3 ; Var TabJalSoc : TTabCptLu) : String3 ;
var QNatCpte: TQuery ;
    JalLu : TCptLu ;
BEGIN
Result:='' ;
If FALSE Then
   BEGIN
   QNatCpte:=OpenSQL('SELECT J_NATUREJAL FROM JOURNAL WHERE J_JOURNAL="'+Trim(OldJal)+'"',True) ;
   if not QNatCpte.Eof then Result:=QNatCpte.Fields[0].AsString ;
   Ferme(QNatCpte) ;
   END Else
   BEGIN
   JalLu:=TrouveJalSoc(Oldjal,TabJalSoc) ;
   Result:=JalLu.Nature ;
   END ;
END ;


procedure AjouteMvt(QAjoute : TQuery ; Var InfoImp : TInfoImport) ;

  function TrimSpe(St: String) : String ;
  BEGIN
  While (pos('"',St)<>0) do Delete(St,pos('"',St),1) ;
  Result:=Trim(St) ;
  END ;

  function IntToStrSpe(St : Integer) : string ;
  var t : string ;
  BEGIN
  t:=IntToStr(st) ;
  if (not isnumeric(t)) then result:='0' else Result:=IntToStr(St) ;
  END ;

BEGIN

If Imp_Methode1 And (Not QAJParam) Then QAjoute.SQL.Add(') VALUES (') ;

With MvtImport^ do
  BEGIN
  if TypeMvt<>'L' then Inc(InfoImp.NbLigIntegre) ;
  Inc(NbLigEnCours) ;
  if (TypeMvt<>'A') and (TypeMvt<>'L') then
  //Cumuls pour Compte-rendu
    if ((MsgBox=nil) or (ChoixFmt.CompteRendu) or (Lequel='FBA')) then
      BEGIN
      InfoImp.TotDeb:=InfoImp.TotDeb+IE_DEBIT ;
      InfoImp.TotCred:=InfoImp.TotCred+IE_CREDIT ;
      END ;
  if Positifs then
     BEGIN
     IE_DEBIT:=Abs(IE_DEBIT) ; IE_CREDIT:=Abs(IE_CREDIT) ;
     IE_DEBITDEV:=Abs(IE_DEBITDEV) ; IE_CREDITDEV:=Abs(IE_CREDITDEV) ;
     IE_DEBITEURO:=Abs(IE_DEBITEURO) ; IE_CREDITEURO:=Abs(IE_CREDITEURO) ;
     END ;
  If Imp_Methode1 And (Not QAJParam) Then
     BEGIN
     QAjoute.SQL.Add(IntToStrSpe(IE_CHRONO)+',"'+TrimSpe(IE_AFFAIRE)+'","'+
     TrimSpe(IE_ETATLETTRAGE)+'","'+TrimSpe(IE_LETTRAGE)+'","'+
     TrimSpe(IE_LETTREPOINTLCR)+'","'+TrimSpe(IE_LIBELLE)+'",');
     QAjoute.SQL.Add('"'+TrimSpe(IE_REFEXTERNE)+'","'+TrimSpe(IE_REFINTERNE)+'","'+
     TrimSpe(IE_REFLIBRE)+'","'+TrimSpe(IE_REFPOINTAGE)+'","'+TrimSpe(IE_REFRELEVE)+'",');
     QAjoute.SQL.Add('"'+TrimSpe(IE_RIB)+'","'+TrimSpe(IE_SECTION)+'","'+
     TrimSpe(IE_AXE)+'","'+ TrimSpe(IE_JOURNAL)+'","'+TrimSpe(IE_ECRANOUVEAU)+'","'+
     TrimSpe(IE_ETABLISSEMENT)+'",') ;

     QAjoute.SQL.Add('"'+TrimSpe(IE_FLAGECR)+'","'+TrimSpe(IE_MODEPAIE)+'","'+
     TrimSpe(IE_NATUREPIECE)+'","'+TrimSpe(IE_REGIMETVA)+'","'+TrimSpe(IE_QUALIFPIECE)+'",');
     QAjoute.SQL.Add('"'+TrimSpe(IE_QUALIFQTE1)+'","'+TrimSpe(IE_QUALIFQTE2)+'","'+
     TrimSpe(IE_SOCIETE)+'","'+TrimSpe(IE_TPF)+'","'+TrimSpe(IE_TVA)+'","'+
     TrimSpe(IE_TYPEANOUVEAU)+'",') ;
     QAjoute.SQL.Add('"'+TrimSpe(IE_TYPEECR)+'","'+TrimSpe(IE_TYPEMVT)+'","'+
     TrimSpe(IE_DEVISE)+'","'+TrimSpe(IE_AUXILIAIRE)+'","'+
     TrimSpe(IE_GENERAL)+'","'+TrimSpe(IE_CONTREPARTIEAUX)+'",') ;

     QAjoute.SQL.Add('"'+TrimSpe(IE_CONTREPARTIEGEN)+'",'+IntToStrSpe(IE_NUMECHE)+',"'+
     TrimSpe(IE_NUMPIECEINTERNE)+'",'+IntToStrSpe(IE_NUMLIGNE)+',');
     QAjoute.SQL.Add(IntToStrSpe(IE_NUMPIECE)+','+IntToStrSpe(IE_NUMVENTIL)+',"'+
     USDateTime(IE_DATECOMPTABLE)+'","'+USDateTime(IE_DATEECHEANCE)+'","'+
     USDateTime(IE_DATEPAQUETMAX)+'",');
     QAjoute.SQL.Add('"'+USDateTime(IE_DATEPAQUETMIN)+'","'+
     USDateTime(IE_DATEPOINTAGE)+'","'+
     USDateTime(IE_DATEREFEXTERNE)+'","'+USDateTime(IE_DATERELANCE)+'",') ;
     QAjoute.SQL.Add('"'+USDateTime(IE_DATETAUXDEV)+'","'+
     USDateTime(IE_DATEVALEUR)+'","'+
     USDateTime(IE_ORIGINEPAIEMENT)+'","'+IE_ECHE+'","'+IE_ENCAISSEMENT+'",');
     QAjoute.SQL.Add('"'+IE_CONTROLE+'","'+IE_LETTRAGEDEV+'","'+
     IE_OKCONTROLE+'","'+IE_SELECTED+'","'+IE_INTEGRE+'","'+IE_TVAENCAISSEMENT+'",');
     QAjoute.SQL.Add('"'+IE_TYPEANALYTIQUE+'","'+IE_VALIDE+'","'+IE_ANA+'",'+
     StrfPoint(IE_POURCENTAGE)+','+StrfPoint(IE_POURCENTQTE1)+',') ;

     QAjoute.SQL.Add(StrfPoint(IE_POURCENTQTE2)+','+StrfPoint(IE_QTE1)+','+StrfPoint(IE_QTE2)+','+
     StrfPoint(IE_QUOTITE)+','+StrfPoint(IE_RELIQUATTVAENC)+','+StrfPoint(IE_TAUXDEV)+',');
     QAjoute.SQL.Add(StrfPoint(IE_TOTALTVAENC)+','+
     StrfPoint(IE_DEBIT)+','+StrfPoint(IE_CREDIT)+','+StrfPoint(IE_CREDITDEV)+','+
     StrfPoint(IE_CREDITEURO)+',');
     QAjoute.SQL.Add(StrfPoint(IE_COUVERTURE)+','+StrfPoint(IE_COUVERTUREDEV)+','+
     StrfPoint(IE_DEBITDEV)+','+StrfPoint(IE_DEBITEURO)+')') ;
     ChangeSQL(QAjoute) ;
     QAjoute.ExecSQL ;
     QAjoute.Close ;
     END Else
     BEGIN
     If QAJParam Then
        BEGIN
        QAjoute.Close ;
        QAjoute.Params[0].AsInteger:=IE_CHRONO ;
        QAjoute.Params[1].AsString:=TrimSpe(IE_AFFAIRE) ;
        QAjoute.Params[2].AsString:=TrimSpe(IE_ETATLETTRAGE) ;
        QAjoute.Params[3].AsString:=TrimSpe(IE_LETTRAGE) ;
        QAjoute.Params[4].AsString:=TrimSpe(IE_LETTREPOINTLCR) ;
        QAjoute.Params[5].AsString:=TrimSpe(IE_LIBELLE) ;
        QAjoute.Params[6].AsString:=TrimSpe(IE_REFEXTERNE) ;
        QAjoute.Params[7].AsString:=TrimSpe(IE_REFINTERNE) ;
        QAjoute.Params[8].AsString:=TrimSpe(IE_REFLIBRE) ;
        QAjoute.Params[9].AsString:=TrimSpe(IE_REFPOINTAGE) ;
        QAjoute.Params[10].AsString:=TrimSpe(IE_REFRELEVE) ;
        QAjoute.Params[11].AsString:=TrimSpe(IE_RIB) ;
        QAjoute.Params[12].AsString:=TrimSpe(IE_SECTION) ;
        QAjoute.Params[13].AsString:=TrimSpe(IE_AXE) ;
        QAjoute.Params[14].AsString:=TrimSpe(IE_JOURNAL) ;
        QAjoute.Params[15].AsString:=TrimSpe(IE_ECRANOUVEAU) ;
        QAjoute.Params[16].AsString:=TrimSpe(IE_ETABLISSEMENT) ;
        QAjoute.Params[17].AsString:=TrimSpe(IE_FLAGECR) ;
        QAjoute.Params[18].AsString:=TrimSpe(IE_MODEPAIE) ;
        QAjoute.Params[19].AsString:=TrimSpe(IE_NATUREPIECE) ;
        QAjoute.Params[20].AsString:=TrimSpe(IE_REGIMETVA) ;
        QAjoute.Params[21].AsString:=TrimSpe(IE_QUALIFPIECE) ;
        QAjoute.Params[22].AsString:=TrimSpe(IE_QUALIFQTE1) ;
        QAjoute.Params[23].AsString:=TrimSpe(IE_QUALIFQTE2) ;
        QAjoute.Params[24].AsString:=TrimSpe(IE_SOCIETE) ; // GG ???
        QAjoute.Params[25].AsString:=TrimSpe(IE_TPF) ;
        QAjoute.Params[26].AsString:=TrimSpe(IE_TVA) ;
        QAjoute.Params[27].AsString:=TrimSpe(IE_TYPEANOUVEAU) ;
        QAjoute.Params[28].AsString:=TrimSpe(IE_TYPEECR) ;
        QAjoute.Params[29].AsString:=TrimSpe(IE_TYPEMVT) ;
        QAjoute.Params[30].AsString:=TrimSpe(IE_DEVISE) ;
        QAjoute.Params[31].AsString:=TrimSpe(IE_AUXILIAIRE) ;
        QAjoute.Params[32].AsString:=TrimSpe(IE_GENERAL) ;
        QAjoute.Params[33].AsString:=TrimSpe(IE_CONTREPARTIEAUX) ;
        QAjoute.Params[34].AsString:=TrimSpe(IE_CONTREPARTIEGEN) ;
        QAjoute.Params[35].AsInteger:=IE_NUMECHE ;
        QAjoute.Params[36].AsString:=TrimSpe(IE_NUMPIECEINTERNE) ;
        QAjoute.Params[37].AsInteger:=IE_NUMLIGNE ;
        QAjoute.Params[38].AsInteger:=IE_NUMPIECE ;
        QAjoute.Params[39].AsInteger:=IE_NUMVENTIL ;
        QAjoute.Params[40].AsDateTime:=IE_DATECOMPTABLE ;
        QAjoute.Params[41].AsDateTime:=IE_DATEECHEANCE ;
        QAjoute.Params[42].AsDateTime:=IE_DATEPAQUETMAX ;
        QAjoute.Params[43].AsDateTime:=IE_DATEPAQUETMIN ;
        QAjoute.Params[44].AsDateTime:=IE_DATEPOINTAGE ;
        QAjoute.Params[45].AsDateTime:=IE_DATEREFEXTERNE ;
        QAjoute.Params[46].AsDateTime:=IE_DATERELANCE ;
        QAjoute.Params[47].AsDateTime:=IE_DATETAUXDEV ;
        QAjoute.Params[48].AsDateTime:=IE_DATEVALEUR ;
        QAjoute.Params[49].AsDateTime:=IE_ORIGINEPAIEMENT ;
        QAjoute.Params[50].AsString:=IE_ECHE ;
        QAjoute.Params[51].AsString:=IE_ENCAISSEMENT ;
        QAjoute.Params[52].AsString:=IE_CONTROLE ;
        QAjoute.Params[53].AsString:=IE_LETTRAGEDEV ;
        QAjoute.Params[54].AsString:=IE_OKCONTROLE ;
        QAjoute.Params[55].AsString:=IE_SELECTED ;
        QAjoute.Params[56].AsString:=IE_INTEGRE ;
        QAjoute.Params[57].AsString:=IE_TVAENCAISSEMENT ;
        QAjoute.Params[58].AsString:=IE_TYPEANALYTIQUE ;
        QAjoute.Params[59].AsString:=IE_VALIDE ;
        QAjoute.Params[60].AsString:=IE_ANA ;
        QAjoute.Params[61].AsFloat:=IE_POURCENTAGE ;
        QAjoute.Params[62].AsFloat:=IE_POURCENTQTE1 ;
        QAjoute.Params[63].AsFloat:=IE_POURCENTQTE2 ;
        QAjoute.Params[64].AsFloat:=IE_QTE1 ;
        QAjoute.Params[65].AsFloat:=IE_QTE2 ;
        QAjoute.Params[66].AsFloat:=IE_QUOTITE ;
        QAjoute.Params[67].AsFloat:=IE_RELIQUATTVAENC ;
        QAjoute.Params[68].AsFloat:=IE_TAUXDEV ;
        QAjoute.Params[69].AsFloat:=IE_TOTALTVAENC ;
        QAjoute.Params[70].AsFloat:=IE_DEBIT ;
        QAjoute.Params[71].AsFloat:=IE_CREDIT ;
        QAjoute.Params[72].AsFloat:=IE_CREDITDEV ;
        QAjoute.Params[73].AsFloat:=IE_CREDITEURO ;
        QAjoute.Params[74].AsFloat:=IE_COUVERTURE ;
        QAjoute.Params[75].AsFloat:=IE_COUVERTUREDEV ;
        QAjoute.Params[76].AsFloat:=IE_DEBITDEV ;
        QAjoute.Params[77].AsFloat:=IE_DEBITEURO ;

        (*
        QAjoute.FindField('IE_LIBRETEXTE0').AsString:= ;
        QAjoute.FindField('IE_LIBRETEXTE1').AsString:= ;
        QAjoute.FindField('IE_LIBRETEXTE2').AsString:= ;
        QAjoute.FindField('IE_LIBRETEXTE3').AsString:= ;
        QAjoute.FindField('IE_LIBRETEXTE4').AsString:= ;
        QAjoute.FindField('IE_LIBRETEXTE5').AsString:= ;
        QAjoute.FindField('IE_LIBRETEXTE6').AsString:= ;
        QAjoute.FindField('IE_LIBRETEXTE7').AsString:= ;
        QAjoute.FindField('IE_LIBRETEXTE8').AsString:= ;
        QAjoute.FindField('IE_LIBRETEXTE9').AsString:= ;
        QAjoute.FindField('IE_TABLE0').AsString:= ;
        QAjoute.FindField('IE_TABLE1').AsString:= ;
        QAjoute.FindField('IE_TABLE2').AsString:= ;
        QAjoute.FindField('IE_TABLE3').AsString:= ;
        QAjoute.FindField('IE_LIBREBOOL0').AsString:= ;
        QAjoute.FindField('IE_LIBREBOOL1').AsString:= ;
        QAjoute.FindField('IE_LIBREMONTANT0').AsString:= ;
        QAjoute.FindField('IE_LIBREMONTANT1').AsString:= ;
        QAjoute.FindField('IE_LIBREMONTANT2').AsString:= ;
        QAjoute.FindField('IE_LIBREMONTANT3').AsString:= ;
        QAjoute.FindField('IE_LIBREDATE').AsString:= ;
        *)
        QAjoute.ExecSQL ;
        END Else
        BEGIN
        QAjoute.Insert ;
        InitNew(QAjoute) ;
        QAjoute.FindField('IE_CHRONO').AsInteger:=IE_CHRONO ;
        QAjoute.FindField('IE_TYPEECR').AsString:=TrimSpe(IE_TYPEECR) ;
        QAjoute.FindField('IE_GENERAL').AsString:=TrimSpe(IE_GENERAL) ;
        QAjoute.FindField('IE_AUXILIAIRE').AsString:=TrimSpe(IE_AUXILIAIRE) ;
        QAjoute.FindField('IE_AXE').AsString:=TrimSpe(IE_AXE) ;
        QAjoute.FindField('IE_SECTION').AsString:=TrimSpe(IE_SECTION) ;
        QAjoute.FindField('IE_JOURNAL').AsString:=TrimSpe(IE_JOURNAL) ;
        QAjoute.FindField('IE_DATECOMPTABLE').AsDateTime:=IE_DATECOMPTABLE ;
        QAjoute.FindField('IE_NUMPIECE').AsInteger:=IE_NUMPIECE ;
        QAjoute.FindField('IE_NUMLIGNE').AsInteger:=IE_NUMLIGNE ;
        QAjoute.FindField('IE_REFINTERNE').AsString:=TrimSpe(IE_REFINTERNE) ;
        QAjoute.FindField('IE_LIBELLE').AsString:=TrimSpe(IE_LIBELLE) ;
        QAjoute.FindField('IE_DEBIT').AsFloat:=IE_DEBIT ;
        QAjoute.FindField('IE_CREDIT').AsFloat:=IE_CREDIT ;
        QAjoute.FindField('IE_DEBITDEV').AsFloat:=IE_DEBITDEV ;
        QAjoute.FindField('IE_CREDITDEV').AsFloat:=IE_CREDITDEV ;
        QAjoute.FindField('IE_DEVISE').AsString:=TrimSpe(IE_DEVISE) ;
        QAjoute.FindField('IE_TAUXDEV').AsFloat:=IE_TAUXDEV ;
        QAjoute.FindField('IE_DATETAUXDEV').AsDateTime:=IE_DATETAUXDEV ;
        QAjoute.FindField('IE_POURCENTAGE').AsFloat:=IE_POURCENTAGE ;
        QAjoute.FindField('IE_OKCONTROLE').AsString:=IE_OKCONTROLE ;
        QAjoute.FindField('IE_SELECTED').AsString:=IE_SELECTED ;
        QAjoute.FindField('IE_NATUREPIECE').AsString:=TrimSpe(IE_NATUREPIECE) ;
        QAjoute.FindField('IE_QUALIFPIECE').AsString:=TrimSpe(IE_QUALIFPIECE) ;
        QAjoute.FindField('IE_REFEXTERNE').AsString:=TrimSpe(IE_REFEXTERNE) ;
        QAjoute.FindField('IE_REFLIBRE').AsString:=TrimSpe(IE_REFLIBRE) ;
        QAjoute.FindField('IE_AFFAIRE').AsString:=TrimSpe(IE_AFFAIRE) ;
        QAjoute.FindField('IE_COUVERTURE').AsFloat:=IE_COUVERTURE ;
        QAjoute.FindField('IE_LETTRAGE').AsString:=TrimSpe(IE_LETTRAGE) ;
        QAjoute.FindField('IE_LETTRAGEDEV').AsString:=IE_LETTRAGEDEV ;
        QAjoute.FindField('IE_REFPOINTAGE').AsString:=TrimSpe(IE_REFPOINTAGE) ;
        QAjoute.FindField('IE_DATEPOINTAGE').AsDateTime:=IE_DATEPOINTAGE ;
        QAjoute.FindField('IE_MODEPAIE').AsString:=TrimSpe(IE_MODEPAIE) ;
        QAjoute.FindField('IE_DATEECHEANCE').AsDateTime:=IE_DATEECHEANCE ;
        QAjoute.FindField('IE_QTE1').AsFloat:=IE_QTE1 ;
        QAjoute.FindField('IE_QTE2').AsFloat:=IE_QTE2 ;
        QAjoute.FindField('IE_QUALIFQTE1').AsString:=TrimSpe(IE_QUALIFQTE1) ;
        QAjoute.FindField('IE_QUALIFQTE2').AsString:=TrimSpe(IE_QUALIFQTE2) ;
        QAjoute.FindField('IE_ECRANOUVEAU').AsString:=TrimSpe(IE_ECRANOUVEAU) ;
        QAjoute.FindField('IE_DATEVALEUR').AsDateTime:=IE_DATEVALEUR ;
        QAjoute.FindField('IE_RIB').AsString:=TrimSpe(IE_RIB) ;
        QAjoute.FindField('IE_REFRELEVE').AsString:=TrimSpe(IE_REFRELEVE) ;
        QAjoute.FindField('IE_COUVERTUREDEV').AsFloat:=IE_COUVERTUREDEV ;
        QAjoute.FindField('IE_NUMECHE').AsInteger:=IE_NUMECHE ;
        QAjoute.FindField('IE_ANA').AsString:=IE_ANA ;
        QAjoute.FindField('IE_ORIGINEPAIEMENT').AsDateTime:=IE_ORIGINEPAIEMENT ;
        QAjoute.FindField('IE_FLAGECR').AsString:=TrimSpe(IE_FLAGECR) ;
        QAjoute.FindField('IE_NUMVENTIL').AsInteger:=IE_NUMVENTIL ;
        QAjoute.FindField('IE_POURCENTQTE1').AsFloat:=IE_POURCENTQTE1 ;
        QAjoute.FindField('IE_POURCENTQTE2').AsFloat:=IE_POURCENTQTE2 ;
        QAjoute.FindField('IE_VALIDE').AsString:=IE_VALIDE ;
        QAjoute.FindField('IE_DATEREFEXTERNE').AsDateTime:=IE_DATEREFEXTERNE ;
        QAjoute.FindField('IE_SOCIETE').AsString:=TrimSpe(IE_SOCIETE) ; // GG ???
        QAjoute.FindField('IE_ETABLISSEMENT').AsString:=TrimSpe(IE_ETABLISSEMENT) ;
        QAjoute.FindField('IE_DEBITEURO').AsFloat:=IE_DEBITEURO ;
        QAjoute.FindField('IE_CREDITEURO').AsFloat:=IE_CREDITEURO ;
        QAjoute.FindField('IE_QUOTITE').AsFloat:=IE_QUOTITE ;
        QAjoute.FindField('IE_TVAENCAISSEMENT').AsString:=IE_TVAENCAISSEMENT ;
        QAjoute.FindField('IE_REGIMETVA').AsString:=TrimSpe(IE_REGIMETVA) ;
        QAjoute.FindField('IE_TVA').AsString:=TrimSpe(IE_TVA) ;
        QAjoute.FindField('IE_TPF').AsString:=TrimSpe(IE_TPF) ;
        QAjoute.FindField('IE_CONTREPARTIEGEN').AsString:=TrimSpe(IE_CONTREPARTIEGEN) ;
        QAjoute.FindField('IE_CONTREPARTIEAUX').AsString:=TrimSpe(IE_CONTREPARTIEAUX) ;
        QAjoute.FindField('IE_DATEPAQUETMIN').AsDateTime:=IE_DATEPAQUETMIN ;
        QAjoute.FindField('IE_DATEPAQUETMAX').AsDateTime:=IE_DATEPAQUETMAX ;
        QAjoute.FindField('IE_LETTREPOINTLCR').AsString:=TrimSpe(IE_LETTREPOINTLCR) ;
        QAjoute.FindField('IE_DATERELANCE').AsDateTime:=IE_DATERELANCE ;
        QAjoute.FindField('IE_CONTROLE').AsString:=IE_CONTROLE ;
        QAjoute.FindField('IE_TOTALTVAENC').AsFloat:=IE_TOTALTVAENC ;
        QAjoute.FindField('IE_RELIQUATTVAENC').AsFloat:=IE_RELIQUATTVAENC ;
        QAjoute.FindField('IE_ETATLETTRAGE').AsString:=TrimSpe(IE_ETATLETTRAGE) ;
        QAjoute.FindField('IE_ENCAISSEMENT').AsString:=IE_ENCAISSEMENT ;
        QAjoute.FindField('IE_TYPEANOUVEAU').AsString:=TrimSpe(IE_TYPEANOUVEAU) ;
        QAjoute.FindField('IE_ECHE').AsString:=IE_ECHE ;
        QAjoute.FindField('IE_TYPEANALYTIQUE').AsString:=IE_TYPEANALYTIQUE ;
        QAjoute.FindField('IE_TYPEMVT').AsString:=TrimSpe(IE_TYPEMVT) ;
        QAjoute.FindField('IE_NUMPIECEINTERNE').AsString:=TrimSpe(IE_NUMPIECEINTERNE) ;
        QAjoute.FindField('IE_INTEGRE').AsString:=IE_INTEGRE ;
        (*
        QAjoute.FindField('IE_LIBRETEXTE0').AsString:= ;
        QAjoute.FindField('IE_LIBRETEXTE1').AsString:= ;
        QAjoute.FindField('IE_LIBRETEXTE2').AsString:= ;
        QAjoute.FindField('IE_LIBRETEXTE3').AsString:= ;
        QAjoute.FindField('IE_LIBRETEXTE4').AsString:= ;
        QAjoute.FindField('IE_LIBRETEXTE5').AsString:= ;
        QAjoute.FindField('IE_LIBRETEXTE6').AsString:= ;
        QAjoute.FindField('IE_LIBRETEXTE7').AsString:= ;
        QAjoute.FindField('IE_LIBRETEXTE8').AsString:= ;
        QAjoute.FindField('IE_LIBRETEXTE9').AsString:= ;
        QAjoute.FindField('IE_TABLE0').AsString:= ;
        QAjoute.FindField('IE_TABLE1').AsString:= ;
        QAjoute.FindField('IE_TABLE2').AsString:= ;
        QAjoute.FindField('IE_TABLE3').AsString:= ;
        QAjoute.FindField('IE_LIBREBOOL0').AsString:= ;
        QAjoute.FindField('IE_LIBREBOOL1').AsString:= ;
        QAjoute.FindField('IE_LIBREMONTANT0').AsString:= ;
        QAjoute.FindField('IE_LIBREMONTANT1').AsString:= ;
        QAjoute.FindField('IE_LIBREMONTANT2').AsString:= ;
        QAjoute.FindField('IE_LIBREMONTANT3').AsString:= ;
        QAjoute.FindField('IE_LIBREDATE').AsString:= ;
        *)
        QAjoute.Post ;
        END ;
     END ;
  END ;
END ;

procedure AjoutePourcentages ;
BEGIN
if (MvtImport^.IE_TYPEECR<>'A') then Exit ;
With MvtImport^,TotLigne do
  BEGIN
  if (IE_POURCENTAGE=0) and (TotEcr<>0) then IE_POURCENTAGE:=Arrondi(((IE_DEBIT-IE_CREDIT)/TotEcr)*100,ADecimP) ;
  if IE_POURCENTQTE1=0 then IE_POURCENTQTE1:=IE_POURCENTAGE ;
  IE_POURCENTQTE2:=IE_POURCENTQTE1 ;
  END ;
END ;

procedure MemoriseMontantGene ;
BEGIN
if (MvtImport^.IE_ANA='-') or (MvtImport^.IE_TYPEECR='A') then Exit ;
With MvtImport^,TotLigne do
  BEGIN
  TotEcr:=IE_DEBIT-IE_CREDIT ;
{
  TotDev:=IE_DEBITDEV-IE_CREDITDEV ;
  TotEuro:=IE_DEBITEURO-IE_CREDITEURO ;
  TotQte1:=IE_QTE1 ;
  TotQte1:=IE_QTE2 ;
  QualQte1:=IE_QUALIFQTE1 ;
  QualQte2:=IE_QUALIFQTE2 ;
}
  END ;
END ;

procedure AjouteAnalytique(Axe : String5 ; Var TabJalSoc : TTabCptLu) ;
var i: integer ;
    QuelAxe : String3 ;
BEGIN
if (Axe='-----') then Exit ;
QuelAxe:='A1' ;
if Axe[2]='X' then QuelAxe:='A2' else
 if Axe[3]='X' then QuelAxe:='A3' else
  if Axe[4]='X' then QuelAxe:='A4' else
   if Axe[5]='X' then QuelAxe:='A5' else
//Inc(NbLigEnCours) ;
Inc(NumVentil) ;
With MvtImport^ do
  BEGIN
  IE_CHRONO:=NbLigEnCours;IE_JOURNAL:=OldJal;IE_DATECOMPTABLE:=OldDateC;
  IE_ANA:='X' ;
  if (NatJal2(OldJal,TabJalSoc)='ANO') then IE_ECRANOUVEAU:='OAN' ;
  IE_TYPEECR:='A';IE_NUMVENTIL:=NumVentil;
  IE_AXE:=QuelAxe;
  i:=Round(Valeur(Copy(QuelAxe,2,1))) ;
  IE_SECTION:=SectAttente[i];
  AjoutePourcentages ;
  IE_QTE1:=0;IE_QTE2:=0;IE_QUALIFQTE1:='...';IE_QUALIFQTE2:='...' ;
  END;
END ;

// Lignes écritures EDI

procedure ImportLigneEDI(St : String ; QAjoute : TQuery ; Var TabJalSoc : TTabCptLu ; Var InfoImp : TInfoImport) ;
var FAnal : TAnal ;
    SQL : String ;
    Axe : String5 ;
    Lettrable : boolean ;
BEGIN
SQL:='' ; Axe:='-----' ;
// Detail uniquement...( pour l'instant ?).
if (Round(Valeur(Copy(St,1,5)))<250) or (Round(Valeur(Copy(St,1,5)))>400) then Exit ;
Case Round(Valeur(Copy(St,1,5))) of
  250  : BEGIN
         Jal:=Copy(St,6,3) ;
         END ;
  285  : BEGIN
         DateC:=Format_DateEDI(Copy(St,97,8)) ;
         NumPiece:=Round(Valeur(Copy(St,6,12))) ;
         if NumPiece=0 then NumPiece:=1 ;
         if Copy(St,85,3)='VAL' then BEGIN QualP:='N' ; END Else //IE_VALIDE:='X' ; END else
          if Copy(St,85,3)='SIM' then QualP:='S' else
           if Copy(St,85,3)='NVL' then QualP:='N' ;
         OkRupt:=TestBreak(InfoImp) ;
         if OkRupt then Exit ;
         MvtImport^:=InitMvtImp^ ;
         With MvtImport^ do
           BEGIN
           IE_CHRONO:=NbLigEnCours ;
           IE_JOURNAL:=Jal ;
           if (NatJal2(Jal,TabJalSoc)='ANO') then IE_ECRANOUVEAU:='OAN' ;
           IE_DATECOMPTABLE:=DateC ;
           IE_NUMPIECE:=NumPiece ;
           IE_NUMLIGNE:=Round(Valeur(Copy(St,18,6))) ;
           if (QualP='N') and (QualifPiece<>'') and (IE_ECRANOUVEAU<>'OAN') then IE_QUALIFPIECE:=QualifPiece
                                                                            else IE_QUALIFPIECE:=QualP ;
           IE_GENERAL:=Copy(St,33,17) ;
           IE_AUXILIAIRE:=Copy(St,59,17) ;
           RecupInfosCpte(IE_GENERAL,IE_AUXILIAIRE) ;
           Lettrable:=False ;
           if IE_AUXILIAIRE='' then NatureCompte(IE_GENERAL,True,Lettrable)
                               else NatureCompte(IE_AUXILIAIRE,False,Lettrable) ;
           if Pointable='X' then BEGIN IE_ECHE:='X' ; END ;
           VideLettrage(Lettrable) ;
           if IsVentilable(IE_GENERAL,Axe,'') then IE_ANA:='X' ;
           IE_TYPEECR:='E' ;
           IE_DATEECHEANCE:=Format_DateEDI(Copy(St,141,8)) ;
           IE_MODEPAIE:=TrouveModePaie('',Valeur(StPoint(Copy(St,163,18))),TRUE) ;
           if Copy(St,187,2)='DP' then IE_DEBIT:=Valeur(StPoint(Copy(St,163,18))) else
           if Copy(St,187,2)='CP' then IE_CREDIT:=StrToFloat(StPoint(Copy(St,163,18))) else
           if Positifs then
             BEGIN
             if Copy(St,187,2)='DN' then IE_CREDIT:=Valeur(StPoint(Copy(St,163,18))) else
             if Copy(St,187,2)='CN' then IE_DEBIT:=Valeur(StPoint(Copy(St,163,18))) ;
             END else
             BEGIN
             if Copy(St,187,2)='DN' then IE_DEBIT:=(Valeur(StPoint(Copy(St,163,18)))*-1) else
             if Copy(St,187,2)='CN' then IE_CREDIT:=(Valeur(StPoint(Copy(St,163,18)))*-1) ;
             END ;
           AjouteDevise(V_PGI.DevisePivot) ;
           if lettrable then
             BEGIN
             Inc(NumEche) ; IE_ECHE:='X' ; IE_NUMECHE:=NumEche ;
             END ;
           MemoriseMontantGene ;
           {
           if (IE_AUXILIAIRE<>OldAuxiliaire) or (IE_AUXILIAIRE='') then InitRuptEche ;
           if ((IE_DATEECHEANCE<>OldDateEche) or (IE_MODEPAIE<>OldModePaie)) and Lettrable then
             BEGIN
             OldDateEche:=IE_DATEECHEANCE ; OldModePaie:=IE_MODEPAIE ;
             Inc(NumEche) ; IE_ECHE:='X' ; IE_NUMECHE:=NumEche ;
             END ;
           }
           //if (Copy(St,65,3)='IMP') then IE_ETAT:='IMP' else Etat:='MAN' ;
           //if (Copy(St,77,3)='TSE') then IE_TVAENCAISSEMENT='X' ;
           END ;
         END ;
  287  : BEGIN
         MvtImport^.IE_QTE1:=Valeur(StPoint(Copy(St,30,16))) ;
         MvtImport^.IE_QUALIFQTE1:=Copy(St,46,3) ;
         END ;
  370  : BEGIN
         With MvtImport^ do
           BEGIN
           IE_REFINTERNE:=Copy(St,33,17) ;
           IE_NATUREPIECE:=Copy(St,56,3) ;
           IE_REGIMETVA:=RegimeTva ;
           if IE_AUXILIAIRE='' then IE_TYPEMVT:=QuelTypeMvt(IE_GENERAL, NatureCompte(IE_GENERAL,True,Lettrable),IE_NATUREPIECE)
                               else IE_TYPEMVT:=QuelTypeMvt(IE_AUXILIAIRE, NatureCompte(IE_AUXILIAIRE,False,Lettrable),IE_NATUREPIECE) ;
           IE_LIBELLE:=Copy(St,65,70) ;
           IE_ETABLISSEMENT:=VH^.EtablisDefaut ;
           IE_SOCIETE:=V_PGI.CodeSociete ;
           if IsVentilable(MvtImport^.IE_GENERAL,Axe,'') then IE_ANA:='X' ;
           END ;
         AjouteMvt(QAjoute,InfoImp) ;
         if (Axe<>'-----') then
           BEGIN
           With FAnal,MvtImport^ do
             BEGIN
             Journal:=Jal ;
             NatPiece:=IE_NATUREPIECE ;
             QualP:=IE_QUALIFPIECE ;
             General:=IE_GENERAL ;
             NumP:=IE_NUMPIECE ;
             NumL:=IE_NUMLIGNE ;
             NumV:=1 ;
             RefInt:=IE_REFINTERNE ;
             Lib:=IE_LIBELLE ;
             Debit:=IE_DEBIT ;
             Credit:=IE_CREDIT ;
             END ;
           MvtImport^:=InitMvtImp^ ;
           MvtImport^.IE_ETABLISSEMENT:=VH^.EtablisDefaut ;
           MvtImport^.IE_SOCIETE:=V_PGI.CodeSociete ;
           With MvtImport^,FAnal do
             BEGIN
             IE_JOURNAL:=Journal ;
             IE_NATUREPIECE:=NatPiece ;
             IE_QUALIFPIECE:=QualP ;
             IE_GENERAL:=General ;
             IE_NUMPIECE:=NumP ;
             IE_NUMLIGNE:=NumL ;
             IE_NUMVENTIL:=NumV ;
             IE_REFINTERNE:=RefInt ;
             IE_LIBELLE:=Lib ;
             IE_DEBIT:=Debit ;
             IE_CREDIT:=Credit ;
             AjouteDevise(V_PGI.DevisePivot) ;
             END ;
           AjouteAnalytique(Axe,TabJalSoc) ; PrepareQuery(QAjoute) ; AjouteMvt(QAjoute,InfoImp) ;
           END ;
         END ;
  END ;
END;

{================= Balance EDI =====================}

procedure ImportBalance(St : string ; NumPiece : integer ; QAjoute : TQuery ; Var TabJalSoc : TTabCptLu ;
                        Var InfoImp : TInfoImport) ;
var NumL : integer ;
    SANouvD,SANouvC,MvtD,MvtC,Mt : Double ;
    CompteGene,Auxi : String17 ;
    Ok,OkColl,Lettrable : boolean ;
    Q : TQuery ;
    Axe : String5 ;
    FAnal : TAnal ;
BEGIN
OkRupt:=TestBreak(InfoImp) ;
if OkRupt then Exit ;
OkColl:=False ; Lettrable:=False ;
if (Round(Valeur(Copy(St,1,5)))=1030) then
  BEGIN
  DateDeb:=Format_DateEDI(Copy(st,66,8)) ;
  DateFin:=Format_DateEDI(Copy(st,74,8)) ;
  Exit ;
  END ;
if (Round(Valeur(Copy(St,1,5)))<>1230) then Exit ;
CompteGene:=Copy(St,6,17) ;
// A nouveaux
SANouvD:=Round(Valeur(Trim(Copy(St,133,20)))) ;
SANouvC:=Round(Valeur(Trim(Copy(St,161,20)))) ;
if (SANouvD=0) and (SANouvC=0) and (NumPiece=1) then Exit ;
// Solde
MvtD:=Round(Valeur(Trim(Copy(St,77,20)))) ;
MvtC:=Round(Valeur(Trim(Copy(St,105,20)))) ;
// Mouvements                   *
Mt:=(MvtD-MvtC)-(SANouvD-SANouvC) ;
if ((Mt=0) and (NumPiece=2)) then Exit ;
if Mt>0 then BEGIN MvtD:=Mt ; MvtC:=0 ; END else BEGIN MvtD:=0 ; MvtC:=Abs(Mt) ; END ;

if (EstCollectif(CompteGene)) then
  BEGIN
  OkColl:=True ;
{$IFDEF SPEC302}
  Q:=OpenSQL('Select SO_CLIATTEND,SO_FOUATTEND FROM SOCIETE',True) ;
  if Not Q.EOF then
    if Copy(CompteGene,1,2)='41' then Auxi:=Q.Fields[0].AsString else Auxi:=Q.Fields[1].AsString ;
  Ferme(Q) ;
{$ELSE}
  if Copy(CompteGene,1,2)='41' then Auxi:=GetParamSoc('SO_CLIATTEND') else Auxi:=GetParamSoc('SO_FOUATTEND') ;
{$ENDIF}
  END ;
InfoImp.NbPiece:=NumPiece ;
Ok:=True ;
While Ok do
  BEGIN
  Inc(NumLignePiece) ;
  NumL:=NbLigEnCours ;
  MvtImport^:=InitMvtImp^ ;
  With MvtImport^ do
    BEGIN
    IE_NUMPIECE:=NumPiece ;
    IE_JOURNAL:='' ;
    IE_CHRONO:=NumL ;
    IE_NATUREPIECE:='OD' ;
    IE_QUALIFPIECE:='N' ;
    IE_NUMLIGNE:=NumLignePiece ;
    IE_GENERAL:=CompteGene ;
    IE_TYPEECR:='E' ;
    IE_REFINTERNE:='' ;
    if OkColl then IE_AUXILIAIRE:=Auxi ;
    RecupInfosCpte(IE_GENERAL,IE_AUXILIAIRE) ;
    if IE_AUXILIAIRE='' then NatureCompte(IE_GENERAL,True,Lettrable)
                        else NatureCompte(IE_AUXILIAIRE,False,Lettrable) ;
    if Pointable='X' then BEGIN IE_ECHE:='X' ; END ;
    VideLettrage(Lettrable) ;
    if Lettrable then
      BEGIN
      IE_ECHE:='X' ;
      IE_NUMECHE:=1 ;
      END ;
    IE_MODEPAIE:=TrouveModePaie('',Mt,True) ;
    Case NumPiece of
      1 : BEGIN // A nouveaux..
          OldDateC:=DateDeb ; ;
          IE_ECRANOUVEAU:='OAN' ;
          if (SANouvD<>0) then BEGIN IE_DEBIT:=SANouvD ; SANouvD:=0 ; END
                          else BEGIN IE_CREDIT:=SANouvC ; SANouvC:=0 ; END ;
          END ;
      2:  BEGIN
          OldDateC:=DateFin ;
          if (MvtD<>0) then BEGIN IE_DEBIT:=MvtD ; MvtD:=0 ; END
                       else BEGIN IE_CREDIT:=MvtC ; MvtC:=0 ; END ;
          END ;
      END ;
    IE_DATECOMPTABLE:=OldDateC ;
    if IsVentilable(IE_GENERAL,Axe,'') then BEGIN IE_ANA:='X' ; END ;
    AjouteDevise(V_PGI.DevisePivot) ;
    END ;
  TypeMvt:=' ' ; MemoriseMontantGene ;
  AjouteMvt(QAjoute,InfoImp) ;
  if (Axe<>'-----') then
    BEGIN
    With FAnal,MvtImport^ do
      BEGIN
      Journal:=Jal ;
      NatPiece:=IE_NATUREPIECE ;
      QualP:=IE_QUALIFPIECE ;
      General:=IE_GENERAL ;
      NumP:=IE_NUMPIECE ;
      NumL:=IE_NUMLIGNE ;
      NumV:=1 ;
      RefInt:=IE_REFINTERNE ;
      Lib:=IE_LIBELLE ;
      Debit:=IE_DEBIT ;
      Credit:=IE_CREDIT ;
      END ;
    MvtImport^:=InitMvtImp^ ;
    MvtImport^.IE_ETABLISSEMENT:=VH^.EtablisDefaut ;
    MvtImport^.IE_SOCIETE:=V_PGI.CodeSociete ;
    With MvtImport^,FAnal do
      BEGIN
      Case NumPiece of
        1 : BEGIN // A nouveaux..
            IE_DATECOMPTABLE:=DateDeb ;
            IE_ECRANOUVEAU:='OAN' ;
            END ;
        2:  BEGIN
            IE_DATECOMPTABLE:=DateFin ;
            END ;
        END ;
      IE_JOURNAL:=Journal ;
      IE_NATUREPIECE:=NatPiece ;
      IE_QUALIFPIECE:=QualP ;
      IE_GENERAL:=General ;
      IE_NUMPIECE:=NumP ;
      IE_NUMLIGNE:=NumL ;
      IE_NUMVENTIL:=1 ;
      IE_REFINTERNE:=RefInt ;
      IE_LIBELLE:=Lib ;
      IE_DEBIT:=Debit ;
      IE_CREDIT:=Credit ;
      AjouteDevise(V_PGI.DevisePivot) ;
      END ;
    AjouteAnalytique(Axe,TabJalSoc) ; PrepareQuery(QAjoute) ; TypeMvt:='A' ; AjouteMvt(QAjoute,InfoImp) ;
    END ;
  if (Not OkColl) or (OkColl and (NumPiece=1) and (SANouvD=0) and (SANouvC=0))
     or (OkColl and (NumPiece=2) and (MvtD=0) and (MvtC=0)) then Ok:=False
  END ;
END ;

procedure MAJContreparties ;
var NatJ : String3 ;
    St : String ;
    CPGene,CPAuxi : String17 ;
BEGIN
// Plus tard
Exit ;
NatJ:=NatJal(MvtImport^.IE_JOURNAL) ;
if (NatJ<>'ACH') and (NatJ<>'VTE') then Exit ;
With MvtImport^ do
  BEGIN
  if (IE_AUXILIAIRE<>'') then
    BEGIN
    Read(Fichier,St) ;
    IE_CONTREPARTIEGEN:=Copy(St,12,13) ;
    IE_CONTREPARTIEAUX:=Copy(St,26,13) ;
    CPGene:=IE_GENERAL ; CPGene:=IE_AUXILIAIRE ;
    END else
    BEGIN
    IE_CONTREPARTIEGEN:=CPGene ;
    IE_CONTREPARTIEAUX:=CPAuxi ;
    END ;
  END ;
END ;

{================= Rapprochement SAARI =====================}

{
procedure AjouteReference(Gene : String17 ; StRef : String ; DD : TDateTime) ;
var TRef : TReference ;
    i : integer ;
BEGIN
if (ListePointe=nil) then ListePointe:=TStringList.Create ;
TRef:=TReference.Create ;
TRef.Ref:=StRef ; TRef.Date:=DD ;
i:=ListePointe.IndexOf(Gene) ;
if i>-1 then ListePointe.Objects[i]:=TRef else ListePointe.AddObject(Gene,TRef) ;
END ;
}

procedure RapprocheLigne(Var InfoImp : TInfoImport) ;
var ExistMvt : TExistMvt ;
    Ref : String ;
BEGIN
if AnnuleImport then BEGIN OkRupt:=True ; Exit ; END ;
InitExistMvt(ExistMvt) ;
With ExistMvt do
  BEGIN
  JOURNAL:=QImpEcr.FindField('IE_JOURNAL').AsString ;
  DATECOMPTABLE:=QImpEcr.FindField('IE_DATECOMPTABLE').AsDateTime ;
  GENERAL:=QImpEcr.FindField('IE_GENERAL').AsString ;
  if QImpEcr.FindField('IE_AUXILIAIRE').AsString<>'' then AUXIANA:=QImpEcr.FindField('IE_AUXILIAIRE').AsString ;
  if QImpEcr.FindField('IE_SECTION').AsString<>'' then AUXIANA:=QImpEcr.FindField('IE_SECTION').AsString ;
  REFINTERNE:=QImpEcr.FindField('IE_REFINTERNE').AsString ;
  LIBELLE:=QImpEcr.FindField('IE_LIBELLE').AsString ;
  DEBIT:=QImpEcr.FindField('IE_DEBIT').AsFloat ;
  CREDIT:=QImpEcr.FindField('IE_CREDIT').AsFloat ;
  InfoImp.TotDeb:=InfoImp.TotDeb+DEBIT ;
  InfoImp.TotCred:=InfoImp.TotCred+CREDIT ;
  END ;
Ref:=Trim(QImpEcr.FindField('IE_REFPOINTAGE').AsString) ;
if ExisteMouvement(ExistMvt) and (Ref<>'') then
   BEGIN
   //i:=ListePointe.IndexOf(ExistMvt.GENERAL) ;
   ExecuteSQL('UPDATE Ecriture Set E_REFPOINTAGE="'+Ref+'"'+
              ', E_DATEPOINTAGE="'+USDateTime(QImpEcr.FindField('IE_DATEPOINTAGE').AsDateTime)+'" WHERE '+RecupWhereExistMvt(ExistMvt))
   END ;
END ;

function BourreOuTronque(St : String ; fb : TFichierBase) : String ;
var Diff : Integer ;
BEGIN
St:=Trim(St) ; Result:=St ;
if (not (LgComptes)) or (St='') then Exit ;
Diff:=Length(St)-LongMax(fb) ;
if Diff>0 then St:=BourreLess(St,fb) ;
Result:=BourreLaDonc(St,fb) ;
END ;

procedure ImporteZones(TypeMvt : String ; St : String ; LongFormat : boolean ; Lettrable : Boolean ; Var TabJalSoc : TTabCptLu) ;
Var Axe : String5 ;
    OnRemplace : Boolean ;
    TTiers : TGTiers ;
    Decal,DecalGene,LgCptes : byte ;
BEGIN
Decal:=0 ; DecalGene:=0 ; LgCptes:=13 ;
// Comptes sur 17 et non 13 caractères
if not VH^.ImportRL and ((Lequel='FEC') or (Lequel='FBA')) and ((Format='HAL') or (Format='HLI')) then
   BEGIN
   DecalGene:=4 ; Decal:=8 ; LgCptes:=17 ;
   END ;
With MvtImport^ do
  BEGIN
  {--------------------- Echéances -----------------------}
  (* GP le 10/07/97 Pour recup pont paye : Date eche et mode paiement non alimenté
  if (((FFormat.Value='SN2') and (TypeMvt='E'))
    or (FFormat.Value='SAA') or (FFormat.Value='HAL'))
    and (Copy(St,78,6)<>'      ') and Lettrable then
  *)
  if Lettrable and (((Format='SN2') and (TypeMvt='E')) or (Format='SAA') or (Format='HAL') or (Format='HLI')) then
      BEGIN
      if (((Format='HAL') or (Format='HLI')) and (Copy(St,80+Decal,8)<>'        ')) then IE_DATEECHEANCE:=Format_Date_HAL(Copy(St,80+Decal,8)) else
        if ((Format<>'HAL') and (Format<>'HLI') and (Copy(St,78,6)<>'      '))  then IE_DATEECHEANCE:=Format_Date(Copy(St,78,6))
                                                                                else IE_DATEECHEANCE:=IE_DATECOMPTABLE ;

      if (Format='HAL') or (Format='HLI') then OnRemplace:=(Copy(St,79+Decal,1)<>' ') And (Copy(St,79+Decal,1)<>'S')
                                          else OnRemplace:=(Copy(St,77,1)<>' ') And (Copy(St,77,1)<>'S') ;

      if ((Format<>'HAL') or (Lequel='FBA')) then
         if (Lequel='FBA') or (Format='HLI') then IE_MODEPAIE:=TrouveModePaie(CHERCHEUNMODE(Copy(St,79+Decal,1),False,False),Valeur(StPoint(Copy(St,89+Decal,20))),OnRemplace)
                                             else IE_MODEPAIE:=TrouveModePaie(CHERCHEUNMODE(Copy(St,77,1),False,False),Valeur(StPoint(Copy(St,85,20))),OnRemplace) ;
      //if (IE_AUXILIAIRE<>OldAuxiliaire) or (IE_AUXILIAIRE='') then InitRuptEche ;
      //if ((IE_DATEECHEANCE<>OldDateEche) or (IE_MODEPAIE<>OldModePaie)) then
      //  BEGIN
        OldModePaie:=IE_MODEPAIE ; OldDateEche:=IE_DATEECHEANCE ;
        Inc(NumEche) ;
      //  END ;
      IE_ECHE:='X' ; IE_NUMECHE:=NumEche ;
      IE_DATEPAQUETMIN:=IE_DATECOMPTABLE ; IE_DATEPAQUETMAX:=IE_DATECOMPTABLE ;
      IE_ETATLETTRAGE:='AL' ;
      END ;
  if not LongFormat then
    BEGIN
    if TypeMvt<>'A' then
      BEGIN
      if (IE_NUMECHE<=1) then Inc(NumLignePiece) ;
      END else
      BEGIN
      if (IE_AUXILIAIRE<>'') and (IE_MODEPAIE='') then
        BEGIN
        TTiers:=TGTiers.Create(IE_AUXILIAIRE) ;
        IE_MODEPAIE:=FirstModePaie(TTiers.ModeRegle) ;
        TTiers.Free ;
        END ;
      // Balance Halley
      if Format='HAL' then IsVentilable(IE_GENERAL,Axe,Copy(St,13,1))
                      else IsVentilable(IE_GENERAL,Axe,Copy(St,11,1)) ;
      if (Axe='-----') then Axe:='X----' ;
      AjouteAnalytique(Axe,TabJalSoc) ;
      if ((Format='HAL') or (Format='HLI')) then
        BEGIN
        if IE_SECTION<>'' then
          BEGIN
          IE_SECTION:=Copy(St,28+4,LgCptes) ;
          if (Format='HAL') then IE_AXE:=Copy(St,117+Decal,2) ;
          END ;
        END else if IE_SECTION<>'' then
           BEGIN
           IE_SECTION:=Copy(St,26,13) ; //IE_AXE:=Copy(St,113,2) ;
           // GP le 12/12/97
           //If Trim(Copy(St,113,2))<>'' Then IE_AXE:=Copy(St,113,2) ;
           END ;
      END ;
    IE_NUMLIGNE:=NumLignePiece ;
    Exit ;
    END;
// Format Halley
  if Trim(Copy(St,119+Decal,1))<>'' then IE_VALIDE:=Copy(St,119+Decal,1) ;
  if Trim(Copy(St,120+Decal,35))<>'' then IE_REFEXTERNE:=Copy(St,120+Decal,35) ;
  IE_DATEREFEXTERNE:=Format_Date_HAL(Copy(St,155+Decal,8)) ;
  //IE_DATECREATION:=Format_Date_HAL(Copy(St,163+Decal,8)) ;
  //IE_DATEMODIFICATION:=Format_Date_HAL(Copy(St,171+Decal,8)) ;
  IE_SOCIETE:=V_PGI.CodeSociete ; //Copy(St,179+Decal,3) ;
  if Trim(Copy(St,182+Decal,3))<>'' then IE_ETABLISSEMENT:=Copy(St,182+Decal,3) ;
  if Trim(Copy(St,185+Decal,17))<>'' then IE_AFFAIRE:=Copy(St,185+Decal,17);
  IE_DEBITEURO:=Valeur(StPoint(Copy(St,202+Decal,20)));

  IE_CREDITEURO:=Valeur(StPoint(Copy(St,222+Decal,20)));
  //IE_TAUXEURO:=Valeur(StPoint(Copy(St,242+Decal,20)));
  if Trim(Copy(St,262+Decal,3))<>'' then IE_DEVISE:=Copy(St,262+Decal,3);
  IE_DEBITDEV:=Valeur(StPoint(Copy(St,265+Decal,20))); IE_CREDITDEV:=Valeur(StPoint(Copy(St,285+Decal,20)));
  if Trim(Copy(St,305+Decal,20))<>'' then IE_TAUXDEV:=Valeur(StPoint(Copy(St,305+Decal,20)));
  IE_DATETAUXDEV:=Format_Date_HAL(Copy(St,325+Decal,8));
  if Trim(Copy(St,333+Decal,20))<>'' then IE_QUOTITE:=Valeur(StPoint(Copy(St,333+Decal,20)));
  if Trim(Copy(St,353+Decal,3))<>'' then IE_ECRANOUVEAU:=Copy(St,353+Decal,3);
  IE_QTE1:=Valeur(StPoint(Copy(St,356+Decal,20)));
  IE_QTE2:=Valeur(StPoint(Copy(St,376+Decal,20)));
  IE_QUALIFQTE1:=Copy(St,396+Decal,3);
  IE_QUALIFQTE2:=Copy(St,399+Decal,3);
  if TypeMvt<>'A' then
    BEGIN
    IE_REFLIBRE:=Copy(St,402+Decal,35);
    if Trim(Copy(St,437+Decal,1))<>'' then IE_TVAENCAISSEMENT:=Copy(St,437+Decal,1);
    if Trim(Copy(St,438+Decal,3))<>'' then IE_REGIMETVA:=Copy(St,438+Decal,3);
    if Trim(Copy(St,441+Decal,3))<>'' then IE_TVA:=Copy(St,441+Decal,3);
    if Trim(Copy(St,444+Decal,3))<>'' then IE_TPF:=Copy(St,444+Decal,3);
    IE_CONTREPARTIEGEN:=Copy(St,447+Decal,17);
    IE_CONTREPARTIEAUX:=Copy(St,464+Decal,17);

    if (TypeMvt='L') then
      BEGIN
      IE_COUVERTURE:=Valeur(StPoint(Copy(St,481+Decal,20)));
      IE_LETTRAGE:=Copy(St,501+Decal,5);
      IE_LETTRAGEDEV:=Copy(St,506+Decal,1);
      IE_DATEPAQUETMAX:=Format_Date_HAL(Copy(St,507+Decal,8));
      IE_DATEPAQUETMIN:=Format_Date_HAL(Copy(St,515+Decal,8));

      IE_COUVERTUREDEV:=Valeur(StPoint(Copy(St,679+Decal,20)));
      if Trim(Copy(St,447+Decal,17))<>'' then IE_ETATLETTRAGE:=Copy(St,699+Decal,3);
      END ;
    IE_REFPOINTAGE:=Copy(St,523+Decal,17);
    IE_DATEPOINTAGE:=Format_Date_HAL(Copy(St,540+Decal,8));
    //IE_LETTREPOINTLCR:=Copy(St,548+Decal,4);
    IE_DATERELANCE:=Format_Date_HAL(Copy(St,552+Decal,8));
    if Trim(Copy(St,560,1))<>'' then IE_CONTROLE:=Copy(St,560,1);
    //IE_TOTALTVAENC:=Valeur(StPoint(Copy(St,561+Decal,20)));
    //IE_RELIQUATTVAENC:=Valeur(StPoint(Copy(St,581+Decal,20)));
    IE_DATEVALEUR:=Format_Date_HAL(Copy(St,601+Decal,8));
    IE_RIB:=Copy(St,609+Decal,35);
    IE_REFRELEVE:=Copy(St,644+Decal,35);

    IE_NUMPIECEINTERNE:=Copy(St,702+Decal,35);
    IE_ENCAISSEMENT:=Copy(St,737+Decal,3);
    IE_TYPEANOUVEAU:=Copy(St,740+Decal,3);
    if Trim(Copy(St,743+Decal,1))<>'' then IE_ECHE:=Copy(St,743+Decal,1);
    if StrToIntDef(Copy(St,117+Decal,2),0)<>0 then IE_NUMECHE:=StrToIntDef(Copy(St,117+Decal,2),0) ;
    if Trim(Copy(St,744+Decal,1))<>'' then IE_ANA:=Copy(St,744+Decal,1) ;
    if Trim(Copy(St,745+Decal,3))<>'' then IE_MODEPAIE:=Copy(St,745+Decal,3) ;
    END else
    BEGIN
    if (Format='HAL') or (Format='HLI') then BEGIN IE_SECTION:=Copy(St,28+DecalGene,LgCptes) ; IE_AXE:=Copy(St,117+Decal,2) ; END
                                        else BEGIN IE_SECTION:=Copy(St,26,13) ; {IE_AXE:=Copy(St,113,2) ;} END ;
    IE_POURCENTAGE:=Valeur(StPoint(Copy(St,402+Decal,20)));
    //'IE_TOTALECRITURE='+StrfPoint(Valeur(Copy(St,422+Decal,20)) ;
    IE_CONTROLE:=Copy(St,422+Decal,1);
    IE_NUMVENTIL:=StrToIntDef(Copy(St,423+Decal,6),0);IE_POURCENTQTE1:=Valeur(StPoint(Copy(St,429+Decal,20)));IE_POURCENTQTE2:=Valeur(StPoint(Copy(St,449+Decal,20)));
    //IE_TOTALDEVISE='+StrfPoint(Valeur(Copy(St,477+Decal,20)) ; ',IE_TOTALEURO='+StrfPoint(Valeur(Copy(St,497+Decal,20)) ;
    //IE_TOTALQTE1='+StrfPoint(Valeur(Copy(St,517+Decal,20)) ; ',IE_TOTALQTE2='+StrfPoint(Valeur(Copy(St,537+Decal,20)) ;
    IE_TYPEANOUVEAU:=Copy(St,469,3);IE_TYPEMVT:=Copy(St,472+Decal,3);
    //'IE_QUALIFECRQTE1="'+Copy(St,563+Decal,3) ;  ',IE_QUALIFECRQTE2="'+Copy(St,566+Decal,3) ;
    if Trim(Copy(St,475+Decal,1))<>'' then IE_TYPEANALYTIQUE:=Copy(St,475+Decal,1) ;
    END ;
  if (TypeMvt<>'A') and (IE_NUMECHE<=1) then Inc(NumLignePiece) ;
  IE_NUMLIGNE:=NumLignePiece ;
  END ;
END ;

Procedure AlimTabCpt(Var InfoImp : TInfoImport) ;
Var CptLu,CptRenvoye : TCptLu ;

BEGIN
Fillchar(CptLu,SizeOf(CptLu),#0) ;
CptLu.Cpt:=Trim(MvtImport^.IE_GENERAL)   ;
If CptLu.Cpt<>'' Then
   BEGIN
   AlimTabCptLu(CptLu,InfoImp.GenLu) ;
   END ;
CptLu.Cpt:=Trim(MvtImport^.IE_AUXILIAIRE) ;
If CptLu.Cpt<>'' Then
   BEGIN
   CptLu.Collectif:=Trim(MvtImport^.IE_GENERAL) ;
   AlimTabCptLu(CptLu,InfoImp.AuxLu) ;
   END ;
CptLu.Cpt:=Trim(MvtImport^.IE_JOURNAL)   ;
If CptLu.Cpt<>'' Then
   BEGIN
   AlimTabCptLu(CptLu,InfoImp.JalLu) ;
   END ;
CptLu.Cpt:=Trim(MvtImport^.IE_SECTION)   ;
If CptLu.Cpt<>'' Then
   BEGIN
   CptLu.Collectif:=Trim(MvtImport^.IE_GENERAL) ;
   CptRenvoye:=AlimTabCptLu(CptLu,InfoImp.AnaLu) ;
   RetoucheTabCptLu(CptRenvoye,InfoImp.GenLu) ;
   END ;
END ;

procedure ImportLigne(St : String ; QAjoute : TQuery ; Var TabJalSoc : TTabCptLu ; Var InfoImp : TInfoImport) ;
var Sens     : String ;
    Axe      : String5 ;
    TypeCpte : String1 ;
    FFic : TFichierBase ;
    Lettrable : Boolean ;
    LgCpte,Decal : Byte ;
BEGIN
if Pos(SepLigneIE,St)=1 then Exit ;
Decal:=0 ; LgCpte:=13 ; Lettrable:=False ;
if (Format='HAL') or (Format='HLI') then
  BEGIN
  //if (Format='SN2') and (TypeMvt='X') then Exit ;
  if not VH^.ImportRL then LgCpte:=17 ;
  Jal:=Copy(St,1,3) ;
  DateC:=Format_Date_HAL(Copy(St,4,8)) ;
  MvtImport^:=InitMvtImp^ ;
  With MvtImport^ do
      BEGIN
      IE_CHRONO:=NbLigEnCours ; IE_JOURNAL:=Jal ; IE_DATECOMPTABLE:=DateC ;
      TypeCpte:=Copy(St,13,1) ; IE_GENERAL:=Copy(St,14,LgCpte) ;
      if ((Lequel='FEC') or ((Lequel='FBA') and (Format='HAL'))) and not VH^.ImportRL then Inc(Decal,4) ;
      TypeMvt:=Copy(St,27+Decal,1) ;
      if (TypeMvt='L') then IE_TYPEECR:='L' else IE_TYPEECR:='E' ;
      if (TypeMvt<>'A') then NumVentil:=0 ;
      case TypeMvt[1] of
        'X','E','L': IE_AUXILIAIRE:=Copy(St,28+Decal,LgCpte) ;
        'A': BEGIN IE_ANA:='X' ; IE_TYPEECR:='A' ; END ;
        END ;
      if ((Lequel='FEC') or ((Lequel='FBA') and (Format='HAL'))) and not VH^.ImportRL then Inc(Decal,4) ;
      IE_REFINTERNE:=Copy(St,41+Decal,13) ;
      IE_LIBELLE:=Copy(St,54+Decal,25) ;
      Sens:=Copy(St,88+Decal,1) ;
      if Sens='D' then IE_DEBIT:=Valeur(StPoint(Copy(St,89+Decal,20))) else
        if Sens='C' then IE_CREDIT:=Valeur(StPoint(Copy(St,89+Decal,20))) ;
      if (TypeMvt<>'A') then
      NatP:=Copy(St,12,2) ;
      if (Format='HAL') and (Lequel<>'FBA') and (Trim(Copy(St,748+Decal,3))<>'') then NatP:=Copy(St,748+Decal,3) ;
      IE_NATUREPIECE:=NatP ;
      QualP:=Copy(St,109+Decal,1) ;
      NumPiece:=Round(Valeur(Copy(St,110+Decal,7))) ;
      END ;
  END else
  BEGIN
  TypeMvt:=Copy(St,25,1) ;
  // On ne prend que les échéances au format SN2
  if (Format='SN2') and (TypeMvt='X') then Exit ;
  Sens:=Copy(St,84,1) ; Jal:=Copy(St,1,3) ;
  DateC:=Format_Date(Copy(St,4,6)) ;
  if (TypeMvt<>'A') then NatP:=Copy(St,10,2) ;
  MvtImport^:=InitMvtImp^ ;
  With MvtImport^ do
    BEGIN
    IE_CHRONO:=NbLigEnCours ; IE_JOURNAL:=Jal ; IE_DATECOMPTABLE:=DateC ; IE_NATUREPIECE:=NatP ;
    TypeCpte:=Copy(St,11,1) ;
    IE_GENERAL:=Copy(St,12,13) ;
    case TypeMvt[1] of
      'X','E','L': IE_AUXILIAIRE:=Copy(St,26,13) ;
      'A': BEGIN IE_ANA:='X' ; IE_TYPEECR:='A' ; END ;
      END ;
    IE_REFINTERNE:=Copy(St,39,13) ;
    IE_LIBELLE:=Copy(St,52,25) ;
    if Sens='D' then IE_DEBIT:=Valeur(StPoint(Copy(St,85,20))) else
      if Sens='C' then IE_CREDIT:=Valeur(StPoint(Copy(St,85,20))) ;
    if (TypeMvt='L') then IE_TYPEECR:='L' else IE_TYPEECR:='E' ;
    if (TypeMvt<>'A') then NumVentil:=0 ;
    QualP:=Copy(St,105,1) ;
    NumPiece:=Round(Valeur(Copy(St,106,7))) ;
    END ;
  END ;
With MvtImport^do
  BEGIN
  if NumPiece=0 then NumPiece:=1 ;
  IE_NUMPIECE:=NumPiece ;
  if (TypeMvt<>'A') and (NatJal2(Jal,TabJalSoc)='ANO') then IE_ECRANOUVEAU:='OAN' ;
  if (QualP='N') and (QualifPiece<>'') and (IE_ECRANOUVEAU<>'OAN') then QualP:=QualifPiece ;
  IE_QUALIFPIECE:=QualP ;
  // Maj de la longueur des comptes du fichier / celles paramétrées.
  IE_GENERAL:=BourreOuTronque(IE_GENERAL,fbGene) ;
  IE_AUXILIAIRE:=BourreOuTronque(IE_AUXILIAIRE,fbAux) ;
  if (IE_TYPEECR='A') then
    BEGIN
    FFic:=fbAxe1 ;
    if (IE_AXE='A2') then FFic:=fbAxe2 else
    if (IE_AXE='A3') then FFic:=fbAxe3 else
    if (IE_AXE='A4') then FFic:=fbAxe4 else
    if (IE_AXE='A5') then FFic:=fbAxe5 ;
    IE_SECTION:=BourreOuTronque(IE_SECTION,FFic) ;
    END ;
  if (TypeMvt[1]<>'E') then InitRuptEche ;
  RecupInfosCpte(IE_GENERAL,IE_AUXILIAIRE) ;
  if (Format<>'HAL') and (Format<>'HLI') and (Pointable='X') then BEGIN IE_ECHE:='X' ; IE_MODEPAIE:=TrouveModePaie('',Valeur(StPoint(Copy(St,85,20))),TRUE) ; END else
   if ((Lequel='FBA') or (Format='HLI')) and (Pointable='X') then BEGIN IE_ECHE:='X' ; IE_MODEPAIE:=TrouveModePaie('',Valeur(StPoint(Copy(St,89+Decal,20))),TRUE) ; END ;
  IE_REGIMETVA:=RegimeTva ;
  if IE_AUXILIAIRE='' then IE_TYPEMVT:=QuelTypeMvt(IE_GENERAL, NatureCompte(IE_GENERAL,True,Lettrable),NatP)
                      else IE_TYPEMVT:=QuelTypeMvt(IE_AUXILIAIRE,NatureCompte(IE_AUXILIAIRE,False,Lettrable),NatP) ;
  if IsVentilable(IE_GENERAL,Axe,TypeCpte) then IE_ANA:='X' ;
  // On n'importe que les mouvements rapprochés
  if (Format='RAP') and (Copy(St,133,1)='0') then Exit ; // IE_REFPOINTAGE:=Copy(St,133,1) ;
  if OkRupt then Exit ;
  END ;
OkRupt:=TestBreak(InfoImp) ;
if OkRupt then Exit ;
if (Format<>'HAL') or (Lequel='FBA') then ImporteZones (TypeMvt,St,False,Lettrable,TabJalSoc)
                                     else ImporteZones (TypeMvt,St,True,Lettrable,TabJalSoc) ;
if (TypeMvt<>'L') and (TypeMvt<>'A') then VideLettrage(Lettrable) ;
if (Format<>'HAL') or ((Lequel='FBA') and (Format<>'EDI')) then AjouteDevise(MvtImport^.IE_DEVISE) ;
MemoriseMontantGene ;
AjoutePourcentages ;
//MAJContreparties ;
AjouteMvt(Qajoute,InfoImp) ;
AlimTabCpt(InfoImp) ;
{For i:=0 to TImpEcr.FieldCount-1 do
    BEGIN
    NomChp:=Copy(TImpEcr.Fields[i].Name,4,Length(TImpEcr.Fields[i].Name)-3) ;
    if MajValCombo(RecupJoin(NomChp,''),Where,Trim(TImpEcr.Fields[i].Value),Msgbox.Mess[10],Msgbox.Mess[11])>0 then ;
    END ;
}
END ;

// Importation du fichier paramétrable

procedure RechercheMultiAxe(Q : TQuery) ;
var QM : TQuery ;
BEGIN
// Recherche des lignes de ventile multi-axe
// Chaque ligne géné est suivie de sa ventile analytique
QM:=TQuery.Create(Application) ; QM.DataBaseName:='SOC' ;
QM.SQL.Add('SELECT * FROM IMPECR WHERE IE_JOURNAL=:P1 AND IE_DATECOMPTABLE=:P2 AND IE_NATUREPIECE=:P3'+
           ' AND IE_NUMPIECE=:P4 AND IE_QUALIFPIECE=:P5 AND IE_GENERAL=:P6 AND IE_NUMLIGNE>:P7') ;
ChangeSQL(QM) ; QM.Prepare ;
Q.First ;
While not Q.Eof do
  BEGIN
  if Q.FindField('IE_TYPEECR').AsString='A' then
    BEGIN
    QM.Close ;
    QM.Params[0].AsString:=Q.FindField('IE_JOURNAL').AsString ;
    QM.Params[1].AsDateTime:=Q.FindField('IE_DATECOMPTABLE').AsDateTime ;
    QM.Params[2].AsString:=Q.FindField('IE_NATUREPIECE').AsString ;
    QM.Params[3].AsInteger:=Q.FindField('IE_NUMPIECE').AsInteger ;
    QM.Params[4].AsString:=Q.FindField('IE_QUALIFPIECE').AsString ;
    QM.Params[5].AsString:=Q.FindField('IE_GENERAL').AsString ;
    QM.Params[6].AsInteger:=Q.FindField('IE_NUMLIGNE').AsInteger ;
    QM.Open ;
    if not QM.Eof then
      if QM.FindField('IE_AXE').AsString<>Q.FindField('IE_AXE').AsString then
         BEGIN
         QM.Prior ; QM.Delete ;
         END ;
    END ;
  Q.Next ;
  END ;
END ;

function ImporteFormatParam(NomFic : String ; Fmt : String ; Var InfoImp : TInfoImport) : boolean ;
var Entete  : TFmtEntete ;
    Detail  : TTabFmtDetail ;
    Debut   : Boolean ;
    Q : TQuery ;
    N,NumLigne : Integer ;
    OkRupt : boolean ;
BEGIN
Result:=False ; Fillchar(InfoImp,SizeOf(InfoImp),#0) ;
if not ChargeFormat(Fichier,NomFic,Lequel,'X',Fmt,Entete,Detail,Debut) then Exit ;
Q:=OpenSQL('SELECT * FROM IMPECR',False) ;
N:=1 ; NumLigne:=0 ;
InitMove(NbLig,'') ;
While not Eof(Fichier) do
  BEGIN
  Q.Insert ; InitNew(Q) ;
  Q.FindField('IE_CHRONO').AsInteger:=N ;
  Q.FindField('IE_ETABLISSEMENT').AsString:=VH^.EtablisDefaut ;
  if (Lequel='FOD') then
     BEGIN
     Q.FindField('IE_TYPEECR').AsString:='A' ;
     Q.FindField('IE_TYPEANALYTIQUE').AsString:='X' ;
     END else
     BEGIN
     Q.FindField('IE_TYPEECR').AsString:='E' ;
     END ;
  LireFormat(Fichier,Entete,Detail,Debut,Q) ;
  Jal:=Q.FindField('IE_JOURNAL').AsString ;
  DateC:=Q.FindField('IE_DATECOMPTABLE').AsDateTime ;
  NumP:=Q.FindField('IE_NUMPIECE').AsInteger ;
  QualP:=Q.FindField('IE_QUALIFPIECE').AsString ;
  NatP:=Q.FindField('IE_NATUREPIECE').AsString ;
  If Imp_Methode1 Then Application.ProcessMessages ;
  if (Lequel='FBE') and (OldNumP<>NumP) then
    BEGIN
    OldNumP:=NumP ;
    if AnnuleImport then break ;
    Inc(InfoImp.NbPiece) ; NumLigne:=0 ;
    END else
    BEGIN
    OkRupt:=TestBreak(InfoImp) ;
    if OkRupt then Break ;
    END ;
  Inc(NumLigne) ;
  if (Lequel<>'FBE') and (Lequel<>'FOD') and (Q.FindField('IE_NUMLIGNE').AsInteger=0) then Q.FindField('IE_NUMLIGNE').AsInteger:=NumLigne ;
  if (Q.FindField('IE_TYPEECR').AsString<>'A') then
     BEGIN
     InfoImp.TotDeb:=InfoImp.TotDeb+Q.FindField('IE_DEBIT').AsFloat ;
     InfoImp.TotCred:=InfoImp.TotCred+Q.FindField('IE_CREDIT').AsFloat ;
     END ;
  Inc(InfoImp.NbLigIntegre) ;
  Q.Post ; Inc(N) ;
  MoveCur(False) ;
  END ;
FiniMove ;
//RechercheMultiAxe(Q) ;
Ferme(Q) ;
Result:=True ;
END ;

Function InitQAjoute : TQuery ;
Var Q : TQuery ;
BEGIN
Q:=TQuery.Create(Application) ; Q.DataBaseName:='SOC' ;
Q.Close ; Q.Sql.Clear ;
If QAJParam Then
   BEGIN
   Q.Close ; Q.SQL.Clear ;
   Q.SQL.Add('INSERT INTO IMPECR (') ;

   Q.SQL.Add('IE_CHRONO,IE_AFFAIRE,IE_ETATLETTRAGE,IE_LETTRAGE,IE_LETTREPOINTLCR,IE_LIBELLE,') ;
   Q.SQL.Add('IE_REFEXTERNE,IE_REFINTERNE,IE_REFLIBRE,IE_REFPOINTAGE,IE_REFRELEVE,') ;
   Q.SQL.Add('IE_RIB,IE_SECTION,IE_AXE,IE_JOURNAL,IE_ECRANOUVEAU,IE_ETABLISSEMENT,') ;

   Q.SQL.Add('IE_FLAGECR,IE_MODEPAIE,IE_NATUREPIECE,IE_REGIMETVA,IE_QUALIFPIECE,');
   Q.SQL.Add('IE_QUALIFQTE1,IE_QUALIFQTE2,IE_SOCIETE,IE_TPF,IE_TVA,IE_TYPEANOUVEAU,');
   Q.SQL.Add('IE_TYPEECR,IE_TYPEMVT,IE_DEVISE,IE_AUXILIAIRE,IE_GENERAL,IE_CONTREPARTIEAUX,') ;

   Q.SQL.Add('IE_CONTREPARTIEGEN,IE_NUMECHE,IE_NUMPIECEINTERNE,IE_NUMLIGNE,');
   Q.SQL.Add('IE_NUMPIECE,IE_NUMVENTIL,IE_DATECOMPTABLE,IE_DATEECHEANCE,IE_DATEPAQUETMAX,');
   Q.SQL.Add('IE_DATEPAQUETMIN,IE_DATEPOINTAGE,IE_DATEREFEXTERNE,IE_DATERELANCE,') ;

   Q.SQL.Add('IE_DATETAUXDEV,IE_DATEVALEUR,IE_ORIGINEPAIEMENT,IE_ECHE,IE_ENCAISSEMENT,');
   Q.SQL.Add('IE_CONTROLE,IE_LETTRAGEDEV,IE_OKCONTROLE,IE_SELECTED,IE_INTEGRE,IE_TVAENCAISSEMENT,');
   Q.SQL.Add('IE_TYPEANALYTIQUE,IE_VALIDE,IE_ANA,IE_POURCENTAGE,IE_POURCENTQTE1,');

   Q.SQL.Add('IE_POURCENTQTE2,IE_QTE1,IE_QTE2,IE_QUOTITE,IE_RELIQUATTVAENC,IE_TAUXDEV,');
   Q.SQL.Add('IE_TOTALTVAENC,IE_DEBIT,IE_CREDIT,IE_CREDITDEV,IE_CREDITEURO,');
   Q.SQL.Add('IE_COUVERTURE,IE_COUVERTUREDEV,IE_DEBITDEV,IE_DEBITEURO) ') ;

   Q.SQL.Add('VALUES ( ') ;
   Q.SQL.Add(':IE_CHRONO,:IE_AFFAIRE,:IE_ETATLETTRAGE,:IE_LETTRAGE,:IE_LETTREPOINTLCR,:IE_LIBELLE,') ;
   Q.SQL.Add(':IE_REFEXTERNE,:IE_REFINTERNE,:IE_REFLIBRE,:IE_REFPOINTAGE,:IE_REFRELEVE,') ;
   Q.SQL.Add(':IE_RIB,:IE_SECTION,:IE_AXE,:IE_JOURNAL,:IE_ECRANOUVEAU,:IE_ETABLISSEMENT,') ;

   Q.SQL.Add(':IE_FLAGECR,:IE_MODEPAIE,:IE_NATUREPIECE,:IE_REGIMETVA,:IE_QUALIFPIECE,');
   Q.SQL.Add(':IE_QUALIFQTE1,:IE_QUALIFQTE2,:IE_SOCIETE,:IE_TPF,:IE_TVA,:IE_TYPEANOUVEAU,');
   Q.SQL.Add(':IE_TYPEECR,:IE_TYPEMVT,:IE_DEVISE,:IE_AUXILIAIRE,:IE_GENERAL,:IE_CONTREPARTIEAUX,') ;

   Q.SQL.Add(':IE_CONTREPARTIEGEN,:IE_NUMECHE,:IE_NUMPIECEINTERNE,:IE_NUMLIGNE,');
   Q.SQL.Add(':IE_NUMPIECE,:IE_NUMVENTIL,:IE_DATECOMPTABLE,:IE_DATEECHEANCE,:IE_DATEPAQUETMAX,');
   Q.SQL.Add(':IE_DATEPAQUETMIN,:IE_DATEPOINTAGE,:IE_DATEREFEXTERNE,:IE_DATERELANCE,') ;

   Q.SQL.Add(':IE_DATETAUXDEV,:IE_DATEVALEUR,:IE_ORIGINEPAIEMENT,:IE_ECHE,:IE_ENCAISSEMENT,');
   Q.SQL.Add(':IE_CONTROLE,:IE_LETTRAGEDEV,:IE_OKCONTROLE,:IE_SELECTED,:IE_INTEGRE,:IE_TVAENCAISSEMENT,');
   Q.SQL.Add(':IE_TYPEANALYTIQUE,:IE_VALIDE,:IE_ANA,:IE_POURCENTAGE,:IE_POURCENTQTE1,');

   Q.SQL.Add(':IE_POURCENTQTE2,:IE_QTE1,:IE_QTE2,:IE_QUOTITE,:IE_RELIQUATTVAENC,:IE_TAUXDEV,');
   Q.SQL.Add(':IE_TOTALTVAENC,:IE_DEBIT,:IE_CREDIT,:IE_CREDITDEV,:IE_CREDITEURO,');
   Q.SQL.Add(':IE_COUVERTURE,:IE_COUVERTUREDEV,:IE_DEBITDEV,:IE_DEBITEURO) ') ;
   ChangeSQL(Q) ; Q.Prepare ;
   END Else
   BEGIN
   Q.RequestLive:=TRUE ;
   Q.SQL.Add('Select * From IMPECR where IE_journal="'+W_W+'"') ;
   ChangeSQL(Q) ; Q.Open ;
   END ;
Result:=Q ;
END ;

// Importation des écritures vers le fichier tampon

Function ImporteEcr(Var StErr : String ; Var InfoImp : TInfoImport) : Boolean ;
var St : String ;
    ExpParam,OkFic : boolean ;
    QAjoute : TQuery ;
    TabJalSoc : TTabCptLu ;
BEGIN
InitMove(NbLig,'') ; MoveCur(False) ; StErr:='' ; Result:=TRUE ;
If Imp_Methode1 And (Not QAJParam) Then BEGIN QAjoute:=TQuery.Create(Application) ; QAjoute.DataBaseName:='SOC' ; END
                Else QAjoute:=InitQAjoute ;
ChargejalSoc(TabJalSoc) ;
try
  SourisSablier ;
  New(MvtImport) ; InitMvtImport ;
  BEGINTRANS ;
  ExecuteSQL('DELETE FROM IMPECR') ;
  COMMITTRANS ;
  ExpParam:=(Lequel='FEC') and ((Format<>'SAA') and (Format<>'EDI')
                           and (Format<>'SN2')  and (Format<>'HAL')
                           and (Format<>'HLI')  and (Format<>'MP')
                           and (Format<>'RAP')) or ((Lequel<>'FEC') and (Lequel<>'FBA')) ;
  BEGINTRANS ;
  if ExpParam then
    BEGIN
    ImporteFormatParam(FileName,Format,InfoImp) ;
    END else
    BEGIN
    RemplirLesListes ; OkFic:=TRUE ;
    if (Lequel='FBA') and (Format='EDI') then NbLig:=Nblig*2 ;
    if (Format='SN2') Or (Format='SAA') Then
       BEGIN // GP à prévoir sur HALLEY et HALLEY étendu
       OkFic:=SoldeAZero(FileName,StErr) ;
       If Not OkFic Then
          BEGIN
          Result:=FALSE ;
          If HShowMessage('0;Il y a des erreurs dans le fichier ;Désirez-vous néanmoins continuer le traitement?;Q;YN;N;N;0','','')=mrYes Then
             BEGIN
             OkFic:=TRUE ; Result:=TRUE ;
             END  ;
          END ;
       END ;
    AssignFile(Fichier,FileName) ; {$I-} Reset(Fichier) ; {$I+}
    If OkFic Then
       BEGIN
       Readln(Fichier,St) ; //Société
       if (Lequel='FBA') and (Format<>'EDI') then
         BEGIN
         DateDeb:=Format_DateEDI(Copy(st,77,8)) ;
         DateFin:=Format_DateEDI(Copy(st,85,8)) ;
         END ;
       While not EOF(Fichier) do
         BEGIN
         PrepareQuery(QAjoute) ;
         Readln(Fichier,St) ;
         if OkRupt or Erreur then Break ;
         if ChoixFmt.Ascii then St:=ASCII2ANSI(St) ;
         if (Lequel<>'FBA') then
           BEGIN
           if (Format='EDI') then ImportLigneEDI(St,QAjoute,TabJalSoc,InfoImp) else ImportLigne(St,QAjoute,TabJalSoc,InfoImp) ;
           END else if (Format<>'EDI') then ImportLigne(St,QAjoute,TabJalSoc,InfoImp) else ImportBalance(St,1,QAjoute,TabJalSoc,InfoImp) ;
         MoveCur(False) ;
         END ;
       END ;
    END ;
  finally
    CloseFile(Fichier) ;
    Dispose(MvtImport) ; Dispose(InitMvtImp) ;
    if Erreur then ROLLBACK else COMMITTRANS ;
    SourisNormale ;
  end ;

if (Lequel='FBA') and (Format='EDI') and not AnnuleImport then
  BEGIN
  try
    AssignFile(Fichier,FileName) ; {$I-} Reset(Fichier) ; {$I+}
    Readln(Fichier,St) ; //Société
    New(MvtImport) ; InitMvtImport ; NumLignePiece:=0 ;
    While not EOF(Fichier) do
      BEGIN
      PrepareQuery(QAjoute) ;
      Readln(Fichier,St) ;
      if OkRupt then Break ;
      if ChoixFmt.Ascii then St:=ASCII2ANSI(St) ;
      ImportBalance(St,2,QAjoute,TabJalSoc,InfoImp) ;
      MoveCur(False) ;
      END ;
    finally
    CloseFile(Fichier) ;
    Dispose(MvtImport) ; Dispose(InitMvtImp) ;
    end ;
  END ;
QAjoute.Close ;
If Imp_Methode1 Then QAjoute.Free Else Ferme(QAjoute) ;
FiniMove ;
END ;

{================= Correspondance pour les Combos(TT) ======================}

function OkValExiste (OkCCO : boolean ; StCode,StWhere,StPrefixe,StTable,Val : String ) : boolean ;
var Q  : TQuery ;
    P  : Integer ;
    SQL,StOrder : String ;
BEGIN
Result:=True ;
if (Val='') then Exit ;
if OKCCO then SQL:='Select '+StCode+', '+StPrefixe+'_LIBELLE, '+StPrefixe+'_ABREGE FROM '+StTable
            else SQL:='Select '+StCode+', '+StPrefixe+'_LIBELLE FROM '+StTable ;
   if StWhere<>'' then SQL:=SQL+' Where '+StWhere+' AND '+StCode+'="'+Val+'"' else SQL:=SQL+' Where '+StCode+'="'+Val+'"' ;
   P:=Pos('DISTINCT',StCode) ; if P>0 then StOrder:=Copy(StCode,p+9,50) else StOrder:=StCode ;
   if OKCCO then SQL:=SQL+' ORDER BY '+StPrefixe+'_TYPE, '+StOrder
            else SQL:=SQL+' ORDER BY '+StOrder ;
Q:=OpenSQL(SQL,TRUE) ;
Result:=not (Q.EOF) ;
Ferme(Q) ;
END ;

function RecupJoin(NomChp : String ; MultiTT : string ) : string ;
BEGIN
Result:='' ;
if (NomChp='CREERPAR') or (NomChp='EXPORTE') or (NomChp='UTILISATEUR') or (NomChp='SOCIETE') then Exit ;
if (NomChp='CLEREPARTITION') then
 BEGIN
 if MultiTT='A1' then Result:='ttCleRepart1' else
  if MultiTT='A2' then Result:='ttCleRepart2' else
   if MultiTT='A3' then Result:='ttCleRepart3' else
    if MultiTT='A4' then Result:='ttCleRepart4' else
     if MultiTT='A5' then Result:='ttCleRepart5' ;
 END else Result:=Get_Join('X_'+NomChp) ;
END ;

function MajValCombo(ttCorr: string ; var Where : String ; Val,Lib,Abr : String) : byte ;
var St,StCode,StWhere,StTable,StPrefixe,StLib : String ;
    OkCCO : boolean ;
    i,First : integer ;
    StCle,StVal : Array[1..3] of String ;

BEGIN
Result:=0 ; Where:='' ;
if (ttCorr<>'') then
    BEGIN
    GetCorrespType(ttCorr,StTable,StCode,StWhere,StPrefixe,StLib) ;
    if ((StWhere<>'') AND (Pos('="'+W_W+'"',StWhere)>0)) then
      BEGIN
      StWhere:=FindEtReplace(StWhere,'="'+W_W+'"','<>"'+W_W+'"',TRUE) ;
      END ;
    OkCCO:=((StTable='COMMUN') or (StTable='CHOIXCOD')) ;
    if not OkValExiste(OkCCO,StCode,StWhere,StPrefixe,StTable,Val) then
      BEGIN
      St:=StWhere ;
      StCle[1]:=StCode ;
      StVal[1]:='"'+Val+'"' ;
      i:=1 ;
      While (St<>'') and (i<3) do
        BEGIN
        i:=i+1 ;
        First:=Pos('="',St) ;
        StCle[i]:=Copy(St,1,First-1) ;
        St:=Copy(St,First+2,Length(St)-First+3) ;
        First:=Pos('"',St) ;
        StVal[i]:='"'+Copy(St,1,First-1)+'"' ;
        St:=Copy(St,First+1,Length(St)-First+1) ;
        END ;
      StWhere:=Copy(StVal[2],2,Length(StVal[2])-2) ;
      StCle[3]:=StLib ;
      StVal[3]:='"'+Lib+'"' ;
//Simon      if OkCCO then AddValCombo((StTable='COMMUN'),StWhere,Val,Lib,Abr,'') else
      ValueInsert(StTable,StCle,StVal,TRUE) ;
      Result:=1 ;
      if OkCCO then Result:=Result+1 ;
      Where:=StWhere ;
      END ;
    END ;
END ;

function ImporteLesEcritures(NomFic : String ; Fmt,QualifP : String3 ; Leq : String ; Pos,lgCptes : boolean ; Msg : THMsgBox ;
                             Var InfoImp : TInfoImport) : boolean ;
Var OkOk : Boolean ;
    StErr : String ;
    FmtFic : Integer ;

BEGIN
Result:=False ; OkOk:=TRUE ; StErr:='' ; FillChar(InfoImp,SizeOf(InfoImp),#0) ;
MsgBox:=Msg ; FileName:=NomFic ; Format:=Fmt ; QualifPiece:=QualifP ; Lequel:=Leq ;
// Pas de format SAARI pour la balance
if (Lequel='FBA') and (Format='SAA') then Format:='HAL' ;
Prefixe:='E' ; Table:='ECRITURE' ;
if Lequel='FBE' then BEGIN Prefixe:='BE' ; Table:='BUDECR' ; END else
  if Lequel='FOD' then BEGIN Prefixe:='Y' ; Table:='ANALYTIQ' ;END ;
if (FileName='') and (Msg<>nil) then BEGIN AnnuleImport:=True ; MsgBox.Execute(40,'','') ; Exit ; END ;
Positifs:=Pos ; LgComptes:=LgCptes ; NbLig:=0 ;
Case GoodSoc of
  2 : if (Msg<>nil) then BEGIN MsgBox.Execute(41,'','') ; Exit ; END
                    else BEGIN CompteRenduBatch(2,InfoImp) ; Exit ; END ;
  0 : BEGIN CompteRenduBatch(0,InfoImp) ; Exit ; END ;
  END ;
(*
If Imp_Methode1 Then BEGIN QAjoute:=TQuery.Create(Application) ; QAjoute.DataBaseName:='SOC' ; END
            Else QAjoute:=InitQAjoute ;
*)
NbLigEnCours:=1; InfoImp.NbLigIntegre:=0 ;
InfoImp.TotDeb:=0 ; InfoImp.TotCred:=0 ;
InfoImp.NbPiece:=0 ; NumP:=0 ; OldNumP:=0 ;
DateC:=iDate1900 ; OldDateC:=iDate1900 ;
Jal:='' ; NatP:=''; QualP:='';
OldJal:=''; OldNatP:=''; OldQualP:='' ;
InitRuptEche ;
AnnuleImport:=False ; Erreur:=False ; OkRupt:=False ;
ChargeDevEtSect ;
OkOk:=ImporteEcr(StErr,InfoImp) ;
If Not OkOk Then
   BEGIN
   AnnuleImport:=FALSE ;
   FmtFic:=0 ;
   if Format='HLI' then FmtFic:=1 else
   if Format='HAL' then FmtFic:=2 ;
   if MsgBox<>nil then
     BEGIN
     if MsgBox.Execute(54,'','')=mrYes then VisuLignesErreurs(FileName,StErr,FmtFic) ;
     END ;
   END ;
VideListe(TDev) ; TDev.Free ;
SourisNormale ;
if (Msg<>nil) and (not AnnuleImport) then
  BEGIN
  //ActivePanels(Self,True,False) ;
  if ChoixFmt.Detruire then DeleteFile(FileName) ;
  if (ChoixFmt.CompteRendu) and (InfoImp.NbLigIntegre>0) then
    BEGIN
    //Q:=OPpenSQL('SELECT SUM(IE_DEBIT),SUM(IE_CREDIT) FROM IMPER WHERE IE_TYPEECR<>"A" AND IE_TYPEECR<>"L"',True) ;
    MsgBox.Execute(33,'',' '+IntToStr(InfoImp.NbLigIntegre)+Chr(10)+MsgBox.Mess[34]+' '+IntToStr(InfoImp.NbPiece)+Chr(10)+
    Chr(10)+MsgBox.Mess[35]+AfficheMontant('#,##0.00',V_PGI.SymbolePivot,InfoImp.TotDeb,True)+Chr(10)+MsgBox.Mess[36]+AfficheMontant('#,##0.00',V_PGI.SymbolePivot,InfoImp.TotCred,True)) ;
    END ;
  END ;
(*
QAjoute.Close ;
If Imp_Methode1 Then QAjoute.Free Else Ferme(Q) ;
*)
Result:=not AnnuleImport ;
END ;

//----------------------------------------------------------------------
//-------------------- Intégration des écritures -----------------------
//----------------------------------------------------------------------

procedure ActiveTotaux(Ok : Boolean ; FAssImp : TForm ; Var InfoImp : TInfoImport) ;
var ttJal,ttNatP : string ;
begin
if FAssImp=nil then Exit ;
if not Ok then BEGIN InfoImp.NbPiece:=0 ; InfoImp.TotDeb:=0 ; InfoImp.TotCred:=0 ; END ;
//if NbPiece<=1 then THLabel(FAssImp.FindComponent('LNbPiece')).Caption:=THMsgBox(FAssImp.FindComponent('HM')).Mess[7] else THLabel(FAssImp.FindComponent('LNbPiece')).Caption:=THMsgBox(FAssImp.FindComponent('HM')).Mess[8] ;
ttJal:='ttJournal' ;
if Prefixe='BE' then ttJal:='ttBudJal' ;
ttNatP:='ttNaturePiece' ;
if Prefixe='BE' then ttNatP:='ttNatEcrBud' ;
THLabel(FAssImp.FindComponent('DateP')).Caption:=DateToStr(QImpEcr.FindField('IE_DATECOMPTABLE').AsDateTime) ;
THLabel(FAssImp.FindComponent('JalP')).Caption:=RechDom(ttJal,QImpEcr.FindField('IE_JOURNAL').AsString,False) ;
THLabel(FAssImp.FindComponent('NatP')).Caption:=RechDom(ttNatP,QImpEcr.FindField('IE_NATUREPIECE').AsString,False) ;
THLabel(FAssImp.FindComponent('PieceP')).Caption:=IntToStr(QImpEcr.FindField('IE_NUMPIECE').AsInteger) ;
if (Prefixe='BE') then
  BEGIN
  THLabel(FAssImp.FindComponent('LigneP')).Visible:=False ;
  THLabel(FAssImp.FindComponent('Slash')).Visible:=False ;
  THLabel(FAssImp.FindComponent('THLigne')).Visible:=False ;
  THLabel(FAssImp.FindComponent('THSlash')).Visible:=False ;
  END else
  if (Prefixe='Y') then THLabel(FAssImp.FindComponent('LigneP')).Caption:=IntToStr(QImpEcr.FindField('IE_NUMVENTIL').AsInteger)
                   else THLabel(FAssImp.FindComponent('LigneP')).Caption:=IntToStr(QImpEcr.FindField('IE_NUMLIGNE').AsInteger) ;
THLabel(FAssImp.FindComponent('ZNb')).Caption:=IntToStr(InfoImp.NbPiece) ;
THLabel(FAssImp.FindComponent('ZNbLig')).Caption:=IntToStr(NbLigEnCours) ;
THNumEdit(FAssImp.FindComponent('ZTotDeb')).Text:=StrfMontant(Abs(InfoImp.TotDeb),20,V_PGI.OkDecV,V_PGI.SymbolePivot,True) ;
THNumEdit(FAssImp.FindComponent('ZTotCred')).Text:=StrfMontant(Abs(InfoImp.TotCred),20,V_PGI.OkDecV,V_PGI.SymbolePivot,True) ;
// GG ??? :
//Application.ProcessMessages ;
END ;

{==================== Mémoire des natures de comptes/ journaux  =====================}

// Mise à jour des Totaux dans les fiches

procedure MajTotCompte(Compte : String17 ; OkANouv : boolean) ;
var i,k      : integer ;
    TTotCpte : TFTotCpte ;
    Exo : String3 ;
BEGIN
if ((DebitMvt=0) and (CreditMvt=0)) then Exit ;
k:=TTotalCpte.IndexOf(Trim(Compte)) ;
if k=-1 then Exit ;
TTotCpte:=TFTotCpte(TTotalCpte.Objects[k]) ;
TTotCpte.LastDate:=OldDateC ;
TTotCpte.LastNum:=NumP ;
TTotCpte.LastNumL:=NumL ;
Exo:=QUELEXODT(OldDateC) ;
for i:=1 to 5 do
  BEGIN
  if ((Exo=VH^.EnCours.Code) and (i=4)) or ((Exo=VH^.Suivant.Code) and (i=3)) then Continue ;
  // Pas les à nouveau
  if (i=5) and not OkANouv then Break ;
  if i=1 then
    BEGIN
    TTotCpte.LesMontants[i,1]:=DebitMvt ;
    TTotCpte.LesMontants[i,2]:=CreditMvt ;
    END else
    BEGIN
    TTotCpte.LesMontants[i,1]:=TTotCpte.LesMontants[i,1]+DebitMvt ;
    TTotCpte.LesMontants[i,2]:=TTotCpte.LesMontants[i,2]+CreditMvt ;
    END ;
  END ;
TTotalCpte.Objects[k]:=TTotCpte ;
END ;

procedure MajTotSection(Section : String17 ; Axe : String3 ; OkANouv : boolean) ;
var TSection : TFSection ;
    i,j : Integer ;
    Exo : String3 ;
BEGIN
if ((DebitMvt=0) and (CreditMvt=0)) then Exit ;
for i:=0 to TListeSection.Count-1 do
  BEGIN
  TSection:=TFSection(TListeSection[i]) ;
  if (TSection.Section=Section) and (TSection.Axe=Axe) then
    BEGIN
    TSection.LastDate:=OldDateC ;
    TSection.LastNum:=NumP ;
    TSection.LastNumL:=NumL ;
    Exo:=QUELEXODT(OldDateC) ;
    for j:=1 to 5 do
      BEGIN
      if Exo=VH^.EnCours.Code then
        BEGIN
        if j=4 then Continue ;
        END else
        if Exo=VH^.Suivant.Code then if j=3 then Continue ;
      // Journal non à nouveau
      if (j=5) and not OkANouv then Break ;
      if j=1 then
        BEGIN
        TSection.LesMontants[j,1]:=DebitMvt ;
        TSection.LesMontants[j,2]:=CreditMvt ;
        END else
        BEGIN
        TSection.LesMontants[j,1]:=TSection.LesMontants[j,1]+DebitMvt ;
        TSection.LesMontants[j,2]:=TSection.LesMontants[j,2]+CreditMvt ;
        END ;
      END ;
    END ;
  TListeSection.Items[i]:=TSection ;
  END ;
END ;

procedure MajTotJal ;
var i,j      : integer ;
    TJal     : TFJal ;
    Exo      : String3 ;
BEGIN
// Journaux
if ((DebitMvt=0) and (CreditMvt=0)) then Exit ;
i:=TListeJal.IndexOf(Trim(OldJal)) ;
if i<0 then Exit ;
TJal:=TFJal(TListeJal.Objects[i]) ;
TJal.LastDate:=OldDateC ;
TJal.LastNum:=NumP ;
Exo:=QUELEXODT(OldDateC) ;
for j:=1 to 4 do
  BEGIN
  if Exo=VH^.EnCours.Code then
    BEGIN
    if j=4 then Continue ;
    END else if Exo=VH^.Suivant.Code then if j=3 then Continue ;
  if j=1 then
    BEGIN
    TJal.LesMontants[j,1]:=DebitMvt ;
    TJal.LesMontants[j,2]:=CreditMvt ;
    END else
    BEGIN
    TJal.LesMontants[j,1]:=TJal.LesMontants[j,1]+DebitMvt ;
    TJal.LesMontants[j,2]:=TJal.LesMontants[j,2]+CreditMvt ;
    END ;
  END ;
TListeJal.Objects[i]:=TJal ;
END ;

{======================= Contôles/Rupture en intégration ==========================}

Function EnRupture(T : TDataSet) : Boolean ;
BEGIN
Result:=False ;
// rupture des écritures budgétaires
if (Prefixe='BE') then
  BEGIN
  If   (T.FindField('IE_JOURNAL').AsString<>OldJal)
    or (T.FindField('IE_NATUREPIECE').AsString<>OldNat)
    or (T.FindField('IE_NUMPIECE').AsInteger<>OldNumP)
    or (T.FindField('IE_QUALIFPIECE').AsString<>OldQualP) then
      BEGIN
      OldJal:=T.Findfield('IE_JOURNAL').AsString ;
      OldDateC:=T.Findfield('IE_DATECOMPTABLE').AsDateTime ;
      OldNat:=Trim(T.Findfield('IE_NATUREPIECE').AsString) ;
      OldNumP:=T.Findfield('IE_NUMPIECE').AsInteger ;
      OldQualP:=T.Findfield('IE_QUALIFPIECE').AsString ;
      Result:=True ;
      END ;
  END else
// rupture des écritures courantes / Analytiques
  BEGIN
  If   (T.FindField('IE_JOURNAL').AsString<>OldJal)
    or (T.FindField('IE_NATUREPIECE').AsString<>OldNat)
    or (T.FindField('IE_DATECOMPTABLE').AsDateTime<>OldDateC)
    or (T.FindField('IE_NUMPIECE').AsInteger<>OldNumP)
    or (T.FindField('IE_QUALIFPIECE').AsString<>OldQualP) then
      BEGIN
      OldJal:=T.Findfield('IE_JOURNAL').AsString ;
      OldNat:=Trim(T.Findfield('IE_NATUREPIECE').AsString) ;
      OldDateC:=T.Findfield('IE_DATECOMPTABLE').AsDateTime ;
      OldNumP:=T.Findfield('IE_NUMPIECE').AsInteger ;
      OldQualP:=T.Findfield('IE_QUALIFPIECE').AsString ;
      Result:=True ;
      END ;
  END ;
//if Result then Inc(NbPiece);
END;

procedure RecupPiece(Num : Integer ; var ImpPiece : Piece) ;
Var QImp : Tquery ;
BEGIN
QImp:=OpenSQL('SELECT IE_JOURNAL,IE_DATECOMPTABLE,IE_NATUREPIECE,IE_NUMPIECE,IE_QUALIFPIECE FROM IMPECR WHERE IE_CHRONO='+IntToStr(Num),True) ;
if Not QImp.Eof then
  BEGIN
  With ImpPiece do
    BEGIN
    Jal:=QImp.Fields[0].AsString ;
    DateP:=QImp.Fields[1].AsDateTime  ;
    Nat:=QImp.Fields[2].AsString ;
    NumP:=QImp.Fields[3].AsInteger ;
    QualP:=QImp.Fields[4].AsString ;
    END ;
  END;
Ferme(QImp) ;
END ;

procedure MajImpErr(ListeErreurs : TStringList ; LaTotale : boolean) ;
var Nb,i,OldChrono : integer ;
    ImpError : TERRORIMP ;
    ImpPiece,OldPiece : Piece ;
    St,QuelleErreur,OldErreur : String ;
BEGIN
OldChrono:=0 ; OldErreur:='' ;
Nb:=ListeErreurs.Count ;
InitMove(Nb,'') ;
for i:=0 to ListeErreurs.Count-1 do
  BEGIN
  if ListeErreurs.Objects[i]<>Nil then
    BEGIN
    ImpError:=TERRORIMP(ListeErreurs.Objects[i]) ;
    QuelleErreur:=ImpError.NumErreur ;
    //if (QuelleErreur<>'G') and (QuelleErreur<>'A') then QuelleErreur:='-' ;
    QuelleErreur:='-' ;
    //if (QuelleErreur='A') then St:=' AND IE_TYPEECR="A" AND IE_ANA="X"' else  St:=' AND IE_OKCONTROLE="X"' ;
    St:=' AND IE_OKCONTROLE="X"' ;
    RecupPiece(ImpError.NumChrono,ImpPiece) ;
    if (ImpPiece.DateP<>OldPiece.DateP) or (ImpPiece.Jal<>OldPiece.Jal)
       or (ImpPiece.NumP<>OldPiece.NumP) or (ImpPiece.Nat<>OldPiece.Nat)
       or (ImpPiece.QualP<>OldPiece.QualP) then
       with ImpPiece do
          ExecuteSQL('UPDATE IMPECR SET IE_OKCONTROLE="'+QuelleErreur+'" WHERE IE_JOURNAL="'+ImpPiece.Jal+'" AND IE_DATECOMPTABLE="'
                     +USDateTime(ImpPiece.DateP)+'" AND IE_NUMPIECE='+IntToStr(ImpPiece.NumP)+' AND IE_QUALIFPIECE="'+ImpPiece.QualP
                     +'" AND IE_NATUREPIECE="'+ImpPiece.Nat+'"'+St)  ;
    if LaTotale then
      BEGIN
      if (ImpError.NumChrono<>OldChrono) or (OldErreur<>'-') then ExecuteSQL('UPDATE IMPECR SET IE_OKCONTROLE="'+ImpError.NumErreur+'" WHERE IE_CHRONO='+IntToStr(ImpError.NumChrono)) ;
      END ;
    OldChrono:=ImpError.NumChrono ; OldErreur:=QuelleErreur ;
    if (ImpPiece.DateP<>OldPiece.DateP) or (ImpPiece.Jal<>OldPiece.Jal)
       or (ImpPiece.NumP<>OldPiece.NumP) or (ImpPiece.Nat<>OldPiece.Nat)
       or (ImpPiece.QualP<>OldPiece.QualP) then
       BEGIN
       OldPiece.DateP:=ImpPiece.DateP ;
       OldPiece.Jal:=ImpPiece.Jal ;
       OldPiece.NumP:=ImpPiece.NumP ;
       OldPiece.Nat:=ImpPiece.Nat ;
       OldPiece.QualP:=ImpPiece.QualP ;
       END ;
    END ;
  MoveCur(False) ;
  END ;
FiniMove ;
END ;

{======================= Vérification des doublons ============================}

procedure MajDoublons(var TLD : TList) ;
var Nb,i : Integer ;
    TDoublon : TFDoublon ;
    OldJal,OldQualifP,SQL,SQLN: string ;
BEGIN
Nb:=TLD.Count ; OldJal:='' ; OldQualifP:='' ;SQL:='' ; SQLN:='' ;
InitMove(Nb,'') ; MoveCur(False) ;
For i:=0 to TLD.Count-1 do
  BEGIN
  TDoublon:=TFDoublon(TLD[i]) ;
  if (TDoublon.JOURNAL=OldJal) and (TDoublon.QUALIFPIECE=OldQualifP) then
      BEGIN
      SQLN:=SQLN+' OR IE_NUMPIECE='+IntToStr(TDoublon.NUMPIECE) ;
      END else
      BEGIN
      if (SQL<>'') then ExecuteSQL(SQL+SQLN+')') ;
      SQL:='UPDATE IMPECR SET IE_OKCONTROLE="D" WHERE'
          +' (IE_JOURNAL="'+TDoublon.JOURNAL+'"'
          +' AND IE_QUALIFPIECE="'+TDoublon.QUALIFPIECE+'")' ;
      SQLN:=' AND (IE_NUMPIECE='+IntToStr(TDoublon.NUMPIECE) ;
      OldJal:=TDoublon.JOURNAL ; OldQualifP:=TDoublon.QUALIFPIECE ;
      END ;
  if i>200 then
      BEGIN
      if (SQL<>'') then ExecuteSQL(SQL+SQLN+')') ;
      SQL:='' ; SQLN:='' ;
      END ;
  MoveCur(False) ;
  END ;
if (SQL<>'') then ExecuteSQL(SQL+SQLN+')') ;
VideListe(TLD) ; TLD.Free; FiniMove ;
END ;

Function ExisteDoublonMvt(ExistMvt : TExistMvt ) : TFDoublon ;
var Q : TQuery ;
    SQL,Pref,Tabl : String ;
    TDoublon : TFDoublon ;
BEGIN
Result:=nil ;
if ExistMvt.General='' then Exit ;
if ExistMvt.ANA='X' then BEGIN  Pref:='Y_' ; Tabl:='ANALYTIQ' ; END
  else BEGIN Pref:='E_' ; Tabl:='ECRITURE' ; END ;
SQL:='SELECT '+Pref+'JOURNAL,'+Pref+'DATECOMPTABLE,'+Pref+'NATUREPIECE,'+Pref+
     'NUMEROPIECE,'+Pref+'QUALIFPIECE FROM '+Tabl+' WHERE ' ;
SQL:=SQL+RecupWhereExistMvt(ExistMvt) ;
Q:=OpenSQL(SQL,True) ;
TDoublon:=nil ;
if not Q.Eof then
  BEGIN
  TDoublon:=TFDoublon.Create ;
  TDoublon.IMPORT:=False ;
  TDoublon.JOURNAL:=Q.Fields[0].asString ; TDoublon.DATEC:=Q.Fields[1].AsDateTime ;
  TDoublon.NATUREPIECE:=Q.Fields[2].asString ; TDoublon.NUMPIECE:=Q.Fields[3].AsInteger ;
  TDoublon.QUALIFPIECE:=Q.Fields[4].AsString ;
  END ;
Result:=TDoublon ;
Ferme(Q) ;
END ;

function ExisteDejaDoublon(TDoublon : TFDoublon ; TLD : TList) : boolean ;
var TDBL : TFDoublon ;
    i : integer ;
BEGIN
Result:=True ;
for i:=0 to TLD.Count-1 do
  BEGIN
  TDBL:=TFDoublon(TLD[i]) ;
  if (TDBL.IMPORT=TDoublon.IMPORT) and (TDBL.JOURNAL=TDoublon.JOURNAL) and (TDBL.DATEC=TDoublon.DATEC) and
     (TDBL.NATUREPIECE=TDoublon.NATUREPIECE) and (TDBL.NUMPIECE=TDoublon.NUMPIECE) and
     (TDBL.QUALIFPIECE=TDoublon.QUALIFPIECE) then Exit ;
  END ;
Result:=False ;
END ;


procedure VerifDoublons ;
var QImp       : TQuery ;
    NumPiece,
    NbErreur,
    NumChrono  : Integer ;
    Trouve     : Boolean ;
    ExistMvt   : TExistMvt ;
    TDoublon   : TFDoublon ;
    X          : DelInfo ;
    TLD        : TList ;
    NatP,QualP : String3 ;
    Fic : String ;
BEGIN
//ActiveAttente(4,True) ;
QImp:=OpenSQL('SELECT IE_JOURNAL,IE_DATECOMPTABLE,IE_GENERAL,IE_DEBIT,IE_CREDIT,IE_REFINTERNE,IE_DEVISE,IE_NATUREPIECE,IE_NUMPIECE,IE_QUALIFPIECE,IE_CHRONO,IE_NUMECHE,IE_AUXILIAIRE FROM IMPECR '
             +' WHERE IE_SELECTED="X" AND IE_TYPEECR<>"L" AND IE_TYPEECR<>"A" AND IE_OKCONTROLE="X"'
             +' ORDER BY IE_JOURNAL,IE_DATECOMPTABLE,IE_GENERAL,IE_AUXILIAIRE,IE_REFINTERNE,IE_DEBIT,IE_CREDIT,IE_ECHE,IE_NATUREPIECE,IE_QUALIFPIECE,IE_NUMPIECE',True) ;
InitMove(RecordsCount(QImp)+2,'') ; MoveCur(False) ;
TLD:=TList.Create ; InitExistMvt(ExistMvt) ;
TErreur:=TList.Create ; TErreur.Clear ; NbErreur:=0 ;
While not QImp.Eof do
  BEGIN
  NumChrono:=QImp.Findfield('IE_CHRONO').AsInteger ;
  With ExistMvt do
    BEGIN
    JOURNAL:=QImp.Fields[0].AsString ;
    DATECOMPTABLE:=QImp.Fields[1].AsDateTime ;
    GENERAL:=QImp.Fields[2].AsString ; AUXIANA:=QImp.Fields[12].AsString ;
    DEBIT:=QImp.Fields[3].AsFloat ; CREDIT:=QImp.Fields[4].AsFloat ;
    REFINTERNE:=QImp.Fields[5].AsString ;
    NUMECHEVENTIL:=QImp.Fields[11].AsInteger ;
    DEVISE:=QImp.Fields[6].AsString ;
    END ;
  If Imp_Methode1 Then Application.ProcessMessages ; Trouve:=True ;
  While (not QImp.Eof) and Trouve do
    BEGIN
    NatP:=QImp.Fields[7].AsString ; NumPiece:=QImp.Fields[8].AsInteger ;
    QualP:=QImp.Fields[9].AsString ;
    QImp.Next ; MoveCur(False) ;
    With ExistMvt do
      BEGIN
      if ((JOURNAL=QImp.Fields[0].AsString) and (DATECOMPTABLE=QImp.Fields[1].AsDateTime) and
         (GENERAL=QImp.Fields[2].AsString) and (DEBIT=QImp.Fields[3].AsFloat) and (CREDIT=QImp.Fields[4].AsFloat) and
         (REFINTERNE=QImp.Fields[5].AsString) and (NUMECHEVENTIL=QImp.Fields[11].AsInteger) and
         (DEVISE=QImp.Fields[6].AsString) and ((NatP=QImp.Fields[7].AsString) and (QualP=QImp.Fields[9].AsString) and
         (NumPiece<>QImp.Fields[8].AsInteger))) then
          BEGIN
          TDoublon:=TFDoublon.Create ;
          TDoublon.IMPORT:=True ;
          TDoublon.JOURNAL:=QImp.Fields[0].asString ; TDoublon.DATEC:=QImp.Fields[1].AsDateTime ;
          TDoublon.NATUREPIECE:=QImp.Fields[7].asString ; TDoublon.NUMPIECE:=QImp.Fields[8].AsInteger ;
          TDoublon.QUALIFPIECE:=QImp.Fields[9].AsString ;
          END else
          BEGIN
          TDoublon:=ExisteDoublonMvt(ExistMvt) ;
          END ;
      if (TDoublon=nil) then Trouve:=False else
        if not ExisteDejaDoublon(TDoublon,TLD) then
          BEGIN
          NbErreur:=NbErreur+1 ;
          TLD.Add(TDoublon) ;
          X:=DelInfo.Create ; X.LeCod:=IntToStr(NbErreur) ; X.LeLib:=IntToStr(NumChrono);
          if TDoublon.IMPORT then Fic:=MsgBox.Mess[44] else Fic:=V_PGI.NomSociete+' '+V_PGI.CodeSociete +' :' ;
          X.LeMess:=Fic+' '+TDoublon.JOURNAL+' : '+DateToStr(TDoublon.DATEC)+' : '+ IntToStr(TDoublon.NumPiece)+' : '+TDoublon.NaturePiece+' : '+TDoublon.QualifPiece ;
          TErreur.Add(X) ;
          END else Trouve:=False ;
      QImp.Next ; MoveCur(False) ;
      END ;
    END ;
  END ;
Ferme(QImp) ; FiniMove ;
//ActiveAttente(4,False) ;
if (TErreur<>nil) and (TErreur.Count>0) then RapportDeSuppression(TErreur,6) ;
VideListe(TErreur) ; TErreur.Free ; If Imp_Methode1 Then Application.ProcessMessages ;
//ActiveAttente(12,True) ;
MajDoublons(TLD) ;
//ActiveAttente(12,False) ;
END ;

{======================= Lancement de l'intégration ==========================}

procedure MajSelection(OkRejetes : boolean) ;
var St : string ;
BEGIN
//ActiveTotaux(False) ;
//ActiveAttente(9,True) ;
InitMove(3,'') ; MoveCur(False) ;
St:='UPDATE IMPECR SET IE_SELECTED="X",IE_OKCONTROLE="X"' ;
St:=St+ ' WHERE IE_TYPEECR<>"L" AND IE_OKCONTROLE<>"D"' ;
if OkRejetes then St:=St+' AND IE_OKCONTROLE<>"X"' ;
MoveCur(False) ; ExecuteSQL(St) ; MoveCur(False) ;
//St:='UPDATE IMPECR SET IE_OKCONTROLE="X" WHERE IE_SELECTED="X"' ;
//MoveCur(False) ; ExecuteSQL(St) ; MoveCur(False) ;
FiniMove ; //ActiveAttente(9,False) ;
END ;

function NatureEtNumeroJournal(OldJal : String3) :String3 ;
var TJal : TFJal ;
    i : integer ;
    TTEcr : TTypeEcr ;
BEGIN
Result:='' ;
if (QImpEcr.Findfield('IE_ANA').AsString='-') and  (QImpEcr.Findfield('IE_TYPEANALYTIQUE').AsString='-') then TTEcr:=EcrGen else TTEcr:=EcrAna ;
if Prefixe='BE' then TTEcr:=EcrBud ;
if Prefixe='Y' then TTEcr:=EcrAna ;
// même jal/ qualif.
if (JalChange=OldJal) and (QualifPiece=OldQualP) then
  BEGIN
  SetIncNum(TTEcr,Facturier,NumP) ; //Result:=NatureJal ;
  Exit ;
  END ;
// journal dans liste
JalChange:=OldJal ; QualifPiece:=OldQualP ;
i:= TListeJal.IndexOf(Trim(OldJal)) ;
if i>-1 then
  BEGIN
  TJal:=TFJal(TListeJal.Objects[i]) ;
  NatureJal:=TJal.Nature ;
  if (OldQualP='N') then Facturier:=TJal.Normal else Facturier:=TJal.Simul ;
  if (Facturier='') then if FAssImp<>nil then BEGIN THMsgBox(FAssImp.FindComponent('HM')).Execute(43,'',' '+OldJal) ; Erreur:=True ; Exit ; END ;
  SetIncNum(TTEcr,Facturier,NumP) ;
  END ;
END ;

procedure AjouteMvtDansLaBase(QI : TQuery ; Var TabJalSoc : TTabCptLu ; Var InfoImp : TInfoImport) ;
var StTable,Pref,StName : String ;
    CpteAuxi,CpteGene : String17 ;
    SectMvt : String17 ;
    Axe  : String3 ;
    i,k  : integer ;
    OkANouv : boolean ;
BEGIN
SectMvt:='' ; Axe:='' ; DebitMvt:=0 ; CreditMvt:=0 ;
StTable:=Table ; Pref:=Prefixe+'_' ;
If (QImpEcr.Findfield('IE_TYPEECR').AsString='A') then
  BEGIN
  StTable:='ANALYTIQ' ; Pref:='Y_' ;
  END ;
QI.Insert ; InitNew(QI) ;
CpteGene:='' ; CpteAuxi:='' ;
for i:=0 to QI.FieldCount-1 do
  BEGIN
  //PrendsLe:=True ;
  StName:=QI.Fields[i].FieldName ;
  StName:=Copy(StName,Pos('_',StName)+1,50) ;
  if (StName='NUMEROPIECE') then QI.Fields[i].AsInteger:=NumP else
  if (StName='NUMLIGNE') then QI.Fields[i].AsInteger:=QImpEcr.FindField('IE_NUMLIGNE').AsInteger+NumLigneGeneA0 else
  if (StName='NUMECHE') then QI.Fields[i].AsInteger:=QImpEcr.FindField('IE_NUMECHE').AsInteger+NumLigneEcheA0 else
  if (StName='NUMVENTIL') then QI.Fields[i].AsInteger:=QImpEcr.FindField('IE_NUMVENTIL').AsInteger+NumLigneVentilA0 else
  if (StName='EXERCICE') then QI.Fields[i].AsString:=QUELEXODT(OldDateC) else
  if (StName='DATECREATION') then QI.Fields[i].AsDateTime:=V_PGI.DateEntree else
  if (StName='DATEMODIF') then QI.Fields[i].AsDateTime:=NowH else
  if (StName='CREERPAR') then QI.Fields[i].AsString:='IMP' else
  if (StName='CONTROLETVA') then QI.Fields[i].AsString:='RIE' else
  if (StName='VALIDE') then BEGIN if NatureJal='ANO' then QI.Fields[i].AsString:='X' ; END else
    // Ecritures budgétaires
  if (StName='BUDJAL') then QI.Fields[i].AsString:=QImpEcr.FindField('IE_JOURNAL').AsString else
  if (StName='BUDGENE') then QI.Fields[i].AsString:=QImpEcr.FindField('IE_GENERAL').AsString else
  if (StName='BUDSECT') then QI.Fields[i].AsString:=QImpEcr.FindField('IE_SECTION').AsString else
  if (StName='NATUREBUD') then QI.Fields[i].AsString:=QImpEcr.FindField('IE_NATUREPIECE').AsString else

  if (StName='DEBIT') then
    BEGIN
    QI.Fields[i].AsFloat:=QImpEcr.FindField('IE_DEBIT').AsFloat ;
    if (QImpEcr.FindField('IE_QUALIFPIECE').AsString='N') then
      BEGIN
      DebitMvt:=QI.Fields[i].AsFloat ;
      END ;
    if (QImpEcr.FindField('IE_TYPEECR').AsString<>'A') then InfoImp.TotDeb:=InfoImp.TotDeb+QI.Fields[i].AsFloat ;
    END else
  if (StName='CREDIT') then
    BEGIN
    QI.Fields[i].Value:=QImpEcr.FindField('IE_CREDIT').AsFloat ;
    if (QImpEcr.FindField('IE_QUALIFPIECE').AsString='N') then
      BEGIN
      CreditMvt:=QI.Fields[i].AsFloat ;
      END ;
    if (QImpEcr.Findfield('IE_TYPEECR').AsString<>'A') then InfoImp.TotCred:=InfoImp.TotCred+QI.Fields[i].AsFloat ;
    END else
  if StName='TOTALECRITURE' then QI.Fields[i].AsFloat:=TotLigne.TotEcr else
  if StName='TOTALDEVISE' then QI.Fields[i].AsFloat:=TotLigne.TotDev else
  if StName='TOTALEURO' then QI.Fields[i].AsFloat:=TotLigne.TotEuro else
  if StName='TOTALQTE1' then QI.Fields[i].AsFloat:=TotLigne.TotQte1 else
  if StName='TOTALQTE2' then QI.Fields[i].AsFloat:=TotLigne.TotQte2 else
  if StName='QUALIFECRQTE1' then QI.Fields[i].AsString:=TotLigne.QualQte1 else
  if StName='QUALIFECRQTE2' then QI.Fields[i].AsString:=TotLigne.QualQte2 else
  if (StName='AUXILIAIRE') then
    BEGIN
    QI.Fields[i].Value:=QImpEcr.FindField('IE_'+StName).AsString ; CpteAuxi:=QI.Fields[i].AsString ;
    END else
  if (StName='GENERAL') then
    BEGIN
    QI.Fields[i].Value:=QImpEcr.FindField('IE_'+StName).AsString ; CpteGene:=QI.Fields[i].AsString ;
    END else
  if (StName='SECTION') then
    BEGIN
    QI.Fields[i].AsString:=QImpEcr.FindField('IE_'+StName).AsString ;
    SectMvt:=QI.Fields[i].AsString ; Axe:=QImpEcr.FindField('IE_AXE').AsString ;
    if (Lequel='FBA') then
      BEGIN
      if (QImpEcr.FindField('IE_SECTION').AsString='')
        and (QImpEcr.FindField('IE_AXE').AsString<>'') then
        BEGIN
        k:=Round(Valeur(Copy(QImpEcr.FindField('IE_AXE').AsString,2,1))) ; QI.Fields[i].AsString:=SectAttente[k] ;
        END ;
      END ;
    END else
  if (StName='REGIMETVA') or (StName='TYPEMVT') then
    //PrendsLe:=False
    else
    BEGIN
    if (QImpEcr.FindField('IE_'+StName)<>Nil) then QI.Fields[i].Value:=QImpEcr.FindField('IE_'+StName).AsVariant ;
    END ;
  END ;

// Mise en mémoire du compte gene /auxi en liste ainsi que ses regimetva,nature,totaux ...
if (Lequel='FEC') or (Lequel='FOD') then
  BEGIN
  if (CpteAuxi='') then CpteEnCours:=CpteGene else CpteEnCours:=CpteAuxi ;
  RecupInfosCpte(CpteGene,CpteAuxi) ;
  if (Lequel='FEC') and (QImpEcr.FindField('IE_ANA').AsString='-') then
    BEGIN
    QI.FindField(Pref+'REGIMETVA').Value:=RegimeTva ;
    QI.FindField(Pref+'TYPEMVT').Value:=QuelTypeMvt(CpteEnCours,NatureCpte,QImpEcr.FindField('IE_NATUREPIECE').AsString) ;
    END ;
  END ;

QI.Post ;

if (Lequel='FEC') or (Lequel='FOD') then
  BEGIN
  if  (QImpEcr.FindField('IE_ANA').AsString='X') then
    With TotLigne do
      BEGIN
      TotEcr:=Abs(QImpEcr.FindField('IE_DEBIT').AsFloat-QImpEcr.FindField('IE_CREDIT').AsFloat) ;
      TotDev:=Abs(QImpEcr.FindField('IE_DEBITDEV').AsFloat-QImpEcr.FindField('IE_CREDITDEV').AsFloat) ;
      TotEuro:=Abs(QImpEcr.FindField('IE_DEBITEURO').AsFloat-QImpEcr.FindField('IE_CREDITEURO').AsFloat) ;
      TotQte1:=QImpEcr.FindField('IE_QTE1').AsFloat ;
      TotQte1:=QImpEcr.FindField('IE_QTE2').AsFloat ;
      QualQte1:=QImpEcr.FindField('IE_QUALIFQTE1').AsString ;
      QualQte2:=QImpEcr.FindField('IE_QUALIFQTE2').AsString ;
      END ;
  END ;
if (Lequel='FEC') or (Lequel='FOD') then
  BEGIN
  OkANouv:=(NatJal2(QImpEcr.FindField('IE_JOURNAL').AsString,TabJalSoc)='ANO') ;
  if (QImpEcr.FindField('IE_TYPEECR').AsString='A') then
     BEGIN
     MajTotSection(SectMvt,Axe,OkANouv)
     END ;
  if ((Lequel='FEC') and (QImpEcr.FindField('IE_TYPEECR').AsString<>'A')) or (Lequel='FOD') then
     BEGIN
     MajTotJal ;
     MajTotCompte(CpteGene,OkANouv) ;
     if CpteAuxi<>'' then MajTotCompte(CpteAuxi,OkANouv) ;
     END ;
  END ;
END ;

//Ecritures de format paramétré (Budget,analytiques...)

procedure IntegreParam(Var TabJalSoc : TTabCptLu ; Var InfoImp : TInfoImport) ;
BEGIN
BEGINTRANS ;
Fillchar(InfoImp,SizeOf(InfoImp),#0) ;
While not QImpEcr.EOF do
  BEGIN
  if EnRupture(QImpEcr) then
    BEGIN
    Inc(InfoImp.NbPiece);
    NumL:=0 ; NumLigneGeneA0:=0 ; NumLigneEcheA0:=0 ; NumLigneVentilA0:=0 ;
    NatureEtNumeroJournal(OldJal) ;
    END ;
  if Erreur then Break ;
  if AnnuleImport then
    if FAssImp<>nil then if (THMsgBox(FAssImp.FindComponent('HM')).Execute(39,'','')=mryes) then Break else AnnuleImport:=False ;
  AjouteMvtDanslaBase(QE,TabJalSoc,InfoImp) ;
  Inc(NbLigEnCours) ; Inc(InfoImp.NbLigIntegre) ;
  ActiveTotaux(True,FAssImp,InfoImp) ;
  QImpEcr.Next ; MoveCur(False) ;
  END ;
if AnnuleImport or Erreur then ROLLBACK else COMMITTRANS ;
END ;

// Ecritures générales
procedure IntegrePieces(Var TabJalSoc : TTabCptLu ; Var InfoImp : TInfoImport) ;
BEGIN
if EnRupture(QImpEcr) then
  BEGIN
  Inc(InfoImp.NbPiece);
  NumL:=0 ; NumLigneGeneA0:=0 ; NumLigneEcheA0:=0 ; NumLigneVentilA0:=0 ;
  NatureEtNumeroJournal(OldJal) ;
  END ;
// Ne prends pas les mouvements à zéro et les ligne d'éché / ventil. correspondantes
if not QImpEcr.Eof then
   BEGIN
   // RAZ n° lignes d'échéances.
   if (NumLigneEcheA0<0) and (QImpEcr.FindField('IE_ECHE').AsString='-') then NumLigneEcheA0:=0 ;
   // RAZ pour n° lignes ventilations.
   if (NumLigneVentilA0<0) and (QImpEcr.FindField('IE_NUMVENTIL').AsInteger=0) then NumLigneVentilA0:=0 ;
   // Ligne géné., éché. ou ventil. à 0.
   if (Arrondi(QImpEcr.FindField('IE_DEBIT').AsFloat,V_PGI.OkDecV)=0) and (Arrondi(QImpEcr.FindField('IE_CREDIT').AsFloat,V_PGI.OkDecV)=0) then
     BEGIN
     if not MvtAZero then
        BEGIN
        // MvtAZero : Pour ne pas prendre en compte toutes les ventilations à zéro
        if (QImpEcr.FindField('IE_ANA').AsString='X') and (QImpEcr.FindField('IE_TYPEECR').AsString='E') then MvtAZero:=True ;
        // décalage des n° lignes ventilations suivantes.
        if (QImpEcr.FindField('IE_NUMVENTIL').AsInteger>=1) then Dec(NumLigneVentilA0) ;
        if (QImpEcr.FindField('IE_TYPEECR').AsString<>'A') then
           BEGIN
           // décalage des n° lignes non échéances suivantes.
           if (QImpEcr.FindField('IE_NUMLIGNE').AsInteger>1)
              and (QImpEcr.FindField('IE_NUMECHE').AsInteger<=1) then Dec(NumLigneGeneA0) ;
           // décalage n° lignes d'échéances suivantes.
           if (QImpEcr.FindField('IE_NUMECHE').AsInteger>=1) then Dec(NumLigneEcheA0) ;
           END ;
        END ;
     Exit ;
     END else
     if MvtAZero then
        BEGIN
        If (QImpEcr.FindField('IE_TYPEECR').AsString='A') then Exit
                                                          else MvtAZero:=False ;
        END ;
   END ;
if (QImpEcr.FindField('IE_TYPEECR').AsString<>'A') then AjouteMvtDansLaBase(QE,TabJalSoc,InfoImp)
                                                   else AjouteMvtDansLaBase(QY,TabJalSoc,InfoImp) ;
END ;

procedure IntegreGenerales(TabJalSoc : TTabCptLu ; Var InfoImp : TInfoImport) ;
BEGIN
if (Format<>'RAP') and (Format<>'CRA') and (Format<>'MP') and (Prefixe='E') then
  BEGIN
  QY:=TQuery.Create(Application) ; QY.DataBaseName:='SOC' ; QY.RequestLive:=True ;
  QY.SQL.Clear ; QY.SQL.Add('SELECT * FROM ANALYTIQ WHERE Y_EXERCICE="'+W_W+'"') ;
  ChangeSQL(QY) ; QY.Prepare ; QY.Open ;
  END ;
Fillchar(InfoImp,SizeOf(InfoImp),#0) ;
BEGINTRANS ;
While not QImpEcr.EOF do
  BEGIN
  if Erreur then Break ;
  if AnnuleImport then
    if FAssImp<>nil then if (THMsgBox(FAssImp.FindComponent('HM')).Execute(39,'','')=mryes) then Break else AnnuleImport:=False ;
  if (Format='RAP') or (Format='CRA') then RapprocheLigne(InfoImp) else IntegrePieces(TabJalSoc,InfoImp) ;
  Inc(NbLigEnCours) ; Inc(InfoImp.NbLigIntegre) ;
  ActiveTotaux(True,FAssImp,InfoImp) ;
  QImpEcr.Next ; MoveCur(False) ;
  END ;
if AnnuleImport or Erreur then ROLLBACK else COMMITTRANS ;
QE.Close ; QE.Free ;
if (Format<>'RAP') and (Format<>'CRA') and (Format<>'MP') then BEGIN QY.Close ; QY.Free ; END ;
END ;

procedure IntegreEcr(FAss : TForm ; Fmt,Leq : String ; Var InfoImp : TInfoImport) ;
Var Nb : Integer ;
     Q : TQuery ;
     TabJalSoc : TTabCptLu ;
BEGIN
FAssImp:=FAss ; Format:=Fmt ; Lequel:=Leq ; FillChar(InfoImp,SizeOf(InfoImp),#0) ;
Prefixe:='E' ; Table:='ECRITURE' ;
if Lequel='FBE' then BEGIN Prefixe:='BE' ; Table:='BUDECR' ; END else
  if Lequel='FOD' then BEGIN Prefixe:='Y' ; Table:='ANALYTIQ' ;END ;
AnnuleImport:=False ; Erreur:=False ;
QImpEcr:=OpenSQL('SELECT * FROM IMPECR WHERE IE_SELECTED="X" AND IE_OKCONTROLE="X" AND IE_INTEGRE="-" AND IE_TYPEECR<>"L" ORDER BY IE_CHRONO',False) ;
QImpEcr.UpdateMode:=upWhereChanged ;
Q:=OpenSQL('SELECT COUNT(IE_CHRONO) FROM IMPECR WHERE IE_SELECTED="X" AND IE_OKCONTROLE="X" AND IE_INTEGRE="-" AND IE_TYPEECR<>"L"',True) ;
Nb:=Q.Fields[0].AsVariant ;
Ferme(Q) ;
if (Nb<=0) and (FAssImp<>nil) then
  BEGIN
  AnnuleImport:=True ;
  Ferme(QImpEcr) ;
  if (Format='RAP') or (Format='CRA') then THMsgBox(FAssImp.FindComponent('HM')).Execute(50,'','')
                                      else THMsgBox(FAssImp.FindComponent('HM')).Execute(42,'','') ;
  Exit ;
  END ;
InitMove(Nb+1,'') ; MoveCur(False) ;
NbLigEnCours:=0 ; InfoImp.NbLigIntegre:=0 ; InfoImp.TotDeb:=0 ; InfoImp.TotCred:=0 ; InfoImp.NbPiece:=0 ;
NumP:=0 ; OldNumP:=0 ;
OldJal:='' ; OldNat:='' ; OldDateC:=iDate1900 ;
OldQualP:='' ; JalChange:='' ; QualifPiece:='' ;
CpteEnCours:='' ; OldCpteEnCours:='' ;
QE:=TQuery.Create(Application) ; QE.DataBaseName:='SOC' ; QE.RequestLive:=True ;
QE.SQL.Clear ; QE.SQL.Add('SELECT * FROM '+Table+' WHERE '+Prefixe+'_EXERCICE="'+W_W+'"') ;
ChangeSQL(QE) ; QE.Prepare ; QE.Open ;
// Listes des fiches Cpte,Tiers,Jal et section
RemplirLesListes ;
ChargejalSoc(TabJalSoc) ;
if (Lequel<>'FEC') or ((Lequel='FEC') and (Format='CRA')) then IntegreParam(TabJalSoc,InfoImp) else IntegreGenerales(TabJalSoc,InfoImp) ;
FiniMove ;
Ferme(QImpEcr) ;
if (Lequel<>'FBE') then
  BEGIN
  if not AnnuleImport then UpdateTotTables('') ;
  TTotalCpte.Clear ; TTotalCpte.Free ;
  QGen.Close ; QGen.Free ; QTiers.Close ; QTiers.Free ;
  END ;
TListeJal.Clear ; TListeJal.Free ;
VideListe(TListeSection) ; TListeSection.Free ;
if not AnnuleImport then ExecuteSQL('UPDATE IMPECR SET IE_INTEGRE="X" WHERE IE_SELECTED="X" AND IE_OKCONTROLE="X" AND IE_TYPEECR<>"L"') ;
if (FAss=nil) then CompteRenduBatch(1,InfoImp) ;
AnnuleImport:=False ;
END ;

{======================= lettrage après intégration ==========================}

procedure Lettrage ;
var L : TL_Rappro ;
    SQL,OldLettrage,OldLett : String ;
    QImp : TQuery ;
    TLett:TList ;
    ExistMvt : TExistMvt ;
    OkALettrer : Boolean ;
    NbLettImp,NbLett,DecDev : integer ;
    Quotite : Double ;
    Auxi,Gene : String17 ;
    OldDev : String3 ;
    QEcrLett : TQuery ;
    StV8 : String ;
BEGIN
//ActiveAttente(6,True) ;
ExecuteSQL('UPDATE IMPECR SET IE_OKCONTROLE="X",IE_SELECTED="X" WHERE IE_TYPEECR="L"') ;
SQL:='SELECT IE_GENERAL,IE_AUXILIAIRE,IE_DATECOMPTABLE,IE_DEVISE,IE_LETTRAGE,IE_JOURNAL,IE_DEBIT,IE_CREDIT,IE_REFINTERNE FROM IMPECR WHERE IE_TYPEECR="L" AND IE_OKCONTROLE="X" AND IE_SELECTED="X"'
     +' ORDER BY IE_AUXILIAIRE,IE_GENERAL,IE_DEVISE,IE_ETATLETTRAGE, IE_LETTRAGE, IE_DATECOMPTABLE, IE_NUMPIECE, IE_NUMLIGNE, IE_NUMECHE ' ;
QImp:=OpenSQL(SQL,True) ;
if (QImp.EOF) then BEGIN Ferme(QImp) ; Exit ; END ;
InitMove(RecordsCount(QImp),'') ;
QEcrLett:=TQuery.Create(Application) ; QEcrLett.DataBaseName:='SOC' ;
QEcrLett.SQL.Add('SELECT E_JOURNAL,E_EXERCICE,E_DATECOMPTABLE,E_NUMEROPIECE,'
                +'E_NATUREPIECE,E_NUMLIGNE,E_NUMECHE, E_GENERAL,E_AUXILIAIRE,E_DATEREFEXTERNE,'
                +'E_REFINTERNE,E_REFLIBRE,E_REFEXTERNE,E_LIBELLE,E_TAUXDEV,E_DEBIT,E_CREDIT,'
                +'E_DEBITDEV,E_CREDITDEV,E_LETTRAGE,E_DATEECHEANCE,E_DEVISE,'
                +'E_DEBITEURO,E_CREDITEURO FROM ECRITURE'
                +'WHERE E_JOURNAL=:JOURNAL AND E_DATECOMPTABLE=:DATECOMPTABLE'
                +'AND E_GENERAL=:GENERAL AND E_AUXILIAIRE=:AUXIANA AND E_DEBIT=:DEBIT'
                +'AND E_CREDIT=:CREDIT AND E_REFINTERNE=:REFINTERNE'
                +'AND E_DEVISE=:DEVISE') ;
StV8:=LWhereV8 ; if StV8<>'' then QEcrLett.SQL.Add(' AND '+StV8) ;
if not QEcrLett.Prepared then BEGIN ChangeSQL(QEcrLett) ; QEcrLett.Prepare ; END ;
TLett:=TList.Create ;
While not QImp.Eof do
  BEGIN
  VideListe(TLett) ; OldLettrage:='' ; NbLett:=0 ; NbLettImp:=0 ;
  OldLett:=QImp.Fields[4].AsString ;
  Auxi:=QImp.Fields[1].AsString ;
  Gene:=QImp.Fields[0].AsString ;
  OldDev:=QImp.Fields[3].AsString ;
  OkALettrer:=True ;
  While not (QImp.Eof) and ((QImp.Fields[1].AsString=Auxi)
    and (QImp.Fields[0].AsString=Gene)
    and (OldLett=QImp.Fields[4].AsString)
    and (OldDev=QImp.Fields[3].AsString)) do
    BEGIN
    Inc(NbLettImp) ;
    InitExistMvt(ExistMvt) ;
    With ExistMvt do
      BEGIN
      JOURNAL:=QImp.Fields[5].AsString ;
      DATECOMPTABLE:=QImp.Fields[2].AsDateTime ;
      GENERAL:=QImp.Fields[0].AsString ;
      AUXIANA:=QImp.Fields[1].AsString ;
      DEVISE:=QImp.Fields[3].AsString ;
      DEBIT:=QImp.Fields[6].AsFloat ; CREDIT:=QImp.Fields[7].AsFloat ;
      REFINTERNE:=QImp.Fields[8].AsString ;
      QEcrLett.ParamByName('AUXIANA').AsString:=AUXIANA ;
      QEcrLett.ParamByName('CREDIT').AsFloat:=CREDIT ;
      QEcrLett.ParamByName('DATECOMPTABLE').AsDateTime:=DATECOMPTABLE ;
      QEcrLett.ParamByName('DEBIT').AsFloat:=DEBIT ;
      QEcrLett.ParamByName('DEVISE').AsString:=DEVISE ;
      QEcrLett.ParamByName('GENERAL').AsString:=GENERAL ;
      QEcrLett.ParamByName('JOURNAL').AsString:=JOURNAL ;
      QEcrLett.ParamByName('REFINTERNE').AsString:=REFINTERNE ;
      END ;
    QEcrLett.Open ;
    if QEcrLett.Eof then OkALettrer:=False ;
    While OkALettrer and (not QEcrLett.Eof) do
      BEGIN
      if (OldLettrage='') then OldLettrage:=QEcrLett.Fields[19].AsString ;
      OkALettrer:=((OldLettrage='') or (QEcrLett.Fields[19].AsString=OldLettrage)) ;
      if OkALettrer then
        BEGIN
        Inc(NbLett) ;
        RecupDevise(QImp.Fields[3].AsString,DecDev,Quotite,TDev) ;
        L:=TL_Rappro.Create ;
        L.General:=QEcrLett.Fields[7].AsString ; L.Auxiliaire:=QEcrLett.Fields[8].AsString ;
        L.DateC:=QEcrLett.Fields[2].AsDateTime ; L.DateE:=QEcrLett.Fields[20].AsDateTime ; L.DateR:=QEcrLett.Fields[9].AsDateTime ;
        L.RefI:=QEcrLett.Fields[10].AsString ; L.RefL:=QEcrLett.Fields[11].AsString ;
        L.RefE:=QEcrLett.Fields[12].AsString ; L.Lib:=QEcrLett.Fields[13].AsString ;
        L.Jal:=QEcrLett.Fields[0].AsString ; L.Numero:=QEcrLett.Fields[3].AsInteger;
        L.NumLigne:=QEcrLett.Fields[5].AsInteger ; L.NumEche:=QEcrLett.Fields[6].AsInteger ;
        L.CodeL:=QEcrLett.Fields[19].AsString ;
        L.TauxDEV:=QEcrLett.Fields[14].AsFloat ;
        L.CodeD:=QEcrLett.Fields[21].AsString ;
        L.Decim:=DecDev ;
        L.Debit:=QEcrLett.Fields[15].AsFloat ; L.Credit:=QEcrLett.Fields[16].AsFloat ;
        L.DebDev:=QEcrLett.Fields[17].AsFloat ; L.CredDev:=QEcrLett.Fields[18].AsFloat ;
        L.DebEuro:=QEcrLett.Fields[22].AsFloat ; L.CredEuro:=QEcrLett.Fields[23].AsFloat ;
        L.Nature:=QEcrLett.Fields[4].AsString ;
        L.Facture:=((L.Nature='FC') or (L.Nature='FF') or (L.Nature='AC') or (L.Nature='AF')) ;
        L.Client:=((L.Nature='FC') or (L.Nature='AC') or (L.Nature='RC') or (L.Nature='OC')) ;
        L.Solution:=0 ; L.Exo:=QEcrLett.Fields[1].AsString ;
        L.EditeEtatTva:=False ; 
        TLett.Add(L) ;
        QEcrLett.Next ;
        END ;
      END ;
    QEcrLett.Close ;
    QImp.Next ; MoveCur(False) ;
    END ;
  OkALettrer:=(OkALettrer and (NbLettImp=NbLett)) ;
  if OkALettrer then LettrerUnPaquet(TLett,True,False) ;
  END ;
VideListe(TLett) ; TLett.Free ; Ferme(QImp) ; QEcrLett.Free ; FiniMove ;
//ActiveAttente(6,False) ;
END ;

Procedure LanceVerifCpta(EnBatch : boolean) ;
Var ListeErreurs : TStringList ;
    TypeV : TTypeVerif ;
BEGIN
//MajSelection;
ListeErreurs:=TStringList.Create ;
TypeV:=tvEcr ;
if Prefixe='BE' then TypeV:=tvEcrBud else
  if Prefixe='Y' then TypeV:=tvEcrAna ;
if not VerPourImp(ListeErreurs,EnBatch,TypeV) then MajImpErr(ListeErreurs,True) ;
END ;

Procedure LanceLettrage ;
var QL : TQuery ;
BEGIN
QL:=OpenSQL('SELECT DISTINCT IE_TYPEECR FROM IMPECR WHERE IE_TYPEECR="L"',True) ;
if not QL.Eof then
   BEGIN
   Application.ProcessMessages ;
   //if Transactions(LanceLettrage,5)<>OeOK then BEGIN MessageAlerte('Enregistrement inaccessible') ; Exit ; END ;
   Lettrage ;
   END ;
Ferme(QL) ;
Application.ProcessMessages ;
END ;

function MajLettrageEtVentileGen(Compte : String17) : boolean ;
var QI : TQuery ;
    TGene : TGGeneral ;
    NumAxe : integer ;
    FAnal : TAnal ;
    StAxe : String ;
BEGIN
Result:=False ;
TGene:=TGGeneral.Create(Compte) ;
if TGene=nil then BEGIN TGene.free ; Exit ; END ;
if TGene.Lettrable then
   BEGIN
   QI:=OpenSQL('SELECT * FROM IMPECR WHERE IE_GENERAL="'+Compte+'" AND IE_ETATLETTRAGE="RI"',False) ;
   QI.UpdateMode:=upWhereChanged ;
   While not QI.Eof do
     BEGIN
     QI.Edit ;
     QI.FindField('IE_ECHE').AsString:='X' ;
     QI.FindField('IE_NUMECHE').AsInteger:=1 ;
     QI.FindField('IE_ETATLETTRAGE').AsString:='AL' ;
     QI.FindField('IE_LETTRAGE').AsString:='' ;
     QI.FindField('IE_DATEECHEANCE').AsDateTime:=QI.FindField('IE_DATECOMPTABLE').AsDateTime ;
     QI.FindField('IE_DATEPAQUETMIN').AsDateTime:=QI.FindField('IE_DATECOMPTABLE').AsDateTime ;
     QI.FindField('IE_DATEPAQUETMAX').AsDateTime:=QI.FindField('IE_DATECOMPTABLE').AsDateTime ; ;
     QI.FindField('IE_MODEPAIE').AsString:=FirstModepaie(TGene.ModeRegle) ;
     QI.Post ;
     QI.Next ;
     if not Result then Result:=True ;
     END ;
   Ferme(QI) ;
   END ;
if TGene.Ventilable[1] or TGene.Ventilable[2] or TGene.Ventilable[3]
   or TGene.Ventilable[4] or TGene.Ventilable[5] then
  BEGIN
  QI:=OpenSQL('SELECT * FROM IMPECR WHERE IE_GENERAL="'+Compte+'" AND IE_ANA="-"',False) ;
  QI.UpdateMode:=upWhereChanged ;
  While not QI.Eof do
    BEGIN
    QI.Edit ;
    QI.FindField('IE_ANA').AsString:='X' ;
    NumAxe:=1 ;
    if TGene.Ventilable[1] then NumAxe:=1 else
     if TGene.Ventilable[2] then NumAxe:=2 else
      if TGene.Ventilable[3] then NumAxe:=3 else
       if TGene.Ventilable[4] then NumAxe:=4 else
        if TGene.Ventilable[5] then NumAxe:=5 ;
    QI.FindField('IE_AXE').AsString:='A'+IntToStr(NumAxe) ;
    QI.FindField('IE_SECTION').AsString:=SectAttente[NumAxe] ;
    With FAnal,MvtImport^ do
      BEGIN
      Journal:=QI.FindField('IE_JOURNAL').AsString ;
      NatPiece:=QI.FindField('IE_NATUREPIECE').AsString ;
      QualP:=QI.FindField('IE_QAUALIFPIECE').AsString ;
      General:=QI.FindField('IE_GENERAL').AsString ;
      NumP:=QI.FindField('IE_NUMPIECE').AsInteger ;
      NumL:=QI.FindField('IE_NUMLIGNE').AsInteger ;
      NumV:=1 ;
      RefInt:=QI.FindField('IE_REFINTERNE').AsString ;
      Lib:=QI.FindField('IE_LIBELLE').AsString ;
      Debit:=QI.FindField('IE_DEBIT').AsFLoat ;
      Credit:=QI.FindField('IE_CREDIT').AsFLoat ;
      END ;
    QI.Post ;
    MvtImport^:=InitMvtImp^ ;
    MvtImport^.IE_ETABLISSEMENT:=VH^.EtablisDefaut ;
    MvtImport^.IE_SOCIETE:=V_PGI.CodeSociete ;
    With MvtImport^,FAnal do
      BEGIN
      IE_JOURNAL:=Journal ;
      IE_NATUREPIECE:=NatPiece ;
      IE_QUALIFPIECE:=QualP ;
      IE_GENERAL:=General ;
      IE_NUMPIECE:=NumP ;
      IE_NUMLIGNE:=NumL ;
      IE_NUMVENTIL:=NumV ;
      IE_REFINTERNE:=RefInt ;
      IE_LIBELLE:=Lib ;
      IE_DEBIT:=Debit ;
      IE_CREDIT:=Credit ;
      AjouteDevise(V_PGI.DevisePivot) ;
      END ;
    MemoriseMontantGene ;
    StAxe:='-----' ; StAxe[NumAxe]:='X' ;
//    AjouteAnalytique(StAxe) ;
//    AjouteMvt(QAjoute) ;
    QI.Next ;
    if not Result then Result:=True ;
    END ;
  Ferme(QI) ;
  END ;
TGene.Free ;
END ;

function MajEtatLettrageTiers(Compte : String17) : boolean ;
var QI : TQuery ;
    TTiers : TGTiers ;
BEGIN
Result:=False ;
TTiers:=TGTiers.Create(Compte) ;
if TTiers=nil then BEGIN TTiers.Free ; Exit ; END ;
if not TTiers.Lettrable then BEGIN TTiers.Free ; Exit ; END ;
QI:=OpenSQL('SELECT * FROM IMPECR WHERE IE_AUXILIAIRE="'+Compte+'" AND IE_ETATLETTRAGE="RI"',False) ;
QI.UpdateMode:=upWhereChanged ;
While not QI.Eof do
  BEGIN
  QI.Edit ;
  QI.FindField('IE_ECHE').AsString:='X' ;
  QI.FindField('IE_NUMECHE').AsInteger:=1 ;
  QI.FindField('IE_ETATLETTRAGE').AsString:='AL' ;
  QI.FindField('IE_LETTRAGE').AsString:='' ;
  QI.FindField('IE_DATEECHEANCE').AsDateTime:=QI.FindField('IE_DATECOMPTABLE').AsDateTime ;
  QI.FindField('IE_DATEPAQUETMIN').AsDateTime:=QI.FindField('IE_DATECOMPTABLE').AsDateTime ;
  QI.FindField('IE_DATEPAQUETMAX').AsDateTime:=QI.FindField('IE_DATECOMPTABLE').AsDateTime ; ;
  QI.FindField('IE_MODEPAIE').AsString:=FirstModepaie(TTiers.ModeRegle) ;
  QI.Post ;
  QI.Next ;
  if not Result then Result:=True ;
  END ;
Ferme(QI) ;
TTiers.Free ;
END ;

function MajIEGeneral(Compte : String17 ; Lettrable,Ventilable : boolean) : boolean ;
var TGene : TGGeneral ;
BEGIN
Result:=False ;
TGene:=TGGeneral.Create(Compte) ;
if Lettrable then ExecuteSQL('UPDATE IMPECR SET IE_ECHE="X",IE_NUMECHE=1,IE_ETATLETTRAGE="AL",'+
                             'IE_LETTRAGE="",IE_DATEECHEANCE=IE_DATECOMPTABLE,IE_DATEPAQUETMIN=IE_DATECOMPTABLE,'+
                             'IE_DATEPAQUETMAX=IE_DATECOMPTABLE,IE_MODEPAIE="'+FirstModePaie(TGene.ModeRegle)+'"'+
                             ' WHERE IE_GENERAL="'+Compte+'" AND IE_ECHE="-"') ;
if Ventilable then ExecuteSQL('UPDATE IMPECR SET IE_ANA="X" WHERE IE_GENERAL="'+Compte+'" AND IE_ANA="-"') ;
TGene.free ;
END ;

function MajIETiers(Compte : String17 ; Lettrable : boolean) : boolean ;
var TTiers : TGTiers ;
BEGIN
Result:=False ;
TTiers:=TGTiers.Create(Compte) ;
if Lettrable then ExecuteSQL('UPDATE IMPECR SET IE_ECHE="X",IE_NUMECHE=1,IE_ETATLETTRAGE="AL",'+
                             'IE_LETTRAGE="",IE_DATEECHEANCE=IE_DATECOMPTABLE,IE_DATEPAQUETMIN=IE_DATECOMPTABLE,'+
                             'IE_DATEPAQUETMAX=IE_DATECOMPTABLE,IE_MODEPAIE="'+FirstModePaie(TTiers.ModeRegle)+'"'+
                             ' WHERE IE_AUXILIAIRE="'+Compte+'" AND IE_ECHE="-"') ;
TTiers.Free ;
END ;

// Infos sur les comptes : V_PGI.SAV seulement

procedure ImporteInfosCptes(var ListeComptes : Array of TStringList)  ;
var P,i : Integer ;
    OldSt,St,S,S1 : String ;
    General,Auxi,Section : String17 ;
    FInfos : TFInfosCompte ;
    DecalDate : byte ;
BEGIN
AssignFile(Fichier,FileName) ; {$I-} Reset(Fichier) ; {$I+}
if (IOResult<>0) then Exit ;
DecalDate:=0 ;
if (Format='HAL') or (Format='HLI') then DecalDate:=2 ;
Readln(Fichier,St) ; Readln(Fichier,St) ; OldSt:='' ;
While Not EOF(Fichier) do
  BEGIN
  P:=Pos(SepLigneIE,St) ; S1:='' ; if (OldSt='') then OldSt:=Copy(St,1,47) ; //Ident. des comptes et typemvt ;
  if P=0 then BEGIN Readln(Fichier,St) ; Continue ; END ;

  Section:='' ; Auxi:='' ;
  General:=Trim(Copy(OldSt,12+DecalDate,17)) ;
  Inc(DecalDate,4) ;
  case OldSt[25+DecalDate] of
    'X','E' : Auxi:=Trim(Copy(OldSt,26+DecalDate,17)) ;
    'A'     : Section:=Trim(Copy(OldSt,26+DecalDate,17)) ;
    END ;
  OldSt:='' ;
  S:=Copy(St,P+1,Length(St)) ;
  P:=Pos(SepLigneIE,S) ;
  if P>0 then
    BEGIN
    if (Auxi<>'') or (Section<>'') then
      BEGIN
      S1:=Copy(S,P+1,Length(St)) ;
      S:=Copy(S,1,P-1) ;
      END ;
    i:=ListeComptes[0].Indexof(General) ;
    if i>-1 then
       BEGIN
       FInfos:=TFInfosCompte(ListeComptes[0].Objects[i]) ;
       FInfos.Lib:=Copy(S,1,35) ;
       FInfos.St:=S ;
       ListeComptes[0].Objects[i]:=FInfos ;
       END ;
    i:=ListeComptes[1].IndexOf(Section) ;
    if i>-1 then
       BEGIN
       FInfos:=TFInfosCompte(ListeComptes[1].Objects[i]) ;
       FInfos.Lib:=Copy(S1,1,35) ;
       FInfos.Axe:=Copy(S1,42,2) ;
       FInfos.St:=S1 ;
       ListeComptes[1].Objects[i]:=FInfos ;
       END ;
    i:=ListeComptes[2].IndexOf(Auxi) ;
    if i>-1 then
       BEGIN
       FInfos:=TFInfosCompte(ListeComptes[2].Objects[i]) ;
       FInfos.Lib:=Copy(S1,1,35) ;
       FInfos.St:=S1 ;
       ListeComptes[2].Objects[i]:=FInfos ;
       END ;
    END ;
  Readln(Fichier,St) ;
  END ;
CloseFile(Fichier) ;
END ;

// Liste des comptes à créer

procedure RempliListesNewCptes(Leq,Fmt,FileN,Prefixe : String ; var ListeComptes : Array of TStringList ; Msg : THMsgBox ; LAide : THLabel ) ;
var QC : TQuery ;
    Ecart,i : integer ;
    FInfos : TFInfosCompte ;
    SQLGen,SQLSec,SQLAux,
    StGenExist,StSecExist,StAuxExist : String ;
BEGIN
Lequel:=Leq ; Format:=Fmt ; FileName:=FileN ;
QC:=TQuery.Create(Application) ;
QC.DatabaseName:=DBSOC.DataBaseName ;
QC.SessionName:=DBSOC.SessionName ;
QC.RequestLive:=False ; QC.UniDirectional:=True ;
InitMove(7,'') ;
MoveCur(False) ;
// Quels comptes pour import des écritures budgetaires / courantes
Ecart:=2 ;
if Prefixe='E' then
  BEGIN
  SQLGen:='SELECT DISTINCT IE_GENERAL,G_LIBELLE FROM IMPECR LEFT JOIN GENERAUX ON IE_GENERAL=G_GENERAL' ;
  SQLSec:='SELECT DISTINCT IE_SECTION,S_LIBELLE,IE_AXE FROM IMPECR LEFT JOIN SECTION ON IE_SECTION=S_SECTION AND IE_AXE=S_AXE' ;
  SQLAux:='SELECT DISTINCT IE_AUXILIAIRE,T_LIBELLE FROM IMPECR LEFT JOIN TIERS ON IE_AUXILIAIRE=T_AUXILIAIRE' ;
  StGenExist:='SELECT G_GENERAL FROM GENERAUX WHERE G_GENERAL=IMPECR.IE_GENERAL' ;
  StAuxExist:='SELECT T_AUXILIAIRE FROM TIERS WHERE T_AUXILIAIRE=IMPECR.IE_AUXILIAIRE' ;
  StSecExist:='SELECT S_SECTION FROM SECTION WHERE S_SECTION=IMPECR.IE_SECTION AND S_AXE=IMPECR.IE_AXE' ;
  END else
if Prefixe='BE' then
  BEGIN
  Ecart:=5 ;
  SQLGen:='SELECT DISTINCT IE_GENERAL,BG_LIBELLE FROM IMPECR LEFT JOIN BUDGENE ON IE_GENERAL=BG_BUDGENE' ;
  SQLSec:='SELECT DISTINCT IE_SECTION,BS_LIBELLE,IE_AXE FROM IMPECR LEFT JOIN BUDSECT ON IE_SECTION=BS_BUDSECT AND IE_AXE=BS_AXE' ;
  SQLAux:='SELECT DISTINCT IE_AUXILIAIRE,T_LIBELLE FROM IMPECR LEFT JOIN TIERS ON IE_AUXILIAIRE=T_AUXILIAIRE' ;
  StGenExist:='SELECT BG_BUDGENE FROM BUDGENE WHERE BG_BUDGENE=IMPECR.IE_GENERAL' ;
  StAuxExist:='SELECT T_AUXILIAIRE FROM TIERS WHERE T_AUXILIAIRE=IMPECR.IE_AUXILIAIRE' ;
  StSecExist:='SELECT BS_BUDSECT FROM BUDSECT WHERE BS_BUDSECT=IMPECR.IE_SECTION AND BS_AXE=IMPECR.IE_AXE' ;
  END ;

// Comptes à Créer
for i:=Low(ListeComptes) to High(ListeComptes) do
  BEGIN
  QC.SQL.Clear ;
  case i of
    0:QC.SQL.Add('SELECT DISTINCT IE_GENERAL FROM IMPECR WHERE NOT EXISTS('+StGenExist+')') ;
    1:QC.SQL.Add('SELECT DISTINCT IE_SECTION,IE_AXE FROM IMPECR WHERE IE_TYPEECR="A" AND NOT EXISTS ('+StSecExist+')') ;
    2:BEGIN
      if Prefixe='BE' then BEGIN MoveCur(False) ; Continue ; END ;
      QC.SQL.Add('SELECT DISTINCT IE_AUXILIAIRE FROM IMPECR WHERE IE_AUXILIAIRE<>"" AND NOT EXISTS('+StAuxExist+')') ;
      END ;
    else Continue ;
    END ;
  LAide.Caption:=Msg.Mess[i+Ecart]+' '+Msg.Mess[7] ;
  If Imp_Methode1 Then Application.ProcessMessages ;
  ChangeSQL(QC) ;
  QC.Open ;
  While not QC.Eof do
    BEGIN
    FInfos:=TFInfosCompte.Create ;
    FInfos.New:=True ; FInfos.Lib:=QC.Fields[0].AsString ;
    if i=1 then BEGIN FInfos.Axe:=QC.Fields[1].AsString ; END ;
    ListeComptes[i].AddObject(QC.Fields[0].AsString,FInfos) ;
    QC.Next ;
    END ;
  QC.CLose ;
  MoveCur(False) ;
  END ;
// Comptes à modifier
for i:=Low(ListeComptes) to High(ListeComptes) do
  BEGIN
  QC.SQL.Clear ;
  case i of
    0:QC.SQL.Add(SqlGen+' WHERE EXISTS('+StGenExist+')') ;
    1:QC.SQL.Add(SqlSec+' WHERE EXISTS('+StSecExist+')') ;
    2:BEGIN
      if Prefixe='BE' then BEGIN MoveCur(False) ; Continue ; END ;
      QC.SQL.Add(SQLAux+' WHERE EXISTS('+StAuxExist+')') ;
      END ;
    else Continue ;
    END ;
  LAide.Caption:=Msg.Mess[i+Ecart]+' '+Msg.Mess[8] ; Application.ProcessMessages ;
  ChangeSQL(QC) ;
  QC.Open ;
  While not QC.Eof do
    BEGIN
    FInfos:=TFInfosCompte.Create ;
    FInfos.New:=False ; FInfos.Lib:=QC.Fields[1].AsString ;
    if i=1 then BEGIN FInfos.Axe:=QC.Fields[2].AsString ; END ;
    ListeComptes[i].AddObject(QC.Fields[0].AsString,FInfos) ;
    QC.Next ;
    END ;
  QC.CLose ;
  MoveCur(False) ;
  END ;
FiniMove ;
QC.Free ;
if V_PGI.SAV and ((Lequel='FEC') and ((Format='HAL')
   or (Format='SAA') or (Format='SN2'))) then ImporteInfosCptes(ListeComptes) ;
END ;

procedure MajIMPECR(ListeComptes : Array of TStringList) ;
var i : integer ;
    FLib : TFInfosCompte ;
    Compte : String17 ;
BEGIN

InitMove (ListeComptes[0].Count+ListeComptes[2].Count,'') ;
// Généraux
for i:=0 to ListeComptes[0].Count-1 do
  BEGIN
  Compte:=ListeComptes[0][i] ;
  FLib:=TFInfosCompte(ListeComptes[0].Objects[i]) ;
  if FLib.New=False then Continue ;
  MajLettrageEtVentileGen(Compte) ;
  If Imp_Methode1 Then Application.ProcessMessages ;
  MoveCur(False) ;
  END ;

// Tiers
for i:=0 to ListeComptes[2].Count-1 do
  BEGIN
  Compte:=ListeComptes[2][i] ;
  FLib:=TFInfosCompte(ListeComptes[2].Objects[i]) ;
  if FLib.New=False then Continue ;
  MajEtatLettrageTiers(Compte) ;
  If Imp_Methode1 Then Application.ProcessMessages ;
  MoveCur(False) ;
  END ;
FiniMove ;
END ;

end.

