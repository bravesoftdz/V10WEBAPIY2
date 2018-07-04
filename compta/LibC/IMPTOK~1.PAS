unit ImpToKill;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, Hctrls, Buttons, ExtCtrls, Ent1, HEnt1, Menus, DB, DBTables, ed_tools,
  ComCtrls, FmtChoix,CPTESAV,MajTable, HStatus, HSysMenu, hmsgbox, IniFiles,ImpFic,
  ImpFicU,RappType,TImpFic ;

procedure VideStringList ( L : TStringList ) ;
procedure VideListeInfoImp(Var InfoImp : TInfoImport ; OnFree : Boolean) ;
function ImporteLesEcritures(Msg : THMsgBox ; Var InfoImp : TInfoImport) : boolean ;
procedure MajImpErr(ListeEntetePieceFausse : TStringList) ;
procedure IntegreEcr(FAss : TForm ; Var InfoImp : TInfoImport) ;
procedure MajOkControle ;
procedure ChargeDevEtSect ;
Function  AlignDroite(St : String ; l : Integer) : String ;
Procedure RempliListeInfoImp(Var InfoImp : TInfoImport) ;
Procedure FaitFichierRejet(InfoImp : TInfoImport) ;
Procedure FaitFichierDoublon(InfoImp : TInfoImport) ;
Procedure ChargeScenarioImport(Var InfoImp : TInfoImport) ;

//procedure AjouteReference(Gene : String17 ; StRef : String ; DD : TDateTime) ;

// Objets d'importation

type
  FMvtImport = ^TFMvtImport ;
  TFMvtImport = RECORD
//String
  IE_AFFAIRE,IE_ETATLETTRAGE,IE_LETTRAGE,IE_LETTREPOINTLCR,IE_LIBELLE,IE_REFEXTERNE,
  IE_REFINTERNE,IE_REFLIBRE,IE_REFPOINTAGE,IE_REFRELEVE,IE_RIB,IE_SECTION,
  IE_NUMPIECEINTERNE,IE_LIBRETEXTE0,IE_LIBRETEXTE1,IE_LIBRETEXTE2,IE_LIBRETEXTE3,
  IE_LIBRETEXTE4,IE_LIBRETEXTE5,IE_LIBRETEXTE6,IE_LIBRETEXTE7,IE_LIBRETEXTE8,
  IE_LIBRETEXTE9,IE_NUMEROIMMO : String ;
  IE_AXE : String[2] ;
// Combos
  IE_JOURNAL,IE_ENCAISSEMENT,IE_ECRANOUVEAU,IE_ETABLISSEMENT,IE_FLAGECR,IE_MODEPAIE,IE_NATUREPIECE,
  IE_REGIMETVA,IE_QUALIFPIECE,IE_QUALIFQTE1,IE_QUALIFQTE2,IE_SOCIETE,IE_TPF,IE_TVA,
  IE_TYPEANOUVEAU,IE_TYPEECR,IE_TYPEMVT,IE_DEVISE,IE_TABLE0,IE_TABLE1,
  IE_TABLE2,IE_TABLE3,IE_CONSO : String3 ;
// String17
  IE_AUXILIAIRE,IE_GENERAL,IE_CONTREPARTIEAUX,IE_CONTREPARTIEGEN : String17 ;
// entiers
  IE_CHRONO : Integer ;
  IE_NUMECHE,IE_NUMLIGNE,IE_NUMPIECE,
  IE_NUMVENTIL : Integer ;
// Dates
  IE_DATECOMPTABLE,IE_DATEECHEANCE,IE_DATEPAQUETMAX,IE_DATEPAQUETMIN,
  IE_DATEPOINTAGE,IE_DATEREFEXTERNE,IE_DATERELANCE,IE_DATETAUXDEV,
  IE_DATEVALEUR,IE_ORIGINEPAIEMENT,IE_LIBREDATE,IE_DATECREATION : TDateTime ;
// booleens
  IE_ECHE,IE_CONTROLE,IE_LETTRAGEDEV,IE_OKCONTROLE,
  IE_SELECTED,IE_TVAENCAISSEMENT,IE_TYPEANALYTIQUE,IE_VALIDE,IE_ANA,IE_INTEGRE,
  IE_LIBREBOOL0,IE_LIBREBOOL1,IE_LETTRAGEEURO,IE_SAISIEEURO  : String1 ;
// doubles
  IE_POURCENTAGE,IE_POURCENTQTE1,IE_POURCENTQTE2,IE_QTE1,IE_QTE2,
  IE_QUOTITE : Double ;
  IE_RELIQUATTVAENC,IE_TAUXDEV,IE_TOTALTVAENC,IE_COTATION : Double ;
  IE_DEBIT,IE_CREDIT,IE_CREDITDEV,IE_CREDITEURO,IE_COUVERTURE,
  IE_COUVERTUREDEV,IE_DEBITDEV,IE_DEBITEURO,IE_COUVERTUREEURO,
  IE_LIBREMONTANT0,IE_LIBREMONTANT1,IE_LIBREMONTANT2,IE_LIBREMONTANT3 : Double ;

  IE_ELEMENTARECUPERER : Boolean ;
  END ;

// Objets d'intégration

type TFDoublon = Class
     IMPORT : boolean ;
     JOURNAL : String3 ;
     DATEC : TDateTime ;
     NATUREPIECE : String3 ;
     NUMPIECE : Integer ;
     QUALIFPIECE : String3 ;
     END ;

type TReference = Class
     Ref : String ;
     Date : TDateTime ;
     END ;

var
    AnnuleImport : boolean ;

Const Imp_Methode1 : Boolean = TRUE ;
Const QAJParam : BOOLEAN = FALSE ;
implementation

uses SaisUtil,HCompte,VERCPTA,LettUtil,LetBatch,RapSuppr,FmtImpor,ImporFmt,VisuEnr ;

Const MaxCptesEnMemoire = 5000 ;

var

    // Variables pour l'importation
    MsgBox            : THMsgBox ;
    MvtAZero,OkRupt   : boolean ;
    NbCptes,NbLig : LongInt ;
    TDev      : TList ;
    MvtImport : FMvtImport ;
    Erreur : boolean ;
    FAssImp : TForm ;
    ListePointe    : TStringList ;
    TLett           : TList ;

procedure VideStringList ( L : TStringList ) ;
Var i : integer ;
BEGIN
if L=Nil then Exit ; if L.Count<=0 then Exit ;
for i:=0 to L.Count-1 do If L.Objects[i]<>NIL Then L.Objects[i].Free ;
L.Clear ;
END ;

procedure VideListeInfoImp(Var InfoImp : TInfoImport ; OnFree : Boolean) ;
BEGIN
VideStringList(InfoImp.LGenLu) ; If OnFree then InfoImp.LGenLu.Free ;
VideStringList(InfoImp.LAuxLu) ; If OnFree then InfoImp.LAuxLu.Free ;
VideStringList(InfoImp.LAnaLu) ; If OnFree then InfoImp.LAnaLu.Free ;
VideStringList(InfoImp.LJalLu) ; If OnFree then InfoImp.LJalLu.Free ;
VideStringList(InfoImp.LMP) ; If OnFree then InfoImp.LMP.Free ;
VideListe(InfoImp.ListeCptFaux) ; If OnFree then InfoImp.ListeCptFaux.Free ;
VideListe(InfoImp.ListePieceFausse) ; If OnFree then InfoImp.ListePieceFausse.Free ;
VideListe(InfoImp.CRListeEnteteDoublon) ; If OnFree then InfoImp.CRListeEnteteDoublon.Free ;
VideStringList(InfoImp.ListeEnteteDoublon) ; If OnFree then InfoImp.ListeEnteteDoublon.Free ;
If OnFree Then InfoImp.ListeEntetePieceFausse.Free ;
END ;


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

Procedure ChargeScenarioImport(Var InfoImp : TInfoImport) ;
BEGIN
ChargeScenario('X',InfoImp.Format,InfoImp.Lequel,InfoImp.SC) ;
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

function GoodSoc(Var InfoImp : TInfoImport) : byte ;
var St, Lib : String ;
    Fichier   : TextFile ;
BEGIN
if not FileExists(InfoImp.NomFic) then BEGIN Result:=2 ; Exit ; END ;
AssignFile(Fichier,InfoImp.Nomfic) ;
{$I-} Reset (Fichier) ; {$I+}
if EOF(Fichier) then BEGIN Result:=2 ; CloseFile(Fichier); Exit ; END ;
Readln(Fichier,St) ;
Lib:=Trim(Copy(St,1,70)) ;
NbLig:=0 ;
While not EOF(Fichier) do
    BEGIN
    if (InfoImp.Format='EDI') and (InfoImp.Lequel='FBA')  then
      if (Copy(St,1,5)='01480') then NbCptes:=Round(Valeur(Trim(Copy(St,9,15)))) ;
    Inc(NbLig) ;
    Readln(Fichier,St) ;
    END ;
CloseFile(Fichier) ;
//if (FFormat.Value='EDI') then BEGIN Result:=1 ; Exit ; END ;
//Q:=OpenSQL('Select SO_LIBELLE From SOCIETE',True) ;
//If (Trim(Q.Fields[0].AsString)<>Lib) then Exit ;
Result:=1 ;
END ;

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
If (J=0) Or (M=0) Or (Y=0) Then Exit ;
If j in [1..31]=FALSE Then Exit ;
If m in [1..12]=FALSE Then Exit ;
Result:=EncodeDate(Y,M,J) ;
END ;

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

Function AlignDroite(St : String ; l : Integer) : String ;
var St1 : String ;
BEGIN
St1:=Trim(st) ;
While Length(St1)<l Do St1:=' '+St1 ;
AlignDroite:=St1 ;
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

Procedure TrouveLeModePaie(StMp : String ; Var InfoImp : TInfoImport ; Var Mvt : TFMvtImport) ;
Var CptLuMP : TCptLu ;
    St : String ;
BEGIN
St:=CHERCHEUNMODE(StMP,False,False) ;
Fillchar(CptLuMP,SizeOf(CptLuMP),#0) ;
CptLuMp.Nature:=St ;
If LeModePaie(InfoImp.LMP,CptLuMp) Then Mvt.IE_MODEPAIE:=Trim(CptLuMP.Cpt) Else
  BEGIN
  Fillchar(CptLuMP,SizeOf(CptLuMP),#0) ;
  CptLuMp.Cpt:=St ;
  If LeModePaie(InfoImp.LMP,CptLuMp) Then Mvt.IE_MODEPAIE:=Trim(CptLuMP.Cpt) ;
  END ;
END ;

function EnRuptAscii(AlimOldIdentPiece : Boolean ; Var IdentPiece,OldIdentPiece : TIdentPiece) : Boolean ;
BEGIN
Result:=False ;
If (IdentPiece.JalP<>OldIdentPiece.JalP) or (IdentPiece.DateP<>OldIdentPiece.DateP) or (IdentPiece.NumP<>OldIdentPiece.NumP)
   or (IdentPiece.QualP<>OldIdentPiece.QualP) or (IdentPiece.NatP<>OldIdentPiece.NatP) then
    BEGIN
    If AlimOldIdentPiece Then
      BEGIN
      OldIdentPiece.JalP:=IdentPiece.JalP ; OldIdentPiece.DateP:=IdentPiece.DateP ;
      OldIdentPiece.NumP:=IdentPiece.NumP ; OldIdentPiece.QualP:=IdentPiece.QualP ;
      OldIdentPiece.NatP:=IdentPiece.NatP ; MvtAZero:=False ;
      END ;
    Result:=True ;
    END ;
END;

function DetecteUneRupture(Var St : String ; Var IdentPiece : TIdentPiece ; Var InfoImp : TInfoImport) : Boolean ;
Var FmtFic : Integer ;
    JalP,NatP : String ;
    DateP : TDateTime ;
    NumP : Integer ;
    Decal : Integer ;
BEGIN
Result:=False ; FmtFic:=0 ; Decal:=0 ;
If Not VH^.ImportRL Then Decal:=8 ;
if InfoImp.Format='SAA' then FmtFic:=0 else if InfoImp.Format='SN2' then FmtFic:=0 else
if InfoImp.Format='HLI' then FmtFic:=1 else if InfoImp.Format='HAL' then FmtFic:=2 else
if InfoImp.Format='CGN' then FmtFic:=3 else if InfoImp.Format='CGE' then FmtFic:=4 Else Exit ;
JalP:='' ; NumP:=0 ; DateP:=0 ;
Case FmtFic Of
  0 : BEGIN
      JalP:=Trim(Copy(St,1,3)) ;
      DateP:=Format_Date(Copy(St,4,6)) ;
      NatP:=Trim(Copy(St,10,2)) ;
      NumP:=Round(Valeur(Copy(St,106,7))) ;
      END ;
  1,2 : BEGIN
        JalP:=Trim(Copy(St,1,3)) ;
        DateP:=Format_Date_HAL(Copy(St,4,8)) ;
        NatP:=Trim(Copy(St,12,2)) ;
        NumP:=Round(Valeur(Copy(St,110+Decal,7))) ;
        END ;
  3,4 : BEGIN
        JalP:=Trim(Copy(St,1,3)) ; DateP:=Format_Date_HAL(Copy(St,4,8)) ;
        NatP:=Copy(St,12,2)  ;
        NumP:=Round(Valeur(Copy(St,152,8))) ;
        END ;
  END ;
If (IdentPiece.JalP<>JalP) or (IdentPiece.DateP<>DateP) or (IdentPiece.NumP<>NumP) or (IdentPiece.NatP<>NatP) then Result:=True ;
END;



function TestBreak : Boolean ;
BEGIN
Result:=False ;
If Imp_Methode1 Then Application.ProcessMessages ;
if AnnuleImport and (MsgBox<>nil) then
  BEGIN
  if (MsgBox.Execute(39,'','')<>mryes) then AnnuleImport:=False ;
  Result:=AnnuleImport ;
  END ;
END ;

procedure ChargeDevEtSect ;
var Q,Q1       : TQuery ;
    TDevise : TFDevise ;
BEGIN
(*
i:=0 ;
Q:=OpenSQL('SELECT X_SECTIONATTENTE FROM AXE ORDER BY X_AXE',True) ;
While not Q.Eof do BEGIN inc(i) ; SectAttente[i]:=Q.Fields[0].AsString ; Q.Next ; END ;
Ferme(Q) ;
*)
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


// RECORD des mouvements à importer ;
procedure InitMvtImport(Var MvtImport : TFMvtImport) ;
BEGIN
Fillchar(MvtImport,SizeOf(MvtImport),#0) ;
With MvtImport do
  BEGIN
  IE_ETATLETTRAGE:='RI'; IE_ECRANOUVEAU:='N';IE_ETABLISSEMENT:=VH^.EtablisDefaut; IE_QUALIFPIECE:='N';
  IE_SOCIETE:=V_PGI.CodeSociete; IE_DEVISE:=V_PGI.DevisePivot;
  IE_NUMPIECE:=1 ; IE_DATEECHEANCE:=iDate1900 ;IE_DATEPAQUETMAX:=iDate1900 ;
  IE_DATEPAQUETMIN:=iDate1900 ; IE_DATEPOINTAGE:=iDate1900 ; IE_DATEREFEXTERNE:=iDate1900 ;
  IE_DATERELANCE:=iDate1900 ; IE_DATETAUXDEV:=iDate1900 ; IE_DATEVALEUR:=iDate1900 ;
  IE_DATECREATION:=Date ;
  IE_ORIGINEPAIEMENT:=iDate1900 ; IE_DATECOMPTABLE:=iDate1900 ; IE_LIBREDATE:=iDate1900 ;
  IE_ECHE:='-' ;IE_ENCAISSEMENT:='RIE';IE_CONTROLE:='-' ;IE_LETTRAGEDEV:='-' ;
  IE_OKCONTROLE:='X';IE_SELECTED:='X' ; IE_TVAENCAISSEMENT:='-' ; IE_TYPEANALYTIQUE:='-' ;
  IE_VALIDE:='-' ; IE_ANA:='-' ;IE_INTEGRE:='-' ; IE_QUOTITE:=1 ;
  IE_LIBREBOOL0:='-' ; IE_LIBREBOOL1:='-' ; IE_SAISIEEURO:='-' ;
  IE_ELEMENTARECUPERER:=FALSE ; IE_COTATION:=1 ;
  END ;
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
QAjoute.SQL.Add('IE_COUVERTURE,IE_COUVERTUREDEV,IE_DEBITDEV,IE_DEBITEURO,') ;

QAjoute.SQL.Add('IE_LIBRETEXTE0,IE_LIBRETEXTE1,IE_LIBRETEXTE2,IE_LIBRETEXTE3,IE_LIBRETEXTE4,') ;
QAjoute.SQL.Add('IE_LIBRETEXTE5,IE_LIBRETEXTE6,IE_LIBRETEXTE7,IE_LIBRETEXTE8,IE_LIBRETEXTE9,') ;
QAjoute.SQL.Add('IE_TABLE0,IE_TABLE1,IE_TABLE2,IE_TABLE3,IE_LIBREBOOL0,IE_LIBREBOOL1,') ;

QAjoute.SQL.Add('IE_LIBREMONTANT0,IE_LIBREMONTANT1,IE_LIBREMONTANT2,IE_LIBREMONTANT3,IE_LIBREDATE,IE_SAISIEEURO,IE_COTATION ') ;
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
  if IE_TYPEECR<>'L' then Inc(InfoImp.NbLigIntegre) ;
//  Inc(NbLigEnCours) ;
  if ((IE_TYPEECR<>'A') and (IE_TYPEECR<>'L')) Or ((IE_TYPEECR='A') AND (IE_TYPEANALYTIQUE='X')) then
  //Cumuls pour Compte-rendu
    if ((MsgBox=nil) or (ChoixFmt.CompteRendu) or (InfoImp.Lequel='FBA')) then
      BEGIN
      InfoImp.TotDeb:=InfoImp.TotDeb+IE_DEBIT ;
      InfoImp.TotCred:=InfoImp.TotCred+IE_CREDIT ;
      END ;
  if InfoImp.ForcePositif then
     BEGIN
     if (IE_DEBIT<0) and (IE_CREDIT=0) then
        BEGIN
        IE_CREDIT:=Abs(IE_DEBIT) ; IE_DEBIT:=0 ;
        IE_CREDITDEV:=Abs(IE_DEBITDEV) ; IE_DEBITDEV:=0 ;
        IE_CREDITEURO:=Abs(IE_DEBITEURO) ; IE_DEBITEURO:=0 ;
        END else
     if (IE_CREDIT<0) and (IE_DEBIT=0) then
        BEGIN
        IE_DEBIT:=Abs(IE_CREDIT) ; IE_CREDIT:=0 ;
        IE_DEBITDEV:=Abs(IE_CREDITDEV) ; IE_CREDITDEV:=0 ;
        IE_DEBITEURO:=Abs(IE_CREDITEURO) ; IE_CREDITEURO:=0 ;
        END ;
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
     StrfPoint(IE_DEBITDEV)+','+StrfPoint(IE_DEBITEURO)+',') ;
     (*
     QAjoute.SQL.Add('"","","","","","","","","","","","","","","-","-",0,0,0,0,"'+USDateTime(IDate1900)+'")') ;
     *)
     QAjoute.SQL.Add('"'+TrimSpe(IE_LIBRETEXTE0)+'","'+TrimSpe(IE_LIBRETEXTE1)+'","'+TrimSpe(IE_LIBRETEXTE2)+'","'+
                     TrimSpe(IE_LIBRETEXTE3)+'","'+TrimSpe(IE_LIBRETEXTE4)+'",') ;
     QAjoute.SQL.Add('"'+TrimSpe(IE_LIBRETEXTE5)+'","'+TrimSpe(IE_LIBRETEXTE6)+'","'+TrimSpe(IE_LIBRETEXTE7)+'","'+
                     TrimSpe(IE_LIBRETEXTE8)+'","'+TrimSpe(IE_LIBRETEXTE9)+'",') ;
     QAjoute.SQL.Add('"'+TrimSpe(IE_TABLE0)+'","'+TrimSpe(IE_TABLE1)+'","'+TrimSpe(IE_TABLE2)+'","'+TrimSpe(IE_TABLE3)+'","'+
                     TrimSpe(IE_LIBREBOOL0)+'","'+TrimSpe(IE_LIBREBOOL1)+'",') ;

     QAjoute.SQL.Add(StrfPoint(IE_LIBREMONTANT0)+','+StrfPoint(IE_LIBREMONTANT1)+','+StrfPoint(IE_LIBREMONTANT2)+','+
                     StrfPoint(IE_LIBREMONTANT3)+',"'+UsDateTime(IE_LIBREDATE)+'","'+IE_SAISIEEURO+'",'+StrfPoint(IE_COTATION)+')') ;
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
        QAjoute.Params[78].AsString:=IE_LIBRETEXTE0 ;
        QAjoute.Params[79].AsString:=IE_LIBRETEXTE1 ;
        QAjoute.Params[80].AsString:=IE_LIBRETEXTE2 ;
        QAjoute.Params[81].AsString:=IE_LIBRETEXTE3 ;
        QAjoute.Params[82].AsString:=IE_LIBRETEXTE4 ;
        QAjoute.Params[83].AsString:=IE_LIBRETEXTE5 ;
        QAjoute.Params[84].AsString:=IE_LIBRETEXTE6 ;
        QAjoute.Params[85].AsString:=IE_LIBRETEXTE7 ;
        QAjoute.Params[86].AsString:=IE_LIBRETEXTE8 ;
        QAjoute.Params[87].AsString:=IE_LIBRETEXTE9 ;
        QAjoute.Params[88].AsString:=IE_TABLE0 ;
        QAjoute.Params[89].AsString:=IE_TABLE1 ;
        QAjoute.Params[90].AsString:=IE_TABLE2 ;
        QAjoute.Params[91].AsString:=IE_TABLE3 ;
        QAjoute.Params[92].AsString:=IE_LIBREBOOL0 ;
        QAjoute.Params[93].AsString:=IE_LIBREBOOL1 ;
        QAjoute.Params[94].AsFloat:=IE_LIBREMONTANT0 ;
        QAjoute.Params[95].AsFloat:=IE_LIBREMONTANT1 ;
        QAjoute.Params[96].AsFloat:=IE_LIBREMONTANT2 ;
        QAjoute.Params[97].AsFloat:=IE_LIBREMONTANT3 ;
        QAjoute.Params[98].AsDateTime:=IE_LIBREDATE ;
        QAjoute.Params[99].AsString:=IE_SAISIEEURO ;
        QAjoute.Params[100].AsFloat:=IE_COTATION ;
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
        QAjoute.FindField('IE_LIBRETEXTE0').AsString:=TrimSpe(IE_LIBRETEXTE0) ;
        QAjoute.FindField('IE_LIBRETEXTE1').AsString:=TrimSpe(IE_LIBRETEXTE1) ;
        QAjoute.FindField('IE_LIBRETEXTE2').AsString:=TrimSpe(IE_LIBRETEXTE2) ;
        QAjoute.FindField('IE_LIBRETEXTE3').AsString:=TrimSpe(IE_LIBRETEXTE3) ;
        QAjoute.FindField('IE_LIBRETEXTE4').AsString:=TrimSpe(IE_LIBRETEXTE4) ;
        QAjoute.FindField('IE_LIBRETEXTE5').AsString:=TrimSpe(IE_LIBRETEXTE5) ;
        QAjoute.FindField('IE_LIBRETEXTE6').AsString:=TrimSpe(IE_LIBRETEXTE6) ;
        QAjoute.FindField('IE_LIBRETEXTE7').AsString:=TrimSpe(IE_LIBRETEXTE7) ;
        QAjoute.FindField('IE_LIBRETEXTE8').AsString:=TrimSpe(IE_LIBRETEXTE8) ;
        QAjoute.FindField('IE_LIBRETEXTE9').AsString:=TrimSpe(IE_LIBRETEXTE9) ;
        QAjoute.FindField('IE_TABLE0').AsString:=TrimSpe(IE_TABLE0) ;
        QAjoute.FindField('IE_TABLE1').AsString:=TrimSpe(IE_TABLE1) ;
        QAjoute.FindField('IE_TABLE2').AsString:=TrimSpe(IE_TABLE2) ;
        QAjoute.FindField('IE_TABLE3').AsString:=TrimSpe(IE_TABLE3) ;
        QAjoute.FindField('IE_LIBREBOOL0').AsString:=TrimSpe(IE_LIBREBOOL0) ;
        QAjoute.FindField('IE_LIBREBOOL1').AsString:=TrimSpe(IE_LIBREBOOL1) ;
        QAjoute.FindField('IE_LIBREMONTANT0').AsFloat:=IE_LIBREMONTANT0 ;
        QAjoute.FindField('IE_LIBREMONTANT1').AsFloat:=IE_LIBREMONTANT1 ;
        QAjoute.FindField('IE_LIBREMONTANT2').AsFloat:=IE_LIBREMONTANT2 ;
        QAjoute.FindField('IE_LIBREMONTANT3').AsFloat:=IE_LIBREMONTANT3 ;
        QAjoute.FindField('IE_LIBREDATE').AsDateTime:=IE_LIBREDATE ;
        QAjoute.FindField('IE_SAISIEEURO').AsString:=IE_SAISIEEURO ;
        QAjoute.FindField('IE_COTATION').AsFloat:=IE_COTATION ;
        QAjoute.Post ;
        END ;
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
BEGIN
(*
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
*)
END ;

Procedure ImportLigneHAL(St : String ; Var InfoImp : TInfoImport ; Var Mvt : TFMvtImport) ;
Var LgCpte,Decal : Byte ;
    Sens,NatP,TypeMvt : String ;

BEGIN
Decal:=0 ; LgCpte:=13 ; if not VH^.ImportRL then LgCpte:=17 ;
InitMvtImport(Mvt) ;
With Mvt do
    BEGIN
    IE_JOURNAL:=Trim(Copy(St,1,3)) ;
    IE_DATECOMPTABLE:=Format_Date_HAL(Copy(St,4,8)) ;
    NatP:=Trim(Copy(St,12,2)) ;
//    TypeCpte:=Copy(St,13,1) ;
    if ((InfoImp.Lequel='FEC') or ((InfoImp.Lequel='FBA') and (InfoImp.Format='HAL'))) and not VH^.ImportRL then Inc(Decal,4) ;
    TypeMvt:=Copy(St,27+Decal,1) ;
    IE_GENERAL:=Trim(Copy(St,14,LgCpte)) ;
    IE_TYPEECR:='E' ;
    case TypeMvt[1] of
      'X','E': BEGIN IE_AUXILIAIRE:=Trim(Copy(St,28+Decal,LgCpte)) ; IE_TYPEECR:='E' ; END ;
      'L': BEGIN IE_AUXILIAIRE:=Trim(Copy(St,28+Decal,LgCpte)) ; IE_TYPEECR:='L' ; END ;
      'A': BEGIN IE_ANA:='X' ; IE_TYPEECR:='A' ; IE_SECTION:=Trim(Copy(St,28+Decal,LgCpte)) ; END ;
      END ;

    if ((InfoImp.Lequel='FEC') or ((InfoImp.Lequel='FBA') and (InfoImp.Format='HAL'))) and not VH^.ImportRL then Inc(Decal,4) ;
    IE_REFINTERNE:=Trim(Copy(St,41+Decal,13)) ;
    IE_LIBELLE:=Trim(Copy(St,54+Decal,25)) ;
    Sens:=Copy(St,88+Decal,1) ;
    if Sens='D' then IE_DEBIT:=Valeur(StPoint(Copy(St,89+Decal,20))) else
      if Sens='C' then IE_CREDIT:=Valeur(StPoint(Copy(St,89+Decal,20))) ;
    if (InfoImp.Format='HAL') and (InfoImp.Lequel<>'FBA') and (Trim(Copy(St,748+Decal,3))<>'') then NatP:=Trim(Copy(St,748+Decal,3)) ;
    IE_NATUREPIECE:=NatP ;
//    IE_QUALIFPIECE:=Trim(Copy(St,109+Decal,1)) ;
    If Trim(Copy(St,109+Decal,1))<>'' Then IE_QUALIFPIECE:=Trim(Copy(St,109+Decal,1)) ;
    IE_NUMPIECE:=Round(Valeur(Copy(St,110+Decal,7))) ;
    END ;
END ;

Procedure ImportLigneAutre(St : String ; Var InfoImp : TInfoImport ; Var Mvt : TFMvtImport) ;
Var    TypeMvt,Sens  : String ;

BEGIN
TypeMvt:=Copy(St,25,1) ;
// On ne prend que les échéances au format SN2
if (InfoImp.Format='SN2') and (TypeMvt='X') then Exit ;
Sens:=Copy(St,84,1) ;
//Jal:=Copy(St,1,3) ;DateC:=Format_Date(Copy(St,4,6)) ; if (TypeMvt<>'A') then NatP:=Copy(St,10,2) ;
InitMvtImport(Mvt) ;
With Mvt do
  BEGIN
  IE_JOURNAL:=Trim(Copy(St,1,3)) ;
  IE_DATECOMPTABLE:=Format_Date(Copy(St,4,6)) ;
  IE_NATUREPIECE:=Trim(Copy(St,10,2)) ;
//  TypeCpte:=Copy(St,11,1) ;
  IE_GENERAL:=Trim(Copy(St,12,13)) ;
  IE_TYPEECR:='E' ;
  case TypeMvt[1] of
    'X','E' : BEGIN IE_AUXILIAIRE:=Trim(Copy(St,26,13)) ; IE_TYPEECR:='E' ; END ;
    'L': BEGIN IE_AUXILIAIRE:=Trim(Copy(St,26,13)) ; IE_TYPEECR:='L' ;END ;
    'A': BEGIN IE_ANA:='X' ; IE_TYPEECR:='A' ; IE_SECTION:=Trim(Copy(St,26,13)) ; END ;
    END ;
  IE_REFINTERNE:=Trim(Copy(St,39,13)) ;
  IE_LIBELLE:=Trim(Copy(St,52,25)) ;
  if Sens='D' then IE_DEBIT:=Valeur(StPoint(Copy(St,85,20))) else
    if Sens='C' then IE_CREDIT:=Valeur(StPoint(Copy(St,85,20))) ;
//  IE_QUALIFPIECE:=Trim(Copy(St,105,1)) ;
  If Trim(Copy(St,105,1))<>'' Then IE_QUALIFPIECE:=Trim(Copy(St,105,1)) ;
  IE_NUMPIECE:=Round(Valeur(Copy(St,106,7))) ;
  END ;
END ;

Procedure AffecteDebitCredit(Sens : String ; Montant : Double ; Var Deb,Cre : Double) ;
BEGIN
Deb:=0 ; Cre:=0 ; If Montant=0 Then Exit ;
if Sens='D' then Deb:=Montant else if Sens='C' then Cre:=Montant ;
END ;

Procedure AffecteMontant(Qui : Char ; Montant : Double ; Sens : String ; Var Mvt : TFMvtImport) ;
Var EnEuro : Boolean ;
BEGIN
EnEuro:=EuroOK And VH^.TenueEuro ;
Case Qui Of
  'F' : If EnEuro Then AffecteDebitCredit(Sens,Montant,MVT.IE_DEBITEURO,MVT.IE_CREDITEURO)
                  Else AffecteDebitCredit(Sens,Montant,MVT.IE_DEBIT,MVT.IE_CREDIT) ;
  'E' : If EnEuro Then AffecteDebitCredit(Sens,Montant,MVT.IE_DEBIT,MVT.IE_CREDIT)
                  Else AffecteDebitCredit(Sens,Montant,MVT.IE_DEBITEURO,MVT.IE_CREDITEURO) ;
  'D' : AffecteDebitCredit(Sens,Montant,MVT.IE_DEBITDEV,MVT.IE_CREDITDEV) ;
  END ;
END ;

Procedure InterpreteMontant(St : String ; Var Mvt : TFMvtImport ; Var IdentPiece : TIdentPiece) ;
Var CodeMontant,Sens : String ;
    IndP,IndD,IndE : Integer ;
    Montant : Array[1..3] Of Double ;
    SMontant : Array[1..3] Of String ;
BEGIN
//CodeMontant:=Trim(Copy(St,173,3)) ;
If IdentPiece.CodeMontant='' Then IdentPiece.CodeMontant:=Trim(Copy(St,173,3)) ;
CodeMontant:=IdentPiece.CodeMontant ;
If (CodeMontant='') Or (CodeMontant='---') Then CodeMontant:='F--' ;
IndP:=Pos('F',CodeMontant) ; IndD:=Pos('D',CodeMontant) ; IndE:=Pos('E',CodeMontant) ;
Fillchar(Montant,SizeOf(Montant),#0) ;
SMontant[1]:=Trim(StPoint(Copy(St,131,20))) ; SMontant[2]:=Trim(StPoint(Copy(St,176,20))) ;
SMontant[3]:=Trim(StPoint(Copy(St,196,20))) ;
If SMontant[1]<>'' Then Montant[1]:=Valeur(SMontant[1]) ;
If SMontant[2]<>'' Then Montant[2]:=Valeur(SMontant[2]) ;
If SMontant[3]<>'' Then Montant[3]:=Valeur(SMontant[3]) ;
Sens:=Copy(St,130,1) ;
If IndP>0 Then AffecteMontant('F',Montant[IndP],Sens,Mvt) ;
If IndD>0 Then AffecteMontant('D',Montant[IndD],Sens,Mvt) ;
If IndE>0 Then AffecteMontant('E',Montant[IndE],Sens,Mvt) ;
END ;


Procedure ImportLigneCEGID(St : String ; Var InfoImp : TInfoImport ; Var Mvt : TFMvtImport ;
                           Var AnalSurUneLigne : Boolean ; Var IdentPiece : TIdentPiece) ;
Var TypeMvt,Sens : String ;
    ElementARecuperer : Boolean ;
BEGIN
ElementARecuperer:=MVT.IE_ELEMENTARECUPERER ;
//ElementARecuperer:=FALSE ;
TypeMvt:=Copy(St,31,1) ;
If AnalSurUneLigne Then TypeMvt:='A' ;
AnalSurUneLigne:=FALSE ;
// On ne prend que les échéances au format SN2
if (InfoImp.Format='SN2') and (TypeMvt='X') then Exit ;
Sens:=Copy(St,84,1) ;
//Jal:=Copy(St,1,3) ;DateC:=Format_Date(Copy(St,4,6)) ; if (TypeMvt<>'A') then NatP:=Copy(St,10,2) ;
If Not ElementARecuperer Then
  BEGIN
  InitMvtImport(Mvt) ;
  With Mvt do
    BEGIN
    IE_JOURNAL:=Trim(Copy(St,1,3)) ; IE_DATECOMPTABLE:=Format_Date_HAL(Copy(St,4,8)) ;
    IE_NATUREPIECE:=Copy(St,12,2)  ; IE_GENERAL:=Trim(Copy(St,14,17)) ;
    IE_TYPEECR:='E' ;
    case TypeMvt[1] of
      'X','E' : BEGIN IE_AUXILIAIRE:=Trim(Copy(St,32,17)) ; IE_TYPEECR:='E' ; END ;
      'L': BEGIN IE_AUXILIAIRE:=Trim(Copy(St,32,17)) ; IE_TYPEECR:='L' ;END ;
      'A','O' : BEGIN IE_ANA:='X' ; IE_TYPEECR:='A' ; IE_SECTION:=Trim(Copy(St,32,17)) ; END ;
      'H' : BEGIN AnalSurUneLigne:=TRUE ; IE_TYPEECR:='E' ; END ;
      END ;
    IE_REFINTERNE:=Trim(Copy(St,49,35)) ;
    IE_LIBELLE:=Trim(Copy(St,84,35)) ;
  //Date d'échéance traitée après selon la caractéristique du compte
    InterpreteMontant(St,Mvt,IdentPiece) ;
  //  IE_QUALIFPIECE:=Copy(St,151,1) ;
    If Trim(Copy(St,151,1))<>'' Then IE_QUALIFPIECE:=Trim(Copy(St,151,1)) ;
    IE_NUMPIECE:=Round(Valeur(Copy(St,152,8))) ;
    if Trim(Copy(St,160,3))<>'' then IE_DEVISE:=Trim(Copy(St,160,3)) ;
    if Trim(Copy(St,163,10))<>'' then IE_TAUXDEV:=Valeur(StPoint(Copy(St,163,10)));
    if Trim(Copy(St,216,3))<>'' then IE_ETABLISSEMENT:=Trim(Copy(St,216,3)) ;
    if Trim(Copy(St,219,2))<>'' then IE_AXE:=Trim(Copy(St,219,2)) ;
    END ;
  END ;
END ;


Procedure InitRetoucheEnr(Var CptLuG,CptLuT,CptLuJ : TCptLu ;
                          Var Lettrable,Pointable,IsTiersT,IsTiersG : Boolean ;
                          Var Mvt : TFMvtImport ; Var InfoImp : TInfoImport) ;
Var IsHT : Boolean ;
    CodeTva,CodeTPF : String ;
BEGIN
Fillchar(CptLuG,SizeOf(CptLuG),#0) ; Fillchar(CptLuT,SizeOf(CptLuT),#0) ;
Fillchar(CptLuJ,SizeOf(CptLuJ),#0) ;
IsHT:=FALSE ; CodeTva:='' ; CodeTPF:='' ;
With Mvt do
  BEGIN
  CptLuJ.Cpt:=IE_JOURNAL ; ChercheCptLu(InfoImp.LJalLu,CptLuJ) ;
  If CptLuJ.Nature='ODA' Then IE_TYPEANALYTIQUE:='X' ;
  CptLuT.Cpt:=IE_AUXILIAIRE ; ChercheCptLu(InfoImp.LAuxLu,CptLuT) ;
  CptLuG.Cpt:=IE_GENERAL ; ChercheCptLu(InfoImp.LGenLu,CptLuG) ;
  If CptLuG.Pointable Then
    BEGIN
    IE_NUMECHE:=1 ; IE_ETATLETTRAGE:='RI' ;
    END ;
  If CptLuG.Ventilable Then IE_ANA:='X' ;
  If CptLuJ.Nature='ANO' then IE_ECRANOUVEAU:='OAN' else if CptLuJ.Nature='CLO' then IE_ECRANOUVEAU:='C' ;
  if (IE_QUALIFPIECE='N') and (InfoImp.ForceQualif<>'') and (IE_ECRANOUVEAU<>'OAN') then IE_QUALIFPIECE:=InfoImp.ForceQualif ;
  if IE_AUXILIAIRE='' then IE_TYPEMVT:=QuelTypeMvt(IE_GENERAL,CptLuG.Nature,IE_NATUREPIECE)
                      else IE_TYPEMVT:=QuelTypeMvt(IE_AUXILIAIRE,CptLuT.Nature,IE_NATUREPIECE) ;
  // Maj de la longueur des comptes du fichier / celles paramétrées.
  (*
  if (Format<>'HAL') and (Format<>'HLI') and (Pointable='X') then BEGIN IE_ECHE:='X' ; IE_MODEPAIE:=TrouveModePaie('',Valeur(StPoint(Copy(St,85,20))),TRUE) ; END else
   if ((Lequel='FBA') or (Format='HLI')) and (Pointable='X') then BEGIN IE_ECHE:='X' ; IE_MODEPAIE:=TrouveModePaie('',Valeur(StPoint(Copy(St,89+Decal,20))),TRUE) ; END ;
  *)
  IsTiersT:=((IE_AUXILIAIRE<>'') And (CptLuT.Lettrable)) ;
  IsTiersG:=((IE_AUXILIAIRE='') And (CptLuG.Lettrable)) ;
  If IsTiersG Then
     BEGIN
     IE_REGIMETVA:=CptLuG.regimeTva ;
     END Else
  If IsTiersT Then
     BEGIN
     IE_REGIMETVA:=CptLuT.regimeTva ;
     END ;
  If MVT.IE_TYPEANALYTIQUE<>'X' Then
    BEGIN
    If (CptLuG.Nature='CHA') Or (CptLuG.Nature='PRO') Or (CptLuG.Nature='IMO') Then
      BEGIN
      IsHT:=TRUE ; CodeTva:=CptLuG.Tva ; CodeTPF:=CptLuG.TPF ;
      If Mvt.IE_TVA='' Then Mvt.IE_TVA:=CodeTva ;
      If Mvt.IE_TPF='' Then Mvt.IE_TPF:=CodeTPF ;
      END ;
    END ;
  END ;
END ;

procedure RAZInfoLettrage(Let,Point : boolean ; Var Mvt : TFMvtImport) ;
BEGIN
With Mvt do
  BEGIN
  IE_LETTRAGE:='';IE_LETTRAGEDEV:='-';IE_COUVERTURE:=0; IE_COUVERTUREDEV:=0;
  IE_COUVERTUREEURO:=0 ; IE_LETTRAGEEURO:='-' ;
  IE_DATEPAQUETMIN:=IDate1900 ; IE_DATEPAQUETMAX:=IDate1900 ;
  IE_ETATLETTRAGE:='RI' ;
  If Point Or Let Then IE_ENCAISSEMENT:=SensEnc(IE_DEBIT,IE_CREDIT) ;
  if Let then
    BEGIN
    IE_ETATLETTRAGE:='AL' ;
    IE_DATEPAQUETMIN:=IE_DATECOMPTABLE ; IE_DATEPAQUETMAX:=IE_DATECOMPTABLE ;
    END ;
  if Point then
    BEGIN
    IE_ETATLETTRAGE:='RI' ; IE_NUMECHE:=1 ; IE_ECHE:='X' ;
//    If IE_MODEPAIE='' Then IE_MODEPAIE:= :
    IE_DATEPAQUETMIN:=IE_DATECOMPTABLE ; IE_DATEPAQUETMAX:=IE_DATECOMPTABLE ;
    END ;
  END ;
END ;

Procedure SetCotation (Var Mvt : TFMvtImport) ;
Var Cote,Taux : Double ;
    Dev       : String3 ;
BEGIN
Taux:=Mvt.IE_TAUXDEV ;
if MVT.IE_DATECOMPTABLE<V_PGI.DateDebutEuro then Cote:=Taux else
   BEGIN
   Dev:=MVT.IE_DEVISE ;
   if ((Dev=V_PGI.DevisePivot) or (Dev=V_PGI.DeviseFongible)) then Cote:=1.0 else
   if V_PGI.TauxEuro<>0 then Cote:=Taux/V_PGI.TauxEuro else Cote:=1 ;
   END ;
Mvt.IE_COTATION:=Cote ;
END ;

Procedure TraiteMontantDevise(Var Mvt : TFMvtImport ; Var IdentPiece : TIdentPiece) ;
Var i : Integer ;
    OkDev : Boolean ;
    EnEuro : Boolean ;
    PbTaux : Boolean ;
    DebitF,CreditF,DebitD,CreditD,DebitE,CreditE : Double ;
    DebitCote,CreditCote,DebitNonCote,CreditNonCote : Double ;
    LaDate : TDateTime ;
    TauxLu : Double ;
    Decim : Integer ;
    TDevise : TFDevise ;
    QuiNonCote : Char ;
{
Après
Taux=Taux*V_PGI.TauxEuro
GetTaux Correct
DeviseToEuro  OK
DeviseToPivot OK
Avant
taux=Taux
GetTaux Correct
DeviseToEuro  OK
DeviseToPivot OK
}
BEGIN
EnEuro:=EuroOK And VH^.TenueEuro ;
If EnEuro Then
  BEGIN
  DebitF:=Mvt.IE_DEBITEURO ; CreditF:=Mvt.IE_CREDITEURO ;
  DebitE:=Mvt.IE_DEBIT ; CreditE:=Mvt.IE_CREDIT ;
  Decim:=V_PGI.OkDecE ;
  END Else
  BEGIN
  DebitF:=Mvt.IE_DEBIT ; CreditF:=Mvt.IE_CREDIT ;
  DebitE:=Mvt.IE_DEBITEURO ; CreditE:=Mvt.IE_CREDITEURO ;
  Decim:=V_PGI.OkDecV ;
  END ;
If MVT.IE_DATECOMPTABLE<V_PGI.DateDebutEuro Then
  BEGIN
  DebitCote:=Arrondi(DebitF,Decim) ; CreditCote:=Arrondi(CreditF,Decim) ;
  DebitNonCote:=Arrondi(DebitE,Decim) ; CreditNonCote:=Arrondi(CreditE,Decim) ;
  QuiNonCote:='E' ;
  END Else
  BEGIN
  DebitCote:=Arrondi(DebitE,Decim) ; CreditCote:=Arrondi(CreditE,Decim) ;
  DebitNonCote:=Arrondi(DebitF,Decim) ; CreditNonCote:=Arrondi(CreditF,Decim) ;
  QuiNonCote:='F' ;
  END ;

OkDev:=FALSE ;
for i:=0 to TDev.Count-1 do
  BEGIN
  TDevise:=TDev[i] ;
  if (TDevise.Code=Mvt.IE_DEVISE) then
    BEGIN
    Mvt.IE_QUOTITE:=TDevise.Quotite ;
    If Mvt.IE_QUOTITE=0 Then Mvt.IE_QUOTITE:=1 ;
    OkDev:=TRUE ;
    IdentPiece.DecimDev:=TDevise.Decimale ;
    Break ;
    END ;
  END ;
If OkDev Then
  BEGIN
  DebitD:=Arrondi(Mvt.IE_DEBITDEV,IdentPiece.DecimDev) ; CreditD:=Arrondi(Mvt.IE_CREDITDEV,IdentPiece.DecimDev) ;
  If DebitD+CreditD<>0 Then
    BEGIN
    MVT.IE_DEBITDEV:=Arrondi(MVT.IE_DEBITDEV,IdentPiece.DecimDev) ;
    MVT.IE_CREDITDEV:=Arrondi(MVT.IE_CREDITDEV,IdentPiece.DecimDev) ;
    If DebitCote+CreditCote<>0 Then
       BEGIN
//       Mvt.IE_TAUXDEV:=(MntP*IE_QUOTITE)/MntD
       If IdentPiece.TauxDev=0 Then
         BEGIN
         If MVT.IE_DATECOMPTABLE<V_PGI.DateDebutEuro
           Then TauxLu:=((DebitCote+CreditCote)*Mvt.IE_QUOTITE)/(DebitD+CreditD)
           Else TauxLu:=(((DebitCote+CreditCote)*Mvt.IE_QUOTITE)/(DebitD+CreditD))*V_PGI.TauxEuro ;
         (*
         Avant la date Cote/Dev
         Apres  la date Cote/dev*TauxEuro
         *)
         Mvt.IE_TAUXDEV:=TauxLu ; IdentPiece.TauxDev:=TauxLu ;
         END Else Mvt.IE_TAUXDEV:=IdentPiece.TauxDev ;
       If DebitNonCote+CreditNonCote<>0 Then
         BEGIN
         END Else
         BEGIN
         If QuiNonCote='E' Then
           BEGIN
           DebitNonCote:=PivotToEuro(DebitCote) ; CreditNonCote:=PivotToEuro(CreditCote) ;
           END Else
           BEGIN
           DebitNonCote:=EuroToPivot(DebitCote) ; CreditNonCote:=EuroToPivot(CreditCote) ;
           END ;
         If Mvt.IE_DATECOMPTABLE<V_PGI.DateDebutEuro Then
           BEGIN
           If EnEuro Then
             BEGIN
             Mvt.IE_DEBIT:=DebitNonCote ;
             Mvt.IE_CREDIT:=CreditNonCote ;
             END Else
             BEGIN
             Mvt.IE_DEBITEURO:=DebitNonCote ;
             Mvt.IE_CREDITEURO:=CreditNonCote ;
             END ;
           END Else
           BEGIN
           If EnEuro Then
             BEGIN
             Mvt.IE_DEBITEURO:=DebitNonCote ;
             Mvt.IE_CREDITEURO:=CreditNonCote ;
             END Else
             BEGIN
             Mvt.IE_DEBIT:=DebitNonCote ;
             Mvt.IE_CREDIT:=CreditNonCote ;
             END ;
           END ;
         END ;
       END Else
       BEGIN
//           If IE_TAUXDEV=0 Then IE_TAUXDEV:=TDevise.TauxDev ;
       PbTaux:=FALSE ;
       If IdentPiece.TauxDev=0 Then
         BEGIN
         TauxLu:=GetTaux(Mvt.IE_DEVISE,LaDate,Mvt.IE_DATECOMPTABLE) ;
         if ((MVT.IE_DATECOMPTABLE<V_PGI.DateDebutEuro) or (Not EstMonnaieIn(MVT.IE_DEVISE))) then PbTaux:=(TauxLu=1) else
            if EstMonnaieIn(Mvt.IE_DEVISE) then PbTaux:=(TauxLu=V_PGI.TauxEuro) ;
         If Not PbTaux Then
           BEGIN
           Mvt.IE_TAUXDEV:=TauxLu ; IdentPiece.TauxDev:=TauxLu ;
           END Else
           BEGIN
           Mvt.IE_TAUXDEV:=0 ; Mvt.IE_DEBIT:=0 ; Mvt.IE_CREDIT:=0 ;
           END ;
         END Else Mvt.IE_TAUXDEV:=IdentPiece.TauxDev ;
       If Not PbTaux Then
         BEGIN
         If EnEuro Then
           BEGIN
           Mvt.IE_DEBIT:=DeviseToEuro(Mvt.IE_DEBITDEV,IdentPiece.TauxDev,Mvt.IE_QUOTITE) ;
           Mvt.IE_CREDIT:=DeviseToEuro(Mvt.IE_CREDITDEV,IdentPiece.TauxDev,Mvt.IE_QUOTITE) ;
           Mvt.IE_DEBITEURO:=DeviseToPivot(Mvt.IE_DEBITDEV,IdentPiece.TauxDev,Mvt.IE_QUOTITE) ;
           Mvt.IE_CREDITEURO:=DeviseToPivot(Mvt.IE_CREDITDEV,IdentPiece.TauxDev,Mvt.IE_QUOTITE) ;
           END Else
           BEGIN
           Mvt.IE_DEBIT:=DeviseToPivot(Mvt.IE_DEBITDEV,IdentPiece.TauxDev,Mvt.IE_QUOTITE) ;
           Mvt.IE_CREDIT:=DeviseToPivot(Mvt.IE_CREDITDEV,IdentPiece.TauxDev,Mvt.IE_QUOTITE) ;
           Mvt.IE_DEBITEURO:=DeviseToEuro(Mvt.IE_DEBITDEV,IdentPiece.TauxDev,Mvt.IE_QUOTITE) ;
           Mvt.IE_CREDITEURO:=DeviseToEuro(Mvt.IE_CREDITDEV,IdentPiece.TauxDev,Mvt.IE_QUOTITE) ;
           END ;
         END ;
       END ;
    END Else
    BEGIN { Code devise avec montant devise=0 }
    Mvt.IE_DEBIT:=0 ; Mvt.IE_CREDIT:=0 ;
    Mvt.IE_DEBITEURO:=0 ; Mvt.IE_CREDITEURO:=0 ;
    END ;
  END Else
  BEGIN { Code Devise inexistant }
  Mvt.IE_TAUXDEV:=1;Mvt.IE_DEBITDEV:=Mvt.IE_DEBIT ;
  Mvt.IE_CREDITDEV:=Mvt.IE_CREDIT;Mvt.IE_QUOTITE:=1 ;
  END ;
END ;

Procedure TaiteMontantPivotEtOppose(Var Mvt : TFMvtImport ; Var IdentPiece : TIdentPiece) ;
Var     EnEuro : Boolean ;
    DebitF,CreditF,DebitE,CreditE : Double ;
    Decim : Integer ;
BEGIN
IdentPiece.DecimDev:=V_PGI.OkDecV ;
IdentPiece.TauxDev:=1 ; Mvt.IE_TAUXDEV:=IdentPiece.TauxDev ;
EnEuro:=EuroOK And VH^.TenueEuro ;
If EnEuro Then
  BEGIN
  DebitF:=Mvt.IE_DEBITEURO ; CreditF:=Mvt.IE_CREDITEURO ;
  DebitE:=Mvt.IE_DEBIT ; CreditE:=Mvt.IE_CREDIT ;
  Decim:=V_PGI.OkDecE ;
  END Else
  BEGIN
  DebitF:=Mvt.IE_DEBIT ; CreditF:=Mvt.IE_CREDIT ;
  DebitE:=Mvt.IE_DEBITEURO ; CreditE:=Mvt.IE_CREDITEURO ;
  Decim:=V_PGI.OkDecV ;
  END ;
If (Arrondi(Mvt.IE_DEBIT+Mvt.IE_CREDIT,Decim)=0) And (Arrondi(Mvt.IE_DEBITEURO+Mvt.IE_CREDITEURO,Decim)<>0)Then
  BEGIN
  If EnEuro Then
    BEGIN
    Mvt.IE_DEBIT:=PivotToEuro(Mvt.IE_DEBITEURO) ;
    Mvt.IE_CREDIT:=PivotToEuro(Mvt.IE_CREDITEURO) ;
    END Else
    BEGIN
    Mvt.IE_DEBIT:=EuroToPivot(Mvt.IE_DEBITEURO) ;
    Mvt.IE_CREDIT:=EuroToPivot(Mvt.IE_CREDITEURO) ;
    END ;
  END Else
If (Arrondi(Mvt.IE_DEBITEURO+Mvt.IE_CREDITEURO,Decim)=0) And (Arrondi(Mvt.IE_DEBIT+Mvt.IE_CREDIT,Decim)<>0) Then
  BEGIN
  If EnEuro Then
    BEGIN
    Mvt.IE_DEBITEURO:=EuroToPivot(Mvt.IE_DEBIT) ;
    Mvt.IE_CREDITEURO:=EuroToPivot(Mvt.IE_CREDIT) ;
    END Else
    BEGIN
    Mvt.IE_DEBITEURO:=PivotToEuro(MVt.IE_DEBIT) ;
    Mvt.IE_CREDITEURO:=PivotToEuro(MVt.IE_CREDIT) ;
    END ;
  END ;
Mvt.IE_DEBITDEV:=Mvt.IE_DEBIT ; Mvt.IE_CREDITDEV:=Mvt.IE_CREDIT ;
END ;

procedure RetoucheMontantCEGID(Var Mvt : TFMvtImport ; Var IdentPiece : TIdentPiece) ;
BEGIN
If Mvt.IE_DEVISE<>V_PGI.DevisePivot Then TraiteMontantDevise(Mvt,IdentPiece)
                                 Else TaiteMontantPivotEtOppose(Mvt,IdentPiece) ;
SetCotation(Mvt) ;
Mvt.IE_DEBIT:=Arrondi(Mvt.IE_DEBIT,V_PGI.OkDecV) ; Mvt.IE_CREDIT:=Arrondi(Mvt.IE_CREDIT,V_PGI.OkDecV) ;
Mvt.IE_DEBITEURO:=Arrondi(Mvt.IE_DEBITEURO,V_PGI.OkDecV) ; Mvt.IE_CREDITEURO:=Arrondi(Mvt.IE_CREDITEURO,V_PGI.OkDecV) ;
END ;

procedure RetoucheMontantAutre(Var Mvt : TFMvtImport ; Var IdentPiece : TIdentPiece) ;
Var EnEuro : Boolean ;
    SauveD,SauveC : Double ;
BEGIN
EnEuro:=EuroOK And VH^.TenueEuro ;
If EnEuro Then
  BEGIN
  SauveD:=Mvt.IE_DEBITEURO ; SauveC:=Mvt.IE_CREDITEURO ;
  Mvt.IE_DEBITEURO:=Mvt.IE_DEBIT ; Mvt.IE_CREDITEURO:=Mvt.IE_CREDIT ;
  Mvt.IE_DEBIT:=SauveD ; Mvt.IE_CREDIT:=SauveC ;
  END ;
If Mvt.IE_DEVISE<>V_PGI.DevisePivot Then TraiteMontantDevise(Mvt,IdentPiece)
                                 Else TaiteMontantPivotEtOppose(Mvt,IdentPiece) ;
SetCotation(Mvt) ;
Mvt.IE_DEBIT:=Arrondi(Mvt.IE_DEBIT,V_PGI.OkDecV) ; Mvt.IE_CREDIT:=Arrondi(Mvt.IE_CREDIT,V_PGI.OkDecV) ;
Mvt.IE_DEBITEURO:=Arrondi(Mvt.IE_DEBITEURO,V_PGI.OkDecV) ; Mvt.IE_CREDITEURO:=Arrondi(Mvt.IE_CREDITEURO,V_PGI.OkDecV) ;
END ;

(*
procedure RetoucheMontantAutre(Var Mvt : TFMvtImport) ;
var i : integer ;
    TDevise : TFDevise ;
BEGIN
With Mvt do
  BEGIN
  if IE_DEVISE=V_PGI.DevisePivot then
    BEGIN
    IE_DEVISE:=V_PGI.DevisePivot;IE_TAUXDEV:=1;IE_DEBITDEV:=IE_DEBIT ;
    IE_CREDITDEV:=IE_CREDIT;IE_QUOTITE:=1 ;
    IE_DEBITEURO:=PivotToEuro(IE_DEBIT) ;
    IE_CREDITEURO:=PivotToEuro(IE_CREDIT) ;
    END else
    BEGIN
    for i:=0 to TDev.Count-1 do
      BEGIN
      TDevise:=TDev[i] ;
      if (TDevise.Code=IE_DEVISE) then
        BEGIN
//        IE_DEVISE:=Code ;
        IE_QUOTITE:=TDevise.Quotite ;
        IE_TAUXDEV:=TDevise.TauxDev ;
        IE_DEBITDEV:=PIVOTTODEVISE(IE_DEBIT,IE_TAUXDEV,IE_QUOTITE,TDevise.Decimale) ;
        IE_CREDITDEV:=PIVOTTODEVISE(IE_CREDIT,IE_TAUXDEV,IE_QUOTITE,TDevise.Decimale) ;
        IE_DEBITEURO:=PivotToEuro(IE_DEBIT) ;
        IE_CREDITEURO:=PivotToEuro(IE_CREDIT) ;
        Break ;
        END ;
      END ;
    END ;
  END ;
END ;
*)


Procedure retoucheEnrHal(St : String ; Var IdentPiece : TIdentPiece ; Var InfoImp : TInfoImport ;
                         Var Mvt : TFMvtImport ; OkRetoucheMontant : Boolean) ;
Var CptLuG,CptLuT,CptLuJ : TCptLu ;
    Lettrable,Pointable,IsTiersT,IsTiersG : Boolean ;
    Decal : Integer ;
BEGIN
Decal:=0 ; if not VH^.ImportRL  then Decal:=8 ;
InitRetoucheEnr(CptLuG,CptLuT,CptLuJ,Lettrable,Pointable,IsTiersT,IsTiersG,Mvt,InfoImp) ;
With Mvt do
  BEGIN
  If (CptLuG.Lettrable Or CptLuT.Lettrable) And (IE_TYPEECR='E') Then
     BEGIN
     if Trim(Copy(St,80+Decal,8))<>'' then IE_DATEECHEANCE:=Format_Date_HAL(Copy(St,80+Decal,8))
                                       else IE_DATEECHEANCE:=IE_DATECOMPTABLE ;
     IE_ECHE:='X' ;
     IE_DATEPAQUETMIN:=IE_DATECOMPTABLE ; IE_DATEPAQUETMAX:=IE_DATECOMPTABLE ;
     IE_ETATLETTRAGE:='AL' ;
     If InfoImp.Format='HAL' Then IE_MODEPAIE:=Trim(Copy(St,745+Decal,3)) Else
       If InfoImp.Format='HLI' Then TrouveLeModePaie(Copy(St,79+Decal,1),InfoImp,Mvt) ;
     END ;
  If CptLuG.Pointable Then
    BEGIN
    If InfoImp.Format='HAL' Then IE_MODEPAIE:=Trim(Copy(St,745+Decal,3)) Else
      If InfoImp.Format='HLI' Then TrouveLeModePaie(Copy(St,79+Decal,1),InfoImp,Mvt) ;
    if Trim(Copy(St,80+Decal,8))<>'' then IE_DATEECHEANCE:=Format_Date_HAL(Copy(St,80+Decal,8))
                                       else IE_DATEECHEANCE:=IE_DATECOMPTABLE ;
    END ;
  If InfoImp.Format='HAL' Then
     BEGIN
     if Trim(Copy(St,119+Decal,1))<>'' then IE_VALIDE:=Copy(St,119+Decal,1) ;
     if Trim(Copy(St,120+Decal,35))<>'' then IE_REFEXTERNE:=Copy(St,120+Decal,35) ;
     IE_DATEREFEXTERNE:=Format_Date_HAL(Copy(St,155+Decal,8)) ;
     //IE_DATECREATION:=Format_Date_HAL(Copy(St,163+Decal,8)) ;
     //IE_DATEMODIFICATION:=Format_Date_HAL(Copy(St,171+Decal,8)) ;
     IE_SOCIETE:=V_PGI.CodeSociete ; //Copy(St,179+Decal,3) ;
     if Trim(Copy(St,182+Decal,3))<>'' then IE_ETABLISSEMENT:=Trim(Copy(St,182+Decal,3)) ;
     if Trim(Copy(St,185+Decal,17))<>'' then IE_AFFAIRE:=Trim(Copy(St,185+Decal,17));
     IE_DEBITEURO:=Valeur(StPoint(Copy(St,202+Decal,20)));

     IE_CREDITEURO:=Valeur(StPoint(Copy(St,222+Decal,20)));
      if Trim(Copy(St,262+Decal,3))<>'' then IE_DEVISE:=Trim(Copy(St,262+Decal,3));
     IE_DEBITDEV:=Valeur(StPoint(Copy(St,265+Decal,20))); IE_CREDITDEV:=Valeur(StPoint(Copy(St,285+Decal,20)));
     if Trim(Copy(St,305+Decal,20))<>'' then IE_TAUXDEV:=Valeur(StPoint(Copy(St,305+Decal,20)));
     IE_DATETAUXDEV:=Format_Date_HAL(Copy(St,325+Decal,8));
     if Trim(Copy(St,333+Decal,20))<>'' then IE_QUOTITE:=Valeur(StPoint(Copy(St,333+Decal,20)));
     if Trim(Copy(St,353+Decal,3))<>'' then IE_ECRANOUVEAU:=Trim(Copy(St,353+Decal,3));
     IE_QTE1:=Valeur(StPoint(Copy(St,356+Decal,20)));
     IE_QTE2:=Valeur(StPoint(Copy(St,376+Decal,20)));
     IE_QUALIFQTE1:=Trim(Copy(St,396+Decal,3));
     IE_QUALIFQTE2:=Trim(Copy(St,399+Decal,3));
     if IE_TYPEECR<>'A' then
       BEGIN
       IE_REFLIBRE:=Trim(Copy(St,402+Decal,35));
       if Trim(Copy(St,437+Decal,1))<>'' then IE_TVAENCAISSEMENT:=Copy(St,437+Decal,1);
       if Trim(Copy(St,438+Decal,3))<>'' then IE_REGIMETVA:=Trim(Copy(St,438+Decal,3));
       if Trim(Copy(St,441+Decal,3))<>'' then IE_TVA:=Trim(Copy(St,441+Decal,3));
       if Trim(Copy(St,444+Decal,3))<>'' then IE_TPF:=Trim(Copy(St,444+Decal,3));
       IE_CONTREPARTIEGEN:=Trim(Copy(St,447+Decal,17));
       IE_CONTREPARTIEAUX:=Trim(Copy(St,464+Decal,17));

       if (IE_TYPEECR='L') then
         BEGIN
         IE_COUVERTURE:=Valeur(StPoint(Copy(St,481+Decal,20)));
         IE_LETTRAGE:=Copy(St,501+Decal,5);
         IE_LETTRAGEDEV:=Copy(St,506+Decal,1);
         IE_DATEPAQUETMAX:=Format_Date_HAL(Copy(St,507+Decal,8));
         IE_DATEPAQUETMIN:=Format_Date_HAL(Copy(St,515+Decal,8));

         IE_COUVERTUREDEV:=Valeur(StPoint(Copy(St,679+Decal,20)));
         if Trim(Copy(St,447+Decal,17))<>'' then IE_ETATLETTRAGE:=Copy(St,699+Decal,3);
         END ;
       IE_REFPOINTAGE:=Trim(Copy(St,523+Decal,17));
       IE_DATEPOINTAGE:=Format_Date_HAL(Copy(St,540+Decal,8));
       //IE_LETTREPOINTLCR:=Copy(St,548+Decal,4);
       IE_DATERELANCE:=Format_Date_HAL(Copy(St,552+Decal,8));
       if Trim(Copy(St,560,1))<>'' then IE_CONTROLE:=Copy(St,560,1);
       //IE_TOTALTVAENC:=Valeur(StPoint(Copy(St,561+Decal,20)));
       //IE_RELIQUATTVAENC:=Valeur(StPoint(Copy(St,581+Decal,20)));
       IE_DATEVALEUR:=Format_Date_HAL(Copy(St,601+Decal,8));
       IE_RIB:=Trim(Copy(St,609+Decal,35));
       IE_REFRELEVE:=Trim(Copy(St,644+Decal,35));

       IE_NUMPIECEINTERNE:=Trim(Copy(St,702+Decal,35));
       IE_ENCAISSEMENT:=Trim(Copy(St,737+Decal,3));
       IE_TYPEANOUVEAU:=Trim(Copy(St,740+Decal,3));
       if Trim(Copy(St,743+Decal,1))<>'' then IE_ECHE:=Copy(St,743+Decal,1);
       if StrToIntDef(Copy(St,117+Decal,2),0)<>0 then IE_NUMECHE:=StrToIntDef(Copy(St,117+Decal,2),0) ;
       if Trim(Copy(St,744+Decal,1))<>'' then IE_ANA:=Copy(St,744+Decal,1) ;
       if Trim(Copy(St,745+Decal,3))<>'' then IE_MODEPAIE:=Trim(Copy(St,745+Decal,3)) ;
       END else
       BEGIN
       IE_POURCENTAGE:=Valeur(StPoint(Copy(St,402+Decal,20)));
       IE_CONTROLE:=Copy(St,422+Decal,1);
       IE_NUMVENTIL:=StrToIntDef(Copy(St,423+Decal,6),0);IE_POURCENTQTE1:=Valeur(StPoint(Copy(St,429+Decal,20)));IE_POURCENTQTE2:=Valeur(StPoint(Copy(St,449+Decal,20)));
       IE_TYPEANOUVEAU:=Trim(Copy(St,469,3));IE_TYPEMVT:=Trim(Copy(St,472+Decal,3));
       if Trim(Copy(St,475+Decal,1))<>'' then IE_TYPEANALYTIQUE:=Copy(St,475+Decal,1) ;
       END ;
     END ;
  END ;
RazInfoLettrage((IsTiersG or IsTiersT),CptLuG.Pointable,Mvt) ;
If OkRetoucheMontant Then RetoucheMontantAutre(Mvt,IdentPiece) ;
END ;
(*
procedure RetoucheMontantCEGID(Var Mvt : TFMvtImport ; Var IdentPiece : TIdentPiece) ;
Var Dev,St : String ;
    i : integer ;
    TDevise : TFDevise ;
    OkDev,PbTaux : Boolean ;
    MntP,MntD,MntE : Double ;
    Q : TQuery ;
BEGIN
pbTaux:=FALSE ;
With Mvt do
  BEGIN
  MntP:=Arrondi(IE_DEBIT+IE_CREDIT,V_PGI.OkDecV) ;
  MntD:=Arrondi(IE_DEBITDEV+IE_CREDITDEV,6) ;
  MntE:=Arrondi(IE_DEBITEURO+IE_CREDITEURO,V_PGI.OkDecE) ;
  If IE_DEVISE='' Then IE_DEVISE:=V_PGI.DevisePivot ;
  OkDev:=FALSE ;
  If IE_DEVISE<>V_PGI.DevisePivot Then
    BEGIN
    for i:=0 to TDev.Count-1 do
      BEGIN
      TDevise:=TDev[i] ;
      if (TDevise.Code=IE_DEVISE) then
        BEGIN
        IE_QUOTITE:=TDevise.Quotite ;
        If IE_QUOTITE=0 Then IE_QUOTITE:=1 ;
        OkDev:=TRUE ;
        Break ;
        END ;
      END ;
    If OkDev Then
      BEGIN
      If MntD<>0 Then
        BEGIN
        If MntP<>0 Then IE_TAUXDEV:=(MntP*IE_QUOTITE)/MntD Else
           BEGIN
//           If IE_TAUXDEV=0 Then IE_TAUXDEV:=TDevise.TauxDev ;
           If IE_TAUXDEV=0 Then
             BEGIN
             St:='SELECT H_TAUXREEL FROM CHANCELL WHERE H_DEVISE="'+IE_DEVISE+'" '+
                 ' AND H_DATECOURS=( '+
                 'SELECT MAX(H_DATECOURS) FROM CHANCELL WHERE H_DEVISE="'+IE_DEVISE+'" '+
                 'AND H_DATECOURS<"'+UsDateTime(IE_DATECOMPTABLE)+'") ' ;
             Q:=OpenSQL(St,TRUE) ;
             If Not Q.EOF Then IE_TAUXDEV:=Q.Fields[0].AsFloat Else PbTaux:=TRUE ;
             Ferme(Q) ;
             END ;
           If PbTaux Then
             BEGIN
             IE_TAUXDEV:=0 ; IE_DEBIT:=0 ; IE_CREDIT:=0 ;
             END Else
             BEGIN
             If IE_TAUXDEV=0 Then IE_TAUXDEV:=1 ;
             IE_DEBIT:=DEVISETOPIVOT(IE_DEBITDEV,IE_TAUXDEV,IE_QUOTITE) ;
             IE_CREDIT:=DEVISETOPIVOT(IE_CREDITDEV,IE_TAUXDEV,IE_QUOTITE) ;
             END ;
           END ;
        END Else
        BEGIN
        IE_DEBIT:=0 ; IE_CREDIT:=0 ;
        END ;
      END Else
      BEGIN
      IE_TAUXDEV:=1;IE_DEBITDEV:=IE_DEBIT ;
      IE_CREDITDEV:=IE_CREDIT;IE_QUOTITE:=1 ;
      END ;
    END Else
    BEGIN
    IE_TAUXDEV:=1;IE_DEBITDEV:=IE_DEBIT ;
    IE_CREDITDEV:=IE_CREDIT;IE_QUOTITE:=1 ;
    END ;
  If Arrondi(IE_DEBITEURO,V_PGI.OkDecE)=0 Then IE_DEBITEURO:=PivotToEuro(IE_DEBIT) ;
  If Arrondi(IE_CREDITEURO,V_PGI.OkDecE)=0 Then IE_CREDITEURO:=PivotToEuro(IE_CREDIT) ;
  END ;
END ;
*)


Procedure retoucheEnrCEGID(St : String ; Var IdentPiece : TIdentPiece ; Var InfoImp : TInfoImport;
                           Var Mvt : TFMvtImport ; OkRetoucheMontant : Boolean) ;
Var CptLuG,CptLuT,CptLuJ,CptLuMP : TCptLu ;
    Lettrable,Pointable,IsTiersT,IsTiersG : Boolean ;
    RIB : String ;
    i : Integer ;
BEGIN
Fillchar(CptLuMP,SizeOf(CptLuMP),#0) ;
InitRetoucheEnr(CptLuG,CptLuT,CptLuJ,Lettrable,Pointable,IsTiersT,IsTiersG,Mvt,InfoImp) ;
With Mvt do
  BEGIN
  If (CptLuG.Lettrable Or CptLuT.Lettrable) And (IE_TYPEECR='E') Then
     BEGIN
     if Trim(Copy(St,122,8))<>'' then IE_DATEECHEANCE:=Format_Date_HAL(Copy(St,122,8))
                                  else IE_DATEECHEANCE:=IE_DATECOMPTABLE ;
     IE_ECHE:='X' ;
     IE_DATEPAQUETMIN:=IE_DATECOMPTABLE ; IE_DATEPAQUETMAX:=IE_DATECOMPTABLE ;
     IE_ETATLETTRAGE:='AL' ;
     If (InfoImp.Format='CGN') Or (InfoImp.Format='CGE') Then IE_MODEPAIE:=Trim(Copy(St,119,3)) ;
     END ;
  If CptLuG.Pointable Then
    BEGIN
    IE_MODEPAIE:=Trim(Copy(St,119,3)) ;
    if Trim(Copy(St,122,8))<>'' then IE_DATEECHEANCE:=Format_Date_HAL(Copy(St,122,8))
                                else IE_DATEECHEANCE:=IE_DATECOMPTABLE ;
    END ;
  If InfoImp.Format='CGE' Then
     BEGIN
     IE_REFEXTERNE:=Trim(Copy(St,223,35)) ;
     If Trim(Copy(St,258,8))<>'' Then IE_DATEREFEXTERNE:=Format_Date_HAL(Copy(St,258,8)) ;
     If Trim(Copy(St,266,8))<>'' Then IE_DATECREATION:=Format_Date_HAL(Copy(St,266,8)) ;
     IE_AFFAIRE:=Trim(Copy(St,277,17)) ;
     If Trim(Copy(St,294,8))<>'' Then IE_DATETAUXDEV:=Format_Date_HAL(Copy(St,294,8)) ;
     If Trim(Copy(St,302,3))<>'' Then IE_ECRANOUVEAU:=Trim(Copy(St,302,3)) ;
     IE_QTE1:=Valeur(StPoint(Copy(St,305,20)));
     IE_QTE2:=Valeur(StPoint(Copy(St,325,20)));
     IE_QUALIFQTE1:=Trim(Copy(St,345,3));
     IE_QUALIFQTE2:=Trim(Copy(St,348,3));
     IE_REFLIBRE:=Trim(Copy(St,351,35)) ;
     if Trim(Copy(St,386,1))<>'' then IE_TVAENCAISSEMENT:=Trim(Copy(St,386,1));
     if Trim(Copy(St,390,3))<>'' then IE_TVA:=Copy(St,390,3);
     if Trim(Copy(St,393,3))<>'' then IE_TPF:=Copy(St,393,3);
     if Trim(Copy(St,396,17))<>'' Then IE_CONTREPARTIEGEN:=Copy(St,396,17);
     if Trim(Copy(St,413,3))<>'' Then  IE_CONTREPARTIEAUX:=Copy(St,413,17);
     IE_REFPOINTAGE:=Copy(St,430,17);
     If Trim(Copy(St,447,8))<>'' Then IE_DATEPOINTAGE:=Format_Date_HAL(Copy(St,447,8)) ;
     If Trim(Copy(St,455,8))<>'' Then IE_DATERELANCE:=Format_Date_HAL(Copy(St,45,8)) ;
     If Trim(Copy(St,463,8))<>'' Then IE_DATEVALEUR:=Format_Date_HAL(Copy(St,463,8)) ;
     Rib:=Trim(Copy(St,471,35)) ;
     If RIB<>'' Then
        BEGIN
        i:=Pos('/',RIB) ;
        If i<=0 Then IE_RIB:=Copy(RIB,1,5)+'/'+Copy(RIB,6,5)+'/'+Copy(RIB,11,11)+'/'+Copy(RIB,22,2)+'/'+Copy(RIB,24,12) Else
          BEGIN
          IE_RIB:=Trim(RIB) ;
          END ;
        END ;
     IE_REFRELEVE:=Copy(St,506,10) ;
     IE_NUMEROIMMO:=Copy(St,516,17) ;
     IE_LIBRETEXTE0:=Copy(St,533,30) ;
     IE_LIBRETEXTE1:=Copy(St,563,30) ;
     IE_LIBRETEXTE2:=Copy(St,593,30) ;
     IE_LIBRETEXTE3:=Copy(St,623,30) ;
     IE_LIBRETEXTE4:=Copy(St,653,30) ;
     IE_LIBRETEXTE5:=Copy(St,683,30) ;
     IE_LIBRETEXTE6:=Copy(St,713,30) ;
     IE_LIBRETEXTE7:=Copy(St,743,30) ;
     IE_LIBRETEXTE8:=Copy(St,773,30) ;
     IE_LIBRETEXTE9:=Copy(St,803,30) ;
     IE_TABLE0:=Copy(St,833,3) ;
     IE_TABLE1:=Copy(St,836,3) ;
     IE_TABLE2:=Copy(St,839,3) ;
     IE_TABLE3:=Copy(St,842,3) ;
     IE_LIBREMONTANT0:=Valeur(StPoint(Copy(St,845,20)));
     IE_LIBREMONTANT1:=Valeur(StPoint(Copy(St,865,20)));
     IE_LIBREMONTANT2:=Valeur(StPoint(Copy(St,885,20)));
     IE_LIBREMONTANT3:=Valeur(StPoint(Copy(St,905,20)));
     If Trim(Copy(St,925,8))<>'' Then IE_LIBREDATE:=Format_Date_HAL(Copy(St,925,8)) ;
     If Trim(Copy(St,933,1))<>'' Then IE_LIBREBOOL0:=Copy(St,933,1) ;
     If Trim(Copy(St,934,1))<>'' Then IE_LIBREBOOL0:=Copy(St,934,1) ;
     IE_CONSO:=Copy(St,935,3) ;
     if (IE_TYPEECR='L') then
       BEGIN
       IE_COUVERTURE:=Valeur(StPoint(Copy(St,938,20)));
       IE_LETTRAGE:=Copy(St,1014,5);
       IE_LETTRAGEDEV:=Copy(St,1019,1);
       if Trim(Copy(St,998,8))<>'' then IE_DATEPAQUETMAX:=Format_Date_HAL(Copy(St,998,8));
       if Trim(Copy(St,1006,8))<>'' then IE_DATEPAQUETMIN:=Format_Date_HAL(Copy(St,1006,8));
       IE_COUVERTUREDEV:=Valeur(StPoint(Copy(St,958,20)));
       IE_COUVERTUREEURO:=Valeur(StPoint(Copy(St,978,20)));
       if Trim(Copy(St,1021,3))<>'' then IE_ETATLETTRAGE:=Copy(St,1021,3);
       END ;
     END ;
  END ;
RazInfoLettrage((IsTiersG or IsTiersT),CptLuG.Pointable,Mvt) ;
If OkRetoucheMontant Then RetoucheMontantCEGID(Mvt,IdentPiece) ;
END ;

Procedure retoucheEnrAutre(St : String ; Var IdentPiece : TIdentPiece ; Var InfoImp : TInfoImport ;
                           Var Mvt : TFMvtImport ; OkRetoucheMontant : Boolean) ;
Var CptLuG,CptLuT,CptLuJ : TCptLu ;
    Lettrable,Pointable,IsTiersT,IsTiersG : Boolean ;
BEGIN
InitRetoucheEnr(CptLuG,CptLuT,CptLuJ,Lettrable,Pointable,IsTiersT,IsTiersG,Mvt,InfoImp) ;
With Mvt do
  BEGIN
  If (CptLuG.Lettrable Or CptLuT.Lettrable) And (IE_TYPEECR='E') Then
     BEGIN
     if ((Trim(Copy(St,78,6))<>''))  then IE_DATEECHEANCE:=Format_Date(Copy(St,78,6)) else IE_DATEECHEANCE:=IE_DATECOMPTABLE ;
     IE_ECHE:='X' ;
     IE_DATEPAQUETMIN:=IE_DATECOMPTABLE ; IE_DATEPAQUETMAX:=IE_DATECOMPTABLE ;
     IE_ETATLETTRAGE:='AL' ;
     TrouveLeModePaie(Copy(St,77,1),InfoImp,Mvt) ;
     END ;
  If CptLuG.Pointable Then
    BEGIN
    TrouveLeModePaie(Copy(St,77,1),InfoImp,Mvt) ;
    if ((Trim(Copy(St,78,6))<>''))  then IE_DATEECHEANCE:=Format_Date(Copy(St,78,6))
                                    else IE_DATEECHEANCE:=IE_DATECOMPTABLE ;
    END ;
  END ;
RazInfoLettrage((IsTiersG or IsTiersT),CptLuG.Pointable,Mvt) ;
If OkRetoucheMontant Then RetoucheMontantAutre(Mvt,IdentPiece) ;
END ;


Function TestRupturePiece(Format : String ; AlimOldIdentPiece : Boolean ; Var IdentPiece,OldIdentPiece : TIdentPiece) : Boolean ;
BEGIN
Result:=FALSE ;
if (Format='RAP') or (Format='CRA') or (Trim(Format)='MP') then Result:=TRUE
                                                           Else Result:=EnRuptAscii(AlimOldIdentPiece,IdentPiece,OldIdentPiece) ;
END ;

Procedure InitAxe(St : String ; Format : String ; Var Mvt : TFMvtImport) ;
BEGIN
If Mvt.IE_TYPEECR<>'A' Then BEGIN Mvt.IE_AXE:='' ; Exit ; END ;
If Mvt.IE_ANA='X' Then
   BEGIN
   If (Format='HAL') Or (Format='CGN') Or (Format='CGE') Then
      BEGIN
      Case Format[1] Of
        'C' : Mvt.IE_AXE:=Trim(Copy(St,219,2)) ;
        'H' : if not VH^.ImportRL Then Mvt.IE_AXE:=Trim(Copy(St,125,2))
                                  Else Mvt.IE_AXE:=Trim(Copy(St,117,2)) ;
        Else Mvt.IE_AXE:='A1' ;
        End ;
      END Else
      BEGIN
      Mvt.IE_AXE:='A1' ;
      END ;
   If (Mvt.IE_AXE<>'A1') And (Mvt.IE_AXE<>'A2') And
      (Mvt.IE_AXE<>'A3') And (Mvt.IE_AXE<>'A4') And
      (Mvt.IE_AXE<>'A5') Then Mvt.IE_AXE:='A1' ;
   END ;
END ;

Function AnalytiqueAIgnorer(Var Mvt : TFMvtImport ; Var InfoImp : TInfoImport ; Var CptLuG : TCptLu) : Boolean ;
BEGIN
Result:=FALSE ;
If (Mvt.IE_TYPEECR='L') Or (Mvt.IE_TYPEANALYTIQUE='X') Then Exit ;
If (Mvt.IE_TYPEECR='A') Then
  BEGIN
  If Not CptLuG.Ventilable Then Result:=TRUE Else
    BEGIN
    If (CptLuG.Axe[1]<>'X') And (Mvt.IE_AXE='A1') Then Result:=TRUE ;
    If (CptLuG.Axe[2]<>'X') And (Mvt.IE_AXE='A2') Then Result:=TRUE ;
    If (CptLuG.Axe[3]<>'X') And (Mvt.IE_AXE='A3') Then Result:=TRUE ;
    If (CptLuG.Axe[4]<>'X') And (Mvt.IE_AXE='A4') Then Result:=TRUE ;
    If (CptLuG.Axe[5]<>'X') And (Mvt.IE_AXE='A5') Then Result:=TRUE ;
    END ;
  END ;
END ;


Function AlimCptMemoire(QFiche : TQfiche ; Var Mvt : TFMvtImport ; Var InfoImp : TInfoImport ) : Integer ;
Var CptLuT,CptLuG,CptLuS,CptLuJ : TCptLu ;
    Lefb : TFichierBase ;
    EstVentilable,EstCollectif : Boolean ;
    NatureCptGen : String ;
BEGIN
Result:=0 ;
Fillchar(CptLuG,SizeOf(CptLuG),#0) ;
Fillchar(CptLuT,SizeOf(CptLuT),#0) ;
Fillchar(CptLuS,SizeOf(CptLuS),#0) ;
Fillchar(CptLuJ,SizeOf(CptLuJ),#0) ;
EstVentilable:=FALSE ;  EstCollectif:=FALSE ;
NatureCptGen:='' ;
Mvt.IE_GENERAL:=BourreOuTronque(Mvt.IE_GENERAL,fbGene) ;
Mvt.IE_AUXILIAIRE:=BourreOuTronque(Mvt.IE_AUXILIAIRE,fbAux) ;
CptLuG.Cpt:=Mvt.IE_GENERAL ;
If AlimLTabCptLu(0,QFiche[0],InfoImp.LGenLu,InfoImp.ListeCptFaux,CptLuG) Then
   BEGIN
   If CptLuG.Cpt='' Then Inc(InfoImp.NbGenFaux) ;
   EstVentilable:=CptLuG.Ventilable ;
   NatureCptGen:=CptLuG.Nature ;
   EstCollectif:=(CptLuG.Nature='COC') Or (CptLuG.Nature='COF') Or
                 (CptLuG.Nature='COS') Or (CptLuG.Nature='COD') ;
   If CptLuG.Ventilable Then
      BEGIN
       Lefb:=fbAxe1 ;
       if (Mvt.IE_AXE='A2') then Lefb:=fbAxe2 else
       if (Mvt.IE_AXE='A3') then Lefb:=fbAxe3 else
       if (Mvt.IE_AXE='A4') then Lefb:=fbAxe4 else
       if (Mvt.IE_AXE='A5') then Lefb:=fbAxe5 ;
       Mvt.IE_SECTION:=BourreOuTronque(Mvt.IE_SECTION,Lefb) ;
      END Else If Mvt.IE_SECTION<>'' Then Mvt.IE_SECTION:=BourreOuTronque(Mvt.IE_SECTION,fbAxe1) ;
   If AnalytiqueAIgnorer(Mvt,InfoImp,CptLuG) Then Result:=1 ;
   END Else Inc(InfoImp.NbGenFaux) ;

CptLuT.Cpt:=Mvt.IE_AUXILIAIRE ; CptLuT.Collectif:=Mvt.IE_GENERAL ;
If AlimLTabCptLu(1,QFiche[1],InfoImp.LAuxLu,InfoImp.ListeCptFaux,CptLuT) Then
   BEGIN
   If Not EstCollectif Then Inc(InfoImp.NbAuxFaux) ;
   END Else
   BEGIN
   If EstCollectif Then Inc(InfoImp.NbAuxFaux) ;
   END ;

CptLuS.Cpt:=Mvt.IE_SECTION ; CptLuS.Collectif:=Mvt.IE_GENERAL ; CptLuS.Axe:=MVT.IE_AXE ;
If AlimLTabCptLu(2,QFiche[2],InfoImp.LAnaLu,InfoImp.ListeCptFaux,CptLuS) Then
   BEGIN
   If (Result=0) And (Not EstVentilable) Then Inc(InfoImp.NbAnaFaux) ;
   END Else
   BEGIN
   If EstVentilable And (MVT.IE_TYPEECR='A') Then Inc(InfoImp.NbAnaFaux) ;
   END ;

CptLuJ.Cpt:=Mvt.IE_JOURNAL ;
If AlimLTabCptLu(3,QFiche[3],InfoImp.LJalLu,InfoImp.ListeCptFaux,CptLuJ) Then Else Inc(InfoImp.NbJalFaux) ;
END ;

Procedure AlimIdentPiece(Var IdentPiece : TIdentPiece ; Var Mvt : TFMvtImport) ;
Var EnEuro : Boolean ;
BEGIN
IdentPiece.JalP:=Mvt.IE_JOURNAL        ; IdentPiece.NatP:=Mvt.IE_NATUREPIECE ;
IdentPiece.QualP:=Mvt.IE_QUALIFPIECE   ; IdentPiece.DateP:=Mvt.IE_DATECOMPTABLE ;
IdentPiece.NumP:=Mvt.IE_NUMPIECE ;
If Mvt.IE_TYPEECR='E' Then
  With IdentPiece Do
    BEGIN
    LignePrec.DP:=LigneEnCours.DP ; LignePrec.CP:=LigneEnCours.CP ;
    LignePrec.DD:=LigneEnCours.DD ; LignePrec.CD:=LigneEnCours.CD ;
    LignePrec.DE:=LigneEnCours.DE ; LignePrec.CE:=LigneEnCours.CE ;

    LigneEnCours.DP:=Mvt.IE_DEBIT ; LigneEnCours.CP:=Mvt.IE_CREDIT ;
    LigneEnCours.DD:=Mvt.IE_DEBITDEV ; LigneEnCours.CD:=Mvt.IE_CREDITDEV ;
    LigneEnCours.DE:=Mvt.IE_DEBITEURO ; LigneEnCours.CE:=Mvt.IE_CREDITEURO ;

    END ;
With IdentPiece Do
  BEGIN
  LignePrec.Gen:=LigneEnCours.Gen  ; LignePrec.Aux:=LigneEnCours.Aux ;
  LignePrec.Sect:=LigneEnCours.Sect ; LignePrec.Eche:=LigneEnCours.Eche ;
  LignePrec.Ana:=LigneEnCours.Ana ; LignePrec.Axe:=LigneEnCours.Axe ;

  LigneEnCours.Gen:=Mvt.IE_GENERAL ; LigneEnCours.Aux:=Mvt.IE_AUXILIAIRE ;
  LigneEnCours.Sect:=MVT.IE_SECTION ; LigneEnCours.Ana:=Mvt.IE_TYPEECR='A' ;
  LigneEnCours.Eche:=Mvt.IE_ECHE='X' ; LigneEnCours.Axe:=Mvt.IE_AXE ;
  END ;
EnEuro:=EuroOK And VH^.TenueEuro ;
If IdentPiece.CodeMontant<>'' Then
  BEGIN
  If IdentPiece.CodeMontant[1]='F' Then
    BEGIN
    If EnEuro Then Mvt.IE_SAISIEEURO:='X' ;
    END Else
  If IdentPiece.CodeMontant[1]='E' Then
    BEGIN
    If Not EnEuro Then Mvt.IE_SAISIEEURO:='X' ;
    END ;
  END ;
END ;

Procedure RAZTotAna(Var IdentPiece : TIdentPiece) ;
BEGIN
IdentPiece.TotDPAna:=0 ; IdentPiece.TotCPAna:=0 ; IdentPiece.TotDDAna:=0 ;
IdentPiece.TotCDAna:=0 ; IdentPiece.TotDEAna:=0 ; IdentPiece.TotCEAna:=0 ;
END ;

Procedure AlimNumeros(OkRupt : Boolean ; Var IdentPiece : TIdentPiece ; Var InfoImp : TInfoImport ; Var Mvt : TFMvtImport) ;
Var OkIncLigne : Boolean ;
    CptLuG : TCptLu ;
    OkSens : Boolean ;
BEGIN
OkIncLigne:=TRUE ;
If OkRupt Then
   BEGIN
   IdentPiece.LigneEnCours.NumLig:=0 ;
   IdentPiece.LigneEnCours.NumEche:=0 ;
   IdentPiece.LigneEnCours.NumV:=0 ;
   IdentPiece.TotDP:=0 ; IdentPiece.TotCP:=0 ; IdentPiece.TotDD:=0 ;
   IdentPiece.TotCD:=0 ; IdentPiece.TotDE:=0 ; IdentPiece.TotCE:=0 ;
   RazTotAna(IdentPiece) ;
   IdentPiece.DernChronoE:=0 ; IdentPiece.DernChronoENonVentilable:=0 ;
   IdentPiece.DernChronoEcrAna:=0 ;
//   Inc(IdentPiece.NumP) ;
   END Else
   BEGIN
   If Mvt.IE_TYPEANALYTIQUE<>'X' Then
     BEGIN
     OkSens:=TRUE ;
     If ((Arrondi(IdentPiece.LigneEnCours.DP,V_PGI.OkDecV)=0) And (Arrondi(IdentPiece.LignePrec.DP,V_PGI.OkDecV)<>0)) Or
        ((Arrondi(IdentPiece.LigneEnCours.CP,V_PGI.OkDecV)=0) And (Arrondi(IdentPiece.LignePrec.CP,V_PGI.OkDecV)<>0)) Then OkSens:=FALSE ;
     If (IdentPiece.LignePrec.Eche) And (IdentPiece.LigneEnCours.Eche) and
        (IdentPiece.LigneEnCours.Gen=IdentPiece.LignePrec.Gen) And
        (IdentPiece.LigneEnCours.Aux=IdentPiece.LignePrec.Aux) And OkSens
        Then OkIncLigne:=FALSE Else IdentPiece.LigneEnCours.NumEche:=0 ;
     If (*(IdentPiece.LignePrec.Ana) And *)(IdentPiece.LigneEnCours.Ana) and
        (IdentPiece.LigneEnCours.Gen=IdentPiece.LignePrec.Gen) Then
  //      (IdentPiece.LigneEnCours.Sect=IdentPiece.LignePrec.Sect)
        BEGIN
        OkIncLigne:=FALSE ;
        If IdentPiece.LigneEnCours.Axe<>IdentPiece.LignePrec.Axe Then IdentPiece.LigneEnCours.NumV:=0 ;
        END Else IdentPiece.LigneEnCours.NumV:=0 ;
     END ;
   END ;
If Mvt.IE_TYPEANALYTIQUE='X' Then
  BEGIN
  OkIncLigne:=FALSE ;
  END ;
If IdentPiece.LigneEnCours.Eche Then Inc(IdentPiece.LigneEnCours.NumEche) ;
If IdentPiece.LigneEnCours.Ana Then Inc(IdentPiece.LigneEnCours.NumV) ;
If OkIncLigne Then Inc(IdentPiece.LigneEnCours.NumLig) ;
Inc(IdentPiece.Chrono) ;
Mvt.IE_CHRONO:=IdentPiece.Chrono ;
Mvt.IE_NUMPIECE:=IdentPiece.NumP ;
Mvt.IE_NUMLIGNE:=IdentPiece.LigneEnCours.NumLig ;
Mvt.IE_NUMECHE:=IdentPiece.LigneEnCours.NumEche ;
Mvt.IE_NUMVENTIL:=IdentPiece.LigneEnCours.NumV ;
If IdentPiece.LigneEnCours.NumV<=0 Then BEGIN IdentPiece.DernChronoEcrAna:=0 ; RAZTotAna(IdentPiece) ; END
                                   Else IdentPiece.DernChronoEcrAna:=IdentPiece.Chrono ;
If Mvt.IE_TYPEECR='E' Then
  BEGIN
  IdentPiece.DernChronoE:=IdentPiece.Chrono ;
  Fillchar(CptLuG,SizeOf(CptLuG),#0) ; CptLuG.Cpt:=MVT.IE_GENERAL ;
  If (Not IdentPiece.LigneEnCours.Ana) And ChercheCptLu(InfoImp.LGenLu,CptLuG) Then
    BEGIN
    If Not CptLuG.Ventilable Then IdentPiece.DernChronoENonVentilable:=IdentPiece.Chrono ;
    END ;
  IdentPiece.TotDP:=Arrondi(IdentPiece.TotDP+MVT.IE_DEBIT,V_PGI.okDecV) ;
  IdentPiece.TotCP:=Arrondi(IdentPiece.TotCP+MVT.IE_CREDIT,V_PGI.okDecV) ;
  IdentPiece.TotDD:=Arrondi(IdentPiece.TotDD+MVT.IE_DEBITDEV,IdentPiece.DecimDev) ;
  IdentPiece.TotCD:=Arrondi(IdentPiece.TotCD+MVT.IE_CREDITDEV,IdentPiece.DecimDev) ;
  IdentPiece.TotDE:=Arrondi(IdentPiece.TotDE+MVT.IE_DEBITEURO,V_PGI.okDecE) ;
  IdentPiece.TotCE:=Arrondi(IdentPiece.TotCE+MVT.IE_CREDITEURO,V_PGI.okDecE) ;
  END ;
If Mvt.IE_TYPEECR='A' Then
  BEGIN
  IdentPiece.TotDPAna:=Arrondi(IdentPiece.TotDPAna+MVT.IE_DEBIT,V_PGI.okDecV) ;
  IdentPiece.TotCPAna:=Arrondi(IdentPiece.TotCPAna+MVT.IE_CREDIT,V_PGI.okDecV) ;
  IdentPiece.TotDDAna:=Arrondi(IdentPiece.TotDDAna+MVT.IE_DEBITDEV,IdentPiece.DecimDev) ;
  IdentPiece.TotCDAna:=Arrondi(IdentPiece.TotCDAna+MVT.IE_CREDITDEV,IdentPiece.DecimDev) ;
  IdentPiece.TotDEAna:=Arrondi(IdentPiece.TotDEAna+MVT.IE_DEBITEURO,V_PGI.okDecE) ;
  IdentPiece.TotCEAna:=Arrondi(IdentPiece.TotCEAna+MVT.IE_CREDITEURO,V_PGI.okDecE) ;
  END ;
END ;

procedure RetoucheFinal(Var IdentPiece : TIdentPiece ; Var Mvt : TFMvtImport) ;
Var TotEcr : Double ;
BEGIN
If (Mvt.IE_TYPEECR='L') Then BEGIN Mvt.IE_OKCONTROLE:='-' ; Mvt.IE_SELECTED:='-' ;  END ;
If IdentPiece.Doublon Then Mvt.IE_OKCONTROLE:='D' ;
if (Mvt.IE_TYPEECR<>'A') then Exit ;
If Mvt.IE_TYPEANALYTIQUE='X' Then
  BEGIN
  Mvt.IE_NUMLIGNE:=0 ; MVT.IE_POURCENTAGE:=0 ;
  END Else
  BEGIN
  TotEcr:=IdentPiece.LigneEnCours.DP+IdentPiece.LigneEnCours.CP ;
  With Mvt do
    BEGIN
    if (IE_POURCENTAGE=0) and (TotEcr<>0) then IE_POURCENTAGE:=Arrondi(((Mvt.IE_DEBIT+Mvt.IE_CREDIT)/TotEcr)*100,ADecimP) ;
    if IE_POURCENTQTE1=0 then IE_POURCENTQTE1:=IE_POURCENTAGE ;
    IE_POURCENTQTE2:=IE_POURCENTQTE1 ;
    END ;
  END ;
END ;

Procedure RetoucheODANAL(Var Mvt : TFMvtImport) ;
BEGIN
if (Mvt.IE_TYPEECR<>'A') then Exit ;
If (Mvt.IE_TYPEANALYTIQUE<>'X') Then Exit ;
Mvt.IE_QUALIFPIECE:='N' ;
Mvt.IE_ECRANOUVEAU:='N';
Mvt.IE_DEVISE:=V_PGI.DevisePivot ;
Mvt.IE_DEBITDEV:=MVT.IE_DEBIT ;
Mvt.IE_CREDITDEV:=MVT.IE_CREDIT ;
Mvt.IE_TAUXDEV:=1 ; MVT.IE_COTATION:=1 ; 
END ;

Function ChercheDoublon(QD : TQDoublon ; Var InfoImp : TInfoImport ; Var Mvt : TFMvtImport) : Boolean ;
Var What : Integer ;
    X    : DelInfo ;
    CleDb : String ;
    StD,StC : String ;
BEGIN
Result:=FALSE ;
What:=0 ; If Mvt.IE_TYPEANALYTIQUE='X' Then What:=1 ;
QD[What].Params[0].AsString:=Mvt.IE_JOURNAL ;
QD[What].Params[1].AsString:=QUELEXODT(Mvt.IE_DATECOMPTABLE) ;
QD[What].Params[2].AsDateTime:=Mvt.IE_DATECOMPTABLE ;
QD[What].Params[3].AsString:=Mvt.IE_GENERAL ;
Case What Of
  0 : QD[What].Params[4].AsString:=Mvt.IE_AUXILIAIRE ;
  1 : QD[What].Params[4].AsString:=Mvt.IE_SECTION ;
  END ;
QD[What].Params[5].AsString:=Mvt.IE_REFINTERNE ;
QD[What].Params[6].AsString:=Mvt.IE_LIBELLE ;
QD[What].Params[7].AsString:=Mvt.IE_DEVISE ;
StD:=Formatfloat('###0.00',Mvt.IE_DEBITDEV) ; StD:=FindEtReplace(StD,',','.',TRUE) ;
StC:=Formatfloat('###0.00',Mvt.IE_CREDITDEV) ; StC:=FindEtReplace(StC,',','.',TRUE) ;
QD[What].Params[8].AsString:=StD ;
QD[What].Params[9].AsString:=StC ;
Case What Of
  1 : QD[What].Params[10].AsString:=Mvt.IE_AXE ;
  END ;
QD[What].Open ;
If Not QD[What].Eof Then Result:=TRUE ;
If Result Then
  BEGIN
  X:=DelInfo.Create ; X.LeCod:=Mvt.IE_Journal ; X.LeLib:=IntToStr(MVT.IE_NumPiece);
  X.LeMess:=TraduireMemoire('N° Pièce')+' : '+IntToStr(QD[What].Fields[3].AsInteger)+' - '+TraduireMemoire('Date')+' '+DateToStr(QD[What].Fields[2].AsDateTime)+
            ' - '+TraduireMemoire('Référence')+' '+QD[What].Fields[7].AsString ;
  InfoImp.CRListeEnteteDoublon.Add(X) ;
  CleDB:=Format_String(Mvt.IE_Journal,3)+Format_String(' ',1)+FormatFloat('00000000',MVT.IE_NumPiece) ;
  InfoImp.ListeEnteteDoublon.Add(CleDB) ;
  END ;
QD[What].Close ;
END ;

Function AnalytiqueAbsenteSurLignePrecedente(Var St,OldSt : String ; Var InfoImp : TInfoImport) : Boolean ;
Var FmtFic : Integer ;
    Enr,OldEnr : TStLue ;
    Cpt : String ;
    CptLu : TCptLu ;
    AnalAbsente : Boolean ;
    Decal,LgCpte : Integer ;
BEGIN
Result:=FALSE ; If OldSt='' Then Exit ; Decal:=0 ;
if InfoImp.Format='SAA' then FmtFic:=0 else if InfoImp.Format='SN2' then FmtFic:=0 else
if InfoImp.Format='HLI' then FmtFic:=1 else if InfoImp.Format='HAL' then FmtFic:=2 else
if InfoImp.Format='CGN' then FmtFic:=3 else if InfoImp.Format='CGE' then FmtFic:=4 Else Exit ;
AlimEnr(St,Enr,FmtFic) ; AlimEnr(OldSt,OldEnr,FmtFic) ;
If OldEnr.TC='O' Then Exit ;
If OldEnr.TC='L' Then Exit ;
Cpt:=Trim(OldEnr.Jal) ;
Fillchar(CptLu,SizeOf(CptLu),#0) ; CptLu.Cpt:=Cpt ;
If ChercheCptLu(InfoImp.LJalLu,CptLu) Then
  BEGIN
  If CptLu.Nature='ODA' Then Exit ;
  END Else Exit ;
Cpt:=BourreOuTronque(Trim(OldEnr.General),fbGene) ;
Fillchar(CptLu,SizeOf(CptLu),#0) ; CptLu.Cpt:=Cpt ;
If ChercheCptLu(InfoImp.LGenLu,CptLu) Then
  BEGIN
  AnalAbsente:=CptLu.Ventilable And ((Enr.TC<>'A') Or (St='')) And (OldEnr.TC<>'H') And (OldEnr.TC=' ') And (CptLu.Axe[1]='X') ;
  If AnalAbsente Then
    BEGIN
    Result:=TRUE ;
    LgCpte:=13 ; if not VH^.ImportRL then LgCpte:=17 ;
    Case FmtFic Of
      0   : BEGIN
            OldSt:=Insere(OldSt,'A',25,1) ;
            OldSt:=Insere(OldSt,Format_String(VH^.Cpta[fbAxe1].Attente,13),26,13) ;
            END ;
      1   : BEGIN
            If Not VH^.ImportRL Then Inc(Decal,4) ;
            OldSt:=Insere(OldSt,'A',27+Decal,1) ;
            OldSt:=Insere(OldSt,Format_String(VH^.Cpta[fbAxe1].Attente,LgCpte),28+Decal,LgCpte) ;
            END ;
      2   : BEGIN
            If Not VH^.ImportRL Then Inc(Decal,4) ;
            OldSt:=Insere(OldSt,'A',27+Decal,1) ;
            OldSt:=Insere(OldSt,Format_String(VH^.Cpta[fbAxe1].Attente,LgCpte),28+Decal,LgCpte) ;
            If Not VH^.ImportRL Then Inc(Decal,4) ;
            OldSt:=Insere(OldSt,'A1',117+Decal,2) ;
            END ;
      3,4 : BEGIN
            OldSt:=Insere(OldSt,'A',31,1) ;
            OldSt:=Insere(OldSt,Format_String(VH^.Cpta[fbAxe1].Attente,17),32,17) ;
            OldSt:=Insere(OldSt,'A1',219,2) ;
            END ;
      END ;
    END ;
  END ;
END ;

Procedure RecupMontant(What : Integer ; EnEuro : Boolean ; Var DSaisi,CSaisi,DOppose,COppose,DDevise,CDevise : Double ; Var IdentPiece : TIdentPiece ;
                       Var QuiOppose : Char ; Var DecimSaisi,DecimOppose : Integer ;
                       Var TotEcrDSaisi,TotEcrCSaisi,TotEcrDOppose,TotEcrCOppose,TotEcrDDevise,TotEcrCDevise : Double) ;
BEGIN
DSaisi:=0 ; CSaisi:=0 ; DOppose:=0 ; COppose:=0 ; QuiOppose:='L' ;
DDevise:=0 ; CDevise:=0 ; DecimSaisi:=V_PGI.okDecV ; DecimOppose:=V_PGI.OkDecE ;
TotEcrDSaisi:=0 ; TotEcrCSaisi:=0 ; TotEcrDOppose:=0 ;
TotEcrCOppose:=0 ;
If IdentPiece.CodeMontant[1]='F' Then
  BEGIN
  If EnEuro Then
    BEGIN
    Case What Of
      0 : BEGIN // Pour Equilibrage Ecriture
          DSaisi:=IdentPiece.TotDE ; CSaisi:=IdentPiece.TotCE ;
          DOppose:=IdentPiece.TotDP ; COppose:=IdentPiece.TotCP ;
          END ;
      1 : BEGIN // Pour Equilibrage Analytique
          DSaisi:=IdentPiece.TotDEAna ; CSaisi:=IdentPiece.TotCEAna ;
          DOppose:=IdentPiece.TotDPAna ; COppose:=IdentPiece.TotCPAna ;
          TotEcrDSaisi:=IdentPiece.LigneEnCours.DE ; TotEcrCSaisi:=IdentPiece.LigneEnCours.CE ;
          TotEcrDOppose:=IdentPiece.LigneEnCours.DP ; TotEcrCOppose:=IdentPiece.LigneEnCours.CP ;
          END ;
      END ;
    QuiOppose:='P' ; DecimSaisi:=V_PGI.OkDecE ; DecimOppose:=V_PGI.OkDecV ;
    END Else
    BEGIN
    Case What Of
      0 : BEGIN // Pour Equilibrage Ecriture
          DSaisi:=IdentPiece.TotDP ; CSaisi:=IdentPiece.TotCP ;
          DOppose:=IdentPiece.TotDE ; COppose:=IdentPiece.TotCE ;
          END ;
      1 : BEGIN // Pour Equilibrage Analytique
          DSaisi:=IdentPiece.TotDPAna ; CSaisi:=IdentPiece.TotCPAna ;
          DOppose:=IdentPiece.TotDEAna ; COppose:=IdentPiece.TotCEAna ;
          TotEcrDSaisi:=IdentPiece.LigneEnCours.DP ; TotEcrCSaisi:=IdentPiece.LigneEnCours.CP ;
          TotEcrDOppose:=IdentPiece.LigneEnCours.DE ; TotEcrCOppose:=IdentPiece.LigneEnCours.CE ;
          END ;
      END ;
    QuiOppose:='E' ; DecimSaisi:=V_PGI.OkDecV ; DecimOppose:=V_PGI.OkDecE ;
    END ;
  END Else
If IdentPiece.CodeMontant[1]='E' Then
  BEGIN
    If EnEuro Then
    BEGIN
    Case What Of
      0 : BEGIN // Pour Equilibrage Ecriture
          DSaisi:=IdentPiece.TotDP ; CSaisi:=IdentPiece.TotCP ;
          DOppose:=IdentPiece.TotDE ; COppose:=IdentPiece.TotCE ;
          END ;
      1 : BEGIN // Pour Equilibrage Analytique
          DSaisi:=IdentPiece.TotDPAna ; CSaisi:=IdentPiece.TotCPAna ;
          DOppose:=IdentPiece.TotDEAna ; COppose:=IdentPiece.TotCEAna ;
          TotEcrDSaisi:=IdentPiece.LigneEnCours.DP ; TotEcrCSaisi:=IdentPiece.LigneEnCours.CP ;
          TotEcrDOppose:=IdentPiece.LigneEnCours.DE ; TotEcrCOppose:=IdentPiece.LigneEnCours.CE ;
          END ;
      END ;
    QuiOppose:='E' ; DecimSaisi:=V_PGI.OkDecV ; DecimOppose:=V_PGI.OkDecE ;
    END Else
    BEGIN
    Case What Of
      0 : BEGIN // Pour Equilibrage Ecriture
          DSaisi:=IdentPiece.TotDE ; CSaisi:=IdentPiece.TotCE ;
          DOppose:=IdentPiece.TotDP ; COppose:=IdentPiece.TotCP ;
          END ;
      1 : BEGIN // Pour Equilibrage Analytique
          DSaisi:=IdentPiece.TotDEAna ; CSaisi:=IdentPiece.TotCEAna ;
          DOppose:=IdentPiece.TotDPAna ; COppose:=IdentPiece.TotCPAna ;
          TotEcrDSaisi:=IdentPiece.LigneEnCours.DE ; TotEcrCSaisi:=IdentPiece.LigneEnCours.CE ;
          TotEcrDOppose:=IdentPiece.LigneEnCours.DP ; TotEcrCOppose:=IdentPiece.LigneEnCours.CP ;
          END ;
      END ;
    QuiOppose:='P' ; DecimSaisi:=V_PGI.OkDecE ; DecimOppose:=V_PGI.OkDecV ;
    END ;
  END Else
If IdentPiece.CodeMontant[1]='D' Then
  BEGIN
    Case What Of
      0 : BEGIN // Pour Equilibrage Ecriture
          DDevise:=IdentPiece.TotDD ; CDevise:=IdentPiece.TotCD ;
          DSaisi:=IdentPiece.TotDP ; CSaisi:=IdentPiece.TotCP ;
          DOppose:=IdentPiece.TotDE ; COppose:=IdentPiece.TotCE ;
          END ;
      1 : BEGIN // Pour Equilibrage Analytique
          DDevise:=IdentPiece.TotDDAna ; CDevise:=IdentPiece.TotCDAna ;
          DSaisi:=IdentPiece.TotDPAna ; CSaisi:=IdentPiece.TotCPAna ;
          DOppose:=IdentPiece.TotDEAna ; COppose:=IdentPiece.TotCEAna ;
          TotEcrDDevise:=IdentPiece.LigneEnCours.DD ; TotEcrCDevise:=IdentPiece.LigneEnCours.CD ;
          TotEcrDSaisi:=IdentPiece.LigneEnCours.DP ; TotEcrCSaisi:=IdentPiece.LigneEnCours.CP ;
          TotEcrDOppose:=IdentPiece.LigneEnCours.DE ; TotEcrCOppose:=IdentPiece.LigneEnCours.CE ;
          END ;
      END ;
  QuiOppose:='D' ;
  END ;
END ;

Procedure AjusteSQLMontant(What : Integer ; Decim : Integer ; NomChampDebit,NomChampCredit : String ;
                           Var DeltaOppose,DebitPLu,CreditPLu : Double ; Var StDebit,StCredit : String) ;
BEGIN
If DeltaOppose<>0 Then
  BEGIN
  If Arrondi(DebitPLu,Decim)<>0 Then
    BEGIN
    Case What Of
      0 : If DeltaOppose>0 Then DebitPLu:=DebitPLu-Abs(DeltaOppose) Else DebitPLu:=DebitPLu+Abs(DeltaOppose) ;
      1 : If DeltaOppose>0 Then DebitPLu:=DebitPLu+Abs(DeltaOppose) Else DebitPLu:=DebitPLu-Abs(DeltaOppose)
      End ;
    StDebit:=NomChampDebit+'='+StrfPoint(DebitPLu)+' ' ;
    END Else If Arrondi(CreditPLu,Decim)<>0 Then
    BEGIN
    Case What Of
      0 : If DeltaOppose>0 Then CreditPLu:=CreditPLu+abs(DeltaOppose) Else CreditPLu:=CreditPLu-Abs(DeltaOppose) ;
      1 : If DeltaOppose>0 Then CreditPLu:=CreditPLu+abs(DeltaOppose) Else CreditPLu:=CreditPLu-Abs(DeltaOppose) ;
      End ;
    StCredit:=NomChampCredit+'='+StrfPoint(CreditPLu)+' ' ;
    END ;
  END ;
END ;
(*
Procedure UpdateEcrPourEquilibre(EnEuro : Boolean ; QuiOppose : Char ; DeltaOppose,DeltaOppose2 : Double ; Var IdentPiece : TIdentPiece) ;
Var Q : TQuery ;
    StSQL,STDEBIT,STCREDIT,STDEBITEURO,STCREDITEURO : String ;
    DebitELu,CreditELu,DebitPLu,CreditPLu : Double ;
    LeChrono : Integer ;
BEGIN
If IdentPiece.DernChronoENonVentilable<>0 Then LeChrono:=IdentPiece.DernChronoENonVentilable
                                          Else LeChrono:=IdentPiece.DernChronoE ;
If LeChrono=0 Then Exit ;
STDEBIT:='' ; STCREDIT:='' ; STDEBITEURO:='' ; STCREDITEURO:='' ;
StSQL:='SELECT IE_DEBIT,IE_CREDIT,IE_DEBITEURO,IE_CREDITEURO FROM IMPECR WHERE IE_CHRONO='+IntToStr(LeChrono) ;
Q:=OpenSQL(StSQL,TRUE) ;
If Not Q.Eof Then
  BEGIN
  DebitPLu:=Q.Fields[0].AsFloat ; CreditPLu:=Q.Fields[1].AsFloat ;
  DebitELu:=Q.Fields[2].AsFloat ; CreditELu:=Q.Fields[3].AsFloat ;
  Ferme(Q) ;
  AjusteSQLMontant(0,V_PGI.OkDecV,'IE_DEBIT','IE_CREDIT',DeltaOppose,DebitPLu,CreditPLu,StDebit,StCredit) ;
  AjusteSQLMontant(0,V_PGI.OkDecE,'IE_DEBITEURO','IE_CREDITEURO',DeltaOppose2,DebitELu,CreditELu,StDebitEuro,StCreditEuro) ;
  StSQL:='' ;
  If StDebit<>'' Then StSQL:=StSQL+StDebit+', ' ;
  If StCredit<>'' Then StSQL:=StSQL+StCredit+', ' ;
  If StDebitEuro<>'' Then StSQL:=StSQL+StDebitEuro+', ' ;
  If StCreditEuro<>'' Then StSQL:=StSQL+StCreditEuro+', ' ;
  If StSQL<>'' Then
    BEGIN
    StSQL:=Copy(StSQL,1,Length(StSQL)-2) ;
    StSQL:='UPDATE IMPECR SET '+StSQL+' WHERE IE_CHRONO='+IntToStr(LeChrono) ;
    ExecuteSQL(StSQL) ;
    END ;
  END Else
  BEGIN
  Ferme(Q) ;
  END ;
END ;
*)
Procedure UpdatePourEquilibre(SurAna,SurODA : Boolean ; EnEuro : Boolean ; QuiOppose : Char ; DeltaOppose,DeltaOppose2 : Double ; Var IdentPiece : TIdentPiece) ;
Var Q : TQuery ;
    StSQL,STDEBIT,STCREDIT,STDEBITEURO,STCREDITEURO,STDEBITDEV,STCREDITDEV : String ;
    DebitDLu,DebitELu,CreditELu,DebitPLu,CreditPLu : Double ;
    DebitP,CreditP,DebitD,CreditD,DebitE,CreditE : Double ;
    LeChrono : Integer ;
    What : Integer ;
BEGIN
If (Not SurAna) Then
  BEGIN
  If IdentPiece.DernChronoENonVentilable<>0 Then LeChrono:=IdentPiece.DernChronoENonVentilable
                                            Else LeChrono:=IdentPiece.DernChronoE ;
  What:=0 ;
  END Else
  BEGIN
  If IdentPiece.DernChronoEcrAna<>0 Then LeChrono:=IdentPiece.DernChronoEcrAna Else Exit ;
  What:=1 ; If SurODA Then What:=0 ;
  END ;
STDEBIT:='' ; STCREDIT:='' ; STDEBITDEV:='' ; STCREDITDEV:='' ; STDEBITEURO:='' ; STCREDITEURO:='' ;
DebitP:=0 ; CreditP:=0 ; DebitD:=0 ; CreditD:=0 ; DebitE:=0 ; CreditE:=0 ;
StSQL:='SELECT IE_DEBIT,IE_CREDIT,IE_DEBITDEV,IE_CREDITDEV,IE_DEBITEURO,IE_CREDITEURO FROM IMPECR WHERE IE_CHRONO='+IntToStr(LeChrono) ;
Q:=OpenSQL(StSQL,TRUE) ;
If Not Q.Eof Then
  BEGIN
  DebitPLu:=Q.Fields[0].AsFloat ; CreditPLu:=Q.Fields[1].AsFloat ;
  DebitDLu:=Q.Fields[2].AsFloat ; CreditELu:=Q.Fields[3].AsFloat ;
  DebitELu:=Q.Fields[4].AsFloat ; CreditELu:=Q.Fields[5].AsFloat ;
  Ferme(Q) ;
  Case QuiOppose Of
    'D' : BEGIN
          AjusteSQLMontant(What,V_PGI.OkDecV,'IE_DEBIT','IE_CREDIT',DeltaOppose,DebitPLu,CreditPLu,StDebit,StCredit) ;
          AjusteSQLMontant(What,V_PGI.OkDecE,'IE_DEBITEURO','IE_CREDITEURO',DeltaOppose2,DebitELu,CreditELu,StDebitEuro,StCreditEuro) ;
          END ;
    'P' : BEGIN
          AjusteSQLMontant(What,V_PGI.OkDecV,'IE_DEBIT','IE_CREDIT',DeltaOppose,DebitPLu,CreditPLu,StDebit,StCredit) ;
          If StDebit<>'' Then StDebitDev:='IE_DEBITDEV='+StrfPoint(DebitPLu) ;
          If StCredit<>'' Then StCreditDev:='IE_CREDITDEV='+StrfPoint(CreditPLu) ;
          END ;
    'E' : BEGIN
          AjusteSQLMontant(What,V_PGI.OkDecE,'IE_DEBITEURO','IE_CREDITEURO',DeltaOppose,DebitELu,CreditELu,StDebitEuro,StCreditEuro) ;
          END ;
(*
    'P' : BEGIN
          If EnEuro Then
            BEGIN
            AjusteSQLMontant(What,V_PGI.OkDecV,'IE_DEBIT','IE_CREDIT',DeltaOppose,DebitPLu,CreditPLu,StDebit,StCredit) ;
            If StDebit<>'' Then StDebitDev:='IE_DEBITDEV='+StrfPoint(DebitPLu) ;
            If StCredit<>'' Then StCreditDev:='IE_CREDITDEV='+StrfPoint(CreditPLu) ;
            END Else
            BEGIN
            AjusteSQLMontant(What,V_PGI.OkDecE,'IE_DEBITEURO','IE_CREDITEURO',DeltaOppose,DebitELu,CreditELu,StDebitEuro,StCreditEuro) ;
            END ;
          END ;
    'E' : BEGIN
          If EnEuro Then AjusteSQLMontant(What,V_PGI.OkDecE,'IE_DEBITEURO','IE_CREDITEURO',DeltaOppose,DebitELu,CreditELu,StDebitEuro,StCreditEuro)
                    Else BEGIN
                         AjusteSQLMontant(What,V_PGI.OkDecV,'IE_DEBIT','IE_CREDIT',DeltaOppose,DebitELu,CreditELu,StDebit,StCredit) ;
                         If StDebit<>'' Then StDebitDev:='IE_DEBITDEV='+StrfPoint(DebitELu) ;
                         If StCredit<>'' Then StCreditDev:='IE_CREDITDEV='+StrfPoint(CreditELu) ;
                         END ;
          END ;
*)
    END ;
  StSQL:='' ;
  If StDebit<>'' Then StSQL:=StSQL+StDebit+', ' ;
  If StCredit<>'' Then StSQL:=StSQL+StCredit+', ' ;
  If StDebitEuro<>'' Then StSQL:=StSQL+StDebitEuro+', ' ;
  If StCreditEuro<>'' Then StSQL:=StSQL+StCreditEuro+', ' ;
  If StDebitDev<>'' Then StSQL:=StSQL+StDebitDev+', ' ;
  If StCreditDev<>'' Then StSQL:=StSQL+StCreditDev+', ' ;
  If StSQL<>'' Then
    BEGIN
    StSQL:=Copy(StSQL,1,Length(StSQL)-2) ;
    StSQL:='UPDATE IMPECR SET '+StSQL+' WHERE IE_CHRONO='+IntToStr(LeChrono) ;
    ExecuteSQL(StSQL) ;
    END ;
  END Else
  BEGIN
  Ferme(Q) ;
  END ;
END ;

Function EcartDeConversionSurPiecePrecedente(Var IdentPiece : TIdentPiece ;
                                             Var InfoImp : TInfoImport ; Var Mvt : TFMvtImport) : Boolean ;
Var Cpt : String ;
    CptLu : TCptLu ;
    DSaisi,CSaisi,DOppose,COppose,DDevise,CDevise : Double ;
    DeltaSaisi,DeltaOppose,DeltaOppose2 : Double ;
    DecimSaisi,DecimOppose : Integer ;
    Pb : Boolean ;
    QuiOppose : Char ;
    EnEuro : Boolean ;
    UpdateImpEcr : Boolean ;
    TotEcrDSaisi,TotEcrCSaisi,TotEcrDOppose,TotEcrCOppose,TotEcrDDevise,TotEcrCDevise : Double ;
    SurODA : Boolean ;
BEGIN
Result:=FALSE ; UpdateImpEcr:=FALSE ;
//If IdentPiece.Date<V_PGI.DateDebutEuro Then Exit ;
EnEuro:=EuroOK And VH^.TenueEuro ;
Cpt:=Trim(IdentPiece.JalP) ; Fillchar(CptLu,SizeOf(CptLu),#0) ; CptLu.Cpt:=Cpt ; SurODA:=FALSE ;
If ChercheCptLu(InfoImp.LJalLu,CptLu) Then SurODA:=(CptLu.Nature='ODA') ;
If SurODA Then RecupMontant(1,EnEuro,DSaisi,CSaisi,DOppose,COppose,DDevise,CDevise,
                            IdentPiece,QuiOppose,DecimSaisi,DecimOppose,
                            TotEcrDSaisi,TotEcrCSaisi,TotEcrDOppose,TotEcrCOppose,TotEcrDDevise,TotEcrCDevise)
          Else RecupMontant(0,EnEuro,DSaisi,CSaisi,DOppose,COppose,DDevise,CDevise,
                            IdentPiece,QuiOppose,DecimSaisi,DecimOppose,
                            TotEcrDSaisi,TotEcrCSaisi,TotEcrDOppose,TotEcrCOppose,TotEcrDDevise,TotEcrCDevise) ;
DeltaSaisi:=0 ; DeltaOppose:=0 ; DeltaOppose2:=0 ;
Case QuiOppose Of
  'D' : BEGIN
        DeltaSaisi:=Arrondi(DDevise-CDevise,IdentPiece.DecimDev) ;
        DeltaOppose:=Arrondi(DSaisi-CSaisi,V_PGI.OkDecV) ;
        DeltaOppose2:=Arrondi(DOppose-COppose,V_PGI.OkDecE) ;
        END ;
  'E' : BEGIN
        DeltaSaisi:=Arrondi(DSaisi-CSaisi,DecimSaisi) ;
        DeltaOppose:=Arrondi(DOppose-COppose,DecimOppose) ;
        END ;
  'P' : BEGIN
        DeltaSaisi:=Arrondi(DSaisi-CSaisi,DecimSaisi) ;
        DeltaOppose:=Arrondi(DOppose-COppose,DecimOppose) ;
        END ;
  END ;
Pb:=((QuiOppose in ['P','E']) And (DeltaOppose<>0) And (DeltaSaisi=0)) Or
    ((QuiOppose in ['D']) And (DeltaSaisi=0) And ((DeltaOppose<>0) Or (DeltaOppose2<>0))) ;
UpdateImpecr:=Pb And ((IdentPiece.DateP<V_PGI.DateDebutEuro) Or (QuiOppose='D') Or SurODA) ;
If UpdateImpEcr Then
  BEGIN
  If SurODA Then UpdatePourEquilibre(TRUE,TRUE,EnEuro,QuiOppose,DeltaOppose,DeltaOppose2,IdentPiece)
            Else UpdatePourEquilibre(FALSE,FALSE,EnEuro,QuiOppose,DeltaOppose,DeltaOppose2,IdentPiece) ;
  END Else
  BEGIN
  If Pb Then
    BEGIN
    Result:=TRUE ;
    MVT.IE_ELEMENTARECUPERER:=TRUE ;
    Mvt.IE_DEBIT:=0 ; Mvt.IE_CREDIT:=0 ; Mvt.IE_DEBITEURO:=0 ; MVT.IE_CREDITEURO:=0 ;
    Mvt.IE_DEBITDEV:=0 ; Mvt.IE_CREDITDEV:=0 ;
    Case QuiOppose Of
      'D' : BEGIN
            Result:=FALSE ;
            MVT.IE_ELEMENTARECUPERER:=FALSE ;
            (*
            If DeltaOppose<>0 Then
              BEGIN
              If DeltaOppose>0 Then Mvt.IE_CREDIT:=DeltaOppose Else Mvt.IE_DEBIT:=Abs(DeltaOppose) ;
              Mvt.IE_GENERAL:=VH^.EccEuroDebit ;
              If DeltaOppose>0 Then Mvt.IE_GENERAL:=VH^.EccEuroCredit ;
              END ;
            If DeltaOppose2<>0 Then
              BEGIN
              If DeltaOppose2>0 Then Mvt.IE_CREDITEURO:=DeltaOppose2 Else Mvt.IE_DEBITEURO:=Abs(DeltaOppose2) ;
              Mvt.IE_GENERAL:=VH^.EccEuroDebit ;
              If DeltaOppose2>0 Then Mvt.IE_GENERAL:=VH^.EccEuroCredit ;
              END ;
            *)
            END ;
      'P' : BEGIN
            If DeltaOppose>0 Then Mvt.IE_CREDIT:=DeltaOppose Else Mvt.IE_DEBIT:=Abs(DeltaOppose) ;
            Mvt.IE_GENERAL:=VH^.EccEuroDebit ;
            If DeltaOppose>0 Then Mvt.IE_GENERAL:=VH^.EccEuroCredit ;
            END ;
      'E' : BEGIN
            If DeltaOppose>0 Then Mvt.IE_CREDITEURO:=DeltaOppose Else Mvt.IE_DEBITEURO:=Abs(DeltaOppose) ;
            Mvt.IE_GENERAL:=VH^.EccEuroDebit ;
            If DeltaOppose>0 Then Mvt.IE_GENERAL:=VH^.EccEuroCredit ;
            END ;
      END ;
    Mvt.IE_LIBELLE:='Ecart de conversion' ; Mvt.IE_AUXILIAIRE:='' ;
    MVT.IE_SECTION:='' ; MVT.IE_AXE:='' ; MVT.IE_ECHE:='-' ;
    MVT.IE_ANA:='-' ; Mvt.IE_NUMECHE:=0 ; Mvt.IE_ETATLETTRAGE:='RI' ;
    MVT.IE_DATEECHEANCE:=iDate1900 ;MVT.IE_DATEPAQUETMAX:=iDate1900 ;
    MVT.IE_DATEPAQUETMIN:=iDate1900 ; MVT.IE_DATEPOINTAGE:=iDate1900 ; MVT.IE_DATEREFEXTERNE:=iDate1900 ;
    MVT.IE_DATERELANCE:=iDate1900 ; MVT.IE_DATETAUXDEV:=iDate1900 ; MVT.IE_DATEVALEUR:=iDate1900 ;
    MVT.IE_DATECREATION:=Date ; MVT.IE_ORIGINEPAIEMENT:=iDate1900 ; MVT.IE_LIBREDATE:=iDate1900 ;
    MVT.IE_ENCAISSEMENT:='RIE';MVT.IE_CONTROLE:='-' ;MVT.IE_LETTRAGEDEV:='-' ; MVT.IE_OKCONTROLE:='X' ;
    MVT.IE_SELECTED:='X' ; MVT.IE_TVAENCAISSEMENT:='-' ; MVT.IE_TYPEANALYTIQUE:='-' ;
    MVT.IE_VALIDE:='-' ; MVT.IE_ANA:='-' ;MVT.IE_INTEGRE:='-' ;
    MVT.IE_LIBREBOOL0:='-' ; MVT.IE_LIBREBOOL1:='-' ;
    Mvt.IE_TYPEECR:='E' ;
    END ;
  END ;
END ;

Function VerifEquilibrageAnalytique(Var St,OldSt : String ; Var IdentPiece : TIdentPiece ;
                                    Var InfoImp : TInfoImport ; Var Mvt : TFMvtImport) : Boolean ;
Var FmtFic : Integer ;
    OldEnr,Enr : TStLue ;
    Cpt : String ;
    CptLu : TCptLu ;
    AVerifier : Boolean ;
    DSaisi,CSaisi,DOppose,COppose,DDevise,CDevise : Double ;
    DeltaSaisi,DeltaOppose,DeltaOppose2 : Double ;
    DecimSaisi,DecimOppose : Integer ;
    TotEcrDSaisi,TotEcrCSaisi,TotEcrDOppose,TotEcrCOppose,TotEcrDDevise,TotEcrCDevise : Double ;
    UpdateImpEcr : Boolean ;
    QuiOppose : Char ;
    EnEuro : Boolean ;
BEGIN
//Exit ;
EnEuro:=EuroOK And VH^.TenueEuro ;
Result:=FALSE ; If OldSt='' Then Exit ;
Fillchar(Enr,SizeOf(Enr),#0) ; Fillchar(OldEnr,SizeOf(OldEnr),#0) ;
FmtFic:=0 ;
if InfoImp.Format='SAA' then FmtFic:=0 else if InfoImp.Format='SN2' then FmtFic:=0 else
if InfoImp.Format='HLI' then FmtFic:=1 else if InfoImp.Format='HAL' then FmtFic:=2 else
if InfoImp.Format='CGN' then FmtFic:=3 else if InfoImp.Format='CGE' then FmtFic:=4 Else Exit ;
AlimEnr(OldSt,OldEnr,FmtFic) ; If St<>'' Then AlimEnr(St,Enr,FmtFic) ;
If OldEnr.TC='O' Then Exit ;
If OldEnr.TC='L' Then Exit ;
Cpt:=Trim(OldEnr.Jal) ; Fillchar(CptLu,SizeOf(CptLu),#0) ; CptLu.Cpt:=Cpt ;
If ChercheCptLu(InfoImp.LJalLu,CptLu) Then
  BEGIN
  If CptLu.Nature='ODA' Then Exit ;
  END Else Exit ;
Cpt:=BourreOuTronque(Trim(OldEnr.General),fbGene) ;
Fillchar(CptLu,SizeOf(CptLu),#0) ; CptLu.Cpt:=Cpt ;
If ChercheCptLu(InfoImp.LGenLu,CptLu) Then
  BEGIN
  AVerifier:=CptLu.Ventilable And (Mvt.IE_TYPEECR='A') And ((St='') Or (Enr.TC<>'A'))  ;
  If AVerifier Then
    BEGIN
    RecupMontant(1,EnEuro,DSaisi,CSaisi,DOppose,COppose,DDevise,CDevise,IdentPiece,QuiOppose,DecimSaisi,DecimOppose,
                 TotEcrDSaisi,TotEcrCSaisi,TotEcrDOppose,TotEcrCOppose,TotEcrDDevise,TotEcrCDevise) ;
    DeltaSaisi:=0 ; DeltaOppose:=0 ; DeltaOppose2:=0 ;
    Case QuiOppose Of
      'D' : BEGIN
            DeltaSaisi:=Arrondi((TotEcrDDevise+TotEcrCDevise)-(DDevise+CDevise),IdentPiece.DecimDev) ;
            DeltaOppose:=Arrondi((TotEcrDSaisi+TotEcrCSaisi)-(DSaisi+CSaisi),V_PGI.OkDecV) ;
            DeltaOppose2:=Arrondi((TotEcrDOppose+TotEcrCOppose)-(DOppose+COppose),V_PGI.OkDecE) ;
            END ;
      'E' : BEGIN
            DeltaSaisi:=Arrondi((TotEcrDSaisi+TotEcrCSaisi)-(DSaisi+CSaisi),DecimSaisi) ;
            DeltaOppose:=Arrondi((TotEcrDOppose+TotEcrCOppose)-(DOppose+COppose),DecimOppose) ;
            END ;
      'P' : BEGIN
            DeltaSaisi:=Arrondi((TotEcrDSaisi+TotEcrCSaisi)-(DSaisi+CSaisi),DecimSaisi) ;
            DeltaOppose:=Arrondi((TotEcrDOppose+TotEcrCOppose)-(DOppose+COppose),DecimOppose) ;
            END ;
      END ;
    UpdateImpEcr:=((QuiOppose in ['P','E']) And (DeltaOppose<>0) And (DeltaSaisi=0)) Or
                  ((QuiOppose in ['D']) And (DeltaSaisi=0) And ((DeltaOppose<>0) Or (DeltaOppose2<>0))) ;
    If UpdateImpEcr Then UpdatePourEquilibre(TRUE,FALSE,EnEuro,QuiOppose,DeltaOppose,DeltaOppose2,IdentPiece) ;
    END ;
  END ;
END ;

Procedure ReAlimCodeMontant(St : String ; Var InfoImp : TInfoImport ; Var IdentPiece : TIdentPiece) ;
Var CodeMontant : String ;
    TauxDev : Double ;
    Decal : Integer ;
    DateC : TDateTime ;
    EnEuro : Boolean ;
BEGIN
EnEuro:=EuroOK And VH^.TenueEuro ;
CodeMontant:='F--' ; TauxDev:=0 ; Decal:=0 ; If Not VH^.ImportRL Then Decal:=8 ; DateC:=0 ;
If (InfoImp.Format='CGN') Or (InfoImp.Format='CGE') Then
  BEGIN
  CodeMontant:=Trim(Copy(St,173,3)) ; If (CodeMontant='') Or (CodeMontant='---') Then CodeMontant:='F--' ;
  if Trim(Copy(St,163,10))<>'' then TauxDev:=Valeur(StPoint(Copy(St,163,10)));
  DateC:=Format_Date_HAL(Copy(St,4,8)) ;
  END Else
If InfoImp.Format='HAL' Then
  BEGIN
  if Trim(Copy(St,305+Decal,20))<>'' then TauxDev:=Valeur(StPoint(Copy(St,305+Decal,20)));
  If EnEuro Then CodeMontant:='E--' Else CodeMontant:='F--' ;
  if Trim(Copy(St,262+Decal,3))<>'' then IF Trim(Copy(St,262+Decal,3))<>V_PGI.DevisePivot Then CodeMontant:='D--' ;
  DateC:=Format_Date_HAL(Copy(St,4,8)) ;
  END Else
If InfoImp.Format='HLI'  Then
  BEGIN
  If EnEuro Then CodeMontant:='E--' Else CodeMontant:='F--' ;
  DateC:=Format_Date_HAL(Copy(St,4,8)) ;
  END Else
  BEGIN
  If EnEuro Then CodeMontant:='E--' Else CodeMontant:='F--' ;
  DateC:=Format_Date(Copy(St,4,6)) ;
  END ;
IdentPiece.CodeMontant:=CodeMontant ;
If TauxDev<>0 Then
  BEGIN
  If DateC>=V_PGI.DateDebutEuro Then TauxDev:=TauxDev*V_PGI.TauxEuro ;
  END ;
IdentPiece.TauxDev:=TauxDev ;
END ;

Function MontantNul(Var St : String ; Var InfoImp : TInfoImport) : Boolean ;
Var FmtFic,Decal : Integer ;
    Montant : Double ;
BEGIN
Result:=FALSE ;
if InfoImp.Format='SAA' then FmtFic:=0 else if InfoImp.Format='SN2' then FmtFic:=0 else
if InfoImp.Format='HLI' then FmtFic:=1 else if InfoImp.Format='HAL' then FmtFic:=2 else
if InfoImp.Format='CGN' then FmtFic:=3 else if InfoImp.Format='CGE' then FmtFic:=4 Else Exit ;
Decal:=0 ; If Not VH^.ImportRL Then Decal:=8 ;
Montant:=0 ;
Case FmtFic Of
  0   : Montant:=Valeur(StPoint(Copy(St,85,20))) ;
  1,2 : Montant:=Valeur(StPoint(Copy(St,89+Decal,20))) ;
  3,4 : Montant:=Valeur(Trim(StPoint(Copy(St,131,20)))) ;
  END ;
Result:=Arrondi(Montant,5)=0 ;
END ;

Function ImportLigne(St : String ; Var OldSt : String ; QAjoute : TQuery ; QFiche : TQFiche ; QDoublon : TQDoublon ;
                     Var InfoImp : TInfoImport ; Var IdentPiece,OldIdentPiece : TIdentPiece ; Var PremFois : Boolean ; FinDeTraitement : Boolean)  : Boolean ;
var EnRupture,DetecteRupture : Boolean ;
    AnalSurUneLigne : Boolean ;
    AnalAInserer : Boolean ;
    EcartDeConvertionAGenerer : Boolean ;
    SauveSt : String ;
    OkRetoucheMontant : Boolean ;
 Label 0 ;
BEGIN
Result:=FALSE ; AnalSurUneLigne:=FALSE ; AnalAInserer:=FALSE ;
MvtImport^.IE_ElementARecuperer:=FALSE ;
if (Pos(SepLigneIE,St)=1) And (St[2]<>SepLigneIE ) then Exit ;
If EstUneLigneCpt(St) Then
  BEGIN
  Result:=TestBreak ; If Result Then Exit ;
  TraiteImportCompte(St,InfoImp,QFiche) ; Exit ;
  END ;
If MontantNul(St,InfoImp) Then Exit ;
EcartDeConvertionAGenerer:=FALSE ;
DetecteRupture:=FALSE ;
If Not PremFois Then
  BEGIN
  DetecteRupture:=DetecteUneRupture(St,IdentPiece,InfoImp) ;
  If FinDeTraitement Then BEGIN OldSt:=St ; St:='' ; END ;
  VerifEquilibrageAnalytique(St,OldSt,IdentPiece,InfoImp,MvtImport^) ;
  If FinDeTraitement Then BEGIN St:=OldSt ; OldSt:='' ; END ;
  END ;
if (InfoImp.Format='RAP') or (InfoImp.Format='CRA') or (Trim(InfoImp.Format)='MP')  Then Else
  BEGIN
  If FinDeTraitement Or DetecteRupture Then
    BEGIN
    EcartDeConvertionAGenerer:=EcartDeConversionSurPiecePrecedente(IdentPiece,InfoImp,MvtImport^) ;
    If FinDeTraitement And (Not EcartDeConvertionAGenerer) Then Exit ;
    END ;
  END ;
SauveSt:=St ;
If AnalytiqueAbsenteSurLignePrecedente(St,OldSt,InfoImp) Then
  BEGIN
  AnalAInserer:=TRUE ;
  St:=OldSt ;
  END ;
0:
If (PremFois Or DetecteRupture) And (Not EcartDeConvertionAGenerer) Then ReAlimCodeMontant(St,InfoImp,IdentPiece) ;
PremFois:=FALSE ;
if (InfoImp.Format='HAL') or (InfoImp.Format='HLI') then ImportLigneHal(St,InfoImp,MvtImport^) Else
  If (InfoImp.Format='CGN') Or (InfoImp.Format='CGE') Then ImportLigneCegid(St,InfoImp,MvtImport^,AnalSurUneLigne,IdentPiece) Else
     ImportLigneAutre(St,InfoImp,MvtImport^) ;
InitAxe(St,InfoImp.Format,MvtImport^) ;
If AlimCptMemoire(QFiche,MvtImport^,InfoImp)=1 Then Exit ; // Ligne analytique sur compte non ventilable : on ignore
OkRetoucheMontant:=True ; If (EcartDeConvertionAGenerer) Then OkRetoucheMontant:=FALSE ;
if (InfoImp.Format='HAL') or (InfoImp.Format='HLI') then RetoucheEnrHal(St,IdentPiece,InfoImp,MvtImport^,OkRetoucheMontant) Else
  If (InfoImp.Format='CGN') Or (InfoImp.Format='CGE') Then RetoucheEnrCegid(St,IdentPiece,InfoImp,MvtImport^,OkRetoucheMontant) Else
     RetoucheEnrAutre(St,IdentPiece,InfoImp,MvtImport^,OkRetoucheMontant) ;
{ A FAIRE : }
if (InfoImp.Format='RAP') and (Copy(St,133,1)='0') then Exit ; // IE_REFPOINTAGE:=Copy(St,133,1) ;
RetoucheODANAL(MvtImport^) ;
AlimIdentPiece(IdentPiece,MvtImport^) ;
EnRupture:=TestRupturePiece(InfoImp.Format,TRUE,IdentPiece,OldIdentPiece) ;
If EnRupture Then
   BEGIN
   Inc(InfoImp.NbPiece) ; Result:=TestBreak ;
   IdentPiece.Doublon:=FALSE ;
   If InfoImp.CtrlDB Then If ChercheDoublon(QDoublon,InfoImp,MvtImport^) Then IdentPiece.Doublon:=TRUE ;
   If Result Then Exit ;
   END ;
AlimNumeros(EnRupture,IdentPiece,InfoImp,MvtImport^) ;
RetoucheFinal(IdentPiece,MvtImport^) ;
OldSt:=St ;
AjouteMvt(Qajoute,InfoImp) ;
If AnalSurUneLigne Then
   BEGIN
   PrepareQuery(QAjoute) ;
   Goto 0 ;
   END ;
If AnalAInserer Then
   BEGIN
   AnalAInserer:=FALSE ; St:=SauveSt ; If Trim(St)='' Then Exit ;
   PrepareQuery(QAjoute) ;
   Goto 0 ;
   END ;
If Not (FinDeTraitement) And EcartDeConvertionAGenerer Then
   BEGIN
   EcartDeConvertionAGenerer:=FALSE ; If Trim(St)='' Then Exit ;
   MvtImport^.IE_ElementARecuperer:=FALSE ;
   PrepareQuery(QAjoute) ;
   Goto 0 ;
   END ;
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

function ImporteFormatParam(Var InfoImp : TInfoImport) : boolean ;
Var Debut   : Boolean ;
    Q : TQuery ;
    N,NumLigne : Integer ;
    OkRupt : boolean ;
    Fichier   : TextFile ;
    Entete  : TFmtEntete ;
    Detail  : TTabFmtDetail ;
BEGIN
Result:=False ;
if not ChargeFormat(Fichier,InfoImp.NomFic,InfoImp.Lequel,'X',InfoImp.Format,Entete,Detail,Debut) then Exit ;
Q:=OpenSQL('SELECT * FROM IMPECR',False) ;
N:=1 ; NumLigne:=0 ;
InitMove(NbLig,'') ;
While not Eof(Fichier) do
  BEGIN
  Q.Insert ; InitNew(Q) ;
  Q.FindField('IE_CHRONO').AsInteger:=N ;
  Q.FindField('IE_ETABLISSEMENT').AsString:=VH^.EtablisDefaut ;
  Q.FindField('IE_DEVISE').AsString:=V_PGI.DevisePivot ;
  Q.FindField('IE_ECRANOUVEAU').AsString:='N' ;
  Q.FindField('IE_QUALIFPIECE').AsString:='N' ;
  Q.FindField('IE_ETATLETTRAGE').AsString:='RI' ;
  if (InfoImp.Lequel='FOD') then
     BEGIN
     Q.FindField('IE_TYPEECR').AsString:='A' ;
     Q.FindField('IE_TYPEANALYTIQUE').AsString:='X' ;
     END else
     BEGIN
     Q.FindField('IE_TYPEECR').AsString:='E' ;
     END ;
  if (InfoImp.Lequel<>'FBE') and (InfoImp.Lequel<>'FOD') and (Q.FindField('IE_NUMLIGNE').AsInteger=0) then BEGIN Inc(NumLigne) ; Q.FindField('IE_NUMLIGNE').AsInteger:=NumLigne ; END ;
  LireFormat(Fichier,Entete,Detail,Debut,Q) ;
  (*
  Jal:=Q.FindField('IE_JOURNAL').AsString ;
  DateC:=Q.FindField('IE_DATECOMPTABLE').AsDateTime ;
  NumP:=Q.FindField('IE_NUMPIECE').AsInteger ;
  QualP:=Q.FindField('IE_QUALIFPIECE').AsString ;
  NatP:=Q.FindField('IE_NATUREPIECE').AsString ;
  if NatJal(Jal)='ANO' then Q.FindField('IE_ECRANOUVEAU').AsString:='OAN' else
    if NatJal(Jal)='CLO' then Q.FindField('IE_ECRANOUVEAU').AsString:='C' ;
  *)
  If Imp_Methode1 Then Application.ProcessMessages ;
  if (InfoImp.Lequel='FBE') (*and (OldNumP<>NumP)*) then
    BEGIN
//    OldNumP:=NumP ;
    if AnnuleImport then break ;
    Inc(InfoImp.NbPiece) ; NumLigne:=0 ;
    END else
    BEGIN
//    OkRupt:=TestBreak(InfoImp) ;
    if OkRupt then Break ;
    END ;
  Q.FindField('IE_DEBITEURO').AsFloat:=PivotToEuro(Q.FindField('IE_DEBIT').AsFloat) ;
  Q.FindField('IE_CREDITEURO').AsFloat:=PivotToEuro(Q.FindField('IE_CREDIT').AsFloat) ;
  if (Q.FindField('IE_DEVISE').AsString=V_PGI.DevisePivot) then
   BEGIN
   Q.FindField('IE_TAUXDEV').AsFloat:=1 ;
   Q.FindField('IE_DATETAUXDEV').AsDateTime:=iDate1900 ;
   Q.FindField('IE_DEBITDEV').AsFloat:=Q.FindField('IE_DEBIT').AsFloat ;
   Q.FindField('IE_CREDITDEV').AsFloat:=Q.FindField('IE_CREDIT').AsFloat ;
   END ;
  if (Q.FindField('IE_TYPEECR').AsString<>'A') then
     BEGIN
     InfoImp.TotDeb:=InfoImp.TotDeb+Q.FindField('IE_DEBIT').AsFloat ;
     InfoImp.TotCred:=InfoImp.TotCred+Q.FindField('IE_CREDIT').AsFloat ;
     END ;
// Même n° de ligne générale et ventil.
  if (InfoImp.Lequel<>'FBE') and (InfoImp.Lequel<>'FOD') and (Q.FindField('IE_NUMLIGNE').AsInteger<NumLigne) then Dec(NumLigne) ;
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
   Q.SQL.Add('IE_COUVERTURE,IE_COUVERTUREDEV,IE_DEBITDEV,IE_DEBITEURO, ') ;

   Q.SQL.Add('IE_LIBRETEXTE0,IE_LIBRETEXTE1,IE_LIBRETEXTE2,IE_LIBRETEXTE3,IE_LIBRETEXTE4,') ;
   Q.SQL.Add('IE_LIBRETEXTE5,IE_LIBRETEXTE6,IE_LIBRETEXTE7,IE_LIBRETEXTE8,IE_LIBRETEXTE9,') ;
   Q.SQL.Add('IE_TABLE0,IE_TABLE1,IE_TABLE2,IE_TABLE3,IE_LIBREBOOL0,IE_LIBREBOOL1,') ;
   Q.SQL.Add('IE_LIBREMONTANT0,IE_LIBREMONTANT1,IE_LIBREMONTANT2,IE_LIBREMONTANT3,IE_LIBREDATE,IE_SAISIEEURO,IE_COTATION) ') ;



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
   Q.SQL.Add(':IE_COUVERTURE,:IE_COUVERTUREDEV,:IE_DEBITDEV,:IE_DEBITEURO, ') ;

   Q.SQL.Add(':IE_LIBRETEXTE0,:IE_LIBRETEXTE1,:IE_LIBRETEXTE2,:IE_LIBRETEXTE3,:IE_LIBRETEXTE4,') ;
   Q.SQL.Add(':IE_LIBRETEXTE5,:IE_LIBRETEXTE6,:IE_LIBRETEXTE7,:IE_LIBRETEXTE8,:IE_LIBRETEXTE9,') ;
   Q.SQL.Add(':IE_TABLE0,:IE_TABLE1,:IE_TABLE2,:IE_TABLE3,:IE_LIBREBOOL0,:IE_LIBREBOOL1,') ;
   Q.SQL.Add(':IE_LIBREMONTANT0,:IE_LIBREMONTANT1,:IE_LIBREMONTANT2,:IE_LIBREMONTANT3,:IE_LIBREDATE,:IE_SAISIEEURO,:IE_COTATION)') ;

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

Procedure InitRequete(Var Q : TQuery ; What : Integer) ;
Var St : String ;
BEGIN
Q:=Nil ;
Case What Of
  0 : St:='SELECT G_NATUREGENE,G_REGIMETVA,G_DATEDERNMVT,G_NUMDERNMVT, '+
          'G_LIGNEDERNMVT,G_DEBITDERNMVT,G_CREDITDERNMVT,G_TOTALDEBIT,G_TOTALCREDIT, '+
          'G_TOTDEBE,G_TOTCREE,G_TOTDEBS,G_TOTCRES,G_TOTDEBANO,G_TOTCREANO,G_VENTILABLE, '+
          'G_VENTILABLE1,G_VENTILABLE2,G_VENTILABLE3,G_VENTILABLE4,G_VENTILABLE5, '+
          'G_LETTRABLE,G_POINTABLE,G_TVA,G_TPF FROM GENERAUX '+
          'WHERE G_GENERAL=:CPTE' ;
  1 : St:='SELECT T_NATUREAUXI,T_REGIMETVA,T_DATEDERNMVT,T_NUMDERNMVT,T_COLLECTIF, '+
          'T_LIGNEDERNMVT,T_DEBITDERNMVT,T_CREDITDERNMVT,T_TOTALDEBIT,T_TOTALCREDIT, '+
          'T_TOTDEBE,T_TOTCREE,T_TOTDEBS,T_TOTCRES,T_TOTDEBANO,T_TOTCREANO,T_LETTRABLE FROM TIERS '+
          'WHERE T_AUXILIAIRE=:CPTE' ;
  2 : St:='SELECT S_SECTION,S_AXE,S_DATEDERNMVT, '+
          'S_NUMDERNMVT,S_LIGNEDERNMVT,S_DEBITDERNMVT,S_CREDITDERNMVT,S_TOTALDEBIT,S_TOTALCREDIT, '+
          'S_TOTDEBE,S_TOTCREE,S_TOTDEBS,S_TOTCRES,S_TOTDEBANO,S_TOTCREANO FROM SECTION '+
          'WHERE S_SECTION=:CPTE AND S_AXE=:AXE' ;
  3 : St:='SELECT J_JOURNAL,J_NATUREJAL,J_COMPTEURNORMAL,J_COMPTEURSIMUL,J_DATEDERNMVT,'+
          'J_NUMDERNMVT,J_DEBITDERNMVT,J_CREDITDERNMVT,J_TOTALDEBIT,J_TOTALCREDIT,'+
          'J_TOTDEBE,J_TOTCREE,J_TOTDEBS,J_TOTCRES FROM JOURNAL '+
          'WHERE J_JOURNAL=:CPTE' ;
  4 : St:='SELECT BJ_BUDJAL,BJ_NATJAL,BJ_COMPTEURNORMAL,BJ_COMPTEURSIMUL '+
          'FROM BUDJAL ORDER BY BJ_BUDJAL' ;
  END ;
Q:=PrepareSQL(St,TRUE) ;
END ;

Procedure InitRequeteDoublon(Var Q : TQuery ; What : Integer) ;
Var St : String ;
BEGIN
Q:=Nil ;
Case What Of
  0 : St:='SELECT E_JOURNAL, E_EXERCICE, E_DATECOMPTABLE, E_NUMEROPIECE, E_NUMLIGNE, E_NUMECHE, E_QUALIFPIECE, E_REFINTERNE '+
          'FROM ECRITURE WHERE E_JOURNAL=:JAL AND E_EXERCICE=:EXO AND E_DATECOMPTABLE=:DATEC AND E_GENERAL=:GEN AND '+
          'E_AUXILIAIRE=:AUX AND E_REFINTERNE=:REF AND E_LIBELLE=:LIB AND E_DEVISE=:DEV AND E_DEBITDEV=:DEB AND '+
          'E_CREDITDEV=:CRE' ;
  1 : St:='SELECT Y_JOURNAL, Y_EXERCICE, Y_DATECOMPTABLE, Y_NUMEROPIECE, Y_NUMLIGNE, Y_QUALIFPIECE, Y_REFINTERNE '+
          'FROM ANALYTIQ WHERE Y_JOURNAL=:JAL AND Y_EXERCICE=:EXO AND Y_DATECOMPTABLE=:DATEC AND Y_GENERAL=:GEN AND '+
          'Y_SECTION=:SECT AND Y_REFINTERNE=:REF AND Y_LIBELLE=:LIB AND Y_DEVISE=:DEV AND Y_DEBITDEV=:DEB AND '+
          'Y_CREDITDEV=:CRE AND Y_AXE=:AXE' ;
  END ;
Q:=PrepareSQL(St,TRUE) ;
END ;

Function ImporteEcr(Var StErr : String ; Var InfoImp : TInfoImport) : Boolean ;
var St,OldSt : String ;
    ExpParam,OkFic,PremFois : boolean ;
    QAjoute : TQuery ;
    IdentPiece,OldIdentPiece : TIdentPiece ;
    QFiche : TQFiche ;
    QDoublon : TQDoublon ;
    DateDeb,DateFin : TDateTime ;
    i : Integer ;
    Fichier   : TextFile ;
    ArretDemande : Boolean ;
BEGIN
InitMove(NbLig,'') ; MoveCur(False) ; StErr:='' ; Result:=TRUE ; ArretDemande:=FALSE ;
Fillchar(IdentPiece,SizeOf(IdentPiece),#0) ; Fillchar(OldIdentPiece,SizeOf(OldIdentPiece),#0) ;
If Imp_Methode1 And (Not QAJParam) Then BEGIN QAjoute:=TQuery.Create(Application) ; QAjoute.DataBaseName:='SOC' ; END
                Else QAjoute:=InitQAjoute ;
For i:=0 To 3 Do InitRequete(QFiche[i],i) ;
If InfoImp.CtrlDB Then For i:=0 To 1 Do InitRequeteDoublon(QDoublon[i],i) ;
try
  SourisSablier ;
  New(MvtImport) ;
  BEGINTRANS ;
  ExecuteSQL('DELETE FROM IMPECR') ;
  COMMITTRANS ;
  ExpParam:=(InfoImp.Lequel='FEC') and ((InfoImp.Format<>'SAA') and (InfoImp.Format<>'EDI')
                           and (InfoImp.Format<>'SN2')  and (InfoImp.Format<>'HAL')
                           and (InfoImp.Format<>'HLI')  and (InfoImp.Format<>'MP')
                           and (InfoImp.Format<>'CGE')  and (InfoImp.Format<>'CGN')
                           and (InfoImp.Format<>'RAP')) or ((InfoImp.Lequel<>'FEC') and (InfoImp.Lequel<>'FBA')) ;
  BEGINTRANS ;
  if ExpParam then
    BEGIN
    ImporteFormatParam(InfoImp) ;
    END else
    BEGIN
    OkFic:=TRUE ; PremFois:=TRUE ;
    if (InfoImp.Lequel='FBA') and (InfoImp.Format='EDI') then NbLig:=Nblig*2 ;
    if (InfoImp.Format='SN2') Or (InfoImp.Format='SAA') Then
       BEGIN // GP à prévoir sur HALLEY et HALLEY étendu
       OkFic:=SoldeAZero(InfoImp.Nomfic,StErr) ;
       If Not OkFic Then Result:=FALSE ;
       END ;
    AssignFile(Fichier,InfoImp.Nomfic) ; {$I-} Reset(Fichier) ; {$I+}
    If OkFic Then
       BEGIN
       Readln(Fichier,St) ; //Société
       if (InfoImp.Lequel='FBA') and (InfoImp.Format<>'EDI') then
         BEGIN
         DateDeb:=Format_DateEDI(Copy(st,77,8)) ;
         DateFin:=Format_DateEDI(Copy(st,85,8)) ;
         END ;
       OldSt:='' ;
       While (not EOF(Fichier)) And (Not ArretDemande) do
         BEGIN
         PrepareQuery(QAjoute) ;
         Readln(Fichier,St) ;
         if Trim(St)='' then Break ;
         if OkRupt or Erreur then Break ;
         if ChoixFmt.Ascii then St:=ASCII2ANSI(St) ;
         if (InfoImp.Lequel<>'FBA') then
           BEGIN
           (*
           if (Format='EDI') then ImportLigneEDI(St,QAjoute,TabJalSoc,InfoImp)
           *)
                             (*else*)
           ArretDemande:=ImportLigne(St,OldSt,QAjoute,QFiche,QDoublon,InfoImp,IdentPiece,OldIdentPiece,PremFois,FALSE) ;
           END else
           BEGIN
           (*
           if (Format<>'EDI') then ImportLigne(St,QAjoute,QFiche,TabJalSoc,InfoImp,IdentPiece,OldIdentPiece)
                              else ImportBalance(St,1,QAjoute,TabJalSoc,InfoImp) ;
           *)
           END ;
//         PremFois:=FALSE ;
         MoveCur(False) ;
         END ;
       St:='' ;
       If (Not PremFois) ANd (OldSt<>'') And AnalytiqueAbsenteSurLignePrecedente(St,OldSt,InfoImp) Then
         BEGIN
         St:=OldSt ; OldSt:='' ;
         PrepareQuery(QAjoute) ;
         ImportLigne(St,OldSt,QAjoute,QFiche,QDoublon,InfoImp,IdentPiece,OldIdentPiece,PremFois,TRUE) ;
         END ;
       If (Not PremFois) Then
         BEGIN
         PrepareQuery(QAjoute) ; St:=OldSt ;
         ImportLigne(St,OldSt,QAjoute,QFiche,QDoublon,InfoImp,IdentPiece,OldIdentPiece,PremFois,TRUE) ;
         END ;
       END ;
    END ;
  finally
    CloseFile(Fichier) ;
    Dispose(MvtImport) ;
    if Erreur then ROLLBACK else COMMITTRANS ;
    SourisNormale ;
  end ;

QAjoute.Close ;
If Imp_Methode1 Then QAjoute.Free Else Ferme(QAjoute) ;
FiniMove ;
//If Result Then VerifExisteCptLu(InfoImp,LgComptes) ;
For i:=0 To 3 Do Ferme(QFiche[i]) ;
If InfoImp.CtrlDB Then For i:=0 To 1 Do Ferme(QDoublon[i]) ;
END ;

function ImporteLesEcritures(Msg : THMsgBox ; Var InfoImp : TInfoImport) : boolean ;
Var OkOk : Boolean ;
    StErr : String ;
    FmtFic : Integer ;
BEGIN
Result:=False ; OkOk:=TRUE ; StErr:='' ;
MsgBox:=Msg ;
// Pas de format SAARI pour la balance
if (InfoImp.Lequel='FBA') and (InfoImp.Format='SAA') then InfoImp.Format:='HAL' ;
InfoImp.Prefixe:='E' ; InfoImp.Table:='ECRITURE' ;
if InfoImp.Lequel='FBE' then BEGIN InfoImp.Prefixe:='BE' ; InfoImp.Table:='BUDECR' ; END else
  if InfoImp.Lequel='FOD' then BEGIN InfoImp.Prefixe:='Y' ; InfoImp.Table:='ANALYTIQ' ;END ;
if (InfoImp.NomFic='') and (Msg<>nil) then BEGIN AnnuleImport:=True ; MsgBox.Execute(40,'','') ; Exit ; END ;
NbLig:=0 ;
Case GoodSoc(InfoImp) of
  2 : if (Msg<>nil) then BEGIN MsgBox.Execute(41,'','') ; Exit ; END
                    else BEGIN CompteRenduBatch(2,InfoImp) ; Exit ; END ;
  0 : BEGIN CompteRenduBatch(0,InfoImp) ; Exit ; END ;
  END ;
(*
If Imp_Methode1 Then BEGIN QAjoute:=TQuery.Create(Application) ; QAjoute.DataBaseName:='SOC' ; END
            Else QAjoute:=InitQAjoute ;
*)
AlimLTabDiv(0,InfoImp.LMP) ;
InfoImp.NbLigIntegre:=0 ;
InfoImp.TotDeb:=0 ; InfoImp.TotCred:=0 ;
InfoImp.NbPiece:=0 ;
AnnuleImport:=False ; Erreur:=False ; OkRupt:=False ;
ChargeDevEtSect ;
OkOk:=ImporteEcr(StErr,InfoImp) ;
If Not OkOk Then
   BEGIN
   AnnuleImport:=FALSE ;
   FmtFic:=0 ;
   if InfoImp.Format='HLI' then FmtFic:=1 else
   if InfoImp.Format='HAL' then FmtFic:=2 Else
   if InfoImp.Format='CGN' then FmtFic:=3 Else
   if InfoImp.Format='CGE' then FmtFic:=4 ;
   if MsgBox<>nil then
     if MsgBox.Execute(54,'','')=mrYes then VisuLignesErreurs(InfoImp.NomFic,StErr,FmtFic,InfoImp.ForceBourrage) ;
   END ;
VideListe(TDev) ; TDev.Free ;
SourisNormale ;
if (Msg<>nil) and (not AnnuleImport) then
  BEGIN
  //ActivePanels(Self,True,False) ;
  if ChoixFmt.Detruire then DeleteFile(InfoImp.NomFic) ;
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

procedure ActiveTotaux(Ok : Boolean ; FAssImp : TForm ; QImpEcr : TQuery ; Var InfoImp : TInfoImport) ;
var ttJal,ttNatP : string ;
begin
if FAssImp=nil then Exit ;
if not Ok then BEGIN InfoImp.NbPiece:=0 ; InfoImp.TotDeb:=0 ; InfoImp.TotCred:=0 ; END ;
//if NbPiece<=1 then THLabel(FAssImp.FindComponent('LNbPiece')).Caption:=THMsgBox(FAssImp.FindComponent('HM')).Mess[7] else THLabel(FAssImp.FindComponent('LNbPiece')).Caption:=THMsgBox(FAssImp.FindComponent('HM')).Mess[8] ;
ttJal:='ttJournal' ;
if InfoImp.Prefixe='BE' then ttJal:='ttBudJal' ;
ttNatP:='ttNaturePiece' ;
if InfoImp.Prefixe='BE' then ttNatP:='ttNatEcrBud' ;
THLabel(FAssImp.FindComponent('DateP')).Caption:=DateToStr(QImpEcr.FindField('IE_DATECOMPTABLE').AsDateTime) ;
THLabel(FAssImp.FindComponent('JalP')).Caption:=RechDom(ttJal,QImpEcr.FindField('IE_JOURNAL').AsString,False) ;
THLabel(FAssImp.FindComponent('NatP')).Caption:=RechDom(ttNatP,QImpEcr.FindField('IE_NATUREPIECE').AsString,False) ;
THLabel(FAssImp.FindComponent('PieceP')).Caption:=IntToStr(QImpEcr.FindField('IE_NUMPIECE').AsInteger) ;
if (InfoImp.Prefixe='BE') then
  BEGIN
  THLabel(FAssImp.FindComponent('LigneP')).Visible:=False ;
  THLabel(FAssImp.FindComponent('Slash')).Visible:=False ;
  THLabel(FAssImp.FindComponent('THLigne')).Visible:=False ;
  THLabel(FAssImp.FindComponent('THSlash')).Visible:=False ;
  END else
  if (InfoImp.Prefixe='Y') then THLabel(FAssImp.FindComponent('LigneP')).Caption:=IntToStr(QImpEcr.FindField('IE_NUMVENTIL').AsInteger)
                   else THLabel(FAssImp.FindComponent('LigneP')).Caption:=IntToStr(QImpEcr.FindField('IE_NUMLIGNE').AsInteger) ;
THLabel(FAssImp.FindComponent('ZNb')).Caption:=IntToStr(InfoImp.NbPiece) ;
THLabel(FAssImp.FindComponent('ZNbLig')).Caption:=IntToStr(InfoImp.NbLigIntegre) ;
THNumEdit(FAssImp.FindComponent('ZTotDeb')).Text:=StrfMontant(Abs(InfoImp.TotDeb),20,V_PGI.OkDecV,V_PGI.SymbolePivot,True) ;
THNumEdit(FAssImp.FindComponent('ZTotCred')).Text:=StrfMontant(Abs(InfoImp.TotCred),20,V_PGI.OkDecV,V_PGI.SymbolePivot,True) ;
// GG ??? :
Application.ProcessMessages ;
END ;

{======================= Contôles/Rupture en intégration ==========================}

Procedure AlimIdentPieceI(T : TDataSet ; Var IdentPiece : TIdentPiece) ;
BEGIN
IdentPiece.JalP:=T.FindField('IE_JOURNAL').AsString ;
IdentPiece.NatP:=T.FindField('IE_NATUREPIECE').AsString ;
IdentPiece.NumP:=T.FindField('IE_NUMPIECE').AsInteger ;
IdentPiece.QualP:=T.FindField('IE_QUALIFPIECE').AsString ;
IdentPiece.DateP:=T.FindField('IE_DATECOMPTABLE').AsDateTime ;
IdentPiece.LigneEnCours.ANouveau:=(T.FindField('IE_ECRANOUVEAU').AsString='H') Or (T.FindField('IE_ECRANOUVEAU').AsString='OAN') ;
If T.FindField('IE_TYPEECR').AsString='E' Then
  With IdentPiece Do
    BEGIN
    LigneEnCours.DP:=T.FindField('IE_DEBIT').AsFloat ; LigneEnCours.CP:=T.FindField('IE_CREDIT').AsFloat ;
    LigneEnCours.DD:=T.FindField('IE_DEBITDEV').AsFloat ; LigneEnCours.CD:=T.FindField('IE_CREDITDEV').AsFloat ;
    LigneEnCours.DE:=T.FindField('IE_DEBITEURO').AsFloat ; LigneEnCours.CE:=T.FindField('IE_CREDITEURO').AsFloat ;
    LigneEnCours.Q1:=T.FindField('IE_QTE1').AsFloat ; LigneEnCours.Q2:=T.FindField('IE_QTE2').AsFloat ;
    LigneEnCours.QualQ1:=T.FindField('IE_QUALIFQTE1').AsString ; LigneEnCours.QualQ2:=T.FindField('IE_QUALIFQTE2').AsString ;
    LigneEnCours.NumLig:=T.FindField('IE_NUMLIGNE').AsInteger ;
    END ;
IdentPiece.LigneEnCours.Debit:=T.FindField('IE_DEBIT').AsFloat ; IdentPiece.LigneEnCours.Credit:=T.FindField('IE_CREDIT').AsFloat ;
IdentPiece.LigneEnCours.Gen:=T.FindField('IE_GENERAL').AsString ;
IdentPiece.LigneEnCours.Aux:=T.FindField('IE_AUXILIAIRE').AsString ;
IdentPiece.LigneEnCours.Sect:=T.FindField('IE_SECTION').AsString ;
IdentPiece.LigneEnCours.Axe:=T.FindField('IE_AXE').AsString ;
IdentPiece.LigneEnCours.ODAnal:=T.FindField('IE_TYPEANALYTIQUE').AsString='X' ;
END ;


Function EnRupture(T : TDataSet ; Var IdentPiece,OldIdentPiece : TIdentPiece ; Var InfoImp : TInfoImport) : Boolean ;
Var OkRupt : Boolean ;
BEGIN
Result:=False ;
if (InfoImp.Prefixe='BE')
  then OkRupt:=(IdentPiece.JalP<>OldIdentPiece.JalP) Or
               (IdentPiece.NatP<>OldIdentPiece.NatP) Or
               (IdentPiece.NumP<>OldIdentPiece.NumP) Or
               (IdentPiece.QualP<>OldIdentPiece.QualP)
  Else OkRupt:=(IdentPiece.JalP<>OldIdentPiece.JalP) Or
               (IdentPiece.NatP<>OldIdentPiece.NatP) Or
               (IdentPiece.NumP<>OldIdentPiece.NumP) Or
               (IdentPiece.DateP<>OldIdentPiece.DateP) Or
               (IdentPiece.QualP<>OldIdentPiece.QualP) ;
If OkRupt Then
  BEGIN
  OldIdentPiece.JalP:=T.Findfield('IE_JOURNAL').AsString ;
  OldIdentPiece.DateP:=T.Findfield('IE_DATECOMPTABLE').AsDateTime ;
  OldIdentPiece.NatP:=Trim(T.Findfield('IE_NATUREPIECE').AsString) ;
  OldIdentPiece.NumP:=T.Findfield('IE_NUMPIECE').AsInteger ;
  OldIdentPiece.QualP:=T.Findfield('IE_QUALIFPIECE').AsString ;
  Result:=TRUE ;
  END ;
//if Result then Inc(NbPiece);
END;

procedure MajImpErr(ListeEntetePieceFausse : TStringList) ;
var Nb,i,OldChrono : integer ;
    ImpError : TERRORIMP ;
    ImpPiece,OldPiece : TIdentPiece ;
    St,QuelleErreur,OldErreur : String ;
    Jal,Qual : String ;
    NumP : Integer ;
BEGIN
OldChrono:=0 ; OldErreur:='' ;
Nb:=ListeEntetePieceFausse.Count ;
InitMove(Nb,'') ;
for i:=0 to ListeEntetePieceFausse.Count-1 do
  BEGIN
  Jal:=Trim(Copy(ListeEntetePieceFausse[i],1,3)) ;
  Qual:=Trim(Copy(ListeEntetePieceFausse[i],4,1)) ;
  NumP:=StrToInt(Copy(ListeEntetePieceFausse[i],5,8)) ;
  (*
  ExecuteSQL('UPDATE IMPECR SET IE_OKCONTROLE="-" WHERE IE_JOURNAL="'+Jal+
             '" AND IE_NUMPIECE='+IntToStr(NumP)+' AND IE_QUALIFPIECE="'+Qual+'"'+
             ' AND IE_OKCONTROLE="X"')  ;
  *)
  ExecuteSQL('UPDATE IMPECR SET IE_OKCONTROLE="-" WHERE IE_JOURNAL="'+Jal+
             '" AND IE_NUMPIECE='+IntToStr(NumP)+' AND IE_OKCONTROLE="X"')  ;
  MoveCur(False) ;
  END ;
FiniMove ;
END ;

{======================= Lancement de l'intégration ==========================}

procedure MajOkControle ;
var St : string ;
BEGIN
St:='UPDATE IMPECR SET IE_OKCONTROLE="X" WHERE IE_TYPEECR<>"L" AND IE_OKCONTROLE<>"D"' ;
ExecuteSQL(St) ;
END ;

Procedure AlimNumDef(QImpEcr : TQuery ; Var IdentPiece : TIdentPiece ; Var InfoImp : TInfoImport) ;
var i : integer ;
    TTEcr : TTypeEcr ;
    CptLu : TCptLu ;
    Facturier : String ;
BEGIN
If InfoImp.ForceNumPiece Then BEGIN IdentPiece.NumPDef:=QImpEcr.Findfield('IE_NUMPIECE').AsInteger ; Exit ; END ;
//if (QImpEcr.Findfield('IE_ANA').AsString='-') and  (QImpEcr.Findfield('IE_TYPEANALYTIQUE').AsString='-') then TTEcr:=EcrGen else TTEcr:=EcrAna ;
if (QImpEcr.Findfield('IE_TYPEECR').AsString<>'A') and  (QImpEcr.Findfield('IE_TYPEANALYTIQUE').AsString='-') then TTEcr:=EcrGen else TTEcr:=EcrAna ;
if InfoImp.Prefixe='BE' then TTEcr:=EcrBud ;
if InfoImp.Prefixe='Y' then TTEcr:=EcrAna ;
CptLu.Cpt:=IdentPiece.JalP ;
If ChercheCptLu(InfoImp.LJalLu,CptLu) Then
  BEGIN
  Facturier:=CptLu.SoucheN ;
  If IdentPiece.QualP<>'N' Then Facturier:=CptLu.SoucheS ;
  if (Facturier='') then if FAssImp<>nil then BEGIN THMsgBox(FAssImp.FindComponent('HM')).Execute(43,'',' '+IdentPiece.JalP) ; Erreur:=True ; Exit ; END ;
  SetIncNum(TTEcr,Facturier,IdentPiece.NumPDef) ;
  END Else if FAssImp<>nil then BEGIN THMsgBox(FAssImp.FindComponent('HM')).Execute(43,'',' '+IdentPiece.JalP) ; Erreur:=True ; Exit ; END ;
END ;

procedure AjouteMvtDansLaBase(QImpEcr,QI : TQuery ; Var IdentPiece : TIdentPiece ; Var InfoImp : TInfoImport) ;
var StTable,Pref,StName : String ;
    CpteAuxi,CpteGene : String17 ;
    SectMvt : String17 ;
    i,k  : integer ;
    OkANouv : boolean ;
    CptLuJ : TCptLu ;
BEGIN
StTable:=InfoImp.Table ; Pref:=InfoImp.Prefixe+'_' ;
If (QImpEcr.Findfield('IE_TYPEECR').AsString='A') then
  BEGIN
  StTable:='ANALYTIQ' ; Pref:='Y_' ;
  END ;
QI.Insert ; InitNew(QI) ;
Fillchar(CptLuJ,SizeOf(CptLuJ),#0) ; CptLuJ.Cpt:=IdentPiece.JalP ;
ChercheCptLu(InfoImp.LJalLu,CptLuJ) ;

for i:=0 to QI.FieldCount-1 do
  BEGIN
  //PrendsLe:=True ;
  StName:=QI.Fields[i].FieldName ;
  StName:=Copy(StName,Pos('_',StName)+1,50) ;
  if (StName='NUMEROPIECE') then QI.Fields[i].AsInteger:=IdentPiece.NumPDef else
  if (StName='NUMLIGNE') then QI.Fields[i].AsInteger:=QImpEcr.FindField('IE_NUMLIGNE').AsInteger else
  if (StName='NUMECHE') then QI.Fields[i].AsInteger:=QImpEcr.FindField('IE_NUMECHE').AsInteger else
  if (StName='NUMVENTIL') then QI.Fields[i].AsInteger:=QImpEcr.FindField('IE_NUMVENTIL').AsInteger else
  if (StName='EXERCICE') then QI.Fields[i].AsString:=QUELEXODT(IdentPiece.DateP) else
  if (StName='DATECREATION') then QI.Fields[i].AsDateTime:=V_PGI.DateEntree else
  if (StName='DATEMODIF') then QI.Fields[i].AsDateTime:=NowH else
  if (StName='CREERPAR') then QI.Fields[i].AsString:='IMP' else
  if (StName='CONTROLETVA') then QI.Fields[i].AsString:='RIE' else
  if (StName='VALIDE') then BEGIN if CptLuJ.Nature='ANO' then QI.Fields[i].AsString:='X' ; END else
    // Ecritures budgétaires
  if (StName='BUDJAL') then QI.Fields[i].AsString:=QImpEcr.FindField('IE_JOURNAL').AsString else
  if (StName='BUDGENE') then QI.Fields[i].AsString:=QImpEcr.FindField('IE_GENERAL').AsString else
  if (StName='BUDSECT') then QI.Fields[i].AsString:=QImpEcr.FindField('IE_SECTION').AsString else
  if (StName='NATUREBUD') then QI.Fields[i].AsString:=QImpEcr.FindField('IE_NATUREPIECE').AsString else

  if (StName='DEBIT') then
    BEGIN
    QI.Fields[i].AsFloat:=QImpEcr.FindField('IE_DEBIT').AsFloat ;
    if (QImpEcr.FindField('IE_TYPEECR').AsString<>'A') then InfoImp.TotDeb:=InfoImp.TotDeb+QI.Fields[i].AsFloat ;
    END else
  if (StName='CREDIT') then
    BEGIN
    QI.Fields[i].Value:=QImpEcr.FindField('IE_CREDIT').AsFloat ;
    if (QImpEcr.Findfield('IE_TYPEECR').AsString<>'A') then InfoImp.TotCred:=InfoImp.TotCred+QI.Fields[i].AsFloat ;
    END else
  if StName='TOTALECRITURE' then
     BEGIN
     If (Not IdentPiece.LigneEnCours.OdAnal) Then QI.Fields[i].AsFloat:=IdentPiece.LigneEnCours.DP+IdentPiece.LigneEnCours.CP  ;
     END else
  if StName='TOTALDEVISE' then
     BEGIN
     If (Not IdentPiece.LigneEnCours.OdAnal) Then QI.Fields[i].AsFloat:=IdentPiece.LigneEnCours.DD+IdentPiece.LigneEnCours.CD ;
     END else
  if StName='TOTALEURO' then
     BEGIN
     If (Not IdentPiece.LigneEnCours.OdAnal) Then QI.Fields[i].AsFloat:=IdentPiece.LigneEnCours.DE+IdentPiece.LigneEnCours.CE ;
     END else
  if StName='TOTALQTE1' then
     BEGIN
     If (Not IdentPiece.LigneEnCours.OdAnal) Then QI.Fields[i].AsFloat:=IdentPiece.LigneEnCours.Q1  ;
     END else
  if StName='TOTALQTE2' then
     BEGIN
     If (Not IdentPiece.LigneEnCours.OdAnal) Then QI.Fields[i].AsFloat:=IdentPiece.LigneEnCours.Q2 ;
     END else
  if StName='QUALIFECRQTE1' then QI.Fields[i].AsString:=IdentPiece.LigneEnCours.QualQ1 else
  if StName='QUALIFECRQTE2' then QI.Fields[i].AsString:=IdentPiece.LigneEnCours.QualQ2 else
  if (StName='AUXILIAIRE') then
    BEGIN
    QI.Fields[i].Value:=QImpEcr.FindField('IE_'+StName).AsString ;
    END else
  if (StName='GENERAL') then
    BEGIN
    QI.Fields[i].Value:=QImpEcr.FindField('IE_'+StName).AsString ;
    END else
  if (StName='SECTION') then
    BEGIN
    QI.Fields[i].AsString:=QImpEcr.FindField('IE_'+StName).AsString ;
    END else
  if (StName='COTATION') then
    BEGIN
    QI.Fields[i].AsFloat:=QImpEcr.FindField('IE_'+StName).AsFloat ;
    END else
    BEGIN
    if (QImpEcr.FindField('IE_'+StName)<>Nil) then QI.Fields[i].Value:=QImpEcr.FindField('IE_'+StName).AsVariant ;
    END ;
  END ;
QI.Post ;
AlimTotauxCptLu(InfoImp,IdentPiece) ;
END ;

//Ecritures de format paramétré (Budget,analytiques...)

procedure IntegreParam(QImpEcr,QE,QY : TQuery ; Var InfoImp : TInfoImport) ;
Var IdentPiece,OldIdentPiece : TIdentPiece ;

BEGIN
Fillchar(IdentPiece,SizeOf(IdentPiece),#0) ; Fillchar(OldIdentPiece,SizeOf(OldIdentPiece),#0) ;
While not QImpEcr.EOF do
  BEGIN
  AlimIdentPieceI(QImpEcr,IdentPiece) ;
  if EnRupture(QImpEcr,IdentPiece,OldIdentPiece,InfoImp) then
    BEGIN
    Inc(InfoImp.NbPiece); AlimNumDef(QImpEcr,IdentPiece,InfoImp) ;
    (*
    NumL:=0 ; NumLigneGeneA0:=0 ; NumLigneEcheA0:=0 ; NumLigneVentilA0:=0 ;
    *)
    END ;
  if Erreur then Break ;
  if AnnuleImport then
    if FAssImp<>nil then if (THMsgBox(FAssImp.FindComponent('HM')).Execute(39,'','')=mryes) then Break else AnnuleImport:=False ;
  AjouteMvtDanslaBase(QImpEcr,QE,IdentPiece,InfoImp) ;
  Inc(InfoImp.NbLigIntegre) ;
  ActiveTotaux(True,FAssImp,QImpEcr,InfoImp) ;
  QImpEcr.Next ; MoveCur(False) ;
  END ;
END ;

// Ecritures générales
procedure IntegrePieces(QImpEcr,QE,QY : TQuery ; Var IdentPiece,OldIdentPiece : TIdentPiece ;
                        Var InfoImp : TInfoImport) ;
BEGIN
AlimIdentPieceI(QImpEcr,IdentPiece) ;
if EnRupture(QImpEcr,IdentPiece,OldIdentPiece,InfoImp) then
  BEGIN
  Inc(InfoImp.NbPiece); AlimNumDef(QImpEcr,IdentPiece,InfoImp) ;
  (*
  NumL:=0 ; NumLigneGeneA0:=0 ; NumLigneEcheA0:=0 ; NumLigneVentilA0:=0 ;
  *)
  END ;
// Ne prends pas les mouvements à zéro et les ligne d'éché / ventil. correspondantes
if not QImpEcr.Eof then
   BEGIN
   (*
   // RAZ n° lignes d'échéances.
   if (NumLigneEcheA0<0) and (QImpEcr.FindField('IE_ECHE').AsString='-') then NumLigneEcheA0:=0 ;
   // RAZ pour n° lignes ventilations.
   if (NumLigneVentilA0<0) and (QImpEcr.FindField('IE_NUMVENTIL').AsInteger=0) then NumLigneVentilA0:=0 ;
   *)
   // Ligne géné., éché. ou ventil. à 0.
   if (Arrondi(QImpEcr.FindField('IE_DEBIT').AsFloat,V_PGI.OkDecV)=0) and (Arrondi(QImpEcr.FindField('IE_CREDIT').AsFloat,V_PGI.OkDecV)=0) then
     BEGIN
     if not MvtAZero then
        BEGIN
        (*
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
        *)
        END ;
     Exit ;
     END else
     if MvtAZero then
        BEGIN
        If (QImpEcr.FindField('IE_TYPEECR').AsString='A') then Exit
                                                          else MvtAZero:=False ;
        END ;
   END ;
if (QImpEcr.FindField('IE_TYPEECR').AsString<>'A') then AjouteMvtDansLaBase(QImpEcr,QE,IdentPiece,InfoImp)
                                                   else AjouteMvtDansLaBase(QImpEcr,QY,IdentPiece,InfoImp) ;
END ;

procedure IntegreGenerales(QImpEcr,QE,QY : TQuery ; Var InfoImp : TInfoImport) ;
Var IdentPiece,OldIdentPiece : TIdentPiece ;
BEGIN
Fillchar(IdentPiece,SizeOf(IdentPiece),#0) ; Fillchar(OldIdentPiece,SizeOf(OldIdentPiece),#0) ;
While not QImpEcr.EOF do
  BEGIN
  if Erreur then Break ;
  if AnnuleImport then
    if FAssImp<>nil then if (THMsgBox(FAssImp.FindComponent('HM')).Execute(39,'','')=mryes) then Break else AnnuleImport:=False ;
  if (InfoImp.Format='RAP') or (InfoImp.Format='CRA') then RapprocheLigne(InfoImp) else
    IntegrePieces(QImpEcr,QE,QY,IdentPiece,OldIdentPiece,InfoImp) ;
  Inc(InfoImp.NbLigIntegre) ;
  ActiveTotaux(True,FAssImp,QImpEcr,InfoImp) ;
  QImpEcr.Next ; MoveCur(False) ;
  END ;
END ;

procedure IntegreEcr(FAss : TForm ; Var InfoImp : TInfoImport) ;
Var Nb : Integer ;
     Q : TQuery ;
     QImpEcr,QE,QY : TQuery ;
BEGIN
FAssImp:=FAss ;
InfoImp.Prefixe:='E' ; InfoImp.Table:='ECRITURE' ;
if InfoImp.Lequel='FBE' then BEGIN InfoImp.Prefixe:='BE' ; InfoImp.Table:='BUDECR' ; END else
  if InfoImp.Lequel='FOD' then BEGIN InfoImp.Prefixe:='Y' ; InfoImp.Table:='ANALYTIQ' ;END ;
AnnuleImport:=False ; Erreur:=False ;
BEGINTRANS ;
QImpEcr:=OpenSQL('SELECT * FROM IMPECR WHERE IE_SELECTED="X" AND IE_OKCONTROLE="X" AND IE_INTEGRE="-" AND IE_TYPEECR<>"L" ORDER BY IE_CHRONO',False) ;
QImpEcr.UpdateMode:=upWhereChanged ;
if (QImpecr.EOF) and (FAssImp<>nil) then
  BEGIN
  AnnuleImport:=True ;
  Ferme(QImpEcr) ;
  if (InfoImp.Format='RAP') or (InfoImp.Format='CRA') then THMsgBox(FAssImp.FindComponent('HM')).Execute(50,'','')
                                      else THMsgBox(FAssImp.FindComponent('HM')).Execute(42,'','') ;
  Exit ;
  END ;
InitMove(1000,'') ; MoveCur(False) ;
InfoImp.NbLigIntegre:=0 ; InfoImp.TotDeb:=0 ; InfoImp.TotCred:=0 ; InfoImp.NbPiece:=0 ;
QE:=PrepareSQL('SELECT * FROM '+InfoImp.Table+' WHERE '+InfoImp.Prefixe+'_EXERCICE="'+W_W+'"',FALSE) ;
QE.Open ;
if (InfoImp.Format<>'RAP') and (InfoImp.Format<>'CRA') and (InfoImp.Format<>'MP') and (InfoImp.Prefixe='E') then
  BEGIN
  QY:=PrepareSQL('SELECT * FROM ANALYTIQ WHERE Y_EXERCICE="'+W_W+'"',FALSE) ;
  QY.Open ;
  END ;
// Listes des fiches Cpte,Tiers,Jal et section
//RemplirLesListes ;
if (InfoImp.Lequel<>'FEC') or ((InfoImp.Lequel='FEC') and (InfoImp.Format='CRA')) then
   BEGIN
   InfoImp.ForceNumPiece:=FALSE ;
   IntegreParam(QImpEcr,QE,QY,InfoImp) ;
   END else IntegreGenerales(QImpEcr,QE,QY,InfoImp) ;
FiniMove ;
if (InfoImp.Format<>'RAP') and (InfoImp.Format<>'CRA') and (InfoImp.Format<>'MP') then BEGIN Ferme(QY) ; END ;
Ferme(QE) ; Ferme(QImpEcr) ;
if AnnuleImport or Erreur then ROLLBACK else COMMITTRANS ;
if (InfoImp.Lequel<>'FBE') then
  BEGIN
  if not AnnuleImport then MajSoldeCompteImport(InfoImp) ;
  END ;
if not AnnuleImport then ExecuteSQL('UPDATE IMPECR SET IE_INTEGRE="X" WHERE IE_SELECTED="X" AND IE_OKCONTROLE="X" AND IE_TYPEECR<>"L"') ;
if (FAss=nil) then CompteRenduBatch(1,InfoImp) ;
AnnuleImport:=False ;
END ;

Procedure RempliListeInfoImp(Var InfoImp : TInfoImport) ;
Var QFiche : TQFiche ;
    CptLu : TCptLu ;
    Q : TQuery ;
    i : Integer ;
BEGIN
For i:=0 To 3 Do InitRequete(QFiche[i],i) ;
//Q:=OpenSQL('Select IE_JOURNAL,IE_GENERAL,IE_AUXILIAIRE,IE_SECTION FROM IMPECR GROUP BY IE_JOURNAL,IE_GENERAL,IE_AUXILIAIRE,IE_SECTION',TRUE) ;
Q:=OpenSQL('Select IE_JOURNAL,IE_GENERAL,IE_AUXILIAIRE,IE_SECTION,IE_AXE FROM IMPECR GROUP BY IE_JOURNAL,IE_GENERAL,IE_AUXILIAIRE,IE_SECTION,IE_AXE',TRUE) ;
While Not Q.Eof Do
  BEGIN
  Fillchar(CptLu,SizeOf(CptLu),#0) ;
  CptLu.Cpt:=Q.Fields[1].AsString ;
  If AlimLTabCptLu(0,QFiche[0],InfoImp.LGenLu,InfoImp.ListeCptFaux,CptLu) Then ;

  Fillchar(CptLu,SizeOf(CptLu),#0) ;
  CptLu.Cpt:=Q.Fields[2].AsString ; CptLu.Collectif:=Q.Fields[1].AsString ; ;
  If AlimLTabCptLu(1,QFiche[1],InfoImp.LAuxLu,InfoImp.ListeCptFaux,CptLu) Then ;

  Fillchar(CptLu,SizeOf(CptLu),#0) ;
  CptLu.Cpt:=Q.Fields[3].AsString ; CptLu.Collectif:=Q.Fields[1].AsString ;
  CptLu.Axe:=Q.Fields[4].AsString ;
  If AlimLTabCptLu(2,QFiche[2],InfoImp.LAnaLu,InfoImp.ListeCptFaux,CptLu) Then ;

  Fillchar(CptLu,SizeOf(CptLu),#0) ;
  CptLu.Cpt:=Q.Fields[0].AsString ;
  If AlimLTabCptLu(3,QFiche[3],InfoImp.LJalLu,InfoImp.ListeCptFaux,CptLu) Then ;
  Q.Next ;
  END ;
Ferme(Q) ;
For i:=0 To 3 Do Ferme(QFiche[i]) ;
END ;

Function PbPiece(St : String ; InfoImp : TInfoImport ; What : Integer) : Boolean ;
Var Decal,DebNum,LongNum : Integer ;
    NumP : Integer ;
    JalP : String ;
    ClePF : String ;
    i : Integer ;
BEGIN
Decal:=0 ; Result:=FALSE ; If not VH^.ImportRL then Inc(Decal,8) ;
If EstUneLigneCpt(St) Then Exit ;
If (InfoImp.Format='SAA') Or (InfoImp.Format='SN2') Then
  BEGIN
  DebNum:=106 ; LongNum:=7 ;
  END Else
If (InfoImp.Format='HLI') Or (InfoImp.Format='HAL') Then
  BEGIN
  DebNum:=110+Decal ; LongNum:=7 ;
  END Else
If (InfoImp.Format='CGN') Or (InfoImp.Format='CGE') Then
  BEGIN
  DebNum:=152 ; LongNum:=8 ;
  END Else Exit ;
JalP:=Copy(St,1,3) ;
NumP:=StrToInt(Trim(Copy(St,DebNum,LongNum))) ;
ClePF:=Format_String(JalP,3)+Format_String(' ',1)+FormatFloat('00000000',NumP) ;
Case What Of
  0 : If InfoImp.ListeEntetePieceFausse.Find(ClePF,i) Then Result:=TRUE ; // Piece Fausse ;
  1 : If InfoImp.ListeEnteteDoublon.Find(ClePF,i) Then Result:=TRUE ; // Doublon ;
  END ;
END ;

Procedure FaitFichierRejet(InfoImp : TInfoImport) ;
Var Fichier,NewFichier1,NewFichier2 : TextFile ;
    St,StNewFichier2 : String ;
    OkFormat : Boolean ;
BEGIN
OkFormat:=(InfoImp.Lequel='FEC') and ((InfoImp.Format='SAA') Or
          (InfoImp.Format='SN2') or (InfoImp.Format='HAL') Or
          (InfoImp.Format='HLI') Or (InfoImp.Format='CGE') Or
          (InfoImp.Format='CGN')) ;
If Not OkFormat Then Exit ;
If (InfoImp.ListeEntetePieceFausse=Nil) Or (InfoImp.ListeEntetePieceFausse.Count=0) Then Exit ;
InitMove(1000,'') ;
{$i-} AssignFile(NewFichier1,InfoImp.NomFicRejet) ;  {$i+}
If IoResult<>0 Then Exit ;
Rewrite(NewFichier1) ;
AssignFile(Fichier,InfoImp.NomFic) ; Reset(Fichier) ;
StNewFichier2:=FileTemp('.PNM') ;
AssignFile(NewFichier2,StNewFichier2) ; Rewrite(NewFichier2) ;
ReadLn(Fichier,St) ; WriteLn(NewFichier1,St) ; WriteLn(NewFichier2,St) ;
While Not EOF(Fichier) do
  BEGIN
  MoveCur(FALSE) ; ReadLn(Fichier,St) ;
  If PbPiece(St,InfoImp,0) Then WriteLn(NewFichier1,St) Else WriteLn(NewFichier2,St) ;
  END ;
FiniMove ;
CloseFile(Fichier) ; CloseFile(NewFichier1) ; CloseFile(NewFichier2) ;
AssignFile(Fichier,InfoImp.NomFic) ; Erase(Fichier) ;
renamefile(StNewFichier2,InfoImp.NomFic) ;
END ;

Procedure FaitFichierDoublon(InfoImp : TInfoImport) ;
Var Fichier,NewFichier1,NewFichier2 : TextFile ;
    St,StNewFichier2 : String ;
    OkFormat : Boolean ;
BEGIN
If Not InfoImp.CtrlDB Then Exit ;
OkFormat:=(InfoImp.Lequel='FEC') and ((InfoImp.Format='SAA') Or
          (InfoImp.Format='SN2') or (InfoImp.Format='HAL') Or
          (InfoImp.Format='HLI') Or (InfoImp.Format='CGE') Or
          (InfoImp.Format='CGN')) ;
If Not OkFormat Then Exit ;
If (InfoImp.ListeEnteteDoublon=Nil) Or (InfoImp.ListeEnteteDoublon.Count=0) Then Exit ;
InitMove(1000,'') ;
{$i-} AssignFile(NewFichier1,InfoImp.NomFicDoublon) ;  {$i+}
If IoResult<>0 Then Exit ;
Rewrite(NewFichier1) ;
AssignFile(Fichier,InfoImp.NomFic) ; Reset(Fichier) ;
StNewFichier2:=FileTemp('.PNM') ;
AssignFile(NewFichier2,StNewFichier2) ; Rewrite(NewFichier2) ;
ReadLn(Fichier,St) ; WriteLn(NewFichier1,St) ; WriteLn(NewFichier2,St) ;
While Not EOF(Fichier) do
  BEGIN
  MoveCur(FALSE) ; ReadLn(Fichier,St) ;
  If PbPiece(St,InfoImp,1) Then WriteLn(NewFichier1,St) Else WriteLn(NewFichier2,St) ;
  END ;
FiniMove ;
CloseFile(Fichier) ; CloseFile(NewFichier1) ; CloseFile(NewFichier2) ;
AssignFile(Fichier,InfoImp.NomFic) ; Erase(Fichier) ;
renamefile(StNewFichier2,InfoImp.NomFic) ;
END ;


end.

