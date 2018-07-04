{$A+,B-,C-,D+,E-,F-,G+,H+,I+,J+,K-,L+,M-,N+,O+,P+,Q-,R-,S-,T-,U+,V+,W+,X+,Y+,Z1}
{$MINSTACKSIZE $00004000}
{$MAXSTACKSIZE $00100000}
{$IMAGEBASE $00400000}
{$APPTYPE GUI}
unit ImpUtil;

interface

uses
  SysUtils, Classes, Controls, Forms, Dialogs,
  StdCtrls, Hctrls, Ent1, HEnt1, DB, uDbxDataSet, ed_tools, Variants,
{$IFNDEF EAGLCLIENT}
  FmtChoix,
  CPTESAV,
  ImpFic,
{$ENDIF}
  HStatus, hmsgbox, IniFiles,
  ImpFicU,RappType,TImpFic,UtilPGI ;

{$IFNDEF EAGLCLIENT}
function ImporteLesEcritures(Msg : THMsgBox ; Var InfoImp : TInfoImport ; ignorel1 : boolean=False) : boolean ;
procedure MajImpErr(ListeEntetePieceFausse : HTStringList) ;
procedure IntegreEcr(FAss : TForm ; Var InfoImp : TInfoImport) ;
procedure MajOkControle ;
Procedure RempliListeInfoImp(Var InfoImp : TInfoImport) ;
Procedure FaitFichierRejet(InfoImp : TInfoImport) ;
Procedure FaitFichierDoublon(InfoImp : TInfoImport) ;
Procedure ChargeScenarioImport(Var InfoImp : TInfoImport ; Auto : Boolean) ;
procedure LettrageImport(Var InfoImp : TInfoImport) ;
Function ErreurMultiDevise(St : String ; InfoImp : TInfoImport ; Mvt : TFMvtImport ) : Boolean ;
{JP 14/11/05 : FQ 16567 : Teste si les comptes n'ont pas de caractères jokers et si c'est
               le cas, on les remplace par le caractères de bourrage}
procedure TesteJokers(var Mvt : TFMvtImport);
{$ENDIF EAGLCLIENT}

// Objets d'importation

Const ImpV3 : Boolean = FALSE ;

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


implementation

uses SaisUtil,SaisComm,LettUtil,LetBatch,RapSuppr,
{$IFNDEF EAGLCLIENT}
ImporFmt,VisuEnr,
{$ENDIF}
SoldeCpt,
ULibExercice,
TiersPayeur, UtilTrans ;


var

    // Variables pour l'importation
    MsgBox            : THMsgBox ;
    MvtAZero,OkRupt   : boolean ;
    NbCptes,NbLig : LongInt ;
    MvtImport : FMvtImport ;
    Erreur : boolean ;
    FAssImp : TForm ;
{$IFNDEF EAGLCLIENT}
Procedure TraitementFinRecupSISCO(Var InfoImp : TInfoImport) ;
Var Q   : TQuery ;
    St : String ;
    NbMvt : Integer ;
    ListeTiers : TStringList ;
    i : Integer ;
BEGIN
If VH^.RecupSISCOPGI Then
  BEGIN
  If Not (InfoImp.LettrageSISCOEnImport) Then Exit ;
  END ;
If (VH^.RecupPCL) Or ((Not VH^.RecupPCL) AND (InfoImp.LettrageSISCOEnImport)) Then Else Exit ;
If (Not RecupSISCO) And (Not RecupSISCOExt) Then Exit ;
begintrans ;
ListeTiers:=TStringList.Create ;
ListeTiers.Sorted:=TRUE ;
ListeTiers.Duplicates:=DupIgnore ;
InitMove(1000,'') ; NbMvt:=0 ;
St:='select e_auxiliaire, e_general, e_etatlettrage , e_lettrage ,min(e_datecomptable),max(e_datecomptable) from ecriture '+
    'where (e_etatlettrage="PL" or e_etatlettrage="TL") and e_qualiforigine="SIS" '+
    'group by e_auxiliaire, e_general, e_etatlettrage , e_lettrage' ;
Q:=OpenSQL(st,TRUE) ;
While Not Q.EOF do
   BEGIN
   MoveCur(FALSE) ;
   St:='UPDATE ECRITURE set E_datepaquetmin="'+USDateTime(Q.Fields[4].AsDateTime)+'", '+
       'E_DatePaquetMax="'+USDateTime(Q.Fields[5].AsDateTime)+'" '+
       'where e_auxiliaire="'+Q.Fields[0].AsString+'" and e_general="'+Q.Fields[1].AsString+'" '+
       'and e_etatlettrage="'+Q.Fields[2].AsString+'" and e_lettrage="'+Q.Fields[3].AsString+'"' ;
   ExecuteSql(St) ;
   ListeTiers.Add(Q.Fields[0].AsString) ;
   Q.Next ;
   if NbMvt>=1000 then BEGIN NbMvt:=0 ; FiniMove ; InitMove(100,'') ; END ;
   END ;
Ferme(Q) ;
For i:=0 To ListeTiers.Count-1 Do ExecuteSQL('UPDATE TIERS SET T_DERNLETTRAGE="AAAA" WHERE T_AUXILIAIRE="'+ListeTiers[i]+'" ') ;
MajTotTousComptes(False,'') ;
ListeTiers.Clear ;
ListeTiers.Free ;
committrans ; FiniMove ;
END ;

Procedure TraitementFinRecupSERVANT(Var InfoImp : TInfoImport) ;
BEGIN
MajTotTousComptes(False,'') ; LettrageSurRegroupement(InfoImp,'E_REFEXTERNE') ;
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

Procedure ChargeScenarioImport(Var InfoImp : TInfoImport ; Auto : Boolean) ;
BEGIN
ChargeScenario('X',InfoImp.Format,InfoImp.Lequel,InfoImp.CodeFormat,InfoImp.SC,Auto) ;
END ;

procedure CompteRenduBatch(QuelleErreur : Integer ; Var InfoImp : TInfoImport) ;
var F   : TextFile ;
    Fic,StName : String ;
BEGIN
StName:='' ;
If VH^.RecupSISCOPGI Then Exit ;
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
Result:=1 ;
END ;

Procedure ForceCommit(Var NbCount : Integer ; OkCommit : Boolean) ;
BEGIN
Inc(NbCount) ;
If NbCount>500 Then
  BEGIN
  NbCount:=0 ; If OkCommit Then BEGIN CommitTrans ; Delay(1000) ; BeginTrans ; END ;
  END ;
END ;

Function Format_Date(Dat : String) : TDateTime ;
var An : String ;
    Y,M,J : Word ;
    StD : String ;
BEGIN
Result:=iDate1900 ;
if Trim(Dat)='' then Exit ;
An:=Copy(Dat,5,2) ;
if ((Round(Valeur(An))>=0) and (Round(Valeur(An))<=80)) or (Round(Valeur(An))>=2000) then An:='20'+An
                                                           else An:='19'+An ;
Y:=Round(Valeur(An)) ;
M:=Round(Valeur(Copy(Dat,3,2))) ;
J:=Round(Valeur(Copy(Dat,1,2))) ;
If (J=0) Or (M=0) Or (Y=0) Then Exit ;
If J in [1..31]=FALSE Then Exit ;
If M in [1..12]=FALSE Then Exit ;
StD:=FormatFloat('00',J)+'/'+FormatFloat('00',M)+'/'+An ;
If Not IsValidDate(StD) Then
  BEGIN
  If m<>2 Then J:=30 Else J:=28 ;
  END ;
Result:=EncodeDate(Y,M,J) ;
END ;

function Format_DateEDI(St : String) : TDateTime ;
var Y,M,J : Word ;
BEGIN
if Trim(St)='' then St:='19000101' ;
Y:=Round(Valeur(Copy(St,1,4))) ;
M:=Round(Valeur(Copy(St,5,2))) ;
J:=Round(Valeur(Copy(St,7,2))) ;
Result:=EncodeDate(Y,M,J) ;
END ;

function TrouveModePaie(CategP : String3) : String3 ;
var Q : TQuery ;
BEGIN
Result:='' ;
Q:=OpenSQL('Select MP_MODEPAIE FROM MODEPAIE WHERE MP_CATEGORIE="'+Trim(CategP)+'" ORDER BY MP_MODEPAIE',True) ;
If Not Q.Eof Then Result:=Q.Fields[0].AsString ;
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

Function TraiteCorresp(Quoi : Integer ; QFiche : TQfiche ; Var Mvt : TFMvtImport ; Var InfoImp : TInfoImport) : Boolean ;
Var OkOk : Boolean ;
    CRType : String ;
    Cpt : String ;
    LeFb : TFichierBase ;
BEGIN
OkOk:=FALSE ; CRTYPE:='IGE' ; LeFb:=fbGene ; Cpt:='' ; Result:=FALSE ;
Case Quoi Of
  0 : BEGIN
      OkOk:=InfoImp.Sc.CorrespGen AND (Mvt.IE_GENERAL<>'') ; CRTYPE:='IGE' ; Cpt:=Mvt.IE_GENERAL ; LeFb:=fbGene ;
      END ;
  1 : BEGIN
      OkOk:=InfoImp.Sc.CorrespAux AND (Mvt.IE_AUXILIAIRE<>'') ; CRTYPE:='IAU' ; Cpt:=Mvt.IE_AUXILIAIRE ; LeFb:=fbAux ;
      END ;
  2 : BEGIN
      OkOk:=InfoImp.Sc.CorrespSect1 AND (Mvt.IE_SECTION<>'') And (Mvt.IE_AXE='A1') ; CRTYPE:='IA1' ; Cpt:=Mvt.IE_SECTION ; LeFb:=fbAxe1 ;
      END ;
  3 : BEGIN
      OkOk:=InfoImp.Sc.CorrespSect2 AND (Mvt.IE_SECTION<>'') And (Mvt.IE_AXE='A2') ; CRTYPE:='IA2' ; Cpt:=Mvt.IE_SECTION ; LeFb:=fbAxe2 ;
      END ;
  4 : BEGIN
      OkOk:=InfoImp.Sc.CorrespSect3 AND (Mvt.IE_SECTION<>'') And (Mvt.IE_AXE='A3') ; CRTYPE:='IA3' ; Cpt:=Mvt.IE_SECTION ; LeFb:=fbAxe3 ;
      END ;
  5 : BEGIN
      OkOk:=InfoImp.Sc.CorrespSect4 AND (Mvt.IE_SECTION<>'') And (Mvt.IE_AXE='A4') ; CRTYPE:='IA4' ; Cpt:=Mvt.IE_SECTION ; LeFb:=fbAxe4 ;
      END ;
  6 : BEGIN
      OkOk:=InfoImp.Sc.CorrespSect5 AND (Mvt.IE_SECTION<>'') And (Mvt.IE_AXE='A5') ; CRTYPE:='IA5' ; Cpt:=Mvt.IE_SECTION ; LeFb:=fbAxe5 ;
      END ;
  7 : BEGIN
      OkOk:=InfoImp.Sc.CorrespMP AND (Mvt.IE_MODEPAIE<>'') ; CRTYPE:='IPM' ; Cpt:=Mvt.IE_MODEPAIE ;
      END ;
  8 : BEGIN
      OkOk:=InfoImp.Sc.CorrespJal AND (Mvt.IE_JOURNAL<>'') ; CRTYPE:='IJA' ; Cpt:=Mvt.IE_JOURNAL ;
      END ;
  END ;
If Not OkOk Then Exit ;
If InfoImp.SC.Majuscule Then Cpt:=AnsiUpperCase(Cpt) ;
QFiche[5].Params[0].AsString:=CRTYPE ;
QFiche[5].Params[1].AsString:=Cpt ;
QFiche[5].Open ;
If Not QFiche[5].Eof Then
  BEGIN
  Result:=TRUE ;
  Case Quoi Of
    0 : Mvt.IE_GENERAL:=BourreOuTronque(QFiche[5].Fields[2].AsString,LeFb) ;
    1 : Mvt.IE_AUXILIAIRE:=BourreOuTronque(QFiche[5].Fields[2].AsString,LeFb) ;
    2..6 : Mvt.IE_SECTION:=BourreOuTronque(QFiche[5].Fields[2].AsString,LeFb) ;
    7 : Mvt.IE_MODEPAIE:=QFiche[5].Fields[2].AsString ;
    8 : Mvt.IE_JOURNAL:=QFiche[5].Fields[2].AsString ;
    END ;
  END ;
QFiche[5].Close ;
END ;

Procedure TrouveLeModePaie(StMp : String ; Var InfoImp : TInfoImport ; Var Mvt : TFMvtImport ; FormatCegid : Boolean ; QFiche : TQfiche ; OkForce : Boolean) ;
Var CptLuMP,CptLuG : TCptLu ;
    St,OldIE_MODEPAIE : String ;
    OkFormatCEGID,ChercheSurCat : Boolean ;
    StCpt : String ;
    OkOk : Boolean ;
BEGIN
If (Trim(Mvt.IE_MODEPAIE)='') And (StMP<>'') And OkForce Then Mvt.IE_MODEPAIE:=StMP ;
OldIE_MODEPAIE:=Mvt.IE_MODEPAIE ;
OkFormatCEGID:=FormatCegid Or InfoImp.SC.CorrespMP ;
If OkFormatCegid Then
  BEGIN
  If InfoImp.SC.CorrespMP Then
    BEGIN
    Mvt.IE_MODEPAIE:=StMp ;
    If TraiteCorresp(7,QFiche,Mvt,InfoImp) Then StMp:=Mvt.IE_MODEPAIE Else Mvt.IE_MODEPAIE:=OldIE_MODEPAIE ;
    END ;
  Fillchar(CptLuMP,SizeOf(CptLuMP),#0) ;
  If Trim(StMp)='' Then StMp:='W_W' ;
  CptLuMp.Cpt:=StMP ;
  ChercheSurCat:=TRUE ; If VH^.RecupSISCOPGI Then ChercheSurCat:=FALSE ;
  If Not LeModePaie(InfoImp.LMP,CptLuMp,ChercheSurCat) Then
    BEGIN
    OkOk:=InfoImp.SC.RMP<>'' ;
    If VH^.RecupSISCOPGI And (StMP<>'') And (StMP<>'W_W') Then
      BEGIN
      CreateMPSISCO(StMP,StMP,'CHQ','NON') ;
      AjouteLTabDiv(0,InfoImp.LMP,StMP,'NON') ;
      Mvt.IE_MODEPAIE:=StMP ;
      OkOk:=FALSE ;
      END ;
    If OkOk Then
      BEGIN
      Mvt.IE_MODEPAIE:=InfoImp.SC.RMP ;
      If VH^.RecupCegid Then
        BEGIN
        StCpt:=Copy(Mvt.IE_GENERAL,1,3) ;
        If StCpt='411' Then Mvt.IE_MODEPAIE:='CHQ' Else
         If StCpt='413' Then Mvt.IE_MODEPAIE:='TNS' Else
          If StCpt='416' Then Mvt.IE_MODEPAIE:='999' Else
           If StCpt='401' Then Mvt.IE_MODEPAIE:='CHQ' Else
             BEGIN
             Fillchar(CptLuG,SizeOf(CptLuG),#0) ;
             CptLuG.Cpt:=Mvt.IE_GENERAL ;
             If ChercheCptLu(InfoImp.LGenLu,CptLuG) Then If CptLuG.Lettrable Then Mvt.IE_MODEPAIE:='CHQ' ;
             END ;
        END ;
      END ;
    END ;
  END Else
  BEGIN
  St:=CHERCHEUNMODE(StMP,False,False) ;
  Fillchar(CptLuMP,SizeOf(CptLuMP),#0) ;
  CptLuMp.Nature:=St ;
  If LeModePaie(InfoImp.LMP,CptLuMp) Then Mvt.IE_MODEPAIE:=Trim(CptLuMP.Cpt) Else
    BEGIN
    Fillchar(CptLuMP,SizeOf(CptLuMP),#0) ;
    CptLuMp.Cpt:=St ;
    If LeModePaie(InfoImp.LMP,CptLuMp) Then Mvt.IE_MODEPAIE:=Trim(CptLuMP.Cpt) Else
      If InfoImp.SC.RMP<>'' Then Mvt.IE_MODEPAIE:=InfoImp.SC.RMP ;
    END ;
  END ;
END ;

function EnRuptAscii(AlimOldIdentPiece : Boolean ; Var IdentPiece,OldIdentPiece : TIdentPiece ; Var InfoImp : TInfoImport) : Boolean ;
var Per1,Per2 : Word ;
    AA,JJ : Word ;
BEGIN
Result:=False ;
If VH^.RecupPCL Or (ModeJal(IdentPiece.JalP,InfoImp)<>mPiece) Then
  BEGIN
  DecodeDate(OldIdentPiece.DateP,AA,Per1,JJ) ; Per1:=AA+Per1 ;
  DecodeDate(IdentPiece.DateP,AA,Per2,JJ) ; Per2:=AA+Per2 ;
  If (IdentPiece.JalP<>OldIdentPiece.JalP) or (Per2<>Per1) or (IdentPiece.NumP<>OldIdentPiece.NumP) then Result:=True ;
  END Else
  BEGIN
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
  END ;
If Result And AlimOldIdentPiece Then
  BEGIN
  OldIdentPiece.JalP:=IdentPiece.JalP ; OldIdentPiece.DateP:=IdentPiece.DateP ;
  OldIdentPiece.NumP:=IdentPiece.NumP ; OldIdentPiece.QualP:=IdentPiece.QualP ;
  OldIdentPiece.NatP:=IdentPiece.NatP ; MvtAZero:=False ;
  END ;
END;

function DetecteUneRupture(Var St : String ; Var IdentPiece : TIdentPiece ; Var InfoImp : TInfoImport ; QFiche : TQfiche) : Boolean ;
Var FmtFic : Integer ;
    JalP,NatP : String ;
    DateP : TDateTime ;
    NumP : Integer ;
    Decal : Integer ;
    TypeP : String ;
    LeJal : TCptImport ;
    Per1,Per2,AA,JJ : Word ;
BEGIN
Result:=False ; Decal:=0 ;
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
      TypeP:=Copy(St,27+Decal,1)
      END ;
  1,2 : BEGIN
        JalP:=Trim(Copy(St,1,3)) ;
        DateP:=Format_Date_HAL(Copy(St,4,8)) ;
        NatP:=Trim(Copy(St,12,2)) ;
        NumP:=Round(Valeur(Copy(St,110+Decal,7))) ;
        TypeP:=Copy(St,25,1) ;
        END ;
  3,4 : BEGIN
        JalP:=Trim(Copy(St,1,3)) ; DateP:=Format_Date_HAL(Copy(St,4,8)) ;
        NatP:=Copy(St,12,2)  ;
        NumP:=Round(Valeur(Copy(St,152,8))) ;
        TypeP:=Copy(St,31,1) ;
        END ;
  END ;
If TypeP='L' Then Exit ;
If InfoImp.SC.CorrespJal Then
  BEGIN
  Fillchar(LeJal,SizeOf(LeJal),#0) ; LeJal.Cpt:=JalP ;
  If TraiteCorrespCpt(8,QFiche,LeJal,InfoImp) Then JalP:=LeJal.Cpt ;
  END ;
If VH^.RecupPCL Or (ModeJal(JalP,InfoImp)<>mPiece) Then
  BEGIN
  DecodeDate(DateP,AA,Per1,JJ) ; Per1:=AA+Per1 ;
  DecodeDate(IdentPiece.DateP,AA,Per2,JJ) ; Per2:=AA+Per2 ;
  If (IdentPiece.JalP<>JalP) or (Per2<>Per1) or (IdentPiece.NumP<>NumP) then Result:=True ;
  END Else
  BEGIN
  If (IdentPiece.JalP<>JalP) or (IdentPiece.DateP<>DateP) or (IdentPiece.NumP<>NumP) or (IdentPiece.NatP<>NatP) then Result:=True ;
  END ;
END;



function TestBreak : Boolean ;
BEGIN
Result:=False ;
Application.ProcessMessages ;
if AnnuleImport and (MsgBox<>nil) then
  BEGIN
  if (MsgBox.Execute(39,'','')<>mryes) then AnnuleImport:=False ;
  Result:=AnnuleImport ;
  END ;
If HalteLa Then
  BEGIN
  AnnuleImport:=TRUE ;
  Result:=AnnuleImport ;
  END ;
END ;


// RECORD des mouvements à importer ;

function TrimSpe(St: String) : String ;
BEGIN
While (pos('"',St)<>0) do Delete(St,pos('"',St),1) ;
Result:=Trim(St) ;
END ;

procedure AjouteMvt(Var InfoImp : TInfoImport) ;

Var
SQL : string;
  function IntToStrSpe(St : Integer) : string ;
  var t : string ;
  BEGIN
  t:=IntToStr(st) ;
  if (not isnumeric(t)) then result:='0' else Result:=IntToStr(St) ;
  END ;

BEGIN

With MvtImport^ do
BEGIN
  if IE_TYPEECR<>'L' then Inc(InfoImp.NbLigIntegre) Else Inc(InfoImp.NbLigLettre) ;
  if ((IE_TYPEECR<>'A') and (IE_TYPEECR<>'L')) Or ((IE_TYPEECR='A') AND (IE_TYPEANALYTIQUE='X')) then
  //Cumuls pour Compte-rendu
     if ((MsgBox=nil) or (ChoixFmt.CompteRendu) or (InfoImp.Lequel='FBA')) then
     BEGIN
      InfoImp.TotDeb:=Arrondi(InfoImp.TotDeb+IE_DEBIT,V_PGI.OkDecV) ;
      InfoImp.TotCred:=Arrondi(InfoImp.TotCred+IE_CREDIT,V_PGI.OkDecV) ;
     END ;
     if InfoImp.ForcePositif then
     BEGIN
     if (IE_DEBIT<0) and (IE_CREDIT=0) then
     BEGIN
        IE_CREDIT:=Abs(IE_DEBIT) ; IE_DEBIT:=0 ;
        IE_CREDITDEV:=Abs(IE_DEBITDEV) ; IE_DEBITDEV:=0 ;
     END else
     if (IE_CREDIT<0) and (IE_DEBIT=0) then
        BEGIN
        IE_DEBIT:=Abs(IE_CREDIT) ; IE_CREDIT:=0 ;
        IE_DEBITDEV:=Abs(IE_CREDITDEV) ; IE_CREDITDEV:=0 ;
        END ;
     END ;
     SQL:= 'INSERT INTO IMPECR (' +
          'IE_CHRONO,IE_AFFAIRE,IE_ETATLETTRAGE,IE_LETTRAGE,IE_LETTREPOINTLCR,IE_LIBELLE,'+
          'IE_REFEXTERNE,IE_REFINTERNE,IE_REFLIBRE,IE_REFPOINTAGE,IE_REFRELEVE,'+
          'IE_RIB,IE_SECTION,IE_AXE,IE_JOURNAL,IE_ECRANOUVEAU,IE_ETABLISSEMENT,'+
          'IE_FLAGECR,IE_MODEPAIE,IE_NATUREPIECE,IE_REGIMETVA,IE_QUALIFPIECE,'+
          'IE_QUALIFQTE1,IE_QUALIFQTE2,IE_SOCIETE,IE_TPF,IE_TVA,IE_TYPEANOUVEAU,'+
          'IE_TYPEECR,IE_TYPEMVT,IE_DEVISE,IE_AUXILIAIRE,IE_GENERAL,IE_CONTREPARTIEAUX,'+
          'IE_CONTREPARTIEGEN,IE_NUMECHE,IE_NUMPIECEINTERNE,IE_NUMLIGNE,'+
          'IE_NUMPIECE,IE_NUMVENTIL,IE_DATECOMPTABLE,IE_DATEECHEANCE,IE_DATEPAQUETMAX,'+
          'IE_DATEPAQUETMIN,IE_DATEPOINTAGE,IE_DATEREFEXTERNE,IE_DATERELANCE,'+
          'IE_DATETAUXDEV,IE_DATEVALEUR,IE_ORIGINEPAIEMENT,IE_ECHE,IE_ENCAISSEMENT,'+
          'IE_CONTROLE,IE_LETTRAGEDEV,IE_OKCONTROLE,IE_SELECTED,IE_INTEGRE,IE_TVAENCAISSEMENT,'+
          'IE_TYPEANALYTIQUE,IE_VALIDE,IE_ANA,IE_POURCENTAGE,IE_POURCENTQTE1,'+
          'IE_POURCENTQTE2,IE_QTE1,IE_QTE2,IE_QUOTITE,IE_RELIQUATTVAENC,IE_TAUXDEV,'+
          'IE_TOTALTVAENC,IE_DEBIT,IE_CREDIT,IE_CREDITDEV,'+
          'IE_COUVERTURE,IE_COUVERTUREDEV,IE_DEBITDEV,'+
          'IE_LIBRETEXTE0,IE_LIBRETEXTE1,IE_LIBRETEXTE2,IE_LIBRETEXTE3,IE_LIBRETEXTE4,'+
          'IE_LIBRETEXTE5,IE_LIBRETEXTE6,IE_LIBRETEXTE7,IE_LIBRETEXTE8,IE_LIBRETEXTE9,'+
          'IE_TABLE0,IE_TABLE1,IE_TABLE2,IE_TABLE3,IE_LIBREBOOL0,IE_LIBREBOOL1,'+
          'IE_LIBREMONTANT0,IE_LIBREMONTANT1,IE_LIBREMONTANT2,IE_LIBREMONTANT3,IE_LIBREDATE,IE_COTATION '+
         ',IE_CODEACCEPT';
         SQL := SQL + ') VALUES ( ';
         SQL := SQL + IntToStrSpe(IE_CHRONO)+',"'+TrimSpe(IE_AFFAIRE)+'","'+
         TrimSpe(IE_ETATLETTRAGE)+'","'+TrimSpe(IE_LETTRAGE)+'","'+
         TrimSpe(IE_LETTREPOINTLCR)+'","'+TrimSpe(IE_LIBELLE)+'",';
         SQL := SQL + '"'+TrimSpe(IE_REFEXTERNE)+'","'+TrimSpe(IE_REFINTERNE)+'","'+
         TrimSpe(IE_REFLIBRE)+'","'+TrimSpe(IE_REFPOINTAGE)+'","'+TrimSpe(IE_REFRELEVE)+'",';
         SQL := SQL + '"'+TrimSpe(IE_RIB)+'","'+TrimSpe(IE_SECTION)+'","'+
         TrimSpe(IE_AXE)+'","'+ TrimSpe(IE_JOURNAL)+'","'+TrimSpe(IE_ECRANOUVEAU)+'","'+
         TrimSpe(IE_ETABLISSEMENT)+'",';

         SQL := SQL + '"'+TrimSpe(IE_FLAGECR)+'","'+TrimSpe(IE_MODEPAIE)+'","'+
         TrimSpe(IE_NATUREPIECE)+'","'+TrimSpe(IE_REGIMETVA)+'","'+TrimSpe(IE_QUALIFPIECE)+'",';
         SQL := SQL + '"'+TrimSpe(IE_QUALIFQTE1)+'","'+TrimSpe(IE_QUALIFQTE2)+'","'+
         TrimSpe(IE_SOCIETE)+'","'+TrimSpe(IE_TPF)+'","'+TrimSpe(IE_TVA)+'","'+
         TrimSpe(IE_TYPEANOUVEAU)+'",' ;
         SQL := SQL + '"'+TrimSpe(IE_TYPEECR)+'","'+TrimSpe(IE_TYPEMVT)+'","'+
         TrimSpe(IE_DEVISE)+'","'+TrimSpe(IE_AUXILIAIRE)+'","'+
         TrimSpe(IE_GENERAL)+'","'+TrimSpe(IE_CONTREPARTIEAUX)+'",';

         SQL := SQL + '"'+TrimSpe(IE_CONTREPARTIEGEN)+'",'+IntToStrSpe(IE_NUMECHE)+',"'+
         TrimSpe(IE_NUMPIECEINTERNE)+'",'+IntToStrSpe(IE_NUMLIGNE)+',';
         SQL := SQL + IntToStrSpe(IE_NUMPIECE)+','+IntToStrSpe(IE_NUMVENTIL)+',"'+
         USDateTime(IE_DATECOMPTABLE)+'","'+USDateTime(IE_DATEECHEANCE)+'","'+
         USDateTime(IE_DATEPAQUETMAX)+'",';
         SQL := SQL + '"'+USDateTime(IE_DATEPAQUETMIN)+'","'+
         USDateTime(IE_DATEPOINTAGE)+'","'+
         USDateTime(IE_DATEREFEXTERNE)+'","'+USDateTime(IE_DATERELANCE)+'",' ;
         SQL := SQL + '"'+USDateTime(IE_DATETAUXDEV)+'","'+
         USDateTime(IE_DATEVALEUR)+'","'+
         USDateTime(IE_ORIGINEPAIEMENT)+'","'+IE_ECHE+'","'+IE_ENCAISSEMENT+'",';
         SQL := SQL + '"'+IE_CONTROLE+'","'+IE_LETTRAGEDEV+'","'+
         IE_OKCONTROLE+'","'+IE_SELECTED+'","'+IE_INTEGRE+'","'+IE_TVAENCAISSEMENT+'",';
         SQL := SQL + '"'+IE_TYPEANALYTIQUE+'","'+IE_VALIDE+'","'+IE_ANA+'",'+
         StrfPoint(IE_POURCENTAGE)+','+StrfPoint(IE_POURCENTQTE1)+',' ;

         SQL := SQL + StrfPoint(IE_POURCENTQTE2)+','+StrfPoint(IE_QTE1)+','+StrfPoint(IE_QTE2)+','+
         StrfPoint(IE_QUOTITE)+','+StrfPoint(IE_RELIQUATTVAENC)+','+StrfPoint(IE_TAUXDEV)+',';
         SQL := SQL + StrfPoint(IE_TOTALTVAENC)+','+
         StrfPoint(IE_DEBIT)+','+StrfPoint(IE_CREDIT)+','+StrfPoint(IE_CREDITDEV)+',';
         SQL := SQL + StrfPoint(IE_COUVERTURE)+','+StrfPoint(IE_COUVERTUREDEV)+','+
         StrfPoint(IE_DEBITDEV)+',' ;
         SQL := SQL + '"'+TrimSpe(IE_LIBRETEXTE0)+'","'+TrimSpe(IE_LIBRETEXTE1)+'","'+TrimSpe(IE_LIBRETEXTE2)+'","'+
                         TrimSpe(IE_LIBRETEXTE3)+'","'+TrimSpe(IE_LIBRETEXTE4)+'",' ;
         SQL := SQL + '"'+TrimSpe(IE_LIBRETEXTE5)+'","'+TrimSpe(IE_LIBRETEXTE6)+'","'+TrimSpe(IE_LIBRETEXTE7)+'","'+
                         TrimSpe(IE_LIBRETEXTE8)+'","'+TrimSpe(IE_LIBRETEXTE9)+'",' ;
         SQL := SQL + '"'+TrimSpe(IE_TABLE0)+'","'+TrimSpe(IE_TABLE1)+'","'+TrimSpe(IE_TABLE2)+'","'+TrimSpe(IE_TABLE3)+'","'+
                         TrimSpe(IE_LIBREBOOL0)+'","'+TrimSpe(IE_LIBREBOOL1)+'",' ;

         SQL := SQL + StrfPoint(IE_LIBREMONTANT0)+','+StrfPoint(IE_LIBREMONTANT1)+','+StrfPoint(IE_LIBREMONTANT2)+','+
                         StrfPoint(IE_LIBREMONTANT3)+',"'+UsDateTime(IE_LIBREDATE)+'",'+StrfPoint(IE_COTATION) ;
         SQL := SQL + ',"'+TrimSpe(IE_CODEACCEPT)+'")' ;
         ExecuteSQL (SQL);
  END ;
END ;

{================= Rapprochement SAARI =====================}

Procedure AffecteMontantAutreAvecForceDevise(Var InfoImp : TInfoImport ;
                                             Var Mvt : TFMvtImport ; St,Sens : String ; Deb,Lg : Integer) ;
BEGIN
    if Sens='D' then Mvt.IE_DEBIT:=Valeur(StPoint(Copy(St,Deb,Lg))) else
      if Sens='C' then Mvt.IE_CREDIT:=Valeur(StPoint(Copy(St,Deb,Lg))) ;
END ;

Procedure ImportLigneHAL(St : String ; Var InfoImp : TInfoImport ; Var Mvt : TFMvtImport) ;
Var LgCpte,Decal : Byte ;
    Sens,NatP,TypeMvt : String ;
    ElementARecuperer : Boolean ;

BEGIN

ElementARecuperer:=MVT.IE_ELEMENTARECUPERER ;
Decal:=0 ; LgCpte:=13 ; if not VH^.ImportRL then LgCpte:=17 ;
If Not ElementARecuperer Then
  BEGIN
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
      AffecteMontantAutreAvecForceDevise(InfoImp,Mvt,St,Sens,89+Decal,20) ;
      (*
      if Sens='D' then IE_DEBIT:=Valeur(StPoint(Copy(St,89+Decal,20))) else
        if Sens='C' then IE_CREDIT:=Valeur(StPoint(Copy(St,89+Decal,20))) ;
      *)
      if (InfoImp.Format='HAL') and (InfoImp.Lequel<>'FBA') and (Trim(Copy(St,748+Decal,3))<>'') then NatP:=Trim(Copy(St,748+Decal,3)) ;
      IE_NATUREPIECE:=NatP ;
  //    IE_QUALIFPIECE:=Trim(Copy(St,109+Decal,1)) ;
      If Trim(Copy(St,109+Decal,1))<>'' Then IE_QUALIFPIECE:=Trim(Copy(St,109+Decal,1)) ;
      IE_NUMPIECE:=Round(Valeur(Copy(St,110+Decal,7))) ;
      END ;
  END ;
END ;

Procedure RecupInfoOrli(St : String ; Var InfoImp : TInfoImport ; Var Mvt : TFMvtImport) ;
Var Sens,Dev,MP,MontantDev,MontantPivot : String ;
    MDev : Double ;
{ Ne Fonctionne que pour des montants Francs ou devise. Pas pour les montants EURO }
BEGIN
Sens:=Copy(St,84,1) ; MontantDev:=trim(Copy(St,113,20)) ; MontantPivot:=Trim(Copy(St,85,20)) ;
MDev:=Valeur(StPoint(MontantDev)) ;
Dev:=Trim(Copy(St,133,3)) ;
If Dev='' Then Dev:=V_PGI.DevisePivot ;
Mvt.IE_DEVISE:=Dev ;
If (MontantDev<>'') And (Arrondi(MDev,3)<>0) And (Mvt.IE_DEVISE<>V_PGI.DevisePivot) Then
   BEGIN
   if Sens='D' then Mvt.IE_DEBITDEV:=MDev Else if Sens='C' then Mvt.IE_CREDITDEV:=MDev ;
   if EstMonnaieIn(Mvt.IE_DEVISE) Then BEGIN Mvt.IE_DEBIT:=0 ; Mvt.IE_CREDIT:=0 ; END ;
   END ;
MP:=Trim(Copy(St,146,3)) ; Mvt.IE_MODEPAIE:=MP ;
Mvt.IE_RIB:=Trim(Copy(St,149,50)) ;
END ;

Procedure ImportLigneAutre(St : String ; Var InfoImp : TInfoImport ; Var Mvt : TFMvtImport) ;
Var    TypeMvt,Sens  : String ;
       ElementARecuperer : Boolean ;
BEGIN
ElementARecuperer:=MVT.IE_ELEMENTARECUPERER ;
TypeMvt:=Copy(St,25,1) ;
// On ne prend que les échéances au format SN2
if (InfoImp.Format='SN2') and (TypeMvt='X') then Exit ;
Sens:=Copy(St,84,1) ;
//Jal:=Copy(St,1,3) ;DateC:=Format_Date(Copy(St,4,6)) ; if (TypeMvt<>'A') then NatP:=Copy(St,10,2) ;
If Not ElementARecuperer Then
  BEGIN
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
    AffecteMontantAutreAvecForceDevise(InfoImp,Mvt,St,Sens,85,20) ;
    (*
    if Sens='D' then IE_DEBIT:=Valeur(StPoint(Copy(St,85,20))) else
      if Sens='C' then IE_CREDIT:=Valeur(StPoint(Copy(St,85,20))) ;
    *)
  //  IE_QUALIFPIECE:=Trim(Copy(St,105,1)) ;
    If Trim(Copy(St,105,1))<>'' Then IE_QUALIFPIECE:=Trim(Copy(St,105,1)) ;
    IE_NUMPIECE:=Round(Valeur(Copy(St,106,7))) ;
    If Sn2Orli And (InfoImp.Format='SN2') Then RecupInfoOrli(St,InfoImp,Mvt) ;
    END ;
  END ;
END ;

Procedure AffecteDebitCredit(Sens : String ; Montant : Double ; Var Deb,Cre : Double) ;
BEGIN
Deb:=0 ; Cre:=0 ; If Montant=0 Then Exit ;
if Sens='D' then Deb:=Montant else if Sens='C' then Cre:=Montant ;
END ;

Procedure AffecteMontant(Qui : Char ; Montant : Double ; Sens : String ; Var Mvt : TFMvtImport ; MontantOrigine : Double) ;
Var EnEuro : Boolean ;
BEGIN
EnEuro:=EuroOK And VH^.TenueEuro ;
Case Qui Of
  'F' : (*If EnEuro Then AffecteDebitCredit(Sens,Montant,MVT.IE_DEBITEURO,MVT.IE_CREDITEURO)
                  Else*) AffecteDebitCredit(Sens,PivotToEuro(Montant),MVT.IE_DEBIT,MVT.IE_CREDIT) ;    // CA - 07/10/2003 - Pb des dossier SI2 avec exo en francs.
  'E' : (*If EnEuro Then*) AffecteDebitCredit(Sens,Montant,MVT.IE_DEBIT,MVT.IE_CREDIT);
                  //Else AffecteDebitCredit(Sens,Montant,MVT.IE_DEBITEURO,MVT.IE_CREDITEURO) ;
  'D' : AffecteDebitCredit(Sens,Montant,MVT.IE_DEBITDEV,MVT.IE_CREDITDEV) ;
  END ;
If (VH^.RecupPCL) And (Arrondi(MontantOrigine,2)<>0) Then
  BEGIN
  Case Qui Of
    'F' : BEGIN
          If EnEuro Then
            BEGIN
            if Sens='D' then MVT.IE_DEBIT:=MontantOrigine else if Sens='C' then MVT.IE_CREDIT:=MontantOrigine ;
            END Else
            BEGIN
            END ;
          END ;
    'E' : BEGIN
          If EnEuro Then
            BEGIN
            END Else
            BEGIN
            if Sens='D' then MVT.IE_DEBIT:=MontantOrigine else if Sens='C' then MVT.IE_CREDIT:=MontantOrigine ;
            END ;
          END ;
    END ;
  END ;
END ;

Function OnForceLaDevise(Var IdentPiece : TIdentPiece ; Var InfoImp : TInfoImport) : String ;
BEGIN
Result:=IdentPiece.CodeMontant ;
If (InfoImp.ForceDevise<>'') And (IdentPiece.CodeMontant<>'') And (IdentPiece.CodeMontant[1] in ['F','E']) Then
  BEGIN
  If InfoImp.ForceDevise='EUR' Then Result[1]:='E' Else
    If InfoImp.ForceDevise='FRF' Then Result[1]:='F' ;
  END ;
END ;

Procedure InterpreteMontant(St : String ; Var Mvt : TFMvtImport ; Var IdentPiece : TIdentPiece ; Var InfoImp : TInfoImport) ;
Var CodeMontant,Sens : String ;
    IndP,IndD,IndE : Integer ;
    Montant : Array[1..3] Of Double ;
    SMontant : Array[1..3] Of String ;
BEGIN
//CodeMontant:=Trim(Copy(St,173,3)) ;
If IdentPiece.CodeMontant='' Then
  BEGIN
  IdentPiece.CodeMontant:=Trim(Copy(St,173,3)) ;
  IdentPiece.CodeMontant:=OnForceLaDevise(IdentPiece,InfoImp) ;
  END ;
CodeMontant:=IdentPiece.CodeMontant ;
{$IFDEF CYBER}
IdentPiece.CodeMontant:=Trim(Copy(St,173,3)) ;
CodeMontant:=IdentPiece.CodeMontant ;
{$ENDIF}
If (CodeMontant='') Or (CodeMontant='---') Then CodeMontant:='F--' ;
IndP:=Pos('F',CodeMontant) ; IndD:=Pos('D',CodeMontant) ; IndE:=Pos('E',CodeMontant) ;
Fillchar(Montant,SizeOf(Montant),#0) ;
SMontant[1]:=Trim(StPoint(Copy(St,131,20))) ; SMontant[2]:=Trim(StPoint(Copy(St,176,20))) ;
SMontant[3]:=Trim(StPoint(Copy(St,196,20))) ;
If SMontant[1]<>'' Then Montant[1]:=Valeur(SMontant[1]) ;
If SMontant[2]<>'' Then Montant[2]:=Valeur(SMontant[2]) ;
If SMontant[3]<>'' Then Montant[3]:=Valeur(SMontant[3]) ;
Sens:=Copy(St,130,1) ;
If IndP>0 Then AffecteMontant('F',Montant[IndP],Sens,Mvt,Montant[2]) ;
If IndD>0 Then AffecteMontant('D',Montant[IndD],Sens,Mvt,Montant[2]) ;
If IndE>0 Then AffecteMontant('E',Montant[IndE],Sens,Mvt,Montant[2]) ;
END ;

Procedure RetoucheRecupPCL(Var Mvt : TFMvtImport) ;
Var S,SE : Double ;
    Cpt : String ;
BEGIN
If Not VH^.REcupPCL Then Exit ;
If VH^.RecupSISCOPGI Then exit ;
If (Mvt.IE_LIBELLE='Equilibrage au jour') Then
  BEGIN
  Cpt:='' ;
  S:=Arrondi(MVT.IE_DEBIT-MVT.IE_CREDIT,2) ;
  //SE:=Arrondi(MVT.IE_DEBITEURO-MVT.IE_CREDITEURO,2) ;
  SE := 0;
  If (S=0) AND (SE=0) Then Exit ;
  If Abs(SE)>1 Then Exit ; If abs(S)>1 Then Exit ;
//  If (S>0) Or (SE>0) Then Cpt:=VH^.EccEuroDebit ;
//  If (S<0) Or (SE<0) Then Cpt:=VH^.EccEuroCredit ;
  If Cpt<>'' Then MVT.IE_GENERAL:=Cpt ;
  END ;

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
  If Not ElementARecuperer Then BEGIN
    InitMvtImport(Mvt) ;
    With Mvt do BEGIN
     try
      IE_JOURNAL:=Trim(Copy(St,1,3)) ;
      IE_DATECOMPTABLE:=Format_Date_HAL(Copy(St,4,8)) ;
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
      InterpreteMontant(St,Mvt,IdentPiece,InfoImp) ;
      //  IE_QUALIFPIECE:=Copy(St,151,1) ;

      {JP 14/11/05 : FQ 16567 : On s'assure que les champs auxiliaire et généraux ne contiennent
                     pas de caractères Jokers, sinon les requêtes dans le PGI risque de connaître
                     quelques petits problèmes : dans ce cas, on les remplace par les caractères
                     de bourage}
      TesteJokers(Mvt);

      If Trim(Copy(St,151,1))<>'' Then IE_QUALIFPIECE:=Trim(Copy(St,151,1)) ;
      IE_NUMPIECE:=Round(Valeur(Copy(St,152,8))) ;
      if Trim(Copy(St,160,3))<>'' then IE_DEVISE:=Trim(Copy(St,160,3)) ;
      If IE_DEVISE='EUR' Then IE_DEVISE:=V_PGI.DevisePivot ;
      If VH^.TenueEuro And (IE_DEVISE=V_PGI.DeviseFongible) Then IE_DEVISE:=V_PGI.DevisePivot ;

      // YMO 11/10/2005 FQ15422 Verifier que le journal est multidevise si ecriture en devise
      // Choix de prendre 'V' pour les devises ('D' sert pour les erreurs doublons)
      if ErreurMultidevise(St, InfoImp, Mvt) then IE_OKCONTROLE := 'V';

      if Trim(Copy(St,163,10))<>'' then IE_TAUXDEV:=Valeur(StPoint(Copy(St,163,10)));
      if Trim(Copy(St,216,3))<>'' then IE_ETABLISSEMENT:=Trim(Copy(St,216,3)) ;
      if Trim(Copy(St,219,2))<>'' then IE_AXE:=Trim(Copy(St,219,2)) ;
      {JP 21/10/03 : Code oublié et récupéré de VDEV}
      try
        if Trim(Copy(St,222,1))<>'' then IE_NUMECHE := StrToInt(Trim(Copy(St,222,1)))
                                    else IE_NUMECHE := -1;
      except
        IE_NUMECHE := -1;
      end;
      RetoucheRecupPCL(Mvt) ;
     except
       on e:Exception do Showmessage(E.MEssage);
     end;
    END ;
  END ;
END ;

Function per4(DD : tDateTime) : String ;
Var Y,M,D : Word ;
BEGIN
DecodeDate(DD,Y,M,D) ;
per4:=Copy(IntToStr(Y),3,2)+FormatFloat('00',M) ;
END ;

Procedure InitRetoucheEnr(Var CptLuG,CptLuT,CptLuJ : TCptLu ;
                          Var Lettrable,Pointable,IsTiersT,IsTiersG,DeBanque : Boolean ;
                          Var Mvt : TFMvtImport ; Var InfoImp : TInfoImport) ;
Var CodeTva,CodeTPF : String ;
    Exo : TExoDate ;
    MonoExo : Boolean ;
BEGIN
Fillchar(CptLuG,SizeOf(CptLuG),#0) ; Fillchar(CptLuT,SizeOf(CptLuT),#0) ;
Fillchar(CptLuJ,SizeOf(CptLuJ),#0) ;
CodeTva:='' ; CodeTPF:='' ; DeBanque:=FALSE ;
With Mvt do
  BEGIN
  If IE_DEVISE='EUR' Then IE_DEVISE:=V_PGI.DevisePivot ;
  If VH^.TenueEuro And (IE_DEVISE=V_PGI.DeviseFongible) Then IE_DEVISE:=V_PGI.DevisePivot ;
  CptLuJ.Cpt:=IE_JOURNAL ; ChercheCptLu(InfoImp.LJalLu,CptLuJ) ;


  {$IFNDEF CMPGIS35} // 15098
    If (CptLuJ.Nature='ODA') or (CptLuJ.Nature='ANA') Then IE_TYPEANALYTIQUE:='X';
  {$ELSE}
    {$IFDEF COMPTA}
    {JP 13/06/03 : Lorsque l'on demande une RAZ, la nature du compte est à '', alors que
                   lorsque l'on importe un deuxième exercice depuis SISCO (donc sans RAZ),
                   la nature est à ODA => IE_TYPEANALYTIQUE:='X', ce qui fait que dans
                   "RetoucheFinal", Mvt.IE_NUMLIGNE est forcé à 0 provoquant ainsi un déséquilibre
                   dans les écritures analytiques et une violation de clé dans l'insertion
                   dans la table ANALYTIQ}
    If CptLuJ.Nature='ODA' Then begin
      CptLuJ.Nature:='OD';
      IE_TYPEANALYTIQUE:='-';
    end;
    {$ELSE}
    If CptLuJ.Nature='ODA' Then
      IE_TYPEANALYTIQUE:='X' ;
    {$ENDIF}
  {$ENDIF}

  If ImportBor Then
    BEGIN
    If (CptLuJ.ModeSaisie='BOR') Then IE_LIBREBOOL1:='B' Else If (CptLuJ.ModeSaisie='LIB') Then IE_LIBREBOOL1:='L' Else IE_LIBREBOOL1:='-' ;
    END Else IE_LIBREBOOL1:='-' ;
  IE_LETTREPOINTLCR:=Per4(MVT.IE_DATECOMPTABLE) ;
  CptLuT.Cpt:=IE_AUXILIAIRE ; ChercheCptLu(InfoImp.LAuxLu,CptLuT) ;
  CptLuG.Cpt:=IE_GENERAL ; ChercheCptLu(InfoImp.LGenLu,CptLuG) ;
  If CptLuG.Pointable Then
    BEGIN
    IE_NUMECHE:=1 ; IE_ETATLETTRAGE:='RI' ; If (CptLuG.Nature='BQE') Or (CptLuG.Nature='CAI') Then DeBanque:=TRUE ;
    END ;
  If CptLuG.Ventilable Then IE_ANA:='X' ;
  If CptLuJ.Nature='ANO' then
    BEGIN
    IE_ECRANOUVEAU:='OAN' ; If InfoImp.SC.ANDetail Then IE_ECRANOUVEAU:='H' ;
    If RecupSISCO Or RecupServant Or RecupSISCOExt Then
      BEGIN
      If QuelExoDate(IE_DATECOMPTABLE,IE_DATECOMPTABLE,MonoExo,Exo) Then
        BEGIN
        If Exo.Code=VH^.EnCours.Code Then IE_ECRANOUVEAU:='H' Else IE_ECRANOUVEAU:='OAN' ;
        IE_DATECOMPTABLE:=Exo.Deb ;
        END ;
      END ;
    END else if CptLuJ.Nature='CLO' then IE_ECRANOUVEAU:='C' ;
  (*
  If (CptLuJ.Nature='BQE') Then
    BEGIN
    If (IE_NATUREPIECE<>'RC') And (IE_NATUREPIECE<>'RF') And (IE_NATUREPIECE<>'OD') And (IE_NATUREPIECE<>'OF') And (IE_NATUREPIECE<>'OC') Then IE_NATUREPIECE:='OD' ;
    END ;
  *)
  if (IE_QUALIFPIECE='N') and (InfoImp.ForceQualif<>'') and (IE_ECRANOUVEAU<>'OAN') and (IE_ECRANOUVEAU<>'H') then IE_QUALIFPIECE:=InfoImp.ForceQualif ;
                                                                                              
  if VH^.RecupPCL And (IE_QUALIFPIECE='N') and (IE_ECRANOUVEAU<>'OAN') and (IE_ECRANOUVEAU<>'H') And
    (CptLuJ.SoucheS='SIS') Then IE_QUALIFPIECE:='S' ;

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
      CodeTva:=CptLuG.Tva ; CodeTPF:=CptLuG.TPF ;
      If Mvt.IE_TVA='' Then Mvt.IE_TVA:=CodeTva ;
      If Mvt.IE_TPF='' Then Mvt.IE_TPF:=CodeTPF ;
      If Mvt.IE_TVAENCAISSEMENT<>'X' Then Mvt.IE_TVAENCAISSEMENT:=CptLuG.TvaEncHT ;
      END ;
    END
  END ;
If InfoImp.SC.ValideEcr Then Mvt.IE_VALIDE:='X' ;
END ;

procedure RAZInfoLettrage(Let,Point,DeBanque : boolean ; Var Mvt : TFMvtImport) ;
BEGIN
With Mvt do
  BEGIN
   if (IE_TYPEECR<>'L') then
     BEGIN
     IE_LETTRAGE:='';IE_LETTRAGEDEV:='-';IE_COUVERTURE:=0; IE_COUVERTUREDEV:=0;
     IE_DATEPAQUETMIN:=IDate1900 ; IE_DATEPAQUETMAX:=IDate1900 ;
     END ;
  IE_ETATLETTRAGE:='RI' ;
  If Point Or Let Then IE_ENCAISSEMENT:=SensEnc(IE_DEBIT,IE_CREDIT) ;
  if Let then
    BEGIN
    IE_ETATLETTRAGE:='AL' ;
    IE_DATEPAQUETMIN:=IE_DATECOMPTABLE ; IE_DATEPAQUETMAX:=IE_DATECOMPTABLE ;
    END ;
  if Point then
    BEGIN
    IE_ETATLETTRAGE:='RI' ;
    If DeBanque Then BEGIN IE_NUMECHE:=1 ; IE_ECHE:='X' ; END Else BEGIN IE_NUMECHE:=0 ; IE_ECHE:='-' ; END ;
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

Procedure TraiteMontantDevise(Var Mvt : TFMvtImport ; Var IdentPiece : TIdentPiece ; Var InfoImp : TInfoImport) ;
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
    DateInutile : TDateTime ;
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
(*If (Sn2ORLI Or InfoImp.SC.CalcTauxDevOut) And (Not EstMonnaieIn(Mvt.IE_DEVISE)) And (Not EnEuro) And
  ((Mvt.IE_DEBIT<>0) Or (Mvt.IE_CREDIT<>0)) Then
  BEGIN
  Mvt.IE_DEBITEURO:=PivotToEuro(Mvt.IE_DEBIT) ;
  Mvt.IE_CREDITEURO:=PivotToEuro(Mvt.IE_CREDIT) ;
  END ;
If (Sn2ORLI Or InfoImp.SC.CalcTauxDevOut) And (Not EstMonnaieIn(Mvt.IE_DEVISE)) And (EnEuro) And
  ((Mvt.IE_DEBITEURO<>0) Or (Mvt.IE_CREDITEURO<>0)) Then
  BEGIN
  Mvt.IE_DEBIT:=PivotToEuro(Mvt.IE_DEBITEURO) ;
  Mvt.IE_CREDIT:=PivotToEuro(Mvt.IE_CREDITEURO) ;
  END ;*)
DebitE := 0; CreditE := 0;
If EnEuro Then
  BEGIN
//  DebitF:=Mvt.IE_DEBITEURO ; CreditF:=Mvt.IE_CREDITEURO ;
  DebitE:=Mvt.IE_DEBIT ; CreditE:=Mvt.IE_CREDIT ;
//  Decim:=V_PGI.OkDecE ;
  END Else
  BEGIN
//  DebitF:=Mvt.IE_DEBIT ; CreditF:=Mvt.IE_CREDIT ;
//  DebitE:=Mvt.IE_DEBITEURO ; CreditE:=Mvt.IE_CREDITEURO ;
//  Decim:=V_PGI.OkDecV ;
  END ;

  DebitF:=Mvt.IE_DEBIT ; CreditF:=Mvt.IE_CREDIT ;
  Decim:=V_PGI.OkDecV ;

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
for i:=0 to InfoImp.ListeDev.Count-1 do
  BEGIN
  TDevise:=InfoImp.ListeDev[i] ;
  if (TDevise.Code=Mvt.IE_DEVISE) then
    BEGIN
    Mvt.IE_QUOTITE:=TDevise.Quotite ;
    If Mvt.IE_QUOTITE=0 Then Mvt.IE_QUOTITE:=1 ;
    OkDev:=TRUE ;
    IdentPiece.DecimDev:=TDevise.Decimale ;
    If EstMonnaieIn(Mvt.IE_DEVISE) And (MVT.IE_DATECOMPTABLE>=V_PGI.DateDebutEuro) Then IdentPiece.TauxDev:=GetTaux(Mvt.IE_DEVISE,DateInutile,Mvt.IE_DATECOMPTABLE) ;
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
             Mvt.IE_DEBIT:=DebitNonCote ;
             Mvt.IE_CREDIT:=CreditNonCote ;

         (*If Mvt.IE_DATECOMPTABLE<V_PGI.DateDebutEuro Then
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
           *)
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
           if ((MVT.IE_DATECOMPTABLE<V_PGI.DateDebutEuro) or (Not EstMonnaieIn(MVT.IE_DEVISE)))  Then Mvt.IE_DATETAUXDEV:=LaDate
                                                                                              Else Mvt.IE_DATETAUXDEV:=V_PGI.DateDebutEuro ;
           END Else
           BEGIN
           Mvt.IE_TAUXDEV:=0 ; Mvt.IE_DEBIT:=0 ; Mvt.IE_CREDIT:=0 ;
           //Mvt.IE_DEBITEURO:=0 ; Mvt.IE_CREDITEURO:=0 ;
           END ;
         END Else Mvt.IE_TAUXDEV:=IdentPiece.TauxDev ;
       If Not PbTaux Then
         BEGIN
         If EnEuro Then
           BEGIN
           Mvt.IE_DEBIT:=DeviseToEuro(Mvt.IE_DEBITDEV,IdentPiece.TauxDev,Mvt.IE_QUOTITE) ;
           Mvt.IE_CREDIT:=DeviseToEuro(Mvt.IE_CREDITDEV,IdentPiece.TauxDev,Mvt.IE_QUOTITE) ;
           //Mvt.IE_DEBITEURO:=DeviseToPivot(Mvt.IE_DEBITDEV,IdentPiece.TauxDev,Mvt.IE_QUOTITE) ;
           //Mvt.IE_CREDITEURO:=DeviseToPivot(Mvt.IE_CREDITDEV,IdentPiece.TauxDev,Mvt.IE_QUOTITE) ;
           END Else
           BEGIN
           Mvt.IE_DEBIT:=DeviseToPivot(Mvt.IE_DEBITDEV,IdentPiece.TauxDev,Mvt.IE_QUOTITE) ;
           Mvt.IE_CREDIT:=DeviseToPivot(Mvt.IE_CREDITDEV,IdentPiece.TauxDev,Mvt.IE_QUOTITE) ;
           //Mvt.IE_DEBITEURO:=DeviseToEuro(Mvt.IE_DEBITDEV,IdentPiece.TauxDev,Mvt.IE_QUOTITE) ;
           //Mvt.IE_CREDITEURO:=DeviseToEuro(Mvt.IE_CREDITDEV,IdentPiece.TauxDev,Mvt.IE_QUOTITE) ;
           END ;
         END ;
       END ;
    END Else
    BEGIN { Code devise avec montant devise=0 }
    Mvt.IE_DEBIT:=0 ; Mvt.IE_CREDIT:=0 ;
    //Mvt.IE_DEBITEURO:=0 ; Mvt.IE_CREDITEURO:=0 ;
    END ;
  END Else
  BEGIN { Code Devise inexistant }
  Mvt.IE_TAUXDEV:=1;Mvt.IE_DEBITDEV:=Mvt.IE_DEBIT ;
  Mvt.IE_CREDITDEV:=Mvt.IE_CREDIT;Mvt.IE_QUOTITE:=1 ;
  END ;
END ;

Procedure TaiteMontantPivotEtOppose(Var Mvt : TFMvtImport ; Var IdentPiece : TIdentPiece) ;
{$IFDEF AVOIR}
Var     EnEuro : Boolean ;
{$ENDIF}
//        PasTouche : Boolean ;
BEGIN
IdentPiece.DecimDev:=V_PGI.OkDecV ;
IdentPiece.TauxDev:=1 ; Mvt.IE_TAUXDEV:=IdentPiece.TauxDev ;
{$IFDEF CYBER}
//PasTouche:=FALSE ;
//If ((MVT.IE_GENERAL=VH^.EccEuroDebit) Or (MVT.IE_GENERAL=VH^.EccEuroCredit)) And (Mvt.IE_TYPEECR='E') Then PasTouche:=TRUE ;
{$ENDIF}
{$IFDEF AVOIR}
EnEuro:=EuroOK And VH^.TenueEuro ;
If (Arrondi(Mvt.IE_DEBIT+Mvt.IE_CREDIT,V_PGI.OkDecV)=0) And (Arrondi(Mvt.IE_DEBITEURO+Mvt.IE_CREDITEURO,V_PGI.OkDecE)<>0) {And (Not PasTouche)} Then
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
If (Arrondi(Mvt.IE_DEBITEURO+Mvt.IE_CREDITEURO,V_PGI.OkDecE)=0) And (Arrondi(Mvt.IE_DEBIT+Mvt.IE_CREDIT,V_PGI.OkDecV)<>0) {And (Not PasTouche)} Then
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
{$ENDIF}
Mvt.IE_DEBITDEV:=Mvt.IE_DEBIT ; Mvt.IE_CREDITDEV:=Mvt.IE_CREDIT ;
END ;

Procedure InverseMontantANA(Var Mvt : TFMvtImport ; Var IdentPiece : TIdentPiece) ;
Var D,C,SP,SD : Double ;
    OkInverse : Boolean ;
BEGIN
D:=Arrondi(MVT.IE_DEBIT+MVT.IE_DEBITDEV,2) ;
C:=Arrondi(MVT.IE_CREDIT+MVT.IE_CREDITDEV,2) ;
If MVT.IE_TYPEECR='E' Then
  BEGIN
  IdentPiece.SensDernLigneGen:=0 ;
  If Arrondi(D,2)<>0 Then IdentPiece.SensDernLigneGen:=1 Else
    If Arrondi(C,2)<>0 Then IdentPiece.SensDernLigneGen:=2 ;
  END Else If MVT.IE_TYPEECR='A' Then
  BEGIN
  OkInverse:=FALSE ;
  If (IdentPiece.SensDernLigneGen=1) And (C<>0) Then OkInverse:=TRUE ;
  If (IdentPiece.SensDernLigneGen=2) And (D<>0) Then OkInverse:=TRUE ;
  If OkInverse Then
    BEGIN
    SP:=MVT.IE_DEBIT ; SD:=MVT.IE_DEBITDEV ;  
    MVT.IE_DEBIT:=(-1)*MVT.IE_CREDIT ; MVT.IE_DEBITDEV:=(-1)*MVT.IE_CREDITDEV ;
    MVT.IE_CREDIT:=(-1)*SP ; MVT.IE_CREDITDEV:=(-1)*SD ;
    END ;
  END ;
END ;

procedure RetoucheMontantCEGID(Var Mvt : TFMvtImport ; Var IdentPiece : TIdentPiece ; Var InfoImp : TInfoImport) ;
BEGIN
If Mvt.IE_DEVISE<>V_PGI.DevisePivot Then TraiteMontantDevise(Mvt,IdentPiece,InfoImp)
                                 Else TaiteMontantPivotEtOppose(Mvt,IdentPiece) ;
SetCotation(Mvt) ;
Mvt.IE_DEBIT:=Arrondi(Mvt.IE_DEBIT,V_PGI.OkDecV) ; Mvt.IE_CREDIT:=Arrondi(Mvt.IE_CREDIT,V_PGI.OkDecV) ;
InverseMontantANA(Mvt,IdentPiece) ;
END ;

procedure RetoucheMontantAutre(Var Mvt : TFMvtImport ; Var IdentPiece : TIdentPiece ; Var InfoImp : TInfoImport) ;
//Var EnEuro : Boolean ;
//    SauveD,SauveC : Double ;
BEGIN
//EnEuro:=EuroOK And VH^.TenueEuro ;
(* GP le 21/06/99
If EnEuro Then
  BEGIN
  SauveD:=Mvt.IE_DEBITEURO ; SauveC:=Mvt.IE_CREDITEURO ;
  Mvt.IE_DEBITEURO:=Mvt.IE_DEBIT ; Mvt.IE_CREDITEURO:=Mvt.IE_CREDIT ;
  Mvt.IE_DEBIT:=SauveD ; Mvt.IE_CREDIT:=SauveC ;
  END ;
*)
If Mvt.IE_DEVISE<>V_PGI.DevisePivot Then TraiteMontantDevise(Mvt,IdentPiece,InfoImp)
                                 Else TaiteMontantPivotEtOppose(Mvt,IdentPiece) ;
SetCotation(Mvt) ;
Mvt.IE_DEBIT:=Arrondi(Mvt.IE_DEBIT,V_PGI.OkDecV) ; Mvt.IE_CREDIT:=Arrondi(Mvt.IE_CREDIT,V_PGI.OkDecV) ;
InverseMontantANA(Mvt,IdentPiece) ;
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
                         Var Mvt : TFMvtImport ; OkRetoucheMontant,EcartDeConvertionAGenerer : Boolean ; QFiche : TQfiche) ;
Var CptLuG,CptLuT,CptLuJ : TCptLu ;
    Lettrable,Pointable,IsTiersT,IsTiersG,DeBanque : Boolean ;
    Decal : Integer ;
BEGIN
Decal:=0 ; if not VH^.ImportRL  then Decal:=8 ;
InitRetoucheEnr(CptLuG,CptLuT,CptLuJ,Lettrable,Pointable,IsTiersT,IsTiersG,DeBanque,Mvt,InfoImp) ;
With Mvt do
  BEGIN
  If (CptLuG.Lettrable Or CptLuT.Lettrable) And ((IE_TYPEECR='E') Or (IE_TYPEECR='L'))Then
     BEGIN
     if Trim(Copy(St,80+Decal,8))<>'' then
       BEGIN
       IE_DATEECHEANCE:=Format_Date_HAL(Copy(St,80+Decal,8)) ;
       If IE_DATEECHEANCE=IDate1900 Then IE_DATEECHEANCE:=IE_DATECOMPTABLE ;
       END else IE_DATEECHEANCE:=IE_DATECOMPTABLE ;
     IE_ECHE:='X' ;
     IE_DATEPAQUETMIN:=IE_DATECOMPTABLE ; IE_DATEPAQUETMAX:=IE_DATECOMPTABLE ;
     If (IE_TYPEECR='E') Or ((IE_ETATLETTRAGE<>'AL') And (IE_ETATLETTRAGE<>'PL') And(IE_ETATLETTRAGE<>'TL')) Then IE_ETATLETTRAGE:='AL' ;
//     If InfoImp.Format='HAL' Then IE_MODEPAIE:=Trim(Copy(St,745+Decal,3)) Else
     If InfoImp.Format='HAL' Then TrouveLeModePaie(Copy(St,745+Decal,3),InfoImp,Mvt,TRUE,QFiche,TRUE) Else
       If InfoImp.Format='HLI' Then TrouveLeModePaie(Copy(St,79+Decal,1),InfoImp,Mvt,FALSE,QFiche,TRUE) ;
     END ;
  If CptLuG.Pointable Then
    BEGIN
    If InfoImp.Format='HAL' Then IE_MODEPAIE:=Trim(Copy(St,745+Decal,3)) Else
      If InfoImp.Format='HLI' Then TrouveLeModePaie(Copy(St,79+Decal,1),InfoImp,Mvt,FALSE,QFiche,TRUE) ;
    if Trim(Copy(St,80+Decal,8))<>'' then IE_DATEECHEANCE:=Format_Date_HAL(Copy(St,80+Decal,8))
                                       else IE_DATEECHEANCE:=IE_DATECOMPTABLE ;
    END ;

  If (InfoImp.Format='HAL') And (Not EcartDeConvertionAGenerer) Then
     BEGIN
     if Trim(Copy(St,119+Decal,1))<>'' then IE_VALIDE:=Copy(St,119+Decal,1) ;
     if Trim(Copy(St,120+Decal,35))<>'' then IE_REFEXTERNE:=Copy(St,120+Decal,35) ;
     IE_DATEREFEXTERNE:=Format_Date_HAL(Copy(St,155+Decal,8)) ;
     //IE_DATECREATION:=Format_Date_HAL(Copy(St,163+Decal,8)) ;
     //IE_DATEMODIFICATION:=Format_Date_HAL(Copy(St,171+Decal,8)) ;
     IE_SOCIETE:=V_PGI.CodeSociete ; //Copy(St,179+Decal,3) ;
     if Trim(Copy(St,182+Decal,3))<>'' then IE_ETABLISSEMENT:=Trim(Copy(St,182+Decal,3)) ;
     if Trim(Copy(St,185+Decal,17))<>'' then IE_AFFAIRE:=Trim(Copy(St,185+Decal,17));
     IE_DEBITDEV:=Valeur(StPoint(Copy(St,265+Decal,20)));
     IE_CREDITDEV:=Valeur(StPoint(Copy(St,285+Decal,20)));
     if Trim(Copy(St,262+Decal,3))<>'' then IE_DEVISE:=Trim(Copy(St,262+Decal,3));
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
       (* IE_CONTROLE utilisé par VERIFCPTA
       if Trim(Copy(St,560,1))<>'' then IE_CONTROLE:=Copy(St,560,1);
       *)
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
       (* IE_CONTROLE utilisé par VERIFCPTA
       IE_CONTROLE:=Copy(St,422+Decal,1);
       *)
       IE_NUMVENTIL:=StrToIntDef(Copy(St,423+Decal,6),0);IE_POURCENTQTE1:=Valeur(StPoint(Copy(St,429+Decal,20)));IE_POURCENTQTE2:=Valeur(StPoint(Copy(St,449+Decal,20)));
       IE_TYPEANOUVEAU:=Trim(Copy(St,469,3));IE_TYPEMVT:=Trim(Copy(St,472+Decal,3));
       if Trim(Copy(St,475+Decal,1))<>'' then IE_TYPEANALYTIQUE:=Copy(St,475+Decal,1) ;
       END ;
     END ;
  END ;
RazInfoLettrage((IsTiersG or IsTiersT),DeBanque,CptLuG.Pointable,Mvt) ;
If OkRetoucheMontant Then RetoucheMontantAutre(Mvt,IdentPiece,InfoImp) ;
If CptLuG.Pointable Or (IsTiersG or IsTiersT) Then Mvt.IE_ENCAISSEMENT:=SensEnc(Mvt.IE_DEBIT,Mvt.IE_CREDIT) ;
Mvt.IE_CODEACCEPT:=MPToACC(Mvt.IE_MODEPAIE) ;
END ;

Procedure RetouchePourRecupSISCO(Var InfoImp : TInfoImport; Var Mvt : TFMvtImport ; IsLettrable : Boolean) ;
(*
Refexterne  : reference au folio SISCO
Reflibre    : Code lettre théorique PGI suite à récup SISCO
Libretexte0 : code stat SISCO
*)
Var St,StF,StL : String ;
    OkOk : Boolean ;
BEGIN
If (Not RecupSISCO) And (Not RecupSISCOExt) Then Exit ;
With Mvt Do
  BEGIN
  St:=IE_REFEXTERNE ; StF:='' ; StL:='' ;
  If St<>'' Then StF:=ReadTokenSt(St) ; If St<>'' Then StL:=ReadTokenSt(St) ;
  IE_REFEXTERNE:=StF ;
  If (Not IsLettrable) Or (Trim(StL)='') Then BEGIN IE_LIBRETEXTE0:=IE_REFLIBRE ; IE_REFLIBRE:='' ; Exit ; END ;
  If Not IsLettrable Then Exit ;
  If Trim(StL)='' Then Exit ;
  IE_LIBRETEXTE0:=IE_REFLIBRE ;
  IE_REFLIBRE:=STL ;
  If VH^.RecupSISCOPGI Then OkOk:=InfoImp.LettrageSISCOEnImport
                       Else OkOk:=VH^.RecupPCL Or InfoImp.LettrageSISCOEnImport ;
  If OkOk Then
    BEGIN
    IE_ETATLETTRAGE:='TL' ;
    IE_LETTRAGE:='AAAA' ;
    IE_COUVERTURE:=IE_DEBIT+IE_CREDIT ;
    IE_COUVERTUREDEV:=IE_DEBITDEV+IE_CREDITDEV ;
    IE_LETTRAGEDEV:='-' ;
    END ;
  END ;
END ;

Procedure retoucheEnrCEGID(St : String ; Var IdentPiece : TIdentPiece ; Var InfoImp : TInfoImport;
                           Var Mvt : TFMvtImport ; OkRetoucheMontant : Boolean ; QFiche : TQfiche) ;
Var CptLuG,CptLuT,CptLuJ,CptLuMP : TCptLu ;
    Lettrable,Pointable,IsTiersT,IsTiersG,DeBanque : Boolean ;
    RIB : String ;
    i : Integer ;
    OkRetoucheMP : Boolean ;
    StP,StN,StI : String ;
    StPer : String ;
BEGIN
Fillchar(CptLuMP,SizeOf(CptLuMP),#0) ; OkRetoucheMP:=FALSE ;
InitRetoucheEnr(CptLuG,CptLuT,CptLuJ,Lettrable,Pointable,IsTiersT,IsTiersG,DeBanque,Mvt,InfoImp) ;
With Mvt do
  BEGIN
  If (CptLuG.Lettrable Or CptLuT.Lettrable) And (IE_TYPEECR='E') Then
     BEGIN
     if Trim(Copy(St,122,8))<>'' then
       BEGIN
       IE_DATEECHEANCE:=Format_Date_HAL(Copy(St,122,8)) ;
       If IE_DATEECHEANCE=IDate1900 Then IE_DATEECHEANCE:=IE_DATECOMPTABLE ;
       END else IE_DATEECHEANCE:=IE_DATECOMPTABLE ;
     IE_ECHE:='X' ;
     IE_DATEPAQUETMIN:=IE_DATECOMPTABLE ; IE_DATEPAQUETMAX:=IE_DATECOMPTABLE ;
     IE_ETATLETTRAGE:='AL' ;
     If (InfoImp.Format='CGN') Or (InfoImp.Format='CGE') Then BEGIN IE_MODEPAIE:=Trim(Copy(St,119,3)) ; OkRetoucheMP:=TRUE ; END ;
     END ;
  If CptLuG.Pointable Then
    BEGIN
    IE_MODEPAIE:=Trim(Copy(St,119,3)) ;
    if Trim(Copy(St,122,8))<>'' then IE_DATEECHEANCE:=Format_Date_HAL(Copy(St,122,8))
                                else IE_DATEECHEANCE:=IE_DATECOMPTABLE ;
    OkRetoucheMP:=TRUE ;
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
     if Trim(Copy(St,387,3))<>'' then IE_REGIMETVA:=Trim(Copy(St,387,3));
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
//     IE_NUMEROIMMO:=Copy(St,516,17) ;
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
     If VH^.RecupSISCOPGI And (Trim(IE_REFPOINTAGE)<>'') Then
       BEGIN
       If CptLuG.Pointable Then
         BEGIN
         StP:='' ; StN:='' ;
         IE_REFPOINTAGE:=Trim(IE_REFPOINTAGE) ;
         if IE_REFPOINTAGE<>'' Then StP:=ReadTokenSt(IE_REFPOINTAGE) ;
         if IE_REFPOINTAGE<>'' Then StN:=ReadTokenSt(IE_REFPOINTAGE) ;
         If StP='X' Then
           BEGIN
           IE_REFPOINTAGE:='POINTE' ;
           IE_DATEPOINTAGE:=FinDeMois(IE_DATECOMPTABLE) ;
           InfoImp.PointageAFaire:=TRUE ;
           StPer:=DateToStr(IE_DATEPOINTAGE) ;
           StI:=CptLuG.Cpt+','+StPer ;
           If pos(StI,InfoImp.StEEXBQ)<=0 Then InfoImp.StEEXBQ:=InfoImp.StEEXBQ+StI+';' ;
           END ;
         If StN<>'' Then IE_LIBRETEXTE9:=StN ;
         END Else
         BEGIN
         IE_REFPOINTAGE:='' ;
         IE_DATEPOINTAGE:=IDate1900 ;
         END ;
       END ;
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
       if Trim(Copy(St,1021,3))<>'' then IE_ETATLETTRAGE:=Copy(St,1021,3);
       END ;
{$IFDEF CYBER}
     If VH^.Cyber Then
       BEGIN
       IE_LIBRETEXTE6:=Copy(St,1014,5);
       END ;
{$ENDIF}
     END ;
  END ;
RazInfoLettrage((IsTiersG or IsTiersT),DeBanque,CptLuG.Pointable,Mvt) ;
If OkRetoucheMP And ((InfoImp.SC.RMP<>'') Or (InfoImp.SC.CorrespMP)) Then TrouveLeModePaie(Mvt.IE_MODEPAIE,InfoImp,Mvt,TRUE,QFiche,FALSE) ;
If OkRetoucheMontant Then RetoucheMontantCEGID(Mvt,IdentPiece,InfoImp) ;
If CptLuG.Pointable Or (IsTiersG or IsTiersT) Then Mvt.IE_ENCAISSEMENT:=SensEnc(Mvt.IE_DEBIT,Mvt.IE_CREDIT) ;
If RecupSISCO Or RecupSISCOExt  Then RetouchePourRecupSISCO(InfoImp,Mvt,(IsTiersG or IsTiersT) );
Mvt.IE_CODEACCEPT:=MPToACC(Mvt.IE_MODEPAIE) ;
END ;

Procedure retoucheEnrAutre(St : String ; Var IdentPiece : TIdentPiece ; Var InfoImp : TInfoImport ;
                           Var Mvt : TFMvtImport ; OkRetoucheMontant : Boolean ; QFiche : TQfiche) ;
Var CptLuG,CptLuT,CptLuJ : TCptLu ;
    Lettrable,Pointable,IsTiersT,IsTiersG,DeBanque : Boolean ;
BEGIN
InitRetoucheEnr(CptLuG,CptLuT,CptLuJ,Lettrable,Pointable,IsTiersT,IsTiersG,DeBanque,Mvt,InfoImp) ;
With Mvt do
  BEGIN
  If (CptLuG.Lettrable Or CptLuT.Lettrable) And ((IE_TYPEECR='E') Or (IE_TYPEECR='L')) Then
     BEGIN
     if ((Trim(Copy(St,78,6))<>''))  then
       BEGIN
       IE_DATEECHEANCE:=Format_Date(Copy(St,78,6)) ;
       If IE_DATEECHEANCE=IDate1900 Then IE_DATEECHEANCE:=IE_DATECOMPTABLE ;
       END else IE_DATEECHEANCE:=IE_DATECOMPTABLE ;
     IE_ECHE:='X' ;
     IE_DATEPAQUETMIN:=IE_DATECOMPTABLE ; IE_DATEPAQUETMAX:=IE_DATECOMPTABLE ;
     If (IE_TYPEECR='E') Or ((IE_ETATLETTRAGE<>'AL') And (IE_ETATLETTRAGE<>'PL') And(IE_ETATLETTRAGE<>'TL')) Then IE_ETATLETTRAGE:='AL' ;
     If (InfoImp.Format='SN2') ANd Sn2Orli Then TrouveLeModePaie(Copy(St,77,1),InfoImp,Mvt,TRUE,QFiche,TRUE)
                                           Else TrouveLeModePaie(Copy(St,77,1),InfoImp,Mvt,FALSE,QFiche,TRUE) ;
     END ;
  If CptLuG.Pointable Then
    BEGIN
    TrouveLeModePaie(Copy(St,77,1),InfoImp,Mvt,FALSE,QFiche,TRUE) ;
    if ((Trim(Copy(St,78,6))<>''))  then IE_DATEECHEANCE:=Format_Date(Copy(St,78,6))
                                    else IE_DATEECHEANCE:=IE_DATECOMPTABLE ;
    END ;
  END ;
RazInfoLettrage((IsTiersG or IsTiersT),DeBanque,CptLuG.Pointable,Mvt) ;
If OkRetoucheMontant Then RetoucheMontantAutre(Mvt,IdentPiece,InfoImp) ;
If CptLuG.Pointable Or (IsTiersG or IsTiersT) Then Mvt.IE_ENCAISSEMENT:=SensEnc(Mvt.IE_DEBIT,Mvt.IE_CREDIT) ;
Mvt.IE_CODEACCEPT:=MPToACC(Mvt.IE_MODEPAIE) ;
END ;


Function TestRupturePiece(Format : String ; AlimOldIdentPiece : Boolean ; Var IdentPiece,OldIdentPiece : TIdentPiece ;
                          Var InfoImp : TInfoImport) : Boolean ;
BEGIN
if (Format='RAP') or (Format='CRA') or (Trim(Format)='MP') then Result:=TRUE
                                                           Else Result:=EnRuptAscii(AlimOldIdentPiece,IdentPiece,OldIdentPiece,InfoImp) ;
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

Procedure TraiteTiersInexistant(RAux,Nat : String ; QFiche : TQfiche ; Var Mvt : TFMvtImport ; Var InfoImp : TInfoImport) ;
Var CptLuT,CptLuG : TCptLu ;
    OkOk : Boolean ;
BEGIN
If (RAux<>'') Then
  BEGIN
  CptLuG.Cpt:=Mvt.IE_GENERAL ;
  If AlimLTabCptLu(0,QFiche[0],InfoImp.LGenLu,NIL,CptLuG) Then
    BEGIN
    OkOk:=(CptLuG.Nature=Nat) And (Nat<>'CHA') ;
    If (VH^.RecupCegid) And (Nat='CHA') Then
      BEGIN
      If CptLuG.EstColl And ((CptLuG.Nature='CHA') Or (CptLuG.Nature='PRO')) Then OkOk:=TRUE ;  
      END ;
    If OkOk Then
      BEGIN
      CptLuT.Cpt:=Mvt.IE_AUXILIAIRE ;
      If Not AlimLTabCptLu(1,QFiche[1],InfoImp.LAuxLu,NIL,CptLuT) Then Mvt.IE_AUXILIAIRE:=BourreOuTronque(Raux,fbAux) ;
      END ;
    END ;
  END ;
END ;

Procedure TraiteSectInexistant(RSect,Axe : String ; QFiche : TQfiche ; Var Mvt : TFMvtImport ; Var InfoImp : TInfoImport) ;
Var CptLuS,CptLuG : TCptLu ;
    leFb : TFichierBase ;
    SectionACreer : Boolean ;
    LibACreer : String ;
    QS : TQuery ;
BEGIN
If Mvt.IE_AXE<>Axe Then Exit ;
If (RSect<>'') Then
  BEGIN
  SectionACreer:=Copy(RSect,1,2)='**' ;
  CptLuG.Cpt:=Mvt.IE_GENERAL ;
  If AlimLTabCptLu(0,QFiche[0],InfoImp.LGenLu,NIL,CptLuG) Then
    BEGIN
    If CptLuG.Ventilable Then
      BEGIN
      CptLuS.Cpt:=Mvt.IE_SECTION ; CptLuS.Axe:=Mvt.IE_AXE ;
//      If CptLuG.Axe[StrToInt(Copy(Axe,2,1))]<>Axe[2] Then Exit ;
      If Mvt.IE_AXE[2]<>Axe[2] Then Exit ;
      Lefb:=fbAxe1 ;
      if (Mvt.IE_AXE='A2') then Lefb:=fbAxe2 else
      if (Mvt.IE_AXE='A3') then Lefb:=fbAxe3 else
      if (Mvt.IE_AXE='A4') then Lefb:=fbAxe4 else
      if (Mvt.IE_AXE='A5') then Lefb:=fbAxe5 ;
      CptLuS.Cpt:=BourreOuTronque(CptLuS.Cpt,leFb) ;
      If Not AlimLTabCptLu(2,QFiche[2],InfoImp.LAnaLu,NIL,CptLuS) Then
        BEGIN
        If SectionACreer Then
          BEGIN
          QS:=OpenSQL('Select * From Section Where S_AXE="'+MVT.IE_AXE+'" AND S_SECTION="'+w_w+'" ',FALSE) ;
          QS.Insert ; InitNew(QS) ;
          QS.FindField('S_SECTION').AsString:=CptLuS.Cpt ;
          LibACreer:=LibSousSection(CptLuS.Cpt,Mvt.IE_AXE) ;
          QS.FindField('S_AXE').AsString:=Mvt.IE_AXE;
          If LibACreer<>'' Then QS.FindField('S_LIBELLE').AsString:=LibACreer
                           Else QS.FindField('S_LIBELLE').AsString:='Créé par importation de mouvements' ;
          QS.FindField('S_ABREGE').AsString:='Import' ;
          If QS.FindField('S_CODEIMPORT')<>NIL Then QS.FindField('S_CODEIMPORT').AsString:=CptLuS.Cpt ;
//          QS.FindField('S_CREERPAR').AsString:='IMP' ;
          SetCreerPar(QS,'S_CREERPAR') ;
          QS.Post ;
          END Else Mvt.IE_SECTION:=BourreOuTronque(RSect,leFb) ;
        END ;
      END ;
    END ;
  END ;
END ;

Procedure TraiteCompteSubstitue(QFiche : TQfiche ; Var Mvt : TFMvtImport ; Var InfoImp : TInfoImport) ;
Var CptLuT,CptLuG : TCptLu ;
//    i : Integer ;
BEGIN
//If InfoImp.SC.UseCorresp Then For i:=0 To 6 Do TraiteCorresp(i,QFiche,Mvt,InfoImp) ;
If Mvt.IE_AUXILIAIRE<>'' Then
  BEGIN
  If (InfoImp.SC.RCollCliT<>'') And (Copy(Mvt.IE_GENERAL,1,Length(InfoImp.SC.RCollCliT))=InfoImp.SC.RCollCliT) Then
     BEGIN
     CptLuT.Cpt:=Mvt.IE_AUXILIAIRE ;
     If AlimLTabCptLu(1,QFiche[1],InfoImp.LAuxLu,InfoImp.ListeCptFaux,CptLuT) Then
       BEGIN
       If (CptLuT.Nature='CLI') And (CptLuT.CollTiers<>'') Then
         BEGIN
         Mvt.IE_GENERAL:=BourreOuTronque(CptLuT.CollTiers,fbGene) ;
         END ;
       END ;
     END ;
  If (InfoImp.SC.RCollFouT<>'') And (Copy(Mvt.IE_GENERAL,1,Length(InfoImp.SC.RCollFouT))=InfoImp.SC.RCollFouT) Then
     BEGIN
     CptLuT.Cpt:=Mvt.IE_AUXILIAIRE ;
     If AlimLTabCptLu(1,QFiche[1],InfoImp.LAuxLu,InfoImp.ListeCptFaux,CptLuT) Then
       BEGIN
       If (CptLuT.Nature='FOU') And (CptLuT.CollTiers<>'') Then
         BEGIN
         Mvt.IE_GENERAL:=BourreOuTronque(CptLuT.CollTiers,fbGene) ;
         END ;
       END ;
     END ;
  END ;
If InfoImp.Sc.RGen<>'' Then
  BEGIN
  CptLuG.Cpt:=Mvt.IE_GENERAL ;
  If Not AlimLTabCptLu(0,QFiche[0],InfoImp.LGenLu,NIL,CptLuG) Then Mvt.IE_GENERAL:=BourreOuTronque(InfoImp.Sc.RGen,fbGene) ;
  END ;
TraiteTiersInexistant(InfoImp.Sc.RCli,'COC',QFiche,Mvt,InfoImp) ;
TraiteTiersInexistant(InfoImp.Sc.RFou,'COF',QFiche,Mvt,InfoImp) ;
TraiteTiersInexistant(InfoImp.Sc.RSal,'COS',QFiche,Mvt,InfoImp) ;
TraiteTiersInexistant(InfoImp.Sc.RDiv,'COD',QFiche,Mvt,InfoImp) ;
If VH^.RecupCEGID Then
  BEGIN
  TraiteTiersInexistant(InfoImp.Sc.RDiv,'CHA',QFiche,Mvt,InfoImp) ;
  END ;
TraiteSectInexistant(InfoImp.Sc.RSect1,'A1',QFiche,Mvt,InfoImp) ;
TraiteSectInexistant(InfoImp.Sc.RSect2,'A2',QFiche,Mvt,InfoImp) ;
TraiteSectInexistant(InfoImp.Sc.RSect3,'A3',QFiche,Mvt,InfoImp) ;
TraiteSectInexistant(InfoImp.Sc.RSect4,'A4',QFiche,Mvt,InfoImp) ;
TraiteSectInexistant(InfoImp.Sc.RSect5,'A5',QFiche,Mvt,InfoImp) ;
END ;

Procedure TraiteScenarioRib(Var Mvt : TFMvtImport ; Var InfoImp : TInfoImport ; Var CptLuT : TCptLu) ;
Var Rib : String ;
BEGIN
If InfoImp.SC.ForceRIB And (CptLuT.Nature='CLI') And (Trim(Mvt.IE_RIB)='') Then
  BEGIN
  Rib:=ChercheRibPrincipal(Mvt.IE_AUXILIAIRE) ;
  If Rib<>'' Then Mvt.IE_RIB:=Rib ;
  END ;
If InfoImp.SC.ForceRIBFou And (CptLuT.Nature='FOU') And (Trim(Mvt.IE_RIB)='') Then
  BEGIN
  Rib:=ChercheRibPrincipal(Mvt.IE_AUXILIAIRE) ;
  If Rib<>'' Then Mvt.IE_RIB:=Rib ;
  END ;
END ;

Function AlimCptMemoire(QFiche : TQfiche ; Var Mvt : TFMvtImport ; Var InfoImp : TInfoImport ) : Integer ;
Var CptLuT,CptLuG,CptLuS,CptLuJ : TCptLu ;
    Lefb : TFichierBase ;
    EstVentilable,EstCollectif : Boolean ;
    NatureCptGen : String ;
    i : Integer ;
BEGIN
Result:=0 ;
Fillchar(CptLuG,SizeOf(CptLuG),#0) ;
Fillchar(CptLuT,SizeOf(CptLuT),#0) ;
Fillchar(CptLuS,SizeOf(CptLuS),#0) ;
Fillchar(CptLuJ,SizeOf(CptLuJ),#0) ;
EstVentilable:=FALSE ;  EstCollectif:=FALSE ;
NatureCptGen:='' ;
If InfoImp.SC.UseCorresp Then
  BEGIN
  For i:=0 To 6 Do TraiteCorresp(i,QFiche,Mvt,InfoImp) ;
  TraiteCorresp(8,QFiche,Mvt,InfoImp) ; // Jal
  END ;
Mvt.IE_GENERAL:=BourreOuTronque(Mvt.IE_GENERAL,fbGene) ;
{$IFDEF CYBER}
If VH^.Cyber Then
  BEGIN
  If Mvt.IE_GENERAL='768100' Then Mvt.IE_GENERAL:='768800' ;
  If Mvt.IE_GENERAL='668100' Then Mvt.IE_GENERAL:='668800' ;
  END ;
{$ENDIF}
If InfoImp.SC.Majuscule Then Mvt.IE_GENERAL:=AnsiUpperCase(Mvt.IE_GENERAL) ;
Mvt.IE_AUXILIAIRE:=BourreOuTronque(Mvt.IE_AUXILIAIRE,fbAux) ;
If InfoImp.SC.Majuscule Then Mvt.IE_AUXILIAIRE:=AnsiUpperCase(Mvt.IE_AUXILIAIRE) ;
If PasDeBlanc And (Mvt.IE_AUXILIAIRE<>'') Then Mvt.IE_AUXILIAIRE:=FindEtReplace(Mvt.IE_AUXILIAIRE,' ','.',TRUE) ;
TraiteCompteSubstitue(QFiche,Mvt,InfoImp) ;
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
       If InfoImp.SC.Majuscule Then Mvt.IE_SECTION:=AnsiUpperCase(AnsiUpperCase(Mvt.IE_SECTION)) ;
      END Else If Mvt.IE_SECTION<>'' Then Mvt.IE_SECTION:=AnsiUpperCase(BourreOuTronque(Mvt.IE_SECTION,fbAxe1)) ;
   If AnalytiqueAIgnorer(Mvt,InfoImp,CptLuG) Then Result:=1 ;
   END Else Inc(InfoImp.NbGenFaux) ;

CptLuT.Cpt:=Mvt.IE_AUXILIAIRE ; CptLuT.Collectif:=Mvt.IE_GENERAL ;
If AlimLTabCptLu(1,QFiche[1],InfoImp.LAuxLu,InfoImp.ListeCptFaux,CptLuT) Then
   BEGIN
   If Not EstCollectif Then Inc(InfoImp.NbAuxFaux) Else TraiteScenarioRIB(Mvt,InfoImp,CptLuT) ;
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
//Var EnEuro : Boolean ;
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
    LigneEnCours.DE:=0 ; LigneEnCours.CE:=0 ;
//    LigneEnCours.DE:=Mvt.IE_DEBITEURO ; LigneEnCours.CE:=Mvt.IE_CREDITEURO ;
    LigneEnCours.ANouveau:=(MVT.IE_ECRANOUVEAU='H') Or (MVT.IE_ECRANOUVEAU='OAN');

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
//EnEuro:=EuroOK And VH^.TenueEuro ;
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
    OkNbEche : Boolean ;
    JalAchVen : Boolean ;
    CptLuJ : TCptLu ;
BEGIN
//If Mvt.IE_TYPEECR='L' Then Exit ;
OkIncLigne:=TRUE ;
If OkRupt Then
   BEGIN
   IdentPiece.LigneEnCours.NumLig:=0 ;
   IdentPiece.LigneEnCours.NumEche:=0 ;
   IdentPiece.LigneEnCours.NumV:=0 ;
   IdentPiece.SensDernLigneGen:=0 ;
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
     OkSens:=TRUE ; JalAchVen:=FALSE ;
     If ((Arrondi(IdentPiece.LigneEnCours.DP,V_PGI.OkDecV)=0) And (Arrondi(IdentPiece.LignePrec.DP,V_PGI.OkDecV)<>0)) Or
        ((Arrondi(IdentPiece.LigneEnCours.CP,V_PGI.OkDecV)=0) And (Arrondi(IdentPiece.LignePrec.CP,V_PGI.OkDecV)<>0)) Then OkSens:=FALSE ;
     OkNbEche:=IdentPiece.LigneEnCours.NumEche<11 ;
     If OkSens And OkNbEche Then
       OkSens:=(((Arrondi(IdentPiece.LigneEnCours.DP,V_PGI.OkDecV)>=0) And (Arrondi(IdentPiece.LignePrec.DP,V_PGI.OkDecV)>=0)) And
                ((Arrondi(IdentPiece.LigneEnCours.CP,V_PGI.OkDecV)>=0) And (Arrondi(IdentPiece.LignePrec.CP,V_PGI.OkDecV)>=0))) Or
               (((Arrondi(IdentPiece.LigneEnCours.DP,V_PGI.OkDecV)<=0) And (Arrondi(IdentPiece.LignePrec.DP,V_PGI.OkDecV)<=0)) And
                ((Arrondi(IdentPiece.LigneEnCours.CP,V_PGI.OkDecV)<=0) And (Arrondi(IdentPiece.LignePrec.CP,V_PGI.OkDecV)<=0))) ;
     If OkSens And OkNbEche Then
       BEGIN
       Fillchar(CptLuJ,SizeOf(CptLuJ),#0) ;
       CptLuJ.Cpt:=Mvt.IE_JOURNAL ;
       If ChercheCptLu(InfoImp.LJalLu,CptLuJ) Then JalAchVen:=(CptLuJ.Nature='ACH') Or (CptLuJ.Nature='VTE') ;
       END ;
     {JP 16/10/2003 : le code de VDEV est plus complet que celui de diffusion et corrige la violation de clé
                      que l'on a lorsque deux lignes consécutives ont le même compte général :
                      il n'y a pas de Inc(IdentPiece.LigneEnCours.NumLig); dans le code de Diffusion

     If (IdentPiece.LignePrec.Eche) And (IdentPiece.LigneEnCours.Eche) and
        (IdentPiece.LigneEnCours.Gen=IdentPiece.LignePrec.Gen) And JalAchVen And
        (IdentPiece.LigneEnCours.Aux=IdentPiece.LignePrec.Aux) And OkSens And (Not IdentPiece.LigneEnCours.ANouveau)
        Then OkIncLigne:=FALSE Else IdentPiece.LigneEnCours.NumEche:=0 ;}

     If (IdentPiece.LignePrec.Eche) And (IdentPiece.LigneEnCours.Eche) and
        (IdentPiece.LigneEnCours.Gen=IdentPiece.LignePrec.Gen) And JalAchVen And
        (IdentPiece.LigneEnCours.Aux=IdentPiece.LignePrec.Aux) And OkSens And (Not IdentPiece.LigneEnCours.ANouveau)
        and (IdentPiece.LigneEnCours.NumEche=IdentPiece.LignePrec.NumEche)
        Then OkIncLigne:=FALSE
        Else begin
          if IdentPiece.LigneEnCours.Eche then begin
          if (Mvt.IE_NUMECHE =-1) then IdentPiece.LigneEnCours.NumEche:=1
                                  else begin
                                    IdentPiece.LigneEnCours.NumEche:=Mvt.IE_NUMECHE;
                                    if not (IdentPiece.LigneEnCours.NumEche=IdentPiece.LignePrec.NumEche) then begin
                                      OkIncLigne:=FALSE;
                                      Inc(IdentPiece.LigneEnCours.NumLig);
                                    end;
                                  end;
            end
          else IdentPiece.LigneEnCours.NumEche:=0;
        end;
     {JP 16/10/2003 : FIN du report de code}
     
     If (InfoImp.FormatOrigine='SIS') Or RECUPSISCO Then BEGIN IdentPiece.LigneEnCours.NumEche:=0 ; OkIncLigne:=TRUE ; END ;
{$IFDEF RECUPPCL}
{$IFDEF CEGID}
     If (IdentPiece.LigneEnCours.Ana) and (Not IdentPiece.LignePrec.Ana) And
        (IdentPiece.LigneEnCours.Gen=IdentPiece.LignePrec.Gen) Then IdentPiece.LigneEnCours.NextNumVCEGID:=-1 ;
        BEGIN
        END ;
{$ENDIF}
{$ENDIF}
     If (*(IdentPiece.LignePrec.Ana) And *)(IdentPiece.LigneEnCours.Ana) and
        (IdentPiece.LigneEnCours.Gen=IdentPiece.LignePrec.Gen) Then
  //      (IdentPiece.LigneEnCours.Sect=IdentPiece.LignePrec.Sect)
        BEGIN
        OkIncLigne:=FALSE ;
        If Not VH^.RecupCegid Then
          If IdentPiece.LigneEnCours.Axe<>IdentPiece.LignePrec.Axe Then IdentPiece.LigneEnCours.NumV:=0 ;
{$IFDEF RECUPPCL}
{$IFDEF CEGID}
        If (Not IdentPiece.LignePrec.Ana) Or (IdentPiece.LigneEnCours.Axe<IdentPiece.LignePrec.Axe) Then
          BEGIN
          Inc(IdentPiece.LigneEnCours.NextNumVCegid) ;
          END ;
        IdentPiece.LigneEnCours.NumV:=IdentPiece.LigneEnCours.NextNumVCegid ;
        If IdentPiece.LigneEnCours.NumV<0 Then Inc(IdentPiece.LigneEnCours.NumV) ;
{$ENDIF}
{$ENDIF}
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

  {JP 10/09/2003 : ces lignes essaient de rationaliser tout ce qui précéde
   commentaire précédent JP 29/07/03 : If IdentPiece.LigneEnCours.Eche Then Inc(IdentPiece.LigneEnCours.NumEche) ;
                                      Cette ligne avait été mise en commentaire par Vincent Laroche cf eQualité Fiche
                                      n°12384. Cependant cela soulève d'autres problèmes car les NumEche sont tous à
                                      zéro et E_EtatLettrage à 'RI' pour les écritures lettrables au lieu de 'AL'.
                                      => revoir la chose avec Vincent Beaurenaut}

  {1/ Si Sisco : IdentPiece.LigneEnCours.Eche on met 1 sinon 0}
  if (InfoImp.FormatOrigine='SIS') or RECUPSISCO then
    Mvt.IE_NUMECHE := IdentPiece.LigneEnCours.NumEche

  {2/ Si le NumEche est renseigné dans le fichier on le reprend par défaut sauf si not IdentPiece.LigneEnCours.Eche}
  else if (Mvt.IE_NUMECHE >= 1) then begin
    if not IdentPiece.LigneEnCours.Eche then begin
      Mvt.IE_NUMECHE := 0;
      IdentPiece.LigneEnCours.NumEche := 0;
    end

    else begin
      IdentPiece.LigneEnCours.NumEche := Mvt.IE_NUMECHE;
      if (Mvt.IE_NUMECHE > 1) and
         (IdentPiece.LigneEnCours.Gen     =  IdentPiece.LignePrec.Gen    ) and
         (IdentPiece.LigneEnCours.Aux     =  IdentPiece.LignePrec.Aux    ) and
         (IdentPiece.LigneEnCours.NumEche <> IdentPiece.LignePrec.NumEche) then begin
        {Actuellement les numéro de ligne sont incrémentés : on reprend le numéro précédent}
        Dec(IdentPiece.LigneEnCours.NumLig);
        Mvt.IE_NUMLIGNE := IdentPiece.LigneEnCours.NumLig;
      end;
    end;
  end

  {3/ Si le compte et le numéro de ligne sont identiques, on est enmutli-eche}
  else if (IdentPiece.LigneEnCours.Gen    = IdentPiece.LignePrec.Gen) and
          (IdentPiece.LigneEnCours.Aux    = IdentPiece.LignePrec.Aux) and
          (IdentPiece.LigneEnCours.NumLig = IdentPiece.LignePrec.NumLig) then begin
    {En fait dans l'immédiat, ce cas ne se présentera jamais car les NumLig sont toujours différents
     cf : if (Mvt.IE_NUMECHE =-1) then IdentPiece.LigneEnCours.NumEche:=1 }
    if IdentPiece.LigneEnCours.Eche then
      Mvt.IE_NUMECHE := IdentPiece.LigneEnCours.NumEche
    else begin
      Mvt.IE_NUMECHE := 0;
      IdentPiece.LigneEnCours.NumEche := 0;
    end;
  end

  {4/ Sinon on considère que l'on est en mono echéance }
  else if IdentPiece.LigneEnCours.Eche then begin
      Mvt.IE_NUMECHE := 1;
      IdentPiece.LigneEnCours.NumEche := 1;
    end
    else begin
      IdentPiece.LigneEnCours.NumEche := 0;
      Mvt.IE_NUMECHE := 0;
    end;

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
  IdentPiece.TotDE:=Arrondi(IdentPiece.TotDE,V_PGI.okDecE) ;
  IdentPiece.TotCE:=Arrondi(IdentPiece.TotCE,V_PGI.okDecE) ;
  END ;
If Mvt.IE_TYPEECR='A' Then
  BEGIN
  IdentPiece.TotDPAna:=Arrondi(IdentPiece.TotDPAna+MVT.IE_DEBIT,V_PGI.okDecV) ;
  IdentPiece.TotCPAna:=Arrondi(IdentPiece.TotCPAna+MVT.IE_CREDIT,V_PGI.okDecV) ;
  IdentPiece.TotDDAna:=Arrondi(IdentPiece.TotDDAna+MVT.IE_DEBITDEV,IdentPiece.DecimDev) ;
  IdentPiece.TotCDAna:=Arrondi(IdentPiece.TotCDAna+MVT.IE_CREDITDEV,IdentPiece.DecimDev) ;
  IdentPiece.TotDEAna:=Arrondi(IdentPiece.TotDEAna,V_PGI.okDecE) ;
  IdentPiece.TotCEAna:=Arrondi(IdentPiece.TotCEAna,V_PGI.okDecE) ;
  END ;
END ;

procedure RetoucheFinal(EnRupture : Boolean ; Var IdentPiece : TIdentPiece ; Var Mvt : TFMvtImport) ;
Var TotEcr : Double ;
BEGIN
If (Mvt.IE_TYPEECR='L') Then BEGIN Mvt.IE_OKCONTROLE:='-' ; Mvt.IE_SELECTED:='-' ;  END ;
If IdentPiece.Doublon Then Mvt.IE_OKCONTROLE:='D' ;
if (Mvt.IE_TYPEECR<>'A') then Exit ;
If Mvt.IE_TYPEANALYTIQUE='X' Then
  BEGIN
  Mvt.IE_NUMLIGNE:=0 ; MVT.IE_POURCENTAGE:=0 ;
  If EnRupture Then Mvt.IE_TYPEMVT:='AE' Else Mvt.IE_TYPEMVT:='AL' ;
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

Function ChercheDoublon(Var InfoImp : TInfoImport ; Var Mvt : TFMvtImport ; PourLettrage : Boolean) : Integer ;
(*
Result = 0 : Mouvement non trouvé dans la base réceptrice
Result = 1 : Mouvement trouvé dans la base réceptrice
Result = 2 : Mouvement trouvé dans la base réceptrice pour la X° fois (X>1) Utilisé en lettrage uniquement
*)
Var What : Integer ;
    X    : DelInfo ;
    CleDb : String ;
    StD,StC : String ;
    QT : TQuery ;
    St,StRef : String ;
    PourCERIC : Boolean ;
    CleL : String ;
    i : Integer ;
BEGIN
Result:=0 ; PourCERIC:=(Pos('50002',V_PGI.Specif)>0) ;
What:=0 ; If Mvt.IE_TYPEANALYTIQUE='X' Then What:=1 ;
StD:=Formatfloat('###0.00',Mvt.IE_DEBITDEV) ; StD:=FindEtReplace(StD,',','.',TRUE) ;
StC:=Formatfloat('###0.00',Mvt.IE_CREDITDEV) ; StC:=FindEtReplace(StC,',','.',TRUE) ;
If PourLettrage Then
  BEGIN
  St:='SELECT E_JOURNAL,E_EXERCICE,E_DATECOMPTABLE,E_NUMEROPIECE,'+
      'E_NATUREPIECE,E_NUMLIGNE,E_NUMECHE, E_GENERAL,E_AUXILIAIRE,E_DATEREFEXTERNE,'+
      'E_REFINTERNE,E_REFLIBRE,E_REFEXTERNE,E_LIBELLE,E_TAUXDEV,E_DEBIT,E_CREDIT,'+
      'E_DEBITDEV,E_CREDITDEV,E_LETTRAGE,E_DATEECHEANCE,E_DEVISE,E_REFINTERNE,E_EDITEETATTVA,' ;
  St:=St+
      ' FROM ECRITURE WHERE E_JOURNAL="'+Mvt.IE_JOURNAL+'" AND '+
      'E_EXERCICE="'+QUELEXODT(Mvt.IE_DATECOMPTABLE)+'" AND E_DATECOMPTABLE="'+USDATETIME(Mvt.IE_DATECOMPTABLE)+'" AND E_GENERAL="'+Mvt.IE_GENERAL+'" AND '+
      'E_AUXILIAIRE="'+Mvt.IE_AUXILIAIRE+'" AND E_REFINTERNE="'+TrimSpe(Mvt.IE_REFINTERNE)+'" AND E_LIBELLE="'+TrimSpe(Mvt.IE_LIBELLE)+'" AND E_DEVISE="'+Mvt.IE_DEVISE+'" AND E_DEBITDEV="'+STD+'" AND '+
      'E_CREDITDEV="'+STC+'" AND E_NATUREPIECE="'+TrimSpe(Mvt.IE_NATUREPIECE)+'" AND E_QUALIFPIECE="N" ' ;
  END Else
  BEGIN
  If Not PourCeric Then
    BEGIN
    Case What Of
      0 : StRef:=' AND E_REFINTERNE="'+TrimSpe(Mvt.IE_REFINTERNE)+'" ' ;
      1 : StRef:=' AND Y_REFINTERNE="'+TrimSpe(Mvt.IE_REFINTERNE)+'" ' ;
      END ;
    END Else StRef:='' ;
  Case What Of
    0 : St:='SELECT E_JOURNAL, E_EXERCICE, E_DATECOMPTABLE, E_NUMEROPIECE, E_NUMLIGNE, E_NUMECHE, E_QUALIFPIECE, E_REFINTERNE '+
            'FROM ECRITURE WHERE E_JOURNAL="'+Mvt.IE_JOURNAL+'" AND E_EXERCICE="'+QUELEXODT(Mvt.IE_DATECOMPTABLE)+'" AND E_DATECOMPTABLE="'+USDATETIME(Mvt.IE_DATECOMPTABLE)+'" AND E_GENERAL="'+Mvt.IE_GENERAL+'" AND '+
            'E_AUXILIAIRE="'+Mvt.IE_AUXILIAIRE+'" '+StRef+' AND E_LIBELLE="'+TrimSpe(Mvt.IE_LIBELLE)+'" AND E_DEVISE="'+Mvt.IE_DEVISE+'" AND E_DEBITDEV="'+STD+'" AND '+
            'E_CREDITDEV="'+STC+'"' ;
    1 : St:='SELECT Y_JOURNAL, Y_EXERCICE, Y_DATECOMPTABLE, Y_NUMEROPIECE, Y_NUMLIGNE, Y_NUMVENTIL, Y_QUALIFPIECE, Y_REFINTERNE '+
            'FROM ANALYTIQ WHERE Y_JOURNAL="'+Mvt.IE_JOURNAL+'" AND Y_EXERCICE="'+QUELEXODT(Mvt.IE_DATECOMPTABLE)+'" AND Y_DATECOMPTABLE="'+USDATETIME(Mvt.IE_DATECOMPTABLE)+'" AND Y_GENERAL="'+Mvt.IE_GENERAL+'" AND '+
            'Y_SECTION="'+Mvt.IE_SECTION+'" '+StRef+' AND Y_LIBELLE="'+TrimSpe(Mvt.IE_LIBELLE)+'" AND Y_DEVISE="'+Mvt.IE_DEVISE+'" AND Y_DEBITDEV="'+STD+'" AND '+
            'Y_CREDITDEV="'+STC+'" AND Y_AXE="'+Mvt.IE_AXE+'" ' ;
    Else Exit ;
    END ;
  END ;
QT:=OpenSQL(St,TRUE) ;
If Not QT.Eof Then Result:=1 ;
If (Not PourLettrage) And (Result=1) Then
  BEGIN
  X:=DelInfo.Create ; X.LeCod:=Mvt.IE_Journal ; X.LeLib:=IntToStr(MVT.IE_NumPiece);
  X.LeMess:=TraduireMemoire('N° Pièce')+' : '+IntToStr(QT.Fields[3].AsInteger)+' - '+TraduireMemoire('Date')+' '+DateToStr(QT.Fields[2].AsDateTime)+
            ' - '+TraduireMemoire('Référence')+' '+QT.Fields[7].AsString ;
  InfoImp.CRListeEnteteDoublon.Add(X) ;
  CleDB:=Format_String(Mvt.IE_Journal,3)+Format_String(' ',1)+FormatFloat('00000000',MVT.IE_NumPiece) ;
  InfoImp.ListeEnteteDoublon.Add(CleDB) ;
  END ;
If PourLettrage And (Result=1) Then
  BEGIN
  Mvt.IE_GENERAL:=QT.FindField('E_GENERAL').AsString ;
  Mvt.IE_AUXILIAIRE:=QT.FindField('E_AUXILIAIRE').AsString ;
  Mvt.IE_DATECOMPTABLE:=QT.FindField('E_DATECOMPTABLE').AsDateTime ;
  Mvt.IE_DATEECHEANCE:=QT.FindField('E_DATEECHEANCE').AsDateTime ;
  Mvt.IE_DATEREFEXTERNE:=QT.FindField('E_DATEREFEXTERNE').AsDateTime ;
  Mvt.IE_REFLIBRE:=QT.FindField('E_REFLIBRE').AsString ;
  Mvt.IE_REFINTERNE:=QT.FindField('E_REFINTERNE').AsString ;
  Mvt.IE_REFEXTERNE:=QT.FindField('E_REFEXTERNE').AsString ;
  Mvt.IE_LIBELLE:=QT.FindField('E_LIBELLE').AsString ;
  Mvt.IE_JOURNAL:=QT.FindField('E_JOURNAL').AsString ;
  Mvt.IE_NUMPIECE:=QT.FindField('E_NUMEROPIECE').AsInteger ;
  Mvt.IE_NUMLIGNE:=QT.FindField('E_NUMLIGNE').AsInteger ;
  Mvt.IE_NUMECHE:=QT.FindField('E_NUMECHE').AsInteger ;
  Mvt.IE_LETTRAGE:=QT.FindField('E_LETTRAGE').AsString ;
  Mvt.IE_TAUXDEV:=QT.FindField('E_TAUXDEV').AsFloat ;
  Mvt.IE_DEVISE:=QT.FindField('E_DEVISE').AsString ;
  Mvt.IE_DEBIT:=QT.FindField('E_DEBIT').AsFloat ;
  Mvt.IE_CREDIT:=QT.FindField('E_CREDIT').AsFloat ;
  Mvt.IE_DEBITDEV:=QT.FindField('E_DEBITDEV').AsFloat ;
  Mvt.IE_CREDITDEV:=QT.FindField('E_CREDITDEV').AsFloat ;
  Mvt.IE_NATUREPIECE:=QT.FindField('E_NATUREPIECE').AsString ;
  Mvt.IE_LIBREBOOL1:=QT.FindField('E_EDITEETATTVA').AsString ;
  If InfoImp.ListeMvtTrouve<>NIL Then
    BEGIN
    CleL:=Format_String(Mvt.IE_LETTRAGE,4)+
          Format_String(QT.FindField('E_EXERCICE').AsString,3)+
          Format_String(Mvt.IE_JOURNAL,3)+
                    DateToStr(Mvt.IE_DATECOMPTABLE)+
          FormatFloat('0000000000',Mvt.IE_NUMPIECE)+
          FormatFloat('00000',Mvt.IE_NUMLIGNE)+
          FormatFloat('00',Mvt.IE_NUMECHE) ;
    If InfoImp.ListeMvtTrouve.Find(CleL,i) Then Result:=2 Else InfoImp.ListeMvtTrouve.Add(CleL) ;
    END ;
  END ;
QT.Close ;


(*
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
*)
END ;

Procedure RetoucheJalCorresp(Var Enr : tStLue ; Var InfoImp : TInfoImport ; QFiche : TQfiche) ;
Var LeJal : tCptImport ;
BEGIN
If InfoImp.SC.CorrespJal Then
  BEGIN
  Fillchar(LeJal,SizeOf(LeJal),#0) ; LeJal.Cpt:=Enr.Jal ;
  If TraiteCorrespCpt(8,QFiche,LeJal,InfoImp) Then Enr.Jal:=LeJal.Cpt ;
  END ;
END ;

Procedure RetoucheGenCorresp(Var Enr : tStLue ; Var InfoImp : TInfoImport ; QFiche : TQfiche) ;
Var LeGen : tCptImport ;
BEGIN
If InfoImp.SC.CorrespGen Then
  BEGIN
  Fillchar(LeGen,SizeOf(LeGen),#0) ; LeGen.Cpt:=Enr.General ;
  If TraiteCorrespCpt(0,QFiche,LeGen,InfoImp) Then Enr.General:=LeGen.Cpt ;
  END ;
END ;


Function AnalytiqueAbsenteSurLignePrecedente(Var St,OldSt : String ; Var InfoImp : TInfoImport ; QFiche : TQfiche; Axe : Integer = 0) : Boolean ;
Var FmtFic : Integer ;
    Enr,OldEnr : TStLue ;
    Cpt,Section : String ;
    CptLu : TCptLu ;
    AnalAbsente : Boolean ;
    Decal,LgCpte : Integer ;
    szAxe : String;
BEGIN
Result:=FALSE ; If OldSt='' Then Exit ; Decal:=0 ;
if InfoImp.Format='SAA' then FmtFic:=0 else if InfoImp.Format='SN2' then FmtFic:=0 else
if InfoImp.Format='HLI' then FmtFic:=1 else if InfoImp.Format='HAL' then FmtFic:=2 else
if InfoImp.Format='CGN' then FmtFic:=3 else if InfoImp.Format='CGE' then FmtFic:=4 Else Exit ;
AlimEnr(St,Enr,FmtFic) ; AlimEnr(OldSt,OldEnr,FmtFic) ;
RetoucheJalCorresp(Enr,InfoImp,QFiche) ; RetoucheJalCorresp(OldEnr,InfoImp,QFiche) ;
RetoucheGenCorresp(Enr,InfoImp,QFiche) ; RetoucheGenCorresp(OldEnr,InfoImp,QFiche) ;
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
    Section:=VH^.Cpta[fbAxe1].Attente ;
    If InfoImp.Sc.RSect1<>'' Then Section:=InfoImp.Sc.RSect1 ;
    szAxe := 'A1';
    case Axe of
      //1: begin Section:=VH^.Cpta[fbAxe1].Attente; if InfoImp.Sc.RSect1<>'' Then Section:=InfoImp.Sc.RSect1;  szAxe := 'A1'; end;
      2: begin Section:=VH^.Cpta[fbAxe2].Attente; if InfoImp.Sc.RSect2<>'' Then Section:=InfoImp.Sc.RSect2; szAxe := 'A2'; end;
      3: begin Section:=VH^.Cpta[fbAxe3].Attente; if InfoImp.Sc.RSect3<>'' Then Section:=InfoImp.Sc.RSect3; szAxe := 'A3'; end;
      4: begin Section:=VH^.Cpta[fbAxe4].Attente; if InfoImp.Sc.RSect4<>'' Then Section:=InfoImp.Sc.RSect4; szAxe := 'A4'; end;
      5: begin Section:=VH^.Cpta[fbAxe5].Attente; if InfoImp.Sc.RSect5<>'' Then Section:=InfoImp.Sc.RSect5; szAxe := 'A5'; end;
    end;
    Case FmtFic Of
      0   : BEGIN
            OldSt:=Insere(OldSt,'A',25,1) ;
            OldSt:=Insere(OldSt,Format_String(Section,13),26,13) ;
            END ;
      1   : BEGIN
            If Not VH^.ImportRL Then Inc(Decal,4) ;
            OldSt:=Insere(OldSt,'A',27+Decal,1) ;
            OldSt:=Insere(OldSt,Format_String(Section,LgCpte),28+Decal,LgCpte) ;
            END ;
      2   : BEGIN
            If Not VH^.ImportRL Then Inc(Decal,4) ;
            OldSt:=Insere(OldSt,'A',27+Decal,1) ;
            OldSt:=Insere(OldSt,Format_String(Section,LgCpte),28+Decal,LgCpte) ;
            If Not VH^.ImportRL Then Inc(Decal,4) ;
            OldSt:=Insere(OldSt,szAxe,117+Decal,2) ;
            END ;
      3,4 : BEGIN
            OldSt:=Insere(OldSt,'A',31,1) ;
            OldSt:=Insere(OldSt,Format_String(Section,17),32,17) ;
            OldSt:=Insere(OldSt,szAxe,219,2) ;
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
    {DebitELu,CreditELu,}DebitPLu,CreditPLu : Double ;
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
StSQL:='SELECT IE_DEBIT,IE_CREDIT FROM IMPECR WHERE IE_CHRONO='+IntToStr(LeChrono) ;
Q:=OpenSQL(StSQL,TRUE) ;
If Not Q.Eof Then
  BEGIN
  DebitPLu:=Q.Fields[0].AsFloat ; CreditPLu:=Q.Fields[1].AsFloat ;
  Ferme(Q) ;
  Case QuiOppose Of
    'D' : BEGIN
          AjusteSQLMontant(What,V_PGI.OkDecV,'IE_DEBIT','IE_CREDIT',DeltaOppose,DebitPLu,CreditPLu,StDebit,StCredit) ;
          //AjusteSQLMontant(What,V_PGI.OkDecE,'IE_DEBITEURO','IE_CREDITEURO',DeltaOppose2,DebitELu,CreditELu,StDebitEuro,StCreditEuro) ;
          END ;
    'P' : BEGIN
          AjusteSQLMontant(What,V_PGI.OkDecV,'IE_DEBIT','IE_CREDIT',DeltaOppose,DebitPLu,CreditPLu,StDebit,StCredit) ;
          If StDebit<>'' Then StDebitDev:='IE_DEBITDEV='+StrfPoint(DebitPLu) ;
          If StCredit<>'' Then StCreditDev:='IE_CREDITDEV='+StrfPoint(CreditPLu) ;
          END ;
    'E' : BEGIN
          //AjusteSQLMontant(What,V_PGI.OkDecE,'IE_DEBITEURO','IE_CREDITEURO',DeltaOppose,DebitELu,CreditELu,StDebitEuro,StCreditEuro) ;
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

Procedure UpdatePourControle(Var IdentPiece : TIdentPiece) ;
Var LeChrono : Integer ;
    StSQL : String ;
BEGIN
If IdentPiece.DernChronoENonVentilable<>0 Then LeChrono:=IdentPiece.DernChronoENonVentilable
                                          Else LeChrono:=IdentPiece.DernChronoE ;
StSQL:='UPDATE IMPECR SET IE_CONTROLE="X" WHERE IE_CHRONO='+IntToStr(LeChrono) ;
ExecuteSQL(StSQL) ;
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
    SurODA,SurAN,PbEquilibreSaisi : Boolean ;
    NoEcart : Boolean ;
BEGIN
Result:=FALSE ; SurAN:=FALSE ; //PbEquilibreSaisi:=FALSE ; UpdateImpEcr:=FALSE ; 
//If IdentPiece.Date<V_PGI.DateDebutEuro Then Exit ;
EnEuro:=EuroOK And VH^.TenueEuro ;
Cpt:=Trim(IdentPiece.JalP) ; Fillchar(CptLu,SizeOf(CptLu),#0) ; CptLu.Cpt:=Cpt ; SurODA:=FALSE ;
If ChercheCptLu(InfoImp.LJalLu,CptLu) Then
  BEGIN
  SurODA:=(CptLu.Nature='ODA') ; SurAN:=(CptLu.Nature='ANO') ;
  END ;
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
PbEquilibreSaisi:=((QuiOppose in ['P','E']) And (DeltaSaisi<>0)) Or
                  ((QuiOppose in ['D']) And (DeltaSaisi<>0)) ;
UpdateImpecr:=Pb And ((IdentPiece.DateP<V_PGI.DateDebutEuro) Or (QuiOppose='D') Or SurODA Or SurAN) ;
If Not VH^.Cyber Then If PbEquilibreSaisi And (Not SurODA) Then UpdatePourControle(IdentPiece) ;
If UpdateImpEcr Or VH^.RecupSISCOPGI Then
  BEGIN
  If SurODA Then UpdatePourEquilibre(TRUE,TRUE,EnEuro,QuiOppose,DeltaOppose,DeltaOppose2,IdentPiece)
            Else UpdatePourEquilibre(FALSE,FALSE,EnEuro,QuiOppose,DeltaOppose,DeltaOppose2,IdentPiece) ;
  END Else
  BEGIN
  NoEcart:=FALSE ;
  If Not (ctxPCL in V_PGI.PGIContexte) Then
    BEGIN
    If EstSpecif('51188') Then
      BEGIN
      If StopEcart(IdentPiece.DateP) Then NoEcart:=TRUE ;
      END ;
    END ;
  If NoEcart Then Pb:=FALSE ;
  If Pb Then
    BEGIN
    Result:=TRUE ;
    MVT.IE_ELEMENTARECUPERER:=TRUE ;
    Mvt.IE_DEBIT:=0 ; Mvt.IE_CREDIT:=0 ;
    //Mvt.IE_DEBITEURO:=0 ; MVT.IE_CREDITEURO:=0 ;
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
            If DeltaOppose>0 Then
              BEGIN
              Mvt.IE_CREDIT:=DeltaOppose ; Mvt.IE_CREDITDEV:=Mvt.IE_CREDIT ;
              END Else
              BEGIN
              Mvt.IE_DEBIT:=Abs(DeltaOppose) ; Mvt.IE_DEBITDEV:=Mvt.IE_DEBIT ;
              END ;
            //Mvt.IE_GENERAL:=VH^.EccEuroDebit ;
            //If DeltaOppose>0 Then Mvt.IE_GENERAL:=VH^.EccEuroCredit ;
            END ;
      'E' : BEGIN
            If DeltaOppose>0 Then
               BEGIN
               //Mvt.IE_CREDITEURO:=DeltaOppose ;
               Mvt.IE_CREDITDEV:=Mvt.IE_CREDIT ;
               END Else
               BEGIN
               //Mvt.IE_DEBITEURO:=Abs(DeltaOppose) ;
               Mvt.IE_CREDITDEV:=Mvt.IE_CREDIT ;
               END ;
            //Mvt.IE_GENERAL:=VH^.EccEuroDebit ;
            //If DeltaOppose>0 Then Mvt.IE_GENERAL:=VH^.EccEuroCredit ;
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
    MVT.IE_LIBREBOOL0:='-' ; //MVT.IE_LIBREBOOL1:='-' ;
    Mvt.IE_TYPEECR:='E' ; Mvt.IE_DEVISE:=V_PGI.DevisePivot ; Mvt.IE_TAUXDEV:=1 ;
    If (VH^.RecupPCL And (CptLu.ModeSaisie='LIB')) Or (ModeJal(Mvt.IE_JOURNAL,InfoImp)=mLib) Then
      BEGIN
      Mvt.IE_DATECOMPTABLE:=FinDeMois(Mvt.IE_DATECOMPTABLE) ;
      END ;
    END ;
  END ;
END ;

Function VerifEquilibrageAnalytique(Var St,OldSt : String ; Var IdentPiece : TIdentPiece ;
                                    Var InfoImp : TInfoImport ; Var Mvt : TFMvtImport  ; QFiche : TQfiche) : Boolean ;
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
//FmtFic:=0 ;
if InfoImp.Format='SAA' then FmtFic:=0 else if InfoImp.Format='SN2' then FmtFic:=0 else
if InfoImp.Format='HLI' then FmtFic:=1 else if InfoImp.Format='HAL' then FmtFic:=2 else
if InfoImp.Format='CGN' then FmtFic:=3 else if InfoImp.Format='CGE' then FmtFic:=4 Else Exit ;
AlimEnr(OldSt,OldEnr,FmtFic) ; If St<>'' Then AlimEnr(St,Enr,FmtFic) ;
RetoucheJalCorresp(Enr,InfoImp,QFiche) ; RetoucheJalCorresp(OldEnr,InfoImp,QFiche) ;
RetoucheGenCorresp(Enr,InfoImp,QFiche) ; RetoucheGenCorresp(OldEnr,InfoImp,QFiche) ;
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
    Dev : String ;
BEGIN
EnEuro:=EuroOK And VH^.TenueEuro ;
CodeMontant:='F--' ; TauxDev:=0 ; Decal:=0 ; If Not VH^.ImportRL Then Decal:=8 ; //DateC:=0 ;
If (InfoImp.Format='CGN') Or (InfoImp.Format='CGE') Then
  BEGIN
  CodeMontant:=Trim(Copy(St,173,3)) ; If (CodeMontant='') Or (CodeMontant='---') Then CodeMontant:='F--' ;
  IdentPiece.CodeMontant:=CodeMontant ;
  CodeMontant:=OnForceLaDevise(IdentPiece,InfoImp) ;
  if Trim(Copy(St,163,10))<>'' then TauxDev:=Valeur(StPoint(Copy(St,163,10)));
  DateC:=Format_Date_HAL(Copy(St,4,8)) ;
  END Else
If InfoImp.Format='HAL' Then
  BEGIN
  if Trim(Copy(St,305+Decal,20))<>'' then TauxDev:=Valeur(StPoint(Copy(St,305+Decal,20)));
  If EnEuro Then CodeMontant:='E--' Else CodeMontant:='F--' ;
  if Trim(Copy(St,262+Decal,3))<>'' then IF Trim(Copy(St,262+Decal,3))<>V_PGI.DevisePivot Then CodeMontant:='D--' ;
  DateC:=Format_Date_HAL(Copy(St,4,8)) ;
  IdentPiece.CodeMontant:=CodeMontant ;
  CodeMontant:=OnForceLaDevise(IdentPiece,InfoImp) ;
  END Else
If InfoImp.Format='HLI'  Then
  BEGIN
  If EnEuro Then CodeMontant:='E--' Else CodeMontant:='F--' ;
  DateC:=Format_Date_HAL(Copy(St,4,8)) ;
  IdentPiece.CodeMontant:=CodeMontant ;
  CodeMontant:=OnForceLaDevise(IdentPiece,InfoImp) ;
  END Else
  BEGIN
  If EnEuro Then CodeMontant:='E--' Else CodeMontant:='F--' ;
  If (InfoImp.Format='SN2') And Sn2Orli Then // Format Orli Avec Devise : Cas non prévu sur le format SAARI standard
    BEGIN
    Dev:=Trim(Copy(St,133,3)) ; If (Dev<>'') And (Dev<>V_PGI.DevisePivot) Then CodeMontant:='D--' ;
    END ;
  DateC:=Format_Date(Copy(St,4,6)) ;
  IdentPiece.CodeMontant:=CodeMontant ;
  CodeMontant:=OnForceLaDevise(IdentPiece,InfoImp) ;
  END ;
IdentPiece.CodeMontant:=CodeMontant ;
If TauxDev<>0 Then
  BEGIN
  If DateC>=V_PGI.DateDebutEuro Then TauxDev:=TauxDev*V_PGI.TauxEuro ;
  END ;
IdentPiece.TauxDev:=TauxDev ;
END ;
(*
{$IFDEF RECUPPCL}
Function EstUneLigneEcart(St : String) : Boolean ;
Var Cpt,CptD,CptC : String ;
BEGIN
Result:=FALSE ;
Cpt:=Trim(Copy(St,14,17)) ;
Cpt:=BourreOuTronque(Cpt,fbGene) ;
CptD:=GetParamSoc('SO_ECCEURODEBIT') ;
CptC:=GetParamSoc('SO_ECCEUROCREDIT') ;
If (Cpt=CptD) Or (Cpt=CptC) Then Result:=TRUE ;
END ;
{$ENDIF}
*)
Function MontantNul(Var St : String ; Var InfoImp : TInfoImport) : Boolean ;
Var FmtFic,Decal : Integer ;
    Montant : Double ;
    TypeMvt : String ;
{$IFDEF CYBER}
    Cpt : String ;
{$ENDIF}
BEGIN
Result:=FALSE ; // Exit ; Pourquoi ce exit ???
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
If (Not result) And (InfoImp.Format='SN2') Then // Pour shunter les ligne 'E' sur le format SAARI négoce V2
  BEGIN
  TypeMvt:=Copy(St,25,1) ; If (TypeMvt='X') then Result:=TRUE ;
  END ;
{$IFDEF CYBER}
If Result Then
  BEGIN
  Cpt:=Trim(Copy(St,14,17)) ; Cpt:=BourreLaDonc(Cpt,fbGene) ;
//  If (Cpt=VH^.EccEuroDebit) Or (Cpt=VH^.EccEuroCredit) Then Result:=FALSE ;
  If Cpt='768100' Then Result:=FALSE ;
  If Cpt='668100' Then Result:=FALSE ;
  END ;
{$ENDIF}

{$IFDEF RECUPPCL}
//If Not Result Then Result:=EstUneLigneEcart(St)  ;
{$ENDIF}
END ;

Function ImportLigne(St : String ; Var OldSt : String ; QFiche : TQFiche ;
                     Var InfoImp : TInfoImport ; Var IdentPiece,OldIdentPiece : TIdentPiece ; Var PremFois : Boolean ; FinDeTraitement,PbAnaSurLastLigne : Boolean)  : Boolean ;
var EnRupture,DetecteRupture : Boolean ;
    AnalSurUneLigne : Boolean ;
    AnalAInserer : Boolean ;
    EcartDeConvertionAGenerer : Boolean ;
    SauveSt : String ;
    SauveSt1, SauveSt2 : String ;
    OkRetoucheMontant : Boolean ;
    iAxe : Integer;
 Label 0,1;
BEGIN

iAxe := -1;
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
  DetecteRupture:=DetecteUneRupture(St,IdentPiece,InfoImp,QFiche) ;
  If FinDeTraitement Then BEGIN OldSt:=St ; St:='' ; END ;
  VerifEquilibrageAnalytique(St,OldSt,IdentPiece,InfoImp,MvtImport^,QFiche) ;
  If FinDeTraitement Then BEGIN St:=OldSt ; OldSt:='' ; END ;
  END ;
SauveSt1:=St ;
if (InfoImp.Format='RAP') or (InfoImp.Format='CRA') or (Trim(InfoImp.Format)='MP')  Then Else
  BEGIN
  If FinDeTraitement Or DetecteRupture Then
    BEGIN
    EcartDeConvertionAGenerer:=EcartDeConversionSurPiecePrecedente(IdentPiece,InfoImp,MvtImport^) ;
    If FinDeTraitement And (Not EcartDeConvertionAGenerer) And (Not PbAnaSurLastLigne) Then Exit ;
    END ;
  END ;
SauveSt:=St;
SauveSt2 := OldSt; // Pour récupérer la ligne sans les paramètres analytiques
If AnalytiqueAbsenteSurLignePrecedente(St,OldSt,InfoImp,QFiche) Or PbAnaSurLastLigne Then
  BEGIN
  If EcartDeConvertionAGenerer Then MvtImport^.IE_ElementARecuperer:=FALSE ;
  AnalAInserer:=TRUE ;
  iAxe := 1;
  If Not PbAnaSurLastLigne Then St:=OldSt ;
  END ;
0:
If (PremFois Or DetecteRupture) And (Not EcartDeConvertionAGenerer) Then ReAlimCodeMontant(St,InfoImp,IdentPiece) ;
PremFois:=FALSE ;
if (InfoImp.Format='HAL') or (InfoImp.Format='HLI') then ImportLigneHal(St,InfoImp,MvtImport^) Else
  If (InfoImp.Format='CGN') Or (InfoImp.Format='CGE') Then ImportLigneCegid(St,InfoImp,MvtImport^,AnalSurUneLigne,IdentPiece) Else
     ImportLigneAutre(St,InfoImp,MvtImport^) ;
InitAxe(St,InfoImp.Format,MvtImport^) ;
If AlimCptMemoire(QFiche,MvtImport^,InfoImp)=1 Then begin
  if AnalAInserer then goto 1
                  else Exit; // Ligne analytique sur compte non ventilable : on ignore
end;
OkRetoucheMontant:=True ; If (EcartDeConvertionAGenerer) Then If Not AnalAInserer Then OkRetoucheMontant:=FALSE ;
if (InfoImp.Format='HAL') or (InfoImp.Format='HLI') then RetoucheEnrHal(St,IdentPiece,InfoImp,MvtImport^,OkRetoucheMontant,EcartDeConvertionAGenerer,QFiche) Else
  If (InfoImp.Format='CGN') Or (InfoImp.Format='CGE') Then RetoucheEnrCegid(St,IdentPiece,InfoImp,MvtImport^,OkRetoucheMontant,QFiche) Else
     RetoucheEnrAutre(St,IdentPiece,InfoImp,MvtImport^,OkRetoucheMontant,QFiche) ;
{ A FAIRE : }
if (InfoImp.Format='RAP') and (Copy(St,133,1)='0') then Exit ; // IE_REFPOINTAGE:=Copy(St,133,1) ;
RetoucheODANAL(MvtImport^) ;
AlimIdentPiece(IdentPiece,MvtImport^) ;
EnRupture:=TestRupturePiece(InfoImp.Format,TRUE,IdentPiece,OldIdentPiece,InfoImp) ;
If EnRupture And (MvtImport^.IE_TYPEECR<>'L') Then
   BEGIN
   Inc(InfoImp.NbPiece) ; Result:=TestBreak ;
   IdentPiece.Doublon:=FALSE ;
   If InfoImp.CtrlDB Then If ChercheDoublon(InfoImp,MvtImport^,FALSE)=1 Then IdentPiece.Doublon:=TRUE ;
   If Result Then Exit ;
   END ;
AlimNumeros(EnRupture,IdentPiece,InfoImp,MvtImport^) ;
RetoucheFinal(EnRupture,IdentPiece,MvtImport^) ;
OldSt:=St ;
AjouteMvt(InfoImp) ;
If AnalSurUneLigne Then
BEGIN
   Goto 0 ;
END ;
1:
If AnalAInserer Then
   BEGIN
   If EcartDeConvertionAGenerer Then
     BEGIN
     EcartDeConversionSurPiecePrecedente(IdentPiece,InfoImp,MvtImport^) ;
     St:=SauveSt1 ;
     end;
   // FQ 15130
   inc(iAxe);
   if iAxe = 6 then begin
     iAxe := -1; St:=SauveSt;
     AnalAInserer:=FALSE ;  If Trim(St)='' Then Exit;
     end
   else begin
     OldSt := SauveSt2; // Répart sur une ligne sans analytique
     St := SauveSt2;
     if AnalytiqueAbsenteSurLignePrecedente(St,OldSt,InfoImp,QFiche, iAxe) then St := OldSt
       else St := SauveSt;
   end;
   If PbAnaSurLastLigne And (Not EcartDeConvertionAGenerer) Then Exit ;
   Goto 0 ;
   END ;
If Not (FinDeTraitement) And EcartDeConvertionAGenerer Then
   BEGIN
   EcartDeConvertionAGenerer:=FALSE ; If Trim(St)='' Then Exit ;
   MvtImport^.IE_ElementARecuperer:=FALSE ;
   Goto 0 ;
   END ;
END ;

Procedure ReInitQ(Q : TQuery ; N : Integer ; Var InfoImp : TInfoImport) ;
BEGIN
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
//if (InfoImp.Lequel<>'FBE') and (InfoImp.Lequel<>'FOD') and (Q.FindField('IE_NUMLIGNE').AsInteger=0) then BEGIN Inc(NumLigne) ; Q.FindField('IE_NUMLIGNE').AsInteger:=NumLigne ; END ;
END ;

Procedure QToMvt(Q : TQuery ; Var Mvt : TFMvtImport) ;
BEGIN
(*
Mvt.:=Q.FindField('').AsString ;
Mvt.:=Q.FindField('').AsDateTime ;
Mvt.:=Q.FindField('').AsFloat ;
Mvt.:=Q.FindField('').AsInteger ;
*)
Mvt.IE_JOURNAL:=Trim(Q.FindField('IE_JOURNAL').AsString) ;
Mvt.IE_GENERAL:=Q.FindField('IE_GENERAL').AsString ;
Mvt.IE_AUXILIAIRE:=Q.FindField('IE_AUXILIAIRE').AsString ;
Mvt.IE_SECTION:=Q.FindField('IE_SECTION').AsString ;
If Trim(Q.FindField('IE_NATUREPIECE').AsString)<>'' Then Mvt.IE_NATUREPIECE:=Q.FindField('IE_NATUREPIECE').AsString ;
If Trim(Q.FindField('IE_QUALIFPIECE').AsString)<>'' Then Mvt.IE_QUALIFPIECE:=Q.FindField('IE_QUALIFPIECE').AsString ;
Mvt.IE_DATECOMPTABLE:=Q.FindField('IE_DATECOMPTABLE').AsDateTime ;
Mvt.IE_NUMPIECE:=Q.FindField('IE_NUMPIECE').AsInteger ;
Mvt.IE_DEBIT:=Q.FindField('IE_DEBIT').AsFloat ;
Mvt.IE_CREDIT:=Q.FindField('IE_CREDIT').AsFloat ;
Mvt.IE_DEBITDEV:=Q.FindField('IE_DEBITDEV').AsFloat ;
Mvt.IE_CREDITDEV:=Q.FindField('IE_CREDITDEV').AsFloat ;
If Trim(Q.FindField('IE_ECHE').AsString)<>'' Then Mvt.IE_ECHE:=Q.FindField('IE_ECHE').AsString ;
If Trim(Q.FindField('IE_ANA').AsString)<>'' Then Mvt.IE_ECHE:=Q.FindField('IE_ANA').AsString ;
If Q.FindField('IE_NUMECHE').AsInteger<>0 Then Mvt.IE_NUMECHE:=Q.FindField('IE_NUMECHE').AsInteger ;
Mvt.IE_TYPEECR:=Q.FindField('IE_TYPEECR').AsString ;
Mvt.IE_TYPEANALYTIQUE:=Q.FindField('IE_TYPEANALYTIQUE').AsString ;
Mvt.IE_AXE:=Q.FindField('IE_AXE').AsString ;
END ;

Procedure MvtToQ(Var Mvt : TFMvtImport ; Q : TQuery ; OnBudget : Boolean) ;
BEGIN
Q.FindField('IE_DEBIT').AsFloat:=Mvt.IE_DEBIT ;
Q.FindField('IE_CREDIT').AsFloat:=Mvt.IE_CREDIT ;
Q.FindField('IE_DEBITDEV').AsFloat:=Mvt.IE_DEBITDEV ;
Q.FindField('IE_CREDITDEV').AsFloat:=Mvt.IE_CREDITDEV ;
Q.FindField('IE_TAUXDEV').AsFloat:=Mvt.IE_TAUXDEV ;
Q.FindField('IE_COTATION').AsFloat:=Mvt.IE_COTATION ;
Q.FindField('IE_ECHE').AsString:=Mvt.IE_ECHE ;
Q.FindField('IE_ANA').AsString:=Mvt.IE_ANA ;
Q.FindField('IE_NUMECHE').AsInteger:=Mvt.IE_NUMECHE ;
If Not OnBudget Then Q.FindField('IE_NUMPIECE').AsInteger:=Mvt.IE_NUMPIECE ;
If OnBudget Then Q.FindField('IE_NUMLIGNE').AsInteger:=0
            Else Q.FindField('IE_NUMLIGNE').AsInteger:=Mvt.IE_NUMLIGNE ;
Q.FindField('IE_NUMVENTIL').AsInteger:=Mvt.IE_NUMVENTIL ;
Q.FindField('IE_NUMECHE').AsInteger:=Mvt.IE_NUMVENTIL ;
Q.FindField('IE_ECRANOUVEAU').AsString:=Mvt.IE_ECRANOUVEAU ;
Q.FindField('IE_OKCONTROLE').AsString:=Mvt.IE_OKCONTROLE ;
Q.FindField('IE_SELECTED').AsString:=Mvt.IE_SELECTED ;
Q.FindField('IE_ECRANOUVEAU').AsString:=Mvt.IE_ECRANOUVEAU ;
Q.FindField('IE_REGIMETVA').AsString:=Mvt.IE_REGIMETVA ;
If OnBudget Then Q.FindField('IE_QUALIFPIECE').AsString:='N'
            Else Q.FindField('IE_QUALIFPIECE').AsString:=Mvt.IE_QUALIFPIECE ;
Q.FindField('IE_TVA').AsString:=Mvt.IE_TVA ;
Q.FindField('IE_TPF').AsString:=Mvt.IE_TPF ;
END ;

function ImporteFormatParam(Var Fichier : TextFile ; QFiche : TQFiche ; Var IdentPiece,OldIdentPiece : TIdentPiece ; Var InfoImp : TInfoImport ; Var Mvt : TFMvtImport) : boolean ;
Var Debut   : Boolean ;
    Q : TQuery ;
    N,NumLigne : Integer ;
    Entete  : TFmtEntete ;
    EnRupture : Boolean ;
    Detail  : TTabFmtDetail ;
    OkInsert,ArretDemande : Boolean ;
    NumP : Integer ;
    CptLu : TCptLu ;
BEGIN
Result:=False ;
Fillchar(IdentPiece,SizeOf(IdentPiece),#0) ; Fillchar(OldIdentPiece,SizeOf(OldIdentPiece),#0) ;
if not ChargeFormat(Fichier,InfoImp.NomFic,InfoImp.Lequel,'X',InfoImp.Format,Entete,Detail,Debut) then Exit ;
Q:=OpenSQL('SELECT * FROM IMPECR',False) ;
N:=1 ; NumLigne:=0 ; ArretDemande:=FALSE ; NumP:=0 ;
InitMove(NbLig,'') ;
While (not Eof(Fichier)) And (Not ArretDemande) do
  BEGIN
  InitMvtImport(Mvt) ; OkInsert:=TRUE ;
  Q.Insert ; InitNew(Q) ;
  LireFormat(Fichier,Entete,Detail,Debut,Q) ;
  ReinitQ(Q,N,InfoImp) ; QToMvt(Q,Mvt) ;
  Application.ProcessMessages ;
  if (InfoImp.Lequel='FBE') then // Ecritures budgétaires
    BEGIN
    if AnnuleImport then break ;
    IdentPiece.TauxDev:=1 ;
    Inc(InfoImp.NbPiece) ; NumLigne:=0 ;
    CptLu.Cpt:=Mvt.IE_JOURNAL ;
    If AlimLTabCptLu(4,QFiche[4],InfoImp.LJalLu,InfoImp.ListeCptFaux,CptLu) Then ;
    IdentPiece.TauxDev:=1 ;
    RetoucheMontantCEGID(Mvt,IdentPiece,InfoImp) ;
    AlimIdentPiece(IdentPiece,Mvt) ;
    EnRupture:=EnRuptAscii(TRUE,IdentPiece,OldIdentPiece,InfoImp) ;
    If EnRupture Then Inc(NumP) ;
    Q.FindField('IE_NUMPIECE').AsInteger:=NumP ;
    If TestBreak Then ArretDemande:=TRUE ;
    END else
    BEGIN
    If AlimCptMemoire(QFiche,Mvt,InfoImp)=1 Then OkInsert:=FALSE ; // Ligne analytique sur compte non ventilable : on ignore
    IdentPiece.TauxDev:=1 ;
    RetoucheEnrCegid('',IdentPiece,InfoImp,Mvt,TRUE,QFiche) ;
    RetoucheODANAL(Mvt) ;
    AlimIdentPiece(IdentPiece,Mvt) ;
    EnRupture:=EnRuptAscii(TRUE,IdentPiece,OldIdentPiece,InfoImp) ;
    If EnRupture Then
      BEGIN
      IdentPiece.Doublon:=FALSE ;
      If InfoImp.CtrlDB Then If ChercheDoublon(InfoImp,MvtImport^,FALSE)=1 Then IdentPiece.Doublon:=TRUE ;
      END ;
    AlimNumeros(EnRupture,IdentPiece,InfoImp,Mvt) ;
    RetoucheFinal(EnRupture,IdentPiece,Mvt) ;
    If EnRupture And TestBreak Then BEGIN OkInsert:=FALSE ; ArretDemande:=TRUE ; END ;
    END ;
  if (Q.FindField('IE_TYPEECR').AsString<>'A') Or (InfoImp.Lequel='FOD') then
     BEGIN
     InfoImp.TotDeb:=InfoImp.TotDeb+Q.FindField('IE_DEBIT').AsFloat ;
     InfoImp.TotCred:=InfoImp.TotCred+Q.FindField('IE_CREDIT').AsFloat ;
     END ;
// Même n° de ligne générale et ventil.
  if (InfoImp.Lequel<>'FBE') and (InfoImp.Lequel<>'FOD') and (Q.FindField('IE_NUMLIGNE').AsInteger<NumLigne) then Dec(NumLigne) ;
  If OkInsert Then
    BEGIN
    Inc(InfoImp.NbLigIntegre) ;
    MvtToQ(Mvt,Q,InfoImp.Lequel='FBE') ;
    Q.Post ; Inc(N) ;
    END Else Q.Cancel ;
  MoveCur(False) ;
  END ;
FiniMove ;
Ferme(Q) ;
Result:=True ;
END ;


Function ImporteEcr(Var StErr : String ; Var InfoImp : TInfoImport ; ignorel1 : boolean=False) : Boolean ;
var St,OldSt : String ;
    ExpParam,OkFic,PremFois : boolean ;
    IdentPiece,OldIdentPiece : TIdentPiece ;
    QFiche : TQFiche ;
//    DateDeb,DateFin : TDateTime ;
    i : Integer ;
    Fichier   : TextFile ;
    ArretDemande : Boolean ;
    AnaGenereSurLastLigne : Boolean ;
    NbCount : Integer ;
BEGIN
InitMove(NbLig,'') ; MoveCur(False) ; StErr:='' ; Result:=TRUE ; ArretDemande:=FALSE ; nbCount:=0 ;
Fillchar(IdentPiece,SizeOf(IdentPiece),#0) ; Fillchar(OldIdentPiece,SizeOf(OldIdentPiece),#0) ;
try
  SourisSablier ;
  New(MvtImport) ;
  ExecuteSQL('DELETE FROM IMPECR') ;
  ExpParam:=(InfoImp.Lequel='FEC') and ((InfoImp.Format<>'SAA') and (InfoImp.Format<>'EDI')
                           and (InfoImp.Format<>'SN2')  and (InfoImp.Format<>'HAL')
                           and (InfoImp.Format<>'HLI')  and (InfoImp.Format<>'MP')
                           and (InfoImp.Format<>'CGE')  and (InfoImp.Format<>'CGN')
                           and (InfoImp.Format<>'RAP')) or ((InfoImp.Lequel<>'FEC') and (InfoImp.Lequel<>'FBA')) ;
  BEGINTRANS ;
  try
      If InfoImp.Lequel='FBE' Then InitRequete(QFiche[4],4) Else
        BEGIN
        For i:=0 To 3 Do InitRequete(QFiche[i],i) ;
        If InfoImp.Sc.UseCorresp Then InitRequete(QFiche[5],5) ;
        END ;
      if ExpParam then
        BEGIN
        ImporteFormatParam(Fichier,QFiche,IdentPiece,OldIdentPiece,InfoImp,MvtImport^) ;
        END else
        BEGIN
        OkFic:=TRUE ; PremFois:=TRUE ;
        if (InfoImp.Lequel='FBA') and (InfoImp.Format='EDI') then NbLig:=Nblig*2 ;
        if ((InfoImp.Format='SN2') Or (InfoImp.Format='SAA')) And (Not Sn2Orli) Then
           BEGIN // GP à prévoir sur HALLEY et HALLEY étendu
           OkFic:=SoldeAZero(InfoImp.Nomfic,StErr) ;
           If Not OkFic Then Result:=FALSE ;
           END ;
        TravailleFichier(FALSE,StErr,InfoImp,QFiche,ignorel1) ;
        AssignFile(Fichier,InfoImp.Nomfic) ; {$I-} Reset(Fichier) ; {$I+}
        If OkFic Then
           BEGIN
           if not ignorel1 then Readln(Fichier,St) ; //Société
           if (InfoImp.Lequel='FBA') and (InfoImp.Format<>'EDI') then
             BEGIN
             END ;
           OldSt:='' ;
           While (not EOF(Fichier)) And (Not ArretDemande) do
             BEGIN
             Readln(Fichier,St) ;
             ForceCommit(NbCount,TRUE) ;
             if Trim(St)='' then Break ;
             if OkRupt or Erreur then Break ;
             if ChoixFmt.Ascii then St:=ASCII2ANSI(St) ;
             if (InfoImp.Lequel<>'FBA') then
               BEGIN
               ArretDemande:=ImportLigne(St,OldSt,QFiche,InfoImp,IdentPiece,OldIdentPiece,PremFois,FALSE,FALSE) ;
               END else
               BEGIN
               if (InfoImp.Format<>'EDI') then ArretDemande:=ImportLigne(St,OldSt,QFiche,InfoImp,IdentPiece,OldIdentPiece,PremFois,FALSE,FALSE) ;
               END ;
             MoveCur(False) ;
             END ;
           St:='' ; AnaGenereSurLastLigne:=FALSE ;
           If (Not PremFois) ANd (OldSt<>'') And AnalytiqueAbsenteSurLignePrecedente(St,OldSt,InfoImp,QFiche) Then
             BEGIN
             St:=OldSt ; OldSt:='' ;
             ImportLigne(St,OldSt,QFiche,InfoImp,IdentPiece,OldIdentPiece,PremFois,TRUE,TRUE) ;
             AnaGenereSurLastLigne:=TRUE ;
             END ;
           If (Not PremFois) And (Not AnaGenereSurLastLigne) Then
             BEGIN
             St:=OldSt ;
             ImportLigne(St,OldSt,QFiche,InfoImp,IdentPiece,OldIdentPiece,PremFois,TRUE,FALSE) ;
             END ;
           END ;
        END ;
        Finally
               if Erreur then ROLLBACK else COMMITTRANS ;
        end;
  finally
    {$I-} CloseFile(Fichier); {$I+} IOResult;
    Dispose(MvtImport) ;
    If InfoImp.Lequel='FBE' Then Ferme(QFiche[4]) Else
      BEGIN
      For i:=0 To 3 Do Ferme(QFiche[i]) ;
      If InfoImp.Sc.UseCorresp Then Ferme(QFiche[5]) ;
      END ;
    SourisNormale ;
  end ;

FiniMove ;
//If Result Then VerifExisteCptLu(InfoImp,LgComptes) ;
END ;

function ImporteLesEcritures(Msg : THMsgBox ; Var InfoImp : TInfoImport ; ignorel1 : boolean=False) : boolean ;
Var OkOk : Boolean ;
    StErr : String ;
    FmtFic : Integer ;
BEGIN
Result:=False ; StErr:='' ;
InfoImp.debutTraitement:=Now ;
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
AlimLTabDiv(0,InfoImp.LMP) ;
If InfoImp.SC.RMR<>'' Then AlimLTabDiv(1,InfoImp.LMR) ;
If InfoImp.SC.RGT<>'' Then AlimLTabDiv(2,InfoImp.LRGT) ;
InfoImp.NbLigIntegre:=0 ;
InfoImp.TotDeb:=0 ; InfoImp.TotCred:=0 ;
InfoImp.NbPiece:=0 ; InfoImp.AuMoinsUnBordereau:=FALSE ; InfoImp.AuMoinsUnePiece:=FALSE ;
AnnuleImport:=False ; Erreur:=False ; OkRupt:=False ;
ChargeDev(InfoImp) ;
OkOk:=ImporteEcr(StErr,InfoImp,ignorel1) ;
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
VideListe(InfoImp.ListeDev) ; InfoImp.ListeDev.Free ; InfoImp.ListeDev:=NIL ;
SourisNormale ;
if (Msg<>nil) and (not AnnuleImport) then
  BEGIN
  //ActivePanels(Self,True,False) ;
//  if ChoixFmt.Detruire then DeleteFile(InfoImp.NomFic) ;
  if (ChoixFmt.CompteRendu) and (InfoImp.NbLigIntegre>0) then
    BEGIN
    //Q:=OPpenSQL('SELECT SUM(IE_DEBIT),SUM(IE_CREDIT) FROM IMPER WHERE IE_TYPEECR<>"A" AND IE_TYPEECR<>"L"',True) ;
    MsgBox.Execute(33,'',' '+IntToStr(InfoImp.NbLigIntegre)+Chr(10)+MsgBox.Mess[34]+' '+IntToStr(InfoImp.NbPiece)+Chr(10)+
    Chr(10)+MsgBox.Mess[35]+AfficheMontant('#,##0.00',V_PGI.SymbolePivot,InfoImp.TotDeb,True)+Chr(10)+MsgBox.Mess[36]+AfficheMontant('#,##0.00',V_PGI.SymbolePivot,InfoImp.TotCred,True)) ;
    END ;
  END ;
Result:=not AnnuleImport ;
If (Not VH^.RecupPCL) And InfoImp.AuMoinsUnBordereau Then AlimSoucheBor(InfoImp.LSoucheBor) ;
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
    LigneEnCours.DE:=0 ;
    LigneEnCours.CE:=0 ;
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
IdentPiece.LigneEnCours.Ana:=T.FindField('IE_TYPEECR').AsString='A' ;
END ;


Function EnRupture(T : TDataSet ; Var IdentPiece,OldIdentPiece : TIdentPiece ; Var InfoImp : TInfoImport) : Boolean ;
Var OkRupt : Boolean ;
    Per1,Per2,AA,JJ : Word ;
BEGIN
Result:=False ;
If VH^.RecupPCL Or (ModeJal(IdentPiece.JalP,InfoImp)<>mPiece) Then
  BEGIN
  DecodeDate(IdentPiece.DateP,AA,Per1,JJ) ; Per1:=AA+Per1 ;
  DecodeDate(OldIdentPiece.DateP,AA,Per2,JJ) ; Per2:=AA+Per2 ;
  OkRupt:=(IdentPiece.JalP<>OldIdentPiece.JalP) Or
          (IdentPiece.NumP<>OldIdentPiece.NumP) Or
          (Per1<>Per2) Or
          (IdentPiece.QualP<>OldIdentPiece.QualP) ;
  END Else
  BEGIN
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
  END ;
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

procedure MajImpErr(ListeEntetePieceFausse : HTStringList) ;
var Nb,i : integer ;
    OldErreur : String ;
    Jal,Qual : String ;
    NumP : Integer ;
BEGIN
OldErreur:='' ;
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
var TTEcr : TTypeEcr ;
    CptLu : TCptLu ;
    Facturier : String ;
    OkForceNumPiece : Boolean ;
BEGIN
OkForceNumPiece:=InfoImp.ForceNumPiece ;
If VH^.RecupSISCOPGI Then
  BEGIN
  CptLu.Cpt:=IdentPiece.JalP ;
  If ChercheCptLu(InfoImp.LJalLu,CptLu) Then
    If CptLu.Nature='ANO' Then OkForceNumPiece:=FALSE ;
  END ;
If OkForceNumPiece Then BEGIN IdentPiece.NumPDef:=QImpEcr.Findfield('IE_NUMPIECE').AsInteger ; Exit ; END ;
//if (QImpEcr.Findfield('IE_ANA').AsString='-') and  (QImpEcr.Findfield('IE_TYPEANALYTIQUE').AsString='-') then TTEcr:=EcrGen else TTEcr:=EcrAna ;
if (QImpEcr.Findfield('IE_TYPEECR').AsString<>'A') and  (QImpEcr.Findfield('IE_TYPEANALYTIQUE').AsString='-') then TTEcr:=EcrGen else TTEcr:=EcrAna ;
if InfoImp.Prefixe='BE' then TTEcr:=EcrBud ;
if InfoImp.Prefixe='Y' then TTEcr:=EcrAna ;
CptLu.Cpt:=IdentPiece.JalP ;
If ChercheCptLu(InfoImp.LJalLu,CptLu) Then
  BEGIN
  If ((CptLu.ModeSaisie='BOR') Or (CptLu.ModeSaisie='LIB')) And ImportBor Then
    BEGIN
    IdentPiece.NumPDef:=SetIncNumSB(InfoImp.LSoucheBOR,IdentPiece.JalP,IdentPiece.DateP) ;
    END Else
    BEGIN
    Facturier:=CptLu.SoucheN ;
    If IdentPiece.QualP<>'N' Then Facturier:=CptLu.SoucheS ;
    if (Facturier='') then if FAssImp<>nil then BEGIN THMsgBox(FAssImp.FindComponent('HM')).Execute(43,'',' '+IdentPiece.JalP) ; Erreur:=True ; Exit ; END ;
    SetIncNum(TTEcr,Facturier,IdentPiece.NumPDef,IdentPiece.DAteP) ;
    END ;
  END Else if FAssImp<>nil then BEGIN THMsgBox(FAssImp.FindComponent('HM')).Execute(43,'',' '+IdentPiece.JalP) ; Erreur:=True ; Exit ; END ;
END ;

procedure AjouteMvtDansLaBase(QImpEcr,QI : TQuery ; Var IdentPiece : TIdentPiece ; Var InfoImp : TInfoImport) ;
var StTable,Pref,StName : String ;
    i  : integer ;
    CptLuJ : TCptLu ;
    LaRefLettrage : String ;
BEGIN
StTable:=InfoImp.Table ; Pref:=InfoImp.Prefixe+'_' ; LaRefLettrage:='' ;
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
// BPY le 01/10/2004 => Fiche n° 12429 : set a 1 au lieux de rien (0)
  if (StName='NUMGROUPEECR') then QI.Fields[i].AsInteger := 1 else
// Fin BPY
  if (StName='NUMLIGNE') then QI.Fields[i].AsInteger:=QImpEcr.FindField('IE_NUMLIGNE').AsInteger else
  if (StName='NUMECHE') then QI.Fields[i].AsInteger:=QImpEcr.FindField('IE_NUMECHE').AsInteger else
  if (StName='NUMVENTIL') then QI.Fields[i].AsInteger:=QImpEcr.FindField('IE_NUMVENTIL').AsInteger else
  if (StName='EXERCICE') then QI.Fields[i].AsString:=QUELEXODT(IdentPiece.DateP) else
  if (StName='DATECREATION') then QI.Fields[i].AsDateTime:=V_PGI.DateEntree else
  if (StName='DATEMODIF') then QI.Fields[i].AsDateTime:=NowH else
  if (StName='CREERPAR') then BEGIN If VH^.RecupSISCOPGI Then QI.Fields[i].AsString:='REC' Else QI.Fields[i].AsString:='IMP' ; END else
  if (StName='CONTROLETVA') then QI.Fields[i].AsString:='RIE' else
  if (StName='NUMTRAITECHQ') And VH^.RecupSISCOPGI And (QImpEcr.FindField('IE_LIBRETEXTE9').AsString<>'') Then
    BEGIN
    QI.Fields[i].AsString:=QImpEcr.FindField('IE_LIBRETEXTE9').AsString ;
    QImpEcr.FindField('IE_LIBRETEXTE9').AsString:='' ;
    END else
  if (StName='VALIDE') then
    BEGIN
    if CptLuJ.Nature='ANO' then QI.Fields[i].AsString:='X' Else QI.Fields[i].Value:=QImpEcr.FindField('IE_'+StName).AsVariant ;
    END else
    // Ecritures budgétaires
  if (StName='BUDJAL') then QI.Fields[i].AsString:=QImpEcr.FindField('IE_JOURNAL').AsString else
  if (StName='MODESAISIE') then
    BEGIN
    If VH^.RecupPCL Or (ModeJal(IdentPiece.JalP,InfoImp)<>mPiece) Then
      BEGIN
      QI.Fields[i].AsString:=CptluJ.ModeSaisie ;
      END Else QI.Fields[i].AsString:='-' ;
    END else
  if (StName='PERIODE') then QI.Fields[i].AsInteger:=GetPeriode(QImpEcr.FindField('IE_DATECOMPTABLE').AsDateTime) else
  if (StName='SEMAINE') then QI.Fields[i].AsInteger:=NumSemaine(QImpEcr.FindField('IE_DATECOMPTABLE').AsDateTime) else
  if (StName='BUDGENE') then QI.Fields[i].AsString:=QImpEcr.FindField('IE_GENERAL').AsString else
  if (StName='BUDSECT') then QI.Fields[i].AsString:=QImpEcr.FindField('IE_SECTION').AsString else
  if (StName='NATUREBUD') then QI.Fields[i].AsString:=QImpEcr.FindField('IE_NATUREPIECE').AsString else
  if (StName='REFRELEVE') then
    BEGIN
    QI.Fields[i].AsString:=QImpEcr.FindField('IE_REFRELEVE').AsString  ;
    LaRefLettrage:=QImpEcr.FindField('IE_REFRELEVE').AsString  ;
    END else
  if (StName='REFLETTRAGE') then
    BEGIN
    If (VH^.RecupLTL Or VH^.RecupSISCOPGI) And InfoImp.LettrageSISCOEnImport Then QI.Fields[i].AsString:='' Else
      QI.Fields[i].AsString:=LaRefLettrage ;
    END else
  If (RecupSISCO And ((VH^.RecupPCL) Or InfoImp.LettrageSISCOEnImport)) And (StName='ETAT') And (QImpEcr.FindField('IE_LETTRAGE').AsString<>'') Then QI.Fields[i].AsString:='---0GM000' Else
  If (StName='ETAT') Then QI.Fields[i].AsString:='0000000000' Else
  If (RecupSISCO Or RecupSISCOExt) And (StName='QUALIFORIGINE') Then QI.Fields[i].AsString:='SIS' Else
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
  if (VH^.RecupPCL) and (StName='AFFAIRE') then
    begin
    QI.Fields[i].AsString:='' ;
    END ELSE
  if (VH^.RecupPCL) and (StName = 'REFREVISION') then
    BEGIN
    // On récupère le contenu de IE_AFFAIRE dans E_REFREVISION (identecr de Sisco II)
    QI.Fields[i].AsString:= QImpEcr.FindField('IE_AFFAIRE').AsString;
    END ELSE
  if (VH^.RecupPCL) and (StName = 'IO') then
    BEGIN
    // pour la synchro avec S1
    QI.Fields[i].AsString:='X';
    END ELSE
  if (StName='COTATION') then
    BEGIN
    QI.Fields[i].AsFloat:=QImpEcr.FindField('IE_'+StName).AsFloat ;
    END else
    BEGIN
    If VH^.RecupCegid And (StName='TABLE3') And (InfoImp.BaseOrigine<>'') Then QI.Fields[i].Value:=InfoImp.BaseOrigine
     Else If VH^.RecupCegid And (StName='LIBRETEXTE1') And (InfoImp.BaseOrigine<>'') And InfoImp.CorrigeANA Then QI.Fields[i].Value:='ANA'
      Else if (QImpEcr.FindField('IE_'+StName)<>Nil) then QI.Fields[i].Value:=QImpEcr.FindField('IE_'+StName).AsVariant ;
    END ;
  END ;
QI.Post ;
AlimTotauxCptLu(InfoImp,IdentPiece) ;
END ;

Procedure AlimListeIntegre(T : TDataSet ; Var IdentPiece : TIdentPiece ; Var InfoImp : TInfoImport) ;
Var StOkCtr,StOkTvaEnc,SOkTP,StDate,CleIT : String ;
    OkTp : Boolean ;
    CptLu : TCptLu ;
    YY,MM,JJ : Word ;
BEGIN
If InfoImp.Lequel<>'FEC' Then Exit ;
CptLu.Cpt:=IdentPiece.JalP ; StOkCtr:='X' ; StOkTvaEnc:='-' ; SOkTP:='-' ; OkTp:=TRUE ;
If ChercheCptLu(InfoImp.LJalLu,CptLu) Then
  BEGIN
  If ((CptLu.Nature='VTE') Or (CptLu.Nature='ACH')) And VH^.OuiTvaEnc And
     ((IdentPiece.NatP='FC') Or (IdentPiece.NatP='AC') Or (IdentPiece.NatP='FF') Or
      (IdentPiece.NatP='AF')) Then StOkTvaEnc:='X' ;
  If IdentPiece.LigneEnCours.ODAnal Then StOkCtr:='-' ;
  If IdentPiece.LigneEnCours.ANouveau Then StOkCtr:='-' ;
  If (CptLu.Nature='ODA') Or (CptLu.Nature='REG') Or (CptLu.Nature='EXT') Or (CptLu.Nature='ECC') Or (CptLu.Nature='ANA') Then BEGIN StOkCtr:='-' ; StOkTvaEnc:='-' ; END ;
  If Not InfoImp.Sc.TraiteCtr Then StOkCtr:='-' ;
  If VH^.RecupLTL And ((CptLu.Nature='VTE') Or (CptLu.Nature='ACH') Or (CptLu.Nature='BQE')) Then StOkCtr:='X' ;
  If Not InfoImp.Sc.TraiteTVA Then StOkTvaEnc:='-' ;

  if Not VH^.OuiTP Then OkTP:=FALSE ;
  If((CptLu.Nature<>'ACH') And (CptLu.Nature<>'VTE')) Then OkTP:=FALSE ;
  if ((IdentPiece.NatP<>'FC') and (IdentPiece.NatP<>'FF') and (IdentPiece.NatP<>'AC') and (IdentPiece.NatP<>'AF')) then OkTP:=FALSE ;
  if ((IdentPiece.NatP='FC') or (IdentPiece.NatP='AC')) then if VH^.JalVTP='' then OkTP:=FALSE ;
  if ((IdentPiece.NatP='FF') or (IdentPiece.NatP='AF')) then if VH^.JalATP='' then OkTP:=FALSE ;
  If IdentPiece.QualP<>'N' Then OkTP:=FALSE ;
  If OkTP Then SOkTP:='X' ;

  (*
  StDate:='12345678' ;
  Move(IdentPiece.DateP,StDate[1],SizeOf(Double)) ;
  *)
  DecodeDate(IdentPiece.DateP,YY,MM,JJ) ;
  StDate:=FormatFloat('0000',YY)+FormatFloat('00',MM)+FormatFloat('00',JJ) ;
  CleIT:=Format_String(IdentPiece.JalP,3)+Format_String(QUELEXODT(IdentPiece.DateP),3)+
  Format_String(IdentPiece.QualP,1)+FormatFloat('00000000',IdentPiece.NumPDef)+StDate+StOkCtr+StOkTvaEnc+SOkTP ;
  If Not VH^.RecupCegid Then InfoImp.ListePieceIntegre.Add(CleIT) ;
  END ;
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
  AlimListeIntegre(QImpEcr,IdentPiece,InfoImp) ;
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
   if (Arrondi(QImpEcr.FindField('IE_DEBIT').AsFloat,V_PGI.OkDecV)=0) and (Arrondi(QImpEcr.FindField('IE_CREDIT').AsFloat,V_PGI.OkDecV)=0)  then
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

procedure IntegreGenerales(QImpEcr,QE,QY : TQuery ; Var InfoImp : TInfoImport ; Var FF : TextFile) ;
Var IdentPiece,OldIdentPiece : TIdentPiece ;
    NbCount : Integer ;
    j : Integer;
BEGIN
j := 10000;
Fillchar(IdentPiece,SizeOf(IdentPiece),#0) ; Fillchar(OldIdentPiece,SizeOf(OldIdentPiece),#0) ; NbCount:=0 ;
While not QImpEcr.EOF do
  BEGIN

  if Erreur then Break ;
  If HalteLa Then BEGIN AnnuleImport:=TRUE ; Break ; End Else
    if AnnuleImport then
      BEGIN
      if FAssImp<>nil then if (THMsgBox(FAssImp.FindComponent('HM')).Execute(39,'','')=mryes) then Break else AnnuleImport:=False ;
      END ;
  IntegrePieces(QImpEcr,QE,QY,IdentPiece,OldIdentPiece,InfoImp) ;
  Inc(InfoImp.NbLigIntegre) ;
  ActiveTotaux(True,FAssImp,QImpEcr,InfoImp) ;
  If VH^.RecupCegid Or VH^.RecupSISCOPGI Then ForceCommit(NbCount,TRUE) Else ForceCommit(NbCount,FALSE) ;
  If VH^.RecupCegid Then Writeln(FF,IntToStr(QImpEcr.FindField('IE_CHRONO').AsInteger)) ;
  {Fermeture du query pour limiter le chargement en cache au fur et à mesure que l'on se déplace dans la table}
  QImpEcr.Next ; MoveCur(False) ;
  if QImpEcr.EOF then begin

    QImpEcr.Close ;
    QImpEcr := OpenSQL('SELECT * FROM IMPECR WHERE IE_SELECTED="X" AND IE_OKCONTROLE="X" AND IE_INTEGRE="-" ' +
                       'AND IE_TYPEECR <> "L" AND IE_CHRONO >= ' + IntToStr(j) + 'AND IE_CHRONO < ' + IntToStr(j + 10000) + ' ORDER BY IE_CHRONO', False);
    j := j + 10000;
  end;
  END ;
END ;

Procedure GetContepartieImp (OkBqe : Boolean ; GeneTreso : String ; TPiece : TList ; O : TOBM ;
                             Lig : Integer ; Var LeGene,LeAux : String ; Var InfoImp : TInfoImport ;
                             IMin : Integer = -1 ; IMax : Integer = -1) ;
Var k : integer ;
    O1 : TOBM ;
    CptLu : TCptLu ;
BEGIN
If IMin=-1 Then IMin:=0 ; If IMAx=-1 Then IMax:=TPiece.Count-1 ;
LeGene:='' ; LeAux:='' ; If TPiece=Nil Then Exit ; If O=Nil Then Exit ;
if OkBqe then
   BEGIN
   {Pièce de banque}
   if O.GetMvt('E_GENERAL')=GeneTreso then
      BEGIN
      {Sur cpte de banque, contrep=première ligne non banque au dessus}
//      for k:=Lig-1 downto 0 do
      for k:=Lig-1 downto IMin do
        BEGIN
        O1:=TOBM(TPiece[k]) ;
        if (O1<>NIL) And (O1.GetMvt('E_GENERAL')<>GeneTreso) then BEGIN LeGene:=O1.GetMvt('E_GENERAL') ; LeAux:=O1.GetMvt('E_AUXILIAIRE') ; Break ; END ;
        END ;
      END else
      BEGIN
      {Sur cpte non banque, contrep=première ligne banque au dessous}
//      for k:=Lig+1 to TPiece.Count-1 do
      for k:=Lig+1 to IMax do
        BEGIN
        O1:=TOBM(TPiece[k]) ;
        if (O1<>NIL) And (O1.GetMvt('E_GENERAL')=GeneTreso) then BEGIN LeGene:=O1.GetMvt('E_GENERAL') ; LeAux:=O1.GetMvt('E_AUXILIAIRE') ; Break ; END ;
        END ;
      END ;
   END else
   BEGIN
   {Cas normal}
   if O.GetMvt('E_AUXILIAIRE')<>'' then
      BEGIN
      {Lecture avant pour trouver "HT"}
//      for k:=Lig+1 to TPiece.Count-1 do
      for k:=Lig+1 to IMax do
          BEGIN
          O1:=TOBM(TPiece[k]) ;
          If O1<>NIL Then if O1.GetMvt('E_TYPEMVT')='HT' then BEGIN LeGene:=O1.GetMvt('E_GENERAL') ; Break ; END ;
          END ;
      if  LeGene='' then
          BEGIN
          {Lecture arrière pour trouver "HT"}
//          for k:=Lig-1 downto 0 do
          for k:=Lig-1 downto IMin do
              BEGIN
              O1:=TOBM(TPiece[k]) ;
              If O1<>NIL Then if O1.GetMvt('E_TYPEMVT')='HT' then BEGIN LeGene:=O1.GetMvt('E_GENERAL') ; Break ; END ;
              END ;
          END ;
      END else
      BEGIN
      {Lecture arrière pour trouver "Tiers"}
//      for k:=Lig-1 downto 0 do
      for k:=Lig-1 downto IMIn do
          BEGIN
          O1:=TOBM(TPiece[k]) ;
          If (O1<>NIL) And (O1.GetMvt('E_AUXILIAIRE')<>'') Then BEGIN LeGene:=O1.GetMvt('E_GENERAL') ; LeAux:=O1.GetMvt('E_AUXILIAIRE') ; Break ; END ;
          END ;
      if  LeGene='' then
          BEGIN
          {Lecture avant pour trouver "Tiers"}
//          for k:=Lig+1 to TPiece.Count-1 do
          for k:=Lig+1 to IMax do
              BEGIN
              O1:=TOBM(TPiece[k]) ;
              If (O1<>NIL) And (O1.GetMvt('E_AUXILIAIRE')<>'') Then BEGIN LeGene:=O1.GetMvt('E_GENERAL') ; LeAux:=O1.GetMvt('E_AUXILIAIRE') ; Break ; END ;
              END ;
          END ;
      END ;
   END ;
{Cas particuliers}
if LeGene<>'' then Exit ;

if ((OkBqe) and (O.GetMvt('E_GENERAL')<>GeneTreso)) then LeGene:=GeneTreso else
  BEGIN
  Fillchar(CptLu,SizeOf(CptLu),#0) ; CptLu.Cpt:=O.GetMvt('E_GENERAL') ; ChercheCptLu(InfoImp.LGenLu,CptLu) ;
  if Not (CptLu.Nature='BQE') then
     BEGIN
//     for k:=Lig+1 to TPiece.Count-1 do
     for k:=Lig+1 to IMax do
       BEGIN
       O1:=TOBM(TPiece[k]) ;
       If (O1<>NIL) Then
         BEGIN
         Fillchar(CptLu,SizeOf(CptLu),#0) ; CptLu.Cpt:=O.GetMvt('E_GENERAL') ; ChercheCptLu(InfoImp.LGenLu,CptLu) ;
         if (CptLu.Nature='BQE') then BEGIN LeGene:=O1.GetMvt('E_GENERAL') ; Break ; END ;
         END ;
       END ;
//     if LeGene='' then for k:=Lig-1 downto 0 do
     if LeGene='' then for k:=Lig-1 downto IMin do
       BEGIN
       O1:=TOBM(TPiece[k]) ;
       If (O1<>NIL) Then
         BEGIN
         Fillchar(CptLu,SizeOf(CptLu),#0) ; CptLu.Cpt:=O.GetMvt('E_GENERAL') ; ChercheCptLu(InfoImp.LGenLu,CptLu) ;
         if (CptLu.Nature='BQE') then BEGIN LeGene:=O1.GetMvt('E_GENERAL') ; Break ; END ;
         END ;
       END ;
     END else
     BEGIN
//     for k:=Lig-1 downto 0 do
     for k:=Lig-1 downto IMin do
       BEGIN
       O1:=TOBM(TPiece[k]) ;
       If (O1<>NIL) Then
         if O1.GetMvt('E_AUXILIAIRE')<>'' then BEGIN LeGene:=O1.GetMvt('E_GENERAL') ; LeAux:=O1.GetMVt('E_AUXILIAIRE') ; Break ; END ;
       END ;
//     if LeGene='' then for k:=Lig+1 to TPiece.Count-1 do
     if LeGene='' then for k:=Lig+1 to IMax do
       BEGIN
       O1:=TOBM(TPiece[k]) ;
       If (O1<>NIL) Then
         if O1.GetMvt('E_AUXILIAIRE')<>'' then BEGIN LeGene:=O1.GetMvt('E_GENERAL') ; LeAux:=O1.GetMVt('E_AUXILIAIRE') ; Break ; END ;
       END ;
     END ;
  END ;
{Si rien trouvé, swaper les lignes 1 et 2}
if LeGene<>'' then Exit ;
//if Lig>1 then
if Lig>IMin+1 then
  BEGIN
//  O1:=NIL ; If TPiece.Count>=1 Then  O1:=TOBM(TPiece[0]) ;
  O1:=NIL ; If IMax-IMin>=0 Then  O1:=TOBM(TPiece[IMin]) ;
  If O1<>NIL Then BEGIN LeGene:=O1.GetMvt('E_GENERAL') ; LeAux:=O1.GetMvt('E_AUXILIAIRE') ; END ;
//  END else if TPiece.Count>=2 then
  END else if IMax-IMin>=1 then
  BEGIN
//  O1:=TOBM(TPiece[1]) ;
  O1:=TOBM(TPiece[IMin+1]) ;
  If O1<>NIL Then BEGIN LeGene:=O1.GetMvt('E_GENERAL') ; LeAux:=O1.GetMvt('E_AUXILIAIRE') ; END ;
  END ;
END ;

Procedure MajCtrAna(O : TOBM ; CtrGene,CtrAux : String) ;
var St : String ;
BEGIN
If O.GetMvt('E_ANA')<>'X' Then Exit ;
St:='UPDATE ANALYTIQ SET Y_CONTREPARTIEGEN="'+CtrGene+'",'+'Y_CONTREPARTIEAUX="'+CtrAux+'" WHERE '
   +' Y_JOURNAL="'+O.GetMvt('E_JOURNAL')+'" AND Y_DATECOMPTABLE="'
   +USDateTime(O.GetMvt('E_DATECOMPTABLE'))+'" AND Y_NATUREPIECE="'
   +O.GetMvt('E_NATUREPIECE')+'" AND Y_NUMEROPIECE='
   +IntToStr(O.GetMvt('E_NUMEROPIECE'))+' AND Y_NUMLIGNE='+IntToStr(O.GetMvt('E_NUMLIGNE'))+' AND '
   +' Y_QUALIFPIECE="'+O.GetMvt('E_QUALIFPIECE')+'" AND Y_EXERCICE="'+O.GetMvt('E_EXERCICE')+'" ' ;
ExecuteSQL(St) ;
END ;

Procedure TraitementSurEcr ( OkCtr,OkTva : Boolean ; M : RMVT ; ForceFlag : boolean ; Var InfoImp : TInfoImport) ;
Var TPiece : TList ;
    Q : TQuery ;
    O : TOBM ;
    NbTiers,Lig,LigTiers,NumBase,i : integer ;
    SorteTva  : TSorteTva ;
    ExigeTva  : TExigeTva ;
    Solde,TTC,Coef : double ;
    Okok,OkTiers,ExisteEnc : boolean ;
    NaturePiece,NatureJal,RegimeTva,RegimeEnc,StE,NatGene,CodeTva : String3 ;
    Auxi,Gene,StUp,SQL,Coll : String ;
    TabTvaEnc : Array[1..5] of Double ;
    CptLu,CptLuJ : TCptLu ;
    CtrGene,CtrAuxi,Cpt : String ;
    OkTvaSurDebit,OkTvaSurEnc,OkMajCtr : Boolean ;
    OnEstSurUnBordereau : Boolean ;
    E_NUMLIGNEENCOURS : Integer ;
    TD,TC : Double ;
    BordereauFini : Boolean ;
    ll : Integer ;
  Label 0 ;
BEGIN
{#TVAENC}
//if Not VH^.OuiTvaEnc then Exit ;
OkTvaSurDebit:=FALSE ; OkTvaSurEnc:=FALSE ;
If OkTva Then BEGIN OkTvaSurDebit:=TRUE ; OkTvaSurEnc:=TRUE ; END ;
if V_PGI.IoError<>oeOk then Exit ;
if M.Jal='' then Exit ;
Okok:=True ; ExisteEnc:=False ;
{Identification nature facture}
NaturePiece:=M.Nature ;
if ((NATUREPIECE<>'FC') and (NATUREPIECE<>'AC') and (NATUREPIECE<>'FF') and (NATUREPIECE<>'AF')) then
  BEGIN
  OkTva:=FALSE ; OkTvaSurDebit:=FALSE ; OkTvaSurEnc:=FALSE ;
  END ;
If (Not OkTva) And (Not OkCtr) Then Exit ;
{Identification journal facture}
NatureJal:='' ; TTC:=0 ;
CptLuJ.Cpt:=M.Jal ;
If ChercheCptLu(InfoImp.LJalLu,CptLuJ) Then NatureJal:=CptLuJ.Nature ;
SorteTva:=stvDivers ;
if NatureJal='VTE' then SorteTva:=stvVente else if NatureJal='ACH' then SorteTva:=stvAchat ;
if SorteTva=stvDivers then
  BEGIN
  OkTva:=FALSE ; OkTvaSurDebit:=FALSE ; OkTvaSurEnc:=FALSE ;
  END ;
{Constitution d'une liste des lignes de l'écriture}
E_NUMLIGNEENCOURS:=0 ;
0:
TPiece:=TList.Create ;
OnEstSurUnBordereau:=(M.ModeSaisieJal='BOR') Or (M.ModeSaisieJal='LIB') ;
SQL:='Select * from ECRITURE Where '+WhereEcriture(tsGene,M,False,OnEstSurUnBordereau) ;
If E_NUMLIGNEENCOURS>0 Then SQL:=SQL+' AND E_NUMLIGNE>'+IntToStr(E_NUMLIGNEENCOURS)+' ' ;
If OnEstSurUnBordereau Then SQL:=SQL+' ORDER BY E_JOURNAL, E_EXERCICE, E_PERIODE, E_NUMEROPIECE, E_NUMLIGNE, E_NUMECHE, E_QUALIFPIECE'
                       Else SQL:=SQL+' ORDER BY E_JOURNAL, E_EXERCICE, E_DATECOMPTABLE, E_NUMEROPIECE, E_NUMLIGNE, E_NUMECHE, E_QUALIFPIECE' ;
Q:=OpenSQL(SQL,True) ;
TD:=0 ; TC:=0 ;
If Q.Eof Then BEGIN Ferme(Q) ; VideListe(TPiece) ; TPiece.Free ; Exit ;  END ;
BordereauFini:=FALSE ;
While Not Q.EOF do
   BEGIN
   O:=TOBM.Create(EcrGen,'',False) ;
   O.ChargeMvt(Q) ; TPiece.Add(O) ;
   If OnEstSurUnBordereau Then
     BEGIN
     TD:=Arrondi(TD+Q.FindField('E_DEBIT').AsFloat,V_PGI.okDecV) ;
     TC:=Arrondi(TC+Q.FindField('E_CREDIT').AsFloat,V_PGI.okDecV) ;
     If (Arrondi(TD-TC,V_PGI.OkDecV)=0)  Then
       BEGIN
       E_NUMLIGNEENCOURS:=Q.FindField('E_NUMLIGNE').AsInteger ;
       BordereauFini:=TRUE ;
       Break ;
       END ;
     END ;
   Q.Next ;
   END ;
If BordereauFini Then BEGIN Q.Next ; If Not Q.Eof Then BordereauFini:=FALSE ; END ;
Ferme(Q) ;
{TVA : Détermination du nombre de tiers}
{CTR : Détermination des contreparties}
NbTiers:=0 ; LigTiers:=-1 ;
for Lig:=0 to TPiece.Count-1 do
    BEGIN
    O:=TOBM(TPiece[Lig]) ; OkTiers:=False ;
    if ((O.GetMvt('E_ECHE')='X') and (O.GetMvt('E_NUMECHE')>=1)) then
       BEGIN
       if O.GetMvt('E_NUMECHE')=1 then OkTiers:=True ;
       TTC:=Arrondi(TTC+O.GetMvt('E_DEBIT')-O.GetMvt('E_CREDIT'),V_PGI.OkDecV) ;
       END ;
    if OkTiers then
       BEGIN
       Inc(NbTiers) ;
       if NbTiers=1 then LigTiers:=Lig else BEGIN Okok:=False ; Break ; END ;
       END ;
    END ;
if ((Not Okok) or (NbTiers<>1) or (LigTiers<0) or (TTC=0)) then
  BEGIN
  OkTva:=FALSE ; OkTvaSurDebit:=FALSE ; OkTvaSurEnc:=FALSE ;
  END ;
If (Not OkTva) And (Not OkCtr) Then BEGIN VideListe(TPiece) ; TPiece.Free ; Exit ; END ;
{Lecture paramètres tiers}
RegimeTva:='' ; RegimeEnc:='' ;
If LigTiers>=0 Then
  BEGIN
  O:=TOBM(TPiece[LigTiers]) ;
  Coll:=O.GetMvt('E_GENERAL') ; if Not EstCollFact(Coll) then BEGIN OkTva:=FALSE ; OkTvaSurEnc:=FALSE ; END ;
  If (Not OkTvaSurDebit) And (Not OkTvaSurEnc) And (Not OkCtr) Then BEGIN VideListe(TPiece) ; TPiece.Free ; Exit ; END ;
  if O.GetMvt('E_AUXILIAIRE')<>'' then
     BEGIN
     Auxi:=O.GetMvt('E_AUXILIAIRE') ;
     CptLu.Cpt:=Auxi ;
     If ChercheCptLu(InfoImp.LAuxLu,CptLu) Then BEGIN RegimeTva:=CptLu.RegimeTva ; RegimeEnc:=CptLu.TvaEnc ; END ;
     END else
     BEGIN
     Gene:=O.GetMvt('E_GENERAL') ;
     (*
     Q:=OpenSQL('Select G_REGIMETVA, G_TVAENCAISSEMENT from GENERAUX Where G_GENERAL="'+Gene+'"',True) ;
     if Not Q.EOF then BEGIN RegimeTva:=Q.Fields[0].AsString ; RegimeEnc:=Q.Fields[1].AsString ; END ;
     Ferme(Q) ;
     *)
     CptLu.Cpt:=Gene ;
     If ChercheCptLu(InfoImp.LGenLu,CptLu) Then BEGIN RegimeTva:=CptLu.RegimeTva ; RegimeEnc:=CptLu.TvaEnc ; END ;
     END ;
  END ;
//if ((RegimeTva='') or ((SorteTva=stvAchat) and (RegimeEnc=''))) then OkTva:=FALSE ;
if (RegimeTva='') Then
  BEGIN
  OkTva:=FALSE ; OkTvaSurDebit:=FALSE ; OkTvaSurEnc:=FALSE ;
  END ;
If ((SorteTva=stvAchat) and (RegimeEnc='')) then
  BEGIN
  OkTva:=FALSE ; OkTvaSurEnc:=FALSE ;
  END ;
If (Not OkTvaSurDebit) And (Not OkTvaSurEnc) And (Not OkCtr) Then BEGIN VideListe(TPiece) ; TPiece.Free ; Exit ; END ;
{Affectation exigibilité Tva}
if SorteTva=stvAchat then ExigeTva:=Code2Exige(RegimeEnc) else ExigeTva:=tvaMixte ;
{Si Fournisseur débit --> débit donc rien à traiter}
if ExigeTva=tvaDebit then OkTvaSurEnc:=FALSE ;
If (Not OkTvaSurEnc) And (Not OkTvaSurDebit) And (Not OkCtr) Then BEGIN VideListe(TPiece) ; TPiece.Free ; Exit ; END ;
{Affectation des régimes, des tva,tpf et enc O/N sur les lignes}
for Lig:=0 to TPiece.Count-1 do
    BEGIN
    O:=TOBM(TPiece[Lig]) ; Gene:=O.GetMvt('E_GENERAL') ;
    CptLu.Cpt:=Gene ;
    If ChercheCptLu(InfoImp.LGenLu,CptLu) Then
      BEGIN
      NatGene:=CptLu.Nature ;
      if ((NatGene='CHA') or (NatGene='PRO') or (NatGene='IMO')) then
         BEGIN
         If OkTva And ((O.GetMvt('E_TVA')='') Or ForceFlag) Then O.PutMvt('E_TVA',CptLu.Tva) ;
         If OkTva And ((O.GetMvt('E_TPF')='') Or ForceFlag) Then O.PutMvt('E_TPF',CptLu.Tpf) ;
         If (O.GetMvt('E_TVAENCAISSEMENT')='') Or ForceFlag Then
           BEGIN
           StE:=FlagEncais(ExigeTva,(CptLu.TvaEncHT='X')) ;
           END else StE:=O.GetMvt('E_TVAENCAISSEMENT') ;
         if StE='X' then ExisteEnc:=True ;
         If O.GetMvt('E_REGIMETVA')='' Then O.PutMvt('E_REGIMETVA',RegimeTva) ;
         If OkTvaSurEnc Then
           BEGIN
           If OkTva And (O.GetMvt('E_TVAENCAISSEMENT')<>StE) Then O.PutMvt('E_TVAENCAISSEMENT',StE) ;
           END Else O.PutMvt('E_TVAENCAISSEMENT','-') ;
         END else
         BEGIN
         If OkTva Then O.PutMvt('E_TVA','') ;
         END ;
      END ;
    END ;
{Si aucune ligne en encaissement --> aucun traitement}
if Not ExisteEnc then BEGIN OkTva:=FALSE ; OkTvaSurEnc:=FALSE ; END ;
If (Not OkTvaSurDebit) And (Not OkTvaSurEnc) And (Not OkCtr) Then BEGIN VideListe(TPiece) ; TPiece.Free ; Exit ; END ;
{Calcul du tableau des bases Tva}
FillChar(TabTvaEnc,Sizeof(TabTvaEnc),#0) ;
for Lig:=0 to TPiece.Count-1 do
    BEGIN
    O:=TOBM(TPiece[Lig]) ;
    CodeTva:=O.GetMvt('E_TVA') ; if CodeTva='' then Continue ;
    if SorteTva=stvVente then Solde:=O.GetMvt('E_CREDIT')-O.GetMvt('E_DEBIT')
                         else Solde:=O.GetMvt('E_DEBIT')-O.GetMvt('E_CREDIT') ;
    If OkTvaSurEnc Then
      BEGIN
      if O.GetMvt('E_TVAENCAISSEMENT')='X' then
         BEGIN
         NumBase:=Tva2NumBase(CodeTva) ;
         if NumBase>0 then TabTvaEnc[NumBase]:=TabTvaEnc[NumBase]+Solde ;
         END else
         BEGIN
         TabTvaEnc[5]:=TabTvaEnc[5]+Solde ;
         END ;
      END ;
    END ;
{Report du tableau des bases sur les échéances}
for Lig:=0 to TPiece.Count-1 do
    BEGIN
    O:=TOBM(TPiece[Lig]) ;
    if ((O.GetMvt('E_ECHE')='X') and (O.GetMvt('E_NUMECHE')>0)) then
       BEGIN
       Solde:=O.GetMvt('E_DEBIT')-O.GetMvt('E_CREDIT') ;
       If (TTC<>0) And OkTvaSurEnc Then
         BEGIN
         Coef:=Solde/TTC ;
         for i:=1 to 4 do If OkTvaSurEnc Then O.PutMvt('E_ECHEENC'+IntToStr(i),Arrondi(TabTvaEnc[i]*Coef,V_PGI.OkDecV)) ;
         If OkTvaSurEnc Then O.PutMvt('E_ECHEDEBIT',Arrondi(TabTvaEnc[5]*Coef,V_PGI.OkDecV)) ;
         If OkTvaSurEnc Then O.PutMvt('E_EMETTEURTVA','X') ;
         END ;
       END ;
    If OkCtr And (O.GetMvt('E_CONTREPARTIEGEN')='') And (O.GetMvt('E_CONTREPARTIEAUX')='') Then
      BEGIN
      GetContepartieImp(NatureJal='BQE',CptLuJ.CollTiers,TPiece,O,Lig,CtrGene,CtrAuxi,InfoImp) ;
      OkMajCtr:=FALSE ;
      If CtrGene<>'' Then BEGIN OkMajCtr:=TRUE ; O.PutMvt('E_CONTREPARTIEGEN',CtrGene) ; END ;
      If CtrAuxi<>'' Then BEGIN OkMajCtr:=TRUE ; O.PutMvt('E_CONTREPARTIEAUX',CtrAuxi) ; END ;
      If OkMajCtr And OkMajCtr Then MajCtrAna(O,CtrGene,CtrAuxi) ;
      END ;
    END ;
{Mise à jour fichier}
for Lig:=0 to TPiece.Count-1 do
    BEGIN
    O:=TOBM(TPiece[Lig]) ;
    StUp:=O.StPourUpdate ; if StUp='' then Continue ;
    SQL:='UPDATE ECRITURE SET '+StUp+' Where  '+WhereEcriture(tsGene,M,False,OnEstSurUnBordereau)
        +' AND E_NUMLIGNE='+IntToStr(O.GetMvt('E_NUMLIGNE'))+' AND E_NUMECHE='+IntToStr(O.GetMvt('E_NUMECHE')) ;
    ll:=ExecuteSQL(SQL) ;
    if ll<>1 then
      BEGIN
      V_PGI.IoError:=oeUnknown ; Break ;
      END ;
    END ;
{Dispose mémoire}
VideListe(TPiece) ; TPiece.Free ;
If OnEstSurUnBordereau And (E_NUMLIGNEENCOURS>0) Then
  BEGIN
  If Not BordereauFini Then Goto 0 ;
  END ;
END ;

Procedure TrouveMinMax(Var IMIN,IMax : Integer ; TPiece : TList) ;
Var Lig : Integer ;
    TD,TC : Double ;
    Cpt : String ;
    O : TOBM ;
    Trouv : Boolean ;
BEGIN
IMin:=IMax+1 ; IMax:=IMin ; Trouv:=FALSE ; TD:=0 ; TC:=0 ;
for Lig:=IMin to TPiece.Count-1 do
  BEGIN
  O:=TOBM(TPiece[Lig]) ;
  TD:=Arrondi(TD+O.GetMvt('E_DEBIT'),V_PGI.okDecV) ;
  TC:=Arrondi(TC+O.GetMvt('E_CREDIT'),V_PGI.okDecV) ;
  Cpt:=O.GetMvt('E_GENERAL') ;
  If (Arrondi(TD-TC,V_PGI.OkDecV)=0) Then
    BEGIN
    IMax:=Lig ; Trouv:=TRUE ; Break ;
    END ;
  END ;
If Not Trouv Then IMax:=TPiece.Count-1 ;
If IMax>TPiece.Count-1 Then IMax:=TPiece.Count-1 ;
If Imax=Imin Then IMax:=TPiece.Count-1 ;
END ;

Procedure TraitementSurEcr2 ( OkCtr,OkTva : Boolean ; M : RMVT ; ForceFlag : boolean ; Var InfoImp : TInfoImport) ;
Var TPiece : TList ;
    Q : TQuery ;
    O : TOBM ;
    NbTiers,Lig,LigTiers,NumBase,i : integer ;
    SorteTva  : TSorteTva ;
    ExigeTva  : TExigeTva ;
    Solde,TTC,Coef : double ;
    Okok,OkTiers,ExisteEnc : boolean ;
    NaturePiece,NatureJal,RegimeTva,RegimeEnc,StE,NatGene,CodeTva : String3 ;
    Auxi,Gene,StUp,SQL,Coll : String ;
    TabTvaEnc : Array[1..5] of Double ;
    CptLu,CptLuJ : TCptLu ;
    CtrGene,CtrAuxi : String ;
    OkTvaSurDebit,OkTvaSurEnc,OkMajCtr : Boolean ;
    OnEstSurUnBordereau : Boolean ;
    ll,IMin,IMax : Integer ;
  Label 0,1 ;
BEGIN
{#TVAENC}
//if Not VH^.OuiTvaEnc then Exit ;
OkTvaSurDebit:=FALSE ; OkTvaSurEnc:=FALSE ;
If OkTva Then BEGIN OkTvaSurDebit:=TRUE ; OkTvaSurEnc:=TRUE ; END ;
if V_PGI.IoError<>oeOk then Exit ;
if M.Jal='' then Exit ;
Okok:=True ; ExisteEnc:=False ;
{Identification nature facture}
NaturePiece:=M.Nature ;
if ((NATUREPIECE<>'FC') and (NATUREPIECE<>'AC') and (NATUREPIECE<>'FF') and (NATUREPIECE<>'AF')) then
  BEGIN
  OkTva:=FALSE ; OkTvaSurDebit:=FALSE ; OkTvaSurEnc:=FALSE ;
  END ;
If (Not OkTva) And (Not OkCtr) Then Exit ;
{Identification journal facture}
NatureJal:='' ; TTC:=0 ;
CptLuJ.Cpt:=M.Jal ;
If ChercheCptLu(InfoImp.LJalLu,CptLuJ) Then NatureJal:=CptLuJ.Nature ;
SorteTva:=stvDivers ;
if NatureJal='VTE' then SorteTva:=stvVente else if NatureJal='ACH' then SorteTva:=stvAchat ;
If NatureJal='ANO' Then Exit ;
if SorteTva=stvDivers then
  BEGIN
  OkTva:=FALSE ; OkTvaSurDebit:=FALSE ; OkTvaSurEnc:=FALSE ;
  END ;
{Constitution d'une liste des lignes de l'écriture}
TPiece:=TList.Create ;
OnEstSurUnBordereau:=(M.ModeSaisieJal='BOR') Or (M.ModeSaisieJal='LIB') ;
SQL:='Select * from ECRITURE Where '+WhereEcriture(tsGene,M,False,OnEstSurUnBordereau) ;
If OnEstSurUnBordereau Then SQL:=SQL+' ORDER BY E_JOURNAL, E_EXERCICE, E_PERIODE, E_NUMEROPIECE, E_NUMLIGNE, E_NUMECHE, E_QUALIFPIECE'
                       Else SQL:=SQL+' ORDER BY E_JOURNAL, E_EXERCICE, E_DATECOMPTABLE, E_NUMEROPIECE, E_NUMLIGNE, E_NUMECHE, E_QUALIFPIECE' ;
Q:=OpenSQL(SQL,True) ;
While Not Q.EOF do
   BEGIN
   O:=TOBM.Create(EcrGen,'',False) ;
   O.ChargeMvt(Q) ; TPiece.Add(O) ;
   Q.Next ;
   END ;
Ferme(Q) ;
{TVA : Détermination du nombre de tiers}
{CTR : Détermination des contreparties}
IMin:=-1 ; IMax:=-1 ;
0:
NbTiers:=0 ; LigTiers:=-1 ;
If OnEstSurUnBordereau Then
  BEGIN
  TrouveMinMax(IMin,IMax,TPiece) ;
  If IMin<>0 Then
    BEGIN
    OkTvaSurDebit:=FALSE ; OkTvaSurEnc:=FALSE ;
    If OkTva Then BEGIN OkTvaSurDebit:=TRUE ; OkTvaSurEnc:=TRUE ; END ;
    Okok:=True ; ExisteEnc:=False ;
    {Identification nature facture}
    O:=TOBM(TPiece[IMin]) ; If O<>NIL Then NATUREPIECE:=O.GetMvt('E_NATUREPIECE') ;
//    NaturePiece:=M.Nature ;
    if ((NATUREPIECE<>'FC') and (NATUREPIECE<>'AC') and (NATUREPIECE<>'FF') and (NATUREPIECE<>'AF')) then
      BEGIN
      OkTva:=FALSE ; OkTvaSurDebit:=FALSE ; OkTvaSurEnc:=FALSE ;
      END ;
    If (Not OkTva) And (Not OkCtr) Then Goto 1 ;
    END ;
  END Else
  BEGIN
  IMin:=0 ; IMAx:=TPiece.Count-1 ;
  END ;
for Lig:=IMin to IMax do
    BEGIN
    O:=TOBM(TPiece[Lig]) ; OkTiers:=False ;
    if ((O.GetMvt('E_ECHE')='X') and (O.GetMvt('E_NUMECHE')>=1)) then
       BEGIN
       if O.GetMvt('E_NUMECHE')=1 then OkTiers:=True ;
       TTC:=Arrondi(TTC+O.GetMvt('E_DEBIT')-O.GetMvt('E_CREDIT'),V_PGI.OkDecV) ;
       END ;
    if OkTiers then
       BEGIN
       Inc(NbTiers) ;
       if NbTiers=1 then LigTiers:=Lig else BEGIN Okok:=False ; Break ; END ;
       END ;
    END ;
if ((Not Okok) or (NbTiers<>1) or (LigTiers<0) or (TTC=0)) then
  BEGIN
  OkTva:=FALSE ; OkTvaSurDebit:=FALSE ; OkTvaSurEnc:=FALSE ;
  END ;
If (Not OkTva) And (Not OkCtr) Then
  BEGIN
  If OnEstSurUnBordereau Then Goto 1 Else
    BEGIN
    VideListe(TPiece) ; TPiece.Free ; Exit ;
    END ;
  END ;
{Lecture paramètres tiers}
RegimeTva:='' ; RegimeEnc:='' ;
If LigTiers>=0 Then
  BEGIN
  O:=TOBM(TPiece[LigTiers]) ;
  Coll:=O.GetMvt('E_GENERAL') ; if Not EstCollFact(Coll) then BEGIN OkTva:=FALSE ; OkTvaSurEnc:=FALSE ; END ;
  If (Not OkTvaSurDebit) And (Not OkTvaSurEnc) And (Not OkCtr) Then
    BEGIN
    If OnEstSurUnBordereau Then Goto 1 Else
      BEGIN
      VideListe(TPiece) ; TPiece.Free ; Exit ;
      END ;
    END ;
  if O.GetMvt('E_AUXILIAIRE')<>'' then
     BEGIN
     Auxi:=O.GetMvt('E_AUXILIAIRE') ;
     CptLu.Cpt:=Auxi ;
     If ChercheCptLu(InfoImp.LAuxLu,CptLu) Then BEGIN RegimeTva:=CptLu.RegimeTva ; RegimeEnc:=CptLu.TvaEnc ; END ;
     END else
     BEGIN
     Gene:=O.GetMvt('E_GENERAL') ;
     (*
     Q:=OpenSQL('Select G_REGIMETVA, G_TVAENCAISSEMENT from GENERAUX Where G_GENERAL="'+Gene+'"',True) ;
     if Not Q.EOF then BEGIN RegimeTva:=Q.Fields[0].AsString ; RegimeEnc:=Q.Fields[1].AsString ; END ;
     Ferme(Q) ;
     *)
     CptLu.Cpt:=Gene ;
     If ChercheCptLu(InfoImp.LGenLu,CptLu) Then BEGIN RegimeTva:=CptLu.RegimeTva ; RegimeEnc:=CptLu.TvaEnc ; END ;
     END ;
  END ;
//if ((RegimeTva='') or ((SorteTva=stvAchat) and (RegimeEnc=''))) then OkTva:=FALSE ;
if (RegimeTva='') Then
  BEGIN
  OkTva:=FALSE ; OkTvaSurDebit:=FALSE ; OkTvaSurEnc:=FALSE ;
  END ;
If ((SorteTva=stvAchat) and (RegimeEnc='')) then
  BEGIN
  OkTva:=FALSE ; OkTvaSurEnc:=FALSE ;
  END ;
If (Not OkTvaSurDebit) And (Not OkTvaSurEnc) And (Not OkCtr) Then
  BEGIN
  If OnEstSurUnBordereau Then Goto 1 Else
    BEGIN
    VideListe(TPiece) ; TPiece.Free ; Exit ;
    END ;
  END ;
{Affectation exigibilité Tva}
if SorteTva=stvAchat then ExigeTva:=Code2Exige(RegimeEnc) else ExigeTva:=tvaMixte ;
{Si Fournisseur débit --> débit donc rien à traiter}
if ExigeTva=tvaDebit then OkTvaSurEnc:=FALSE ;
If (Not OkTvaSurEnc) And (Not OkTvaSurDebit) And (Not OkCtr) Then
  BEGIN
  If OnEstSurUnBordereau Then Goto 1 Else
    BEGIN
    VideListe(TPiece) ; TPiece.Free ; Exit ;
    END ;
  END ;
{Affectation des régimes, des tva,tpf et enc O/N sur les lignes}
for Lig:=IMin to IMax do
    BEGIN
    O:=TOBM(TPiece[Lig]) ; Gene:=O.GetMvt('E_GENERAL') ;
    CptLu.Cpt:=Gene ;
    If ChercheCptLu(InfoImp.LGenLu,CptLu) Then
      BEGIN
      NatGene:=CptLu.Nature ;
      if ((NatGene='CHA') or (NatGene='PRO') or (NatGene='IMO')) then
         BEGIN
         If OkTva And ((O.GetMvt('E_TVA')='') Or ForceFlag) Then O.PutMvt('E_TVA',CptLu.Tva) ;
         If OkTva And ((O.GetMvt('E_TPF')='') Or ForceFlag) Then O.PutMvt('E_TPF',CptLu.Tpf) ;
         If (O.GetMvt('E_TVAENCAISSEMENT')='') Or ForceFlag Then
           BEGIN
           StE:=FlagEncais(ExigeTva,(CptLu.TvaEncHT='X')) ;
           END else StE:=O.GetMvt('E_TVAENCAISSEMENT') ;
         if StE='X' then ExisteEnc:=True ;
         If O.GetMvt('E_REGIMETVA')='' Then O.PutMvt('E_REGIMETVA',RegimeTva) ;
         If OkTvaSurEnc Then
           BEGIN
           If OkTva And (O.GetMvt('E_TVAENCAISSEMENT')<>StE) Then O.PutMvt('E_TVAENCAISSEMENT',StE) ;
           END Else O.PutMvt('E_TVAENCAISSEMENT','-') ;
         END else
         BEGIN
         If OkTva Then O.PutMvt('E_TVA','') ;
         END ;
      END ;
    END ;
{Si aucune ligne en encaissement --> aucun traitement}
if Not ExisteEnc then BEGIN OkTva:=FALSE ; OkTvaSurEnc:=FALSE ; END ;
If (Not OkTvaSurDebit) And (Not OkTvaSurEnc) And (Not OkCtr) Then
  BEGIN
  If OnEstSurUnBordereau Then Goto 1 Else
    BEGIN
    VideListe(TPiece) ; TPiece.Free ; Exit ;
    END ;
  END ;
{Calcul du tableau des bases Tva}
FillChar(TabTvaEnc,Sizeof(TabTvaEnc),#0) ;
for Lig:=IMin to IMax do
    BEGIN
    O:=TOBM(TPiece[Lig]) ;
    CodeTva:=O.GetMvt('E_TVA') ; if CodeTva='' then Continue ;
    if SorteTva=stvVente then Solde:=O.GetMvt('E_CREDIT')-O.GetMvt('E_DEBIT')
                         else Solde:=O.GetMvt('E_DEBIT')-O.GetMvt('E_CREDIT') ;
    If OkTvaSurEnc Then
      BEGIN
      if O.GetMvt('E_TVAENCAISSEMENT')='X' then
         BEGIN
         NumBase:=Tva2NumBase(CodeTva) ;
         if NumBase>0 then TabTvaEnc[NumBase]:=TabTvaEnc[NumBase]+Solde ;
         END else
         BEGIN
         TabTvaEnc[5]:=TabTvaEnc[5]+Solde ;
         END ;
      END ;
    END ;
{Report du tableau des bases sur les échéances}
for Lig:=IMin to IMax do
    BEGIN
    O:=TOBM(TPiece[Lig]) ;
    if ((O.GetMvt('E_ECHE')='X') and (O.GetMvt('E_NUMECHE')>0)) then
       BEGIN
       Solde:=O.GetMvt('E_DEBIT')-O.GetMvt('E_CREDIT') ;
       If (TTC<>0) And OkTvaSurEnc Then
         BEGIN
         Coef:=Solde/TTC ;
         for i:=1 to 4 do If OkTvaSurEnc Then O.PutMvt('E_ECHEENC'+IntToStr(i),Arrondi(TabTvaEnc[i]*Coef,V_PGI.OkDecV)) ;
         If OkTvaSurEnc Then O.PutMvt('E_ECHEDEBIT',Arrondi(TabTvaEnc[5]*Coef,V_PGI.OkDecV)) ;
         If OkTvaSurEnc Then O.PutMvt('E_EMETTEURTVA','X') ;
         END ;
       END ;
    If OkCtr And (O.GetMvt('E_CONTREPARTIEGEN')='') And (O.GetMvt('E_CONTREPARTIEAUX')='') Then
      BEGIN
      GetContepartieImp(NatureJal='BQE',CptLuJ.CollTiers,TPiece,O,Lig,CtrGene,CtrAuxi,InfoImp,IMin,IMax) ;
      OkMajCtr:=FALSE ;
      If CtrGene<>'' Then BEGIN OkMajCtr:=TRUE ; O.PutMvt('E_CONTREPARTIEGEN',CtrGene) ; END ;
      If CtrAuxi<>'' Then BEGIN OkMajCtr:=TRUE ; O.PutMvt('E_CONTREPARTIEAUX',CtrAuxi) ; END ;
      If OkMajCtr And OkMajCtr Then MajCtrAna(O,CtrGene,CtrAuxi) ;
      END ;
    END ;
1:
If OnEstSurUnBordereau And (IMax<TPiece.Count-1) Then GoTo 0 ;
{Mise à jour fichier}
for Lig:=0 to TPiece.Count-1 do
    BEGIN
    O:=TOBM(TPiece[Lig]) ;
    StUp:=O.StPourUpdate ; if StUp='' then Continue ;
    SQL:='UPDATE ECRITURE SET '+StUp+' Where  '+WhereEcriture(tsGene,M,False,OnEstSurUnBordereau)
        +' AND E_NUMLIGNE='+IntToStr(O.GetMvt('E_NUMLIGNE'))+' AND E_NUMECHE='+IntToStr(O.GetMvt('E_NUMECHE')) ;
    ll:=ExecuteSQL(SQL) ;
    if ll<>1 then
      BEGIN
      V_PGI.IoError:=oeUnknown ; Break ;
      END ;
    END ;
{Dispose mémoire}
VideListe(TPiece) ; TPiece.Free ;
END ;


Function StDate2Date(St : String ; Var YY,MM,JJ : Word) : tDateTime ;
BEGIN
YY:=StrToInt(Copy(St,1,4)) ; MM:=StrToInt(Copy(St,5,2)) ; JJ:=StrToInt(Copy(St,7,2)) ;
Result:=EncodeDate(YY,MM,JJ) ;
END ;

Procedure TraitementSurEcritureIntegrees(Var InfoImp : TInfoImport) ;
Var St{,St1} : String ;
    JalP,ExoP,QualP : String ;
    NumP : Integer ;
    DateP : TDateTime ;
    i : Integer ;
    M : RMvt ;
    Q : TQuery ;
    FicRap : TextFile ;
    StR,OldJal : String ;
    OkRapport : Boolean ;
    NumDeb,NumFin : Integer ;
    PasDeMvtIntegres : Boolean ;
    OkCtr,OkTvaEnc,OkTP : Boolean ;
    SDateP : String ;
    PerP,OldPerP,PP : Word ;
    YY,JJ,OldYY : Word ;
BEGIN
If InfoImp.Lequel<>'FEC' Then Exit ;
PasDeMvtIntegres:=FALSE ;
If (InfoImp.ListePieceIntegre=NIL) Or (InfoImp.ListePieceIntegre.Count=0) Then PasDeMvtIntegres:=TRUE ;
OkRapport:=FALSE ; NumDeb:=999999999 ; NumFin:=0 ; OldJal:='' ; OldPerP:=0 ; OldYY:=0 ;
EcrireRapportDebutIntegration(PasDeMvtIntegres,InfoImp,FicRap,OkRapport) ;
If PasDeMvtIntegres Then Exit ;
InitMove(InfoImp.ListePieceIntegre.Count,'') ;
For i:=0 To InfoImp.ListePieceIntegre.Count-1 Do
  BEGIN
  MoveCur(FALSE) ;
  St:=InfoImp.ListePieceIntegre[i] ;
  JalP:=Trim(Copy(St,1,3)) ; 
  SDateP:=Copy(St,16,8) ; StDate2Date(SdateP,YY,PerP,JJ) ; PerP:=PerP+YY ;
  If i=0 Then BEGIN OldJal:=JalP ; OldPerP:=PerP ; OldYY:=YY ; END ;
  If OkRapport And ((JalP<>OldJal) Or (VH^.RecupPCL And (PerP<>OldPerP)) Or (ModeJal(OldJal,InfoImp)<>mPiece))Then
    BEGIN
    StR:=' Journal '+OldJal+' - Pièce(s) Importée(s) du N° '+IntToStr(NumDeb)+' au N° '+IntToStr(NumFin) ;
    If VH^.RecupPCL Then StR:=StR+' période du '+FormatFloat('00',OldPerP-OldYY)+'/'+FormatFloat('0000',OldYY) ;
    Writeln(FicRap,StR) ;
    NumDeb:=999999999 ; NumFin:=0 ;
    OldJal:=JalP ; OldPerP:=PerP ; OldYY:=YY ;
    END ;
  ExoP:=Copy(St,4,3) ;
  QualP:=Copy(St,7,1) ;
  NumP:=StrToInt(Copy(St,8,8)) ;
  If NumP<NumDeb Then NumDeb:=NumP ;
  If NumP>NumFin Then NumFin:=NumP ;
  (*
  St1:=Copy(St,16,8) ;
  Move(St1[1],DateP,8) ;
  *)
  DateP:=StDate2Date(SdateP,YY,PP,JJ) ;
  OkCtr:=St[24]='X' ; OkTvaEnc:=St[25]='X' ; OkTP:=St[26]='X' ;
  If OkCtr Or OkTvaEnc Or OkTP Then
    BEGIN
    Q:=OpenSQL('Select * from Ecriture where E_JOURNAL="'+JalP+'"'
              +' AND E_EXERCICE="'+ExoP+'"'
              +' AND E_DATECOMPTABLE="'+USDateTime(DateP)+'"'
              +' AND E_NUMEROPIECE='+IntToStr(NumP)
              +' AND E_QUALIFPIECE="'+QualP+'" ',True) ;
    if Not Q.EOF then
      BEGIN
      M:=MvtToIdent(Q,fbGene,False) ; Ferme(Q) ;
      If OkCtr Or OkTvaEnc Then TraitementSurEcr2(OkCtr,OkTvaEnc,M,FALSE,InfoImp) ;
      If OkTP Then GenerePiecesPayeur(M) ;
      END Else Ferme(Q) ;
    END ;
  (*
  CleIT:=StOkTraite+StOkTvaEnc+Format_String(IdentPiece.JalP,3)+Format_String(QUELEXODT(IdentPiece.DateP),3)+
  Format_String(IdentPiece.QualP,1)+FormatFloat('00000000',IdentPiece.NumPDef)+StDate ;
  *)
  END ;
FiniMove ;
If OkRapport Then
  BEGIN
//  Str:=' ' ; Writeln(FicRap,StR) ;
  StR:=TraduireMemoire(' Journal ')+OldJal+TraduireMemoire(' - Pièce(s) Importée(s) du N° ')+IntToStr(NumDeb)+TraduireMemoire(' au N° ')+IntToStr(NumFin) ;
  If VH^.RecupPCL Then StR:=StR+' période du '+FormatFloat('00',OldPerP-oldYY)+'/'+FormatFloat('0000',oldYY) ;
  Writeln(FicRap,StR) ;
  EcrireSeparateur(FicRap) ;
  StR:=TraduireMemoire('Traitement commencé le '+DateToStr(InfoImp.DebutTraitement)+' à '+TimeToStr(InfoImp.DebutTraitement)) ;
  Writeln(FicRap,StR) ;
  InfoImp.FinTraitement:=Now ;
  StR:=TraduireMemoire('Traitement terminé le '+DateToStr(InfoImp.FinTraitement)+' à '+TimeToStr(InfoImp.FinTraitement)) ;
  Writeln(FicRap,StR) ;
  EcrireSeparateur(FicRap) ;
  EcrireRapportRejets(FicRap,InfoImp) ;
  EcrireRapportDoublons(FicRap,InfoImp) ;
  END ;
If OkRapport Then CloseFile(FicRap) ;
END ;

Function StReqAnal(Q : TQuery ; JalP,ExoP : String ; DateP : tDateTime) : String ; { Requete paramatrée sur la table ANALYTIQE}
Var St : String ;
BEGIN
St:=' Select Sum(Y_DEBIT) as DP, Sum(Y_CREDIT) as CP, ' ;
St:=St+' Sum(Y_DEBITDEV) as DD, Sum(Y_CREDITDEV) as CD, ' ;
{JP 21/06/06 : Un oubli !!
St:=St+' Sum(Y_DEBITEURO) as DE, Sum(Y_CREDITEURO) as CE, ' ;}
St:=St+' Max(Y_NUMVENTIL) AS MAXNV, Y_AXE AS THEAXE From ANALYTIQ Where ' ;
St:=St+' Y_JOURNAL="'+JalP+'" ' ;
St:=St+' And Y_EXERCICE="'+ExoP+'" ' ;
St:=St+' And Y_DATECOMPTABLE="'+USDateTime(DateP)+'" ' ;
St:=St+' And Y_NUMEROPIECE='+IntToStr(Q.FindField('E_NUMEROPIECE').AsInteger)+' ' ;
St:=St+' And Y_NUMLIGNE='+IntToStr(Q.FindField('E_NUMLIGNE').AsInteger)+' ' ;
St:=St+' And Y_DEVISE="'+Q.FindField('E_DEVISE').AsString+'" ' ;
St:=St+' And Y_ETABLISSEMENT="'+Q.FindField('E_ETABLISSEMENT').AsString+'" ' ;
St:=St+' And Y_QUALIFPIECE="'+Q.FindField('E_QUALIFPIECE').AsString+'" ' ;
St:=St+' Group By Y_AXE ' ;
Result:=St ;
END ;

Procedure AjouteMvtPbA(Q : tQuery ; DP,CP,DD,CD,DE,CE,MP,MD,ME : Double ; Axe,JalP,ExoP : String ; DateP : tDateTime ; NumV : Integer) ;
Var QY : TQuery ;
BEGIN
QY:=PrepareSQL('SELECT * FROM ANALYTIQ WHERE Y_EXERCICE="'+W_W+'"',FALSE) ;
QY.Open ; QY.Insert ; InitNew(QY) ;
QY.FindField('Y_GENERAL').AsString:=Q.FindField('E_GENERAL').AsString ;
QY.FindField('Y_AXE').AsString:=Axe ;
QY.FindField('Y_DATECOMPTABLE').AsDateTime:=DateP ;
QY.FindField('Y_NUMEROPIECE').AsInteger:=Q.FindField('E_NUMEROPIECE').AsInteger ;
QY.FindField('Y_NUMLIGNE').AsInteger:=Q.FindField('E_NUMLIGNE').AsInteger ;
QY.FindField('Y_SECTION').AsString:=VH^.Cpta[AxeToFb(Axe)].Attente ;
QY.FindField('Y_EXERCICE').AsString:=ExoP ;
QY.FindField('Y_DEBIT').AsFloat:=DP ;
QY.FindField('Y_CREDIT').AsFloat:=CP ;
QY.FindField('Y_REFINTERNE').AsString:='Ré-équilibrage pendant import' ;
QY.FindField('Y_LIBELLE').AsString:='' ;
QY.FindField('Y_NATUREPIECE').AsString:=Q.FindField('E_NATUREPIECE').AsString ;
QY.FindField('Y_QUALIFPIECE').AsString:=Q.FindField('E_QUALIFPIECE').AsString ;
QY.FindField('Y_TYPEANALYTIQUE').AsString:='-' ;
QY.FindField('Y_ETABLISSEMENT').AsString:=Q.FindField('E_ETABLISSEMENT').AsString ;
QY.FindField('Y_DEVISE').AsString:=Q.FindField('E_DEVISE').AsString ;
QY.FindField('Y_DEBITDEV').AsFloat:=DD ;
QY.FindField('Y_CREDITDEV').AsFloat:=CD ;
QY.FindField('Y_TAUXDEV').AsFloat:=Q.FindField('E_TAUXDEV').AsFloat ;
If Arrondi(MP,V_PGI.OkDecV)<>0 Then QY.FindField('Y_POURCENTAGE').AsFloat:=((DP+CP)*100)/MP ;
QY.FindField('Y_TOTALECRITURE').AsFloat:=MP ;
QY.FindField('Y_JOURNAL').AsString:=JalP ;
QY.FindField('Y_NUMVENTIL').AsInteger:=NumV+1 ;
QY.FindField('Y_TOTALDEVISE').AsFloat:=MD ;
{JP 22/06/06 : Oubli de champ Euro !!
QY.FindField('Y_TOTALEURO').AsFloat:=ME ;}
QY.FindField('Y_ECRANOUVEAU').AsString:=Q.FindField('E_ECRANOUVEAU').AsString ;
//QY.FindField('Y_CREERPAR').AsString:='IMP' ;
SetCreerPar(QY,'Y_CREERPAR') ;
QY.FindField('Y_CONTREPARTIEGEN').AsString:=Q.FindField('E_CONTREPARTIEGEN').AsString ;
QY.FindField('Y_CONTREPARTIEAUX').AsString:=Q.FindField('E_CONTREPARTIEAUX').AsString ;
QY.FindField('Y_AUXILIAIRE').AsString:=Q.FindField('E_AUXILIAIRE').AsString ;
QY.FindField('Y_PERIODE').AsInteger:=Q.FindField('E_PERIODE').AsInteger ;
QY.FindField('Y_SEMAINE').AsInteger:=Q.FindField('E_SEMAINE').AsInteger ;
QY.Post ;
Ferme(QY) ;
END ;

Procedure TraitementSurAnalytiqueDesequilibree(Var InfoImp : TInfoImport) ;
Var St,St1 : String ;
    JalP,ExoP,GenP,RefP,SDateP : String ;
    DateP : TDateTime ;
    Pb : Boolean ;
    i,j,NumV : Integer ;
    M : RMvt ;
    Q,QA,QDev : TQuery ;
    SetAxe : Set Of Byte ;
    AxeEnCours : String ;
    IAxeEnCours : Byte ;
    AxeOk : Array[1..5] Of Boolean ;
    Decim : Integer ;
    Dev : String ;
    MP,MD,ME,MPA,MDA,MEA,DeltaP,DeltaD,DeltaE,DP,CP,DD,CD,DE,CE : Double ;
BEGIN
If InfoImp.Lequel<>'FEC' Then Exit ; If Not InfoImp.SC.ShuntPbAna Then Exit ;
If Not InfoImp.PbAna Then Exit ; If InfoImp.ListePbAna=Nil Then Exit ;
If InfoImp.ListePbAna.Count<=0 Then Exit ;
If (InfoImp.ListePieceIntegre=NIL) Or (InfoImp.ListePieceIntegre.Count=0) Then Exit ;
InitMove(InfoImp.ListePbAna.Count,'') ; ME := 0;
For i:=0 To InfoImp.ListePbAna.Count-1 Do
  BEGIN
  MoveCur(FALSE) ;
  St:=InfoImp.ListePbAna[i] ;
  JalP:='' ; ExoP:='' ; GenP:='' ; RefP:='' ; Pb:=FALSE ;
  If St<>'' Then JalP:=ReadTokenSt(St) Else Pb:=TRUE ;
  If St<>'' Then SDateP:=ReadTokenSt(St) Else Pb:=TRUE ;
  If St<>'' Then GenP:=ReadTokenSt(St) Else Pb:=TRUE ;
  If St<>'' Then RefP:=ReadTokenSt(St) Else Pb:=TRUE ;
  If Pb Then Continue ;
  DateP:=StrToDate(SDateP) ; ExoP:=QUELEXODT(DateP) ;
  St1:='SELECT E_GENERAL, E_AUXILIAIRE, E_CONTREPARTIEGEN, E_CONTREPARTIEAUX, '+
       'E_DEBIT,E_CREDIT, E_DEBITDEV, E_CREDITDEV,  '+
       'E_DEVISE, E_TAUXDEV, E_COTATION, G_VENTILABLE1, G_VENTILABLE2, G_VENTILABLE3,'+
       'G_VENTILABLE4, G_VENTILABLE5,E_NUMEROPIECE, E_NUMLIGNE, E_ETABLISSEMENT, E_QUALIFPIECE, E_DEVISE, ' + {E_SAISIEEURO, '+}
       'E_NATUREPIECE, E_ETABLISSEMENT, E_TAUXDEV, E_ECRANOUVEAU, E_SEMAINE, E_PERIODE '+
       'FROM ECRITURE LEFT JOIN GENERAUX ON E_GENERAL=G_GENERAL '+
       'where E_JOURNAL="'+JalP+'" AND E_EXERCICE="'+ExoP+'" AND E_DATECOMPTABLE="'+USDateTime(DateP)+'" '+
       'AND E_GENERAL="'+GenP+'" AND E_REFINTERNE="'+REFP+'" ' ;
  If VH^.RecupSISCOPGI Then St1:=St1+ ' AND E_CREERPAR="REC" ' Else St1:=St1+' AND E_CREERPAR="IMP" ' ;
  Q:=OpenSQL(St1,TRUE) ;
  While Not Q.Eof Do
    BEGIN
    SetAxe:=[] ; FillChar(AxeOk,SizeOf(AxeOk),#0) ;
    For j:=1 To 5 Do If Q.FindField('G_VENTILABLE'+IntToStr(j)).AsString='X' Then SetAxe:=SetAxe+[j] ;
    Dev:=Q.FindField('E_DEVISE').AsString ; Decim:=V_PGI.OkDecV ;
    If Dev=V_PGI.DevisePivot Then
      BEGIN
      Decim:=V_PGI.OkDecV ;
      END Else
      BEGIN
      QDev:=OpenSQL('SELECT D_DECIMALE FROM DEVISE WHERE D_DEVISE="'+Dev+'" ',TRUE) ;
      If Not QDev.Eof Then Decim:=QDev.Fields[0].AsInteger ;
      Ferme(QDev) ;
      END ;
    St1:=StReqAnal(Q,JalP,ExoP,DateP) ;
    QA:=OpenSQL(St1,TRUE) ;
    MP:=Arrondi(Q.FindField('E_DEBIT').AsFloat+Q.FindField('E_CREDIT').AsFloat,Decim) ;
    MD:=Arrondi(Q.FindField('E_DEBITDEV').AsFloat+Q.FindField('E_CREDITDEV').AsFloat,Decim) ;
    While Not QA.Eof Do
      BEGIN
      AxeEnCours:=QA.FindField('THEAXE').AsString ;
      NumV:=QA.FindField('MAXNV').AsInteger ;
      If Length(AxeEnCours)>=2 Then IAxeEnCours:=StrToInt(Copy(AxeEnCours,2,1)) Else IAxeEnCours:=0 ;
      If IAxeEnCours In [1..5] Then AxeOk[IAxeEnCours]:=TRUE ;
      MPA:=Arrondi(QA.FindField('DP').AsFloat+QA.FindField('CP').AsFloat,Decim) ;
      MDA:=Arrondi(QA.FindField('DD').AsFloat+QA.FindField('CD').AsFloat,Decim) ;
      {JP 22/06/06 : Reste euros !!
      MEA:=Arrondi(QA.FindField('DE').AsFloat+QA.FindField('CE').AsFloat,Decim) ;
      DeltaE:=Arrondi(ME-MEA,Decim) ;}
      DeltaE:= 0;
      DeltaP:=Arrondi(MP-MPA,Decim) ; DeltaD:=Arrondi(MD-MDA,Decim) ;
      If (DeltaP<>0) Or (DeltaD<>0) Or (DeltaE<>0) Then
        BEGIN
        DP:=0 ; CP:=0 ; DD:=0 ; CD:=0 ; DE:=0 ; CE:=0 ;
        If Arrondi(Q.FindField('E_DEBIT').AsFloat+Q.FindField('E_DEBITDEV').AsFloat,2)<>0 Then
          BEGIN
          If DeltaP>0 Then BEGIN DP:=DeltaP ; DD:=DeltaD ; {DE:=DeltaE ;} END Else
            If DeltaP<0 Then BEGIN DP:=DeltaP ; DD:=DeltaD ; {DE:=DeltaE} ; END ;
          END Else
          BEGIN
          If DeltaP>0 Then BEGIN CP:=DeltaP ; CD:=DeltaD ; {CE:=DeltaE} ; END Else
            If DeltaP<0 Then BEGIN CP:=DeltaP ; CD:=DeltaD ; {CE:=DeltaE} ; END ;
          END ;
        AjouteMvtPbA(Q,DP,CP,DD,CD,DE,CE,MP,MD,ME,AxeEnCours,JalP,ExoP,DateP,NumV) ;
        END ;
      QA.Next ;
      END ;
    For j:=1 To 5 Do If (j In SetAxe) And (Not AxeOk[j]) Then
      BEGIN
      DP:=0 ; CP:=0 ; DD:=0 ; CD:=0 ; DE:=0 ; CE:=0 ; AxeEnCours:='A'+IntToStr(j) ;
      If Arrondi(Q.FindField('E_DEBIT').AsFloat+Q.FindField('E_DEBITDEV').AsFloat,2)<>0 Then
        BEGIN DP:=MP ; DD:=MD ; DE:=ME ; END Else
        BEGIN CP:=MP ; CD:=MD ; CE:=ME ; END ;
      AjouteMvtPbA(Q,DP,CP,DD,CD,DE,CE,MP,MD,ME,AxeEnCours,JalP,ExoP,DateP,0) ;
      END ;
    Ferme(QA) ;
    Q.Next ;
    END ;
  Ferme(Q) ;
  END ;
FiniMove ;
END ;

{JP 21/06/06 : Reprise du code de IntegreEcr auquel on force la requête}
function IntegreEcrParBloc(FAss : TForm; Var InfoImp : TInfoImport; ClauseWhere : string) : Boolean;
Var
  QImpEcr,QE,QY : TQuery ;
  OkIntegre : Boolean ;
  FF : TextFile ;
begin
  Result := True;

  FAssImp:=FAss ; OkIntegre:=TRUE ;
  InfoImp.Prefixe:='E' ; InfoImp.Table:='ECRITURE' ;
  if InfoImp.Lequel='FBE' then BEGIN InfoImp.Prefixe:='BE' ; InfoImp.Table:='BUDECR' ; END else
  if InfoImp.Lequel='FOD' then BEGIN InfoImp.Prefixe:='Y' ; InfoImp.Table:='ANALYTIQ' ;END ;
  AnnuleImport:=False ; Erreur:=False ;

  {JP 21/06/06 : Ajout d'une éventuelle restriction de la requête}
//  QImpEcr:=OpenSQL('SELECT * FROM IMPECR WHERE IE_SELECTED="X" AND IE_OKCONTROLE="X" AND IE_INTEGRE="-" ' +
  //                 'AND IE_TYPEECR <> "L" ' + ClauseWhere + ' ORDER BY IE_CHRONO', False);
  QImpEcr:=OpenSQL('SELECT * FROM IMPECR WHERE IE_SELECTED="X" AND IE_OKCONTROLE="X" AND IE_INTEGRE="-" ' +
                   'AND IE_TYPEECR <> "L" AND IE_CHRONO < 10000 ORDER BY IE_CHRONO', False);
  try
    QImpEcr.UpdateMode:=upWhereChanged ;
    if (QImpecr.EOF) then
    BEGIN
      AnnuleImport:=True ;
      Ferme(QImpEcr) ;
      If (FAssImp<>nil) Then
      BEGIN
        if (InfoImp.Format='RAP') or (InfoImp.Format='CRA') then THMsgBox(FAssImp.FindComponent('HM')).Execute(50,'','')
                                           else THMsgBox(FAssImp.FindComponent('HM')).Execute(42,'','') ;
      END ;
      TraitementSurEcritureIntegrees(InfoImp) ;
      EcrireRapportDivers(0,InfoImp) ;
      OkIntegre:=FALSE ;
    END ;

    If OkIntegre Then
    BEGIN
      Try
        BEGINTRANS ;
        If VH^.RecupCEGID Then
          BEGIN
          AssignFile(FF,'C:\CHRONOCEGID.TXT') ; Rewrite(FF) ;
          END ;
        InitMove(1000,'') ; MoveCur(False) ;
        InfoImp.NbLigIntegre:=0 ; InfoImp.TotDeb:=0 ; InfoImp.TotCred:=0 ; InfoImp.NbPiece:=0 ;
        If InfoImp.Lequel='FBE' Then QE:=PrepareSQL('SELECT * FROM '+InfoImp.Table+' WHERE '+InfoImp.Prefixe+'_BUDJAL="'+W_W+'"',FALSE)
                                Else QE:=PrepareSQL('SELECT * FROM '+InfoImp.Table+' WHERE '+InfoImp.Prefixe+'_JOURNAL="'+W_W+'"',FALSE) ;
        QE.Open ;
        if (InfoImp.Format<>'RAP') and (InfoImp.Format<>'CRA') and (InfoImp.Format<>'MP') and (InfoImp.Prefixe='E') then
        BEGIN
          QY:=PrepareSQL('SELECT * FROM ANALYTIQ WHERE Y_JOURNAL="'+W_W+'"',FALSE) ;
          QY.Open ;
        END ;

        if (InfoImp.Lequel<>'FEC') or ((InfoImp.Lequel='FEC') and (InfoImp.Format='CRA')) then
        BEGIN
          InfoImp.ForceNumPiece:=FALSE ;
          IntegreParam(QImpEcr,QE,QY,InfoImp) ;
        END
        else
          IntegreGenerales(QImpEcr,QE,QY,InfoImp,FF) ;
        FiniMove ;

        if (InfoImp.Format<>'RAP') and (InfoImp.Format<>'CRA') and (InfoImp.Format<>'MP') and (InfoImp.Prefixe='E') then BEGIN Ferme(QY) ; END ;
        if AnnuleImport or Erreur then ROLLBACK else COMMITTRANS ;
        if (InfoImp.Lequel<>'FBE') then
        BEGIN
          if not AnnuleImport then MajSoldeCompteImport(InfoImp) ;
        END ;

        if not AnnuleImport then
          {JP 21/06/06 : Ajout d'une éventuelle clause de limitation des enregistrement}
          ExecuteSQL('UPDATE IMPECR SET IE_INTEGRE="X" WHERE IE_SELECTED="X" AND IE_OKCONTROLE="X" AND IE_TYPEECR<>"L" ' + ClauseWhere) ;

        if (FAss=nil) then CompteRenduBatch(1,InfoImp) ;
        if not AnnuleImport then TraitementSurEcritureIntegrees(InfoImp) ;
        if not AnnuleImport then TraitementSurAnalytiqueDesequilibree(InfoImp) ;

        {JP 21/06/06 : Pour Interrompre la boucle dans IntegreEcr}
        Result := (not AnnuleImport) and (not Erreur);

        AnnuleImport:=False ;
        if (not AnnuleImport) then
        BEGIN
          If RecupSISCO And (VH^.RecupPCL Or InfoImp.LettrageSISCOEnImport) Then TraitementFinRecupSISCO(InfoImp) ;
          If RecupSERVANT then TraitementFinRecupSERVANT(InfoImp) ;
        END ;

        {$i-}
        If VH^.RecupCegid Then CloseFile(FF) ;
        {$i+}
      Except
        RollBack ;
        Result := False;
        {$i-}
        If VH^.RecupCegid Then CloseFile(FF) ;
        {$i+}
      end ;
    end;

  finally
      If OkIntegre Then
      begin
        Ferme(QE);
        Ferme(QY);
        Ferme(QImpEcr);
      end;
  end;
end;

{JP 21/06/06 : Gestion des gros volumes : le code IntegreEcr a été déplacé dans IntegreEcrParBloc.
               la diférence entre IntegreEcrParBloc et l'ancien IntegreEcr est que l'on limite la
               résultat de la requête sur ImpEcr par bloc (par pièce, par journaux ou toute la table).
               Le traitement est ainsi appelé en boucle}
procedure IntegreEcr(FAss : TForm ; Var InfoImp : TInfoImport) ;
(*
var
  Ini : TIniFile;
  Cut : TQuery;
  Typ : string;
  Req : string;
  OK  : Boolean;
  *)
begin
  (*
  Ini := TIniFile.Create(ExtractFilePath(Application.ExeName) + 'recup.ini');
  Typ := Ini.ReadString('RUPTURE', 'TYPE', '');
  FreeAndNil(Ini);
  if Typ <> '' then begin
    if UpperCase(Typ) = 'JOURNAL' then
      Cut := OpenSQL('SELECT DISTINCT IE_JOURNAL FROM IMPECR ORDER BY IE_JOURNAL', True)
    else
      Cut := OpenSQL('SELECT IE_NUMPIECE, IE_JOURNAL, IE_QUALIFPIECE, MONTH(IE_DATECOMPTABLE) AS IE_MONTH FROM IMPECR ' +
                     'GROUP BY IE_NUMPIECE, IE_JOURNAL, IE_QUALIFPIECE ' +
                     'ORDER BY IE_JOURNAL, IE_NUMPIECE', True);
    Ok := True;
    try
      if Cut.EOF then
        {A priori, cela signifie qu'il n'y a pas d'écriture !!
         On reprend le code comme il était autrefois}
        IntegreEcrParBloc(FAss, InfoImp, '')
      else
        while not Cut.EOF do begin
          if UpperCase(Typ) = 'JOURNAL' then
            Req := ' AND IE_JOURNAL = "' + Cut.FindField('IE_JOURNAL').AsString + '"'
          else
            Req := ' AND IE_JOURNAL = "'     + Cut.FindField('IE_JOURNAL'    ).AsString + '"' +
                   ' AND IE_QUALIFPIECE = "' + Cut.FindField('IE_QUALIFPIECE').AsString + '"' +
                   ' AND MONTH(IE_DATECOMPTABLE) = ' + Cut.FindField('IE_MONTH').AsString + ' ' +
                   ' AND IE_NUMPIECE = '     + Cut.FindField('IE_NUMPIECE'   ).AsString;

          if not IntegreEcrParBloc(FAss, InfoImp, Req) then Break;
          Cut.Next;
        end;
    finally
      Ferme(Cut);
    end;
  end
  else
    *)
    {Pas de rupture demandée, on reprend le code comme il était autrefois}
    IntegreEcrParBloc(FAss, InfoImp, '');
end;

Procedure RecupMvtImport(Q : TQuery ; Var Mvt : TFMvtImport) ;
BEGIN
InitMvtImport(Mvt) ;
If Not Q.Eof Then With Mvt Do
  BEGIN
  IE_GENERAL:=Q.FindField('IE_GENERAL').AsString ;
  IE_AUXILIAIRE:=Q.FindField('IE_AUXILIAIRE').AsString ;
  IE_DATECOMPTABLE:=Q.FindField('IE_DATECOMPTABLE').ASDateTime ;
  IE_DEVISE:=Q.FindField('IE_DEVISE').AsString ;
  IE_LETTRAGE:=Q.FindField('IE_LETTRAGE').AsString ;
  IE_JOURNAL:=Q.FindField('IE_JOURNAL').AsString ;
  IE_DEBIT:=Q.FindField('IE_DEBIT').AsFloat ;
  IE_CREDIT:=Q.FindField('IE_CREDIT').ASFloat ;
  IE_REFINTERNE:=Q.FindField('IE_REFINTERNE').AsString ;
  IE_LIBELLE:=Q.FindField('IE_LIBELLE').AsString ;
  IE_DEBITDEV:=Q.FindField('IE_DEBITDEV').AsFloat ;
  IE_CREDITDEV:=Q.FindField('IE_CREDITDEV').ASFloat ;
  IE_NATUREPIECE:=Q.FindField('IE_NATUREPIECE').AsString ;
  END
END ;

Procedure PrepareLettrage(OkInc : Boolean ; Var InfoImp : TInfoImport ; TLett : TList ; Var NbLett : Integer ; Var Mvt : TFMvtImport) ;
Var DecDev : Integer ;
    Quotite : Double ;
    L : TL_Rappro ;
BEGIN
If OkInc Then Inc(NbLett) ;
RecupDevise(Mvt.IE_DEVISE,DecDev,Quotite,InfoImp.ListeDev) ;
L:=TL_Rappro.Create ;
L.General:=Mvt.IE_GENERAL ; L.Auxiliaire:=Mvt.IE_AUXILIAIRE ;
L.DateC:=Mvt.IE_DATECOMPTABLE ; L.DateE:=Mvt.IE_DATEECHEANCE ; L.DateR:=Mvt.IE_DATEREFEXTERNE ;
L.RefI:=Mvt.IE_REFINTERNE ; L.RefL:=Mvt.IE_REFLIBRE ;
L.RefE:=Mvt.IE_REFEXTERNE; L.Lib:=Mvt.IE_LIBELLE ;
L.Jal:=Mvt.IE_JOURNAL ; L.Numero:=Mvt.IE_NUMPIECE ;
L.NumLigne:=Mvt.IE_NUMLIGNE ; L.NumEche:=Mvt.IE_NUMECHE ;
L.CodeL:=Mvt.IE_LETTRAGE ;
L.TauxDEV:=Mvt.IE_TAUXDEV ;
L.CodeD:=Mvt.IE_DEVISE ;
L.Decim:=DecDev ;
L.Debit:=Mvt.IE_DEBIT ; L.Credit:=Mvt.IE_CREDIT ;
L.DebDev:=Mvt.IE_DEBITDEV ; L.CredDev:=Mvt.IE_CREDITDEV ;
//L.DebEuro:=Mvt.IE_DEBITEURO ; L.CredEuro:=Mvt.IE_CREDITEURO ;
//L.DebEuro:=0 ; L.CredEuro:=0 ;
L.Nature:=Mvt.IE_NATUREPIECE ;
L.Facture:=((L.Nature='FC') or (L.Nature='FF') or (L.Nature='AC') or (L.Nature='AF')) ;
L.Client:=((L.Nature='FC') or (L.Nature='AC') or (L.Nature='RC') or (L.Nature='OC')) ;
L.Solution:=0 ; L.Exo:=QUELEXODT(Mvt.IE_DATECOMPTABLE) ;
L.EditeEtatTva:=Mvt.IE_LIBREBOOL1='X' ;
//L.saisieEuro:= TRUE;
TLett.Add(L) ;
END ;

Procedure ChercheMvtEnPlusLettres(TLett : TList ; Var NbLett : Integer ; Var Mvt : TFMvtImport) ;
//Var i : Integer ;
BEGIN
(*
for i:=0 to LEcr.Count-1 do
  BEGIN
  TL:=TL_Rappro(LEcr[i]) ;
  if Total then
     BEGIN
     if Not LettEuro then
        BEGIN
        if TL.Debit<>0 then TL.Couv:=TL.Debit else TL.Couv:=TL.Credit ;
        END else
        BEGIN
        if TL.DebEuro<>0 then TL.CouvEuro:=TL.DebEuro else TL.CouvEuro:=TL.CredEuro ;
        END ;
     END else
     BEGIN
     if TL.Debit>0 then LD.Add(LEcr[i]) else LC.Add(LEcr[i]) ;
     END ;
  END ;
*)
END ;

procedure LettrageImport(Var InfoImp : TInfoImport) ;
var SQL,OldLettrage : String ;
    QImp : TQuery ;
    TLett:TList ;
    OkALettrer : Boolean ;
    NbLettImp,NbLett : integer ;
    Auxi,Gene,CodeLettre,Devise : String17 ;
BEGIN
SQL:='SELECT IE_GENERAL,IE_AUXILIAIRE,IE_DATECOMPTABLE,IE_DEVISE,IE_LETTRAGE,IE_JOURNAL,IE_DEBIT,IE_CREDIT,IE_REFINTERNE,'
     +'IE_LIBELLE,IE_DEBITDEV,IE_CREDITDEV,IE_NATUREPIECE  FROM IMPECR WHERE IE_TYPEECR="L" '
     +' ORDER BY IE_AUXILIAIRE,IE_GENERAL,IE_DEVISE,IE_ETATLETTRAGE, IE_LETTRAGE, IE_DATECOMPTABLE, IE_NUMPIECE, IE_NUMLIGNE, IE_NUMECHE ' ;
QImp:=OpenSQL(SQL,True) ;
if (QImp.EOF) then BEGIN Ferme(QImp) ; Exit ; END ;
InitMove(RecordsCount(QImp),'') ;
TLett:=TList.Create ;
InfoImp.ListeMvtTrouve:=HTStringList.Create ;
InfoImp.ListeMvtTrouve.Sorted:=TRUE ;
InfoImp.ListeMvtTrouve.Duplicates:=DupIgnore ;
New(MvtImport) ;
While not QImp.Eof do
  BEGIN
  VideListe(TLett) ; OldLettrage:='' ; NbLett:=0 ; NbLettImp:=0 ;
  RecupMvtImport(QImp,MvtImport^) ;
  Auxi:=MvtImport^.IE_AUXILIAIRE ;
  Gene:=MvtImport^.IE_GENERAL ;
  CodeLettre:=MvtImport^.IE_LETTRAGE ;
  Devise:=MvtImport^.IE_DEVISE ;
  OkALettrer:=True ;
  While not (QImp.Eof) and ((QImp.Fields[1].AsString=Auxi)
    and (QImp.Fields[0].AsString=Gene)
    and (QImp.Fields[4].AsString=CodeLettre)
    and (QImp.Fields[3].AsString=Devise)) do
    BEGIN
    Inc(NbLettImp) ;
    RecupMvtImport(QImp,MvtImport^) ;
    OkALettrer:=ChercheDoublon(InfoImp,MvtImport^,TRUE)=1 ;
    if OkALettrer then PrepareLettrage(TRUE,InfoImp,TLett,NbLett,MvtImport^) ;
    QImp.Next ; MoveCur(False) ;
    END ;
  OkALettrer:=(OkALettrer and (NbLettImp=NbLett)) ;
  if OkALettrer then
    BEGIN
    ChercheMvtEnPlusLettres(Tlett,NbLett,MvtImport^) ;
    LettrerUnPaquet(TLett,False,False) ;
    END ;
  END ;
FiniMove ;
VideListe(TLett) ; TLett.Free ;
InfoImp.ListeMvtTrouve.Clear ; InfoImp.ListeMvtTrouve.Free ; InfoImp.ListeMvtTrouve:=Nil ;
Ferme(QImp) ;
Dispose(MvtImport) ;
END ;


Procedure RempliListeInfoImp(Var InfoImp : TInfoImport) ;
Var QFiche : TQFiche ;
    CptLu : TCptLu ;
    Q : TQuery ;
    i : Integer ;
BEGIN
If InfoImp.Lequel='FBE' Then
  BEGIN
  InitRequete(QFiche[4],4) ;
  //Q:=OpenSQL('Select IE_JOURNAL,IE_GENERAL,IE_AUXILIAIRE,IE_SECTION FROM IMPECR GROUP BY IE_JOURNAL,IE_GENERAL,IE_AUXILIAIRE,IE_SECTION',TRUE) ;
  Q:=OpenSQL('Select IE_JOURNAL FROM IMPECR GROUP BY IE_JOURNAL',TRUE) ;
  While Not Q.Eof Do
    BEGIN
    Fillchar(CptLu,SizeOf(CptLu),#0) ;
    CptLu.Cpt:=Q.Fields[0].AsString ;
    If AlimLTabCptLu(4,QFiche[4],InfoImp.LJalLu,InfoImp.ListeCptFaux,CptLu) Then ;
    Q.Next ;
    END ;
  Ferme(Q) ;
  Ferme(QFiche[4]) ;
  END Else
  BEGIN
  For i:=0 To 3 Do InitRequete(QFiche[i],i) ;
  If InfoImp.Sc.UseCorresp Then InitRequete(QFiche[5],5) ;
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
  If InfoImp.Sc.UseCorresp Then Ferme(QFiche[5]) ;
  END ;
END ;

Function PbPiece(St : String ; InfoImp : TInfoImport ; What : Integer ; QFiche : TQFiche) : Boolean ;
Var Decal,DebNum,LongNum : Integer ;
    NumP : Integer ;
    JalP,JalP1 : String ;
    ClePF,St1 : String ;
    i : Integer ;
    LeJal : tCptImport ;
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
St1:=Trim(Copy(St,DebNum,LongNum)) ;
If Not VerifEntier(St1) Then Exit ;
NumP:=StrToInt(Trim(Copy(St,DebNum,LongNum))) ;
JalP1:=JalP ;
If (InfoImp.SC.CorrespJal) And (QFiche[5]<>NIL) Then
  BEGIN
  Fillchar(LeJal,SizeOf(LeJal),#0) ; LeJal.Cpt:=JalP ;
  If TraiteCorrespCpt(8,QFiche,LeJal,InfoImp) Then JalP1:=LeJal.Cpt ;
  END ;
ClePF:=Format_String(JalP1,3)+Format_String(' ',1)+FormatFloat('00000000',NumP) ;
Case What Of
  0 : If InfoImp.ListeEntetePieceFausse.Find(ClePF,i) Then Result:=TRUE ; // Piece Fausse ;
  1 : If InfoImp.ListeEnteteDoublon.Find(ClePF,i) Then Result:=TRUE ; // Doublon ;
  END ;
END ;

Procedure FaitFichierRejet(InfoImp : TInfoImport) ;
Var Fichier,NewFichier1 : TextFile ;
    St : String ;
    OkFormat : Boolean ;
    QFiche : TQFiche ;
BEGIN
OkFormat:=(InfoImp.Lequel='FEC') and ((InfoImp.Format='SAA') Or
          (InfoImp.Format='SN2') or (InfoImp.Format='HAL') Or
          (InfoImp.Format='HLI') Or (InfoImp.Format='CGE') Or
          (InfoImp.Format='CGN')) ;
If Not OkFormat Then Exit ;
If Trim(InfoImp.NomFicRejet)='' Then Exit ;
If (InfoImp.ListeEntetePieceFausse=Nil) Or (InfoImp.ListeEntetePieceFausse.Count=0) Then Exit ;
InitMove(1000,'') ;
{$i-} AssignFile(NewFichier1,InfoImp.NomFicRejet) ;  {$i+}
If IoResult<>0 Then Exit ;
Rewrite(NewFichier1) ;
AssignFile(Fichier,InfoImp.NomFic) ; Reset(Fichier) ;
QFiche[5]:=Nil ;
If InfoImp.Sc.UseCorresp Then InitRequete(QFiche[5],5) ;
ReadLn(Fichier,St) ; WriteLn(NewFichier1,St) ; //WriteLn(NewFichier2,St) ;
While Not EOF(Fichier) do
  BEGIN
  MoveCur(FALSE) ; ReadLn(Fichier,St) ;
  If Trim(St)<>'' Then If PbPiece(St,InfoImp,0,QFiche) Then WriteLn(NewFichier1,St) ; //Else WriteLn(NewFichier2,St) ; END ;
  END ;
FiniMove ;
If InfoImp.Sc.UseCorresp Then Ferme(QFiche[5]) ;
CloseFile(Fichier) ; CloseFile(NewFichier1) ;
END ;

Procedure FaitFichierDoublon(InfoImp : TInfoImport) ;
Var Fichier,NewFichier1: TextFile ;
    St : String ;
    OkFormat : Boolean ;
    QFiche : TQFiche ;
BEGIN
If Not InfoImp.CtrlDB Then Exit ;
If Trim(InfoImp.NomFicDoublon)='' Then Exit ;
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
QFiche[5]:=Nil ;
If InfoImp.Sc.UseCorresp Then InitRequete(QFiche[5],5) ;
ReadLn(Fichier,St) ; WriteLn(NewFichier1,St) ;
While Not EOF(Fichier) do
  BEGIN
  MoveCur(FALSE) ; ReadLn(Fichier,St) ;
  If Trim(St)<>'' Then If PbPiece(St,InfoImp,1,QFiche) Then WriteLn(NewFichier1,St)  ;
  END ;
FiniMove ;
If InfoImp.Sc.UseCorresp Then Ferme(QFiche[5]) ;
CloseFile(Fichier) ; CloseFile(NewFichier1) ;
END ;

{***********A.G.L.***********************************************
Auteur  ...... : Yann MORENO
Créé le ...... : 11/10/2005
Modifié le ... : 17/11/2005
Description .. : Verifie si l'on n'essaie pas d'importer une ecriture en devise
Suite ........ : sur un jal non multidevise
Mots clefs ... : FQ15422 multidevise devise
*****************************************************************}
Function ErreurMultiDevise(St : String ; InfoImp : TInfoImport ; Mvt : TFMvtImport ) : Boolean ;
Var
  FmtFic : Integer ;
BEGIN

Result:=FALSE ;

if InfoImp.Format='SAA' then FmtFic:=0 else if InfoImp.Format='SN2' then FmtFic:=0 else
if InfoImp.Format='HLI' then FmtFic:=1 else if InfoImp.Format='HAL' then FmtFic:=2 else
if InfoImp.Format='CGN' then FmtFic:=3 else if InfoImp.Format='CGE' then FmtFic:=4 Else Exit ;

// 17/11/05 On ne relit pas le fichier, on utilise Mvt (prend en charge les traitements antérieurs sur la monnaie)
If (FmtFic > 1) and (St <> '')
and Not EXISTESQL('SELECT * FROM JOURNAL WHERE J_JOURNAL = "'+Mvt.IE_JOURNAL+'" AND J_MULTIDEVISE = "X"')
and (Mvt.IE_DEVISE <> V_PGI.DevisePivot) then
  Result := True
else
  Result := False;

END;

{JP 14/11/05 : FQ 16567 : Teste si les comptes n'ont pas de caractères jokers et si c'est
               le cas, on les remplace par le caractères de bourrage
{---------------------------------------------------------------------------------------}
procedure TesteJokers(var Mvt : TFMvtImport);
{---------------------------------------------------------------------------------------}

    {-----------------------------------------------------------------------}
    procedure RemplaceJoker(var s : string; c : Char);
    {-----------------------------------------------------------------------}
    begin
      s := FindEtReplace(s, '*' , c, True);
      s := FindEtReplace(s, '_' , c, True);
      s := FindEtReplace(s, '%' , c, True);
      s := FindEtReplace(s, '?' , c, True);
      s := FindEtReplace(s, '"' , c, True);
      s := FindEtReplace(s, '''', c, True);
    end;
var
  s : string;
begin
  if ExisteCarInter(Mvt.IE_GENERAL) then begin
    s := Mvt.IE_GENERAL;
    RemplaceJoker(s, VH^.Cpta[fbGene].Cb);
    Mvt.IE_GENERAL := s;
  end;

  if ExisteCarInter(Mvt.IE_AUXILIAIRE) then begin
    s := Mvt.IE_AUXILIAIRE;
    RemplaceJoker(s, VH^.Cpta[fbAux ].Cb);
    Mvt.IE_AUXILIAIRE := s;
  end;

  if ExisteCarInter(Mvt.IE_SECTION) then begin
    s := Mvt.IE_SECTION;
    case Mvt.IE_AXE[2] of
      '1' : RemplaceJoker(s, VH^.Cpta[fbAxe1].Cb);
      '2' : RemplaceJoker(s, VH^.Cpta[fbAxe2].Cb);
      '3' : RemplaceJoker(s, VH^.Cpta[fbAxe3].Cb);
      '4' : RemplaceJoker(s, VH^.Cpta[fbAxe4].Cb);
      '5' : RemplaceJoker(s, VH^.Cpta[fbAxe5].Cb);
    end;
    Mvt.IE_SECTION := s;
  end;
end;

(*
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
If Trim(InfoImp.NomFicRejet)='' Then Exit ;
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
  If Trim(St)<>'' Then BEGIN If PbPiece(St,InfoImp,0) Then WriteLn(NewFichier1,St) Else WriteLn(NewFichier2,St) ; END ;
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
If Trim(InfoImp.NomFicDoublon)='' Then Exit ;
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
  If Trim(St)<>'' Then BEGIN If PbPiece(St,InfoImp,1) Then WriteLn(NewFichier1,St) Else WriteLn(NewFichier2,St) ; END ;
  END ;
FiniMove ;
CloseFile(Fichier) ; CloseFile(NewFichier1) ; CloseFile(NewFichier2) ;
AssignFile(Fichier,InfoImp.NomFic) ; Erase(Fichier) ;
renamefile(StNewFichier2,InfoImp.NomFic) ;
END ;
*)

{$ENDIF EAGLCLINT}
end.

