unit UtilAutoPCL;

interface

Uses inifiles, dbTables, Classes, Sysutils, Forms, TImpFic, FileCtrl, Hent1, HCtrls,ImpFicU ;

Function LireFileIni(NomIni : String ; Var EE : tEnvImpAuto) : Integer ;
Function StErrIni(i : integer) : String ;
Procedure CopieAS400(St,Qui : String) ;
Procedure litSUParam(Var Contact,Rva : String) ;

implementation

Const _NomSectionRepertoire = 'REPERTOIRE' ;
      _NomSectionNomFichier = 'NOMFICHIER' ;
      _NomSectionParametres = 'PARAMETRES' ;
      _NomSectionExe = 'EXE' ;
      _NomSectionExport = 'EXPORT' ;
      _NomSectionRepertoireExport = 'REPERTOIREEXPORT' ;
      _Rep='DIR' ;
      _RepExport='DIR' ;
      _Ecriture='ECRITURE' ;
      _Tiers='TIERS' ;
      _Import='IMPORT' ;
      _Format='FORMAT' ;
      _CodeFormat='CODEFORMAT' ;
      _User='USER' ;
      _Exe='PROG' ;
      _ExeExport='PROG' ;
      _Editeur='EDITEUR' ;
      _VideRep='A_VIDER' ;

Function SectionPresente(St : String ; LNOMSECTION : TStrings) : Boolean ;
Var i : Integer ;
BEGIN
Result:=TRUE ;
For i:=0 To LNOMSECTION.Count-1 Do If LNOMSECTION[i]=St Then Exit ;
Result:=FALSE ;
End ;

Procedure Alim(Var MF : tMasqueFic ; St : String ) ;
Var i : Integer ;
BEGIN
For i:=0 To 9 Do BEGIN If MF[i]='' Then BEGIN MF[i]:=St ; Exit ; END ; END ;
END ;

Procedure LireSection(Qui : Byte ; IniFile : TIniFile ; LL : TStrings ; NomSection : String ; Var EE : tEnvImpAuto) ;
Var i : Integer ;
    St,St1 : String ;
BEGIN
For i:=0 TO LL.Count-1 Do
  BEGIN
  St:=LL[i] ;
  St1:=IniFile.ReadString(NomSection,St,'') ;
  Case Qui Of
    0 : BEGIN //Répertoire
        If St=_Rep Then EE.Rep:=St1 ;
        If St=_VideRep Then EE.VideRep:=St1='OUI' ;
        END ;
    1 : BEGIN //Fichier
        If Pos(_Ecriture,St)>0 Then Alim(EE.MasqueFic[mMasqueEcr],St1) ;
        If Pos(_Tiers,St)>0 Then Alim(EE.MasqueFic[mMasqueAux],St1) ;
        END ;
    2 : BEGIN //Paramètres
        If St=_Import Then EE.Import:=St1 ;
        If St=_Format Then EE.Format:=St1 ;
        If St=_CodeFormat Then EE.CodeFormat:=St1 ;
        If St=_User Then EE.User:=St1 ;
        If St=_Editeur Then EE.Editeur:=St1 ;
        END ;
    3 : BEGIN
        If St=_Exe Then EE.Exe:=St1 ;
        END ;
    4 : BEGIN
        If St=_ExeExport Then EE.ExeExport:=St1 ;
        END ;
    5 : BEGIN //Répertoire
        If St=_RepExport Then EE.RepExport:=St1 ;
        END ;
    END ;
  END ;
END ;

Function FMVide(Var FM : tMasqueFic) : Boolean ;
Var i : Integer ;
BEGIN
Result:=TRUE ; For i:=0 To 9 Do If FM[i]<>'' Then BEGIN Result:=FALSE ; Exit ; End ;
END ;

Function IniOk(Var EE : tEnvImpAuto) : Integer ;
Var i : Integer ;
BEGIN
Result:=0 ;
If EE.Rep='' Then BEGIN Result:=1 ; Exit ; END ;
If Not DirectoryExists(EE.Rep) Then BEGIN Result:=2 ; Exit ; END ;
If FMVide(EE.MasqueFic[mMasqueEcr]) And FMVide(EE.MasqueFic[mMasqueAux]) Then BEGIN Result:=3 ; Exit ; END ;
If EE.Import='' Then BEGIN Result:=4 ; Exit ; END ;
If EE.Format='' Then BEGIN Result:=5 ; Exit ; END ;
If EE.User='' Then BEGIN Result:=6 ; Exit ; END ;
END ;

Function StErrIni(i : integer) : String ;
BEGIN
Result:='' ;
Case i Of
 -1 : Result:='0;Importation automatique;Fichier INI inexistant;W;O;O;O;' ;
  1 : Result:='0;Importation automatique;Fichier INI : [DIR] non renseigné;W;O;O;O;' ;
  2 : Result:='0;Importation automatique;Fichier INI : [DIR] répertoire inexistant;W;O;O;O;' ;
  3 : Result:='0;Importation automatique;Fichier INI : [NOMFICHIER] masque de fichier non renseigné ;W;O;O;O;' ;
  4 : Result:='0;Importation automatique;Fichier INI : [IMPORT] Non renseigné;W;O;O;O;' ;
  5 : Result:='0;Importation automatique;Fichier INI : [FORMAT] Non renseigné;W;O;O;O;' ;
  6 : Result:='0;Importation automatique;Fichier INI : [USER] Non renseigné;W;O;O;O;' ;
  END ;
END ;

Function LireFileIni(NomIni : String ; Var EE : tEnvImpAuto) : Integer ;
var IniFile : TIniFile ;
    StPath : String ;
    LL : TStringList ;
    LNOMSECTION : TStringList;
begin
Fillchar(EE,SizeOf(EE),#0) ;
StPath:=ExtractFilePath(Application.ExeName) ;
LNOMSECTION:=TStringList.Create ; LL:=TStringList.Create ;
NomIni:=StPath+NomIni ;
if FileExists(NomIni) then
  BEGIN
  IniFile :=TIniFile.Create(NomINI);
  IniFile.ReadSections(LNOMSECTION) ;
  If SectionPresente(_NomSectionRepertoire,LNOMSECTION) Then
    BEGIN
    IniFile.ReadSection(_NomSectionRepertoire,LL) ; LireSection(0,IniFile,LL,_NomSectionRepertoire,EE) ;
    END ;
  LL.Clear ;
  If SectionPresente(_NomSectionNomFichier,LNOMSECTION) Then
    BEGIN
    IniFile.ReadSection(_NomSectionNomFichier,LL) ; LireSection(1,IniFile,LL,_NomSectionNomFichier,EE) ;
    END ;
  LL.Clear ;
  If SectionPresente(_NomSectionParametres,LNOMSECTION) Then
    BEGIN
    IniFile.ReadSection(_NomSectionParametres,LL) ; LireSection(2,IniFile,LL,_NomSectionParametres,EE) ;
    END ;
  LL.Clear ;
  If SectionPresente(_NomSectionExe,LNOMSECTION) Then
    BEGIN
    IniFile.ReadSection(_NomSectionExe,LL) ; LireSection(3,IniFile,LL,_NomSectionExe,EE) ;
    END ;
  LL.Clear ;
  If SectionPresente(_NomSectionExport,LNOMSECTION) Then
    BEGIN
    IniFile.ReadSection(_NomSectionExport,LL) ; LireSection(4,IniFile,LL,_NomSectionExport,EE) ;
    END ;
  LL.Clear ;
  If SectionPresente(_NomSectionRepertoireExport,LNOMSECTION) Then
    BEGIN
    IniFile.ReadSection(_NomSectionRepertoireExport,LL) ; LireSection(5,IniFile,LL,_NomSectionRepertoireExport,EE) ;
    END ;
  If LNOMSECTION<>NIL Then LNOMSECTION.Free ; If LL<>NIL Then LL.Free ;
  IniFile.Free;
  Result:=IniOk(EE) ;
  END Else Result:=-1 ;
END ;

Procedure AlimNomFicToKill(ListeFichier : TStringList ; LeChemin,LeNom : String) ;
Var SearchRec : TSearchRec ;
    St,St1,St2 : String ;
BEGIN
If FindFirst(LeNom, faAnyFile, SearchRec)=0 Then
  BEGIN
  St:=LeChemin+SearchRec.Name ; ListeFichier.Add(St) ;
  While (FindNext(SearchRec) = 0) Do
    BEGIN
    St:=LeChemin+SearchRec.Name ; ListeFichier.Add(St) ;
    END ;
  END ;
FindClose(SearchRec);
END ;

Procedure RecupNomFicToKill(Var EE : tEnvImpAuto ; ListeFichier : TStringList ; Qui : String) ;
Var SearchRec : TSearchRec ;
    LeChemin,LePrefixe,LeSuffixe,LeNom,St,St1 : String ;
    Ind  : tTypeMasque ;
    i : Integer ;
BEGIN
  LeChemin:=EE.RepExport ;
  If LeChemin<>'' Then If LeChemin[Length(LeChemin)]<>'\' Then LeChemin:=LeChemin+'\' ;
  LeNom:=LeChemin+Qui+'*.*' ;
  AlimNomFicToKill(ListeFichier,LeChemin,LeNom) ;
END ;

Procedure VideRepertoire(Var EE : tEnvImpAuto ; Qui : String) ;
Var StNom : String ;
    LL : TStringList ;
    i : Integer ;
    F : TExtFile ;
BEGIN
If EE.RepExport='' Then Exit ;
//If Not InfoImp.EE.VideRep Then Exit ;
LL:=tStringList.Create ;
RecupNomFicToKill(EE,LL,Qui) ;
For i:=0 To ll.Count-1 Do
  BEGIN
  {$i-}
  AssignFile(F,LL[i]) ; Erase(F) ;
  {$i+}
  END ;
LL.Clear ; LL.Free ;
END ;


Procedure CopieAS400(St,Qui : String) ;
Var EE : tEnvImpAuto ;
    Err : Integer ;
BEGIN
Fillchar(EE,SizeOf(EE),#0) ;
Err:=LireFileIni('IMPSU.INI',EE) ;
If Err=0 Then
  BEGIN
  FichierOnDisk(St,FALSE) ;
  If EE.ExeExport<>'' Then
    BEGIN
    FileExec(EE.ExeExport,FALSE,TRUE) ;
    Delay(15000) ;
    VideRepertoire(EE,Qui) ;
    END ;
  END ;
END ;

Procedure litSUParam(Var Contact,Rva : String) ;
Var Q : TQuery ;
BEGIN
Contact:='' ; RVA:='' ;
{$IFDEF SYSTEMU}
 {$IFDEF SPEC302}
  Q:=OpenSQL('SELECT SO_RVA,SO_CONTACT FROM SOCIETE',TRUE) ;
  If Not Q.Eof Then BEGIN Rva:=Q.Fields[0].AsString ; Contact:=Q.Fields[1].AsString ; END ;
  Ferme(Q) ;
 {$ELSE}
  Q:=OpenSQL('SELECT SOC_DATA FROM PARAMSOC WHERE SOC_NOM="SO_RVA" ',TRUE) ;
  If Not Q.Eof Then Rva:=Q.Fields[0].AsString ;
  Ferme(Q) ;
  Q:=OpenSQL('SELECT SOC_DATA FROM PARAMSOC WHERE SOC_NOM="SO_CONTACT" ',TRUE) ;
  If Not Q.Eof Then Contact:=Result+Q.Fields[0].AsString ;
  Ferme(Q) ;
 {$ENDIF}
{$ENDIF}
END ;

end.
