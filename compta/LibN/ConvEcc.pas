unit ConvEcc ;

interface

Uses Classes,SysUtils, Ent1, HEnt1, IniFiles, DBTables, Forms, HCtrls, TImpFic ;

procedure LireFileIni(Var FEC : tFicEnCours) ;

implementation

Const
  NomSectionEEcc = 'EN COURS ENTETE' ;
  NomSectionDEcc = 'EN COURS DETAIL' ;
  _NEB_NBFIXE = 'NOMBRE DE LIGNES A FIXE' ;
  _NEB_Sens     = 'SENS SOLDE COMPTABLE' ;
  _NEB_onnaie  = 'MONNAIE' ;
  _NEB_Comptable  = 'SOLDE COMPTABLE' ;
  _NEB_SoldeEnCoursNonEchu  = 'SOLDE EN COURS TRAITE NON ECHU' ;
  _NEB_SoldeEnCoursEchuNonRegle  = 'SOLDE EN COURS TRAITE ECHU NON REGLE' ;
  _NEB_DebitEnCoursNonEchu  = 'DEBIT EN COURS TRAITE NON ECHU' ;
  _NEB_DebitEnCoursEchuNonRegle  = 'DEBIT EN COURS TRAITE ECHU NON REGLE' ;
  _NEB_CreditEnCoursNonEchu  = 'CREDIT EN COURS TRAITE NON ECHU' ;
  _NEB_CreditEnCoursEchuNonRegle  = 'CREDIT EN COURS TRAITE ECHU NON REGLE' ;
  _NEB_Dev = 'DEVISE' ;

Function FaitStSQL(FEC : tFicEnCours) : String ;
Var i : Integer ;
BEGIN
Result:='' ;
For i:=0 To FEC.LSQL.Count-1 Do
  BEGIN
  Result:=Result+FEC.LSQL[i]+', ' ;
  END ;
END ;

Function SectionPresente(St : String ; LNOMSECTION : TStrings) : Boolean ;
Var i : Integer ;
BEGIN
Result:=TRUE ;
For i:=0 To LNOMSECTION.Count-1 Do If LNOMSECTION[i]=St Then Exit ;
Result:=FALSE ;
End ;

Procedure LireUneSection(IniFile : TIniFile ; LSECTION : TStrings ; j : Integer ; Var FEC : tFicEnCours) ;
Var i : Integer ;
    St,St1,St2 : String ;
BEGIN
For i:=0 TO LSECTION.Count-1 Do
  BEGIN
  St:=LSECTION[i] ;
  St1:=IniFile.ReadString(FEC.TSI[j].NomSection,St,'');
  If St1<>'' Then
    BEGIN
    St2:=St+';'+St1 ; FEC.TSI[j].LSection.Add(St2) ;
    If Pos('_',St)>0 Then FEC.LSQL.Add(St) ;
    END ;
  END ;
END ;


procedure LireFileIni(Var FEC : tFicEnCours) ;
var IniFile : TIniFile ;
    LNOMSECTION,LSECTION : TStringList ;
    i : Integer ;
begin
LNOMSECTION:=TStringList.Create ; LSECTION:=TStringList.Create ;
FEC.LSQL.Clear ;                  
IniFile :=TIniFile.Create(FEC.NomIni);
IniFile.ReadSections(LNOMSECTION) ;
For i:=0 To MaxSectionIni Do If FEC.TSI[i].NomSection<>'' Then
  If SectionPresente(FEC.TSI[i].NomSection,LNOMSECTION) Then
    BEGIN
    FEC.TSI[i].LSection.Clear ;
    If Pos('ENTETE',FEC.TSI[i].NomSection)>0 Then FEC.TSI[i].TypeSection:=tsEntete Else FEC.TSI[i].TypeSection:=tsDetail ;
    IniFile.ReadSection(FEC.TSI[i].NomSection,LSECTION) ; LireUneSection(IniFile,LSECTION,i,FEC) ;
    END ;
If LNOMSECTION<>NIL Then LNOMSECTION.Free ;
If LSECTION<>NIL Then LSECTION.Free ;
IniFile.Free;
END ;
(*
Function lireFormatEccIni(Var EE : tEnrEntree) : Boolean ;
BEGIN
Result:=FALSE ; If Not CreateFileIni(FALSE) Then BEGIN LireFileIni(EE) ; Result:=TRUE ; END;
END ;
*)

(*
Function CreateFileEccIni(OkCreate : Boolean ; NomSection : String) : Boolean ;
var IniFile : TIniFile ;
    StPath : String ;
begin
Result:=False ;
if not FileExists('ee') then
  BEGIN
  Result:=True ;
  If OkCreate Then
    BEGIN
    IniFile :=TIniFile.Create('ee');
    // Création par défaut
    IniFile.WriteString(NomSection,_NECCNBFIXE,'0;0;') ;
    IniFile.WriteString(NomSection,_NECCMonnaie,'F;') ;
    IniFile.WriteString(NomSection,'T_AUXILIAIRE','4;8;') ;
    IniFile.WriteString(NomSection,_NECCComptable,'18;10;&[d"0"];') ;
    IniFile.WriteString(NomSection,_NECCSens,'28;1;&[DEBIT"D";CREDIT"C"];') ;
    IniFile.WriteString(NomSection,_NECCSoldeEnCoursNonEchu,'29;10;&[d"0"];') ;
    IniFile.WriteString(NomSection,_NECCSoldeEnCoursEchuNonRegle,'39;10;&[d"0"];') ;
    IniFile.WriteString(NomSection,_NECCDev,'39;3;') ;
    IniFile.Free;
    END ;
  END ;
END ;
*)

end.
