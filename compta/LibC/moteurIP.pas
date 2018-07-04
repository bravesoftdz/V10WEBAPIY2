unit moteurIP;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  ExtCtrls, ComCtrls, HStatus, DBTables, ImpUtil, FmtChoix, Ent1, HEnt1,
  HCtrls,DB,MajTable,TImpFic,LicUtil,ed_tools,VerCpta, StdCtrls, ImpFicU,RappType ;

Type TEnvImpAuto = Record
                   Format,Lequel,NomFic,CodeFormat : String ;
                   End ;

procedure MoteurImportAuto(Var Env : TEnvImpAuto ; MessImport : tlabel) ;

implementation

Procedure RepositionneNomFic(Var InfoImp : TInfoImport) ;
Var LePath,LeNom,Lextension : String ;
    i : Integer ;
BEGIN
If InfoImp.NomFic='AUTO' Then Exit ;
LePath:=ExtractFilePath(InfoImp.NomFic) ;
LExtension:=ExtractFileExt(InfoImp.NomFic) ;
LeNom:=ExtractFileName(InfoImp.NomFic) ;
i:=Pos('.',LeNom) ; If i>0 Then LeNom:=Copy(LeNom,1,i-1) ;
InfoImp.NomFicDoublon:=NewNomFicEtDir(InfoImp.NomFic,'Dbl','Doublons') ;
InfoImp.NomFicRejet:=NewNomFicEtDir(InfoImp.NomFic,'Err','Rejets') ;
InfoImp.NomFicRapport:=NewNomFicEtDir(InfoImp.NomFic,'Rap','Rapports') ;
CreateFicRapport(InfoImp) ;
END ;

Procedure RecupNomFic(Var InfoImp : TInfoImport ; ListeFichier : TStringList) ;
Var SearchRec : TSearchRec ;
    LeChemin,LePrefixe,LeSuffixe,LeNom : String ;
BEGIN
LeChemin:='' ; LePrefixe:='*' ; LeSuffixe:='*' ;
If InfoImp.Sc.Chemin<>'' Then LeChemin:=InfoImp.Sc.Chemin+'\' ;
If InfoImp.Sc.Prefixe<>'' Then LePrefixe:=InfoImp.Sc.Prefixe+'*' ;
If InfoImp.Sc.Suffixe<>'' Then LeSuffixe:=InfoImp.Sc.Suffixe ;
LeNom:=LeChemin+LePrefixe+'.'+LeSuffixe ;
If FindFirst(LeNom, faAnyFile, SearchRec)=0 Then
  BEGIN
  If (Pos('Dbl',SearchRec.Name)=0) And  (Pos('Err',SearchRec.Name)=0) Then ListeFichier.Add(LeChemin+SearchRec.Name) ;
  While (FindNext(SearchRec) = 0) Do
    If (Pos('Dbl',SearchRec.Name)=0) And  (Pos('Err',SearchRec.Name)=0) Then ListeFichier.Add(LeChemin+SearchRec.Name) ;
  END ;
FindClose(SearchRec);
END ;

Function FichierSuivant(ListeFichier : TStringList ; Var InfoImp : TInfoImport ; Var NumFic : Integer) : Boolean ;
BEGIN
Result:=FALSE ;
If NumFic<=ListeFichier.Count-1 Then
  BEGIN
  InfoImp.NomFic:=ListeFichier[NumFic] ;
  RepositionneNomFic(InfoImp) ;
  Inc(NumFic) ;
  END Else Exit ;
Result:=TRUE ;
END ;
(*
Procedure CreateListe(Var InfoImp :TInfoImport) ;
BEGIN
InfoImp.LGenLu:=TStringList.Create ;
InfoImp.LAuxLu:=TStringList.Create ;
InfoImp.LAnaLu:=TStringList.Create ;
InfoImp.LJalLu:=TStringList.Create ;
InfoImp.LMP:=TStringList.Create ;
InfoImp.LMR:=TStringList.Create ;
InfoImp.LRGT:=TStringList.Create ;
InfoImp.ListeCptFaux:=TList.Create ;
InfoImp.ListePieceFausse:=TList.Create ;
InfoImp.ListeEntetePieceFausse:=TStringList.Create ;
InfoImp.ListeEnteteDoublon:=TStringList.Create ;
InfoImp.ListePieceIntegre:=TStringList.Create ;
InfoImp.CRListeEnteteDoublon:=TList.Create ;
END ;
*)
Procedure AlimInfoImp(Var InfoImp :TInfoImport ; Var Env : TEnvImpAuto) ;
BEGIN
FillChar(InfoImp,SizeOf(InfoImp),#0) ;
CreateListeImp(InfoImp) ;
InfoImp.Lequel:=Env.Lequel ;
InfoImp.Format:=Env.Format ;
If InfoImp.Format='ORLI' Then BEGIN InfoImp.Format:='SN2' ; Sn2Orli:=TRUE ; END ;
InfoImp.ImportAuto:=TRUE ;
InfoImp.NomFic:=Env.NomFic ;
InfoImp.CodeFormat:=Env.CodeFormat ;
InfoImp.ForceQualif:='' ;
InfoImp.ForcePositif:=FALSE ;
InfoImp.ForceBourrage:=TRUE ;
InfoImp.CtrlDB:=FALSE ;
InfoImp.ForceNumPiece:=FALSE ;
If RecupSISCO Then InfoImp.Format:='SIS' ;
If RecupSERVANT Then InfoImp.Format:='CGE' ;
ChargeScenarioImport(InfoImp,True) ;
If RecupSISCO Then InfoImp.Format:='CGE' ;
If InfoImp.Sc.EstCharge Then
  BEGIN
  If InfoImp.SC.Doublon Then
    BEGIN
    InfoImp.CtrlDB:=TRUE ;
    If InfoImp.SC.ForcePiece Then InfoImp.ForceNumPiece:=TRUE ;
    END ;
  END ;
RepositionneNomFic(InfoImp) ;
END ;


Procedure initImporte(Var InfoImp :TInfoImport) ;
BEGIN
InfoImp.NbGenFaux:=0 ; InfoImp.NbAuxFaux:=0 ;
InfoImp.NbAnaFaux:=0 ; InfoImp.NbJalFaux:=0 ;
VideListeInfoImp(InfoImp,FALSE) ;
InfoImp.ListeEnteteDoublon.Sorted:=TRUE ;
InfoImp.ListeEnteteDoublon.Duplicates:=DupIgnore ;
END ;


Procedure initVerif(Var InfoImp :TInfoImport) ;
BEGIN
InfoImp.ListeEntetePieceFausse.Clear ;
InfoImp.ListeEntetePieceFausse.Sorted:=TRUE ;
InfoImp.ListeEntetePieceFausse.Duplicates:=DupIgnore ;
VideListe(InfoImp.ListePieceFausse) ;
InfoImp.ListePbAna.Clear ;
InfoImp.ListePbAna.Sorted:=TRUE ;
InfoImp.ListePbAna.Duplicates:=DupIgnore ;
END ;

Procedure InitIntegre(Var InfoImp :TInfoImport) ;
BEGIN
InfoImp.ListePieceIntegre.Clear ;
InfoImp.ListePieceIntegre.Sorted:=TRUE ;
InfoImp.ListePieceIntegre.Duplicates:=DupIgnore ;
END ;

Procedure DetruitFichier(Var InfoImp :TInfoImport) ;
Var Fichier : TextFile ;
    FichierOk : Boolean ;
    i : Integer ;
BEGIN
If InfoImp.Sc.DetruitFic Then
  BEGIN
  i:=0 ;
  Repeat
    Inc(i) ;
    AssignFile(Fichier,InfoImp.NomFic) ;
    {$I-} Reset (Fichier) ; {$I+}
    FichierOk:=ioresult=0 ;
    If FichierOk Then Close(Fichier) ;
  Until FichierOk Or (i>200) ;
  If FichierOk Then
    BEGIN
    {$i-}
    AssignFile(Fichier,InfoImp.NomFic) ; Erase(Fichier) ;
    {$i+}
    END ;
  END ;
END ;


procedure MoteurImportAuto(Var Env : TEnvImpAuto ; MessImport : tlabel) ;
var Soc : String ;
    InfoImp : PtTInfoImport ;
    Qui : String ;
    Auto,PlusDeFichierAImporter,OkFic : Boolean ;
    ListeFichier : TStringList ;
    NumFic : Integer ;
begin
PlusDeFichierAImporter:=TRUE ;
Auto:=FALSE ;
New(InfoImp) ; AlimInfoImp(InfoImp^,Env) ;
If InfoImp.NomFic='AUTO' Then Auto:=TRUE ;
If Auto Then
  BEGIN
  ListeFichier:=TStringList.Create ;
  RecupNomFic(InfoImp^,ListeFichier) ;
  END ;
NumFic:=0 ;
Repeat
  OkFic:=TRUE ; If Auto Then BEGIN OkFic:=FichierSuivant(ListeFichier,InfoImp^,NumFic) ; VideListeInfoImp(InfoImp^,FALSE) ; END ;
  If OkFic Then
    BEGIN
    PlusDeFichierAImporter:=FALSE ;
    If MessImport<>NIL Then
      BEGIN
      MessImport.Visible:=TRUE ;
      MessImport.Caption:='Importation en cours du fichier '+InfoImp^.NomFic ;
      END ;
    If ImporteLesEcritures(nil,InfoImp^) then
      BEGIN
      InitVerif(InfoImp^) ;
      if not VerPourImp2(InfoImp^.ListeEntetePieceFausse,InfoImp^.ListePieceFausse,TRUE,tvEcr) then
        BEGIN
        MajImpErr(InfoImp.ListeEntetePieceFausse) ;
        END ;
    //     if AucunMvtControlesOk then BEGIN HM.Execute(42,'','') ; Exit ; END ;
      InitIntegre(InfoImp^) ;
      IntegreEcr(Nil,InfoImp^) ;
      END else VH^.ImportRL:=FALSE ;
    FaitFichierDoublon(InfoImp^) ; FaitFichierRejet(InfoImp^) ; DetruitFichier(InfoImp^) ;
    END Else PlusDeFichierAImporter:=TRUE ;
  If Not Auto Then PlusDeFichierAImporter:=TRUE ;
Until PlusDeFichierAImporter ;
If Auto Then ListeFichier.Free ;
VideListeInfoImp(InfoImp^,TRUE) ;
Dispose(InfoImp) ;
SourisNormale ;
END ;


end.
