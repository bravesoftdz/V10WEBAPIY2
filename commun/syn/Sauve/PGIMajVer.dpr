program PGIMajVer;

uses
  MajStruc in 'C:\Users\sidos\Projets\PGIMAJVER\20071024\MajStruc.pas',
  Forms,
  Hent1,
  AfInitSoc in '..\..\..\Gescom\LibA\AfInitSoc.pas',
  AglInitCommun in '..\Lib\AGLINITCOMMUN.PAS',
  ASSISTINITSOC in '..\..\..\commun\LIB\ASSISTINITSOC.PAS' {FAssistInitSoc},
  CORRESP_TOM in '..\..\..\COMPTA\LIBN\CORRESP_TOM.PAS',
  DicoAf in '..\..\..\Gescom\LibA\DicoAf.pas',
  Dispo in '..\..\..\GESCOM\LIB\Dispo.PAS',
  DiversCommunTOT in '..\..\..\commun\LIB\DiversCommunTOT.pas',
  ECRPIECE_TOF in '..\..\..\commun\LIB\ECRPIECE_TOF.PAS',
  EntGC in '..\Lib\ENTGC.PAS',
  EntPGI in '..\Lib\ENTPGI.PAS',
  EuroPGI in '..\..\..\commun\LIB\EuroPGI.pas',
  FichComm in '..\..\..\commun\LIB\FichComm.pas',
  Fiche,
  FichGrid,
  FichList,
  galFileTools in '..\..\..\CegidPGI\Lib\galFileTools.pas',
  galSystem in '..\..\..\CegidPGI\Lib\galSystem.pas',
  GRS1,
  HManifest in '..\..\..\commun\LIB\HManifest.pas',
  ImportObjetLot_TOF in '..\..\..\commun\LIB\ImportObjetLot_TOF.pas',
  LettAuto in '..\..\..\commun\Lib\Lettauto.pas',
  MajSocParLot_TOF in '..\..\..\commun\LIB\MajSocParLot_TOF.PAS',
  MDispAdm in '..\LIB\MDispAdm.pas' {gctie: TDataModule},
  menuolg {FMenuG},
  Mul,
  SplashG,
  StkMoulinette in '..\..\..\GESCOM\LIB\StkMoulinette.pas',
  StkNature in '..\..\..\GESCOM\LIB\StkNature.PAS',
  Tablette,
  UserGrp_tom in '..\Lib\USERGRP_TOM.PAS',
  UtilGA in '..\..\..\GESCOM\LIBA\UtilGA.pas',
  utilPGI in '..\Lib\UTILPGI.PAS',
  UtilRessource in '..\..\..\GESCOM\LIBA\UtilRessource.pas',
  UtilSoc in '..\Lib\UTILSOC.PAS',
  UtofExportConfidentialite in '..\..\..\PROSPECT\LIB\UtofExportConfidentialite.pas',
  UTOFYYIMPORTOBJET in '..\..\..\commun\LIB\UTOFYYIMPORTOBJET.PAS',
  UTOMUTILISAT in '..\Lib\UTOMUTILISAT.PAS',
  wCommuns in '..\Lib\WCOMMUNS.PAS',
  wJetons in '..\..\..\commun\LIB\wJetons.pas',
  wrapport in '..\..\..\commun\LIB\wrapport.pas',
  WTOM in '..\..\..\GESCOM\LIBW\WTOM.PAS',
  uTOMComm in '..\Lib\UTOMCOMM.PAS',
  yTarifs_TOM in '..\..\..\GESCOM\LIB\yTarifs_TOM.pas',
  yTarifsBascule in '..\..\..\GESCOM\LIB\yTarifsBascule.pas',
  yTarifsCommun in '..\..\..\GESCOM\LIB\yTarifsCommun.pas',
  yTarifsFourchette_TOM in '..\..\..\GESCOM\LIB\yTarifsFourchette_TOM.Pas',
  UObjFiltres in '..\..\..\treso\Lib\UObjFiltres.pas',
  UProcGen in '..\..\..\treso\Lib\UProcGen.pas',
  uRecupSQLModele in '..\..\..\commun\LIB\uRecupSQLModele.pas',
  UtilConfid in '..\..\..\GESCOM\LIB\UtilConfid.pas',
  BUNDLEDETAIL_TOF in '..\..\..\commun\LIB\BUNDLEDETAIL_TOF.PAS',
  Constantes in '..\..\..\treso\Lib\Constantes.pas',
  EBizUtil in '..\..\..\GESCOM\LIBec\EBizUtil.pas',
  BackupRestore_TOF in '..\..\..\commun\LIB\BackupRestore_TOF.pas',
  gcLienPiece in '..\..\..\GESCOM\LIB\gcLienPiece.pas' {dgcLienPiece},
  uTablesGed in '..\..\..\commun\LIB\uTablesGed.pas',
  ListUsersAutorises in '..\..\..\commun\Lib\ListUsersAutorises.pas',
  MajEnRafale in '..\LIB\MajEnRafale.PAS',
  AccesPortail_TOF in '..\..\..\commun\LIB\AccesPortail_TOF.pas',
  BasculeArticleLie in '..\..\..\Gescom\Lib\BasculeArticleLie.pas',
  YYBUNDLE_TOF in '..\Lib\YYBUNDLE_TOF.PAS',
  UTomEtabliss in '..\Lib\UTOMETABLISS.PAS',
  UTOMBanque in '..\Lib\UTOMBANQUE.PAS',
  TomAgence in '..\..\..\treso\Lib\TomAgence.pas',
  TomCalendrier in '..\..\..\treso\Lib\TomCalendrier.pas',
  Commun in '..\..\..\treso\Lib\Commun.pas',
  ULibIdentBancaire in '..\Lib\ULIBIDENTBANCAIRE.PAS',
  TomIdentificationBancaire in '..\..\..\commun\Lib\TomIdentificationBancaire.pas',
  uLibRevision in '..\..\..\compta\LibN\uLibRevision.pas',
  MajHalleyUtil in '..\..\..\COMMUN\LIB\MajHalleyUtil.pas',
  MajHalley in '..\LIB\MajHalley.pas',
  MajHalleyProc in '..\..\..\commun\Lib\MajHalleyProc.pas',
  Confidentialite_TOF in '..\Lib\CONFIDENTIALITE_TOF.PAS',
  YPlanning in '..\Lib\YPLANNING.PAS',
  YRESSOURCE in '..\Lib\YRESSOURCE.PAS',
  YYYPLANNING_MUL_TOF in '..\Lib\YYYPLANNING_MUL_TOF.PAS',
  YYYRESSOURCE_MUL_TOF in '..\..\..\commun\Lib\YYYRESSOURCE_MUL_TOF.PAS',
  AJOUTSOCRGT_TOF in '..\..\..\COMMUN\LIB\AJOUTSOCRGT_TOF.PAS';

{$R *.RES}

begin
  Application.Initialize;
  InitAgl;
  Application.Title := 'Administration Sociétés';
  Application.CreateForm(TFMenuG, FMenuG);
  Application.CreateForm(Tgctie, gctie);
  InitApplication ;
  SplashScreen:=NIL  ;
  if paramcount=0 then
     BEGIN
     SplashScreen := TSplashScreen.Create(Nil) ;
     SplashScreen.Show ;
     SplashScreen.Update ;
     END ;
  Application.Run;
end.
