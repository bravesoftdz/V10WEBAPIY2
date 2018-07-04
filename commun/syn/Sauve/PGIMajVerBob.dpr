program PGIMajVerBob;

uses
  Forms,
  Hent1,
  AfInitSoc in '..\..\Gescom\LibA\AfInitSoc.pas',
  AglInitCommun in '..\LIB\AglInitCommun.pas',
  AssistInitSoc in '..\Lib\AssistInitSoc.pas' {FAssistInitSoc},
  Confidentialite_TOF in '..\LIB\Confidentialite_TOF.pas',
  CORRESP_TOM in '..\..\compta\LibN\CORRESP_TOM.PAS',
  DicoAf in '..\..\Gescom\LibA\DicoAf.pas',
  Dispo in '..\..\GESCOM\LIB\Dispo.PAS',
  DiversCommunTOT in '..\Lib\DiversCommunTOT.pas',
  ECRPIECE_TOF in '..\LIB\ECRPIECE_TOF.PAS',
  EntGC in '..\LIB\EntGC.pas',
  EntPGI in '..\LIB\EntPGI.pas',
  EuroPGI in '..\Lib\EuroPGI.pas',
  FichComm in '..\Lib\FichComm.pas',
  Fiche,
  FichGrid,
  FichList,
  galFileTools in '..\..\CegidPGI\Lib\galFileTools.pas',
  galSystem in '..\..\CegidPGI\Lib\galSystem.pas',
  GRS1,
  HManifest in '..\LIB\HManifest.pas',
  ImportObjetLot_TOF in '..\LIB\ImportObjetLot_TOF.pas',
  LettAuto in '..\Lib\Lettauto.pas',
  MajHalley in '..\LIB\MajHalley.pas',
  MajHalleyProc in '..\LIB\MajHalleyProc.pas',
  MAJHALLEYUTIL in '..\LIB\MAJHALLEYUTIL.PAS',
  MajSocParLot_TOF in '..\LIB\MajSocParLot_TOF.PAS',
  MDispAdm in '..\LIB\MDispAdm.pas' {gctie: TDataModule},
  menuolg,
  Mul,
  SplashG,
  StkMoulinette in '..\..\GESCOM\LIB\StkMoulinette.pas',
  StkNature in '..\..\Gescom\Lib\StkNature.PAS',
  utilConfid in '..\..\Gescom\Lib\utilConfid.PAS',
  Tablette,
  TiersUtil in '..\LIB\TiersUtil.pas',
  UserGrp_tom in '..\LIB\UserGrp_tom.pas',
  UtilGA in '..\..\GESCOM\LIBA\UtilGA.pas',
  utilPGI in '..\LIB\utilPGI.pas',
  UtilRessource in '..\..\Gescom\LibA\UtilRessource.pas',
  UtilSoc in '..\LIB\UtilSoc.pas',
  UtofExportConfidentialite in '..\..\Prospect\lib\UtofExportConfidentialite.pas',
  UTOFYYIMPORTOBJET in '..\LIB\UTOFYYIMPORTOBJET.PAS',
  UTOMUTILISAT in '..\LIB\UTOMUTILISAT.PAS',
  wcommuns in '..\LIB\wcommuns.pas',
  wJetons in '..\LIB\wJetons.pas',
  wrapport in '..\LIB\wrapport.pas',
  WTOM in '..\LIB\WTOM.PAS',
  yTarifs_TOM in '..\..\GESCOM\LIB\yTarifs_TOM.pas',
  yTarifsBascule in '..\..\GESCOM\LIB\yTarifsBascule.pas',
  yTarifsCommun in '..\..\GESCOM\LIB\yTarifsCommun.pas',
  yTarifsFourchette_TOM in '..\..\GESCOM\LIB\yTarifsFourchette_TOM.pas',
  UObjFiltres in '..\..\treso\Lib\UObjFiltres.pas',
  UProcGen in '..\..\treso\Lib\UProcGen.pas',
  uRecupSQLModele in '..\Lib\uRecupSQLModele.pas';

{$R *.RES}

begin
  Application.Initialize;
  InitAgl;
  Application.Title := 'PGIMajVer';
  Application.CreateForm(TFMenuG, FMenuG);
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
