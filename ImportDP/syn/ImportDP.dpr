program ImportDP;

uses
  SysUtils,
  Forms,
  Hent1,
  Assist in '..\..\..\..\..\..\..\Agl7\mac7\assist.pas' {FAssist},
  Mul in '..\..\..\..\..\..\..\Agl7\mac7\Mul.pas' {FMul},
  FichList in '..\..\..\..\..\..\..\Agl7\mac7\FichList.pas' {FFicheListe},
  Fiche in '..\..\..\..\..\..\..\Agl7\mac7\Fiche.pas' {FFiche},
  FichGrid in '..\..\..\..\..\..\..\Agl7\mac7\FichGrid.pas' {FFicheGrid},
  GRS1 in '..\..\..\..\..\..\..\Agl7\mac7\GRS1.pas' {FGRS1},
  QRS1 in '..\..\..\..\..\..\..\Agl7\mac7\QRS1.pas' {FQRS1},
  Vierge in '..\..\..\..\..\..\..\Agl7\mac7\Vierge.pas' {FVierge},
  Tablette {FTablette},
  Prefs {FPrefs},
  ParamSoc {FParamSoc},
  MenuOLG {FMenuG},
  SplashG {SplashScreen},
  edtQR {EditQR},
  UImportDP in '..\LIB\UImportDP.pas',
  UtilTrans in '..\..\compta\LibC\UtilTrans.pas',
  FicheLanceImport in '..\LIB\FicheLanceImport.pas' {FImportDp},
  MDispImportDp in '..\LIB\MDispImportDp.pas' {FMenuDisp: TDataModule},
  galFileTools in '..\..\CegidPgi\Lib\galFileTools.pas',
  UTobTrans in '..\..\commun\Lib\UTobTrans.pas',
  galSystem in '..\..\CegidPgi\Lib\galSystem.pas',
  UtilGed in '..\..\DP\Lib\UtilGed.pas',
  UtilDossier in '..\..\DP\Lib\UtilDossier.pas',
  YDocuments_Tom in '..\..\commun\Lib\YDocuments_Tom.pas',
  YFiles_Tom in '..\..\commun\Lib\YFiles_Tom.pas',
  DpTableauBordLibrairie in '..\..\DP\Lib\DpTableauBordLibrairie.pas',
  uYFILESTD2 in '..\lib\uYFILESTD2.pas';

{$R *.RES}


begin
 //POUR VOIR LES FUITES MEMOIRES
{$ifdef MEMCHECK}
   MemCheckLogFileName:=ChangeFileExt(Application.exename,'.log');
   MemChk;
// utiliser AllocMem(x); => permet au niveau de chaque création d'objets
// d'avoir une trace dans monexe.log => seul moyen de repérer l'objet qui leak
// x étant un entier al mano !
{$endif}

  Application.Initialize;
  InitAgl() ;
  Application.Title := 'Utilitaire Import Dossier Permanent';
  Application.CreateForm(TFMenuG, FMenuG);
  Application.CreateForm(TFMenuDisp, FMenuDisp);
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
