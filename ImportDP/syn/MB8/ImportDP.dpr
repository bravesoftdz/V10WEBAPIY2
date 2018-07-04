program ImportDP;

uses
  SysUtils,
  Forms,
  Hent1,
  Assist in 'C:\Agl7\mac7\assist.pas' {FAssist},
  Mul in 'C:\Agl7\mac7\Mul.pas' {FMul},
  FichList in 'C:\Agl7\mac7\FichList.pas' {FFicheListe},
  Fiche in 'C:\Agl7\mac7\Fiche.pas' {FFiche},
  FichGrid in 'C:\Agl7\mac7\FichGrid.pas' {FFicheGrid},
  GRS1 in 'C:\Agl7\mac7\GRS1.pas' {FGRS1},
  QRS1 in 'C:\Agl7\mac7\QRS1.pas' {FQRS1},
  Tablette in 'C:\Agl7\mac7\Tablette.pas' {FTablette},
  Prefs in 'C:\Agl7\mac7\Prefs.pas' {FPrefs},
  ParamSoc in 'C:\Agl7\mac7\ParamSoc.pas' {FParamSoc},
  MenuOLG in 'C:\Agl7\mac7\MenuOLG.pas' {FMenuG},
  SplashG in 'C:\Agl7\mac7\SplashG.pas' {SplashScreen},
  edtQR in 'C:\Agl7\mac7\Edtqr.pas' {EditQR},
  UImportDP in '..\..\LIB\UImportDP.pas',
  UtilTrans in '..\..\..\COMPTA\LIBC\UTILTRANS.PAS',
  FicheLanceImport in '..\..\LIB\FICHELANCEIMPORT.PAS' {FImportDp},
  MDispImportDp in '..\..\LIB\MDispImportDp.pas' {FMenuDisp: TDataModule},
  Vierge in 'C:\Agl7\mac7\Vierge.pas' {FVierge},
  galSystem in '..\..\..\..\CegidPgi\Lib\galSystem.pas',
  galFileTools in '..\..\..\..\CegidPgi\Lib\galFileTools.pas',
  ExceptInfo in '..\..\..\..\COMMUN\LIB\ExceptInfo.pas',
  UTobTrans in '..\..\..\..\CegidPgi\Lib\UTobTrans.pas';

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
