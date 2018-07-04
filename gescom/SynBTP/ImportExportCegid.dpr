program ImportExportCegid;
uses
  SysUtils,
  Forms,
  Hent1,
  Assist {FAssist},
  edtQR {EditQR},
  Fiche {FFiche},
  FichGrid {FFicheGrid},
  FichList {FFicheListe},
  GRS1 {FGRS1},
  MenuOLG {FMenuG},
  Mul {FMul},
  QRS1 {FQRS1},
  SplashG {SplashScreen},
  Tablette {FTablette},
  Vierge {FVierge},
  MenuDisp in '..\ImpExpCegid\lib\MenuDisp.pas' {FMenuDisp: TDataModule},
  UApplication in '..\LibBTP\UApplication.pas',
  UFicheImportExport in '..\ImpExpCegid\lib\UFicheImportExport.pas' {FexpImpCegid},
  CommonTools in '..\..\commun\Lib\CommonTools.pas',
  UConnectWSConst in '..\..\CONNECTWS\UConnectWSConst.pas';

//

// FIN NEW

{$R *.RES}

begin
{$ifdef MEMCHECK}
  MemCheckLogFileName:=ChangeFileExt(Application.exename,'.log');
  MemChk;
{$endif}
  GetInfoMajApplication;
  Application.Initialize;
  InitAgl;
  Application.Title := 'IMPORT EXPORT DONNEES CEGID';
  Application.CreateForm(TFMenuG, FMenuG);
  Application.CreateForm(TFMenuDisp, FMenuDisp);
  initApplication ;
  SplashScreen:=Nil ;
  if ParamCount=0 then
     BEGIN
     SplashScreen:=TSplashScreen.Create(Application) ;
     SplashScreen.Show ;
     SplashScreen.Update ;
     END ;
  Application.Run ;
end.
