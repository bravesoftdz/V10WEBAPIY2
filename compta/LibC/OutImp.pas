unit OutImp ;

interface

Uses WinProcs,Forms,Ent1,Sysutils,HEnt1,HmsgBox,dbTables, Hctrls ;

type tProcImport = procedure (Qui : PChar ; PDBSOC,PV,PVH : Pointer) ; stdcall;

Var LanceurExeExterne : tProcImport ;

const ModuleImpExp : longint = 0;

Procedure LanceLibraryExterne(NomExe,NomFonction,Soc,Lequel : String) ;

implementation
(*
Procedure LanceExe(NomExe,NomFonction,Soc : String ; LC : TLigneCommandeExe) ;
Var LigneCommande,St,St1,LePath : String ;
    i : Integer ;
BEGIN
St1:=Application.ExeName ; LePath:=ExtractFilePath(St1) ;
NomExe:=LePath+NomExe ;
LigneCommande:='"CCS7" "'+NomFonction+'" "'+SOC+'" "'+V_PGI.USerName+'" ' ;
For i:=1 To 10 Do If LC[i]<>'' Then LigneCommande:=LigneCommande+'"'+LC[i]+'" ' ;
St:=NomExe+' '+LigneCommande ;
FileExec(St,FALSE,TRUE) ;
END ;
*)

Procedure LanceLibraryExterne(NomExe,NomFonction,Soc,Lequel : String) ;
Var Qui,DN : PChar ;
    St,St1,LePath,NomLibrary,Serie : String ;
BEGIN
St1:=Application.ExeName ; LePath:=ExtractFilePath(St1) ;
NomLibrary:=LePath+'CCIMP.DLL' ;
if ModuleImpExp<>0 then BEGIN FreeLibrary(ModuleImpExp); ModuleImpExp:=0 ; END ;
if ModuleImpExp=0 then
   BEGIN
   ModuleImpExp := LoadLibrary(PChar(NomLibrary));
   if ModuleImpExp<>0 then LanceurExeExterne:=GetProcAddress(ModuleImpExp, 'LanceurExeExterne') ;
   END ;
if ModuleImpExp<>0 then
   BEGIN
   Serie:='CCS7' ;
   if (EstSerie(S3)) Then Serie:='CCS3' ;
   if (EstSerie(S5)) Then Serie:='CCS5' ;
   St:=Serie+';'+NomFonction+';'+Soc+';'+V_PGI.UserName+';'+Lequel+';;;' ;
   Qui:=PChar(St) ; DN:=PChar(DBSOC.DriverName) ;
   ClearLesTablettes ;
   LanceurExeExterne(Qui,DBSOC.Handle,V,VH);
   if ModuleImpExp<>0 then BEGIN FreeLibrary(ModuleImpExp); ModuleImpExp:=0 ; END ;
   END Else HShowMessage('0;Import - Export;Le fichier CCIMP.DLL n''est pas installé !;W;O;O;O','','') ;
END ;

initialization

Finalization
if ModuleImpExp<>0 then FreeLibrary(ModuleImpExp);
end.
