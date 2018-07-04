unit Usisco;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  hmsgbox, HSysMenu, ExtCtrls, Ent1, HEnt1, HCtrls, DBTables, HStatus,
  HTB97, StdCtrls, ImpFicU, SISCO, TImpFic, Buttons, Menus, HPanel, UiUtil,
  HFLabel ;

procedure Un2DeuxSISCO(Var InfoImp :TInfoImport) ;
Procedure ChangeNomFicPourBigVol(Var InfoImp :TInfoImport ; St : String) ;

implementation

Procedure ChangeNomFicPourBigVol(Var InfoImp :TInfoImport ; St : String) ;
BEGIN
InfoImp.NomFicCptCGE:=NewNomFic(InfoImp.NomFicCpt,'CGE'+St) ;
InfoImp.NomFicMvtCGE:=NewNomFic(InfoImp.NomFicMvt,'CGE'+St) ;
END ;

Procedure PrepareTRS(Var InfoImp :TInfoImport) ;
BEGIN
InfoImp.NomFicOrigine:=InfoImp.NomFic ;
InfoImp.NomFicCpt:=NewNomFicEtDir(InfoImp.NomFic,'Cpt','SISCO') ;
InfoImp.NomFicMvt:=NewNomFicEtDir(InfoImp.NomFic,'Mvt','SISCO') ;
InfoImp.NomFicCptCGE:=NewNomFic(InfoImp.NomFicCpt,'CGE') ;
InfoImp.NomFicMvtCGE:=NewNomFic(InfoImp.NomFicMvt,'CGE') ;
END ;

Function NatureEnr(St : String) : Integer ;
Var Code1,Code2 : String ;
(*
Cpt 03 05 06 07 08 09 10 C 30 50 P L D I f s R S
Mvt 00 01 02 03 04 11 M J F E 12 p t r A B H v w G
*)

BEGIN
If Not VH^.RecupCEGID Then
  BEGIN
  Result:=2 ;
  Code1:=Copy(St,1,1) ; Code2:=Copy(St,1,2) ;
  If (Code1='C') Or (Code1='P') Or (Code1='L') Or (Code1='D') Or (Code1='I') Or
     (Code1='f') Or (Code1='s') Or (Code1='R') Or (Code1='S') Or
     (Code2='05') Or (Code2='06') Or (Code2='07') Or (Code2='08') Or (Code2='09') Or
     (Code2='10') Or (Code2='30') Or (Code2='50') Then Result:=0 ;
  If (Code1='M') Or (Code1='J') Or (Code1='F') Or (Code1='E') Or (Code1='p') Or
     (Code1='t') Or (Code1='r') Or (Code1='A') Or (Code1='B') Or
     (Code1='H') Or (Code1='v') Or (Code1='w') Or (Code1='G') Or
     (Code2='11') Or (Code2='12') Then Result:=1 ;
  END Else
  BEGIN
  Result:=2 ;
  Code1:=Copy(St,1,1) ; Code2:=Copy(St,1,2) ;
  If (Code1='C') Or (Code1='P') Or (Code1='L') Or (Code1='D') Or (Code1='I') Or
     (Code1='f') Or (Code1='s') Or (Code1='R') Or (Code1='S') Or
     (Code2='05') Or (Code2='06') Or (Code2='07') Or (Code2='08') Or (Code2='09') Or
     (Code2='10') Or (Code2='30') Or (Code2='50') Or (Code1='P') Or (Code1='L') Then Result:=0 ;
  If (Code1='M') Or (Code1='J') Or (Code1='F') Or (Code1='E') Or (Code1='p') Or
     (Code1='t') Or (Code1='r') Or (Code1='A') Or (Code1='B') Or
     (Code1='H') Or (Code1='v') Or (Code1='w') Or (Code1='G') Or
     (Code2='00') Or (Code2='01') Or (Code2='02') Or (Code2='04') Or
     (Code2='11') Or (Code2='12') Then Result:=1 ;
  END ;
END ;

Procedure PrepareFichierPhase1(Var InfoImp :TInfoImport) ;
Var Fichier,NewFichier1,NewFichier2 : TextFile ;
    St : String ;
    Pb : Boolean ;
    What : Integer ;
BEGIN
Pb:=FALSE ;
InitMove(1000,TraduireMemoire('Chargement du fichier en cours...')) ;
AssignFile(Fichier,InfoImp.NomFicOrigine) ; Reset(Fichier) ;
AssignFile(NewFichier1,InfoImp.NomFicCpt) ; Rewrite(NewFichier1) ;
AssignFile(NewFichier2,InfoImp.NomFicMvt) ; Rewrite(NewFichier2) ;
//TRS.NbLCpt:=0 ; TRS.NbLMvt:=0 ;
While (Not EOF(Fichier)) And (Not Pb)  do
  BEGIN
  MoveCur(FALSE) ;
  ReadLn(Fichier,St) ;
  What:=NatureEnr(St) ;
  Case What Of
    0 : BEGIN Writeln(NewFichier1,St) ; (*Inc(TRS.NbLCpt) ;*) END ; // Cpt
    1 : BEGIN Writeln(NewFichier2,St) ; (*Inc(TRS.NbLMvt) ;*) END ; // Mvt
    2 : BEGIN Writeln(NewFichier1,St) ; Writeln(NewFichier2,St) ; (*Inc(TRS.NbLCpt) ; Inc(TRS.NbLMvt) ;*) END ;// Les 2
    END ;
  END ;
FiniMove ; Flush(NewFichier1) ; Flush(NewFichier2) ;
CloseFile(Fichier) ; CloseFile(NewFichier1) ; CloseFile(NewFichier2) ;
FichierOnDisk(InfoImp.NomFicCpt,TRUE) ; FichierOnDisk(InfoImp.NomFicMvt,TRUE) ;
FiniMove ;
END ;



procedure Un2DeuxSISCO(Var InfoImp :TInfoImport) ;
begin
PrepareTRS(InfoImp) ;
PrepareFichierPhase1(InfoImp) ;
end;


end.
