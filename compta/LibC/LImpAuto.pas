unit LImpAuto ;

interface

Uses HEnt1,MajTable,Ent1,MajStruc,HCtrls,SysUtils,LicUtil,Forms,
     General, Tiers, Section, HDebug,{$IFNDEF DBXPRESS}dbtables{$ELSE}uDbxDataSet{$ENDIF},Dialogs, LanceIE,HStatus  ;

Procedure LanceurExeExterne(Qui2 : PChar ; PDBSOC,PV,PVH : Pointer ); export ; StdCall ;

implementation

Procedure LanceurExeExterne(Qui2 : PChar ; PDBSOC,PV,PVH : Pointer) ;
Var Auto,OkConnection,ConnectionExterne : Boolean ;
    Qui1,Qui,Quoi,Soc,User,Lequel,Format,NomFic : string ;
    OldV,OldVH : Pointer ;
    OldVStatus : THStatusBar ;
BEGIN
if LaVariable(PV^).LaSerie=S7 then Application.HelpFile:=ExtractFilePath(Application.ExeName)+'CCS7.HLP' else
 if LaVariable(PV^).LaSerie=S5 then Application.HelpFile:=ExtractFilePath(Application.ExeName)+'CCS5.HLP' ;
Auto:=FALSE ; OkConnection:=FALSE ; ConnectionExterne:=FALSE ;
Qui1:=StrPas(Qui2) ;
Qui:=ReadtokenSt(Qui1) ; Quoi:=ReadtokenSt(Qui1) ; Soc:=ReadtokenSt(Qui1) ;
User:=ReadtokenSt(Qui1) ;Lequel:=ReadtokenSt(Qui1) ;Format:=ReadtokenSt(Qui1) ;NomFic:=ReadtokenSt(Qui1) ;
If User<>'' Then Auto:=TRUE ;
If Not Auto Then {LanceMenuManuel} Else
  BEGIN
  OkConnection:=(PDBSoc<>NIL) And (PV<>NIL) And (PVH<>NIL) ;
  If OkConnection Then
    BEGIN
    OldV:=V ; OldVH:=VH ; V:=PV ; VH:=PVH ; OldVStatus:=VStatus ;
    DBSOC:=TDataBase.Create(Application) ;
    DBSoc.Name:='ZEZE' ; DBSOc.DatabaseName:='SOC' ; DBSoc.DriverName:=DriverToBDESt(V_PGI.Driver) ;  //Astuce HTable
    DBSoc.Handle:=PDBSoc ;
    END Else
    BEGIN
    V_PGI.UserName:=Trim(User) ;
    V_PGI.PassWord:=CryptageSt(DayPass(Date)) ;
    OkConnection:=ConnecteHalley(Soc,FALSE,@ChargeMagHalley,NIL,NIL) ;
    If OkConnection Then ConnectionExterne:=TRUE ;
    END;
  if OkConnection then
    BEGIN
    If Not Assigned(ProcZoomGene)    Then ProcZoomGene    :=FicheGene    ;
    If Not Assigned(ProcZoomTiers)   Then ProcZoomTiers   :=FicheTiers   ;
    If Not Assigned(ProcZoomSection) Then ProcZoomSection :=FicheSection ;
    If (Qui='CCS7') Or (Qui='CCS5') Then
      BEGIN
      If Quoi='IMPORT' Then
        BEGIN
        If Lequel<>'' Then LanceImportExt(Lequel) Else LanceImportExt('FEC') ;
        END Else If Quoi='EXPORT' Then
        BEGIN
        If Lequel<>'' Then LanceExportExt(Lequel) Else LanceExportExt('FEC') ;
        END ;
      END ;
    If ConnectionExterne Then DeconnecteHalley else
      BEGIN
      ClearLesTablettes ;
      V:=OldV ; VH:=OldVH  ; DBSoc.Handle:=Nil ; DBSoc.Free ; VStatus:=OldVStatus ;
      END ;
    END ;
  END ;
END ;

initialization
{Titres et série du programme}
Apalatys:= 'CEGID' ;
NomHalley:= 'Comptabilité S7' ;
TitreHalley := 'CEGID Comptabilité S7' ;
Copyright   := '© Copyright CEGID' ;
{Paramètres}
(*
ChargeXuelib ;
V_PGI.VersionReseau:=True ;
//V_PGI.SAV:=False ;
V_PGI.OutLook:=False ;
V_PGI.NumVersion:='3.00' ;
V_PGI.NumBuild:='095' ;
V_PGI.DateVersion:=EncodeDate(1998,12,21) ;
V_PGI.Synap:=FALSE ;
V_PGI.VersionDEV:=True ;
V_PGI.ImpMatrix := True ;
VH^.GereSousPlan:=FALSE ;
Application.HelpFile := ExtractFilePath(Application.ExeName) + 'Halley.hlp' ;
VH^.LaSerie:=CCS7 ;
V_PGI.MultiUserLogin:=TRUE ;
V_PGI.VersionDemo:=FALSE ;
V_PGI.MultiUserLogin:=True ;
(*
VH^.SerProdCompta:='0701030' ;
VH^.SerProdBudget:='0702030' ;
*)

end.
