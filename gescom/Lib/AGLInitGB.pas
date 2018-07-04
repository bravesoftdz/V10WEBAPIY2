unit AGLInitGB;

interface

uses M3VM,TestGB ;

implementation

uses HCtrls,HMsgBox,M3FP, SysUtils, Fe_Main, Forms, DB, Fiche, Mul, FichList ;

procedure AGLtestGB( parms: array of variant; nb: integer ) ;
begin
TestGBV(String(parms[0])) ;
end;

//////////////////////////////////////////////////////////////////////////////
procedure initM3GB();
begin
  RegisterAglProc('TestGB', FALSE , 1, AGLTestGB) ;
end;

Initialization
initM3GB();
finalization
end.
