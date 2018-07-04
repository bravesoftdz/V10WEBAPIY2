unit DiagEUR;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, Hcompte, Ent1, HEnt1, Mask, Hctrls, HSysMenu, {$IFNDEF DBXPRESS}dbtables,
  ExtCtrls{$ELSE}uDbxDataSet{$ENDIF}, HStatus,
  ExtCtrls, Db, hmsgbox  ;

Procedure MessEuro1 ;

type
  TFDiagEur1 = class(TForm)
    HMTrad: THSystemMenu;
    Panel1: TPanel;
    Att2: THLabel;
    Att3: THLabel;
    Att4: THLabel;
    HLabel3: THLabel;
    HLabel4: THLabel;
    HLabel5: THLabel;
    HLabel6: THLabel;
  private
    { Déclarations privées }
  public
    { Déclarations publiques }
  end;

implementation

{$R *.DFM}

Procedure MessEuro1 ;
var TT : TRepDbl;
BEGIN
TT:=TRepDbl.Create(Application) ;
 Try
  TT.ShowModal ;
 Finally
  TT.Free ;
 End ;
END ;



end.
