unit CorresMP;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Tablette, HSysMenu, hmsgbox, Db, {$IFNDEF DBXPRESS}dbtables{$ELSE}uDbxDataSet{$ENDIF}, StdCtrls, DBCtrls, Buttons,
  ExtCtrls, Grids, DBGrids, HDB, Ent1, HEnt1, Hqry, HTB97, HCtrls ;

procedure CorrespMP(Enc : Boolean) ;

type
  TFCorrespMP = class(TFTablette)
    Msg: THMsgBox;
    procedure FormShow(Sender: TObject);
  private
    { Déclarations privées }
  public
    { Déclarations publiques }
  end;

implementation

{$R *.DFM}

procedure CorrespMP(Enc : Boolean) ;
var X : TFCorrespMP ;
begin
if Blocage(['nrCloture'],False,'nrAucun') then Exit ;
X:=TFCorrespMP.Create(Application) ;
  try
  if Enc then X.FQuoi:='TTCORRESPMPENC' else X.FQuoi:='TTCORRESPMPDEC' ;
  X.ShowModal ;
  finally
  X.Free ;
  end ;
Screen.Cursor:=SyncrDefault ;
end ;

procedure TFCorrespMP.FormShow(Sender: TObject);
begin
inherited;
  begin
  FListe.Columns[0].Width:=FListe.Columns[0].Width+FListe.Columns[2].Width ;
  FListe.Columns[2].width:=0 ;
  FListe.Columns[0].Title.Caption:=Msg.Mess[0] ;
  FListe.Columns[1].Title.Caption:=Msg.Mess[1] ;
  if FQuoi='TTCORRESPMPENC' then Caption:=Msgbox.Mess[2]
                            else Caption:=Msgbox.Mess[3] ;
  UpdateCaption(Self) ;
  end ;
end;

end.
