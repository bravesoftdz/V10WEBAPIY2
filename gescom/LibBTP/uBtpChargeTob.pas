{***********UNITE*************************************************
Auteur  ...... : Xavier PERSOUYRE
Créé le ...... : 07/06/2006
Modifié le ... :   /  /
Description .. : Unité de gestion des différents Affaires
Suite ........ : Exemple du Design Pattern : Singleton
Mots clefs ... :
*****************************************************************}
unit uBtpChargeTob;

interface

uses Windows,
  classes,
  sysutils,
  HCtrls,
  utob;

type
  MaTob = class (Tob)
  public
    procedure ReLoad() ;
    constructor Create () ; reintroduce ; overload ;
  end ;

function BTP_ChargeTob(StArguments : String) : MaTOB;

implementation

var
  FDatas  : MaTOB;

  StSQl   : String;

{***********A.G.L.***********************************************
Auteur  ...... : Xavier PERSOUYRE
Créé le ...... : 07/06/2006
Modifié le ... :   /  /
Description .. : Chargement des données
Mots clefs ... :
*****************************************************************}
procedure Load () ;
begin

  FDatas := MaTOB.Create () ;

  FDatas.LoadDetailFromSQL (StSql) ;

end ;

{***********A.G.L.***********************************************
Auteur  ...... : Xavier PERSOUYRE
Créé le ...... : 07/06/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
function BTP_ChargeTob(StArguments : String) : MaTOB;
begin

  If StArguments = '' then
     Begin
     Result := nil;
     exit;
     end;

  if StSql <> StArguments then
     Begin
     FreeAndNil(FDatas);
     StSql := StArguments;
     end;

  if not assigned(FDatas) then Load() ;

  Result := FDatas;

end;

{ MaTob }

constructor MaTob.Create;
begin
  inherited Create ('', Nil, -1) ;
end;

procedure MaTob.ReLoad;
begin
  if assigned(FDatas) then
    FDatas.ClearDetail ;
  Load () ;
end;

initialization
  FDatas := nil;

finalization
  if assigned(FDatas) then
    FreeAndNil(FDatas);

end.

