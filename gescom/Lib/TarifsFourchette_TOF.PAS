{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 14/06/2002
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : TarifsFourchette ()
Mots clefs ... : TOF;TarifsFourchette
*****************************************************************}
Unit TarifsFourchette_TOF ;

Interface

Uses
  StdCtrls,
  Controls,
  Classes,
  {$IFNDEF EAGLCLIENT}
    db,
    dbtables,
  {$ENDIF}
  forms,
  sysutils,
  ComCtrls,
  HCtrls,
  HEnt1,
  HMsgBox,
  UTOF,
  wTOF,
  SaisieList,
  uTableFiltre,
  uTob;

Type
  TOF_TarifsFourchette = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnClose                  ; override ;
    private
    TF               : TTAbleFiltre;
    iIdentifiantYTA  : integer;
  end ;

Implementation

uses
  Wcommuns;

{--------------------------------------------------------------------------------
---------------------------------------------------------------------------------}
procedure TOF_TarifsFourchette.OnNew ;
begin
  Inherited ;
end ;

{--------------------------------------------------------------------------------
---------------------------------------------------------------------------------}
procedure TOF_TarifsFourchette.OnDelete ;
begin
  Inherited ;
end ;

{--------------------------------------------------------------------------------
---------------------------------------------------------------------------------}
procedure TOF_TarifsFourchette.OnUpdate ;
begin
  Inherited ;
end ;

{--------------------------------------------------------------------------------
---------------------------------------------------------------------------------}
procedure TOF_TarifsFourchette.OnLoad ;
var
  TobYTA : TOB;
  Q      : TQuery;
begin
  Inherited ;
  Q := OpenSQL('SELECT * FROM YTARIFSAQUI WHERE YTA_IDENTIFIANT='+intToStr(iIdentifiantYTA), true );
  TobYTA := TOB.create('YTARIFSAQUI',nil,-1);
  Try
    TobYTA.selectDb('',Q,true);
    TobYTA.PutEcran( Ecran, TWinControl(GetControl('PanTete')) );
  finally
    TobYTA.free;
    Ferme(Q);
  end;
end ;

{--------------------------------------------------------------------------------
---------------------------------------------------------------------------------}
procedure TOF_TarifsFourchette.OnArgument (S : String ) ;
var
  sParam, sWhere, sFonctionnalite : String;
begin
  Inherited ;
  iIdentifiantYTA := ValeurI(GetArgumentValue(S, 'YTF_IDENTIFIANTYTA'));
  if (Ecran<>nil) then
    TF := TFSaisieList(Ecran).LeFiltre;

  sPAram := S;
  ReadToKenSt(sParam);
  sWhere := 'WHERE ';
  sWhere := sWhere+ReadTokenSt(sPAram)+' AND ';
  sWhere := sWhere+ReadTokenSt(sPAram);
  TF.WhereTable := sWhere ;

  //Fixer la liste par rapport à la fonctionnalité
  sFonctionnalite := GetArgumentValue(S, 'YTA_FONCTIONNALITE');
  TFSaisieList(ecran).Dbliste               := 'YTARIFSTRANCHE'+sFonctionnalite;
  TFsaisieList(ecran).Lefiltre.LaGridListe  := 'YTARIFSTRANCHE'+sFonctionnalite;

end ;

{--------------------------------------------------------------------------------
---------------------------------------------------------------------------------}
procedure TOF_TarifsFourchette.OnClose ;
begin
  Inherited ;
end ;

{--------------------------------------------------------------------------------
---------------------------------------------------------------------------------}
Initialization
  registerclasses ( [ TOF_TarifsFourchette ] ) ;
end.
