{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 31/10/2002
Modifié le ... :   /  /
Description .. : Source TOM de la TABLE : PARACTIONS (PARACTIONS)
Mots clefs ... : TOM;PARACTIONS
*****************************************************************}
Unit UtomParChainages ;

Interface

Uses Classes,
{$IFNDEF EAGLCLIENT}
     db, 
     Fiche,
{$ELSE}
     eFiche,
{$ENDIF}
     forms,
     sysutils,
     HCtrls,
     HEnt1,
     UTOM, UTob, EntRT;
const
  // libellés des messages
  TexteMessage: array[1..1] of string 	= (
    {1}  'Vous ne pouvez pas effacer ce type de chaînage, il est actuellement utilisé dans les chaînages'
    );
Type
  TOM_PARCHAINAGES = Class (TOM)
    procedure OnNewRecord                ; override ;
    procedure OnDeleteRecord             ; override ;
    procedure OnUpdateRecord             ; override ;
    procedure OnAfterUpdateRecord        ; override ;
    procedure OnLoadRecord               ; override ;
    procedure OnChangeField ( F: TField) ; override ;
    procedure OnArgument ( S: String )   ; override ;
    procedure OnClose                    ; override ;
    procedure OnCancelRecord             ; override ;
    Private
      stProduitpgi : string;
    end ;

Implementation

procedure TOM_PARCHAINAGES.OnNewRecord ;
begin
  Inherited ;
  if stProduitpgi='GRF' then
     SetField('RPG_PRODUITPGI','GRF')
  else
     SetField('RPG_PRODUITPGI','GRC');

end ;

procedure TOM_PARCHAINAGES.OnDeleteRecord ;
var TobTypeEncours : tob;
begin
  Inherited ;
  if (Ecran<>nil) and ( ECran is TFFiche ) then
    TFFiche(Ecran).retour := GetField('RPG_CHAINAGE');
  if ExisteSQL( 'SELECT RCH_NUMERO FROM ACTIONSCHAINEES WHERE RCH_CHAINAGE="'+GetField('RPG_CHAINAGE')
     +'" and RCH_PRODUITPGI="'+stProduitpgi+'"') then
  begin
    LastError:=1;
    LastErrorMsg:=TraduireMemoire(TexteMessage[LastError]);
    exit ;
  end ;

  VH_RT.TobTypesChainage.Load;

  TobTypeEncours:=VH_RT.TobTypesChainage.FindFirst(['RPG_CHAINAGE','RPG_PRODUITPGI'],[GetField('RPG_CHAINAGE'),stProduitpgi],TRUE) ;
  if (TobTypeEncours <> Nil) then
     TobTypeEncours.free;
end ;

procedure TOM_PARCHAINAGES.OnUpdateRecord ;
var TobTypeEncours : tob;
begin
  Inherited ;
  VH_RT.TobTypesChainage.Load;

  if ( ds.State = dsEdit ) then
  begin
    TobTypeEncours:=VH_RT.TobTypesChainage.FindFirst(['RPG_CHAINAGE','RPG_PRODUITPGI'],[GetField('RPG_CHAINAGE'),stProduitpgi],TRUE) ;
    if (TobTypeEncours <> Nil) then
       TobTypeEncours.GetEcran(TForm(Ecran),nil);
  end
  else
    if ( ds.State = dsInsert ) then
    begin
      TobTypeEncours:=TOB.create ('PARCHAINAGES',VH_RT.TobTypesChainage,-1);
      if (TobTypeEncours <> Nil) then
       begin
       TobTypeEncours.GetEcran(TForm(Ecran),nil);
       TobTypeEncours.PutValue('RPG_PRODUITPGI',stProduitpgi);
       end;
    end;   
end ;

procedure TOM_PARCHAINAGES.OnAfterUpdateRecord ;
begin
  Inherited ;
  if (Ecran<>nil) and ( Ecran is TFFiche ) then
    begin
    TFFiche(Ecran).retour := GetField('RPG_CHAINAGE');
    end;

end ;

procedure TOM_PARCHAINAGES.OnLoadRecord ;
begin
  Inherited ;
end ;

procedure TOM_PARCHAINAGES.OnChangeField ( F: TField ) ;
begin
  Inherited ;
end ;

procedure TOM_PARCHAINAGES.OnArgument ( S: String ) ;
var Critere,ChampMul,ValMul : string;
    x : integer;
begin
  Inherited ;
  stProduitpgi:='';  
  Repeat
      Critere:=uppercase(ReadTokenSt(S)) ;
      if Critere<>'' then
          begin
          x:=pos('=',Critere);
          if x<>0 then
             begin
             ChampMul:=copy(Critere,1,x-1);
             ValMul:=copy(Critere,x+1,length(Critere));
             if ChampMul='PRODUITPGI' then
                stProduitpgi := ValMul;
             end;
          end;
  until  Critere='';
  if stProduitpgi = 'GRF' then
    begin
    SetControlVisible ('RPG_TABLELIBRECH1',false);
    SetControlVisible ('TRPG_TABLELIBRECH1',false);
    SetControlVisible ('RPG_TABLELIBRECH2',false);
    SetControlVisible ('TRPG_TABLELIBRECH2',false);
    SetControlVisible ('RPG_TABLELIBRECH3',false);
    SetControlVisible ('TRPG_TABLELIBRECH3',false);
    end
  else
    begin
    SetControlVisible ('RPG_TABLELIBRECHF1',false);
    SetControlVisible ('TRPG_TABLELIBRECHF1',false);
    SetControlVisible ('RPG_TABLELIBRECHF2',false);
    SetControlVisible ('TRPG_TABLELIBRECHF2',false);
    SetControlVisible ('RPG_TABLELIBRECHF3',false);
    SetControlVisible ('TRPG_TABLELIBRECHF3',false);
    end;

end ;

procedure TOM_PARCHAINAGES.OnClose ;
begin
  Inherited ;
end ;

procedure TOM_PARCHAINAGES.OnCancelRecord ;
begin
  Inherited ;
end ;

Initialization
  registerclasses ( [ TOM_PARCHAINAGES ] ) ;
end.

