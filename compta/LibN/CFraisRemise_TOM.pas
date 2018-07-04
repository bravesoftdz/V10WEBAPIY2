{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 04/02/2004
Modifié le ... :   /  /
Description .. : Source TOM de la TABLE : CFRAISREMISE (CFRAISREMISE)
Mots clefs ... : TOM;CFRAISREMISE
*****************************************************************}
Unit CFRAISREMISE_TOM ;

Interface

Uses StdCtrls, 
     Controls, 
     Classes, 
{$IFNDEF EAGLCLIENT}
     db, 
     dbtables, 
     Fiche, 
     FichList, 
{$else}
     eFiche,
     eFichList,
{$ENDIF}
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     UTOM,
     UTob ;

Procedure FicheFraisRemiseESPCP ( CondRemise,Frais : String ; action : TActionFiche ; Argument : String ) ; //XMG 05/04/04

Type
  TOM_CFRAISREMISE = Class (TOM)
   private
    CCB_GENERALBanque : String ;
   protected
    procedure OnNewRecord                ; override ;
    procedure OnDeleteRecord             ; override ;
    procedure OnUpdateRecord             ; override ;
    procedure OnAfterUpdateRecord        ; override ;
    procedure OnLoadRecord               ; override ;
    procedure OnChangeField ( F: TField) ; override ;
    procedure OnArgument ( S: String )   ; override ;
    procedure OnClose                    ; override ;
    procedure OnCancelRecord             ; override ;
   public
    end ;

Implementation

uses
  {$IFDEF MODENT1}
  CPTypeCons,
  CPProcMetier,
  {$ENDIF MODENT1}
  AGLInit, HDB, Ent1, ULibEncaDecaESP, LookUp, utilDiv, //XMG 05/04/04
{$IFDEF eAGLClient}
    MaineAGL         // AGLLanceFiche (et non pas AGLMaine !!)
{$ELSE}
    FE_Main
{$ENDIF eAGLClient}
    ;
//==================================================
// Definition des variable
//==================================================
var
    MESS : array [0..5] of string = (
     {0}   'Vous devez renseigner une libellé!',
     {1}   'le calcule des frais est nulle!',
     {2}   'Vous devez renseigner un compte général pour les frais!',
     {3}   'Le compte pour les frais renseigné n''existe pas!',
     {4}   'Le journal choisi n''appartiant pas au compte général de la banque choisi!',
     {5}   'On n''a pas renseigner le journal de génération!'
    ) ;


//XMG 05/04/04 début
Procedure FicheFraisRemiseESPCP ( CondRemise,Frais : String ; action : TActionFiche ; Argument : String ) ;
Begin
  if (Action<>taConsult) and (not ExisteSQL('select CFR_CONDREMISE from CFRAISREMISE where CFR_CONDREMISE="'+CondRemise+'" and (CFR_FRAISREMISE="'+Frais+'" or "'+Frais+'"="")')) then
     Action:=taCreat ;
  Argument:=actiontostring(Action)+';'+Argument ;
  AGLLanceFiche('CP','CPESPFRAISREMISE',CondRemise,Frais,Argument)
End ;
//XMG 05/04/04 fin

procedure TOM_CFRAISREMISE.OnNewRecord ;
begin
  Inherited ;
  SetField('CFR_TYPESFRAIS','REM') ;
  SetField('CFR_BASECALCUL','REM') ;
  SetField('CFR_TVA','') ;
end ;

procedure TOM_CFRAISREMISE.OnDeleteRecord ;
begin
  Inherited ;
end ;

procedure TOM_CFRAISREMISE.OnUpdateRecord ;
begin
  Inherited ;
  if trim(GetField('CFR_LIBELLE'))='' then
     Lasterror:=1
  else
  if Arrondi(getfield('CFR_FRAISMIN')+Getfield('CFR_FRAISFIXE')+GetField('CFR_TAUXFRAIS'),CFR_NbrDecTauxFrais)=0 then
     Lasterror:=2
  else
  if GetField('CFR_GENERALFRAIS')='' then
     LastError:=3
  else
  if (not LookUpValueExist( GetControl('CFR_GENERALFRAIS'))) then
     LastError:=4
  else
  if trim(getField('CFR_JOURNAL'))='' then
     LastError:=6
  else
  if not ExisteSQL('select J_JOURNAL from JOURNAL where J_JOURNAL="'+GetField('CFR_JOURNAL')+'" and J_NATUREJAL="BQE" and J_CONTREPARTIE="'+CCB_GENERALBanque+'"') then
     LastError:=5
  ;
  if LastError>0 then
     LastErrorMsg:=Mess[LastError-1] ;
end ;

procedure TOM_CFRAISREMISE.OnAfterUpdateRecord ;
begin
  Inherited ;
end ;

procedure TOM_CFRAISREMISE.OnLoadRecord ;
begin
  Inherited ;
end ;

procedure TOM_CFRAISREMISE.OnChangeField ( F: TField ) ;
var St : String ; 
begin
  Inherited ;
  if (F.FieldName='CFR_GENERALFRAIS') and (Ds.State in [dsEdit,dsInsert]) and (F.AsString<>'') then begin
     St:=Bourreladonc(F.AsString,fbGene) ;
     if St<>F.AsString then F.AsString:=St ;
  End ;
end ;

procedure TOM_CFRAISREMISE.OnArgument ( S: String ) ;
var StrMask : string ;
begin
  Inherited ;
  if Assigned(Ecran) then begin
    TFFicheListe(Ecran).TypeAction:=stringtoaction('ACTION='+TrouveArgument(s,'ACTION',TFFicheListe(Ecran).TypeAction)) ;
    CCB_GENERALBanque:=TrouveArgument(s,'Banque=','') ;
    StrMask:=strfMask(V_PGI.OkDecV,'',TRUE,FALSE) ;
    THDBEdit(GetControl('CFR_FRAISMIN')).DisplayFormat:=strMask ;
    THDBEdit(GetControl('CFR_FRAISFIXE')).DisplayFormat:=strMask ;
    THDBEdit(GetControl('CFR_TAUXFRAIS')).DisplayFormat:=strfMask(CFR_NbrDecTauxFrais,'',TRUE,FALSE) ;
  End ;
end ;

procedure TOM_CFRAISREMISE.OnClose ;
begin
  Inherited ;
end ;

procedure TOM_CFRAISREMISE.OnCancelRecord ;
begin
  Inherited ;
end ;

Initialization
  registerclasses ( [ TOM_CFRAISREMISE ] ) ; 
end.
