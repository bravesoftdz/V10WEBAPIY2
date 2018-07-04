{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 01/10/2002
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : AFTIERS_MUL_JP ()
Mots clefs ... : TOF;AFTIERS_MUL_JP
*****************************************************************}
Unit UTOFAFTIERS_MUL_JP ;

Interface

uses  windows,Classes,UTOF, DicoAf, Hent1,SysUtils,forms,HCtrls,UtilPGI,
      HMsgBox,Ent1,AGlINit,UtilGc,UtofAftraducChampLibre
{$IFDEF EAGLCLIENT}
      ,Maineagl,emul
{$ELSE}
		 ,mul,FE_Main
{$ENDIF}
        ;

Type
  TOF_AFTIERS_MUL_JP = Class (TOF_AFTRADUCCHAMPLIBRE)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnClose                  ; override ;
  end ;

  Procedure AFLanceFiche_Mul_Tiers_DP (Range,Argument:string);

Implementation

procedure TOF_AFTIERS_MUL_JP.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_AFTIERS_MUL_JP.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_AFTIERS_MUL_JP.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_AFTIERS_MUL_JP.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_AFTIERS_MUL_JP.OnArgument (S : String ) ;
begin
  inherited ;

  // mcd 24/04/02If ctxTempo in V_PGI.PGIContexte  then
  If not(ctxSCot in V_PGI.PGIContexte)  then
  begin
    SetControlVisible ('TT_MOISCLOTURE' , False);
    SetControlVisible ('T_MOISCLOTURE' , False);
    SetControlVisible ('TT_MOISCLOTURE_' , False);
    SetControlVisible ('T_MOISCLOTURE_' , False);
  end;

  if not ExJaiLeDroitConcept (TConcept (gcCLICreat), False) then
     SetControlVisible('BINSERT',False) ;
end ;

procedure TOF_AFTIERS_MUL_JP.OnClose ;
begin
  Inherited ;
end ;


Procedure AFLanceFiche_Mul_Tiers_DP (Range,Argument:string);
begin
     AGLLanceFiche ('AFF','AFTIERS_MUL_JP', Range, '', Argument);
end;




Initialization
              registerclasses ( [ TOF_AFTIERS_MUL_JP ] ) ;
end.
 
