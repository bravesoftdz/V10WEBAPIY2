{***********UNITE*************************************************
Auteur  ...... : BM
Créé le ...... : 05/11/2002
Modifié le ... :   /  /
Description .. : Source TOM Générique et TOM héritantes pour les
                 TABLES : TOM_JUFORMEJUR, TOM_JUTYPECIVIL,
                 TOM_JUTYPEEVT, TOM_JUTYPEPER
Mots clefs ... : TOM;JUFORMEJUR;JUTYPECIVIL;JUTYPEEVT;JUTYPEPER
*****************************************************************}

Unit UTOMTypePer;
//////////////////////////////////////////////////////////////////
Interface
//////////////////////////////////////////////////////////////////
Uses
{$IFNDEF EAGLCLIENT}
     db,
{$ELSE}
    UTOB,
{$ENDIF}
{$ifdef GRC}
      Hent1,
{$ENDIF}
     Classes, UTOMTypeParam;
//////////////////////////////////////////////////////////////////
Type
   TOM_JUTYPEPER   = Class (TOM_TYPEPARAM)
         procedure OnArgument ( sParams_p : String )   ; override;
         procedure OnLoadRecord               ; override;
   end ;
//////////////////////////////////////////////////////////////////
Implementation
//////////////////////////////////////////////////////////////////
{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. MERIAUX
Créé le ...... : 22/07/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOM_JUTYPEPER.OnArgument ( sParams_p : String ) ;
begin
   Inherited ;
   sTypeElement_f := 'de personne';
   if Getcontrol('JTP_NATUREAUXI') <>Nil then
   begin
      SetControlVisible ('JTP_NATUREAUXI',false);
      SetControlVisible ('TJTP_NATUREAUXI',false);
(* mcd 16/11/2006 c echamp n'est plus géré euqlité  11150
{$ifdef GRC}
      If ctxGRC in V_PGI.PGICOntexte then
      begin
         SetControlVisible ('JTP_NATUREAUXI',true);
         SetControlVisible ('TJTP_NATUREAUXI',true);
      end;
{$endif}  fin mcd 16/11/2006*)
   end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B. Mériaux
Créé le ...... : 03/10/2005
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOM_JUTYPEPER.OnLoadRecord;
var
   bEnable_l : boolean;
begin
   inherited;
{$ifdef GRC}
   If ctxGRC in V_PGI.PGICOntexte then
   begin
      bEnable_l := (DS.State in [dsInsert]) or (GetField('JTP_PREDEFINI') <> 'CEG');
      SetControlEnabled('JTP_FAMPER', bEnable_l);
      SetControlEnabled('JTP_TYPEPERLIB', bEnable_l);
      SetControlEnabled('JTP_TYPEPERABREG', bEnable_l);
      SetControlEnabled('JTP_RACINE', bEnable_l);
      SetControlEnabled('JTP_AFFICHE', bEnable_l);
      SetControlEnabled('JTP_DEFAUTLIB1', bEnable_l);
      SetControlEnabled('JTP_DEFAUTLIB2', bEnable_l);
      SetControlEnabled('JTP_DEFAUTLIB3', bEnable_l);
   end;
{$endif}
end;
//////////////////////////////////////////////////////////////////

Initialization
   registerclasses ( [ TOM_JUTYPEPER ] ) ;
end.
