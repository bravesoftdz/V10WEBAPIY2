{***********UNITE*************************************************
Auteur  ...... : Joël TRIFILIEFF
Créé le ...... : 03/10/2000
Modifié le ... : 09/07/2002
Description .. : Source TOF de la FICHE : UTOFGCPIECECOURS_TIER ()
Mots clefs ... : TOF;PIECECOURSTIERS
*****************************************************************}
unit UTofGcPieceCours_Tier;

interface
uses  StdCtrls,Controls,Classes,forms,sysutils,ComCtrls,
      HCtrls,HEnt1,HMsgBox,UTOF,
{$IFDEF EAGLCLIENT}
      UTob,Maineagl,//Fe_Main,
{$ELSE}
      {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}mul, db,DBGrids,Fe_Main,
{$ENDIF}
{$IFDEF NOMADE}
      UtilPOP,
{$ENDIF}
      HDimension,UTOM,AGLInit,M3FP;

function AGLLancePieceCoursTiers( Parms : array of variant ; nb : integer ) : variant ;

Type
     TOF_GcPieceCours_Tier = Class (TOF)
        procedure OnLoad ; override ;
        procedure OnArgument (stArgument : String ) ; override ;
     END ;

implementation

function AGLLancePieceCoursTiers( Parms : array of variant ; nb : integer ) : variant ;
var Nat,Cod,Range,Lequel,Argument : string ;
begin
Nat:=string(Parms[1]);
Cod:=string(Parms[2]);
Range:=string(Parms[3]);
Lequel:=string(Parms[4]);
Argument:=string(Parms[5]);
Result:=AGLLanceFiche(Nat,Cod,Range,Lequel,Argument);
end;

procedure TOF_GcPieceCours_Tier.OnLoad();
Var StCaption ,Libelle:string;
    Q:TQuery;
begin
{$IFDEF NOMADE} //Gestion du <<Tous>> dans le Mul
  if (GetControlText('GP_NATUREPIECEG')='') or (GetControlText('GP_NATUREPIECEG')='<<Tous>>') then
     SetControlText('XX_WHERE',GetNaturePOP('GP_NATUREPIECEG'));
{$ENDIF} //NOMADE
stCaption:= GetControlText('GP_TIERS');
Q:=OpenSQL('SELECT T_LIBELLE FROM TIERS WHERE T_TIERS="'+StCaption+'"',True) ;
Libelle:=Q.Findfield('T_LIBELLE').AsString;
Ecran.Caption:='Pièces en cours :  '+stCaption+' - '+Libelle;
Ferme(Q) ;
end;

procedure TOF_GCPIECECOURS_TIER.OnArgument (stArgument : String ) ;
Var x : integer;
    stArg,critere,ChampMul : string;
{$IFDEF NOMADE}
    NatureDoc : THDBValComboBox;
{$ENDIF} //NOMADE
begin
  Inherited ;
{$IFDEF NOMADE} //Limite liste de nature des pièces
NatureDoc := THDBValComboBox(GetControl('GP_NATUREPIECEG'));
NatureDoc.Plus := GetNaturePOP('GPP_NATUREPIECEG');
{$ENDIF} //NOMADE

stArg:=stArgument;
Repeat
    Critere:=UpperCase(Trim(ReadTokenSt(stArg)));
    if Critere<>'' then
        BEGIN
        x:=pos('=',Critere);
        if x<>0 then
           BEGIN
           ChampMul:=copy(Critere,1,x-1);
           END;
           if ChampMul='ACTION' then SetControlText('TYPEACTION',Critere);
        END;
until Critere='';

if GetControl('bInsert')<>nil then
    SetControlVisible('bInsert',StringToAction(GetControlText('TYPEACTION'))<>taConsult);

end ;


Initialization
registerclasses([TOF_GcPieceCours_Tier]);
RegisterAglFunc('LancePieceCoursTiers',True,5,AGLLancePieceCoursTiers) ;
end.

