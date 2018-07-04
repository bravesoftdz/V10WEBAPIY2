unit UtofCubePropositions;

interface

uses  Classes,forms,
      UTOF,
{$Ifdef GIGI}
      EntGC,  Hent1,
{$ENDIf}
     UtilSelection,Cube,UtilRT;

Type
    TOF_CubePropositions = Class (TOF)
        procedure OnArgument (stArgument : String ) ; override ;
        procedure OnLoad ; override;
     END;

implementation
uses ParamSoc;
procedure TOF_CubePropositions.OnArgument(stArgument : String );
var F : TForm;
    stSQL:string;
begin
inherited ;
F := TForm (Ecran);
MulCreerPagesCL(F,'NOMFIC=GCTIERS');
if GetParamSocSecur('SO_RTGESTINFOS00V',False) then
  MulCreerPagesCL(Ecran,'NOMFIC=RTPERSPECTIVES');

stSQL:='LIGNE ';
stSQL:=stSQL+'left join PIECE on GL_NATUREPIECEG=GP_NATUREPIECEG AND  GL_SOUCHE=GP_SOUCHE AND GL_NUMERO=GP_NUMERO AND  GL_INDICEG=GP_INDICEG ';
stSQL:=stSQL+'left join PERSPECTIVES on RPE_PERSPECTIVE=GP_PERSPECTIVE ';
stSQL:=stSQL+'left join RTINFOS00V on RDV_CLEDATA=CAST(GP_PERSPECTIVE AS VARCHAR(17))';
stSQL:=stSQL+'left join PROSPECTS on RPE_AUXILIAIRE = RPR_AUXILIAIRE left Join Tiers on T_AUXILIAIRE=RPE_AUXILIAIRE ';
stSQL:=stSQL+'left join TIERSCOMPL on YTC_AUXILIAIRE = RPE_AUXILIAIRE';
TFCube(ecran).FromSQL:=stSql;
TFCube(ecran).WhereSQL:='GL_TYPELIGNE="ART" AND GP_PERSPECTIVE<>0  and (RPE_VARIANTE=0 or RPE_PERSPECTIVE=RPE_VARIANTE)';
{$Ifdef GIGI}
 if (GetControl('RPE_OPERATION') <> nil) and Not GetParamsocSecur('SO_AFRTOPERATIONS',False)
    then  begin
    SetControlVisible('RPE_OPERATION',false);
    SetControlVisible('TRPE_OPERATION',false);
    end;
 if (GetControl('RPE_PROJET') <> nil) and Not GetParamsocSecur('SO_RTPROJGESTION',False)
    then  begin
    SetControlVisible('RPE_PROJET',false);
    SetControlVisible('TRPE_PROJET',false);
    end;
 if (GetControl('TRPE_REPRESENTANT') <> nil) then  SetControlVisible('TRPE_REPRESENTANT',false);
 if (GetControl('RPE_REPRESENTANT') <> nil) then  SetControlVisible('RPE_REPRESENTANT',false);
 if (GetControl('T_ZONECOM') <> nil) then  SetControlVisible('T_ZONECOM',false);
 if (GetControl('TT_ZONECOM') <> nil) then  SetControlVisible('TT_ZONECOM',false);
 SetControlText ('TRPE_TYPETIERS', TraduireMemoire('Nature tiers'));
 SetControlProperty ('RPE_TYPETIERS', 'Complete', true);
 SetControlProperty ('RPE_TYPETIERS', 'Datatype', 'TTNATTIERS');
 SetControlProperty ('RPE_TYPETIERS', 'Plus', VH_GC.AfNatTiersGRCGI);
{$endif}
end;

procedure TOF_CubePropositions.OnLoad;
begin
inherited;
SetControlText('XX_WHERE',RTXXWhereConfident('CON')) ;
end;

Initialization
registerclasses([TOF_CubePropositions]);

end.
