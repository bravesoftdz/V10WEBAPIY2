{***********UNITE*************************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 10/09/2002
Modifié le ... : 30/09/2002 Refonte complète du code pour traiter le
               : référentiel métier de façon à gérer les min,max ...
Description .. : Gestion de la modification du coeff
Mots clefs ... : PAIE;PGMODIFCOEFF
*****************************************************************}
{
PT1   : 07/06/2004 VG V_50 Affichage des éléments saisissables en fonction du
                           paramétrage stocké dans la table EVOLUTIONSAL
}
unit UTOMModifCoeff;

interface
uses     UTOM,
{$IFNDEF EAGLCLIENT}
         DB,{$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ELSE}
         UTOB,
{$ENDIF}
         Hctrls,HEnt1,Classes,Entpaie, sysutils;

Type
      TOM_ModifCoeff = Class (TOM)
      public
        procedure OnArgument (stArgument : String ) ; override ;
        procedure OnNewRecord ; override ;
        procedure OnLoadRecord ; override ;
        procedure OnUpdateRecord  ; override ;
        procedure OnAfterUpdateRecord ; override;
        procedure OnChangeField (F : TField) ; override ;
        procedure OnClose ; override ;
      private
        procedure GestionAffich (Quoi : string);
     END ;

implementation

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 10/09/2002
Modifié le ... :   /  /
Description .. : OnArgument
Mots clefs ... : PAIE;PGMODIFCOEFF
*****************************************************************}
procedure TOM_ModifCoeff.OnArgument(stArgument: String);
var
Arg, Orig, St, St1, State, StParamEvo, StQualEvo : String;
ParamEvo : boolean;
Q : TQuery;
begin
Inherited ;
Arg := stArgument;
State := Trim(ReadTokenPipe(Arg,';'));
Orig := Trim(ReadTokenPipe(Arg,';')) ;
if (Orig = 'COE') then
   begin
   GestionAffich ('EMP');
   GestionAffich ('QUA');            //PT1
   GestionAffich ('COE');
   GestionAffich ('IND');
   GestionAffich ('NIV');
   GestionAffich ('NBA');
   GestionAffich ('NB2');
//PT1
   StParamEvo:='SELECT * FROM EVOLUTIONSAL';
   Q := OpenSql(StParamEvo, True);
   if (Not Q.EOF) then
      begin
      ParamEvo:= (Q.FindField('PVL_ANCIENPOSTEP').AsString='X');
      SetControlEnabled('TPMC_NBANCIEN', ParamEvo);
      SetControlEnabled('PMC_NBANCIEN', ParamEvo);
      SetControlEnabled('TPMC_NBANCIENSUP', ParamEvo);
      SetControlEnabled('PMC_NBANCIENSUP', ParamEvo);
      SetControlEnabled('LBLNBA', ParamEvo);
      ParamEvo:= (Q.FindField('PVL_COEFFICIENTP').AsString='X');
      SetControlEnabled('TPMC_COEFFICIENT', ParamEvo);
      SetControlEnabled('PMC_COEFFICIENT', ParamEvo);
      SetControlEnabled('LBLCOE', ParamEvo);
      ParamEvo:= (Q.FindField('PVL_LIBEMPLOIP').AsString='X');
      SetControlEnabled('TPMC_LIBELLEEMPLOI', ParamEvo);
      SetControlEnabled('PMC_LIBELLEEMPLOI', ParamEvo);
      SetControlEnabled('LBLEMP', ParamEvo);
      ParamEvo:= (Q.FindField('PVL_QUALIFP').AsString='X');
      SetControlEnabled('TPMC_QUALIFICATION', ParamEvo);
      SetControlEnabled('PMC_QUALIFICATION', ParamEvo);
      SetControlEnabled('LBLQUA', ParamEvo);
      ParamEvo:= (Q.FindField('PVL_NIVEAUP').AsString='X');
      SetControlEnabled('TPMC_NIVEAU', ParamEvo);
      SetControlEnabled('PMC_NIVEAU', ParamEvo);
      SetControlEnabled('LBLNIV', ParamEvo);
      ParamEvo:= (Q.FindField('PVL_INDICEP').AsString='X');
      SetControlEnabled('TPMC_INDICE', ParamEvo);
      SetControlEnabled('PMC_INDICE', ParamEvo);
      SetControlEnabled('LBLIND', ParamEvo);
      StQualEvo := Q.FindField('PVL_EVOLUANT').AsString;
      if (StQualEvo='001') then
         begin
         SetControlEnabled('TPMC_COEFFICIENT', True);
         SetControlEnabled('PMC_COEFFICIENT', True);
         SetControlEnabled('LBLCOE', True);
         end
      else
      if (StQualEvo='002') then
         begin
         SetControlEnabled('TPMC_QUALIFICATION', True);
         SetControlEnabled('PMC_QUALIFICATION', True);
         SetControlEnabled('LBLQUA', True);
         end
      else
      if (StQualEvo='003') then
         begin
         SetControlEnabled('TPMC_NIVEAU', True);
         SetControlEnabled('PMC_NIVEAU', True);
         SetControlEnabled('LBLNIV', True);
         end
      else
      if (StQualEvo='004') then
         begin
         SetControlEnabled('TPMC_INDICE', True);
         SetControlEnabled('PMC_INDICE', True);
         SetControlEnabled('LBLIND', True);
         end;
      end;
   Ferme(Q);
//FIN PT1
   end
else
   begin
   st := VH_Paie.PGChampBudget ;
   st1 := ReadTokenSt (st);
   while St1 <> '' do
      begin
      GestionAffich (St1);
      st1 := ReadTokenSt (st);
      end;
   GestionAffich ('MIN');
   GestionAffich ('MAX');
   end;
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 10/09/2002
Modifié le ... :   /  /
Description .. : OnNewRecord
Mots clefs ... : PAIE;PGMODIFCOEFF
*****************************************************************}
procedure TOM_ModifCoeff.OnNewRecord;
var
NumOrdre : integer;
QCoeff : TQuery;
begin
Inherited ;

NumOrdre:=1;
QCoeff:=OpenSQL('SELECT MAX(PMC_MODIFCOEFF) AS NUMMAX FROM MODIFCOEFF',TRUE);
if Not QCoeff.EOF then
   try
   NumOrdre:= QCoeff.Fields[0].AsInteger+1;
   except
         on E: EConvertError do
            NumOrdre:= 1;
   end;
Ferme(QCoeff);
SetField ('PMC_MODIFCOEFF',NumOrdre);
SetField ('PMC_SALAIREMIN', 0);
SetField ('PMC_SALAIREMAX', 0);
SetField ('PMC_DATEANCIEN', Idate1900);
SetField ('PMC_DATEANCIENFIN', Idate1900);
SetField ('PMC_DATEAPPLIC', Idate1900);
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 10/09/2002
Modifié le ... :   /  /
Description .. : OnLoadRecord
Mots clefs ... : PAIE;PGMODIFCOEFF
*****************************************************************}
procedure TOM_ModifCoeff.OnLoadRecord;
begin
Inherited ;
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 10/09/2002
Modifié le ... :   /  /
Description .. : OnUpdateRecord
Mots clefs ... : PAIE;PGMODIFCOEFF
*****************************************************************}
procedure TOM_ModifCoeff.OnUpdateRecord;
begin

end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 10/09/2002
Modifié le ... :   /  /
Description .. : OnAfterUpdateRecord
Mots clefs ... : PAIE;PGMODIFCOEFF
*****************************************************************}
procedure TOM_ModifCoeff.OnAfterUpdateRecord;
begin
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 10/09/2002
Modifié le ... :   /  /
Description .. : OnChangeField
Mots clefs ... : PAIE;PGMODIFCOEFF
*****************************************************************}
procedure TOM_ModifCoeff.OnChangeField;
begin
if ((F.FieldName=('PMC_NBANCIEN')) or (F.FieldName=('PMC_NBANCIENSUP'))) then
   ExecuteSQL('UPDATE MODIFCOEFF SET'+
              ' PMC_DATEANCIEN="'+UsDateTime(IDate1900)+'" WHERE'+
              ' PMC_MODIFCOEFF=0'); {DB2 SB 15/04/2005 }
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 10/09/2002
Modifié le ... :   /  /
Description .. : OnClose
Mots clefs ... : PAIE;PGMODIFCOEFF
*****************************************************************}
procedure TOM_ModifCoeff.OnClose;
begin
end;

procedure TOM_ModifCoeff.GestionAffich (Quoi : string);
var Champ : String;
begin
if Quoi = 'COE' then Champ := 'PMC_COEFFICIENT'
 else if Quoi = 'QUA' then Champ := 'PMC_QUALIFICATION'
  else if Quoi = 'IND' then Champ := 'PMC_INDICE'
   else if Quoi = 'NIV' then Champ := 'PMC_NIVEAU'
    else if Quoi = 'EMP' then Champ := 'PMC_LIBELLEEMPLOI'
     else if Quoi = 'MIN' then Champ := 'PMC_SALAIREMIN'
      else if Quoi = 'MAX' then Champ := 'PMC_SALAIREMAX'
       else if Quoi = 'NBA' then Champ := 'PMC_NBANCIEN'
        else if Quoi = 'NB2' then Champ := 'PMC_NBANCIENSUP'
         else exit;
SetControlVisible (Champ,true);
SetControlVisible ('T'+Champ,true);
SetControlVisible ('LBL'+Quoi,true);
end;


Initialization
registerclasses([TOM_ModifCoeff]) ;

end.

