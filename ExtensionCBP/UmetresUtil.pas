unit UmetresUtil;

interface

uses hctrls,
  sysutils,HEnt1,
{$IFNDEF EAGLCLIENT}
  uDbxDataSet, DB,
{$ELSE}
  uWaini,
  {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} DB,
{$ENDIF}
  uhttp,paramsoc,
  utob,uEntCommun;

function GenereLaReferenceMetre(ThisTOB : TOB; Action : TactionFiche; AValider : boolean): string;

implementation

function GenereLaReferenceMetre(ThisTOB : TOB; Action : TactionFiche; AValider : boolean): string;
begin
  result := '';

  if ThisTOB.NomTable = 'PIECE' then
     begin
     if (Action <> taCreat) or (AValider) then
        begin
        Result := ThisTOB.GetValue('GP_NATUREPIECEG')+'-'+ThisTOB.GetValue('GP_SOUCHE')+'-'+
                  inttostr(ThisTOB.GetValue('GP_NUMERO'))+'-'+IntTostr(ThisTOB.GetValue('GP_INDICEG'))+'-';
        end
     else
        begin
        Result := ThisTOB.GetValue('GP_NATUREPIECEG')+'-'+ThisTOB.GetValue('GP_SOUCHE')+'-'+'NOUVEAU-0-';
        end;
     end;

  if ThisTOB.NomTable = 'LIGNE' then
     begin
     Result := ThisTOB.GetValue('GL_NATUREPIECEG')+'-'+ThisTOB.GetValue('GL_SOUCHE')+'-'+
               IntToStr(ThisTOB.GetValue('GL_NUMERO'))+'-'+IntToStr(ThisTOB.GetValue('GL_INDICEG'))+'-'+
               IntToStr(ThisTOB.GetValue('GL_NUMLIGNE'))+'-----';
     end;

  if ThisTOB.NomTable = 'LIGNEOUV' then
     begin
     Result := ThisTOB.GetValue('BLO_NATUREPIECEG')+'-';
     result := result +ThisTOB.GetValue('BLO_SOUCHE')+'-';
     result := result + IntToStr(ThisTOB.GetValue('BLO_NUMERO'))+'-';
     result := result+IntToStr(ThisTOB.GetValue('BLO_INDICEG'))+'-' ;
     result := result+IntToStr(ThisTOB.GetValue('BLO_NUMLIGNE'))+'-';
     result := result+IntToStr(ThisTOB.GetValue('BLO_N1'))+'-';
     result := result+IntToStr(ThisTOB.GetValue('BLO_N2'))+'-';
     result := result+IntToStr(ThisTOB.GetValue('BLO_N3'))+'-';
     result := result+IntToStr(ThisTOB.GetValue('BLO_N4'))+'-';
     result := result+IntToStr(ThisTOB.GetValue('BLO_N5'));
     end;

  if ThisTOB.NomTable = 'ARTICLE' then
     begin
     result := ThisTOB.GetValue('GA_ARTICLE')+';';
     end;

  if ThisTOB.NomTable = 'NOMENENT' then
     begin
     result := ThisTOB.GetValue('GNE_NOMENCLATURE')+';';
     result := result + ThisTOB.GetValue('GNE_ARTICLE')+';';
     end;

  if ThisTOB.NomTable = 'NOMENLIG' then
     begin
     result := ThisTOB.GetValue('GNL_NOMENCLATURE')+';';
     result := result + ThisTOB.GetValue('GNL_ARTICLE')+';';
     result := result + inttostr(ThisTOB.GetValue('GNL_NUMLIGNE'))+';';
     end;

end;

end.
