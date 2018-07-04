unit CalcMulProspect;

interface

uses
hctrls
,Hent1
,UtilTOM
, sysutils
{$IFDEF EAGLCLIENT}
, menuolx
, utob
, emul
{$ELSE}
, mul
, menuolg
, DB
{$IFNDEF DBXPRESS},dbtables{BDE}{$ELSE},uDbxDataSet{$ENDIF}
{$ENDIF}
,hmsgbox
;
Const
  selectRacOper : String = ' AND RAC_OPERATION in ("&#@OPERAGRC")';
  selectTypeAct : String = ' AND RAC_TYPEACTION in ("&#@TYPEACTGRC")';
  selectRpeOper : String = ' AND RPE_OPERATION in ("&#@OPERAGRC")';
  selectEtatAct : String = ' AND RAC_ETATACTION IN ("&#@ETATACTGRC")';
  selectEtatPro : String = ' AND RPE_ETATPER IN ("&#@ETATPROPOGRC")';

Function RTProcCalcMul (Nom,Params,WhereSQL : string ; TT : TDataset ;Total : Boolean) : string ;

implementation

Function RTProcCalcMul (Nom,Params,WhereSQL : string ; TT : TDataset ;Total : Boolean) : string ;
var
  Ecran: TFMul;
  sql: string;

  function TransformeValeur(valeurOri : string) : String;
  var stVal,valeurListe : string;
  begin
  // in ("&#@OPERAGRC")
    result:='';
    if valeurOri='' then
      exit;
    valeurListe:=valeurOri;
    repeat
      stVal:=ReadToKenSt(valeurListe);
      if stVal <> '' then
        if result = '' then
          result :=stVal
        else
          result:=result+'","'+stVal;
    until stVal='';
  end;

  function ChangeInContentAndRun: string;
  var st: string;
      DateDeb,DateFin : TDateTime;
      Q: TQuery;
  begin
    result := '';
    Ecran := nil;
    if (FMenuG.PRien.InsideForm is TFMul) then
      Ecran := (FMenuG.PRien.InsideForm as TFMul);
    if (not assigned(Ecran)) or (assigned(Ecran) and (TFMul(Ecran).name <> 'RTTIERS_PILOTACTC')) then
    begin
      pgiinfo('Cette formule ne doit être utilisée que dans le pilotage de l''activité commerciale');
      exit;
    end;

    st  := TransformeValeur(TOMTOFGetControlText(Ecran, 'OPERATION'));
    if (st = '') or (st = '<<Tous>>') then
      begin
      Delete(sql, Pos(selectRacOper,sql), length(selectRacOper));
      Delete(sql, Pos(selectRpeOper,sql), length(selectRpeOper));
      end
    else
      sql := StringReplace(sql,'&#@OPERAGRC',st,[rfReplaceAll]);

    st  := TransformeValeur(TOMTOFGetControlText(Ecran, 'TYPEACTION'));
    if (st = '') or (st = '<<Tous>>') then
      Delete(sql, Pos(selectTypeAct,sql), length(selectTypeAct))
    else
      sql := StringReplace(sql,'&#@TYPEACTGRC',st,[rfReplaceAll]);

    st  := TransformeValeur(TOMTOFGetControlText(Ecran, 'ETATACTION'));
    if (st = '') or (st = '<<Tous>>') then
      Delete(sql, Pos(selectEtatAct,sql), length(selectTypeAct))
    else
      sql := StringReplace(sql,'&#@ETATACTGRC',st,[rfReplaceAll]);

    st  := TransformeValeur(TOMTOFGetControlText(Ecran, 'ETATPROPOSITION'));
    if (st = '') or (st = '<<Tous>>') then
      Delete(sql, Pos(selectEtatPro,sql), length(selectTypeAct))
    else
      sql := StringReplace(sql,'&#@ETATPROPOGRC',st,[rfReplaceAll]);

    DateDeb := StrToDate(TOMTOFGetControlText(Ecran, 'DATEACTION'));
    DateFin := PlusDate(StrToDate(TOMTOFGetControlText(Ecran, 'DATEACTION_')),1,'J');
    sql := StringReplace(sql,'&#@DATEACTDEB',UsDateTime(DateDeb),[rfReplaceAll]);
    sql := StringReplace(sql,'&#@DATEACTFIN',UsDateTime(DateFin),[rfReplaceAll]);

    DateDeb := StrToDate(TOMTOFGetControlText(Ecran, 'DATEPROPOSITION'));
    DateFin := PlusDate(StrToDate(TOMTOFGetControlText(Ecran, 'DATEPROPOSITION_')),1,'J');
    sql:=StringReplace(sql,'&#@DATEPRODEB',UsDateTime(DateDeb),[rfReplaceAll]);
    sql:=StringReplace(sql,'&#@DATEPROFIN',UsDateTime(DateFin),[rfReplaceAll]);

    try
      Q := OpenSQL(sql,true,-1,'',true);
      if assigned(Q) and (not Q.EOF) then
        begin
        result := Q.Fields[0].AsString;
        if (Nom='RTRPELIBELLE') and (result = 'Perdue') then
          result:=result+' : '+Q.Fields[1].AsString;
        end;
    finally
      Ferme(Q);
    end;
  end;

begin
Result := '';
if Nom='RTISFAX' then
   begin
   Result := 'N';
   //if (pos('((T_FAX <> "") or (C_FAX <> ""))',WhereSQL) <> 0) then  Result := 'O';
   if (pos('(T_FAX <> "")',WhereSQL) <> 0) or (pos('(C_FAX <> "")',WhereSQL) <> 0) then  Result := 'O';
   end
else if Nom='RTARSLIBELLE' then
   begin
   sql := 'SELECT ##TOP 1## ARS_LIBELLE FROM ACTIONS '+
          'LEFT JOIN TIERS ON RAC_AUXILIAIRE=T_AUXILIAIRE '+
          'LEFT JOIN RESSOURCE ON ARS_RESSOURCE = RAC_INTERVENANT '+
          'WHERE RAC_AUXILIAIRE="'+TT.FindField('T_AUXILIAIRE').AsString+'" AND RAC_OPERATION in ("&#@OPERAGRC") AND RAC_TYPEACTION in ("&#@TYPEACTGRC") '+
          'AND RAC_ETATACTION IN ("&#@ETATACTGRC") AND RAC_DATEACTION >= "&#@DATEACTDEB" AND '+
          'RAC_DATEACTION < "&#@DATEACTFIN" ';
   result := ChangeInContentAndRun;
   end
else if Nom='RTSOMMEACT' then
   begin
   sql := 'SELECT COUNT(*) FROM ACTIONS '+
          'LEFT JOIN TIERS ON RAC_AUXILIAIRE = T_AUXILIAIRE '+
          'WHERE RAC_AUXILIAIRE="'+TT.FindField('T_AUXILIAIRE').AsString+'" AND RAC_OPERATION in ("&#@OPERAGRC") AND '+
          'RAC_TYPEACTION in ("&#@TYPEACTGRC") AND RAC_ETATACTION IN ("&#@ETATACTGRC") AND '+
          'RAC_DATEACTION >="&#@DATEACTDEB" AND RAC_DATEACTION < "&#@DATEACTFIN" ';
   result := ChangeInContentAndRun;
   end
else if Nom='RTRPELIBELLE' then
   begin
   sql := 'SELECT ##TOP 1## CO_LIBELLE,RPE_COMMENTPERTE FROM PERSPECTIVES '+
          'LEFT JOIN TIERS ON RPE_AUXILIAIRE = T_AUXILIAIRE '+
          'LEFT JOIN COMMUN ON CO_TYPE="RSP" AND CO_CODE = RPE_ETATPER '+
          'WHERE RPE_AUXILIAIRE = "'+TT.FindField('T_AUXILIAIRE').AsString+'" AND RPE_OPERATION in ("&#@OPERAGRC") '+
          'AND RPE_ETATPER IN ("&#@ETATPROPOGRC") AND RPE_DATEREALISE >= "&#@DATEPRODEB" AND RPE_DATEREALISE < "&#@DATEPROFIN"';
   result := ChangeInContentAndRun;
   end
else if Nom='RTSOMMERPE' then
   begin
   sql := 'SELECT SUM(RPE_MONTANTPER) FROM PERSPECTIVES '+
          'LEFT JOIN TIERS ON RPE_AUXILIAIRE = T_AUXILIAIRE '+
          'WHERE RPE_AUXILIAIRE = "'+TT.FindField('T_AUXILIAIRE').AsString+'" AND '+
          'RPE_OPERATION in ("&#@OPERAGRC") AND RPE_ETATPER IN ("&#@ETATPROPOGRC") AND '+
          'RPE_DATEREALISE >= "&#@DATEPRODEB" AND RPE_DATEREALISE < "&#@DATEPROFIN"';
   result := ChangeInContentAndRun;
   end


else Result:=DefProcCalcMul(Nom,Params,WhereSQL,TT,Total);
end;

end.
