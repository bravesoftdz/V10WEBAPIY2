unit CalcMulPGI;

interface

uses
  {$IFNDEF EAGLCLIENT}
    DB,
    UEdtComp,
    {$IFNDEF DBXPRESS}dbtables{BDE},{$ELSE}uDbxDataSet,{$ENDIF}
  {$ENDIF !EAGLCLIENT}
  utob,
  HEnt1,
  SysUtils
  ;

{$IFDEF VER150}
function PGIProcCalcMul(Nom,Params,WhereSQL : hstring ; TT : TDataset ;Total : Boolean) : hstring ;
{$ELSE VER150}
function PGIProcCalcMul(Nom,Params,WhereSQL : string ; TT : TDataset ;Total : Boolean) : string ;
{$ENDIF VER150}

implementation
uses
{$IFDEF VER150}
  Variants,
{$ENDIF}

  {$IFDEF GRC}
    CalcMulProspect,
  {$ENDIF GRC}
  {$IFDEF GPAOLIGHT}
    CalcMulGPAO,
  {$ENDIF GPAOLIGHT}
  {$IFDEF CHR}
    CalcMulCHR,
  {$ENDIF CHR}
  {$IFDEF FOS5}
    CalcMulFO,
  {$ENDIF FOS5}
  {$IFDEF STK}
    EntGC,
    UtilArticle,
    DispoDetail,
    StockUtil,
  {$ENDIF STK}
  {$IFDEF ACCESSCM}
    UFormuleMul,
  {$ENDIF ACCESSCM}
  { Objectifs }
  {$IFNDEF PGISIDE}
  {$IF (Defined(GESCOM) AND Defined(BUSINESSPLACE)) OR Defined(PAIEGRH)}
    UFormuleMulBP,
  {$IFEND (Defined(GESCOM) AND Defined(BUSINESSPLACE)) OR Defined(PAIEGRH)}
  {$ENDIF !PGISIDE}
  hctrls,
  MeaUtil,
  wCommuns
  ;

function PGIProcCalcMul(Nom,Params,WhereSQL : hstring ; TT : TDataset ;Total : Boolean) : hstring ;
var
  s2, s3, s4, s5, s6, s7, s8, s9: string;
  sParamsSave : string;
  s1 : hstring;
{$IFDEF STK}
  FieldName, TableName : string;
  QualifUniteSto, UniteQteVte, UniteQteAch, UniteQteProd, UniteQteConso: string;
  CoefConvQteAch, CoefConvQteVte, CoefConvQtePro, CoefConvQteCon: Double;
  GAFields                          : MyArrayValue;
  Quantite                          : double;
{$ENDIF !STK}
{$IFDEF AFFAIRE}
  vSt : String;
  vQr : TQuery;
{$ENDIF AFFAIRE}

  function GetFormuleType: string; // afficher le type de frais dans liste TIERSFRAIS
  var
    Prefixe, TypePort, ParamsSave: string;
  begin
    Result := '';
    ParamsSave := Params;
    Prefixe := wExtractGuillemet(ReadTokenPipe(ParamsSave, ';'));
    try
      if Prefixe = 'GTF' then
      begin
        TypePort:= wGetSqlFieldValue('GPO_TYPEPORT', 'PORT', 'GPO_CODEPORT="' + TT.FindField('GTF_CODEPORT').Value + '"');
        if TypePort <> '' then
          Result := RechDom('GCTYPEPORT', TypePort, False)
      end
    except
    end;
  end;

  procedure GetParam;
  begin
    sParamsSave := Params;
    try
      s1 := ReadTokenPipe(sParamsSave, ';');
      s2 := ReadTokenPipe(sParamsSave, ';');
      s3 := ReadTokenPipe(sParamsSave, ';');
      s4 := ReadTokenPipe(sParamsSave, ';');
      s5 := ReadTokenPipe(sParamsSave, ';');
      s6 := ReadTokenPipe(sParamsSave, ';');
      s7 := ReadTokenPipe(sParamsSave, ';');
      s8 := ReadTokenPipe(sParamsSave, ';');
      s9 := ReadTokenPipe(sParamsSave, ';');
    finally { en cas de planté, on conserve les paramètres }
      sParamsSave := Params
    end
  end;

  {$IFDEF STK}
  function CheckAndGetUniteLib(const Mesure: String): String;
  var
    T: Tob;
  begin
    VH_GC.MTobMEA.Load;
    T := VH_GC.MTobMEA.FindFirst(['GME_MESURE'], [Mesure], True);
    if (Mesure = '') or not Assigned(T) then
      Result := ''
    else
      Result := T.GetString('GME_LIBELLE')
  end;
{$ENDIF STK}

begin
  Result := '';

  {$IFDEF STK}
  GAFields := nil;
  {$ENDIF STK}

  {$IFDEF GRC}
  if copy(Nom, 1, 2) = 'RT' then
  begin
    Result := RTProcCalcMul(Nom,Params,WhereSQL,TT,Total);
    Exit;
  end else
  {$ENDIF}
  if Nom = 'GPICONEMEMO' then
  begin
    try
      Result := iif(TT.FindField(ReadTokenPipe(Params, ';') + '_WBMEMO').Value = wTrue, '#ICO#' + '4', '');
    except
    end;
  end else
  {$IFDEF GPAOLIGHT}
  if copy(Nom, 1, 2) = 'GP' then
  begin
    Result := wProcCalcMul(Nom,Params,WhereSQL,TT,Total);
    Exit;
  end else
  {$ENDIF GPAOLIGHT}
  {$IFDEF AFFAIRE} // si pbm cf gm / cb de la Ga, ce code était dans mdispga ,pourquoi ?
  if Nom='LIENOLE' then
  begin
    vSt := 'SELECT LO_LIBELLE FROM LIENSOLE WHERE LO_EMPLOIBLOB = "' + params + '"';
    vSt := vSt + ' AND LO_IDENTIFIANT = "' + TT.findfield('AFF_AFFAIRE').asstring + '"';
    Try
      vQR := OpenSql(vSt, True);
      if Not vQR.Eof then
        result := vQR.findfield('LO_LIBELLE').AsString
      else
        result := '';
    Finally
      if vQR <> nil then ferme(vQR);
    End
  end else
  {$ENDIF AFFAIRE}
  {$IFDEF CHR}
  if copy(Nom, 1, 2) = 'HR' then
  begin
    Result := HRProcCalcMul(Nom,Params,WhereSQL,TT,Total);
    Exit;
  end else
  {$ENDIF CHR}
  {$IFDEF FOS5}
  if copy(Nom, 1, 2) = 'FO' then
  begin
    Result := FOProcCalcMul(Nom,Params,WhereSQL,TT,Total);
    Exit;
  end else
  {$ENDIF FOS5}

  { Formules communes à tous les marchés }
  if Nom='DATETOHEURE' then
  begin
     if Assigned(TT.FindField(Params)) then Result := FormatdateTime('hh:nn',TT.FindField(Params).AsDateTime)
  end
  else if Nom = 'SETFORMULETYPE' then Result := GetFormuleType
  {$IFDEF STK}
  else if Nom = 'CONVERTFLUX' then
  begin
    GetParam;

    Quantite := 0;

    CoefConvQteAch := 1;
    CoefConvQteVte := 1;
    CoefConvQtePro := 1;
    CoefConvQteCon := 1;

    { Quantite }
    if Total then
    begin
      { FieldName }
      if      s1 = 'GQ_DISPO'          then FieldName := 'SUM(GQ_PHYSIQUE-GQ_BLOCAGE-GQ_RESERVECLI) AS GQ_DISPO'
      else if s1 = 'GQ_PROJE'          then FieldName := 'SUM(GQ_PHYSIQUE-GQ_RESERVECLI+GQ_RESERVEFOU) AS GQ_PROJE'
      else if s1 = 'ABS(GSM_QPREVUE)'  then FieldName := 'SUM(ABS(GSM_QPREVUE))'
      else if s1 = 'ABS(GSM_QPREPA)'   then FieldName := 'SUM(ABS(GSM_QPREPA))'
      else if s1 = 'ABS(GSM_QRUPTURE)' then FieldName := 'SUM(ABS(GSM_QRUPTURE))'
      else if s1 = 'ABS(GSM_PHYSIQUE)' then FieldName := 'SUM(GSM_PHYSIQUE)'
      else if s1 = 'ABS(GSM_QRESTE)'   then FieldName := 'SUM(IIF(GSM_ETATMVT="SOL", 0, ABS(GSM_QPREVUE-GSM_QPREPA-GSM_PHYSIQUE))) AS GSM_QRESTE'
      else FieldName := 'SUM(' + s1 + ')'
      ;

      { TableName }
      if      Pos('GSM_',  s1) > 0 then
      begin
        if pos('WOL_', WhereSQL) > 0 then //Liste GCSTKRUPSTA Rupture sur ordre sous traitance d'achat
          TableName := 'STKMOUVEMENT '
                     + ' LEFT JOIN ARTICLE ON GA_ARTICLE=GSM_ARTICLE'
                     + ' LEFT JOIN STKNATURE ON GSN_QUALIFMVT=GSM_QUALIFMVT'
                     + ' LEFT JOIN WORDRELIG ON WOL_NATURETRAVAIL=GSM_NATUREORI AND WOL_LIGNEORDRE=GSM_NUMEROORI'
        else
          TableName := 'STKMOUVEMENT '
                     + ' LEFT JOIN ARTICLE ON GA_ARTICLE=GSM_ARTICLE'
                     + ' LEFT JOIN STKNATURE ON GSN_QUALIFMVT=GSM_QUALIFMVT';
      end
      else if Pos('GQ_', s1)   > 0 then
      begin
        TableName := 'DISPO LEFT JOIN ARTICLE ON GA_ARTICLE=GQ_ARTICLE';
        if VH_GC.GCIfDefCEGID then
          TableName := TableName + ' LEFT OUTER JOIN EMPLACEMENT GEM2 ON GQ_EMPLACEMENT=GEM2.GEM_EMPLACEMENT'
      end
{$IFDEF STK}
      else if Pos('GQD_', s1)  > 0 then TableName := GetGQDTableName + ' LEFT JOIN ARTICLE ON (GA_ARTICLE=GQD_ARTICLE)'
{$ENDIF STK}      
      ;

      if TableName <> '' then
      begin
        GAFields := wGetSQLFieldsValues([FieldName, 'MIN(GA_QUALIFUNITESTO)', 'MIN(' + GAUniteQteVente + ')', 'MIN(' + GAUniteQteAchat + ')', 'MIN(' + GAUniteQteProd + ')', 'MIN(' + GAUniteQteConso + ')', 'MIN(GA_COEFCONVQTEACH)', 'MIN(GA_COEFCONVQTEVTE)', 'MIN(GA_COEFPROD)', 'MIN(GA_COEFCONVCONSO)'], TableName, WhereSQL);
        if string(GAFields[0]) <> '' then
        begin
          Quantite       := valeur(GAFields[0]);
          QualifUniteSto := GAFields[1];
          UniteQteVte    := GAFields[2];
          UniteQteAch    := GAFields[3];
          UniteQteProd   := GAFields[4];
          UniteQteConso  := GAFields[5];
          CoefConvQteAch := valeur(GAFields[6]);
          CoefConvQteVte := valeur(GAFields[7]);
          CoefConvQtePro := valeur(GAFields[8]);
          CoefConvQteCon := valeur(GAFields[9]);
        end;
      end;
    end
    else
    begin
      if      s1 = 'GQ_DISPO'          then Quantite := Valeur(VarToStr(tt.FindField('GQ_PHYSIQUE').Value - tt.FindField('GQ_BLOCAGE').Value - tt.FindField('GQ_RESERVECLI').Value))
      else if s1 = 'GQ_PROJE'          then Quantite := Valeur(VarToStr(tt.FindField('GQ_PHYSIQUE').Value - tt.FindField('GQ_RESERVECLI').Value + tt.FindField('GQ_RESERVEFOU').Value))
      else if s1 = 'ABS(GSM_QPREVUE)'  then Quantite := Abs(Valeur(VarToStr(tt.FindField('GSM_QPREVUE').Value)))
      else if s1 = 'ABS(GSM_QPREPA)'   then Quantite := Abs(Valeur(VarToStr(tt.FindField('GSM_QPREPA').Value)))
      else if s1 = 'ABS(GSM_QRUPTURE)' then Quantite := Abs(Valeur(VarToStr(tt.FindField('GSM_QRUPTURE').Value)))
      else if s1 = 'ABS(GSM_PHYSIQUE)' then Quantite := Abs(Valeur(VarToStr(tt.FindField('GSM_PHYSIQUE').Value)))
      else if s1 = 'ABS(GSM_QRESTE)'   then Quantite := iif(VarToStr(tt.FindField('GSM_ETATMVT').Value) <> 'SOL', Abs(Valeur(VarToStr(tt.FindField('GSM_QPREVUE').Value - tt.FindField('GSM_QPREPA').Value - tt.FindField('GSM_PHYSIQUE').Value))), 0)
      else                                  Quantite := Valeur(VarToStr(tt.FindField(s1).Value));

      QualifUniteSto := VarToStr(tt.FindField('GA_QUALIFUNITESTO').Value);
      UniteQteVte    := VarToStr(tt.FindField(GAUniteQteVente).Value);
      UniteQteAch    := VarToStr(tt.FindField(GAUniteQteAchat).Value);
      UniteQteProd   := VarToStr(tt.FindField(GAUniteQteProd).Value);
      UniteQteConso  := VarToStr(tt.FindField(GAUniteQteConso).Value);
      CoefConvQteAch := 1;
      CoefConvQteVte := 1;
      CoefConvQtePro := 1;
      CoefConvQteCon := 1;
      if Assigned(TT.FindField('GA_COEFCONVQTEACH')) and (TT.FindField('GA_COEFCONVQTEACH').Value <> null)then
        CoefConvQteAch := VarAsType(TT.FindField('GA_COEFCONVQTEACH').Value, varDouble);
      if Assigned(TT.FindField('GA_COEFCONVQTEVTE')) and (TT.FindField('GA_COEFCONVQTEVTE').Value <> null) then
        CoefConvQteVte := VarAsType(TT.FindField('GA_COEFCONVQTEVTE').Value, varDouble);
      if Assigned(TT.FindField('GA_COEFPROD')) and (TT.FindField('GA_COEFPROD').Value <> null) then
        CoefConvQtePro := VarAsType(TT.FindField('GA_COEFPROD').Value, varDouble);
      if Assigned(TT.FindField('GA_COEFCONVCONSO')) and (TT.FindField('GA_COEFCONVCONSO').Value <> null) then
        CoefConvQteCon := VarAsType(TT.FindField('GA_COEFCONVCONSO').Value, varDouble);
    end;

    { Conversion }
    if      VH_GC.PmFlux.Flux = 'STO' then Result := FloatToStr(Quantite)
    else if VH_GC.PmFlux.Flux = 'VTE' then Result := FloatToStr(wDivise(Quantite, CoefConvQteVte))
    else if VH_GC.PmFlux.Flux = 'ACH' then Result := FloatToStr(wDivise(Quantite, CoefConvQteAch))
    else if VH_GC.PmFlux.Flux = 'PRO' then Result := FloatToStr(wDivise(Quantite, CoefConvQtePro))
    else if VH_GC.PmFlux.Flux = 'CON' then Result := FloatToStr(wDivise(Quantite, CoefConvQteCon));
  end
  {$ENDIF STK}
  {$IFDEF STK}
  else if Nom = 'MESUREFLUX' then
  begin
    GetParam;
    if      VH_GC.PmFlux.Flux = 'STO' then Result := CheckAndGetUniteLib(VarToStr(TT.FindField('GA_QUALIFUNITESTO').Value))
    else if VH_GC.PmFlux.Flux = 'VTE' then Result := CheckAndGetUniteLib(VarToStr(TT.FindField(GAUniteQteVente).Value))
    else if VH_GC.PmFlux.Flux = 'ACH' then Result := CheckAndGetUniteLib(VarToStr(TT.FindField(GAUniteQteAchat).Value))
    else if VH_GC.PmFlux.Flux = 'PRO' then Result := CheckAndGetUniteLib(VarToStr(TT.FindField(GAUniteQteProd).Value))
    else if VH_GC.PmFlux.Flux = 'CON' then Result := CheckAndGetUniteLib(VarToStr(TT.FindField(GAUniteQteConso).Value))
  end
  {$ENDIF STK}
  {$IFDEF STK}
  else if Nom = 'STKAFFECTATION' then   { afficher la REFAFFECTATION en littéral }
  begin
    GetParam;
    if (s1 <> '') and Assigned(TT.FindField(s1 + '_REFAFFECTATION')) then
    begin
      Result := GetLibAffectation(TT.FindField(s1 + '_REFAFFECTATION').AsString);
    end;  
  end
  {$ENDIF STK}
  else if (Nom = 'TARIFTIERS') and Assigned(TT.FindField('YTS_TARIFTIERS')) then
  begin
    if TT.FindField('YTS_TARIFTIERS').AsString <> '' then
    begin
      if Params = 'FOU' then
        Result := RechDom('TTTARIFFOURNISSEUR', TT.FindField('YTS_TARIFTIERS').AsString, False)
      else if Params = 'CLI' then
        Result := RechDom('TTTARIFCLIENT'     , TT.FindField('YTS_TARIFTIERS').AsString, False)
      else
        Result := TraduireMemoire('Erreur')
    end
    else
      Result := ''
  end
  else if (Nom = 'LIBELLETIERS') and Assigned(TT.FindField('YTS_TIERS')) then
  begin
    if TT.FindField('YTS_TIERS').AsString <> '' then
      Result := RechDom('GCTIERS', TT.FindField('YTS_TIERS').AsString, False)
    else
      Result := ''
  end
  else if (Nom = 'LIBELLEARTICLE') and Assigned(TT.FindField('YTS_ARTICLE')) then
  begin
    if TT.FindField('YTS_ARTICLE').AsString <> '' then
      Result := RechDom('GCARTICLE', TT.FindField('YTS_ARTICLE').AsString, False)
    else
      Result := ''
  end
  else if (Nom = 'TARIFTIERS') and Assigned(TT.FindField('WPF_TARIFTIERS')) then
  begin
    if TT.FindField('WPF_TARIFTIERS').AsString <> '' then
    begin
      if Params = 'FOU' then
        Result := RechDom('TTTARIFFOURNISSEUR', TT.FindField('WPF_TARIFTIERS').AsString, False)
      else if Params = 'CLI' then
        Result := RechDom('TTTARIFCLIENT'     , TT.FindField('WPF_TARIFTIERS').AsString, False)
      else
        Result := TraduireMemoire('Erreur')
    end
    else
      Result := ''
  end
  //GP_20071206_MM_GC15600 Déb
  else if (Nom = 'TARIFTIERS') then
  begin
    GetParam;
    if Assigned(TT.FindField(s2+'_TARIFTIERS')) and (TT.FindField(s2+'_TARIFTIERS').AsString<>'') then
    begin
      if s1 = 'FOU' then
        Result := RechDom('TTTARIFFOURNISSEUR', TT.FindField(s2+'_TARIFTIERS').AsString, False)
      else if s1 = 'CLI' then
        Result := RechDom('TTTARIFCLIENT'     , TT.FindField(s2+'_TARIFTIERS').AsString, False)
      else
        Result := TraduireMemoire('Erreur')
    end
    else
      Result := ''
  end
  //GP_20071206_MM_GC15600 Fin
  {$IFDEF STK}
	  else if Nom = 'ARTICLETOCODEARTICLE' then { Convertit un article en codearticle }
	  begin
	    if Assigned(TT.FindField(Params)) then
	      Result := wGetCodeArticleFromArticle(TT.FindField(Params).AsString)
	    else
	      Result := ''
	  end
  {$ENDIF STK}
{$IFDEF ACCESSCM}
  else if FormuleMul0(Nom, Params, WhereSQL,TT,Total,S1) then
  begin
     Result := S1;
     { Objectifs }
     {$IF (Defined(GESCOM) AND Defined(BUSINESSPLACE)) OR Defined(PAIEGRH)}
     if (Result = '') and FormuleMulOBJPREV(Nom, Params, WhereSQL, TT, Total, S1) then
       Result := S1
     {$IFEND (Defined(GESCOM) AND Defined(BUSINESSPLACE)) OR Defined(PAIEGRH)}
  end
{$ENDIF ACCESSCM}
{ Objectifs }
{$IFNDEF PGISIDE}
{$IF (Defined(GESCOM) AND Defined(BUSINESSPLACE)) OR Defined(PAIEGRH)}
  else if FormuleMulOBJPREV(Nom, Params, WhereSQL, TT, Total, S1) then
     Result := S1
{$IFEND (Defined(GESCOM) AND Defined(BUSINESSPLACE)) OR Defined(PAIEGRH)}
{$ENDIF !PGISIDE}
  else Result:=DefProcCalcMul(Nom,Params,WhereSQL,TT,Total);
end;

end.
