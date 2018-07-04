{***********UNITE*************************************************
Auteur  ...... : Thibaut SUBLET
Créé le ...... : 16/04/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
unit CalcMulGPAO;

interface

uses
  {$IFDEF EAGLCLIENT}
    utob,
    Maineagl,
  {$ELSE}
    {$IFNDEF DBXPRESS}dbtables{BDE},{$ELSE}uDbxDataSet,{$ENDIF}
    uTob,
    DB,
    UEdtComp,
  {$ENDIF}
  Variants,
  HCtrls,
  wCommuns,
  Math,
  EntGC,
  {$IFDEF GPAO}
  {$IFDEF AFFAIRE}
  AffaireUtil,
  {$ENDIF AFFAIRE}
  TiersUtil,
  DispoDetail,
  Dispo,
  MeaUtil,
  EntGP,
  {$ENDIF GPAO}
  sysutils
  ;

{$IFDEF EAGLCLIENT}
function wProcCalcMul(Nom, Params, Where: string; TT: TOB; Total: boolean): string ;
{$ELSE}
function wProcCalcMul(Nom, Params, Where: string; TT: TDataSet; Total: boolean): string ;
{$ENDIF}

implementation

uses
  {$IFDEF GRC}
  {$ENDIF GRC}
  {$IFDEF GPAO}
  	wOrdreLig,
    wOrdreBes,
    wOrdrePhase,
  {$ENDIF GPAO}
  {$IFDEF AFFAIRE}
  {$ENDIF AFFAIRE}
  hEnt1,
  hMsgBox
  ;

{$IFDEF EAGLCLIENT}
function wProcCalcMul(Nom, Params, Where: string; TT: TOB; Total: boolean): string ;
{$ELSE}
function wProcCalcMul(Nom, Params, Where: string; TT: TDataSet; Total: boolean): string ;
{$ENDIF}
var
  {$IFDEF GPAO}
  Physique, QBesLance   : Double;
  Etatlig, QReste 			: string;
  MyArgument, Depot			: string;
  NumImage : integer;
  Tablette : string;
  ValChamp : variant;
  {$ENDIF GPAO}
  Prefixe, ParamsSave   : string;
  s1, s2, s3, s4, s5, s6: string;

  function GetCleGQ: tCleGQ;
  begin
    if Prefixe = 'WOB' then
    begin
      Result.Article := string(TT.FindField('WOB_COMPOSANT').Value);
      Result.Depot := string(TT.FindField('WOB_DEPOT').Value);
    end;
  end;
  function CtrlUnite ( Chp : hstring) : hstring;
  begin
    if copy(chp,4,1)='_' then
      Result := TT.FindField(chp).AsString
    else
      Result := Chp;
  end;

begin
  { pour rappeler DefProcCalcMul à la fin de cette fonction }
  ParamsSave := Params;

  s1 := ReadTokenPipe(ParamsSave, ';');
  s2 := ReadTokenPipe(ParamsSave, ';');
  s3 := ReadTokenPipe(ParamsSave, ';');
  s4 := ReadTokenPipe(ParamsSave, ';');
  s5 := ReadTokenPipe(ParamsSave, ';');
  s6 := ReadTokenPipe(ParamsSave, ';');

  Prefixe := wExtractGuillemet(s1);
  if Nom = 'SETICONE' then
  begin
    Result := '#ICO#' + intToStr(trunc(Random*25));
  end
  else if Nom = 'GPICONEANNULEE' then
  begin
    try
      Result := iif((TT.FindField('WJA_ANNULEE')<>nil) and (TT.FindField('WJA_ANNULEE').Value = wTrue), '#ICO#' + '80', '');
    except
    end;
  end
  else if Nom = 'GPICONEEXCEPTION' then
  begin
    try
      if (TT.FindField(Prefixe + '_CIRCUIT')=nil) or (TT.FindField(Prefixe + '_CIRCUIT').Value = '') then
        Result := ''
      else
        Result := '#ICO#' + '57';
    except
    end;
  end
  else if Nom = 'GPICONERESSOURCE' then
  begin
    try
      Result := iif((TT.FindField('C0')<>nil) and (TT.FindField('C0').Value > 0), '#ICO#' + '81', '');
    except
    end;
  end
{$IFDEF GPAO}
  else if Nom = 'GPICONETYPELIGNEPDR' then
  begin
    try
      if Copy(TT.Fields[0].FieldName,1,3)='QWL' then
        ValChamp := TT.FindField( 'QWL_TYPELIGNEPDR' ).Value
      else if Copy(TT.Fields[0].FieldName,1,3)='QWH' then
        ValChamp := TT.FindField( 'QWH_TYPELIGNEPDR' ).Value;
      Tablette :='QMESLIGNERELEVE';
      NumImage :=  GetNumImage( Tablette , TTToNum( Tablette ), ValChamp);
      Result := '#ICO#' + IntToStr(NumImage);
    except
    end;
  end
  else if Nom = 'GPICONEPHASE' then
  begin
    {Geré par le source FSL}
  end
  else if Nom = 'GPCOLOR' then
  begin
    try
      if (TT.Fields[0]<>nil) and (Copy(TT.Fields[0].FieldName,1,3)='QPH') then
      begin
        ValChamp := 0;
        if TT.FindField( 'QPH_PHASE' )<>nil then
          ValChamp := RechDom('QUTPHASECOUL', TT.FindField( 'QPH_PHASE' ).Value , False);
        if ValChamp>0 then
          Result := '#COL#' + IntToStr(ValChamp)
        else
          Result := '';
      end;
    except
    end;
  end
{$ENDIF GPAO}
  else if Nom = 'GPICONENOMENCLATURE' then
  begin
    try
      Result := iif(TT.FindField('C0').Value > 0, '#ICO#' + '55', '');
    except
      if V_Pgi.SAV then PGIError('Vérifier la liste', 'Mode Sav');
    end;
  end
  else if Nom = 'GPICONEFAISAMANQUE' then
  begin
    try
      Result := iif(TT.FindField('WND_VALLIBRE6').Value > 0, '#ICO#' + '48', '');
    except
    end;
  end
  {$IFDEF GPAO}
  else if Nom = 'GPFORMULERESTE' then   // pour appliquer une formule s'il ne s'agit pas d'un MUL
  begin
    try
      if Prefixe = 'WOB' then
      begin
        EtatLig   := TT.FindField(Prefixe + '_ETATBES').Value;
        QBesLance := TT.FindField('WOB_QBESSAIS').AsFloat * wDivise(wGetFieldFromWOP('WOP_QLANSAIS', GetCleWOP(TT.FindField('WOB_NATURETRAVAIL').AsString, TT.FindField('WOB_LIGNEORDRE').AsInteger, TT.FindField('WOB_OPECIRC').AsString)), wGetFieldFromWOP('WOP_QACCSAIS', GetCleWOP(TT.FindField('WOB_NATURETRAVAIL').AsString, TT.FindField('WOB_LIGNEORDRE').AsInteger, TT.FindField('WOB_OPECIRC').AsString)));
        QReste    := FloatToStr(Max(0, QBesLance - TT.FindField('WOB_QCONSAIS').AsFloat));
      end
      else QReste :='0';
      if (EtatLig = 'SOL') or (EtatLig = 'TER') then Result := '0'
      else Result := Qreste;
    except
    end;
  end
  {$ENDIF GPAO}
  {$IFDEF GPAO}
  else if Nom = 'GPFORMULEQSTO' then   // afficher le stock physique
  begin
    try
      if Prefixe = 'WOB' then
      begin
				{ Dans le cas de la liste à servir de production }
  	  	MyArgument:= 'QUALIFMVT=CAF'
                  + ';STATUTFLUX=-PVE'
                  + ';REFAFFECTATION='     + GetRefAffectationWOB(GetCleWOB(TT.FindField('WOB_NATURETRAVAIL').AsString,TT.FindField('WOB_LIGNEORDRE').AsInteger,TT.FindField('WOB_OPECIRC').AsString,TT.FindField('WOB_LIENNOME').AsInteger))
                  + ';REFAFFECTATION_WOL=' + GetRefAffectationWOL(GetCleWOL(GetCleWOB(TT.FindField('WOB_NATURETRAVAIL').AsString,TT.FindField('WOB_LIGNEORDRE').AsInteger,TT.FindField('WOB_OPECIRC').AsString,TT.FindField('WOB_LIENNOME').AsInteger)))
                  + ';REFAFFECTATION_WOP=' + GetRefAffectationWOP(GetCleWOP(GetCleWOB(TT.FindField('WOB_NATURETRAVAIL').AsString,TT.FindField('WOB_LIGNEORDRE').AsInteger,TT.FindField('WOB_OPECIRC').AsString,TT.FindField('WOB_LIENNOME').AsInteger)))
                  + ';REFAFFECTATION_T='   + GetRefAffectationT(wGetFieldFromWOL('WOL_TIERS'    , GetCleWOL(GetCleWOB(TT.FindField('WOB_NATURETRAVAIL').AsString,TT.FindField('WOB_LIGNEORDRE').AsInteger,TT.FindField('WOB_OPECIRC').AsString,TT.FindField('WOB_LIENNOME').AsInteger))))
			{$IFDEF AFFAIRE}
  	  		        + ';REFAFFECTATION_AFF=' + GetRefAffectationAFF(wGetFieldFromWOL('WOL_AFFAIRE', GetCleWOL(GetCleWOB(TT.FindField('WOB_NATURETRAVAIL').AsString,TT.FindField('WOB_LIGNEORDRE').AsInteger,TT.FindField('WOB_OPECIRC').AsString,TT.FindField('WOB_LIENNOME').AsInteger))))
			{$ENDIF AFFAIRE}
									;
				if (TT.FindField('ULS_DEPOTORG')<>nil) and (TT.FindField('ULS_DEPOTORG').AsString<>'')then
          Depot:= TT.FindField('ULS_DEPOTORG').AsString
        else
        	Depot:= TT.FindField('WOB_DEPOT').AsString;
				Physique := GetQDispo(Dispo.GetCleGQ(Depot,String(TT.FindField('WOB_COMPOSANT').Value)),
        						iif((TT.FindField('WOB_CODELAS').AsInteger<>0) or (TT.FindField('WOB_CODELASST').AsInteger<>0),MyArgument,'STATUTDISPO=LBR'), 'STO');

        Physique := wDivise(Physique, TT.FindField('WOB_COEFLIEN').AsFloat); { Unité Stock -> Unité de Qte de lien }
        Result := FloatToStr(Physique);
      end
      else Result :='0';
    except
    end;
  end
  else if Nom = 'GPCONVHHCC' then // Conversion Temps
  begin
    try
      ValChamp := TT.FindField( s1 ).Value;
      if ValChamp >0 then
      begin
        Result := FloatToStr( wConversionTemps (ValChamp ,'HC' , GetParamMESTpsUnitDef));
      end
      else
        Result := '0';
    except
    end;
  end
  else if Nom = 'GPCONVHHCCUNIT' then   // pour convertir le temps de HHCC dans l'unité d'affichage
    Result := FloatToStr(wConversionTemps(TT.FindField(s1).AsFloat, 'HC', CtrlUnite(s2) ))
  else if Nom = 'GPCONVTPS' then   // pour convertir un temps
    Result := FloatToStr(wConversionTemps(TT.FindField(s1).AsFloat, CtrlUnite(s2) , CtrlUnite(s3) ))
  else if Nom = 'GPCONVTPSSYN' then   // pour convertir un temps dans l'unité de synthése
    Result := FloatToStr(wConversionTemps(TT.FindField(s1).AsFloat, 'HC' , GetParamMESTpsUnitSyn))
  else if Nom = 'GPHEURE' then   // Extraction de l'heure dans une date
  begin
    ValChamp := TT.FindField( s1 ).AsFloat;
    if (valChamp >1000) or (valChamp - Abs(ValChamp)>0 ) then
      Result := TimeToStr(ValChamp);
  end
  {$ENDIF GPAO}
  ;
end;

end.
