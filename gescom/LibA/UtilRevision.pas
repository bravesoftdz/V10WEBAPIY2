unit UtilRevision;

interface

uses
  HEnt1, HCtrls,
  {$IFNDEF EAGLCLIENT}
  {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
  EdtREtat,
  {$ELSE}
  UtileAGL,
  {$ENDIF}
  UTob, Ed_tools,
  Controls,
  Classes,
  dicoBTP, Sysutils, paramsoc, EntGC,
  forms;

function VerifieCoherence(Affaire: string; pBoSilence: Boolean): integer;
procedure ValeurIndiceRegularisation(MaTob: Tob; pDate: TDateTime; pStIndice, pStPubCode: string; var pDtIndice: TDateTime; var pDbValeur: Double);
procedure FlagueLesLignesRegularisables(Ensilence: boolean);
procedure LanceEtatIndicesNonSaisis;
procedure LanceEtatDernieresValeursIndicesSaisies; 
function MaxRevision(pStAffaire: string): Integer;
function ToutesFormulesParam(pStAffaire: string): Boolean;
function MajDatesRevisionsApresDesapplication(pStAffaire, pStFormule: string; pDtDesapplication: TDateTime): Boolean;
  
procedure DebutLog(pStFileName: string; var pSLLog: TStringList);
procedure FinLog(pStFileName: string; var pSLLog: TStringList);

const TexteMessage: array[1..3] of string = (
    {1}   'Traitement terminé.',
    {2}   'Marquage des lignes régularisables.',                                               
    {3}   '%s Affaire %s. La facture n° %s n''est pas régularisable, car l''indice %s du mois de %s n''est pas connu.'
    );

implementation

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 22/04/2003
Modifié le ... :
Description .. : la cherche des indices de regularisation est faites
                 sur la valeur du mois de l'indice
                 retourne -1 si pas de valeur trouvée
Mots clefs ... :
*****************************************************************}

function Existe(T: Tob; Value, Champ: string): Boolean;
var i: integer;
begin
  i := 0;
  if value = '' then
    result := True
  else
  begin
    result := False;
    while (i < T.detail.count) and (not result) do
    begin
      result := (T.detail[0].GetValue(champ) = Value);
      inc(i);
    end;
  end;
end;

procedure fabriqueTobformuleAffaire(TobForAff: tob; Affaire: string);
var TobFilleForAff: Tob;
  st: string;
  Q: Tquery;
begin

  st := 'select gl_forcode1  forcode from ligne  where GL_AFFAIRE="' + Affaire + '"  union ';
  st := st + 'select gl_forcode2  forcode from ligne  where GL_AFFAIRE="' + Affaire + '" union ';
  st := st + 'select aff_forcode1 forcode from Affaire where Aff_AFFAIRE="' + Affaire + '"union ';
  st := st + 'select aff_forcode2 forcode from Affaire where Aff_AFFAIRE="' + Affaire + '" ';
  Q := nil;
  Q := OpenSQL(St, TRUE);
  while not Q.eof do
  begin
    if not Existe(TobForAff, Q.FindField('forcode').asstring, 'forcode') then
    begin
      TobFilleForAff := TOB.Create('Mes Formules dans les affaire', TobForAff, -1);
      TobFilleForAff.AddChampSup('forcode', true);
      TobFilleForAff.AddChampSup('Action', true);
      TobFilleForAff.putValue('forcode', Q.FindField('forcode').asstring);
      TobFilleForAff.putValue('Action', 'Creation');
    end;
    Q.Next;
  end;
  ferme(Q);
end;

procedure fabriqueTobformuleParamFor(uneTob: tob; Affaire: string);
var st: string;
  Q: Tquery;
  TobFille: Tob;
begin
  st := 'select afc_forcode from afparamformule  where Afc_AFFAIRE="' + Affaire + '" ';
  Q := nil;
  Q := OpenSQL(St, TRUE);
  while not Q.eof do
  begin
    TobFille := TOB.Create('Mes Formules dans les parametres', unetob, -1);
    TobFille.AddChampSup('forcode', true);
    TobFille.AddChampSup('Action', true);
    TobFille.putValue('forcode', Q.FindField('afc_forcode').asstring);
    TobFille.putValue('Action', 'Suppression');
    Q.next;
  end;
  ferme(Q);
end;

function ExisteAndFlag(TobForParam: Tob; Value, Champ: string): Boolean;
var i: integer;
begin
  i := 0;
  result := False;
  while (i < TobForParam.detail.count) and (not result) do
  begin
    if (TobForParam.detail[i].GetValue(champ) = Value) then
    begin
      result := true;
      TobForParam.detail[i].putValue('Action', 'Ok');
    end;
    inc(i);
  end;
end;

function ComparelesTobs(TobForAff, TobForParam: tob): boolean;
var i: integer;
begin
  result := True;
  try
    for i := 0 to TobForAff.detail.count - 1 do
    begin
      if ExisteAndFlag(TobForParam, tobForAff.detail[i].getvalue('forcode'), 'forcode') then
        tobForAff.detail[i].putvalue('Action', 'Ok');
    end;
  except
    result := false;
  end;
end;

procedure CrelesManquant(TobForAff: tob; Affaire: string);
var i, j: integer;
  stSql: string;
begin
  for i := 0 to TobForAff.detail.count - 1 do
  begin
    if tobForAff.detail[i].getvalue('Action') = 'Creation' then
    begin
      stSql := 'insert into  afparamformule (afc_forcode,afc_affaire,AFC_APPLIQUERCOEF,AFC_MODELECTURE,AFC_LASTDATECALC, ';
      stSql := stSql + ' AFC_NEXTDATECALC,AFC_NEXTDATEAPP,AFC_LASTDATEAPP,AFC_DATEINITIALE,AFC_PREMIEREDATE, ';
      stSql := stSql + ' AFC_SEUILMIN,AFC_SEUILMAX,AFC_PERIODREV,AFC_PARAMFORMULEOK,AFC_REVFINMOIS,AFC_NUMMOISREV,AFC_NUMJOURREV,AFC_PARAMDESC,';
      for j := 1 to 10 do
      begin
        stSql := stSql + 'AFC_VALINITIND' + inttostr(j) + ',AFC_INDAFF' + inttostr(j) + ',AFC_PUBCODE' + inttostr(j);
        if j < 10 then stSql := stSql + ',';
      end;
      stSql := stSql + ' ) values ("' + tobForAff.detail[i].getvalue('forcode') + '","' + Affaire + '","-","MFA",';
      stSql := stSql + '"' + usdatetime(iDate2099) + '","' + usdatetime(iDate2099) + '","' + usdatetime(iDate2099) + '","' + usdatetime(iDate2099) + '","' +
        usdatetime(iDate2099) + '","' + usdatetime(iDate2099) + '",';
      stSql := stSql + '0,0,"T","-","-",1,1,"",';
      for j := 1 to 10 do
      begin
        stSql := stSql + '0,"",""';
        if j < 10 then stSql := stSql + ',';
      end;
      stSql := stSql + ')';
      executesql(stSql);
    end;
  end;
end;

function SupprimeLesInutiles(TobForParam: tob; Affaire: string; pBoSilence: Boolean): boolean;
var i: integer;
  st: string;
begin
  result := true;
  try
    if TobForParam.detail <> nil then
      for i := 0 to TobForParam.detail.count - 1 do // GM le 7/7/03
      begin
        if TobForParam.detail[i].getvalue('Action') = 'Suppression' then
          if pBoSilence then
          begin
            st := 'delete from  afparamformule where afrc_forcode = "' + TobForParam.detail[i].getvalue('forcode') + '" and afc_affaire ="' + Affaire + '"';
            executesql(st);
            st := 'delete from  afrevision where afr_forcode = "' + TobForParam.detail[i].getvalue('forcode') + '" and afr_affaire ="' + Affaire + '"';
            executesql(st);
          end
          else if PGIAskAF('Suppression de la formule ' + TobForParam.detail[i].getvalue('forcode') + ' pour cette affaire ?', '') = mrYes then
          begin
            st := 'delete from  afparamformule where afc_forcode = "' + TobForParam.detail[i].getvalue('forcode') + '" and afc_affaire ="' + Affaire + '"';
            executesql(st);
            st := 'delete from  afrevision where afr_forcode = "' + TobForParam.detail[i].getvalue('forcode') + '" and afr_affaire ="' + Affaire + '"';
            executesql(st);
          end;
      end;
  except
    result := false;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : PB
Créé le ...... : 23/04/2003
Modifié le ... :   /  /
Description .. : Vérification des paramétrages des formules
Mots clefs ... :
*****************************************************************}

function VerifieCoherence(Affaire: string; pBoSilence: Boolean): integer;
var TobForAff, TobForParam: Tob;
begin
  (* but des opérations
  une tob pour les lignes et les affaires va etre créer.
  elle sera faite a la main pour eviter les doublons.
  Une tob pour les parametres de formule va etre créer.
  Ensuite on va comparer ces deux tobs.
  si une formule existe dans les 2 tobs => Ok
  Si une formule existe dans TobForAff et pas dans TobForParam
  =>il faut creer un enregistrement dans ParamFormule
  Si une formule n'existe pas dans TobForAff et existe dans TobForParam
  =>il faut supprimmer l'enregistrement dans ParamFormule apres confirmation.
  *)
  TobForAff := TOB.Create('Mes Formules dans les affaires', nil, -1);
  TobForParam := TOB.Create('Mes formules dans les paramsformules', nil, -1);
  fabriqueTobformuleAffaire(TobForAff, Affaire);
  fabriqueTobformuleParamFor(TobForParam, Affaire);
  if CompareLesTobs(TobForAff, TobForParam) then
    CrelesManquant(TobForAff, Affaire);
  SupprimeLesInutiles(TobForParam, Affaire, pBoSilence);
  TobForAff.free;
  // je recompte les AfrevParamformules pour voir si je dois envoyer la suite ou pas
  TobForAff := TOB.Create('Mes Formules dans les affaires', nil, -1);
  fabriqueTobformuleAffaire(TobForAff, Affaire);
  result := TobForAff.Detail.count;
  TobForAff.free;
  TobForParam.free;
end;

procedure ValeurIndiceRegularisation(MaTob: Tob; pDate: TDateTime; pStIndice, pStPubCode: string; var pDtIndice: TDateTime; var pDbValeur: Double);
var
  vDtDebut    : TDateTime;
  vDtFin      : TDateTime;
  vTob        : Tob;

begin

  vDtDebut := DebutdeMois(pDate);
  vDtFin := FinDeMois(pDate);
 
  if pStPubCode = '' then
    vTob := maTob.FindFirst(['AFV_INDCODE', 'AFV_DEFINITIF'],
                            [pStIndice, 'X'], True)
  else
    vTob := maTob.FindFirst(['AFV_INDCODE', 'AFV_PUBCODE', 'AFV_DEFINITIF'],
                            [pStIndice, pStPubCode, 'X'], True);

  while vTob <> nil do
  begin
    if (vDtDebut <= vTob.GetValue('AFV_INDDATEVAL')) and
        (vDtFin >= vTob.GetValue('AFV_INDDATEVAL')) then
      begin
        if (vTob.GetValue('AFV_INDCODESUIV') = '') or
           (pdate < vTob.GetValue('AFV_INDDATEFIN')) then
        begin
          pDbValeur := vTob.GetValue('AFV_INDVALEUR');
          pDtIndice := vTob.GetValue('AFV_INDDATEVAL');
          if vTob.GetValue('AFV_COEFRACCORD') <> 0 then
            pDbValeur := pDbValeur * vTob.GetValue('AFV_COEFRACCORD');
          break;
        end
        else
        begin
          ValeurIndiceRegularisation(MaTob, pdate,
                                     vTob.GetValue('AFV_INDCODESUIV'), '',
                                     pDtIndice, pDbValeur);
        end;
      end
    else
    begin
      if pStPubCode = '' then
        vTob := maTob.FindNext(['AFV_INDCODE', 'AFV_DEFINITIF'],
                               [pStIndice, 'X'], True)
      else
        vTob := maTob.FindNext(['AFV_INDCODE', 'AFV_PUBCODE', 'AFV_DEFINITIF'],
                               [pStIndice, pStPubCode, 'X'], True);
    end;
  end;
end;

function StringGereNull(V :Variant) : String ;
begin
  try
  result:=V ;
  except
  result:='' ;
  end ;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : PB
Créé le ...... : 10/06/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure ParcoursEtFlague(TobLigne, TobValindice: tob);
var
  i, j            : Integer;
  StSql           : String;
  LeIndice        : String;
  ValeurIndice    : double;
  LaDate          : TDateTime;
  DateValeur      : TDateTime;
  OLD_GL_DATEPIECE: TdateTime;
  OLD_GL_SOUCHE   : string;
  OLD_GL_INDICEG  : integer;
  OLD_GL_NUMERO   : integer;
  Actualisable    : Boolean;
  MemePiece       : Boolean;
  PieceEntiere    : boolean;
  vSLLog          : TStringList;
  Year, Month, Day: Word;

  function StringGereNull(V :Variant) : String ;
  begin
    try
      result:=V;
    except
      result:='';
    end ;
  end;

begin
  DebutLog('', vSLLog);
  InitMoveProgressForm(nil, 'Traitement en cours...', '', TobLigne.detail.count, false, true);
  i := 0;
  PieceEntiere := GetParamSoc('SO_AFREGULGLOBAL');

  OLD_GL_DATEPIECE  := 0;
  OLD_GL_SOUCHE     := '';
  OLD_GL_INDICEG    := 0;
  OLD_GL_NUMERO   := 0;
  Actualisable := True;
  while i < TobLigne.detail.count do
  begin

    MemePiece := (OLD_GL_DATEPIECE = TobLigne.detail[i].getvalue('GL_DATEPIECE')) and
      (OLD_GL_SOUCHE = TobLigne.detail[i].getvalue('GL_SOUCHE')) and
      (OLD_GL_INDICEG = TobLigne.detail[i].getvalue('GL_INDICEG')) and
      (OLD_GL_NUMERO = TobLigne.detail[i].getvalue('GL_NUMERO'));

    if not MemePiece then Actualisable := True; // je reinitialise on change de piece
    if Actualisable then
    begin
      for j := 1 to 20 do
      begin
        LeIndice  := StringGereNull(TobLigne.detail[i].getvalue('IND' + inttostr(j)));
        LaDate    := StrToDate(StringGereNull(TobLigne.detail[i].getvalue('GL_DATEPIECE')));
        decodeDate(LaDate, year, month, day);
        if LeIndice <> '' then
        begin
          ValeurIndice := -1;
          ValeurIndiceRegularisation(TobValindice, LaDate, LeIndice, '', DateValeur, ValeurIndice);
          if ValeurIndice = -1 then
          begin
            Actualisable := False;
            vSLLog.add(format(TexteMessage[3],
                              [dateToStr(now),
                              TobLigne.detail[i].getvalue('GL_AFFAIRE'),
                              TobLigne.detail[i].getvalue('GL_NUMERO'),
                              LeIndice,
                              intToStr(month) + '/' + intToStr(year)]));
            break;
          end;
        end;
      end;
    end;
    if Actualisable then
    begin // juste la ligne
      stSql := 'Update LIGNE SET GL_REGULARISABLE="X" where ';
      stSql := stSql + 'GL_DATEPIECE="' + usdatetime(TobLigne.detail[i].getvalue('GL_DATEPIECE')) + '"';
      stSql := stSql + 'AND GL_SOUCHE="' + TobLigne.detail[i].getvalue('GL_SOUCHE') + '"';
      stSql := stSql + 'AND GL_INDICEG=' + inttostr(TobLigne.detail[i].getvalue('GL_INDICEG'));
      stSql := stSql + 'AND GL_NUMERO=' + inttostr(TobLigne.detail[i].getvalue('GL_NUMERO'));
      stSql := stSql + 'AND GL_NUMLIGNE=' + inttostr(TobLigne.detail[i].getvalue('GL_NUMLIGNE'));
    end
    else
    begin
      stSql := 'Update LIGNE SET GL_REGULARISABLE="-" where ';
      stSql := stSql + 'GL_DATEPIECE="' + usdatetime(TobLigne.detail[i].getvalue('GL_DATEPIECE')) + '"';
      stSql := stSql + 'AND GL_SOUCHE="' + TobLigne.detail[i].getvalue('GL_SOUCHE') + '"';
      stSql := stSql + 'AND GL_INDICEG=' + inttostr(TobLigne.detail[i].getvalue('GL_INDICEG'));
      stSql := stSql + 'AND GL_NUMERO=' + inttostr(TobLigne.detail[i].getvalue('GL_NUMERO'));
      if not PieceEntiere then // toute la piece ou seulement cette ligne
        stSql := stSql + 'AND GL_NUMLIGNE=' + inttostr(TobLigne.detail[i].getvalue('GL_NUMLIGNE'));
    end;
    executesql(stSql);

    // si global et non actualisable tant que meme piece on boucle
    if PieceEntiere and (not Actualisable) then
    begin
      OLD_GL_DATEPIECE := TobLigne.detail[i].getvalue('GL_DATEPIECE');
      OLD_GL_SOUCHE := TobLigne.detail[i].getvalue('GL_SOUCHE');
      OLD_GL_INDICEG := TobLigne.detail[i].getvalue('GL_INDICEG');
      OLD_GL_NUMERO := TobLigne.detail[i].getvalue('GL_NUMERO');
      inc(i);
           
      while (i < TobLigne.detail.count) and
            (OLD_GL_DATEPIECE = TobLigne.detail[i].getvalue('GL_DATEPIECE')) and
            (OLD_GL_SOUCHE = TobLigne.detail[i].getvalue('GL_SOUCHE')) and
            (OLD_GL_INDICEG = TobLigne.detail[i].getvalue('GL_INDICEG')) and
            (OLD_GL_NUMERO = TobLigne.detail[i].getvalue('GL_NUMERO')) do
      begin                                                             
        inc(i);
        MoveCurProgressForm('');
      end;
    end
    else
    begin
      OLD_GL_DATEPIECE := TobLigne.detail[i].getvalue('GL_DATEPIECE');
      OLD_GL_SOUCHE := TobLigne.detail[i].getvalue('GL_SOUCHE');
      OLD_GL_INDICEG := TobLigne.detail[i].getvalue('GL_INDICEG');
      OLD_GL_NUMERO := TobLigne.detail[i].getvalue('GL_NUMERO');
      inc(i);
      MoveCurProgressForm('');
    end;
  end;
  FinLog('', vSLLog);
  FiniMoveProgressForm;
end;

{***********A.G.L.***********************************************
Auteur  ...... : PB
Créé le ...... : 02/07/2003
Modifié le ... :   /  /
Description .. : dans ce traitement est pris en compte le paramsoc SO_REGULARISATION
                 si les pieces doivent etre régularisées "entieres" alors il suffit de trouver
                 une ligne non régularisable pour flaguer toutes les autres lignes de la piece a '-'.
                 Lors de la sélection qui suit (pour selection des factures regularisables),
                 SO_REGULARISATION n'entre plus en compte car dans le cas des pieces régularisées "entieres"
                 aucune d'elles ne peut etre mixte (du fait du flagage "PieceEntiere")
Mots clefs ... :
*****************************************************************}
procedure FlagueLesLignesRegularisables(Ensilence: boolean);
var
  TobValindice  : Tob;
  TobLigne      : Tob;
  Q             : Tquery;
  Qligne        : Tquery;
  vSt           : String;
  vStLig        : String;

begin

  TobValindice := TOB.Create('Mes Valeurs indices', nil, -1);
  TobLigne := TOB.Create('Mes lignes', nil, -1);
  
  Try
    vSt := 'Select AFV_INDCOMMENT,AFV_PUBCODE,AFV_INDCODE,AFV_INDDATEVAL, ';
    vSt := vSt + ' AFV_DEFINITIF,AFV_INDVALEUR,AFV_INDDATEFIN,AFV_INDCODESUIV, ';
    vSt := vSt + ' AFV_COEFPASSAGE,AFV_PUBCODESUIV,AFV_COEFRACCORD ';
    vSt := vSt + ' FROM AFVALINDICE order by AFV_INDDATEVAL desc';
    Q := nil;
    try
      Q := OpenSQL(vSt, TRUE);
      TobValindice.LoadDetailDB('', '', '', Q, false);
    finally
      Ferme(Q);
    end;

    vStLig := 'SELECT GL_AFFAIRE,GL_DATEPIECE,GL_REGULARISABLE,GL_NATUREPIECEG, ';
    vStLig := vStLig + 'GL_SOUCHE, GL_NUMERO, GL_INDICEG, GL_NUMLIGNE, ';
    vStLig := vStLig + 'F2.AFE_INDCODE10 IND20, F2.AFE_INDCODE9  IND19,';
    vStLig := vStLig + 'F2.AFE_INDCODE8  IND18, F2.AFE_INDCODE7  IND17,';
    vStLig := vStLig + 'F2.AFE_INDCODE6  IND16, F2.AFE_INDCODE5  IND15,';
    vStLig := vStLig + 'F2.AFE_INDCODE4  IND14, F2.AFE_INDCODE3  IND13,';
    vStLig := vStLig + 'F2.AFE_INDCODE2  IND12, F2.AFE_INDCODE1  IND11,';
    vStLig := vStLig + 'F1.AFE_INDCODE10 IND10, F1.AFE_INDCODE9  IND9,';
    vStLig := vStLig + 'F1.AFE_INDCODE8  IND8,  F1.AFE_INDCODE7  IND7,';
    vStLig := vStLig + 'F1.AFE_INDCODE6  IND6,  F1.AFE_INDCODE5  IND5,';
    vStLig := vStLig + 'F1.AFE_INDCODE4  IND4,  F1.AFE_INDCODE3  IND3,';
    vStLig := vStLig + 'F1.AFE_INDCODE2  IND2,  F1.AFE_INDCODE1  IND1 ';
    vStLig := vStLig + ' FROM LIGNE LEFT JOIN AFFORMULE F1  ON LIGNE.GL_FORCODE1 = F1.AFE_FORCODE ';
    vStLig := vStLig + ' LEFT JOIN AFFORMULE F2  ON LIGNE.GL_FORCODE2 = F2.AFE_FORCODE ';
    vStLig := vStLig + ' where GL_REGULARISE="-" and GL_REGULARISABLE="-" ';
    vStLig := vStLig + ' and ((gl_forcode1<>"") or (gl_forcode2<>""))';
    vStLig := vStLig + ' and (GL_NATUREPIECEG = "FAC")';

    Qligne := nil;
    try
      Qligne := OpenSQL(vStLig, TRUE);
      TobLigne.LoadDetailDB('', '', '', Qligne, false);
    finally
      Ferme(Qligne);
    end;
    ParcoursEtFlague(TobLigne, TobValindice);
  finally
    TobValindice.free;
    TobLigne.free;
  end;

  if not Ensilence then PgiinfoAF(TexteMessage[1], TexteMessage[2]);
end;

procedure LanceEtatIndicesNonSaisis;
var
  StWhere     : string;
  DateLimite  : TdateTime;
begin
  DateLimite := plusdate(date, -1, 'A');
  StWhere := ' ((AFV_DEFINITIF="X") AND (((AFP_PUBPERIOD="A") AND (LADATE < "' + UsDateTime(DateLimite) + '")) ';
  DateLimite := plusdate(date, -1, 'S');
  StWhere := StWhere + ' OR ((AFP_PUBPERIOD="H") AND (LADATE < "' + UsDateTime(DateLimite) + '")) ';
  DateLimite := plusMois(date, -1);
  StWhere := StWhere + ' OR ((AFP_PUBPERIOD="M") AND (LADATE < "' + UsDateTime(DateLimite) + '")) ';
  DateLimite := plusMois(date, -6);
  StWhere := StWhere + ' OR ((AFP_PUBPERIOD="S") AND (LADATE < "' + UsDateTime(DateLimite) + '")) ';
  DateLimite := plusMois(date, -3);
  StWhere := StWhere + ' OR ((AFP_PUBPERIOD="T") AND (LADATE < "' + UsDateTime(DateLimite) + '")))) ';
//  StWhere := StWhere + 'OR ISNULL(LADATE,"12/31/2099")="12/31/2099"';
  StWhere := StWhere + 'OR LADATE IS NULL';
  LanceEtat('E', 'AIN', 'AIN', True, False, False, nil, StWhere, '', False);
end;
 
procedure LanceEtatDernieresValeursIndicesSaisies;
begin
  LanceEtat('E', 'AIV', 'AIV', True, False, False, nil, '', '', False);
end;

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 10/04/2003
Modifié le ... :   /  /
Description .. : le max est toute formule confondue
Mots clefs ... :
*****************************************************************}

function MaxRevision(pStAffaire: string): Integer;
var
  vSt: string;
  vQr: TQuery;

begin
  vSt := 'SELECT MAX(AFR_NUMEROLIGNE) AS NUM FROM AFREVISION ';
  vSt := vSt + 'WHERE AFR_AFFAIRE = "' + pStAffaire + '"';

  vQr := nil;
  try
    vQR := OpenSql(vSt, True);
    if not vQR.Eof then
      result := vQr.FindField('NUM').AsInteger + 1
    else
      result := 1;
  finally
    if vQR <> nil then ferme(vQR);
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 10/07/2003
Modifié le ... :   /  /
Description .. : Recherche si une formule n'est pas paramétrée
Mots clefs ... :
*****************************************************************}

function ToutesFormulesParam(pStAffaire: string): Boolean;
var
  i: Integer;
  vSt: string;
  vSt1: string;
  vQr: TQuery;
  vQr1: TQuery;
  vQr2: TQuery;
  vTob: Tob;
  vTob1: Tob;
  vTob2: Tob;
  vTob3: Tob;

begin

  result := True;
  vSt := 'SELECT DISTINCT GL_FORCODE1, GL_FORCODE2 FROM LIGNE ';
  vSt := vSt + 'WHERE GL_AFFAIRE ="' + pStAffaire + '"';

  vSt1 := 'SELECT DISTINCT AFF_FORCODE1, AFF_FORCODE2 ';
  vSt1 := vSt1 + ' FROM AFFAIRE WHERE AFF_AFFAIRE ="' + pStAffaire + '"';

  vQr := nil;
  vQr1 := nil;
  vQr2 := nil;

  vTob := Tob.create('formule ligne', nil, -1);
  vTob1 := Tob.create('les params', nil, -1);
  vTob2 := Tob.create('formule entete', nil, -1);

  try
    vQr := OpenSQL(vSt, True);
    vTob.LoadDetailDB('formule ligne', '', '', vQr, False, true);

    vQr1 := OpenSQL(vSt1, True);
    vTob1.LoadDetailDB('formule entete', '', '', vQr, False, true);

    if (not vQr.EOF) or (not vQr1.EOF) then
    begin

      try
        vSt := 'SELECT DISTINCT AFC_FORCODE FROM AFPARAMFORMULE ';
        vSt := vSt + 'WHERE AFC_AFFAIRE = "' + pStAffaire + '"';

        vQr2 := OpenSQL(vSt, True);
        vTob2.LoadDetailDB('les params', '', '', vQr2, false, true);

        for i := 0 to vTob.detail.count - 1 do
        begin
          if vTob.Detail[i].GetValue('GL_FORCODE1') <> '' then
          begin
            vTob3 := vTob2.FindFirst(['AFC_FORCODE'], [vTob.Detail[i].GetValue('GL_FORCODE1')], true);
            if vTob3 = nil then
            begin
              result := false;
              break;
            end
            else if vTob.Detail[i].GetValue('GL_FORCODE2') <> '' then
            begin
              vTob3 := vTob2.FindFirst(['AFC_FORCODE'], [vTob.Detail[i].GetValue('GL_FORCODE2')], true);
              if vTob3 = nil then
              begin
                result := false;
                break;
              end;
            end;
          end;
        end;

        if result then
          for i := 0 to vTob1.detail.count - 1 do
          begin
            if vTob1.Detail[i].GetValue('AFF_FORCODE1') <> '' then
            begin
              vTob3 := vTob2.FindFirst(['AFR_FORCODE'], [vTob1.Detail[i].GetValue('AFF_FORCODE1')], true);
              if vTob3 = nil then
              begin
                result := false;
                break;
              end
              else if vTob1.Detail[i].GetValue('AFF_FORCODE2') <> '' then
              begin
                vTob3 := vTob2.FindFirst(['AFR_FORCODE'], [vTob1.Detail[i].GetValue('AFF_FORCODE2')], true);
                if vTob3 = nil then
                begin
                  result := false;
                  break;
                end;
              end;
            end;
          end;
      finally
        Ferme(vQr2);
        vTob2.Free;
      end;
    end
  finally
    Ferme(vQr);
    vTob.Free;
    Ferme(vQr1);
    vTob1.Free;
  end;
end;

procedure DebutLog(pStFileName: string; var pSLLog: TStringList);
begin
  pSLLog := TStringList.create;
  try
    if pStFileName = '' then
    begin
      if getParamSoc('SO_AFREVPATH') = '' then
        pSLLog.LoadFromFile(ExtractFileDrive(Application.ExeName) + '\RevisionPrix.log')
      else
        pSLLog.LoadFromFile(getParamSoc('SO_AFREVPATH') + '\RevisionPrix.log');
    end
    else
      pSLLog.LoadFromFile(pStFileName);

  except
  end;
end;

procedure FinLog(pStFileName: string; var pSLLog: TStringList);
begin
  if pStFileName = '' then
  begin
    if getParamSoc('SO_AFREVPATH') = '' then
      pSLLog.SaveToFile(ExtractFileDrive(Application.ExeName) + '\RevisionPrix.log')
    else                                      
    begin
      if copy(getParamSoc('SO_AFREVPATH'), length(getParamSoc('SO_AFREVPATH')) - 1, 1) = '\' then
        pSLLog.SaveToFile(getParamSoc('SO_AFREVPATH') + 'RevisionPrix.log')
      else
        pSLLog.SaveToFile(getParamSoc('SO_AFREVPATH') + '\RevisionPrix.log');
    end
  end    
  else
    pSLLog.SaveToFile(pStFileName);

  pSLLog.Free;
end;

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 22/07/2003
Modifié le ... :   /  /
Description .. : Mise a jour des dates après désapplication des coefficients
Mots clefs ... :
*****************************************************************}

function MajDatesRevisionsApresDesapplication(pStAffaire, pStFormule: string; pDtDesapplication: TDateTime): Boolean;
var
  vSt: string;
  vQr: TQuery;
begin

  result := True;

  // tester si il y a un document de facturation
  // si oui, on ne peut pas désappliquer
  vSt := 'SELECT AFR_DATECALCCOEF FROM AFREVISION';
  vSt := vSt + ' WHERE AFR_AFFAIRE = "' + pStAffaire + '"';
  vSt := vSt + ' AND AFR_FORCODE = "' + pStFormule + '"';
  vSt := vSt + ' AND AFR_DATECALCCOEF = "' + usDateTime(pDtDesapplication) + '"';
  vSt := vSt + ' AND AFR_NATUREPIECEG <> "' + VH_GC.AFNatAffaire + '"';
  vSt := vSt + ' AND AFR_NATUREPIECEG <> "..."';

  vQr := nil;
  try
    vQr := OpenSQL(vSt, TRUE);
    if not vQr.eof then Result := False;
  finally
    Ferme(vQr);
  end;

  if Result then
  begin
    vSt := 'SELECT MAX(AFR_DATECALCCOEF) AS DATECALCCOEF FROM AFREVISION';
    vSt := vSt + ' WHERE AFR_AFFAIRE = "' + pStAffaire + '"';
    vSt := vSt + ' AND AFR_FORCODE = "' + pStFormule + '"';
    vSt := vSt + ' AND AFR_DATECALCCOEF < "' + usDateTime(pDtDesapplication) + '"';

    vQr := nil;
    try
      vQr := OpenSQL(vSt, TRUE);
      if vQr.FindField('DATECALCCOEF').AsString = '' then
      begin
        vSt := 'UPDATE AFPARAMFORMULE SET AFC_LASTDATEAPP = "' + usDateTime(iDate2099) + '"';
        vSt := vSt + ',AFC_NEXTDATEAPP = "' + usDateTime(pDtDesapplication) + '"';
        vSt := vSt + ' WHERE AFC_AFFAIRE = "' + pStAffaire + '"';
        vSt := vSt + ' AND AFC_FORCODE = "' + pStFormule + '"';
      end
      else
      begin
        vSt := 'UPDATE AFPARAMFORMULE SET AFC_LASTDATEAPP = "' + usDateTime(vQr.FindField('DATECALCCOEF').AsDateTime) + '"';
        vSt := vSt + ',AFC_NEXTDATEAPP = "' + usDateTime(pDtDesapplication) + '"';
        vSt := vSt + ' WHERE AFC_AFFAIRE = "' + pStAffaire + '"';
        vSt := vSt + ' AND AFC_FORCODE = "' + pStFormule + '"';
      end;
      ExecuteSql(vSt);

    finally
      Ferme(vQr);
    end;
  end;
end;

end.
