unit UAFO_REVPRIXLECTURECOEF;

interface
                                            
uses
  HEnt1, HCtrls, sysutils, forms,
  {$IFNDEF EAGLCLIENT}
  {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
  {$ENDIF}
  UTob, paramsoc,
  Controls, classes,
  dicobtp,
  AGLInit,
  UAFO_REVPRIXCALCULCOEF, utilRevision, EntGC, FactComm;


  function LectureCoefFacture(pTobPiece, pTobRevision : Tob; pDtEcheance : TdateTime; pStLogFile : String; pBoLignes : Boolean) : Boolean;
  function LectureCoefFactureRegul(pTobPiece, pTobRevision : Tob; pDtEcheance : TdateTime; var pInNumRev : Integer; pBoLignes : Boolean) : Boolean;

  procedure AddCoefCalc(pTobRevision, pTobRevisionCalc : Tob);
  procedure UpdateParamFormule(pTobParamFormule, pTobParamFormuleCalc : Tob);
  function LoadParamFormule(var pTob : Tob; pStAffaire : String; pDtEcheance : TdateTime) : Integer;
  function CoefCalculeEtApplique(pTobFormule, pTobRevision : Tob) : Boolean;
  function CoefCalculeNonApplique(pTobFormule, pTobRevision : Tob) : Boolean;
  function CoefNonCalcule(pTobFormule, pTobRevision : Tob) : Boolean;
  function AffaireExisteFormule(pStAffaire : String; var pStForCode : String) : Boolean;
  procedure SelectActivite(pTobOrig, pTobDest : Tob);

const
	TexteMsg: array[1..9] of string 	= (
          {1} '%s Affaire : %s. Erreur lors du chargement de la formule %s.',
          {2} '%s Affaire : %s. Le coefficient de la formule %s est calculé et n''est pas applicable.',
          {3} '%s Affaire : %s. Erreur lors du calcul du coefficient de la formule %s.',
          {4} '%s Affaire : %s. Le coefficient calculé n''est pas le dernier coefficient attendu.',
          {5} '%s Affaire : %s. Le coefficient n''est pas calculé et n''est pas applicable.',
          {6} '%s Affaire : %s. La lecture des coefficients du %s est terminée.',
          {7} '%s Affaire : %s. La formule n''est pas paramètrée, la facture n''a pas été générée.',
          {8} '%s Affaire : %s. Erreur lors de la lecture du coefficient du %s, la facture n''a pas été générée.',
          {9} '%s Affaire : %s. La date de révision est inférieure à la première date de révision.'
          );


implementation

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 10/05/2002
Modifié le ... :   /  /
Description .. : Calcul de la date de revision attendue pour cette revision
                 Test : le temps entre la derniere date de revision et la
                        date de facturation doit etre inferieure a la periodicite
                        de revision
                        si la revision est calculée, mais pas appliquer, lancer l'application

Mots clefs ... :
*****************************************************************}
function LoadParamFormule(var pTob : Tob; pStAffaire : String; pDtEcheance : TdateTime) : Integer;
var
  vQr             : TQuery;
  vSt             : String;
  vInNbMois       : Integer;
  vInNbSemaine    : Integer;
  vDtDateButoire  : TDateTime;
  i               : Integer;
  vDtRevision     : TDateTime;

begin

  result := 1;
  if not assigned(pTob) then
  begin
    vSt := 'SELECT AFC_FORCODE, AFC_AFFAIRE, AFC_PREMIEREDATE, ';
    vSt := vSt + ' AFC_LASTDATECALC, AFC_NEXTDATEAPP, ';
    vSt := vSt + ' AFC_LASTDATEAPP, AFC_APPLIQUERCOEF, AFC_PERIODREV ';
    vSt := vSt + ' FROM AFPARAMFORMULE';
    vSt := vSt + ' WHERE AFC_AFFAIRE = "' + pStAffaire + '"';

    pTob := TOB.Create('AFPARAMFORMULE', nil, -1);
    vQr := nil;
    vQR := OpenSql(vSt, True,-1,'',true);
    try
      if Not vQR.Eof then
        pTob.LoadDetailDB('AFPARAMFORMULE','','',vQR,False,True);
    Finally
      if vQR <> nil then ferme(vQR);
    end;
  end;

  if assigned(pTob) and (pTob.detail.count > 0) then
  begin
    pTob.detail[0].AddChampSup('COEFCALCULE', True);
    pTob.detail[0].AddChampSup('DATEREVISION', True);
    for i := 0 to pTob.detail.count - 1 do
    begin
      if pTob.detail[i].getValue('AFC_LASTDATECALC') = iDate2099 then
        pTob.detail[i].PutValue('COEFCALCULE', '-')
      else
      begin
        vInNbSemaine := 0;
        vInNbMois := 0;
        if pTob.detail[i].getValue('AFC_PERIODREV') = 'A' then
          vInNbMois := 12
        else if pTob.detail[i].getValue('AFC_PERIODREV')  = 'H' then
          vInNbSemaine := 1
        else if pTob.detail[i].getValue('AFC_PERIODREV')  = 'M' then
          vInNbMois := 1
        else if pTob.detail[i].getValue('AFC_PERIODREV')  = 'S' then
          vInNbMois := 6
        else if pTob.detail[i].getValue('AFC_PERIODREV')  = 'T' then
          vInNbMois := 3;

        if vInNbSemaine <> 0 then
          vDtDateButoire := PlusDate(pTob.detail[i].getValue('AFC_LASTDATECALC'), vInNbSemaine, 'S')
        else
          vDtDateButoire := PlusMois(pTob.detail[i].getValue('AFC_LASTDATECALC'), vInNbMois);

        // la date butoire doit être strictement supérieure à la date
        // d'échéance car on attend un calcul du nouveau coef si égalité
        if (vDtDateButoire > pDtEcheance)  then
           pTob.detail[i].PutValue('COEFCALCULE', 'X')
        else
           pTob.detail[i].PutValue('COEFCALCULE', '-');

        vDtRevision := pTob.detail[i].getValue('AFC_LASTDATECALC');

        // on calcule la date de revision utilisée pour le calcul
        // supérieur ou égal
        if vInNbSemaine = 0 then
//        while PlusMois(vDtRevision, - vInNbMois) >=  pDtEcheance do
        while vDtRevision >  pDtEcheance do
          vDtRevision := PlusMois(vDtRevision, - vInNbMois);

        // 26/08/2003
        // ajout d'un test qui permet de vérifié que la date de révision
        // trouvée n'est pas inférieure à la première date de révision possible

        // C.B 22/10/03
        // si AFC_PREMIEREDATE est une fin de mois
        // 1. on compare la premiere date de révision à la fin de mois de la date de revision
        // 2. et si le mois de la première date de révision est le le mois de revision attendu,
        // on force la date de revision a la fin de mois
        if pTob.detail[i].getValue('AFC_PREMIEREDATE') = FinDeMois(pTob.detail[i].getValue('AFC_PREMIEREDATE')) then
        begin
          if pTob.detail[i].getValue('AFC_PREMIEREDATE') = FinDeMois(vDtRevision) then
            vDtRevision := FinDeMois(vDtRevision)
          else if FinDeMois(pTob.detail[i].getValue('AFC_PREMIEREDATE')) > FinDeMois(vDtRevision) then
            result := -2;
        end
        else if pTob.detail[i].getValue('AFC_PREMIEREDATE') > vDtRevision then
          result := -2;

        pTob.detail[i].PutValue('DATEREVISION', vDtRevision);

      end;
    end;
  end
  else
    result := -1;
end;

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 30/04/2003
Modifié le ... :   /  /
Description .. : Recherche si la ligne est une ligne à regulariser ou non
Mots clefs ... :
*****************************************************************}
procedure LoadRegularise(pStAffaire : String; pTob : Tob; pBolignes : Boolean);
var
  vSt : String;
  vQr : TQuery;
  i   : Integer;

Begin

  if pBoLignes then
  begin
    vSt := vSt + 'SELECT AFE_FORCODE, AFE_INDTYPE ';
    vSt := vSt + ' FROM LIGNE, AFFORMULE ';
    vSt := vSt + ' WHERE GL_AFFAIRE = "' + pStAffaire + '"';
    vSt := vSt + ' AND GL_FORCODE1 = AFE_FORCODE ';

    vSt := vSt + ' UNION ';
    vSt := vSt + ' SELECT AFE_FORCODE, AFE_INDTYPE ';
    vSt := vSt + ' FROM LIGNE, AFFORMULE ';
    vSt := vSt + ' WHERE GL_AFFAIRE = "' + pStAffaire + '"';
    vSt := vSt + ' AND GL_FORCODE2 = AFE_FORCODE ';
  end
  else
  begin
    vSt := 'SELECT AFE_FORCODE, AFE_INDTYPE ';
    vSt := vSt + ' FROM AFFAIRE, AFFORMULE ';
    vSt := vSt + ' WHERE AFF_AFFAIRE = "' + pStAffaire + '"';
    vSt := vSt + ' AND AFF_FORCODE1 = AFE_FORCODE ';

    vSt := vSt + ' UNION ';

    vSt := vSt + ' SELECT AFE_FORCODE, AFE_INDTYPE ';
    vSt := vSt + ' FROM AFFAIRE, AFFORMULE ';
    vSt := vSt + ' WHERE AFF_AFFAIRE = "' + pStAffaire + '"';
    vSt := vSt + ' AND AFF_FORCODE2 = AFE_FORCODE ';
  end;

  vQR := OpenSql(vSt, True,-1,'',true);
  Try
    if Not vQR.Eof then
    begin
      pTob.LoadDetailDB('AFFORMULE','','',vQR,False,True);
      for i := 0 to pTob.detail.count -1 do
      begin
        if pTob.Detail[i].Getvalue('AFE_INDTYPE') <> 'REE' then
          pTob.detail[i].AddChampSupValeur('REGULARISE', 'X')
        else
          pTob.detail[i].AddChampSupValeur('REGULARISE', '-');
      end;
    end;
  Finally
    if vQR <> nil then ferme(vQR);
  End;
End;

procedure SelectActivite(pTobOrig, pTobDest : Tob);
var
  vSt       : String;
  vQr       : TQuery;
  vStPiece  : String;

begin

  // recherche si la ligne provient de l'activité
  vStPiece := intToStr(pTobOrig.GetValue('GL_NUMERO')) + ';';
  vStPiece := vStPiece + intToStr(pTobOrig.GetValue('GL_INDICEG'));
                  
  vSt := 'SELECT ACT_NUMORDRE, ACT_ARTICLE FROM ACTIVITE';       
  vSt := vSt + ' WHERE ACT_NUMPIECEORIG like ';
  vSt := vSt + ' "%' + VH_GC.AFNatAffaire + ';AFF;' + vStPiece + '%"';

  vQR := OpenSql(vSt, True,-1,'',true);
  try
    if Not vQR.Eof then
      pTobDest.LoadDetailDB('ACTIVITE','','', vQr, False, True);
  Finally
    if vQR <> nil then ferme(vQR);
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 10/05/2003
Modifié le ... :
Description .. : creation des lignes de revision pour la facture
Mots clefs ... :
*****************************************************************}
function PrepareLignesRevision(pTobPiece, pTobRevision, pTobRevisionAff,
                               pTobParamFormule : Tob; pDtEcheance : TDateTime;
                               pBoLignes : Boolean) : Boolean;
var
  j                 : Integer;
  vTobTheRev1       : Tob;
  vTobTheRev2       : Tob;
  vTobTheFormule    : Tob;
  vStAffaire        : String;
  vRdRemise         : Double;
  vRdPrixHt         : Double;
  vRdPrixHtDev      : Double;
  vRdPrixOrigine    : Double;
  vRdPrixOrigineDev : Double;
  vInAFFAIRE        : Integer;
  vInNUMERO         : Integer;
  vInINDICEG        : Integer;
  vInNUMORDRE       : Integer;
  vInARTICLE        : Integer;
  vInFORCODE1       : Integer;
  vInFORCODE2       : Integer;
  vInNumRev         : Integer;
  vTobFor1          : Tob;
  vTobFor2          : Tob;
  Year, Month, Day  : Word;
  vInValNumOrdre    : Integer;
  vTobActivite      : Tob;
  LaTobActivite     : Tob;
  madate : tdatetime;
  
begin

  result := True;
  vStAffaire := pTobPiece.detail[0].GetValue('GL_AFFAIRE');
  vInNumRev := MaxRevision(vStAffaire);

  vTobActivite := Tob.create('ACTIVITE', nil, -1);
  SelectActivite(pTobPiece.detail[0], vTobActivite);

  // chargement des formules de l'affaire
  vInAFFAIRE  := pTobPiece.detail[0].GetNumChamp('GL_AFFAIRE');

//  vTobRegularise := TOB.create('REGULARISE',nil,-1);
  Try
//    LoadRegularise(vStAffaire, vTobRegularise, True);

    vInNUMERO   := pTobPiece.detail[0].GetNumChamp('GL_NUMERO');
    vInINDICEG  := pTobPiece.detail[0].GetNumChamp('GL_INDICEG');
    vInNUMORDRE := pTobPiece.detail[0].GetNumChamp('GL_NUMORDRE');
    vInFORCODE1 := pTobPiece.detail[0].GetNumChamp('GL_FORCODE1');
    vInFORCODE2 := pTobPiece.detail[0].GetNumChamp('GL_FORCODE2');
    vInARTICLE  := pTobPiece.detail[0].GetNumChamp('GL_ARTICLE');

    // pour chaque lignes de la piece
    for j := 0 to pTobPiece.detail.count -1 do
    begin

      if (pTobPiece.detail[j].GetValue('GL_TYPELIGNE') = 'ART') and
         (pTobPiece.detail[j].GetValue('GL_TYPEARTICLE') <> 'POU') then
      begin
        vTobTheRev1 := nil;
        vTobTheRev2 := nil;

        // recherche s'il existe un coef1 puis un coef2 a la ligne
        if pTobPiece.detail[j].GetValeur(vInFORCODE1) <> '' then
        begin
          vTobTheFormule := pTobParamFormule.FindFirst(['AFC_FORCODE'],
                                                       [pTobPiece.detail[j].GetValeur(vInFORCODE1)],
                                                       True);

          LaTobActivite := vTobActivite.FindFirst(['ACT_ARTICLE'], [pTobPiece.detail[j].GetValeur(vInARTICLE)], True);
          if LaTobActivite <> nil then
            vInValNumOrdre := LaTobActivite.GetValue('ACT_NUMORDRE')
          else
            vInValNumOrdre := pTobPiece.detail[j].GetValeur(vInNUMORDRE);
          madate := vTobTheFormule.GetValue('DATEREVISION');
          vTobTheRev1 := pTobRevisionAff.FindFirst(['AFR_AFFAIRE',
                                            'AFR_NATUREPIECEG',
                                            'AFR_NUMERO',
                                            'AFR_INDICEG',
                                            'AFR_NUMORDRE',
                                            'AFR_DATECALCCOEF',
                                            'AFR_FORCODE',
                                            'AFR_NUMFORMULE',
                                            'AFR_OKCOEFAPPLIQUE'],
                                            [pTobPiece.detail[j].GetValeur(vInAFFAIRE),
                                             VH_GC.AFNatAffaire, //pTobPiece.detail[j].GetValeur(vInNATUREPIECEG),
                                             pTobPiece.detail[j].GetValeur(vInNUMERO),
                                             pTobPiece.detail[j].GetValeur(vInINDICEG),
                                             vInValNumOrdre,
                                             vTobTheFormule.GetValue('DATEREVISION'),
                                             pTobPiece.detail[j].GetValeur(vInFORCODE1),
                                             1,
                                             'X'],
                                             True);

          // bug si la date est 31 a l'origine elle peut etre a 30 ensuite
          if vTobTheRev1 = nil then
          begin
            decodeDate(vTobTheFormule.GetValue('DATEREVISION'), Year, Month, Day);
            if Day = 30 then
              vTobTheRev1 := pTobRevisionAff.FindFirst(['AFR_AFFAIRE',
                                                'AFR_NATUREPIECEG',
                                                'AFR_NUMERO',
                                                'AFR_INDICEG',
                                                'AFR_NUMORDRE',
                                                'AFR_DATECALCCOEF',
                                                'AFR_FORCODE',
                                                'AFR_NUMFORMULE',
                                                'AFR_OKCOEFAPPLIQUE'],
                                                [pTobPiece.detail[j].GetValeur(vInAFFAIRE),
                                                 VH_GC.AFNatAffaire,
                                                 pTobPiece.detail[j].GetValeur(vInNUMERO),
                                                 pTobPiece.detail[j].GetValeur(vInINDICEG),
                                                 vInValNumOrdre,
                                                 plusDate(vTobTheFormule.GetValue('DATEREVISION'),1,'J'),
                                                 pTobPiece.detail[j].GetValeur(vInFORCODE1),
                                                 1,
                                                 'X'],
                                                 True);
          end;
        end; // end formule 1

        if pTobPiece.detail[j].GetValeur(vInFORCODE2) <> '' then
        begin
          vTobTheFormule := pTobParamFormule.FindFirst(['AFC_FORCODE'],
                                                       [pTobPiece.detail[j].GetValeur(vInFORCODE2)],
                                                       True);

          LaTobActivite := vTobActivite.FindFirst(['ACT_ARTICLE'], [pTobPiece.detail[j].GetValeur(vInARTICLE)], True);
          if LaTobActivite <> nil then
            vInValNumOrdre := LaTobActivite.GetValue('ACT_NUMORDRE')
          else
            vInValNumOrdre := pTobPiece.detail[j].GetValeur(vInNUMORDRE);

          vTobTheRev2 := pTobRevisionAff.FindFirst(['AFR_AFFAIRE',
                                            'AFR_NATUREPIECEG',
                                            'AFR_NUMERO',
                                            'AFR_INDICEG',
                                            'AFR_NUMORDRE',
                                            'AFR_DATECALCCOEF',
                                            'AFR_FORCODE',
                                            'AFR_NUMFORMULE',
                                            'AFR_OKCOEFAPPLIQUE'],
                                            [pTobPiece.detail[j].GetValeur(vInAFFAIRE),
                                             VH_GC.AFNatAffaire,
                                             pTobPiece.detail[j].GetValeur(vInNUMERO),
                                             pTobPiece.detail[j].GetValeur(vInINDICEG),
                                             vInValNumOrdre,
                                             vTobTheFormule.Getvalue('DATEREVISION'),
                                             pTobPiece.detail[j].GetValeur(vInFORCODE2),
                                             2,
                                             'X'],
                                             True);

          // bug si la date est 31 a l'origine elle peut etre a 30 ensuite
          if vTobTheRev2 = nil then
          begin
            decodeDate(vTobTheFormule.GetValue('DATEREVISION'), Year, Month, Day);
            if Day = 30 then
              vTobTheRev2 := pTobRevisionAff.FindFirst(['AFR_AFFAIRE',
                                                'AFR_NATUREPIECEG',
                                                'AFR_NUMERO',
                                                'AFR_INDICEG',
                                                'AFR_NUMORDRE',
                                                'AFR_DATECALCCOEF',
                                                'AFR_FORCODE',
                                                'AFR_NUMFORMULE',
                                                'AFR_OKCOEFAPPLIQUE'],
                                                [pTobPiece.detail[j].GetValeur(vInAFFAIRE),
                                                 VH_GC.AFNatAffaire,
                                                 pTobPiece.detail[j].GetValeur(vInNUMERO),
                                                 pTobPiece.detail[j].GetValeur(vInINDICEG),
                                                 vInValNumOrdre,
                                                 PlusDate(vTobTheFormule.Getvalue('DATEREVISION'),1,'J'),
                                                 pTobPiece.detail[j].GetValeur(vInFORCODE2),
                                                 2,
                                                 'X'],
                                                 True);
          end;
        end;

        // initialisation de la clé
        // calcul de la valeur remise avec coef
        // calcul du prix unitaire
        vRdRemise := 0;
        vRdPrixHt := 0;
        vRdPrixHtDev := 0;
        vRdPrixOrigine := 0;
        vRdPrixOrigineDev := 0;

        if assigned(vTobTheRev1) then
        begin
          // creation de l'enregistrement dans la tob revision
          vTobFor1 := tob.create('AFREVISION', nil, -1);
          vTobFor1.Dupliquer(vTobTheRev1, True, True);
          vTobFor1.PutValue('AFR_NUMEROLIGNE', vInNumRev);
          vTobFor1.ChangeParent(pTobRevision, -1);

          // lien avec la tob des pieces
          if pTobPiece.detail[j].GetValue('GL_FORCODE1') = '' then
            pTobPiece.detail[j].PutValue('GL_FORCODE1', vTobTheRev1.getvalue('AFR_FORCODE'));

          pTobPiece.detail[j].PutValue('GLLAFFAIRE1', vTobTheRev1.getvalue('AFR_AFFAIRE'));
          pTobPiece.detail[j].PutValue('GLLFORCODE1', vTobTheRev1.getvalue('AFR_FORCODE'));
          pTobPiece.detail[j].PutValue('GLLNUMEROLIGNE1', vInNumRev);

          vRdRemise     := pTobPiece.detail[j].GetValue('GL_VALEURREMDEV') * vTobTheRev1.getvalue('AFR_COEFAPPLIQUE');
          vRdPrixHt     := pTobPiece.detail[j].GetValue('GL_PUHT') * vTobTheRev1.getvalue('AFR_COEFAPPLIQUE');
          vRdPrixHtDev  := pTobPiece.detail[j].GetValue('GL_PUHTDEV') * vTobTheRev1.getvalue('AFR_COEFAPPLIQUE');

          vRdPrixOrigine :=  pTobPiece.detail[j].GetValue('GL_PUHT');
          vRdPrixOrigineDev := pTobPiece.detail[j].GetValue('GL_PUHTDEV');

{          vTobTheRegul := vTobRegularise.FindFirst(['AFE_FORCODE'], [vTobTheRev1.GetValue('AFR_FORCODE')], True);
          if vTobTheRegul.GetValue('REGULARISE') = 'X' then
            pTobPiece.detail[j].PutValue('GL_REGULARISE', 'X')
          else
            pTobPiece.detail[j].PutValue('GL_REGULARISE', '-');
}
          vInNumRev := vInNumRev + 1;
        end
        else if (pTobPiece.detail[j].GetValeur(vInFORCODE1) <> '') then
          result := False;

        if assigned(vTobTheRev2) then
        begin
          // creation de l'enregistrement dans la tob revision
          vTobFor2 := tob.create('AFREVISION', nil, -1);
          vTobFor2.Dupliquer(vTobTheRev2, True, True);
          vTobFor2.PutValue('AFR_NUMEROLIGNE', vInNumRev);
          vTobFor2.ChangeParent(pTobRevision, -1);

          // lien avec la tob des pieces
          if pTobPiece.detail[j].GetValue('GL_FORCODE2') = '' then
            pTobPiece.detail[j].PutValue('GL_FORCODE2', vTobTheRev2.getvalue('AFR_FORCODE'));

          pTobPiece.detail[j].PutValue('GLLAFFAIRE2', vTobTheRev2.getvalue('AFR_AFFAIRE'));
          pTobPiece.detail[j].PutValue('GLLFORCODE2', vTobTheRev2.getvalue('AFR_FORCODE'));
          pTobPiece.detail[j].PutValue('GLLNUMEROLIGNE2', vInNumRev);

          if vRdRemise = 0 then
            vRdRemise := pTobPiece.detail[j].GetValue('GL_VALEURREMDEV') * vTobTheRev2.getvalue('AFR_COEFAPPLIQUE')
          else
            vRdRemise := vRdRemise * vTobTheRev2.getvalue('AFR_COEFAPPLIQUE');

          if vRdPrixHt = 0 then
            vRdPrixHt := pTobPiece.detail[j].GetValue('GL_PUHT') * vTobTheRev2.getvalue('AFR_COEFAPPLIQUE')
          else
            vRdPrixHt := vRdPrixHt * vTobTheRev2.getvalue('AFR_COEFAPPLIQUE');

          if vRdPrixHtDev = 0 then
            vRdPrixHtDev := pTobPiece.detail[j].GetValue('GL_PUHTDEV') * vTobTheRev2.getvalue('AFR_COEFAPPLIQUE')
          else
            vRdPrixHtDev := vRdPrixHtDev * vTobTheRev2.getvalue('AFR_COEFAPPLIQUE');

          if vRdPrixOrigine = 0 then
            vRdPrixOrigine := pTobPiece.detail[j].GetValue('GL_PUHT')
          else
            vRdPrixOrigine := vRdPrixOrigine;

          if vRdPrixOrigineDev = 0 then
            vRdPrixOrigineDev := pTobPiece.detail[j].GetValue('GL_PUHTDEV')
          else
            vRdPrixOrigineDev := vRdPrixOrigineDev;

          vInNumRev := vInNumRev + 1;
        end
        else if (pTobPiece.detail[j].GetValeur(vInFORCODE2) <> '') then
          result := False;

        if assigned(vTobTheRev1) or assigned(vTobTheRev2) then
        begin
          pTobPiece.detail[j].PutValue('GL_VALEURREMDEV', vRdRemise);
          pTobPiece.detail[j].PutValue('GL_PUHT', vRdPrixHt);
          pTobPiece.detail[j].PutValue('GL_PUHTDEV', vRdPrixHtDev);
          pTobPiece.detail[j].PutValue('GL_PUHTORIGINE', vRdPrixOrigine);
          pTobPiece.detail[j].PutValue('GL_PUHTORIGINEDEV', vRdPrixOrigineDev);
        end;
      end; // end if
    end; // end for

  Finally
//    vTobRegularise.Free;
    vTobActivite.Free;
  end;
end;

procedure PrepareLignesRevisionRegul(pTobPiece, pTobRevision, pTobRevisionAff, pTobParamFormule : Tob; pDtEcheance : TDateTime; var pInNumRev : Integer; pBoLignes : Boolean);
var
  j                 : Integer;
  vTobTheRev1       : Tob;
  vTobTheRev2       : Tob;
//  vTobTheFormule    : Tob;
  vStAffaire        : String;
  vRdRemise         : Double;
  vRdPrixHt         : Double;
  vRdPrixHtDev      : Double;
  vRdPrixOrigine    : Double;
  vRdPrixOrigineDev : Double;
  vTobRegularise    : Tob;
  vTobTheRegul      : Tob;
  vInAFFAIRE        : Integer;
  vInFORCODE1       : Integer;
  vInFORCODE2       : Integer;
  vTobFor1          : Tob;
  vTobFor2          : Tob;
//  Year, Month, Day  : Word;

begin

  vStAffaire := pTobPiece.detail[0].GetValue('GL_AFFAIRE');

  // chargement des formules de l'affaire
  vTobRegularise := TOB.create('REGULARISE',nil,-1);
  Try
    LoadRegularise(vStAffaire, vTobRegularise, True);

    vInAFFAIRE      := pTobPiece.detail[0].GetNumChamp('GL_AFFAIRE');
    vInFORCODE1     := pTobPiece.detail[0].GetNumChamp('GL_FORCODE1');
    vInFORCODE2     := pTobPiece.detail[0].GetNumChamp('GL_FORCODE2');

    // pour chaque piece
    //k := 0;
    for j := 0 to pTobPiece.detail.count -1 do
    begin
      if (pTobPiece.detail[j].GetValue('GL_TYPELIGNE') = 'ART') and
         (pTobPiece.detail[j].GetValue('GL_TYPEARTICLE') <> 'POU') then
      begin
        // on ne traite que les premières lignes des paires d'articles pour les regul ...
        //if (frac(k/2) = 0) then
        if pTobPiece.detail[j].GetValue('GLLDATEPIECE') <> iDate1900 then
        begin
          vTobTheRev1 := nil;
          vTobTheRev2 := nil;

          // recherche s'il existe un coef1 puis un coef2 a la ligne
          if pTobPiece.detail[j].GetValeur(vInFORCODE1) <> '' then
          begin
            //vTobTheFormule := pTobParamFormule.FindFirst(['AFC_FORCODE'],
            //                                             [pTobPiece.detail[j].GetValeur(vInFORCODE1)],
            //                                             True);

            vTobTheRev1 := pTobRevisionAff.FindFirst(['AFR_AFFAIRE',
                                              'AFR_NATUREPIECEG',
//                                              'AFR_NUMERO',
                                              'AFR_DATECALCCOEF',
                                              'AFR_FORCODE',
                                              'AFR_NUMFORMULE',
                                              'AFR_OKCOEFAPPLIQUE'],
                                              [pTobPiece.detail[j].GetValeur(vInAFFAIRE),
                                               VH_GC.AFNatAffaire,
//                                               pInNumPieceAff,
                                               //vTobTheFormule.GetValue('DATEREVISION'),
                                               //pDtEcheance,
                                               pTobPiece.detail[j].GetValue('GLLDATEPIECE'),
                                               pTobPiece.detail[j].GetValeur(vInFORCODE1),
                                               1,
                                               'X'],
                                               True);
          end;

          if pTobPiece.detail[j].GetValeur(vInFORCODE2) <> '' then
          begin
            //vTobTheFormule := pTobParamFormule.FindFirst(['AFC_FORCODE'],
            //                                             [pTobPiece.detail[j].GetValeur(vInFORCODE2)],
            //                                             True);

            vTobTheRev2 := pTobRevisionAff.FindFirst(['AFR_AFFAIRE',
                                              'AFR_NATUREPIECEG',
//                                              'AFR_NUMERO',
                                              'AFR_DATECALCCOEF',
                                              'AFR_FORCODE',
                                              'AFR_NUMFORMULE',
                                              'AFR_OKCOEFAPPLIQUE'],
                                              [pTobPiece.detail[j].GetValeur(vInAFFAIRE),
                                               VH_GC.AFNatAffaire,
//                                               pInNumPieceAff,
                                               //vTobTheFormule.Getvalue('DATEREVISION'),
                                               //pDtEcheance,
                                               pTobPiece.detail[j].GetValue('GLLDATEPIECE'),
                                               pTobPiece.detail[j].GetValeur(vInFORCODE2),
                                               2,
                                               'X'],
                                               True);

          end;

          // initialisation de la clé
          // calcul de la valeur remise avec coef
          // calcul du prix unitaire
          vRdRemise := 0;
          vRdPrixHt := 0;
          vRdPrixHtDev := 0;
          vRdPrixOrigine := 0;
          vRdPrixOrigineDev := 0;

          if assigned(vTobTheRev1) then
          begin
            // creation de l'enregistrement dans la tob revision
            vTobFor1 := tob.create('AFREVISION', nil, -1);
            vTobFor1.Dupliquer(vTobTheRev1, True, True);
            vTobFor1.PutValue('AFR_NUMEROLIGNE', pInNumRev);
            vTobFor1.ChangeParent(pTobRevision, -1);

            // lien avec la tob des pieces
            if pTobPiece.detail[j].GetValue('GL_FORCODE1') = '' then
              pTobPiece.detail[j].PutValue('GL_FORCODE1', vTobTheRev1.getvalue('AFR_FORCODE'));

            pTobPiece.detail[j].PutValue('GLLAFFAIRE1', vTobTheRev1.getvalue('AFR_AFFAIRE'));
            pTobPiece.detail[j].PutValue('GLLFORCODE1', vTobTheRev1.getvalue('AFR_FORCODE'));
            pTobPiece.detail[j].PutValue('GLLNUMEROLIGNE1', pInNumRev);

            vRdRemise     := pTobPiece.detail[j].GetValue('GL_VALEURREMDEV') * vTobTheRev1.getvalue('AFR_COEFAPPLIQUE');
            vRdPrixHt     := pTobPiece.detail[j].GetValue('GL_PUHT') * vTobTheRev1.getvalue('AFR_COEFAPPLIQUE');
            vRdPrixHtDev  := pTobPiece.detail[j].GetValue('GL_PUHTDEV') * vTobTheRev1.getvalue('AFR_COEFAPPLIQUE');

            vRdPrixOrigine :=  pTobPiece.detail[j].GetValue('GL_PUHT');
            vRdPrixOrigineDev := pTobPiece.detail[j].GetValue('GL_PUHTDEV');

            vTobTheRegul := vTobRegularise.FindFirst(['AFE_FORCODE'], [vTobTheRev1.GetValue('AFR_FORCODE')], True);
            if vTobTheRegul.GetValue('REGULARISE') = 'X' then
              pTobPiece.detail[j].PutValue('GL_REGULARISE', 'X')
            else
              pTobPiece.detail[j].PutValue('GL_REGULARISE', '-');

            pInNumRev := pInNumRev + 1;
          end;

          if assigned(vTobTheRev2) then
          begin
            // creation de l'enregistrement dans la tob revision
            vTobFor2 := tob.create('AFREVISION', nil, -1);
            vTobFor2.Dupliquer(vTobTheRev2, True, True);
            vTobFor2.PutValue('AFR_NUMEROLIGNE', pInNumRev);
            vTobFor2.ChangeParent(pTobRevision, -1);

            // lien avec la tob des pieces
            if pTobPiece.detail[j].GetValue('GL_FORCODE2') = '' then
              pTobPiece.detail[j].PutValue('GL_FORCODE2', vTobTheRev2.getvalue('AFR_FORCODE'));

            pTobPiece.detail[j].PutValue('GLLAFFAIRE2', vTobTheRev2.getvalue('AFR_AFFAIRE'));
            pTobPiece.detail[j].PutValue('GLLFORCODE2', vTobTheRev2.getvalue('AFR_FORCODE'));
            pTobPiece.detail[j].PutValue('GLLNUMEROLIGNE2', pInNumRev);

            if vRdRemise = 0 then
              vRdRemise := pTobPiece.detail[j].GetValue('GL_VALEURREMDEV') * vTobTheRev2.getvalue('AFR_COEFAPPLIQUE')
            else
              vRdRemise := vRdRemise * vTobTheRev2.getvalue('AFR_COEFAPPLIQUE');

            if vRdPrixHt = 0 then
              vRdPrixHt := pTobPiece.detail[j].GetValue('GL_PUHT') * vTobTheRev2.getvalue('AFR_COEFAPPLIQUE')
            else
              vRdPrixHt := vRdPrixHt * vTobTheRev2.getvalue('AFR_COEFAPPLIQUE');

            if vRdPrixHtDev = 0 then
              vRdPrixHtDev := pTobPiece.detail[j].GetValue('GL_PUHTDEV') * vTobTheRev2.getvalue('AFR_COEFAPPLIQUE')
            else
              vRdPrixHtDev := vRdPrixHtDev * vTobTheRev2.getvalue('AFR_COEFAPPLIQUE');

            if vRdPrixOrigine = 0 then
              vRdPrixOrigine := pTobPiece.detail[j].GetValue('GL_PUHT')
            else
              vRdPrixOrigine := vRdPrixOrigine;

            if vRdPrixOrigineDev = 0 then
              vRdPrixOrigineDev := pTobPiece.detail[j].GetValue('GL_PUHTDEV')
            else
              vRdPrixOrigineDev := vRdPrixOrigineDev;

            pInNumRev := pInNumRev + 1;
          end;

          if assigned(vTobTheRev1) or assigned(vTobTheRev2) then
          begin
            pTobPiece.detail[j].PutValue('GL_VALEURREMDEV', vRdRemise);
            pTobPiece.detail[j].PutValue('GL_PUHT', vRdPrixHt);
            pTobPiece.detail[j].PutValue('GL_PUHTDEV', vRdPrixHtDev);
            pTobPiece.detail[j].PutValue('GL_PUHTORIGINE', vRdPrixOrigine);
            pTobPiece.detail[j].PutValue('GL_PUHTORIGINEDEV', vRdPrixOrigineDev);
          end;
        end; // end if regul
        //k := k + 1;
      end;
    end; // end for

  Finally
    vTobRegularise.Free;
  end;

end;
                                         
// Avec la gestion au niveau des entetes
{
procedure PrepareLignesRevision(pTobPiece, pTobRevision, pTobRevisionAff, pTobParamFormule : Tob; pDtEcheance : TDateTime; pBoLigneRegul : Boolean; pInNumPieceAff : Integer);
var
  j                 : Integer;
  vTobTheRev1       : Tob;
  vTobTheRev2       : Tob;
  vTobTheFormule    : Tob;
  vStAffaire        : String;
  vRdRemise         : Double;
  vRdPrixHt         : Double;
  vRdPrixHtDev      : Double;
  vRdPrixOrigine    : Double;
  vRdPrixOrigineDev : Double;
  vTobRegularise    : Tob;
  vTobTheRegul      : Tob;
  vInAFFAIRE        : Integer;
  vInNUMERO         : Integer;
  vInINDICEG        : Integer;
  vInNUMORDRE       : Integer;
  vInFORCODE1       : Integer;
  vInFORCODE2       : Integer;
  vInNumRev         : Integer;
  vTobFor1          : Tob;
  vTobFor2          : Tob;
  var Year, Month, Day : Word;
begin

  vStAffaire := pTobPiece.detail[0].GetValue('GL_AFFAIRE');
  vInNumRev := MaxRevision(vStAffaire);

  // chargement des formules de l'affaire
  vTobRegularise := TOB.create('REGULARISE',nil,-1);
  Try
    LoadRegularise(vStAffaire, vTobRegularise);

    vInAFFAIRE      := pTobPiece.detail[0].GetNumChamp('GL_AFFAIRE');
    vInNUMERO       := pTobPiece.detail[0].GetNumChamp('GL_NUMERO');
    vInINDICEG      := pTobPiece.detail[0].GetNumChamp('GL_INDICEG');
    vInNUMORDRE     := pTobPiece.detail[0].GetNumChamp('GL_NUMORDRE');
    vInFORCODE1     := pTobPiece.detail[0].GetNumChamp('GL_FORCODE1');
    vInFORCODE2     := pTobPiece.detail[0].GetNumChamp('GL_FORCODE2');

    // pour chaque piece
    for j := 0 to pTobPiece.detail.count -1 do
    begin
      // on ne traite que les lignes paire pour les regul ...
      if (not pBoLigneRegul) or ((pBoLigneRegul) and (frac(j/2) = 0)) then
      begin
        if (pTobPiece.detail[j].GetValue('GL_TYPELIGNE') = 'ART') and
           (pTobPiece.detail[j].GetValue('GL_TYPEARTICLE') <> 'POU') then
        begin
          vTobTheRev1 := nil;
          vTobTheRev2 := nil;

          // recherche s'il existe un coef1 puis un coef2 a la ligne
          if pTobPiece.detail[j].GetValeur(vInFORCODE1) <> '' then
          begin
            vTobTheFormule := pTobParamFormule.FindFirst(['AFC_FORCODE'],
                                                         [pTobPiece.detail[j].GetValeur(vInFORCODE1)],
                                                         True);

            vTobTheRev1 := pTobRevisionAff.FindFirst(['AFR_AFFAIRE',
                                              'AFR_NATUREPIECEG',
                                              'AFR_NUMERO',
                                              'AFR_INDICEG',
                                              'AFR_NUMORDRE',
                                              'AFR_DATECALCCOEF',
                                              'AFR_FORCODE',
                                              'AFR_NUMFORMULE',
                                              'AFR_OKCOEFAPPLIQUE'],
                                              [pTobPiece.detail[j].GetValeur(vInAFFAIRE),
                                               VH_GC.AFNatAffaire,
                                               pTobPiece.detail[j].GetValeur(vInNUMERO),
                                               pTobPiece.detail[j].GetValeur(vInINDICEG),
                                               pTobPiece.detail[j].GetValeur(vInNUMORDRE),
                                               vTobTheFormule.GetValue('DATEREVISION'),
                                               pTobPiece.detail[j].GetValeur(vInFORCODE1),
                                               1,
                                               'X'],
                                               True);

            // bug si la date est 31 a l'origine elle peut etre a 30 ensuite
            if vTobTheRev1 = nil then
            begin
              decodeDate(vTobTheFormule.GetValue('DATEREVISION'), Year, Month, Day);
              if Day = 30 then
                vTobTheRev1 := pTobRevisionAff.FindFirst(['AFR_AFFAIRE',
                                                  'AFR_NATUREPIECEG',
                                                  'AFR_NUMERO',
                                                  'AFR_INDICEG',
                                                  'AFR_NUMORDRE',
                                                  'AFR_DATECALCCOEF',
                                                  'AFR_FORCODE',
                                                  'AFR_NUMFORMULE',
                                                  'AFR_OKCOEFAPPLIQUE'],
                                                  [pTobPiece.detail[j].GetValeur(vInAFFAIRE),
                                                   VH_GC.AFNatAffaire,
                                                   pTobPiece.detail[j].GetValeur(vInNUMERO),
                                                   pTobPiece.detail[j].GetValeur(vInINDICEG),
                                                   pTobPiece.detail[j].GetValeur(vInNUMORDRE),
                                                   plusDate(vTobTheFormule.GetValue('DATEREVISION'),1,'J'),
                                                   pTobPiece.detail[j].GetValeur(vInFORCODE1),
                                                   1,
                                                   'X'],
                                                   True);
            end;
          end;

          if pTobPiece.detail[j].GetValeur(vInFORCODE2) <> '' then
          begin
            vTobTheFormule := pTobParamFormule.FindFirst(['AFC_FORCODE'],
                                                         [pTobPiece.detail[j].GetValeur(vInFORCODE2)],
                                                         True);

            vTobTheRev2 := pTobRevisionAff.FindFirst(['AFR_AFFAIRE',
                                              'AFR_NATUREPIECEG',
                                              'AFR_NUMERO',
                                              'AFR_INDICEG',
                                              'AFR_NUMORDRE',
                                              'AFR_DATECALCCOEF',
                                              'AFR_FORCODE',
                                              'AFR_NUMFORMULE',
                                              'AFR_OKCOEFAPPLIQUE'],
                                              [pTobPiece.detail[j].GetValeur(vInAFFAIRE),
                                               VH_GC.AFNatAffaire,
                                               pTobPiece.detail[j].GetValeur(vInNUMERO),
                                               pTobPiece.detail[j].GetValeur(vInINDICEG),
                                               pTobPiece.detail[j].GetValeur(vInNUMORDRE),
                                               vTobTheFormule.Getvalue('DATEREVISION'),
                                               pTobPiece.detail[j].GetValeur(vInFORCODE2),
                                               2,
                                               'X'],
                                               True);

            // bug si la date est 31 a l'origine elle peut etre a 30 ensuite
            if vTobTheRev2 = nil then
            begin
              decodeDate(vTobTheFormule.GetValue('DATEREVISION'), Year, Month, Day);
              if Day = 30 then
                vTobTheRev2 := pTobRevisionAff.FindFirst(['AFR_AFFAIRE',
                                                  'AFR_NATUREPIECEG',
                                                  'AFR_NUMERO',
                                                  'AFR_INDICEG',
                                                  'AFR_NUMORDRE',
                                                  'AFR_DATECALCCOEF',
                                                  'AFR_FORCODE',
                                                  'AFR_NUMFORMULE',
                                                  'AFR_OKCOEFAPPLIQUE'],
                                                  [pTobPiece.detail[j].GetValeur(vInAFFAIRE),
                                                   VH_GC.AFNatAffaire,
                                                   pTobPiece.detail[j].GetValeur(vInNUMERO),
                                                   pTobPiece.detail[j].GetValeur(vInINDICEG),
                                                   pTobPiece.detail[j].GetValeur(vInNUMORDRE),
                                                   PlusDate(vTobTheFormule.Getvalue('DATEREVISION'),1,'J'),
                                                   pTobPiece.detail[j].GetValeur(vInFORCODE2),
                                                   2,
                                                   'X'],
                                                   True);
            end;
          end;

          // si non, recherche s'il existe un coef1 puis un coef2 a l'entete
          if not assigned(vTobTheRev1) then
          begin
            // recherche de la formule
            vTobTheRev1 := pTobRevisionAff.FindFirst(['AFR_AFFAIRE',
                                             'AFR_NATUREPIECEG',
                                             'AFR_SOUCHE',
                                             'AFR_NUMERO',
                                             'AFR_INDICEG',
                                             'AFR_NUMORDRE',
                                             'AFR_NUMFORMULE',
                                             'AFR_OKCOEFAPPLIQUE'],
                                             [pTobPiece.detail[j].GetValue('GL_AFFAIRE'),
                                              cStVide,
                                              cStVide,
                                              cInVide,
                                              cInVide,
                                              cInVide,
                                              1,
                                              'X'],
                                              True);

            if assigned(vTobTheRev1) then
            begin

              // recherche de la date
              vTobTheFormule := pTobParamFormule.FindFirst(['AFC_FORCODE'],
                                                           [vTobTheRev1.GetValue('AFR_FORCODE')],
                                                           True);

              // recherche de la formule
              vTobTheRev1 := pTobRevisionAff.FindFirst(['AFR_AFFAIRE',
                                               'AFR_NATUREPIECEG',
                                               'AFR_SOUCHE',
                                               'AFR_NUMERO',
                                               'AFR_INDICEG',
                                               'AFR_NUMORDRE',
                                               'AFR_DATECALCCOEF',
                                               'AFR_NUMFORMULE',
                                               'AFR_OKCOEFAPPLIQUE'],
                                               [pTobPiece.detail[j].GetValue('GL_AFFAIRE'),
                                                cStVide,
                                                cStVide,
                                                cInVide,
                                                cInVide,
                                                cInVide,
                                                vTobTheFormule.Getvalue('DATEREVISION'),
                                                1,
                                                'X'],
                                                True);

              // bug si la date est 31 a l'origine elle peut etre a 30 ensuite
              if vTobTheRev1 = nil then
              begin
                decodeDate(vTobTheFormule.GetValue('DATEREVISION'), Year, Month, Day);
                if Day = 30 then
                  vTobTheRev1 := pTobRevisionAff.FindFirst(['AFR_AFFAIRE',
                                                 'AFR_NATUREPIECEG',
                                                 'AFR_SOUCHE',
                                                 'AFR_NUMERO',
                                                 'AFR_INDICEG',
                                                 'AFR_NUMORDRE',
                                                 'AFR_DATECALCCOEF',
                                                 'AFR_NUMFORMULE',
                                                 'AFR_OKCOEFAPPLIQUE'],
                                                 [pTobPiece.detail[j].GetValue('GL_AFFAIRE'),
                                                  cStVide,
                                                  cStVide,
                                                  cInVide,
                                                  cInVide,
                                                  cInVide,
                                                  PlusDate(vTobTheFormule.Getvalue('DATEREVISION'), 1, 'J'),
                                                  1,
                                                  'X'],
                                                  True);
              end;
            end;
          end;

          // si non, recherche s'il existe un coef1 puis un coef2 a l'entete
          if not assigned(vTobTheRev2) then
          begin
            vTobTheRev2 := pTobRevisionAff.FindFirst(['AFR_AFFAIRE',
                                             'AFR_NATUREPIECEG',
                                             'AFR_SOUCHE',
                                             'AFR_NUMERO',
                                             'AFR_INDICEG',
                                             'AFR_NUMORDRE',
                                             'AFR_NUMFORMULE',
                                             'AFR_OKCOEFAPPLIQUE'],
                                             [pTobPiece.detail[j].GetValue('GL_AFFAIRE'),
                                              cStVide,
                                              cStVide,
                                              cInVide,
                                              cInVide,
                                              cInVide,
                                              2,
                                              'X'],
                                              True);

            // recherche de la date
            if assigned(vTobTheRev2) then
            begin
              vTobTheFormule := pTobParamFormule.FindFirst(['AFC_FORCODE'],
                                                           [vTobTheRev2.GetValue('AFR_FORCODE')],
                                                           True);

              // recherche de la formule
              vTobTheRev2 := pTobRevisionAff.FindFirst(['AFR_AFFAIRE',
                                               'AFR_NATUREPIECEG',
                                               'AFR_SOUCHE',
                                               'AFR_NUMERO',
                                               'AFR_INDICEG',
                                               'AFR_NUMORDRE',
                                               'AFR_DATECALCCOEF',
                                               'AFR_NUMFORMULE',
                                               'AFR_OKCOEFAPPLIQUE'],
                                               [pTobPiece.detail[j].GetValue('GL_AFFAIRE'),
                                                cStVide,
                                                cStVide,
                                                cInVide,
                                                cInVide,
                                                cInVide,
                                                vTobTheFormule.Getvalue('DATEREVISION'),
                                                2,
                                                'X'],
                                                True);

              // bug si la date est 31 a l'origine elle peut etre a 30 ensuite
              if vTobTheRev2 = nil then
              begin
                decodeDate(vTobTheFormule.GetValue('DATEREVISION'), Year, Month, Day);
                if Day = 30 then
                  vTobTheRev2 := pTobRevisionAff.FindFirst(['AFR_AFFAIRE',
                                                            'AFR_NATUREPIECEG',
                                                            'AFR_SOUCHE',
                                                            'AFR_NUMERO',
                                                            'AFR_INDICEG',
                                                            'AFR_NUMORDRE',
                                                            'AFR_DATECALCCOEF',
                                                            'AFR_NUMFORMULE',
                                                            'AFR_OKCOEFAPPLIQUE'],
                                                            [pTobPiece.detail[j].GetValue('GL_AFFAIRE'),
                                                             cStVide,
                                                             cStVide,
                                                             cInVide,
                                                             cInVide,
                                                             cInVide,
                                                             PlusDate(vTobTheFormule.Getvalue('DATEREVISION'), 1, 'J'),
                                                             2,
                                                             'X'],
                                                             True);
              end;
            end;
          end;

          // si oui
          // initialisation de la clé
          // calcul de la valeur remise avec coef
          // calcul du prix unitaire
          vRdRemise := 0;
          vRdPrixHt := 0;
          vRdPrixHtDev := 0;
          vRdPrixOrigine := 0;
          vRdPrixOrigineDev := 0;

          if assigned(vTobTheRev1) then
          begin
            // creation de l'enregistrement dans la tob revision
            vTobFor1 := tob.create('AFREVISION', nil, -1);
            vTobFor1.Dupliquer(vTobTheRev1, True, True);
            vTobFor1.PutValue('AFR_NUMEROLIGNE', vInNumRev);
            vTobFor1.ChangeParent(pTobRevision, -1);

            // lien avec la tob des pieces
            if pTobPiece.detail[j].GetValue('GL_FORCODE1') = '' then
              pTobPiece.detail[j].PutValue('GL_FORCODE1', vTobTheRev1.getvalue('AFR_FORCODE'));

            pTobPiece.detail[j].PutValue('GLLAFFAIRE1', vTobTheRev1.getvalue('AFR_AFFAIRE'));
            pTobPiece.detail[j].PutValue('GLLFORCODE1', vTobTheRev1.getvalue('AFR_FORCODE'));
            pTobPiece.detail[j].PutValue('GLLNUMEROLIGNE1', vInNumRev);

            vRdRemise     := pTobPiece.detail[j].GetValue('GL_VALEURREMDEV') * vTobTheRev1.getvalue('AFR_COEFAPPLIQUE');
            vRdPrixHt     := pTobPiece.detail[j].GetValue('GL_PUHT') * vTobTheRev1.getvalue('AFR_COEFAPPLIQUE');
            vRdPrixHtDev  := pTobPiece.detail[j].GetValue('GL_PUHTDEV') * vTobTheRev1.getvalue('AFR_COEFAPPLIQUE');

            vRdPrixOrigine :=  pTobPiece.detail[j].GetValue('GL_PUHT');
            vRdPrixOrigineDev := pTobPiece.detail[j].GetValue('GL_PUHTDEV');

            vTobTheRegul := vTobRegularise.FindFirst(['AFE_FORCODE'], [vTobTheRev1.GetValue('AFR_FORCODE')], True);
            if vTobTheRegul.GetValue('REGULARISE') = 'X' then
              pTobPiece.detail[j].PutValue('GL_REGULARISE', 'X')
            else
              pTobPiece.detail[j].PutValue('GL_REGULARISE', '-');

            vInNumRev := vInNumRev + 1;
          end;

          if assigned(vTobTheRev2) then
          begin
            // creation de l'enregistrement dans la tob revision
            vTobFor2 := tob.create('AFREVISION', nil, -1);
            vTobFor2.Dupliquer(vTobTheRev2, True, True);
            vTobFor2.PutValue('AFR_NUMEROLIGNE', vInNumRev);
            vTobFor2.ChangeParent(pTobRevision, -1);

            // lien avec la tob des pieces
            if pTobPiece.detail[j].GetValue('GL_FORCODE2') = '' then
              pTobPiece.detail[j].PutValue('GL_FORCODE2', vTobTheRev2.getvalue('AFR_FORCODE'));

            pTobPiece.detail[j].PutValue('GLLAFFAIRE2', vTobTheRev2.getvalue('AFR_AFFAIRE'));
            pTobPiece.detail[j].PutValue('GLLFORCODE2', vTobTheRev2.getvalue('AFR_FORCODE'));
            pTobPiece.detail[j].PutValue('GLLNUMEROLIGNE2', vInNumRev);

            if vRdRemise = 0 then
              vRdRemise := pTobPiece.detail[j].GetValue('GL_VALEURREMDEV') * vTobTheRev2.getvalue('AFR_COEFAPPLIQUE')
            else
              vRdRemise := vRdRemise * vTobTheRev2.getvalue('AFR_COEFAPPLIQUE');

            if vRdPrixHt = 0 then
              vRdPrixHt := pTobPiece.detail[j].GetValue('GL_PUHT') * vTobTheRev2.getvalue('AFR_COEFAPPLIQUE')
            else
              vRdPrixHt := vRdPrixHt * vTobTheRev2.getvalue('AFR_COEFAPPLIQUE');

            if vRdPrixHtDev = 0 then
              vRdPrixHtDev := pTobPiece.detail[j].GetValue('GL_PUHTDEV') * vTobTheRev2.getvalue('AFR_COEFAPPLIQUE')
            else
              vRdPrixHtDev := vRdPrixHtDev * vTobTheRev2.getvalue('AFR_COEFAPPLIQUE');

            if vRdPrixOrigine = 0 then
              vRdPrixOrigine := pTobPiece.detail[j].GetValue('GL_PUHT')
            else
              vRdPrixOrigine := vRdPrixOrigine;

            if vRdPrixOrigineDev = 0 then
              vRdPrixOrigineDev := pTobPiece.detail[j].GetValue('GL_PUHTDEV')
            else
              vRdPrixOrigineDev := vRdPrixOrigineDev;

            vInNumRev := vInNumRev + 1;
          end;
 
          if assigned(vTobTheRev1) or assigned(vTobTheRev2) then
          begin
            pTobPiece.detail[j].PutValue('GL_VALEURREMDEV', vRdRemise);
            pTobPiece.detail[j].PutValue('GL_PUHT', vRdPrixHt);
            pTobPiece.detail[j].PutValue('GL_PUHTDEV', vRdPrixHtDev);
            pTobPiece.detail[j].PutValue('GL_PUHTORIGINE', vRdPrixOrigine);
            pTobPiece.detail[j].PutValue('GL_PUHTORIGINEDEV', vRdPrixOrigineDev);
          end;
        end; // end if
      end; // end if regul
    end; // end for

  Finally
    vTobRegularise.Free;
  end;

end;
}

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 25/04/2002
Modifié le ... :
Description .. : Recupere les lignes de factures et les complète
                 et alimente la TobRevision
Mots clefs ... :
*****************************************************************}
function LectureCoefFacture(pTobPiece, pTobRevision : Tob; pDtEcheance : TdateTime; pStLogFile : String; pBoLignes : Boolean) : Boolean;
var
  vRev                  : TCALCULCOEF;
  vSLLog                : TStringList;
  i                     : Integer;
  vTobParamFormule      : Tob;
  vTobRevision          : Tob;
  vTobRevisionCalc      : Tob;
  vTobParamFormuleCalc  : Tob;
  vQr                   : TQuery;
  vSt                   : String;
  vTobFille             : Tob;
  vStForCode            : String;
  vInResult             : Integer;
  
begin

  result := True;
  DebutLog(pStLogFile, vSLLog);
  vTobRevision := Tob.create('AFREVISION', nil, -1);

  // chargement de tous les coef calculés
  vSt := 'SELECT * FROM AFREVISION ';
  vSt := vSt + ' WHERE AFR_AFFAIRE = "' + pTobPiece.detail[0].GetValue('GL_AFFAIRE') + '"';
  vSt := vSt + ' AND AFR_COEFREGUL = "-"';

  if pBoLignes then
  begin
    vSt := vSt + ' AND AFR_NATUREPIECEG <> "..."';
    vSt := vSt + ' AND AFR_SOUCHE <> "..."';
  end
  else
  begin
    vSt := vSt + ' AND AFR_NATUREPIECEG = "..."';
    vSt := vSt + ' AND AFR_SOUCHE = "..."';
  end;

  vSt := vSt + ' ORDER BY AFR_DATECALCCOEF DESC';

  vQr := OpenSql(vSt, True,-1,'',true);

  Try
    vInResult := LoadParamFormule(vTobParamFormule, pTobPiece.detail[0].GetValue('GL_AFFAIRE'), pDtEcheance);
    if vInResult = 1 then
    begin
      vTobRevision.LoadDetailDB('AFREVISION','','', vQr, False, True);

      for i := 0 to vTobParamFormule.detail.count - 1 do
      begin

        // Cas 1 : recherche du coefficient appliqué
        if not CoefCalculeEtApplique(vTobParamFormule.detail[i], vTobRevision) then
          if CoefCalculeNonApplique(vTobParamFormule.detail[i], vTobRevision) then
          begin
            // cas 2 : coefficient applicable
            if vTobParamFormule.detail[i].GetValue('AFC_APPLIQUERCOEF') = 'X' then
              // Appliquer Coefficient
            begin
              vRev := TCALCULCOEF.Create;
              try
                if not vRev.LoadFormule(pTobPiece.detail[0].GetValue('GL_AFFAIRE'), vTobParamFormule.detail[i].GetValue('AFC_FORCODE')) then
                begin
                  result := False;
                  vSLLog.Add(format(TexteMsg[1], [DateToStr(now),
                                                  pTobPiece.detail[0].GetValue('GL_AFFAIRE'),
                                                  vTobParamFormule.detail[i].GetValue('AFC_FORCODE')]));
                  break;
                end
                else
                  vRev.AppliquerCoef(false, True);

              finally
                vRev.Free;
              end;
            end

            // cas 3 : coefficient non applicable
            else
            begin
              result := false;
              vSLLog.Add(format(TexteMsg[2],
                                [DateToStr(now),
                                 pTobPiece.detail[0].GetValue('GL_AFFAIRE'),
                                 vTobParamFormule.detail[i].GetValue('AFC_FORCODE')]));
              break;
            end
          end

          else if CoefNonCalcule(vTobParamFormule.detail[i], vTobRevision) then
          begin
            if vTobParamFormule.detail[i].GetValue('AFC_APPLIQUERCOEF') = 'X' then

            // cas 4 : calculé et appliquer coefficient
            begin
              vRev := TCALCULCOEF.Create;
              vTobRevisionCalc := Tob.Create('Revision mere', nil, -1);
              vTobParamFormuleCalc := Tob.create('Liste des Formules', nil, -1);
              try
                theTob := TOB.Create('Retour des saisies ',nil,-1);
                vTobFille :=TOB.Create('Une Ligne de la grille',theTob,-1);
                vTobFille.AddChampSupValeur('FORMULE',vTobParamFormule.detail[i].GetValue('AFC_FORCODE')) ;
                vTobFille.AddChampSupValeur('PROCHAINE', 'X') ;
                vTobFille.AddChampSupValeur('RECALCUL', 'X') ;

                if not vRev.CalculCoefFacturation(pTobPiece.detail[0].GetValue('GL_AFFAIRE'), theTob, vTobRevisionCalc, vTobParamFormuleCalc, True) then
                begin
                  result := false;
                  vSLLog.Add(format(TexteMsg[3], [DateToStr(now),
                                      pTobPiece.detail[0].GetValue('GL_AFFAIRE'),
                                      vTobParamFormule.detail[i].GetValue('AFC_FORCODE')]));
                  break;
                end
                else
                // met a jour les tob vTobRevision et vTobParamFormule
                // avec les données qui viennent d'etre calculées
                begin
                  AddCoefCalc(vTobRevision, vTobRevisionCalc);
                  UpdateParamFormule(vTobParamFormule, vTobParamFormuleCalc);
                  // le coef calculé n'est pas le bon, on ne stoppe pas pour
                  // calculer le coeff de toutes les formules
                  LoadParamFormule(vTobParamFormule, pTobPiece.detail[0].GetValue('GL_AFFAIRE'), pDtEcheance);
                  if not CoefCalculeEtApplique(vTobParamFormule.detail[i], vTobRevision) then
                  begin
                    result := false;
                    if result then vSLLog.Add(format(TexteMsg[4], [datetostr(now),
                                                                   pTobPiece.detail[0].GetValue('GL_AFFAIRE')]));
                  end
                end;

              finally
                vRev.Free;
                theTob.Free;
                theTob := nil;
                vTobRevisionCalc.free;
                vTobParamFormuleCalc.free;
              end
            end

            // cas 5 : Coef non applicable
            else
            begin
              result := false;
              vSLLog.Add(format(TexteMsg[5], [DateToStr(now),
                                              pTobPiece.detail[0].GetValue('GL_AFFAIRE')]));
              break;
            end;
          end;
      end;

      // si le calcul des coefficients est ok
      if result then
      begin

        // revision des lignes
        if pBoLignes then
        begin
          // creation des lignes dans AFRevision
          // vTobRevision : données de revisions de l'affaire
          // pTobRevision : données de revision du document à créer
          if PrepareLignesRevision(pTobPiece, pTobRevision, vTobRevision, vTobParamFormule,
                                pDtEcheance, pBolignes) then
            vSLLog.Add(format(TexteMsg[6], [DateToStr(now),
                                            pTobPiece.detail[0].GetValue('GL_AFFAIRE'),
                                            datetostr(pDtEcheance)]))
          else
          begin
            vSLLog.Add(format(TexteMsg[8], [DateToStr(now),
                                            pTobPiece.detail[0].GetValue('GL_AFFAIRE'),
                                            datetostr(pDtEcheance)]));

            result := false;
          end;
        end
        //revision du document
        // à développer ...
        else
        begin

        end;
      end;
    end
     
    // on vérifie si il y a une formule a l'entete ou dans les lignes
    // si oui, la formule n'a pas été paramètrée
    else if vInResult = -1 then
    begin
      vStForCode := '';
      if AffaireExisteFormule(pTobPiece.detail[0].GetValue('GL_AFFAIRE'), vStForCode) then
      begin
        vSLLog.Add(format(TexteMsg[7], [DateToStr(now),
                                        pTobPiece.detail[0].GetValue('GL_AFFAIRE'),
                                        vStForCode]));
        result := False;
      end;
    end

    // la date de révision est inférieure à la première date de révision
    else if vInResult = -2 then
    begin
        vSLLog.Add(format(TexteMsg[9], [DateToStr(now),
                                        pTobPiece.detail[0].GetValue('GL_AFFAIRE')]));
        result := False;
    end;

  finally
    vTobParamFormule.free;
    vTobRevision.free;
    ferme(vQr);
    FinLog(pStLogFile, vSLLog);
  end;
end;

function AffaireExisteFormule(pStAffaire : String; var pStForCode : String) : Boolean;
var
  vQr : TQuery;
  vSt : String;

begin
  result := True;

  // chargement de tous les coef calculés
  vSt := 'SELECT AFF_FORCODE1, AFF_FORCODE2, GL_FORCODE1, GL_FORCODE2 ';
  vSt := vSt + ' FROM AFFAIRE LEFT OUTER JOIN LIGNE ON AFF_AFFAIRE = GL_AFFAIRE ';
  vSt := vSt + ' WHERE AFF_AFFAIRE = "' + pStAffaire + '"';
 
  vQr := nil;
  Try
    vQr := OpenSql(vSt, True,-1,'',true);
    if (vQr.FindField('AFF_FORCODE1').AsString = '') and
       (vQr.FindField('AFF_FORCODE2').AsString = '') and
       (vQr.FindField('GL_FORCODE1').AsString = '') and
       (vQr.FindField('GL_FORCODE2').AsString = '') then
      result := False
    else
    begin
      if vQr.FindField('AFF_FORCODE1').AsString <> '' then pStForCode := vQr.FindField('AFF_FORCODE1').AsString
      else if vQr.FindField('AFF_FORCODE2').AsString <> '' then pStForCode := vQr.FindField('AFF_FORCODE2').AsString
      else if vQr.FindField('GL_FORCODE1').AsString <> '' then pStForCode := vQr.FindField('GL_FORCODE1').AsString
      else if vQr.FindField('GL_FORCODE2').AsString <> '' then pStForCode := vQr.FindField('GL_FORCODE2').AsString;
    end;
  Finally
    ferme(vQr);
  End;
end;

// prise en compte des periodicités
function CoefCalculeEtApplique(pTobFormule, pTobRevision : Tob) : Boolean;
var
  vTob            : Tob;
begin
    
  if pTobFormule.Getvalue('COEFCALCULE') = 'X' then
  begin
    // si coef calculé, on recherche la derniere date applique
    vTob := pTobRevision.FindFirst(['AFR_AFFAIRE',
                                    'AFR_FORCODE',
                                    'AFR_DATECALCCOEF',
                                    'OKCOEFAPPLIQUE'],
                                    [pTobFormule.GetValue('AFC_AFFAIRE'),
                                     pTobFormule.GetValue('AFC_FORCODE'),
                                     pTobFormule.GetValue('AFC_LASTDATEAPP'),
                                    'X'], True);
    if vTob <> nil then
     result := True
    else
      result := False;
  end
  else
    result := false;
end;

function CoefCalculeNonApplique(pTobFormule, pTobRevision : Tob) : Boolean;
var
  vTob : Tob;
begin

  if pTobFormule.Getvalue('COEFCALCULE') = 'X' then
  begin
    vTob := pTobRevision.FindFirst(['AFR_AFFAIRE',
                                    'AFR_FORCODE',
                                    'AFR_DATECALCCOEF',
                                    'AFR_OKCOEFAPPLIQUE'],
                                    [pTobFormule.GetValue('AFC_AFFAIRE'),
                                     pTobFormule.GetValue('AFC_FORCODE'),
                                     pTobFormule.GetValue('AFC_LASTDATECALC'),
                                     '-'], True);
    if vTob <> nil then
      result := True
    else
      result := false;
  end
  else
    result := false;
end;

// le coef n'est pas calculé si il n'a jamais été calculé ou
// si le dernier calculé et appliqué n'est pas celui attendu
function CoefNonCalcule(pTobFormule, pTobRevision : Tob) : Boolean;
var
  vTob : Tob;

begin
  if pTobFormule.Getvalue('COEFCALCULE') = 'X' then
  begin
    // si coef calculé, on recherche la derniere date applique
    vTob := pTobRevision.FindFirst(['AFR_AFFAIRE',
                                    'AFR_FORCODE',
                                    'AFR_DATECALCCOEF',
                                    'OKCOEFAPPLIQUE'],
                                    [pTobFormule.GetValue('AFC_AFFAIRE'),
                                     pTobFormule.GetValue('AFC_FORCODE'),
                                     pTobFormule.GetValue('AFC_LASTDATEAPP'),
                                    'X'], True);
    if vTob <> nil then
      result := False
    else
      result := True;
  end
  else
    result := True;
end;

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 17/06/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure AddCoefCalc(pTobRevision, pTobRevisionCalc : Tob);
var i : Integer;
begin
  if assigned(pTobRevisionCalc) then
    for i := pTobRevisionCalc.detail.count -1 to 0 do
      pTobRevisionCalc.detail[i].ChangeParent(pTobRevision, 0);
end;

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 17/06/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure UpdateParamFormule(pTobParamFormule, pTobParamFormuleCalc : Tob);
var
  vTobParamFormule : Tob;

begin
  if assigned(pTobParamFormuleCalc) then
    vTobParamFormule := pTobParamFormule.FindFirst(['AFC_AFFAIRE',
                                'AFC_FORCODE'],
                                [pTobParamFormuleCalc.GetValue('AFC_AFFAIRE'),
                                pTobParamFormuleCalc.GetValue('AFC_FORCODE')],
                                True)
  else
    vTobParamFormule := nil;

  if vTobParamFormule <> nil then
  begin
    vTobParamFormule.PutValue('AFC_LASTDATEAPP', pTobParamFormuleCalc.GetValue('AFC_LASTDATEAPP'));
    vTobParamFormule.PutValue('AFC_LASTDATECALC', pTobParamFormuleCalc.GetValue('AFC_LASTDATECALC'));
    vTobParamFormule.PutValue('AFC_NEXTDATECALC', pTobParamFormuleCalc.GetValue('AFC_NEXTDATECALC'));
    vTobParamFormule.PutValue('AFC_NEXTDATEAPP', pTobParamFormuleCalc.GetValue('AFC_NEXTDATEAPP'));
    vTobParamFormule.PutValue('COEFCALCULE', '-');
  end;
end;
{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 30/07/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
function NumPieceAffaire(pStAffaire : String) : Integer;
var
  vQr : TQuery;
  vSt : String;
Begin
  vSt := 'SELECT GP_NUMERO FROM PIECE WHERE GP_AFFAIRE = "' + pStAffaire + '"';
  vQr := nil;
  vQR := OpenSql(vSt, True,-1,'',true);
  try
    if Not vQR.Eof then
      result := vQr.FindField('GP_NUMERO').AsInteger
    else
      result := 0;
  Finally
    if vQR <> nil then ferme(vQR);
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : CB
Créé le ...... : 25/04/2002
Modifié le ... :
Description .. : Recupere les lignes de factures et les complète
                 et alimente la TobRevision
Mots clefs ... :
*****************************************************************}
function LectureCoefFactureRegul(pTobPiece, pTobRevision : Tob; pDtEcheance : TdateTime; var pInNumRev : Integer; pBoLignes : Boolean) : Boolean;
var
  vSLLog                : TStringList;
  vTobParamFormule      : Tob;
  vTobRevision          : Tob;
  vQr                   : TQuery;
  vSt                   : String;
  vInResult             : Integer;

begin

  result := True;
  DebutLog('', vSLLog);
  vTobRevision := Tob.create('AFREVISION', nil, -1);

  // chargement de tous les coef calculés
  vSt := 'SELECT * FROM AFREVISION ';
  vSt := vSt + ' WHERE AFR_AFFAIRE = "' + pTobPiece.detail[0].GetValue('GL_AFFAIRE') + '"';
  vSt := vSt + ' AND AFR_COEFREGUL = "X"';

  if pBoLignes then
  begin
    vSt := vSt + ' AND AFR_NATUREPIECEG <> "..."';
    vSt := vSt + ' AND AFR_SOUCHE <> "..."';
  end
  else
  begin
    vSt := vSt + ' AND AFR_NATUREPIECEG = "..."';
    vSt := vSt + ' AND AFR_SOUCHE = "..."';
  end;

  vSt := vSt + ' ORDER BY AFR_DATECALCCOEF DESC';

  vQr := OpenSql(vSt, True,-1,'',true);
  Try
    vInResult := LoadParamFormule(vTobParamFormule, pTobPiece.detail[0].GetValue('GL_AFFAIRE'), pDtEcheance);
    if vInResult = 1 then
    begin
      vTobRevision.LoadDetailDB('AFREVISION','','', vQr, False, True);

      if pBoLignes then
      begin
        // creation des lignes dans AFRevision
        // vTobRevision : données de revisions de l'affaire
        // pTobRevision : données de revision du document à créer
        PrepareLignesRevisionRegul(pTobPiece, pTobRevision, vTobRevision,
                              vTobParamFormule, pDtEcheance, pInNumRev, pBoLignes);
        vSLLog.Add(format(TexteMsg[6], [DateToStr(now), pTobPiece.detail[0].GetValue('GL_AFFAIRE'), datetostr(pDtEcheance)]));
      end
      else
      begin

      end;
    end

    else if vInResult = -1 then
    begin
        vSLLog.Add(format(TexteMsg[7], [DateToStr(now),
                                        pTobPiece.detail[0].GetValue('GL_AFFAIRE'),
                                        '']));
    end
    else if vInResult = -2 then
    begin
        vSLLog.Add(format(TexteMsg[9], [DateToStr(now),
                                        pTobPiece.detail[0].GetValue('GL_AFFAIRE')]));
    end;

  finally
    vTobParamFormule.free;
    vTobRevision.free;
    ferme(vQr);
    FinLog('', vSLLog);
  end;
end;

end.
