{***********UNITE*************************************************
Auteur  ...... : Olivier TARCY
Créé le ...... : 16/04/2002
Modifié le ... : 22/04/2002
Description .. : Focntions utilisées pour l'interface comptable
Mots clefs ... : BOS5;INTERCOMPTA
*****************************************************************}
unit UtilInterCompta;

interface

uses {$IFDEF VER150} variants,{$ENDIF} HEnt1, HCtrls, UTOB, Controls, ComCtrls, StdCtrls, Ent1,
  {$IFNDEF EAGLCLIENT}
  DB, {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
  {$ENDIF}
  VentilCpta, SysUtils, ParamSoc, EntGC;

function ValeurAxeAna(TobAna: TOB): TOB;
function CalculIteration(CodeAxe, CodeType, CodeEtab, CodeStruct: string): integer;
function LibChamp(Champ, Code: string): string;
function IsNatCaisse(NatureATester: string) : boolean;

implementation

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Olivier TARCY
Créé le ...... : 16/04/2002
Modifié le ... : 22/04/2002
Description .. : Retourne les valeurs des axes analytiques pour chaque
Suite ........ : pièce comptable
Mots clefs ... : BOS5;INTERCOMPTA
*****************************************************************}

function ValeurAxeAna(TobAna: TOB): TOB;
var CodeAxe, CodeType, CodeEtab, CodeStruct, CodeValeur, CodeSav, Valeur, valeur1, valeur2, stSQL, isLibelle: string;
  i_tob, i_cstte, i_compteur, i1, i2, i_compteur1, iLSE, i_compteur2, iSEC, {NumA,} Longueur: integer;
  OkCstte: boolean;
  TobTemp, TobMerge, TobCstte: TOB;
  Q1: TQuery;
begin
  //gcvoirtob(tobana) ;
  if (TobAna <> nil) and (TobAna.Detail.count > 0) then
  begin
    TobTemp := TOB.Create('Merge', nil, -1);
    i_tob := 0;
    i_cstte := 0;
    CodeSav := '';
    {NumA := 1;}
    Valeur := '';

    //
    // Première lecture
    //
    CodeAxe := TobAna.Detail[i_tob].GetValue('GDA_AXE');
    CodeType := TobAna.Detail[i_tob].GetValue('GDA_TYPECOMPTE');
    CodeEtab := TobAna.Detail[i_tob].GetValue('GDA_ETABLISSEMENT');
    CodeStruct := TobAna.Detail[i_tob].GetValue('GDA_TYPESTRUCTANA');

    //
    // chargement de l'axe en TOB pour pouvoir intégrer les constantes si nécessaire
    //
    stSQL := 'select * from DECOUPEANA where ' +
      'GDA_AXE="' + CodeAxe + '" AND ' +
      'GDA_TYPECOMPTE="' + CodeType + '" AND ' +
      'GDA_ETABLISSEMENT="' + CodeEtab + '" ' +
      'order by GDA_AXE,GDA_TYPECOMPTE,GDA_ETABLISSEMENT,GDA_TYPESTRUCTANA';
    TobCstte := Tob.Create('Cstte', nil, -1);
    Q1 := OpenSQL(stSQL, True);
    TobCstte.LoadDetailDB('DECOUPEANA', '', '', Q1, False);
    Ferme(Q1);

    //
    // Initialisation des compteurs
    //
    i1 := CalculIteration(CodeAxe, CodeType, CodeEtab, 'SEC');
    i2 := CalculIteration(CodeAxe, CodeType, CodeEtab, 'LSE');
    i_compteur := i1 + i2;
    i_compteur1 := i1;
    iLSE := 0;
    i_compteur2 := i2;
    iSEC := 0;

    //
    //  Construction des axes analytiques en TOB
    //
    repeat
      if (i_tob > TobAna.Detail.Count - 1) then i_tob := TobAna.Detail.Count - 1;
      CodeAxe := TobAna.Detail[i_tob].GetValue('GDA_AXE');
      CodeType := TobAna.Detail[i_tob].GetValue('GDA_TYPECOMPTE');
      CodeEtab := TobAna.Detail[i_tob].GetValue('GDA_ETABLISSEMENT');
      Longueur := TobAna.Detail[i_tob].GetValue('GDA_LONGUEUR');
      isLibelle := TobAna.Detail[i_tob].GetValue('GDA_ISLIBELLE');
      CodeStruct := TobAna.Detail[i_tob].GetValue('GDA_TYPESTRUCTANA');

      OkCstte := False;
      if TobCstte.Detail[i_cstte].GetValue('GDA_LIBCHAMP') <> TobAna.Detail[i_tob].GetValue('GDA_LIBCHAMP')
        then OkCstte := true;

      if (not OkCstte) and (TobAna.Detail[i_tob].GetValue('VALEUR') <> #0) then
        CodeValeur := TobAna.Detail[i_tob].GetValue('VALEUR') else
        CodeValeur := '';

      if isLibelle = 'X' then
      begin
        // ce n'est pas une constante
        if not OkCstte then
        begin
          if (CodeValeur = Null) then isLibelle := ''
          else isLibelle := LibChamp(TobAna.Detail[i_tob].GetValue('GDA_LIBCHAMP'), CodeValeur);
          {if CodeAxe = 'A1' then NumA := 1 else NumA := 2;}
          if ExisteSQL('SELECT GDA_TYPESTRUCTANA FROM DECOUPEANA WHERE GDA_AXE="' + CodeAxe +
            '" AND ' + 'GDA_TYPECOMPTE="' + CodeType + '" AND ' +
            'GDA_ETABLISSEMENT="' + CodeEtab + '" AND ' + 'GDA_TYPESTRUCTANA="' + CodeStruct + '"') then
            //isLibelle := BourreChamps (isLibelle, Longueur, TFichierBase(Ord(fbAxe1)+NumA-1), True) ;
            TobAna.Detail[i_tob].PutValue('VALEUR', isLibelle);
        end else
          // c'est une constante
        begin
          isLibelle := TobCstte.Detail[i_cstte].GetValue('GDA_LIBCHAMP');
          //isLibelle := BourreChamps (isLibelle, TobCstte.Detail [ i_cstte ].GetValue ( 'GDA_LONGUEUR' ), TFichierBase(Ord(fbAxe1)+NumA-1), True) ;
        end;
      end;

      //
      // la concaténation est terminée
      //
      if CodeStruct = 'SEC' then
        if not OkCstte then
        begin
          if TobAna.Detail[i_tob].GetValue('VALEUR') = Null then
            TobAna.Detail[i_tob].PutValue('VALEUR', '')
          else TobAna.Detail[i_tob].PutValue('VALEUR', copy(TobAna.Detail[i_tob].GetValue('VALEUR'), 1, Longueur));
        end;
      if OkCstte then
      begin
        Longueur := TobCstte.Detail[i_cstte].GetValue('GDA_LONGUEUR');
        isLibelle := copy(TobCstte.Detail[i_Cstte].GetValue('GDA_CSTTE'), 1, Longueur);
      end;

      if OkCstte then Valeur := Valeur + isLibelle else
        if TobAna.Detail[i_tob].GetValue('VALEUR') <> #0 then
        Valeur := Valeur + TobAna.Detail[i_tob].GetValue('VALEUR');

      if CodeStruct = 'LSE' then
        if Valeur <> '' then Valeur := Valeur + ' ';

      //  incrémentation des indices du nb de copde section et libellé
      if OkCstte then
      begin
        if TobCstte.Detail[i_cstte].GetValue('GDA_TYPESTRUCTANA') = 'SEC' then
          Inc(iSEC) else Inc(iLSE);
      end else
        if CodeStruct = 'SEC' then Inc(iSEC) else Inc(iLSE);

      // test de fin de traitement de la chaine des codes section
      if ((i_compteur1 > 0) and ((iSEC) mod i_compteur1 = 0) and (iSEC > 0)) then
      begin
        valeur1 := TrimRight(Valeur);
        Valeur := '';
        i_compteur1 := 99; // pour éviter la division par zéro
      end else
        // test de fin de traitement de la chaine des codes libellé
        if ((i_compteur2 <> 0) and ((iLSE) mod i_compteur2 = 0) and (iLSE > 0)) then
      begin
        valeur2 := TrimRight(Valeur);
        Valeur := '';
        i_compteur2 := 99; // pour éviter la division par zéro
      end;
      // test de fin de traitement
      if ((i_cstte + 1) mod i_compteur = 0) then
      begin
        i_compteur1 := i1;
        iLSE := 0;
        i_compteur2 := i2;
        iSEC := 0;
        TobMerge := TOB.Create('Fille', TobTemp, -1);
        TobMerge.AddChampSupValeur('GDA_AXE', CodeAxe);
        TobMerge.AddChampSupValeur('GDA_TYPECOMPTE', CodeType);
        TobMerge.AddChampSupValeur('GDA_ETABLISSEMENT', CodeEtab);
        TobMerge.AddChampSupValeur('GDA_VALEUR1', valeur1);
        TobMerge.AddChampSupValeur('GDA_VALEUR2', valeur2);
        valeur1 := '';
        valeur2 := '';
      end;
      if ((i_cstte + 1) mod i_compteur <> 0) then Inc(i_cstte) else i_cstte := 0;
      if (i_tob = TobAna.Detail.Count - 1) and (i_cstte <> 0) then Dec(i_tob);
      if ((i_tob = TobAna.Detail.Count - 1) and (i_cstte = 0)) or (not OkCstte) then Inc(i_tob); // cas ou on a terminé par une constante
    until (i_tob > TobAna.Detail.Count - 1);
    TobAna.Dupliquer(TobTemp, True, True);
    TobTemp.Free;
    TobCstte.Free;
  end;
  //gcvoirtob(tobana) ;
  Result := TobAna;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Olivier TARCY
Créé le ...... : 09/07/2002
Modifié le ... : 09/07/2002
Description .. : Calcul des compteurs pour connaitre le nombre d'itérations à
Suite ........ :  effectuer pour les codes sections et les libellés section pour
Suite ........ : former les chaines concaténées
Mots clefs ... : BOS5;INTERCOMPTA;
*****************************************************************}

function CalculIteration(CodeAxe, CodeType, CodeEtab, CodeStruct: string): integer;
var Q1: TQuery;
  stSQL: string;
begin
  //
  // Calcul du nombre d'enregistrements consécutifs à traiter pour former la chaîne concaténée
  //
  stSQL := 'SELECT COUNT(GDA_AXE) AS COMPTEUR FROM DECOUPEANA WHERE ' +
    'GDA_AXE="' + CodeAxe + '" AND ' +
    'GDA_TYPECOMPTE="' + CodeType + '" AND ' +
    'GDA_ETABLISSEMENT="' + CodeEtab + '" AND ' +
    'GDA_TYPESTRUCTANA="' + CodeStruct + '"';
  Q1 := OpenSQL(stSQL, True);
  Result := Q1.FindField('COMPTEUR').AsInteger;
  //      Valeur := '' ;
  Ferme(Q1);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Olivier TARCY
Créé le ...... : 16/04/2002
Modifié le ... : 22/04/2002
Description .. : Retourne le libellé du champ
Mots clefs ... : BOS5;INTERCOMPTA
*****************************************************************}

function LibChamp(Champ, Code: string): string;
var Req: string;
  wlib, Lib: string;
  TobTrt: TOB;
  QQ: TQuery;
  i: integer;
begin
  Req := 'SELECT CO_CODE,CO_LIBELLE,CO_LIBRE FROM COMMUN WHERE CO_TYPE="GCS" AND ' +
    ' AND CO_LIBRE LIKE "%' + Champ + ';X%MODE%" ORDER BY CO_CODE';
  QQ := OpenSQL(Req, True);
  if QQ.EOF then exit;
  TobTrt := Tob.create('liste champs', nil, -1);
  TobTrt.LoadDetailDB('', '', '', QQ, False);
  Ferme(QQ);

  for i := 0 to TobTrt.detail.count - 1 do
  begin
    Lib := TobTrt.detail[i].GetValue('CO_LIBELLE');
    wLib := ReadTokenSt(Lib);
    if wlib <> '&#@' then
    begin
      wlib := ReadTokenSt(Lib);
      Result := RechDom(wlib, Code, False);
    end;
  end;
  TobTrt.free;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : JT
Créé le ...... : 20/06/2003
Modifié le ... : 20/06/2003
Description .. : Test par nature si PIEDECHE doit correspondre à un règlement
Mots clefs ... : INTERCOMPTA;COMPTADIFF
*****************************************************************}
function IsNatCaisse(NatureATester: string) : boolean;
begin
  if NatureATester <> '' then
    Result := (GetInfoParPiece(NatureATester,'GPP_EQUIPIECE') = 'FFO')
    else
    Result := False;
end;

end.
