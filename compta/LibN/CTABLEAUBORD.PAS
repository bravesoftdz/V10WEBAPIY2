unit CTableauBord;

interface


procedure InitialiserTableCompta;

implementation

uses
  {$IFDEF MODENT1}
  CPTypeCons,
  CPProcMetier,
  CPVersion,
  {$ENDIF MODENT1}
  SysUtils,
  {$IFNDEF EAGLCLIENT}
  {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
  {$ENDIF}
  HCtrls,
  HMsgBox,
  utob,
  Ent1,
  HEnt1,
  Calcole,
  Classes,
  ED_TOOLS,
  galOutil;


procedure MajTableauBordCompta ( bIsDossierCabinet : boolean);

type
  TB_RUBRIQUE = (CRP001, CRP000, CRP002, SIG003, SIG011, SIG010, SIG012, CRP004, CRC006,
    SIG013, SIG002, SIG014, CRC009, CRC013, CRP016, CRP007, CRC015, CRP013,
    CRC018, CRC021, CRC022);

const
  TB_INFORUBRIQUE: array[TB_RUBRIQUE] of string =
  {@CRP/001}('701:7069999999(SM);708:7088999999(SM);7091:709699999(SM);7098(SM);'
    {@CRP/000}, '707(SM);7097(SM);'
    {@CRP/002}, '713(SM);'
    {@SIG/003}, '72(SM);'
    {@SIG/011}, '607(SM);6037(SM);6087(SM);6097(SM);'
    {@SIG/010}, '601(SM);602(SM);6081(SM);6082(SM);6091(SM);6092(SM);6031(SM);6032(SM);604(SM);605(SM);6084(SM);6085(SM);6086(SM);6088(SM);6094(SM);6095(SM);6096(SM);6098(SM);611(SM);'
    {@SIG/012}, '606(SM);6122(SM);613(SM);614(SM);615(SM);616(SM);617(SM);618(SM);621(SM);6221:622899999(SM);623(SM);624(SM);625(SM);626(SM);627(SM);628(SM);629(SM);619(SM);6125(SM);'
    {@CRP/004}, '74(SM);'
    {@CRC/006}, '630:6349999999(SM);6351:6378999999(SM);'
    {@SIG/013}, '641:6418999999(SM);644(SM);642(SM);645:648999999(SM);'
    {@SIG/002}, '7811(SM);7815(SM);78161(SM);78162(SM);78173(SM);78174(SM);791(SM);'
    {@SIG/014}, '73(SM);751(SM);752(SM);753(SM);754(SM);758(SM);'
    {@CRC/009}, '6811(SM);6812(SM);6815(SM);68161(SM);68162(SM);68173(SM);68174(SM);'
    {@CRC/013}, '651(SM);652(SM);653(SM);654(SM);658(SM);'
    {@CRP/016}, '755(SM);655(SM);'
    {@CRP/007}, '761:7689999999(SM);7864(SM);7865(SM);78662:7866599999(SM);796(SM);'
    {@CRC/015}, '661:6689999999(SM);6861(SM);6864(SM);6865(SM);68662:6866599999(SM);6868(SM);'
    {@CRP/013}, '771(SM);772(SM);775(SM);777(SM);778(SM);7871(SM);7872(SM);7873(SM);7874(SM);7875(SM);7876(SM);797(SM);'
    {@CRC/018}, '671(SM);672(SM);675(SM);678(SM);6871(SM);6872(SM);6873(SM);6874(SM);6875(SM);6876(SM);'
    {@CRC/021}, '691(SM);'
    { vu avec CJ le 15/10/2003, on considère 689 et 789 en SM }
    {@CRC/022}, '689(SD);789(SC);695:697(SM);6981(SM);6989(SM);6990(SM);6995:6997(SM);');
  TB_SQLSELECT = 'SELECT SUM(G_TOTDEBP) G_TOTDEBP,SUM(G_TOTCREP) G_TOTCREP,SUM(G_TOTDEBE) G_TOTDEBE,SUM(G_TOTCREE) G_TOTCREE,SUM(G_TOTDEBS) G_TOTDEBS,SUM(G_TOTCRES) G_TOTCRES FROM GENERAUX ';
  procedure TBMajInfoExo(T: TOB);
  var
    Exo: TExoDate;
    stMillesime: string;
    i: integer;
    stLibelle : string;
  begin
    for i := 0 to 2 do
    begin
      case i of
        0:
          begin
            Exo := VH^.Precedent;
            stMillesime := 'N-';
            stLibelle := 'N-1';
          end;
        1:
          begin
            Exo := VH^.Encours;
            stMillesime := 'N';
            stLibelle := 'N';
          end;
        2:
          begin
            Exo := VH^.Suivant;
            stMillesime := 'N+';
            stLibelle := 'N+1';
          end;
      end;
      T.Detail[i].PutValue('DTC_NODOSSIER', V_PGI.NoDossier);
      T.Detail[i].PutValue('DTC_MILLESIME', stMillesime);
      T.Detail[i].PutValue('DTC_LIBEXERCICE', stLibelle);
      T.Detail[i].PutValue('DTC_DUREE', Exo.NombrePeriode);
      T.Detail[i].PutValue('DTC_DATEDEB', Exo.Deb);
      T.Detail[i].PutValue('DTC_DATEFIN', Exo.Fin);
    end;
  end;

  function WhereSQLDepuisRubrique(stRubrique: string): string;
  var
    stSQL, St: string;
    i : integer;
  begin
    stSQL := '';
    { Suppression des infos de solde dans la rubrique }
    i := Pos('(',stRubrique);
    while ( i > 0)  do
    begin
      Delete(stRubrique,i,4);
      i := Pos('(',stRubrique);
    end;
    while (stRubrique <> '') do
    begin
      St := ReadTokenSt(stRubrique);
      if stSQL <> '' then stSQL := stSQL + ' OR ';
      if Pos(':', St) > 0 then
        stSQL := stSQL + '(G_GENERAL>="' + BourreEtLess(ReadTokenPipe(St, ':'), fbGene) + '" AND G_GENERAL<="' + BourreEtLess(St, fbGene) + '")'
      else stSQL := stSQL + '(G_GENERAL LIKE "' + St + '%")';
    end;
    Result := ' WHERE ' + stSQL;
  end;

  procedure CumulDepuisRubrique ( TR : array of TB_RUBRIQUE; bDebitNeg : boolean; var CumulPrec : double; var CumulExo : double; var CumulSuiv : double);
  var i : integer;
      St, stWhere, stSQL : string;
      Q : TQuery;
  begin
    St := ''; CumulPrec := 0; CumulExo := 0; CumulSuiv := 0;
    for i := 0 to High ( TR ) do
      St := St + TB_INFORUBRIQUE[TR[i]];
    stWhere := WhereSQLDepuisRubrique(St);
    stSQL := TB_SQLSELECT + stWhere ;
    Q := OpenSQL ( stSQL, True );
    try
      if bDebitNeg then
      begin
        CumulPrec := Q.FindField('G_TOTCREP').AsFloat - Q.FindField('G_TOTDEBP').AsFloat;
        CumulExo := Q.FindField('G_TOTCREE').AsFloat - Q.FindField('G_TOTDEBE').AsFloat;
        CumulSuiv := Q.FindField('G_TOTCRES').AsFloat - Q.FindField('G_TOTDEBS').AsFloat;
      end else
      begin
        CumulPrec := Q.FindField('G_TOTDEBP').AsFloat - Q.FindField('G_TOTCREP').AsFloat;
        CumulExo := Q.FindField('G_TOTDEBE').AsFloat - Q.FindField('G_TOTCREE').AsFloat;
        CumulSuiv := Q.FindField('G_TOTDEBS').AsFloat - Q.FindField('G_TOTCRES').AsFloat;
      end;
    finally
      Ferme ( Q );
    end;
  end;

  procedure TBChiffreAffaire(T: TOB);
  var  CumulPrec, CumulExo, CumulSuiv : double;
  begin
    CumulDepuisRubrique (  [CRP001,CRP000],False,CumulPrec,CumulExo,CumulSuiv);
    T.Detail[0].PutValue('DTC_CA', -CumulPrec);
    T.Detail[1].PutValue('DTC_CA', -CumulExo);
    T.Detail[2].PutValue('DTC_CA', -CumulSuiv);
  end;

  procedure TBMargeTotale(T: TOB);
  var  CumulPrec, CumulExo, CumulSuiv : double;
  begin
    CumulDepuisRubrique (  [CRP002,SIG003,SIG011,SIG010],False,CumulPrec,CumulExo,CumulSuiv);
    T.Detail[0].PutValue('DTC_MARGETOTALE', T.Detail[0].GetValue('DTC_CA') - CumulPrec);
    T.Detail[1].PutValue('DTC_MARGETOTALE', T.Detail[1].GetValue('DTC_CA') - CumulExo);
    T.Detail[2].PutValue('DTC_MARGETOTALE', T.Detail[2].GetValue('DTC_CA') - CumulSuiv);
  end;

  procedure TBValeurAjoutee(T: TOB);
  var  CumulPrec, CumulExo, CumulSuiv : double;
  begin
    CumulDepuisRubrique (  [SIG012],False,CumulPrec,CumulExo,CumulSuiv);
    T.Detail[0].PutValue('DTC_VALEURAJOUTEE', T.Detail[0].GetValue('DTC_MARGETOTALE') - CumulPrec);
    T.Detail[1].PutValue('DTC_VALEURAJOUTEE', T.Detail[1].GetValue('DTC_MARGETOTALE') - CumulExo);
    T.Detail[2].PutValue('DTC_VALEURAJOUTEE', T.Detail[2].GetValue('DTC_MARGETOTALE') - CumulSuiv);
  end;

  procedure TBExcedentBrut(T: TOB);
  var  CumulPrec, CumulExo, CumulSuiv : double;
  begin
    CumulDepuisRubrique (  [CRP004,CRC006,SIG013],False,CumulPrec,CumulExo,CumulSuiv);
    T.Detail[0].PutValue('DTC_EXCEDBRUT', T.Detail[0].GetValue('DTC_VALEURAJOUTEE') - CumulPrec);
    T.Detail[1].PutValue('DTC_EXCEDBRUT', T.Detail[1].GetValue('DTC_VALEURAJOUTEE') - CumulExo);
    T.Detail[2].PutValue('DTC_EXCEDBRUT', T.Detail[2].GetValue('DTC_VALEURAJOUTEE') - CumulSuiv);
  end;

  procedure TBResultatExploitation(T: TOB);
  var  CumulPrec, CumulExo, CumulSuiv : double;
  begin
    CumulDepuisRubrique (  [SIG002,SIG014,CRC009,CRC013],False,CumulPrec,CumulExo,CumulSuiv);
    T.Detail[0].PutValue('DTC_RESULTEXPLOIT', T.Detail[0].GetValue('DTC_EXCEDBRUT') - CumulPrec);
    T.Detail[1].PutValue('DTC_RESULTEXPLOIT', T.Detail[1].GetValue('DTC_EXCEDBRUT') - CumulExo);
    T.Detail[2].PutValue('DTC_RESULTEXPLOIT', T.Detail[2].GetValue('DTC_EXCEDBRUT') - CumulSuiv);
  end;

  procedure TBResultatCourant(T: TOB);
  var  CumulPrec, CumulExo, CumulSuiv : double;
  begin
    CumulDepuisRubrique (  [CRP016,CRP007,CRC015],False,CumulPrec,CumulExo,CumulSuiv);
    T.Detail[0].PutValue('DTC_RESULTCOURANT', T.Detail[0].GetValue('DTC_RESULTEXPLOIT') - CumulPrec);
    T.Detail[1].PutValue('DTC_RESULTCOURANT', T.Detail[1].GetValue('DTC_RESULTEXPLOIT') - CumulExo);
    T.Detail[2].PutValue('DTC_RESULTCOURANT', T.Detail[2].GetValue('DTC_RESULTEXPLOIT') - CumulSuiv);
  end;

  procedure TBResultatExercice(T: TOB);
  var  CumulPrec, CumulExo, CumulSuiv : double;
  begin
    CumulDepuisRubrique (  [CRP013,CRC018,CRC021,CRC022],False,CumulPrec,CumulExo,CumulSuiv);
    T.Detail[0].PutValue('DTC_RESULTEXERC', T.Detail[0].GetValue('DTC_RESULTCOURANT') - CumulPrec);
    T.Detail[1].PutValue('DTC_RESULTEXERC', T.Detail[1].GetValue('DTC_RESULTCOURANT') - CumulExo);
    T.Detail[2].PutValue('DTC_RESULTEXERC', T.Detail[2].GetValue('DTC_RESULTCOURANT') - CumulSuiv);
  end;

  procedure TBSubventionAgricole(T: TOB);
  var  Q        : TQuery;
       stChamp  : string;
       stSQL    : string;
  begin
    stSQL := 'SELECT SUBSTRING(G_GENERAL,1,4) COMPTE, ';
    stSQL := stSQL + ' SUM(G_TOTCREP) TOTCREP,SUM(G_TOTDEBP) TOTDEBP,SUM(G_TOTCREE) TOTCREE,SUM(G_TOTDEBE) TOTDEBE,SUM(G_TOTCRES) TOTCRES ,SUM(G_TOTDEBS) TOTDEBS ';
    stSQL := stSQL + ' FROM GENERAUX ';
    stSQL := stSQL + ' WHERE G_GENERAL LIKE "7451%" OR G_GENERAL LIKE "7452%" OR G_GENERAL LIKE "7453%" OR G_GENERAL LIKE "7458%" ';
    stSQL := stSQL + ' GROUP BY SUBSTRING(G_GENERAL,1,4) ';
    Q := OpenSQL (stSQL, True,-1,'',True);
    try
      while not Q.Eof do
      begin
        stChamp := '';
        if (Q.FindField('COMPTE').AsString = '7451') then stChamp := 'DTC_SUBVPRO'
        else if (Q.FindField('COMPTE').AsString = '7452') then stChamp := 'DTC_SUBVSTR'
        else if (Q.FindField('COMPTE').AsString = '7453') then stChamp := 'DTC_SUBVREV'
        else if (Q.FindField('COMPTE').AsString = '7458') then stChamp := 'DTC_SUBVAUT';
        if stChamp <> '' then
        begin
          T.Detail[0].PutValue(stChamp, Q.FindField('TOTDEBP').AsFloat - Q.FindField('TOTCREP').AsFloat );
          T.Detail[1].PutValue(stChamp, Q.FindField('TOTDEBE').AsFloat - Q.FindField('TOTCREE').AsFloat);
          T.Detail[2].PutValue(stChamp, Q.FindField('TOTDEBS').AsFloat - Q.FindField('TOTCRES').AsFloat);
        end;
        Q.Next;
      end;
    finally
      Ferme (Q);
    end;
  end;

  procedure TBNombreEcriture(T: TOB);
  var
    stSQL: string;
    Q: TQuery;
  begin
    stSQL := 'SELECT E_EXERCICE,COUNT(*) NBECRITURE FROM ECRITURE WHERE E_QUALIFPIECE="N" AND (E_EXERCICE="' + VH^.Precedent.Code + '" OR ' +
      ' E_EXERCICE="' + VH^.EnCours.Code + '" OR E_EXERCICE="' + VH^.Suivant.Code + '") GROUP BY E_EXERCICE';
    Q := OpenSQL(stSQL, True);
    if not Q.Eof then
    begin
      while not Q.Eof do
      begin
        if Q.FindField('E_EXERCICE').AsString = VH^.Precedent.Code then
          T.Detail[0].PutValue('DTC_NBECRITURE', Q.FindField('NBECRITURE').AsInteger)
        else if Q.FindField('E_EXERCICE').AsString = VH^.Encours.Code then
          T.Detail[1].PutValue('DTC_NBECRITURE', Q.FindField('NBECRITURE').AsInteger)
        else if Q.FindField('E_EXERCICE').AsString = VH^.Suivant.Code then
          T.Detail[2].PutValue('DTC_NBECRITURE', Q.FindField('NBECRITURE').AsInteger);
        Q.Next;
      end;
    end;
    Ferme(Q);
  end;

  procedure TBInfoSaisie(T: TOB);

    procedure _TBInfoDerniereSaisie (stSQL : string; Exo : string; Indice : integer; T : TOB );
    var lQ : TQuery;
    begin
      stSQL := StringReplace(stSQL , '&#@', Exo,[rfReplaceAll, rfIgnoreCase]);
      lQ := OpenSQL (stSQL, True);
      try
        if not lQ.Eof then
        begin
          T.Detail[Indice].PutValue('DTC_DERNSAISIE', lQ.FindField('E_DATECOMPTABLE').AsDateTime);
          T.Detail[Indice].PutValue('DTC_DATESAISIE', lQ.FindField('E_DATEMODIF').AsDateTime);
          T.Detail[Indice].PutValue('DTC_DERNJOURNAL', lQ.FindField('E_JOURNAL').AsString);
          T.Detail[Indice].PutValue('DTC_UTILISATSAISIE', lQ.FindField('E_UTILISATEUR').AsString);
        end else
        begin
          T.Detail[Indice].PutValue('DTC_DERNSAISIE', iDate1900);
          T.Detail[Indice].PutValue('DTC_DATESAISIE', iDate1900);
          T.Detail[Indice].PutValue('DTC_DERNJOURNAL', '');
          T.Detail[Indice].PutValue('DTC_UTILISATSAISIE', '');
        end;
      finally
        Ferme (lQ);
      end;
    end;

  var
    stSQL: string;
    Q: TQuery;
  begin
    stSQL := 'SELECT E_EXERCICE,MAX(E_DATECOMPTABLE) MAXDATE FROM ECRITURE WHERE E_QUALIFPIECE="N" AND (E_EXERCICE="' + VH^.Precedent.Code + '" OR ' +
      ' E_EXERCICE="' + VH^.EnCours.Code + '" OR E_EXERCICE="' + VH^.Suivant.Code + '") GROUP BY E_EXERCICE';
    Q := OpenSQL(stSQL, True);
    if not Q.Eof then
    begin
      while not Q.Eof do
      begin
        if Q.FindField('E_EXERCICE').AsString = VH^.Precedent.Code then
          T.Detail[0].PutValue('DTC_ENCOURSSAISIE', Q.FindField('MAXDATE').AsDateTime)
        else if Q.FindField('E_EXERCICE').AsString = VH^.Encours.Code then
          T.Detail[1].PutValue('DTC_ENCOURSSAISIE', Q.FindField('MAXDATE').AsDateTime)
        else if Q.FindField('E_EXERCICE').AsString = VH^.Suivant.Code then
          T.Detail[2].PutValue('DTC_ENCOURSSAISIE', Q.FindField('MAXDATE').AsDateTime);
        Q.Next;
      end;
    end;
    Ferme(Q);
    // FQ 18150
    (*
    stSQL := 'SELECT E_EXERCICE,MAX(E_DATEMODIF) MAXDATEMODIF, MAX(E_DATECOMPTABLE) MAXDATE, MAX(E_JOURNAL) MAXJAL, MAX(E_UTILISATEUR) MAXUTILISATEUR FROM ECRITURE WHERE E_QUALIFPIECE="N" AND (E_EXERCICE="'
      + VH^.Precedent.Code + '" OR ' +
      ' E_EXERCICE="' + VH^.EnCours.Code + '" OR E_EXERCICE="' + VH^.Suivant.Code + '") GROUP BY E_EXERCICE';
    *)
    stSQL := 'select e_datecomptable,e_datemodif,e_journal,e_utilisateur from ecriture '
          + 'WHERE E_QUALIFPIECE="N" AND E_EXERCICE="&#@"'
          + ' AND  e_datemodif=(select max(e_datemodif) from ecriture '
          + ' where E_QUALIFPIECE="N" AND E_EXERCICE="&#@")';
    { Pour l'exercice précédent }
    _TBInfoDerniereSaisie (stSQL, VH^.Precedent.Code,0,T);
    _TBInfoDerniereSaisie (stSQL, VH^.EnCours.Code,1,T);
    _TBInfoDerniereSaisie (stSQL, VH^.Suivant.Code,2,T);
  end;

  //--- Ajout CAT : 17/03/2004
  function TBNombreEntree (DateDebut : TDateTime; DateFin : TDateTime) : Integer;
  var ChSql    : String;
      RSql     : Tquery;
  begin
   ChSql:='SELECT COUNT (*) AS NOMBRE FROM IMMOLOG WHERE IL_DATEOP>="'+USDATETIME(DateDebut)+'" AND IL_DATEOP<="'+USDATETIME (DateFin)+'" AND IL_TYPEOP="ACQ"';
   RSql:=OpenSql (ChSql,True);

   if (not RSql.Eof) then
    Result:=RSql.FindField ('NOMBRE').AsInteger
   else
    Result:=0;

   Ferme (Rsql);
  end;

  //--- Ajout CAT : 17/03/2004
  function TBNombreSortie (DateDebut : TDateTime; DateFin : TDateTime) : Integer;
  var ChSql    : String;
      RSql     : Tquery;
  begin
   ChSql:='SELECT COUNT (*) AS NOMBRE FROM IMMOLOG WHERE IL_DATEOP>="'+USDATETIME(DateDebut)+'" AND IL_DATEOP<="'+USDATETIME (DateFin)+'" AND IL_TYPEOP="CES"';
   RSql:=OpenSql (ChSql,True);

   if (not RSql.Eof) then
    Result:=RSql.FindField ('NOMBRE').AsInteger
   else
    Result:=0;

   Ferme (Rsql);
  end;

  //--- Ajout CAT : 17/03/2004
  function TBNombreLigne (DateFin : TDateTime) : Integer;
  var ChSql    : String;
      RSql     : Tquery;
  begin
//  FQ 18152
//   ChSql:='SELECT COUNT (*) AS NOMBRE FROM IMMOLOG WHERE IL_DATEOP>="'+USDATETIME(DateDebut)+'" AND IL_DATEOP<="'+USDATETIME (DateFin)+'"';
    ChSql:='SELECT COUNT (*) NOMBRE FROM IMMO WHERE I_DATEPIECEA<="'+USDATETIME(DateFin)+'" AND I_ETAT<>"FER"';
    RSql:=OpenSql (ChSql,True);

    if (not RSql.Eof) then
      Result:=RSql.FindField ('NOMBRE').AsInteger
    else
     Result:=0;

    Ferme (Rsql);
  end;

  //--- Ajout CAT : 17/03/2004
  function TBNombreImmo () : Integer;
  var ChSql    : String;
      RSql     : Tquery;
  begin
   ChSql:='SELECT COUNT (*) AS NOMBRE FROM IMMO';
   RSql:=OpenSql (ChSql,True);

   if (not RSql.Eof) then
    Result:=RSql.FindField ('NOMBRE').AsInteger
   else
    Result:=0;

   Ferme (Rsql);
  end;

  //--- Ajout CAT : 17/03/2004
  procedure TBInfoImmo (T:Tob);
  var ChSql : String;
      RSql  : Tquery;
  begin
   ChSql:='SELECT EX_Exercice,EX_DATEDEBUT,EX_DATEFIN FROM EXERCICE WHERE EX_EXERCICE="'+VH^.Precedent.Code+'" OR EX_EXERCICE="'+VH^.EnCours.Code+'" OR EX_EXERCICE="'+VH^.Suivant.Code+'"';
   RSql:=OpenSql (ChSql,True);

   While not RSql.Eof do
    begin
     if (RSql.FindField ('Ex_Exercice').AsString=VH^.Precedent.Code) then
      begin
       T.Detail [0].PutValue ('DTC_NbEntreeImmo',TBNombreEntree(RSql.FindField ('EX_DATEDEBUT').AsDateTime,RSql.FindField ('Ex_DATEFIN').AsDateTime));
       T.Detail [0].PutValue ('DTC_NbSortieImmo',TBNombreSortie(RSql.FindField ('Ex_DATEDEBUT').AsDateTime,RSql.FindField ('Ex_DATEFIN').AsDateTime));
       T.Detail [0].PutValue ('DTC_NbLigneImmo',TBNombreLigne (RSql.FindField ('Ex_DATEFIN').AsDateTime));
       //T.Detail [0].PutValue ('DTC_NbImmo',TBNombreImmo ())
      end
     else
      if (RSql.FindField ('Ex_Exercice').AsString=VH^.EnCours.Code) then
       begin
        T.Detail [1].PutValue ('DTC_NbEntreeImmo',TBNombreEntree(RSql.FindField ('Ex_DATEDEBUT').AsDateTime,RSql.FindField ('Ex_DATEFIN').AsDateTime));
        T.Detail [1].PutValue ('DTC_NbSortieImmo',TBNombreSortie(RSql.FindField ('Ex_DATEDEBUT').AsDateTime,RSql.FindField ('Ex_DATEFIN').AsDateTime));
        T.Detail [1].PutValue ('DTC_NbLigneImmo',TBNombreLigne (RSql.FindField ('Ex_DATEFIN').AsDateTime));
        //T.Detail [1].PutValue ('DTC_NbImmo',TBNombreImmo ())
       end
      else
       if (RSql.FindField ('Ex_Exercice').AsString=VH^.Suivant.Code) then
        begin
         T.Detail [2].PutValue ('DTC_NbEntreeImmo',TBNombreEntree(RSql.FindField ('Ex_DATEDEBUT').AsDateTime,RSql.FindField ('Ex_DATEFIN').AsDateTime));
         T.Detail [2].PutValue ('DTC_NbSortieImmo',TBNombreSortie(RSql.FindField ('Ex_DATEDEBUT').AsDateTime,RSql.FindField ('Ex_DATEFIN').AsDateTime));
         T.Detail [2].PutValue ('DTC_NbLigneImmo',TBNombreLigne (RSql.FindField ('Ex_DATEFIN').AsDateTime));
         //T.Detail [2].PutValue ('DTC_NbImmo',TBNombreImmo ())
        end;
     RSql.Next
    end;

   Ferme (RSql);
  end;

var
  T, TPrec, TEnc, TSuiv: TOB;
  Annee,Mois,Jour : Word;
  stMois : string;
begin
  try
    BeginTrans;

    InitMoveProgressForm(nil, 'Mise à jour des cumuls', 'Travail en cours ... Veuillez patienter ...', 6, False, True);
    MoveCurProgressForm('Initialisation des données...');
    { Effacement des enregistrements de la table }
    ExecuteSQL('DELETE from DPTABCOMPTA where DTC_NODOSSIER="' + V_PGI.NoDossier + '"');
    { Création des TOB }
    T := TOB.Create('', nil, -1);
    try
      TPrec := TOB.Create('DPTABCOMPTA', T, -1);
      TEnc := TOB.Create('DPTABCOMPTA', T, -1);
      TSuiv := TOB.Create('DPTABCOMPTA', T, -1);
      { Mise à jour des informations exercice }
      MoveCurProgressForm('Mise à jour des données générales...');
      TBMajInfoExo(T);
      if not bIsDossierCabinet then
      begin
        MoveCurProgressForm('Mise à jour des rubriques...');
        TBChiffreAffaire(T);
        TBMargeTotale(T);
        TBValeurAjoutee(T);
        TBExcedentBrut(T);
        TBResultatExploitation(T);
        TBResultatCourant(T);
        TBResultatExercice(T);
        TBSubventionAgricole(T);
        MoveCurProgressForm('Mise à jour des informations de saisie...');
        TBNombreEcriture(T);
        TBInfoSaisie(T);
        TBInfoImmo (T);
      end;
      { Suppression des exercices non définis }
      if TPrec.GetValue('DTC_DATEDEB')=0 then TPrec.Free;
      if TEnc.GetValue('DTC_DATEDEB')=0 then TEnc.Free;
      if TSuiv.GetValue('DTC_DATEDEB')=0 then TSuiv.Free;
      { Enregistrement dans la base }
      MoveCurProgressForm('Enregistrement des données...');
      T.InsertDB(Nil);
      { Mise à jour des informations DP - FQ 16123 }
      DecodeDate(VH^.Encours.Fin, Annee, Mois, Jour);
      stMois:=Format ('%2.2d',[Mois]);
      ExecuteSQL (
        'update JURIDIQUE set JUR_DATEFINEX="'+USDateTime(VH^.Encours.Fin)+'", '+
        'JUR_DATEDEBUTEX="'+USDateTime(VH^.Encours.Deb) + '", ' +
        'JUR_DUREEEX='+IntToStr(VH^.EnCours.NombrePeriode) + ', '+
        'JUR_DUREEEXPREC='+IntToStr(VH^.Precedent.NombrePeriode) + ' ' +
       'where JUR_GUIDPERDOS="'+GetGuidPer(V_PGI.NoDossier)+'"');
      ExecuteSQL ('UPDATE ANNUAIRE SET ANN_MOISCLOTURE="'+stMois+'" '+
                 'WHERE ANN_GUIDPER="'+GetGuidPer(V_PGI.NoDossier)+'"');
      { Destruction }
    finally
      T.Free;
    end;
    FiniMoveProgressForm;

    CommitTrans;
  except
    RollBack;
  end;
end;

//-------------------------------------------------
//--- Fonction : IntialiserTableCompta
//--- Objet    : Initialise la table DPTABCOMPTA
//-------------------------------------------------
procedure InitialiserTableCompta;
begin
  {$IFNDEF TT}
  MajTableauBordCompta (IsDossierCabinet(V_PGI.NoDossier));
  {$ENDIF}
end;

end.

