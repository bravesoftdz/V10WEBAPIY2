{***********UNITE*************************************************
Auteur  ...... : Xavier PERSOUYRE
Créé le ...... : 23/01/2006
Modifié le ... :   /  /
Description .. : Unité regroupant toutes les procédures et les méthodes
Suite ........ : permettant de changer la façon dont sont gérées les clefs
Suite ........ : des tables de la GED et des tables applicatives rattachées.
Mots clefs ... :                                               MenuOlg  uGedFiles
*****************************************************************}
unit uTablesGed;

interface

uses uTob,
  MajHalleyUtil,
  db,
{$IFNDEF DBXPRESS}dbtables,
{$ELSE}uDbxDataSet,
{$ENDIF}
  Sysutils;

function TraitementTablesGed(): boolean;
// XP 06.03.2006 Procédure exportée suite demande C. AYEL
procedure TraitementTableEcriture();
procedure CreateTGGUID();

procedure Affiche(msg: String); //js1 030506 rajout log
function IsModeTGGUID():boolean; //js1 040506 rajout test mode tgguid avec .ini comme pour vision sav
implementation

uses Hctrls,
  MajTable,
  Hent1,
  utilpgi //js1 utilisation de ismssql
  ,wcommuns //js1 wprogressform
  ,CbpMCD
  ;

{***********A.G.L.***********************************************
Auteur  ...... : Xavier PERSOUYRE
Créé le ...... : 31/01/2006
Modifié le ... :   /  /
Description .. : Permet de vérifier que le traitement de la GED a déjà été fait
Mots clefs...:
*****************************************************************}
function TraitementDejaFait(): boolean;
var
  OldValue: boolean;
begin
  OldValue := V_PGI.EnableDeShare;
  V_PGI.EnableDeShare := False;
  try
    // Est-ce que le champ existe dans le dictionnaire
    if ChampToNum('YFI_FILEGUID') >= 0 then
      Result :=
        ExisteSQL('SELECT YFI_FILEGUID FROM YFILES WHERE YFI_FILEGUID = "MCD{123456789}"')
    else
      result := False;
  finally
    V_PGI.EnableDeShare := OldValue;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Xavier PERSOUYRE
Créé le ...... : 31/01/2006
Modifié le ... :   /  /
Description .. : Permet de vérifier que le traitement de la GED pour cette
Suite ........ : table a été fait ou pas avec lecture de la table YFILES avec
Suite ........ : un GUID = TABLE
Mots clefs...:
*****************************************************************}
function TraitementTableDejaFait(const NouveauChamp: string): boolean;
var
  OldValue: boolean;
begin
  OldValue := V_PGI.EnableDeShare;
  V_PGI.EnableDeShare := False;
  try
    // Est-ce que le champ existe dans le dictionnaire
    if ChampToNum('YFI_FILEGUID') >= 0 then
      Result :=
        ExisteSQL('SELECT YFI_FILEGUID FROM YFILES WHERE YFI_FILEGUID = "MCD{' +
        UpperCase(NouveauChamp) + '}"')
    else
      result := False;
  finally
    V_PGI.EnableDeShare := OldValue;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Xavier PERSOUYRE
Créé le ...... : 31/01/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TraitementTableFait(const NouveauChamp: string);
var
  eMax: integer;
  OldValue: boolean;
begin
  OldValue := V_PGI.EnableDeShare;
  V_PGI.EnableDeShare := False;
  try
    eMax := 10;
    with TOB.Create('YFILES', nil, -1) do
    begin
      PutValue('YFI_FILEGUID', 'MCD{' + UpperCase(NouveauChamp) + '}');
      PutValue('YFI_FILEID', AGLGetIntGUID());
      // XP 07.03.2006 Pour règler un problème d'enregistrement en double toujours avec les Int
      while true do
      begin
        try
          if InsertDb(nil) then
            Break
          else
            raise Exception.Create('Encore une erreur d''insert...');
        except
          on E: exception do
          begin
            if eMax = 0 then
              raise
            else
              PutValue('YFI_FILEID', AGLGetIntGUID());
            Dec(eMax);
          end;
        end;
      end;
      Free;
    end;
  finally
    V_PGI.EnableDeShare := OldValue;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Xavier PERSOUYRE
Créé le ...... : 31/01/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TraitementFait();
var
  OldValue: boolean;
begin
  OldValue := V_PGI.EnableDeShare;
  V_PGI.EnableDeShare := False;
  try
    with TOB.Create('YFILES', nil, -1) do
    begin
      PutValue('YFI_FILEGUID', 'MCD{123456789}');
      PutValue('YFI_FILEID', AGLGetIntGUID());
      // les index ne sont pas encore à jour !
      try
        InsertDb(nil);
      except
        on E: exception do
          raise;
      end;
      Free;
    end;
  finally
    V_PGI.EnableDeShare := OldValue;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Xavier PERSOUYRE
Créé le ...... : 04/04/2006
Modifié le ... :   /  /
Description .. : Permet de supprimer les éléments des tables qui se trouvent
Suite ........ : dans la base DOSSIER, mais qui sont marquées dans la
Suite ........ : table DESHARE comme étant redirigée.
Mots clefs ... :
*****************************************************************}
procedure FlingueTable(const MaTable: string);
var
  Old: Boolean;
begin
  Old := V_PGI.EnableDeShare;
  V_PGI.EnableDeShare := true;
  try
    if TableSurAutreBase(MaTable) and TableExiste(MaTable) then
    begin
      V_PGI.EnableDeShare := False;
      try
        if MaTable = 'YFILES' then
        begin
          // XP 05.04.2006
          if ChampPhysiqueExiste('YFILES', 'YFI_FILEGUID') then
            ExecuteSql('DELETE FROM YFILES WHERE YFI_FILEGUID not LIKE "MCD%"');
        end
        else
          ExecuteSql('DELETE FROM ' + MaTable);
      finally
        V_PGI.EnableDeShare := True;
      end;
    end;
  finally
    V_PGI.EnableDeShare := Old;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Xavier PERSOUYRE
Créé le ...... : 19/01/2006
Modifié le ... :   /  /
Description .. : Cette procédure va créer le champ PREFIXE+SUFFIXE
Suite ........ : dans la table TABLE et en fonction de UpdateChamp mettre
Suite ........ : à jour avec PgiGuid().
Mots clefs ... :
*****************************************************************}
function TraitementTableGed(const Table, NouveauChamp, AncienChamp, TableMaitre,
  NouveauChampMaitre, AncienChampMaitre: string; const UpdateChamp: boolean =
  False; const ActiveDeShare: boolean = False): boolean;
var
  s: string;
  Prefixe: string;
  SaveDEShare: boolean;
  Q:TQuery;
	Mcd : IMCDServiceCOM;
begin
  result := False;

  wMoveProgressForm(TraduireMemoire('Table Traitée : ') + Table);

  // XP 31.03.2006
  SaveDeShare := V_PGI.EnableDESHare;
  V_PGI.enableDEShare := ActiveDeShare;
  try
    // Si la table n'existe pas rien à faire...
    if TableExiste(Table) then
    begin
    // pgr et js1 : chargement table tempo tg_guid avec liens avant traitement
      if not UpdateChamp and IsModeTGGUID then
      begin
         //js1 27042006 cas des tables avec des liens recursifs ==> on aliasse
            if isMssql or isDb2 or (V_PGI.Driver in [dbORACLE9,dbORACLE10]) then
            begin
              ExecuteSQL('INSERT INTO TG_GUID SELECT "' + TableMaitre + '" AS TG_TABLE1, "' + AncienChampMaitre + '" AS TG_CHAMP1,'
                        + ' "' + Table + '" AS TG_TABLE2, "' + AncienChamp + '" AS TG_CHAMP2, T1.'
                        + AncienChampMaitre + ' AS TG_VALEUR1, T2.' + AncienChamp + ' AS TG_VALEUR2,'
                        + ' NULL AS TG_GCHAMP1, NULL AS TG_GCHAMP2, NULL AS TG_GVALEUR1, NULL AS TG_GVALEUR2'
                        +  ' FROM '+ TableMaitre + ' T1 FULL OUTER JOIN ' + Table
                        + ' T2 ON (T1.' + AncienChampMaitre + ' = T2.' + AncienChamp + ')');
          // js1 020506 on compte les liens avant trt ==> on logge
              Q := OpenSQL('SELECT COUNT(*) FROM '+ TableMaitre + ' T1 FULL OUTER JOIN ' + Table
                        + ' T2 ON (T1.' + AncienChampMaitre + ' = T2.' + AncienChamp + ')', True);
              If not Q.Eof then
              Affiche('Table : ' + Table + ' / TableMaitre : ' + TableMaitre + ' / Nombre de liens avant traitement : '+Q.Fields[0].AsString);
              Ferme(Q);
            end
            else
            begin
              ExecuteSQL('INSERT INTO TG_GUID SELECT "' + TableMaitre + '" AS TG_TABLE1, "' + AncienChampMaitre + '" AS TG_CHAMP1,'
                        + ' "' + Table + '" AS TG_TABLE2, "' + AncienChamp + '" AS TG_CHAMP2, T1.'
                        + AncienChampMaitre + ' AS TG_VALEUR1, T2.' + AncienChamp + ' AS TG_VALEUR2,'
                        + ' NULL AS TG_GCHAMP1, NULL AS TG_GCHAMP2, NULL AS TG_GVALEUR1, NULL AS TG_GVALEUR2'
                        +  ' FROM '+ TableMaitre + ' T1 , ' + Table
                        + ' T2 WHERE (T1.' + AncienChampMaitre + ' = T2.' + AncienChamp + '(+)) UNION'
                        + ' SELECT "' + TableMaitre + '" AS TG_TABLE1, "' + AncienChampMaitre + '" AS TG_CHAMP1,'
                        + ' "' + Table + '" AS TG_TABLE2, "' + AncienChamp + '" AS TG_CHAMP2, T1.'
                        + AncienChampMaitre + ' AS TG_VALEUR1, T2.' + AncienChamp + ' AS TG_VALEUR2,'
                        + ' NULL AS TG_GCHAMP1, NULL AS TG_GCHAMP2, NULL AS TG_GVALEUR1, NULL AS TG_GVALEUR2'
                        +  ' FROM '+ TableMaitre + ' T1 , ' + Table
                        + ' T2 WHERE (T1.' + AncienChampMaitre + '(+) = T2.' + AncienChamp + ')');

              Q := OpenSQL('SELECT COUNT(*) FROM '+ TableMaitre + ' T1 , ' + Table
                        + ' T2 WHERE (T1.' + AncienChampMaitre + ' = T2.' + AncienChamp + '(+))'
                        + ' UNION '
                        + 'SELECT COUNT(*) FROM '+ TableMaitre + ' T1 , ' + Table
                        + ' T2 WHERE (T1.' + AncienChampMaitre + '(+) = T2.' + AncienChamp + ')', True);
              If not Q.Eof then
              Affiche('Table : ' + Table + ' / TableMaitre : ' + TableMaitre + ' / Nombre de liens avant traitement : '+Q.Fields[0].AsString);
              Ferme(Q);
            end;
       end;
      // XP 13-02-2006 Dans le cas ou le champ n'existe pas encore en version 714
      if (AncienChamp = '') or ((AncienChamp <> '') and
        ChampPhysiqueExiste(Table,
        AncienChamp)) then
        // XP 31-01-2006 Est-ce que le traitement a déjà été fait ?
        if not TraitementTableDejaFait(NouveauChamp) then
        begin
          FlingueTable(Table);

          //try
            // Le préfixe
          s := NouveauChamp;
          Prefixe := ReadTokenPipe(s, '_');

          if UpdateChamp then
            s := 'CHAR(36)'
          else
            s := 'VarChar(36)';

          AddChamp(Table, NouveauChamp, s, Prefixe, False, '', 'Id', '', False);

          if Prefixe = 'YFI' then
          begin
            SetLength(V_PGI.DEChamps[PrefixeToNum(Prefixe)], 0);
            ChargeDeChamps(PrefixeToNum(Prefixe), Prefixe, TRUE);
            // PCS 10/02/2006   sionon les TOB sont pas à jour !
          end;

          // XP 21-02-2006 Pour des problèmes d'orphelins présents dans les tables "secondaires",
          // toutes les tables seront initialisés avec un GUID
          // ExecuteSql('UPDATE ' + Table + ' SET ' + NouveauChamp + ' = PgiGuid');
          if UpdateChamp then
          begin
            ExecuteSql('UPDATE ' + Table + ' SET ' + NouveauChamp +
              ' = PgiGuid');
          end
          else
          begin
            // La requête de mise à jour
            if Table <> TableMaitre then
              ExecuteSql('UPDATE ' + Table + ' SET ' + NouveauChamp + '=(SELECT '
                +
                NouveauChampMaitre + ' FROM ' + TableMaitre + ' WHERE ' +
                AncienChampMaitre + ' = ' + AncienChamp + ')')
            else
            begin
              // XP 15.03.2006
              if isMssql then
              begin
                ExecuteSql('UPDATE T2 SET T2.' + NouveauChamp + '=T1.' +
                  NouveauChampMaitre +
                  ' FROM ' + TableMaitre + ' T2, ' + Table + ' T1 ' +
                  ' WHERE T2.' + AncienChamp + '=T1.' + AncienChampMaitre);
              end
              else
              begin
                ExecuteSql('UPDATE ' + TableMaitre + ' T2 SET T2.' + NouveauChamp
                  + '=(SELECT T1.' + NouveauChampMaitre +
                  ' FROM ' + TableMaitre + ' T1 WHERE T2.' + AncienChamp + '=T1.'
                  + AncienChampMaitre + ')');
              end;
            end;
          end;
          // XP 21-02-2006 Pour des problèmes d'orphelins présents dans les tables "secondaires",
          // toutes les tables seront initialisés avec un GUID
          // ExecuteSql('UPDATE ' + Table + ' SET ' + NouveauChamp + ' = PgiGuid WHERE (' + NouveauChamp + ' = "") OR (' + NouveauChamp + ' IS NULL)');

          // Traitement effectué avec succès
          TraitementTableFait(NouveauChamp);

          result := True;
          // except
          //   on E: Exception do
          //     raise;
          // end;
        end;

    if not UpdateChamp and IsModeTGGUID then
    begin
    //js1 020506 on update la table avec les liens apres trt
        ExecuteSQL('UPDATE TG_GUID SET TG_GCHAMP1="' + NouveauChampMaitre + '", TG_GCHAMP2="'
                      + NouveauChamp + '", TG_GVALEUR1 =(SELECT MAX(' + NouveauChampMaitre
                      + ') FROM ' + TableMaitre + ' WHERE ' + AncienChampMaitre + ' = TG_VALEUR1)'
                      + ', TG_GVALEUR2=(SELECT MAX(' + NouveauChamp
                      + ') FROM ' + Table + ' WHERE ' + AncienChamp + ' = TG_VALEUR2)'
                      + ' WHERE TG_TABLE1 ="' + TableMaitre + '" AND TG_CHAMP1="' + AncienChampMaitre
                      + '" AND TG_TABLE2="' + Table + '" AND TG_CHAMP2="' + AncienChamp +'"');

    //js1 on compte les liens apres trt et on logge
        if isMssql or isDb2 or (V_PGI.Driver in [dbORACLE9,dbORACLE10]) then
        begin
          Q := OpenSQL('SELECT COUNT(*) FROM '+ TableMaitre + ' T1 FULL OUTER JOIN ' + Table
                      + ' T2 ON (T1.' + NouveauChampMaitre + ' = T2.' + NouveauChamp + ')', True);
          If not Q.Eof then
           Affiche('Table : ' + Table + ' / TableMaitre : ' + TableMaitre + ' / Nombre de liens apres traitement : '+Q.Fields[0].AsString);
          Ferme(Q);
        end
        else
        begin
          Q := OpenSQL('SELECT COUNT(*) FROM '+ TableMaitre + ' T1 , ' + Table
                    + ' T2 WHERE (T1.' + NouveauChampMaitre + ' = T2.' + NouveauChamp + '(+))'
                    + ' UNION '
                    + 'SELECT COUNT(*) FROM '+ TableMaitre + ' T1 , ' + Table
                    + ' T2 WHERE (T1.' + NouveauChampMaitre + '(+) = T2.' + NouveauChamp + ')', True);
          If not Q.Eof then
          Affiche('Table : ' + Table + ' / TableMaitre : ' + TableMaitre + ' / Nombre de liens apres traitement : '+Q.Fields[0].AsString);
          Ferme(Q);
        end;
    end;
        end;
  finally
    V_PGI.EnableDESHare := SaveDeShare;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Xavier PERSOUYRE
Créé le ...... : 09/03/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TraitementTableF2042();
var
  SavEnableDeShare: boolean;
  Q: TQuery;
  Q1: TQuery;
  T: TOB;
begin
  if TableExiste('F2042_VALRUB') and not TraitementTableDejaFait('F2042_VALRUB')
    then
  begin
    SavEnableDeShare := V_PGI.enableDEShare;
    V_PGI.enableDEShare := True;
    T := TOB.Create('', nil, -1);
    try
      Q := OpenSql('SELECT * FROM F2042_VALRUB WHERE F42_NUMRUB = "00500"',
        True);
      if not Q.Eof then
      begin
        Q.First;
        while not Q.EOf do
        begin
          with TOB.Create('F2042_VALRUB', T, -1) do
          begin
            if SelectDb('', Q) then
            begin
              PutValue('F42_NUMRUB', '00501');

              q1 :=
                OpenSql('SELECT ANN_GUIDPER FROM ANNUAIRE WHERE ANN_CODEPER = '
                +
                GetString('F42_VALRUB'), True);
              if not Q1.Eof then
              begin
//{$IFDEF UNICODE}
//                PutValue('F42_VALRUB', Q1.FindField('ANN_GUIDPER').AsValue);
//{$ELSE}
                PutValue('F42_VALRUB', Q1.FindField('ANN_GUIDPER').AsString);
//{$ENDIF}
              end
              else
                PutValue('F42_VALRUB', '');
              Ferme(Q1);
            end
            else
            begin
              Free;
              raise
                Exception.Create('Erreur en lecture de F2042_VALRUB sur SelectDb.');
            end;
          end;
          Q.Next;
        end;
        Ferme(Q);
        Q := nil;

{        // Ecriture des nouveaux enregistrements avec 00501
        if T.Detail.COunt > 0 then
        begin
          try
            // XP 20.03.2006 Suite demande de D.W. T.InsertDb(nil);
            //T.InsertOrUpdateDb();
            T.InsertDb(nil); //js1 010606 retour au comportement initial demande DW
          except
            on E: Exception do
              raise Exception.Create('Erreur en écriture de F2042_VALRUB. ' +
                E.Message);
          end;
        end;}
      end;
      //dw 060606 meme traitement pour les 625
      Q := OpenSql('SELECT * FROM F2042_VALRUB WHERE F42_NUMRUB = "00625"',
        True);
      if not Q.Eof then
      begin
        Q.First;
        while not Q.EOf do
        begin
          with TOB.Create('F2042_VALRUB', T, -1) do
          begin
            if SelectDb('', Q) then
            begin
              PutValue('F42_NUMRUB', '00502');

              q1 :=
                OpenSql('SELECT ANN_GUIDPER FROM ANNUAIRE WHERE ANN_CODEPER = '
                +
                GetString('F42_VALRUB'), True);
              if not Q1.Eof then
              begin
//{$IFDEF UNICODE}
//                PutValue('F42_VALRUB', Q1.FindField('ANN_GUIDPER').AsValue);
//{$ELSE}
                PutValue('F42_VALRUB', Q1.FindField('ANN_GUIDPER').AsString);
//{$ENDIF}
              end
              else
                PutValue('F42_VALRUB', '');
              Ferme(Q1);
            end
            else
            begin
              Free;
              raise
                Exception.Create('Erreur en lecture de F2042_VALRUB sur SelectDb.');
            end;
          end;
          Q.Next;
        end;
        Ferme(Q);
        Q := nil;

        // Ecriture des nouveaux enregistrements avec 00501
        if T.Detail.COunt > 0 then
        begin
          try
            // XP 20.03.2006 Suite demande de D.W. T.InsertDb(nil);
            //T.InsertOrUpdateDb();
            T.InsertDb(nil); //js1 010606 retour au comportement initial demande DW
          except
            on E: Exception do
              raise Exception.Create('Erreur en écriture de F2042_VALRUB. ' +
                E.Message);
          end;
        end;
      end;
    finally
      if assigned(Q) then
        Ferme(Q);
      V_PGI.enableDEShare := SavEnableDeShare;
      FreeAndNil(T);
    end;
    TraitementTableFait('F2042_VALRUB');
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Xavier PERSOUYRE
Créé le ...... : 26/01/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TraitementTableEcriture();
var
  S: string;
  SavEnableDeShare: boolean;
begin
  //try
  // XP 20.03.2006
  if not TraitementTableDejaFait('EC_DOCGUID') and TableExiste('ECRCOMPL') then
  begin
    // mis ici car ce champ peux déjà exister depyuis 730 et obligatoire pour la requête qui suit.
    // Phase 2 : Création du champ EC_DOCGIUD
    AddChamp('ECRCOMPL', 'EC_DOCGUID', 'VarChar(36)', 'EC', False, '', 'Doc Id',
      'L', False);

{$IFDEF AUCASOU}
    // phase 1 : Création des ECRCOMPL dans le cas ou E_DOCID <> 0 et que n'existe pas dans ECRCOMPL
    S :=
      'INSERT INTO ECRCOMPL SELECT E_JOURNAL EC_JOURNAL, E_EXERCICE EC_EXERCICE, E_DATECOMPTABLE EC_DATECOMPTABLE, E_NUMEROPIECE EC_NUMEROPIECE, ';
    s := s +
      'E_NUMLIGNE EC_NUMLIGNE, E_NUMECHE EC_NUMECHE, E_QUALIFPIECE EC_QUALIFPIECE ';
    s := s + ', "' + UsDateTime(IDate1900) + '" EC_CUTOFFDEB, "' +
      UsDateTime(IDate1900) + '" EC_CUTOFFFIN, "' + UsDateTime(IDate1900) +
      '" EC_CUTOFFDATECALC ';
    s := s +
      ', "" EC_CLEECR, "" EC_DOCGUID FROM ECRITURE LEFT JOIN ECRCOMPL ON ';
    s := s +
      'E_JOURNAL=EC_JOURNAL AND E_EXERCICE=EC_EXERCICE AND E_DATECOMPTABLE=EC_DATECOMPTABLE AND E_NUMEROPIECE=EC_NUMEROPIECE AND ';
    s := s + 'E_NUMLIGNE=EC_NUMLIGNE AND E_NUMECHE=EC_NUMECHE ';
    s := s +
      'WHERE EC_JOURNAL IS NULL AND (E_DOCID <> 0 AND E_DOCID <> -1 AND (E_DOCID IS NOT NULL))';
    //s := s + 'WHERE EC_JOURNAL IS NULL AND (E_DOCID <> 0 OR E_DOCID <> -1)';
{$ENDIF AUCASOU}

    // phase 1 : Création des ECRCOMPL dans le cas ou E_DOCID <> 0 et que n'existe pas dans ECRCOMPL
    S :=
      'INSERT INTO ECRCOMPL SELECT E_EXERCICE EC_EXERCICE, E_JOURNAL EC_JOURNAL, E_DATECOMPTABLE EC_DATECOMPTABLE, E_NUMEROPIECE EC_NUMEROPIECE, ';
    s := s +
      'E_NUMLIGNE EC_NUMLIGNE, E_QUALIFPIECE EC_QUALIFPIECE, E_NUMECHE EC_NUMECHE ';
    s := s + ', "' + UsDateTime(IDate1900) + '" EC_CUTOFFDEB, "' +
      UsDateTime(IDate1900) + '" EC_CUTOFFFIN, "' + UsDateTime(IDate1900) +
      '" EC_CUTOFFDATECALC ';
    s := s +
      ', "" EC_CLEECR, "" EC_DOCGUID FROM ECRITURE LEFT JOIN ECRCOMPL ON ';
    s := s +
      ' E_EXERCICE=EC_EXERCICE AND E_JOURNAL=EC_JOURNAL AND E_DATECOMPTABLE=EC_DATECOMPTABLE AND E_NUMEROPIECE=EC_NUMEROPIECE AND ';
    s := s +
      'E_NUMLIGNE=EC_NUMLIGNE AND E_QUALIFPIECE=EC_QUALIFPIECE AND E_NUMECHE=EC_NUMECHE ';
    s := s +
      'WHERE EC_JOURNAL IS NULL AND (E_DOCID <> 0 AND E_DOCID <> -1 AND (E_DOCID IS NOT NULL))';
    //s := s + 'WHERE EC_JOURNAL IS NULL AND (E_DOCID <> 0 OR E_DOCID <> -1)';

    ExecuteSql(S);

    // Phase 2 : Update
    SavEnableDeShare := V_PGI.enableDEShare;
    V_PGI.enableDEShare := True;
    try
      s :=
        'UPDATE ECRCOMPL SET EC_DOCGUID=(SELECT YDO_DOCGUID FROM YDOCUMENTS WHERE YDO_DOCID=(SELECT E_DOCID FROM ECRITURE WHERE';
      s := s +
        ' E_JOURNAL=EC_JOURNAL AND E_EXERCICE=EC_EXERCICE AND E_DATECOMPTABLE=EC_DATECOMPTABLE AND E_NUMEROPIECE=EC_NUMEROPIECE AND ';
      s := s +
        'E_NUMLIGNE=EC_NUMLIGNE AND E_NUMECHE=EC_NUMECHE AND E_QUALIFPIECE=EC_QUALIFPIECE)) ';
      ExecuteSql(S);
    finally
      V_PGI.enableDEShare := SavEnableDeShare;
    end;

    TraitementTableFait('EC_DOCGUID');
  end;
  //except
  //  on E: Exception do
  //    raise;
  //end;
end;

procedure ExecuteSql_(const Table, Sql: string; const ActiveDeShare: boolean =
  False);
var
  OldValue: boolean;
begin
  OldValue := V_PGI.EnableDeShare;
  V_PGI.enableDEShare := ActiveDeShare;
  try
    if TableExiste(Table) then
    begin
      ExecuteSql(Sql);
    end;
  finally
    V_PGI.enableDEShare := OldValue;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Xavier PERSOUYRE
Créé le ...... : 23/01/2006
Modifié le ... :   /  /
Description .. : Point d'entré de l'unité
Mots clefs ... :
*****************************************************************}
function TraitementTablesGed(): boolean;
begin

  wInitProgressForm(nil, traduireMemoire('Passage des identifiants en GUID'), '', 107 , False, True);
  // XP 24.02.2006
  if IsMonoOuCommune() then
  begin
    // XP 07.03.2006 Blindage avec contrôle existance de la table
    ExecuteSQL_('NFILEPARTS',
      'DELETE FROM NFILEPARTS WHERE NOT EXISTS (SELECT 1 FROM NFILES WHERE NFS_FILEID=NFI_FILEID)');
    ExecuteSQL_('YFILEPARTS',
      'DELETE FROM YFILEPARTS WHERE NOT EXISTS (SELECT 1 FROM YFILES WHERE YFP_FILEID=YFI_FILEID)');
    ExecuteSQL_('YMSGFILES',
      'DELETE FROM YMSGFILES WHERE NOT EXISTS (SELECT 1 FROM YFILES WHERE YMG_FILEID=YFI_FILEID)');
    ExecuteSQL_('YMSGFILES',
      'DELETE FROM YMSGFILES WHERE NOT EXISTS (SELECT 1 FROM YMESSAGES WHERE YMG_MSGID=YMS_MSGID)');
    ExecuteSQL_('YDOCFILES',
      'DELETE FROM YDOCFILES WHERE NOT EXISTS (SELECT 1 FROM YFILES WHERE YDF_FILEID=YFI_FILEID)');
    ExecuteSQL_('YDOCFILES',
      'DELETE FROM YDOCFILES WHERE NOT EXISTS (SELECT 1 FROM YDOCUMENTS WHERE YDF_DOCID=YDO_DOCID)');
    ExecuteSQL_('DPDOCUMENT',
      'DELETE FROM DPDOCUMENT WHERE NOT EXISTS (SELECT 1 FROM YDOCUMENTS WHERE DPD_DOCID=YDO_DOCID)');
    ExecuteSQL_('YMODELES',
      'DELETE FROM YMODELES WHERE NOT EXISTS (SELECT 1 FROM YDOCUMENTS WHERE YMO_DOCID=YDO_DOCID)');
    ExecuteSQL_('ANNUBIS',
      'DELETE FROM ANNUBIS WHERE NOT EXISTS (SELECT 1 FROM ANNUAIRE WHERE ANB_CODEPER=ANN_CODEPER)');
    ExecuteSQL_('ANNULIEN',
      'DELETE FROM ANNULIEN WHERE NOT EXISTS (SELECT 1 FROM ANNUAIRE WHERE ANL_CODEPER=ANN_CODEPER)');
    ExecuteSQL_('ANNULIEN',
      'DELETE FROM ANNULIEN WHERE NOT EXISTS (SELECT 1 FROM ANNUAIRE WHERE ANL_CODEPERDOS=ANN_CODEPER)');
    ExecuteSQL_('DPCONTINTV',
      'DELETE FROM DPCONTINTV WHERE NOT EXISTS (SELECT 1 FROM ANNUAIRE WHERE DCI_NODP=ANN_CODEPER)');
    ExecuteSQL_('DPCONTROLE',
      'DELETE FROM DPCONTROLE WHERE NOT EXISTS (SELECT 1 FROM ANNUAIRE WHERE DCL_NODP=ANN_CODEPER)');
    ExecuteSQL_('DPCONVENTION',
      'DELETE FROM DPCONVENTION WHERE NOT EXISTS (SELECT 1 FROM ANNUAIRE WHERE DCV_NODP=ANN_CODEPER)');
    ExecuteSQL_('DPECO',
      'DELETE FROM DPECO WHERE NOT EXISTS (SELECT 1 FROM ANNUAIRE WHERE DEC_NODP=ANN_CODEPER)');
    ExecuteSQL_('DPFISCAL',
      'DELETE FROM DPFISCAL WHERE NOT EXISTS (SELECT 1 FROM ANNUAIRE WHERE DFI_NODP=ANN_CODEPER)');
    ExecuteSQL_('DPINSTIT',
      'DELETE FROM DPINSTIT WHERE NOT EXISTS (SELECT 1 FROM ANNUAIRE WHERE DIS_NODP=ANN_CODEPER)');
    ExecuteSQL_('DPMVTCAP',
      'DELETE FROM DPMVTCAP WHERE NOT EXISTS (SELECT 1 FROM ANNUAIRE WHERE DPM_NODP=ANN_CODEPER)');
    ExecuteSQL_('DPORGA',
      'DELETE FROM DPORGA WHERE NOT EXISTS (SELECT 1 FROM ANNUAIRE WHERE DOR_NODP=ANN_CODEPER)');
    ExecuteSQL_('DPPATRIM',
      'DELETE FROM DPPATRIM WHERE NOT EXISTS (SELECT 1 FROM ANNUAIRE WHERE DPA_NODP=ANN_CODEPER)');
    ExecuteSQL_('DPPERSO',
      'DELETE FROM DPPERSO WHERE NOT EXISTS (SELECT 1 FROM ANNUAIRE WHERE DPP_NODP=ANN_CODEPER)');
    ExecuteSQL_('DPSOCIAL',
      'DELETE FROM DPSOCIAL WHERE NOT EXISTS (SELECT 1 FROM ANNUAIRE WHERE DSO_NODP=ANN_CODEPER)');
    ExecuteSQL_('DPSOCIALCAISSE',
      'DELETE FROM DPSOCIALCAISSE WHERE NOT EXISTS (SELECT 1 FROM ANNUAIRE WHERE DSC_NODP=ANN_CODEPER)');
    ExecuteSQL_('HISTOANNULIEN',
      'DELETE FROM HISTOANNULIEN WHERE NOT EXISTS (SELECT 1 FROM ANNUAIRE WHERE HNL_CODEPER=ANN_CODEPER)');
    ExecuteSQL_('HISTOANNULIEN',
      'DELETE FROM HISTOANNULIEN WHERE NOT EXISTS (SELECT 1 FROM ANNUAIRE WHERE HNL_CODEPERDOS=ANN_CODEPER)');
    ExecuteSQL_('JUBAUXFONDS',
      'DELETE FROM JUBAUXFONDS WHERE NOT EXISTS (SELECT 1 FROM ANNUAIRE WHERE JBF_CODEPERDOS=ANN_CODEPER)');
    ExecuteSQL_('JUHISTOREM',
      'DELETE FROM JUHISTOREM WHERE NOT EXISTS (SELECT 1 FROM ANNUAIRE WHERE JHM_CODEPERDOS=ANN_CODEPER)');
    ExecuteSQL_('JUHISTOREM',
      'DELETE FROM JUHISTOREM WHERE NOT EXISTS (SELECT 1 FROM ANNUAIRE WHERE JHM_CODEPER=ANN_CODEPER)');
    ExecuteSQL_('JUMVTCOMPTES',
      'DELETE FROM JUMVTCOMPTES WHERE NOT EXISTS (SELECT 1 FROM ANNUAIRE WHERE JMC_CODEPERDOS=ANN_CODEPER)');
    ExecuteSQL_('JUMVTCOMPTES',
      'DELETE FROM JUMVTCOMPTES WHERE NOT EXISTS (SELECT 1 FROM ANNUAIRE WHERE JMC_CODEPER=ANN_CODEPER)');
    ExecuteSQL_('JUMVTREM',
      'DELETE FROM JUMVTREM WHERE NOT EXISTS (SELECT 1 FROM ANNUAIRE WHERE JMR_CODEPERDOS=ANN_CODEPER)');
    ExecuteSQL_('JUMVTREM',
      'DELETE FROM JUMVTREM WHERE NOT EXISTS (SELECT 1 FROM ANNUAIRE WHERE JMR_CODEPER=ANN_CODEPER)');
    ExecuteSQL_('JUMVTTITRES',
      'DELETE FROM JUMVTTITRES WHERE NOT EXISTS (SELECT 1 FROM ANNUAIRE WHERE JMT_CODEPERDOS=ANN_CODEPER)');
    ExecuteSQL_('JUMVTTITRES',
      'DELETE FROM JUMVTTITRES WHERE NOT EXISTS (SELECT 1 FROM ANNUAIRE WHERE JMT_CODEPER=ANN_CODEPER)');
    ExecuteSQL_('JURIDIQUE',
      'DELETE FROM JURIDIQUE WHERE NOT EXISTS (SELECT 1 FROM ANNUAIRE WHERE JUR_CODEPERDOS=ANN_CODEPER)');
    ExecuteSQL_('JUSYSCLE',
      'DELETE FROM JUSYSCLE WHERE NOT EXISTS (SELECT 1 FROM ANNUAIRE WHERE JSC_CODEI=ANN_CODEPER)');
    ExecuteSQL_('LIENSOLE',
      'delete from liensole where (lo_tableblob="ANN" or lo_tableblob="JUR") and not exists (select 1 from annuaire where ann_codeper=lo_rangblob)');
    // XP 24.02.2006
  end;

  //PGR & JS1 Création de la table d'analyse

  if IsModeTGGUID then CreateTGGUID();
  TraitementTableGed('YFILES', 'YFI_FILEGUID', '', '', '', '', True);

  TraitementTableGed('YFILEPARTS', 'YFP_FILEGUID', 'YFP_FILEID', 'YFILES',
    'YFI_FILEGUID', 'YFI_FILEID', False);

  TraitementTableGed('YDOCFILES', 'YDF_FILEGUID', 'YDF_FILEID', 'YFILES',
    'YFI_FILEGUID', 'YFI_FILEID', False);

  TraitementTableGed('YMAILFILES', 'YMF_FILEGUID', 'YMF_FILEID', 'YFILES',
    'YFI_FILEGUID', 'YFI_FILEID', False);

  TraitementTableGed('YMSGFILES', 'YMG_FILEGUID', 'YMG_FILEID', 'YFILES',
    'YFI_FILEGUID', 'YFI_FILEID', False);

  TraitementTableGed('NFILES', 'NFI_FILEGUID', '', '', '', '', True);

  TraitementTableGed('NFILEPARTS', 'NFS_FILEGUID', 'NFS_FILEID', 'NFILES',
    'NFI_FILEGUID', 'NFI_FILEID', False);

  TraitementTableGed('YFILESTD', 'YFS_FILEGUID', 'YFS_FILEID', 'NFILES',
    'NFI_FILEGUID', 'NFI_FILEID', False);

  // XP 08.09.2006 Correction de PGV_FILECV par PGV_FILEIDCV
  TraitementTableGed('PCURRICULUM', 'PGV_FILEGUIDCV', 'PGV_FILEIDCV', 'YFILES',
    'YFI_FILEGUID', 'YFI_FILEID', True);
  TraitementTableGed('PCURRICULUM', 'PGV_FILEGUIDLM', 'PGV_FILEIDLM', 'YFILES',
    'YFI_FILEGUID', 'YFI_FILEID', True);
  TraitementTableGed('PCURRICULUM', 'PGV_FILEGUIDDIVERS', 'PGV_FILEIDDIVERS',
    'YFILES', 'YFI_FILEGUID', 'YFI_FILEID', True);

  TraitementTableGed('PSUIVITCANDIDAT', 'PS2_FILEGUID', 'PS2_FILEID', 'YFILES',
    'YFI_FILEGUID', 'YFI_FILEID', True);

  TraitementTableGed('PSUIVICV', 'PS3_FILEGUID', 'PS3_FILEID', 'YFILES',
    'YFI_FILEGUID', 'YFI_FILEID', True);

  // XP 26.04.2006
  TraitementTableGed('PCVFILE', 'PXF_FILEGUID', 'PXF_FILEID', 'YFILES', 'YFI_FILEGUID', 'YFI_FILEID', True);

  //XP 120606
  TraitementTableGed('PECOLES', 'PY8_NOTEGUID', 'PY8_NOTE', 'YFILES', 'YFI_FILEGUID', 'YFI_FILEID', True);

  // TraitementTableGed('ANNUAIRE', 'ANN_CODEPER', 'ANN_GUIDPER', '', '', '', True);
  if TraitementTableGed('ANNUAIRE', 'ANN_GUIDPER', '', '', '', '', True) then
  begin
    if IsMonoOuCommune then
    begin
      ExecuteSQL('UPDATE LIENSOLE SET'
        + ' LO_IDENTIFIANT=(SELECT ANN_GUIDPER FROM ANNUAIRE WHERE (LO_TABLEBLOB="ANN" OR LO_TABLEBLOB="JUR") AND ANN_CODEPER=LO_RANGBLOB ),'
        + ' LO_EMPLOIBLOB="ACT", '
        + ' LO_RANGBLOB=1'
        +
        ' WHERE (LO_TABLEBLOB="ANN" OR LO_TABLEBLOB="JUR") AND LO_IDENTIFIANT="OLE_ACTIVITE"');

      ExecuteSQL('UPDATE LIENSOLE SET'
        + ' LO_IDENTIFIANT=(SELECT ANN_GUIDPER FROM ANNUAIRE WHERE (LO_TABLEBLOB="ANN" OR LO_TABLEBLOB="JUR") AND ANN_CODEPER=LO_RANGBLOB ),'
        + ' LO_EMPLOIBLOB="BLO",'
        + ' LO_RANGBLOB=2'
        +
        ' WHERE (LO_TABLEBLOB="ANN" OR LO_TABLEBLOB="JUR") AND LO_IDENTIFIANT="OLE_BLOCNOTES"');

      ExecuteSQL('UPDATE LIENSOLE SET'
        + ' LO_IDENTIFIANT=(SELECT ANN_GUIDPER FROM ANNUAIRE WHERE (LO_TABLEBLOB="ANN" OR LO_TABLEBLOB="JUR") AND ANN_CODEPER=LO_RANGBLOB ),'
        + ' LO_EMPLOIBLOB="REG",'
        + ' LO_RANGBLOB=3'
        +
        ' WHERE (LO_TABLEBLOB="ANN" OR LO_TABLEBLOB="JUR") AND LO_IDENTIFIANT="OLE_REGMATTXT"');

      ExecuteSQL('UPDATE LIENSOLE SET'
        + ' LO_IDENTIFIANT=(SELECT ANN_GUIDPER FROM ANNUAIRE WHERE (LO_TABLEBLOB="ANN" OR LO_TABLEBLOB="JUR") AND ANN_CODEPER=LO_RANGBLOB ),'
        + ' LO_EMPLOIBLOB="OBJ",'
        + ' LO_RANGBLOB=4'
        +
        ' WHERE (LO_TABLEBLOB="ANN" OR LO_TABLEBLOB="JUR") AND LO_IDENTIFIANT="OLE_OBJETSOC"');
    end;
  end;
  TraitementTableGed('ANNUAIRE', 'ANN_GUIDCJ', 'ANN_CODECJ', 'ANNUAIRE',
    'ANN_GUIDPER', 'ANN_CODEPER', False);
  TraitementTableGed('ANNUAIRE', 'ANN_PERASS1GUID', 'ANN_PERASS1CODE',
    'ANNUAIRE', 'ANN_GUIDPER', 'ANN_CODEPER', False);
  TraitementTableGed('ANNUAIRE', 'ANN_PERASS2GUID', 'ANN_PERASS2CODE',
    'ANNUAIRE', 'ANN_GUIDPER', 'ANN_CODEPER', False);
  // XP 27.02.2006 Suppression demandée par MD TraitementTableGed('ANNUAIRE', 'ANN_GUIDPERINT', 'ANN_NOPER', 'ANNUAIRE', 'ANN_GUIDPER', 'ANN_CODEPER', False);

  TraitementTableGed('ANNUBIS', 'ANB_GUIDPER', 'ANB_CODEPER', 'ANNUAIRE',
    'ANN_GUIDPER', 'ANN_CODEPER', False);

  TraitementTableGed('ANNULIEN', 'ANL_GUIDPER', 'ANL_CODEPER', 'ANNUAIRE',
    'ANN_GUIDPER', 'ANN_CODEPER', False);
  TraitementTableGed('ANNULIEN', 'ANL_GUIDPERDOS', 'ANL_CODEPERDOS', 'ANNUAIRE',
    'ANN_GUIDPER', 'ANN_CODEPER', False);
  TraitementTableGed('ANNULIEN', 'ANL_PERASS1GUID', 'ANL_PERASS1CODE',
    'ANNUAIRE', 'ANN_GUIDPER', 'ANN_CODEPER', False);
  TraitementTableGed('ANNULIEN', 'ANL_PERASS2GUID', 'ANL_PERASS2CODE',
    'ANNUAIRE', 'ANN_GUIDPER', 'ANN_CODEPER', False);
  TraitementTableGed('ANNULIEN', 'ANL_PERASS3GUID', 'ANL_PERASS3CODE',
    'ANNUAIRE', 'ANN_GUIDPER', 'ANN_CODEPER', False);
  TraitementTableGed('ANNULIEN', 'ANL_COOPTGUID', 'ANL_COOPTCODE', 'ANNUAIRE',
    'ANN_GUIDPER', 'ANN_CODEPER', False);
  TraitementTableGed('ANNULIEN', 'ANL_CRSPGUID', 'ANL_CRSPCODE', 'ANNUAIRE',
    'ANN_GUIDPER', 'ANN_CODEPER', False);
  TraitementTableGed('ANNULIEN', 'ANL_MDRPGUID', 'ANL_MDRPCODE', 'ANNUAIRE',
    'ANN_GUIDPER', 'ANN_CODEPER', False);

  TraitementTableGed('DOSSIER', 'DOS_GUIDPER', 'DOS_CODEPER', 'ANNUAIRE',
    'ANN_GUIDPER', 'ANN_CODEPER', False); // ??

  TraitementTableGed('DPCONTINTV', 'DCI_GUIDPER', 'DCI_NODP', 'ANNUAIRE',
    'ANN_GUIDPER', 'ANN_CODEPER', False);

  TraitementTableGed('DPCONTROLE', 'DCL_GUIDPER', 'DCL_NODP', 'ANNUAIRE',
    'ANN_GUIDPER', 'ANN_CODEPER', False);

  TraitementTableGed('DPCONVENTION', 'DCV_GUIDPER', 'DCV_NODP', 'ANNUAIRE',
    'ANN_GUIDPER', 'ANN_CODEPER', False);

  TraitementTableGed('DPECO', 'DEC_GUIDPER', 'DEC_NODP', 'ANNUAIRE',
    'ANN_GUIDPER', 'ANN_CODEPER', False);

  TraitementTableGed('DPFISCAL', 'DFI_GUIDPER', 'DFI_NODP', 'ANNUAIRE',
    'ANN_GUIDPER', 'ANN_CODEPER', False);
  TraitementTableGed('DPFISCAL', 'DFI_GUIDPEROGA', 'DFI_NOOGADP', 'ANNUAIRE',
    'ANN_GUIDPER', 'ANN_CODEPER', False);
  TraitementTableGed('DPFISCAL', 'DFI_GUIDPERTETEGRD', 'DFI_NOREFTETEGRDP',
    'ANNUAIRE', 'ANN_GUIDPER', 'ANN_CODEPER', False);

  TraitementTableGed('DPINSTIT', 'DIS_GUIDPER', 'DIS_NODP', 'ANNUAIRE',
    'ANN_GUIDPER', 'ANN_CODEPER', False);

  TraitementTableGed('DPMVTCAP', 'DPM_GUIDPER', 'DPM_NODP', 'ANNUAIRE',
    'ANN_GUIDPER', 'ANN_CODEPER', False);

  TraitementTableGed('DPORGA', 'DOR_GUIDPER', 'DOR_NODP', 'ANNUAIRE',
    'ANN_GUIDPER', 'ANN_CODEPER', False);

  TraitementTableGed('DPPATRIM', 'DPA_GUIDPER', 'DPA_NODP', 'ANNUAIRE',
    'ANN_GUIDPER', 'ANN_CODEPER', False);

  TraitementTableGed('DPPERSO', 'DPP_GUIDPER', 'DPP_NODP', 'ANNUAIRE',
    'ANN_GUIDPER', 'ANN_CODEPER', False);

  TraitementTableGed('DPSOCIAL', 'DSO_GUIDPER', 'DSO_NODP', 'ANNUAIRE',
    'ANN_GUIDPER', 'ANN_CODEPER', False);

  TraitementTableGed('DPSOCIALCAISSE', 'DSC_GUIDPER', 'DSC_NODP', 'ANNUAIRE',
    'ANN_GUIDPER', 'ANN_CODEPER', False);

  TraitementTableGed('HISTOANNULIEN', 'HNL_GUIDPER', 'HNL_CODEPER', 'ANNUAIRE',
    'ANN_GUIDPER', 'ANN_CODEPER', False);
  TraitementTableGed('HISTOANNULIEN', 'HNL_GUIDPERDOS', 'HNL_CODEPERDOS',
    'ANNUAIRE', 'ANN_GUIDPER', 'ANN_CODEPER', False);
  TraitementTableGed('HISTOANNULIEN', 'HNL_PERASS1GUID', 'HNL_PERASS1CODE',
    'ANNUAIRE', 'ANN_GUIDPER', 'ANN_CODEPER', False);
  TraitementTableGed('HISTOANNULIEN', 'HNL_COOPTGUID', 'HNL_COOPTCODE',
    'ANNUAIRE', 'ANN_GUIDPER', 'ANN_CODEPER', False);

  TraitementTableGed('JURIDIQUE', 'JUR_GUIDPERDOS', 'JUR_CODEPERDOS',
    'ANNUAIRE',
    'ANN_GUIDPER', 'ANN_CODEPER', False);

  TraitementTableGed('JUBAUXFONDS', 'JBF_GUIDPERDOS', 'JBF_CODEPERDOS',
    'JURIDIQUE', 'JUR_GUIDPERDOS', 'JUR_CODEPERDOS', False);
  TraitementTableGed('JUHISTOREM', 'JHM_GUIDPERDOS', 'JHM_CODEPERDOS',
    'JURIDIQUE', 'JUR_GUIDPERDOS', 'JUR_CODEPERDOS', False);
  TraitementTableGed('JUMVTCOMPTES', 'JMC_GUIDPERDOS', 'JMC_CODEPERDOS',
    'JURIDIQUE', 'JUR_GUIDPERDOS', 'JUR_CODEPERDOS', False);
  TraitementTableGed('JUMVTREM', 'JMR_GUIDPERDOS', 'JMR_CODEPERDOS',
    'JURIDIQUE',
    'JUR_GUIDPERDOS', 'JUR_CODEPERDOS', False);
  TraitementTableGed('JUMVTTITRES', 'JMT_GUIDPERDOS', 'JMT_CODEPERDOS',
    'JURIDIQUE', 'JUR_GUIDPERDOS', 'JUR_CODEPERDOS', False);

  TraitementTableGed('JUBAUXFONDS', 'JBF_GUIDBENEF', 'JBF_CODEBENEF',
    'ANNUAIRE',
    'ANN_GUIDPER', 'ANN_CODEPER', False);
  TraitementTableGed('JUDOSINFO', 'JDI_GUIDPER', 'JDI_CODEPER', 'ANNUAIRE',
    'ANN_GUIDPER', 'ANN_CODEPER', False);
  TraitementTableGed('JUEVENEMENT', 'JEV_GUIDPER', 'JEV_CODEPER', 'ANNUAIRE',
    'ANN_GUIDPER', 'ANN_CODEPER', False);
  TraitementTableGed('JUGRPSOC', 'JGR_GUIDPER', 'JGR_CODEPER', 'ANNUAIRE',
    'ANN_GUIDPER', 'ANN_CODEPER', False);
  TraitementTableGed('JUHISTOREM', 'JHM_GUIDPER', 'JHM_CODEPER', 'ANNUAIRE',
    'ANN_GUIDPER', 'ANN_CODEPER', False);
  TraitementTableGed('JUMVTCOMPTES', 'JMC_GUIDPER', 'JMC_CODEPER', 'ANNUAIRE',
    'ANN_GUIDPER', 'ANN_CODEPER', False);
  TraitementTableGed('JUMVTREM', 'JMR_GUIDPER', 'JMR_CODEPER', 'ANNUAIRE',
    'ANN_GUIDPER', 'ANN_CODEPER', False);
  TraitementTableGed('JUMVTTITRES', 'JMT_GUIDPER', 'JMT_CODEPER', 'ANNUAIRE',
    'ANN_GUIDPER', 'ANN_CODEPER', False);

  TraitementTableGed('JUSYSCLE', 'JSC_GUIDI', 'JSC_CODEI', 'ANNUAIRE',
    'ANN_GUIDPER', 'ANN_CODEPER', False);

  // M. DESGOUTTE MSGID
  TraitementTableGed('YMESSAGES', 'YMS_MSGGUID', 'YMS_MSGID', '', '', '', True);
  TraitementTableGed('YMSGFILES', 'YMG_MSGGUID', 'YMG_MSGID', 'YMESSAGES',
    'YMS_MSGGUID', 'YMS_MSGID', False);

  // M. DESGOUTTE DOCID
  TraitementTableGed('YDOCUMENTS', 'YDO_DOCGUID', 'YDO_DOCID', '', '', '',
    True);
  TraitementTableGed('YDOCFILES', 'YDF_DOCGUID', 'YDF_DOCID', 'YDOCUMENTS',
    'YDO_DOCGUID', 'YDO_DOCID', False);
  TraitementTableGed('DPDOCUMENT', 'DPD_DOCGUID', 'DPD_DOCID', 'YDOCUMENTS',
    'YDO_DOCGUID', 'YDO_DOCID', False);
  TraitementTableGed('YMODELES', 'YMO_DOCGUID', 'YMO_DOCID', 'YDOCUMENTS',
    'YDO_DOCGUID', 'YDO_DOCID', False);
  // XP 27.02.2006 Ajouté par MD
  TraitementTableGed('YREPERTPERSO', 'YRP_GUIDREP', 'YRP_CODEREP', '', '', '',
    True);
  TraitementTableGed('JUEVENEMENT', 'JEV_GUIDEVT', 'JEV_NOEVT', '', '', '',
    True);
  TraitementTableGed('JUDOSOPACT', 'JOA_GUIDACT', 'JOA_NOACT', '', '', '',
    True);
  TraitementTableGed('JUDOSOPACT', 'JOA_GUIDEVT', 'JOA_NOEVT', 'JUEVENEMENT',
    'JEV_GUIDEVT', 'JEV_NOEVT', False);

  // P. BASSET
  TraitementTableGed('L02_ENTDECLA', 'L02_GUIDPER', 'L02_CODEPER', 'ANNUAIRE',
    'ANN_GUIDPER', 'ANN_CODEPER', False);
  TraitementTableGed('L02_ENTDECLA', 'L02_GUIDPERPERE', 'L02_LIACODEPERPERE',
    'ANNUAIRE', 'ANN_GUIDPER', 'ANN_CODEPER', False);
  TraitementTableGed('L02_ENTDECLA', 'L02_GUIDPERFILS', 'L02_LIACODEPERFILS',
    'ANNUAIRE', 'ANN_GUIDPER', 'ANN_CODEPER', False);

  // XP 08.03.2006 Demande de P. BASSET
  TraitementTableGed('F2072_ASSOCIE', 'A72_GUIDPER', 'A72_CODEPER', 'ANNUAIRE',
    'ANN_GUIDPER', 'ANN_CODEPER', False, True);
  TraitementTableGed('F2072_GERANT', 'G72_GUIDPER', 'G72_CODEPER', 'ANNUAIRE',
    'ANN_GUIDPER', 'ANN_CODEPER', False, True);
  TraitementTableGed('NBOOKDFPF', 'NBF_FILEGUID', 'NBF_NUMFILE', 'NFILES',
    'NFI_FILEGUID', 'NFI_FILEID', False);

  // C. DUMAS
  TraitementTableGed('RTDOCUMENT', 'RTD_DOCGUID', 'RTD_DOCID', 'YDOCUMENTS',
    'YDO_DOCGUID', 'YDO_DOCID', False);

  // XP 15.03.2006 Demande de P. Basset
  TraitementTableGed('F2036_ASSOCIE', 'A36_GUIDPER', 'A36_CODEPER', 'ANNUAIRE',
    'ANN_GUIDPER', 'ANN_CODEPER', False, True);
  TraitementTableGed('F2036_ASSOCIEP', 'P36_GUIDPER', 'P36_CODEPER', 'ANNUAIRE',
    'ANN_GUIDPER', 'ANN_CODEPER', False, True);
  TraitementTableGed('F9065_MEMBRE', 'B95_GUIDPER', 'B95_CODEPER', 'ANNUAIRE',
    'ANN_GUIDPER', 'ANN_CODEPER', False, True);

  // XP 15.03.2006 Demande de C. AYEL
  TraitementTableGed('ICCGENERAUX', 'ICG_GUIDPER', 'ICG_CODEPER', 'ANNUAIRE',
    'ANN_GUIDPER', 'ANN_CODEPER', False, True);

  // XP 15.03.2006 Demande de P. DUMET
  TraitementTableGed('EMPLOIINTERIM', 'PEI_AGENCEINTGU', 'PEI_AGENCEINTERIM',
    'ANNUAIRE', 'ANN_GUIDPER', 'ANN_CODEPER', False, True);
  TraitementTableGed('EMPLOIINTERIM', 'PEI_CENTREFORMGU', 'PEI_CENTREFORM',
    'ANNUAIRE', 'ANN_GUIDPER', 'ANN_CODEPER', False, True);
  TraitementTableGed('FORMATIONS', 'PFO_ORGCOLLECTGU', 'PFO_ORGCOLLECT',
    'ANNUAIRE', 'ANN_GUIDPER', 'ANN_CODEPER', False, True);
  TraitementTableGed('FORMATIONS', 'PFO_CENTREFORMGU', 'PFO_CENTREFORM',
    'ANNUAIRE', 'ANN_GUIDPER', 'ANN_CODEPER', False, True);
  TraitementTableGed('INVESTFORMATION', 'PIF_ORGCOLLECTGU', 'PIF_ORGCOLLECT',
    'ANNUAIRE', 'ANN_GUIDPER', 'ANN_CODEPER', False, True);
  TraitementTableGed('SESSIONSTAGE', 'PSS_ORGCOLLECTSGU', 'PSS_ORGCOLLECTS',
    'ANNUAIRE', 'ANN_GUIDPER', 'ANN_CODEPER', False, True);
  TraitementTableGed('SESSIONSTAGE', 'PSS_ORGCOLLECTPGU', 'PSS_ORGCOLLECTP',
    'ANNUAIRE', 'ANN_GUIDPER', 'ANN_CODEPER', False, True);
  TraitementTableGed('SESSIONSTAGE', 'PSS_CENTREFORMGU', 'PSS_CENTREFORM',
    'ANNUAIRE', 'ANN_GUIDPER', 'ANN_CODEPER', False, True);
  TraitementTableGed('STAGE', 'PST_ORGCOLLECTSGU', 'PST_ORGCOLLECTS',
    'ANNUAIRE',
    'ANN_GUIDPER', 'ANN_CODEPER', False, True);
  TraitementTableGed('STAGE', 'PST_ORGCOLLECTPGU', 'PST_ORGCOLLECTP',
    'ANNUAIRE',
    'ANN_GUIDPER', 'ANN_CODEPER', False, True);
  TraitementTableGed('STAGE', 'PST_CENTREFORMGU', 'PST_CENTREFORM', 'ANNUAIRE',
    'ANN_GUIDPER', 'ANN_CODEPER', False, True);
  TraitementTableGed('ENVOIFORMATION', 'PVF_ORGCOLLECTGU', 'PVF_ORGCOLLECT',
    'ANNUAIRE', 'ANN_GUIDPER', 'ANN_CODEPER', False, True);
  TraitementTableGed('VISITEMEDTRAV', 'PVM_MEDTRAVGU', 'PVM_MEDTRAV',
    'ANNUAIRE', 'ANN_GUIDPER', 'ANN_CODEPER', False, True);
  TraitementTableGed('RETENUESALAIRE', 'PRE_BENEFRSGU', 'PRE_BENEFICIAIRERS',
    'ANNUAIRE', 'ANN_GUIDPER', 'ANN_CODEPER', False, True);
  TraitementTableGed('ETABCOMPL', 'ETB_CODEDDTEFPGU', 'ETB_CODEDDTEFP',
    'ANNUAIRE', 'ANN_GUIDPER', 'ANN_CODEPER', False, True);
  TraitementTableGed('ETABCOMPL', 'ETB_MEDTRAVGU', 'ETB_MEDTRAV', 'ANNUAIRE',
    'ANN_GUIDPER', 'ANN_CODEPER', False, True);

  // XP 22.03.2006 pour D.W.
  TraitementTableGed('FPDOSSIER', 'FPO_GUIDPER', 'FPO_CODEPER', 'ANNUAIRE',
    'ANN_GUIDPER', 'ANN_CODEPER', False, True);
  TraitementTableGed('FPBIENS', 'FPB_GUIDPER', 'FPB_N01201', 'ANNUAIRE',
    'ANN_GUIDPER', 'ANN_CODEPER', False, True);
  TraitementTableGed('FPPERSONNE', 'FPR_GUIDPER', 'FPR_CODEPER', 'ANNUAIRE',
    'ANN_GUIDPER', 'ANN_CODEPER', False, True);
  TraitementTableGed('F2044_DEC', 'RF4_GUIDPER', 'RF4_N00500', 'ANNUAIRE',
    'ANN_GUIDPER', 'ANN_CODEPER', False, True);
  TraitementTableGed('F2725_DEC', 'FSF_GUIDPER', 'FSF_N04001', 'ANNUAIRE',
    'ANN_GUIDPER', 'ANN_CODEPER', False, True);

  // JS1 02.11.2006 pour D.W.
  TraitementTableGed('FTP_LIEU', 	'FLI_GUIDPER', 'FLI_LICODE',  'ANNUAIRE',
    'ANN_GUIDPER', 'ANN_CODEPER', False, True);
  // XP 28.03.2006 pour P.B.
  TraitementTableGed('CA3_PARAMETRE', 'CP3_GUIDPER', 'CP3_CODEPER', 'ANNUAIRE',
    'ANN_GUIDPER', 'ANN_CODEPER', False);

  // XP 29.03.2006
  TraitementTableGed('TRFDOSSIER', 'TRD_GUIDPER', 'TRD_CODEPER', 'ANNUAIRE',
    'ANN_GUIDPER', 'ANN_CODEPER', False);

  // XP 03.04.2006 par MD
  TraitementTableGed('FEMPRUNT', 'EMP_GUIDORGPRETEUR', 'EMP_ORGPRETEUR',
    'ANNUAIRE', 'ANN_GUIDPER', 'ANN_CODEPER', False, True);

{$IFDEF AUCASOU}
  // XP 11.04.2006 par CV
  TraitementTableGed('F2072_ASSOCIE', 'A72_GUIDPERSCI', 'A72_NUMSCI', 'ANNUAIRE','ANN_GUIDPER', 'ANN_CODEPER', False, True);
{$ENDIF}

   // XP 26-01-2006 Cas particulier de la table ECRITURE(E) et ECRCOMPL(EC)
  TraitementTableEcriture();

  // XP 09.03.2006 Pour D.W.
  TraitementTableF2042();

  // Traitement effectué avec succès
  TraitementFait();

  wFiniProgressForm;

  Result := True;

 // except
    // Result := False;
  // end;
end;

//pgr et js1 260406 creation table tempo d'analyse du traitement guid
procedure CreateTGGUID();
var
  PNewStructure, PNewField: TOB;
  i: Integer;
begin
//js1 020506 on supprime tabel tempo eventuellement existante
  if TableExiste('TG_GUID') then DBDeleteTable(DBSOC,V_PGI.Driver,'TG_GUID',true);

  PNewStructure := TOB.Create('TG_GUID', nil, -1);
  PNewStructure.AddChampSupValeur('DT_NOMTABLE', 'TG_GUID');
  PNewStructure.AddChampSupValeur('DT_CLE1', 'TG_TABLE1');
  PNewStructure.AddChampSupValeur('DT_UNIQUE1', '-');
  PNewStructure.AddChampSupValeur('DT_CLE2', 'TG_CHAMP1');
  PNewStructure.AddChampSupValeur('DT_UNIQUE2', '-');
  PNewStructure.AddChampSupValeur('DT_CLE3', 'TG_TABLE2');
  PNewStructure.AddChampSupValeur('DT_UNIQUE3', '-');
  PNewStructure.AddChampSupValeur('DT_CLE4', 'TG_TABLE2');
  PNewStructure.AddChampSupValeur('DT_UNIQUE4', '-');
  PNewStructure.AddChampSupValeur('DT_PREFIXE', 'TG');

  for i := 2 to MaxIndexes do
    PNewStructure.AddChampSupValeur('DT_CLE' + IntToStr(i), '');

  PNewField := TOB.Create('FIELD', pNewStructure, -1);
  PNewField.AddChampSupValeur('DH_NOMCHAMP', 'TG_TABLE1');
  PNewField.AddChampSupValeur('DH_TYPECHAMP', 'VARCHAR(50)');
  PNewField.AddChampSupValeur('DH_PREFIX', 'TG');
  PNewField.AddChampSupValeur('DH_NULLABLE', '-');

  PNewField := TOB.Create('FIELD', pNewStructure, -1);
  PNewField.AddChampSupValeur('DH_NOMCHAMP', 'TG_CHAMP1');
  PNewField.AddChampSupValeur('DH_TYPECHAMP', 'VARCHAR(50)');
  PNewField.AddChampSupValeur('DH_PREFIX', 'TG');
  PNewField.AddChampSupValeur('DH_NULLABLE', '-');

  PNewField := TOB.Create('FIELD', pNewStructure, -1);
  PNewField.AddChampSupValeur('DH_NOMCHAMP', 'TG_TABLE2');
  PNewField.AddChampSupValeur('DH_TYPECHAMP', 'VARCHAR(50)');
  PNewField.AddChampSupValeur('DH_PREFIX', 'TG');
  PNewField.AddChampSupValeur('DH_NULLABLE', '-');

  PNewField := TOB.Create('FIELD', pNewStructure, -1);
  PNewField.AddChampSupValeur('DH_NOMCHAMP', 'TG_CHAMP2');
  PNewField.AddChampSupValeur('DH_TYPECHAMP', 'VARCHAR(50)');
  PNewField.AddChampSupValeur('DH_PREFIX', 'TG');
  PNewField.AddChampSupValeur('DH_NULLABLE', '-');

  PNewField := TOB.Create('FIELD', pNewStructure, -1);
  PNewField.AddChampSupValeur('DH_NOMCHAMP', 'TG_VALEUR1');
  PNewField.AddChampSupValeur('DH_TYPECHAMP', 'INT');
  PNewField.AddChampSupValeur('DH_PREFIX', 'TG');
  PNewField.AddChampSupValeur('DH_NULLABLE', 'X');

  PNewField := TOB.Create('FIELD', pNewStructure, -1);
  PNewField.AddChampSupValeur('DH_NOMCHAMP', 'TG_VALEUR2');
  PNewField.AddChampSupValeur('DH_TYPECHAMP', 'INT');
  PNewField.AddChampSupValeur('DH_PREFIX', 'TG');
  PNewField.AddChampSupValeur('DH_NULLABLE', 'X');

  PNewField := TOB.Create('FIELD', pNewStructure, -1);
  PNewField.AddChampSupValeur('DH_NOMCHAMP', 'TG_GCHAMP1');
  PNewField.AddChampSupValeur('DH_TYPECHAMP', 'VARCHAR(50)');
  PNewField.AddChampSupValeur('DH_PREFIX', 'TG');
  PNewField.AddChampSupValeur('DH_NULLABLE', 'X');

  PNewField := TOB.Create('FIELD', pNewStructure, -1);
  PNewField.AddChampSupValeur('DH_NOMCHAMP', 'TG_GCHAMP2');
  PNewField.AddChampSupValeur('DH_TYPECHAMP', 'VARCHAR(50)');
  PNewField.AddChampSupValeur('DH_PREFIX', 'TG');
  PNewField.AddChampSupValeur('DH_NULLABLE', 'X');

  PNewField := TOB.Create('FIELD', pNewStructure, -1);
  PNewField.AddChampSupValeur('DH_NOMCHAMP', 'TG_GVALEUR1');
  PNewField.AddChampSupValeur('DH_TYPECHAMP', 'VARCHAR(50)');
  PNewField.AddChampSupValeur('DH_PREFIX', 'TG');
  PNewField.AddChampSupValeur('DH_NULLABLE', 'X');

  PNewField := TOB.Create('FIELD', pNewStructure, -1);
  PNewField.AddChampSupValeur('DH_NOMCHAMP', 'TG_GVALEUR2');
  PNewField.AddChampSupValeur('DH_TYPECHAMP', 'VARCHAR(50)');
  PNewField.AddChampSupValeur('DH_PREFIX', 'TG');
  PNewField.AddChampSupValeur('DH_NULLABLE', 'X');

  DbCreateTable(DBSOC, PNewStructure, V_PGI.Driver, False);
end;

//js1 020506 on logge
procedure Affiche(msg: String);
begin
  LogAgl('***  ' + msg + '  ***');
end;

//js1 040506 on teste la presence d'un .ini pour modetgguid
function IsModeTGGUID():boolean;
begin
  // pour forcer mode tg_guid
  result := FileExists(ExtractFilePath('PGIMAJVER.exe') + 'GUID_34A1A000-6432-406A-B6B2-ED4CAB625CD0.ini')
end;
end.

