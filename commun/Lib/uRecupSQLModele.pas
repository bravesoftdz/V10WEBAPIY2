unit uRecupSQLModele;

interface
uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms,
{$IFDEF EAGLCLIENT}
  UTob,
{$ELSE}
{$IFNDEF DBXPRESS}dbtables{$ELSE}uDbxDataSet{$ENDIF}, DB,
{$ENDIF}
  UtilPGI, // EstMultiSoc
  ed_tools, HEnt1, Hctrls;

{$IFNDEF ERADIO}
function RecupSQLModele(Tipe, Nat, Modele, CodeL: string; Auxi, Gene: String17; WhereSQL: string = ''): string;
{$ENDIF !ERADIO}
function RecupSQLEtat(Tipe, Nat, Modele, CodeL: string; Auxi, Gene: String17; WhereSQL: string; var OrderSQL: string): string;
function RecupSQLComplet(Tipe, Nat, Modele: string; var SWhere, SGroup, SOrder: string): string;

procedure RecupSQLMultiSoc(vType, vNature, vModele: string; var vSQL: string);

implementation

{$IFNDEF ERADIO}
{$IFDEF EAGLCLIENT}

function RecupSQLModele(Tipe, Nat, Modele, CodeL: string; Auxi, Gene: String17; WhereSQL: string = ''): string;
var Q: TQuery;
  p, po: integer;
  Sql, SqlW, StOrder, SAVLangue: string;
  Tsql: TStrings;
begin
  if WhereSQL = '' then SqlW := ' WHERE T_AUXILIAIRE="' + Auxi + '" '
  else SqlW := WhereSQL;
  if (Nat = 'REL') or (Nat = 'RLC') then Sql := 'SELECT * FROM TIERS LEFT OUTER JOIN ECRITURE ON E_AUXILIAIRE=T_AUXILIAIRE ' + SqlW;
  if Nat = 'RLV' then
  begin
    SqlW := ' WHERE E_AUXILIAIRE="' + Auxi + '" AND E_GENERAL="' + Gene + '" '
      + 'AND E_ETATLETTRAGE="TL" AND E_LETTRAGE="' + CodeL + '" AND E_FLAGECR="ROR"'
      + ' ORDER BY E_AUXILIAIRE, E_GENERAL, E_ETATLETTRAGE, E_LETTRAGE, '
      + 'E_DATECOMPTABLE, E_DATEECHEANCE, E_NUMECHE ';
    Sql := 'SELECT ECRITURE.*, TIERS.*, RIB.*, (E_DEBIT-E_CREDIT) E_SOLDE FROM ECRITURE '
      + 'LEFT JOIN TIERS ON E_AUXILIAIRE=T_AUXILIAIRE '
      + 'LEFT JOIN RIB ON (E_AUXILIAIRE=R_AUXILIAIRE AND R_PRINCIPAL="X")'
      + SqlW;
  end;
// BPY le 18/11/2003 : fiche 13055
//Q:=OpenSQL('SELECT * FROM MODEDATA WHERE MD_CLE="'+Tipe+Nat+Modele+'SQL'+V_PGI.LanguePrinc+'"',True) ;
  SAVLangue := V_PGI.LanguePrinc;
  if (not ExisteSQL('SELECT MD_CLE FROM MODEDATA WHERE MD_CLE="' + Tipe + Nat + Modele + 'SQL' + V_PGI.LanguePrinc + '"')) then
  begin
    Q := OpenSQL('SELECT MD_CLE FROM MODEDATA WHERE MD_CLE LIKE "' + Tipe + Nat + Modele + 'SQL%"', True,-1,'',true);
    if (not Q.EOF) then V_PGI.LanguePrinc := Copy(Q.FindField('MD_CLE').AsString, Length(Tipe + Nat + Modele + 'SQL') + 1, Length(Q.FindField('MD_CLE').AsString));
    Ferme(Q);
  end;
  Q := OpenSQL('SELECT * FROM MODEDATA WHERE MD_CLE="' + Tipe + Nat + Modele + 'SQL' + V_PGI.LanguePrinc + '"', True,-1,'',true);
  V_PGI.LanguePrinc := SAVLangue;
// Fin BPY
//FetchSQLODBC(Q) ;
  if not Q.EOF then
  begin
    Sql := '';
    TSql := TStringList.Create;
    TSql.Text := TmemoField(Q.FindField('MD_DATA')).AsString;
    SQL := SQL + ' ' + SQLToString(TSQL, True);
  //For i:=0 to TSQL.Count-1 do SQL:=SQL+TSQL[i] ;
    TSql.Free;
    p := Pos(' WHERE', Sql);
  { GP 17/02/98 : }
    StOrder := '';
    if Nat = 'RLC' then
    begin
      StOrder := ''; po := Pos('ORDER BY', Sql);
      if (po > 0) and (p > 0) then
      begin
        StOrder := Copy(Sql, po, Length(Sql) - po + 1);
        System.Delete(Sql, po, Length(Sql) - po + 1);
      end;
    end;
  { fin GP 17/02/98 }
    if p <> 0 then Sql := Copy(Sql, 1, p);
    Sql := Sql + SqlW + ' ' + StOrder;
  end;
  Ferme(Q); Result := Sql;
end;

{$ELSE}

function RecupSQLModele(Tipe, Nat, Modele, CodeL: string; Auxi, Gene: String17; WhereSQL: string = ''): string;
var Q: TQuery;
  s: TMemoryStream;
  p, po: integer;
  Sql, SqlW, StOrder, SAVLangue: string;
  Tsql: TStrings;
begin
  if WhereSQL = '' then SqlW := ' WHERE T_AUXILIAIRE="' + Auxi + '" '
  else SqlW := WhereSQL;
  if (Nat = 'REL') or (Nat = 'RLC') then Sql := 'SELECT * FROM TIERS LEFT OUTER JOIN ECRITURE ON E_AUXILIAIRE=T_AUXILIAIRE ' + SqlW;
  if Nat = 'RLV' then
  begin
    SqlW := ' WHERE E_AUXILIAIRE="' + Auxi + '" AND E_GENERAL="' + Gene + '" '
      + 'AND E_ETATLETTRAGE="TL" AND E_LETTRAGE="' + CodeL + '" AND E_FLAGECR="ROR"'
      + ' ORDER BY E_AUXILIAIRE, E_GENERAL, E_ETATLETTRAGE, E_LETTRAGE, '
      + 'E_DATECOMPTABLE, E_DATEECHEANCE, E_NUMECHE ';
    Sql := 'SELECT ECRITURE.*, TIERS.*, RIB.*, (E_DEBIT-E_CREDIT) E_SOLDE FROM ECRITURE '
      + 'LEFT JOIN TIERS ON E_AUXILIAIRE=T_AUXILIAIRE '
      + 'LEFT JOIN RIB ON (E_AUXILIAIRE=R_AUXILIAIRE AND R_PRINCIPAL="X")'
      + SqlW;
  end;
// BPY le 18/11/2003 : fiche 13055
//Q:=OpenSQL('SELECT * FROM MODEDATA WHERE MD_CLE="'+Tipe+Nat+Modele+'SQL'+V_PGI.LanguePrinc+'"',True) ;
  SAVLangue := V_PGI.LanguePrinc;
  if (not ExisteSQL('SELECT MD_CLE FROM MODEDATA WHERE MD_CLE="' + Tipe + Nat + Modele + 'SQL' + V_PGI.LanguePrinc + '"')) then
  begin
    Q := OpenSQL('SELECT MD_CLE FROM MODEDATA WHERE MD_CLE LIKE "' + Tipe + Nat + Modele + 'SQL%"', True,-1,'',true);
    if (not Q.EOF) then V_PGI.LanguePrinc := Copy(Q.FindField('MD_CLE').AsString, Length(Tipe + Nat + Modele + 'SQL') + 1, Length(Q.FindField('MD_CLE').AsString));
    Ferme(Q);
  end;
  Q := OpenSQL('SELECT * FROM MODEDATA WHERE MD_CLE="' + Tipe + Nat + Modele + 'SQL' + V_PGI.LanguePrinc + '"', True,-1,'',true);
  V_PGI.LanguePrinc := SAVLangue;
// Fin BPY
  FetchSQLODBC(Q);
  if not Q.EOF then
  begin
    s := TMemoryStream.Create; s.clear;
    TBlobField(Q.FindField('MD_DATA')).SaveToStream(s);
    s.position := 0; Sql := '';
    TSql := TStringList.Create;
    TSql.LoadFromStream(s);
    SQL := SQL + ' ' + SQLToString(TSQL, True);
  //For i:=0 to TSQL.Count-1 do SQL:=SQL+TSQL[i] ;
    s.Free; TSql.Free;
    p := Pos(' WHERE', Sql);
  { GP 17/02/98 : }
    StOrder := '';
    if Nat = 'RLC' then
    begin
      StOrder := ''; po := Pos('ORDER BY', Sql);
      if (po > 0) and (p > 0) then
      begin
        StOrder := Copy(Sql, po, Length(Sql) - po + 1);
        System.Delete(Sql, po, Length(Sql) - po + 1);
      end;
    end;
  { fin GP 17/02/98 }
    if p <> 0 then Sql := Copy(Sql, 1, p);
    Sql := Sql + SqlW + ' ' + StOrder;
  end;
  Ferme(Q); Result := Sql;
end;
{$ENDIF EAGLCLIENT}
{$ENDIF !ERADIO}

{$IFDEF EAGLCLIENT}

function RecupSQLEtat(Tipe, Nat, Modele, CodeL: string; Auxi, Gene: String17; WhereSQL: string; var OrderSQL: string): string;
var Q: TQuery;
  p, po: integer;
  Sql, SqlW, StOrder, SAVLangue: string;
  Tsql: TStrings;
  nbWhere, PCumul : integer;
  SqlWhere : string;
begin
  SqlW := WhereSQL;
// Modif Marché Mode pour impression en langues étrangères : OT
  SAVLangue := V_PGI.LanguePrinc;
  if not ExisteSQL('SELECT MD_CLE FROM MODEDATA WHERE MD_CLE="' + Tipe + Nat + Modele + 'SQL' + V_PGI.LanguePrinc + '"') then
  begin
    Q := OpenSQL('SELECT MD_CLE FROM MODEDATA WHERE MD_CLE LIKE "' + Tipe + Nat + Modele + 'SQL%"', True,-1,'',true);
    if not Q.EOF then
      V_PGI.LanguePrinc := Copy(Q.FindField('MD_CLE').AsString, Length(Tipe + Nat + Modele + 'SQL') + 1, Length(Q.FindField('MD_CLE').AsString));
    Ferme(Q);
  end;
  Q := OpenSQL('SELECT * FROM MODEDATA WHERE MD_CLE="' + Tipe + Nat + Modele + 'SQL' + V_PGI.LanguePrinc + '"', True, -1, 'MODEDATA',true);
  V_PGI.LanguePrinc := SAVLangue;
  if not Q.EOF then
  begin
    Sql := '';
    TSql := TStringList.Create;
    TSql.Text := TmemoField(Q.FindField('MD_DATA')).AsString;
    SQL := SQL + ' ' + SQLToString(TSQL, True);
    TSql.Free;
    
    pCumul := Pos(' WHERE',Sql) + 6 ;
    SqlWhere := Copy (Sql, pCumul, Length (Sql));
    p := pCumul;
    nbWhere := 0;
    while pos (' WHERE', SqlWhere) > 0 do
    begin
      pCumul := Pos(' WHERE', SqlWhere) + 6 ;
      SqlWhere := Copy (SqlWhere, pCumul, Length (SqlWhere));
      Inc (nbWhere);
      p := p + pCumul;
    end;
    p := p - (6 + nbWhere);

    StOrder := ''; po := Pos('ORDER BY', Sql);
    if (po > 0) and (p > 0) then
    begin
      StOrder := Copy(Sql, po, Length(Sql) - po + 1);
      System.Delete(Sql, po, Length(Sql) - po + 1);
    end;
    if p <> 0 then Sql := Copy(Sql, 1, p);
    Sql := Sql + SqlW;
    OrderSQL := StOrder;
  end;
  Ferme(Q); Result := Sql;
end;

{$ELSE}

function RecupSQLEtat(Tipe, Nat, Modele, CodeL: string; Auxi, Gene: String17; WhereSQL: string; var OrderSQL: string): string;
var Q: TQuery;
  s: TMemoryStream;
  p, po: integer;
  Sql, SqlW, StOrder, SAVLangue: string;
  Tsql: TStrings;
  nbWhere, PCumul : integer;
  SqlWhere : string;
begin
  SqlW := WhereSQL;
// Modif Marché Mode pour impression en langues étrangères : OT
  SAVLangue := V_PGI.LanguePrinc;
  if not ExisteSQL('SELECT MD_CLE FROM MODEDATA WHERE MD_CLE="' + Tipe + Nat + Modele + 'SQL' + V_PGI.LanguePrinc + '"') then
  begin
    Q := OpenSQL('SELECT MD_CLE FROM MODEDATA WHERE MD_CLE LIKE "' + Tipe + Nat + Modele + 'SQL%"', True,-1,'',true);
{$IFNDEF EAGLSERVER}
    if not Q.EOF then
      V_PGI.LanguePrinc := Copy(Q.FindField('MD_CLE').AsString, Length(Tipe + Nat + Modele + 'SQL') + 1, Length(Q.FindField('MD_CLE').AsString));
{$ENDIF !EAGLSERVER}
    Ferme(Q);
  end;
  Q := OpenSQL('SELECT * FROM MODEDATA WHERE MD_CLE="' + Tipe + Nat + Modele + 'SQL' + V_PGI.LanguePrinc + '"', True,-1,'',true);
{$IFNDEF EAGLSERVER}
  V_PGI.LanguePrinc := SAVLangue;
{$ENDIF !EAGLSERVER}
  FetchSQLODBC(Q);
  if not Q.EOF then
  begin
    s := TMemoryStream.Create; s.clear;
    TBlobField(Q.FindField('MD_DATA')).SaveToStream(s);
    s.position := 0; Sql := '';
    TSql := TStringList.Create;
    TSql.LoadFromStream(s);
    SQL := SQL + ' ' + SQLToString(TSQL, True);
    s.Free; TSql.Free;

    pCumul := Pos(' WHERE',Sql) + 6 ;
    SqlWhere := Copy (Sql, pCumul, Length (Sql));
    p := pCumul;
    nbWhere := 0;
    while pos (' WHERE', SqlWhere) > 0 do
    begin
      pCumul := Pos(' WHERE', SqlWhere) + 6 ;
      SqlWhere := Copy (SqlWhere, pCumul, Length (SqlWhere));
      Inc (nbWhere);
      p := p + pCumul;
    end;
    p := p - (6 + nbWhere);
    
  { GP 17/02/98 : }
    StOrder := '';
    StOrder := ''; po := Pos('ORDER BY', Sql);
    if (po > 0) and (p > 0) then
    begin
      StOrder := Copy(Sql, po, Length(Sql) - po + 1);
      System.Delete(Sql, po, Length(Sql) - po + 1);
    end;
    if p <> 0 then Sql := Copy(Sql, 1, p);
    Sql := Sql + SqlW;
    OrderSQL := StOrder;
  end;
  Ferme(Q); Result := Sql;
end;
{$ENDIF EAGLCLIENT}

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 20/02/2004
Modifié le ... : 20/02/2004
Description .. : Recupertation de la requete SQL d'un etat
Suite ........ : Recuperation complete
Suite ........ : avec :
Suite ........ :   - le select dans la variable de retour
Suite ........ :   - la clause "where" dans "SWhere"
Suite ........ :   - la clause "group by" dans "SGroup"
Suite ........ :   - et la clause "order by" dans "SOrder"
Mots clefs ... : SQL;ETAT;
*****************************************************************}

function RecupSQLComplet(Tipe, Nat, Modele: string; var SWhere, SGroup, SOrder: string): string;
var
  Q: TQuery;
  pw, po, pg: integer;
  SQL, Langue: string;
  Tsql: TStrings;
begin
    // recuperation d'une langue d'etat valide !
  if (not ExisteSQL('SELECT MD_CLE FROM MODEDATA WHERE MD_CLE="' + Tipe + Nat + Modele + 'SQL' + V_PGI.LanguePrinc + '"')) then
  begin
    Q := OpenSQL('SELECT MD_CLE FROM MODEDATA WHERE MD_CLE LIKE "' + Tipe + Nat + Modele + 'SQL%"', true,-1,'',true);
    if (not Q.EOF) then Langue := Copy(Q.FindField('MD_CLE').AsString, Length(Tipe + Nat + Modele + 'SQL') + 1, Length(Q.FindField('MD_CLE').AsString));
    Ferme(Q);
  end
  else Langue := V_PGI.LanguePrinc;

    // Selection du SQL de l'etat
  Q := OpenSQL('SELECT * FROM MODEDATA WHERE MD_CLE="' + Tipe + Nat + Modele + 'SQL' + Langue + '"', true,-1,'',true);
    // si existant
  if (not Q.EOF) then
  begin
        // recup de la requete
    TSql := TStringList.Create;
    TSql.Text := TmemoField(Q.FindField('MD_DATA')).AsString;
    SQL := SQLToString(TSQL, True);
        // free des objets temp
    TSql.Free;

        // uppercase
    SQL := UpperCase(SQL);

        // recup de la clause "order by"
    po := Pos(' ORDER BY ', SQL);
    if (po > 0) then
    begin
      SOrder := Trim(Copy(SQL, po, Length(SQL) - po + 1));
      System.Delete(SQL, po, Length(SQL) - po + 1);
    end
    else SOrder := '';

        // recup de la clause "group by"
    pg := Pos(' GROUP BY ', SQL);
    if (pg > 0) then
    begin
      SGroup := Trim(Copy(SQL, pg, Length(SQL) - pg + 1));
      System.Delete(SQL, pg, Length(SQL) - pg + 1);
    end
    else SGroup := '';

        // recup de la clause "where"
    pw := Pos(' WHERE ', SQL);
    if (pw > 0) then
    begin
      SWhere := Trim(Copy(SQL, pw, Length(SQL) - pw + 1));
      System.Delete(SQL, pw, Length(SQL) - pw + 1);
    end
    else SWhere := '';
  end
  else SQL := '';

  Ferme(Q);
  Result := Trim(SQL);
end;


{***********A.G.L.***********************************************
Auteur  ...... : Compta
Créé le ...... : 23/05/2005
Modifié le ... :   /  /
Description .. : Permet de remplacer le nom des tables GENERAUX,
Suite ........ : SECTION ou TIERS par la vue équivalente en mode
Suite ........ : multisociété.
Suite ........ :
Suite ........ : PLUS UTILISE DEPUIS L'EXISTENCE DU
Suite ........ : V_PGI.ENABLETABLETOVUE
Suite ........ :
Mots clefs ... :
*****************************************************************}

procedure RecupSQLMultiSoc(vType, vNature, vModele: string; var vSQL: string);
var lReqSQL: string;
  lGroup: string;
  lOrder: string;
  lWhere: string;

  procedure _TraiteTable(vNomTable: string; var vStSQL: string);
  begin
    if pos(vNomTable, vStSQL) < 0 then Exit;
      // Remplacement de la table par la vue
    vStSQL := FindEtReplace(vStSQL, vNomTable, '##MS', True);
    vStSQL := FindEtReplace(vStSQL, '##MS', vNomTable + 'MS', True);
  end;

begin

  // Abandon si pas de traitement nécessaire
  if not EstMultiSoc then Exit;

  // Récupération de la requete SQL sauf si commence par select
  if copy(vSQL, 1, 5) = 'SELECT'
    then lReqSQL := vSQL
  else begin
    lReqSQL := RecupSQLComplet(vType, vNature, vModele, lWhere, lGroup, lOrder);
         //RecupSQLEtat( vType, vNature, vModele, '', '', '', vSQL, lOrderSQL ) ;
    if lWhere <> '' then
    begin
      lReqSQL := lReqSQL + ' WHERE ' + lWhere;
      if vSQL <> '' then
        lReqSQL := lReqSQL + ' AND ' + vSQL;
    end
    else
    begin
      if vSQL <> '' then
        lReqSQL := lReqSQL + ' WHERE ' + vSQL;
    end;
    if lGroup <> '' then
      lReqSQL := lReqSQL + ' GROUP BY ' + lGroup;
    if lOrder <> '' then
      lReqSQL := lReqSQL + ' ORDER BY ' + lOrder;
  end;

  // Rempacement des tables par les vues si besoin
  if EstTablePartagee('GENERAUX') then
    _TraiteTable('GENERAUX', lReqSQL);

  if EstTablePartagee('TIERS') then
    _TraiteTable('TIERS', lReqSQL);

  if EstTablePartagee('SECTION') then
    _TraiteTable('SECTION', lReqSQL);

  // Modification de la requête initiale
  vSQL := lReqSQL;

end;


end.
