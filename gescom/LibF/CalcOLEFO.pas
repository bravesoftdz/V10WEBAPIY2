{***********UNITE*************************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 25/11/2002
Modifié le ... : 25/11/2002
Description .. : Fonctions @ pour le générateur d'état et de ticket
Mots clefs ... : FO
*****************************************************************}
unit CalcOLEFO;      

interface

uses
  SysUtils, Classes,
  {$IFNDEF EAGLCLIENT}
  dbtables, db,
  {$ENDIF}
  Hctrls, UTob, ED_Tools, Hent1;

function FOCalcOLEEtat(sf, sp: string): variant;

implementation

uses
  Ent1, EntGC, 
  MFOSOLDECLIENT_TOF,
  TickUtilFO, FOUtil, FODefi;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 23/07/2001
Modifié le ... : 23/07/2001
Description .. : Retourne le nom et le code de chacune des dimensions
Suite ........ : d'un article
Mots clefs ... : FO
*****************************************************************}

function FOGetArtDim(CodeArticle: string; Abrege, Compact, AvecTitre: boolean; Separateur: string): string;
var QQ: TQuery;
  Sql: string;
  TOBArt: TOB;
begin
  Result := '';
  Sql := 'select GA_ARTICLE,GA_GRILLEDIM1,GA_GRILLEDIM2,GA_GRILLEDIM3,GA_GRILLEDIM4,GA_GRILLEDIM5 '
    + 'from ARTICLE where GA_ARTICLE="' + CodeArticle + '"';
  QQ := OpenSQL(Sql, True);
  if not QQ.EOF then
  begin
    TOBArt := TOB.Create('ARTICLE', nil, -1);
    TOBArt.SelectDB('', QQ, False);
    Result := FOLibelleDimension(TOBArt, Abrege, Compact, AvecTitre, Separateur);
    TOBArt.Free;
  end;
  Ferme(QQ);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 23/07/2001
Modifié le ... : 23/07/2001
Description .. : Retourne la valeur d'un paramètre de la caisse
Mots clefs ... : FO
*****************************************************************}

function FOGetPCaisse(NomChamp, Libelle: string): string;
var TypeChamp, Stg: string;
  Valeur: Variant;
begin
  Result := '';
  if (VH_GC.TOBPCaisse <> nil) and (VH_GC.TOBPCaisse.FieldExists(NomChamp)) then
  begin
    TypeChamp := ChampToType(NomChamp);
    Valeur := VH_GC.TOBPCaisse.GetValue(NomChamp);
    if (TypeChamp = 'INTEGER') or (TypeChamp = 'SMALLINT') then
      Result := IntToStr(Valeur)
    else if (TypeChamp = 'DOUBLE') or (TypeChamp = 'RATE') then
      Result := FloatToStr(Valeur)
    else if TypeChamp = 'DATE' then
      Result := DateTimeToStr(Valeur)
    else if (TypeChamp = 'DATA') or (TypeChamp = 'BLOB') then
      Result := ''
    else
    begin
      Stg := VarAsType(Valeur, VarOleStr);
      if Stg = #0 then Stg := '';
      Result := Stg;
      if (Libelle = 'L') or (Libelle = 'A') then
        Result := RechDom(Get_Join(NomChamp), Result, (Libelle = 'A'));
    end;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 18/11/2002
Modifié le ... : 18/11/2002
Description .. : Retourne la valeur d'un paramètre d'un code démarque
Mots clefs ... : FO
*****************************************************************}

function FOGetDemarque(Code, Option: string): string;
var TOBDem, TOBL: TOB;
  TobLocal: boolean;
begin
  Result := FODecodeLibreDemarque(nil, Option);
  if Code = '' then Exit;
  if VH_GC.TOBPCaisse = nil then Exit;
  if VH_GC.TOBPCaisse.GetValue('GPK_DEMSAISIE') = '-' then Exit;
  TobLocal := False;
  TOBDem := FOGetPTobDemarque;
  if TOBDem = nil then
  begin
    TOBDem := TOB.Create('Types remise', nil, -1);
    TOBDem.LoadDetailDB('TYPEREMISE', '"' + Code + '"', '', nil, False);
    TobLocal := True;
  end;
  if TOBDem <> nil then
  begin
    TOBL := TOBDem.FindFirst(['GTR_TYPEREMISE'], [Code], False);
    if TOBL <> nil then Result := FODecodeLibreDemarque(TOBL, Option);
    if TobLocal then TOBDem.Free;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : O. TARCY
Créé le ...... : 23/07/2001
Modifié le ... : 25/03/2002
Description .. : Calcul les lignes suivantes du Ticket Z :
Suite ........ :  * nombre de tickets,
Suite ........ :  * le nombres d'articles vendus,
Suite ........ :  * le cumul des remises.
Suite ........ :
Suite ........ : Appeler depuis le générateur de ticket par
Suite ........ : '@FOGETPARAM(DIVERS;...'
Mots clefs ... : FO
*****************************************************************}

function FONbEnrgt(Champ, Caisse, Numz, GroupBy: string): double;
var QQ: TQuery;
  SQL: string;
  Compteur: integer;
begin
  SQL := '';
  if (Champ = 'GL_QTEFACT') or (Champ = 'GL_TOTREMLIGNEDEV') then
    SQL := 'SELECT SUM(' + Champ + ') AS TOT'
  else
    SQL := 'SELECT ' + Champ;
  SQL := SQL + ' FROM JOURSCAISSE ' +
    'LEFT OUTER JOIN PIECE ON JOURSCAISSE.GJC_CAISSE=PIECE.GP_CAISSE ' +
    'AND JOURSCAISSE.GJC_NUMZCAISSE=PIECE.GP_NUMZCAISSE ' +
    'AND JOURSCAISSE.GJC_DATEOUV=PIECE.GP_DATEPIECE ' +
    'LEFT OUTER JOIN LIGNE ON LIGNE.GL_NATUREPIECEG=PIECE.GP_NATUREPIECEG ' +
    'AND LIGNE.GL_SOUCHE=PIECE.GP_SOUCHE ' +
    'AND LIGNE.GL_NUMERO=PIECE.GP_NUMERO ' +
    'AND LIGNE.GL_INDICEG=PIECE.GP_INDICEG ' +
    'WHERE GJC_NUMZCAISSE=' + NumZ + ' AND GJC_CAISSE= "' + Caisse + '" ' +
    'AND GP_NATUREPIECEG=' + FOGetNatureTicket(False, True) + ' ' +
    'AND GL_TYPEARTICLE IN ("MAR","NOM")';
  if GroupBy = 'TRUE' then SQL := SQL + ' GROUP BY ' + Champ;
  QQ := OpenSQL(SQL, True);
  if (Champ = 'GL_QTEFACT') or (Champ = 'GL_TOTREMLIGNEDEV') then
    Result := QQ.FindField('TOT').AsFloat
  else
  begin
    Compteur := 0;
    while not QQ.EOF do
    begin
      Inc(Compteur);
      QQ.Next;
    end;
    Result := Compteur;
  end;
  Ferme(QQ);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : O. TARCY
Créé le ...... : 23/07/2001
Modifié le ... : 23/07/2001
Description .. : Retourne le nombre d'entrées de la période
Mots clefs ... : FO
*****************************************************************}

function FOGetNbEntrees(Etab, Caisse, DeNumz, ANumZ: string): integer;
var FONbEntrees: TQuery;
  DeDate, ADate: TDateTime;
begin
  Result := 0;
  DeDate := Date;
  ADate := Date;
  // Récupération de la date de début
  FONbEntrees := OpenSQL('SELECT GJC_DATEOUV FROM JOURSCAISSE WHERE GJC_CAISSE="' + Caisse +
    '" AND GJC_NUMZCAISSE=' + DeNumZ, True);
  if not FONbEntrees.EOF then
    DeDate := FONbEntrees.FindField('GJC_DATEOUV').AsDateTime;
  Ferme(FONbEntrees);
  // Récupération de la date de fin
  FONbEntrees := OpenSQL('SELECT GJC_DATEOUV,GJC_DATEFERME FROM JOURSCAISSE WHERE GJC_CAISSE="' + Caisse +
    '" AND GJC_NUMZCAISSE=' + ANumZ, True);
  if not FONbEntrees.EOF then
    if FONbEntrees.FindField('GJC_DATEOUV').AsDateTime >= FONbEntrees.FindField('GJC_DATEFERME').AsDateTime then
      ADate := FONbEntrees.FindField('GJC_DATEOUV').AsDateTime else
      ADate := FONbEntrees.FindField('GJC_DATEFERME').AsDateTime;
  Ferme(FONbEntrees);
  if DeDate > ADate then DeDate := ADate;
  // Calcul du nb d'entrées entre les 2 dates
  FONbEntrees := OpenSQL('SELECT SUM(GJE_NBENTREES) AS NBENTREES FROM JOURSETAB WHERE GJE_JOURNEE>="' + USDateTime(DeDate) +
    '" AND GJE_JOURNEE<="' + USDateTime(ADate) + '" AND GJE_ETABLISSEMENT="' + Etab + '"', True);
  if not FONbEntrees.EOF then Result := FONbEntrees.FindField('NBENTREES').AsInteger;
  Ferme(FONbEntrees);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : O. TARCY
Créé le ...... : 23/07/2001
Modifié le ... : 23/07/2001
Description .. : Récupère le nombre de ticket par vendeur pour le
Suite ........ : récapitulatif par vendeur
Suite ........ :
Suite ........ : Appeler depuis le générateur de ticket par
Suite ........ : '@FOGETPARAM(..;NBTICKETS;...'
Mots clefs ... : FO
*****************************************************************}

function FOGetNbTickets(Vendeur, MinZ, MaxZ, Caisse: string): double;
var SQL: string;
  nbTickets: integer;
  QQ: TQuery;
begin
  nbTickets := 0;
  SQL := 'SELECT DISTINCT GL_REPRESENTANT,GL_NUMERO ' +
    'FROM COMMERCIAL,PIECE ' +
    'LEFT OUTER JOIN LIGNE ON GL_NATUREPIECEG=GP_NATUREPIECEG ' +
    'AND GL_SOUCHE=GP_SOUCHE ' +
    'AND GL_NUMERO=GP_NUMERO ' +
    'AND GL_INDICEG=GP_INDICEG ' +
    'WHERE GL_TYPEARTICLE IN ("MAR","NOM") AND GP_NATUREPIECEG='+ FOGetNatureTicket(False, True) +' AND ' +
    'GL_REPRESENTANT="' + Vendeur + '" AND GCL_TYPECOMMERCIAL="VEN" AND ' +
    'GP_NUMZCAISSE <=' + MaxZ + ' AND GP_NUMZCAISSE >=' + MinZ + ' AND ' +
    FOSelectNumZDate('GP', StrToInt(MinZ), StrToInt(MaxZ), Caisse) + ' ' +
    'GROUP BY GL_REPRESENTANT,GL_NUMERO';
  QQ := OpenSQL(SQL, True);
  while not QQ.EOF do
  begin
    Inc(nbTickets);
    QQ.Next;
  end;
  Ferme(QQ);
  Result := StrToFloat(IntToStr(nbTickets));
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : O. TARCY
Créé le ...... : 23/07/2001
Modifié le ... : 23/07/2001
Description .. : Concatène et formate le nom et prénom d'un vendeur
Mots clefs ... : FO
*****************************************************************}

function FOConcateneVendeur(Nom: string): string;
var QQ: TQuery;
begin
  QQ := OpenSQL('SELECT GCL_LIBELLE,GCL_PRENOM FROM COMMERCIAL WHERE GCL_COMMERCIAL="' + Nom + '"', True);
  if not QQ.EOF then
  begin
    Nom := QQ.FindField('GCL_LIBELLE').AsString + ' ' + Copy(QQ.FindField('GCL_PRENOM').AsString, 1, 1) + '.';
    Result := Nom;
  end else
    Result := '';
  Ferme(QQ);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 23/07/2001
Modifié le ... : 23/07/2001
Description .. : Gère l'impression d'un montant en lettre sur plusieurs lignes
Mots clefs ... : FO
*****************************************************************}

function FOChiffreLettre(Montant: double; Devise: string; Taille, Ligne, MaxLigne: integer): string;
var Texte: string;
  Ind: integer;
  StgLst: TStringList;
  NoWrap: boolean;
begin
  Result := '';
  Texte := Trim(ChiffreLettre(Montant, FODonneMaskDevise(Devise)));
  if Length(Texte) > (MaxLigne * Taille) then
  begin
    if Ligne = 1 then Result := FOStrFMontant(Montant, 12, Devise);
  end else
  begin
    StgLst := TStringList.Create;
    StgLst.Text := WrapText(Texte, #13#10, ['.', ' ', #9, '-'], Taille - 1);
    if StgLst.Count <= MaxLigne then
    begin
      NoWrap := False;
      for Ind := 0 to StgLst.Count - 1 do if Length(StgLst.Strings[Ind]) > Taille then NoWrap := True;
    end else NoWrap := True;
    if NoWrap then
    begin
      StgLst.Clear;
      while Length(Texte) > 0 do
      begin
        StgLst.Add(Copy(Texte, 1, Taille));
        Delete(Texte, 1, Taille);
      end;
    end;
    if (Ligne > 0) and (Ligne <= StgLst.Count) then Result := Trim(StgLst.Strings[Ligne - 1]);
    StgLst.Free;
  end;
  // Ajout du caractère * devant et derrière
  if (Taille - Length(Result)) > 1 then
    Result := '*' + Result;
  if (Taille - Length(Result)) > 0 then
    Result := Result + StringOfChar('*', (Taille - Length(Result)));
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 14/01/2003
Modifié le ... : 14/01/2003
Description .. : Extrait le nombre saisi pour un code pièce ou billet
Mots clefs ... : FO
*****************************************************************}

function FOExtractPieceBillet(PieceBillet, Code, DetailPiecBil: string): string;
var St1, St2: string;
begin
  Result := '';
  St1 := Copy(PieceBillet, 1, 1) + Code + '=';
  while DetailPiecBil <> '' do
  begin
    St2 := ReadTokenSt(DetailPiecBil);
    if St1 = Copy(St2, 1, Length(St1)) then
    begin
      Result := Copy(St2, Length(St1) + 1, MaxInt);
      break;
    end;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 25/11/2002
Modifié le ... : 25/11/2002
Description .. : Extraction des arguments des fonctions @
Mots clefs ... : FO
*****************************************************************}

function ReadParam(var St: string): string;
var Ind: integer;
begin
  Ind := Pos(';', St);
  if Ind <= 0 then Ind := Length(St) + 1;
  Result := Copy(St, 1, Ind - 1);
  Delete(St, 1, Ind);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : N. ACHINO
Créé le ...... : 25/11/2002
Modifié le ... : 25/11/2002
Description .. : Fonctions @ pour le générateur d'état et de ticket
Mots clefs ... : FO
*****************************************************************}

function FOCalcOLEEtat(sf, sp: string): variant;
var
  s1, s2, s3, s4, s5: string;
begin
  Result := '';
  if sf = 'FOGETPARAM' then
  begin
    s1 := ReadParam(sp);
    s2 := ReadParam(sp);
    s3 := ReadParam(sp);
    s4 := ReadParam(sp);
    s5 := ReadParam(sp);
    if s1 = 'CAISSE' then
      Result := FOCaisseCourante // Numéro de caisse
    else if s1 = 'DEVISE' then
      Result := V_PGI.SymbolePivot // Devise de tenue du dossier
    else if s1 = 'NUMZ' then
      Result := FOGetNumZCaisse(FOCaisseCourante, 'MAX') // Numéro de ticket Z
    else if s1 = 'DIVERS' then
      Result := FONbEnrgt(s2, s3, s4, s5)
    else if s1 = 'VENDEURS' then
      Result := FOConcateneVendeur(s2) // Nombre d'enregistrements pour le champ s2
    else if s1 = 'NBENTREES' then
      Result := FOGetNbEntrees(s2, s3, s4, s5) // Nombre d'entrées
    else if s2 = 'NBTICKETS' then
      Result := FOGetNbTickets(s1, s3, s4, s5); // Nombre de tickets pour récap vendeur
  end
  else if sf = 'FOCHFLET' then
  begin
    s1 := ReadParam(sp); // Montant
    s2 := ReadParam(sp); // Devise
    s3 := ReadParam(sp); // Taille de la zone
    s4 := ReadParam(sp); // Ligne
    s5 := ReadParam(sp); // Nombre de lignes
    Result := FOChiffreLettre(Valeur(s1), s2, StrToInt(s3), StrToInt(s4), StrToInt(s5));
  end else
  if sf = 'FOGETARTDIM' then
  begin
    s1 := ReadParam(sp); // Code article
    s2 := ReadParam(sp); // Impression des libellés abrégés
    s3 := ReadParam(sp); // Impression compacte
    s4 := ReadParam(sp); // Titres des dimensions
    s5 := ReadParam(sp); // Séparateurs
    Result := FOGetArtDim(s1, (s2 <> '-'), (s3 <> '-'), (s4 <> '-'), s5);
  end else
  if sf = 'FOGETPCAIS' then
  begin
    s1 := ReadParam(sp); // Nom du champ
    s2 := ReadParam(sp); // Libellé, Abrégé ou code
    Result := FOGetPCaisse(s1, s2);
  end else
  if sf = 'FOGETDEMARQUE' then
  begin
    s1 := ReadParam(sp); // Code démarque
    s2 := ReadParam(sp); // 1=Client obligatoire ou 2=Imprimable en caisse (CC_LIBRE)
    Result := FOGetDemarque(s1, s2);
  end else
  if sf = 'FOEXTRACTPIECEBILLET' then
  begin
    s1 := ReadParam(sp); // Type pièce ou billet
    s2 := ReadParam(sp); // Code de la pièce ou du billet
    s3 := sp; // Détail des pièces et des billets
    Result := FOExtractPieceBillet(s1, s2, s3);
  end else
  if sf = 'FOSOLDECLIENT' then
  begin
    s1 := ReadParam(sp); // Code Tiers
    s2 := ReadParam(sp); // pour tous les établissements
    if (s1 <> '') and (s1 <> FODonneClientDefaut) then
      Result := FOCalculSoldeClient(s1, (s2 = 'X'));
  end else
  if sf = 'FOTOTALMODEPAIE' then
  begin
    s1 := ReadParam(sp); // Préfixe
    s2 := ReadParam(sp); // Caisse
    s3 := ReadParam(sp); // Z de
    s4 := ReadParam(sp); // Z à
    Result := FOMakeWhereRecapVendeurs(s1, s2, Valeuri(s3), Valeuri(s4))+ ' AND MP_TYPEMODEPAIE<>"'+ TYPEPAIERESTEDU +'"';
  end else
  if sf= 'FORECUPLIBFAMN2' then
  begin
    s1 := ReadParam(sp);
    Result := RechDom('GCFAMILLENIV2',s1,false);
  end;
end;

end.
