unit UtilRessource;

interface
uses
  ParamSoc,
  HEnt1,
  HCtrls,
  Utob,
  HStatus,
{$IFNDEF EAGLCLIENT}
{$IFNDEF DBXPRESS}dbtables {BDE},   
{$ELSE}uDbxDataSet,
{$ENDIF}
{$ENDIF !EAGLCLIENT}

  // pour compil dans GI/GA et éviter à PGIMAJVER d'avoir des includes inutilisés
{$IFDEF EAGLCLIENT}
  Maineagl,
{$ELSE}
  db,
{$IFNDEF EAGLSERVER}
{$IFNDEF ERADIO}
  Fe_Main,
{$ENDIF !ERADIO}
{$ENDIF EAGLSERVER}
{$ENDIF}
{$IFNDEF EAGLSERVER}
{$IFNDEF ERADIO}
  LookUp,
  UtofRessource_Mul,
{$ENDIF !ERADIO}
{$ENDIF EAGLSERVER}
  sysutils ,Utom
  ;

function GetRessourceRecherche(G_CodeRessource: THCritMaskEdit;
  stWhere, stRange, stMul: string): string;
{$IFNDEF EAGLSERVER}
{$IFNDEF ERADIO}
procedure GetRessourceMul(G_CodeRessource: THCritMaskEdit; stWhere, stRange, stMul: string);
procedure GetRessourceLookUp(G_CodeRessource: THCritMaskEdit; stWhere: string);
function GetRechRessource(G_CodeRessource: THCritMaskEdit;
  stWhere, stRange, stMul: string): string;
{$ENDIF !ERADIO}
{$ENDIF EAGLSERVER}

procedure LibelleRessource(CodeRess: string; var lib1, lib2: string);
//FV1 : 09/01/2018 - FS#2840 - FLIPO - afficher l'adresse mail du responsable à l'édition du tableau de bord
procedure EMailRessource(CodeRess: string; var Email : string);
function ExisteRessource(Coderess: string): boolean;

function ExisteLienSalarie(Slr: string; Rsc: string = ''): boolean;
function JaiLeDroitRessource(pStRes: string): Boolean; //mcd 14/03/2005 // C.B déplacé 18/04/2005
{ Creation de Ressource via une tob externe }
function CreateRessourceFromTob(Ressource: String; TobDET: Tob): String;
function UpdateRessourceFromTob(Ressource: String; TobDET: Tob): String;

function GetChampsRessource(CodeRessource, NomChamp: String): String;

//FV1 : 12/02/2014 - FS#706 - NCN : Dans la fiche ressource, le nom du salarié associé ne s'affiche pas                                                                   
procedure LibelleSalarie(CodeRess: string; var lib1, lib2: string);

const // Zones communes entre le Salarié et la Ressource
  RSLinkedFields: array[0..14, 0..1] of string = (
    ('ARS_LIBELLE', 'PSA_LIBELLE'),
    ('ARS_LIBELLE2', 'PSA_PRENOM'),
    ('ARS_ADRESSE1', 'PSA_ADRESSE1'),
    ('ARS_ADRESSE2', 'PSA_ADRESSE2'),
    ('ARS_ADRESSE3', 'PSA_ADRESSE3'),
    ('ARS_CODEPOSTAL', 'PSA_CODEPOSTAL'),
    ('ARS_VILLE', 'PSA_VILLE'),
    ('ARS_PAYS', 'PSA_PAYS'),
    ('ARS_ETABLISSEMENT', 'PSA_ETABLISSEMENT'),
    ('ARS_TELEPHONE', 'PSA_TELEPHONE'),
    ('ARS_TELEPHONE2', 'PSA_PORTABLE'),
    ('ARS_AUXILIAIRE', 'PSA_AUXILIAIRE'),
    ('ARS_IMMAT', 'PSA_NUMEROSS'),
    ('ARS_STANDCALEN', 'PSA_CALENDRIER'),
    ('ARS_CALENSPECIF', 'PSA_STANDCALEND')
    );

  // mcd  16/07/02 fct reprise de Activiteutil pour éviter include en cascade  si pas GA
function AFConversionUnite(AcsCodeSource, AcsCodeCible: string; AcdValeurSource: Double): Double;
function ConversionUnite(AcsCodeSource, AcsCodeCible: string; AcdValeurSource: Double): Double;
procedure ChangeUnite;

implementation
uses
  EntGC
  {$IFDEF PGISIDE}
  ,UtilGC
  {$ENDIF PGISIDE}
;

{***********************Recherche d'Ressource ******************************************}

//mcd 27/03/2006 pour création ressource depuis TOB (pgiside)
function CreateRessourceFromTob(Ressource: String; TobDet: Tob): String;
var
  iChamp: Integer;
  FieldName: String;
  Exist, IsValid: Boolean;
  TobRessource : Tob;
  TomRessource: Tom;
begin
  Result := '';
  { Contrôle existence du Ressource }
  if ExisteSQL('Select ARS_RESSOURCE from RESSOURCE where ARS_RESSOURCE="' + Ressource + '"') then
  begin
    Result := UpdateRessourceFromTob(Ressource, TobDet);
    Exit;
  end;

  if TobDet <> nil then
  begin
    { Crée une TobRessource et contrôle l'existence des champs de la tob passée en paramêtre }
    TobRessource := Tob.Create('RESSOURCE', nil, -1);
    try
      { Recopie la tob à mettre à jour dans la tob Ressource }
      iChamp := 999; Exist := True;
      while (iChamp < (1000 + (TobDet.ChampsSup.Count - 1))) and Exist do
      begin
        Inc(iChamp);
        { Vérifie si le champ fait partie de la table Ressource }
        if Copy(TobDet.GetNomChamp(iChamp), 0, Pos('_', TobDet.GetNomChamp(iChamp)) - 1) = 'T' then
          Exist := TobRessource.FieldExists(TobDet.GetNomChamp(iChamp));
      end;
      if Exist then
      begin
        { Vérifie les données en passant par une TomArticle }
        TomRessource := CreateTOM('RESSOURCE', nil, false, true);
        TomRessource.InitTOB(TobRessource);
        { Recopie la tob à mettre à jour dans la tobR essource }
        iChamp := 999;
        while (iChamp < (1000 + (TobDet.ChampsSup.Count - 1))) do
        begin
          Inc(iChamp);
          FieldName := TobDet.GetNomChamp(iChamp);
          TobRessource.PutValue(FieldName,TobDet.GetValue(FieldName));
        end;
					//on force les valeurs des dates et boolean, comme dans OnNewRecord TOm
        if TobRessource.GetValue('ARS_FINDISPO') = iDate1900 then
          TobREssource.PutValue('ARS_FINDISPO', iDate2099);
        if TobRessource.GetValue('ARS_DATESORTIE') = iDate1900 then
          TobREssource.PutValue('ARS_DATESORTIE', iDate2099);
				if TobRessource.getValue('ARS_TYPERESSOURCE')='SAL'  then
					TobRessource.putvalue('ARS_ESTHUMAIN','X');
				if TobRessource.getValue('ARS_MENSUALISE')='-'  then
					TobRessource.putvalue('ARS_MENSUALISE','X');  //pour même init que OnNEwRecord, sinon oblige à avoir champ dans fichier
{$ifdef GCGC}
        if (VH_GC.AFResCalculPR) then
          TobRessource.putvalue('ARS_CALCULPR', 'X')
        else
          TobRessource.putvalue('ARS_CALCULPR', '-');
{$endif}
        try
          IsValid := TomRessource.VerifTOB(TobRessource);
          Result := TomRessource.LastErrorMsg;
        finally
          TomRessource.Free;
        end;
        if IsValid then
        begin
          try
            TobRessource.InsertDB(nil, False); { Enregistre la TobRessource }
            TobDet.Dupliquer(TobRessource, False, True);
          except
            on E: Exception do
            begin
              {$IFDEF PGISIDE}
              MAJJnalEvent('PSE', 'ERR', TraduireMemoire('Mise à jour de la table RESSOURCE'), E.Message);
              {$ENDIF PGISIDE}
              Result := E.Message;
            end;
          end;
        end;
      end
      else
        Result := TraduireMemoire('Le champ : ') + FieldName + TraduireMemoire(' n''existe pas dans la table : ') + 'RESSOURCE';
    finally
      TobRessource.Free;
    end;
  end
  else
    Result := TraduireMemoire('Paramètres d''appel incorrect (Tob non renseignée)');
  if Result <> '' then
    TobDet.AddChampSupValeur('Error', Result);
end;

//==============================================================================
//  fonction de creation et/ou modif d'un Ressource à partir d'une tob.
//  elle applique les regles de validation definies dans la tom Ressource
//  et retourne un message d'erreur ainsi qu'un champ Error eventuel.
//  les champs de donnees à renseigner sont fournis en tant que champs sup
//  dans TobDet.
//==============================================================================
function UpdateRessourceFromTob(Ressource: String; TobDet: Tob): String;
var
  iChamp: Integer;
  FieldName: String;
  RessourceExist, Exist, IsValid: Boolean;
  TobRessource: Tob;
  TomRessource: Tom;
begin
  Result := '';
  { Contrôle existence du tiers }
  RessourceExist := ExisteSQL('Select ARS_RESSOURCE from RESSOURCE where ARS_RESSOURCE="' + Ressource + '"');

  if TobDet <> nil then
  begin
    { Crée une TobRessource et contrôle l'existence des champs de la tob passée en paramêtre }
    TobRessource := Tob.Create('RESSOURCE', nil, -1);
    if RessourceExist then
      TobRessource.SelectDB('"' + Ressource + '"', nil);
    try
      { Recopie la tob à mettre à jour dans la tob Ressource }
      iChamp := 999; Exist := True;
      while (iChamp < (1000 + (TobDet.ChampsSup.Count - 1))) and Exist do
      begin
        Inc(iChamp);
        { Vérifie si le champ fait partie de la table Ressource }
        if Copy(TobDet.GetNomChamp(iChamp), 0, Pos('_', TobDet.GetNomChamp(iChamp)) - 1) = 'T' then
          Exist := TobRessource.FieldExists(TobDet.GetNomChamp(iChamp)) ;
      end;
      if Exist then
      begin
        { Vérifie les données en passant par une TomArticle }
        TomRessource := CreateTOM('RESSOURCE', nil, false, true);
        { Recopie la tob à mettre à jour dans la tob article }
        iChamp := 999;
        while (iChamp < (1000 + (TobDet.ChampsSup.Count - 1))) do
        begin
          Inc(iChamp);
          FieldName := TobDet.GetNomChamp(iChamp);
          TobRessource.PutValue(FieldName, TobDet.GetValue(FieldName));
        end;
        try
          IsValid := TomRessource.VerifTOB(TobRessource);
          Result := TomRessource.LastErrorMsg;
        finally
          TomRessource.Free;
        end;
        if IsValid then
        begin
          try
            TobRessource.InsertOrUpdateDB; { Enregistre la TobRessource }
            TobDet.Dupliquer(TobRessource, False, True);
          except
            on E: Exception do
            begin
              {$IFDEF PGISIDE}
              MAJJnalEvent('PSE', 'ERR', TraduireMemoire('Mise à jour de la table RESSOURCE'), E.Message);
              {$ENDIF PGISIDE}
              Result := E.Message;
            end;
          end;
        end;
      end
      else
        Result := TraduireMemoire('Le champ : ') + FieldName + TraduireMemoire(' n''existe pas dans la table : ') + 'RESSOURCE';
    finally
      TobRessource.Free;
    end;
  end
  else
    Result := TraduireMemoire('Paramètres d''appel incorrect (Tob non renseignée)');
  if Result <> '' then
    TobDet.AddChampSupValeur('Error', Result);
end;


{***********A.G.L.***********************************************
Auteur  ...... : Dominique BROSSET
Créé le ...... : 27/03/2000
Modifié le ... :   /  /
Description .. : Rercherche d'un Ressource
Mots clefs ... : RECHERCHE;Ressource
*****************************************************************}

function GetRessourceRecherche(G_CodeRessource: THCritMaskEdit; stWhere, stRange, stMul: string): string;
var G_Ressource: THCritMaskEdit;
begin

  Result := '';

  if (VH_GC.AFRechResAv) then
    GetRessourceMul(G_CodeRessource, stWhere, stRange, stMul)
  else
    GetRessourceLookUp(G_CodeRessource, stwhere);

  Result := G_CodeRessource.Text;


  {*

  G_Ressource:=Nil;

  if G_CodeRessource = nil then
  begin
    G_Ressource := THCritMaskEdit.create(nil);
    G_Ressource.Text := '';
  end
  else
    G_Ressource := nil;

  if (G_CodeRessource = nil) then
  begin
    if (G_CodeRessource = nil) or (VH_GC.AFRechResAv) then
    begin
      GetRessourceMul(G_Ressource, stWhere, stRange, stMul);
      Result := G_Ressource.Text;
    end
    else
    begin
      GetRessourceMul(G_CodeRessource, stWhere, stRange, stMul);
      Result := G_CodeRessource.Text;
    end;
  end
  else
  begin
    GetRessourceLookUp(G_CodeRessource, stwhere);
  end;

  Result := G_CodeRessource.Text;


  if G_CodeRessource = nil then G_Ressource.Free;
  *}

end;

{$IFNDEF EAGLSERVER}
{$IFNDEF ERADIO}
procedure GetRessourceMul(G_CodeRessource: THCritMaskEdit; stWhere, stRange, stMul: string);
var
  CodeRes: string;
begin
  CodeRes := '';
  //if ctxAffaire in V_PGI.PGIContexte then Domaine := 'AFF' else Domaine := 'GC';
  //if stMul <> '' then CodeRes := AGLLanceFiche (Domaine, stMul, stWhere, '', stRange)
  //else
  if stRange = 'GRC' then
    CodeRes := AGLLanceFiche('RT', 'RTRessource_MUL', stWhere, '', stRange)
  else
    //mcd 12/06/02CodeRes := AGLLanceFiche ('AFF', 'RessourceRECH_MUL', stWhere, '', stRange);
    CodeRes := AFLanceFiche_Rech_Ressource(stRange,stWhere);
  //
  if CodeRes <> '' then // Coderes pour laisser la valeur initiale si pas de selection
  begin
    G_CodeRessource.Text := CodeRes;
  end;
end;

procedure GetRessourceLookUp(G_CodeRessource: THCritMaskEdit; stWhere: string);
var
  G_Ressource: THCritMaskEdit;
  stW,WW,Where : string;
begin

  G_Ressource := G_CodeRessource;

  //stW := stWhere;

  {*
  WW := READTOKENPipe (stW,'=');
  if stw <> '' then
  begin
    if Pos('"',stW) > 0 then stW := StringReplace_(stW,'"','',[rfReplaceAll]);
  	Where := WW+'="'+Stw+'"'
  end else
  begin
  	Where := '';
  end;
  *}

  LookupList(G_Ressource, 'RESSOURCES', 'RESSOURCE', 'ARS_RESSOURCE', 'ARS_LIBELLE', stWhere, 'ARS_RESSOURCE', False, 6);

  if (G_Ressource.Text <> '') and (G_Ressource.Text <> G_CodeRessource.Text) then G_CodeRessource.Text := G_Ressource.Text;
end;


function GetRechRessource(G_CodeRessource: THCritMaskEdit;
  stWhere, stRange, stMul: string): string;
var
  G_Ressource: THCritMaskEdit;
begin
  Result := '';
  if G_CodeRessource = nil then
  begin
    G_Ressource := THCritMaskEdit.create(nil);
    G_Ressource.Text := '';
  end
  else
    G_Ressource := nil;

  if (G_CodeRessource = nil){$IFDEF GCGC} or (VH_GC.AFRechResAv){$ENDIF} then
    // pour la Paie
  begin
    if G_CodeRessource = nil then
    begin
      GetRessourceMul(G_Ressource, stWhere, stRange, stMul);
      Result := G_Ressource.Text;
    end
    else
    begin
      GetRessourceMul(G_CodeRessource, stWhere, stRange, stMul);
      Result := G_CodeRessource.Text;
    end;
  end
  else
  begin
    LookUpCombo(G_CodeRessource);
    Result := G_CodeRessource.Text;
  end;
  if G_CodeRessource = nil then
    G_Ressource.Free;
end;
{$ENDIF !ERADIO}
{$ENDIF EAGLSERVER}

{***********A.G.L.***********************************************
Auteur  ...... : G.Merieux
Créé le ...... : 16/10/2001
Modifié le ... :   /  /
Description .. : Recherche des libelle 1 et 2 de la ressource
Mots clefs ... : RESSOURCE       
*****************************************************************}

procedure LibelleRessource(CodeRess: string; var lib1, lib2: string);
var QQ : TQuery;
    Req: string;
begin

  lib1 := '';
  lib2 := '';

  Req := 'SELECT ARS_LIBELLE,ARS_LIBELLE2 FROM RESSOURCE WHERE ARS_RESSOURCE="' + CodeRess + '"';

  QQ := OpenSQL(Req, false);

  if not QQ.EOF then
  begin
    lib1 := QQ.findField('ARS_LIBELLE').asString;
    lib2 := QQ.findField('ARS_LIBELLE2').asString;
  end; // mcd 30/04/02

  ferme(QQ);
end;
//FV1 : 09/01/2018 - FS#2840 - FLIPO - afficher l'adresse mail du responsable à l'édition du tableau de bord
procedure EMailRessource(CodeRess: string; var Email : string);
var QQ : TQuery;
    Req : String;
begin

  Email := '';

  Req := 'SELECT ARS_EMAIL FROM RESSOURCE WHERE ARS_RESSOURCE="' + CodeRess + '"';

  QQ := OpenSQL(Req, false);

  if not QQ.EOF then
  begin
    Email := QQ.findField('ARS_EMAIL').asString;
  end;

  ferme(QQ);
end;

{***********A.G.L.***********************************************
Auteur  ...... : MC DESSEIGNET
Créé le ...... : 19/02/2003
Modifié le ... :   /  /
Description .. : fonction squi regarde si une ressource existe
Suite ........ : Utiliser pour tester validité champ libre ressource
Mots clefs ... : RESSOURCE;EXISTE
*****************************************************************}

function ExisteRessource(CodeRess: string): boolean;
begin
  Result := ExisteSql('SELECT ARS_RESSOURCE FROM RESSOURCE WHERE ARS_RESSOURCE="' + CodeRess + '"');
end;

{***********A.G.L.***********************************************
Auteur  ...... : TG
Créé le ...... : 04/04/2003
Modifié le ... : 04/04/2003
Description .. : Teste si le salarié Slr est associé à une ressource
Suite ........ : [autre que Rsc si spécifié]
Mots clefs ... : RESSOURCE;SALARIE;LIEN
*****************************************************************}

function ExisteLienSalarie(Slr: string; Rsc: string = ''): boolean;
begin
  result := GetParamSocSecur('SO_PGLIENRESSOURCE', false) and
    (Slr <> '') and
    ExisteSQL('SELECT ARS_RESSOURCE FROM RESSOURCE WHERE ARS_SALARIE="' + Slr + '" AND ARS_RESSOURCE<>"' + Rsc + '"');
end;

{***********A.G.L.***********************************************
Auteur  ...... : Pierre Lenormand
Créé le ...... : 16/03/2006
Modifié le ... :   /  /
Description .. : Pour une unité source, une unité cible, une valeur source en entrées, renvoie la valeur cible en sortie = valeur source exprimée en unité cible
Mots clefs ... : CONVERSION; UNITE; QUOTITE; MEA ; GIGA
*****************************************************************}
function ConversionUnite(AcsCodeSource, AcsCodeCible: string; AcdValeurSource: Double): Double;
// exemples :
// 2 Jours en heures ? : ConversionUnite('J', 'H', 2) renvoie 16  (2 * 8 heures/jour)
// 1 semaine en jours ? : ConversionUnite('S', 'J', 1) renvoie 5   (1 * 40 heures/semaine divisé par 8 heures/jour)
// 1 heure en semaine ? : ConversionUnite('H', 'S', 1) renvoie 1/40   (1 * 1 heure divisé par 40 heures/semaine)
var
  QS, QC: TQuery;
begin
  Result := AcdValeurSource;

  if (AcsCodeSource = '') or (AcsCodeCible = '') then
    exit; // pas d'untité, inutile de passer par la fct

  Result := 0;

  if AcdValeurSource = 0.0 then
    Exit;

  QS := nil;
  QC := nil;
  try
    // SELECT * : nombre de champs limité et nombre d'enregistrements restreint
    QS := OpenSQL('SELECT GME_QUALIFMESURE,GME_QUOTITE FROM MEA WHERE GME_MESURE="' + AcsCodeSource + '"', True);
    if QS.Eof then
      Exit;
    QC := OpenSQL('SELECT GME_QUALIFMESURE,GME_QUOTITE FROM MEA WHERE GME_MESURE="' + AcsCodeCible + '"', True);
    if QC.Eof then
      Exit;

    // Si les deux codes d'entrée ne sont pas du même type, on sort
    if (QS.FindField('GME_QUALIFMESURE').AsString <> QC.FindField('GME_QUALIFMESURE').AsString) then
      exit;
    // si l'une des quotités est nulle, on renvoit 0
    if (QS.FindField('GME_QUOTITE').AsFloat = 0.0) or (QC.FindField('GME_QUOTITE').AsFloat = 0.0) then
      exit;

    // Calcul de la valeur de sortie
    Result := (AcdValeurSource * QS.FindField('GME_QUOTITE').AsFloat) / QC.FindField('GME_QUOTITE').AsFloat;
  finally
    Ferme(QS);
    Ferme(QC);
  end;
end;


{***********A.G.L.***********************************************
Auteur  ...... : Pierre Lenormand
Créé le ...... : 31/01/2000
Modifié le ... :   /  /
Description .. : Pour une unité source, une unité cible, une valeur source en entrées, renvoie la valeur cible en sortie = valeur source exprimée en unité cible
Mots clefs ... : CONVERSION; UNITE; QUOTITE; MEA ; GIGA
*****************************************************************}

function AFConversionUnite(AcsCodeSource, AcsCodeCible: string; AcdValeurSource: Double): Double;
// exemples :
// 2 Jours en heures ? : ConversionUnite('J', 'H', 2) renvoie 16  (2 * 8 heures/jour)
// 1 semaine en jours ? : ConversionUnite('S', 'J', 1) renvoie 5   (1 * 40 heures/semaine divisé par 8 heures/jour)
// 1 heure en semaine ? : ConversionUnite('H', 'S', 1) renvoie 1/40   (1 * 1 heure divisé par 40 heures/semaine)
var
	TOBSource, TOBCible: TOB;
  dQuotS, dQuotC: double;
begin
  Result := AcdValeurSource;
// PL le 27/01/06 : correction pour éviter le cas où l'on n'avait pas d'unité
// j'ai aussi passé le Result := 0 après...
  if (AcsCodeSource = '') or (AcsCodeCible = '') then
    exit; // pas d'unité, inutile de passer par la fct

  Result := 0;

{$IFDEF GCGC} //pour la paie

    if (VH_GC.MTOBMEA <> nil) then
    begin

      TOBSource := VH_GC.MTOBMEA.FindFirst(['GME_MESURE'], [AcsCodeSource], true);

      TOBCible := VH_GC.MTOBMEA.FindFirst(['GME_MESURE'], [AcsCodeCible], true);

      if (TOBSource = nil) or (TOBCible = nil) then
        exit;

      // Si les deux codes d'entrée ne sont pas du même type, on sort
      if (TOBSource.GetValue('GME_QUALIFMESURE') <> TOBCible.GetValue('GME_QUALIFMESURE')) then
        exit;

      dQuotS := TOBSource.GetValue('GME_QUOTITE');
      dQuotC := TOBCible.GetValue('GME_QUOTITE');
      // si l'une des quotité est nulle, on renvoit 0
      if (dQuotS = 0) or (dQuotC = 0) then
        exit;

      // Calcul de la valeur de sortie
      Result := (AcdValeurSource * dQuotS) / dQuotC;
    end
    else // si (VH_GC.TOBMEA = nil)
{$ENDIF}
    	Result := ConversionUnite(AcsCodeSource, AcsCodeCible, AcdValeurSource);

end;


procedure ChangeUnite;
var
  TobRes, TobACT, TobDEt: TOB;
  QQ1, QQ2: Tquery;
  ii, jj: integer;
  Qte: double;
begin
  TobRes := nil;
  QQ1 := nil;
  try
    SourisSablier;
    InitMove(3, '');

    if not (ExisteSQL('SELECT ACT_TYPEACTIVITE FROM ACTIVITE')) then
      exit; // rien dans la table activite dans la base, on sort...
    MoveCur(False);

    // mise à jour de la zone qteUniteref  , si même unité que dans les paramètres
    try
      ExecuteSql('UPDATE ACTIVITE SET ACT_QTEUNITEREF=ACT_QTE WHERE ACT_TYPEARTICLE="PRE" AND ACT_UNITE="' +
        GetParamSocSecur('SO_AFMESUREACTIVITE', '') + '"');
    except
      V_PGI.IoError := oeSaisie;
    end;

    MoveCur(False);

    if V_PGI.IoError <> oeOk then
      exit;

    TobRes := Tob.Create('les ress', nil, -1);
    // on traite l'activité par ressource pour éviter de récupérer trop d'enrgt à la fois
    QQ1 := OpenSql('SELECT ARS_RESSOURCE FROM RESSOURCE', true);

    if not (QQ1.eof) then
    begin
      TOBRes.LoadDetailDB('RESSOURCE', '', '', QQ1, False);
      TOBDet := TOB.Create('Les Ressources', TobRes, -1);
      TOBDet.LoadFromSt('ARS_RESSOURCE;', ';', ';', ';');

      InitMove(TobRes.Detail.count, '');
      for ii := 0 to TobRes.Detail.count - 1 do
      begin
        if V_PGI.IoError <> oeOk then
          exit;
        QQ2 := nil;
        try
          // PL le 15/04/03 : modif clé activite
    //      QQ2 := Opensql('SELECT ACT_TYPEACTIVITE,ACT_AFFAIRE,ACT_RESSOURCE,ACT_DATEACTIVITE,ACT_FOLIO,ACT_TYPEARTICLE,ACT_NUMLIGNE,ACT_QTE,ACT_QTEUNITEREF,ACT_UNITE FROM ACTIVITE WHERE ACT_RESSOURCE="'
    //            + TobRes.detail[ii].GetValue('ARS_RESSOURCE')+'" AND ACT_TYPEARTICLE="PRE" AND ACT_UNITE<>"" AND ACT_UNITE <>"'
    //            + GetParamSoc ('SO_AFMESUREACTIVITE') +'"',True);
          QQ2 := Opensql('SELECT ACT_TYPEACTIVITE,ACT_AFFAIRE,ACT_NUMLIGNEUNIQUE,ACT_QTE,ACT_QTEUNITEREF,ACT_UNITE FROM ACTIVITE WHERE ACT_RESSOURCE="'
            + TobRes.detail[ii].GetValue('ARS_RESSOURCE') + '" AND ACT_TYPEARTICLE="PRE" AND ACT_UNITE<>"" AND ACT_UNITE <>"'
            + GetParamSocSecur('SO_AFMESUREACTIVITE', '') + '"', True);
          if not (QQ2.Eof) then
          begin
            TobAct := Tob.Create('les Act', nil, -1);
            try
              TobAct.LOadDetailDb('ACTIVITE', '', '', QQ2, False);
              for jj := 0 to TobAct.Detail.count - 1 do
              begin
                Tobdet := TObAct.detail[jj];
                Qte := Arrondi(ConversionUnite(TOBDet.Getvalue('ACT_UNITE'),
                  GetParamSocSecur('SO_AFMESUREACTIVITE', ''), TobDet.Getvalue('ACT_QTE')), GetParamSocSecur('SO_DECQTE', 0));
                // PL le 16/01/03 : lorsque la fonction est lancée par le PGIMajver, la TOB ne tiens pas encore compte de la structure de
                // la table avec les nouveaux champs (ACT_QTEUNITEREF vient d'être créé). On est donc obligé de faire un UPDATE par ligne
                //            TobDet.Putvalue ('ACT_QTEUNITEREF', Arrondi(Qte, GetParamSoc ('SO_DECQTE')));

                try
                  // PL le 15/04/03 : modif clé activite
                  ExecuteSql('UPDATE ACTIVITE SET ACT_QTEUNITEREF=' + StrfPoint(Qte) +
                    ' WHERE ACT_TYPEACTIVITE="' + TobDet.Getvalue('ACT_TYPEACTIVITE') + '"'
                    + ' AND ACT_AFFAIRE="' + TobDet.Getvalue('ACT_AFFAIRE') + '"'
                    + ' AND ACT_NUMLIGNEUNIQUE=' + inttostr(TobDet.Getvalue('ACT_NUMLIGNEUNIQUE')));
                except
                  V_PGI.IoError := oeSaisie;
                end;
              end;

              //          if (Not TobAct.UpdateDb(false, true)) then
              //            V_PGI.IoError := oeUnknown;
              //////////////////////////////////////////////////////// Fin PL le 16/01/03
            finally
              TobAct.free;
            end;
          end;
        finally
          Ferme(QQ2);
        end;

        MoveCur(False);
      end; // boucle for
    end;

  finally
    Ferme(QQ1);
    TobRes.free;
    FiniMove;
    SourisNormale;
  end;
end;


function JaiLeDroitRessource(pStRes: string): Boolean;
var
  QQ: Tquery;
  vStEquipe: string;
  vStTypeRessource: string;

begin
  result := True;
  // confidentialité
  if (GetParamsocSecur('SO_AFTypeConPla', '') = 'EQU') and (pStRes <> '') then
  begin
    QQ := nil;
    try
      QQ := OpenSQL('SELECT ARS_TYPERESSOURCE, ARS_EQUIPERESS FROM RESSOURCE WHERE ARS_RESSOURCE="' + pStRes + '"', True);
      if not QQ.EOF then
      begin
        vStEquipe := QQ.FindField('ARS_EQUIPERESS').AsString;
        vStTypeRessource := QQ.FindField('ARS_TYPERESSOURCE').AsString;
{$IFDEF GCGC}
        // on ne fait le controle que pour les ressources qui ont une gestion d'équipe
        if pos(vStTypeRessource, GetParamSocSecur('SO_AFTYPERESSOURCE', '')) = 0 then
          if pos(vStEquipe, Vh_GC.AfEquipe) = 0 then
            Result := False;
{$ENDIF GCGC}
      end;
    finally
      Ferme(QQ);
    end;
  end;
end;

// BDU - 16/11/06, Récupération des valeurs de la table RESSOURCE
function GetChampsRessource(CodeRessource, NomChamp: String): String;
var
  Q: TQuery;
begin
  Result := '';
  if CodeRessource = '' then
    Exit;
  if not (ctxAffaire in V_PGI.PGIContexte) and not (ctxGCAFF in V_PGI.PGIContexte) then
    Exit;
  Q := nil;
  try
    Q := OPENSQL('SELECT ' + NomChamp + ' From RESSOURCE WHERE ARS_RESSOURCE="' +
      CodeRessource + '"', True);
    if not Q.EOF then
      Result := Q.Findfield(NomChamp).AsString;
  finally
    Ferme(Q);
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : f.VAUTRAIN
Créé le ...... : 16/10/2001
Modifié le ... :   /  /
Description .. : Recherche des libelle 1 et 2 du Salarié
Mots clefs ... : RESSOURCE
*****************************************************************}

procedure LibelleSalarie(CodeRess: string; var lib1, lib2: string);
var StSQL : String;
    QQ    : TQuery;
begin

  lib1 := '';
  lib2 := '';

  StSQL := 'SELECT PSA_LIBELLE,PSA_PRENOM FROM SALARIES WHERE PSA_SALARIE="' + CodeRess + '"';
  QQ := OpenSQL(StSQL, false);
  //
  if not QQ.EOF then
  begin
    lib1 := QQ.findField('PSA_LIBELLE').asString;
    lib2 := QQ.findField('PSA_PRENOM').asString;
  end;

  ferme(QQ);
  
end;

end.
