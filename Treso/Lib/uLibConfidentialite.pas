{ Gestion des confidentialités
--------------------------------------------------------------------------------------
    Version  |   Date  | Qui |   Commentaires
--------------------------------------------------------------------------------------
 8.10.001.004  08/08/07  JP   Création de l'unité
 
--------------------------------------------------------------------------------------}
unit uLibConfidentialite;


interface

uses
  {$IFDEF EAGLCLIENT}
  eMul, eQRS1,
  {$ELSE}
  Mul, QRS1,
  {$ENDIF EAGLCLIENT}
  Stat,Cube, HCtrls, Controls, HEnt1, StdCtrls, Classes, ComCtrls, UTob, UTOF;

const
  {Liste des types de confidentialité que l'on trouve dans la tablette TRTYPECONFIDENTIALITE
   et dans CO_ABREGE de la Tablette TRCHAMPSCONFIDENTIEL.
   Voir la méthode : TListeConf.InitConfidentialite}
  tyc_All         = '@AL'; {A ne pas utiliser dans les tablettes}
  tyc_Banque      = 'BQE';
  tyc_EtabBqe     = 'PQE';
  tyc_Flux        = 'FLX';
  {A compléter ....}

  MaxTypeConfidentialite = 2;

  TableauLeftJoin : array [0..1, 0..MaxTypeConfidentialite] of string  =
                    ((tyc_Banque, tyc_EtabBqe, tyc_Flux),
                    ('',
                     'PQ_BANQUE = BQ_BANQUE',
                     'TFT_FLUX = TE_CODEFLUX'));

  TableauTables : array [0..1, 0..MaxTypeConfidentialite] of string  =
                    ((tyc_Banque, tyc_EtabBqe, tyc_Flux   ),
                     ('BANQUECP', 'BANQUES', 'FLUXTRESO'));

  {Valeurs de la tablette TRPRODUITS}
  typ_Compta      = 'CPT';
  typ_Treso       = 'TRE';
  typ_ComptaTreso = 'CET';

  NOMCONFIDENTIEL = 'SO_TRUSERCONFIDENTIEL';

type
  {Objet qui contient la liste des types de confidentialité et mémorise si ils sont actifs
   Cet objet repose sur les constantes de type "tyc_" déclarées ci-dessus}
  TListeConf = class
  private
    FListeConfidentialite : TStringList;
    procedure MajConfidentialite (TypeConf : string; Actif : Boolean);
    procedure InitConfidentialite(Actif : Boolean = False);
  public
    {ListTypeConf : Type confidentialités que l'on veut les activer par défaut}
    constructor Create(ListTypeConf : string = '');
    destructor  Destroy; override;

    {Active les confidentialités passées en paramètre et séparées par des ";"}
    procedure   SetOnConfidentialite (ListeTypeConf : string);
    {Désactive les confidentialités passées en paramètre et séparées par des ";"}
    procedure   SetOffConfidentialite(ListeTypeConf : string);
    {Contrôle si la confidentialité passée en paramètre est active}
    function    GereConfidentialite  (TypeConf : string) : Boolean;
    {Retourne la liste des types de confidentialité actifs}
    function    GetListeTypeConf     : string;
  end;

  {Classe stockant toutes les enregistrements de PROSPECTCONF pour un utilisateur (que l'on
   gère les confidentialité par UserGrp ou par Utilisat) et pour un produit. Il est possible
   de limiter les confidentialités mémorisées en en donnant la liste au constructeur (TypeConf)}
  TObjConfidentialite = class
  private
    FWhereProduit : string;
    FIntervenant  : string;
    FListeConf    : TOB;

    procedure ChargeConfident   (TypeConf : string = '');
    function  GetLeftJoinFromTab (UnTypeConf : string) : string;
    class function GetProduit    (Produit  : TContexte_PGI) : string;
    class function GetIntervenant(Utilisat : string) : string;
  public
    constructor Create(CodeUser : string; Produit : TContexte_PGI = ctxTreso; TypeConf : string = '');
    destructor  Destroy; override;
    function    GetWhereSQLConf(TypeConf : string) : HString;
    function    GetLeftJoinConf(TypeConf : string) : string;
    function    GetChampsConf  (TypeConf : string) : string;
    {Fonction générique pour récupérer une clause Where sur les confidentialités par requête}
    class function GetWhereConf(CodeUser, TypeConf : string; Produit : TContexte_PGI = ctxTreso) : string;
  end;

  {Ancêtre pour les Muls, Cubes, Stats et QRS1.
   Pour gérer les confidentialités, il suffit dans la classe héritière de renseigner, avant le
   inherited, TypeConfidentialite. TypeConfidentialite est une succession de String3 séparés par des ";"
   De préférence, essayer d'utiliser les constantes de type 'tyc_' déclarées en entête de l'unité}
  TOFCONF = class(TOF)
  private
    FObjConfidentialite  : TObjConfidentialite;
    FObjTypeConf         : TListeConf;

    FTypeConfidentialite : string;
    FGereConfidentialite : Boolean;

    //function  GetNomTables : string;
    function  IsEcranAConfidentialite : Boolean;
    procedure InitConfident;
    procedure SetTypeConfidentialite(Value : string);
  public
    XX_WHERECONF         : THEdit;
    XX_WHEREJOIN         : THEdit;
    XX_FROM              : THEdit;

    procedure OnArgument(S : string); override;
    procedure OnUpdate              ; override;
    procedure OnClose               ; override;
    procedure OnLoad                ; override;

    property  TypeConfidentialite : string              read FTypeConfidentialite write SetTypeConfidentialite;
    property  ObjTypeConf         : TListeConf          read FObjTypeConf         write FObjTypeConf;
    property  ObjConfidentialite  : TObjConfidentialite read FObjConfidentialite;
  end;

{Debut : Méthode pour le paramétrage des confidentialités}
procedure GereTabletteUser(pControl : TControl);
procedure GereCaptionUser (pControl : TLabel);
function  FiltreTabletteProduit : string;
function  FiltreTabletteTypeCnf(aProduit : string) : string;
function  FiltreTableChamps    (aTypConf : string) : string;
function  GetLaDBListe         : string;
{Fin : Méthode pour le paramétrage des confidentialités}

{Fonction de codage du pagecontrol "Avancé" en SQL}
function  CodeSQL              (Champ, Operateur : string; Value : HString) : HString;
{Retourne la table concernée par un type de confidentialité}
function  GetTableFromConf     (aTypConf : string) : string;
{Purge de la table PROSPECTCONF des éléments issus de la compta et de la Tréso}
function  VideConfidentialites : Boolean;
{Retourne le contexte en fonction de l'exécutable}
function  GetExeContext        : TContexte_PGI;
{Test si la table passées en paramètre est une vue}
function  EstUneVue            (NomTable : string) : string;
{Ajoute un alias à une requête ou un morceau de requête}
function  AliasSQL(Requete, PfxChp, Alias : string) : string;


implementation

uses
  {$IFDEF MODENT1}
  CPVersion,
  {$ENDIF MODENT1}
  {$IFNDEF EAGLCLIENT}
    HDB,
    {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
  {$ENDIF EAGLCLIENT}
  UProcGen, ParamSoc, Ent1, SysUtils, HMsgBox;


{---------------------------------------------------------------------------------------}
procedure GereTabletteUser(pControl : TControl);
{---------------------------------------------------------------------------------------}
var
  lDataType : string;
begin
  if GetParamSocSecur(NOMCONFIDENTIEL, False) then lDataType := 'TTUSERGROUPE'
                                              else lDataType := 'TTUTILISATEUR';

       if (pControl is THEdit              ) then (pControl as THEdit              ).DataType := lDataType
  else if (pControl is THValComboBox       ) then (pControl as THValComboBox       ).DataType := lDataType
  else if (pControl is THMultiValComboBox  ) then (pControl as THMultiValComboBox  ).DataType := lDataType
  {$IFNDEF EAGLCLIENT}
  else if (pControl is THDBEdit            ) then (pControl as THDBEdit            ).DataType := lDataType
  else if (pControl is THDBValComboBox     ) then (pControl as THDBValComboBox     ).DataType := lDataType
  else if (pControl is THDBMultiValComboBox) then (pControl as THDBMultiValComboBox).DataType := lDataType
  {$ENDIF EAGLCLIENT}
end;

{---------------------------------------------------------------------------------------}
function  GetLaDBListe : string;
{---------------------------------------------------------------------------------------}
begin
  if GetParamSocSecur(NOMCONFIDENTIEL, False) then Result := 'TRMULCONFIDENTGRP'
                                              else Result := 'TRMULCONFIDENTUS';
end;

{---------------------------------------------------------------------------------------}
procedure GereCaptionUser(pControl : TLabel);
{---------------------------------------------------------------------------------------}
var
  lCaption : string;
begin
  if GetParamSocSecur(NOMCONFIDENTIEL, False) then lCaption := TraduireMemoire('Groupe utilisateurs')
                                              else lCaption := TraduireMemoire('Utilisateur');
  pControl.Caption := lCaption;
end;

{---------------------------------------------------------------------------------------}
function  FiltreTabletteProduit : string;
{---------------------------------------------------------------------------------------}
begin
  if (CtxTreso in V_PGI.PGIContexte) then Result := 'AND CO_CODE <> "' + typ_Compta + '"'
  else if not EstComptaTreso         then Result := 'AND CO_CODE = "' + typ_Compta + '"'
                                     else Result := 'AND CO_CODE <> "' + typ_Treso + '"';
end;

{---------------------------------------------------------------------------------------}
function  FiltreTabletteTypeCnf(aProduit : string) : string;
{---------------------------------------------------------------------------------------}
begin
  if aProduit = '' then begin
    if CtxTreso in V_PGI.PGIContexte then
      Result := 'AND CO_LIBRE <> "' + typ_Compta + '"'
    else if not EstComptaTreso then
      Result := 'AND CO_LIBRE = "' + typ_Compta + '"'
    else if CtxCompta in V_PGI.PGIContexte then
      Result := 'AND CO_LIBRE <> "' + typ_Treso + '"';
  end
  else
    Result := 'AND CO_LIBRE = "' + aProduit + '"';
end;

{Test si la table passées en paramètre est une vue}
{---------------------------------------------------------------------------------------}
function  EstUneVue(NomTable : string) : string;
{---------------------------------------------------------------------------------------}
var
  n : Integer;
begin
  Result := '';
  for n := Low(V_PGI.DEVues) to High(V_PGI.DEVues) do begin
    if UpperCase(NomTable) = UpperCase(V_PGI.DEVues[n].Nom) then begin
      Result := V_PGI.DEVues[n].SQL;
      Break;
    end;
  end;
end;

{Ajoute un alias à une requête ou un morceau de requête
 Exemple : AliasSQL(WHERE BQ_GENERAL = "RRR" AND BQ_NODOSSIER = "00000", BQ, BQE) renvoie
           WHERE BQE.BQ_GENERAL = "RRR" AND BQE.BQ_NODOSSIER = "00000"   
{---------------------------------------------------------------------------------------}
function AliasSQL(Requete, PfxChp, Alias : string) : string;
{---------------------------------------------------------------------------------------}
begin
  Result := Requete;
  if (pfxChp <> '') and (Alias <> '') then
    Result := TrTrouveEtRemplace(Result, PfxChp + '_', Alias + '.' + PfxChp + '_', True);
end;

{---------------------------------------------------------------------------------------}
function  FiltreTableChamps(aTypConf : string) : string;
{---------------------------------------------------------------------------------------}
begin
  if aTypConf = '' then Result := 'AND CO_ABREGE = "WWW"'
                   else Result := 'AND CO_ABREGE = "' + aTypConf + '"';
end;

{Retourne la table concernée par un type de confidentialité
{---------------------------------------------------------------------------------------}
function  GetTableFromConf(aTypConf : string) : string;
{---------------------------------------------------------------------------------------}
var
  n : Integer;
  T : string;
begin
  Result := '';
  if aTypConf[Length(aTypConf)] <> ';' then aTypConf := aTypConf + ';';
  T := ReadTokenSt(aTypConf);
  while T <> '' do begin
    for n := 0 to MaxTypeConfidentialite do begin
      if T = TableauTables[0, n] then begin
        Result := Result + TableauTables[1, n] + ',';
        Break;
      end;
    end;
    T := ReadTokenSt(aTypConf);
  end;
  System.Delete(Result, Length(Result), 1);
end;

{Retourne le contexte en fonction de l'exécutable
{---------------------------------------------------------------------------------------}
function GetExeContext : TContexte_PGI;
{---------------------------------------------------------------------------------------}
begin
       if ctxTreso  in V_PGI.PGIContexte then Result := ctxTreso
  else if CtxCompta in V_PGI.PGIContexte then Result := ctxCompta
                                         else Result := ctxCompta;
end;

{Valeurs possible de Operateur
 VV : n'est pas vide
 V  : est Vide
 M  : ne contient pas
 L  : contient
 J  : N'est pas dans
 I  : est dans
 G  : n'est pas entre
 E  : est entre
 D  : ne commnce pas par
 C  : Commence par
 et =, <>, <=, <, >=, >=
{---------------------------------------------------------------------------------------}
function CodeSQL(Champ, Operateur : string; Value : HString) : HString;
{---------------------------------------------------------------------------------------}
var
  St, Val1, Val2: hstring;
  Typ: string;
  ValIsNull, OkChamp: boolean;
  stCh : string;
begin
  if (Operateur = 'V') or (Operateur = 'VV') then
    Value := '??';

  if (Champ = '') or (Operateur = '') or (Value = '') then begin
    result := '';
    Exit;
  end;

  ValIsNull := (Value = 'NULL') or (Value = 'NUL');

  if ValIsNull then
    St := ''

  else if (Operateur = 'I') or (Operateur = 'J') then
    St := ''

  else if (Operateur = 'E') or (Operateur = 'G') then
    St := ' BETWEEN '

  else if (Operateur = 'C') or (Operateur = 'L') then
    St := ' LIKE '

  else if (Operateur = 'D') or (Operateur = 'M') then
    St := ' NOT LIKE '

  else if (Operateur = 'V') then
    st := ''

  else if (Operateur = 'VV') then
    st := ''

  else
    St := Operateur;

  {Vide, pas vide égale ou différent}
  if ((Value = ' ') and ((Operateur = '=') or (Operateur = '<>'))) or
     (Operateur = 'V') or (Operateur = 'VV') then begin
    {Il faut supprimer les . des champs DEBUT}
    stCh := Champ;

    while Pos('.', stCh) > 0 do
      ReadTokenPipe(stCh, '.');

    Typ := ChampToType(stCh);

    if (Typ = '??') or (Typ = 'INTEGER') or (Typ = 'SMALLINT') or
       (Typ = 'DOUBLE') or (Typ = 'RATE') or (Typ = 'EXTENDED') or
       (Typ = 'DATE') or (Typ = 'BLOB') then begin
      Result := '';
      exit;
    end;

    Value := '';

    if (Operateur = '=') or (Operateur = 'V') then
      St := ' = "" '

    else if (Operateur = '<>') or (Operateur = 'VV') then
      St := ' <> "" ';
  end

  else if ValIsNull then begin
    if Operateur = '=' then
      St := ' IS NULL '

    else if Operateur = '<>' then
      St := ' IS NOT NULL ';
  end

  {Between et not between}
  else if (Operateur = 'E') or (Operateur = 'G') then begin
    Typ := ChampToType(Champ);

    if Typ <> 'DATE' then
      Value := FindetReplace_(Value, ':', ';', True)

    else
      {Pour les dates avec heure il vaut mieux mettre des espace pour séparer les termes sinon ...}
      Value := FindetReplace_(Value, ' : ', ';', True);

    Value := FindetReplace_(Value, ' et ', ';', True);
    Value := FindetReplace_(Value, ' ET ', ';', True);
    Value := FindetReplace_(Value, ' Et ', ';', True);
    Val1 := ReadTokenSt(Value);
    Val2 := ReadTokenSt(Value);

    if Val2 = '' then
      Val2 := Val1;

    if (Typ = 'INTEGER') or (Typ = 'SMALLINT') then
      St := St + IntToStr(ValeurI(Val1)) + ' AND ' + IntToStr(ValeurI(Val2))

    else if (Typ = 'DOUBLE') or (Typ = 'RATE') or (Typ = 'EXTENDED') then
      St := St + StrfPoint(Valeur(Val1)) + ' AND ' + StrfPoint(Valeur(Val2))

    else if Typ = 'DATE' then
      St := St + '"' + UsDateTime(TextToDate(Val1)) + '" AND "' + UsDateTime(TextToDate(Val2)) + '"'

    else
      St := St + '"' + Val1 + '" AND "' + Val2 + '"';

    if Operateur = 'G' then
      St := ' NOT ' + St;
  end

  {In ou Not In}
  else if (Operateur = 'I') or (Operateur = 'J') then begin
    Typ := ChampToType(Champ);
    Value := FindetReplace_(Value, ' ou ', ';', True);
    Value := FindetReplace_(Value, ' OU ', ';', True);
    Value := FindetReplace_(Value, ' Ou ', ';', True);

    Val2 := '';

    while Value <> '' do begin
      Val1 := ReadTokenSt(Value);

      if Val1 <> '' then begin
        if (Typ = 'INTEGER') or (Typ = 'SMALLINT') then
          St := St + IntToStr(ValeurI(Val1)) + ', '

        else if (Typ = 'DOUBLE') or (Typ = 'RATE') or (Typ = 'EXTENDED') then
          St := St + StrfPoint(Valeur(Val1)) + ' ,'

        else if Typ = 'DATE' then
          St := St + '"' + UsDateTime(TextToDate(Val1)) + '", '
        else
          St := St + '"' + Val1 + '", ';

        Val2 := 'Ok';
      end;
    end;

    if Val2 <> '' then begin
      System.Delete(St, Length(St) - 1, 2);
      St := ' IN (' + St + ')';

      if Operateur = 'J' then
        St := ' NOT ' + St;
    end
  end

  else begin
    Typ := ChampToType(Champ);
    {Si de type inconnu soit ?? alors controler que c'est une formule...}
    if Typ = '??' then
      Typ := FormuleToType(Champ);

    OkChamp := (Copy(Value, 1, 1) = '[') and (Copy(Value, Length(Value), 1) = ']');

    if Okchamp then
      St := St + Value

    else if (Typ = 'INTEGER') or (Typ = 'SMALLINT') then
      St := St + IntToStr(ValeurI(Value))

    else if (Typ = 'DOUBLE') or (Typ = 'RATE') or (Typ = 'EXTENDED') then
      St := St + StrfPoint(Valeur(Value))

    else if Typ = 'DATE' then
      St := St + '"' + UsDateTime(TextToDate(Value)) + '"'

    else begin
      St := St + '"';

      if ((Operateur = 'L') or (Operateur = 'M')) then
        St := St + '%';

      St := St + CheckdblQuote(Value);

      if ((Operateur = 'C') or (Operateur = 'L') or (Operateur = 'D') or (Operateur = 'M')) then
        St := St + '%';

      St := St + '"';
    end;
  end;

  Result := Champ + St;
end;

{Purge de la table PROSPECTCONF des éléments issus de la compta et de la Tréso
{---------------------------------------------------------------------------------------}
function  VideConfidentialites : Boolean;
{---------------------------------------------------------------------------------------}
begin
  Result := True;
  if PGIAsk(TraduireMemoire('Vous êtes sur le point de supprimer le paramétrage des confidentialités.') + #13 +
            TraduireMemoire('Souhaitez-vous abandonner ?')) = mrYes then begin
    Result := False;
    Exit;
  end;
  try
    ExecuteSQL('DELETE FROM PROSPECTCONF WHERE RTC_PRODUITPGI IN ("' + typ_Compta + '", "' +
               typ_ComptaTreso + '", "' + typ_Treso + '")');
  except
    on E : Exception do begin
      PGIError(TraduireMemoire('Une erreur est intervenue lors de la réinitialisation des confidentialités avec le message : ')
               + #13#13 + E.Message);
      Result := False;
    end;
  end;
end;

{---------------------------------------------------------------------------------------}
{----------------------------     TOBJCONFIDENTIALITE     ------------------------------}
{---------------------------------------------------------------------------------------}

{---------------------------------------------------------------------------------------}
constructor TObjConfidentialite.Create(CodeUser : string; Produit : TContexte_PGI = ctxTreso; TypeConf : string = '');
{---------------------------------------------------------------------------------------}
begin
  FListeConf := TOB.Create('AZERTY', nil, -1);
  {Prépare la clause where sur les codes produit}
  FWhereProduit := GetProduit(Produit);
  {Initialise FIntervenant qui peut être un User ou UserGroupe à partir de l'utilisateur}
  FIntervenant := GetIntervenant(CodeUser);
  {Chargement des confidentialités de l'utilisateur}
  ChargeConfident(TypeConf);
end;

{---------------------------------------------------------------------------------------}
destructor TObjConfidentialite.Destroy;
{---------------------------------------------------------------------------------------}
begin
  if Assigned(FListeConf) then FreeAndNil(FListeConf);
  inherited;
end;

{---------------------------------------------------------------------------------------}
procedure TObjConfidentialite.ChargeConfident(TypeConf : string = '');
{---------------------------------------------------------------------------------------}
var
  S : string;
begin
  s := 'SELECT RTC_TYPECONF, RTC_SQLCONF, RTC_CHAMP1, RTC_CHAMP2, RTC_CHAMP3 FROM PROSPECTCONF WHERE RTC_INTERVENANT = "' + FIntervenant +
       '" AND ' + FWhereProduit;
  if TypeConf <> '' then begin
    if TypeConf[Length(TypeConf)] <> ';' then TypeConf := TypeConf + ';';
    S := S + ' AND RTC_TYPECONF IN (' + GetClauseIn(TypeConf) + ')';
  end;
  FListeConf.LoadDetailFromSQL(S);
end;

{---------------------------------------------------------------------------------------}
function TObjConfidentialite.GetLeftJoinFromTab(UnTypeConf : string) : string;
{---------------------------------------------------------------------------------------}
var
  n : Integer;
begin
  Result := '';
  if Length(UnTypeConf) < 3 then Exit;
  if Length(UnTypeConf) > 3 then UnTypeConf := Copy(UnTypeConf, 1, 3);
  for n := 0 to MaxTypeConfidentialite do begin
    if UnTypeConf = TableauLeftJoin[0, n] then begin
      Result := TableauLeftJoin[1, n];
      Break;
    end;
  end;
end;

{Remarque : par défaut, tout est ouvert ; ce n'est que s'il y a une confidentialité que l'on filtre
{---------------------------------------------------------------------------------------}
function TObjConfidentialite.GetWhereSQLConf(TypeConf : string) : HString;
{---------------------------------------------------------------------------------------}
var
  F : TOB;
  C : string;
begin
  Result := '';
  if TypeConf = '' then Exit;

  if TypeConf[Length(TypeConf)] <> ';' then TypeConf := TypeConf + ';';
  C := ReadTokenSt(TypeConf);
  while C <> '' do begin
    F := FListeConf.FindFirst(['RTC_TYPECONF'], [C], True);
    if Assigned(F) then begin
      if Result = '' then Result := '(' + F.GetString('RTC_SQLCONF') + ')'
                     else Result := Result + ' AND (' + F.GetString('RTC_SQLCONF') + ')';
    end;
    C := ReadTokenSt(TypeConf);
  end;
end;

{---------------------------------------------------------------------------------------}
function TObjConfidentialite.GetLeftJoinConf(TypeConf : string) : string;
{---------------------------------------------------------------------------------------}
var
  C : string;
begin
  Result := '';
  if TypeConf = '' then Exit;

  if TypeConf[Length(TypeConf)] <> ';' then TypeConf := TypeConf + ';';
  C := ReadTokenSt(TypeConf);
  while C <> '' do begin
    Result := Result + GetLeftJoinFromTab(C);
    C := ReadTokenSt(TypeConf);
    if (c <> '') and (Result <> '') then Result := Result + ' AND ';
  end;
end;

{---------------------------------------------------------------------------------------}
function TObjConfidentialite.GetChampsConf(TypeConf: string) : string;
{---------------------------------------------------------------------------------------}
var
  C : string;
  F : TOB;
  n : Integer;
begin
  Result := '';
  if TypeConf = '' then Exit;

  C := ReadTokenSt(TypeConf);
  while C <> '' do begin
    F := FListeConf.FindFirst(['RTC_TYPECONF'], [C], True);
    if Assigned(F) then begin
      for n := 1 to 3 do
        if F.GetString('RTC_CHAMP' + IntToStr(n)) <> '' then
          Result := Result + F.GetString('RTC_CHAMP' + IntToStr(n)) + ';';
    end;
    C := ReadTokenSt(TypeConf);
  end;
end;

{---------------------------------------------------------------------------------------}
class function TObjConfidentialite.GetProduit(Produit : TContexte_PGI) : string;
{---------------------------------------------------------------------------------------}
begin
  if Produit = ctxTreso then
    Result := 'RTC_PRODUITPGI IN ("' + typ_ComptaTreso + '", "' + typ_Treso + '")'
  else begin
    if EstComptaTreso then Result := 'RTC_PRODUITPGI IN ("' + typ_Compta + '", "' + typ_ComptaTreso + '")'
                      else Result := 'RTC_PRODUITPGI  = "' + typ_Compta + '"'
  end;
end;

{---------------------------------------------------------------------------------------}
class function TObjConfidentialite.GetIntervenant(Utilisat : string) : string;
{---------------------------------------------------------------------------------------}
var
  Q : TQuery;
begin
  {Je ne sais pas si c'est utile, peut devrais-je me contenter de V_PGI.Groupe}
  if GetParamSocSecur(NOMCONFIDENTIEL, False) then begin
    Q := OpenSQL('SELECT US_GROUPE FROM UTILISAT WHERE US_UTILISATEUR = "' + Utilisat + '"', True);
    if not Q.EOF then Result := Q.Fields[0].AsString
                 else Result := V_PGI.Groupe;
    Ferme(Q);
  end
  else
    Result := Utilisat;
end;

{---------------------------------------------------------------------------------------}
class function TObjConfidentialite.GetWhereConf(CodeUser, TypeConf : string; Produit : TContexte_PGI = ctxTreso): string;
{---------------------------------------------------------------------------------------}
var
  Q : TQuery;
  S : string;
begin
  Result := '';
  
  if (CodeUser = '') or (TypeConf = '') then Exit;

  s := 'SELECT RTC_SQLCONF FROM PROSPECTCONF WHERE RTC_INTERVENANT = "' + GetIntervenant(CodeUser) +
       '" AND ' + GetProduit(Produit) + ' AND RTC_TYPECONF = "' + TypeConf + '"';

  Q := OpenSQL(S, True);
  if not Q.EOF then Result := Q.FindField('RTC_SQLCONF').AsString;
  Ferme(Q);
end;


{---------------------------------------------------------------------------------------}
{----------------------------            TOFCONF          ------------------------------}
{---------------------------------------------------------------------------------------}

{---------------------------------------------------------------------------------------}
procedure TOFCONF.OnArgument(S : string);
{---------------------------------------------------------------------------------------}
begin
  inherited;
  InitConfident;
end;

{---------------------------------------------------------------------------------------}
procedure TOFCONF.OnClose;
{---------------------------------------------------------------------------------------}
begin
  if Assigned(FObjTypeConf) then FreeAndNil(FObjTypeConf);
  if Assigned(XX_WHERECONF) then FreeAndNil(XX_WHERECONF);
  if Assigned(XX_FROM) then FreeAndNil(XX_FROM);
  if Assigned(XX_WHEREJOIN) then FreeAndNil(XX_WHEREJOIN);
  if Assigned(FObjConfidentialite) then FreeAndNil(FObjConfidentialite);
  inherited;
end;

(*
{---------------------------------------------------------------------------------------}
function TOFCONF.GetNomTables : string;
{---------------------------------------------------------------------------------------}
var
  e, i, j : WideString;
  a, b, c, d, f, g, h, k, aTable : string;
  m, n : Boolean;
  lTable : string;
begin
  Result := '';

  {1/ AUTOMATISATION POUR UN MUL}
  if Ecran is TFMul then begin
    a := TFMul(Ecran).DBListe;
    {Récupération des tables de la DBListe}
    ChargeHListe(a, aTable, b, c, d, e, f, g, h, i, j, k, m, n);
    if aTable = '' then Exit;

    {Gestion du cas des vues : on regarde si les champs de confidentialités sont dans la requête}
    c := EstUneVue(aTable);
    m := False;
    if c <> '' then begin
      d := ObjConfidentialite.GetChampsConf(ObjTypeConf.GetListeTypeConf);
      e := ReadTokenSt(d);
      while e <> '' do begin
        {On regarde si le champ en cours figure dans la requête de la vue}
        if Pos(e, c) = 0 then begin
          m := True;
          Break;
        end;
        e := ReadTokenSt(d);
      end;
      {Tous les champs figure dans la vue}
      if not m then Exit;
    end;

    b := aTable;

    {Récupération des tables des confidentialités}
    lTable := GetTableFromConf(ObjTypeConf.GetListeTypeConf);
    {On s'assure que les tables confidentialités ne sont pas redondantes avec le paramétrage}
    a := Trim(ReadTokenPipe(lTable, ','));
    while a <> '' do begin
      {En tréso, on ne gère pas BanqueCP car elle est déjà gérée dans tous les écrans}
      if (CtxTreso in V_PGI.PGIContexte) xor (UpperCase(a) = 'BANQUECP') then
        if Pos(a, aTable) = 0 then b := b + ',' + a;

      a := Trim(ReadTokenPipe(lTable, ','));
    end;
    if b <> aTable then Result := b;
  end

  {2/ AUTOMATISATION POUR UN CUBDE}
  else if (Ecran is TFCube) or (Ecran is TFStat) then begin
    (* A REVOIR !!!!!
    {Récupération de la requête}
    if Ecran is TFCube then aTable := UpperCase(TFCube(Ecran).FromSQL)
                       else aTable := UpperCase(TFStat(Ecran).FSQL.Text);
    a := '';
    b := '';
    d := '';
    e := '';

    {Récupération des tables du cube}
    for p := Low(V_PGI.DETables) to High(V_PGI.DETables) do begin
      if (Pos(' ' + V_PGI.DETables[p].Nom, aTable) > 0) or
         (Pos(',' + V_PGI.DETables[p].Nom, aTable) > 0) then
        a := a + V_PGI.DETables[p].Nom + ';';
    end;
    {On regarde si les confidentialite sont présentes dans la requête}
    g := GetTableFromConf(ObjTypeConf.GetListeTypeConf);
    c := ReadTokenPipe(g, ',');
    while c <> '' do begin
      if Pos(c + ';', a) = 0 then d := d + c + ';';
      c := ReadTokenPipe(g, ',');
    end;

    for p := Low(V_PGI.DEVues) to High(V_PGI.DEVues) do begin
      if (V_PGI.DETables[p].Nom <> '') and
         ((Pos(' ' + V_PGI.DEVues[p].Nom, aTable) > 0) or
          (Pos(',' + V_PGI.DEVues[p].Nom, aTable) > 0)) then
        b := V_PGI.DEVues[p].SQL + ' ';
    end;
    if Trim(b) <> '' then begin
      a := '';
      {Récupération des tables des vues}
      for p := Low(V_PGI.DETables) to High(V_PGI.DETables) do begin
        if (V_PGI.DETables[p].Nom <> '') and
           ((Pos(' ' + V_PGI.DETables[p].Nom, b) > 0) or
            (Pos(',' + V_PGI.DETables[p].Nom, b) > 0)) then
          a := a + V_PGI.DETables[p].Nom + ';';
      end;
      {On regarde si les confidentialite sont présentes dans la requête}
      g := GetTableFromConf(ObjTypeConf.GetListeTypeConf);
      c := ReadTokenPipe(g, ',');
      while c <> '' do begin
        if Pos(c + ';', a) = 0 then d := d + c + ';';
        c := ReadTokenPipe(g, ',');
      end;
    end;

    Result := d;
    
  end;
end;
*)
{---------------------------------------------------------------------------------------}
procedure TOFCONF.OnLoad;
{---------------------------------------------------------------------------------------}
begin
  inherited;

  if Assigned(ObjTypeConf) then begin
    {Cela me semble difficilement gérable, sauf peut-être pour les Muls et encore ...
    lTable := GetNomTables;
    if lTable <> '' then begin
      XX_FROM.Text := lTable;
      XX_WHEREJOIN.Text := FObjConfidentialite.GetLeftJoinConf(ObjTypeConf.GetListeTypeConf);
    end;
    }
    XX_WHERECONF.Text := FObjConfidentialite.GetWhereSQLConf(ObjTypeConf.GetListeTypeConf);
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOFCONF.OnUpdate;
{---------------------------------------------------------------------------------------}
begin
  inherited;

end;

{---------------------------------------------------------------------------------------}
procedure TOFCONF.InitConfident;
{---------------------------------------------------------------------------------------}
var
  FPages : TPageControl;
begin
  {Création du XX_WHERE}
  FPages := TPageControl(Ecran.FindComponent('PAGES'));
  if Assigned(FPages) then begin
    XX_WHERECONF := THEdit.Create(Ecran);
    XX_WHERECONF.Parent := FPages.Pages[0];
    XX_WHERECONF.Visible := False;
    XX_WHERECONF.Name := 'XX_WHERECONF';
    XX_WHERECONF.Text := '';

    XX_WHEREJOIN := THEdit.Create(Ecran);
    XX_WHEREJOIN.Parent := FPages.Pages[0];
    XX_WHEREJOIN.Visible := False;
    XX_WHEREJOIN.Name := 'XX_WHEREJOIN';
    XX_WHEREJOIN.Text := '';

    XX_FROM := THEdit.Create(Ecran);
    XX_FROM.Parent := FPages.Pages[0];
    XX_FROM.Visible := False;
    XX_FROM.Name := 'XX_FROM';
    XX_FROM.Text := '';
  end
  else begin
    FGereConfidentialite := False;
    if V_PGI.SAV then PGIBox(TraduireMemoire('PageControl introuvable : pas de confidentialité possible'));
    Exit;
  end;

  {Gestion de l'objet confidentialité}
  if FGereConfidentialite then
    FObjConfidentialite := TObjConfidentialite.Create(V_PGI.User, GetExeContext, ObjTypeConf.GetListeTypeConf);
end;

{---------------------------------------------------------------------------------------}
procedure TOFCONF.SetTypeConfidentialite(Value : string);
{---------------------------------------------------------------------------------------}
begin
  {La gestion de la confidentialité n'est effective que si on est sur un Mul, un QRS1, un Cube ou une Stat
   Et pour le moment, faute de pouvoir tout automatiser, il faut que value soit renseigné
   10/10/07 : dans TRQRJALPORTEFEUIL_TOF, je mets volontairement Value à vide pour ne pas gérer les confidentialités}
  FGereConfidentialite := IsEcranAConfidentialite and (Value <> '');

  if not FGereConfidentialite then begin
    if Assigned(ObjTypeConf) then
      ObjTypeConf.SetOffConfidentialite(tyc_All);
  end
  else begin
    if not Assigned(ObjTypeConf) then
      ObjTypeConf := TListeConf.Create(Value)
    else begin
      if Value = '' then
        ObjTypeConf.SetOnConfidentialite(tyc_All)
      else begin
        ObjTypeConf.SetOffConfidentialite(tyc_All);
        ObjTypeConf.SetOnConfidentialite(Value);
      end;
    end;
  end;
end;

{---------------------------------------------------------------------------------------}
function TOFCONF.IsEcranAConfidentialite : Boolean;
{---------------------------------------------------------------------------------------}
var
  sClass : string;
begin
  sClass := UpperCase(Ecran.ClassName);
  Result := (sClass = 'TFMUL') or (sClass = 'TFQRS1') or (sClass = 'TFSTAT') or (sClass = 'TFCUBE');
end;


{---------------------------------------------------------------------------------------}
{----------------------------          TLISTECONF         ------------------------------}
{---------------------------------------------------------------------------------------}

{---------------------------------------------------------------------------------------}
constructor TListeConf.Create(ListTypeConf : string = '');
{---------------------------------------------------------------------------------------}
begin
  FListeConfidentialite := TStringList.Create;
  FListeConfidentialite.Sorted := True;
  FListeConfidentialite.Duplicates := dupIgnore;

  {Initialisation de toutes les confidentialités à False}
  InitConfidentialite;

  {Si l'on a passé des confidentialités au constructeur, c'est que l'on veut les activer par défaut}
  if ListTypeConf <> '' then SetOnConfidentialite(ListTypeConf);
end;

{---------------------------------------------------------------------------------------}
destructor TListeConf.Destroy;
{---------------------------------------------------------------------------------------}
begin
  if Assigned(FListeConfidentialite) then FreeAndNil(FListeConfidentialite);
  inherited;
end;

{---------------------------------------------------------------------------------------}
procedure TListeConf.MajConfidentialite(TypeConf: string; Actif: Boolean);
{---------------------------------------------------------------------------------------}
var
  OffOk : Integer;
  OnOk  : Integer;
begin
  OffOk := FListeConfidentialite.IndexOf(TypeConf + ';-;');
  OnOk  := FListeConfidentialite.IndexOf(TypeConf + ';X;');

  {La confidentialité est stockée dans le sens passé en paramètre}
  if (Actif and (OnOk > -1)) or (not Actif and (OffOk > -1)) then Exit;

  {La confidentialité est stockée mais dans le sens opposé à celui passé en paramètre}
  if (not Actif and (OnOk > -1)) or (Actif and (OffOk > -1)) then begin
    {On commence par détruire la référence}
    if Actif then FListeConfidentialite.Delete(OffOk)
             else FListeConfidentialite.Delete(OnOk);
  end;

  {La confidentialité est n'est pas stockée, on l'ajoute}
  if Actif then FListeConfidentialite.Add(TypeConf + ';X;')
           else FListeConfidentialite.Add(TypeConf + ';-;');

end;

{Contrôle si la confidentialité passée en paramètre est active.
{---------------------------------------------------------------------------------------}
function TListeConf.GereConfidentialite(TypeConf : string): Boolean;
{---------------------------------------------------------------------------------------}
begin
  Result := FListeConfidentialite.IndexOf(TypeConf + ';X;') > -1;
end;

{Désactive les confidentialités passées en paramètre et séparées par des ";".
 Pour les valeurs, il est préférable d'utiliser les constantes tyc_
{---------------------------------------------------------------------------------------}
procedure TListeConf.SetOffConfidentialite(ListeTypeConf : string);
{---------------------------------------------------------------------------------------}
var
  ch : string;
begin
  if ListeTypeConf = '' then Exit;

  if ListeTypeConf = tyc_All then
    InitConfidentialite
  else begin
    if ListeTypeConf[Length(ListeTypeConf)] <> ';' then ListeTypeConf := ListeTypeConf + ';';
    ch := ReadTokenSt(ListeTypeConf);
    while ch <> '' do begin
      MajConfidentialite(ch, False);
      ch := ReadTokenSt(ListeTypeConf);
    end;
  end;
end;

{Active les confidentialités passées en paramètre et séparées par des ";".
 Pour les valeurs, il est préférable d'utiliser les constantes tyc_
{---------------------------------------------------------------------------------------}
procedure TListeConf.SetOnConfidentialite(ListeTypeConf : string);
{---------------------------------------------------------------------------------------}
var
  ch : string;
begin
  if ListeTypeConf = '' then Exit;

  if ListeTypeConf = tyc_All then
    InitConfidentialite(True)
  else begin
    if ListeTypeConf[Length(ListeTypeConf)] <> ';' then ListeTypeConf := ListeTypeConf + ';';
    ch := ReadTokenSt(ListeTypeConf);
    while ch <> '' do begin
      MajConfidentialite(ch, True);
      ch := ReadTokenSt(ListeTypeConf);
    end;
  end;
end;

{Retourne la liste des types de confidentialité actifs
{---------------------------------------------------------------------------------------}
function TListeConf.GetListeTypeConf : string;
{---------------------------------------------------------------------------------------}
var
  n     : Integer;
  Ch    : string;
  Conf  : string;
//  AllOk : Boolean;
begin
  Result := '';
//  AllOk := True;
  for n := 0 to FListeConfidentialite.Count - 1 do begin
    Ch := FListeConfidentialite[n];
    Conf := ReadTokenSt(Ch);
    Ch := ReadTokenSt(Ch);
    if Ch = 'X' then Result := Result + Conf + ';'
                //else AllOk := False;
  end;
//  if AllOk then Result := '';
end;

{---------------------------------------------------------------------------------------}
procedure TListeConf.InitConfidentialite(Actif : Boolean = False);
{---------------------------------------------------------------------------------------}
begin
  MajConfidentialite(tyc_Banque, Actif);
  MajConfidentialite(tyc_Flux  , Actif);
  {A compléter ...}
end;

end.
