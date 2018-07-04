{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 09/10/2007
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : ANARCHIREGLPAIE ()
Mots clefs ... : TOF;ANARCHIREGLPAIE
*****************************************************************}
{
PT1 24/10/2007 GGU V_8 Passage en TreeTobFrame
PT2 23/11/2007 GGU V_8 Gestion multi-dossier
PT3 26/11/2007 GGU V_8 Gestion des éléments nationaux
PT4 26/11/2007 GGU V_8 Gestion des tables dynamiques
PT5 26/11/2007 GGU V_8 Optimisations
PT6 28/11/2007 GGU V_8 Tri des éléments
PT7 13/12/2007 GGU V_8 En cwas, si la RechDom ne trouve pas l'élément, il renvoi
                       'Error' alors qu'en 2/3 RechDom renvoi ''. -> Uniformisation
PT8 : 19/12/2007 GGU V_80 Compatibilité Delphi 5 / Version 7 pour reprise de l'outils en V7
}
unit UTofPG_AnalyseArchiReglPaie;

Interface

Uses StdCtrls, 
     Controls, 
     Classes, 
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul,
{$else}
     eMul,
{$ENDIF}
     forms,
     sysutils,
     ComCtrls,
     HCtrls, 
     HEnt1, 
     HMsgBox, 
     UTOF, UTob
     , PGTreeTobFrame, hpanel
{$IFNDEF VER130} //PT8
     , Utofpg_Diagnostic
{$ENDIF}
     , HTB97   //PT1
     ;

Type

  { Objet qui permet de gérer les informations nécéssaires à la recherche des
    éléments utilisés pour un type d'élément. (Où chercher les éléments utilisés
    et où chercher leurs types respectifs) -> Fait le liens entre le
    TRubriqueDispatcher à une nature particulière d'éléments (par exemple
    variable, rémunération, cotisation, cumul, profil...)
   }
  TRubriqueDispatcherConnecteur = Class
  Public
    TobMereRubriques : Tob;
    TypeRubriques : String;
    ListeChampsTypeElt, ListeChampsCodeElt : Array of String;
    ChampsCode : String;
    ChampsType : String;//PT2
    TabletteType, TabletteLibelle : String;
    Prefixe : String; //PT2
    Constructor Create(stSQL, NomChampsCode : String; TypeDesRubriques : String; ListeDesChampsTypeElt, ListeDesChampsCodeElt : Array of String);
    Destructor Destroy(); override ;
    Procedure GetPredefiniNodossierVisibleElt(var TypeElt, Predef, Nodoss : String; CodeElt, NoDossUtilisateur : String);//PT2
    Function GetVisibleElt(CodeElt, NoDossUtilisateur : String) : Tob;  //PT2
  end;

  { Type des procédures pouvant être "dispatchées" sur tous les types d'éléments }
  TProcedureDispatchRub = procedure(Connecteur: TRubriqueDispatcherConnecteur) of object;

  { Type des gérant tous les types d'éléments }
//PT2   TProcedureDispatchRubElt = procedure(TobElt: Tob; TypeElt, ChampType, ChampCode : String) of object;
  TProcedureDispatchRubElt = procedure(TobElt: Tob; TypeElt, ChampType, ChampCode, ChampNoDossier : String) of object;

  { Objet qui permet d'executer une procédure sur l'ensemble des rubriques,
    quelque soit le type de la rubrique. -> Gère l'abstraction du type d'élément
   }
  TRubriqueDispatcher = Class
  Private
    Connecteurs : Array of TRubriqueDispatcherConnecteur;
  Public
    Function GetConnecteur(TypeDeRubriques : String) : TRubriqueDispatcherConnecteur;
    Procedure DispatchRubriques(ProcedureToDispatch : TProcedureDispatchRub);
    Procedure DispatchRubriquesElt(ProcedureToDispatch : TProcedureDispatchRubElt);
    Constructor Create();
    Destructor Destroy(); override ;
  end;

  TOF_ANARCHIREGLPAIE = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  private
    TobAnalyse : Tob;
    RubriqueDispatcher : TRubriqueDispatcher;
    TreeTobFrame : TFFrameTreeTob;   //PT1
//PT2    procedure LectureRub(TobElt: Tob; TypeElt, ChampType, ChampCode : String);
    Procedure LectureRub(TobElt : Tob; TypeElt, ChampType, ChampCode, ChampNoDossier : String); //PT2
    Procedure LectureType(Connecteur: TRubriqueDispatcherConnecteur);
    procedure RechercheRubriquesLiees(TobMere, TobElt: Tob;
      NatureMere, ChampType, ChampCode, TabletteType: String);
    Procedure AnalyseElt(TobElt : Tob; Nature, Code, Predefini, Dossier : String);
    Function AddEltOnTobAnalyse(Nature, Code, Predefini, Dossier : String) : Tob;
    Procedure AnalyseInversee;
  public
    Procedure OnBTANALYSESELECTELTClick(Sender : TObject);
  end ;

  { Fonction d'uniformisation des codes de nature (VAR pour les variables de paie,
    ELT pour les éléments dynamiques... }
  Function RechTypeToCodeType(TypeElt, RechType : String) : String;

Implementation

Uses
  UTomVariablePaie, ed_tools//, ConvUtils
  , PGTobOutils
{$IFNDEF VER130} //PT8
  , StrUtils
{$ENDIF}
  ;

{$IFDEF VER130} //PT8
function LeftStr(const AText: AnsiString; const ACount: Integer): AnsiString;
begin
  result := Copy(AText, 0, ACount);
end;

Function IconeToNature(stIcone : String) : String;
begin
  if stIcone = '#ICO#58' then   { Variables }
    result := 'VAR'
  else if stIcone = '#ICO#67' then { Elements nationaux  }
    result := 'ELT'
  else if stIcone = '#ICO#77' then  { Tables dynamiques }
    result := 'TDY'
  else if stIcone = '#ICO#75' then  { Eléments dynamiques }
    result := 'EDY'
  else if stIcone = '#ICO#53' then  { Cotisation }
    result := 'COT'
  else if stIcone = '#ICO#52' then  { Remunération  }
    result := 'REM'
  else if stIcone = '#ICO#69' then  { Remunération }
    result := 'CUM'
  else if stIcone = '#ICO#76' then  { Profil }
    result := 'PRO'
  else if stIcone = '#ICO#21' then  { Valeurs }
    result := 'VAL'
  else if stIcone = '#ICO#41' then  { Nature vide (Lignes de mises en forme) }
    result := ''
  else   { Autres }
    result := '';

end;

Function NatureToIcone(stNature : String) : String;
begin
  if pos('#ICO', stNature) > 0 then
  begin
    result := stNature;
    exit;
  end;
  if stNature = 'VAR' then   { Variables }
    result := '#ICO#58'
  else if stNature = 'ELT' then { Elements nationaux  }
    result := '#ICO#67'
  else if stNature = 'TDY' then  { Tables dynamiques }
    result := '#ICO#77'
  else if stNature = 'EDY' then  { Eléments dynamiques }
    result := '#ICO#75'
  else if stNature = 'COT' then  { Cotisation }
    result := '#ICO#53'
  else if stNature = 'REM' then  { Remunération  }
    result := '#ICO#52'
  else if stNature = 'CUM' then  { Remunération }
    result := '#ICO#69'
  else if stNature = 'PRO' then  { Profil }
    result := '#ICO#76'
  else if stNature = 'VAL' then  { Valeurs }
    result := '#ICO#21'
  else if stNature = '' then     { Nature vide (Lignes de mises en forme) }
    result := '#ICO#41'
  else   { Autres }
    result := '#ICO#25';
end;

Procedure TobNatureToIcone(var LaTob : Tob);
var
  IndexTob : Integer;
  TempTob : Tob;
begin
  if not Assigned(LaTob) then exit;
  LaTob.PutValue('ARP_NATUREARCHI', NatureToIcone(LaTob.GetString('ARP_NATUREARCHI')));
  for IndexTob := 0 to LaTob.detail.count -1 do
  begin
    TempTob := LaTob.detail[IndexTob];
    TobNatureToIcone(TempTob);
  end;
end;
{$ENDIF}

Procedure AddEmptyFields(TheTob : Tob);
begin
  TheTob.AddChampSupValeur('ARP_TYPENATUREARCHI', '');
  TheTob.AddChampSupValeur('ARP_PREDEFINI', '');
  TheTob.AddChampSupValeur('ARP_NODOSSIER', '');
  TheTob.AddChampSupValeur('ARP_ORIGINE', '');
end;

{***********A.G.L.***********************************************
Auteur  ...... : GGU
Créé le ...... : 12/11/2007
Modifié le ... :   /  /
Description .. : Fonction d'uniformisation des codes de nature (VAR pour
Suite ........ : les variables de paie, ELT pour les éléments dynamiques...
Mots clefs ... :
*****************************************************************}
Function RechTypeToCodeType(TypeElt, RechType : String) : String;
  function RechVarTypeToCodeType(RechType : String) : String;
  begin
    if RechType = '03' then
      result := 'VAR'
    else if RechType = '02' then
      result := 'ELT'
    else if RechType = '04' then
      result := 'VAL'
    else if (RechType >= '05') and (RechType <= '10') then
      result := 'REM'
    else if (RechType >= '12') and (RechType <= '14') then
     result := 'COT'
    else if (RechType >= '225') then   //PT4
     result := 'TDY'
    else
      result := '';
  end;
  function RechCotTypeToCodeType(RechType : String) : String;
  begin
    if RechType = 'VAR' then
      result := 'VAR'
    else if RechType = 'ELN' then
      result := 'ELT'
    else
      result := '';
  end;
  function RechProTypeToCodeType(RechType : String) : String;
  begin
    if RechType = 'BAS' then
      result := 'COT'
    else     if RechType = 'COT' then
      result := 'COT'
    else     if RechType = 'AAA' then
      result := 'REM'
    else
      result := '';
  end;
  function RechRemTypeToCodeType(RechType : String) : String;
  begin
    if RechType = '03' then
      result := 'VAR'
    else if RechType = '02' then
      result := 'ELT'
    else if (RechType >= '04') and (RechType <= '07') then
      result := 'REM'
    else if RechType = '09' then
      result := 'CUM'
    else
      result := '';
  end;
begin
  if TypeElt = 'VAR' then
    result := RechVarTypeToCodeType(RechType)
  else if TypeElt = 'COT' then
    result := RechCotTypeToCodeType(RechType)
  else if TypeElt = 'REM' then
    result := RechRemTypeToCodeType(RechType)
  else if TypeElt = 'CUM' then
    result := RechProTypeToCodeType(RechType)
  else if TypeElt = 'PRO' then
    result := RechProTypeToCodeType(RechType)
  else
    result := '';
end;

procedure TOF_ANARCHIREGLPAIE.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_ANARCHIREGLPAIE.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_ANARCHIREGLPAIE.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_ANARCHIREGLPAIE.OnLoad ;
begin
  Inherited ;
end ;

Procedure TOF_ANARCHIREGLPAIE.RechercheRubriquesLiees(TobMere, TobElt : Tob; NatureMere, ChampType, ChampCode, TabletteType : String);
{$IFDEF VER130} //PT8
  //Fonction de UTomVariablePaie depuis la V8
  function RechercheNomTablette(TypeBase, NatureVariable: String; Presence : Boolean): string;
  begin
    if      TypeBase = '02' then
      result := 'PGELEMENTNAT'
    else if TypeBase = '30' then
      result := 'PGZONEELTDYN'
    else if TypeBase = '31' then
      result := 'PGZONELIBRESAL'
    else if (TypeBase = '12') or (TypeBase = '13') or (TypeBase = '14') or (TypeBase = '25') or (TypeBase = '26') then
      result := 'PGCOTISATION'
    else if (TypeBase = '16') or (TypeBase = '17') or (TypeBase = '18') or (TypeBase = '19') then
      result := 'PGBASECOTISATION'
    else if (TypeBase = '05') or (TypeBase = '06') or (TypeBase = '07') or (TypeBase = '08') or (TypeBase = '09') or (TypeBase = '10') then
      result := 'PGREMUNERATION'
    else if TypeBase = '03' then
      if not Presence then
        result := 'PGVARIABLE'
      else
        result := 'PGVARIABLECEGSALVAL'
    else if TypeBase = '15' then
      result := 'PGCHAMPSALARIE'
    else if TypeBase = '20' then
      result := 'PGTABINTAGE'
    else if TypeBase = '21' then
      result := 'PGTABINTANC'
    else if TypeBase = '22' then
      result := 'PGTABINTDIV'
    else if TypeBase = '220' then
      result := 'PGTABINTDIW'
    else if TypeBase = '225' then
      result := 'PGTABLEDIMDSAL'
    else if TypeBase = '500' then
      result := 'PGCOMPTEURPRESJ'
    else if TypeBase = '501' then
      result := 'PGVARIABLEPREJ'
    else if TypeBase = '502' then
      result := 'PGCOMPTEURPRESH'
    else if TypeBase = '503' then
      result := 'PGVARIABLEPREH'
    else if TypeBase = '504' then
      result := 'PGCOMPTEURPRESM'
    else if TypeBase = '505' then
      result := 'PGVARIABLEPREM'
    else if (TypeBase = '506') or (TypeBase = '507') or (TypeBase = '508') then
      result := 'PGDROITJOURNEETYPE'
    else if (TypeBase = '510') or (TypeBase = '512') or (TypeBase = '513') or (TypeBase = '514') or
            (TypeBase = '515') or (TypeBase = '516') or (TypeBase = '517') or (TypeBase = '518') then
      result := 'PGMOTIFEVENEMENT'
    else if TypeBase = '520' then
      result := 'PGCOMPTEURPRESC'
    else if TypeBase = '521' then
      result := 'PGVARIABLEPREC'
    else if TypeBase = '530' then
      result := 'PGVARIABLEPRECEG'
    else
      if NatureVariable = 'CUM' then
        result := 'PGCUMULPAIE'
      else
        result := '';
  end;
{$ENDIF}
var
  NewTempRechTob : Tob;
  RechType, RechTypeLib, RechCode, RechCodeLib : String;
  TabletteCode, TypeMere, CodeType : String;
  Predef, NoDoss, TypeElt : String; //PT2
  RubDispatcConnecteur : TRubriqueDispatcherConnecteur;
  LibDossier : String; //PT7
  Function GetTabletteName(TypeDeMere, RechercheType : String) : String;
  begin
    if TypeDeMere = 'VAR' then
    begin
      Result := RechercheNomTablette(RechercheType, TobElt.GetString('PVA_NATUREVAR'), False)
    end else if TypeDeMere = 'REM' then
    begin
      if (RechercheType = '02') then
        Result := 'PGELEMENTNAT'
      else if (RechercheType = '03') then
        Result := 'PGVARIABLE'
      else if (RechercheType > '03') and (RechercheType < '08') then
        Result := 'PGREMUNERATION'
      else if (RechercheType = '09') then
        Result := 'PGCUMULPAIE'
      else
        Result := '';
    end else if TypeDeMere = 'COT' then
    begin
      if      (RechercheType = 'ELN') then
        Result := 'PGELEMENTNAT'
      else if (RechercheType = 'VAR') then
        Result := 'PGVARIABLE'
      else if (RechercheType = 'ELV') then
        Result := ''
      else if (RechercheType = 'VAL') then
        Result := ''
      else if (RechercheType = 'ZCU') then
        Result := 'PGCUMULPAIE'
      else if (RechercheType = 'BDC') then
        Result := 'PGBASECOTISATION'
      else if (RechercheType = 'EXB') then
        Result := 'PGEXONERATIONBAS'
      else if (RechercheType = 'EXC') then
        Result := 'PGEXONERATIONCOT'
      else
        Result := '';
    end else if (TypeDeMere = 'CUM') or (TypeDeMere = 'PRO') then
    begin
      if      (RechercheType = 'BAS') then
        Result := 'PGCOTISATION'
      else if (RechercheType = 'COT') then
        Result := 'PGCOTISATION'
      else if (RechercheType = 'AAA') then
        Result := 'PGREMUNERATION'
      else
        Result := '';
    end;
  end;
begin
  TypeMere := NatureMere;
  RechType := TobElt.GetString(ChampType);
  CodeType := RechTypeToCodeType(TypeMere, RechType);
  if RechType <> '' then
  begin
    RechTypeLib := RechDom(TabletteType, RechType, False);
    RechCode := TobElt.GetString(ChampCode);
    TabletteCode := GetTabletteName(TypeMere, RechType);
    RechCodeLib := RechDom(TabletteCode, RechCode, False);
    NewTempRechTob := Tob.Create('Rubrique liée', TobMere, -1);
    NewTempRechTob.AddChampSupValeur('ARP_NATUREARCHI', CodeType);
    NewTempRechTob.AddChampSupValeur('ARP_CODE', RechCode);
    //Debut PT2
    { On recherche le type, le prédéfini et le nodossier de la rubrique trouvée,
      cad ceux qui sont visibles pour l'élément analysé (m dossier ou commun) }
    RubDispatcConnecteur := RubriqueDispatcher.GetConnecteur(CodeType);
    TypeElt := '';
    Predef := '';
    NoDoss := '';
    if Assigned(RubDispatcConnecteur) then
      RubDispatcConnecteur.GetPredefiniNodossierVisibleElt(TypeElt, Predef, NoDoss, RechCode, TobElt.GetString(RubriqueDispatcher.GetConnecteur(TypeMere).Prefixe+'_NODOSSIER'));
    NewTempRechTob.AddChampSupValeur('ARP_TYPENATUREARCHI', TypeElt);
    NewTempRechTob.AddChampSupValeur('ARP_PREDEFINI', Predef);
    NewTempRechTob.AddChampSupValeur('ARP_NODOSSIER', NoDoss);
    if CodeType <> 'TDY' then
    begin
      LibDossier := RechDom('YYDOSSIER', NoDoss, False);  //PT7
      if LibDossier = 'Error' then LibDossier := '';  //PT7
      NewTempRechTob.AddChampSupValeur('ARP_ORIGINE', RechDom('YYPREDEFINI', Predef, False)+' '+LibDossier); //PT7
    end else //Pour les tables dynamiques, on ne peut pas savoir quel sera la table utilisée si on a pas le contexte
      NewTempRechTob.AddChampSupValeur('ARP_ORIGINE', '');
//    NewTempRechTob.AddChampSupValeur('ARP_ORIGINE', RechDom('YYPREDEFINI', Predef, False)+' '+RechDom('YYDOSSIER', NoDoss, False));
    //Fin PT2
    NewTempRechTob.AddChampSupValeur('ARP_MESSAGE', RechTypeLib+' : ('+RechCode+') '+RechCodeLib);
  end;
end;

//PT2 Procedure TOF_ANARCHIREGLPAIE.LectureRub(TobElt : Tob; TypeElt, ChampType, ChampCode : String);
Procedure TOF_ANARCHIREGLPAIE.LectureRub(TobElt : Tob; TypeElt, ChampType, ChampCode, ChampNoDossier : String); //PT2
var
  NewTempRechTob, TobMere, TobBranche : Tob;
  RechType, RechTypeLib, RechCode, RechCodeLib, CodeType, TabletteCode : String;
  NeedVerif : Boolean;
  RechPredefini, RechNoDossier, LType, LPredef, LNoDoss, AType, APredef, ANoDoss : String; //PT2
  Connecteur : TRubriqueDispatcherConnecteur;

  TempTob :Tob; //PT6
  PosInTob, IndexBranche : Integer; //PT6
  stTempNATUREARCHI, stTempNODOSSIER, stTempCODE : String; // PT6

  LibDossier : String; //PT7

  Procedure GetInfosElt(Elt : Tob; NatureElt : String; var Code, TypeNature, Predef, NoDoss, TabletteCode : String);
  var
    Connect : TRubriqueDispatcherConnecteur;
  begin
    Connect := RubriqueDispatcher.GetConnecteur(NatureElt);
    Code := TobElt.GetString(Connect.ChampsCode);
    if Connect.TabletteType <> '' then
      TypeNature := RechDom(Connect.TabletteType, TobElt.GetString(Connect.ChampsType), False)
    else
      TypeNature := TobElt.GetString(Connect.ChampsType);
    Predef := TobElt.GetString(Connect.Prefixe+'_PREDEFINI');
    NoDoss := TobElt.GetString(Connect.Prefixe+'_NODOSSIER');
    TabletteCode := Connect.TabletteLibelle;
  end;
begin
  RechType := TobElt.GetString(ChampType);
  RechPredefini := TobElt.GetString(ChampNoDossier);//PT2
  RechNoDossier := TobElt.GetString(ChampNoDossier);//PT2
  if RechType <> '' then
  begin
    RechCode := TobElt.GetString(ChampCode);
    RechNoDossier := TobElt.GetString(ChampNoDossier);  //PT2
    { On recherche si cet élément fait parti des éléments à analyser }
    CodeType := RechTypeToCodeType(TypeElt, RechType);
    TobMere := TobAnalyse.FindFirst(['ARP_NATURE'], [CodeType], False);
    //debut PT2
//    if Assigned(TobMere) then
//      TobMere := TobMere.FindFirst(['ARP_CODE'], [RechCode], False);
    if Assigned(TobMere) then
    begin
      if RechNoDossier <> '000000' then
      begin
        Connecteur := RubriqueDispatcher.GetConnecteur(CodeType);
        if Assigned(Connecteur) then
          Connecteur.GetPredefiniNodossierVisibleElt(AType, APredef, ANoDoss, RechCode, RechNoDossier);
      end else
        ANoDoss := RechNoDossier;
      TobMere := TobMere.FindFirst(['ARP_CODE', 'ARP_NODOSSIER'], [RechCode, ANoDoss], False);
    end;
    //fin PT2
    { Si oui, on l'écrit }
    if Assigned(TobMere) then
    begin
      TobBranche := TobMere.findFirst(['CODEBRANCHE'], [2], False);//On se met dans la branche "participe au calcul de "
      NeedVerif := False;
      if not Assigned(TobBranche) then
      begin
        TobBranche := Tob.Create( 'Participe à', TobMere, -1);
        TobBranche.AddChampSupValeur('ARP_NATUREARCHI', '');
        TobBranche.AddChampSupValeur('ARP_CODE', TobBranche.NomTable);
        TobBranche.AddChampSupValeur('ARP_MESSAGE', TobBranche.NomTable);
        TobBranche.AddChampSupValeur('CODEBRANCHE', 2);
        AddEmptyFields(TobBranche);
      end else begin
        NeedVerif := True;
      end;
      TobMere := TobBranche;
      RechType := TypeElt;
      GetInfosElt(TobElt, TypeElt, RechCode, LType, LPredef, LNoDoss, TabletteCode);
{      if TypeElt = 'VAR' then
      begin
        RechCode := TobElt.GetString('PVA_VARIABLE');
        LType := GetTypeNatureArchi(TobElt, TypeElt);
        LPredef := TobElt.GetString('PVA_PREDEFINI');
        LNoDoss := TobElt.GetString('PVA_NODOSSIER');
        TabletteCode := 'PGVARIABLE';
      end else if TypeElt = 'COT' then
      begin
        RechCode := TobElt.GetString('PCT_RUBRIQUE');
        LType := RechDom('PGTHEMECOT', TobElt.GetString('PCT_THEMECOT'), False);
        LPredef := TobElt.GetString('PCT_PREDEFINI');
        LNoDoss := TobElt.GetString('PCT_NODOSSIER');
        TabletteCode := 'PGCOTISATION';
      end else if TypeElt = 'REM' then
      begin
        RechCode := TobElt.GetString('PRM_RUBRIQUE');
        LType := TobElt.GetString('PRM_NATURERUB');
        LPredef := TobElt.GetString('PRM_PREDEFINI');
        LNoDoss := TobElt.GetString('PRM_NODOSSIER');
        TabletteCode := 'PGREMUNERATION';
      end else if TypeElt = 'CUM' then
      begin
        RechCode := TobElt.GetString('PCR_CUMULPAIE'); //PCL_CUMULPAIE ?
        LType := TobElt.GetString('PCR_NATURERUB');
        LPredef := TobElt.GetString('PCR_PREDEFINI');
        LNoDoss := TobElt.GetString('PCR_NODOSSIER');
        TabletteCode := 'PGCUMULPAIE';
      end else if TypeElt = 'PRO' then
      begin
        RechCode := TobElt.GetString('PPM_PROFIL');
        LType := TobElt.GetString('PPM_NATURERUB');
        LPredef := TobElt.GetString('PPM_PREDEFINI');
        LNoDoss := TobElt.GetString('PPM_NODOSSIER');
        TabletteCode := 'PGPROFIL';
      //Debut PT3
      end else if TypeElt = 'ELT' then
      begin
        RechCode := TobElt.GetString('PEL_CODEELT');
        LType := TobElt.GetString('PEL_THEMEELT');
        LPredef := TobElt.GetString('PEL_PREDEFINI');
        LNoDoss := TobElt.GetString('PEL_NODOSSIER');
        TabletteCode := 'PGELEMENTNAT';
      //Fin PT3
      end;  }


      //debut PT6 Tri des éléments
      //On recherche a quel index on doit mettre cet élément
      //Tri par nature / code
      PosInTob := -2;
      if NeedVerif then
      begin
        PosInTob := -1;
        for IndexBranche := 0 to TobBranche.Detail.Count -1 do
        begin
          TempTob := TobBranche.Detail[IndexBranche];
          stTempNATUREARCHI := TempTob.GetString('ARP_NATUREARCHI');
          { Le test est efectué pour les 2 valeurs du champs ARP_NATUREARCHI (transformé ou non en icone) }
          if   (stTempNATUREARCHI = RechType)
            or (stTempNATUREARCHI = NatureToIcone(RechType)) then
          begin
            stTempCODE := TempTob.GetString('ARP_CODE');
            if (stTempCODE < RechCode) then
            begin
              PosInTob := TempTob.GetIndex;
              continue;
            end else if (stTempCODE = RechCode) then
            begin
              stTempNODOSSIER := TempTob.GetString('ARP_NODOSSIER');
              if (stTempNODOSSIER = RechNoDossier) then
              begin
                { Déjà présent }
                exit;
              end else begin
                if stTempNODOSSIER < RechNoDossier then
                begin
                  PosInTob := TempTob.GetIndex;
                  continue;
                end else begin
                  break;
                end;
              end;
            end else break;
          end else continue;
        end;
      end;
//      { on vérifie qu'il n'est pas déjà présent }
//      if NeedVerif and Assigned(TobBranche.findFirst(['ARP_NATUREARCHI', 'ARP_NODOSSIER', 'CODE'], [RechType, RechNoDossier, RechCode], False)) then //PT2
//        exit;
//      { On vérifie aussi qu'il n'est pas déjà présent avec le champs ARP_NATUREARCHI
//       qui a déjà été transformé en icone }
//      if NeedVerif and Assigned(TobBranche.findFirst(['ARP_NATUREARCHI', 'ARP_NODOSSIER', 'CODE'], [NatureToIcone(RechType), RechNoDossier, RechCode], False)) then  //PT2
//        exit;
      //Fin PT6
      RechTypeLib := RechDom('PGARCHIREGLPAIENATURE', RechType, False);
      RechCodeLib := RechDom(TabletteCode, RechCode, False);
      NewTempRechTob := Tob.Create('Rubrique', TobMere, PosInTob + 1);
      NewTempRechTob.AddChampSupValeur('ARP_NATUREARCHI', RechType);
      NewTempRechTob.AddChampSupValeur('ARP_CODE', RechCode);
      NewTempRechTob.AddChampSupValeur('CODE', RechCode);
      //DEBUT PT2
      NewTempRechTob.AddChampSupValeur('ARP_TYPENATUREARCHI', LType);
      NewTempRechTob.AddChampSupValeur('ARP_PREDEFINI', LPredef);
      NewTempRechTob.AddChampSupValeur('ARP_NODOSSIER', LNoDoss);
      if RechType <> 'TDY' then
      begin
        LibDossier := RechDom('YYDOSSIER', LNoDoss, False);  //PT7
        if LibDossier = 'Error' then LibDossier := '';  //PT7
        NewTempRechTob.AddChampSupValeur('ARP_ORIGINE', RechDom('YYPREDEFINI', LPredef, False)+' '+LibDossier); //PT7
      end else //Pour les tables dynamiques, on ne peut pas savoir quel sera la table utilisée si on a pas le contexte
        NewTempRechTob.AddChampSupValeur('ARP_ORIGINE', '');
//      NewTempRechTob.AddChampSupValeur('ARP_ORIGINE', RechDom('YYPREDEFINI', LPredef, False)+' '+RechDom('YYDOSSIER', LNoDoss, False));
      //FIN PT2
      NewTempRechTob.AddChampSupValeur('ARP_MESSAGE', RechTypeLib+' : ('+RechCode+') '+RechCodeLib);
    end;
  end;
end;

procedure TOF_ANARCHIREGLPAIE.OnArgument (S : String ) ;
var
  Param, TempNature, TempCode, TempPredefini, TempDossier : String;
  TempTob : Tob;
  TV : TTreeView;
  ListeColonnesTV : String;
  IndexElement, IndexNature : Integer;
  LinkedCtrls: TTreeTobLinkedCtrls; //PT1
begin
  Inherited ;
  { Gestion du bouton d'analyse de l'élément sélectionné }
  (GetControl('BTANALYSESELECTELT') as TToolbarButton97).OnClick := OnBTANALYSESELECTELTClick;
  If Assigned(TobAnalyse) then FreeAndNil(TobAnalyse);
  InitMoveProgressForm(nil, 'Chargement des données', 'Chargement en cours...', 9, False, True );
  RubriqueDispatcher := TRubriqueDispatcher.Create;
  MoveCurProgressForm('Chargement des paramètres');
  TV := (GetControl('TVANALYSE') as TTreeView);
  ListeColonnesTV := '';
  TobAnalyse := Tob.Create('Analyse du réglementaire', nil, -1);
  try
    TobAnalyse.AddChampSupValeur('ARP_NATURE', '');
    TobAnalyse.AddChampSupValeur('ARP_NATUREARCHI', '');
    TobAnalyse.AddChampSupValeur('ARP_CODE', 'Analyse du réglementaire');
    TobAnalyse.AddChampSupValeur('ARP_MESSAGE', 'Analyse du réglementaire');
    AddEmptyFields(TobAnalyse);
    ListeColonnesTV := ListeColonnesTV + 'ARP_MESSAGE;';
    TempTob := TOB.Create('Variables', TobAnalyse, -1);
    TempTob.AddChampSupValeur('ARP_NATURE', 'VAR');
    TempTob.AddChampSupValeur('ARP_NATUREARCHI', 'VAR');
    TempTob.AddChampSupValeur('ARP_CODE', 'Variables');
    TempTob.AddChampSupValeur('ARP_MESSAGE', 'Variables');
    AddEmptyFields(TempTob);
    TempTob := TOB.Create('Rémunérations', TobAnalyse, -1);
    TempTob.AddChampSupValeur('ARP_NATURE', 'REM');
    TempTob.AddChampSupValeur('ARP_NATUREARCHI', 'REM');
    TempTob.AddChampSupValeur('ARP_CODE', 'Rémunérations');
    TempTob.AddChampSupValeur('ARP_MESSAGE', 'Rémunérations');
    AddEmptyFields(TempTob);
    TempTob := TOB.Create('Cotisations', TobAnalyse, -1);
    TempTob.AddChampSupValeur('ARP_NATURE', 'COT');
    TempTob.AddChampSupValeur('ARP_NATUREARCHI', 'COT');
    TempTob.AddChampSupValeur('ARP_CODE', 'Cotisations');
    TempTob.AddChampSupValeur('ARP_MESSAGE', 'Cotisations');
    AddEmptyFields(TempTob);
    TempTob := TOB.Create('Cumuls', TobAnalyse, -1);
    TempTob.AddChampSupValeur('ARP_NATURE', 'CUM');
    TempTob.AddChampSupValeur('ARP_NATUREARCHI', 'CUM');
    TempTob.AddChampSupValeur('ARP_CODE', 'Cumuls');
    TempTob.AddChampSupValeur('ARP_MESSAGE', 'Cumuls');
    AddEmptyFields(TempTob);
    TempTob := TOB.Create('Eléments nationaux', TobAnalyse, -1);
    TempTob.AddChampSupValeur('ARP_NATURE', 'ELT');
    TempTob.AddChampSupValeur('ARP_NATUREARCHI', 'ELT');
    TempTob.AddChampSupValeur('ARP_CODE', 'Eléments nationaux');
    TempTob.AddChampSupValeur('ARP_MESSAGE', 'Eléments nationaux');
    AddEmptyFields(TempTob);
//Debut PT4
    TempTob := TOB.Create('Tables dynamiques', TobAnalyse, -1);
    TempTob.AddChampSupValeur('ARP_NATURE', 'TDY');
    TempTob.AddChampSupValeur('ARP_NATUREARCHI', 'TDY');
    TempTob.AddChampSupValeur('ARP_CODE', 'Tables dynamiques');
    TempTob.AddChampSupValeur('ARP_MESSAGE', 'Tables dynamiques');
    AddEmptyFields(TempTob);
//Fin PT4
    TempTob := TOB.Create('Profils de paie', TobAnalyse, -1);
    TempTob.AddChampSupValeur('ARP_NATURE', 'PRO');
    TempTob.AddChampSupValeur('ARP_NATUREARCHI', 'PRO');
    TempTob.AddChampSupValeur('ARP_CODE', 'Profils de paie');
    TempTob.AddChampSupValeur('ARP_MESSAGE', 'Profils de paie');
    AddEmptyFields(TempTob);
    ListeColonnesTV := ListeColonnesTV + 'ARP_MESSAGE;';
    { Lecture des données à analyser }
    Param := readtokenst(S);
    While Param <> '' do
    begin
      TempNature := READTOKENPipe(Param, '|');
      TempCode := Trim(READTOKENPipe(Param, '|'));
      TempPredefini := READTOKENPipe(Param, '|');
      TempDossier := READTOKENPipe(Param, '|');
      AddEltOnTobAnalyse(TempNature, TempCode, TempPredefini, TempDossier);
      Param := readtokenst(S);
    end;
    ListeColonnesTV := ListeColonnesTV + 'ARP_MESSAGE;'; //'"("|ARP_CODE|") "|ARP_LIBELLE;';
    { Analyse des données et mise à jour de la tob }
    MoveCurProgressForm('Analyse en cours...');
    ListeColonnesTV := ListeColonnesTV + 'ARP_MESSAGE;ARP_MESSAGE;';
    for IndexNature := 0 to TobAnalyse.Detail.Count -1 do
    begin
      for IndexElement := 0 to TobAnalyse.Detail.Items[IndexNature].Detail.Count -1 do
      begin
        TempTob := TobAnalyse.Detail.Items[IndexNature].Detail.Items[IndexElement];
        TempNature    := TempTob.GetString('ARP_NATURE');
        TempCode      := Trim(TempTob.GetString('ARP_CODE'));
        TempPredefini := TempTob.GetString('ARP_PREDEFINI');
        TempDossier   := TempTob.GetString('ARP_NODOSSIER');
        AnalyseElt(TempTob, TempNature, TempCode, TempPredefini, TempDossier);
      end;
    end;
    { Calculs auxquels participe X }
    AnalyseInversee;
    { Affichage }
    MoveCurProgressForm('Affichage...');
//PT1    TobAnalyse.PutTreeView(TV, nil, ListeColonnesTV);
  finally
    FiniMoveProgressForm;
  end;
  //Debut PT1
  TV.Visible := False;
  // Ajout des icones à la place des codes Nature
  TobNatureToIcone(TobAnalyse);
  LinkedCtrls := TTreeTobLinkedCtrls.Create();
  try
    LinkedCtrls.BtPrintGrid  := (GetControl('BImprimer') as TToolbarButton97);
    LinkedCtrls.BtExpand     := (GetControl('BTEXPAND') as TToolbarButton97);
//    LinkedCtrls.BtParamListe := (GetControl('BPARAMLISTE') as TToolbarButton97);
    LinkedCtrls.BtSearch     := (GetControl('BRechercher') as TToolbarButton97);
    if Assigned(TreeTobFrame) then FreeAndNil(TreeTobFrame);
    TreeTobFrame := TFFrameTreeTob.Create(Ecran, THPanel(GetControl('PNFRAMETREETOB')), TobAnalyse, 'PGDIAGTREETYPEORI', LinkedCtrls);
  finally
    LinkedCtrls.Free;
  end;
  //Fin PT1
end ;

procedure TOF_ANARCHIREGLPAIE.OnClose ;
begin
  Inherited ;
  If Assigned(TobAnalyse) then FreeAndNil(TobAnalyse);
  FreeAndNil(RubriqueDispatcher);
end ;

procedure TOF_ANARCHIREGLPAIE.OnDisplay () ;
begin
  Inherited ;
end ;

procedure TOF_ANARCHIREGLPAIE.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_ANARCHIREGLPAIE.LectureType(Connecteur: TRubriqueDispatcherConnecteur);
begin
  RubriqueDispatcher.DispatchRubriquesElt(LectureRub);
end;

{ TRubriqueDispatcher }

constructor TRubriqueDispatcher.Create;
begin
{$IFNDEF VER130} //PT8
  SetLength(Connecteurs, 7);
{$ELSE}
  SetLength(Connecteurs, 6);
{$ENDIF}
  { Variables }
  MoveCurProgressForm('Chargement des variables');
  Connecteurs[0] := TRubriqueDispatcherConnecteur.Create(
//PT2                  ' SELECT PVA_VARIABLE, PVA_DATEVALIDITE, PVA_LIBELLE, PVA_THEMEVAR, PVA_PREDEFINI, '
                  '@@ SELECT PVA_VARIABLE, PVA_DATEVALIDITE, PVA_LIBELLE, PVA_THEMEVAR, PVA_PREDEFINI, '//PT2
                + ' PVA_NODOSSIER, PVA_NATUREVAR, '//PT8 PVA_VARPERIODICITE, '
                + ' PVA_BASE0, PVA_BASE1, PVA_BASE2, PVA_BASE3, PVA_BASE4, PVA_BASE5, PVA_BASE6, PVA_BASE7, PVA_BASE8, PVA_BASE9, '
                + ' PVA_RESTHEN0, PVA_RESTHEN1, PVA_RESTHEN2, PVA_RESTHEN3, '
                + ' PVA_RESELSE0, PVA_RESELSE1, PVA_RESELSE2, PVA_RESELSE3, '
                + ' PVA_TYPEBASE0, PVA_TYPEBASE1, PVA_TYPEBASE2, PVA_TYPEBASE3, PVA_TYPEBASE4, PVA_TYPEBASE5, PVA_TYPEBASE6, PVA_TYPEBASE7, PVA_TYPEBASE8, PVA_TYPEBASE9, '
                + ' PVA_TYPERESTHEN0, PVA_TYPERESTHEN1, PVA_TYPERESTHEN2, PVA_TYPERESTHEN3, '
                + ' PVA_TYPERESELSE0, PVA_TYPERESELSE1, PVA_TYPERESELSE2, PVA_TYPERESELSE3 '
//PT2                + ' FROM VARIABLEPAIE WHERE ##PVA_PREDEFINI## AND PVA_TYPEVARIABLE = "PAI"',
{$IFNDEF VER130} //PT8
                + ' FROM VARIABLEPAIE WHERE PVA_TYPEVARIABLE = ''PAI''', //PT2
{$ELSE}
                + ' FROM VARIABLEPAIE', //PT2
{$ENDIF}
                  'PVA_VARIABLE',
                  'VAR',
                  ['PVA_TYPEBASE0', 'PVA_TYPEBASE1', 'PVA_TYPEBASE2', 'PVA_TYPEBASE3', 'PVA_TYPEBASE4',
                 'PVA_TYPEBASE5', 'PVA_TYPEBASE6', 'PVA_TYPEBASE7', 'PVA_TYPEBASE8', 'PVA_TYPEBASE9',
                 'PVA_TYPERESTHEN0', 'PVA_TYPERESTHEN1', 'PVA_TYPERESTHEN2', 'PVA_TYPERESTHEN3',
                 'PVA_TYPERESELSE0', 'PVA_TYPERESELSE1', 'PVA_TYPERESELSE2', 'PVA_TYPERESELSE3'
                 ],
                ['PVA_BASE0', 'PVA_BASE1', 'PVA_BASE2', 'PVA_BASE3', 'PVA_BASE4',
                 'PVA_BASE5', 'PVA_BASE6', 'PVA_BASE7', 'PVA_BASE8', 'PVA_BASE9',
                 'PVA_RESTHEN0', 'PVA_RESTHEN1', 'PVA_RESTHEN2', 'PVA_RESTHEN3',
                 'PVA_RESELSE0', 'PVA_RESELSE1', 'PVA_RESELSE2', 'PVA_RESELSE3'
                 ]
                );
  Connecteurs[0].ChampsType := 'PVA_NATUREVAR';
  Connecteurs[0].TabletteType := 'PGNATUREVAR';
  Connecteurs[0].TabletteLibelle := 'PGVARIABLE';
  { Rémunérations }
  MoveCurProgressForm('Chargement des rémunérations');
  Connecteurs[1] := TRubriqueDispatcherConnecteur.Create(
//PT2                  ' SELECT PRM_RUBRIQUE, PRM_LIBELLE, PRM_ABREGE, PRM_CODECALCUL, PRM_SENSBUL, '
                  '@@ SELECT PRM_RUBRIQUE, PRM_LIBELLE, PRM_ABREGE, PRM_CODECALCUL, PRM_SENSBUL, ' //PT2
                + ' PRM_THEMEREM, PRM_DATEMODIF, PRM_BASEREM, PRM_BASEMTQTE, PRM_TAUXREM, '
                + ' PRM_TAUXMTQTE, PRM_COEFFREM, PRM_COEFFMTQTE, PRM_IMPRIMABLE, PRM_BASEIMPRIMABLE, '
                + ' PRM_TAUXIMPRIMABLE, PRM_COEFFIMPRIM, PRM_MONTANT, PRM_MONTANTMTQTE, PRM_BASETYPARR, '
                + ' PRM_BASEMTARR, PRM_MTTYPARR, PRM_MTMTARR, PRM_DADS, PRM_PREDEFINI, PRM_DATECREATION, '
                + ' PRM_VISIBLE, PRM_SOUSTYPEREM, PRM_NODOSSIER, PRM_TYPEBASE, PRM_TYPETAUX, PRM_TYPECOEFF, '
                + ' PRM_TYPEMONTANT, PRM_TYPEMINI, PRM_VALEURMINI, PRM_TYPEMAXI, PRM_VALEURMAXI, PRM_DU, '
                + ' PRM_AU, PRM_MOIS1, PRM_MOIS2, PRM_MOIS3, PRM_MOIS4, '
//PT8                + ' PRM_MOIS5, PRM_MOIS6, '
                + ' PRM_NATURERUB, '
                + ' PRM_ORDREETAT, PRM_DECBASE, PRM_DECTAUX, PRM_DECCOEFF, PRM_DECMONTANT, PRM_AVANTAGENAT, '
                + ' PRM_RETSALAIRE, PRM_FRAISPROF, PRM_CHQVACANCE, PRM_IMPOTRETSOURC, PRM_INDEXPATRIAT, '
                + ' PRM_SOMISOL, PRM_INDEMCONVENT, PRM_INDEMPREAVIS, PRM_INDEMCP, PRM_INDEMLEGALES, '
                + ' PRM_INDEMINTEMP, PRM_INDEMCOMPCP, PRM_HCHOMPAR, PRM_BTPSALAIRE, PRM_BTPARRET, '
                + ' PRM_ACTIVITE, PRM_BLOCNOTE, PRM_SAISIEARRETAC, PRM_EXCLMAINT, PRM_EXCLENVERS, '
                + ' PRM_METHODARRONDI, PRM_REMLIEREGUL'//PT8, PRM_LIBCONTRAT, PRM_PRIMEASSEDIC '
//PT2                + ' FROM REMUNERATION WHERE ##PRM_PREDEFINI##',
                + ' FROM REMUNERATION', //PT2
                  'PRM_RUBRIQUE',
                  'REM',
                  ['PRM_TYPEBASE', 'PRM_TYPETAUX', 'PRM_TYPECOEFF', 'PRM_TYPEMONTANT', 'PRM_TYPEMINI',
                 'PRM_TYPEMAXI'
                 ],
                ['PRM_BASEREM', 'PRM_TAUXREM', 'PRM_COEFFREM', 'PRM_MONTANT', 'PRM_VALEURMINI',
                 'PRM_VALEURMAXI'
                 ]
                );
  Connecteurs[1].ChampsType := 'PRM_THEMEREM';
  Connecteurs[1].TabletteType := 'PGTHEMEREM';
  Connecteurs[1].TabletteLibelle := 'PGREMUNERATION';
  { Cotisations }
  MoveCurProgressForm('Chargement des cotisations');
  Connecteurs[2] := TRubriqueDispatcherConnecteur.Create(
//PT2                  ' SELECT PCT_RUBRIQUE, PCT_LIBELLE, PCT_PREDEFINI, PCT_THEMECOT, PCT_DATEMODIF, PCT_ORGANISME, '
                  '@@ SELECT PCT_RUBRIQUE, PCT_LIBELLE, PCT_PREDEFINI, PCT_THEMECOT, PCT_DATEMODIF, PCT_ORGANISME, '//PT2
                + ' PCT_BASECOTISATION, PCT_TAUXSAL, PCT_FFSAL, PCT_TAUXPAT, PCT_FFPAT, PCT_IMPRIMABLE, PCT_BASEIMP, '
                + ' PCT_TXSALIMP, PCT_FFSALIMP, PCT_TXPATIMP, PCT_FFPATIMP, PCT_CODETRANCHE, PCT_BASEMTARR, '
                + ' PCT_BASETYPARR, PCT_FFSALARR, PCT_FFSALTYPARR, PCT_FFPATARR, PCT_FFPATTYPARR, PCT_ABREGE, '
                + ' PCT_DATECREATION, PCT_NODOSSIER, PCT_TYPEBASE, PCT_TYPETAUXSAL, PCT_TYPETAUXPAT, PCT_TYPEFFSAL, '
                + ' PCT_TYPEFFPAT, PCT_NATURERUB, PCT_PLAFOND, PCT_TYPEMINISAL, PCT_VALEURMINISAL, PCT_TYPEMAXISAL, '
                + ' PCT_VALEURMAXISAL, PCT_TYPEMINIPAT, PCT_VALEURMINIPAT, PCT_TYPEMAXIPAT, PCT_VALEURMAXIPAT, '
                + ' PCT_DU, PCT_AU, PCT_MOIS1, PCT_MOIS2, PCT_MOIS3, PCT_MOIS4, PCT_DECBASE, PCT_DECTXSAL, '
                + ' PCT_DECMTSAL, PCT_DECTXPAT, PCT_DECMTPAT, PCT_DECBASECOT, PCT_TRANCHE1, PCT_TRANCHE2, PCT_TRANCHE3, '
                + ' PCT_TYPEPLAFOND, PCT_TYPETRANCHE1, PCT_TYPETRANCHE2, PCT_TYPETRANCHE3, PCT_SOUMISREGUL, '
                + ' PCT_ORDREETAT, PCT_ORDREAT, PCT_BASECSGCRDS, PCT_CASPART, PCT_BASECRDS, PCT_BRUTSS, PCT_PLAFONDSS, '
                + ' PCT_BRUTFISC, PCT_NETIMPO, PCT_DADSTOTIMPTSS, PCT_DADSMONTTSS, '
//PT8                + ' PCT_DADSBASEPREV, '
                + ' PCT_TYPEREGUL, '
                + ' PCT_ASSIETTEBRUT, PCT_SOUMISMALAD, PCT_PRECOMPTEASS, PCT_ETUDEDROIT, PCT_REDUCBASSAL, PCT_REDUCREPAS, '
                + ' PCT_ALLEGEMENTA2, PCT_MAJORATA2, PCT_MINORATA2, PCT_DADSEXOBASE, PCT_DADSEPARGNE, PCT_DADSEXOCOT, '
                + ' PCT_PRESFINMOIS, PCT_ACTIVITE, PCT_BLOCNOTE '
//PT2                + ' FROM COTISATION WHERE ##PCT_PREDEFINI##',
                + ' FROM COTISATION', //PT2
                  'PCT_RUBRIQUE',
                  'COT',
                  ['PCT_TYPEBASE', 'PCT_TYPETAUXSAL', 'PCT_TYPETAUXPAT', 'PCT_TYPEFFSAL', 'PCT_TYPEFFPAT',
                 'PCT_TYPEMINISAL', 'PCT_TYPEMAXISAL', 'PCT_TYPEMINIPAT', 'PCT_TYPEMAXIPAT', 'PCT_TYPEPLAFOND',
                 'PCT_TYPETRANCHE1', 'PCT_TYPETRANCHE2', 'PCT_TYPETRANCHE3'
                 ],
                ['PCT_BASECOTISATION', 'PCT_TAUXSAL', 'PCT_TAUXPAT', 'PCT_FFSAL', 'PCT_FFPAT',
                 'PCT_VALEURMINISAL', 'PCT_VALEURMAXISAL', 'PCT_VALEURMINIPAT', 'PCT_VALEURMAXIPAT', 'PCT_PLAFOND',
                 'PCT_TRANCHE1', 'PCT_TRANCHE2', 'PCT_TRANCHE3'
                 ]
                );
  Connecteurs[2].ChampsType := 'PCT_THEMECOT';
  Connecteurs[2].TabletteType := 'PGTHEMECOT';
  Connecteurs[2].TabletteLibelle := 'PGCOTISATION';
  { Cumuls }
  MoveCurProgressForm('Chargement des cumuls');
  Connecteurs[3] := TRubriqueDispatcherConnecteur.Create(
//PT2                  ' SELECT PCR_NATURERUB, PCR_RUBRIQUE, PCR_CUMULPAIE, PCR_PREDEFINI, PCR_NODOSSIER, PCR_LIBELLE, '
                  '@@ SELECT PCR_NATURERUB, PCR_RUBRIQUE, PCR_CUMULPAIE, PCR_PREDEFINI, PCR_NODOSSIER, PCR_LIBELLE, '//PT2
                + ' PCR_SENS, PCR_DATECREATION, PCR_DATEMODIF '
//PT2                + ' FROM CUMULRUBRIQUE WHERE ##PCR_PREDEFINI##',
                + ' FROM CUMULRUBRIQUE',//PT2
                  'PCR_CUMULPAIE',
                  'CUM',
                  ['PCR_NATURERUB'],
                  ['PCR_RUBRIQUE' ]
                );
  Connecteurs[3].ChampsType := 'PCR_SENS';
  Connecteurs[3].TabletteType := '';
  Connecteurs[3].TabletteLibelle := 'PGCUMULPAIE';
  { Profils }
  MoveCurProgressForm('Chargement des profils');
  Connecteurs[4] := TRubriqueDispatcherConnecteur.Create(
//PT2                  ' SELECT PPM_TYPEPROFIL, PPM_PROFIL, PPM_NATURERUB, PPM_RUBRIQUE, PPM_LIBELLE, PPM_SUPPRESSION, '
                  '@@ SELECT PPM_TYPEPROFIL, PPM_PROFIL, PPM_NATURERUB, PPM_RUBRIQUE, PPM_LIBELLE, PPM_SUPPRESSION, '//PT2
                + ' PPM_VISIBILITE, PPM_PREDEFINI, PPM_IMPRIMABLE, PPM_NODOSSIER, PPM_DATECREATION, PPM_DATEMODIF '
//PT2                + ' FROM PROFILRUB WHERE ##PPM_PREDEFINI##',
                + ' FROM PROFILRUB',//PT2
                  'PPM_PROFIL',
                  'PRO',
                  ['PPM_NATURERUB'],
                  ['PPM_RUBRIQUE' ]
                );
  Connecteurs[4].ChampsType := 'PPM_TYPEPROFIL';
  Connecteurs[4].TabletteType := 'PGTYPEPROFIL';
  Connecteurs[4].TabletteLibelle := 'PGPROFIL';
//Debut PT3
  { Eléments nationaux }
  MoveCurProgressForm('Chargement des éléments nationaux');
  Connecteurs[5] := TRubriqueDispatcherConnecteur.Create(
                  '@@ SELECT PEL_CODEELT, PEL_DATEVALIDITE, PEL_LIBELLE, PEL_THEMEELT, PEL_PREDEFINI, PEL_NODOSSIER '
                + 'FROM ELTNATIONAUX',
                  'PEL_CODEELT',
                  'ELT',
//PT8 En Delphi 5, un bug coromp la mémoire si on passe un tableau vide en paramètre
                  [''],
                  ['']
                );
  Connecteurs[5].ChampsType := 'PEL_THEMEELT';
  Connecteurs[5].TabletteType := 'PGTHEMEELEMENTNAT';
  Connecteurs[5].TabletteLibelle := 'PGELEMENTNAT';
//Fin PT3
//Debut PT4
{$IFNDEF VER130} //PT8
  { Tables dynamiques }
  MoveCurProgressForm('Chargement des tables dynamiques');
  Connecteurs[6] := TRubriqueDispatcherConnecteur.Create(
                  '@@ SELECT PTE_CODTABL, PTE_PREDEFINI, PTE_NODOSSIER, PTE_DTVALID, PTE_NIVSAIS, PTE_VALNIV, '
                 +'PTE_PRIORITENIV, PTE_LIBELLE, PTE_NATURETABLE, PTE_TYPERESULTAT '
                 +' FROM TABLEDIMENT',
                  'PTE_CODTABL',
                  'TDY',
                  [],
                  []
                );
  Connecteurs[6].ChampsType := 'PTE_NATURETABLE';
  Connecteurs[6].TabletteType := 'PGNATURETABLE';
  Connecteurs[6].TabletteLibelle := 'PGTABLEDIM';
{$ENDIF}
//Fin PT4
end;

destructor TRubriqueDispatcher.Destroy;
var
  Index : Integer;
begin
  for Index := 0 to Length(Connecteurs) -1 do
  begin
    FreeAndNil(Connecteurs[Index]);
  end;
  SetLength(Connecteurs, 0);
  inherited;
end;

procedure TRubriqueDispatcher.DispatchRubriques(
  ProcedureToDispatch: TProcedureDispatchRub);
var
  Index : Integer;
begin
  for Index := 0 to Length(Connecteurs) -1 do
  begin
    ProcedureToDispatch(Connecteurs[Index]);
  end;
end;

procedure TRubriqueDispatcher.DispatchRubriquesElt(
  ProcedureToDispatch: TProcedureDispatchRubElt);
var
  IndexConnecteur, IndexDetail, IndexElt : Integer;
  TempEltTob : Tob;
  IndexDetail2, countDetail1, CountDetail2 : Integer; // PT5
begin
  for IndexConnecteur := 0 to Length(Connecteurs) -1 do
  begin
    countDetail1 := (Connecteurs[IndexConnecteur].TobMereRubriques.Detail.count -1);
    For IndexDetail := 0 to countDetail1 do
    begin
      //PT5
      CountDetail2 := (Connecteurs[IndexConnecteur].TobMereRubriques.Detail[IndexDetail].Detail.count -1);
      For IndexDetail2 := 0 to CountDetail2 do
      begin

        TempEltTob := Connecteurs[IndexConnecteur].TobMereRubriques.Detail[IndexDetail].detail[IndexDetail2];
        for IndexElt := 0 to Length(Connecteurs[IndexConnecteur].ListeChampsTypeElt) -1 do
        begin
          ProcedureToDispatch(TempEltTob, Connecteurs[IndexConnecteur].TypeRubriques, Connecteurs[IndexConnecteur].ListeChampsTypeElt[IndexElt], Connecteurs[IndexConnecteur].ListeChampsCodeElt[IndexElt], Connecteurs[IndexConnecteur].Prefixe+'_NODOSSIER');
        end;

      end;
    end;
  end;
end;

function TRubriqueDispatcher.GetConnecteur(
  TypeDeRubriques: String): TRubriqueDispatcherConnecteur;
var
  Index : Integer;
begin
  Result := nil;
  for Index := 0 to Length(Connecteurs) -1 do
  begin
    if Connecteurs[Index].TypeRubriques = TypeDeRubriques then
    begin
      Result := Connecteurs[Index];
      exit;
    end;
  end;
end;

{ TRubriqueDispatcherConnecteur }

constructor TRubriqueDispatcherConnecteur.Create(stSQL, NomChampsCode : String; TypeDesRubriques : String; ListeDesChampsTypeElt, ListeDesChampsCodeElt : Array of String);
var
  indexTab : Integer;
  TempTob : Tob;
begin
  if Length(ListeDesChampsTypeElt) <> Length(ListeDesChampsCodeElt) then
  begin
    Raise Exception.Create('Vous devez spécifier autant de type d''éléments que d''éléments.');
  end;
  //Debut PT5
  //  TobMereRubriques := Tob.create('Définitions des rubriques de type '+TypeRubriques, nil, -1);
  //TobMereRubriques.LoadDetailFromSQL(stSQL);
  TempTob := Tob.create('Définitions des rubriques de type '+TypeRubriques, nil, -1);
  TempTob.LoadDetailFromSQL(stSQL);
  Prefixe := LeftStr(NomChampsCode, 3);   //PT2
  { On éclate la tob des compteurs par salarié }
  EclateTob(TempTob, TobMereRubriques, Prefixe+'_NODOSSIER', False);
  FreeAndNil(TempTob);
  //Fin PT5
  TypeRubriques := TypeDesRubriques;
  If Not ((Length(ListeDesChampsTypeElt) = 1) And (ListeDesChampsTypeElt[0] = '')) Then //PT8
  Begin
    SetLength(ListeChampsTypeElt, Length(ListeDesChampsTypeElt));
    for indexTab := 0 to Length(ListeDesChampsTypeElt) -1 do
    begin
      ListeChampsTypeElt[indexTab] := ListeDesChampsTypeElt[indexTab];
    end;
  End Else Begin  //PT8
    SetLength(ListeChampsTypeElt,0);
  end;
  If Not ((Length(ListeDesChampsCodeElt) = 1) And (ListeDesChampsCodeElt[0] = '')) Then //PT8
  Begin
    SetLength(ListeChampsCodeElt, Length(ListeDesChampsCodeElt));
    for indexTab := 0 to Length(ListeDesChampsCodeElt) -1 do
    begin
      ListeChampsCodeElt[indexTab] := ListeDesChampsCodeElt[indexTab];
    end;
  End Else Begin  //PT8
    SetLength(ListeChampsCodeElt,0);
  end;
  ChampsCode := NomChampsCode;
  ChampsType := '';//PT2
end;

destructor TRubriqueDispatcherConnecteur.Destroy;
begin
  if Assigned(TobMereRubriques) then FreeAndNil(TobMereRubriques);
  inherited;
end;

//debut PT2
procedure TRubriqueDispatcherConnecteur.GetPredefiniNodossierVisibleElt(
  var TypeElt, Predef, Nodoss: String; CodeElt, NoDossUtilisateur: String);
var
  TempTob : Tob;
begin
  Predef := '';
  Nodoss := '';
  TypeElt := '';
  TempTob := GetVisibleElt(CodeElt, NoDossUtilisateur);
  if Assigned(TempTob) then
  begin
    Predef := TempTob.GetString(Prefixe+'_PREDEFINI');
    Nodoss := TempTob.GetString(Prefixe+'_NODOSSIER');
    if ChampsType <> '' then
    begin
      if TabletteType <> '' then
        TypeElt := RechDom(TabletteType, TempTob.GetString(ChampsType), False)
      else
        TypeElt := TempTob.GetString(ChampsType);
    end;
  end;
end;

Function TRubriqueDispatcherConnecteur.GetVisibleElt(CodeElt, NoDossUtilisateur : String) : Tob;
var
  TempTob : Tob;
begin
  TempTob := nil;
  // Pour les tables dynamiques, on ne peut pas savoir quel niveau sera utilisé
  // donc on récupère le 1er qu'on trouve
  if TypeRubriques = 'TDY' then
  begin
    TempTob := TobMereRubriques.FindFirst([ChampsCode], [CodeElt], True);
  end else begin
    if NoDossUtilisateur <> '000000' then
    begin
      //PT5
      TempTob := TobMereRubriques.FindFirst([Prefixe+'_NODOSSIER'], [NoDossUtilisateur], False);
      if Assigned(TempTob) then
        TempTob := TempTob.FindFirst([ChampsCode], [CodeElt], False);
    end;
    if not Assigned(TempTob) then
    begin
      //PT5
      TempTob := TobMereRubriques.FindFirst([Prefixe+'_NODOSSIER'], ['000000'], False);
      if Assigned(TempTob) then
        TempTob := TempTob.FindFirst([ChampsCode], [CodeElt], False);
    end;
  end;
  result := TempTob;
end;
//Fin PT2

procedure TOF_ANARCHIREGLPAIE.AnalyseElt(TobElt : Tob; Nature, Code, Predefini, Dossier : String);
var
  NewTempTob, TempEltTob : Tob;
  TempEltTob2 : Tob;//PT5
begin
  if Nature = 'VAR' then { Rubriques qui participent au calcul de la variable X }
  begin
    NewTempTob := Tob.Create('Eléments qui participent au calcul de la variable', TobElt, -1);
    NewTempTob.AddChampSupValeur('ARP_CODE', NewTempTob.NomTable);
    NewTempTob.AddChampSupValeur('ARP_MESSAGE', NewTempTob.NomTable);
    NewTempTob.AddChampSupValeur('CODEBRANCHE', 1);
    NewTempTob.AddChampSupValeur('ARP_NATUREARCHI', '');
    NewTempTob.AddChampSupValeur('ARP_TYPENATUREARCHI', '');
    NewTempTob.AddChampSupValeur('ARP_PREDEFINI', '');
    NewTempTob.AddChampSupValeur('ARP_NODOSSIER', '');
    NewTempTob.AddChampSupValeur('ARP_ORIGINE', '');
    //PT5
    TempEltTob := RubriqueDispatcher.getConnecteur('VAR').TobMereRubriques.FindFirst(['PVA_NODOSSIER'], [Dossier], False);
    if Assigned(TempEltTob) then
      TempEltTob := TempEltTob.FindFirst(['PVA_VARIABLE', 'PVA_PREDEFINI'], [Code, Predefini], False);
    if Assigned(TempEltTob) then
    begin
      RechercheRubriquesLiees(NewTempTob, TempEltTob, Nature, 'PVA_TYPEBASE0', 'PVA_BASE0', 'PGTYPECHAMPVAR');
      RechercheRubriquesLiees(NewTempTob, TempEltTob, Nature, 'PVA_TYPEBASE1', 'PVA_BASE1', 'PGTYPECHAMPVAR');
      RechercheRubriquesLiees(NewTempTob, TempEltTob, Nature, 'PVA_TYPEBASE2', 'PVA_BASE2', 'PGTYPECHAMPVAR');
      RechercheRubriquesLiees(NewTempTob, TempEltTob, Nature, 'PVA_TYPEBASE3', 'PVA_BASE3', 'PGTYPECHAMPVAR');
      RechercheRubriquesLiees(NewTempTob, TempEltTob, Nature, 'PVA_TYPEBASE4', 'PVA_BASE4', 'PGTYPECHAMPVAR');
      RechercheRubriquesLiees(NewTempTob, TempEltTob, Nature, 'PVA_TYPEBASE5', 'PVA_BASE5', 'PGTYPECHAMPVAR');
      RechercheRubriquesLiees(NewTempTob, TempEltTob, Nature, 'PVA_TYPEBASE6', 'PVA_BASE6', 'PGTYPECHAMPVAR');
      RechercheRubriquesLiees(NewTempTob, TempEltTob, Nature, 'PVA_TYPEBASE7', 'PVA_BASE7', 'PGTYPECHAMPVAR');
      RechercheRubriquesLiees(NewTempTob, TempEltTob, Nature, 'PVA_TYPEBASE8', 'PVA_BASE8', 'PGTYPECHAMPVAR');
      RechercheRubriquesLiees(NewTempTob, TempEltTob, Nature, 'PVA_TYPEBASE9', 'PVA_BASE9', 'PGTYPECHAMPVAR');
      RechercheRubriquesLiees(NewTempTob, TempEltTob, Nature, 'PVA_TYPERESTHEN0', 'PVA_RESTHEN0', 'PGTYPECHAMPVAR');
      RechercheRubriquesLiees(NewTempTob, TempEltTob, Nature, 'PVA_TYPERESTHEN1', 'PVA_RESTHEN1', 'PGTYPECHAMPVAR');
      RechercheRubriquesLiees(NewTempTob, TempEltTob, Nature, 'PVA_TYPERESTHEN2', 'PVA_RESTHEN2', 'PGTYPECHAMPVAR');
      RechercheRubriquesLiees(NewTempTob, TempEltTob, Nature, 'PVA_TYPERESTHEN3', 'PVA_RESTHEN3', 'PGTYPECHAMPVAR');
      RechercheRubriquesLiees(NewTempTob, TempEltTob, Nature, 'PVA_TYPERESELSE0', 'PVA_RESELSE0', 'PGTYPECHAMPVAR');
      RechercheRubriquesLiees(NewTempTob, TempEltTob, Nature, 'PVA_TYPERESELSE1', 'PVA_RESELSE1', 'PGTYPECHAMPVAR');
      RechercheRubriquesLiees(NewTempTob, TempEltTob, Nature, 'PVA_TYPERESELSE2', 'PVA_RESELSE2', 'PGTYPECHAMPVAR');
      RechercheRubriquesLiees(NewTempTob, TempEltTob, Nature, 'PVA_TYPERESELSE3', 'PVA_RESELSE3', 'PGTYPECHAMPVAR');
    end;
  end else if Nature = 'REM' then { Rubriques de la rémunération X }
  begin
    NewTempTob := Tob.Create('Eléments qui participent au calcul de la rémunération', TobElt, -1);
    NewTempTob.AddChampSupValeur('ARP_MESSAGE', NewTempTob.NomTable);
    NewTempTob.AddChampSupValeur('CODEBRANCHE', 1);
    NewTempTob.AddChampSupValeur('ARP_CODE', NewTempTob.NomTable);
    NewTempTob.AddChampSupValeur('ARP_NATUREARCHI', '');
    NewTempTob.AddChampSupValeur('ARP_TYPENATUREARCHI', '');
    NewTempTob.AddChampSupValeur('ARP_PREDEFINI', '');
    NewTempTob.AddChampSupValeur('ARP_NODOSSIER', '');
    NewTempTob.AddChampSupValeur('ARP_ORIGINE', '');
    //PT5
    TempEltTob := RubriqueDispatcher.getConnecteur('REM').TobMereRubriques.FindFirst(['PRM_NODOSSIER'], [Dossier], False);
    if Assigned(TempEltTob) then
      TempEltTob := TempEltTob.FindFirst(['PRM_RUBRIQUE', 'PRM_PREDEFINI'], [Code, Predefini], False);
    if Assigned(TempEltTob) then
    begin
      RechercheRubriquesLiees(NewTempTob, TempEltTob, Nature, 'PRM_TYPEBASE', 'PRM_BASEREM', 'PGTYPECHAMPREM');
      RechercheRubriquesLiees(NewTempTob, TempEltTob, Nature, 'PRM_TYPETAUX', 'PRM_TAUXREM', 'PGTYPECHAMPREM');
      RechercheRubriquesLiees(NewTempTob, TempEltTob, Nature, 'PRM_TYPECOEFF', 'PRM_COEFFREM', 'PGTYPECHAMPREM');
      RechercheRubriquesLiees(NewTempTob, TempEltTob, Nature, 'PRM_TYPEMONTANT', 'PRM_MONTANT', 'PGTYPECHAMPREM');
      RechercheRubriquesLiees(NewTempTob, TempEltTob, Nature, 'PRM_TYPEMINI', 'PRM_VALEURMINI', 'PGTYPECHAMPREM');
      RechercheRubriquesLiees(NewTempTob, TempEltTob, Nature, 'PRM_TYPEMAXI', 'PRM_VALEURMAXI', 'PGTYPECHAMPREM');
    end;
  end else if Nature = 'COT' then { Rubriques de la cotisation X }
  begin
    NewTempTob := Tob.Create('Eléments qui participent au calcul de la cotisation', TobElt, -1);
    NewTempTob.AddChampSupValeur('ARP_MESSAGE', NewTempTob.NomTable);
    NewTempTob.AddChampSupValeur('CODEBRANCHE', 1);
    NewTempTob.AddChampSupValeur('ARP_CODE', NewTempTob.NomTable);
    NewTempTob.AddChampSupValeur('ARP_NATUREARCHI', '');
    NewTempTob.AddChampSupValeur('ARP_TYPENATUREARCHI', '');
    NewTempTob.AddChampSupValeur('ARP_PREDEFINI', '');
    NewTempTob.AddChampSupValeur('ARP_NODOSSIER', '');
    NewTempTob.AddChampSupValeur('ARP_ORIGINE', '');
    //PT5
    TempEltTob := RubriqueDispatcher.getConnecteur('COT').TobMereRubriques.FindFirst(['PCT_NODOSSIER'], [Dossier], False);
    if Assigned(TempEltTob) then
      TempEltTob := TempEltTob.FindFirst(['PCT_RUBRIQUE', 'PCT_PREDEFINI'], [Code, Predefini], False);
    if Assigned(TempEltTob) then
    begin
      RechercheRubriquesLiees(NewTempTob, TempEltTob, Nature, 'PCT_TYPEBASE', 'PCT_BASECOTISATION', 'PGTYPEBASECOT');
      RechercheRubriquesLiees(NewTempTob, TempEltTob, Nature, 'PCT_TYPETAUXSAL', 'PCT_TAUXSAL', 'PGTYPECHAMPCOT');
      RechercheRubriquesLiees(NewTempTob, TempEltTob, Nature, 'PCT_TYPETAUXPAT', 'PCT_TAUXPAT', 'PGTYPECHAMPCOT');
      RechercheRubriquesLiees(NewTempTob, TempEltTob, Nature, 'PCT_TYPEFFSAL', 'PCT_FFSAL', 'PGTYPECHAMPCOT');
      RechercheRubriquesLiees(NewTempTob, TempEltTob, Nature, 'PCT_TYPEFFPAT', 'PCT_FFPAT', 'PGTYPECHAMPCOT');
      RechercheRubriquesLiees(NewTempTob, TempEltTob, Nature, 'PCT_TYPEMINISAL', 'PCT_VALEURMINISAL', 'PGTYPEMINMAX');
      RechercheRubriquesLiees(NewTempTob, TempEltTob, Nature, 'PCT_TYPEMAXISAL', 'PCT_VALEURMAXISAL', 'PGTYPEMINMAX');
      RechercheRubriquesLiees(NewTempTob, TempEltTob, Nature, 'PCT_TYPEMINIPAT', 'PCT_VALEURMINIPAT', 'PGTYPEMINMAX');
      RechercheRubriquesLiees(NewTempTob, TempEltTob, Nature, 'PCT_TYPEMAXIPAT', 'PCT_VALEURMAXIPAT', 'PGTYPEMINMAX');
      RechercheRubriquesLiees(NewTempTob, TempEltTob, Nature, 'PCT_TYPEPLAFOND', 'PCT_PLAFOND', 'PGTYPEPLAFOND');
      RechercheRubriquesLiees(NewTempTob, TempEltTob, Nature, 'PCT_TYPETRANCHE1', 'PCT_TRANCHE1', 'PGTYPETRANCHE');
      RechercheRubriquesLiees(NewTempTob, TempEltTob, Nature, 'PCT_TYPETRANCHE2', 'PCT_TRANCHE2', 'PGTYPETRANCHE');
      RechercheRubriquesLiees(NewTempTob, TempEltTob, Nature, 'PCT_TYPETRANCHE3', 'PCT_TRANCHE3', 'PGTYPETRANCHE');
    end;
  end else if Nature = 'CUM' then  { Rubriques du Cumul X }
  begin
    NewTempTob := Tob.Create('Eléments qui participent au calcul du cumul', TobElt, -1);
    NewTempTob.AddChampSupValeur('ARP_MESSAGE', NewTempTob.NomTable);
    NewTempTob.AddChampSupValeur('CODEBRANCHE', 1);
    NewTempTob.AddChampSupValeur('ARP_CODE', NewTempTob.NomTable);
    NewTempTob.AddChampSupValeur('ARP_NATUREARCHI', '');
    NewTempTob.AddChampSupValeur('ARP_TYPENATUREARCHI', '');
    NewTempTob.AddChampSupValeur('ARP_PREDEFINI', '');
    NewTempTob.AddChampSupValeur('ARP_NODOSSIER', '');
    NewTempTob.AddChampSupValeur('ARP_ORIGINE', '');
    //PT5
    TempEltTob2 := RubriqueDispatcher.getConnecteur('CUM').TobMereRubriques.FindFirst(['PCR_NODOSSIER'], [Dossier], False);
    if Assigned(TempEltTob2) then
    begin
      TempEltTob := TempEltTob2.FindFirst(['PCR_CUMULPAIE', 'PCR_PREDEFINI'], [Code, Predefini], False);
      While Assigned(TempEltTob) do
      begin
        RechercheRubriquesLiees(NewTempTob, TempEltTob, Nature, 'PCR_NATURERUB', 'PCR_RUBRIQUE', 'PGNATURERUBRIQUE');   //PGTYPEBASECOT
        TempEltTob := TempEltTob2.FindNext(['PCR_CUMULPAIE', 'PCR_PREDEFINI'], [Code, Predefini], False);
      end;
    end;
  end else if Nature = 'PRO' then  { Rubriques du Profil X }
  begin
    NewTempTob := Tob.Create('Eléments qui sont associés au profil', TobElt, -1);
    NewTempTob.AddChampSupValeur('ARP_MESSAGE', NewTempTob.NomTable);
    NewTempTob.AddChampSupValeur('CODEBRANCHE', 1);
    NewTempTob.AddChampSupValeur('ARP_CODE', NewTempTob.NomTable);
    NewTempTob.AddChampSupValeur('ARP_NATUREARCHI', '');
    NewTempTob.AddChampSupValeur('ARP_TYPENATUREARCHI', '');
    NewTempTob.AddChampSupValeur('ARP_PREDEFINI', '');
    NewTempTob.AddChampSupValeur('ARP_NODOSSIER', '');
    NewTempTob.AddChampSupValeur('ARP_ORIGINE', '');
    //PT5
    TempEltTob2 := RubriqueDispatcher.getConnecteur('PRO').TobMereRubriques.FindFirst(['PPM_NODOSSIER'], [Dossier], False);
    if Assigned(TempEltTob2) then
    begin
      TempEltTob := TempEltTob2.FindFirst(['PPM_PROFIL', 'PPM_PREDEFINI'], [Code, Predefini], False);
//      TempEltTob := RubriqueDispatcher.getConnecteur('PRO').TobMereRubriques.FindFirst(['PPM_PROFIL', 'PPM_PREDEFINI', 'PPM_NODOSSIER'], [Code, Predefini, Dossier], False);
      While Assigned(TempEltTob) do
      begin
        RechercheRubriquesLiees(NewTempTob, TempEltTob, Nature, 'PPM_NATURERUB', 'PPM_RUBRIQUE', 'PGNATURERUBRIQUE');     //PGPROFIL
        TempEltTob := TempEltTob2.FindNext(['PPM_PROFIL', 'PPM_PREDEFINI'], [Code, Predefini], False);
//        TempEltTob := RubriqueDispatcher.getConnecteur('PRO').TobMereRubriques.FindNext(['PPM_PROFIL'], [Code], False);
      end;
    end;
  end;
end;

procedure TOF_ANARCHIREGLPAIE.OnBTANALYSESELECTELTClick(Sender: TObject);
var
  Nature, Code, Predefini, Dossier : String;
  TempTob : Tob;
begin
  { On récupère l'élément sélectionné dans la grille }
  TempTob := TreeTobFrame.TobRow[TreeTobFrame.CurrentRow];
  Nature    := IconeToNature(TempTob.GetString('ARP_NATUREARCHI'));
  if nature = '' then exit;
  Code      := TempTob.GetString('ARP_CODE');
  { Evite l'analyse des lignes de présentation }
  if   (Code = 'Analyse du réglementaire') or (Code = 'Variables')
    or (Code = 'Rémunérations') or (Code = 'Cotisations')
    or (Code = 'Cumuls') or (Code = 'Eléments nationaux')
    or (Code = 'Tables dynamiques') or (Code = 'Profils de paie') then
    exit;
  Predefini := TempTob.GetString('ARP_PREDEFINI');
  Dossier   := TempTob.GetString('ARP_NODOSSIER');
//  { On vérifie qu'on ai pas déjà analyser l'élément }
//  RechTob := TobAnalyse.FindFirst(['ARP_NATURE'], [Nature], False);
//  if Assigned(RechTob) then
//    RechTob := RechTob.FindFirst(['ARP_CODE'], [Code], False);
//  if Assigned(RechTob) then
//    exit;
  InitMoveProgressForm(nil, 'Analyse du réglementaire', 'Analyse en cours...', 2, False, True );
  try
    TempTob := AddEltOnTobAnalyse(Nature, Code, Predefini, Dossier);
    if Assigned(TempTob) then
    begin
      MoveCurProgressForm('Analyse du nouvel élément');
      AnalyseElt(TempTob, Nature, Code, Predefini, Dossier);
      { Calculs auxquels participe X }
      AnalyseInversee;
      { Affichage }
      MoveCurProgressForm('Affichage...');
      { Ajout des icones à la place des codes Nature }
      TobNatureToIcone(TobAnalyse);
      TreeTobFrame.RefreshTreeTob;
    end;
  finally
    FiniMoveProgressForm;
  end;

end;

function TOF_ANARCHIREGLPAIE.AddEltOnTobAnalyse(Nature, Code, Predefini,
  Dossier : String): Tob;
var
  TempTob : Tob; //, RechTob : Tob;
  RubDispatchConnecteur : TRubriqueDispatcherConnecteur; //PT2
  TpTob : Tob; //PT2
  TypeElt : String; //PT2

  TriTempTob :Tob; //PT6
  PosInTob, IndexBranche : Integer; //PT6
  stTempNATUREARCHI, stTempCODE : String; // PT6
  stTempDossier, stTempPredefini : String; //PT6

  LibDossier : String; //PT7

begin
  result := Nil;
  if nature = '' then exit;
  TempTob := TobAnalyse.FindFirst(['ARP_NATURE'], [Nature], False);
  if Assigned(TempTob) then
  begin


    //debut PT6 Tri des éléments
    //On recherche a quel index on doit mettre cet élément
    //Tri par nature / code
    PosInTob := -1;
    for IndexBranche := 0 to TempTob.Detail.Count -1 do
    begin
      TriTempTob := TempTob.Detail[IndexBranche];
      stTempNATUREARCHI := TriTempTob.GetString('ARP_NATUREARCHI');
      if   (stTempNATUREARCHI = Nature)
        or (stTempNATUREARCHI = NatureToIcone(Nature)) then
      begin
        stTempCODE := TriTempTob.GetString('ARP_CODE');
        if (stTempCODE < Code) then
        begin
          PosInTob := TriTempTob.GetIndex;
          continue;
        end else if (stTempCODE = Code) then
        begin
          stTempDossier := TriTempTob.GetString('ARP_NODOSSIER');
          stTempPredefini := TriTempTob.GetString('ARP_PREDEFINI');
          if   (    (stTempDossier = Dossier)
                and (stTempPredefini = Predefini) )
            or (Nature = 'TDY') then
          begin
            { Déjà présent }
            exit;
          end else begin
            if stTempDossier < Dossier then
            begin
              PosInTob := TriTempTob.GetIndex;
              continue;
            end else begin
              if stTempDossier > Dossier then
              begin
                break;
              end else begin
                if stTempPredefini < Predefini then
                begin
                  PosInTob := TriTempTob.GetIndex;
                  continue;
                end else break;
              end;
            end;
          end;
        end else break;
      end else continue;
    end;
//    { On vérifie qu'on ai pas déjà analyser l'élément }
//    if Nature = 'TDY' then
//    begin
//      RechTob := TempTob.FindFirst(['ARP_CODE'], [Code], False);
//      if Assigned(RechTob) then
//        exit;
//    end else begin
//      RechTob := TempTob.FindFirst(['ARP_CODE', 'ARP_PREDEFINI', 'ARP_NODOSSIER'], [Code, Predefini, Dossier], False);
//      if Assigned(RechTob) then
//        exit;
//    end;
//    TempTob := Tob.create('Analyse d''un élément ', TempTob, -1);
    TempTob := Tob.create('Analyse d''un élément ', TempTob, PosInTob + 1);
    //Fin PT6
    TempTob.AddChampSupValeur('ARP_NATURE', Nature);
    TempTob.AddChampSupValeur('ARP_NATUREARCHI', Nature);
    TempTob.AddChampSupValeur('ARP_CODE', Code);
    if Nature = 'VAR' then
      TempTob.AddChampSupValeur('ARP_LIBELLE', RechDom('PGVARIABLE', Code, False))
    else if Nature = 'REM' then
      TempTob.AddChampSupValeur('ARP_LIBELLE', RechDom('PGREMUNERATION', Code, False))
    else if Nature = 'COT' then
      TempTob.AddChampSupValeur('ARP_LIBELLE', RechDom('PGCOTISATION', Code, False))
    else if Nature = 'CUM' then
      TempTob.AddChampSupValeur('ARP_LIBELLE', RechDom('PGCUMULPAIE', Code, False))
    else if Nature = 'ELT' then
      TempTob.AddChampSupValeur('ARP_LIBELLE', RechDom('PGELEMENTNAT', Code, False))
    else if Nature = 'TDY' then  //PT4
      TempTob.AddChampSupValeur('ARP_LIBELLE', RechDom('PGTABLEDIM', Code, False))
    else if Nature = 'PRO' then
      TempTob.AddChampSupValeur('ARP_LIBELLE', RechDom('PGPROFIL', Code, False));
    TempTob.AddChampSupValeur('ARP_PREDEFINI', Predefini);
    TempTob.AddChampSupValeur('ARP_NODOSSIER', Dossier);
    if Nature <> 'TDY' then
    begin
      LibDossier := RechDom('YYDOSSIER', Dossier, False);  //PT7
      if LibDossier = 'Error' then LibDossier := '';  //PT7
      TempTob.AddChampSupValeur('ARP_ORIGINE', RechDom('YYPREDEFINI', Predefini, False)+' '+LibDossier); //PT7
    end else
      TempTob.AddChampSupValeur('ARP_ORIGINE', '');  //Pour les tables dynamiques, on ne peut pas savoir quel sera la table utilisée si on a pas le contexte
    //Debut PT2
    RubDispatchConnecteur := RubriqueDispatcher.GetConnecteur(Nature);
    if Assigned(RubDispatchConnecteur) then
    begin
      //PT5
//      TpTob := RubDispatchConnecteur.TobMereRubriques.FindFirst([RubDispatchConnecteur.ChampsCode, RubDispatchConnecteur.Prefixe+'_PREDEFINI', RubDispatchConnecteur.Prefixe+'_NODOSSIER'], [Code, Predefini, Dossier], False);
      TpTob := RubDispatchConnecteur.TobMereRubriques.FindFirst([RubDispatchConnecteur.Prefixe+'_NODOSSIER'], [Dossier], False);
      if Assigned(TpTob) then
        TpTob := TpTob.FindFirst([RubDispatchConnecteur.ChampsCode, RubDispatchConnecteur.Prefixe+'_PREDEFINI'], [Code, Predefini], False);
      if Assigned(TpTob) then
      begin
        if RubDispatchConnecteur.TabletteType <> '' then
          TypeElt := RechDom(RubDispatchConnecteur.TabletteType, TpTob.GetString(RubDispatchConnecteur.ChampsType), False)
        else
          TypeElt := TpTob.GetString(RubDispatchConnecteur.ChampsType);
        TempTob.AddChampSupValeur('ARP_TYPENATUREARCHI', TypeElt);
      end else
        TempTob.AddChampSupValeur('ARP_TYPENATUREARCHI', '');
    end else
      TempTob.AddChampSupValeur('ARP_TYPENATUREARCHI', '');
    //Fin PT2
    TempTob.AddChampSupValeur('ARP_MESSAGE', '('+Code+') '+TempTob.GetString('ARP_LIBELLE'));
    result := TempTob;
  end;
end;

procedure TOF_ANARCHIREGLPAIE.AnalyseInversee;
begin
  RubriqueDispatcher.DispatchRubriques(LectureType);
end;

Initialization
  registerclasses ( [ TOF_ANARCHIREGLPAIE ] ) ;
end.
