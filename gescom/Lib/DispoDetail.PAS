{***********UNITE*************************************************
Auteur  ...... : Jean-Luc Sauzet
Créé le ...... : 23/04/2003
Modifié le ... :   /  /
Description .. : Fonctions & procédure métier
Mots clefs ... : DISPODETAIL
*****************************************************************}
unit DispoDetail;

interface

{$IFDEF STK}
uses
  Forms,
  Windows,
  Classes,
  SysUtils,
  Hctrls,
  Hent1,
  uTob,
  uTom,
  Dialogs,
  Menus,
  Graphics,
  SaisieList,
  uTableFiltre,
  messages,
  Controls,
  Hmsgbox,

  {$IFNDEF EAGLCLIENT}
    DBtables,
    Db,
    FE_Main,
  {$ELSE}
    MainEAGL,
  {$ENDIF}

  Dispo,
  wAction
  ;

Type tCleGQD =  Record
                  Article       : string;
                  Depot         : string;
                  StatutDispo   : string;
                  StatutFlux    : String;
                  Emplacement   : string;
                  LotInterne    : string;
                  SerieInterne  : string;
                  RefProprio    : string;
                  IndiceArticle : string;
                  Marque        : string;
                  ChoixQualite  : string;
                  RefAffectation: string;
                end;

function CreateGQD(CleGQD: tCleGQD): boolean;
function AddToGQD(CleGQD: tCleGQD; Physique: Double): boolean;

{ Call }
procedure CallMulGQD(CleGQ: TCleGQ; Flux: String);
procedure CallFicGQD(Article, Depot: string; Identifiant: integer; Flux: string; Action: String; UniteSto: String =''; UniteVte: String =''; UniteAch: String =''; UnitePro: String =''; UniteCon: String ='');
function CallFicPickingStk(CleGQ: tCleGQ; QDemSais: double; Flux: string; Contexte: String; TobDispo, TobServi: tob; FormCaption : string = ''; DispoCaption: string = ''; ServiCaption: string = ''; Info: string = ''): boolean;

{ Where }
function WhereGQD(CleGQD: tCleGQD): string; overload;
function WhereGQD(CleGQ: tCleGQ): string; overload;

{ Get }
function GetTobDispo(CleGQ: tCleGQ; Contexte: string; Var TobDispo: Tob): boolean;
function GetTobServi(Mode: string; QDemStoc: double; TobDispo: Tob; var TobServi: Tob): boolean;
function GetFieldFromGQD(FieldName: string; CleGQD: tCleGQD): Variant;
function GetRang(TobSource, TobTarget: tob): integer;

{ AddTo }
function AddToTob(TobSource, TobTarget: Tob; Physique: double): double;

{ Mouvements Manuel }
procedure EntreeGQD(CleGQ: tCleGQ; Flux: string);
procedure SortieGQD(CleGQ: tCleGQ; Flux: string);

const
	TableName        = 'DISPODETAIL';
  Wpa_CodeParam    = 'WPA_ENTREEGQ';
  Wpa_CodeArticle  = 'WPA_VARCHAR01';
  Wpa_Depot        = 'WPA_COMBO01';
  Wpa_StatutDispo  = 'WPA_COMBO04';
  Wpa_StatutFlux   = 'WPA_COMBO05';
  Wpa_Emplacement  = 'WPA_VARCHAR02';
  Wpa_SerieInterne = 'WPA_VARCHAR05';
  Wpa_Physique     = 'WPA_DOUBLE01';
  Wpa_Unite        = 'WPA_COMBO02';
  Wpa_BlocNote     = 'WPA_BLOCNOTE';

{$ENDIF}


implementation

{$IFDEF STK}

Uses
  EntGP,
  StkMouvement,
  wParam,
  UtilArticle,
  wCommuns,
  Math,
  uPickingStk,
  wConversionUnite
  ;

{***********A.G.L.***********************************************
Auteur  ...... : Jean-Luc Sauzet
Créé le ...... : 24/04/2003
Modifié le ... :   /  /
Description .. : Création d'un nouvel enregistrement GQD
Mots clefs ... :
*****************************************************************}
function CreateGQD(CleGQD: tCleGQD): boolean;
var
  sql: string;
  TobGQD: Tob;
  TomGQD: Tom;
begin
  { Test de l'existance }
  sql := 'SELECT 0'
       + ' FROM ' + TableName
       + ' WHERE ' + WhereGQD(CleGQD);
  if not existeSql(sql) then
  begin
    tobGQD := Tob.Create(TableName, nil, -1);
    tomGQD := CreateTOM(TableName, nil, false, true);
    try
      { NewRecord }
      tomGQD.InitTOB(tobGQD);

      { Data }
      TobGQD.P('GQD_ARTICLE',         CleGQD.Article);
      TobGQD.P('GQD_DEPOT',           CleGQD.Depot);
      TobGQD.P('GQD_STATUTDISPO',     CleGQD.StatutDispo);
      TobGQD.P('GQD_STATUTFLUX',      CleGQD.StatutFlux);
      TobGQD.P('GQD_EMPLACEMENT',     CleGQD.Emplacement);
      TobGQD.P('GQD_LOTINTERNE',      CleGQD.LotInterne);
      TobGQD.P('GQD_SERIEINTERNE',    CleGQD.SerieInterne);
      TobGQD.P('GQD_REFPROPRIO',      CleGQD.RefProprio);
      TobGQD.P('GQD_INDICEARTICLE',   CleGQD.IndiceArticle);
      TobGQD.P('GQD_MARQUE',          CleGQD.Marque);
      TobGQD.P('GQD_CHOIXQUALITE',    CleGQD.ChoixQualite);
      TobGQD.P('GQD_REFAFFECTATION',  CleGQD.RefAffectation);

      { UpDateRecord }
      tobGQD.AddChampSupValeur('IKC', 'C');
      Result := tomGQD.VerifTOB(tobGQD) and tobGQD.InsertDb(nil);
    finally
      tomGQD.free;
      tobGQD.free;
    end;
  end
  else
    Result := true;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Jean-Luc Sauzet
Créé le ...... : 07/05/2003
Modifié le ... :   /  /
Description .. : Ajoute la quantité à la cle
Mots clefs ... :
*****************************************************************}
function AddToGQD(CleGQD: tCleGQD; Physique: Double): boolean;
var
	Sql : string;
begin
  Sql := 'UPDATE DISPODETAIL'
	  	 + ' SET GQD_PHYSIQUE=GQD_PHYSIQUE+' + StrFPoint(Physique)
		   + ' WHERE ' + WhereGQD(CleGQD)
			 ;
  Result := (ExecuteSQL(Sql) = 1) or (CreateGQD(CleGQD) and AddToGQD(CleGQD, Physique));
end;

{***********A.G.L.***********************************************
Auteur  ...... : Jean-Luc Sauzet
Créé le ...... : 07/05/2003
Modifié le ... :   /  /
Description .. : Renvoie un Where GQD
Mots clefs ... :
*****************************************************************}
function WhereGQD(CleGQD: tCleGQD): string; overload;
begin
  Result := '('
          + '     GQD_ARTICLE="'        + CleGQD.Article + '"'
          + ' AND GQD_DEPOT="'          + CleGQD.Depot + '"'
          + ' AND GQD_STATUTDISPO="'    + CleGQD.StatutDispo + '"'
          + ' AND GQD_STATUTFLUX="'     + CleGQD.StatutFlux + '"'
          + ' AND GQD_EMPLACEMENT="'    + CleGQD.Emplacement + '"'
          + ' AND GQD_LOTINTERNE="'     + CleGQD.LotInterne + '"'
          + ' AND GQD_SERIEINTERNE="'   + CleGQD.SerieInterne + '"'
          + ' AND GQD_REFPROPRIO="'     + CleGQD.RefProprio + '"'
          + ' AND GQD_INDICEARTICLE="'  + CleGQD.IndiceArticle + '"'
          + ' AND GQD_MARQUE="'         + CleGQD.Marque + '"'
          + ' AND GQD_CHOIXQUALITE="'   + CleGQD.ChoixQualite + '"'
          + ' AND GQD_REFAFFECTATION="' + CleGQD.RefAffectation + '"'
          + ')'
end;

function WhereGQD(CleGQ: tCleGQ): string; overload;
begin
  Result := '('
          + '     GQD_ARTICLE="'        + CleGQ.Article + '"'
          + ' AND GQD_DEPOT="'          + CleGQ.Depot + '"'
          + ')'
end;


{***********A.G.L.***********************************************
Auteur  ...... : Jean-Luc Sauzet
Créé le ...... : 19/05/2003
Modifié le ... :   /  /
Description .. : Renvoie une tob contenant les lignes de GQD disponibles
Suite ........ : pour la demande.
Mots clefs ... :
*****************************************************************}
function GetTobDispo(CleGQ: tCleGQ; Contexte: string; Var TobDispo: Tob): boolean;
var
  sql : string;

  function GetWhere: string;
  var
    QualifMvt     : string;
    RefAffectation: string;
  begin
    QualifMvt := GetArgumentValue(Contexte, 'QUALIFMVT');
    RefAffectation := GetArgumentValue(Contexte, 'REFAFFECTATION');

    Result := WhereGQD(CleGQ);

    if QualifMvt = 'SPR' then
    begin
      if V_Pgi.SAV then
      begin
        if RefAffectation = '' then
          PgiError('Attention, il faut passer RefAffectation dans l''argument de Dispo.GetTobDispo', 'Dispo.GetTobDispo');
      end;
      Result := Result
              + ' AND GQD_STATUTDISPO <> "BLQ"'
              + ' AND GQD_STATUTFLUX IN ("STD", "CON")'
              + ' AND (GQD_REFAFFECTATION="" OR GQD_REFAFFECTATION="' + RefAffectation + '" )'
    end
    else if QualifMvt = 'SMA' then 
    begin
      Result := Result
              + ' AND GQD_STATUTDISPO <> "BLQ"'
              + ' AND GQD_STATUTFLUX IN ("STD")'
              + ' AND GQD_REFAFFECTATION=""'
    end
    else if QualifMvt = 'CSD' then
    begin

    end
    else if V_Pgi.SAV then
    begin
      PgiError('QualifMvt non défini dans le calcul de la TobDispo', 'Dispo.GetTobDispo');
    end;
  end;

begin
  sql := 'SELECT *'
       + ' FROM ' + TableName
       + ' WHERE (' + GetWhere + ')'
       + ' AND GQD_PHYSIQUE > 0'
       + ' ORDER BY GQD_REFAFFECTATION DESC, GQD_DATEENTREELOT'
       ;
  Result := wLoadTobFromSql(TableName, Sql, TobDispo);
end;

{***********A.G.L.***********************************************
Auteur  ...... : Jean-Luc Sauzet
Créé le ...... : 07/08/2003
Modifié le ... :   /  /
Description .. : Renvoie la TobServi selon le mode demandé
Mots clefs ... :
*****************************************************************}
function GetTobServi(Mode: string; QDemStoc: double; TobDispo: Tob; var TobServi: Tob): boolean;
var
  iDispo  : integer;
begin
  Result := false;
  if QDemStoc <= 0 then exit;

  if Mode = 'FIFO' then
  begin
    Result := true;

    { Tri de la TobDispo à l'envers car ja vais parcourir la tob de la fin }
    TobDispo.Detail.Sort('GQD_REFAFFECTATION;-GQD_DATEENTREELOT');

    for iDispo := TobDispo.Detail.Count-1 downto 0 do
    begin
      QDemStoc := QDemStoc - AddToTob(TobDispo.Detail[iDispo], TobServi, Min(QDemStoc, TobDispo.Detail[iDispo].G('GQD_PHYSIQUE')));
      if QDemStoc <= 0 then break;
    end;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Jean-Luc Sauzet
Créé le ...... : 13/05/2003
Modifié le ... :   /  /
Description .. : Entrée exeptionnelle
Mots clefs ... :
*****************************************************************}
procedure EntreeGQD(CleGQ: tCleGQ; Flux: string);
  function GetArgument: string;
  begin
    Result := 'ARTICLE=' + CleGQ.Article
            + ';DEPOT=' + CleGQ.Depot
            + ';FLUX=' + Flux
  end;
begin
  AGLLanceFiche('GC', 'GSMPHYSIQUE_FIC', '', '', 'MONOFICHE;ACTION=CREATION;' + GetArgument);
end;

procedure SortieGQD(CleGQ: tCleGQ; Flux: string);
var
  TobDispo, TobServi: Tob;

  function GetContexte: string;
  begin
    Result := 'QUALIFMVT=SMA'
  end;

  function CreateGSM(TobServi: Tob; IdAction: integer): Double;
  var
    ActionGSM : tActionGSM;

    function GetTobData(var TobData: Tob): double;
    var
      iServi, iChamp     : integer;
      FieldName, Prefixe : string;
      t                  : tob;
    begin
      Result := 0;
      Prefixe := TobServi.NomTable;

      for iServi := 0 to TobServi.Detail.Count-1 do
      begin
        t := Tob.Create('DETAIL', TobData, -1);

        { Champs de GSM }
        t.AddChampSupValeur('GSM_QUALIFMVT', 'SMA');
        t.AddChampSupValeur('GSM_ETATMVT'  , 'SOL');

        { Champs Réels }
        for iChamp := 1 to TobServi.Detail[iServi].NbChamps do
        begin
          FieldName := StringReplace(TobServi.Detail[iServi].GetNomChamp(iChamp), Prefixe + '_', 'GSM_', [rfReplaceAll]);
          t.AddChampSupValeur(FieldName, TobServi.Detail[iServi].G(TobServi.Detail[iServi].GetNomChamp(iChamp)));
        end;

        { Champs virtuels }
        for iChamp := 1000 to (1000 + TobServi.Detail[iServi].ChampsSup.Count - 1) do
        begin
          FieldName := StringReplace(TobServi.Detail[iServi].GetNomChamp(iChamp), Prefixe + '_', 'GSM_', [rfReplaceAll]);
          t.AddChampSupValeur(FieldName, TobServi.Detail[iServi].G(TobServi.Detail[iServi].GetNomChamp(iChamp)));
        end;

        { Sortie: Négatif }
        t.P('GSM_PHYSIQUE', -t.G('GSM_PHYSIQUE'));

        { Action }
        t.AddChampSupValeur('GSM_IDACTION', IdAction);

        { Cumul }
        Result := Result + TobServi.Detail[iServi].G(Prefixe + '_PHYSIQUE');
      end;
    end;

  begin
    Result := 0;

    { Exit }
    if TobServi.Detail.Count <= 0 then exit;

    ActionGSM := tActionGSM.Create;
    ActionGSM.TobData := Tob.Create('GSM', nil, -1);
    try
      Result := GetTobData(ActionGSM.TobData);

      if ActionGSM.TobData.detail.Count > 0 then
      begin
        V_PGI.IoError := Transactions(ActionGSM.DoCreate, 1);
        if V_PGI.IoError = oeOk then
        begin

        end
      end;
    finally
      ActionGSM.TobData.free;
      ActionGSM.free;
    end;
  end;

begin
  TobDispo     := Tob.Create('GQD', nil, -1);
  TobServi     := Tob.Create('GQD', nil, -1);
  try
    GetTobDispo(CleGQ, GetContexte, TobDispo);
    if CallFicPickingStk(CleGQ, 0, Flux, '', TobDispo, TobServi, TraduireMemoire('Sortie exeptionnelle'), TraduireMemoire('Stock disponible'), TraduireMemoire('Stock à sortir')) then
      CreateGSM(TobServi, 0);
  finally
    TobDispo.free;
    TobServi.free;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Jean-Luc Sauzet
Créé le ...... : 13/05/2003
Modifié le ... :   /  /
Description .. : Renvoi un champ de GQ
Mots clefs ... :
*****************************************************************}
function GetFieldFromGQD(FieldName: string; CleGQD: tCleGQD): Variant;
begin
  Result := wGetSqlFieldValue(FieldName, TableName, WhereGQD(CleGQD));
end;

{***********A.G.L.***********************************************
Auteur  ...... : Thibaut SUBLET
Créé le ...... : 22/05/2003
Modifié le ... :   /  /
Description .. : Call du mul du détail du dispo
Mots clefs ... :
*****************************************************************}
procedure CallMulGQD(CleGQ: TCleGQ; Flux: String);

  function GetRange: string;
  begin
    Result := 'GQD_ARTICLE='+ CleGQ.Article
            + ';GQD_DEPOT=' + CleGQ.Depot
            + ';GQD_CODEARTICLE=' + wGetCodeArticleFromArticle(CleGQ.Article)
  end;

  function GetArgument: string;
  begin
    Result := 'FLUX=' + Flux;
  end;

begin
  wForceCumulsOnViewFields('GCDISPODETAIL', 'PHYSIQUESTO;PHYSIQUEVTE;PHYSIQUEACH;PHYSIQUEPRO;PHYSIQUECON;');
  AglLanceFiche('GC', 'GCDISPODETAIL_MUL', GetRange, '', 'ACTION=MODIFICATION;' + GetArgument);
end;

{***********A.G.L.***********************************************
Auteur  ...... : Dominique Sclavopoulos
Créé le ...... : 18/06/2003
Modifié le ... :   /  /
Description .. : Appel de la fiche DISPODETAIL
Mots clefs ... :
*****************************************************************}
procedure CallFicGQD(Article, Depot: string; Identifiant: integer; Flux: string; Action: String; UniteSto, UniteVte, UniteAch, UnitePro, UniteCon: String);
begin
  AGLLanceFiche('GC', 'GCDISPODETAIL_FIC', '', Article + ';' + Depot + ';' + IntToStr(Identifiant), 'MONOFICHE;ACTION=' + Action + ';FLUX=' + Flux + ';UNITESTO=' + UniteSto + ';UNITEVTE=' + UniteVte + ';UNITEACH=' + UniteAch + ';UNITEPRO=' + UnitePro + ';UNITECON=' + UniteCon);
end;

(*
{***********A.G.L.***********************************************
Auteur  ...... : Thibaut SUBLET
Créé le ...... : 05/06/2003
Modifié le ... :   /  /
Description .. : Recalcul du détail du physique
Mots clefs ... : DISPODETAIL;STOCK;RECALCUL
*****************************************************************}
function RecalculGQD(TypeAction: wTypeAction; Action: Integer; Argument: string): boolean;
var
  Where: String;

  { Control }
  function RecalculIsNeeded: Boolean;
  var
    Q: TQuery;
    sSql: String;
  const
    ControlFieldName = 'CONTROL_GQD';
  begin
    sSql := 'SELECT (SUM(GSM_PHYSIQUE) - GQD_PHYSIQUE) as ' + ControlFieldName
            + ' FROM STKMOUVEMENT'
            + ' LEFT JOIN DISPODETAIL ON ('
            + ' GSM_ARTICLE=GQD_ARTICLE'
            + ' AND GSM_DEPOT=GQD_DEPOT'
            + ' AND GSM_STATUTDISPO=GQD_STATUTDISPO'
            + ' AND GSM_STATUTFLUX=GQD_STATUTFLUX'
            + ' AND GSM_EMPLACEMENT=GQD_EMPLACEMENT'
            + ' AND GSM_LOTINTERNE=GQD_LOTINTERNE'
            + ' AND GSM_SERIEINTERNE=GQD_SERIEINTERNE'
            + ' AND GSM_REFPROPRIO=GQD_REFPROPRIO'
            + ' AND GSM_INDICEARTICLE=GQD_INDICEARTICLE'
            + ' AND GSM_MARQUE=GQD_MARQUE'
            + ' AND GSM_CHOIXQUALITE=GQD_CHOIXQUALITE'
            + ' AND GSM_REFAFFECTATION=GQD_REFAFFECTATION)'
            + ' WHERE ' + Where
            + ' AND GSM_STKTYPEMVT="PHY"'
            + ' GROUP BY GQD_PHYSIQUE'
            ;
    Q := OpenSQL(sSql, True, -1);
    try
      Result := Q.FindField(ControlFieldName).AsFloat <> 0;
    finally
      Ferme(Q);
    end;
  end;

  { Suppression du détail du physique }
  function DeleteGQD: Boolean;
  begin
    Result := wDeleteTable('DISPODETAIL', Where);
  end;

  { Recalcul }
  function RecalculIt: Boolean;
  var
    sSql: String;
  begin
    sSql := 'UPDATE DISPODETAIL'
          + ' SET GQD_PHYSIQUE=(SELECT SUM(GSM_PHYSIQUE) FROM STKMOUVEMENT WHERE ' + StringReplace(Where, 'GQD_', 'GSM_', [rfIgnoreCase, rfReplaceAll]) + ' AND GSM_STKTYPEMVT="PHY")'
          + ' WHERE ' + Where
          ;
    Result := ExecuteSQL(sSql) > 0;
  end;

begin
  Result := False;
  Where := GetArgumentValue(Argument, 'WHERE');

  PutToDebugLog(GetAction(TypeAction), True);

  if RecalculIsNeeded then
  begin
    //if DeleteGQD then
      Result := RecalculIt;
  end
  else if V_PGI.SAV and not GetArgumentBoolean(Argument, 'MULTISELECTED') then
    PGIInfo(TraduireMemoire('Le recalcul n''est pas nécessaire.'));

  PutToDebugLog(GetAction(TypeAction), False, 'Fin du recalcul : Recalcul ' + iif(Result, 'effectué.', 'non nécessaire.'));
end;
*)

{***********A.G.L.***********************************************
Auteur  ...... : Jean-Luc Sauzet
Créé le ...... : 01/08/2003
Modifié le ... :   /  /
Description .. : Appel de la fiche Picking stock
Mots clefs ... :
*****************************************************************}
function CallFicPickingStk(CleGQ: tCleGQ; QDemSais: double; Flux: string; Contexte: String; TobDispo, TobServi: tob; FormCaption : string = ''; DispoCaption: string = ''; ServiCaption: string = ''; Info: string = ''): boolean;
var
  PickingStk: TfPickingStk;
begin
  Result := false;
  PickingStk := TfPickingStk.Create(Application.MainForm);
  try
    { Informations }
    PickingStk.Caption := iif(FormCaption = '', TraduireMemoire('Fiche picking'), FormCaption);
    PickingStk.PnTitreDispo.Caption := iif(DispoCaption = '', TraduireMemoire('Stock disponible'), DispoCaption);
    PickingStk.PnTitreServi.Caption := iif(ServiCaption = '', TraduireMemoire('Stock servi')     , ServiCaption);
    PickingStk.mmInfo.Text := Info;

    PickingStk.TobDispo := Tob.Create('GQD', nil, -1);
    PickingStk.TobServi := Tob.Create('GQD', nil, -1);
    try
      { Données }
      PickingStk.TobDispo.Dupliquer(TobDispo, true, true);
      PickingStk.TobServi.Dupliquer(TobServi, true, true);
      PickingStk.CleGQ := CleGQ;
      PickingStk.Contexte := Contexte;
      PickingStk.Flux := Flux;
      PickingStk.Mesure := GetMesureFromFlux(CleGQ.Article, Flux);
      PickingStk.Quotite := GetQuotiteMEA(PickingStk.Mesure);

      { Quantité }
      PickingStk.NeQDemandee.Value := QDemSais;

      { Fiche }
      PickingStk.ShowModal;

      { Résultat }
      if PickingStk.ModalResult = mrOk then
      begin
        Result := true;
        TobDispo.Dupliquer(PickingStk.TobDispo, true, true);
        TobServi.Dupliquer(PickingStk.TobServi, true, true);
      end;
    finally
      PickingStk.TobDispo.free;
      PickingStk.TobServi.free;
    end;
  finally
    PickingStk.free;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Jean-Luc Sauzet
Créé le ...... : 18/08/2003
Modifié le ... :   /  /
Description .. : Passe la quantité physique de la TobSource vers la
Suite ........ : TobTarget.
Suite ........ : Si la TobSOuce passe à 0 alors on vire la ligne.
Suite ........ : Si la TobTarget n'existe pas, alors on la crée
Mots clefs ... :
*****************************************************************}
function AddToTob(TobSource, TobTarget: Tob; Physique: double): double;
var
  iTarget: integer;

  procedure CreateTobTarget;
  var
    iChamp: integer;
    t     : Tob;
  begin
    t := Tob.Create(TableName, TobTarget, -1);

    for iChamp := 1 to TobSource.NbChamps do
      t.P(TobSource.GetNomChamp(iChamp), TobSource.G(TobSource.GetNomChamp(iChamp)));

    { Forcer la quantité à 0 }
    t.P('GQD_PHYSIQUE', 0);

    { Trier la TobTarget }
    TobTarget.Detail.Sort('-GQD_REFAFFECTATION;GQD_DATEENTREELOT');
  end;
begin
  Result := 0;

  Physique := Min(Physique, TobSource.G('GQD_PHYSIQUE'));
  if Physique <= 0 then exit;

  iTarget := GetRang(TobSource, TobTarget);
  if iTarget >= 0 then
  begin
    TobSource.P('GQD_PHYSIQUE', TobSource.G('GQD_PHYSIQUE') - Physique);
    TobTarget.Detail[iTarget].P('GQD_PHYSIQUE', TobTarget.Detail[iTarget].G('GQD_PHYSIQUE') + Physique);
    Result := Physique;

    { On vire l'enregistrement si vide }
    if TobSource.G('GQD_PHYSIQUE') <= 0 then
      TobSource.free;
  end
  else
  begin
    CreateTobTarget;
    Result := AddToTob(TobSource, TobTarget, Physique);
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Jean-Luc Sauzet
Créé le ...... : 07/08/2003
Modifié le ... :   /  /
Description .. : Renvoie le Rang de la TobTarget correspondant à la clef
Suite ........ : de la TobSource
Mots clefs ... :
*****************************************************************}
function GetRang(TobSource, TobTarget: tob): integer;
var
  iTarget: integer;
begin
  Result := -1;

  for iTarget := 0 to TobTarget.Detail.Count-1 do
  begin
    if    (TobSource.G('GQD_ARTICLE')        = TobTarget.Detail[iTarget].G('GQD_ARTICLE'))
      and (TobSource.G('GQD_DEPOT')          = TobTarget.Detail[iTarget].G('GQD_DEPOT'))
      and (TobSource.G('GQD_STATUTDISPO')    = TobTarget.Detail[iTarget].G('GQD_STATUTDISPO'))
      and (TobSource.G('GQD_STATUTFLUX')     = TobTarget.Detail[iTarget].G('GQD_STATUTFLUX'))
      and (TobSource.G('GQD_EMPLACEMENT')    = TobTarget.Detail[iTarget].G('GQD_EMPLACEMENT'))
      and (TobSource.G('GQD_LOTINTERNE')     = TobTarget.Detail[iTarget].G('GQD_LOTINTERNE'))
      and (TobSource.G('GQD_SERIEINTERNE')   = TobTarget.Detail[iTarget].G('GQD_SERIEINTERNE'))
      and (TobSource.G('GQD_REFPROPRIO')     = TobTarget.Detail[iTarget].G('GQD_REFPROPRIO'))
      and (TobSource.G('GQD_INDICEARTICLE')  = TobTarget.Detail[iTarget].G('GQD_INDICEARTICLE'))
      and (TobSource.G('GQD_MARQUE')         = TobTarget.Detail[iTarget].G('GQD_MARQUE'))
      and (TobSource.G('GQD_CHOIXQUALITE')   = TobTarget.Detail[iTarget].G('GQD_CHOIXQUALITE'))
      and (TobSource.G('GQD_REFAFFECTATION') = TobTarget.Detail[iTarget].G('GQD_REFAFFECTATION')) then
    begin
      Result := iTarget;
      break;
    end;
  end;
end;

{$ENDIF}
end.
