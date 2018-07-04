{{***********UNITE*************************************************
Auteur  ...... : Jean-Luc Sauzet
Créé le ...... : 25/05/2003
Modifié le ... :   /  /    
Description .. : Recalcul les compteurs de Dispo basé sur les mouvements 
Suite ........ : de réservation
Mots clefs ... : 
*****************************************************************}
unit Dispo;

interface

uses
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
  Hmsgbox,

  {$IFNDEF EAGLCLIENT}
    DBtables,
    Db,
    FE_Main,
  {$ELSE}
    MainEAGL,
  {$ENDIF}
  {$IFDEF GPAO}
    wAction,
  {$ENDIF GPAO}  

   wCommuns
    ;

type
  tCleGQ = record
    Article     : string;
    Depot       : String;
    Cloture     : Boolean;
    DateCloture : TDateTime;
  end;

function CreateGQ(CleGQ: tCleGQ): boolean;
function AddToGQ(CleGQ: tCleGQ; FieldName: string; Physique: Double): boolean;
function RemoveFromGQ(CleGQ: tCleGQ; FieldName: string; Physique: Double): boolean;

{ Call }
procedure CallMulGQ(CleGQ: tCleGQ; Flux: string);
procedure CallFicGQ(CleGQ: TCleGQ; Flux: string; Action: String; UniteSto: String =''; UniteVte: String =''; UniteAch: String =''; UnitePro: String =''; UniteCon: String ='');

{ Get }
function GetFieldFromGQ(FieldName: string; CleGQ: tCleGQ): Variant;

{ Where }
function WhereGQ(CleGQ: tCleGQ): string;

{ Exists }
function ExistsGQ(CleGQ: TCleGQ; WithAlert: Boolean = False): Boolean;

{ Utilitaires }
{$IFDEF STK}
function RecalculCompteursGQ(TypeAction: wTypeAction; Action: integer; Argument: string): boolean;
{$ENDIF}

const
	TableName = 'DISPO';

implementation

Uses
  EntGP,
  UtilArticle,
{$IFDEF STK}
  StkMouvement,
  StkNature,
{$ENDIF}
  wParam,
  wJetons
  ;

{***********A.G.L.***********************************************
Auteur  ...... : Jean-Luc Sauzet
Créé le ...... : 24/04/2003
Modifié le ... :   /  /
Description .. : Ajoute la quantité à Dispo
Mots clefs ... :
*****************************************************************}
function AddToGQ(CleGQ: tCleGQ; FieldName: string; Physique: Double): boolean;
var
	Sql : string;
begin
  Sql := 'UPDATE DISPO'
	  	 + ' SET ' + FieldName + '=' + FieldName + '+' + StrFPoint(Physique)
		   + ' WHERE ' + WhereGQ(CleGQ)
			 ;
  Result := (ExecuteSQL(Sql) = 1) or (CreateGQ(CleGQ) and AddToGQ(CleGQ, FieldName, Physique));
end;

{***********A.G.L.***********************************************
Auteur  ...... : Jean-Luc Sauzet
Créé le ...... : 24/04/2003
Modifié le ... :   /  /
Description .. : Retire de Dispo la quantité
Mots clefs ... :
*****************************************************************}
function RemoveFromGQ(CleGQ: tCleGQ; FieldName: string; Physique: Double): boolean;
var
	Sql       : string;
begin
  Sql := 'UPDATE DISPO'
	  	 + ' SET ' + FieldName + '=' + FieldName + '-' + StrFPoint(Physique)
		   + ' WHERE ' + WhereGQ(CleGQ)
			 ;
  Result := (ExecuteSQL(Sql) = 1) or (CreateGQ(CleGQ) and RemoveFromGQ(CleGQ, FieldName, Physique));
end;

{***********A.G.L.***********************************************
Auteur  ...... : Jean-Luc Sauzet
Créé le ...... : 06/05/2003
Modifié le ... :   /  /    
Description .. : Création d'un enregistrement Dispo
Mots clefs ... : 
*****************************************************************}
function CreateGQ(CleGQ: tCleGQ): boolean;
var
  TobGQ: Tob;
  TomGQ: Tom;
begin
  tobGQ := Tob.Create(TableName, nil, -1);
  tomGQ := CreateTOM(TableName, nil, false, true);
  try
    { NewRecord }
    tomGQ.InitTOB(tobGQ);

    { Data }
    TobGQ.P('GQ_ARTICLE', CleGQ.Article);
    TobGQ.P('GQ_DEPOT', CleGQ.Depot);
    TobGQ.P('GQ_CLOTURE', BoolToStr(CleGQ.Cloture));
    TobGQ.P('GQ_DATECLOTURE', CleGQ.DateCloture);

    { UpDateRecord }
    TobGQ.AddChampSupValeur('IKC', 'C');
    Result := tomGQ.VerifTOB(tobGQ) and tobGQ.InsertDb(nil);
  finally
    tomGQ.free;
    tobGQ.free;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Jean-Luc Sauzet
Créé le ...... : 13/05/2003
Modifié le ... :   /  /    
Description .. : Appel de la fiche Mul GQ
Mots clefs ... : 
*****************************************************************}
procedure CallMulGQ(CleGQ: tCleGQ; Flux: string);
  function GetRange: string;
  begin
    Result := 'GQ_ARTICLE='+ CleGQ.Article
            + ';GQ_DEPOT=' + CleGQ.Depot
            + ';GA_CODEARTICLE=' + wGetCodeArticleFromArticle(CleGQ.Article)
            + ';GQ_CLOTURE=' + BoolToStr(CleGQ.Cloture)
            + ';GQ_DATECLOTURE=' + DateToStr(CleGQ.DateCloture)
  end;

  function GetArgument: string;
  begin
    Result := 'FLUX=' + Flux;
  end;
begin
{$IFDEF STK}
  wForceCumulsOnViewFields('GCDISPODEPART', 'PHYSIQUESTO;PHYSIQUEVTE;PHYSIQUEACH;PHYSIQUEPRO;PHYSIQUECON;'
                                          + 'RESERVESTO;RESERVEVTE;RESERVEACH;RESERVEPRO;RESERVECON;'
                                          + 'DISPOABSSTO;DISPOABSVTE;DISPOABSACH;DISPOABSPRO;DISPOABSCON;'
                                          + 'ATTENDUSTO;ATTENDUVTE;ATTENDUACH;ATTENDUPRO;ATTENDUCON;'
                                          + 'PROJABSSTO;PROJABSVTE;PROJABSACH;PROJABSPRO;PROJABSCON;');
  AglLanceFiche('GC', 'GCDISPODEPART_MUL', GetRange, '', 'ACTION=MODIFICATION;' + GetArgument);
{$ELSE}
  AglLanceFiche('W', 'WDISPOART_MUL', GetRange, '', 'ACTION=MODIFICATION;' + GetArgument);
{$ENDIF}
end;

{***********A.G.L.***********************************************
Auteur  ...... : Jean-Luc Sauzet
Créé le ...... : 13/05/2003
Modifié le ... :   /  /    
Description .. : Renvoi un champ de GQ
Mots clefs ... : 
*****************************************************************}
function GetFieldFromGQ(FieldName: string; CleGQ: tCleGQ): Variant;
begin
  Result := wGetSqlFieldValue(FieldName, TableName, WhereGQ(CleGQ));
end;

{***********A.G.L.***********************************************
Auteur  ...... : Jean-Luc Sauzet
Créé le ...... : 13/05/2003
Modifié le ... :   /  /    
Description .. : Construit un Where sur le GQ
Mots clefs ... :
*****************************************************************}
function WhereGQ(CleGQ: tCleGQ): string;
begin
  Result := 'GQ_ARTICLE="' + CleGQ.Article + '"'
          + ' AND GQ_DEPOT="' + CleGQ.Depot + '"'
          + ' AND GQ_CLOTURE="' + BoolToStr(CleGQ.Cloture) + '"'
          + ' AND GQ_DATECLOTURE="' + UsDateTime(CleGQ.DateCloture) +'"'
end;

procedure CallFicGQ(CleGQ: TCleGQ; Flux: string; Action: String; UniteSto, UniteVte, UniteAch, UnitePro, UniteCon: String);
var
  ReloadFic, Return: String;
  Unites: MyArrayValue;
begin
  Unites := nil;
  repeat
    Return := GetArgumentValue(AGLLanceFiche('GC', 'GCDISPODEPART_FIC', '', CleGQ.Article + ';' + CLeGQ.Depot + ';' + BoolToStr(CleGQ.Cloture) + ';' + DateToStr(CleGQ.DateCloture), 'MONOFICHE;ACTION=' + Action + ';FLUX=' + Flux + ';UNITESTO=' + UniteSto + ';UNITEVTE=' + UniteVte + ';UNITEACH=' + UniteAch + ';UNITEPRO=' + UnitePro + ';UNITECON=' + UniteCon), 'RELOADFIC');
    if Return <> '' then
    begin
      ReloadFic := StringReplace(Return, '~', ';', [rfReplaceAll, rfIgnoreCase]);
      CleGQ.Article := ReadTokenSt(ReloadFic);
      CleGQ.Depot := ReadTokenSt(ReloadFic);
      CleGQ.Cloture := StrToBool(ReadTokenSt(ReloadFic));
      CleGQ.DateCloture := StrToDate(ReadTokenSt(ReloadFic));
{$IFDEF GPAO}
      Unites := wGetFieldsFromGA(['GA_QUALIFUNITESTO', 'GA_UNITEQTEVTE', 'GA_UNITEQTEACH', 'GA_UNITEPROD', 'GA_UNITECONSO'], CleGQ.Article);
      UniteSto := Unites[0]; UniteVte := Unites[1]; UniteAch := Unites[2]; UnitePro := Unites[3]; UniteCon := Unites[4];
{$ELSE GPAO}
      UniteSto := wGetFieldFromGA('GA_QUALIFUNITESTO', CleGQ.Article);
      UniteVte := UniteSto; UniteAch := UniteSto; UnitePro := UniteSto; UniteCon := UniteSto;
{$ENDIF GPAO}
    end;
  until Return = '';
end;

{$IFDEF STK}
{***********A.G.L.***********************************************
Auteur  ...... : Jean-Luc Sauzet
Créé le ...... : 26/05/2003
Modifié le ... : 27/05/2003
Description .. : Recalcul dynamique des compteur Dispo
Suite ........ :
Mots clefs ... :
*****************************************************************}
function RecalculCompteursGQ(TypeAction: wTypeAction; Action: integer; Argument: string): boolean;
var
  iCSK, iGSN                : Integer;
  TobGQD, T                 : Tob;
  sql, WhereGQ, WherePhy,
  WherePre, sSql, WhereGSM,
  sField                    : String;
const
  ControlFieldName = 'CONTROL_GQD';
  CalculGQD_PHYSIQUE = 'PHYSIQUE_GQD';

  function GetCompteur: string;
  begin
    Result := VH_GP.TobCSK.Detail[iCSK].G('CO_ABREGE');
  end;

  function GetStkTypeMvt: string;
  begin
    Result := Vh_Gp.TobGSN.Detail[iGSN].G('GSN_STKTYPEMVT');
  end;

  function GetQualifMvt: string;
  begin
    Result := Vh_Gp.TobGSN.Detail[iGSN].G('GSN_QUALIFMVT');
  end;

  procedure MiseAZeroDuCompteur(Compteur: string);
  begin
    sql := 'UPDATE DISPO'
         + ' SET ' + Compteur + '=0'
         + iif(WhereGQ <> '', ' WHERE ' + WhereGQ, '')
         ;
    ExecuteSql(Sql);
  end;

  procedure MiseAZeroDesNull(Compteur: string);
  begin
    Sql := 'UPDATE DISPO'
         + ' SET ' + Compteur + '=0'
         + ' WHERE ' + Compteur + ' IS NULL'
         + iif(WhereGQ <> '', ' AND ' + WhereGQ, '')
         ;
    ExecuteSql(Sql);
  end;

  function GetWherePhy(QualifMvt: string): string;
  begin
    Result := '(GSM_STKTYPEMVT="PHY" AND GSM_QUALIFMVT="' + QualifMvt + '")';
  end;

  function GetWherePre(StkTypeMvt, QualifMvt: string): string;
  begin
    Result := '(GSM_STKTYPEMVT="' + StkTypeMvt + '" AND GSM_QUALIFMVT="' + QualifMvt + '")';
  end;

  function MakeWhereDelGQD(T: Tob): String;
  begin
    Result := '(GQD_ARTICLE="' + T.G('GSM_ARTICLE') + '"'
            + ' AND GQD_DEPOT="' + T.G('GSM_DEPOT') + '"'
            + ' AND GQD_STATUTDISPO="' + T.G('GSM_STATUTDISPO') + '"'
            + ' AND GQD_STATUTFLUX="' + T.G('GSM_STATUTFLUX') + '"'
            + ' AND GQD_EMPLACEMENT="' + T.G('GSM_EMPLACEMENT') + '"'
            + ' AND GQD_LOTINTERNE="' + T.G('GSM_LOTINTERNE') + '"'
            + ' AND GQD_SERIEINTERNE="' + T.G('GSM_SERIEINTERNE') + '"'
            + ' AND GQD_REFPROPRIO="' + T.G('GSM_REFPROPRIO') + '"'
            + ' AND GQD_INDICEARTICLE="' + T.G('GSM_INDICEARTICLE') + '"'
            + ' AND GQD_MARQUE="' + T.G('GSM_MARQUE') + '"'
            + ' AND GQD_CHOIXQUALITE="' + T.G('GSM_CHOIXQUALITE') + '"'
            + ' AND GQD_REFAFFECTATION="' + T.G('GSM_REFAFFECTATION') + '")'
  end;

  function GetWhereGSM: string;
  begin
    Result := StringReplace(StringReplace(WhereGQ, ' AND GQ_CLOTURE="' + wFalse + '" AND GQ_DATECLOTURE="' + USDATETIME(iDate1900) + '"', '', [rfIgnoreCase, rfReplaceAll]), 'GQ_', 'GSM_', [rfIgnoreCase, rfReplaceAll]);
  end;

  function GetWhereGQD: string;
  begin
    Result := StringReplace(WhereGSM, 'GSM_', 'GQD_', [rfIgnoreCase, rfReplaceAll]);
  end;

  { Traitement de l'intégrité entre GSM_PHYSIQUE & GQD_PHYSIQUE ****************** DEBUT ******************* }
  procedure UpdateGQDByGSM;
  var
    iTobSTKError  : Integer;
    WhereDelGQD   : String;
    OneGQDIsFalse : Boolean;
    TobSTKError   : Tob;
  begin

    { récupère le Where GQ et le transforme en Where GQD }
    WhereGSM := GetWhereGSM;
    { supprime les enregistrements de DISPODETAIL non présents dans STKMOUVEMENT }
    ExecuteSQL('DELETE FROM DISPODETAIL'
             + ' WHERE NOT EXISTS(SELECT NULL'
                               + ' FROM STKMOUVEMENT'
                               + ' WHERE GSM_ARTICLE=GQD_ARTICLE'
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
             + ' AND ' + GetWhereGQD);
    { Ramène le contrôle du physique entre GQD & GSM et le total de GSM }
    sSql := 'SELECT GSM_ARTICLE, GSM_CODEARTICLE, GSM_DEPOT, GSM_STATUTDISPO, GSM_STATUTFLUX, GSM_EMPLACEMENT'
          + ', GSM_LOTINTERNE, GSM_SERIEINTERNE, GSM_REFPROPRIO, GSM_INDICEARTICLE, GSM_MARQUE, GSM_CHOIXQUALITE'
          + ', GSM_REFAFFECTATION, ISNULL((SUM(GSM_PHYSIQUE) - GQD_PHYSIQUE), -1) AS ' + ControlFieldName
          + ', SUM(GSM_PHYSIQUE) AS ' + CalculGQD_PHYSIQUE
          + ' FROM STKMOUVEMENT'
          + ' LEFT JOIN DISPODETAIL'
          + ' ON (GSM_ARTICLE=GQD_ARTICLE'
          +     ' AND GSM_DEPOT=GQD_DEPOT'
          +     ' AND GSM_STATUTDISPO=GQD_STATUTDISPO'
          +     ' AND GSM_STATUTFLUX=GQD_STATUTFLUX'
          +     ' AND GSM_EMPLACEMENT=GQD_EMPLACEMENT'
          +     ' AND GSM_LOTINTERNE=GQD_LOTINTERNE'
          +     ' AND GSM_SERIEINTERNE=GQD_SERIEINTERNE'
          +     ' AND GSM_REFPROPRIO=GQD_REFPROPRIO'
          +     ' AND GSM_INDICEARTICLE=GQD_INDICEARTICLE'
          +     ' AND GSM_MARQUE=GQD_MARQUE'
          +     ' AND GSM_CHOIXQUALITE=GQD_CHOIXQUALITE'
          +     ' AND GSM_REFAFFECTATION=GQD_REFAFFECTATION)'
          +     ' WHERE ' + WhereGSM
          + ' AND GSM_STKTYPEMVT="PHY"'
          + ' GROUP BY GSM_ARTICLE, GSM_CODEARTICLE, GSM_DEPOT, GSM_STATUTDISPO, GSM_STATUTFLUX, GSM_EMPLACEMENT, GSM_LOTINTERNE, GSM_SERIEINTERNE, GSM_REFPROPRIO, GSM_INDICEARTICLE, GSM_MARQUE, GSM_CHOIXQUALITE, GSM_REFAFFECTATION, GQD_PHYSIQUE'
          + ' HAVING ((SUM(GSM_PHYSIQUE) - GQD_PHYSIQUE) <> 0) OR ((SUM(GSM_PHYSIQUE) - GQD_PHYSIQUE) is null)'
          ;
    TobSTKError := Tob.Create('STKMOUVEMENT', nil, -1);
    try
      if wLoadTobFromSQL('STKMOUVEMENT', sSql, TobSTKError) then
      begin
        TobGQD := Tob.Create('DISPODETAIL_VIRTUELLE', nil, -1);
        try
          { on boucle sur les valeurs ramenées par la requête... }
          OneGQDIsFalse := False;
          for iTobSTKError := 0 to Pred(TobSTKError.Detail.Count) do
          begin
            { ...si l'on a besoin de recalculer... }
            if TobSTKError.Detail[iTobSTKError].G(ControlFieldName) <> 0 then
            begin
              OneGQDIsFalse := True;
              { ...composition du where afin de supprimer les enregistrements de GQD erronés (déphasage avec GSM) }
              WhereDelGQD := WhereDelGQD + iif(WhereDelGQD = '', '', ' OR ') + MakeWhereDelGQD(TobSTKError.Detail[iTobSTKError]);
              { on recréé les données de GQD en mémoire }
              T := Tob.Create('DISPODETAIL', TobGQD, -1);
              wCopyTobBySuffixe(TobSTKError.Detail[iTobSTKError], T, 'GSM', 'GQD');
              T.P('GQD_IDENTIFIANT', wSetJeton('GQD'));
              { on récupère la vrai valeur du physique }
              T.P('GQD_PHYSIQUE', TobSTKError.Detail[iTobSTKError].G(CalculGQD_PHYSIQUE));
            end;
          end;
          WhereDelGQD := '(' + WhereDelGQD + ')';
          { Suppression des enregistrements de DISPODETAIL erronés }
          if OneGQDIsFalse then
          begin
            wDeleteTable('DISPODETAIL', WhereDelGQD);
            { Recontruit et Recalcul DISPODETAIL }
            TobGQD.InsertDB(nil, True);
          end;
        finally
          TobGQD.Free;
        end;
      end;
    finally
      TobSTKError.Free;
    end;
  end;
  { Traitement de l'intégrité enter GSM_PHYSIQUE & GQD_PHYSIQUE ******************* FIN ******************** }

begin
  Result := false;

  { Argument }
  whereGQ := GetArgumentValue(Argument, 'WHERE');

  { Tob système }
  GetTobGSN;
  GetTobCSK;

  wInitProgressForm(nil, GetLibelleAction(TypeAction), '', VH_GP.TobCSK.Detail.Count, False, True);
  try
    for iCSK := 0 to VH_GP.TobCSK.Detail.Count-1 do
    begin
      wMoveProgressForm(VH_GP.TobCSK.Detail[iCSK].G('CO_LIBELLE'));
      sField := GetCompteur;
      if sField <> '' then
      begin
        { Mise à Zero du compteur }
        MiseAZeroDuCompteur(sField);

        if sField = 'GQ_PHYSIQUE' then
        begin
          { MàJ du DISPODETAIL depuis STKMOUVEMENT }
          UpdateGQDByGSM;
          { MàJ du physique par rapport au DISPODETAIL <= MàJ par le STKMOUVEMENT }
          ExecuteSQL('UPDATE DISPO'
                   + ' SET GQ_PHYSIQUE=(SELECT SUM(GQD_PHYSIQUE)'
                   +                   ' FROM DISPODETAIL WHERE'
                   +                   ' GQ_ARTICLE=GQD_ARTICLE'
                   +                   ' AND GQ_DEPOT=GQD_DEPOT)'
                   + ' WHERE ' + WhereGQ
                   );
        end
        else
        begin
          { Je tourne sur les Nature de mouvements}
          WherePhy := '';
          WherePre := '';
          for iGSN := 0 to Vh_GP.TobGSN.Detail.Count-1 do
          begin
            if (Pos(VH_GP.TobCSK.Detail[iCSK].G('CO_CODE'), Vh_GP.TobGSN.Detail[iGSN].G('GSN_QTEPLUS')) > 0) then
            begin
              if GetStkTypeMvt = 'PHY' then
                WherePhy := WherePhy + iif(WherePhy <> '', ' OR ', '') + GetWherePhy(GetQualifMvt)
              else
                WherePre := WherePre + iif(WherePre <> '', ' OR ', '') + GetWherePre(GetStkTypeMvt, GetQualifMvt);
            end;
          end;
          if WherePhy <> '' then
          begin
            { MàJ du PHYSIQUE dans DISPO }
            Sql := 'UPDATE DISPO'
                 + ' SET ' + sField + '=(SELECT SUM(GSM_PHYSIQUE)'
                 +                      ' FROM STKMOUVEMENT'
                 +                      ' WHERE GSM_ARTICLE=GQ_ARTICLE AND GSM_DEPOT=GQ_DEPOT AND (' + WherePhy + '))'
                 + iif(WhereGQ <> '', ' WHERE ' + WhereGQ, '')
                 ;
            Result := ExecuteSql(Sql) > 0;
          end
          else
          if WherePre <> '' then
          begin
            sql := 'UPDATE DISPO'
                 + ' SET ' + sField + '=(SELECT SUM(ABS(GSM_QPREVUE)-ABS(GSM_PHYSIQUE))'
                 +                          ' FROM STKMOUVEMENT'
                 +                          ' WHERE GSM_ARTICLE=GQ_ARTICLE AND GSM_DEPOT=GQ_DEPOT AND GSM_ETATMVT<>"SOL" AND (' + WherePre + '))'
                 + iif(WhereGQ <> '', ' WHERE ' + WhereGQ, '')
                 ;
            Result := ExecuteSql(Sql) > 0;
          end;
        end;

        { Mise à Zero des Nulls }
        MiseAZeroDesNull(sField);
      end;
    end;
  finally
    wFiniProgressForm;
  end;
end;
{$ENDIF}

function ExistsGQ(CleGQ: TCleGQ; WithAlert: Boolean = False): Boolean;
begin
  Result := ExisteSQL('SELECT NULL FROM ' + TableName + ' WHERE ' + WhereGQ(CleGQ));

  if WithAlert and (not Result) then
    PgiError(TraduireMemoire('Aucune information n''est disponible sur le stock de l''article') + ' ' + CleGQ.Article + ' ' + TraduireMemoire('pour le dépôt') + ' ' + CleGQ.Depot);
end;

end.
