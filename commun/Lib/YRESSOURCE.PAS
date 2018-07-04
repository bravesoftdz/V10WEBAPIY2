{***********UNITE*************************************************
Auteur  ...... : Jean-Luc Sauzet
Créé le ...... : 15/01/2007
Modifié le ... :   /  /
Description .. : Fonctions & procédure métier
Mots clefs ... : YRESSOURCE
*****************************************************************}
unit YRESSOURCE;

interface

uses
{$IFNDEF EAGLCLIENT}
  {$IFNDEF DBXPRESS}dbtables{$ELSE}uDbxDataSet{$ENDIF},
{$IFNDEF EAGLSERVER}
  {$IFNDEF ERADIO}
    FE_Main,
  {$ENDIF !ERADIO}
{$ENDIF EAGLSERVER}
{$ELSE}
  MainEAGL,
{$ENDIF}
{$IFNDEF EAGLSERVER}
  wParam,
{$ENDIF EAGLSERVER}
  wCommuns
  ;  

  { Méthodes }
  function CreateYRS(Const Salarie, Ressource, User: string): integer;  { Création d'enregistrements dans YRESSOURCE }
  function UpdateYRS(Const Salarie, Ressource, User: string): integer;  { Modification d'enregistrements dans YRESSOURCE }
  function DeleteYRS(Const Salarie, Ressource, User: string): integer;  { Suppression d'enregistrements dans YRESSOURCE }
  function GetLastErrorMsgYRS(Const LastError: integer): String;        { Renvoie le texte de l'erreur }

  function WhereYRS(Const Salarie, Ressource, User: string): String;
  function ExistYRS(Const Salarie, Ressource, User: string): boolean;
  function ExistYRSSalarie(Const Salarie: string): boolean;
  function ExistYRSRessource(Const Ressource: string): boolean;
  function ExistYRSUser (const strUser, srtRessource : string) : boolean; // $$$ JP 23/04/07  
  function GetYRS_GUID(Const Salarie, Ressource, User: string): String;
  procedure RecalculYRS(Silence:Boolean=False);                 { Recalcul de YRESSOURCE }

const

  TexteMessage: array[1..4] of string = (
    {1} 'Traitement des ressources en cours, veuillez patienter...',
    {2} 'Traitement en cours ...',
    {3} 'Traitement des salariés en cours, veuillez patienter...',
    {4} 'Traitement des utilisateurs en cours, veuillez patienter...'
    );

implementation

uses
  Hctrls, uTob, uTom, yPlanning, YRESSOURCE_TOM, ed_tools, hent1, sysutils, HMsgBox;

{***********A.G.L.***********************************************
Auteur  ...... : Jean-Luc SAUZET
Créé le ...... : 23/01/2007
Modifié le ... :   /  /
Description .. : Création d'un enregistrement dans YRS:
Description .. : -Crée la ligne si elle n'existe pas
Description .. : -Met à jour si un des éléments existe déjà
Description .. :
Description .. :
Mots clefs ... :
*****************************************************************}
function CreateYRS(Const Salarie, Ressource, User: string): integer;
var
  TomYRS: Tom;
  TobYRS: Tob;

  // $$$ JP 23/04/07
  // 20080828_C.B_FQ15215_debut
  // il est possible d'avoir 2 ressource différentes pour un même user
  // si une seule de ces 2 ressources est de type salarié, alors c'est ok
  // on crée les 2 enregistrements
  // si les 2 ressources sont de type salarié -> erreur
  // si la ressource est celle qui est déjà dans la base, alors on update,
  // sinon on créer
  function DoubleUserOk (const strUser : String) : Boolean;
  var
    Sql : string;
    TobRes : Tob;
  begin
    Result := True;
    if strUser <> '' then
    begin
      TobRes := Tob.Create('#RESSOURCE', nil, -1);
      try
      Sql := 'SELECT ARS_RESSOURCE FROM RESSOURCE WHERE ARS_UTILASSOCIE = "'
             + strUser + '" AND ARS_TYPERESSOURCE = "SAL"';
      if TobRes.LoadDetailDBFromSql('RESSOURCE', Sql) then
        result := TobRes.Detail.Count < 2;

      finally
        TobRes.Free;
      end;
    end;
  end;
  // 20080828_C.B_FQ15215_fin

begin

  Result := 0;
  if not GestionYPlanning then exit;
  if ExistYRSSalarie(Salarie) then
    UpdateYRS(Salarie, Ressource, User)
  else if ExistYRSRessource(Ressource) then
    UpdateYRS(Salarie, Ressource, User)

  // $$$ JP 23/04/07: il faut aussi vérifier s'il n'y a pas de ressource lié au user
  // 20080828_C.B_FQ15215_debut
  // on ne teste pas ici le cas d'una autre ressource avec le même user, c'est le test suivant
  else if ExistYRSUser(User, Ressource) then
     UpdateYRS (Salarie, Ressource, User)

  // il est possible d'avoir 2 ressource différentes pour un même user
  // si une seule de ces 2 ressources est de type salarié, alors c'est ok, on peut créer le nouvel enregistrement
  // on vérifie donc si la ressource ou le salarie à le même code user
  // et si un des 2 n'est pas salarié
  else if (not DoubleUserOk(User)) then
  begin
    result := -1;
    PGIBox(traduireMemoire(format(GetLastErrorMsgYRS(1), [User, Ressource])), traduireMemoire(GetLastErrorMsgYRS(3)));
  end

  // 20080828_C.B_FQ15215_fin

  else
  begin
    TomYRS := CreateTOM('YRESSOURCE', nil, false, true);
    TobYRS := Tob.Create('YRESSOURCE', nil, -1);
    try
      { OnNewRecord }
      TomYRS.InitTOB(TobYRS);

      { Données}
      TobYRS.SetString('YRS_SALARIE'  , Salarie);
      TobYRS.SetString('YRS_RESSOURCE', Ressource);
      TobYRS.SetString('YRS_USER'     , User);

      { OnUpdateRecord }
      TobYRS.AddChampSupValeur('IKC', 'C', false);
      if TomYRS.VerifTOB(TobYRS) then
      begin
        if TobYRS.InsertDb(nil) then
        begin
          { OnAfterUpdaterecord }
          TomYRS.AfterVerifTOB(TobYRS);
        end;
      end
      else
        Result := TomYRS.LastError;
    finally
      TobYRS.free;
      TomYRS.free;
    end;
  end;
end;

function UpdateYRS(Const Salarie, Ressource, User: string): integer;
var
  i     : integer;
  Sql   : string;
  TomYRS: Tom;
  TobYRS: Tob;
begin
  Result := 0;

  if not GestionYPlanning then exit;

  // C.B 02/04/2007
  // si une seule donnée est renseignée, on sort
  if ((Salarie = '') and (Ressource = '')) or
     ((User = '') and (Ressource = '')) or
     ((Salarie = '') and (User = '')) then exit;

  TomYRS := CreateTOM('YRS', nil, false, true);
  TobYRS := Tob.Create('YRESSOURCE', nil, -1);
  try
    Sql := 'SELECT *'
         + ' FROM YRESSOURCE'
         + ' WHERE 1=0'
         + iif (Salarie   <> '', ' OR YRS_SALARIE="'   + Salarie   + '"', '')
         + iif (Ressource <> '', ' OR YRS_RESSOURCE="' + Ressource + '"', '')
         + iif (User      <> '', ' OR YRS_USER="'      + User      + '"', '')
         ;
    if TobYRS.LoadDetailDBFromSql('YRESSOURCE', Sql) then
    begin
      for i := 0 to TobYRS.Detail.Count-1 do
      begin
        if (TobYRS.Detail[i].GetString('YRS_SALARIE') = '') and (TobYRS.Detail[i].GetString('YRS_RESSOURCE') = Ressource) then
          TobYRS.Detail[i].SetString('YRS_SALARIE', Salarie);
        if (TobYRS.Detail[i].GetString('YRS_RESSOURCE') = '') and (TobYRS.Detail[i].GetString('YRS_SALARIE') = Salarie) then
          TobYRS.Detail[i].SetString('YRS_RESSOURCE', Ressource);

        //C.B 02/04/2007
        // on laisse la mauvaise donnée si remise à blanc pour l'instant
        if (TobYRS.Detail[i].GetString('YRS_USER') = '') then
          TobYRS.Detail[i].SetString('YRS_USER', User);

        { OnUpdateRecord }
        TobYRS.AddChampSupValeur('IKC', 'M', false);
        if TomYRS.VerifTOB(TobYRS.Detail[i]) then
        begin
          if TobYRS.Detail[i].UpdateDb then
          begin
            { OnAfterUpdaterecord }
            TomYRS.AfterVerifTOB(TobYRS.Detail[i]);
          end
          else
          begin
            Result := -1;
            Break;
          end;
        end
        else
        begin
          Result := TomYRS.LastError;
          Break;
        end;
      end;
    end;
  finally
    TobYRS.free;
    TomYRS.free;
  end;

  { Gestion des doublons }
  { Il se peut qu'il y ai des doublons entre Salarie et Ressource }
  { => Ne faire q'un seul enregistrement }
  {  1° Mettre à jour YPLANNING }
  {  2° Supprimer le doublon }

  if (Salarie <> '') and (Ressource <> '') then
  begin
    TobYRS := Tob.Create('YRESSOURCE', nil, -1);
    try
      Sql := 'SELECT YRS_GUID'
           + ' FROM YRESSOURCE'
           + ' WHERE YRS_SALARIE="' + Salarie + '"'
           + ' AND YRS_RESSOURCE="' + Ressource + '"'
           ;
      if TobYRS.LoadDetailDBFromSql('YRS', Sql) then
      begin
        if TobYRS.Detail.Count > 1 then
        begin
          { 1° }
          Sql := 'UPDATE YPLANNING'
               + ' SET YPL_GUIDYRS="' + TobYRS.Detail[0].GetString('YRS_GUID') + '"'
               + ' WHERE 1=0'
               ;
          for i := 1 to TobYRS.Detail.Count-1 do
          begin
            Sql := Sql + ' OR YPL_GUIDYRS="' + TobYRS.Detail[i].GetString('YRS_GUID') + '"'
          end;
          ExecuteSql(Sql);

          { 2° }
          Sql := 'DELETE YRESSOURCE'
               + ' WHERE 1=0'
               ;
          for i := 1 to TobYRS.Detail.Count-1 do
          begin
            Sql := Sql + ' OR YRS_GUID="' + TobYRS.Detail[i].GetString('YRS_GUID') + '"'
          end;
          ExecuteSql(Sql);
        end;
      end;
    finally
      TobYRS.free;
    end;
  end;
                       
  { Gestion des doublons }
  { Il se peut qu'il y ai des doublons entre utilisateur et Ressource }
  { => Ne faire q'un seul enregistrement }
  {  1° Mettre à jour YPLANNING }
  {  2° Supprimer le doublon }                        
  if (User <> '') and (Ressource <> '') then
  begin
    TobYRS := Tob.Create('YRESSOURCE', nil, -1);
    try
      Sql := 'SELECT YRS_GUID'
           + ' FROM YRESSOURCE'
           + ' WHERE YRS_USER="' + User + '"'
           + ' AND YRS_RESSOURCE="' + Ressource + '"'
           ;
      if TobYRS.LoadDetailDBFromSql('YRS', Sql) then
      begin
        if TobYRS.Detail.Count > 1 then
        begin
          { 1° }
          Sql := 'UPDATE YPLANNING'
               + ' SET YPL_GUIDYRS="' + TobYRS.Detail[0].GetString('YRS_GUID') + '"'
               + ' WHERE 1=0'
               ;
          for i := 1 to TobYRS.Detail.Count-1 do
          begin
            Sql := Sql + ' OR YPL_GUIDYRS="' + TobYRS.Detail[i].GetString('YRS_GUID') + '"'
          end;
          ExecuteSql(Sql);

          { 2° }
          Sql := 'DELETE YRESSOURCE'
               + ' WHERE 1=0'
               ;
          for i := 1 to TobYRS.Detail.Count-1 do
          begin
            Sql := Sql + ' OR YRS_GUID="' + TobYRS.Detail[i].GetString('YRS_GUID') + '"'
          end;
          ExecuteSql(Sql);
        end;
      end;
    finally
      TobYRS.free;
    end;
  end;

end;

{***********A.G.L.***********************************************
Auteur  ...... : Jean-Luc SAUZET
Créé le ...... : 29/01/2007
Modifié le ... : 29/01/2007
Description .. : Suppression dans la table YRessource:
Suite ........ : Update à blanc du delete demandé.
Suite ........ : Delete de l'éventuelle ligne qui contient les trois codes à
Suite ........ : blanc.
Mots clefs ... :
*****************************************************************}
function DeleteYRS(Const Salarie, Ressource, User: string): integer;
begin
  Result := 0;
  if not GestionYPlanning then exit;
  if Salarie <> '' then
    wDeleteTable('YRESSOURCE', 'YRS_SALARIE="' + Salarie + '"');
  if Ressource <> '' then
    wDeleteTable('YRESSOURCE', 'YRS_RESSOURCE="' + Ressource + '"');
  if User <> '' then
    wDeleteTable('YRESSOURCE', 'YRS_USER="' + User + '"');
     
  wDeleteTable('YRESSOURCE', 'YRS_RESSOURCE="" AND YRS_SALARIE="" AND YRS_USER=""')
end;

function GetLastErrorMsgYRS(Const LastError: integer): String;
var
  TomYRS: Tom;
begin
  TomYRS := CreateTOM('YRESSOURCE', nil, false, true);
  try
    Result := Tom_YRESSOURCE(TomYRS).GetLastErrorMsg(LastError);
  finally
    TomYRS.free;
  end;
end;

function WhereYRS(Const Salarie, Ressource, User: string): String;
begin
  Result := '';

  if Salarie <> '' then
    Result := 'YRS_SALARIE="' + Salarie + '"';

  if (Result = '') and (Ressource <> '') then
    Result := 'YRS_RESSOURCE="' + Ressource + '"';

  if (Result = '') and (User <> '') then
    Result := 'YRS_USER="' + User + '"';
end;

function ExistYRS(Const Salarie, Ressource, User: string): boolean;
var
  Sql: string;
begin
  Sql := 'SELECT 1'
       + ' FROM YRESSOURCE'
       + ' WHERE ' + WhereYRS(Salarie, Ressource, User)
       ;
  Result := ExisteSql(Sql);
end;

function ExistYRSSalarie(Const Salarie: string): boolean;
var
  Sql: string;
begin
  if Salarie <> '' then
  begin
    Sql := 'SELECT 1'
         + ' FROM YRESSOURCE'
         + ' WHERE YRS_SALARIE="' + Salarie + '"'
         ;
    Result := ExisteSql(Sql);
  end
  else
    result := False;
end;

function ExistYRSRessource(Const Ressource: string): boolean;
var
  Sql: string;
begin
  if Ressource <> '' then
  begin
    Sql := 'SELECT 1'
         + ' FROM YRESSOURCE'
         + ' WHERE YRS_RESSOURCE="' + Ressource + '"'
         ;
    Result := ExisteSql(Sql);
  end
  else
    result := False;
end;

// $$$ JP 23/04/07
function ExistYRSUser (const strUser, srtRessource : string) : boolean;
var
  Sql: string;
begin
  Result := FALSE;
  if strUser <> '' then
  begin
      Sql := 'SELECT 1 FROM YRESSOURCE WHERE YRS_USER="' + strUser + '"'
             + ' AND (YRS_RESSOURCE = "' + srtRessource + '"'
             + ' OR YRS_RESSOURCE = "")' ;
                                            
      Result := ExisteSql (Sql);
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Céline BOUET
Créé le ...... : 15/02/2007
Modifié le ... : 15/02/2007
Description .. : retourne la clé de la table YRESSOURCE
Mots clefs ... :
*****************************************************************}
function GetYRS_GUID(Const Salarie, Ressource, User: string): String;
var
  vTob : Tob;
  Sql : String;

begin

  vTob := Tob.Create('YRESSOURCE', nil, -1);
  Try
    Sql := 'SELECT YRS_GUID FROM YRESSOURCE WHERE ' + WhereYRS(Salarie, Ressource, User);
    vTob.LoadDetailFromSQL(Sql);
    if vTob.Detail.count = 1 then
      result := vTob.Detail[0].GetString('YRS_GUID')
    else
      result := '';
  Finally
    vTob.Free;
  End;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Céline BOUET
Créé le ...... : 11/05/2007
Modifié le ... : 11/05/2007
Description .. : recalcul la table YRESSOURCE à partir des tables
                 SALARIE, RESSOURCE et UTILISAT
                 1) suppression des données
                 2) ecriture directe des données de la table ressource
                 qui contient ARS_RESSOURCE, ARS_SALARIE et ARS_UTILASSOCIE
                 3) ecriture en utilisant les méthodes de controle de yressource
                 des données de la table Salarié avec champ PCN_SALARIE
                 4) ecriture en utilisant les méthodes de controle de yressource
                 des données de la table Utilisat avec champ US_UTILISATEUR et US_AUXILIAIRE (si renseigné)
Mots clefs ... :
*****************************************************************}
procedure RecalculYRS(Silence:Boolean=False);
var
  i     : Integer;
  vSt   : String;
  vTob  : Tob;

begin

  // 1) suppression des données
  executeSql('DELETE FROM YRESSOURCE');

  // 2) ecriture des données de la table ressource
  //    qui contient ARS_RESSOURCE, ARS_SALARIE et ARS_UTILASSOCIE
  vSt := 'SELECT ARS_RESSOURCE, ARS_SALARIE, ARS_UTILASSOCIE FROM RESSOURCE ORDER BY ARS_RESSOURCE';
  vTob := Tob.Create('#RESSOURCE', nil, -1);
  try
    vTob.LoadDetailDBFromSql('#RESSOURCE', vSt);

    if not Silence then
    begin
      SourisSablier;
      InitMoveProgressForm(nil, '',traduireMemoire(TexteMessage[1]), vTob.detail.count, False, True);
    end;

    for i:= 0 to vTob.detail.count -1 do
    begin
      //20080828_C.B_GA15215_debut
      MoveCurProgressForm();
      if (CreateYRS(vTob.detail[i].GetString('ARS_SALARIE'),
                vTob.detail[i].GetString('ARS_RESSOURCE'),
                vTob.detail[i].GetString('ARS_UTILASSOCIE')) <> 0) then
      begin
        if (not silence) then
          PGIBox(traduireMemoire(format(GetLastErrorMsgYRS(2), [vTob.detail[i].GetString('ARS_RESSOURCE')])), traduireMemoire(GetLastErrorMsgYRS(3)));
        exit;
      end;
      //20080828_C.B_GA15215_fin
    end;

  finally
    vTob.Free;
    if not Silence then
    begin
      //MoveCurProgressForm(traduireMemoire(TexteMessage[2])); //20080828_C.B_GA15215
      FiniMoveProgressForm;
      SourisNormale;
    end;
  end;

  // 3) ecriture en utilisant les méthodes de controle de yressource
  // des données de la table Salarié avec champ PCN_SALARIE
  vSt := 'SELECT PSA_SALARIE FROM SALARIES';
  vTob := Tob.Create('#SALARIES', nil, -1);
  try
    vTob.LoadDetailDBFromSql('#SALARIES', vSt);

    if not Silence then
    begin
      SourisSablier;
      InitMoveProgressForm(nil, '',traduireMemoire(TexteMessage[3]), vTob.detail.count, False, True);
    end;

    //20080828_C.B_GA15215_debut
    for i := 0 to vTob.detail.count -1 do
    begin
      MoveCurProgressForm();
      if (CreateYRS(vTob.detail[i].GetString('PSA_SALARIE'), '', '') <> 0) then
      begin
        if (not silence) then
          PGIBox(traduireMemoire(format(GetLastErrorMsgYRS(4), [vTob.detail[i].GetString('PSA_SALARIE')])), traduireMemoire(GetLastErrorMsgYRS(3)));
        exit;
      end;
    end;
    //20080828_C.B_GA15215_fin

  finally
    vTob.Free;
    if not Silence then
    begin
      //MoveCurProgressForm(traduireMemoire(TexteMessage[2])); //20080828_C.B_GA15215
      FiniMoveProgressForm;
      SourisNormale;
    end;
  end;

  //4) ecriture en utilisant les méthodes de controle de yressource
  //   des données de la table Utilisat avec champ US_UTILISATEUR et US_AUXILIAIRE (si renseigné)
  vSt := 'SELECT US_UTILISATEUR, US_AUXILIAIRE FROM UTILISAT';
  vTob := Tob.Create('#UTILISAT', nil, -1);
  try
    vTob.LoadDetailDBFromSql('#UTILISAT', vSt);

    if not Silence then
    begin
      SourisSablier;
      InitMoveProgressForm(nil, '',traduireMemoire(TexteMessage[4]), vTob.detail.count, False, True);
    end;

    //20080828_C.B_GA15215_debut
    for i := 0 to vTob.detail.count -1 do
    begin
      MoveCurProgressForm();
      if (CreateYRS(vTob.detail[i].GetString('US_AUXILIAIRE'), '',
                vTob.detail[i].GetString('US_UTILISATEUR')) <> 0) then
      begin
        if (not silence) then
          PGIBox(traduireMemoire(format(GetLastErrorMsgYRS(5), [vTob.detail[i].GetString('US_UTILISATEUR')])), traduireMemoire(GetLastErrorMsgYRS(3)));
        exit;
      end;
    end;
    //20080828_C.B_GA15215_fin

  finally
    vTob.Free;
    if not Silence then
    begin 
      //MoveCurProgressForm(traduireMemoire(TexteMessage[2])); //20080828_C.B_GA15215
      FiniMoveProgressForm;
      SourisNormale;
    end;
  end;

end;

end.
