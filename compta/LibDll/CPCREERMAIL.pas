unit CPCREERMAIL;

interface

uses
  Sysutils, Classes, HEnt1, UTob;

type
  TCreerMail = class
    Dossier : string;
  private
    procedure UpdateBAP(TBap : TOB);
  public
    class function CreerMails(T : TOB) : string;
    constructor Create(Param : TOB);
    destructor Destroy; override;
    function LanceTraitement : string;
    function GereMailInitiateur : string;
  end;

implementation


uses
  uLibBonAPayer, HCtrls, uFichierLog{, UTaskJob};

{---------------------------------------------------------------------------------------}
class function TCreerMail.CreerMails(T : TOB) : string;
{---------------------------------------------------------------------------------------}
var
  Obj : TCreerMail;
  ch  : string; 
begin
  Result := '';
  try
    Obj := TCreerMail.Create(T);
    try
      if Obj.Dossier <> '' then begin
        {Préparation des mails pour les viseurs}
        Result := Obj.LanceTraitement;
        {Préparation des mails de retour aux initiateurs des BAP en cas de "Refus"}
        ch := Obj.GereMailInitiateur;
        if Result = '' then Result := ch
                       else Result := Result + #13#13 + ch;
      end
      else
        Result := TraduireMemoire('Dossier non défini');
      if Result = '' then ObjLog.Add('CreerMails : Result = OK')
                     else ObjLog.Add('CreerMails : Result = ' + Result);
    finally
      FreeAndNil(Obj);
    end;
  except
    on E : Exception do begin
      Result := TraduireMemoire('Erreur lors de la tâche ') + 'CreerMails'#13#13 + E.Message;
      ObjLog.Add('CreerMails : ' + E.Message);
    end;
  end;
end;

{---------------------------------------------------------------------------------------}
constructor TCreerMail.Create(Param : TOB);
{---------------------------------------------------------------------------------------}
begin
  Dossier := Param.GetString('DOSSIER');
end;

{---------------------------------------------------------------------------------------}
destructor TCreerMail.Destroy;
{---------------------------------------------------------------------------------------}
begin
  inherited;
end;

{---------------------------------------------------------------------------------------}
function TCreerMail.LanceTraitement : string;
{---------------------------------------------------------------------------------------}
var
  TBap : TOB;
  TTac : TOB;
  TMai : TOB;
  T    : TOB;
  F    : TObjTacheBap;
  M    : TOB;
  n    : Integer;
begin
  Result := '';
  try
    TBap := TOB.Create('', nil, -1);
    TTac := TOB.Create('', nil, -1);
    TMai := TOB.Create('', nil, -1);
    try
      ObjLog.Add('CreerMails : Avant Chargement des Tobs');
      TMai.LoadDetailFromSQL('SELECT CMA_BLOCNOTE, CMA_CODEMAIL, CMA_NATUREMAIL FROM CPMAILBAP');
      {On récupère tous les BAP dont la relance 2 n'a pas été faite et dont la date d'échéance (du BAP) est dépassée}
      TBap.LoadDetailFromSQL('SELECT BAP_DATEMAIL, BAP_DATERELANCE1, BAP_DATERELANCE2, BAP_ECHEANCEBAP, ' +
                   'BAP_RELANCE1, BAP_RELANCE2, BAP_STATUTBAP, BAP_VISEUR1, BAP_ALERTE1, BAP_ALERTE2, ' +
                   'BAP_VISEUR2, BAP_JOURNAL, BAP_EXERCICE, BAP_DATECOMPTABLE, BAP_NUMEROPIECE, BAP_NUMEROORDRE, ' +
                   'BAP_CIRCUITBAP, BAP_CODEVISA, U1.US_EMAIL ADDVIS1, U2.US_EMAIL ADDVIS2 FROM CPBONSAPAYER ' +
                   'LEFT JOIN UTILISAT U1 ON U1.US_UTILISATEUR = BAP_VISEUR1 ' +
                   'LEFT JOIN UTILISAT U2 ON U2.US_UTILISATEUR = BAP_VISEUR2 ' +
                   'WHERE BAP_STATUTBAP = "' + sbap_Encours + '" AND (((BAP_RELANCE2 = "-" OR BAP_RELANCE2 IS NULL) AND ' +
                   'BAP_ECHEANCEBAP < "' + UsDateTime(Date) + '") OR ((BAP_RELANCE1 = "-" OR BAP_RELANCE1 IS NULL) AND ' +
                   'BAP_ECHEANCEBAP < "' + UsDateTime(Date) + '") OR (BAP_ALERTE1 = "-" OR BAP_ALERTE1 IS NULL)) ORDER BY BAP_VISEUR1, BAP_VISEUR2');
      ObjLog.Add('CreerMails : Après Chargement des Tobs');

      for n := 0 to TBap.Detail.Count - 1 do begin
        T := TBap.Detail[n];

        if T.GetString('BAP_ALERTE1') <> 'X' then begin
          F := TObjTacheBap.CreateObj(nam_Alerte, T.GetString('BAP_VISEUR1'), Date, TTac);
          F.ViseurPrincipal := True;
          M := TMai.FindFirst(['CMA_NATUREMAIL'], [nam_Alerte], True);
          F.PrepareTache(T, M);

          {Mise à jour des champs idoines dans la table BAP}
          T.SetString('BAP_ALERTE1', 'X');
          T.SetDateTime('BAP_DATEMAIL', Date);
        end
        else if T.GetString('BAP_RELANCE1') <> 'X' then begin
          {Création d'un mail de relance au viseur 1}
          F := TObjTacheBap.CreateObj(nam_Relance, T.GetString('BAP_VISEUR1'), Date, TTac);
          F.ViseurPrincipal := True;
          M := TMai.FindFirst(['CMA_NATUREMAIL'], [nam_Relance], True);
          F.PrepareTache(T, M);

          {Création d'un mail d'alerte suppléant au viseur 2}
          F := TObjTacheBap.CreateObj(nam_Suppleant, T.GetString('BAP_VISEUR2'), Date, TTac);
          F.ViseurPrincipal := False;
          M := TMai.FindFirst(['CMA_NATUREMAIL'], [nam_Suppleant], True);
          F.PrepareTache(T, M);

          {Mise à jour des champs idoines dans la table BAP}
          T.SetString('BAP_RELANCE1', 'X');
          T.SetString('BAP_ALERTE2', 'X');
          T.SetDateTime('BAP_DATERELANCE1', Date);
        end
        else if T.GetString('BAP_RELANCE2') <> 'X' then begin
          {Création d'un mail de relance au viseur 1}
          F := TObjTacheBap.CreateObj(nam_Relance, T.GetString('BAP_VISEUR1'), Date, TTac);
          F.ViseurPrincipal := True;
          M := TMai.FindFirst(['CMA_NATUREMAIL'], [nam_Relance], True);
          F.PrepareTache(T, M);

          {Création d'un mail de relance au viseur 2}
          F := TObjTacheBap.CreateObj(nam_Relance, T.GetString('BAP_VISEUR2'), Date, TTac);
          F.ViseurPrincipal := False;
          M := TMai.FindFirst(['CMA_NATUREMAIL'], [nam_Relance], True);
          F.PrepareTache(T, M);

          {Mise à jour des champs idoines dans la table BAP}
          T.SetString('BAP_RELANCE2', 'X');
          T.SetDateTime('BAP_DATERELANCE2', Date);
        end;
      end;

      ObjLog.Add('CreerMails : Avant InsertOrUpdateDB');
      UpdateBAP(TBap);
      TTac.InsertOrUpdateDB;
      ObjLog.Add('CreerMails : Apès InsertOrUpdateDB');
    finally
      FreeAndNil(TBap);
      FreeAndNil(TTac);
      FreeAndNil(TMai);
    end;
  except
    on E : Exception do begin
      Result := TraduireMemoire('Erreur lors de la tâche ') + psrv_Alerte + ' (LanceTraitement)' + #13#13 + E.Message;
      ObjLog.Add('CreerMails : ' + E.Message);
      raise E;
    end;
  end;
end;

{On envoie un mail d'alerte à l'initiateur lors de bap bloqués ou refusés :
 -> à l'échéance du BAP
 -> à sept jours de l'échéance de facture
{---------------------------------------------------------------------------------------}
function TCreerMail.GereMailInitiateur : string;
{---------------------------------------------------------------------------------------}
var
  TBap : TOB;
  TTac : TOB;
  TMai : TOB;
  T    : TOB;
  F    : TObjTacheBap;
  M    : TOB;
  n    : Integer;
begin
  Result := '';
  try
    TBap := TOB.Create('', nil, -1);
    TTac := TOB.Create('', nil, -1);
    TMai := TOB.Create('', nil, -1);
    try
      ObjLog.Add('GereMailInitiateur : Avant Chargement des Tobs');
      TMai.LoadDetailFromSQL('SELECT CMA_BLOCNOTE, CMA_CODEMAIL, CMA_NATUREMAIL FROM CPMAILBAP');
      {On récupère tous les BAP bloqués, refusés, analytiques dont la date d'échéance égale celle du jour ou
       dont la date d'échéance de la facture est dans 7 jours}
      TBap.LoadDetailFromSQL('SELECT BAP_DATEECHEANCE, BAP_ECHEANCEBAP, BAP_STATUTBAP, BAP_CREATEUR, ' +
                   'BAP_JOURNAL, BAP_EXERCICE, BAP_DATECOMPTABLE, BAP_NUMEROPIECE, BAP_NUMEROORDRE, ' +
                   'BAP_CIRCUITBAP, BAP_CODEVISA, BAP_CREATEUR, U1.US_EMAIL ADDVIS1 FROM CPBONSAPAYER ' +
                   'LEFT JOIN UTILISAT U1 ON US_UTILISATEUR = BAP_CREATEUR ' +
                   'WHERE BAP_STATUTBAP IN ("' + sbpa_bloque + '", "' + sbap_Analytique + '", "' + sbap_Refuse + '") AND ' +
                   '(BAP_ECHEANCEBAP = "' + UsDateTime(Date) + '" OR BAP_DATEECHEANCE = "' + UsDateTime(Date + 5) + '" ) ORDER BY BAP_CREATEUR');

      ObjLog.Add('GereMailInitiateur : Après Chargement des Tobs');

      for n := 0 to TBap.Detail.Count - 1 do begin
        T := TBap.Detail[n];

        if T.Getstring('BAP_STATUTBAP') = sbap_Refuse  then begin
          F := TObjTacheBap.CreateObj(nam_Refuse, T.GetString('BAP_CREATEUR'), Date, TTac);
          F.ViseurPrincipal := True;
          M := TMai.FindFirst(['CMA_NATUREMAIL'], [nam_Refuse], True);
          F.PrepareTache(T, M);
        end
        else begin
          F := TObjTacheBap.CreateObj(nam_Bloque, T.GetString('BAP_CREATEUR'), Date, TTac);
          F.ViseurPrincipal := True;
          M := TMai.FindFirst(['CMA_NATUREMAIL'], [nam_Bloque], True);
          F.PrepareTache(T, M);
        end;
      end;

      ObjLog.Add('GereMailInitiateur : Avant InsertOrUpdateDB');
      TTac.InsertOrUpdateDB;
      ObjLog.Add('GereMailInitiateur : Après InsertOrUpdateDB');

    finally
      FreeAndNil(TBap);
      FreeAndNil(TTac);
      FreeAndNil(TMai);
    end;
  except
    on E : Exception do begin
      Result := TraduireMemoire('Erreur lors de la tâche ') + psrv_Alerte + ' (GereMailInitiateur)' + #13#13 + E.Message;
      ObjLog.Add('GereMailInitiateur : ' + E.Message);
      raise E;
    end;
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TCreerMail.UpdateBAP(TBap : TOB);
{---------------------------------------------------------------------------------------}
var
  n : Integer;
  T : TOB;
  S : string;
begin
  try
    for n := 0 to TBap.Detail.Count - 1 do begin
      T := TBap.Detail[n];
      S := 'UPDATE CPBONSAPAYER SET BAP_DATEMAIL = "' + UsDateTime(T.GetDateTime('BAP_DATEMAIL')) + '", ' +
                                   'BAP_DATERELANCE1 = "' + UsDateTime(T.GetDateTime('BAP_DATERELANCE1')) + '", ' +
                                   'BAP_DATERELANCE2 = "' + UsDateTime(T.GetDateTime('BAP_DATERELANCE2')) + '", ' +
                                   'BAP_ALERTE1 = "' + T.GetString('BAP_ALERTE1') + '", ' +
                                   'BAP_ALERTE2 = "' + T.GetString('BAP_ALERTE2') + '", ';
      S := S +                     'BAP_RELANCE1 = "' + T.GetString('BAP_RELANCE1') + '", ' +
                                   'BAP_RELANCE2 = "' + T.GetString('BAP_RELANCE2') + '", ' +
                                   'BAP_DATEMODIF = "' + UsDateTime(Now) + '", ' +
                                   'BAP_MODIFICATEUR = "' + USERTACHE  + '" ';
      S := S +   ' WHERE BAP_JOURNAL = "' + T.GetString('BAP_JOURNAL') + '" AND ' +
                        'BAP_EXERCICE = "' + T.GetString('BAP_EXERCICE') + '" AND ' +
                        'BAP_DATECOMPTABLE = "' + UsDateTime(T.GetDateTime('BAP_DATECOMPTABLE')) + '" AND ' +
                        'BAP_NUMEROPIECE = ' + T.GetString('BAP_NUMEROPIECE') + ' AND ' +
                        'BAP_NUMEROORDRE = ' + T.GetString('BAP_NUMEROORDRE');

      ObjLog.Add('UpdateBAP : ' + S);
      ExecuteSQL(S);
    end;
  except
    on E : Exception do begin
      ObjLog.Add('UpdateBAP : ' + E.Message);
      raise;
    end;
  end;
end;

end.

