unit CPENVOYERMAIL;

interface

uses
  Sysutils, Classes, HEnt1, UTOB;

type
  TEnvoyerMail = class
    Dossier : string;
  private
    function SendMails : string;
  public
    class function EnvoyerMails(T : TOB) : string;
    constructor Create     (Param : TOB);
    destructor Destroy; override;
  end;



implementation

uses
  uLibBonAPayer, HCtrls, uFichierLog;

{---------------------------------------------------------------------------------------}
constructor TEnvoyerMail.Create(Param : TOB);
{---------------------------------------------------------------------------------------}
begin
  Dossier := Param.GetString('DOSSIER');
end;

{---------------------------------------------------------------------------------------}
destructor TEnvoyerMail.Destroy;
{---------------------------------------------------------------------------------------}
begin
  inherited;

end;

{---------------------------------------------------------------------------------------}
class function TEnvoyerMail.EnvoyerMails(T : TOB) : string;
{---------------------------------------------------------------------------------------}
var
  Obj : TEnvoyerMail;
begin
  Result := '';
  try
    Obj := TEnvoyerMail.Create(T);
    try
      if Obj.Dossier <> '' then Result := Obj.SendMails
                           else Result := TraduireMemoire('Dossier non défini');
      if Result = '' then ObjLog.Add('EnvoyerMails : Result = OK')
                     else ObjLog.Add('EnvoyerMails : Result = ' + Result);
    finally
      FreeAndNil(Obj);
    end;
  except
    on E : Exception do begin
      Result := TraduireMemoire('Erreur lors de la tâche ') + psrv_EnvoiMail + ' (EnvoyerMails)'  + #13#13 + E.Message;
      ObjLog.Add('EnvoyerMails : ' + E.Message);
    end;
  end;
end;

{---------------------------------------------------------------------------------------}
function TEnvoyerMail.SendMails : string;
{---------------------------------------------------------------------------------------}
var
  Obj : TObjEnvoiMail;
  TT  : TOB;
  FF  : TOB;
  SQL : string;
  n   : Integer;
begin
  try
    Result := '';
    ObjLog.Add('EnvoyerMails : Début');

    TT := TOB.Create('', nil, -1);
    try
      SQL := 'SELECT * FROM CPTACHEBAP WHERE CTA_DATEENVOI <= "' + UsDateTime(Date) + '" ' +
             'AND (CTA_ENVOYE = "-" OR CTA_ENVOYE IS NULL)';
      TT.LoadDetailDBFromSQL('CPTACHEBAP', SQL);
      for n := TT.Detail.Count - 1 downto 0 do begin
        FF := TT.Detail[n];
        Obj := TObjEnvoiMail.Create('', '', 1, FF);
        try
          if not Obj.EnvoieMail then
            Result := TraduireMemoire(DateTimeToStr(Date) + ' : Certains mails pour les bons à payer n''ont pu être envoyés faute d''un paramétrage complet');
        finally
          FreeAndNil(Obj);
        end;
      end;
      ObjLog.Add('EnvoyerMails : Fin');
    finally
      FreeAndNil(TT);
    end;
  except
    on E : Exception do begin
      Result := TraduireMemoire('Erreur lors de la tâche ') + psrv_EnvoiMail + ' (SendMails)' + #13#13 + E.Message;
      ObjLog.Add('EnvoyerMails : ' + E.Message);
    end;
  end;
end;

end.
