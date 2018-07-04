unit CPETAPEBAP;

interface

uses
  Sysutils, Classes, HEnt1, UTob;

type
  TCreerEtapeBap = class
    Dossier : string;
  public
    class function CreerEtape(T : TOB) : string;
    constructor Create(Param : TOB);
    destructor Destroy; override;
    function LanceTraitement : string;
  end;

implementation


uses
  uLibBonAPayer, HCtrls, uFichierLog,
  Registry, Windows, eHttp;

{---------------------------------------------------------------------------------------}
class function TCreerEtapeBap.CreerEtape(T : TOB) : string;
{---------------------------------------------------------------------------------------}
var
  Obj : TCreerEtapeBap;
begin
  Result := '';
  try
    ObjLog.Add('Avant création de l''objet Etape');
    Obj := TCreerEtapeBap.Create(T);
    try
      if Obj.Dossier <> '' then
        Result := Obj.LanceTraitement
      else
        Result := TraduireMemoire('Dossier non défini');

      if Result = '' then ObjLog.Add('CreerMails : Result = OK')
                     else ObjLog.Add('CreerMails : Result = ' + Result);
    finally
      FreeAndNil(Obj);
    end;
  except
    on E : Exception do begin
      Result := TraduireMemoire('Erreur lors de la tâche ') + psrv_Etape + #13#13 + E.Message;
      ObjLog.Add('CreerEtape : ' + E.Message);
    end;
  end;
end;

{---------------------------------------------------------------------------------------}
constructor TCreerEtapeBap.Create(Param : TOB);
{---------------------------------------------------------------------------------------}
begin
  Dossier := Param.GetString('DOSSIER');
end;

{---------------------------------------------------------------------------------------}
destructor TCreerEtapeBap.Destroy;
{---------------------------------------------------------------------------------------}
begin
  inherited;
end;

{---------------------------------------------------------------------------------------}
function TCreerEtapeBap.LanceTraitement : string;
{---------------------------------------------------------------------------------------}
var
  TBap : TOB;
  T    : TOB;
  n    : Integer;
  SQL  : string;
  oBap : TBonAPayer;
begin
  Result := '';
  try
  (*
    sServer := '';
    sFrom   := '';

    with TRegistry.Create do begin
      RootKey := HKEY_LOCAL_MACHINE;
      if OpenKey('SOFTWARE\CEGID_RM\PgiService\Journal', False) then begin
        if ValueExists('smtpServer') then sServer := ReadString('smtpServer');
        if ValueExists('smtpFrom')   then sFrom   := ReadString('smtpFrom');
        CloseKey;
      end;
      Free;
    end;

    SendMailSmtp(sServer, sFrom, 'jpasteris@cegid.fr', 'JOB TO DO', 'Avant tâche');
    *)
    TBap := TOB.Create('µµµ', nil, -1);
    try
      {Récupération de tous les bap qui ont le statut à validé, mais qui n'ont pas d'autres étapes postérieures ainsi
       que les ligne de circuits correspondants à l'étape suivante}
      SQL := 'SELECT BAP_DATEMAIL, BAP_DATEMODIF, BAP_DATERELANCE1, BAP_DATERELANCE2, BAP_ECHEANCEBAP, ' +
             'BAP_MODIFICATEUR, BAP_NBJOUR, BAP_RELANCE1, BAP_RELANCE2, BAP_STATUTBAP, BAP_VISEUR, BAP_VISEUR1, ';
      SQL := SQL + 'BAP_VISEUR2, BAP_JOURNAL, BAP_EXERCICE, BAP_DATECOMPTABLE, BAP_NUMEROPIECE, BAP_NUMEROORDRE, ' +
                   'CCI_VISEUR1, CCI_VISEUR2, CCI_NBJOUR FROM CPBONSAPAYER C1 LEFT JOIN CPCIRCUIT ON ' +
                   'CCI_CIRCUITBAP = BAP_CIRCUITBAP AND CCI_NUMEROORDRE = BAP_NUMEROORDRE + 1 ';
      SQL := SQL + 'WHERE BAP_STATUTBAP = "' + sbap_Valide + '" AND NOT EXISTS (SELECT BAP_JOURNAL ' +
                   'FROM CPBONSAPAYER C2 WHERE C1.BAP_JOURNAL = C2.BAP_JOURNAL AND C1.BAP_EXERCICE = C2.BAP_EXERCICE AND ' +
                   'C1.BAP_DATECOMPTABLE = C2.BAP_DATECOMPTABLE AND C1.BAP_NUMEROPIECE = C2.BAP_NUMEROPIECE AND ' +
                   'C1.BAP_NUMEROORDRE < C2.BAP_NUMEROORDRE)';

      ObjLog.Add(SQL);
      TBap.LoadDetailFromSQL(SQL);
      for n := 0 to TBap.Detail.Count - 1 do begin
        T := TBap.Detail[n];
        oBap := TBonAPayer.Create(T, Length(Trim(T.GetString('CCI_VISEUR1'))) = 0);
        if Assigned(oBap) then FreeAndNil(oBap);
      end;
  //    SendMailSmtp(sServer, sFrom, 'jpasteris@cegid.fr', 'JOB IS DONE', 'Après tâche');
    finally
      FreeAndNil(TBap);
    end;
  except
    on E : Exception do begin
      Result := TraduireMemoire('Erreur lors de la tâche ') + psrv_Etape + ' (LanceTraitement)' + #13#13 + E.Message;
      ObjLog.Add('LanceTraitement : ' + E.Message);
    end;
  end;
end;

end.

