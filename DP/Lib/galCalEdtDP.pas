unit galCalEdtDP;

interface

uses
     Forms,SysUtils, HEnt1,
{$IFNDEF EAGLCLIENT}
     {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}
{$ENDIF}
     controls, HCtrls, Classes, UTOB, stdctrls, windows;

function CalcOLEEtatDP(sf,sp : string) : variant ;


///////////// IMPLEMENTATION /////////////
implementation

uses CalcOleAffaire
     ,dpPublicationCD        // $$$ JP 06/10/04 - pour accéder aux données de l'assistant publication CD (pour jaquette)
     ;


function CalcOLEEtatDP (sf,sp:string):variant;
var Q: TQuery;
begin
     Result := '';

  // nom du domaine en fonction du code
  if sf='GETNOMDOMAINE' then
    begin
    if sp='A' then
      Result := 'Affaire'
    else if sp='B' then
      Result := 'BTP'
    else if sp='C' then
      Result := 'Comptabilité'
    else if sp='D' then
      Result := 'Dossier permanent'
    else if sp='E' then
      Result := 'E_business'
    else if sp='K' then
      Result := 'Net service'
    else if sp='F' then
      Result := 'Fiscalité personnelle'
    else if sp='G' then
      Result := 'Gestion commerciale'
    else if sp='H' then
      Result := 'CHR'
    else if sp='I' then
      Result := 'Immobilisations'
    else if sp='J' then
      Result := 'Juridique'
    else if sp='K' then
      Result := 'Transport'
    else if sp='M' then
      Result := 'MCC'
    else if sp='N' then
      Result := 'Note de frais'
    else if sp='O' then
      Result := 'Conso'
    else if sp='P' then
      Result := 'Paie / GRH'
    else if sp='Q' then
      Result := 'Synaptique'
    else if sp='R' then
      Result := 'Prospect'
    else if sp='S' then
      Result := 'Système'
    else if sp='T' then
      Result := 'CTNS'
    else if sp='V' then
      Result := 'TDI'
    else if sp='X' then
      Result := 'Sysload'
    else if sp='Y' then
      Result := 'Commun';
    end

  // SO_VERSIONBASE
  else if sf='GETVERSIONBASE' then
    begin
    Q := OpenSQL('SELECT SO_VERSIONBASE FROM SOCIETE', True);
    if not Q.Eof then Result := Q.FindField('SO_VERSIONBASE').AsString;
    Ferme(Q);
    end

  // $$$ JP 06/10/04 - fonctions pour la jaquette CD GED
  // $$$ JP 03/12/04 - gestion casse du code fonction, et test d'égalité (et pas "différent")
  else if UpperCase (sf) = 'CDVALUE' then //if sf<>'CDValue' then
  begin
       if (FPublicationCD <> nil) and (FPublicationCD.m_TOBData <> nil) then
       begin
            sp := StringReplace (sp, '@', '', [rfReplaceAll, rfIgnoreCase]);

            // Si on demande la liste, on considère que c'est le résumé détaillé: armoire + classeur
            if Copy (sp, 1, 10) = 'PUB_RESUME' then
                 Result := FPublicationCD.GetResume (StrToInt (Copy (sp, 11, Length (sp)-10)))
            else
                 Result := FpublicationCD.m_TOBData.GetString (sp);
       end
       else
           Result := sp;
  end

  // AUTRES FONCTIONS NECESSAIRES : CELLES DE LA GI !!
  else
    Result := AFCalcOLEEtat(sf, sp);
end;

end.
