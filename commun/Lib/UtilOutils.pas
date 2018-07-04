{***********UNITE*************************************************
Auteur  ...... : MNG
Créé le ...... : 08/10/2009
Pour ..........: FQ;034;16583
Description .. : deplacement des fonctions pour qu'elles puissent être utilisées des appli Gx
Suite ........ : appel dans PgiMajVer et menu outils Gx
*****************************************************************}

unit UtilOutils;

interface

{ TS : 25/11/8 variable globale => fonction }
function IsDossierPCL: Boolean;
function ExecuteSQLContOnExcept(const Sql: WideString): Integer;
function usDateTime_ (LaDate : TDateTime ) : string;

implementation
uses hent1 , hctrls, sysutils, hmsgbox, cbptrace, DateUtils;

function IsDossierPCL: Boolean;
begin
  {$if Defined(MAJPCL)}
    Result := V_PGI.DossierPCL;
  {$elseif Defined(MAJBOB)}
    Result := V_PGI.BobDeMiseAJour <> '';
  {$else}
    Result := False;
  {$ifend}
end;

//===============================================================
//
// ExecuteSQL Continue On Exception en mode SAV
// 05/04/07. ALR & LA
//===============================================================
function ExecuteSQLContOnExcept(const Sql: WideString): Integer;
begin
  result := -1;

  try
    result := ExecuteSQL(sql);
  except
    on e: exception do
    begin
      if  (v_pgi.SAV) then
        PgiError(E.message);
      // nb: si pas  sav ... raise=> pgierror + haut
      trace.traceError('SQL','Exception '+e.message );
      trace.traceError('SQL',sql);
      // ALR 05/04/07
      if  (not v_pgi.SAV) then
        raise;
    end;
  end;
end;

function usDateTime_ (LaDate : TDateTime ) : string;
var aa,mm,dd,hh,min,SS,mil : Word;
begin
	Result := USDATETIME(LaDate);
end;

end.
