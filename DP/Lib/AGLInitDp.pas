unit AGLInitDp;
// certaines fonctions issues de AGLInitJur (src\juridique\lib)

interface
uses HEnt1, HCtrls, M3FP,
{$IFDEF EAGLCLIENT}
     MaineAGL,
{$ELSE}
     Fe_Main,
{$ENDIF}
     Web, MailOL, AGLInit;

//////////// IMPLEMENTATION /////////////
implementation

uses dpOutils, DpJurOutils
{$IFDEF BUREAU}
     ,entDP, sysutils, Controls, hmsgbox, galsystem
{$ENDIF}
     ;


procedure AGLFonctionInternet(parms : array of variant; nb: integer );
begin
  if String(Parms[0])='SITEWEB' then
    LanceWeb(String(Parms[1]),FALSE)
  else if String(Parms[0])='EMAIL' then
    SendMail('',String(Parms[1]),'',nil,'',FALSE);
end;

{$IFDEF BUREAU}
procedure AGLFonctionCTI (parms:array of variant; nb:integer);
var
   strTelephone   :string;
begin
     if string (Parms [0]) = 'MAKECALL' then
     begin
          strTelephone := Trim (string (Parms [1]));
          if strTelephone <> '' then
             if PgiAsk ('Appeler le ' + strTelephone + ' ?') = mrYes then
             begin
                  VH_DP.ctiAlerte.Close;
                  VH_DP.ctiAlerte.MakeCall (strTelephone);
             end;
     end;
end;

Function AGLAffecteGroupe(parms:array of variant; nb:integer): variant;
var
   StrDossier, strGroupeConf :string;
begin
     if Length(Parms) = 2 then
     begin
        StrDossier := string(parms[0]) ;
        strGroupeConf := string(parms[1]) ;
        Result := AffecteGroupe(StrDossier,strGroupeConf);
     end;
end;
{$ENDIF}


Function AGLActionToString( parms: array of variant; nb: integer ): variant;
begin
  Result := ActionToString(Parms[0]);
end;


Function AGLStringToAction( parms: array of variant; nb: integer ): variant;
begin
  Result := StringToAction(Parms[0]);
end;


Function AGLReadToken( parms: array of variant; nb: integer ) : variant ;
// fct qui recoit l'indice + le champ. Recherche dans la chaine
// la xième Valeur de celle-ci et la retourne.
var tmp, tmp1 : string;
    indice,ii : integer;
begin
  Indice := Integer(Parms[0]) ;
  tmp := String(Parms[1]) ;
  Result:='';
  ii:=0;
  Repeat
    tmp1 := ReadTokenSt(tmp);
    if (ii= indice) then begin result:=tmp1; break; end; ;
    Inc (ii);
  until (tmp1 = '') ;
end;


// MD
procedure AGLDispatchTT (parms: array of variant; nb: integer ) ;
begin
  if Assigned(V_PGI.DispatchTT) then V_PGI.DispatchTT(Parms[0],StringToAction(parms[1]),Parms[2],Parms[3],Parms[4]) ;
end;


procedure AGLDupliquerElement( parms: array of variant; nb: integer );
var
  Table,sFiche,sCle,sDom : string;
begin
  Table := String(parms[0]);
  sDom := 'JUR';
  //NCX & XP 15/11/00 : permet l'affichage en mode création lors de la duplication.
  // Pour JUBIBACTION ou 'JUBIBRUBRIQUE', voir AGLDupliquerElement dans AGLInitJur
  if True then
    begin
    sCle := DupliquerElement(String(parms[0]),String(parms[1]),String(parms[2]));
    if Table='ANNUAIRE' then
      // ####existe aussi DupliquerElementAnn ...
      begin sFiche := 'ANNUAIRE'; sDom:='YY'; end
    else if Table='JUEVENEMENT' then
      sFiche := 'EVENEMENT';
    AGLLanceFiche ( sDom, sFiche, sCle, sCle, '' ) ;
    end ;
end;


initialization
  RegisterAglProc( 'FonctionInternet', FALSE , 1, AGLFonctionInternet);
  RegisterAglFunc( 'ActionToString',False,1,AGLActionToString);
  RegisterAglFunc( 'StringToAction',False,1,AGLStringToAction);
  RegisterAglFunc( 'ReadToken', FALSE , 2, AGLReadToken);
  // MD
  RegisterAglProc( 'DispatchTT', FALSE , 4, AGLDispatchTT);
  RegisterAglProc( 'DupliquerElement', FALSE , 4, AGLDupliquerElement);

  // $$$ JP 10/06/2005 - lancement appel CTI depuis une fiche décla
{$IFDEF BUREAU}
  RegisterAglProc ( 'FonctionCTI', FALSE, 2, AGLFonctionCTI);
  // MB : 23/10/2007 - Ajout de l'affectation d'un dossier a un groupe de travail.
  RegisterAglFunc( 'AffecteGroupe', FALSE , 2, AGLAffecteGroupe);
{$ENDIF}

finalization


end.
