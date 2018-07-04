unit AglInitCommun;

interface

implementation

uses
  M3FP, Hent1, HCtrls,  Sysutils,ShellApi,Windows,
  variants,
  {$IFNDEF EAGLSERVER}
    {$IFNDEF ERADIO}
      AglInit,
    {$ENDIF !ERADIO}
  {$ENDIF !EAGLSERVER}
  {$IFDEF EAGLCLIENT}
    etablette,utob,
  {$ELSE  EAGLCLIENT}
    {$IFNDEF ERADIO}
      {$IFNDEF EAGLSERVER}
        tablette,
      {$ENDIF !EAGLSERVER}
    {$ENDIF !ERADIO}
    {$IFNDEF DBXPRESS}dbtables{BDE},{$ELSE}uDbxDataSet,{$ENDIF}
  {$ENDIF EAGLCLIENT}
  ParamSoc;

function AGLSerieS3 ( parms: array of variant; nb: integer ) : variant ;
begin
Result:=0 ;
if V_PGI.LaSerie=S3 then Result:=1 ;
end;

function AGLSerieLine ( parms: array of variant; nb: integer ) : variant ;
begin

	Result := 0;

//uniquement en line
//	Result := 1;


end;


{***********A.G.L.***********************************************
Auteur  ...... : AV
Créé le ...... : 26/01/2001
Modifié le ... : 26/01/2001
Description .. : Utilisation des paramètres société dans les scripts
Mots clefs ... : SCRIPT;PARAMSOC;GETPARAMSOC
*****************************************************************}
function AGLGetParamSoc ( parms: array of variant; nb: integer ) : variant ;
var ch : string;
begin
ch := string(Parms[0]) ;
result := GetParamSoc (ch);
end;

{***********A.G.L.***********************************************
Auteur  ...... : AV
Créé le ...... : 26/01/2001
Modifié le ... : 26/01/2001
Description .. : Utilisation des paramètres société dans les scripts
Mots clefs ... : SCRIPT;PARAMSOC;GETPARAMSOC
*****************************************************************}
function AGLGetParamSocSecur ( parms: array of variant; nb: integer ) : variant ;
var ch : string;
begin
ch := string(Parms[0]) ;
result := GetParamSocSecur (ch,Parms[1]);
end;

{$IFNDEF EAGLSERVER}
{$IFNDEF ERADIO}
{***********A.G.L.***********************************************
Auteur  ...... : SD
Créé le ...... : 15/03/2001
Modifié le ... :
Description .. : gestion manuelle du paramtable pour les familles hierarchique
                 utilisé lors de la création des familles dans les scripts
Mots clefs ... : SCRIPT;PARAMTABLERAMSOC
*****************************************************************}
function AGLParamTable ( parms: array of variant; nb: integer ) : variant ;
var ch : string;
begin
ch := string(Parms[0]) ;
ParamTable(ch,taCreat,0,Nil,3) ;
end;
{$ENDIF !ERADIO}
{$ENDIF !EAGLSERVER}

{***********A.G.L.***********************************************
Auteur  ...... : SD
Créé le ...... : 15/03/2001
Modifié le ... :
Description .. : Rechargement manuel d'une tablette dans un script
Mots clefs ... : SCRIPT;PARAMTABLERAMSOC
*****************************************************************}
function AGLAvertirTable ( parms: array of variant; nb: integer ) : variant ;
var ch : string;
begin
ch := string(Parms[0]) ;
AvertirTable (ch) ;
end;

{$IFNDEF EAGLSERVER}
{$IFNDEF ERADIO}
////////////////// transmission du mode action de la fiche pour ses sous-fiche ////////////////////
function AGLActionForm (parms: array of variant; nb: integer) : variant ;
begin
result := ActionToString(TActionFiche(parms[0]));
end;
{$ENDIF !ERADIO}
{$ENDIF !EAGLSERVER}

Function  AGLRTRecupMessagerie( parms: array of variant; nb: integer ): variant ;
begin
result:='';
if V_PGI.MailMethod = mmOUTLOOK then
   result:='Outlook';
if V_PGI.MailMethod = mmNotes then
   result:='Notes';
end;

function AGLRTSQLField( parms: array of variant; nb: integer ):variant ;
var QQ:TQuery;
begin
Result:='';
QQ:=TQuery(longint(parms[0]));
if QQ<> nil then Result:= QQ.Findfield(string(parms[1])).asstring;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Thierry Petetin
Créé le ...... : 28/06/2002
Modifié le ... :   /  /
Description .. : 'Accès' à V_PGI depuis le script
Mots clefs ... :
*****************************************************************}
function AGL_V_PGI(Parms: Array of variant; Nb: Integer ): Variant;
begin
  if UpperCase(Parms[0]) = 'CONTEXTE' then
  begin
    if UpperCase(Parms[1]) = 'CTXGPAO' then Result := (ctxGPAO in V_PGI.PGIContexte) else
    if UpperCase(Parms[1]) = 'CTXFO' then Result := (ctxFO in V_PGI.PGIContexte) else
    if UpperCase(Parms[1]) = 'CTXCHR' then Result := (ctxChr in V_PGI.PGIContexte) else
    if UpperCase(Parms[1]) = 'CTXMODE' then Result := (ctxMode in V_PGI.PGIContexte)
    else Result:=False;
  end
//GP_20080110_DKZ_AGL14241
  else if UpperCase(Parms[0]) = 'USER' then
  begin
    Result := V_PGI.User;
  end
//GP_20080110_DKZ_AGL14241
  else
    Result := null;
end;
Function AGLReadToken( parms: array of variant; nb: integer ) : variant ;
var  tmp, tmp1 : string;
    indice,ii : integer;
begin
    // fct qui recoit l'indice + le champ. Recherche dans la chaine
    // la xième valeur de celle-ci et la retourne.
Indice := Integer(Parms[0]) ;
tmp := String(Parms[1]) ;
result:='';
ii:=0;
Repeat
    tmp1 := ReadTokenSt(tmp);
    if (ii= indice) then begin result:=tmp1; break; end; ;
    Inc (ii);
    until (tmp1 = '') ;

end;

{***********A.G.L.***********************************************
Auteur  ...... : BRUN Matthieu
Créé le ...... : 17/01/2007
Modifié le ... :   /  /
Description .. : Lancement de n'importe quoi
Mots clefs ... :
*****************************************************************}
function AGLShellExecute( parms: array of variant; nb: integer ) : variant ;
var
  v : integer ;
  commande, parametre, startdir : Pchar ;
  Visible : boolean ;
begin
   if Length(parms) = 4 then
   begin
      commande := Pchar(String(parms[0]));
      parametre := Pchar(String(parms[1]));
      startdir := Pchar(String(parms[2]));
      Visible := boolean(parms[3]);
      if Visible then
         v := SW_SHOWNORMAL
      else
         v := SW_HIDE ;
      ShellExecute(0,'open',Pchar(Commande),Pchar(Parametre), Pchar(StartDir),v);
   end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BRUN Matthieu
Créé le ...... : 17/01/2007
Modifié le ... :   /  /
Description .. : Lancement d'une url
Mots clefs ... :
*****************************************************************}
function AGLExecuteUrl( parms: array of variant; nb: integer ) : variant ;
begin
   AGLShellExecute([parms[0],'','',true],nb);
end;


Initialization
RegisterAglFunc( 'SERIES3', FALSE , 0, AGLSerieS3);
RegisterAglFunc( 'SerieLine', FALSE , 0, AGLSerieLine);
RegisterAglFunc( 'AGLGetParamSoc', False, 1, AGLGetParamSoc );
RegisterAglFunc( 'AGLGetParamSocSecur', False, 2, AGLGetParamSocSecur );
{$IFNDEF EAGLSERVER}
  {$IFNDEF ERADIO}
    RegisterAglFunc( 'AGLParamTable', False, 1, AGLParamTable );
    RegisterAglFunc( 'AGLActionForm', FALSE, 1, AGLActionForm );
  {$ENDIF !ERADIO}
{$ENDIF !EAGLSERVER}
RegisterAglFunc( 'AGLAvertirTable', False, 1, AGLAvertirTable );
RegisterAglFunc( 'RTRecupMessagerie', FALSE , 1, AGLRTRecupMessagerie);
RegisterAglFunc( 'AGL_V_PGI', False, 2, Agl_V_Pgi);
RegisterAglFunc( 'RTSQLField', False,2,AGLRTSQLField) ;
RegisterAglFunc( 'ReadToken', FALSE , 2, AGLReadToken);
RegisterAglFunc( 'ShellExecute', FALSE , 4, AGLShellExecute);
RegisterAglFunc( 'ExecuteUrl', FALSE , 1, AGLExecuteUrl);

end.


