{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 12/06/2008
Modifié le ... :   /  /
Description .. : Utilitaires Oracle
Mots clefs ... : ORACLE
*****************************************************************}
Unit UtilOracle ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
{$IFDEF VER150} {D7}
  Variants,
{$ENDIF}
{$IFNDEF EAGLCLIENT}
     db,fe_main,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     majtable,
{$else}
     UtileAGL,
{$ENDIF}
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     Grids,
     HEnt1,galSystem  ,
     PGIAppli,
     HMsgBox, HTB97,HPanel, HsysMenu,
     UTOF,
     UTOB;

function okJeuCarOra : boolean;

Implementation

Const NomExec = 'PGIMajVer.exe';

{JS1 12/06/2008 : Test du jeu de caractères Oracle
si KO, on bloque la maj de base
si on ne fait pas cela, les bases sont corrompues après traitement
Fonction ecrite à partir du script transmis par C Lavenne}
Function okJeuCarOra : Boolean;
var
  cset   : string;
  eansi  : integer;
  e850   : integer;
  lQMD   : TQuery ;
  iPos   : integer;
Begin
  result  := true;
  cset    := '';
  eansi   :=0;
  e850    :=0;  

  lQMD   := OpenSQL('@@select * from nls_database_parameters where parameter =' + '''NLS_CHARACTERSET''', True ) ;
  if not lQMD.Eof then
    cset := lQMD.FindField('VALUE').AsString ;
  Ferme( lQMD ) ;
  if cset = '' then Exit ;

  iPos := pos(cset,'WE8MSWIN1252,WE8ISO8859P1,WE8ISO8859P15');

  if iPos > 0 then
  begin
    lQMD   := OpenSQL('@@select count(*) AS EANSI from menu where instr(mn_libelle, chr(233))>0', True ) ;
    if not lQMD.Eof then
      eansi := lQMD.FindField('EANSI').AsInteger ;
    Ferme( lQMD ) ;

    lQMD   := OpenSQL('@@select count(*) AS E850  from menu where instr(mn_libelle, chr(130))>0', True ) ;
    if not lQMD.Eof then
      e850 := lQMD.FindField('E850').AsInteger ;
    Ferme( lQMD ) ;
    if (eansi + e850) = 0 then exit;

    if (eansi*100/(eansi+e850)) < 97 then
    begin
      PGIBOX ( '        Votre configuration ORACLE est incorrecte.'
      +chr(13)+'Veuillez contacter CEGID pour mettre à jour votre configuration.'
      +chr(13)+'   ORA-20000 : CEGID, Jeu de caractères incorrect ('+ IntToStr(eansi) +'/'+ IntToStr(e850) +').' , 'Attention');
      Result := False;
    end;
  end
  else Result := True;

End;


end.
