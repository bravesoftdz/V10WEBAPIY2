unit AGLInitDpJur;

interface
uses
   UTOB;
type PTOB = ^TOB;

function AGLAjoutOutlook(parms : array of variant; nb : integer) : integer;
function AGLSetValChamp( parms : array of variant; nb : integer) : integer;

//////////// IMPLEMENTATION /////////////

implementation

uses
   M3FP,

   {$IFDEF VER150}
   Variants,
   {$ENDIF}

   DpJurOutils, hctrls;

{*****************************************************************
Auteur ....... : ??
Date ......... : ??/??/??
Procédure .... : AGLAjoutOutlook
Description .. :
Paramètres ... : Tableau de variants : 0 famille évènement, 1 n° évènement
                 Nombre d'arguments
*****************************************************************}
function AGLAjoutOutlook(parms : array of variant; nb : integer) : integer;
begin
   result := 0;
   if (string(parms[0])<>'REU') and
      (string(parms[0])<>'TAC') and
      (string(parms[0])<>'MSG') then
      exit;
   result := AjoutOutlook( string(parms[0]), string(parms[1]));
end;

{*****************************************************************
Auteur ....... : BM
Date ......... : 10/12/02
Procédure .... : AGLSetValChamp
Description .. : Préparation pour la mise à jour du champ d'un enregistrement dans une table
Paramètres ... : Tableau de variants : 0 table, 1 nom = valeur du champ, 2 à n clé de l'enregistrement
                 Nombre d'arguments
*****************************************************************}
function AGLSetValChamp( parms : array of variant; nb : integer) : integer;
var
   sTable_l, sChamp_l, sChampVal_l, sValeur_l : string;
   tsCle_l, tsCleVal_l : array of string;
   nParmInd_l, nCleInd_l : integer;
begin
   sTable_l    := VarToStr(parms[0]);
   sValeur_l := VarToStr(parms[1]);
   sChamp_l    := ReadTokenPipe( sValeur_l, '=');
   sChampVal_l := ReadTokenSt( sValeur_l );

   SetLength( tsCle_l, nb - 2 );
   SetLength( tsCleVal_l, nb - 2 );
   nCleInd_l := 0;
   for nParmInd_l := 2 to nb - 1 do
   begin
      sValeur_l := VarToStr(parms[nParmInd_l]);
      tsCle_l[nCleInd_l] := ReadTokenPipe( sValeur_l, '=');
      tsCleVal_l[nCleInd_l] := ReadTokenSt( sValeur_l );
      inc(nCleInd_l);
   end;
   SetValChamp( sTable_l, tsCle_l, tsCleVal_l, sChamp_l, sChampVal_l );
   result := 1;
end;


// $$$ JP 10/08/2004 - en provenance de AGLInitJur
procedure AGLLanceWord (parms:array of variant; nb:integer);
begin
     LanceWord (string (Parms[0]), string (Parms[1]));
end;

{*****************************************************************
Auteur ....... : ??
Date ......... : ??/??/??
Procédure .... : initM3Commun
Description .. :
Paramètres ... : 
*****************************************************************}
procedure initM3Commun();
begin
//   RegisterAglFunc( 'AjoutOutlook', FALSE , 2, AGLAjoutOutlook);
//   RegisterAglFunc( 'AGLSetValChamp', FALSE , 5, AGLSetValChamp);
end;

//////////// IMPLEMENTATION //////////////

Initialization
              RegisterAglProc ('LanceWord', FALSE , 2, AGLLanceWord);
              initM3Commun ();

finalization

end.
