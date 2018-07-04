{***********UNITE*************************************************
Auteur  ...... : SBO
Créé le ...... : 01/12/2003
Modifié le ... :   /  /
Description .. :
Suite ........ : Fonctions utilisées dans la compta pour le synchro avec la
Suite ........ : tréso
Mots clefs ... : TRESO;TRSYNCHRO;TRESOSYNCHRO
*****************************************************************}
unit ULibTrSynchro;

interface

uses
{$IFDEF EAGLCLIENT}
{$ELSE}
     DB,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ENDIF}
     SysUtils,    // IntToStr
     Utob,        // TOB
     Hent1,       // TActionFiche
     HMsgBox,     // PGIBox
     HCtrls,      // USDateTime
     SaisUtil,    // EstCptSyn
     uLibEcriture,// TInfoEcriture
     ParamSoc;    // GetParamSocSecur

// Fonctions de gestion du lien Compta-Tréso version avec des TOBs
Function  ExisteReferencesTresoTOB( vEcr : TOB ) : Boolean ;
Function  DetruitReferencesTresoTOB( vEcr : TOB ) : Integer ;
Function  DetruitEcritureTresoTOB( vEcr : TOB ) : Integer ;
Function  MajTresoEcritureTOB( vEcr : TOB ; vAction : TActionFiche ; vInfo : TInfoEcriture = nil ) : Boolean ;
Function  MajTresoSupprEcritureTOB( vEcr : TOB ) : Boolean ;
{JP 17/11/04 : FQ 14972 : maj de E_TRESOSYNCHROo depuis modification des entêtes de pièces,
               unité dans laquelle on a aussi les Tob journaux et généraux. Cette fonction
               effectue le même traitement que EstCptSyn mais avec des Tobs}
procedure MajE_TRESOSYNCHROTob(vEcr, vJal, vGen : TOB);
{JP 12/04/05 : Destruction d'une écriture à partir d'un RMVT}
procedure DetruitEcritureTresoRMVT(Mvt : RMVT);
{JP 12/04/05 : Renvoie vide si E_TRESOSYNCHRO n'a pas besoin d'être modifié, sinon renvoie
               la partie de reuqête concernant E_TRESOSYNCHRO}
function  MajETresoSynchro(Mvt : RMVT) : string;

// SBO 16/01/2005 : Equivalent EstCptSyn sans requête
function  EstCptSynTob( vTobEcr : Tob ; vInfo : TInfoEcriture = nil ) : Boolean ;

implementation

uses
  SaisComm,
  {$IFDEF NOVH}
  ULibCpContexte,
  {$ENDIF NOVH}
  Commun, UtilPgi, {EstMultiSoc}
  Constantes
  , Ent1
  , CPProcGen
  ;



{***********A.G.L.***********************************************
Auteur  ...... : SBO
Créé le ...... : 27/11/2003
Modifié le ... :   /  /
Description .. :
Suite ........ : Traitement de maj de la table TRECRITURE de la tréso à la
Suite ........ : création ou la modification d'une écriture compta
Mots clefs ... :
*****************************************************************}
Function  MajTresoEcritureTOB( vEcr : TOB ; vAction : TActionFiche ; vInfo : TInfoEcriture = nil  ) : Boolean ;
var lBoSyn : Boolean ;
begin

  Result := True ;

  // Traitements valables uniquement pour les écritures normales
  if ( vEcr = nil ) or ( vEcr.GetValue('E_QUALIFPIECE') <> 'N' ) or
     (( vEcr.GetString('E_ECRANOUVEAU') <> 'H' ) and
      ( vEcr.GetString('E_ECRANOUVEAU') <> 'N' )) or
     (( vEcr.GetString('E_ECRANOUVEAU') = 'H' ) and
      {$IFDEF NOVH}
      ((GetParamSocSecur('SO_EXOV8','') <> '') and (vEcr.GetDateTime('E_DATECOMPTABLE') < TCPContexte.GetCurrent.Exercice.ExoV8.Deb) )) then begin
      {$ELSE}
      ((GetParamSocSecur('SO_EXOV8','') <> '') and (vEcr.GetDateTime('E_DATECOMPTABLE') < VH^.ExoV8.Deb) )) then begin
      {$ENDIF NOVH}
    {JP 17/04/07 : c'est plus prudent !!}
    if Assigned(vEcr) then vEcr.PutValue('E_TRESOSYNCHRO', 'RIE');
    Exit;
  end;

  // ------------------
  // ---- CREATION ----
  // ------------------
  if (vAction = taCreat) or (vAction = taCreatOne) or (vAction = taCreatEnSerie) then
  begin

    // -> Notifier la création à la tréso si la ligne porte sur un compte de banque
    if Assigned( vInfo )
      then lBoSyn := EstCptSynTob( vEcr, vInfo )
      else lBoSyn := EstCptSyn( vEcr.GetString('E_GENERAL'), vEcr.GetString('E_JOURNAL'), vEcr.GetString('E_NATUREPIECE') <> 'ECC' ) ;

    if lBoSyn then
      vEcr.PutValue('E_TRESOSYNCHRO', 'CRE' ) ;

  end ;

  // ----------------------
  // ---- MODIFICATION ----
  // ----------------------
  if (vAction = taModif) or (vAction = taModifEnSerie) then
  begin

    // Effacer la référence tréso si elle existe
    if vEcr.GetValue('E_NUMLIGNE') = 1  then
      DetruitReferencesTresoTOB( vEcr ) ;

    // MAJ de champ de synchro tréso : E_TRESOSYNCHRO
    if Assigned( vInfo )
      then lBoSyn := EstCptSynTob( vEcr, vInfo )
      else lBoSyn := EstCptSyn( vEcr.GetString('E_GENERAL'), vEcr.GetString('E_JOURNAL'), vEcr.GetString('E_NATUREPIECE') <> 'ECC' ) ;

    if lBoSyn
      then vEcr.PutValue('E_TRESOSYNCHRO', 'CRE' )
      else vEcr.PutValue('E_TRESOSYNCHRO', 'RIE' ) ; // On n'est plus sur un compte de banque...

  end ;

end ;

{SBO 01/12/03 : Vérifie une écriture de tréso existe avec les références de la pièce (quelque soit la ligne ou le compte...)
 JP 09/08/06 : gestion du multi sociétés en Tréso  : je pars du principe que la pièce comptable
               appartient au dossier V_PGI.SchemaName !!}
Function  ExisteReferencesTresoTOB( vEcr : TOB ) : Boolean ;
var
  lStSQL : String ;
begin
  {La trésorerie ne figure que dans une base, paramétrée dans le paramsoc SO_TRBASETRESO}
  lStSQL := 'SELECT TE_CPNUMLIGNE FROM ' + GetTableDossier(GetParamSocSecur('SO_TRBASETRESO', ''), 'TRECRITURE') + ' WHERE '
            {$IFDEF NOVH}
            + 'TE_NODOSSIER = "' + TCPContexte.GetCurrent.TresoNoDossier + '" AND ' {JP 09/08/06 : TRESO MS}
            {$ELSE}
            + 'TE_NODOSSIER = "' + VH^.TresoNoDossier + '" AND ' {JP 09/08/06 : TRESO MS}
            {$ENDIF NOVH}
            + 'TE_JOURNAL = "' + vEcr.GetString('E_JOURNAL') + '" AND '
            + 'TE_EXERCICE = "' + vEcr.GetString('E_EXERCICE') + '" AND '
            + 'TE_NUMEROPIECE=' + vEcr.GetString('E_NUMEROPIECE');
  if (vEcr.GetString('E_MODESAISIE') = 'BOR') or (vEcr.GetString('E_MODESAISIE') = 'LIB') then
    lStSQL := lStSQL + 'AND TE_NUMTRANSAC LIKE "' + TRANSACIMPORT + vEcr.GetString('E_SOCIETE')
                     + vEcr.GetString('E_JOURNAL') + Copy(vEcr.GetString('E_PERIODE'), 3, 4) + '%" ';
  Result := ExisteSQL( lStSQL ) ;
end ;


{SBO 01/12/03 : Supprime les écritures de tréso référençant la pièce contenu l'ecriture est passé en paramètre.
 JP 09/08/06 : gestion du multi sociétés en Tréso  : je pars du principe que la pièce comptable
               appartient au dossier V_PGI.SchemaName !!
 JP 18/06/07 : FQ 20766 : ajout d'un espace avant le AND TE_NUMTRANSA}
Function  DetruitReferencesTresoTOB( vEcr : TOB ) : Integer ;
var lStSQL : String ;
begin
  {La trésorerie ne figure que dans une base, paramétrée dans le paramsoc SO_TRBASETRESO}
  lStSQL := 'DELETE FROM ' + GetTableDossier(GetParamSocSecur('SO_TRBASETRESO', ''), 'TRECRITURE') + ' WHERE '
            {$IFDEF NOVH}
            + 'TE_NODOSSIER = "' + TCPContexte.GetCurrent.TresoNoDossier + '" AND ' {JP 09/08/06 : TRESO MS}
            {$ELSE}
            + 'TE_NODOSSIER = "' + VH^.TresoNoDossier + '" AND ' {JP 09/08/06 : TRESO MS}
            {$ENDIF NOVH}
            + 'TE_JOURNAL = "' + vEcr.GetString('E_JOURNAL') + '" AND '
            + 'TE_EXERCICE = "' + vEcr.GetString('E_EXERCICE') + '" AND '
            + 'TE_NUMEROPIECE=' + vEcr.GetString('E_NUMEROPIECE');
  if (vEcr.GetString('E_MODESAISIE') = 'BOR') or (vEcr.GetString('E_MODESAISIE') = 'LIB') then
    {18/06/07 : FQ 20766 : ajout d'un espace avant le AND TE_NUMTRANSA}
    lStSQL := lStSQL + ' AND TE_NUMTRANSAC LIKE "' + TRANSACIMPORT + vEcr.GetString('E_SOCIETE')
                     + vEcr.GetString('E_JOURNAL') + Copy(vEcr.GetString('E_PERIODE'), 3, 4) + '%" ';
  Result := ExecuteSQL( lStSQL ) ;
end ;


{SBO 28/11/03 : Supprime les écritures de tréso référençant l'ecriture contenu dans la Tob
                Equivalent de DetruitEcritureTreso et Equivalent de DetruitEcritureTresoRMVT mais avec une TOB
 JP 09/08/06 : gestion du multi sociétés en Tréso  : je pars du principe que la pièce comptable
               appartient au dossier V_PGI.SchemaName !!}
Function  DetruitEcritureTresoTOB( vEcr : TOB ) : Integer ;
var
  lStSQL : String ;
begin
  {La trésorerie ne figure que dans une base, paramétrée dans le paramsoc SO_TRBASETRESO}
  lStSQL := 'DELETE FROM ' + GetTableDossier(GetParamSocSecur('SO_TRBASETRESO', ''), 'TRECRITURE') + ' WHERE '
            {$IFDEF NOVH}
            + 'TE_NODOSSIER = "' + TCPContexte.GetCurrent.TresoNoDossier + '" AND ' {JP 09/08/06 : TRESO MS}
            {$ELSE}
            + 'TE_NODOSSIER = "' + VH^.TresoNoDossier + '" AND ' {JP 09/08/06 : TRESO MS}
            {$ENDIF NOVH}
            + 'TE_JOURNAL="' + vEcr.GetString('E_JOURNAL') + '" AND '
            + 'TE_EXERCICE = "' + vEcr.GetString('E_EXERCICE') + '" AND '
            + 'TE_NUMEROPIECE=' + vEcr.GetString('E_NUMEROPIECE') + ' AND '
            + 'TE_CPNUMLIGNE=' + vEcr.GetString('E_NUMLIGNE')  + ' AND '
            + 'TE_NUMECHE=' + vEcr.GetString('E_NUMECHE')  + ' ' ;
  if (vEcr.GetString('E_MODESAISIE') = 'BOR') or (vEcr.GetString('E_MODESAISIE') = 'LIB') then
    lStSQL := lStSQL + 'AND TE_NUMTRANSAC LIKE "' + TRANSACIMPORT + vEcr.GetString('E_SOCIETE')
                     + vEcr.GetString('E_JOURNAL') + Copy(vEcr.GetString('E_PERIODE'), 3, 4) + '%" ';
  Result := ExecuteSQL( lStSQL ) ;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : SBO
Créé le ...... : 27/11/2003
Modifié le ... :   /  /
Description .. :
Suite ........ : Traitement de maj de la table TRECRITURE de la tréso à la
Suite ........ : suppression d'une ecriture compta
Mots clefs ... :
*****************************************************************}
Function  MajTresoSupprEcritureTOB( vEcr : TOB ) : Boolean ;
begin

  Result := True ;

  // Traitements valables uniquement pour les écritures normales
  if ( vEcr = nil ) or ( vEcr.GetValue('E_QUALIFPIECE') <> 'N' ) or
     (( vEcr.GetString('E_ECRANOUVEAU') <> 'H' ) and
      ( vEcr.GetString('E_ECRANOUVEAU') <> 'N' )) or
     (( vEcr.GetString('E_ECRANOUVEAU') = 'H' ) and
      {$IFDEF NOVH}
      ((GetParamSocSecur('SO_EXOV8','') <> '') and (vEcr.GetDateTime('E_DATECOMPTABLE') < TCPContexte.GetCurrent.Exercice.ExoV8.Deb) )) then
      {$ELSE}
      ((GetParamSocSecur('SO_EXOV8','') <> '') and (vEcr.GetDateTime('E_DATECOMPTABLE') < VH^.ExoV8.Deb) )) then
      {$ENDIF NOVH}
    Exit ;

  // ---------------------
  // ---- SUPPRESSION ----
  // ---------------------
  // Effacer les références tréso poru toute la pièce comptable en 1 fois.
  // Traitement effectuer sur la première ligne uniquement.
  if vEcr.GetValue('E_NUMLIGNE') = 1  then
      DetruitReferencesTresoTOB( vEcr ) ;

  // Remarque : le statut E_TRESOSYNCHRO de la pièce détruite est placé à RIE dans le module SupprEcr.pas

end ;

{JP 17/11/04 : FQ 14972 : maj de E_TRESOSYNCHROo depuis modification des entêtes de pièces,
               unité dans laquelle on a aussi les Tob journaux et généraux. Cette fonction
               effectue le même traitement que EstCptSyn mais avec des Tobs}
{---------------------------------------------------------------------------------------}
procedure MajE_TRESOSYNCHROTob(vEcr, vJal, vGen : TOB);
{---------------------------------------------------------------------------------------}
var
  Cpte  : string;
  Jnal  : string;
  n     : Integer;

    {----------------------------------------------------------------------}
    function Synchronisable(NonEcc : Boolean) : Boolean;
    {----------------------------------------------------------------------}
    begin
      {La tob général n'est pas forcément assignée}
      if (vGen = nil) or (vJal = nil) then begin
        Result := EstCptSyn(Cpte, Jnal, NonEcc);
        Exit;
      end;

      {Est-ce un compte de banque}
      Result := vGen.FindFirst(['G_GENERAL', 'G_NATUREGENE'], [Cpte, 'BQE'], True) <> nil;
      if Result then Exit;

      {S'agit-il d'un compte collectif / TIC / TID et d'un journal d'Achat / Vente / OD mais pas une écriture d'écart de change}
      Result := ((vJal.FindFirst(['J_JOURNAL', 'J_NATUREJAL'], [Jnal, 'ACH'], True) <> nil) or
                 (vJal.FindFirst(['J_JOURNAL', 'J_NATUREJAL'], [Jnal, 'VTE'], True) <> nil) or
                 ((vJal.FindFirst(['J_JOURNAL', 'J_NATUREJAL'], [Jnal, 'OD'], True) <> nil) and NonEcc))
                and
                ((vGen.FindFirst(['G_GENERAL', 'G_NATUREGENE'], [Cpte, 'COD'], True) <> nil) or
                 (vGen.FindFirst(['G_GENERAL', 'G_NATUREGENE'], [Cpte, 'COC'], True) <> nil) or
                 (vGen.FindFirst(['G_GENERAL', 'G_NATUREGENE'], [Cpte, 'COF'], True) <> nil) or
                 (vGen.FindFirst(['G_GENERAL', 'G_NATUREGENE'], [Cpte, 'COS'], True) <> nil) or
                 (vGen.FindFirst(['G_GENERAL', 'G_NATUREGENE'], [Cpte, 'TIC'], True) <> nil) or
                 (vGen.FindFirst(['G_GENERAL', 'G_NATUREGENE'], [Cpte, 'TID'], True) <> nil));
      if Result then Exit;

      {JP 09/08/06 : Gestion des comptes courants.
                     Il s'agit d'un compte divers figurant dans CLIENSSOC et d'un journal de banque ou d'OD paramétré
                     Cette fonction étant appelée depuis la modification d'entête de pièce, je pars
                     du principe qu'il n'y a pas lieu de gérer le dossier pour le ParamSoc et CLIENSSOC}
      Result := EstMultiSoc and
                (vGen.FindFirst(['G_GENERAL', 'G_NATUREGENE'], [Cpte, 'DIV'], True) <> nil) and
                ((vJal.FindFirst(['J_JOURNAL', 'J_NATUREJAL'], [Jnal, 'BQE'], True) <> nil) or
                 (vJal.FindFirst(['J_JOURNAL', 'J_NATUREJAL'], [Jnal, 'OD' ], True) <> nil)) and
                ExisteSQL('SELECT CLS_GENERAL FROM CLIENSSOC WHERE CLS_GENERAL = "' + Cpte + '"');
    end;
begin
  for n := 0 to vEcr.Detail.Count - 1 do begin
    Cpte := vEcr.Detail[n].GetString('E_GENERAL');
    Jnal := vEcr.Detail[n].GetString('E_JOURNAL');
    if Synchronisable(vEcr.Detail[n].GetString('E_NATUREPIECE') <> 'ECC') then
      vEcr.Detail[n].SetString('E_TRESOSYNCHRO', 'CRE')
    else
      vEcr.Detail[n].SetString('E_TRESOSYNCHRO', 'RIE');
  end;
end;

{JP 12/04/05 : Destruction d'une écriture à partir d'un RMVT
 JP 09/08/06 : gestion du multi sociétés en Tréso  : je pars du principe que la pièce comptable
               appartient au dossier V_PGI.SchemaName !!
 JP 18/06/07 : FQ 20766 : ajout d'un espace avant le AND TE_NUMTRANSA}
{---------------------------------------------------------------------------------------}
procedure DetruitEcritureTresoRMVT(Mvt : RMVT);
{---------------------------------------------------------------------------------------}
var
  lStSQL : string;
  Period : Integer;
begin

  {La trésorerie ne figure que dans une base, paramétrée dans le paramsoc SO_TRBASETRESO}
  lStSQL := 'DELETE FROM ' + GetTableDossier(GetParamSocSecur('SO_TRBASETRESO', ''), 'TRECRITURE') + ' WHERE '
            {$IFDEF NOVH}
            + 'TE_NODOSSIER = "' + TCPContexte.GetCurrent.TresoNoDossier + '" AND ' {JP 09/08/06 : TRESO MS}
            {$ELSE}
            + 'TE_NODOSSIER = "' + VH^.TresoNoDossier + '" AND ' {JP 09/08/06 : TRESO MS}
            {$ENDIF NOVH}
            + 'TE_JOURNAL = "' + Mvt.Jal + '" AND '
            + 'TE_EXERCICE = "' + Mvt.Exo + '" AND '
            + 'TE_NUMEROPIECE = ' + IntToStr(Mvt.Num) + ' AND '
            + 'TE_CPNUMLIGNE = ' + IntToStr(Mvt.NumLigne) + ' AND '
            + 'TE_NUMECHE = ' + IntToStr(Mvt.NumEche);
  if Mvt.Bordereau then begin
    Period := GetPeriode(Mvt.DateC);
    Period := Period - 200000; {pour supprimer le 2 et le 0 du début de période}
    {18/06/07 : FQ 20766 : ajout d'un espace avant le AND TE_NUMTRANSA}
    lStSQL := lStSQL + ' AND TE_NUMTRANSAC LIKE "' + TRANSACIMPORT + V_PGI.CodeSociete + Mvt.Jal + IntToStr(Period) + '%" ';
  end;
  ExecuteSQL(lStSQL);
end;

{JP 12/04/05 : Renvoie vide si E_TRESOSYNCHRO n'a pas besoin d'être modifié, sinon renvoie
               la partie de reuqête concernant E_TRESOSYNCHRO
{---------------------------------------------------------------------------------------}
function MajETresoSynchro(Mvt : RMVT) : string;
{---------------------------------------------------------------------------------------}
begin
  Result := '';

  {$IFDEF NOVH}
  if (Mvt.Simul <> 'N') or ((GetParamSocSecur('SO_EXOV8','') <> '') and (Mvt.DateC < TCPContexte.GetCurrent.Exercice.ExoV8.Deb)) then
  {$ELSE}
  if (Mvt.Simul <> 'N') or ((GetParamSocSecur('SO_EXOV8','') <> '') and (Mvt.DateC < VH^.ExoV8.Deb)) then
  {$ENDIF NOVH}
    Exit ;
  {On exclut les écriture d'à-nouveaux typées "OAN"}
  if Mvt.ANouveau and
     ExisteSQL('SELECT E_ECRANOUVEAU FROM ECRITURE WHERE E_ECRANOUVEAU = "OAN" AND ' + WhereEcriture(tsGene, Mvt, True)) then
    Exit;

  {On regarde si le compte est synchronisable}
  if EstCptSyn( Mvt.General, Mvt.Jal, Mvt.Nature <> 'ECC' ) then begin
    Result := ', E_TRESOSYNCHRO = "' + ets_Nouveau + '" ';
    {JP 15/05/07 : FQ 10467 : destruction du flux Tréso, pour que la Tréso prenne en compte la modification}
    DetruitEcritureTresoRMVT(Mvt);
  end;
end;

{SBO 16/01/05 : Retourne True si le compte est synchronisable en Trésorerie.
                Equivalent de EstCptSyn sans requête (SaisUtil.pas)
 JP 09/08/06 : Gestion des comptes courants/groupes
{---------------------------------------------------------------------------------------}
function  EstCptSynTob( vTobEcr : Tob ; vInfo : TInfoEcriture = nil ) : Boolean ;
{---------------------------------------------------------------------------------------}
var lBoDetruireInfo : Boolean ;
begin

  result := False ;

  if not Assigned( vTobEcr ) then
    exit ;

  {JP 17/04/07 : C'est plus prudent}
  if vTobEcr.GetValue('E_QUALIFPIECE') <> 'N' then Exit;

  // Gestion auto du TInfoEcriture
  if vInfo = nil then
    begin
    vInfo           := TInfoEcriture.Create ;
    lBoDetruireInfo := true ;
    end
  else lBoDetruireInfo := false ;

  // Le compte est-il un compte de banque ?
  if vInfo.LoadCompte( vTobEcr.GetString('E_GENERAL') ) then
    result := vInfo.Compte.GetValue('G_NATUREGENE') = 'BQE' ;

  // Autres cas...
  if not result then
    begin
    result := ( vTobEcr.GetString('E_NATUREPIECE') <> 'ECC' ) and
              ( vInfo.LoadJournal( vTobEcr.GetString('E_JOURNAL') ) and ( pos( vInfo.Journal.GetValue('J_NATUREJAL'), 'ACH;VTE;OD') > 0 ) ) and
              ( vInfo.LoadCompte(  vTobEcr.GetString('E_GENERAL') ) and ( pos( vInfo.Compte.GetValue('G_NATUREGENE'), 'COD;COC;COF;COS;TIC;TID') > 0 ) ) ;
    end ;
  {JP 09/08/06 : Gestion des comptes courants}
  if not Result then
    Result := EstMultiSoc and
              (vInfo.Compte.GetValue('G_NATUREGENE') = 'DIV') and
              ((vInfo.Journal.GetValue('J_NATUREJAL') = 'BQE') or
               (vInfo.Journal.GetValue('J_NATUREJAL') = 'OD')) and
              ExisteSQL('SELECT CLS_GENERAL FROM ' + GetTableDossier(vInfo.Dossier, 'CLIENSSOC') +
                        ' WHERE CLS_GENERAL = "' + vTobEcr.GetString('E_GENERAL') + '"');

  // Libératin du TInfoEcriture si gestion auto...
  if lBoDetruireInfo then vInfo.Free ;

end;


end.
