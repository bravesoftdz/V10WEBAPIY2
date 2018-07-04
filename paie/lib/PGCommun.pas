{***********UNITE*************************************************
Auteur  ...... : Paie Pgi
Créé le ...... : 24/04/2003
Modifié le ... :   /  /
Description .. : Fonction de la paie pouvant être commune aux applis...
Mots clefs ... : PAIE;COMMUN
*****************************************************************}
{
PT1   : 28/08/2003 PH V_421 NVelle methode creation ressource=saisie manuelle
PT2   : 05/11/2003 SB V_50 Maj zone
PT3   : 15/06/2004 PH V_50 Traitement du code auxiliaire
PT4   : 28/06/2004 PH V_50 Publication de la fonction cration d'un etablissement
                           social
PT5   : 14/01/2005 SB V_50 FQ 11320 : Affectation de la zone calendrier
PT6   : 31/01/2005 SB V_60 Intégration de proc. communes
PT7   : 03/02/2005 SB V_60 FQ 11692 Initialisation des dates à Idate1900
PT8   : 14/02/2005 SB V_60 FQ 11349 Refonte effacement ressource
PT9   : 28/02/2006 SB V_65 FQ 10397 Gestion d'annulation
---- JL 20/03/2006 modification clé annuaire ----
PT10  : 19/05/2006 SB V_65 Appel des fonctions communes pour identifier
                           V_PGI.driver
PT11  : 09/06/2006 SB V_65 FQ 13202 Function pour ajout clause sur option
                           contrôle absence/motif d'absence
PT12  : 17/10/2006 SB V_70 Refonte Fn pour utilisation portail => exporter vers
                           pgcalendrier
PT13  : 18/12/2006 SB V_70 Correction conversion erroné si format date anglais
PT14  : 25/01/2007 SB V_70 Compatibilité format date internationalisation
PT15  : 02/02/2007 SB V_70 Dpcmt de la function PGEncodeDateBissextile dans
                           pgcalendrier
PT16  : 28/08/2007 FC V_72 Extraction de fonctions dans ULibEditionPaie
PT17  : 29/05/2008 VG V_80 Résolution d'un Acces Violation
}
unit PGCommun;

interface

uses SysUtils,
  {$IFNDEF EAGLCLIENT}
  {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
 {$IFNDEF EAGLSERVER}
  Fe_Main,
  {$ENDIF}
  {$ELSE}
  MaineAgl,
  {$ENDIF}
  HCtrls,
  Utob,
  HEnt1,
  HmsgBox,
  ParamSoc,
  utilpgi
  ,ULibEditionPaie //PT16
  ;

//Fonctions publiques utilisées par d'autres appli.
function PGCreeRessource(Matricule, Ressource: string; CreeEnrTable: Boolean): Boolean;
function PGGetRessource(Matricule, Libelle, Prenom: string): string;
function PGRendRessInitiale(Initiale1, Initiale2: string): string;
function PGSupprRessource(ressource: string): boolean;
function PGCreerEtabComp(etablissement, Libelle, Localisation: string): Boolean; // PT4

//Fonctions publiques utilisées par la paie  PT6
procedure InitialiseTobAbsenceSalarie(TCP: tob);
//function  PGEncodeDateBissextile(AA, MM, JJ: WORD): TDateTime; PT15
//PT16 function  CalculPeriode(DTClot, DTValidite: TDatetime): integer;


function  PGisMssql: boolean;  { PT10 }
function  PGisOracle: boolean; { PT10 }
function  PGisDB2: boolean;    { PT10 }
function  PGisSYBASE: boolean; { PT10 }

implementation
uses PgCalendrier;


{***********A.G.L.***********************************************
Auteur  ...... : PH
Créé le ...... : 25/06/2004
Modifié le ... :   /  /
Description .. : Fonction de création d'un enregistrement dans la table
Suite ........ : ETABCOMPL
Mots clefs ... : PAIE;ETABCOMPL
*****************************************************************}
function PGCreerEtabComp(etablissement, Libelle, Localisation: string): Boolean;
var
  T, T1: TOB;
  DateFerm: TDateTime;
  ANow, MNow, DNow: Word;
begin
  Decodedate(dATE, ANow, MNow, DNow);
  T := TOB.Create('Mon etablissement', nil, -1);
  T1 := TOB.Create('ETABCOMPL', T, -1);
  T1.PutValue('ETB_CONGESPAYES', '-');
  DateFerm := EncodeDate(ANow, 05, 31);
  if DateFerm <= Now then DateFerm := EncodeDate(ANow + 1, 05, 31);
  T1.PutValue('ETB_DATECLOTURECPN', DateFerm);
  T1.PutValue('ETB_JOURHEURE', (GetParamSoc('SO_PGJOURHEURE')));
  {PT39
     T1.PutValue ('ETB_CONGESPAYES', 'X');
  }
  T1.PutValue('ETB_CONGESPAYES', '-');
  T1.PutValue('ETB_CODESECTION', '1');
  //FIN PT39
  T1.PutValue('ETB_BCJOURPAIEMENT', 0);
  T1.PutValue('ETB_JEDTDU', 0);
  T1.PutValue('ETB_JEDTAU', 0);
  T1.PutValue('ETB_PRUDHCOLL', '1');
  T1.PutValue('ETB_PRUDHSECT', '4');
  T1.PutValue('ETB_PRUDHVOTE', '1');
  T1.PutValue('ETB_ISLICSPEC', '-');
  T1.PutValue('ETB_ISOCCAS', '-');
  T1.PutValue('ETB_ISLABELP', '-');
  T1.PutValue('ETB_ETABLISSEMENT', etablissement);
  T1.PutValue('ETB_LIBELLE', LIBELLE);
  T1.PutValue('ETB_NBJOUTRAV', 0);
  T1.PutValue('ETB_NBREACQUISCP', 0);
  T1.PutValue('ETB_NBACQUISCP', '');
  T1.PutValue('ETB_NBRECPSUPP', 0); { PT52-2 }
  T1.PutValue('ETB_TYPDATANC', '1'); //PORTAGE CWAS Champ combo non integer
  T1.PutValue('ETB_DATEACQCPANC', '');
  T1.PutValue('ETB_VALORINDEMCP', '');
  T1.PutValue('ETB_RELIQUAT', '');
  T1.PutValue('ETB_PROFILCGE', '');
  T1.PutValue('ETB_BASANCCP', '');
  T1.PutValue('ETB_VALANCCP', '');
  T1.PutValue('ETB_PERIODECP', 0);
  T1.PutValue('ETB_1ERREPOSH', '');
  T1.PutValue('ETB_2EMEREPOSH', '');
  T1.PutValue('ETB_MVALOMS', '');
  T1.PutValue('ETB_VALODXMN', 0);
  T1.PutValue('ETB_TYPDATANC', '');
  T1.PutValue('ETB_MVALOMS', '');
  T1.PutValue('ETB_PCTFRAISPROF', 0);
  T1.PutValue('ETB_HORAIREETABL', 0);
  T1.PutValue('ETB_DATEVALTRANS', Idate1900);
  T1.PutValue('ETB_SEUILTEMPSPAR', 0);
  T1.PutValue('ETB_PRORATATVA', 0);
  T1.PutValue('ETB_JOURPAIEMENT', 0);
  T1.PutValue('ETB_EDITBULCP', ''); //PT46
  //PT51  : 05/04/2004 PH V_50 FQ 11231 Prise en compte Alsace LOrraine
  T1.PutValue('ETB_REGIMEALSACE', '-');
  T1.PutValue('ETB_MEDTRAV', -1); { PT54 }
  T1.PutValue('ETB_CODEDDTEFP', -1); { PT54 }
  T1.PutValue('ETB_MEDTRAVGU', '');
  T1.PutValue('ETB_CODEDDTEFPGU', '');
  T1.PutValue('ETB_SUBROGATION','-');
  Result := TRUE ;
  try
  T.InsertDb(nil, FALSE);
  except
  result := FALSE;
  PGIError ('Erreur de création de l''établissement social','Création établissement social');
  end ;
  FreeAndNil(T);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Paie Pgi
Créé le ...... : 24/04/2003
Modifié le ... : 24/04/2003
Description .. : Création d'une ressource de type SAL en fonction du
Suite ........ : descriptif des params soc de la paie
Mots clefs ... : PAIE;RESSOURCE
*****************************************************************}
function PGCreeRessource(Matricule, Ressource: string; CreeEnrTable: Boolean): Boolean;
var
  Q: TQuery;
  Tob_Sal, Tob_Ress, T_Res: Tob;
begin
  result := False;
  { En saisie libre et sans la ressource transmise on sort!! }
  if (GetParamSoc('SO_PGTYPEAFFECTRES') = '') and (ressource = '') then exit
  else
  begin
    { Chargement des données salariés }
    Q := OpenSql('SELECT PSA_ETABLISSEMENT,PSA_LIBELLE,PSA_PRENOM,PSA_DATESORTIE,' +
      'PSA_ADRESSE1,PSA_ADRESSE2,PSA_ADRESSE3,PSA_CODEPOSTAL,' +
      'PSA_VILLE,PSA_PAYS,PSA_TELEPHONE,PSA_PORTABLE,PSA_NUMEROSS,PSA_AUXILIAIRE,'+
      'PSA_CALENDRIER ' + { PT5 }
      'FROM SALARIES WHERE PSA_SALARIE="' + Matricule + '"', True);
    if Q.eof then
    begin
      Ferme(Q);
      Exit;
    end
    else
    begin
      Tob_Sal := Tob.Create('Les salariés', nil, -1);
      Tob_Sal.LoadDetailDB('Les salariés', '', '', Q, False);
    end;
    Ferme(Q);
    { Affectation de la ressource selon le paramètre société et si non saisie }
    if Ressource = '' then Ressource := PGGetRessource(Matricule, Tob_Sal.Detail[0].GetValue('PSA_LIBELLE'), Tob_Sal.Detail[0].GetValue('PSA_PRENOM'));
    { Création de la ressource via la tob }
    if (CreeEnrTable) and (Ressource <> '') then
    begin
      Tob_Ress := Tob.create('MERE_RESSOURCE', nil, -1);
      T_Res := Tob.create('RESSOURCE', Tob_Ress, -1);
      T_Res.AddChampSupValeur('ARS_TYPERESSOURCE', 'SAL');
      T_Res.AddChampSupValeur('ARS_RESSOURCE', Ressource);
      T_Res.AddChampSupValeur('ARS_SALARIE', Matricule);
      T_Res.AddChampSupValeur('ARS_ETABLISSEMENT', Tob_Sal.detail[0].GetValue('PSA_ETABLISSEMENT'));
      T_Res.AddChampSupValeur('ARS_LIBELLE', Tob_Sal.detail[0].GetValue('PSA_LIBELLE'));
      T_Res.AddChampSupValeur('ARS_LIBELLE2', Tob_Sal.detail[0].GetValue('PSA_PRENOM'));
      T_Res.AddChampSupValeur('ARS_DATESORTIE', Tob_Sal.detail[0].GetValue('PSA_DATESORTIE')); //PT2
      T_Res.AddChampSupValeur('ARS_ADRESSE1', Tob_Sal.detail[0].GetValue('PSA_ADRESSE1'));
      T_Res.AddChampSupValeur('ARS_ADRESSE2', Tob_Sal.detail[0].GetValue('PSA_ADRESSE2'));
      T_Res.AddChampSupValeur('ARS_ADRESSE3', Tob_Sal.detail[0].GetValue('PSA_ADRESSE3')); //PT2
      T_Res.AddChampSupValeur('ARS_CODEPOSTAL', Tob_Sal.detail[0].GetValue('PSA_CODEPOSTAL'));
      T_Res.AddChampSupValeur('ARS_VILLE', Tob_Sal.detail[0].GetValue('PSA_VILLE'));
      T_Res.AddChampSupValeur('ARS_PAYS', Tob_Sal.detail[0].GetValue('PSA_PAYS'));
      T_Res.AddChampSupValeur('ARS_TELEPHONE', Tob_Sal.detail[0].GetValue('PSA_TELEPHONE'));
      T_Res.AddChampSupValeur('ARS_TELEPHONE2', Tob_Sal.detail[0].GetValue('PSA_PORTABLE'));
      T_Res.AddChampSupValeur('ARS_IMMAT', Tob_Sal.detail[0].GetValue('PSA_NUMEROSS'));
      T_Res.AddChampSupValeur('ARS_AUXILIAIRE', Tob_Sal.detail[0].GetValue('PSA_AUXILIAIRE')); // PT3
      T_Res.AddChampSupValeur('ARS_ESTHUMAIN', 'X');
      T_Res.AddChampSupValeur('ARS_MENSUALISE', 'X');
      T_Res.AddChampSupValeur('ARS_FINDISPO', iDate2099);
      T_Res.AddChampSupValeur('ARS_STANDCALEN', Tob_Sal.detail[0].GetValue('PSA_CALENDRIER')); { PT5 }
      if T_Res.InsertOrUpdateDB then result := True;
      Tob_Ress.free;
    end
    else
      // PT1  PH 28/08/2003 V_421 Creation ok que si le code ressource est saisi
      if Ressource <> '' then result := True;
    Tob_Sal.free;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Paie Pgi
Créé le ...... : 24/04/2003
Modifié le ... :   /  /
Description .. : Composition de la ressource en fonction de la
Suite ........ : sélection des params soc paie
Mots clefs ... : PAIE;RESSOURCE
*****************************************************************}
function PGGetRessource(Matricule, Libelle, Prenom: string): string;
var
  Initiale1, Initiale2: string;
begin
  Initiale1 := '';
  Initiale2 := '';
  if GetParamSoc('SO_PGTYPEAFFECTRES') = '' then result := ''
  else if GetParamSoc('SO_PGTYPEAFFECTRES') = 'MAT' then Result := Matricule
  else if GetParamSoc('SO_PGTYPEAFFECTRES') = 'RAC' then Result := 'SAL' + Matricule
    // PT1  PH 28/08/2003 V_421 NVelle methode creation ressource=saisie manuelle
  else if GetParamSoc('SO_PGTYPEAFFECTRES') = 'SAI' then
 {$IFNDEF EAGLSERVER}
   Result := AglLanceFiche('PAY', 'CREATRESSOURCE', '', '', '')
 {$ENDIF}  
  else
  begin
    if GetParamSoc('SO_PGNBCARNOMRES') > 0 then
      Initiale1 := UpperCase(Copy(Libelle, 1, StrToInt(GetParamSoc('SO_PGNBCARNOMRES'))));

    if GetParamSoc('SO_PGNBCARPRENRES') > 0 then
      Initiale2 := UpperCase(Copy(Prenom, 1, StrToInt(GetParamSoc('SO_PGNBCARPRENRES'))));
    if GetParamSoc('SO_PGTYPEAFFECTRES') = 'RNP' then Result := PGRendRessInitiale(Initiale1, Initiale2)
    else if GetParamSoc('SO_PGTYPEAFFECTRES') = 'RPN' then Result := PGRendRessInitiale(Initiale2, Initiale1);
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Paie Pgi
Créé le ...... : 24/04/2003
Modifié le ... :   /  /
Description .. : Composition de la ressource de type SAL prefixé d'initiales
Suite ........ : et d'incrément
Mots clefs ... : PAIE;RESSOURCE
*****************************************************************}
function PGRendRessInitiale(Initiale1, Initiale2: string): string;
var
Q : TQuery;
Num : string;
begin
{Recherche des ressources existantes dans l'ordre décroissant}
Q:= OpenSql ('SELECT ARS_RESSOURCE'+
             ' FROM RESSOURCE WHERE'+
             ' ARS_TYPERESSOURCE="SAL" AND'+
             ' ARS_RESSOURCE LIKE "'+Initiale1+Initiale2+'%"'+
             ' ORDER BY ARS_RESSOURCE DESC', True);
{si ressource de type initiale inexistante}
if (Q.eof) then
   Result:= Initiale1+Initiale2+'1'
else
   begin
{sinon recherche suffixe ressource numérique pour incrément}
{PT17
   while (Isnumeric(Copy(Q.FindField('ARS_RESSOURCE').AsString, 3, Length(Q.FindField('ARS_RESSOURCE').AsString))) = False) and (Q.eof = False) do
}
   while (Q.eof=False) and
         (Isnumeric (Copy(Q.FindField ('ARS_RESSOURCE').AsString, 3,
                     Length (Q.FindField ('ARS_RESSOURCE').AsString)))=False) do
         Q.next;
   if (not Q.eof) then
      begin {si trouvé valeur numerique, on incrémente de 1}
      Num:= Copy (Q.FindField ('ARS_RESSOURCE').AsString, 3,
                  Length (Q.FindField ('ARS_RESSOURCE').AsString));
      Result:= Initiale1+Initiale2+IntToStr (StrToInt (Num)+1);
      end
   else
      Result:= Initiale1+Initiale2+'1';
   end;
Ferme (Q);
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Paie Pgi
Créé le ...... : 24/04/2003
Modifié le ... :   /  /
Description .. : Suppression des informations salariés associées à la
Suite ........ : ressource
Mots clefs ... : PAIE;RESSOURCE
*****************************************************************}
function PGSupprRessource(ressource: string): boolean;
var
  Tob_Ress: Tob;
begin
  result := False;
  Tob_Ress := Tob.create('RESSOURCE', nil, -1);
  Tob_Ress.SelectDB('"' + ressource + '"', nil);
  //Tob_Ress.PutValue('ARS_TYPERESSOURCE' , 'SAL');
  //Tob_Ress.PutValue('ARS_RESSOURCE'     , Ressource);
  Tob_Ress.PutValue('ARS_SALARIE', '');
(*  Tob_Ress.PutValue('ARS_ETABLISSEMENT', '');  PT8 Mise en commentaire
  Tob_Ress.PutValue('ARS_LIBELLE', '');
  Tob_Ress.PutValue('ARS_LIBELLE2', '');
  Tob_Ress.PutValue('ARS_DATESORTIE', idate1900); //PT2
  Tob_Ress.PutValue('ARS_ADRESSE1', '');
  Tob_Ress.PutValue('ARS_ADRESSE2', '');
  Tob_Ress.PutValue('ARS_ADRESSE3', ''); //PT2
  Tob_Ress.PutValue('ARS_CODEPOSTAL', '');
  Tob_Ress.PutValue('ARS_VILLE', '');
  Tob_Ress.PutValue('ARS_PAYS', '');
  Tob_Ress.PutValue('ARS_TELEPHONE', '');
  Tob_Ress.PutValue('ARS_TELEPHONE2', '');
  Tob_Ress.PutValue('ARS_IMMAT', '');
  //  Tob_Ress.PutValue('ARS_AUXILIAIRE'    , '');
  Tob_Ress.PutValue('ARS_STANDCALEN', ''); { PT5 }*)
  if Tob_Ress.UpdateDB then result := True;
  Tob_Ress.free;
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Paie Pgi
Créé le ...... : 03/02/2003
Modifié le ... :   /  /
Description .. : fonction d'initialisation d'un mouvement congés payés
Suite ........ : A appeler avant tout renseignement de nouveau mvt
Mots clefs ... : PAIE;CP
*****************************************************************}
//PT6
procedure InitialiseTobAbsenceSalarie(TCP: tob);
begin
  TCP.Putvalue('PCN_TYPEMVT'             , 'CPA' );
  TCP.Putvalue('PCN_SALARIE'             , ''    );
  TCP.Putvalue('PCN_DATEDEBUT'           , Idate1900 );   //PT7
  TCP.Putvalue('PCN_DATEFIN'             , Idate1900 );   //PT7
  TCP.Putvalue('PCN_ORDRE'               , 0     );
  TCP.Putvalue('PCN_MVTPRIS'             , ''    );
  TCP.PutValue('PCN_PERIODECP'           , -1    );
  TCP.PutValue('PCN_PERIODEPY'           , -1    );
  TCP.Putvalue('PCN_TYPECONGE'           , ''    );
  TCP.Putvalue('PCN_SENSABS'             , ''    );
  TCP.Putvalue('PCN_LIBELLE'             , ''    );
  TCP.Putvalue('PCN_GENERECLOTURE'       ,'-'    );
  TCP.Putvalue('PCN_DATESOLDE'           , Idate1900 );   //PT7
  TCP.Putvalue('PCN_DATEVALIDITE'        , Idate1900 );   //PT7
  TCP.Putvalue('PCN_DATEDEBUTABS'        , Idate1900 );   //PT7
  TCP.Putvalue('PCN_DEBUTDJ'             , ''    );
  TCP.Putvalue('PCN_DATEFINABS'          , Idate1900 );   //PT7
  TCP.Putvalue('PCN_FINDJ'               , ''    );
  TCP.Putvalue('PCN_DATEPAIEMENT'        , Idate1900 );   //PT7
  TCP.Putvalue('PCN_CODETAPE'            , '...' );
  TCP.Putvalue('PCN_JOURS'               , 0     );
  TCP.Putvalue('PCN_HEURES'              , 0     );
  TCP.Putvalue('PCN_BASE'                , 0     );
  TCP.Putvalue('PCN_NBREMOIS'            , 0     );
  TCP.Putvalue('PCN_CODERGRPT'           , 0     );
  TCP.PutValue('PCN_MVTDUPLIQUE'         , '-'   );
  TCP.Putvalue('PCN_ABSENCE'             , 0     );
  TCP.Putvalue('PCN_MODIFABSENCE'        , '-'   );
  TCP.Putvalue('PCN_APAYES'              , 0     );
  TCP.Putvalue('PCN_VALOX'               , 0     );
  TCP.Putvalue('PCN_VALOMS'              , 0     );
  TCP.Putvalue('PCN_VALORETENUE'         , 0     );
  TCP.Putvalue('PCN_VALOMANUELLE'        , 0     );
  TCP.Putvalue('PCN_MODIFVALO'           , '-'   );
  TCP.Putvalue('PCN_PERIODEPAIE'         ,''     );
  TCP.Putvalue('PCN_ETABLISSEMENT'       , ''    );
  TCP.Putvalue('PCN_TRAVAILN1'           , ''    );
  TCP.Putvalue('PCN_TRAVAILN2'           , ''    );
  TCP.Putvalue('PCN_TRAVAILN3'           , ''    );
  TCP.Putvalue('PCN_TRAVAILN4'           , ''    );
  TCP.Putvalue('PCN_CODESTAT'            , ''    );
  TCP.Putvalue('PCN_SAISIEDEPORTEE'       , '-'   );
  TCP.Putvalue('PCN_VALIDSALARIE'        , ''    );
  TCP.Putvalue('PCN_VALIDRESP'           , ''    );
  TCP.Putvalue('PCN_EXPORTOK'            , ''    );
  TCP.Putvalue('PCN_LIBCOMPL1'           , ''    );
  TCP.Putvalue('PCN_LIBCOMPL2'           , ''    );
  TCP.Putvalue('PCN_VALIDABSENCE'        , ''    );
  TCP.Putvalue('PCN_OKFRACTION'          , ''    );
  TCP.Putvalue('PCN_NBJCARENCE'          , 0     );
  TCP.Putvalue('PCN_NBJCALEND'           , 0     );
  TCP.Putvalue('PCN_NBJIJSS'             , 0     );
  TCP.Putvalue('PCN_IJSSSOLDEE'          , '-'   );
  TCP.Putvalue('PCN_GESTIONIJSS'         , '-'   );
  TCP.Putvalue('PCN_ETATPOSTPAIE'        , 'VAL' ); { PT9 }
end;


(*  PT15 Mise en commentaire, déplacé dans pgcalendrier //PT6
function PGEncodeDateBissextile(AA, MM, JJ: WORD): TDateTime;
begin
//  if IsValidDate(IntToStr(JJ) + '/' + IntToStr(MM) + '/' + IntToStr(AA)) then
//    Result := encodedate(AA, MM, JJ); PT13 Mise en commentaire
  if (MM = 2) and (JJ = 29) {((JJ = 28) or ) PT14 } then //Année bissextile
  begin
    Result := encodedate(AA, MM, 1);
    Result := FindeMois(Result);
  end
else
    Result := encodedate(AA, MM, JJ);    { PT13 }
end;

//PT6   *)
(*PT16
function CalculPeriode(DTClot, DTValidite: TDatetime): integer;
var
  Dtdeb, DtFin, DtFinS: TDATETIME;
  aa, mm, jj: word;
  i: integer;
begin
  result := -1;
  if DTClot <= idate1900 then exit;
  Decodedate(DTclot, aa, mm, jj);
  DtDeb := PGEncodeDateBissextile(AA - 1, MM, JJ) + 1;
  DtFin := DtClot;
  DtFinS := PGEncodeDateBissextile(AA + 1, MM, JJ);
  if Dtvalidite > Dtfins then
  begin
    result := -9;
    exit;
  end;
  if DtValidite > DtClot then exit;
  result := 0;
  i := 0;
  while not ((DTValidite >= DtDeb) and (DTValidite <= DtFin)) do
  begin
    i := i + 1;
    if i > 50 then exit; // pour ne pas boucler au cas où....
    result := result + 1;
    DtFin := DtDeb - 1;
    Decodedate(DTFin, aa, mm, jj);
    DtDeb := PGEncodeDateBissextile(AA - 1, MM, JJ) + 1;
  end;
end; *)


function PGisMssql: boolean;
begin
  Result := isMssql;    { PT10 }
end;

function PGisOracle: boolean;
begin
  Result := isOracle;  { PT10 }
end;

function PGisDB2: boolean;
begin
  Result := isDB2;     { PT10 }
end;

function PGisSYBASE: boolean;
begin
  Result := V_PGI.Driver in [dbSYBASE] { PT10 }
end;



end.

