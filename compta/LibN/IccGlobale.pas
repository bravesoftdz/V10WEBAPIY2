{***********UNITE*************************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 04/03/2003
Modifié le ... :   /  /    
Description .. : Passage en eAGL
Suite ........ : GCO - 26/03/2004 - FB 11733
Mots clefs ... : 
*****************************************************************}
unit IccGlobale;

interface

uses Classes,
{$IFDEF EAGLCLIENT}

{$ELSE}
     {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}
{$ENDIF}
     uTOB,
     Hctrls,
     Ent1,
     Hent1,
     ParamSoc,
     AGLInit,
     HmsgBox,
     sysUtils
     ;

// Calcul le taux maximum légal applicable pour l' exercice
function CalculTauxMaxLegal( DateDebut,DateFin : TDateTime ) : Double;

// Test si la société est soumise à l'imposition
function TestImpositionSociete : Boolean;

// Fonction de répupération des ecritures dans les comte ICC
function TraitementEcriture( vCompte : String; vD1, vD2 : TDateTime ; vModeMiseAJour : Boolean ) : Boolean;

// Compare le Solde du compte ICC et le solde du compte General
function GetSoldeComptable ( vStCompte : String ; D1, D2 : TDateTime  ): Double;

// Recherche si nouvelles écritures sur le compte général pour les dates vD1 et vD2
function ExisteNouvelleEcriture ( vStCompte : String ; vD1, vD2 : TDateTime ) : Boolean;

// Recherche la valeur de départ pour la clé ICE_LIGNE
function GetNumLigneFromEcriture( vStGeneral : String ; vD1 : TDateTime ) : Integer;

type IccData =
  record
    NombreJours         : Integer ;
    DateDu              : TDateTime ;
    DateAu              : TDateTime ;
    Methode             : Word ;
    TauxLegal           : Double  ;
    TotInt              : Double  ;
    TotIntNonDedCapital : Double  ;
    TotIntNonDedTaux    : Double  ;
    TotIntNonDed        : Double  ;
    TotIntDed           : Double  ;
    SoldeDifferent      : Boolean ;
    NouvelleEcriture    : Boolean ;
    TauxAbsent          : Boolean ;
    TauxProvisoire      : Boolean ;
    AvecRecapitulatif   : Boolean ;
    SoumisIs            : Boolean ;
  end;

var
  ICC_Data : IccData;

Type
  TICCCalcul = class (TPersistent)
    constructor Create;
    destructor Destroy; override;
    public
      procedure CalculTOB(T :TOB; QCompte : TQuery);
      function SoustraitDate(Date1,Date2 : TDateTime; Mode : Integer) :Double;
      function CalculCapital(LaDate : TDateTime; LeSolde : Double): Double;
      procedure RechercheEcriture( LaTob : TOB );
      procedure RechercheVariationTaux ( LaTob : TOB );
      function  RechercheTaux ( lDate : TDatetime ) : Double;

      // Calcul le solde, le capital, le NB de jours, la Base, les ontérets déductibles
      procedure CalcDonnees(Tob1 : TOB);

      // Calcul les intérets non déductibles sur Capital
      procedure NonDeductibleSurCapital( LaTob : TOB );

      // Calcul les intérets non déductibles sur Taux
      procedure NonDeductibleSurTaux( LaTob : TOB );

    private
      FGroupe, FLimitation, FDirigeant : Boolean;
      FCompte : String;
      FTaux : Double;
      FDateBascule : TDateTime;

    end;

implementation

uses
  uLibExercice ,
  uLibEcriture ,
  uLibWindows,
  CalCole;       // GetCumul

////////////////////////////////////////////////////////////////////////////////
function TestImpositionSociete : Boolean;
var Q : TQuery;
Begin
  Result := False;
  Q := nil;
  try
    try
      // GCO - 29/08/2007 - FQ 21321
      Q := OpenSQL('SELECT DFI_IMPODIR, DFI_OPTIONAUTID FROM DPFISCAL LEFT JOIN ' +
                   'DOSSIER ON DOS_GUIDPER = DFI_GUIDPER WHERE ' +
                   'DOS_NODOSSIER = "' + V_PGI.NoDossier + '"', True);
      if Not Q.Eof then
      begin
        if (Q.FindField('DFI_IMPODIR').AsString = 'IS') or
           ((Q.FindField('DFI_IMPODIR').AsString = 'IR') and (Q.FindField('DFI_OPTIONAUTID').AsString = 'X')) then
          Result := True
        else
          Result := False;
      end;
    except
      on E: Exception do PgiError('Erreur SQL : ' + E.Message, 'Fonction : TestImpositionSociete');
    end;
  finally
    if Assigned(Q) then Ferme(Q);
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 09/04/2001
Modifié le ... : 09/04/2001
Description .. : Calcul le Taux Maximum légal par rapport à la
                 date d'exercice de l'exercice de référence
Mots clefs ... : TAUX MAXIMUM LEGAL
*****************************************************************}
function CalculTauxMaxLegal(DateDebut,DateFin : TDateTime): Double;
var D1,D2 : TDateTime;
    PremMois,PremAnnee,NbMois,NombreDeMois : Word;
    NbMoisRestant : Integer;
    LPremiereIteration : Boolean;
    ValRetour : Double;
    Q : TQuery;
begin
  ValRetour := 0;
  lPremiereIteration := True;
  Icc_Data.TauxProvisoire := False;

  // Calcul du nombre de mois sur la période d'exercice
  if DateFin <> FinDeMois(DateFin) then
  begin
    DateFin := PlusMois(DateFin,-1);
    DateFin := FinDeMois(DateFin);
  end;

  NombreMois(DateDebut,DateFin,PremMois,PremAnnee,NbMois);
  NombreDeMois := NbMois; // Nombre de mois de l'exercice
  NbMoisRestant := NbMois;

  D1 := CDateDebutTrimestre(DateDebut);
  D2 := CDateFinTrimestre(DateFin);

  Q := nil;
  try
    try
      Q := OpenSQL('SELECT ICN_CODETXMOYEN, ICN_DATEVALEUR, ICN_TXMOYEN, ICN_ETATICC FROM ICCELTNAT WHERE ' +
                   '(ICN_CODETXMOYEN = "N02" and ICN_DATEVALEUR >= "' + UsDateTime(D1) + '" and ICN_DATEVALEUR < "' +
                   UsDateTime(D2) +'") ORDER BY ICN_CODETXMOYEN, ICN_DATEVALEUR',True);

      while not Q.Eof do
      begin
        if (Q.FindField('ICN_ETATICC').AsString = 'PRO') and (not Icc_Data.TauxProvisoire) then
          Icc_Data.TauxProvisoire := True;

        if NbMoisRestant <> 0 then
        begin
          if lPremiereIteration then
          begin
            lPremiereIteration := False;
            D1 := DateDebut;
          end
          else
            D1 := Q.FindField('ICN_DATEVALEUR').AsDateTime;

          if ((NbMoisRestant=3) or (NbMoisRestant=4) or (NbMoisRestant=5)) and (DateFin <> CDateFinTrimestre(DateFin)) then
            D2 := DateFin
          else
            D2 := CDateFinTrimestre(Q.FindField('ICN_DATEVALEUR').AsDateTime);

          NombreMois(D1,D2,PremMois,PremAnnee,NbMois);
          NbMoisRestant := NbMoisRestant - NbMois;

          ValRetour := ValRetour + (NbMois * Q.FindField('ICN_TXMOYEN').AsFloat);

        end;
        Q.Next;
      end;
      ValRetour := ValRetour / NombreDeMois;
    except
      ValRetour := 0;
    end;
  finally
    if Assigned(Q) then Ferme(Q);
  end;

  if ValRetour = 0 then
  begin
    // On prend le dernier élémént National car les taux sont absents pour la période
    Q := nil;
    try
      Q := OpenSQl('SELECT ICN_TXMOYEN,ICN_DATEVALEUR FROM ICCELTNAT ORDER BY ICN_DATEVALEUR DESC',True);
      if not Q.EOF then
      begin
        ValRetour := Q.FindField('ICN_TXMOYEN').AsFloat;
        Icc_Data.TauxAbsent := True;
      end;

    finally
      if Assigned(Q) then Ferme(Q);
    end;
  end;

  Result := ValRetour;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 03/10/2001
Modifié le ... :   /  /
Description .. : Importe toutes les écritures de la compta dans le compte ICC
Mots clefs ... :
*****************************************************************}
function TraitementEcriture( vCompte : String; vD1, vD2 : TDateTime; vModeMiseAJour : Boolean ) : Boolean;
var lQ : TQuery;
    lTobEcriture, lTobDetail : TOB;
    i : integer;
    lNumLigne : integer;

    //  Fonction qui récupère et charge la TOB lTobEcriture avec les flux à récupérer
    //  Retour : True si des écritures ont été trouvées
    //           False si pas d'écriture ou Erreur
    function PrepareTob : Boolean ;
    begin
      Result := False;

      lQ := nil;
      try
        if vModeMiseAjour then
          lQ := OpenSql('Select E_GENERAL, E_NUMLIGNE, E_EXERCICE, E_LIBELLE, E_DEBIT, E_CREDIT, E_DATECOMPTABLE from ECRITURE where ' +
                        '(E_GENERAL = "' + vCompte + '") and ' +
                        '(E_QUALIFORIGINE <> "ICC") and ' +
                        '(E_QUALIFPIECE = "N") and ' +
                        '(E_ECRANOUVEAU = "N" or E_ECRANOUVEAU = "H") and ' +
                        //'(E_EXERCICE = "' + VH^.CPExoRef.Code + '") and ' +
                        '(E_DATECOMPTABLE >= "' + UsDateTime(vD1) + '") and ' +
                        '(E_DATECOMPTABLE <= "' + UsDateTime(vD2) + '") ' +
                        'Order by E_GENERAL, E_EXERCICE, E_DATECOMPTABLE, E_NUMEROPIECE, E_NUMLIGNE', True)
        else
          lQ := OpenSQL('Select E_GENERAL, E_NUMLIGNE, E_EXERCICE, E_LIBELLE, E_DEBIT, E_CREDIT, E_DATECOMPTABLE from ECRITURE where ' +
                        '(E_GENERAL = "' + vCompte + '") and ' +
                        '(E_QUALIFPIECE = "N") and ' +
                        '(E_ECRANOUVEAU = "N" or E_ECRANOUVEAU = "H") and ' +
                        //'(E_EXERCICE = "' + VH^.CPExoRef.Code + '") and ' +
                        '(E_DATECOMPTABLE >= "' + UsDateTime(vD1) + '") and ' +
                        '(E_DATECOMPTABLE <= "' + UsDateTime(vD2) + '") ' +
                        'Order by E_GENERAL, E_EXERCICE, E_DATECOMPTABLE, E_NUMEROPIECE, E_NUMLIGNE', True);

        if not lQ.EOF then
        begin
          lQ.First;
          while not lQ.Eof do
          begin
            lTobDetail := Tob.Create('ICCECRITURE', lTobEcriture, -1);
            lTobDetail.PutValue('ICE_GENERAL',    lQ.FindField('E_GENERAL').AsString);
            lTobDetail.PutValue('ICE_LIGNE',      lQ.FindField('E_NUMLIGNE').AsInteger);
            lTobDetail.PutValue('ICE_LIBELLE',    Copy(lQ.FindField('E_LIBELLE').AsString,1,30));
            lTobDetail.PutValue('ICE_DATEVALEUR', lQ.FindField('E_DATECOMPTABLE').AsDateTime);
            lTobDetail.PutValue('ICE_DEBIT',      lQ.FindField('E_DEBIT').AsFloat);
            lTobDetail.PutValue('ICE_CREDIT',     lQ.FindField('E_CREDIT').AsFloat);
            lQ.Next;
          end;
          Result := True;
        end;

      finally
        if Assigned( lQ ) then Ferme( lQ );
      end;

    end; // Fin de la fonction PrepareTob

    procedure AjouteTob ;
    begin
      lQ := nil;
      try
        if vModeMiseAJour then
        begin
          lQ := OpenSQL('Select ICE_GENERAL, ICE_LIGNE, ICE_DATEVALEUR, ICE_LIBELLE, ICE_DEBIT, ICE_CREDIT from ICCECRITURE where (ICE_GENERAL = "' + vCompte + '") and ' +
                        '(ICE_DATEVALEUR >= "' + UsDateTime(vD1) + '") ' +
                        'Order by ICE_GENERAL, ICE_LIGNE', True);
        end
        else
        begin
          lQ := OpenSQL('Select ICE_GENERAL, ICE_LIGNE, ICE_DATEVALEUR, ICE_LIBELLE, ICE_DEBIT, ICE_CREDIT from ICCECRITURE where (ICE_GENERAL = "' + vCompte + '") and ' +
                        '(ICE_DATEVALEUR > "' + UsDateTime(vD2) + '") ' +
                        'Order by ICE_GENERAL, ICE_LIGNE', True);
        end;
        if not lQ.EOF then
          lTobEcriture.LoadDetailDB('ICCECRITURE', '', '', lQ, True, False);
      finally
        if Assigned( lQ ) then Ferme( lQ );
      end;
    end;

    function UpDatedataBase : Boolean ;
    begin
      Result := False;

      try
        // Suppression des anciennes ecritures ICC concernées
        ExecuteSQL('Delete from ICCECRITURE where ' +
                   '(ICE_GENERAL = "' + vCompte + '") and ' +
                   '(ICE_DATEVALEUR >= "' + UsDateTime(vD1) + '")');
                   //'(ICE_DATEVALEUR <= "' + UsDateTime(vD2) + '")');

        // MAJ du Flag des ecritures de la compta récupérées }
        if vModeMiseAJour then
          ExecuteSQL('Update Ecriture set E_QUALIFORIGINE = "ICC" where '+
                     '(E_GENERAL = "' + vCompte + '") and ' +
                     '(E_DATECOMPTABLE >= "' + UsDateTime(vD1) + '") and ' +
                     '(E_DATECOMPTABLE <= "' + UsDateTime(vD2) + '") and ' +
                     //'(E_EXERCICE = "' + VH^.CPExoRef.Code + '") and ' +
                     '(E_QUALIFORIGINE <> "ICC")')
        else
          ExecuteSQL('Update Ecriture set E_QUALIFORIGINE = "ICC" where '+
                     '(E_GENERAL = "' + vCompte + '") and ' +
                     '(E_DATECOMPTABLE >= "' + UsDateTime(vD1) + '") and ' +
                     '(E_DATECOMPTABLE <= "' + UsDateTime(vD2) + '")');
                     //'(E_EXERCICE = "' + VH^.CPExoRef.Code + '")');

        // Insertion des ecritures dans la table ICCECRITURE
        lTobEcriture.SetAllModifie(True);
        lTobEcriture.InsertDB(nil,False);

        Result := True;
      except
        on E : Exception do
          PgiInfo(E.Message, 'Traitement annulé. Une erreur s''est produite sur le compte ' + vCompte);
      end;

    end;

begin
  Result := False;

  try
    lTobEcriture := Tob.Create('', nil, -1);

    if not PrepareTob then
    begin
      // Pas d'écritures trouvées ou Erreur lors de la requête
      Result := True;
      Exit;
    end;

    // Ajoute les anciennes Tob Existantes dans lTobEcriture
    AjouteTob;

    // Tri lTobEcriture par Date
    lTobEcriture.Detail.Sort('ICE_DATEVALEUR');
    // Récupération du numéro de ligne à utliser
    lNumLigne := GetNumLigneFromEcriture( vCompte, vD1 );
    // Renumérotation lTobEcriture
    for i:=0 to lTobEcriture.Detail.Count -1 do
    begin
     lTobEcriture.Detail[i].PutValue('ICE_LIGNE', lNumLigne + 1 );
     inc(LNumLigne);
    end;
    //lTobEcriture.SaveToFile('C:\LaTob.txt', False, True, True);

    if UpdateDataBase then Result := True;

  finally
    if lTobEcriture <> nil then lTobEcriture.Free;
  end;

end;

{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 09/10/2001
Modifié le ... : 10/10/2001
Description .. : Fonction qui compare le solde du compte ICC et le solde du
Suite ........ : compte général
Mots clefs ... :
*****************************************************************}
function GetSoldeComptable( vStCompte : String; D1, D2 : TDateTime): Double;
var //Q : TQuery;
    //SoldeIcc : Double;
    SoldeCompta : Double;
//    MtDeb : Double;
//    MtCre : Double;
    LesBalances : TStringList;
    LeResult : TabloExt;
begin
(*
//  Result      := False;
  SoldeIcc    := 0;
//  SoldeCompta := 0;

  // Solde du Compte ICC
  Q := nil;
  try
    Q := OpenSQL('Select Icg_Soldedebex from iccgeneraux where icg_general = "' + vStCompte + '"',True);
    if not Q.EOF then
      SoldeIcc := Q.FindField('Icg_Soldedebex').AsFloat;
  finally
      if Assigned(Q) then  Ferme(Q);
  end;

  // Solde des écritures du Compte ICc
  Q := nil;
  try
    // Solde du Compte ICC
    Q := OpenSQL('Select sum(ice_debit - ice_credit) as Total from iccecriture where ' +
                 '(ice_general = "' + vStCompte +'") and ' +
                 '(ice_datevaleur >= "' + USDateTime(D1) + '") and ' +
                 '(ice_datevaleur <= "' + USDateTime(D2) + '")', True);

    if not Q.EOF then
      SoldeIcc := SoldeIcc + Q.FindField('Total').AsFloat;
  finally
    if Assigned(Q) then  Ferme(Q);
  end; *)

  // Solde du compte général
  //SoldeCompta := CGetCumulParDate( D1 , D2 , vStCompte );
  //CGetBalanceParcompte ('', vStCompte , D1 , D2 , MtDeb , MtCre , SoldeCompta );
  LesBalances := TStringList.Create;
  SoldeCompta := GetCumul('GEN', vStCompte, '', '', '', '', '', D1, D2, True, True, LesBalances, LeResult , False);
  LesBalances.Free;

  Result := SoldeCompta;

end;

{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 09/10/2001
Modifié le ... : 10/10/2001
Description .. : Vérification si des nouvelles écritures ont été saisies sur le
Suite ........ : compte passé en paramètre
Mots clefs ... :
*****************************************************************}
function ExisteNouvelleEcriture ( vStCompte : String ; vD1, vD2 : TDateTime ) : Boolean;
var st : string;
begin
  st := 'Select E_QUALIFORIGINE from ECRITURE where '+
                      '(E_GENERAL = "' + vStCompte + '") and ' +
                      '(E_QUALIFPIECE = "N") and '+
                      '(E_ECRANOUVEAU = "N" or E_ECRANOUVEAU = "H") and ' +
                      '(E_DATECOMPTABLE >= "' + UsDateTime( vD1 ) + '") and ' +
                      '(E_DATECOMPTABLE <= "' + UsDateTime( vD2 ) + '") and ' +
                      '(E_QUALIFORIGINE <> "ICC")';
  Result := ExisteSQL(st);
end;

{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 29/11/2001
Modifié le ... :   /  /
Description .. : Recherche la valeur de départ pour la clé ICE_LIGNE
Mots clefs ... :
*****************************************************************}
function GetNumLigneFromEcriture( vStGeneral : String ; vD1 : TDateTime ) : Integer;
var Sql : TQuery;
begin
//  Result := 0;
  Sql := nil;
  try
    Sql := OpenSql('Select Max(ICE_LIGNE) As CleUnique from ICCECRITURE where ' +
                   '(ICE_GENERAL = "' + vStGeneral + '") and ' +
                   '(ICE_DATEVALEUR < "' + USDateTime( vD1 ) + '")', True);

    Result := Sql.FindField('CleUnique').AsInteger;

    if Result = 0 then Result := 1;
  finally
    if Assigned(SQL) then Ferme(SQL);
  end;
end;











































{==============================================================================}
{ TICCCalcul }
{==============================================================================}
constructor TICCCalcul.Create;
begin

end;

{==============================================================================}
Procedure TICCCalcul.CalculTOB(T :TOB; QCompte : TQuery);
Var i : Integer;
    TobFille : TOB;
    Solde : Double;
Begin
  FCompte := QCompte.FindField('ICG_GENERAL').AsString;                // Numéro du compte
  FGroupe := (QCompte.FindField('ICG_GROUPE').AsString = 'X');         // Compte de Groupe
  FLimitation := (QCompte.FindField('ICG_LIMITATION').AsString = 'X'); // Limitation légale du taux
  FDirigeant := (QCompte.FindField('ICG_DIRIGEANT').AsString = 'X');   // Dirigeant ou plus de 50 % du capital
  FTaux := QCompte.FindField('ICG_TXANNUEL').AsFloat;                  // Taux du Compte
  FDateBascule := QCompte.FindField('ICG_DATEEFFET').AsDateTime;       // La date d'effet

  // CALCUL DU SOLDE ANTERIEUR A LA DATE DE DEBUT DU CALCUL
  Solde := QCompte.FindField('ICG_SOLDEDEBEX').AsFloat;

  i:=0;
  while (i <> T.Detail.Count) do
  begin
    if (T.Detail[i].GetValue('ICE_DATEVALEUR') < Icc_Data.DateDu) then
    begin
      Solde := Solde + T.Detail[i].GetValue('ICE_DEBIT') - T.Detail[i].GetValue('ICE_CREDIT');
      T.Detail[i].Free;
    end
    else
      Inc(i);
  end;

  // Prise en compte du solde antérieur par rapport à la date de début du calcul
  if Solde <> 0 then
  begin
    if (T.Detail.Count = 0) or (T.Detail[0].GetValue('ICE_DATEVALEUR') <> Icc_Data.DateDu) then
    begin
      TobFille := Tob.Create('ICCECRITURE', T, 0);
      TobFille.PutValue('ICE_GENERAL', FCompte);
      TobFille.PutValue('ICE_DATEVALEUR', Icc_Data.DateDu);
      TobFille.PutValue('ICE_LIBELLE', 'RECAPITULATIF');

      if Solde >= 0 then
      begin
        TobFille.PutValue('ICE_CREDIT',0);
        TobFille.PutValue('ICE_DEBIT',Abs(Solde));
      end
      else
      begin
        TobFille.PutValue('ICE_CREDIT',Abs(Solde));
        TobFille.PutValue('ICE_DEBIT',0);
      end;
    end
    else
    begin // Ajout du Solde au montant de la premiere ecriture
      if Solde >= 0 then
        T.Detail[0].PutValue('ICE_DEBIT',Abs(Solde) + T.Detail[0].GetValue('ICE_DEBIT'))
      else
        // GCO - 26/03/2004 - FB 11733
        //T.Detail[0].PutValue('ICE_CREDIT',Solde - T.Detail[0].GetValue('ICE_CREDIT'));
        T.Detail[0].PutValue('ICE_CREDIT',Abs(Solde) + T.Detail[0].GetValue('ICE_CREDIT'));
    end;
  end;

  if FDirigeant then
  begin
    // AJOUT DE L ECRITURE CAUSE PAR LA DATE D'EFFET : LE COMPTE DEVIENT DE TYPE DIRIGEANT
    if (FDateBascule >= Icc_Data.DateDu) and (FDateBascule <= Icc_Data.DateAu) then
    begin
      TobFille := T.FindFirst(['ICE_DATEVALEUR'],[FDateBascule],False);
      if TobFille = nil then // Pas trouvé
      begin
        TobFille := Tob.Create('ICCECRITURE',T,-1);
        TobFille.PutValue('ICE_DATEVALEUR',FDateBascule);
        TobFille.PutValue('ICE_CREDIT',0);
        TobFille.PutValue('ICE_DEBIT',0);
      end;
    end;

    // Ajout des ecritures des autres comptes de type dirigeant
    RechercheEcriture(T);
  end;

  //Recherche les variations possibles du tauxx du compte
  RechercheVariationTaux(T);

  T.Detail.Sort('ICE_DATEVALEUR');

  // Ajout des nouveaux champs pour effectuer les calculs
  if T.Detail.Count <> 0 then // GCO - 18/01/2005 - FQ 14905
  begin
    T.Detail[0].AddChampSupValeur('ICE_SOLDE', 0, True);
    T.Detail[0].AddChampSupValeur('ICE_SOLDE2', '', True);
    T.Detail[0].AddChampSupValeur('ICE_CAPITAL', 0, True);
    T.Detail[0].AddChampSupValeur('ICE_CAPITAL2', 0, True);
    T.Detail[0].AddChampSupValeur('ICE_TAUX', 0, True);
    T.Detail[0].AddChampSupValeur('ICE_TAUXMAX', Icc_Data.TauxLegal, True);
    T.Detail[0].AddChampSupValeur('ICE_NBJOURS', 0, True);
    T.Detail[0].AddChampSupValeur('ICE_BASE', 0, True);
    T.Detail[0].AddChampSupValeur('ICE_INTERET', 0, True);
    T.Detail[0].AddChampSupValeur('ICE_SURCAPITAL', 0, True);
    T.Detail[0].AddChampSupValeur('ICE_SURTAUX', 0, True);
  end;
  CalcDonnees(T);
End;

{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE & Christophe AYEL
Créé le ...... : 06/04/2000
Modifié le ... : 06/04/2000
Description .. : Calcul le nombre de jours entre 2 dates pour une
Suite ........ : année pour 360 jours ou 365 jours.
Mots clefs ... : DATE;360 JOURS
*****************************************************************}
Function TICCCalcul.SoustraitDate(Date1, Date2 : TDateTime ; Mode : Integer): Double;
var nbJourD1,nbJourD2 : Double;
    PremMois, PremAnnee, NbMois : Word;
    Year, Month, Day : Word;
    Year_f, Month_f, Day_f : Word;
Begin
  (* ---- Mode 360 jours par Année ---- *)
  if Mode = 360 then
  begin
    if Date1 = Date2+1 then
    begin
      Result := 0;
      Exit;
    end;

    (* --- Principe du jour à la veille --- *)
    DecodeDate(Date1,Year,Month,Day);
    if Day = 31 then
      Date1 := Date1 - 1;

    if Date2 <> ICC_Data.DateAU then
    begin
      DecodeDate(Date2,Year,Month,Day);
      // Si day = 30, ça veut dire qu'il etait saisie au 31, donc on le ramene au 29
      if Day = 30 then
        Date2 := Date2 - 1;
    end;

    //-- Calcul du nombre de mois plein ---
    NombreMois(Date1,Date2,PremMois,PremAnnee,NbMois);

    if NbMois >= 2 then
    begin
      DecodeDate(Date1,Year,Month,Day);
      (*if (Month = 2) then
        Date2 := EncodeDate(Year_f, Month, 28)
        else
          Date2 := EncodeDate(Year_f, Month, 30);
      end;*)
      NbJourD1 := 30 - Day + 1;

      // Nombre de jours du dernier mois de la fourchette
      //DecodeDate(Date2,Year,Month,Day);
      //DecodeDate(FinDeMois(Date2),Year_f,Month_f,Day_f);

      // Calcul du nombre de jour de D2
      DecodeDate(DebutDeMois(Date2), Year, Month, Day);
      DecodeDate(Date2, Year_f, Month_f, Day_f);
      if Date2 = FinDeMois(Date2) then
        NbJourD2 := 30 - Day + 1
      else
        nbJourD2 := Day_F - Day + 1;

      // Ok on ne touche plus
      if NbMois > 2 then
        Result := (NbMois-2) * 30 + nbJourD1 + nbJourD2
      else
        Result := nbJourD1 + nbJourD2;
    end
    else
    begin
      DecodeDate(Date1, Year, Month, Day);
      DecodeDate(Date2, Year_f, Month_f, Day_f);

      if Date2 = FinDeMois(Date2) then
        Result := 30 - Day + 1
      else
        Result := Day_F - Day + 1;
    end;

    // if Date1 = VH^.Encours.Deb then
    // Result := Result + 1;

    //if Date2 = VH^.Encours.Fin then
    // Result := Result + 1;

  end
  else (* ---- Mode nombre de jours réels dans l'année (365 ou 366 jours) ---- *)
    Result := (Date2 - Date1) + 1;
End;

{=============== A.G.L.=========================================================
Auteur  ...... : Gilles COSTE
Créé le ...... : 29/03/2000
Modifié le ... : 29/03/2000
Description .. : Recherche des ecritures des autres comptes de capital pour
                 rechercher des changements qui vont modifier dans le calcul
Mots clefs ... :
===============================================================================}
procedure TICCCalcul.RechercheEcriture(LaTob: TOB);
var Q : TQuery;
    TobCalcul,TobTemp : Tob;
    i : integer;
begin
  Q := nil;
  try
    Q := OpenSQL('Select ice_datevaleur from iccecriture left join iccgeneraux on ice_general=icg_general ' +
                 'where (ice_datevaleur >= "' + USDateTime(Icc_Data.DateDu) + '" and ice_datevaleur >= "' + UsDateTime(FDateBascule) + '" and ice_datevaleur <= "' + UsDateTime(Icc_Data.DateAu) + '" and ' +
                 'icg_dirigeant ="X" and icg_general <> "' + FCompte + '"' + ') group by ice_datevaleur',True);

    TobCalcul := Tob.Create('TobCalcul', nil, -1);
    TobCalCul.LoadDetailDB('ICCECRITURE', '', '', Q, True, False);
  finally
    if Assigned(Q) then Ferme(Q);
  end;

  TobCalcul.Detail.Sort('ICE_DATEVALEUR');

  (* ---- Suppresion des lignes de TobCalcul qui sont déjà dans LaTob ---- *)
  (* ---- Il ne peut en y avoir qu'une par date à cause du groupe by ICE_DATEVALEUR sur les autres comptes --- *)
  (* ---- AJOUT DES AUTRES ECRITURES DE LA TOBCALCUL DANS LATOB ---- *)
  i := 0;
  while (i <> TobCalcul.Detail.Count) do
  begin
    TobTemp := LaTob.FindFirst(['ICE_DATEVALEUR'],[TobCalcul.Detail[i].GetValue('ICE_DATEVALEUR')],False);
    if TobTemp = nil then
    begin // Pas Trouvé dans LATOB
      TobTemp := Tob.Create('ICCECRITURE', LaTob, -1);
      //TobTemp.Dupliquer(TobCalcul.Detail[i],True,True,False);
      TobTemp.PutValue('ICE_GENERAL', FCompte);
      TobTemp.PutValue('ICE_DATEVALEUR', TobCalcul.Detail[i].GetValue('ICE_DATEVALEUR'));
    end
    else (* --- Trouvé dans LATOB alors Suppression dans TobCalcul--- *)
      TobCalcul.Detail[i].Free;
  end;
  TobCalcul.Free;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 26/11/2001
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TICCCalcul.RechercheVariationTaux(LaTob: TOB);
var Sql : TQuery;
    lTobFille : Tob;
begin
  Sql := nil;
  try
    Sql := OpenSQL('Select Icd_datedu,icd_dateau,icd_taux from ICCTauxCompte where ' +
                   'icd_general = "' + FCompte + '" and ' +
                   'icd_datedu >= "' + USDateTime(Icc_Data.DateDu) + '" and '+
                   'icd_dateau <= "' + UsDateTime(Icc_Data.DateAu) + '" order by icd_general,icd_datedu',True);
    if not (Sql.EOF) then
    begin
      Sql.First;
      while not Sql.Eof do
      begin
        lTobFille := LaTob.FindFirst(['ICE_DATEVALEUR'],[Sql.FindField('ICD_DATEDU').AsDateTime],False);
        if lTobFille = Nil then
        begin
          lTobFille := Tob.Create('ICCECRITURE', LaTob, -1);
          lTobFille.PutValue('ICE_GENERAL', FCompte);
          lTobFille.PutValue('ICE_DATEVALEUR', Sql.FindField('ICD_DATEDU').AsDateTime);
        end;

        lTobFille := LaTob.FindFirst(['ICE_DATEVALEUR'],[Sql.FindField('ICD_DATEAU').AsDateTime+1],False);
        if lTobFille = Nil then
        begin
          lTobFille := Tob.Create('ICCECRITURE', LaTob, -1);
          lTobFille.PutValue('ICE_GENERAL', FCompte);
          lTobFille.PutValue('ICE_DATEVALEUR', Sql.FindField('ICD_DATEAU').AsDateTime+1);
        end;

        Sql.Next;
      end;
    end;

  finally
    if Assigned(Sql) then Ferme(Sql);
  end;

end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : COSTE Gilles
Créé le ...... : 05/09/2001
Modifié le ... : 05/09/2001
Description .. : Recherche le taux à appliquer pour unen ecriure ICC
Mots clefs ... :
*****************************************************************}
function TICCCalcul.RechercheTaux(lDate: TDatetime): Double;
var SQL : TQuery;
begin
//  Result := 0;
  SQL := nil;
  try
    SQL := OpenSql('Select ICD_TAUX from ICCTAUXCOMPTE where (ICD_GENERAL = "' + FCompte + '") and ' +
                   '(ICD_DATEDU <= "' + USDateTime(lDate) + '") and ' +
                   '(ICD_DATEAU >= "' + USDateTime(lDate) + '") ',True);

    if not SQL.EOF then
      Result := SQL.FindField('ICD_TAUX').AsFloat
    else
        Result := FTaux;
  finally
    if Assigned(SQL) then Ferme(SQL);
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : COSTE Gilles
Créé le ...... : 29/08/2001
Modifié le ... : 29/08/2000
Description .. :
Suite ........ : lmk
Mots clefs ... : TOTO
*****************************************************************}
function TICCCalcul.CalculCapital(LaDate: TDateTime; LeSolde: Double): Double;
var Q : TQuery;
    TotCapital : Double;
    TotComptes : Double;
begin
  Result := 0;
  TotCapital := 0;
  TotComptes := 0;
  Q := nil;

  // Solde début exercice du compte de capital
  try
    Q := OpenSQL('Select icg_general, icg_soldedebex from iccgeneraux where (ICG_GENERAL = "'+GetParamSocSecur('SO_ICCCOMPTECAPITAL','')+'")',True);
    if not Q.EOF then
      TotCapital := Q.Findfield('icg_soldedebex').AsFloat;
  finally
    if Assigned(Q) then Ferme(Q);
  end;

  // Cumul des écritures du compte de capital avant LADATE
  Q := nil;
  try
    Q := OpenSQL('Select sum(ICE_DEBIT - ICE_CREDIT) as Cumul from iccecriture where (ice_general = "' + GetParamSocSecur('SO_ICCCOMPTECAPITAL','') + '") and (ICE_DATEVALEUR < "' + USDateTime(LaDate) + '")', True);
    TotCapital := (TotCapital + Q.FindField('CUMUL').AsFloat) * 1.5 ;
  finally
    if Assigned(Q) then Ferme(Q);
  end;

  // On retourne zéro si le compte de capital est débiteur
  if TotCapital >= 0 then Exit;

  // Solde de début des autres comptes concernés
  Q := nil;
  try
    Q := OpenSQL('SELECT SUM(ICG_SOLDEDEBEX)AS CUMUL FROM ICCGENERAUX WHERE (ICG_DIRIGEANT = "X") AND (ICG_GENERAL <> "'+GetParamSocSecur('SO_ICCCOMPTECAPITAL','')+'")',True);
    if not Q.Eof then
      TotComptes := Q.FindField('Cumul').AsFloat;
  finally
    if Assigned(Q) then Ferme(Q);
  end;

  // Cumul des écritures des autres comptes concernés
  Q := nil;
  try
    Q := OpenSQL('Select sum( ice_debit - ice_credit ) as Cumul from iccecriture left join iccgeneraux on ice_general = icg_general ' +
                 'where (icg_dirigeant = "X") and (icg_general <> "' + GetParamSocSecur('SO_ICCCOMPTECAPITAL','')+'") and (ice_datevaleur < "' + UsDateTime(LaDate) + '")',True);
    if not Q.Eof then
      TotComptes := TotComptes + Q.FindField('Cumul').AsFloat;
  finally
    if Assigned(Q) then Ferme(Q);
  end;

  // Vérification des résultats
  if (TotComptes <> 0) then
  begin
    Result := (LeSolde * TotCapital) / TotComptes;

    if Result > 0 then
      Result := 0
    else
    begin
      if Result < LeSolde then
        Result := LeSolde;
    end;
  end;

end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : COSTE Gilles
Créé le ...... : 29/08/2001
Modifié le ... : 29/08/2001
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TICCCalcul.CalcDonnees(Tob1 : TOB);
var Debit, Credit : Double;
    i : integer;
    Base, Interet : Double;
    D1, D2 : TDatetime;
begin
  // Initialisation des variables
  Debit  := 0;
  Credit := 0;

  i := 0;
  while (i <> Tob1.Detail.Count) do
  begin
    // Solde
    Debit := Debit + Tob1.Detail[i].GetValue('ICE_DEBIT');
    Credit := Credit + Tob1.Detail[i].GetValue('ICE_CREDIT');

    Tob1.Detail[i].PutValue('ICE_SOLDE',  Debit - Credit);
    Tob1.Detail[i].PutValue('ICE_SOLDE2', AfficheDBCR(Debit - Credit));

    // Nombre de jours
    D1 := Tob1.Detail[i].GetValue('ICE_DATEVALEUR');
    // Si fin de la TOB, on prend la Date de fin du calcul, sinon Date de la TOB Fille suivante -1 jour
    if i = Tob1.Detail.Count -1 then
      D2 := Icc_Data.DateAu
    else
      D2 := Tob1.Detail[i+1].GetValue('ICE_DATEVALEUR')-1;

    Tob1.Detail[i].PutValue('ICE_NBJOURS',FormatFloat('###0',SoustraitDate(D1,D2,ICC_Data.NombreJours)));

    if Tob1.Detail[i].GetValue('ICE_NBJOURS') <> 0 then
    begin
      // Capital autorise
      if (FDirigeant) and (Tob1.Detail[i].GetValue('ICE_DATEVALEUR') >= FDateBascule) then
      begin
        //if Tob1.Detail[i].GetValue('ICE_SOLDE') <> 0 then
        if Tob1.Detail[i].GetValue('ICE_SOLDE') < 0 then
        begin
          if i <> Tob1.Detail.Count-1 then
            Tob1.Detail[i].PutValue('ICE_CAPITAL',CalculCapital(Tob1.Detail[i+1].GetValue('ICE_DATEVALEUR'),Debit-Credit))
          else
            // GCO - 26/09/2007 - FQ 21353 ajout de +1 à la date
            Tob1.Detail[i].PutValue('ICE_CAPITAL',CalculCapital(Icc_Data.DateAu+1,Debit-Credit));
        end;
      end;

      Tob1.Detail[i].PutValue('ICE_CAPITAL2',AfficheDBCR(Tob1.Detail[i].GetValue('ICE_CAPITAL')));

      // Taux du compte
      Tob1.Detail[i].PutValue('ICE_TAUX',RechercheTaux(Tob1.Detail[i].GetValue('ICE_DATEVALEUR')));

      // Base du calcul
      Base := (-1) * (Tob1.Detail[i].GetValue('ICE_SOLDE') * Tob1.Detail[i].GetValue('ICE_NBJOURS')) / ICC_Data.NombreJours;
      Tob1.Detail[i].PutValue('ICE_BASE',Base);

      // Interets déductibles
      if (FGroupe) or ((not FGroupe) and (Tob1.Detail[i].GetValue('ICE_SOLDE') <= 0)) then // MBAMF
      begin
        Interet := (Base * Tob1.Detail[i].GetValue('ICE_TAUX')) / 100;
        Tob1.Detail[i].PutValue('ICE_INTERET',Interet);
      end;

      // Intérets non déductibles
      if (not FGroupe) then
      begin
        if (FDirigeant) and
           (Tob1.Detail[i].GetValue('ICE_SOLDE') < 0) and
           (Tob1.Detail[i].GetValue('ICE_CAPITAL') <= 0) and
           (Tob1.Detail[i].GetValue('ICE_DATEVALEUR') >= FDateBascule) then
          NonDeductibleSurCapital(Tob1.Detail[i]);

        if (FLimitation) and
           (Tob1.Detail[i].GetValue('ICE_SOLDE') < 0) and
           (Tob1.Detail[i].GetValue('ICE_TAUX') > Tob1.Detail[i].GetValue('ICE_TAUXMAX')) then
          NonDeductibleSurTaux(Tob1.Detail[i]);
      end;

      Inc(i);
    end
    else
      Tob1.Detail[i].Free;
  end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 17/09/2001
Modifié le ... : 17/09/2001
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TICCCalcul.NonDeductibleSurCapital( LaTob : TOB );
var Valeur : Double;
begin
//  Valeur := 0;

  if Icc_Data.Methode = 0 then
  begin
    // Méthode ICC
    if (LaTob.GetValue('ICE_TAUX') > LaTob.GetValue('ICE_TAUXMAX')) then
      Valeur := (LaTob.GetValue('ICE_SOLDE') - LaTob.GetValue('ICE_CAPITAL')) * LaTob.GetValue('ICE_TAUXMAX') * LaTob.GetValue('ICE_NBJOURS')
    else
      Valeur := (LaTob.GetValue('ICE_SOLDE') - LaTob.GetValue('ICE_CAPITAL')) * LaTob.GetValue('ICE_TAUX') * LaTob.GetValue('ICE_NBJOURS');
  end
  else
  begin
    // Méthode Lefevre
    //if (LaTob.GetValue('ICE_TAUX') > LaTob.GetValue('ICE_TAUXMAX')) then
    //  Valeur := (LaTob.GetValue('ICE_SOLDE') - LaTob.GetValue('ICE_CAPITAL')) *  LaTob.GetValue('ICE_TAUX') * LaTob.GetValue('ICE_NBJOURS')
    //else
      Valeur := (LaTob.GetValue('ICE_SOLDE') - LaTob.GetValue('ICE_CAPITAL')) *  LaTob.GetValue('ICE_TAUX') * LaTob.GetValue('ICE_NBJOURS');
  end;

  Valeur := Arrondi(Valeur / (100 * ICC_Data.NombreJours),V_PGI.OkDecV);

  LaTob.PutValue('ICE_SURCAPITAL', Abs(Valeur));
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 17/09/2001
Modifié le ... : 17/09/2001
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TICCCalcul.NonDeductibleSurTaux(LaTob: TOB);
var Valeur : Double;
begin
//  Valeur := 0;

  if Icc_Data.Methode = 0 then
    // Méthode ICC
    Valeur := LaTob.GetValue('ICE_SOLDE') * (LaTob.GetValue('ICE_TAUX')- LaTob.GetValue('ICE_TAUXMAX')) * LaTob.GetValue('ICE_NBJOURS')
  else
    // Méthode Lefevre
    Valeur := LaTob.GetValue('ICE_CAPITAL') * (LaTob.GetValue('ICE_TAUX')- LaTob.GetValue('ICE_TAUXMAX')) * LaTob.GetValue('ICE_NBJOURS');

  Valeur := Arrondi(Valeur / (100 * ICC_Data.NombreJours),V_PGI.OkDecV);

  LaTob.PutValue('ICE_SURTAUX',Abs(Valeur));

end;

destructor TICCCalcul.Destroy;
begin
  inherited;
end;

END.
