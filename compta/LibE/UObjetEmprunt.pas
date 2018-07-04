unit UObjetEmprunt;

interface

uses
    classes,
    Sysutils,
    HCtrls,
    Math,
    HEnt1,
    UTob,
    dialogs,
{$IFDEF VER150}
    Variants,
{$ENDIF}
{$IFDEF EAGLCLIENT}
    UtileAgl,
 {$ELSE}
    {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
 {$ENDIF}
    paramsoc;

type
    // -------------------------------------------------------------------------
    //
    //                      T Y P E S   E N U M E R E S
    //
    // -------------------------------------------------------------------------
    TTypeRecherche = (trCapital, trDuree, trTaux, trVersement, trAucune);

    TTestEmpruntResult = (terOk, terInsuffVers, terTropVers, terInsuffAmort, terTropAmort);

    TColonneEcheance = (ceDate, ceTauxInteret, ceTauxAssurance, ceAmortissement, ceVersement, ceMntAssurance);

    TStatutEmprunt = (seVide, seCalcule, seModifie, seSaisi, seSimule);


    // -------------------------------------------------------------------------
    //
    //                O B J E T   M E T I E R   E M P R U N T
    //
    // -------------------------------------------------------------------------
    TObjetEmprunt = class(TOB)
    private
       fRdTotalReport : double;      // montant total du report de charges (int+ass)

       fBoTauxReel         : Boolean;              // Utilisé pour le calcul en taux variable
       fRdReportDeBase_int : double;            // Reports de charges de base à utiliser
       fRdReportDeBase_ass : double;            // dès le début du calcul de l'emprunt

       fObTobExercices : TOB;

//       fInMoisExercice : Integer;
//       fInJourExercice : Integer;
       // Evenements

       // Fonctions d'accès aux propriétés
       function  GetTauxTotal : Double;
       function  GetVersementMaximum : double;
       Function  GetNbEchAn : Integer;
       Function  GetNbMoisEch : Integer;
       Procedure SetValeurPK(pInValeur : Integer);

       // Fonctions de calculs
       Function  Tricher(i : Integer; dette : double) : Boolean;
       Function  CalcTauxEcheance(pInIndexe : Integer; pRdTauxAnnuel : Double) : Double;
       Function  CalcTauxEqui(tx, nb_ech : Double) : Double;
       Function  CalcTauxPro(tx, nb_ech : Double) : Double;
       Function  CalculerK : Double;
       Function  EmpruntValide(pInTypeRecherche : TTypeRecherche) : TTestEmpruntResult;
       Procedure CalculerEmprunt(pInTypeRecherche : TTypeRecherche; pBoArrDerEch : Boolean);
       function  CalculerNbEch(pInTypeRecherche : TTypeRecherche) : Integer;
       Function  DateEcheance(pInEcheance : Integer) : TDateTime;
       Procedure AffMessage(pStMessage : String);
       Procedure Simulation1(pInTypeRecherche : TTypeRecherche);
       Procedure Simulation2(pInTypeRecherche : TTypeRecherche);
       Procedure Simulation3(pInTypeRecherche : TTypeRecherche);
       Function  Genius(pInTypeRecherche : TTypeRecherche) : Double;
//       Procedure SetDateFinExercice(pDtDate : TDateTime);

    protected

    public

       // Propriétés de l'emprunt
       Echeances         : TOB;
       TauxVariables     : TOB;

//       Function  GetText : String;

       Constructor Create(LenomTable : String; LeParent : TOB; IndiceFils : Integer); Override;
       Constructor CreerEmprunt;
       Destructor  Destroy; Override;

       Function  DBChargerEmprunt(pInCodeEmprunt : Integer = -1) : Boolean;
       Procedure DBSupprimerEmprunt;
       Procedure DBSupprimerEcheances;
       Procedure DBSupprimeEtRemplaceEmprunt;
       Procedure DBSupprimeEtRemplaceEcheances;

       Procedure Assigner(pOmEmprunt : TObjetEmprunt);
       Procedure Init;

       Procedure Simulation(pInTypeRecherche : TTypeRecherche);
       Function  Ajustement(pRdMontant : Double) : Boolean;
       function  CalcAmortCst : double;
       function  CalcVersCst : double;
       Function  Arrondir(pRdMontant : Double; pRdParam : Double; pRdUnit : Integer ) : Double;
       Function  FormaterMontant(pRdMontant : Double) : String;
       Function  FormaterTaux(pRdTaux : Double) : String;
       Function  ProchaineDateExercice(pDtDate : TDateTime) : TDateTime;

       Class Function CalculerPeriode(PerExo : Integer ; pDtDateRef : TDateTime; pInN : Integer; Var pDtDateDebPer : TDateTime; Var pDtDateFinPer : TDateTime; pObTobPeriode : TOB = Nil) : Boolean;
       Class Function CalculerExercice(pDtDateRef : TDateTime; pInN : Integer; Var pDtDateDebExe : TDateTime; Var pDtDateFinExe : TDateTime; pObTobExercice : TOB = Nil) : Boolean;

       // gestion des taux variables
       Procedure SupprimerUnTaux(pDtDate : TDateTime);
       Procedure AjouterTaux(pDtDate : TDateTime; pRdTaux : Double);
       Function  RechercherTaux(pDtDate : TDateTime) : Double;

       // Prochain n° d'emprunt
       Function EmpruntSuivant : Integer;

       //
       Procedure CreerEcheances(pNbEcheances : Integer);
       Procedure CreerTaux(pNbTaux : Integer);
       Procedure EcheancesVersGrille(pObGrille : THGrid);

       Property TauxTotal                     : Double    Read  GetTauxTotal;
       Property VersementMaximum              : double    Read  GetVersementMaximum;
       Property TotalReport                   : double    Read  fRdTotalReport;
       Property EMP_NBECHAN                   : Integer   Read  GetNbEchAn;
       Property EMP_NBMOISECH                 : Integer   Read  GetNbMoisEch;
       Property ValeurPK                      : Integer   Write SetValeurPK;
//       Property JourExercice                  : Integer   Read  fInJourExercice Write fInJourExercice;
//       Property MoisExercice                  : Integer   Read  fInMoisExercice Write fInMoisExercice;
//       Property DateFinExercice               : TDateTime Write SetDateFinExercice;

    published


    end;

    function TrouverCompteAssocies(anom : string):string;
const

  // -------------------------------------------------------------------------
  //
  //                          C O N T A N T E S
  //
  // -------------------------------------------------------------------------

   // Constantes des tablettes

   // Types de période
   tpMois      = 'MOI';
   tpTrimestre = 'TRI';
   tpSemestre  = 'SEM';
   tpAnnee     = 'ANN';

   // Types de taux
   ttConstant     = 'CST';
   ttVariable     = 'VAR';
   ttArythmetique = 'ARY';
   ttGeometrique  = 'GEO';

   // Banquaire / équivalent
   beBanquaire  = 'TBQ';
   beEquivalent = 'TEQ';

   // Type de l'assurance
   taAucune  = 'NON';
   taMontant = 'MNT';
   taTaux    = 'TAU';

   // Type de versement
   tvConstant = 'CST';
   tvVariable = 'VAR';

   // Base de l'assurance
   tbaDette = 'DET';
   tbaSolde = 'SOL';

   // Constantes diverses
   cInMaxEch   = 600;         // Nombre maxi d'échéances
   cRdErreur   = 1/100;       // Erreur admise
   cRdTxMarge  = 1.95;        // Taux de marge
   cInNbDecMnt = 2;           // Nombre de décimales des montants
   cInNbDecTx  = 6;           // Nombre de décimales des taux
   cInMaxIter  = 1000;        // Nombre maxi d'itérations dans Genius
   cRdBorneSup = 0.95;        // Utilisée dans les calculs
   cRdBorneInf = 0.005;       // "                       "

   // Colonnes de la grille des échéances
   cColNumEch        = 0;
   cColDate          = 1;
   cColTxInteret     = 2;
   cColTxAssurance   = 3;
   cColTxGlobal      = 4;
   cColInteret       = 5;
   cColAssurance     = 6;
   cColTotInteret    = 7;
   cColAmortissement = 8;
   cColVersement     = 9;
   cColSolde         = 10;
   cColReport        = 11;
   cColDette         = 12;
   cColCumulA        = 13;
   cColCumulI        = 14;
   cColCumulM        = 15;
   cColTypeLigne     = 16;
   cColModif         = 17;
   cNbColEcheances   = 18;

   cChampsColonnes = ';ECH_DATE;ECH_TAUXINT;ECH_TAUXASS;;ECH_INTERET;ECH_ASSURANCE;' +
                     ';ECH_AMORTISSEMENT;ECH_VERSEMENT;ECH_SOLDE;;ECH_DETTE;' +
                     'ECH_CUMULA;ECH_CUMULI;ECH_CUMULV;;ECH_MODIF';

   // Types de lignes dans la grille des échéances
   cStLigneEcheance    = 'E';
   cStLigneTotal       = 'T';

   // Formats des colonnes
   cFormatTaux    = '##0.000000';
   cFormatMontant = '###,###,##0.00';



implementation

// ----------------------------------------------------------------------------
// Nom    :
// Date   : 13/11/2001
// Auteur : D. ZEDIAR
// Objet  :
// ----------------------------------------------------------------------------
// Historique :
// ----------------------------------------------------------------------------




// ----------------------------------------------------------------------------
//
//                    F O N C T I O N S   D I V E R S E S
//
// ----------------------------------------------------------------------------



// ----------------------------------------------------------------------------
//
//     G E S T I O N   D E   L ' O B J E T   M E T I E R   E M P R U N T
//
// ----------------------------------------------------------------------------
{***********A.G.L.***********************************************
Auteur  ...... : Zediar
Créé le ...... : 10/12/2001
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
// ----------------------------------------------------------------------------
// Nom    : FormaterMontant
// Date   : 22/11/2001
// Auteur : D. ZEDIAR
// Objet  : Formatage (string) d'un montant
// ----------------------------------------------------------------------------
// Historique :
// ----------------------------------------------------------------------------
Function TObjetEmprunt.FormaterMontant(pRdMontant : Double) : String;
begin
   Result := StrfMontant(pRdMontant,15,cInNbDecMnt,'',False);
end;

Function  TObjetEmprunt.FormaterTaux(pRdTaux : Double) : String;
begin
   Result := StrfMontant(pRdTaux,15,cInNbDecTx,'',False);
end;


// ----------------------------------------------------------------------------
// Nom    : Arrondir
// Date   : 21/11/2001
// Auteur : D. ZEDIAR
// Objet  : fonction d'arrondi
// montant = valeur à arrondir
//    param = paramètre d'arrondi en puissance de 10 ex = pour arrondir au 1/10
//                                                        unit = -1
//                                                        pour arrondir à 100
//                                                        unit = 2
//    unit = unités du paramètre d'arrondi
//    montant_arr valeur arrondie retournée
// ----------------------------------------------------------------------------
// Historique :
// ----------------------------------------------------------------------------
function TObjetEmprunt.Arrondir( pRdMontant : Double; pRdParam : Double; pRdUnit : Integer ) : Double;
var
   dec : Double;
begin
   pRdUnit  := -pRdUnit;
   Result   := arrondi(pRdMontant,pRdUnit+1);
   pRdParam := pRdParam / 10;
   dec      := frac(Result * Power(10,pRdUnit));

   if (dec <> 0) and (dec >= pRdParam) then
      Result := Result + (1 - dec) *  Power(10,-pRdUnit)
   else
      Result := pRdMontant - dec * Power(10,-pRdUnit);
end;


// ----------------------------------------------------------------------------
// Nom    : Create
// Date   : 13/11/2001
// Auteur : D. ZEDIAR
// Objet  : Constructeur
// ----------------------------------------------------------------------------
// Historique :
// ----------------------------------------------------------------------------
Constructor TObjetEmprunt.Create(LenomTable : String; LeParent : TOB; IndiceFils : Integer);
begin
   Inherited;
   Echeances         := Tob.Create('',Self,-1);
   TauxVariables     := Tob.Create('',Self,-1);

   fObTobExercices := Tob.Create('',Nil,-1);
   fObTobExercices.LoadDetailFromSQL('Select EX_DATEDEBUT, EX_DATEFIN From EXERCICE Order By EX_DATEDEBUT');

//   fInMoisExercice := 1;
//   fInJourExercice := 1;

   init;
end;

Constructor TObjetEmprunt.CreerEmprunt;
begin
   Create('FEMPRUNT',Nil,-1);
end;

// ----------------------------------------------------------------------------
// Nom    : destroy
// Date   : 13/11/2001
// Auteur : D. ZEDIAR
// Objet  : destructeur
// ----------------------------------------------------------------------------
// Historique :
// ----------------------------------------------------------------------------
Destructor  TObjetEmprunt.destroy;
begin
   fObTobExercices.Free;

   inherited;
end;



// Chargement de l'emprunt
Function TObjetEmprunt.DBChargerEmprunt(pInCodeEmprunt : Integer = -1) : Boolean;
var
   lObQuery : TQuery;
begin
   Init;
   Result := False;

   if pInCodeEmprunt = -1 then
      pInCodeEmprunt := GetValue('EMP_CODEEMPRUNT');

   // Chargement de l'entête
   lObQuery := Nil;
   try
      lObQuery := OpenSQL('Select * From FEMPRUNT Where EMP_CODEEMPRUNT=' + IntToStr(pInCodeEmprunt), True);
      if lObQuery.Eof then
         Exit;
      Self.SelectDB('',lObQuery);
   finally
      if Assigned(lObQuery) Then Ferme(lObQuery);
   end;

   // Chargement des échéances
   lObQuery := Nil;
   try
      lObQuery := OpenSQL('Select * From FECHEANCE Where ECH_CODEEMPRUNT=' + IntToStr(pInCodeEmprunt), True);
      Echeances.LoadDetailDB('FECHEANCE','','ECH_DATE',lObQuery,False);
   finally
      if Assigned(lObQuery) Then Ferme(lObQuery);
   end;

   // Chargement des taux variables
   lObQuery := Nil;
   try
      lObQuery := OpenSQL('Select * From FTAUXVAR Where TXV_CODEEMPRUNT=' + IntToStr(pInCodeEmprunt), True);
      TauxVariables.LoadDetailDB('FTAUXVAR','','TXV_DATE',lObQuery,False);
   finally
      if Assigned(lObQuery) Then Ferme(lObQuery);
   end;

   Result := True;

end;


// Suppression de l'emprunt
Procedure TObjetEmprunt.DBSupprimerEmprunt;
var
   lStCodeEmprunt : String;
begin
   lStCodeEmprunt := IntToStr(GetValue('EMP_CODEEMPRUNT'));

   DBSupprimerEcheances;

   ExecuteSQL('Delete From FEMPRUNT  Where EMP_CODEEMPRUNT=' + lStCodeEmprunt);
end;


// Suppression des taux variables et des échéances
Procedure TObjetEmprunt.DBSupprimerEcheances;
var
   lStCodeEmprunt : String;
begin
   lStCodeEmprunt := IntToStr(GetValue('EMP_CODEEMPRUNT'));

   ExecuteSQL('Delete From FTAUXVAR  Where TXV_CODEEMPRUNT=' + lStCodeEmprunt);
   ExecuteSQL('Delete From FECHEANCE Where ECH_CODEEMPRUNT=' + lStCodeEmprunt);
end;


// Ajout et remplacement
Procedure TObjetEmprunt.DBSupprimeEtRemplaceEmprunt;
begin
   DBSupprimerEmprunt;

   SetAllModifie(True);
   Echeances.SetAllModifie(True);
   TauxVariables.SetAllModifie(True);

   InsertDBByNivel(False);
end;


// Ajout et remplacement des échéances et des taux variables
Procedure TObjetEmprunt.DBSupprimeEtRemplaceEcheances;
var
   lStCodeEmprunt : String;
begin
   lStCodeEmprunt := IntToStr(GetValue('EMP_CODEEMPRUNT'));
   ExecuteSQL('Delete From FTAUXVAR  Where TXV_CODEEMPRUNT=' + lStCodeEmprunt);
   ExecuteSQL('Delete From FECHEANCE Where ECH_CODEEMPRUNT=' + lStCodeEmprunt);

   Echeances.SetAllModifie(True);
   TauxVariables.SetAllModifie(True);
   Echeances.InsertDBByNivel(False);
   TauxVariables.InsertDBByNivel(False);
end;



// ----------------------------------------------------------------------------
// Nom    : Init
// Date   : 13/11/2001
// Auteur : D. ZEDIAR
// Objet  : RAZ de l'emprunt tout entier
// ----------------------------------------------------------------------------
// Historique :
// ----------------------------------------------------------------------------
Procedure TObjetEmprunt.Init;
begin
   Inherited;

   InitValeurs;

   Putvalue('EMP_DATEDEBUT',Date);
   Putvalue('EMP_DATECONTRAT',Date);
   Putvalue('EMP_ORGPRETEUR',-1);
   Putvalue('EMP_GUIDORGPRETEUR','');
   Putvalue('EMP_CODEEMPDEV',-1);
   Putvalue('EMP_EMPPERIODE',tpMois);
   Putvalue('EMP_DUREE',1);
   Putvalue('EMP_EMPTYPETAUX',ttConstant);
   Putvalue('EMP_EMPBQEQ',beBanquaire);
   Putvalue('EMP_EMPTYPEASSUR',taAucune);
   Putvalue('EMP_EMPTYPEVERS',tvConstant);
   Putvalue('EMP_AMORTTXINI','X');
   Putvalue('EMP_EMPBASEASSUR',tbaDette);
   Putvalue('EMP_ACTUANNU','-');
   fRdTotalReport      := 0;
   fBoTauxReel         := False;
   fRdReportDeBase_int := 0;
   fRdReportDeBase_ass := 0;

   Echeances.Detail.Clear;
   TauxVariables.Detail.Clear;

end;

function TrouverCompteAssocies(anom : string):string;
var binf, bsup,ssql,xres,snom: string;
   lObQuery : TQuery;
   lObTob : TOb;
begin
  lObTob:=nil;
  xres:='';
  snom:= 'FP'+copy(anom,5,length(anom)-4);
  lObQuery := Nil;
  try
    binf:=GetParamSoc('SO_' + snom + 'INF');
    bsup:=GetParamSoc('SO_' + snom + 'SUP');
    ssql:='SELECT G_GENERAL FROM GENERAUX WHERE G_GENERAL<="'+bsup+'" AND G_GENERAL>="' +binf + '"';
    if ExisteSQL (ssql) then
    begin
        lObQuery := OpenSQL(ssql, True);
        if lObQuery.Eof then
           Exit;
        lObTob := Tob.Create('',Nil,-1);
        lObTob.LoadDetailFromSQL(ssql);
        if lObTob.Detail.Count = 0 then
         xres := ''
        else
         xres := lObTob.Detail[0].GetValue('G_GENERAL');
    end;
  finally
      if Assigned(lObTob) Then lObTob.Free;
      if Assigned(lObQuery) Then Ferme(lObQuery);
      lObQuery := Nil;
      result:=xres;
  end;
end;

// ----------------------------------------------------------------------------
// Nom    : GetTauxTotal
// Date   : 13/11/2001
// Auteur : D. ZEDIAR
// Objet  : Propriété taux total (int+ass)
// ----------------------------------------------------------------------------
// Historique :
// ----------------------------------------------------------------------------
function TObjetEmprunt.GetTauxTotal : Double;
begin
   Result := GetValue('EMP_TAUXAN');
   if GetValue('EMP_EMPTYPEASSUR') = taTaux then
      Result := result + GetValue('EMP_VALASSUR');
end;


// ----------------------------------------------------------------------------
// Nom    : GetVersementMaximum
// Date   : 13/11/2001
// Auteur : D. ZEDIAR
// Objet  : Renvoie le montant maximum de versement
// ----------------------------------------------------------------------------
// Historique :
// ----------------------------------------------------------------------------
function TObjetEmprunt.GetVersementMaximum : double;
var
   lInCpt : Integer;
   lTbEch : Tob;
begin
   Result := 0;
   for lInCpt := 0 to GetValue('EMP_NBECHTOT') - 1 do
   begin
      lTbEch := Echeances.Detail[lInCpt];
      if lTbEch.GetValue('ECH_VERSEMENT') > Result then
         Result := lTbEch.GetValue('ECH_VERSEMENT');
   end;
end;


// Nombre d'échéances par an
Function  TObjetEmprunt.GetNbEchAn : Integer;
var
   lStPeriode : String;
begin
   lStPeriode := GetValue('EMP_EMPPERIODE');

   if lStPeriode = tpMois then
      Result := 12
   else
      if lStPeriode = tpTrimestre then
         Result := 4
      else
         if lStPeriode = tpSemestre then
            Result := 2
         else
            if lStPeriode = tpAnnee then
               Result := 1
            else
               Result := 0;
end;

// Nombre de mois par échéances
Function  TObjetEmprunt.GetNbMoisEch : Integer;
var
   lStPeriode : String;
begin
   lStPeriode := GetValue('EMP_EMPPERIODE');

   if lStPeriode = tpMois then
      Result := 1
   else
      if lStPeriode = tpTrimestre then
         Result := 3
      else
         if lStPeriode = tpSemestre then
            Result := 6
         else
            if lStPeriode = tpAnnee then
               Result := 12
            else
               Result := 0;
end;


// ----------------------------------------------------------------------------
// Nom    : SetValeurPK
// Date   : 12/12/2001
// Auteur : D. ZEDIAR
// Objet  :
// ----------------------------------------------------------------------------
// Historique :
// ----------------------------------------------------------------------------
Procedure TObjetEmprunt.SetValeurPK(pInValeur : Integer);
var
   lInCpt : Integer;
begin
   Putvalue('EMP_CODEEMPRUNT',pInValeur);

   for lInCpt := 0 to Echeances.Detail.Count - 1 do
      Echeances.Detail[lInCpt].Putvalue('ECH_CODEEMPRUNT',pInValeur);

   for lInCpt := 0 to TauxVariables.Detail.Count - 1 do
      TauxVariables.Detail[lInCpt].Putvalue('TXV_CODEEMPRUNT',pInValeur);
end;


// ----------------------------------------------------------------------------
// Nom    : Genius
// Date   : 13/11/2001
// Auteur : D. ZEDIAR
// Objet  : calcul de l'élément contant par itération
// ----------------------------------------------------------------------------
// Historique :
// ----------------------------------------------------------------------------
Function TObjetEmprunt.Genius(pInTypeRecherche : TTypeRecherche) : Double;
var
// ele       : valeur de l'élément recherché
// anc_ele   : ancienne valeur de l'élément recherché
// delta_ele : variation entre l'élément trouvé et l'ancien élément
// k         : valeur calculée du capital à partir des éléments calculés
// pente     : rapport entre la variation et la nouvelle valeur calculée
// borne     : valeur qui termine l'approximation
// iter      : nombre de tentative de rapprochement de la valeur idéale

  Ech       : Tob;
  iter      : Integer;
  ele       : Double;
  delta_ele : Double;
  anc_ele   : Double;
  k         : Double;
  pente     : Double;
  borne     : Double;
//  ele_min : Double;
//  ele_max : Double;

  // Assigne l'élément constant selon le type de la recherche
  procedure AssigneValeurElement(pRdValeur : Double);
  begin
     case pInTypeRecherche of
        trDuree     : Putvalue('EMP_DUREE',pRdValeur);
        trVersement : Putvalue('EMP_VERSCST',pRdValeur);
        trCapital   : Putvalue('EMP_CAPITAL',pRdValeur);
        trTaux      : Putvalue('EMP_TAUXAN',pRdValeur);
     end;
  end;

  // Retourne l'élément constant selon le type de la recherche
  function RetourneValeurElement : Double;
  begin
     case pInTypeRecherche of
        trDuree     : Result := GetValue('EMP_DUREE');
        trVersement : Result := GetValue('EMP_VERSCST');
        trCapital   : Result := GetValue('EMP_CAPITAL');
        trTaux      : Result := GetValue('EMP_TAUXAN');
        else          Result := 0;
     end;
  end;

begin

   iter           := 0;
   fRdTotalReport := 0;
   ele            := 0;
   delta_ele      := 0;
//   ele_max       := 100;
//   ele_min       := 0;

   // Détermine la valeur de base pour ele (Valeur utilisée pour la première itération)
   case pInTypeRecherche of
      // Recherche de durée
      trDuree :
         ele := GetValue('EMP_NBECHTOT');
      // Recherche de montant de versement
      trVersement :
         begin
            //
            if ((GetValue('EMP_EMPTYPEVERS') = tvConstant) or (GetValue('EMP_AMORTCST') = 0)) and
               (GetValue('EMP_NBECHTOT') <> 0) then
               ele := GetValue('EMP_CAPITAL') / GetValue('EMP_NBECHTOT')
            else
               ele := 10000;
         end;
      // Recherche de capital
      trCapital :
         begin
            if RetourneValeurElement = 0 then
               ele := 100000
            else
               ele := RetourneValeurElement;
         end;
      // Recherche de taux
      trTaux :
         begin
            if RetourneValeurElement = 0 then
               ele := 100 //ele_max
            else
               ele := RetourneValeurElement;
         end;
   end;

   AssigneValeurElement(ele);

   // calcul de la borne
   if pInTypeRecherche = trDuree then
   begin
      if GetValue('EMP_VERSCST') <> 0 then
         borne := GetValue('EMP_CAPITAL') / GetValue('EMP_VERSCST')
      else
         if GetValue('EMP_AMORTCST') <> 0 then
            borne := GetValue('EMP_CAPITAL') / GetValue('EMP_AMORTCST')
         else
            borne := GetValue('EMP_CAPITAL') / 100 * 0.01;
   end
   else
   begin
      if pInTypeRecherche = trTaux then
         borne := 0.000001
      else
         borne := 0.0001;
   end;
   // début de la boucle d'approximation
   if GetValue('EMP_NBECHTOT') > 0 then
      repeat
         inc(iter);
         anc_ele := ele;

         // calcul de l'emprunt
         if pInTypeRecherche = trDuree then
            CalculerEmprunt(pInTypeRecherche, True)
         else
            CalculerEmprunt(pInTypeRecherche, False);

         // Calcul de l'amortissement de l'emprunt
         k := CalculerK;

         // Calcul de la prochaine valeur d'itération par approximation linéaire
         // En fonction de l'amortissement calculé et de l'amortissement souhaité
         case pInTypeRecherche of
            trCapital :
               begin
                  ele := k - fRdTotalReport;
                  delta_ele := anc_ele - ele;
               end;
            trDuree :
               begin
                  delta_ele := (k - fRdTotalReport) - GetValue('EMP_CAPITAL');
                  if (delta_ele > borne) and (delta_ele > 0) then
                     ele := ele - 1;
                  if (- delta_ele > - borne) and (delta_ele < 0) then
                     ele := ele + 1;
                  if ele < cInMaxEch then
                     Putvalue('EMP_NBECHTOT', Int(ele))
                  else
                     Putvalue('EMP_NBECHTOT', cInMaxEch);
               end;
            trTaux :
               begin
                  delta_ele := GetValue('EMP_CAPITAL') - (k - fRdTotalReport);
                  if k <> 0 then
                     pente := delta_ele / k
                  else
                  begin
                     if GetValue('EMP_CAPITAL') <> 0 then
                        pente := delta_ele/GetValue('EMP_CAPITAL')
                     else
                        pente := 0;
                  end;
                  ele := anc_ele * (1 - pente);
                  if ele < 0 then ele := abs(anc_ele) - 10;
{
                  // Autre méthode : dicotomie
                  delta_ele := EMP_CAPITAL - (k - fRdTotalReport);
                  if delta_ele > 0 then
                     ele_max := ele
                  else
                     ele_min := ele;
                  if delta_ele <> 0 then
                  begin
                     ele := (ele_max + ele_min)/2;
                     delta_ele := ele - anc_ele;
                  end;
}
               end;
            trVersement :
               begin
                  delta_ele := k - fRdTotalReport - GetValue('EMP_CAPITAL');
                  if k <> 0 then
                     pente := delta_ele / k
                  else
                  begin
                     if GetValue('EMP_CAPITAL') <> 0 then
                        pente := delta_ele/GetValue('EMP_CAPITAL')
                     else
                        pente := 0;
                  end;
                  ele := anc_ele * (1 - pente);
               end;
         end;

         // Met à jour l'élément concerné par la recherche
         if pInTypeRecherche = trDuree then
            AssigneValeurElement(GetValue('EMP_NBECHTOT') * EMP_NBMOISECH)
         else
            AssigneValeurElement(ele);

         // Erreur si plus de 600 itérations sans résultat probant
         if iter >= cInMaxIter then // si les calculs bouclent on gere l'erreur et on trace
         begin
            if delta_ele = 0 then break;
            // !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
            // Ajouter ici la gestion de l'erreur
            // !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
//            AffMessage('Plus de ' + IntToStr(cInMaxIter) + ' itérations !!!');
            break;
         end;
      Until not ((delta_ele < -borne) or (delta_ele > borne));

   //
   if ele < 0 then ele := 0;

   // Ajout djamel : ajustement du nombre d'échéances si il en manque une ou
   // si il y en a une de trop dans le cas d'une recherche de durée
   if (pInTypeRecherche = trDuree) and (GetValue('EMP_NBECHTOT') > 0) then
   begin
      Ech := Echeances.Detail[GetValue('EMP_NBECHTOT') - 1];
      // Versement ou amortissement ?
      if GetValue('EMP_EMPTYPEVERS') = tvConstant then
      begin
         // Une échéance de trop
         if Ech.GetValue('ECH_VERSEMENT') <= 0 then
            Putvalue('EMP_NBECHTOT',GetValue('EMP_NBECHTOT') - 1);
         // Manque une échéance
         if Ech.GetValue('ECH_VERSEMENT') > GetValue('EMP_VERSCST') then
            Putvalue('EMP_NBECHTOT',GetValue('EMP_NBECHTOT') + 1);
      end
      else
      begin
         // Une échéance de trop
         if Ech.GetValue('ECH_AMORTISSEMENT') <= 0 then
            Putvalue('EMP_NBECHTOT', GetValue('EMP_NBECHTOT') - 1);
         // Manque une échéance
         if Ech.GetValue('ECH_AMORTISSEMENT') > GetValue('EMP_AMORTCST') then
            Putvalue('EMP_NBECHTOT', GetValue('EMP_NBECHTOT') + 1);
      end;

      Result := GetValue('EMP_NBECHTOT') * EMP_NBMOISECH;
   end
   else
      Result := ele;

end;



// ----------------------------------------------------------------------------
// Nom    : Tricher
// Date   : 21/11/2001
// Auteur : D. ZEDIAR
// Objet  : Indique si l'on peut 'tricher' sur la dernière échéance
// ----------------------------------------------------------------------------
// Historique :
// ----------------------------------------------------------------------------
function TObjetEmprunt.Tricher(i : Integer; dette : double) : Boolean;
var
  Ech : Tob;
begin
  Result := False;

  Ech := Echeances.Detail[i];

  if ((dette - Ech.GetValue('ECH_AMORTISSEMENT')) >= 0) and
     ((dette - Ech.GetValue('ECH_AMORTISSEMENT')) < (cRdErreur * Ech.GetValue('ECH_VERSEMENT'))) and
     (Ech.GetValue('ECH_VERSEMENT') <> 0) then
     Result := True;

  if ((Ech.GetValue('ECH_AMORTISSEMENT') - dette) >= 0) and
     ((Ech.GetValue('ECH_AMORTISSEMENT') - dette) < cRdErreur * (Ech.GetValue('ECH_INTERET') + Ech.GetValue('ECH_ASSURANCE'))) and
     (Ech.GetValue('ECH_VERSEMENT') <> 0) then
     Result := True;
end;


// ----------------------------------------------------------------------------
// Nom    : CalcTauxEqui
// Date   : 21/11/2001
// Auteur : D. ZEDIAR
// Objet  : Fonction de calcul du taux bancaire à partir du taux actuariel
// ----------------------------------------------------------------------------
// Historique :
// ----------------------------------------------------------------------------
function TObjetEmprunt.CalcTauxEqui(tx, nb_ech : Double) : Double;
begin
   if nb_ech = 0 then
      Result := 0.0
   else
      Result := 100 * (power(1 + tx/100, 1/nb_ech) - 1);
end;

// ----------------------------------------------------------------------------
// Nom    : CalcTauxPro
// Date   : 21/11/2001
// Auteur : D. ZEDIAR
// Objet  : Fonction de calcul du taux bancaire à partir du taux actuariel
// ----------------------------------------------------------------------------
// Historique :
// ----------------------------------------------------------------------------
Function TObjetEmprunt.CalcTauxPro(tx, nb_ech : Double) : Double;
begin
   if nb_ech = 0 then
      Result := 0
   else
      Result := tx/nb_ech;
end;


// ----------------------------------------------------------------------------
// Nom    : CalcTauxEcheance
// Date   : 21/11/2001
// Auteur : D. ZEDIAR
// Objet  : Calcul un taux pour une échéance donnée (taux de l'emprunt ou taux
//          de l'assurance) en fonction du type de taux de l'emprunt
// ----------------------------------------------------------------------------
// Historique :
// ----------------------------------------------------------------------------
Function TObjetEmprunt.CalcTauxEcheance( pInIndexe : Integer ; pRdTauxAnnuel : Double) : Double;
begin
   // Valeur de base du taux
   if GetValue('EMP_EMPBQEQ') = beEquivalent then
      Result := CalcTauxEqui(pRdTauxAnnuel,EMP_NBECHAN)
   else
      Result := CalcTauxPro(pRdTauxAnnuel,EMP_NBECHAN);

   // calcul des taux
   if (GetValue('EMP_EMPTYPETAUX') <> ttVariable) or
      ((GetValue('EMP_EMPTYPETAUX') = ttVariable) and (GetValue('EMP_AMORTTXINI') = 'X')) then
   begin
      if (pInIndexe > 0) then
      begin
         if GetValue('EMP_EMPTYPETAUX') = ttArythmetique then
         begin
            if GetValue('EMP_EMPBQEQ') = beEquivalent then
               Result := CalcTauxEqui(pRdTauxAnnuel + (GetValue('EMP_RAISON') * pInIndexe),EMP_NBECHAN)
            else
               result := CalcTauxPro(pRdTauxAnnuel + (GetValue('EMP_RAISON') * pInIndexe),EMP_NBECHAN);
         end
         else
         begin
            if GetValue('EMP_EMPTYPETAUX') = ttGeometrique then
            begin
               if GetValue('EMP_EMPBQEQ') = beEquivalent then
                  Result := CalcTauxEqui(pRdTauxAnnuel * power(GetValue('EMP_RAISON'){/1000},pInIndexe),EMP_NBECHAN)
               else
                  Result := CalcTauxPro(pRdTauxAnnuel * power(GetValue('EMP_RAISON'){/1000},pInIndexe),EMP_NBECHAN);
            end;
         end;
      end;

   end;
end;

// ----------------------------------------------------------------------------
// Nom    : CalculerEmprunt
// Date   : 21/11/2001
// Auteur : D. ZEDIAR
// Objet  : Calcul des échéances de l'emprunt
// ----------------------------------------------------------------------------
// Historique :
// ----------------------------------------------------------------------------
Procedure TObjetEmprunt.CalculerEmprunt(pInTypeRecherche : TTypeRecherche; pBoArrDerEch : Boolean);
var
   Ech            : Tob;
   i              : Word;
   solde_report   : double;
   solde          : double;
   dette          : double;
   report2        : double;
   report_s       : double;
   amort          : double;
   interet        : double;
   montant        : double;
   rompus_int     : double;
   rompus_ass     : double;
   rompus_amo     : double;
   report_int     : double;
   report_ass     : double;
   lRdAmorReelEch : double;
   fRdReportInt : double;
   fRdReportAss   : double;

   lRdBaseAssurance : Double;

   lRdTauxGlobal    : Double;
   lRdTauxHorsAssur : Double;

   lRdUnitFrac : Double;

   // Répercute les reports d'int est d'ass sur l'amortissement
   Procedure RepercuterReport(var pRdReport : double; pInDiffere : Integer);
   begin
      if (pRdReport > 0) and (i >= pInDiffere) then
      begin
         if pRdReport <= lRdAmorReelEch then
         begin
            lRdAmorReelEch := lRdAmorReelEch - pRdReport;
            pRdReport := 0;
         end
         else
         begin
            pRdReport := pRdReport - lRdAmorReelEch;
            lRdAmorReelEch := 0;
         end;
      end
      else
   end;


   // Applique le total des arrondis si ce total deviens significatif (> 10^cInNbDecMnt)
   procedure AppliquerRompus(var pRdRompus : Double);
   begin
      if (pRdRompus >= lRdUnitFrac) then
      begin
         Ech.Putvalue('ECH_INTERET', Ech.GetValue('ECH_INTERET') + lRdUnitFrac);
         pRdRompus := pRdRompus - lRdUnitFrac;
      end
      else
      begin
         if (- pRdRompus >= lRdUnitFrac) then
         begin
            Ech.Putvalue('ECH_INTERET', Ech.GetValue('ECH_INTERET') - lRdUnitFrac);
            pRdRompus := pRdRompus + lRdUnitFrac;
         end
         else
         begin
            // Si on est sur la dernière échéance, on répercute rompus sur les intérets
            // Sinon, on garde rompus
            if i = GetValue('EMP_NBECHTOT') - 1 then
            begin
               Ech.Putvalue('ECH_INTERET', Ech.GetValue('ECH_INTERET') + pRdRompus);
               Ech.Putvalue('ECH_INTERET', Arrondi(Ech.Getvalue('ECH_INTERET'), cInNbDecMnt));
               pRdRompus := 0;
            end;
         end;
      end;
   end;

begin
   report_int     := fRdReportDeBase_int;
   report_ass     := fRdReportDeBase_ass;
   report2        := 0;
   amort          := 0;
   interet        := 0;
   montant        := 0;
   rompus_int     := 0;
   rompus_ass     := 0;
   rompus_amo     := 0;
   fRdTotalReport := 0;
   fRdReportInt   := 0;
   fRdReportAss   := 0;
   lRdUnitFrac := power(10,cInNbDecMnt);

   solde := GetValue('EMP_CAPITAL');
   dette := GetValue('EMP_CAPITAL') + fRdReportDeBase_int + fRdReportDeBase_ass;
   solde_report := solde + fRdReportDeBase_int + fRdReportDeBase_ass;

   Putvalue('EMP_TOTINT',   0);
   Putvalue('EMP_TOTASS',   0);
   Putvalue('EMP_TOTAMORT', 0);
   Putvalue('EMP_TOTVERS',  0);

   if GetValue('EMP_NBECHTOT') > Echeances.Detail.count then
      CreerEcheances(GetValue('EMP_NBECHTOT'));

   // Boucle sur les échéances
   for i := 0 to GetValue('EMP_NBECHTOT') - 1 do
   begin
      Ech := Echeances.Detail[i];

      if not fBoTauxReel then
//         Echeances.Blank(i);
         Ech.InitValeurs;

      // Versement ou amortissement constant
      if not fBoTauxReel then
      begin
         if GetValue('EMP_EMPTYPEVERS') = tvConstant then
         begin
            Ech.Putvalue('ECH_VERSEMENT', GetValue('EMP_VERSCST'));
            Ech.Putvalue('ECH_AMORTISSEMENT', 0);
         end
         else
         begin
            Ech.Putvalue('ECH_VERSEMENT', 0);
            Ech.Putvalue('ECH_AMORTISSEMENT', GetValue('EMP_AMORTCST'));
         end;
      end;

      // calcul des taux
      if fBoTauxReel then
         lRdTauxHorsAssur := RechercherTaux(DateEcheance(i))
      else
         lRdTauxHorsAssur := GetValue('EMP_TAUXAN');

      if (GetValue('EMP_EMPTYPEASSUR') = taTaux) and
         ((GetValue('EMP_EMPBQEQ') = beEquivalent) or (GetValue('EMP_EMPTYPETAUX') = ttGeometrique)) then
      begin
         // Si le calcul du taux fait intervenir des exponetiations, on travail sur un
         // taux global et des rapports
         lRdTauxGlobal := CalcTauxEcheance(i,lRdTauxHorsAssur + GetValue('EMP_VALASSUR'));
         Ech.PutValue( 'ECH_TAUXINT', lRdTauxGlobal * GetValue('EMP_TAUXAN') / (GetValue('EMP_TAUXAN') + GetValue('EMP_VALASSUR')) );
         Ech.PutValue( 'ECH_TAUXASS', lRdTauxGlobal * GetValue('EMP_VALASSUR') / (GetValue('EMP_TAUXAN') + GetValue('EMP_VALASSUR')) );
      end
      else
      begin
         Ech.PutValue( 'ECH_TAUXINT', CalcTauxEcheance(i,lRdTauxHorsAssur) );
         if GetValue('EMP_EMPTYPEASSUR') = taTaux then
            Ech.Putvalue('ECH_TAUXASS', CalcTauxEcheance(i,GetValue('EMP_VALASSUR')))
         else
            Ech.Putvalue('ECH_TAUXASS', 0);
      end;

      //
      // calcul de l'assurance
      //

      // A-t-on un différé d'assurance
      if (i >= GetValue('EMP_DIFASS')) and (GetValue('EMP_VALASSUR') <> 0) then
      begin
         // Pas de differé d'assurance et assurance <> 0
         // Base dette ou solde ?
         if GetValue('EMP_EMPTYPEASSUR') = taMontant then
            Ech.Putvalue('ECH_ASSURANCE', GetValue('EMP_VALASSUR'))
         else
         begin
            if GetValue('EMP_EMPTYPEASSUR') = taTaux then
            begin
               if GetValue('EMP_EMPBASEASSUR') = tbaSolde then
                  lRdBaseAssurance := solde
               else
                  lRdBaseAssurance := dette;
               Ech.Putvalue('ECH_ASSURANCE', Ech.GetValue('ECH_TAUXASS') * lRdBaseAssurance/100);
            end
         end;
      end
      else
      begin
        // On a un différé d'assurance ou pas d'assurance du tout
        Ech.Putvalue('ECH_ASSURANCE', 0);
        // Report d'assurance
        if GetValue('EMP_EMPTYPEASSUR') = taMontant then
           report_ass := report_ass + GetValue('EMP_VALASSUR')
        else
           if GetValue('EMP_EMPTYPEASSUR') = taTaux then
              report_ass := report_ass + Ech.GetValue('ECH_TAUXASS') * solde_report / 100;
      end;

      //
      // calcul des interets
      //

      // Pas d'intérets si remboursement par anticipation
//      if DateEcheance(i) < EMP_DATECONTRAT.Asdate then
//         ECH_INTERET.Value := 0
//      else
      begin
         // Différé d'interets ?
         if i < GetValue('EMP_DIFINT') then
         begin
            // Report des intérêts
            report_int := report_int + (Ech.GetValue('ECH_TAUXINT') * solde_report/100);
         end
         else
         begin
            // Pas de différé d'intéret

            //
            if (Ech.GetValue('ECH_VERSEMENT') = 0) and
               (GetValue('EMP_EMPTYPEVERS') = tvVariable) and
               (GetValue('EMP_AMORTCST') = 0) then
            begin
               // Versement variable (amortissement constant)
               // Différé de remboursement ?
               if i < GetValue('EMP_DIFREMB') then
                  Ech.Putvalue('ECH_INTERET', Ech.GetValue('ECH_TAUXINT') * dette/100)
               else
               begin
                  report_s                := Ech.GetValue('ECH_TAUXINT') * dette/100;
                  report_int              := report_int + report_s;
                  report2                 := report2 + report_s;
                  Ech.Putvalue('ECH_INTERET', 0);
               end
            end
            else
            begin
               // Versement constant (amortissement variable)
               Ech.Putvalue('ECH_INTERET', Ech.GetValue('ECH_TAUXINT') * dette/100);

               // Interets > versement ?
               if (Ech.GetValue('ECH_INTERET') + Ech.GetValue('ECH_ASSURANCE') > Ech.GetValue('ECH_VERSEMENT')) and
                  (Ech.GetValue('EMP_EMPTYPEVERS') = tvVariable) and
                  (GetValue('EMP_AMORTCST') = 0) then
               begin
                  report_s := (Ech.GetValue('ECH_INTERET')  + Ech.GetValue('ECH_ASSURANCE') - Ech.GetValue('ECH_VERSEMENT'));
                  report_int := report_int + report_s;
                  report2  := report2 + report_s;

                  if i < GetValue('EMP_DIFREMB') then
                     Ech.Putvalue('ECH_INTERET', Ech.GetValue('ECH_TAUXINT') * dette/100)
                  else
                     Ech.Putvalue('ECH_INTERET', Ech.GetValue('ECH_VERSEMENT'));
               end;
            end;
         end;
      end;

      //
      // On arrondi à 2 décimales intéret et assurance, le delta étant stocké dans rompus_int
      // rompus_int (ou rompus_int) est le cumul des parties fractionnaires des intérets (ou des assurances)
      //
      if i >= GetValue('EMP_DIFINT') then
      begin
         rompus_int := rompus_int + Ech.GetValue('ECH_INTERET') - Arrondi(Ech.GetValue('ECH_INTERET'), cInNbDecMnt);
         Ech.Putvalue('ECH_INTERET', Arrondi(Ech.GetValue('ECH_INTERET'), cInNbDecMnt) );
         AppliquerRompus(rompus_int);
      end;

      if i >= GetValue('EMP_DIFASS') then
      begin
         rompus_ass := rompus_ass + Ech.GetValue('ECH_ASSURANCE') - Arrondi(Ech.GetValue('ECH_ASSURANCE'), cInNbDecMnt);
         Ech.Putvalue('ECH_ASSURANCE', Arrondi(Ech.GetValue('ECH_ASSURANCE'), cInNbDecMnt) );  //
         AppliquerRompus(rompus_ass);
      end;

      //
      // calcul des amortissements et des versements
      //

      // Y a-t-il différé de remboursement ?
      if i >= GetValue('EMP_DIFREMB') then
      begin
         // Pas de différé de remboursement
         if Ech.GetValue('ECH_AMORTISSEMENT') <> 0 then
         begin
            if (Ech.GetValue('ECH_AMORTISSEMENT') > dette) and (pInTypeRecherche = trVersement) then
               Ech.Putvalue('ECH_AMORTISSEMENT', dette);
            Ech.Putvalue( 'ECH_VERSEMENT', Ech.GetValue('ECH_AMORTISSEMENT') + Ech.GetValue('ECH_INTERET') + Ech.GetValue('ECH_ASSURANCE') );
         end
         else
         begin
            Ech.Putvalue( 'ECH_AMORTISSEMENT', Ech.GetValue('ECH_VERSEMENT') - Ech.GetValue('ECH_INTERET') - Ech.GetValue('ECH_ASSURANCE') );
            if (Ech.GetValue('ECH_AMORTISSEMENT') > dette) and
               (GetValue('EMP_EMPTYPEVERS') = tvVariable) and
               (pInTypeRecherche = trVersement) then
            begin
               Ech.Putvalue( 'ECH_AMORTISSEMENT', dette );
               Ech.Putvalue( 'ECH_VERSEMENT', Ech.GetValue('ECH_AMORTISSEMENT') + Ech.GetValue('ECH_INTERET') + Ech.GetValue('ECH_ASSURANCE') );
            end;
         end;
      end
      else
      begin
         // Il y a un différé de remboursement
         Ech.Putvalue( 'ECH_AMORTISSEMENT', 0 );
         Ech.Putvalue( 'ECH_VERSEMENT', Ech.GetValue('ECH_INTERET') + Ech.GetValue('ECH_ASSURANCE') );
      end;


      //
      // On arrondi à 2 décimales l'amortissement
      //
      if GetValue('EMP_EMPTYPEVERS') = tvConstant then
      begin
         rompus_amo := rompus_amo + (Ech.GetValue('ECH_AMORTISSEMENT') - Arrondi(Ech.GetValue('ECH_AMORTISSEMENT'), cInNbDecMnt));
         Ech.Putvalue( 'ECH_AMORTISSEMENT', Arrondi(Ech.GetValue('ECH_AMORTISSEMENT'), cInNbDecMnt) );

         if rompus_amo >= 1 then
         begin
            Ech.Putvalue( 'ECH_AMORTISSEMENT', Ech.GetValue('ECH_AMORTISSEMENT') + 1 );
            rompus_amo := rompus_amo - 1.0;
         end
         else
         begin
            if - rompus_amo >= 1 then
            begin
               Ech.Putvalue( 'ECH_AMORTISSEMENT', Ech.GetValue('ECH_AMORTISSEMENT') - 1 );
               rompus_amo := rompus_amo + 1;
            end
            else
            begin
               if i = GetValue('EMP_NBECHTOT') then
               begin
                  Ech.Putvalue( 'ECH_AMORTISSEMENT', Ech.GetValue('ECH_AMORTISSEMENT') + rompus_amo );
                  rompus_amo := 0;
                  Ech.Putvalue( 'ECH_AMORTISSEMENT', Arrondi(Ech.GetValue('ECH_AMORTISSEMENT'), cInNbDecMnt) );
               end;
            end;
         end;
      end;

      //
      // tricherie sur la dernière échéance pour que le plan tombe juste à 0
      //
      if (pBoArrDerEch) and (i = GetValue('EMP_NBECHTOT') - 1) then
      begin
         if pInTypeRecherche <> trDuree then
         begin
            if Tricher(i,dette) = true then
            begin
               Ech.Putvalue( 'ECH_INTERET', Ech.GetValue('ECH_INTERET') - (dette - Ech.GetValue('ECH_AMORTISSEMENT')) );
               Ech.Putvalue( 'ECH_AMORTISSEMENT', dette );
               Ech.Putvalue( 'ECH_VERSEMENT', Ech.GetValue('ECH_AMORTISSEMENT') + Ech.GetValue('ECH_INTERET') + Ech.GetValue('ECH_ASSURANCE') );
            end
            else
            begin
               if (Ech.GetValue('ECH_AMORTISSEMENT') - dette) >= 0 then
                  Ech.Putvalue('ECH_AMORTISSEMENT', dette);
               Ech.Putvalue('ECH_VERSEMENT', Ech.GetValue('ECH_AMORTISSEMENT') + Ech.GetValue('ECH_INTERET') + Ech.GetValue('ECH_ASSURANCE') );
            end;
         end
         else
         begin
            if ( (GetValue('EMP_VERSCST') <> 0) and
                 (dette < cRdTxMarge * (GetValue('EMP_VERSCST') - Ech.GetValue('ECH_INTERET') - Ech.GetValue('ECH_ASSURANCE'))) ) or
               ( (GetValue('EMP_AMORTCST') <> 0) and
                 (dette < cRdTxMarge * GetValue('EMP_AMORTCST'))) then
            begin
               Ech.Putvalue('ECH_VERSEMENT', Ech.GetValue('ECH_VERSEMENT') - (dette - Ech.GetValue('ECH_AMORTISSEMENT')) );
               Ech.Putvalue('ECH_AMORTISSEMENT', dette );
               Ech.Putvalue('ECH_VERSEMENT', Ech.GetValue('ECH_AMORTISSEMENT') + Ech.GetValue('ECH_INTERET') + Ech.GetValue('ECH_ASSURANCE') )
            end;
         end;

         Ech.Putvalue('ECH_VERSEMENT', Ech.GetValue('ECH_VERSEMENT') + rompus_int + rompus_ass );
      end;

      // modif du 5/10/92 CRI
      if (Ech.GetValue('ECH_INTERET') + Ech.GetValue('ECH_AMORTISSEMENT') + Ech.GetValue('ECH_ASSURANCE') <> Ech.GetValue('ECH_VERSEMENT')) and
         (pInTypeRecherche = trCapital) and
         (pBoArrDerEch) then
      begin
         if Ech.GetValue('ECH_AMORTISSEMENT') <> 0 then
         begin
            Ech.Putvalue('ECH_AMORTISSEMENT', Ech.GetValue('ECH_VERSEMENT') - Ech.GetValue('ECH_INTERET') - Ech.GetValue('ECH_ASSURANCE') );
            Ech.Putvalue('ECH_AMORTISSEMENT', Arrondi(Ech.GetValue('ECH_AMORTISSEMENT'), cInNbDecMnt) );
         end
         else
         begin
            Ech.Putvalue('ECH_INTERET', Ech.GetValue('ECH_VERSEMENT') );
            Ech.Putvalue('ECH_INTERET', Arrondi(Ech.GetValue('ECH_INTERET'), cInNbDecMnt) );
         end;
      end;

      // calcul du solde et du report
      if report_int < 0 then report_int := 0;
      if report_ass < 0 then report_ass := 0;

      // Met le total du report dans fRdReportInt et fRdReportAss
      if report_int > fRdReportInt then
         fRdReportInt := report_int;
      if report_ass > fRdReportAss then
         fRdReportAss := report_ass;

      // Le report est répercuté sur le solde dans la mesure du possible
      lRdAmorReelEch := Ech.GetValue('ECH_AMORTISSEMENT');
      RepercuterReport(report_int,GetValue('EMP_DIFINT'));
      RepercuterReport(report_ass,GetValue('EMP_DIFASS'));
      solde := solde - lRdAmorReelEch;

      // Actualisation annuelle
      if (((i + 1) mod EMP_NBECHAN) = 0) or (GetValue('EMP_ACTUANNU') = '-') then
//         solde_report := EMP_CAPITAL.AsFloat + report;
         solde_report := solde + report_int + report_ass;

      // calcul des totaux
      amort   := amort   + Ech.GetValue('ECH_AMORTISSEMENT');
      interet := interet + Ech.GetValue('ECH_INTERET') + Ech.GetValue('ECH_ASSURANCE');
      montant := montant + Ech.GetValue('ECH_AMORTISSEMENT') + Ech.GetValue('ECH_INTERET') + Ech.GetValue('ECH_ASSURANCE');

      if solde < 0 then solde := 0;

      dette := solde + report_int + report_ass;

      // maj des échéances
      Ech.Putvalue('ECH_REPORTINT', report_int);
      Ech.Putvalue('ECH_REPORTASS', report_ass);
      Ech.Putvalue('ECH_SOLDE', solde);
      Ech.Putvalue('ECH_DETTE', dette);

      Ech.Putvalue('ECH_CUMULA', amort);
      Ech.Putvalue('ECH_CUMULI', interet);
      Ech.Putvalue('ECH_CUMULV', montant);

      // mise à jour des totaux généraux
      Putvalue('EMP_TOTINT', GetValue('EMP_TOTINT') + Ech.GetValue('ECH_INTERET') );
      Putvalue('EMP_TOTASS', GetValue('EMP_TOTASS') + Ech.GetValue('ECH_ASSURANCE') );
      Putvalue('EMP_TOTAMORT', GetValue('EMP_TOTAMORT') + Ech.GetValue('ECH_AMORTISSEMENT') );
      Putvalue('EMP_TOTVERS', GetValue('EMP_TOTVERS') + Ech.GetValue('ECH_ASSURANCE') +
                              Ech.GetValue('ECH_INTERET') + Ech.GetValue('ECH_AMORTISSEMENT') );

   end;

   // Arrondi les amortissements ou les versements
   if pBoArrDerEch then
   begin
      Putvalue('EMP_TOTAMORT', 0);
      Putvalue('EMP_TOTVERS', 0);
      For i := 0 to GetValue('EMP_NBECHTOT') - 1 do
      begin
         Ech := Echeances.Detail[i];
         if GetValue('EMP_EMPTYPEVERS') = tvConstant then
            Ech.Putvalue('ECH_VERSEMENT', Ech.GetValue('ECH_AMORTISSEMENT') + Ech.GetValue('ECH_ASSURANCE') + Ech.GetValue('ECH_INTERET') )
         else
            Ech.Putvalue('ECH_AMORTISSEMENT', Ech.GetValue('ECH_VERSEMENT') - Ech.GetValue('ECH_ASSURANCE') - Ech.GetValue('ECH_INTERET') );

         Putvalue('EMP_TOTVERS', GetValue('EMP_TOTVERS') + Ech.GetValue('ECH_VERSEMENT') );
         Putvalue('EMP_TOTAMORT', GetValue('EMP_TOTAMORT') + Ech.GetValue('ECH_AMORTISSEMENT') );
      end;
   end;


   fRdTotalReport := fRdReportInt + fRdReportAss;
   if fRdTotalReport < 0 then fRdTotalReport := 0;

end;


// ----------------------------------------------------------------------------
// Nom    : CalculerK
// Date   : 21/11/2001
// Auteur : D. ZEDIAR
// Objet  : Fonction qui calcule le montant du capital emprunté
//          Cette fonction recalcul le montant du capital amorti pour
//          chaque échéance et fait la somme sachant que pour calculer
//          l'amortissement de l'échéance 1 on utilise les infos de
//          l'échéance n
// ----------------------------------------------------------------------------
// Historique :
// ----------------------------------------------------------------------------
Function TObjetEmprunt.CalculerK : Double;
var
   Ech       : Tob;
   i         : Integer;
   somme     : Double;
   puissance : double;
   nombre    : double;
   moy_tx    : Double;
   som_tx    : Double;
   lRdValeur : Double;
   lRdDifInt : Integer;
   lRdDifAss : Integer;
begin
   som_tx    := 0;
   somme     := 0;

   lRdDifInt := GetValue('EMP_DIFINT') - GetValue('EMP_DIFREMB');
   if lRdDifInt < 0 then lRdDifInt := 0;

   lRdDifAss := GetValue('EMP_DIFASS') - GetValue('EMP_DIFREMB');
   if lRdDifAss < 0 then lRdDifAss := 0;

   for i := 1 + GetValue('EMP_DIFREMB') to GetValue('EMP_NBECHTOT') do
   begin
      Ech := Echeances.Detail[i - 1];
      puissance := GetValue('EMP_DIFREMB') - i;
      if GetValue('EMP_EMPTYPETAUX') = ttConstant then
      begin
         // si taux constant
         // Teste les différés
         nombre := 1;
         if GetValue('EMP_NBECHTOT') - i >= lRdDifInt then
         begin
               nombre := nombre + Ech.GetValue('ECH_TAUXINT') / 100;
         end;
         if (GetValue('EMP_EMPTYPEASSUR') = taTaux) and (GetValue('EMP_NBECHTOT') - i >= lRdDifAss) then
            nombre := nombre + Ech.GetValue('ECH_TAUXASS') / 100;
      end
      else
      begin
         // si taux variable/géométrique/arithmétique
         lRdValeur := 0;
         if GetValue('EMP_NBECHTOT') - i >= lRdDifInt then
            lRdValeur := Ech.GetValue('ECH_TAUXINT');
         if (GetValue('EMP_EMPTYPEASSUR') = taTaux) and (GetValue('EMP_NBECHTOT') - i >= lRdDifAss) then
            lRdValeur := lRdValeur + Ech.GetValue('ECH_TAUXASS');
         som_tx := som_tx + lRdValeur ;
         moy_tx := som_tx/(i - GetValue('EMP_DIFREMB'));
         nombre := (1 + moy_tx/100);
      end;
      lRdValeur := Ech.GetValue('ECH_VERSEMENT');

      // Prise en compte du montant de l'assurance
      if GetValue('EMP_EMPTYPEASSUR') = taMontant then
         lRdValeur := lRdValeur - GetValue('EMP_VALASSUR');

      // Valeur de l'amortissement pour cette échéance
      lRdValeur := lRdValeur * power(nombre, puissance);

      // Sommage des ammortissements
      Somme := Somme + lRdValeur;
   end;

   Result := somme;
end;


// ----------------------------------------------------------------------------
// Nom    : CalculerNbEch
// Date   : 21/11/2001
// Auteur : D. ZEDIAR
// Objet  : Détermine le nombre d'échéances total nécéssaire au calcul
//          de l'emprunt
// ----------------------------------------------------------------------------
// Historique :
// ----------------------------------------------------------------------------
function TObjetEmprunt.CalculerNbEch(pInTypeRecherche : TTypeRecherche) : Integer;
var
   k, decimale : Double;
begin

   k := GetValue('EMP_CAPITAL') + fRdTotalReport;

   if pInTypeRecherche <> trDuree then
   begin
      {if ((EMP_EMPTYPEVERS = tvVariable) and (EMP_AMORTCST = 0)) or
         (EMP_EMPTYPETAUX = ttVariable) then
      begin
         // !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
         // Gérer ce cas par la suite
         // !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
         //for (EMP_NBECHTOT = 0, cur_ech = emp->deb_ech; cur_ech ; EMP_NBECHTOT++,cur_ech = mem_next(cur_ech));
         //tab_ech = (struct s_ech *) TMmalloc(sizeof (struct s_ech) * EMP_NBECHTOT);
      end
      else}
         Putvalue('EMP_NBECHTOT', Int(GetValue('EMP_DUREE') / EMP_NBMOISECH));
   end
   else
   begin
      // Gestion élément recherché = durée
      Putvalue('EMP_NBECHTOT', cInMaxEch + 1);

      if (GetValue('EMP_AMORTCST') <> 0) and (GetValue('EMP_EMPTYPEVERS') <> tvConstant) then
      begin
         // Cas de l'amortissement constant
         Putvalue('EMP_NBECHTOT', Int(Arrondir(k / GetValue('EMP_AMORTCST'),10,0) + GetValue('EMP_DIFREMB')) );

         decimale := k / GetValue('EMP_AMORTCST') + GetValue('EMP_DIFREMB') - GetValue('EMP_NBECHTOT');

         if decimale < cRdBorneInf then
            Putvalue('EMP_NBECHTOT', GetValue('EMP_NBECHTOT') - 1);

         if decimale > cRdBorneSup then
            Putvalue('EMP_NBECHTOT', GetValue('EMP_NBECHTOT') + 1 )
         else
         begin
            if decimale < cRdBorneInf then
               Putvalue('EMP_NBECHTOT', GetValue('EMP_NBECHTOT') + 1 );
         end;
      end
      else
      begin
         // Cas du versement constant
         if (GetValue('EMP_VERSCST') <> 0) and (GetValue('EMP_EMPTYPEVERS') = tvConstant) then
         begin
            Putvalue('EMP_NBECHTOT', Int(Arrondir(k / GetValue('EMP_VERSCST'),10,0) + GetValue('EMP_DIFREMB')) );

            decimale := k / GetValue('EMP_VERSCST') + GetValue('EMP_DIFREMB') - GetValue('EMP_NBECHTOT');

            if decimale < cRdBorneInf then
               Putvalue('EMP_NBECHTOT', GetValue('EMP_NBECHTOT') - 1 );

            if decimale > cRdBorneSup then
               Putvalue('EMP_NBECHTOT', GetValue('EMP_NBECHTOT') + 1 )
            else
            begin
               if decimale < cRdBorneInf then
                  Putvalue('EMP_NBECHTOT', GetValue('EMP_NBECHTOT') + 1 );
            end;
         end
         else
            Putvalue('EMP_NBECHTOT', cInMaxEch);
      end;

      if GetValue('EMP_NBECHTOT') <= cInMaxEch then
         Putvalue('EMP_DUREE', GetValue('EMP_NBECHTOT') * EMP_NBMOISECH )
      else
      begin
         Putvalue('EMP_NBECHTOT', Int(GetValue('EMP_DUREE') / EMP_NBMOISECH) );
//         EMP_NBECHTOT.Value = 0;
//         emp->duree = 0.0;
//         ret = ERR;
      end;

      Putvalue('EMP_NBECHTOT', GetValue('EMP_NBECHTOT') + 1);
   end;

   Result := GetValue('EMP_NBECHTOT');

end;



// ----------------------------------------------------------------------------
// Nom    : EmpruntValide
// Date   : 21/11/2001
// Auteur : D. ZEDIAR
// Objet  : Cette fonction tèste la validité de l'emprunt
// ----------------------------------------------------------------------------
// Historique :
// ----------------------------------------------------------------------------
function TObjetEmprunt.EmpruntValide(pInTypeRecherche : TTypeRecherche) : TTestEmpruntResult;
var
//   i         : Integer;
   duree     : Double;
//   versement : Double;
begin

//   versement := 0;

   duree := GetValue('EMP_NBECHTOT') - GetValue('EMP_DIFREMB');

   Result := terOk;

   {
   if pInAction <> tActModif then
   begin
      if (EMP_EMPTYPEVERS.AsString = tvVariable) and (EMP_AMORTCST.AsFloat = 0) then
      begin
         if versement < EMP_CAPITAL.AsFloat then
            result := terInsuffVers
      end
      else
      begin
         if (TauxTotal <> 0) and (EMP_EMPTYPETAUX.AsString = ttConstant) then
         begin
            if ( (EMP_VERSCST.AsFloat * duree) > (EMP_CAPITAL.AsFloat * (1 + TauxTotal/100 * duree)) ) and
               ( pInTypeRecherche <> trDuree ) and
               ( pInTypeRecherche <> trVersement ) and
               ( EMP_VERSCST.AsFloat <> 0) and
               ( EMP_CAPITAL.AsFloat <> 0) then
               result := terTropVers;
         end;

         if ( (EMP_VERSCST.AsFloat * duree) < EMP_CAPITAL.AsFloat ) and
            ( pInTypeRecherche <> trDuree ) and
            ( pInTypeRecherche <> trVersement ) and
            ( EMP_VERSCST.AsFloat <> 0 ) and
            ( EMP_CAPITAL.AsFloat <> 0 ) then
            result := terInsuffVers;
      end;
   end;
   }

   // !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
   // revoir les arrondis ici car ils doivent prendre en compte le facteur 1000
   // !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
   if (Arrondir(GetValue('EMP_AMORTCST') * duree,0,2) < Arrondir(GetValue('EMP_CAPITAL') + fRdTotalReport,10,2)) and
      (GetValue('EMP_AMORTCST') <> 0) and
      (pInTypeRecherche <> trDuree) and
      (GetValue('EMP_CAPITAL') <> 0) then
      Result := terInsuffAmort;

   if (Arrondir(GetValue('EMP_AMORTCST') * duree,10,2) > Arrondir(GetValue('EMP_CAPITAL') + fRdTotalReport,0,2)) and
      (GetValue('EMP_AMORTCST') <> 0) and
      (pInTypeRecherche <> trDuree) and
      (GetValue('EMP_CAPITAL') <> 0) then
      Result := terTropAmort;

end;


// ----------------------------------------------------------------------------
// Nom    : Simulation
// Date   : 21/11/2001
// Auteur : D. ZEDIAR
// Objet  : Fonction de simulation de l'emprunt
// ----------------------------------------------------------------------------
// Historique :
// ----------------------------------------------------------------------------
Procedure TObjetEmprunt.Simulation(pInTypeRecherche : TTypeRecherche);
begin

    CreerEcheances(0);

    if GetValue('EMP_EMPTYPEVERS') = tvConstant then
       Putvalue('EMP_AMORTCST', 0)
    else
       PutValue('EMP_VERSCST', 0);

    if GetValue('EMP_EMPTYPETAUX') = ttVariable then
    begin
       // Taux variable
       if GetValue('EMP_AMORTTXINI') = 'X' then
          // Amortissement calculé sur le taux initial
          Simulation3(pInTypeRecherche)
       else
          // Amortissement non calculé sur le taux initial
          Simulation2(pInTypeRecherche);
    end
    else
       // Taux non variable (constant, arythmétique, géométrique)
       Simulation1(pInTypeRecherche);
end;


// ----------------------------------------------------------------------------
// Nom    : Simulation1
// Date   : 21/11/2001
// Auteur : D. ZEDIAR
// Objet  : Fonction de simulation de l'emprunt lorsque le taux n'est pas
//          variable
// ----------------------------------------------------------------------------
// Historique :
// ----------------------------------------------------------------------------
Procedure TObjetEmprunt.Simulation1(pInTypeRecherche : TTypeRecherche);
var
   lInCpt     : Integer;
   lOmEmprunt : TObjetEmprunt;
   lObEch     : Tob;
begin

   // Emprunt de travail
   lOmEmprunt := TObjetEmprunt.CreerEmprunt;

   try
      // Création des échéances
      lOmEmprunt.Assigner(Self);
      lOmEmprunt.CalculerNbEch(pInTypeRecherche);    // Calcul EMP_NBECHTOT
      lOmEmprunt.CreerEcheances(0);
      lOmEmprunt.CreerEcheances(cInMaxEch);             // allocation de la mémoire 'une fois pour toute'

      // Calcul du versement constant de base
      if lOmEmprunt.GetValue('EMP_NBECHTOT') <> 0 then
      begin
         if ((lOmEmprunt.GetValue('EMP_VERSCST') * lOmEmprunt.GetValue('EMP_NBECHTOT')) < lOmEmprunt.GetValue('EMP_CAPITAL')) and
            (pInTypeRecherche = trVersement) and
            (lOmEmprunt.GetValue('EMP_EMPTYPEVERS') = tvConstant) then
            lOmEmprunt.Putvalue('EMP_VERSCST', lOmEmprunt.GetValue('EMP_CAPITAL')/lOmEmprunt.GetValue('EMP_NBECHTOT') );
      end
      else
         lOmEmprunt.Putvalue('EMP_VERSCST', 0);

      // Tèste de validité de l'emprunt
      case lOmEmprunt.EmpruntValide(pInTypeRecherche) of
         terInsuffVers  : AffMessage('Pas assez de versements');
         terTropVers    : AffMessage('Trop de versements');
         terInsuffAmort : AffMessage('Pas assez d''amortissements');
         terTropAmort   : AffMessage('Trop d''amortissements');
      end;

      // Calcul itératif de l'élément recherché sur l'emprunt de travail
      case pInTypeRecherche of
         trTaux : // simulation du taux
            begin
               lOmEmprunt.Putvalue('EMP_TAUXAN', Arrondi(lOmEmprunt.Genius(pInTypeRecherche),cInNbDecTx) );
//               lOmEmprunt.EMP_TAUXAN := lOmEmprunt.Genius(pInTypeRecherche);
            end;
         trCapital : // simulation du capital
            begin
               lOmEmprunt.PutValue('EMP_CAPITAL', Arrondi(lOmEmprunt.Genius(pInTypeRecherche),cInNbDecMnt) );
            end;
         trDuree : // simulation de la duree
            begin
               lOmEmprunt.PutValue('EMP_DUREE', Int(lOmEmprunt.Genius(pInTypeRecherche)) );
               lOmEmprunt.PutValue('EMP_DUREE', lOmEmprunt.GetValue('EMP_NBECHTOT') * lOmEmprunt.EMP_NBMOISECH );
            end;
         trVersement : // simulation des versements
            begin
               // Lance le calcul seleument si on est en versement constant
               // (l'amortissement constant étant saisi)
               if lOmEmprunt.GetValue('EMP_EMPTYPEVERS') = tvConstant then
                  lOmEmprunt.Putvalue('EMP_VERSCST', lOmEmprunt.Genius(pInTypeRecherche));
            end;
      end;

      // Maj et calcul de l'emprunt courant avec les infos de l'emprunt de travail
      Self.Assigner(lOmEmprunt);
      Self.CreerEcheances(GetValue('EMP_NBECHTOT'));
      CalculerEmprunt(pInTypeRecherche,True);

      // calcul des dates des échéances
      for lInCpt := 0 to Self.Echeances.Detail.Count - 1 do
      begin
         lObEch := Echeances.Detail[lInCpt];
         lObEch.PutValue('ECH_DATE', DateEcheance(lInCpt) );
         lObEch.PutValue('ECH_CODEEMPRUNT', GetValue('EMP_CODEEMPRUNT') );
         lObEch.PutValue('ECH_DATEEXERCICE', ProchaineDateExercice(lObEch.GetValue('ECH_DATE')) );
      end;

   finally
      lOmEmprunt.Free;
   end;

end;


// Création de n échéances vides
Procedure TObjetEmprunt.CreerEcheances(pNbEcheances : Integer);
var
   lInCpt : Integer;
begin
   if pNbEcheances < Echeances.Detail.Count then
   begin
      // Suppression d'enregistrements
      While Echeances.Detail.Count > pNbEcheances do
         Echeances.Detail[pNbEcheances].Free;
   end
   else
   begin
      // Ajout d'enregistrements
      For lInCpt := Echeances.Detail.Count to pNbEcheances - 1 do
          TOB.Create('FECHEANCE',Echeances,-1);
   end;
end;


Procedure TObjetEmprunt.CreerTaux(pNbTaux : Integer);
var
   lInCpt : Integer;
begin
   if pNbTaux < TauxVariables.Detail.Count then
   begin
      // Suppression d'enregistrements
      While TauxVariables.Detail.Count > pNbTaux do
         TauxVariables.Detail[pNbTaux].Free;
   end
   else
   begin
      // Ajout d'enregistrements
      For lInCpt := TauxVariables.Detail.Count to pNbTaux - 1 do
          TOB.Create('FTAUXVAR',TauxVariables,-1);
   end;
end;


// ----------------------------------------------------------------------------
// Nom    : Simulation1
// Date   : 21/11/2001
// Auteur : D. ZEDIAR
// Objet  : Renvoie la date d'une échéance
// ----------------------------------------------------------------------------
// Historique :
// ----------------------------------------------------------------------------
function TObjetEmprunt.DateEcheance(pInEcheance : Integer) : TDateTime;
var
   lInMois, lInJour, lInAn, lInMois1 : Word;
begin
   Result := PlusMois(GetValue('EMP_DATEDEBUT'),pInEcheance * EMP_NBMOISECH);

   DecodeDate(GetValue('EMP_DATEDEBUT'),lInAn,lInMois,lInJour);
   DecodeDate(GetValue('EMP_DATEDEBUT')+1,lInAn,lInMois1,lInJour);
   if lInMois <> lInMois1 then
      Result := findemois(Result);
end;

// ----------------------------------------------------------------------------
// Nom    : Simulation2
// Date   : 21/11/2001
// Auteur : D. ZEDIAR
// Objet  : Simulation avec amortissement non calculé sur le taux initial
// ----------------------------------------------------------------------------
// Historique :
// ----------------------------------------------------------------------------
Procedure TObjetEmprunt.Simulation2(pInTypeRecherche : TTypeRecherche);
var
   lInCpt1       : Integer;
   lInCpt2       : Integer;
   lRdTaux       : Double;
   lRdTauxActuel : Double;
   lOmEmp        : TObjetEmprunt;
   lInCpt3       : Integer;
   lObEch        : Tob;
   lObEch2       : Tob;
begin
   // Execute une simulation classique
   Simulation1(pInTypeRecherche);

   // Découpe l'emprunt en autant d'emprunts que de taux
   lRdTaux  := GetValue('EMP_TAUXAN');
   lOmEmp   := TObjetEmprunt.CreerEmprunt;

   for lInCpt1 := 0 to Echeances.Detail.Count - 1 do
   begin
      // Détecte un changement de taux
      lRdTauxActuel := RechercherTaux(DateEcheance(lInCpt1));
      if lRdTaux <> lRdTauxActuel then
      begin
         lRdTaux := lRdTauxActuel;

         // Calcul un emprunt en découpant celui de base
         // Détermination des paramètres de l'emprunt à calculer :
         lOmEmp.Init;
         lOmEmp.Assigner(Self);
         lOmEmp.PutValue('EMP_EMPTYPEVERS', tvConstant );
         lOmEmp.PutValue('EMP_DATEDEBUT', DateEcheance(lInCpt1) );
         lOmEmp.PutValue('EMP_EMPTYPETAUX', ttConstant );
         lOmEmp.PutValue('EMP_DUREE', GetValue('EMP_DUREE') - EMP_NBMOISECH * lInCpt1 );
         lOmEmp.PutValue('EMP_TAUXAN', lRdTauxActuel );
         lOmEmp.PutValue('EMP_NBECHTOT', 0 );

         lOmEmp.fRdReportDeBase_int := 0;
         lOmEmp.fRdReportDeBase_ass := 0;
         if lInCpt1 > 0 then
         begin
            lObEch := Echeances.Detail[lInCpt1-1];
            lOmEmp.PutValue('EMP_CAPITAL', lObEch.GetValue('ECH_SOLDE') );
            lOmEmp.fRdReportDeBase_int := lObEch.GetValue('ECH_REPORTINT');
            lOmEmp.fRdReportDeBase_ass := lObEch.GetValue('ECH_REPORTASS');
         end;

         if GetValue('EMP_DIFREMB') > lInCpt1 then
            lOmEmp.PutValue('EMP_DIFREMB', GetValue('EMP_DIFREMB') - lInCpt1 )
         else
            lOmEmp.PutValue('EMP_DIFREMB', 0 );

         if GetValue('EMP_DIFINT') > lInCpt1 then
            lOmEmp.PutValue('EMP_DIFINT', GetValue('EMP_DIFINT') - lInCpt1 )
         else
            lOmEmp.PutValue('EMP_DIFINT', 0 );

         if GetValue('EMP_DIFASS') > lInCpt1 then
            lOmEmp.PutValue('EMP_DIFASS', GetValue('EMP_DIFASS') - lInCpt1 )
         else
            lOmEmp.PutValue('EMP_DIFASS', 0 );

         // Calcul de l'emprunt
         lOmEmp.Simulation1(pInTypeRecherche);

         // Recopie le résultat de la simulation dans l'emprunt courant
         for lInCpt2 := lInCpt1 to Echeances.Detail.Count - 1 do
         begin
            lObEch  := Echeances.Detail[lInCpt2];
            lObEch2 := lOmEmp.Echeances.Detail[lInCpt2 - lInCpt1];
//            lOmEmp.Echeances.ItemIndex := lInCpt2 - lInCpt1;
            For lInCpt3 := 1 to lObEch.NbChamps do
                lObEch.PutValue(lObEch.GetNomChamp(lInCpt3),lObEch2.GetValue(lObEch2.GetNomChamp(lInCpt3)) );
//                Echeances.Fields[lInCpt3].Value := lOmEmp.Echeances.Fields[lInCpt3].Value;
         end;
      end;
   end;

   // Recalcul de l'échéance totale et des cumuls pour chaque échéance
   PutValue('EMP_TOTINT', 0);
   PutValue('EMP_TOTASS', 0);
   PutValue('EMP_TOTAMORT', 0);
   PutValue('EMP_TOTVERS', 0);
   for lInCpt1 := 0 to Echeances.Detail.Count - 1 do
   begin
      lObEch  := Echeances.Detail[lInCpt1];

      PutValue('EMP_TOTINT',   GetValue('EMP_TOTINT') + lObEch.GetValue('ECH_INTERET') );
      PutValue('EMP_TOTASS',   GetValue('EMP_TOTASS') + lObEch.GetValue('ECH_ASSURANCE') );
      PutValue('EMP_TOTAMORT', GetValue('EMP_TOTAMORT') + lObEch.GetValue('ECH_AMORTISSEMENT') );
      PutValue('EMP_TOTVERS',  GetValue('EMP_TOTVERS') + lObEch.GetValue('ECH_ASSURANCE') +
                               lObEch.GetValue('ECH_INTERET') + lObEch.GetValue('ECH_AMORTISSEMENT') );

      lObEch.PutValue('ECH_CUMULA', GetValue('EMP_TOTAMORT') );
      lObEch.PutValue('ECH_CUMULI', GetValue('EMP_TOTINT') + GetValue('EMP_TOTASS') );
      lObEch.PutValue('ECH_CUMULV', GetValue('EMP_TOTVERS') );
   end;

end;

// ----------------------------------------------------------------------------
// Nom    : Simulation3
// Date   : 21/11/2001
// Auteur : D. ZEDIAR
// Objet  : Simulation avec amortissement calculé sur le taux initial
// ----------------------------------------------------------------------------
// Historique :
// ----------------------------------------------------------------------------
Procedure TObjetEmprunt.Simulation3(pInTypeRecherche : TTypeRecherche);
begin
   // Première phase : calcul classique de l'emprunt (en taux constant)
   PutValue('EMP_EMPTYPETAUX', ttConstant);
   fBoTauxReel := False;
   Simulation1(pInTypeRecherche);

   // Seconde phase : répercute les vrais taux (variables) sur l'emprunt
   PutValue('EMP_EMPTYPETAUX', ttVariable);
   fBoTauxReel := True;
   CalculerEmprunt(pInTypeRecherche,True);

   fBoTauxReel := False;
end;


// ----------------------------------------------------------------------------
// Nom    : CalcAmortCst
// Date   : 21/11/2001
// Auteur : D. ZEDIAR
// Objet  : fonction qui calcule le montant de l'amortissement pour l'emprunt
//          (pour amortissement constant)
// ----------------------------------------------------------------------------
// Historique :
// ----------------------------------------------------------------------------
function TObjetEmprunt.CalcAmortCst : double;
var
   lObEmp : TObjetEmprunt;
begin

   // Init de l'emprunt de travail
   lObEmp := TObjetEmprunt.CreerEmprunt;

   try
      lObEmp.Init;
      lObEmp.Assigner(self);
      lObEmp.PutValue('EMP_AMORTCST', 0);
      lObEmp.PutValue('EMP_VERSCST', 0);
      lObEmp.PutValue('EMP_EMPTYPEVERS', tvConstant);
      lObEmp.PutValue('EMP_EMPTYPETAUX', ttConstant);

      // Lance une simulation sur l'emprunt de travail ; élément recherché : versement
      lObEmp.Simulation1(trVersement);

      // Calcul de l'amortissement constant par échéance
      if (lObEmp.GetValue('EMP_DUREE')/lObEmp.EMP_NBMOISECH - GetValue('EMP_DIFREMB')) > 0 then
         Result := (lObEmp.GetValue('EMP_TOTAMORT'))/(lObEmp.GetValue('EMP_DUREE')/lObEmp.EMP_NBMOISECH - lObEmp.GetValue('EMP_DIFREMB'))
      else
         Result := 0;

   finally
      lObEmp.Free;
   end;

end;


// ----------------------------------------------------------------------------
// Nom    : CalcVersCst
// Date   : 21/11/2001
// Auteur : D. ZEDIAR
// Objet  : fonction qui calcule le montant du versement constant
//          (pour les simulations d'emprunt)
// ----------------------------------------------------------------------------
// Historique :
// ----------------------------------------------------------------------------
function TObjetEmprunt.CalcVersCst : double;
var
   lObEmp : TObjetEmprunt;
begin

   // Init de l'emprunt de travail
   lObEmp := TObjetEmprunt.CreerEmprunt;

   try
      lObEmp.Init;
      lObEmp.Assigner(self);
      lObEmp.PutValue('EMP_AMORTCST', 0 );
      lObEmp.PutValue('EMP_VERSCST', 0);
      lObEmp.PutValue('EMP_EMPTYPEVERS', tvConstant);
      if GetValue('EMP_EMPTYPETAUX') = ttVariable then
         lObEmp.PutValue('EMP_EMPTYPETAUX', ttConstant);

      // Lance une simulation sur l'emprunt de travail ; élément recherché : versement
      lObEmp.Simulation1(trVersement);

      // Calcul de l'amortissement constant par échéance
//      lObEmp.Echeances.ItemIndex := lObEmp.EMP_DIFREMB.AsInteger;
      Result := lObEmp.Echeances.Detail[lObEmp.GetValue('EMP_DIFREMB')].GetValue('ECH_VERSEMENT');

   finally
      lObEmp.Free;
   end;

end;


// ----------------------------------------------------------------------------
// Nom    : ProchaineDateExercice
// Date   : 18/12/2001
// Auteur : D. ZEDIAR
// Objet  : Calcul de la prochaine date d'exercice
// ----------------------------------------------------------------------------
// Historique :
// ----------------------------------------------------------------------------
Function TObjetEmprunt.ProchaineDateExercice(pDtDate : TDateTime) : TDateTime;
var
   lDtdateDeb, lDtDateFin : TDateTime;
begin
   if CalculerExercice(pDtDate,0,lDtdateDeb, lDtDateFin, fObTobExercices) then
      Result := lDtDateFin
   else
      Result := 0;
end;

{
var
   lInJour, lInMois, lInAn : Word;
begin
   DecodeDate(pDtDate,lInAn,lInMois,lInJour);

   if lInMois = fInMoisExercice then
   begin
      if lInJour > fInJourExercice then
         lInAn := lInAn + 1;
   end
   else
      if lInMois > fInMoisExercice then
         lInAn := lInAn + 1;

   Result := EncodeDate(lInAn,fInMoisExercice,fInJourExercice);
end;
}

{Procedure TObjetEmprunt.SetDateFinExercice(pDtDate : TDateTime);
var
   lInAn, lInMois, lInJour : Word;
begin
   DecodeDate(pDtDate,lInAn,lInMois,lInJour);
   JourExercice := lInJour;
   MoisExercice := lInMois;
end;
}
// ----------------------------------------------------------------------------
// Nom    : Ajustement
// Date   : 22/11/2001
// Auteur : D. ZEDIAR
// Objet  : Effectue un ajustement proportionnel de l'emprunt
// ----------------------------------------------------------------------------
// Historique :
// ----------------------------------------------------------------------------
Function TObjetEmprunt.Ajustement(pRdMontant : Double) : Boolean;
var
   lRdMntInt   : Double;
   lRdDelta    : Double;
   lRdDeltaEch : Double;
   lRdTxAjus   : Double;
   lInCpt      : Integer;
   lObEch      : Tob;
   PrecSolde : Double;

begin
   Result := False;

   // L'ajustement prend en compte l'assurance seleument si on est en taux
   lRdMntInt := GetValue('EMP_TOTINT');
   if GetValue('EMP_EMPTYPEASSUR') = taTaux then
      lRdMntInt := lRdMntInt + GetValue('EMP_TOTASS');

   if pRdMontant = lRdMntInt then
      Exit;

   lRdDelta := pRdMontant - lRdMntInt;
   lRdDeltaEch := lRdDelta / Echeances.Detail.Count;

   PrecSolde:= GetValue('EMP_CAPITAL');

   for lInCpt := 0 to Echeances.Detail.Count - 1 do
   begin
      lObEch := Echeances.Detail[lInCpt];

      lObEch.PutValue('ECH_VERSEMENT', lObEch.GetValue('ECH_VERSEMENT') + lRdDeltaEch );
      lObEch.PutValue('ECH_CUMULV', lObEch.GetValue('ECH_CUMULV') + lRdDeltaEch );

      lRdTxAjus := (lRdDeltaEch / lObEch.GetValue('ECH_VERSEMENT')) + 1;

      lObEch.PutValue('ECH_INTERET', lObEch.GetValue('ECH_INTERET') * lRdTxAjus );
      if GetValue('EMP_EMPTYPEASSUR') = taTaux then
         lObEch.PutValue('ECH_ASSURANCE', lObEch.GetValue('ECH_ASSURANCE') * lRdTxAjus );

      lObEch.PutValue('ECH_CUMULI', lObEch.GetValue('ECH_CUMULI') * lRdTxAjus );

      lObEch.PutValue('ECH_CUMULA', lObEch.GetValue('ECH_CUMULA') * lRdTxAjus );

      lObEch.PutValue('ECH_AMORTISSEMENT',  lObEch.GetValue('ECH_AMORTISSEMENT') * lRdTxAjus );

    {FQ20429  11.06.07  YMO Méthode de calcul}
      lObEch.PutValue('ECH_SOLDE', PrecSolde  - lObEch.GetValue('ECH_AMORTISSEMENT') );
{     lObEch.PutValue('ECH_SOLDE',         lObEch.GetValue('ECH_SOLDE')         - lObEch.GetValue('ECH_AMORTISSEMENT') * lRdTxAjus );
      lObEch.PutValue('ECH_DETTE',         lObEch.GetValue('ECH_DETTE')         - lObEch.GetValue('ECH_AMORTISSEMENT') * lRdTxAjus );
}
      PrecSolde:=lObEch.GetValue('ECH_SOLDE');

      lObEch.PutValue('ECH_MODIF', 'X');
   end;

   PutValue('EMP_TOTVERS', GetValue('EMP_TOTVERS') + lRdDelta );
   If (GetValue('EMP_TOTINT') + GetValue('EMP_TOTASS'))<>0 then {FQ21976 09.12.2007 YMO}
     lRdTxAjus := GetValue('EMP_TOTINT') / (GetValue('EMP_TOTINT') + GetValue('EMP_TOTASS'));
   PutValue('EMP_TOTINT', pRdMontant * lRdTxAjus );
   PutValue('EMP_TOTASS', pRdMontant * (1 - lRdTxAjus) );

   Result := True;

end;


// ----------------------------------------------------------------------------
// Nom    : Assign
// Date   : 21/11/2001
// Auteur : D. ZEDIAR
// Objet  : Copie des infos d'un emprunt dans l'emprunt courant
// ----------------------------------------------------------------------------
// Historique :
// ----------------------------------------------------------------------------
Procedure TObjetEmprunt.Assigner(pOmEmprunt : TObjetEmprunt);
var
   lInCpt         : Integer;
   lObTaux        : TOB;
   lObTauxOrigine : TOB;
begin
   Inherited;

   PutValue('EMP_CODEEMPRUNT',  pOmEmprunt.GetValue('EMP_CODEEMPRUNT') );
   PutValue('EMP_LIBEMPRUNT',   pOmEmprunt.GetValue('EMP_LIBEMPRUNT') );
   PutValue('EMP_NUMCOMPTE',    pOmEmprunt.GetValue('EMP_NUMCOMPTE') );
   PutValue('EMP_DATEDEBUT',    pOmEmprunt.GetValue('EMP_DATEDEBUT') );
   PutValue('EMP_DATECONTRAT',  pOmEmprunt.GetValue('EMP_DATECONTRAT') );
   PutValue('EMP_ORGPRETEUR',   pOmEmprunt.GetValue('EMP_ORGPRETEUR') );
   PutValue('EMP_GUIDORGPRETEUR',   pOmEmprunt.GetValue('EMP_GUIDORGPRETEUR') );
   PutValue('EMP_CAPITAL',      pOmEmprunt.GetValue('EMP_CAPITAL') );
   PutValue('EMP_EMPPERIODE',   pOmEmprunt.GetValue('EMP_EMPPERIODE') );
   PutValue('EMP_DUREE',        pOmEmprunt.GetValue('EMP_DUREE') );
   PutValue('EMP_EMPTYPETAUX',  pOmEmprunt.GetValue('EMP_EMPTYPETAUX') );
   PutValue('EMP_TAUXAN',       pOmEmprunt.GetValue('EMP_TAUXAN') );
   PutValue('EMP_EMPBQEQ',      pOmEmprunt.GetValue('EMP_EMPBQEQ') );
   PutValue('EMP_EMPTYPEASSUR', pOmEmprunt.GetValue('EMP_EMPTYPEASSUR') );
   PutValue('EMP_VALASSUR',     pOmEmprunt.GetValue('EMP_VALASSUR') );
   PutValue('EMP_DIFREMB',      pOmEmprunt.GetValue('EMP_DIFREMB') );
   PutValue('EMP_DIFINT',       pOmEmprunt.GetValue('EMP_DIFINT') );
   PutValue('EMP_DIFASS',       pOmEmprunt.GetValue('EMP_DIFASS') );
   PutValue('EMP_EMPTYPEVERS',  pOmEmprunt.GetValue('EMP_EMPTYPEVERS') );
   PutValue('EMP_AMORTCST',     pOmEmprunt.GetValue('EMP_AMORTCST') );
   PutValue('EMP_VERSCST',      pOmEmprunt.GetValue('EMP_VERSCST') );
   PutValue('EMP_AMORTTXINI',   pOmEmprunt.GetValue('EMP_AMORTTXINI') );
   PutValue('EMP_RAISON',       pOmEmprunt.GetValue('EMP_RAISON') );
   PutValue('EMP_EMPBASEASSUR', pOmEmprunt.GetValue('EMP_EMPBASEASSUR') );
   PutValue('EMP_ACTUANNU',     pOmEmprunt.GetValue('EMP_ACTUANNU') );
   PutValue('EMP_NBECHTOT',     pOmEmprunt.GetValue('EMP_NBECHTOT') );
   PutValue('EMP_GUIDORGPRETEUR',   pOmEmprunt.GetValue('EMP_GUIDORGPRETEUR') );
   fRdReportDeBase_int       := pOmEmprunt.fRdReportDeBase_int;
   fRdReportDeBase_ass       := pOmEmprunt.fRdReportDeBase_ass;
//   fInMoisExercice           := pOmEmprunt.fInMoisExercice;
//   fInJourExercice           := pOmEmprunt.fInJourExercice;

   // Assignation des taux variables
   CreerTaux(0);
   for lInCpt := 0 to pOmEmprunt.TauxVariables.Detail.Count - 1 do
   begin
      lObTauxOrigine := pOmEmprunt.TauxVariables.Detail[lInCpt];
      lObTaux := TOB.Create('FTAUXVAR',TauxVariables,-1);
      lObTaux.Putvalue('TXV_CODEEMPRUNT',lObTauxOrigine.GetValue('TXV_CODEEMPRUNT'));
      lObTaux.Putvalue('TXV_DATE',lObTauxOrigine.GetValue('TXV_DATE'));
      lObTaux.Putvalue('TXV_TAUX',lObTauxOrigine.GetValue('TXV_TAUX'));
   end;

end;


// ----------------------------------------------------------------------------
// Nom    : AffMessage
// Date   : 21/11/2001
// Auteur : D. ZEDIAR
// Objet  : Lance l'évènement d'affichage de message
// ----------------------------------------------------------------------------
// Historique :
// ----------------------------------------------------------------------------
Procedure TObjetEmprunt.AffMessage(pStMessage : String);
begin
//   if Assigned(fOnMessage) then
//      fOnMessage(Self,pStMessage);
end;

// ----------------------------------------------------------------------------
// Nom    : GetText
// Date   : 21/11/2001
// Auteur : D. ZEDIAR
// Objet  : retourne le tableau d'amortissement de l'emprunt sous forme de
//          texte
// ----------------------------------------------------------------------------
// Historique :
// ----------------------------------------------------------------------------
{Function TObjetEmprunt.GetText : String;
var
   lInCpt : Integer;
   lObEch : Tob;
begin

   Result := '';

   Result := Justifier('n',4,' ');
   Result := Result + Justifier('Date',12,' ');
   Result := Result + Justifier('Tx. Int',8,' ');
   Result := Result + Justifier('Tx. As.',8,' ');
   Result := Result + Justifier('Amortissement',15,' ');
   Result := Result + Justifier('Intéret',12,' ');
   Result := Result + Justifier('Assurance',12,' ');
   Result := Result + Justifier('Versement',15,' ');
   Result := Result + Justifier('Solde',15,' ');
   Result := Result + Justifier('Dette',15,' ');

   Result := Result + chr(13);

   Result := Result + ' ' + Justifier('',3,'-');
   Result := Result + ' ' + Justifier('',11,'-');
   Result := Result + ' ' + Justifier('',7,'-');
   Result := Result + ' ' + Justifier('',7,'-');
   Result := Result + ' ' + Justifier('',14,'-');
   Result := Result + ' ' + Justifier('',11,'-');
   Result := Result + ' ' + Justifier('',11,'-');
   Result := Result + ' ' + Justifier('',14,'-');
   Result := Result + ' ' + Justifier('',14,'-');
   Result := Result + ' ' + Justifier('',14,'-');

   for lInCpt := 0 to Echeances.Detail.Count - 1 do
   begin
       Result := Result + chr(13);
       lObEch := Echeances.Detail[lInCpt];
       Result := Result + Justifier(IntTostr(lInCpt + 1),4,' ');
       Result := Result + Justifier(DateTostr(lObEch.GetValue('ECH_DATE')),12,' ');
       Result := Result + Justifier(FormaterMontant(lObEch.GetValue('ECH_TAUXINT'),4),8,' ');
       Result := Result + Justifier(FormaterMontant(lObEch.GetValue('ECH_TAUXASS'),4),8,' ');
       Result := Result + Justifier(FormaterMontant(lObEch.GetValue('ECH_AMORTISSEMENT'),2),15,' ');
       Result := Result + Justifier(FormaterMontant(lObEch.GetValue('ECH_INTERET'),2),12,' ');
       Result := Result + Justifier(FormaterMontant(lObEch.GetValue('ECH_ASSURANCE'),2),12,' ');
       Result := Result + Justifier(FormaterMontant(lObEch.GetValue('ECH_VERSEMENT'),2),15,' ');
       Result := Result + Justifier(FormaterMontant(lObEch.GetValue('ECH_SOLDE'),2),15,' ');
       Result := Result + Justifier(FormaterMontant(lObEch.GetValue('ECH_DETTE'),2),15,' ');
   end;

   // Totaux
   Result := Result + chr(13);

   Result := Result + ' ' + StringOfChar(' ',3);
   Result := Result + ' ' + StringOfChar(' ',11);
   Result := Result + ' ' + StringOfChar(' ',7);
   Result := Result + ' ' + StringOfChar(' ',7);
   Result := Result + ' ' + StringOfChar('-',14);
   Result := Result + ' ' + StringOfChar('-',11);
   Result := Result + ' ' + StringOfChar('-',11);
   Result := Result + ' ' + StringOfChar('-',14);
   Result := Result + ' ' + StringOfChar(' ',14);
   Result := Result + ' ' + StringOfChar(' ',14);

   Result := Result + chr(13);

   Result := Result + ' ' + StringOfChar(' ',3);
   Result := Result + ' ' + StringOfChar(' ',11);
   Result := Result + ' ' + StringOfChar(' ',7);
   Result := Result + ' ' + StringOfChar(' ',7);
   Result := Result + Justifier(FormaterMontant(GetValue('EMP_TOTAMORT'),2),15,' ');
   Result := Result + Justifier(FormaterMontant(GetValue('EMP_TOTINT'),2),12,' ');
   Result := Result + Justifier(FormaterMontant(GetValue('EMP_TOTASS'),2),12,' ');
   Result := Result + Justifier(FormaterMontant(GetValue('EMP_TOTVERS'),2),15,' ');
   Result := Result + ' ' + StringOfChar(' ',14);
   Result := Result + ' ' + StringOfChar(' ',14);
end;
}



// ----------------------------------------------------------------------------
//
//             G E S T I O N   D E S   T A U X   V A R I A B L E S
//
// ----------------------------------------------------------------------------

// ----------------------------------------------------------------------------
// Nom    : SupprimerUnTaux
// Date   : 13/11/2001
// Auteur : D. ZEDIAR
// Objet  : Supprime un taux à partir d'une date
// ----------------------------------------------------------------------------
// Historique :
// ----------------------------------------------------------------------------
Procedure TObjetEmprunt.SupprimerUnTaux(pDtDate : TDateTime);
var
   lInIndexe : Integer;
begin
   For lInIndexe := 0 to TauxVariables.Detail.Count - 1 do
   begin
      if TauxVariables.Detail[lInIndexe].GetValue('TXV_DATE') = pDtDate then
      begin
         TauxVariables.Detail[lInIndexe].Free;
         Break;
      end;
   end;
end;


// ----------------------------------------------------------------------------
// Nom    : AjouterTaux
// Date   : 21/11/2001
// Auteur : D. ZEDIAR
// Objet  : Ajout d'un taux variable à la liste
// ----------------------------------------------------------------------------
// Historique :
// ----------------------------------------------------------------------------
Procedure TObjetEmprunt.AjouterTaux(pDtDate : TDateTime; pRdTaux : Double);
var
   lInCpt    : Integer;
   lBoFound  : Boolean;
   lObTaux   : Tob;
begin
   lObTaux := nil;
   lBoFound := False;
   For lInCpt := 0 to TauxVariables.Detail.Count - 1 do
   begin
      lObTaux := TauxVariables.Detail[lInCpt];
      if lObTaux.GetValue('TXV_DATE') = pDtDate then
      begin
         lBoFound := True;
         Break;
      end;
   end;

   if not lBoFound then
   begin
      lObTaux := Tob.Create('FTAUXVAR',TauxVariables,-1);
      lObTaux.PutValue('TXV_CODEEMPRUNT', GetValue('EMP_CODEEMPRUNT') );
      lObTaux.PutValue('TXV_DATE', pDtDate);
   end;

   lObTaux.PutValue('TXV_TAUX', pRdTaux);

end;


// ----------------------------------------------------------------------------
// Nom    : RechercherTaux
// Date   : 21/11/2001
// Auteur : D. ZEDIAR
// Objet  : retourne le taux variable pour une date donnée
// ----------------------------------------------------------------------------
// Historique :
// ----------------------------------------------------------------------------
Function TObjetEmprunt.RechercherTaux(pDtDate : TDateTime) : Double;
var
   lInCpt     : Integer;
   lDtMaxDate : TDateTime;
   lTbTaux    : Tob;
begin
   Result := GetValue('EMP_TAUXAN');
   lDtMaxDate := 0;

   For lInCpt := 0 to TauxVariables.Detail.Count - 1 do
   begin
      lTbTaux := TauxVariables.Detail[lInCpt];
      if (lTbTaux.GetValue('TXV_DATE') <= pDtDate) and
         (lTbTaux.GetValue('TXV_DATE') > lDtMaxDate) then
      begin
         lDtMaxDate := lTbTaux.GetValue('TXV_DATE');
         Result := lTbTaux.GetValue('TXV_TAUX');
      end;
   end;
end;


Function TObjetEmprunt.EmpruntSuivant : Integer;
var
   lTbEmprunt : Tob;
begin
   Result := 1;
   lTbEmprunt := Tob.Create('',Nil,-1);
   try
      lTbEmprunt.LoadDetailFromSQL('Select Max(EMP_CODEEMPRUNT) As NB From FEMPRUNT');
      if lTbEmprunt.Detail.Count > 0 then
      begin
      try
      if not VarIsNull(lTbEmprunt.Detail[0].GetValue('NB')) then
     //   Result := lTbEmprunt.Detail[0].GetValue('NB') + 1; Modif Chri.A du 10/05/06 compatibilité avec Oracle
        Result := lTbEmprunt.Detail[0].GetInteger('NB') + 1;
      except
        Result:=1;
      end;
      end;
   finally
      lTbEmprunt.Free;
   end;
end;


// ----------------------------------------------------------------------------
// Nom    : EcheancesVersGrille
// Date   : 06/11/2002
// Auteur : D. ZEDIAR
// Objet  : Remplissage d'une grille avec les échéances de l'emprunt
// ----------------------------------------------------------------------------
// Historique :
// ----------------------------------------------------------------------------
Procedure TObjetEmprunt.EcheancesVersGrille(pObGrille : THGrid);
var
   lInCpt          : Integer;
   lInRow          : Integer;
   lRdTotInt       : Double;
   lRdTotAss       : Double;
   lRdTotAmort     : Double;
   lRdTotVers      : Double;
   lDtDateExercice : TDateTime;
   lBoFinExercice  : Boolean;
   lObEch          : Tob;
   lObEchSuivante  : Tob;

   Procedure TotalExercice;
   begin
      lInRow := lInRow + 1;
      pObGrille.RowCount := lInRow + 1;
      pObGrille.Cells[cColTypeLigne,lInRow]     := cStLigneTotal;
      pObGrille.Cells[cColModif,lInRow]         := '';
      pObGrille.Cells[cColNumEch,lInRow]        := 'E';
      pObGrille.Cells[cColDate,lInRow]          := DateToStr(lObEch.GetValue('ECH_DATEEXERCICE'));
      pObGrille.Cells[cColTxInteret,lInRow]     := '';
      pObGrille.Cells[cColTxAssurance,lInRow]   := '';
      pObGrille.Cells[cColTxGlobal,lInRow]      := '';
      pObGrille.Cells[cColAmortissement,lInRow] := FormaterMontant(lRdTotAmort);
      pObGrille.Cells[cColInteret,lInRow]       := FormaterMontant(lRdTotInt);
      pObGrille.Cells[cColAssurance,lInRow]     := FormaterMontant(lRdTotAss);
      pObGrille.Cells[cColTotInteret,lInRow]    := FormaterMontant(lRdTotInt+lRdTotAss);
      pObGrille.Cells[cColVersement,lInRow]     := FormaterMontant(lRdTotVers);
      pObGrille.Cells[cColSolde,lInRow]         := '';
      pObGrille.Cells[cColReport,lInRow]        := '';
      pObGrille.Cells[cColDette,lInRow]         := '';
      pObGrille.Cells[cColCumulA,lInRow]        := '';
      pObGrille.Cells[cColCumulI,lInRow]        := '';
      pObGrille.Cells[cColCumulM,lInRow]        := '';
      pObGrille.Cells[cColModif,lInRow]         := '';

      pObGrille.RowLengths[lInRow] := -1;

      // Réinit des totaux
      lRdTotInt   := 0;
      lRdTotAss   := 0;
      lRdTotAmort := 0;
      lRdTotVers  := 0;
   end;

begin
   // Titres des colonnes
   pObGrille.Cells[cColNumEch,0]        := 'n';
   pObGrille.Cells[cColDate,0]          := 'Date';
   pObGrille.Cells[cColTxInteret,0]     := 'Tx. Int.';
   pObGrille.Cells[cColTxAssurance,0]   := 'Tx. Ass';
   pObGrille.Cells[cColTxGlobal,0]      := 'Taux';
   pObGrille.Cells[cColInteret,0]       := 'Intérets';
   pObGrille.Cells[cColAssurance,0]     := 'Assurance';
   pObGrille.Cells[cColTotInteret,0]    := 'Int. + Ass.';
   pObGrille.Cells[cColAmortissement,0] := 'Amortissement';
   pObGrille.Cells[cColVersement,0]     := 'Versement';
   pObGrille.Cells[cColSolde,0]         := 'Solde';
   pObGrille.Cells[cColReport,0]        := 'Report';
   pObGrille.Cells[cColDette,0]         := 'Dette';
   pObGrille.Cells[cColCumulA,0]        := 'Cumul A.';
   pObGrille.Cells[cColCumulI,0]        := 'Cumul I.';
   pObGrille.Cells[cColCumulM,0]        := 'Cumul V.';
   pObGrille.Cells[cColTypeLigne,0]     := 'Type';
   pObGrille.Cells[cColModif,0]         := 'Modif';

   pObGrille.ColAligns[cColNumEch]        := taCenter;
   pObGrille.ColAligns[cColDate]          := taRightJustify;
   pObGrille.ColAligns[cColTxInteret]     := taRightJustify;
   pObGrille.ColAligns[cColTxAssurance]   := taRightJustify;
   pObGrille.ColAligns[cColTxGlobal]      := taRightJustify;
   pObGrille.ColAligns[cColAmortissement] := taRightJustify;
   pObGrille.ColAligns[cColInteret]       := taRightJustify;
   pObGrille.ColAligns[cColAssurance]     := taRightJustify;
   pObGrille.ColAligns[cColTotInteret]    := taRightJustify;
   pObGrille.ColAligns[cColVersement]     := taRightJustify;
   pObGrille.ColAligns[cColSolde]         := taRightJustify;
   pObGrille.ColAligns[cColReport]        := taRightJustify;
   pObGrille.ColAligns[cColDette]         := taRightJustify;
   pObGrille.ColAligns[cColCumulA]        := taRightJustify;
   pObGrille.ColAligns[cColCumulI]        := taRightJustify;
   pObGrille.ColAligns[cColCumulM]        := taRightJustify;

   pObGrille.ColTypes[cColDate]   := 'D';
   pObGrille.ColFormats[cColDate] := ShortDateFormat;

   pObGrille.ColTypes[cColAmortissement] := 'R';
   pObGrille.ColTypes[cColAssurance]     := 'R';
   pObGrille.ColTypes[cColTxInteret]     := 'R';
   pObGrille.ColTypes[cColTxAssurance]   := 'R';
   pObGrille.ColTypes[cColInteret]       := 'R';
   pObGrille.ColTypes[cColVersement]     := 'R';
   pObGrille.ColTypes[cColSolde]         := 'R';

   pObGrille.ColFormats[cColAmortissement] := cFormatMontant;
   pObGrille.ColFormats[cColAssurance]     := cFormatMontant;
   pObGrille.ColFormats[cColTxInteret]     := cFormatTaux;
   pObGrille.ColFormats[cColTxAssurance]   := cFormatTaux;
   pObGrille.ColFormats[cColInteret]       := cFormatMontant;
   pObGrille.ColFormats[cColVersement]     := cFormatMontant;
   pObGrille.ColFormats[cColSolde]         := cFormatMontant;

   // Init des totaux
   lRdTotInt    := 0 ;
   lRdTotAss    := 0 ;
   lRdTotAmort  := 0 ;
   lRdTotVers   := 0 ;

   lInRow := 0;

   // Remplissage de la grille
   For lInCpt := 1 to Echeances.Detail.Count do
   begin
      lObEch := Echeances.Detail[lInCpt - 1];

      // Totaux
      lRdTotInt   := lRdTotInt   + Arrondi(lObEch.GetValue('ECH_INTERET')       ,cInNbDecMnt);
      lRdTotAss   := lRdTotAss   + Arrondi(lObEch.GetValue('ECH_ASSURANCE')     ,cInNbDecMnt);
      lRdTotAmort := lRdTotAmort + Arrondi(lObEch.GetValue('ECH_AMORTISSEMENT') ,cInNbDecMnt);
      lRdTotVers  := lRdTotVers  + Arrondi(lObEch.GetValue('ECH_VERSEMENT')     ,cInNbDecMnt);

      //
      lInRow := lInRow + 1;
      pObGrille.RowCount := lInRow + 1;

      // Ligne échéance
      lObEch.PutLigneGrid(pObGrille,lInRow,False,False,cChampsColonnes);
      pObGrille.Cells[cColTypeLigne,lInRow]     := cStLigneEcheance;
      pObGrille.Cells[cColNumEch,lInRow]        := IntToStr(lInCpt);
      pObGrille.Cells[cColTxGlobal,lInRow]      := FormaterTaux(lObEch.GetValue('ECH_TAUXINT') + lObEch.GetValue('ECH_TAUXASS'));
      pObGrille.Cells[cColTotInteret,lInRow]    := FormaterMontant(lObEch.GetValue('ECH_INTERET') + lObEch.GetValue('ECH_ASSURANCE'));

      // totaux fin d'exercice :
      // Détecte une rupture dans la date d'exercice
      if lInCpt = GetValue('EMP_NBECHTOT') then
         lBoFinExercice := True
      else
      begin
         //
         lObEchSuivante  := Echeances.Detail[lInCpt];
         lDtDateExercice := lObEchSuivante.GetValue('ECH_DATEEXERCICE');
         lBoFinExercice  := lObEch.GetValue('ECH_DATEEXERCICE') <> lDtDateExercice;
      end;

      // Affichage des totaux de l'exercice si fin d'exercice
      if lBoFinExercice then
         TotalExercice;
   end;

   // Dernière ligne de totalisation
   if pObGrille.Cells[cColTypeLigne,lInRow] = cStLigneEcheance then
      TotalExercice;

end;

{***********A.G.L.***********************************************
Auteur  ...... : YMO
Créé le ...... : 16/03/2007
Modifié le ... :   /  /
Description .. : Calcul des périodes en remplacement des exercices pour
Suite ........ : les états
Mots clefs ... :
*****************************************************************}
Class Function TObjetEmprunt.CalculerPeriode(PerExo : Integer ; pDtDateRef : TDateTime; pInN : Integer; Var pDtDateDebPer : TDateTime; Var pDtDateFinPer : TDateTime; pObTobPeriode : TOB = Nil) : Boolean;
var
   lInJour, lInMois, lInAn : Word;
begin
   {25/04/07 Si l'exercice n'existe pas, on garde 12 mois glissants}
   If (PerExo=2) and (not ExisteSQl('SELECT EX_DATEDEBUT, EX_DATEFIN FROM EXERCICE'
                +' WHERE EX_DATEDEBUT<="' + UsDateTime(pDtDateRef)
                +'" AND EX_DATEFIN>="'+ UsDateTime(pDtDateRef) + '"')) then
      PerExo:=1;

  If PerExo=1 then
  begin
    {12 mois glissants avec comme fin de période la date de référence}
    DecodeDate(pDtDateRef,lInAn,lInMois,lInJour);
    pDtDateDebPer := EncodeDate(lInAn +pInN -1, lInMois, lInJour)+1;
    pDtDateFinPer := EncodeDate(lInAn +pInN, lInMois, lInJour);;
    Result := True;
  end
  else
  If PerExo=2 then
  begin
    {25/04/07 X mois glissants avec comme début de période le début d'exercice}
    Result := CalculerExercice(pDtDateRef, pInN, pDtDateDebPer, pDtDateFinPer, pObTobPeriode);
    DecodeDate(pDtDateRef,lInAn,lInMois,lInJour);
    pDtDateFinPer := EncodeDate(lInAn +pInN, lInMois, lInJour);
  end
  else
    Result := CalculerExercice(pDtDateRef, pInN, pDtDateDebPer, pDtDateFinPer, pObTobPeriode);

end;

// ----------------------------------------------------------------------------
// Nom    : CalculerExercice
// Date   : 16/06/2003
// Auteur : D. ZEDIAR
// Objet  : calcul d'une date d'exercice
// ----------------------------------------------------------------------------
// Historique :
// ----------------------------------------------------------------------------
Class Function TObjetEmprunt.CalculerExercice(pDtDateRef : TDateTime; pInN : Integer; Var pDtDateDebExe : TDateTime; Var pDtDateFinExe : TDateTime; pObTobExercice : TOB = Nil) : Boolean;
var
   lBoSupprimerTob         : Boolean;
   lInCpt                  : Integer;
   lInIndiceExe            : Integer;
   lInJour, lInMois, lInAn : Word;

   Procedure CreerExerciceVirtuelAnterieur;
   var
      lDtDate1, lDtDate2 : TDateTime;
   begin
      With Tob.Create('EXERCICE',pObTobExercice,0) do
      begin
         lDtDate2 := pObTobExercice.Detail[1].GetValue('EX_DATEDEBUT')-1;
         lDtDate1 := PlusMois(lDtDate2,-12) + 1;
         PutValue('EX_DATEDEBUT',lDtDate1);
         PutValue('EX_DATEFIN',lDtDate2);
      end;
   end;

   Procedure CreerExerciceVirtuelPosterieur;
   var
      lDtDate1, lDtDate2 : TDateTime;
   begin
      With Tob.Create('EXERCICE',pObTobExercice,pObTobExercice.Detail.Count) do
      begin
         lDtDate1 := pObTobExercice.Detail[pObTobExercice.Detail.Count-2].GetValue('EX_DATEFIN')+1;
         lDtDate2 := PlusMois(lDtDate1,12) - 1;
         PutValue('EX_DATEDEBUT',lDtDate1);
         PutValue('EX_DATEFIN',lDtDate2);
      end;
   end;

begin

   if Assigned(pObTobExercice) then
      lBoSupprimerTob := false
   else
   begin
      lBoSupprimerTob := True;
      pObTobExercice := Tob.Create('',Nil,-1);
      pObTobExercice.LoadDetailFromSQL('Select EX_DATEDEBUT, EX_DATEFIN From EXERCICE Order By EX_DATEDEBUT');
   end;

   try
      if pObTobExercice.Detail.Count > 0 then
      begin
         pDtDateDebExe := pObTobExercice.Detail[0].GetValue('EX_DATEDEBUT');
         pDtDateFinExe := pObTobExercice.Detail[pObTobExercice.Detail.Count-1].GetValue('EX_DATEFIN');

         if pDtDateRef < pDtDateDebExe then
         begin
            CreerExerciceVirtuelAnterieur;
            Result := CalculerExercice(pDtDateRef,pInN,pDtDateDebExe,pDtDateFinExe,pObTobExercice);
         end
         else
            if pDtDateRef >= pDtDateFinExe + 1 then
            begin
               CreerExerciceVirtuelPosterieur;
               Result := CalculerExercice(pDtDateRef,pInN,pDtDateDebExe,pDtDateFinExe,pObTobExercice);
            end
            else
            begin
               // Rechèrche l'exercice 0
               lInIndiceExe := -1;
               for lInCpt := 0 to pObTobExercice.Detail.Count - 1 do
                  With pObTobExercice.Detail[lInCpt] do
                     if (GetValue('EX_DATEDEBUT') <= pDtDateRef) And (GetValue('EX_DATEFIN')+1 > pDtDateRef) then
                     begin
                        lInIndiceExe := lInCpt;
                        Break;
                     end;

               if lInIndiceExe = -1 then
               begin
                  DecodeDate(pDtDateRef,lInAn,lInMois,lInJour);
                  pDtDateDebExe := EncodeDate(lInAn + pInN,1,1);
                  pDtDateFinExe := EncodeDate(lInAn + pInN,12,31);
                  Result := True;
               end
               else
               begin
                  // Rechèrche de l'exercice N
                  lInIndiceExe := lInIndiceExe + pInN;
                  if lInIndiceExe < 0 then
                  begin
                     CreerExerciceVirtuelAnterieur;
                     Result := CalculerExercice(pDtDateRef,pInN,pDtDateDebExe,pDtDateFinExe,pObTobExercice);
                  end
                  else
                    begin
                        if lInIndiceExe >= pObTobExercice.Detail.Count then
                        begin
                           CreerExerciceVirtuelPosterieur;
                           Result := CalculerExercice(pDtDateRef,pInN,pDtDateDebExe,pDtDateFinExe,pObTobExercice);
                        end
                        else
                        begin
                           With pObTobExercice.Detail[lInIndiceExe] do
                           begin
                               pDtDateDebExe := GetValue('EX_DATEDEBUT');
                               pDtDateFinExe := GetValue('EX_DATEFIN');
                           end;
                           Result := True;
                        end;
                    end;
               end;
            end;

      end
      else
      begin
         DecodeDate(pDtDateRef,lInAn,lInMois,lInJour);
         pDtDateDebExe := EncodeDate(lInAn + pInN,1,1);
         pDtDateFinExe := EncodeDate(lInAn + pInN,12,31);
         Result := True;
      end;

   finally
      if lBoSupprimerTob then pObTobExercice.Free;
   end;

end;

end.
