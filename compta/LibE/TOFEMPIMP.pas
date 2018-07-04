{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 13/12/2001
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : CAMORTCST ()
Mots clefs ... : TOF;TOFAMORTCST
*****************************************************************}
Unit TOFEMPIMP;

Interface

Uses StdCtrls,
     Controls,
     Graphics,
     Classes,
     buttons,
     HDB,
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}
{$ENDIF}
{$IFDEF VER150}
     Variants,
{$ENDIF}
     forms,
     UTOB,
     Windows,
     sysutils,
     ComCtrls,
     HCtrls,
     //PGIEnv,
     HEnt1,
     HMsgBox,
     HTB97,
     UTOF,
{$IFDEF EAGLCLIENT}
     eQRS1,
     MaineAgl,
{$ELSE}
     QRS1,
     FE_Main,
{$ENDIF}
     Ent1,
     TOFEMPCHOIXIMP,
     UObjetEmprunt,
     dialogs,
     Math,  // RoundTo
     (* Ajout GED *)
     UGedViewer;
     (* Fin Ajout GED *)

Const {YMO 19/03/2007 Détemrine si 'lon travaille par exo ou 12 mois glissants}
  PerExo: integer = 1; //1 : 12 mois glissants ; 2 : Exercice

Type
  TOF_EMPIMP = Class (TOF)
    private
       fOnbValiderClick : TNotifyEvent;
       fInCodeEmprunt   : Integer;
       fChainePeriode1  : String;
       fChainePeriode2  : String;

       procedure bValiderClick(Sender : TObject);
       procedure ChoixEmpruntClick(Sender : TObject);

       Function ChoixDateReference : TDateTime;
       Function ExtractionDonneesTabAmo : Boolean;
       Function ExtractionDonneesEcheancier : Boolean;
       Function ExtractionDonneesPrevisionnel : Boolean;
       Function ExtractionDonneesTableauEmprunt : Boolean;
       Function ExtractionDonneesTableauFF : Boolean;

    published
       procedure OnArgument (S : String ) ; override ;
       procedure OnClose                  ; override ;
  end ;



Procedure LancerEditions;

// Impression d'un emprunt
Procedure ImprimeEmprunt(pInCodeEmprunt : Integer; pInLibEmprunt : String; pBoListeExportation : Boolean = False; pCodeEtat : String='AMO');  {FQ20074  01.06.05  YMO}

// Impression de l'échéancier sur 12 mois
//Procedure ImprimeEcheancier;

// Impression du prévisionnel sur 5 ans
//Procedure ImprimePrevisionnel;

// Impression du tableau des emprunts
//Procedure ImprimeTableauEmprunts;

// Détermine les dates d'un exercice par raport à une date de référence
//Function CalculerExercice(pDtDateRef : TDateTime; pInN : Integer; Var pDtDateDebExe : TDateTime; Var pDtDateFinExe : TDateTime) : Boolean;


Implementation

uses
  {$IFDEF MODENT1}
  CPTypeCons,
  {$ENDIF MODENT1}
  uLibExercice ;

Procedure LancerEditions;
begin
   AGLLanceFiche('FP','FEMPIMP','','',';;');
end;



// ----------------------------------------------------------------------------
// Nom    : ImprimeEmprunt
// Date   : 13/09/2001
// Auteur : D. ZEDIAR
// Objet  : Impression d'un emprunt
// ----------------------------------------------------------------------------
// Historique :
// ----------------------------------------------------------------------------
Procedure ImprimeEmprunt(pInCodeEmprunt : Integer; pInLibEmprunt : String; pBoListeExportation : Boolean = False; pCodeEtat : String='AMO');  {FQ20074  01.06.05  YMO}
var
   lStArg : String;
begin
   lStArg := pCodeEtat+';';
   {FQ20223 23.05.07 YMO affichage de tous les emprunts}
   If pInCodeEmprunt>0 then lStArg := lStArg + IntToStr(pInCodeEmprunt)+';'+pInLibEmprunt+';';
   if pBoListeExportation then
      lStArg := lStArg + 'EXPORTATION';

   AGLLanceFiche('FP','FEMPIMP','','',lStArg);
end;


Function TOF_EMPIMP.ChoixDateReference : TDateTime;
begin
   try
      Result := StrToDate(GetControlText('edDateRef'));
   except
      Result := Date;
   end;
end;

Function TOF_EMPIMP.ExtractionDonneesTabAmo : Boolean;
var
   lStSQL                  : String;
   lTobDonnees             : Tob;
   lTobEcheance            : Tob;
   lTobEcheancier          : Tob;
   lInAncCodeEmprunt       : Integer;
   lInCpt1                 : Integer;
begin
   //Result := False;
   ExecuteSQL('Delete From FECHEANCIER');

   // Requête SQL
   lStSQL := ' Select * From FEMPRUNT a';
   lStSQL := lStSQL + ' Left Outer Join ANNUAIRE b On ANN_GUIDPER = a.EMP_GUIDORGPRETEUR';
   lStSQL := lStSQL + ' WHERE a.EMP_CODEEMPRUNT = ' + IntToStr(fInCodeEmprunt);
   lStSQL := lStSQL + ' Order By EMP_CODEEMPRUNT';

   lTobDonnees    := Tob.Create('',Nil,-1);
   lTobEcheancier := Tob.Create('FECHEANCIER',Nil,-1);
   try
      // Extraction des données
      lTobDonnees.LoadDetailFromSQL(lStSQL);

      lInAncCodeEmprunt :=-1;
      // Boucle
      For lInCpt1 := 0 to lTobDonnees.Detail.Count - 1 do
      begin
           lTobEcheance := lTobDonnees.Detail[lInCpt1];
           lTobEcheancier.InsertDB(Nil);
           lInAncCodeEmprunt := lTobEcheance.GetValue('EMP_CODEEMPRUNT');
           lTobEcheancier.Putvalue('ECR_CODEEMPRUNT',lInAncCodeEmprunt);
           lTobEcheancier.PutValue('ECR_LIBELLE2',IntToStr(lTobEcheance.GetValue('EMP_CODEEMPRUNT')) + ' - ' + lTobEcheance.GetValue('EMP_LIBEMPRUNT'));
           lTobEcheancier.PutValue('ECR_LIBELLE2',VarToStr(lTobEcheance.GetValue('ANN_NOM1')));
           lTobEcheancier.PutValue('ECR_DEVISE',VarToStr(lTobEcheance.GetValue('EMP_DEVISE')));

      end;

      // Nécéssaire
      if lInAncCodeEmprunt <> -1 then
         lTobEcheancier.InsertDB(Nil);

   finally
      lTobDonnees.Free;
      lTobEcheancier.Free;
   end;

   Result := true;

end;


// ----------------------------------------------------------------------------
// Nom    : ImprimeEcheancier
// Date   : 13/09/2001
// Auteur : D. ZEDIAR
// Objet  : Impression de l'échéancier sur 12 mois
// ----------------------------------------------------------------------------
// Historique :
// ----------------------------------------------------------------------------
{Procedure ImprimeEcheancier;
begin
   if ExtractionDonneesEcheancier then
      AglLanceFiche('FP','FEMPIMP','','','ECR;');
end;
}
Function TOF_EMPIMP.ExtractionDonneesEcheancier : Boolean;
var
   lDtDateRef              : TDateTime;
   lInMoisDeBase           : Word;
   lInDateDu               : TDateTime;
   lInDateAu               : TDateTime;
   lStSQL                  : String;
   lTobDonnees             : Tob;
   lTobEcheance            : Tob;
   lTobEcheancier          : Tob;
   lInCpt1                 : Integer;
   lInCpt2                 : Integer;
   lInAncCodeEmprunt       : Integer;
   lStNomChamps            : String;
   lInAn, lInMois, lInJour : Word;

begin
   Result := False;

   // Saisie de la date de référence
   lDtDateRef := ChoixDateReference;
   if lDtDateRef <= 0 then Exit;

   // Dates limites

   Decodedate(lDtDateRef, lInAn, lInMois, lInJour);
   {FQ20137 YMO 30.05.2007  On prend la valeur du mois sans se soucier du jour d'échéance}
   lInDateDu := Encodedate(lInAn, lInMois, 1);
   {FQ20137 YMO 21.06.2007  01/01/07 au 01/01/08 et non 31/12/08}
   lInDateAu := Encodedate(lInAn+1, lInMois, 1);

   lInMoisDeBase := lInMois;
   ExecuteSQL('Delete From FECHEANCIER');

   // Requête SQL
   lStSQL := ' Select * From FEMPRUNT a';
   lStSQL := lStSQL + ' Left Outer Join ANNUAIRE b On ANN_GUIDPER = a.EMP_GUIDORGPRETEUR';
   lStSQL := lStSQL + ' Left Join FECHEANCE c On c.ECH_CODEEMPRUNT = a.EMP_CODEEMPRUNT';
   lStSQL := lStSQL + ' Where ECH_DATE >= "' + UsDateTime(lInDateDu) + '" And';
   lStSQL := lStSQL + ' ECH_DATE < "' + UsDateTime(lInDateAu) + '"';
   If fInCodeEmprunt<>0 then  {FQ20641  08.06.07  YMO Choix d'un seul emprunt}
      lStSQL := lStSQL + ' AND a.EMP_CODEEMPRUNT = ' + IntToStr(fInCodeEmprunt);
   lStSQL := lStSQL + ' Order By ECH_CODEEMPRUNT, ECH_DATE';

   lTobDonnees    := Tob.Create('',Nil,-1);
   lTobEcheancier := Tob.Create('FECHEANCIER',Nil,-1);
   try
      // Initialisation des dates de l'échéancier
      For lInCpt1 := 0 to 11 do
         lTobEcheancier.PutValue('ECR_DATE' + IntToStr(lInCpt1 + 1),PlusMois(lDtDateRef,lInCpt1));

      // Extraction des données
      lTobDonnees.LoadDetailFromSQL(lStSQL);

      // Boucle
      lInAncCodeEmprunt := -1;
      For lInCpt1 := 0 to lTobDonnees.Detail.Count - 1 do
      begin
         lTobEcheance := lTobDonnees.Detail[lInCpt1];

         // Rupture sur le code emprunt
         if lInAncCodeEmprunt <> lTobEcheance.GetValue('ECH_CODEEMPRUNT') then
         begin
            if lInAncCodeEmprunt <> -1 then
               lTobEcheancier.InsertDB(Nil);
            For lInCpt2 := 1 to 12 do
               lTobEcheancier.PutValue('ECR_MONTANT' + IntToStr(lInCpt2),0);
            lInAncCodeEmprunt := lTobEcheance.GetValue('ECH_CODEEMPRUNT');
            lTobEcheancier.Putvalue('ECR_CODEEMPRUNT',lInAncCodeEmprunt);
            lTobEcheancier.PutValue('ECR_LIBELLE1',IntToStr(lTobEcheance.GetValue('EMP_CODEEMPRUNT')) + ' - ' + lTobEcheance.GetValue('EMP_LIBEMPRUNT'));
            lTobEcheancier.PutValue('ECR_LIBELLE2',VarToStr(lTobEcheance.GetValue('ANN_NOM1')));
            lTobEcheancier.PutValue('ECR_DEVISE',VarToStr(lTobEcheance.GetValue('EMP_DEVISE')));
         end;

         // Totalisation sur le bon mois
         Decodedate(lTobEcheance.GetValue('ECH_DATE'), lInAn, lInMois, lInJour);

           if lInMois < lInMoisDeBase then
              lInMois := 13 - (lInMoisDeBase - lInMois)
           else
              lInMois := lInMois - lInMoisDeBase + 1;
           lStNomChamps := 'ECR_MONTANT' + IntToStr(lInMois);

           lTobEcheancier.PutValue(lStNomChamps,lTobEcheancier.GetValue(lStNomChamps) + lTobEcheance.GetValue('ECH_VERSEMENT'));


      end;

      // Nécessaire
      if lInAncCodeEmprunt <> -1 then
         lTobEcheancier.InsertDB(Nil);

   finally
      lTobDonnees.Free;
      lTobEcheancier.Free;
   end;

   Result := true;

end;


// ----------------------------------------------------------------------------
// Nom    : ImprimePrevisionnel
// Date   : 13/09/2001
// Auteur : D. ZEDIAR
// Objet  : Impression du prévisionnel sur 5 ans
// ----------------------------------------------------------------------------
// Historique :
// ----------------------------------------------------------------------------
{
Procedure ImprimePrevisionnel;
begin
   if ExtractionDonneesPrevisionnel then
      AglLanceFiche('FP','FEMPIMP','','','PRE;');
end;
}

Function TOF_EMPIMP.ExtractionDonneesPrevisionnel : Boolean;
var
   lDtDatesDebEx1          : Array [0..5] Of TDateTime;
   lDtDatesFinEx1          : Array [0..5] Of TDateTime;

   lDtDatesDebEx           : Array [0..6] Of TDateTime;
   lDtDatesFinEx           : Array [0..6] Of TDateTime;
   lDtDateDeba             :TDateTime;
   lInAn, lInMois, lInJour : Word;
//   lInMois2, lInJour2      : Word;
   lDtDateRef              : TDateTime;
   lInCpt1, lInCpt2        : Integer;
   lTbEcheances            : Tob;
   lTbDetail               : Tob;
   lStSQL                  : String;
   lBoManquant             : Boolean;
   lRdProrata              : Double;
   lRdIntCourus            : Array [0..6] Of Double;
   lRdAssCourus            : Array [0..6] Of Double;
   lRdIntCCA               : Array [0..6] Of Double;
   lRdAssCCA               : Array [0..6] Of Double;
   lRdTotVers              : Array [0..6] Of Double;
   lRdTotAmo               : Array [0..6] Of Double;
   lRdTotInt               : Array [0..6] Of Double;
   lRdTotAss               : Array [0..6] Of Double;
   lRdSolde                : Array [0..6] Of Double;
   lRdDette                : Array [0..6] Of Double;
   lDtDateEch              : TDateTime;
   lRdIntEch               : Double;
   lRdAssEch               : Double;
   lRdSoldePrec            : Double;
   lDtDateEchPrec          : TDateTime;
   lRdIntEchPrec           : Double;
   lRdAssEchPrec           : Double;
   lInAncCodeEmprunt       : Integer;
   lRdTotVersEmp           : Double;
   lRdCumulVersEmp         : Double;
   lTbPrevisionnel         : Tob;
   lTbTableauFF            : Tob;
   lBoAncZoomOLE           : Boolean;
   nbExos                  : Integer;
begin
   Result := False;

   if Not V_PGI.OKOuvert then Exit ;

   nbExos:=6;

   // Saisie de la date de référence
   lDtDateRef := ChoixDateReference;
   if lDtDateRef <= 0 then Exit;

   ExecuteSQL('Delete From FECHEANCIER');

   // Détermine les dates des exercices N+1 à N+5
//   DecodeDate(VH^.Encours.Deb, lInAn, lInMois, lInJour);
//   DecodeDate(lDtDateRef, lInAn, lInMois2, lInJour2);

   For lInCpt1 := 0 to 5 do
   begin
//      lDtDatesDebEx[lInCpt1] := EncodeDate(lInAn + lInCpt1, lInMois, lInJour);
//      lDtDatesFinEx[lInCpt1] := EncodeDate(lInAn + lInCpt1 + 1, lInMois, lInJour);

      if not TObjetEmprunt.CalculerPeriode(PerExo,lDtDateRef,lInCpt1,lDtDatesDebEx1[lInCpt1],lDtDatesFinEx1[lInCpt1]) then
      begin
         PGIError('Un problème est survenu dans le calcul des dates '+fChainePeriode1+' !',TitreHalley);
         Exit;
      end;

      lRdIntCourus[lInCpt1]  := 0;
      lRdAssCourus[lInCpt1]  := 0;
      lRdIntCCA[lInCpt1]     := 0;
      lRdAssCCA[lInCpt1]     := 0;
      lRdTotAmo[lInCpt1]     := 0;
      lRdTotInt[lInCpt1]     := 0;
      lRdTotAss[lInCpt1]     := 0;
      lRdTotVers[lInCpt1]    := 0;
      lRdSolde[lInCpt1]      := 0;
      lRdDette[lInCpt1]      := 0;

      lDtDatesDebEx[lInCpt1+1] :=  lDtDatesDebEx1[lInCpt1];
      lDtDatesFinEx[lInCpt1+1] := lDtDatesFinEx1[lInCpt1];

   end;

   DecodeDate(lDtDatesDebEx1[0],lInAn,lInMois,lInJour);
   lDtDatesDebEx[0] := EncodeDate(lInAn - 1,lInMois,lInJour);
   DecodeDate(lDtDatesFinEx1[0],lInAn,lInMois,lInJour);
   lDtDatesFinEx[0] := EncodeDate(lInAn - 1,lInMois,lInJour);

   lTbEcheances := Tob.Create('',Nil,-1);
   lTbPrevisionnel := Tob.Create('',Nil,-1);
   try

      // Extraction des données
      lStSQL := 'Select * From FEMPRUNT, FECHEANCE';
      lStSQL := lStSQL + ' Where ECH_CODEEMPRUNT = EMP_CODEEMPRUNT';
      lStSQL := lStSQL + ' And EMP_CODEEMPRUNT In';
      lStSQL := lStSQL + ' (Select Distinct ECH_CODEEMPRUNT From FECHEANCE';
      lStSQL := lStSQL + ' Where ECH_DATE >= "' + UsDateTime(lDtDatesDebEx[0]) + '"';
      lStSQL := lStSQL + ' And ECH_DATE < "' + UsDateTime(lDtDatesFinEx[6]) + '")';
      If fInCodeEmprunt<>0 then  {FQ20641  08.06.07  YMO Choix d'un seul emprunt}
        lStSQL := lStSQL + ' AND EMP_CODEEMPRUNT = ' + IntToStr(fInCodeEmprunt);
      lStSQL := lStSQL + ' Order By ECH_CODEEMPRUNT, ECH_DATE';

      lTbEcheances.LoadDetailFromSQL(lStSQL);
      lDtDateEchPrec    := -1;
      lInAncCodeEmprunt := -1;
      lRdSoldePrec      := 0;
      lRdIntEchPrec     := 0;
      lRdAssEchPrec     := 0;
      lRdTotVersEmp     := 0;
      lRdCumulVersEmp   := 0;
      lBoManquant       := False;

      // Boucle sur les échéances
      For lInCpt1 := 0 to lTbEcheances.Detail.Count - 1 do
      begin
         lTbDetail := lTbEcheances.Detail[lInCpt1];

         // Rupture sur le code emprunt
         if lTbDetail.GetValue('EMP_CODEEMPRUNT') <> lInAncCodeEmprunt then
         begin
            lDtDateEchPrec    := -1;
            lRdSoldePrec      := 0;
            lRdIntEchPrec     := 0;
            lRdAssEchPrec     := 0;
            lRdTotVersEmp     := lTbDetail.GetValue('EMP_TOTVERS');
            lRdCumulVersEmp   := 0;
            lBoManquant       := lTbDetail.GetValue('EMP_DATECONTRAT') < lTbDetail.GetValue('EMP_DATEDEBUT');
            lInAncCodeEmprunt := lTbDetail.GetValue('EMP_CODEEMPRUNT');
         end;

         lDtDateEch      := lTbDetail.GetValue('ECH_DATE');
         lRdIntEch       := lTbDetail.GetValue('ECH_INTERET');
         lRdAssEch       := lTbDetail.GetValue('ECH_ASSURANCE');
         lRdCumulVersEmp := lRdCumulVersEmp + lTbDetail.GetValue('ECH_VERSEMENT');

         // Trouve à quel exercice appartient l'échéance
         For lInCpt2 := 0 to 6 do
         begin
            //
            if (lDtDateEch >= lDtDatesDebEx[lInCpt2]) And
               (lDtDateEch <= lDtDatesFinEx[lInCpt2]) Then   {FQ21744 06/11/2007  YMO}
            begin
               // Totalisations
               lRdTotAmo[lInCpt2]  := lRdTotAmo[lInCpt2]  + lTbDetail.GetValue('ECH_AMORTISSEMENT');
               lRdTotInt[lInCpt2]  := lRdTotInt[lInCpt2]  + lTbDetail.GetValue('ECH_INTERET');
               lRdTotAss[lInCpt2]  := lRdTotAss[lInCpt2]  + lTbDetail.GetValue('ECH_ASSURANCE');
               lRdTotVers[lInCpt2] := lRdTotVers[lInCpt2] + lTbDetail.GetValue('ECH_VERSEMENT');
            end;

            // Solde
            if lInCpt2 > 0 then
               if (lDtDatesFinEx[lInCpt2-1] > lDtDateEchPrec) and (lDtDatesFinEx[lInCpt2-1] <= lDtDateEch) then
               begin
                  lRdSolde[lInCpt2 - 1] := lRdSolde[lInCpt2 - 1] + lRdSoldePrec;
                  lRdDette[lInCpt2 - 1] := lRdDette[lInCpt2 - 1] + lRdTotVersEmp - lRdCumulVersEmp + lTbDetail.GetValue('ECH_VERSEMENT');
               end;


            // Charges réelles

            if lDtDateEchPrec > -1 then
            begin
            // Courus n
                if lBoManquant then
                   if (lDtDatesFinEx[lInCpt2]-1 > lDtDateEchPrec) and
                      (lDtDatesFinEx[lInCpt2] -1 <= lDtDateEch) then
                   begin

                      // Intérets manquants
    //                  lRdProrata := (lDtDatesFinEx[lInCpt2] - 1 - lDtDateEchPrec) / (lDtDateEch - lDtDateEchPrec);
                      lRdProrata := (lDtDatesFinEx[lInCpt2] - lDtDateEchPrec) / (lDtDateEch - lDtDateEchPrec);
    //Correction CAD du 07/12/04 : ne pas enlever un jour dans (Date de clôture - Date échéance précédente)
                      lRdIntCourus[lInCpt2] := lRdIntCourus[lInCpt2] + lRdProrata * lRdIntEch;
                      lRdAssCourus[lInCpt2] := lRdAssCourus[lInCpt2] + lRdProrata * lRdAssEch;

                      {FQ18447  18.06.07  YMO Suppression de la variation
                      if lInCpt2 < 5 then
                      begin
                         lRdIntCourus[lInCpt2 + 1] := lRdIntCourus[lInCpt2 + 1] - lRdProrata * lRdIntEch;
                         lRdAssCourus[lInCpt2 + 1] := lRdAssCourus[lInCpt2 + 1] - lRdProrata * lRdAssEch;
                      end; }
                   end;

                if Not lBoManquant then
                begin
                   if (lDtDatesFinEx[lInCpt2] > lDtDateEchPrec) and {FQ18447  18.06.07  YMO}
                      (lDtDatesFinEx[lInCpt2] <= lDtDateEch) then
                   begin
                      lRdProrata := (lDtDateEch - lDtDatesFinEx[lInCpt2]) / (lDtDateEch - lDtDateEchPrec);

                      {FQ18447  18.06.07  YMO Suppression de la variation
                      if lInCpt2 > 0 then
                      begin
                         lRdIntCCA[lInCpt2 - 1] := lRdIntCCA[lInCpt2 - 1] - lRdProrata * lRdIntEchPrec;
                         lRdAssCCA[lInCpt2 - 1] := lRdAssCCA[lInCpt2 - 1] - lRdProrata * lRdAssEchPrec;
                      end;}

                      lRdIntCCA[lInCpt2] := lRdIntCCA[lInCpt2] + lRdProrata * lRdIntEchPrec;
                      lRdAssCCA[lInCpt2] := lRdAssCCA[lInCpt2] + lRdProrata * lRdAssEchPrec;
                   end;
                end;
            end
            else
            begin
            //***********CAD 11/01/2005 Calcul intérêts courus premier exercice**************
            // Si date de mise à disposition des fonds < date de première échéance et
            // date de mise à disposition des fonds comprise dans l'exercice et
            // date de la première échéance non comprise dans l'exercice.
                if lBoManquant then
                begin
                    lDtDateEchPrec := lTbDetail.GetValue('EMP_DATECONTRAT');
                 // si la date de fin d'exercice est entre l'échéance précédente
                 // et l'échéance courante
                   if (lDtDatesDebEx[lInCpt2] < lDtDateEchPrec) and
                      (lDtDatesFinEx[lInCpt2] > lDtDateEchPrec) and
                      (lDtDatesFinEx[lInCpt2] < lDtDateEch) then
                   begin
                                    // Intérets manquants
                      lRdProrata := (lDtDatesFinEx[lInCpt2] - lDtDateEchPrec) / (lDtDateEch - lDtDateEchPrec);
                      lRdIntCourus[lInCpt2] := lRdIntCourus[lInCpt2] + lRdProrata * lRdIntEch;
                      lRdAssCourus[lInCpt2] := lRdAssCourus[lInCpt2] + lRdProrata * lRdAssEch;

                      if lInCpt2 < 5 then
                      begin
                         lRdIntCourus[lInCpt2 + 1] := lRdIntCourus[lInCpt2 + 1] - lRdProrata * lRdIntEch;
                         lRdAssCourus[lInCpt2 + 1] := lRdAssCourus[lInCpt2 + 1] - lRdProrata * lRdAssEch;
                      end;                    
                   end;
                end;
            //*************************************************************************
            end;
         end;

(* plus besoin car on calcule une échéance en plus ! CAD le 27/10/2003
         // Solde
         if (lDtDatesFinEx[4] > lDtDateEchPrec) and (lDtDatesFinEx[4] <= lDtDateEch) then
         begin
            showmessage('FInEx4 : ' + DateToStr(lDtDatesFinEx[4]) + '  EchPrec : ' +  DateToStr(lDtDateEchPrec) + chr(10) + chr(13) +
                        'DateEch : ' + DateToStr(lDtDateEch));
            lRdSolde[4] := lRdSolde[4] + lRdSoldePrec;
            lRdDette[4] := lRdDette[4] + lRdTotVersEmp - lRdCumulVersEmp + lTbDetail.GetValue('ECH_VERSEMENT');
         end;
                                           *)
         // Echéance suivante
         lDtDateEchPrec := lDtDateEch;
         lRdIntEchPrec  := lRdIntEch;
         lRdAssEchPrec  := lRdAssEch;
         lRdSoldePrec   := lTbDetail.GetValue('ECH_SOLDE');

      end;

      // Enregistrement en base
      for lInCpt1 := 0 to 11 do
         Tob.Create('FECHEANCIER',lTbPrevisionnel,-1).PutValue('ECR_CODEEMPRUNT',lInCpt1);

      // Titres des lignes
      lTbPrevisionnel.Detail[0].PutValue('ECR_LIBELLE1','Versements sur '+fChainePeriode2+' :');
      lTbPrevisionnel.Detail[1].PutValue('ECR_LIBELLE1','Capital remboursé sur '+fChainePeriode2+' :');

      lTbPrevisionnel.Detail[2].PutValue('ECR_LIBELLE1','Charges payées :');
      lTbPrevisionnel.Detail[2].PutValue('ECR_LIBELLE2','Intérêts');
      lTbPrevisionnel.Detail[3].PutValue('ECR_LIBELLE2','Assurance');

      lTbPrevisionnel.Detail[4].PutValue('ECR_LIBELLE1','Charges constatées d''avance :');
      lTbPrevisionnel.Detail[4].PutValue('ECR_LIBELLE2','Intérêts');
      lTbPrevisionnel.Detail[5].PutValue('ECR_LIBELLE2','Assurance');

      lTbPrevisionnel.Detail[6].PutValue('ECR_LIBELLE1','Charges à payer (intérêts courus) :');
      lTbPrevisionnel.Detail[6].PutValue('ECR_LIBELLE2','Intérêts');
      lTbPrevisionnel.Detail[7].PutValue('ECR_LIBELLE2','Assurance');

      lTbPrevisionnel.Detail[8].PutValue('ECR_LIBELLE1','Charges réelles de '+fChainePeriode2+' :');
      lTbPrevisionnel.Detail[8].PutValue('ECR_LIBELLE2','Intérêts');
      lTbPrevisionnel.Detail[9].PutValue('ECR_LIBELLE2','Assurance');

      lTbPrevisionnel.Detail[10].PutValue('ECR_LIBELLE1','Capital restant dû :');
      lTbPrevisionnel.Detail[11].PutValue('ECR_LIBELLE1','Annuités restant à verser :');

      lTbPrevisionnel.Detail[0].PutValue('ECR_DEVISE','1');
      lTbPrevisionnel.Detail[1].PutValue('ECR_DEVISE','2');
      lTbPrevisionnel.Detail[2].PutValue('ECR_DEVISE','3');
      lTbPrevisionnel.Detail[3].PutValue('ECR_DEVISE','3');
      lTbPrevisionnel.Detail[4].PutValue('ECR_DEVISE','4');
      lTbPrevisionnel.Detail[5].PutValue('ECR_DEVISE','4');
      lTbPrevisionnel.Detail[6].PutValue('ECR_DEVISE','5');
      lTbPrevisionnel.Detail[7].PutValue('ECR_DEVISE','5');
      lTbPrevisionnel.Detail[8].PutValue('ECR_DEVISE','6');
      lTbPrevisionnel.Detail[9].PutValue('ECR_DEVISE','6');
      lTbPrevisionnel.Detail[10].PutValue('ECR_DEVISE','7');
      lTbPrevisionnel.Detail[11].PutValue('ECR_DEVISE','8');

      // Init des montants des lignes
      For lInCpt1 := 1 to 6 do
      begin
       //  lTbPrevisionnel.Detail[0].PutValue('ECR_DATE'+IntToStr(lInCpt1+1),lDtDatesFinEx[lInCpt1]-1);
       //Correction CAD du 28/10/2003
         lTbPrevisionnel.Detail[0].PutValue('ECR_DATE'+IntToStr(lInCpt1),lDtDatesFinEx[lInCpt1]);
       //Fin Correction CAD du 28/10/2003
         lTbPrevisionnel.Detail[0].PutValue('ECR_MONTANT'+IntToStr(lInCpt1),lRdTotVers[lInCpt1]);
         lTbPrevisionnel.Detail[1].PutValue('ECR_MONTANT'+IntToStr(lInCpt1),lRdTotAmo[lInCpt1]);

         lTbPrevisionnel.Detail[2].PutValue('ECR_MONTANT'+IntToStr(lInCpt1),lRdTotInt[lInCpt1]);
         lTbPrevisionnel.Detail[3].PutValue('ECR_MONTANT'+IntToStr(lInCpt1),lRdTotAss[lInCpt1]);

         lTbPrevisionnel.Detail[4].PutValue('ECR_MONTANT'+IntToStr(lInCpt1),lRdIntCCA[lInCpt1]);
         lTbPrevisionnel.Detail[5].PutValue('ECR_MONTANT'+IntToStr(lInCpt1),lRdAssCCA[lInCpt1]);

         lTbPrevisionnel.Detail[6].PutValue('ECR_MONTANT'+IntToStr(lInCpt1),lRdIntCourus[lInCpt1]);
         lTbPrevisionnel.Detail[7].PutValue('ECR_MONTANT'+IntToStr(lInCpt1),lRdAssCourus[lInCpt1]);

         lTbPrevisionnel.Detail[8].PutValue('ECR_MONTANT'+IntToStr(lInCpt1),lRdTotInt[lInCpt1]+lRdIntCCA[lInCpt1]+lRdIntCourus[lInCpt1]);
         lTbPrevisionnel.Detail[9].PutValue('ECR_MONTANT'+IntToStr(lInCpt1),lRdTotAss[lInCpt1]+lRdAssCCA[lInCpt1]+lRdAssCourus[lInCpt1]);

         lTbPrevisionnel.Detail[10].PutValue('ECR_MONTANT'+IntToStr(lInCpt1),lRdSolde[lInCpt1]);
         lTbPrevisionnel.Detail[11].PutValue('ECR_MONTANT'+IntToStr(lInCpt1),lRdDette[lInCpt1]);
      end;

      lTbPrevisionnel.InsertDBByNivel(False);

   finally
      lTbPrevisionnel.Free;
      lTbEcheances.Free;
   end;

   Result := True;


end;

Function TOF_EMPIMP.ExtractionDonneesTableauFF : Boolean;
var
   lDtDatesDebEx           : Array [0..2] Of TDateTime;
   lDtDatesFinEx           : Array [0..2] Of TDateTime;
   lDtDateDeba             :TDateTime;
   lInAn, lInMois, lInJour : Word;
   lDtDateRef              : TDateTime;
   lEmprunt, lExo, lEche   : Integer;
   lTobEmprunt             : Tob;
   lTbEcheances            : Tob;
   lTbDetail, lTbDetailE   : Tob;
   lStSQL, lStSQLE         : String;
   lBoManquant             : Boolean;
   lRdProrata              : Double;
   lRdIntCourus            : Array [0..2] Of Double;
   lRdAssCourus            : Array [0..2] Of Double;
   lRdIntCCA               : Array [0..2] Of Double;
   lRdAssCCA               : Array [0..2] Of Double;
   lRdTotAmo               : Array [0..2] Of Double;
   lRdTotInt               : Array [0..2] Of Double;
   lRdTotAss               : Array [0..2] Of Double;
   lDtDateEch              : TDateTime;
   lRdIntEch               : Double;
   lRdAssEch               : Double;
   lDtDateEchPrec          : TDateTime;
   lRdIntEchPrec           : Double;
   lRdAssEchPrec           : Double;
   lInAncCodeEmprunt       : Integer;
   lRdCumulVersEmp         : Double;
   lTbTableauFF            : Tob;
   lBoAncZoomOLE           : Boolean;

   // Essaie de remplacement de fonction de calcul des périodes
   // pour FQ 20231 SBO 22/08/2007
   function _DeterminePeriode : boolean ;
   var lZExo   : TZExercice ;
       lExoRef : TExoDate ;
       lExoP   : TExoDate ;
       lExoS   : TExoDate ;
     begin
//       result := false ;

     lZExo   := ctxExercice ; //TZExercice.Create ;
     lExoRef := lZExo.QuelExoDate( lDtDateRef ) ;

     // Si pas d'exo ref on sort, on recherche un éventuelle exo précédent...
     if lExoRef.Code = '' then
       begin
       // 12 mois plus tôt...y'avait-il un exercice ? ( si on est en N+1 inexistant...)
       // C'est arbitraire mais vu qu'on est hors exo...
       lExoP := lZExo.QuelExoDate( PlusMois(lDtDateRef , -12) ) ;
       if lExoP.Code <> '' then
         begin
         lExoRef.Deb := lExoP.Fin + 1 ;
         lExoRef.Fin := PlusMois(lExoP.Fin, 12 ) ;
         end
       else
         begin
         // 24 mois plus tôt...y'avait-il un exercice ? ( si on est en N+2... avec N+1 inexistant...)
         // C'est arbitraire mais vu qu'on est hors exo...
         lExoP := lZExo.QuelExoDate( PlusMois(lDtDateRef , -24) ) ;
         if lExoP.Code <> '' then
           begin
           lExoRef.Deb := PlusMois(lExoP.Fin + 1, 12) ;
           lExoRef.Fin := PlusMois(lExoP.Fin, 24 ) ;
           end
         // Sinon on sort avec result à false
         else
           begin
           //exit ;
           lExoRef.Deb := PlusMois(lDtDateRef + 1, -12) ;
           lExoRef.Fin := lDtDateRef ;
           end

         end ;
       end ;

       // Affectation des date de la période de référence
       lDtDatesDebEx[1] := lExoRef.Deb ;
       lDtDatesFinEx[1] := lDtDateRef  ;

       // période précedente
       lDtDatesFinEx[0] := lDtDatesDebEx[1] - 1 ;
       lExoP := lZExo.QuelExoDate( lDtDatesFinEx[0] ) ;
       if lExoP.code <> ''
         then lDtDatesDebEx[0] := lExoP.Deb
         else lDtDatesDebEx[0] := PlusMois( lDtDatesDebEx[1], -12 ) ;

       // période suivante
       lDtDatesDebEx[2] := lExoRef.Fin + 1 ;
       lExoS := lZExo.QuelExoDate( lDtDatesDebEx[2] ) ;
       if lExoS.code <> ''
         then lDtDatesFinEx[2] := lExoS.Fin
         else lDtDatesFinEx[2] := PlusMois( lExoRef.Fin, 12 ) ;//PlusMois( lDtDatesFinEx[1], 12 ) ;

       result := True ;

     end ;
begin
   Result := False;
   lRdProrata:=0;

   if Not V_PGI.OKOuvert then Exit ;

   // Saisie de la date de référence
   lDtDateRef := ChoixDateReference;
   if lDtDateRef <= 0 then Exit;

   ExecuteSQL('DELETE FROM FECHEANCIER');

   DecodeDate(lDtDateRef,lInAn,lInMois,lInJour);
{
// Ancienne méthode
   For lExo := 0 to 2 do // On travaille sur 3 exercices lexo=0 : N-1 ; lexo=1 : N-1 ; lexo=2 : N+1
   begin

        if not TObjetEmprunt.CalculerPeriode(2,EncodeDate(lInAn- 1,lInMois,lInJour),lExo,lDtDatesDebEx[lExo],lDtDatesFinEx[lExo]) then
        begin
           PGIError('Un problème est survenu dans le calcul des dates '+fChainePeriode1+' !',TitreHalley);
           Exit;
        end;
   end;
}
   // Essaie de remplacement de fonction de calcul des périodes
   // pour FQ 20231 SBO 22/08/2007
   if not _DeterminePeriode then
     begin
     PGIError('La date de référence saisie ne permet pas le calcul de l''état !',TitreHalley);
     Exit;
     end ;

   lTobEmprunt := Tob.Create('',Nil,-1);
   lTbEcheances := Tob.Create('',Nil,-1);
   lTbTableauFF := Tob.Create('',Nil,-1);

   try

   // Emprunts à prendre en compte
   lStSQLE := 'Select EMP_CODEEMPRUNT, EMP_LIBEMPRUNT FROM FEMPRUNT';
   lStSQLE  := lStSQLE + ' WHERE EMP_CODEEMPRUNT IN';
   lStSQLE := lStSQLE + ' (SELECT DISTINCT ECH_CODEEMPRUNT FROM FECHEANCE';
   lStSQLE := lStSQLE + ' Where ECH_DATE >= "' + UsDateTime(lDtDatesDebEx[0]) + '"';
   lStSQLE := lStSQLE + ' And ECH_DATE < "' + UsDateTime(lDtDatesFinEx[2]) + '")';
   lStSQLE := lStSQLE + ' And EMP_STATUT<>4'; {19/03/07 Suppression des simulations}
   If fInCodeEmprunt<>0 then  {FQ20641  08.06.07  YMO Choix d'un seul emprunt}
      lStSQLE := lStSQLE + ' AND EMP_CODEEMPRUNT = ' + IntToStr(fInCodeEmprunt);

   lStSQLE := lStSQLE + ' ORDER BY EMP_CODEEMPRUNT';

   lTobEmprunt.LoadDetailFromSQL(lStSQLE);

     // Boucle sur les emprunts
     For lEmprunt := 0 to lTobEmprunt.Detail.Count-1 do
     begin

       For lExo := 0 to 2 do
       begin
          lRdIntCourus[lExo]  := 0;
          lRdAssCourus[lExo]  := 0;
          lRdIntCCA[lExo]     := 0;
          lRdAssCCA[lExo]     := 0;
          lRdTotAmo[lExo]     := 0;
          lRdTotInt[lExo]     := 0;
          lRdTotAss[lExo]     := 0;
       end;

      lTbDetailE := lTobEmprunt.Detail[lEmprunt];

      // Extraction des données de l'emprunt
      lStSQL := 'SELECT * FROM FEMPRUNT, FECHEANCE';
      lStSQL := lStSQL + ' WHERE ECH_CODEEMPRUNT = EMP_CODEEMPRUNT';
      lStSQL := lStSQL + ' AND EMP_CODEEMPRUNT ="'+lTbDetailE.GetString('EMP_CODEEMPRUNT')+'"';
      lStSQL := lStSQL + ' AND ECH_DATE >= "' + UsDateTime(lDtDatesDebEx[0]) + '"';
      lStSQL := lStSQL + ' AND ECH_DATE < "' + UsDateTime(lDtDatesFinEx[2]) + '"';
      lStSQL := lStSQL + ' ORDER BY ECH_CODEEMPRUNT, ECH_DATE';

      lTbEcheances.LoadDetailFromSQL(lStSQL);
      lDtDateEchPrec    := -1;
      lInAncCodeEmprunt := -1;
      lRdIntEchPrec     := 0;
      lRdAssEchPrec     := 0;
      lRdCumulVersEmp   := 0;
      lBoManquant       := False;

      // Boucle sur les échéances
      For lEche := 0 to lTbEcheances.Detail.Count - 1 do
      begin
         lTbDetail := lTbEcheances.Detail[lEche];

         // Rupture sur le code emprunt
         if lTbDetail.GetValue('EMP_CODEEMPRUNT') <> lInAncCodeEmprunt then
         begin
            lDtDateEchPrec    := -1;
            lRdIntEchPrec     := 0;
            lRdAssEchPrec     := 0;
            lRdCumulVersEmp   := 0;
            lBoManquant       := lTbDetail.GetValue('EMP_DATECONTRAT') < lTbDetail.GetValue('EMP_DATEDEBUT');
            lInAncCodeEmprunt := lTbDetail.GetValue('EMP_CODEEMPRUNT');
         end;

         lDtDateEch      := lTbDetail.GetValue('ECH_DATE');
         lRdIntEch       := lTbDetail.GetValue('ECH_INTERET');
         lRdAssEch       := lTbDetail.GetValue('ECH_ASSURANCE');
         lRdCumulVersEmp := lRdCumulVersEmp + lTbDetail.GetValue('ECH_VERSEMENT');

         // On travaille sur 3 exercices lexo=0 : N-1 ; lexo=1 : N ; lexo=2 : N+1
         For lexo := 0 to 2 do
         begin

            if (lDtDateEch >= lDtDatesDebEx[lExo]) And
               (lDtDateEch < lDtDatesFinEx[lExo]) Then
            begin
               // Totalisations
               lRdTotAmo[lExo]  := RoundTo(lRdTotAmo[lExo]+ lTbDetail.GetValue('ECH_AMORTISSEMENT'), -2);
               lRdTotInt[lExo]  := RoundTo(lRdTotInt[lExo]+ lTbDetail.GetValue('ECH_INTERET'), -2);
               lRdTotAss[lExo]  := RoundTo(lRdTotAss[lExo]+ lTbDetail.GetValue('ECH_ASSURANCE'), -2);
            end;

            // Charges réelles
            if lDtDateEchPrec > -1 then
            begin
            // Courus n
                if lBoManquant then    {11.07.07 FQ20231+FQ20929 YMO Optimisation dates de debut et fin prises pour calcul IC et CCA}
                   if ((lExo=0) and (lDtDatesDebEx[lExo+1]- 2> lDtDateEchPrec) and (lDtDatesDebEx[lExo+1]- 2<= lDtDateEch))
                   or ((lExo=1) and (lDtDatesFinEx[lExo]- 1> lDtDateEchPrec) and (lDtDatesFinEx[lExo]- 1<= lDtDateEch)) then
                   begin
                      lRdProrata:= 0;
                      {Intérets manquants}
                      if lexo=0 then lRdProrata := (lDtDatesDebEx[lExo+1]-1-lDtDateEchPrec)/ (lDtDateEch-lDtDateEchPrec)
                      else
                      if lexo=1 then lRdProrata := (lDtDatesFinEx[lExo]-lDtDateEchPrec)/ (lDtDateEch-lDtDateEchPrec);

                      lRdIntCourus[lExo]:= RoundTo(lRdIntCourus[lExo]+ lRdProrata* lRdIntEch, -2);
                      lRdAssCourus[lExo]:= RoundTo(lRdAssCourus[lExo]+ lRdProrata* lRdAssEch, -2);
                   end;

                if Not lBoManquant then
                begin
                   {FQ18684 FQ20231 05.06.07  YMO Choix}
                   {11.07.07 FQ20231+FQ20929 YMO Optimisation dates de debut et fin prises pour calcul IC et CCA}
                   if  ((lExo=1) and (lDtDatesDebEx[lExo]-1> lDtDateEchPrec) and (lDtDatesDebEx[lExo]-1<= lDtDateEch))
                   or  ((lExo=2) and (lDtDatesFinEx[lExo-1]> lDtDateEchPrec) and (lDtDatesFinEx[lExo-1]<= lDtDateEch)) then
                   begin
                      {Charges calc. d'avance}
                      If lexo=1 then lRdProrata:= (lDtDateEch- lDtDatesDebEx[lExo]+1)/ (lDtDateEch- lDtDateEchPrec)
                      else
                      if lexo=2 then lRdProrata:= (lDtDateEch- lDtDatesFinEx[lExo-1])/ (lDtDateEch- lDtDateEchPrec);

                      lRdIntCCA[lExo]:= RoundTo(lRdIntCCA[lExo]+ lRdProrata* lRdIntEchPrec, -2);
                      lRdAssCCA[lExo]:= RoundTo(lRdAssCCA[lExo]+ lRdProrata* lRdAssEchPrec, -2);
                   end;
                end;
            end
            else
            begin
            //***********CAD 11/01/2005 Calcul intérêts courus premier exercice**************
            // Si date de mise à disposition des fonds < date de première échéance et
            // date de mise à disposition des fonds comprise dans l'exercice et
            // date de la première échéance non comprise dans l'exercice.
                if lBoManquant then
                begin
                    lDtDateEchPrec := lTbDetail.GetValue('EMP_DATECONTRAT');
                 // si la date de fin d'exercice est entre l'échéance précédente
                 // et l'échéance courante
                   if (lDtDatesDebEx[lExo]< lDtDateEchPrec) and
                      (lDtDatesFinEx[lExo]> lDtDateEchPrec) and
                      (lDtDatesFinEx[lExo]< lDtDateEch) then
                   begin
                      // Intérets manquants
                      lRdProrata := RoundTo((lDtDatesFinEx[lExo]- lDtDateEchPrec)/ (lDtDateEch- lDtDateEchPrec), -2);

                      lRdIntCourus[lExo]:= RoundTo(lRdIntCourus[lExo]+ lRdProrata* lRdIntEch, -2);
                      lRdAssCourus[lExo]:= RoundTo(lRdAssCourus[lExo]+ lRdProrata* lRdAssEch, -2);

                   end;
                end;
            //*************************************************************************
            end;
         end;

         // Echéance suivante
         lDtDateEchPrec := lDtDateEch;
         lRdIntEchPrec  := lRdIntEch;
         lRdAssEchPrec  := lRdAssEch;

      end;

      // Enregistrement en base : 3 lignes (intérêts - assurance - total)
      Tob.Create('FECHEANCIER',lTbTableauFF,-1).PutValue('ECR_CODEEMPRUNT',lTbDetailE.GetString('EMP_CODEEMPRUNT'));
      Tob.Create('FECHEANCIER',lTbTableauFF,-1).PutValue('ECR_CODEEMPRUNT',lTbDetailE.GetString('EMP_CODEEMPRUNT'));

      // Titres des lignes
      lTbTableauFF.Detail[2*lEmprunt].PutValue('ECR_LIBELLE1', lTbDetailE.GetString('EMP_LIBEMPRUNT'));
      lTbTableauFF.Detail[2*lEmprunt].PutValue('ECR_LIBELLE2', 'Frais financiers');
      lTbTableauFF.Detail[2*lEmprunt+1].PutValue('ECR_LIBELLE2', 'Assurance');

      lTbTableauFF.Detail[2*lEmprunt].PutValue('ECR_DATE1', lDtDatesDebEx[1]);
      lTbTableauFF.Detail[2*lEmprunt+1].PutValue('ECR_DATE1', lDtDatesDebEx[1]);
      lTbTableauFF.Detail[2*lEmprunt].PutValue('ECR_DATE2', lDtDatesFinEx[1]);
      lTbTableauFF.Detail[2*lEmprunt+1].PutValue('ECR_DATE2', lDtDatesFinEx[1]);

      // Init des montants des lignes

      //Int.+Ass. courus M-1  {FQ20457  31.05.07 YMO Arrondis à 2 décimales}
      lTbTableauFF.Detail[2*lEmprunt].PutValue('ECR_MONTANT1', RoundTo(lRdIntCourus[0], -2));
      lTbTableauFF.Detail[2*lEmprunt+1].PutValue('ECR_MONTANT1', RoundTo(lRdAssCourus[0], -2));
      //Int.+Ass. avance M
      lTbTableauFF.Detail[2*lEmprunt].PutValue('ECR_MONTANT2', RoundTo(lRdIntCCA[1], -2));
      lTbTableauFF.Detail[2*lEmprunt+1].PutValue('ECR_MONTANT2', RoundTo(lRdAssCCA[1], -2));
      //Int.+Ass. Charges payées M
      lTbTableauFF.Detail[2*lEmprunt].PutValue('ECR_MONTANT3', RoundTo(lRdTotInt[1], -2));
      lTbTableauFF.Detail[2*lEmprunt+1].PutValue('ECR_MONTANT3', RoundTo(lRdTotAss[1], -2));
      //Int.+Ass. courus M
      lTbTableauFF.Detail[2*lEmprunt].PutValue('ECR_MONTANT4', RoundTo(lRdIntCourus[1], -2));
      lTbTableauFF.Detail[2*lEmprunt+1].PutValue('ECR_MONTANT4', RoundTo(lRdAssCourus[1], -2));
      //Int.+Ass. avance M+1
      lTbTableauFF.Detail[2*lEmprunt].PutValue('ECR_MONTANT5', RoundTo(lRdIntCCA[2], -2));
      lTbTableauFF.Detail[2*lEmprunt+1].PutValue('ECR_MONTANT5', RoundTo(lRdAssCCA[2], -2));
      //Int.+Ass. Calcul charges réelles
      lTbTableauFF.Detail[2*lEmprunt].PutValue('ECR_MONTANT6', RoundTo(lRdTotInt[1]+lRdIntCCA[1]-lRdIntCourus[0]+lRdIntCourus[1]-lRdIntCCA[2], -2));
      lTbTableauFF.Detail[2*lEmprunt+1].PutValue('ECR_MONTANT6', RoundTo(lRdTotAss[1]+lRdAssCCA[1]-lRdAssCourus[0]+lRdAssCourus[1]-lRdAssCCA[2], -2)); {27.06.07 Correction signes}

   end;

   lTbTableauFF.InsertDBByNivel(False);

   finally
      lTbTableauFF.Free;
      lTbEcheances.Free;
      lTobEmprunt.Free;
   end;

   Result := True;


end;


// ----------------------------------------------------------------------------
// Nom    : OnArgument
// Date   : 13/09/2003
// Auteur : D. ZEDIAR
// Objet  :
// ----------------------------------------------------------------------------
// Historique :
// ----------------------------------------------------------------------------
{Procedure ImprimeTableauEmprunts;
begin
   if ExtractionDonneesTableauEmprunt then
      AglLanceFiche('FP','FEMPIMP','','','TEM;');
end;
}

// ----------------------------------------------------------------------------
// Nom    : ExtractionDonneesTableauEmprunt
// Date   : 13/06/2003
// Auteur : D. ZEDIAR
// Objet  :
// ----------------------------------------------------------------------------
// Historique :
// ----------------------------------------------------------------------------
Function  TOF_EMPIMP.ExtractionDonneesTableauEmprunt : Boolean;
var
   lDtDateRef   : TDateTime;
   lDtDate1     : TDateTime;
   lDtDate2     : TDateTime;
   lDtDateDebN0 : TDateTime;
   lDtDateDebN1 : TDateTime;
   lDtDateDebN2 : TDateTime;
   lDtDateDebN6 : TDateTime;
   lDtDateFinN0 : TDateTime;
   lDtDateFinN1 : TDateTime;
   lDtDateFinN2 : TDateTime;
   lDtDateFinN6 : TDateTime;

//   lInAn    : Word;
//   lInMois  : Word;
//   lInJour  : Word;
//   lInMois2 : Word;
//   lInJour2 : Word;

   lTbEmprunts       : TOB;
   lTbEcheance       : TOB;
   lTbEcheancier     : TOB;
   lInCpt            : Integer;
   lStSQL, lStCode   : String;
   lInAncCodeEmprunt : Integer;
   lDtDateEch        : TDateTime;

   lBoManquant    : Boolean;
//   lRdProrata     : Double;
   lRdIntEchPrec  : Double;
   lRdAssEchPrec  : Double;
   lDtDateEchPrec : TDateTime;
   lRdIntCourus   : Double;
   lRdAssCourus   : Double;
   lRdIntCCA      : Double;
   lRdAssCCA      : Double;
   lRdInterets    : Double;

   lStNumCompte  : String;
   lDtDateDebut,lDateContrat, lDateContratFixe : TDateTime;
   lStOrgPreteur : String;
   lRdCapital    : Double;
   lInDuree      : Integer;
   lRdTauxAn     : Double;
   lStDevise     : String;

   lRdCapitalAvantExe : Double;  // Capital restant du avant l'exercice
   lRdIntExe          : Double;  // Intérets sur l'exercice
   lRdAssExe          : Double;  // Intérets sur l'exercice
   lRdRembAvantExe    : Double;  // Remboursement avant exercice
   lRdRembExe         : Double;  // Remboursement sur l'exercice
   lRdCapitalApresExe : Double;  // Capital restant du après l'exercice
   lRdRembExe1        : Double;  // Remboursement sur N+1
   lRdRembExe1a5      : Double;  // Remboursement sur N+2 à N+5
   lRdRembPlus5       : Double;  // Au delà de 5 ans

   function Neant(pStValeur : String) : string;
   begin
      if pStValeur = '' then
         Result := '-'
      else
         Result := pStValeur;
   end;

   Procedure RuptureCodeEmprunt(pBoForcerRupture : Boolean = False);
   begin
      if pBoForcerRupture or
         (lInAncCodeEmprunt <> lTbEcheance.GetValue('EMP_CODEEMPRUNT')) then
      begin
         if lInAncCodeEmprunt <> -1 then
         begin
            // Insersion d'une ligne dans la table FECHEANCIER
            lTbEcheancier.PutValue('ECR_CODEEMPRUNT',lInAncCodeEmprunt);  // code de l'emprunt

            lTbEcheancier.PutValue('ECR_DATE3',    lDtDate1);   // Exercice
            lTbEcheancier.PutValue('ECR_DATE4',    lDtDate2);   // Exercice
            lTbEcheancier.PutValue('ECR_DATE5',    lDateContratFixe);  {FQ19987 20/04/07 YMO}

            lTbEcheancier.PutValue('ECR_LIBELLE1', lStNumCompte);   // Numéro de compte
            lTbEcheancier.PutValue('ECR_DATE1',    lDtDateDebut);   // 1ière échéance
            lTbEcheancier.PutValue('ECR_LIBELLE2', lStOrgPreteur);  // Organisme prêteur
            lTbEcheancier.PutValue('ECR_MONTANT1', lRdCapital);     // Capital
            lTbEcheancier.PutValue('ECR_MONTANT2', lInDuree);       // Durée en mois
            lTbEcheancier.PutValue('ECR_MONTANT3', lRdTauxAn);      // Taux annuel
            lTbEcheancier.PutValue('ECR_DEVISE',   lStDevise);      // devise de l'emprunt
            lTbEcheancier.PutValue('ECR_DATE2',    lDtDateEch);     // Dernière échéance

            lRdInterets := lRdIntExe + lRdAssExe + lRdIntCourus + lRdAssCourus + lRdIntCCA + lRdAssCCA;
            lTbEcheancier.PutValue('ECR_MONTANT4' ,lRdCapitalAvantExe);    // Capital restant du avant l'exercice
            lTbEcheancier.PutValue('ECR_MONTANT5' ,lRdInterets);           // Intérets sur l'exercice
            lTbEcheancier.PutValue('ECR_MONTANT6' ,lRdRembAvantExe);       // Remboursement avant exercice
            lTbEcheancier.PutValue('ECR_MONTANT7' ,lRdRembExe);            // Remboursement sur l'exercice
            lTbEcheancier.PutValue('ECR_MONTANT8' ,lRdCapitalApresExe);    // Capital restant du après l'exercice
            lTbEcheancier.PutValue('ECR_MONTANT9' ,lRdRembExe1);           // Remboursement sur N+1
            lTbEcheancier.PutValue('ECR_MONTANT10',lRdRembExe1a5);         // Remboursement sur N+2 à N+5
            lTbEcheancier.PutValue('ECR_MONTANT11',lRdRembPlus5);          //

            lTbEcheancier.Modifie := True;
            lTbEcheancier.InsertDB(Nil);
         end;

         lStNumCompte  := Neant(lTbEcheance.GetValue('EMP_LIBEMPRUNT'));
         lDtDateDebut  := lTbEcheance.GetValue('EMP_DATEDEBUT');
         //Ajout CAD du 16/01/2004
         lDateContrat := lTbEcheance.GetValue('EMP_DATECONTRAT');
         //FIn ajout CAD du 16/01/2004
         lStOrgPreteur := Neant(VarToStr(lTbEcheance.GetValue('ANN_NOMPER')));
         lRdCapital    := lTbEcheance.GetValue('EMP_CAPITAL');
         lInDuree      := lTbEcheance.GetValue('EMP_DUREE');
         lRdTauxAn     := lTbEcheance.GetValue('EMP_TAUXAN');
         lStDevise     := lTbEcheance.GetValue('EMP_DEVISE');
         lDateContratFixe := lTbEcheance.GetValue('EMP_DATECONTRAT');

         lInAncCodeEmprunt  := lTbEcheance.GetValue('EMP_CODEEMPRUNT');
         lRdCapitalAvantExe := 0;
         lRdIntExe          := 0;
         lRdAssExe          := 0;
         lRdIntEchPrec      := 0;
         lRdAssEchPrec      := 0;
         lRdIntCourus       := 0;
         lRdAssCourus       := 0;
         lRdIntCCA          := 0;
         lRdAssCCA          := 0;
         lDtDateEchPrec     := 0;
         lRdRembAvantExe    := 0;
         lRdRembExe         := 0;
         lRdCapitalApresExe := 0;
         lRdRembExe1        := 0;
         lRdRembExe1a5      := 0;
         lRdRembPlus5       := 0;

         lDtDateEchPrec := -1;
         lBoManquant    := lTbEcheance.GetValue('EMP_DATECONTRAT') < lTbEcheance.GetValue('EMP_DATEDEBUT');
      end;
   end;

begin
   Result := False;

   if Not V_PGI.OKOuvert then Exit ;

   // Saisie de la date de référence
   lDtDateRef := ChoixDateReference;
   if lDtDateRef <= 0 then Exit;

   ExecuteSQL('Delete From FECHEANCIER');

   // Détermine les dates des exercices N et N+1
//   DecodeDate(VH^.Encours.Deb, lInAn, lInMois, lInJour);
//   DecodeDate(lDtDateRef, lInAn, lInMois2, lInJour2);
   if (not TObjetEmprunt.CalculerPeriode(2,lDtDateRef,0,lDtDate1,lDtDate2)) or
      (not TObjetEmprunt.CalculerPeriode(1,lDtDateRef,1,lDtDateDebN1,lDtDateFinN1)) or
      (not TObjetEmprunt.CalculerPeriode(1,lDtDateRef,2,lDtDateDebN2,lDtDateFinN2)) or
      (not TObjetEmprunt.CalculerPeriode(1,lDtDateRef,6,lDtDateDebN6,lDtDateFinN6)) then
   begin
      PGIError('Un problème est survenu dans le calcul des dates '+fChainePeriode1+' !',TitreHalley);
      Exit;
   end;

   lDtDateDebN0 := lDtDate1;
   lDtDateFinN0 := lDtDate2;

//   lDtDateDebN0      := EncodeDate(lInAn,lInMois,lInJour);
//   lDtDateDebN1      := EncodeDate(lInAn+1,lInMois,lInJour);
//   lDtDateDebN2      := EncodeDate(lInAn+2,lInMois,lInJour);
//   lDtDateDebN6      := EncodeDate(lInAn+6,lInMois,lInJour);

   //
   lTbEcheancier := TOB.Create('FECHEANCIER',Nil,-1);
   lTbEmprunts := TOB.Create('',Nil,-1);

   If fInCodeEmprunt>0 then
      lStCode := ' AND ECH_CODEEMPRUNT = ' + IntToStr(fInCodeEmprunt)
   else
      lStCode:='';

   try

      lStSQL := ' Select EMP_CODEEMPRUNT,ECH_DATE,ECH_SOLDE,ECH_INTERET,ECH_AMORTISSEMENT,ECH_CUMULA,' +
                ' EMP_LIBEMPRUNT, EMP_NUMCOMPTE, EMP_DATEDEBUT, EMP_ORGPRETEUR, EMP_GUIDORGPRETEUR, EMP_CAPITAL, EMP_DUREE,' +
                ' EMP_TAUXAN, EMP_DEVISE, ANN_NOMPER, EMP_DATECONTRAT, ECH_ASSURANCE' +
                ' From FECHEANCE, FEMPRUNT' +
                ' Left Outer Join ANNUAIRE On ANN_GUIDPER = EMP_GUIDORGPRETEUR' +
                ' Where EMP_STATUT<>4' + {19/03/07 Suppression des simulations}
                ' AND EMP_DATEDEBUT<= "' + USDateTime(lDtDateFinN0) + '"' +   //06.07 YMO On ne prend que les emprunts deja démarrés..
                ' And ECH_CODEEMPRUNT = EMP_CODEEMPRUNT' + lStCode +
                ' Order By EMP_CODEEMPRUNT, ECH_DATE';

      lTbEmprunts.LoadDetailFromSQL(lStSQL);

      lInAncCodeEmprunt := -1;

      // Pour éviter les warnings
      lStNumCompte        := '';
      lDtDateDebut        := 0;
      lDateContrat:=0;
      lStOrgPreteur       := '';
      lRdCapital          := 0;
      lInDuree            := 0;
      lRdTauxAn           := 0;
      lStDevise           := '';
      lDtDateEch          := 0;
      lRdCapitalAvantExe  := 0;
      lRdIntExe           := 0;
      lRdRembAvantExe     := 0;
      lRdRembExe          := 0;
      lRdCapitalApresExe  := 0;
      lRdRembExe1         := 0;
      lRdRembExe1a5       := 0;
      lRdRembPlus5        := 0;
      lDtDateEchPrec      := -1;

      // Boucle sur les échéances des emprunts
      For lInCpt := 0 to lTbEmprunts.Detail.Count - 1 do
      begin
         lTbEcheance := lTbEmprunts.Detail[lInCpt];

         RuptureCodeEmprunt;

         // Rupture sur le code emprunt
         lDtDateEch := lTbEcheance.GetValue('ECH_DATE');

         // Avant Exercice
         if lDtDateEch < lDtDateDebN0 then
         begin
            //S'il existe au moins une échéance antérieure à l'exercice de départ du tableau,
            //LDateContrat est remis à 0.
            lDateContrat:=0;
            // Capital restant dû avant exercice
            lRdCapitalAvantExe := lTbEcheance.GetValue('ECH_SOLDE');
            // Remboursement avant exercice
            lRdRembAvantExe := lTbEcheance.GetValue('ECH_CUMULA');
         end;

         if (lDateContrat <> 0)and (lDtDateEch >= lDtDateDebN0) then
         begin
             //Si premier exercice et date de conclusion du contrat antérieure
             // à la date de début d'exercice : le montant de l'emprunt dû en début d'exercice
             // doit être renseigné : lDateContrat
             if lDateContrat < lDtDateDebN0 then
             begin
                // Capital restant dû avant exercice
                lRdCapitalAvantExe := lTbEcheance.GetValue('EMP_CAPITAL');
             end;
         end;
         // Sur l'exercice
         if (lDtDateEch >= lDtDateDebN0) and (lDtDateEch <= lDtDateFinN0) then
         begin
            // Intérets sur l'exercice
            lRdIntExe := lRdIntExe + lTbEcheance.GetValue('ECH_INTERET');
            lRdAssExe := lRdAssExe + lTbEcheance.GetValue('ECH_ASSURANCE');

            // Remboursement sur l'exercice
            lRdRembExe := lRdRembExe + lTbEcheance.GetValue('ECH_AMORTISSEMENT');
            // Capital après exercice
            lRdCapitalApresExe := lTbEcheance.GetValue('ECH_SOLDE');
         end;
            //Si date contrat < début exercice suivant et pas de remboursement sur l'exercice
            //et l'emprunt n'est pas totalement remboursé,
            // le capital restant dû en fin d'exercice = capital de l'emprunt.
         if (lDateContrat<>0) and (lRdCapitalApresExe=0) and (lDateContrat <= lDtDateFinN0) then
         lRdCapitalApresExe := lTbEcheance.GetValue('EMP_CAPITAL');


{         // Charges réelles
         if lBoManquant then
         begin
            // Intérêts courus
            if (lDtDateDebN1 - 1 >  lDtDateEchPrec) and
               (lDtDateDebN1 - 1 <= lDtDateEch) then
            begin
               // Intérets manquants
               lRdProrata   := (lDtDateDebN1 - 1 - lDtDateEchPrec) / (lDtDateEch - lDtDateEchPrec);
               lRdIntCourus := lRdIntCourus + lRdProrata * lTbEcheance.GetValue('ECH_INTERET');
               lRdAssCourus := lRdAssCourus + lRdProrata * lTbEcheance.GetValue('ECH_ASSURANCE');
            end;

            if (lDtDateDebN0 - 1 >  lDtDateEchPrec) and
               (lDtDateDebN0 - 1 <= lDtDateEch) then
            begin
               lRdProrata   := (lDtDateDebN0 - 1 - lDtDateEchPrec) / (lDtDateEch - lDtDateEchPrec);
               lRdIntCourus := lRdIntCourus - lRdProrata * lTbEcheance.GetValue('ECH_INTERET');
               lRdAssCourus := lRdAssCourus - lRdProrata * lTbEcheance.GetValue('ECH_ASSURANCE');
            end;

         end
         else
         begin
            // Charges constatées d'avance
            if (lDtDateDebN0 - 1 > lDtDateEchPrec) and
               (lDtDateDebN0 - 1 <= lDtDateEch) then
            begin
               lRdProrata := (lDtDateEch - lDtDateDebN0 + 1) / (lDtDateEch - lDtDateEchPrec);
               lRdIntCCA := lRdIntCCA + lRdProrata * lRdIntEchPrec;
               lRdAssCCA := lRdAssCCA + lRdProrata * lRdAssEchPrec;
            end;

            if (lDtDateDebN1 - 1 > lDtDateEchPrec) and
               (lDtDateDebN1 - 1 <= lDtDateEch) then
            begin
               lRdProrata := (lDtDateEch - lDtDateDebN1 + 1) / (lDtDateEch - lDtDateEchPrec);
               lRdIntCCA := lRdIntCCA - lRdProrata * lRdIntEchPrec;
               lRdAssCCA := lRdAssCCA - lRdProrata * lRdAssEchPrec;
            end;
         end;
}
         // N+1  {22.05.07 YMO FQ20224}
         if (lDtDateEch >= lDtDateDebN1) and (lDtDateEch <= lDtDateFinN1) then
         begin
            // Remboursement sur N+1
            lRdRembExe1 := lRdRembExe1 + lTbEcheance.GetValue('ECH_AMORTISSEMENT');
//            lRdRembExe1 := lTbEcheance.GetValue('ECH_SOLDE');
         end;

         // N+2 à N+5
         if (lDtDateEch >= lDtDateDebN2) and (lDtDateEch < lDtDateDebN6) then
         begin
            // Remboursement de N+1 à N+5
            lRdRembExe1a5 := lRdRembExe1a5 + lTbEcheance.GetValue('ECH_AMORTISSEMENT');
//            lRdRembExe1a5 := lTbEcheance.GetValue('ECH_SOLDE');
         end;

         // > N+5
         if lDtDateEch >= lDtDateDebN6 then
         begin
            // Remboursement après 5 ans
            lRdRembPlus5 := lRdRembPlus5 + lTbEcheance.GetValue('ECH_AMORTISSEMENT');
//            lRdRembPlus5 := lTbEcheance.GetValue('ECH_SOLDE');
         end;

         lDtDateEchPrec := lDtDateEch;
         lRdIntEchPrec  := lTbEcheance.GetValue('ECH_INTERET');
         lRdAssEchPrec  := lTbEcheance.GetValue('ECH_ASSURANCE');
      end;

      if lTbEmprunts.Detail.Count > 0 then
         RuptureCodeEmprunt(True);

   finally
      lTbEcheancier.Free;
      lTbEmprunts.Free;
   end;


   Result := True;
end;


// ----------------------------------------------------------------------------
// Nom    : ChoixEmpruntClick
// Date   : 13/09/2002
// Auteur : D. ZEDIAR
// Objet  : Sélection d'un emprunt
// ----------------------------------------------------------------------------
// Historique :
// ----------------------------------------------------------------------------
{Function TObjetEmprunt.CalculerExercice(pDtDateRef : TDateTime; pInN : Integer; Var pDtDateDebExe : TDateTime; Var pDtDateFinExe : TDateTime) : Boolean;
var
   lObTobExe  : Tob;
begin
   Result := False;

   lObTobExe := Tob.Create('',Nil,-1);

   try
      lObTobExe.LoadDetailFromSQL('Select * From EXERCICE Where EX_DATEFIN >= "' + USDATETIME(pDtDateRef) + '" Order By EX_DATEDEBUT');

      if lObTobExe.Detail.Count = 0 then
      begin
         pDtDateDebExe := 0;
         pDtDateFinExe := 0;
         PGIError('La date de référence ne permet pas le calcul des dates '+fChainePeriode1+' !', TitreHalley);
         Exit;
      end
      else
         if lObTobExe.Detail.Count <= pInN then
         begin
            pDtDateDebExe := lObTobExe.Detail[lObTobExe.Detail.Count - 1].GetValue('EX_DATEFIN') + 1;
            pDtDateDebExe := PLUSMOIS(pDtDateDebExe,12*(pInN-lObTobExe.Detail.Count));
            pDtDateFinExe := PLUSMOIS(pDtDateDebExe,12);
         end
         else
         begin
            pDtDateDebExe := lObTobExe.Detail[pInN].GetValue('EX_DATEDEBUT');
            pDtDateFinExe := lObTobExe.Detail[pInN].GetValue('EX_DATEFIN');
         end;

      Result := True;

   finally
      lObTobExe.Free;
   end;

end;
}

// ----------------------------------------------------------------------------
// Nom    : OnArgument
// Date   : 13/09/2002
// Auteur : D. ZEDIAR
// Objet  : Fenêtre d'impression
// ----------------------------------------------------------------------------
// Historique :
// ----------------------------------------------------------------------------
procedure TOF_EMPIMP.OnArgument (S : String ) ;
var
  lStEtat    : String;
  lStEmprunt : String;
  FBoRegroupement : Boolean;

  function NomCabinet : String;
  var
     lObTob : TOB;
  begin
     lObTob := TOB.Create('',Nil,-1);
     Result :=  '';
     try
                                                                                                           //CAD 040723Copy(V_PGI.DBSocName)
        lObTob.LoadDetailFromSQL('Select DOS_SOCIETE, DOS_LIBELLE From DOSSIER Where DOS_NODOSSIER = "' + Copy(V_PGI.DbName,3,1000) + '"');
        if lObTob.Detail.Count > 0 then
           Result := lObTob.Detail[0].GetValue('DOS_SOCIETE') + ' - ' + lObTob.Detail[0].GetValue('DOS_LIBELLE');
     finally
        lObTob.Free;
     end;
  end;

begin
  Inherited ;

  // Initialisation des controles
  TPageControl(GetControl('Pages')).Pages[0].Tabvisible := False;
  SetControlChecked('fCouleur', true);
  fOnbValiderClick := TToolbarButton97(GetControl('bValider')).OnClick;

  TToolbarButton97(GetControl('bValider')).OnClick := bValiderClick;
  THEdit(GetControl('EDEMPRUNT')).OnElipsisClick   := ChoixEmpruntClick;

  fInCodeEmprunt := 0;
  {YMO 19/03/2007 Selon que l'on travaille en exercice ou 12 mois glissants}
  If PerExo =1 then
  begin
    fChainePeriode1 := 'de période';
    fChainePeriode2 := 'la période';
  end
  else
  begin
    fChainePeriode1 := 'd''exercice';
    fChainePeriode2 := 'l''exercice';
  end;

  // Gestion du mode "MULTI-DOSSIER"
  FBoRegroupement := pos( 'MULTIDOSSIER', S ) > 0 ;
  SetControlVisible( 'BVMULTIDOSSIER',  FBoRegroupement )  ;
  SetControlVisible( 'TMULTIDOSSIER', FBoRegroupement )  ;
  SetControlVisible( 'MULTIDOSSIER', FBoRegroupement )  ;

  // Etat à imprimer
  lStEtat                 := ReadTokenSt(S);
  TFQRS1(Ecran).CodeEtat  := lStEtat;
  TFQRS1(Ecran).ChoixEtat := (lStEtat = '');

  // Code de l'emprunt associé
  lStEmprunt := ReadTokenSt(S);

  if lStEmprunt <> '' then
  begin
     fInCodeEmprunt := StrToInt(lStEmprunt); 
     lStEmprunt := lStEmprunt+ ' - ' +ReadTokenSt(S);
     SetControlText('EDEMPRUNT', lStEmprunt );
  end;

  {  FQ19986 YMO 18/04/07 Sélections possibles
     SetControlVisible('laEmprunt', False);
     SetControlVisible('edEmprunt', False);
     SetControlVisible('ladateref', False);
     SetControlVisible('eddateref', False);
   }

  // Liste d'exportation
  SetControlChecked('fListe', S <> '');
  SetControlVisible('fListe', lStEtat = '');
  SetcontrolVisible('Pages',  S = '');

  {FQ 19982 YMO 17/04/2007 Date de fin d'exercice}
  if VH^.EnCours.Fin > 0 then
     SetControlText('edDateRef',DateToStr(VH^.EnCours.Fin))
  else
     SetControlText('edDateRef',DateToStr(SysUtils.Date));

  // Lancement de l'aperçu
  if lStEtat <> '' then
     TToolBarButton97(GetControl('bValider')).Click;

  // Entête et pied de page
{  SetControlText('ENTETE1','');
  SetControlText('ENTETE2',NomCabinet);
  if V_PGI_Env.DBSocName =  V_PGI_Env.DBComName then
     SetControlText('ENTETE3','')
  else
     SetControlText('ENTETE3',V_PGI.CodeSociete + ' - ' + V_PGI.NomSociete);
  SetControlText('ENTETE4',TitreHalley + ' v' + v_pgi.NumVersion);
  SetControlText('ENTETE5','Imprimé le ' + DateToStr(Now));

  SetControlText('PIED1','');
  SetControlText('PIED2','');
}
end ;


procedure TOF_EMPIMP.OnClose ;
begin
  Inherited ;
end ;


// ----------------------------------------------------------------------------
// Nom    : bValiderClick
// Date   : 13/09/2002
// Auteur : D. ZEDIAR
// Objet  : Lancement de l'impression
// ----------------------------------------------------------------------------
// Historique :
// ----------------------------------------------------------------------------
procedure TOF_EMPIMP.bValiderClick(Sender : TObject);
begin
   SetControlText('XX_WHERE', '');

   if TFQRS1(Ecran).CodeEtat = 'ECR' then
   begin
      if ExtractionDonneesEcheancier then
         fOnbValiderClick(Self);
   end
   else
   if TFQRS1(Ecran).CodeEtat = 'PRE' then
   begin
      if ExtractionDonneesPrevisionnel then
         fOnbValiderClick(Self);
   end
   else
   if TFQRS1(Ecran).CodeEtat = 'TFF' then
   begin
      if ExtractionDonneesTableauFF then
         fOnbValiderClick(Self);
   end
   else
   if TFQRS1(Ecran).CodeEtat = 'TEM' then
   begin
      if ExtractionDonneesTableauEmprunt then
         fOnbValiderClick(Self);
   end
   else
   begin
      if fInCodeEmprunt = 0 then
         PGIInfo('Vous devez sélectionner un emprunt !', TitreHalley)
      else
      begin
         SetControlText('XX_WHERE', 'EMP_CODEEMPRUNT = ' + IntToStr(fInCodeEmprunt));
         if ExtractionDonneesTabAmo then
         fOnbValiderClick(Self);
      end;
   end;
end;


// ----------------------------------------------------------------------------
// Nom    : ChoixEmpruntClick
// Date   : 13/09/2002
// Auteur : D. ZEDIAR
// Objet  : Sélection d'un emprunt
// ----------------------------------------------------------------------------
// Historique :
// ----------------------------------------------------------------------------
procedure TOF_EMPIMP.ChoixEmpruntClick(Sender : TObject);
var
   lStEmprunt     : String;
begin
   lStEmprunt := AGLLanceFiche('FP','FMULEMPRUNT','','',';;SELECTION'); {22.05.07 YMO FQ20250}
   if lStEmprunt <> '' then
   begin
      fInCodeEmprunt := StrToInt(ReadTokenSt(lStEmprunt));
      SetControlText('EDEMPRUNT', IntToStr(fInCodeEmprunt) + ' - ' + lStEmprunt );
      SetControlText('XX_WHERE',  'EMP_CODEEMPRUNT = ' + IntTostr(fInCodeEmprunt) );
   end;
end;


Initialization
  registerclasses ( [ TOF_EMPIMP] ) ;
end.
