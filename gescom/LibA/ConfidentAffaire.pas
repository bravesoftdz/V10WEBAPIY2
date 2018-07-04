unit ConfidentAffaire;

interface

uses hmsgbox, hent1,Hctrls,sysUtils,M3FP;

Function VoirValorisationAct : Boolean;
Function SaisieActiviteManager : Boolean;
Function AffichageValorisation : Boolean;
Function AffichageFNF : Boolean;
Function CreationAffaireAutorise : Boolean;
Function ModifAffaireAutorise : Boolean;
Function SuppAffaireAutorise : Boolean;

Function JaiLeDroitTable (Table : String) : Boolean;
Function AGLJaiLeDroitFiche (parms: array of variant; nb: integer) : Variant;
function AFJaiLeDroitNatureGCCreat(Nature : string; Typepiece :string = 'Origine'; NatureOrig : string = '') : boolean;

implementation
uses UFonctionsCBP;

Function VoirValorisationAct : Boolean;
BEGIN
Result := False;
END;

// Gestion confidentialité sur l'accès réduit à une ressource sur la saisie d'activité
// Si SaisieActiviteManager=true, l'accès est ouvert à l'ensemble des ressources
// sinon, l'accès est limité à la ressource liée au user courant
// CC31,CC32,CC33 ... sont des "variable" définie dans Hent1n et correspond au concept réservé pour GI/GA
Function SaisieActiviteManager : Boolean;
BEGIN
Result := ExJaiLeDroitConcept ( cc31 , false);
END;

// Gestion confidentialité sur l'accès à la valorisation sur la saisie d'activité et sur la fiche ressource
// Si AffichageValorisation=true, la valorisation est visible sinon elle ne l'est pas...
Function AffichageValorisation : Boolean;
BEGIN
Result := ExJaiLeDroitConcept ( cc32 , false);
END;

// Gestion confidentialité sur l'accès à la saisie/lecture du Facturable/Non facturable sur la saisie d'activité
Function AffichageFNF : Boolean;
BEGIN
Result := ExJaiLeDroitConcept ( cc33 , false);
END;


Function CreationAffaireAutorise : Boolean;
BEGIN
Result := ExJaiLeDroitConcept ( cc34 , false);
END;

Function ModifAffaireAutorise : Boolean;
BEGIN
Result := ExJaiLeDroitConcept ( cc35 , false);
END;

Function SuppAffaireAutorise : Boolean;
BEGIN
Result := ExJaiLeDroitConcept ( cc37 , false);
END;


// Gestion confidentialité des accès aux tables
Function JaiLeDroitTable (Table : String) : Boolean;
         // en fct de la tablette reçu, regarde le tag associé
var text, titre :string;
begin
Result := true; // par défaut ok
if ctxGCAFF in V_PGI.PGIContexte then Exit;
{$ifndef AFFAIRE}
Exit;
{$ENDIF}
Text :='Vous n''avez pas le droit d''utiliser cette option';
titre := 'Accès Interdit !';

if Table ='TTFORMEJURIDIQUE' then Result:= JaiLeDroitTag(74156) else
if Table ='TTCIVILITE' then Result:= JaiLeDroitTag(74157) else
if Table ='TTLANGUE' then Result:= JaiLeDroitTag(74161) else
if Table ='GCEMPLOIBLOB' then Result:= JaiLeDroitTag(74204) else
if Table ='GCCOMMENTAIREENT' then Result:= JaiLeDroitTag(74211) else
if Table ='GCCOMMENTAIRELIGNE' then Result:= JaiLeDroitTag(74212) else
if Table ='GCCOMMENTAIREPIED' then Result:= JaiLeDroitTag(74213) else
if Table ='DPCODENAF' then Result:= JaiLeDroitTag(74311) else
if Table ='GCCOMPTATIERS' then Result:= JaiLeDroitTag(74312) else
if Table ='GCORIGINETIERS' then Result:= JaiLeDroitTag(74316) else
if Table ='GCZONECOM' then Result:= JaiLeDroitTag(74313) else
if Table ='TTSECTEUR' then Result:= JaiLeDroitTag(74314) else
if Table ='TTTARIFCLIENT' then Result:= JaiLeDroitTag(74315) else
if Table ='TTFONCTION' then Result:= JaiLeDroitTag(74321) else
if Table ='AFCOMPTAAFFAIRE' then Result:= JaiLeDroitTag(74331) else
if Table ='AFDEPARTEMENT' then Result:= JaiLeDroitTag(74332) else
if Table ='AFTLIENAFFTIERS' then Result:= JaiLeDroitTag(74333) else
if Table ='AFTRESILAFF' then Result:= JaiLeDroitTag(74334) else
if Table ='AFFAIREPART1' then Result:= JaiLeDroitTag(74335) else
if Table ='GCFAMILLENIV1' then Result:= JaiLeDroitTag(74341) else
if Table ='GCFAMILLENIV2' then Result:= JaiLeDroitTag(74342) else
if Table ='GCFAMILLENIV3' then Result:= JaiLeDroitTag(74343) else
if Table ='GCLIBFAMILLE' then Result:= JaiLeDroitTag(74344) else
if Table ='GCCOMPTAARTICLE' then Result:= JaiLeDroitTag(74345) else
if Table ='GCTARIFARTICLE' then Result:= JaiLeDroitTag(74346) else
if Table ='AFTTARIFRESSOURCE' then Result:= JaiLeDroitTag(74346) else
if Table ='AFTNIVEAUDIPLOME' then Result:= JaiLeDroitTag(74354) else
if Table ='AFTTYPERESSOURCE' then Result:= JaiLeDroitTag(74355) else
if Table ='AFTSTANDCALEN' then Result:= JaiLeDroitTag(74361) else
if Table ='AFTYPEHEURE' then Result:= JaiLeDroitTag(74371) else
if Table ='GCCATTIERSTABLE' then Result:= JaiLeDroitTag(74400) else
if Table ='GCCATARTTABLE' then Result:= JaiLeDroitTag(74401) else
if Table ='AFCARESTABLE' then Result:= JaiLeDroitTag(74402) else
if Table ='AFCAAFFTABLE' then Result:= JaiLeDroitTag(74403) else
if Table ='GCCATETABLISSTABLE' then Result:= JaiLeDroitTag(74404) else
if Table ='AFCATIERSTABLE' then Result:= JaiLeDroitTag(74405) else
if Table ='GCLIBRETIERS1' then Result:= JaiLeDroitTag(74411) else
if Table ='GCLIBRETIERS2' then Result:= JaiLeDroitTag(74412) else
if Table ='GCLIBRETIERS3' then Result:= JaiLeDroitTag(74413) else
if Table ='GCLIBRETIERS4' then Result:= JaiLeDroitTag(74414) else
if Table ='GCLIBRETIERS5' then Result:= JaiLeDroitTag(74415) else
if Table ='GCLIBRETIERS6' then Result:= JaiLeDroitTag(74416) else
if Table ='GCLIBRETIERS7' then Result:= JaiLeDroitTag(74417) else
if Table ='GCLIBRETIERS8' then Result:= JaiLeDroitTag(74418) else
if Table ='GCLIBRETIERS9' then Result:= JaiLeDroitTag(74419) else
if Table ='GCLIBRETIERSA' then Result:= JaiLeDroitTag(74420) else
if Table ='GCLIBREFOU1' then Result:= JaiLeDroitTag(74561) else
if Table ='GCLIBREFOU2' then Result:= JaiLeDroitTag(74562) else
if Table ='GCLIBREFOU3' then Result:= JaiLeDroitTag(74563) else
if Table ='AFTLIBREAFF1' then Result:= JaiLeDroitTag(74461) else
if Table ='AFTLIBREAFF2' then Result:= JaiLeDroitTag(74462) else
if Table ='AFTLIBREAFF3' then Result:= JaiLeDroitTag(74463) else
if Table ='AFTLIBREAFF4' then Result:= JaiLeDroitTag(74464) else
if Table ='AFTLIBREAFF5' then Result:= JaiLeDroitTag(74465) else
if Table ='AFTLIBREAFF6' then Result:= JaiLeDroitTag(74466) else
if Table ='AFTLIBREAFF7' then Result:= JaiLeDroitTag(74467) else
if Table ='AFTLIBREAFF8' then Result:= JaiLeDroitTag(74468) else
if Table ='AFTLIBREAFF9' then Result:= JaiLeDroitTag(74469) else
if Table ='AFTLIBREAFFA' then Result:= JaiLeDroitTag(74489) else
if Table ='AFTLIBRERES1' then Result:= JaiLeDroitTag(74431) else
if Table ='AFTLIBRERES2' then Result:= JaiLeDroitTag(74432) else
if Table ='AFTLIBRERES3' then Result:= JaiLeDroitTag(74433) else
if Table ='AFLIBRETACHE1' then Result:= JaiLeDroitTag(74581) else
if Table ='AFLIBRETACHE2' then Result:= JaiLeDroitTag(74582) else
if Table ='AFLIBRETACHE3' then Result:= JaiLeDroitTag(74583) else
if Table ='AFLIBRETACHE4' then Result:= JaiLeDroitTag(74584) else
if Table ='AFLIBRETACHE5' then Result:= JaiLeDroitTag(74585) else
if Table ='AFLIBRETACHE6' then Result:= JaiLeDroitTag(74586) else
if Table ='AFLIBRETACHE7' then Result:= JaiLeDroitTag(74587) else
if Table ='AFLIBRETACHE8' then Result:= JaiLeDroitTag(74588) else
if Table ='AFLIBRETACHE9' then Result:= JaiLeDroitTag(74589) else
if Table ='AFLIBRETACHEA' then Result:= JaiLeDroitTag(74590) else
if Table ='AFFAMILLETACHE' then Result:= JaiLeDroitTag(74381) else
if Table ='GCLIBREART1' then Result:= JaiLeDroitTag(74441) else
if Table ='GCLIBREART2' then Result:= JaiLeDroitTag(74442) else
if Table ='GCLIBREART3' then Result:= JaiLeDroitTag(74443) else
if Table ='GCLIBREART4' then Result:= JaiLeDroitTag(74444) else
if Table ='GCLIBREART5' then Result:= JaiLeDroitTag(74445) else
if Table ='GCLIBREART6' then Result:= JaiLeDroitTag(74446) else
if Table ='GCLIBREART7' then Result:= JaiLeDroitTag(74447) else
if Table ='GCLIBREART8' then Result:= JaiLeDroitTag(74448) else
if Table ='GCLIBREART9' then Result:= JaiLeDroitTag(74449) else
if Table ='GCLIBREARTA' then Result:= JaiLeDroitTag(74450) else

if Table ='GCLIBREPIECE1' then Result:= JaiLeDroitTag(74451) else
if Table ='GCLIBREPIECE2' then Result:= JaiLeDroitTag(74452) else
if Table ='GCLIBREPIECE3' then Result:= JaiLeDroitTag(74453) else
if Table ='YYLIBREET1' then Result:= JaiLeDroitTag(74470) else
if Table ='YYLIBREET2' then Result:= JaiLeDroitTag(74471) else
if Table ='YYLIBREET3' then Result:= JaiLeDroitTag(74472) else
if Table ='YYLIBREET4' then Result:= JaiLeDroitTag(74473) else
if Table ='YYLIBREET5' then Result:= JaiLeDroitTag(74474) else
if Table ='YYLIBREET6' then Result:= JaiLeDroitTag(74475) else
if Table ='YYLIBREET7' then Result:= JaiLeDroitTag(74476) else
if Table ='YYLIBREET8' then Result:= JaiLeDroitTag(74477) else
if Table ='YYLIBREET9' then Result:= JaiLeDroitTag(74478) else
if Table ='YYLIBREETA' then Result:= JaiLeDroitTag(74479) else
if Table ='YYLIBRECON1' then Result:= JaiLeDroitTag(74541) else
if Table ='YYLIBRECON2' then Result:= JaiLeDroitTag(74542) else
if Table ='YYLIBRECON3' then Result:= JaiLeDroitTag(74543) else
;

if result = False then PgiInfo (Text,Titre);
end;


// Gestion confidentialité des accès aux fiches
Function AGLJaiLeDroitFiche (parms: array of variant; nb: integer) : Variant;
         // en fct de la Fiche reçu, regarde le tag associé
var text, titre, fiche, typfic ,action,Zaction:string;
ind : integer;
begin
result:=TRUE;
if ctxGCAFF in V_PGI.PGIContexte then Exit;
{$ifndef AFFAIRE}
Exit;
{$ENDIF}
Text :='Vous n''avez pas le droit d''utiliser cette option';
titre := 'Accès Interdit !';
Fiche := STring(Parms[0]);
fiche:=AnsiUppercase (Fiche);
  // on récupère l'aciton pour , par la suite,
  //permettre une confidentialité sur creation/modif/consul ....
ZAction:= STring(Parms[1]);
Zaction:=AnsiUppercase (Zaction);
ind :=Pos('ACTION',Zaction);
if (ind <>0) then begin
   Delete(ZAction,1,ind+6) ;
   Action:=AnsiUppercase(ReadTokenSt(ZAction)) ;
   end;
if nb = 3 then begin
   TypFic := STring(Parms[2]);
   typfic:= AnsiUppercase (TypFic);
   end;

   // nom des fiches mises en ordre alpha
if Fiche ='ACTIVITEMUL' then Result:= JaiLeDroitTag(72102)   //Modif activite
else if Fiche ='AFFAIRE' then begin    //fiche affaire et devis
   if (copy (typfic,1,6)='STATUT') then typfic:=COpy(Typfic,8,1);
   if Copy(TypFic,1,1) = 'A' then begin
                       Result:= JaiLeDroitTag(-71103);
                       If Result=True then Result:= JaiLeDroitTag(71107);
                       If Result=True then Result:= JaiLeDroitTag(71103);
                       end;
   if Copy(TypFic,1,1) = 'P' then Result:= JaiLeDroitTag(71102);
   end
else if Fiche ='AFTIERS' then
   begin
   Result:= JaiLeDroitTag(-71101); //fiche client
   If Result=True then Result:= JaiLeDroitTag(71101);
   If Result=True then Result:= JaiLeDroitTag(71106);
   end
else if Fiche ='AFTBVIEWER' then Result:= JaiLeDroitTag(72401) //Tableau bord
else if Fiche ='AFGRANDLIVRE' then Result:= JaiLeDroitTag(72402) //GL
else if Fiche ='PIECE' then begin         // piece (facture,avoir, cde ...
       If Copy(TypFic, 1,3) = 'FAC' then result:=JaiLeDroitTag(73101);
       If Copy(TypFic, 1,3) = 'AVC' then result:=JaiLeDroitTag(73102);
       If Copy(TypFic, 1,3) = 'FPR' then result:=JaiLeDroitTag(73103);
       If Copy(TypFic, 1,3) = 'APR' then result:=JaiLeDroitTag(73105);
       If Copy(TypFic, 1,2) = 'CC'  then result:=JaiLeDroitTag(71103);
       If Copy(TypFic, 1,2) = 'DE'  then result:=JaiLeDroitTag(71102);
    end
else if Fiche ='RESSOURCE' then Result:= JaiLeDroitTag(71104)
else if Fiche ='ARTICLE' then begin         //fiche artcile, prest,frais
   if Copy(TypFic,1,1) = 'M' then Result:= JaiLeDroitTag(71513)     //MAR fourniture
   else if Copy(TypFic,1,1) = 'F' then Result:= JaiLeDroitTag(71512)  //FRA frais
   else if Copy(TypFic,1,1) = 'C' then Result:= JaiLeDroitTag(71514)  //CTR contrat
   else if Copy(TypFic,1,1) = 'B' then Result:= JaiLeDroitTag(71515)  //POU pourcenatge
        else Result:= JaiLeDroitTag(71511);
   end
else if Fiche ='COMMERCIAL' then Result:= JaiLeDroitTag(71501)   //apporteur et GCCOMEMRCIAL
else if Fiche ='COMPETENCE' then Result:= JaiLeDroitTag(74353)   //Competence ressorue
else if Fiche ='CONTACT' then Result:= JaiLeDroitTag(71503)   //contact client
else if Fiche ='FONCTION' then Result:= JaiLeDroitTag(74352)  //focntion ressource
else if Fiche ='IMPCALEN' then Result:= JaiLeDroitTag(71241)
else if Fiche ='IMPCALREGLE' then Result:= JaiLeDroitTag(71242);
if result = False then PgiInfo (Text,Titre);
end;                 


{*******************************************************************************
Auteur  ...... : Alain Buchet
Créé le ...... : 07/02/2008
Modifié le ... :
Description .. : Fiche qualité 14525
Description .. : Appel depuis JaiLeDroitNatureGCCreat en contexte Affaire
Description .. : Droit de création d'une pièce lors d'un accès
Description .. : NatureOrig pour la génération de pièces de natures différentes
Mots clefs ... : Nature :pièce à créer - NatureOrig: pièce orig à transformer -
***********************************************************************************}
function AFJaiLeDroitNatureGCCreat(Nature : string; Typepiece :string = 'Origine'; NatureOrig : string = '') : boolean;
begin
    if Typepiece <> 'Destination' then
    begin
    //Saisie vente - Gestion d'affaire
      if Nature = 'FPR' then Result := JaiLeDroitTag(73103) //Facture provisoire
      else if Nature = 'FAC' then Result := JaiLeDroitTag(73101) //Facture
      else if Nature = 'AVC' then Result := JaiLeDroitTag(73102) //Avoir client
      else if Nature = 'APR' then Result := JaiLeDroitTag(73105) //Avoir provisoire
    //Saisie achat - Gestion d'affaire
      else if Nature = 'DEF' then Result := JaiLeDroitTag(138105) //Proposition d'achat
      else if Nature = 'CF' then Result := JaiLeDroitTag(138101) //Commande fournisseur
      else if Nature = 'BLF' then Result := JaiLeDroitTag(138102) //Livraison fournisseur
      else if Nature = 'AF' then Result := JaiLeDroitTag(138104) //avoir fournisseur financier
      else if Nature = 'FF' then Result := JaiLeDroitTag(138103) //Facture fournisseur
      else if Nature = 'PRF' then Result := JaiLeDroitTag(31204) //Préparation de reception
      else Result := true;
    end else
    begin
    //Génération vente - Gestion d'affaire
      if Nature = 'FPR' then
        begin
          Result := JaiLeDroitTag(73103); //Facture provisoire
          if Result then Result := JaiLeDroitTag(73403); // Validation des documents provisoires
        end
      else if Nature = 'FAC' then
        begin
          Result := JaiLeDroitTag(73101); //Facture
          if Result then Result := JaiLeDroitTag(30405); // Générer Facture
        end
      else if Nature = 'AVC' then
        begin
          Result := JaiLeDroitTag(73102); //Avoir client
          if Result then
            if NatureOrig = 'FAC' then
              Result := JaiLeDroitTag(73408); // Générer Facture en avoir client (avoir financier)
        end
      else if Nature = 'APR' then //Avoir provisoire
        Result := JaiLeDroitTag(73105)
      //Génération achat - Gestion d'affaire
      else if Nature = 'CF' then //Commande fournisseur
        begin
          Result := JaiLeDroitTag(138101);
          if Result then Result := JaiLeDroitTag(138304); // Générer cde
          if Result then
            if NatureOrig = 'DEF' then
            Begin
              Result := JaiLeDroitTag(31451); // Générer manuellement une cde à partir d'une proposition
              if result then Result := JaiLeDroitTag(138323); // Générer automatiquement une cde à partir d'une proposition
            end
        end
        else if Nature = 'PRF' then //Préparation réception
        begin
          Result := JaiLeDroitTag(31204);
          if Result then Result := JaiLeDroitTag(31513); // Générer Préparation reception
          if Result then
            if NatureOrig = 'CF' then
            Begin
              Result := JaiLeDroitTag(31452); // Générer manuellement une Prépa reception à partir d'une cde
              if result then Result := JaiLeDroitTag(31513); // Générer automatiquement une Prépa réception à partir d'une cde
            end
        end
        else if Nature = 'BLF' then // Réception fournisseur
        begin
          Result := JaiLeDroitTag(138102);
          if Result then Result := JaiLeDroitTag(138302); // Générer réception
          if Result then
            if NatureOrig = 'CF' then
            Begin
              Result := JaiLeDroitTag(31453); // Générer manuellement une réception à partir d'une commande
              if result then Result := JaiLeDroitTag(138321); // Générer automatiquement une réception à partir d'une commande
            end
            else if NatureOrig = 'PRF' then
            Begin
              Result := JaiLeDroitTag(31455); // Générer manuellement une réception à partir d'une préparation
              if result then Result := JaiLeDroitTag(31523); // Générer automatiquement une réception à partir d'une préparation
            end
        end
      else if Nature = 'AF' then  //Avoir fournisseur
        begin
          Result := JaiLeDroitTag(138104);
        end
      else if Nature = 'FF' then  //Facture fournisseur
        begin
          Result := JaiLeDroitTag(138103);
          if Result then Result := JaiLeDroitTag(138303); // Générer Facture
          if Result then
            if NatureOrig = 'CF' then
            Begin
              Result := JaiLeDroitTag(31454); // Générer manuellement une facture à partir d'une cde
            end
            else if NatureOrig = 'PRF' then
            Begin
              Result := JaiLeDroitTag(31456); // Générer manuellement une facture à partir d'une Prépa Livr
            end
            else if NatureOrig = 'BLF' then
            Begin
              Result := JaiLeDroitTag(31457); // Générer manuellement une facture à partir d'une Livraison
              if result then Result := JaiLeDroitTag(138322); // Générer automatiquement une facture à partir d'une Livraison
            end
        end
      else Result := true;
    end;
end;



Initialization
RegisterAglFunc( 'JaiLeDroitFiche', FALSE , 2, AGLJaiLeDroitFiche);
end.
