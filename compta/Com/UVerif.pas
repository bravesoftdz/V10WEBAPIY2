unit UVerif;

interface

uses SysUtils, Classes, UTOB, MesgErrCom ;

type
  ProcCharge = procedure (Ligne : string) of object;
  TTypeLigne=(ENT,PS1,PS2,PS3,PS4,PS5,EXO,CGN,CAE,JAL,ECR,ETB,MDP);

Erreur = record
       NumeroLigne : integer;
       Champ       : string;
       MsgErreur   : array[1..10] of string;
  end;

AffComEnreg = record
                    Typeligne   : array[1..7] of string;
                    TOBCharge   : ProcCharge;
                    MTOB        : TOB;
                    FTOBErreur  : TOB;
                    MTOBTout    : TOB;
                    ColName     : string;
                    Verif       : TTypeLigne;
                    CodeErreur  : integer;
                    LongLigne   : integer;
              end;

type
 TVerif = class
        TypeLigne: TTypeLigne ;
       private
             FUNCTION  VerifNature(champ:string ; TabNature:array of String ; nb:integer) : boolean;
             FUNCTION  EstVide(var TD:TOB ; Champ:string ; Erreur:string) : boolean;
             FUNCTION  ControlDate(date:string):boolean;
             FUNCTION  ValidMoisJour(annee:integer ; mois:integer ; jour:integer) : boolean;
             FUNCTION  ValidBorne(test:integer ; borne1:integer ; borne2:integer) : boolean;
             FUNCTION  EnleveVide(champ:string) : string;
             function  VerifExistLigne(var MTOB : TOB; MessgErreur : string;champ:string): integer;
             FUNCTION  DATEVALIDE(var TD :TOB ; Champ : string ; MsgErreur : string) : boolean;
             FUNCTION  VerifLongueur(TD : TOB ; Champ : string; borne1,borne2:integer;MsgErr:string):boolean;
             FUNCTION  RechercheCode(VALEUR :string ; MTOB:TOB ; Champ: string):integer;

             PROCEDURE OUI_NON(var TD:TOB ; Champ:string ; Erreur:string);
             PROCEDURE RechercheErrCmptGnrx(var TD:TOB);
             PROCEDURE VerifEntete(var MTOB:TOB);
             PROCEDURE VerifPS1(var MTOB:TOB);
             PROCEDURE VerifPS2(var MTOB:TOB);
             PROCEDURE VerifPS3(var MTOB:TOB);
             PROCEDURE VerifPS5(var MTOB:TOB);
             PROCEDURE VerifEcriture(var Enreg:array of AffComEnreg ; i : integer);
             PROCEDURE VerifCmptGnrx(var MTOB:TOB);
             PROCEDURE VerifExercice(var MTOB:TOB);
             PROCEDURE VerifJournaux(var MTOB:TOB);
             PROCEDURE VerifCmpTiers(var MTOB:TOB);
             PROCEDURE VerifEtab(var MTOB:TOB);
             function  VerifLettrage(var TD:TOB):boolean;
             PROCEDURE InitTabNatureJournaux;
             PROCEDURE InitTabNatureCmptGnrx;
             PROCEDURE InitTabTypePieceEcr;
             PROCEDURE AlimTOBErreur(var TD:TOB ;champ : string ; MsgErreur : string);
        public
             procedure ControlTRA(var Enreg: array of AffComEnreg ; i : integer);
             PROCEDURE RecupLigneErreur(var Enreg: AffComEnreg);
             FUNCTION  RechercheLigne(Enreg: array of AffComEnreg ; typeRecherch : TTypeLigne):TOB;
end;

var TabNatureJournaux:array [1..12] of String;
    TabNatureCmptGnrx:array[1..13] of String;
    TabTypePieceEcr  :array[1..10] of String;
    nbNatureJournaux : integer;
    nbNatureCmptGnrx : integer;
    nbTypePieceEcr   : integer;      

implementation

{***********A.G.L.***********************************************
Auteur  ...... : Marion Koszczuk
Créé le ...... : 15/06/2004
Modifié le ... :   /  /
Description .. : Procedure appelée lors du chargement des lignes du fichier,
Suite ........ : elle controle les lignes selon le type .
Mots clefs ... : CONTROL TRA
*****************************************************************}
procedure TVerif.ControlTRA(var Enreg: array of AffComEnreg ; i : integer);
var TypeLigne : TTypeLigne ;
begin
     InitTabNatureJournaux;
     InitTabNatureCmptGnrx;
     InitTabTypePieceEcr;
     TypeLigne:=Enreg[i].Verif;

     Case TypeLigne of
       ENT:  VerifEntete(Enreg[i].MTOB);//MARCHE
       PS1:  VerifPS1(Enreg[i].MTOB);
       PS2:  VerifPS2(Enreg[i].MTOB); //MARCHE
       PS3:  VerifPS3(Enreg[i].MTOB);
       PS5:  VerifPS5(Enreg[i].MTOB);
       EXO:  VerifExercice(Enreg[i].MTOB);
       CGN:  VerifCmptGnrx(Enreg[i].MTOB); //MARCHE
       CAE:  VerifCmpTiers(Enreg[i].MTOB);//MARCHE
       JAL:  VerifJournaux(Enreg[i].MTOB);
       ECR:  VerifEcriture(Enreg,i);
       ETB:  VerifEtab(Enreg[i].MTOB);
       MDP:  exit;
     End;
end;


{***********A.G.L.***********************************************
Auteur  ...... : Marion Koszczuk
Créé le ...... : 15/06/2004
Modifié le ... : 15/06/2004
Description .. : Initialisation du tableau contenant les natures des journaux.
Suite ........ : Pour ensuite pouvoir tester la validité de la nature
Mots clefs ... : TABLEAU NATURE JOURNAL
*****************************************************************}
procedure TVerif.InitTabNatureJournaux;
var i : integer;
begin
  i := 1;
  TabNatureJournaux[i] := 'ACH'; i := i+1; TabNatureJournaux[i] := 'ANA'; i := i+1;
  TabNatureJournaux[i] := 'BQE'; i := i+1; TabNatureJournaux[i] := 'CAI'; i := i+1;
  TabNatureJournaux[i] := 'CLO'; i := i+1; TabNatureJournaux[i] := 'ECC'; i := i+1;
  TabNatureJournaux[i] := 'EXT'; i := i+1; TabNatureJournaux[i] := 'OD';  i := i+1;
  TabNatureJournaux[i] := 'ODA'; i := i+1; TabNatureJournaux[i] := 'REG'; i := i+1;
  TabNatureJournaux[i] := 'VTE'; i := i+1; TabNatureJournaux[i] := 'ANO';
  nbNatureJournaux := i;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Marion Koszczuk
Créé le ...... : 15/06/2004
Modifié le ... :   /  /    
Description .. : Initialisation du tableau contenant les natures des journaux.
Suite ........ : Pour ensuite pouvoir tester la validité de la nature
Mots clefs ... : TABLEAU NATURE COMPT GNRX
*****************************************************************}
procedure TVerif.InitTabNatureCmptGnrx;
var i : integer;
begin
  i := 1;
  TabNatureCmptGnrx[i] := 'BQE'; i := i+1; TabNatureCmptGnrx[i] := 'CAI'; i := i+1;
  TabNatureCmptGnrx[i] := 'CHA'; i := i+1; TabNatureCmptGnrx[i] := 'COC'; i := i+1;
  TabNatureCmptGnrx[i] := 'COD'; i := i+1; TabNatureCmptGnrx[i] := 'COF'; i := i+1;
  TabNatureCmptGnrx[i] := 'COS'; i := i+1; TabNatureCmptGnrx[i] := 'DIV'; i := i+1;
  TabNatureCmptGnrx[i] := 'EXT'; i := i+1; TabNatureCmptGnrx[i] := 'IMO'; i := i+1;
  TabNatureCmptGnrx[i] := 'PRO'; i := i+1; TabNatureCmptGnrx[i] := 'TIC'; i := i+1;
  TabNatureCmptGnrx[i] := 'TID';
  nbNatureCmptGnrx := i;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Marion Koszczuk
Créé le ...... : 15/06/2004
Modifié le ... :   /  /    
Description .. : Initialisation du tableau contenant les types de pieces des 
Suite ........ : ecritures.
Suite ........ : Pour ensuite pouvoir tester la validité des pieces
Mots clefs ... : TABLEAU TYPE PIECE ECRITURE
*****************************************************************}
procedure TVerif.InitTabTypePieceEcr;
var i : integer;
begin
  i := 1;
  TabTypePieceEcr[i] := 'FC'; i := i+1; TabTypePieceEcr[i] := 'AC'; i := i+1;
  TabTypePieceEcr[i] := 'RC'; i := i+1; TabTypePieceEcr[i] := 'FF'; i := i+1;
  TabTypePieceEcr[i] := 'AF'; i := i+1; TabTypePieceEcr[i] := 'RF'; i := i+1;
  TabTypePieceEcr[i] := 'OD'; i := i+1; TabTypePieceEcr[i] := 'OC'; i := i+1;
  TabTypePieceEcr[i] := 'OF';
  nbTypePieceEcr := i;
end;

{------------------------------------------------------------------------------
---------------------------*FONCTIONS DE CONTROLE*-----------------------------
-------------------------------------------------------------------------------
-------------------------------------------------------------------------------}
{***********A.G.L.***********************************************
Auteur  ...... : Marion Koszczuk
Créé le ...... : 15/06/2004
Modifié le ... :   /  /    
Description .. : fonction qui renvoie si la TOB Mere passée en param existe 
Suite ........ : et si elle contient des TOB Filles. 
Mots clefs ... : VERIF LIGNES OBLIGATOIRES
*****************************************************************}
function TVerif.VerifExistLigne(var MTOB : TOB; MessgErreur : string;champ:string): integer;
var nbligne : integer;
    TOBFille : TOB;
begin
     if(Assigned(MTOB))then
     begin
          nbligne:= MTOB.Detail.Count;
          Result := nbligne;
     end else
     begin
          TOBFille := TOB.Create('',MTOB,-1);
          AlimTOBErreur(TOBFille,champ,MessgErreur);
          Result := 0;
     end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Marion Koszczuk
Créé le ...... : 15/06/2004
Modifié le ... :   /  /    
Description .. : Procedure qui supprim les lignes contenant des erreurs de la 
Suite ........ : TOB courante et les inseres dans la TOB Erreur et la TOB 
Suite ........ : tout contenant toutes les lignes.
Mots clefs ... : REMPLI TOB ERREUR
*****************************************************************}
PROCEDURE TVerif.RecupLigneErreur(var Enreg: AffComEnreg);
var TD,MTOB,TOBTmp,TOBTout : TOB;
    i                      : integer;
begin
     MTOB := Enreg.MTOB;
     Enreg.MTOBTout := TOB.Create('TOBMereTout',nil,-1); 
     i:=0;
     if Assigned(MTOB) then
     begin
          while (i<=MTOB.Detail.Count-1) do
          begin
               TOBTmp := MTOB.Detail[i];
               if TOBTmp.GetValue('Erreur')='1' then
               begin
                    TD := TOB.Create('',Enreg.FTOBErreur,-1);
                    TD.Dupliquer(TOBTmp,TRUE,TRUE,FALSE);
                    TOBTout := TOB.Create('',Enreg.MTOBTout,-1);
                    TOBTout.Dupliquer(TOBTmp,TRUE,TRUE,FALSE);
                    TOBTout.AddChampSupValeur('Position ligne',TOBTout.GetValue('Position ligne')+' '+'!!!ERREUR!!!');
                    TOBTmp.Free;
               END ELSE
               BEGIN
                    TOBTout := TOB.Create('',Enreg.MTOBTout,-1);
                    TOBTout.Dupliquer(TOBTmp,TRUE,TRUE,FALSE);
                    i:=i+1;
               END;
          end;
     end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Marion Koszczuk
Créé le ...... : 15/06/2004
Modifié le ... :   /  /    
Description .. : Insere le message d'erreur et cree un champ erreur à 1 dans 
Suite ........ : la TOB courante pour pouvoir ensuite recup les lignes 
Suite ........ : erreurs.
Mots clefs ... : INSERE MESSAGE ERREUR
*****************************************************************}
PROCEDURE TVerif.AlimTOBErreur(var TD:TOB;champ:string;MsgErreur:string);
var Valeur : string;
begin
     if TD.FieldExists(champ)=TRUE then
     begin
          if champ='Montant1' then
          begin
               Valeur := FloatToStr(TD.GetValue(champ));
          end else
              Valeur := TD.GetValue(champ);
          TD.PutValue(champ,valeur +' '+MsgErreur);
     end else
          TD.AddChampSupValeur(champ,' '+MsgErreur);
   TD.AddChampSupValeur('Erreur',1);
end;

{***********A.G.L.***********************************************
Auteur  ...... : Marion Koszczuk
Créé le ...... : 15/06/2004
Modifié le ... :   /  /    
Description .. : Control des lignes de type Entete: existance de l'entete, des 
Suite ........ : champs obligatoire et verif le format de certains champs.
Mots clefs ... : CONTROL ENTETE
*****************************************************************}
PROCEDURE TVerif.VerifEntete(var MTOB:TOB);
var TD           : TOB;
    i,nbligne    : integer;
BEGIN
     TD := TOB.Create('',nil,-1);
     nbligne := VerifExistLigne(MTOB,ENTETE_INEXSISTANT,'Identifiant');
     if nbligne>0 then
     begin
          for i := 0 to nbligne-1 do
          begin
               TD := MTOB.Detail[i];
               if(TD.GetValue('Identifiant')<>'!')then
               begin
                    //CONTROL TYPE FICHIER
                    IF ( EstVide(TD,'Type Fichier',ENTETE_TYPEMANQUANT)=FALSE ) then
                    BEGIN
                         IF ( (Trim(TD.GetValue('Type Fichier'))<>'JRL') and (Trim(TD.GetValue('Type Fichier'))<>'DOS') and (Trim(TD.GetValue('Type Fichier'))<>'BAL') and (Trim(TD.GetValue('Type Fichier'))<>'SYN') )then
                             AlimTOBErreur(TD,'Type Fichier',ENTETE_TYPEINCORRECT);
                    End;
                    //CONTROL ORIGINE FICHIER
                    IF (EstVide(TD,'Origine Fichier',ENTETE_ORIGINE2) = FALSE) then
                    BEGIN
                         IF ((TD.GetValue('Origine Fichier')<>'CLI') and (TD.GetValue('Origine Fichier')<>'EXP')) then
                            AlimTOBErreur(TD,'Origine Fichier',ENTETE_ORIGINE);
                    END;
                    //CONTROL FORMAT DU FICHIER
                    IF (EstVide(TD,'Format',ENTETE_FORMAT2)= FALSE) then
                    BEGIN
                         IF ((TD.GetValue('Format')<>'STD') and (TD.GetValue('Format')<>'ETE')) then
                           AlimTOBErreur(TD,'Format',ENTETE_FORMAT);
                    END;
                    //CONTROL SI TYPE='BAL' alors FORMAT='STD'
                    IF ((EstVide(TD,'Type Fichier',ENTETE_TYPEMANQUANT)= FALSE) and (EstVide(TD,'Format',ENTETE_FORMAT2)= FALSE)) then
                    BEGIN
                         IF ((TD.GetValue('Type Fichier')='BAL') and  (TD.GetValue('FORMAT')<>'STD')) then
                           AlimTOBErreur(TD,'Type Fichier',ENTETE_FORMAT_TYPE);
                    END;
                    //CONTROL VERSION DU FICHIER
                    EstVide(TD,'Version',ENTETE_VERSION);
                   //CONTROL NUMDOSSCAB
                   //EstVide(TD,'NumDoCab',ENTETE_NUMDOCAB);    voir avec synchro S1
                   //CONTROL FORMAT DATE BASCULE
                     IF(EstVide(TD,'DateBascule','')=FALSE)then
                        DATEVALIDE(TD,'DateBascule',ENTETE_DTEBASCUL);
                    //CONTROL FORMAT DATE ARR PERIODIQUE
                    IF(EstVide(TD,'DteArrPeriodiq','')=FALSE)then
                       DATEVALIDE(TD,'DteArrPeriodiq',ENTETE_DTEARRPER);
                    MTOB.Detail[i]:=TD;
               end else
               begin
                    MTOB.Detail[i].PutValue('Identifiant','! : dossier existant dans la base');
               end;
           end;
     end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Marion Koszczuk
Créé le ...... : 15/06/2004
Modifié le ... :   /  /
Description .. : Verif existance lignes PARAM G1
Mots clefs ... : CONTROL PARAMG1
*****************************************************************}
Procedure TVerif.VerifPS1(var MTOB:TOB);
begin
     VerifExistLigne(MTOB,ERR_PARAMETRE+'1'+ERR_PARAMETRESUITE,'Identifiant');
end;

{***********A.G.L.***********************************************
Auteur  ...... : Marion Koszczuk
Créé le ...... : 15/06/2004
Modifié le ... :   /  /    
Description .. : Verif existance ParamG2, champs obligatoires et format de 
Suite ........ : certains champs
Mots clefs ... : CONTROL PARAMG2
*****************************************************************}
Procedure TVerif.VerifPS2(var MTOB:TOB);
var TD         : TOB;
    i,nbligne  :integer;
BEGIN
     TD := TOB.Create('',nil,-1);
     nbligne := VerifExistLigne(MTOB,ERR_PARAMETRE+'2'+ERR_PARAMETRESUITE,'Identifiant');
     If (nbligne>0) then
     BEGIN
          for i := 0 to nbligne-1 do
          begin
               TD := MTOB.Detail[i];
               //longueur des comptes entre 6 et 17 ERR_LGCPTEPS2
               IF(EstVide(TD,'Lggen',ERR_LGCPTEPS2) = FALSE) then
                  VerifLongueur(TD,'Lggen',17,6,ERR_LGCPTEPS2);
               //longueur des compte de tiers entre 6 et 17
               IF(EstVide(TD,'Lgaux',ERR_LGCPTEPS2) = FALSE) then
                  VerifLongueur(TD,'Lgaux',17,6,ERR_LGCPTEPS2);
               //longueur section analytique entre 3 et 17 ERR_SECTIONANA
               IF(EstVide(TD,'Lgsec1',ERR_LGSECTION) = FALSE) then
                  VerifLongueur(TD,'Lgsec1',17,3,ERR_LGSECTION);
               IF(EstVide(TD,'Lgsec2',ERR_LGSECTION) = FALSE) then
                  VerifLongueur(TD,'Lgsec2',17,3,ERR_LGSECTION);
               IF(EstVide(TD,'Lgsec3',ERR_LGSECTION) = FALSE) then
                  VerifLongueur(TD,'Lgsec3',17,3,ERR_LGSECTION);
               IF(EstVide(TD,'Lgsec4',ERR_LGSECTION) = FALSE) then
                  VerifLongueur(TD,'Lgsec4',17,3,ERR_LGSECTION);
               IF(EstVide(TD,'Lgsec5',ERR_LGSECTION) = FALSE) then
                  VerifLongueur(TD,'Lgsec5',17,3,ERR_LGSECTION);
               MTOB.Detail[i] := TD;
          END;
      //num compte general d'attente dans CGE
      //num compte client d'attente dans CAE
      //num compte FOURNISSEUR d'attente dans CAE
      //num compte salarie d'attente dans CAE
      //NUM COMPTE TIERS DANS CAE
     END;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Marion Koszczuk
Créé le ...... : 15/06/2004
Modifié le ... :   /  /
Description .. : Verif existance lignes PARAMG3
Mots clefs ... : CONTROL PARAMG3
*****************************************************************}
Procedure TVerif.VerifPS3(var MTOB:TOB);
begin
    VerifExistLigne(MTOB,ERR_PARAMETRE+'3'+ERR_PARAMETRESUITE,'Identifiant');
end;

{***********A.G.L.***********************************************
Auteur  ...... : Marion Koszczuk
Créé le ...... : 15/06/2004
Modifié le ... :   /  /
Description .. : Verif Existance lignes PARAMG5, champs obligatoires et
Suite ........ : format de certains champs
Mots clefs ... : CONTROL PARAMG5
*****************************************************************}
Procedure TVerif.VerifPS5(var MTOB:TOB);
var TD         : TOB;
    i,nbligne  : integer;
BEGIN
     TD      := TOB.Create('',nil,-1);
     nbligne := VerifExistLigne(MTOB,ERR_PARAMETRE+'5'+ERR_PARAMETRESUITE,'Identifiant');
     If (nbligne>0) then
     BEGIN
        For i := 0 to nbligne-1 do
        BEGIN
           TD := MTOB.Detail[i];
           //DEVISE
           IF(EstVide(TD,'Devise',ERR_DEVISE) = FALSE) then
           BEGIN
                IF ( ((Trim(TD.GetValue('Devise')) )<>'FRF') and ((Trim( TD.GetValue('Devise') ))<>'EUR') ) then
                begin
                   AlimTOBErreur(TD,'Devise',ERR_DEVISE2);
                end else
                begin
                   //verif q TENUEEURO=X si Devise=EUR
                   if((TD.GetValue('Devise')='EUR') and ( Trim (TD.GetValue('TenueEuro')) <>'X')) then
                   begin
                      AlimTOBErreur(TD,'TenueEuro',ERR_PARAMTENUE);
                   end else
                   begin
                        if ((TD.GetValue('Devise')<>'EUR') and  (Trim (TD.GetValue('TenueEuro'))<>'-') ) then
                            AlimTOBErreur(TD,'TenueEuro',ERR_PARAMTENUE);
                   end;
                end;
           END;
           //Verif existence du nombre de decimal de la devise de tenue
           EstVide(TD,'DecDevise',ERR_PARAMPS5);
           //TENUE MONNAIE
           IF(EstVide(TD,'TenueEuro',ERR_TENUE) = FALSE) then
              OUI_NON(TD,'TenueEuro',ERR_MONNAIETENUE);
           //VERIF PRESENCE JOURNAL ECAR DE CONVERSION
           EstVide(TD,'JalECC',ERR_JALPS5);  //VERIFIER AVEC JAL
           //Verif presence numero de compte utilisé au debit
           EstVide(TD,'CpteCDD',ERR_JALPS5);  //verif avec CGN
           //Verif presence numero de compte utilisé au credit
           EstVide(TD,'CpteCDC',ERR_JALPS5); //verif avec CGN
           //VERIF FORMAT GERER ANALYTIQU
           if EstVide(TD,'GereANA','')=FALSE then
              OUI_NON(TD,'GereANA',ERR_PS5_GEREANA);
           MTOB.Detail[i] := TD;
        END;
    END;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Marion Koszczuk
Créé le ...... : 25/05/2004
Modifié le ... : 03/06/2004
Description .. : CONTROL SUR LES LIGNE DE TYPE
Suite ........ : EXERCICE.Verification des champs de chaque ligne
Mots clefs ... : CONTROLE CHAMP EXO
*****************************************************************}
Procedure TVerif.VerifExercice(var MTOB:TOB);
var TD     : TOB;
    i,nbligne,nb  : integer;
    DateValid  : boolean;
    etat,DteFin: string;
BEGIN
     TD := TOB.Create('',nil,-1);
     nbligne := VerifExistLigne(MTOB,ERR_EXOMANQUANT,'Identifiant');
     If (nbligne>0) then
     BEGIN
          For i := 0 to nbligne-1 do
          BEGIN
               TD := MTOB.Detail[i];
               //présence du code exercice
               IF(EstVide(TD,'Code',ERR_CODE_EXERCICE) = FALSE) then
               Begin
                    nb := RechercheCode(TD.GetValue('Code'), MTOB,'Code');
                    if nb>1 then
                    begin
                         AlimTOBErreur(TD,'Code',ERR_CODE_DOUBLON);
                    end else
                    begin
                         //validité du code exercice
                         IF ((StrToInt(TD.GetValue('Code'))<001) or (StrToInt(TD.GetValue('Code'))>999)) then
                            AlimTOBErreur(TD,'Code',ERR_CODE_EXERCICE2);
                    End;
               end;

               //VERIF DU CHAMP DATE DEBUT
               IF (EstVide(TD,'DateDebut',ERR_DATEDEB_EXERCICE) = FALSE) then
               BEGIN
                    DateValid:=DATEVALIDE(TD,'DateDebut',ERR_EXERCICEDATEFIN);
                    IF ((DateValid=TRUE) and ((Copy(TD.GetValue('DateDebut'),1,2)<>'01'))) then
                        AlimTOBErreur(TD,'DateDebut',ERR_EXERCICEDATEDEB);
               END;
               //VERIF DU CHAMP DATE FIN
               DteFin := Copy(TD.GetValue('DateFin'),1,2);
               IF(EstVide(TD,'DateFin',ERR_DATEFIN_EXERCICE) = FALSE) then
               BEGIN
                     DateValid:=DATEVALIDE(TD,'DateFin',ERR_EXERCICEDATEFIN);
                     if ((DateValid=TRUE) and ((DteFin<>'30') and (DteFin<>'31')and (DteFin<>'28')and (DteFin<>'29')))then
                         AlimTOBErreur(TD,'DateFin',ERR_EXERCICEDATEFIN);
               END;

               //VERIF ETAT CPTA
               etat:=TD.GetValue('EtatCpta');
               IF (EstVide(TD,'EtatCpta',ERR_ETAT_INEXISTANT) = FALSE) then
               BEGIN
                    IF ((etat<>'OUV') and (etat<>'CDE') and (etat<>'CPR') and (etat<>'CLO') and (etat<>'NON')) then
                    begin
                       AlimTOBErreur(TD,'EtatCpta',ERR_EXERCICEETAT);
                    end else
                    begin
                         if(etat='OUV')then
                         begin
                              nb :=  RechercheCode('OUV', MTOB,'EtatCpta');
                              if nb>2 then
                                 AlimTOBErreur(TD,'EtatCpta',ERR_ETAT_EXERCICE);
                         end;
                     end;
               END;

               //VERIF ETATANO
               etat:=Trim(TD.GetValue('EtatAno'));
               IF ((etat<>'OAN') and (etat<>'H') and (etat<>'')) then
                   AlimTOBErreur(TD,'EtatAno',ERR_ETATANO_INEXIST);
               MTOB.Detail[i] := TD;
          END;
     END;
END;

PROCEDURE TVerif.VerifEtab(var MTOB:TOB);
var i,nbligne : integer;
    TD        : TOB;
begin
   nbligne := VerifExistLigne(MTOB,ENTETE_INEXSISTANT,'Code');
   if(nbligne>0)then
   begin
       FOR i := 0 to  nbligne-1 do
       begin
            TD := MTOB.Detail[i];
            //verif présence du code etablissement
           EstVide(TD,'Code',ERR_JRL_CODE);
           //verif presence du libelle
           EstVide(TD,'Libelle',ERR_JRL_LIBELLE);
           MTOB.Detail[i] := TD ;
       end;
   end;
end;


{***********A.G.L.***********************************************
Auteur  ...... : Marion Koszczuk
Créé le ...... : 26/05/2004
Modifié le ... :   /  /
Description .. : Controle des champ 'LETTRABLE' ,'POINTABLE', 'AXE
Suite ........ : Ventilable1,2,3,4,5' Complément de controle POUR
Suite ........ : COMPTES GENERAUX
Mots clefs ... : CONTROLE CMPT GNRX 2
*****************************************************************}
procedure TVerif.RechercheErrCmptGnrx(var TD:TOB);
BEGIN
     //champ lettrabl
     If(EstVide(TD,'Lettrabl',ERR_GENERAUX_CDLETTE) = FALSE) then//LETTRABLE
        OUI_NON(TD,'Lettrabl',ERR_GENERAUX_CDLETTE);
     //champ pointable
     If(EstVide(TD,'Pointabl',ERR_GENERAUX_CDPOINT) = FALSE) then//CONTROL CODE POINTABLE
        OUI_NON(TD,'Pointabl',ERR_GENERAUX_CDPOINT);
     //champ ventilable axe1
     If(EstVide(TD,'VentilAx1',ERR_GENERAUX_VENTIL1) = FALSE) then//CONTROL VENTILABLE AXE1
        OUI_NON(TD,'VentilAx1',ERR_GENERAUX_VENTIL1);
     //champ ventilable axe2
     If(EstVide(TD,'VentilAx2',ERR_GENERAUX_VENTIL2) = FALSE) then//CONTROL VENTILABLE AXE2
        OUI_NON(TD,'VentilAx2',ERR_GENERAUX_VENTIL2);
     //champ ventilable axe3
     If(EstVide(TD,'VentilAx3',ERR_GENERAUX_VENTIL3) = FALSE) then//CONTROL VENTILABLE AXE3
        OUI_NON(TD,'VentilAx3',ERR_GENERAUX_VENTIL3);
     //champ ventilable axe4
     If(EstVide(TD,'VentilAx4',ERR_GENERAUX_VENTIL4) = FALSE) then//CONTROL VENTILABLE AXE4
        OUI_NON(TD,'VentilAx4',ERR_GENERAUX_VENTIL4);
    //champ ventilable axe5
    If(EstVide(TD,'VentilAx5',ERR_GENERAUX_VENTIL5) = FALSE) then//CONTROL VENTILABLE AXE5
       OUI_NON(TD,'VentilAx5',ERR_GENERAUX_VENTIL5);
END;

{***********A.G.L.***********************************************
Auteur  ...... : Marion Koszczuk
Créé le ...... : 25/05/2004
Modifié le ... :   /  /
Description .. : CONTROL SUR LES LIGNES DE TYPE COMPTES GNRX
Mots clefs ... : CONTROL CPTE GNRX
*****************************************************************}
Procedure TVerif.VerifCmptGnrx(var MTOB:TOB);
var TD          : TOB;
    i,nbligne   : integer;
    Nature,sens : string;
BEGIN
     TD := TOB.Create('',nil,-1);
     nbligne := VerifExistLigne(MTOB,ERR_COMPTEGENEMANQ,'Identifiant');
     If (nbligne>0) then
     BEGIN
          for i := 0 to nbligne-1 do
          begin
               TD := MTOB.Detail[i];
               Nature := TD.GetValue('Nature');
               //controle de l'existannce du code
               EstVide(TD,'Code',ERR_GENERAUX_CODE);
               //control existance libelle compte
               EstVide(TD,'Libelle',ERR_GENERAUX_LIB);
               //CONTROLE DE LA NATURE DU COMPTE
               If (EstVide(TD,'Nature',ERR_GENERAUX_NATURE) = False) then
               BEGIN
                    If (VerifNature(Nature,TabNatureCmptGnrx,nbNatureCmptGnrx)= FALSE) then
                    begin
                         AlimTOBErreur(TD,'Nature',ERR_GENERAUX_NATURE);
                    end else
                    BEGIN
                         //CONTROLE CODE LETTRABLE AVEC NATURE TIC OU TID
                         If(EstVide(TD,'Lettrabl',ERR_GENERAUX_CDLETTE) = FALSE) then
                         Begin
                              IF ((TD.GetValue('Lettrabl')='X') and ((TD.GetValue('Nature')<>'TIC') and (TD.GetValue('Nature')<>'TID')))then
                                  AlimTOBErreur(TD,'Lettrabl',ERR_GENERAUX_LETTR);
                         end;
                   end;
               END;
               //control Sens du compte
               sens := Trim(TD.GetValue('Sens'));

               If((sens<>'')and(Sens<>'D') and (Sens<>'C') and (Sens<>'M')) then
                  AlimTOBErreur(TD,'Sens',ERR_GENERAUX_SENS);
              RechercheErrCmptGnrx(TD);
              MTOB.Detail[i] := TD;
         END;
     End;
END;


{***********A.G.L.***********************************************
Auteur  ...... : Marion Koszczuk
Créé le ...... : 02/06/2004
Modifié le ... :   /  /
Description .. : Control sur les champ de type Compte Tiers
Mots clefs ... : CONTROL COMPTE TIER
*****************************************************************}
procedure TVerif.VerifCmpTiers(var MTOB:TOB);
var TD         : TOB;
    i,nbligne  : integer;
Begin
     TD := TOB.Create('',nil,-1);
     nbligne := VerifExistLigne(MTOB,ERR_COMPTEAUXIMANQ,'libelle');
     if (nbligne>0) then
     BEGIN
        for i := 0 to nbligne-1 do
        begin
            TD := MTOB.Detail[i];
            //VERIF Presence Code compte auxiliare
            EstVide(TD,'Code',ERR_CMPTTIER_CODE);
            //Verif Presence Libelle
            EstVide(TD,'libelle',ERR_CMPTTIER_LIBEL);
            //VERIF NATURE
            if(EstVide(TD,'Nature',ERR_CMPTTIER_NATURE)=FALSE)then
            begin
                 if ((TD.GetValue('Nature')<>'AUC') and (TD.GetValue('Nature')<>'AUD') and (TD.GetValue('Nature')<>'CLI') and (TD.GetValue('Nature')<>'DIV') and (TD.GetValue('Nature')<>'FOU') and (TD.GetValue('Nature')<>'SAL')) THEN
                     AlimTOBErreur(TD,'Nature',ERR_CMPTTIER_NATURE);
            END;
            //verif lettrable
            if (EstVide(TD,'Lettrabl',ERR_CMPTTIER_LETTRA) = FALSE) then//CONTROL Letrable
               OUI_NON(TD,'Lettrabl',ERR_CMPTTIER_LETTRA);
            if(EstVide(TD,'Collectif',ERR_CMPTTIER_CGN)=FALSE)then
            begin
                //verif avec CGN ERR_CMPTTIER_CGN2
            end;
            //verif Adress1
            EstVide(TD,'Adr1',ERR_CMPTTIER_ADR1);
            //verif Adress2
            EstVide(TD,'Adr2',ERR_CMPTTIER_ADR2);
            //verif domiciliation
            if (EstVide(TD,'Domiciliation',ERR_CMPTTIER_DOMICIL)) then
            begin
               //  verif dans RIB
            end;
            //verif etablissement bancaire
            if (EstVide(TD,'Etablssmnt',ERR_CMPTTIER_ETAB)) then
            begin
                //verif dans rib correspond au 1er element
            end;
            //verif guichet
            if (EstVide(TD,'Guichet',ERR_CMPTTIER_GUICH))then
            begin
               //verif dans rib correspond au 2eme element
            end;
            //verif numero de compte
            if(EstVide(TD,'NumCpt',ERR_CMPTTIER_CMPT))then
            begin
               //verif dans rib correspond au 3eme element
            end;
            //verif cle du rib
            if(EstVide(TD,'Cle',ERR_CMPTTIER_CLE))then
            begin
               //verif dans rib correspond au 4eme element
            end;
            //verif q le pay est saisi
            EstVide(TD,'Pay',ERR_CMPTTIER_PAY);
            //verif format multidevise
            if(EstVide(TD,'MultiDevise','')=FALSE)then
               OUI_NON(TD,'MultiDevise',ERR_CMPTTIER_MDEV);
            //verif q le regime tva est saisi
            EstVide(TD,'RegimTVA',ERR_CMPTTIER_RTVA);
            //verif q le mod de reglement est saisi
            if(EstVide(TD,'ModReglmnt',ERR_CMPTTIER_MRGLMT)=FALSE)then
            begin
                 //verif avec code dans MDR
            end;
            //verif X ou - du contact principal
            if(EstVide(TD,'CPpal','')=FALSE)then
               OUI_NON(TD,'CPpal',ERR_CMPTTIER_CPPAL);
            //verif X ou - du RIB principal
            if(EstVide(TD,'RibPpal','')=FALSE)then
               OUI_NON(TD,'RibPpal',ERR_CMPTTIER_RIBPAL);
            MTOB.Detail[i]:=TD;
        end;
     END;
end;


{***********A.G.L.***********************************************
Auteur  ...... : Marion Koszczuk
Créé le ...... : 25/05/2004
Modifié le ... :   /  /
Description .. : CONTROL SUR LES LIGNES DE TYPE JOURNAUX
Mots clefs ... : CONTROL LIGNE JOURNAUX
*****************************************************************}
procedure TVerif.VerifJournaux(var MTOB:TOB);
var TD         : TOB;
    i, nbligne : integer;
    Nature     : string;
BEGIN
     TD := TOB.Create('',nil,-1);
     nbligne := VerifExistLigne(MTOB,ERR_JALMANQUANT,'Identifiant');
     If (nbligne>0) then
     BEGIN
          For i := 0 to nbligne-1 do
          begin
               TD := MTOB.Detail[i];
               //VERIF PRESENCE CODE
               EstVide(TD,'CODE',ERR_JOURN_CODE);
               //VERIF PRESENCE LIBELLE
               EstVide(TD,'Libelle',ERR_JRL_LIBELLE);
               Nature := Trim(TD.GetValue('Nature'));
               //VERIF NATURE
               If (EstVide(TD,'Nature',ERR_JOURN_NATURE1) = FALSE) then
               BEGIN
                    IF (VerifNature(Nature,TabNatureJournaux,nbNatureJournaux)= FALSE) then
                    begin
                       AlimTOBErreur(TD,'Nature',ERR_JOURN_NATURE2);
                    end else
                    begin
                         //verif la presence d'un num compte si nature=BQE ou CAI
                         if (((Nature='BQE') or (Nature='CAI'))and (EstVide(TD,'Compte','')=TRUE)) then
                             AlimTOBErreur(TD,'Compte',ERR_JOURN_CPT);
                         //verif presence axe analytiq pour nature=ODA
                         if((Nature='ODA') and  (EstVide(TD,'Axe','')=TRUE))then
                            AlimTOBErreur(TD,'Axe',ERR_JOURN_AXE);
                    end;
               END;
               //VERIF MODE DE SAISIE
               IF (EstVide(TD,'Mode Saisie','') = FALSE)then
               BEGIN
                    If ((Trim(TD.GetValue('Mode Saisie'))<>'-') and (TD.GetValue('Mode Saisie')<>'BOR') and (TD.GetValue('Mode Saisie')<>'LIB')) then
                        AlimTOBErreur(TD,'Mode Saisie',ERR_JOURN_MODSAISI);
               END;
               MTOB.Detail[i] := TD;
          END;
    END;
END;


{***********A.G.L.***********************************************
Auteur  ...... : Marion Koszczuk
Créé le ...... : 25/05/2004
Modifié le ... :   /  /
Description .. : Control sur les lignes de type Ecriture
Mots clefs ... : CONTROL ECRITURE
*****************************************************************}
procedure TVerif.VerifEcriture(var Enreg:array of AffComEnreg ; i : integer);
var   TD,TOBM,TOBJ,TOBTmp,TOBC,TOBE : TOB;
      j, nbligne1                   : integer;
      typeEcrit,valcode             : string;
begin
     TOBM := TOB.Create('',nil,-1);
     TOBM := Enreg[i].MTOB;
     TOBJ := RechercheLigne(Enreg,JAL);  //recherche la tob mere des lignes journal
     TOBC := RechercheLigne(Enreg,CGN); // recherche la tob mere des lignes compt generaux
     nbligne1 := VerifExistLigne(TOBM,ERR_INEXISTANT,'Code');
     If (nbligne1>0) then
     BEGIN
          for j := 0 to nbligne1-1 do
          begin
               TD := TOBM.Detail[j];
              //VERIF QUE LE JOURNAL EXISTE
              if((TOBJ<>nil)and (TOBJ.Detail.Count>0))then
              begin
                   TOBTmp := TOBJ.FindFirst(['Code'],[TD.GetValue('Code')],TRUE);
                   if (TOBTmp<>nil) then
                   begin
                       if ((TOBTmp.GetValue('Mode Saisie')='BOR') or (TOBTmp.GetValue('Mode Saisie')='LIB'))then
                       begin
                            if(TD.GetValue('TypeEcriture')<>'N')then
                               AlimTOBErreur(TD,'TypeEcriture',ERR_MODSAIS_JOURN);
                       end;
                   end;
              end;
              //VERIF COMPTE GENERAL EXISTE
              if((TOBC<>nil)and (TOBC.Detail.Count>0) and(EstVide(TD,'General',ERR_CMPT_GNRL)=FALSE))then
              begin
                   TOBTmp := TOBC.FindFirst(['Code'],[TD.GetValue('General')],TRUE);
                   if TOBTmp=nil then
                   begin
                       AlimTOBErreur(TD,'General',ERR_CMPT_GNRL);
                   end else
                   begin
                        if (TD.GetValue('TypeCpte')='X') then
                        begin
                           if ((TOBTmp.GetValue('Nature')<>'COS') and (TOBTmp.GetValue('Nature')<>'COF') and (TOBTmp.GetValue('Nature')<>'COD') and (TOBTmp.GetValue('Nature')<>'COC'))then
                           begin
                                if TOBTmp=nil then
                                   AlimTOBErreur(TD,'General',ERR_CMPTGNR_NAT);
                           end;
                        end;
                   end;
              end;
              valcode:=TD.GetValue('Code');
              TOBE := RechercheLigne(Enreg,ENT); //recherch lign entete
              if ((TOBE<>nil) and (TOBE.Detail.Count>0)) then
              begin
                   if ((Copy(TOBE.Detail[0].GetValue('Identifiant'),1,1)='!') or (Trim(TOBE.GetValue('Type Fichier'))='JRL')) then
                   begin
                        //verif Journal  ANO et ECRANOUVEAU OAN
                        if (Trim(TD.GetValue('Code'))='ANO') then
                        begin
                             if (Trim(TD.GetValue('Nature mouvement'))='OAN') then
                                AlimTOBErreur(TD,'Nature mouvement',ERR_ANOUVEAU);
                        end;
                   end;
              end;
               //verif date comptable
               if(EstVide(TD,'DteCptable',ERR_DTECMPABL)=FALSE)then
                  DATEVALIDE(TD,'DteCptable',ERR_DTECMPABL);
               //verif  TYPE PIECE
               if(EstVide(TD,'TypePiece',ERR_TYPPIEC)=FALSE)then
               begin
                   if(VerifNature(TD.GetValue('TypePiece'),TabTypePieceEcr,nbTypePieceEcr)= FALSE)then
                      AlimTOBErreur(TD,'TypePiece',ERR_TYPPIEC2);
               end;

              //verif sens du mouvement
              if(EstVide(TD,'Sens',ERR_SENS1)=FALSE)then
              begin
                   if((TD.GetValue('Sens')<>'D') and (TD.GetValue('Sens')<>'C')) then
                     AlimTOBErreur(TD,'Sens',ERR_SENS2);
               end;
               //verif date echeance
               if(EstVide(TD,'Echeance','')=FALSE)then
                   DATEVALIDE(TD,'Echeance',ERR_DTEECH);
               //verif montant 1
//pb d'affichage sur toberreur erreur convert
//             if (TD.GetValue('Montant1')=0)then
//                  AlimTOBErreur(TD,'Montant1',ERR_MONTANT1);

               //verif premier caractere code montant
               if(EstVide(TD,'CodeMontant',ERR_CODEMONTANT)=FALSE)then
               begin
                   if (Trim(Copy(TD.GetValue('CodeMontant'),1,1))='')then
                       AlimTOBErreur(TD,'CodeMontant',ERR_CODEMONTANT);
               end;
               //Verif type ecriture
               typeEcrit := Trim(TD.GetValue('TypeEcriture'));
               if(EstVide(TD,'TypeEcriture',ERR_TYPE_ECR)=FALSE)then
               begin
                    if((typeEcrit<>'N') and (typeEcrit<>'S') and (typeEcrit<>'A') and (typeEcrit<>'U') and (typeEcrit<>'R'))then
                        AlimTOBErreur(TD,'TypeEcriture',ERR_TYPE_ECR);
               end;
               typeEcrit := TD.GetValue('Typecpte'); // ajout me
               //Verif Axe
               if((typeEcrit='A') or (typeEcrit='H'))then
               begin
                    if(EstVide(TD,'Axe',ERR_AXE_ECR)=TRUE)then
                       AlimTOBErreur(TD,'Axe',ERR_AXE_ECR);
               end
               else   // ajout me
               //Verif Presence EtatLettrage
               VerifLettrage(TD);
              TOBM.Detail[j] := TD;
           end;
     end;
end;

function TVerif.VerifLettrage(var TD:TOB):boolean;
var EtatLet  : string;
    vraifaux : boolean;
begin
   vraifaux := TRUE;
   EtatLet  := Trim(TD.GetValue('EtaLettrag'));
   if(EstVide(TD,'EtaLettrag',ERR_MOUVMNT)=FALSE) THEN
   begin
        if ((EtatLet<>'RI') and (EtatLet<>'AL') and (EtatLet<>'PL') and (EtatLet<>'TL')) then
        begin
           AlimTOBErreur(TD,'EtaLettrag',ERR_MOUVMNT);
           vraifaux := FALSE;
        end else
        begin
            If ((CompareStr('AAAA',TD.GetValue('Lettrage'))=0) and (EtatLet<>'TL')) then
            begin
                 AlimTOBErreur(TD,'EtaLettrag',ERR_MOUVMNT_CODE);
            end else
            begin
                 if ((CompareStr('aaAA',TD.GetValue('Lettrage'))=0) and (EtatLet<>'PL')) then
                     AlimTOBErreur(TD,'EtaLettrag',ERR_MOUVMNT_CODE);
            end;
        end;
   end else
       VraiFaux := FALSE;
  Result := VraiFaux;
end;


{***********A.G.L.***********************************************
Auteur  ...... : Marion Koszczuk
Créé le ...... : 15/06/2004
Modifié le ... :   /  /    
Description .. : Verifie que la valeur de cetains champs soit compris entre 
Suite ........ : deux valeurs entière. Appel à la fonction ValidBorne
Mots clefs ... : VERIF VALEUR LONGUEUR
*****************************************************************}
FUNCTION TVerif.VerifLongueur(TD : TOB ; Champ : string; borne1,borne2:integer;MsgErr:string):boolean;
var val    : integer;
    chaine : string;
begin
     chaine := Trim(TD.GetValue(Champ));
     val := StrToInt(chaine);
     if ValidBorne(val,borne1,borne2)=FALSE THEN
     begin
         AlimTOBErreur(TD,Champ,MsgErr);
         Result := FALSE;
     end else
         Result := TRUE;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Marion Koszczuk
Créé le ...... : 26/05/2004
Modifié le ... :   /  /
Description .. : Verifie si la nature du champ passée en parametre
Suite ........ : correspond à une chaine de caractere autoriséé pour ce
Suite ........ : champ et contenu dans un tableau.
Suite ........ : Peut s'adapter a d'autre champ q Nature
Mots clefs ... : VERIF NATURE
*****************************************************************}
FUNCTION TVerif.VerifNature(champ:string;TabNature:array of String;nb:integer):boolean;
var i     : integer;
Begin
   i := 0;
   while ((i<=nb-1) and ((Trim(champ)) <> TabNature[i])) do
   begin
        i := i+1;
   end;
   if ((i<=nb-1)and ((Trim(champ))=TabNature[i]))  then
   begin
        Result := TRUE;
   end else
        Result := FALSE;
END;

{***********A.G.L.***********************************************
Auteur  ...... : Marion Koszczuk
Créé le ...... : 17/05/2004
Modifié le ... :   /  /
Description .. : Fonction qui retourne soit la chaine passée en param en
Suite ........ : enlevant les espaces si c'est une chaine vide soit la chaine
Suite ........ : telle quelle sinon
Mots clefs ... :
*****************************************************************}
function TVerif.EnleveVide(champ:string) : string;//MARCHE
var varTrim : string;
begin
     varTrim := Trim(champ);
     if varTrim='' then
        champ:=varTrim;
     Result := champ;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Marion Koszczuk
Créé le ...... : 25/05/2004
Modifié le ... :   /  /
Description .. : Fonction qui teste si le champ de la TOB est vide et ajoute
Suite ........ : le message d'erreur si c''est le cas dans le champ concerné.
Mots clefs ... : EST VIDE, CHAMP MANQUANT
*****************************************************************}
Function TVerif.EstVide(var TD:TOB ; Champ:string ; Erreur:string):boolean;
begin
     IF (EnleveVide(TD.GetValue(Champ))='' ) then
     BEGIN
        if(Erreur<>'')then
             AlimTOBErreur(TD,Champ,Erreur);
        Result := TRUE;
     END ELSE
         Result := FALSE;
end;


{***********A.G.L.***********************************************
Auteur  ...... : Marion Koszczuk
Créé le ...... : 26/05/2004
Modifié le ... :   /  /
Description .. : FONCTION QUI TESTE SI LE CHAMP EST X OU - C'EST
Suite ........ : à DIRE SI Oui (X) ou Non(-)
Mots clefs ... : X OU -
*****************************************************************}
Procedure TVerif.OUI_NON(var TD:TOB ;Champ:string ; Erreur:string);
Begin
     IF( (TD.GetValue(Champ)<>'X') and (TD.GetValue(Champ)<>'-') ) then
       AlimTOBErreur(TD,Champ,Erreur);
END;

{***********A.G.L.***********************************************
Auteur  ...... : Marion Koszczuk
Créé le ...... : 02/06/2004
Modifié le ... : 15/06/2004
Description .. : Recherche la tob Mere correspondant au type passé en 
Suite ........ : parametre.
Mots clefs ... : RECHERCHE TYPE LIGNE 
*****************************************************************}
function  TVerif.RechercheLigne(Enreg: array of AffComEnreg ; typeRecherch : TTypeLigne ):TOB;
var indice,nbEnreg,trouv    : integer;
begin
   indice := 0;
   trouv  := 0;
   nbEnreg := Length(Enreg);
   while ((indice<nbEnreg-1) and (trouv=0) ) do
   begin
        if(Enreg[indice].Verif=typeRecherch) then
        begin
             trouv:=1;
        end else
           indice := indice+1;
   end;
   if (trouv=1) then
   begin
        Result := Enreg[indice].MTOB;
   end else
       Result := nil;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Marion Koszczuk
Créé le ...... : 27/05/2004
Modifié le ... :   /  /
Description .. : function qui test la validité de deux nombre en les
Suite ........ : comparant aux bornes les concernat.
Suite ........ : Ex: un jour compris entre 1 et 31 sera testé avec borne1=1
Suite ........ : et borne2=31
Mots clefs ... : TEST valeur entre 2 bornes
*****************************************************************}
function TVerif.ValidBorne(test:integer;borne1:integer;borne2:integer):boolean;
var Valid : boolean;
begin
   Valid := TRUE;
   if( (test>borne1) or (test<borne2) ) then
        Valid := FALSE;
   Result := Valid;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Marion Koszczuk
Créé le ...... : 27/05/2004
Modifié le ... :   /  /
Description .. : Fonction qui test la validité du jour par rapport au mois.
Suite ........ : Ex: le mois de février n'a jamais 30 jours ou le mois d'avril
Suite ........ : jamais 31.
Mots clefs ... : VALIDE JOUR MOIS
*****************************************************************}
function TVerif.ValidMoisJour(annee:integer;mois:integer;jour:integer):boolean;
var Valid    : boolean;
    bisextil : INTEGER;
BEGIN
    Valid := TRUE;
    Case mois of
        02: begin
               bisextil := 0;
               if ( (annee mod 4 = 0) or (annee mod 400 = 0)) then
                   bisextil := 1;
               if ( (bisextil = 1) and (jour>29) ) then
               begin
                    Valid := FALSE;
                    Result := Valid;
                    Exit;
               End;
               If( (jour>28) and (bisextil=0) ) then
               begin
                   Valid := FALSE;
                   Result := Valid;
                   exit;
               end;
           end;
 04,06,09,11: begin
                if jour>30 then
                begin
                    Valid := FALSE;
                    Result := Valid;
                    exit;
                end;
             end;
     END;
     Result:=Valid;
END;

{***********A.G.L.***********************************************
Auteur  ...... : Marion Koszczuk
Créé le ...... : 27/05/2004
Modifié le ... :   /  /
Description .. : Recup le jour, mois , année de la date passée en param, les 
Suite ........ : converti en Int et appel les fonction ValidBorne et
Suite ........ : ValidMoisJour
Mots clefs ... : CONTROL DATE
*****************************************************************}
FUNCTION TVerif.ControlDate(date:string):boolean;
var Valid             : boolean;
    dateTmp           : PChar;
    jour,mois,annee,i : integer;
begin
    Valid := TRUE;
    if (length(date)<>8)then
    begin
       Valid  := FALSE;
       Result := Valid;
       exit;
    end else
        begin
             datetmp := PChar(date);
             for i:=0 to 7 do
             begin
                  if (not (datetmp[i] in ['0'..'9'])) then
                  begin
                     Result := False;
                     exit;
                  end;
             end;
             jour  := StrToINT(Copy(date,1,2));
             mois  := StrToINT(Copy(date,3,2));
             annee := StrToINT(Copy(date,5,4));
             IF ValidBorne(jour,31,1)=False then
             BEGIN
                Result := False;
                EXIT;
             END ELSE
             BEGIN
                 if ValidBorne(mois,12,1)=FALSE then
                 BEGIN
                      Result := False;
                      EXIT;
                 end else
                 BEGIN
                    if ValidMoisJour(annee,mois,jour)=FALSE then
                        Valid := FALSE;
                 END;
             END;
         END;
    Result := Valid;
END;

{***********A.G.L.***********************************************
Auteur  ...... : Marion Koszczuk
Créé le ...... : 15/06/2004
Modifié le ... :   /  /    
Description .. : Verif la validité de la date si <>01011900 en appelant la 
Suite ........ : fonction Control Date
Mots clefs ... : VALIDATION DATE
*****************************************************************}
FUNCTION TVerif.DATEVALIDE(var TD :TOB ; Champ : string ; MsgErreur : string) : boolean;
var date : string;
begin
     date := TD.Getvalue(Champ);
     if date<>'01011900' then
     begin
          if (ControlDate(date)=FALSE) then
          begin
                AlimTOBErreur(TD,Champ,MsgErreur);
                Result := FALSE;
          end else
              Result := TRUE;
     end else
         Result := TRUE;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Marion Koszczuk
Créé le ...... : 15/06/2004
Modifié le ... :   /  /    
Description .. : Recherche a l'interieur d'une tob une valeur d'un champ , 
Suite ........ : tous deux passé en paramètre.
Suite ........ : On renvoi le nombre de fois que ce champ a été trouvé.
Mots clefs ... : RECHERCHE EXISTE VALEUR DANS CHAMPS
*****************************************************************}
FUNCTION TVerif.RechercheCode(VALEUR :string ; MTOB:TOB ; Champ: string):integer;
var TOBTMP : TOB;
    nb     : integer;
begin
   nb := 0;
   TOBTMP := MTOB.FindFirst([Champ],[VALEUR],TRUE);
   while TOBTMP<>nil do
   begin
         nb := nb+1;
         TOBTMP := MTOB.FindNext([Champ],[VALEUR],TRUE);
   end;
   Result := nb;
end;

end.
