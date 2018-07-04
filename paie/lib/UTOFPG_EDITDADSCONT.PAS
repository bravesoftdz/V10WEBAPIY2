{***********UNITE*************************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 07/03/2006
Modifié le ... :   /  /    
Description .. : Source TOF de la FICHE : PG_EDITDADSCONT ()
Mots clefs ... : TOF;PG_EDITDADSCONT
*****************************************************************}
{
PT1   : 07/11/2006 VG V_70 Correction gestion du décalage
PT2   : 14/11/2006 VG V_70 Les données de l'établissement n'étaient pas
                           réinitialisées entre chaque calcul - FQ N°13684
PT3   : 01/12/2006 VG V_70 Correction affichage
PT4   : 03/01/2007 VG V_72 Etat vide si DADS-U TDS seule - FQ N°13804
PT5   : 11/01/2007 VG V_72 Suppression des enregistrements ayant le champ
                           exercicedads mal alimenté - FQ N°13827
PT6   : 30/01/2007 FC V_80 Mise en place filtrage des habilitations/populations
                           correction le 27/03/07
PT7-1 : 03/09/2007 VG V_80 Adaptation cahier des charges V8R05
PT7-2 : 26/09/2007 VG V_80 Ajout contrôle de cohérence entre le code NAF et le
                           code section prud'homale - FQ N°13766
PT8   : 05/11/2007 VG V_80 Adaptation cahier des charges V8R06
PT9   : 28/11/2007 VG V_80 Gestion du champ ET_FICTIF - FQ N°13925
PT10-1: 03/12/2007 VG V_80 Prise en compte du trimestre civil dans le cas d'une
                           déclaration trimestrielle - FQ N°13245
PT10-2: 03/12/2007 VG V_80 Amélioration traitement - FQ N°13648
PT10-3: 03/12/2007 NA V_80 FQ 13623 Edition de l'exercice sur l'état  + correction plantage si honoraire
}
Unit UTOFPG_EDITDADSCONT ;

Interface

Uses Classes,
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
     FE_Main,
     qrs1,
{$ELSE}
     MaineAgl,
     eMul,
     eqrs1,
{$ENDIF}
     uTob,
     sysutils,
     HCtrls,
     UTOF,
     PgDADSControles,
     PgDADSCommun,
     PgDADSOutils,
     HTB97,
     paramsoc,
     pgoutils2,
     HMsgBox,
     EntPaie,
     StdCtrls,
     HEnt1,
     ed_tools,
     HStatus,
     HDebug,
     Controls,
     P5Def,
     uTobDebug;

Type
  TOF_PG_EDITDADSCONT = Class (TOF)
  public
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  private
    Arg, PerDeb, PerFin : string;
    Car, Nature, THAnneeRef : THValCombobox;
    Neant : TCheckBox;
    SALAR : THEdit;

    procedure InitControle;
    procedure CalculEntreprise;
    procedure RemplitTOBErreur;
    procedure ControleDossier;
    procedure ChargerLexiqueDads (Annee, NatureDADSU:String);
    function AffecteDebut(T : TOB) : integer;
    function AffecteFin(T : TOB) : integer;
    procedure ControleStandard;
    procedure DecalageChange(Sender: TObject);
    procedure DateDecaleChange(Sender: TObject);
    procedure DateDecalageChange(Sender: TObject);
    procedure Parametrage ();
    procedure CarDateChange (Sender: Tobject);
    procedure NatureChange (Sender: Tobject);
    procedure NeantChange (Sender: TObject);
    procedure CreeEtab();
    procedure SalarieExit (Sender: TObject);

  end;

Implementation


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 07/03/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... : PAYE;PGDADSU;PGCONTROLE
*****************************************************************}
procedure TOF_PG_EDITDADSCONT.OnNew;
begin
Inherited ;
end ;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 07/03/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... : PAYE;PGDADSU;PGCONTROLE
*****************************************************************}
procedure TOF_PG_EDITDADSCONT.OnDelete;
begin
Inherited ;
end ;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 07/03/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... : PAYE;PGDADSU;PGCONTROLE
*****************************************************************}
procedure TOF_PG_EDITDADSCONT.OnUpdate;
begin
Inherited;
if (Arg <> 'S') then
   Parametrage;
ChargeTOBDADS;
FreeAndNil (TErreur);
TErreur:= TOB.Create ('DADSCONTROLE', Nil, -1);
try
   begintrans;
   if (Arg='D') then                      //Contrôle dossier avant préparation
      begin
      InitMoveProgressForm (NIL,'Edition en cours', 'Veuillez patienter SVP ...',
                            1000,FALSE,TRUE);
      InitMove(1000, '');
      DeleteErreur ('', 'SOK');
      DeleteErreur ('', 'SKO');
      InitControle;
      CalculEntreprise;
      if (GetCheckBoxState ('CBRECONTROL')=CbChecked) then
         ControleDossier;
      RemplitTOBErreur;
      TErreur.Detail.Sort ('PSU_ORIGINECRIT;PSU_SALARIE;PSU_ORDRE;PSU_SEGMENT');
//      TobDebug (TErreur);
      end
   else                                   //Contrôle standard avant envoi
      begin
      ControleStandard;
      end;
   CommitTrans;
except
   Rollback;
   end;

TFQRS1(Ecran).LaTob:= TErreur;
FiniMove;
FiniMoveProgressForm;
//TobDebug (TDADSUD);
FreeTobDADS;
FreeAndNil (TEtab);
FreeAndNil (TLexique);
end ;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 07/03/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... : PAYE;PGDADSU;PGCONTROLE
*****************************************************************}
procedure TOF_PG_EDITDADSCONT.OnLoad;
begin
Inherited ;
end ;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 07/03/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... : PAYE;PGDADSU;PGCONTROLE
*****************************************************************}
procedure TOF_PG_EDITDADSCONT.OnArgument (S : String );
var
ModifDateDecale : TToolbarButton97;
Decal : THValComboBox;
Decalage : THEdit;
JourJ : TDateTime;
AnneeA, Jour, MoisM : Word;
AnneeE, AnneePrec, ComboExer, MoisE, StPlus : string;
Num : integer;
begin
Inherited ;
Arg:= S;

SetControlText ('ARG', Arg);
if (Arg='S') then
   begin
   SetControlVisible ('TANNEE', False);
   SetControlVisible ('ANNEE', False);
   SetControlVisible ('TDECALAGE', False);
   SetControlVisible ('DECALAGE', False);
   SetControlVisible ('LNATURE', False);
   SetControlVisible ('CNATURE', False);
   SetControlVisible ('LCAR', False);
   SetControlVisible ('CCAR', False);
   SetControlVisible ('LETAB', False);
   SetControlVisible ('MCETAB', False);
   SetControlVisible ('L_DU', False);
   SetControlVisible ('DATEDEB', False);
   SetControlVisible ('L_AU', False);
   SetControlVisible ('DATEFIN', False);
   SetControlVisible ('BMODIFDATE', False);
   SetControlVisible ('TDADSFRACTION', False);
   SetControlVisible ('DADSFRACTION', False);
   SetControlVisible ('CBRECONTROL', False);
   SetControlVisible ('CHNEANT', False);
   SetControlVisible ('CHPREVOYANCE', False);
   SetControlVisible ('CHHONORAIRE', False);
   SetControlVisible ('CBPRUDH', False);
//PT3
   SetControlVisible ('CBBLOQUANT', False);
   SetControlVisible ('CHCOMPL', False);
//FIN PT3   
   SetControlProperty ('PSALARIE', 'TabVisible', False);
   SetControlProperty ('PCOMPLEMENT', 'TabVisible', False);
   SetControlVisible ('TFICHIER', True);
   SetControlVisible ('FICHIER', True);
   end
else
   begin
   THAnneeRef:= THValCombobox(GetControl('ANNEE'));
   SetControlText ('DOSSIER', GetParamSoc ('SO_LIBELLE'));
   JourJ:= Date;
   DecodeDate (JourJ, AnneeA, MoisM, Jour);
   if MoisM>9 then
      AnneePrec:= IntToStr(AnneeA)
   else
      AnneePrec:= IntToStr(AnneeA-1);

   if RendExerSocialPrec (MoisE, AnneeE, ComboExer, DebExer, FinExer, AnneePrec) = TRUE then
      THAnneeRef.Value:= ComboExer
   else
      PGIBox ('L''exercice '+AnneePrec+' n''existe pas', Ecran.Caption);

   SetControlText ('DATEDEB', DateToStr(DebExer));
   SetControlText ('DATEFIN', DateToStr(FinExer));
   if (VH_Paie.PGDecalage) then
      SetControlText ('DECALAGE', '03')
   else
      SetControlText ('DECALAGE', '01');
   SetControlText ('CNATURE', '0151');
   SetControlText ('CCAR', 'A00');
   SetControlText ('DADSFRACTION', '1');
   SetControlText ('MCETAB', '<<Tous>>');

   If THAnneeRef <> NIL then
      THAnneeRef.OnChange:= CarDateChange;

   Nature:= THValComboBox(GetControl('CNATURE'));
   if Nature <> nil then
      Nature.OnChange:= NatureChange;

   Car:= THValComboBox(GetControl('CCAR'));
   if Car <> nil then
      Car.OnChange:= CarDateChange;

// Gestion du bouton de modification des dates
   ModifDateDecale:= TToolbarButton97(GetControl('BMODIFDATE'));
   if ModifDateDecale<>nil then
      ModifDateDecale.OnClick:= DateDecaleChange;

   Decal:= THValComboBox(GetControl('DECALAGE'));
   if Decal <> nil then
      Decal.OnChange:= DecalageChange;

   Decalage:= THEDIT(GetControl('DATEDECALAGE'));
   if Decalage <> nil then
{PT1
      Decalage.OnChange:= DateDecalageChange;
}
      Decalage.OnExit:= DateDecalageChange;
//FIN PT1

   Neant:= TCheckBox (GetControl ('CHNEANT'));
   if (Neant<> nil) then
      Neant.OnClick:= NeantChange;

   Parametrage;
   end;
SetControlChecked ('CBRECONTROL', True);

SALAR:= THEdit (GetControl ('PSA_SALARIE'));
if SALAR <> nil then
   SALAR.OnExit:= SalarieExit;

for Num := 1 to VH_Paie.PGNbreStatOrg do
    begin
    if Num > 4 then
       Break;
    VisibiliteChampSalarie (IntToStr (Num),
                            GetControl ('PSA_TRAVAILN'+IntToStr (Num)),
                            GetControl ('TPSA_TRAVAILN'+IntToStr (Num)));
    end;
VisibiliteStat (GetControl ('PSA_CODESTAT'), GetControl ('TPSA_CODESTAT'));

MajExercice;   //PT5

//PT9
StPlus:= ' WHERE ET_FICTIF<>"X"';
SetControlProperty ('MCETAB', 'Plus', StPlus);
//FIN PT9
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 07/03/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... : PAYE;PGDADSU;PGCONTROLE
*****************************************************************}
procedure TOF_PG_EDITDADSCONT.OnClose;
begin
Inherited ;
FreeAndNil (TErreur);
end ;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 07/03/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... : PAYE;PGDADSU;PGCONTROLE
*****************************************************************}
procedure TOF_PG_EDITDADSCONT.OnDisplay ();
begin
Inherited ;
end ;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 07/03/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... : PAYE;PGDADSU;PGCONTROLE
*****************************************************************}
procedure TOF_PG_EDITDADSCONT.OnCancel ();
begin
Inherited ;
end ;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 31/03/2006
Modifié le ... :   /  /
Description .. : Chargement des données nécessaires au contrôle
Mots clefs ... : PAYE;PGDADSU;PGCONTROLE
*****************************************************************}
procedure TOF_PG_EDITDADSCONT.InitControle ();
var
StLexique, StTypeDecla, TD2 : string;
QRechLexique : TQuery;
begin
PGExercice:= RechDom ('PGANNEESOCIALE', GetControlText('ANNEE'), False);
TD2:= Copy (GetControlText ('CNATURE'),1,2);
StTypeDecla:= GetDADSUNature (TD2);
StLexique:= 'SELECT *'+
            ' FROM DADSLEXIQUE WHERE'+
            ' PDL_EXERCICEDEB <= "'+PGExercice+'" AND'+
            ' (PDL_EXERCICEFIN >= "'+PGExercice+'" OR'+
            ' PDL_EXERCICEFIN = "")'+StTypeDecla+
            ' ORDER BY PDL_DADSSEGMENT';
QRechLexique:= OpenSql (StLexique, TRUE);
TLexique:= TOB.Create ('DADSLEXIQUE', NIL, -1);
TLexique.LoadDetailDB ('DADSLEXIQUE','','',QRechLexique,False);
Ferme (QRechLexique);
end ;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 31/03/2006
Modifié le ... :   /  /
Description .. : Calcul des données S20 et S80 pour contrôles (pas
Suite ........ : d'alimentation de la table DADSDETAIL)
Mots clefs ... : PAYE;PGDADSU;PGCONTROLE
*****************************************************************}
procedure TOF_PG_EDITDADSCONT.CalculEntreprise ();
var
ParamCalc : TParamCalc;
BufDest, BufOrig, StEtab : string;
QRechEtab : TQuery;
Requete, Etablis, Habilitation: string;  //PT6
j: Integer;                              //PT6
begin
BufOrig:= GetParamSoc ('SO_SIRET');
ForceNumerique(BufOrig, BufDest);
ParamCalc.Siren:= Copy(BufDest, 1, 9);

ParamCalc.TypeDecla:= GetControlText ('CNATURE');

ParamCalc.TypeD:= TypeD;
ParamCalc.Libelle:= GetParamSoc ('SO_LIBELLE');
ParamCalc.Du:= GetControlText ('DATEDEB');
ParamCalc.Au:= GetControlText ('DATEFIN');
ParamCalc.Fraction:= GetControlText ('DADSFRACTION');
ParamCalc.FractionTot:= Getparamsoc('SO_PGFRACTION');

//DEB PT6
Habilitation := '';
if Assigned(MonHabilitation) and (MonHabilitation.LeSQL<>'') then
begin // Champs etablissement seul comme critere de la population
  Requete := MonHabilitation.LesEtab;
  Etablis := ReadTokenSt(Requete);
  j := 0;
  while Etablis <> '' do
  begin
    j := j + 1;
    if Etablis <> '' then
    begin
      if j > 1 then Habilitation := Habilitation + ',';
      Habilitation := Habilitation + '"' + Etablis + '"';
    end;
    Etablis := ReadTokenSt(Requete);
  end;
  if j > 0 then
    Habilitation := ' WHERE ET_ETABLISSEMENT IN (' + Habilitation + ')';
end;
//FIN PT6

StEtab:= 'SELECT ET_SIRET, ET_ETABLISSEMENT, ET_LIBELLE, ET_ADRESSE1,'+
         ' ET_ADRESSE2, ET_ADRESSE3, ET_CODEPOSTAL, ET_VILLE, ET_PAYS,'+
         ' ET_FICTIF'+
         ' FROM ETABLISS' + Habilitation;     //PT6
QRechEtab:=OpenSql(StEtab,TRUE);
TEtab:= TOB.Create('Les Etablissements', NIL, -1);
TEtab.LoadDetailDB('ETABLISS','','',QRechEtab,False);
Ferme(QRechEtab);
ParamCalc.TobEtab:= TEtab;

ParamCalc.Siege:= GetParamSoc('SO_ETABLISDEFAUT');
ParamCalc.Depose:= GetParamSoc('SO_PGCDEPOSE');
ParamCalc.Client:= GetParamSoc('SO_PGNUMCLIENT');
ParamCalc.Siret:= GetParamSoc('SO_PGSIRETDESTIN');
{PT7-1
ParamCalc.Communication:= GetParamSoc('SO_PGCOMMCRE');
if (GetParamSoc('SO_PGCOMMCRE')='03') then
   ParamCalc.Coordonnee:= GetParamSoc('SO_PGCOORDCRE')
else
if (GetParamSoc('SO_PGCOMMCRE')='05') then
   begin
   ParamCalc.Civilite:= GetParamSoc('SO_PGCIVILCRE');
   ParamCalc.Personne:= GetParamSoc('SO_PGNOMCRE');
   end;
}
ParamCalc.Coordonnee:= GetParamSoc('SO_PGCOORDCRE');
//FIN PT7-1
ParamCalc.Routage:= GetParamSoc('SO_PGROUTAGE');
ParamCalc.Routage1:= GetParamSoc('SO_PGROUTAGE1');
ParamCalc.Routage2:= GetParamSoc('SO_PGROUTAGE2');
ParamCalc.Caracteristique:= GetControlText ('CCAR');;

CalculS20_Entreprise(ParamCalc);
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 15/06/2006
Modifié le ... :   /  /
Description .. : Confection de la requête en fonction des paramètres indiqués
Suite ........ : sur la QRS1
Mots clefs ... : PAYE;PGDADSU;PGCONTROLE
*****************************************************************}
procedure TOF_PG_EDITDADSCONT.RemplitTOBErreur;
var
Etab, EtabSelect, StBloquant, StCompl, StDate, StEtabEtab, StEtabGen : String;
StEtabSal, StHonor, StSelect, StSupp, StTypeDecla, TD2 : String;
QRechEtab, QRechHonor, QRechSelect : TQuery;
TErreurD, TEtabDADS, TEtabDADSD, TLexiqueD : TOB;
Habilitation:String;//PT6
LesEtablis,Etabl : String; //PT6
j : integer;              //PT6
begin
TD2:= Copy (GetControlText ('CNATURE'),1,2);
//PT8
if (TD2>='03') then
   StDate:= ' PDS_DATEDEBUT>="'+UsDateTime (DebExer)+'" AND'+
            ' PDS_DATEFIN<="'+UsDateTime (FinExer)+'" AND'
else
   StDate:= '';
//FIN PT8
StTypeDecla:= GetDADSUNature (TD2);

if (GetCheckBoxState ('CBBLOQUANT')=CbChecked) then
   StBloquant:= ' PSU_CTRLBLOQUANT="X" AND'
else
   StBloquant:= '';

//DONNEES ENTREPRISE
StSelect:= 'SELECT DADSCONTROLE.*, DADSLEXIQUE.PDL_DADSNOM'+
           ' FROM DADSCONTROLE'+
           ' LEFT JOIN DADSLEXIQUE ON'+
           ' PSU_SEGMENT=PDL_DADSSEGMENT WHERE'+
           ' PSU_SALARIE LIKE "++%" AND'+
           ' PSU_TYPE="'+TypeD+'" AND'+
           ' PSU_DATEDEBUT>="'+UsDateTime(DebExer)+'" AND'+
           ' PSU_DATEFIN<="'+UsDateTime(FinExer)+'" AND'+
           ' PSU_EXERCICEDADS="'+PGExercice+'" AND'+StBloquant+
           ' PSU_ORIGINECRIT = "S20" AND'+
           ' PDL_EXERCICEDEB <= "'+PGExercice+'" AND'+
           ' (PDL_EXERCICEFIN >= "'+PGExercice+'" OR'+
           ' PDL_EXERCICEFIN = "")'+StTypeDecla+
           ' ORDER BY PSU_SALARIE, PSU_ORDRE, PSU_SEGMENT';
QRechSelect:= OpenSql(StSelect,TRUE);
TErreur.LoadDetailDB('DADSCONTROLE','','',QRechSelect,False);
Ferme(QRechSelect);

if (THMultiValCombobox(GetControl('MCETAB')).Tous) then
{PT9
   StEtabGen:= ''
}
   begin
   StEtabEtab:= '';
   TEtabDADS:= TOB.Create ('Les etablissements', nil, -1);
   ChargeEtabNonFictif (TEtabDADS);
   if (TEtabDADS<>nil) then
      begin
      TEtabDADSD:= TEtabDADS.FindFirst ([''], [''], False);
      if (TEtabDADSD<>nil) then
         begin
         Etab:= TEtabDADSD.GetValue ('ET_ETABLISSEMENT');
         while (Etab<>'') do
               begin
               StEtabSal:= StEtabSal+' PDS_DONNEEAFFICH="'+Etab+'"';
               StEtabEtab:= StEtabEtab+' PSU_SALARIE="**'+Etab+'"';
               TEtabDADSD:= TEtabDADS.FindNext ([''], [''], False);
               if (TEtabDADSD<>nil) then
                  Etab:= TEtabDADSD.GetValue ('ET_ETABLISSEMENT')
               else
                  Etab:='';
               if (Etab<>'') then
                  begin
                  StEtabSal:= StEtabSal+' OR';
                  StEtabEtab:= StEtabEtab+' OR';
                  end;
               end;
         StEtabGen:= ' AND'+
                     ' PSU_SALARIE IN (SELECT PDS_SALARIE FROM DADSDETAIL WHERE'+
                     ' PDS_TYPE = "'+TypeD+'" AND'+
                     ' PDS_ORDRE > 0 AND'+StDate+
                     ' PDS_EXERCICEDADS = "'+PGExercice+'" AND'+
                     ' PDS_SEGMENT="S41.G01.00.005" AND ('+StEtabSal+'))';
         StEtabEtab:= ' AND'+
                      ' ('+StEtabEtab+')';
         end
      else
         StEtabGen:= '';
      end
   else
      StEtabGen:= '';
   end
//FIN PT9
else
   begin
   EtabSelect:= GetControlText('MCETAB');
   Etab:= ReadTokenpipe(EtabSelect, ';');
   While (Etab <> '') do
         begin
         StEtabSal:= StEtabSal+' PDS_DONNEEAFFICH="'+Etab+'"';
         StEtabEtab:= StEtabEtab+' PSU_SALARIE="**'+Etab+'"';
         Etab:= ReadTokenpipe(EtabSelect,';');
         if (Etab <> '') then
            begin
            StEtabSal:= StEtabSal+' OR';
            StEtabEtab:= StEtabEtab+' OR';
            end;
         end;
   StEtabGen:= ' AND'+
               ' PSU_SALARIE IN (SELECT PDS_SALARIE FROM DADSDETAIL WHERE'+
               ' PDS_TYPE = "'+TypeD+'" AND'+
               ' PDS_ORDRE > 0 AND'+StDate+
               ' PDS_EXERCICEDADS = "'+PGExercice+'" AND'+
               ' PDS_SEGMENT="S41.G01.00.005" AND ('+StEtabSal+'))';
   StEtabEtab:= ' AND'+
                ' ('+StEtabEtab+')';
   end;

StCompl:= '';
if (GetControlText ('PSA_SALARIE')<> '') then
   StCompl:= ' AND PSA_SALARIE = "'+GetControlText ('PSA_SALARIE')+'"';
if (GetControlText ('PSA_LIBELLE')<> '') then
   StCompl:= StCompl+' AND PSA_LIBELLE = "'+GetControlText ('PSA_LIBELLE')+'"';
if (GetControlText ('PSA_TRAVAILN1')<> '') then
   StCompl:= StCompl+' AND PSA_TRAVAILN1 = "'+GetControlText ('PSA_TRAVAILN1')+'"';
if (GetControlText ('PSA_TRAVAILN2')<> '') then
   StCompl:= StCompl+' AND PSA_TRAVAILN2 = "'+GetControlText ('PSA_TRAVAILN2')+'"';
if (GetControlText ('PSA_TRAVAILN3')<> '') then
   StCompl:= StCompl+' AND PSA_TRAVAILN3 = "'+GetControlText ('PSA_TRAVAILN3')+'"';
if (GetControlText ('PSA_TRAVAILN4')<> '') then
   StCompl:= StCompl+' AND PSA_TRAVAILN4 = "'+GetControlText ('PSA_TRAVAILN4')+'"';
if (GetControlText ('PSA_CODESTAT')<> '') then
   StCompl:= StCompl+' AND PSA_CODESTAT = "'+GetControlText ('PSA_CODESTAT')+'"';
if (GetControlText ('PSA_DADSDATE')<> '  /  /    ') then
   StCompl:= StCompl+' AND PSA_DADSDATE = "'+GetControlText ('PSA_DADSDATE')+'"';

//DEB PT6
if Assigned(MonHabilitation) and (MonHabilitation.LeSQL <> '') then
  StCompl := StCompl + ' AND ' + MonHabilitation.LeSQL;
//FIN PT6

StEtabGen:= StEtabGen+' AND'+
            ' PSU_SALARIE IN (SELECT PSA_SALARIE FROM SALARIES WHERE'+
            ' PSA_DADSFRACTION="'+GetControlText ('DADSFRACTION')+'"'+StCompl+')';

//DONNEES SALARIE
if (GetControlText('CHNEANT')<>'X') then
   begin
   if (GetControlText('CHPREVOYANCE')='-') then
      StSupp:= ' AND PSU_SEGMENT NOT LIKE "S45%"'
   else
      StSupp:= '';
   StSelect:= 'SELECT DADSCONTROLE.*, DADSLEXIQUE.PDL_DADSNOM'+
              ' FROM DADSCONTROLE'+
              ' LEFT JOIN DADSLEXIQUE ON'+
              ' PSU_SEGMENT=PDL_DADSSEGMENT WHERE'+
              ' PSU_TYPE="'+TypeD+'" AND'+
              ' PSU_DATEDEBUT>="'+UsDateTime(DebExer)+'" AND'+
              ' PSU_DATEFIN<="'+UsDateTime(FinExer)+'" AND'+
              ' PSU_EXERCICEDADS="'+PGExercice+'" AND'+StBloquant+
              ' PSU_ORIGINECRIT LIKE "S%" AND'+
              ' PDL_EXERCICEDEB <= "'+PGExercice+'" AND'+
              ' (PDL_EXERCICEFIN >= "'+PGExercice+'" OR'+
              ' PDL_EXERCICEFIN = "")'+StTypeDecla+StEtabGen+StSupp+
              ' ORDER BY PSU_SALARIE, PSU_ORDRE, PSU_SEGMENT';
   QRechSelect:= OpenSql(StSelect,TRUE);
   TErreur.LoadDetailDB('DADSCONTROLE','','',QRechSelect,True);
   Ferme(QRechSelect);

//DONNEES HONORAIRES
   if (GetControlText('CHHONORAIRE')='X') then
      begin
{PT9
      if (THMultiValCombobox(GetControl('MCETAB')).Tous) then
}
      if ((THMultiValCombobox(GetControl('MCETAB')).Tous) and
         (TEtabDADS=nil)) then
//FIN PT9
         StEtabGen:= ''
      else
         StEtabGen:= ' AND'+
                     ' PSU_ORDRE IN (SELECT PDS_ORDRE FROM DADSDETAIL WHERE'+
                     ' PDS_SEGMENT="S70.G01.00.014" AND'+
                     ' PDS_DATEDEBUT >= "'+UsDateTime(DebExer)+'" AND'+
                     ' PDS_DATEFIN <= "'+UsDateTime(FinExer)+'" AND'+
                     ' PDS_EXERCICEDADS = "'+PGExercice+'" AND ('+
                     StEtabSal+'))';
      StHonor:= 'SELECT DADSCONTROLE.*, DADSLEXIQUE.PDL_DADSNOM'+
                ' FROM DADSCONTROLE'+
                ' LEFT JOIN DADSLEXIQUE ON'+
                ' PSU_SEGMENT=PDL_DADSSEGMENT WHERE'+
                ' PSU_SALARIE LIKE "--H%" AND'+
                ' PSU_TYPE="'+TypeD+'" AND'+
                ' PSU_DATEDEBUT>="'+UsDateTime(DebExer)+'" AND'+
                ' PSU_DATEFIN<="'+UsDateTime(FinExer)+'" AND'+StBloquant+
                ' PSU_ORDRE IN (SELECT PDS_ORDRE FROM DADSDETAIL WHERE'+
                ' PDS_SEGMENT="S70.G01.01.002" AND'+
                ' PDS_SALARIE LIKE "--H%" AND PDS_TYPE="'+TypeD+'" AND'+
                ' PDS_DATEDEBUT>="'+UsDateTime(DebExer)+'" AND'+
             //   ' PDS_DATEFIN<="'+UsDateTime(FinExer)+'") AND'+ // PT10-3
             //   ' PDS_EXERCICEDADS = "'+PGExercice+'" AND'+     // PT10-3
                ' PDS_DATEFIN<="'+UsDateTime(FinExer)+'" AND'+    // PT10-3
                ' PDS_EXERCICEDADS = "'+PGExercice+'") AND'+      // PT10-3
                ' PDL_EXERCICEDEB <= "'+PGExercice+'" AND'+
                ' (PDL_EXERCICEFIN >= "'+PGExercice+'" OR'+
                ' PDL_EXERCICEFIN = "")'+StTypeDecla+StEtabGen+
                ' ORDER BY PSU_SALARIE, PSU_ORDRE, PSU_SEGMENT';
      QRechHonor:= OpenSql(StHonor,TRUE);
      TErreur.LoadDetailDB ('DADSCONTROLE','','',QRechHonor,True);
      Ferme (QRechHonor);
      end;
   end;

//DONNEES ETABLISSEMENT
//DEB PT6
Habilitation := '';
if Assigned(MonHabilitation) and (MonHabilitation.LeSQL <> '') then
begin
  LesEtablis := MonHabilitation.LesEtab;
  Etabl := ReadTokenSt(LesEtablis);
  Habilitation := '';
  j := 0;
  while Etabl <> '' do
  begin
    j := j + 1;
    if Etabl <> '' then
    begin
      if j > 1 then Habilitation := Habilitation + ',';
      Habilitation := Habilitation + '"' + Etabl + '"';
    end;
    Etabl := ReadTokenSt(LesEtablis);
  end;
  if j > 0 then
    // La zone PSU_SALARIE contient plusieurs types d'info
    // Dans le cas d'un établissement il se compose de '**' + établissement
    Habilitation := ' AND RIGHT(PSU_SALARIE,3) IN (' + Habilitation + MonHabilitation.SqlPop + ')';
end;
//FIN PT6
StSelect:= 'SELECT DADSCONTROLE.*, DADSLEXIQUE.PDL_DADSNOM'+
           ' FROM DADSCONTROLE'+
           ' LEFT JOIN DADSLEXIQUE ON'+
           ' PSU_SEGMENT=PDL_DADSSEGMENT WHERE'+
           ' PSU_TYPE="'+TypeD+'" AND'+
           ' PSU_DATEDEBUT>="'+UsDateTime(DebExer)+'" AND'+
           ' PSU_DATEFIN<="'+UsDateTime(FinExer)+'" AND'+
           ' PSU_EXERCICEDADS="'+PGExercice+'" AND'+StBloquant+
           ' PSU_ORIGINECRIT = "S80" AND'+
           ' PDL_EXERCICEDEB <= "'+PGExercice+'" AND'+
           ' (PDL_EXERCICEFIN >= "'+PGExercice+'" OR'+
           ' PDL_EXERCICEFIN = "")'+StTypeDecla+StEtabEtab+ Habilitation + {PT6}
           ' ORDER BY PSU_SALARIE, PSU_ORDRE, PSU_SEGMENT';
QRechSelect:= OpenSql(StSelect,TRUE);
TErreur.LoadDetailDB ('DADSCONTROLE','','',QRechSelect,True);
Ferme(QRechSelect);

//Boucle pour ajouter le champ PDL_DADSNOM
TErreurD:= TErreur.FindFirst ([''], [''], False);
While TErreurD <> NIL do
    begin
    TLexiqueD:= TLexique.FindFirst (['PDL_DADSSEGMENT'],
                                    [TErreurD.GetValue ('PSU_SEGMENT')],FALSE);
    TErreurD.AddChampSupValeur ('PDL_DADSNOM',
                                TLexiqueD.GetValue ('PDL_DADSNOM'), False);
    TErreurD:= TErreur.FindNext ([''], [''], False);
    end;
FreeAndNil (TEtabDADS);         //PT9
//TobDebug (TErreur);
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 31/03/2006
Modifié le ... :   /  /
Description .. : Procédure d'alimentation de la table DADSCONTROLE
Mots clefs ... : PAYE;PGDADSU;PGCONTROLE
*****************************************************************}
procedure TOF_PG_EDITDADSCONT.ControleDossier ();
var
TDADSUDetail, TErreurD, TEtabDADS, TEtabDADSD, TInact, TInactD, TPrudh : TOB;
TPrudhD, TSal, TSalD : TOB;
Buffer, CodeSal, Etab, EtabSelect, Matricule, MatriculePrec, Ordre : string;
StCompl, StDate, StEtab, StEtabEtab, StEtabGen, StHonor, StInact : string;
StNbreH, StPrecSeg, StPrudh, StSelect, StSupp, StTypeDecla, TD2 : string;
QRechEtab, QRechHonor, QRechInact, QRechPrudh, QRechSelect : TQuery;
Ircantec, Next, S41G02, S44, S90000 : boolean;
i, j, Nbre30, NbreSeg, NbreSeg2, NbreSeg3, Reponse : integer;
ErreurDADSU : TControle;
begin
Nbre30:= 0;

TD2:= Copy (GetControlText ('CNATURE'),1,2);
//PT8
if (TD2>='03') then
   StDate:= ' PDS_DATEDEBUT>="'+UsDateTime (DebExer)+'" AND'+
            ' PDS_DATEFIN<="'+UsDateTime (FinExer)+'" AND'
else
   StDate:= '';
//FIN PT8
StTypeDecla:= GetDADSUNature (TD2);

ErreurDADSU.TypeD:= TypeD;
ErreurDADSU.Exercice:= PGExercice;

TErreurD := TErreur.FindFirst([''],[''],FALSE);
if ((TErreurD <> nil)) then
   begin
   Reponse:= PGIAsk ('Des anomalies ont déjà été constatées. Voulez-vous#13#10'+
                     'réinitialiser l''état de contrôle ?',
                     Ecran.Caption);
   if (Reponse <> MrYes) then
      exit;
   end;

TErreur.DeleteDB;

//TobDebug (TDADSUD);
ControleTOB (TDADSUD, False, True);
TDADSUD.ClearDetail;

//Données SALARIE
if ((GetControlText ('CHNEANT')<>'X') and (TD2<>'12')) then
   begin
   if (GetControlText('CHPREVOYANCE')='-') then
      StSupp:= ' AND PDS_SEGMENT NOT LIKE "S45%"'
   else
      StSupp:= '';

   if (TD2='08') then
      StSupp:= StSupp+' AND PDS_SALARIE IN (SELECT PDS_SALARIE'+
               ' FROM DADSDETAIL WHERE'+
               ' PDS_TYPE = "'+TypeD+'" AND'+
               ' PDS_ORDRE > 0 AND'+StDate+
               ' PDS_EXERCICEDADS = "'+PGExercice+'" AND'+
               ' PDS_SEGMENT="S45.G01.00.001")';

   StEtab:= '';
   if (THMultiValCombobox(GetControl('MCETAB')).Tous) then
{PT9
      StEtabGen:= ''
}
      begin
      StEtabEtab:= '';
      TEtabDADS:= TOB.Create ('Les etablissements', nil, -1);
      ChargeEtabNonFictif (TEtabDADS);
      if (TEtabDADS<>nil) then
         begin
         TEtabDADSD:= TEtabDADS.FindFirst ([''], [''], False);
         if (TEtabDADSD<>nil) then
            begin
            Etab:= TEtabDADSD.GetValue ('ET_ETABLISSEMENT');
            while (Etab<>'') do
                  begin
                  StEtab:= StEtab+' PDS_DONNEEAFFICH="'+Etab+'"';
                  TEtabDADSD:= TEtabDADS.FindNext ([''], [''], False);
                  if (TEtabDADSD<>nil) then
                     Etab:= TEtabDADSD.GetValue ('ET_ETABLISSEMENT')
                  else
                     Etab:='';
                  if (Etab<>'') then
                     StEtab:= StEtab+' OR';
                  end;
            StEtabGen:= ' AND'+
                        ' PDS_SALARIE IN (SELECT PDS_SALARIE'+
                        ' FROM DADSDETAIL WHERE'+
                        ' PDS_TYPE = "'+TypeD+'" AND'+
                        ' PDS_ORDRE > 0 AND'+StDate+
                        ' PDS_EXERCICEDADS = "'+PGExercice+'" AND'+
                        ' PDS_SEGMENT="S41.G01.00.005" AND ('+StEtab+'))';
            end
         else
            StEtabGen:= '';
         end
      else
         StEtabGen:= '';
      end
//FIN PT9
   else
      begin
      EtabSelect:= GetControlText('MCETAB');
      Etab:= ReadTokenpipe(EtabSelect, ';');
      While (Etab <> '') do
            begin
            StEtab:= StEtab+' PDS_DONNEEAFFICH="'+Etab+'"';
            Etab := ReadTokenpipe(EtabSelect,';');
            if (Etab <> '') then
               StEtab:= StEtab+' OR';
            end;
      StEtabGen:= ' AND'+
                  ' PDS_SALARIE IN (SELECT PDS_SALARIE FROM DADSDETAIL'+
                  ' WHERE PDS_TYPE = "'+TypeD+'" AND'+
                  ' PDS_ORDRE > 0 AND'+StDate+
                  ' PDS_EXERCICEDADS = "'+PGExercice+'" AND'+
                  ' PDS_SEGMENT="S41.G01.00.005" AND ('+StEtab+'))';
      end;

   StCompl:= '';
   if (GetControlText ('PSA_SALARIE')<> '') then
      StCompl:= ' AND PSA_SALARIE = "'+GetControlText ('PSA_SALARIE')+'"';
   if (GetControlText ('PSA_LIBELLE')<> '') then
      StCompl:= StCompl+' AND PSA_LIBELLE = "'+GetControlText ('PSA_LIBELLE')+'"';
   if (GetControlText ('PSA_TRAVAILN1')<> '') then
      StCompl:= StCompl+' AND PSA_TRAVAILN1 = "'+GetControlText ('PSA_TRAVAILN1')+'"';
   if (GetControlText ('PSA_TRAVAILN2')<> '') then
      StCompl:= StCompl+' AND PSA_TRAVAILN2 = "'+GetControlText ('PSA_TRAVAILN2')+'"';
   if (GetControlText ('PSA_TRAVAILN3')<> '') then
      StCompl:= StCompl+' AND PSA_TRAVAILN3 = "'+GetControlText ('PSA_TRAVAILN3')+'"';
   if (GetControlText ('PSA_TRAVAILN4')<> '') then
      StCompl:= StCompl+' AND PSA_TRAVAILN4 = "'+GetControlText ('PSA_TRAVAILN4')+'"';
   if (GetControlText ('PSA_CODESTAT')<> '') then
      StCompl:= StCompl+' AND PSA_CODESTAT = "'+GetControlText ('PSA_CODESTAT')+'"';
   if (GetControlText ('PSA_DADSDATE')<> '  /  /    ') then
      StCompl:= StCompl+' AND PSA_DADSDATE = "'+GetControlText ('PSA_DADSDATE')+'"';

    //DEB PT6
    if Assigned(MonHabilitation) and (MonHabilitation.LeSQL <> '') then
      StCompl := StCompl + ' AND ' + MonHabilitation.LeSQL;
    //FIN PT6

   StSelect:= 'SELECT PDS_SALARIE, PDS_SEGMENT, PDS_DATEDEBUT, PDS_DATEFIN,'+
              ' PDS_EXERCICEDADS, PDS_DONNEE, PDS_ORDRE, PDS_ORDRESEG'+
              ' FROM DADSDETAIL'+
              ' LEFT JOIN SALARIES ON'+
              ' PDS_SALARIE = PSA_SALARIE'+
              ' LEFT JOIN DADSLEXIQUE ON'+
              ' PDS_SEGMENT=PDL_DADSSEGMENT WHERE'+
              ' PDS_TYPE = "'+TypeD+'" AND'+
              ' PDS_ORDRE >= 0 AND'+StDate+
              ' PDS_EXERCICEDADS = "'+PGExercice+'"'+StSupp+' AND'+
              ' PDS_SEGMENT NOT LIKE "S41.G02%" AND'+
              ' PSA_DADSFRACTION="'+GetControlText ('DADSFRACTION')+'" AND'+
              ' PDL_EXERCICEDEB <= "'+PGExercice+'" AND'+
              ' (PDL_EXERCICEFIN >= "'+PGExercice+'" OR'+
              ' PDL_EXERCICEFIN = "")'+StTypeDecla+StEtabGen+StCompl+
              ' ORDER BY PDS_SALARIE,PDS_DATEDEBUT,PDS_ORDRE,'+
              ' PDS_ORDRESEG,PDS_SEGMENT,PDS_DATEFIN';
   QRechSelect:= OpenSql(StSelect,TRUE);
   TSal:= TOB.Create('Les salariés', NIL, -1);
   TSal.LoadDetailDB('DADSDETAIL','','',QRechSelect,False);
   Ferme(QRechSelect);

   Next:= False;
   S44:= False;
   S90000:= False;
   S41G02:= True;
   FreeAndNil (TPrudh);
   NbreSeg3:= 0;
   NbreSeg:= TSal.FillesCount(1);

   TSalD:= TSal.FindFirst([''],[''],FALSE);
   if (TSalD <> nil) then
      begin
      CodeSal:= TSalD.GetValue('PDS_SALARIE');
      For i := 1 to NbreSeg do
          begin
          if ((TSalD<>nil) and
             ((TSalD.GetValue('PDS_SEGMENT')='S30.G01.00.001') or
             (TSalD.GetValue('PDS_SEGMENT')>'S41.G02')) and
             (GetControlText('CBPRUDH')='X') and (S41G02=False) and
             (TPrudh<>nil)) then
             begin
             NbreSeg3:= TPrudh.FillesCount(1);
             TPrudhD:= TPrudh.FindFirst([''],[''],FALSE);
             For j := 1 to NbreSeg3 do
                 begin
                 TDADSUDetail:= Tob.Create ('DADSDETAIL', TDADSUD, -1);
                 TDADSUDetail.PutValue ('PDS_SALARIE',
                                        TPrudhD.GetValue('PDS_SALARIE'));
                 TDADSUDetail.PutValue ('PDS_TYPE', TypeD);
                 TDADSUDetail.PutValue ('PDS_ORDRE',
                                        TPrudhD.GetValue('PDS_ORDRE'));
                 TDADSUDetail.PutValue ('PDS_DATEDEBUT', FinExer);
                 TDADSUDetail.PutValue ('PDS_DATEFIN', FinExer);
                 TDADSUDetail.PutValue ('PDS_EXERCICEDADS', PGExercice);
                 TDADSUDetail.PutValue ('PDS_ORDRESEG',
                                        TPrudhD.GetValue('PDS_ORDRESEG'));
                 TDADSUDetail.PutValue ('PDS_SEGMENT',
                                        TPrudhD.GetValue('PDS_SEGMENT'));
                 TDADSUDetail.PutValue ('PDS_DONNEE',
                                        TPrudhD.GetValue('PDS_DONNEE'));
                 MoveCurProgressForm ('Contrôle salarié '+CodeSal);
                 MoveCur(False);
                 TPrudhD:= TPrudh.FindNext([''],[''],FALSE);
                 end;
             S41G02:= True;
             end;

          if ((TSalD <> nil) and
             ((TSalD.GetValue('PDS_SEGMENT') = 'S30.G01.00.001') or
             (TSalD.GetValue('PDS_SEGMENT') = 'S41.G01.00.001')) and
             ((TD2 = '01') or (TD2 = '03') or (TD2 = '07'))) then
             begin
             if ((StPrecSeg > 'S30') and (S44 = False)) then
                begin
                if ((S90000 = False) and (Ircantec=False)) then
                   begin
                   TDADSUDetail:= Tob.Create ('DADSDETAIL', TDADSUD, -1);
                   TDADSUDetail.PutValue ('PDS_SALARIE', CodeSal);
                   TDADSUDetail.PutValue ('PDS_TYPE', TypeD);
                   TDADSUDetail.PutValue ('PDS_ORDRE',
                                          TSalD.GetValue('PDS_ORDRE'));
                   TDADSUDetail.PutValue ('PDS_DATEDEBUT', DebExer);
                   TDADSUDetail.PutValue ('PDS_DATEFIN', FinExer);
                   TDADSUDetail.PutValue ('PDS_EXERCICEDADS', PGExercice);
                   TDADSUDetail.PutValue ('PDS_ORDRESEG', 751);
                   TDADSUDetail.PutValue ('PDS_SEGMENT', 'S44.G01.00.001');
                   TDADSUDetail.PutValue ('PDS_DONNEE', '01');
                   MoveCurProgressForm ('Contrôle salarié '+CodeSal);
                   MoveCur(False);

                   TDADSUDetail:= Tob.Create ('DADSDETAIL', TDADSUD, -1);
                   TDADSUDetail.PutValue ('PDS_SALARIE', CodeSal);
                   TDADSUDetail.PutValue ('PDS_TYPE', TypeD);
                   TDADSUDetail.PutValue ('PDS_ORDRE',
                                          TSalD.GetValue('PDS_ORDRE'));
                   TDADSUDetail.PutValue ('PDS_DATEDEBUT', DebExer);
                   TDADSUDetail.PutValue ('PDS_DATEFIN', FinExer);
                   TDADSUDetail.PutValue ('PDS_EXERCICEDADS', PGExercice);
                   TDADSUDetail.PutValue ('PDS_ORDRESEG', 702);
                   TDADSUDetail.PutValue ('PDS_SEGMENT', 'S44.G01.00.002');
                   TDADSUDetail.PutValue ('PDS_DONNEE',
                                          IntToStr(StrToInt(StNbreH)*100));
                   MoveCurProgressForm ('Contrôle salarié '+CodeSal);
                   MoveCur(False);
                   end;
                end;
             S44:= False;
             end;

          if ((TSalD <> nil) and
             (TSalD.GetValue('PDS_SEGMENT') = 'S41.G01.00.001')) then
             begin
             if (((TD2 = '01') or (TD2 = '03') or (TD2 = '07'))) then
                begin
                S90000:= False;
                Ircantec:= False;
                end;
             ErreurDADSU.Salarie:= TSalD.GetValue('PDS_SALARIE');
             ErreurDADSU.Num:= TSalD.GetValue('PDS_ORDRE');
             ErreurDADSU.DateDeb:= TSalD.GetValue('PDS_DATEDEBUT');
             ErreurDADSU.DateFin:= TSalD.GetValue('PDS_DATEFIN');
             ErreurDADSU.Segment:= TSalD.GetValue('PDS_SEGMENT');
             end;

          if ((TSalD <> nil) and
             (TSalD.GetValue('PDS_SEGMENT')='S30.G01.00.001')) then
             begin
             ErreurDADSU.Salarie:= TSalD.GetValue('PDS_SALARIE');
             Nbre30:= Nbre30+1;
             if (Nbre30 <> 1) then
                begin
                StInact:= 'SELECT PDS_SALARIE, PDS_SEGMENT, PDS_DATEDEBUT,'+
                          ' PDS_DONNEE, PDS_ORDRE, PDS_ORDRESEG'+
                          ' FROM DADSDETAIL'+
                          ' LEFT JOIN SALARIES ON'+
                          ' PDS_SALARIE = PSA_SALARIE'+
                          ' LEFT JOIN DADSLEXIQUE ON'+
                          ' PDS_SEGMENT=PDL_DADSSEGMENT WHERE'+
                          ' PDS_SALARIE = "'+Matricule+'" AND'+
                          ' PDS_TYPE = "'+TypeD+'" AND'+
                          ' PDS_ORDRE < 0 AND'+
                          ' PDS_EXERCICEDADS = "'+PGExercice+'" AND'+
                          ' PSA_DADSFRACTION="'+GetControlText ('DADSFRACTION')+'" AND'+
                          ' PDL_EXERCICEDEB <= "'+PGExercice+'" AND'+
                          ' (PDL_EXERCICEFIN >= "'+PGExercice+'" OR'+
                          ' PDL_EXERCICEFIN = "")'+StTypeDecla+
                          ' ORDER BY PDS_DATEDEBUT,PDS_ORDRE DESC,'+
                          ' PDS_ORDRESEG,PDS_SEGMENT,PDS_DATEFIN';
                QRechInact:=OpenSql(StInact,TRUE);
                TInact:= TOB.Create('Les éléments Inactivité', NIL, -1);
                TInact.LoadDetailDB('DADSDETAIL','','',QRechInact,False);
                Ferme(QRechInact);

                NbreSeg2:= TInact.FillesCount(1);
                TInactD:= TInact.FindFirst([''],[''],FALSE);
                For j := 1 to NbreSeg2 do
                    begin
                    TDADSUDetail:= Tob.Create ('DADSDETAIL', TDADSUD, -1);
                    TDADSUDetail.PutValue ('PDS_SALARIE',
                                           TInactD.GetValue('PDS_SALARIE'));
                    TDADSUDetail.PutValue ('PDS_TYPE', TypeD);
                    TDADSUDetail.PutValue ('PDS_ORDRE',
                                           TInactD.GetValue('PDS_ORDRE'));
                    TDADSUDetail.PutValue ('PDS_DATEDEBUT', DebExer);
                    TDADSUDetail.PutValue ('PDS_DATEFIN', FinExer);
                    TDADSUDetail.PutValue ('PDS_EXERCICEDADS', PGExercice);
                    TDADSUDetail.PutValue ('PDS_ORDRESEG',
                                           TInactD.GetValue('PDS_ORDRESEG'));
                    TDADSUDetail.PutValue ('PDS_SEGMENT',
                                           TInactD.GetValue('PDS_SEGMENT'));
                    TDADSUDetail.PutValue ('PDS_DONNEE',
                                           TInactD.GetValue('PDS_DONNEE'));
                    MoveCurProgressForm ('Contrôle salarié '+CodeSal);
                    MoveCur(False);

                    TInactD:= TInact.FindNext([''],[''],FALSE);
                    end;
                FreeAndNil(TInact);
                end;
             end;

          if ((TSalD <> nil) and
             (TSalD.GetValue('PDS_SEGMENT') = 'S41.G01.00.019')) then
             begin
             MatriculePrec:= Matricule;
             Matricule:= TSalD.GetValue('PDS_DONNEE');
             if (GetControlText('CBPRUDH')='X') then
                begin
                S41G02:= False;
                if (MatriculePrec<>Matricule) then
                   begin
{PT9
                   if (THMultiValCombobox(GetControl('MCETAB')).Tous) then
}
                   if ((THMultiValCombobox(GetControl('MCETAB')).Tous) and
                      (TEtabDADS=nil)) then
//FIN PT9
                      StEtabGen:= ''
                   else
                      StEtabGen:= ' AND'+
                                  ' PDS_ORDRE IN (SELECT PDS_ORDRE FROM'+
                                  ' DADSDETAIL WHERE'+
                                  ' PDS_TYPE = "'+TypeD+'" AND'+
                                  ' PDS_ORDRE > 0 AND'+StDate+
                                  ' PDS_SEGMENT NOT LIKE "S41.G02%" AND'+
                                  ' PDS_SEGMENT="S41.G01.00.005" AND ('+StEtab+'))';
                   StPrudh:= 'SELECT *'+
                             ' FROM DADSDETAIL WHERE'+
                             ' PDS_SALARIE = "'+Matricule+'" AND'+
                             ' PDS_TYPE = "'+TypeD+'" AND'+
                             ' PDS_ORDRE >= 0 AND'+StDate+
                             ' PDS_EXERCICEDADS = "'+PGExercice+'" AND'+
                             ' PDS_SEGMENT IN (SELECT PDL_DADSSEGMENT FROM'+
                             ' DADSLEXIQUE WHERE PDL_DADSSEGMENT LIKE "S41.G02%" AND'+
                             ' PDL_EXERCICEDEB <= "'+PGExercice+'" AND'+
                             ' (PDL_EXERCICEFIN >= "'+PGExercice+'" OR'+
                             ' PDL_EXERCICEFIN = "")'+StTypeDecla+')'+StEtabGen+
                             ' ORDER BY PDS_DATEDEBUT,PDS_ORDRE, PDS_ORDRESEG,'+
                             ' PDS_SEGMENT,PDS_DATEFIN';

                   QRechPrudh:= OpenSql(StPrudh,TRUE);
                   FreeAndNil (TPrudh);
                   TPrudh:= TOB.Create('Les éléments prud''hommaux', NIL, -1);
                   TPrudh.LoadDetailDB('DADSDETAIL','','',QRechPrudh,False);
                   Ferme(QRechPrudh);
                   end;
                end;
             end;

          if ((TSalD <> nil) and
             (TSalD.GetValue('PDS_SEGMENT') = 'S41.G01.01.001')) then
             begin
             if (TD2='02') then
                begin
                if ((TSalD.GetValue ('PDS_DONNEE')<>'I0001') and
                   (TSalD.GetValue ('PDS_DONNEE')<>'I0002') and
                   (TSalD.GetValue ('PDS_DONNEE')<>'F0002') and
                   (TSalD.GetValue ('PDS_DONNEE')<>'R0001') and
                   (TSalD.GetValue ('PDS_DONNEE')<>'CNBF')) then
                begin
                Buffer:= 'S41.G01.01';
                Ordre:= TSalD.GetValue('PDS_ORDRE');             //PT4
                While Buffer = 'S41.G01.01' do
                      begin
                      TSalD:= TSalD.FindNext([''],[''],FALSE);
                      if TSalD <> nil then
                         Buffer:= Copy(TSalD.GetValue('PDS_SEGMENT'),1,10)
                      else
                         Buffer:= '';
                      Next:= True;
                      end;

                TDADSUDetail:= Tob.Create ('DADSDETAIL', TDADSUD, -1);
                TDADSUDetail.PutValue ('PDS_SALARIE', CodeSal);
                TDADSUDetail.PutValue ('PDS_TYPE', TypeD);
{PT4
                TDADSUDetail.PutValue ('PDS_ORDRE',
                                       TSalD.GetValue('PDS_ORDRE'));
}
                TDADSUDetail.PutValue ('PDS_ORDRE', Ordre);
//FIN PT4                
                TDADSUDetail.PutValue ('PDS_DATEDEBUT', DebExer);
                TDADSUDetail.PutValue ('PDS_DATEFIN', FinExer);
                TDADSUDetail.PutValue ('PDS_EXERCICEDADS', PGExercice);
                TDADSUDetail.PutValue ('PDS_ORDRESEG', 111);
                TDADSUDetail.PutValue ('PDS_SEGMENT', 'S41.G01.01.001');
                TDADSUDetail.PutValue ('PDS_DONNEE', '90000');
                MoveCur(False);
                end;
                S90000:= True;
                end
             else
                begin
                if (TSalD.GetValue('PDS_DONNEE') = '90000') then
                   S90000:= True;
                end;
             if ((TSalD <> nil) and
                (Pos('I', TSalD.GetValue('PDS_DONNEE'))=1)) then
                Ircantec:= True;
             end;

          if ((TSalD <> nil) and
             (TSalD.GetValue('PDS_SEGMENT') = 'S44.G01.00.001')) then
             begin
             if ((S90000=True) or (Ircantec=True)) then
                begin
                Buffer:= 'S44';
                While Buffer = 'S44' do
                      begin
                      TSalD:= TSal.FindNext([''],[''],FALSE);
                      if TSalD <> nil then
                         begin
                         Buffer:= Copy(TSalD.GetValue('PDS_SEGMENT'),1,3);
                         CodeSal:= TSalD.GetValue('PDS_SALARIE');
                         Ordre:= TSalD.GetValue('PDS_ORDRE');    //PT4
                         end
                      else
                         Buffer:= '';
                      Next:= True;
                      end;
                end;

             if ((TD2='01') or (TD2='03') or (TD2='07')) then
                S44:= True;
             end;

          if (((TSalD = nil) or (TSalD.GetValue('PDS_SEGMENT') > 'S44')) and
             ((TD2 = '01') or (TD2 = '03') or (TD2 = '07'))) then
             begin
             if ((StPrecSeg > 'S30') and (S44 = False) and
                (S90000 = False) and (Ircantec=False)) then
                begin
                TDADSUDetail:= Tob.Create ('DADSDETAIL', TDADSUD, -1);
                TDADSUDetail.PutValue ('PDS_SALARIE', CodeSal);
                TDADSUDetail.PutValue ('PDS_TYPE', TypeD);
{PT4
                TDADSUDetail.PutValue ('PDS_ORDRE',
                                       TSalD.GetValue('PDS_ORDRE'));
}
                TDADSUDetail.PutValue ('PDS_ORDRE', Ordre);
//FIN PT4                
                TDADSUDetail.PutValue ('PDS_DATEDEBUT', DebExer);
                TDADSUDetail.PutValue ('PDS_DATEFIN', FinExer);
                TDADSUDetail.PutValue ('PDS_EXERCICEDADS', PGExercice);
                TDADSUDetail.PutValue ('PDS_ORDRESEG', 751);
                TDADSUDetail.PutValue ('PDS_SEGMENT', 'S44.G01.00.001');
                TDADSUDetail.PutValue ('PDS_DONNEE', '01');
                MoveCurProgressForm ('Contrôle salarié '+CodeSal);
                MoveCur (False);

                TDADSUDetail:= Tob.Create ('DADSDETAIL', TDADSUD, -1);
                TDADSUDetail.PutValue ('PDS_SALARIE', CodeSal);
                TDADSUDetail.PutValue ('PDS_TYPE', TypeD);
                TDADSUDetail.PutValue ('PDS_ORDRE',
                                       TSalD.GetValue('PDS_ORDRE'));
                TDADSUDetail.PutValue ('PDS_DATEDEBUT', DebExer);
                TDADSUDetail.PutValue ('PDS_DATEFIN', FinExer);
                TDADSUDetail.PutValue ('PDS_EXERCICEDADS', PGExercice);
                TDADSUDetail.PutValue ('PDS_ORDRESEG', 702);
                TDADSUDetail.PutValue ('PDS_SEGMENT', 'S44.G01.00.002');
                TDADSUDetail.PutValue ('PDS_DONNEE',
                                       IntToStr(StrToInt(StNbreH)*100));
                MoveCurProgressForm ('Contrôle salarié '+CodeSal);
                MoveCur(False);
                end;
             S44:= True;
             end;

          if (Next = True) then
             Next:= False
          else
             begin
             MoveCurProgressForm ('Contrôle salarié '+CodeSal);
             MoveCur(False);
             if TSalD = nil then
                StPrecSeg:= Copy(TSalD.GetValue('PDS_SEGMENT'),1,3);
             TDADSUDetail:= Tob.Create ('DADSDETAIL', TDADSUD, -1);
             TDADSUDetail.Dupliquer (TSalD, True, True);
             TSalD:= TSal.FindNext([''],[''],FALSE);
             end;
          if TSalD = nil then
             break
          else
             CodeSal:= TSalD.GetValue('PDS_SALARIE');
          end;

      if ((GetControlText('CBPRUDH')='X') and (S41G02=False) and
         (TPrudh<>nil)) then
         begin
         NbreSeg3:= TPrudh.FillesCount(1);
         TPrudhD:= TPrudh.FindFirst([''],[''],FALSE);
         For j := 1 to NbreSeg3 do
             begin
             TDADSUDetail:= Tob.Create ('DADSDETAIL', TDADSUD, -1);
             TDADSUDetail.PutValue ('PDS_SALARIE',
                                    TPrudhD.GetValue('PDS_SALARIE'));
             TDADSUDetail.PutValue ('PDS_TYPE', TypeD);
             TDADSUDetail.PutValue ('PDS_ORDRE', TPrudhD.GetValue('PDS_ORDRE'));
             TDADSUDetail.PutValue ('PDS_DATEDEBUT', DebExer);
             TDADSUDetail.PutValue ('PDS_DATEFIN', FinExer);
             TDADSUDetail.PutValue ('PDS_EXERCICEDADS', PGExercice);
             TDADSUDetail.PutValue ('PDS_ORDRESEG',
                                    TPrudhD.GetValue('PDS_ORDRESEG'));
             TDADSUDetail.PutValue ('PDS_SEGMENT',
                                    TPrudhD.GetValue('PDS_SEGMENT'));
             TDADSUDetail.PutValue ('PDS_DONNEE',
                                    TPrudhD.GetValue('PDS_DONNEE'));
             MoveCurProgressForm ('Contrôle salarié '+CodeSal);
             MoveCur (False);
             TPrudhD:= TPrudh.FindNext([''],[''],FALSE);
             end;
         S41G02:= True;
         end;

      if ((TD2='01') or (TD2='03') or (TD2='07')) then
         begin
         if ((StPrecSeg > 'S30') and (S44 = False) and (Ircantec=False)) then
            begin
            TDADSUDetail:= Tob.Create ('DADSDETAIL', TDADSUD, -1);
            TDADSUDetail.PutValue ('PDS_SALARIE',
                                   TSalD.GetValue('PDS_SALARIE'));
            TDADSUDetail.PutValue ('PDS_TYPE', TypeD);
            TDADSUDetail.PutValue ('PDS_ORDRE', TSalD.GetValue('PDS_ORDRE'));
            TDADSUDetail.PutValue ('PDS_DATEDEBUT', DebExer);
            TDADSUDetail.PutValue ('PDS_DATEFIN', FinExer);
            TDADSUDetail.PutValue ('PDS_EXERCICEDADS', PGExercice);
            TDADSUDetail.PutValue ('PDS_ORDRESEG', 751);
            TDADSUDetail.PutValue ('PDS_SEGMENT', 'S44.G01.00.001');
            TDADSUDetail.PutValue ('PDS_DONNEE', '01');
            MoveCurProgressForm ('Contrôle salarié '+CodeSal);
            MoveCur(False);

            TDADSUDetail:= Tob.Create ('DADSDETAIL', TDADSUD, -1);
            TDADSUDetail.PutValue ('PDS_SALARIE', CodeSal);
            TDADSUDetail.PutValue ('PDS_TYPE', TypeD);
            TDADSUDetail.PutValue ('PDS_ORDRE', TSalD.GetValue('PDS_ORDRE'));
            TDADSUDetail.PutValue ('PDS_DATEDEBUT', DebExer);
            TDADSUDetail.PutValue ('PDS_DATEFIN', FinExer);
            TDADSUDetail.PutValue ('PDS_EXERCICEDADS', PGExercice);
            TDADSUDetail.PutValue ('PDS_ORDRESEG', 702);
            TDADSUDetail.PutValue ('PDS_SEGMENT', 'S44.G01.00.002');
            TDADSUDetail.PutValue ('PDS_DONNEE',
                                   IntToStr(StrToInt(StNbreH)*100));
            MoveCurProgressForm ('Contrôle salarié '+CodeSal);
            MoveCur(False);
            end;
         end;
      end;
   FreeAndNil(TSal);

   StInact:= 'SELECT PDS_SALARIE, PDS_SEGMENT, PDS_DATEDEBUT, PDS_DONNEE,'+
             ' PDS_ORDRE, PDS_ORDRESEG'+
             ' FROM DADSDETAIL'+
             ' LEFT JOIN SALARIES ON'+
             ' PDS_SALARIE = PSA_SALARIE'+
             ' LEFT JOIN DADSLEXIQUE ON'+
             ' PDS_SEGMENT=PDL_DADSSEGMENT WHERE'+
             ' PDS_SALARIE = "'+Matricule+'" AND'+
             ' PDS_TYPE = "'+TypeD+'" AND'+
             ' PDS_ORDRE < 0 AND'+
             ' PDS_EXERCICEDADS = "'+PGExercice+'" AND'+
             ' PSA_DADSFRACTION="'+GetControlText ('DADSFRACTION')+'" AND'+
             ' PDL_EXERCICEDEB <= "'+PGExercice+'" AND'+
             ' (PDL_EXERCICEFIN >= "'+PGExercice+'" OR'+
             ' PDL_EXERCICEFIN = "")'+StTypeDecla+
             ' ORDER BY PDS_DATEDEBUT,PDS_ORDRE DESC, PDS_ORDRESEG,'+
             ' PDS_SEGMENT, PDS_DATEFIN';
   QRechInact:= OpenSql(StInact,TRUE);
   TInact:= TOB.Create ('Les éléments Inactivité', NIL, -1);
   TInact.LoadDetailDB ('DADSDETAIL','','',QRechInact,False);
   Ferme(QRechInact);

   NbreSeg2:= TInact.FillesCount(1);
   TInactD:= TInact.FindFirst([''],[''],FALSE);
   if (TInactD <> nil) then
      begin
      For j:=1 to NbreSeg2 do
          begin
          TDADSUDetail:= Tob.Create ('DADSDETAIL', TDADSUD, -1);
          TDADSUDetail.PutValue ('PDS_SALARIE',
                                 TInactD.GetValue('PDS_SALARIE'));
          TDADSUDetail.PutValue ('PDS_TYPE', TypeD);
          TDADSUDetail.PutValue ('PDS_ORDRE', TInactD.GetValue('PDS_ORDRE'));
          TDADSUDetail.PutValue ('PDS_DATEDEBUT', DebExer);
          TDADSUDetail.PutValue ('PDS_DATEFIN', FinExer);
          TDADSUDetail.PutValue ('PDS_EXERCICEDADS', PGExercice);
          TDADSUDetail.PutValue ('PDS_ORDRESEG',
                                 TInactD.GetValue('PDS_ORDRESEG'));
          TDADSUDetail.PutValue ('PDS_SEGMENT',
                                 TInactD.GetValue('PDS_SEGMENT'));
          TDADSUDetail.PutValue ('PDS_DONNEE', TInactD.GetValue('PDS_DONNEE'));
          MoveCurProgressForm ('Contrôle salarié '+CodeSal);
          MoveCur(False);
          TInactD:= TInact.FindNext([''],[''],FALSE);
          end;
      end;
   FreeAndNil(TInact);
   end;

//Données Honoraires
if ((GetControlText ('CHNEANT')<>'X') and
   (GetControlText('CHHONORAIRE')='X')) then
   begin
   DeleteErreur ('', 'S70');
{PT9
   if (THMultiValCombobox(GetControl('MCETAB')).Tous) then
}
   if ((THMultiValCombobox(GetControl('MCETAB')).Tous) and
      (TEtabDADS=nil)) then
//FIN PT9
      StEtabGen:=''
   else
      StEtabGen:= ' AND'+
                  ' PDS_ORDRE IN (SELECT PDS_ORDRE FROM DADSDETAIL WHERE'+
                  ' PDS_SEGMENT="S70.G01.00.014" AND'+
                  ' PDS_DATEDEBUT >= "'+UsDateTime(DebExer)+'" AND'+
                  ' PDS_DATEFIN <= "'+UsDateTime(FinExer)+'" AND'+
                  ' PDS_EXERCICEDADS = "'+PGExercice+'" AND ('+
                  StEtab+'))';
   StHonor:= 'SELECT PDS_SALARIE, PDS_SEGMENT, PDS_DONNEE, PDS_ORDRE,'+
             ' PDS_ORDRESEG'+
             ' FROM DADSDETAIL'+
             ' LEFT JOIN DADSLEXIQUE ON'+
             ' PDS_SEGMENT=PDL_DADSSEGMENT WHERE'+
             ' PDS_SALARIE LIKE "--H%" AND'+
             ' PDS_TYPE="'+TypeD+'" AND'+
             ' PDS_DATEDEBUT>="'+UsDateTime(DebExer)+'" AND'+
             ' PDS_DATEFIN<="'+UsDateTime(FinExer)+'" AND'+
             ' PDS_EXERCICEDADS = "'+PGExercice+'" AND'+
             ' PDS_ORDRE IN (SELECT PDS_ORDRE FROM DADSDETAIL WHERE'+
             ' PDS_SEGMENT="S70.G01.01.002" AND'+
             ' PDS_SALARIE LIKE "--H%" AND PDS_TYPE="'+TypeD+'" AND'+
             ' PDS_DATEDEBUT>="'+UsDateTime(DebExer)+'" AND'+
             ' PDS_DATEFIN<="'+UsDateTime(FinExer)+'") AND'+
             ' PDS_EXERCICEDADS = "'+PGExercice+'" AND'+
             ' PDL_EXERCICEDEB <= "'+PGExercice+'" AND'+
             ' (PDL_EXERCICEFIN >= "'+PGExercice+'" OR'+
             ' PDL_EXERCICEFIN = "")'+StTypeDecla+StEtabGen+
             ' ORDER BY PDS_ORDRE,PDS_ORDRESEG,PDS_SEGMENT,'+
             ' PDS_DATEDEBUT,PDS_DATEFIN';
   QRechHonor:=OpenSql(StHonor,TRUE);
   TSal:= TOB.Create('Les éléments honoraires', NIL, -1);
   TSal.LoadDetailDB ('DADSDETAIL','','',QRechHonor,False);
   Ferme(QRechHonor);

   NbreSeg:= TSal.FillesCount(1);
   TSalD:= TSal.FindFirst([''],[''],FALSE);
   if (TSalD <> nil) then
      begin
      For i := 1 to NbreSeg do
          begin
          TDADSUDetail:= Tob.Create ('DADSDETAIL', TDADSUD, -1);
          TDADSUDetail.PutValue ('PDS_SALARIE',
                                 TSalD.GetValue('PDS_SALARIE'));
          TDADSUDetail.PutValue ('PDS_TYPE', TypeD);
          TDADSUDetail.PutValue ('PDS_ORDRE', TSalD.GetValue('PDS_ORDRE'));
          TDADSUDetail.PutValue ('PDS_DATEDEBUT', DebExer);
          TDADSUDetail.PutValue ('PDS_DATEFIN', FinExer);
          TDADSUDetail.PutValue ('PDS_EXERCICEDADS', PGExercice);
          TDADSUDetail.PutValue ('PDS_ORDRESEG',
                                 TSalD.GetValue('PDS_ORDRESEG'));
          TDADSUDetail.PutValue ('PDS_SEGMENT',
                                 TSalD.GetValue('PDS_SEGMENT'));
          TDADSUDetail.PutValue ('PDS_DONNEE', TSalD.GetValue('PDS_DONNEE'));
          MoveCurProgressForm ('Contrôle Honoraire '+
                               TSalD.GetValue('PDS_SALARIE'));
          MoveCur(False);
          TSalD:= TSal.FindNext([''],[''],FALSE);
          end;
      end;

   StHonor:= 'SELECT PDS_ORDRE'+
             ' FROM DADSDETAIL WHERE'+
             ' PDS_SALARIE LIKE "--H%" AND'+
             ' PDS_TYPE="'+TypeD+'" AND'+
             ' PDS_DATEDEBUT>="'+UsDateTime(DebExer)+'" AND'+
             ' PDS_DATEFIN<="'+UsDateTime(FinExer)+'" AND'+
             ' PDS_EXERCICEDADS = "'+PGExercice+'" AND'+
             ' PDS_ORDRE NOT IN (SELECT PDS_ORDRE FROM DADSDETAIL WHERE'+
             ' PDS_SEGMENT="S70.G01.01.002" AND'+
             ' PDS_SALARIE LIKE "--H%" AND PDS_TYPE="'+TypeD+'" AND'+
             ' PDS_DATEDEBUT>="'+UsDateTime(DebExer)+'" AND'+
             ' PDS_DATEFIN<="'+UsDateTime(FinExer)+'" AND'+
             ' PDS_EXERCICEDADS = "'+PGExercice+'")'+
             ' GROUP BY PDS_ORDRE';
   QRechHonor:= OpenSql(StHonor,TRUE);
   TInact:= TOB.Create('Les éléments honoraires non repris', NIL, -1);
   TInact.LoadDetailDB ('DADSDETAIL','','',QRechHonor,False);
   Ferme(QRechHonor);
   NbreSeg:= TInact.FillesCount(1);
   if (NbreSeg<>0) then
      begin
      TInactD:= TInact.FindFirst([''],[''],FALSE);
      For i:= 1 to NbreSeg do
          begin
          ErreurDADSU.Salarie:= TInactD.GetValue('PDS_SALARIE');
          ErreurDADSU.Num:= TInactD.GetValue('PDS_TYPE');
          ErreurDADSU.DateDeb:= TInactD.GetValue('PDS_DATEFIN');
          ErreurDADSU.DateFin:= TInactD.GetValue('PDS_DATEFIN');
          ErreurDADSU.Segment:= TInactD.GetValue('PDS_SEGMENT');
          ErreurDADSU.Explication:= 'Honoraire non intégré car aucun'+
                                    ' montant de rémunération n''est'+
                                    ' renseigné';
          ErreurDADSU.CtrlBloquant:= False;

          EcrireErreur (ErreurDADSU);
          TInactD:= TInact.FindNext([''],[''],FALSE);
          end;
      end;
   FreeAndNil(TInact);
   FreeAndNil(TSal);
   end;

CreeEtab;
//TobDebug (TDADSUD);
ControleTOB (TDADSUD, False, True);

if (TErreur.Detail.Count=0) then
   begin
   ErreurDADSU.Salarie:= GetParamSoc('SO_LIBELLE');
   ErreurDADSU.Num:= 0;
   ErreurDADSU.DateDeb:= StrToDate (GetControlText ('DATEDEB'));
   ErreurDADSU.DateFin:= StrToDate (GetControlText ('DATEFIN'));
   ErreurDADSU.Segment:= 'SOK';
   ErreurDADSU.Explication:= 'Le dossier ne comporte pas d''erreur';
   ErreurDADSU.CtrlBloquant:= False;

   EcrireErreur (ErreurDADSU);
   end
else
   TErreur.InsertOrUpdateDb (FALSE);

FreeAndNil (TPrudh);
FreeAndNil (TEtabDADS);     //PT9
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 21/06/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... : PAYE;PGDADSU;PGCONTROLE
*****************************************************************}
procedure TOF_PG_EDITDADSCONT.ChargerLexiqueDads (Annee, NatureDADSU:String);
var
StQ, StTypeDecla : string;
Q : TQuery;
begin
StTypeDecla:= GetDADSUNature (NatureDADSU);
StQ:= 'SELECT *'+
      ' FROM DADSLEXIQUE WHERE'+
      ' PDL_EXERCICEDEB <= "'+Annee+'" AND'+
      ' (PDL_EXERCICEFIN >= "'+Annee+'" OR PDL_EXERCICEFIN = "")'+
      ' ORDER BY PDL_DADSSEGMENT';
Q:= OpenSQL (StQ, True);
TLexique.LoadDetailDB('DADSLEXIQUE','','',Q,False);
Ferme(Q);
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 21/06/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... : PAYE;PGDADSU;PGCONTROLE
*****************************************************************}
function TOF_PG_EDITDADSCONT.AffecteDebut(T : TOB) : integer;
begin
T.PutValue ('PDS_DATEDEBUT', StrToDate (PerDeb));
result:= 0;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 21/06/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... : PAYE;PGDADSU;PGCONTROLE
*****************************************************************}
function TOF_PG_EDITDADSCONT.AffecteFin(T : TOB) : integer;
begin
T.PutValue ('PDS_DATEFIN', StrToDate (PerFin));
result:= 0;
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 31/03/2006
Modifié le ... :   /  /
Description .. : Procédure de contrôle sans mise en table des erreurs mais avec
Suite ........ : constitution d'une TOB basée sur la structure de la table
Suite ........ : DADSCONTROLE
Mots clefs ... : PAYE;PGDADSU;PGCONTROLE
*****************************************************************}
procedure TOF_PG_EDITDADSCONT.ControleStandard ();
var
DateDeb, DateDebMMJJ, DateFin, DateFinMMJJ, ExerciceDads, FileN, st, Exer : String;
Fraction, Libelle, NatureDADSU, Nom, Periodicite, Prenom, S, Segment : String;
Structure, ValeurSeg : String;
FLect : TextFile;
TD, TErreurDetail, TobEmetteur, TobEmetteurD, TobEntreprise, TobEtabliss : TOB;
TobFille, TobHonoraire, TobSalaries, TobTotaux, TobTotauxD : TOB;
TobUneEntreprise, TobUneEntrepriseD, TobUnEtab, TobUnEtabD : TOB;
TobUnHonoraire, TobUnHonoraireD, TobUnSalarie, TobUnSalarieD : TOB;
Structure2003 : boolean;
LL, Ordre, Virgule : integer;
Q: TQuery; // PT10-3
AnneeRef : THValCombobox; // PT10-3


begin
//Récupération du fichier
FileN:= GetControlText ('FICHIER');

if (FileN='') then
   begin
   PGIBox ('Aucun fichier sélectionné', Ecran.Caption);
   exit;
   end;

AssignFile (FLect, FileN);
Reset (FLect);

TD:= Tob.Create ('DetailDADSU', Nil, -1);
TobEmetteur:= Tob.Create ('S10_M', TD, -1);
TobEntreprise:= Tob.Create ('S20_M', TD, -1);
TobSalaries:= Tob.Create ('S30_M', TD, -1);
TobHonoraire:= Tob.Create ('S70_M', TD, -1);
TobEtabliss:= Tob.Create ('S80_M', TD, -1);
TobTotaux:= Tob.Create ('S90_M', TD, -1);
TLexique:= Tob.Create ('DADSLEXIQUE', Nil, -1);

ExerciceDads:= '';
Structure2003:= False;

AnneeRef:= THValCombobox(GetControl('ANNEE')); // PT10-3

Readln (FLect,S);
LL:= Length (S);
Virgule:= Pos(',', S);
Segment:= Copy (S,1,Virgule-1);
Structure:= Copy (S,1,3);
ValeurSeg:= Copy (S, Virgule+2, LL-(Virgule+2));

while (not(eof(FLect))) do
      begin
      if (Segment = 'S10.G01.00.001.001') then
         begin
         Ordre:= 0;
         Libelle:= '';
         TypeD:= '';

         while (Structure = 'S10') do
               begin
               TobEmetteurD:= Tob.Create ('S10_F', TobEmetteur, -1);
               TobEmetteurD.AddChampSup ('PDS_SALARIE', False);
               TobEmetteurD.AddChampSup ('PDS_LIBELLE', False);
               TobEmetteurD.AddChampSup ('PDS_TYPE', False);
               TobEmetteurD.AddChampSup ('PDS_ORDRE', False);
               TobEmetteurD.AddChampSup ('PDS_DATEDEBUT', False);
               TobEmetteurD.AddChampSup ('PDS_DATEFIN', False);
               TobEmetteurD.AddChampSup ('PDS_EXERCICEDADS', False);
               TobEmetteurD.AddChampSup ('PDS_SEGMENT', False);
               TobEmetteurD.AddChampSup ('PDS_DONNEE', False);
               if (Segment='S10.G01.00.002') then
                  begin
                  Libelle:= ValeurSeg;
                  TobEmetteur.PutValueAllFille ('PDS_SALARIE', Libelle);
                  TobEmetteur.PutValueAllFille ('PDS_LIBELLE', Libelle);
                  end
               else
               if (Libelle <> '') then
                  begin
                  TobEmetteurD.PutValue ('PDS_SALARIE', Libelle);
                  TobEmetteurD.PutValue ('PDS_LIBELLE', Libelle);
                  end;

               if (Segment='S10.G01.01.001.001') then
                  Ordre:= Ordre+1;
               TobEmetteurD.PutValue ('PDS_ORDRE', Ordre);
               TobEmetteurD.PutValue ('PDS_SEGMENT', Segment);
               TobEmetteurD.PutValue ('PDS_DONNEE', ValeurSeg);
               Readln (FLect,S);
               LL:= Length (S);
               Virgule:= Pos(',', S);
               Segment:= Copy (S,1,Virgule-1);
               Structure:= Copy (S,1,3);
               ValeurSeg:= Copy (S, Virgule+2, LL-(Virgule+2));
               end;
         end;

      if (Segment = 'S20.G01.00.001') then
         begin
         Ordre:= 0;
         Libelle:= '';
         TobUneEntreprise:= Tob.Create ('S20_P', TobEntreprise, -1);

         while (Structure = 'S20') do
               begin
               TobUneEntrepriseD:= Tob.Create ('S20_F', TobUneEntreprise, -1);
               TobUneEntrepriseD.AddChampSup ('PDS_SALARIE', False);
               TobUneEntrepriseD.AddChampSup ('PDS_LIBELLE', False);
               TobUneEntrepriseD.AddChampSup ('PDS_TYPE', False);
               TobUneEntrepriseD.AddChampSup ('PDS_ORDRE', False);
               TobUneEntrepriseD.AddChampSup ('PDS_DATEDEBUT', False);
               TobUneEntrepriseD.AddChampSup ('PDS_DATEFIN', False);
               TobUneEntrepriseD.AddChampSup ('PDS_EXERCICEDADS', False);
               TobUneEntrepriseD.AddChampSup ('PDS_SEGMENT', False);
               TobUneEntrepriseD.AddChampSup ('PDS_DONNEE', False);
               if (Segment = 'S20.G01.00.002') then
                  begin
                  Libelle:= ValeurSeg;
                  TobUneEntreprise.PutValueAllFille ('PDS_SALARIE', Libelle);
                  TobUneEntreprise.PutValueAllFille ('PDS_LIBELLE', Libelle);
                  SetControlText ('DOSSIER', Libelle);
                  end
               else
               if (Libelle <> '') then
                  begin
                  TobUneEntrepriseD.PutValue ('PDS_SALARIE', Libelle);
                  TobUneEntrepriseD.PutValue ('PDS_LIBELLE', Libelle);
                  end;

               If (Segment='S20.G01.00.003.001') then
                  begin
                  DateDeb:= Copy (ValeurSeg, 1, 2)+'/'+Copy (ValeurSeg, 3, 2)+
                            '/'+Copy (ValeurSeg, 5, 4);
                  DateDebMMJJ:= Copy (ValeurSeg, 3, 2)+Copy (ValeurSeg, 1, 2);
                  end;

               if (Segment='S20.G01.00.003.002') then
                  begin
                  DateFin:= Copy (ValeurSeg, 1, 2)+'/'+Copy (ValeurSeg, 3, 2)+
                            '/'+Copy (ValeurSeg, 5, 4);
                  DateFinMMJJ:= Copy (ValeurSeg, 3, 2)+Copy (ValeurSeg, 1, 2);
                  ExerciceDads:= Copy (ValeurSeg, 5, 4);
                  // Deb PT10-3
                  //  SetControlText ('ANNEE', ExerciceDads);
                  st := 'SELECT PEX_EXERCICE FROM EXERSOCIAL WHERE PEX_ANNEEREFER = "'+ExerciceDADS+'"';
                  Q:= opensql(St, true);
                  if not Q.EOF then exer := Q.findfield('PEX_EXERCICE').asstring else exer := '';
                  AnneeRef.Value:= Exer;
                  // fin PT10-3
                  end;

               If (Segment='S20.G01.00.003') then
               // Recherche de l'année de la DADS
                  begin
                  ExerciceDads:= Copy (ValeurSeg, 5, 4);
                  Structure2003:= True;
                  // deb PT10-3
                  //  SetControlText ('ANNEE', ExerciceDads);
                  st := 'SELECT PEX_EXERCICE FROM EXERSOCIAL WHERE PEX_ANNEEREFER = "'+ExerciceDADS+'"';
                  Q:= opensql(St, true);
                  if not Q.EOF then exer := Q.findfield('PEX_EXERCICE').asstring else exer := '';
                  AnneeRef.Value:= Exer;
                  // fin PT10-3
                  end;

               if (Segment='S20.G01.00.004.001') then
               //Recherche du type de DADS
                  begin
                  NatureDADSU:= ValeurSeg;
                  if (Structure2003=True) then
                     begin
                     if (NatureDADSU<>'04') then
                        begin
                        DateDeb:= '01/01/'+ExerciceDads;
                        DateFin:= '31/12/'+ExerciceDads;
                        TypeD:= '001';
                        end
                     else
                        begin
                        DateDeb:= '01/04/'+IntToStr(StrToInt(ExerciceDads)-1);
                        DateFin:= '31/03/'+ExerciceDads;
                        TypeD:= '002';
                        end;
                     end;
                  if (NatureDADSU='04') then
                     ExerciceDads:= IntToStr(StrToInt(ExerciceDads)-1);
                  end;

               if (Segment='S20.G01.00.004.002') then
               //Recherche du type de DADS
                  begin
                  if (ValeurSeg='52') then
                     TypeD:= '2'+Copy (TypeD, 2, 2);

                  if (ValeurSeg='55') then
                     SetControlChecked ('CHNEANT', True);
                  end;

               If (Segment='S20.G01.00.005') then
                  Fraction:= Copy (ValeurSeg, 1, 1);

               If (Segment='S20.G01.00.018') then
               // Recherche de la périodicité de la DADS
                  begin
                  if (NatureDADSU<>'04') then
                     TypeD:= Copy (TypeD, 1, 1)+'01'
                  else
                  if ((Copy (ValeurSeg, 1, 1)='A')) then
                     TypeD:= Copy (TypeD, 1, 1)+'02'
                  else
                  if ((Copy (ValeurSeg, 1, 1)='S')) then
                     TypeD:= Copy (TypeD, 1, 1)+'03'
                  else
                  if ((Copy (ValeurSeg, 1, 1)='T')) then
                     TypeD:= Copy (TypeD, 1, 1)+'04'
                  else
                  if ((Copy (ValeurSeg, 1, 1)='M')) then
                     TypeD:= Copy (TypeD, 1, 1)+'05';
                  Periodicite:= ValeurSeg;
                  end;

               TobUneEntrepriseD.PutValue ('PDS_ORDRE', Ordre);
               TobUneEntrepriseD.PutValue ('PDS_SEGMENT', Segment);
               TobUneEntrepriseD.PutValue ('PDS_DONNEE', ValeurSeg);
               Readln (FLect,S);
               LL:= Length (S);
               Virgule:= Pos(',', S);
               Segment:= Copy (S,1,Virgule-1);
               Structure:= Copy (S,1,3);
               ValeurSeg:= Copy (S, Virgule+2, LL-(Virgule+2));
               end;
         if (Structure>'S20') then
            begin
            TobEmetteur.PutValueAllFille ('PDS_TYPE', TypeD);
            TobEmetteur.PutValueAllFille ('PDS_EXERCICEDADS', ExerciceDads);
            TobEmetteur.PutValueAllFille ('PDS_DATEDEBUT', StrToDate (DateDeb));
            TobEmetteur.PutValueAllFille ('PDS_DATEFIN', StrToDate (DateFin));
            TobUneEntreprise.PutValueAllFille ('PDS_TYPE', TypeD);
            TobUneEntreprise.PutValueAllFille ('PDS_EXERCICEDADS',
                                               ExerciceDads);
            TobUneEntreprise.PutValueAllFille ('PDS_DATEDEBUT',
                                               StrToDate (DateDeb));
            TobUneEntreprise.PutValueAllFille ('PDS_DATEFIN',
                                               StrToDate (DateFin));
            SetControlText ('DATEDEB', DateDeb);
            SetControlText ('DATEFIN', DateFin);
            SetControlText ('CNATURE', NatureDADSU+'51');
            SetControlText ('CCAR', Periodicite);
            SetControlText ('DADSFRACTION', Fraction);
            end;
         end;

{Construction tob qui contient les définitions des champs correspondant aux
segments à partir table DADSLEXIQUE}
      if ((ExerciceDads<>'') and (NatureDADSU<>'') and
         (TLexique.Detail.Count=0)) then
         ChargerLexiqueDads(ExerciceDads, NatureDADSU);

      if (Segment='S30.G01.00.001') then
      //Si nouveau salarié
         begin
         Ordre:= 0;
         Libelle:= '';
         Nom:= '';
         Prenom:= '';
         TobUnSalarie:= Tob.Create ('S30_P', TobSalaries, -1);

         while (Structure = 'S30') do
               begin
               TobUnSalarieD:= Tob.Create ('S30_F', TobUnSalarie, -1);
               TobUnSalarieD.AddChampSup ('PDS_SALARIE', False);
               TobUnSalarieD.AddChampSup ('PDS_LIBELLE', False);
               TobUnSalarieD.AddChampSup ('PDS_TYPE', False);
               TobUnSalarieD.AddChampSup ('PDS_ORDRE', False);
               TobUnSalarieD.AddChampSup ('PDS_DATEDEBUT', False);
               TobUnSalarieD.AddChampSup ('PDS_DATEFIN', False);
               TobUnSalarieD.AddChampSup ('PDS_EXERCICEDADS', False);
               TobUnSalarieD.AddChampSup ('PDS_SEGMENT', False);
               TobUnSalarieD.AddChampSup ('PDS_DONNEE', False);
               TobUnSalarieD.PutValue ('PDS_TYPE', TypeD);
               TobUnSalarieD.PutValue ('PDS_ORDRE', Ordre);
               TobUnSalarieD.PutValue ('PDS_SEGMENT', Segment);
               TobUnSalarieD.PutValue ('PDS_DONNEE', ValeurSeg);
               TobUnSalarieD.PutValue ('PDS_DATEDEBUT', StrToDate (DateDeb));
               TobUnSalarieD.PutValue ('PDS_DATEFIN', StrToDate (DateFin));
               TobUnSalarieD.PutValue ('PDS_EXERCICEDADS', ExerciceDads);
               if (Segment = 'S30.G01.00.002') then
                  Nom:= ValeurSeg
               else
               if (Segment = 'S30.G01.00.003') then
                  Prenom:= ValeurSeg
               else
               if (Segment = 'S30.G01.00.004') then
                  Nom:= ValeurSeg;
               Readln (FLect,S);
               LL:= Length (S);
               Virgule:= Pos(',', S);
               Segment:= Copy (S,1,Virgule-1);
               Structure:= Copy (S,1,3);
               ValeurSeg:= Copy (S, Virgule+2, LL-(Virgule+2));
               end;
         end;

      if (Segment='S41.G01.00.001') then
      //Si nouvelle période
         begin
         PerDeb:= '';
         PerFin:= '';

         while ((Structure = 'S41') or (Structure = 'S42') or
               (Structure = 'S43') or (Structure = 'S44') or
               (Structure = 'S45') or (Structure = 'S46') or
               (Structure = 'S51') or (Structure = 'S53') or
               (Structure = 'S66')) do
               begin
               if ((Segment = 'S41.G01.00.001') and (Ordre <>0)) then
                  begin
                  TobUnSalarie.ParcoursTraitement (['PDS_ORDRE'], [Ordre],
                                                   False, AffecteDebut);
                  TobUnSalarie.ParcoursTraitement (['PDS_ORDRE'], [Ordre],
                                                   False, AffecteFin);
                  end;
               TobUnSalarieD:= Tob.Create ('S30_F', TobUnSalarie, -1);
               TobUnSalarieD.AddChampSup ('PDS_SALARIE', False);
               TobUnSalarieD.AddChampSup ('PDS_LIBELLE', False);
               TobUnSalarieD.AddChampSup ('PDS_TYPE', False);
               TobUnSalarieD.AddChampSup ('PDS_ORDRE', False);
               TobUnSalarieD.AddChampSup ('PDS_DATEDEBUT', False);
               TobUnSalarieD.AddChampSup ('PDS_DATEFIN', False);
               TobUnSalarieD.AddChampSup ('PDS_EXERCICEDADS', False);
               TobUnSalarieD.AddChampSup ('PDS_SEGMENT', False);
               TobUnSalarieD.AddChampSup ('PDS_DONNEE', False);
               if (Segment = 'S41.G01.00.001') then
                  begin
                  Ordre:= Ordre+1;
{PT7-1
                  PerDeb:= Copy (ValeurSeg, 3, 2)+Copy (ValeurSeg, 1, 2);
                  if (((DateDebMMJJ<=PerDeb) and (PerDeb<=DateFinMMJJ)) or
                     (PerDeb<=DateDebMMJJ)) then
                     PerDeb:= Copy (ValeurSeg, 1, 2)+'/'+Copy (ValeurSeg, 3, 2)+
                              '/'+ExerciceDADS
                  else
                     PerDeb:= Copy (ValeurSeg, 1, 2)+'/'+Copy (ValeurSeg, 3, 2)+
                              '/'+IntToStr(StrToInt(ExerciceDads)-1);
}
                  PerDeb:= Copy (ValeurSeg, 1, 2)+'/'+Copy (ValeurSeg, 3, 2)+
                           '/'+Copy (ValeurSeg, 5, 4);
//FIN PT7-1
                  end;

               if (Segment = 'S41.G01.00.003') then
{PT7-1
                  begin
                  PerFin:= Copy (ValeurSeg, 3, 2)+Copy (ValeurSeg, 1, 2);
                  if (((DateDebMMJJ<=PerFin) and (PerFin<=DateFinMMJJ)) or
                     (PerFin<=DateDebMMJJ)) then
                     PerFin:= Copy (ValeurSeg, 1, 2)+'/'+Copy (ValeurSeg, 3, 2)+
                              '/'+ExerciceDADS
                  else
                     PerFin:= Copy (ValeurSeg, 1, 2)+'/'+Copy (ValeurSeg, 3, 2)+
                              '/'+IntToStr(StrToInt(ExerciceDads)-1);
                  end;
}
                     PerFin:= Copy (ValeurSeg, 1, 2)+'/'+Copy (ValeurSeg, 3, 2)+
                              '/'+Copy (ValeurSeg, 5, 4);
//FIN PT7-1

               if (Segment = 'S41.G01.00.009') then
                  SetControlText ('DECALAGE', ValeurSeg);

               if (Segment = 'S41.G01.00.019') then
                  begin
                  Libelle:= ValeurSeg;
                  TobUnSalarie.PutValueAllFille ('PDS_SALARIE', Libelle);
                  end
               else
               if (Libelle <> '') then
                  TobUnSalarieD.PutValue ('PDS_SALARIE', Libelle);

               if (Segment = 'S41.G02.00.008') then
                  SetControlText ('CBPRUDH', 'X'); 

               if (Segment = 'S45.G01.00.001') then
                  SetControlText ('CHPREVOYANCE', 'X');

               TobUnSalarieD.PutValue ('PDS_TYPE', TypeD);
               TobUnSalarieD.PutValue ('PDS_ORDRE', Ordre);
               TobUnSalarieD.PutValue ('PDS_SEGMENT', Segment);
               TobUnSalarieD.PutValue ('PDS_DONNEE', ValeurSeg);
               TobUnSalarieD.PutValue ('PDS_EXERCICEDADS', ExerciceDads);
               TobUnSalarieD.PutValue ('PDS_DATEDEBUT', IDate1900);
               TobUnSalarieD.PutValue ('PDS_DATEFIN', IDate1900);
               Readln (FLect,S);
               LL:= Length (S);
               Virgule:= Pos(',', S);
               Segment:= Copy (S,1,Virgule-1);
               Structure:= Copy (S,1,3);
               ValeurSeg:= Copy (S, Virgule+2, LL-(Virgule+2));
               end;

         TobUnSalarie.ParcoursTraitement (['PDS_ORDRE'], [Ordre], False,
                                          AffecteDebut);
         TobUnSalarie.ParcoursTraitement (['PDS_ORDRE'], [Ordre], False,
                                          AffecteFin);
         TobUnSalarie.PutValueAllFille ('PDS_LIBELLE', Nom+' '+Prenom);
         end;

      if ((Segment='S70.G01.00.001') or (Segment='S70.G01.00.002.001')) then
      //Si nouvel honoraire
         begin
         Ordre:= 0;
         Libelle:= '';
         SetControlText ('CHHONORAIRE', 'X');

         while (Structure = 'S70') do
               begin
               if ((Segment='S70.G01.00.001') or
                  (Segment='S70.G01.00.002.001')) then
                  begin
                  Ordre:= Ordre+1;
                  Libelle:= ValeurSeg;
                  TobUnHonoraire:= Tob.Create ('S70_P', TobHonoraire, -1);
                  end;

               TobUnHonoraireD:= Tob.Create ('S70_F', TobUnHonoraire, -1);
               TobUnHonoraireD.AddChampSup ('PDS_SALARIE', False);
               TobUnHonoraireD.AddChampSup ('PDS_LIBELLE', False);
               TobUnHonoraireD.AddChampSup ('PDS_TYPE', False);
               TobUnHonoraireD.AddChampSup ('PDS_ORDRE', False);
               TobUnHonoraireD.AddChampSup ('PDS_DATEDEBUT', False);
               TobUnHonoraireD.AddChampSup ('PDS_DATEFIN', False);
               TobUnHonoraireD.AddChampSup ('PDS_EXERCICEDADS', False);
               TobUnHonoraireD.AddChampSup ('PDS_SEGMENT', False);
               TobUnHonoraireD.AddChampSup ('PDS_DONNEE', False);
               TobUnHonoraireD.PutValue ('PDS_SALARIE', '--H'+IntToStr (Ordre));
               TobUnHonoraireD.PutValue ('PDS_TYPE', TypeD);
               TobUnHonoraireD.PutValue ('PDS_ORDRE', Ordre);
               TobUnHonoraireD.PutValue ('PDS_SEGMENT', Segment);
               TobUnHonoraireD.PutValue ('PDS_DONNEE', ValeurSeg);
               TobUnHonoraireD.PutValue ('PDS_EXERCICEDADS', ExerciceDads);
               TobUnHonoraireD.PutValue ('PDS_DATEDEBUT', StrToDate (DateDeb));
               TobUnHonoraireD.PutValue ('PDS_DATEFIN', StrToDate (DateFin));
               if ((Segment='S70.G01.00.001') or
                  (Segment='S70.G01.00.002.001')) then
                  begin
                  Libelle:= ValeurSeg;
                  if (Segment='S70.G01.00.001') then
                     TobUnHonoraireD.PutValue ('PDS_LIBELLE', Libelle);
                  end
               else
               if (Segment='S70.G01.00.002.002') then
                  begin
                  Libelle:= Libelle+' '+ValeurSeg;
                  TobUnHonoraire.PutValueAllFille ('PDS_LIBELLE', Libelle);
                  end
               else
               if (Libelle <> '') then
                  TobUnHonoraireD.PutValue ('PDS_LIBELLE', Libelle);
               Readln (FLect,S);
               LL:= Length (S);
               Virgule:= Pos(',', S);
               Segment:= Copy (S,1,Virgule-1);
               Structure:= Copy (S,1,3);
               ValeurSeg:= Copy (S, Virgule+2, LL-(Virgule+2));
               end;
         end;

      if (Segment='S80.G01.00.001.001') then
      //Si nouvel établissement
         begin
         Ordre:= 0;
         Libelle:= '';

         while ((Structure = 'S80') or (Structure = 'S85')) do
               begin
               if (Segment='S80.G01.00.001.001') then
                  begin
                  Libelle:= '';
                  Ordre:= Ordre+1;
                  TobUnEtab:= Tob.Create ('S80_P', TobEtabliss, -1);
                  end;

               TobUnEtabD:= Tob.Create ('S80_F', TobUnEtab, -1);
               TobUnEtabD.AddChampSup ('PDS_SALARIE', False);
               TobUnEtabD.AddChampSup ('PDS_LIBELLE', False);
               TobUnEtabD.AddChampSup ('PDS_TYPE', False);
               TobUnEtabD.AddChampSup ('PDS_ORDRE', False);
               TobUnEtabD.AddChampSup ('PDS_DATEDEBUT', False);
               TobUnEtabD.AddChampSup ('PDS_DATEFIN', False);
               TobUnEtabD.AddChampSup ('PDS_EXERCICEDADS', False);
               TobUnEtabD.AddChampSup ('PDS_SEGMENT', False);
               TobUnEtabD.AddChampSup ('PDS_DONNEE', False);
               TobUnEtabD.PutValue ('PDS_SALARIE', '**'+IntToStr (Ordre));
               TobUnEtabD.PutValue ('PDS_TYPE', TypeD);
               TobUnEtabD.PutValue ('PDS_ORDRE', Ordre);
               TobUnEtabD.PutValue ('PDS_EXERCICEDADS', ExerciceDads);
               TobUnEtabD.PutValue ('PDS_DATEDEBUT', StrToDate (DateDeb));
               TobUnEtabD.PutValue ('PDS_DATEFIN', StrToDate (DateFin));
               TobUnEtabD.PutValue ('PDS_SEGMENT', Segment);
               TobUnEtabD.PutValue ('PDS_DONNEE', ValeurSeg);
               if (Segment='S80.G01.00.002') then
                  begin
                  Libelle:= ValeurSeg;
                  TobUnEtab.PutValueAllFille ('PDS_LIBELLE', Libelle);
                  end
               else
               if (Libelle <> '') then
                  TobUnEtabD.PutValue ('PDS_LIBELLE', Libelle);

               Readln (FLect,S);
               LL:= Length (S);
               Virgule:= Pos(',', S);
               Segment:= Copy (S,1,Virgule-1);
               Structure:= Copy (S,1,3);
               ValeurSeg:= Copy (S, Virgule+2, LL-(Virgule+2));
               end;
         end;

      if (Segment = 'S90.G01.00.001') then
         begin
         Ordre:= 0;
         Libelle:= 'TOTAUX';

         while (Structure = 'S90') do
               begin
               TobTotauxD:= Tob.Create ('S90_F', TobTotaux, -1);
               TobTotauxD.AddChampSup ('PDS_SALARIE', False);
               TobTotauxD.AddChampSup ('PDS_LIBELLE', False);
               TobTotauxD.AddChampSup ('PDS_TYPE', False);
               TobTotauxD.AddChampSup ('PDS_ORDRE', False);
               TobTotauxD.AddChampSup ('PDS_DATEDEBUT', False);
               TobTotauxD.AddChampSup ('PDS_DATEFIN', False);
               TobTotauxD.AddChampSup ('PDS_EXERCICEDADS', False);
               TobTotauxD.AddChampSup ('PDS_SEGMENT', False);
               TobTotauxD.AddChampSup ('PDS_DONNEE', False);
               TobTotauxD.PutValue ('PDS_SALARIE', Libelle);
               TobTotauxD.PutValue ('PDS_LIBELLE', Libelle);
               TobTotauxD.PutValue ('PDS_TYPE', TypeD);
               TobTotauxD.PutValue ('PDS_ORDRE', Ordre);
               TobTotauxD.PutValue ('PDS_DATEDEBUT', StrToDate (DateDeb));
               TobTotauxD.PutValue ('PDS_DATEFIN', StrToDate (DateFin));
               TobTotauxD.PutValue ('PDS_EXERCICEDADS', ExerciceDads);
               TobTotauxD.PutValue ('PDS_SEGMENT', Segment);
               TobTotauxD.PutValue ('PDS_DONNEE', ValeurSeg);
               Readln (FLect,S);
               LL:= Length (S);
               Virgule:= Pos(',', S);
               Segment:= Copy (S,1,Virgule-1);
               Structure:= Copy (S,1,3);
               ValeurSeg:= Copy (S, Virgule+2, LL-(Virgule+2));
               end;
         end;
      end;
CloseFile(FLect);

InitMoveProgressForm (NIL,'Edition en cours', 'Veuillez patienter SVP ...',
                      TD.FillesCount (4),FALSE,TRUE);
InitMove (TD.FillesCount (4), '');

//TobDebug (TD);
ControleTOB (TD, True, False);

if (TErreur.Detail.Count=0) then
   begin
   TErreurDetail:= TOB.Create ('DADSCONTROLE', TErreur, -1);
   TErreurDetail.PutValue ('PSU_SALARIE', GetControlText ('DOSSIER'));
   TErreurDetail.PutValue ('PSU_ORIGINECRIT', 'SOK');
   TErreurDetail.PutValue ('PSU_ORDRE', 0);
   TErreurDetail.PutValue ('PSU_DATEDEBUT',
                           StrToDate (GetControlText ('DATEDEB')));
   TErreurDetail.PutValue ('PSU_DATEFIN',
                           StrToDate (GetControlText ('DATEFIN')));
   TErreurDetail.PutValue ('PSU_SEGMENT', 'SOK');
   TErreurDetail.PutValue ('PSU_EXERCICEDADS', ExerciceDADS);
   TErreurDetail.PutValue ('PSU_EXPLICATION',
                           'Le dossier ne comporte pas d''erreur');
   TErreurDetail.PutValue ('PSU_CTRLBLOQUANT', '-');
   end;

FreeAndNil (TD);
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 12/04/2006
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_PG_EDITDADSCONT.DecalageChange(Sender: TObject);
begin
if (Arg <> 'S') then
   Parametrage;
Decalage:= GetControlText ('DECALAGE');
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 31/03/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_PG_EDITDADSCONT.DateDecaleChange(Sender: TObject);
begin
AglLanceFiche ('PAY', 'DADS_DATE', '', '', '');
SetControlText ('DATEDEB', DateToStr(DebExer));
SetControlText ('DATEFIN', DateToStr(FinExer));
SetControlText ('DATEDECALAGE', DateToStr(FinExer));
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 12/04/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_PG_EDITDADSCONT.DateDecalageChange(Sender: TObject);
begin
DateDecalage:= StrToDate(GetControlText ('DATEDECALAGE'));
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 31/03/2006
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_PG_EDITDADSCONT.Parametrage();
var
Car1, StExer, TD2 : string;
QExer : TQuery;
NbMois : integer;
begin
PGAnnee:= THAnneeRef.value;

StExer:= 'SELECT PEX_DATEDEBUT, PEX_DATEFIN, PEX_ANNEEREFER'+
         ' FROM EXERSOCIAL WHERE'+
         ' PEX_EXERCICE="'+PGAnnee+'"';
QExer:= OpenSQL (StExer,TRUE) ;
if (NOT QExer.EOF) then
   PGExercice:= QExer.FindField ('PEX_ANNEEREFER').AsString
else
   PGExercice:= '';

if ((Nature.Value>='0400') and (Nature.Value<'0500')) then
   begin
   if ((Car.Value>='M00') and (Car.Value <='M99')) then
      begin
      TypeD:= '005';
      if (Car.Value = 'M01') then
         begin
         DebExer:= StrToDate ('01/01/'+PGExercice);
         FinExer:= StrToDate ('31/01/'+PGExercice);
         end
      else
      if (Car.Value = 'M02') then
         begin
         DebExer:= StrToDate ('01/02/'+PGExercice);
         FinExer:= FinDeMois (StrToDate ('28/02/'+PGExercice));
         end
      else
      if (Car.Value = 'M03') then
         begin
         DebExer:= StrToDate ('01/03/'+PGExercice);
         FinExer:= StrToDate ('31/03/'+PGExercice);
         end
      else
      if (Car.Value = 'M04') then
         begin
         DebExer:= StrToDate ('01/04/'+IntToStr (StrToInt (PGExercice)-1));
         FinExer:= StrToDate ('30/04/'+IntToStr (StrToInt (PGExercice)-1));
         end
      else
      if (Car.Value = 'M05') then
         begin
         DebExer:= StrToDate ('01/05/'+IntToStr (StrToInt (PGExercice)-1));
         FinExer:= StrToDate ('31/05/'+IntToStr (StrToInt (PGExercice)-1));
         end
      else
      if (Car.Value = 'M06') then
         begin
         DebExer:= StrToDate ('01/06/'+IntToStr (StrToInt (PGExercice)-1));
         FinExer:= StrToDate ('30/06/'+IntToStr (StrToInt (PGExercice)-1));
         end
      else
      if (Car.Value = 'M07') then
         begin
         DebExer:= StrToDate ('01/07/'+IntToStr (StrToInt (PGExercice)-1));
         FinExer:= StrToDate ('31/07/'+IntToStr (StrToInt (PGExercice)-1));
         end
      else
      if (Car.Value = 'M08') then
         begin
         DebExer:= StrToDate ('01/08/'+IntToStr (StrToInt (PGExercice)-1));
         FinExer:= StrToDate ('31/08/'+IntToStr (StrToInt (PGExercice)-1));
         end
      else
      if (Car.Value = 'M09') then
         begin
         DebExer:= StrToDate ('01/09/'+IntToStr (StrToInt (PGExercice)-1));
         FinExer:= StrToDate ('30/09/'+IntToStr (StrToInt (PGExercice)-1));
         end
      else
      if (Car.Value = 'M10') then
         begin
         DebExer:= StrToDate ('01/10/'+IntToStr (StrToInt (PGExercice)-1));
         FinExer:= StrToDate ('31/10/'+IntToStr (StrToInt (PGExercice)-1));
         end
      else
      if (Car.Value = 'M11') then
         begin
         DebExer:= StrToDate ('01/11/'+IntToStr (StrToInt (PGExercice)-1));
         FinExer:= StrToDate ('30/11/'+IntToStr (StrToInt (PGExercice)-1));
         end
      else
      if (Car.Value = 'M12') then
         begin
         DebExer:= StrToDate ('01/12/'+IntToStr (StrToInt (PGExercice)-1));
         FinExer:= StrToDate ('31/12/'+IntToStr (StrToInt (PGExercice)-1));
         end;
      end;
   if ((Car.Value>='T00') and (Car.Value <='T99')) then
      begin
      TypeD:= '004';
      if (Car.Value = 'T01') then
         begin
{PT10-1
         DebExer:= StrToDate ('01/04/'+IntToStr (StrToInt (PGExercice)-1));
         FinExer:= StrToDate ('30/06/'+IntToStr (StrToInt (PGExercice)-1));
}
         DebExer:= StrToDate ('01/01/'+PGExercice);
         FinExer:= StrToDate ('31/03/'+PGExercice);
         end
      else
      if (Car.Value = 'T02') then
         begin
{
         DebExer:= StrToDate ('01/07/'+IntToStr (StrToInt (PGExercice)-1));
         FinExer:= StrToDate ('30/09/'+IntToStr (StrToInt (PGExercice)-1));
}
         DebExer:= StrToDate ('01/04/'+IntToStr (StrToInt (PGExercice)-1));
         FinExer:= StrToDate ('30/06/'+IntToStr (StrToInt (PGExercice)-1));
         end
      else
      if (Car.Value = 'T03') then
         begin
{
         DebExer:= StrToDate ('01/10/'+IntToStr (StrToInt (PGExercice)-1));
         FinExer:= StrToDate ('31/12/'+IntToStr (StrToInt (PGExercice)-1));
}
         DebExer:= StrToDate ('01/07/'+IntToStr (StrToInt (PGExercice)-1));
         FinExer:= StrToDate ('30/09/'+IntToStr (StrToInt (PGExercice)-1));
         end
      else
      if (Car.Value = 'T04') then
         begin
{
         DebExer:= StrToDate ('01/01/'+PGExercice);
         FinExer:= StrToDate ('31/03/'+PGExercice);
}
         DebExer:= StrToDate ('01/10/'+IntToStr (StrToInt (PGExercice)-1));
         FinExer:= StrToDate ('31/12/'+IntToStr (StrToInt (PGExercice)-1));
//FIN PT10-1
         end;
      end;
   if ((Car.Value>='S00') and (Car.Value <='S99')) then
      begin
      PGIBox ('Périodicité interdite', Ecran.Caption);
      Car.Value:= 'A00';
      end;
   if (Car.Value='A00') then
      begin
      TypeD:= '002';
      DebExer:= StrToDate ('01/04/'+IntToStr (StrToInt (PGExercice)-1));
      FinExer:= StrToDate ('31/03/'+PGExercice);
      end;
   end
else
   begin
   TypeD:= '001';
   DebExer:= QExer.FindField ('PEX_DATEDEBUT').AsDateTime;
   FinExer:= QExer.FindField ('PEX_DATEFIN').AsDateTime;
   end;
Ferme (QExer);

if (GetControlText ('CHCOMPL')='X') then
   TypeD:= '2'+Copy (TypeD, 2, 2);

TD2:= Copy (Nature.Value, 1, 2);
if (TD2>='03') then
   SetControlVisible ('CHHONORAIRE', False)
else
   SetControlVisible ('CHHONORAIRE', True);
if (TD2<>'01') then
   SetControlVisible ('CHPREVOYANCE', False)
else
   SetControlVisible ('CHPREVOYANCE', True);
{PT10-2
if ((TD2='03') or (TD2='08')) then
}
if ((TD2='01') or (TD2='03') or (TD2='08')) then
   SetControlText ('CHPREVOYANCE', 'X')
else
   SetControlText ('CHPREVOYANCE', '-');
if ((TD2='04') or (TD2='09')) then
   SetControlVisible ('CHNEANT', False)
else
   SetControlVisible ('CHNEANT', True);
if ((TD2<>'01') and (TD2<>'02')) then
   SetControlVisible ('CBPRUDH', False)
else
   SetControlVisible ('CBPRUDH', True);
if (TD2<'03') then
   SetControlText ('CBPRUDH', 'X')
else
   SetControlText ('CBPRUDH', '-');

Car1:= Copy (Car.Value, 1, 1);
if (Car1='A') then
   NbMois:= 12
else
if (Car1='S') then
   NbMois:= 6
else
if (Car1='T') then
   NbMois:= 3
else
if (Car1='M') then
   NbMois:= 1
else
   NbMois:= 12;

Decalage:= GetControlText ('DECALAGE');
if (Decalage = '02') then  //Adoption du décalage
   begin
   if (Car1 <> 'M') then
      begin
      FinExer:= PlusDate (DebExer, NbMois-1, 'M');
      FinExer:= PlusDate (FinExer, -1, 'J');
      SetControlText ('L_DATEDECALAGE', 'Date d''adoption');
      SetControlVisible ('L_DATEDECALAGE', True);
      SetControlVisible ('DATEDECALAGE', True);
      SetControlVisible ('BMODIFDATE', True);
      end
   else
      begin
      PGIBox ('Vous devez effectuer une DADS-U Néant', 'Adoption du décalage');
      SetControlEnabled ('BCALCULER', False);
      end;
   end
else
if (Decalage = '04') then             //Suppression du décalage
   begin
   FinExer:= PlusDate (DebExer, NbMois+1, 'M');
   FinExer:= PlusDate (FinExer, -1, 'J');
   SetControlText ('L_DATEDECALAGE', 'Date de suppression');
   SetControlVisible ('L_DATEDECALAGE', True);
   SetControlVisible ('DATEDECALAGE', True);
   SetControlVisible ('BMODIFDATE', True);
   end
else
   begin
   SetControlVisible ('L_DATEDECALAGE', False);
   SetControlVisible ('DATEDECALAGE', False);
   if (((Decalage = '01') and (VH_Paie.PGDecalage=False)) or
      ((Decalage = '03') and (VH_Paie.PGDecalage=True))) then
      SetControlVisible ('BMODIFDATE', False)
   else
      begin
      if (Decalage = '01') then
         begin
         DebExer:= PlusDate (DebExer, 1, 'M');
         FinExer:= PlusDate (FinExer, 1, 'M');
         end
      else
      if (Decalage = '03') then
         begin
         DebExer:= PlusDate (DebExer, -1, 'M');
         FinExer:= PlusDate (FinExer, -1, 'M');
         end;
      SetControlVisible ('BMODIFDATE', True);
      end;
   end;
SetControlText('DATEDECALAGE', DateToStr(FinExer));

SetControlText ('DATEDEB', DateToStr(DebExer));
SetControlText ('DATEFIN', DateToStr(FinExer));
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 12/04/2006
Modifié le ... :   /  /
Description .. : Gestion de modification de la date
Mots clefs ... : PAIE;PGDADSU
*****************************************************************}
procedure TOF_PG_EDITDADSCONT.CarDateChange(Sender: TObject);
begin
if (Arg <> 'S') then
   Parametrage;
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 12/04/2006
Modifié le ... :   /  /
Description .. : Gestion de modification de la nature
Mots clefs ... : PAIE;PGDADSU
*****************************************************************}
procedure TOF_PG_EDITDADSCONT.NatureChange(Sender: TObject);
begin
if (Car <> nil) then
   begin
   if ((Nature.Value>='0400') and (Nature.Value<'0500')) then
      begin
      Car.Enabled := True;
      Car.Visible := True;
      SetControlEnabled('LCAR', True);
      SetControlVisible('LCAR', True);
      end
   else
      begin
      Car.Enabled := False;
      Car.Visible := False;
      Car.Value := 'A00';
      SetControlEnabled('LCAR', False);
      SetControlVisible('LCAR', False);
      end;
   end;

if ((Nature.Value>='0400') and (Nature.Value<'0500') and
   (VH_Paie.PGBTP = False)) then
   begin
   PGIBox('Vous ne gérez pas le module BTP', 'Edition du contrôle DADS-U');
   Nature.Value := '0151';
   end;
if (Arg <> 'S') then
   Parametrage;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 13/04/2006
Modifié le ... :   /  /
Description .. : Modification de la CheckBox Neant
Mots clefs ... : PAIE;PGDADSU
*****************************************************************}
procedure TOF_PG_EDITDADSCONT.NeantChange(Sender: TObject);
begin
if (GetControlText('CHNEANT') = 'X') then
   begin
   SetControlEnabled ('CHPREVOYANCE', False);
   SetControlEnabled ('CHHONORAIRE', False);
   SetControlText('CHPREVOYANCE', '');
   SetControlText('CHHONORAIRE', '');
   end
else
   begin
   SetControlEnabled ('CHPREVOYANCE', True);
   SetControlEnabled ('CHHONORAIRE', True);
   end;
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 06/07/2006
Modifié le ... :   /  /
Description .. : Création des enregistrements ETABLISSEMENT
Mots clefs ... : PAIE,PGDADSU
*****************************************************************}
procedure TOF_PG_EDITDADSCONT.CreeEtab();
var
Buffer, BufDest, BufOrig, CodeAPE, CodeIso, Etab, EtabSelect, Libelle : string;
NomRue, numero, StDate, StEtab, StNbSal, StTestSal : string;
QNbSal, QRechEtab, QTestSal : TQuery;
TEtabD, TEtabSelect : TOB;
NbSal, Num, TaxeSal : integer;
EtabAcreer : TypeEtab;
ErreurDADSU : TControle;
begin
DeleteErreur ('', 'S80');     //PT9
TEtabSelect:= TOB.Create('Les établissements', NIL, -1);
EtabSelect:= GetControlText('MCETAB');
If not THMultiValCombobox(GetControl('MCETAB')).Tous then
   begin
   Etab:= ReadTokenpipe(EtabSelect, ';');
   While (Etab <> '') do
         begin
         TEtabD:= Tob.Create('UnEtab', TEtabSelect, -1);
         TEtabD.AddChampSupValeur('CODE', Etab);
         Etab:= ReadTokenpipe(EtabSelect,';');
         end;
   end
else
   RemplirListe ('TTETABLISSEMENT' ,'ET_FICTIF<>"X"' ,TEtabSelect);   //PT9

TEtabD := TEtabSelect.FindFirst([''],[''],FALSE);
if TEtabD <> nil then
   begin
   Num := 1;
   Etab := '';
   While (TEtabD<>nil) do
         begin
         InitEtab (EtabAcreer);                  //PT2
         EtabAcreer.Code:= '**'+TEtabD.GetValue('CODE');

         ErreurDADSU.Salarie:= '**'+TEtabD.GetValue('CODE');
         ErreurDADSU.TypeD:= TypeD;
         ErreurDADSU.Num:= Num;
         ErreurDADSU.DateDeb:= StrToDate (GetControlText ('DATEDEB'));
         ErreurDADSU.DateFin:= StrToDate (GetControlText ('DATEFIN'));
         ErreurDADSU.Exercice:= PGExercice;
{PT9
         DeleteErreur (ErreurDADSU.Salarie, 'S80');
}

         if (TEtabD.GetValue('CODE') <> Etab) then
            begin
            StEtab := 'SELECT ET_ETABLISSEMENT, ET_SIRET, ET_LIBELLE,'+
                      ' ET_ADRESSE1, ET_ADRESSE2, ET_ADRESSE3, ET_CODEPOSTAL,'+
                      ' ET_VILLE, ET_PAYS, ET_APE, ETB_PRUDH, ETB_PRORATATVA'+
                      ' FROM ETABLISS'+
                      ' LEFT JOIN ETABCOMPL ON'+
                      ' ETB_ETABLISSEMENT=ET_ETABLISSEMENT WHERE'+
                      ' ET_ETABLISSEMENT="'+TEtabD.GetValue('CODE')+'"';
            QRechEtab:=OpenSql(StEtab,TRUE);
            if (not QRechEtab.EOF) then
               begin
               ErreurDADSU.Libelle:= QRechEtab.FindField('ET_LIBELLE').AsString;
               BufOrig := QRechEtab.FindField('ET_SIRET').AsString;
               ForceNumerique(BufOrig, BufDest);
               if ControlSiret(BufDest)=False then
                  begin
                  ErreurDADSU.Segment:= 'S80.G01.00.001.001';
                  ErreurDADSU.Explication:= 'Le SIRET de l''établissement '+
                                            TEtabD.GetValue('CODE')+' n''est'+
                                            ' pas valide';
                  ErreurDADSU.CtrlBloquant:= True;
                  EcrireErreur (ErreurDADSU);
                  end;

               if (QRechEtab.FindField('ET_CODEPOSTAL').AsString <> '') then
                  begin
                  if ((QRechEtab.FindField('ET_PAYS').AsString = 'FRA') and
                     ((QRechEtab.FindField('ET_CODEPOSTAL').AsString < '00000') or
                     (QRechEtab.FindField('ET_CODEPOSTAL').AsString > '99999'))) then
                     begin
                     ErreurDADSU.Segment:= 'S80.G01.00.003.010';
                     ErreurDADSU.Explication:= 'L''adresse de l''établissement '+
                                               TEtabD.GetValue('CODE')+' est'+
                                               ' mal renseignée';
                     ErreurDADSU.CtrlBloquant:= True;
                     EcrireErreur (ErreurDADSU);
                     end;
                  end
               else
                  begin
                  ErreurDADSU.Segment:= 'S80.G01.00.003.010';
                  ErreurDADSU.Explication:= 'L''adresse de l''établissement '+
                                            TEtabD.GetValue('CODE')+' est'+
                                            ' mal renseignée';
                  ErreurDADSU.CtrlBloquant:= True;
                  EcrireErreur (ErreurDADSU);
                  end;

               if QRechEtab.FindField('ET_VILLE').AsString = '' then
                  begin
                  ErreurDADSU.Segment:= 'S80.G01.00.003.012';
                  ErreurDADSU.Explication:= 'L''adresse de l''établissement '+
                                             TEtabD.GetValue('CODE')+' est'+
                                             ' mal renseignée';
                  ErreurDADSU.CtrlBloquant:= True;
                  EcrireErreur (ErreurDADSU);
                  end;

               EtabAcreer.Siret:= BufDest;
               EtabAcreer.Libelle:= QRechEtab.FindField('ET_LIBELLE').AsString;

               Buffer:= '';
               EtabAcreer.Adr_Adresse1:= QRechEtab.FindField('ET_ADRESSE1').AsString;
               EtabAcreer.Adr_Adresse2:= QRechEtab.FindField('ET_ADRESSE2').AsString;
               EtabAcreer.Adr_Adresse3:= QRechEtab.FindField('ET_ADRESSE3').AsString;
               if (EtabAcreer.Adr_Adresse2 <> '') then
                  begin
                  if (EtabAcreer.Adr_Adresse3 <> '') then
                     Buffer:= EtabAcreer.Adr_Adresse2+' '+
                              EtabAcreer.Adr_Adresse3
                  else
                     Buffer:= EtabAcreer.Adr_Adresse2;
                  end
               else
                  if (EtabAcreer.Adr_Adresse3 <> '') then
                     Buffer:= EtabAcreer.Adr_Adresse3;

               if (Buffer <> '') then
                  EtabAcreer.Adr_Complement:= Buffer;

               if (EtabAcreer.Adr_Adresse1 <> '') then
                  AdresseNormalisee (EtabAcreer.Adr_Adresse1, numero, NomRue);

               if (numero <> '') then
                  EtabAcreer.Adr_Numero:= numero;
               if (NomRue <> '') then
                  EtabAcreer.Adr_NomRue:= NomRue;

               EtabAcreer.Adr_CP:= QRechEtab.FindField('ET_CODEPOSTAL').AsString;
               EtabAcreer.Adr_Ville:= PGUpperCase(QRechEtab.FindField('ET_VILLE').AsString);

               if ((QRechEtab.FindField('ET_PAYS').AsString <> 'FRA') and
                  (QRechEtab.FindField('ET_PAYS').AsString <> '')) then
                  begin
                  PaysISOLib (QRechEtab.FindField('ET_PAYS').AsString, CodeIso,
                              Libelle);
                  EtabAcreer.Adr_PaysCode:= CodeIso;
                  EtabAcreer.Adr_PaysNom:= Libelle;
                  end;

               StDate:= UsDateTime(StrToDate('31/12/'+PGExercice));
               StNbSal:= 'SELECT COUNT(PSA_SALARIE) AS NBSAL'+
                         ' FROM SALARIES WHERE'+
                         ' (PSA_DATESORTIE="'+UsDateTime(IDate1900)+'" OR'+
                         ' PSA_DATESORTIE IS NULL OR'+
                         ' PSA_DATESORTIE>="'+StDate+'") AND'+
                         ' PSA_DATEENTREE<="'+StDate+'" AND'+
                         ' PSA_DATEENTREE IS NOT NULL AND'+
                         ' PSA_ETABLISSEMENT="'+QRechEtab.FindField('ET_ETABLISSEMENT').AsString+'"';
               QNbSal:= OpenSql (StNbSal,True);

               if (not QNbSal.EOF) then
                  NbSal:= QNbSal.FindField('NBSAL').asinteger
               else
                  NbSal:= 0;
               Ferme (QNbSal);

               EtabAcreer.Effectif:= IntToStr(NbSal);

               If (NbSal=0) then
                  begin
                  StTestSal:= 'SELECT COUNT(DISTINCT(PDS_SALARIE)) AS NBPAIE'+
                              ' FROM DADSDETAIL WHERE'+
                              ' PDS_SEGMENT="S41.G01.00.005" AND'+
                              ' PDS_DONNEEAFFICH="'+QRechEtab.FindField('ET_ETABLISSEMENT').AsString+'" AND'+
                              ' PDS_DATEFIN <= "'+UsDateTime(FinExer)+'" AND'+
                              ' PDS_DATEDEBUT >= "'+UsDateTime(DebExer)+'" AND'+
                              ' PDS_EXERCICEDADS = "'+PGExercice+'"';
                  QTestSal:= OpenSql(StTestSal,True);

                  if (QTestSal.FindField('NBPAIE').AsInteger=0) then
                     EtabAcreer.SansSal:= '01';
                  Ferme(QTestSal);
                  end;

               TaxeSal:= QRechEtab.FindField('ETB_PRORATATVA').AsInteger;
               if (TaxeSal=0) then
                  EtabAcreer.TaxeSal:= '02'
               else
                  EtabAcreer.TaxeSal:= '01';

               CodeAPE:= QRechEtab.FindField('ET_APE').AsString;
               if (RechDom('YYCODENAF', CodeAPE, TRUE)='') then
                  begin
                  ErreurDADSU.Segment:= 'S80.G01.00.006';
                  ErreurDADSU.Explication:= 'L''établissement '+
                                            QRechEtab.FindField('ET_ETABLISSEMENT').AsString+
                                            ' a pour code APE "'+CodeAPE+'"'+
                                            ' qui est inconnu';
                  ErreurDADSU.CtrlBloquant:= True;
                  EcrireErreur (ErreurDADSU);
                  end;
               EtabAcreer.NAF:= CodeAPE;

               if (GetControlText('CBPRUDH')='X') then
{PT7-1
                  EtabAcreer.Prudh:= QRechEtab.FindField ('ETB_PRUDH').AsString;
}
                  begin
                  EtabAcreer.PrudhSect:= QRechEtab.FindField ('ETB_PRUDH').AsString;
//PT7-2
                  if ((((CodeAPE='725Z') or (CodeAPE='748B') or
                     (CodeAPE='921G') or (CodeAPE='924Z') or
                     ((Copy (CodeAPE, 1, 2)='05') and (CodeAPE<>'050C')) or
                     ((Copy (CodeAPE, 1, 2)>='10') and
                     (Copy (CodeAPE, 1, 2)<='36') and (CodeAPE<>'151F')) or
                     (Copy (CodeAPE, 1, 2)='40') or
                     (Copy (CodeAPE, 1, 2)='41') or
                     (Copy (CodeAPE, 1, 2)='45')) and
                     (EtabAcreer.PrudhSect<>'01')) or (((CodeAPE='151F') or
                     (CodeAPE='741J') or (CodeAPE='747Z') or (CodeAPE='748A') or
                     (CodeAPE='748G') or (CodeAPE='748H') or (CodeAPE='851H') or
                     (CodeAPE='922F') or (Copy (CodeAPE, 1, 2)='37') or
                     ((Copy (CodeAPE, 1, 2)>='50') and
                     (Copy (CodeAPE, 1, 2)<='52')) or
                     (Copy (CodeAPE, 1, 2)='55') or
                     ((Copy (CodeAPE, 1, 2)>='60') and
                     (Copy (CodeAPE, 1, 2)<='71') and (CodeAPE<>'602C') and
                     (CodeAPE<>'660G') and (CodeAPE<>'701C')) or
                     (Copy (CodeAPE, 1, 2)='90') or
                     ((Copy (CodeAPE, 1, 2)='93') and (CodeAPE<>'930K'))) and
                     (EtabAcreer.PrudhSect<>'02')) or (((CodeAPE='050C') or
                     (Copy (CodeAPE, 1, 2)='01') or
                     (Copy (CodeAPE, 1, 2)='02')) and
                     (EtabAcreer.PrudhSect<>'03')) or (((CodeAPE='602C') or
                     (CodeAPE='660G') or (CodeAPE='701C') or (CodeAPE='930K') or
                     ((Copy (CodeAPE, 1, 2)>='72') and
                     (Copy (CodeAPE, 1, 2)<='75') and (CodeAPE<>'725Z') and
                     (CodeAPE<>'741J') and (CodeAPE<>'747Z') and
                     (CodeAPE<>'748A') and (CodeAPE<>'748B') and
                     (CodeAPE<>'748G') and (CodeAPE<>'748H')) or
                     (Copy (CodeAPE, 1, 2)='80') or
                     ((Copy (CodeAPE, 1, 2)='85') and (CodeAPE<>'851H')) or
                     (Copy (CodeAPE, 1, 2)='91') or
                     ((Copy (CodeAPE, 1, 2)='92') and (CodeAPE<>'921G') and
                     (CodeAPE<>'922F') and (CodeAPE<>'924Z')) or
                     ((Copy (CodeAPE, 1, 2)>='95') and
                     (Copy (CodeAPE, 1, 2)<='97')) or
                     (Copy (CodeAPE, 1, 2)='99')) and
                     (EtabAcreer.PrudhSect<>'04'))) then
                     begin
                     ErreurDADSU.Segment:= 'S80.G01.00.007.001';
                     ErreurDADSU.Explication:= 'La section prud''hommale'+
                                               ' établissement n''est pas'+
                                               ' compatible avec le code APE';
                     ErreurDADSU.CtrlBloquant:= True;
                     EcrireErreur (ErreurDADSU);
                     end;
//FIN PT7-2
                  EtabAcreer.PrudhDero:= '';
                  end;
//FIN PT7-1
               if (ErrorDADSU=0) then
                  CalculS80_Etablissement(EtabAcreer);
               end;
            Ferme(QRechEtab);

            Etab := TEtabD.GetValue('CODE');
            Num := Num+1;
            end;
         TEtabD := TEtabSelect.FindNext([''],[''],FALSE);
         end;
   end;
FreeAndNil (TEtabSelect);
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 02/10/2006
Modifié le ... :   /  /
Description .. : Sortie de la zone de saisie du code salarié
Mots clefs ... : PAIE;PGDADSU
*****************************************************************}
procedure TOF_PG_EDITDADSCONT.SalarieExit(Sender: TObject);
begin
if (isnumeric(SALAR.Text) AND (VH_PAIE.PgTypeNumSal='NUM')) then
   SALAR.Text:=ColleZeroDevant(StrToInt(SALAR.Text),10);
end;


Initialization
registerclasses ( [ TOF_PG_EDITDADSCONT ] ) ;
end.
