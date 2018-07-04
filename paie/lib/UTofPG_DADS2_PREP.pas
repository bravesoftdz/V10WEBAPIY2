{***********UNITE*************************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 17/06/2003
Modifié le ... :
Description .. : Unité de gestion de la fiche de gestion de paramétrage
Suite ........ : entête entreprise
Mots clefs ... : PAIE;PGDADSB
*****************************************************************}
{
PT1   : 14/01/2004 VG V_50 Adaptation cahier des charges 2003
PT2-1 : 31/03/2004 VG V_50 La totalisation entreprise était fausse
PT2-2 : 31/03/2004 VG V_50 Le nom de naissance était mal alimenté
PT2-3 : 31/03/2004 VG V_50 La rémunération nette ne prenait pas en compte les
                           revenus d'activité
PT2-4 : 31/03/2004 VG V_50 Les avantages en nature et frais professionnels
                           étaient mal cadrés
PT3   : 14/05/2004 VG V_50 Ajout du traitement lié aux raccourcis clavier
PT4   : 19/05/2004 VG V_50 On complète le message lorsque le siret n'est pas
                           valide - FQ N°11152
PT5   : 08/10/2004 VG V_50 Adaptation cahier des charges 2004
PT6   : 11/03/2005 VG V_60 Gestion des champs de l'enregistrement "Honoraires"
                           qui pourraient être mal initialisés - FQ N°12090
PT7   : 05/01/2006 VG V_65 Adaptation cahier des charges 2005
PT8   : 02/03/2006 VG V_65 Prise en compte des modifications de structure
PT9   : 13/10/2006 VG V_70 Suppression du fichier de contrôle - mise en table
                           des erreurs
PT10  : 14/11/2006 VG V_70 Adaptation cahier des charges 2006
PT11  : 13/02/2007 VG V_70 Correction alimentation 2ème période en numérique
                           FQ N°13922
PT12  : 20/07/2007 VG V_70 Fichiers en base
PT13  : 14/09/2007 MF V_80 On passe le n° de dossier en paramètre à la fonction AGL__YFILESTD_IMPORT
PT14  : 12/11/2007 NA V_80 Modif réglementaire 2007
PT15  : 18/01/2008 NA V_81 FQ 15126 Assujetti taxe apprentissage et formation professionnelle (stocké dans DADSDETAIL)
}
unit UTofPG_DADS2_PREP;

interface

uses
         {$IFDEF VER150}
         Variants,
         {$ENDIF}
         UTOF,
         sysutils,
         hctrls,
         Vierge,
         ParamSoc,
         UTOB,
{$IFDEF COMPTA}
         UtilPGI,         //VerifSiret
         UtilTrans,       //EpureChar
{$ELSE}
         PgOutils2,
{$ENDIF}
         HMsgBox,
         PgDADSCommun,
         HEnt1,
         ed_tools,
         Classes,
         controls,
         hstatus,
{$IFNDEF EAGLCLIENT}
         FE_Main,
         windows,
         {$IFNDEF DBXPRESS} dbTables {$ELSE} uDbxDataSet {$ENDIF},
{$ELSE}
         MaineAGL,
         entpaie,
         windows,
{$ENDIF}
         StdCtrls,
         uYFILESTD;

Type TotalSalEtab = record
     TotalS100 : double;
     TotalS102103 : double;
     TotalS105 : double;
     TotalS117 : double;
     TotalS125 : double;
     TotalS132 : double;
     TotalS134 : double;
     TotalS135 : double;
     TotalS136 : double;
     TotalS146 : double; // pt14
     TotalTS : double;
     TotalEff : double;
     TotalBaseApp : double; // pt15
     TotalBaseFPC : double; // pt15
     TotalBaseFPCCDD : double; // pt15
end;

Type TotalHonEtab = record
     TotalH038 : double;
     TotalH039 : double;
     TotalH040 : double;
     TotalH041 : double;
     TotalH042 : double;
     TotalH043 : double;
     TotalH044 : double;
     TotalH045 : double;
     TotalH046 : double;
     TotalH047 : double;
     TotalH048 : double;
end;

Type TotalEntreprise = record
     TotEtab : integer;
     TotSal : integer;
     TotHon : integer;
     TotalEtabS : TotalSalEtab;
     TotalEtabH : TotalHonEtab;
end;

// deb pt15
type TTAXE = record
     TCodeEtab : string;    //Etablissement
     TASSAPP : string;      //Assujetti taxe apprentissage
     TTotalTA : double;     //Total base TA
     TASSFORMPROF : string; //Assujetti formation prof
     TBASEFPCCDD : double; //Total Base FPC CDD
     TBASEFPC : double;   //Total base  FPC
end;
// fin pt15

Type
      TOF_PG_DADS2_PREP = Class (TOF)
      public
        procedure OnArgument (stArgument : String ) ; override ;

      private
        TEtabR, TEtabResult : TOB;
        ErrorDADSU : integer;
        Nature, Validite : THValComboBox;
        FDADS : TextFile;

        procedure ChargeZones;
        procedure Generation (Sender: TObject);
        function CreeEtab : boolean;
        procedure GenereFichier;
        procedure EcritureFicEntre;
        procedure EcritureFicEtab(TEtabD : TOB; var Taxe : TTAXE); // pt15
        procedure EcritureFicSal(TSalD : TOB; var TotalEtabS : TotalSalEtab);
        procedure EcritureFicHon(THonD : TOB; var TotalEtabH : TotalHonEtab);
        procedure EcritureFicTotEtab(Siret, Section, DValidite, TypeDADS,
                                     CodeEtab : string;
                                     Taxe : TTAXE;
                                     var TotalEtabS : TotalSalEtab;
                                     TotalEtabH : TotalHonEtab);     // pt15
        procedure EcritureFicTotDecla(TotalEntre : TotalEntreprise);
        procedure InitTotalEtab(var TotalEtabS : TotalSalEtab; var TotalEtabH : TotalHonEtab);
        procedure Parametrage(Sender: TObject);
        procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
     END ;

     procedure LanceFiche_DADS2Prep;

implementation
uses Pg_OutilsEnvoi;

procedure LanceFiche_DADS2Prep;
begin
  AGLLanceFiche('PAY', 'DADS2_PREP', '', '', '');
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 17/06/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... : PAIE;PGDADSB
*****************************************************************}
procedure TOF_PG_DADS2_PREP.OnArgument(stArgument: String);
var
Mes, StPcl : string;
begin
Inherited ;
ErrorDADSU := 0;                

ChargeZones;

TFVierge(Ecran).BValider.OnClick := Generation;
TFVierge(Ecran).OnKeyDown:=FormKeyDown;	//PT3
//PT12
{Recherche fichier <> CEG  et = DA2 pour savoir si client utilise fichier en
base pour le traitement des TD bilatéral}
if (V_PGI.ModePcl = '1') then
   StPcl:= ' AND NOT (YFS_PREDEFINI="DOS" AND'+
           ' YFS_NODOSSIER<>"'+V_PGI.NoDossier+'")'
else
   StPcl:= '';

If not existeSQL ('SELECT YFS_NOM'+
                  ' FROM YFILESTD WHERE'+
                  ' YFS_CODEPRODUIT="PAIE" AND'+
                  ' YFS_CRIT1="DA2" AND'+
                  ' YFS_PREDEFINI<>"CEG"'+StPcl) then
   begin
   Mes:= 'Cette version apporte un nouveau fonctionnement dans#13#10'+
         'la gestion du stockage des fichiers.#13#10#13#10'+
         'Par défaut l''utilisation reste identique.#13#10#13#10'+
         'Toutefois nous vous invitons à mettre en place cette nouvelle'+
         ' fonctionnalité';
   SetControlChecked ('SANSFICHIERBASE', true);
   HShowMessage ('1;Gestion des fichiers en base ;'+Mes+';E;HO;;;;;41910025', '', '');
   end;
//FIN PT12
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 17/06/2003
Modifié le ... :   /  /
Description .. : Chargement des éléments de la fiche
Mots clefs ... : PAIE,PGDADSB
*****************************************************************}
procedure TOF_PG_DADS2_PREP.ChargeZones;
var
Buffer, BufDest, BufOrig, NomRue, numero : string;
Jour : TDateTime;
AnneeA, JourJ, MoisM : Word;
AnneePrec : string;
begin
{ FQ 20736 BVE 18.06.07 }
{$IFDEF COMPTA}           
{ FQ 20731 BVE 18.06.07 }
SetControlProperty('CSECTION','Plus',' AND CO_CODE = "01"');
{ END FQ 20731 }
{$ENDIF}
SetControlText('CSECTION', '01');
{ END FQ 20736 }

Validite := THValComboBox (GetControl ('CEXERCICE'));
Nature := THValComboBox (GetControl ('CTYPE'));

Jour := Date;
DecodeDate(Jour, AnneeA, MoisM, JourJ);
if MoisM>9 then
   AnneePrec := IntToStr(AnneeA)
else
   AnneePrec := IntToStr(AnneeA-1);

if Validite <> NIL then
   begin
   Validite.value:=copy(AnneePrec, 1, 1)+copy(AnneePrec, 3, 2);
   PGAnnee := Validite.value;
   PGExercice := AnneePrec;
   Validite.OnChange := Parametrage;
   end;

if Nature <> nil then
   begin
{ FQ 20736 BVE 18.06.07 }
{$IFDEF COMPTA}
   Nature.Plus := ' AND (CO_CODE = "1" OR CO_CODE = "3" OR CO_CODE = "4")';
   Nature.Value := '1';
{$ELSE}
   Nature.Plus := '';
   Nature.Value := '0';
{$ENDIF}
{ END FQ 20736 }
   Nature.OnChange :=  Parametrage;
   end;

SetControlText('ERAISONSOC', GetParamSoc('SO_LIBELLE'));

BufOrig := GetParamSoc('SO_SIRET');
{$IFDEF COMPTA}
  EpureChar(BufOrig, BufDest);
  if not VerifSiret(BufDest) then
{$ELSE}
ForceNumerique(BufOrig, BufDest);
if ControlSiret(BufDest)=False then
{$ENDIF}
//PT4
   PGIBox('Le SIRET n''est pas valide.#13#10'+
          'Vous devez le vérifier en y accédant par le module#13#10'+
          'Paramètres/menu Comptabilité/#13#10'+
          'commande Paramètres comptables/Coordonnées.#13#10'+
          'Si vous travaillez en environnement multi-dossiers, vous#13#10'+
          'pouvez y accéder par le Bureau PGI/Annuaire',
          'Saisie Entreprise DADS-U');
//FIN PT4
SetControlText('ESIREN', Copy(BufDest, 1, 9));

if (GetParamSoc('SO_ADRESSE2') <> '') then
   begin
   if (GetParamSoc('SO_ADRESSE3') <> '') then
      Buffer := GetParamSoc('SO_ADRESSE2')+' '+GetParamSoc('SO_ADRESSE3')
   else
      Buffer := GetParamSoc('SO_ADRESSE2');
   end
else
   if (GetParamSoc('SO_ADRESSE3') <> '') then
      Buffer := GetParamSoc('SO_ADRESSE3');
if (Buffer <> '') then
   SetControlText('EADRCOMP', Buffer);

if (GetParamSoc('SO_ADRESSE1') <> '') then
   AdresseNormalisee (GetParamSoc('SO_ADRESSE1'), numero, NomRue);

if (numero <> '') then
   SetControlText('EADRNUM', numero);

if (NomRue <> '') then
   SetControlText('EADRNOMVOIE', NomRue);

SetControlText('EADRCP', GetParamSoc('SO_CODEPOSTAL'));
SetControlText('EADRBURDISTRIB', GetParamSoc('SO_VILLE'));
SetControlText('EAPE', GetParamSoc('SO_APE'));
SetControlText('CDEPOSE', GetParamSoc('SO_ETABLISDEFAUT'));
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 17/06/2003
Modifié le ... :   /  /
Description .. : Validation
Mots clefs ... : PAIE;PGDADSB
*****************************************************************}
procedure TOF_PG_DADS2_PREP.Generation (Sender: TObject);
var
EtabOK : boolean;
begin
if ((GetControlText('CSECTION') = '02') and
   (GetControlText('CTYPE') <> '2')) then
   begin
   PGIBox ('La section et le type sont incompatibles',
          'Préparation TD Bilatéral');
   SetControlText('CTYPE', '2');
   exit;
   end;
try
   begintrans;
   if ErrorDADSU=0 then
      begin
      EtabOK := CreeEtab;
      if (EtabOK = True) then
         GenereFichier;
      end;
   CommitTrans;
Except
   Rollback;
   PGIBox ('Une erreur est survenue lors de la mise à jour de la base',
          'Préparation TD Bilatéral');
   END;
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 17/06/2003
Modifié le ... :   /  /
Description .. : Generation du fichier DADSB
Mots clefs ... : PAIE,PGDADSB
*****************************************************************}
procedure TOF_PG_DADS2_PREP.GenereFichier;
var
TEtabD, TGener, TGenerD : TOB;
Buffer, CodeEtab, Guid1, Millesime, Nomcomplet, NomFic, NomFicRapport : string;
Siret, StCompte, StDelete, StExer, StGener, StPCL, StQQ : string;
CodeRetour, FileAttrs, Nbre, Ordre : integer;
QCompte, QQ, QRechDelete, QRechExer, QRechGener : TQuery;
TotalEtabS : TotalSalEtab;
TotalEtabH : TotalHonEtab;
TotalEntre : TotalEntreprise;
sr : TSearchRec;
Size : double;
EnregEnvoi : TEnvoiSocial;
NoDossier : string; // PT13
Taxe : TTaxe; // pt15
begin
TotalEntre.TotEtab:=0;
TotalEntre.TotSal:=0;
TotalEntre.TotHon:=0;
TotalEntre.TotalEtabS.TotalS100:=0;
TotalEntre.TotalEtabS.TotalS102103:=0;
TotalEntre.TotalEtabS.TotalS105:=0;
TotalEntre.TotalEtabS.TotalS117:=0;
TotalEntre.TotalEtabS.TotalS125:=0;
TotalEntre.TotalEtabS.TotalS132:=0;
TotalEntre.TotalEtabS.TotalS134:=0;
TotalEntre.TotalEtabS.TotalS135:=0;
TotalEntre.TotalEtabS.TotalS136:=0;
TotalEntre.TotalEtabS.TotalS146:=0;   // pt14
TotalEntre.TotalEtabS.TotalTS:=0;
TotalEntre.TotalEtabS.TotalEff:=0;
TotalEntre.TotalEtabS.TotalBaseApp := 0; // pt15
TotalEntre.TotalEtabS.TotalBaseFPC := 0; // pt15
TotalEntre.TotalEtabS.TotalBaseFPCCDD := 0; // pt15
TotalEntre.TotalEtabH.TotalH038:=0;
TotalEntre.TotalEtabH.TotalH039:=0;
TotalEntre.TotalEtabH.TotalH040:=0;
TotalEntre.TotalEtabH.TotalH041:=0;
TotalEntre.TotalEtabH.TotalH042:=0;
TotalEntre.TotalEtabH.TotalH043:=0;
TotalEntre.TotalEtabH.TotalH044:=0;
TotalEntre.TotalEtabH.TotalH045:=0;
TotalEntre.TotalEtabH.TotalH046:=0;
TotalEntre.TotalEtabH.TotalH047:=0;
TotalEntre.TotalEtabH.TotalH048:=0;

{$IFDEF COMPTA}
  EpureChar(GetParamSoc('SO_SIRET'), Buffer);
{$ELSE}
ForceNumerique(GetParamSoc('SO_SIRET'), Buffer);
{$ENDIF}
{$IFDEF EAGLCLIENT}
NomFicRapport := VH_Paie.PgCheminEagl+'\'+Buffer+'_TDB_PGI_PREP.log';
{$ELSE}
NomFicRapport := V_PGI.DatPath+'\'+Buffer+'_TDB_PGI_PREP.log';
{$ENDIF}

{PT9
AssignFile(FRapport, NomFicRapport);
ReWrite(FRapport);
Writeln(FRapport, '');
Writeln(FRapport, 'Début de préparation TD Bilatéral : '+DateTimeToStr(Now));
}

Size:=0;
StCompte := 'SELECT count(*) AS NBENREG'+
            ' FROM DADS2SALARIES WHERE'+
            ' PD2_VALIDITE="'+PGExercice+'" AND'+
            ' PD2_SECTIONETAB="'+GetControlText('CSECTION')+'" AND'+
            ' PD2_TYPEDADS="'+Nature.Value+'"';
QCompte:=OpenSql(StCompte,TRUE);
if (Not QCompte.EOF) then
   Nbre := QCompte.FindField('NBENREG').Asinteger
else
   Nbre := 0;
Ferme(QCompte);

StCompte := 'SELECT count(*) AS NBENREG'+
            ' FROM DADS2HONORAIRES WHERE'+
            ' PDH_VALIDITE="'+PGExercice+'" AND'+
            ' PDH_SECTIONETAB="'+GetControlText('CSECTION')+'" AND'+
            ' PDH_TYPEDADS="'+Nature.Value+'"';
QCompte:=OpenSql(StCompte,TRUE);
if (Not QCompte.EOF) then
   Nbre := Nbre+QCompte.FindField('NBENREG').Asinteger;
Ferme(QCompte);
InitMoveProgressForm (NIL,'Traitement en cours',
                     'Veuillez patienter SVP ...',Nbre,FALSE,TRUE);
InitMove(Nbre,'');

//Ouverture du fichier
{PT12
NomFic := 'D2'+Nature.Value+GetControlText('CSECTION')+GetParamSoc('SO_SIRET')+
          PGExercice+'.DA2';
}
NomFic:= 'D2'+Nature.Value+GetControlText('CSECTION')+Buffer+PGExercice+'.DA2';
//FIN PT12
{$IFDEF EAGLCLIENT}
Nomcomplet := VH_Paie.PgCheminEagl+'\'+NomFic;
{$ELSE}
Nomcomplet := V_PGI.DatPath+'\'+NomFic;
{$ENDIF}
if FileExists(Nomcomplet) then
      DeleteFile(PChar(Nomcomplet));

try
   begintrans;
   AssignFile(FDADS, Nomcomplet);
{$i-} ReWrite(FDADS); {$i+}
   If IoResult<>0 Then
      Begin
      PGIBox ('Fichier inaccessible : '+Nomcomplet,
             'Préparation TD Bilatéral');
      Exit ;
      End;

//Données ENTREPRISE
   EcritureFicEntre;

//Données ETABLISSEMENT
   TEtabD := TEtabR.FindFirst([''],[''],FALSE);
   While (TEtabD<>nil) do
         begin
         Siret:= TEtabD.getValue('SIRET');
         CodeEtab := TEtabD.getValue('CODEETAB');
         EcritureFicEtab(TEtabD, Taxe);    // pt15
         TotalEntre.TotEtab := TotalEntre.TotEtab+1;

//Données SALARIE
         if ((GetControlText('CTYPE') = '0') or
            (GetControlText('CTYPE') = '2') or
            (GetControlText('CTYPE') = '5')) then
            begin
            StGener := 'SELECT *'+
                       ' FROM DADS2SALARIES WHERE'+
                       ' PD2_VALIDITE="'+PGExercice+'" AND'+
                       ' PD2_ETABLISSEMENT="'+CodeEtab+'" AND'+
                       ' PD2_SECTIONETAB="'+GetControlText('CSECTION')+'"';
            QRechGener:=OpenSql(StGener,TRUE);
            TGener := TOB.Create('Les éléments salariés', NIL, -1);
            TGener.LoadDetailDB('DADS2SALARIES','','',QRechGener,False);
            Ferme(QRechGener);
            TGenerD := TGener.FindFirst([''],[''],FALSE);
            While (TGenerD <> nil) do
               begin
               EcritureFicSal(TGenerD, TotalEtabS);
               TotalEntre.TotSal := TotalEntre.TotSal+1;
               TGenerD := TGener.FindNext([''],[''],FALSE);
               end;
            TotalEntre.TotalEtabS.TotalS100 := TotalEntre.TotalEtabS.TotalS100+
                                               TotalEtabS.TotalS100;
            TotalEntre.TotalEtabS.TotalS102103 := TotalEntre.TotalEtabS.TotalS102103+
                                                  TotalEtabS.TotalS102103;
            TotalEntre.TotalEtabS.TotalS105 := TotalEntre.TotalEtabS.TotalS105+
                                               TotalEtabS.TotalS105;
            TotalEntre.TotalEtabS.TotalS117 := TotalEntre.TotalEtabS.TotalS117+
                                               TotalEtabS.TotalS117;
            TotalEntre.TotalEtabS.TotalS125 := TotalEntre.TotalEtabS.TotalS125+
                                               TotalEtabS.TotalS125;
            TotalEntre.TotalEtabS.TotalS132 := TotalEntre.TotalEtabS.TotalS132+
                                               TotalEtabS.TotalS132;
            TotalEntre.TotalEtabS.TotalS134 := TotalEntre.TotalEtabS.TotalS134+
                                               TotalEtabS.TotalS134;
            TotalEntre.TotalEtabS.TotalS135 := TotalEntre.TotalEtabS.TotalS135+
                                               TotalEtabS.TotalS135;
            TotalEntre.TotalEtabS.TotalS136 := TotalEntre.TotalEtabS.TotalS136+
                                               TotalEtabS.TotalS136;
            TotalEntre.TotalEtabS.TotalS146 := TotalEntre.TotalEtabS.TotalS146+
                                               TotalEtabS.TotalS146;        // pt14
            FreeAndNil(TGener);
            end;

//Données Honoraires
         if ((GetControlText('CTYPE') = '0') or
            (GetControlText('CTYPE') = '1') or
            (GetControlText('CTYPE') = '4')) then
            begin
            StGener := 'SELECT *'+
                       ' FROM DADS2HONORAIRES WHERE'+
                       ' PDH_VALIDITE="'+PGExercice+'" AND'+
                       ' PDH_ETABLISSEMENT="'+CodeEtab+'" AND'+
                       ' PDH_SECTIONETAB="'+GetControlText('CSECTION')+'"';
            QRechGener:=OpenSql(StGener,TRUE);
            TGener := TOB.Create('Les éléments Honoraires', NIL, -1);
            TGener.LoadDetailDB('DADS2HONORAIRES','','',QRechGener,False);
            Ferme(QRechGener);
            TGenerD := TGener.FindFirst([''],[''],FALSE);
            While (TGenerD <> nil) do
               begin
               EcritureFicHon(TGenerD, TotalEtabH);
               TotalEntre.TotHon := TotalEntre.TotHon+1;
               TGenerD := TGener.FindNext([''],[''],FALSE);
               end;
            TotalEntre.TotalEtabH.TotalH038 := TotalEntre.TotalEtabH.TotalH038+
                                               TotalEtabH.TotalH038;
            TotalEntre.TotalEtabH.TotalH039 := TotalEntre.TotalEtabH.TotalH039+
                                               TotalEtabH.TotalH039;
            TotalEntre.TotalEtabH.TotalH040 := TotalEntre.TotalEtabH.TotalH040+
                                               TotalEtabH.TotalH040;
            TotalEntre.TotalEtabH.TotalH041 := TotalEntre.TotalEtabH.TotalH041+
                                               TotalEtabH.TotalH041;
            TotalEntre.TotalEtabH.TotalH042 := TotalEntre.TotalEtabH.TotalH042+
                                               TotalEtabH.TotalH042;
            TotalEntre.TotalEtabH.TotalH043 := TotalEntre.TotalEtabH.TotalH043+
                                               TotalEtabH.TotalH043;
            TotalEntre.TotalEtabH.TotalH044 := TotalEntre.TotalEtabH.TotalH044+
                                               TotalEtabH.TotalH044;
            TotalEntre.TotalEtabH.TotalH045 := TotalEntre.TotalEtabH.TotalH045+
                                               TotalEtabH.TotalH045;
            TotalEntre.TotalEtabH.TotalH046 := TotalEntre.TotalEtabH.TotalH046+
                                               TotalEtabH.TotalH046;
            TotalEntre.TotalEtabH.TotalH047 := TotalEntre.TotalEtabH.TotalH047+
                                               TotalEtabH.TotalH047;
            TotalEntre.TotalEtabH.TotalH048 := TotalEntre.TotalEtabH.TotalH048+
                                               TotalEtabH.TotalH048;
            FreeAndNil(TGener);
            end;

//Total Etablissement
         if ((GetControlText('CTYPE') <> '3') and
            (GetControlText('CTYPE') <> '6')) then
          //  EcritureFicTotEtab(Siret, GetControlText('CSECTION'), PGExercice,
          //                    Nature.Value, CodeEtab, TotalEtabS, TotalEtabH); pt15
         EcritureFicTotEtab(Siret, GetControlText('CSECTION'), PGExercice,
                               Nature.Value, CodeEtab, taxe, TotalEtabS, TotalEtabH); // pt15
         TotalEntre.TotalEtabS.TotalTS := TotalEntre.TotalEtabS.TotalTS+
                                          TotalEtabS.TotalTS;
         TotalEntre.TotalEtabS.TotalEff := TotalEntre.TotalEtabS.TotalEff+
                                           TotalEtabS.TotalEff;

         TotalEntre.TotalEtabS.TotalBaseApp := TotalEntre.TotalEtabS.TotalBaseApp+
                                           TotalEtabS.TotalBaseAPP;                  // pt15
         TotalEntre.TotalEtabS.TotalBaseFPC := TotalEntre.TotalEtabS.TotalBaseFPC+
                                           TotalEtabS.TotalBaseFPC;                  // pt15
         TotalEntre.TotalEtabS.TotalBaseFPCCDD := TotalEntre.TotalEtabS.TotalBaseFPCCDD+
                                           TotalEtabS.TotalBaseFPCCDD;                  // pt15
                                           
         TEtabD := TEtabR.FindNext([''],[''],FALSE);
         InitTotalEtab(TotalEtabS, TotalEtabH);
         end;
   FreeAndNil(TEtabR);

//Données DECLARATION
   EcritureFicTotDecla(TotalEntre);

   CloseFile(FDADS);
{$IFDEF COMPTA}
   NoDossier := V_Pgi.NoDossier;
{$ELSE}
   NoDossier := PgRendNoDossier(); // PT13
{$ENDIF}

//PT12
   if (GetCheckBoxState ('SANSFICHIERBASE')=cbUnchecked) then
// mise en base du fichier TD bilatéral confectionné
      begin
      CodeRetour:= AGL_YFILESTD_IMPORT (NomComplet, 'PAIE', NomFic, 'DA2', 'DA2',
                                        RechDom ('PGANNEE', GetControlText ('CEXERCICE'), False),
                                        Nature.Value, GetControlText('CSECTION'),
                                        '', '-', '-', '-', '-', '-', 'FRA',
                                        'DOS', NomFic, NoDossier);   // PT13
      if (CodeRetour<>-1) then
         PGIInfo (AGL_YFILESTD_GET_ERR (CodeRetour)+'#13#10'+NomComplet);
      end;
//FIN PT12

   FileAttrs := 0;
   FileAttrs := FileAttrs + faAnyFile;
   if FindFirst(Nomcomplet, FileAttrs, sr) = 0 then
      begin
      if (sr.Attr and FileAttrs) = sr.Attr then
         Size := Arrondi(sr.Size/1024, 2);
      sysutils.FindClose(sr);
      end;

   if (ErrorDADSU = 0) then
      begin
//Mise à jour de la table ENVOISOCIAL
      DebExer:= StrToDate ('01/01/'+PGExercice);
      FinExer:= StrToDate ('31/12/'+PGExercice);

      Millesime:= PGExercice;
      StExer:= 'SELECT CO_CODE'+
               ' FROM COMMUN WHERE'+
               ' CO_TYPE="PGA" AND'+
               ' CO_LIBELLE="'+Millesime+'"';
      QRechExer:=OpenSQL(StExer,TRUE) ;
      if Not QRechExer.EOF then
         Millesime := QRechExer.FindField ('CO_CODE').AsString;
      Ferme (QRechExer);

      StDelete := 'DELETE FROM ENVOISOCIAL WHERE'+
                  ' PES_TYPEMESSAGE= "DA2" AND '+
                  ' PES_MILLESSOC = "'+Millesime+'" AND'+
                  ' PES_DATEDEBUT = "'+UsDateTime(DebExer)+'" AND'+
                  ' PES_DATEFIN = "'+UsDateTime(FinExer)+'" AND'+
                  ' PES_SIRETDO = "'+GetControlText('ESIREN')+'" AND'+
                  ' PES_FICHIERRECU = "'+NomFic+'"';
      ExecuteSQL(StDelete) ;

      Ordre:= 1;
      StDelete := 'SELECT MAX(PES_CHRONOMESS) AS MAXI FROM ENVOISOCIAL';
      QRechDelete:=OpenSQL(StDelete,TRUE) ;
      if Not QRechDelete.EOF then
         try
         Ordre:= QRechDelete.FindField ('MAXI').AsInteger+1;
         except
               on E: EConvertError do
                  Ordre:= 1;
         end;
      Ferme(QRechDelete);

      ChargeTOBENVOI;

//PT12
// Récupération du GUID pour maj ENVOISOCIAL
      Guid1:= '';
      if (GetCheckBoxState ('SANSFICHIERBASE')=cbUnchecked) then
         begin

         if (V_PGI.ModePcl='1') then
            StPcl:= ' AND (YFS_PREDEFINI="DOS" AND'+
                    ' YFS_NODOSSIER="'+V_PGI.NoDossier+'")'
         else
           StPcl:= '';
//Recherche enregistrement dans YFILESTD
         StQQ:= 'SELECT YFS_FILEGUID'+
                ' FROM YFILESTD WHERE'+
                ' YFS_NOM="'+NomFic+'" AND'+
                ' YFS_CODEPRODUIT="PAIE" AND'+
                ' YFS_CRIT1="DA2" AND'+
                ' YFS_CRIT2="'+RechDom('PGANNEE', GetControlText ('CEXERCICE'), False)+'" AND'+
                ' YFS_CRIT3="'+Nature.Value+'" AND'+
                ' YFS_CRIT4="'+GetControlText('CSECTION')+'" AND'+
                ' YFS_PREDEFINI<>"CEG"'+StPcl;
         QQ:= OpenSQL (StQQ, TRUE);
         if not QQ.EOF then
            Guid1 := QQ.FindField ('YFS_FILEGUID').AsString;
         Ferme(QQ);
         end;
//FIN PT12

      EnregEnvoi.Ordre := Ordre;
      EnregEnvoi.TypeE := 'DA2';
      EnregEnvoi.Millesime := Millesime;
      EnregEnvoi.Periodicite := 'A';
      EnregEnvoi.DateD := DebExer;
      EnregEnvoi.DateF := FinExer;
      EnregEnvoi.Siret := GetControlText('ESIREN');
      EnregEnvoi.Libelle := GetControlText('ERAISONSOC');
      EnregEnvoi.Size := size;
      EnregEnvoi.NomFic := ExtractFileName(NomFic);
      EnregEnvoi.Statut := '';
      EnregEnvoi.Monnaie := 'EUR';
      EnregEnvoi.Inst := 'ZIMP';
      EnregEnvoi.Guid1:= Guid1;                              //PT12
      CreeEnvoi (EnregEnvoi);
      LibereTOBENVOI;

      FiniMove;
      FiniMoveProgressForm;

//PT12
      if (GetCheckBoxState ('SANSFICHIERBASE')=cbUnchecked) then
// suppression du fichier sur disque
         begin
         if FileExists (NomComplet) then
            DeleteFile(PChar(NomComplet));
         end;
//FIN PT12

{$IFDEF COMPTA}
      PGIBox ('Préparation terminée.#13#10'+
              ' L''étape suivante est la génération des envois qui créera#13#10'+
              ' un fichier pouvant comporter plusieurs entreprises.', 
              'Préparation TD Bilatéral');
{$ELSE}
      PGIBox ('Préparation terminée.#13#10'+
              ' L''étape suivante est la génération des envois qui créera#13#10'+
              ' un fichier pouvant comporter plusieurs entreprises.#13#10'+
              ' Module Paramètres, Menu Gestion envois,#13#10'+
              ' Commande TD Bilatéral - Envois TD Bilatéral',
              'Préparation TD Bilatéral');
{$ENDIF}
      end
   else
      begin
      FiniMove;
      FiniMoveProgressForm;
      DeleteFile(PChar(Nomcomplet));
      PGIAsk ('Préparation annulée car vous avez '+IntToStr(ErrorDADSU)+' erreurs.',
              'Préparation TD Bilatéral');
      end;
{PT9
   Writeln(FRapport, 'Préparation TD Bilatéral terminée : '+DateTimeToStr(Now));
   CloseFile(FRapport);
   ShellExecute( 0, PCHAR('open'),PChar('WordPad'), PChar(NomFicRapport),Nil,SW_RESTORE);
}   
   CommitTrans;
Except
   Rollback;
   CloseFile(FDADS);
   PGIBox ('Une erreur est survenue lors de la mise à jour de la base',
          'Préparation TD Bilatéral');
   END;
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 17/06/2003
Modifié le ... :   /  /
Description .. : Création des enregistrements ETABLISSEMENT
Mots clefs ... : PAIE,PGDADSB
*****************************************************************}
function TOF_PG_DADS2_PREP.CreeEtab : boolean;
var
Buffer, BufDest, BufOrig, CodeAPE, Etab, NomRue, numero, StEtab : string;
StPrefixe : string;
QRechEtab : TQuery;
TEtab, TEtabD: TOB;
reponse, TaxeSal : integer;
begin
result := False;
//Données Salarié-Honoraires-Etab
if ((Nature.Value = '0') or (Nature.Value = '3') or (Nature.Value = '6')) then
   begin
   StPrefixe := 'PD2_';
   StEtab:= 'SELECT DISTINCT PD2_ETABLISSEMENT, PDH_ETABLISSEMENT, PD2_SIRET '+
            ' FROM DADS2SALARIES'+
            ' LEFT JOIN DADS2HONORAIRES ON'+
            ' PD2_VALIDITE=PDH_VALIDITE AND'+
            ' PD2_SECTIONETAB=PDH_SECTIONETAB WHERE'+
            ' PD2_VALIDITE = "'+PGExercice+'" AND'+
            ' PD2_SECTIONETAB = "'+GetControlText('CSECTION')+'"'+
            ' ORDER BY PD2_ETABLISSEMENT';
   end
else
if ((Nature.Value = '2') or (Nature.Value = '5')) then
   begin
   StPrefixe := 'PD2_';
   StEtab:= 'SELECT PD2_SALARIE, PD2_ETABLISSEMENT, PD2_SIRET '+
            ' FROM DADS2SALARIES WHERE'+
            ' PD2_VALIDITE = "'+PGExercice+'"'+
            ' ORDER BY PD2_ETABLISSEMENT';
   end
else
if ((Nature.Value = '1') or (Nature.Value = '4')) then
   begin
   StPrefixe := 'PDH_';
   StEtab:= 'SELECT PDH_HONORAIRE, PDH_ETABLISSEMENT'+
            ' FROM DADS2HONORAIRES WHERE'+
            ' PDH_VALIDITE = "'+PGExercice+'"'+
            ' ORDER BY PDH_ETABLISSEMENT';
   end;

QRechEtab:=OpenSql(StEtab,TRUE);
TEtab := TOB.Create('Les établissements', NIL, -1);
TEtab.LoadDetailDB('DADSDBilatérale','','',QRechEtab,False);
Ferme(QRechEtab);

TEtabD := TEtab.FindFirst([''],[''],FALSE);
if TEtabD <> nil then
   begin
   result := True;
   Etab := '';
   TEtabR:=TOB.Create('ETABDADS', Nil, -1);
   While (TEtabD<>nil) do
         begin
         if (TEtabD.GetValue(StPrefixe+'ETABLISSEMENT') <> Etab) then
            begin
            StEtab := 'SELECT ET_ETABLISSEMENT, ET_SIRET, ET_APE, ET_LIBELLE,'+
                      ' ET_ADRESSE1, ET_ADRESSE2, ET_ADRESSE3, ET_CODEPOSTAL,'+
                      ' ET_VILLE, ET_ACTIVITE, ETB_PRORATATVA'+
                      ' FROM ETABLISS'+
                      ' LEFT JOIN ETABCOMPL ON'+
                      ' ETB_ETABLISSEMENT=ET_ETABLISSEMENT WHERE'+
                      ' ET_ETABLISSEMENT="'+TEtabD.GetValue(StPrefixe+'ETABLISSEMENT')+'"';
            QRechEtab:=OpenSql(StEtab,TRUE);
            if (not QRechEtab.EOF) then
               begin
               BufOrig := QRechEtab.FindField('ET_SIRET').AsString;
{$IFDEF COMPTA}
              EpureChar(BufOrig, BufDest);
              if not VerifSiret(BufDest) then
{$ELSE}
               ForceNumerique(BufOrig, BufDest);
               if ControlSiret(BufDest)=False then
{$ENDIF}
                  begin

                  PGIBox ('Le SIRET '+TEtabD.GetValue('PD2_SIRET')+
                         ' n''est pas valide.', 'Préparation TD Bilatéral');
                  ErrorDADSU := 1;
                  end;
               TEtabResult := TOB.Create ('SiretTraites',TEtabR,-1);
               TEtabResult.AddChampSupValeur ('SIRET',BufDest,FALSE) ;
               TEtabResult.AddChampSupValeur ('SECTION',
                                             GetControlText('CSECTION'),FALSE) ;
               TEtabResult.AddChampSupValeur ('VALIDITE', PGExercice, FALSE) ;
               TEtabResult.AddChampSupValeur ('TYPE',
                                             GetControlText('CTYPE'),FALSE) ;
               CodeAPE := QRechEtab.FindField('ET_APE').AsString;
               if (RechDom('YYCODENAF', CodeAPE, TRUE)='') then
                  begin
                  if (ErrorDADSU = 0) then
                     begin
                     reponse := PGIAsk ('L''établissement "'+QRechEtab.FindField('ET_ETABLISSEMENT').AsString+'"#13#10'+
                                       ' a pour code APE "'+CodeAPE+'" qui est inconnu. #13#10'+
                                       'Voulez-vous continuer ?',
                                       'Préparation TD Bilatéral');
                     if (reponse=mrNo) then
                        ErrorDADSU := 1;
                     end
                  else
                     PGIBox('L''établissement "'+QRechEtab.FindField('ET_ETABLISSEMENT').AsString+'"#13#10'+
                           ' a pour code APE "'+CodeAPE+'" qui est inconnu. #13#10',
                           'Préparation TD Bilatéral');
                  end;
               TEtabResult.AddChampSupValeur ('APE',
                                             QRechEtab.FindField('ET_APE').AsString,FALSE);
               TEtabResult.AddChampSupValeur ('MODIFSIRET',
                                             '00000000000000',FALSE) ;
               TEtabResult.AddChampSupValeur ('LIBELLE',
                                             QRechEtab.FindField('ET_LIBELLE').AsString,FALSE) ;
               if (QRechEtab.FindField('ET_ADRESSE2').AsString <> '') then
                  begin
                  if (QRechEtab.FindField('ET_ADRESSE3').AsString <> '') then
                     Buffer := QRechEtab.FindField('ET_ADRESSE2').AsString+' '+
                               QRechEtab.FindField('ET_ADRESSE3').AsString
                  else
                     Buffer := QRechEtab.FindField('ET_ADRESSE2').AsString;
                  end
               else
                  if (QRechEtab.FindField('ET_ADRESSE3').AsString <> '') then
                     Buffer := QRechEtab.FindField('ET_ADRESSE3').AsString;
               TEtabResult.AddChampSupValeur ('ADRCOMPL',Buffer,FALSE) ;

               if (QRechEtab.FindField('ET_ADRESSE1').AsString <> '') then
                  AdresseNormalisee (QRechEtab.FindField('ET_ADRESSE1').AsString,
                                    numero, NomRue);

               TEtabResult.AddChampSupValeur ('ADRNUM',numero,FALSE) ;

               TEtabResult.AddChampSupValeur ('ADRBIS',' ',FALSE) ;
               TEtabResult.AddChampSupValeur ('ADRNOM',NomRue,FALSE) ;
               TEtabResult.AddChampSupValeur ('ADRINSEE','',FALSE) ;
               TEtabResult.AddChampSupValeur ('ADRCOMMUNE','',FALSE) ;
               if ((QRechEtab.FindField('ET_CODEPOSTAL').AsString <> '') and
                  (QRechEtab.FindField('ET_CODEPOSTAL').AsString > '00000') and
                  (QRechEtab.FindField('ET_CODEPOSTAL').AsString < '99999')) then
                  TEtabResult.AddChampSupValeur ('ADRCP',
                                                QRechEtab.FindField('ET_CODEPOSTAL').AsString,FALSE)
               else
                  begin
                  PGIBox ('L''adresse de l''établissement '+
                         QRechEtab.FindField('ET_ETABLISSEMENT').AsString+
                         ' est mal renseignée', 'Préparation TD Bilatéral');
                  ErrorDADSU := 1;
                  end;
               if QRechEtab.FindField('ET_VILLE').AsString <> '' then
                  TEtabResult.AddChampSupValeur ('ADRDISTRIB',
                                                QRechEtab.FindField('ET_VILLE').AsString,FALSE)
               else
                  begin
                  PGIBox ('L''adresse de l''établissement '+
                         QRechEtab.FindField('ET_ETABLISSEMENT').AsString+
                         ' est mal renseignée',
                         'Préparation TD Bilatéral');
                  ErrorDADSU := 1;
                  end;
               TEtabResult.AddChampSupValeur ('PROFESSION',
                                             QRechEtab.FindField('ET_ACTIVITE').AsString,FALSE) ;
               TaxeSal := QRechEtab.FindField('ETB_PRORATATVA').AsInteger;
               if (TaxeSal=0) then
                  TEtabResult.AddChampSupValeur ('TAXESAL','N',FALSE)
               else
                  TEtabResult.AddChampSupValeur ('TAXESAL','I',FALSE) ;
               TEtabResult.AddChampSupValeur ('CODEETAB',
                                             QRechEtab.FindField('ET_ETABLISSEMENT').AsString,FALSE) ;
               end;

            Ferme(QRechEtab);
            Etab := TEtabD.GetValue(StPrefixe+'ETABLISSEMENT');
            end;
         TEtabD := TEtab.FindNext([''],[''],FALSE);
         end;
   end
else
   PGIBox ('Données non trouvées. Vérifiez vos paramètres',
           'Préparation TD Bilatéral');
FreeandNil(TEtab);
if (ErrorDADSU <> 0) then
   result := False;
//LibereTOBDADS;
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 06/08/2003
Modifié le ... :   /  /
Description .. : Fonction permettant d'écrire les données concernant
Suite ........ : l'entreprise
Mots clefs ... : PAIE,PGDADSB
*****************************************************************}
procedure TOF_PG_DADS2_PREP.EcritureFicEntre;
var
Buf, Buffer, BufOrig, BufDest, Numero, NomRue, NumEtab, StEtab : string;
QRechEtab : TQuery;
begin

Buffer := FORMAT_STRING(GetControlText('ESIREN'), 9);                               //003
Buffer := Buffer+StringOfChar('0', 12);                                             //004
Buffer := Buffer+'010';                                                             //009
Buffer := Buffer+StringOfChar(' ', 14);                                             //010
Buffer := Buffer+FORMAT_STRING(GetControlText('EAPE'), 4);                          //012
Buffer := Buffer+StringOfChar(' ', 5);                                              //013
Buffer := Buffer+FORMAT_STRING(GetControlText('ERAISONSOC'), 50);                   //015
Buffer := Buffer+FORMAT_STRING(GetControlText('EADRCOMP'), 32);                     //019
Buffer := Buffer+StringOfChar(' ', 1);                                              //020
Buffer := Buffer+FORMAT_STRING(GetControlText('EADRNUM'), 4);                       //022
Buffer := Buffer+FORMAT_STRING(GetControlText('CADRBIS'), 1);                       //023
Buffer := Buffer+StringOfChar(' ', 1);                                              //024
Buffer := Buffer+FORMAT_STRING(GetControlText('EADRNOMVOIE'), 26);                  //025
Buffer := Buffer+FORMAT_STRING(GetControlText('EADRINSEE'), 5);                     //027
Buffer := Buffer+StringOfChar(' ', 1);                                              //028
Buffer := Buffer+FORMAT_STRING(GetControlText('EADRNOMCOM'), 26);                   //029
Buffer := Buffer+FORMAT_STRING(GetControlText('EADRCP'), 5);                        //031
Buffer := Buffer+StringOfChar(' ', 1);                                              //032
Buffer := Buffer+FORMAT_STRING(GetControlText('EADRBURDISTRIB'), 26);               //033
Buffer := Buffer+StringOfChar('0', 8);                                              //035
{PT1
Buffer := Buffer+'?';                                                               //036
}
if (GetControlText('CTYPE') = '1') then
   Buffer := Buffer+'X'                                                             //036
else
   Buffer := Buffer+'G';                                                            //036

NumEtab := GetControlText('CDEPOSE');
StEtab := 'SELECT ET_ETABLISSEMENT, ET_SIRET, ET_ADRESSE1, ET_ADRESSE2,'+
          ' ET_ADRESSE3, ET_CODEPOSTAL, ET_VILLE'+
          ' FROM ETABLISS WHERE'+
          ' ET_ETABLISSEMENT="'+NumEtab+'"';
QRechEtab:=OpenSql(StEtab,TRUE);
if (not QRechEtab.EOF) then
   begin
   BufOrig := QRechEtab.FindField('ET_SIRET').AsString;
{$IFDEF COMPTA}
  EpureChar(BufOrig, BufDest);
  if not VerifSiret(BufDest) then
{$ELSE}
   ForceNumerique(BufOrig, BufDest);
   if ControlSiret(BufDest)=False then
{$ENDIF}
      PGIBox ('Le SIRET de l''établissement déposant n''est pas valide.',
              'Préparation TD Bilatéral');

   Buffer := Buffer+FORMAT_STRING(BufDest, 14);                                     //037
   Buffer := Buffer+StringOfChar(' ', 5);                                           //038

   if (QRechEtab.FindField('ET_ADRESSE2').AsString <> '') then
      begin
      if (QRechEtab.FindField('ET_ADRESSE3').AsString <> '') then
         Buf := QRechEtab.FindField('ET_ADRESSE2').AsString+' '+
                   QRechEtab.FindField('ET_ADRESSE3').AsString
      else
         Buf := QRechEtab.FindField('ET_ADRESSE2').AsString;
      end
   else
      if (QRechEtab.FindField('ET_ADRESSE3').AsString <> '') then
         Buf := QRechEtab.FindField('ET_ADRESSE3').AsString;

   Buffer := Buffer+FORMAT_STRING(Buf, 32);                                         //039
   Buffer := Buffer+StringOfChar(' ', 1);                                           //040

   if (QRechEtab.FindField('ET_ADRESSE1').AsString <> '') then
      AdresseNormalisee (QRechEtab.FindField('ET_ADRESSE1').AsString, numero,
                         NomRue);

   Buffer := Buffer+FORMAT_STRING(numero,4) ;                                       //042
   Buffer := Buffer+StringOfChar(' ', 1);                                           //043
   Buffer := Buffer+StringOfChar(' ', 1);                                           //044
   Buffer := Buffer+FORMAT_STRING(NomRue,26) ;                                      //045
   Buffer := Buffer+StringOfChar(' ', 5);                                           //047
   Buffer := Buffer+StringOfChar(' ', 1);                                           //048
   Buffer := Buffer+StringOfChar(' ', 26) ;                                         //049
   Buffer := Buffer+FORMAT_STRING(QRechEtab.FindField('ET_CODEPOSTAL').AsString,5) ;//051
   Buffer := Buffer+StringOfChar(' ', 1);                                           //052
   Buffer := Buffer+FORMAT_STRING(QRechEtab.FindField('ET_VILLE').AsString,26) ;    //053
   Buffer := Buffer+StringOfChar(' ', 1);                                           //054
   Buffer := Buffer+StringOfChar(' ', 288);                                         //055
   end;

Writeln(FDADS, Buffer);
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 30/06/2003
Modifié le ... :   /  /
Description .. : Fonction permettant d'écrire les données concernant
Suite ........ : les établissements dans le pré-fichier
Mots clefs ... : PAIE,PGDADSB
*****************************************************************}
procedure TOF_PG_DADS2_PREP.EcritureFicEtab(TEtabD : TOB; var Taxe : TTAXE); // pt15
var
Buffer , codeetab, strech, clesal, asstaxeapp, assformprof, segment, donnee , exerciceTD: string;  // pt15
BaseApp, BaseFPCCDD, BaseFPC : double; // pt15
Qrechdads : Tquery; // pt15

begin
// deb pt15
asstaxeapp := 'N';  Assformprof := 'N'; BaseFPCCDD := 0; BaseApp := 0; BaseFPC := 0;

ExerciceTD := TEtabD.getValue('VALIDITE');
CodeEtab := TEtabD.getValue('CODEETAB');
clesal := '**' + codeetab;

if (GetControlText('CTYPE') <> '1') then     // si non honoraires seuls
begin
  // Recherche si l'établissement est assujetti à la taxe apprentissage et formation continu
  // ou non dans DADSDETAIL
  strech := 'SELECT PDS_DONNEE , PDS_SEGMENT, PDS_DONNEEAFFICH FROM DADSDETAIL WHERE '+
          ' PDS_SALARIE = "'+clesal+'" AND'+
          ' PDS_TYPE="001" AND'+
          ' PDS_ORDRE=2 AND'+
          ' PDS_EXERCICEDADS = "'+ExerciceTD+'" AND'+
          ' PDS_SEGMENT LIKE "S80.G62%" '+
          ' order by PDS_SALARIE, PDS_TYPE, PDS_ORDRE, PDS_SEGMENT';
  Qrechdads := opensql(strech, true);

  while (not Qrechdads.EOF) do
  begin
    segment := Qrechdads.findfield('PDS_SEGMENT').asstring;
    donnee := Qrechdads.findfield('PDS_DONNEE').asstring;
    if (segment = 'S80.G62.05.001') and (donnee = '01') then asstaxeapp := 'I';
    if (segment = 'S80.G62.10.001') and (donnee = '01') then assformprof := 'I';
    if (segment = 'S80.G62.05.002.001') then  baseApp := strtofloat (Qrechdads.findfield('PDS_DONNEEAFFICH').asstring);
    if (segment = 'S80.G62.10.002.001') then BaseFPCCDD := strtofloat (Qrechdads.findfield('PDS_DONNEEAFFICH').asstring);
    if (segment = 'S80.G62.10.003.001') then BaseFPC := strtofloat (Qrechdads.findfield('PDS_DONNEEAFFICH').asstring);
    Qrechdads.next;
  end;
  ferme (Qrechdads);
  if basefpccdd <0 then basefpccdd := 0;
  if basefpc <0 then basefpc := 0;

end;

Taxe.TcodeEtab := Codeetab;
Taxe.TAssApp := asstaxeapp;
Taxe.TTotalTA := baseapp;
Taxe.Tassformprof := assformprof;
Taxe.Tbasefpccdd := basefpccdd;
Taxe.TbaseFPC := baseFPC;
// fin pt15

Buffer := FORMAT_STRING(TEtabD.getValue('SIRET'), 14);            //002
Buffer := Buffer+FORMAT_STRING(TEtabD.getValue('SECTION'), 2);    //005
Buffer := Buffer+FORMAT_STRING(TEtabD.getValue('VALIDITE'), 4);   //006
Buffer := Buffer+FORMAT_STRING(TEtabD.getValue('TYPE'), 1);       //007
Buffer := Buffer+'020';                                           //009
Buffer := Buffer+StringOfChar(' ', 14);                           //013
Buffer := Buffer+FORMAT_STRING(TEtabD.getValue('APE'), 4);        //014
Buffer := Buffer+StringOfChar(' ', 1);                            //015
Buffer := Buffer+FORMAT_STRING(TEtabD.getValue('MODIFSIRET'), 14);//016
Buffer := Buffer+StringOfChar(' ', 41);                           //022
Buffer := Buffer+FORMAT_STRING(TEtabD.getValue('LIBELLE'), 50);   //024
Buffer := Buffer+FORMAT_STRING(TEtabD.getValue('ADRCOMPL'), 32);  //027
Buffer := Buffer+StringOfChar(' ', 1);                            //028
Buffer := Buffer+FORMAT_STRING(TEtabD.getValue('ADRNUM'), 4);     //030
Buffer := Buffer+FORMAT_STRING(TEtabD.getValue('ADRBIS'), 1);     //031
Buffer := Buffer+StringOfChar(' ', 1);                            //032
Buffer := Buffer+FORMAT_STRING(TEtabD.getValue('ADRNOM'), 26);    //033
Buffer := Buffer+FORMAT_STRING(TEtabD.getValue('ADRINSEE'), 5);   //035
Buffer := Buffer+StringOfChar(' ', 1);                            //036
Buffer := Buffer+FORMAT_STRING(TEtabD.getValue('ADRCOMMUNE'), 26);//037
Buffer := Buffer+FORMAT_STRING(TEtabD.getValue('ADRCP'), 5);      //039
Buffer := Buffer+StringOfChar(' ', 1);                            //040
Buffer := Buffer+FORMAT_STRING(TEtabD.getValue('ADRDISTRIB'), 26);//041
Buffer := Buffer+FORMAT_STRING(TEtabD.getValue('PROFESSION'), 40);//045
Buffer := Buffer+StringOfChar(' ', 53);                           //046
Buffer := Buffer+FORMAT_STRING(TEtabD.getValue('TAXESAL'), 1);    //047
//Buffer := Buffer+StringOfChar(' ', 301);                        //048      PT14
//Buffer := Buffer+StringOfChar('N', 1);                           //048      PT14
//Buffer := Buffer+StringOfChar('N', 1);                           //049      PT14
Buffer := Buffer+Format_string(Asstaxeapp, 1);                      //048      PT15
Buffer := Buffer+format_string(Assformprof, 1);                     //049      PT15

Buffer := Buffer+StringOfChar(' ', 299);                          //050      PT14

Writeln(FDADS, Buffer);
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 01/07/2003
Modifié le ... :   /  /
Description .. : Fonction permettant d'écrire les données concernant
Suite ........ : les salariés dans le pré-fichier
Mots clefs ... : PAIE,PGDADSB
*****************************************************************}
procedure TOF_PG_DADS2_PREP.EcritureFicSal(TSalD : TOB; var TotalEtabS : TotalSalEtab);
var
BufChamp, Buffer, Dept, NumSS : string;
aa, mm, jj : word;
begin
Buffer := FORMAT_STRING(TSalD.getValue('PD2_SIRET'), 14);                      //002
Buffer := Buffer+FORMAT_STRING(TSalD.getValue('PD2_SECTIONETAB'), 2);          //005
Buffer := Buffer+FORMAT_STRING(TSalD.getValue('PD2_VALIDITE'), 4);             //006
Buffer := Buffer+FORMAT_STRING(GetControlText('CTYPE'), 1);                    //007
Buffer := Buffer+'200';                                                        //009
Buffer := Buffer+StringOfChar(' ', 10);                                        //012
NumSS := FORMAT_STRING(TSalD.getValue('PD2_NUMEROSS'), 15);
Buffer := Buffer+NumSS;                                                        //013

DecodeDate(TSalD.GetValue('PD2_DATENAISSANCE'), aa, mm, jj);
Buffer := Buffer+FormatFloat('00', jj)+FormatFloat('00', mm)+IntToStr(aa);                        //025
Dept := TSalD.GetValue('PD2_DEPART');
if (Dept = '20A') then
   Dept := '2A';
if (Dept = '20B') then
   Dept := '2B';
Buffer := Buffer+FORMAT_STRING(Dept, 2);                                       //029
if (Dept = '99') then
   begin
   Buffer := Buffer+Copy(NumSS, 8, 3);                                         //030
   Buffer := Buffer+FORMAT_STRING(RechDom('TTPAYS', TSalD.GetValue('PD2_PAYS'),
                                          FALSE), 26);                         //031
   end
else
   begin
   Buffer := Buffer+'000';                                                     //030
   Buffer := Buffer+FORMAT_STRING(TSalD.GetValue('PD2_COMMUNENAISS'), 26);     //031
   end;
Buffer := Buffer+FORMAT_STRING(TSalD.GetValue('PD2_CIVILITE'), 3);             //034
if ((TSalD.GetValue('PD2_NOMJF') = '') or
   (TSalD.GetValue('PD2_NOMJF') = null)) then
   Buffer := Buffer+FORMAT_STRING(TSalD.GetValue('PD2_NOMMARITAL'), 30)        //035
else
   Buffer := Buffer+FORMAT_STRING(TSalD.GetValue('PD2_NOMJF'), 30);            //035
Buffer := Buffer+FORMAT_STRING(TSalD.GetValue('PD2_PRENOMDADS'), 20);          //036
if ((TSalD.GetValue('PD2_NOMJF') <> '') and
   (TSalD.GetValue('PD2_NOMJF') <> null)) then
   Buffer := Buffer+FORMAT_STRING(TSalD.GetValue('PD2_NOMMARITAL'), 30)        //037
else
   Buffer := Buffer+StringOfChar(' ', 30);                                      //037
Buffer := Buffer+FORMAT_STRING(TSalD.GetValue('PD2_ADRCOMPL'), 32);            //040
Buffer := Buffer+' ';                                                          //041
Buffer := Buffer+FORMAT_STRING(TSalD.GetValue('PD2_ADRNUM'), 4);               //043
Buffer := Buffer+FORMAT_STRING(TSalD.GetValue('PD2_ADRBISTER'), 1);            //044
Buffer := Buffer+' ';                                                          //045
Buffer := Buffer+FORMAT_STRING(TSalD.GetValue('PD2_ADRNOM'), 26);              //046
Buffer := Buffer+FORMAT_STRING(TSalD.GetValue('PD2_ADRCOMMINSEE'), 5);         //048
Buffer := Buffer+' ';                                                          //049
Buffer := Buffer+FORMAT_STRING(TSalD.GetValue('PD2_ADRCOMMUNE'), 26);          //050
Buffer := Buffer+FORMAT_STRING(TSalD.GetValue('PD2_CODEPOSTAL'), 5);           //052
Buffer := Buffer+' ';                                                          //053
Buffer := Buffer+FORMAT_STRING(TSalD.GetValue('PD2_BUREAUDISTRIB'), 26);       //054
Buffer := Buffer+FORMAT_STRING(TSalD.GetValue('PD2_EMPLOIQUALIF'), 30);        //058
Buffer := Buffer+FORMAT_STRING(TSalD.GetValue('PD2_CODEEMPLOI'), 4);           //059
{PT8
Buffer := Buffer+StringOfChar(' ', 2);
}
if (TSalD.GetValue('PD2_CHQEMPLOI')= 'X') then
   Buffer := Buffer+'C'                                                        //060
else
   Buffer := Buffer+' ';                                                       //060
Buffer := Buffer+StringOfChar(' ', 1);                                         //061
//FIN PT8
Buffer := Buffer+FORMAT_STRING(TSalD.GetValue('PD2_PERIODE1'), 8);             //065
{PT11
Buffer := Buffer+FORMAT_STRING(TSalD.GetValue('PD2_PERIODE2'), 8);
}
if (TSalD.GetValue('PD2_PERIODE2') <> '') then
   Buffer:= Buffer+FORMAT_STRING(TSalD.GetValue('PD2_PERIODE2'), 8)            //070
else
   Buffer:= Buffer+StringOfChar('0', 8);                                       //070
//FIN PT11   
Buffer := Buffer+StringOfChar(' ', 2);                                         //075
Buffer := Buffer+FORMAT_STRING(TSalD.GetValue('PD2_CONDEMPLOI'), 1);           //076
Buffer := Buffer+FORMAT_STRING(TSalD.GetValue('PD2_TRAVETRANG'), 1);           //077
if (TSalD.GetValue('PD2_SORTIEDEFINIT')= 'X') then
   Buffer := Buffer+'D'                                                        //078
else
   Buffer := Buffer+' ';                                                       //078
Buffer := Buffer+StringOfChar(' ', 80);
Buffer := Buffer+FormatFloat('0000000', TSalD.GetValue('PD2_BASEBRUTE'));      //100
TotalEtabS.TotalS100 := TotalEtabS.TotalS100+TSalD.GetValue('PD2_BASEBRUTE');
{PT7
Buffer := Buffer+StringOfChar(' ', 7);
}
{PT8
Buffer := Buffer+StringOfChar('0', 7);
}
// Buffer := Buffer+FormatFloat('0000000', TSalD.GetValue('PD2_INDEMIMPATRI'));   //102  // pt14
Buffer := Buffer+StringofChar('0', 7);  // PT14
//FIN PT8
Buffer := Buffer+FormatFloat('0000000', TSalD.GetValue('PD2_INDEMEXPATRI'));   //103
{PT8
TotalEtabS.TotalS102103 := TotalEtabS.TotalS102103+
                           TSalD.GetValue('PD2_INDEMEXPATRI');
}
TotalEtabS.TotalS102103 := TotalEtabS.TotalS102103+
//                           TSalD.GetValue('PD2_INDEMIMPATRI')+                 PT14
                           TSalD.GetValue('PD2_INDEMEXPATRI');
//FIN PT8
Buffer := Buffer+FormatFloat('0000000', TSalD.GetValue('PD2_AVANTAGENATV'));   //105
TotalEtabS.TotalS105 := TotalEtabS.TotalS105+TSalD.GetValue('PD2_AVANTAGENATV');
if ((TSalD.GetValue('PD2_AVANTAGENATN')<>'') and
   (TSalD.GetValue('PD2_AVANTAGENATN')<>null)) then
   begin
   BufChamp:= TSalD.GetValue('PD2_AVANTAGENATN');                              //106
   if (Copy (BufChamp, 1, 1)='N') then
      Buffer:= Buffer+Copy (BufChamp, 1, 1)
   else
      Buffer:= Buffer+' ';
   if (Copy (BufChamp, 2, 1)='L') then
      Buffer:= Buffer+Copy (BufChamp, 2, 1)
   else
      Buffer:= Buffer+' ';
   if (Copy (BufChamp, 3, 1)='V') then
      Buffer:= Buffer+Copy (BufChamp, 3, 1)
   else
      Buffer:= Buffer+' ';
   if (Copy (BufChamp, 4, 1)='A') then
      Buffer:= Buffer+Copy (BufChamp, 4, 1)
   else
      Buffer:= Buffer+' ';
   end
else
   Buffer:= Buffer+StringOfChar(' ', 4);
if ((TSalD.GetValue('PD2_NTIC')<>'') and
   (TSalD.GetValue('PD2_NTIC')<>null)) then
   Buffer:= Buffer+TSalD.GetValue('PD2_NTIC')                                  //111
else
   Buffer:= Buffer+' ';
Buffer := Buffer+FormatFloat('000000', TSalD.GetValue('PD2_RETENUESALV'));     //112
if (TSalD.GetValue('PD2_REMPOURBOIRE')= 'X') then
   Buffer := Buffer+'P'                                                        //113
else
   Buffer := Buffer+' ';                                                       //113
Buffer := Buffer+FormatFloat('0000000', TSalD.GetValue('PD2_FRAISPROFM'));     //117
TotalEtabS.TotalS117 := TotalEtabS.TotalS117+TSalD.GetValue('PD2_FRAISPROFM');
if ((TSalD.GetValue('PD2_FRAISPROFN')<>'') and
   (TSalD.GetValue('PD2_FRAISPROFN')<>null)) then
   begin
   BufChamp:= TSalD.GetValue('PD2_FRAISPROFN');                                //118
   if (Copy (BufChamp, 1, 1)='F') then
      Buffer:= Buffer+Copy (BufChamp, 1, 1)
   else
      Buffer:= Buffer+' ';
   if (Copy (BufChamp, 2, 1)='R') then
      Buffer:= Buffer+Copy (BufChamp, 2, 1)
   else
      Buffer:= Buffer+' ';
   if (Copy (BufChamp, 3, 1)='P') then
      Buffer:= Buffer+Copy (BufChamp, 3, 1)
   else
      Buffer:= Buffer+' ';
   if (Copy (BufChamp, 4, 1)='D') then
      Buffer:= Buffer+Copy (BufChamp, 4, 1)
   else
      Buffer:= Buffer+' ';
   end
else
   Buffer:= Buffer+StringOfChar(' ', 4);
Buffer := Buffer+FormatFloat('000000', TSalD.GetValue('PD2_CHQVACANCEM'));     //125
TotalEtabS.TotalS125 := TotalEtabS.TotalS125+TSalD.GetValue('PD2_CHQVACANCEM');
Buffer := Buffer+StringOfChar(' ', 5);                                         //126
Buffer := Buffer+FormatFloat('0000000', TSalD.GetValue('PD2_TOTALIMPOSAB'));   //132
TotalEtabS.TotalS132 := TotalEtabS.TotalS132+TSalD.GetValue('PD2_TOTALIMPOSAB');
Buffer := Buffer+FORMAT_STRING(TSalD.GetValue('PD2_REMMULTIETAB'), 6);         //133
Buffer := Buffer+FormatFloat('0000000', TSalD.GetValue('PD2_BASEIMPO1'));      //134
TotalEtabS.TotalS134 := TotalEtabS.TotalS134+TSalD.GetValue('PD2_BASEIMPO1');
Buffer := Buffer+FormatFloat('0000000', TSalD.GetValue('PD2_BASEIMPO2'));      //135
TotalEtabS.TotalS135 := TotalEtabS.TotalS135+TSalD.GetValue('PD2_BASEIMPO2');
Buffer := Buffer+FormatFloat('000000', TSalD.GetValue('PD2_RETENUESOURC'));    //136
TotalEtabS.TotalS136 := TotalEtabS.TotalS136+TSalD.GetValue('PD2_RETENUESOURC');
Buffer := Buffer+FormatFloat('000000', TSalD.GetValue('PD2_EPARGNERETRAI'));   //137
Buffer := Buffer+' ';                                                          //138
Buffer := Buffer+FORMAT_STRING(TSalD.GetValue('PD2_CODEABSENCE'), 1);          //139
Buffer := Buffer+StringOfChar(' ', 5);                                         //140
Buffer := Buffer+FormatFloat('0000000', TSalD.GetValue('PD2_REVENUSACTIV'));   //142
TotalEtabS.TotalS102103 := TotalEtabS.TotalS102103+
                           TSalD.GetValue('PD2_REVENUSACTIV');
Buffer := Buffer+FormatFloat('0000', TSalD.GetValue('PD2_NBREHEURE'));         //143
{PT8
Buffer := Buffer+StringOfChar('0', 7);
}
Buffer := Buffer+FormatFloat('0000000', TSalD.GetValue('PD2_AUTREREVENUS'));   //144
TotalEtabS.TotalS102103 := TotalEtabS.TotalS102103+
                           TSalD.GetValue('PD2_AUTREREVENUS');
//FIN PT8
{PT10
Buffer := Buffer+StringOfChar(' ', 109);
}
Buffer := Buffer+StringOfChar('0', 6);                                         //145
//Buffer := Buffer+StringOfChar(' ', 103);                                     //146    PT14
Buffer := Buffer+ FormatFloat('0000000', TSalD.GetValue('PD2_INDEMIMPATRI'));  //146    PT14
TotalEtabS.TotalS146 := TotalEtabS.TotalS146+TSalD.GetValue('PD2_INDEMIMPATRI');//      PT14                                   //146    PT14
Buffer := Buffer+StringOfChar(' ', 96);                                        //148    PT14
//FIN PT10

MoveCurProgressForm ('Ecriture 200');
MoveCur(False) ;
Writeln(FDADS, Buffer);
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 03/07/2003
Modifié le ... :   /  /
Description .. : Fonction permettant d'écrire les données concernant
Suite ........ : les honoraires dans le pré-fichier
Mots clefs ... : PAIE,PGDADSB
*****************************************************************}
procedure TOF_PG_DADS2_PREP.EcritureFicHon(THonD : TOB; var TotalEtabH : TotalHonEtab);
var
BufChamp, Buffer : string;
begin
Buffer := FORMAT_STRING(THonD.getValue('PDH_SIRET'), 14);                      //002
Buffer := Buffer+FORMAT_STRING(THonD.getValue('PDH_SECTIONETAB'), 2);          //005
Buffer := Buffer+FORMAT_STRING(THonD.getValue('PDH_VALIDITE'), 4);             //006
Buffer := Buffer+FORMAT_STRING(GetControlText('CTYPE'), 1);                    //007
Buffer := Buffer+'210';                                                        //009
Buffer := Buffer+FORMAT_STRING(THonD.getValue('PDH_SIRETBEN'), 14);            //012
Buffer := Buffer+FORMAT_STRING(THonD.getValue('PDH_NOMBEN'), 30);              //013
Buffer := Buffer+FORMAT_STRING(THonD.GetValue('PDH_PRENOMBEN'), 20);           //014
Buffer := Buffer+FORMAT_STRING(THonD.GetValue('PDH_RAISONSOCBEN'), 50);        //015
Buffer := Buffer+FORMAT_STRING(THonD.GetValue('PDH_PROFESSIONBEN'), 30);       //016
Buffer := Buffer+FORMAT_STRING(THonD.GetValue('PDH_ADRCOMPL'), 32);            //019
Buffer := Buffer+' ';                                                          //020
Buffer := Buffer+FORMAT_STRING(THonD.GetValue('PDH_ADRNUM'), 4);               //022
Buffer := Buffer+FORMAT_STRING(THonD.GetValue('PDH_ADRBISTER'), 1);            //023
Buffer := Buffer+' ';                                                          //024
Buffer := Buffer+FORMAT_STRING(THonD.GetValue('PDH_ADRNOM'), 26);              //025
Buffer := Buffer+FORMAT_STRING(THonD.GetValue('PDH_ADRCOMMINSEE'), 5);         //027
Buffer := Buffer+' ';                                                          //028
Buffer := Buffer+FORMAT_STRING(THonD.GetValue('PDH_ADRCOMMUNE'), 26);          //029
Buffer := Buffer+FORMAT_STRING(THonD.GetValue('PDH_CODEPOSTAL'), 5);           //032
Buffer := Buffer+' ';                                                          //033
Buffer := Buffer+FORMAT_STRING(THonD.GetValue('PDH_BUREAUDISTRIB'), 26);       //034
Buffer := Buffer+FormatFloat('0000000000', THonD.GetValue('PDH_REMHONOR'));    //038
TotalEtabH.TotalH038 := TotalEtabH.TotalH038+THonD.GetValue('PDH_REMHONOR');
Buffer := Buffer+FormatFloat('0000000000', THonD.GetValue('PDH_REMCOMMISS'));  //039
TotalEtabH.TotalH039 := TotalEtabH.TotalH039+THonD.GetValue('PDH_REMCOMMISS');
Buffer := Buffer+FormatFloat('0000000000', THonD.GetValue('PDH_REMCOURTAGE')); //040
TotalEtabH.TotalH040 := TotalEtabH.TotalH040+THonD.GetValue('PDH_REMCOURTAGE');
Buffer := Buffer+FormatFloat('0000000000', THonD.GetValue('PDH_REMRISTOURNE'));//041
TotalEtabH.TotalH041 := TotalEtabH.TotalH041+THonD.GetValue('PDH_REMRISTOURNE');
Buffer := Buffer+FormatFloat('0000000000', THonD.GetValue('PDH_REMJETON'));    //042
TotalEtabH.TotalH042 := TotalEtabH.TotalH042+THonD.GetValue('PDH_REMJETON');
Buffer := Buffer+FormatFloat('0000000000', THonD.GetValue('PDH_REMAUTEUR'));   //043
TotalEtabH.TotalH043 := TotalEtabH.TotalH043+THonD.GetValue('PDH_REMAUTEUR');
Buffer := Buffer+FormatFloat('0000000000', THonD.GetValue('PDH_REMINVENT'));   //044
TotalEtabH.TotalH044 := TotalEtabH.TotalH044+THonD.GetValue('PDH_REMINVENT');
Buffer := Buffer+FormatFloat('0000000000', THonD.GetValue('PDH_REMAUTRE'));    //045
TotalEtabH.TotalH045 := TotalEtabH.TotalH045+THonD.GetValue('PDH_REMAUTRE');
Buffer := Buffer+FormatFloat('0000000000', THonD.GetValue('PDH_REMINDEMNITE'));//046
TotalEtabH.TotalH046 := TotalEtabH.TotalH046+THonD.GetValue('PDH_REMINDEMNITE');
Buffer := Buffer+FormatFloat('0000000000', THonD.GetValue('PDH_REMAVANTAGE')); //047
TotalEtabH.TotalH047 := TotalEtabH.TotalH047+THonD.GetValue('PDH_REMAVANTAGE');
Buffer := Buffer+FormatFloat('0000000000', THonD.GetValue('PDH_RETENUESOURC'));//048
TotalEtabH.TotalH048 := TotalEtabH.TotalH048+THonD.GetValue('PDH_RETENUESOURC');
{PT2-4
Buffer := Buffer+FORMAT_STRING(THonD.GetValue('PDH_AVANTAGENATN'), 4);
Buffer := Buffer+' ';
Buffer := Buffer+FORMAT_STRING(THonD.GetValue('PDH_CHARGEINDEMN'), 3);
Buffer := Buffer+FORMAT_STRING(THonD.GetValue('PDH_TAUXSOURCE'), 2);
}
//PT6
if ((THonD.GetValue('PDH_AVANTAGENATN')<>'') and
   (THonD.GetValue('PDH_AVANTAGENATN')<>null)) then
   begin
   BufChamp:= THonD.GetValue('PDH_AVANTAGENATN');                              //050
   if (Copy (BufChamp, 1, 1)='N') then
      Buffer:= Buffer+Copy (BufChamp, 1, 1)
   else
      Buffer:= Buffer+' ';
   if (Copy (BufChamp, 2, 1)='L') then
      Buffer:= Buffer+Copy (BufChamp, 2, 1)
   else
      Buffer:= Buffer+' ';
   if (Copy (BufChamp, 3, 1)='V') then
      Buffer:= Buffer+Copy (BufChamp, 3, 1)
   else
      Buffer:= Buffer+' ';
   if (Copy (BufChamp, 4, 1)='A') then
      Buffer:= Buffer+Copy (BufChamp, 4, 1)
   else
      Buffer:= Buffer+' ';
   end
else
   Buffer:= Buffer+StringOfChar(' ', 4);
{PT5
Buffer := Buffer+' ';
}
if ((THonD.GetValue('PDH_NTIC')<>'') and
   (THonD.GetValue('PDH_NTIC')<>null)) then
   Buffer:= Buffer+THonD.GetValue('PDH_NTIC')                                  //055
else
   Buffer:= Buffer+' ';
//FIN PT5
if ((THonD.GetValue('PDH_CHARGEINDEMN')<>'') and
   (THonD.GetValue('PDH_CHARGEINDEMN')<>null)) then
   begin
   BufChamp:= THonD.GetValue('PDH_CHARGEINDEMN');                              //057
   if (Copy (BufChamp, 1, 1)='F') then
      Buffer:= Buffer+Copy (BufChamp, 1, 1)
   else
      Buffer:= Buffer+' ';
   if (Copy (BufChamp, 2, 1)='R') then
      Buffer:= Buffer+Copy (BufChamp, 2, 1)
   else
      Buffer:= Buffer+' ';
   if (Copy (BufChamp, 3, 1)='P') then
      Buffer:= Buffer+Copy (BufChamp, 3, 1)
   else
      Buffer:= Buffer+' ';
   end
else
   Buffer:= Buffer+StringOfChar(' ', 3);
if ((THonD.GetValue('PDH_TAUXSOURCE')<>'') and
   (THonD.GetValue('PDH_TAUXSOURCE')<>null)) then
   begin
   BufChamp:= THonD.GetValue('PDH_TAUXSOURCE');                                //061
   if (Copy (BufChamp, 1, 1)='R') then
      Buffer:= Buffer+Copy (BufChamp, 1, 1)
   else
      Buffer:= Buffer+' ';
   if (Copy (BufChamp, 2, 1)='D') then
      Buffer:= Buffer+Copy (BufChamp, 2, 1)
   else
      Buffer:= Buffer+' ';
   end
else
   Buffer:= Buffer+StringOfChar(' ', 2);
//FIN PT6
Buffer := Buffer+FormatFloat('0000000000', THonD.GetValue('PDH_TVAAUTEUR'));   //064
Buffer := Buffer+StringOfChar(' ', 245);                                       //065

MoveCurProgressForm ('Ecriture 210');
MoveCur(False) ;
Writeln(FDADS, Buffer);
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 03/07/2003
Modifié le ... :   /  /
Description .. : Fonction permettant d'écrire les données concernant
Suite ........ : les Totaux établissement dans le pré-fichier
Mots clefs ... : PAIE,PGDADSB
*****************************************************************}
//procedure TOF_PG_DADS2_PREP.EcritureFicTotEtab(Siret, Section, DValidite,
//                                              TypeDADS, CodeEtab : string;
//                                               var TotalEtabS : TotalSalEtab;
//                                               TotalEtabH : TotalHonEtab);
procedure TOF_PG_DADS2_PREP.EcritureFicTotEtab(Siret, Section, DValidite,     // pt15
                                            TypeDADS, CodeEtab : string;
                                            taxe : TTAXE;
                                            var TotalEtabS : TotalSalEtab;
                                            TotalEtabH : TotalHonEtab);
var
Buffer, SignNom, SignPrenom, SignTelFonc, StDate, StEntree, StSortie : string;
TaxeSal : double;
QEntree, QSortie : TQuery;
NbSal : integer;
begin
if (CodeEtab = GetControlText('CDEPOSE')) then
   TaxeSal := StrToFloat(GetControlText('ETAXESAL'))
else
   TaxeSal := 0;
SignNom := GetControlText('ESIGNNOM');
SignPrenom := GetControlText('ESIGNPRENOM');
SignTelFonc := GetControlText('ESIGNTELFONC');
NbSal := 0;
Buffer := FORMAT_STRING(Siret, 14);                                            //002
Buffer := Buffer+FORMAT_STRING(Section, 2);                                    //005
Buffer := Buffer+FORMAT_STRING(DValidite, 4);                                  //006
Buffer := Buffer+FORMAT_STRING(TypeDADS, 1);                                   //007
Buffer := Buffer+'300';                                                        //009
Buffer := Buffer+StringOfChar(' ', 36);                                        //012
Buffer := Buffer+FormatFloat('000000000000', TotalEtabS.TotalS100);            //018
Buffer := Buffer+FormatFloat('000000000000', TotalEtabS.TotalS102103);         //019
Buffer := Buffer+FormatFloat('000000000000', TotalEtabS.TotalS105);            //020
Buffer := Buffer+FormatFloat('000000000000', TotalEtabS.TotalS117);            //021
Buffer := Buffer+FormatFloat('000000000000', TotalEtabS.TotalS125);            //022
Buffer := Buffer+FormatFloat('000000000000', TotalEtabS.TotalS132);            //025
Buffer := Buffer+FormatFloat('000000000000', TotalEtabS.TotalS134);            //026
Buffer := Buffer+FormatFloat('000000000000', TotalEtabS.TotalS135);            //027
Buffer := Buffer+FormatFloat('000000000000', TotalEtabS.TotalS136);            //028
Buffer := Buffer+FormatFloat('000000000000', TotalEtabH.TotalH038);            //031
Buffer := Buffer+FormatFloat('000000000000', TotalEtabH.TotalH039);            //032
Buffer := Buffer+FormatFloat('000000000000', TotalEtabH.TotalH040);            //033
Buffer := Buffer+FormatFloat('000000000000', TotalEtabH.TotalH041);            //034
Buffer := Buffer+FormatFloat('000000000000', TotalEtabH.TotalH042);            //035
Buffer := Buffer+FormatFloat('000000000000', TotalEtabH.TotalH043);            //036
Buffer := Buffer+FormatFloat('000000000000', TotalEtabH.TotalH044);            //037
Buffer := Buffer+FormatFloat('000000000000', TotalEtabH.TotalH045);            //038
Buffer := Buffer+FormatFloat('000000000000', TotalEtabH.TotalH046);            //039
Buffer := Buffer+FormatFloat('000000000000', TotalEtabH.TotalH047);            //040
Buffer := Buffer+FormatFloat('000000000000', TotalEtabH.TotalH048);            //041
Buffer := Buffer+StringOfChar(' ', 12);                                        //042
Buffer := Buffer+FormatFloat('000000000000', TaxeSal);                         //046
TotalEtabS.TotalTS := TaxeSal;
//Buffer := Buffer+StringOfChar(' ', 12);                                      //047 PT14
Buffer := Buffer+FormatFloat('000000000000', TotalEtabS.TotalS146);            //047 PT14                                       

StDate:=UsDateTime(StrToDate('31/12/'+PGExercice));
StEntree := 'SELECT COUNT(PSA_SALARIE) AS NBENTREE'+
            ' FROM SALARIES '+
            ' WHERE PSA_DATEENTREE<"'+StDate+'"'+
            ' AND PSA_DATEENTREE<>"'+UsDateTime(IDate1900)+'"'+
            ' AND PSA_DATEENTREE IS NOT NULL'+
            ' AND PSA_ETABLISSEMENT="'+CodeEtab+'"';
QEntree:=OpenSql(StEntree,True);

StSortie := 'SELECT COUNT(PSA_SALARIE) AS NBSORTIE'+
            ' FROM SALARIES'+
            ' WHERE PSA_DATESORTIE<"'+StDate+'"'+
            ' AND PSA_DATESORTIE<>"'+UsDateTime(IDate1900)+'"'+
            ' AND PSA_DATESORTIE IS NOT NULL'+
            ' AND PSA_ETABLISSEMENT="'+CodeEtab+'"';
QSortie:=OpenSql(StSortie,True);

if ((not QEntree.EOF) and (not QSortie.EOF)) then
   NbSal:=(QEntree.FindField('NBENTREE').asinteger)-
          (QSortie.FindField('NBSORTIE').asinteger);
Ferme(QEntree);
Ferme(QSortie);
Buffer := Buffer+FormatFloat('000000', NbSal);                                 //049
TotalEtabS.TotalEff := NbSal;
//Buffer := Buffer+StringOfChar(' ', 134);                                     //050   PT14
// pt15  Buffer := Buffer+StringOfChar('0', 12);                               //050   PT14
// pt15  Buffer := Buffer+StringOfChar('0', 12);                               //051   PT14
// pt15  Buffer := Buffer+StringOfChar('0', 12);                               //052   PT14
Buffer := Buffer + FormatFloat('000000000000', Taxe.TTotalTA);                 //050 pt15
Buffer := Buffer + FormatFloat('000000000000', Taxe.TBaseFPC);                 //051 pt15
Buffer := Buffer + formatFloat('000000000000', Taxe.TBaseFPCCDD);              //052 pt15
TotalEtabS.TotalBaseApp:= Taxe.TTotalTA;                                       // pt15
TotalEtabS.TotalBaseFPC:= Taxe.TBaseFPC;                                       // pt15
TotalEtabS.TotalBaseFPCCDD:= Taxe.TBaseFPCCDD;                                 // pt15
Buffer := Buffer+StringOfChar(' ', 98);                                        //054   PT14

Buffer := Buffer+FORMAT_STRING(SignNom, 30);                                   //065
Buffer := Buffer+FORMAT_STRING(SignPrenom, 20);                                //066
Buffer := Buffer+FORMAT_STRING(SignTelFonc, 39);                               //067
Buffer := Buffer+StringOfChar(' ', 107);                                       //068

Writeln(FDADS, Buffer);
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 03/07/2003
Modifié le ... :   /  /
Description .. : Fonction permettant d'écrire les données concernant
Suite ........ : les Totaux entreprise dans le pré-fichier
Mots clefs ... : PAIE,PGDADSB
*****************************************************************}
procedure TOF_PG_DADS2_PREP.EcritureFicTotDecla(TotalEntre : TotalEntreprise);
var
Buffer : string;
begin
Buffer := FORMAT_STRING(GetControlText('ESIREN'), 9);                            //002
Buffer := Buffer+StringOfChar('9', 12);                                          //004
Buffer := Buffer+'310';                                                          //009
Buffer := Buffer+FormatFloat('00000', TotalEntre.TotEtab);                       //011
Buffer := Buffer+FormatFloat('000000', TotalEntre.TotSal);                       //012
Buffer := Buffer+FormatFloat('000000', TotalEntre.TotHon);                       //013
Buffer := Buffer+StringOfChar(' ', 36);                                          //015
Buffer := Buffer+FormatFloat('000000000000', TotalEntre.TotalEtabS.TotalS100);   //021
Buffer := Buffer+FormatFloat('000000000000', TotalEntre.TotalEtabS.TotalS102103);//022
Buffer := Buffer+FormatFloat('000000000000', TotalEntre.TotalEtabS.TotalS105);   //023
Buffer := Buffer+FormatFloat('000000000000', TotalEntre.TotalEtabS.TotalS117);   //024
Buffer := Buffer+FormatFloat('000000000000', TotalEntre.TotalEtabS.TotalS125);   //025
Buffer := Buffer+FormatFloat('000000000000', TotalEntre.TotalEtabS.TotalS132);   //027
Buffer := Buffer+FormatFloat('000000000000', TotalEntre.TotalEtabS.TotalS134);   //028
Buffer := Buffer+FormatFloat('000000000000', TotalEntre.TotalEtabS.TotalS135);   //029
Buffer := Buffer+FormatFloat('000000000000', TotalEntre.TotalEtabS.TotalS136);   //030
Buffer := Buffer+FormatFloat('000000000000', TotalEntre.TotalEtabH.TotalH038);   //033
Buffer := Buffer+FormatFloat('000000000000', TotalEntre.TotalEtabH.TotalH039);   //034
Buffer := Buffer+FormatFloat('000000000000', TotalEntre.TotalEtabH.TotalH040);   //035
Buffer := Buffer+FormatFloat('000000000000', TotalEntre.TotalEtabH.TotalH041);   //036
Buffer := Buffer+FormatFloat('000000000000', TotalEntre.TotalEtabH.TotalH042);   //037
Buffer := Buffer+FormatFloat('000000000000', TotalEntre.TotalEtabH.TotalH043);   //038
Buffer := Buffer+FormatFloat('000000000000', TotalEntre.TotalEtabH.TotalH044);   //039
Buffer := Buffer+FormatFloat('000000000000', TotalEntre.TotalEtabH.TotalH045);   //040
Buffer := Buffer+FormatFloat('000000000000', TotalEntre.TotalEtabH.TotalH046);   //041
Buffer := Buffer+FormatFloat('000000000000', TotalEntre.TotalEtabH.TotalH047);   //042
Buffer := Buffer+FormatFloat('000000000000', TotalEntre.TotalEtabH.TotalH048);   //043
Buffer := Buffer+StringOfChar(' ', 12);                                          //045
Buffer := Buffer+FormatFloat('000000000000', TotalEntre.TotalEtabS.TotalTS);     //047
//Buffer := Buffer+StringOfChar(' ', 12);                                        //048
Buffer := Buffer+FormatFloat('000000000000', TotalEntre.TotalEtabS.TotalS146);   //048  PT14                                     //048  PT14


Buffer := Buffer+FormatFloat('000000', TotalEntre.TotalEtabS.TotalEff);          //050
//Buffer := Buffer+StringOfChar(' ', 313);                                       //051  PT14
// pt15 Buffer := Buffer+StringOfChar('0', 12);                                          //051  PT14
// pt15 Buffer := Buffer+StringOfChar('0', 12);                                          //052  PT14
// pt15 Buffer := Buffer+StringOfChar('0', 12);                                          //053  PT14
Buffer := Buffer + FormatFloat('000000000000',TotalEntre.TotalEtabS.TotalBaseApp); // pt15
Buffer := Buffer + FormatFloat('000000000000',TotalEntre.TotalEtabS.TotalBaseFPC); // pt15
Buffer := Buffer + FormatFloat('000000000000',TotalEntre.TotalEtabS.TotalBaseFPCCDD); // pt15
Buffer := Buffer+StringOfChar(' ', 277);                                         //054  PT14

Writeln(FDADS, Buffer);
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 04/07/2003
Modifié le ... :   /  /
Description .. : Fonction permettant d'initialise les données etablissement
Mots clefs ... : PAIE,PGDADSB
*****************************************************************}
procedure TOF_PG_DADS2_PREP.InitTotalEtab(var TotalEtabS : TotalSalEtab; var TotalEtabH : TotalHonEtab);
begin
TotalEtabS.TotalS100:=0;
TotalEtabS.TotalS102103:=0;
TotalEtabS.TotalS105:=0;
TotalEtabS.TotalS117:=0;
TotalEtabS.TotalS125:=0;
TotalEtabS.TotalS132:=0;
TotalEtabS.TotalS134:=0;
TotalEtabS.TotalS135:=0;
TotalEtabS.TotalS136:=0;
TotalEtabS.TotalS146:=0;  // pt14
TotalEtabS.TotalTS:=0;
TotalEtabS.TotalEff:=0;

TotalEtabH.TotalH038:=0;
TotalEtabH.TotalH039:=0;
TotalEtabH.TotalH040:=0;
TotalEtabH.TotalH041:=0;
TotalEtabH.TotalH042:=0;
TotalEtabH.TotalH043:=0;
TotalEtabH.TotalH044:=0;
TotalEtabH.TotalH045:=0;
TotalEtabH.TotalH046:=0;
TotalEtabH.TotalH047:=0;
TotalEtabH.TotalH048:=0;
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 22/04/2003
Modifié le ... :   /  /
Description .. : Gestion de modification de la nature
Mots clefs ... : PAIE;PGDADSB
*****************************************************************}
procedure TOF_PG_DADS2_PREP.Parametrage(Sender: TObject);
begin
PGAnnee:= Validite.value;
PGExercice:= RechDom ('PGANNEE', PGAnnee, False);

DebExer:= StrToDate ('01/01/'+PGExercice);
FinExer:= StrToDate ('31/12/'+PGExercice);
end;

//PT3
{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 12/05/2004
Modifié le ... :   /  /
Description .. : Complément des raccourcis claviers
Mots clefs ... : PAIE;PGDADSB
*****************************************************************}
procedure TOF_PG_DADS2_PREP.FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
TFVierge(Ecran).FormKeyDown(Sender, Key, Shift);
case Key of
     VK_F6: if ((GetControlVisible('BValider')) and
               (GetControlEnabled('BValider'))) then
               TFVierge(Ecran).BValider.Click; //Calcul des éléments
     end;
end;
//FIN PT3

Initialization
registerclasses([TOF_PG_DADS2_PREP]) ;

end.
