{***********UNITE*************************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 28/09/2001
Modifié le ... :
Description .. : Unité de gestion du multicritère DADSU pour la construction
Suite ........ : du fichier
Mots clefs ... : PAIE,PGDADSU
*****************************************************************}
{
PT1   : 22/01/2002 VG V571 Ajout de PGExercice
PT2   : 26/04/2002 VG V571 Par défaut, année précédente jusqu'au mois de
                           septembre . A partir du mois d'octobre, année en
                           cours (pour tests)
PT3   : 03/06/2002 VG V582 Gestion de la DADSU BTP                           
PT4-1 : 29/01/2003 VG V591 Affichage d'un message si l'exercice par défaut
                           n'existe pas - FQ N°10469
PT4-2 : 12/02/2003 VG V_42 FQ N°10469 bis
PT5   : 12/03/2003 VG V_42 Ajout d'un message d'avertissement lors de la
                           génération du pré-fichier de type DADS-U complète
                           FQ N°10567                          
PT6   : 16/09/2003 VG V_42 Adaptation cahier des charges V7R01
PT7   : 31/03/2004 VG V_50 Ajout d'un bouton sur le multi-critère pour
                           visualiser le fichier de log
PT8   : 14/05/2004 VG V_50 Ajout du traitement lié aux raccourcis clavier                           
PT9   : 05/07/2004 VG V_50 Adaptation cahier des charges V8R00
PT10  : 07/10/2005 VG V_60 Adaptation cahier des charges V8R02
PT11-1: 17/10/2006 VG V_70 Traitement DADS-U complémentaire
PT11-2: 17/10/2006 VG V_70 Suppression du fichier de contrôle - mise en table
                           des erreurs
PT11-3: 17/10/2006 VG V_70 Gestion du décalage de paie - FQ N°12860
PT11-4: 17/10/2006 VG V_70 Adaptation cahier des charges DADS-U V8R04
PT12  : 06/11/2006 VG V_70 Correction gestion du décalage
PT13  : 11/01/2007 VG V_72 Suppression des enregistrements ayant le champ
                           exercicedads mal alimenté - FQ N°13827
PT14  : 06/02/2007 FC V_72 Mise en place filtrage des habilitations/populations
PT15  : 29/06/2007 VG V_72 Fichiers en base
PT16  : 05/11/2007 VG V_80 Adaptation cahier des charges V8R06
PT17  : 28/11/2007 VG V_80 Gestion du champ ET_FICTIF - FQ N°13925
PT18  : 03/12/2007 VG V_80 Prise en compte du trimestre civil dans le cas d'une
                           déclaration trimestrielle - FQ N°13245
}
unit UTofPG_MulDADSFichier;

interface
uses UTOF,
     HCtrls,
     Hqry,
     HTB97,
     paramsoc,
     ShellAPI,
     windows,
{$IFNDEF EAGLCLIENT}
     {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
     Mul,
     FE_Main,
{$ELSE}
     MaineAgl,
     emul,
{$ENDIF}
     UTob,
     HEnt1,
     Pgoutils2,
     classes,
     PgDADSCommun,
     sysutils,
     EntPaie,
     hmsgbox,
     Controls,
     PgDADSOutils,
     P5Def, //PT14
     StdCtrls;
     {,PGVisuObjet;}

Type
     TOF_PGMULDADSFICHIER= Class (TOF)
       public
       procedure OnArgument(Arguments : String ) ; override ;
       procedure OnUpdate ; override;

       private
       WW : THEdit;     // Clause XX_WHERE
       Car, Exercice, Nature : THValComboBox;
       Etablissement : THMultiValComboBox;
       Q_Mul:THQuery ;  // Query pour changer la liste associee
       SansFichierBase : string;

       procedure ActiveWhere (Sender: TObject);
       procedure DateChange (Sender: Tobject);
       procedure CarChange(Sender: TObject);
       procedure NatureChange (Sender: Tobject);
       procedure EtabChange (Sender: TObject);
       procedure Parametrage();
       procedure Lancement (Sender: TObject);
       procedure ChargeZones ();
       procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
       procedure DateDecaleChange(Sender: TObject);
     END ;

implementation

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 02/10/2001
Modifié le ... :   /  /
Description .. : XX_WHERE
Mots clefs ... : PAIE;PGDADSU
*****************************************************************}
procedure TOF_PGMULDADSFICHIER.ActiveWhere(Sender: TObject);
var
Habilitation, StDateFin : String;  //PT14
begin
FinDAnnee:= StrToDate ('31/12/'+Copy (DateToStr (FinExer), 7, 4));  //PT16
if (WW <> NIL) then
   begin
   if Q_Mul <> NIL then
      TFMul(Ecran).SetDBListe('PGDADSPERIODE');

//DEB PT14
   Habilitation:= '';
   if Assigned (MonHabilitation) and (MonHabilitation.LeSQL<>'') then
      Habilitation:= ' AND PDE_SALARIE IN (SELECT PSA_SALARIE FROM SALARIES'+
                     ' WHERE '+MonHabilitation.LeSQL+')';
//FIN PT14

{PT16
   WW.Text:= ' PDE_TYPE="'+TypeD+'" AND'+
             ' PDE_DATEDEBUT<="'+UsDateTime(FinExer)+'" AND'+
             ' PDE_DATEFIN>="'+UsDateTime(DebExer)+'" AND'+
             ' PDE_EXERCICEDADS="'+PGExercice+'"' + Habilitation;
}
   if ((Nature.Value>='0100') and (Nature.Value<'0300')) then
      StDateFin:= ' PDE_DATEDEBUT<="'+UsDateTime(FinDAnnee)+'" AND'
   else
      StDateFin:= ' PDE_DATEDEBUT<="'+UsDateTime(FinExer)+'" AND';
   WW.Text:= ' PDE_TYPE="'+TypeD+'" AND'+StDateFin+
             ' PDE_EXERCICEDADS="'+PGExercice+'"'+Habilitation;
//FIN PT16
   TFMul(Ecran).BOuvrir.Enabled:= False;
   end;
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 02/10/2001
Modifié le ... :   /  /    
Description .. : OnArgument
Mots clefs ... : PAIE;PGDADSU
*****************************************************************}
procedure TOF_PGMULDADSFICHIER.OnArgument(Arguments: String);
var
ModifDateDecale : TToolbarButton97;
Etab, Mes, StEtab, StPcl, StPlus : string;
TEtabDADS, TEtabDADSD : Tob;
begin
inherited ;
TFMul (Ecran).BOuvrir.Enabled:= False;
Exercice:= THValComboBox (GetControl ('CEXERCICE'));
Nature:= THValComboBox (GetControl ('CNATURE'));
Etablissement:= THMultiValComboBox (GetControl ('PDS_DONNEEAFFICH'));        //PT17
SetControlText ('PSA_DADSFRACTION', '1');
Car:= THValComboBox (GetControl ('CCAR'));

if (Exercice<>NIL) then
   Exercice.OnChange:= DateChange;

if (Nature<>nil) then
   Nature.OnChange:=  NatureChange;

if (Car<>nil) then
   begin
   Car.Value:= 'A00';
   Car.OnChange:= CarChange;
   end;

//PT17
if (Etablissement<>nil) then
   Etablissement.OnChange:= EtabChange;
//FIN PT17
   
//PT11-1
if (Arguments='C') then
   SetControlChecked ('CHCOMPL', True);   
//FIN PT11-1

ChargeZones;

WW:= THEdit (GetControl ('XX_WHERE'));
ActiveWhere (NIL);

TFMul (Ecran).BOuvrir.OnClick:= Lancement;

TFMul (Ecran).OnKeyDown:= FormKeyDown;

{PT11-2
Visualiser := TToolbarButton97(GetControl('BVISUALISER'));
if Visualiser <> NIL then
   Visualiser.OnClick := VisuClick;
}   

//PT11-3
// Gestion du bouton de modification des dates
ModifDateDecale:= TToolbarButton97 (GetControl ('BMODIFDATE'));
if ModifDateDecale<>nil then
   ModifDateDecale.OnClick:= DateDecaleChange;
//FIN PT11-3

MajExercice;                 //PT13
//PT15
{Recherche fichier <> CEG  et = DAD pour savoir si client utilise fichier en
base pour le traitement des DADS-U}
if (V_PGI.ModePcl = '1') then
   StPcl:= ' AND NOT (YFS_PREDEFINI="DOS" AND'+
           ' YFS_NODOSSIER<>"'+V_PGI.NoDossier+'")'
else
   StPcl:= '';

If not existeSQL ('SELECT YFS_NOM'+
                  ' FROM YFILESTD WHERE'+
                  ' YFS_CODEPRODUIT="PAIE" AND'+
                  ' YFS_CRIT1="DAD" AND'+
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
//FIN PT15

//PT17
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
            StEtab:= StEtab+Etab;
            TEtabDADSD:= TEtabDADS.FindNext ([''], [''], False);
            if (TEtabDADSD<>nil) then
               Etab:= TEtabDADSD.GetValue ('ET_ETABLISSEMENT')
            else
               Etab:='';
            if (Etab<>'') then
               StEtab:= StEtab+';';
            end;
      SetControlText ('PDS_DONNEEAFFICH', StEtab);
      end;
   end;
FreeAndNil (TEtabDADS);

StPlus:= ' WHERE ET_FICTIF<>"X"';
SetControlProperty ('PDS_DONNEEAFFICH', 'Plus', StPlus);
//FIN PT17
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 13/02/2002
Modifié le ... :   /  /
Description .. : OnUpdate
Mots clefs ... : PAIE;PGDADSU
*****************************************************************}
procedure TOF_PGMULDADSFICHIER.OnUpdate;
begin
inherited ;
{PT12
Parametrage;
}
TFMul(Ecran).BOuvrir.Enabled:= True;
//PT4-2
if (PGAnnee = '') then
   TFMul(Ecran).BOuvrir.Enabled:= False;
//FIN PT4-2

end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 02/10/2001
Modifié le ... :   /  /
Description .. : Gestion de modification de date
Mots clefs ... : PAIE;PGDADSU
*****************************************************************}
procedure TOF_PGMULDADSFICHIER.DateChange(Sender: TObject);
begin
Parametrage;
ActiveWhere (Sender);
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 03/10/2001
Modifié le ... :   /  /
Description .. : Gestion de modification de la nature
Mots clefs ... : PAIE;PGDADSU
*****************************************************************}
procedure TOF_PGMULDADSFICHIER.NatureChange(Sender: TObject);
Var
Nat2 : string;
begin
{PT11-1
if (Car <> nil) then
   begin
   if ((Nature.Value>='0400') and (Nature.Value<'0500')) then
}
Nat2:= Copy(GetControlText ('CNATURE'),1,2);
if ((Nat2='04') or (Nat2='09')) then
//FIN PT11-1
   begin
   if (VH_Paie.PGBTP = False) then
      begin
      PGIBox('Vous ne gérez pas le module BTP', 'DADS-U standard');
{PT11-1
         Nature.Value := '0151';
}
      SetControlText ('CNATURE', '0151');
      end
   else
      begin
{PT11-1
      Car.Enabled := True;
      Car.Visible := True;
}
      SetControlEnabled('CCAR', True);
      SetControlVisible('CCAR', True);
//FIN PT11-1
      SetControlEnabled('LCAR', True);
      SetControlVisible('LCAR', True);
      end;
   end
else
   begin
{PT11-3
   Car.Enabled := False;
   Car.Visible := False;
   Car.Value := 'A00';
}
   SetControlEnabled('CCAR', False);
   SetControlVisible('CCAR', False);
   SetControlText ('CCAR', 'A00');
{PT12
   SetControlText ('L_DDU', '01/01/'+RechDom ('PGANNEESOCIALE',
                                              GetControlText ('CEXERCICE'), FALSE));
   SetControlText ('L_DAU', '31/12/'+RechDom ('PGANNEESOCIALE',
                                              GetControlText ('CEXERCICE'), FALSE));
}
//FIN PT11-3
   SetControlEnabled('LCAR', False);
   SetControlVisible('LCAR', False);
   end;

Parametrage;
ActiveWhere (Sender);
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 13/02/2002
Modifié le ... :   /  /
Description .. : Gestion de modification de la caractéristique
Mots clefs ... : PAIE;PGDADSU
*****************************************************************}
procedure TOF_PGMULDADSFICHIER.CarChange(Sender: TObject);
begin
Parametrage;
ActiveWhere (Sender);
end;


//PT17
{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 28/11/2007
Modifié le ... :   /  /
Description .. : Gestion de modification de l'établissement
Mots clefs ... : PAIE;PGDADSU
*****************************************************************}
procedure TOF_PGMULDADSFICHIER.EtabChange(Sender: TObject);
Var
Etab, StEtab : string;
TEtabDADS, TEtabDADSD : Tob;
begin
if ((GetControlText ('PDS_DONNEEAFFICH')='') or
   (GetControlText ('PDS_DONNEEAFFICH')='<<Tous>>')) then
   begin
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
               StEtab:= StEtab+Etab;
               TEtabDADSD:= TEtabDADS.FindNext ([''], [''], False);
               if (TEtabDADSD<>nil) then
                  Etab:= TEtabDADSD.GetValue ('ET_ETABLISSEMENT')
               else
                  Etab:='';
               if (Etab<>'') then
                  StEtab:= StEtab+';';
               end;
         SetControlText ('PDS_DONNEEAFFICH', StEtab);
         end;
      end;
   FreeAndNil (TEtabDADS);
   end;
end;
//FIN PT17


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 13/02/2002
Modifié le ... :   /  /
Description .. : Gestion de modification de la nature
Mots clefs ... : PAIE;PGDADSU
*****************************************************************}
procedure TOF_PGMULDADSFICHIER.Parametrage();
var
StExer : string;
QExer : TQuery;
begin
PGAnnee := Exercice.value;
StExer:= 'SELECT PEX_DATEDEBUT, PEX_DATEFIN, PEX_ANNEEREFER'+
         ' FROM EXERSOCIAL WHERE'+
         ' PEX_EXERCICE="'+PGAnnee+'"';
QExer:=OpenSQL(StExer,TRUE) ;
if (NOT QExer.EOF) then
   PGExercice := QExer.FindField ('PEX_ANNEEREFER').AsString
else
   PGExercice := '';

if ((Nature.Value>='0400') and (Nature.Value<'0500')) then
   begin
   if ((Car.Value>='M00') and (Car.Value <='M99')) then
      begin
      TypeD := '005';
      if (Car.Value = 'M01') then
         begin
         DebExer := StrToDate('01/01/'+PGExercice);
         FinExer := StrToDate('31/01/'+PGExercice);
         end
      else
      if (Car.Value = 'M02') then
         begin
         DebExer := StrToDate('01/02/'+PGExercice);
         FinExer := FinDeMois(StrToDate('28/02/'+PGExercice));
         end
      else
      if (Car.Value = 'M03') then
         begin
         DebExer := StrToDate('01/03/'+PGExercice);
         FinExer := StrToDate('31/03/'+PGExercice);
         end
      else
      if (Car.Value = 'M04') then
         begin
         DebExer := StrToDate('01/04/'+IntToStr(StrToInt(PGExercice)-1));
         FinExer := StrToDate('30/04/'+IntToStr(StrToInt(PGExercice)-1));
         end
      else
      if (Car.Value = 'M05') then
         begin
         DebExer := StrToDate('01/05/'+IntToStr(StrToInt(PGExercice)-1));
         FinExer := StrToDate('31/05/'+IntToStr(StrToInt(PGExercice)-1));
         end
      else
      if (Car.Value = 'M06') then
         begin
         DebExer := StrToDate('01/06/'+IntToStr(StrToInt(PGExercice)-1));
         FinExer := StrToDate('30/06/'+IntToStr(StrToInt(PGExercice)-1));
         end
      else
      if (Car.Value = 'M07') then
         begin
         DebExer := StrToDate('01/07/'+IntToStr(StrToInt(PGExercice)-1));
         FinExer := StrToDate('31/07/'+IntToStr(StrToInt(PGExercice)-1));
         end
      else
      if (Car.Value = 'M08') then
         begin
         DebExer := StrToDate('01/08/'+IntToStr(StrToInt(PGExercice)-1));
         FinExer := StrToDate('31/08/'+IntToStr(StrToInt(PGExercice)-1));
         end
      else
      if (Car.Value = 'M09') then
         begin
         DebExer := StrToDate('01/09/'+IntToStr(StrToInt(PGExercice)-1));
         FinExer := StrToDate('30/09/'+IntToStr(StrToInt(PGExercice)-1));
         end
      else
      if (Car.Value = 'M10') then
         begin
         DebExer := StrToDate('01/10/'+IntToStr(StrToInt(PGExercice)-1));
         FinExer := StrToDate('31/10/'+IntToStr(StrToInt(PGExercice)-1));
         end
      else
      if (Car.Value = 'M11') then
         begin
         DebExer := StrToDate('01/11/'+IntToStr(StrToInt(PGExercice)-1));
         FinExer := StrToDate('30/11/'+IntToStr(StrToInt(PGExercice)-1));
         end
      else
      if (Car.Value = 'M12') then
         begin
         DebExer := StrToDate('01/12/'+IntToStr(StrToInt(PGExercice)-1));
         FinExer := StrToDate('31/12/'+IntToStr(StrToInt(PGExercice)-1));
         end;
      end;
   if ((Car.Value>='T00') and (Car.Value <='T99')) then
      begin
      TypeD := '004';
      if (Car.Value = 'T01') then
         begin
{PT18
         DebExer := StrToDate('01/04/'+IntToStr(StrToInt(PGExercice)-1));
         FinExer := StrToDate('30/06/'+IntToStr(StrToInt(PGExercice)-1));
}
         DebExer := StrToDate('01/01/'+PGExercice);
         FinExer := StrToDate('31/03/'+PGExercice);
         end
      else
      if (Car.Value = 'T02') then
         begin
{
         DebExer := StrToDate('01/07/'+IntToStr(StrToInt(PGExercice)-1));
         FinExer := StrToDate('30/09/'+IntToStr(StrToInt(PGExercice)-1));
}
         DebExer := StrToDate('01/04/'+IntToStr(StrToInt(PGExercice)-1));
         FinExer := StrToDate('30/06/'+IntToStr(StrToInt(PGExercice)-1));
         end
      else
      if (Car.Value = 'T03') then
         begin
{
         DebExer := StrToDate('01/10/'+IntToStr(StrToInt(PGExercice)-1));
         FinExer := StrToDate('31/12/'+IntToStr(StrToInt(PGExercice)-1));
}
         DebExer := StrToDate('01/07/'+IntToStr(StrToInt(PGExercice)-1));
         FinExer := StrToDate('30/09/'+IntToStr(StrToInt(PGExercice)-1));
         end
      else
      if (Car.Value = 'T04') then
         begin
{
         DebExer := StrToDate('01/01/'+PGExercice);
         FinExer := StrToDate('31/03/'+PGExercice);
}
         DebExer := StrToDate('01/10/'+IntToStr(StrToInt(PGExercice)-1));
         FinExer := StrToDate('31/12/'+IntToStr(StrToInt(PGExercice)-1));
//FIN PT18         
         end;
      end;
   if ((Car.Value>='S00') and (Car.Value <='S99')) then
      begin
{PT11-4
      TypeD := '003';
      if (Car.Value = 'S01') then
         begin
         DebExer := StrToDate('01/04/'+IntToStr(StrToInt(PGExercice)-1));
         FinExer := StrToDate('30/09/'+IntToStr(StrToInt(PGExercice)-1));
         end
      else
      if (Car.Value = 'S02') then
         begin
         DebExer := StrToDate('01/10/'+IntToStr(StrToInt(PGExercice)-1));
         FinExer := StrToDate('31/03/'+PGExercice);
         end;
}
      PGIBox ('Périodicité interdite', TFMul(Ecran).Caption);
      Car.Value:= 'A00';
//FIN PT11-4
      end;
   if Car.Value='A00' then
      begin
      TypeD := '002';
      DebExer := StrToDate('01/04/'+IntToStr(StrToInt(PGExercice)-1));
      FinExer := StrToDate('31/03/'+PGExercice);
      end;
   end
else
   begin
   TypeD := '001';
   DebExer := QExer.FindField ('PEX_DATEDEBUT').AsDateTime;
   FinExer := QExer.FindField ('PEX_DATEFIN').AsDateTime;
   end;
Ferme(QExer);

//PT11-1
if (GetControlText ('CHCOMPL')='X') then
   TypeD:= '2'+Copy (TypeD, 2, 2);
//FIN PT11-1

//PT12
SetControlText ('L_DDU', DateToStr (DebExer));
SetControlText ('L_DAU', DateToStr (FinExer));
//FIN PT12
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 02/10/2001
Modifié le ... :   /  /
Description .. : Lancement de la création de fichier
Mots clefs ... : PAIE;PGDADSU
*****************************************************************}
procedure TOF_PGMULDADSFICHIER.Lancement(Sender: TObject);
var
reponse : integer;
begin
reponse:= mrYes;

if ((VH_Paie.PGPCS2003=FALSE) and (PGExercice ='2003')) then
   PGIBox('Vous ne devez pas envoyer la DADS-U 2003 sans avoir réaffecté#13#10'+
          'les codes PCS. Pour cela, choisissez le menu#13#10'+
          '"Paramètres\Calcul de la Paie\Gestion des codes PCS"',
          Ecran.Caption);

if (Nature.Value='0151') then
   reponse:= PGIAsk('Attention, les fichiers de type "DADS-U Complète"#13#10'+
                    ' sont destinés à être envoyés par le biais d''un#13#10'+
                    ' portail Internet pour l''ensemble des organismes'+
                    ' sociaux.#13#10'+
                    ' Ce type de fichier ne doit pas être utilisé pour#13#10'+
                    ' un envoi à une caisse unique. Exemple : T.D.S.#13#10'+
                    ' Voulez-vous continuer ?',Ecran.Caption);

if (reponse=mrNo) then
   exit;

//PT11-2
reponse:= PGIAsk ('Avez-vous édité l''état "Contrôle DADS-U" ?#13#10'+
                  'Cet état recense les anomalies à corriger. Si vous#13#10'+
                  'ne procédez pas à la correction des anomalies, votre#13#10'+
                  'fichier DADS-U risque d''être refusé',Ecran.Caption);

if (reponse=mrNo) then
   begin
   reponse:= PGIAsk ('Voulez-vous continuer ?',Ecran.Caption);

   if (reponse=mrNo) then
      exit;
   end;
//FIN PT11-2

{PT11-1
AglLanceFiche ('PAY','DADS_ENTREPRISE', '', '' ,
              Exercice.Value+'|'+GetControlText('PDS_DONNEEAFFICH')+'|'+
              GetControlText('L_DDU')+'|'+GetControlText('L_DAU')+'|'+
              Nature.Value+'|'+GetControlText('PSA_DADSFRACTION')+'|'+
              Car.Value);
}
{PT15
AglLanceFiche ('PAY','DADS_ENTREPRISE', '', '' ,
              Exercice.Value+'|'+GetControlText('PDS_DONNEEAFFICH')+'|'+
              GetControlText('L_DDU')+'|'+GetControlText('L_DAU')+'|'+
              Nature.Value+'|'+GetControlText('PSA_DADSFRACTION')+'|'+
              Car.Value+'|'+GetControlText('CHCOMPL'));
}
if (GetCheckBoxState ('SANSFICHIERBASE')=CbChecked) then
// dépôt des fichiers sur répertoire
   SansFichierBase:= 'X'
else
// dépôt des fichiers en base
   SansFichierBase:= '-';

AglLanceFiche ('PAY','DADS_ENTREPRISE', '', '' ,
              Exercice.Value+'|'+GetControlText('PDS_DONNEEAFFICH')+'|'+
              GetControlText('L_DDU')+'|'+GetControlText('L_DAU')+'|'+
              Nature.Value+'|'+GetControlText('PSA_DADSFRACTION')+'|'+
              Car.Value+'|'+GetControlText('CHCOMPL')+'|'+SansFichierBase);
//FIN PT15
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 03/10/2001
Modifié le ... :   /  /
Description .. : Chargement des éléments de la fiche
Mots clefs ... : PAIE,PGDADSU
*****************************************************************}
procedure TOF_PGMULDADSFICHIER.ChargeZones();
var
AnneeE, AnneePrec, MoisE, ComboExer : string;
JourJ : TDateTime;
AnneeA, Jour, MoisM : Word;
begin
JourJ := Date;
DecodeDate(JourJ, AnneeA, MoisM, Jour);
if MoisM>9 then
   AnneePrec := IntToStr(AnneeA)
else
   AnneePrec := IntToStr(AnneeA-1);

if RendExerSocialPrec (MoisE, AnneeE, ComboExer, DebExer, FinExer, AnneePrec) = TRUE then
   begin
   if Exercice <> NIL then
      begin
      Exercice.Value := ComboExer;
      PGExercice := AnneeE;
      end;
   end
else
   begin
   PGIBox('L''exercice '+AnneePrec+' n''existe pas', Ecran.Caption);
   end;

if Exercice <> nil then
   PGAnnee := Exercice.value;

if (Nature <> nil) then
   Nature.Value := '0151';

if Car <> nil then
   begin
   Car.Enabled := False;
   Car.Value := 'A00';
   Car.Visible := False;
   end;

SetControlEnabled('LCAR', False);
SetControlVisible('LCAR', False);
{PT10
SetControlEnabled('CANNEE', False);
}
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 31/03/2004
Modifié le ... :   /  /
Description .. : Procédure exécutée lors du click sur le bouton "Visualiser"
Mots clefs ... : PAIE,PGDADSU
*****************************************************************}
(*PT11-2
procedure TOF_PGMULDADSFICHIER.VisuClick(Sender: TObject);
var
NomFic, St : String;
begin
ForceNumerique(GetParamSoc('SO_SIRET'), st);
if ControlSiret(st)=False then
   PGIAsk('Le SIRET de la société n''est pas valide.#13#10'+
          'Vous devez le vérifier en y accédant par le module#13#10'+
          'Paramètres/menu Comptabilité/commande Paramètres comptables/Coordonnées.#13#10'+
          'Si vous travaillez en environnement multi-dossiers,#13#10'+
          'vous pouvez y accéder par le Bureau PGI/Annuaire',
          'Visualisation du contrôle de la génération');

{$IFDEF EAGLCLIENT}
NomFic := VH_Paie.PgCheminEagl+'\'+st+'_DADSU_PGI_PREP.log';
{$ELSE}
NomFic := V_PGI.DatPath+'\'+st+'_DADSU_PGI_PREP.log';
{$ENDIF}
ShellExecute( 0, PCHAR('open'),PChar('WordPad'), PChar(NomFic),Nil,SW_RESTORE);
end;
*)

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 12/05/2004
Modifié le ... :   /  /
Description .. : Complément des raccourcis claviers
Mots clefs ... : PAIE;PGDADSU
*****************************************************************}
procedure TOF_PGMULDADSFICHIER.FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
TFMul(Ecran).FormKeyDown(Sender, Key, Shift);
case Key of
     VK_F6: if ((TFMul(Ecran).BOuvrir.Visible) and
               (TFMul(Ecran).BOuvrir.Enabled)) then
               TFMul(Ecran).BOuvrir.Click; //Lancement de la préparation
{PT11-2
     ord('I'): if ((GetControlVisible('BVISUALISER')) and
                  (GetControlEnabled('BVISUALISER')) and (ssCtrl in Shift)) then
                  Visualiser.Click; //Visualisation des éléments
}                  
     end;
end;

//PT11-3
{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 24/03/2006
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_PGMULDADSFICHIER.DateDecaleChange(Sender: TObject);
begin
AglLanceFiche ('PAY', 'DADS_DATE', '', '', '');
SetControlText ('L_DDU', DateToStr(DebExer));
SetControlText ('L_DAU', DateToStr(FinExer));
ActiveWhere (Sender);    //PT12
end;
//FIN PT11-3

Initialization
registerclasses([TOF_PGMULDADSFICHIER]);
end.
