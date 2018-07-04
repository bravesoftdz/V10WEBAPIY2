{***********UNITE*************************************************
Auteur  ...... : PAIE - MF
Créé le ...... : 28/05/2002
Modifié le ... :   /  /
Description .. : Confection des fichier Ducs Edi dossier
Mots clefs ... : PAIE , PGDUCSEDI
*****************************************************************}
{ PT1 : 11/10/2002 : V585  MF  Correction de la requête de
                               la fct Activewhere pour tenir compte du n°
                               d'établissement
  PT2 : 23/10/2002 : V585  MF   fiche qualité 10252 : le fichier .log est
                                maintenant nommé DUCSEDI+Siret.log
  PT3 : 06/01/2003 : V591  MF
                                1-Rectification du paiement pour les caisses IRC
                                (ne traitait le paiement que si paiement groupé
                                ou ducs dossier et non paiement groupé)
                                2-contrôles lors de la génération du fichier
                                absence de RIB
                                3-Initialisation enregistrement EnregDucsEdi
  PT4 : 04/03/2003 : V42   MF
                                1- initialisation du n° de centre payeur à 1
  PT5 : 10/03/2003 : V42   MF
                                1- IRC : segment DLI gp3 (n° de ligne): Il ne doit
                             pas y avoi de rupture de numérotation entre les
                             centres payeurs.
  PT6 : 21/07/2003 : V_421 MF
                                1- La notion d'établissement pricipal n'est
                                plus utilisée
                                2- Traitement du Centre Payeur (Caisses IRC)
  PT7 : 12/09/2003 : V_421 MF
                                1-Suite tests sur soc 614 : correction du traitement
                                de Caisse destinataire (une caisse destinataire,
                                non ducs dossier, sans regroupement apparait
                                dans la liste)
                                2-pgibox remplacé par piginfo (fin de traitement)
  PT8 : 16/09/2003 : V_421 MF
                                FQ 10788 : traitement des raccoucis clavier
                                sur la saisie des dates de période..
  PT9 : 18/09/2003 : V_421 MF
                                1-PgiBox remplace Message alerte
  PT10 : 20/04/2004 : V_50 MF   Ajout d'un contrôle concernant l'existence du
                                répertoire de stockage (CheminEagl).
                                L'ensemble du traitement est abandoné si ce
                                chemin n'existe pas
  PT11 : 09/02/2006 : V_65 MF   DUCS EDI 4.2
  PT12 : 07/07/2006 : V_70 MF   DUCS EDI 4.2 : modif traitement suite tests
                                Ducs Dossier en VLU
  PT13 : 29/01/2007 : V_70 MF   Modifs Ducs Edi V4.2
  PT14 : 09/02/2007 : V_702 FC Suivant un paramètre société, contrer l'habilitation qui est faite automatique
                            en effaçant le contenu du champ XX_WHERE2
  PT15 : 11/04/2007 : V_702 MF  Modifs Mise en base des fichiers Ducs EDI
  PT16 : 25/10/2007 : V_80  MF  Ne pas rendre obligatoire la coche paiement groupé pour les IRC
}
unit UtofPg_MulDucsEdi;

interface
uses

{$IFDEF EAGLCLIENT}
    eMul,
{$ELSE}
    HDB,
    Mul,
{$ENDIF}
    UTOF,classes,ed_tools,
    hmsgbox,HStatus,sysutils,
    PGDucsEdiOutils,
    HCtrls,Hqry,UTOB,HEnt1,
    StdCtrls, // PT15
    ParamSoc//PT14
    ;
Type
     TOF_PGMULDUCSEDI= Class (TOF)


      public
       procedure OnArgument(Arguments : String ) ; override ;
       procedure OnUpdate ; override;
       procedure OnLoad;  override ;

      private
       WW : THEdit;
// PT6-1       EtabPrinc : string;
       DateD,DateD_ : THEdit;   // PT8
       Q_Mul:THQuery;
       procedure ActiveWhere(Okok : Boolean);
       procedure LanceConfection (Sender: TObject);
       procedure RecupCleDucs(var EnregDucsEdi : TDucsEdi);


     END ;
type EConvertError = class(Exception);      // PT8
implementation

{***********A.G.L.Privé.*****************************************
Auteur  ...... : PAIE - MF
Créé le ...... : 28/05/2002
Modifié le ... : 28/05/2002
Description .. : On Argument
Suite ........ : Récupération de l'établissement par défaut.
Suite ........ : Initialisation bouton (mouette verte)
Mots clefs ... : PAIE, PGDUCSEDI
*****************************************************************}
procedure TOF_PGMULDUCSEDI.OnArgument(Arguments: String);
var
  StPcl : string;    // PT15
  Mes   : string;    // PT15

begin
inherited ;

  TFMul(Ecran).BOuvrir.Enabled := False;
  TFMul(Ecran).BOuvrir.OnClick := LanceConfection;

  //PT6-1 EtabPrinc := GetParamSoc('SO_ETABLISDEFAUT');
// d  PT8
  DateD:= THEdit(getcontrol('XX_VARIABLED'));
  DateD_:= THEdit(getcontrol('XX_VARIABLED_'));
// f PT8

  WW:=THEdit (GetControl ('XX_WHERE'));
// d PT15
// Recherche fichier <> CEG  et = DUC pour savoir si client utilise fichier en base
// pour le traitement des Ducs EDI
    if (V_PGI.ModePcl = '1') then
       StPcl:= ' AND NOT (YFS_PREDEFINI="DOS" AND YFS_NODOSSIER<>"'+V_PGI.NoDossier+'")'
    else
       StPcl:= '';
    If not existeSQL ('SELECT YFS_NOM'+
                      ' FROM YFILESTD WHERE'+
                      ' YFS_CODEPRODUIT="PAIE" AND'+
                      ' YFS_CRIT1="DUC" AND'+
                      ' YFS_PREDEFINI<>"CEG"'+StPcl) then
       begin
       Mes:= 'Cette version apporte un nouveau fonctionnement dans#13#10'+
             'la gestion du stockage des fichiers.#13#10#13#10'+
             'Par défaut l''utilisation reste identique.#13#10#13#10'+
             'Toutefois nous vous invitons à mettre en place cette nouvelle fonctionnalité';
       SetControlChecked ('SANSFICHIERBASE', true);
       HShowMessage ('1;Gestion des fichiers en base ;'+Mes+';E;HO;;;;;41910025', '', '');
       end;
// f PT15
end; {OnArgument}
{***********A.G.L.Privé.*****************************************
Auteur  ...... : PAIE - MF
Créé le ...... : 28/05/2002
Modifié le ... :   /  /
Description .. : OnLoad : Chargement de la fiche
Mots clefs ... : PAIE, PGDUCSEDI
*****************************************************************}
procedure TOF_PGMULDUCSEDI.OnLoad;
var
   Okok : Boolean;
begin
inherited ;
  Okok := TRUE;
  ActiveWhere (Okok);

end;  {fin OnLoad}

{***********A.G.L.Privé.*****************************************
Auteur  ...... : PAIE - MF
Créé le ...... : 28/05/2002
Modifié le ... :   /  /
Description .. : OnUpdate
Mots clefs ... : PAIE, PGDUCSEDI
*****************************************************************}
procedure TOF_PGMULDUCSEDI.OnUpdate;
begin
inherited ;
  TFMul(Ecran).BOuvrir.Enabled := True;
end; {OnUpdate}

{***********A.G.L.Privé.*****************************************
Auteur  ...... : PAIE - MF
Créé le ...... : 28/05/2002
Modifié le ... :   /  /
Description .. : Lancement de la confection des fichiers Ducs Edi pour les
Suite ........ : déclarations sélectionnées
Mots clefs ... : PAIE, PGDUCSEDI
*****************************************************************}
procedure TOF_PGMULDUCSEDI.LanceConfection(Sender: TObject);
var
{$IFDEF EAGLCLIENT}
   Liste : THGrid;
{$ELSE}
   Liste : THDBGrid;
{$ENDIF}
//   TProgress: TQRProgressForm ; PORTAGECWAS

//   Pages : TPageControl;
   i                            : integer;
   EnregDucsEdi                 : TDucsEdi;
   EnrDucsEdiGroupee            : TDucsEdiGroupee;
   Anomalie, dossOK, OuvLog     : boolean;
   SiretPrecedent               : string;
   TSiret, TSiretFille          : TOB;
   NoLg                         : integer; // PT5
// d  PT9  DUCS EDI 4.2
{$IFNDEF DUCS41}
   LesDucsGroupee               : TOB ;
   EffGlobal                    : boolean;
   TotHommes                    : double;
   TotFemmes                    : double;
   TMtsEtabs                    : TOB ; //PT13
{$ENDIF}
// f PT9  DUCS EDI 4.2
   SansFichierBase              : boolean; // PT15

begin
  Anomalie := False;
// d PT15
  if (GetCheckBoxState ('SANSFICHIERBASE')=CbChecked) then
    // dépôt des fichiers sur répertoire
    SansFichierBase := true
  else
    // dépôt des fichiers en base
    SansFichierBase := false;
// f PT15


  i := 0;
{$IFDEF EAGLCLIENT}
  Liste := THGrid(GetControl('FListe'));
{$ELSE}
  Liste := THDBGrid(GetControl('FListe'));
{$ENDIF}
  if Liste <> NIL then
    if (Liste.NbSelected=0) and (not Liste.AllSelected) then
    begin
        PgiBox('Aucun élément sélectionné',Ecran.Caption); // PT9-1
      exit;
    end;

  {Sélection des DUCS}
  if Liste <> NIL then
  begin
    SiretPrecedent := '';
    TSiret := TOB.Create ('SiretLog',NIL,-1);
    if (Liste.AllSelected=TRUE) then
    begin
      InitMoveProgressForm (NIL,'Confection de fichier en cours', 'Veuillez patienter SVP ...',i,FALSE,TRUE);
      InitMove(TFmul(Ecran).Q.RecordCount,'');
      TFmul(Ecran).Q.First;
      while not TFmul(Ecran).Q.EOF do
      begin
// d PT13
{$IFNDEF DUCS41}
        if (TMtsEtabs <> NIL) then
           FreeAndNil (TMtsEtabs);
{$ENDIF}
// f PT13
        RecupCleDucs(EnregDucsEdi);
        if (SiretPrecedent <> EnregDucsEdi.Siret) then
          OuvLog := True
        else
          OuvLog := False;
        if (OuvLog = True)  then
        begin
          if (SiretPrecedent <> '') then
            FermFicDucsEdiLOG();
          OuvFicDucsEdiLOG(EnregDucsEdi.Siret);
// d PT10
          if (EnregDucsEdi.Siret = '') then
            begin
              break;
            end;
// f PT10
          SiretPrecedent := EnregDucsEdi.Siret;
          TSiretFille := TOB.Create ('SiretTraites',TSiret,-1);
          TSiretFille.AddChampSupValeur('SIRET','',FALSE);
          TSiretFille.PutValue('SIRET',EnregDucsEdi.Siret);
          TSiretFille.AddChampSupValeur('ANOMALIE','',FALSE);
          TSiretFille.PutValue('ANOMALIE','-');
        end;

        InitVarDucsPDU(EnregDucsEdi);  {DUCSENTETE}
        InitVarDucsPOG(EnregDucsEdi);  {ORGANISMEPAIE}
        InitVarDucsPDD(EnregDucsEdi);  {DUCSDETAIL}
        InitVarDucsPET(EnregDucsEdi);  {EMETTEURSOCIAL}
        InitVarDucsET(EnregDucsEdi);   {ETABLISS}
// d PT12
{        if ((copy(EnregDucsEdi.TypDucs,1,1) = 'U') and
            (EnregDucsEdi.DucsDossier) and
            (EnregDucsEdi.PaiementGroupe)) then    }
// f PT12
{           InitEffectifsUNEDIC (EnregDucsEdi);} {UNEDIC Multi Etab - VLU}
        
        InitCotisation(EnregDucsEdi, EnrDucsEdiGroupee,'');  {Lignes détail}
        dossOK := true;
        if (ControlVarDucsEnTete(EnregDucsEdi) = False) then
          dossOK := false;
        if (ControlVarDucsgrp1(EnregDucsEdi) = False) then
          dossOK := false;
        if (ControlCotisation(EnregDucsEdi, EnrDucsEdiGroupee,'') = False) then
          dossOK := False;
        if (ControlPaiement(EnregDucsEdi) = False) then
          dossOK := False;
        if (dossOK = false) then
        begin
          Abandonfichier('');
          MoveCurProgressForm(EnregDucsEdi.Organisme);
          TFmul(Ecran).Q.Next;
          Anomalie := True;
          TSiretFille.PutValue('ANOMALIE','X');
          Continue;
        end;
        OuvreFichierDucsEdi(EnregDucsEdi);  
        EnTeteDeclaration(EnregDucsEdi);
// d PT9  DUCS EDI 4.2
{$IFNDEF DUCS41}
        EffGlobal := False;
        if ((copy(EnregDucsEdi.TypDucs,1,1) = 'A') and
            (EnregDucsEdi.DucsDossier = TRUE) and
            (EnregDucsEdi.PaiementGroupe = TRUE)) or
            ((copy(EnregDucsEdi.TypDucs,1,1) = 'I') and
            (EnregDucsEdi.DucsDossier = TRUE) {and
            (EnregDucsEdi.PaiementGroupe = TRUE)}) or //PT16
            ((copy(EnregDucsEdi.TypDucs,1,1) = 'U') and
            (EnregDucsEdi.DucsDossier = TRUE) and
            (EnregDucsEdi.PaiementGroupe = TRUE)) then
        {Multi-établissement : Traitement de plusieurs groupes 2,
         3, 4 et 5 selon le cas
         Acoss : ducs dossier - paiement groupé (pls GP2+GP3+GP4+GP5)
         Irc   : Multi centre payeur (pls GP2+GP3)
         Unedic : ducs dossier - paiement groupé (pls GP2 et 1 GP3)}
        begin
          TotHommes := EnregDucsEdi.TotHommes ;
          TotFemmes := EnregDucsEdi.TotFemmes ;
          MultiGp2Gp3TOB(EnregDucsEdi.Etab ,EnregDucsEdi,EnrDucsEdiGroupee,DossOK, NoLg); // 10/05 on réactive la ligne
          if (copy(EnregDucsEdi.TypDucs,1,1) = 'A') or
             (copy(EnregDucsEdi.TypDucs,1,1) = 'U') then

          LesDucsGroupee := TOB_DucsGroupee.FindFirst([''],[''],TRUE);
          While LesDucsGroupee <> NIL  do
          begin
              InitEffGlobal (EnregDucsEdi,EnrDucsEdiGroupee,LesDucsGroupee);  //10/05 on réactive la ligne
              EnregDucsEdi.TotHommes := EnregDucsEdi.TotHommes+EnrDucsEdiGroupee.TotHommes ;
              EnregDucsEdi.TotFemmes := EnregDucsEdi.TotFemmes+EnrDucsEdiGroupee.TotFemmes ;
              EffGlobal := true;
              LesDucsGroupee := TOB_DucsGroupee.FindNext ([''],[''],TRUE);
          end;
        end;
{$ENDIF}
// f PT9  DUCS EDI 4.2
        Groupe1Declaration(EnregDucsEdi);
// d PT9  DUCS EDI 4.2
{$IFNDEF DUCS41}
        if (EffGlobal) then
        begin
           EnregDucsEdi.TotHommes := TotHommes;
           EnregDucsEdi.TotFemmes := TotFemmes;
           EffGlobal := false;
        end;
{$ENDIF}
// f PT9  DUCS EDI 4.2
        Groupe2Declaration (EnregDucsEdi, EnrDucsEdiGroupee,'');

        if (copy(EnregDucsEdi.TypDucs,1,1) = 'A') then
        { ACOSS }
        begin
          Groupe3ACOSS (EnregDucsEdi, EnrDucsEdiGroupee,'');
          if (EnregDucsEdi.NbTransport > 1) then
            Groupe4_5ACOSS(EnregDucsEdi, EnrDucsEdiGroupee,'');
{??          if (EnregDucsEdi.DucsDossier = TRUE) and
             (EnregDucsEdi.PaiementGroupe = TRUE) then
          begin

          end;??}
        end;

        if (copy(EnregDucsEdi.TypDucs,1,1) = 'I') then
        { IRC }
// d PT5
          begin
            NoLg := 0;
// PT5      Groupe3IRC(EnregDucsEdi, EnrDucsEdiGroupee,'');
            Groupe3IRC(EnregDucsEdi, EnrDucsEdiGroupee,'', NoLg);
          end;
// f¨PT5

        if (copy(EnregDucsEdi.TypDucs,1,1) = 'U') then
        { UNEDIC }
          Groupe3UNEDIC(EnregDucsEdi);

        if ((copy(EnregDucsEdi.TypDucs,1,1) = 'A') and
            (EnregDucsEdi.DucsDossier = TRUE) and
            (EnregDucsEdi.PaiementGroupe = TRUE)) or
            ((copy(EnregDucsEdi.TypDucs,1,1) = 'I') and
            (EnregDucsEdi.DucsDossier = TRUE) {and
            (EnregDucsEdi.PaiementGroupe = TRUE)}) or //PT16
            ((copy(EnregDucsEdi.TypDucs,1,1) = 'U') and
            (EnregDucsEdi.DucsDossier = TRUE) and
            (EnregDucsEdi.PaiementGroupe = TRUE)) then
        {Multi-établissement : Traitement de plusieurs groupes 2,
         3, 4 et 5 selon le cas
         Acoss : ducs dossier - paiement groupé (pls GP2+GP3+GP4+GP5)
         Irc   : Multi centre payeur (pls GP2+GP3)
         Unedic : ducs dossier - paiement groupé (pls GP2 et 1 GP3)}
        begin
// d PT5
// PT5    MultiGp2Gp3(EnregDucsEdi.Etab ,EnregDucsEdi,EnrDucsEdiGroupee,DossOK);
{$IFNDEF DUCS41}
          MultiGp2Gp3(EnregDucsEdi.Etab ,EnregDucsEdi,EnrDucsEdiGroupee,DossOK, NoLg,TMtsEtabs);  //PT13
{$ELSE}
          MultiGp2Gp3(EnregDucsEdi.Etab ,EnregDucsEdi,EnrDucsEdiGroupee,DossOK, NoLg);  //PT13
{$ENDIF}
// f PT5          MultiGp2Gp3(EtabPrinc,EnregDucsEdi,EnrDucsEdiGroupee,DossOK);
          if (dossOK = False) then
          begin
            Abandonfichier('G');
            MoveCurProgressForm(EnregDucsEdi.Organisme);
            Anomalie := True;
            TSiretFille.PutValue('ANOMALIE','X');
            TFmul(Ecran).Q.Next;
            Continue;
          end;
// d PT3-1
        end;
{            if (copy(EnregDucsEdi.TypDucs,1,1) <>'I') or
               (EnregDucsEdi.PaiementGroupe = True) then }

// f PT3-1
//              begin
                FinDeclaration(EnregDucsEdi, EnrDucsEdiGroupee);
                if (EnregDucsEdi.ModePaiement <> '') then
// d PT3-2              PaiementEdi(EnregDucsEdi,EnrDucsEdiGroupee,'');
{$IFNDEF DUCS41}
                  PaiementEdi(EnregDucsEdi,EnrDucsEdiGroupee,'',DossOk,TMtsEtabs); //PT13
{$ELSE}
                  PaiementEdi(EnregDucsEdi,EnrDucsEdiGroupee,'',DossOk); //PT13
{$ENDIF}
                  if (dossOK = False) then
                    begin
                      Abandonfichier('G');
                      MoveCurProgressForm(EnregDucsEdi.Organisme);
                      TFmul(Ecran).Q.Next;
                      Anomalie := True;
                      TSiretFille.putValue('ANOMALIE','X');
                      Continue;
                      end;
// f PT3-2
//              end;
{$IFNDEF DUCS41}
            FinFin (EnregDucsEdi, TMtsEtabs, SansFichierBase); //PT13 PT15
            MajEnvoiSocialDUCS (EnregDucsEdi, SansFichierBase); // PT15
{$ELSE}
            FinFin (EnregDucsEdi); //PT13
            MajEnvoiSocialDUCS (EnregDucsEdi);
{$ENDIF}

            MoveCurProgressForm(EnregDucsEdi.Organisme);
            TFmul(Ecran).Q.Next;

            END;
      Liste.AllSelected:=False;
      FiniMove;
      FiniMoveProgressForm;
      TFMul(Ecran).bSelectAll.Down := False;

     end
    else
     begin
      InitMoveProgressForm (NIL,'Confection de fichier en cours', 'Veuillez patienter SVP ...',i,FALSE,TRUE);

      InitMove(Liste.NbSelected,'');
      for i:=0 to Liste.NbSelected-1 do
          BEGIN
          Liste.GotoLeBOOKMARK(i);
{$IFDEF EAGLCLIENT}
         TFmul(Ecran).Q.TQ.Seek(Liste.Row-1) ;
{$ELSE}
{$ENDIF}

          MoveCur(False);
// d PT13
{$IFNDEF DUCS41}
        if (TMtsEtabs <> NIL) then
           FreeAndNil (TMtsEtabs);
{$ENDIF}
// f PT13

          RecupCleDucs(EnregDucsEdi);
// d PT2
          if (SiretPrecedent <> EnregDucsEdi.Siret) then
            OuvLog := True
          else
            OuvLog := False;
          If (OuvLog = True) then
           begin
            if  (SiretPrecedent <> '') then
              FermFicDucsEdiLOG();
            OuvFicDucsEdiLOG(EnregDucsEdi.Siret);
// d PT10
            if (EnregDucsEdi.Siret = '') then
              begin
                break;
              end;
// f PT10
            SiretPrecedent := EnregDucsEdi.Siret;
            TSiretFille := TOB.Create ('SiretTraites',TSiret,-1);
            TSiretFille.AddChampSupValeur('SIRET','',FALSE);
            TSiretFille.PutValue('SIRET',EnregDucsEdi.Siret);
            TSiretFille.AddChampSupValeur('ANOMALIE','',FALSE);
            TSiretFille.PutValue('ANOMALIE','-');
           end;
// f PT2
          InitVarDucsPDU(EnregDucsEdi);  // DUCSENTETE
          InitVarDucsPOG(EnregDucsEdi);  // ORGANISMEPAIE
          InitVarDucsPDD(EnregDucsEdi);  // DUCSDETAIL
          InitVarDucsPET(EnregDucsEdi);  // EMETTEURSOCIAL
          InitVarDucsET(EnregDucsEdi);   // ETABLISS
// d PT12
{          if ((copy(EnregDucsEdi.TypDucs,1,1) = 'U') and
              (EnregDucsEdi.DucsDossier) and
              (EnregDucsEdi.PaiementGroupe)) then  }
// f PT12
{             InitEffectifsUNEDIC (EnregDucsEdi); // UNEDIC Multi Etab - VLU}
          InitCotisation(EnregDucsEdi, EnrDucsEdiGroupee,'');  // Lignes détail
          dossOK := true;
          if (ControlVarDucsEnTete(EnregDucsEdi) = False) then
             dossOK := False;
          if (ControlVarDucsgrp1(EnregDucsEdi) = False) then
             dossOK := False;
          if (ControlCotisation(EnregDucsEdi, EnrDucsEdiGroupee,'') = False) then
             dossOK := False;
          if (ControlPaiement(EnregDucsEdi) = False) then
             dossOK := False;
          if (dossOK = False) then
             begin
               Abandonfichier('');
               MoveCurProgressForm(EnregDucsEdi.Organisme);
               Anomalie := True;
               TSiretFille.PutValue('ANOMALIE','X');        // PT2
               Continue;
             end;

          OuvreFichierDucsEdi(EnregDucsEdi);
          EnTeteDeclaration(EnregDucsEdi);
// d PT9  DUCS EDI 4.2
{$IFNDEF DUCS41}
        EffGlobal := false;
        if ((copy(EnregDucsEdi.TypDucs,1,1) = 'A') and
            (EnregDucsEdi.DucsDossier = TRUE) and
            (EnregDucsEdi.PaiementGroupe = TRUE)) or
            ((copy(EnregDucsEdi.TypDucs,1,1) = 'I') and
            (EnregDucsEdi.DucsDossier = TRUE) {and
            (EnregDucsEdi.PaiementGroupe = TRUE)}) or //PT16
            ((copy(EnregDucsEdi.TypDucs,1,1) = 'U') and
            (EnregDucsEdi.DucsDossier = TRUE) and
            (EnregDucsEdi.PaiementGroupe = TRUE)) then
        {Multi-établissement : Traitement de plusieurs groupes 2,
         3, 4 et 5 selon le cas
         Acoss : ducs dossier - paiement groupé (pls GP2+GP3+GP4+GP5)
         Irc   : Multi centre payeur (pls GP2+GP3)
         Unedic : ducs dossier - paiement groupé (pls GP2 et 1 GP3)}
        begin
          TotHommes := EnregDucsEdi.TotHommes ;
          TotFemmes := EnregDucsEdi.TotFemmes ;
          MultiGp2Gp3TOB(EnregDucsEdi.Etab ,EnregDucsEdi,EnrDucsEdiGroupee,DossOK, NoLg);
          if (copy(EnregDucsEdi.TypDucs,1,1) = 'A') or
             (copy(EnregDucsEdi.TypDucs,1,1) = 'U') then

          LesDucsGroupee := TOB_DucsGroupee.FindFirst([''],[''],TRUE);
          While LesDucsGroupee <> NIL  do
          begin
              InitEffGlobal (EnregDucsEdi,EnrDucsEdiGroupee,LesDucsGroupee);
              EnregDucsEdi.TotHommes := EnregDucsEdi.TotHommes+EnrDucsEdiGroupee.TotHommes ;
              EnregDucsEdi.TotFemmes := EnregDucsEdi.TotFemmes+EnrDucsEdiGroupee.TotFemmes ;
              EffGlobal := true;
              LesDucsGroupee := TOB_DucsGroupee.FindNext ([''],[''],TRUE);
          end;
        end;
{$ENDIF}
// f PT9  DUCS EDI 4.2

          Groupe1Declaration(EnregDucsEdi);
// d PT9  DUCS EDI 4.2
{$IFNDEF DUCS41}
        if (EffGlobal) then
        begin
           EnregDucsEdi.TotHommes := TotHommes;
           EnregDucsEdi.TotFemmes := TotFemmes;
           EffGlobal := false;
        end;
{$ENDIF}
// f PT9  DUCS EDI 4.2

          Groupe2Declaration (EnregDucsEdi, EnrDucsEdiGroupee,'');
          if (copy(EnregDucsEdi.TypDucs,1,1) = 'A') then
            // ACOSS
            begin
               Groupe3ACOSS (EnregDucsEdi, EnrDucsEdiGroupee,'');
               if (EnregDucsEdi.NbTransport > 1) then
                 Groupe4_5ACOSS(EnregDucsEdi, EnrDucsEdiGroupee,'');
            end;
          if (copy(EnregDucsEdi.TypDucs,1,1) = 'I') then
            // IRC
// d PT5
          begin
            NoLg := 0;
//          Groupe3IRC(EnregDucsEdi, EnrDucsEdiGroupee,'');
            Groupe3IRC(EnregDucsEdi, EnrDucsEdiGroupee,'', NoLg);
          end;
// f PT5
          if (copy(EnregDucsEdi.TypDucs,1,1) = 'U') then        //A voir si utile
//PT9  DUCS EDI 4.2
{* pour ducs edi v4.1
          if ((EnregDucsEdi.DucsDossier = TRUE) and
               (EnregDucsEdi.PaiementGroupe = TRUE))  then
            begin
              InitCotisation (EnregDucsEdi, EnrDucsEdiGroupee,'G');
            end;*}
// PT9  DUCS EDI 4.2
            // UNEDIC
            Groupe3UNEDIC(EnregDucsEdi);

          if ((copy(EnregDucsEdi.TypDucs,1,1) = 'A') and
              (EnregDucsEdi.DucsDossier = TRUE) and
              (EnregDucsEdi.PaiementGroupe = TRUE)) or
             ((copy(EnregDucsEdi.TypDucs,1,1) = 'I') and
              (EnregDucsEdi.DucsDossier = TRUE) {and
              (EnregDucsEdi.PaiementGroupe = TRUE)}) or // PT16
             ((copy(EnregDucsEdi.TypDucs,1,1) = 'U') and
              (EnregDucsEdi.DucsDossier = TRUE) and
              (EnregDucsEdi.PaiementGroupe = TRUE)) then
            // Multi-établissement : Traitement de plusieurs groupes 2,
            //                       3, 4 et 5 selon le cas
            // Acoss : ducs dossier - paiement groupé (pls GP2+GP3+GP4+GP5)
            // Irc   : Multi centre payeur (pls GP2+GP3)
            // Unedic : ducs dossier - paiement groupé (pls GP2 et 1 GP3)
            begin
//PT6-1            MultiGp2Gp3(EtabPrinc,EnregDucsEdi,EnrDucsEdiGroupee,DossOK);
// PT5              MultiGp2Gp3(EnregDucsEdi.Etab,EnregDucsEdi,EnrDucsEdiGroupee,DossOK);
{$IFNDEF DUCS41}
              MultiGp2Gp3(EnregDucsEdi.Etab,EnregDucsEdi,EnrDucsEdiGroupee,DossOK, NoLg,TMtsEtabs);     //PT13
{$ELSE}
              MultiGp2Gp3(EnregDucsEdi.Etab,EnregDucsEdi,EnrDucsEdiGroupee,DossOK, NoLg);     //PT13
{$ENDIF}
              if (dossOK = False) then
                begin
                 Abandonfichier('G');
                 MoveCurProgressForm(EnregDucsEdi.Organisme);
                 Anomalie := True;
                 TSiretFille.putValue('ANOMALIE','X'); // PT2
                 Continue;
                end;
// d PT3-1
            end;
          if (copy(EnregDucsEdi.TypDucs,1,1) <> 'I') or
               (EnregDucsEdi.PaiementGroupe = True) or
               ((EnregDucsEdi.ducsDossier = False) and
                (EnregDucsEdi.PaiementGroupe = False))then
          begin
            FinDeclaration(EnregDucsEdi,EnrDucsEdiGroupee);
            if (EnregDucsEdi.ModePaiement <> '') then
// d PT3-2          PaiementEdi(EnregDucsEdi,EnrDucsEdiGroupee,'');
{$IFNDEF DUCS41}
              PaiementEdi(EnregDucsEdi,EnrDucsEdiGroupee,'',DossOK,TMtsEtabs); //PT13
{$ELSE}
              PaiementEdi(EnregDucsEdi,EnrDucsEdiGroupee,'',DossOK); //PT13
{$ENDIF}
              if (dossOK = False) then
                begin
                 Abandonfichier('G');
                 MoveCurProgressForm(EnregDucsEdi.Organisme);
                 TFmul(Ecran).Q.Next;
                 Anomalie := True;
                 TSiretFille.putValue('ANOMALIE','X');
                 Continue;
                end;
          end;
{$IFNDEF DUCS41}
          FinFin (EnregDucsEdi, TMtsEtabs, SansFichierBase);    //PT13 PT15
          MajEnvoiSocialDUCS (EnregDucsEdi, SansFichierBase); // PT15
{$ELSE}
          FinFin (EnregDucsEdi);    //PT13
          MajEnvoiSocialDUCS (EnregDucsEdi); 
{$ENDIF}

          MoveCurProgressForm(EnregDucsEdi.Organisme);
          END;

      Liste.ClearSelected;
      FiniMove;
      FiniMoveProgressForm;
    end;
      PgiInfo('Traitement terminé','Confection de fichiers Ducs-EDI');
//PT7-2     PGIBox('Traitement terminé','Confection de fichiers Ducs-EDI');
      if (Anomalie = True) then
        // Il y a des anomalies : confection du fichier abandonnée
        PGIBox('Attention : Il y a des anomalies','Confection de fichiers Ducs-EDI');

// d PT10
      if (EnregDucsEdi.Siret <> '') then
        begin
          FermFicDucsEdiLOG();
// d PT2
          EdiFicDucsEdiLOG(TSiret);
        end;
// f PT10

      if TSiret <> NIL then
       begin
        TSiret.Free;
       end;
// f PT2
   END;

end;
{***********A.G.L.Privé.*****************************************
Auteur  ...... : PAIE - MF
Créé le ...... : 28/05/2002
Modifié le ... :   /  /    
Description .. : Récupération des critères
Mots clefs ... : PAIE, PGDUCSEDI
*****************************************************************}
procedure TOF_PGMULDUCSEDI.RecupCleDucs(var EnregDucsEdi : TDucsEdi);
begin
// d PT3-3
  with EnregDucsEdi do
   begin
     Etab :='';
     Organisme :='';
     DateDebut := IDate1900;
     DateFin := IDate1900;
     NoDucs := 0;
     Siret :='';
     NatureOrg :='';
     TypDucs :='';
     Periode :='';
     // Infos DUCSENTETE
     Abrege :='';
     IdentQual :='';
     IdentEmet :='';
     IdentDest :='';
     DucsDossier :=False;
     Declarant :='';
     EmettSoc :='';
     TelephoneDecl :='';
     FaxDeclarant :='';
     DeclarantSuite :='';
     ApePdu :='';
     NbSalFpe := 0.0;
     TotHommes := 0.0;
     TotFemmes := 0.0;
     TotApprenti := 0.0;
     Paiement := IDate1900;
     Regularisation := 0.0;
     Acompte := 0.0;
     MtTotRegul := 0.0;
     MtTotAcpte := 0.0;
     NumeroPdu :='';
     MonnaieTenue :='';
     NumeroInterne :='';
     // Infos DUCSDETAIL
     MtDeclare  := 0.0;
     MtTransport := 0.0;
     MtTotal := 0.0;
     MtAPayer := 0.0;
     NbTransport := 0;
     NbCot := 0;
     CleIBAN :='';
     ErrTransport :=False;
     ErrCotisqual :=False;
     ErrInstitution :=False;
     ErrCondition :=False;
     Neant :=False;
     // Infos ORGANISMEPAIE
     SiretPog :='';
     LibellePog :='';
     Adresse1Pog :='';
     Adresse2Pog :='';
     VillePog :='';
     CpPog :='';
     ContactPog :='';
     TelPog :='';
     FaxPog :='';
     InstitutionPog :='';
     AdherContact :='';
     NoContEmet :='';
     PaiementGroupe :=False;
     ModePaiement :='';
     IdentOPS :='';
//PT9  DUCS EDI 4.2     CodAppliPog :='';
     ServUniqPog :=False;
     // Infos EMETTEURSOCIAL
     SiretPet :='';
     AdressePet :='';
     VillePet :='';
     CpPet :='';
     Adresse2Pet :='';
     // Infos ETABLISS
     LibelleET :='';
     Adresse1ET :='';
     Adresse2ET :='';
     VilleET :='';
     CpET :='';
     TelET :='';
     FaxET :='';
     Juridique :='';
// PT4     CentrePayeur := 0;
     NoCentrePayeur := 1; // PT6-2
   end;
// f PT3-3
  EnregDucsEdi.Etab:=TFmul(Ecran).Q.FindField('PDU_ETABLISSEMENT').AsString;
  EnregDucsEdi.Organisme:=TFmul(Ecran).Q.FindField('PDU_ORGANISME').AsString;
  EnregDucsEdi.DateDebut :=TFmul(Ecran).Q.FindField('PDU_DATEDEBUT').AsDateTime;
  EnregDucsEdi.DateFin :=TFmul(Ecran).Q.FindField('PDU_DATEFIN').AsDateTime;
  EnregDucsEdi.NoDucs :=TFmul(Ecran).Q.FindField('PDU_NUM').AsInteger;
  EnregDucsEdi.Siret :=TFmul(Ecran).Q.FindField('PDU_SIRET').AsString;
end;
{***********A.G.L.Privé.*****************************************
Auteur  ...... : PAIE
Créé le ...... : 17/04/2002
Modifié le ... :   /  /    
Description .. : Complément au where pour traiter les cas particuliers
Suite          :  . Dans le cas d'une ducs dossier ACOSS ou IRC avec
Suite          : paiement groupé il n'est affiché que l'élément
Suite          : correpondant à l'établissement principal, même s'il
Suite          : y a des ruptures d'édition
Mots clefs ... : PAIE, PGDUCSEDI
*****************************************************************}
procedure TOF_PGMULDUCSEDI.ActiveWhere(Okok : Boolean);
var
   St : STring;
   D : Variant; // PT8
   DateOk : Boolean;
  Where2:THEdit;//PT14
begin
  WW.Text:='';
  st := '';

{PT6-1 St := '(PDU_ETABLISSEMENT="'+EtabPrinc+'" and PDU_ORGANISME '+
      'IN(SELECT org1.POG_ORGANISME FROM ORGANISMEPAIE org1 WHERE '+
      'org1.POG_DUCSDOSSIER="X" and org1.POG_PAIEGROUPE="X" AND '+
      'org1.POG_NATUREORG="100" AND org1.POG_ETABLISSEMENT=PDU_ETABLISSEMENT)) '+
      'OR '+
      '(PDU_ETABLISSEMENT="'+EtabPrinc+'" and PDU_ORGANISME '+
      'IN(SELECT org2.POG_ORGANISME FROM ORGANISMEPAIE org2 WHERE '+
      'org2.POG_DUCSDOSSIER="X" and org2.POG_PAIEGROUPE<>"X" AND '+
      'org2.POG_NATUREORG="300" AND org2.POG_ETABLISSEMENT=PDU_ETABLISSEMENT)) '+
      'OR '+
      '((PDU_ORGANISME NOT IN(SELECT org3.POG_ORGANISME FROM ORGANISMEPAIE org3 '+
      'WHERE '+
      'org3.POG_NATUREORG ="100" AND org3.POG_ETABLISSEMENT=PDU_ETABLISSEMENT)) '+
      ' AND '+
      '(PDU_ORGANISME NOT IN(SELECT org4.POG_ORGANISME FROM ORGANISMEPAIE org4 '+
      'WHERE '+
      'org4.POG_NATUREORG="300" AND org4.POG_ETABLISSEMENT=PDU_ETABLISSEMENT)) '+
      'OR '+
      '(PDU_ORGANISME NOT IN(SELECT org5.POG_ORGANISME FROM ORGANISMEPAIE org5 '+
      'WHERE '+
      'org5.POG_DUCSDOSSIER="X" AND org5.POG_ETABLISSEMENT=PDU_ETABLISSEMENT)))';}
// d PT8
  DateOk := True;
  try
    D := StrToDate(DateD.text);
  except
     on EConvertError do
       DateOk := False;
  end;
  try
    D := StrToDate(DateD_.text);
  except
     on EConvertError do
       DateOk := False;
  end;
  if (DateOk = True) then
    St:='((PDU_DATEDEBUT >="'+UsDateTime (StrToDate(DateD.text))+'") AND '+
        '(PDU_DATEDEBUT <="'+UsDateTime (StrToDate(DateD_.text))+'") AND '+
        '(PDU_DATEFIN >="'+UsDateTime (StrToDate(DateD.text))+'") AND '+
        '(PDU_DATEFIN <= "'+UsDateTime(StrToDate(DateD_.text))+'")) AND ';

  if (DateOk = True) then
    St := St + '(';
// f PT8
  St := St+
        '(PDU_ORGANISME '+
        'IN(SELECT org1.POG_ORGANISME FROM ORGANISMEPAIE org1 WHERE '+
        'org1.POG_DUCSDOSSIER="X"  AND '+
        'org1.POG_CAISSEDESTIN="X" AND '+
        'org1.POG_ETABLISSEMENT=PDU_ETABLISSEMENT)) '+
        'OR '+
        '(PDU_ORGANISME '+
        'IN(SELECT org2.POG_ORGANISME FROM ORGANISMEPAIE org2 WHERE '+
        'org2.POG_REGROUPEMENT<>"" AND  '+
        'org2.POG_CAISSEDESTIN="X" AND '+
        'org2.POG_NATUREORG="300" AND org2.POG_ETABLISSEMENT=PDU_ETABLISSEMENT)) '+
        'OR '+
        '(PDU_ORGANISME IN(SELECT org3.POG_ORGANISME FROM ORGANISMEPAIE org3 '+
        'WHERE '+
        'org3.POG_CAISSEDESTIN<>"X" AND '+
        'org3.POG_DUCSDOSSIER<>"X"  AND '+
        'org3.POG_REGROUPEMENT="" AND '+
        'org3.POG_ETABLISSEMENT=PDU_ETABLISSEMENT)) '+
//PT7-1
        'OR '+
        '(PDU_ORGANISME IN(SELECT org4.POG_ORGANISME FROM ORGANISMEPAIE org4 '+
        'WHERE '+
        'org4.POG_CAISSEDESTIN="X" AND '+
        'org4.POG_DUCSDOSSIER<>"X"  AND '+
        'org4.POG_REGROUPEMENT="" AND '+
        'org4.POG_ETABLISSEMENT=PDU_ETABLISSEMENT))';
// d PT8
  if (DateOk = True) then
    St := St + ')';
// f PT8
  if St <> '' then WW.Text := st;
  if Q_Mul <> nil then
  begin
    TFMul(Ecran).SetDBListe('PGDUCSENTETE');
  end;

  //DEB PT14
  if GetParamSocSecur('SO_PGDRTVISUETAB',True) then
  begin
    Where2 := THEdit(GetControl('XX_WHERE2'));
    if Where2 <> nil then SetControlText('XX_WHERE2', '');
  end;
  //FIN PT14
end;  // fin ActiveWhere

Initialization
registerclasses([TOF_PGMULDUCSEDI]);
end.
