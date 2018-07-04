{***********UNITE*************************************************
Auteur  ...... : PH
Créé le ...... : 12/11/2001
Modifié le ... : 15/11/2001
Description .. : Unit de lancement de la confection du fichiers
Suite ........ : DADSU à envoyer avec totalisation des S90
Suite ........ : Ecriture et controle entete emetteur
Suite ........ : Mise à jour des champs de la table des envois du social
Mots clefs ... : PAIE;DADSU;DUCSEDI
*****************************************************************}
{
PT1   : 30/01/2002 VG V571 L'enregistrement SIRET du destinataire du CRE n'était
                           pas découpé
PT2   : 13/06/2002 MF V582 Ajout du traitement des Ducs EDI (Pour ne pas
                           créer une nouvelle fiche. A voir pour prochaine
                           Socref nvelle fiche PREP_COPAYM + unité
                           UtofPG_PrepDucsEdi et suppression du PT2)
PT3-1 : 15/07/2002 VG V585 Limitation de l'adresse à 32 caractères
PT3-2 : 15/07/2002 VG V585 Adaptation cahier des charges V6R02
PT4   : 07/10/2002 VG V585 FQ N° 10258 - Ergonomie
PT5   : 10/10/2002 MF V585 récup. paramètres code appli et serveur
PT6-1 : 14/10/2002 VG V585 Changement du nom du fichier DADS-U
PT6-2 : 14/10/2002 VG V585 Si le complément d'adresse n'est pas renseigné, il ne
                           faut plus créer le segment.
PT7   : 11/02/2003 VG V_42 On permet de choisir dorénavant un répertoire de
                           travail pour déposer le fichier (celui défini dans
                           les paramètres société) - FQ N°10471
PT8   : 13/02/2003 MF V_42 Suppression du fichier émis précédemment :
                           suite modif liste PGENVOIDUCS modification du code
                           source (remplace une requête)
PT9   : 02/04/2003 VG V_42 Correction pour la DADS-U BTP : Modification du nom
                           du fichier final
PT10  : 21/08/2003 VG V_42 Traitement de la DADS Bilatérale
PT11  : 16/09/2003 VG V_42 Adaptation cahier des charges V7R01
PT12  : 18/09/2003 MF V_421
                           1- PgiBox remplace MessageAlerte
                           2- Mise au point CWAS
PT13  : 23/09/2003 MF V_42 FQ 10829 : pour les envois DUCS sélection possible
                           de l'émetteur.
PT14  : 02/10/2003 MF V_42 FQ 10872 Suite tests  sous XP. correction du
                           plantage sur FileExists qd le nom du fichier est
                           absent
PT15  : 07/01/2004 VG V_50 Correction du contrôle du nombre d'honoraires.
                           On compte le nombre d'honoraires pour personne
                           morale + nombre d'honoraires pour personne physique
                           FQ N°11040
PT16-1: 05/07/2004 VG V_50 Adaptation cahier des charges V8R00
PT16-2: 05/07/2004 VG V_50 Ergonomie
PT17  : 05/07/2004 VG V_50 Adaptation cahier des charges V8R01
PT18  : 08/10/2004 VG V_50 Modification du nom du fichier TD bilatéral par
                           défaut
PT19  : 18/02/2005 MF V_60 FQ : 11997 On permet la modification du nom de
                           fichier envoyé.
PT20  : 07/10/2005 VG V_60 Adaptation cahier des charges V8R02
PT21  : 02/12/2005 VG V_65 Correction pour prendre en compte l'état de la coche
                           de la fenêtre "gestion des envois" dans le fichier
                           FQ N°12733
PT22  : 09/02/2006 MF V_65 DUCS EDI V 4.2
PT23  : 22/03/2006 VG V_65 Ajout des nouveaux champs emetteur
PT24  : 17/10/2006 VG V_70 Adaptation cahier des charges DADS-U V8R04
PT26  : 09/01/2007 VG V_72 Mauvaise alimentation de la civilité du destinataire
                           du CRE - FQ N°13823
PT27  : 11/01/2007 VG V_72 Mauvaise alimentation du nom du 3ème contact emetteur
                           FQ N°13831
PT28  : 29/06/2007 VG V_72 Fichiers en base
PT29  : 03/04/2007 MF V_80 Mise en place des WebServices Jedeclare
PT29-1: 17/08/2007 VG V_80 Adaptation cahier des charges V8R05 - FQ N°14424
PT29-2: 20/08/2007 VG V_80 Adaptation cahier des charges V8R05 - FQ N°14432
PT29-3: 29/08/2007 VG V_80 Adaptation cahier des charges V8R05
PT30  : 16/11/2007 VG V_80 Suppression du fichier Dossier exporté sur le disque
                           local
PT31  : 26/11/2007 VG V_80 Erreur dans la reprise des critères du fichier en
                           base - FQ N°14989
}
unit UTofPG_PrepDadsU;

interface
uses
  {$IFDEF VER150}
  Variants,
  {$ENDIF}
  StdCtrls, Controls, Classes, sysutils, ComCtrls, HTB97,
  Dialogs, ed_tools, Mailol,
  {$IFDEF EAGLCLIENT}
  UTob, HQry,
  {$ELSE}
  {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} Qre, HDB,
  {$ENDIF}
  HCtrls, HEnt1, EntPaie,
  {$IFDEF COMPTA}
  UtilTrans,
  {$ELSE}
  PgOutils2,
  {$ENDIF}
  {$IFDEF PAIEGRH}
  Paie_cjdc_lib,Paie_Cjdc,  //PT29 JDC
  PgDADSOutils,
  {$ENDIF}
  HMsgBox,
  UTOF,
  Vierge,
  HStatus,
  uYFILESTD;

type
  TOF_PG_PrepDadsU = class(TOF)
  private
    Pan: TPageControl;
    Tbsht: TTabSheet;
    Trace, TraceErr: TListBox;
    QMul: TQUERY; // Query recuperee du mul
    FEcrt, FLect, FZ: TextFile; // Canaux des fichiers
    FileN, FileM, FileZ: string; // Nom des fichiers
    NumEnvoi: Integer; // Numero de l'envoi
    Reel, Emet, Monnaie, Support: string; // Reel = X sinon - ...
    S90: array[1..9] of double; // Totalisations réelles
    S90F: array[1..9] of double; // Totalisations théoriques
    S90Z1: double; // Totalisation des S10 et S90 écrites par la fonction
    BtnLance: TToolbarButton97;
    resultat: integer;
    Destin, periodicite, LSupport: string;
    FichierCopaym: string;
{$IFNDEF DUCS41}
// DUCS 4.2
    TypeDecl : string;
{$ENDIF}
    procedure MajEnvoiDucsEdi(Sender: TObject);
    procedure LancePrepDadsU(Sender: TObject);
    procedure LancePrepDadsB(Sender: TObject);
    procedure ImprimeClick(Sender: TObject);
    procedure AbandonTraitement;
    function TraiteFichier: Boolean;
    function TraiteFichierDADSB: Boolean;
    function TraiteEmetteur: Boolean;
{$IFDEF PAIEGRH}
    function EcritureLigne(Segment, Chaine, Oblig: string): Boolean;
{$ENDIF}
    procedure ReecritChrono(Chrono: Integer);
//    function CumulS90F(Num: Integer; Nombre, St: string): Boolean;
  public
    procedure OnArgument(Arguments: string); override;
  end;

implementation
uses
  UTofPG_MulEnvoiSocial;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : PH
Créé le ...... : 12/11/2001
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_PG_PrepDadsU.LancePrepDadsU(Sender: TObject);
var
ListeJ: HTStrings;
St, S, Chaine: string;
Chrono, i, reponse, rep: integer;
okok: Boolean;
QQ: TQuery;
Nbre: Double;
// d PT29
{$IFDEF PAIEGRH}
Transmission : TTRANSMISSION;
RetourAff_RecepEnvoi_Web  : boolean;
Q1: Tquery;
{$ENDIF}
// f PT29
begin
for i := 1 to 9 do
    S90[i] := 0;
for i := 1 to 9 do
    S90F[i] := 0;

Pan := TPageControl(GetControl('PANELPREP'));
Tbsht := TTabSheet(GetControl('TBSHTTRACE'));
Trace := TListBox(GetControl('LSTBXTRACE'));
TraceErr := TListBox(GetControl('LSTBXERROR'));
if (Trace = nil) or (TraceErr = nil) then
   begin
   PGIBox('La préparation de la DADS-U ne peut pas être lancée', 'Car les composants trace ne sont pas disponibles');
   exit;
   end;

if GetControlText('NOMFIC') = '' then
   begin
   PgiBox('Vous devez indiquer un nom de fichier', Ecran.Caption);
   exit;
   end;

BtnLance.Enabled := FALSE;
if (Pan <> nil) and (Tbsht <> nil) then
   Pan.ActivePage := Tbsht;

if QMul = nil then
   begin
   PGIBox('Erreur sur la liste des fichiers à traiter', 'Abandon du traitement');
   exit;
   end;

if (Grille = nil) then
   begin
   PgiBox('Aucun élément sélectionné', Ecran.Caption);
   exit;
   end;

if (Grille.NbSelected = 0) and (not Grille.AllSelected) then
   begin
   MessageAlerte('Aucun élément sélectionné');
   exit;
   end;

Trace.Items.Add('Debut du traitement');
if support <> 'DTK' then
   begin
//DUCS EDI V 4.2
{$IFNDEF DUCS41}
// DUCS 4.2
     if (TypeDecl = 'DUC') then
       FileN := VH_Paie.PGCheminEagl + '\' + GetControlText('NOMFIC')
     else
{$ENDIF}
     begin
   FileN := '$DAT' + '\' + GetControlText('NOMFIC');
   FileN := ChangeStdDatPath(FileN);
     end;
   end
else
   FileN := 'A:\' + GetControlText('NOMFIC');

if support = 'TRA' then
   FileN := VH_Paie.PGCheminEagl + '\' + GetControlText('NOMFIC');

if FileExists(FileN) then
   begin
   reponse := HShowMessage('5;;Voulez-vous supprimer le fichier DADS-U ' + FileN + ';Q;YN;Y;N', '', '');
   if reponse = 6 then
      DeleteFile(PChar(FileN))
   else
      exit;
   end;

AssignFile(FEcrt, FileN);
{$I-}
ReWrite(FEcrt);
{$I+}

if IoResult <> 0 then
   begin
   PGIBox('Fichier inaccessible : ' + FileN, 'Abandon du traitement');
   Exit;
   end;

NumEnvoi := 1;
QQ := OpenSql('SELECT MAX(PES_NUMEROENVOI) FROM ENVOISOCIAL', TRUE);
if not QQ.EOF then
   try
   NumEnvoi:= QQ.Fields[0].AsInteger + 1;
   except
         on E: EConvertError do
            NumEnvoi:= 1;
   end;
Ferme(QQ);
Trace.Items.Add('Traitement de l''émetteur');
okok := TraiteEmetteur;
Trace.Items.Add('Fin traitement de l''émetteur');
if not okok then
   AbandonTraitement;

if ((Grille.nbSelected) > 0) and (not Grille.AllSelected) and okok then
   begin
   InitMoveProgressForm (nil, 'Début du traitement',
                        'Veuillez patienter SVP ...', Grille.nbSelected, FALSE,
                        TRUE);

   InitMove(Grille.NbSelected, '');
   for i := 0 to Grille.NbSelected - 1 do
       begin
       Grille.GotoLeBOOKMARK(i);
       MoveCur(False);
{$IFDEF EAGLCLIENT}
       QMul.Seek(Grille.Row - 1);
{$ENDIF}
       okok := TraiteFichier;
       if not okok then
          begin
          AbandonTraitement;
          break;
          end;
       st := QMul.findfield('PES_LIBELLE').asString;
       MoveCurProgressForm(St);
       end;
   FiniMove;
   FiniMoveProgressForm;
   end;

if (Grille.AllSelected = TRUE) and okok then
   begin
   InitMoveProgressForm (nil, 'Début du traitement',
                        'Veuillez patienter SVP ...', QMul.RecordCount, FALSE,
                        TRUE);

   InitMove(QMul.RecordCount, '');
   QMul.First;
   while not QMul.EOF do
         begin
         MoveCur(False);
         okok := TraiteFichier;
         if not okok then
            begin
            AbandonTraitement;
            break;
            end;
         st := QMul.findfield('PES_LIBELLE').asString;
         MoveCurProgressForm(St);
         QMul.Next;
         end;
   FiniMove;
   FiniMoveProgressForm;
   end;

// controle des totalisations
Trace.Items.Add('Début de traitement de controle des totalisations');
for i := 1 to 2 do
    begin
    if S90[i] <> S90F[i] then
       begin
       okok := FALSE;
       TraceErr.Items.Add('Erreur dans le comptage des totalisations');
       TraceErr.Items.Add('Trouvées : ' + FloatToStr(S90F[i]) + ' Comptées :' +
                         FloatToStr(S90[i]) + ' Segment SG90.G0x.00' + IntToStr(i));
       end;
    if i = 1 then st := ' Rubriques';
    if i = 2 then st := ' Entreprises';

    Trace.Items.Add('Totalisations trouvées : ' + FloatToStr(S90F[i]) + ' Comptées :' +
                   FloatToStr(S90[i]) + ' pour le segment S90.G01.00' + IntToStr(i) + st);
    end;

Trace.Items.Add('Fin de traitement de controle des totalisations');
// si tout est ok alors on ecrit les S90
if okok then
   begin
   S90[1] := S90[1] + S90Z1; // Pour comptabliser les S10 et S90
   Trace.Items.Add('Ecritures des totalisations');
   Nbre := 0;
   for i := 1 to 2 do
       if S90[i] <> 0 then
          Nbre := Nbre + 1;
   S90[1] := S90[1] + Nbre; // Pour comptabliser les S90 que l'on va écrire
   for i := 1 to 2 do
       begin
       if S90[i] <> 0 then
          begin
          st := 'G01.';
          Chaine := FloatToStr(S90[i]);
          S := 'S90.' + st + '00.00' + IntToStr(i) + ',''' + Chaine + '''';
{$I-}
          Writeln(FEcrt, S);
{$I+}
          if IoResult <> 0 then
             begin
             PGIBox('Erreur d''écriture du fichier : ' + FileN, 'Abandon du traitement');
             TraceErr.Items.Add('Erreur d''écriture du fichier totalisation : ' + S);
             AbandonTraitement;
             okok := FALSE;
             break;
             end;
          end;
       end;
   Trace.Items.Add('Fin d''écritures des totalisations');
   end; // fin de la boucle totalisation des S90

closeFile(FEcrt);
// Traitement de mise à jour des enregistrements de la table envoisocial
if okok then
   begin // B1
   Trace.Items.Add('Mise à jour de la liste des fichiers envoyés');
   if ((Grille.nbSelected) > 0) and (not Grille.AllSelected) then
      begin // B2
      for i := 0 to Grille.NbSelected - 1 do
          begin // B3
          Grille.GotoLeBOOKMARK(i);
{$IFDEF EAGLCLIENT}
          QMul.Seek(Grille.Row - 1);
{$ENDIF}
          Chrono := QMul.findfield('PES_CHRONOMESS').asInteger;
          ReecritChrono(Chrono);
          end; // B3
      Grille.ClearSelected;
      end; // B2
   if (Grille.AllSelected = TRUE) then
      begin // B4
      QMul.First;
      while not QMul.EOF do
            begin // B5
            Chrono := QMul.findfield('PES_CHRONOMESS').asInteger;
            ReecritChrono(Chrono);
            QMul.Next;
            end; // B5
      Grille.AllSelected := False;
      end; // B4
   Trace.Items.Add('Fin de mise à jour de la liste des fichiers envoyés');
   end // B1
else
   DeleteFile(PChar(FileN));

if TraceErr.items.Count >= 1 then
   Trace.Items.Add('Fin du traitement, consultez vos anomalies')
else
   Trace.Items.Add('Fin de traitement');

Pan.ActivePage := Tbsht;
if TraceErr.Items.Count > 0 then
   begin
//Génération d'un fichier de log
{$IFNDEF EAGLCLIENT}
   if MessageDlg('Voulez-vous générez le fichier ErreurDADSU.log sous le répertoire ' +
      V_PGI.DatPath, mtConfirmation, [mbYes, mbNo], 0) = mrYes then
      begin
      if V_PGI.DatPath <> '' then
         FileZ := V_PGI.DatPath + '\ErreurDADSU.log'
      else
         FileZ := 'C:\ErreurDADSU.log';
{$ELSE}
   if MessageDlg('Voulez-vous générez le fichier ErreurDADSU.log sous le répertoire ' +
      VH_Paie.PGCheminEagl, mtConfirmation, [mbYes, mbNo], 0) = mrYes then
      begin
      if VH_Paie.PGCheminEagl <> '' then
         FileZ := VH_Paie.PGCheminEagl + '\ErreurDADSU.log'
      else
         FileZ := 'C:\ErreurDADSU.log';
{$ENDIF}
      if SupprimeFichier(FileZ) = False then
         exit;

      AssignFile(FZ, FileZ);
{$I-}
      ReWrite(FZ);
{$I+}
      if IoResult <> 0 then
         begin
         PGIBox('Fichier inaccessible : ' + FileZ, 'Abandon du traitement');
         Exit;
         end;
      writeln(FZ, 'Création fichier DADS-U : Gestion des messages d''erreur.');
      for i := 0 to TraceErr.Items.Count - 1 do
          begin
          St := TraceErr.Items.Strings[i];
          writeln(FZ, St);
          end;

      CloseFile(FZ);
      PGIInfo('La génération du fichier d''erreurs est terminée', Ecran.Caption);
      end;
   end;

// Traitement specifique envoi fichier par messagerie Outlook ou autre en fonction des préférences
if support = 'TEL' then
   begin
   rep := mrYes;
   if TraceErr.Items.Count - 1 > 0 then
      begin
      Rep := PGIAsk('Vous avez des erreurs#13#10Voulez vous quand même envoyer votre fichier', 'Emission par Internet');
    end;
    if rep = mrYes then
    begin
      ListeJ := HTStringList.Create;
      ListeJ.Add('Veuillez trouver ci-joint notre déclaration annuelle DADS-U');
      ListeJ.Add('Cordialement');
      SendMail('Déclaration DADS-U', '', '', ListeJ, FileN, FALSE);
      ListeJ.Clear;
      ListeJ.Free;
    end;
  end;
// d PT29
{$IFDEF PAIEGRH}
  if support = 'JDC' then
  begin
    rep := mrYes;
    if TraceErr.Items.Count - 1 > 0 then
    begin
      Rep := PGIAsk('Vous avez des erreurs#13#10Voulez vous quand même envoyer votre fichier', 'Emission par Internet');
    end;
    if rep = mrYes then
    begin
      if (Emet <> '') then
      begin
        St := 'SELECT * FROM EMETTEURSOCIAL WHERE PET_EMETTSOC="' + Emet + '"';
        Q1 := OpenSql(St, TRUE);
        if not Q1.EOF then
        begin
          if (InitInfoWS(Q1.FindField('PET_SIRET').AsString,Transmission,True)) then
          begin
            Transmission.ModeEnvoi := '0';      // dépôt
           	Transmission.FichierChemin := FileN;
            Transmission.FichierRequete  := '';
            Transmission.FichierEnvoi := GetControlText('NOMFIC');
            Transmission.Teleprocedure := DADSU;
//      Transmission.idWS := 'cegid.developpement';
//      Transmission.mpWS := 'b4ie65xw';
            Transmission.urlWS := '';
            RetourAff_RecepEnvoi_Web := Aff_RecepEnvoi_Web (Transmission)
          end
          else
            PGIBox('Vous devez émettre le fichier ' + FileN, Ecran.Caption);
        end
        else
          PGIBox('Emetteur mal paramétré ' + Emet, 'transmission non effectuée');
        ferme (Q1);
      end
      else
        PGIBox('Renseigner le champ ''émis par'' ' + Emet, 'Emetteur absent');
    end;
  end;
{$ENDIF}
// f PT29 JDC

  if support <> 'TEL' then
    PGIBox('Vous devez émettre le fichier ' + FileN, Ecran.Caption)
  else
    PGIBox('Fin de la procédure', Ecran.Caption);
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 18/08/2003
Modifié le ... :   /  /
Description .. : Envoi de la DADS Bilatérale
Mots clefs ... :
*****************************************************************}
procedure TOF_PG_PrepDadsU.LancePrepDadsB(Sender: TObject);
var
ListeJ: HTStrings;
St: string;
Chrono, i, reponse, rep: integer;
okok: Boolean;
QQ: TQuery;
begin
okok:= True;
Pan:= TPageControl(GetControl('PANELPREP'));
Tbsht:= TTabSheet(GetControl('TBSHTTRACE'));
Trace:= TListBox(GetControl('LSTBXTRACE'));
TraceErr:= TListBox(GetControl('LSTBXERROR'));
if (Trace = nil) or (TraceErr = nil) then
   begin
   PGIBox('La préparation TD Bilatéral ne peut pas être lancée',
          'car les composants trace ne sont pas disponibles');
   exit;
   end;

if GetControlText('NOMFIC') = '' then
   begin
   PgiBox('Vous devez indiquer un nom de fichier', Ecran.Caption);
   exit;
   end;

BtnLance.Enabled := FALSE;
if (Pan <> nil) and (Tbsht <> nil) then
   Pan.ActivePage := Tbsht;
if QMul = nil then
   begin
   PGIBox('Erreur sur la liste des fichiers à traiter',
          'Abandon du traitement');
   exit;
   end;

if (Grille = nil) then
   begin
   MessageAlerte('Grille de données inexistantes');
   exit;
   end;

if (Grille.NbSelected = 0) and (not Grille.AllSelected) then
   begin
   PgiBox('Aucun élément sélectionné', Ecran.Caption);
   exit;
   end;

Trace.Items.Add('Debut du traitement');
if support <> 'DTK' then
   begin
//DUCS EDI V 4.2
{$IFNDEF DUCS41}
// DUCS 4.2
     if (TypeDecl = 'DUC') then
       FileN := VH_Paie.PGCheminEagl+'\'+GetControlText('NOMFIC')
     else
{$ENDIF}     
     begin
   FileN:= '$DAT'+'\'+GetControlText('NOMFIC');
   FileN:= ChangeStdDatPath(FileN);
     end;
   end
else
   FileN:= 'A:\'+GetControlText('NOMFIC');

if support = 'TRA' then
   FileN:= VH_Paie.PGCheminEagl+'\'+GetControlText('NOMFIC');

if FileExists(FileN) then
   begin
   reponse:= HShowMessage('5;;Voulez-vous supprimer le fichier existant ? ;Q;' +
                          'YN;Y;N', '', '');
   if reponse = 6 then
      DeleteFile(PChar(FileN))
   else
      exit;
   end;

AssignFile(FEcrt, FileN);
{$I-}
ReWrite(FEcrt);
{$I+}
if IoResult <> 0 then
   begin
   PGIBox('Fichier inaccessible : '+FileN, 'Abandon du traitement');
   Exit;
   end;

NumEnvoi := 1;
QQ:= OpenSql('SELECT MAX(PES_NUMEROENVOI) FROM ENVOISOCIAL', TRUE);
if not QQ.EOF then
   try
   NumEnvoi:= QQ.Fields[0].AsInteger + 1;
   except
         on E: EConvertError do
            NumEnvoi:= 1;
   end;
Ferme(QQ);

if ((Grille.nbSelected) > 0) and (not Grille.AllSelected) and okok then
   begin
   InitMoveProgressForm (nil, 'Début du traitement',
                         'Veuillez patienter SVP ...', Grille.nbSelected, FALSE,
                         TRUE);

   InitMove(Grille.NbSelected, '');
   for i := 0 to Grille.NbSelected - 1 do
       begin
       Grille.GotoLeBOOKMARK(i);
       MoveCur(False);
       okok := TraiteFichierDADSB;
       if not okok then
          begin
          AbandonTraitement;
          break;
          end;
       st:= QMul.findfield('PES_LIBELLE').asString;
       MoveCurProgressForm(St);
       end;
   FiniMove;
   FiniMoveProgressForm;
   end;

if (Grille.AllSelected = TRUE) and okok then
   begin
   InitMoveProgressForm (nil, 'Début du traitement',
                         'Veuillez patienter SVP ...', QMul.RecordCount, FALSE,
                         TRUE);

   InitMove(QMul.RecordCount, '');
   QMul.First;
   while not QMul.EOF do
         begin
         MoveCur(False);
         okok:= TraiteFichierDADSB;
         if not okok then
            begin
            AbandonTraitement;
            break;
            end;
         st:= QMul.findfield('PES_LIBELLE').asString;
         MoveCurProgressForm(St);
         QMul.Next;
         end;
   FiniMove;
   FiniMoveProgressForm;
   end;

closeFile(FEcrt);
// Traitement de mise à jour des enregistrements de la table envoisocial
if okok then
   begin
   Trace.Items.Add('Mise à jour de la liste des fichiers envoyés');
   if ((Grille.nbSelected) > 0) and (not Grille.AllSelected) then
      begin
      for i := 0 to Grille.NbSelected - 1 do
          begin
          Grille.GotoLeBOOKMARK(i);
          Chrono:= QMul.findfield('PES_CHRONOMESS').asInteger;
          ReecritChrono(Chrono);
          end;
      Grille.ClearSelected;
      end;
   if (Grille.AllSelected = TRUE) then
      begin
      QMul.First;
      while not QMul.EOF do
            begin
            Chrono:= QMul.findfield('PES_CHRONOMESS').asInteger;
            ReecritChrono(Chrono);
            QMul.Next;
            end;
      Grille.AllSelected := False;
      end;
   Trace.Items.Add('Fin de mise à jour de la liste des fichiers envoyés');
   end
else
   DeleteFile(PChar(FileN));

if TraceErr.items.Count >= 1 then
   Trace.Items.Add('Fin du traitement, consultez vos anomalies')
else
   Trace.Items.Add('Fin de traitement');
Pan.ActivePage := Tbsht;
if TraceErr.Items.Count > 0 then
   begin
//Génération d'un fichier de log
{$IFNDEF EAGLCLIENT}
   if MessageDlg ('Voulez-vous générez le fichier ErreurTDB.log sous le répertoire '+
                  V_PGI.DatPath, mtConfirmation, [mbYes, mbNo], 0) = mrYes then
      begin
      if V_PGI.DatPath <> '' then
         FileZ:= V_PGI.DatPath+'\ErreurTDB.log'
      else
         FileZ:= 'C:\ErreurTDB.log';
{$ELSE}
   if MessageDlg ('Voulez-vous générez le fichier ErreurTDB.log sous le répertoire ' +
                  VH_Paie.PGCheminEagl, mtConfirmation, [mbYes, mbNo], 0) = mrYes then
      begin
      if VH_Paie.PGCheminEagl <> '' then
         FileZ:= VH_Paie.PGCheminEagl+'\ErreurTDB.log'
      else
         FileZ:= 'C:\ErreurTDB.log';
{$ENDIF}
      if SupprimeFichier(FileZ) = False then
         exit;
      AssignFile(FZ, FileZ);
{$I-}
      ReWrite(FZ);
{$I+}
      if IoResult <> 0 then
         begin
         PGIBox('Fichier inaccessible : '+FileZ, 'Abandon du traitement');
         Exit;
         end;
      writeln(FZ, 'Création fichier TD Bilatéral : Gestion des messages d''erreur.');
      for i := 0 to TraceErr.Items.Count - 1 do
          begin
          St:= TraceErr.Items.Strings[i];
          writeln(FZ, St);
          end;
      CloseFile(FZ);
      PGIInfo('La génération du fichier d''erreurs est terminée', Ecran.Caption);
      end;
   end;

// Traitement specifique envoi fichier par messagerie Outlook ou autre en fonction des préférences
if support = 'TEL' then
   begin
   rep:= mrYes;
   if TraceErr.Items.Count - 1 > 0 then
      begin
      Rep:= PGIAsk ('Vous avez des erreurs. #13#10' +
                    'Voulez vous quand même envoyer votre fichier ?',
                    'Emission par Internet');
      end;
   if rep = mrYes then
      begin
      ListeJ:= HTStringList.Create;
      ListeJ.Add ('Veuillez trouver ci-joint notre déclaration des salaires');
      ListeJ.Add ('et/ou honoraires sur support magnétique (TD Bilatéral)');
      ListeJ.Add ('Cordialement');
      SendMail('Déclaration TD Bilatéral', '', '', ListeJ, FileN, FALSE);
      ListeJ.Clear;
      ListeJ.Free;
      end;
   end;

if support <> 'TEL' then
   PGIBox('Vous devez émettre le fichier '+FileN, Ecran.Caption)
else
   PGIBox('Fin de la procédure', Ecran.Caption);
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : PH
Créé le ...... : 12/11/2001
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_PG_PrepDadsU.ImprimeClick(Sender: TObject);
{$IFNDEF EAGLCLIENT}
var
  MPages: tpagecontrol;
{$ENDIF}
begin
  {$IFNDEF EAGLCLIENT}
  MPages := TPageControl(getcontrol('PANELPREP'));
  if MPages <> nil then
    PrintPageDeGarde(MPages, TRUE, TRUE, FALSE, Ecran.Caption, 0);
  {$ENDIF}
end;



{***********A.G.L.Privé.*****************************************
Auteur  ...... : PH
Créé le ...... : 12/11/2001
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_PG_PrepDadsU.OnArgument(Arguments: string);
var
F: TFVierge;
BImprime: ttoolbarbutton97;
AnneePrec, Dest, LeType, NomFic, S, st: string;
begin
inherited;
st:= Trim(Arguments);
S:= ReadTokenSt(st); // Type de traitement DADS-U
LeType:= S;
{$IFNDEF DUCS41}
// DUCS 4.2
TypeDecl := LeType;
{$ENDIF}
SetControlText ('LBLTYPE', GetControlText('LBLTYPE')+' '+
               RechDom('PGENVOISOCIAL', S, FALSE));
S:= ReadTokenSt(st); // Millesime
SetControlText ('LBLMILLESIME', GetControlText('LBLMILLESIME')+' '+
               RechDom('PGANNEE', S, FALSE));
S:= ReadTokenSt(st); // Périodicité
SetControlText ('LBLPERIODICITE', GetControlText('LBLPERIODICITE')+' '+
               RechDom('PGPERIODICITEDUCS', S, FALSE));
Dest:= ReadTokenSt(st); // Destinataire
if (LeType <> 'DA2') then
   SetControlText ('LBLDESTINATAIRE', GetControlText('LBLDESTINATAIRE')+' '+
                  RechDom('PGINSTITUTION', Dest, FALSE));
Emet:= ReadTokenSt(st); // Emetteur
SetControlText ('LBLEMETTEUR', GetControlText('LBLEMETTEUR')+' '+
               Copy(RechDom('PGEMETTEURSOC', Emet, FALSE), 1, 25));
S:= ReadTokenSt(st); // Support
Support:= S;
if (LeType = 'DUC') then
   LSupport:= RechDom('PGSUPPORTEDI', S, FALSE);
SetControlText ('LBLSUPPORT', GetControlText('LBLSUPPORT')+' '+
               RechDom('PGSUPPORTEDI', S, FALSE));
Reel:= ReadTokenSt(st); // Envoi Reel
if (Reel = 'X') then
   SetControlProperty('LBLREEL', 'Checked', TRUE)
else
   SetControlProperty('LBLREEL', 'Checked', FALSE);
S:= ReadTokenSt(st); // Monnaie
Monnaie:= S;
SetControlText ('LBLMONNAIE', GetControlText('LBLMONNAIE')+' '+
               RechDom('PGMONNAIE', S, FALSE));
if (LeType = 'DUC') then
   begin
   SetControlText('TBSHTTRACE', 'Traitement de la DUCS');
   s:= ReadTokenSt(st); // serveur unique
   s:= ReadTokenSt(st); // nom FichierCopaym
   FichierCopaym:= S;
   SetControlText('NOMFIC', FichierCopaym);
   s:= ReadTokenSt(st); // code retour traduction
   resultat:= StrToInt(s);
   BtnLance:= TToolbarButton97(GetControl('BLANCE'));
   if (BtnLance <> nil) then
      BtnLance.OnClick:= MajEnvoiDucsEdi;
   end
else
   if (LeType = 'DAD') then
      begin
      AnneePrec:= Copy(GetControlText('LBLMILLESIME'), 11, 4);

      if (Dest = 'ZBTP') then
         NomFic:= 'DADSU'+AnneePrec +'_BTP.TXT'
      else
         NomFic:= 'DADSU'+AnneePrec+'.TXT';
      SetControlText('NOMFIC', NomFic);
      BtnLance:= TToolbarButton97(GetControl('BLANCE'));
      if BtnLance <> nil then
      BtnLance.OnClick:= LancePrepDadsU;
      BImprime:= ttoolbarbutton97(getcontrol('BIMPRIMER'));
      if Bimprime <> nil then
         Bimprime.Onclick:= ImprimeClick;
      end
else
   if (LeType = 'DA2') then
      begin
      AnneePrec:= Copy(GetControlText('LBLMILLESIME'), 11, 4);
      if (Reel = 'X') then
         NomFic := 'SRN' + AnneePrec + '.TXT'
      else
         NomFic := 'SRN' + AnneePrec + 'T.TXT';
      SetControlText('NOMFIC', NomFic);
      BtnLance := TToolbarButton97(GetControl('BLANCE'));
      if BtnLance <> nil then
         BtnLance.OnClick := LancePrepDadsB;
      BImprime := ttoolbarbutton97(getcontrol('BIMPRIMER'));
      if Bimprime <> nil then
         Bimprime.Onclick := ImprimeClick;
      end;

if not (Ecran is TFVierge) then exit;
F := TFVierge(Ecran);
if F <> nil then
{$IFDEF EAGLCLIENT}
   QMUL := THQuery(F.FMULQ).TQ;
{$ELSE}
   QMUL := F.FMULQ;
{$ENDIF}
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : PH
Créé le ...... : 12/11/2001
Modifié le ... :   /  /
Description .. : Fonction de prise en compte du fichier code retour pour
Suite ........ : savoir
Mots clefs ... :
*****************************************************************}
function TOF_PG_PrepDadsU.TraiteFichier: Boolean;
var
CodeRetour, LL, i : Integer;
Crit2, Crit3, Crit4, Crit5, S, St : string;
FichierEnBase : boolean;
begin
result:= FALSE;
FichierEnBase:= False;  //PT30

{$IFNDEF EAGLCLIENT}
FileM:= V_PGI.DatPath+'\'+QMul.findfield ('PES_FICHIERRECU').AsString;
{$ELSE}
FileM:= VH_Paie.PGCheminEagl+'\'+QMul.findfield ('PES_FICHIERRECU').AsString;
{$ENDIF}

//PT28
//Récupération du fichier DADS-U dans la base, FileM = Chemin et nom du fichier
if (QMul.findfield ('PES_GUID1').AsString<>'') then
   begin
   FichierEnBase:= True;                       //PT30
   Crit2:= Copy (QMul.findfield ('PES_FICHIERRECU').AsString, 20, 4);
   Crit3:= Copy (QMul.findfield ('PES_FICHIERRECU').AsString, 3, 4);
   Crit4:= Copy (QMul.findfield ('PES_FICHIERRECU').AsString, 7, 3);
{PT31
   Crit5:= Copy (QMul.findfield ('PES_FICHIERRECU').AsString, 18, 1);
}
   Crit5:= Copy (QMul.findfield ('PES_FICHIERRECU').AsString, 19, 1);
//FIN PT31
   CodeRetour:= AGL_YFILESTD_EXTRACT (FileM, 'PAIE',
                                      QMul.findfield ('PES_FICHIERRECU').AsString ,
                                      'DAD', Crit2, Crit3, Crit4, Crit5, False,
                                      'FRA', 'DOS');
   if (CodeRetour<>-1) then
      PGIInfo (AGL_YFILESTD_GET_ERR (CodeRetour)+'#13#10'+
               QMul.findfield ('PES_FICHIERRECU').AsString);

   end;
//FIN PT28

Trace.Items.Add ('Traitement de la DADS-U de l''entreprise '+
                 QMul.findfield ('PES_LIBELLE').AsString);
if not FileExists (FileM) then
   begin
   PgiBox ('Fichier DADS-U inexistant pour l''entreprise '+FileM,
           Ecran.caption);
   TraceErr.Items.Add ('Fichier DADS-U inexistant pour l''entreprise '+
                       QMul.findfield ('PES_FICHIERRECU').AsString);
   exit;
   end;
AssignFile(FLect, FileM);
{$I-}
Reset(FLect);
{$I+}
if (IoResult<>0) then
   begin
   PGIBox ('Fichier inaccessible : '+FileM, 'Abandon du traitement');
   TraceErr.Items.Add ('Fichier inaccessible : '+FileM);
   Exit;
   end;

while not eof (FLect) do
      begin
{$I-}
      Readln (FLect, S);
{$I+}
      if (IoResult<>0) then
         begin
         PGIBox ('Erreur de lecture du fichier : '+FileM,
                 'Abandon du traitement');
         TraceErr.Items.Add ('Erreur de lecture du fichier : '+FileM);
         closeFile (FLect);
         Exit;
         end;
      st:= Copy (S, 1, 3);
      if (st <> 'S10') and (st <> 'S90') then
         begin
         S90F[1]:= S90F[1]+1;
         if Copy(S, 1, 14)='S20.G01.00.001' then
            S90F[2]:= S90F[2]+1;
         end;
      if (Copy (S, 1, 3)='S90') then
         begin
         i:= StrToInt (Copy (S, 14, 1));
         LL:= Length (S);
// On recupère le montant en enlevant le dernier caractere
         st:= Copy (S, 17, LL - 17);
         if (not IsNumeric (st)) then
            begin
            PgiBox ('Totalisation erronée '+Copy (S, 1, 14), Ecran.caption);
            TraceErr.Items.Add ('Totalisation erronée '+Copy (S, 1, 14));
            closeFile (FLect);
            exit;
            end;
         S90[i]:= S90[i]+StrToFloat (St);
         end
      else
         begin
{$I-}
         Writeln (FEcrt, S);
{$I+}
         if (IoResult<>0) then
            begin
            PGIBox ('Erreur d''écriture du fichier : '+FileN,
                    'Abandon du traitement');
            TraceErr.Items.Add ('Erreur d''écriture du fichier : '+FileN);
            closeFile (FLect);
            Exit;
            end;
         end;
      end;
closeFile (FLect);

//PT30
if ((FichierEnBase) and (FileExists(FileM))) then
   DeleteFile (PChar (FileM));
//FIN PT30

Trace.Items.Add ('Fin de traitement de l''entreprise '+
                 QMul.findfield('PES_LIBELLE').AsString);
result:= TRUE;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 09/09/2003
Modifié le ... :   /  /
Description .. : Construction du fichier final de la TD Bilatéral
Mots clefs ... : PAIE;DADSB
*****************************************************************}
function TOF_PG_PrepDadsU.TraiteFichierDADSB: Boolean;
var
Crit2, Crit3, Crit4, S : string;
CodeRetour : integer;
FichierEnBase : Boolean;
begin
result := FALSE;
FichierEnBase:= False;  //PT30

{$IFNDEF EAGLCLIENT}
FileM := V_PGI.DatPath + '\' + QMul.findfield('PES_FICHIERRECU').AsString;
{$ELSE}
FileM := VH_PAIE.PGCheminEagl + '\' + QMul.findfield('PES_FICHIERRECU').AsString;
{$ENDIF}

//PT28
//Récupération du fichier DADS-U dans la base, FileM = Chemin et nom du fichier
if (QMul.findfield ('PES_GUID1').AsString<>'') then
   begin
   FichierEnBase:= True;                       //PT30
   Crit2:= Copy (QMul.findfield ('PES_FICHIERRECU').AsString, 20, 4);
   Crit3:= Copy (QMul.findfield ('PES_FICHIERRECU').AsString, 3, 1);
   Crit4:= Copy (QMul.findfield ('PES_FICHIERRECU').AsString, 4, 2);
   CodeRetour:= AGL_YFILESTD_EXTRACT (FileM, 'PAIE',
                                      QMul.findfield ('PES_FICHIERRECU').AsString ,
                                      'DA2', Crit2, Crit3, Crit4, '', False,
                                      'FRA', 'DOS');
   if (CodeRetour<>-1) then
      PGIInfo (AGL_YFILESTD_GET_ERR (CodeRetour)+'#13#10'+
               QMul.findfield ('PES_FICHIERRECU').AsString);

   end;
//FIN PT28

Trace.Items.Add('Traitement TD Bilatéral de l''entreprise ' + QMul.findfield('PES_LIBELLE').AsString);
if not FileExists(FileM) then
   begin
   PgiBox('Fichier TD Bilatéral inexistant pour l''entreprise ' + FileM, Ecran.caption);
   TraceErr.Items.Add('Fichier TD Bilatéral inexistant pour l''entreprise '+
                      QMul.findfield('PES_FICHIERRECU').AsString);
   exit;
   end;
AssignFile(FLect, FileM);
{$I-}
Reset(FLect);
{$I+}
if IoResult <> 0 then
   begin
   PGIBox('Fichier inaccessible : ' + FileM, 'Abandon du traitement');
   TraceErr.Items.Add('Fichier inaccessible : ' + FileM);
   Exit;
   end;

while not eof(FLect) do
      begin
{$I-}
      Readln(FLect, S);
{$I+}
      if IoResult <> 0 then
         begin
         PGIBox('Erreur de lecture du fichier : ' + FileM, 'Abandon du traitement');
         TraceErr.Items.Add('Erreur de lecture du fichier : ' + FileM);
         closeFile(FLect);
         Exit;
         end;
{$I-}
      Writeln(FEcrt, S);
{$I+}
      if IoResult <> 0 then
         begin
         PGIBox('Erreur d''écriture du fichier : ' + FileN, 'Abandon du traitement');
         TraceErr.Items.Add('Erreur d''écriture du fichier : ' + FileN);
         closeFile(FLect);
         Exit;
         end;
      end;
closeFile(FLect);

//PT30
if ((FichierEnBase) and (FileExists(FileM))) then
   DeleteFile (PChar (FileM));
//FIN PT30

Trace.Items.Add('Fin de traitement de l''entreprise ' + QMul.findfield('PES_LIBELLE').AsString);
result := TRUE;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : PH
Créé le ...... : 12/11/2001
Modifié le ... :   /  /
Description .. : Traitement des enregistrements emetteur
Mots clefs ... :
*****************************************************************}
function TOF_PG_PrepDadsU.TraiteEmetteur: Boolean;
{$IFDEF PAIEGRH}
var
CodeIso, Libelle, st, st2: string;
Q1: Tquery;
Ok: Boolean;
i, Long, Long2: Integer;
{$ENDIF PAIEGRH}
begin
result := FALSE;
{$IFDEF PAIEGRH}
St := 'SELECT * FROM EMETTEURSOCIAL WHERE PET_EMETTSOC="' + Emet + '"';
Q1 := OpenSql(St, TRUE);
if not Q1.EOF then
   i:=1
else
   i:=0;
ok:=False;

while i = 1 do
      begin
      st := Q1.FindField('PET_SIRET').AsString;
      Ok := EcritureLigne('S10.G01.00.001.001', Copy(St, 1, 9), 'O');
      if not Ok then
         break;
      Ok := EcritureLigne('S10.G01.00.001.002', Copy(St, 10, 5), 'O');
      if not Ok then
         break;
      st := Q1.FindField('PET_RAISONSOC').AsString;
      Ok := EcritureLigne('S10.G01.00.002', St, 'O');
      if not Ok then
         break;
      st := Q1.FindField('PET_COMPLADR').AsString;
      st2 := Q1.FindField('PET_COMPLADR2').AsString;
      Long := Length(st);
      Long2 := Length(st2);
      if ((Long <> 0) and (Long2 <> 0)) then
         begin
         st := Copy(st + ' ', 1, Long + 1);
         st := Copy(st + st2, 1, Long + 1 + Long2);
         end;
      Long := Length(st);
{PT24
      if (Long > 32) then
         st := Copy(st, 1, 32);
}
      if (Long > 38) then
         st := Copy(st, 1, 38);
//FIN PT24
      Ok := EcritureLigne('S10.G01.00.003.001', St, 'C');
      if not Ok then
         break;
      st := Q1.FindField('PET_NUMEROVOIE').AsString;
      Ok := EcritureLigne('S10.G01.00.003.003', St, 'C');
      if not Ok then
         break;
      st := Q1.FindField('PET_BISTER').AsString;
      Ok := EcritureLigne('S10.G01.00.003.004', St, 'C');
      if not Ok then
         break;
      st := Q1.FindField('PET_NOMVOIE').AsString;
      Ok := EcritureLigne('S10.G01.00.003.006', St, 'C');
      if not Ok then
         break;
      st := Q1.FindField('PET_INSEECOMMUNE').AsString;
      Ok := EcritureLigne('S10.G01.00.003.007', St, 'F');
      if not Ok then
         break;
      st := Q1.FindField('PET_VILLE').AsString;
{PT24
      Ok := EcritureLigne('S10.G01.00.003.009', Copy(St, 1, 26), 'C');
}
      Ok := EcritureLigne('S10.G01.00.003.009', St, 'C');
      if not Ok then
         break;
      st := Q1.FindField('PET_CODEPOSTAL').AsString;
      if (Q1.FindField('PET_PAYS').AsString <> 'FRA') then
{PT24
         Ok := EcritureLigne('S10.G01.00.003.010', Copy(St, 1, 10), 'O')
}
         Ok := EcritureLigne('S10.G01.00.003.010', St, 'O')
      else
         begin
         if ((st > '00000') and (st < '99999')) then
            Ok := EcritureLigne('S10.G01.00.003.010', Copy(St, 1, 5), 'O')
         else
            Ok := EcritureLigne('S10.G01.00.003.010', '99999', 'O');
         end;
      if not Ok then break;
      st := Q1.FindField('PET_BURDISTRIB').AsString;
      st := PGUpperCase(st);
      Ok := EcritureLigne('S10.G01.00.003.012', St, 'O');
      if not Ok then
         break;
      st := Q1.FindField('PET_PAYS').AsString;
{PT29-1
      if ((st <> 'FRA') and (st <> '')) then
}
      if ((st<>'FRA') and (st<>'GUF') and (st<>'GLP') and (st<>'MCO') and
         (st<>'MTQ') and (st<>'NCL') and (st<>'PYF') and (st<>'SPM') and
         (st<>'REU') and (st<>'ATF') and (st<>'WLF') and (st<>'MYT') and
         (st<>'')) then
//FIN PT29-1
         begin
         PaysISOLib(st, CodeIso, Libelle);
         Ok := EcritureLigne('S10.G01.00.003.013', Copy(CodeIso, 1, 3), 'C');
         if not Ok then
            break;
{PT24
         Ok := EcritureLigne('S10.G01.00.003.014', Copy(Libelle, 1, 32), 'C');
}
         Ok := EcritureLigne('S10.G01.00.003.014', Copy(Libelle, 1, 38), 'C');
         if not Ok then
            break;
         end;
      st := IntToStr(NumEnvoi);
      Ok := EcritureLigne('S10.G01.00.004', St, 'O');
      if not Ok then
         break;
      Ok := EcritureLigne('S10.G01.00.005', 'CEGID PAIE PGI', 'F');
      if not Ok then
         break;
      Ok := EcritureLigne('S10.G01.00.006', 'CEGID', 'F');
      if not Ok then
         break;
      Ok := EcritureLigne('S10.G01.00.007', V_PGI.NumVersion, 'F');
      if not Ok then
         break;
      Ok := EcritureLigne('S10.G01.00.009', '40', 'O');
      if not Ok then
         break;
      Reel:= GetControlText ('LBLREEL');
      if Reel = 'X' then
         st := '02'
      else
         St := '01';
      Ok := EcritureLigne('S10.G01.00.010', St, 'O');
      if not Ok then
         break;
{PT24
      Ok := EcritureLigne('S10.G01.00.011', 'V08R02', 'O');
}
{PT29-1
      Ok := EcritureLigne('S10.G01.00.011', 'V08R04', 'O');
}
      Ok := EcritureLigne('S10.G01.00.011', CDC, 'O');
      if not Ok then
         break;
      Ok := EcritureLigne('S10.G01.00.012', '01', 'O');
      if not Ok then
         break;
      st := Q1.FindField('PET_CREDADSU').AsString;
      if (St='X') then
         begin
         st:= Q1.FindField ('PET_SIRET').AsString;
{PT29-3
         if ((st <> '') and (Q1.FindField('PET_CREMEDIA').AsString = '03')) then
}
         if (st<>'') then
            begin
            Ok:= EcritureLigne ('S10.G01.00.013.001', Copy (st, 1, 9), 'O');
            if not Ok then
               break;
            Ok:= EcritureLigne ('S10.G01.00.013.002', Copy (st, 10, 5), 'O');
            if not Ok then
               break;
{PT29-2
            st := Q1.FindField('PET_CREMEDIA').AsString;
            Ok := EcritureLigne('S10.G01.00.014', st, 'O');
            if not Ok then
               break;
}
            st := Q1.FindField('PET_CREINDICATIF').AsString;
            Ok := EcritureLigne('S10.G01.00.015.001', st, 'O');
            if not Ok then
               break;
            end;
{PT29-2
         if ((st <> '') and (Q1.FindField('PET_CREMEDIA').AsString = '05')) then
            begin
            Ok := EcritureLigne('S10.G01.00.013.001', Copy(st, 1, 9), 'C');
            if not Ok then
               break;
            Ok := EcritureLigne('S10.G01.00.013.002', Copy(st, 10, 5), 'C');
            if not Ok then
               break;
            st := Q1.FindField('PET_CREMEDIA').AsString;
            Ok := EcritureLigne('S10.G01.00.014', st, 'C');
            if not Ok then
               break;
            st := Q1.FindField('PET_CRECIVILITE').AsString;
            if (St = 'MR') then
               Ok:= EcritureLigne('S10.G01.00.015.002', '01', 'C')
            else
            if (St = 'MME') then
               Ok:= EcritureLigne('S10.G01.00.015.002', '02', 'C')
            else
            if (St = 'MLE') then
               Ok:= EcritureLigne('S10.G01.00.015.002', '03', 'C');
            if not Ok then
               break;
            st := Q1.FindField('PET_CRENOM').AsString;
            Ok := EcritureLigne('S10.G01.00.015.003', st, 'C');
            if not Ok then
               break;
            end;
}
         end;

// Gestion des contacts
      st:= Q1.FindField('PET_CIVIL1DADSU').AsString;
      if (St = 'MR') then
         Ok := EcritureLigne('S10.G01.01.001.001', '01', 'O')
      else
      if (St = 'MME') then
         Ok := EcritureLigne('S10.G01.01.001.001', '02', 'O')
      else
      if (St = 'MLE') then
         Ok := EcritureLigne('S10.G01.01.001.001', '03', 'O');
      if not Ok then
         break;
      st := Q1.FindField('PET_CONT1DUDS').AsString;
      Ok := EcritureLigne('S10.G01.01.001.002', st, 'O');
      if not Ok then
         break;
      st := Q1.FindField('PET_DOMAINEDUDS1').AsString;
      Ok := EcritureLigne('S10.G01.01.002', st, 'O');
      if not Ok then
         break;
      st := Q1.FindField('PET_APPEL1DUDS').AsString;
      Ok := EcritureLigne('S10.G01.01.005', st, 'O');
      if not Ok then
         break;
      st := Q1.FindField('PET_TEL1DADSU').AsString;
      Ok := EcritureLigne('S10.G01.01.006', st, 'O');
      if not Ok then
         break;
      st := Q1.FindField('PET_FAX1DADSU').AsString;
      Ok := EcritureLigne('S10.G01.01.007', st, 'C');
      if not Ok then
         break;
      st := Q1.FindField('PET_CONT2DUDS').AsString;
      if St <> '' then
         begin
         st:= Q1.FindField('PET_CIVIL2DADSU').AsString;
         if (St = 'MR') then
            Ok := EcritureLigne('S10.G01.01.001.001', '01', 'O')
         else
         if (St = 'MME') then
            Ok := EcritureLigne('S10.G01.01.001.001', '02', 'O')
         else
         if (St = 'MLE') then
            Ok := EcritureLigne('S10.G01.01.001.001', '03', 'O');
         if not Ok then
            break;
         st := Q1.FindField('PET_CONT2DUDS').AsString;
         Ok := EcritureLigne('S10.G01.01.001.002', st, 'O');
         if not Ok then
            break;
         st := Q1.FindField('PET_DOMAINEDUDS2').AsString;
         Ok := EcritureLigne('S10.G01.01.002', st, 'O');
         if not Ok then
            break;
         st := Q1.FindField('PET_APPEL2DUDS').AsString;
      Ok := EcritureLigne('S10.G01.01.005', st, 'O');
      if not Ok then
         break;
      st := Q1.FindField('PET_TEL2DADSU').AsString;
      Ok := EcritureLigne('S10.G01.01.006', st, 'O');
      if not Ok then
         break;
      st := Q1.FindField('PET_FAX2DADSU').AsString;
      Ok := EcritureLigne('S10.G01.01.007', st, 'C');
      if not Ok then
         break;
         end;
      st := Q1.FindField('PET_CONT3DUDS').AsString;
      if St <> '' then
         begin
         st:= Q1.FindField('PET_CIVIL3DADSU').AsString;
         if (St = 'MR') then
            Ok := EcritureLigne('S10.G01.01.001.001', '01', 'O')
         else
         if (St = 'MME') then
            Ok := EcritureLigne('S10.G01.01.001.001', '02', 'O')
         else
         if (St = 'MLE') then
            Ok := EcritureLigne('S10.G01.01.001.001', '03', 'O');
         if not Ok then
            break;
         st := Q1.FindField('PET_CONT3DUDS').AsString;      //PT27
         Ok := EcritureLigne('S10.G01.01.001.002', st, 'O');
         if not Ok then
            break;
         st := Q1.FindField('PET_DOMAINEDUDS3').AsString;
         Ok := EcritureLigne('S10.G01.01.002', st, 'O');
         if not Ok then
            break;
         st := Q1.FindField('PET_APPEL3DUDS').AsString;
      Ok := EcritureLigne('S10.G01.01.005', st, 'O');
      if not Ok then
         break;
      st := Q1.FindField('PET_TEL3DADSU').AsString;
      Ok := EcritureLigne('S10.G01.01.006', st, 'O');
      if not Ok then
         break;
      st := Q1.FindField('PET_FAX3DADSU').AsString;
      Ok := EcritureLigne('S10.G01.01.007', st, 'C');
      if not Ok then
         break;
         end;
      i := 0; // Pour sortir
      end;
Ferme(Q1);
if ok then
   result := TRUE;
{$ENDIF PAIEGRH}
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : PH
Créé le ...... : 12/11/2001
Modifié le ... :   /  /
Description .. : Fonction d'ecriture et de controle du fichier dads-U
Mots clefs ... :
*****************************************************************}
{$IFDEF PAIEGRH}
function TOF_PG_PrepDadsU.EcritureLigne(Segment, Chaine, Oblig: string): Boolean;
begin
result := FALSE;
if (Chaine = '') and (Oblig = 'O') then
   begin
   TraceErr.Items.Add('Le segment '+Segment+' ne possède pas de valeur');
   exit;
   end;

if (Chaine = '') and ((Oblig = 'F') or (Oblig = 'C')) then
   begin
   result:= TRUE;
   exit;
   end;

writeln(FEcrt, Segment+','''+Chaine+'''');
S90Z1:= S90Z1+1;
result:= TRUE;
end;
{$ENDIF}

{***********A.G.L.Privé.*****************************************
Auteur  ...... : PH
Créé le ...... : 12/11/2001
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_PG_PrepDadsU.AbandonTraitement;
begin
  TraceErr.items.add('Une erreur est survenue lors de la conception du fichier');
  TraceErr.items.add('Le traitement est interrompu');
  Trace.items.add('Le traitement est abandonné, vérifiez vos erreurs !');
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : PH
Créé le ...... : 12/11/2001
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_PG_PrepDadsU.ReecritChrono(Chrono: Integer);
var
st : string;
begin
if (Emet='') then
   st:= 'UPDATE ENVOISOCIAL SET'+
        ' PES_ENVOIREEL="'+Reel+'",'+
        ' PES_NUMEROENVOI='+IntToStr (NumEnvoi)+','+
        ' PES_FICHIEREMIS="'+GetControlText ('NOMFIC')+'",'+
        ' PES_SUPPORTEMIS="'+Support+'",'+
        ' PES_STATUTENVOI="TRA",'+
        ' PES_PREPARELE="'+UsDateTime (Date)+'",'+
        ' PES_ENVOYEPAR="'+V_PGI.User+'" WHERE'+
        ' PES_CHRONOMESS='+IntToStr (Chrono)
else
   st:= 'UPDATE ENVOISOCIAL SET'+
        ' PES_ENVOIREEL="'+Reel+'",'+
        ' PES_NUMEROENVOI='+IntToStr (NumEnvoi)+','+
        ' PES_FICHIEREMIS="'+GetControlText ('NOMFIC')+'",'+
        ' PES_EMETSOC="'+Emet+'",'+
        ' PES_SUPPORTEMIS="'+Support+'",'+
        ' PES_STATUTENVOI="TRA",'+
        ' PES_PREPARELE="'+UsDateTime (Date)+'",'+
        ' PES_ENVOYEPAR="'+V_PGI.User+'" WHERE'+
        ' PES_CHRONOMESS='+IntToStr (Chrono);
ExecuteSql (st);
end;


{***********A.G.L.Privé.*****************************************
Auteur  ...... : PAIE - MF
Créé le ...... : 18/06/2002
Modifié le ... :   /  /
Description .. : Mise à jour de la table ENVOISOCIAL + contrôle des
Suite ........ : éventuelles erreurs + lancement messagerie si
Suite ........ : télétransmission internet
Mots clefs ... : PAIE; DUCSEDI
*****************************************************************}
procedure TOF_PG_PrepDadsU.MajEnvoiDucsEdi(Sender: TObject);
var
  ListeJ: HTStrings;
  St: string;
  Chrono, i, rep: integer;
  okok: Boolean;
  FicAPurger: string;
  FileARenommer : string;
  NomModifie     : boolean;
// d PT29
{$IFDEF PAIEGRH}
  Transmission : TTRANSMISSION;
  RetourAff_RecepEnvoi_Web  : boolean;
  Q1: Tquery;
{$ENDIF}
// f PT29
begin

  okok := True;
  Pan := TPageControl(GetControl('PANELPREP'));
  Tbsht := TTabSheet(GetControl('TBSHTTRACE'));
  Trace := TListBox(GetControl('LSTBXTRACE'));
  TraceErr := TListBox(GetControl('LSTBXERROR'));
  if (Trace = nil) or (TraceErr = nil) then
  begin
    PGIBox('La mise à jour des envois ne peut pas être lancée', 'Car les composants trace ne sont pas disponibles');
    exit;
  end;
  if GetControlText('NOMFIC') = '' then
  begin
    PgiBox('Vous devez indiquer un nom de fichier', Ecran.Caption);
    exit;
  end;
//Contrôle si nom de fichier modifié
  if (FichierCopaym <> GetControlText('NOMFIC')) then
    NomModifie := true
  else
    NomModifie := false;

  if support <> 'DTK' then
{$IFNDEF DUCS41}
// DUCS 4.2
{$IFNDEF EAGLCLIENT}
    if (TypeDecl = 'DUC') then
      FileN := VH_Paie.PGCheminEagl + '\' + GetControlText('NOMFIC')
    else
      FileN := V_PGI.DatPath + '\' + GetControlText('NOMFIC')
{$ELSE}
    FileN := VH_Paie.PGCheminEagl + '\' + GetControlText('NOMFIC')
{$ENDIF}
{$ENDIF}

{$IFDEF DUCS41}
// DUCS 4.1
{$IFNDEF EAGLCLIENT}
    FileN := V_PGI.DatPath + '\' + GetControlText('NOMFIC')
{$ELSE}
    FileN := VH_Paie.PGCheminEagl + '\' + GetControlText('NOMFIC')
{$ENDIF}
{$ENDIF}

  else
    FileN := 'A:\' + GetControlText('NOMFIC');
  if support = 'TRA' then
    FileN := VH_Paie.PGCheminEagl + '\' + GetControlText('NOMFIC');

  if (NomModifie) then
  begin
    if support <> 'DTK' then
{$IFNDEF DUCS41}
// DUCS 4.2
{$IFNDEF EAGLCLIENT}
    if (TypeDecl = 'DUC') then
      FileARenommer := VH_Paie.PGCheminEagl + '\' + FichierCopaym
    else
      FileARenommer := V_PGI.DatPath + '\' + FichierCopaym
{$ELSE}
    FileARenommer := VH_Paie.PGCheminEagl + '\' + FichierCopaym
{$ENDIF}
{$ENDIF}

{$IFDEF DUCS41}
// DUCS 4.1
{$IFNDEF EAGLCLIENT}
      FileARenommer := V_PGI.DatPath + '\' + FichierCopaym
{$ELSE}
      FileARenommer := VH_Paie.PGCheminEagl + '\' + FichierCopaym
{$ENDIF}
{$ENDIF}

    else
      FileARenommer := 'A:\' + FichierCopaym;

    if support = 'TRA' then
      FileARenommer := VH_Paie.PGCheminEagl + '\' + FichierCopaym ;

    if (not RenameFile(FileARenommer, FILEN)) then
    begin
      if (PGIAsk('Ce fichier existe déjà. Voulez-vous le remplacer?', Ecran.Caption) = mrYes) then
      begin
        if FileExists(FILEN) then
          DeleteFile(PChar(FILEN));
        RenameFile(FileARenommer, FILEN)
      end
      else
        exit;
    end;
  end;

  BtnLance.Enabled := FALSE;
  if (Pan <> nil) and (Tbsht <> nil) then Pan.ActivePage := Tbsht;
  if QMul = nil then
  begin
    PGIBox('Erreur sur la liste des fichiers à traiter', 'Abandon du traitement');
    exit;
  end;
  if (Grille = nil) then
  begin
    MessageAlerte('Grille de données inexistantes');
    exit;
  end;
  if (Grille.NbSelected = 0) and (not Grille.AllSelected) then
  begin
    PgiBox('Aucun élément sélectionné', Ecran.Caption);
    exit;
  end;
  Trace.Items.Add('Début du traitement');

  if (resultat = 0) then
  begin
    if (Reel = 'X') then
      Trace.Items.add('Envoi réel ' + Destin + ' ' + periodicite + ' - ' + GetControlText('NOMFIC') + ' support ' + LSupport)
    else
      Trace.Items.add('Envoi de test ' + Destin + ' ' + periodicite + ' - ' + GetControlText('NOMFIC') + ' support ' + LSupport);
  end
  else
  begin
    okok := False;
    if (resultat = -1) then
      TraceErr.Items.Add('Impossible d''ouvrir ou de créer le fichier Copaym.ref');
    if (resultat = -2) then
      TraceErr.Items.Add(' Type de message inconnu.');
    if (resultat = -3) then
      TraceErr.Items.Add('Impossible de charger la table de description des messages.');
    if (resultat = -4) then
      TraceErr.Items.Add('Impossible d''ouvir le fichier maquette.');
    if (resultat = -5) then
      TraceErr.Items.Add('Impossible d''ouvir le fichier pré-copaym.');
    if (resultat = -6) then
      TraceErr.Items.Add('Erreur lors de la génération du fichier copaym.');
  end;

  // Traitement de mise à jour des enregistrements de la table envoisocial
  if okok then
  begin // B1
    Trace.Items.Add('Mise à jour de la liste des fichiers envoyés');
    if ((Grille.nbSelected) > 0) and (not Grille.AllSelected) then
    begin // B2
      for i := 0 to Grille.NbSelected - 1 do
      begin // B3
        Grille.GotoLeBOOKMARK(i);
        {$IFDEF EAGLCLIENT}
        QMul.Seek(Grille.Row - 1);
        {$ENDIF}
        Chrono := QMul.findfield('PES_CHRONOMESS').asInteger;
        if (QMul.findfield('PES_FICHIEREMIS').asString <> '') and
          (QMul.findfield('PES_FICHIEREMIS').asString <> NULL) then
        begin
          {$IFNDEF EAGLCLIENT}
//DUCS EDI V 4.2
          if (QMul.findfield('PES_CODAPPLI').asString = 'V42') then
            FicAPurger := VH_Paie.PGCheminEagl + '\' + QMul.findfield('PES_FICHIEREMIS').asString
          else
          FicAPurger := V_PGI.DatPath + '\' + QMul.findfield('PES_FICHIEREMIS').asString;
          {$ELSE}
          FicAPurger := VH_Paie.PGCheminEagl + '\' + QMul.findfield('PES_FICHIEREMIS').asString;
          {$ENDIF}
          if (FicAPurger <> '') and FileExists(FicAPurger) and (FicAPurger <> FILEN) then
            DeleteFile(PChar(FicAPurger));
        end;
        ReecritChrono(Chrono);
      end; // B3
      Grille.ClearSelected;
    end; // B2
    if (Grille.AllSelected = TRUE) then
    begin // B4
      QMul.First;
      while not QMul.EOF do
      begin // B5
        Chrono := QMul.findfield('PES_CHRONOMESS').asInteger;
        if (QMul.findfield('PES_FICHIEREMIS').asString <> '') and
          (QMul.findfield('PES_FICHIEREMIS').asString <> NULL) then
        begin
          {$IFNDEF EAGLCLIENT}
//DUCS EDI V 4.2
          if (QMul.findfield('PES_CODAPPLI').asString = 'V42') then
            FicAPurger := VH_Paie.PGCheminEagl + '\' + QMul.findfield('PES_FICHIEREMIS').asString
          else
          FicAPurger := V_PGI.DatPath + '\' + QMul.findfield('PES_FICHIEREMIS').asString;
          {$ELSE}
          FicAPurger := VH_Paie.PGCheminEagl + '\' + QMul.findfield('PES_FICHIEREMIS').asString;
          {$ENDIF}
          if (FicAPurger <> '') and FileExists(FicAPurger) then
            DeleteFile(PChar(FicAPurger));
        end;
        ReecritChrono(Chrono);
        QMul.Next;
      end; // B5
      Grille.AllSelected := False;
    end; // B4
    Trace.Items.Add('Fin de mise à jour de la liste des fichiers envoyés');
  end // B1
  else DeleteFile(PChar(FileN));

  if TraceErr.items.Count >= 1 then Trace.Items.Add('Fin du traitement, consultez vos anomalies')
  else Trace.Items.Add('Fin de traitement');
  Pan.ActivePage := Tbsht;
  if TraceErr.Items.Count > 0 then
  begin
    //Génération d'un fichier de log
    {$IFNDEF EAGLCLIENT}
    if MessageDlg('Voulez-vous générez le fichier ErreurDUCS.log sous le répertoire ' +
      V_PGI.DatPath, mtConfirmation, [mbYes, mbNo], 0) = mrYes then
    begin
      if V_PGI.DatPath <> '' then FileZ := V_PGI.DatPath + '\ErreurDUCS.log'
      else FileZ := 'C:\ErreurDUCS.log';
      {$ELSE}
    if MessageDlg('Voulez-vous générez le fichier ErreurDUCS.log sous le répertoire ' +
      VH_Paie.PGCheminEagl, mtConfirmation, [mbYes, mbNo], 0) = mrYes then
    begin
      if VH_Paie.PGCheminEagl <> '' then FileZ := VH_Paie.PGCheminEagl + '\ErreurDUCS.log'
      else FileZ := 'C:\ErreurDUCS.log';
      {$ENDIF}
      if SupprimeFichier(FileZ) = False then exit;
      AssignFile(FZ, FileZ);
      {$I-}ReWrite(FZ);
      {$I+}
      if IoResult <> 0 then
      begin
        PGIBox('Fichier inaccessible : ' + FileZ, 'Abandon du traitement');
        Exit;
      end;
      writeln(FZ, 'Création fichier DUCS : Gestion des messages d''erreur.');
      for i := 0 to TraceErr.Items.Count - 1 do
      begin
        St := TraceErr.Items.Strings[i];
        writeln(FZ, St);
      end;
      CloseFile(FZ);
      PGIInfo('La génération du fichier d''erreurs est terminée', Ecran.Caption);
    end;
  end;
  // Traitement specifique envoi fichier par messagerie Outlook ou autre en
  // fonction des préférences
  if support = 'TEL' then
  begin
    rep := mrYes;
    if TraceErr.Items.Count - 1 > 0 then
    begin
      Rep := PGIAsk('Vous avez des erreurs#13#10Voulez vous quand même envoyer votre fichier', 'Emission par Internet');
    end;
    if rep = mrYes then
    begin
      ListeJ := HTStringList.Create;
      ListeJ.Add('Veuillez trouver ci-joint notre déclaration Ducs EDI');
      ListeJ.Add('Cordialement');
//DUCS EDI V 4.2
{$IFDEF DUCS41}
// DUCS 4.1
      SendMail('Déclaration DUCS EDI', '', '', ListeJ, FileN, FALSE);
{$ELSE}
// DUCS 4.2
      SendMail('DUCS EDI', '', '', ListeJ, FileN, FALSE);
{$ENDIF}
      ListeJ.Clear;
      ListeJ.Free;
    end;
// d PT29 JDC
{$IFNDEF PAIEGRH}
  end;
{$ELSE}
  end
  else
  if support = 'JDC' then
  begin
    rep := mrYes;
    if TraceErr.Items.Count - 1 > 0 then
    begin
      Rep := PGIAsk('Vous avez des erreurs#13#10Voulez vous quand même envoyer votre fichier', 'Emission par Internet');
    end;
    if rep = mrYes then
    begin
      if (Emet <> '') then
      begin
        St := 'SELECT * FROM EMETTEURSOCIAL WHERE PET_EMETTSOC="' + Emet + '"';
        Q1 := OpenSql(St, TRUE);
        if not Q1.EOF then
        begin
          if (InitInfoWS(Q1.FindField('PET_SIRET').AsString,Transmission,True)) then
          begin
            Transmission.ModeEnvoi := '0';      // dépôt
           	Transmission.FichierChemin := FileN;
            Transmission.FichierRequete  := '';
            Transmission.FichierEnvoi := GetControlText('NOMFIC');
            Transmission.Teleprocedure := DUCS;
//      Transmission.idWS := 'cegid.developpement';
//      Transmission.mpWS := 'b4ie65xw';
            Transmission.urlWS := '';
            RetourAff_RecepEnvoi_Web := Aff_RecepEnvoi_Web (Transmission)
          end
          else
            PGIBox('Vous devez émettre le fichier ' + FileN, Ecran.Caption);
        end
        else
          PGIBox('Emetteur mal paramétré ' + Emet, 'transmission non effectuée');
        ferme (Q1);
      end
      else
        PGIBox('Renseigner le champ ''émis par'' ' + Emet, 'Emetteur absent');
    end;
  end;
{$ENDIF}  
// f PT29 JDC
  if (support <> 'TEL') and (support <> 'JDC')then   // PT29
    PGIBox('Vous devez émettre le fichier ' + FileN, Ecran.Caption)
  else
    PGIBox('Fin de la procédure', Ecran.Caption);
end;
initialization
  registerclasses([TOF_PG_PrepDadsU]);
end.

