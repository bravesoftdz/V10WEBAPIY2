{***********UNITE*************************************************
Auteur  ...... : PAIE - MF
Créé le ...... : 18/06/2002
Modifié le ... :   /  /
Description .. : Traitement des déclarations DUCS EDI dossier
Suite ........ : sélectionnées afin de confectionner un fichier au format
Suite ........ : copaym prêt à l'envoi
Mots clefs ... : PAIE; DUCSEDI
*****************************************************************}

{
 PT1 : 07/08/2002 : V585  MF  Traitement code application et serveur unique
 PT2 : 10/10/2202 : V585  MF
                              1- récup. des critères Codappli et serveur unique
                              2- Traitement des fichiers temporaires quand une
                              erreur est survenue pendant les traitement des
                              déclarations sélectionnées.
 PT3 : 06/01/2003 : V591  MF
                              1- Correction des avertissements de compile
 PT4 : 15/01/2003 : V591  MF
                              1- Correction indicateur de test
                              2- Correction du traitement lié à l'indicateur
                                 serveur unique.
                              3- modification des messages d'information
                                 lors du traitement des fichiers de
                                 travail ( n° erreur E/S , nouveau
                                 message si COPAYM.PGI non supprimé)
 PT5 : 18/09/2003 : V_421 MF
                              1- Mise au point CWAS : utilisation des répertoires
                              2- PGIError remplace PGIBox
 PT6 : 18/03/2004 : V_50  MF  FQ 11182  : correction de la mise sur répertoire de
                              travail
 PT7 : 22/04/2004 : V_50  MF  Ajout contrôle Coplib.a existe. Sinon abandon
                              traitement
 PT8 : 30/04/2004 : V_50  MF
                              1- Correction du traitement dans le cas où "serveur
                              unique" est coché. Le message contenait 2 fois la
                              même déclaration, la 1ère étant incomplète.

                              2- Modification de la gestion de l'ensemble des
                              fichiers temporaires. Le nom du fichier est
                              complété du Code Utilisateur. (concerne le fichier
                              COPAYMxxx.PGI). Le fichier COPAYMxxx.PGI est
                              supprimé avant le lancement du traitement. Un
                              message avec validation permet de continuer le
                              traitement.
                              La suppression de l'ensemble des autres fichiers
                              temporaires est fait par TraducCopaym.dll (à partir
                              de la V 1.0.0.6). Fichiers concernés : tracexxx.log,
                              cop_is2xxx.tmp, Copaymxxx.tmp.
 PT9 : 09/02/2006 : V_650  MF DUCS EDI V4.2
 PT10 : 11/04/2007 : V_702 MF Modifs Mise en base des fichiers Ducs EDI
 PT11 : 11/10/2007 : V_800 MF FQ 14845 modif répertoire d'ouverture du coplib.a
 PT12 : 16/11/2007 : V_800 MF Fichiers en base : suppression du fichier ducs dossier extrait
                                et déposé sur réperoire temporaire
}

unit PgEnvoiDucsEdi;

interface
uses
  ed_tools, Hstatus, HMsgBox, Controls, HCtrls,
  uYFILESTD,  // PT10
  EntPaie, sysutils, CbpPath, WINDOWS, Hent1, // PT11
  {$IFDEF EAGLCLIENT}
//unused  UtileAGL,
  emul;
{$ELSE}
  HDB, mul;
{$ENDIF}

{$IFDEF EAGLCLIENT}
procedure LanceEnvoiDucsEdi(liste: THGrid; Ecran: eMul.TFMul; st: string; var fichier_copaym: string; var resultat: integer);
procedure UneDeclaration(Ecran: eMul.TFMul; var FEcrt: TextFile; var FLect: TextFile; Reel, FileN: string; var FileM: string; var PremiereDucs: boolean; var resultat: integer);
  // PT4-2
{$ELSE}
procedure LanceEnvoiDucsEdi(liste: THDBGrid; Ecran: Mul.TFMul; st: string; var fichier_copaym: string; var resultat: integer);
procedure UneDeclaration(Ecran: Mul.TFMul; var FEcrt: TextFile; var FLect: TextFile; Reel, FileN: string; var FileM: string; var PremiereDucs: boolean; var resultat: integer);
  // PT4-2
{$ENDIF}

function TraducCopaym(Nomfichier: string; TypeMess: string; TypeOrganisme: string; PathDat: string; PathStd: string): integer; pascal; external 'TraducCopaym.dll';
//        function TraducCopaym( Nomfichier:string;  TypeMess:string; TypeOrganisme:string; PathDat : string; PathStd : string) : integer ; pascal ; external 'TraducCopaym.ddd';
var
  ServeurUnique: boolean;
  CodeAppli: string;

implementation
{$IFDEF EAGLCLIENT}
{***********A.G.L.Privé.*****************************************
Auteur  ...... : PAIE - MF
Créé le ...... : 18/06/2002
Modifié le ... :   /  /
Description .. : Confectionne le fichier COPAYM.PGI des déclarations
Suite ........ : concaténées et lance la traduction au format copaym
Suite ........ : (TraducCopaym.dll)
Mots clefs ... : PAIE ; DUCSEDI
*****************************************************************}
procedure LanceEnvoiDucsEdi(liste: THGrid; Ecran: eMul.TFMul; st: string; var fichier_copaym: string; var resultat: integer);
{$ELSE}
procedure LanceEnvoiDucsEdi(liste: THDBGrid; Ecran: Mul.TFMul; st: string; var fichier_copaym: string; var resultat: integer);
{$ENDIF}
var
  // PT3-1   i, Chrono, Nbre : integer;
  i: integer;
  //   TProgress,TT: TQRProgressForm ;
  // PT3-1   reponse    : word;
  // PT3-1   LibAppli : string;
  S, SIn, SOut: string;
  FileN, FileM, Support, Reel, FileTrace: string;
  FEcrt, FLect, FE, FL, FTrace: TextFile;
  PremiereDucs: boolean;
  nomfichier, typemessage, typeorganisme: string;
  FileFich_copaym, Dfichier_copaym: string;
  cycle: integer; // PT2-2
  NbreOK: integer; // PT2-2
  NoErreur: integer; // PT4-3
  CheminMaquette, Maquette: string;
  Utilisateur : string; // PT8_2

begin
  NbreOK := 0; // PT3-1
  S := ReadTokenSt(st); // type message
  S := ReadTokenSt(st); // millésime
  S := ReadTokenSt(st); // périodicité
  S := ReadTokenSt(st); // institution
  S := ReadTokenSt(st); // émetteur
  Support := ReadTokenSt(st); // support  PT4-1
  Reel := ReadTokenSt(st); //PT4-1 envoi réel ou test
  S := ReadTokenSt(st); // monnaie
  ServeurUnique := False;

// PT9 DUCS EDI V4.2
//  CodeAppli := ReadTokenSt(st); // code application
  S := ReadTokenSt(st); // Serveur unique
  if (S = 'X') then
    ServeurUnique := True;

  Utilisateur := V_PGI.User;

     // Ouverture du fichier COPAYM.PGI
     // plusieurs d'essais d'ouverture pour permettre à un autre envoi en cours
     // de se terminer et de libérer le fichier
// d PT9 DUCS EDI V4.2
{$IFNDEF DUCS41}
// DUCS 4.2
    FileN := VH_Paie.PgCheminEagl + '\COPAYM' + Utilisateur + '.PGI';
{$ELSE}
// DUCS 4.1
{$IFDEF EAGLCLIENT}
    FileN := VH_Paie.PgCheminEagl + '\COPAYM' + Utilisateur + '.PGI';
{$ELSE}
    FileN := V_PGI.DatPath + '\COPAYM' + Utilisateur + '.PGI';
{$ENDIF}
{$ENDIF}
// f PT9 DUCS EDI V4.2

  cycle := 0;
  while (FileExists(FileN)) and (cycle < 40) do
  begin
    cycle := cycle + 1;
  end;

  if (cycle >= 40) then
  begin
    if PGIAsk ('Fichier inaccessible : ' + FileN + ' #13#10 Voulez-vous continuer le traitement',
      'Un autre envoi est peut être en cours.') = mrYes then
    begin
      DeleteFile(PChar(FileN)); // Supp COPAYM.PGI
      if FileExists(FileN) then
      begin
        PgiInfo('LE FICHIER ' + FileN + ' n''a pas pu être supprimé.',
                'Traduction COPAYM');
        resultat := -1;
        Exit;
      end;
    end
    else
    begin
//    PGIError('Fichier inaccessible : ' + FileN + ' #13#10 Abandon du traitement',
//      'Un autre envoi est peut être en cours.'); // PT5-2
      resultat := -1;
      if (Liste.AllSelected = TRUE) then
        Liste.AllSelected := FALSE
      else
        Liste.ClearSelected;
      exit;
    end;
  end;

  AssignFile(FEcrt, FileN);
  {$I-}
  ReWrite(FEcrt);
  {$I+}
  NoErreur := IoResult;
  if NoErreur <> 0 then
  begin
    PGIError('(' + IntToStr(NoErreur) + ') Fichier inaccessible : ' + FileN,
      'Abandon du traitement');
    closeFile(FEcrt);
    //  déselection des déclarations
    resultat := -1;
    if (Liste.AllSelected = TRUE) then
      Liste.AllSelected := FALSE
    else
      Liste.ClearSelected;
    Exit;
  end;

  PremiereDucs := True;

  if (Liste.AllSelected = TRUE) then
    // Toutes les déclarations on été sélectionnées
  begin
    InitMoveProgressForm(nil, 'Confection du fichier au format COPAYM en cours', 'Veuillez patienter SVP ...', TFmul(Ecran).Q.RecordCount, FALSE, TRUE);

    InitMove(TFmul(Ecran).Q.RecordCount, '');
    TFmul(Ecran).Q.First;
    while not TFmul(Ecran).Q.EOF do
    // Traitement d'une déclaration
    begin

      resultat := 0; // indicateur de traitement OK
      UneDeclaration(Ecran, FEcrt, FLect, Reel, FileN, FileM, PremiereDucs, resultat);
      if (resultat = 0) then NbreOK := NbreOK + 1; // décompte du nombre
      // de déclarations OK
      MoveCurProgressForm('');
      TFmul(Ecran).Q.Next;

    end; // fin while Not TFmul(Ecran).Q.EOF

    if (ServeurUnique = FALSE) and (NbreOK <> 0) then
      // Un seul segment UNZ pour toutes les déclarations
    begin
      SOut := 'UNZ   1';
      {$I-}
      Writeln(FEcrt, SOut);
      {$I+}
      NoErreur := IoResult;
      if NoErreur <> 0 then
      begin
        PGIError('(' + IntToStr(NoErreur) + ') Erreur d''écriture du fichier : ' +
          FileN, 'Abandon du traitement');
        closeFile(FEcrt);
        resultat := -1;
        Exit;
      end;
    end;

    CloseFile(FEcrt); // fermeture fichier COPAYM.PGI

    //   Liste.AllSelected:=False;
    FiniMove;
    FiniMoveProgressForm;
    TFMul(Ecran).bSelectAll.Down := False;

  end // fin     if (Liste.AllSelected=TRUE)
  else
    // seules quelques déclarations ont été sélectionnées
  begin
    InitMoveProgressForm(nil, 'Confection de fichier au format COPAYM en cours', 'Veuillez patienter SVP ...', Liste.NbSelected, FALSE, TRUE);

    InitMove(Liste.NbSelected, '');
    for i := 0 to Liste.NbSelected - 1 do
      // Traitement d'une déclaration
    begin
      Liste.GotoLeBOOKMARK(i);
      {$IFDEF EAGLCLIENT}
      TFmul(Ecran).Q.TQ.Seek(Liste.Row - 1);
      {$ELSE}
      {$ENDIF}
      MoveCur(False);

      resultat := 0; // indicateur de traitement OK
      UneDeclaration(Ecran, FEcrt, FLect, Reel, FileN, FileM, PremiereDucs, resultat); // PT2-2
      if (resultat = 0) then
      //  décompte du nombre de déclarations OK
        NbreOK := NbreOK + 1;


      MoveCurProgressForm('');
      TFmul(Ecran).Q.Next;
    end; // fin      for i:=0 to Liste.NbSelected-1

    if (ServeurUnique = FALSE) and (NbreOK <> 0) then
    // Un seul segment UNZ pour toutes les déclarations
    begin
      SOut := 'UNZ   1';
      {$I-}
      Writeln(FEcrt, SOut);
      {$I+}
      NoErreur := IoResult;
      if NoErreur <> 0 then
      begin
        PGIBox('(' + IntToStr(NoErreur) + ') Erreur d''écriture du fichier : ' + FileN, 'Abandon du traitement');
        closeFile(FEcrt);
        resultat := -1;
        Exit;
      end;
    end;
    CloseFile(FEcrt); // fermeture fichier COPAYM.PGI

    //      Liste.ClearSelected;
    FiniMove;
    FiniMoveProgressForm;
  end; // fin else

  if (NbreOK <> 0) then // PT2-2
    // des déclarations ont été traitées
  begin // PT2-2
    // Traduction au foramt COPAYM (TraducCopaym.dll)
    nomfichier := 'COPAYM' + Utilisateur + '.PGI';
    typemessage := 'COP';
    typeorganisme := 'ACOSS';

    //d PT5-1
{PT11
    Maquette := ChangeStdDatPath('$STD\Coplib.a');
    if FileExists(Maquette) then
      cheminMaquette := ExtractFileDir(Maquette)
    else
    begin
      Maquette := 'Coplib.a';
      if FileExists(Maquette) then
      begin
        Maquette := ExpandFileName(Maquette);
        cheminMaquette := ExtractFileDir(Maquette);
      end
        // PT7
      else
      begin
        PGIBox('Maquette Coplib.a absente', 'Abandon du traitement');
        resultat := -1;
        if FileExists(FileN) then
          DeleteFile(PChar(FileN)); // Supp COPAYM.PGI
        if FileExists(FileN) then
          PgiInfo('LE FICHIER ' + FileN + ' n''a pas pu être supprimé.',
            'Traduction COPAYM');
        Exit;
      end;
      // PT7
    end;*}
    Maquette := TCbpPath.GetCegidDistriApp+'\Coplib.a';
    if FileExists(Maquette) then
      cheminMaquette := ExtractFileDir(Maquette)
    else
    begin
      Maquette := TCbpPath.GetCegidDistriStd+'\Coplib.a';
      if FileExists(Maquette) then
      begin
        cheminMaquette := ExtractFileDir(Maquette);
      end
      else
      begin
        PGIBox('Maquette Coplib.a absente', 'Abandon du traitement');
        resultat := -1;
        if FileExists(FileN) then
          DeleteFile(PChar(FileN)); // Supp COPAYM.PGI
        if FileExists(FileN) then
          PgiInfo('LE FICHIER ' + FileN + ' n''a pas pu être supprimé.',
            'Traduction COPAYM');
        Exit;
      end;
    end;
// d PT9 DUCS EDI V4.2
{$IFNDEF DUCS41}
// DUCS 4.2
    resultat := TraducCopaym(nomfichier,
                             typemessage,
                             typeorganisme,
                             VH_Paie.PgCheminEagl,
                             cheminMaquette);
{$ELSE}
// DUCS 4.1
{$IFDEF EAGLCLIENT}
    resultat := TraducCopaym(nomfichier,
                             typemessage,
                             typeorganisme,
                             VH_Paie.PgCheminEagl,
                             cheminMaquette);
{$ELSE}
    resultat := TraducCopaym(nomfichier,
                               typemessage,
                               typeorganisme,
                               V_PGI.DatPath,
                               V_PGI.StdPath);
{$ENDIF}
{$ENDIF}
// f PT9 DUCS EDI V4.2

    // récup nom du fichier final
// d PT9 DUCS EDI V4.2
{$IFNDEF DUCS41}
// DUCS 4.2
    FileTrace := VH_Paie.PgCheminEagl + '\trace' + Utilisateur + '.log';
{$ELSE}
// DUCS 4.1
{$IFDEF EAGLCLIENT}
    FileTrace := VH_Paie.PgCheminEagl + '\trace' + Utilisateur + '.log';
{$ELSE}
    FileTrace := V_PGI.DatPath + '\trace' + Utilisateur + '.log';
{$ENDIF}
{$ENDIF}
// f PT9 DUCS EDI V4.2
    AssignFile(FTrace, FileTrace);
    {$I-}
    Reset(FTrace);
    {$I+}
    NoErreur := IoResult;
    if NoErreur <> 0 then
    begin
      PGIError('(' + IntToStr(NoErreur) + ') Fichier inaccessible : ' + FileTrace,
        'Abandon du traitement'); // PT4-3 PT5-2
      Exit;
    end;
    if (resultat = 0) then
    begin
      while not eof(FTrace) do
      begin
        {$I-}
        Readln(FTrace, fichier_copaym);
        {$I+}
        NoErreur := IoResult;
        if NoErreur <> 0 then
        begin
          PGIError('(' + IntToStr(NoErreur) + ') Erreur de lecture du fichier : ' +
            fichier_copaym, 'Abandon du traitement'); // PT4-3 PT5-2
          closeFile(FTrace);
          resultat := -1; // PT2-2
          Exit;
        end;
      end;
    end
    else
      PGIError('Anomalie dans le traducteur : ' + IntToStr(resultat),
        'TraducCopaym'); // PT5-2

    CloseFile(FTrace);
    DeleteFile(PChar(FileTrace));

    // support disquette
    if (support = 'DTK') then
    begin
      Dfichier_copaym := 'A:\' + fichier_copaym;
// d  PT9 DUCS EDI V4.2
{$IFNDEF DUCS41}
// DUCS 4.2
      FileFich_copaym := VH_Paie.PgCheminEagl + '\' + fichier_copaym;
{$ELSE}
// DUCS 4.1
{$IFDEF EAGLCLIENT}
      FileFich_copaym := VH_Paie.PgCheminEagl + '\' + fichier_copaym;
{$ELSE}
      FileFich_copaym := V_PGI.DatPath + '\' + fichier_copaym;
{$ENDIF}
{$ENDIF}
// f PT9 DUCS EDI V4.2
      AssignFile(FL, FileFich_copaym);
      {$I-}
      Reset(FL);
      {$I+}

      AssignFile(FE, Dfichier_copaym);
      {$I-}
      ReWrite(FE);
      {$I+}
      // d PT4-3
      NoErreur := IoResult;
      if NoErreur <> 0 then
      begin
        PGIError('(' + IntToStr(NoErreur) + ') Fichier inaccessible : ' +
          Dfichier_copaym, 'Abandon du traitement'); // PT5-2
        if FileExists(FileN) then
          DeleteFile(PChar(FileN)); // Supp COPAYM.PGI
        resultat := -1;
        Exit;
      end;
      // f PT4-3
      while not eof(FL) do
      begin
        {$I-}
        Readln(FL, Sin);
        {$I+}
        SOut := Sin;
        {$I-}
        Writeln(FE, SOut);
        {$I+}

      end;
      CloseFile(FL);
      CloseFile(FE);
      DeleteFile(PChar(FileFich_copaym));
    end; // fin if (support = 'DTK')
    // D PT6
    // support répertoire Travail
    if (support = 'TRA') then
    begin
      Dfichier_copaym := VH_Paie.PgCheminEagl + '\' + fichier_copaym;

// d PT9 DUCS EDI V4.2
{$IFNDEF DUCS41}
// DUCS 4.2
      FileFich_copaym := VH_Paie.PgCheminEagl + '\' + fichier_copaym;
{$ELSE}
// DUCS 4.1
{$IFDEF EAGLCLIENT}
      FileFich_copaym := VH_Paie.PgCheminEagl + '\' + fichier_copaym;
{$ELSE}
      FileFich_copaym := V_PGI.DatPath + '\' + fichier_copaym;
{$ENDIF}
{$ENDIF}
// f PT9 DUCS EDI V4.2

      AssignFile(FL, FileFich_copaym);
      {$I-}
      Reset(FL);
      {$I+}

      AssignFile(FE, Dfichier_copaym);
      {$I-}
      ReWrite(FE);
      {$I+}
      NoErreur := IoResult;
//d PT8-2
      if (NoErreur = 32) then
      // cas où le répertoire partagé = répertoire $DAT
      begin
        CloseFile(FL);
      end
      else
      begin
        if NoErreur <> 0 then
        begin
          PGIError('(' + IntToStr(NoErreur) + ') Fichier inaccessible : ' +
            Dfichier_copaym, 'Abandon du traitement');
          if FileExists(FileN) then
            DeleteFile(PChar(FileN)); // Supp COPAYM.PGI
          resultat := -1;
          Exit;
        end;
      end;
      if (NoErreur <> 32) then
      begin
        while not eof(FL) do
        begin
          {$I-}
          Readln(FL, Sin);
          {$I+}
          SOut := Sin;
          {$I-}
          Writeln(FE, SOut);
          {$I+}

        end;
        CloseFile(FL);
        CloseFile(FE);
        DeleteFile(PChar(FileFich_copaym));
      end;
    end; // fin if (support = 'TRA')
  end;

  if FileExists(FileN) then
    DeleteFile(PChar(FileN)); // Supp COPAYM.PGI
  if FileExists(FileN) then
    PgiInfo('LE FICHIER ' + FileN + ' n''a pas pu être supprimé.', 'Traduction COPAYM')
  else
    PGIInfo('Traduction terminée', 'Traduction COPAYM');
end;
// fin   LanceEnvoiDucsEdi

{***********A.G.L.Privé.*****************************************
Auteur  ...... : PAIE - MF
Créé le ...... : 18/06/2002
Modifié le ... :   /  /
Description .. : Traitement d'une déclaration : intégration dans le fichier
Suite ........ : COPAYM.PGI
Mots clefs ... : PAIE ; DUCSEDI
*****************************************************************}
{$IFDEF EAGLCLIENT}
procedure UneDeclaration(Ecran: eMul.TFMul; var FEcrt: TextFile; var FLect: TextFile; Reel, FileN: string; var FileM: string; var PremiereDucs: boolean; var resultat: integer);
{$ELSE}
procedure UneDeclaration(Ecran: Mul.TFMul; var FEcrt: TextFile; var FLect: TextFile; Reel, FileN: string; var FileM: string; var PremiereDucs: boolean; var resultat: integer);
{$ENDIF}
var
  SIn, SOut: string;
  NoErreur: integer;
  CodeRetour : integer; // PT10
  Crit2 : string; // PT10
  FichierEnBase : boolean; // PT12
begin
  FichierEnBase := False;  // PT12
//d PT10
  Crit2 := ''; 
  // Récupération du fichier Ducs EDI dans la base, FileM = Chemin et nom du fichier
  if (TFmul(Ecran).Q.findfield('PES_GUID1').AsString <> '') then
  begin
    FichierEnBase := True;    // PT12
    Crit2 := Copy(TFmul(Ecran).Q.findfield('PES_FICHIERRECU').AsString, 1, 3);
    CodeRetour := AGL_YFILESTD_EXTRACT (FileM,
                                        'PAIE',
                                        TFmul(Ecran).Q.findfield('PES_FICHIERRECU').AsString ,
                                        'DUC', Crit2, '', '','',
                                        False, 'FRA', 'DOS');
    if (CodeRetour <> -1 ) then
      PGIInfo(AGL_YFILESTD_GET_ERR(CodeRetour) +
                                #13#10 +
                                TFmul(Ecran).Q.findfield('PES_FICHIERRECU').AsString);

  end
  else
  begin
//f PT10
    // Récup nom du fichier
{$IFDEF EAGLCLIENT}
    FileM := VH_Paie.PgCheminEagl + '\' + TFmul(Ecran).Q.findfield('PES_FICHIERRECU').AsString;
{$ELSE}
// d PT9 DUCS EDI V4.2
    if (TFmul(Ecran).Q.findfield('PES_CODAPPLI').AsString = 'V42') then
      FileM := VH_Paie.PgCheminEagl + '\' + TFmul(Ecran).Q.findfield('PES_FICHIERRECU').AsString
    else
      FileM := V_PGI.DatPath + '\' + TFmul(Ecran).Q.findfield('PES_FICHIERRECU').AsString;
// f PT9 DUCS EDI V4.2
{$ENDIF}
  end; // PT10

  // Ouverture fichier DUCS Dossier
  if not FileExists(FileM) then
  begin
    PgiBox('Fichier inexistant : ' + FileM, 'Déclaration non traitée');
    resultat := -1; // PT2-2
    exit;
  end;
  AssignFile(FLect, FileM);
  {$I-}
  Reset(FLect);
  {$I+}
  NoErreur := IoResult;
  if NoErreur <> 0 then
  begin
    PGIBox('(' + IntToStr(NoErreur) + ') Fichier inaccessible : ' + FileM, 'Déclaration non traitée'); // PT4-3
    closeFile(FLect);
    resultat := -1;
    Exit;
  end;

  // Lecture du fichier DUCS Dossier tant que non Fin de Fichier
  // et mise à jour fichier COPAYM.PGI
  while not eof(FLect) do
  begin
    {$I-}
    Readln(FLect, Sin);
    {$I+}
    NoErreur := IoResult;
    if NoErreur <> 0 then
    begin
      PGIBox('(' + IntToStr(NoErreur) + ') Erreur de lecture du fichier : ' + FileM, 'Déclaration non traitée'); // PT4-3
      closeFile(FLect);
      resultat := -1;
      Exit;
    end;

    if (Copy(Sin, 1, 3) <> 'TYP') or (PremiereDucs = True) then
    begin
      // Un seul enregistrement TYP
{* PT9 DUCS EDI V4.2
pour ducs edi v4.1
      if (Copy(Sin, 1, 3) = 'TYP') then
      begin
        SOut := Copy(Sin, 1, 12) + CodeAppli;
      end;*}
      if (Copy(Sin, 1, 3) = 'UNB') then
      begin
        // Enregistrement UNB : si serveur Unique autant d'UNB
        //                      que de déclarations
        if (ServeurUnique = True) or
          ((ServeurUnique = False) and (PremiereDucs = True)) then
        begin
          if (Reel = 'X') then
            SOut := Copy(Sin, 1, 90) + ' '
          else
            SOut := Copy(Sin, 1, 90) + '1';
        end;
        PremiereDucs := False;
      end
      else
        // Autres Enregistrements
        SOut := Sin;

      // PT4-2           PremiereDucs := False;
      {$I-}
      if (SOut <> '') then Writeln(FEcrt, SOut);
      {$I+}
      NoErreur := IoResult;
      if NoErreur <> 0 then
      begin
        PGIBox('(' + IntToStr(NoErreur) + ') Erreur d''écriture du fichier : ' + FileN, 'Déclaration non traitée'); // PT4-3
        closeFile(FLect);
        resultat := -1; // PT2-2
        Exit;
      end;
    end; // fin if (Copy(Sin,1,3) <> 'TYP') or  (PremiereDucs = True)
  end; // fin while not eof(FLect)

  if (ServeurUnique = True) then
  // Un segment UNZ par déclaration
  begin
    SOut := 'UNZ   1';
    {$I-}
    Writeln(FEcrt, SOut);
    {$I+}
    NoErreur := IoResult;
    if NoErreur <> 0 then
    begin
      PGIBox('(' + IntToStr(NoErreur) + ') Erreur d''écriture du fichier : ' + FileN, 'Déclaration non traitée'); // PT4-3
      closeFile(FLect);
      resultat := -1; // PT2-2
      Exit;
    end;
  end;

  CloseFile(FLect); // fermeture fichier DUCS Dossier
// d PT12
  // suppression du fichier DUCS Dossier exporté sur le disque local
  if   (FichierEnBase) and (FileExists(FileM)) then
        DeleteFile(PChar(FileM));
// f PT12
end;

end.

