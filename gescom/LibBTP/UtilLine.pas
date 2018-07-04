unit UtilLine;

interface
Uses WinInet,
     Classes,
{$IFDEF EAGLCLIENT}
      MaineAGL,
{$ELSE}
    	{$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
      DB, Fe_Main,
{$ENDIF}
     HMSGBox,
     Controls,
     ParamSoc,
     HDB,
	   HEnt1,
     Ent1,
     HCtrls,
     Windows,
     SysUtils,
     ShellApi,
     UToz,
     Forms,
     SplashS1,
     RestaureS1,
     FileCtrl,
     MajTable,
     UTob,
     HTB97,
     StdCtrls,
     UTOF,
     AglMail,
     MailOl,
     MajExe;

    Function ControleFichierLog(PathRapport : String) : boolean;
    Function DetectionConnexionInternet: Boolean;
    //
    Procedure BackupPGIS1(ok_Connect : boolean = true);
    Procedure ExecuteWait(Programme, Parametre :PChar; Visible :boolean);
    procedure RestorePGIS1;
    //

Type ParamConnection = RECORD
	   		Path						: String;
        PathSauve				: String;
	  	  FichierSQL			: String;
  	    FichierRapport	: String;
	      FichierSauve		: String;
	    	StNom 					: String;
		    StServer				: String;
	  	  StUser					: String;
		    StPassWord			: String;
		    StBase					: String;
		    StConnexion			: String;
        ZipFile					: Boolean;
        ConnectBase			: Boolean;
        NbSauvegarde		: Integer;
     end;
    //
     ParamMail = Record
        EMail						: String;
        CopieA					: String;
        Sujet						: HString;
        Fichier					: String;
        MessMail	 			: String;
     End;
    //
    Function AfficheSplash  : TSplash_S1;
    Function AfficheRestaure: TRestaure_S1;
    Function CompressionFichier(StConnection : ParamConnection; MessAttente : TSplash_S1) : Boolean;
    Function DetermineParametre : ParamConnection;
    Function DoBackup(StConnection : ParamConnection) : Boolean;
    Function DoRestore(StConnection : ParamConnection) : Boolean;
		Function GetFileTimes (const FileName: string; var DateAcces : TDateTime) : Boolean;
    Function ListingRepSauve(RepSauvegarde : String) : Integer;
    Function RechercheFichier(path : String; Messattente : TRestaure_S1) : string;
    //
    Procedure ChargeComboRestaure(MessAttente : TRestaure_S1; Path : String);
    //
    procedure ChargeEcran(Ecran : TForm; TobParam : Tob);
    procedure DescentChargeUnNiveau (Composant : TWinControl; TobParam : Tob);
    procedure ChargeComponent (composant : TWinControl; LaTob : Tob);
    //
    procedure StockeInfos(Ecran : TForm; TobParam : Tob);
		procedure DescentStockeUnNiveau (Composant : TWinControl; TobParam : Tob);
    procedure StockeComponent (composant : TWinControl; TobParam : Tob);
    Procedure EnvoieMail(ParametreMail : ParamMail);

implementation

Var	FicSauveBat	 : Textfile;
    TitreMess		 : String;
    //
    Ok_Sauve		 : Boolean;

    StConnection : ParamConnection;

//const TypeChampsTraitable : string = 'TEdit;THedit;THNumEdit;TCheckBox;THCheckbox;{THDBSpinEdit;}THSpinEdit;TGroupBox;THValComboBox;THDBMultiValComboBox;THMultiValComboBox;THCritMaskEdit; TPanel; THPanel; TRadioButton';
const TypeChampsTraitable : string = 'TEdit;THedit;THNumEdit;TCheckBox;THCheckbox;THSpinEdit;TGroupBox;THgroupBox;THValComboBox;THMultiValComboBox;THCritMaskEdit;TPanel;THPanel';

const
    // libellés des messages
  	TexteMessage: array[ 1..19 ] of hstring = (
     	 {1}  'Vous devez être administrateur pour effectuer une %s',
       {2}  'Assurez-vous d''être le seul connecté avant d''effectuer une %s',
    	 {3}  'Désirez-vous sauvegarder la société ',
    	 {4}  'La %s n''a pas été effectuée.',
    	 {5}  'Le fichier de %s existe désirez-vous le remplacer ?',
    	 {6}  'Répertoire inexistant désirez-vous le créer ?',
    	 {7}  '%s effectuée avec succès',
       {8}  'Serveur de connexion introuvable veuillez vérifier les paramètres de Connection',
       {9}  'Voulez-vous afficher le rapport ?',
       {10} '%s en cours ...',
       {11} 'Le serveur d''application n''a pas été trouvé, veuillez vérifier les paramètres de connection',
       {12} 'Désirez-vous restaurer la société ',
       {13} 'Vous vous apprêtez à restaurer une sauvegarde de votre dossier, les données sauvegardées vont remplacer celles de votre dossier actuel.#13 Désirez-vous continuer ?',
       {14} 'La %s a été annulée.',
			 {15} 'Fichier .BAK inexistant.#13 Désirez-vous le rechercher ?',
 			 {16} 'Compression du fichier de sauvegarde en cours...',
       {17} 'La compression du fichier de sauvegarde n''a pas pu être exécutée',
       {18} 'Une %s existe déjà pour cette journée. #13 Désirez-vous l''écraser ?',
       {19} 'Assurez-vous d''être le seul utilisateur pendant la %s !'
        );


function DetectionConnexionInternet: Boolean;
var
  dwFlags: DWord;
begin
  dwFlags := INTERNET_CONNECTION_MODEM or INTERNET_CONNECTION_LAN or INTERNET_CONNECTION_PROXY;
  Result := InternetGetConnectedState(@dwFlags, 0);
end;

//------------------------------------------------------------------------------
// ALPY 20/07/2007 - reecriture de la fonction pour prise en compte de la gestion
// des sauvegardes multiples.
//------------------------------------------------------------------------------
Procedure BackupPGIS1(ok_Connect : boolean = true);
//Var StConnection : ParamConnection;
begin

		TitreMess := 'Sauvegarde';

	  // FF - 12/11/2007 - FQ-15116 - Pas de sauvegarde si non administrateur
	  // Le code ci-dessous est celui qui était en place dans les autres fonctions
    if not V_PGI.Superviseur then
       begin
       PgiInfo(Format(TexteMessage[1],[TitreMess]), TitreMess);
       exit;
       end;

	  if not OKToutSeul then
       Begin
       PgiInfo(Format(TexteMessage[2],[TitreMess]), TitreMess);
       exit ;
       end;

	  //Demande de confirmation de sauvegarde...
    if (PGIAsk(TraduireMemoire(TexteMessage[3] + V_PGI.NomSociete + ' ?'), TitreMess) <> mrYes) then
		   begin
       PGIInfo(Format(TexteMessage[4],[TitreMess]), TitreMess);
	     exit;
		   end;

    StConnection := DetermineParametre;

    StConnection.ConnectBase := Ok_Connect;
    StConnection.FichierSQL := 'Sauvegarde.SQL';

    //Création du fichier bat et execution pour sauvegarde de la base !!!!
    if DoBackup(StConnection) then
       PGIInfo(Format(TexteMessage[7],[TitreMess]), TitreMess)
    Else
       PGIInfo(Format(TexteMessage[4],[TitreMess]), TitreMess);

end;

{***********A.G.L.***********************************************
Auteur  ...... : Yann-Cyril PELUD
Créé le ...... : 03/12/2002
Modifié le ... :   /  /
Description .. : Restore spécifique S1 qui permet de gérer les base MSDE
								 YCP 09/10/2006 changement de méthode pour le backup/restore
Suite ........ : directement dans S1.exe
Mots clefs ... : S1;DATABASE;MSDE
*****************************************************************}
procedure RestorePGIS1;
begin

 		TitreMess := 'Restauration';

		if not V_PGI.Superviseur then
       begin
       PGIInfo(Format(TexteMessage[1],[TitreMess]), TitreMess);
       exit;
       end;

		if not OKToutSeul then
       Begin
			 PGIInfo(Format(TexteMessage[2],[TitreMess]), TitreMess);
       exit;
       end;

    if PgiAsk(TexteMessage[13], TitreMess)=mrno then
 		   begin
    	 PGIBox(Format(TexteMessage[14],[TitreMess]), TitreMess);
	     exit;
		   end;

    StConnection := DetermineParametre;

    StConnection.FichierSQL := 'Restauration.SQL';

    //Création du fichier SQL et execution pour Restauration de la base !!!!
    if DoRestore(StConnection) then
       PGIInfo(Format(TexteMessage[7],[TitreMess]), TitreMess)
    Else
       PGIInfo(Format(TexteMessage[4],[TitreMess]), TitreMess);

end;


function DetermineParametre : ParamConnection;
Begin

    Result.FichierRapport := 'SaveRapport.LOG';

    //Détermination du nom du fichier de sauvegarde
    Result.FichierSauve := V_PGI.DBName + '.BAK';

    //Chargement du chemin de sauvegarde
    Result.Path := IncludeTrailingBackslash(V_PGI.StdPath);
    //Result.PathSauve := IncludeTrailingBackslash(V_PGI.CheminSauve);
    Result.PathSauve :=  IncludeTrailingBackslash(GetParamsocSecur('SO_FICHIERSAUVE', IncludeTrailingBackslash(V_PGI.CheminSauve)));

    //Chargement des variables système
    Result.StNom 		 := V_PGI.CurrentAlias;
    Result.StServer	 := V_PGI.StServer;
    Result.StUser		 := V_PGI.User;
    Result.StPassWord := V_PGI.Password;
    Result.StBase		 := V_PGI.DBName;

    Result.ZipFile := GetParamsocSecur('SO_ZIPFILE', false);

    Result.NbSauvegarde := GetParamsocSecur('SO_HOWLASTSAVE', 0);

    //Recherche de la clé de connexion : Clef : HKEY_LOCAL_MACHINE\SOFTWARE\ODBC\ODBC.INI\<Nom de connexion>\Server
		//result.StConnexion := GetFromRegistry(HKEY_LOCAL_MACHINE, 'Software\ODBC\ODBC.INI\' + Result.StServer, 'Server', Result.StConnexion, True);
  Result.StConnexion := 'Provider=Microsoft.Jet.OLEDB.4.0;User ID=Admin;Data Source='+ Result.StBase +
  ';Mode=Share Deny None;Extended Properties="";Jet OLEDB:System database="";Jet OLEDB:Registry Path="";'+
  'Jet OLEDB:Database Password="";Jet OLEDB:Engine Type=5;Jet OLEDB:Database Locking Mode=1;'+
  'Jet OLEDB:Global Partial Bulk Ops=2;Jet OLEDB:Global Bulk Transactions=1;Jet OLEDB:New Database Password="";'+
  'Jet OLEDB:Create System Database=False;Jet OLEDB:Encrypt Database=False;Jet OLEDB:Don''t Copy Locale on Compact=False;'+
  'Jet OLEDB:Compact Without Replica Repair=False;Jet OLEDB:SFP=False';

end;

//Cette fonction va permettre de créer dans le répertoire sauve, un fichier bat qui
//permettra de gérer la sauvegarde automatique de la base !!!
function DoBackup(StConnection : ParamConnection) : Boolean;
var ZoneTxt 		: String;
    stParametre	: String;
    PathSQL			: String;
    PathRapport	: String;
    MessAttente : TSplash_S1;
    Ok_Connect	: Boolean;
    NbFile			: Integer;
    I						: Integer;
begin

    Result := False;

    NbFile := 0;

    if StConnection.StServer = '' then
       Begin
       PGIInfo(TexteMessage[11], TitreMess);
	     exit;
       end;

    if StConnection.StConnexion = '' then
       Begin
       PGIInfo(TexteMessage[8], TitreMess);
	     exit;
       end;

   //création du repertoire de sauvegarde
   If Not DirectoryExists(StConnection.Path) then
      begin
      if (PGIAsk(TraduireMemoire(TexteMessage[6]), TitreMess) = mrYes) then
         ForceDirectories(StConnection.path)
      else
         SelectDirectory(StConnection.Path, [sdAllowCreate, sdPerformCreate, sdPrompt],1000)
      end;

   if Not DirectoryExists(StConnection.PathSauve) then
      begin
      if (PGIAsk(TraduireMemoire(TexteMessage[6]), TitreMess) = mrYes) then
         ForceDirectories(StConnection.pathSauve)
      else
         SelectDirectory(StConnection.Pathsauve, [sdAllowCreate, sdPerformCreate, sdPrompt],1000)
      end;

   //Contrôle du nombre de sauvegarde déjà présentent...
   NbFile := ListingRepSauve(StConnection.Pathsauve);

   //Si .BAK existe demande suppression
   if Not Ok_Sauve then
      Begin
      if (PGIAsk(Format(TraduireMemoire(TexteMessage[18]), [TitreMess]), TitreMess) <> mrYes) then
         Begin
         result := false;
  	     exit;
         end;
      end;

   //si le nombre de fichier = 0 le fichier reste un .BAK
   if NbFile = 0 then
      StConnection.FichierSauve := V_PGI.DBName + '.BAK'
   //si le nombre de fichier est inférieur au nombre de fichier à garder on renome le .BAK existant en BNbFile
   Else if NBFile < StConnection.NbSauvegarde then
      RenameFile(StConnection.Pathsauve + V_PGI.DBName + '.BAK', StConnection.Pathsauve + V_PGI.DBName + '.B' + IntToStr(NBFile))
	 //Sauvegarde des fichiers histo des sauvegardes si le nombre de fichiers = nombre total on décale de 1
   Else if NbFile = StConnection.NbSauvegarde then
      Begin
      //boucle sur l'ensemble des fichiers Bxx ou Zxx
      For i := 1 to NbFile do
          Begin
          if I = 1 then
             Deletefile(StConnection.Pathsauve + V_PGI.DBName + '.B' + intToStr(I))
          else if I = NbFile then
             RenameFile(StConnection.Pathsauve + V_PGI.DBName + '.BAK', StConnection.Pathsauve + V_PGI.DBName + '.B' + IntToStr(I-1))
          Else
             RenameFile(StConnection.Pathsauve + V_PGI.DBName + '.B' + intToStr(I), StConnection.Pathsauve + V_PGI.DBName + '.B' + IntToStr(I-1))
          end;
      end;

   Messattente := AfficheSplash;
   Messattente.Show;

	 //Arrêter la connexion à la base de donnée courrante...
   MessAttente.Label1.Caption := 'Deconnection de la base ' + stConnection.StBase + ' en cours...';
   Messattente.Refresh;
	 Logout;

   //DeconnecteHalley;

   MessAttente.Label1.Caption := Format(TexteMessage[10],[TitreMess]);
   MessAttente.Refresh;

   //Contrôle existence du fichier .SQL si existe suppression
   PathSQL := StConnection.Path + StConnection.FichierSQL; //'C:\PGI00\STD' + FichierSQL;
   if FileExists(PathSQL) then Deletefile(PathSQL);

   //Contrôle Existence du Fichier .LOG si existe suppression
   PathRapport := StConnection.Path + StConnection.FichierRapport; //'C:\PGI00\STD' + SaveRapport.Log';
   if FileExists(PathRapport) then Deletefile(PathRapport);

   //Ordre SQL de Sauvegarde
   MessAttente.Label1.Caption := Format(TexteMessage[19],[TitreMess]);
   MessAttente.Refresh;

   ZoneTxt := 'BACKUP DATABASE [' + StConnection.StBase + ']';
   ZoneTxt := ZoneTxt + ' TO DISK = "' + StConnection.PathSauve + StConnection.FichierSauve + '"' ;
   ZoneTxt := ZoneTxt + ' WITH FORMAT, NAME = "' + stConnection.StBase;
   ZoneTxt := ZoneTxt + '-Sauvegarde Complète Base de données "';
   ZoneTxt := zonetxt + ', SKIP, NOREWIND, NOUNLOAD, STATS = 10';

   Executesql(ZoneTxt);

   //Si le fichier .BAK existe on concidère que tout est ok
   if FileExists(StConnection.PathSauve + StConnection.FichierSauve) then  Result := True;

   //Controle du paramétrage si l'on doit zipper le fichier
   if Result then
	    if StConnection.ZipFile then
         if not Compressionfichier(StConnection,MessAttente) then
            PGIInfo(TexteMessage[17], TitreMess);

   //Reconnection de la base
   MessAttente.Label1.Caption := 'Connection à la base ' + stConnection.StBase + ' en cours...';
   Messattente.Refresh;

   if StConnection.ConnectBase then
	    Ok_Connect := connecteHalley(StConnection.Stbase, False, CHARGESOCIETEHALLEY, nil, nil, nil, False, False);

   MessAttente.Free;

end;

//Cette fonction va permettre de créer dans le répertoire sauve, un fichier bat qui
//permettra de gérer la Restauration automatique de la base !!!
function DoRestore(StConnection : ParamConnection) : Boolean;
var ZoneTxt 		: String;
    stParametre	: String;
    PathSQL			: String;
    PathRapport	: String;
    NouveauNom	: String;
    MessAttente : TRestaure_S1;
begin

    Result := False;

    if StConnection.StServer = '' then
       Begin
       PGIInfo(TexteMessage[11], TitreMess);
	     exit;
       end;

    if StConnection.StConnexion = '' then
       Begin
       PGIInfo(TexteMessage[8], TitreMess);
	     exit;
       end;

    Messattente := AfficheRestaure;
    //
	  MessAttente.TRestaure.Visible := True;
	  MessAttente.ComboBox1.Visible := True;
    //
    //Chargement du Combo de sélection  du fichier à restaurer
 	  ChargeComboRestaure(MessAttente, StConnection.PathSauve);

    //Affichage de la fiche de sélection pour restauration...
    Messattente.ShowModal;
    //
    Messattente.Show;

    //Si click sur bouton abandon
    if Messattente.ComboBox1.Items.Strings[Messattente.ComboBox1.ItemIndex] = '' then
       Begin
       Result := False;
       PGIBox(Format(TexteMessage[14],[TitreMess]), TitreMess);
	     exit;
       end;

    //si click sur bouton valide !!!
    StConnection.FichierSauve := Messattente.ComboBox1.Items.Strings[Messattente.ComboBox1.ItemIndex];
    Stconnection.PathSauve := IncludeTrailingBackslash (StConnection.PathSauve) + StConnection.FichierSauve;
    //

    //Vérification si le fichier .BAK de sauvegarde existe !!!
    IF Not FileExists(Stconnection.PathSauve) then
       begin
       if (PGIAsk(TraduireMemoire(TexteMessage[15]), TitreMess) = mrYes) then
          Begin
          Messattente.ODGetInfosBAK.Filter := 'Fichier sauvegarde SQL (*.BAK)|*.BAK|Fichier Compressé (*.ZIP)|*.ZIP|Tous les fichiers (*.*)|*.*';
          Stconnection.PathSauve := RechercheFichier(StConnection.Path, Messattente);
          end
       else
           Begin
				   MessAttente.Free;
        	 exit;
           end;
       end;

   //Contrôle existence du fichier .BAT si existe suppression
   PathSQL := StConnection.Path + StConnection.FichierSQL;
   if FileExists(PathSQL) then Deletefile(PathSQL);

   //Contrôle Existence du Fichier .LOG si existe suppression
   PathRapport := StConnection.Path + StConnection.FichierRapport;
   if FileExists(PathRapport) then Deletefile(PathRapport);


	 //Arrêter la connexion à la base de donnée courrante...
   MessAttente.Label1.Caption := 'Deconnection de la base ' + stConnection.StBase + ' en cours...';
   Messattente.Refresh;
	 Logout;
   DeconnecteHalley;

   MessAttente.Label1.Caption := Format(TexteMessage[19],[TitreMess]);
   Messattente.Refresh;

   //Restauration de la base de données
	 ZoneTxt := 'RESTORE DATABASE [' +  stConnection.StBase + '] ';
   ZoneTxt := ZoneTXt + 'FROM  DISK = "' + StConnection.Pathsauve + '" ';
   ZoneTxt := ZoneTxt + 'WITH FILE = 1, ';
	 ZoneTxt := ZoneTxt + 'NOUNLOAD, REPLACE, STATS = 10';

   Executesql(ZoneTxt);

   if FileExists(PathRapport) then Result := ControleFichierLog(PathRapport);

   MessAttente.Label1.Caption := 'Connection à la base ' + stConnection.StBase + ' en cours...';
   Messattente.Refresh;

   Result := connecteHalley(StConnection.Stbase, False, CHARGESOCIETEHALLEY, nil, nil, nil, False, False);

   MessAttente.Free;

end;

Procedure ChargeComboRestaure(MessAttente : TRestaure_S1; Path : String);
Var Chemin 	: String;
		Extend	: String;
    Info		: TSearchRec;
begin

  Chemin := Path;

  //Pour être sur que la barre oblique finisse le nom du chemin

  Chemin := IncludeTrailingBackslash(Chemin);

  if StConnection.ZipFile then
     Extend := '*.Z*'
  else
	   Extend := '*.B*';

  //Recherche de la première entrée du répertoire
  If FindFirst(Chemin  + Extend, faAnyFile,Info)=0 Then
     Begin
     Repeat
     If Not ((Info.Attr And faDirectory)=0) Then // On ne gère que les fichiers si 0 ce sont les répertoires
     Else
        begin
        MessAttente.ComboBox1.Items.Add(Info.FindData.cFileName);
        end;
     Until FindNext(Info)<>0; // Il faut ensuite rechercher l'entrée suivante }
     //Dans le cas ou une entrée au moins est trouvée il faut }
     //appeler FindClose pour libérer les ressources de la recherche
     FindClose(Info);
     End;

  Messattente.ComboBox1.ItemIndex := 0;

end;

Function RechercheFichier(path : String; MessAttente : TRestaure_S1) : string;
Var Extension : String;
    X : Integer;
begin

  Result := '';

  if Path = '' then
	   Messattente.ODGetInfosBAK.InitialDir := 'C:\'
  else
  	 Messattente.ODGetInfosBAK.InitialDir := Path;

  if Messattente.ODGetInfosBAK.Execute then
  	 begin
     if Messattente.ODGetInfosBAK.FileName <> '' then
    		begin
        Result := Messattente.ODGetInfosBAK.FileName;
        end;
  	 end;

end;

Function ControleFichierLog(PathRapport : String) : boolean;
var Fichier : TextFile;
		lig 		: string;
begin

  //Assignation du fichier
	assignfile(Fichier, PathRapport);

  //Acces au début du fichier
  reset(Fichier);

	//Lecture du fichier log pour vérification si erreur ou non !!!!
  //cette boucle répètera la lecture d'une ligne tant que nous n'aurons pas atteint la fin du document EOF
  while not eof(Fichier) do
      begin
			Readln(Fichier,lig);
			end;

  if pos('successfully', Lig) <> 0 then
     Begin
     Result := true;
     exit;
     end
  Else
     Begin
     Result := False;
     PGIINfo(Lig, TitreMess);
     End;

	//cette procédure ferme le fichier
	CloseFile(Fichier);

end;


//Affichage d'une fenetre Splash pour Sauvegarde...
Function AfficheSplash : TSplash_S1;
begin

     result := TSplash_S1.Create(Application);

     result.Animate1.Active := true;

End;

//Affichage de la fenêtre de restauration
Function AfficheRestaure : TRestaure_S1;
begin

    result := TRestaure_S1.Create(Application);

    result.Label1.Caption := Format(TexteMessage[10],[TitreMess]);
    result.Refresh;

    result.Animate1.Active := true;

End;

{***********A.G.L.***********************************************
Auteur  ...... : Franck Vautrain
Créé le ...... : 07/04/2008
Modifié le ... :   /  /
Description .. : Procedure permettant l'éxecution en arrière plan d'une
Suite ........ : commande DOS
Mots clefs ... : COMMANDE DOS, EXECUTION, CACHÉ
*****************************************************************}
Procedure ExecuteWait(Programme, Parametre : PChar; Visible :boolean);
var ShExecInfo : TShellExecuteInfo;
begin

    FillChar(ShExecInfo, SizeOf(ShExecInfo), 0);

    with ShExecInfo do
         begin
         cbSize := SizeOf(ShExecInfo);
         fMask  := SEE_MASK_NOCLOSEPROCESS;
         lpFile := Programme;      { le nom du programme }
         lpVerb := 'open';
         lpDirectory := PChar(ExtractFilepath(programme));
         lpParameters := Parametre; {les parametres}
         if visible Then
            nShow  := SW_SHOW
         else
            nShow  := SW_HIDE;
    		 end;

    {on execute le programme}
    if ShellExecuteEx(@ShExecInfo) then
       begin
       WaitForSingleObject(ShExecInfo.hProcess, INFINITE); { on attends un temps indefinie que l'appli s'arrete }
       end;

end;

//lister le répertoire de sauvegarde pour vérification du nombre de sauvegarde déjà existantes

Function ListingRepSauve(RepSauvegarde : String) : Integer;
Var Chemin 	: String;
		Extend	: String;
    DateFile: TDateTime;
    Info		: TSearchRec;
begin

  { Pour être sur que la barre oblique finisse le nom du chemin }
  RepSauvegarde := IncludeTrailingBackslash(RepSauvegarde);

  Ok_Sauve := True;
  Result := 0;

  if StConnection.ZipFile then
     Extend := '*.Z*'
  else
	   Extend := '*.B*';

  { Recherche de la première entrée du répertoire }
  If FindFirst(RepSauvegarde + Extend, faAnyFile,Info)=0 Then
     Begin
     Repeat
     If Not ((Info.Attr And faDirectory)=0) Then // On ne gère que les fichiers si 0 ce sont les répertoires
     Else
        begin
        //Vérification si fichier trouvé = fichier de sauvegarde (Bak ou Zip)
        Extend := ExtractFileExt(Info.FindData.CfileName);
        if (Extend = '.BAK') Or (Extend = '.ZIP') then
           Begin
	         if GetFileTimes(RepSauvegarde + Info.FindData.CfileName, DateFile) then
              if DateToStr(DateFile) = DateToStr(Now) then Ok_Sauve := False;
           end;
        Result := Result + 1;
        end;
     Until FindNext(Info)<>0; // Il faut ensuite rechercher l'entrée suivante }
     { Dans le cas ou une entrée au moins est trouvée il faut }
     { appeler FindClose pour libérer les ressources de la recherche }
     FindClose(Info);
     End;

end;

Function CompressionFichier(StConnection : ParamConnection;MessAttente : TSplash_S1) : Boolean;
var Zip 	: TOZ;
    PathZip : String;
    FileZip	: String;
    PathSave: String;
    ok_Zip 	: boolean;
Begin
  //
	MessAttente.Label1.Caption := Format(TexteMessage[16],[TitreMess]);
	MessAttente.Refresh;
  //
  PathZip := IncludeTrailingBackslash(StConnection.PathSauve);
  FileZip := StConnection.StBase + '.ZIP';
  PathSave:= PathZip + StConnection.FichierSauve;

  //On compresse le fichier de sauvegarde dans le répertoire
  Result := False;
  Zip := TOZ.Create;

  try
    Zip.OpenZipFile(PathZip + FileZip, moCreate);
    Zip.ActiveSwitchDisqueDefault(False, False, True, False);
    Zip.ZipFileWithPath := True;
    Zip.NiveauDeCompression := 9;
    // test validité
    if zip.OpenSession(osAdd) then
       begin
        // ajout fichier
        Zip.ProcessFile(PathSave);
        Zip.CloseSession;
        if FileExists(PathSave) then Deletefile(PathSave);
        Result := True;
      end
      else
        Zip.CancelSession;
    finally
      Zip.Free;
    end;

end;

//
function GetFileTimes (const FileName: string; var DateAcces : TDateTime) : Boolean;
var h							: THandle;
    Info1					: TFileTime;
    Info2					: TFileTime;
    Info3					: TFileTime;
    SysTimeStruct	: SYSTEMTIME;
    TimeZoneInfo	: TTimeZoneInformation;
    Bias					: Double;
begin

	//Info1 : Date de Création
  //Info2 : Date de dernier Accès
  //Info3 : Date de Dernière Modif

  Result := false;

  Bias   := 0;
  h      := FileOpen(FileName, fmOpenRead or fmShareDenyNone);

  if h > 0 then
  	 begin
        try
        	if GetTimeZoneInformation(TimeZoneInfo) <> $FFFFFFFF then
         	  Bias := TimeZoneInfo.Bias / 1440; // 60x24
        	GetFileTime(h, @Info1, @Info2, @Info3);
        	if FileTimeToSystemTime(Info2, SysTimeStruct) then
         	  DateAcces := SystemTimeToDateTime(SysTimeStruct) - Bias;
        	Result := True;
        finally
        	FileClose(h);
        end;
     end;

end;

procedure ChargeEcran(Ecran : TForm; TobParam : Tob);
var indice : integer;
		NomToFind : string;
begin

	for indice := 0 to ecran.ControlCount -1 do
     begin
  	 NomToFind := ecran.Controls [Indice].Name;
     if (Pos(UpperCase(ecran.controls[indice].ClassType.className),UpperCase(TypeChampsTraitable))=0) then continue;
     if (TWinControl(ecran.controls[indice]).ControlCount  > 0) and
        (ecran.controls[indice].ClassType.className<>'THMultiValComboBox') and
//        (ecran.controls[indice].ClassType.className<>'THDBMultiValComboBox') and
        (ecran.controls[indice].ClassType.className<>'THValComboBox') and
        (ecran.controls[indice].ClassType.className<>'TCheckBox') and
        (ecran.controls[indice].ClassType.className<>'THEdit') and
        (ecran.controls[indice].ClassType.className<>'TEdit') and
        (ecran.controls[indice].ClassType.className<>'THNumEdit') and
//        (ecran.controls[indice].ClassType.className<>'THDBSpinEdit') and
        (ecran.controls[indice].ClassType.className<>'THSpinEdit') and
        (ecran.controls[indice].ClassType.className<>'THCritMaskEdit') then
    	  DescentChargeUnNiveau (TWinControl(ecran.controls[indice]), TobParam)
     else
      	ChargeComponent (TWinControl(ecran.controls[indice]), TobParam);
     end;


end;

procedure DescentChargeUnNiveau (Composant : TWinControl; TobParam : Tob);
var indice : integer;
		NomToFind : string;
begin

	for indice := 0 to composant.Controlcount -1 do
     begin
  	 NomToFind := Composant.Controls [Indice].Name;
     if (Pos(UpperCase(Composant.Controls [Indice].ClassType.className),UpperCase(TypeChampsTraitable))=0) then continue;
  	 if (TwinControl(composant.controls[indice]).ControlCount > 0 ) and
        (composant.controls[indice].ClassType.className<>'THMultiValComboBox') and
//        (composant.controls[indice].ClassType.className<>'THDBMultiValComboBox') and
        (composant.controls[indice].ClassType.className<>'THValComboBox') and
        (composant.controls[indice].ClassType.className<>'TCheckBox') and
        (composant.controls[indice].ClassType.className<>'THEdit') and
        (composant.controls[indice].ClassType.className<>'TEdit') and
        (composant.controls[indice].ClassType.className<>'THNumEdit') and
        (composant.controls[indice].ClassType.className<>'THSpinEdit') and
//        (composant.controls[indice].ClassType.className<>'THDBSpinEdit') and
        (composant.controls[indice].ClassType.className<>'THCritMaskEdit') then
    	  DescentChargeUnNiveau (TWinControl(composant.Controls[indice]), TobParam)
     else
    	  ChargeComponent (TWinControl(composant.Controls[indice]), TobParam);
     end;

end;

procedure ChargeComponent (composant : TWinControl; LaTob : Tob);
var NomToFind 	: string;
		TheTOBValue : TOB;
    SpinVal	  	: Integer;
    Value				: Variant;
begin

  NomToFind := composant.Name;

  TheTOBValue := LaTob.FindFirst(['SOC_NOM'],[NomToFind],true);

  if TheTobValue = Nil then exit;

  Value := TheTOBValue.GetValue('SOC_DATA');

	if (composant IS THSpinEdit) then
  Begin
     SpinVal := Value;
     THSpinEdit(composant).Value := SpinVal;
  end else if (composant IS TEdit) then
  begin
     TEdit(composant).Text := Value
  end else if (composant IS THEdit) Then
  begin
     THEdit(composant).Text := Value
	end Else if (composant IS THNumEdit) Then
  begin
  	 THNumEdit(composant).Text := Value
  end Else if (composant IS THValComboBox ) Then
  begin
	   THValComboBox(composant).Value := Value
  end Else if (composant IS THMultiValComboBox ) Then
  begin
	   THMultiValComboBox(composant).Value := Value
{  Else if (composant IS THDBMultiValComboBox) then
	   THDBMultiValComboBox(composant).Value := Value
}
  end else if (composant IS THCritMaskEdit ) then
  begin
     if (composant is THCritMaskEdit) and (ThCritMaskEdit(composant).OpeType = otDate) then
        THCritMaskEdit(composant).Text := DateToStr(Value)
     else
        THCritMaskEdit(composant).Text := Value
  end else if (composant IS TCheckBox) Then
  Begin
     If TheTOBValue.getValue('SOC_DATA')='X' THEN
        TCheckBox(composant).Checked := True
     Else
	      TCheckBox(composant).Checked := False;
  end else if (composant IS THCheckBox) then
  Begin
     If TheTOBValue.getValue('SOC_DATA')='X' THEN
        THCheckBox(composant).Checked := True
     Else
	      THCheckBox(composant).Checked := False;
  end;
  {
  else if composant IS THDBSpinEdit then
     Begin
     SpinVal := Value;
     THSpinEdit(composant).Value := SpinVal;
     end;
  }

end;

procedure StockeInfos(Ecran : TForm; TobParam : Tob);
var indice 		: integer;
		NomToFind : string;
begin

	for indice := 0 to ecran.ControlCount -1 do
  		begin
    	NomToFind := ecran.Controls [Indice].Name;
      if (Pos(upperCase(ecran.controls[indice].ClassType.className),UpperCase(TypeChampsTraitable))=0) then continue;
      if (TWinControl(ecran.controls[indice]).ControlCount  > 0) and
//         (ecran.controls[indice].ClassType.className<>'THDBMultiValComboBox') and
         (ecran.controls[indice].ClassType.className<>'THMultiValComboBox') and
         (ecran.controls[indice].ClassType.className<>'THValComboBox') and
         (ecran.controls[indice].ClassType.className<>'TCheckBox') and
         (ecran.controls[indice].ClassType.className<>'THEdit') and
         (ecran.controls[indice].ClassType.className<>'THNumEdit') and
         (ecran.controls[indice].ClassType.className<>'TEdit') and
         (ecran.controls[indice].ClassType.className<>'THSpinEdit') and
//         (ecran.controls[indice].ClassType.className<>'THDBSpinEdit') and
         (ecran.controls[indice].ClassType.className<>'THCritMaskEdit') then
		     DescentStockeUnNiveau (TWinControl(ecran.controls[indice]), TobParam)
      else
         StockeComponent(TWinControl(ecran.controls[indice]), Tobparam);
  	  end;

end;

procedure DescentStockeUnNiveau (Composant : TWinControl; TobParam : Tob);
var indice 		: integer;
		NomToFind : string;
begin

	for indice := 0 to composant.Controlcount -1 do
      begin
  	  NomToFind := Composant.Controls [Indice].Name;
      if (Pos(UpperCase(Composant.Controls [Indice].ClassType.className),UpperCase(TypeChampsTraitable))=0) then continue;
  	  if (TwinControl(composant.controls[indice]).ControlCount > 0) and
         (Composant.controls[indice].ClassType.className<>'THMultiValComboBox') and
//         (Composant.controls[indice].ClassType.className<>'THDBMultiValComboBox') and
         (Composant.controls[indice].ClassType.className<>'THValComboBox') and
         (Composant.controls[indice].ClassType.className<>'TCheckBox') and
         (Composant.controls[indice].ClassType.className<>'THEdit') and
         (Composant.controls[indice].ClassType.className<>'TEdit') and
			   (Composant.controls[indice].ClassType.className<>'THNumEdit') and
         (Composant.controls[indice].ClassType.className<>'THSpinEdit') and
//         (Composant.controls[indice].ClassType.className<>'THDBSpinEdit') and
         (Composant.controls[indice].ClassType.className<>'THCritMaskEdit') then
     	   DescentStockeUnNiveau (TWinControl(composant.Controls[indice]), TobParam)
      else
         StockeComponent (TWinControl(composant.Controls[indice]), TobParam);
	    end;

end;


procedure StockeComponent (composant : TWinControl; TobParam : Tob);
var	NomToFind 	: String;
		Value				: string;
		TheTOBValue : TOB;
begin

  NomToFind := composant.name;

  TheTOBValue := TobParam.findFirst(['SOC_NOM'],[NomToFind],true);

  if TheTOBValue = nil then exit;

  if Composant IS THSpinEdit then
  begin
  	Value := IntToStr(THSpinEdit(Composant).value);
  end else if (Composant IS TEdit) Then
  begin
  	Value := TEdit(composant).Text
  end Else if (composant IS THEdit) Then
  begin
  	Value := THEdit(composant).Text
  end Else if (composant IS THNumEdit) then
  begin
  	Value := THNumEdit(composant).Text
  end Else if (Composant IS THMultiValComboBox ) Then
  begin
  	Value := THMultiValCombobox(composant).Value
  end Else if (Composant IS THValComboBox ) Then
  begin
  	Value := THValCombobox(composant).Value
  end Else if (composant IS THCritMaskEdit ) then
  begin
    if (composant is THCritMaskEdit) and (ThCritMaskEdit(composant).OpeType = otDate) then
    	Value := FloatToStr(StrToDate(THCritMaskEdit(Composant).Text))
  Else
  	Value := THCritMaskEdit(Composant).Text
  end else if Composant IS TCheckBox then
  begin
    if TCheckBox(composant).Checked Then Value := 'X'
    															  Else Value := '-'
  end else if Composant IS THCheckBox then
  begin
    if THCheckBox(Composant).Checked Then Value := 'X'
                                     Else Value := '-'
  end;
  (*
  else if Composant IS THDBSpinEdit then
  Value := IntToStr(THDBSpinEdit(Composant).value);
  *)

  TheTOBValue.PutValue('SOC_DATA',Value);

end;

Procedure EnvoieMail(ParametreMail : ParamMail);
Var Corps			 : HTStringList;
    ResultMail : TResultMailForm;
    Email,CopieA : string;
Begin

  if ParametreMail.EMail = '' then
     Begin
     Pgibox('Adresse Mail à blanc', 'Paramétres Généraux');
     exit;
     end;

  if pos('@', ParametreMail.Email) = 0 then
     Begin
     Pgibox('Adresse Mail : ' + ParametreMail.Email + ' Invalide', 'Paramétres Généraux');
     exit;
     end;

  TRY
     Corps := HTStringList.Create;
     Corps.Add('');
     Corps.Text := ParametreMail.MessMail;
     Email := ParametreMail.EMail;
     CopieA := ParametreMail.CopieA;
     //Envoie du mail
     ResultMail := AglMailForm(ParametreMail.Sujet, Email,CopieA, Corps,ParametreMail.Fichier,false);
     if ResultMail = rmfOkButNotSend then SendMail(ParametreMail.Sujet,Email,CopieA,Corps,ParametreMail.Fichier,True,1);
  FINALLY
     Corps.Free;
  END;


end;


end.
