{***********UNITE*************************************************
Auteur  ...... : Xavier PERSOUYRE
Créé le ...... : 09/08/2001
Modifié le ... :   /  /
Description .. : Dans cette unité, on trouve LA FONCTION qui sera
Suite ........ : appelée par l'AGL pour lancer les options.
Suite ........ : Ainsi que la gestion de l'HyperZoom, de la gestion de
Suite ........ : l'arobase, ...
Suite ........ : C'est aussi dans cette unité que l'on défini le fichier ini
Suite ........ : utilisé, le nom de l'application, sa version, que l'on lance la
Suite ........ : sérialisation, les différentes possibilités d'action sur la mise à
Suite ........ : jour de structure, ...
Mots clefs ... : IMPORTANT;STRCTURE;MENU;SERIALISATION
*****************************************************************}
unit Menucom;
{$ifndef COMSX}
ZZZZZZZZ // pour être mettre dans le projet COMSX
{$endif}

interface
Uses
  Windows, Messages,
  HEnt1,EntPGI,
{$ifdef eAGLClient}
  MenuOLX,MaineAGL,eTablette, UtilEagl,CEGIDIEU,
{$else}
  MenuOLG, FE_Main,
  UGedFiles,
  YNewDocument,
{$endif eAGLClient}
  BobGestion,
  CPGENERAUX_TOM, CPTIERS_TOM, CPJOURNAL_TOM, CPSECTION_TOM,
  Forms,sysutils,HMsgBox, Classes, Controls, HPanel,
  hctrls, MajTable, Utilsoc,
  ExtCtrls, inifiles, UTOB, galOutil, UNeActions,
  UNetListe, ULibCpContexte, uTofConsEcr, Expecc, ImgList, utilPGI; // AJOUT SUR 7XX

Procedure InitApplication ;

type
  TFMenuDisp = class(TDatamodule)
    ImageList1  : TImageList;
    Timer1      : TTimer;
    procedure Timer1Timer(Sender: TObject);
    end ;

Var FMenuDisp : TFMenuDisp ;
    Soc,User,Motdepass    : string;
    OkExport,Minimise : Boolean;

procedure LanceMenuAuto;

implementation

{$R *.DFM}

Uses
  {$IFDEF MODENT1}
  CPVersion,
  CPTypeCons,
  {$ENDIF MODENT1}
     LicUtil,Ent1,
     ParamSoc,
     UTOFMULPARAMGEN,
     UTOFCPMULMVT,
     Resulexo,
     UAssistComsx, UAssistImport,
     CPJALIMPORT_TOF;
{$IFNDEF CERTIFNF}
Procedure GoMenu(SuperUser : Boolean) ;
BEGIN
if (V_PGI.ModePCL='1') then
begin
     if GetParamSocSecur ('SO_CPLIENGAMME', 'S5') <> 'S1' then
        FMenuG.SetFunctions([TraduireMemoire('Recevoir'),
                       TraduireMemoire('Envoyer'),
                       TraduireMemoire('Réinitialisation des blocages')
                      ],
                      [45100, 45000,  45200],
                      [46, 40, 126])
     Else
        FMenuG.SetFunctions([TraduireMemoire('Recevoir'),
                       TraduireMemoire('Envoyer'),
                       TraduireMemoire('Mouvements modifiés'),
                       TraduireMemoire('Mouvements Lettrés'),
                       TraduireMemoire('Réinitialisation des blocages')
                      ],
                      [45100,45000, 46000, 47000, 45200],
                      [46, 40, 41, 45, 126]) ;
end
else
begin
  FMenuG.SetFunctions([TraduireMemoire('Recevoir'),
                       TraduireMemoire('Envoyer'),
                       TraduireMemoire('Encours-client'),
                       TraduireMemoire('Réinitialisation des blocages')
                      ],
                      [45100, 45000, 48000, 45200],
                      [46, 40, 47, 126])
end;
END ;

{$ELSE}

Procedure GoMenu(SuperUser : Boolean) ;
BEGIN
if (V_PGI.ModePCL='1') then
begin
     if GetParamSocSecur ('SO_CPLIENGAMME', 'S5') <> 'S1' then
        FMenuG.SetFunctions([TraduireMemoire('Recevoir'),
                       TraduireMemoire('Envoyer'),
                       TraduireMemoire('Rapport d''intégration'),
                       TraduireMemoire('Réinitialisation des blocages')
                      ],
                      [45100, 45000, 45300, 45200],
                      [46, 40, 128, 126])
     Else
        FMenuG.SetFunctions([TraduireMemoire('Recevoir'),
                       TraduireMemoire('Envoyer'),
                       TraduireMemoire('Mouvements modifiés'),
                       TraduireMemoire('Mouvements Lettrés'),
                       TraduireMemoire('Rapport d''intégration'),
                       TraduireMemoire('Réinitialisation des blocages')
                      ],
                      [45100,45000, 46000, 47000, 45300, 45200],
                      [46, 40, 41, 45, 128, 126]) ;
end
else
begin
  FMenuG.SetFunctions([TraduireMemoire('Recevoir'),
                       TraduireMemoire('Envoyer'),
                       TraduireMemoire('Encours-client'),
                       TraduireMemoire('Rapport d''intégration'),
                       TraduireMemoire('Réinitialisation des blocages')
                      ],
                      [45100, 45000, 48000, 45300, 45200],
                      [46, 40, 47, 128, 126])
end;
END ;
{$ENDIF}

FUNCTION CHARGEMAGHALLEYGG : Boolean ;
Var SuperUser : Boolean ;
BEGIN
Result:=CHARGEMAGHALLEY ;
SuperUser:=(V_PGI.PassWord=CryptageSt(DayPass(Date))) ;
GoMenu(SuperUser) ;
ChargeMenuPop(integer(hm2),FMenuG.DispatchX) ;

END ;

//exemple rem c:\pgi00\app\COMSX.EXE /TRF=répertoire;Fichier d'import;IMPORT;SOCIETE;UTILISTATEUR;Adresse Email;TYPE écriture;TYPE écriture;Fichier RAPPORT;Gestion DOUBLON;Controle Paramsoc; Mode SANSECHEC;Option Minimisé
//exemple C:\PGI00\APP\COMSX.EXE /TRF=C:\tmp\FFF\;JOURNALGI.TRA;IMPORT;SECALDOS;CEGID;mentressangle@cegid.fr;N;N;RAPPORT.LOG;FALSE;FALSE;SANSECHEC;Minimized

Procedure ScruteRepertoire (StArg : string; Attributs: integer; Commande : string);
var
  FichierTrouve            : string;
  Resultat,NoSEqNet        : Integer;
  SearchRec                : TSearchRec;
  St                       : string;
  Repertoire               : string;
  filtre,tmp               : string;
  Minim,Echec              : string;
  FileNonZip,FileZip       : string;
  FicIni                   : TIniFile;
  sTempo                   : string;
  StArgum                  : string;
  OutTOBNet                : TOB;
  function RenseigneCommande : Boolean;
  begin
          if (pos('/INI=', StArg) <> 0) then
          begin
                tmp    := ReadTokenPipe(StArg, ';');
                sTempo := Copy (tmp, pos('/INI=', tmp)+5, length(tmp));
                if not FileExists(sTempo) then
                begin
                    PGIInfo ('Le fichier .ini'+ sTempo+ ' n''existe pas','');
                    Result := FALSE;
                    exit;
                end;
                Minim                    := '';
                if Minimise then Minim   := 'Minimized';
                FicIni                   := TIniFile.Create(sTempo);
                Repertoire               := FicIni.ReadString (Commande, 'REPERTOIRE', '');
                Filtre                   := FicIni.ReadString (Commande, 'NOMFICHIER', '');
                StArgum := Minim;
                FicIni.free;
          end
          else
          begin
                if pos('Minimized', StArg) <> 0 then Minim := 'Minimized';
                if pos('SANSECHEC', StArg) <> 0 then Echec := 'SANSECHEC';
                St            :=UpperCase(Trim(ReadTokenPipe(StArg,'=')));
                Repertoire    := ReadTokenPipe(StArg, ';');  // répertoire
                Filtre        := ReadTokenPipe(StArg, ';');   // nom du fichier
                ReadTokenPipe(StArg, ';');
                tmp           := ReadTokenPipe(StArg, ';'); // societe
                tmp           := ReadTokenPipe(StArg, ';'); // user
                StArgum       := ReadTokenPipe(StArg, ';')
                 + ';' + ReadTokenPipe(StArg, ';')    // type d'écriture reçu  exemple simule
                 + ';' + ReadTokenPipe(StArg, ';') // type d'écriture à intégrer exemple Normal
                 + ';' + Minim
                 + ';' + ReadTokenPipe(StArg, ';')  // rapport
                 + ';' + Echec                    // sans echec
                 + ';' + ReadTokenPipe(StArg, ';')  // doublon
                 + ';' + ReadTokenPipe(StArg, ';');  // Controleparam
          end;
          Result := TRUE;
  end;
      procedure EcrireDansfichierCom (Fichier : string; Chaine : string);
      var
        F: TextFile;
      begin
      {$I-}
        AssignFile(F, Fichier);
        Rewrite(F) ;
        writeln(F, 'Compte rendu d''import/export');
        writeln(F, Chaine);
        CloseFile(F);
      {$I+}
      end;
begin
  if not RenseigneCommande then exit;
  if (Repertoire = '') and (IsDossierNetExpert(V_PGI.NoDossier, NoSEqNet)) then
  begin
       if pos ('.NSV', UpperCase (Filtre)) <> 0 then
           OutTOBNet := InterrogeCorbeil('PAIE', 'NSV')
       else
           OutTOBNet := InterrogeCorbeil('COMPTA', 'TRA');
       if OutTOBNet <> nil then
       begin
           if OutTOBNet.GetValue('ERROR') = 0 then
           begin
              ImportDonnees(Filtre+';'+Minim, FALSE, OutTOBNet, sTempo, Commande);
           end
           else
              EcrireDansfichierCom ('ListeCom'+ V_PGI.NoDossier+'.ERR',OutTOBNet.GetValue('ERRORLIB'));
           OutTOBNet.free;
       //    NelibereTob;
       end;
          exit;
  end;
  if not FileExists(Repertoire+filtre) then  // pour S1 car il risque d'envoyer des fichiers zippés
  begin
                  FileNonZip := filtre;
                  FileZip := ReadTokenPipe (FileNonZip, '.')+'.ZIP';
                  if FileExists(Repertoire+FileZip) then filtre := FileZip;
  end;
  if Repertoire[length (Repertoire)] <> '\' then Repertoire := Repertoire + '\';

  Resultat := FindFirst (Repertoire + filtre, Attributs, SearchRec);
  if Resultat <> 0 then
                      EcrireDansfichierCom (Repertoire +'ListeCom'+ V_PGI.NoDossier+'.ERR', Repertoire + filtre + ' inexistant');
  while Resultat = 0 do
  begin
    Application.ProcessMessages;
    if ((SearchRec.Attr and faDirectory) <= 0) then // On a trouvé un Fichier (et non un dossier)
      FichierTrouve := SearchRec.Name
    else begin Resultat:= FindNext (SearchRec); continue; end;
    ImportDonnees (Repertoire+';'+FichierTrouve+';'+StArgum, Minimise, nil, sTempo, Commande);
    Resultat:= FindNext (SearchRec);
  end;

  FindClose (SearchRec);// libération de la mémoire
end;

{***********A.G.L.***********************************************
Auteur  ...... : Xavier PERSOUYRE
Créé le ...... : 09/08/2001
Modifié le ... :   /  /
Description .. : Cette fonction est appellée par l'AGL à chaque sélection
Suite ........ : d'une option de menu, en lui indiquant le TAG du menu en
Suite ........ : question. Ce qui déclenche l'action en question.
Suite ........ : L'AGL lance aussi cette fonction directement afin d'offrir à
Suite ........ : l'application la possibilité d'agir avant ou après la connexion,
Suite ........ : et avant ou après la déconnexion.
Suite ........ : Cette fonction prend aussi en paramètre retourForce et
Suite ........ : SortieHalley. Si RetourForce est à True, alors l'AGL
Suite ........ : remontera au niveau des modules, si SortieHalley est à
Suite ........ : True, alors ...
Mots clefs ... : MENU;OPTION;DISPATCH
*****************************************************************}
type
  TComsxContext = class (TCPContexte)
         //   constructor Create( MySession : TISession ) ; override ;
  end;

function AvecLigneCommande : boolean;
var i : integer;
    St : string;
begin
  Result := False;
  for i:=1 to ParamCount do
  begin
    St := ParamStr(i);
    if ((pos('/INI', St) <> 0) or (pos('/TRF', St) <> 0)) then
    begin
      Result := True;
      break;
    end;
  end;
end;

Procedure DispatchCom ( Num : Integer ; PRien : THPanel ; var retourforce,sortiehalley : boolean);
var
St,opt,Nom,Arg   : string;
Numopt,i         : integer;
Okarg            : Boolean;
Tmp,Fichier      : string;
FicIni           : TIniFile;
ListeFile        : string;
Commande         : string;
sDom             : string;
//pContext         : TCPContexte;
//pContext   : TComsxContext;
BEGIN
     Numopt := Num;
   case Numopt of
     16 :
     begin
       BOB_IMPORT_PCL_STD ('CCS5','Paramétrage Comptabilité');
(*
      {$IFNDEF EAGLCLIENT}
        { Importation des bobs CCS5 }
        BOB_IMPORT_PCL('CCS5');
      {$ENDIF}
*)
     end;
     10 :
     begin
{$IFNDEF EAGLCLIENT}
{$IFDEF SCANGED}
// AJOUT SUR 7XX  il n'y a pas de séria sur la Comsx on ne test pas VH^.OkModGed
            if V_PGI.RunFromLanceur then
              InitializeGedFiles(V_PGI.DefaultSectionDbName, nil)
            else
              InitializeGedFiles(V_PGI.DbName, nil);
{$ENDIF}
{$ENDIF}
          Okarg := FALSE;
          ChargeNoDossier;  // AJOUT SUR 7XX
          //pContext   := TComsxContext(TComsxContext.GetCurrent) ;
          //pContext := TCPContexte.GetCurrent;
           //PGIINfo(GetEncours.Code) ;
           //PGIINFO(IntTostr(GetInfoCpta(fbGene).lg));
          if ParamCount > 0 then
          begin
              for i:=1 to ParamCount do
              BEGIN
                  St := ParamStr(i);
                  Nom:=UpperCase(Trim(ReadTokenPipe(St,'='))) ;
                  Arg := St;
                  if (Nom = '/TRF') or (Nom = '/INI') then
                  begin
                        OkArg := TRUE;
                        While St <> '' do
                        begin
                              Opt := ReadTokenPipe(St, ';');
                              (*Test2=societe; CEGID=user;Email;fichier tob tier;-si ecriture sur un seul etablissement
                              S5=type; JRL=nature;fichier TRA; exercice;date1;date2; journaux
                              /TRF=EXPORT;TEST2;CEGID;mentressangle@cegid.fr;C:\import\journal.txt;X;S5;JRL;C:\import\pgi00001.TRA;001;01/01/2003;31/12/2003;TYPE=[S;N];JOURNAL=[VP;AP];Fichierrapport.txt
                              /TRF=EXPORT;000001;CEGID;mentressangle@cegid.fr;;X;SISCO;JRL;C:\import\SI00001.TRT;;01/01/2002;31/12/2002;;JOURNAL=[VEN];Fichierrapport.txt
                              *)
                              if Opt = 'EXPORT' then
                              begin
                                 ExportDonnees(Arg, Minimise);
                                 OkExport := FALSE;
                                 break;
                              end;
                                (* exemple
                                /TRF=c:\sav\;pgi00001.TRA;IMPORT;SOCIETE;USERS;Email;Typecrrecu;typecrremplace;Calcul de numéro de pièce(O/N); Minimiser l'écran; SANSECHEC=pour ne pas créer le compte sinon rien
                                /TRF=C:\tmp\;JOURNAL.TRA;IMPORT;10000;CEGID;Email;S;N;Minimized;SANSECHEC
                                ou /USER=CEGID /PASSWORD=CEGID /DOSSIER=10000 /INI=C:\tmp\10000\COMSX.INI;IMPORT
                                ou pour TRFS5 /USER=CEGID /PASSWORD=46171FC5C3 /DATE=08/01/2004 /DOSSIER=###DOSSIER### /MAJSTRUCTURE=TRUE /TRF=C:\tmp\00003\;PGBAILL_??.TRA;IMPORT;Minimized /APPSIG:PGI-ID0001CIMP_CP:00003 /APPPAR:$T
                                *)

                              if (Opt = 'IMPORT') then
                              begin
                                    if (pos('.INI', UPPERCASE(ParamStr(i))) <> 0) then
                                    begin
                                              Fichier    := ParamStr(i);
                                              tmp := ReadTokenPipe(Fichier, ';');
                                              Fichier := Copy (tmp, pos('/INI=', tmp)+5, length(tmp));
                                              if FileExists(Fichier) then
                                              begin
                                                    FicIni       := TIniFile.Create(Fichier);
                                                    ListeFile    := FicIni.ReadString ('COMMANDE', 'LISTECOMMANDE', '');
                                                    FicIni.Free;
                                              end;
                                   end;
                                   if ListeFile = '' Then
                                   begin
                                      Commande := 'COMMANDE';
                                      ScruteRepertoire (ParamStr(i), FaDirectory + faHidden + faSysFile, Commande);
                                   end
                                   else
                                   begin
                                       Commande := ListeFile;
                                       While Commande <> '' do
                                       begin
                                            Commande    := ReadTokenPipe(ListeFile, ';');
                                            if Commande <> '' then
                                                        ScruteRepertoire (ParamStr(i), FaDirectory + faHidden + faSysFile, Commande);
                                       end;
                                   end;
                                   OkExport := FALSE;
                                 break;
                              end;
                              if Opt = 'ENCOURS' then  // fiche 10561
                              begin
                                 ExportEnCours(Arg);
                                 OkExport := FALSE;
                                 break;
                              end;

                        end;
                  end;
              END;
              if OkArg then
              begin
                  // Fermeture de l'application
                  Application.ProcessMessages;
                  if FMenuG <> nil then
                  begin
                      FMenuG.Quitter;
                      {$IFNDEF EAGLCLIENT}
                      // AJOUT SUR 7XX if VH^.OkModGed then
                      FinalizeGedFiles;
                      {$ENDIF}
                      Exit;
                  end;
              end;
          end;

     end;
     11 : ;
     12 : BEGIN
             // Interdiction de lancer en direct
          if (Not V_PGI.RunFromLanceur) and (V_PGI.ModePCL='1') and (not AvecLigneCommande) then
          begin
            FMenuG.FermeSoc;
            FMenuG.Quitter;
            exit;
          end;
          if V_PGI.ModePCL = '1' then
          begin
            sDom := '00280011' ;

            VH^.SerProdSCAN         := '00610080';
            FMenuG.SetSeria(sDom, [VH^.SerProdSCAN],
                                  ['Cegid Expert SCAN']) ;
          end;

          RenseignelaSerie(ExeCOMSX) ;
          VH^.RecupComSx := TRUE;
          VH^.ModeSilencieux := True;
{$ifndef eAGLClient}
          FMenuG.OkVerifStructure:=FALSE ;
{$endif}
          if ((V_PGI.ModePCL='1')) then V_PGI.PGIContexte:=V_PGI.PGIContexte+[ctxPCL] ;
          END ;
     13 :
{$IFNDEF EAGLCLIENT}
         begin
          FinalizeGedFiles;
        end
{$ENDIF}
      ;

     15 : ;
     100 : ;
     4006 : AGLLanceFiche ('YY','YYDEFPRT','','','TYPE=E');

      //Menu Pop Compta
     // export
     45000 :
        ExportDonnees('');
     // import
     45100 :
        ImportDonnees ('');
     //   AGlLanceFiche ('CP','IMPORTFICHECOM','','','');
     // Liste des mouvements modifiés
     46000 :
        AGlLanceFiche ('CP','MULMVTMODIFIES','','','');

     // liste des mouvement modifies et pointées
    47000 : AGlLanceFiche ('CP','MULMVTPOINTE','','','');
    // Encours-client
    48000 : ExportEnCours;
    25710: CCLanceFiche_MULGeneraux('P;7112000'); // Comptes généraux
    25720: CPLanceFiche_MULTiers('P;7145000'); // Comptes auxiliaires
    25730: CPLanceFiche_MULSection('P;7178000'); // Sections analytiques
    25740: CPLanceFiche_MULJournal('P;7205000'); // Journaux
    25750: MultiCritereMvt(taConsult, 'N', False); // Consultation des écritures
    25760: ResultatDeLexo;
    25755 : OperationsSurComptes('','','') ;
    45200 :
        begin
        PGIInfo ('Attention, êtes vous certain de vouloir  lancer ce traitement ? '+#10#13+ ' Assurez vous qu'' aucun autre utilisateur n''est connecté à ce module en cours de réception.');
        if PGIAsk ('Confirmez-vous le traitement ?', 'Réinitialisation des blocages')=mrYes then
            ExecuteSQL ('DELETE FROM COURRIER WHERE MG_TYPE=5555');
        end;
    45300 : CPLanceFiche_JournalDesimports('');
    25770 :
    begin
          ParamSociete(False,BrancheParamSocAVirer,'SCO_PARAMETRESGENERAUX','',RechargeParamSoc,ChargePageSoc,SauvePageSoc,InterfaceSoc,1105000) ;
          CHARGEMAGHALLEY;
    end;
    else HShowMessage('2;?caption?;Fonction non disponible : ;W;O;O;O;',TitreHalley,IntToStr(Num)) ;
  end ;
END ;

{***********A.G.L.***********************************************
Auteur  ...... : Xavier PERSOUYRE
Créé le ...... : 09/08/2001
Modifié le ... :   /  /
Description .. : Cette procédure permet d'initiliaser certaines référence de
Suite ........ : fonction, les modules des menus gérés par l'application, ...
Suite ........ :
Suite ........ : Cette procédure est appelée directement dans le source du
Suite ........ : projet.
Mots clefs ... : INITILISATION
*****************************************************************}

Procedure AfterProtec ( sAcces : String ) ;
BEGIN
VH^.OkModCompta:=True ; VH^.OkModBudget:=TRUE ; VH^.OkModImmo:=TRUE ;
//V_PGI.Monoposte:=TRUE ;
V_PGI.VersionDemo:=FALSE ;
if (V_PGI.ModePCL='1') then
VH^.OkModSCAN         :=(sAcces[1]='X');
END ;


procedure InitApplication ;
var
I         : integer;
St,Nom    : string;
begin
     { Référence à la fonction PRINCIPALE qui permet de lancer les actions en fonction de la
       sélection d'une option des menus.
       }
     FMenuG.OnDispatch:=DispatchCom ;

     { Référence à une fonction qui est lancée après la connexion à la base et le chargement du dictionnaire }
     FMenuG.OnChargeMag:=ChargeMagHalleyGG ;

     { Référence à une fonction qui est lancée avant la mise à jour de structure }
     FMenuG.OnMajAvant:=Nil ;

{$IFNDEF EAGLCLIENT}
     { Référence à une fonction qui est lancée pendant la mise à jour de structure }
     FMenuG.OnMajpendant:=Nil ;
     FMenuG.Timer1.Enabled := FALSE;
{$ENDIF}

     { Référence à une fonction qui est lancée après la mise à jour de structure }
     FMenuG.OnMajApres:=Nil ;
     FMenuG.OnChangeModule:=nil ;
     OkExport := FALSE;
     Minimise := FALSE;
     for i:=1 to ParamCount do
     BEGIN
                  St := ParamStr(i);
                  // A partir du bureau
                  if pos('###DOSSIER###', St) <> 0  then
                  begin
                       V_PGI.MultiUserLogin := TRUE;  break;
                  end;
                  Nom:=UpperCase(Trim(ReadTokenPipe(St,'='))) ;
                  if (Nom = '/TRF') or (Nom = '/INI') then
                  begin
                        if pos('Minimized', St) <> 0 then Minimise  := TRUE;
                        While St <> '' do
                        begin
                              Nom := ReadTokenPipe(St, ';');
                              if (Nom = 'EXPORT') or (Nom = 'IMPORT')
                              or (Nom = 'ENCOURS') then   // fiche 10561
                              begin
                                 Soc       :=ReadTokenPipe(St, ';') ;
                                 User      :=ReadTokenPipe(St, ';') ;
                                 Motdepass :=  DayPass(Date);
                                 OkExport  := TRUE;
                              end;

                        end;
                  end;
   end;
FMenuG.OnAfterProtec:=AfterProtec ;
RenseignelaSerie(ExeCOMSX) ;
{$IFDEF TESTSIC}
{$IFDEF EAGLCLIENT}
SaveSynRegKey('eAGLHost', 'CWAS-DEV3:80', true);
FCegidIE.HostN.Enabled := false;
{$ENDIF EAGLCLIENT}
{$ENDIF TESTSIC}
END ;

procedure DispatchTT(Num : Integer; Action : TActionFiche; Lequel,TT,Range : String) ;
begin
  case Num of
    // general
    1,7112 : FicheGene(Nil,'',LeQuel,Action,0);
    // tiers
    2,7145 : FicheTiers(Nil,'',LeQuel,Action,1) ;
    // journal
    4,7211 : FicheJournal(Nil,'',Lequel,Action,0) ;
    // Section
    71781 : FicheSection(nil,'A1',Lequel,Action,1);
    71782 : FicheSection(nil,'A2',Lequel,Action,1);
    71783 : FicheSection(nil,'A3',Lequel,Action,1);
    71784 : FicheSection(nil,'A4',Lequel,Action,1);
    71785 : FicheSection(nil,'A5',Lequel,Action,1);

    end ;

end ;

Procedure InitLaVariablePGI;
Begin
Apalatys:='CEGID' ;
Copyright:='© Copyright ' + Apalatys ;

HalSocIni:='CEGIDPGI.INI' ;
V_PGI.OutLook:=FALSE ;
V_PGI.VersionDemo:=FALSE ;
V_PGI.MenuCourant:=0 ;
V_PGI.VersionDEV:=TRUE ;
V_PGI.ImpMatrix := True ;
V_PGI.OKOuvert:=FALSE ;
V_PGI.Halley:=TRUE ;
V_PGI.OfficeMsg:=True ;
V_PGI.NumMenuPop:=27 ;
V_PGI.DispatchTT:=DispatchTT ;
V_PGI.ParamSocLast:=False ;
V_PGI.RAZForme:=TRUE ;
V_PGI.NoModuleButtons:=False ;
V_PGI.PGIContexte:=[ctxCompta] ;
V_PGI.BlockMAJStruct:=True ;
V_PGI.EuroCertifiee:=False ;

V_PGI.VersionReseau:=True ;
V_PGI.CegidAPalatys:=FALSE ;
V_PGI.CegidBureau:=TRUE ;
V_PGI.StandardSurDP:=True ;
V_PGI.MajPredefini:=False ;
V_PGI.MultiUserLogin:=False ;

V_PGI.NumMenuPop:=27 ;
V_PGI.OfficeMsg:=True ;
V_PGI.NoModuleButtons:=FALSE ;
V_PGI.NbColModuleButtons:=1 ;
V_PGI.LaSerie := S5;
RenseignelaSerie(ExeCOMSX) ;
if GetSynRegKey('Outlook',2,True)=2 then
  SaveSynRegKey('Outlook',False,True);
ChargeXuelib ;
end;

procedure TFMenuDisp.Timer1Timer(Sender: TObject);
begin
     if Minimise then FMenuG.WindowState  := wsMinimized;
     Timer1.Enabled := FALSE;
     V_PGI.MultiUserLogin := TRUE;
     if not OkExport then exit;
     if User = '' then exit;

     FMenuG.FDossier.ItemIndex := FMenuG.FDossier.Items.Indexof(Soc);
     FMenuG.FUser.text := User ;
     FMenuG.FPassword.Text := Motdepass;
     FMenug.bConnect.Click;
end;



procedure LanceMenuAuto;
var
Nom, St, Value     : string;
i,j                : integer;
Connect,B1,B2      : Boolean;
Num                : integer;
P2                 : THPanel ;
FRejet             : TextFile;
SRep               : string;
begin
      V_PGI.Debug:=FALSE ;
      V_PGI.Versiondev:=FALSE ;
      V_PGI.Synap:=FALSE ;
      VH^.GrpMontantMin:=0 ;
      VH^.GrpMontantMax:=1000000 ;
      V_PGI.DateEntree := Date ;
      VH^.Mugler:=FALSE ;
      V_PGI.Halley:=TRUE ;
      V_PGI.NumVersion:='3' ;
      V_PGI.Versiondev:=FALSE ;
      V_PGI.Synap:=FALSE ;
      VH^.GereSousPlan:=True ;
      V_PGI.Halley:=TRUE ;
      VH^.ModeSilencieux := True;
      V_PGI.MultiUserLogin := TRUE;
      if ((V_PGI.ModePCL='1')) then
           V_PGI.PGIContexte:=V_PGI.PGIContexte+[ctxPCL] ;
      Connect := FALSE;
      P2 := nil;
      for i:=1 to ParamCount do
      begin
        St:=ParamStr(i) ;
        Nom:=UpperCase(Trim(ReadTokenPipe(St,'='))) ;
        Value:=UpperCase(Trim(St)) ;
        //Paramètres de connexion
        if Nom='/USER'     then V_PGI.UserLogin :=Value;
        if Nom='/PASSWORD' then V_PGI.PassWord:=Value;
        if Nom='/DATE'     then BEGIN  V_PGI.DateEntree:=StrToDate(Value) ; END ;
        if Nom='/DOSSIER'  then
        BEGIN
             if V_PGI.PassWord = '000' then V_PGI.PassWord := CryptageSt(DayPass(Date));
             V_PGI.CurrentAlias := Value;
             V_PGI.DefaultSectionName := '';
{$IFNDEF EAGLCLIENT}
             j := pos('@', Value);
             if (j > 0) then
             begin
                  V_PGI.DefaultSectionName := Copy(Value, j + 1, 255);
                  V_PGI.CurrentAlias := Copy(Value, 1, j - 1);
                  if (j > 3) and (Copy(Value, 1, 2) = 'DB') then
                  V_PGI.NoDossier := Copy(Value, 3, j - 3); // sinon reste à '000000'
             end;
             V_PGI.RunFromLanceur := (V_PGI.DefaultSectionName <> '');
             Connect := ConnecteHalley(V_PGI.CurrentAlias,FALSE,@ChargeMagHalley,NIL,NIL,NIL);
{$ELSE}
             V_PGI.RunFromLanceur := (V_PGI.DefaultSectionName <> '');
             Connect := ConnecteHalley(V_PGI.CurrentAlias,FALSE,@ChargeMagHalley,NIL,NIL);
{$ENDIF}
             if not Connect then
             begin
                  if i < ParamCount then  // ajout me reprendre dans V7.0
                  begin
                        St:=ParamStr(i+1) ;
                        Nom:=UpperCase(Trim(ReadTokenPipe(St,'='))) ;
                        Value:=UpperCase(Trim(St)) ;
                        if (Nom='/INI') then
                        begin
                             SRep := ExtractFileDir (Value);
                             AssignFile(FRejet, SRep+'\ListeCom'+ V_PGI.NoDossier+'.ERR');
                             try Rewrite(FRejet) ; except end;
                             Writeln(FRejet, 'Erreur connexion à la base');
                             Flush(FRejet);
                             CloseFile(FRejet);
                        end;
                  end;
                  break;
             end;
             if v_pgi.DefaultSectionName <> '' then
             if (v_pgi.DBName = v_pgi.DefaultSectionDBName) then
                v_pgi.InBaseCommune := true;
        END ;
        if (Nom='/INI') or (Nom='/TRF')  then
        BEGIN
                        if (not Connect) then
                        begin
                             if V_PGI.PassWord = '000' then V_PGI.PassWord := CryptageSt(DayPass(Date));
{$IFNDEF EAGLCLIENT}
                             Connect := ConnecteHalley(V_PGI.CurrentAlias,TRUE,@ChargeMagHalley,NIL,NIL,NIL);
{$ELSE}
                             Connect := ConnecteHalley(V_PGI.CurrentAlias,FALSE,@ChargeMagHalley,NIL,NIL);
{$ENDIF}
                        end;
                        if pos('Minimized', St) <> 0 then Minimise  := TRUE;
                        While St <> '' do
                        begin
                              Nom := ReadTokenPipe(St, ';');              // fiche 10561
                              if (Nom = 'EXPORT') or (Nom = 'IMPORT') or (Nom = 'ENCOURS')then
                              begin
                                   Num := 10;  B1 := false; B2 := false;
                                   DispatchCom (Num, P2, B1, B2);
                                   break;
                              end;
                        end;
        END;
       end;
       Application.ProcessMessages;
       if Connect then begin
{$IFNDEF EAGLCLIENT}
          Logout ;
          DeconnecteHalley ;
{$ENDIF}
       end
       else
       begin
                  // Fermeture de l'application
                  Application.ProcessMessages;
                  if FMenuG <> nil then
                  begin
                       FMenuG.ForceClose := True ;
                       PostMessage(FmenuG.Handle,WM_CLOSE,0,0) ;
                       FMenuG.Close ;
                  end;
{$IFNDEF EAGLCLIENT}
                  If (DBSOC<>NIL) AND (DBSOC.Connected) then DBSOC.Connected:=FALSE ;
{$ENDIF}
                  VH^.STOPRSP:=TRUE ;
                  Exit;
       end;
{$IFNDEF EAGLCLIENT}
       If (DBSOC<>NIL) AND (DBSOC.Connected) then DBSOC.Connected:=FALSE ;
{$ENDIF}
       VH^.STOPRSP:=TRUE ;
       SourisNormale ;

end;



Initialization
  ProcChargeV_PGI :=  InitLaVariablePGI ;
  RegisterClassContext(TComsxContext) ;

end.
