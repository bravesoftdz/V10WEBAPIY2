
unit UImportDP;

interface

uses

  SysUtils,
  Classes,
  Controls,
  StdCtrls,
  HmsgBox,
  HCtrls,
  UTob,
  Hqry,
  HEnt1,
  {$IFNDEF EAGLCLIENT}
   {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}
  {$ELSE}
  {$ENDIF}
  ed_tools,
  Buttons ,Forms, dialogs,
  UtilTrans, ParamSoc, CBPPath;

Const version = '820';

type
  TtobArray = Array of TOB;


  TImportDP = class
  private
    listeFic_ : TStrings;//liste des fichiers générés
    listeIdent_ : TStrings;//liste des identifiants rencontrés
    listeChps_ : TStrings;//liste des champs d'une table
    listeTypes_ : TStrings;//type associé à chaque champ
    tabOffset_ : array of integer;//tableau des offset pour chaque champ
    nomFic_ : string;//chemin complet du fichier avec le nom
    pathFic_ : string;//chemin du fichier sans le nom
    fichier_ : TextFile;//fichier
    Tobs_ : TtobArray;//tableau des tobs
    entete_ : string;//ligne d'en-tête du fichier
    TobErreur  : Tob; //Gestion des erreurs lors de l'importation
    TobCIgnore : Tob; //Gestion des champs à ignorer
    MaxId : integer; //Traitement spécif des groupes de travail : MaxID de GRPDONNEES
    TobLDO, TobTempo : Tob; //Traitement spécif des groupes de travail : gestion de LIENDOSGRP

  //Méthodes privées
    function __initExtract(ident : string) : string;
    function __ligneToTob( ligne, ident, nomTable : string; idx, tabIdx : integer) : integer;
    function __ligneToTobSpecif( ligne, ident, nomTable : string; idx, tabIdx : integer) : integer;
    function __ajusterTaille(s :string) : string;
    function __SplitFile : integer;
    function __ConstruireTOB (subfic, ident : string; index : integer) : integer;
    procedure __setListes(nom, typ : string; offset : integer);
//    function __convertValue(UneValeur, typ : string; l : integer) : variant;
    function __FindTobTiers : integer;
    function ConvertirDonnee (UneDonnee : String; UnType : String; Taille : Integer; var Erreur : Boolean) : Variant;
    function IgnoreChamp (ListeChpsIgn : array of string) : integer;
    function TraiteImportGED (ident : string; idx : integer) : integer;
    function TraiteDoublonsGEDDP (strDossier, CodeGed, Libelle : string) : boolean;
    function TraiteDoublonsGEDSTD (CodeProduit, Nom, Langue, Predefini, Crit1, CodeDossier : string) : boolean;

//    function __getTobTiers (t_aux : string; idx : integer) : TOB;

    //procedure GenererFichierLog;

  public
    CheminFichierLog : String;
    FichierLogGen : Boolean;
    ChkFichierLog : Boolean;
    IsAutomate : Boolean;
    EcraserEnreg : Boolean;
    StopSurErreur : Boolean;
    //Méthodes publiques
    constructor Create;
    destructor Free;
    function LanceImport(fic : string; Auto : Boolean; Ecraser : Boolean; ArretErreur : Boolean) : integer;
    procedure GenererFichierLog(strErreur : string);

  //POUR TESTS EXPORT
    procedure __addEntete(var f : textFile);
    function generateFile(identTable : array of string; nomFic : string) : integer;
    function writeligne(rsql : tquery; var f : textfile; nomTable,identTable : string) : integer;
    function correctlength(s :string; l : integer) : string;
    function __varToStr(rsql : tquery; idx : integer) : string;
  end;

  function getVersion(fic : string) : string;
  procedure LancerExportTRA;
  procedure MyAfterImport (Sender: TObject; FileGUID: String; var Cancel: Boolean) ;

{-----------------IMPLEMENTATION-----------------------------------------------}
implementation

uses
{$IFNDEF EAGLCLIENT}
  utobtrans,
{$ENDIF}
galSystem,
utilged,
UGedFiles,
{$IFNDEF BUREAU}
uyfilestd2,
{$ENDIF}
DpTableauBordLibrairie;

//----------------------------
//--- Nom : LancerExportTRA
//----------------------------
procedure LancerExportTRA;
var Resultat     : Integer;
    ExportImport : TImportDP;
    SaveDialog   : TSaveDialog;
begin
 //--- Récupération du chemin d'export
 SaveDialog:=TSaveDialog.Create (Application);
 SaveDialog.Filename := 'Export.Tra';
 SaveDialog.Filter := 'Fichiers TRA (*.Tra)|*.Tra';
 SaveDialog.InitialDir := 'C:\';

 if (SaveDialog.Execute) then
  begin
   //--- Lancement de l'export
   if (UpperCase(ExtractFileExt(SaveDialog.FileName)) = '.TRA') then
    begin
     if (FileExists (SaveDialog.FileName)) then
      if PgiAsk ('Le fichier '+SaveDialog.Filename+' existe déjà sur le disque. Voulez-vous le remplacer ?',TitreHalley)= mrNo then
       begin
        SaveDialog.Free;
        Exit;
       end;
     ExportImport := TImportDP.Create;
     Resultat:=ExportImport.GenerateFile(['CAE','T','YTC','ANN','ANB','ANL','US','C','DOS','DAP','DOG','DCI','DCL','DCV','DEC','DFI','DPM','DOR','DPP','DSO','DSC','DTC','DT1','DTP','CC','YX','ARS','GA','PY','ET','MR','R','MDP','ANP','RPR','AFO','JUR','ROP','RD2','RPJ','RDQ','RPE','RDV','RPT','RAC','RD1','RAI','RD6'], SaveDialog.FileName);
     if (Resultat<>0) then
      PgiInfo('Erreur : Impossible d''exporter au format TRA.', TitreHalley);
     FreeAndNil(ExportImport);
    end
   else
    PgiInfo('Le nom du fichier doit comporter l''extension TRA.', TitreHalley);
  end;

 SaveDialog.Free;
end;

function getVersion(fic : string) : string;
var
  f : textFile;
  entete : string;
begin
  result := '';
  If Not FileExists(fic) then
    begin
    result := '';
    exit;
    end;
  AssignFile(f,fic);
  reset(f);
  readln(f, entete);
  result := copy(entete, 7,3);
  closeFile(f);
end;

{supervision du traitement global de l'import}
function TImportDP.LanceImport(fic : string; Auto : Boolean; Ecraser : Boolean; ArretErreur : Boolean) : integer;
var
{  SEntete,SNomChamp : String;
  SEnregistrement   : String;
  Indice,IndiceChp  : Integer; }
  i,vrt,ret_ged     : integer;
  isdel, chk        : boolean;
  nodoss, strgrp    : string;
  nomTable          : string;
begin
  CheminFichierLog:='';
  FichierLogGen:=False;
  IsAutomate := Auto;
  EcraserEnreg := Ecraser;
  StopSurErreur := ArretErreur;
  MaxId := -1;
  nodoss := '';
  strgrp := '';
  nomTable := '';

  if Not FileExists(fic) then//on quitte si le fichier n'existe pas
    begin
    GenererFichierLog('Le fichier n''existe pas !');
    result := -1;
    exit;
    end;

  result := 0;//fct optimiste !
  nomFic_ := fic;
  pathFic_ := ExtractFilePath(nomFic_);
  AssignFile(fichier_, nomFic_);
//splitter en petits fichiers
  if __SplitFile <> 0 then
    begin
    GenererFichierLog('Impossible de découper le fichier en plusieurs parties !');
    result := -1;
    exit;
    end;
  try
    Setlength(Tobs_, listeIdent_.Count);//allocation du tableau.
  except
    On EOutOfMemory Do result := -1;
  end;

//Creation tob champs à ignorer
  if IgnoreChamp(['DOS_GROUPECONF', 'DOS_APPLISPROTEC', 'ANN_CODENAFOLD', 'ANN_CODENAF2OLD']) <> 0 then
    begin
    result := -1;
    exit;
    end;

//pour chaque fichier : construire TOB
  For i:= 0 to listeFic_.Count-1 do
    begin
    result := 0;
    vrt := __construireTOB(listeFic_[i], listeIdent_[i], i);//remplit la TOB à partir du fichier
    if vrt <> 0 then result := -1;
    if (result = -1) and StopSurErreur then Break;
    end;
  // liberation tob champs à ignorer
  TobCIgnore.Free;
  Sleep(1000);
  InitMoveProgressForm(nil,'Execution Import', 'Traitement en cours...', listeFic_.Count, False, True);
//insérer toutes les TOBs
  Try
    BeginTrans;//Transaction
    chk := True;
    if result = -1 then Raise Exception.Create('Erreur d''import du TRA.');
    For i:= 0 to listeFic_.Count-1 do
      begin
      //Specif groupes de travail traité après la boucle
      if listeIdent_[i] = 'DOG' then
        begin
        deleteFile(listeFic_[i]);//Supprimer le fichier temporaire associé
        continue;
        end;
      isdel := False;
      listeChps_ := TStringList.Create;
      listeTypes_ := TStringList.Create;
      nomTable := __initExtract(listeIdent_[i]);//pour récupérer le nom de la table
      //Translation nouvelle table groupes de travail
      if nomTable = 'CHOIXCODGRP' then nomTable := 'GRPDONNEES';
      MoveCurProgressForm('Insertion dans la base des donnees de '+nomTable);
      //Traitement spécif si Import GED
      if (listeIdent_[i] = 'GEDD') or (listeIdent_[i] = 'GEDS') then
        begin
        InitializeGedFiles(V_PGI.DbName, MyAfterImport);
        ret_ged:=TraiteImportGED(listeIdent_[i], i);
        FinalizeGedFiles;
        if (TobErreur)<>nil then GenererFichierLog('');
        if ret_ged<>-1 then Raise Exception.Create('Erreur d''import des fichiers en GED (erreur '+string(ret_ged)+').');
        deleteFile(listeFic_[i]);//Supprimer le fichier temporaire associé
        continue;
        end;

      if Not ExisteSQL('SELECT 1 FROM '+nomTable) then isdel := True;
      //libération
      listeChps_.Free;
      listeTypes_.Free;
      Finalize(TabOffset_);
      chk := False;
      If isdel and (listeIdent_[i] <> 'T') then
        {$IFNDEF EAGLCLIENT}
          // #### TOBTrans(Tobs_[i]).BatchInsertDB(Q)
          // #### MB : il vous faut désormais utiliser UTobTrans
          // #### de V_800\CorrectionsV800\Commun\Lib et pas V_800\CegidPgi\Lib !!!
          TOBTrans(Tobs_[i]).BatchInsertDB
          //Tobs_[i].InsertDB(Nil,True)
        {$ELSE}
          Tobs_[i].InsertDB(Nil,True)
        {$ENDIF}
      else
        Tobs_[i].InsertOrUpdateDB(True);

{      for Indice:=0 to Tobs_[i].Detail.count-1 do
       begin
        chk:=Tobs_[i].Detail [Indice].InsertOrUpdateDB;
        if Not chk then
         begin
          //Raise Exception.Create('Erreur d''insertion de la TOB dans la base.');//contrôle sur l'insertion des TOBs
          SEntete:='';SEnregistrement:='';
          for IndiceChp:=0 to Tobs_[i].Detail [Indice].NbChamps do
           begin
            SNomChamp:=Tobs_[i].Detail [Indice].GetNomChamp(IndiceChp);
            if SEntete<>'' then SEntete:=SEntete+' | '+SNomChamp else SEntete:=SNomChamp;
            if SEnregistrement<>'' then
             SEnregistrement:=SEnregistrement+' | '+Tobs_[i].Detail [Indice].GetValue (SNomChamp).AsString
            else
             SEnregistrement:=Tobs_[i].Detail [Indice].GetValue (SNomChamp).AsString;
           end;
          GenererFichierLog (SEntete,Senregistrement);
         end;
       end;  }

      // If not chk then PgiInfo('Erreur InsertDB', TitreHalley);
      chk := True;
      deleteFile(listeFic_[i]);//Supprimer le fichier temporaire associé
      end;
    TobErreur.Free;
    FiniMoveProgressForm();
    CommitTrans;

    // Traitement spécif LIENDOSGRP après le commit pour que les infos de GRPDONNEES soient à jour
    if TobLDO <> nil then
      begin
      InitMoveProgressForm(nil,'Execution Import', 'Traitement en cours...', TobLDO.Detail.Count - 1, False, True);
      For i:=0 to TobLDO.Detail.Count - 1 do
        begin
        MoveCurProgressForm('Insertion dans la base des donnees de LIENDOSGRP');
        nodoss := TobLDO.Detail[i].GetString('NODOSS');
        strgrp := TobLDO.Detail[i].GetString('STRGRP');
        AffecteGroupe(nodoss,strgrp);
        end;
      TobLDO.Free;
      FiniMoveProgressForm();
      end;
  Except
    if not chk then GenererFichierLog ('Attention, aucune donnée intégrée !'+#13+#10
                      +'Insert ou Update de la table '+nomTable+' (identifiant:'+listeIdent_[i]+') impossible.'+#13+#10
                      +'Veuillez vérifier les données du fichier TRA correspondant à cette table.');
    RollBack;
    result := -1;
    FiniMoveProgressForm();
  End;
end;
{-----------------------------------------------------------------------------}
{Sépare le fichier TRA en plusieurs petits regroupant les lignes par leur identifiant}
function TImportDP.__SplitFile : integer;
var
  subFic  : TextFile;
  ligne, nomSubFic : string;
  tmp : string;
  position : integer;
  bResteUn : Boolean;
begin
  try
    bResteUn := False;
    reset(fichier_);
    readln(fichier_, entete_);//sauvegarde la ligne d'en-tête
    InitMoveProgressForm(nil,'Analyse du fichier d''import', 'Traitement en cours...', 20000, False, True);
    While Not EOF(fichier_) do//parcours du fichier
      begin
      readln(fichier_, ligne);//lit chaque ligne
      tmp := __ajusterTaille(copy(ligne, 4, 3));//récupère l'identifiant
      //if tmp = 'RPR' then continue;// exclusion de la table PROSPECTS en import
      if (tmp = 'CC') and (__ajusterTaille(copy(ligne, 7, 3))= 'UCO') then tmp := 'CCG';//traitement spécif groupes de travail
      if (tmp = 'GED') then //traitement spécif import des fichiers en GED
        begin
        tmp := __ajusterTaille(copy(ligne, 4, 4));
        if (tmp<>'GEDD') and (tmp<>'GEDS') then continue;
        end;
      if (tmp = 'BCN') then //traitement spécif champs Bloc-notes
        begin
        tmp := __ajusterTaille(copy(ligne, 4, 6));
        if (tmp<>'BCNC') and (tmp<>'BCNT') and (tmp<>'BCNRAC') and (tmp<>'BCNRPE') and (tmp<>'BCNROP') and (tmp<>'BCNRD2') then continue; //Ajouter ici les identifiants à prendre en compte au cas où 'BCN' devienne un vrai préfixe de table
        end;

      if listeIdent_.IndexOf(tmp) = -1 then//si absent de la liste
        begin
        if listeIdent_.Count <> 0 then//ce n'est pas la 1ère ligne du fichier
          begin
          closeFile(subFic);//close car un autre est déjà ouvert
          bResteUn := False;
          end;
        listeIdent_.Add(tmp);//ajout à la liste des identifiants
        nomSubFic := tempFileName;//génère un fichier temporaire
        listeFic_.Add(nomSubFic);//l'ajoute à la liste des fichiers
        assignFile(subFic, nomSubFic);
        bResteUn := True;
        rewrite(subFic);
        end
      else//si présent dans la liste et créé avant, on réouvre l'ancien fichier
        begin
        position := listeIdent_.IndexOf(tmp);
        if listeFic_.IndexOf(nomsubfic)<>position then
          begin
          closeFile(subFic);
          bResteUn := False;
          nomSubFic := listeFic_[position];
          assignfile(subFic,nomSubFic);
          bResteUn := True;
          append(subFic);
          end;
        end;
      MoveCurProgressForm('Analyse de la section '+tmp);
      //écrire ligne dans le fichier
      writeln(subFic, ligne);

      end;//WHILE
    FiniMoveProgressForm();
    if bResteUn then closeFile(subFic);//fermer le dernier sous-fichier
    closeFile(fichier_);
    result := 0;
  except
    result := -1;
  end;
end;
{-----------------------------------------------------------------------------}
{Construit une TOB à partir du fichier subfic}
function TImportDP.__ConstruireTOB (subfic, ident : string; index : integer) : integer;
var
  fic : TextFile;
  ligne, nomTable, nodoss, nodossprec, strgrp, strgrpprec : string;
  chk, tabIdx, i, pos : integer;
  S_Sql : string;
  Q : TQuery;
  tobF : TOB;
begin
  result := 0;
  tabIdx := -1;
  nodoss := '';
  nodossprec := '';
  strgrp := '';
  strgrpprec := '';
  listeChps_ := TStringList.Create;
  listeTypes_ := TStringList.Create;
  nomTable := __initExtract(ident);//initialise les listes
  if nomTable <> '' then //si pas d'erreur dans initExtract
    begin
    //Tobs_[index] := TOB.Create('TOB '+nomTable,nil, -1);//allocation TOB mere (virtuelle)  de la table
    if (nomTable = 'CHOIXCODGRP') then
      Tobs_[index] := TOB.Create('GRPDONNEES',nil, -1)//allocation TOB mere de la table
    else
      Tobs_[index] := TOB.Create(nomTable,nil, -1);//allocation TOB mere de la table

    //on regarde si la table TIERS a déjà été traitée. Si oui, on récupère l'index de sa TOB dans le tableau
    if nomTable = 'TIERS' then
      tabIdx := __FindTobTiers;

    //spécif groupes de travail : recherche du maxID dans GRPDONNEES si traitement de CHOIXCODGRP ou DOSSIERGRP
    if (nomTable = 'CHOIXCODGRP') then
      begin
      // Valeur Max de Grp_id
      S_Sql := 'SELECT MAX(GRP_ID) MAXID FROM GRPDONNEES WHERE GRP_NOM = "GROUPECONF"';
      Q := OpenSQL(S_Sql,FALSE);
      if not Q.EOF then MaxId := Q.FindField('MAXID').AsInteger;
      Ferme(Q);
      end;

    assignFile(fic, subfic);
    Reset(fic);
    InitMoveProgressForm(nil,'Préparation Import', 'Traitement en cours...', 20000, True, True);
    While not EOF(fic) do //lecture ligne par ligne du fichier
      begin
      readln(fic, ligne);
      //extraction des données

      TobErreur:=nil;
      result := 0;
      if (nomTable='CHOIXCODGRP') or (nomTable='DOSSIERGRP') or (nomTable='GEDDP') or (nomTable='GEDSTD') then
        chk := __ligneToTobSpecif(ligne, ident, nomTable, index, tabIdx)
      else
        chk := __ligneToTob(ligne, ident, nomTable, index, tabIdx);

      MoveCurProgressForm('Création de la tob '+ident) ;
      if chk <> 0 then result := -1;

      if (TobErreur)<>nil then GenererFichierLog('');
      TobErreur.Free;

      if (result = -1) and StopSurErreur then Break;

      end;//WHILE
    closeFile(fic);

    //Post-traitement spécif LIENDOSGRP
    if (nomTable = 'DOSSIERGRP') and (TobTempo <> nil)then
      begin
      MoveCurProgressForm('Création de la tob '+ident) ;
      if TobLDO = nil then TobLDO := TOB.Create('', nil, -1);
      TobTempo.Detail.Sort('NODOSS');
      For i:=0 to TobTempo.Detail.Count -1 do
        begin
        nodoss := TobTempo.Detail[i].GetString('NODOSS');
        strgrp := TobTempo.Detail[i].GetString('STRGRP');
        if nodoss <> nodossprec then
          begin
          tobF := TOB.Create('', TobLDO, -1);
          tobF.AddChampSupValeur('NODOSS',nodoss);
          tobF.AddChampSupValeur('STRGRP',strgrp);
          //tobF.SetString('NODOSS',nodoss);
          //tobF.SetString('STRGRP',strgrp);
          end
        else
          begin
          pos := TobLDO.Detail.Count - 1;
          strgrpprec := TobLDO.Detail[pos].GetString('STRGRP');
          TobLDO.Detail[pos].SetString('STRGRP',strgrpprec+';'+strgrp);
          //tobF.SetString('STRGRP',tobF.GetString('STRGRP')+';'+strgrp);
          end;
        nodossprec := nodoss;
        end;//FOR
      end;//Post-traitement
    FiniMoveProgressForm();
    end//IF
  else result := -1;//erreur dans l'initialisation des listes

  //libération
  listeChps_.Free;
  listeTypes_.Free;
  Finalize(TabOffset_);
end;
{-----------------------------------------------------------------------------}
{Prend en paramètre une ligne du fichier et met les données dans la tob}
function TImportDP.__ligneToTob( ligne, ident, nomTable : string; idx, tabIdx : integer) : integer;
var
  tobF : TOB;
  i,c : integer;
  val : string;
  Enregistrement : String;
  UneDonnee      : Variant;
  Erreur,EnregOk : Boolean;
  UneTobErreur   : Tob;
begin
 EnregOk := True;
 Enregistrement := Ligne;
 Erreur  := False;
 result  := 0;
 tobF := TOB.Create(nomTable, Tobs_[idx], -1);
 tobF.InitValeurs;
 c := listeChps_.Count-1;
 for i := 0 to c do
  begin
   if TobCIgnore.FieldExists(listeChps_[i]) then continue; // si champ à exclure, on l'ignore et continue
   if (Not tobF.FieldExists(listeChps_[i])) then continue;//si champ n'existe pas, on l'ignore
   //extrait la donnée de la ligne et enlève les espaces à gauche ou à droite
   val := __ajusterTaille(copy(ligne, TabOffset_[i], TabOffset_[i+1] - TabOffset_[i]));
   if (val <> '') And (val<> '@') then//on n'insère pas les chamsp vides
    begin
     try
      //Si table TIERS, on regarde si l'autre partie a déjà été traitée afin de réutiliser les TOBs
      //et ne pas en créer d'autres.
{      if (nomTable = 'TIERS') And (i = 0) And (tabIdx > -1) then
          begin//recherche de la tob correspondante à la ligne en cours d'extraction
            TobT := Tobs_[tabIdx].FindFirst(['T_AUXILIAIRE'],[val],true);
            if TobT <> Nil then//si est trouvée
              begin
              TobF.Free;//suppression de la 1ere tob
              TobF := TobT;//on récupère l'ancienne
              end;
          end;
}        //conversion de la donnée dans le type voulu
        //donnee := __convertValue(val, ListeTypes_[i], TabOffset_[i+1] - TabOffset_[i]);

        UneDonnee:=ConvertirDonnee (Val, ListeTypes_[i], TabOffset_[i+1] - TabOffset_[i], Erreur);
        if (ChkFichierLog) and (Erreur) then
          begin
          if TobErreur=nil then TobErreur:=Tob.Create ('',nil,-1);
          UneTobErreur:=TOB.Create('', TobErreur, -1);
          UneTobErreur.AddChampSupValeur ('Enregistrement',Enregistrement);
          UneTobErreur.AddChampSupValeur ('Table',NomTable);
          UneTobErreur.AddChampSupValeur ('Champ',ListeChps_[i]);
          UneTobErreur.AddChampSupValeur ('Valeur',Val);
          EnregOk:=False;
          Erreur:=False;
          end;

        tobF.PutValue(listeChps_[i],Unedonnee);

      except
        on E: Exception do
          begin
          GenererFichierLog('ERREUR : '+E.Message);
          if not IsAutomate then pgiinfo('ERREUR  : '+E.Message);
          result := -1;
          end;
      end;//except
    end;//IF
  end;//FOR

  //--- On rejete l'enregistrement
  if not (EnregOk) then
    begin
    TobF.Free;
    result := -1;
    end
  else
    if EcraserEnreg AND ( (copy(ident,1,3)<>'BCN') or ((copy(ident,1,3)='BCN') and (Length(ident)=3)) ) then TobF.SetAllModifie(True);

end;
{-----------------------------------------------------------------------------}
{Idem ci-dessus mais traitement spécifique des groupes de travail}
function TImportDP.__ligneToTobSpecif( ligne, ident, nomTable : string; idx, tabIdx : integer) : integer;
var tobF : TOB;
    i,j,c : integer;
    val : string;
    Enregistrement : String;
    UneDonnee      : Variant;
    Erreur,EnregOk : Boolean;
    UneTobErreur   : Tob;
    Q : TQuery;
    ok : Boolean;

    procedure TraiteErreur;
    begin
      if TobErreur=nil then TobErreur:=Tob.Create ('',nil,-1);
      UneTobErreur:=TOB.Create('', TobErreur, -1);
      UneTobErreur.AddChampSupValeur ('Enregistrement',Enregistrement);
      UneTobErreur.AddChampSupValeur ('Table',NomTable);
      UneTobErreur.AddChampSupValeur ('Champ',ListeChps_[i]);
      UneTobErreur.AddChampSupValeur ('Valeur',Val);
      EnregOk:=False;
      Erreur:=False;
    end;
    
    function AffecteChemin(Path : string ; Fic : string) : string ;
    begin
      Fic:=Path + ExtractFileName(Fic);//On extrait le nom de fichier au cas où
      if Not FileExists(Fic) then
        begin
        Val:=Val+' (fichier inexistant)';
        TraiteErreur;
        Result:='';
        end
      else
        Result:=Fic;
    end;

begin
 EnregOk := True;
 Enregistrement := Ligne;
 Erreur  := False;
 result  := 0;

 //Création des tobs spécifiques : DOSSIERGRP alimente LIENDOSGRP et CHOIXCODGRP alimente GRPDONNEES
 if nomTable = 'DOSSIERGRP' then
    begin
    if TobTempo = nil then TobTempo := TOB.Create('', nil, -1);
    tobF := TOB.Create('', TobTempo, -1);
    tobF.AddChampSup('NODOSS',False);
    tobF.AddChampSup('STRGRP',False);
    end
 else if nomTable = 'CHOIXCODGRP' then
    begin
    tobF := TOB.Create('GRPDONNEES', Tobs_[idx], -1);
    tobF.SetString('GRP_NOM','GROUPECONF');
    end
 else
    tobF := TOB.Create(nomTable, Tobs_[idx], -1);

 c := listeChps_.Count-1;
 for i := 0 to c do
  begin
   //extrait la donnée de la ligne et enlève les espaces à gauche ou à droite
   val := __ajusterTaille(copy(ligne, TabOffset_[i], TabOffset_[i+1] - TabOffset_[i]));
   if (val <> '') And (val<> '@') then//on n'insère pas les champs vides
    begin
     try
        UneDonnee:=ConvertirDonnee (Val, ListeTypes_[i], TabOffset_[i+1] - TabOffset_[i], Erreur);

        if (ChkFichierLog) and (Erreur) then TraiteErreur;
{          if TobErreur=nil then TobErreur:=Tob.Create ('',nil,-1);
          UneTobErreur:=TOB.Create('', TobErreur, -1);
          UneTobErreur.AddChampSupValeur ('Enregistrement',Enregistrement);
          UneTobErreur.AddChampSupValeur ('Table',NomTable);
          UneTobErreur.AddChampSupValeur ('Champ',ListeChps_[i]);
          UneTobErreur.AddChampSupValeur ('Valeur',Val);
          EnregOk:=False;
          Erreur:=False;  }

        // Traitement spécif DOSSIERGRP
        if nomTable = 'DOSSIERGRP' then
          begin
          if listeChps_[i] = 'DOG_NODOSSIER' then tobF.SetString('NODOSS',UneDonnee);
          if listeChps_[i] = 'DOG_GROUPECONF' then tobF.SetString('STRGRP',UneDonnee);

{          // Recherche dans les tobs filles d'import de DOSSIERGRP
          if listeChps_[i] = 'DOG_NODOSSIER' then
            begin
            Tmp := TobLDO.FindFirst(['NODOSS'],[UneDonnee],false) ;
            if Tmp = nil then
              begin
              tobF := TOB.Create('', TobLDO, -1);
              tobF.AddChampSup('NODOSS',False);
              tobF.AddChampSup('STRGRP',False);
              tobF.SetString('NODOSS',UneDonnee);
              end;

            //For j:=0 to Tobs_[idx].Detail.Count - 1 do
              //begin
              //if Tobs_[idx].Detail[j].GetString('NODOSS') = UneDonnee then
              //end;
            end;

          if listeChps_[i] = 'DOG_GROUPECONF' then
            begin
            if Tmp <> nil then
              begin
              For j:=0 to TobLDO.Detail.Count - 1 do
                begin
                nodoss := TobLDO.Detail[j].GetString('NODOSS');
                strgrp := TobLDO.Detail[j].GetString('STRGRP');
                if Tmp.GetString('NODOSS')=nodoss then TobLDO.Detail[j].SetString('STRGRP',strgrp+';'+UneDonnee);
                end;
              //TobLDO.FindFirst(['NODOSS'],[Tmp.GetString('NODOSS')],false);
              //j := Tmp.GetIndex;
              //if j > -1 then TobLDO.Detail[j].SetString('STRGRP',';'+UneDonnee);
              end
            else
              tobF.SetString('STRGRP',UneDonnee);
            end;
}
          end;

        // Traitement spécif CHOIXCODGRP
        if nomTable = 'CHOIXCODGRP' then
          begin
          if listeChps_[i] = 'CC_CODE' then
            begin
            ok := false;
            // Recherche si Code existe déjà dans la table, on récupère l'ID
            Q := OpenSQL('SELECT * FROM GRPDONNEES',False);
            while not Q.Eof do
              begin
              if Q.FindField('GRP_CODE').AsString = UneDonnee then
                begin
                //id := Q.FindField('GRP_ID').AsInteger;
                //tobF.SetInteger('GRP_ID',id);
                tobF.SetInteger('GRP_ID',Q.FindField('GRP_ID').AsInteger);
                ok := true;
                break;
                end;
              Q.Next;
              end;
            Ferme(Q);

            // Si pas trouvé précédemment, recherche dans les tobs filles d'import de CHOIXCODGRP, pour éviter les doublons
            if not ok then
              begin
              For j := 0 to Tobs_[idx].Detail.Count - 1 do
                begin
                //code := Tobs_[idx].Detail[j].GetString('GRP_CODE');
                //if code = UneDonnee then
                if Tobs_[idx].Detail[j].GetString('GRP_CODE') = UneDonnee then
                  begin
                  //id := Tobs_[idx].Detail[j].GetInteger('GRP_ID');
                  //tobF.SetInteger('GRP_ID',id);
                  tobF.SetInteger('GRP_ID',Tobs_[idx].Detail[j].GetInteger('GRP_ID'));
                  tobF.SetInteger('GRP_IDPERE',-1);
                  ok := true;
                  break;
                  end;
                end;
              end;

            // Si pas trouvé précédemment, c'est un nouveau groupe, donc affectation d'un nouvel ID
            if not ok then
              begin
              MaxId := MaxId+1;
              tobF.SetInteger('GRP_ID',MaxId);
              tobF.SetInteger('GRP_IDPERE',-1);
              ok := true;
              end;

            if ok then tobF.SetString('GRP_CODE',UneDonnee);
            end

          else if listeChps_[i] = 'CC_LIBELLE' then
            tobF.SetString('GRP_LIBELLE',UneDonnee)
          else
            continue;

          end;//Fin spécif CHOIXCODGRP

        // Traitement spécif GED
        if (nomTable = 'GEDDP') or (nomTable = 'GEDSTD') then
          begin
          if listeChps_[i] = 'FICHIER' then UneDonnee:=AffecteChemin(pathFic_, UneDonnee);
          tobF.AddChampSupValeur(listeChps_[i],UneDonnee);
          end;

      except
        on E: Exception do
          begin
          GenererFichierLog('ERREUR : '+E.Message);
          if not IsAutomate then pgiinfo('ERREUR  : '+E.Message);
          result := -1;
          end;
      end;//except
    end;//IF
  end;//FOR

  //--- On rejete l'enregistrement
  if not (EnregOk) then
    begin
    TobF.Free;
    result := -1;
    end;

end;
{-----------------------------------------------------------------------------}
{Permet d'exclure des champs en import, tout en les gardant dans le format général}
function TImportDP.IgnoreChamp(ListeChpsIgn : array of string) : integer;
var
  i : integer;
begin
  try
    if TobCIgnore = nil then TobCIgnore := TOB.Create('',nil,-1);
    For i := 0 to high(ListeChpsIgn) do
    begin
      TobCIgnore.AddChampSup(ListeChpsIgn[i], False); //initalise la liste
    end;
    result := 0;
  except
    result := -1;
  end;
end;

//-------------------------------------------------------
//--- Nom : ConvertirDonnee
//-------------------------------------------------------
function TImportDP.ConvertirDonnee (UneDonnee : String; UnType : String; Taille : Integer; var Erreur : Boolean) : Variant;
begin
 Erreur:=False;

 case ord(UpperCase (UnType)[1]) of
  //--- Convertion CHAINE
  65 : Result:=copy(uneDonnee,1, Taille);
  //--- Convertion BOOLEAN
  66 : Result:=copy(uneDonnee,1, Taille);
  //--- Convertion DATE
  68 : try
        if UneDonnee = '' then
         result := idate1900
        else
         result:=EncodeDate (strToInt(copy(UneDonnee,5,4)),strToInt(copy(UneDonnee, 3,2)),strToInt(copy(UneDonnee,1,2)));

       except
        On EConvertError Do
         begin
          Erreur:=True;
          result := idate1900;//échec => date au 01/01/1900 par défaut;
         end;
       end;
  //--- Conversion ENTIER / DECIMALE quel que soit le separateur de decimales
  78  : if (pos (',',UneDonnee)=0) and (pos ('.',UneDonnee)=0) then
         begin
          try
           result := StrToInt(UneDonnee);
          except
           On EConvertError Do
            begin
             Erreur:=True;
             result := 0;
            end;
          end;
         end
        else
         begin
          if (Pos(',',UneDonnee)<>0) or (Pos('.',UneDonnee)<>0) then
           begin
            try
              //Result:=Valeur(UneDonnee);
              Result:=Valeur(StringReplace (UneDonnee,'.',',',[RFREPLACEALL]));
            except
             On EConvertError Do
              begin
               Erreur:=True;
               result := 0;
              end;
            end;
           end;
         end;
  //--- Conversion TEXT limité à 1500 caractères
  84  : begin
        UneDonnee := FindEtReplace(UneDonnee, Chr(3), Chr(13)+Chr(10), TRUE);
        UneDonnee := FindEtReplace(UneDonnee, Chr(4), Chr(0), TRUE);
        Result:=copy(uneDonnee,1, Taille);
        end;
 end;
end;

{------------------------------------------------------------------------------}
{Traitement spécifique de l'import de fichiers en GED}
function TImportDP.TraiteImportGED (ident : string; idx : integer) : integer;
var
  ListParam : TStrings;
  SChamp, strDossier : string;
  indice, indicechp : integer;
  Res : boolean;
  UneTobErreur   : Tob;
begin
  Res:=True;
  result:=-1;
  ListParam:=TStringList.Create;
  for indice:=0 to Tobs_[idx].Detail.count-1 do
  begin
    ListParam.Clear;
    SChamp := '';
    for indicechp:=0 to listeChps_.Count-1 do
    begin
      SChamp:= Tobs_[idx].Detail[indice].GetValue(listeChps_[indicechp]);
      if SChamp = #0 then SChamp := '';
      if (copy(listeChps_[indicechp],1,5) = 'BCRIT') and (SChamp = '') then SChamp:='-';
      ListParam.Add(SChamp);
    end;

    //Selon type d'import, appels des fonctions (voir ordre des champs du fichier dans __initExtract)
    if ident = 'GEDD' then
      begin
      if ListParam[4] = '' then
        strDossier := Trim (DonnerDossierFromGuid (DonnerGuidFromTiers (ListParam[5])))
      else
        strDossier := Trim (ListParam[4]);
      // si code dossier inconnu par rapport au tiers fourni, on n'autorise pas l'insertion du document + génère erreur dans le log
      if strDossier = '' then
        begin
        if TobErreur=nil then TobErreur:=Tob.Create ('',nil,-1);
        UneTobErreur:=TOB.Create('', TobErreur, -1);
        UneTobErreur.AddChampSupValeur ('Enregistrement',ListParam[0]);
        UneTobErreur.AddChampSupValeur ('Table',ident);
        UneTobErreur.AddChampSupValeur ('Champ',ListParam[5]);
        UneTobErreur.AddChampSupValeur ('Valeur','code dossier inconnu par rapport au tiers fourni');
        exit;
        end;
      //Paramètres : strDossier,CodeGed,Libelle
      Res:=TraiteDoublonsGEDDP(strDossier,ListParam[3],ListParam[6]);
      //Paramètres : Fichier, CodeDossier, CodeTiers, CodeGed, Description, Auteur, Annee, Mois
      if Res then Res:=TraiteAjoutFichierDansGed(ListParam[0],strDossier,ListParam[3],ListParam[6],ListParam[7],ListParam[8],ListParam[9],True);
      if Not Res then result:=0;
      end
    else if ident = 'GEDS' then
      begin
      //Paramètres : CodeProduit, Nom, Langue, Predefini, Crit1, CodeDossier
      Res:=TraiteDoublonsGEDSTD(ListParam[1],ListParam[2],ListParam[7],ListParam[6],ListParam[8],ListParam[4]);
      //Paramètres : Fichier, CodeProduit, Nom, Extension,
      //             Crit1, Crit2, Crit3, Crit4, Crit5,
      //             BCrit1, BCrit2, BCrit3, BCrit4, BCrit5,
      //             Langue, Predefini, Libelle, Dossier
      {$IFNDEF BUREAU}
      if Res then result:=AGL_YFILESTD_IMPORT(ListParam[0],ListParam[1],ListParam[2],ListParam[3],
                               ListParam[8],ListParam[9],ListParam[10],ListParam[11],ListParam[12],
                               ListParam[13],ListParam[14],ListParam[15],ListParam[16],ListParam[17],
                               ListParam[7],ListParam[6],ListParam[5],ListParam[4]);
      {$ENDIF}
      end

    else //normalement jamais là
      result:=0;
  end;
  ListParam.Free;
end;


function TImportDP.TraiteDoublonsGEDDP (strDossier, CodeGed, Libelle : string) : boolean;
var
  Q : TQuery;
  SWhere, Cle : string;
begin
  Cle:='';
  SWhere := 'WHERE DPD_NODOSSIER = "' + strDossier + '"';
  SWhere := SWhere + ' AND DPD_CODEGED = "' + CodeGed + '"';
  SWhere := SWhere + ' AND DPD_DOCGUID = YDO_DOCGUID';
  SWhere := SWhere + ' AND YDO_LIBELLEDOC = "' + Libelle + '"';
  Q := OpenSql('SELECT DPD_DOCGUID FROM DPDOCUMENT, YDOCUMENTS '+SWhere,True);
  if Q.Eof then
    Result := True
  else
  begin
    while Not Q.Eof do
    begin
      Cle:=Q.FindField ('DPD_DOCGUID').AsString;
      if Cle<>'' then
      begin
        ExecuteSql ('DELETE DPDOCUMENT WHERE DPD_DOCGUID="'+Cle+'"');
        ExecuteSql ('DELETE YDOCUMENTS WHERE YDO_DOCGUID="'+Cle+'"');
        ExecuteSql ('DELETE YFILES FROM YFILES, YDOCFILES WHERE YFI_FILEGUID = YDF_FILEGUID AND YDF_DOCGUID="'+Cle+'" AND ');
        ExecuteSql ('DELETE YFILEPARTS FROM YFILEPARTS, YDOCFILES WHERE YFP_FILEGUID = YDF_FILEGUID AND YDF_DOCGUID="'+Cle+'"');
        ExecuteSql ('DELETE YDOCFILES WHERE YDF_DOCGUID="'+Cle+'"');
      end;
      Q.Next;
    end;
    Result := True;
  end;
  Ferme(Q);
end;

function TImportDP.TraiteDoublonsGEDSTD (CodeProduit, Nom, Langue, Predefini, Crit1, CodeDossier : string) : boolean;
var
  Q : TQuery;
  SWhere, Cle : string;
begin
  Cle:='';
  //Si predefini = standard, on ne tient pas compte du n° de dossier, on force à 000000
  if UpperCase(Predefini)='STD' then CodeDossier:='000000';
  SWhere := 'WHERE YFS_CODEPRODUIT = "' + CodeProduit + '"';
  SWhere := SWhere + ' AND YFS_NOM = "' + Nom + '"';
  SWhere := SWhere + ' AND YFS_LANGUE = "' + Langue + '"';
  SWhere := SWhere + ' AND YFS_PREDEFINI = "' + Predefini + '"';
  SWhere := SWhere + ' AND YFS_CRIT1 = "' + Crit1 + '"';
  SWhere := SWhere + ' AND YFS_NODOSSIER = "' + CodeDossier + '"';
  Q := OpenSql('SELECT YFS_FILEGUID FROM YFILESTD '+SWhere,True);
  if Q.Eof then
    Result := True
  else
  begin
    while Not Q.Eof do
    begin
      Cle:=Q.FindField ('YFS_FILEGUID').AsString;
      if Cle<>'' then
      begin
        ExecuteSql ('DELETE NFILEPARTS WHERE NFS_FILEGUID="'+Cle+'"');
        ExecuteSql ('DELETE NFILES WHERE NFI_FILEGUID="'+Cle+'"');
        ExecuteSql ('DELETE YFILESTD WHERE YFS_FILEGUID="'+Cle+'"');
      end;
      Q.Next;
    end;
    Result := True;
  end;
  Ferme(Q);
end;

{------------------------------------------------------------------------------}
{Renvoie la valeur convertie dans le type voulu}
{function TImportDP.__convertValue(Unevaleur, typ : string; l : integer) : variant;
var
  d : TDateTime;
  n : Extended;
  i,dd,mm,yyyy : integer;
  t : string;
begin
  t := LowerCase(typ[1]);
  if (t = 'a') Or (t = 'b') then //chaîne ou booléen
    result := copy(unevaleur,1, l)
  else if t = 'd' then //date
    begin
    try
      //extraction des valeurs des jour, mois, année
      if UneValeur = '' then
        result := idate1900
      else
        begin
        dd := strToInt(copy(UneValeur,1,2));
        mm := strToInt(copy(UneValeur, 3,2));
        yyyy := strToInt(copy(UneValeur,5,4));
        d := EncodeDate(yyyy, mm, dd);//création de la date
        result := d;
        end;
    except
       On EConvertError Do
        begin
         UneTobErreur:=TOB.Create('', TobErreur, -1);
         UneTobErreur.AddChampSupValeur ('Table','')
         UneTobErreur.AddChampSupValeur ('Champ','')
         Erreur:=True;
         result := idate1900;//échec => date au 01/01/1900 par défaut;
        end;
    end;//except
    end
  else if (t = 'n') And (Pos('.',UneValeur)=0) then//entier
    begin
    try
      i := StrToInt(UneValeur);
      result := i;
    except
      On EConvertError Do
       begin
        Erreur:=True;
        result := 0;//échec
       end;
    end;
    end
  else if (t = 'n') And (Pos('.',UneValeur)<>0) then//numeric
    begin
    try
      n:= Valeur(StringReplace (UneValeur,'.',',',[RFREPLACEALL]));
      Result := n;
    except
      On EConvertError Do
       begin
        Erreur:=True;
        result := 0;
       end;
    end;//try
    end;
end;}
{-----------------------------------------------------------------------------}
{parcourt les TOB du tableau Tobs_ à la recherche de la TOB TIERS. Si elle existe,
 renvoie son index.}
function TImportDP.__FindTobTiers : integer;
var
  i,c : integer;
begin
  result := -1;
  c := High(Tobs_);
  For i := 0 to c do
    begin
    //if Tobs_[i].NomTable = 'TOB TIERS' then
    if Tobs_[i].NomTable = 'TIERS' then
      begin
      result := i;
      exit;
      end
    end;
end;
{-----------------------------------------------------------------------------}
{Renvoie le nom de la table associée à l'identifiant passé en paramètre
 & remplit les listes des champs}
function TImportDP.__initExtract(ident : string) : string;
begin
  try
  listeChps_.Clear;
  listeTypes_.Clear;
  if ident = 'CAE' then //données de base du TIERS
    begin
    result := 'TIERS';
    setLength(tabOffset_, 52);
    __setListes('T_AUXILIAIRE','a17',7);
    __setListes('T_LIBELLE','a35',24);
    __setListes('T_NATUREAUXI','a3',59);
    __setListes('T_LETTRABLE','b1',62);
    __setListes('T_COLLECTIF','a17',63);
    __setListes('T_EAN','a17',80);
    __setListes('T_TABLE0','a17',97);
    __setListes('T_TABLE1','a17',114);
    __setListes('T_TABLE2','a17',131);
    __setListes('T_TABLE3','a17',148);
    __setListes('T_TABLE4','a17',165);
    __setListes('T_TABLE5','a17',182);
    __setListes('T_TABLE6','a17',199);
    __setListes('T_TABLE7','a17',216);
    __setListes('T_TABLE8','a17',233);
    __setListes('T_TABLE9','a17',250);
    __setListes('T_ADRESSE1','a35',267);
    __setListes('T_ADRESSE2','a35',302);
    __setListes('T_ADRESSE3','a35',337);
    __setListes('T_CODEPOSTAL','a9',372);
    __setListes('T_VILLE','a35',381);
    __setListes('DOMICILIATION','a24',416);
    __setListes('ETABLISSEMENT','a5',440);
    __setListes('GUICHET','a5',445);
    __setListes('COMPTE','a11',450);
    __setListes('CLE','a2',461);
    __setListes('T_PAYS','a3',463);
    __setListes('T_ABREGE','a17',466);
    __setListes('T_LANGUE','a3',483);
    __setListes('T_MULTIDEVISE','b1',486);
    __setListes('T_DEVISE','a3',487);
    __setListes('T_TELEPHONE','a25',490);
    __setListes('T_FAX','a25',515);
    __setListes('T_REGIMETVA','a3',540);
    __setListes('T_MODEREGLE','a3',543);
    __setListes('T_COMMENTAIRE','a35',546);
    __setListes('T_NIF','a17',581);
    __setListes('T_SIRET','a17',598);
    __setListes('T_APE','a5',615);
    __setListes('CONTACT :  NOM ','a35',620);
    __setListes('CONTACT :  SERVICE ','a35',655);
    __setListes('CONTACT :  FONCTION ','a35',690);
    __setListes('CONTACT :  TELEPHONE ','a25',725);
    __setListes('CONTACT :  FAX ','a25',750);
    __setListes('CONTACT :  TELEX ','a25',775);
    __setListes('CONTACT :  RVA ','a50',800);
    __setListes('CONTACT : CIVILITE','a3',850);
    __setListes('CONTACT : PRINCIPAL','b1',853);
    __setListes('T_FORMEJURIDIQUE','a3',854);
    __setListes('RIB PRINCIPAL','b1',857);
    __setListes('T_TVAENCAISSEMENT','a3',858);
    tabOffset_[51] := 861;
    end
  else if ident = 'T' then //données complémentaires sur le TIERS
    begin
    result := 'TIERS';
    SetLength(tabOffset_, 101);//dimensionne le tableau
    __setListes('T_AUXILIAIRE','a17',7);
    __setListes('T_NATUREAUXI','a3',24);
    __setListes('T_TIERS','a17',27);
    __setListes('T_PARTICULIER','b1',44);
    __setListes('T_TELEX','a25',45);
    __setListes('T_TELEPHONE2','a25',70);
    __setListes('T_FACTURE','a17',95);
    __setListes('T_PAYEUR','a17',112);
    __setListes('T_APPORTEUR','a17',129);
    __setListes('T_SOCIETEGROUPE','a17',146);
    __setListes('T_PRESCRIPTEUR','a17',163);
    __setListes('T_TRANSPORTEUR','a17',180);
    __setListes('T_COEFCOMMA','n15.4',197);
    __setListes('T_SECTEUR','a3',212);
    __setListes('T_ZONECOM','a3',215);
    __setListes('T_REPRESENTANT','a17',218);
    __setListes('T_TARIFTIERS','a3',235);
    __setListes('T_REMISE','n15.2',238);
    __setListes('T_FRANCO','n15.2',253);
    __setListes('T_ESCOMPTE','n15.2',268);
    __setListes('T_QUALIFESCOMPTE','a3',283);
    __setListes('T_JOURPAIEMENT1','n6',286);
    __setListes('T_JOURPAIEMENT2','n6',292);
    __setListes('T_FACTUREHT','b1',298);
    __setListes('T_SOUMISTPF','b1',299);
    __setListes('T_CORRESP1','a17',300);
    __setListes('T_CORRESP2','a17',317);
    __setListes('T_DATECREATION','d8',334);
    __setListes('T_DATEMODIF','d8',342);
    __setListes('T_DATEOUVERTURE','d8',350);
    __setListes('T_DATEFERMETURE','d8',358);
    __setListes('T_FERME','b1',366);
    __setListes('T_DATEDERNMVT','d8',367);
    __setListes('T_DEBITDERNMVT','n15.2',375);
    __setListes('T_CREDITDERNMVT','n15.2',390);
    __setListes('T_NUMDERNMVT','n6',405);
    __setListes('T_LIGNEDERNMVT','n6',411);
    __setListes('T_CONFIDENTIEL','b1',417);
    __setListes('T_SOLDEPROGRESSIF','b1',418);
    __setListes('T_SAUTPAGE','b1',419);
    __setListes('T_TOTAUXMENSUELS','b1',420);
    __setListes('T_COUTHORAIRE','n15.2',421);
    __setListes('T_SOCIETE','a3',436);
    __setListes('T_RELANCEREGLEMENT','a3',439);
    __setListes('T_RELANCETRAITE','a3',442);
    __setListes('T_RESIDENTETRANGER','a3',445);
    __setListes('T_NATUREECONOMIQUE','a3',448);
    __setListes('T_MOTIFVIREMENT','a3',451);
    __setListes('T_LETTREPAIEMENT','a3',454);
    __setListes('T_RELEVEFACTURE','b1',457);
    __setListes('T_FREQRELEVE','a3',458);
    __setListes('T_JOURRELEVE','n6',461);
    __setListes('T_UTILISATEUR','a3',467);
    __setListes('T_RVA','a250',470);
    __setListes('T_EMAIL','a250',720);
    __setListes('T_DERNLETTRAGE','a4',970);
    __setListes('T_JURIDIQUE','a3',974);
    __setListes('T_PROFIL','a3',977);
    __setListes('T_DATEDERNRELEVE','d8',980);
    __setListes('T_SCORERELANCE','n6',988);
    __setListes('T_SCORECLIENT','n6',994);
    __setListes('T_PAYEURECLATEMENT','b1',1000);
    __setListes('T_DATEDERNPIECE','d8',1001);
    __setListes('T_NUMDERNPIECE','n6',1009);
    __setListes('T_TOTDERNPIECE','n15.2',1015);
    __setListes('T_CONSO','a3',1030);
    __setListes('T_CREDITDEMANDE','n15.2',1033);
    __setListes('T_CREDITACCORDE','n15.2',1048);
    __setListes('T_DOSSIERCREDIT','a35',1063);
    __setListes('T_DATECREDITDEB','d8',1098);
    __setListes('T_DATECREDITFIN','d8',1106);
    __setListes('T_CREDITPLAFOND','n15.2',1114);
    __setListes('T_DATEPLAFONDDEB','d8',1129);
    __setListes('T_DATEPLAFONDFIN','d8',1137);
    __setListes('T_NIVEAURISQUE','a3',1145);
    __setListes('T_AVOIRRBT','b1',1148);
    __setListes('T_ISPAYEUR','b1',1149);
    __setListes('T_CODEIMPORT','a17',1150);
    __setListes('T_NATIONALITE','a3',1167);
    __setListes('T_PRENOM','a35',1170);
    __setListes('T_JOURNAISSANCE','n6',1205);
    __setListes('T_MOISNAISSANCE','n6',1211);
    __setListes('T_ANNEENAISSANCE','n6',1217);
    __setListes('T_SEXE','a3',1223);
    __setListes('T_DEBRAYEPAYEUR','b1',1226);
    __setListes('T_MOISCLOTURE','n6',1227);
    __setListes('T_EURODEFAUT','b1',1233);
    __setListes('T_PASSWINTERNET','a20',1234);
    __setListes('T_PUBLIPOSTAGE','b1',1254);
    __setListes('T_ORIGINETIERS','a3',1255);
    __setListes('T_DATEPROCLI','d8',1258);
    __setListes('T_ETATRISQUE','a3',1266);
    __setListes('T_DOMAINE','a3',1269);
    __setListes('T_DATEINTEGR','d8',1272);
    __setListes('T_NIVEAUIMPORTANCE','a3',1280);
    __setListes('T_ENSEIGNE','a35',1283);
    __setListes('T_REGION','a9',1318);
    __setListes('T_COMPTATIERS','a3',1327);
    __setListes('T_DELAIMOYEN','n6',1330);
    __setListes('T_EMAILING','b1',1336);
    tabOffset_[100] := 1337;
    end
  else if ident = 'YTC' then
    begin
    result := 'TIERSCOMPL';
    setLength(tabOffset_, 86);
    __setListes('YTC_AUXILIAIRE','a17',7);
    __setListes('YTC_TIERS','a17',24);
    __setListes('YTC_TABLELIBRETIERS1','a6',41);
    __setListes('YTC_TABLELIBRETIERS2','a6',47);
    __setListes('YTC_TABLELIBRETIERS3','a6',53);
    __setListes('YTC_TABLELIBRETIERS4','a6',59);
    __setListes('YTC_TABLELIBRETIERS5','a6',65);
    __setListes('YTC_TABLELIBRETIERS6','a6',71);
    __setListes('YTC_TABLELIBRETIERS7','a6',77);
    __setListes('YTC_TABLELIBRETIERS8','a6',83);
    __setListes('YTC_TABLELIBRETIERS9','a6',89);
    __setListes('YTC_TABLELIBRETIERSA','a6',95);
    __setListes('YTC_VALLIBRE1','n20.2',101);
    __setListes('YTC_VALLIBRE2','n20.2',121);
    __setListes('YTC_VALLIBRE3','n20.2',141);
    __setListes('YTC_DATELIBRE1','d8',161);
    __setListes('YTC_DATELIBRE2','d8',169);
    __setListes('YTC_DATELIBRE3','d8',177);
    __setListes('YTC_TABLELIBREFOU1','a6',185);
    __setListes('YTC_TABLELIBREFOU2','a6',191);
    __setListes('YTC_TABLELIBREFOU3','a6',197);
    __setListes('YTC_VALLIBREFOU1','n20.2',203);
    __setListes('YTC_VALLIBREFOU2','n20.2',223);
    __setListes('YTC_VALLIBREFOU3','n20.2',243);
    __setListes('YTC_DATELIBREFOU1','d8',263);
    __setListes('YTC_DATELIBREFOU2','d8',271);
    __setListes('YTC_DATELIBREFOU3','d8',279);
    __setListes('YTC_TEXTELIBRE1','a35',287);
    __setListes('YTC_TEXTELIBRE2','a35',322);
    __setListes('YTC_TEXTELIBRE3','a35',357);
    __setListes('YTC_BOOLLIBRE1','b1',392);
    __setListes('YTC_BOOLLIBRE2','b1',393);
    __setListes('YTC_BOOLLIBRE3','b1',394);
    __setListes('YTC_RESSOURCE1','a17',395);
    __setListes('YTC_RESSOURCE2','a17',412);
    __setListes('YTC_RESSOURCE3','a17',429);
    __setListes('YTC_DOCIDENTITE','a3',446);
    __setListes('YTC_DOCDATEDELIV','d8',449);
    __setListes('YTC_DOCDATEEXPIR','d8',457);
    __setListes('YTC_DOCOBSERV','a35',465);
    __setListes('YTC_DOCORIGINE','a35',500);
    __setListes('YTC_FAMREG','a3',535);
    __setListes('YTC_DATECREATION','d8',538);
    __setListes('YTC_DATEMODIF','d8',546);
    __setListes('YTC_CREATEUR','a3',554);
    __setListes('YTC_UTILISATEUR','a3',557);
    __setListes('YTC_COMMSPECIAL','a17',560);
    __setListes('YTC_REPRESENTANT2','a17',577);
    __setListes('YTC_REPRESENTANT3','a17',594);
    __setListes('YTC_TAUXREPR1','n20.2',611);
    __setListes('YTC_TAUXREPR2','n20.2',631);
    __setListes('YTC_TAUXREPR3','n20.2',651);
    __setListes('YTC_TARIFSPECIAL','a17',671);
    __setListes('YTC_INCOTERM','a3',688);
    __setListes('YTC_MODEEXP','a3',691);
    __setListes('YTC_LIEUDISPO','a3',694);
    __setListes('YTC_TIERSLIVRE','a17',697);
    __setListes('YTC_NADRESSELIV','n11',714);
    __setListes('YTC_NADRESSEFAC','n11',725);
    __setListes('YTC_STATIONEDI','a17',736);
    __setListes('YTC_NOTRECODETIERS','a17',753);
    __setListes('YTC_NOTRECODCOMPTA','a17',770);
    __setListes('YTC_SCHEMAGEN','a17',787);
    __setListes('YTC_ACCELERATEUR','b1',804);
    __setListes('YTC_DAS2','b1',805);
    __setListes('YTC_PROFESSION','a35',806);
    __setListes('YTC_REMUNERATION','a3',841);
    __setListes('YTC_INDEMNITE','a3',844);
    __setListes('YTC_AVANTAGE','a3',847);
    __setListes('YTC_TYPEFOURNI','a35',850);
    __setListes('YTC_SECTEURGEO','a3',885);
    __setListes('YTC_TIMBRE','n20.2',888);
    __setListes('YTC_SURTAXE','n20.2',908);
    __setListes('YTC_MODELEBON','a3',928);
    __setListes('YTC_EDITRA','b1',931);
    __setListes('YTC_MODELETXT','a3',932);
    __setListes('YTC_TIERSEXPE','a40',935);
    __setListes('YTC_QUALIFPOIDS','a3',975);
    __setListes('YTC_DOUBLON','a3',978);
    __setListes('YTC_APPORTEUR2','a17',981);
    __setListes('YTC_APPORTEUR3','a17',998);
    __setListes('YTC_ASEGMENTER','b1',1015);
    __setListes('YTC_ASEGMENTERGRP','b1',1016);
    __setListes('YTC_PARTENAIRE','b1',1017);
    __setListes('YTC_AVECSOUSTRAIT','b1',1018);
    tabOffset_[85] := 1019;
    end
  else if ident = 'ANN' then
    begin
    result := 'ANNUAIRE';
    setLength(tabOffset_, 113);
    __setListes('ANN_CODEPER','n11',7);
    __setListes('ANN_DATECREATION','d8',18);
    __setListes('ANN_DATEMODIF','d8',26);
    __setListes('ANN_UTILISATEUR','a3',34);
    __setListes('ANN_APPMODIF','a17',37);
    __setListes('ANN_TIERS','a17',54);
    __setListes('ANN_ETATCPTABLE','a3',71);
    __setListes('ANN_GRPCONF','a3',74);
    __setListes('ANN_NOMPER','a25',77);
    __setListes('ANN_FAMPER','a3',102);
    __setListes('ANN_TYPEPER','a17',105);
    __setListes('ANN_PPPM','a3',122);
    __setListes('ANN_FORME','a9',125);
    __setListes('ANN_FORMEGEN','a9',134);
    __setListes('ANN_GRPSOC','a17',143);
    __setListes('ANN_NOM1','a35',160);
    __setListes('ANN_NOM2','a35',195);
    __setListes('ANN_NOM3','a35',230);
    __setListes('ANN_CV','a17',265);
    __setListes('ANN_CVA','a9',282);
    __setListes('ANN_CVL','a35',291);
    __setListes('ANN_TYPECIV','a9',326);
    __setListes('ANN_NATIONALITE','a3',335);
    __setListes('ANN_LANGUE','a3',338);
    __setListes('ANN_DEVISE','a3',341);
    __setListes('ANN_VOIENO','a5',344);
    __setListes('ANN_VOIENOCOMPL','a3',349);
    __setListes('ANN_VOIETYPE','a3',352);
    __setListes('ANN_VOIENOM','a35',355);
    __setListes('ANN_ALRUE1','a35',390);
    __setListes('ANN_ALRUE2','a35',425);
    __setListes('ANN_ALRUE3','a35',460);
    __setListes('ANN_ALVILLE','a35',495);
    __setListes('ANN_ALCP','a9',530);
    __setListes('ANN_PAYS','a3',539);
    __setListes('ANN_ALRESID','a17',542);
    __setListes('ANN_ALBAT','a4',559);
    __setListes('ANN_ALESC','a4',563);
    __setListes('ANN_ALETA','a4',567);
    __setListes('ANN_ALNOAPP','a4',571);
    __setListes('ANN_APNOM','a35',575);
    __setListes('ANN_APRUE1','a35',610);
    __setListes('ANN_APRUE2','a35',645);
    __setListes('ANN_APRUE3','a35',680);
    __setListes('ANN_APCPVILLE','a35',715);
    __setListes('ANN_APPAYS','a35',750);
    __setListes('ANN_CHGTADR','b1',785);
    __setListes('ANN_TEL1','a25',786);
    __setListes('ANN_TEL2','a25',811);
    __setListes('ANN_FAX','a25',836);
    __setListes('ANN_MINITEL','a25',861);
    __setListes('ANN_EMAIL','a250',886);
    __setListes('ANN_SITEWEB','a250',1136);
    __setListes('ANN_SEXE','a3',1386);
    __setListes('ANN_DATENAIS','d8',1389);
    __setListes('ANN_VILLENAIS','a35',1397);
    __setListes('ANN_NODEPTNAIS','a3',1432);
    __setListes('ANN_DEPTNAIS','a35',1435);
    __setListes('ANN_PROFESSION','a35',1470);
    __setListes('ANN_SITUFAM','a3',1505);
    __setListes('ANN_REGMAT','a3',1508);
    __setListes('ANN_CODECJ','n11',1511);
    __setListes('ANN_EVTFAMDATE','d8',1522);
    __setListes('ANN_CAPITAL','n20.2',1530);
    __setListes('ANN_CAPDEV','a3',1550);
    __setListes('ANN_CAPLIB','n20.2',1553);
    __setListes('ANN_CAPNBTITRE','n20.2',1573);
    __setListes('ANN_CAPVN','n20.2',1593);
    __setListes('ANN_SIREN','a17',1613);
    __setListes('ANN_CLESIRET','a5',1630);
    __setListes('ANN_RCS','b1',1635);
    __setListes('ANN_RCSVILLE','a35',1636);
    __setListes('ANN_RCSDATE','d8',1671);
    __setListes('ANN_RCSGEST','a35',1679);
    __setListes('ANN_RCSNOREF','n11',1714);
    __setListes('ANN_RM','b1',1725);
    __setListes('ANN_RMDEP','a3',1726);
    __setListes('ANN_RMANNEE','a4',1729);
    __setListes('ANN_RMNOREF','n11',1733);
    __setListes('ANN_CODENAFOLD','a4',1744);
    __setListes('ANN_DEBACT','d8',1748);
    __setListes('ANN_MOISCLOTURE','a3',1756);
    __setListes('ANN_ENSEIGNE','a70',1759);
    __setListes('ANN_PERASS1CODE','n11',1829);
    __setListes('ANN_PERASS1QUAL','a70',1840);
    __setListes('ANN_PERASS2CODE','n11',1910);
    __setListes('ANN_PERASS2QUAL','a70',1921);
    __setListes('ANN_NOPER','n11',1991);
    __setListes('ANN_CODEINSTIT','a4',2002);
    __setListes('ANN_REGROUPEMENT','a4',2006);
    __setListes('ANN_NOADMIN','a25',2010);
    __setListes('ANN_COMPLTNOADMIN','a25',2035);
    __setListes('ANN_NOIDENTIF','a20',2060);
    __setListes('ANN_NOTEOUV','b1',2080);
    __setListes('ANN_NATUREAUXI','a3',2081);
    __setListes('ANN_CONFIDENTIEL','b1',2084);
    __setListes('ANN_CLETELEPHONE','a25',2085);
    __setListes('ANN_CLETELEPHONE2','a25',2110);
    __setListes('ANN_GUIDPER','a36',2135);
    __setListes('ANN_CODEINSEE','a4',2171);
    __setListes('ANN_FORMEGRPPRIVE','a3',2175);
    __setListes('ANN_FORMESTE','a3',2178);
    __setListes('ANN_FORMESCI','a3',2181);
    __setListes('ANN_COOP','b1',2184);
    __setListes('ANN_NOM4','a35',2185);
    __setListes('ANN_PAYSNAIS','a3',2220);
    __setListes('ANN_TYPESCI','a3',2223);
    __setListes('ANN_CODENAF2OLD','a4',2226);
    __setListes('ANN_GUIDCJ','a36',2230);
    __setListes('ANN_CODENAF','a5',2266);
    __setListes('ANN_CODENAF2','a5',2271);
    __setListes('ANN_NUMEROSS','a15',2276);
    tabOffset_[112] := 2291;
    end
  else if ident = 'ANB' then
    begin
    result := 'ANNUBIS';
    setLength(tabOffset_, 34);
    __setListes('ANB_CODEPER','n11',7);
    __setListes('ANB_DATECREATION','d8',18);
    __setListes('ANB_DATEMODIF','d8',26);
    __setListes('ANB_UTILISATEUR','a3',34);
    __setListes('ANB_APPMODIF','a17',37);
    __setListes('ANB_OLDVOIENO','a5',54);
    __setListes('ANB_OLDVOIENOCOMPL','a3',59);
    __setListes('ANB_OLDVOIETYPE','a3',62);
    __setListes('ANB_OLDVOIENOM','a35',65);
    __setListes('ANB_OLDRUE1','a35',100);
    __setListes('ANB_OLDRUE2','a35',135);
    __setListes('ANB_OLDRUE3','a35',170);
    __setListes('ANB_OLDCP','a9',205);
    __setListes('ANB_OLDVILLE','a35',214);
    __setListes('ANB_GESTIONFISC','b1',249);
    __setListes('ANB_GESTIONPATRIM','b1',250);
    __setListes('ANB_COTISETNS','b1',251);
    __setListes('ANB_VALLIBRE1','n20.2',252);
    __setListes('ANB_VALLIBRE2','n20.2',272);
    __setListes('ANB_VALLIBRE3','n20.2',292);
    __setListes('ANB_DATELIBRE1','d8',312);
    __setListes('ANB_DATELIBRE2','d8',320);
    __setListes('ANB_DATELIBRE3','d8',328);
    __setListes('ANB_CHOIXLIBRE1','a3',336);
    __setListes('ANB_CHOIXLIBRE2','a3',339);
    __setListes('ANB_CHOIXLIBRE3','a3',342);
    __setListes('ANB_CHARLIBRE1','a35',345);
    __setListes('ANB_CHARLIBRE2','a35',380);
    __setListes('ANB_CHARLIBRE3','a35',415);
    __setListes('ANB_BOOLLIBRE1','b1',450);
    __setListes('ANB_BOOLLIBRE2','b1',451);
    __setListes('ANB_BOOLLIBRE3','b1',452);
    __setListes('ANB_GUIDPER','a36',453);
    tabOffset_[33] := 489;
    end
  else if ident = 'ANL' then
    begin
    result := 'ANNULIEN';
    setLength(tabOffset_, 113);
    __setListes('ANL_CODEPER','n11',7);
    __setListes('ANL_CODEPERDOS','n11',18);
    __setListes('ANL_TYPEDOS','a3',29);
    __setListes('ANL_NOORDRE','n11',32);
    __setListes('ANL_CODEDOS','a8',43);
    __setListes('ANL_DATECREATION','d8',51);
    __setListes('ANL_DATEMODIF','d8',59);
    __setListes('ANL_UTILISATEUR','a3',67);
    __setListes('ANL_APPMODIF','a17',70);
    __setListes('ANL_FONCTION','a9',87);
    __setListes('ANL_FORME','a9',96);
    __setListes('ANL_ORGANISME','a4',105);
    __setListes('ANL_NOMPER','a25',109);
    __setListes('ANL_TYPEPER','a17',134);
    __setListes('ANL_TYPEPERDOS','a17',151);
    __setListes('ANL_RACINE','a4',168);
    __setListes('ANL_TIERS','b1',172);
    __setListes('ANL_INFO','a17',173);
    __setListes('ANL_AFFICHE','a9',190);
    __setListes('ANL_TRI','n11',199);
    __setListes('ANL_EFDEB','d8',210);
    __setListes('ANL_EFFIN','d8',218);
    __setListes('ANL_REFLIEN','a200',226);
    __setListes('ANL_TTNBPP','n11',426);
    __setListes('ANL_TTNBUS','n11',437);
    __setListes('ANL_TTNBNP','n11',448);
    __setListes('ANL_TTNBTOT','n11',459);
    __setListes('ANL_TTNBTOTUS','n11',470);
    __setListes('ANL_TTMONTANT','n20.2',481);
    __setListes('ANL_TTPCTCAP','n20.9',501);
    __setListes('ANL_TTPCTBENEF','n20.9',521);
    __setListes('ANL_TTPCTVOIX','n20.9',541);
    __setListes('ANL_TTNO','a70',561);
    __setListes('ANL_VOIXAGO','n20.2',631);
    __setListes('ANL_VOIXAGE','n20.2',651);
    __setListes('ANL_INDIVIS','b1',671);
    __setListes('ANL_PERASS1CODE','n11',672);
    __setListes('ANL_PERASS1QUAL','a70',683);
    __setListes('ANL_REPRESENTANT','a35',753);
    __setListes('ANL_PERASS2CODE','n11',788);
    __setListes('ANL_PERASS2QUAL','a70',799);
    __setListes('ANL_PERASS3CODE','n11',869);
    __setListes('ANL_PERASS3QUAL','a70',880);
    __setListes('ANL_CONV','b1',950);
    __setListes('ANL_CONVTXT','a250',951);
    __setListes('ANL_CONVDATE','d8',1201);
    __setListes('ANL_CONVDATEAPP','d8',1209);
    __setListes('ANL_CONVLIB','b1',1217);
    __setListes('ANL_CONVSUITE','b1',1218);
    __setListes('ANL_TRAVCONT','b1',1219);
    __setListes('ANL_TRAVTXT','a70',1220);
    __setListes('ANL_TXDETDIRECT','n20.9',1290);
    __setListes('ANL_TXDETINDIRECT','n20.9',1310);
    __setListes('ANL_TXDETTOTAL','n20.9',1330);
    __setListes('ANL_METHODECONSO','a35',1350);
    __setListes('ANL_GRPFISCAL','b1',1385);
    __setListes('ANL_NOADHESION','a200',1386);
    __setListes('ANL_NODHCOMPL','a20',1586);
    __setListes('ANL_NUMAFFILIATION','a20',1606);
    __setListes('ANL_LIEUADHESION','a35',1626);
    __setListes('ANL_INTERLOCUTEUR','a35',1661);
    __setListes('ANL_BANQCODE','a5',1696);
    __setListes('ANL_BANQGUICHET','a5',1701);
    __setListes('ANL_BANQCOMPTE','a17',1706);
    __setListes('ANL_BANQCLERIB','a2',1723);
    __setListes('ANL_BANQCOND','a35',1725);
    __setListes('ANL_BANQTXDEB','n20.9',1760);
    __setListes('ANL_BANQDECOUVERT','n20.2',1780);
    __setListes('ANL_SATISFCLIENT','a35',1800);
    __setListes('ANL_SATISFCAB','a35',1835);
    __setListes('ANL_PERIODICITDUC','a3',1870);
    __setListes('ANL_REGIMEFISC','a3',1873);
    __setListes('ANL_DIVANNEE','a4',1876);
    __setListes('ANL_DIVBRUT','n20.2',1880);
    __setListes('ANL_DIVNET','n20.2',1900);
    __setListes('ANL_DIVAF','n20.2',1920);
    __setListes('ANL_NOMDATE','d8',1940);
    __setListes('ANL_NOMACTE','a3',1948);
    __setListes('ANL_RENDATE','d8',1951);
    __setListes('ANL_EXPDATE','d8',1959);
    __setListes('ANL_EXPEXE','a4',1967);
    __setListes('ANL_COOPTPROV','b1',1971);
    __setListes('ANL_COOPTCODE','n11',1972);
    __setListes('ANL_MDSALAIRE','n20.2',1983);
    __setListes('ANL_MDEXP','b1',2003);
    __setListes('ANL_MDRN','b1',2004);
    __setListes('ANL_MDRP','b1',2005);
    __setListes('ANL_MDNRNR','b1',2006);
    __setListes('ANL_MDRPNO','n11',2007);
    __setListes('ANL_MDRPCODE','n11',2018);
    __setListes('ANL_CRSPNO','n11',2029);
    __setListes('ANL_CRSPCODE','n11',2040);
    __setListes('ANL_CJASS','b1',2051);
    __setListes('ANL_APPNAT','b1',2052);
    __setListes('ANL_APPNATM','n20.2',2053);
    __setListes('ANL_APPNATT','n20.2',2073);
    __setListes('ANL_APPNUMM','n20.2',2093);
    __setListes('ANL_APPNUMT','n20.2',2113);
    __setListes('ANL_APPNUML','n20.2',2133);
    __setListes('ANL_ACTNAT','a3',2153);
    __setListes('ANL_ACTDATE','d8',2156);
    __setListes('ANL_ACTNBRE','n20.2',2164);
    __setListes('ANL_ACTMONT','n20.2',2184);
    __setListes('ANL_TITNATURE','a3',2204);
    __setListes('ANL_TITTYPE','a3',2207);
    __setListes('ANL_GUIDPER','a36',2210);
    __setListes('ANL_GUIDPERDOS','a36',2246);
    __setListes('ANL_DECLARATION','a3',2282);
    __setListes('ANL_ESTNONCADRE','b1',2285);
    __setListes('ANL_EVOLREV','n19.4',2286);
    __setListes('ANL_ALSACEMOSELLE','b1',2305);
    __setListes('ANL_GUIDCONV','a36',2306);
    tabOffset_[112] := 2342;
    end
  else if ident = 'US'  then
    begin
    result := 'UTILISAT';
    setLength(tabOffset_, 29);
    __setListes('US_UTILISATEUR','a3',7);
    __setListes('US_LIBELLE','a35',10);
    __setListes('US_ABREGE','a35',45);
    __setListes('US_CONTROLEUR','b1',80);
    __setListes('US_PASSWORD','a35',81);
    __setListes('US_PRESENT','b1',116);
    __setListes('US_DATECONNEXION','d8',117);
    __setListes('US_AUXILIAIRE','a17',125);
    __setListes('US_GROUPE','a3',142);
    __setListes('US_SUPERVISEUR','b1',145);
    __setListes('US_FONCTION','a100',146);
    __setListes('US_SUIVILOG','b1',246);
    __setListes('US_QRCOULEUR','b1',247);
    __setListes('US_EMAIL','a250',248);
    __setListes('US_EMAILPASSWORD','a35',498);
    __setListes('US_EMAILPOPSERVER','a128',533);
    __setListes('US_EMAILSMTPSERVER','a128',661);
    __setListes('US_EMAILLOGIN','a250',789);
    __setListes('US_NBTENTATIVE','n11',1039);
    __setListes('US_NBCONNEXION','n11',1050);
    __setListes('US_DATECHANGEPWD','d8',1061);
    __setListes('US_DESACTIVE','b1',1069);
    __setListes('US_MULTICONNEXION','b1',1070);
    __setListes('US_DOMAINEAUTHENT','a70',1071);
    __setListes('US_AUTHENTNT','a3',1141);
    __setListes('US_SSO','a3',1144);
    __setListes('US_AUTHENTFORTE','a3',1147);
    __setListes('US_AUTHENTPARAM','a250',1150);
    tabOffset_[28] := 1400;
    end
  else if ident = 'C'   then
    begin
    result := 'CONTACT';
    setLength(tabOffset_, 61);
    __setListes('C_TYPECONTACT','a3',7);
    __setListes('C_AUXILIAIRE','a36',10);
    __setListes('C_TIERS','a17',46);
    __setListes('C_NUMEROCONTACT','n11',63);
    __setListes('C_NATUREAUXI','a3',74);
    __setListes('C_PRINCIPAL','b1',77);
    __setListes('C_NOM','a35',78);
    __setListes('C_PRENOM','a35',113);
    __setListes('C_SERVICE','a35',148);
    __setListes('C_FONCTION','a35',183);
    __setListes('C_TELEPHONE','a25',218);
    __setListes('C_CLETELEPHONE','a25',243);
    __setListes('C_FAX','a25',268);
    __setListes('C_CLEFAX','a25',293);
    __setListes('C_TELEX','a25',318);
    __setListes('C_CLETELEX','a25',343);
    __setListes('C_RVA','a250',368);
    __setListes('C_SOCIETE','a3',618);
    __setListes('C_CIVILITE','a3',621);
    __setListes('C_FONCTIONCODEE','a3',624);
    __setListes('C_LIPARENT','a3',627);
    __setListes('C_JOURNAIS','n11',630);
    __setListes('C_MOISNAIS','n11',641);
    __setListes('C_ANNEENAIS','n11',652);
    __setListes('C_SEXE','a3',663);
    __setListes('C_PUBLIPOSTAGE','b1',666);
    __setListes('C_TEXTELIBRE1','a35',667);
    __setListes('C_TEXTELIBRE2','a35',702);
    __setListes('C_TEXTELIBRE3','a35',737);
    __setListes('C_BOOLLIBRE1','b1',772);
    __setListes('C_BOOLLIBRE2','b1',773);
    __setListes('C_BOOLLIBRE3','b1',774);
    __setListes('C_LIBRECONTACT1','a3',775);
    __setListes('C_LIBRECONTACT2','a3',778);
    __setListes('C_LIBRECONTACT3','a3',781);
    __setListes('C_LIBRECONTACT4','a3',784);
    __setListes('C_LIBRECONTACT5','a3',787);
    __setListes('C_LIBRECONTACT6','a3',790);
    __setListes('C_LIBRECONTACT7','a3',793);
    __setListes('C_LIBRECONTACT8','a3',796);
    __setListes('C_LIBRECONTACT9','a3',799);
    __setListes('C_LIBRECONTACTA','a3',802);
    __setListes('C_DATELIBRE1','d8',805);
    __setListes('C_DATELIBRE2','d8',813);
    __setListes('C_DATELIBRE3','d8',821);
    __setListes('C_VALLIBRE1','n20.2',829);
    __setListes('C_VALLIBRE2','n20.2',849);
    __setListes('C_VALLIBRE3','n20.2',869);
    __setListes('C_LIENTIERS','a17',889);
    __setListes('C_EMAILING','b1',906);
    __setListes('C_FERME','b1',907);
    __setListes('C_DATEFERMETURE','d8',908);
    __setListes('C_SERVICECODE','a3',916);
    __setListes('C_DATECREATION','d8',919);
    __setListes('C_DATEMODIF','d8',927);
    __setListes('C_CREATEUR','a3',935);
    __setListes('C_UTILISATEUR','a3',938);
    __setListes('C_GUIDPER','a36',941);
    //__setListes('C_BLOCNOTE','t1500',977);
    __setListes('C_GUIDPERANL','a36',977);
    __setListes('C_NUMEROADRESSE','n11',1013);
    tabOffset_[60] := 1024;
    end
  else if ident = 'DOS' then
    begin
    result := 'DOSSIER';
    setLength(tabOffset_, 35);
    __setListes('DOS_NODOSSIER','a8',7);
    __setListes('DOS_CODEPER','n11',15);
    __setListes('DOS_SOCIETE','a3',26);
    __setListes('DOS_LIBELLE','a35',29);
    __setListes('DOS_NODISQUE','n11',64);
    __setListes('DOS_NODISQUELOC','n11',75);
    __setListes('DOS_LASERIE','a2',86);
    __setListes('DOS_GROUPECONF','a3',88);
    __setListes('DOS_VERROU','a3',91);
    __setListes('DOS_DATEDEPART','d8',94);
    __setListes('DOS_DATERETOUR','d8',102);
    __setListes('DOS_UTILISATEUR','a3',110);
    __setListes('DOS_CABINET','b1',113);
    __setListes('DOS_NETEXPERT','b1',114);
    __setListes('DOS_NECPSEQ','n11',115);
    __setListes('DOS_NERECDATE','d8',126);
    __setListes('DOS_NERECNBFIC','n11',134);
    __setListes('DOS_NECPDATEARRET','d8',145);
    __setListes('DOS_PASSWORD','a17',153);
    __setListes('DOS_PWDGLOBAL','b1',170);
    __setListes('DOS_APPLISPROTEC','a255',171);
    __setListes('DOS_NETEXPERTGED','b1',426);
    __setListes('DOS_NECDKEY','a16',427);
    __setListes('DOS_USRS1','a17',443);
    __setListes('DOS_PWDS1','a17',460);
    __setListes('DOS_EWSCREE','b1',477);
    __setListes('DOS_SERIAS1','b1',478);
    __setListes('DOS_WINSTALL','a8',479);
    __setListes('DOS_TYPEMAJ','a3',487);
    __setListes('DOS_DETAILMAJ','a35',490);
    __setListes('DOS_GUIDPER','a36',525);
    __setListes('DOS_ABSENT','a1',561);
    __setListes('DOS_WINOUV','a3',562);
    __setListes('DOS_WINSTR','a35',565);
    tabOffset_[34] := 600;
    end
  else if ident = 'DAP' then
    begin
    result := 'DOSSAPPLI';
    setLength(tabOffset_, 5);
    __setListes('DAP_NODOSSIER', 'a8', 7);
    __setListes('DAP_NOMEXEC', 'a35', 15);
    __setListes('DAP_DATEMODIF', 'd8', 50);
    __setListes('DAP_UTILISATEUR', 'a3', 58);
    tabOffset_[4] := 61;
    end
  else if ident = 'DOG' then
    begin
    result := 'DOSSIERGRP';
    setLength(tabOffset_, 3);
    __setListes('DOG_NODOSSIER', 'a17', 7);
    __setListes('DOG_GROUPECONF', 'a17', 24);
    tabOffset_[2] := 41;
    end
{ else if ident = 'GRP' then
    begin
    result := 'GRPDONNEES';
    setLength(tabOffset_, 6);
    __setListes('GRP_ID', 'n11', 7);
    __setListes('GRP_IDPERE', 'n11', 18);
    __setListes('GRP_CODE', 'a17', 29);
    __setListes('GRP_LIBELLE', 'a35', 46);
    __setListes('GRP_NOM', 'a35', 81);
    tabOffset_[5] := 116;
    end
  else if ident = 'LDO' then
    begin
    result := 'LIENDOSGRP';
    setLength(tabOffset_, 5);
    __setListes('LDO_GRPID', 'n11', 7);
    __setListes('LDO_MARK', 'a1', 18);
    __setListes('LDO_NODOSSIER', 'a17', 19);
    __setListes('LDO_NOM', 'a35', 36);
    tabOffset_[4] := 71;
    end
    }
  else if ident = 'DCI' then
    begin
    result := 'DPCONTINTV';
    setLength(tabOffset_, 18);
    __setListes('DCI_NODP','n11',7);
    __setListes('DCI_NOORDRE','n11',18);
    __setListes('DCI_SALARIE','a17',29);
    __setListes('DCI_NOMCOLLAB','a35',46);
    __setListes('DCI_NODOSSREF','a17',81);
    __setListes('DCI_DATEINTERV','d8',98);
    __setListes('DCI_MODEINTERV','a3',106);
    __setListes('DCI_RESUME','a255',109);
    __setListes('DCI_PERSPECTIVES','a200',364);
    __setListes('DCI_MISSIONCOMPL','b1',564);
    __setListes('DCI_ARELANCER','b1',565);
    __setListes('DCI_DATERELANCE','d8',566);
    __setListes('DCI_DATESIGNAT','d8',574);
    __setListes('DCI_DATECREATION','d8',582);
    __setListes('DCI_DATEMODIF','d8',590);
    __setListes('DCI_UTILISATEUR','a3',598);
    __setListes('DCI_GUIDPER','a36',601);
    tabOffset_[17] := 637;
    end
  else if ident = 'DCL' then
    begin
    result := 'DPCONTROLE';
    setLength(tabOffset_, 20);
    __setListes('DCL_NODP','n11',7);
    __setListes('DCL_NOORDRE','n11',18);
    __setListes('DCL_TYPECTRL','a3',29);
    __setListes('DCL_DETAILCTRL','a3',32);
    __setListes('DCL_DATENOTIF','d8',35);
    __setListes('DCL_VERIFICATEUR','a35',43);
    __setListes('DCL_EXERCDEB','n11',78);
    __setListes('DCL_EXERCFIN','n11',89);
    __setListes('DCL_NATUREREDR','a35',100);
    __setListes('DCL_REDRENVISAG','n20.2',135);
    __setListes('DCL_REDRACCEPT','n20.2',155);
    __setListes('DCL_DATEPAIEMENT','d8',175);
    __setListes('DCL_ETATREDR','a35',183);
    __setListes('DCL_DATEFINREDR','d8',218);
    __setListes('DCL_DATECREATION','d8',226);
    __setListes('DCL_DATEMODIF','d8',234);
    __setListes('DCL_UTILISATEUR','a3',242);
    __setListes('DCL_NOREFORG','n11',245);
    __setListes('DCL_GUIDPER','a36',256);
    tabOffset_[19] := 292;
    end
  else if ident = 'DCV' then
    begin
    result := 'DPCONVENTION';
    setLength(tabOffset_, 18);
    __setListes('DCV_NODP','n11',7);
    __setListes('DCV_NOORDRE','n11',18);
    __setListes('DCV_STEPERSONNE','a35',29);
    __setListes('DCV_BENEFICIAIRE','a35',64);
    __setListes('DCV_OBJET','a35',99);
    __setListes('DCV_MODALITE','a35',134);
    __setListes('DCV_NATURECONV','a35',169);
    __setListes('DCV_DATEAUTOR','d8',204);
    __setListes('DCV_DATEAPPROB','d8',212);
    __setListes('DCV_COMMUNTOT','b1',220);
    __setListes('DCV_NATUREOPE','a3',221);
    __setListes('DCV_MTTFACTURE','n20.2',224);
    __setListes('DCV_MTTPAYE','n20.2',244);
    __setListes('DCV_DATECREATION','d8',264);
    __setListes('DCV_DATEMODIF','d8',272);
    __setListes('DCV_UTILISATEUR','a3',280);
    __setListes('DCV_GUIDPER','a36',283);
    tabOffset_[17] := 319;
    end
{  else if ident = 'DPD' then
    begin
    result := 'DPDOCUMENT';
    setLength(tabOffset_, 12);
    __setListes('DPD_DOCID','n11',7);
    __setListes('DPD_NODOSSIER','a8',18);
    __setListes('DPD_CODEGED','a3',26);
    __setListes('DPD_DATECREATION','d8',29);
    __setListes('DPD_DATEMODIF','d8',37);
    __setListes('DPD_UTILISATEUR','a3',45);
    __setListes('DPD_EWSID','a38',48);
    __setListes('DPD_EWSPUBLIABLE','b1',86);
    __setListes('DPD_EWSAPUBLIER','b1',87);
    __setListes('DPD_EWSDATEPUBL','d8',88);
    __setListes('DPD_EWSPUBLIEUR','a35',96);
    tabOffset_[11] := 131;
    end }
  else if ident = 'DEC' then
    begin
    result := 'DPECO';
    setLength(tabOffset_, 11);
    __setListes('DEC_NODP','n11',7);
    __setListes('DEC_CA','n20.2',18);
    __setListes('DEC_DATECA','d8',38);
    __setListes('DEC_ORIGCA','a20',46);
    __setListes('DEC_OLDCA','n20.2',66);
    __setListes('DEC_RESULTAT','n20.2',86);
    __setListes('DEC_DATERESULTAT','d8',106);
    __setListes('DEC_ORIGRESULTAT','a20',114);
    __setListes('DEC_OLDRESULTAT','n20.2',134);
    __setListes('DEC_GUIDPER','a36',154);
    tabOffset_[10] := 190;
    end
  else if ident = 'DFI' then
    begin
    result := 'DPFISCAL';
    setLength(tabOffset_, 106);
    __setListes('DFI_NODP','n11',7);
    __setListes('DFI_ACTIVFISC','a3',18);
    __setListes('DFI_NOFRP','a35',21);
    __setListes('DFI_IMPODIR','a3',56);
    __setListes('DFI_OPTIONAUTID','b1',59);
    __setListes('DFI_REGIMFISCDIR','a3',60);
    __setListes('DFI_OPTIONRDSUP','b1',63);
    __setListes('DFI_DATEOPTRDSUP','d8',64);
    __setListes('DFI_OPTIONRSS','b1',72);
    __setListes('DFI_DATEOPTRSS','d8',73);
    __setListes('DFI_OPTIONREPORT','b1',81);
    __setListes('DFI_DATEOPTREPORT','d8',82);
    __setListes('DFI_IMPOINDIR','a3',90);
    __setListes('DFI_OPTIONRISUP','b1',93);
    __setListes('DFI_EXIGIBILITE','a3',94);
    __setListes('DFI_OPTIONEXIG','b1',97);
    __setListes('DFI_DATEOPTEXIG','d8',98);
    __setListes('DFI_TYPECA12','a3',106);
    __setListes('DFI_TAUXTVA1','n20.9',109);
    __setListes('DFI_TAUXTVA2','n20.9',129);
    __setListes('DFI_TAUXTVA3','n20.9',149);
    __setListes('DFI_PERIODIIMPIND','a3',169);
    __setListes('DFI_JOURDECLA','n11',172);
    __setListes('DFI_REGLEMENTEURO','b1',183);
    __setListes('DFI_MODEPAIEFISC','a3',184);
    __setListes('DFI_NOINTRACOMM','a20',187);
    __setListes('DFI_TAXESALAIRES','b1',207);
    __setListes('DFI_PRORATATVA','n20.2',208);
    __setListes('DFI_DATEPRORTVA','d8',228);
    __setListes('DFI_PRORTVAREVIS','n20.2',236);
    __setListes('DFI_AUTRESTAXES','b1',256);
    __setListes('DFI_TYPETAXETVA','a3',257);
    __setListes('DFI_INTEGRAFISC','b1',260);
    __setListes('DFI_TETEGROUPE','b1',261);
    __setListes('DFI_NOREFTETEGRDP','n11',262);
    __setListes('DFI_EXONERE','b1',273);
    __setListes('DFI_DATEDEBEXO','d8',274);
    __setListes('DFI_DATEFINEXO','d8',282);
    __setListes('DFI_EXONERATION','a3',290);
    __setListes('DFI_EXONERETP','b1',293);
    __setListes('DFI_DATEDEBEXOTP','d8',294);
    __setListes('DFI_DATEFINEXOTP','d8',302);
    __setListes('DFI_EXONERATIONTP','a35',310);
    __setListes('DFI_ACPTEJUIN','b1',345);
    __setListes('DFI_DECLA1003R','b1',346);
    __setListes('DFI_CAMIONSCARS','b1',347);
    __setListes('DFI_TYPETAXEPRO','a3',348);
    __setListes('DFI_ACTIVTAXEPRO','a3',351);
    __setListes('DFI_COTISMIN','b1',354);
    __setListes('DFI_ALLEGETRANS','b1',355);
    __setListes('DFI_ABATTEFIXE','b1',356);
    __setListes('DFI_DEGREVTREDUC','b1',357);
    __setListes('DFI_MTTDEGREVTRED','n20.2',358);
    __setListes('DFI_ADHEREOGA','b1',378);
    __setListes('DFI_NOOGADP','n11',379);
    __setListes('DFI_NOADHOGA','a20',390);
    __setListes('DFI_DATEADHOGA','d8',410);
    __setListes('DFI_NODOSSINSPEC','a6',418);
    __setListes('DFI_NOINSPECTION','a6',424);
    __setListes('DFI_REGIMFUSION','b1',430);
    __setListes('DFI_DATEREGFUS','d8',431);
    __setListes('DFI_EXISTEPVLT','b1',439);
    __setListes('DFI_DROITSAPPORT','b1',440);
    __setListes('DFI_DATEAPPORT','d8',441);
    __setListes('DFI_CTRLFISC','b1',449);
    __setListes('DFI_DEMATERIATDFC','b1',450);
    __setListes('DFI_DATECONVTDFC','d8',451);
    __setListes('DFI_MONDECLAEURO','b1',459);
    __setListes('DFI_RESULTFISC','n20.2',460);
    __setListes('DFI_REINTEGR','n20.2',480);
    __setListes('DFI_REDUC','n20.2',500);
    __setListes('DFI_DATEMAJBENEF','d8',520);
    __setListes('DFI_ORIGMAJBENEF','a35',528);
    __setListes('DFI_OLDBENEFFISC','n20.2',563);
    __setListes('DFI_OLDREINTEGR','n20.2',583);
    __setListes('DFI_OLDREDUC','n20.2',603);
    __setListes('DFI_CA','n20.2',623);
    __setListes('DFI_ANNEECIVILE','n11',643);
    __setListes('DFI_ORIGMAJCA','a35',654);
    __setListes('DFI_OLDCA','n20.2',689);
    __setListes('DFI_DATECREATION','d8',709);
    __setListes('DFI_DATEMODIF','d8',717);
    __setListes('DFI_DATEOPTRG','d8',725);
    __setListes('DFI_UTILISATEUR','a3',733);
    __setListes('DFI_BENEFFISC','n20.2',736);
    __setListes('DFI_DECLARATION','b1',756);
    __setListes('DFI_CLEINSPECT','a3',757);
    __setListes('DFI_REGIMEINSPECT','a3',760);
    __setListes('DFI_REDEVABILITE','a1',763);
    __setListes('DFI_DOMHORSFRANCE','a3',764);
    __setListes('DFI_BAFORFAIT','a3',767);
    __setListes('DFI_DISTRIBDIVID','b1',770);
    __setListes('DFI_TAXEPROF','b1',771);
    __setListes('DFI_CONTREVENUSLOC','b1',772);
    __setListes('DFI_TAXEFONCIERE','b1',773);
    __setListes('DFI_CONTSOCSOLDOC','b1',774);
    __setListes('DFI_TAXEGRDSURF','b1',775);
    __setListes('DFI_TAXEANNIMM','b1',776);
    __setListes('DFI_TAXEVEHICSOC','b1',777);
    __setListes('DFI_VIGNETTEAUTO','b1',778);
    __setListes('DFI_IMPSOLFORTUNE','b1',779);
    __setListes('DFI_GUIDPER','a36',780);
    __setListes('DFI_GUIDPEROGA','a36',816);
    __setListes('DFI_REGLEFISC','a3',852);
    __setListes('DFI_TVATERR','a3',855);
    tabOffset_[105] := 858;
    end
{  else if ident = 'DIS' then
    begin
    result := 'DPINSTIT';
    setLength(tabOffset_,10);
    __setListes('DIS_NODP','n11',7);
    __setListes('DIS_CODEINSTIT','a4',18);
    __setListes('DIS_TYPECAISSE','a3',22);
    __setListes('DIS_NOADMIN','a25',25);
    __setListes('DIS_COMPLTNOADMIN','a25',50);
    __setListes('DIS_NOIDENTIF','a20',75);
    __setListes('DIS_DATECREATION','d8',95);
    __setListes('DIS_DATEMODIF','d8',103);
    __setListes('DIS_UTILISATEUR','a3',111);
    tabOffset_[9] := 114;
    end}
  else if ident = 'DPM' then
    begin
    result := 'DPMVTCAP';
    setLength(tabOffset_, 11);
    __setListes('DPM_NODP','n11',7);
    __setListes('DPM_NOORDRE','n11',18);
    __setListes('DPM_DATE','d8',29);
    __setListes('DPM_LIBELLE','a35',37);
    __setListes('DPM_SENS','a3',72);
    __setListes('DPM_NBTITRES','n11',75);
    __setListes('DPM_VALNOM','n20.2',86);
    __setListes('DPM_MONTANT','n20.2',106);
    __setListes('DPM_NATUREAUGM','a3',126);
    __setListes('DPM_GUIDPER','a36',129);
    tabOffset_[10] := 165;
    end
{  else if ident = 'DPO' then
    begin
    result := 'DPOBLIGATIONREALISE';
    setLength(tabOffset_, 7);
    __setListes('DPO_CODEOBLIGATION','a3',7);
    __setListes('DPO_NODOSSIER','a8',10);
    __setListes('DPO_DATEOBLIGATION','d8',18);
    __setListes('DPO_DATEREALISE','d8',26);
    __setListes('DPO_UTILISATEUR','a3',34);
    __setListes('DPO_COMMENTAIRE','a255',37);
    tabOffset_[6] := 292;
    end}
{  else if ident = 'DPT' then
    begin
    result := 'DPOBLIGATIONTMP';
    setLength(tabOffset_, 3);
    __setListes('DPT_DATEOBLIGATION','d8', 7);
    __setListes('DPT_CODEOBLIGATION','a3', 15);
    tabOffset_[2] := 18;
    end}
  else if ident = 'DOR' then
    begin
    result := 'DPORGA';
    setLength(tabOffset_, 76);
    __setListes('DOR_NODP','n11',7);
    __setListes('DOR_REGLEFISC','a3',18);
    __setListes('DOR_TYPESCI','a3',21);
    __setListes('DOR_DATEDEBUTEX','d8',24);
    __setListes('DOR_DATEFINEX','d8',32);
    __setListes('DOR_DUREE','n11',40);
    __setListes('DOR_DUREEPREC','n11',51);
    __setListes('DOR_LOCAGERANCE','b1',62);
    __setListes('DOR_NOREFPROPDP','n11',63);
    __setListes('DOR_DUREEVIE','n11',74);
    __setListes('DOR_DATECREAENT','d8',85);
    __setListes('DOR_TENUEEUROCPTA','b1',93);
    __setListes('DOR_TENUEEUROPAIE','b1',94);
    __setListes('DOR_TENUEEUROGC','a1',95);
    __setListes('DOR_ETABLTS','b1',96);
    __setListes('DOR_NBETABLTS','n11',97);
    __setListes('DOR_DOSSREF','b1',108);
    __setListes('DOR_NODOSSREF','a17',109);
    __setListes('DOR_CAUSERATTACH','a35',126);
    __setListes('DOR_ATTACHEMENT','b1',161);
    __setListes('DOR_NBRATTACH','n11',162);
    __setListes('DOR_DEPENDANCE','b1',173);
    __setListes('DOR_COMPTESCONSO','b1',174);
    __setListes('DOR_PROBLEME','b1',175);
    __setListes('DOR_TYPEPROBLEME','a3',176);
    __setListes('DOR_REGLEBNC','a3',179);
    __setListes('DOR_SECTIONBNC','a3',182);
    __setListes('DOR_CONVENTIONNE','b1',185);
    __setListes('DOR_HONORBNC','a3',186);
    __setListes('DOR_VALEURSNIR','n20.2',189);
    __setListes('DOR_DATESNIR','d8',209);
    __setListes('DOR_ORIGINESNIR','a35',217);
    __setListes('DOR_VALSNIRPREC','n20.2',252);
    __setListes('DOR_SALARPREP','b1',272);
    __setListes('DOR_SPECIALMED','a3',273);
    __setListes('DOR_FORMEASSO','a3',276);
    __setListes('DOR_CATEGASSO','a9',279);
    __setListes('DOR_SECTEURASSO','a3',288);
    __setListes('DOR_TYPEASSO','a3',291);
    __setListes('DOR_ETBLTSOINS','b1',294);
    __setListes('DOR_ADHECOTI','b1',295);
    __setListes('DOR_DONMANUEL','b1',296);
    __setListes('DOR_DONNATURE','b1',297);
    __setListes('DOR_VTEPUBLI','b1',298);
    __setListes('DOR_BIENFAIS','b1',299);
    __setListes('DOR_DONATION','b1',300);
    __setListes('DOR_SUBVENTION','b1',301);
    __setListes('DOR_BILLETERIE','b1',302);
    __setListes('DOR_REMBTPARMBS','b1',303);
    __setListes('DOR_DATECREATION','d8',304);
    __setListes('DOR_DATEMODIF','d8',312);
    __setListes('DOR_UTILISATEUR','a3',320);
    __setListes('DOR_EXERCICE','a3',323);
    __setListes('DOR_DATEENTREECAB','d8',326);
    __setListes('DOR_ORIGINECLT','a3',334);
    __setListes('DOR_DATESUPPR','d8',337);
    __setListes('DOR_MOTIFSUPPR','a3',345);
    __setListes('DOR_ETABLISSEMENT','a3',348);
    __setListes('DOR_FERME','b1',351);
    __setListes('DOR_UTILRESPCOMPTA','a3',352);
    __setListes('DOR_UTILRESPSOCIAL','a3',355);
    __setListes('DOR_UTILRESPJURID','a3',358);
    __setListes('DOR_UTILCHEFGROUPE','a3',361);
    __setListes('DOR_CABINETASSOC','a3',364);
    __setListes('DOR_NONTRAITE','b1',367);
    __setListes('DOR_MOTIFNONTRAITE','a25',368);
    __setListes('DOR_WETABENTITE','a3',393);
    __setListes('DOR_WCATREVENUS','a3',396);
    __setListes('DOR_WTYPEIMPO','a3',399);
    __setListes('DOR_WREGIMEIMPO','a3',402);
    __setListes('DOR_WREGIMETVA','a3',405);
    __setListes('DOR_WNATUREATTEST','a3',408);
    __setListes('DOR_WMODETRAITCPT','a3',411);
    __setListes('DOR_WTYPEBNC','a3',414);
    __setListes('DOR_GUIDPER','a36',417);
    tabOffset_[75] := 453;
    end
{  else if ident = 'DPA' then
    begin
    result := 'DPPATRIM';
    setLength(tabOffset_, 46);
    __setListes('DPA_NODP','n11',7);
    __setListes('DPA_NOPATRIM','n11',18);
    __setListes('DPA_CLASSEPAT','a3',29);
    __setListes('DPA_PATANNEXE','a17',32);
    __setListes('DPA_ORIGPROPENT','a3',49);
    __setListes('DPA_ORIGPROPSOR','a3',52);
    __setListes('DPA_PROPRIO','a3',55);
    __setListes('DPA_TYPEPROPRIO','a3',58);
    __setListes('DPA_QUOTITE','n20',61);
    __setListes('DPA_ISF','b1',81);
    __setListes('DPA_LIBELLE','a35',82);
    __setListes('DPA_ADRESSE1','a35',117);
    __setListes('DPA_ADRESSE2','a35',152);
    __setListes('DPA_ADRESSE3','a35',187);
    __setListes('DPA_CODEPOSTAL','a9',222);
    __setListes('DPA_VILLE','a35',231);
    __setListes('DPA_PAYS','a3',266);
    __setListes('DPA_VALEURACHAT','n20',269);
    __setListes('DPA_MONTANTN','n20',289);
    __setListes('DPA_ANNEE','n11',309);
    __setListes('DPA_MONTANTN1','n20',320);
    __setListes('DPA_MONTANTN2','n20',340);
    __setListes('DPA_MONTANTN3','n20',360);
    __setListes('DPA_MONTANTN4','n20',380);
    __setListes('DPA_PRIXVENTE','n20',400);
    __setListes('DPA_DEFIMMEUBLE','a3',420);
    __setListes('DPA_PATRIMOINE','a3',423);
    __setListes('DPA_LOCATAIRE','a35',426);
    __setListes('DPA_PERISSOL','b1',461);
    __setListes('DPA_PARTISF','n20',462);
    __setListes('DPA_CATEGORIEISF','a17',482);
    __setListes('DPA_NATBIEN','a35',499);
    __setListes('DPA_SUPERFICIE','a20',534);
    __setListes('DPA_SUPERFHABIT','a20',554);
    __setListes('DPA_NBPIECES','n11',574);
    __setListes('DPA_FERMAGE','n20',585);
    __setListes('DPA_DATEBAIL','d8',605);
    __setListes('DPA_DUREEBAIL','n11',613);
    __setListes('DPA_NATEXO','a3',624);
    __setListes('DPA_CREANCIER','a35',627);
    __setListes('DPA_DATEDEBUT','d8',662);
    __setListes('DPA_DATEFIN','d8',670);
    __setListes('DPA_DATECREATION','d8',678);
    __setListes('DPA_DATEMODIF','d8',686);
    __setListes('DPA_UTILISATEUR','a3',694);
    tabOffset_[45] := 697;
    end}
  else if ident = 'DPP' then
    begin
    result := 'DPPERSO';
    setlength(tabOffset_, 95);
    __setListes('DPP_NODP','n11',7);
    __setListes('DPP_NOMCJ','a35',18);
    __setListes('DPP_PRENOMCJ','a35',53);
    __setListes('DPP_DATENAISSCJ','d8',88);
    __setListes('DPP_VILLENAISSCJ','a35',96);
    __setListes('DPP_DEPTNAISSCJ','a9',131);
    __setListes('DPP_PROFESSIONCJ','a35',140);
    __setListes('DPP_NOMUSAGE','a35',175);
    __setListes('DPP_RESIDENCE','a30',210);
    __setListes('DPP_BATIMENT','a5',240);
    __setListes('DPP_ETAGE','a5',245);
    __setListes('DPP_ESCALIER','a5',250);
    __setListes('DPP_APPARTEMENT','a5',255);
    __setListes('DPP_REGIMEMATRI','a3',260);
    __setListes('DPP_EXPLOITDOM','b1',263);
    __setListes('DPP_VOIENO','a5',264);
    __setListes('DPP_VOIENOCOMPL','a3',269);
    __setListes('DPP_VOIETYPE','a3',272);
    __setListes('DPP_VOIENOM','a35',275);
    __setListes('DPP_ADRESSE1','a35',310);
    __setListes('DPP_ADRESSE2','a35',345);
    __setListes('DPP_ADRESSE3','a35',380);
    __setListes('DPP_CODEPOSTAL','a9',415);
    __setListes('DPP_VILLE','a35',424);
    __setListes('DPP_PAYS','a3',459);
    __setListes('DPP_GESTIONFISC','b1',462);
    __setListes('DPP_NODOSSFISC','a5',463);
    __setListes('DPP_PENSION','b1',468);
    __setListes('DPP_CJPENSION','b1',469);
    __setListes('DPP_75ANSCOMBAT','b1',470);
    __setListes('DPP_ENFDECEDES','b1',471);
    __setListes('DPP_ENFIMPOSES','b1',472);
    __setListes('DPP_ANNEENAISENF','a4',473);
    __setListes('DPP_PENSIONVEUVE','b1',477);
    __setListes('DPP_ENFCJDECEDE','a4',478);
    __setListes('DPP_CJDECESAN','b1',482);
    __setListes('DPP_PARENTISOLE','b1',483);
    __setListes('DPP_PERSACHARGE','b1',484);
    __setListes('DPP_ENFMIN','n11',485);
    __setListes('DPP_ENFMINHAND','n11',496);
    __setListes('DPP_AUTPERSHAND','n11',507);
    __setListes('DPP_ENFMAJ','n11',518);
    __setListes('DPP_ENFMARIES','n11',529);
    __setListes('DPP_IMPOTTOTAL','n20.2',540);
    __setListes('DPP_ANREFIMPOT','n11',560);
    __setListes('DPP_TAUXQF','n20.9',571);
    __setListes('DPP_NATREVENUS','a20',591);
    __setListes('DPP_ISF','b1',611);
    __setListes('DPP_FORFAITMOB','b1',612);
    __setListes('DPP_GESTIONPATRIM','b1',613);
    __setListes('DPP_PERSLIEES','b1',614);
    __setListes('DPP_COTISETNS','b1',615);
    __setListes('DPP_NODOSSTNS','a5',616);
    __setListes('DPP_TYPETNS','a3',621);
    __setListes('DPP_REVSOCN','n20.2',624);
    __setListes('DPP_REVSOCN1','n20.2',644);
    __setListes('DPP_REVSOCN2','n20.2',664);
    __setListes('DPP_REVSOCN3','n20.2',684);
    __setListes('DPP_REVSOCN4','n20.2',704);
    __setListes('DPP_COTISSOCN','n20.2',724);
    __setListes('DPP_DATEREVSOC','d8',744);
    __setListes('DPP_ORIGREVSOC','a20',752);
    __setListes('DPP_COTISSOCN1','n20.2',772);
    __setListes('DPP_COTISSOCN2','n20.2',792);
    __setListes('DPP_COTISSOCN3','n20.2',812);
    __setListes('DPP_COTISSOCN4','n20.2',832);
    __setListes('DPP_COTISFACN','n20.2',852);
    __setListes('DPP_STATUTCJ','a3',872);
    __setListes('DPP_DATECOTISSOC','d8',875);
    __setListes('DPP_ORIGCOTISSOC','a20',883);
    __setListes('DPP_COTISFACN1','n20.2',903);
    __setListes('DPP_COTISFACN2','n20.2',923);
    __setListes('DPP_COTISFACN3','n20.2',943);
    __setListes('DPP_COTISFACN4','n20.2',963);
    __setListes('DPP_EXONEREMALAD','b1',983);
    __setListes('DPP_DATECOTISFAC','d8',984);
    __setListes('DPP_ORIGCOTISFAC','a20',992);
    __setListes('DPP_PERIODIMALAD','a3',1012);
    __setListes('DPP_PERIODIVIEIL','a3',1015);
    __setListes('DPP_COMPLCOM','b1',1018);
    __setListes('DPP_DATECOMPLCOM','d8',1019);
    __setListes('DPP_CLASSCOMPLCOM','a20',1027);
    __setListes('DPP_EXONERECHOM','b1',1047);
    __setListes('DPP_ACTIVACCESS','b1',1048);
    __setListes('DPP_DATECREATION','d8',1049);
    __setListes('DPP_DATEMODIF','d8',1057);
    __setListes('DPP_UTILISATEUR','a3',1065);
    __setListes('DPP_STATUTCJCOLL','b1',1068);
    __setListes('DPP_PERIODICJVIEIL','a3',1069);
    __setListes('DPP_75ANSCOMBATDCD','a1',1072);
    __setListes('DPP_GUIDPER','a36',1073);
    __setListes('DPP_STATUTSOCIAL','a3',1109);
    __setListes('DPP_COUVSOCIALAUTO','b1',1112);
    __setListes('DPP_RETRAITE','n11',1113);
    tabOffset_[94] := 1124;
    end
  else if ident = 'DSO' then
    begin
    result := 'DPSOCIAL';
    setLength(tabOffset_, 109);
    __setListes('DSO_NODP','n11',7);
    __setListes('DSO_EFFECTIF','n11',18);
    __setListes('DSO_DATEEFFECTIF','d8',29);
    __setListes('DSO_ORIGEFFECTIF','a20',37);
    __setListes('DSO_EFFMOYEN','n11',57);
    __setListes('DSO_OLDEFFMOYEN','n11',68);
    __setListes('DSO_EFFHANDICAP','n11',79);
    __setListes('DSO_EFFAPPRENTIS','n11',90);
    __setListes('DSO_CE','b1',101);
    __setListes('DSO_PARTICIPATION','b1',102);
    __setListes('DSO_ACCORDS','b1',103);
    __setListes('DSO_PAIEDECALEE','b1',104);
    __setListes('DSO_MTTDADS','n20.2',105);
    __setListes('DSO_DATEDADS','d8',125);
    __setListes('DSO_ORIGDADS','a20',133);
    __setListes('DSO_MTTDAS2','n20.2',153);
    __setListes('DSO_CTRLSOC','b1',173);
    __setListes('DSO_PAIEANNEESOC','n11',174);
    __setListes('DSO_PAIEAPPOINTS','b1',185);
    __setListes('DSO_PAIESTATS','b1',186);
    __setListes('DSO_PAIEANALYTIQUE','b1',187);
    __setListes('DSO_PAIEGENECRIT','b1',188);
    __setListes('DSO_CONVENCOLLEC','a3',189);
    __setListes('DSO_DATEDEBEX','d8',192);
    __setListes('DSO_DATEFINEX','d8',200);
    __setListes('DSO_TAXESALARIE','b1',208); //LM20070412 champ obsolet
    __setListes('DSO_TXSALPERIODIC','a3',209);
    __setListes('DSO_TAXEAPPRENT','b1',212);
    __setListes('DSO_PARTFORMCONT','b1',213);
    __setListes('DSO_PARTCONSTRUC','b1',214);
    __setListes('DSO_PAIECAB','b1',215);
    __setListes('DSO_PAIEENT','b1',216);
    __setListes('DSO_PAIEENTSYS','a3',217);
    __setListes('DSO_REGPERS','a3',220);
    __setListes('DSO_DECUNEMB','a3',223);
    __setListes('DSO_ELTVARIAENT','b1',226);
    __setListes('DSO_DATEDERPAIE','d8',227);
    __setListes('DSO_GESTCONGES','b1',235);
    __setListes('DSO_MUTSOCAGR','b1',236);
    __setListes('DSO_INTERMSPEC','b1',237);
    __setListes('DSO_BTP','b1',238);
    __setListes('DSO_TICKETREST','b1',239);
    __setListes('DSO_GESTIONETS','b1',240);
    __setListes('DSO_TELEDADS','b1',241);
    __setListes('DSO_PLANPAIEACT','b1',242);
    __setListes('DSO_CDD','b1',243);
    __setListes('DSO_CDI','b1',244);
    __setListes('DSO_ABATFRAISPRO','b1',245);
    __setListes('DSO_TPSPARTIEL','b1',246);
    __setListes('DSO_TPSPARTIEL30','b1',247);
    __setListes('DSO_CIE','b1',248);
    __setListes('DSO_CEC','b1',249);
    __setListes('DSO_CES','b1',250);
    __setListes('DSO_CRE','b1',251);
    __setListes('DSO_EMBSAL1','b1',252);
    __setListes('DSO_EMBSAL23','b1',253);
    __setListes('DSO_CONTAPPRENT','b1',254);
    __setListes('DSO_CONTQUAL','b1',255);
    __setListes('DSO_CONTORIENT','b1',256);
    __setListes('DSO_EXOCHARGES','b1',257);
    __setListes('DSO_DATEEXSOC','d8',258);
    __setListes('DSO_COMITEENT','b1',266);
    __setListes('DSO_DELEGUEPERS','b1',267);
    __setListes('DSO_DELEGUESYND','b1',268);
    __setListes('DSO_EXISTTICKREST','b1',269);
    __setListes('DSO_TAUXACCTRAV','n20.2',270);
    __setListes('DSO_VERSETRANS','b1',290);
    __setListes('DSO_PERURSSAF','a3',291);
    __setListes('DSO_PERASSEDIC','a3',294);
    __setListes('DSO_PERIRC','a3',297);
    __setListes('DSO_REMBULLPAIE','n11',300);
    __setListes('DSO_DECLHANDICAP','b1',311);
    __setListes('DSO_DECLMO','b1',312);
    __setListes('DSO_DATECREATION','d8',313);
    __setListes('DSO_DATEMODIF','d8',321);
    __setListes('DSO_UTILISATEUR','a3',329);
    __setListes('DSO_GUIDPER','a36',332);
    __setListes('DSO_RETCOLLECTIF','a36',368);
    __setListes('DSO_PREVCOLLECT','a36',369);
    __setListes('DSO_EPARGNESAL','a36',370);
    __setListes('DSO_RETCOLLECTIF39','a36',371);
    __setListes('DSO_IFC','a36',372);
    __setListes('DSO_STATUT','a3',373);
    __setListes('DSO_CONJOINTAVEC','b1',376);
    __setListes('DSO_CONTRATRET','b1',377);
    __setListes('DSO_CONTRATPREVOY','b1',378);
    __setListes('DSO_CONJOINTSTATUT','a3',379);
    __setListes('DSO_TRANSPORT','b1',382);
    __setListes('DSO_TXTRANSPORT','n19.4',383);
    __setListes('DSO_ACCIDENT','b1',402);
    __setListes('DSO_TXACCIDENT','n19.4',403);
    __setListes('DSO_ART83','b1',422);
    __setListes('DSO_ART83SALTRA','n19.4',423);
    __setListes('DSO_ART83PATTRA','n19.4',442);
    __setListes('DSO_ART83SALTRB','n19.4',461);
    __setListes('DSO_ART83PATTRB','n19.4',480);
    __setListes('DSO_PREV','b1',499);
    __setListes('DSO_PREVSALTRA','n19.4',500);
    __setListes('DSO_PREVPATTRA','n19.4',519);
    __setListes('DSO_PREVSALTRB','n19.4',538);
    __setListes('DSO_PREVPATTRB','n19.4',557);
    __setListes('DSO_COMPSANTE','b1',576);
    __setListes('DSO_COMPSANTESAL','n19.4',577);
    __setListes('DSO_COMPSANTEPAT','n19.4',596);
    __setListes('DSO_ART83SALTRC','n19.4',615);
    __setListes('DSO_ART83PATTRC','n19.4',634);
    __setListes('DSO_PREVSALTRC','n19.4',653);
    __setListes('DSO_PREVPATTRC','n19.4',672);
    tabOffset_ [108] := 691;
    end
  else if ident = 'DSC' then
    begin
    result := 'DPSOCIALCAISSE';
    setLength(tabOffset_, 13);
    __setListes('DSC_NODP','n11',7);
    __setListes('DSC_CODECAISSE','a15',18);
    __setListes('DSC_ORDRE','n11',33);
    __setListes('DSC_NOMCAISSE','a25',44);
    __setListes('DSC_NATUREDUCS','a25',69);
    __setListes('DSC_PERIODICITE','a25',94);
    __setListes('DSC_GUIDPER','a36',119);
    __setListes('DSC_NUMAFFILIATION','a20',155);
    __setListes('DSC_NUMINTERNE','a35',175);
    __setListes('DSC_SIRET','a17',210);
    __setListes('DSC_INSTITUTION','a4',227);
    __setListes('DSC_REGROUPEMENT','a4',231);
    tabOffset_[12] := 235;
    end
  else if ident = 'DTC' then
    begin
    result := 'DPTABCOMPTA';
    setLength(tabOffset_, 23);
    __setListes('DTC_NODOSSIER','a8',7);
    __setListes('DTC_LIBEXERCICE','a35',15);
    __setListes('DTC_MILLESIME','a2',50);
    __setListes('DTC_DUREE','n11',52);
    __setListes('DTC_DATEDEB','d8',63);
    __setListes('DTC_DATEFIN','d8',71);
    __setListes('DTC_CA','n20.2',79);
    __setListes('DTC_MARGETOTALE','n20.2',99);
    __setListes('DTC_VALEURAJOUTEE','n20.2',119);
    __setListes('DTC_EXCEDBRUT','n20.2',139);
    __setListes('DTC_RESULTEXPLOIT','n20.2',159);
    __setListes('DTC_RESULTCOURANT','n20.2',179);
    __setListes('DTC_RESULTEXERC','n20.2',199);
    __setListes('DTC_NBECRITURE','n11',219);
    __setListes('DTC_ENCOURSSAISIE','d8',230);
    __setListes('DTC_DERNSAISIE','d8',238);
    __setListes('DTC_DERNJOURNAL','a3',246);
    __setListes('DTC_UTILISATSAISIE','a3',249);
    __setListes('DTC_DATESAISIE','d8',252);
    __setListes('DTC_NBENTREEIMMO','n11',260);
    __setListes('DTC_NBSORTIEIMMO','n11',271);
    __setListes('DTC_NBLIGNEIMMO','n11',282);
    tabOffset_[22] := 293;
    end
  else if ident = 'DT1' then
    begin
    result := 'DPTABGENPAIE';
    setLength(tabOffset_, 9);
    __setListes('DT1_NODOSSIER','a8',7);
    __setListes('DT1_MOIS','a2',15);
    __setListes('DT1_ANNEE','a4',17);
    __setListes('DT1_NBBULLETINS','n11',21);
    __setListes('DT1_TOTALBRUT','n20.2',32);
    __setListes('DT1_NBPRESENTS','n11',52);
    __setListes('DT1_NBENTREES','n11',63);
    __setListes('DT1_NBSORTIES','n11',74);
    tabOffset_[8] := 85;
    end
  else if ident = 'DTP' then
    begin
    result := 'DPTABPAIE';
    setLength(tabOffset_, 12);
    __setListes('DTP_NODOSSIER','a8',7);
    __setListes('DTP_LIBEXERCICE','a35',15);
    __setListes('DTP_MILLESIME','a2',50);
    __setListes('DTP_DATEDEB','d8',52);
    __setListes('DTP_DATEFIN','d8',60);
    __setListes('DTP_DECALAGEPAIE','b1',68);
    __setListes('DTP_MONTANTDADS','n20.2',69);
    __setListes('DTP_EFFECTIF','n11',89);
    __setListes('DTP_NBENTREES','n11',100);
    __setListes('DTP_NBSORTIES','n11',111);
    __setListes('DTP_MODEREGL','a3',122);
    tabOffset_[11] := 125;
    end
  else if ident = 'CC' then
    begin
    result := 'CHOIXCOD';
    SetLength(tabOffset_, 6);
    __setListes('CC_TYPE', 'a3', 7);
    __setListes('CC_CODE', 'a17', 10);
    __setListes('CC_LIBELLE', 'a105', 27);
    __setListes('CC_ABREGE', 'a17', 132);
    __setListes('CC_LIBRE', 'a35', 149);
    tabOffset_[5] := 184;
    end
  else if ident = 'CCG' then
    begin
    result := 'CHOIXCODGRP';
    SetLength(tabOffset_, 6);
    __setListes('CC_TYPE', 'a3', 7);
    __setListes('CC_CODE', 'a17', 10);
    __setListes('CC_LIBELLE', 'a105', 27);
    __setListes('CC_ABREGE', 'a17', 132);
    __setListes('CC_LIBRE', 'a35', 149);
    tabOffset_[5] := 184;
    end
  else if ident = 'YX' then
    begin
    result := 'CHOIXEXT';
    SetLength(tabOffset_, 6);
    __setListes('YX_TYPE', 'a3', 7);
    __setListes('YX_CODE', 'a17', 10);
    __setListes('YX_LIBELLE', 'a105', 27);
    __setListes('YX_ABREGE', 'a17', 132);
    __setListes('YX_LIBRE', 'a35', 149);
    tabOffset_[5] := 184;
    end
  else if Ident = 'ARS' then
    begin
     Result:='RESSOURCE';
     SetLength (TabOffset_,123);
     __setListes('ARS_TYPERESSOURCE','a3',7);
     __setListes('ARS_RESSOURCE','a17',10);
     __setListes('ARS_AUXILIAIRE','a17',27);
     __setListes('ARS_SALARIE','a17',44);
     __setListes('ARS_LIBELLE','a35',61);
     __setListes('ARS_LIBELLE2','a35',96);
     __setListes('ARS_ADRESSE1','a35',131);
     __setListes('ARS_ADRESSE2','a35',166);
     __setListes('ARS_ADRESSE3','a35',201);
     __setListes('ARS_CODEPOSTAL','a9',236);
     __setListes('ARS_VILLE','a35',245);
     __setListes('ARS_PAYS','a3',280);
     __setListes('ARS_TELEPHONE','a25',283);
     __setListes('ARS_TELEPHONE2','a25',308);
     __setListes('ARS_IMMAT','a35',333);
     __setListes('ARS_DEBUTDISPO','d8',368);
     __setListes('ARS_FINDISPO','d8',376);
     __setListes('ARS_TAUXUNIT','n15.2',384);
     __setListes('ARS_UNITETEMPS','a3',399);
     __setListes('ARS_TAUXCHARGEPAT','n15.2',402);
     __setListes('ARS_TAUXFRAISGEN1','n15.2',417);
     __setListes('ARS_TAUXFRAISGEN2','n15.2',432);
     __setListes('ARS_COEFMETIER','n15.2',447);
     __setListes('ARS_COEFPV','n15.2',462);
     __setListes('ARS_CALCULPR','b1',477);
     __setListes('ARS_CALCULPV','b1',478);
     __setListes('ARS_TARIFRESSOURCE','a3',479);
     __setListes('ARS_TAUXREVIENTUN','n15.2',482);
     __setListes('ARS_PVHT','n15.2',497);
     __setListes('ARS_PVTTC','n15.2',512);
     __setListes('ARS_PVHTCALCUL','n15.2',527);
     __setListes('ARS_DATEPRIX','d8',542);
     __setListes('ARS_TAUXREVIENTUN2','n15.2',550);
     __setListes('ARS_PVHT2','n15.2',565);
     __setListes('ARS_PVTTC2','n15.2',580);
     __setListes('ARS_PVHTCALCUL2','n15.2',595);
     __setListes('ARS_DATEPRIX2','d8',610);
     __setListes('ARS_TAUXREVIENTUN3','n15.2',618);
     __setListes('ARS_PVHT3','n15.2',633);
     __setListes('ARS_PVTTC3','n15.2',648);
     __setListes('ARS_PVHTCALCUL3','n15.2',663);
     __setListes('ARS_DATEPRIX3','d8',678);
     __setListes('ARS_TAUXREVIENTUN4','n15.2',686);
     __setListes('ARS_PVHT4','n15.2',701);
     __setListes('ARS_PVTTC4','n15.2',716);
     __setListes('ARS_PVHTCALCUL4','n15.2',731);
     __setListes('ARS_DATEPRIX4','d8',746);
     __setListes('ARS_COEFPRPV','n15.4',754);
     __setListes('ARS_MARGEMINI','n15.4',769);
     __setListes('ARS_QUALIFMARGE','a3',784);
     __setListes('ARS_UTILASSOCIE','a3',787);
     __setListes('ARS_GROUPERES','a3',790);
     __setListes('ARS_SOCIETE','a3',793);
     __setListes('ARS_FERME','b1',796);
     __setListes('ARS_FONCTION1','a17',797);
     __setListes('ARS_DATEFONC1','d8',814);
     __setListes('ARS_FONCTION2','a17',822);
     __setListes('ARS_DATEFONC2','d8',839);
     __setListes('ARS_FONCTION3','a17',847);
     __setListes('ARS_DATEFONC3','d8',864);
     __setListes('ARS_FONCTION4','a17',872);
     __setListes('ARS_DATEFONC4','d8',889);
     __setListes('ARS_CREATEUR','a3',897);
     __setListes('ARS_CREERPAR','a3',900);
     __setListes('ARS_UTILISATEUR','a3',903);
     __setListes('ARS_DATECREATION','d8',906);
     __setListes('ARS_DATEMODIF','d8',914);
     __setListes('ARS_VALLIBRE1','n15.2',922);
     __setListes('ARS_VALLIBRE2','n15.2',937);
     __setListes('ARS_VALLIBRE3','n15.2',952);
     __setListes('ARS_DATELIBRE1','d8',967);
     __setListes('ARS_DATELIBRE2','d8',975);
     __setListes('ARS_DATELIBRE3','d8',983);
     __setListes('ARS_LIBRERES1','a6',991);
     __setListes('ARS_LIBRERES2','a6',997);
     __setListes('ARS_LIBRERES3','a6',1003);
     __setListes('ARS_LIBRERES4','a6',1009);
     __setListes('ARS_LIBRERES5','a6',1015);
     __setListes('ARS_LIBRERES6','a6',1021);
     __setListes('ARS_LIBRERES7','a6',1027);
     __setListes('ARS_LIBRERES8','a6',1033);
     __setListes('ARS_LIBRERES9','a6',1039);
     __setListes('ARS_LIBRERESA','a6',1045);
     __setListes('ARS_CHARLIBRE1','a35',1051);
     __setListes('ARS_CHARLIBRE2','a35',1086);
     __setListes('ARS_CHARLIBRE3','a35',1121);
     __setListes('ARS_BOOLLIBRE1','b1',1156);
     __setListes('ARS_BOOLLIBRE2','b1',1157);
     __setListes('ARS_BOOLLIBRE3','b1',1158);
     __setListes('ARS_STANDCALEN','a3',1159);
     __setListes('ARS_ARTICLE','a35',1162);
     __setListes('ARS_ACTIVITEPAIE','b1',1197);
     __setListes('ARS_CALENSPECIF','b1',1198);
     __setListes('ARS_EMAIL','a25',1199);
     __setListes('ARS_DEPARTEMENT','a6',1224);
     __setListes('ARS_ETABLISSEMENT','a3',1230);
     __setListes('ARS_GENERIQUE','b1',1233);
     __setListes('ARS_RESSOURCELIE','a17',1234);
     __setListes('ARS_TAUXSIMUL','n15.2',1251);
     __setListes('ARS_SECTIONPDR','a17',1266);
     __setListes('ARS_RUBRIQUEPDR','a17',1283);
     __setListes('ARS_ESTHUMAIN','b1',1300);
     __setListes('ARS_DATESORTIE','d8',1301);
     __setListes('ARS_SERVCPTAPRINC1','a3',1309);
     __setListes('ARS_SERVCPTASEC1','a3',1312);
     __setListes('ARS_SERVCPTAPRINC2','a3',1315);
     __setListes('ARS_SERVCPTASEC2','a3',1318);
     __setListes('ARS_DATESERVCPTAP2','d8',1321);
     __setListes('ARS_DATESERVCPTAS2','d8',1329);
     __setListes('ARS_PRINCIPALE1','b1',1337);
     __setListes('ARS_PRINCIPALE2','b1',1338);
     __setListes('ARS_PRINCIPALE3','b1',1339);
     __setListes('ARS_PRINCIPALE4','b1',1340);
     __setListes('ARS_DATEFINFONC3','d8',1341);
     __setListes('ARS_DATEFINFONC4','d8',1349);
     __setListes('ARS_CODEFORFAIT','a3',1357);
     __setListes('ARS_COEFCORRFCT1','n15.9',1360);
     __setListes('ARS_COEFCORRFCT2','n15.9',1375);
     __setListes('ARS_COEFCORRFCT3','n15.9',1390);
     __setListes('ARS_COEFCORRFCT4','n15.9',1405);
     __setListes('ARS_REGION','a9',1420);
     __setListes('ARS_SECTEURGEO','a3',1429);
     tabOffset_[122]:=1432;
    end
  else if Ident = 'GA' then
    begin
     Result:='ARTICLE';
     SetLength (TabOffset_,141);
     __setListes('GA_ARTICLE','a35',7);
     __setListes('GA_CODEARTICLE','a18',42);
     __setListes('GA_TYPEARTICLE','a3',60);
     __setListes('GA_CODEBARRE','a18',63);
     __setListes('GA_QUALIFCODEBARRE','a3',81);
     __setListes('GA_FAMILLENIV1','a3',84);
     __setListes('GA_FAMILLENIV2','a3',87);
     __setListes('GA_FAMILLENIV3','a3',90);
     __setListes('GA_LIBELLE','a70',93);
     __setListes('GA_LIBCOMPL','a35',163);
     __setListes('GA_COMPTAARTICLE','a3',198);
     __setListes('GA_COMMENTAIRE','a200',201);
     __setListes('GA_DESIGNATION1','a100',401);
     __setListes('GA_DESIGNATION2','a100',501);
     __setListes('GA_DPA','n15.2',601);
     __setListes('GA_DPR','n15.2',616);
     __setListes('GA_PVHT','n15.2',631);
     __setListes('GA_PVTTC','n15.2',646);
     __setListes('GA_PAHT','n15.2',661);
     __setListes('GA_PRHT','n15.2',676);
     __setListes('GA_PMAP','n15.2',691);
     __setListes('GA_PMRP','n15.2',706);
     __setListes('GA_MARGEMINI','n15.2',721);
     __setListes('GA_QUALIFMARGE','a3',736);
     __setListes('GA_FERME','b1',739);
     __setListes('GA_INTERDITVENTE','b1',740);
     __setListes('GA_POIDSNET','n15.4',741);
     __setListes('GA_POIDSBRUT','n15.4',756);
     __setListes('GA_QUALIFPOIDS','a3',771);
     __setListes('GA_VOLUME','n15.4',774);
     __setListes('GA_QUALIFVOLUME','a3',789);
     __setListes('GA_LINEAIRE','n15.4',792);
     __setListes('GA_QUALIFLINEAIRE','a3',807);
     __setListes('GA_SURFACE','n15.4',810);
     __setListes('GA_QUALIFSURFACE','a3',825);
     __setListes('GA_HEURE','n15.4',828);
     __setListes('GA_QUALIFHEURE','a3',843);
     __setListes('GA_QUALIFUNITESTO','a3',846);
     __setListes('GA_QUALIFUNITEVTE','a3',849);
     __setListes('GA_QUALIFUNITEACT','a3',852);
     __setListes('GA_TARIFARTICLE','a3',855);
     __setListes('GA_PAYSORIGINE','a3',858);
     __setListes('GA_CODEDOUANIER','a17',861);
     __setListes('GA_POIDSDOUA','n15.4',878);
     __setListes('GA_PRIXPOURQTE','n15.2',893);
     __setListes('GA_DATECREATION','d8',908);
     __setListes('GA_DATEMODIF','d8',916);
     __setListes('GA_DATEINTEGR','d8',924);
     __setListes('GA_DATESUPPRESSION','d8',932);
     __setListes('GA_TENUESTOCK','b1',940);
     __setListes('GA_LOT','b1',941);
     __setListes('GA_NUMEROSERIE','b1',942);
     __setListes('GA_CONTREMARQUE','b1',943);
     __setListes('GA_QUALIFEMBALLAGE','a3',944);
     __setListes('GA_PCB','n15.2',947);
     __setListes('GA_FAMILLETAXE1','a3',962);
     __setListes('GA_FAMILLETAXE2','a3',965);
     __setListes('GA_FAMILLETAXE3','a3',968);
     __setListes('GA_FAMILLETAXE4','a3',971);
     __setListes('GA_FAMILLETAXE5','a3',974);
     __setListes('GA_ACTION','a3',977);
     __setListes('GA_CREATEUR','a3',980);
     __setListes('GA_SOCIETE','a3',983);
     __setListes('GA_CREERPAR','a3',986);
     __setListes('GA_VALLIBRE1','n15.2',989);
     __setListes('GA_VALLIBRE2','n15.2',1004);
     __setListes('GA_VALLIBRE3','n15.2',1019);
     __setListes('GA_DATELIBRE1','d8',1034);
     __setListes('GA_DATELIBRE2','d8',1042);
     __setListes('GA_DATELIBRE3','d8',1050);
     __setListes('GA_LIBREART1','a6',1058);
     __setListes('GA_LIBREART2','a6',1064);
     __setListes('GA_LIBREART3','a6',1070);
     __setListes('GA_LIBREART4','a6',1076);
     __setListes('GA_LIBREART5','a6',1082);
     __setListes('GA_LIBREART6','a6',1088);
     __setListes('GA_LIBREART7','a6',1094);
     __setListes('GA_LIBREART8','a6',1100);
     __setListes('GA_LIBREART9','a6',1106);
     __setListes('GA_LIBREARTA','a6',1112);
     __setListes('GA_CHARLIBRE1','a35',1118);
     __setListes('GA_CHARLIBRE2','a35',1153);
     __setListes('GA_CHARLIBRE3','a35',1188);
     __setListes('GA_BOOLLIBRE1','b1',1223);
     __setListes('GA_BOOLLIBRE2','b1',1224);
     __setListes('GA_BOOLLIBRE3','b1',1225);
     __setListes('GA_REMISEPIED','b1',1226);
     __setListes('GA_ESCOMPTABLE','b1',1227);
     __setListes('GA_COMMISSIONNABLE','b1',1228);
     __setListes('GA_DIMMASQUE','a3',1229);
     __setListes('GA_FONCTION','a35',1232);
     __setListes('GA_ACTIVITEEFFECT','b1',1267);
     __setListes('GA_ACTIVITEREPRISE','a3',1268);
     __setListes('GA_REMISELIGNE','b1',1271);
     __setListes('GA_COEFFG','n15.2',1272);
     __setListes('GA_TYPEEMPLACE','a3',1287);
     __setListes('GA_COLLECTION','a3',1290);
     __setListes('GA_CODEDIM1','a3',1293);
     __setListes('GA_CODEDIM2','a3',1296);
     __setListes('GA_CODEDIM3','a3',1299);
     __setListes('GA_CODEDIM4','a3',1302);
     __setListes('GA_CODEDIM5','a3',1305);
     __setListes('GA_GRILLEDIM1','a3',1308);
     __setListes('GA_GRILLEDIM2','a3',1311);
     __setListes('GA_GRILLEDIM3','a3',1314);
     __setListes('GA_GRILLEDIM4','a3',1317);
     __setListes('GA_GRILLEDIM5','a3',1320);
     __setListes('GA_PRIXUNIQUE','b1',1323);
     __setListes('GA_STATUTART','a3',1324);
     __setListes('GA_CALCPRIXHT','a3',1327);
     __setListes('GA_CALCPRIXTTC','a3',1330);
     __setListes('GA_COEFCALCHT','n15.2',1333);
     __setListes('GA_COEFCALCTTC','n15.2',1348);
     __setListes('GA_CALCAUTOHT','b1',1363);
     __setListes('GA_CALCAUTOTTC','b1',1364);
     __setListes('GA_FOURNPRINC','a17',1365);
     __setListes('GA_DPRAUTO','b1',1382);
     __setListes('GA_ARRONDIPRIX','a3',1383);
     __setListes('GA_ARRONDIPRIXTTC','a3',1386);
     __setListes('GA_TYPENOMENC','a3',1389);
     __setListes('GA_SUBSTITUTION','a35',1392);
     __setListes('GA_REMPLACEMENT','a35',1427);
     __setListes('GA_TYPEARTFINAN','a3',1462);
     __setListes('GA_INVISIBLEWEB','b1',1465);
     __setListes('GA_DOMAINE','a3',1466);
     __setListes('GA_URLWEB','a35',1469);
     __setListes('GA_CALCPRIXPR','a3',1504);
     __setListes('GA_NATUREPRES','a17',1507);
     __setListes('GA_REGROUPELIGNE','a3',1524);
     __setListes('GA_LISTEREGROUPE','a200',1527);
     __setListes('GA_PRINCIPALEXTRA','a3',1727);
     __setListes('GA_PRIXPASMODIF','b1',1730);
     __setListes('GA_HRSTAT','a3',1731);
     __setListes('GA_QTEDEFAUT','a3',1734);
     __setListes('GA_PRIXPOURQTEAC','n15.2',1737);
     __setListes('GA_REFCONSTRUC','a35',1752);
     __setListes('GA_EAN','a17',1787);
     __setListes('GA_GEREPARC','b1',1804);
     __setListes('GA_FAMILLEVALO','a3',1805);
     __setListes('GA_CALCULPA','b1',1808);
     tabOffset_[140]:=1809;
    end
  else if ident = 'PY' then
    begin
    result := 'PAYS';
    SetLength(tabOffset_, 15);
    __setListes('PY_PAYS', 'a3', 7);
    __setListes('PY_LIBELLE', 'a35', 10);
    __setListes('PY_CODEBANCAIRE', 'a3', 45);
    __setListes('PY_CODEISO2', 'a2', 48);
    __setListes('PY_ABREGE', 'a17', 50);
    __setListes('PY_REGION', 'a1', 67);
    __setListes('PY_NATIONALITE', 'a35', 68);
    __setListes('PY_MEMBRECEE', 'b1', 103);
    __setListes('PY_DEVISE', 'a3', 104);
    __setListes('PY_LANGUE', 'a3', 107);
    __setListes('PY_LIEUDISPO', 'a3', 110);
    __setListes('PY_LIMITROPHE', 'b1', 113);
    __setListes('PY_CODEDI', 'a17', 114);
    __setListes('PY_CODEINSEE', 'a5', 131);
    tabOffset_[14] := 136;
    end
  else if ident = 'ET' then
    begin
    result := 'ETABLISS';
    SetLength(tabOffset_, 37);
    __setListes('ET_ETABLISSEMENT', 'a3', 7);
    __setListes('ET_LIBELLE', 'a35', 10);
    __setListes('ET_ABREGE', 'a17', 45);
    __setListes('ET_JURIDIQUE', 'a3', 62);
    __setListes('ET_ADRESSE1', 'a35', 65);
    __setListes('ET_ADRESSE2', 'a35', 100);
    __setListes('ET_ADRESSE3', 'a35', 135);
    __setListes('ET_CODEPOSTAL', 'a9', 170);
    __setListes('ET_VILLE', 'a35', 179);
    __setListes('ET_PAYS', 'a3', 214);
    __setListes('ET_LANGUE', 'a3', 217);
    __setListes('ET_TELEPHONE', 'a25', 220);
    __setListes('ET_APE', 'a5', 245);
    __setListes('ET_SIRET', 'a17', 250);
    __setListes('ET_LIBREET1', 'a6', 267);
    __setListes('ET_LIBREET2', 'a6', 273);
    __setListes('ET_LIBREET3', 'a6', 279);
    __setListes('ET_LIBREET4', 'a6', 285);
    __setListes('ET_LIBREET5', 'a6', 291);
    __setListes('ET_LIBREET6', 'a6', 297);
    __setListes('ET_LIBREET7', 'a6', 303);
    __setListes('ET_LIBREET8', 'a6', 309);
    __setListes('ET_LIBREET9', 'a6', 315);
    __setListes('ET_LIBREETA', 'a6', 321);
    __setListes('ET_DATELIBRE1', 'd8', 327);
    __setListes('ET_DATELIBRE2', 'd8', 335);
    __setListes('ET_DATELIBRE3', 'd8', 343);
    __setListes('ET_BOOLLIBRE1', 'b1', 351);
    __setListes('ET_BOOLLIBRE2', 'b1', 352);
    __setListes('ET_BOOLLIBRE3', 'b1', 353);
    __setListes('ET_CHARLIBRE1', 'a35', 354);
    __setListes('ET_CHARLIBRE2', 'a35', 389);
    __setListes('ET_CHARLIBRE3', 'a35', 424);
    __setListes('ET_VALLIBRE1', 'n15.2', 459);
    __setListes('ET_VALLIBRE2', 'n15.2', 474);
    __setListes('ET_VALLIBRE3', 'n15.2', 489);
    tabOffset_[36] := 504;
    end
  else if ident ='MR' then
    begin
    result := 'MODEREGL';
    SetLength(tabOffset_, 52);
    __setListes('MR_MODEREGLE','a3', 7);
    __setListes('MR_LIBELLE','a70', 10);
    __setListes('MR_APARTIRDE','a3', 80);
    __setListes('MR_PLUSJOUR','n3', 83);
    __setListes('MR_ARRONDIJOUR','a3', 86);
    __setListes('MR_NOMBREECHEANCE','n2', 89);
    __setListes('MR_SEPAREPAR','a3', 91);
    __setListes('MR_MONTANTMIN','n20', 94);
    __setListes('MR_REMPLACEMIN','a3', 114);
    __setListes('MR_MP1','a3', 117);
    __setListes('MR_TAUX1','n6.2', 120);
    __setListes('MR_MP2','a3', 126);
    __setListes('MR_TAUX2','n6.2', 129);
    __setListes('MR_MP3','a3', 135);
    __setListes('MR_TAUX3','n6.2', 138);
    __setListes('MR_MP4','a3', 144);
    __setListes('MR_TAUX4','n6.2', 147);
    __setListes('MR_MP5','a3', 153);
    __setListes('MR_TAUX5','n6.2', 156);
    __setListes('MR_MP6','a3', 162);
    __setListes('MR_TAUX6','n6.2', 165);
    __setListes('MR_MP7','a3', 171);
    __setListes('MR_TAUX7','n6.2', 174);
    __setListes('MR_MP8','a3', 180);
    __setListes('MR_TAUX8','n6.2', 183);
    __setListes('MR_MP9','a3', 189);
    __setListes('MR_TAUX9','n6.2', 192);
    __setListes('MR_MP10','a3', 198);
    __setListes('MR_TAUX10','n6.2', 201);
    __setListes('MR_MP11','a3', 207);
    __setListes('MR_TAUX11','n6.2', 210);
    __setListes('MR_MP12','a3', 216);
    __setListes('MR_TAUX12','n6.2', 219);
    __setListes('MR_ESC1','b1', 225);
    __setListes('MR_ESC2','b1', 226);
    __setListes('MR_ESC3','b1', 227);
    __setListes('MR_ESC4','b1', 228);
    __setListes('MR_ESC5','b1', 229);
    __setListes('MR_ESC6','b1', 230);
    __setListes('MR_ESC7','b1', 231);
    __setListes('MR_ESC8','b1', 232);
    __setListes('MR_ESC9','b1', 233);
    __setListes('MR_ESC10','b1', 234);
    __setListes('MR_ESC11','b1', 235);
    __setListes('MR_ESC12','b1', 236);
    __setListes('MR_REPARTECHE','a3', 237);
    __setListes('MR_ABREGE','a17', 240);
    __setListes('MR_SOCIETE','a3', 257);
    __setListes('MR_MODEGUIDE','b1', 260);
    __setListes('MR_ECARTJOURS','a50', 261);
    __setListes('MR_EINTEGREAUTO','b1', 311);
    tabOffset_[51] := 312;
    end
  else if ident = 'R' then
    begin
    result := 'RIB';
    SetLength(tabOffset_, 19);
    __setListes('R_AUXILIAIRE','a17',7);
    __setListes('R_NUMERORIB','n6',24);
    __setListes('R_PRINCIPAL','b1',30);
    __setListes('R_ETABBQ','a5',31);
    __setListes('R_GUICHET','a5',36);
    __setListes('R_NUMEROCOMPTE','a11',41);
    __setListes('R_CLERIB','a2',52);
    __setListes('R_DOMICILIATION','a24',54);
    __setListes('R_VILLE','a35',78);
    __setListes('R_PAYS','a3',113);
    __setListes('R_DEVISE','a3',116);
    __setListes('R_CODEBIC','a35',119);
    __setListes('R_SOCIETE','a3',154);
    __setListes('R_SALAIRE','b1',157);
    __setListes('R_ACOMPTE','b1',158);
    __setListes('R_FRAISPROF','b1',159);
    __setListes('R_NATECO','a3',160);
    __setListes('R_CODEIBAN','a70',163);
    tabOffset_[18] := 233;
    end
  else if ident = 'MDP' then
    begin
    result := 'MODEPAIE';
    SetLength(tabOffset_, 33);
    __setListes('MP_MODEPAIE', 'a3', 7);
    __setListes('MP_LIBELLE', 'a35', 10);
    __setListes('MP_CATEGORIE', 'a3', 45);
    __setListes('MP_CODEACCEPT', 'a3', 48);
    __setListes('MP_LETTRECHEQUE', 'b1', 51);
    __setListes('MP_LETTRETRAITE', 'b1', 52);
    __setListes('MP_CONDITION', 'b1', 53);
    __setListes('MP_MONTANTMAX', 'n20.2', 54);
    __setListes('MP_REMPLACEMAX', 'a3', 74);
    __setListes('MP_ABREGE', 'a17', 77);
    __setListes('MP_FORMATCFONB', 'a3', 94);
    __setListes('MP_EDITABLE', 'a70', 97);
    __setListes('MP_POINTABLE', 'b1', 167);
    __setListes('MP_DELAIRETIMPAYE', 'n6', 168);
    __setListes('MP_GENERAL', 'a17', 174);
    __setListes('MP_ENCAISSEMENT', 'a3', 191);
    __setListes('MP_COREXPDEBIT', 'a17', 194);
    __setListes('MP_COREXPCREDIT', 'a17', 211);
    __setListes('MP_CPTEREGLE', 'a17', 228);
    __setListes('MP_JALREGLE', 'a3', 245);
    __setListes('MP_EDITCHEQUEFO', 'b1', 248);
    __setListes('MP_ENVOITPEFO', 'b1', 249);
    __setListes('MP_TYPEMODEPAIE', 'a3', 250);
    __setListes('MP_UTILFO', 'b1', 253);
    __setListes('MP_DEVISEFO', 'a3', 254);
    __setListes('MP_ARRONDIFO', 'a3', 257);
    __setListes('MP_CLIOBLIGFO', 'b1', 260);
    __setListes('MP_MONTANTMIN', 'n15.2', 261);
    __setListes('MP_AVECINFOCOMPL', 'b1', 276);
    __setListes('MP_AVECNUMAUTOR', 'b1', 277);
    __setListes('MP_COPIECBDANSCTRL', 'b1', 278);
    __setListes('MP_AFFICHNUMCBUS', 'b1', 279);
    tabOffset_[32] := 280;
    end
  else if ident = 'ANP' then
    begin
    result := 'ANNUPARAM';
    SetLength(tabOffset_, 14);
    __setListes('ANP_TEXTELIBRE1', 'a35',    7 );
    __setListes('ANP_TEXTELIBRE2', 'a35',    42 );
    __setListes('ANP_TEXTELIBRE3', 'a35',    77 );
    __setListes('ANP_DATELIBRE1', 'a35',     112 );
    __setListes('ANP_DATELIBRE2', 'a35',     147 );
    __setListes('ANP_DATELIBRE3', 'a35',     182 );
    __setListes('ANP_MONTANTLIBRE1', 'a35',  217 );
    __setListes('ANP_MONTANTLIBRE2', 'a35',  252 );
    __setListes('ANP_MONTANTLIBRE3', 'a35',  287 );
    __setListes('ANP_COCHELIBRE1', 'a35',    322 );
    __setListes('ANP_COCHELIBRE2', 'a35',    357 );
    __setListes('ANP_COCHELIBRE3', 'a35',    392 );
    __setListes('ANP_CODE', 'a1',            427 );
    tabOffset_[13] := 428;
    end
  else if ident = 'RPR' then
    begin
    result := 'PROSPECTS';
    SetLength(tabOffset_, 151);
    __setListes('RPR_AUXILIAIRE','a17',7);
    __setListes('RPR_RPRLIBTEXTE0','a35',24);
    __setListes('RPR_RPRLIBTEXTE1','a35',59);
    __setListes('RPR_RPRLIBTEXTE2','a35',94);
    __setListes('RPR_RPRLIBTEXTE3','a35',129);
    __setListes('RPR_RPRLIBTEXTE4','a35',164);
    __setListes('RPR_RPRLIBTEXTE5','a35',199);
    __setListes('RPR_RPRLIBTEXTE6','a35',234);
    __setListes('RPR_RPRLIBTEXTE7','a35',269);
    __setListes('RPR_RPRLIBTEXTE8','a35',304);
    __setListes('RPR_RPRLIBTEXTE9','a35',339);
    __setListes('RPR_RPRLIBVAL0','n20.2',374);
    __setListes('RPR_RPRLIBVAL1','n20.2',394);
    __setListes('RPR_RPRLIBVAL2','n20.2',414);
    __setListes('RPR_RPRLIBVAL3','n20.2',434);
    __setListes('RPR_RPRLIBVAL4','n20.2',454);
    __setListes('RPR_RPRLIBVAL5','n20.2',474);
    __setListes('RPR_RPRLIBVAL6','n20.2',494);
    __setListes('RPR_RPRLIBVAL7','n20.2',514);
    __setListes('RPR_RPRLIBVAL8','n20.2',534);
    __setListes('RPR_RPRLIBVAL9','n20.2',554);
    __setListes('RPR_RPRLIBTABLE0','a3',574);
    __setListes('RPR_RPRLIBTABLE1','a3',577);
    __setListes('RPR_RPRLIBTABLE2','a3',580);
    __setListes('RPR_RPRLIBTABLE3','a3',583);
    __setListes('RPR_RPRLIBTABLE4','a3',586);
    __setListes('RPR_RPRLIBTABLE5','a3',589);
    __setListes('RPR_RPRLIBTABLE6','a3',592);
    __setListes('RPR_RPRLIBTABLE7','a3',595);
    __setListes('RPR_RPRLIBTABLE8','a3',598);
    __setListes('RPR_RPRLIBTABLE9','a3',601);
    __setListes('RPR_RPRLIBTABLE10','a3',604);
    __setListes('RPR_RPRLIBTABLE11','a3',607);
    __setListes('RPR_RPRLIBTABLE12','a3',610);
    __setListes('RPR_RPRLIBTABLE13','a3',613);
    __setListes('RPR_RPRLIBTABLE14','a3',616);
    __setListes('RPR_RPRLIBTABLE15','a3',619);
    __setListes('RPR_RPRLIBTABLE16','a3',622);
    __setListes('RPR_RPRLIBTABLE17','a3',625);
    __setListes('RPR_RPRLIBTABLE18','a3',628);
    __setListes('RPR_RPRLIBTABLE19','a3',631);
    __setListes('RPR_RPRLIBTABLE20','a3',634);
    __setListes('RPR_RPRLIBTABLE21','a3',637);
    __setListes('RPR_RPRLIBTABLE22','a3',640);
    __setListes('RPR_RPRLIBTABLE23','a3',643);
    __setListes('RPR_RPRLIBTABLE24','a3',646);
    __setListes('RPR_RPRLIBTABLE25','a3',649);
    __setListes('RPR_RPRLIBTABLE26','a3',652);
    __setListes('RPR_RPRLIBTABLE27','a3',655);
    __setListes('RPR_RPRLIBTABLE28','a3',658);
    __setListes('RPR_RPRLIBTABLE29','a3',661);
    __setListes('RPR_RPRLIBTABLE30','a3',664);
    __setListes('RPR_RPRLIBTABLE31','a3',667);
    __setListes('RPR_RPRLIBTABLE32','a3',670);
    __setListes('RPR_RPRLIBTABLE33','a3',673);
    __setListes('RPR_RPRLIBTABLE34','a3',676);
    __setListes('RPR_RPRLIBMUL0','a80',679);
    __setListes('RPR_RPRLIBMUL1','a80',759);
    __setListes('RPR_RPRLIBMUL2','a80',839);
    __setListes('RPR_RPRLIBMUL3','a80',919);
    __setListes('RPR_RPRLIBMUL4','a80',999);
    __setListes('RPR_RPRLIBMUL5','a80',1079);
    __setListes('RPR_RPRLIBMUL6','a80',1159);
    __setListes('RPR_RPRLIBMUL7','a80',1239);
    __setListes('RPR_RPRLIBMUL8','a80',1319);
    __setListes('RPR_RPRLIBMUL9','a80',1399);
    __setListes('RPR_RPRLIBBOOL0','b1',1479);
    __setListes('RPR_RPRLIBBOOL1','b1',1480);
    __setListes('RPR_RPRLIBBOOL2','b1',1481);
    __setListes('RPR_RPRLIBBOOL3','b1',1482);
    __setListes('RPR_RPRLIBBOOL4','b1',1483);
    __setListes('RPR_RPRLIBBOOL5','b1',1484);
    __setListes('RPR_RPRLIBBOOL6','b1',1485);
    __setListes('RPR_RPRLIBBOOL7','b1',1486);
    __setListes('RPR_RPRLIBBOOL8','b1',1487);
    __setListes('RPR_RPRLIBBOOL9','b1',1488);
    __setListes('RPR_RPRLIBDATE0','d8',1489);
    __setListes('RPR_RPRLIBDATE1','d8',1497);
    __setListes('RPR_RPRLIBDATE2','d8',1505);
    __setListes('RPR_RPRLIBDATE3','d8',1513);
    __setListes('RPR_RPRLIBDATE4','d8',1521);
    __setListes('RPR_RPRLIBDATE5','d8',1529);
    __setListes('RPR_RPRLIBDATE6','d8',1537);
    __setListes('RPR_RPRLIBDATE7','d8',1545);
    __setListes('RPR_RPRLIBDATE8','d8',1553);
    __setListes('RPR_RPRLIBDATE9','d8',1561);
    __setListes('RPR_DATECREATION','d8',1569);
    __setListes('RPR_DATEMODIF','d8',1577);
    __setListes('RPR_CREATEUR','a3',1585);
    __setListes('RPR_UTILISATEUR','a3',1588);
    __setListes('RPR_RPRLIBVAL10','n19.4',1591);
    __setListes('RPR_RPRLIBVAL11','n19.4',1610);
    __setListes('RPR_RPRLIBVAL12','n19.4',1629);
    __setListes('RPR_RPRLIBVAL13','n19.4',1648);
    __setListes('RPR_RPRLIBVAL14','n19.4',1667);
    __setListes('RPR_RPRLIBVAL15','n19.4',1686);
    __setListes('RPR_RPRLIBVAL16','n19.4',1705);
    __setListes('RPR_RPRLIBVAL17','n19.4',1724);
    __setListes('RPR_RPRLIBVAL18','n19.4',1743);
    __setListes('RPR_RPRLIBVAL19','n19.4',1762);
    __setListes('RPR_RPRLIBDATE10','d8',1781);
    __setListes('RPR_RPRLIBDATE11','d8',1789);
    __setListes('RPR_RPRLIBDATE12','d8',1797);
    __setListes('RPR_RPRLIBDATE13','d8',1805);
    __setListes('RPR_RPRLIBDATE14','d8',1813);
    __setListes('RPR_RPRLIBDATE15','d8',1821);
    __setListes('RPR_RPRLIBDATE16','d8',1829);
    __setListes('RPR_RPRLIBDATE17','d8',1837);
    __setListes('RPR_RPRLIBDATE18','d8',1845);
    __setListes('RPR_RPRLIBDATE19','d8',1853);
    __setListes('RPR_RPRLIBBOOL10','b1',1861);
    __setListes('RPR_RPRLIBBOOL11','b1',1862);
    __setListes('RPR_RPRLIBBOOL12','b1',1863);
    __setListes('RPR_RPRLIBBOOL13','b1',1864);
    __setListes('RPR_RPRLIBBOOL14','b1',1865);
    __setListes('RPR_RPRLIBBOOL15','b1',1866);
    __setListes('RPR_RPRLIBBOOL16','b1',1867);
    __setListes('RPR_RPRLIBBOOL17','b1',1868);
    __setListes('RPR_RPRLIBBOOL18','b1',1869);
    __setListes('RPR_RPRLIBBOOL19','b1',1870);
    __setListes('RPR_RPRLIBMUL10','a80',1871);
    __setListes('RPR_RPRLIBMUL11','a80',1951);
    __setListes('RPR_RPRLIBMUL12','a80',2031);
    __setListes('RPR_RPRLIBMUL13','a80',2111);
    __setListes('RPR_RPRLIBMUL14','a80',2191);
    __setListes('RPR_RPRLIBMUL15','a80',2271);
    __setListes('RPR_RPRLIBMUL16','a80',2351);
    __setListes('RPR_RPRLIBMUL17','a80',2431);
    __setListes('RPR_RPRLIBMUL18','a80',2511);
    __setListes('RPR_RPRLIBMUL19','a80',2591);
    __setListes('RPR_RPRLIBTEXTE10','a35',2671);
    __setListes('RPR_RPRLIBTEXTE11','a35',2706);
    __setListes('RPR_RPRLIBTEXTE12','a35',2741);
    __setListes('RPR_RPRLIBTEXTE13','a35',2776);
    __setListes('RPR_RPRLIBTEXTE14','a35',2811);
    __setListes('RPR_RPRLIBTEXTE15','a35',2846);
    __setListes('RPR_RPRLIBTEXTE16','a35',2881);
    __setListes('RPR_RPRLIBTEXTE17','a35',2916);
    __setListes('RPR_RPRLIBTEXTE18','a35',2951);
    __setListes('RPR_RPRLIBTEXTE19','a35',2986);
    __setListes('RPR_RPRLIBTEXTE20','a35',3021);
    __setListes('RPR_RPRLIBTEXTE21','a35',3056);
    __setListes('RPR_RPRLIBTEXTE22','a35',3091);
    __setListes('RPR_RPRLIBTEXTE23','a35',3126);
    __setListes('RPR_RPRLIBTEXTE24','a35',3161);
    __setListes('RPR_RPRLIBTEXTE25','a35',3196);
    __setListes('RPR_RPRLIBTEXTE26','a35',3231);
    __setListes('RPR_RPRLIBTEXTE27','a35',3266);
    __setListes('RPR_RPRLIBTEXTE28','a35',3301);
    __setListes('RPR_RPRLIBTEXTE29','a35',3336);
    tabOffset_[150] := 3371;
    end
  else if Ident = 'AFO' then
    begin
    Result:='FONCTION';
    SetLength (TabOffset_,14);
    __setListes('AFO_FONCTION','a17',7);
    __setListes('AFO_LIBELLE','a35',24);
    __setListes('AFO_UNITETEMPS','a3',59);
    __setListes('AFO_TAUXREVIENTUN','n15.2',62);
    __setListes('AFO_TARIFRESSOURCE','a3',77);
    __setListes('AFO_PVHT','n15.2',80);
    __setListes('AFO_PVTTC','n15.2',95);
    __setListes('AFO_ORDREETAT','n11',110);
    __setListes('AFO_LIBREFON1','a6',121);
    __setListes('AFO_LIBREFON2','a6',127);
    __setListes('AFO_LIBREFON3','a6',133);
    __setListes('AFO_LIBREFON4','a6',139);
    __setListes('AFO_LIBREFON5','a6',145);
    tabOffset_[13]:=151;
    end
  else if Ident = 'JUR' then
    begin
    Result:='JURIDIQUE';
    SetLength (TabOffset_,18);
    __setListes('JUR_GUIDPERDOS','a36',7);
    __setListes('JUR_TYPEDOS','a3',43);
    __setListes('JUR_NOORDRE','n11',46);
    __setListes('JUR_CODEDOS','a8',57);
    __setListes('JUR_DATECREATION','d8',65);
    __setListes('JUR_DATEMODIF','d8',73);
    __setListes('JUR_UTILISATEUR','a3',81);
    __setListes('JUR_NOMDOS','a25',84);
    __setListes('JUR_DOSLIBELLE','a70',109);
    __setListes('JUR_FORME','a9',179);
    __setListes('JUR_CAPITAL','n19.4',188);
    __setListes('JUR_CAPDEV','a3',207);
    __setListes('JUR_NBDROITSVOTE','n11',210);
    __setListes('JUR_NBTITRESCLOT','n11',221);
    __setListes('JUR_MOISCLOTURE','a3',232);
    __setListes('JUR_VALNOMINCLOT','n19.4',235);
    __setListes('JUR_TYPEACTIV','a3',254);
    tabOffset_[17]:=257;
    end
  else if Ident = 'ROP' then
    begin
    Result:='OPERATIONS';
    SetLength (TabOffset_,43);
    __setListes('ROP_OPERATION','a17',7);
    __setListes('ROP_PRODUITPGI','a3',24);
    __setListes('ROP_LIBELLE','a35',27);
    __setListes('ROP_DATEDEBUT','d8',62);
    __setListes('ROP_DATEFIN','d8',70);
    __setListes('ROP_BUDGET','n20.4',78);
    __setListes('ROP_COUT','n20.4',98);
    __setListes('ROP_CREATEUR','a3',118);
    __setListes('ROP_UTILISATEUR','a3',121);
    __setListes('ROP_DATECREATION','d8',124);
    __setListes('ROP_DATEMODIF','d8',132);
    __setListes('ROP_OBJETOPE','a3',140);
    __setListes('ROP_FERME','b1',143);
    __setListes('ROP_OBJETOPEF','a3',144);
    __setListes('ROP_ROPTABLELIBRE1','a6',147);
    __setListes('ROP_ROPTABLELIBRE2','a6',153);
    __setListes('ROP_ROPTABLELIBRE3','a6',159);
    __setListes('ROP_ROPTABLELIBRE4','a6',165);
    __setListes('ROP_ROPTABLELIBRE5','a6',171);
    __setListes('ROP_ROPTABLELIBRE6','a6',177);
    __setListes('ROP_ROPTABLELIBRE7','a6',183);
    __setListes('ROP_ROPTABLELIBRE8','a6',189);
    __setListes('ROP_ROPTABLELIBRE9','a6',195);
    __setListes('ROP_ROPTABLELIBREA','a6',201);
    __setListes('ROP_ROPVALLIBRE1','n19.4',207);
    __setListes('ROP_ROPVALLIBRE2','n19.4',226);
    __setListes('ROP_ROPVALLIBRE3','n19.4',245);
    __setListes('ROP_ROPDATELIBRE1','d8',264);
    __setListes('ROP_ROPDATELIBRE2','d8',272);
    __setListes('ROP_ROPDATELIBRE3','d8',280);
    __setListes('ROP_ROPCHARLIBRE1','a35',288);
    __setListes('ROP_ROPCHARLIBRE2','a35',323);
    __setListes('ROP_ROPCHARLIBRE3','a35',358);
    __setListes('ROP_ROPBOOLLIBRE1','b1',393);
    __setListes('ROP_ROPBOOLLIBRE2','b1',394);
    __setListes('ROP_ROPBOOLLIBRE3','b1',395);
    __setListes('ROP_FRAISCREATION','n19.4',396);
    __setListes('ROP_FRAISROUTAGE','n19.4',415);
    __setListes('ROP_FRAISSALLE','n19.4',434);
    __setListes('ROP_FRAISMAILING','n19.4',453);
    __setListes('ROP_FRAISCOCKTAIL','n19.4',472);
    __setListes('ROP_FRAISAUTRES','n19.4',491);
    tabOffset_[42]:=510;
    end
  else if Ident = 'RD2' then
    begin
    Result:='RTINFOS002';
    SetLength (TabOffset_,36);
    __setListes('RD2_CLEDATA','a25',7);
    __setListes('RD2_RD2LIBTEXTE0','a35',32);
    __setListes('RD2_RD2LIBTEXTE1','a35',67);
    __setListes('RD2_RD2LIBTEXTE2','a35',102);
    __setListes('RD2_RD2LIBTEXTE3','a35',137);
    __setListes('RD2_RD2LIBTEXTE4','a35',172);
    __setListes('RD2_RD2LIBVAL0','n19.4',207);
    __setListes('RD2_RD2LIBVAL1','n19.4',226);
    __setListes('RD2_RD2LIBVAL2','n19.4',245);
    __setListes('RD2_RD2LIBVAL3','n19.4',264);
    __setListes('RD2_RD2LIBVAL4','n19.4',283);
    __setListes('RD2_RD2LIBTABLE0','a3',302);
    __setListes('RD2_RD2LIBTABLE1','a3',305);
    __setListes('RD2_RD2LIBTABLE2','a3',308);
    __setListes('RD2_RD2LIBTABLE3','a3',311);
    __setListes('RD2_RD2LIBTABLE4','a3',314);
    __setListes('RD2_RD2LIBMUL0','a80',317);
    __setListes('RD2_RD2LIBMUL1','a80',397);
    __setListes('RD2_RD2LIBMUL2','a80',477);
    __setListes('RD2_RD2LIBMUL3','a80',557);
    __setListes('RD2_RD2LIBMUL4','a80',637);
    __setListes('RD2_RD2LIBBOOL0','b1',717);
    __setListes('RD2_RD2LIBBOOL1','b1',718);
    __setListes('RD2_RD2LIBBOOL2','b1',719);
    __setListes('RD2_RD2LIBBOOL3','b1',720);
    __setListes('RD2_RD2LIBBOOL4','b1',721);
    __setListes('RD2_RD2LIBDATE0','d8',722);
    __setListes('RD2_RD2LIBDATE1','d8',730);
    __setListes('RD2_RD2LIBDATE2','d8',738);
    __setListes('RD2_RD2LIBDATE3','d8',746);
    __setListes('RD2_RD2LIBDATE4','d8',754);
    __setListes('RD2_DATECREATION','d8',762);
    __setListes('RD2_DATEMODIF','d8',770);
    __setListes('RD2_CREATEUR','a3',778);
    __setListes('RD2_UTILISATEUR','a3',781);
    tabOffset_[35]:=784;
    end
  else if Ident = 'RPJ' then
    begin
    Result:='PROJETS';
    SetLength (TabOffset_,36);
    __setListes('RPJ_PROJET','a17',7);
    __setListes('RPJ_LIBELLE','a35',24);
    __setListes('RPJ_INTERVENANT','a17',59);
    __setListes('RPJ_NUMEROCONTACT','n11',76);
    __setListes('RPJ_DATEDEBUT','d8',87);
    __setListes('RPJ_ETAT','b1',95);
    __setListes('RPJ_AUXILIAIRE','a17',96);
    __setListes('RPJ_TIERS','a17',113);
    __setListes('RPJ_CREATEUR','a3',130);
    __setListes('RPJ_UTILISATEUR','a3',133);
    __setListes('RPJ_DATECREATION','d8',136);
    __setListes('RPJ_DATEMODIF','d8',144);
    __setListes('RPJ_RPJTABLELIBRE1','a6',152);
    __setListes('RPJ_RPJTABLELIBRE2','a6',158);
    __setListes('RPJ_RPJTABLELIBRE3','a6',164);
    __setListes('RPJ_RPJTABLELIBRE4','a6',170);
    __setListes('RPJ_RPJTABLELIBRE5','a6',176);
    __setListes('RPJ_TYPEPROJET','a3',182);
    __setListes('RPJ_RPJTABLELIBRE6','a6',185);
    __setListes('RPJ_RPJTABLELIBRE7','a6',191);
    __setListes('RPJ_RPJTABLELIBRE8','a6',197);
    __setListes('RPJ_RPJTABLELIBRE9','a6',203);
    __setListes('RPJ_RPJTABLELIBREA','a6',209);
    __setListes('RPJ_RPJVALLIBRE1','n19.4',215);
    __setListes('RPJ_RPJVALLIBRE2','n19.4',234);
    __setListes('RPJ_RPJVALLIBRE3','n19.4',253);
    __setListes('RPJ_RPJDATELIBRE1','d8',272);
    __setListes('RPJ_RPJDATELIBRE2','d8',280);
    __setListes('RPJ_RPJDATELIBRE3','d8',288);
    __setListes('RPJ_RPJCHARLIBRE1','a35',296);
    __setListes('RPJ_RPJCHARLIBRE2','a35',331);
    __setListes('RPJ_RPJCHARLIBRE3','a35',366);
    __setListes('RPJ_RPJBOOLLIBRE1','b1',401);
    __setListes('RPJ_RPJBOOLLIBRE2','b1',402);
    __setListes('RPJ_RPJBOOLLIBRE3','b1',403);
    tabOffset_[35]:=404;
    end
  else if Ident = 'RDQ' then
    begin
    Result:='RTINFOS00Q';
    SetLength (TabOffset_,36);
    __setListes('RDQ_CLEDATA','a25',7);
    __setListes('RDQ_RDQLIBTEXTE0','a35',42);
    __setListes('RDQ_RDQLIBTEXTE1','a35',77);
    __setListes('RDQ_RDQLIBTEXTE2','a35',112);
    __setListes('RDQ_RDQLIBTEXTE3','a35',147);
    __setListes('RDQ_RDQLIBTEXTE4','a35',182);
    __setListes('RDQ_RDQLIBVAL0','n19.4',217);
    __setListes('RDQ_RDQLIBVAL1','n19.4',236);
    __setListes('RDQ_RDQLIBVAL2','n19.4',255);
    __setListes('RDQ_RDQLIBVAL3','n19.4',274);
    __setListes('RDQ_RDQLIBVAL4','n19.4',293);
    __setListes('RDQ_RDQLIBTABLE0','a3',312);
    __setListes('RDQ_RDQLIBTABLE1','a3',315);
    __setListes('RDQ_RDQLIBTABLE2','a3',318);
    __setListes('RDQ_RDQLIBTABLE3','a3',321);
    __setListes('RDQ_RDQLIBTABLE4','a3',324);
    __setListes('RDQ_RDQLIBMUL0','a80',327);
    __setListes('RDQ_RDQLIBMUL1','a80',407);
    __setListes('RDQ_RDQLIBMUL2','a80',487);
    __setListes('RDQ_RDQLIBMUL3','a80',567);
    __setListes('RDQ_RDQLIBMUL4','a80',647);
    __setListes('RDQ_RDQLIBBOOL0','b1',727);
    __setListes('RDQ_RDQLIBBOOL1','b1',728);
    __setListes('RDQ_RDQLIBBOOL2','b1',729);
    __setListes('RDQ_RDQLIBBOOL3','b1',730);
    __setListes('RDQ_RDQLIBBOOL4','b1',731);
    __setListes('RDQ_RDQLIBDATE0','d8',732);
    __setListes('RDQ_RDQLIBDATE1','d8',740);
    __setListes('RDQ_RDQLIBDATE2','d8',748);
    __setListes('RDQ_RDQLIBDATE3','d8',756);
    __setListes('RDQ_RDQLIBDATE4','d8',764);
    __setListes('RDQ_DATECREATION','d8',772);
    __setListes('RDQ_DATEMODIF','d8',780);
    __setListes('RDQ_CREATEUR','a3',788);
    __setListes('RDQ_UTILISATEUR','a3',791);
    tabOffset_[35]:=794;
    end
  else if Ident = 'RPE' then
    begin
    Result:='PERSPECTIVES';
    SetLength (TabOffset_,68);
    __setListes('RPE_PERSPECTIVE','n11',7);
    __setListes('RPE_VARIANTE','n11',18);
    __setListes('RPE_TYPEPERSPECTIV','a3',29);
    __setListes('RPE_LIBELLE','a35',32);
    __setListes('RPE_TIERS','a17',67);
    __setListes('RPE_AUXILIAIRE','a17',84);
    __setListes('RPE_TYPETIERS','a3',101);
    __setListes('RPE_REPRESENTANT','a17',104);
    __setListes('RPE_INTERVENANT','a17',121);
    //__setListes('RPE_INTERVINT','a80',138);
    __setListes('RPE_OPERATION','a17',138);
    __setListes('RPE_NUMEROACTION','n11',155);
    __setListes('RPE_PROJET','a17',166);
    __setListes('RPE_AFFAIRE','a17',183);
    __setListes('RPE_DATEREALISE','d8',200);
    __setListes('RPE_ETATPER','a3',208);
    __setListes('RPE_MONTANTPER','n20.4',211);
    __setListes('RPE_POURCENTAGE','n15.9',231);
    //__setListes('RPE_PIECE','a35',326);
    __setListes('RPE_TABLELIBREPER1','a3',246);
    __setListes('RPE_TABLELIBREPER2','a3',249);
    __setListes('RPE_TABLELIBREPER3','a3',252);
    __setListes('RPE_DATECREATION','d8',255);
    __setListes('RPE_CREATEUR','a3',263);
    __setListes('RPE_DATEMODIF','d8',266);
    __setListes('RPE_UTILISATEUR','a3',274);
    __setListes('RPE_MOTIFPERTE1','a3',277);
    __setListes('RPE_MOTIFPERTE2','a3',280);
    __setListes('RPE_NUMEROCONTACT','n11',283);
    __setListes('RPE_DESTMAIL','a80',294);
    __setListes('RPE_CONCURRENTS','a85',374);
    __setListes('RPE_DATEFINVIE','d8',459);
    __setListes('RPE_MOTIFSIGNATURE','a3',467);
    __setListes('RPE_COMMENTPERTE','a200',470);
    __setListes('RPE_RPETABLELIBRE1','a6',670);
    __setListes('RPE_RPETABLELIBRE2','a6',676);
    __setListes('RPE_RPETABLELIBRE3','a6',682);
    __setListes('RPE_RPETABLELIBRE4','a6',688);
    __setListes('RPE_RPETABLELIBRE5','a6',694);
    __setListes('RPE_APPORTEUR1','a17',700);
    __setListes('RPE_MOTIFRETRIB1','a35',717);
    __setListes('RPE_MONTANTRETRIB1','n20.4',752);
    __setListes('RPE_APPEXTERNE1','b1',772);
    __setListes('RPE_APPORTEUR2','a17',773);
    __setListes('RPE_MOTIFRETRIB2','a35',790);
    __setListes('RPE_MONTANTRETRIB2','n20.4',825);
    __setListes('RPE_APPEXTERNE2','b1',845);
    __setListes('RPE_APPORTEUR3','a17',846);
    __setListes('RPE_MOTIFRETRIB3','a35',863);
    __setListes('RPE_MONTANTRETRIB3','n20.4',898);
    __setListes('RPE_APPEXTERNE3','b1',918);
    __setListes('RPE_TYPEPREVIMISS','a14',919);
    __setListes('RPE_RPETABLELIBRE6','a6',933);
    __setListes('RPE_RPETABLELIBRE7','a6',939);
    __setListes('RPE_RPETABLELIBRE8','a6',945);
    __setListes('RPE_RPETABLELIBRE9','a6',951);
    __setListes('RPE_RPETABLELIBREA','a6',957);
    __setListes('RPE_RPEVALLIBRE1','n19.4',963);
    __setListes('RPE_RPEVALLIBRE2','n19.4',982);
    __setListes('RPE_RPEVALLIBRE3','n19.4',1001);
    __setListes('RPE_RPEDATELIBRE1','d8',1020);
    __setListes('RPE_RPEDATELIBRE2','d8',1028);
    __setListes('RPE_RPEDATELIBRE3','d8',1036);
    __setListes('RPE_RPECHARLIBRE1','a35',1044);
    __setListes('RPE_RPECHARLIBRE2','a35',1079);
    __setListes('RPE_RPECHARLIBRE3','a35',1114);
    __setListes('RPE_RPEBOOLLIBRE1','b1',1149);
    __setListes('RPE_RPEBOOLLIBRE2','b1',1150);
    __setListes('RPE_RPEBOOLLIBRE3','b1',1151);
    tabOffset_[67]:=1152;
    end
  else if Ident = 'RDV' then
    begin
    Result:='RTINFOS00V';
    SetLength (TabOffset_,36);
    __setListes('RDV_CLEDATA','a25',7);
    __setListes('RDV_RDVLIBTEXTE0','a35',42);
    __setListes('RDV_RDVLIBTEXTE1','a35',77);
    __setListes('RDV_RDVLIBTEXTE2','a35',112);
    __setListes('RDV_RDVLIBTEXTE3','a35',147);
    __setListes('RDV_RDVLIBTEXTE4','a35',182);
    __setListes('RDV_RDVLIBVAL0','n19.4',217);
    __setListes('RDV_RDVLIBVAL1','n19.4',236);
    __setListes('RDV_RDVLIBVAL2','n19.4',255);
    __setListes('RDV_RDVLIBVAL3','n19.4',274);
    __setListes('RDV_RDVLIBVAL4','n19.4',293);
    __setListes('RDV_RDVLIBTABLE0','a3',312);
    __setListes('RDV_RDVLIBTABLE1','a3',315);
    __setListes('RDV_RDVLIBTABLE2','a3',318);
    __setListes('RDV_RDVLIBTABLE3','a3',321);
    __setListes('RDV_RDVLIBTABLE4','a3',324);
    __setListes('RDV_RDVLIBMUL0','a80',327);
    __setListes('RDV_RDVLIBMUL1','a80',407);
    __setListes('RDV_RDVLIBMUL2','a80',487);
    __setListes('RDV_RDVLIBMUL3','a80',567);
    __setListes('RDV_RDVLIBMUL4','a80',647);
    __setListes('RDV_RDVLIBBOOL0','b1',727);
    __setListes('RDV_RDVLIBBOOL1','b1',728);
    __setListes('RDV_RDVLIBBOOL2','b1',729);
    __setListes('RDV_RDVLIBBOOL3','b1',730);
    __setListes('RDV_RDVLIBBOOL4','b1',731);
    __setListes('RDV_RDVLIBDATE0','d8',732);
    __setListes('RDV_RDVLIBDATE1','d8',740);
    __setListes('RDV_RDVLIBDATE2','d8',748);
    __setListes('RDV_RDVLIBDATE3','d8',756);
    __setListes('RDV_RDVLIBDATE4','d8',764);
    __setListes('RDV_DATECREATION','d8',772);
    __setListes('RDV_DATEMODIF','d8',780);
    __setListes('RDV_CREATEUR','a3',788);
    __setListes('RDV_UTILISATEUR','a3',791);
    tabOffset_[35]:=794;
    end
  else if Ident = 'RPT' then
    begin
    Result:='PERSPECTIVESTIERS';
    SetLength (TabOffset_,3);
    __setListes('RPT_PERSPECTIVE','n11',7);
    __setListes('RPT_TIERS','a17',18);
    tabOffset_[2]:=35;
    end
  else if Ident = 'RAC' then
    begin
    Result:='ACTIONS';
    SetLength (TabOffset_,84);
    __setListes('RAC_NUMACTION','n11',7);
    __setListes('RAC_LIBELLE','a35',18);
    __setListes('RAC_AUXILIAIRE','a17',53);
    __setListes('RAC_TIERS','a17',70);
    __setListes('RAC_TYPEACTION','a3',87);
    __setListes('RAC_PRODUITPGI','a3',90);
    __setListes('RAC_INTERVENANT','a17',93);
    __setListes('RAC_INTERVINT','a85',110);
    __setListes('RAC_OPERATION','a17',195);
    __setListes('RAC_NUMCHAINAGE','n11',212);
    __setListes('RAC_NUMACTGEN','n11',223);
    __setListes('RAC_PROJET','a17',234);
    __setListes('RAC_DATEACTION','d8',251);
    __setListes('RAC_HEUREACTION','d8',259);
    __setListes('RAC_DUREEACT','d8',267);
    __setListes('RAC_DUREEACTION','n19.4',275);
    __setListes('RAC_DATEECHEANCE','d8',294);
    __setListes('RAC_ETATACTION','a3',302);
    __setListes('RAC_COUTACTION','n20.4',305);
    __setListes('RAC_UTILISATEUR','a3',325);
    __setListes('RAC_CREATEUR','a3',328);
    __setListes('RAC_CHRONOMETRE','n20.4',331);
    __setListes('RAC_TABLELIBRE1','a3',351);
    __setListes('RAC_TABLELIBRE2','a3',354);
    __setListes('RAC_TABLELIBRE3','a3',357);
    __setListes('RAC_DATECREATION','d8',360);
    __setListes('RAC_DATEMODIF','d8',368);
    __setListes('RAC_NUMEROCONTACT','n11',376);
    __setListes('RAC_DESTMAIL','a80',387);
    __setListes('RAC_NIVIMP','a3',467);
    __setListes('RAC_PERSPECTIVE','n11',470);
    __setListes('RAC_GESTRAPPEL','b1',481);
    __setListes('RAC_DELAIRAPPEL','a3',482);
    __setListes('RAC_DATERAPPEL','d8',485);
    __setListes('RAC_CHAINAGE','a17',493);
    __setListes('RAC_NUMLIGNE','n11',510);
    __setListes('RAC_MAILENVOYE','b1',521);
    __setListes('RAC_DERNIERMAIL','d8',522);
    __setListes('RAC_MAILAUTO','b1',530);
    __setListes('RAC_AFFAIRE','a17',531);
    __setListes('RAC_AFFAIRE0','b1',548);
    __setListes('RAC_AFFAIRE1','a14',549);
    __setListes('RAC_AFFAIRE2','a13',563);
    __setListes('RAC_AFFAIRE3','a12',576);
    __setListes('RAC_AVENANT','a2',588);
    __setListes('RAC_TABLELIBREF1','a3',590);
    __setListes('RAC_TABLELIBREF2','a3',593);
    __setListes('RAC_TABLELIBREF3','a3',596);
    __setListes('RAC_IDENTPARC','n11',599);
    __setListes('RAC_IDENTIFIANT','n11',610);
    __setListes('RAC_QORIGINERQ','a3',621);
    __setListes('RAC_QUALITETYPE','a3',624);
    __setListes('RAC_TYPELOCA','a3',627);
    __setListes('RAC_QDEMDEROGNUM','n11',630);
    __setListes('RAC_QPLANCORRNUM','n11',641);
    __setListes('RAC_QNCNUM','n11',652);
    __setListes('RAC_NATURETRAVAIL','a3',663);
    __setListes('RAC_LIGNEORDRE','n11',666);
    //__setListes('RAC_AREALISERPAR','a17',677);
    __setListes('RAC_QPCTAVANCT','n11',677);
    //__setListes('RAC_REALISEELE','d8',705);
    //__setListes('RAC_REALISEEPAR','a17',713);
    //__setListes('RAC_VERIFIEELE','d8',730);
    //__setListes('RAC_VERIFIEEPAR','a17',738);
    __setListes('RAC_EFFICACITE','a3',688);
    //__setListes('RAC_EFFJUGEELE','d8',758);
    //__setListes('RAC_EFFJUGEEPAR','a17',766);
    //__setListes('RAC_CLOTUREELE','d8',783);
    //__setListes('RAC_CLOTUREEPAR','a17',791);
    __setListes('RAC_PRIVATE','b1',691);
    __setListes('RAC_RACTABLELIBRE1','a6',692);
    __setListes('RAC_RACTABLELIBRE2','a6',698);
    __setListes('RAC_RACTABLELIBRE3','a6',704);
    __setListes('RAC_RACTABLELIBRE4','a6',710);
    __setListes('RAC_RACTABLELIBRE5','a6',716);
    __setListes('RAC_RACTABLELIBRE6','a6',722);
    __setListes('RAC_RACTABLELIBRE7','a6',728);
    __setListes('RAC_RACTABLELIBRE8','a6',734);
    __setListes('RAC_RACTABLELIBRE9','a6',740);
    __setListes('RAC_RACTABLELIBREA','a6',746);
    __setListes('RAC_RACVALLIBRE1','n19.4',752);
    __setListes('RAC_RACVALLIBRE2','n19.4',771);
    __setListes('RAC_RACVALLIBRE3','n19.4',790);
    __setListes('RAC_RACDATELIBRE1','d8',809);
    __setListes('RAC_RACDATELIBRE2','d8',817);
    __setListes('RAC_RACDATELIBRE3','d8',825);
    __setListes('RAC_RACCHARLIBRE1','a35',833);
    __setListes('RAC_RACCHARLIBRE2','a35',868);
    __setListes('RAC_RACCHARLIBRE3','a35',903);
    __setListes('RAC_RACBOOLLIBRE1','b1',938);
    __setListes('RAC_RACBOOLLIBRE2','b1',939);
    __setListes('RAC_RACBOOLLIBRE3','b1',940);
    tabOffset_[83]:=941;
    end
  else if Ident = 'RD1' then
    begin
    Result:='RTINFOS001';
    SetLength (TabOffset_,36);
    __setListes('RD1_CLEDATA','a25',7);
    __setListes('RD1_RD1LIBTEXTE0','a35',32);
    __setListes('RD1_RD1LIBTEXTE1','a35',67);
    __setListes('RD1_RD1LIBTEXTE2','a35',102);
    __setListes('RD1_RD1LIBTEXTE3','a35',137);
    __setListes('RD1_RD1LIBTEXTE4','a35',172);
    __setListes('RD1_RD1LIBVAL0','n19.4',207);
    __setListes('RD1_RD1LIBVAL1','n19.4',226);
    __setListes('RD1_RD1LIBVAL2','n19.4',245);
    __setListes('RD1_RD1LIBVAL3','n19.4',264);
    __setListes('RD1_RD1LIBVAL4','n19.4',283);
    __setListes('RD1_RD1LIBTABLE0','a3',302);
    __setListes('RD1_RD1LIBTABLE1','a3',305);
    __setListes('RD1_RD1LIBTABLE2','a3',308);
    __setListes('RD1_RD1LIBTABLE3','a3',311);
    __setListes('RD1_RD1LIBTABLE4','a3',314);
    __setListes('RD1_RD1LIBMUL0','a80',317);
    __setListes('RD1_RD1LIBMUL1','a80',397);
    __setListes('RD1_RD1LIBMUL2','a80',477);
    __setListes('RD1_RD1LIBMUL3','a80',557);
    __setListes('RD1_RD1LIBMUL4','a80',637);
    __setListes('RD1_RD1LIBBOOL0','b1',717);
    __setListes('RD1_RD1LIBBOOL1','b1',718);
    __setListes('RD1_RD1LIBBOOL2','b1',719);
    __setListes('RD1_RD1LIBBOOL3','b1',720);
    __setListes('RD1_RD1LIBBOOL4','b1',721);
    __setListes('RD1_RD1LIBDATE0','d8',722);
    __setListes('RD1_RD1LIBDATE1','d8',730);
    __setListes('RD1_RD1LIBDATE2','d8',738);
    __setListes('RD1_RD1LIBDATE3','d8',746);
    __setListes('RD1_RD1LIBDATE4','d8',754);
    __setListes('RD1_DATECREATION','d8',762);
    __setListes('RD1_DATEMODIF','d8',770);
    __setListes('RD1_CREATEUR','a3',778);
    __setListes('RD1_UTILISATEUR','a3',781);
    tabOffset_[35]:=784;
    end
  else if Ident = 'RD6' then
    begin
    Result:='RTINFOS006';
    SetLength (TabOffset_,36);
    __setListes('RD6_CLEDATA','a25',7);
    __setListes('RD6_RD6LIBTEXTE0','a35',32);
    __setListes('RD6_RD6LIBTEXTE1','a35',67);
    __setListes('RD6_RD6LIBTEXTE2','a35',102);
    __setListes('RD6_RD6LIBTEXTE3','a35',137);
    __setListes('RD6_RD6LIBTEXTE4','a35',172);
    __setListes('RD6_RD6LIBVAL0','n19.4',207);
    __setListes('RD6_RD6LIBVAL1','n19.4',226);
    __setListes('RD6_RD6LIBVAL2','n19.4',245);
    __setListes('RD6_RD6LIBVAL3','n19.4',264);
    __setListes('RD6_RD6LIBVAL4','n19.4',283);
    __setListes('RD6_RD6LIBTABLE0','a3',302);
    __setListes('RD6_RD6LIBTABLE1','a3',305);
    __setListes('RD6_RD6LIBTABLE2','a3',308);
    __setListes('RD6_RD6LIBTABLE3','a3',311);
    __setListes('RD6_RD6LIBTABLE4','a3',314);
    __setListes('RD6_RD6LIBMUL0','a80',317);
    __setListes('RD6_RD6LIBMUL1','a80',397);
    __setListes('RD6_RD6LIBMUL2','a80',477);
    __setListes('RD6_RD6LIBMUL3','a80',557);
    __setListes('RD6_RD6LIBMUL4','a80',637);
    __setListes('RD6_RD6LIBBOOL0','b1',717);
    __setListes('RD6_RD6LIBBOOL1','b1',718);
    __setListes('RD6_RD6LIBBOOL2','b1',719);
    __setListes('RD6_RD6LIBBOOL3','b1',720);
    __setListes('RD6_RD6LIBBOOL4','b1',721);
    __setListes('RD6_RD6LIBDATE0','d8',722);
    __setListes('RD6_RD6LIBDATE1','d8',730);
    __setListes('RD6_RD6LIBDATE2','d8',738);
    __setListes('RD6_RD6LIBDATE3','d8',746);
    __setListes('RD6_RD6LIBDATE4','d8',754);
    __setListes('RD6_DATECREATION','d8',762);
    __setListes('RD6_DATEMODIF','d8',770);
    __setListes('RD6_CREATEUR','a3',778);
    __setListes('RD6_UTILISATEUR','a3',781);
    tabOffset_[35]:=784;
    end
  else if Ident = 'RAI' then
    begin
    Result:='ACTIONINTERVENANT';
    SetLength (TabOffset_,5);
    __setListes('RAI_AUXILIAIRE','a17',7);
    __setListes('RAI_NUMACTION','n11',24);
    __setListes('RAI_RESSOURCE','a17',35);
    __setListes('RAI_GUID','a36',52);
    tabOffset_[4]:=88;
    end
  else if Ident = 'GEDD' then
    begin
    Result:='GEDDP';
    SetLength (TabOffset_,11);
    __setListes('FICHIER','a255',8);
    __setListes('NOM','a70',263);
    __setListes('EXTENSION','a5',333);
    __setListes('CODEGED','a6',338);
    __setListes('CODEDOSSIER','a8',344);
    __setListes('CODETIERS','a8',352);
    __setListes('LIBELLE','a70',360);
    __setListes('AUTEUR','a35',430);
    __setListes('ANNEE','a4',465);
    __setListes('MOIS','a2',469);
    tabOffset_[10]:=471;
    end
  else if Ident = 'GEDS' then
    begin
    Result:='GEDSTD';
    SetLength (TabOffset_,19);
    __setListes('FICHIER','a255',8);
    __setListes('CODEPRODUIT','a17',263);
    __setListes('NOM','a70',280);
    __setListes('EXTENSION','a5',350);
    __setListes('CODEDOSSIER','a8',355);
    __setListes('LIBELLE','a35',363);
    __setListes('PREDEFINI','a3',398);
    __setListes('LANGUE','a17',401);
    __setListes('CRIT1','a17',404);
    __setListes('CRIT2','a17',421);
    __setListes('CRIT3','a17',438);
    __setListes('CRIT4','a17',455);
    __setListes('CRIT5','a17',472);
    __setListes('BCRIT1','b1',489);
    __setListes('BCRIT2','b1',490);
    __setListes('BCRIT3','b1',491);
    __setListes('BCRIT4','b1',492);
    __setListes('BCRIT5','b1',493);
    tabOffset_[18]:=494;
    end
  else if ident = 'BCNC' then
    begin
    Result:='CONTACT';
    SetLength (TabOffset_,5);
    __setListes('C_TYPECONTACT','a3',10);
    __setListes('C_AUXILIAIRE','a36',13);
    __setListes('C_NUMEROCONTACT','n11',49);
    __setListes('C_BLOCNOTE','t1500',60);
    tabOffset_[4]:=1560;
    end
  else if ident = 'BCNT' then
    begin
    Result:='TIERS';
    SetLength (TabOffset_,3);
    __setListes('T_AUXILIAIRE','a17',10);
    __setListes('T_BLOCNOTE','t1500',27);
    tabOffset_[2]:=1527;
    end
  else if ident = 'BCNRAC' then
    begin
    Result:='ACTIONS';
    SetLength (TabOffset_,4);
    __setListes('RAC_NUMACTION','n11',10);
    __setListes('RAC_AUXILIAIRE','a17',21);
    __setListes('RAC_BLOCNOTE','t1500',38);
    tabOffset_[3]:=1538;
    end
  else if ident = 'BCNRPE' then
    begin
    Result:='PERSPECTIVES';
    SetLength (TabOffset_,3);
    __setListes('RPE_PERSPECTIVE','n11',10);
    __setListes('RPE_BLOCNOTE','t1500',21);
    tabOffset_[2]:=1521;
    end
  else if ident = 'BCNROP' then
    begin
    Result:='OPERATIONS';
    SetLength (TabOffset_,3);
    __setListes('ROP_OPERATION','a17',10);
    __setListes('ROP_BLOCNOTE','t1500',27);
    tabOffset_[2]:=1527;
    end
  else if ident = 'BCNRD2' then
    begin
    Result:='RTINFOS002';
    SetLength (TabOffset_,3);
    __setListes('RD2_CLEDATA','a25',10);
    __setListes('RD2_BLOCNOTE','t1500',35);
    tabOffset_[2]:=1535;
    end

  // 3OB le 06/02/2008 : ajout pour gestion des tables de segmentation
  // RP0 13/10/08 : CHOIXEXT est déjà définie plus haut
  {
  else if ident = 'YX' then // Etiquette du modèle , liste des questionnaires, questions et reponses associées
    begin
      result := 'CHOIXEXT';
      SetLength(tabOffset_, 6);
      __setListes('YX_TYPE', 'a3', 7);
      __setListes('YX_CODE', 'a17', 10);
      __setListes('YX_LIBELLE', 'a105', 27);
      __setListes('YX_ABREGE', 'a17', 132);
      __setListes('YX_LIBRE', 'a35', 149);
      tabOffset_[5] := 184;
    end
  }
  else if ident = 'ASG' then // Détail du modèle (questionnaires associés)
    begin
      result := 'AFSEGMODELEDET';
      SetLength(tabOffset_, 4); //dimensionne le tableau
      __setListes('ASG_MODELESEG', 'a17', 7);
      __setListes('ASG_QUESTIONN', 'a17', 24);
      __setListes('ASG_NUMORDRE', 'n6', 41);
      tabOffset_[3] := 47;
    end
  else if ident = 'ASD' then // Détail des questionnaires (questions associées)
    begin
      result := 'AFSEGQUESTIONDET';
      setLength(tabOffset_, 5);
      __setListes('ASD_QUESTIONN', 'a17', 7);
      __setListes('ASD_QUESTSEG', 'a17', 24);
      __setListes('ASD_POIDSQ', 'n8.2', 41);
      __setListes('ASD_VALEURREP', 'n6', 49);
      tabOffset_[4] := 55;
    end
  else if ident = 'ASR' then // Réponses possibles aux questions
    begin
      result := 'AFSEGREPONSE';
      setLength(tabOffset_, 5);
      __setListes('ASR_QUESTSEG', 'a17', 7);
      __setListes('ASR_CODEREPS', 'a5', 24);
      __setListes('ASR_TEXTEREP', 'a70', 29);
      __setListes('ASR_VALEURREP', 'n6', 99);
      tabOffset_[4] := 105;
    end
  else if ((ident = 'ASE') or (ident = 'ASE_1') or (ident = 'ASE_2'))  then // Explications sur segmentation
    begin
      result := 'AFSEGEXPRES'+copy(ident,4,2);
      setLength(tabOffset_, 6);
      __setListes('ASE_CODE', 'a17', 7);
      __setListes('ASE_LIBELLE', 'a105', 24);
      __setListes('ASE_TYPE', 'a5', 129);
      __setListes('ASE_RESULTAT', 'a17', 134);
      __setListes('ASE_EXPLICATION', 't1500', 151);
      tabOffset_[5] := 1651;
    end
  else if ident = 'ASB' then // Combinaisons des resultats de segmentation
    begin
      result := 'AFSEGCOMBRES';
      setLength(tabOffset_, 6);
      __setListes('ASB_CODEQ1', 'a17', 7);
      __setListes('ASB_RESULT1', 'a17', 24);
      __setListes('ASB_CODEQ2', 'a17', 41);
      __setListes('ASB_RESULT2', 'a17', 58);
      __setListes('ASB_MEMOCOMB', 't1500', 75);
      tabOffset_[5] := 1575;
    end
    // 3OB : Fin
  else
    result := '';//SI IDENTIFIANT NON-RECONNU
  except
    On EOutOfMemory Do result := '';
  end;
end;
{-----------------------------------------------------------------------------}
{enlève les espaces situées avant et après la chaîne}
function TImportDP.__ajusterTaille(s :string) : string;
var
  tmp : string;
  i :integer;
begin
  i := 0;
  while i< length(s) do//enlève les espaces à gauche
    begin
    inc(i);
    if s[i]<> ' ' then break;
    end;
  tmp := copy(s, i, length(s)-i+1);
  i := length(tmp)+1;
  while i>1 do//enlève les espaces à droite
    begin
    dec(i);
    if tmp[i]<>' ' then break;
    end;
  tmp := copy(tmp,1, i);
  if tmp = ' ' then result := ''
  else result := tmp;
end;
{-----------------------------------------------------------------------------}
{ajoute le nom du champ, son type et son offset dans la liste correspondante}
procedure TImportDP.__setListes(nom, typ : string; offset : integer);
begin
  ListeChps_.Add(nom);
  ListeTypes_.Add(typ);
  TabOffset_[ListeChps_.Count-1] := offset;
end;
{-----------------------------------------------------------------------------}
constructor TImportDP.Create;
begin
  Setlength(Tobs_, 0);
  //creation des listes
  listeFic_ := TStringList.Create;
  listeIdent_ := TStringList.Create;
  listeFic_.Clear;
  listeIdent_.Clear;
//  SetLength(Tobs, 26);//allocation du tableau. -> est faite plus tard
end;
{-----------------------------------------------------------------------------}
destructor TImportDP.Free;
var i : integer;
begin
  for i := 0 to listeIdent_.Count-1 do
    Tobs_[i].Free;//libère les tobs
  //libère les tableaux dynamiques
  Finalize(Tobs_);
  //libération des listes
  listeFic_.Free;
  listeIdent_.Free;
end;
{-----------------------------------------------------------------------------}

{******************************************************************************
 *                 FONCTIONS D'EXPORT POUR TESTS                              *
 ******************************************************************************}

{Ajoute au fichier la ligne d'en-tête}
procedure TImportDP.__addEntete(var f : textFile);
var
  s : string;
begin
  s := '***ETT'+version;//numéro de version du format TRADP utilisé
  s := s+correctLength('BUREAU PGI',35);//nom du prog générant le fichier;
  s := s+FormatDateTime('ddmmyyyyhhmm',now);//date et heure de génération du fichier
  s := s+correctLength(V_PGI.UserName, 35);//utilisateur générant le fichier
  s := s+correctLength(GetParamSoc('SO_LIBELLE'), 35);//raison sociale
  writeln(f, s);
end;
{-----------------------------------------------------------------------------}
 {Ajuste la chaîne à la longueur désirée}
function TImportDP.correctlength(s :string; l : integer) : string;
begin
  if length(s) > l then//chaîne trop longue : on coupe
    result := copy(s,1 , l)
  else if length(s) < l then//chaîne trop courte : on bourre à droite
    begin
    if V_PGI.SAV And (s = '') then
      result := s+'@'+format_string(' ', l-length(s)-1)//version de debugage : fichier inexploitable !
    else
      result := s+format_string(' ', l-length(s))//version normale;
    end
  else if length(s) = l then//bonne longueur
    result := s;
end;
{------------------------------------------------------------------------------}
{Récupère une donnée et la convertit en string}
function TImportDP.__varToStr(rsql : tquery; idx : integer) : string;
var
  l,d : integer;
  t,s : string;
begin
  t := copy(listeTypes_[idx], 1,1);
  l := tabOffset_[idx+1] - tabOffset_[idx];//longueur de la chaine
//  if
  if (t = 'a') Or (t = 'b') then//chaîne ou booléen
    begin
    s := rsql.FindField(listeChps_[idx]).AsString;
    s := FindEtReplace(s, Chr(9), ' ', TRUE);
    s := FindEtReplace(s, Chr(13)+Chr(10), ' ', TRUE);
    result := correctLength(s, l);
    end
  else if t = 't' then//text limité à 1500 caractères
    begin
    s := rsql.FindField(listeChps_[idx]).AsString;
    //s := FindEtReplace(s, Chr(9), Chr(2), TRUE);
    s := FindEtReplace(s, Chr(13)+Chr(10), Chr(3), TRUE);
    s := FindEtReplace(s, Chr(0), Chr(4), TRUE);
    result := correctLength(s, l);
    end
  else if t = 'd' then//date
    result := correctLength(DateToStr(rsql.findField(listeChps_[idx]).AsDateTime), 8)
  else if t = 'n' then//numérique
    begin
    if pos('.', listeTypes_[idx]) = 0 then //entier
      result := alignDroite(StrfMontant(rsql.findField(listeChps_[idx]).AsFloat,l,0,'',false), l)
    else//float
      begin
      d := strtoint(copy(listetypes_[idx],pos('.',listetypes_[idx])+1,2));//nombre de décimales
      result := alignDroite(StrfMontant(rsql.findField(listeChps_[idx]).AsFloat,l ,d,'',false), l);
      end;
    end;//IF T='n'
end;
{------------------------------------------------------------------------------}
{écrit une ligne du TRA à partir d'un enregistrement}
function TImportDP.writeligne( rsql : tquery; var f : textfile; nomTable,identTable : string) : integer;
var
  s  : string;
  formatDate : string;
  i,c : integer;
  tobtest : TOB;
begin
  result := 0;
  s := '';
  formatDate := ShortDateFormat;//sauvegarde du format de date courant
  ShortDateFormat := 'ddmmyyyy';
  tobtest := toB.Create(nomTable,nil, -1);//creation d'une tob pour tester si les champs sont réels
  if nomTable = 'DOSSIERGRP' then//spécif tables groupes de travail
    begin
    tobtest.AddChampSup('DOG_NODOSSIER', True);
    tobtest.AddChampSup('DOG_GROUPECONF', True);
    end;
  s := '***'+correctLength(identTable,3);//ajout de l'identifiant du champ du TRA
  c:= High(tabOffset_)-1;
  for i := 0 to c do
    begin
    if tobtest.FieldExists(listechps_[i]) then//vérification que le champ existe
      begin
      try
      s := s+__varToStr(rsql, i)//si oui, on l'extrait de la réponse à la requête
      except
      pgiinfo('Erreur de conversion de la donnée '+listeChps_[i]);
      end;
      end
    else
      s := s+format_string(' ', tabOffset_[i+1] - tabOffset_[i]);//si non, on bourre.
    end;
  tobtest.Free;
  writeln(f, s);
  Flush(f);
  ShortDateFormat := formatDate;//restauration du format de date courant
end;
{------------------------------------------------------------------------------}
{Génère un fichier TRA contenant la table voulue}
function TImportDP.generateFile(identTable : array of string; nomFic : string) : integer;
var
  Rsql : TQuery;
  f : TextFile;
  i : integer;
  nomTable : string;
  critereSup : string;
  SelectFrom : String;  // 3OB le 03/04/08 pour ajout select distinct
begin
  assignFile(f, nomFic);
  rewrite(f);//ouverture en écriture, écrasement si existe déjà
{$IFNDEF GIGI}
  __addEntete(f);//écrit l'entête (infos générales)
{$ELSE}
  writeln(f, identTable[7]);   // 3OB le 17/03/2008 : écriture du code modèle en première ligne
{$ENDIF}
  InitMoveProgressForm(nil,'Execution Export', 'Traitement en cours', high(identTable), True, True);
  listeChps_ := TStringList.Create;
  listeTypes_ := TStringList.Create;
  For i := 0 to high(identTable) do
  begin
  try
    Listechps_.Clear;
    ListeTypes_.Clear;
    if (identTable[i] = 'GEDD') or (identTable[i] = 'GEDS') then continue; //Formats pour GED valables en import uniquement
    nomTable := __initExtract(identTable[i]);//initalise les listes
    SelectFrom:='SELECT * FROM ';  // 3OB le 03/4/08
    //selection des tablettes du DP ET YY  parmis celles de choixcod et choixext
    if nomTable = 'CHOIXCOD' then
     begin
      // OLD critereSup := ' WHERE (CC_TYPE IN ("DCR","DLT","DMS","DMT","DNT","DOC","DRA","DRI","DRT","DST","DTB","DTM","DWI","AN1","AN2","AN3","DCU","GDM","SRV","UCO","VOT","YCT","YDA","YEW","YNA","YPP","YTH")) OR (CC_TYPE="CIV" AND CC_CODE<>"DR")'
      CritereSup:=' WHERE CC_TYPE IN ("DRA","DMS","DOC","DCR","DWI","DMT","DNT","DRI","DRT","DTB","DTM","AN1","AN2","AN3","YEW","VOT","LGU","GZC","FN1","FN2","FN3","TX1","TX2","TX3","TX4"';
      CritereSup:=CritereSup+',"TX5","GCT","GCA","JUR","GCO","GOR","TRC","SCC","RTV","TAR","FON","NVR","INC","MEX","FVS","SRV","LPA","PRO","GDM","CIV","LIP","TRE","TAS","AP1","AP2","AP3","ONB"';
      CritereSup:=CritereSup+',"ZLA","ZLC","ZLT","RLZ","ROO","ROF","RTP","RMP","RMS")';
      // Traitement spécif groupes de travail
      CritereSup:=CritereSup+' UNION SELECT "UCO",GRP_CODE,GRP_LIBELLE,"","" FROM GRPDONNEES order by CC_TYPE';
     end
    else if nomTable = 'CHOIXEXT' then
     begin
      // OLD critereSup := ' WHERE YX_TYPE IN ("ANP","DAU","DL1","DL2","DL3","DL4","DL5","DL6","DL7","DL8","DL9","DLA","EL1","EL2","EL3","EL4","EL5","EL6","EL7","EL8","EL9","ELA","LB1","LB2","LB3","LB4","LB5","LB6","LB7","LB8","LB9","LBA","SPE","YIN","DC2","DC2","DCG","DCI","DLC")'
{$IFNDEF GIGI}
      CritereSup:=' WHERE YX_TYPE IN ("DCG","DCI","DLC","DC2","ANP","DAU","LB1","LB2","LB3","LB4","LB5","LB6","LB7","LB8","LB9","LBA","LA1","LA2","LA3","LA4","LA5","LA6","LA7","LA8","LA9"';
      CritereSup:=CritereSup+',"LAA","LC1","LC2","LC3","LT1","LT2","LT3","LT4","LT5","LT6","LT7","LT8","LT9","LTA","LR1","LR2","LR3","LR4","LR5","LR6","LR7","LR8","LR9","LRA","EL1","EL2","EL3"';
      CritereSup:=CritereSup+',"EL4","EL5","EL6","EL7","EL8","EL9","ELA","LF1","LF2","LF3","LF4","LF5","LF6","LF7","LF8","LF9","LFA","OR1","OR2","OR3","OR4","OR5","RR1","RR2","RR3","RR4","RR5")';
{$ELSE}
        CritereSup := ',afsegmodeledet,afsegquestiondet where yx_type="SG2" and yx_code="' + identTable[7] + '" or' +
                      ' yx_type="SG3" and yx_code=asg_questionn and asg_modeleseg="' + identTable[7] + '" or' +
                      ' yx_type="SG1" and yx_code=asd_questseg and asd_questionn=asg_questionn and asg_modeleseg="'+ identTable[7] +
                      '" order by yx_type';
        SelectFrom := 'SELECT DISTINCT CHOIXEXT.* FROM '; 
//        CritereSup := ' WHERE YX_TYPE IN ("SG2","SG3","SG1")';
{$ENDIF}
     end
        // 3OB le 06/02/2008 : Traitement des tables pour segmentation - le code modèle est passé dans le 5ème poste.
      else
        if nomTable = 'AFSEGMODELEDET' then
        CritereSup := ' WHERE ASG_MODELESEG="' + identTable[7] + '"'
      else
        if nomTable = 'AFSEGQUESTIONDET' then
        CritereSup := ',AFSEGMODELEDET WHERE ASD_QUESTIONN=ASG_QUESTIONN and ASG_MODELESEG="' + identTable[7] + '"'
      else
        if nomTable = 'AFSEGREPONSE' then
        CritereSup := ',AFSEGQUESTIONDET,AFSEGMODELEDET WHERE ASD_QUESTIONN=ASG_QUESTIONN and ASR_QUESTSEG=ASD_QUESTSEG and ASG_MODELESEG="' + identTable[7] + '"'
      else
        if nomTable = 'AFSEGEXPRES_1' then
        begin
          nomTable := 'AFSEGEXPRES';
          CritereSup := ',AFSEGMODELEDET WHERE ASE_CODE=ASG_QUESTIONN and ASG_MODELESEG="' + identTable[7] + '"';
        end
      else
        if nomTable = 'AFSEGEXPRES_2' then
        begin
          nomTable := 'AFSEGEXPRES';
          CritereSup := ',AFSEGQUESTIONDET,AFSEGMODELEDET WHERE ASD_QUESTIONN=ASG_QUESTIONN and ASE_CODE=ASD_QUESTSEG and ASG_MODELESEG="' + identTable[7] + '"';
        end

        // 3OB : Fin
      else //Pas de critere pour les autres tables
      critereSup := '';
    if nomTable = 'DOSSIERGRP' then
      Rsql := OpenSql('SELECT GRP_CODE AS DOG_GROUPECONF, LDO_NODOSSIER AS DOG_NODOSSIER FROM GRPDONNEES,LIENDOSGRP '+
                      'WHERE GRP_NOM = LDO_NOM AND GRP_NOM = "GROUPECONF" AND GRP_ID = LDO_GRPID AND LDO_MARK = "X"', true)
    else
        // 3OB le 06/02/2008 : ajout si table non ''
        if nomTable <> '' then
      begin
//      Rsql := OpenSql('SELECT * FROM '+nomTable+critereSup, true);  //3OB le 03/4/08
      Rsql := OpenSql(SelectFrom+nomTable+critereSup, true);
    While Not Rsql.Eof do//parcours des résultats
      begin
      //MoveCurProgressForm('Création du fichier de sauvegarde table '+NomTable) ;
      writeLigne(rsql, f, nomTable,identTable[i]);//écrit une ligne dans le fichier
      Rsql.Next;
      end;
    MoveCurProgressForm('Création du fichier de sauvegarde table '+NomTable) ;
    result := 0;
    try
      rsql.free;
      //libération
      Finalize(TabOffset_);
    except
      PgiInfo('erreur dans la libération de mémoire');
        end;
    end;
  except
    PgiInfo('ERREUR DE GENERATION DU FICHIER TRA');
    result := -1;
  end;//TRY
  end; //FOR
  listeChps_.Free;
  listeTypes_.Free;
  FiniMoveProgressForm();
  closeFile(f);
  result := 0;
end;

{------------------------------------------------------------------------------}
procedure TImportDP.GenererFichierLog(strErreur : string);
var Fichier       : TextFile;
    Indice        : Integer;
begin
 FichierLogGen:=True;
 if (CheminFichierLog='') then CheminFichierLog:=TCBPPath.GetCegidUserDocument + '\ImportDp.log';

 AssignFile(Fichier,CheminFichierLog);
 if not (FileExists (CheminFichierLog)) then
  rewrite(Fichier)
 else
  Append (Fichier);

 Writeln(Fichier, '--- '+DateTimeToStr (Now)+' -----------------------------------');

 if strErreur<>'' then
 begin
    Writeln(Fichier, 'Erreur : '+strErreur);
 end;

 if TobErreur<>nil then
 begin
    for indice:=0 to TobErreur.Detail.count-1 do
    begin
      Writeln(Fichier, 'Ligne  : '+TobErreur.Detail [indice].GetValue ('Enregistrement'));
      Writeln(Fichier, 'Table  : '+TobErreur.Detail [indice].GetValue ('Table'));
      Writeln(Fichier, 'Champ  : '+TobErreur.Detail [indice].GetValue ('Champ'));
      Writeln(Fichier, 'Valeur : '+TobErreur.Detail [indice].GetValue ('Valeur'));
    end;
 end;

 WriteLn(Fichier, '');
 CloseFile (Fichier);
end;


//-----------------------------------------------------------------------------------
//--- Si on met une valeur supérieur à 3 caractère au champ CC_CODE qui
//--- est de type COMBO nous avons un plantage de l'appli avec le message suivant
//-----------------------------------------------------------------------------------
{function CrasherTob (SChemin : String) : Integer;
var UneTob : Tob;
    Tobs_ : TtobArray;//tableau des tobs
    NbreTob : Integer;
    Indice : Integer;
    Chk : Boolean;
begin
 NbreTob:=1;

 try
  Setlength(Tobs_,NbreTob);//allocation du tableau.
 except
  On EOutOfMemory Do result := -1;
 end;

 for Indice:=0 to NbreTob-1 do
  begin
   Tobs_[Indice]:=Tob.Create ('',nil,-1);
   TobLoadFromFile (SChemin,nil,Tobs_[Indice]);

   Try
    beginTrans;
    Tobs_[indice].SaveToFile ('C:\SauveTob',False,False,False);
    Chk:=Tobs_[Indice].insertOrUpDateDB;
    commitTrans;
   Except
    RollBack;
   End;

//   Tobs_[Indice].Free;
  end;
end; }

procedure MyAfterImport (Sender: TObject; FileGUID: String; var Cancel: Boolean);
begin
end;

end.
