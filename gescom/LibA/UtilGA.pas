unit UtilGa;

// source contenant toutes les fct génériques GI/GA, pas liée à une fct
// et demandent peu de uses en cascade..

interface
Uses
{$IFDEF EAGLCLIENT}

{$ELSE}
      {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
  {$ifdef GIGI}
       galoutil,
  {$ENDIF}
{$ENDIF}
       hctrls, menus, Dicobtp, ent1, Classes, SysUtils, entgc ,Utob
       ,ed_tools ,ParamSoc ,HmsgBox ,Controls, CBPPath
{$ifdef GIGI}
       , PgiEnv
{$ENDIF}
			 , Masks
       , uYFILESTD2
       , ShellApi
       ;




Procedure ChargePopUp(Menu:TpopUpMenu;Fct:TnotifyEvent;table,req:string);
Procedure AnnuleCutOff;
{$IFNDEF EAGLCLIENT}
Procedure GIGABasculeAna(Parle:boolean);
{$ifdef GIGI}
Procedure AnnuleGI;
{$ENDIF}
{$ENDIF}


function GetAppliPathStd:string;
function GetAppliPathDat:string;
function IsValidFileName (strFileName: string): boolean;
function CodeProduitSelonContexte (Nature : string = '') : String;
function FindFreeFileName (strFileName: string; strCrit1: string; strCodeProduit: string): string;
function ExtractDocBaseFile (strFile: string; strCrit1: string; strOrigine: string; bOpenFile: boolean; strNature : string = ''): string;


implementation

uses windows, filectrl, hent1;


function IsValidFileName (strFileName: string): boolean;
begin
  Result := not MatchesMask (strFileName, '*[\/:*?"<>|]*')
end;

function CodeProduitSelonContexte (Nature : string = '') : String;
begin
  if Nature = 'GCO' then Result := 'GCS5'
  else if Nature = 'MLC' then Result := 'GRC'
  else if Nature = 'MLF' then Result := 'GRF'
  else if ctxScot in V_PGI.PGIContexte then Result := 'GIGA'
  else Result := 'GAS5';
end;

{***********A.G.L.***********************************************
Auteur  ...... : MC DESSEIGNET
Créé le ...... : 14/04/2003
Modifié le ... :   /  /    
Description .. : fonction qu sera utiliser depuis le DP
Suite ........ : Va permettre d'annuler l'assistant de création de la GI lancé
Suite ........ : à tort.
Suite ........ : Va détruire enrgt dans dossappli
Suite ........ : mais permet aussi éventuellement de tout détruire les tables 
Suite ........ : de la GI déjà créer
Suite ........ : Affaire, activite, piece...
Mots clefs ... : ANNULATION GI
*****************************************************************}
{$IFNDEF EAGLCLIENT}
{$ifdef GIGI}
Procedure AnnuleGI;
var titre : string;
begin
Titre :='Annulation GI';
If ExisteSql('SELECT AFF_AFFAIRE FROM AFFAIRE') then
  begin
  if PgiAsk('Il existe des missions. Voulez-vous les détruire ?',Titre) = mrNo then Exit
   else
    begin
    If PgiAsk('Voulez-vous vraiment détruire votre dossier GI ?',titre) = mrNo then exit
    else if PgiAsk ('Le traitement sera irréversible. Voulez-vous annuler l''opération ?',titre) = mrYes then exit
         else
          begin   // plusieurs confirmation on annule tout
          InitMoveProgressForm (nil,'Traitement en cours...','', 10, false, true) ;
          ExecuteSql ('DELETE FROM AFFTIERS');
          ExecuteSql ('DELETE FROM FACTAFF');
          ExecuteSql ('DELETE FROM PIECE');
          ExecuteSql ('DELETE FROM LIGNE');
          ExecuteSql ('DELETE FROM PIEDBASE');
          ExecuteSql ('DELETE FROM PIEDECHE');
          ExecuteSql ('DELETE FROM PIEDPORT');
          ExecuteSql ('DELETE FROM ACTIVITE');
          ExecuteSql ('DELETE FROM AFCUMUL');
          ExecuteSql ('DELETE FROM AFFAIRE');
          FiniMoveProgressForm;
          end;
   end;
  end;
if GetFlagAppli ('CGIS5.EXE') then   // PL le 17/04/03 : chgt IsAppliActive de AssistPL en GetFlagAppli de galoutil
  begin
  If PgiAsk ('Confirmez-vous la désactiviation de la Gestion Interne',titre) = MrYes then
    begin
    executeSql ('DELETE FROM DB000000.DBO.DOSSAPPLI WHERE DAP_NODOSSIER="'+ V_PGI_ENV.nodossier+ '" AND DAP_NOMEXEC="CGIS5.EXE"');
    PgiInfo ('Traitement terminé',titre);
    end
  else exit;
  end
else begin
  PgiInfo ('La Gestion Interne n''est pas activée dans cette base',Titre);
  end;
end;
{$endif}
{$endif}

{***********A.G.L.***********************************************
Auteur  ...... : MC DESSEIGNET
Créé le ...... : 15/10/2002
Modifié le ... : 15/10/2002
Description .. : Fct qui permet de renseigner un popupmenu avec les info
Suite ........ : d'une tablette. reçoit le menu popup, le code de la tablette
Suite ........ : (obligatoirement dans commun, avec le libellé dans
Suite ........ : CO_LIBELLE, et le champ dans CO_ABREGE), la fct à
Suite ........ : activer sur le clic du pop up
Suite ........ : reçoit aussi la requête SQL à ajouter au test du type de
Suite ........ : table dans le cas ou en fct des paramsoc, on doit prendre
Suite ........ : ou non certain code
Mots clefs ... : POPUPMENU;TABLETTE
*****************************************************************}
Procedure ChargePopUp(Menu:TPopUpMenu;Fct:TnotifyEvent;table,req:string);
var     T     : TMenuItem ;
 QQ: Tquery;
 ii:integer;
 lib:string;
begin
PurgePopup(Menu) ;
QQ:=   OpenSql('SELECT CO_LIBELLE,CO_ABREGE FROM COMMUN WHERE CO_TYPE="'+table+'"'+req,true);
if Not QQ.EOF then begin
    While not QQ.EOF do
      BEGIN
      T:=TMenuItem.Create(Menu) ;
      T.Name:=QQ.FindField('CO_ABREGE').asString;
      lib:= QQ.FindField('CO_LIBELLE').asString;
      ii:= pos('PARTIE',AnsiUppercase(lib));
      if ( ii >0 )then
        begin
        if  copy (lib,ii+7,1) ='1' then T.Caption:=VH_GC.CleAffaire.Co1Lib
        else  if  copy (lib,ii+7,1) ='2' then if VH_GC.CleAffaire.Co2Visible then T.Caption:=VH_GC.CleAffaire.Co2Lib
        else  if  copy (lib,ii+7,1) ='3' then if VH_GC.CleAffaire.Co3Visible then T.Caption:=VH_GC.CleAffaire.Co3Lib;
        end
      else  T.Caption := TraduitGa(lib);
      T.OnClick:=Fct ;
      Menu.Items.Add(T) ;
      QQ.Next; // il y a peu d'enrgt dans cette tablette, on peut faire le QQ.next
      end;
   end;
Ferme(QQ) ;
End;


{$IFNDEF EAGLCLIENT}
Procedure GIGABasculeAna (Parle:boolean);
var qq:Tquery;
TobAna,TobAnaDet,TobAnaNou,TobAnaNouDet:TOB;
Wi,jj:Integer;
text:string;
begin

if ExisteSql('SELECT GDA_AXE FROM DECOUPEANA') then exit;  // il existe déjà des info dans la table. on ne fait rien
if not (ExisteSql('SELECT AST_AXE FROM STRUCRANAAFFAIRE')) then exit; // il n'y a rien dans la table ana ... on sort
// fct qui permet de basculer de la table GI/GA : STRUCTANAAFFAIRE
// à la table générique DECOUPEANA
InitMoveProgressForm (Nil,'Modification Analytique','ddd',10, False,True);
TobAna:=Tob.Create('STRUCRANAAFFAIRE',NIL,-1);
  // select * normal, il faut tout lire pour tout transferer
QQ := OpenSQL('SELECT * FROM STRUCRANAAFFAIRE',true) ;
If Not QQ.EOF then TobAna.LoadDetailDB('STRUCRANAAFFAIRE','','',QQ,True);
Ferme(QQ);
If TObAna.detail.count >0 then begin
    TobAnaNou:=Tob.Create('DECOUPEANA',NIL,-1);
    for wi := 0 to TObAna.detail.count-1 do
      begin
        MoveCurProgressForm('');
        TobAnaDet := TobAna.detail[wi];
        TobAnaNouDet:= Tob.create('DECOUPEANA',TobAnanOu,-1);
        TobAnaNouDet.putvalue('GDA_AXE',TobAnaDet.Getvalue('AST_AXE'));
        TobAnaNouDet.putvalue('GDA_TYPECOMPTE','VEN');
        TobAnaNouDet.putvalue('GDA_ETABLISSEMENT',TobAnaDet.Getvalue('AST_ETABLISSEMENT'));
        TobAnaNouDet.putvalue('GDA_TYPESTRUCTANA',TobAnaDet.Getvalue('AST_TYPESTRUCRANA'));
        TobAnaNouDet.putvalue('GDA_RANG',TobAnaDet.Getvalue('AST_RANG'));
        QQ := OpenSQL('SELECT CO_LIBRE FROM COMMUN WHERE CO_TYPE="ACL" AND CO_CODE="'
           + TobAnaDet.Getvalue('AST_CHAMPS')+ '"',true) ;
        If Not QQ.EOF then begin
           Text:= QQ.FindField('CO_LIBRE').AsString;
           jj:=Pos(';',Text);
           if JJ >0 then TobAnaNouDet.putvalue('GDA_LIBCHAMP',Copy(Text,1,jj-1));
           end;
        Ferme(QQ);
        TobAnaNouDet.putvalue('GDA_ISLIBELLE','-');
        TobAnaNouDet.putvalue('GDA_POSITION',TobAnaDet.Getvalue('AST_POSITION'));
        TobAnaNouDet.putvalue('GDA_LONGUEUR',TobAnaDet.Getvalue('AST_LONGUEUR'));
        TobAnaNouDet.putvalue('GDA_CREATSSSECTION',TobAnaDet.Getvalue('AST_CREATSSSECTION'));
        end;
    TobAnaNou.InsertOrUpdatedb(False);
    TobAnaNou.free;
    SetParamSOc ('SO_GCAXEANALYTIQUE',TRUE);  //mcd 08/01/03 force analytique à vrai pour OK depuis pgimajver
    end;
FiniMoveProgressForm;
TobAna.free;
if Parle then PgiBoxAf('Traitement termine','');
end;
{$ENDIF}

{***********A.G.L.***********************************************
Auteur  ...... : MC DESSEIGNET
Créé le ...... : 02/06/2003
Modifié le ... :   /  /    
Description .. : Fct qui permet de détruire tous les CUtOff qui ont pu être
Suite ........ : générer dans le fichier Afcumul.
Suite ........ : permet ensuite de changer le mode de génération du cut off
Mots clefs ... : CUTOFF
*****************************************************************}
Procedure AnnuleCutOff;
var titre : string;
begin
Titre :='Annulation CutOff';
If ExisteSql('SELECT ACU_TYPEAC FROM AFCUMUL WHERE ACU_TYPEAC="CAC" OR ACU_TYPEAC="CVE"') then
  begin
  if PgiAsk('Il existe des Cut Off. Voulez-vous les détruire ?',Titre) = mrNo then Exit
   else
    begin
    If PgiAsk('Voulez-vous vraiment détruire vos cut off ?',titre) = mrNo then exit
    else if PgiAsk ('Le traitement sera irréversible. Voulez-vous annuler l''opération ?',titre) = mrYes then exit
         else
          begin   // plusiers confirmation on annule tout
          InitMoveProgressForm (nil,'Traitement en cours...','', 10, false, true) ;
          ExecuteSql ('DELETE FROM AFCUMUL WHERE ACU_TYPEAC="CAC" OR ACU_TYPEAC="CVE"');
          FiniMoveProgressForm;
          end;
   end;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : J Pomat
Créé le ...... : 07/07/2003
Modifié le ... :   /  /
Description .. : Récupère "à tout prix" le répertoire des standards pour
Suite ........ : GI/GA (normalement C:\PGI00\STD\GIS5 ou GAS5)
Mots clefs ... : CHEMIN;STD;STANDARD
*****************************************************************}
function GetAppliPathStd:string;
var cDrive    : char;
    CPath     : String;
begin
     Result := '';

     //on remplace V_PGI.StdPath par CPath
     CPath := IncludeTrailingBackSlash(TCBPPath.GetCegidDistriStd);

{$IFNDEF EAGLCLIENT}
     if (V_PGI <> nil) AND (Trim (CPath) <> '') then
  {$IFDEF GIGI}
         Result := Trim (CPath) + '\GIS5'
  {$ELSE}
         Result := Trim (CPath) + '\GAS5'
  {$ENDIF}
     else
     begin
          // Répertoire non connu: pas normal, mais on essaie de trouver PGI00\STD sur les lecteurs locaux
          cDrive := 'A';
          while cDrive <> ' ' do
          begin
               case GetDriveType (pchar (cDrive + ':\')) of
                    0,
                    1,
                    DRIVE_REMOVABLE,
                    DRIVE_REMOTE,
                    DRIVE_CDROM:
                    if cDrive = 'Z' then
                       cDrive := ' '
                    else
                       Inc (cDrive);
               else
  {$IFDEF GIGI}
                   Result := cDrive + ':\PGI00\STD\GIS5';
  {$ELSE}
                   Result := cDrive + ':\PGI00\STD\GAS5';
  {$ENDIF}
                   if DirectoryExists (Result) = FALSE then
                   begin
                        Result := '';
                        Inc (cDrive);
                   end
                   else
                       cDrive := ' ';
               end;
          end;
     end;
{$ENDIF}
end;

{***********A.G.L.***********************************************
Auteur  ...... : J Pomat
Créé le ...... : 07/07/2003
Modifié le ... :   /  /
Description .. : // $$$ JP 07/07/2003: récupère "à tout prix" le répertoire
Suite ........ : partagé des DAT pour GI/GA (normalement
Suite ........ : \\serveur\PGI01$\DAT\GIS5 ou GAS5)
Mots clefs ... : CHEMIN;DAT
*****************************************************************}
function GetAppliPathDat:string;
var
   cDrive    :char;
begin
     Result := '';

{$IFNDEF EAGLCLIENT}
     if (V_PGI <> nil) AND (Trim (V_PGI.DatPath) <> '') then
  {$IFDEF GIGI}
        Result := Trim (V_PGI.DatPath) + '\GIS5';
  {$ELSE}
        Result := Trim (V_PGI.DatPath) + '\GAS5';
  {$ENDIF}
{$ENDIF}
end;


function FindFreeFileName (strFileName: string; strCrit1: string; strCodeProduit: string): string;
var
  iNumDup: integer;
  iPos,ii: integer;
  strExtension,OldName: string;
begin
  // on travail sur le nom du fichier seulement, sans chemins
  OldName :=strFileName;
  strFileName := Trim (ExtractFileName (strFileName));
  strExtension := ExtractFileExt (strFileName);

  // Par défaut, le nom d'origine est OK (généralement le cas)
  Result := strFileName;

  // Si indexation demandée, on recherche le séparateur d'index dans le nom ("_")
  strFileName := ChangeFileExt (strFileName, '');
  iPos := Length (strFileName);
  if iPos > 30 then
   Begin
			//dans YYFilseSTD 35c seulement... on réduit le nom du fichier
			//A supprimer en V8 quand nom fichier sera permit sur plus de 35c
   strFileName := Trim(Copy (StrFileName,1,28)) + StrExtension;
   ii := Pos(Copy(strfilename,1,10),Oldname);
   RenameFile (Oldname, Copy(OldName,1,II-1) + Strfilename);
   iPos := Length (strFileName);
   Result := strFileName;
   strFileName := ChangeFileExt (strFileName, '');
	 end;
  while (iPos > 0) and (strFileName [iPos] <> '_') do
    Dec (iPos);
  if iPos > 0 then
    Delete (strFileName, iPos, Length (strFileName) - iPos + 1);

  // Tant que le nom de fichier existe déjà en base, on construit un nom approché: on concatène un n° d'ordre
  iNumDup := 0;
  while (iNumDup < 1000) and (ExisteSQL ('SELECT 1 FROM YFILESTD WHERE YFS_CODEPRODUIT="' + strCodeProduit + '" AND YFS_NOM="' + Result + '" AND YFS_LANGUE="FRA" AND YFS_CRIT1="' + strCrit1 + '"') = TRUE) do
  begin
    Inc (iNumDup);
    Result := strFileName + '_' + IntToStr (iNumDup) + strExtension;
  end;
  if iNumDup = 1000 then
    Result := '';
end;


function ExtractDocBaseFile (strFile: string; strCrit1: string; strOrigine: string; bOpenFile: boolean; strNature : string = ''): string;
var
  strExtFichier: string;
  TOBFiles: TOB;
  i: integer;
begin
  // Extraction du fichier ET de tous les fichiers dépendants dans un répertoire temporaire
  // BDU - 24/11/06, Remplace GIGA par un appel à la fonction CodeProduitSelonContexte
  if AGL_YFILESTD_EXTRACT (Result, CodeProduitSelonContexte(strNature), strFile, strCrit1,
    '', '', '', '', FALSE, 'FRA', strOrigine) = -1 then
  begin
    TOBFiles := TOB.Create ('les fichiers', nil, -1);
    try
      // $$$ JP 30/11/05 - Plus la peine de piloter Excel pour ouvrir en lecture seul, le fichier est stocké en base
      // $$$ JP 30/11/05 - Il faut extraire au même endroit toutes les extensions pour ce type/nature de document (exemple: .xla, .dot, ...)
      // $$$ JP 30/11/05 - Il faut déplacer les fichiers extension dans le répertoire du fichier à ouvrir
      // BDU - 24/11/06, Remplace GIGA par un appel à la fonction CodeProduitSelonContexte
      TOBFiles.LoadDetailFromSQL (Format('SELECT YFS_NOM FROM YFILESTD ' +
        'WHERE YFS_CODEPRODUIT = "%s" ' +
        'AND YFS_CRIT1 = "%s" ' +
        'AND YFS_CRIT2 = "EXT" AND ' +
        'YFS_PREDEFINI = "CEG"',
        [CodeProduitSelonContexte(strNature), strCrit1]));
      for i := 0 to TOBFiles.Detail.Count - 1 do
        // BDU - 24/11/06, Remplace GIGA par un appel à la fonction CodeProduitSelonContexte
        if AGL_YFILESTD_EXTRACT (strExtFichier, CodeProduitSelonContexte(strNature),
          TOBFiles.Detail [i] .GetString ('YFS_NOM'), strCrit1, 'EXT') = -1 then
          MoveFileEx (pchar (strExtFichier), pchar (ExtractFilePath (Result) + ExtractFileName (strExtFichier)), MOVEFILE_REPLACE_EXISTING);
    finally
      FreeAndNil (TOBFiles);
    end;

    // Lancement du fichier à exécuter
    if (bOpenFile = TRUE) and (FileExists (Result) = TRUE) then
      ShellExecute (0, pchar ('open'), pchar (Result), nil, nil, SW_RESTORE);
  end;
end;

end.

