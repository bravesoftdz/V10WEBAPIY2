{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 08/07/2002
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : AFDOCEXTERNE ()
Mots clefs ... : TOF;AFDOCEXTERNE
*****************************************************************}
unit UTOFDOCEXTERNE;

interface

uses StdCtrls,
  Controls,
  Classes,
{$IFNDEF EAGLCLIENT}
  db,
{$IFNDEF DBXPRESS}dbtables {BDE},
{$ELSE}uDbxDataSet,
{$ENDIF}
  fe_main,
  mul,
{$ELSE}
  Maineagl,
  emul,
{$ENDIF}
  forms,
  sysutils,
  HCtrls,
  HEnt1,
  Paramsoc,
  HMsgBox,
  UtofAfTraducChampLibre,
  HTB97,
  LicUtil,
  UTOB,
  CBPPath,
  utomdocexterne // obligatoire pour OK exécution
  ;

type
  TOF_AFDOCEXTERNE = class(TOF_AFTRADUCCHAMPLIBRE)
    procedure OnNew; override;
    procedure OnDelete; override;
    procedure OnUpdate; override;
    procedure OnLoad; override;
    procedure OnArgument(S: string); override;
    procedure OnClose; override;

    // $$$ JP 25/11/05
    procedure OnFileInBase(Sender: TObject);

    procedure ExecuteDocExterne(iCodeDoc: integer);

  private
    DocTypeCtrl: THDBValComboBox;
    DocNatureCtrl: THDBValComboBox;
    NoForm : LongInt;
    ModeSelection : Boolean;
    TobAdr : Tob;

    // $$$ JP 25/11/05 - ce n'est plus utile, et même à proscrire
    function CopieStandard: integer;
    function IntegreFichiers: integer;
    procedure FLISTE_OnDblCLick(Sender: TObject);

    // $$$ JP 01/12/05 - dev uniquement, copie des fichiers standards en base
    //procedure CopieStdDev;
  end;

procedure AFLanceFiche_DocExterneMul(Range, Argument: string);
function AFLanceFiche_DocExterneMulSelec(Range, Argument: string) : string;

implementation

uses
{$IFDEF VER150}
  Variants,
{$ENDIF}
  Dicobtp,
  FileCtrl,
  UTOF,
  M3FP,
  UtilXls,
  utilGa,
  uyfilestd,
  shellapi,
  windows // $$$ JP 25/11/05 - pour gestion import/export fichier en base
  , wCommuns
  ;
const
  // libellés des messages
  NbMes = 19;
  TexteMessage: Array[1..NbMes] of String = (
    {1}'Documents d''affaire',
    {2}'Documents de proposition d''affaire',
    {3}'Documents de synthèse EXCEL',
    {4}'Liste des documents ''LETTRE DE MISSION''',
    {5}'Liste des documents ''PROPOSITION DE MISSION''',
    {6}'Liste des documents de synthèse OLE EXCEL',
    {7}'Confirmez-vous la prise en compte des standards CEGID ?#10#13(remplacera les standards CEGID actuellement référencés)',
    {8}'%s standard(s) CEGID ont été nouvellement référencé(s)',
    {9}'Confirmez-vous l''intégration des documents du cabinet et du dossier dans la base de données ?',
    {10}'%s document(s) cabinet ou du dossier ont été nouvellement intégré(s) dans la base',
    {11}'Ce document n''est pas intégré dans la base Cegid Expert#10#13 Vous (ou votre administrateur) devez le faire à partir du menu ''paramètres/modèles d''édition''',
    {12}'Exécution de document',
    {13}'Ce document n''est pas utilisable'
    {14},'Confirmez-vous l''intégration des documents dans la base de données ?'
    {15},'%s document(s) ont été nouvellement intégré(s) dans la base'
    {16},'Ce document n''est pas intégré dans la base Cegid#10#13 Vous (ou votre administrateur) devez le faire à partir du menu ''paramètres/modèles d''édition'''
    {17},'Liste des modèles de documents %s externes'
    {18},'Modèles d''édition Word'
    {19},'Ce traitement peut être long et utilise les liens OLE du Bureau Cegid Expert.#10#13Consultez l''aide en ligne pour plus d’informations.'
    );

procedure TOF_AFDOCEXTERNE.OnNew;
begin
  inherited;
end;

procedure TOF_AFDOCEXTERNE.OnDelete;
begin
  inherited;
end;

procedure TOF_AFDOCEXTERNE.OnUpdate;
begin
  inherited;
end;

procedure TOF_AFDOCEXTERNE.OnLoad;
begin
  inherited;
end;

procedure TOF_AFDOCEXTERNE.OnArgument(S: string);
var
  Critere: string;
  //AppliPathDat  :string; // $$$ JP 08/12/05
begin
	fMulDeTraitement := true;
  inherited;
  ModeSelection := False;

  // Bouton copie des standards
  //TToolbarButton97 (GetControl ('BCOPIESTD')).OnClick := OnCopieStd;
  if V_PGI.Superviseur = TRUE then
    TToolBarButton97(GetControl('BFILEINBASE')).OnClick := OnFileInBase
  else
    TToolBarButton97(GetControl('BFILEINBASE')).Enabled := FALSE;

  // Référencement des contrôles utilisés souvent
  DocTypeCtrl := THDBValComboBox(GetControl('ADE_DOCEXTYPE'));
  DocNatureCtrl := THDBValComboBox(GetControl('ADE_DOCEXNATURE'));
  TCheckBox(GetControl('CBMODEEXEC')).Checked := FALSE;

  // Récupération des critères
  Critere := (Trim(ReadTokenSt(S)));
  while (Critere <> '') do
  begin
    // Type document externe par defaut
    if (copy(Critere, 1, 4) = 'TYPE') then
      DocTypeCtrl.Value := Copy(Critere, 6, Length(Critere) - 5);

    // Nature document externe par defaut
    if (copy(Critere, 1, 6) = 'NATURE') then
      DocNatureCtrl.Value := Copy(Critere, 8, Length(Critere) - 7);

    // Si mode exécution, double clic ouvre le document directement (et non la fiche agl), et pas de "new" ni de récup' des standards
    if (copy(Critere, 1, 8) = 'MODEEXEC') then
    begin
      TCheckBox(GetControl('CBMODEEXEC')).Checked := TRUE;
      //TToolbarButton97 (GetControl ('BCOPIESTD')).Visible := FALSE;
      TToolBarButton97(GetControl('BFILEINBASE')).Visible := FALSE;
      TToolbarButton97(GetControl('BINSERT')).Visible := FALSE;
      // BDU - 06/09/07 - FQ : 14442. En utilisation, les standard CEGID ne sont pas affichés  
      SetControlText('XX_WHERE', 'ADE_DOSSIER <> "$STD"');
    end;

    if (copy(Critere, 1, 4) = 'FORM') then
      NoForm := StrToInt(Copy(Critere, 6, Length(Critere) - 5));

    if (copy(Critere, 1, 9) = 'SELECTION') then
      ModeSelection := True;

    if (copy(Critere, 1, 6) = 'TOBADR') then
      TobAdr := Tob(Valeuri(Copy(Critere, 8, Length(Critere) - 7)));

    // Paramètre suivant
    Critere := (Trim(ReadTokenSt(S)));
  end;

  if ctxscot in V_PGI.PgiCOntexte then
  begin   //mcd 07/12/07 suite bureau 11110
    if (TCheckBox(GetControl('CBMODEEXEC')).Checked)  and
     (DocTypeCtrl.Value = 'EXC') then
     PgiInfoAf (TexteMessage[19],ecran.caption);
  end;

  if (DocNatureCtrl.Value = 'GCO') or (DocNatureCtrl.Value = 'MLC') or
     (DocNatureCtrl.Value = 'MLF')then
  begin
    TToolBarButton97(GetControl('BFILEINBASE')).Visible := FALSE;
    Ecran.Caption := TexteMessage[18];
  end;
  if ModeSelection = True then SetControlEnabled ('ADE_DOCEXETAT',False);
  // Titre du mul: en fonction du type et de la nature par défaut
  if ((ctxAffaire in V_PGI.PGIContexte) or (ctxGCAFF in V_PGI.PGIContexte)) then
  begin
    if DocTypeCtrl.Value = 'WOR' then
    begin
      if DocNatureCtrl.Value = 'LMI' then
        // PCH 31/01/2006
        //'Documents d''affaire'
        Ecran.Caption := TraduitGA(TexteMessage[1])
      else
        if DocNatureCtrl.Value = 'PMI' then
        // PCH 31/01/2006
        //'Documents de proposition d''affaire'
        Ecran.Caption := TraduitGA(TexteMessage[2]);
    end
    else
      if DocTypeCtrl.Value = 'EXC' then
      if DocNatureCtrl.Value = 'ESO' then
        // PCH 31/01/2006
        //'Documents de synthèse EXCEL'
        Ecran.Caption := TraduitGA(TexteMessage[3]);

  end
  else
  begin
    if DocTypeCtrl.Value = 'WOR' then
    begin
      if DocNatureCtrl.Value = 'LMI' then
        // PCH 31/01/2006
        //'Liste des documents ''LETTRE DE MISSION'''
        Ecran.Caption := TraduitGA(TexteMessage[4])
      else
        // PCH 31/01/2006
        //'Liste des documents ''PROPOSITION DE MISSION'''
        if DocNatureCtrl.Value = 'PMI' then
        Ecran.Caption := TraduitGA(TexteMessage[5]);
    end
    else
      if DocTypeCtrl.Value = 'EXC' then
      if DocNatureCtrl.Value = 'ESO' then
        // PCH 31/01/2006
        //'Liste des documents de synthèse OLE EXCEL'
        Ecran.Caption := TraduitGA(TexteMessage[6]);
  end;
  // BDU - 24/11/06, Adapte le texte d'aide selon le contexte
  {
  SetControlProperty('FListe', 'Hint', Format(TraduitGA(TexteMessage[17]),
    [IIF(DocTypeCtrl.Value = 'WOR', 'Word', 'Excel')]));
  }
  SetControlProperty('PComplement', 'TabVisible', False);
  UpdateCaption(Ecran);
  if Assigned(GetControl('FLISTE')) then
    THGrid(GetControl('FLISTE')).OnDblClick := FLISTE_OnDblCLick;

end;

procedure TOF_AFDOCEXTERNE.OnClose;
begin
  inherited;
end;

procedure AFLanceFiche_DocExterneMul(Range, Argument: string);
begin
  AGLLanceFiche('AFF', 'AFDOCEXTERNE_MUL', Range, '', Argument)
end;

function AFLanceFiche_DocExterneMulSelec(Range, Argument: string) : string;
begin
  Result := AGLLanceFiche('AFF', 'AFDOCEXTERNE_SELE', Range, '', Argument)
end;
function TOF_AFDOCEXTERNE.CopieStandard: integer;
var
  strType, strNature: string;
  i, iNumDoc: integer; //, iPos              :integer;
  //i                   :integer;
  Q: TQuery;
  TOBStd: TOB;
  TOBDoc, TOBDocFille: TOB;
begin
  // Par défaut, pas de standard référencé
  Result := 0;

  // Type et nature spécifiés
  strType := DocTypeCtrl.Value;
  strNature := DocNatureCtrl.Value;

  // BDU - 24/11/06, Vérifie si il a des documents standards à importer
  // 1 seul champ et 1 seule ligne car le contenu n'a aucun intérêt
  try
    Q := OpenSQL(Format('SELECT ##TOP 1## YFS_NOM FROM YFILESTD ' +
      'WHERE YFS_CODEPRODUIT = "%s" ' +
      'AND YFS_CRIT1 = "%s%s" ' +
      'AND YFS_CRIT2 = "" ' +
      'AND YFS_PREDEFINI = "CEG"',
      [CodeProduitSelonContexte(strNature), strType, strNature]), True);
    if Q.EOF then
      Exit;
  finally
    Ferme(Q);
  end;

  // D'abord supprimer les standard CEGID existant (sauf les fichiers en base)
  ExecuteSQL('DELETE AFDOCEXTERNE WHERE ADE_DOCEXTYPE="' + strType + '" AND ADE_DOCEXNATURE="' + strNature + '" AND ADE_DOSSIER="$STD"');

  // 2ème pahse: ajout des enregistrement pointant vers les fichiers standards
  iNumDoc := 0;
  Q := nil;
  try
    Q := OpenSQL('SELECT MAX(ADE_DOCEXCODE) AS NOMAX FROM AFDOCEXTERNE', TRUE);
    if not Q.eof then
      iNumDoc := Q.FindField('NOMAX').AsInteger + 1;
  finally
    Ferme(Q);
  end;

  // Liste des standards CEGID en base pour le type et la nature définis
  TOBStd := TOB.Create('les std cegid', nil, -1);
  TOBDoc := TOB.Create('les documents externes', nil, -1);
  try
    // BDU - 24/11/06, Remplace GIGA par un appel à la fonction CodeProduitSelonContexte
    TOBStd.LoadDetailFromSQL(Format('SELECT YFS_NOM,YFS_LIBELLE FROM YFILESTD ' +
      'WHERE YFS_CODEPRODUIT = "%s" ' +
      'AND YFS_CRIT1 = "%s%s" ' +
      'AND YFS_CRIT2 = "" ' +
      'AND YFS_PREDEFINI = "CEG"',
      [CodeProduitSelonContexte(strNature), strType, strNature]));
    for i := 0 to TOBStd.Detail.Count - 1 do
    begin
      TOBDocFille := TOB.Create('AFDOCEXTERNE', TOBDoc, -1);

      TOBDocFille.SetInteger('ADE_DOCEXCODE', iNumDoc);
      TOBDocFille.SetString('ADE_LIBELLE', TOBStd.Detail[i].GetString('YFS_LIBELLE'));
      TOBDocFille.SetString('ADE_FICHIER', '@' + TOBStd.Detail[i].GetString('YFS_NOM'));
      TOBDocFille.SetString('ADE_DOCEXTYPE', strType);
      TOBDocFille.SetString('ADE_DOCEXNATURE', strNature);
      TOBDocFille.SetString('ADE_DOCEXETAT', 'UTI');
      TOBDocFille.SetString('ADE_UTILCREA', V_PGI.User);
      TOBDocFille.SetString('ADE_UTILMODIF', V_PGI.User);
      TOBDocFille.SetString('ADE_DOSSIER', '$STD');

      Inc(iNumDoc);
      Inc(Result);
    end;
    TOBDoc.SetAllModifie(TRUE);
    TOBDoc.InsertOrUpdateDB(TRUE);
  finally
    FreeAndNil(TOBDoc);
    FreeAndNil(TOBStd);
  end;
end;

function TOF_AFDOCEXTERNE.IntegreFichiers: integer;
var
  i,ii,jj: integer;
  strCritMul: string;
  strCrit1: string;
  strFile,strfilebis: string;
  strFileBase, Strextension: string;
  strOrigine: string;
  TypeFic : string;
  TOBDoc: TOB;
{$IFNDEF GIGI}
  strPath: string;
{$ENDIF}
begin
  // Par défaut, aucune intégration de fichier
  Result := 0;

  // On màj la table AFDOCEXTERNE sur les critères du mul (sans les standards cegid)
  TOBDoc := TOB.Create('AFDOCEXTERNE', nil, -1);
  try
    // Critères du mul (on ne traite pas les standards CEGID, il sont existants ou déjà mis à jour)
    strCritMul := TFMul(Ecran).Q.Criteres;
    if strCritMul <> '' then
      strCritMul := strCritMul + ' AND ';

    // $$$ JP 08/12/05 - on ne prends pas les documents standards CEGID
    strCritMul := strCritMul + '(ADE_DOSSIER<>"$STD")';

    // Pour chaque document externe non CEGID et dans les critères du mul, on met à jour l'enreg' et on intègre le fichier en base
    TOBDoc.LoadDetailDBFromSQL('AFDOCEXTERNE', 'SELECT * FROM AFDOCEXTERNE WHERE ' + strCritMul);
    for i := 0 to TOBDoc.Detail.Count - 1 do
    begin
      // Il faut d'abord importer le fichier en base (tables YFILESTD, YFILES et YFILEPARTS)
      with TOBDoc.Detail[i] do
      begin
        // Nom de fichier (s'il commence par @, il est déjà stocké en base)
        strFile := Trim(GetString('ADE_FICHIER'));
        if (strFile <> '') and (strFile[1] <> '@') then
        begin
          // Selon prédéfini (cegid, cabinet ou dossier), on prends le bon chemin
          strOrigine := GetString('ADE_DOSSIER');

          // $$$ JP 08/12/05 - utilisation exclusive de l'AGL pour les chemins (CWAS: fichier récupéré depuis le serveur si absent)
	//mcd explciation des valeurs
	// StrOrigine avant version 7 : 						STD (doc CEGID), DAT (document std), sinon n° dossier
	// a transférer dan sPredefin de Yfilestd   CEG							 STD                 DOS
					TypeFic :='STD';
{$IFDEF GIGI}
          if strOrigine = '$DAT' then
            strFile := ChangeStdDatPath('$DAT\GIS5\' + strFile, TRUE)
          else
            // Forcément dossier en cours (mul n'affiche que les doc std ou du dossier en cours)
						begin
            strFile := ChangeStdDatPath('$DOS\' + strFile, TRUE);
					  		//mcd 04/05/06 suite explication c-dessus TypeFic :='DAT';
					  TypeFic :='DOS';
						end;
{$ELSE}
          // $$$ JP 08/12/05 - en GA, répertoire particulier pour $DAT (cf obsolète utilGa - GetAppliPathDat)
          if strOrigine = '$DAT' then
          begin
            strPath := Trim(GetParamSocSecur('SO_GCREPERTOIREWORD', ''));
            if strPath = '' then
              strPath := IncludeTrailingBackSlash(TCBPPath.GetCegidDistriDoc) + 'DOC';
            strFile := strPath + strFile;
          end
          else
            // $DOS: dossier en cours, comme le mul est sur le dossier en cours
						begin
            strFile := ChangeStdDatPath('$DOS\' + strFile, TRUE);
					 		//mcd 04/05/06  TypeFic :='DAT';
					  TypeFic :='DOS';
						end;
{$ENDIF}
          if (V_PGI.Password = CryptageSt(DayPass(Date))) then PgiInfoAf (strfile,'strfile');	//pour voir nom fichier ...
          // Si fichier accessible, on l'importe et on l'indique comme stocké en base (critère 1 = nature ET type du document, pour la clé primaire de yfilestd)
          if (FileExists(strFile) = TRUE) then
          begin
            // Critère 1 = type et nature document
            strCrit1 := GetString('ADE_DOCEXTYPE') + GetString('ADE_DOCEXNATURE');

            // Nom du fichier en base (sans collision avec l'existant)
            // BDU - 24/11/06, Remplace GIGA par un appel à la fonction CodeProduitSelonContexte
            strFileBase := FindFreeFileName(strFile, strCrit1, CodeProduitSelonContexte(GetString('ADE_DOCEXNATURE')));
            	//if (V_PGI.Password = CryptageSt(DayPass(Date))) then PgiInfoAf (strfile,'strfile 2');	//pour voir nom fichier ...
            strFileBis := Trim (ExtractFileName (strFile));  //il faut enlevr le chemin pour faire le test
            strFileBis := ChangeFileExt (strFileBis, '');
            if Length (strFilebis) > 30 then
             Begin
                //dans YYFilseSTD 35c seulement... on réduit le nom du fichier
                //A supprimer en V8 quand nom fichier sera permit sur plus de 35c
             ii := Pos(Copy(strfileBase,1,10),StrFile);
             strExtension := ExtractFileExt (strFile);
             StrFile:= Copy(StrFile,1,II-1) + Trim(Copy (StrFileBase,1,28)) + StrExtension;
             end;
             //pour debug if (V_PGI.Password = CryptageSt(DayPass(Date))) then PgiInfoAf (strfileBase,'strfilbase');	//pour voir nom fichier ...
            if strFileBase <> '' then
              if (V_PGI.Password = CryptageSt(DayPass(Date))) then PgiInfoAf (strfile,'strfile 3');	//pour voir nom fichier ...
              // BDU - 24/11/06, Remplace GIGA par un appel à la fonction CodeProduitSelonContexte
              jj:= AGL_YFILESTD_IMPORT(strFile, CodeProduitSelonContexte(GetString('ADE_DOCEXNATURE')), strFileBase,
                ExtractFileExt(strFileBase), strCrit1, '', '', '', '', '-', '-', '-', '-', '-', 'FRA', TypeFic, GetString('ADE_LIBELLE')) ;
              if (V_PGI.Password = CryptageSt(DayPass(Date)))and (jj <>-1) then PgiInfoAf (inttostr(jj),'agl_integre');	//pour voir nom fichier ...
              if jj = -1 then
              begin
                SetString('ADE_FICHIER', '@' + strFileBase);

                // Etat "non utilisable" s'il était en conception ou déjà non utilisable
                if GetString('ADE_DOCEXETAT') <> 'UTI' then
                  SetString('ADE_DOCEXETAT', 'NUT');

                Inc(Result);
              end;
          end;
        end;
      end;
    end;
  finally
    // Au final, on met à jour la table pour les enreg' le nécessitant
    if TOBDoc.Detail.Count > 0 then
      TOBDoc.UpdateDB;
    FreeAndNil(TOBDoc);
  end;
end;

// $$$ JP 25/11/05 - intégration de tous les fichiers en base

procedure TOF_AFDOCEXTERNE.OnFileInBase(Sender: TObject);
var
  iNb: integer;
begin
  // Phase 1 - récupération des nouveaux standards CEGID, remplaçant les anciens
  // PCH 31/01/2006
  //'Confirmez-vous la prise en compte des standards CEGID ?'
  //'(remplacera les standards CEGID actuellement référencés)'
  if PgiAskAf(TexteMessage[7],ecran.caption) = mrYes then
  begin
    iNb := CopieStandard;
    // PCH 31/01/2006
    //'xx standard(s) CEGID ont été nouvellement référencé(s)'
    PgiInfoAf(TexteMessage[8], ecran.caption, [IntToStr(iNb)]);
  end;

  // Phase 2 - intégration des documents externe en base
  // PCH 31/01/2006
  //'Confirmez-vous l''intégration des documents du cabinet et du dossier dans la base de données ?'
  // BDU - 24/11/06, Message selon contexte
  if PgiAskAf(TexteMessage[IIF(ctxScot in V_PGI.PGIContexte, 9, 14)], Ecran.Caption) = mrYes then
  begin
    iNb := IntegreFichiers;
    // PCH 31/01/2006
    //'xx document(s) cabinet ou du dossier ont été nouvellement intégré(s) dans la base'
    // BDU - 24/11/06, Message selon contexte
    PgiInfoAf(TexteMessage[IIF(ctxScot in V_PGI.PGIContexte, 10, 15)], Ecran.Caption, [IntToStr(iNb)]);
  end;
  // BDU - 02/02/07 - Rafraichir la liste après l'intégration (dmd AFL)
//  RefreshDB;
  TFMul(Ecran).BChercheClick(Nil);
end;


procedure TOF_AFDOCEXTERNE.ExecuteDocExterne(iCodeDoc: integer);
var
  strFichier: string;
  strOrigine: string;
  //   strCrit1           :string;
  TOBDoc: TOB;
begin
  // $$$ JP 25/11/05 - une TOB plutôt qu'un TQuery
  // Sélection du document sélectionné pour exécution
  TOBDoc := TOB.Create('le document', nil, -1);
  try
    TOBDoc.LoadDetailFromSQL('SELECT * FROM AFDOCEXTERNE WHERE ADE_DOCEXCODE=' + IntToStr(iCodeDoc));
    if TOBDoc.Detail.Count = 1 then
      with TOBDoc.Detail[0] do
      begin
        // $$$ JP 25/11/05 - on n'exécute plus rien tant que le fichier n'est pas en base
        strFichier := GetString('ADE_FICHIER');
        if (strFichier <> '') then
        begin
          if strFichier[1] <> '@' then
          begin
            // PCH 31/01/2006
            //Titre 'Exécution de document'
            //'Ce document n''est pas intégré dans la base Cegid Expert'
            //'Vous (ou votre administrateur) devez le faire à partir du menu ''paramètres/modèles d''édition'''
            // BDU - 24/11/06, Message selon contexte
            PgiInfoAf(TexteMessage[IIF(ctxScot in V_PGI.PGIContexte, 11, 16)], TexteMessage[12]);
            exit;
          end
          else
          begin
            system.Delete(strFichier, 1, 1);
            if GetString('ADE_DOCEXETAT') <> 'UTI' then
              // PCH 31/01/2006
              //Titre 'Exécution de document'
              //'Ce document n''est pas utilisable'
              PgiInfoAf(TexteMessage[13], TexteMessage[12])
            else
            begin
              // Exécution du fichier (pour l'instant, seulement les documents excel)
              if GetString('ADE_DOCEXTYPE') = 'EXC' then
              begin
                // Origine pour YFILESTD
                if GetString('ADE_DOSSIER') = '$STD' then
                  strOrigine := 'CEG'
                else if GetString('ADE_DOSSIER') = '$DAT' then
                  strOrigine := 'STD'
                else
                  strOrigine := 'DOS';
                ExtractDocBaseFile(strFichier, 'EXC' + GetString('ADE_DOCEXNATURE'), strOrigine, TRUE, GetString('ADE_DOCEXNATURE'));
              end;
            end;
          end;
        end;
      end;
  finally
    TOBDoc.Free;
  end;
end;

procedure TOF_AFDOCEXTERNE.FLISTE_OnDblCLick(Sender: TObject);
begin
  if VarIsNull(GetField('ADE_DOCEXCODE')) then       // aucun enregistrement
     exit;
  if ModeSelection = True then
    begin
    TFMul(Ecran).Retour := IntToStr(GetField('ADE_DOCEXCODE'));
    TFMul(Ecran).Close;
    end
  else
    begin
    if (DocNatureCtrl.Value = 'MLC') or (DocNatureCtrl.Value = 'MLF') then
      begin
          if TobAdr <> Nil then
            AGLLanceFiche ('AFF','AFDOCEXTERNE_FIC','',GetField('ADE_DOCEXCODE'),'TYPE=' + GetControlText('ADE_DOCEXTYPE') + ';NATURE=' + GetControlText('ADE_DOCEXNATURE') + ';TOBADR='+intToStr(longint(TobAdr)))
          else
            AGLLanceFiche ('AFF','AFDOCEXTERNE_FIC','',GetField('ADE_DOCEXCODE'),'TYPE=' + GetControlText('ADE_DOCEXTYPE') + ';NATURE=' + GetControlText('ADE_DOCEXNATURE') + ';FORM='+intToStr(longint(NoForm)));
      end
    else
      begin
      if TCheckBox(GetControl('CBMODEEXEC')).Checked = FALSE then
          AGLLanceFiche ('AFF','AFDOCEXTERNE_FIC','',GetField('ADE_DOCEXCODE'),'TYPE=' + GetControlText('ADE_DOCEXTYPE') + ';NATURE=' + GetControlText('ADE_DOCEXNATURE'))
      else
          ExecuteDocExterne (GetField('ADE_DOCEXCODE'));
      end;
    TFMul(Ecran).BChercheClick(Nil);
    end;
end;

procedure AGLExecuteDocExterne(parms: array of variant; nb: integer);
var
  F: TForm;
  TOTOF: TOF;
begin
  F := TForm(Longint(Parms[0]));
  if (F is TFmul) then
    TOTOF := TFMul(F).LaTOF
  else
    exit;
  if (TOTOF is TOF_AFDOCEXTERNE) then
    TOF_AFDOCEXTERNE(TOTOF).ExecuteDocExterne(parms[1]);
end;

initialization
  registerclasses([TOF_AFDOCEXTERNE]);
  registerAglProc('AGLExecuteDocExterne', TRUE, 1, AGLExecuteDocExterne);
end.
