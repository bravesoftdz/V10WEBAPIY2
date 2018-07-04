{***********UNITE*************************************************
Auteur  ...... : JP
Créé le ...... : 08/07/2002
Modifié le ... :   /  /
Description .. : Source TOM de la TABLE : AFDOCEXTERNE
Mots clefs ... : TOM;AFDOCEXTERNE
*****************************************************************}
unit UTOMDOCEXTERNE;

interface                        

uses
  Controls,
  Classes,
  HQry,
{$IFNDEF EAGLCLIENT}
  db,
{$IFNDEF DBXPRESS}dbtables {BDE},
{$ELSE}uDbxDataSet,
{$ENDIF}
  Fiche,
  FE_Main,
  dbctrls,
  mul,
{$ELSE}
  eFiche,
  MaineAGL,
  utileAGL,
  emul,
{$ENDIF}
  gr_editbox,
  forms,
  sysutils,
  HCtrls,
  HEnt1,
  HMsgBox,
  UTom,
  windows,
  HTB97,
  HDB,
  AGLUtilOLE,
  utob // $$$ JP 30/11/05
  ,wCommuns
  ,UtilWord
  ;
type
  TOM_AFDOCEXTERNE = class(TOM)
    procedure OnNewRecord; override;
    procedure OnDeleteRecord; override;
    procedure OnUpdateRecord; override;
    procedure OnAfterUpdateRecord; override;
    procedure OnLoadRecord; override;
    procedure OnChangeField(F: TField); override;
    procedure OnArgument(S: string); override;
    procedure OnClose; override;
    procedure OnCancelRecord; override;

    procedure OnDuplique(Sender: TObject);

{$IFDEF AFFAIRE}
    procedure OnTestDoc(Sender: TObject);
{$ENDIF AFFAIRE}
    procedure OnElaborationDoc(Sender: TObject);
    procedure OnPrendreDoc(Sender: TObject);
    procedure OnRendreDoc(Sender: TObject);
    procedure OnAnnulerPrendreDoc(Sender: TObject);
    //procedure OnChoixFichier   (Sender:TObject);

  private
    m_TypeDefaut: string;
    m_NatureDefaut: string;
    m_strDupFichier: string;
    m_strDupLibelle: string; // $$$ JP 29/11/05 - obligé pour avoir un libellé correct lors de la duplication
    m_strDupOrigine: string; // $$$ JP 30/11/05 - obligé de connaitre l'origine du fichier à dupliquer
    m_bInCreate: boolean;
    m_bInPrendre: boolean; // $$$ JP 13/12/05 - pour savoir si on peut enregistrer l'état "en cours de conception"
    NoForm : LongInt;
    TobAdr : Tob;

    function GetMaxNoDoc: string;
    //function  NewFileName (strExtension:string)      :string;
    function SelectFileName: string;
    //function  DefaultFileName                        :string;
    procedure EnableControls;
    //function  NomCompletFichier (strFichier:string)  :string;
    procedure DecomposeNom(var pstUtilisateur, pstFichier: String);
  end;

implementation
uses
  dicoBTP, //utofafdocselectfile,
{$IFDEF AFFAIRE}
{$IFNDEF EAGLCLIENT}
  utoflanceparser,
{$ENDIF}

{$ENDIF}
  UtilXls,
  utilGa,
  uyfilestd,
  shellapi,
  menus,
  dialogs // $$$ JP 28/11/05 - pour nouvelle gestion fichier en base
  ;
const
  // libellés des messages
  TexteMessage: array[1..21] of string = (
    {1}  'Sélectionnez le document à intégrer'
    {2}, 'Fichiers Microsoft Word (*.dot;*.dotx)|*.dot;*.dotx'
    {3}, 'Fichiers Microsoft Excel (*.xls;*.xlsx)|*.xls;*.xlsx'
    {4}, 'Est-ce un document spécifique au dossier en cours?'
    {5}, 'nouveau document'
    {6}, 'Un document doit avoir un libellé.'
    {7}, 'Le document n''est pas en cours de conception.'
    {8}, 'Le nom de fichier (issu du libellé) doit être valide.'
    {9}, 'Le fichier ''%s'' est absent#10 il ne peut pas être intégré en base.'
    {10}, 'Le fichier: ''%s'' est absent#10 Impossible de l''ouvrir en modification.'
    {11}, 'Veuillez associer un fichier au document (existant ou nouveau).'
    {12}, 'Confirmez-vous l''extraction de ''%s'' pour modification?'
    {13}, 'Prendre le document'
    {14}, 'Le document a été extrait pour modification dans:#10 %s #10Veillez à ne pas supprimer ce document ou ce répertoire, tant que ce fichier modifié n''aura pas été réintégré ou abandonné.'
    {15}, 'Confirmez-vous la validation des modifications apportées à ''%s'' ?'
    {16}, 'Etes-vous certain de valider ces modifications ?'
    {17}, 'Rendre le document'
    {18}, 'Vous n''êtes pas le correcteur actuel de ce document.#10En tant qu''administrateur, désirez-vous annuler les modifications réalisées actuellement par %s ?'
    {19}, 'Confirmez-vous l''annulation des modifications apportées à ''%s'' ?'
    {20}, 'Etes-vous certain d''annuler ces modifications ?'
    {21}, 'Veuillez fermer le document avant d''essayer d''intégrer les modifications'
     );

function TOM_AFDOCEXTERNE.GetMaxNoDoc: string;
var
  Q: TQuery;
begin
  Result := '0';
  Q := nil;
  try
    Q := OpenSQL('SELECT MAX(ADE_DOCEXCODE) AS NOMAX FROM AFDOCEXTERNE', TRUE);
    if not Q.eof then
      Result := IntToStr(Q.FindField('NOMAX').AsInteger + 1);
  finally
    Ferme(Q);
  end;
end;


function TOM_AFDOCEXTERNE.SelectFileName: string;
var
  //  strFichier  :string; //, strFicComplet  :string;
  strType: string;
  FileDialog: TOpenDialog; // $$$ JP 28/11/05 - sélection de fichier
  // $$$ JP 07/07/03   AppliPathStd, AppliPathDat   :string;
begin

  // Par défaut, nom de fichier vide (nouveau)
  Result := '';

  // Type de fichier à insérer (Word, Excel ...)
  strType := GetField('ADE_DOCEXTYPE');

  // Sélection de fichier (soit dans std, soit dans dat, soit dans rep. dossier)
  // $$$ JP 28/11/05 - sélection de n'importe quel fichier désormais, vu qu'il sera intégré en base à la validation
  FileDialog := TOpenDialog.Create(Ecran);
  try
    FileDialog.Title := 'Sélectionnez le document à intégrer';
    if strType = 'WOR' then
    begin
      FileDialog.Filter := 'Fichiers Microsoft Word (*.dot;*.dotx)|*.dot;*.dotx';
      FileDialog.DefaultExt := 'dotx';
    end
    else
      if strType = 'EXC' then
    begin
      FileDialog.Filter := TraduitGa(TexteMessage[3]);		//'Fichiers Microsoft Excel (*.xls)|*.xls';
      FileDialog.DefaultExt := 'xlsx';
    end;
    FileDialog.FilterIndex := 1;
    if FileDialog.Execute = TRUE then
      Result := FileDialog.FileName;
  finally
    FreeAndNil(FileDialog);
  end;

end;


procedure TOM_AFDOCEXTERNE.EnableControls;
var
  strOrigine: string;
  strType: string;
  strNature: string;
  strEtat: string;
  strFichier: string;
  strUser: string;
  bInModif: boolean;
  bNewVersion: boolean;
begin
  strOrigine := GetField('ADE_DOSSIER');
  strType := GetField('ADE_DOCEXTYPE');
  strNature := GetField('ADE_DOCEXNATURE');
  strEtat := GetField('ADE_DOCEXETAT');
  strFichier := GetField('ADE_FICHIER');
  bInModif := DS.State in [dsInsert, dsEdit];
  if (strFichier = '') or (strFichier[1] <> '@') then
    bNewVersion := FALSE
  else
  begin
    bNewVersion := TRUE;

    // Utilisateur modificateur
    if strEtat = 'ELA' then
    begin
      // BDU - 07/02/07 - FQ 13723 : Remplacer par un appel à la procédure DecomposeNom
      {
      strUser := Copy(strFichier, 2, 3);
      Delete(strFichier, 1, 5);
      }
      DecomposeNom(strUser, strFichier);
    end;
  end;

  // Bouton de gestion document (fichier)
  // $$$ JP 12/12/05 - plus de considération si en modif ou pas
  SetControlEnabled('MNPRENDRE', (bInModif = FALSE) and (strEtat <> 'ELA') and (strOrigine <> '$STD') and (bNewVersion = TRUE));
  SetControlEnabled('MNRENDRE', (bInModif = FALSE) and (strEtat = 'ELA') and (bNewVersion = TRUE) and (strUser = V_PGI.User));
  SetControlEnabled('MNANNULERPRENDRE', (bInModif = FALSE) and (strEtat = 'ELA') and (bNewVersion = TRUE) and ((strUser = V_PGI.User) or (V_PGI.Superviseur = TRUE)));
  SetControlEnabled('MNMODIFIER', (bInModif = FALSE) and (strEtat = 'ELA') and (bNewVersion = TRUE) and (strUser = V_PGI.User));
  SetControlEnabled('MNTESTER', (bInModif = FALSE) and (bNewVersion = TRUE) and (strType = 'WOR') and ((strNature = 'LMI') or (strNature = 'PMI')));

  // $$$ JP 12/12/05 - il faut cacher et non pas invalider les boutons standards (sinon "collision" avec agl/eagl)
  SetControlVisible('BVALIDER', (bInModif = TRUE) or (bNewVersion = TRUE));
  	//mcd 18/04/2006 12841 SetControlVisible('BDELETE', (bInModif = FALSE) and (strOrigine <> '$STD') and (bNewVersion = TRUE));
  SetControlVisible('BDELETE', (bInModif = FALSE) and (strOrigine <> '$STD'));
  SetControlEnabled('BDUPLIQUER', (bInModif = FALSE) and (strEtat <> 'ELA') and (bNewVersion = TRUE));

  SetControlEnabled('ADE_DOCEXCODE', FALSE);
  SetControlEnabled('ADE_LIBELLE', strOrigine <> '$STD');
  SetControlEnabled('ADE_DOCEXETAT', strEtat <> 'ELA');

  // Si document pour dossier uniquement, on affiche la mention (à coté du n° de doc)
  SetControlVisible('LDOSSPROP', strOrigine = V_PGI.NoDossier);
  SetControlVisible('LCEGIDPROP', strOrigine = '$STD');
end;

procedure TOM_AFDOCEXTERNE.OnNewRecord;
var
  strLibelle: string;
begin
  inherited;

  // BDU - 24/11/06, Initialise ces champs à vide
  SetControlText('HLUTILCREA', '');
  SetControlText('HLUTILMODIF', '');
  SetControlText('HLDATECREATION', '');
  SetControlText('HLDATEMODIF', '');

  m_bInCreate := TRUE;

  SetField('ADE_DOCEXCODE', StrToInt(GetMaxNoDoc()));
  SetField('ADE_DOCEXETAT', 'NUT');
  SetField('ADE_DOCEXTYPE', m_TypeDefaut);
  SetField('ADE_DOCEXNATURE', m_NatureDefaut);

  // Construction du nom de fichier en fonction si le doc est local au dossier, ou dans les standards
  if (not (ctxScot in V_PGI.PGIContexte) and ((ctxAffaire in V_PGI.PGIContexte) or (ctxGCAFF in V_PGI.PGIContexte))) or (m_NatureDefaut = 'GCO') or (m_NatureDefaut = 'MLC') or (m_NatureDefaut = 'MLF') then
    SetField('ADE_DOSSIER', '$DAT')
  else
  begin
				//Est-ce un document spécifique au dossier en cours?'
    if PgiAskAf(TexteMessage[4], Ecran.Caption) = mrYes then
      SetField('ADE_DOSSIER', V_PGI.NoDossier)
    else
      SetField('ADE_DOSSIER', '$DAT');
  end;

  // $$$ JP 28/11/05 - si nom de fichier pas encore défini (pas de duplication), on le sélectionne
  if m_strDupFichier = '' then
    m_strDupFichier := SelectFileName;
  // SetField('ADE_UTILCREA', V_PGI.UserName);
  SetField('ADE_UTILCREA', V_PGI.User);

  // Calcul du libellé par défaut: nom du fichier sans l'extension, sans le '@', sans le chemin complet
  if m_strDupFichier = '' then
    strLibelle := TraduitGa(TexteMessage[5])		//'nouveau document'
  else
    if m_strDupFichier[1] = '@' then
    strLibelle := m_strDupLibelle
  else
    strLibelle := ChangeFileExt(ExtractFileName(m_strDupFichier), '');
  SetField('ADE_LIBELLE', strLibelle);

  // Boutons et menus actifs ou non
  EnableControls;
  THEdit(GetControl('ADE_LIBELLE')).SetFocus;
end;

procedure TOM_AFDOCEXTERNE.OnDeleteRecord;
var
  Guid: String;
  strFichier: string;
  strReq: string;
  TOBFile: TOB;
begin
  inherited;

  // On récupère le FILEID à partir de la clé de YFILESTD
  strFichier := GetField('ADE_FICHIER');
  // BDU - 24/11/06, Pendant l'élaboration, le nom contenu dans
  // ADE_FICHIER est différent de celui contenu dans YFS_NOM
  if GetField('ADE_DOCEXETAT') = 'ELA' then
    strFichier := Format('@%s', [ExtractFileName(strFichier)]);
  if (strFichier <> '') and (strFichier[1] = '@') then
  begin
    Delete(strFichier, 1, 1);
    Guid := '';
    TOBFile := TOB.Create('le fileid', nil, -1);
    try
      // BDU - 24/11/06, Remplace GIGA par un appel à la fonction CodeProduitSelonContexte
      strReq := Format('SELECT YFS_FILEGUID FROM YFILESTD ' +
        'WHERE YFS_CODEPRODUIT = "%s" ' +
        'AND YFS_NOM = "%s" ' +
        'AND YFS_LANGUE = "FRA" ' +
        'AND YFS_CRIT1 = "%s%s"',
        [CodeProduitSelonContexte(GetField('ADE_DOCEXNATURE')), strFichier,
        GetField('ADE_DOCEXTYPE'), GetField('ADE_DOCEXNATURE')]);
      TOBFile.LoadDetailFromSQL(strReq);
      if TOBFile.Detail.Count > 0 then
        Guid := TOBFile.Detail[0].GetString('YFS_FILEGUID');
    finally
      FreeAndNil(TOBFile);
    end;

    // On supprime les enreg' des 3 tables ayant référence sur ce fileid
    if Guid<> '' then
    begin
      ExecuteSQL('DELETE FROM NFILES WHERE NFI_FILEGUID="' + Guid +'"');
      ExecuteSQL('DELETE FROM NFILEPARTS WHERE NFS_FILEGUID="' + Guid+'"');
      ExecuteSQL('DELETE FROM YFILESTD WHERE YFS_FILEGUID="' + Guid+'"');
    end;
  end;
end;

procedure TOM_AFDOCEXTERNE.OnUpdateRecord;
var
  strFichier,StrFichierOld: string;
  strExtension: string;
  ii : integer;
begin
  inherited;

  // Il faut au moins un libellé
  if GetField('ADE_LIBELLE') = '' then
  begin
    LastError := 1;
    LastErrorMsg := TraduitGa(TExteMessage[6]);		//'Un document doit avoir un libellé';
    SetFocusControl('ADE_LIBELLE');
    exit;
  end;

  // $$$ JP 13/12/05 - il ne faut pas pouvoir forcer le mode "en cours de conception", sauf si fichier extrait
  if (m_bInPrendre = FALSE) and (GetField('ADE_DOCEXETAT') = 'ELA') then
  begin
    LastError := 1;
    LastErrorMsg := TraduitGa(TexteMessage[7]);	//'Le document n''est pas en cours de conception';
    SetFocusControl('ADE_DOCEXETAT');
    exit;
  end;

  // Utilisteur modificateur
  // SetField('ADE_UTILMODIF', V_PGI.UserName);
  SetField('ADE_UTILMODIF', V_PGI.User);

  // Si on est en train de créer un enreg, il faut calculer le nom à donner au fichier en base
  if m_bInCreate = TRUE then
  begin
    // Extension de fichier pour le type
    strExtension := GetField('ADE_DOCEXTYPE');
    if strExtension = 'WOR' then
      strExtension := '.dotx'
    else
      if strExtension = 'EXC' then
      strExtension := '.xlsx'
    else
      strExtension := '';

    // $$$ JP 08/12/05 - il faut toujours donner au nom de fichier le libellé saisi
    strFichier := GetField('ADE_LIBELLE') + strExtension;

    // Il faut que le nom de fichier soit valide
    if IsValidFileName(strFichier) = FALSE then
    begin
      LastError := 1;
      LastErrorMsg := TraduitGa(TexteMessage[8]);		//'Le nom de fichier (issu du libellé) doit être valide';
      SetFocusControl('ADE_LIBELLE');
      exit;
    end;

    // On construit le nom du fichier pour éviter "collision" avec un fichier existant en base (!: nom en base peut être différent du nom sur disque)
		StrFichierOld:=StrFichier;
    // BDU - 24/11/06, Remplace GIGA par un appel à la fonction CodeProduitSelonContexte
    strFichier := FindFreeFileName(strFichier, GetField('ADE_DOCEXTYPE') +
      GetField('ADE_DOCEXNATURE'), CodeProduitSelonContexte(GetField('ADE_DOCEXNATURE')));
    if strFichier <> '' then
      begin
      SetField('ADE_FICHIER', '@' + strFichier);
				//mcd 26/04/2006 il faut traiter le cas ou le fichier est trop lon et a été renommer
      if Length (strFichierOld) > 35 then
       Begin
          //dans YYFilseSTD 35c seulement... on réduit le nom du fichier
          //A supprimer en V8 quand nom fichier sera permit sur plus de 35c
       ii := Pos(Copy(strfichier,1,10),m_strDupFichier);
       strExtension := ExtractFileExt (m_strDupFichier);
       // BDU - 24/11/06, En duplication, il ne faut pas modifier cette variable.
       // Par contre en "création pure" il faut renseigner la variable
       // avec le nom du fichier d'origine qui a été sélectionné par l'utilisateur
       if m_strDupFichier[1] <> '@' then
         m_strDupFichier:= Copy(m_strDupFichier,1,II-1) + Trim(Copy (StrFichier,1,28)) + StrExtension;
       end;
      end;
  end;
end;

procedure TOM_AFDOCEXTERNE.OnAfterUpdateRecord;
var
  strFileToImport: string;
  strFichier: string;
  strCrit1, std: string;
begin
  inherited;
       // Critère 1 du yfilestd = type + nature du document
  if m_bInCreate = TRUE then
  begin
    strCrit1 := GetField('ADE_DOCEXTYPE') + GetField('ADE_DOCEXNATURE');

    // Extraction du fichier à copier (soit duplication, soit création à partir du modèle vide toujours existant en base sous le nom "$NEWFILE")
    if m_strDupFichier = '' then
      // BDU - 24/11/06, Remplace GIGA par un appel à la fonction CodeProduitSelonContexte
      // Les nouveaux modèles sont nommés N + Nature du document (LMI, PMI ou ESO)
      // L'extension est DOC si la nature est LMI ou PMI, XLS si la nature est ESO
      AGL_YFILESTD_EXTRACT(strFileToImport, CodeProduitSelonContexte(GetField('ADE_DOCEXNATURE')),
        Format('N%s.%s', [GetField('ADE_DOCEXNATURE'),
        IIF(GetField('ADE_DOCEXNATURE') = 'ESO', 'XLS', 'DOC')]), strCrit1, 'NEW')
    else
      if m_strDupFichier[1] = '@' then
    begin
      Delete(m_strDupFichier, 1, 1);
      // BDU - 24/11/06, Remplace GIGA par un appel à la fonction CodeProduitSelonContexte
      AGL_YFILESTD_EXTRACT(strFileToImport, CodeProduitSelonContexte(GetField('ADE_DOCEXNATURE')), m_strDupFichier,
        strCrit1, '', '', '', '', FALSE, 'FRA', m_strDupOrigine);
    end
    else
      strFileToImport := m_strDupFichier;

    // On importe le fichier dans la base sous le nom calculé dans OnUpdateRecord (en enlevant le préfixe "@", bien sûr)
    if FileExists(strFileToImport) = TRUE then
    begin
      strFichier := GetField('ADE_FICHIER');
      Delete(strFichier, 1, 1);
      // BDU - 12/06/07. FQ 14248
      {
      if GetField('ADE_DOSSIER')  =V_PGI.Nodossier  then Std :='DOS' else std:='STD';
      }
      if GetField('ADE_DOSSIER') = '$DAT' then
        Std := 'STD'
      else
        Std := 'DOS';
      // BDU - 24/11/06, Remplace GIGA par un appel à la fonction CodeProduitSelonContexte
      AGL_YFILESTD_IMPORT(strFileToImport, CodeProduitSelonContexte(GetField('ADE_DOCEXNATURE')),
        strFichier, ExtractFileExt(strFileToImport), strCrit1, '', '', '', '', '-', '-', '-', '-',
        '-', 'FRA', std, GetField('ADE_LIBELLE'), GetField('ADE_DOSSIER')); // GA_20080313_BDU_GA15054. Ajout de GetField('ADE_DOSSIER')
    end
    else
    			//  Le fichier " %s " est absent'#10' il ne peut pas être intégré en base'
      PgiInfo(format(traduitga(TExteMessage[9]),[strFileToImport]),Ecran.caption);

    // Fin de création document (fichier importé en base)
    m_bInCreate := FALSE;
    m_strDupFichier := '';
    m_strDupLibelle := '';
    m_strDupOrigine := '';

    // On quitte, car souci de gestion enreg' AGL apparement
    Ecran.ModalResult := mrOK;
    //end;

    // Si nouvel enreg, on ferme la fenêtre
//     if m_bInCreate = TRUE then
//          Ecran.ModalResult := mrOK;
  end
  else
    m_bInPrendre := FALSE;

  // On replace les boutons & menus actifs ou non, en mode Browse
  EnableControls;
end;

procedure TOM_AFDOCEXTERNE.OnLoadRecord;
var
  TOBUser: TOB;
  strUser: string;
begin
  inherited;

  // $$$ JP 13/12/05 - il faut alimenter les zones dates/util création/modification
  SetControlText('HLDATECREATION', GetField('ADE_DATECREATION'));
  SetControlText('HLDATEMODIF', GetField('ADE_DATEMODIF'));

  TOBUser := TOB.Create('le user', nil, -1);
  try
    // Utilisateur créateur
    strUser := GetField('ADE_UTILCREA');
    TOBUser.LoadDetailFromSQL('SELECT US_LIBELLE FROM UTILISAT WHERE US_UTILISATEUR="' + strUser + '"');
    if TOBUser.Detail.Count > 0 then
      strUser := strUser + ' - ' + TOBUser.Detail[0].GetString('US_LIBELLE');
    SetControlText('HLUTILCREA', strUser);

    // Utilisateur modificateur
    strUser := GetField('ADE_UTILMODIF');
    TOBUser.LoadDetailFromSQL('SELECT US_LIBELLE FROM UTILISAT WHERE US_UTILISATEUR="' + strUser + '"');
    if TOBUser.Detail.Count > 0 then
      strUser := strUser + ' - ' + TOBUser.Detail[0].GetString('US_LIBELLE');
    SetControlText('HLUTILMODIF', strUser);
  finally
    FreeAndNil(TOBUser);
  end;

  // Autorisation des boutons ou pas
  EnableControls;
end;

procedure TOM_AFDOCEXTERNE.OnChangeField(F: TField);
begin
  inherited;
end;

procedure TOM_AFDOCEXTERNE.OnArgument(S: string);
var
  Critere: string;
{$IFNDEF GIGI}
  // ComboNat: THValComboBox; // remis GM car pbm compile à voir avec JP
{$ENDIF}

begin
  inherited;

  TToolBarButton97(GetControl('BDUPLIQUER')).OnClick := OnDuplique;

  // $$$ JP 28/11/05 - controle sur fichier
{$IFDEF AFFAIRE}
  TMenuItem(GetControl('MNTESTER')).OnClick := OnTestDoc;
{$ENDIF AFFAIRE}
  TMenuItem(GetControl('MNMODIFIER')).OnClick := OnElaborationDoc;
  TMenuItem(GetControl('MNPRENDRE')).OnClick := OnPrendreDoc;
  TMenuItem(GetControl('MNRENDRE')).OnClick := OnRendreDoc;
  TMenuItem(GetControl('MNANNULERPRENDRE')).OnClick := OnAnnulerPrendreDoc;

  m_bInCreate := FALSE;
  m_bInPrendre := FALSE;
  m_TypeDefaut := '';
  m_NatureDefaut := '';
  m_strDupFichier := '';
  m_strDupLibelle := '';
  m_strDupOrigine := '';
  SetControlEnabled('ADE_DOCEXTYPE', TRUE);
  SetControlEnabled('ADE_DOCEXNATURE', TRUE);
  Critere := (Trim(ReadTokenSt(S)));
  while (Critere <> '') do
  begin
    // Type document externe par defaut
    if (copy(Critere, 1, 4) = 'TYPE') then
    begin
      m_TypeDefaut := Copy(Critere, 6, Length(Critere) - 5);
      SetControlEnabled('ADE_DOCEXTYPE', FALSE);
    end;

    // Nature document externe par defaut
    if (copy(Critere, 1, 6) = 'NATURE') then
    begin
      m_NatureDefaut := Copy(Critere, 8, Length(Critere) - 7);
      SetControlEnabled('ADE_DOCEXNATURE', FALSE);
    end;

    if (copy(Critere, 1, 4) = 'FORM') then
      NoForm := StrToInt(Copy(Critere, 6, Length(Critere) - 5));

    if (copy(Critere, 1, 6) = 'TOBADR') then
      TobAdr := Tob(Valeuri(Copy(Critere, 8, Length(Critere) - 7)));

    // Paramètre suivant
    Critere := (Trim(ReadTokenSt(S)));
  end;

  // en GA, libellé dans les combos
  // $$$ JP 30/11/05 - fonctionne pas: en GI, on passe également par ce code!!!
{$IFNDEF GIGI}
  //if (ctxAffaire in V_PGI.PGIContexte) or (ctxGCAFF in V_PGI.PGIContexte)) then
  //begin
  // BDU - 24/11/06, Les libellés sont adaptés dans la tablette
  (*
  ComboNat := THValComboBox(Getcontrol('ADE_DOCEXNATURE'));
  ComboNat.Items[ComboNat.Values.IndexOf('LMI')] := 'Documents d''affaire';
  ComboNat.Items[ComboNat.Values.IndexOf('PMI')] := 'Proposition d''affaire';
  *)
  //end;
{$ENDIF}
  if (m_NatureDefaut = 'GCO') or (m_NatureDefaut = 'MLC') or (m_NatureDefaut = 'MLF') then TMenuItem(GetControl('MNTESTER')).Visible := False;
end;

procedure TOM_AFDOCEXTERNE.OnClose;
begin
  inherited;
end;

procedure TOM_AFDOCEXTERNE.OnCancelRecord;
begin
  inherited;
end;

procedure TOM_AFDOCEXTERNE.OnDuplique(Sender: TObject);
begin
  // Copie du fichier associé
  m_strDupFichier := GetField('ADE_FICHIER');
  m_strDupLibelle := GetField('ADE_LIBELLE'); // $$$ JP 29/11/05
  if GetField('ADE_DOSSIER') = '$STD' then
    m_strDupOrigine := 'CEG'
  else if GetField('ADE_DOSSIER') = '$DAT' then   //modif mcd 20/04/2006
    m_strDupOrigine := 'STD'
  else
    m_strDupOrigine := 'DOS';

  // Type et nature à dupliquer
  m_TypeDefaut := GetField('ADE_DOCEXTYPE');
  m_NatureDefaut := GetField('ADE_DOCEXNATURE');

  // nouvel enregistrement associé au fichier copié
  TFFiche(Ecran).Binsert.Click;
end;

procedure TOM_AFDOCEXTERNE.OnElaborationDoc(Sender: TObject);
var
  strFichier: string;
  strUser: string;
  F  : TFMul ;
  DSMul : THquery;
  stTitre : string;
begin

  // Code utilisateur modifiant le fichier, + le nom complet du fichier sur disque
  strFichier := GetField('ADE_FICHIER'); //F.Q.FindField ('ADE_FICHIER').AsString;
  if strFichier = '' then
    exit;

  // Code utilisateur & nom du fichier
  // BDU - 07/02/07 - FQ 13723 : Remplacer par un appel à la procédure DecomposeNom
  {
  Delete(strFichier, 1, 1);
  strUser := Copy(strFichier, 1, 3);
  Delete(strFichier, 1, 4);
  }
  DecomposeNom(strUser, strFichier);
  if strUser <> V_PGI.User then
    exit;

  // Ouverture du document
  // $$$ JP 28/11/05 - plus simple: on lance le fichier lui-même, windows se déme.....
  if FileExists(strFichier) = TRUE then
    begin
    if (m_NatureDefaut = 'MLC') or (m_NatureDefaut = 'MLF') then
      begin
        if TobAdr <> Nil then LancePublipostage('TOB',strFichier,'',TobAdr)
        else
          begin
          F := TFMul  (LongInt (NoForm)) ;
          stTitre :=  F.Q.titres ;
          DSMul:=THquery(F.Q) ;
          LancePublipostage('OPEN',strFichier,'',Nil,StTitre,DSMul,False);
          end;
      end
    else
      ShellExecute(0, pchar('open'), pchar(strFichier), nil, nil, SW_RESTORE)
    end
  else
    		//'Le fichier: " %s" est absent'#10' Impossible de l''ouvrir en modification'
    PgiInfo(format(traduitga(TexteMessage[10]),[strFichier]),Ecran.caption);
end;

{$IFDEF AFFAIRE}
procedure TOM_AFDOCEXTERNE.OnTestDoc(Sender: TObject);
var
  strChamps, strArg: string;
  strRech: string;
begin
  // Si pas de fichier associé, on fait rien
  if GetField('ADE_FICHIER') = '' then
  begin    //'Veuillez associer un fichier au document (existant ou nouveau)'
    PgiInfoAf(TexteMessage[11], Ecran.Caption);
    exit;
  end;

  // Sélection d'une affaire pour test lettre ou proposition de mission
  if GetField('ADE_DOCEXTYPE') = 'WOR' then
    if (GetField('ADE_DOCEXNATURE') = 'LMI') or (GetField('ADE_DOCEXNATURE') = 'PMI') then
    begin
      if GetField('ADE_DOCEXNATURE') = 'PMI' then
      begin
        strChamps := 'AFF_STATUTAFFAIRE=PRO';
        strArg := 'STATUT:PRO';
      end
      else
      begin
        strChamps := 'AFF_STATUTAFFAIRE=AFF';
        strArg := 'STATUT:AFF';
      end;
      strArg := strArg + ';NOCHANGESTATUT';

      // Recherche d'une affaire
      strRech := AGLLanceFiche('AFF', 'AFFAIRERECH_MUL', strChamps, '', strArg);
      if strRech <> '' then
      begin
        strArg := 'AFFAIRE=' + ReadTokenSt(strRech) + ';MODELE=' + IntToStr(GetField('ADE_DOCEXCODE')) + ';ETAT=ELA;TITRE=Test document;NATURE=' + GetField('ADE_DOCEXNATURE');
        AFLanceFiche_LanceParser('', strArg);
      end;
    end
    else
      if GetField('ADE_DOCEXTYPE') = 'EXC' then
      if (GetField('ADE_DOCEXNATURE') = 'ESO') then
      begin
        //$$$jp: todo
      end;
end;
{$ENDIF AFFAIRE}

procedure TOM_AFDOCEXTERNE.OnPrendreDoc(Sender: TObject);
var
  strFichier: string;
  strOrigine: string;
  F  : TFMul ;
  DSMul : THquery;
  stTitre : string;
begin
  strFichier := GetField('ADE_FICHIER');
  if (strFichier <> '') and (strFichier[1] = '@') then
  begin
    Delete(strFichier, 1, 1);
    			//'Confirmez-vous l''extraction de "%s" pour modification?', 'Prendre le document'
    if PgiAsk(format(traduitga(TexteMessage[12]),[GetField('ADE_LIBELLE')]),TexteMessage[13]) = mrYes then
    begin
      if GetField('ADE_DOSSIER') = '$STD' then
        strOrigine := 'CEG'
      else if GetField('ADE_DOSSIER') = '$DAT' then   //mcd 20/04/2006
        strOrigine := 'STD'
      else
        strOrigine := 'DOS';
      if (m_NatureDefaut = 'MLC') or (m_NatureDefaut = 'MLF') then
        begin
        strFichier := ExtractDocBaseFile(strFichier, GetField('ADE_DOCEXTYPE') + GetField('ADE_DOCEXNATURE'), strOrigine, False, GetField('ADE_DOCEXNATURE'));
        if TobAdr <> Nil then LancePublipostage('TOB',strFichier,'',TobAdr)
        else
          begin
          F := TFMul  (LongInt (NoForm)) ;
          DSMul:=THquery(F.Q) ;
          stTitre :=  F.Q.titres ;
          LancePublipostage('OPEN',strFichier,'',Nil,stTitre,DSMul,False);
          end;
        end
      else
        strFichier := ExtractDocBaseFile(strFichier, GetField('ADE_DOCEXTYPE') + GetField('ADE_DOCEXNATURE'), strOrigine, TRUE, GetField('ADE_DOCEXNATURE'));
      if strFichier <> '' then
      begin
        // Il faut avertir l'utilisateur où se trouve le document, pour qu'il évite de supprimer le rép' temporaire
        	//'Le document a été extrait pour modification dans:'#10' " #10' Veillez à ne pas supprimer ce document ou ce répertoire, tant que ce fichier modifié n''aura pas été réintégré ou abandonné'
        // BDU - 24/11/06, Pas de message en GA
        if ctxScot in V_PGI.PGIContexte then
          PgiInfo(format(traduitga(TexteMessage[14]),[strFichier]),Ecran.caption);

        // Mise à jour de l'enreg pour indiquer qu'il est pris
        if not (DS.State in [dsInsert, dsEdit]) then
          DS.Edit;
        SetField('ADE_FICHIER', '@' + V_PGI.User + '@' + strFichier);
        SetField('ADE_DOCEXETAT', 'ELA');
        m_bInPrendre := TRUE;
        TFFiche(Ecran).Bouge(nbPost); //TFFiche (Ecran).BValider.Click;

        EnableControls;
      end;
    end;
  end;
end;

procedure TOM_AFDOCEXTERNE.OnRendreDoc(Sender: TObject);
var
  strFichier: string;
  strUser, std, dossier: string;

  function TestFichierFermer(Nom: String): Boolean;
  var
    X: Integer;
  begin
    X := FileOpen(Nom, fmOpenWrite or fmShareExclusive);
    if X > 0 then
    begin
      Result := True;
      FileClose(X);
    end
    else
      Result := False;
  end;
  
begin
  strFichier := GetField('ADE_FICHIER');
  if (strFichier <> '') and (strFichier[1] = '@') then
  begin
    // Code utilisateur qui a pris le document - il faut être le même pour le rendre
    // BDU - 07/02/07 - FQ 13723 : Remplacer par un appel à la procédure DecomposeNom
    {
    Delete(strFichier, 1, 1);
    strUser := Copy(strFichier, 1, 3);
    }
    DecomposeNom(strUser, strFichier);
    if strUser <> V_PGI.User then
      exit;

    // Il faut aussi bien sûr que le fichier sur disque existe bien
    {
    Delete(strFichier, 1, 4);
    }
    if FileExists(strFichier) = FALSE then
      exit;

    if TestFichierFermer(strFichier) then
    begin
      // Après confirmation, on ré-importe le fichier dans la base
        //Confirmez-vous la validation des modifications apportées à %s  ?'
      if PgiAsk(format(traduitga(TexteMEssage[15]),[GetField('ADE_LIBELLE')]),TexteMessage[17]) = mrYes then
          //Etes-vous certain de valider ces modifications ?'
      if PgiAskAf(TexteMessage[16], TexteMessage[17]) = mrYes then
        begin
          // BDU - 12/06/07. FQ 14248
          {
          if GetField('ADE_DOSSIER')  =V_PGI.Nodossier  then Std :='DOS' else std:='STD';
          if Std ='DOS' then Dossier := V_PGI.Nodossier else Dossier :=''; // mcd 11/05/2006 12937 perte de l'info
          }
          if GetField('ADE_DOSSIER') = '$DAT' then
          begin
            Std := 'STD';
            Dossier := '';
          end
          else
          begin
            Std := 'DOS';
            Dossier := GetField('ADE_DOSSIER');
          end;
          // BDU - 30/01/07 - FQ : 13442 : Le test est déplacé avant les messages de confirmation
          // if TestFichierFermer(strFichier) then
          // begin
            // BDU - 24/11/06, Remplace GIGA par un appel à la fonction CodeProduitSelonContexte
          if AGL_YFILESTD_IMPORT(strFichier, CodeProduitSelonContexte(GetField('ADE_DOCEXNATURE')),
             ExtractFileName(strFichier), ExtractFileExt(strFichier), GetField('ADE_DOCEXTYPE') + GetField('ADE_DOCEXNATURE'),
             '', '', '', '', '-', '-', '-', '-', '-', 'FRA', std, GetField('ADE_LIBELLE'),dossier) = -1 then
            begin
              // Suppression copie du fichier en local
              DeleteFile(pchar(strFichier));

              // Màj de l'enregistrement pour indiquer que le document n'est plus en cours de modif'
              if not (DS.State in [dsInsert, dsEdit]) then
                DS.Edit;
              SetField('ADE_FICHIER', '@' + ExtractFileName(strFichier));
              SetField('ADE_DOCEXETAT', 'NUT');
              TFFiche(Ecran).Bouge(nbPost); //TFFiche (Ecran).BValider.Click;

              EnableControls; //Ecran.ModalResult := mrOK;
            end;
          // end
          // else
            // PgiBoxAf(TraduitGA(TexteMessage[21]), Ecran.Caption);
        end;
    end
    else
      PgiBoxAf(TraduitGA(TexteMessage[21]), Ecran.Caption);
  end;
end;

procedure TOM_AFDOCEXTERNE.OnAnnulerPrendreDoc(Sender: TObject);
var
  strFichier: string;
  strUser: string;
begin
  strFichier := GetField('ADE_FICHIER');
  if (strFichier <> '') and (strFichier[1] = '@') then
  begin
    // Code utilisateur qui a pris le document - il faut être le même pour le rendre, ou bien administrateur
    // BDU - 07/02/07 - FQ 13723 : Remplacer par un appel à la procédure DecomposeNom
    {
    Delete(strFichier, 1, 1);
    strUser := Copy(strFichier, 1, 3);
    }
    DecomposeNom(strUser, strFichier);
    if strUser <> V_PGI.User then
					//Vous n''êtes pas le correcteur actuel de ce document'#10' En tant qu''administrateur, désirez-vous annuler les modifications réalisées actuellement par
      // BDU - 07/02/07 - FQ 13723 : Remplacer mrYes par mrNo
      // if (V_PGI.Superviseur = FALSE) or (PgiAsk(format(traduitga(TexteMessage[18]),[strUser])) =mryes) then
      if (not V_PGI.Superviseur) or (PgiAsk(Format(TraduitGa(TexteMessage[18]),[strUser])) = mrNo) then
        Exit;

    // Pas nécessaire que le fichier soit encore présent sur disque, puisqu'on annule les modifications
    // BDU - 07/02/07 - FQ 13723 : Remplacer par un appel à la procédure DecomposeNom
    {
    Delete(strFichier, 1, 4);
    }

    // Après confirmation, on ré-importe le fichier dans la base
				//confirmez-vous l''annulation des modifications apportées à
    if PgiAsk(format(traduitga(TexteMessage[19]),[GetField('ADE_LIBELLE')]),Ecran.caption) = mrYes then
      if PgiAskAf(TexteMessage[20],Ecran.caption) = mrYes then
      begin
        // Suppression copie du fichier en local
        DeleteFile(pchar(strFichier));

        // Màj de l'enregistrement pour indiquer que le document n'est plus en cours de modif'
        if not (DS.State in [dsInsert, dsEdit]) then
          DS.Edit;
        SetField('ADE_FICHIER', '@' + ExtractFileName(strFichier));
        SetField('ADE_DOCEXETAT', 'NUT');
        TFFiche(Ecran).Bouge(nbPost); //TFFiche (Ecran).BValider.Click;

        EnableControls; //Ecran.ModalResult := mrOK;
      end;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BDU
Créé le ...... : 07/02/2007
Modifié le ... :   /  /
Description .. : Décompose le nom du fichier en :
Suite ........ :  - le nom de l'utilisateur entres les @
Suite ........ :  - le nom du fichier après le second @
Mots clefs ... :
*****************************************************************}
procedure TOM_AFDOCEXTERNE.DecomposeNom(var pstUtilisateur, pstFichier: String);
begin
  // pstFichier     =      @xxx@yyyy
  Delete(pstFichier, 1, 1);
  // pstFichier     =      xxx@yyyy
  pstUtilisateur := Copy(pstFichier, 1, Pos('@', pstFichier) - 1);
  // pstUtilisateur =      xxx
  Delete(pstFichier, 1, Pos('@', pstFichier));
  // pstFichier     =      yyyy
end;

initialization
  registerclasses([TOM_AFDOCEXTERNE]);
end.
