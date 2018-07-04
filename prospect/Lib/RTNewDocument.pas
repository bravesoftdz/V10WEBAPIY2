unit RTNewDocument;

interface

uses
  SysUtils, Classes, Controls, Forms, Dialogs,
  Vierge, StdCtrls, HRichOLE, Mask, Hctrls, ExtCtrls,
  HTB97,
{$IFDEF EAGLCLIENT}
  MaineAGL,
{$ELSE}
  FE_Main,
  {$IFNDEF DBXPRESS}dbtables{BDE},{$ELSE}uDbxDataSet,{$ENDIF}
{$ENDIF}
  wcommuns,

  UIUtil, UTob, HPanel, HMsgBox, HEnt1, AGLInitGC, UtilGC, TiersUtil, UtilPGI,
  ComCtrls, HRichEdt, HSysMenu
  ;

type
  TFRTNewDocument = class(TFVierge)
    PnlProprietes: THPanel;
    LblFichier: THLabel;
    LblDossier: THLabel;
    LblProprietes: THLabel;
    LblLibelleDoc: THLabel;
    LblEmplacement: THLabel;
    LblDateDeb: THLabel;
    LblNatDocument: THLabel;
    LblBlocnote: THLabel;
    LblMotsCles: THLabel;
    BDocument: TToolbarButton97;
    Fichier: THCritMaskEdit;
    YDO_LIBELLEDOC: THEdit;
    YDO_BLOCNOTE: THRichEditOLE;
    YDO_MOTSCLES: THMemo;
    RTD_DATERECEPTION: THCritMaskEdit;
    Dossier: THCritMaskEdit;
    YDO_NATDOC: THValComboBox;
    NomDossier: THCritMaskEdit;
    S_NUMCHAINAGE: THCritMaskEdit;
    LBLCHAINAGE: THLabel;
    LIBCHAINAGE: THCritMaskEdit;
    RTD_TABLELIBREGED1: THValComboBox;
    TRTD_TABLELIBREGED1: THLabel;
    RTD_TABLELIBREGED2: THValComboBox;
    TRTD_TABLELIBREGED2: THLabel;
    RTD_TABLELIBREGED3: THValComboBox;
    TRTD_TABLELIBREGED3: THLabel;
    S_NUMACTION: THCritMaskEdit;
    LBLACTION: THLabel;
    LIBACTION: THCritMaskEdit;
    S_PERSPECTIVE: THCritMaskEdit;
    LIBPERSPECTIVE: THCritMaskEdit;
    LBLPERSPECTIVE: THLabel;
    LblPrevenir: THLabel;
    RTD_RESSOURCE: THCritMaskEdit;
    NomRessource: THCritMaskEdit;
    BMail: TToolbarButton97;

    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
    procedure FormCreate(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure BValiderClick(Sender: TObject);
    procedure OnDocChange(Sender: TObject);
    procedure YDO_LIBELLEDOCExit(Sender: TObject);
    procedure YDO_MOTSCLESExit(Sender: TObject);
    procedure BDocumentClick(Sender: TObject);
    procedure DossierElipsisClick(Sender: TObject);
    procedure FichierElipsisClick(Sender: TObject);
    procedure DossierExit(Sender: TObject);
    procedure S_NUMCHAINAGEExit(Sender: TObject);
    procedure S_NUMACTIONElipsisClick(Sender: TObject);
    procedure S_NUMACTIONExit(Sender: TObject);
    procedure S_PERSPECTIVEExit(Sender: TObject);
    procedure RTD_RESSOURCEElipsisClick(Sender: TObject);
    procedure RTD_RESSOURCEExit(Sender: TObject);
    procedure BMailClick(Sender: TObject);
  private
    { Déclarations privées }
    SDocGUID, SDocGUIDASupprimer : String;

    TypeGed    : String;
    TobDocGed  : TOB;      // Tob de l'enregistrement DPDOCUMENT ou YMODELE (tables métier)
    TobDoc     : TOB;      // Tob de l'enregistrement YDOCUMENTS (gestion documentaire)
    SFileGUID     : String;  // Id si le fichier à archiver est déjà dans YFILES

    DefName    : String;   // Nom de fichier par défaut
    bDuringLoad, bChange, bDocChange_f : Boolean ; // Chgt dans les données
    Reference1,Infos,Objet  : String;
    NumChainage,NumAction,NumProposition : integer;
    ModifLien : Boolean;
    function  SauveEnregMsg : Boolean;
    procedure ActualiseTitre;
    function RechLibRef1(Ref1:string): string;
    function RechLibRef2(Rech,Ref2:string): string;
    function RechNumChainage : string;
    procedure EnvoiMail (MailAuto: string);
  public
    { Déclarations publiques }
    bModifie   : Boolean;
    NoDossier  : String;
    CodeGed    : String;
    bModele    : Boolean; // Gestion d'un modèle de doc
    bNoChangeCodeGed : Boolean; // $$$ JP 28/09/04 - pour interdir sur demande la modification dans la hiérarchie GED
  end;

  TParamGedDoc = record
    SDocGUID     : String;
    NoDossier : String;
    CodeGed   : String;
    SFileGUID : String;
    TypeGed   : String;
    DefName   : String;
    Objet     : String;
    Infos     : String;
    ModifLien : Boolean;
    
  end;


function ShowNewDocument(ParamGedDoc:TParamGedDoc; bInside:Boolean=False; bModele:Boolean=False; bNoChangeCodeGed:boolean=FALSE):string;
{                        iDocId : Integer ; sNoDossier, sCodeGed : String ;
                         iFileId : Integer=0 ; sTypeGed : String='DOC' ;
                         bInside : Boolean=False; sDefName: String='') : String;
}

////////// IMPLEMENTATION /////////////
implementation


uses
     {$IFDEF EWS}
     UtileWS,
     {$ENDIF}
     UtilDossier, UtilGed, UGedFiles, {galOutil,} ParamSoc;


function ShowNewDocument (ParamGedDoc:TParamGedDoc; bInside:Boolean=False; bModele:Boolean=False; bNoChangeCodeGed:boolean=FALSE):string;
// Retourne :
// - ##NONE## si l'enregistrement n'est pas validé
// - le NoDossier;CodeGed si validé (rq : le CodeGed peut être '')
var
  F : TFRTNewDocument;
  PP : THPanel;
  Critere,ChampMul,ValMul,stInfos : string;
  x : integer;
begin
  Result := '##NONE##';
  F := TFRTNewDocument.Create(Application);
  With ParamGedDoc do
    begin
    F.SDocGUID := SDocGUID;
    F.SFileGUID := SFileGUID;
    
    F.NoDossier := NoDossier;
    F.CodeGed := CodeGed;
    if TypeGed='' then
      F.TypeGed := 'DOC'
    else
      F.TypeGed := TypeGed;
    F.DefName := DefName;
    F.Infos := Infos;
    F.Objet := Objet;
    F.ModifLien := ModifLien;
    end;
  F.bModele := bModele;

  stInfos := F.Infos;
  Repeat
    Critere:=uppercase(Trim(ReadTokenSt(stInfos))) ;
    if Critere<>'' then
    begin
        x:=pos('=',Critere);
        if x<>0 then
        begin
           ChampMul:=copy(Critere,1,x-1);
           ValMul:=copy(Critere,x+1,length(Critere));
           if ChampMul='TIERS' then
           begin
              F.Reference1 := ValMul;
           end
           else if (ChampMul='CHAINAGE') then
                begin
                F.NumChainage := Valeuri(ValMul);
                end
           else if (ChampMul='ACTION') then
                begin
                F.NumAction := Valeuri(ValMul);
                end
           else if  (ChampMul='PROPOSITION') then
                begin
                F.NumProposition := Valeuri(ValMul);
                end;
        end;
    end;
  until  Critere='';

  PP := Nil;
  if bInside then
    begin
    // Ferme la précédente fiche éventuelle
    if PrepareInside then PP := FindInsidePanel;
    end;
  if PP<>Nil then
    begin
    InitInside(F,PP);
    F.Show;
    end
  else
    begin
    try
      // pour affichage des commentaires en inside, elle est sizeable, et là on veut pas
      F.BorderStyle      := bsDialog;
      F.bNoChangeCodeGed := bNoChangeCodeGed;
      F.ShowModal ;
    finally
      if F.bModifie then Result := F.NoDossier+';'+F.CodeGed;
      F.Free ;
      end;
    end ;
end;

{$R *.DFM}

procedure TFRTNewDocument.FormShow(Sender: TObject);
var Q : TQuery;
begin
  inherited;
  // Gestion d'un modèle de document : pas la même table métier
  if bModele then
    begin
    LblDossier.Visible := False;
    Dossier.Visible := False;
    LblMotsCles.Visible := False;
    YDO_MOTSCLES.Visible := False;
    LblNatDocument.Visible := False;
    YDO_NATDOC.Visible := False;
    LblBlocnote.Visible := False;
    YDO_BLOCNOTE.Visible := False;
    end;

  // Toujours un dossier de rattachement
  if NoDossier='' then NoDossier := '000000';

  TobDoc := Tob.Create('YDOCUMENTS', nil, -1);
  if bModele then TobDocGed := Tob.Create('YMODELES', nil, -1)
  else TobDocGed := Tob.Create('RTDOCUMENT', nil, -1);

  // Si enreg existant, renseigne la clé primaire
  if SDocGUID<>'' then
    begin
    if bModele then BDocument.Visible := true;
    if bModele then TobDocGed.PutValue('YMO_DOCGUID', SDocGUID)
    else TobDocGed.PutValue('RTD_DOCGUID', SDocGUID);
    TobDoc.PutValue('YDO_DOCGUID', SDocGUID);
    end;
  TobDocGed.LoadDB;
  TobDoc.LoadDB;

  // Nouveau document
  if SDocGUID='' then
    begin
    if Reference1 <> '' then Dossier.Text := Reference1;
    if NumChainage <> 0 then S_NumChainage.Text := IntToStr(NumChainage);
    if NumAction <> 0 then S_NumAction.Text := IntToStr(NumAction);
    if NumProposition <> 0 then S_Perspective.Text := IntToStr(NumProposition);
    if Objet = 'ACT' then
       begin
       S_NumChainage.Text := RechNumChainage;
       end;
    end
  // Document existant
  else
    begin
    TobDocGed.PutEcran(Self);
    TobDoc.PutEcran(Self);
    Dossier.Text := TobDocGed.GetValue('RTD_TIERS');
    if TobDocGed.GetValue('RTD_NUMCHAINAGE') <> 0 then S_NumChainage.Text := TobDocGed.GetValue('RTD_NUMCHAINAGE');
    if TobDocGed.GetValue('RTD_NUMACTION') <> 0 then S_NumAction.Text := TobDocGed.GetValue('RTD_NUMACTION');
    if TobDocGed.GetValue('RTD_PERSPECTIVE') <> 0 then S_PERSPECTIVE.Text := TobDocGed.GetValue('RTD_PERSPECTIVE');

    if SDocGUID<>'' then
      begin
      Q := OpenSQL('SELECT YDF_FILEGUID FROM YDOCFILES WHERE YDF_DOCGUID= "'+SDocGUID
                  +'" AND YDF_FILEROLE="PAL"', True,-1,'',true);
      if Not Q.Eof then SFileGUID := Q.FindField('YDF_FILEGUID').AsString;
      Ferme(Q);
      end;
    end;

  ActualiseTitre;

  if ModifLien = False then
     begin
{     if Dossier.Text <> '' then Dossier.enabled := False;
     if S_NumChainage.Text <> '' then S_NumChainage.enabled := False;
     if S_NumAction.Text <> '' then S_NumAction.enabled := False;
     if S_PERSPECTIVE.Text <> '' then S_PERSPECTIVE.enabled := False;  }
     Dossier.enabled := False;
     if (Objet = 'CHA') or (Objet = 'ACT') then S_NumChainage.enabled := False;
     if Objet = 'ACT' then S_NumAction.enabled := False;
     if Objet = 'PRO' then S_PERSPECTIVE.enabled := False;
     end;

  if Dossier.text <> '' then
    begin
    NomDossier.text := RechLibRef1(Dossier.text);
    S_NUMCHAINAGE.plus := 'RCH_PRODUITPGI="GRC" and RCH_TIERS="'+Dossier.text+'"';
    S_NUMACTION.plus := Dossier.text;
    S_PERSPECTIVE.plus := TiersAuxiliaire(Dossier.text, False);
    S_NUMCHAINAGE.DataType := 'RTCHAINAGE';
    S_PERSPECTIVE.DataType := 'RTPERSPECTIVES';
    end
  else
    begin
    S_NUMCHAINAGE.plus := '######';
    S_NUMACTION.plus := '######';
    S_PERSPECTIVE.plus := '######';
    end;
  if S_NUMCHAINAGE.text <> '' then LIBCHAINAGE.text := RechLibRef2('CHA',S_NUMCHAINAGE.text);
  if S_NUMACTION.text <> '' then LIBACTION.text := RechLibRef2('ACT',S_NUMACTION.text);
  if S_PERSPECTIVE.text <> '' then LIBPERSPECTIVE.text := RechLibRef2('PRO',S_PERSPECTIVE.text);
  if RTD_RESSOURCE.text <> '' then NOMRESSOURCE.text := RechLibRef2('RES',RTD_RESSOURCE.text);

  // Fichier déjà présent dans YFILES
  if SFileGUID<>'' then
    begin
    Fichier.Enabled := False;
    Fichier.ElipsisAutoHide := True;
    Fichier.Text := GetFileNameGed(SFileGUID);
    end;

  // Description par défaut du document
  if (SDocGUID='') and (YDO_LIBELLEDOC.Text='') then
    begin
    if DefName<>'' then YDO_LIBELLEDOC.Text := DefName
    else YDO_LIBELLEDOC.Text := Fichier.Text;
    end;

  // Retaillage
  if Not BDocument.Visible then Fichier.Width := YDO_LIBELLEDOC.Width;

  // Chargement terminé
  bDuringLoad := False;
  MajChampsLibresGED(TForm(Longint(PnlProprietes.owner)));
  if (SDocGUID='') and (Fichier.Text <> '') then bChange := True;

end;


procedure TFRTNewDocument.FormCreate(Sender: TObject);
begin
  inherited;
  bDuringLoad := True;
  bChange := False;
  bDocChange_f := False;
  bModifie := False;
end;


procedure TFRTNewDocument.FormClose(Sender: TObject;
  var Action: TCloseAction);
begin
  inherited;
  if TobDocGed<>Nil then TobDocGed.Free;
  if TobDoc<>Nil then TobDoc.Free;
end;


procedure TFRTNewDocument.FormCloseQuery(Sender: TObject;
  var CanClose: Boolean);
var rep: Integer;
begin
  inherited;
  // Changement demandé
  if bChange then
    begin
    if bDocChange_f then
       rep := PGIAskCancel('Le document lié a changé. Voulez-vous enregistrer les modifications ?', Self.Caption)
    else
       rep := PGIAskCancel('Voulez-vous enregistrer les modifications ?', Self.Caption);
    if rep=mrCancel then
      CanClose := False
    else if rep=mrNo then
      begin
      bChange := False;
      bDocChange_f := false;
      end
    else if rep=mrYes then
      CanClose := SauveEnregMsg;
    end;
end;


procedure TFRTNewDocument.BValiderClick(Sender: TObject);
begin
    YDO_LIBELLEDOC.SetFocus;  // Pr forcer la sortie du code tiers
  // Par défaut, le bouton BValider referme la fiche
  if Not SauveEnregMsg then
    // donc on l'en empêche si SauveEnregMsg n'a pas pu valider
    ModalResult := mrNone
  else
    inherited;
end;


procedure TFRTNewDocument.OnDocChange(Sender: TObject);
begin
  if not bDuringLoad then bChange := True;
end;


procedure TFRTNewDocument.YDO_LIBELLEDOCExit(Sender: TObject);
begin
  // Descriptif "minimum" du document
  if (YDO_LIBELLEDOC.Text='') and (Fichier.Text<>'') then
    YDO_LIBELLEDOC.Text := ChangeFileExt(ExtractFileName(Fichier.Text), '');
  ActualiseTitre;
end;


procedure TFRTNewDocument.YDO_MOTSCLESExit(Sender: TObject);
begin
  YDO_MOTSCLES.Text := UpperCase(YDO_MOTSCLES.Text);
end;

procedure TFRTNewDocument.BDocumentClick(Sender: TObject);
var
   sDoc_l : string;
begin
   // Extrait le document
//   sDoc_l := GetParamSocDp('SO_MDREPDOCUMENTS') + '\' + Fichier.text;
   sDoc_l := Fichier.text;
   V_GedFiles.Extract( sDoc_l, SFileGUID );
   if LanceAppliMaquette( sDoc_l, ExtractFileExt(sDoc_l) ) then
   begin
      // on sait pas ce qui a pu changer => suppression de l'ancien modèle
      SDocGUIDASupprimer := SDocGUID;
      // se considère comme en modification
      bChange := True;
      bDocChange_f := True;
   end;
end;


procedure TFRTNewDocument.DossierElipsisClick(Sender: TObject);
begin
  DispatchRecherche (Dossier, 2, '', '', '');
  if (Dossier.Text <> '') then
     begin
     NomDossier.text := RechLibRef1(Dossier.Text);
     Reference1 := Dossier.text;
     S_NUMCHAINAGE.plus := 'RCH_PRODUITPGI="GRC" and RCH_TIERS="'+Dossier.text+'"';
     S_NUMACTION.plus := Dossier.text;
     S_PERSPECTIVE.plus := TiersAuxiliaire(Dossier.text, False);
     S_NUMCHAINAGE.DataType := 'RTCHAINAGE';
     S_PERSPECTIVE.DataType := 'RTPERSPECTIVES';
     end;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : B.MERIAUX
Créé le ...... : 23/09/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TFRTNewDocument.FichierElipsisClick(Sender: TObject);
var
   ODModele_l : TOpenDialog;
begin
   //Open document a fusionner
   ODModele_l := TOpenDialog.Create(Application);
   try
   begin
      ODModele_l.Title := 'Sélectionner un fichier';
      ODModele_l.Filter := 'Tous les fichiers (*.*)|*.*';
      ODModele_l.InitialDir := GetParamsocSecur('SO_RTDIRSORTIE','');
      if ODModele_l.execute then
         Fichier.Text := ODModele_l.FileName;
   end;
   finally
      ODModele_l.Free;
   end;
end;


procedure TFRTNewDocument.ActualiseTitre;
var titre: String;
begin
  titre := '';
  // Type d'enregistrement
  if TypeGed='COM' then
    titre := 'Commentaire'
  else if bModele then
    titre := 'Modèle'
  else
    titre := TraduireMemoire('Document');
  // Nouvel enregistrement
  if SDocGUID ='' then titre := TraduireMemoire('Nouveau ') + LowerCase(titre) ;
  // Libellé détaillé
  if YDO_LIBELLEDOC.Text<>'' then
    titre := titre + ' - ' + YDO_LIBELLEDOC.Text;
  Self.Caption := titre;
  UpdateCaption(Self);
end;


function TFRTNewDocument.SauveEnregMsg: Boolean;
// Sauvegarde d'un nouveau document
// True si on a passé les tests de validité
var TobDocFiles : Tob;
    NewDocument : Boolean;
begin
  Result := False;
  NewDocument := False;

  // N'arrive jamais, le 000000 est par défaut
  if NoDossier='' then
    begin
    PGIInfo('Pas de dossier choisi : impossible d''enregistrer', Self.Caption);
    exit;
    end;
  // Contrôles généraux
  if RechLibRef1(Dossier.text) = '' then
    begin
    PGIInfo('Le code tiers n''existe pas', Self.Caption);
    Dossier.SetFocus;
    exit;
    end;
  if (S_NUMCHAINAGE.text <> '')  then
    begin
    if Not IsNumeric(S_NUMCHAINAGE.Text) or (RechLibRef2('CHA',S_NUMCHAINAGE.text) = '') then
      begin
      PGIInfo('Le chaînage n''existe pas ou ne concerne pas ce tiers', Self.Caption);
      S_NUMCHAINAGE.SetFocus;
      exit;
      end;
    end;
  if (S_NUMACTION.text <> '')  then
    begin
    if Not IsNumeric(S_NUMACTION.Text) or (RechLibRef2('ACT',S_NUMACTION.text) = '') then
      begin
      PGIInfo('L''action n''existe pas ou ne concerne pas ce tiers', Self.Caption);
      S_NUMACTION.SetFocus;
      exit;
      end;
    if RechNumChainage <> S_NUMCHAINAGE.text then
      begin
      PGIInfo('Le chaînage et l''action sont incompatibles', Self.Caption);
      S_NUMACTION.SetFocus;
      exit;
      end;
    end;
  if (S_PERSPECTIVE.text <> '')  then
    begin
    if Not IsNumeric(S_PERSPECTIVE.Text) or (RechLibRef2('PRO',S_PERSPECTIVE.text) = '') then
      begin
      PGIInfo('La proposition n''existe pas ou ne concerne pas ce tiers', Self.Caption);
      S_PERSPECTIVE.SetFocus;
      exit;
      end;
    end;
  if (RTD_RESSOURCE.text <> '')  then
    begin
    if RechLibRef2('RES',RTD_RESSOURCE.text) = '' then
      begin
      PGIInfo('La ressource n''existe pas', Self.Caption);
      RTD_RESSOURCE.SetFocus;
      exit;
      end;
    end;
  if YDO_LIBELLEDOC.Text='' then
    begin
    PGIInfo('Entrez le libellé descriptif du document', Self.Caption);
    YDO_LIBELLEDOC.SetFocus;
    exit;
    end;

  // En cas de modification du document WORD, ...
  if SDocGUIDASupprimer <> '' then
    begin
    // Dès qu'on est en phase de validation, on ne modifiera plus le document
    BDocument.Enabled := False;
    // Purge du modèle
    ExecuteSQL('DELETE FROM YMODELES WHERE YMO_DOCGUID='+SDocGUIDASupprimer);
    SupprimeDocument(SDocGUIDASupprimer);
    // On est maintenant sur un nouveau document
    SDocGUIDASupprimer := '';
    SDocGUID := '';
    // le fichier à insérer est comme un nouveau fichier
    SFileGUID := '';
//    Fichier.text := GetParamSocDp('SO_MDREPDOCUMENTS') + '\' + Fichier.text;
    Fichier.text := Fichier.text;
    // sinon ne validera pas la création si aucune modif autre que le document
    // et si la clé calculée retombe sur le même n° de clé (dernier docid)
    TobDoc.SetAllModifie(True);
    TobDocGed.SetAllModifie(True);
    end;

  // Contrôles avant validation d'un nouvel enreg
  if (SDocGUID='') and (SFileGUID='') and (TypeGed<>'COM') then
    BEGIN
    if Fichier.Text='' then
      begin
      PGIInfo('Aucun fichier à insérer.', Self.Caption);
      Fichier.SetFocus;
      exit;
      end;
    if Not FileExists(Fichier.Text) then
      begin
      PGIInfo(Format(TraduireMemoire('Le fichier %s n''existe plus.'),[Fichier.Text]), Self.Caption);
      Fichier.SetFocus;
      exit;
      end;
    END;

  // Nouvel enreg
  if SDocGUID='' then
    begin
    // Document non encore référencé
    if (SFileGUID='') and (TypeGed<>'COM') then
      begin
      // Enregistre le fichier dans YFILES, YFILEPARTS...
      SFileGUID := V_GedFiles.Import(Fichier.Text);
      if SFileGUID = '' then
        begin
        PGIInfo(Format(TraduireMemoire('Le fichier %s ne peut pas être importé dans la GED.'),[Fichier.Text]), TitreHalley);
        SFileGUID := '';
        exit;
        end;
      end;
    end;

  // Récup des modifs
  TobDocGed.GetEcran(Self);
  TobDoc.GetEcran(Self);

  TobDoc.PutValue('YDO_ANNEE',FormatDateTime('yyyy', StrToDateTime(RTD_DATERECEPTION.Text)));
  TobDoc.PutValue('YDO_MOIS',FormatDateTime('mm', StrToDateTime(RTD_DATERECEPTION.Text)));

  // Clé 2 : on remet à jour car a pu changer, même sur un enreg existant
  // #### pb, doit-on mettre à jour la référence documentaire YDO_DOCREF ?
  if Not bModele then
    begin
    TobDocGed.PutValue('RTD_TIERS', Dossier.text);
    TobDocGed.PutValue('RTD_NUMCHAINAGE', Valeuri(S_NUMCHAINAGE.text));
    TobDocGed.PutValue('RTD_NUMACTION', Valeuri(S_NUMACTION.text));
    TobDocGed.PutValue('RTD_PERSPECTIVE', Valeuri(S_PERSPECTIVE.text));
    end;

  // Nouvel enreg
  if SDocGUID='' then
    begin
    NewDocument := true;
    // Clé 1
    SDocGUID :=  AGLGetGuid; // NouvelleCle('YDO_DOCGUID', 'YDOCUMENTS');
    if bModele then TobDocGed.PutValue('YMO_DOCGUID', SDocGUID )
    else TobDocGed.PutValue('RTD_DOCGUID', SDocGUID );
    TobDoc.PutValue('YDO_DOCGUID', SDocGUID );

    // MD 31/05/06 : Ancienne clé mise à -2
    // pour éviter pb de maj des clés étrangères (comme dans ECRCOMPL)
    // qui prennent les enreg ydo_docid à 0 pour des enreg ged valides
    // lors de la conversion des integer en guid...
    TobDoc.PutValue('YDO_DOCID', -2);
    TobDocGed.PutValue('RTD_DOCID', -2);

    // #### Calcul d'une référence documentaire à revoir
    TobDoc.PutValue('YDO_DOCREF',SFileGUID + ' ' + SDocGUID );

    TobDoc.PutValue('YDO_AUTEUR',V_PGI.UserName );

    // #### pour l'instant, insère un seul fichier pour le document
    // avec comme rôle "PRINCIPAL"
    if (SFileGUID<>'') and (TypeGed<>'COM') then
      begin
      TobDocFiles := Tob.Create('YDOCFILES', Nil, -1);
      TobDocFiles.PutValue('YDF_DOCGUID', SDocGUID);
      TobDocFiles.PutValue('YDF_FILEGUID', SFileGUID);
      // MD 31/05/06
      TobDocFiles.PutValue('YDF_DOCID', -2);

      TobDocFiles.LoadDB;
      TobDocFiles.PutValue('YDF_FILEROLE', 'PAL');
      TobDocFiles.InsertOrUpdateDB;
      TobDocFiles.Free;
      // Un vrai nouveau document inséré, avec fichier associé
      end;
    end;

  Try
    bModifie := True; // Code retour pour la fonction appelante
    TobDoc.InsertOrUpdateDB;
    TobDocGed.InsertOrUpdateDB;
  except
    PGIInfo('Erreur lors de la sauvegarde des modifications.', Self.Caption);
  end;

  if GetParamSocSecur('SO_RTGEDENVOIAUTO',False) and (RTD_RESSOURCE.Text <> '') and (NewDocument) then EnvoiMail ('X');
  SDocGUIDASupprimer := '';

  bChange := False;
  bDocChange_f := False;
  Result := True;
end;

function TFRTNewDocument.RechLibRef1(Ref1:string): string;
var Q : TQuery;
begin
Q:=OpenSQL('SELECT T_LIBELLE FROM TIERS WHERE T_TIERS = "'+ Ref1 +'"', True,-1,'',true);
if not Q.Eof then  Result := Q.FindField('T_LIBELLE').asstring;
Ferme(Q) ;
end;

function TFRTNewDocument.RechLibRef2(Rech,Ref2:string): string;
var Q : TQuery;
begin
if Rech = 'CHA' then
   begin
   Q:=OpenSQL('SELECT RCH_LIBELLE FROM ACTIONSCHAINEES WHERE RCH_TIERS="'+Dossier.text+'" AND RCH_NUMERO = '+ Ref2 , True,-1,'',true);
   if not Q.Eof then  Result := Q.FindField('RCH_LIBELLE').asstring;
   Ferme(Q) ;
   end;
if Rech = 'ACT' then
   begin
   Q:=OpenSQL('SELECT RAC_LIBELLE FROM ACTIONS WHERE RAC_TIERS="'+Dossier.text+'" AND RAC_NUMACTION = '+ Ref2 , True,-1,'',true);
   if not Q.Eof then  Result := Q.FindField('RAC_LIBELLE').asstring;
   Ferme(Q) ;
   end;
if Rech = 'PRO' then
   begin
   Q:=OpenSQL('SELECT RPE_LIBELLE FROM PERSPECTIVES WHERE RPE_TIERS="'+Dossier.text+'" AND RPE_PERSPECTIVE = '+ Ref2 , True,-1,'',true);
   if not Q.Eof then  Result := Q.FindField('RPE_LIBELLE').asstring;
   Ferme(Q) ;
   end;
if Rech = 'RES' then
   begin
   Q:=OpenSQL('SELECT ARS_LIBELLE FROM RESSOURCE WHERE ARS_RESSOURCE="'+Ref2+'"', True,-1,'',true);
   if not Q.Eof then  Result := Q.FindField('ARS_LIBELLE').asstring;
   Ferme(Q) ;
   end;
end;

function TFRTNewDocument.RechNumChainage: string;
var Q : TQuery;
begin
   Q:=OpenSQL('SELECT RAC_NUMCHAINAGE FROM ACTIONS WHERE RAC_TIERS="'+Dossier.text+'" AND RAC_NUMACTION = '+ S_NUMACTION.Text , True,-1,'',true);
   if not Q.Eof then  Result := IntToStr(Q.FindField('RAC_NUMCHAINAGE').asinteger);
   if Valeuri(Result) = 0 then Result := ''; 
   Ferme(Q) ;
end;

procedure TFRTNewDocument.DossierExit(Sender: TObject);
begin
  inherited;
  NomDossier.Text := RechLibRef1 (Dossier.Text);
  S_NUMCHAINAGE.plus := 'RCH_PRODUITPGI="GRC" and RCH_TIERS="'+Dossier.text+'"';
  S_NUMACTION.plus := Dossier.text;
  S_PERSPECTIVE.plus := TiersAuxiliaire(Dossier.text, False);
  if Dossier.Text <> '' then
     begin
     S_NUMCHAINAGE.DataType := 'RTCHAINAGE';
     S_PERSPECTIVE.DataType := 'RTPERSPECTIVES'
     end
  else
     begin
     S_NUMCHAINAGE.DataType := '';
     S_PERSPECTIVE.DataType := '';
     end;
end;

procedure TFRTNewDocument.S_NUMCHAINAGEExit(Sender: TObject);
begin
  inherited;
  if Not IsNumeric (S_NUMCHAINAGE.Text) then
    begin
    LibChainage.Text := '';
    S_NUMCHAINAGE.Text := '';
    end
  else LibChainage.Text := RechLibRef2 ('CHA',S_NUMCHAINAGE.Text);
end;

procedure TFRTNewDocument.S_NUMACTIONElipsisClick(Sender: TObject);
var z_Tiers,z_action : string;
begin
  inherited;
  if Dossier.Text = '' then z_Tiers := '######' else z_Tiers := Dossier.Text;
  z_action:=AGLLanceFiche('RT','RTACTIONS_MUL_GED','RAC_TIERS='+z_Tiers+';RAC_NUMCHAINAGE='+S_NUMCHAINAGE.Text,'','') ;
  if z_action <> '' then S_NUMACTION.Text:=z_action;
end;

procedure TFRTNewDocument.S_NUMACTIONExit(Sender: TObject);
begin
  inherited;
  if Not IsNumeric (S_NUMACTION.Text) then LibAction.Text := ''
  else LibAction.Text := RechLibRef2 ('ACT',S_NUMACTION.Text);
  if LibAction.Text <> '' then
     begin
     if (S_NUMCHAINAGE.Enabled = True) and (S_NUMCHAINAGE.Text = '') then
        begin
        S_NUMCHAINAGE.Text := RechNumChainage;
        if S_NUMCHAINAGE.Text <> '' then LibChainage.Text := RechLibRef2 ('CHA',S_NUMCHAINAGE.Text);
        end;
     end;
end;

procedure TFRTNewDocument.S_PERSPECTIVEExit(Sender: TObject);
begin
  inherited;
  if Not IsNumeric (S_PERSPECTIVE.Text) then LibPerspective.Text := ''
  else LibPerspective.Text := RechLibRef2 ('PRO',S_PERSPECTIVE.Text);
end;

procedure TFRTNewDocument.RTD_RESSOURCEElipsisClick(Sender: TObject);
begin
  inherited;
  DispatchRecherche (RTD_RESSOURCE, 3, '', '', '');
  if (RTD_RESSOURCE.Text <> '') then
     begin
     NomRessource.text := RechLibRef2 ('RES',RTD_RESSOURCE.Text);
     end;
end;

procedure TFRTNewDocument.RTD_RESSOURCEExit(Sender: TObject);
begin
  inherited;
  NomRessource.text := RechLibRef2 ('RES',RTD_RESSOURCE.Text);
end;

procedure TFRTNewDocument.BMailClick(Sender: TObject);
begin
  inherited;
  if RTD_RESSOURCE.Text <> '' then EnvoiMail ('-');
end;

procedure TFRTNewDocument.EnvoiMail(MailAuto : string);
var Q : TQuery;
    Aqui,Objet : string;
    Liste :HTStringList;
begin
  inherited;
  NomRessource.SetFocus;  // Pr forcer la sortie du code ressource
  Q:=OpenSQL('SELECT ARS_EMAIL FROM RESSOURCE WHERE ARS_RESSOURCE="'+RTD_RESSOURCE.Text+'"', True,-1,'',true);
  if not Q.Eof then  Aqui := Q.FindField('ARS_EMAIL').asstring;
  Ferme(Q) ;
  if Aqui <> '' then
  begin
    Objet := TraduireMemoire('Document : ')+YDO_LIBELLEDOC.Text;
    Liste:=HTStringList.Create ;
    Liste.Add(TraduireMemoire('Date du document : ')+FormatDateTime ('dd/mm/yyyy',StrToDateTime(RTD_DATERECEPTION.Text)));
    Liste.Add(TraduireMemoire('Ce document concerne : '));
    Liste.Add(TraduireMemoire('     le client : ')+NomDossier.Text);
    if LibChainage.Text <> '' then Liste.Add(TraduireMemoire('     le chaînage : ')+LibChainage.Text);
    if LibAction.Text <> '' then Liste.Add(TraduireMemoire('     l''action : ')+LibAction.Text);
    if LibPerspective.Text <> '' then Liste.Add(TraduireMemoire('     la proposition : ')+LibPerspective.Text);
    PGIEnvoiMail (Objet,Aqui,'',Liste,'',(MailAuto='X'),1,'','');
    Liste.free;
  end
  else
    PGIInfo(TraduireMemoire('Envoi de message impossible : pas d''adresse e-mail'),'');
end;

end.
