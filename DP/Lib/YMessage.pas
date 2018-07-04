unit YMessage;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  Vierge, StdCtrls, ComCtrls, HRichEdt, HRichOLE, OleCtrls, SHDocVw_TLB,
  Menus, ImgList, ExtCtrls, Mask, Hctrls, HTB97,
  UTob, HMsgBox, Hent1, UGedFiles, HStatus,
{$IFDEF EAGLCLIENT}
  MaineAGL, MenuOLX, UtileAGL,
{$ELSE}
  {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet, {$ENDIF}
  FE_Main, MenuOlg, EdtREtat,
{$ENDIF}
  HSysMenu, HImgList;
                
const
     ccSupprimerMessage = 187400;
     
type

  TFMessage = class(TFVierge)
    BWebPrecedent: TToolbarButton97;
    PnlHaut: TPanel;
    Image1: TImage;
    TDateMessage: THLabel;
    DateMessage: TLabel;
    TYMS_SUJET: THLabel;
    TYMS_ALLDESTCC: THLabel;
    TYMS_TEL: THLabel;
    TYMS_LIBFROM: THLabel;
    TYMS_ALLDEST: THLabel;
    YMS_LIBFROM: TLabel;
    ALLDEST: THLabel;
    EnCopie: TLabel;
    ALLDESTCC: THLabel;
    TDateCreationMsg: THLabel;
    DateCreationMsg: TLabel;
    GrpDossier: TGroupBox;
    ANN_NOM2: TLabel;
    ANN_ALRUE1: TLabel;
    ANN_ALCP: TLabel;
    ANN_ALVILLE: TLabel;
    ANN_ALRUE2: TLabel;
    ANN_NOM1: TLabel;
    TANN_TEL1: TLabel;
    TANN_FAX: TLabel;
    ANN_FAX: TLabel;
    ANN_TEL1: TLabel;
    YMS_NODOSSIER: THCritMaskEdit;
    YMS_URGENT: TCheckBox;
    YMS_PRIVE: TCheckBox;
    YMS_SUJET: TEdit;
    YMS_TEL: TEdit;
    VSplitter: TSplitter;
    SaveDialog: TSaveDialog;
    ImageListFile: THImageList;
    ImageListLargeIcon: THImageList;
    FilePopupMenu: TPopupMenu;
    Ouvrir: TMenuItem;
    EnregistrerSous: TMenuItem;
    N1: TMenuItem;
    Visualiser: TMenuItem;
    PanelFiles: TPanel;
    PanelToolbar: TPanel;
    BOuvrirFichier: TToolbarButton97;
    BEnregistrerSous: TToolbarButton97;
    BVisualiserFichier: TToolbarButton97;
    ListViewFiles: TListView;
    Splitter1: TSplitter;        
    PageControl: TPageControl;
    TabMessage: TTabSheet;
    Web0: TWebBrowser_V1;
    TabText: TTabSheet;
    HroMessage: THRichEditOLE;
    BExtraireEml: TToolbarButton97;
    BTRANSFERER: TToolbarButton97;
    BREPONDRE: TToolbarButton97;
    BDeclasser: TToolbarButton97;
    BClasser: TToolbarButton97;
    BREPONDRETOUS: TToolbarButton97;
    BPrint: TToolbarButton97;
    procedure BPrintClick(Sender: TObject);
    procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    procedure FormShow(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
    procedure FormDestroy(Sender: TObject);
    procedure BWebPrecedentClick(Sender: TObject);
    procedure BDeleteClick(Sender: TObject);
    procedure OnMsgChange(Sender : TObject);
    procedure BValiderClick(Sender: TObject);
    procedure YMS_NODOSSIERElipsisClick(Sender: TObject);
    procedure YMS_NODOSSIERExit(Sender: TObject);
    procedure ANN_NOM1Click(Sender: TObject);
    procedure YMS_SUJETExit(Sender: TObject);
    procedure GrpDossierMouseMove(Sender: TObject; Shift: TShiftState; X, Y: Integer);
    procedure ANN_NOM1MouseMove(Sender: TObject; Shift: TShiftState; X, Y: Integer);
    procedure BClasserClick(Sender: TObject);
    procedure BDeclasserClick(Sender: TObject);
    procedure BREPONDREClick(Sender: TObject);
    procedure BREPONDRETOUSClick (Sender: TObject);
    procedure BTRANSFERERClick(Sender: TObject);
    procedure BExtraireEmlClick(Sender: TObject);
    procedure BOuvrirFichierClick(Sender: TObject);
    procedure BVisualiserFichierClick(Sender: TObject);
    procedure BEnregistrerSousClick(Sender: TObject);
    procedure PageControlChange(Sender: TObject);
    procedure DetailDestClick (Sender : TObject);

  private
    { Déclarations privées }
    FTempFileNames: TTempFileNames;
    FIconFiles: TIconFiles;
    FMailID  : Integer;
    FMsgGuid : String;

    TobListeDestA               : TOB;
    TobListeDestCc              : TOB;

    FBoiteDenvoi : Boolean;
    TobMsg : TOB; // pour la partie modifiable
    bDuringLoad, bMsgChange : Boolean;
    procedure AfficheCoordonneesDossier(guidper: String);
    function  SauveEnregMsg: Boolean;
    function  DonnerDateCreationMsg: String;

    // $$$ JP 16/06/05
{$IFDEF BUREAU}
    procedure OnTelClick (Sender:TObject);
{$ENDIF}

    procedure AddAssociateFile(FileGUID: String);
    procedure Navigate(WeBName: string; URL: string);
    procedure UpdatePageControl(Index: Integer = -1);

  public
    { Déclarations publiques }
    FNoDossier : String;

  end;

Type
  TMailFile = class(TObject)
  public
    FileGUID: String;
    FileName: string;
    FileIconID: Integer;
    TempFileName: string;
  end;

function ShowFicheMessage(MsgGuid: String; BoiteDenvoi: Boolean) : String;


////////// IMPLEMENTATION /////////////
implementation

uses
     UtilMessages, YNewDocument, YNewMessage, galOutil, UMailBox, EntDp,
     YMessageDetailDestinataire, AnnOutils;

function ShowFicheMessage(MsgGuid: String; BoiteDenvoi: Boolean) : String;
// Retourne NoDossier si on veut accéder au dossier client en sortie de fiche
var
  F : TFMessage;
begin
  // fiche utilisée uniquement pour des messages existants
  if MsgGuid='' then exit;

  F := TFMessage.Create(Application);
  F.FMsgGuid := MsgGuid;
  F.FBoiteDenvoi := BoiteDenvoi;
  try
    F.ShowModal ;
    Result := F.FNoDossier;
  finally
    F.Free ;
  end ;
end;

{$R *.DFM}

procedure TFMessage.FormShow(Sender: TObject);
var
    guidperdos         : String;
    tMsgFiles, tEnreg  : Tob;
    i                  : Integer;
    cMailFile          : TMailFile;
    sExtensionMsg, tmp : String;
    Q                  : TQuery;
begin
  inherited;

  //--- Récupération des infos en base
  TobMsg := Tob.Create('YMESSAGES', nil, -1);
  TobMsg.PutValue('YMS_MSGGUID', FMsgGuid);
  TobMsg.LoadDB;

  // Force "Message lu"
  if (Not FBoiteDenvoi) and (TobMsg.GetValue('YMS_LU')<>'X') then
    begin
    TobMsg.PutValue('YMS_LU', 'X');
    TobMsg.InsertOrUpdateDB;
    end;

  // Lien éventuel avec un mail d'origine dans YMAILS
  // (Rq : sera à 0 si ce n'est pas un message de type mail)
  FMailID := TobMsg.GetValue('YMS_MAILID');

  // Initialisations
  HroMessage.Text := '';
  Navigate('Web0', '');

  //--- Affichage des zones
  TobMsg.PutEcran(Self);

  //--- Onglets Message et Texte
  // Rq : zone mémo (HRicheditOle) non renseignée par putecran car pas même nom de zone...
  if TobMsg.GetValue('YMS_MEMO')<>'' then
    begin
    StringToRich (HroMessage, TobMsg.GetValue('YMS_MEMO'));
    TabText.TabVisible := True;
    end
  else if FMailId>0 then
    // on aura théoriquement au moins un contenu html dans ce cas (exemple : newsletters)
    TabText.TabVisible := False;

  //--- Zones calculées
  DateMessage.Caption     := FormatDateTime ('aaaa DD/MM/YYYY HH:NN:SS', TobMsg.GetValue ('YMS_DATEMODIF')); // $$$ JP 16/01/07: FormatDateTime ('aaaa DD/MM/YYYY HH:NN:SS', TobMsg.GetValue('YMS_DATE'));
  DateCreationMsg.Caption := DonnerDateCreationMsg;
  if FBoiteDenvoi then
    begin
    TDateMessage.Caption := 'Emis le';
    // Masqué, car identique à DateMessage
    DateCreationMsg.Visible := False;
    TDateCreationMsg.Visible := False;
    end;

  EnCopie.Visible := (TobMsg.GetValue('YMS_ENCOPIE')='X');

  //--- Titre du message = type + sujet
  Self.Caption := RechDom ('YYMSGTYPE', TobMsg.GetValue('YMS_MSGTYPE'),False)
  + ' - ' + TobMsg.GetValue('YMS_SUJET');

  //--- Dossier associé
  guidperdos := GetGuidPer (TobMsg.GetValue('YMS_NODOSSIER'));
  AfficheCoordonneesDossier (guidperdos);

  //--- Titre des onglets
  // si on ouvre un msg de la boite d'envoi, le vrai texte envoyé est dans l'onglet texte,
  // pas dans l'onglet message (tant qu'on n'a pas d'éditeur html pour modifier directement
  // le msg), donc on affiche en priorité l'onglet Texte
  if FBoiteDenvoi and TabMessage.TabVisible and TabText.TabVisible then
    begin
    TabMessage.Caption := 'Message initial';
    TabText.Caption := 'Texte envoyé';
    PageControl.ActivePage := TabText;
    end;

  //--- Liste des destinataires
  TobListeDestA:=Tob.Create ('',nil,-1);
  TobListeDestCc:=Tob.Create ('',nil,-1);

  InitialiserTobListeDestinataire (TobListeDestA,DonnerListeDestinataireFromSql (TobMsg.GetValue ('YMS_MSGGUID'),'A',TobMsg.GetValue('YMS_ALLDEST')));
  InitialiserTobListeDestinataire (TobListeDestCc,DonnerListeDestinataireFromSql (TobMsg.GetValue ('YMS_MSGGUID'),'Cc',TobMsg.GetValue('YMS_ALLDESTCc')));

  AllDest.Caption:=DonnerListeDestinataireFromTob (TobListeDestA,False);
  AllDestCc.Caption:=DonnerListeDestinataireFromTob (TobListeDestCc,False);

  ALLDEST.Hint:='Double-cliquer pour agrandir';
  ALLDESTCC.Hint:='Double-cliquer pour agrandir';

  //--- Bouton de Classement
  if FBoiteDenvoi then
    begin
    // MD 09/06/06 - Classement autorisé dans la boite d'envoi
    // BClasser.Visible := False;
    // BDeclasser.Visible := False;
    end
  else
    begin
    BClasser.Visible := (TobMsg.GetValue('YMS_TRAITE')<>'X');
    BDeclasser.Visible := (TobMsg.GetValue('YMS_TRAITE')='X');
    // Outil de debug : extraction du fichier eml complet
    if V_PGI.SAV and (FMailId>0) then BExtraireEml.Visible := True;
    end;

  //--- Pièces jointes + contenu eml + contenu html

  // Nb d'enreg avec ATTACHED='-' : pour éliminer l'enreg eml si on a l'html
  sExtensionMsg := '';
  Q := OpenSQL('SELECT YFI_FILENAME FROM YMSGFILES, YFILES WHERE YMG_MSGGUID="'+FMsgGuid+'" AND YMG_FILEGUID=YFI_FILEGUID AND YMG_ATTACHED="-"', True, -1, '', True);
  While Not Q.Eof do
    begin
    tmp := ExtractFileExt(Q.FindField('YFI_FILENAME').AsString);
    if (tmp<>'.eml') or (sExtensionMsg='') then sExtensionMsg := tmp;
    Q.Next;
    end;
  Ferme(Q);

  // Lecture des enreg
  tMsgFiles := TOB.Create('Fichiers attachés', nil, -1);
  tMsgFiles.LoadDetailFromSQL('SELECT YFI_FILEGUID, YFI_FILENAME, YMG_ATTACHED FROM YMSGFILES, YFILES'
   + ' WHERE YMG_MSGGUID="'+FMsgGuid+'" AND YMG_FILEGUID=YFI_FILEGUID');
  for i:=0 to tMsgFiles.Detail.Count-1 do
    begin
    tEnreg := Tob(tMsgFiles.Detail[i]);

    // Contenu html
    if (tEnreg.GetValue('YMG_ATTACHED') = BoolToSTR(False)) then
    begin
      if (TObject(PageControl.Pages[0].Tag) = nil)
      and (ExtractFileExt(tEnreg.GetValue('YFI_FILENAME'))=sExtensionMsg) then
      begin
        cMailFile := TMailFile.Create;
        cMailFile.FileGUID := tEnreg.GetValue('YFI_FILEGUID');
        cMailFile.FileName := tEnreg.GetValue('YFI_FILENAME');
        cMailFile.FileIconID := -1;
        // cMailFile.TempFileName := V_GedFiles.TempFileName(V_GedFiles.FileExtension(tEnreg.GetValue('YFI_FILENAME')));
        cMailFile.TempFileName := V_GedFiles.TempFileName(ExtractFileExt(tEnreg.GetValue('YFI_FILENAME')));

        V_GedFiles.Extract(cMailFile.TempFileName, cMailFile.FileGUID);
        FTempFileNames.Add(cMailFile.TempFileName);
        PageControl.Pages[0].Tag := Integer(cMailFile);
      end;
    end
    // pièce jointe classique
    else
      AddAssociateFile(tEnreg.GetValue('YFI_FILEGUID'));
    end;

  tMsgFiles.Free;

  // Affichage du contenu html dans l'onglet Message
  if TObject(PageControl.Pages[0].Tag) is TMailFile then
    UpdatePageControl(0)
  else
    PageControl.Pages[0].TabVisible := False; // ou TabMessage.TabVisible := False !

  PageControlChange(Self);

  // Affichage du panel des pièces jointes si nécessaire
  if ListViewFiles.Items.Count > 0 then
    begin
    PanelFiles.Visible := True;
    VSplitter.Visible := True;
    end;

  bDuringLoad := False;
end;


procedure TFMessage.AfficheCoordonneesDossier(guidper: String);
var T: Tob;
begin
  if (guidper='') then
    begin
    ANN_NOM1.Caption := '';
    ANN_NOM2.Caption := '';
    ANN_ALRUE1.Caption := '';
    ANN_ALRUE2.Caption := '';
    ANN_ALCP.Caption := '';
    ANN_ALVILLE.Caption := '';
    ANN_TEL1.Caption := '';
    ANN_FAX.Caption := '';
    TANN_TEL1.Visible := False;
    TANN_FAX.Visible := False;
    end
  else
    begin
    TANN_TEL1.Visible := True;
    TANN_FAX.Visible := True;
    T := Tob.Create ('La personne du dossier', Nil, -1);
    T.LoadDetailFromSQL ('SELECT ANN_NOM1, ANN_NOM2, ANN_ALRUE1, ANN_ALRUE2, ANN_ALCP,'
    +' ANN_ALVILLE, ANN_TEL1, ANN_FAX FROM ANNUAIRE WHERE ANN_GUIDPER="'+guidper+'"');
    // 1er enreg trouvé
    if T.Detail.Count>0 then Tob(T.Detail[0]).PutEcran(Self);
    T.Free;
    end;

    // $$$ JP 16/06/05
{$IFDEF BUREAU}
    if VH_DP.ctiAlerte <> nil then
    begin
         ANN_TEL1.OnClick := OnTelClick;

         // $$$ JP 16/01/07: voir qu'on peut lancer un appel sur le n° de tel dossier
         ANN_TEL1.Cursor := crHandPoint; // OnClick := OnTelClick;
         ANN_TEL1.Hint   := 'Appeler';
    end;
{$ENDIF}
end;

{$IFDEF BUREAU}
procedure TFMessage.OnTelClick (Sender:TObject);
begin
     if PgiAsk ('Appeler le ' + ANN_TEL1.Caption + ' ?') = mrYes then
        VH_DP.ctiAlerte.MakeCall (ANN_TEL1.Caption);
end;
{$ENDIF}

function TFMessage.SauveEnregMsg:Boolean;
var
  guidperdos: string;
begin
  Result := False;
  // peu de zones modifiables
  // TobMsg.GetEcran (Self);
  TobMsg.GetEcranChamp ('YMS_URGENT', YMS_URGENT);
  TobMsg.GetEcranChamp ('YMS_PRIVE', YMS_PRIVE);
  TobMsg.GetEcranChamp ('YMS_NODOSSIER', YMS_NODOSSIER);
  TobMsg.GetEcranChamp ('YMS_SUJET', YMS_SUJET);
  // #### Pour le cas où le contenu du message serait modifiable ...
  // TobMsg.PutValue ('YMS_MEMO', RichToString(HroMessage));

  guidperdos := GetGuidPer (YMS_NODOSSIER.Text);
  if (guidperdos='') And (YMS_NODOSSIER.Text<>'') then
      PGIInfo ('Le N° de dossier est incorrect, veuillez le modifier.', YMS_SUJET.Text)
  else
  begin
      Result := True;
      TobMsg.InsertOrUpdateDB;
      bMsgChange := False;
  end;

  AfficheCoordonneesDossier (guidperdos);
end;

// $$$ JP 16/01/07: désormais, YMS_DATE est correct (sauf heure, fuseau horaire pas encore géré)
function TFMessage.DonnerDateCreationMsg:string;
begin
     Result := '';

     if FMsgGuid<>'' then
         // Mail existant
         Result := FormatDateTime ('aaaa DD/MM/YYYY HH:NN:SS', TobMsg.GetValue('YMS_DATE'))
     else
         // Mail en cours de création
         Result := FormatDateTime ('aaaa DD/MM/YYYY HH:NN:SS', Now);
end;
{function TFMessage.DonnerDateCreationMsg: String;
begin
  Result := '';
  // Mail existant
  if FMsgGuid<>'' then
    begin
    // Mail issu de l'extérieur
    if TobMsg.GetString ('YMS_STRDATE')<>'' then
      Result := DecodeStrDate (TobMsg.GetString ('YMS_STRDATE'))
    else
      // (sera masqué en cas de boite d'envoi)
      Result := FormatDateTime ('aaaa DD/MM/YYYY HH:NN:SS', TobMsg.GetValue('YMS_DATE'));
    end
  // Mail en cours de création
  else
    Result := FormatDateTime ('aaaa DD/MM/YYYY HH:NN:SS', Now);
end;}

procedure TFMessage.BPrintClick(Sender: TObject);
begin
  inherited;

  LanceEtat ('E', 'SYS', 'MAI', TRUE, FALSE, FALSE, nil, 'YMS_MSGGUID="' + TobMsg.GetValue('YMS_MSGGUID')+'"', TobMsg.GetValue ('YMS_SUJET'), FALSE);

end;

procedure TFMessage.FormKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  case Key of

  // F10 : Envoyer
  VK_F10: if (BValider.Visible) and (BValider.Enabled) then
            begin key := 0; BValider.Click; end;

  // Ctrl + P : imprimer
  80:     if (Shift = [ssCtrl]) and (BPrint.Visible) and (BPrint.Enabled) then
            begin key := 0; BPrint.Click; end;

  // Ctrl + R : répondre
  82:     if (Shift = [ssCtrl]) and (BREPONDRE.Visible) and (BREPONDRE.Enabled) then
            begin key := 0; BREPONDRE.Click; end;

  // Ctrl + suppr : supprimer
  VK_DELETE :
          if (Shift = [ssCtrl]) and (BDelete.Visible) and (BDelete.Enabled) then
            begin key := 0; BDelete.Click; end;

  else
    inherited;
  end;
end;


procedure TFMessage.AddAssociateFile (FileGUID: String);
var
  tobFiles: Tob;
  cMailFile: TMailFile;
  ListItem: TListItem;
begin
  tobFiles := Tob.Create('_YFILES_', nil, -1);
  tobFiles.LoadDetailFromSQL('SELECT YFI_FILEGUID, YFI_FILENAME FROM YFILES WHERE YFI_FILEGUID = "' + FileGUID + '"');

  if tobFiles.Detail.Count = 0 then
  begin
    tobFiles.Free;
    Exit;
  end;

  cMailFile := TMailFile.Create;
  cMailFile.FileGUID := FileGUID;
  cMailFile.FileName := tobFiles.Detail[0].GetValue('YFI_FILENAME');
  cMailFile.FileIconID := FIconFiles.AddIcons(ExtractFileExt(cMailFile.FileName));

  ListItem := ListViewFiles.Items.Add;
  ListItem.Caption := cMailFile.FileName;
  ListItem.ImageIndex := FIconFiles[cMailFile.FileIconID].LargeID;
  ListItem.Data := cMailFile;

  tobFiles.Free;
end;

procedure TFMessage.FormCreate(Sender: TObject);
begin
  inherited;

  bDuringLoad := True;
  bMsgChange := False;

  FMailId := 0;
  FTempFileNames := TTempFileNames.Create;
  FIconFiles := TIconFiles.Create;
  FIconFiles.SmallImageList := ImageListFile;
  FIconFiles.LargeImageList := ImageListLargeIcon;
  FMsgGuid := '';

  //PGR 01/09/2005
  VStatus.DisablePaint := True;

  // $$$ JP 14/12/06
  if JaiLeDroitConceptBureau (ccSupprimerMessage) = FALSE then
     BDelete.Visible := FALSE;
end;

procedure TFMessage.FormCloseQuery(Sender: TObject; var CanClose: Boolean);
var
    rep: Integer;
begin
  // Changement demandé
  if bMsgChange then
    begin
    rep := PGIAskCancel ('Voulez-vous enregistrer les modifications ?', Self.Caption);
    if rep=mrCancel then
      CanClose := False
    else if rep=mrNo then
      bMsgChange := False
    else if rep=mrYes then
      begin
          if not SauveEnregMsg then
            CanClose := False;
      end;
    end;
  // Le code retour ne doit plus être renseigné
  if Not CanClose then FNoDossier := '';

  inherited;
end;

procedure TFMessage.FormDestroy(Sender: TObject);
var
  iFileAttached, iCtrl: Integer;
begin
  for iCtrl := 0 to Self.ComponentCount - 1 do
  begin
    if Self.Components[iCtrl] is TWebBrowser_V1 then
      Navigate (Self.Components[iCtrl].Name, '');
  end;

  Application.ProcessMessages;

  if TobListeDestA<>nil then TobListeDestA.Free;
  if TobListeDestCc<>nil then TobListeDestCc.Free;

  if TobMsg<>Nil then TobMsg.Free;
  FTempFileNames.Free;
  FIconFiles.Free;
  TMailFile(PageControl.Pages[0].Tag).Free;

  for iFileAttached := 0 to ListViewFiles.Items.Count - 1 do
    TMailFile(ListViewFiles.Items[iFileAttached].Data).Free;

  //PGR 01/09/2005
  VStatus.DisablePaint := False;

  inherited;
end;

procedure TFMessage.BWebPrecedentClick(Sender: TObject);
var
  Web: TWebBrowser_V1;
begin
  inherited;
  Web := TWebBrowser_V1(Self.FindComponent('Web' + IntToStr(PageControl.ActivePageIndex)));

  if Web = nil then Exit;

  try
    Web.GoBack;
  except
  end;
end;

procedure TFMessage.BDeleteClick(Sender: TObject);
var supprok : Boolean;
    sUserMail : String;
begin
  // SauveEnregMsg; => pas utile (permet si on ne supprime pas, d'annuler les dern. modifs)
  if PGIAsk('Voulez-vous vraiment supprimer le message ?', Self.Caption)<>mrYes then exit;
  // Clé pour recherche YMAILS associé
  sUserMail := '';
  if FMailId<>0 then sUserMail := TobMsg.GetValue('YMS_USERMAIL');

  // Delete table principale
  supprok := True;
  if FMsgGuid<>'' then
    supprok := DeleteMessage (FMsgGuid);

  // Delete YMAILS et dépendances
  if (supprok) and (FMailId<>0) then
    // on ne doit pas passer V_PGI.User car parfois BAL globale
    UserMailDelete (FMailId, sUserMail);

  ModalResult := mrOk; // ou Self.Close;
end;

procedure TFMessage.OnMsgChange(Sender: TObject);
begin
  if not bDuringLoad then bMsgChange := True;
end;

procedure TFMessage.BValiderClick(Sender: TObject);
begin
  // Herite seulement si SauveEnregMsg réussit car sinon on repasse dans YMS_NODOSSIERExit
  if not SauveEnregMsg then
    ModalResult := mrNone
  else
    inherited;
end;

procedure TFMessage.YMS_NODOSSIERElipsisClick(Sender: TObject);
var St, codper : String;
begin
  inherited;
   // retourne NoDossier;CodePer;Nom1
   St := AGLLanceFiche ('YY','YYDOSSIER_SEL', '','',YMS_NODOSSIER.Text);
   if St<>'' then
     begin
     YMS_NODOSSIER.Text := READTOKENST (St);
     codper := READTOKENST (St);
     AfficheCoordonneesDossier (codper);
     end;
end;

procedure TFMessage.YMS_NODOSSIERExit(Sender: TObject);
var
  guidperdos : String;
begin
  inherited;
  guidperdos := GetGuidPer (YMS_NODOSSIER.Text);
  if (guidperdos='') And (YMS_NODOSSIER.Text<>'') then
  begin
    PGIInfo ('Le N° de dossier est incorrect, veuillez le modifier.', YMS_SUJET.Text);
    YMS_NODOSSIER.Setfocus;
  end;
  AfficheCoordonneesDossier (guidperdos);
end;

Procedure TFMessage.DetailDestClick (Sender : TObject);
begin
 AfficherDetailDestinataire (TobMsg.GetValue('YMS_MSGGUID'),TobListeDestA,TobListeDestCc);
end;

procedure TFMessage.ANN_NOM1Click(Sender: TObject);
begin
  inherited;
  // No de dossier à accéder en sortie de fiche
  FNoDossier := YMS_NODOSSIER.Text;
  if FNoDossier='' then exit;

  ModalResult := mrOk; // ou Self.Close;
end;

procedure TFMessage.YMS_SUJETExit(Sender: TObject);
begin
  inherited;
  // Titre du message = type + sujet
  Self.Caption := RechDom ('YYMSGTYPE', TobMsg.GetValue('YMS_MSGTYPE'),False) + ' - ' + YMS_SUJET.Text;
end;

procedure TFMessage.GrpDossierMouseMove(Sender: TObject;
  Shift: TShiftState; X, Y: Integer);
begin
  inherited;
  // dégrise le lien hypertexte
  ANN_NOM1.Color := clBtnFace;
end;

procedure TFMessage.ANN_NOM1MouseMove(Sender: TObject; Shift: TShiftState;
  X, Y: Integer);
begin
  inherited;
  // lorsqu'on passe sur la zone, elle apparait grisée
  ANN_NOM1.Color := clBtnShadow;
end;

procedure TFMessage.BClasserClick(Sender: TObject);
// Le message est traité, classement éventuel dans la "GED".
var i, rep : Integer;
    fichierjoint : TMailFile;
    ParGed : TParamGedDoc;
begin
  inherited;

  if PGIAsk ('Voulez-vous classer/archiver ce message ?', YMS_SUJET.Text)=mrNo then
    exit;

  rep := mrOk;
  if YMS_NODOSSIER.Text='' then
    begin
    rep := PGIAskCancel ('Ce message n''est pas associé à un dossier client, '+#13#10
               +' souhaitez-vous le classer dans le dossier du cabinet ?'+#13#10
               +' (sinon, il sera supprimé...)', YMS_SUJET.Text);
    // sortie
    if rep=mrCancel then
      exit
    // met un no dossier par défaut
    else if rep=mrYes then
      YMS_NODOSSIER.Text := NoDossierBaseCommune;
      // ci-dessous pas bon car écrasé par GetEcranChamp de SauveEnregMsg
      // TobMsg.PutValue('YMS_NODOSSIER', NoDossierBaseCommune);

    end;

  // Archivage
  // Message clôturé = traité
  TobMsg.PutValue('YMS_TRAITE', 'X');

  if not SauveEnregMsg then
  begin
    TobMsg.PutValue('YMS_TRAITE', '-');
    exit;
  end;

  if VH_DP.SeriaGed then
    begin
    // Propose choix d'un classement GED pour les pièces jointes
    for i := 0 to ListViewFiles.Items.Count-1 do
      begin
      // Objets créés par AddAssociateFile
      fichierjoint := TMailFile(ListViewFiles.Items[i].Data);
      ParGed.SDocGuid := '';
      ParGed.NoDossier := TobMsg.GetValue('YMS_NODOSSIER');
      ParGed.CodeGed := '';
      ParGed.SFileGuid := fichierjoint.FileGuid;

      ShowNewDocument (ParGed);
      end;
    end;

  // Message non rattaché à un dossier => à supprimer
  if rep=mrNo then DeleteMessage (FMsgGuid);

  ModalResult := mrOk; // ou Self.Close;
end;

procedure TFMessage.BDeclasserClick(Sender: TObject);
// Le message est traité, classement éventuel dans la "GED".
begin
  inherited;
  if PGIAsk ('Voulez-vous déclasser ce message ?', YMS_SUJET.Text)=mrNo then
    exit;

  TobMsg.PutValue('YMS_TRAITE', '-');

  // Message déclôturé
  if Not SauveEnregMsg then
  begin
    TobMsg.PutValue('YMS_TRAITE', 'X');
    exit;
  end;

  ModalResult := mrOk; // ou Self.Close;
end;

procedure TFMessage.BREPONDREClick(Sender: TObject);
begin
  inherited;
  if SauveEnregMsg then
    ShowNewMessage ('', TobMsg.GetValue('YMS_MSGGUID'), 'RE:');
end;

procedure TFMessage.BREPONDRETOUSClick(Sender: TObject);
begin
  inherited;
  if SauveEnregMsg then
  begin
    ShowNewMessage ('', TobMsg.GetValue('YMS_MSGGUID'), 'RE:', '', '', '', TRUE);
    ModalResult := mrOk; // ou Self.Close;
  end;
end;

procedure TFMessage.BTRANSFERERClick(Sender: TObject);
begin
  inherited;
  if SauveEnregMsg then
    ShowNewMessage ('', TobMsg.GetValue('YMS_MSGGUID'), 'TR:');
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : MD
Créé le ...... : 17/03/2006
Modifié le ... :   /  /
Description .. : Bouton de sauvegarde du message d'origine si existe
Suite ........ : (=celui reçu par la messagerie AGL dans YMAILS)
Suite ........ :
Mots clefs ... :
*****************************************************************}
procedure TFMessage.BExtraireEmlClick(Sender: TObject);
var
    tMsgFiles, TEnreg : Tob;
    sTempFileName : String;
    i : Integer;
begin
  inherited;
  if FMsgGuid='' then exit;

  tMsgFiles := Tob.Create('_YMSGFILES_', nil, -1);
  tMsgFiles.LoadDetailFromSQL('SELECT YMG_FILEGUID, YFI_FILENAME FROM YMSGFILES, YFILES WHERE YMG_FILEGUID=YFI_FILEGUID'
   + ' AND YMG_MSGGUID="'+FMsgGuid+'" AND YMG_ATTACHED="-"');
  for i := 0 to tMsgFiles.Detail.Count-1 do
    begin
    TEnreg := tMsgFiles.Detail[i];

    // 2 enreg avec ATTACHED='-' : on élimine l'enreg html
    if (ExtractFileExt(tEnreg.GetValue('YFI_FILENAME'))='.eml') then
      begin
      // Nom du fichier temporaire
      sTempFileName := V_GedFiles.TempPath + TEnreg.GetValue('YFI_FILENAME');
      if FileExists (sTempFileName) then DeleteFile (sTempFileName);
      // Extraction
      if V_GedFiles.Extract(sTempFileName, TEnreg.GetString('YMG_FILEGUID')) then
        PGIInfo ('Le mail a été extrait dans '+sTempFileName, TitreHalley)
      else
        PGIInfo ('Impossible d''extraire le mail dans '+sTempFileName, TitreHalley);
      break;
      end;
    end;

  tMsgFiles.Free;

end;


procedure TFMessage.Navigate(WeBName, URL: string);
var
  Flags, TargetFrameName, PostData, Headers: OleVariant;
  Web: TWebBrowser_V1;
begin
  if URL = '' then URL := 'about:blank';

  if (Self.FindComponent(WebName) = nil) then
  begin
    Web := TWebBrowser_V1.Create(Self);
    Web.Name := WebName;
    TOleControl(Web).Parent := PageControl.ActivePage;
    Web.Align := alClient;
    Web.Visible := True;
  end
  else
    Web := TWebBrowser_V1(Self.FindComponent(WebName));

  try
    Web.Navigate(URL, Flags, TargetFrameName, PostData, Headers);
  except
  end;
end;

procedure TFMessage.BOuvrirFichierClick(Sender: TObject);
var
  TempFileName: string;
  MailFile: TMailFile;

begin
  if ListViewFiles.Selected=Nil then exit;

  MailFile := TMailFile(ListViewFiles.Selected.Data);

  TempFileName := V_GedFiles.TempFileName(ExtractFileExt(MailFile.FileName));
  V_GedFiles.Extract(TempFileName, MailFile.FileGUID);
  FTempFileNames.Add(TempFileName);
  V_GedFiles.OpenFile(TempFileName);
end;

procedure TFMessage.BVisualiserFichierClick(Sender: TObject);
var
  cMailFile: TMailFile;
  iTab: Integer;
  TabSheet: TTabSheet;
begin
  if ListViewFiles.Selected=Nil then exit;

  cMailFile := TMailFile(ListViewFiles.Selected.Data);

  for iTab := 0 to PageControl.PageCount - 1 do
  begin
    if (TObject(PageControl.Pages[iTab].Tag) is TMailFile) and (cMailFile = TMailFile(PageControl.Pages[iTab].Tag)) then
    begin
      UpdatePageControl(iTab);
      Exit;
    end;
  end;

  cMailFile.TempFileName := V_GedFiles.TempFileName(ExtractFileExt(cMailFile.FileName));
  V_GedFiles.Extract(cMailFile.TempFileName, cMailFile.FileGUID);
  FTempFileNames.Add(cMailFile.TempFileName);

  TabSheet := TTabSheet.Create(Self);
  TabSheet.PageControl := PageControl;
  TabSheet.TabVisible := True;
  TabSheet.Caption := cMailFile.FileName;
  TabSheet.Tag := Integer(cMailFile);
  TabSheet.ImageIndex := FIconFiles[cMailFile.FileIconID].SmallID;
  PageControl.ActivePage := TabSheet;

  UpdatePageControl;
end;

procedure TFMessage.BEnregistrerSousClick(Sender: TObject);
var
  MailFile: TMailFile;

begin
  if ListViewFiles.Selected=Nil then exit;

  MailFile := TMailFile(ListViewFiles.Selected.Data);

  SaveDialog.FileName := MailFile.FileName;

  if SaveDialog.Execute then
    V_GedFiles.Extract(SaveDialog.FileName, MailFile.FileGUID);
end;

procedure TFMessage.UpdatePageControl(Index: Integer);
var
  cMailFile: TMailFile;
  Url: string;
begin
  if Index <> - 1 then PageControl.ActivePageIndex  := Index;

  if TObject(PageControl.ActivePage.Tag) is TMailFile then
  begin
    cMailFile := TMailFile(PageControl.ActivePage.Tag);

    if (cMailFile.TempFileName <> '') and V_GedFiles.FileExists(cMailFile.TempFileName) then
    begin
      Url := 'file:///' + ExpandFileName (cMailFile.TempFileName);
      Navigate ('Web' + IntToStr (PageControl.ActivePageIndex), Url);
    end;
  end;
end;

procedure TFMessage.PageControlChange(Sender: TObject);
begin
  inherited;
  BWebPrecedent.Enabled := False;
  if PageControl.ActivePage<>Nil then
    if TObject(PageControl.Pages[0].Tag) is TMailFile then BWebPrecedent.Enabled := True;
end;


end.
