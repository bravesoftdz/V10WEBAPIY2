unit YNewMessage;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  HTB97, StdCtrls, ExtCtrls, Hctrls, Mask, ComCtrls, HRichEdt, HRichOLE,
  OleCtrls, SHDocVw_TLB, UIUtil,
{$IFDEF EAGLCLIENT}
  MaineAGL,
{$ELSE}
  FE_Main, {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ENDIF}
  UTob, HMsgBox,
  HEnt1, HSysMenu, Menus, MailOl, RtfCounter, M3FP, UGedFiles, Buttons;

type
  TFNewMessage = class(TForm)
    Dock971: TDock97;
    PBouton: TToolWindow97;
    BValider: TToolbarButton97;
    BFerme: TToolbarButton97;
    HelpBtn: TToolbarButton97;
    bDefaire: TToolbarButton97;
    Binsert: TToolbarButton97;
    BDelete: TToolbarButton97;
    BImprimer: TToolbarButton97;
    PCVisu: TPageControl;
    PGeneral: TTabSheet;
    YMS_MEMO: THRichEditOLE;
    PnlHaut: TPanel;
    TYMS_SUJET: THLabel;
    Image1: TImage;
    TYMS_ALLDEST: THLabel;
    TYMS_NODOSSIER: THLabel;
    TYMS_MSGTYPE: THLabel;
    TYMS_DATE: THLabel;
    LaDate: TLabel;
    YMS_SUJET: TEdit;
    ALLDEST: THCritMaskEdit;
    YMS_MSGTYPE: THValComboBox;
    YMS_NODOSSIER: THCritMaskEdit;
    YMS_URGENT: TCheckBox;
    YMS_PRIVE: TCheckBox;
    ALLDESTCC: THCritMaskEdit;
    TALLDESTCC: THLabel;
    YMS_TEL: TEdit;
    TANN_TEL1: THLabel;
    BENVOYER: TToolbarButton97;
    ANN_NOM1: TLabel;
    HMTrad: THSystemMenu;
    PPiecesJointes: TTabSheet;
    OpenDlg: TOpenDialog;
    LstFiles: TListView;
    PopFiles: TPopupMenu;
    mnSupprimer: TMenuItem;
    mnAjouter: TMenuItem;
    PnlPiecejointe: TPanel;
    LJOINDRE: THLabel;
    BJOINDRE: TToolbarButton97;
    Mem_A: THMemo;
    Mem_Cc: THMemo;
    procedure FormCreate(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure FormCloseQuery(Sender: TObject; var CanClose: Boolean);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure YMS_NODOSSIERElipsisClick(Sender: TObject);
    procedure YMS_NODOSSIERExit(Sender: TObject);
    procedure BValiderClick(Sender: TObject);
    procedure OnMsgChange(Sender: TObject);
    procedure BENVOYERClick(Sender: TObject);
    procedure BDeleteClick(Sender: TObject);
    procedure DestinatairesElipsisClick(Sender: TObject);
    procedure YMS_SUJETExit(Sender: TObject);
    procedure BJOINDREClick(Sender: TObject);
    procedure mnSupprimerClick(Sender: TObject);
    procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);

  private
    { Déclarations privées }
    TobMsg, TobMsgFiles         : TOB;         // Enreg. message en cours
    TobMsgOrig, TobMsgFilesOrig : TOB;     // Enreg. message d'origine en cas de TR: ou RE:
    TobMsgAdr                   : TOB;

    TobListeDestA               : TOB;
    TobListeDestCc              : TOB;

    FRepondreATous          : Boolean ;
    FMsgGUID                : String ;      // Guid message ('' tant qu'on enregistre pas)
    FMsgGUIDOrig            : String ;      // Guid message origine en cas de réponse/transfert

    FTypeResp               : String ;      // '' classique, 'RE:' répondre, 'TR:' transférer
    FDefaultPieces          : String ;      // Liste de pièces à inclure à la création du mail
    bDuringLoad, bChange    : Boolean ;     // Chgt dans les données
    LstASuppr               : HTStringList ; // Liste des pièces jointes (fileid) à supprimer
                                            // si modif d'un message enregistré
    FDestParDefaut          : String ;      // Destinataire à renseigner
    FNumDossier             : String ;      // Numéro de dossier

    function ControlerDestinataire (UneListeDestinataire : String): Boolean;
    function TypeDestinataire (UnDest: String) : Integer;
    function IsCodeUserValide (UnDest : String) : Boolean;
    function IsListeDistributionValide (UnDest : String) : Boolean;
    function IsAdresseEmailValide (UnDest : String) : Boolean;

    function  SauveEnregMsg : Boolean;
    procedure AffichePiecesJointes;
    procedure CopieRefPiecesJointes(DeMsgGUID, VersMsgGUID : String);
    function  PutFichiersJointsOnDisk: String;
    procedure VireFichiersJointsOnDisk(sFichiers: String);
    procedure ActuZonesCalculees;
    procedure ActualiseNomEtTelephone(NoDossier: String);
    function  RepondreTransferer(TobOrig : Tob) : String;

  public
    { Déclarations publiques }
    procedure AjoutePieceJointe(nomfichier: String);

  end;

function ShowNewMessage(MsgGUID : String; MsgGUIDOrig: String=''; sTypeResp: String='';
  sDestParDefaut: String=''; sNumDossier: String=''; sDefaultPieces: String=''; RepondreATous: Boolean=FALSE) : Boolean;

////////// IMPLEMENTATION /////////////
implementation

uses UtilDossier,  // pour NouvelleCle
     UtilMessages, YRechercherDestinataire, YMessageDetailDestinataire;


function ShowNewMessage(MsgGUId: String; MsgGUIDOrig: String=''; sTypeResp: String='';
 sDestParDefaut: String=''; sNumDossier: String=''; sDefaultPieces: String=''; RepondreATous: Boolean=FALSE) : Boolean;
// MsgGuid : '' en création, ou <>'' pour message enregistré non encore envoyé
// MsgGuidOrig : <>'' si réponse ou transfert d'un message
var
  FNewMessage : TFNewMessage;
begin
  FNewMessage := TFNewMessage.Create(Application);
  FNewMessage.FMsgGUID       := MsgGUID;
  FNewMessage.FMsgGUIDOrig   := MsgGUIDOrig;
  FNewMessage.FTypeResp      := sTypeResp;
  FNewMessage.FDestParDefaut := sDestParDefaut;
  FNewMessage.FNumDossier    := sNumDossier;
  FNewMessage.FDefaultPieces := sDefaultPieces;
  FNewMessage.FRepondreATous := RepondreATous;
  try
    Result := (FNewMessage.ShowModal = mrOk);
  finally
    FNewMessage.Free ;
  end ;
end;

Procedure AGLEnvoiMessage( parms: array of variant; nb: integer );
var  // F : TForm;
     sDest : String;
begin
  // F:=TForm(Longint(Parms[0]));
  sDest := String(Parms[1]) ;

  ShowNewMessage('', '', '', sDest);
end;

{$R *.DFM}

procedure TFNewMessage.FormCreate(Sender: TObject);
begin
  bDuringLoad := True;
  bChange     := False;
  LstASuppr := HTStringList.Create;
end;


procedure TFNewMessage.FormShow(Sender: TObject);
var RSql                                : TQuery;
    ListePiece                          : String;
    UnDestinataire, UnePiece, SSql      : String;
    ListeDestinataire                   : String;
    EmailCurrentUser                    : String;
    NomDest                             : String;
    Index                               : Integer;
begin
  //--- Préparation des infos
  TobMsg := Tob.Create('YMESSAGES', nil, -1);
  TobMsgAdr:=Tob.Create ('YMSGADDRESS',nil,-1);

  //--- Renseigne la clé primaire
  if FMsgGUID<>'' then TobMsg.PutValue('YMS_MSGGUID', FMsgGUID);
  TobMsg.LoadDB;

  //--- Si on part d'un message d'origine (Cas du répondre à tous), on prépare la tob
  if (FMsgGUIDOrig<>'') then
   begin
    TobMsgOrig := TOB.Create('YMESSAGES', nil, -1);
    //--- Sauf si le message est introuvable
    if Not TobMsgOrig.SelectDB('"'+FMsgGUIDOrig+'"', nil ) then
     begin
      FreeAndNil(TobMsgOrig);
      FMsgGUIDOrig := '';
     end;
   end;

  //--- Préparation des tobs destinataires
  TobListeDestA:=Tob.Create ('',nil,-1);
  TobListeDestCc:=Tob.Create ('',nil,-1);

  if FMsgGUID<>'' then // Si message pré-enregistré
   begin
    InitialiserTobListeDestinataire (TobListeDestA,DonnerListeDestinataireFromSql (FMsgGUID,'A',TobMsg.GetValue('YMS_ALLDEST')),FMsgGUID);
    InitialiserTobListeDestinataire (TobListeDestCc,DonnerListeDestinataireFromSql (FMsgGUID,'Cc',TobMsg.GetValue('YMS_ALLDESTCc')),FMsgGUID);

    TobMsg.PutEcran(Self);
    ActuZonesCalculees;
    try YMS_SUJET.SetFocus; except; end;
   end
  else // Si c'est un nouveau message
   begin
    BDelete.Enabled := False;
    TobMsg.PutValue('YMS_USERFROM', V_PGI.User);
    TobMsg.PutValue('YMS_LIBFROM', V_PGI.UserName);
    TobMsg.PutValue('YMS_BROUILLON', 'X');
    TobMsg.PutValue('YMS_TRAITE', '-');
    TobMsg.PutValue('YMS_LU', '-');
    TobMsg.PutValue('YMS_ENCOPIE', '-');
    TobMsg.PutValue('YMS_ATTACHED', '-');
    TobMsg.PutValue('YMS_DATE', Now);
    TobMsg.PutValue('YMS_MSGTYPE', 'MAI');
    TobMsg.PutValue('YMS_INBOX', 'X');
    ANN_NOM1.Caption := '';
    YMS_MSGTYPE.Value := 'MAI';
    YMS_NODOSSIER.Text := FNumDossier;
    ActualiseNomEtTelephone(YMS_NODOSSIER.Text);

    // Remplissage du destinataire par défaut
    if FDestParDefaut<>'' then
     begin
      FDestParDefaut:=FDestParDefaut+'||'+';';
      InitialiserTobListeDestinataire (TobListeDestA,FDestParDefaut);
     end;

    //--- Si on part d'un message d'origine (Cas du répondre a tous)
    if (FMsgGUIDOrig<>'') and (TobMsgOrig<>Nil) then
     begin
      TobMsg.PutValue('YMS_MSGTYPE',   TobMsgOrig.GetValue('YMS_MSGTYPE'));
      TobMsg.PutValue('YMS_NODOSSIER', TobMsgOrig.GetValue('YMS_NODOSSIER'));

      //--- Ajoute l'entête RE: ou TR: si il n'y est pas déjà
      if Copy(TobMsgOrig.GetValue('YMS_SUJET'), 1, 3)<>FTypeResp then
       TobMsg.PutValue('YMS_SUJET',   FTypeResp+' '+TobMsgOrig.GetValue('YMS_SUJET'))
      else
       TobMsg.PutValue('YMS_SUJET',   TobMsgOrig.GetValue('YMS_SUJET'));

      TobMsg.PutValue('YMS_TEL',       TobMsgOrig.GetValue('YMS_TEL'));
      TobMsg.PutValue('YMS_MEMO',      RepondreTransferer(TobMsgOrig) );
      TobMsg.PutValue('YMS_MAILID',    TobMsgOrig.GetValue('YMS_MAILID'));
      TobMsg.PutValue('YMS_USERMAIL',  TobMsgOrig.GetValue('YMS_USERMAIL'));
      TobMsg.PutValue('YMS_ATTACHED',  TobMsgOrig.GetValue('YMS_ATTACHED'));
      TobMsg.PutValue('YMS_URGENT',    TobMsgOrig.GetValue('YMS_URGENT'));
      TobMsg.PutValue('YMS_PRIVE',     TobMsgOrig.GetValue('YMS_PRIVE'));

      //--- (Cas du message Transfert) l'avant-dernier destinataire est conservé comme intermédiaire
      if FTypeResp='TR:' then TobMsg.PutValue('YMS_USERVIA', TobMsgOrig.GetValue('YMS_USERDEST'));

      //--- (Cas du répondre à tous) on recherche l'emetteur du message pour compléter la liste des destinataires
      if FTypeResp='RE:' then
       begin
        //--- C'est un émetteur interne
        TobMsg.PutValue('YMS_USERDEST', TobMsgOrig.GetValue('YMS_USERFROM'));
        TobMsg.PutValue('YMS_LIBDEST', TobMsgOrig.GetValue('YMS_LIBFROM'));
        UnDestinataire:=TobMsgOrig.GetValue('YMS_USERFROM');

        //--- C'est un émetteur externe (on recherche son email)
        if (UnDestinataire='') then
         begin
          //--- On recherche l'email dans YMSGADDRESS
          SSql:='SELECT YMR_EMAILADDRESS FROM YMSGADDRESS WHERE YMR_MSGGUID="'+FMsgGuidOrig+'" AND YMR_ADTYPE=1 ORDER BY YMR_ADORDRE';
          RSql:=OpenSql (SSql,True);
          if not (Rsql.Eof) then
           UnDestinataire:=RSql.FindField ('YMR_EMAILADDRESS').AsString;
          Ferme (RSql);

          //--- On recherche l'email dans ANNUAIRE
          if (UnDestinataire='') then
           begin
            SSql:='SELECT YMA_FROM FROM YMAILS, YMESSAGES WHERE YMA_UTILISATEUR=YMS_USERMAIL AND YMA_MAILID=YMS_MAILID AND YMS_MSGGUID="'+FMsgGuidOrig+'"';
            RSql:=OpenSql (SSql,True);
            if not (Rsql.Eof) then
             UnDestinataire:=ExtractEmail (RSql.FindField ('YMA_FROM').AsString);
            Ferme (RSql);

            if (UnDestinataire='') and (TobMsgOrig.GetValue ('YMS_NODOSSIER')<>'') then
             begin
              SSql:='SELECT ANN_EMAIL FROM ANNUAIRE,DOSSIER WHERE ANN_GUIDPER=DOS_GUIDPER AND DOS_NODOSSIER="'+TobMsgOrig.GetValue ('YMS_NODOSSIER')+'"';
              RSql:=OpenSql (SSql,True);
              if not (Rsql.Eof) then
               UnDestinataire:=RSql.FindField ('ANN_EMAIL').AsString;
              Ferme (RSql);
             end;
           end;
         end;

        //--- Ajoute l'emetteur à la liste des destinataires
        if (UnDestinataire<>'') then
         begin
          ListeDestinataire:=UnDestinataire+'|'+DonnerLibelleDestinataire (FMsgGUIDOrig,UnDestinataire)+'|'+';';
          InitialiserTobListeDestinataire (TobListeDestA,ListeDestinataire);
         end;

        //--- Initialise les destinataires à partir du mail d'origine
        if (FRepondreATous) then
         begin
          InitialiserTobListeDestinataire (TobListeDestA,DonnerListeDestinataireFromSql (FMsgGUIDOrig,'A',TobMsgOrig.GetValue('YMS_ALLDEST')),FMsgGUIDOrig);
          InitialiserTobListeDestinataire (TobListeDestCc,DonnerListeDestinataireFromSql (FMsgGUIDOrig,'Cc',TobMsgOrig.GetValue('YMS_ALLDESTCc')),FMsgGUIDOrig);

          //RP0 04/12/07 FQ 11868 : retirer le user courant ou son email de la liste des destinataires
          EmailCurrentUser:='';
          RSql := OpenSQL('SELECT US_EMAIL FROM UTILISAT WHERE US_UTILISATEUR="'+V_PGI.User+'"', True);
          if Not RSql.Eof then
            EmailCurrentUser := RSql.FindField('US_EMAIL').AsString;
          Ferme(RSql);

          For Index:=TobListeDestA.Detail.Count-1 downto 0 do
          begin
            NomDest := Uppercase(TobListeDestA.Detail[Index].GetString('Nom'));
            if (NomDest = Uppercase(V_PGI.User)) or (NomDest = Uppercase(EmailCurrentUser)) then
              TobListeDestA.Detail[Index].Free;
          end;

          For Index:=TobListeDestCc.Detail.Count-1 downto 0 do
          begin
            NomDest := Uppercase(TobListeDestCc.Detail[Index].GetString('Nom'));
            if (NomDest = Uppercase(V_PGI.User)) or (NomDest = Uppercase(EmailCurrentUser)) then
              TobListeDestCc.Detail[Index].Free;
          end;

         end;


        //--- Email ??????????
        {EmailOrigine:='';
        SSql:='SELECT US_EMAIL FROM UTILISAT WHERE US_UTILISATEUR="'+TobMsgOrig.GetValue('YMS_USERDEST')+'"';
        RSql:=OpenSql (SSql,True);
        if not (RSql.eof) then
         begin
          if ((RSql.FindField ('US_EMAIL').AsString)<>'') then
           EmailOrigine:=RSql.FindField ('US_EMAIL').AsString;
         end;
        Ferme (RSql);}

       end;

     //--- Affiche les données du message d'origine
     TobMsg.PutEcran(Self);
     ActuZonesCalculees;
    end;
   end;

  //--- Affichage des informations
  try YMS_SUJET.SetFocus; except; end;

  AllDest.Text:=DonnerListeDestinataireFromTob (TobListeDestA,False);
  AllDestCc.Text:=DonnerListeDestinataireFromTob (TobListeDestCc,False);
  //GUH31.10.07
  Mem_A.Text := DonnerListeDestinataireFromTob (TobListeDestA,TRUE);
  Mem_Cc.Text := DonnerListeDestinataireFromTob (TobListeDestCc,TRUE);

  LaDate.Caption := FormatDateTime('aaaa DD/MM/YYYY HH:NN:SS', TobMsg.GetValue('YMS_DATE'));
  Self.Caption := 'Nouveau message - '+YMS_SUJET.Text;

  //--- Traitement des pieces jointes, Liste des fichiers issus de YMSGFILES ou de FDefaultPieces
  ListePiece := FDefaultPieces;
  while (ListePiece<>'') do
   begin
    UnePiece:=ReadTokenSt(ListePiece);
    AjoutePieceJointe(UnePiece);
   end;
  AffichePiecesJointes;

  bDuringLoad := False;
end;


procedure TFNewMessage.FormCloseQuery(Sender: TObject;
  var CanClose: Boolean);
var rep: Integer;
begin
  inherited;
  // Changement demandé
  if bChange then
    begin
    rep := PGIAskCancel('Voulez-vous enregistrer les modifications ?', Self.Caption);
    if rep=mrCancel then
      CanClose := False
    else if rep=mrNo then
      bChange := False
    else if rep=mrYes then
      CanClose := SauveEnregMsg;
    end;
end;


procedure TFNewMessage.FormClose(Sender: TObject;
  var Action: TCloseAction);
var i: Integer;
begin
  for i:=0 to LstFiles.Items.Count-1 do
    if LstFiles.Items[i].Data<>Nil then Tob(LstFiles.Items[i].Data).Free;
  // Même si ses tob filles sont purgées par la ligne du dessus, pas grave :
  if TobMsgFilesOrig<>Nil then TobMsgFilesOrig.Free;
  if TobMsgFiles<>Nil then TobMsgFiles.Free;
  if TobMsg<>Nil then TobMsg.Free;
  if TobMsgAdr<>nil then TobMsgAdr.Free;
  if TobMsgOrig<>Nil then TobMsgOrig.Free;
  if TobListeDestA<>nil then TobListeDestA.Free;
  if TobListeDestCc<>nil then TobListeDestCc.Free;

  LstASuppr.Free;
end;


procedure TFNewMessage.YMS_NODOSSIERElipsisClick(Sender: TObject);
var St       : String;
    sGuidPer : String;
    Q        : TQuery;
begin
   // retourne NoDossier;Nom1
   St := AGLLanceFiche('YY','YYDOSSIER_SEL', '','',YMS_NODOSSIER.Text);
   if St<>'' then
     begin
     YMS_NODOSSIER.Text := READTOKENST(St);
     // on utilise pas GetNomDossier, car on a besoin de ANN_TEL1 en plus
     ANN_NOM1.Caption := '';
     sGuidPer := READTOKENST(St);
     if sGuidPer<>'' then
       begin
       Q := OpenSQL('SELECT ANN_NOM1, ANN_TEL1 FROM ANNUAIRE WHERE ANN_GUIDPER="'+sGuidPer+'"', True);
       if Not Q.Eof then
         begin
         ANN_NOM1.Caption := Q.FindField('ANN_NOM1').AsString;
         YMS_TEL.Text := Q.FindField('ANN_TEL1').AsString;
         end;
       Ferme(Q);
       end;
     end;
end;

procedure TFNewMessage.YMS_NODOSSIERExit(Sender: TObject);
begin
  ActualiseNomEtTelephone(YMS_NODOSSIER.Text);
end;

function TFNewMessage.SauveEnregMsg : Boolean;
// Sauvegarde d'un message brouillon (non envoyé)
// True si tout ok
var i : Integer;
    elt : TListItem;
    tMsgFile, tobFiles : TOB;
    sFileGUID : String;

    UneTobMsgAdrEnreg    : Tob;
    UneListeDestinataire : String;
    Indice               : Integer;
    strSubject           : String;
begin
  Result := False;
  // Contrôle avant validation
  if YMS_SUJET.Text='' then
    begin
    PGIInfo('Le sujet n''est pas renseigné.', Self.Caption);
    YMS_SUJET.SetFocus;
    exit;
    end;

  // Récup des modifs
  TobMsg.GetEcran(Self);

  // Cas d'un nouvel enreg, on lui attribue la clé
  if FMsgGUID='' then
   begin
    FMsgGUID:=AGLGetGUID();
    TobMsg.PutValue('YMS_MSGGUID', FMsgGUID);
   end;

  UneListeDestinataire:=AllDest.Text;
  // FQ 11873 - élimine les guillemets
  VireGuillemets(UneListeDestinataire);
  AllDest.Text := UneListeDestinataire;
  TobMsg.PutValue ('YMS_ALLDEST',DonnerListeDestinataireFromTob (TobListeDestA,False));
  SupprimerDestinataireDansTob (TobListeDestA, UneListeDestinataire);
  AjouterDestinataireDansTob (TobListeDestA, UneListeDestinataire);
  TobMsg.PutValue ('YMS_ALLDEST',DonnerListeDestinataireFromTob (TobListeDestA,False));

  UneListeDestinataire:=AllDestCc.Text;
  // FQ 11873 - élimine les guillemets
  VireGuillemets(UneListeDestinataire);
  AllDestCc.Text := UneListeDestinataire;
  TobMsg.PutValue ('YMS_ALLDESTCc',DonnerListeDestinataireFromTob (TobListeDestCc,False));
  SupprimerDestinataireDansTob (TobListeDestCc, UneListeDestinataire);
  AjouterDestinataireDansTob (TobListeDestCc, UneListeDestinataire);
  TobMsg.PutValue ('YMS_ALLDESTCc',DonnerListeDestinataireFromTob (TobListeDestCc,False));

  // suite FQ 11873 - dans le sujet aussi
  strSubject:=YMS_SUJET.Text;
  VireGuillemets(strSubject);
  YMS_SUJET.Text:=strSubject;
  TobMsg.PutValue ('YMS_SUJET',strSubject);

  //--- Traitement des destinataires A
  for Indice:=0 to TobListeDestA.Detail.count-1 do
   begin
    UneTobMsgAdrEnreg:=Tob.Create ('YMSGADDRESS',TobMsgAdr,-1);
    UneTobMsgAdrEnreg.PutValue ('YMR_MSGGUID',FMsgGUID);
    UneTobMsgAdrEnreg.PutValue ('YMR_ADTYPE',2);
    UneTobMsgAdrEnreg.PutValue ('YMR_ADORDRE',Indice+1);
    UneTobMsgAdrEnreg.PutValue ('YMR_EMAILADDRESS',TobListeDestA.Detail [Indice].GetValue ('Nom'));
    UneTobMsgAdrEnreg.putValue ('YMR_EMAILNAME',TobListeDestA.Detail [Indice].GetValue ('Libelle'));
   end;

  //--- Traitement des destinataires Cc
  for Indice:=0 to TobListeDestCc.Detail.count-1 do
   begin
    UneTobMsgAdrEnreg:=Tob.Create ('YMSGADDRESS',TobMsgAdr,-1);
    UneTobMsgAdrEnreg.PutValue ('YMR_MSGGUID',FMsgGUID);
    UneTobMsgAdrEnreg.PutValue ('YMR_ADTYPE',3);
    UneTobMsgAdrEnreg.PutValue ('YMR_ADORDRE',Indice+1);
    UneTobMsgAdrEnreg.PutValue ('YMR_EMAILADDRESS',TobListeDestCc.Detail [Indice].GetValue ('Nom'));
    UneTobMsgAdrEnreg.putValue ('YMR_EMAILNAME',TobListeDestCc.Detail [Indice].GetValue ('Libelle'));
   end;

  // Y'a t'il des pièces jointes
  if LstFiles.Items.Count>0 then TobMsg.PutValue('YMS_ATTACHED', 'X')
  else TobMsg.PutValue('YMS_ATTACHED', '-');
  TobMsg.InsertOrUpdateDB;

  //--- Effacement de
  ExecuteSQL ('DELETE FROM YMSGADDRESS WHERE YMR_MSGGUID="'+FMsgGuid+'"');
  TobMsgAdr.InsertOrUpdateDB;

  // Pièces jointes à supprimer dans la GED (cas de modif d'un brouillon pré-enregistré)
  // => par contre, si nouveau message issu d'un TR: / RE:, on n'a rien à virer
  // car les fichiers ne sont pas encore réf. dans YMSGFILES...
  if (FMsgGUIDOrig='') then // de toutes façons, LstASuppr n'est pas rempli dans ce cas
   begin
    for i:=0 to LstASuppr.Count-1 do
      begin
       sFileGUID :=LstASuppr[i];
       ExecuteSQL('DELETE FROM YMSGFILES WHERE YMG_MSGGUID="'+FMsgGUID+'" AND YMG_FILEGUID="'+sFileGUID+'"');
       tobFiles := Tob.Create('_YFILES_', nil, -1);
       tobFiles.LoadDetailDBFromSQL('YFILES', 'SELECT * FROM YFILES WHERE YFI_FILEGUID="'+sFileGUID+'"');
       tobFiles.DeleteDBTom('YFILES');
       tobFiles.Free;
      end;
   end;

  // Cherche nouvelles pièces jointes
  for i:=0 to LstFiles.Items.Count-1 do
    begin
    elt := LstFiles.Items[i];

    // Fichier déjà dans la GED
    if elt.Data<>Nil then
      BEGIN
      // tob sur YMSGFILES = enreg existant => rien à modifier
      if TOB(elt.Data).FieldExists('YMG_FILEGUID') then Continue

      // fichier issu d'un message d'origine : est déjà dans la GED
      // manque juste de le référencer au message en cours
      else if (FTypeResp='TR:') and (TobMsgOrig<>Nil) then // + sûr que FMsgIdOrig>0
        begin
        if TOB(elt.Data).GetValue('YFI_FILEGUID')='' then Continue;
        ExecuteSQL('INSERT YMSGFILES (YMG_MSGGUID, YMG_FILEGUID, YMG_ATTACHED)'
         + ' VALUES ("'+FMsgGUID+'","'+TOB(elt.Data).GetValue('YFI_FILEGUID')+'", "X")' );
        // fichier suivant
        Continue;
        end;
      END

    // nouveau fichier à insérer
    else
      BEGIN
      if Not FileExists(elt.Caption) then
        PGIInfo('Le fichier '+elt.Caption+' n''existe plus.', 'Pièces jointes')
      else
        begin
        // Enregistre la pièce jointe dans YFILES, YFILEPARTS...
        sFileGUID := V_GedFiles.Import(elt.Caption);
        // et YMSGFILES
        tMsgFile := TOB.Create('YMSGFILES', Nil, -1);

        if Not tMsgFile.SelectDB('"'+FMsgGUID+'";"'+sFileGUID+'"', Nil, True) then
          begin
          // clé primaire
          tMsgFile.PutValue('YMG_MSGGUID', FMsgGUID);
          tMsgFile.PutValue('YMG_FILEGUID', sFileGUID);
          tMsgFile.PutValue('YMG_ATTACHED', 'X');
          tMsgFile.SetAllModifie(True);
          // maj dans base
          tMsgFile.InsertDB(Nil);
          end;

        tMsgFile.Free;
        end;
      END;

    end;

  // On redessine pour que les .Data soient renseignés dans le listview
  AffichePiecesJointes;

  bChange := False;
  Result := True;
end;


procedure TFNewMessage.BValiderClick(Sender: TObject);
begin
 // Par défaut, le bouton BValider referme la fiche
 if Not SauveEnregMsg then ModalResult := mrNone;
 // donc on l'en empêche si SauveEnregMsg n'a pas pu valider
end;

procedure TFNewMessage.OnMsgChange(Sender: TObject);
begin
  if not bDuringLoad then bChange := True;
end;

//-------------------------------------------------------------------
//--- Nom   : ControlerDestinataire
//--- Objet : Vérifie l'intégrité des adresse email, élimine les
//---         adresses non reconnues de la liste
//-------------------------------------------------------------------
function TFNewMessage.ControlerDestinataire (UneListeDestinataire : String): Boolean;
var UnDest : String;
begin
 Result:=True;
 while (UneListeDestinataire<>'') do
  begin
   UnDest:=ReadToKenPipe (UneListeDestinataire,';');

   Case (TypeDestinataire (Undest)) of
    1 : if not (IsCodeUserValide (UnDest)) then
         begin
          PgiInfo (UnDest+' n''est pas un destinataire interne correct.',TitreHalley);
          Result:=False;
          Exit;
         end;

    2 : if not (IsListeDistributionValide (UnDest)) then
         begin
          PgiInfo (UnDest+' n''est pas une liste de distribution correcte.',TitreHalley);
          Result:=False;
          Exit;
         end;

    3 : if not (IsAdresseEmailValide (UnDest)) then
         begin
          PgiInfo (UnDest+' n''est pas une adresse Email correcte.', TitreHalley);
          Result:=False;
          Exit;
         end;
   end
  end;
end;

//----------------------------------------------------------
//--- Nom : IsCodeUserValide
//----------------------------------------------------------
function TFNewMessage.TypeDestinataire (UnDest: String) : Integer;
begin
 //--- Code utilisateur
 if (Length (UnDest)<4) then
   Result:=1
 //--- Liste de distribution
 else if (Pos('@',undest)=0) then
   Result:=2
 //--- Adresse Email
 else
   Result:=3;
end;

//----------------------------------------------------------
//--- Nom : IsCodeUserValide
//----------------------------------------------------------
function TFNewMessage.IsCodeUserValide (UnDest : String) : Boolean;
begin
 Result:=ExisteSQL('SELECT US_UTILISATEUR FROM UTILISAT WHERE US_UTILISATEUR="'+UnDest+'"')
end;

//----------------------------------------------------------
//--- Nom : IsListeDistributionValide
//----------------------------------------------------------
function TFNewMessage.IsListeDistributionValide (UnDest : String) : Boolean;
begin
 Result:=ExisteSQL ('SELECT YLI_LISTEDISTRIB FROM YLISTEDISTRIB WHERE YLI_LISTEDISTRIB="'+UnDest+'"')
end;

//----------------------------------------------------------
//--- Nom : IsAdresseEmailValide
//----------------------------------------------------------
function TFNewMessage.IsAdresseEmailValide (UnDest : String) : Boolean;
var LaPosition : Integer;
begin
 Result:=False;
 LaPosition:=pos ('@',UnDest);
 if (LaPosition>0) then
  begin
   UnDest [LaPosition]:='0';

   LaPosition:=pos ('@',UnDest);
   if (LaPosition>0) then Exit;

   LaPosition:=pos ('.',UnDest);
   if (LaPosition=0) then Exit;
  end;
 Result:=True;
end;

//-------------------------------------------------------------------
//--- Nom   : BEnvoyerClick
//-------------------------------------------------------------------
procedure TFNewMessage.BENVOYERClick(Sender: TObject);
var UnDestinataire                            : String;
    SDestinataireExtA, SDestinataireExtCc     : String;
    FichiersJoints, Txt, SNewMsgGuid          : String;
    Contenu                                   : HTStrings;
    bEnvoiOk                                  : Boolean;
    Indice                                    : Integer;
    UneTobMsgAdrEnreg                         : Tob;
    UneListeDestinataireA                     : String;
    UneListeDestinataireCc                    : String;
    SNewListeDestinataireA                    : String;
    SNewListeDestinataireCc                   : String;
    tmp                                       : String;
begin
 // FQ 11873 - élimine les guillemets
 tmp := AllDest.Text;
 VireGuillemets(tmp);
 AllDest.Text := tmp;
 SupprimerDestinataireDansTob (TobListeDestA, AllDest.Text);
 AjouterDestinataireDansTob (TobListeDestA,AllDest.Text);
 UneListeDestinataireA:=DonnerListeDestinataireFromTob (TobListeDestA,False);

 // FQ 11873 - élimine les guillemets
 tmp := AllDestCC.Text;
 VireGuillemets(tmp);
 AllDestCC.Text := tmp;
 SupprimerDestinataireDansTob (TobListeDestCc, AllDestCc.Text);
 AjouterDestinataireDansTob (TobListeDestCc,AllDestCc.Text);
 UneListeDestinataireCc:=DonnerListeDestinataireFromTob (TobListeDestCc,False);

 // suite FQ 11873 - dans le sujet aussi
 tmp := YMS_SUJET.Text;
 VireGuillemets(tmp);
 YMS_SUJET.Text:=tmp;

 //--- Contrôles avant validation
 if (UneListeDestinataireA='') and (UneListeDestinataireCc='') then
  begin
   PGIInfo('Aucun destinataire sélectionné.', Self.Caption);
   ALLDEST.SetFocus;
   exit;
  end;

 if Not ControlerDestinataire(UneListeDestinataireA) then
  begin
   ALLDEST.SetFocus;
   exit;
  end;

 if Not ControlerDestinataire(UneListeDestinataireCc) then
  begin
   ALLDESTCC.SetFocus;
   exit;
  end;

  //--- Enregistrement du message + traitement des pièces jointes (TobMsgFiles, LstFiles)
  if (not SauveEnregMsg) then exit;

  //--- Création d'un nouvel enregistrement : Maj des valeurs suivantes
  TobMsg.PutValue('YMS_LIBFROM', V_PGI.UserName);
  TobMsg.PutValue('YMS_BROUILLON', '-');
  TobMsg.PutValue('YMS_INBOX', '-');

  //--- Eclatement des listes de distibution des destinataires A
  SNewListeDestinataireA:=''; SDestinataireExtA:='';
  While (UneListeDestinataireA<>'') do
   begin
    UnDestinataire:=ReadTokenSt(UneListeDestinataireA);
    if (Pos('@', UnDestinataire)=0) and (length (UnDestinataire)>3) then // Liste de distribution
     SNewListeDestinataireA:=SNewListeDestinataireA+DonnerElementListeDistribution (UnDestinataire)
    else // Le destinataire n'est pas une liste de distribution
     SNewListeDestinataireA:=SNewListeDestinataireA+UnDestinataire+';';
   end;

  while (SNewListeDestinataireA<>'') do
   begin
    UnDestinataire:=ReadToKenSt (SNewListeDestinataireA);

    //--- Destinataire interne
    if Pos('@', UnDestinataire)=0 then
     begin
      TobMsg.PutValue('YMS_USERDEST', UnDestinataire);
      TobMsg.PutValue('YMS_LIBDEST', GetNomUser(UnDestinataire));
      TobMsg.PutValue('YMS_ENCOPIE', '-');

      sNewMsgGUID := AGLGetGUID ();
      TobMsg.PutValue('YMS_MSGGUID', sNewMsgGUID);
      TobMsg.SetAllModifie(True);
      TobMsg.InsertDb(Nil);

      CopieRefPiecesJointes(FMsgGUID, sNewMsgGUID);

      //--- Traitement des destinataires A
      for Indice:=0 to TobListeDestA.Detail.count-1 do
       begin
        UneTobMsgAdrEnreg:=Tob.Create ('YMSGADDRESS',TobMsgAdr,-1);
        UneTobMsgAdrEnreg.PutValue ('YMR_MSGGUID',SNewMsgGUID);
        UneTobMsgAdrEnreg.PutValue ('YMR_ADTYPE',2);
        UneTobMsgAdrEnreg.PutValue ('YMR_ADORDRE',Indice+1);
        UneTobMsgAdrEnreg.PutValue ('YMR_EMAILADDRESS',TobListeDestA.Detail [Indice].GetValue ('Nom'));
        UneTobMsgAdrEnreg.putValue ('YMR_EMAILNAME',TobListeDestA.Detail [Indice].GetValue ('Libelle'));
       end;

      //--- Traitement des destinataires Cc
      for Indice:=0 to TobListeDestCc.Detail.count-1 do
       begin
        UneTobMsgAdrEnreg:=Tob.Create ('YMSGADDRESS',TobMsgAdr,-1);
        UneTobMsgAdrEnreg.PutValue ('YMR_MSGGUID',SNewMsgGUID);
        UneTobMsgAdrEnreg.PutValue ('YMR_ADTYPE',3);
        UneTobMsgAdrEnreg.PutValue ('YMR_ADORDRE',Indice+1);
        UneTobMsgAdrEnreg.PutValue ('YMR_EMAILADDRESS',TobListeDestCc.Detail [Indice].GetValue ('Nom'));
        UneTobMsgAdrEnreg.putValue ('YMR_EMAILNAME',TobListeDestCc.Detail [Indice].GetValue ('Libelle'));
       end;
     end
    //--- Destinataire externe
    else
     begin
      if YMS_MSGTYPE.Value='MAI' then
        begin
         if SDestinataireExtA<>'' then SDestinataireExtA:=SDestinataireExtA+';';
          SDestinataireExtA:=SDestinataireExtA+UnDestinataire;
        end;
     end;
   end;

  //--- Eclatement des listes de distibution des destinataires Cc
  SNewListeDestinataireCc:='';SDestinataireExtCc:='';
  While (UneListeDestinataireCc<>'') do
   begin
    UnDestinataire:=ReadTokenSt(UneListeDestinataireCc);
    if (Pos('@', UnDestinataire)=0) and (length (UnDestinataire)>3) then // Liste de distribution
     SNewListeDestinataireCc:=SNewListeDestinataireCc+DonnerElementListeDistribution (UnDestinataire)
    else // Le destinataire n'est pas une liste de distribution
     SNewListeDestinataireCc:=SNewListeDestinataireCc+UnDestinataire+';';
   end;

  while (SNewListeDestinataireCc<>'') do
   begin
    UnDestinataire:=ReadToKenSt (SNewListeDestinataireCc);

    //--- Destinataire interne
    if Pos('@', UnDestinataire)=0 then
     begin
      TobMsg.PutValue('YMS_USERDEST', UnDestinataire);
      TobMsg.PutValue('YMS_LIBDEST', GetNomUser(UnDestinataire));
      TobMsg.PutValue('YMS_ENCOPIE', 'X');

      sNewMsgGUID := AGLGetGUID;
      TobMsg.PutValue('YMS_MSGGUID',sNewMsgGUID);
      TobMsg.SetAllModifie(True);
      TobMsg.InsertDb(Nil);

      CopieRefPiecesJointes(FMsgGUID,sNewMsgGUID);

      //--- Traitement des destinataires A
      for Indice:=0 to TobListeDestA.Detail.count-1 do
       begin
        UneTobMsgAdrEnreg:=Tob.Create ('YMSGADDRESS',TobMsgAdr,-1);
        UneTobMsgAdrEnreg.PutValue ('YMR_MSGGUID',FMsgGUID);
        UneTobMsgAdrEnreg.PutValue ('YMR_ADTYPE',2);
        UneTobMsgAdrEnreg.PutValue ('YMR_ADORDRE',Indice+1);
        UneTobMsgAdrEnreg.PutValue ('YMR_EMAILADDRESS',TobListeDestA.Detail [Indice].GetValue ('Nom'));
        UneTobMsgAdrEnreg.putValue ('YMR_EMAILNAME',TobListeDestA.Detail [Indice].GetValue ('Libelle'));
       end;

      //--- Traitement des destinataires Cc
      for Indice:=0 to TobListeDestCc.Detail.count-1 do
       begin
        UneTobMsgAdrEnreg:=Tob.Create ('YMSGADDRESS',TobMsgAdr,-1);
        UneTobMsgAdrEnreg.PutValue ('YMR_MSGGUID',FMsgGUID);
        UneTobMsgAdrEnreg.PutValue ('YMR_ADTYPE',3);
        UneTobMsgAdrEnreg.PutValue ('YMR_ADORDRE',Indice+1);
        UneTobMsgAdrEnreg.PutValue ('YMR_EMAILADDRESS',TobListeDestCc.Detail [Indice].GetValue ('Nom'));
        UneTobMsgAdrEnreg.putValue ('YMR_EMAILNAME',TobListeDestCc.Detail [Indice].GetValue ('Libelle'));
       end;
     end
    //--- Destinataire Externe
    else
     begin
      if YMS_MSGTYPE.Value='MAI' then
       begin
        if SDestinataireExtCc<>'' then SDestinataireExtCc:=SDestinataireExtCc+';';
         SDestinataireExtCc:=SDestinataireExtCc+UnDestinataire;
        end;
      end;
   end;

  //--- Envoi du mail
  bEnvoiOk := True;
  if (SDestinataireExtA<>'') or (SDestinataireExtCc<>'') then
   begin
    //--- Les pièces jointes doivent être sur disque pour l'envoi
    FichiersJoints := PutFichiersJointsOnDisk;
    Contenu := HTStringList.Create;
    //--- GetRTFStringText renvoie vide si pas Rtf...
    Txt := GetRTFStringText(TobMsg.GetValue('YMS_MEMO'));
    if Txt<>'' then Contenu.Text:=Txt else Contenu.Text:=TobMsg.GetValue('YMS_MEMO');

    //--- Si aucun destinataire externe, on met les CC en destinataires
    if SDestinataireExtA='' then
     begin
      SDestinataireExtA:=SDestinataireExtCc;
      SDestinataireExtCc:='';
     end;

    //--- Envoi du mail
    SendMail(TobMsg.GetValue('YMS_SUJET'), SDestinataireExtA, SDestinataireExtCc, Contenu, FichiersJoints, True);
    if GetLastMailError<>'' then
      begin
       bEnvoiOk := False;
       PGIInfo('L''envoi de ce mail vers l''extérieur a échoué. '+ #13#10
             + ' Retournez ultérieurement dans <Brouillons>' + #13#10
             + ' et envoyez de nouveau ce mail aux destinataires.', Self.Caption);
       // Le message doit rester dans les brouillons
       TobMsg.PutValue('YMS_MSGGUID', FMsgGUID);
       TobMsg.PutValue('YMS_BROUILLON', 'X');
       TobMsg.PutValue('YMS_INBOX', 'X');
       TobMsg.InsertOrUpdateDB;

       for Indice:=0 to TobMsgAdr.Detail.Count-1 do
        begin
         if (TobMsgAdr.Detail [Indice].GetValue ('YMS_MSGGUID')=FMsgGuid) then
          TobMsgAdr.Detail [Indice].InsertOrUpdateDB;
        end;
      end;

    Contenu.Free;
    VireFichiersJointsOnDisk(FichiersJoints);
   end;

  bChange := False;

  //--- Le brouillon n'en est plus un
  if bEnvoiOk and (FMsgGUID<>'') then
   begin
    ExecuteSQL('UPDATE YMESSAGES SET YMS_BROUILLON="-" WHERE YMS_MSGGUID="'+FMsgGUID+'"');
    TobMsgAdr.InsertOrUpdateDB;
   end;

  ModalResult := mrOk;
end;

procedure TFNewMessage.BDeleteClick(Sender: TObject);
begin
  if FMsgGUID<>'' then
    begin
    DeleteMessage(FMsgGUID);
    ModalResult := mrOk; // ou Self.Close;
    end;
end;

procedure TFNewMessage.YMS_SUJETExit(Sender: TObject);
begin
  Self.Caption := 'Nouveau message - '+YMS_SUJET.Text;
end;

procedure TFNewMessage.DestinatairesElipsisClick(Sender: TObject);
begin

 InitialiserTobListeDestinataire (TobListeDestA,DonnerListeDestinatairefromSql (TobMsg.GetValue ('YMS_MSGGUID'),'A',ALLDEST.Text));
 InitialiserTobListeDestinataire (TobListeDestCc,DonnerListeDestinatairefromSql (TobMsg.GetValue ('YMS_MSGGUID'),'Cc',ALLDESTCC.Text));

 RechercherDestinataire(YMS_NoDossier.Text,TobListeDestA,TobListeDestCc,YMS_MSGTYPE.Value='MAI');

 AllDest.Text:=DonnerListeDestinataireFromTob (TobListeDestA,False);
 AllDestCc.Text:=DonnerListeDestinataireFromTob (TobListeDestCc,False);

 //GUH31.10.07 > affichage avec nom complet dans un Tmemo en lecture seule
 Mem_A.Text := DonnerListeDestinataireFromTob (TobListeDestA,True);
 Mem_Cc.Text:=DonnerListeDestinataireFromTob (TobListeDestCc,True);

 //GUH31.10.07 - mise en commentaire
 //AllDest.HInt:=DonnerListeDestinataireFromTob (TobListeDestA,True);
 //AllDestCc.HInt:=DonnerListeDestinataireFromTob (TobListeDestCc,True);

 YMS_SUJET.SetFocus;
end;


procedure TFNewMessage.AffichePiecesJointes;
var i: Integer;
    elt : TListItem;
    TEnreg : Tob;
begin
  // Ouverture d'un nouveau message
 if (FMsgGUID='') then
    begin
    // phase de transfert d'un message d'origine : on affiche ses pièces jointes
    if (FTypeResp='TR:') and (TobMsgOrig<>Nil) then
      begin
      TobMsgFilesOrig := TOB.Create('Fichiers origines attachés', nil, -1);
      // si issu d'un mail entrant, les pièces jointes sont listées uniqmt dans YMAILFILES
      {if (TobMsgOrig.GetValue('YMS_MSGTYPE')='MAI') and (TobMsgOrig.GetValue('YMS_MAILID')>0) then
        TobMsgFilesOrig.LoadDetailFromSQL('SELECT YFI_FILEID, YFI_FILENAME FROM YMAILFILES, YFILES'
         + ' WHERE YMF_MAILID='+IntToStr(TobMsgOrig.GetValue('YMS_MAILID'))
         + ' AND YMF_UTILISATEUR="'+TobMsgOrig.GetValue('YMS_USERMAIL')+'"'
         + ' AND YMF_FILEID=YFI_FILEID'
         // seult les vrais pièces jointes
         + ' AND YMF_ATTACHED="X"')
      // sinon dans YMSGFILES
      else }
      // MD 14/01/05 - En fait toujours à prendre dans YMSGFILES
      TobMsgFilesOrig.LoadDetailFromSQL('SELECT YFI_FILEGUID, YFI_FILENAME FROM YMSGFILES, YFILES'
       + ' WHERE YMG_MSGGUID="'+FMsgGUIDOrig+'" AND YMG_FILEGUID=YFI_FILEGUID');

      for i:=0 to TobMsgFilesOrig.Detail.Count-1 do
        begin
        TEnreg := TobMsgFilesOrig.Detail[i];
        elt := LstFiles.Items.Add;
        // On garde une copie de la tob...
        elt.Data := TEnreg;
        elt.Caption := TEnreg.GetValue('YFI_FILENAME');
        end;
      end;
    end
  // ou bien récup des infos
  else
    begin
    // purge préalable
    if TobMsgFiles<>Nil then TobMsgFiles.Free;
    LstFiles.Items.Clear;
    TobMsgFiles := TOB.Create('Fichiers attachés', nil, -1);
    // on charge les fichiers du messages => YMG_FILEID servira pour identifier
    // les pièces jointes déjà dans la GED
    TobMsgFiles.LoadDetailFromSQL('SELECT YMG_FILEGUID, YFI_FILENAME FROM YMSGFILES, YFILES'
     + ' WHERE YMG_MSGGUID="'+FMsgGUID+'" AND YMG_FILEGUID=YFI_FILEGUID');
    // TobMsgFiles.PutListViewDetail(LstFiles, True, True, 'YFI_FILENAME');
    // => non, ne renseignait pas Data !
    for i:=0 to TobMsgFiles.Detail.Count-1 do
      begin
      TEnreg := TobMsgFiles.Detail[i];
      elt := LstFiles.Items.Add;
      elt.Data := TEnreg;
      elt.Caption := TEnreg.GetValue('YFI_FILENAME');
      end;
    // TobMsgFiles.Free; Non, car purgerait les tob filles référencées dans .Data !
    end;
  PCVisu.ActivePage := PGeneral;
  PPiecesJointes.TabVisible := (LstFiles.Items.Count>0);
end;


procedure TFNewMessage.BJOINDREClick(Sender: TObject);
begin
  // Choix d'un fichier
  if OpenDlg.Execute then
    AjoutePieceJointe(OpenDlg.FileName);
end;


procedure TFNewMessage.mnSupprimerClick(Sender: TObject);
var T: Tob;
    fileid : Integer;
begin
  if LstFiles.ItemFocused=Nil then exit;

  // suppression demandée d'une pièce jointe (repère : la tob fille stockée dans data)
  if (FMsgGUID<>'') and (FMsgGUIDOrig='') then
    begin
    // sauf si on ne l'a pas encore intégrée dans YFILES...
    // (car on vient de la rajouter dans le brouillon, mais sans sauver)
    if LstFiles.ItemFocused.Data<>Nil then
      begin
      T := Tob(LstFiles.ItemFocused.Data);
      fileid := T.GetValue('YFI_FILEID');
      if fileid>0 then
        // on ne supprime pas directement dans les tables,
        // mais on ajoute dans une liste temporaire (garde possibilité bouton "Annuler")
        if LstASuppr.IndexOf(IntToStr(fileid))<0 then LstASuppr.Add(IntToStr(fileid));
      end;
    end;

  // Dans tous les cas, retrait de l'icône du listview
  LstFiles.ItemFocused.Delete;

  // Si plus besoin affichage listview
  if LstFiles.Items.Count=0 then
    begin
    PCVisu.ActivePage := PGeneral;
    PPiecesJointes.TabVisible := False;
    end;

  bChange := True;
end;


procedure TFNewMessage.CopieRefPiecesJointes(DeMsgGUID, VersMsgGUID : String);
// recopie les pièces jointes d'un message vers un autre message
// (juste la référence aux pièces jointes, car on ne va pas dupliquer les données)
begin
  if (DeMsgGUID='') or (VersMsgGUID='') then exit;
  ExecuteSQL('INSERT INTO YMSGFILES (YMG_MSGGUID, YMG_FILEGUID, YMG_ATTACHED)'
   + ' SELECT "'+VersMsgGUID+'" AS YMG_MSGGUID, YMG_FILEGUID, YMG_ATTACHED'
   + ' FROM YMSGFILES WHERE YMG_MSGGUID="'+DeMsgGUID+'"');
end;

(* mcd  09/02/2006 ?? fct non utilisée !!!
//$$$ JP 06/12/05 - warning delphi
{$IFNDEF BUREAU}
procedure TFNewMessage.CopieRefUnePieceJointe(DeMsgId, VersMsgId, FileId : Integer);
begin
  if (DeMsgId=0) or (VersMsgId=0) or (FileId=0) then exit;
  ExecuteSQL('INSERT INTO YMSGFILES SELECT '+IntToStr(VersMsgId)+' AS YMG_MSGID, YMG_FILEID'
   +' FROM YMSGFILES WHERE YMG_MSGID='+IntToStr(DeMsgId)+' AND YMG_FILEID='+IntToStr(FileId));
end;
{$ENDIF}  *)

function TFNewMessage.PutFichiersJointsOnDisk : String;
// Dépose sur disque les fichiers joints au message en cours
// Retourne la liste des fichiers séparée par des ;
var TmpPath, FilePath : String;
    TEnreg : TOB;
    i : Integer;
begin
  Result := '';
  // Message pas encore enregistré
  if FMsgGUID='' then exit;

  TmpPath := V_GedFiles.TempPath ;
  // Pièces jointes au message en cours
  // chargées par AffichePiecesJointes
  for i:=0 to TobMsgFiles.Detail.Count-1 do
    begin
    TEnreg := Tob(TobMsgFiles.Detail[i]);
    // Extrait le fichier sous son nom réel (significatif lors de l'envoi du mail)
    // et pas sous un tempname généré automatiqmt...
    FilePath := TmpPath + TEnreg.GetValue('YFI_FILENAME');
    V_GedFiles.Extract(FilePath , TEnreg.GetString('YMG_FILEGUID'));
    if Result<>'' then Result := Result + ';';
    Result := Result + FilePath;
    end;
end;


procedure TFNewMessage.VireFichiersJointsOnDisk(sFichiers: String);
var lesfichiers, tmpfile : String;
begin
  if sFichiers='' then exit;
  lesfichiers := sFichiers;
  tmpfile := ReadTokenSt(lesfichiers);
  While tmpfile<>'' do
    begin
    DeleteFile(tmpfile);
    tmpfile := ReadTokenSt(lesfichiers);
    end;
end;

procedure TFNewMessage.ActuZonesCalculees;
begin
  ActualiseNomEtTelephone(TobMsg.GetValue('YMS_NODOSSIER'));
  // Des précisions dans la bulle d'aide
  //  ALLDEST.Hint := DecodeDestinataires(TobMsg.GetValue ('YMS_ALLDEST'));
  //  ALLDESTCC.Hint := DecodeDestinataires(TobMsg.GetValue ('YMS_ALLDESTCC'));
end;

procedure TFNewMessage.ActualiseNomEtTelephone(NoDossier : String);
var
    Q : TQuery;
begin
  if NoDossier='' then
    begin
    ANN_NOM1.Caption := '';
    YMS_TEL.Text := '';
    end
  else
    begin
    Q := OpenSQL('SELECT ANN_NOM1, ANN_TEL1 FROM ANNUAIRE, DOSSIER WHERE ANN_GUIDPER=DOS_GUIDPER AND DOS_NODOSSIER="'+NoDossier+'"', True);
    if Not Q.Eof then
      begin
      ANN_NOM1.Caption := Q.FindField('ANN_NOM1').AsString;
      YMS_TEL.Text := Q.FindField('ANN_TEL1').AsString;
      end
    else
      begin
      ANN_NOM1.Caption := '';
      YMS_TEL.Text := '';
      end;
    Ferme(Q);
    end;
end;

function TFNewMessage.RepondreTransferer(TobOrig : Tob) : String;
// Retourne le corps du message d'origine,
// complété d'un entête pour la réponse ou le transfert
var
   sDateMail : string;
begin
  sDateMail := FormatDateTime('dd/mm/yyyy hh:nn:ss', TobOrig.GetValue('YMS_DATE'));
  Result := ''#$D#$A +
            '-----Message d''origine-----'#$D#$A +
            'De : ' + TobOrig.GetValue('YMS_LIBFROM') + ''#$D#$A +
            'A : '  + TobOrig.GetValue('YMS_ALLDEST') + ''#$D#$A +
            'CC : ' + TobOrig.GetValue('YMS_ALLDESTCC') + ''#$D#$A +
            'Envoyé : ' +  sDateMail + ''#$D#$A +
            'Objet : ' + TobOrig.GetValue('YMS_SUJET') + ''#$D#$A ;
  if IsRTF( TobOrig.Getvalue('YMS_MEMO') ) then
    Result := Result + GetRTFStringText( TobOrig.Getvalue('YMS_MEMO') )
  else
    Result := Result + TobOrig.Getvalue('YMS_MEMO');
end;

procedure TFNewMessage.AjoutePieceJointe(nomfichier: String);
var elt : TListItem;
    i : Integer;
begin
  if nomfichier='' then exit;

  // Fichier déjà inséré
  for i:=0 to LstFiles.Items.Count-1 do
    begin
    if UpperCase(ExtractFileName( LstFiles.Items[i].Caption ))
     = UpperCase(ExtractFileName( nomfichier )) then exit; // SORTIE
    end;

  // on rajoute simplement dans le listview (sans tob dans la propriété Data)
  // => au moment d'enregistrer les modifs, seules les lignes qui n'ont pas de tob
  // (donc qui ne correspondent pas à des enreg dans YFILES) seront traitées
  elt := LstFiles.Items.Add;
  elt.Caption := nomfichier;

  // après ajout, on doit voir l'onglet des pièces jointes
  PPiecesJointes.TabVisible := True;

  // on a modifié quelque chose
  bChange := True;
end;

procedure TFNewMessage.FormKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin

  case Key of

  // F10 : Envoyer
  VK_F10: begin key := 0; BEnvoyer.Click; end;

  // Ctrl + S : enregistrer
  83:     if Shift = [ssCtrl] then
            begin key := 0; BValider.Click; end;
  end;

end;


Initialization
  RegisterAglProc( 'EnvoiMessage', True, 4, AGLEnvoiMessage);

end.

