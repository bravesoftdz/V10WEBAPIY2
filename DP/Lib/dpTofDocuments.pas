{***********UNITE*************************************************
Auteur  ...... :  MD
Créé le ...... : 12/03/2003
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : DPDOCUMENTS ()
Mots clefs ... : TOF;DPDOCUMENTS
*****************************************************************}
Unit dpTofDocuments ;

Interface

Uses Windows,
     StdCtrls,
     Controls,
     Classes,
{$IFDEF EAGLCLIENT}
     MaineAGL, eMul, UTob,
{$ELSE}
     Fe_Main, Mul,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
{$ENDIF}
     HQry,
     forms,
     sysutils,
     ComCtrls,
     HDB,
     HCtrls,
     HEnt1,
     HMsgBox,
     HStatus,
     UTOF,
     Lookup,
     HTB97,
     UGedDP,
     Menus,
     UGedFiles,
     YNEwDocument,
     AnnOutils,
     utofzoneslibres, // $$$ JP 04/04/06
     CBPPath;

Type

  TDateExercice = Record
                   DateDebut  : TDateTime;
                   DateFin    : TDateTime;
                  end;

  TOF_DPDOCUMENTS = Class (TOF_ZONESLIBRES) // $$$ JP 04/04/06 (TOF)
    procedure OnClose; override ;
    procedure OnUpdate                 ; override ;
    procedure OnArgument ( S : String ) ; override ;


    procedure BEffacerCbArmoire_OnClick(Sender: TObject);
    procedure BEffacerCbClasseur_OnClick(Sender: TObject);
    procedure BEffacerCbIntercalaire_OnClick(Sender: TObject);

    procedure DPD_NODOSSIER_OnElipsisClick(Sender: TObject);
    procedure DPD_NODOSSIER_OnExit(Sender: TObject);
    procedure FListeOnDblClick(Sender: TObject);
    procedure CHKDocumentCabinet_OnClick  (Sender :TObject);
    procedure CODEGED1_OnChange(Sender: TObject);
    procedure CODEGED2_OnChange(Sender: TObject);
    procedure CODEGED3_OnChange(Sender: TObject);
    procedure Exercice_OnChange(Sender: TObject);
    procedure ArboGed_OnClick(Sender : Tobject);
    procedure BDELETE_OnClick(Sender: TObject);
    procedure BEXTRAIRE_OnClick (Sender : Tobject);
    function DonnerNomFichierGed (FileGuid : String) : String;

    function DonnerDroitDocument (DocGuid : String) : String;
    function DonnerLibelleDocument (DocGuid : String) : String;
    function DonnerEtatDocument (DocGuid : String) : String;

  private
    Qry : THQuery;
    Lst : THDBGrid;
    cbCodeGed1, cbCodeGed2, cbCodeGed3 : THValComboBox;
    cbExercice : THValComboBox;
    RbArboGed : THRadioGroup;
    XXWhereComplement : String;
    BRafraichir      : Boolean;

    procedure BPROPRIETE_OnClick(Sender: TObject);
    procedure BChercheClick(Sender: TObject);
    procedure Form_OnKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
    function  InitialiserXXWhere (Sender: TObject) : String;

    // $$$ JP 19/05/06 FQ 11038
    procedure BCurDossClick (Sender:TObject);
  end ;

Var TabExercice      : Array [1..NbMaxExercice] of TDateExercice;

Procedure InitialiserComboExercice (ComboExercice : TControl);

/////////// IMPLEMENTATION /////////////
Implementation

uses
  UDossierSelect,

  {$IFDEF VER150}
  Variants,
  {$ENDIF}

  DpJurOutils, galOutil, UtilGed, UtilMulTraitmt, GalMenuDisp,
  dpOutils;


procedure TOF_DPDOCUMENTS.OnUpdate ;
begin
  Inherited ;

  // $$$ JP 03/05/06 - FQ 11021: Activation du bouton supprimer APRES le OnUpdate de la classe ancêtre
  with TToolbarButton97 (GetControl ('BDELETE')) do
  begin
       Enabled := JaileDroitTag (ccSupprimerDocument);
       if Enabled = TRUE then
          OnClick := BDELETE_OnClick;
  end;
end ;

procedure TOF_DPDOCUMENTS.OnArgument ( S : String ) ;
begin
  Inherited ;
  // S reçoit 'GLOBAL' pour une recherche indépendante des dossiers

  // composants
  Qry := TFMul(Ecran).Q;
  Lst := TFMul(Ecran).FListe;

  // Rajout critère sur l'origine des documents
  XXWhereComplement:='';
  BRafraichir:=False;

  // No dossier
  if GetControl('DPD_NODOSSIER')<>Nil then
  begin
    THEdit(GetControl('DPD_NODOSSIER')).OnElipsisClick := DPD_NODOSSIER_OnElipsisClick;
    THEdit(GetControl('DPD_NODOSSIER')).OnExit := DPD_NODOSSIER_OnExit;
    if Pos ('GLOBAL', S) < 1 then
    begin
     SetControlText ('DPD_NODOSSIER', VH_DOSS.NoDossier); // $$$ JP 03/05/06 FQ 11021 - if S<>'GLOBAL' then SetControlText('DPD_NODOSSIER', VH_Doss.NoDossier);
     SetControlText ('DOS_LIBELLE', VH_DOSS.LibDossier);  // $$$ JP 17/05/06 FQ 11039 - il faut le libellé dès le début de la recherche
    end
    else
    begin
     // $$$ JP 19/05/06 FQ 11038: en recherche globale, inutile de voir le bouton "dossier en cours"
     SetControlEnabled ('BCURDOSS', FALSE);
     // CAT : on élimine les faux documents (c'est à dire les Commentaires)
     XXWhereComplement:=' AND YDF_DOCGUID IS NOT NULL';
     SetControlVisible ('EXERCICE', FALSE);
     SetControlVisible ('TEXERCICE', FALSE);
    end;
  end;

  if GetControl('CHKDOCUMENTCABINET')<>Nil then THCheckBox (GetControl('CHKDOCUMENTCABINET')).OnClick:=CHKDocumentCabinet_OnClick;

  if GetControl('BEFFACERCBARMOIRE')<>Nil then TToolbarButton97 (GetControl('BEFFACERCBARMOIRE')).OnClick:=BEffacerCbArmoire_OnClick;
  if GetControl('BEFFACERCBCLASSEUR')<>Nil then TToolbarButton97 (GetControl('BEFFACERCBCLASSEUR')).OnClick:=BEffacerCbClasseur_OnClick;
  if GetControl('BEFFACERCBINTERCALAIRE')<>Nil then TToolbarButton97 (GetControl('BEFFACERCBINTERCALAIRE')).OnClick:=BEffacerCbIntercalaire_OnClick;

  //--- CAT : 02/06/2005 Propriété du document
  if (GetControl ('BPROPRIETE')<>Nil) then
   begin
    if JaiLeDroitTag (CcModifierDocument) then
     begin
      TToolBarButton97(GetControl('BPropriete')).OnClick:=BPROPRIETE_OnClick;
      TMenuItem(GetControl('mPropriete')).OnClick:=BPROPRIETE_OnClick;
     end
    else
     begin
      TToolBarButton97(GetControl('BPropriete')).Enabled:=False;
      TMenuItem(GetControl('mPropriete')).Enabled:=False;
     end;
   end;

  if (GetControl ('BEXTRAIRE')<>Nil) then
   begin
    if (JaiLeDroitConceptBureau (ccExtraireDocument) = FALSE) then
     TToolBarButton97(GetControl('BEXTRAIRE')).Enabled:=False
    else
     TToolBarButton97(GetControl('BEXTRAIRE')).OnClick:=BEXTRAIRE_OnClick;
   end;

  TToolBarButton97(GetControl('BCHERCHE')).OnClick := BCHERCHEClick;

  //--- CAT : 01/06/2005 Armoire, classeur, Intercalaire
  cbCodeGed1 := THValComboBox(GetControl('CODEGED1'));
  cbCodeGed1.OnChange := CODEGED1_OnChange;
  cbCodeGed2 := THValComboBox(GetControl('CODEGED2'));
  cbCodeGed2.OnChange := CODEGED2_OnChange;
  cbCodeGed3 := THValComboBox(GetControl('CODEGED3'));
  cbCodeGed3.OnChange := CODEGED3_OnChange;

  //--- CAT : 30/01/2007 Gestion des exercices
  cbExercice := THValComboBox(GetControl('EXERCICE'));
  InitialiserComboExercice (ThValComboBox(GetControl('EXERCICE')));
  cbExercice.OnChange := EXERCICE_OnChange;

  RbArboGed := THRadioGroup (GetControl('DPD_ARBOGED'));
  RbArboGed.OnClick:=ArboGed_OnClick;

  if (THRadioGroup (GetControl ('DPD_ARBOGED')).ItemIndex=0) then
   THValComboBox (GetControl ('CodeGed1')).Plus:='D'
  else
   if (THRadioGroup (GetControl ('DPD_ARBOGED')).ItemIndex=1) then
    THValComboBox (GetControl ('CodeGed1')).Plus:='A';

  THDBGrid(GetControl('FLISTE')).OnDblClick := FListeOnDblClick;
  Ecran.OnKeyDown := Form_OnKeyDown;

  // Noms des zones libres assistants
  AfficheLibTablesLibres (Self);
  // FQ 11555 SetControlProperty ('PINFOSLIBRES', 'TabVisible', GetControl ('YDO_LIBREGED1').Visible or GetControl ('YDO_LIBREGED2').Visible or GetControl ('YDO_LIBREGED3').Visible or GetControl ('YDO_LIBREGED4').Visible or GetControl ('YDO_LIBREGED5').Visible);

  // $$$ JP 19/05/06 FQ 11038: Bouton de reprise du n° de dossier en cours
  TToolBarButton97 (GetControl ('BCURDOSS')).OnClick := BCurDossClick;

  // FQ 11554
  SetControlProperty ('YDO_LIBELLEDOC', 'MaxLength', 70);
  SetControlProperty ('YDO_AUTEUR',     'MaxLength', 35);
end ;

//----------------------------
//--- Nom : BPropriete_Onclick
//----------------------------
procedure TOF_DPDOCUMENTS.BPROPRIETE_OnClick(Sender: TObject);
var sResultDoc_l : String;
    ParGed       : TParamGedDoc;
    Q : TQuery;
begin
 if VarIsNull(GetField('DPD_DOCGUID')) then exit;
 ParGed.SDocGUID  := GetField('DPD_DOCGUID');
 // On peut être positionné sur un commentaire ou une url => dans ce cas pas de fileguid
 if VarIsNull(GetField('YDF_FILEGUID')) then
  begin
   ParGed.SFileGUID := '';
   Q := OpenSQL('SELECT YDO_DOCTYPE FROM YDOCUMENTS WHERE YDO_DOCGUID="'+GetField('DPD_DOCGUID')+'"', True);
   if Not Q.Eof then
    begin
     if Q.FindField('YDO_DOCTYPE').AsString='URL' then
      ParGed.TypeGed := 'URL'
     else
      ParGed.TypeGed := 'COM';
    end
   else
    ParGed.TypeGed := 'COM';
   Ferme(Q);
  end
 else
  begin
   ParGed.SFileGUID := GetField('YDF_FILEGUID');
   ParGed.TypeGed   := 'DOC';
  end;
 ParGed.NoDossier := GetField('DPD_NODOSSIER');
 ParGed.CodeGed   := GetField('DPD_CODEGED');
 sResultDoc_l := ShowNewDocument(ParGed);
 TFMul(Ecran).BChercheClick(Sender);
 BRafraichir:=True;
end;


procedure TOF_DPDOCUMENTS.DPD_NODOSSIER_OnElipsisClick(Sender: TObject);
var St, sGuidper : String;
begin
  // retourne NoDossier;GuidPer;Nom1
  St := AGLLanceFiche('YY','YYDOSSIER_SEL', '','',GetControlText('DPD_NODOSSIER'));
  if St<>'' then
    begin
    SetControlText('DPD_NODOSSIER', READTOKENST(St));
    sGuidper := READTOKENST(St);
    SetControlCaption('DOS_LIBELLE', READTOKENST(St));
    // ou GetNomCompPer(Guidper)
    end
  else
    begin
    SetControlText('DPD_NODOSSIER', '');
    SetControlCaption('DOS_LIBELLE', '');
    end;
end;

procedure TOF_DPDOCUMENTS.DPD_NODOSSIER_OnExit(Sender: TObject);
var sGuidPer : String;
begin
    sGuidper := GetGuidPer(GetControlText('DPD_NODOSSIER'));
    SetControlCaption('DOS_LIBELLE', GetNomCompPer(sGuidper));
end;

procedure TOF_DPDOCUMENTS.FListeOnDblClick(Sender: TObject);
begin
  inherited;
  if (Not VarIsNull(GetField('DPD_CODEGED'))) and
     ((JaileDroitTag (ccConsulterDocumentAutre)) or (VarToStr (GetField('YDO_UTILISATEUR'))=V_PGI.User)) then
   begin
    TFMul(Ecran).Retour := VarToStr(GetField('DPD_CODEGED')) + ';'
                         + VarToStr(GetField('DPD_NODOSSIER')) + ';'
                         + VarToStr(GetField('YDF_FILEGUID')) + ';'
                         + VarToStr(GetField('DPD_DOCGUID')) ;
   end;
  TFMul(Ecran).ModalResult := mrOk; // ou Self.Close
end;


procedure TOF_DPDOCUMENTS.CODEGED1_OnChange(Sender: TObject);
begin
  SetControlText('DPD_CODEGED', cbCodeGed1.Value);
end;

procedure TOF_DPDOCUMENTS.CODEGED2_OnChange(Sender: TObject);
begin
  SetControlText('DPD_CODEGED', cbCodeGed2.Value);
end;

procedure TOF_DPDOCUMENTS.CODEGED3_OnChange(Sender: TObject);
begin
  SetControlText('DPD_CODEGED', cbCodeGed3.Value);
end;

procedure TOF_DPDOCUMENTS.EXERCICE_OnChange(Sender: TObject);
var Annee,Mois,Jour : Word;
begin
 if (CbExercice.Value<>'') then
  begin
   DecodeDate (TabExercice [StrToInt (CbExercice.Value)].DateDebut,Annee,Mois,Jour);
   SetControlText ('ANNEEDEB',IntToStr (Annee)); SetControlText ('MOISDEB',IntToStr (Mois));
   DecodeDate (TabExercice [StrToInt (CbExercice.Value)].DateFin,Annee,Mois,Jour);
   SetControlText ('ANNEEFIN',IntToStr(Annee)); SetControlText ('MOISFIN',IntToStr(Mois));
  end
 else
  begin
   SetControlText ('ANNEEDEB','1900'); SetControlText ('MOISDEB','01');
   SetControlText ('ANNEEFIN','2099'); SetControlText ('MOISFIN','12');
  end;
end;

procedure TOF_DPDOCUMENTS.Form_OnKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  case Key of

  // F11 = affichage du popup
  VK_F11 :  begin
            TPopupMenu(GetControl('PopUpDocument')).Popup(Mouse.CursorPos.X, Mouse.CursorPos.Y);
            // évite de traiter le F11 par défaut de Pgi (=vk_choixmul=double clic ds grid)
            Key := 0;
            end;

  // Suppression : Ctrl + Suppr
  VK_DELETE :
       if (Shift = [ssCtrl]) and (TFMul(Ecran).FListe.Focused) and GetControl('BDELETE').Visible then
         BDELETE_OnClick(Nil);

  else
       TFMul(Ecran).FormKeyDown(Sender, Key, Shift);
  end;
end;

procedure TOF_DPDOCUMENTS.BDELETE_OnClick(Sender: TObject);
var Msg : String;
    i : Integer;
begin
  if (Lst.nbSelected=0) and (Not Lst.AllSelected) then
    begin
    PGIInfo('Aucun document sélectionné.', TitreHalley);
    exit;
    end;

  Msg := 'Vous allez supprimer la référence de tous les documents sélectionnés.'+#13#10
   + ' (Rq : si un document est référencé par ailleurs, il ne sera pas supprimé de la GED).'+#13#10
   + ' Confirmez-vous la suppression ?';

  if PGIAsk(Msg, TitreHalley)=mrNo then exit;

  // liste des documents à supprimer
  if Lst.AllSelected then
    BEGIN
{$IFDEF EAGLCLIENT}
    if not TFMul(Ecran).Fetchlestous then
      PGIInfo('Impossible de récupérer tous les enregistrements')
    else
{$ENDIF}
      begin
      Qry.First;
      while Not Qry.EOF do
        begin
        // Suppression en cascade dans les tables, avec tests de dépendance
        if (DonnerDroitDocument (Qry.FindField('DPD_DOCGUID').AsString)='RO') then
         PgiInfo('Le document '+DonnerLibelleDocument (Qry.FindField('DPD_DOCGUID').AsString)+' est en lecture seule, impossible de le supprimer.', TitreHalley)
        else
         SupprimeDocumentGed(Qry.FindField('DPD_DOCGUID').AsString);
        Qry.Next;
        end;
      end;
    END
  else
    BEGIN
    InitMove(Lst.NbSelected,'');
    for i:=0 to Lst.NbSelected-1 do
      begin
      MoveCur(False);
      Lst.GotoLeBookmark(i);
{$IFDEF EAGLCLIENT}
      Qry.TQ.Seek(Lst.Row - 1) ;
{$ENDIF}

      if (DonnerDroitDocument (Qry.FindField('DPD_DOCGUID').AsString)='RO') then
       PgiInfo('Le document '+DonnerLibelleDocument (Qry.FindField('DPD_DOCGUID').AsString)+' est en lecture seule, impossible de le supprimer.', TitreHalley)
      else
       if (DonnerEtatDocument (Qry.FindField('DPD_DOCGUID').AsString)='EXT') then
        PgiInfo('Le document '+DonnerLibelleDocument (Qry.FindField('DPD_DOCGUID').AsString)+' a été extrait, impossible de le supprimer.', TitreHalley)
       else
        SupprimeDocumentGed(Qry.FindField('DPD_DOCGUID').AsString)
      end;
    FiniMove;
    END;
  // déselectionne
  FinTraitmtMul(TFMul(Ecran));
  // actualisation
  TFMul(Ecran).BChercheClick(Sender);
  BRafraichir:=True;
end;

// $$$ JP 19/05/06 FQ 11038
procedure TOF_DPDOCUMENTS.BCurDossClick (Sender:TObject);
begin
     SetControlText    ('DPD_NODOSSIER', VH_DOSS.NoDossier);
     SetControlCaption ('DOS_LIBELLE',   VH_DOSS.LibDossier);
end;

//--- CAT 24/11/2006
procedure TOF_DPDOCUMENTS.ArboGed_OnClick (Sender : Tobject);
begin

 if (THRadioGroup (GetControl ('DPD_ARBOGED')).ItemIndex=0) then
  THValComboBox (GetControl ('CodeGed1')).Plus:='D'
 else
  if (THRadioGroup (GetControl ('DPD_ARBOGED')).ItemIndex=1) then
   THValComboBox (GetControl ('CodeGed1')).Plus:='A';

 THValComboBox (GetControl ('CodeGed1')).Value := '';
 THValComboBox (GetControl ('CodeGed2')).Value := '';
 THValComboBox (GetControl ('CodeGed2')).Values.Clear;
 THValComboBox (GetControl ('CodeGed2')).Items.Clear;
 THValComboBox (GetControl ('CodeGed3')).Value := '';
 THValComboBox (GetControl ('CodeGed3')).Values.Clear;
 THValComboBox (GetControl ('CodeGed3')).Items.Clear;
end;

//--- CAT 08/01/2007
procedure TOF_DPDOCUMENTS.BEXTRAIRE_OnClick (Sender : Tobject);
var CheminFichierGed : String;
    NomFichierGed    : String;
    DroitDoc         : String;
    EtatDoc          : String;
    i                : Integer;
begin
 if (Lst.nbSelected=0) and (Not Lst.AllSelected) then
  begin
   PGIInfo('Aucun document sélectionné.', TitreHalley);
   exit;
  end;

 // liste des documents à extraire
 if Lst.AllSelected then
   begin
   {$IFDEF EAGLCLIENT}
    if not TFMul(Ecran).Fetchlestous then
     PGIInfo('Impossible de récupérer tous les enregistrements')
    else
   {$ENDIF}
     begin
      Qry.First;
      while Not Qry.EOF do
        begin
         CheminFichierGed:=TCbpPath.GetCegidUserRoamingAppData+'\GED\'+Qry.FindField ('DPD_DOCGUID').AsString;
         NomFichierGed:=DonnerNomFichierGed (Qry.FindField ('YDF_FILEGUID').AsString);
         if not DirectoryExists (CheminFichierGed) then
         if not ForceDirectories(CheminFichierGed) then exit;

         DroitDoc:=DonnerDroitDocument (Qry.FindField('DPD_DOCGUID').AsString);
         EtatDoc:=DonnerEtatDocument (Qry.FindField('DPD_DOCGUID').AsString);
         if (DroitDoc='RO') or (DroitDoc='NE') then
          PgiInfo('Impossible d''extraire le document '+DonnerLibelleDocument (Qry.FindField('DPD_DOCGUID').AsString)+#13#10
                + 'Vous n''avez pas les droits d''extraction pour ce document, ou bien il est en lecture seule.', TitreHalley)
         else
          if (EtatDoc='EXT') then
           PgiInfo('Le document '+DonnerLibelleDocument (Qry.FindField('DPD_DOCGUID').AsString)+' a déjà été extrait.')
          else
           begin
            if (V_GedFiles.Extract(CheminFichierGed+'\'+NomFichierGed, Qry.FindField ('YDF_FILEGUID').AsString)) then
             ExecuteSql ('UPDATE YDOCUMENTS SET YDO_DOCSTATE="EXT", YDO_DATEMODIF="'+UsDateTime (Now)+'", YDO_UTILISATEUR="'+V_PGI.USER+'" WHERE YDO_DOCGUID="'+Qry.FindField ('DPD_DOCGUID').AsString+'"')
            else
             PgiInfo('Impossible d''extraire le document "'+Qry.FindField ('YDO_LIBELLEDOC').AsString+'".', TitreHalley);
           end;

         Qry.Next;
        end;

     end;

   end
  else
    BEGIN
     InitMove(Lst.NbSelected,'');
     for i:=0 to Lst.NbSelected-1 do
       begin
        MoveCur(False);
        Lst.GotoLeBookmark(i);
        {$IFDEF EAGLCLIENT}
         Qry.TQ.Seek(Lst.Row - 1) ;
        {$ENDIF}

         CheminFichierGed:=TCbpPath.GetCegidUserRoamingAppData+'\GED\'+Qry.FindField ('DPD_DOCGUID').AsString;
         NomFichierGed:=DonnerNomFichierGed (Qry.FindField ('YDF_FILEGUID').AsString);
         if not DirectoryExists (CheminFichierGed) then
         if not ForceDirectories(CheminFichierGed) then exit;

         DroitDoc:=DonnerDroitDocument (Qry.FindField('DPD_DOCGUID').AsString);
         EtatDoc:=DonnerEtatDocument (Qry.FindField('DPD_DOCGUID').AsString);

         if (DroitDoc='RO') or (DroitDoc='NE') then
          PgiInfo('Impossible d''extraire le document '+DonnerLibelleDocument (Qry.FindField('DPD_DOCGUID').AsString)+#13#10
           +'Vous n''avez pas les droits d''extraction pour ce document, ou bien il est en lecture seule.', TitreHalley)
         else
          if (EtatDoc='EXT') then
           PgiInfo('Le document '+DonnerLibelleDocument (Qry.FindField('DPD_DOCGUID').AsString)+' a déjà été extrait.')
          else
           begin
            if (V_GedFiles.Extract(CheminFichierGed+'\'+NomFichierGed, Qry.FindField ('YDF_FILEGUID').AsString)) then
             ExecuteSql ('UPDATE YDOCUMENTS SET YDO_DOCSTATE="EXT", YDO_DATEMODIF="'+UsDateTime (Now)+'", YDO_UTILISATEUR="'+V_PGI.USER+'" WHERE YDO_DOCGUID="'+Qry.FindField ('DPD_DOCGUID').AsString+'"')
            else
             PgiInfo('Impossible d''extraire le document "'+Qry.FindField ('YDO_LIBELLEDOC').AsString+'".', TitreHalley);
           end;
       end;
     FiniMove;
    END;

  PgiInfo ('Opération terminée.',TitreHalley);

  // déselectionne
  FinTraitmtMul(TFMul(Ecran));
  // actualisation
  TFMul(Ecran).BChercheClick(Sender);
  BRafraichir:=True;
end;

//--- CAT 08/01/2007
function TOF_DPDOCUMENTS.DonnerNomFichierGed (FileGuid : String) : String;
var SSql     : String;
    RSql     : TQuery;
begin
 Result:='';
 SSql:='SELECT YFI_FILENAME FROM YFILES WHERE YFI_FILEGUID="'+FileGuid+'"';
 RSql:=OpenSql (SSql,False);
 if (not RSql.Eof) then
  Result:=RSql.FindField ('YFI_FILENAME').AsString;
 Ferme (RSql);
end;

//----------------------------------
//--- Nom : DonnerEtatDocument
//----------------------------------
function TOF_DPDOCUMENTS.DonnerEtatDocument (DocGuid : String) : String;
var SSql     : String;
    RSql     : TQuery;
begin
 Result:='';
 SSql:='SELECT YDO_DOCSTATE FROM YDOCUMENTS WHERE YDO_DOCGUID="'+DocGuid+'"';
 RSql:=OpenSql (SSql,False);
 if (not RSql.Eof) then
  Result:=RSql.FindField ('YDO_DOCSTATE').AsString;
 Ferme (RSql);
end;

//----------------------------------
//--- Nom : DonnerDroitDocument
//----------------------------------
function TOF_DPDOCUMENTS.DonnerDroitDocument (DocGuid : String) : String;
var SSql     : String;
    RSql     : TQuery;
begin
 Result:='';
 SSql:='SELECT YDO_DROITGED FROM YDOCUMENTS WHERE YDO_DOCGUID="'+DocGuid+'"';
 RSql:=OpenSql (SSql,False);
 if (not RSql.Eof) then
  Result:=RSql.FindField ('YDO_DROITGED').AsString;
 Ferme (RSql);
end;

//----------------------------------
//--- Nom : DonnerLibelleDocument
//----------------------------------
function TOF_DPDOCUMENTS.DonnerLibelleDocument (DocGuid : String) : String;
var SSql     : String;
    RSql     : TQuery;
begin
 Result:='';
 SSql:='SELECT YDO_LIBELLEDOC FROM YDOCUMENTS WHERE YDO_DOCGUID="'+DocGuid+'"';
 RSql:=OpenSql (SSql,False);
 if (not RSql.Eof) then
  Result:=RSql.FindField ('YDO_LIBELLEDOC').AsString;
 Ferme (RSql);
end;

//------------------------------------
//--- Nom : BEffacerCbArmoire_Click
//------------------------------------
procedure TOF_DPDOCUMENTS.BEffacerCbArmoire_OnClick(Sender: TObject);
begin
 THValComboBox (GetControl ('CodeGed1')).Value := '';
end;

//-------------------------------------
//--- Nom : BEffacerCbClasseur_Click
//-------------------------------------
procedure TOF_DPDOCUMENTS.BEffacerCbClasseur_OnClick(Sender: TObject);
begin
 THValComboBox (GetControl ('CodeGed2')).Value := '';
end;

//-----------------------------------------
//--- Nom : BEffacerCbIntercalaire_Click
//-----------------------------------------
procedure TOF_DPDOCUMENTS.BEffacerCbIntercalaire_OnClick(Sender: TObject);
begin
 THValComboBox (GetControl ('CodeGed3')).Value := '';
end;

//---------------------------------------
//--- Nom : CHKDocumentCabinet_Onclick
//---------------------------------------
procedure TOF_DPDOCUMENTS.CHKDocumentCabinet_OnClick(Sender: TObject);
begin
 GetControl('DPD_NODOSSIER').Enabled:=not (THCheckBox (GetControl('CHKDOCUMENTCABINET')).Checked);
 GetControl('TDOSSIER').Enabled:=not (THCheckBox (GetControl('CHKDOCUMENTCABINET')).Checked);
 THEdit (GetControl('DPD_NODOSSIER')).Text:='';
end;

//------------------------------------
//--- Nom : InitialiserComboExercice
//------------------------------------
procedure InitialiserComboExercice (ComboExercice : TControl);
var DateDebEx, DateFinEx : TDateTime;
    DureeExercice        : Integer;
    NumExercice,Indice   : Integer;
begin
 if Not (ComboExercice is THComboBox) then exit;
 DonnerDateExercice (DateDebEx,DateFinEx,DureeExercice);
 THValCombobox (ComboExercice).Items.Clear;

 THValCombobox (ComboExercice).Items.add ('<<Tous>>');
 THValCombobox (ComboExercice).Values.add ('');

 for Indice:=1 DownTo -2 do
  begin
   NumExercice:=Indice+3;
   TabExercice [NumExercice].DateDebut:=PlusMois (DateDebEx,(Indice*DureeExercice));
   TabExercice [NumExercice].DateFin:=PlusMois (DateFinEx,(Indice*DureeExercice));
   THValCombobox (ComboExercice).Items.add ('Du '+DateToStr (TabExercice [NumExercice].DateDebut)+' au '+DateToStr (TabExercice [NumExercice].DateFin));
   THValCombobox (ComboExercice).Values.add (IntToStr (NumExercice));
  end;

 THValCombobox (ComboExercice).Text:=THValCombobox (ComboExercice).Items[2];
end;

//--------------------------
//--- Nom : BChercheClick
//--------------------------
procedure TOF_DPDOCUMENTS.BCHERCHEClick(Sender: TObject);
begin
 SetControlText('XX_WHERE', '('+InitialiserXXWhere (Sender)+')');
 TFMul(Ecran).BChercheClick(Sender);
end;

//--------------------------
//--- Nom : OnClose
//--------------------------
procedure TOF_DPDOCUMENTS.OnClose;
begin
 inherited ;
 if (BRafraichir) then FMenuDisp.GedDp.AlimTreeViewGedDp;
end ;

//-------------------------------
//--- Nom : InitialiserXXWhere
//-------------------------------
function TOF_DPDOCUMENTS.InitialiserXXWhere (Sender: TObject) : String;
var DateDebEx, DateFinEx      : TDateTime;
    ChDocCabinet, ChDocPrive  : String;
    ChDocEtat, ChDocAnneeMois : String;
    ChDocDroit                : String;
begin
 ChDocCabinet:='';ChDocPrive:='';ChDocEtat:='';ChDocDroit:='';ChDocAnneeMois:='';

 //--- Gestion des critères DATE
 if (GetControl('AnneeDeb')<>nil) and (GetControl('MoisDeb')<>nil) and (GetControl('AnneeFin')<>nil) and (GetControl('MoisFin')<>nil) then
  begin
   DateDebEx:=DebutDeMois (StrToDate ('01/'+IntToStr (THSpinEdit(GetControl('MOISDEB')).Value)+'/'+IntToStr(THSpinEdit(GetControl('ANNEEDEB')).Value)));
   DateFinEx:=FinDeMois(StrToDate ('01/'+IntToStr (THSpinEdit(GetControl('MOISFIN')).Value)+'/'+IntToStr(THSpinEdit(GetControl('ANNEEFIN')).Value)));

   ChDocAnneeMois:=' AND ("'+FormatDateTime ('yyyymmdd', DateDebEx) + '"<=YDO_ANNEE||YDO_MOIS||"01" AND YDO_ANNEE||YDO_MOIS||"01"<="' + FormatDateTime ('yyyymmdd', DateFinEx) + '")';
  end;

 //--- Gestion critère DOCUMENT CABINET
 if (GetControl ('CHKDOCUMENTCABINET')<>nil) then
  begin
   if (THCheckBox (GetControl ('CHKDOCUMENTCABINET')).Checked) then
    ChDocCabinet:=' AND DPD_CODEGED<>"###" AND DPD_NODOSSIER=""';
  end;

 //--- Gestion critère DOCUMENT PRIVE
 if (GetControl ('CHKDOCUMENTPRIVE')<>nil) then
  begin
   if (THCheckBox (GetControl ('CHKDOCUMENTPRIVE')).State=CbGrayed) then
    ChDocPrive:=' AND ((DPD_UTILISATEUR="'+V_PGI.User+'" AND YDO_PRIVE="X") OR YDO_PRIVE="-")'
   else
    begin
     if (THCheckBox (GetControl ('CHKDOCUMENTPRIVE')).Checked) then
      ChDocPrive:=' AND (DPD_UTILISATEUR="'+V_PGI.User+'" AND YDO_PRIVE="X")'
     else
      ChDocPrive:=' AND (YDO_PRIVE="-")';
    end;
  end;

 //--- Gestion Critère DROIT DOCUMENT
 if (GetControl ('LISTEDROIT')<>nil) then
  begin
   //--- Sans restriction
   if (THValComboBox(GetControl('LISTEDROIT')).Value='') and (THValComboBox(GetControl('LISTEDROIT')).Text<>'<<Tous>>') and (THValComboBox(GetControl('LISTEDROIT')).Text<>'') then
    ChDocDroit:= ' AND (YDO_DROITGED="")';
   //--- Non extratible
   if (THValComboBox(GetControl('LISTEDROIT')).Value='NE') then ChDocDroit:= ' AND (YDO_DROITGED="NE")';
   //--- En lecture seule
   if (THValComboBox(GetControl('LISTEDROIT')).Value='RO') then ChDocDroit:= ' AND (YDO_DROITGED="RO")';
  end;

 //--- Gestion critère ETAT DOCUMENT
 if (GetControl ('LISTEETAT')<>nil) then
  begin
   //--- Actif
   if (THValComboBox(GetControl('LISTEETAT')).Value='1') then ChDocEtat:= ' AND ((YDO_DOCSTATE="") AND (YDO_DATEFIN>="'+UsDateTime (Now)+'"))';
   //--- Vérouillé
   if (THValComboBox(GetControl('LISTEETAT')).Value='2') then ChDocEtat:= ' AND (YDO_DOCSTATE<>"")';
   //--- Périmé
   if (THValComboBox(GetControl('LISTEETAT')).Value='3') then ChDocEtat:= ' AND (YDO_DATEFIN <"'+UsDateTime (Now)+'")';
  end;

  Result:='('+GererCritereGroupeConfTous+') AND (DPD_CODEGED<>"###")'+XXWhereComplement+ChDocAnneeMois+ChDocCabinet+ChDocPrive+ChDocDroit+ChDocEtat;
end;

initialization
  registerclasses ( [ TOF_DPDOCUMENTS ] ) ;
end.

