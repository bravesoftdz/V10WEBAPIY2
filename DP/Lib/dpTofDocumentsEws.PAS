{***********UNITE*************************************************
Auteur  ...... :  MP
Créé le ...... : 30/06/2004
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : DPDOCUMENTSEWS ()
Mots clefs ... : TOF;DPDOCUMENTSEWS
*****************************************************************}
Unit dpTofDocumentsEws ;

Interface

Uses Windows,
     StdCtrls,
     Controls,
     Classes,

     db,dbtables, //mp a suppr

{$IFDEF EAGLCLIENT}
     MaineAGL, eMul,
{$ELSE}
     Fe_Main, Mul,
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
     uHttpCS,
     uHttp,
     Paramsoc;

Type
  TOF_DPDOCUMENTSEWS = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
  private
    Qry : THQuery;
    Lst : THDBGrid;
    cbCodeGed1, cbCodeGed2 : THValComboBox;
    chkDejaPublie : TCheckBox;
    procedure DPD_NODOSSIER_OnElipsisClick(Sender: TObject);
    procedure LibEws_OnElipsisClick(Sender: TObject);
    procedure BEffaceEws_OnClick(Sender: TObject);
    procedure FListe_OnDblClick(Sender: TObject);
    procedure CODEGED1_OnChange(Sender: TObject);
    procedure CODEGED2_OnChange(Sender: TObject);
    procedure CHKDEJAPUBLIE_OnClick(Sender: TObject);
    // procedure BPUBLICATION_OnClick(Sender: TObject);
    // procedure BANNULAPUBLIER_OnClick(Sender: TObject);
    procedure FiltreDates;
    procedure MarquageAPublier(bMarque: Boolean);
    procedure BOuvrir_OnClick(Sender: TObject);
  end ;


/////////// IMPLEMENTATION /////////////
Implementation

uses
  {$IFDEF EWS}
  UtileWS,
  {$ENDIF}
  DpJurOutils, galOutil, UtilGed, UtilMulTraitmt, galTraitementCourant;


procedure TOF_DPDOCUMENTSEWS.OnNew ;
begin
  Inherited ;
end ;


procedure TOF_DPDOCUMENTSEWS.OnDelete ;
begin
  Inherited ;
end ;


procedure TOF_DPDOCUMENTSEWS.OnUpdate ;
begin
  Inherited ;
end ;


procedure TOF_DPDOCUMENTSEWS.OnLoad ;
begin
  Inherited ;
  TPageControl(GetControl('Pages')).ActivePageIndex := 0;
end ;


procedure TOF_DPDOCUMENTSEWS.OnArgument (S : String ) ;
var ChXXWhere : String;
begin
  Inherited ;
  // composants
  Qry := TFMul(Ecran).Q;
  Lst := TFMul(Ecran).FListe;

  // Critère issu des groupes de travail
  ChXXWhere:=GererCritereGroupeConfTous ();
  SetControlText('XX_WHERE', '('+ChXXWhere+') AND DPD_CODEGED<>"###"');

  // No dossier
  THEdit(GetControl('DPD_NODOSSIER')).OnElipsisClick := DPD_NODOSSIER_OnElipsisClick;

  // Case "Déja Publiés" (documents)
  chkDejaPublie := TCheckBox(GetControl('CHKPUBLIE'));
  chkDejaPublie.OnClick := CHKDEJAPUBLIE_OnClick;

  // Case "A Publier" (documents)
  // #### pourquoi il fait <>'X' ????
  // SetControlVisible('DPD_EWSAPUBLIER', False);

  // Identifiant Ews
  THEdit(GetControl('LIBEWS')).OnElipsisClick := LibEws_OnElipsisClick;

  // Armoire, classeur
  cbCodeGed1 := THValComboBox(GetControl('CODEGED1'));
  cbCodeGed1.OnChange := CODEGED1_OnChange;
  cbCodeGed2 := THValComboBox(GetControl('CODEGED2'));
  cbCodeGed2.OnChange := CODEGED2_OnChange;
  THDBGrid(GetControl('FLISTE')).OnDblClick := FListe_OnDblClick;

  // #### Non utilisés pour l'instant
  SetControlVisible('BPUBLICATION', False);
  SetControlVisible('BANNULAPUBLIER', False);
  // TToolbarButton97(GetControl('BPUBLICATION')).OnClick := BPUBLICATION_OnClick;
  // TToolbarButton97(GetControl('BANNULAPUBLIER')).OnClick := BANNULAPUBLIER_OnClick;
  TToolBarButton97(GetControl('BEffaceEws')).OnClick := BEffaceEws_OnClick;
  TToolBarButton97(GetControl('BOuvrir')).OnClick := BOuvrir_OnClick;

  FiltreDates;
end ;


procedure TOF_DPDOCUMENTSEWS.OnClose ;
begin
  Inherited ;
end ;


procedure TOF_DPDOCUMENTSEWS.OnDisplay () ;
begin
  Inherited ;
end ;


procedure TOF_DPDOCUMENTSEWS.OnCancel () ;
begin
  Inherited ;
end ;


procedure TOF_DPDOCUMENTSEWS.DPD_NODOSSIER_OnElipsisClick(Sender: TObject);
var St, codper : String;
begin
  // retourne NoDossier;GuidPer;Nom1
  St := AGLLanceFiche('YY','YYDOSSIER_SEL', '','',GetControlText('DPD_NODOSSIER'));
  if St<>'' then
    begin
    SetControlText('DPD_NODOSSIER', READTOKENST(St));
    codper := READTOKENST(St);
    SetControlCaption('DOS_LIBELLE', READTOKENST(St));
    // ou GetNomCompPer(codper)
    end
  else
    begin
    SetControlText('DPD_NODOSSIER', '');
    SetControlCaption('DOS_LIBELLE', '');
    end;
end;


procedure TOF_DPDOCUMENTSEWS.LibEws_OnElipsisClick(Sender: TObject);
var EwsId, EwsLib : String;
begin
  {$IFDEF EWS}
  // Appel objet com pour sélection d'un noeud eWS
  if EwsSelectionneNoeud(EwsId, EwsLib) then
    begin
    SetControlText('DPD_EWSID', EwsId);
    SetControlText('LIBEWS', EwsLib);
    end
  else
    begin
    SetControlText('DPD_EWSID', '');
    SetControlText('LIBEWS', '');
    end;
  {$ENDIF}
end;


procedure TOF_DPDOCUMENTSEWS.BEffaceEws_OnClick(Sender: TObject);
begin
  SetControlText('DPD_EWSID', '');
  SetControlText('LIBEWS', '');
end;

procedure TOF_DPDOCUMENTSEWS.FListe_OnDblClick(Sender: TObject);
begin
  inherited;
end;


procedure TOF_DPDOCUMENTSEWS.CHKDEJAPUBLIE_OnClick(Sender: TObject);
begin
  FiltreDates;
end;


procedure TOF_DPDOCUMENTSEWS.CODEGED1_OnChange(Sender: TObject);
begin
  SetControlText('DPD_CODEGED', cbCodeGed1.Value);
end;


procedure TOF_DPDOCUMENTSEWS.CODEGED2_OnChange(Sender: TObject);
begin
  SetControlText('DPD_CODEGED', cbCodeGed2.Value);
end;


{procedure TOF_DPDOCUMENTSEWS.BPUBLICATION_OnClick(Sender: TObject);
begin
  MarquageAPublier(True);
end;

procedure TOF_DPDOCUMENTSEWS.BANNULAPUBLIER_OnClick(Sender: TObject);
begin
  MarquageAPublier(False);
end; }

procedure TOF_DPDOCUMENTSEWS.FiltreDates;
Var ChXXWhere : String;
begin
 ChXXWhere:=GererCritereGroupeConfTous ();
 if chkDejaPublie.State = cbGrayed then SetControlText('XX_WHERE', ChXXWhere)
 else if chkDejaPublie.State = cbChecked then SetControlText('XX_WHERE', '('+ChXXWhere+') AND DPD_EWSDATEPUBL<>"'+UsDateTime(iDate1900)+'"')
 else if chkDejaPublie.State = cbUnchecked then SetControlText('XX_WHERE', '('+ChXXWhere+') AND DPD_EWSDATEPUBL="'+UsDateTime(iDate1900)+'"');
end;

procedure TOF_DPDOCUMENTSEWS.MarquageAPublier(bMarque: Boolean);
// bMarque :
// True => pour publier;
// False => pour annuler la publication
var Msg : String;
    i : Integer;
begin
  if (Lst.nbSelected=0) and (Not Lst.AllSelected) then
    begin
    PGIInfo('Aucun document sélectionné.', TitreHalley);
    exit;
    end;

  if bMarque then
    Msg := 'Vous demandez le marquage pour Publication sur eWS de tous les documents sélectionnés.'+#13#10
     + ' Confirmez-vous la demande de publication ?'
  else
    Msg := 'Vous demandez l''annulation de Publication sur eWS de tous les documents sélectionnés.'+#13#10
     + ' Confirmez-vous l''annulation de publication ?';

  if PGIAsk(Msg, TitreHalley)=mrNo then exit;

  // liste des documents à publier
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
        MarquageEwsAPublier(Qry.FindField('DPD_DOCGUID').AsString, bMarque);
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
      MarquageEwsAPublier(Qry.FindField('DPD_DOCGUID').AsString, bMarque);
      end;
    FiniMove;
    END;
  // déselectionne
  FinTraitmtMul(TFMul(Ecran));
  // actualisation
  TFMul(Ecran).BChercheClick(Nil);
end;


procedure TOF_DPDOCUMENTSEWS.BOuvrir_OnClick(Sender: TObject);
// Publication sur Ews des documents sélectionnés
var EwsId, EwsLib : String;
    L : THDBGrid;
begin
  {$IFDEF EWS}
  // récup des objets
  L := TFMul(Ecran).FListe;
  if (L.NbSelected=0) and (not L.AllSelected) then
    begin
    PGIInfo('Aucun document sélectionné.', TitreHalley);
    exit;
    end;

  if PGIAsk('Vous allez sélectionner une destination dans l''arborescence eWS'+#13#10
   +' pour publier tous les documents sélectionnés. Confirmez-vous ?')=mrNo then
    exit;

  // Appel objet com pour sélection d'un noeud eWS
  EwsId := '';
  EwsLib := '';
  EwsSelectionneNoeud(EwsId, EwsLib);
  if EwsId='' then
    begin
    PGIInfo('Aucune destination sélectionnée.');
    exit;
    end;

  // Fiche permettant de contrôler le déroulement du traitement
  FicheTraitementCourant(Ecran, 'Publier sur eWS', EwsId);
  {$ENDIF}
end;

Initialization
  registerclasses ( [ TOF_DPDOCUMENTSEWS ] ) ;
end.

