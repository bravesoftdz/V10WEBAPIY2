{***********UNITE*************************************************
Auteur  ...... : Bernadette Tynévez
Créé le ...... : 01/10/2007
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : AMLISTEREGFR ()
Mots clefs ... : TOF;AMLISTEREGFR
*****************************************************************}
Unit AMLISTEREGFR_TOF ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
     HQry,
     Menus,
     {Graphics,}
     Windows,
     {dialogs,}
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     mul,
     fe_main,
     HDB,
{$else}
     eMul,
     MaineAgl,
{$ENDIF}
     uTob,
     forms,
     sysutils,
     ComCtrls,
     { AGL }
     HCtrls,
     HEnt1,
     HMsgBox,
     UTOF,
     ParamSoc,
     AGLInit,
     HStatus,
     HTB97,
     Math,
     { Amortissement}
     ImEnt,
     Outils,
     ImPlan,
     ImOutGen,
     ImOuPlan,
     AmListe_TOF,
     ImmoREGFR_TOM,
     AmDupREGFR_TOF,
     AmCrREGFR_TOF,
     Immo_TOM,
     ImmoCpte_TOM ;

Type TypeOpeREGFR = (toValidREGFR);

Type
  TOF_AMLISTEREGFR = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;

  private
    { Contrôles }
    POPZoomAction: TPopupMenu;
    POPZoomCreation: TPopupMenu;
    POPUpListe : TPopupMenu;

   {$IFDEF EAGLCLIENT}
    FListe: THGrid;
   {$ELSE}
    FListe: THDBGrid;
   {$ENDIF}

    { Variables }
    fAction: TActionFiche;
    fWhereOrig: string;
    fOnSaveKeyDownEcran: procedure(Sender: TObject; var Key: Word; Shift:TShiftState) of object;
    fCodeImmo :string;
    fTobAss : TOB;
    fValid : integer;

    { Méthodes }
    procedure OnKeyDownEcran(Sender: TObject; var Key: Word; Shift:TShiftState);
    procedure OnRowEnterListe(Sender: TObject); overload ;
    procedure OnListeDblClick(Sender: TObject);
    procedure OnKeyDownFListe(Sender: TObject; var Key: Word; Shift:TShiftState);
    procedure AfterShow;
    procedure OnFlipSelectionGrille (Sender : TObject);
    procedure OnSelectAllClick ( Sender : TObject );
    procedure RecupereAllSelect ( T : TOB );
    procedure EnableBoutonsAction;
    function  ListeVide: boolean;
    procedure RafraichirListeImmo(stCode: string; Selection : boolean);
    procedure RecupereCritereXX_WHERE(Sender: TObject);
    procedure OnSupprimeClick(Sender: TObject);
    procedure OnPopZoomCreationPopup(Sender: TObject);
    procedure EnableZoomAction;
    procedure EnableZoomActionSerie;
    procedure OnPopZoomActionPopup(Sender: TObject);
    procedure OnPopUpListePopup(Sender: TObject);
    procedure OnCreationPROClick(Sender: TObject);
    procedure OnCreationSerieClick(Sender: TObject);
    procedure OnDupliqueImmoClick(Sender: TObject);
    function  ConfirmationSuppressionImmo ( CodeImmo : string ) : boolean;
    procedure SuppressionFicheImmo;
    procedure OperationSerie (TypeOpe : TypeOpeREGFR);
    function  ExecuteOperation (TypeOpe : TypeOpeREGFR; CodeI : string):integer;
    procedure OnValidationClick(Sender: TObject);
    { partie rendue publique
    procedure ValidationImmo (CodeI: string, TAss :TOB);
    procedure AttribuerComptesAssocies (var RcompteAss:TCompteAss; T:TOB; CompteImmo:string);
    procedure VerifComptesAssocies (Ass : TCompteAss; TAss : TOB);
    procedure AjouteCompte (TAss : TOB ; Compte : string);
    procedure EcrireComptesAssocies (TAss : TOB); }
  end ;

////////////////////////////////////////////////////////////////////////////////
procedure AMLanceFiche_ListeImmosREGFR(Action: TActionFiche = taModif; Param: string = '');
procedure AMValidationImmo (CodeI : string; var fValid : integer; TAss : TOB);
////////////////////////////////////////////////////////////////////////////////


Implementation

const
  HM: array[0..10] of string =
{0}  ('Evaluation des immobilisations'
{1}   ,'Veuillez définir la date du passage au régime réel.'
{2}   ,'Aucune immobilisation sélectionnée'
{3}   ,'Validation'
{4}   ,'La date du passage au régime réel '
{5}   ,' ne coïncide pas avec le début de l''exercice en cours.'
{6}   ,'La validation de l''immobilisation a échoué.'
{7}   ,'La suppression de l''immobilisation a échoué.'
{8}   ,'Confirmez-vous la validation de cette immobilisation ?'
{9}   ,'La validation va porter sur l''ensemble de la sélection. Confirmez-vous l''opération ?'
{10}  ,' n''est pas une année correcte pour dater le passage au régime réel.'
      );

procedure AMLanceFiche_ListeImmosREGFR(Action: TActionFiche = taModif; Param: string = '');
var stMsg : string;
    stAction: string;
begin
  // Paramètres d'appel
  if (Action = taConsult) then
    stAction := 'ACTION=CONSULTATION'
  else
    stAction := 'ACTION=MODIFICATION';
  stAction := stAction + ';' + Param;

  // Contrôle date de changement de régime
  if not AMDatePassageREGFR ('IFR', 'IFR', stMsg) then // dans ImEnt
     PGIBox(stMsg)
  else
     AGLLanceFiche('AM', 'AMLISTEREGFR', '', '', stAction);
end;


procedure TOF_AMLISTEREGFR.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_AMLISTEREGFR.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_AMLISTEREGFR.OnUpdate ;
begin
  Inherited ;

  // Requête a été faite sur le ONLOAD => Récup valeur initiale du critère complémentaire
  SetControlText('XX_WHERE', fWhereOrig);
  // Revoir boutons au cas où liste non vide
  EnableBoutonsAction;
end ;

procedure TOF_AMLISTEREGFR.OnLoad ;
begin
  Inherited ;

  TFMul(Ecran).Q.Manuel := True;
  TFMul(Ecran).SETDBLISTE ('AMLISTIMMOREGFR');
  TFMul(Ecran).FNomFiltre := 'AMLISTIMMOREGFR';
  TFMul(Ecran).Q.Manuel := False; // pour avoir l'exécution automatique de la requête

  // Appliquer les critères retenus dans chaque onglet du MUL
  RecupereCritereXX_WHERE (nil);
end ;


procedure TOF_AMLISTEREGFR.OnArgument (S : String ) ;
var St : string;
begin
  Inherited ;

////////////////////////////////////////////////////////////////////////////
//       Découpage paramètre en entrée
////////////////////////////////////////////////////////////////////////////

  // Action
  St := ReadTokenSt(S);
  if St = 'ACTION=MODIFICATION' then  fAction := taModif
  else   fAction := taConsult;

  // // Aide en ligne
  // Ecran.HelpContext :=


////////////////////////////////////////////////////////////////////////////
//         Initialisations MUL et Fliste
////////////////////////////////////////////////////////////////////////////

{$IFDEF EAGLCLIENT}
  TheMulQ := TFMul(Ecran).Q.TQ;
  FListe := THGrid(GetControl('FLISTE'));
  FListe.SortEnabled := False;
{$ELSE}
  TheMulQ := TFMul(Ecran).Q;
  FListe := THDBGrid(GetControl('FLISTE'));
  FListe.OnRowEnter := OnRowEnterListe;
  FListe.SortEnabled := True;
{$ENDIF}
  // Cacher l'onglet Compléments
  SetControlVisible('PComplement', False);


////////////////////////////////////////////////////////////////////////////
//                 Initialisations des zones
////////////////////////////////////////////////////////////////////////////

  SetControlVisible('bSelectAll', True);
  SetControlVisible('bZoomAction', True);
  SetControlEnabled('bZoomAction', False);
  // TOB des comptes associés créés lors de la validation des immos
  fTobAss := TOB.Create('', nil, -1) ;

  // // PopUps
  POPZoomCreation := TPopUpMenu(GetControl('POPZOOMCREATION', True));
  POPZoomAction := TPopUpMenu(GetControl('POPZOOMACTION', True));
  POPUpListe := TPopUpMenu (GetControl('POPUPLISTE', True));
  EnableBoutonsAction;

  // // Onglet STANDARDS
  SetControlProperty('IR_VO', 'DisplayFormat', StrfMask(V_PGI.OkDecV,'', True));
  SetControlProperty('IR_VO_', 'DisplayFormat', StrfMask(V_PGI.OkDecV,'', True));
  // Exclure certains items de TIMETHODEIMMO
  SetControlProperty('IR_METHODEECO','Plus',' AND CO_CODE<>"VAR"');



////////////////////////////////////////////////////////////////////////////
//                   Evènements
////////////////////////////////////////////////////////////////////////////

  // // Evènements généraux
  fListe.OnDblClick := OnListeDblClick;
  fListe.OnKeyDown := OnKeyDownFListe;
  fOnSaveKeyDownEcran := Ecran.OnKeyDown;
  Ecran.OnKeyDown := OnKeyDownEcran;
  TFMul(Ecran).OnAfterFormShow := AfterShow;
  fListe.OnFlipSelection := OnFlipSelectionGrille;
  TFMul(Ecran).bSelectAll.OnClick := OnSelectAllClick;
  TToolbarButton97(GetControl('bDelete')).OnClick := OnSupprimeClick;

  // // EVENTS Popup Création
  PopZoomCreation.OnPopup := OnPopZoomCreationPopup;
  TMenuItem(GetControl('PleinePropriete')).OnClick := OnCreationPROClick;
  TMenuItem(GetControl('CreationSerie')).OnClick := OnCreationSerieClick;
  TMenuItem(GetControl('DupliqueImmo')).OnClick := OnDupliqueImmoClick;

  // // EVENTS Popup Action
  PopZoomAction.OnPopup := OnPopZoomActionPopup;
  TMenuItem(GetControl('Validation')).OnClick := OnValidationClick;

  // // EVENTS Popup F11
  PopUpListe.OnPopup := OnPopUpListePopUp;
  TMenuItem(GetControl('PleinePropriete1')).OnClick := OnCreationPROClick;
  TMenuItem(GetControl('CreationSerie1')).OnClick := OnCreationSerieClick;
  TMenuItem(GetControl('DupliqueImmo1')).OnClick := OnDupliqueImmoClick;
  TMenuItem(GetControl('Validation1')).OnClick := OnValidationClick;

   // Critères initiaux à ajouter au XX_WHERE
  fWhereOrig := GetControlText('XX_WHERE');
end ;

procedure TOF_AMLISTEREGFR.OnClose ;
begin
  fTobAss.Free;
  TheMulQ := nil;
  Inherited ;
end ;

procedure TOF_AMLISTEREGFR.OnDisplay () ;
begin
  Inherited ;
  EnableBoutonsAction;
  if (not ListeVide) then
      OnRowEnterListe(nil);
end ;

procedure TOF_AMLISTEREGFR.OnCancel () ;
begin
  Inherited ;
end ;



{***********A.G.L.***********************************************
Auteur  ...... : Bernadette Tynévez
Créé le ...... : 02/10/2007
Modifié le ... :   /  /
Description .. : Gestion touche dans l'écran
Mots clefs ... :
*****************************************************************}
procedure TOF_AMLISTEREGFR.OnKeyDownEcran(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
  case Key of
    VK_DELETE :
      begin
        if (Shift = [ssCtrl]) and
           (GetControlVisible ('bDelete')) and
           (TToolbarButton97(GetControl('bDelete')).Enabled) then
        begin
          Key := 0;
          TToolbarButton97(GetControl('bDelete')).Click;
        end;
      end;
  else
    fOnSaveKeyDownEcran(Sender, Key, Shift);
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Bernadette Tynévez
Créé le ...... : 02/10/2007
Modifié le ... :   /  /
Description .. : Clic sur une immo ou défilement vertical dans la liste
Mots clefs ... :
*****************************************************************}
procedure TOF_AMLISTEREGFR.OnRowEnterListe(Sender: TObject);
begin
  exit;
end;


{***********A.G.L.***********************************************
Auteur  ...... : Bernadette Tynévez
Créé le ...... : 02/10/2007
Modifié le ... :   /  /
Description .. : Evènement MUL
Mots clefs ... :
*****************************************************************}
procedure TOF_AMLISTEREGFR.AfterShow;
begin
  OnRowEnterListe(nil);
end;


{***********A.G.L.***********************************************
Auteur  ...... : Bernadette Tynévez
Créé le ...... : 02/10/2007
Modifié le ... :   /  /
Description .. : Ouverture menu Action
Mots clefs ... :
*****************************************************************}
procedure TOF_AMLISTEREGFR.EnableBoutonsAction;
begin
  SetControlVisible('BZoomAction', (not ListeVide));
  SetControlVisible('BDelete', (not ListeVide));
end;


{***********A.G.L.***********************************************
Auteur  ...... : Bernadette Tynévez
Créé le ...... : 02/10/2007
Modifié le ... :   /  /
Description .. : Ouverture fiche immo
Mots clefs ... :
*****************************************************************}
procedure TOF_AMLISTEREGFR.OnListeDblClick(Sender: TObject);
var
  stCode: string;
begin
{$IFDEF EAGLCLIENT}
   TheMulQ := TFMul(Ecran).Q.TQ;
{$ELSE}
   TheMulQ := TFMul(Ecran).Q;
{$ENDIF}

  if TheMulQ.RecordCount <> 0 then
  begin
    stCode := GetField('IR_COMPTEIMMO');

    AMLanceFiche_FicheREGFR(stCode, fAction);
    //RechargeListeChoix;
    // Réactualiser, l'immo courante est perdue
    RafraichirListeImmo (stCode, True);
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Bernadette Tynévez
Créé le ...... : 02/10/2007
Modifié le ... :   /  /
Description .. : Gestion de touche dans la liste
Mots clefs ... :
*****************************************************************}
procedure TOF_AMLISTEREGFR.OnKeyDownFListe(Sender: TObject; var Key: Word; Shift: TShiftState);
begin

  case Key of
  VK_F11: begin
          Key := 0;
          if (fAction <> taConsult) then
             TPopupMenu(GetControl('PopupListe')).Popup(Mouse.CursorPos.x, Mouse.CursorPos.y);
          end;
  VK_F5: begin
         Key := 0;
         OnListeDblClick(nil);
         end;
  end;

end;

{***********A.G.L.***********************************************
Auteur  ...... : Bernadette Tynévez
Créé le ...... : 02/10/2007
Modifié le ... :   /  /
Description .. : Touche ESPACE de sélection d'une immo
Mots clefs ... :
*****************************************************************}
procedure TOF_AMLISTEREGFR.OnFlipSelectionGrille(Sender: TObject);
begin
  inherited;
  SetControlEnabled('BZOOMACTION', (FListe.NbSelected > 0) or (FListe.AllSelected));
end;

{***********A.G.L.***********************************************
Auteur  ...... : Bernadette Tynévez
Créé le ...... : 02/10/2007
Modifié le ... :   /  /
Description .. : CTL+A Sélection de toute la liste
Mots clefs ... :
*****************************************************************}
procedure TOF_AMLISTEREGFR.OnSelectAllClick(Sender: TObject);
begin
  TFMul(Ecran).bSelectAllClick ( Sender );
  OnFlipSelectionGrille(fListe);
end;

{***********A.G.L.***********************************************
Auteur  ...... : Bernadette Tynévez
Créé le ...... : 02/10/2007
Modifié le ... :   /  /
Description .. : Charger en TOB toutes les immos de la multisélection
Mots clefs ... :
*****************************************************************}
procedure TOF_AMLISTEREGFR.RecupereAllSelect(T: TOB);
var
 stWhere : string;
 Q : TQuery;
begin
  if T = nil then exit;

  // RecupWhereCritere de HQRY reprend le XX_WHERE par défaut => compléter par le filtre en cours
  stWhere := RecupWhereCritere(TFMul(Ecran).Pages);
  RecupereCritereXX_WHERE (nil);

  if (GetControlText('XX_WHERE') = '') then
    Q := OpenSql (
    'SELECT IR_COMPTEIMMO,IR_LIBELLE, IR_DATEACHAT, IR_COMPTEREF, IR_METHODEECO, IR_VO, IR_COEF FROM IMMOREGFR ' +
                stWhere , True)
  else
    Q := OpenSql (
    'SELECT IR_COMPTEIMMO,IR_LIBELLE, IR_DATEACHAT, IR_COMPTEREF, IR_METHODEECO, IR_VO, IR_COEF FROM IMMOREGFR ' +
                stWhere + ' AND ' + GetControlText('XX_WHERE'), True);

  // Remettre la valeur par défaut
  SetControlText('XX_WHERE', fWhereOrig);
  try
  if not Q.Eof then
      T.LoadDetailDB ('', '', '', Q, False);
  finally
    Ferme (Q);
  end;
end;


{***********A.G.L.***********************************************
Auteur  ...... : Bernadette Tynévez
Créé le ...... : 02/10/2007
Modifié le ... :   /  /
Description .. : Test liste vide
Mots clefs ... :
*****************************************************************}
function TOF_AMLISTEREGFR.ListeVide: boolean;
begin
{$IFDEF EAGLCLIENT}
   TheMulQ := TFMul(Ecran).Q.TQ;
   Result := (TheMulQ = nil) or ((TheMulQ <> nil) and (TheMulQ.Detail.Count = 0));
{$ELSE}
   TheMulQ := TFMul(Ecran).Q;
   Result := TheMulQ.isEmpty;
{$ENDIF}
end;


{***********A.G.L.***********************************************
Auteur  ...... : Bernadette Tynévez
Créé le ...... : 02/10/2007
Modifié le ... :   /  /
Description .. : Réactualiser la liste des immos,
Suite ........ : les immos validées ne sont plus présentes
Mots clefs ... :
*****************************************************************}
procedure TOF_AMLISTEREGFR.RafraichirListeImmo(stCode: string; Selection : boolean);
begin
  // Relancer la requête de multisélection
  TFMul(ECRAN).BChercheClick(nil);

  // Sauf si on revient d'une validation d'immo, se repositionner sur l'enreg courant
  if (StCode <> '') and Selection then
  begin
    {$IFDEF EAGLCLIENT}
    TheMulQ := TFMul(Ecran).Q.TQ;
    {$ELSE}
    TheMulQ := TFMul(Ecran).Q;
    {$ENDIF}
    TheMulQ.Locate('IR_COMPTEIMMO', stCode, []) ;
  end;
end;


{***********A.G.L.***********************************************
Auteur  ...... : Bernadette Tynévez
Créé le ...... : 02/10/2007
Modifié le ... :   /  /
Description .. : Recomposer le critère complémentaire XX_WHERE
Suite ........ : suivant les sélections constatées dans les onglets du MUL
Mots clefs ... :
*****************************************************************}
procedure TOF_AMLISTEREGFR.RecupereCritereXX_WHERE(Sender: TObject);
begin

end;

{***********A.G.L.***********************************************
Auteur  ...... : Bernadette Tynévez
Créé le ...... : 02/10/2007
Modifié le ... :   /  /
Description .. : Bouton Suppression
Mots clefs ... :
*****************************************************************}
procedure TOF_AMLISTEREGFR.OnSupprimeClick(Sender: TObject);
var
  i: integer;
  stCode: string;
  TListeAll : TOB;
begin
  if (FListe.NbSelected = 0) and (not FListe.AllSelected) then
  begin
    PGIBox(HM[2]);
    exit;
  end;

  stCode := GetField('IR_COMPTEIMMO');

  if FListe.AllSelected then
  begin
    InitMove(TheMulQ.RecordCount, '');
    TListeAll := TOB.Create ('',nil,-1);
    try
      RecupereAllSelect(TListeAll);
      for i := 0 to TListeAll.Detail.Count - 1 do
      begin
        MoveCur(False);
        fCodeImmo := TListeAll.Detail[i].GetValue('IR_COMPTEIMMO');

        if ConfirmationSuppressionImmo ( fCodeImmo ) then
        begin
           if Transactions(SuppressionFicheImmo, 3) <> oeOk then
           begin
              MessageAlerte(HM[7]);
              Break;
           end;
        end;
      end;
      FListe.AllSelected := False;
      TToolbarButton97(GetControl('bSelectAll')).Down := False;
    finally
        FreeAndNil(TListeAll);
    end;
  end
  else
  begin
    InitMove(FListe.NbSelected, '');
    for i := 0 to FListe.NbSelected - 1 do
    begin
      MoveCur(False);
      FListe.GotoLeBookmark(i);
      fCodeImmo := GetField('IR_COMPTEIMMO');
      if ConfirmationSuppressionImmo ( fCodeImmo ) then
      begin
         if Transactions(SuppressionFicheImmo, 3) <> oeOk then
         begin
            MessageAlerte(HM[7]);
            Break;
         end;
      end;
    end;
    FListe.ClearSelected;
  end;
  FiniMove;

  RafraichirListeImmo (stCode, False);
  SetControlEnabled('BZoomAction', False);
end;


{***********A.G.L.***********************************************
Auteur  ...... : Bernadette Tynévez
Créé le ...... : 02/10/2007
Modifié le ... :   /  /
Description .. : Recomposition du menu CREATION en fonction de la sélection
Mots clefs ... :
*****************************************************************}
procedure TOF_AMLISTEREGFR.OnPopZoomCreationPopup(Sender: TObject);
begin
{$IFDEF EAGLCLIENT}
   TheMulQ := TFMul(Ecran).Q.TQ;
{$ELSE}
   TheMulQ := TFMul(Ecran).Q;
{$ENDIF}
 SetControlVisible('DupliqueImmo',(not TheMulQ.EOF) and (FListe.NbSelected <= 1));
end;

{***********A.G.L.***********************************************
Auteur  ...... : Bernadette Tynévez
Créé le ...... : 02/10/2007
Modifié le ... :   /  /
Description .. : CREATION Immo en pleine propriété
Mots clefs ... :
*****************************************************************}
procedure TOF_AMLISTEREGFR.OnCreationPROClick(Sender: TObject);
var
  stCode: string;
begin
  stCode := '';
  if TheMulQ.RecordCount <> 0 then
     stCode := GetField('IR_COMPTEIMMO');

  AMLanceFiche_FicheREGFR ('', taCreatOne);
  // Réactualiser, l'immo courante est perdue
  RafraichirListeImmo (stCode, True);
end;


{***********A.G.L.***********************************************
Auteur  ...... : Bernadette Tynévez
Créé le ...... : 02/10/2007
Modifié le ... :   /  /
Description .. : CREATION de plusieurs immos en série
Mots clefs ... :
*****************************************************************}
procedure TOF_AMLISTEREGFR.OnCreationSerieClick(Sender: TObject);
var
  stCode: string;
begin
  stCode := '';
  if not TheMulQ.Eof then
     stCode := GetField('IR_COMPTEIMMO');
  AMLanceFiche_FicheREGFR('', taCreat);

  // Réactualiser, l'immo courante est perdue
  RafraichirListeImmo (stCode, True);
end;


{***********A.G.L.***********************************************
Auteur  ...... : Bernadette Tynévez
Créé le ...... : 02/10/2007
Modifié le ... :   /  /
Description .. : CREATION d'immo par duplication d'une immo
Mots clefs ... :
*****************************************************************}
procedure TOF_AMLISTEREGFR.OnDupliqueImmoClick(Sender: TObject);
var
  stCode: string;
  TobW : TOB;
begin
  stCode := GetField('IR_COMPTEIMMO');

  TobW := TOB.Create('', nil, -1) ;
  // Stocker la Tob du nouveau code dans la TOB générale TheTob accessible partout
  TheTob := TobW;

  if (AMLanceFiche_DupREGFR (stCode) = mrYes) then
     // TheTob remise à nil => TobW contient la valeur
     AMLanceFiche_FicheREGFR (TobW.GetValue ('CODE'), taModif);

  TobW.Free;
  // Réactualiser, l'immo courante est perdue
  RafraichirListeImmo (stCode, True);
end;


{***********A.G.L.***********************************************
Auteur  ...... : Bernadette Tynévez
Créé le ...... : 02/10/2007
Modifié le ... :   /  /
Description .. : Confirmation de suppression d'une immo
Mots clefs ... :
*****************************************************************}
function TOF_AMLISTEREGFR.ConfirmationSuppressionImmo ( CodeImmo : string ) : boolean;
var stCaption : string;
begin
  stCaption := TraduireMemoire('Immobilisation n°')+ CodeImmo;
  Result := (PGIAsk('Confirmez-vous la suppression de l''enregistrement ?',stCaption)=mrYes);
end;

{***********A.G.L.***********************************************
Auteur  ...... : Bernadette Tynévez
Créé le ...... : 02/10/2007
Modifié le ... :   /  /
Description .. : Suppression d'une immo
Mots clefs ... :
*****************************************************************}
procedure TOF_AMLISTEREGFR.SuppressionFicheImmo;
begin
  if ExecuteSQL('DELETE FROM IMMOREGFR WHERE IR_COMPTEIMMO = "' + fCodeImmo + '"') <> 1 then
     V_PGI.IoError := oeSaisie
  else
     V_PGI.IoError := oeOk;
end;



{***********A.G.L.***********************************************
Auteur  ...... : Bernadette Tynévez
Créé le ...... : 02/10/2007
Modifié le ... :   /  /
Description .. : Ouverture des options des menus sans immo sélectionnée
Mots clefs ... :
*****************************************************************}
procedure TOF_AMLISTEREGFR.EnableZoomAction;
begin
  SetControlVisible('DupliqueImmo', not ListeVide);
  SetControlVisible('Validation', (fListe.NbSelected = 1));
end;

{***********A.G.L.***********************************************
Auteur  ...... : Bernadette Tynévez
Créé le ...... : 02/10/2007
Modifié le ... :   /  /
Description .. : Ouverture des options des menus avec immos sélectionnées
Mots clefs ... :
*****************************************************************}
procedure TOF_AMLISTEREGFR.EnableZoomActionSerie;
begin
  SetControlVisible('DupliqueImmo', False);
  SetControlVisible('Validation', True);
end;


{***********A.G.L.***********************************************
Auteur  ...... : Bernadette Tynévez
Créé le ...... : 02/10/2007
Modifié le ... :   /  /
Description .. : Menu ACTION en fonction de la sélection
Mots clefs ... :
*****************************************************************}
procedure TOF_AMLISTEREGFR.OnPopZoomActionPopup(Sender: TObject);
begin
  if (FListe.NbSelected <= 1) and (not FListe.AllSelected) then
     EnableZoomAction
  else
     EnableZoomActionSerie;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Bernadette Tynévez
Créé le ...... : 16/10/2007
Modifié le ... :   /  /
Description .. : Menu F11
Mots clefs ... :
*****************************************************************}
procedure TOF_AMLISTEREGFR.OnPopUpListePopUp(Sender: TObject);
begin
  OnPopZoomActionPopup(Sender);
  SetControlVisible('DupliqueImmo1', (not ListeVide) and (FListe.NbSelected <= 1));
  SetControlVisible('Trait2', GetControlVisible('Validation'));
  SetControlVisible('Validation1', GetControlVisible('Validation'));
end;


{***********A.G.L.***********************************************
Auteur  ...... : Bernadette Tynévez
Créé le ...... : 02/10/2007
Modifié le ... :   /  /
Description .. : Branchement sur opération élémentaire ou en série
Suite ........ :
Mots clefs ... :
*****************************************************************}
procedure TOF_AMLISTEREGFR.OperationSerie (TypeOpe : TypeOpeREGFR);
var
  i : integer;
  stCode, Titre : string;
  TListeAll : TOB;
  Faire, mr: integer;
begin
  stCode := GetField('IR_COMPTEIMMO');

  case TypeOpe of
  toValidREGFR : Titre := HM[3];
  end;

  // Demande de confirmation par l'utilisateur
  if (fListe.AllSelected) or (fListe.NbSelected > 1) then
    Faire := PGIAskCancel(HM[9], Ecran.Caption)
  else
    Faire := PGIAskCancel(HM[8], Ecran.Caption);

  case Faire of

  mrYes :
    begin
    // Création Lieu Géo = 001 si n'existe pas
    CreationLieuParDefaut;

    fValid := 0;
    fTobAss.ClearDetail;

    // Validation de toutes les immos affichées
    if fListe.AllSelected then
    begin
       TListeAll := TOB.Create ('', nil, -1);
       try
         RecupereAllSelect(TListeAll);

         if (TListeAll.Detail.Count > 0) then
         begin
            // On se positionne sur un enregistrement de la sélection
            if (fListe.NbSelected >= 1) then fListe.GotoLeBookmark(0);

            // On boucle sur les opérations individuelles
            for i := 0 to TListeAll.Detail.Count - 1 do
            begin
               mr := ExecuteOperation (TypeOpe, TListeAll.Detail[i].GetValue('IR_COMPTEIMMO'));
               if mr <> mrYes then break;
            end;
            // Afficher le Compte-rendu
            // Stocker la Tob des comptes asociés dans la TOB générale TheTob accessible partout
            TheTob := fTobAss;
            AMLanceCompteRendu_REGFR (IntToStr (fValid));
         end;
         fListe.AllSelected := False;
       finally
         FreeAndNil (TListeAll);
       end;

    end
    // Validation de une ou plusieurs immos mais pas toutes celles affichées
    else if (fListe.NbSelected >= 1) then
    begin

       // On se positionne sur un enregistrement de la sélection
       if (fListe.NbSelected >= 1) then fListe.GotoLeBookmark(0);

       // On boucle sur les opérations individuelles
       for i := 0 to fListe.NbSelected - 1 do
       begin
         fListe.GotoLeBookmark(i);
         mr := ExecuteOperation(TypeOpe, GetField('IR_COMPTEIMMO'));
         if mr <> mrYes then break;
       end;
       // Afficher le Compte-rendu
       // Stocker la Tob des comptes asociés dans la TOB générale TheTob accessible partout
       TheTob := fTobAss;
       AMLanceCompteRendu_REGFR (IntToStr (fValid));

       fListe.ClearSelected;
    end;
    // Rafraichir sans repositionner => immo n'est plus dans la liste
    RafraichirListeImmo (stCode, False);
    end;

  mrNo : begin
         // Rafraichir avec repositionnement sur l'immo courante, car GotoLeBookmark repositionne
         RafraichirListeImmo (stCode, True);
         end;
  end;
end;


{***********A.G.L.***********************************************
Auteur  ...... : Bernadette Tynévez
Créé le ...... : 02/10/2007
Modifié le ... :   /  /
Description .. : Branchement sur l'opération élémentaire
Mots clefs ... :
*****************************************************************}
function TOF_AMLISTEREGFR.ExecuteOperation (TypeOpe : TypeOpeREGFR; CodeI : string):integer;
begin
  result := mrYes;

  case TypeOpe of
       toValidREGFR :
       begin
          // Transaction pour éviter d'ajouter dans une table et pas dans les autres
          BEGINTRANS ;
          try
            AMValidationImmo (CodeI, fValid, fTobAss);
            COMMITTRANS ;
          except
            ROLLBACK ;
            PGIBox(HM[6], HM[0]);
            result := mrNo;
          end;
       end;
  end;
end;


{***********A.G.L.***********************************************
Auteur  ...... : Bernadette Tynévez
Créé le ...... : 02/10/2007
Modifié le ... :
Description .. : Ajout dans la table IMMO/IMMOLOG/IMMOAMOR
Mots clefs ... :
*****************************************************************}
{procedure TOF_AMLISTEREGFR.ValidationImmo (CodeI : string; TAss : TOB);
var CodeImmo :string;
    T1, TCalcul, TFiche, TAmor : Tob;
    RCompteAss : TCompteAss;
    PremMois, PremAnnee, NbMois: Word;
    Plan : TPlanAmort;
    Infos: TInfoLog;
begin

  // T1 = enreg lu dans IMMOREGFR
  T1 := TOB.Create('IMMOREGFR', nil, -1) ;
  T1.SelectDB ('"'+ CodeI +'"', nil);
  if ExisteSQL
     ('SELECT I_IMMO FROM IMMO WHERE I_IMMO="' + T1.GetValue('IR_COMPTEIMMO') + '"') then
     CodeImmo := NouveauCodeImmo
  else
     CodeImmo := T1.GetValue('IR_COMPTEIMMO');

  // Récup calculs durée amort., VNC, V retenue etc.
  TCalcul := TOB.Create('', nil, -1) ;
  AMCalculREGFR (T1, TCalcul);


// TABLE IMMO --------------------------------------------------------------------------

  TFiche := TOB.Create('IMMO', nil, -1) ;
  //TAss := TOB.Create('', nil, -1) ;
  TAmor  := TOB.Create('', nil, -1);

  try
    with TFiche do
    begin
    PutValue('I_IMMO', CodeImmo);
    PutValue('I_LIBELLE', T1.GetValue('IR_LIBELLE'));

    PutValue('I_COMPTEIMMO', T1.GetValue('IR_COMPTEREF'));
    PutValue('I_MONTANTHT', TCalcul.GetValue('VRETENUE'));
    PutValue('I_MONTANTBASEAMORT', TCalcul.GetValue('VRETENUE'));
    PutValue('I_NATUREIMMO', 'PRO');
    PutValue('I_ABREGE', Copy (T1.GetValue('IR_LIBELLE'), 1, 17));
    PutValue('I_DATEPIECEA', T1.GetValue('IR_DATEACHAT'));
    PutValue('I_DATEAMORT', T1.GetValue('IR_DATEACHAT'));
    PutValue('I_BASETAXEPRO', TCalcul.GetValue('VRETENUE'));

    PutValue('I_BASEECO', TCalcul.GetValue('VRETENUE'));
    PutValue('I_METHODEECO', T1.GetValue('IR_METHODEECO'));
    PutValue('I_DUREEECO', TCalcul.GetValue('DUREE_AMORT'));
    if T1.GetValue('IR_METHODEECO') = 'DEG' then
       PutValue('I_TAUXECO', GetTaux(T1.GetValue('IR_METHODEECO'),
                                     T1.GetValue('IR_DATEACHAT'),
                                     T1.GetValue('IR_DATEACHAT'),
                                     T1.GetValue('IR_DUREERESTM')))
    else if T1.GetValue('IR_METHODEECO') = 'LIN' then
            PutValue('I_TAUXECO', GetTaux(T1.GetValue('IR_METHODEECO'),
                                          T1.GetValue('IR_DATEACHAT'),
                                          T1.GetValue('IR_DATEACHAT'),
                                          TCalcul.GetValue('DUREE_AMORT')));
    AttribuerComptesAssocies (RCompteAss, TFiche, T1.GetValue('IR_COMPTEREF'));
    // Mise en TOB des cptes absents du plan de cptes
    VerifComptesAssocies (RCompteAss, TAss);
    PutValue('I_SOCIETE', V_PGI.CodeSociete);
    if T1.GetValue('IR_METHODEECO') = 'LIN' then
       PutValue('I_JOURNALA', '***');
    PutValue('I_QUALIFIMMO', 'R');
    PutValue('I_ETABLISSEMENT', VHImmo^.EtablisDefaut);
    PutValue('I_QUANTITE', 1);
    PutValue('I_PAYS', 'FRA');
    PutValue('I_LIEUGEO', '001');
    PutValue('I_TYPELOYERCB', 'LCO');
    PutValue('I_PERIODICITE', 'MEN');
    PutValue('I_VERSEMENTCB', 'AVA');
    PutValue('I_DATEDEBUTECH', T1.GetValue('IR_DATEACHAT'));
    PutValue('I_DATEFINECH', VHImmo^.Encours.Fin);
    PutValue('I_NATUREBIEN', 'GEB');
    PutValue('I_ETAT', 'OUV');
    PutValue('I_DEVISE', V_PGI.DevisePivot);
    PutValue('I_OPERATION', '-');
    PutValue('I_OPEMUTATION', '-');
    PutValue('I_OPEECLATEMENT', '-');
    PutValue('I_OPECESSION', '-');
    PutValue('I_OPECHANGEPLAN', '-');
    PutValue('I_OPELIEUGEO', '-');
    PutValue('I_OPEETABLISSEMENT', '-');
    PutValue('I_OPELEVEEOPTION', '-');
    PutValue('I_OPEMODIFBASES', '-');
    PutValue('I_OPEDEPREC', '-');
    PutValue('I_OPEREG', '-');
    PutValue('I_OPEPRIME', '-');
    PutValue('I_OPEREDUCPRI', '-');
    PutValue('I_OPESBV', '-');
    PutValue('I_OPEARD', '-');
    PutValue('I_OPEREMPL', '-');
    PutValue('I_VENTILABLE', '-');
    PutValue('I_VENTILABLE1', '-');
    PutValue('I_VENTILABLE2', '-');
    PutValue('I_VENTILABLE3', '-');
    PutValue('I_VENTILABLE4', '-');
    PutValue('I_VENTILABLE5', '-');
    PutValue('I_REPRISEECO', TCalcul.GetValue('AMORT'));
    NOMBREMOIS(T1.GetValue('IR_DATEACHAT'), VHImmo^.Encours.Deb, PremMois, PremAnnee, NbMois);
    PutValue('I_DUREEREPRISE', MaxIntValue([0, NbMois - 1]));
    PutValue('I_BASEAMORDEBEXO', TCalcul.GetValue('VRETENUE'));
    PutValue('I_BASEAMORFINEXO', TCalcul.GetValue('VRETENUE'));
    PutValue('I_COMPTEREF', T1.GetValue('IR_COMPTEREF'));
    PutValue('I_VAOACEDEE', T1.GetValue('IR_COMPTEREF'));
    PutValue('I_VALEURACHAT', TCalcul.GetValue('VRETENUE'));
    PutValue('I_SURAMORT', '-');
    PutValue('I_SUSDEF','A');
    PutValue('I_REGLECESSION','NOR');
    PutValue('I_NONDED','-');
    PutValue('I_DATEDEBECO', T1.GetValue('IR_DATEACHAT'));
    PutValue('I_DATEDEBFIS', T1.GetValue('IR_DATEACHAT'));
    PutValue('I_DPI','-');
    PutValue('I_DPIEC', '-');
    PutValue('I_SUBVENTION', 'NON');
    PutValue('I_SBVEC', '-');
    PutValue('I_PFR', 'X');  // immo issue du régime forfait
    PutValue('I_ACHFOR', T1.GetValue('IR_DATEACHAT'));
    PutValue('I_PRIXACFOR', T1.GetValue('IR_VO'));
    PutValue('I_VNCFOR', TCalcul.GetValue('VRETENUE'));
    PutValue('I_DURRESTFOR', TCalcul.GetValue('DUREE_AMORT'));
    PutValue('I_VERSION', V_PGI.NumVersion);
    PutValue('I_ARD', '-');
    end;

// TABLE IMMOAMOR --------------------------------------------------------------------------

    Plan := TPlanAmort.Create(true);
    Plan.CalculTOB( TFiche, iDate1900);
    Plan.SauveTOB ( TAmor );
    TFiche.PutValue('I_PLANACTIF', Plan.NumSeq);
    TFiche.PutValue('I_DATEDERMVTECO', Plan.GetDateFinAmortEx(Plan.AmortEco));

// ENREGISTREMENT ---------------------------------------------------------------------------

    TFiche.InsertDB(nil);
    TAmor.InsertDB(nil);

// TABLE IMMOLOG --------------------------------------------------------------------------

    Infos.TVARecuperable := 0;
    Infos.TVARecuperee := 0;
    EnregLogAcquisition (CodeImmo, T1.GetValue('IR_DATEACHAT'), Plan.NumSeq, Infos);
    FreeAndNil( Plan );
    fValid := fValid + 1;
    EcrireComptesAssocies (TAss);

// SUPPRESSION IMMOREGFR ---------------------------------------------------------------
    //ExecuteSQL('DELETE FROM IMMOREGFR WHERE IR_COMPTEIMMO = "' + CodeI + '"');

  finally
    T1.Free ;
    TCalcul.Free ;
    TFiche.Free ;
    TAmor.Free ;
  end;
end;   }



{***********A.G.L.***********************************************
Auteur  ...... : Bernadette Tynévez
Créé le ...... : 12/10/2007
Modifié le ... :   /  /
Description .. : Plan comptable associé au compte d'immo passé en entrée
Mots clefs ... :
*****************************************************************}
{procedure TOF_AMLISTEREGFR.AttribuerComptesAssocies (var RCompteAss: TCompteAss; T : TOB; CompteImmo :string);
begin
  // Chercher les comptes associés au compte d'abord dans IMMOCPTE, sinon les attribuer par défaut
  RecupereComptesAssocies (nil, CompteImmo, RCompteAss);
  //GetComptesAssociesParDefaut(RCompteAss, CompteImmo);

  with T do
  begin
    PutValue('I_COMPTEAMORT', RCompteAss.Amort);
    PutValue('I_COMPTEDOTATION', RCompteAss.Dotation);
    PutValue('I_COMPTEDEROG', RCompteAss.Derog );
    PutValue('I_REPRISEDEROG', RCompteAss.RepriseDerog);
    PutValue('I_PROVISDEROG', RCompteAss.ProvisDerog);
    PutValue('I_DOTATIONEXC', RCompteAss.DotationExcep);
    PutValue('I_VACEDEE', RCompteAss.VaCedee);
    PutValue('I_AMORTCEDE', RCompteAss.AmortCede);
    PutValue('I_VAOACEDEE', RCompteAss.VoaCede);
    PutValue('I_REPEXPLOIT', RCompteAss.RepriseExploit);
    PutValue('I_REPEXCEP', RCompteAss.RepriseExcep);
  end;
end; }

{***********A.G.L.***********************************************
Auteur  ...... : Bernadette Tynévez
Créé le ...... : 12/10/2007
Modifié le ... :   /  /
Description .. : Mise en TOB des comptes associés manquant dans le plan de comptes
Mots clefs ... :
*****************************************************************}
{rocedure TOF_AMLISTEREGFR.VerifComptesAssocies (Ass : TCompteAss; TAss : TOB);
begin
  if not (Ass.Immo = '') then
     if not  Presence ( 'GENERAUX','G_GENERAL', Ass.Immo) then
        AjouteCompte (TAss, Ass.Immo);
  if not (Ass.Amort = '') then
     if not  Presence ( 'GENERAUX','G_GENERAL', Ass.Amort) then
        AjouteCompte (TAss, Ass.Amort);
  if not (Ass.Dotation = '') then
     if not  Presence ( 'GENERAUX','G_GENERAL', Ass.Dotation) then
        AjouteCompte (TAss, Ass.Dotation);
  if not (Ass.Derog = '') then
     if not  Presence ( 'GENERAUX','G_GENERAL', Ass.Derog) then
        AjouteCompte (TAss, Ass.Derog);
  if not (Ass.RepriseDerog = '') then
     if not  Presence ( 'GENERAUX','G_GENERAL', Ass.RepriseDerog) then
        AjouteCompte (TAss, Ass.RepriseDerog);
  if not (Ass.DotationExcep = '') then
     if not  Presence ( 'GENERAUX','G_GENERAL', Ass.DotationExcep) then
        AjouteCompte (TAss, Ass.DotationExcep);
  if not (Ass.VaCedee = '') then
     if not  Presence ( 'GENERAUX','G_GENERAL', Ass.VaCedee) then
        AjouteCompte (TAss, Ass.VaCedee);
  if not (Ass.AmortCede = '') then
     if not  Presence ( 'GENERAUX','G_GENERAL', Ass.AmortCede) then
        AjouteCompte (TAss, Ass.AmortCede);
  if not (Ass.VoaCede = '') then
     if not  Presence ( 'GENERAUX','G_GENERAL', Ass.VoaCede) then
        AjouteCompte (TAss, Ass.VoaCede);
  if not (Ass.ProvisDerog = '') then
     if not  Presence ( 'GENERAUX','G_GENERAL', Ass.ProvisDerog) then
        AjouteCompte (TAss, Ass.ProvisDerog);
  if not (Ass.RepriseExcep = '') then
     if not  Presence ( 'GENERAUX','G_GENERAL', Ass.RepriseExcep) then
        AjouteCompte (TAss, Ass.RepriseExcep);
  if not (Ass.RepriseExploit = '') then
     if not  Presence ( 'GENERAUX','G_GENERAL', Ass.RepriseExploit) then
        AjouteCompte (TAss, Ass.RepriseExploit);
end; }


{***********A.G.L.***********************************************
Auteur  ...... : Bernadette Tynévez
Créé le ...... : 12/10/2007
Modifié le ... :   /  /
Description .. : Ajout compte associé dans TOB
Mots clefs ... :
*****************************************************************}
{procedure TOF_AMLISTEREGFR.AjouteCompte (TAss : TOB; Compte : string);
var TFille : TOB;
    i : integer;
    bTrouve : boolean;
begin
  bTrouve := false;

  for i:=0 to TAss.Detail.Count-1 do
  begin
    if TAss.Detail[i].GetValue ('COMPTE') = Compte then
    begin
      bTrouve := True;
      break;
    end;
  end;
  if not bTrouve then
  begin
    TFille := TOB.Create('', TAss, -1);
    TFille.AddChampSupValeur ('COMPTE', Compte);
  end;
end; }


{procedure TOF_AMLISTEREGFR.EcrireComptesAssocies (TAss : TOB);
var T: Tob;
    stGeneral, stLibelle : string;
    i: integer;
begin
  if TAss <> nil then
  begin
     T := TOB.Create ('GENERAUX', nil, -1);

     try

       for i :=0 to TAss.Detail.Count - 1 do
       begin

         if not Presence ('GENERAUX','G_GENERAL', TAss.Detail[i].GetValue('COMPTE')) then
         begin
           stGeneral := TAss.Detail[i].GetValue('COMPTE');
           stLibelle := stGeneral;
           // Création du compte
           T.PutValue ('G_GENERAL', stGeneral);
           T.PutValue ('G_LIBELLE', stLibelle);
//$IFDEF SERIE1
//$ELSE
           T.PutValue ('G_ABREGE',Copy (stLibelle,1,17));
           T.PutValue ('G_SENS','M');
//$ENDIF
           if stGeneral[1] = '2' then T.PutValue ('G_NATUREGENE','IMO')
           else if stGeneral[1] = '6' then T.PutValue ('G_NATUREGENE','CHA')
           else if stGeneral[1] = '7' then T.PutValue ('G_NATUREGENE','PRO')
           else T.PutValue ('G_NATUREGENE','DIV');
           T.PutValue ('G_VENTILABLE1','-'); T.PutValue ('G_VENTILABLE2','-');
           T.PutValue ('G_VENTILABLE3','-'); T.PutValue ('G_VENTILABLE4','-');
           T.PutValue ('G_VENTILABLE5','-');
//$IFDEF SERIE1
//$ELSE
           T.PutValue ('G_CREERPAR','REG');
//$ENDIF
           T.PutValue ('G_CONFIDENTIEL','0');
           T.InsertDB (nil);
         end;
       end;

     finally
       T.Free;
     end;

  end;
end; }


{***********A.G.L.***********************************************
Auteur  ...... : Bernadette Tynévez
Créé le ...... : 02/10/2007
Modifié le ... :   /  /
Description .. : ACTION Validation
Mots clefs ... :
*****************************************************************}
procedure TOF_AMLISTEREGFR.OnValidationClick(Sender: TObject);
begin
  OperationSerie (toValidREGFR);
end;



{////////////////////////////////////////////////////////////////////////////////
/////////////////////////////////////////////////////////////////////////////////

     P R O C E D U R E S / F O N C T I O N S   P U B L I Q U E S

/////////////////////////////////////////////////////////////////////////////////
////////////////////////////////////////////////////////////////////////////////}


{***********A.G.L.***********************************************
Auteur  ...... : Bernadette Tynévez
Créé le ...... : 02/10/2007
Modifié le ... :
Description .. : Ajout dans la table IMMO/IMMOLOG/IMMOAMOR
Mots clefs ... :
*****************************************************************}
procedure AMValidationImmo (CodeI : string; var fValid : integer; TAss : TOB);

  // ------------------------------------------------------------------------
  // Récup comptes associés du compte donné en entrée
  procedure AttribuerComptesAssocies (var RCompteAss: TCompteAss; T : TOB; CompteImmo :string);
  begin
  // Chercher les comptes associés au compte d'abord dans IMMOCPTE, sinon les attribuer par défaut
  RecupereComptesAssocies (nil, CompteImmo, RCompteAss);
  //GetComptesAssociesParDefaut(RCompteAss, CompteImmo);

  with T do
  begin
    PutValue('I_COMPTEAMORT', RCompteAss.Amort);
    PutValue('I_COMPTEDOTATION', RCompteAss.Dotation);
    PutValue('I_COMPTEDEROG', RCompteAss.Derog );
    PutValue('I_REPRISEDEROG', RCompteAss.RepriseDerog);
    PutValue('I_PROVISDEROG', RCompteAss.ProvisDerog);
    PutValue('I_DOTATIONEXC', RCompteAss.DotationExcep);
    PutValue('I_VACEDEE', RCompteAss.VaCedee);
    PutValue('I_AMORTCEDE', RCompteAss.AmortCede);
    PutValue('I_VAOACEDEE', RCompteAss.VoaCede);
    PutValue('I_REPEXPLOIT', RCompteAss.RepriseExploit);
    PutValue('I_REPEXCEP', RCompteAss.RepriseExcep);
  end;
  end;
  // ------------------------------------------------------------------------
  // Test comptes associés existent dans le plan de comptes
  procedure VerifComptesAssocies (Ass : TCompteAss; TAss : TOB);

    // Ajoute un compte dans la TOB TAss
    procedure AjouteCompte (TAss : TOB; Compte : string);
    var TFille : TOB;
      i : integer;
      bTrouve : boolean;
    begin
      bTrouve := false;

      for i:=0 to TAss.Detail.Count-1 do
      begin
        if TAss.Detail[i].GetValue ('COMPTE') = Compte then
        begin
          bTrouve := True;
          break;
        end;
      end;
      if not bTrouve then
      begin
        TFille := TOB.Create('', TAss, -1);
        TFille.AddChampSupValeur ('COMPTE', Compte);
      end;
    end;

  begin
  if not (Ass.Immo = '') then
     if not  Presence ( 'GENERAUX','G_GENERAL', Ass.Immo) then
        AjouteCompte (TAss, Ass.Immo);
  if not (Ass.Amort = '') then
     if not  Presence ( 'GENERAUX','G_GENERAL', Ass.Amort) then
        AjouteCompte (TAss, Ass.Amort);
  if not (Ass.Dotation = '') then
     if not  Presence ( 'GENERAUX','G_GENERAL', Ass.Dotation) then
        AjouteCompte (TAss, Ass.Dotation);
  if not (Ass.Derog = '') then
     if not  Presence ( 'GENERAUX','G_GENERAL', Ass.Derog) then
        AjouteCompte (TAss, Ass.Derog);
  if not (Ass.RepriseDerog = '') then
     if not  Presence ( 'GENERAUX','G_GENERAL', Ass.RepriseDerog) then
        AjouteCompte (TAss, Ass.RepriseDerog);
  if not (Ass.DotationExcep = '') then
     if not  Presence ( 'GENERAUX','G_GENERAL', Ass.DotationExcep) then
        AjouteCompte (TAss, Ass.DotationExcep);
  if not (Ass.VaCedee = '') then
     if not  Presence ( 'GENERAUX','G_GENERAL', Ass.VaCedee) then
        AjouteCompte (TAss, Ass.VaCedee);
  if not (Ass.AmortCede = '') then
     if not  Presence ( 'GENERAUX','G_GENERAL', Ass.AmortCede) then
        AjouteCompte (TAss, Ass.AmortCede);
  if not (Ass.VoaCede = '') then
     if not  Presence ( 'GENERAUX','G_GENERAL', Ass.VoaCede) then
        AjouteCompte (TAss, Ass.VoaCede);
  if not (Ass.ProvisDerog = '') then
     if not  Presence ( 'GENERAUX','G_GENERAL', Ass.ProvisDerog) then
        AjouteCompte (TAss, Ass.ProvisDerog);
  if not (Ass.RepriseExcep = '') then
     if not  Presence ( 'GENERAUX','G_GENERAL', Ass.RepriseExcep) then
        AjouteCompte (TAss, Ass.RepriseExcep);
  if not (Ass.RepriseExploit = '') then
     if not  Presence ( 'GENERAUX','G_GENERAL', Ass.RepriseExploit) then
        AjouteCompte (TAss, Ass.RepriseExploit);
  end;
  // ------------------------------------------------------------------------
  // Ajout comptes de TAss dans le plan de comptes si pas déjà fait pour une immo précédant celle_ci
  procedure EcrireComptesAssocies (TAss : TOB);
  var T: Tob;
    stGeneral, stLibelle : string;
    i: integer;
  begin
  if TAss <> nil then
  begin
     T := TOB.Create ('GENERAUX', nil, -1);

     try

       for i :=0 to TAss.Detail.Count - 1 do
       begin

         if not Presence ('GENERAUX','G_GENERAL', TAss.Detail[i].GetValue('COMPTE')) then
         begin
           stGeneral := TAss.Detail[i].GetValue('COMPTE');
           stLibelle := stGeneral;
           // Création du compte
           T.PutValue ('G_GENERAL', stGeneral);
           T.PutValue ('G_LIBELLE', stLibelle);
{$IFDEF SERIE1}
{$ELSE}
           T.PutValue ('G_ABREGE',Copy (stLibelle,1,17));
           T.PutValue ('G_SENS','M');
{$ENDIF}
           if stGeneral[1] = '2' then T.PutValue ('G_NATUREGENE','IMO')
           else if stGeneral[1] = '6' then T.PutValue ('G_NATUREGENE','CHA')
           else if stGeneral[1] = '7' then T.PutValue ('G_NATUREGENE','PRO')
           else T.PutValue ('G_NATUREGENE','DIV');
           T.PutValue ('G_VENTILABLE1','-'); T.PutValue ('G_VENTILABLE2','-');
           T.PutValue ('G_VENTILABLE3','-'); T.PutValue ('G_VENTILABLE4','-');
           T.PutValue ('G_VENTILABLE5','-');
{$IFDEF SERIE1}
{$ELSE}
           T.PutValue ('G_CREERPAR','REG');
{$ENDIF}
           T.PutValue ('G_CONFIDENTIEL','0');
           T.InsertDB (nil);
         end;
       end;

     finally
       T.Free;
     end;

  end;
  end;
  // ------------------------------------------------------------------------

var CodeImmo :string;
    DatePassage : TDateTime;
    T1, TCalcul, TFiche, TAmor : Tob;
    RCompteAss : TCompteAss;
    PremMois, PremAnnee, NbMois: Word;
    Plan : TPlanAmort;
    Infos: TInfoLog;
begin
  DatePassage := EncodeDate (StrToInt(GetParamSocSecur('SO_DATEREGREEL','')), 1, 1);

  // T1 = enreg lu dans IMMOREGFR

  T1 := TOB.Create('IMMOREGFR', nil, -1) ;
  T1.SelectDB ('"'+ CodeI +'"', nil);
  if ExisteSQL
     ('SELECT I_IMMO FROM IMMO WHERE I_IMMO="' + T1.GetValue('IR_COMPTEIMMO') + '"') then
     CodeImmo := NouveauCodeImmo
  else
     CodeImmo := T1.GetValue('IR_COMPTEIMMO');

  // Récup calculs durée amort., VNC, V retenue etc.
  TCalcul := TOB.Create('', nil, -1) ;
  AMCalculREGFR (T1, TCalcul);


  // TABLE IMMO --------------------------------------------------------------------------

  TFiche := TOB.Create('IMMO', nil, -1) ;
  //TAss := TOB.Create('', nil, -1) ;
  TAmor  := TOB.Create('', nil, -1);

  try
    with TFiche do
    begin
    PutValue('I_IMMO', CodeImmo);
    PutValue('I_LIBELLE', T1.GetValue('IR_LIBELLE'));

    PutValue('I_COMPTEIMMO', T1.GetValue('IR_COMPTEREF'));
    PutValue('I_MONTANTHT', TCalcul.GetValue('VRETENUE'));
    PutValue('I_MONTANTBASEAMORT', TCalcul.GetValue('VRETENUE'));
    PutValue('I_NATUREIMMO', 'PRO');
    PutValue('I_ABREGE', Copy (T1.GetValue('IR_LIBELLE'), 1, 17));
    PutValue('I_DATEPIECEA', DatePassage);
    PutValue('I_DATEAMORT', DatePassage);
    PutValue('I_BASETAXEPRO', TCalcul.GetValue('VRETENUE'));

    PutValue('I_BASEECO', TCalcul.GetValue('VRETENUE'));
    PutValue('I_METHODEECO', T1.GetValue('IR_METHODEECO'));
    PutValue('I_DUREEECO', T1.GetValue('IR_DUREERESTM'));
    PutValue('I_TAUXECO', GetTaux(T1.GetValue('IR_METHODEECO'),
                                     DatePassage,
                                     DatePassage,
                                     T1.GetValue('IR_DUREERESTM')));
    AttribuerComptesAssocies (RCompteAss, TFiche, T1.GetValue('IR_COMPTEREF'));
    // Mise en TOB des cptes absents du plan de cptes
    VerifComptesAssocies (RCompteAss, TAss);
    PutValue('I_SOCIETE', V_PGI.CodeSociete);
    if T1.GetValue('IR_METHODEECO') = 'LIN' then
       PutValue('I_JOURNALA', '***');
    PutValue('I_QUALIFIMMO', 'R');
    PutValue('I_ETABLISSEMENT', VHImmo^.EtablisDefaut);
    PutValue('I_QUANTITE', 1);
    PutValue('I_PAYS', 'FRA');
    PutValue('I_LIEUGEO', '001');
    PutValue('I_TYPELOYERCB', 'LCO');
    PutValue('I_PERIODICITE', 'MEN');
    PutValue('I_VERSEMENTCB', 'AVA');
    PutValue('I_DATEDEBUTECH', DatePassage);
    PutValue('I_DATEFINECH', VHImmo^.Encours.Fin);
    PutValue('I_NATUREBIEN', 'GEB');
    PutValue('I_ETAT', 'OUV');
    PutValue('I_DEVISE', V_PGI.DevisePivot);
    PutValue('I_OPERATION', '-');
    PutValue('I_OPEMUTATION', '-');
    PutValue('I_OPEECLATEMENT', '-');
    PutValue('I_OPECESSION', '-');
    PutValue('I_OPECHANGEPLAN', '-');
    PutValue('I_OPELIEUGEO', '-');
    PutValue('I_OPEETABLISSEMENT', '-');
    PutValue('I_OPELEVEEOPTION', '-');
    PutValue('I_OPEMODIFBASES', '-');
    PutValue('I_OPEDEPREC', '-');
    PutValue('I_OPEREG', '-');
    PutValue('I_OPEPRIME', '-');
    PutValue('I_OPEREDUCPRI', '-');
    PutValue('I_OPESBV', '-');
    PutValue('I_OPEARD', '-');
    PutValue('I_OPEREMPL', '-');
    PutValue('I_VENTILABLE', '-');
    PutValue('I_VENTILABLE1', '-');
    PutValue('I_VENTILABLE2', '-');
    PutValue('I_VENTILABLE3', '-');
    PutValue('I_VENTILABLE4', '-');
    PutValue('I_VENTILABLE5', '-');
    NOMBREMOIS(DatePassage, VHImmo^.Encours.Deb, PremMois, PremAnnee, NbMois);
    PutValue('I_DUREEREPRISE', MaxIntValue([0, NbMois - 1]));
    PutValue('I_BASEAMORDEBEXO', TCalcul.GetValue('VRETENUE'));
    PutValue('I_BASEAMORFINEXO', TCalcul.GetValue('VRETENUE'));
    PutValue('I_COMPTEREF', T1.GetValue('IR_COMPTEREF'));
    PutValue('I_VAOACEDEE', T1.GetValue('IR_COMPTEREF'));
    PutValue('I_VALEURACHAT', TCalcul.GetValue('VRETENUE'));
    PutValue('I_SURAMORT', '-');
    PutValue('I_SUSDEF','A');
    PutValue('I_REGLECESSION','NOR');
    PutValue('I_NONDED','-');
    PutValue('I_DATEDEBECO', DatePassage);
    PutValue('I_DATEDEBFIS', DatePassage);
    PutValue('I_DPI','-');
    PutValue('I_DPIEC', '-');
    PutValue('I_SUBVENTION', 'NON');
    PutValue('I_SBVEC', '-');
    PutValue('I_PFR', 'X');             // immo issue du régime forfait
    PutValue('I_ACHFOR', T1.GetValue('IR_DATEACHAT'));
    PutValue('I_PRIXACFOR', T1.GetValue('IR_VO'));
    PutValue('I_PRIXACFORC', TCalcul.GetValue('VO_Reevaluee'));
    PutValue('I_VNCFOR', TCalcul.GetValue('VRETENUE'));
    PutValue('I_DURRESTFOR', T1.GetValue('IR_DUREERESTM')); // TCalcul.GetValue('DUREE_AMORT'));
    PutValue('I_VERSION', V_PGI.NumVersion);
    PutValue('I_ARD', '-');
    PutValue('I_BLOCNOTE', T1.GetValue('IR_BLOCNOTE'));
    PutValue('I_AMTFORC', TCalcul.GetValue('AMORT'));
    end;

// TABLE IMMOAMOR --------------------------------------------------------------------------

    Plan := TPlanAmort.Create(true);
    Plan.CalculTOB( TFiche, iDate1900);
    Plan.SauveTOB ( TAmor );
    TFiche.PutValue('I_PLANACTIF', Plan.NumSeq);
    TFiche.PutValue('I_DATEDERMVTECO', Plan.GetDateFinAmortEx(Plan.AmortEco));

// ENREGISTREMENT ---------------------------------------------------------------------------

    TFiche.InsertDB(nil);
    TAmor.InsertDB(nil);

// TABLE IMMOLOG --------------------------------------------------------------------------

    Infos.TVARecuperable := 0;
    Infos.TVARecuperee := 0;
    EnregLogAcquisition (CodeImmo, DatePassage, Plan.NumSeq, Infos);
    FreeAndNil( Plan );
    fValid := fValid + 1;
    EcrireComptesAssocies (TAss);

// SUPPRESSION IMMOREGFR ---------------------------------------------------------------

   ExecuteSQL('DELETE FROM IMMOREGFR WHERE IR_COMPTEIMMO = "' + CodeI + '"');

  finally
    T1.Free ;
    TCalcul.Free ;
    TFiche.Free ;
    TAmor.Free ;
  end;
end;





Initialization
  registerclasses ( [ TOF_AMLISTEREGFR ] ) ;
end.
