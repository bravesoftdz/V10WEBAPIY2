{***********UNITE*************************************************
Auteur  ...... : Bruno DUCLOS
Créé le ...... : 19/04/2007
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : AFMAJCHAMPSAUTO ()
Mots clefs ... : TOF;AFMAJCHAMPSAUTO
*****************************************************************}
Unit UtofAfMajChampsAuto ;

Interface

Uses StdCtrls,
     Controls,
     Classes,
     {$IFNDEF EAGLCLIENT}
       db,
       {$IFNDEF DBXPRESS}
         dbtables,
       {$ELSE}
         uDbxDataSet,
       {$ENDIF}
       mul,
     {$ELSE}
       eMul,
     {$ENDIF}
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HEnt1,
     HMsgBox,
     UTOF,
     Htb97,
     UTob,
     Vierge,
     ParamSoc,
     {$IFNDEF EAGLCLIENT}
     Fe_Main
     {$ELSE}
     MainEAgl
     {$ENDIF}
     ;

Type
  TTypeMaj = (tmGeneral, tmTiers);

  TChampMajAuto = class
    constructor Create(ptmTypeMaj: TTypeMaj); overload;
    destructor Destroy; override;
  private
    ftbChampMaj: TOB;
    ftmTypeMaj:  TTypeMaj;
    fboAuMoinsUn: Boolean;
  end;

  TOF_AFMAJCHAMPSAUTO = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnDisplay                ; override ;
    procedure OnClose                  ; override ;
    procedure OnCancel                 ; override ;
    procedure ListeDisponibleOnClick(Sender: TObject);
    procedure ListeDisponibleOnDblClick(Sender: TObject);
    procedure ListeAffecteOnClick(Sender: TObject);
    procedure ListeAffecteOnDblClick(Sender: TObject);
    procedure btnVersDroiteOnClick(Sender: TObject);
    procedure btnVersGaucheOnClick(Sender: TObject);
    procedure btnValiderOnClick(Sender: TObject);
    procedure btnVersDroiteTousOnClick(Sender: TObject);
    procedure btnVersGaucheTousOnClick(Sender: TObject);
    procedure Maj;
  private
    btnValider,
    btnVersDroite,
    btnVersGauche,
    btnVersDroiteTous,
    btnVersGaucheTous: TToolBarButton97;
    lstDisponible,
    lstAffecte: THListBox;
    ListeArguments: TStringList;
  end;

function ReferenceObjet(Reference: String; var Index: Integer): TChampMajAuto;
function MajAutoCommence(ptmTypeMaj: TTypeMaj): String;
procedure MajAutoTermine(pboMenu: Boolean; Reference: String);
function MajAutoSelectionneChamp(pstArgument, Reference: String): String;
function MajAutoValeur(pstChamp, pstQuoi, Reference: String): String;
procedure MajAutoAjoute(pslListe: TStringList; pstSource, pstDestination, Reference: String);
procedure MajAutoChangeValeur(pstChamp: String; pvaValeur: Variant; Reference: String);
procedure MajAutoInit(Reference: String);
function MajAutoTraitement(pboConfirmation: Boolean; Reference: String): Boolean;
function MajAutoChampTraite(pstChamp, Reference: String): Boolean;

Implementation

uses TntWideStrings,
  TntStdCtrls,
  DicoBTP
 ,CbpMCD
 ,CbpEnumerator
;

var
  ListeObjetChamp: TStringList;
  objChamp: TChampMajAuto;

const
  NbMessage = 6;
  TexteMessage: Array[1..NbMessage] of String = (
    'Déplace le champ dans la liste des champs à mettre à jour',
    'Déplace le champ dans la liste des champs disponibles',
    'Déplace tous les champs dans la liste des champs à mettre à jour',
    'Déplace tous les champs dans la liste des champs disponibles',
    'Des informations présentes dans vos affaires ont été modifiées, ' +
      'souhaitez-vous lancer la mise à jour ?',
    'affaire'
    );

procedure TOF_AFMAJCHAMPSAUTO.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_AFMAJCHAMPSAUTO.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_AFMAJCHAMPSAUTO.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_AFMAJCHAMPSAUTO.OnLoad ;
begin
  Inherited ;
end ;

procedure TOF_AFMAJCHAMPSAUTO.OnArgument (S : String ) ;
var
  vinIndex,
  vinCompteur: Integer;
  vstTemp: String;
begin
  Inherited ;
  S := UpperCase(S);
  // Prépare la liste des disponibles
  if Assigned(GetControl('LISTEDISPONIBLE')) then
  begin
    lstDisponible := THListBox(GetControl('LISTEDISPONIBLE'));
    lstDisponible.OnClick := ListeDisponibleOnClick;
    lstDisponible.OnDblClick := ListeDisponibleOnDblClick;
  end;
  // Prépare la liste des affectés
  if Assigned(GetControl('LISTEAFFECTE')) then
  begin
    lstAffecte := THListBox(GetControl('LISTEAFFECTE'));
    lstAffecte.OnClick := ListeAffecteOnClick;
    lstAffecte.OnDblClick := ListeAffecteOnDblClick;
  end;
  // Prépare le bouton de transfert vers la liste des affectés
  if Assigned(GetControl('BTNVERSDROITE')) then
  begin
    btnVersDroite := TToolBarButton97(GetControl('BTNVERSDROITE'));
    btnVersDroite.OnClick := btnVersDroiteOnClick;
    btnVersDroite.Hint := TraduitGa(TexteMessage[1]);
    btnVersDroite.ShowHint := True;
  end;
  // Prépare le bouton de transfert vers la liste des disponibles
  if Assigned(GetControl('BTNVERSGAUCHE')) then
  begin
    btnVersGauche := TToolBarButton97(GetControl('BTNVERSGAUCHE'));
    btnVersGauche.OnClick := btnVersGaucheOnClick;
    btnVersGauche.Hint := TraduitGa(TexteMessage[2]);
    btnVersGauche.ShowHint := True;
  end;
  // Prépare le bouton de transfert vers la liste des affectés
  if Assigned(GetControl('BTNVERSDROITETOUS')) then
  begin
    btnVersDroiteTous := TToolBarButton97(GetControl('BTNVERSDROITETOUS'));
    btnVersDroiteTous.OnClick := btnVersDroiteTousOnClick;
    btnVersDroiteTous.Hint := TraduitGa(TexteMessage[3]);
    btnVersDroiteTous.ShowHint := True;
  end;
  // Prépare le bouton de transfert vers la liste des disponibles
  if Assigned(GetControl('BTNVERSGAUCHETOUS')) then
  begin
    btnVersGaucheTous := TToolBarButton97(GetControl('BTNVERSGAUCHETOUS'));
    btnVersGaucheTous.OnClick := btnVersGaucheTousOnClick;
    btnVersGaucheTous.Hint := TraduitGa(TexteMessage[4]);
    btnVersGaucheTous.ShowHint := True;
  end;
  // Prépare le bouton de validation
  if Assigned(GetControl('BVALIDER')) then
  begin
    btnValider := TToolBarButton97(GetControl('BVALIDER'));
    btnValider.OnClick := btnValiderOnClick;
  end;
  ListeArguments := TStringList.Create;
  S := StringReplace(S, ';', #1310, [rfReplaceAll]);
  ListeArguments.Text := S;
  objChamp := ReferenceObjet(ListeArguments.Values['OBJ'], vinIndex);
  if Assigned(objChamp) then
  begin
    // Charge les champs dans les listes
    for vinCompteur := 0 to objChamp.ftbChampMaj.Detail.Count - 1 do
    begin
      vstTemp := objChamp.ftbChampMaj.Detail[vinCompteur].GetString('TABLE');
      // Les champs de la table LIGNE ne sont pas affichés, ils sont rajoutés ensuite
      // en fonction des champs de la table PIECE sélectionné
      if (vstTemp <> 'LIGNE') and (vstTemp <> 'LIGNECOMPL') then // GA_20080520_BDU_GA3613. Ajout de LIGNECOMPL
      begin
        // BDU - 09/10/07. Ajout des champs de la table AFFAIRE avec prise en compte PARAMSOC
        if objChamp.ftbChampMaj.Detail[vinCompteur].GetString('TABLE') = 'AFFAIRE' then
        begin
          vstTemp := Format('%s dans %s', [objChamp.ftbChampMaj.Detail[vinCompteur].GetString('LIBELLE'),
            TraduitGa(TexteMessage[6])]);
        end
        else
          vstTemp := objChamp.ftbChampMaj.Detail[vinCompteur].GetString('LIBELLE');
        // Appel depuis le menu, tous les champs sont sélectionnés
        if objChamp.ftmTypeMaj = tmGeneral then
          lstAffecte.Items.AddObject(vstTemp,
            TObject(objChamp.ftbChampMaj.Detail[vinCompteur].GetInteger('NUMERO')))
        // Appel depuis la fiche tiers, les champs modifiés sont sélectionnés
        else if objChamp.ftmTypeMaj = tmTiers then
        begin
          if objChamp.ftbChampMaj.Detail[vinCompteur].GetBoolean('MODIFIE') then
            lstAffecte.Items.AddObject(vstTemp,
              TObject(objChamp.ftbChampMaj.Detail[vinCompteur].GetInteger('NUMERO')));
        end;
      end;
    end;
  end;
  if lstDisponible.Items.Count > 0 then
    lstDisponible.ItemIndex := 0;
  if lstAffecte.Items.Count > 0 then
    lstAffecte.ItemIndex := 0;
  Maj;
end ;

procedure TOF_AFMAJCHAMPSAUTO.OnClose ;
begin
  Inherited ;
  FreeAndNil(ListeArguments);
end ;

procedure TOF_AFMAJCHAMPSAUTO.OnDisplay () ;
begin
  Inherited ;
  //
end ;

procedure TOF_AFMAJCHAMPSAUTO.OnCancel () ;
begin
  Inherited ;
end ;

procedure TOF_AFMAJCHAMPSAUTO.ListeAffecteOnClick(Sender: TObject);
begin
  Maj;
end;

procedure TOF_AFMAJCHAMPSAUTO.ListeAffecteOnDblClick(Sender: TObject);
begin
  btnVersGaucheOnClick(nil);
  Maj;
end;

procedure TOF_AFMAJCHAMPSAUTO.ListeDisponibleOnClick(Sender: TObject);
begin
  Maj;
end;

procedure TOF_AFMAJCHAMPSAUTO.ListeDisponibleOnDblClick(Sender: TObject);
begin
  btnVersDroiteOnClick(nil);
  Maj;
end;

procedure TOF_AFMAJCHAMPSAUTO.Maj;
begin
  // Le bouton est actif si il y a au moins un élément dans les disponibles
  if Assigned(btnVersDroite) then
    btnVersDroite.Enabled := (lstDisponible.Items.Count > 0) and (lstDisponible.ItemIndex <> -1);
  // Le bouton est actif si il y a au moins un élément dans les affectés
  if Assigned(btnVersGauche) then
    btnVersGauche.Enabled := (lstAffecte.Items.Count > 0) and (lstAffecte.ItemIndex <> -1);
  // Le bouton est actif si il y a au moins un élément dans les affectés
  if Assigned(btnValider) then
    btnValider.Enabled := lstAffecte.Items.Count > 0;
  // Le bouton est actif si il a au moins deux éléments dans les disponibles
  if Assigned(btnVersDroiteTous) then
    btnVersDroiteTous.Enabled := LstDisponible.Items.Count > 1;
  // Le bouton est actif si il a au moins deux éléments dans les affectés
  if Assigned(btnVersGaucheTous) then
    btnVersGaucheTous.Enabled := LstAffecte.Items.Count > 1;
end;

procedure TOF_AFMAJCHAMPSAUTO.btnVersDroiteOnClick(Sender: TObject);
var
  vinDisponible,
  vinAffecte: Integer;
begin
  vinDisponible := lstDisponible.ItemIndex;
  if vinDisponible <> -1 then
  begin
    // Copie l'élément sélectionné vers la liste affecté
    vinAffecte := lstAffecte.Items.AddObject(lstDisponible.Items[vinDisponible],
      lstDisponible.Items.Objects[vinDisponible]);
    lstAffecte.ItemIndex := vinAffecte;
    // Supprime l'élément sélectionné
    lstDisponible.Items.Delete(vinDisponible);
    if vinDisponible < lstDisponible.Items.Count - 1 then
      lstDisponible.ItemIndex := vinDisponible
    else
      lstDisponible.ItemIndex := lstDisponible.Items.Count - 1;
  end;
  Maj;
end;

procedure TOF_AFMAJCHAMPSAUTO.btnVersGaucheOnClick(Sender: TObject);
var
  vinDisponible,
  vinAffecte: Integer;
begin
  //
  vinAffecte := lstAffecte.ItemIndex;
  if vinAffecte <> -1 then
  begin
    // Copie l'élément sélectionné vers la liste disponible
    vinDisponible := lstDisponible.Items.AddObject(lstAffecte.Items[vinAffecte],
      lstAffecte.Items.Objects[vinAffecte]);
    lstDisponible.ItemIndex := vinDisponible;
    // Supprime l'élément sélectionné
    lstAffecte.Items.Delete(vinAffecte);
    if vinAffecte < lstAffecte.Items.Count - 1 then
      lstAffecte.ItemIndex := vinAffecte
    else
      lstAffecte.ItemIndex := lstAffecte.Items.Count - 1;
  end;
  Maj;
end;

procedure TOF_AFMAJCHAMPSAUTO.btnVersDroiteTousOnClick(Sender: TObject);
var
  vinCompteur: Integer;
begin
  for vinCompteur := lstDisponible.Items.Count - 1 downto 0 do
  begin
    lstDisponible.ItemIndex := vinCompteur;
    btnVersDroiteOnClick(nil);
  end;
end;

procedure TOF_AFMAJCHAMPSAUTO.btnVersGaucheTousOnClick(Sender: TObject);
var
  vinCompteur: Integer;
begin
  for vinCompteur := lstAffecte.Items.Count - 1 downto 0 do
  begin
    lstAffecte.ItemIndex := vinCompteur;
    btnVersGaucheOnClick(nil);
  end;
end;

procedure TOF_AFMAJCHAMPSAUTO.btnValiderOnClick(Sender: TObject);
var
  vinNumero,
  vinCompteur: Integer;
  vstSeparateur,
  vstTemp: String;
  vtbCourante: TOB;
begin
  vstTemp := '';
  vstSeparateur := '';
  for vinCompteur := 0 to lstAffecte.Items.Count - 1 do
  begin
    // Récupère le numéro d'index
    vinNumero := Integer(lstAffecte.Items.Objects[vinCompteur]);
    // Cherche dans la tob le champ ayant ce numéro d'index
    vtbCourante := objChamp.ftbChampMaj.FindFirst(['NUMERO'], [vinNumero], False);
    vstTemp := Format('%s%s%s', [vstTemp, vstSeparateur, vtbCourante.GetString('CO_LIBRE')]);
    vstSeparateur := ';';
  end;
  TFVierge(Ecran).Retour := vstTemp;
end;

{ TChampMajAuto }

constructor TChampMajAuto.Create(ptmTypeMaj: TTypeMaj);
var
  vinIndex,
  vinCompteur: Integer;
  vstTemp, // BDU - 09/10/07
  vstRequete,
  vstLibelle: String;
  vtbNouvelle: TOB; // BDU - 09/10/07
Mcd : IMCDServiceCOM;
begin
  inherited Create;
  MCD := TMCD.GetMcd;
  if not mcd.loaded then mcd.WaitLoaded();
  //
  ftmTypeMaj := ptmTypeMaj;
  fboAuMoinsUn := False;
  ftbChampMaj := TOB.Create('les champs', nil, -1);
  vstRequete := 'SELECT CO_LIBELLE, CO_LIBRE, CO_ABREGE ' +
    'FROM COMMUN ' +
    // Le type de la tablette AFMAJAUTO
    'WHERE CO_TYPE = "MAA" ';
  // Mise à jours des champs de la table TIERS vers les tables PIECE, LIGNE et AFFAIRE
  if ptmTypeMaj in [tmTiers, tmGeneral] then
  begin
    // Les champs "source" doivent faire partie des tables TIERS ou TIERSCOMPL
    vstRequete := vstRequete + 'AND ((CO_LIBELLE LIKE "T\_%" ESCAPE "\") OR (CO_LIBELLE LIKE "YTC\_%" ESCAPE "\")) AND ' +
      // Les champs "destination" doivent faire partie des tables PIECE, LIGNE ou AFFAIRE
      // GA_20080520_BDU_GA3613. Ajout LIGNECOMPL
      '((CO_LIBRE LIKE "GP\_%" ESCAPE "\") OR (CO_LIBRE LIKE "GL\_%" ESCAPE "\") ' +
      'OR (CO_LIBRE LIKE "AFF\_%" ESCAPE "\") OR (CO_LIBRE LIKE "GLC\_%" ESCAPE "\")) AND ' +
      // Le code doit commencer par A
      '(CO_CODE LIKE "A%")';
    // BDU - 25/04/07 - Demande CCO, représentant selon paramètre société
    if not GetParamSocSecur('SO_GERECOMMERCIAL', False) then
    begin
      vstRequete := vstRequete + ' AND (CO_LIBELLE NOT LIKE "%_REPRESENTANT%")';
      vstRequete := vstRequete + ' AND (CO_LIBELLE NOT LIKE "YTC_TAUXREPR%")'; // GA_20080520_BDU_GA3613
    end;
    // GA_20080520_BDU_GA3613_DEBUT.
    // Ces champ ne sont pas visibles si pas de gestion des commerciaux ou commission à la ligne
    if not GetParamSocSecur('SO_GERECOMMERCIAL', False) or
      GetParamSocSecur('SO_COMMISSIONLIGNE', False) then
    begin
      vstRequete := vstREquete + ' AND (CO_LIBRE NOT LIKE "GL_REPRESENTANT%")' +
        ' AND (CO_LIBRE NOT LIKE "GLC_REPRESENTANT%")' +
        ' AND (CO_LIBRE NOT LIKE "GL_COMMISSIONR%")' +
        ' AND (CO_LIBRE NOT LIKE "GLC_COMMISSIONR%")';
    end;
    // GA_20080520_BDU_GA3613_FIN
  end
  // Si aucun paramètre, la requête ne renverra rien
  else
    vstRequete := vstRequete + 'AND (1 = 0)';
  ftbChampMaj.LoadDetailDBFromSQL('les champs maj', vstRequete);
  // BDU - 09/10/07. Ajout des champs de la table AFFAIRE avec prise en compte PARAMSOC
  // Le champ Responsable
  vstTemp := getParamSocSecur('SO_AFRESPAFF', '');
  if Copy(vstTemp, 1, 2) = 'RE' then
  begin
    vstTemp := Copy(vstTemp, 3, 1);
    vtbNouvelle := TOB.Create('les champs maj', ftbChampMaj, -1);
    vtbNouvelle.AddChampSupValeur('CO_LIBELLE', Format('YTC_RESSOURCE%s', [vstTemp]));
    vtbNouvelle.AddChampSupValeur('CO_LIBRE', 'AFF_RESPONSABLE'); // Format('AFF_RESPONSABLE', [vinCompteur]));
    vtbNouvelle.AddChampSupValeur('CO_ABREGE', '');
  end;
  // Les champs ressources
  for vinCompteur := 1 to 3 do
  begin
    vstTemp := GetParamSocSecur(Format('SO_AFRESS%dAFF', [vinCompteur]), '');
    if Copy(vstTemp, 1, 2) = 'RE' then
    begin
      vstTemp := Copy(vstTemp, 3, 1);
      vtbNouvelle := TOB.Create('les champs maj', ftbChampMaj, -1);
      vtbNouvelle.AddChampSupValeur('CO_LIBELLE', Format('YTC_RESSOURCE%s', [vstTemp]));
      vtbNouvelle.AddChampSupValeur('CO_LIBRE', Format('AFF_RESSOURCE%d', [vinCompteur]));
      vtbNouvelle.AddChampSupValeur('CO_ABREGE', '');
    end;
  end;
  for vinCompteur := 0 to ftbChampMaj.Detail.Count - 1 do
  begin
    // Numéro d'index utilisé pour retrouver les informations du champ lors de la validation
    ftbChampMaj.Detail[vinCompteur].AddChampSupValeur('NUMERO', vinCompteur);
    // Valeur après modification.
    ftbChampMaj.Detail[vinCompteur].AddChampSupValeur('VALEUR', '');
    // Le champ n'est pas modifié
    ftbChampMaj.Detail[vinCompteur].AddChampSupValeur('MODIFIE', False);
    // Récupération du libellé du champ
    vstLibelle := ChampToLibelle(ftbChampMaj.Detail[vinCompteur].GetString('CO_LIBRE'));
    ftbChampMaj.Detail[vinCompteur].AddChampSupValeur('LIBELLE', vstLibelle);
    // Récupération du nom de la table
    vstLibelle := ftbChampMaj.Detail[vinCompteur].GetString('CO_LIBRE');
    vstLibelle := Copy(vstLibelle, 1, Pos('_', vstLibelle) - 1);
    vstLibelle := PrefixeToTable(vstLibelle);
    ftbChampMaj.Detail[vinCompteur].AddChampSupValeur('TABLE', vstLibelle);
    // Type du champ
    vinIndex := ChampToNum(ftbChampMaj.Detail[vinCompteur].GetString('CO_LIBRE'));
    ftbChampMaj.Detail[vinCompteur].AddChampSupValeur('TYPE', mcd.getField(ftbChampMaj.Detail[vinCompteur].GetString('CO_LIBRE')).tipe);
  end;
end;

destructor TChampMajAuto.Destroy;
begin
  FreeAndNil(ftbChampMaj);
  inherited Destroy;
end;

function MajAutoSelectionneChamp(pstArgument, Reference: String): String;
begin
  // Affiche la fenêtre de sélection des champs
  Result := AGLLanceFiche('AFF', 'AFMAJCHAMPSAUTO', '', '', pstArgument + 'OBJ=' + Reference);
end;

function MajAutoValeur(pstChamp, pstQuoi, Reference: String): String;
var
  vtbCourante: TOB;
  vinIndex: Integer;
begin
  // Récupère la valeur de l'un des champs contenu dans la TOB de l'objet
  Result := '';
  objChamp := ReferenceObjet(Reference, vinIndex);
  if Assigned(objChamp) then
  begin
    vtbCourante := objChamp.ftbChampMaj.FindFirst(['CO_LIBRE'], [pstChamp], False);
    if Assigned(vtbCourante) then
      Result := vtbCourante.GetString(pstQuoi);
  end;
end;

function MajAutoCommence(ptmTypeMaj: TTypeMaj): String;
var
  S: String;
begin
  if not Assigned(ListeObjetChamp) then
    ListeObjetChamp := TStringList.Create;
  objChamp := TChampMajAuto.Create(ptmTypeMaj);
  S := '';
  if Assigned(objChamp) then
  begin
    S := FormatDateTime('hhnnsszzz', Time);
    ListeObjetChamp.AddObject(S, objChamp);
  end;
  Result := S;
end;

procedure MajAutoTermine(pboMenu: Boolean; Reference: String);
var
  vinIndex: Integer;
begin
  objChamp := ReferenceObjet(Reference, vinIndex);
  if Assigned(objChamp) and ((pboMenu and (objChamp.ftmTypeMaj = tmGeneral)) or
    ((not pboMenu) and (objChamp.ftmTypeMaj = tmTiers))) then
  begin
    FreeAndNil(objChamp);
    ListeObjetChamp.Delete(vinIndex);
  end;
  if ListeObjetChamp.Count = 0 then // BDU - 08/01/2008. Libération de la liste si elle est vide
    FreeAndNil(ListeObjetChamp);
end;

procedure MajAutoAjoute(pslListe: TStringList; pstSource, pstDestination, Reference: String);
var
  vinIndex,
  vinCompteur: Integer;
  vstTemp: String;
  vtbCourante: TOB;
begin
  objChamp := ReferenceObjet(Reference, vinIndex);
  if Assigned(objChamp) then
  begin
    for vinCompteur := 0 to pslListe.Count - 1 do
    begin
      vstTemp := pslListe[vinCompteur];
      // Si le préfixe du champ est égal au préfixe passé dans pstSource
      if Copy(vstTemp, 1, Length(pstSource)) = pstSource then
      begin
        // Remplace le préfixe pstSource par le préfixe pstDestination (ex : GP_ESCOMPTE -> GL_ESCOMPTE)
        vstTemp := StringReplace(vstTemp, pstSource, pstDestination, [rfIgnoreCase]);
        // Si le champ avec le préfixe pstDestination existe, on l'ajoute dans la liste
        vtbCourante := objChamp.ftbChampMaj.FindFirst(['CO_LIBRE'], [vstTemp], False);
        if Assigned(vtbCourante) then
          pslListe.Add(vstTemp);
      end;
    end;
  end;
end;

procedure MajAutoChangeValeur(pstChamp: String; pvaValeur: Variant; Reference: String);
var
  vinIndex: Integer;
  vtbCourante: TOB;
begin
  objChamp := ReferenceObjet(Reference, vinIndex);
  if Assigned(objChamp) then
  begin
    // Pour tous les champs dont le CO_LIBELLE est égale au paramètre pstChamp
    vtbCourante := objChamp.ftbChampMaj.FindFirst(['CO_LIBELLE'], [pstChamp], False);
    while Assigned(vtbCourante) do
    begin
      // Modification de la valeur courante
      vtbCourante.SetString('VALEUR', pvaValeur);
      // Le flag MODIFIE est mis à VRAI
      vtbCourante.SetBoolean('MODIFIE', True);
      vtbCourante := objChamp.ftbChampMaj.FindNext(['CO_LIBELLE'], [pstChamp], False);
      objChamp.fboAuMoinsUn := True;
    end;
  end;
end;

procedure MajAutoInit(Reference: String);
var
  vinIndex,
  vinCompteur: Integer;
begin
  objChamp := ReferenceObjet(Reference, vinIndex);
  if Assigned(objChamp) then
  begin
    for vinCompteur := 0 to objChamp.ftbChampMaj.Detail.Count - 1 do
    begin
      // La valeur courante est vidée
      objChamp.ftbChampMaj.Detail[vinCompteur].SetString('VALEUR', '');
      // Le flag MODIFIE est mis à FAUX
      objChamp.ftbChampMaj.Detail[vinCompteur].SetBoolean('MODIFIE', False);
    end;
    objChamp.fboAuMoinsUn := False;
  end;
end;

function MajAutoTraitement(pboConfirmation: Boolean; Reference: String): Boolean;
var
  vinIndex: Integer;
begin
  Result := False;
  objChamp := ReferenceObjet(Reference, vinIndex);
  if Assigned(objChamp) then
  begin
    Result := objChamp.fboAuMoinsUn;
    if Result and pboConfirmation then
      Result := PGIAskAF(TexteMessage[5], '') = mrYes;
  end;
end;

function MajAutoChampTraite(pstChamp, Reference: String): Boolean;
var
  vinIndex: Integer;
begin
  Result := False;
  objChamp := ReferenceObjet(Reference, vinIndex);
  if Assigned(objChamp) then
    Result := Assigned(objChamp.ftbChampMaj.FindFirst(['CO_LIBELLE'], [pstChamp], False));
end;

function ReferenceObjet(Reference: String; var Index: Integer): TChampMajAuto;
var
  X: Integer;
begin
  Result := nil;
  X := ListeObjetChamp.IndexOf(Reference);
  if X <> -1 then
  begin
    Result := TChampMajAuto(ListeObjetChamp.Objects[X]);
    Index := X;
  end;
end;
Initialization
  RegisterClasses ( [ TOF_AFMAJCHAMPSAUTO ] ) ;
end.

