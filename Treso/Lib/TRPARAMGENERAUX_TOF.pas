{-------------------------------------------------------------------------------------
    Version   |  Date   | Qui | Commentaires
--------------------------------------------------------------------------------------
08.10.001.001  02/08/07   JP   Création de l'unité : Gestion des ParamSoc communs
--------------------------------------------------------------------------------------}
unit TRPARAMGENERAUX_TOF;

{FONCTIONNEMENT
 1/ Mettre Le composant idoine sur la fiche portant le nom du ParamSoc
 2/ Vérifier dans les methodes privées que le Type du composant est bien géré
 3/ S'il y a des traitements spécifiques les déclarer dans la partie public de la classe
    et laisser le Private pour les fonctions génériques
 4/ Ne pas écrire directement dans la TOB, FTobParamSoc mais utiliser SetValeur
 5/ Pour brancher un évènement sur un composant, le faire dans AffecteEvenement
 6/ Pour récupérer des composants
 7/ AfterRead : pour des traitements Spécifiques après le chargement des ParamSoc.
 8/ BeforeWrite : pour d'éventuels traitements de validation
 9/ les champs de la TOB
      - PARAMSOC : Nom du ParamSoc (<=> au nom du composant)
      - OLDVALUE : Valeur au chargement (à ne pas modifier)
      - NEWVALUE : Dernière valeur stockée
      - MODIFIE  : Si la valeur a été modifiée
      - TYPEPS   : Type de ParamSoc : valeurs définies en constantes dans la partie implementation
 }

interface

uses
  {$IFDEF EAGLCLIENT}
  MaineAGL,
  {$ELSE}
  FE_MAIN, {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
  {$ENDIF}
  StdCtrls, Controls, UTob, Classes, SysUtils, UTOF, ComCtrls;


type
  TOF_TRPARAMGENERAUX = class (TOF)
    procedure OnArgument (S : string); override;
    procedure OnUpdate               ; override;
    procedure OnClose                ; override;
    procedure OnLoad                 ; override;
  private
    FTobParamSoc : TOB;
    FCanClose    : Boolean;

    {Regarde si le composant correspond à un ParamSoc}
    function TesteComposant(Cmp : TComponent) : Boolean;
    {Charge le ParamSoc et appelle AddParamSoc}
    procedure AjouteParam  (Cmp : TComponent);
    {Ajoute un ParamSoc dans la FTobParamSoc}
    procedure AddParamSoc  (NomParam : string; TypeParam : Char; Value : Variant);
    {Modifie la valeur d'un ParamSoc dans la FTobParamSoc}
    procedure SetValeur    (NomParam : string; Value : Variant);
    {Récupère la valeur d'un ParamSoc dans la FTobParamSoc}
    function  GetValeur    (NomParam : string; Defaut : Variant; OldValue : Boolean = True) : Variant;
    {Récupère et stocke dans FTobParamSoc la valeur d'un composant}
    procedure MajValeur(Cmp : TComponent);
    {Remplit FTobParamSoc avec les valeurs de ParamSoc}
    procedure LitValeurs;
    {Ecrit dans ParamSoc les valeurs FTobParamSoc}
    procedure EcritValeurs;
    {Affecte un évènement sur un comosant pour des traitements spécifiques}
    procedure AffecteEvenement;
    {Récupération de composants}
    procedure GetControles;
    {Donne le focus à un ParamSoc en se positionnant éventuellement sur le bon onglet}
    procedure SetFocusControl(NomParam : string);
    {Pour brancher des traitements après le chargement des ParamSoc}
    procedure AfterRead;
    {Pour brancher des traitements avant l'écriture des ParamSoc}
    function BeforeWrite : Boolean;
  protected
    procedure PointageTresoClick(Sender : TObject);
    procedure OnEcranClosqeQuery(Sender : TObject; var CanClose : Boolean);
  public
    pcGeneral : TPageControl;
    {Si on change le mode de confidentialité}
    function VerifConfidentialite : Boolean;
    {Active désactive les ParamSoc au Chargement de la fiche}
    procedure SetEnabled;
    {Gestion de l'affichage}
    procedure GereAffichage;
  end;


function TRLanceFiche_ParamGeneraux(Arguments : string) : string;

implementation

uses
  {$IFDEF MODENT1}
  CPVersion,
  {$ENDIF MODENT1}
  Vierge, HCtrls, HEnt1, HMsgBox, HTB97, ParamSoc, ULibConfidentialite, Ent1,
  UtilPgi, Forms;

const
  tps_String  = 'S';
  tps_Boolean = 'B';
  tps_Integer = 'I';


{---------------------------------------------------------------------------------------}
function TRLanceFiche_ParamGeneraux(Arguments : string) : string;
{---------------------------------------------------------------------------------------}
begin
  Result := AglLanceFiche('TR', 'TRPARAMGENERAUX', '', '', Arguments);
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRPARAMGENERAUX.OnArgument(S : string);
{---------------------------------------------------------------------------------------}
begin
  inherited;
  Ecran.HelpContext := 50000158;


  FTobParamSoc := TOB.Create('A_PARAMSOC', nil, -1);
  FCanClose := True;
  {Récupèration des composants}
  GetControles;
  {Branchement des évènements}
  AffecteEvenement;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRPARAMGENERAUX.OnLoad;
{---------------------------------------------------------------------------------------}
begin
  inherited;
  {Chargement de la Tob}
  LitValeurs;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRPARAMGENERAUX.OnClose;
{---------------------------------------------------------------------------------------}
begin
  if Assigned(FTobParamSoc) then FreeAndNil(FTobParamSoc);
  inherited;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRPARAMGENERAUX.OnUpdate;
{---------------------------------------------------------------------------------------}
begin
  inherited;
  EcritValeurs;
end;

{Ecrit dans ParamSoc les valeurs FTobParamSoc
{---------------------------------------------------------------------------------------}
procedure TOF_TRPARAMGENERAUX.EcritValeurs;
{---------------------------------------------------------------------------------------}
var
  n : Integer;
  C : TComponent;
  F : TOB;
begin
  {Mise à jour de FTobParamSoc}
  for n := 0 to Ecran.ComponentCount - 1 do begin
    C := Ecran.Components[n];
    MajValeur(C);
  end;

  {Traitements éventuels de validation}
  if not BeforeWrite then begin
    PGIInfo(TraduireMemoire('Mise à jour abandonnée.'), Ecran.Caption);
    Exit;
  end;

  {Ecriture dans la base}
  for n := 0 to FTobParamSoc.Detail.Count - 1 do begin
    F := FTobParamSoc.Detail[n];
    if F.GetString('MODIFIE') = 'X' then begin
      if F.GetString('TYPEPS') = tps_Boolean then
        SetParamSoc(F.GetString('PARAMSOC'), F.GetValue('NEWVALUE') = 'X')
      else
        SetParamSoc(F.GetString('PARAMSOC'), F.GetValue('NEWVALUE'));
      F.PutValue('OLDVALUE', F.GetValue('NEWVALUE'));
      F.PutValue('MODIFIE', '-');
    end;
  end;

  PGIInfo(TraduireMemoire('La mise à jour s''est correctement effectuée.'), Ecran.Caption);
end;

{Remplit FTobParamSoc avec les valeurs de ParamSoc
{---------------------------------------------------------------------------------------}
procedure TOF_TRPARAMGENERAUX.LitValeurs;
{---------------------------------------------------------------------------------------}
var
  n : Integer;
  C : TComponent;
begin
  for n := 0 to Ecran.ComponentCount - 1 do begin
    C := Ecran.Components[n];
    if TesteComposant(C) then AjouteParam(C);
  end;
  {traitements à faire avant l'affichage}
  AfterRead;
end;

{Ajoute un ParamSoc dans la FTobParamSoc
{---------------------------------------------------------------------------------------}
procedure TOF_TRPARAMGENERAUX.AddParamSoc(NomParam : string; TypeParam : Char; Value : Variant);
{---------------------------------------------------------------------------------------}
var
  F : TOB;
begin
  F := TOB.Create('F_PARAMSOC', FTobParamSoc, -1);
  F.AddChampSupValeur('PARAMSOC', NomParam);
  F.AddChampSupValeur('OLDVALUE', Value);
  F.AddChampSupValeur('NEWVALUE', Value);
  F.AddChampSupValeur('MODIFIE' , '-');
  F.AddChampSupValeur('TYPEPS'  , TypeParam);
  try
         if TypeParam = tps_Boolean then SetControlChecked(NomParam, Value = 'X')
    else if TypeParam = tps_String  then SetControlText(NomParam, Value)
    else if TypeParam = tps_Integer then THNumEdit(GetControl(NomParam)).Value := Value;
  except
    PgiError(TraduireMemoire('Erreur lors du chargement du ParamSoc : ') + NomParam, Ecran.Caption);
  end;
end;

{Charge le ParamSoc et appelle AddParamSoc
{---------------------------------------------------------------------------------------}
procedure TOF_TRPARAMGENERAUX.AjouteParam(Cmp : TComponent);
{---------------------------------------------------------------------------------------}
var
  Typ : Char;
  Val : Variant;
begin
  {A compléter au fur et à mesure des besoins}
  if Cmp is THCheckBox then Typ := tps_Boolean
  else if (Cmp is THEdit) or (Cmp is THValComboBox) or (Cmp is THMultiValComboBox) then Typ := tps_String
  else if (Cmp is THNumEdit) then Typ := tps_Integer
  else Typ := tps_String;

  case Typ of
    tps_Boolean : if GetParamSocSecur(UpperCase(Cmp.Name), False) then Val := 'X'
                                                                  else Val := '-';
    tps_String  : Val := GetParamSocSecur(UpperCase(Cmp.Name), '');
    tps_Integer : Val := GetParamSocSecur(UpperCase(Cmp.Name), 0);
  else
    Val := GetParamSocSecur(UpperCase(Cmp.Name), '');
  end;
  AddParamSoc(UpperCase(Cmp.Name), Typ, Val);
end;

{Récupère et stocke dans FTobParamSoc la valeur d'un composant}
{---------------------------------------------------------------------------------------}
procedure TOF_TRPARAMGENERAUX.MajValeur(Cmp : TComponent);
{---------------------------------------------------------------------------------------}
begin
  if TesteComposant(Cmp) then begin
    if Cmp is THCheckBox then begin
      if (Cmp as THCheckBox).Checked then SetValeur(UpperCase(Cmp.Name), 'X')
                                     else SetValeur(UpperCase(Cmp.Name), '-');
    end
    else if (Cmp is THEdit) or (Cmp is THValComboBox) or (Cmp is THMultiValComboBox) then
      SetValeur(UpperCase(Cmp.Name), GetControlText(UpperCase(Cmp.Name)))
    else if (Cmp is THNumEdit) then
      SetValeur(UpperCase(Cmp.Name), (Cmp as THNumEdit).Value)
    else begin
      PgiError(TraduireMemoire('Impossible de mettre à jour le ParamSoc :') + UpperCase(Cmp.Name), Ecran.Caption);
      SetFocusControl(UpperCase(Cmp.Name));
    end;
  end;
end;

{Regarde si le composant correspond à un ParamSoc
{---------------------------------------------------------------------------------------}
function TOF_TRPARAMGENERAUX.TesteComposant(Cmp : TComponent) : Boolean;
{---------------------------------------------------------------------------------------}
begin
  {A compléter au fur et à mesure des besoins}
  Result := (Cmp is THEdit) or (Cmp is THCheckBox) or (Cmp is THValComboBox) or
            (Cmp is THMultiValComboBox) or (Cmp is THNumEdit);
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRPARAMGENERAUX.SetFocusControl(NomParam : string);
{---------------------------------------------------------------------------------------}
var
  Cmp : TComponent;
  Tbs : TTabSheet;
begin
  Cmp := Ecran.FindComponent(NomParam);
  if Assigned(Cmp) and (Cmp is TWinControl) then begin
    {Le composant peut prendre le focus, on lui donne ...}
    if TWinControl(Cmp).CanFocus then
      TWinControl(Cmp).SetFocus
    {... sinon, on s'assure que le composant est visible et actif ...}
    else if TWinControl(Cmp).Visible and TWinControl(Cmp).Enabled then begin
      {... alors, on recherche le bon Tabsheet}
      Tbs := nil;
      if TWinControl(Cmp).Parent is TTabSheet then
        Tbs := TTabSheet(TWinControl(Cmp).Parent)
         {Pour le cas où le ParamSoc serait dans un conteneur (GroupBox, Panel)}
      else if TWinControl(Cmp).Parent.Parent is TTabSheet then
        Tbs := TTabSheet(TWinControl(Cmp).Parent.Parent);
      if Assigned(Tbs) then pcGeneral.ActivePageIndex := Tbs.PageIndex;
    end;
  end;
end;

{---------------------------------------------------------------------------------------}
function TOF_TRPARAMGENERAUX.GetValeur(NomParam : string; Defaut : Variant; OldValue : Boolean = True) : Variant;
{---------------------------------------------------------------------------------------}
var
  F : TOB;
begin
  Result := Defaut;
  F := FTobParamSoc.FindFirst(['PARAMSOC'], [NomParam], True);
  if Assigned(F) then begin
    if OldValue then Result := F.GetValue('OLDVALUE')
                else Result := F.GetValue('NEWVALUE');
  end
  else if V_PGI.SAV then
    PGIInfo(TraduireMemoire('Paramètre société introuvable : ') + NomParam, Ecran.Caption);
end;

{Modifie la valeur d'un ParamSoc dans la FTobParamSoc
{---------------------------------------------------------------------------------------}
procedure TOF_TRPARAMGENERAUX.SetValeur(NomParam : string; Value : Variant);
{---------------------------------------------------------------------------------------}
var
  F : TOB;
begin
  F := FTobParamSoc.FindFirst(['PARAMSOC'], [NomParam], True);
  if Assigned(F) then begin
    if F.GetValue('NEWVALUE') <> Value then begin
      F.PutValue('NEWVALUE', Value);
      F.PutValue('MODIFIE' , 'X');
    end;
  end
  else
    PGIError(TraduireMemoire('Paramètre société introuvable : ') + NomParam, Ecran.Caption);
end;

{Affecte un évènement sur un comosant pour des traitements spécifiques
{---------------------------------------------------------------------------------------}
procedure TOF_TRPARAMGENERAUX.AffecteEvenement;
{---------------------------------------------------------------------------------------}
begin
  TFVierge(Ecran).OnCloseQuery := OnEcranClosqeQuery;
  (GetControl('SO_TRPOINTAGETRESO') as THCheckBox).OnClick := PointageTresoClick;
end;

{Récupération de composants
{---------------------------------------------------------------------------------------}
procedure TOF_TRPARAMGENERAUX.GetControles;
{---------------------------------------------------------------------------------------}
begin
  pcGeneral := TPageControl(GetControl('PCGENERAL'));
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRPARAMGENERAUX.OnEcranClosqeQuery(Sender : TObject; var CanClose : Boolean);
{---------------------------------------------------------------------------------------}
begin
  CanClose  := FCanClose;
  FCanClose := True;
  TFVierge(Ecran).FormCloseQuery(Sender, CanClose);
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRPARAMGENERAUX.PointageTresoClick(Sender : TObject);
{---------------------------------------------------------------------------------------}
begin
  if (GetValeur('SO_TRPOINTAGETRESO', '-') = '-') and (Sender as TCheckBox).Checked then begin
    if PGIAsk(TraduireMemoire('Vous êtez sur le point de changer le mode de pointage.') + #13 +
              TraduireMemoire('Il est recommander d''avoir un pointage à jour') + #13#13 +
              TraduireMemoire('Confirmez-vous votre choix ?'), Ecran.Caption) = mrNo then
      (Sender as TCheckBox).Checked := False;
  end
  else if (GetValeur('SO_TRPOINTAGETRESO', '-') = 'X') and not (Sender as TCheckBox).Checked then begin
    if PGIAsk(TraduireMemoire('Il est dangereux de changer le mode de pointage.') + #13#13 +
              TraduireMemoire('Souhaitez-vous annuler la modification ?'), Ecran.Caption) = mrYes then
      (Sender as TCheckBox).Checked := True;
  end;
end;

{Pour brancher des traitements après le chargement des ParamSoc
{---------------------------------------------------------------------------------------}
procedure TOF_TRPARAMGENERAUX.AfterRead;
{---------------------------------------------------------------------------------------}
begin
  {Rend visible au invisible des composants}
  GereAffichage;
  {Désactive/Active les paramSoc}
  SetEnabled;
  pcGeneral.ActivePage := TTabSheet(GetControl('TSIDENTITE'));
  Application.ProcessMessages;
  {C'est de la bidouille mais si la fiche n'a pas été enregistrée sur le TSIDENTITE, il y a un
   problème de rafraîchissement}
  pcGeneral.ActivePageIndex := 1;
  pcGeneral.ActivePageIndex := 0;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRPARAMGENERAUX.GereAffichage;
{---------------------------------------------------------------------------------------}
begin
  {Pour le moment la confidentialité n'est visible qu'en Tréso}
  SetControlVisible(NOMCONFIDENTIEL , (ctxTreso in V_PGI.PGIContexte));
  SetControlVisible('LBETEBAC'      , not (CtxPcl in V_PGI.PGIContexte));
  SetControlVisible('SO_REPETEBAC3' , not (CtxPcl in V_PGI.PGIContexte));
  SetControlVisible('SO_POINTAGEJAL', not (ctxTreso in V_PGI.PGIContexte));
  TTabSheet(GetControl('TSTRESO')).Visible := EstComptaTreso;
end;

{Active désactive les ParamSoc au Chargement de la fiche
{---------------------------------------------------------------------------------------}
procedure TOF_TRPARAMGENERAUX.SetEnabled;
{---------------------------------------------------------------------------------------}
begin
  {Le mode de pointage n'est actif que si l'on n'est pas en Tréso et si la Table EEXBQ est vide}
  SetControlEnabled('SO_POINTAGEJAL', not ExisteSQL('SELECT EE_GENERAL FROM EEXBQ'));
  {Les paramêtre généraux de la société sont désactivés en PCL}
  SetControlEnabled('SO_SOCIETE'   , not (CtxPcl in V_PGI.PGIContexte));
  SetControlEnabled('SO_LIBELLE'   , not (CtxPcl in V_PGI.PGIContexte));
  SetControlEnabled('SO_ADRESSE1'  , not (CtxPcl in V_PGI.PGIContexte));
  SetControlEnabled('SO_ADRESSE2'  , not (CtxPcl in V_PGI.PGIContexte));
  SetControlEnabled('SO_ADRESSE3'  , not (CtxPcl in V_PGI.PGIContexte));
  SetControlEnabled('SO_CODEPOSTAL', not (CtxPcl in V_PGI.PGIContexte));
  SetControlEnabled('SO_VILLE'     , not (CtxPcl in V_PGI.PGIContexte));
  SetControlEnabled('SO_PAYS'      , not (CtxPcl in V_PGI.PGIContexte));
  SetControlEnabled('SO_TELEPHONE' , not (CtxPcl in V_PGI.PGIContexte));
  SetControlEnabled('SO_FAX'       , not (CtxPcl in V_PGI.PGIContexte));
  SetControlEnabled('SO_MAIL'      , not (CtxPcl in V_PGI.PGIContexte));
  SetControlEnabled('SO_SIRET'     , not (CtxPcl in V_PGI.PGIContexte));
  SetControlEnabled('SO_RC'        , not (CtxPcl in V_PGI.PGIContexte));
  SetControlEnabled('SO_CAPITAL'   , not (CtxPcl in V_PGI.PGIContexte));
  SetControlEnabled('SO_APE'       , not (CtxPcl in V_PGI.PGIContexte));
  {On ne peut modifier le dossier de Tréso que s'il est vide
   09/10/07 : Avec du menu Cash Polling, il ne faut plus modifier le paramsoc à la main}
  SetControlEnabled('SO_TRBASETRESO' , False);
  SetControlEnabled('LSO_TRBASETRESO', False);
end;

{Pour brancher des traitements avant l'écriture des ParamSoc
{---------------------------------------------------------------------------------------}
function TOF_TRPARAMGENERAUX.BeforeWrite : Boolean;
{---------------------------------------------------------------------------------------}
begin
  Result := True;
  Result := Result and VerifConfidentialite;

  FCanClose := Result;
end;

{Si on change le mode de confidentialité
{---------------------------------------------------------------------------------------}
function TOF_TRPARAMGENERAUX.VerifConfidentialite : Boolean;
{---------------------------------------------------------------------------------------}
var
  F : TOB;
begin
  Result := True;
  F := FTobParamSoc.FindFirst(['PARAMSOC'], [NOMCONFIDENTIEL], True);
  if Assigned(F) then begin
    if F.GetValue('NEWVALUE') <> F.GetValue('OLDVALUE') then begin
      case PGIAskCancel(TraduireMemoire('Attention : vous venez de modifier la gestion des confidentialités.') + #13 +
                         TraduireMemoire('Tout le paramétrage des confidentialités va être supprimé.') + #13#13 +
                         TraduireMemoire('Confirmez-vous votre choix ?'), Ecran.Caption) of
        {On empêche la validation}
        mrCancel : Result := False;
        {On annule la modification}
        mrNo     : begin
                     TCheckBox(GetControl(NOMCONFIDENTIEL)).Checked := False;
                     F.PutValue('NEWVALUE', F.GetValue('OLDVALUE'));
                     F.PutValue('MODIFIE', '-');
                    end;
        {On vide la table PROSPECTCONF}
        mrYes    : Result := VideConfidentialites;
      end;
    end;
  end;
  if not Result then SetFocusControl(NOMCONFIDENTIEL);
end;

initialization
  RegisterClasses([TOF_TRPARAMGENERAUX]);


end.
