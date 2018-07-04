{-------------------------------------------------------------------------------------
    Version    |  Date  | Qui | Commentaires
--------------------------------------------------------------------------------------
 6.0X.xxx.xxx   02/11/04  JP   Création de l'unité
 6.20.001.003   10/02/05  JP   FQ 10200 : application de la TVA aux commissions en plus des Frais
 7.09.001.00    04/10/06  JP   Gestion du multi sociétés
 7.09.001.001   23/10/06  JP   Gestion des filtres Multi sociétés
 7.09.001.001   06/11/06  JP   FQ 10380 : Reprise des portefeuilles existants
 7.09.001.006   19/02/07  JP   Ajout de précautions sur sur la FQ 10380
 8.00.001.001   15/12/06  JP   FQ 10373 : Les nombres de parts passe d'un Integer à un Double
 8.00.001.018   06/06/07  JP   FQ 10422 : accélération de la saisie
 8.00.001.019   10/06/07  JP   FQ 10409 : accès au ventes
 8.00.001.024   10/07/07  JP   FQ 10501 : Corrections de l'initialisation du NumOpcvm en 2/3 si on est entré
                               dans la fiche en modification et que l'on enchaîne sur une création
--------------------------------------------------------------------------------------}
unit TROPCVM_TOM;

interface

uses {$IFDEF VER150} variants,{$ENDIF}
  {$IFDEF EAGLCLIENT}
  MaineAGL, eFiche,
  {$ELSE}
   db, {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF} HDB, FE_Main, Fiche,
  {$ENDIF}
  StdCtrls, Controls, Classes, SysUtils, ComCtrls, HCtrls, HEnt1, UTOM, UTob, Menus,
  Constantes, UObjOPCVM;

type
  TOM_TROPCVM = class(TOM)
    procedure OnNewRecord              ; override;
    procedure OnUpdateRecord           ; override;
    procedure OnAfterUpdateRecord      ; override;
    procedure OnArgument   (S : string); override;
    procedure OnClose                  ; override;
    procedure OnDeleteRecord           ; override;
    procedure OnChangeField(F : TField); override;
    procedure OnLoadRecord             ; override;
    procedure OnCancelRecord           ; override;
  private
    FCodTransac  : string;
    FNumTransac  : string;
    FCompteur    : string;
    FCotation    : Double;
    ObjEnvers    : TObjOPCVMCalcul;
    ObjVente     : TObjOPCVMVente;

    procedure SetNumTransac(Value : string);
    procedure SetCodTransac(Value : string);
    procedure SetCotation  (Value : Double);
    function  GetCotation  : Double;
    function  IsNumActif   : Boolean;
    function  IsDSModif    : Boolean;
    procedure LancerCalcul (Sender, ZoneDev : string);
    procedure InitEcr      (var tEcr : TOB);
    procedure MajZones     ;
    procedure GereMenuPop  ;
    procedure SetInfosDossier; {04/10/06 : Gestion du NoDossier et du code société}

  protected
    oDevise    : string;  {Ancienne devise}
    nDevise    : string;  {Nouvelle devise}
    VenteOk    : Boolean; {A True si on est vente, à False si on est en achat}
    CalVenteOk : Boolean; {A True si les ventes demandent à être recalculées}
    nbAVendre  : Double;  {15/12/06; FQ 10373 : Les nombres de parts passe d'un Integer à un Double}

    procedure LancerFrais    (Sender : TObject);
    procedure LancerCours    (Sender : TObject);
    procedure LancerDevise   (Sender : TObject);
    procedure LancerVentes   (Sender : TObject); {10/06/07 : FQ 10409}
    procedure GetCours       (Sender : TObject);
    procedure OpcvmChange    (Sender : TObject);
    procedure PortefChange   (Sender : TObject);
    procedure NbVenteChange  (Sender : TObject);
    procedure DateVenteChange(Sender : TObject);
    procedure CkRepriseChange(Sender : TObject); {19/02/07}
    {Calcul d'une transaction d'achat}
    procedure CalcAchatClick (Sender : TObject);
    {Calcul d'une transaction de vente}
    procedure CalcVenteClick (Sender : TObject);
    {Calcul d'une transaction d'achat en devise pivot à l'envers}
    procedure CalcEnvEurClick(Sender : TObject);
    {Calcul d'une transaction d'achat en devise à l'envers}
    procedure CalcEnvDevClick(Sender : TObject);
    procedure OnAfterShow;
    procedure GererControls;
    procedure GererAffichage;
    procedure MajTotauxAchat;
    procedure MajTotauxVente;

    {Routines de vente partielle}
    function  IsNbPartsOk : Boolean; {Vérifie si le nombre de parts est cohérent}
    procedure CalculVente ; {Calcul les montants de la vente}
  public
    pcControl : TPageControl;
    ppMenu    : TPopupMenu;

    procedure MettreLeFocus(ControlName : string; NumPage : Integer);
    procedure GererLesCours(Chp : string);
    procedure AfficherCours(dt : TDateTime; Opcvm : string);
    procedure ChargerDrapeaux;

    property Cotation   : Double  read GetCotation write SetCotation;
    property NumTransac : string  read FNumTransac write SetNumTransac;
    property CodTransac : string  read FCodTransac write SetCodTransac;
    property NumActif   : Boolean read IsNumActif;
  end ;

procedure TRLanceFiche_OPCVM(Dom, Fiche, Range, Lequel, Arguments : string);

implementation

uses
  ExtCtrls, Commun, HMsgBox, HTB97, TRMULFRAIS_TOF, TRCOTATIONOPCVM_TOF, Windows, Messages,
  TRCOTATIONOPCVM_TOM, CPCHANCELL_TOF, Forms, UProcGen, Buttons, UProcCommission, AglInit,
  ParamSoc, UObjEtats;

{---------------------------------------------------------------------------------------}
procedure TRLanceFiche_OPCVM(Dom, Fiche, Range, Lequel, Arguments : string);
{---------------------------------------------------------------------------------------}
var
  S : string;
  c : string;
  O : Boolean;
  Q : TQuery;
begin
  S := Arguments;
  {Suppression de l'action}
  C := ReadTokenSt(S);

  {On regarde si on l'on veut générer une vente}
  O := GetParamSocSecur('SO_TRFIFO', True) and (ReadTokenSt(S) = 'VEN') and (Pos('MODIF', UpperCase(C)) > 0);
  {Si on est en vente, il faut que l'on soit sur la plus vieille transaction disponible}
  if O then begin
    Q := OpenSQL('SELECT TOP_CODEOPCVM, TOP_GENERAL, TOP_DATEACHAT FROM TROPCVM WHERE TOP_NUMOPCVM = "' + LEQUEL + '"', True);
    if not Q.Eof then begin
      S := 'SELECT TOP_CODEOPCVM FROM TROPCVM WHERE (TOP_NBPARTACHETE - TOP_NBPARTVENDU) > 0 AND ' +
           'TOP_GENERAL = "'   + Q.FindField('TOP_GENERAL').AsString   + '" AND ' +
           'TOP_CODEOPCVM = "' + Q.FindField('TOP_CODEOPCVM').AsString + '" AND ' +
           'TOP_DATEACHAT < "' + UsDateTime(Q.FindField('TOP_DATEACHAT').AsDateTime) + '"';
      O := ExisteSQL(S);
    end
    else
      O := True;
    {Il y a des OPCVM plus ancienne, on passe en lecture seule}
    if O then begin
      HShowMessage('0;Vente d''OPCVM;Il existe une transaction antérieure à vendre pour cet OPCVM pour ce compte.'#13 +
                   'Il n''est pas possible de vendre cette transaction;W;O;O;O;', '', '');
      Arguments := ActionToString(taConsult) + ';VEN;';
    end;
    Ferme(Q);
  end;

  AglLanceFiche(Dom, Fiche, Range, Lequel, Arguments);
end;

{---------------------------------------------------------------------------------------}
procedure TOM_TROPCVM.OnArgument(S : string);
{---------------------------------------------------------------------------------------}
begin
  inherited;
  Ecran.HelpContext := 150;
  {Suppression de l'action}
  ReadTokenSt(S);
  {On regarde si on est en duplication}
  VenteOk := ReadTokenSt(S) = 'VEN';
  GererControls;
  GererAffichage;
  ObjEnvers := TObjOPCVMCalcul.Create;
  ObjVente := TObjOPCVMVente.Create(vop_Partiel);
  {23/10/06 : Gestion des filtres multi sociétés sur banquecp et dossier}
  {$IFDEF EAGLCLIENT}
  THEdit(GetControl('TOP_GENERAL'  )).Plus := FiltreBanqueCp(THEdit(GetControl('TOP_GENERAL')).DataType, '', '');
  {$ELSE}
  THDBEdit(GetControl('TOP_GENERAL')).Plus := FiltreBanqueCp(THDBEdit(GetControl('TOP_GENERAL')).DataType, '', '');
  {$ENDIF EAGLCLIENT}
end;

{---------------------------------------------------------------------------------------}
procedure TOM_TROPCVM.OnClose;
{---------------------------------------------------------------------------------------}
begin
  if Assigned(ObjEnvers) then FreeAndNil(ObjEnvers);
  if Assigned(ObjVente)  then FreeAndNil(ObjVente);
  inherited;
end;

{---------------------------------------------------------------------------------------}
procedure TOM_TROPCVM.OnNewRecord;
{---------------------------------------------------------------------------------------}
begin
  inherited;
  SetField('TOP_SOCIETE', V_PGI.CodeSociete);
  SetField('TOP_NODOSSIER', V_PGI.NoDossier); {04/10/06}
  SetField('TOP_USERCREATION', V_PGI.User);
  SetField('TOP_COMPTA', '-');
  SetField('TOP_VALBO', '-');
  SetField('TOP_STATUT', '-');
  {06/06/07 : FQ 10422 : Initialisation de la date d'achat}
  SetField('TOP_DATEACHAT', V_PGI.DateEntree);
  {Nécessaire car ce champ constitue l'index et en création il est réactivé par l'AGL}
  SetControlEnabled('TOP_NUMOPCVM', False);
end;

{---------------------------------------------------------------------------------------}
procedure TOM_TROPCVM.OnDeleteRecord;
{---------------------------------------------------------------------------------------}
begin
  inherited;
end;

{---------------------------------------------------------------------------------------}
procedure TOM_TROPCVM.OnloadRecord;
{---------------------------------------------------------------------------------------}
begin
  inherited;
  {Si on n'est pas en insertion, on interdit de modifier la transsaction car elle influe
   sur TOP_NUMOPCVM qui constitue la clef primaire}
  SetControlEnabled('TOP_TRANSACTION', NumActif);
  FCompteur := '';
  MajTotauxAchat;
  MajTotauxVente;

  MettreLeFocus('TOP_PORTEFEUILLE', 0); {16/05/06 : FQ 10352}
end;

{---------------------------------------------------------------------------------------}
procedure TOM_TROPCVM.OnChangeField(F : TField);
{---------------------------------------------------------------------------------------}
var
  ch : string;
begin
  inherited;
  if F.FieldName = 'TOP_NUMOPCVM' then
    NumTransac := F.AsString
  else if F.FieldName = 'TOP_TRANSACTION' then begin
    CodTransac := F.AsString;
    {$IFDEF EAGLCLIENT}
    {06/06/07 : FQ 10422 : Accélérateur de saisie}
    SetField('TOP_LIBELLE', THValComboBox(GetControl('TOP_TRANSACTION')).Text);
    {$ELSE}
    SetField('TOP_LIBELLE', THDBValComboBox(GetControl('TOP_TRANSACTION')).Text);
    {$ENDIF EAGLCLIENT}
  end
  else if F.FieldName = 'TOP_GENERAL' then
    SetInfosDossier
  else if F.FieldName = 'TOP_DATEACHAT' then
    GererLesCours('TOP_DATEACHAT')
  else if (F.FieldName = 'TOP_LIBELLE') then begin
    ch := F.AsString;
    {GetField semble changer le TField, d'où quelques précautions. J'ai fait une fiche qualité}
    if (VarToStr(GetField('TOP_REFERENCE')) = '') then
    SetField('TOP_REFERENCE', ch)
  end
  else if F.FieldName = 'TOP_CODEOPCVM' then
    GererLesCours('TOP_CODEOPCVM')
  else if F.FieldName = 'TOP_COTATIONACH' then begin
    if not IsDSModif then
      Cotation := F.AsFloat;
  end
  else if F.FieldName = 'TOP_NBPARTVENDU' then begin
    {10/06/07 : FQ 10409 : On n'affiche l'accès à la vente que si on n'est pas en insertion !}
    ppMenu.Items[3].Visible := not (DS.State = dsInsert) and (F.AsFloat > 0); {Menu séparateur}
    ppMenu.Items[4].Visible := not (DS.State = dsInsert) and (F.AsFloat > 0);
  end
end;

{---------------------------------------------------------------------------------------}
procedure TOM_TROPCVM.OnUpdateRecord;
{---------------------------------------------------------------------------------------}
var
  mResult : TModalResult;
begin
  inherited;
  {Si les ventes n'ont pas été calculées et demandent à l'être}
  if VenteOk and CalVenteOk then begin
    {26/06/07 : FQ 10488 : mauvaise gestion du HShowMessage que je remplace par PGIAskCancel}
    mResult := PGIAskCancel(TraduireMemoire('La vente demande a être recalculée. Voulez-vous relancer le calcul ?'), Ecran.Caption);
    if mResult in [mrYes, mrCancel] then begin
      if mResult = mrYes then
        CalcVenteClick(GetControl('BVENTE'));
      LastError := 2;
    end
    else
      CalVenteOk := False;
  end;

  if GetField('TOP_GENERAL') = '' then begin
    LastError := 2;
    LastErrorMsg := TraduireMemoire('Veuillez renseigner le compte bancaire.');
    MettreLeFocus('TOP_GENERAL', 0);
  end
  else if GetField('TOP_LIBELLE') = '' then begin
    LastError := 2;
    LastErrorMsg := TraduireMemoire('Veuillez renseigner le libellé de la transaction.');
    MettreLeFocus('TOP_LIBELLE', 0);
  end
  else if GetField('TOP_CODEOPCVM') = '' then begin
    LastError := 2;
    LastErrorMsg := TraduireMemoire('Veuillez renseigner l''OPCVM.');
    MettreLeFocus('TOP_CODEOPCVM', 0);
  end
  else if GetField('TOP_PORTEFEUILLE') = '' then begin
    LastError := 2;
    LastErrorMsg := TraduireMemoire('Veuillez renseigner le portefeuille.');
    MettreLeFocus('TOP_PORTEFEUILLE', 0);
  end
  else if GetField('TOP_TRANSACTION') = '' then begin
    LastError := 2;
    LastErrorMsg := TraduireMemoire('Veuillez renseigner la transaction.');
    MettreLeFocus('TOP_TRANSACTION', 0);
  end
  else if GetField('TOP_BASECALCUL') = '' then begin
    LastError := 2;
    LastErrorMsg := TraduireMemoire('Veuillez renseigner la base de calcul.');
    MettreLeFocus('TOP_BASECALCUL', 1);
  end
  {15/12/06 : FQ 10373 : On passe d'un Integer à un Double}
  else if Valeur(GetField('TOP_NBPARTACHETE')) <= 0 then begin
    LastError := 2;
    LastErrorMsg := TraduireMemoire('Veuillez renseigner le nombre de parts achetées.');
    MettreLeFocus('', 1);
  end
  else if VarToDateTime(GetField('TOP_DATEACHAT')) <= 2 then begin
    LastError := 2;
    LastErrorMsg := TraduireMemoire('Veuillez renseigner la date d''achat.');
    MettreLeFocus('', 1);
  end;

  if LastError = 0 then begin
    if VenteOk then begin
      {Si on est en vente partiel}
      if ObjVente.EnAttente then begin
        BeginTrans;
        try
          ObjVente.PutVenteInBase;
          CommitTrans;
          {On ne le repasse à False que si tout s'est bien passé !!!}
          ObjVente.EnAttente := False;
        except
          on E : Exception do begin
            RollBack;
            LastError := 2;
            LastErrorMsg := E.Message;
          end;
        end;
      end;
    end
    else begin
      {06/11/06 : FQ 10380 : Gestion de la reprise des OPCVM}
      if (GetCheckBoxState('CKREPRISE') = cbChecked) then begin
        mResult := HShowMessage('0;' + Ecran.Caption + ';Vous avez désigner cette OPCVM comme une reprise de portefeuille.'#13 +
                        'Les flux de Trésorerie et de Comptabilité liés à l''achat ne seront pas générés.'#13#13 +
                        'Attention ! En cochant cette case, il ne sera plus possible de modifier cette saisie.'#13#13 +
                        'Confirmez-vous ce choix ?;W;YNC;N;C;', '', '');
        {19/02/07 : Ajout d'un second message de confirmation}
        if mResult = mrYes then begin
          mResult := HShowMessage('0;' + Ecran.Caption + ';Considérer une OPCVM comme une reprise de portefeuille est irréversible.'#13 +
                                   'Souhaitez-vous revenir en mode standard ?;W;YNC;Y;C;', '', '');
          if mResult = mrYes then mResult := mrNo
          else if mResult = mrNo then mResult := mrYes;
        end;
        if mResult = mrYes then begin
          {Lors d'une reprise, il ne faut pas générer de flux de Trésorerie ni d'écritures comptables :
           On met donc à jour les champs de validation BO et d'intégration en compta}
          SetField('TOP_VALBO'     , 'X');
          SetField('TOP_COMPTA'    , 'X');
          SetField('TOP_USERVBO'   , CODEREPRISEVM);
          SetField('TOP_USERCOMPTA', CODEREPRISEVM);
          SetField('TOP_DATEVBO'   , V_PGI.DateEntree);
          SetField('TOP_DATECOMPTA', V_PGI.DateEntree);
        end
        else if mResult = mrCancel then begin
          LastError := 2;
          SetControlChecked('CKREPRISE', False);
          SetFocusControl('CKREPRISE');
        end;
      end;
    end;{Else if venteok}
  end; {Last error = 0}
end;

{---------------------------------------------------------------------------------------}
procedure TOM_TROPCVM.OnAfterUpdateRecord;
{---------------------------------------------------------------------------------------}
begin
  inherited;
  {Si on est en achat}
  if not VenteOk then begin
    {Mise à jour du compteur de transactions}
    if FCompteur <> '' then
      SetNum(CODEMODULE, V_PGI.CodeSociete, CodTransac, FCompteur);
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOM_TROPCVM.MajTotauxAchat;
{---------------------------------------------------------------------------------------}
var
  Tot    : Double;
  TotEur : Double;
begin
  {Calcul des totaux d'achat}
  Tot := VarToDle(GetField('TOP_FRAISACHDEV')) + VarToDle(GetField('TOP_TVAACHATDEV')) +
         VarToDle(GetField('TOP_COMACHATDEV')) + VarToDle(GetField('TOP_MONTANTACHDEV'));
  TotEur := VarToDle(GetField('TOP_FRAISACH')) + VarToDle(GetField('TOP_TVAACHAT')) +
            VarToDle(GetField('TOP_COMACHAT')) + VarToDle(GetField('TOP_MONTANTACH'));
  SetControlText('EDTOTALDEVACH', FloatToStr(Tot));
  SetControlText('EDTOTALEURACH', FloatToStr(TotEur));
end;

{---------------------------------------------------------------------------------------}
procedure TOM_TROPCVM.MajTotauxVente;
{---------------------------------------------------------------------------------------}
var
  Tot : Double;
  TotEur : Double;
begin
  {Calcul des totaux de vente}
  Tot := - VarToDle(GetField('TOP_FRAISVENDEV')) - VarToDle(GetField('TOP_TVAVENTEDEV')) +
         - VarToDle(GetField('TOP_COMVENTEDEV')) + VarToDle(GetField('TOP_MONTANTVENDEV'));
  TotEur := - VarToDle(GetField('TOP_FRAISVEN')) - VarToDle(GetField('TOP_TVAVENTE')) +
            - VarToDle(GetField('TOP_COMVENTE')) + VarToDle(GetField('TOP_MONTANTVEN'));
  SetControlText('EDTOTALDEVVEN', FloatToStr(Tot));
  SetControlText('EDTOTALEURVEN', FloatToStr(TotEur));
  {Calcul du prox moyen de vente
   15/12/06 : FQ 10373 : On passe d'un Integer à un Double}
  if VarToDle(GetField('TOP_NBPARTVENDU')) <> 0 then begin
    Tot := VarToDle(GetField('TOP_MONTANTVENDEV')) / VarToDle(GetField('TOP_NBPARTVENDU'));
    SetControlText('EDMOYENNE', FloatToStr(Tot));
  end;
end;

{Appel du mul des frais
{---------------------------------------------------------------------------------------}
procedure TOM_TROPCVM.LancerFrais(Sender : TObject);
{---------------------------------------------------------------------------------------}
begin
  TRLanceFiche_MulFrais('TR', 'TRMULFRAIS', '', '' , '');
end;

{Appel du mul des cours des OPCVM
{---------------------------------------------------------------------------------------}
procedure TOM_TROPCVM.LancerCours(Sender : TObject);
{---------------------------------------------------------------------------------------}
begin
  TRLanceFiche_TRCotationOPCVM('TR', 'TRCOTATIONOPCVM', '', '', GetField('TOP_CODEOPCVM'));
end;

{Appel de la fiche de chancellerie
{---------------------------------------------------------------------------------------}
procedure TOM_TROPCVM.LancerDevise(Sender : TObject);
{---------------------------------------------------------------------------------------}
begin
  if VenteOk then begin
    if GetField('TOP_STATUT') = 'X' then
      FicheChancel(GetControlText('DEV1'), True, StrToDate(GetControlText('TOP_DATEFIN')), taModif, True)
    else
      FicheChancel(GetControlText('DEV1'), True, V_PGI.DateEntree, taModif, True);
  end
  else
    FicheChancel(GetControlText('DEV1'), True, StrToDate(GetControlText('TOP_DATEACHAT')), taModif, True);
end;

{Lancement de la fiche des cours des OPCVM
{---------------------------------------------------------------------------------------}
procedure TOM_TROPCVM.GetCours(Sender : TObject);
{---------------------------------------------------------------------------------------}
var
  Mnt : string;
  Dt  : TDateTime;
begin
  {le GetField n'est pas forcément à jour, car les speedbutons ne prennent pas le focus,
   on travaille dont avec le GetControlText()}
  dt := StrToDateTime(GetControlText('TOP_DATEACHAT'));
  if dt = 0 then begin
    HShowMessage('1;' + Ecran.Caption + ';Veuillez saisir une date d''achat;W;O;O;O;', '', '');
    Exit;
  end;

  {Appel de la fiche de saisie des cours des OPCVM}
  Mnt := GetField('TOP_CODEOPCVM') + ';' + DateToStr(Dt) + ';';
  Mnt := TRLanceFiche_TRCoursOPCVM('TR', 'TRCOURSOPCVM', '', Mnt, ';DIRECT;' + Mnt);

  {S'il y a un retour, on affecte le champ}
  if Mnt <> '' then
    SetField('TOP_PRIXUNITAIRE', FormateMontant(Valeur(Mnt), 5));
end;

{Recherche du cours le plus récent pour l'OPCVM courante
{---------------------------------------------------------------------------------------}
procedure TOM_TROPCVM.AfficherCours(dt : TDateTime; Opcvm : string);
{---------------------------------------------------------------------------------------}
var
  Q : TQuery;
begin
  Q := OpenSQL('SELECT TTO_COTATION FROM TRCOTATIONOPCVM WHERE TTO_CODEOPCVM = "' + Opcvm +
               '" AND TTO_DATE BETWEEN "' + UsDateTime(DebutAnnee(dt - 5)) + '" ' +
               ' AND "' + UsDateTime(dt) + '" ORDER BY TTO_DATE DESC', True);
  {Mise à jour du prix unitaire}
  if not Q.EOF then
    SetField('TOP_PRIXUNITAIRE', FormateMontant(Q.FindField('TTO_COTATION').AsFloat, 5));
  Ferme(Q);
end;

{Gestion des cours des devises et des OPCVM
{---------------------------------------------------------------------------------------}
procedure TOM_TROPCVM.GererLesCours(Chp : string);
{---------------------------------------------------------------------------------------}
var
  d : TDateTime;
  c : string;
begin
  {Si on n'est pas en modification ou en insertion}
  if not IsDSModif then Exit;

  c := VarToStr(GetField('TOP_CODEOPCVM'));
  d := VarToDateTime(GetField('TOP_DATEACHAT'));

  {Si la date est inférieure au 01/01/1900 ou si le code OPCVM }
  if (d <= iDate1900) or (c = '') then begin
    Cotation := 1.0000001;
    SetField('TOP_PRIXUNITAIRE', 0);
    Exit;
  end;

  if Chp = 'TOP_DATEACHAT' then begin
    AfficherCours(d, c);
    Cotation := RetPariteEuro(nDevise, d, True);
  end;

  if Chp = 'TOP_CODEOPCVM' then begin
    AfficherCours(d, c);
    if oDevise <> nDevise then
      Cotation := RetPariteEuro(nDevise, d, True);
  end;
end;

{Appel du mul des cours des OPCVM
{---------------------------------------------------------------------------------------}
procedure TOM_TROPCVM.OpcvmChange(Sender : TObject);
{---------------------------------------------------------------------------------------}
begin
  {Gestion de la devise}
  ChargerDrapeaux;
end;

{On filtre les OPCVM en fonction du portefeuille
{---------------------------------------------------------------------------------------}
procedure TOM_TROPCVM.PortefChange(Sender : TObject);
{---------------------------------------------------------------------------------------}
begin
  {$IFDEF EAGLCLIENT}
  THValComboBox(GetControl('TOP_CODEOPCVM')).Plus := 'TOF_PORTEFEUILLE = "' +
                                                     GetControlText('TOP_PORTEFEUILLE') + '"';
  {$ELSE}
  THDBValComboBox(GetControl('TOP_CODEOPCVM')).Plus := 'TOF_PORTEFEUILLE = "' +
                                                       GetControlText('TOP_PORTEFEUILLE') + '"';
  {$ENDIF}
end;

{---------------------------------------------------------------------------------------}
procedure TOM_TROPCVM.NbVenteChange(Sender : TObject);
{---------------------------------------------------------------------------------------}
begin
  CalVenteOk := True;
end;

{---------------------------------------------------------------------------------------}
procedure TOM_TROPCVM.DateVenteChange(Sender : TObject);
{---------------------------------------------------------------------------------------}
begin
  CalVenteOk := True;
  if StrToDate(GetControlText('EDDATEVENTE')) < VarToDateTime(GetField('TOP_DATEACHAT')) then begin
    HShowMessage('0;' + Ecran.Caption + ';La Date de vente ne peut être antérieure à la date d''achat.;W;O;O;O;', '', '');
    SetControlText('EDDATEVENTE', GetControlText('TOP_DATEACHAT'));
    SetFocusControl('EDDATEVENTE');
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOM_TROPCVM.ChargerDrapeaux;
{---------------------------------------------------------------------------------------}
var
  Dev : string;
  n   : Byte;
  Vis : Boolean;
  lb  : THLabel;
begin
  Dev := GetControlText('TOP_CODEOPCVM');
  {Par défaut on met la devise pivot}
  if (Dev = '') or (Dev = #0 ) then Dev := V_PGI.DevisePivot
                               else Dev := RetDeviseOPCVM(Dev);

  oDevise := nDevise;
  nDevise := Dev;

  {Chargement du drapeau relatif à la devise de l'OPCVM en cours}
  AssignDrapeau(Timage(GetControl('IDEV')), Dev);
  for n := 1 to 3 do begin
    Timage(GetControl('IDEV' + IntToStr(n))).Picture.Assign(Timage(GetControl('IDEV')).Picture.Graphic);
    SetControlCaption('DEV' + IntToStr(n), Dev);
  end;

  {Gestion des libellés et des zones affichant les contreparties en euros
   si l'opcvm est en devises}
  Vis := not (UpperCase(Dev) = UpperCase(V_PGI.DevisePivot));
  SetControlVisible('BTOTEUR', Vis and not ((TFFiche(Ecran).FTypeAction = taConsult) or VenteOk));
  SetControlVisible('LBEUR'  , Vis and ((TFFiche(Ecran).FTypeAction = taConsult) or VenteOk));
  SetControlVisible('LBCOURS', Vis);
  SetControlVisible('DEVEUR' , Vis);
  SetControlVisible('IDEVEUR', Vis);
  SetControlVisible('IDEV4'  , Vis); {20/05/05 : Oubli de ma part}
  SetControlVisible('EDTOTALEURACH', Vis);
  SetControlVisible('EDTOTALEURVEN', Vis);
  SetControlVisible('TOP_COTATIONACH', Vis);
  SetControlVisible('TOP_COTATIONVEN', Vis);
  {Gestion de l'affichage en devise pivot si l'OPCVM est en devise}
  if Vis then begin
    Timage(GetControl('IDEV4')).Picture.Assign(Timage(GetControl('IDEV')).Picture.Graphic);
    AssignDrapeau(Timage(GetControl('IDEVEUR')), V_PGI.DevisePivot);
    SetControlCaption('DEVEUR', V_PGI.DevisePivot);
  end;

  {On offre la possiblité d'accéder à la chancellerie que si on est en devises}
  ppMenu.Items[2].Visible := Vis;
  {Gestion en première page de l'affichage de la devise de la transaction}
  lb := THLabel(GetControl('LBDEVISE'));
  lb.Caption := TraduireMemoire('Les montants sont exprimés en ' + RechDom('TTDEVISE', Dev, False));
  Application.ProcessMessages;
  TImage(GetControl('IDEV')).Left := lb.Width + lb.Left + 4;
end;

{---------------------------------------------------------------------------------------}
procedure TOM_TROPCVM.OnAfterShow;
{---------------------------------------------------------------------------------------}
begin
  {$IFNDEF EAGLCLIENT}
  {Au chargement le OnChange du combo est exécuté en eAgl mais pas en 2/3, d'où la ligne ci-dessous}
  ChargerDrapeaux;
  {$ENDIF}
end;

{---------------------------------------------------------------------------------------}
procedure TOM_TROPCVM.MettreLeFocus(ControlName : string; NumPage : Integer);
{---------------------------------------------------------------------------------------}
begin
  if pcControl.ActivePageIndex <> NumPage then
    pcControl.ActivePageIndex := NumPage;
  SetFocusControl(ControlName);
end;

{---------------------------------------------------------------------------------------}
procedure TOM_TROPCVM.SetNumTransac(Value : string);
{---------------------------------------------------------------------------------------}
begin
  if (Value = '') or (FNumTransac = Value) then
    Exit;

  if not NumActif then begin
    {JP 10/07/07 : FQ 10501 : on passe ici à la suite de l'affectation dans le changefield et
                   dans le setcodetransac : dans le cas du CodeTranc, on envoie comme paramètre
                   un valeur Combo => si on n'est pas en insertion, on n'affecte pas FNumTransac}
    if Length(Value) > 3 then
      FNumTransac := Value;
    Exit;
  end;

  {Value vaut le code transaction}
  FNumTransac := InitNumTransac(CODEMODULE, V_PGI.CodeSociete, Value);
  {FCompteur contient le numéro auto-incrémenté qui sera mis à jour dans la table COMPTEURTRESO}
  FCompteur := Copy(FNumTransac, Length(FNumTransac) - 6, 7) ;

  SetField('TOP_NUMOPCVM', FNumTransac) ;
end;

{---------------------------------------------------------------------------------------}
procedure TOM_TROPCVM.SetCodTransac(Value : string);
{---------------------------------------------------------------------------------------}
begin
  if FCodTransac <> Value then begin
    if (Value <> '') then
      NumTransac := Value;
    FCodTransac := Value;
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOM_TROPCVM.SetCotation(Value : Double);
{---------------------------------------------------------------------------------------}
begin
  if FCotation <> Value then begin
    FCotation := Value;
    if IsDSModif then begin
      SetField('TOP_COTATIONACH', FormateMontant(Value, 5));
      SetField('TOP_DEVISE', nDevise);
    end;
  end;
end;

{---------------------------------------------------------------------------------------}
function TOM_TROPCVM.GetCotation : Double;
{---------------------------------------------------------------------------------------}
begin
  if FCotation = 0 then begin
    if not IsDSModif then ForceUpdate;
    Cotation := 1;
  end;
  Result := Valeur(FormateMontant(FCotation, 5));
end;

{---------------------------------------------------------------------------------------}
function TOM_TROPCVM.IsNumActif : Boolean;
{---------------------------------------------------------------------------------------}
begin
  {Est-on en train de créer un enregistrement}
  Result := (TFFiche(Ecran).TypeAction in [taCreat..taCreatOne]) or (DS.State = dsInsert);
end;

{---------------------------------------------------------------------------------------}
function TOM_TROPCVM.IsDSModif : Boolean;
{---------------------------------------------------------------------------------------}
begin
  {Est-on en train de créer ou de modifier un enregistrement}
  Result := DS.State in [dsInsert, dsEdit];
end;

{---------------------------------------------------------------------------------------}
procedure TOM_TROPCVM.GererControls;
{---------------------------------------------------------------------------------------}
begin
  {Accès aux écrans de paramétage}
  TToolbarButton97(GetControl('BPARAMCOURS')).OnClick := GetCours;
  TToolbarButton97(GetControl('BTOTDEV'    )).OnClick := CalcEnvDevClick;
  TToolbarButton97(GetControl('BTOTEUR'    )).OnClick := CalcEnvEurClick;
  {$IFDEF EAGLCLIENT}
  THValComboBox(GetControl('TOP_CODEOPCVM')).OnChange := OpcvmChange;
  THValComboBox(GetControl('TOP_PORTEFEUILLE')).OnChange := PortefChange;
  {$ELSE}
  THDBValComboBox(GetControl('TOP_CODEOPCVM')).OnChange := OpcvmChange;
  THDBValComboBox(GetControl('TOP_PORTEFEUILLE')).OnChange := PortefChange;
  {$ENDIF}

  {Boutons de calcul}
  TBitBtn(GetControl('BACHAT')).OnClick := CalcAchatClick;
  TBitBtn(GetControl('BVENTE')).OnClick := CalcVenteClick;

  TFFiche(Ecran).OnAfterFormShow := OnAfterShow;
  pcControl := TFFiche(Ecran).Pages;

  GereMenuPop;

  if VenteOk then begin
    THNumEdit(GetControl('EDAVENDRE'  )).OnChange := NbVenteChange;
    THEdit   (GetControl('EDDATEVENTE')).OnExit   := DateVenteChange;
  end;
  {26/05/05 : On déactive le bouton si on est en vente}
  SetControlEnabled('BINSERT', GetControlEnabled('BINSERT') and not VenteOk);

  {19/02/07 : pour passer le DS en mode Edit}
  TCheckBox(GetControl('CKREPRISE')).OnClick := CkRepriseChange;
end;

{---------------------------------------------------------------------------------------}
procedure TOM_TROPCVM.GererAffichage;
{---------------------------------------------------------------------------------------}
var
  LectOk : Boolean;
begin
  {Pour désactiver les boutons si on est en lecture seule}
  LectOk := TFFiche(Ecran).FTypeAction = taConsult;
  {Le group box GBDEFINITION ne peut être actif qu'en achat}
  SetControlEnabled('TOP_PORTEFEUILLE', not VenteOk);
  SetControlEnabled('TOP_CODEOPCVM', not VenteOk);
  SetControlEnabled('TOP_TRANSACTION', not VenteOk);
  SetControlEnabled('TOP_LIBELLE', not VenteOk);
  {Le group box GBIDENTIFICATION ne peut être actif qu'en achat}
  SetControlEnabled('TOP_GENERAL', not VenteOk);
  {Le group box GBACHAT ne peut être actif qu'en achat}
  SetControlEnabled('TOP_REFERENCE', not VenteOk);
  SetControlEnabled('TOP_BASECALCUL', not VenteOk);
  SetControlEnabled('TOP_DATEACHAT', not VenteOk);
  SetControlEnabled('TOP_NBPARTACHETE', not VenteOk);
  SetControlEnabled('BACHAT', not VenteOk and not LectOk);
  SetControlEnabled('BPARAMCOURS', not VenteOk and not LectOk);
  {Le group box GBVENTE ne peut être actif qu'en Vente}
  SetControlEnabled('EDAVENDRE', VenteOk);
  SetControlEnabled('EDDATEVENTE', VenteOk);
  SetControlEnabled('BVENTE', VenteOk and not LectOk);

  {Gestion des controles du GroupBox GBMONTANT}
  SetControlEnabled('EDTOTALDEVACH', not VenteOk);
  SetControlEnabled('EDTOTALEURACH', not VenteOk);
  SetControlVisible('LBDEV', VenteOk or LectOk);
  SetControlVisible('LBEUR', VenteOk or LectOk);
  SetControlVisible('BTOTDEV', not VenteOk and not LectOk);
  SetControlVisible('BTOTEUR', not VenteOk and not LectOk);
  SetControlVisible('CKREPRISE', not VenteOk and (TFFiche(Ecran).FTypeAction <> taConsult)); {06/11/06 : FQ 10380 et 19/02/07}
end;

{Calcul d'une transaction d'achat}
{---------------------------------------------------------------------------------------}
procedure TOM_TROPCVM.CalcAchatClick (Sender : TObject);
{---------------------------------------------------------------------------------------}
var
  MntTemp  : Double;
  MntTVA   : Double;
  MntTotal : Double;
  MntEur   : Double;
  Obj      : TObjComOPCVM;
  tEcr     : TOB;
begin
  {On s'assure que le compte général est renseigné, sinon les Frais et commissions ne seront pas calculés}
  if GetField('TOP_GENERAL') = '' then begin
    PGIBox(TraduireMemoire('Veuillez renseigner le compte bancaire.'));
    MettreLeFocus('TOP_GENERAL', 0);
    Exit;
  end;

  ForceUpdate;
  {Calcul du montant horstaxes et commissions}
  MntTemp := Valeur(GetField('TOP_PRIXUNITAIRE')) * Valeur(GetField('TOP_NBPARTACHETE'));
  SetField('TOP_MONTANTACHDEV', FormateMontant(MntTemp, 5));
  MntTotal := MntTemp;
  {Création de l'objet gérant le calcul des commissions et des frais}
  Obj := TObjComOPCVM.Create(VarToStr(GetField('TOP_GENERAL')), VarToStr(GetField('TOP_TRANSACTION')),
                             VarToDateTime(GetField('TOP_DATEACHAT')), False, False);
  tEcr := TOB.Create('TRECRITURE', nil, -1);
  try
    {Gestion des devises : on gère tout en devise pivot pour éviter les problèmes car on peut se retrouver
     avec trois devises : celle du frais, celle de l'OPCVM et celle du compte !!!}
    tEcr.SetDouble('TE_COTATION'  , Valeur(GetField('TOP_COTATIONACH')));
    tEcr.SetDouble('TE_MONTANTDEV', MntTemp / Cotation);
    tEcr.SetString('TE_DEVISE'    , V_PGI.DevisePivot);

    {Mise en place des moyens nécessaires au calcul des frais et commissions}
    Obj.GenererCommissions(tEcr);
    {Récupération du montant total en devise pivot}
    SetField('TOP_MONTANTACH', FormateMontant(tEcr.GetDouble('TE_MONTANT'), 5));
    MntEur := tEcr.GetDouble('TE_MONTANT');

    {Calcul des frais : en trésorerie, ils sont stockés TTC, donc pour l'affichage il faut un découpage}
    MntTemp := Obj.MntFrais;
    {MntTVA  := Obj.MntTVA; FQ 10200 : la Tva est gérée en totalité avec les commissions}

    {Conversion dans la devise de l'OPCVM}
    SetField('TOP_FRAISACHDEV', FormateMontant(MntTemp * Cotation, 5));
    SetField('TOP_FRAISACH'   , FormateMontant(MntTemp, 5));
    MntTotal := MntTotal + (MntTemp * Cotation);
    MntEur   := MntEur + MntTemp;

    {Calcul des Commissions}
    MntTemp := Obj.MntCommission;
    MntTVA  := Obj.MntTVA;
    SetField('TOP_COMACHAT', Valeur(FormateMontant(MntTemp, 5)));
    SetField('TOP_TVAACHAT'   , FormateMontant(MntTVA, 5));
    SetField('TOP_COMACHATDEV', Valeur(FormateMontant(MntTemp * Cotation, 5)));
    SetField('TOP_TVAACHATDEV', FormateMontant(MntTVA * Cotation, 5));
    MntTotal := MntTotal + MntTemp * Cotation  + MntTVA * Cotation;
    MntEur   := MntEur + MntTemp + MntTVA;

    {Calcul du Total}
    SetControlText('EDTOTALDEVACH', FormateMontant(MntTotal, 5));
    SetControlText('EDTOTALEURACH', FormateMontant(MntEur, 5));
  finally
    if Assigned(Obj ) then FreeAndNil(Obj );
    if Assigned(tEcr) then FreeAndNil(tEcr);
  end;
end;

{Calcul d'une transaction de vente : il s'agit forcément d'une vente partielle
{---------------------------------------------------------------------------------------}
procedure TOM_TROPCVM.CalcVenteClick(Sender : TObject);
{---------------------------------------------------------------------------------------}
begin
  {On s'assure que le nombre de parts à vendre est cohérents}
  if IsNbPartsOk then begin
    ForceUpdate;
    {Mise à jour du type de vente}
    SetField('TOP_TYPEVENTE', GetTypeVente(GetField('TOP_TYPEVENTE'), vop_Partiel));
    {Calcul des commissions et frais}
    CalculVente;
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOM_TROPCVM.InitEcr(var tEcr : TOB);
{---------------------------------------------------------------------------------------}
begin
  tEcr.SetString('TE_GENERAL'   , VarToStr(GetField('TOP_GENERAL')));
  tEcr.SetString('TE_CODEFLUX'  , VarToStr(GetField('TOP_TRANSACTION')));
  tEcr.SetDouble('TE_SOLDEDEV'  , Valeur(GetField('TOP_PRIXUNITAIRE')));
  tEcr.SetDouble('TE_COTATION'  , Valeur(GetField('TOP_COTATIONACH')));
  tEcr.SetDouble('TE_MONTANTDEV', Valeur(GetControlText('EDTOTALEURACH')));
  tEcr.SetDateTime('TE_DATECOMPTABLE', VarToDateTime(GetField('TOP_DATEACHAT')));
end;

{---------------------------------------------------------------------------------------}
procedure TOM_TROPCVM.MajZones;
{---------------------------------------------------------------------------------------}
var
  MntTemp  : Double;
  MntTotal : Double;
  MntEur   : Double;
begin
  MntTotal := 0;
  MntEur   := 0;
  MntTemp := ObjEnvers.Montant;
  {Récupération du montant total en devise pivot}
  SetField('TOP_MONTANTACH', FormateMontant(MntTemp, 5));
  SetField('TOP_MONTANTACHDEV', FormateMontant(MntTemp * Cotation, 5));
  MntTotal := MntTotal + (MntTemp * Cotation);
  MntEur   := MntEur + MntTemp;

  {Calcul des frais}
  MntTemp := ObjEnvers.Frais;
  SetField('TOP_FRAISACHDEV', FormateMontant(MntTemp * Cotation, 5));
  SetField('TOP_FRAISACH'   , FormateMontant(MntTemp, 5));
  MntTotal := MntTotal + (MntTemp * Cotation);
  MntEur   := MntEur + MntTemp;

  {Calcul des Commissions}
  MntTemp := ObjEnvers.Commission;
  SetField('TOP_COMACHAT', FormateMontant(MntTemp, 5));
  SetField('TOP_COMACHATDEV', FormateMontant(MntTemp * Cotation, 5));
  MntTotal := MntTotal + (MntTemp * Cotation);
  MntEur   := MntEur + MntTemp;

  {10/02/05 : FQ 10200 : le calcul des frais doit se faire après celui des commissions}
  MntTemp := ObjEnvers.TvaFrais;
  SetField('TOP_TVAACHATDEV', FormateMontant(MntTemp * Cotation, 5));
  SetField('TOP_TVAACHAT'   , FormateMontant(MntTemp, 5));
  MntTotal := MntTotal + (MntTemp * Cotation);
  MntEur   := MntEur + MntTemp;

  {Calcul du Total}
  SetControlText('EDTOTALDEVACH', FormateMontant(MntTotal, 5));
  SetControlText('EDTOTALEURACH', FormateMontant(MntEur, 5));
  SetField('TOP_NBPARTACHETE', ObjEnvers.NbPart);
end;

{Initialisation, calcul et maj des zones p
{---------------------------------------------------------------------------------------}
procedure TOM_TROPCVM.LancerCalcul(Sender, ZoneDev : string);
{---------------------------------------------------------------------------------------}
var
  tEcr : TOB;
  x    : Double;
begin
  if VenteOk then Exit;
  {On s'assure que le compte général est renseigné, sinon les Frais et commissions ne seront pas calculés}
  if GetField('TOP_GENERAL') = '' then begin
    PGIBox(TraduireMemoire('Veuillez renseigner le compte bancaire.'));
    MettreLeFocus('TOP_GENERAL', 0);
    Exit;
  end;

  ForceUpdate;

  if TWinControl(GetControl(Sender)) = Ecran.ActiveControl then begin
    {Les SpeedButtons ne prenant pas le Focus, on simule une tabulation pour être sûr que
     le GetControlText renvoie bien la valeur saisie}
    PostMessage(TWinControl(GetControl(Sender)).Handle, WM_KEYDOWN, VK_TAB, 0);
    Application.ProcessMessages;
  end;

  if Sender = 'EDTOTALDEVACH' then
    x := Valeur(GetControlText(Sender)) / Cotation
  else
    x := Valeur(GetControlText(Sender)) * Cotation;

  {Mise à jour du champ en devise pivot}
  SetControlText(ZoneDev, FloatToStr(x));

  tEcr := TOB.Create('TRECRITURE', nil, -1);
  try
    {Initialisation du record}
    InitEcr(tEcr);
    {Si le prix unitaire n'est pas rensegne, il est inutile d'aller plus loin}
    if tEcr.GetDouble('TE_SOLDEDEV') = 0 then begin
      HShowMessage('0;' + Ecran.Caption +';Veuillez renseigner le cours de l''OPCVM.;W;O;O;O;', '', '');
    end;
    {Calcul récursif}
    ObjEnvers.CalculerNbPart(tEcr);
    {Mise à jour de l'écran}
    MajZones;
  finally
    if Assigned(tEcr) then FreeAndNil(tEcr);
  end;
end;

{Calcul d'une transaction d'achat en devise à l'envers}
{---------------------------------------------------------------------------------------}
procedure TOM_TROPCVM.CalcEnvDevClick(Sender : TObject);
{---------------------------------------------------------------------------------------}
begin
  if VenteOk then Exit;
  LancerCalcul('EDTOTALDEVACH', 'EDTOTALEURACH');
end;

{Calcul d'une transaction d'achat en devise pivot à l'envers}
{---------------------------------------------------------------------------------------}
procedure TOM_TROPCVM.CalcEnvEurClick(Sender : TObject);
{---------------------------------------------------------------------------------------}
begin
  if VenteOk then Exit;
  LancerCalcul('EDTOTALEURACH', 'EDTOTALDEVACH');
end;

{Vérifie si le nombre de parts est cohérent
{---------------------------------------------------------------------------------------}
function TOM_TROPCVM.IsNbPartsOk : Boolean;
{---------------------------------------------------------------------------------------}
var
  neVendre : THNumEdit;
begin
  nbAVendre := 0;
  Result := False;
  neVendre := THNumEdit(GetControl('EDAVENDRE'));

  if neVendre.Value < 0 then neVendre.Value := Abs(neVendre.Value);

  if neVendre.Value > Valeur(GetField('TOP_NBPARTACHETE')) then
    HShowMessage('0;' + Ecran.Caption + ';Le nombre de parts à vendre est supérieur au nombre de parts achetées.;W;O;O;O;', '', '')
  else if (neVendre.Value + BufferAvantModif.GetDouble('TOP_NBPARTVENDU')) > Valeur(GetField('TOP_NBPARTACHETE')) then
    HShowMessage('0;' + Ecran.Caption + ';Le nombre de parts vendues et à vendre est supérieur au nombre de parts achetées.;W;O;O;O;', '', '')
  else if neVendre.Value = 0 then
    HShowMessage('0;' + Ecran.Caption + ';Veuillez sélectionner un nombre de parts à vendre.;W;O;O;O;', '', '')
  else begin
    Result := True;
    {15/12/06; FQ 10373 : Les nombres de parts passe d'un Integer à un Double}
    nbAVendre := Valeur(neVendre.Text);
  end;
end;

{Calcul les montants de la vente
{---------------------------------------------------------------------------------------}
procedure TOM_TROPCVM.CalculVente;
{---------------------------------------------------------------------------------------}
var
  T : TOB;
  D : Double;
  Obj : TObjCalculPMVR;
begin
  {Charge l'enregistrement courant}
  ObjVente.tOPCVM.ClearDetail;
  T := TOB.Create('TROPCVM', ObjVente.tOPCVM, -1);
  T.SetString('TOP_NUMOPCVM', GetField('TOP_NUMOPCVM'));
  T.LoadDB;
  T.AddChampSupValeur('NBPARTS', Valeur(GetControlText('EDAVENDRE')));

  {Si la génération de la vente s'est bien passée, on met à jour les champs de la Tom}
  if ObjVente.GenererVente(T, True, VarToDateTime(GetControlText('EDDATEVENTE')), D) then begin
    SetField('TOP_NBPARTVENDU'  , T.GetValue('TOP_NBPARTVENDU'  ));
    SetField('TOP_DATEFIN'      , T.GetValue('TOP_DATEFIN'      ));
    SetField('TOP_MONTANTVENDEV', T.GetValue('TOP_MONTANTVENDEV'));
    SetField('TOP_MONTANTVEN'   , T.GetValue('TOP_MONTANTVEN'   ));
    SetField('TOP_FRAISVENDEV'  , T.GetValue('TOP_FRAISVENDEV'  ));
    SetField('TOP_FRAISVEN'     , T.GetValue('TOP_FRAISVEN'     ));
    SetField('TOP_TVAVENTEDEV'  , T.GetValue('TOP_TVAVENTEDEV'  ));
    SetField('TOP_TVAVENTE'     , T.GetValue('TOP_TVAVENTE'     ));
    SetField('TOP_COMVENTEDEV'  , T.GetValue('TOP_COMVENTEDEV'  ));
    SetField('TOP_COMVENTE'     , T.GetValue('TOP_COMVENTE'     ));
    SetField('TOP_COTATIONVEN'  , T.GetValue('TOP_COTATIONVEN'  ));
    MajTotauxVente;
    Obj := TObjCalculPMVR.CreateAvecTob(ObjVente.tVente, ObjVente.tOPCVM, T.GetString('TOP_GENERAL'), T.GetString('TOP_CODEOPCVM'));
    try
      Obj.CalculVente(True);
      Obj.RecupVentes(ObjVente.tVente);
    finally
      FreeAndNil(Obj);
    end;

    CalVenteOk := False;
  end;
end;


{---------------------------------------------------------------------------------------}
procedure TOM_TROPCVM.OnCancelRecord;
{---------------------------------------------------------------------------------------}
begin
  inherited;
  SetControlEnabled('BINSERT', not VenteOk);
  {26/05/05 : Lors de l'annulation il faut mettre à jour les zones qui ne correspondent à
              pas des champs de la table TROPCVM}
  if VenteOk then MajTotauxVente
             else MajTotauxAchat;
end;

{04/10/06 : Gestion du NoDossier et du code société}
{---------------------------------------------------------------------------------------}
procedure TOM_TROPCVM.SetInfosDossier;
{---------------------------------------------------------------------------------------}
var
  r : TDblResult;
begin
  if IsTresoMultiSoc then begin
    if VarToStr(GetField('TOP_GENERAL')) <> '' then begin
      r := GetInfosSocFromBQ(VarToStr(GetField('TOP_GENERAL')), True);
      SetField('TOP_SOCIETE', r.RC);
      SetField('TOP_NODOSSIER', r.RE);
      SetControlText('LSOCIETE', r.RT);
    end
    else begin
      SetField('TOP_SOCIETE', V_PGI.CodeSociete);
      SetField('TOP_NODOSSIER', V_PGI.NoDossier);
      SetControlText('LSOCIETE', V_PGI.NomSociete);
    end;
  end
  else
    SetControlText('LSOCIETE', V_PGI.NomSociete)
end;

{19/02/07 : La checkBox n'étant pas un champ, il faut mettre à jour l'état du datasource
{---------------------------------------------------------------------------------------}
procedure TOM_TROPCVM.CkRepriseChange(Sender: TObject);
{---------------------------------------------------------------------------------------}
begin
  if not (DS.State in [dsInsert, dsEdit]) then DS.Edit;
end;

{10/06/07 : FQ 10409 : Visualisation des ventes
{---------------------------------------------------------------------------------------}
procedure TOM_TROPCVM.LancerVentes(Sender : TObject);
{---------------------------------------------------------------------------------------}
var
  T   : TOB;
  SQL : string;
  Obj : TObjEtats;
begin
  SQL := 'SELECT TVE_LIBELLE, TVE_NUMVENTE, TVE_NUMLIGNE, TVE_DATEVENTE, TVE_NBVENDUE, ' +
         'TVE_COURSEUR, TVE_MONTANT, TVE_PMVALUENETTE, TVE_RENDEMENTNET FROM TRVENTEOPCVM ' +
         'WHERE TVE_NUMTRANSAC = "' + VarToStr(GetField('TOP_NUMOPCVM')) + '"';
  T := TOB.Create('_VENTE', nil, -1);
  try
    T.LoadDetailFromSQL(SQl);
    if T.Detail.Count > 0 then begin
      Obj := TObjEtats.Create(TraduireMemoire('Ventes sur la transaction ') + VarToStr(GetField('TOP_NUMOPCVM')), T);
      try
        Obj.MajTitre(0, TraduireMemoire('Libellé'));
        Obj.MajTitre(1, TraduireMemoire('N° Vente'));
        Obj.MajTitre(2, TraduireMemoire('Ligne'));
        Obj.MajTitre(3, TraduireMemoire('Date Vente'));
        Obj.MajTitre(4, TraduireMemoire('Parts Vendues'));
        Obj.MajTitre(5, TraduireMemoire('Cours (' + V_PGI.DevisePivot + ')'));
        Obj.MajTitre(6, TraduireMemoire('Mnt vente (' + V_PGI.DevisePivot + ')'));
        Obj.MajTitre(7, TraduireMemoire('P/M Value Nette (' + V_PGI.DevisePivot + ')'));
        Obj.MajTitre(8, TraduireMemoire('Rendement Net (en %)'));

        Obj.MajAlign(0, ali_Gauche);
        Obj.MajAlign(1, ali_Droite);
        Obj.MajAlign(2, ali_Droite);
        Obj.MajAlign(3, ali_Centre);
        Obj.MajAlign(4, ali_Droite);
        Obj.MajAlign(5, ali_Droite);
        Obj.MajAlign(6, ali_Droite);
        Obj.MajAlign(7, ali_Droite);
        Obj.MajAlign(8, ali_Droite);

        Obj.Imprimer;
      finally
        FreeAndNil(Obj);
      end;
    end
    else
      PGIBox(TraduireMemoire('Aucune vente n''a été trouvée pour cette transaction'));
  finally
    FreeAndNil(T);
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOM_TROPCVM.GereMenuPop;
{---------------------------------------------------------------------------------------}
begin
  ppMenu    := TPopupMenu(GetControl('PPPARAM'));
  ppMenu.Items[0].OnClick := LancerCours;
  ppMenu.Items[1].OnClick := LancerFrais;
  ppMenu.Items[2].OnClick := LancerDevise;
  {ppMenu.Items[3] : c'est un séparateur}
  {10/06/07 : FQ 10409 : Visualisation des ventes}
  ppMenu.Items[4].OnClick := LancerVentes;
end;

initialization
  RegisterClasses([TOM_TROPCVM]);

end.
