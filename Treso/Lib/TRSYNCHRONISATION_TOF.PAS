{ Unité : Source TOF de la FICHE : TRSYNCHRONISATION
--------------------------------------------------------------------------------------
    Version   |  Date  | Qui |  Commentaires
--------------------------------------------------------------------------------------
 0.91          02/09/03  JP   Création de l'unité
 1.20.001.001  24/03/04  JP   Refonte de l'unité pour mettre en place l'échéancier
                              et la gestion des filtres
 1.50.001.001  08/04/04  JP   Mise en place de la synchronisation de s flux prévisionnels
 1.50.001.001  07/05/04  JP   Suppression des comptes de caisse de la Trésorerie IN ("BQE, "CAI") => = "BQE"
 6.00.001.001  09/07/04  JP   Correction du doublement des écritures bancaires en première synchronisation FFInit
                              et de "G_NATUREGENE est une colonne invalide"
 6.00.015.001  20/09/04  JP   FQ 10144 : gestion d'un écran pour les écritures non traitées
                              Ajout de la case à cocher gestion ckReglement
 6.00.015.001  24/09/04  JP   FQ 10104 : aménagement pour tenir compte de la gestion de
                              la synchronisation dans SaisieTr (saisie d'effets et règlement)
 6.20.001.003  10/02/05  JP   FQ 10201 : Erreur de frappe dans la première synchronisation
 6.30.001.002  28/02/05  JP   Gestion des Banques prévisionnelles : ajout d'une nouvelle valeur à E_TRESOSYNCHRO
 6.50.001.001  26/05/05  JP   Amélioration de la gestion de la date première de première synchro
                              Optimisation de la clause Where des requêtes de traitement.
 6.50.001.009  22/07/05  JP   Optimisation du recalcul des soldes avec ajout d'une option
                              FQ 10268 : nouvelle gestion des modes de paiement vides
 7.09.001.001  08/08/06  JP   Mise en place de la sycnhronisation multi sociétés : SynchroniseUnDossier
 7.09.001.001  03/10/06  JP   Mise en place de la première sycnhronisation multi sociétés : FFInit
 7.09.001.001  20/11/06  JP   Gestion des dates antérieures au millésime courant
 8.01.001.010  03/04/07  JP   Mise en place de la référence de pointage
 8.01.001.017  29/05/07  JP   Ajout du champ E_NATURETRESO pour TE_CODERAPPRO
 8.10.001.004  08/08/07  JP   Gestion des confidentialités
 8.10.006.001  28/11/07  JP   Mauvaise gestion des comptes courants dans la clause where (pb vu chez SIC)
--------------------------------------------------------------------------------------}
unit TRSYNCHRONISATION_TOF ;
                           
interface

uses {$IFDEF VER150} variants,{$ENDIF}
  StdCtrls, Controls, Classes,
  {$IFNDEF EAGLCLIENT}
  {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} FE_Main,
  {$ELSE}
  MaineAGL,
  {$ENDIF}
  Forms, sysutils, ComCtrls, HCtrls, HEnt1, HMsgBox, uTob, UTOF, HTB97, UObjGen,
  Constantes, ParamSoc, Ent1, UtilSoc, Graphics,
  {Pour la gestion des filtres}
  Menus, UObjFiltres;

type
  TOF_TRSYNCHRONISATION = Class (TOF)
    procedure OnUpdate             ; override;
    procedure OnLoad               ; override;
    procedure OnArgument(S :string); override;
    procedure OnClose              ; override;
  private
    edDateDe     : THEdit;
    edDateA      : THEdit;
    cbGeneral    : THMultiValComboBox;
    ckEcheancier : TCheckBox;
    ckEffet      : TCheckBox;
    ckReglement  : TCheckBox; {20/09/04}
    ckBancaire   : TCheckBox; {08/08/06}

    ClauseWhere   : string;
    FiltreComptes : Boolean;
    FiltreDate    : Boolean;
    ObjParam      : TObjBanquePrevi;
    Dt1ereSyncho  : TDateTime; {26/05/05}

    function  ConstituerClauseWh(NoDos : string) : string;
    procedure PremiereClick    (Sender : TOBject);
    procedure ckEcheancierClick(Sender : TOBject);
    procedure ckReglementClick (Sender : TOBject); {20/09/04}
    procedure ckEffetClick     (Sender : TOBject);
    procedure ckBancaireClick  (Sender : TOBject); {08/08/06}
    procedure BParamSocClick   (Sender : TObject);
    procedure DossierChange    (Sender : TObject);
    function  FFInit           (NomBase: string ) : Boolean; {Première synchronisation}
    procedure MajLibelle; {Affiche les banques prévisionnelles}
    procedure PoseFiltreSurCpte; {Filtre les comptes sur la nature et le dossier}
  protected
    ObjFiltre : TObjFiltre;
    TobRejets : TOB;
    TobDate   : TOB;
    lCalSolde : TStringList;
    NbEcrit   : Integer;

    {FQ 10239 : s'assure de la cohérence des dates avec la date de première synchronisation}
    procedure GestionDates(var dtMin, dtMax : TDateTime);
    {26/05/05 : Initialistion de la date de première synchronisation}
    procedure InitDt1ereSyncho;
    procedure SynchroniseUnDossier (NomBase : string);
    function  PremiereSynchro : Boolean;
    procedure TermineSynchronisation;
    procedure MessageFinal;
    {14/09/06 : Retourne la clause}
    function  GetWhereGene(NoDos : string; ConfOnly : Boolean) : string;
  end;

procedure TRLanceFiche_Synchro(Dom, Fiche, Range, Lequel, Arguments : string);

implementation

uses
  {$IFDEF TRCONF}
  ULibConfidentialite,
  {$ENDIF TRCONF}
  HStatus {InitMove}, UProcGen {GetClauseIn}, AccImportTreso, UtilPgi {GetBaseMS},
  Commun{GetInfosFromDossier}, TRGRILLE_TOF, AglInit, UProcSolde;

{---------------------------------------------------------------------------------------}
procedure TRLanceFiche_Synchro(Dom, Fiche, Range, Lequel, Arguments : string);
{---------------------------------------------------------------------------------------}
begin
  AGLLanceFiche(Dom, Fiche, Range, Lequel, Arguments);
end;

{-----------------------------------------------------------------}
function TOF_TRSYNCHRONISATION.FFInit(NomBase : string) : Boolean;
{-----------------------------------------------------------------}
var
  Tab : string;
begin
  {03/10/06 : Ajout de la base devant la table si besoin}
  if NomBase <> V_PGI.SchemaName then Tab := GetTableDossier(NomBase, 'ECRITURE')
                                 else Tab := 'ECRITURE';

  Result := True;
  BeginTrans;
  try
    {REQUETES A SUPPRIMER APRES LA V6 : ELLES CONCERNENT LES DEUX/TROIS CLIENTS INSTALLES DEPUIS LA V5}
    {JP 08/04/04 : Avec l'intégration de l'échéancier, on a modifié les natures : il n'y a plus que trois valeurs
                   possibles R(éalisé), P(révisionnelle), S(imulation)
                   Mise à jour des natures et de ce Champ}
    ExecuteSQL('UPDATE TRECRITURE SET TE_NATURE = "R" WHERE TE_NATURE = "IC"');
    ExecuteSQL('UPDATE TRECRITURE SET TE_NATURE = "P" WHERE TE_NATURE IN ("V", "") OR TE_NATURE IS NULL');
    {28/07/04 : initialisation du Champ TE_COMMISSION}
  ////  ExecuteSQL('UPDATE TRECRITURE SET TE_COMMISSION = "S" WHERE TE_COMMISSION = "" OR TE_COMMISSION IS NULL');
    {FIN : REQUETES A SUPPRIMER APRES LA V6 : ELLES CONCERNENT LES DEUX/TROIS CLIENTS INSTALLES DEPUIS LA V5}


    {JP 08/04/04 : modifications des requêtes pour tenir compte de l'échéancier}
    Initmove(14, TraduireMemoire('Première synchronisation'));
    {On fait un premier filtre sur les écritures qui sont à priori traitables}
    ExecuteSQL('UPDATE ' + Tab + ' SET E_TRESOSYNCHRO = "TMP" WHERE E_QUALIFPIECE = "N" AND E_ECRANOUVEAU IN ("H", "N") ' +
               'AND E_DATECOMPTABLE  >= "' + USDateTime(StrToDate(GetControlText('EDANTERIEURE'))) +
               {09/07/04 : Pour les quelques clients installés, sinon leurs écritures de banques étaient de nouveau synchronisées}
               '" AND E_TRESOSYNCHRO <> "' + ets_Synchro + '"');
    MoveCur(2);
    {Mise à jour des écritures bancaires à importer}
    ExecuteSQL('UPDATE ' + Tab + ' SET E_TRESOSYNCHRO = "CRE" WHERE E_TRESOSYNCHRO = "TMP" AND E_GENERAL IN ' +
                       '(SELECT G_GENERAL FROM GENERAUX WHERE G_NATUREGENE = "BQE")');
    MoveCur(2);
    {Mise à jour des écritures prévisionnelles à importer}
    ExecuteSQL('UPDATE ' + Tab + ' SET E_TRESOSYNCHRO = "CRE" WHERE E_TRESOSYNCHRO = "TMP" AND E_ETATLETTRAGE IN ("AL", "PL") AND ' +
                 '(E_GENERAL IN (SELECT G_GENERAL FROM GENERAUX WHERE G_NATUREGENE IN ("COD", "COC", "COF", "COS", "TIC", "TID"))' + {10/02/05 : FQ 10201}
                  'AND ((E_JOURNAL IN (SELECT J_JOURNAL FROM JOURNAL WHERE J_NATUREJAL IN ("ACH", "VTE"))) ' +
                        'OR ' +
                        '(E_JOURNAL IN (SELECT J_JOURNAL FROM JOURNAL WHERE J_NATUREJAL = "OD") AND E_NATUREPIECE <> "ECC")))');

    MoveCur(4);

    {03/10/06 : Gestion des comptes courants}
    if IsTresoMultiSoc then
      ExecuteSQL('UPDATE ' + Tab + ' SET E_TRESOSYNCHRO = "CRE" WHERE E_TRESOSYNCHRO = "TMP" AND ' +
                   '(E_GENERAL IN (SELECT G_GENERAL FROM GENERAUX WHERE G_NATUREGENE = "DIV")' +
                    'AND E_GENERAL IN (SELECT CLS_GENERAL FROM ' + GetTableDossier(NomBase, 'CLIENSSOC') + '))');

    MoveCur(4);
    {Mise à jour des écritures que l'on ne traite pas}
    ExecuteSQL('UPDATE ' + Tab + ' SET E_TRESOSYNCHRO = "' + ets_Rien + '" WHERE E_QUALIFPIECE <> "N" OR E_TRESOSYNCHRO = "TMP"' +
               'OR E_DATECOMPTABLE  < "' + USDateTime(StrToDate(GetControlText('EDANTERIEURE'))) + '" OR ' +
               '(E_ECRANOUVEAU <> "H" AND E_ECRANOUVEAU <> "N")');
    MoveCur(2);
    FiniMove;

  { **** REQUÊTE UNIQUE SI LE DÉCOUPAGE NE SE RÉVÈLE PAS INTÉRRESSANT ****
  UPDATE ECRITURE SET E_TRESOSYNCHRO = "CRE" WHERE E_QUALIFPIECE = "N" AND E_ECRANOUVEAU IN ("H", "N") AND  E_DATECOMPTABLE >= "01/01/2003" AND (
            (E_ETATLETTRAGE IN ("AL", "PL") AND
                 (E_GENERAL IN ( SELECT G_GENERAL FROM GENERAUX WHERE G_NATUREGENE IN ("COD", "COC", "COF", "COS", "TID","TID") )
                  AND ((E_JOURNAL IN (SELECT J_JOURNAL FROM JOURNAL WHERE J_NATUREJAL IN ("ACH", "VTE"))) OR
                      (E_JOURNAL IN (SELECT J_JOURNAL FROM JOURNAL WHERE J_NATUREJAL = "OD") AND E_NATUREPIECE <> "ECC"))))
            OR
            (E_GENERAL IN (SELECT G_GENERAL FROM GENERAUX WHERE G_NATUREGENE = "BQE") ))
  }
    CommitTrans;
  except
    on E : Exception do begin
      Result := False;
      RollBack;
      PGIError(E.Message);
    end;
  end;
end;

{---------------------------------------------------------------------------------------}
function TOF_TRSYNCHRONISATION.PremiereSynchro : Boolean;
{---------------------------------------------------------------------------------------}
var
  s : string;
  b : Boolean;
begin
  Result := True;
  {Gestion de la première synchronisation :
   On commence par mettre à jour la table ECRITURE, ce qui peut être long ...}
  if GetCheckBoxState('CKPREMIERE') = cbChecked then begin
    if HShowMessage('0;Première synchronisation;Vous avez choisi de ne pas synchroniser les écritures'#13 +
                    'antérieures au ' + GetControlText('EDANTERIEURE') + '.'#13 +
                    'Ces écritures ne figureront jamais dans la Trésorerie !'#13#13 +
                    'Comfirmez-vous le traitement ?;Q;YN;N;N', '', '') = mrYes then begin

      {03/10/06 : Gestion de la première synchronisation en Multi sociétés :
       - elle s'applique à tous les dossiers du regroupement de Trésorerie}
      if IsTresoMultiSoc then s := FiltreBaseTreso(True)
                         else s := V_PGI.SchemaName + ';';

      if s = '' then begin
        PGIError(TraduireMemoire('Impossible de récupérer le nom des bases.'));
        Result := False;
        Exit;
      end
      else if IsTresoMultiSoc then begin
        if HShowMessage('0;Première synchronisation;La première synchronisation va s''appliquer à toutes'#13 +
                        'les bases rattachées à la base courante.'#13 +
                        'Comfirmez-vous le traitement ?;Q;YN;N;N', '', '') = mrNo then begin
          Result := False;
          Exit;
        end;
      end;

      {03/10/06 : La première synchronisation se fait base par base}
      b := True;
      while (s <> '') and b do
        b := FFInit(ReadTokenSt(s));

      if b then begin
        SetParamSoc('SO_PREMIERESYNCHRO', True);
        {11/04/05 : FQ 10239 : mémorisation de la date de première synchronisation}
        SetParamSoc('SO_TRDATEPREMSYNCHRO', StrToDate(GetControlText('EDANTERIEURE')));
        {26/05/05 : Mise à jour de Dt1ereSyncho}
        InitDt1ereSyncho;
        {On cache le panel de première synchronisation}
        SetControlVisible('GBPREMIERE', False);
        SetControlChecked('CKPREMIERE', False);
        Result := True;
      end
      else
        Result := False;
    end
    else
      Result := False;
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRSYNCHRONISATION.OnUpDate;
{---------------------------------------------------------------------------------------}
var
  dtMin  : TDateTime;
  dtMax  : TDateTime;
  s      : string;
  Q      : TQuery;
begin
  inherited;
  dtMin := iDate1900;
  dtMax := iDate1900;
  s     := '';

  {On vide les objets de fin de traitements}
  LibereListe(lCalSolde, False);
  TobRejets.ClearDetail;
  TobDate  .ClearDetail;
  NbEcrit := 0;

  {20/09/04 : si on ne demande ni les écritures de règlement, ni les écritures échénacées, on sort}
  if not ckEcheancier.Checked and not ckReglement.Checked then begin
    HShowMessage('0;' + Ecran.Caption + ';Veuillez sélectionner un type d''écritures à traiter.;W;O;O;O', '', '');
    Exit;
  end;

  {13/09/06 : le filtre sur les comptes n'est plus suffisant : il faut ajouter la nature des comptes}
  if cbGeneral.Enabled and not cbGeneral.Tous and (cbGeneral.Text <> '') then begin
    Q := OpenSQL('SELECT BQ_NODOSSIER, BQ_GENERAL FROM BANQUECP WHERE NOT (BQ_CODE IN (' + GetClauseIn(cbGeneral.Value) + '))', True);
    try
      while not Q.EOF do begin
        s := s + Q.FindField('BQ_NODOSSIER').AsString + '|' + Q.FindField('BQ_GENERAL').AsString + ';';
        Q.Next;
      end;
    finally
      Ferme(Q);
    end;

    ObjParam.BqFiltre := s;
    S := '';
  end
  else if ckBancaire.Checked and ckBancaire.Visible then begin
    Q := OpenSQL('SELECT BQ_NODOSSIER, BQ_GENERAL FROM BANQUECP WHERE BQ_NATURECPTE <> "' + tcb_Bancaire + '"', True);
    try
      while not Q.EOF do begin
        s := s + Q.FindField('BQ_NODOSSIER').AsString + '|' + Q.FindField('BQ_GENERAL').AsString + ';';
        Q.Next;
      end;
    finally
      Ferme(Q);
    end;

    ObjParam.BqFiltre := s;
    S := '';
  end
  else
    ObjParam.BqFiltre := '';

  if not PremiereSynchro then Exit;

  {FQ 10239 : on s'assure que la fourchette de date n'est pas antérieure à la date de première synchronisation}
  if edDateDe .Enabled then begin
    {Initialisation des dates}
    GestionDates(dtMin, dtMax);

    if dtMax < Dt1ereSyncho then begin
      PgiError('La fourchette de dates est antérieure à la date de première synchronisation.'#13#13 +
               'Le traitement est abandonné.');
      Exit;
    end

    else if dtMin < Dt1ereSyncho then begin
      if HShowMessage('0;' + Ecran.Caption + ';La date de première synchronisation appartient à la sélection.'#13 +
                      'Seules les écritures postérieures à cette date seront traités.'#13#13 +
                      'Souhaitez-vous poursuivre ?;Q;YN;N;N', '', '') = mrNo then
        Exit;
    end;
  end;

  {Si on demande la synchronisation des écritures prévisionnelles, mais qu'aucune
   banque prévisionnelle n'est paramétrée}
  if ckEcheancier.Checked and not ObjParam.IsAllOk(s) then begin
    {On avertit et on sort}
    HShowMessage('0;' + Ecran.Caption + ';' + s + ';W;O;O;O;', '', '');
    Exit;
  end
  {Si on demande la synchronisation des écritures prévisionnelles, mais certaines
   banque prévisionnelles ne sont pas paramétrées}
  else if ckEcheancier.Checked and (S <> '') then begin
    {On demande confirmation de la synchronisation partielle}
    if HShowMessage('0;' + Ecran.Caption + ';' + s + ';Q;YNC;N;C;', '', '') <> mrYes then Exit;
  end;

  {26/05/05 : Il me semble pas mal de rajouter un test}
  if Dt1ereSyncho = iDate1900 then begin
    if HShowMessage('0;' + Ecran.Caption + ';La première synchronisation n''a encore pas été effectuée'#13 +
                    'ou bien la date de première synchronisation n''a pas été mise à jour.'#13#13 +
                    'Souhaitez-vous poursuivre ?;Q;YN;N;N', '', '') = mrNo then
      Exit;
  end;

  {08/08/06 : Gestion du multi sociétés}
  if IsTresoMultiSoc then begin
    s := GetControlText('MCDOSSIER');
    if (s = '') or THMultiValComboBox(GetControl('MCDOSSIER')).Tous then
      {Récupère les dossier dont la base courante est la base Tréso}
      s := FiltreBaseTreso;

    repeat
      SynchroniseUnDossier(ReadTokenSt(S));
    until s = '';
  end
  else
    SynchroniseUnDossier('');
  {Recalcul des soldes et gestion du message final}
  TermineSynchronisation;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRSYNCHRONISATION.OnArgument(S : String ) ;
{---------------------------------------------------------------------------------------}
var
  Composants : TControlFiltre;
begin
  inherited;
  Ecran.HelpContext := 50000136;
  edDateDe     := THEdit(GetControl('EDDATEDEB'));
  edDateA      := THEdit(GetControl('EDDATEDEB_'));
  cbGeneral    := THMultiValComboBox(GetControl('CBGENERAL'));
  ckEcheancier := TCheckBox(GetControl('CKECHEANCIER'));
  ckEffet      := TCheckBox(GetControl('CKEFFET'));
  ckReglement  := TCheckBox(GetControl('CKREGLEMENT')); {20/09/04}
  ckEcheancier.OnClick := ckEcheancierClick;
  ckEffet     .OnClick := ckEffetClick;
  ckReglement .OnClick := ckReglementClick; {20/09/04}

  {Gestion du paramétrage des banques prévisionnelles}
  TToolBarButton97(GetControl('BPARAMSOC')).OnClick := BParamSocClick;
  ObjParam  := TObjBanquePrevi.Create;
  MajLibelle;

  Ecran.Refresh;
  Application.ProcessMessages;

  {La gestion de la propriété "Enabled" des zones est faite par scripte sur le OnClick
   des radioButtons}
  FiltreComptes := cbGeneral.Enabled;
  FiltreDate    := edDateDe .Enabled;

  {On ne donne la possibilité de faire la première synchro, que si elle n'a pas déjà été
   effectuée}
  SetControlVisible('GBPREMIERE', GetParamSocSecur('SO_PREMIERESYNCHRO', False) = False);
  TCheckBox(GetControl('CKPREMIERE')).OnClick := PremiereClick;

  Composants.PopupF   := TPopUpMenu      (Getcontrol('POPF'));
  Composants.Filtres  := THValComboBox   (Getcontrol('FFILTRES'));
  Composants.Filtre   := TToolBarButton97(Getcontrol('BFILTRE'));
  Composants.PageCtrl := TPageControl    (Getcontrol('PCPAGECONTROL'));
  ObjFiltre := TObjFiltre.Create(Composants, 'TRSYNCHRONIS');

  {08/08/06 : gestion du multi sociétés}
  SetControlVisible('MCDOSSIER' , IsTresoMultiSoc);
  SetControlVisible('LBDOSSIER' , IsTresoMultiSoc);

  if IsTresoMultiSoc then begin
    {Pour n'afficher que les bases qui ont la base courante comme base de référence}
    THMultiValComboBox(GetControl('MCDOSSIER')).Plus := 'DOS_NOMBASE ' + FiltreBaseTreso(False);
    THMultiValComboBox(GetControl('MCDOSSIER')).OnChange := DossierChange;
  end;

  ckBancaire := TCheckBox(GetControl('CKBANCAIRE'));
  ckBancaire.Visible := IsTresoMultiSoc;
  ckBancaire.OnClick := ckBancaireClick;
  ckBancaireClick(ckBancaire);
  
  {Pour la gestion des écritures non traitées par la synchronisation car "incomplètes"}
  TobRejets := TOB.Create('%%%%', nil, -1);
  TobDate   := TOB.Create('%%%%', nil, -1);
  {Liste contenant les objets nécessaire au recalcul des soldes}
  lCalSolde := TStringList.Create;
  SetControlVisible('CKSOLDECOMPLET', not IsNewSoldes);
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRSYNCHRONISATION.OnLoad;
{---------------------------------------------------------------------------------------}
begin
  inherited;
  ObjFiltre.Charger;
  {26/05/05 : On initialise la date de première synchronisation : Dt1ereSyncho}
  InitDt1ereSyncho;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRSYNCHRONISATION.OnClose;
{---------------------------------------------------------------------------------------}
begin
  inherited;
  if Assigned(ObjFiltre) then FreeAndNil(ObjFiltre);
  if Assigned(ObjParam)  then FreeAndNil(ObjParam);
  if Assigned(TobRejets) then FreeAndNil(TobRejets);
  if Assigned(TobDate  ) then FreeAndNil(TobDate);
  if Assigned(lCalSolde) then LibereListe(lCalSolde, True);
end;

{13/09/06 : Filtre de la combo des generaux en fonction de la nature des comptes et des dossiers
{---------------------------------------------------------------------------------------}
procedure TOF_TRSYNCHRONISATION.PoseFiltreSurCpte;
{---------------------------------------------------------------------------------------}
var
  c : string;
  s : string;
  {w : string;
  T : TOB;
  F : TOB;}
begin
  cbGeneral.Plus := '';
  {08/08/07 : Sauf erreur de ma part, tout cela est géré dans FiltreBanqueCp
  if ckBancaire.Checked or not ckBancaire.Visible then
    cbGeneral.Plus := BQCLAUSEWHERE;

  s := GetControlText('MCDOSSIER');
  if (s <> '') and not THMultiValComboBox(GetControl('MCDOSSIER')).Tous then begin
    w := '';
    T := TOB.Create('_DOSSIER', nil, -1);
    T.LoadDetailFromSQL('SELECT DOS_NODOSSIER, DOS_NOMBASE FROM DOSSIER WHERE DOS_NOMBASE IN (' + GetClauseIn(s) + ')');
    try
      repeat
        c := ReadTokenSt(s);
        F := T.FindFirst(['DOS_NOMBASE'], [c], True);
        if Assigned(F) then w := w + ' "' + F.GetString('DOS_NODOSSIER') + '",';
      until s = '';
    finally
      FreeAndNil(T);
    end;

    if w <> '' then begin
      System.Delete(w, Length(w), 1);
      if cbGeneral.Plus <> '' then cbGeneral.Plus := cbGeneral.Plus + ' AND BQ_NODOSSIER IN (' + w + ')'
                              else cbGeneral.Plus := 'BQ_NODOSSIER IN (' + w + ')';
    end;
  end;
  }
  C := '';
  {Si on ne demane que les comptes bancaires ou si l'on n'est pas en multi sociétés}
  if ckBancaire.Checked or not ckBancaire.Visible then
    c := tcb_Bancaire + ';';
  if not THMultiValComboBox(GetControl('MCDOSSIER')).Tous then s := GetControlText('MCDOSSIER')
                                                          else s := '';
  FiltreBanqueCp(cbGeneral.DataType, c, s);
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRSYNCHRONISATION.MajLibelle;
{---------------------------------------------------------------------------------------}
var
  lb : THLabel;
begin
  lb := THLabel(GetControl('LBCOC'));
  lb.Caption := ObjParam.GetLibelle(nc_Client);
       if ObjParam.IsValide (nc_Client)      then lb.Font.Color := clBlue
  else if ObjParam.GetCompte(nc_Client) = '' then lb.Font.Color := clRed
                                             else lb.Font.Color := clPurple;

  lb := THLabel(GetControl('LBCOF'));
  lb.Caption := ObjParam.GetLibelle(nc_Fourni);
       if ObjParam.IsValide (nc_Fourni)      then lb.Font.Color := clBlue
  else if ObjParam.GetCompte(nc_Fourni) = '' then lb.Font.Color := clRed
                                             else lb.Font.Color := clPurple;

  lb := THLabel(GetControl('LBCOD'));
  lb.Caption := ObjParam.GetLibelle(nc_Divers);
       if ObjParam.IsValide (nc_Divers)      then lb.Font.Color := clBlue
  else if ObjParam.GetCompte(nc_Divers) = '' then lb.Font.Color := clRed
                                             else lb.Font.Color := clPurple;

  lb := THLabel(GetControl('LBCOS'));
  lb.Caption := ObjParam.GetLibelle(nc_Salari);
       if ObjParam.IsValide (nc_Salari)      then lb.Font.Color := clBlue
  else if ObjParam.GetCompte(nc_Salari) = '' then lb.Font.Color := clRed
                                             else lb.Font.Color := clPurple;

  lb := THLabel(GetControl('LBTID'));
  lb.Caption := ObjParam.GetLibelle(nc_Debite);
       if ObjParam.IsValide (nc_Debite)      then lb.Font.Color := clBlue
  else if ObjParam.GetCompte(nc_Debite) = '' then lb.Font.Color := clRed
                                             else lb.Font.Color := clPurple;

  lb := THLabel(GetControl('LBTIC'));
  lb.Caption := ObjParam.GetLibelle(nc_Credit);
       if ObjParam.IsValide (nc_Credit)      then lb.Font.Color := clBlue
  else if ObjParam.GetCompte(nc_Credit) = '' then lb.Font.Color := clRed
                                             else lb.Font.Color := clPurple;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRSYNCHRONISATION.PremiereClick(Sender : TOBject);
{---------------------------------------------------------------------------------------}
begin
  SetControlVisible('LBAVER', TCheckBox(Sender).Checked);
  SetControlVisible('LBANTERIEURE', TCheckBox(Sender).Checked);
  SetControlVisible('EDANTERIEURE', TCheckBox(Sender).Checked);
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRSYNCHRONISATION.ckEcheancierClick(Sender : TOBject);
{---------------------------------------------------------------------------------------}
begin
  SetControlEnabled('CKEFFET'  , ckEcheancier.Checked);
  SetControlEnabled('CBNATCPTE', ckEcheancier.Checked);
  SetControlEnabled('LBPREVI'  , ckEcheancier.Checked);
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRSYNCHRONISATION.ckEffetClick(Sender : TObject);
{---------------------------------------------------------------------------------------}
begin
  if ckEffet.Checked then
    THMultiValComboBox(GetControl('CBNATCPTE')).DataType := 'TRNATURETIERS'
  else
    THMultiValComboBox(GetControl('CBNATCPTE')).DataType := 'TRNATURECOMPTE';
  SetControlText('CBNATCPTE', '');
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRSYNCHRONISATION.BParamSocClick(Sender : TObject);
{---------------------------------------------------------------------------------------}
begin
  ParamSociete(False, '', 'SCO_SYNCHRONISATION', '', RechargeParamSoc, ChargePageSoc, SauvePageSoc, InterfaceSoc, 500014);
  ObjParam.Recharger(True) {26/07/05 : FQ 10280 : pour Relancer la requête sur les PramSoc};
  MajLibelle;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRSYNCHRONISATION.DossierChange(Sender : TObject);
{---------------------------------------------------------------------------------------}
begin
  PoseFiltreSurCpte;
end;

{Constitution de la clause where en fonction des critères retenus sur la fiche
 Result contient la clause pour la table ECRITURE et ClauseWhere pour TRECRITURE
{---------------------------------------------------------------------------------------}
function TOF_TRSYNCHRONISATION.ConstituerClauseWh(NoDos : string) : string;
{---------------------------------------------------------------------------------------}
var
  c      : string;
  InSel  : string; {09/07/04}
  dtInit : string;
  WhBqe  : string; {26/05/05}
  dtMin  : TDateTime;
  dtMax  : TDateTime;

    {---------------------------------------------------------------}
    procedure AddCompteCourant;
    {---------------------------------------------------------------}
    begin

      if not IsTresoMultiSoc then Exit; {26/11/08 : IsTresoMultiSoc et non EstMultiSoc}
      if ckBancaire.Checked then Exit;
      {28/11/07 : On ne considère comme compte courant que les 451 et  517 (pour les gens qui ne veulent
                  pas avoir à saisir des opérations pour le bilan sur le 517). Ajout d'un message sur CLIENSSOC}
      WhBqe := ' (' + WhBqe + ' OR (G_NATUREGENE = "DIV" AND J_NATUREJAL IN ("BQE", "OD") ' +
               'AND (E_GENERAL LIKE "451%" OR E_GENERAL LIKE "517%") ))';
    end;

begin
  {FQ 10239 : Initialisation des dates}
  GestionDates(dtMin, dtMax);
  {11/04/05 : FQ 10239 : on ne traite que les écritures postérieures à la première synchronisation}
  if Dt1ereSyncho > iDate1900 then
    dtInit := UsDateTime(Dt1ereSyncho);

  FiltreComptes := cbGeneral.Enabled;
  FiltreDate    := edDateDe .Enabled;
  Result        := '';

  {On a sélectionné une fourchette de dates ...}
  if FiltreDate then begin
    {FQ 10239 : on prend la plus grande des dates entre le ParamSoc et la fourchette basse}
    if dtMin < Dt1ereSyncho then
      dtMin := Dt1ereSyncho;

    c := ' BETWEEN "' + USDateTime(dtMin) + '" AND "' + USDateTime(dtMax) + '"';
    ClauseWhere := 'WHERE (E_DATECOMPTABLE' + c + ') AND';
    Result      := 'WHERE (E_DATECOMPTABLE' + c + ') AND';
  end

  {... Sinon exclusion des écritures antérieures à la date de prémière synchronisation (FQ 10239)}
  else if dtInit <> '' then begin
    ClauseWhere := 'WHERE (E_DATECOMPTABLE >= "' + dtInit + '") AND';
    Result      := 'WHERE (E_DATECOMPTABLE >= "' + dtInit + '") AND';
  end
  {Si l'utilisateur passe outre les avertissements}
  else begin
    ClauseWhere := 'WHERE ';
    Result      := 'WHERE ';
  end;

  c := GetWhereGene(NoDos, not FiltreComptes);
  {Si on n'a rien saisie dans la combo ...}
  if c <> '' then begin
    {26/05/05 : Complétion de ClauseWhere qui sera utilisée pour les écritures de Tréso qui ont été pointées}
    ClauseWhere := ClauseWhere + ' E_GENERAL IN (' + c + ') AND';
    {26/05/05 : Filtre sur les comptes pour la requête principale de synchronisation ...}
    WhBqe       := ' E_GENERAL IN (' + c + ')';
  end
  else begin
    {... Sinon, filtre sur les comptes de nature bancaire pour la requête principale de synchronisation}
    WhBqe := ' G_NATUREGENE = "BQE"';
    {08/08/06 : ajout de la gestion des comptes courants / Titres
     28/11/07 : ne faire le AddCompteCourant que else et pas si on a choisit les comptes à synchroniser}
    AddCompteCourant;
  end;

  {14/09/04 : FQ 10151 : appliquer le filtre que sur les comptes d'effets (a priori 403 et 413)}
  if ckEffet.Checked and ckEffet.Enabled then c := ' AND G_EFFET = "X"'
                                         else c := '';

  {JP 08/04/04 : Gestion de l'échéancier}
  if ckEcheancier.Checked then begin
    {S'il y a un filtre sur les natures de comptes}
    if (GetControlText('CBNATCPTE') <> '') and not THMultiValComboBox(GetControl('CBNATCPTE')).Tous then begin
      if ckReglement.Checked then
        InSel := ' ((G_NATUREGENE IN (' + GetClauseIn(GetControlText('CBNATCPTE')) + ')' + c + ') OR (' + WhBqe + '))'
      else
        InSel := ' (G_NATUREGENE IN (' + GetClauseIn(GetControlText('CBNATCPTE')) + ')' + c + ')' ;

      {13/09/04 : Bel oubli !!  Les écritures n'étaient donc pas filtrées.}
      Result := Result + InSel
    end
    {Si pas de filtre sur la nature des comptes mais que suivi d'effets est coché}
    else if ckEffet.Checked then begin
      {Il faut ici faire attention, car le suivi d'effets est un filtre sur les natures de comptes :
       Donc même si CBNATCPTE est vide il faut filtres sur les collectifs clients et fournisseurs}
      if ckReglement.Checked then
        InSel := ' ((G_NATUREGENE IN ("COC", "COF")' + c + ') OR (' + WhBqe + '))'
      else
        InSel := ' (G_NATUREGENE IN ("COC", "COF")' + c + ')';

      Result := Result + InSel
    end
    else begin
      {Aucun filtre sur les écritures prévisionnelles : on limite aux collectifs et tiers}
      if ckReglement.Checked then
        InSel := ' (G_NATUREGENE IN ("COD", "COC", "COF", "COS", "TIC", "TID") OR (' + WhBqe + '))'
      else
        InSel := ' (G_NATUREGENE IN ("COD", "COC", "COF", "COS", "TIC", "TID"))';

      Result := Result + InSel
    end;
  end
  else
    Result := Result + WhBqe;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRSYNCHRONISATION.ckReglementClick;
{---------------------------------------------------------------------------------------}
begin
  SetControlEnabled('CKCOMPTES', ckReglement.Checked);
  SetControlEnabled('CBGENERAL', ckReglement.Checked and (GetCheckBoxState('CKCOMPTES') =  cbChecked));
end;

{FQ 10239 : s'assure de la cohérence des dates avec la date de première synchronisation
{---------------------------------------------------------------------------------------}
procedure TOF_TRSYNCHRONISATION.GestionDates(var dtMin, dtMax: TDateTime);
{---------------------------------------------------------------------------------------}
begin
  dtMin := StrToDate(edDateDe.Text);
  dtMax := StrToDate(edDateA.Text);

  if dtMin > dtMax then begin
    dtMin := dtMax;
    dtMax := StrToDate(edDateDe.Text);
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRSYNCHRONISATION.InitDt1ereSyncho;
{---------------------------------------------------------------------------------------}
begin
  Dt1ereSyncho := iDate1900;

  if VarToStr(GetParamSocSecur('SO_TRDATEPREMSYNCHRO', '')) = '' then Exit;
  if VarToDateTime(GetParamSocSecur('SO_TRDATEPREMSYNCHRO', 0)) <= iDate1900 then Exit;

  Dt1ereSyncho := VarToDateTime(GetParamSocSecur('SO_TRDATEPREMSYNCHRO', idate1900));
end;

{08/08/06 : synchronisation par dossier dans le cadre du multi sociétés
{---------------------------------------------------------------------------------------}
procedure TOF_TRSYNCHRONISATION.SynchroniseUnDossier(NomBase : string);
{---------------------------------------------------------------------------------------}
const
  Ws = ' AND (E_QUALIFORIGINE <> "' + QUALIFTRESO + '" OR E_QUALIFORIGINE IS NULL) '; {18/05/04 : ajout du IS NULL pour Oracle}
  Wa = ' AND E_QUALIFORIGINE = "' + QUALIFTRESO + '" ';
  {28/02/05 : Ajout de la gestion des banques prévisionnelles en lien avec les modifications de CCMP}
  Wh = ' E_TRESOSYNCHRO IN ("' + ets_Pointe + '", "' + ets_Lettre + '", "' + ets_Nouveau + '", "' + ets_BqPrevi + '") ' +
       {13/09/06 : Exclusion du journal d'a-nouveaux}
       'AND NOT (J_NATUREJAL IN ("ANO", "CLO")) ' +
       {17/04/07 : Ajout du E_QUALIFPIECE}
       'AND E_QUALIFPIECE = "N" ';
var
  Q : TQuery;
  T : TOB;
  R : TOB;
  S : string;
  w : string;
  c : string;
  G : TGrpObjet;
  VisuErrPaie : Boolean;
  CalcTotSold : Boolean;
begin
  if IsTresoMultiSoc then begin
    {18/12/06 : Si nom base = '', c'est que l'on est sur une société autonome dans une regroupement multi-sociétés}
    if NomBase = '' then NomBase := V_PGI.SchemaName;
    G.NoDossier := GetInfosFromDossier('DOS_NOMBASE', NomBase, 'DOS_NODOSSIER');
    G.Societe   := GetInfosFromDossier('DOS_NOMBASE', NomBase, 'DOS_SOCIETE');
  end
  else begin
    G.NoDossier := V_PGI.NoDossier;
    G.Societe   := V_PGI.CodeSociete;
  end;

  G.NomBase := NomBase;

  S := '';
  c := ConstituerClauseWh(G.NoDossier);
  if c <> '' then w := c + ' AND ' + Wh + Ws
             else w := 'WHERE ' + Wh + Ws;

  S := 'SELECT E_NUMLIGNE, E_NUMECHE, E_MODEPAIE, E_QUALIFPIECE, E_DATEVALEUR, E_NATURETRESO, ' + {29/05/07 : E_NATURETRESO}
       'E_JOURNAL, E_NUMEROPIECE, E_EXERCICE, E_DATECOMPTABLE, E_GENERAL, E_CONTREPARTIEGEN, ' +
       'E_REFINTERNE, E_SOCIETE, E_ETABLISSEMENT, E_LIBELLE, E_TRESOSYNCHRO, E_DEVISE, ' +
       'E_DEBIT, E_CREDIT, E_COUVERTURE, E_DEBITDEV, E_CREDITDEV, E_COUVERTUREDEV, E_COTATION, ' +
       'E_DATEPOINTAGE, E_BANQUEPREVI, G_NATUREGENE, E_DATEECHEANCE, G_EFFET, J_NATUREJAL, ' +
       'E_PERIODE, J_MODESAISIE, ' + {11/10/06 : FQ 10365 : pour les pièces bordereau}
       'E_REFPOINTAGE ' + {03/04/07 : Nouveau pointage : ajout du champ TE_REFPOINTAGE dans TRECRITURE}
       'FROM ' + GetTableDossier(NomBase, 'ECRITURE') + ' ' +
       'LEFT JOIN JOURNAL ON J_JOURNAL = E_JOURNAL ' +
       'LEFT JOIN GENERAUX ON G_GENERAL = E_GENERAL ' + w;

  G.oBqPrev := ObjParam;
  Q := OpenSQL(S, True);
  T := TOB.Create('****', nil, -1);
  R := TOB.Create('****', nil, -1);
  try
    T.LoadDetailDB('****', '', '', Q, False);
    Ferme(Q);
    {03/04/07 : Ajout de la référence de pointage}
    R.LoadDetailFromSQL('SELECT E_DATEPOINTAGE, E_NUMLIGNE, E_NUMEROPIECE, E_EXERCICE, E_GENERAL, ' +
                        'E_NUMECHE, E_JOURNAL, E_DATECOMPTABLE, E_REFINTERNE, E_REFPOINTAGE ' +
                        'FROM ' + GetTableDossier(NomBase, 'ECRITURE') + ' ' +
                        ClauseWhere + ' E_TRESOSYNCHRO = "' + ets_Pointe + '" ' + wa, True);
    {22/07/05 : FQ 10268 : Ajout de deux options de traitement}
    VisuErrPaie := GetCheckBoxState('CKMODEPAIE'    ) = cbChecked;
    CalcTotSold := GetCheckBoxState('CKSOLDECOMPLET') = cbChecked;

    G.TobRejet  := TobRejets;
    G.TobDate   := TobDate;
    G.ListSolde := lCalSolde;

    {Synchronisation des écritures}
    NbEcrit := NbEcrit + Synchronisation(T, R, G, VisuErrPaie, CalcTotSold);
  finally
    FreeAndNil(T);
    FreeAndNil(R);
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRSYNCHRONISATION.ckBancaireClick(Sender : TOBject);
{---------------------------------------------------------------------------------------}
begin
  PoseFiltreSurCpte;
end;

{Mise à jour de TobRejets et lCalSolde avec les données du dossier courant
{---------------------------------------------------------------------------------------}
procedure TOF_TRSYNCHRONISATION.TermineSynchronisation;
{---------------------------------------------------------------------------------------}
begin
  {Recalcul complet des soldes uniquement si demandé}
  if (NbEcrit > 0) and (GetCheckBoxState('CKSOLDECOMPLET') = cbChecked) then
    {... on recalcule les soldes sur les comptes concernés}
    MultiRecalculSolde(lCalSolde);

  if not V_PGI.SilentMode then
    MessageFinal;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRSYNCHRONISATION.MessageFinal;
{---------------------------------------------------------------------------------------}
var
  s : string;
begin
  {Constitution du message récapitulatif}
  if NbEcrit = 0 then
    s := TraduireMemoire('Aucune écriture n''a été traitée.')
  else if NbEcrit = 1 then
    {FQ 10152 : la précédente phrase laissait penser qu'il y avait p nouvelles dans la trésorerie : ce n'est pas le
                cas, il peut s'agir d'une simple mise à jour (qui peut être à blanc si les champs modifiés en compta
                ne concerne pas la Trésorerie : lorsque l'on modifie en saisie, on recrée l'écriture et on passe
                donc E_TRESOSYNCHRO à "CRE")}
    s := TraduireMemoire('Une écriture a été traitée.')
  else
    s := IntToStr(NbEcrit) + TraduireMemoire(' écritures ont été traitées.');

  {20/11/06 : gestion des dates antérieures à l'année en cours}  
  if TobDate.Detail.Count > 0 then begin
    TheTob := TobDate;
    HShowMessage('0;Synchronisation;Certaines écritures sont d''une année antérieure à l''année en cours :'#13 +
                  'Cela va impacter les soldes de réinitialisation s''ils ont été saisis.;W;O;O;O;', '', '');
    TRAfficheGrille(TraduireMemoire('Liste des écritures antérieures à l''année courante.'), ag_DateSynchro);
  end;

  {20/09/04 : FQ 10144 : l'ancien système d'affichage des rejets par un showmessage ne s'avérait pas très
              "ergonomique" s'il y avait beaucoup d'écritures non traitées}
  if TobRejets.Detail.Count > 0 then begin
    TheTob := TobRejets;
    s := s + #13#13 + 'Cependant certaines écritures n''ont pas été synchronisées :'#13 +
                      'souhaitez-vous en voir le détail ?';
    if HShowMessage('0;Synchronisation;' + s + ';Q;YN;Y;Y;', '', '') = mrYes then
      TRAfficheGrille(TraduireMemoire('Liste des écritures non synchronisées en raison d''un problème de paramétrage.'), ag_Synchro);
  end
  else
    HShowMessage('0;Synchronisation;' + s + ';Q;O;O;O;', '', '');
end;


{---------------------------------------------------------------------------------------}
function TOF_TRSYNCHRONISATION.GetWhereGene(NoDos : string; ConfOnly : Boolean): string;
{---------------------------------------------------------------------------------------}
var
  Q : Tquery;
  s : string;
  c : string;
begin
  Result := '';
  c := '';
  {$IFDEF TRCONF}
  c := TObjConfidentialite.GetWhereConf(V_PGI.User, tyc_Banque);
  {$ENDIF TRCONF}
  if (c = '') and ConfOnly then Exit;

  if c <> '' then C := ' AND (' + c + ') ';

  s := GetClauseIn(cbGeneral.Value);
  {01/02/07 : cela me parait plus prudent !!}
  if s = '' then
    Q := OpenSQL('SELECT BQ_GENERAL FROM BANQUECP WHERE BQ_NODOSSIER = "' + NoDos + '"' + c, True)
  else
    Q := OpenSQL('SELECT BQ_GENERAL FROM BANQUECP WHERE BQ_CODE IN (' + s + ') AND BQ_NODOSSIER = "' + NoDos + '"' + c, True);
  try
    s := '';
    {08/08/07 : S'il n'y a pas d'enregistrements, c'est que tous les comptes sont filtrés => il n'y donc aucune
                écritures bancaires à synchroniser, mais en plus comme l'utilisateur n'a accès à aucun compte
                du dossier, il ne faut pas qu'il puisse synchroniser des mouvements prévisionnels sur les comptes.
                Pour être sûr de l'exclusion, on crée un code général bidon}
    if Q.EOF then
      s := '"A-AAAAA@A"'
    else begin
      while not Q.EOF do begin
        s := s + '"' + Q.FindField('BQ_GENERAL').AsString + '",';
        Q.Next;
      end;
      if s <> '' then
        System.Delete(s, Length(s), 1);
    end;
    Result := s;
  finally
    Ferme(Q);
  end;
end;

initialization
  RegisterClasses([TOF_TRSYNCHRONISATION]);

end.


