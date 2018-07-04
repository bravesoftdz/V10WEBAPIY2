{-------------------------------------------------------------------------------------
    Version   |   Date | Qui | Commentaires
--------------------------------------------------------------------------------------
               28/03/02  BT   Création de la fiche
 6.xx.xxx.xxx  19/07/04  JP   Gestion des commissions de mouvements
 6.0X.001.001  19/10/04  JP   Gestion du plafond pour la CPFD
                              Gestion du type de calcul : par niveau / par tranche
 6.50.001.004  15/06/05  JP   FQ 10272 : le calcul du taux par tranche est érroné
 6.50.001.020  26/09/05  JP   - FQ 10284 : Maintenant on affiche les dates de traitement
                              - Correction du calcul (cf fonction Standard) ou l'on ne
                                prenait pas en compte le dernier jour de la période cf ##DERNIER##.
                                Par contre je ne calcule pas d'agios sur les opérations du
                                dernier jour, ceux-ci doivent figurer sur le premier jour du
                                trimestre suivant, du moins me semble-t-il (?)
 7.00.001.001  28/12/05  JP   FQ 10314 : Echelles d'intérêts multi-comptes
                              ##SOLDE## : Les soldes détailés et cumulés ne donnaient pas le même résultat
                              que le standard, car le dernier jour de la période n'était pas comptabilisé.
                              A priori, il faudrait décalé les agios d'une ligne vers le Bas et faire figurer
                              sur le jour les agios entre la dernière opération et la dite opération et non
                              les agios qui courent à partir de la date : ce qui annule mon commentaire (##DERNIER##)
                              Dans l'immédiat, il suffit de décalé les dates de sélection d'un jour en arrière
 7.06.001.001  24/08/06  JP   Exclusion des comptes courants et titre du traitement
 7.06.001.001  23/10/06  JP   Filtre sur les dossiers du regroupement Tréso
 7.06.001.001  09/11/06  JP   FQ 10256 : déplacement des réinitialisation au 01/01 au matin
 8.00.001.013  02/05/07  JP   Plantage avec Oracle sur la requête sur BANQUECP
 8.00.001.019  10/06/07  JP   FQ 10432 : Affichage du détail des opérations d'une ligne
 8.10.001.004  08/08/07  JP   Gestion des confidentialités
--------------------------------------------------------------------------------------}
unit TofEchellesInterets ;

interface

uses {$IFDEF VER150} variants,{$ENDIF}
  StdCtrls, Controls, Classes, UProcGen, Math, ParamSoc, UObjGen,
  {$IFDEF EAGLCLIENT}
  eMul, MaineAGL, UtileAGL,
  {$ELSE}
  PrintDBG, FE_Main, Mul, {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} db, EDtREtat,
  {$ENDIF}
  ExtCtrls, Forms, sysutils, HCtrls, HEnt1, UTOB, UTOF, Graphics, Grids;

type
  TOF_ECHELLESINTERETS = class (TOF)
    procedure OnUpdate              ; override;
    procedure OnLoad                ; override;
    procedure OnClose               ; override;
    procedure OnArgument(S : string); override;
  private
    {28/12/05 : FQ 10314 : Compte de référence dans le cas d'un traitement multi-comptes}
    FGeneral : string;
    function GetGeneral : string;
  protected
    TobListe,
    TobGrid   : TOB;
    Grid      : THGrid;
    lComMvt   : TStringList;
    MntComMvt : Double;
    TxComMVT  : Double;
    edGeneral : THedit;
    typePre   : THValComboBox;
    edDateDe  : THEdit;
    edDateA   : THEdit; {FQ 10284}
    bAgrandir,
    bReduire  : TNotifyEvent;
    FormatDevise: string;
    SoldeEstime : Double;
    ObjTaux     : TObjTaux;

    {28/12/05 : FQ 10314 : Nouveaux composants dans le cadre des échelles multi-comptes}
    ckFusion : TCheckBox;
    rgFusion : TRadioGroup;
    cbBanque : THValComboBox;
    cbAgence : THValComboBox;
    cbCompte : THMultiValComboBox;

    procedure CalculComMVT(Montant : Double; Flux : string);
    function  GetTxComMvt : Double;
    function  CalculPFD (TobCond : TOB) : Double;
    procedure Cumulee   (TobCond : TOB; DateFin : TDateTime; SoldePrec : Double);
    procedure Detaillee (TobCond : TOB; DateFin : TDateTime; SoldePrec : Double);
    procedure Standard  (TobCond : TOB; DateFin : TDateTime; SoldePrec : Double);
    function  CalculTaux(TobCond : TOB; NbJ : Double; var CurDate : TDateTime; var Nb, Tx : Double) : Double;
    {28/12/05 : FQ 10314 : Nouvelles méthodes dans le cadre des échelles multi-comptes}
    function  ConstruitRequete(var FiltreeOk : Boolean; var cNat : string) : string;
    function  GetWhereBancaire : string;
    function  TesteSaisie : Boolean; {Vérifie si les critères sont bons}
    function  GetSoldeGeneral(DateDe, Nature : string) : Double; {Récupération du solde de départ}
    function  CalculSolde(ClauseNat : string) : Double; {Recalcul des soldes quand on ne peut travailler sur ceux de la requête}
    procedure GestionDevise; {Gestion de la devise d'affichage et du nombre de décimales}
    procedure GestionInfos; {Affichage de la source des informations}
    function  GetInfosBanques : TQuery; {Charge BanqueCP pour le Ticket d'agios}

    procedure GetCellCanvas  (ACol, ARow: LongInt; Canvas: TCanvas; AState: TGridDrawState);
    procedure GeneralOnChange(Sender : TObject);
    procedure CritereClick   (Sender : TObject);
    procedure bAgrandirClick (Sender : TObject);
    procedure bReduireClick  (Sender : TObject);
    procedure bImprimerClick (Sender : TObject);
    procedure bExportClick   (Sender : TObject);
    procedure bAgiosClick    (Sender : TObject);
    procedure bConditionClick(Sender : TObject);
    procedure TyEchelleChange(Sender : TObject);
    {28/12/05 : FQ 10314 : Nouveaux évènements dans le cadre des échelles multi-comptes}
    procedure ckFusionClick  (Sender : TObject);
    procedure rgFusionClick  (Sender : TObject);
    {10/06/07 : FQ 10432 : affichage du détail des opération}
    procedure FListeDblClick (Sender : TObject);

    {$IFDEF EAGLCLIENT}
    procedure bChercheClick  (Sender: TObject);
    {$ENDIF}
    {28/12/05 : FQ 10314 : Compte de référence dans le cas d'un traitement multi-comptes}
    property General : string read GetGeneral;
  end ;

procedure TRLanceFiche_EchellesInterets(Dom: string; Fiche: string; Range: string; Lequel: string; Arguments: string);

implementation

uses
  {$IFDEF TRCONF}
  ULibConfidentialite,
  {$ENDIF TRCONF}
  Commun, HTB97, HMsgBox, HXlsPas, HPanel, TomConditionDec,
  Constantes, TofDetailFlux, AglInit, UProcSolde;

const
  COL_TYPE = 0; // Colonnes dans la grille

{---------------------------------------------------------------------------------------}
procedure TRLanceFiche_EchellesInterets(Dom: string; Fiche: string; Range: string; Lequel: string; Arguments: string);
{---------------------------------------------------------------------------------------}
begin
  AglLanceFiche(Dom, Fiche, Range, Lequel, Arguments );
end ;

{---------------------------------------------------------------------------------------}
procedure TOF_ECHELLESINTERETS.OnArgument (S : string ) ;
{---------------------------------------------------------------------------------------}
var
  Q : TQuery;
  W : string;{02/05/07 : Plantage Oracle}
begin
  Inherited ;

  Ecran.HelpContext := 150;
  {14/06/04 : La tob n'est plus (re)créer dans le OnUpdate afin de pouvoir l'utiliser en impression en eAgl}
  TobGrid := Tob.Create('', nil, -1);
  {Création de l'objet contenant les cours des taux}
  ObjTaux := TObjTaux.Create(StrToDate(DebutTrimestre(V_PGI.DateEntree)),StrToDate(FinTrimestre(V_PGI.DateEntree)), True) ;

  edDateDe := THEdit(GetControl('DATEDE'));
  edDateA  := THEdit(GetControl('DATEDE_')); {FQ 10284}

  edGeneral := THEdit(GetControl('EDGENERAL'));
  typePre := THValComboBox(GetControl('TYPE')) ;
  bAgrandir := TToolbarButton97(GetControl('BAGRANDIR')).OnClick;
  bReduire := TToolbarButton97(GetControl('BREDUIRE')).OnClick;
  TToolbarButton97(GetControl('BAGRANDIR'   )).OnClick := bAgrandirClick;
  TToolbarButton97(GetControl('BREDUIRE'    )).OnClick := bReduireClick;
  TToolbarButton97(GetControl('BIMPRIMER'   )).OnClick := BImprimerClick;
  TToolbarButton97(GetControl('BEXPORT'     )).OnClick := bExportClick;
  TToolbarButton97(GetControl('BTICKETAGIOS')).OnClick := bAgiosClick;
  TToolbarButton97(GetControl('BCONDITIONS' )).OnClick := bConditionClick;

  {$IFDEF EAGLCLIENT}
  TToolbarButton97(GetControl('BCHERCHE')).OnClick := bChercheClick;
  SetControlVisible('BPREV', False);
  SetControlVisible('BNEXT', False);
  {$ENDIF}
  Grid := THGrid(GetControl('GRID'));
  Grid.GetCellCanvas := GetCellCanvas;
  Grid.OnDblClick := FListeDblClick; {10/06/07 : FQ 10432}

  {26/09/05 : FQ 10284 : On affiche les dates du trimestre}
  edDateDe.Text := DebutTrimestre(Now);
  edDateA .Text := FinTrimestre(Now);

  {23/10/06 : On limite l'affichage aux comptes courants des dossiers du regroupement Tréso}
  THEdit(GetControl('EDGENERAL')).Plus := FiltreBanqueCp(tcp_Bancaire, '', '');

  {04/10/06 : Initialisation de EDGENERAL}
  w := FiltreBanqueCp(tcp_Bancaire, tcb_Bancaire, '');
  if w <> '' then w := ' WHERE ' + w;
  Q := OpenSQL('SELECT ##TOP 1## BQ_CODE FROM BANQUECP' + w, True);
  if not Q.EOF then
    SetControlText('EDGENERAL', Q.Fields[0].AsString);
  Ferme(Q);

  GeneralOnChange(edGeneral);
  edGeneral.OnChange := GeneralOnChange;
  TypePre.OnChange   := TyEchelleChange;
  edDateDe.OnChange  := CritereClick;
  edDateA .OnChange  := CritereClick; {FQ 10284}

  {19/07/04 : chargement des codes flux/rubriques soumis à commission de mouvements}
  lComMvt := TStringList.Create;
  lComMvt.Duplicates := dupIgnore;
  Q := OpenSQL('SELECT TFT_FLUX FROM TRLISTEFLUX WHERE TFT_COMMOUVEMENT = "X"', True);
  try
    while not Q.EOF do begin
      lComMvt.Add(Q.FindField('TFT_FLUX').AsString);
      Q.Next;
    end;
  finally
    Ferme(Q);
  end;

  {28/12/05 : FQ 10314 : Nouveaux composants dans le cadre des échelles multi-comptes}
  ckFusion := TCheckBox         (GetControl('CKFUSION'));
  rgFusion := TRadioGroup       (GetControl('RGFUSION'));
  cbBanque := THValComboBox     (GetControl('CBBANQUE'));
  cbAgence := THValComboBox     (GetControl('CBAGENCE'));
  cbCompte := THMultiValComboBox(GetControl('CBCOMPTE'));

  {23/10/06 : Gestion du multi sociétés}
  cbCompte.Plus := FiltreBanqueCp(tcp_Bancaire, '', '');

  ckFusion.OnClick := ckFusionClick;
  rgFusion.OnClick := rgFusionClick;

  TobListe := nil; {Pour éviter toute désagréable surprise lorsque l'on teste "l'assignation" de la tob}
  FGeneral := '';
  {24/08/06 : exclusion des comptes titres et courants}
  SetPlusBancaire(cbAgence , 'TRA', CODECOURANTS + ';' + CODETITRES + ';');
  SetPlusBancaire(cbBanque , 'PQ' , CODECOURANTS + ';' + CODETITRES + ';');
end;

{---------------------------------------------------------------------------------------}
procedure TOF_ECHELLESINTERETS.bAgrandirClick(Sender: TObject);
{---------------------------------------------------------------------------------------}
begin
  THPanel(GetControl('BOTTOMPANEL')).Visible := False;
  bAgrandir(Sender);
end;

{---------------------------------------------------------------------------------------}
procedure TOF_ECHELLESINTERETS.bReduireClick(Sender: TObject);
{---------------------------------------------------------------------------------------}
begin
  THPanel(GetControl('BOTTOMPANEL')).Visible := True;
  bReduire(Sender);
end;

{---------------------------------------------------------------------------------------}
procedure TOF_ECHELLESINTERETS.BImprimerClick(Sender: TObject);
{---------------------------------------------------------------------------------------}
{$IFDEF EAGLCLIENT}
var
  T : TOB;
  n : Integer;
{$ENDIF}
begin
{$IFDEF EAGLCLIENT}
  {15/06/04 : Gestion des états en eAgl}
  {Suppression de la ligne total}
  T := TobGrid.Detail[TobGrid.Detail.Count - 1];
  if Assigned(T) then FreeAndNil(T);

  {Mise au format de l'état (facultatif !?)}
  for n := 0 to TobGrid.Detail.Count - 1 do begin
    T := TobGrid.Detail[n];
    if T.GetString('INTDEB') = '' then T.SetDouble('INTDEB', 0.00);
  end;

  {Pour simplifier les choses, j'ai fais trois états}
  case TypePre.ItemIndex of
    0 : LanceEtatTOB('E', 'ECT', 'ECC', TobGrid, True, False, False, nil, '', Ecran.Caption, False);
    1 : LanceEtatTOB('E', 'ECT', 'ECD', TobGrid, True, False, False, nil, '', Ecran.Caption, False);
    2 : LanceEtatTOB('E', 'ECT', 'ECS', TobGrid, True, False, False, nil, '', Ecran.Caption, False);
  end ;
{$ELSE}
  PrintDBGrid(Grid, nil, TForm(Ecran).Caption + ' (' + TypePre.Text + ')', '');
{$ENDIF}
end;

{Imprime la grille rajoutée au lieu de la DB grille
{---------------------------------------------------------------------------------------}
procedure TOF_ECHELLESINTERETS.bExportClick(Sender: TObject);
{---------------------------------------------------------------------------------------}
begin
  if ExJaiLeDroitConcept(ccExportListe, True) and TFMul(Ecran).SD.Execute then
  ExportGrid(Grid, Nil, TFMul(Ecran).SD.FileName, TFMul(Ecran).SD.FilterIndex, True);
end;

{---------------------------------------------------------------------------------------}
procedure TOF_ECHELLESINTERETS.bAgiosClick(Sender: TObject);
{---------------------------------------------------------------------------------------}
var
  TobA : TOB;
  T    : TOB;
  Q    : TQuery;
begin
  T := TOB.Create('$TISKETAGIOS', nil, -1);
  try
    Q := GetInfosBanques; {28/12/05 : Charges les comptes sélectionnés}
    while not Q.EOF do begin
      TobA := TOB.Create('$TISKETAGIOS', T, -1);
      TobA.AddChampSupValeur('AG_COMPFD'  , Abs(THNumEdit(GetControl('COMPFD'  )).Value));
      TobA.AddChampSupValeur('AG_AGIOS'   , Abs(THNumEdit(GetControl('AGIOS'   )).Value));
      TobA.AddChampSupValeur('AG_MPFD'    , Abs(THNumEdit(GetControl('MPFD'    )).Value));
      TobA.AddChampSupValeur('AG_INTDEBIT', Abs(THNumEdit(GetControl('INTDEBIT')).Value));
      TobA.AddChampSupValeur('AG_COMMVT'  , Abs(THNumEdit(GetControl('COMMVT'  )).Value));

      {26/09/05 : FQ 10284 : Maintenant les dates peuvent être saisies et donc on n'est pas
                  nécessairement sur un trimestre civil}
      TobA.AddChampSupValeur('AG_DTDEB'   , edDateDe.Text);
      TobA.AddChampSupValeur('AG_DTFIN'   , edDateA .Text);

      TobA.AddChampSupValeur('AG_DEVISE'  , RechDom('TTDEVISE', GetControlText('DEV'), False));
      TobA.AddChampSupValeur('AG_DEV'     , GetControlText('DEV'));
      TobA.AddChampSupValeur('AG_COMPTE'  , Q.Fields[0].AsString);
      TobA.AddChampSupValeur('AG_ADR1'    , Q.Fields[1].AsString);
      TobA.AddChampSupValeur('AG_ADR2'    , Q.Fields[2].AsString);
      TobA.AddChampSupValeur('AG_CP'      , Q.Fields[3].AsString);
      TobA.AddChampSupValeur('AG_VILLE'   , Q.Fields[4].AsString);
      TobA.AddChampSupValeur('AG_PAYS'    , RechDom('TTPAYS', Q.Fields[5].AsString, False));
      TobA.AddChampSupValeur('AG_BIC'     , Q.Fields[6].AsString);
      if Copy(Q.Fields[5].AsString, 1, 2) = 'FR' then
        TobA.AddChampSupValeur('AG_RIB'   , Q.Fields[11].AsString)
      else
        TobA.AddChampSupValeur('AG_RIB'   , Q.Fields[7].AsString + ' - ' +
                                            Q.Fields[8].AsString + ' - ' +
                                            Q.Fields[9].AsString + ' - ' +
                                            Q.Fields[10].AsString);
      TobA.AddChampSupValeur('AG_LIBELLE' , GetParamSocSecur('SO_LIBELLE'   , ''));
      TobA.AddChampSupValeur('AG_SOADR1'  , GetParamSocSecur('SO_ADRESSE1'  , ''));
      TobA.AddChampSupValeur('AG_SOADR2'  , GetParamSocSecur('SO_ADRESSE2'  , ''));
      TobA.AddChampSupValeur('AG_SOCP'    , GetParamSocSecur('SO_CODEPOSTAL', ''));
      TobA.AddChampSupValeur('AG_SOVILLE' , GetParamSocSecur('SO_VILLE'     , ''));
      TobA.AddChampSupValeur('AG_SOPAYS'  , RechDom('TTPAYS', GetParamSocSecur('SO_PAYS', 'FRA'), False));
      TobA.AddChampSupValeur('AG_SOTEL'   , GetParamSocSecur('SO_TELEPHONE', ''));
      TobA.AddChampSupValeur('AG_SOLDE'   , SoldeEstime);
      Q.Next;
    end;
    Ferme(Q);
    if ckFusion.Checked then LanceEtatTOB('E', 'ECT', 'TAM', T, True, False, False, nil, '', 'Tickets d''agios', False)
                        else LanceEtatTOB('E', 'ECT', 'TAG', T, True, False, False, nil, '', 'Tickets d''agios', False);
  finally
    FreeAndNil(T);
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_ECHELLESINTERETS.bConditionClick(Sender : TObject);
{---------------------------------------------------------------------------------------}
begin
  TRLanceFiche_ConditionDec('TR','TRCONDITIONDEC', '', 'DEC;' + General, 'ACTION=MODIFICATION');
  {Pour le raffraîchissement de l'écran}
  TFMul(Ecran).BCherche.Click;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_ECHELLESINTERETS.GetCellCanvas(ACol,ARow: LongInt; Canvas: TCanvas; AState: TGridDrawState);
{---------------------------------------------------------------------------------------}
var
  S: String;
begin
  if not (gdSelected in AState) then
  begin
    S := Grid.Cells[COL_TYPE, ARow];
    if (S = '+') or (S = '-') then // Sur total
    begin
      Canvas.Font.Style := [fsBold];
      if V_PGI.NumAltCol=0 then Canvas.Brush.Color := clInfoBk // Grid.AlternateColor ?
		           else Canvas.Brush.Color := AltColors[V_PGI.NumAltCol];
    end ;
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_ECHELLESINTERETS.GeneralOnChange(Sender: TObject);
{---------------------------------------------------------------------------------------}
begin
  {28/12/05 : Gestion du format des devises et du nombres de décimales d'affichages}
  GestionDevise;
  CritereClick(Sender);
end;

{---------------------------------------------------------------------------------------}
procedure TOF_ECHELLESINTERETS.CritereClick(Sender: TObject);
{---------------------------------------------------------------------------------------}
begin
  if V_PGI.AutoSearch then // Si MulAutoSearch...
    TFMul(Ecran).ChercheClick;
end;

{***********A.G.L.Privé.*****************************************
Auteur  ...... : Bruno TREDEZ
Créé le ...... : 29/03/2002
Modifié le ... :   /  /
Description .. : Contruit une TOB pour la grille, en fonction du type
Suite ........ : d'affichage (Standard, cumulé ou détaillé)
Mots clefs ... : GRID;AFFICHAGE
*****************************************************************}
procedure TOF_ECHELLESINTERETS.Standard(TobCond : TOB; DateFin : TDateTime; SoldePrec : Double);
var
  TobG     : TOB;
  N, J     : Integer;
  TotalDeb,
  TotalCre : Double;
  Taux     : Double;
  Base     : Integer;
  NbJours,
  Nombre   : Double;
  CurDate  : TDateTime;
  MntBrut,
  Solde,
  Int      : Double;

    {-------------------------------------------------------------------}
    procedure CreeTobDetail(I: Integer);
    {-------------------------------------------------------------------}
    var
      n : Integer;
      p : Integer;{##DERNIER##}
    begin
      {Récupération du solde en Valeur}
      Solde   := TobListe.Detail[I].GetDouble('TE_SOLDEDEVVALEUR');
      CurDate := TobListe.Detail[I].GetDouble('TE_DATEVALEUR');
      {Récupération du nombre de jours entre l'opération en cours et l'opération suivante}
      if I = TobListe.Detail.Count - 1 then begin
        NbJours := DateFin - CurDate;
        {##DERNIER## : avant, on faisait -1, ce qui fait que le dernier jour n'était pas traité ...}
        p := Trunc(NbJours);
      end
      else begin
        NbJours := TobListe.Detail[I + 1].GetDateTime('TE_DATEVALEUR') - CurDate;
        {##DERNIER## : ... par contre, ici il est normal de faire - 1 car à NbJours on est
                       sur la date de valeur de l'écriture suivante}
        p := Trunc(NbJours) - 1;
      end;

      {Calcul SoldeJours}
      Nombre  := {NbJours * }Solde;

      {19/07/04 : Calcul des commissions de mouvements}
      CalculComMVT(TobListe.Detail[I].GetDouble('TE_MONTANTDEV'), TobListe.Detail[I].GetString('TE_CODEFLUX'));

      {Calcul du taux}
      MntBrut := CalculTaux(TobCond, 1, DateFin, Nombre, Taux);

      Int := MntBrut{Nombre * Taux} / (100 * CalcNbJourParAnBase(CurDate, Base));
      {Remplissage jusqu'au flux suivant}
      for n := 0 to p {Trunc(NbJours) - 1} do begin {##DERNIER##}
        TobG := TOB.Create('', TobGrid, -1);
        TobG.AddChampSupValeur('COL_TYPE', '1');
        TobG.AddChampSupValeur('DATEVALEUR', DateToStr(CurDate));
        TobG.AddChampSupValeur('SOLDE', Solde);
        TobG.AddChampSupValeur('NOMBRE', Nombre);
        {On affiche le taux que s'il y a des intérêts débiteurs}
        if Nombre < 0 then TobG.AddChampSupValeur('TAUX', FormatFloat('0.0000', Taux))
                      else TobG.AddChampSupValeur('TAUX', '');
        TobG.AddChampSupValeur('INTDEB', '');
        TobG.AddChampSupValeur('INTCRE', '');

        {Calcul des cumuls et maj des totaux}
        if Int < 0 then begin
          TobG.PutValue('INTDEB', Int);
          TotalDeb := TotalDeb + Int;
        end
        else if Int > 0 then begin
          TobG.PutValue('INTCRE', Int);
          TotalCre := TotalCre + Int;
        end;
        CurDate := CurDate + 1;
      end;
    end;

begin
  TotalDeb := 0;
  TotalCre := 0;
  Taux := Valeur(TobCond.GetString('TCN_MAJOTAUX1'));
  Base := ValeurI(TobCond.GetString('TCN_BASECALCUL'));

  CurDate := StrToDate(edDateDe.Text); {FQ 10284}

  {Création d'une première ligne si la date de première opération est supérieure à la date
   de début du trimestre}
  if SoldePrec <> -922337203685477 then begin
    NbJours := TobListe.Detail[0].GetDateTime('TE_DATEVALEUR') - CurDate;
    Nombre  := {NbJours * }SoldePrec;
    MntBrut := CalculTaux(TobCond, 1, DateFin, Nombre, Taux);

    Int := MntBrut {Nombre * Taux} / (100 * CalcNbJourParAnBase(CurDate, Base));

    {On crée des Tobs détail jusqu'à la date du premier flux}
    for J := 0 to Trunc(NbJours) - 1 do begin
      TobG := TOB.Create('', TobGrid, -1);
      TobG.AddChampSupValeur('COL_TYPE', '1');
      TobG.AddChampSupValeur('DATEVALEUR', DateToStr(CurDate));
      TobG.AddChampSupValeur('SOLDE', SoldePrec);
      TobG.AddChampSupValeur('NOMBRE', Nombre);
      {On affiche le taux que s'il y a des intérêts débiteurs}
      if Nombre < 0 then TobG.AddChampSupValeur('TAUX', FormatFloat('0.0000', Taux))
                    else TobG.AddChampSupValeur('TAUX', '');
      TobG.AddChampSupValeur('INTDEB', '');
      TobG.AddChampSupValeur('INTCRE', '');

      {Calcul des cumuls et maj des totaux}
      if Int < 0 then begin
        TobG.PutValue('INTDEB', Int);
        TotalDeb := TotalDeb + Int;
      end
      else if Int > 0 then begin
        TobG.PutValue('INTCRE', Int);
        TotalCre := TotalCre + Int;
      end;
      CurDate := CurDate + 1;
    end;
  end;

  for N:=0 to TobListe.Detail.Count - 1 do
    CreeTobDetail(N);

  {Ligne contenant les totaux}
  TobG := TOB.Create('', TobGrid, -1);
  TobG.AddChampSupValeur('COL_TYPE', '+');
  TobG.AddChampSupValeur('DATEVALEUR', '');
  TobG.AddChampSupValeur('SOLDE', '');
  TobG.AddChampSupValeur('NOMBRE', '');
  TobG.AddChampSupValeur('TAUX', 'Total');
  TobG.AddChampSupValeur('INTDEB', TotalDEB);
  TobG.AddChampSupValeur('INTCRE', TotalCRE);
end;

{---------------------------------------------------------------------------------------}
procedure TOF_ECHELLESINTERETS.Cumulee(TobCond : TOB; DateFin : TDateTime; SoldePrec : Double);
{---------------------------------------------------------------------------------------}
var
  TobG     : TOB;
  N        : Integer;
  TotalDeb,
  TotalCre : Double;
  Taux     : Double;
  Base     : Integer;
  NbJours,
  Nombre   : Double;
  CurDate  : TDateTime;
  MntBrut,
  Solde,
  Int      : Double;

    {-------------------------------------------------------------------}
    procedure CreeTobDetail(I: Integer);
    {-------------------------------------------------------------------}
    begin
      Solde   := TobListe.Detail[I].GetDouble('TE_SOLDEDEVVALEUR');
      CurDate := TobListe.Detail[I].GetDateTime('TE_DATEVALEUR');

      TobG := TOB.Create('', TobGrid, -1);
      TobG.AddChampSupValeur('COL_TYPE', '1');
      TobG.AddChampSupValeur('DATEVALEUR', DateToStr(CurDate));
      TobG.AddChampSupValeur('SOLDE', Solde);

      if I = TobListe.Detail.Count - 1 then NbJours := DateFin - CurDate + 1 {28/12/05 : ##SOLDE##}
                                       else NbJours := TobListe.Detail[I + 1].GetDateTime('TE_DATEVALEUR') - CurDate;

      TobG.AddChampSupValeur('NOMBREJ', NbJours);
      Nombre := NbJours * Solde;

      {19/07/04 : Calcul des commissions de mouvements}
      CalculComMVT(TobListe.Detail[I].GetDouble('TE_MONTANTDEV'), TobListe.Detail[I].GetString('TE_CODEFLUX'));

      {Calcul du taux}
      MntBrut := CalculTaux(TobCond, NbJours, DateFin, Nombre, Taux);

      TobG.AddChampSupValeur('NOMBRE', Nombre);
      {On affiche le taux que s'il y a des intérêts débiteurs}
      if Nombre < 0 then TobG.AddChampSupValeur('TAUX', FormatFloat('0.0000', Taux))
                    else TobG.AddChampSupValeur('TAUX', '');
      TobG.AddChampSupValeur('INTDEB', '');
      TobG.AddchampSupValeur('INTCRE', '');

      Int := (MntBrut{Nombre * Taux} / (100 * CalcNbJourParAnBase(CurDate, Base)));
      if Int < 0 then begin
        TobG.PutValue('INTDEB', Int);
        TotalDeb := TotalDeb + Int;
      end
      else if Int > 0 then begin
        TobG.PutValue('INTCRE', Int);
        TotalCre := TotalCre + Int;
      end;
    end;

begin
  TotalDeb := 0;
  TotalCre := 0;
  Taux := Valeur(TobCond.GetString('TCN_MAJOTAUX1'));
  Base := ValeurI(TobCond.GetString('TCN_BASECALCUL'));
  CurDate := TobListe.Detail[0].GetDateTime('TE_DATEVALEUR');

  {Création d'une première ligne si la date de première opération est supérieure à la date
   de début du trimestre}
  if SoldePrec <> -922337203685477 then begin
    TobG := TOB.Create('', TobGrid, -1);
    TobG.AddChampSupValeur('COL_TYPE', '1');
    TobG.AddChampSupValeur('DATEVALEUR', StrToDate(edDateDe.Text)); {FQ 10286}
    TobG.AddChampSupValeur('SOLDE', SoldePrec);
    NbJours := TobListe.Detail[0].GetDateTime('TE_DATEVALEUR') - TobG.GetDateTime('DATEVALEUR');
    TobG.AddChampSupValeur('NOMBREJ', NbJours);
    Nombre  := NbJours * SoldePrec;
    TobG.AddChampSupValeur('NOMBRE', Nombre);
    MntBrut := CalculTaux(TobCond, NbJours, DateFin, Nombre, Taux);
    {On affiche le taux que s'il y a des intérêts débiteurs}
    if Nombre < 0 then TobG.AddChampSupValeur('TAUX', FormatFloat('0.0000', Taux))
                  else TobG.AddChampSupValeur('TAUX', '');

    TobG.AddChampSupValeur('INTDEB', '');
    TobG.AddchampSupValeur('INTCRE', '');
    Int := MntBrut{Nombre * Taux} / (100 * CalcNbJourParAnBase(CurDate, Base));
    if Int < 0 then begin
      TobG.PutValue('INTDEB', Int);
      TotalDeb := TotalDeb + Int;
    end
    else if Int > 0 then begin
      TobG.PutValue('INTCRE', Int);
      TotalCre := TotalCre + Int;
    end;
  end;

  for N := 0 to TobListe.Detail.Count - 1 do
    CreeTobDetail(N);

  {Ligne contenant les totaux}
  TobG := TOB.Create('', TobGrid, -1);
  TobG.AddChampSupValeur('COL_TYPE', '+');
  TobG.AddChampSupValeur('DATEVALEUR', '');
  TobG.AddChampSupValeur('SOLDE', '');
  TobG.AddChampSupValeur('NOMBREJ', '');
  TobG.AddChampSupValeur('NOMBRE', '');
  TobG.AddChampSupValeur('TAUX', 'Total');
  TobG.AddChampSupValeur('INTDEB', TotalDeb);
  TobG.AddChampSupValeur('INTCRE', TotalCre);
end;

{Calcul des intérêts en mode détaillé : dans ce cas il n'y pas de regroupement par date
 de valeur
{---------------------------------------------------------------------------------------}
procedure TOF_ECHELLESINTERETS.Detaillee(TobCond : TOB; DateFin : TDateTime; SoldePrec : Double);
{---------------------------------------------------------------------------------------}
var
  TobG     : TOB;
  I        : Integer;
  Base     : Integer;
  CurDate,
  NextDate : TDateTime;
  First    : Boolean;
  Montant  : Double;
  NbJours,
  Nombre   : Double;
  TotalDeb,
  TotalCre : Double;
  Taux     : Double;
  MntBrut,
  Solde,
  Int      : Double;
begin
  TotalDeb := 0;
  TotalCre := 0;
  Montant  := 0;
  TobG := nil;
  Taux := Valeur(TobCond.GetString('TCN_MAJOTAUX1'));
  Base := ValeurI(TobCond.GetString('TCN_BASECALCUL'));

  First := True;
  CurDate := TobListe.Detail[0].GetDateTime('TE_DATEVALEUR');

  {Création d'une première ligne si la date de première opération est supérieure à la date
   de début du trimestre}
  if SoldePrec <> - 922337203685477 then begin
    TobG := TOB.Create('', TobGrid, -1);
    TobG.AddChampSupValeur('COL_TYPE', '1');
    TobG.AddChampSupValeur('DATEVALEUR', StrToDate(edDateDe.Text)); {FQ 10284}
    TobG.AddChampSupValeur('MOUVEMENT', '');
    TobG.AddChampSupValeur('SOLDE', SoldePrec);
    NbJours := TobListe.Detail[0].GetDateTime('TE_DATEVALEUR') - TobG.GetDateTime('DATEVALEUR');
    Nombre  := NbJours * SoldePrec;
    TobG.AddChampSupValeur('NOMBRE', Nombre);
    MntBrut := CalculTaux(TobCond, NbJours, DateFin, Nombre, Taux);
    {On affiche le taux que s'il y a des intérêts débiteurs}
    if Nombre < 0 then TobG.AddChampSupValeur('TAUX', FormatFloat('0.0000', Taux))
                  else TobG.AddChampSupValeur('TAUX', '');

    TobG.AddChampSupValeur('INTDEB', '');
    TobG.AddchampSupValeur('INTCRE', '');
    Int := MntBrut{Nombre * Taux} / (100 * CalcNbJourParAnBase(CurDate, Base));
    if Int < 0 then begin
      TobG.PutValue('INTDEB', Int);
      TotalDeb := TotalDeb + Int;
    end
    else if Int > 0 then begin
      TobG.PutValue('INTCRE', Int);
      TotalCre := TotalCre + Int;
    end;
  end;

  for I := 0 to TobListe.Detail.Count - 1 do begin
    {Si on est sur la dernière ligne, on prend DateFin de la requête}
    if I = TobListe.Detail.Count - 1 then NextDate := DateFin + 1 {28/12/05 ##SOLDE##}
                                     else NextDate := TobListe.Detail[I+1].GetDateTime('TE_DATEVALEUR');

    {Si c'est la première ligne pour une date, on récupère la date}
    if First then begin
      TobG := TOB.Create('', TobGrid, -1);
      TobG.AddChampSupValeur('COL_TYPE', '1');
      TobG.AddChampSupValeur('DATEVALEUR', DateToStr(CurDate));
      Montant := 0;
    end;

    {19/07/04 : Calcul des commissions de mouvements}
    CalculComMVT(TobListe.Detail[I].GetDouble('TE_MONTANTDEV'), TobListe.Detail[I].GetString('TE_CODEFLUX'));

    {Calcul des mouvements sur la journée}
    Montant := Montant + TobListe.Detail[I].GetDouble('TE_MONTANTDEV');

    {Si les dates sont différentes, c'est qu'il s'agit de la dernière ligne pour la date en cours}
    if CurDate <> NextDate then begin
      {Récupération du Montant en devise de l'opération}
      TobG.AddChampSupValeur('MOUVEMENT', Montant);

      {Récupération du solde}
      Solde := TobListe.Detail[I].GetDouble('TE_SOLDEDEVVALEUR');
      TobG.AddChampSupValeur('SOLDE', Solde);

      {Calcul de la base}
      NbJours := NextDate - CurDate;
      Nombre  := NbJours * Solde;
      TobG.AddChampSupValeur('NOMBRE', Nombre);

      MntBrut := CalculTaux(TobCond, NbJours, DateFin, Nombre, Taux);
      {On affiche le taux que s'il y a des intérêts débiteurs}
      if Nombre < 0 then TobG.AddChampSupValeur('TAUX', FormatFloat('0.0000', Taux))
                    else TobG.AddChampSupValeur('TAUX', '');

      TobG.AddChampSupValeur('INTDEB', '');
      TobG.AddchampSupValeur('INTCRE', '');

      Int := MntBrut{Nombre * Taux} / (100 * CalcNbJourParAnBase(CurDate, Base));
      if Int < 0 then begin
        TobG.PutValue('INTDEB', Int);
        TotalDeb := TotalDeb + Int;
      end
      else if Int > 0 then begin
        TobG.PutValue('INTCRE', Int);
        TotalCre := TotalCre + Int;
      end;
    end;

    First := CurDate <> NextDate;
    CurDate := NextDate;
  end;

  {Ligne contenant les totaux}
  TobG := TOB.Create('', TobGrid, -1);
  TobG.AddChampSupValeur('COL_TYPE', '+');
  TobG.AddChampSupValeur('DATEVALEUR', '');
  TobG.AddChampSupValeur('MOUVEMENT', '');
  TobG.AddChampSupValeur('SOLDE', '');
  TobG.AddChampSupValeur('NOMBRE', '');
  TobG.AddChampSupValeur('TAUX', 'Total');
  TobG.AddChampSupValeur('INTDEB', TotalDeb);
  TobG.AddChampSupValeur('INTCRE', TotalCre);
end;

{On récupère le taux applicable à la date de fin de trimestre
{---------------------------------------------------------------------------------------}
function TOF_ECHELLESINTERETS.CalculTaux(TobCond : TOB; NbJ : Double; var CurDate : TDateTime;
                                         var Nb, Tx : Double) : Double;
{---------------------------------------------------------------------------------------}

    {--------------------------------------------------------------------------}
    function RetourneTaux(Ind : Char) : Double;
    {--------------------------------------------------------------------------}
    begin
    {JP 30/04/04 : suite à un mail de Régis, je divise TCN_MULTAUX par 100
        03/05/04 : ajout de la valeur absolue "Abs(Nb)", sinon on prenait toujours le TCN_TAUXREF1 !!! FQ 10056}
      Result := ObjTaux.GetTaux(TobCond.GetString('TCN_TAUXREF' + Ind), DateToStr(CurDate));
      Result := Result * TobCond.GetDouble('TCN_MULTAUX' + Ind)/100 + TobCond.GetDouble('TCN_MAJOTAUX' + Ind);
    end;

var
  Coeff : ShortInt;
begin
  Result := 0;
  Coeff  := - 1;
  {19/10/04 : Gestion du calcul par tranche}
  if TobCond.GetString('TCN_TYPECALCFRAIS') = tcf_Tranche then begin
    {Pour gérer le sens du résultat, les TCN_PLAFOND étant, théoriquement, des montants positifs}
    if Nb > 0 then Coeff := 1;

    if Nb = 0 then begin
      Tx := RetourneTaux('1');
      Exit;
    end
    else if Abs(Nb) <= TobCond.GetDouble('TCN_PLAFOND1') * NbJ then
      {FQ 10272 : il n'y a aucune raison de multiplié par NbJ, c'est juste les plafonds qu'il faut multiplier
                  pour les ramener à la même base jour que Nb}
      Result := Nb * {NbJ *} RetourneTaux('1')

    else if Abs(Nb) <= TobCond.GetDouble('TCN_PLAFOND2') * NbJ then
      Result := TobCond.GetDouble('TCN_PLAFOND1') * NbJ * RetourneTaux('1') * Coeff +
                (Abs(Nb) - TobCond.GetDouble('TCN_PLAFOND1') * NbJ) * RetourneTaux('2') * Coeff

    else if Abs(Nb) > TobCond.GetDouble('TCN_PLAFOND2') * NbJ then
      Result := TobCond.GetDouble('TCN_PLAFOND1') * NbJ * RetourneTaux('1') * Coeff +
                {FQ 10272 : il faut appliquer Nbj à TobCond.GetDouble('TCN_PLAFOND2') aussi}
                ((TobCond.GetDouble('TCN_PLAFOND2') - TobCond.GetDouble('TCN_PLAFOND1')) * NbJ) * RetourneTaux('2') * Coeff +
                (Abs(Nb) - TobCond.GetDouble('TCN_PLAFOND2') * NbJ) * RetourneTaux('3') * Coeff;

    {On calcul le taux en fonction du résultat}
    Tx := Abs(Result / Nb);
  end
  else begin
    {Comme Nb est le produit du nombre de jours par le solde, on multiplie les plafonds par le Nombre de jours}
    if Abs(Nb) <= TobCond.GetDouble('TCN_PLAFOND1') * NbJ then
      Tx := RetourneTaux('1')
    else if Abs(Nb) <= TobCond.GetDouble('TCN_PLAFOND2') * NbJ then
      Tx := RetourneTaux('2')
    else if Abs(Nb) > TobCond.GetDouble('TCN_PLAFOND2') * NbJ then
      Tx := RetourneTaux('3');
    Result := Tx * Nb;
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_ECHELLESINTERETS.OnUpdate ;
{---------------------------------------------------------------------------------------}
var
  Q         : TQuery;
  SQL       : string;
  NbCol     : Integer;
  DateDe,
  DateA     : TDateTime;
  n         : Integer;
  Montant   : Double;
  {IntCre, En Attente}
  IntDeb    : Double;
  TobL,
  TobCondition,
  TobCond   : TOB;
  CurRow    : HTStrings;
  SoldePrec : Double;
  FiltreeOk : Boolean;
  ClauseNat : string;
begin
  inherited;
  {28/12/05 : Teste si les critères sont bons}
  if not TesteSaisie then Exit;

  MntComMvt := 0;
  {Récupération du taux de commission de mouvements défini dans le contrat}
  TxComMVT := GetTxComMvt;

  TobListe := TOB.Create('', nil, -1);
  TobCondition := TOB.Create('', nil, -1);
  TobGrid.ClearDetail;
  try
    Grid.VidePile(False);

    {Les échelles sont affichées par trimestre, sauf sur le trimestre en cours où on
     s'arrête à la date d'entrée}
    DateDe := StrToDate(edDateDe.Text);
    DateA  := StrToDate(edDateA .Text);

    ObjTaux.ChangerDates(DateDe, DateA, True);

    {28/12/05 : FQ 10314 : Nouvelle gestion de la requête}
    Q := OpenSQL(ConstruitRequete(FiltreeOk, ClauseNat), True);
    TobListe.LoadDetailDB('1', '', '', Q, False, True);
    Ferme(Q);

    {28/12/05 : FQ 10314 : Calcul des soldes si ceux issus de la requête ne sont pas bons :
                soit parce que les natures sont filtrées, soit parce qu'il y a fusion de comptes}
    SoldePrec := CalculSolde(ClauseNat);

    {Récupération des conditions de découvert}
    GetConditionDec('DEC', General, TobCondition);
    TobCond := TobCondition.Detail[0];

    {Affichage des conditions de découvert}
    TobCond.PutEcran(Ecran, TGroupBox(GetControl('CONDITIONS')));

    {Construction de la TOB à afficher}
    NbCol := 8;
    case TypePre.ItemIndex of
      0 : begin
            if TobListe.Detail.Count > 0 then
              Cumulee(TobCond, DateA, SoldePrec);
          end;
      1 : begin
            if TobListe.Detail.Count > 0 then
              Detaillee(TobCond, DateA, SoldePrec);
          end;
      2 : begin
            NbCol := 7 ;
            if TobListe.Detail.Count > 0 then
              Standard(TobCond, DateA, SoldePrec);
          end;
    end ;

    if TobGrid.Detail.Count > 0 then begin
      Montant := CalculPFD(TobCond);
      TobL := TobGrid.Detail[TobGrid.Detail.Count - 1];
      IntDeb := Abs(TobL.GetDouble('INTDEB'));
      {IntCre := TobL.GetDouble('INTCRE'); En attente}
      {MntComMvt := -CalcCMVT(Montant, TobCond);
       le taux est fixe en général 0.025%
       Son assiette est : les mvts débiteurs - agios
                                             - virements d'équilibrage de bq à bq
                                             - annulations d'écriture
                                             - achats de change
                                             - débits étrangers
                                             - souscription (sicav ...)
       car ces mouvements sont déjà soumis à commission.
       Bref, pour le moment on laisse tomber}
      SoldeEstime := Valeur(TobGrid.Detail[TobGrid.Detail.Count - 2].GetString('SOLDE'));
    end
    else begin
      Montant   := 0;
      IntDeb    := 0;
      {IntCre    := 0; En attente}
      SoldeEstime := 0;
    end;

    {Affichage du tickets des agios}
    THNumEdit(GetControl('COMPFD'   )).Value := Montant;
    THNumEdit(GetControl('INTDEBIT' )).Value := Abs(IntDeb);
    {Pour le moment, on ne gère pas les comptes rémunérés
    THNumEdit(GetControl('INTCREDIT')).Value := IntCre;}

    THNumEdit(GetControl('COMMVT'   )).Value := MntComMvt;
    THNumEdit(GetControl('AGIOS'    )).Value := Montant + Abs(IntDeb) + MntComMvt;

    {Formatage des colonnes}
    Grid.ColCount := NbCol;
    Grid.ColWidths[COL_TYPE] := -1; {Invisible}
    Grid.ColAligns[1] := taCenter; {Date}
    Grid.ColWidths[1] := 80;

    SQL := '0.0000'; {Pour Taux}
    for N := 2 to NbCol - 1 do begin
      Grid.ColWidths [N] := 95;
      Grid.ColAligns [N] := taRightJustify;
      Grid.ColFormats[N] := FormatDevise;
    end;

    {Attention à bien définir le format avant le PutGridDetail, sinon les montants sont arrondis}
    if TypePre.ItemIndex in [0, 1] then Grid.ColFormats[5] := SQL
    else if TypePre.ItemIndex = 2  then Grid.ColFormats[4] := SQL;

    {Affichage de la TOB des intérêts}
    TobGrid.PutGridDetail(Grid, True, True, '', True);

    {Mise à jour titre colonnes. Elle doit s'effectuer après le PutGridDetail}
    CurRow := Grid.Rows[0];
    {0 : Cumulee
     1 : Detaillee
     2 : Standard}
    case TypePre.ItemIndex of
      0,
      1 : begin
            CurRow[1]:=TraduireMemoire('Date de Valeur');
            if TypePre.ItemIndex = 0 then begin
              CurRow[2] := TraduireMemoire('Solde');
              Grid.ColFormats[3] := '#';
              CurRow[3] := TraduireMemoire('Nombre de jour');
            end
            else begin
              CurRow[2] := TraduireMemoire('Mouvement');
              CurRow[3] := TraduireMemoire('Solde');
            end;
            CurRow[4] := TraduireMemoire('Nombre');
            CurRow[5] := TraduireMemoire('Taux');
            CurRow[6] := TraduireMemoire('Int débiteurs');
  //          CurRow[7]:=TraduireMemoire('Int créditeurs');
            CurRow[7] := '';
            CurRow[0] := '';
            {JP 07/10/03 : Pour le moment on cache la colonne créditeur}
            Grid.ColWidths[7] := - 1;
            Grid.ColWidths[0] := - 1;
          end;
      2 : begin
            CurRow[1] := TraduireMemoire('Date de Valeur');
            CurRow[2] := TraduireMemoire('Solde');
            CurRow[3] := TraduireMemoire('Nombre');
            Grid.ColFormats[4] := SQL;
            CurRow[4] := TraduireMemoire('Taux');
            CurRow[5] := TraduireMemoire('Int débiteurs');
  //          CurRow[6]:=TraduireMemoire('Int créditeurs');
            CurRow[6] := '';
            {JP 07/10/03 : Pour le moment on cache la colonne créditeur}
            Grid.ColWidths[6] := - 1;
            Grid.ColWidths[0] := - 1;
            CurRow[0] := '';
          end;
    end;
    Grid.Refresh;
  finally
    FreeAndNil(TobListe);
    FreeAndNil(TobCondition);
  end;

  {28/12/05 : Affichage de la source des informations}
  GestionInfos;
end ;

{Recherche du Plus Fort Découvertt : MPFD.
 Le calcul de la commission sur le PFD se fait de la manière suivante :
    1/ Sur chacun des mois on prend le plus fort découvert
    2/ On somme ces 3 montants auxquels on applique un taux fixe TCN_CPFD
    3/ On plafonne le resultat à la 1/2 des intérêts débiteurs}
{---------------------------------------------------------------------------------------}
function TOF_ECHELLESINTERETS.CalculPFD(TobCond : TOB) : Double;
{---------------------------------------------------------------------------------------}
var
  a, m, j : Word;
  CurMois : Word;
  MntTmp  : Double;
  MntAuto : Double;
  Montant : Double;
  MntPFD  : array [0..2] of Double;
  n, p    : Integer;
  UneOpe  : Boolean;
begin
  Result  := 0;
  for p := 0 to 2 do MntPFD[p] := 0;
  MntTmp  := 0;
  Montant := 0;
  MntAuto := 0;

  {S'il faut déduire MPFD du montant de l'autorisation de découvert}
  if TobCond.GetString('TCN_LIEAUTO') = 'X' then
    MntAuto := Abs(TobCond.GetDouble('TCN_AUTORISATION'));

  {Récupération du mois de début de trimestre}
  DecodeDate(ObjTaux.DateDeb, a, CurMois, j);
  {On exécute le traitement pour les trois mois du trimestre}
  for p := 0 to 2 do begin
    UneOpe := False;
    {On ne traite pas le dernier enregistrement de la Tob qui contient les totaux}
    for n := 0 to TobGrid.Detail.Count - 2 do begin
      DecodeDate(VarToDateTime(TobGrid.Detail[n].GetDateTime('DATEVALEUR')), a, m, j);
      {Si on a changé de mois et passé le mois en cours}
      if m > (CurMois + p) then Break
                           else Montant := TobGrid.Detail[n].GetDouble('SOLDE');

      if m = (CurMois + p) then begin
        {Ce morceau de code est inutile dans le cas d'un affichage standard car TobGrid a
         une tob fille par jour.
         S'il s'agit de la première opération du mois}
        if not UneOpe then begin
          {on regarde si l'opération est au premier du mois}
          if (j > 1) and (n > 0) then
            {dans ce cas, on récupère les solde courant de la fin du mois précédent}
            MntTmp := TobGrid.Detail[n - 1].GetDouble('SOLDE');

          {Pour signifier qu'il y a au moins une opération dans le mois}
          UneOpe := True;
        end;

        if Montant < MntTmp then MntTmp := Montant;
      end;
    end;

    {Récupération du plus fort découvert auquel on retranche l'autorisation de découvert}
    if UneOpe then
      MntPFD[p] := MntTmp + MntAuto
    else
      {S'il n'y a pas de mouvement sur le mois en cours, on récupère le solde du mois précédents}
      MntPFD[p] := Montant  + MntAuto;

    if MntPFD[p] > 0 then MntPFD[p] := 0;
    MntTmp  := 0;
    Montant := 0;
  end;

  {Cumul des trois plus forts découverts}
  for p := 0 to 2 do Result := Result + MntPFD[p];
  {Affichage du plus fort découvert}
  THNumEdit(GetControl('MPFD')).Value := Abs(Result);

  {Calcul de la commission}
  Result := Result * TobCond.GetDouble('TCN_CPFD') / 100;
  {19/10/04 : On plafonne la commission des intérêts débiteurs}
  if TobCond.GetDouble('TCN_PLAFONDCPFD') > 0 then
    Result := Abs(Max(Result, TobCond.GetDouble('TCN_PLAFONDCPFD') * TobGrid.Detail[TobGrid.Detail.Count - 1].GetDouble('INTDEB') / 100))
  else
    Result := Abs(Max(Result, 0.5 * TobGrid.Detail[TobGrid.Detail.Count - 1].GetDouble('INTDEB')));
end;

{---------------------------------------------------------------------------------------}
procedure TOF_ECHELLESINTERETS.OnLoad;
{---------------------------------------------------------------------------------------}
begin
  inherited;
  if TypePre.ItemIndex < 0 then
    TypePre.ItemIndex := 2;
  Ecran.Refresh;
end ;

{---------------------------------------------------------------------------------------}
procedure TOF_ECHELLESINTERETS.OnClose;
{---------------------------------------------------------------------------------------}
begin
  FreeAndNil(ObjTaux);
  FreeAndNil(TobGrid);
  if Assigned(lComMvt) then FreeAndNil(lComMvt);
  inherited;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_ECHELLESINTERETS.TyEchelleChange(Sender : TObject);
{---------------------------------------------------------------------------------------}
begin
  {Raffraîchissement de l'écran}
  CritereClick(Sender);
end;

{$IFDEF EAGLCLIENT}
{---------------------------------------------------------------------------------------}
procedure TOF_ECHELLESINTERETS.bChercheClick(Sender : TObject);
{---------------------------------------------------------------------------------------}
begin
  {JP 07/08/03 : Dans le chargement en CWAS, la recherche est lancé, ce qui pose un petit problème
                 car il n'y a pas de liste associée}
  Load;
  Update;
end;

{$ENDIF}

{---------------------------------------------------------------------------------------}
procedure TOF_ECHELLESINTERETS.CalculComMVT(Montant : Double; Flux : string);
{---------------------------------------------------------------------------------------}
begin
  {La commission s'applique sur les mouvements débiteurs et dont le code flux/rubrique
   est paramétré comme assujeti à cette commission}
  {21/07/04 : Il semblerait (trésorière de Cegid dixit) que la commission de mouvement puisse s'appliquer sur des montant positifs}
  if {(Montant < 0) and} (lComMvt.IndexOf(Flux) > -1) then
    MntComMvt := MntComMvt + Abs(Montant * TxComMVT / 100);
end;

{---------------------------------------------------------------------------------------}
function TOF_ECHELLESINTERETS.GetTxComMvt : Double;
{---------------------------------------------------------------------------------------}
var
  Q : TQuery;
  D : TDateTime;
begin
  {On récupère le contrat valide à la fin du trimestre}
  D := StrToDate(edDateA.Text); {FQ 10284}
  Q := OpenSQL('SELECT TRC_COMMVT FROM TRCONTRAT WHERE TRC_DEBCONTRAT <= "' + UsDateTime(D) +
               '" AND TRC_FINCONTRAT >= "' + UsDateTime(D) +  '" AND ' +
                  'TRC_AGENCE IN (SELECT BQ_AGENCE FROM BANQUECP WHERE BQ_CODE = "' + General + '")', True);
  {Le paramétrage dans contrat n'est peut-être pas très utile car dans la littérature on parle
   0.025% systématiquement}
  if not Q.EOF then Result := Q.FindField('TRC_COMMVT').AsFloat
               else Result := 0.0250;
  Ferme(Q);
end;

{FQ 10314 : Activation des controls de la fusion si celle-ci est demandée
{---------------------------------------------------------------------------------------}
procedure TOF_ECHELLESINTERETS.ckFusionClick(Sender : TObject);
{---------------------------------------------------------------------------------------}
begin
  rgFusion .Enabled := ckFusion.Checked;
  cbCompte .Enabled := ckFusion.Checked;
  cbAgence .Enabled := ckFusion.Checked;
  cbbanque .Enabled := ckFusion.Checked;
  edGeneral.Enabled := not ckFusion.Checked;
  SetControlEnabled('TTCE_COMPTE1', not ckFusion.Checked);
  SetControlEnabled('DEV'         , not ckFusion.Checked);
  SetControlEnabled('IDEV'        , not ckFusion.Checked);
  {28/12/05 : Gestion du format des devises et du nombres de décimales d'affichages}
  GestionDevise;
  {Rafraîchissement du mul}
  CritereClick(Sender);
end;

{FQ 10314 : Affichage de la combo idoine en fonction du type de fusion
{---------------------------------------------------------------------------------------}
procedure TOF_ECHELLESINTERETS.rgFusionClick(Sender : TObject);
{---------------------------------------------------------------------------------------}
begin
  cbCompte.Visible := rgFusion.ItemIndex = 2;
  cbAgence.Visible := rgFusion.ItemIndex = 1;
  cbbanque.Visible := rgFusion.ItemIndex = 0;
  {28/12/05 : Gestion du format des devises et du nombres de décimales d'affichages}
  GestionDevise;
  {Rafraîchissement du mul}
  CritereClick(Sender);
end;

{10/06/07 : Construit la clause Where sur les comptes, agences et banques
{---------------------------------------------------------------------------------------}
function TOF_ECHELLESINTERETS.GetWhereBancaire : string;
{---------------------------------------------------------------------------------------}
begin
  Result := 'AND TE_GENERAL = "' + General + '" ';
  
  if ckFusion.Checked then begin
    case rgFusion.ItemIndex of
      {Par banque}
      0 : Result := 'AND BQ_BANQUE = "' + cbBanque.Value + '" ';

      {Par agence}
      1 : Result := 'AND BQ_AGENCE = "' + cbAgence.Value + '" ';

     {Sélection libre}
      2 : Result := 'AND TE_GENERAL IN (' + GetClauseIn(cbCompte.Value) + ') ';
    end;
  end;
end;

{FQ 10314 : Constitution de la requête avec les jointures sur agences et banques
{---------------------------------------------------------------------------------------}
function TOF_ECHELLESINTERETS.ConstruitRequete(var FiltreeOk : Boolean; var cNat : string) : string;
{---------------------------------------------------------------------------------------}
var
  sNat : string;
  Wh   : string;
begin
  {TE_CODEFLUX : nécessaire pour les commissions de mouvements
   TE_GENERAL  : nécessaire pour les conditions de valeurs}
  if ckFusion.Checked then {Si plusieurs compte, le calcul se fait en devise pivot}
    Result := 'SELECT TE_GENERAL, TE_DATEVALEUR, TE_SOLDEDEVVALEUR, TE_MONTANT AS TE_MONTANTDEV, TE_CODEFLUX FROM TRECRITURE '
  else
    Result := 'SELECT TE_GENERAL, TE_DATEVALEUR, TE_SOLDEDEVVALEUR, TE_MONTANTDEV, TE_CODEFLUX FROM TRECRITURE ';

  Result := Result + 'LEFT JOIN BANQUECP ON BQ_CODE = TE_GENERAL ';

  Wh := 'WHERE TE_DATEVALEUR >= "' + UsDateTime(StrToDate(edDateDe.Text)) + '"' +
        ' AND TE_DATEVALEUR <= "'  + UsDateTime(StrToDate(edDateA .Text)) + '" ';

  Result := Result + Wh + GetWhereBancaire;

  {$IFDEF TRCONF}
  wh := TObjConfidentialite.GetWhereConf(V_PGI.User, tyc_Banque);
  if wh <> '' then wh := ' AND (' + wh + ') ';
  Result := Result + Wh;
  {$ENDIF TRCONF}

  {Si l'on demande une échelle détaillée et qu'on ne demande pas toutes les écritures}
  sNat := GetControlText('TYPEFLUX');
  if (Pos('<<' , sNat) = 0) and (Trim(sNat) <> '') then begin {<<Tous>> ou vide}
    cNat := ' AND TE_NATURE IN (' + GetClauseIn(sNat) + ') ';

    Result := Result + cNat;
    FiltreeOk := True;
  end
  else begin
    FiltreeOk := False;
    cNat := '';
  end;

  {En détaillé, pas de regroupement des montants sur une date, mais on récupère en premier le solde
   du compte d'où le order by}
  Result := Result + 'ORDER BY TE_CLEVALEUR';
end;

{FQ 10314 : On s'assure que le(s) compte(s) ou agence ou banque a bien été sélectionné
{---------------------------------------------------------------------------------------}
function TOF_ECHELLESINTERETS.TesteSaisie: Boolean;
{---------------------------------------------------------------------------------------}
begin
  FGeneral := '';
  Result := True;
  if ckFusion.Checked then begin
    case rgFusion.ItemIndex of
      {Par banque}
      0 : if cbBanque.Value = '' then begin
            Result := False;
            HShowMessage('0;' + Ecran.Caption + ';Veuillez sélectionner une banque.;W;O;O;O;', '', '');
          end;
      {Par agence}
      1 : if cbAgence.Value = '' then begin
            Result := False;
            HShowMessage('0;' + Ecran.Caption + ';Veuillez sélectionner une agence.;W;O;O;O;', '', '');
          end;
      {Par Liste de comptes}
      2 : if cbCompte.Value = '' then begin
            Result := False;
            HShowMessage('0;' + Ecran.Caption + ';Veuillez sélectionner au moins un compte.;W;O;O;O;', '', '');
          end;
    end;
  end
  else begin
    if General = '' then begin
      Result := False;
      HShowMessage('0;' + Ecran.Caption + ';Veuillez sélectionner un compte.;W;O;O;O;', '', '');
    end;
  end;
end;

{28/12/05 : FQ 10314 : Calcul des soldes de début de trimestre
{---------------------------------------------------------------------------------------}
function TOF_ECHELLESINTERETS.GetSoldeGeneral(DateDe, Nature : string) : Double;
{---------------------------------------------------------------------------------------}
var
  aDate : string;
begin
  Result := 0;
  {Sur une fusion de comptes}
  if ckFusion.Checked then begin
    case rgFusion.ItemIndex of
      {Par banque}
      0 : Result := GetSoldeBanque(cbBanque.Value, DateDe, Nature, True);
      {Par Agence}
      1 : Result := GetSoldeBanque(cbAgence.Value, DateDe, Nature, True, True);
      {Liste de comptes}
      2 : Result := GetSoldeMultiComptes(cbCompte.Value, DateDe, Nature, True);
    end;
  end

  {Sur un compte}
  else begin
    if Nature <> '' then Result := GetSoldeInit(General, DateDe, Nature, True)
                    else Result := GetSoldeValeur(General, DateToStr(StrToDate(DateDe) - 1), aDate);
  end;
end;

{28/12/05 : FQ 10314 : Determination du Général de référence pour les conditions de découvert
            et du contrat bancaire pour la commission de mouvement :
            1/ Si un seul compte genéral, pas de problème
            2/ Si sélection de comptes, on prend le premier dans le MutliValCombo
            3/ Par agence ou par banque :
                  a/ soit le premier compte dans la liste des écritures
                  b/ soit le premier compte de l'agnec ou de la banque que l'on trouve
               Le compte sélectionné est stocké dans la variable FGeneral
{---------------------------------------------------------------------------------------}
function TOF_ECHELLESINTERETS.GetGeneral : string;
{---------------------------------------------------------------------------------------}
var
  s : string;
  Q : TQuery;
begin
  if not ckFusion.Checked then
    Result := edGeneral.Text
  else begin
    case rgFusion.ItemIndex of
      {Par banque ou par agence}
      0, 1 : begin
               if Assigned(TobListe) and (TobListe.Detail.Count > 0) then begin
                 Result := TobListe.Detail[0].GetString('TE_GENERAL');
                 FGeneral := Result;
               end

               else begin
                 if FGeneral <> '' then
                   Result := FGeneral

                 else begin
                   s := 'SELECT BQ_CODE FROM BANQUECP WHERE ';
                   if rgFusion.ItemIndex = 0 then s := s + 'BQ_BANQUE = "' + cbBanque.Value + '"'
                                             else s := s + 'BQ_AGENCE = "' + cbAgence.Value + '"';
                   Q := OpenSQL(s, True);
                   if not Q.EOF then begin
                     Result := Q.FindField('BQ_CODE').AsString;
                     FGeneral := Result;
                   end;
                   Ferme(Q);
                 end;
               end;
             end;
      {Liste de comptes}
      2 : begin
            s := cbCompte.Value;
            Result := ReadTokenSt(s);
          end;
    end;
  end;
end;

{28/12/05 : FQ 10314 : si on traite plusieurs comptes, on ne peut partir de TE_SOLDEDEVVALEUR qui
            est autonome pour chaque compte. On part donc du solde initial auquel on va ajouter les
            montants de chaque ligne afin d'obtenir un solde Cumulé
            L'ancien mode de calcul est conservé en fin d'unité
{---------------------------------------------------------------------------------------}
function TOF_ECHELLESINTERETS.CalculSolde(ClauseNat : string) : Double;
{---------------------------------------------------------------------------------------}
var
  n       : Integer;
  DateDe  : TDateTime;
  DateA   : TDateTime;
  Montant : Double;
  TobL    : TOB;
  SimplOk : Boolean;
  DebExOk : Boolean;
begin
  DateDe  := StrToDate(edDateDe.Text);
  DateA   := StrToDate(edDateA .Text);
  {REMARQUE : lorsque Result = -922337203685477, cela signifie qu'il y a des écritures le premier jour
              du trimestre et qu'il ne faut donc pas créer d'enregistrement pour le premier du trimestre}
  Result  := -922337203685477; {Limite inférieure d'un currency}
  Montant := 0;

  {On travaille sur un compte et il n'y a pas de filtre sur les natures, on travaillera sur
   les soldes issus de la requête
   11/07/07 : Nouvelle gestion des soldes : on calcule systématiquement les soldes par requête}
  SimplOk := (ClauseNat = '') and not ckFusion.Checked and not IsNewSoldes;

  if (TobListe.Detail.Count > 0) then begin
    if not SimplOk or (TobListe.Detail[0].GetDateTime('TE_DATEVALEUR') > DateDe) then
      {28/12/05 : FQ 10314 : Nouvelle gestion du calcul des soldes
       19/08/08 : Vu le fonctionnement de la fonction GetSoldeGeneral, il faut prendre DateDe -1}
      Montant := GetSoldeGeneral(DateToStr(DateDe), ClauseNat);

    {S'il n'y a pas d'écriture au premier jour, il faut créer une tob fille à TobGrid sur ce jour qui
     reprend le dernier solde trouvé : on ne peut partir du solde de la première fille de TobListe}
    if (TobListe.Detail[0].GetDateTime('TE_DATEVALEUR') > DateDe) then
      Result := Montant;
  end;

  {Récuperation du solde de la période précedente, si la première date est supérieure à celle
   du début du trimestre.
   Par exemple : debut de trimestre 01/10, première opération sur le trimestre 05/10 => il faut calculer
   les agios du 01/10 au 05/10. Il faut donc récupérer le solde en cours au 01/10 qui est celui figurant
   sur la dernière opération du trimestre pécédent, par exemple le 27/09.

   Par ailleurs, si on ne traite pas toutes les écritures, on ne peut pas travailler avec les soldes de
   la table TRECRITURE : il faut les recalculer

   Enfin, il reste le cas particulier du premier trimestre, avec le solde forcé. S'il n'y a pas d'opération
   le 01/01, il ne faut pas prendre le solde de la dernière opération sur l'année précédente mais celui du
   01/01 et s'il y a des opérations, on soustrait au solde forcé les opérations du 01/01}
  if (TobListe.Detail.Count > 0) and not SimplOk then begin
    {Recalcul des soldes}

    {Si le 01/01 n'appartient pas à la fourchette de dates DebExOk = True}
    DebExOk := not Between(DebutAnnee(DateA), DateDe, DateA);

    {Recalcul des soldes en additionnant les opérations du jour au solde précédent}
    for n := 0{p} to TobListe.Detail.Count - 1 do begin
      {Le calcul des soldes se fait plus tard avec le solde de réinitialisation}
      if not DebExOk and (TobListe.Detail[n].GetDateTime('TE_DATEVALEUR') = DebutAnnee(DateA)) then begin
        {Calcul du solde initial au matin du 01/01}
        Montant := GetSoldeGeneral(DateToStr(DebutAnnee(DateA)), ClauseNat);
        DebExOk := True;
      end;

      Montant := Montant + TobListe.Detail[n].GetDouble('TE_MONTANTDEV');
      TobListe.Detail[n].PutValue('TE_SOLDEDEVVALEUR', Montant);
    end;
  end {TobListe.Detail.Count > 0}

  {Il n'y a pas de mouvements sur le trimestre : on regarde si le dernier solde était négatif et si oui
   on va créer une ligne au premier du trimestre avec ce solde.
   Remarque : S'il y a une écriture de réinitialisation on ne passe pas par ici}
  else if TobListe.Detail.Count = 0 then begin
    {Récupération du dernier solde disponible}
    Montant := GetSoldeGeneral(DateToStr(DateDe), ClauseNat);
    if Montant < 0 then begin
      TobL := TOB.Create('1', TobListe, -1);
      TobL.AddChampSupValeur('TE_DATEVALEUR', DateDe);
      TobL.AddChampSupValeur('TE_SOLDEDEVVALEUR', Montant);
      TobL.AddChampSupValeur('TE_MONTANTDEV', 0.0);
    end
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_ECHELLESINTERETS.GestionDevise;
{---------------------------------------------------------------------------------------}
var
  DevISO : string;
  Nb     : Integer;
begin
  {S'il y a un filtre, peut être appelé avant l'affectation des composants}
  if GetCheckBoxState('CKFUSION') <> cbChecked then DevISO := RetDeviseCompte(GetControlText('EDGENERAL'))
                                               else DevIso := V_PGI.DevisePivot;
  SetControlCaption('DEV', DevISO);
  AssignDrapeau(TImage(GetControl('IDEV')), DevISO);
  if DevISO = V_PGI.DevisePivot then Nb := V_PGI.OkDecV
                                else Nb := CalcDecimaleDevise(DevISO);
  FormatDevise := StrFMask(Nb, '', True); {Pour le format de colonnes}
end;

{---------------------------------------------------------------------------------------}
procedure TOF_ECHELLESINTERETS.GestionInfos;
{---------------------------------------------------------------------------------------}
var
  s : string;
begin
  SetControlVisible('LBCONDITION', ckFusion.Checked);
  if ckFusion.Checked then begin
    s := TraduireMemoire('Les conditions de découverts sont celles du compte ') + RechDom('TRBANQUECP', General, False);
    s := s + #13+ TraduireMemoire('Les montants sont exprimés en devise pivot');
    SetControlCaption('LBCONDITION', s);
  end;
end;

{---------------------------------------------------------------------------------------}
function TOF_ECHELLESINTERETS.GetInfosBanques : TQuery;
{---------------------------------------------------------------------------------------}
var
  Wh : string;
begin
  if ckFusion.Checked then begin
    case rgFusion.ItemIndex of
      {Par banque}
      0 : Wh := 'WHERE BQ_BANQUE = "' + cbBanque.Value + '"'; 
      {Par Agence}
      1 : Wh := 'WHERE BQ_AGENCE = "' + cbAgence.Value + '"';
      {Liste de comptes}
      2 : Wh := 'WHERE BQ_CODE IN (' + GetClauseIn(cbCompte.Value)  + ')';
    end;
  end

  {Sur un compte}
  else
    Wh := 'WHERE BQ_CODE = "' + General  + '"';

  Result := OpenSQL('SELECT BQ_DOMICILIATION, BQ_ADRESSE1, BQ_ADRESSE2, BQ_CODEPOSTAL, BQ_VILLE, BQ_PAYS, ' +
                    'BQ_CODEBIC, BQ_ETABBQ, BQ_GUICHET, BQ_NUMEROCOMPTE, BQ_CLERIB, BQ_CODEIBAN  FROM ' +
                    'BANQUECP ' + Wh , True)
end;

{10/06/07 : FQ 10432 : Affichage du détail des opérations
{---------------------------------------------------------------------------------------}
procedure TOF_ECHELLESINTERETS.FListeDblClick(Sender : TObject);
{---------------------------------------------------------------------------------------}
var
  TT  : TOB;
  Lib : string;
begin
  TT := InitTobPourDetail;
  try
    TT.SetString('GENERAL', GetControlText('EDGENERAL'));
    TT.SetInteger('APPEL', Ord(dsSolde));
    TT.SetInteger('PERIODICITE', Ord(tp_1));
    if (GetControlText('TYPEFLUX') <> '') and not THMultiValComboBox(GetControl('TYPEFLUX')).Tous then
      TT.SetString('NATURE', 'AND TE_NATURE IN (' + GetClauseIn(GetControlText('TYPEFLUX')) + ')');
    {ON affiche les opérations dont la date de valeur correspond à la ligne en cours}
    TT.SetString('DATEOPE', 'TE_DATEVALEUR');
    Lib := TraduireMemoire('Détail des opérations du ') + Grid.Cells[1, Grid.Row];
    TT.SetString('LIBELLE', Lib);
    TT.SetDateTime('DATEDEB', StrToDate(Grid.Cells[1, Grid.Row]));
    TT.SetDateTime('DATEFIN', StrToDate(Grid.Cells[1, Grid.Row]));
    TT.SetString('TYPE', na_Prevision);
    TT.SetString('LEFTJOIN', 'LEFT JOIN BANQUECP ON BQ_CODE = TE_GENERAL');
    TT.SetString('WHERE', GetWhereBancaire);

    (*
    TT.SetString('DATEOPE', sDateOpe);
    TT.SetDateTime('DATEDEB', ColToDate(PtClick.X - COL_DATEDE));
    TT.SetDateTime('DATEFIN', ColToDate(PtClick.X - COL_DATEDE, True));

    Lib := 'Flux ' + CurRow[COL_LIBELLE] + ' en date de valeur';


    TT.SetString('CODE', CurRow[COL_TYPE]);
    TT.SetString('LIBNAT', RechDom('TRNATURE', CurRow[COL_NATURE], False));

    TT.SetString('MOIS', ' du ' + Grid.Cells[PtClick.X, 0]);
    *)
    TheTob := TT;
    {Appel de la fiche}
    TRLanceFiche_DetailFlux('TR', 'TRDETAILFLUX', '', '', ';ECHINT;');

  finally
    if Assigned(TT) then FreeAndNil(TT);
  end;
      (*
  case TypePre.ItemIndex of
    0 : begin
          if TobListe.Detail.Count > 0 then
            Cumulee(TobCond, DateA, SoldePrec);
        end;
    1 : begin
          if TobListe.Detail.Count > 0 then
            Detaillee(TobCond, DateA, SoldePrec);
        end;
    2 : begin
          NbCol := 7 ;
          if TobListe.Detail.Count > 0 then
            Standard(TobCond, DateA, SoldePrec);
        end;
  end ;
  *)
end;

initialization
  RegisterClasses([TOF_ECHELLESINTERETS]);


(* Ancien Calcul des soldes
    {Récuperation du solde de la période précedente, si la première date est supérieure à celle
     du début du trimestre.
     Par exemple : debut de trimestre 01/10, première opération sur le trimestre 05/10 => il faut calculer
     les agios du 01/10 au 05/10. Il faut donc récupérer le solde en cours au 01/10 qui est celui figurant
     sur la dernière opération du trimestre pécédent, par exemple le 27/09.

     Par ailleurs, si on ne traite pas toutes les écritures, on ne peut pas travailler avec les soldes de
     la table TRECRITURE : il faut les recalculer

     Enfin, il reste le cas particulier du premier trimestre, avec le solde forcé. S'il n'y a pas d'opération
     le 01/01, il ne faut pas prendre le solde de la dernière opération sur l'année précédente mais celui du
     01/01 et s'il y a des opérations, on soustrait au solde forcé les opérations du 01/01}
    if TobListe.Detail.Count > 0 then begin
      {Si l'on n'a pas d'écritures au premier jour du trimestre}
      if (TobListe.Detail[0].GetDateTime('TE_DATEVALEUR') > DateDe) then begin
        {Si l'on filtre les natures}
        if FiltreeOk then
          SoldePrec := GetSoldeInit(cbGeneral.Value, DateToStr(DateDe), ClauseNat, True)
        else begin
          {Sinon, on récupère le dernier solde dans la table (-1 var l'inégalité n'est pas stricte)}
          SoldePrec := GetSoldeValeur(cbGeneral.Value, DateToStr(DateDe - 1), SQL);
        end;
      end
      else begin
        {Il y a des opérations le premier jour du trimestre}
        if FiltreeOk then
          {On récupère le solde à la fin du premier jour du trimestre, d'où le DateDe + 1, car GetSoldeInit
           repose sur une inégalité stricte where te_datevaleur < DateDe + 1}
          SoldePrec := GetSoldeInit(cbGeneral.Value, DateToStr(DateDe + 1), ClauseNat, True)
        else
          SoldePrec := -922337203685477; {Limite inférieure d'un currency}
      end;
      {REMARQUE : lorsque SoldePrec = -922337203685477, cela signifie qu'il y a des écritures le premier jour
                  du trimestre et qu'il ne faut donc pas créer d'enregistrement pour le premier du trimestre}

      {Recalcul des soldes}
      if FiltreeOk then begin
        Montant := SoldePrec;
        p := 0;
        {Le cas du premier trimestre est particulier s'il y a des opérations le 01/01,
         car il faut gérer le solde forcé en fin de journée et recalculer les soldes du 01/01 à rebours}
        if (DebutAnnee(DateDe) = DateDe) and (TobListe.Detail[0].GetDateTime('TE_DATEVALEUR') = DateDe) then begin
          for n := 0 to TobListe.Detail.Count - 1 do
            {On cherche la première ligne dont la date de valeur soit supérieure au 01/01}
            if TobListe.Detail[n].GetDateTime('TE_DATEVALEUR') > DateDe then begin
              p := n - 1;
              Break;
            end;

          {Il y a un montant forcé}
          if TobListe.Detail[p].GetDouble('TE_MONTANTDEV') = 0 then begin
            {On recalcule les soldes du premier janvier à rebours en partant du solde forcé}
            for n := p - 1 downto 0 do begin
              Montant := Montant - TobListe.Detail[n].GetDouble('TE_MONTANTDEV');
              TobListe.Detail[n].PutValue('TE_SOLDEDEVVALEUR', Montant);
            end;
            {On incrémente P, pour ne pas traiter ensuite l'enregistrement du montant forcé}
            Inc(p);
            Montant := SoldePrec;
          end
          {Il n'y a pas de montant forcé, même gestion que pour les autre trimestres}
          else begin
            TobListe.Detail[0].PutValue('TE_SOLDEDEVVALEUR', Montant);
            SoldePrec := -922337203685477; {Limite inférieure d'un currency}
            p := 1;
          end;
        end

        {S'il y a des écritures le premier jour, le solde récupéré est celui de la fin de journée...}
        else if TobListe.Detail[0].GetDateTime('TE_DATEVALEUR') = DateDe then begin
          TobListe.Detail[0].PutValue('TE_SOLDEDEVVALEUR', Montant);
          SoldePrec := -922337203685477; {Limite inférieure d'un currency}
          p := 1;
        end
        {... sinon, il faut ajouter les opérations du jours au solde précédent}
        else begin
          Montant := Montant + TobListe.Detail[0].GetDouble('TE_MONTANTDEV');
          TobListe.Detail[0].PutValue('TE_SOLDEDEVVALEUR', Montant);
          p := 1;
        end;
        {Recalcul des soldes en aditionnnant les opérations du jour au solde récédent}
        for n := p to TobListe.Detail.Count - 1 do begin
          Montant := Montant + TobListe.Detail[n].GetDouble('TE_MONTANTDEV');
          TobListe.Detail[n].PutValue('TE_SOLDEDEVVALEUR', Montant);
        end;
      end;{FiltreeOk}
    end {TobListe.Detail.Count > 0}

    {Il n'y a pas de mouvements sur le trimestre : on regarde si le dernier solde était négatif et si oui
     on va créer une ligne au premier du trimestre avec ce solde}
    else begin
      SQL := 'SELECT TE_SOLDEDEVVALEUR FROM TRECRITURE WHERE ' + 'TE_DATEVALEUR <= "' +
             UsDateTime(DateDe) + '" AND AND TE_GENERAL = "' + cbGeneral.Value + '" ' + ClauseNat +
             'ORDER BY TE_CLEVALEUR DESC';
      Q := OpenSQL(SQL, True);
      if (not Q.EOF) and (Q.Fields[0].AsFloat < 0) then begin
        TobL := TOB.Create('1', TobListe, -1);
        TobL.AddChampSupValeur('TE_DATEVALEUR', DateDe);
        TobL.AddChampSupValeur('TE_SOLDEDEVVALEUR', Q.Fields[0].AsFloat);
        TobL.AddChampSupValeur('TE_MONTANTDEV', 0.0);
      end;
      Ferme(Q);
      {Pour ne pas aller rechercher le solde précédent}
      SoldePrec := -922337203685477;
    end;
*)

end.

