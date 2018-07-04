{-------------------------------------------------------------------------------------
  Version    |  Date   | Qui |   Commentaires
--------------------------------------------------------------------------------------
  1.01        19/01/04   JP   Création de l'unité
1.50.000.000  22/04/04   JP   Nouvelle gestion des natures (OnUpdate, ModifierLibelle)
1.50.000.000  12/05/04   JP   Gestion des comptes non mouvementés dans la gestion par solde
6.30.001.002  08/03/05   JP   FQ 10217 : gestion de la date d'opération lors de l'appel de la saisie.
                              ATTENTION : si on met en place le suivi hebdomadaire, il faudra prendre
                              des précautions !!!! cf. UFicTableau.GetCurDate
6.50.001.002  01/06/05   JP   FQ 10266 : La gestion à l'appel de CommanderDetail est érroné depuis la mise
                              en place des suivis hebdomadaire et mensuel sur les autres fiches de suivi
7.00.001.001  16/01/06   JP   FQ 10328 : Refonte de la gestion des devises
7.00.001.009  29/05/06   JP   FQ 10345 : Problème de nature sur le calcul des soldes en affichage par banque
7.09.001.001  10/10/06   JP   FQ 10376 : Affichage du jour devant la date => fonction UFicTableau.GetDateStr
7.09.001.001  10/11/06   JP   Dans le cadre de la FQ 10256, mise aux "normes" des autres fiches de suivi
7.09.001.001  15/11/06   JP   FQ 10302 : Afficher les soldes (par banque) même si les banques ne sont pas mouvementées
7.09.001.005  07/02/07   JP   On n'affiche le détail des comptes que sur les comptes du regroupement Tréso
--------------------------------------------------------------------------------------}

unit TRSUIVITRESO_TOF ;

interface

uses
  StdCtrls, Controls, Classes, UFicTableau,
  {$IFNDEF EAGLCLIENT}
  db, {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
  {$ENDIF}
  SysUtils, HCtrls, HEnt1, UTOB, Grids, Graphics, Windows, UObjGen;

type
  TObjSuiv = class
    CumSynch : Double;
    CumTreso : Double;
    LibBque  : string;
    Ind1     : Integer;
    Ind2     : Integer;
  end;

  TOF_TRSUIVITRESO = class(FICTAB)
    procedure OnUpdate              ; override;
    procedure OnArgument(S : string); override;
    procedure OnClose               ; override;
  private
    {13/11/06 : Adaptation à la gestion mensuelle}
    AnneeDe : Integer;
    MoisDe  : Integer;

    function  GetObjSuiv(CodBq, TitreDt, LibBq : string) : TObjSuiv;
    function  AnneeFromCol(n : Integer) : Integer;
    function  MoisFromCol (n : Integer) : Integer;
    function  IsColReInit (n : Integer) : Boolean;
    procedure SetMoisAnne(DateRef : TDateTime);
    function GereSoldeBanque(Banque, Endate, Nature : string; EnValeur : Boolean; AgenceOk : Boolean = False; NoDossier : string = '') : Double;
  protected
    DetailCpt : Boolean;
    txConv    : Double;
    lIndice   : TStringList;
    rbBanque  : TRadioButton;

    {Méthodes}
    procedure AfficherCompte (DateDe: TDateTime);
    procedure AfficherSoldes (DateDe: TDateTime);
    procedure GererSoldeInit (Taux : Double);
    {Evènements}
    procedure FormOnKeyDown  (Sender : TObject; var Key: Word; Shift: TShiftState); override;
    procedure DessinerCells  (Sender : TObject; Col, Row : Integer; Rect : TRect; State : TGridDrawState); override;
    procedure GeneralOnChange(Sender : TObject);
    procedure DeviseOnChange (Sender : TObject);
    procedure bGraphClick    (Sender : TObject); override;
    procedure bGraph1Click   (Sender : TObject);
    procedure PopupSimulation(Sender : TObject); override;
    {Clause where sur les banques et les comptes}
    function  GetWhereBqCpt : string; override;
  public
    function  CommandeDetail(CurRow : HTStrings; var TT : TOB) : Boolean; override;
    procedure CritereClick   (Sender : TObject); override;
  end;

function TRLanceFiche_TRSUIVITRESO(Dom: string; Fiche: string; Range: string; Lequel: string; Arguments: string): String;

implementation

uses
  ExtCtrls, Commun, TRGRAPHTROCTA_TOF, TRSAISIEFLUX_TOF, HTB97, UProcGen, Constantes,
  UProcSolde, HMsgBox, UTobDebug;

const
  LIBMAT = 'Solde initial';
  LIBSYN = 'Solde après synchronisation';
  LIBTRE = 'Solde après opé. financières';

{---------------------------------------------------------------------------------------}
function TRLanceFiche_TRSUIVITRESO(Dom: string; Fiche: string; Range: string; Lequel: string; Arguments: string): String;
{---------------------------------------------------------------------------------------}
begin
  Result := FICTAB.FicLance(Dom, Fiche, Range, Lequel, Arguments);
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRSUIVITRESO.OnArgument (S : String ) ;
{---------------------------------------------------------------------------------------}
begin
  TypeDesc := dsTreso;
  inherited ;
  LargeurCol := 100;

  txConv := 1; {Par défaut, on est en devise pivot}

  InArgument := True;

  TRadioButton(GetControl('DATEVAL')).OnClick := CritereClick;
  rbBanque :=  TRadioButton(GetControl('RBBANQUE'));
  rbBanque.OnClick := CritereClick;
  TRadioButton(GetControl('RBSOLDE')).OnClick := CritereClick;

  cbDevise.OnChange := DeviseOnChange;

  cbBanque.OnChange := GeneralOnChange;

  {Initialisation de la devise}
  Devise := V_PGI.DevisePivot;
  MajComboDevise(Devise);

  SetControlVisible('BGRAPH', not rbBanque.Checked);
  SetControlVisible('BGRAPH1', not rbBanque.Checked);
  TToolbarButton97(GetControl('BGRAPH1')).OnClick := bGraph1Click;
  TToolbarButton97(GetControl('BGRAPH1')).Hint := TraduireMemoire('Graph des cumuls des opérations');
  TToolbarButton97(GetControl('BGRAPH' )).Hint := TraduireMemoire('Graph des soldes');
  lIndice := TStringList.Create;
  InArgument := False;
end ;

{---------------------------------------------------------------------------------------}
procedure TOF_TRSUIVITRESO.OnClose;
{---------------------------------------------------------------------------------------}
begin
  LibereListe(lIndice, True);
  inherited;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRSUIVITRESO.DessinerCells(Sender: TObject; Col, Row: Integer; Rect: TRect; State: TGridDrawState);
{---------------------------------------------------------------------------------------}
var
  S, sType : string;
begin
  inherited;
  S := Grid.Cells[Col, Row];
  with Grid.Canvas do begin
    {Lignes de valeurs}
    sType := Grid.Cells[COL_TYPE, Row];

    {Pour les lignes de soldes après synchronisation ou Tresorerie, les montants sont
     mis en gras s'il y a eu des mouvements pour le compte/banque ce jour-là}
    if (sType = '+') or (sType = '-') then
      if ListeSolde.IndexOf(Grid.Cells[Col, 0] + ';' + Grid.Cells[0, Row] + ';' + Grid.Cells[1, Row]) > -1 then
        Font.Style := [fsBold];

    {Un affichage par solde, sur les totaux}
    if ((sType = 'µ') or (sType = '£') or (sType = '§')) then begin
      Font.Style := [fsBold];
      Brush.Color := AltColors[V_PGI.NumAltCol];
    end;

    if gdSelected in State then begin
      Brush.Color := clHighLight;
      Font.Color := clHighLightText;
    end

    {On est sur des détails}
    else if (sType = '+') or (sType = '-') or (sType = '@') then begin
      if (Row mod 2 = 0) then Brush.Color := clBtnFace
                         else Brush.Color := clInfoBk;
    end

    {On est sur des libellés}
    else if sType = '' then begin
      Brush.Color := clUltraLight;
      Font.Style := [fsBold];
    end;

    {si par solde et au premier janvier}
//    if (Col = ColRouge) and (ListeRouge.IndexOf(Grid.Cells[0, Row] + ';' +Grid.Cells[1, Row]) > -1) then
//    if IsColAvecSoldeInit(Col - COL_DATEDE) and (ListeRouge.IndexOf(Grid.Cells[0, Row] + ';' +Grid.Cells[1, Row]) > -1) then
    if IsColReInit(Col - COL_DATEDE) and
      (ListeRouge.IndexOf(Grid.Cells[0, Row] + ';' +Grid.Cells[1, Row] + ';' + RetourneCol(Col - COL_DATEDE)) > -1) then
      Font.Color := clRed;
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRSUIVITRESO.FormOnKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
{---------------------------------------------------------------------------------------}
var
  I : Integer;
begin
  if (Shift = [ssCtrl]) then begin
    case Key of
      VK_RIGHT : for I := Grid.Col + 1 to Grid.ColCount - 1 do
                   {Si la date appartient à la liste c'est qu'il y a des écritures}
                   if ListeSolde.IndexOf(Grid.Cells[I, 0] + ';' +  Grid.Cells[0, Grid.Row] + ';' +  Grid.Cells[1, Grid.Row]) > -1 then begin
                     Grid.Col := I;
                     Key := 0;
                     Exit;
                   end;

      VK_LEFT : for I := Grid.Col - 1 downto COL_DATEDE do
                  {Si la date appartient à la liste c'est qu'il y a des écritures}
                   if ListeSolde.IndexOf(Grid.Cells[I, 0] + ';' +  Grid.Cells[0, Grid.Row] + ';' +  Grid.Cells[1, Grid.Row]) > -1 then begin
                    Grid.Col := I;
                    Key := 0;
                    Exit;
                  end;
    end;
  end;

  FormKeyDown(Sender, Key, Shift); // Pour touches standard AGL
end ;

{---------------------------------------------------------------------------------------}
procedure TOF_TRSUIVITRESO.GeneralOnChange(Sender: TObject);
{---------------------------------------------------------------------------------------}
var
  Q : TQuery;
begin
  Devise := '';
  if Trim(cbBanque.Value) <> '' then begin
    Q := OpenSQL('SELECT DISTINCT(BQ_DEVISE) FROM BANQUECP WHERE BQ_BANQUE IN (' + GetClauseIn(cbBanque.Value) + ')', True);
    try
      if Q.RecordCount = 1 then Devise := Q.Fields[0].AsString
                           else Devise := V_PGI.DevisePivot;
    finally
      Ferme(Q);
    end;
  end;

  if Devise = '' then Devise := V_PGI.DevisePivot;

  {Mise à jour de la combo des devises}
  MajComboDevise(Devise);
  {Affichage de l'éventuelle image}
  AssignDrapeau(TImage(GetControl('IDEV')), Devise);
  CritereClick(cbBanque);
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRSUIVITRESO.DeviseOnChange(Sender : TObject);
{---------------------------------------------------------------------------------------}
begin
  Devise := cbDevise.Value;
  {Affichage de l'éventuelle image}
  AssignDrapeau(TImage(GetControl('IDEV')), Devise);
  CritereClick(cbDevise);
end;

{---------------------------------------------------------------------------------------}
function TOF_TRSUIVITRESO.CommandeDetail(CurRow : HTStrings; var TT : TOB) : Boolean;
{---------------------------------------------------------------------------------------}
var
  Par     : string;
  Q       : TQuery;
  TobL    : TOB;
  Code    : string;
  General : string;
begin
  Result := inherited CommandeDetail(CurRow, TT);
  Par := '';

  {A cette date, il n'y a pas d'écriture, on n'affiche donc pas le détail}
  if ListeSolde.IndexOf(Grid.Cells[PtClick.X, 0] + ';' +  Grid.Cells[0, PtClick.Y] + ';' +  Grid.Cells[1, PtClick.Y]) = -1 then Exit;

  Code      := CurRow[COL_TYPE];
  General   := CurRow[COL_CODE];
  if General = '' then Exit; {Sur banque}

  TobL := TobListe.FindFirst(['CODEBANQUE'], [General], True);
  if Assigned(TobL) then begin
    if GetCheckBoxState('CKCOMPTE') in [cbChecked] then Par := ' du compte ' + TobL.getvalue('LIBELLE')
                                                   else Par := ' de la banque ' + TobL.getvalue('LIBELLE');
  end;

  {On n'affiche le détail que sur les lignes après synchronisation et après trésorerie}
  if Code = '+' then begin
    Code := QUALIFCOMPTA;
    if Periodicite = tp_30 then
      LibNature := 'Écritures comptables' + Par + ' du ' + DateToStr(ColToDate(PtClick.X - COL_DATEDE)) + ' au ' + DateToStr(ColToDate(PtClick.X - COL_DATEDE, True))
    else
      LibNature := 'Écritures comptables' + Par + ' du ' + DateToStr(DateDepart + PtClick.X - COL_DATEDE);
  end
  else if Code = '-' then begin
    Code := QUALIFTRESO;
    if Periodicite = tp_30 then
      LibNature := 'Écritures de Trésorerie' + Par + ' du ' + DateToStr(ColToDate(PtClick.X - COL_DATEDE)) + ' au ' + DateToStr(ColToDate(PtClick.X - COL_DATEDE, True))
    else
      LibNature := 'Écritures de Trésorerie' + Par + ' du ' + DateToStr(DateDepart + PtClick.X - COL_DATEDE);
  end
  else Exit;

  Par := IntToStr(Ord(dsTreso));

  {L'affichage se fait par banque, on récupère les comptes de la banque en cours}
  if not (GetCheckBoxState('CKCOMPTE') in [cbChecked]) then begin
    Q := OpenSQl('SELECT BQ_CODE FROM BANQUECP WHERE BQ_BANQUE = "' + General + '"', True);
    try
      General := '';
      if Q.RecordCount = 0 then
        Exit
      else if Q.RecordCount = 1 then
        General := 'TE_GENERAL = "' + Q.Fields[0].AsString + '"'
      else begin
        General := 'TE_GENERAL IN (';
        while not Q.EOF do begin
          General := General + '"' + Q.Fields[0].AsString + '",';
          Q.Next;
        end;
        System.Delete(General, Length(General), 1);
        General := General + ')';
      end;
    finally
      Ferme(Q)
    end;
  end
  else
    General := 'TE_GENERAL = "' + General + '"';

  {Passé en argument : Mode ; Type de date ; Code flux ; Compte ; Date ; Libellé Nature ; Nature ; Libellé code flux}
  Code := Par + ';' + sDateOpe + ';' + Code + ';' + General + ';' +
          DateToStr(DateDepart + PtClick.X - COL_DATEDE) +';' +
          LibNature + ';' + Nature + ';' + CurRow[COL_LIBELLE];
  {01/06/05 : FQ 10266 : suite à la FQ 10237, La valeur de "Code" ne servait à rien car il s'agissait maintenant d'une
              variable Locale et non plus d'un paramètre de CommanderDeatil repris dans UFicTableau}
  TT.AddChampSupValeur('ARGTRESO', Code);
  TT.AddChampSupValeur('DATEDEB', iDate1900);
  TT.AddChampSupValeur('DATEFIN', iDate1900);
  if Periodicite = tp_30 then begin
    TT.SetDateTime('DATEDEB', ColToDate(PtClick.X - COL_DATEDE));
    TT.SetDateTime('DATEFIN', ColToDate(PtClick.X - COL_DATEDE, True));
  end;
  TT.SetString('REGROUP', GetControlText('REGROUPEMENT'));
  Result := True;
end;

{JP 27/10/03 : Lancement de l'écran de saisie des écritures de simulation
{---------------------------------------------------------------------------------------}
procedure TOF_TRSUIVITRESO.PopupSimulation(Sender: TObject);
{---------------------------------------------------------------------------------------}
var
  General : string;
begin
  inherited;
  General := Grid.Cells[COL_CODE, Grid.Row];
  {08/03/05 : FQ 10217 : se positionner sur la date d'opération, si on est en date d'opération}
  TRLanceFiche_SaisieFlux('TR', 'TRSAISIEFLUX', '', '', General + ';;' + DateCourante + ';');
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRSUIVITRESO.GererSoldeInit(Taux : Double);
{---------------------------------------------------------------------------------------}
var
  TobL,
  TobG,
  TobM : TOB;
  Q    : TQuery;
  Sufx : string;
  SQL  : string;
  Mnt  : Double;
  Dat  : string;
  soc  : string;
  MntB : Double;

          {----------------------------------------------------------------}
          function UsDateInit(T : TOB) : string;
          {----------------------------------------------------------------}
          begin
            if Periodicite = tp_1 then
              Result := USDateTime(T.GetDateTime(sDateOpe))
            else
              Result := USDateTime(EncodeDate(T.GetValue('TE_ANNEE'), T.GetValue('TE_MOIS'), 1));
          end;
var
  DtVal : TDateTime;
  Annee : Integer;
  Mois  : Integer;
begin
  DtVal := Date;
  Annee := 1;
  Mois  := 1;

  if not rbDateOpe.Checked then Sufx := 'VALEUR'
                           else Sufx := '';
  ListeRouge.Clear;
  TobListe.Detail.Sort('TE_QUALIFORIGINE');
  TobL := TobListe.FindFirst(['TE_QUALIFORIGINE'], [CODEREINIT], False);

  {Si pas de réinitialisation}
  while Assigned(TobL) do begin
    {13/11/03 : Gestion du NoDossier}
    Soc := GetWhereDossier;
    if Soc <> '' then Soc := ' AND ' + Soc;

    {Constitution de la requête qui va récupérer les soldes forcés en début de millésime
     en fonction des critères de sélection}
    if DetailCpt then SQL := 'SELECT TE_GENERAL AS COMPTE,'
                 else SQL := 'SELECT BQ_BANQUE AS COMPTE,';
    SQL := SQL + ' TE_SOLDEDEV' + Sufx + ', TE_COTATION FROM TRECRITURE';

    if not DetailCpt then SQL := SQL + ' LEFT JOIN BANQUECP ON TE_GENERAL = BQ_CODE';
    SQL := SQL + ' WHERE TE_QUALIFORIGINE = "' +
             CODEREINIT + '" AND ' + sDateOpe + ' = "' +  UsDateInit(TobL) + '"';

    SQL := SQL + Soc +  ' ORDER BY COMPTE';
    Q := OpenSQL(SQL, True);

    try
      {On boucle sur les enregistrements de réinitialisation de la TobGrid}
      while Assigned(TobL) do begin
        TobG := TobGrid.FindFirst(['TYPE', 'CODE'], ['-', TobL.getvalue('CODEBANQUE')], True);
        if Periodicite = tp_1 then
          DtVal := TobL.GetDateTime('TE_DATEVALEUR')
        else begin
          Mois  := TobL.GetInteger('TE_MOIS');
          Annee := TobL.GetInteger('TE_ANNEE');
        end;

        {On recherche dans la requête le solde forcé pour le compte en cours avec conversion
         du montant en euro}
        if Assigned(TobG) and Q.Locate('COMPTE', TobG.GetValue('CODE'), []) then begin
          Mnt := 0;
          {Si on travaille par banque, il est possible qu'il y ait plusieurs comptes donc plusieurs
           montants de réinitialisation. Je ne fait pas de SUM dans la requête car les comptes peuvent
           être dans des devises différentes}
          while not Q.EOF and (Q.Fields[0].AsString = TobG.GetValue('CODE')) do begin
            if (Q.Fields[2].AsFloat <> 0) and MntPivot then Mnt := Mnt + (Q.Fields[1].AsFloat / Q.Fields[2].AsFloat)
                                                       else Mnt := Mnt + Q.Fields[1].AsFloat;
            Q.Next;
          end;
          {Puis conversion du solde calculé dans la devise d'affichage}
          MntB := Mnt * Taux;
          {13/11/06 : si le 01/01 est le premier jour de la sélection, on met le solde sur la TobMatin}
          if ((TobL.GetDateTime(sDateOpe) = DateDepart) and (Periodicite = tp_1)) or
             ((Periodicite = tp_30) and (DateDepart = EncodeDate(TobL.GetValue('TE_ANNEE'), TobL.GetValue('TE_MOIS'), 1))) then begin
            Dat := GetTitreFromDate(TobL.GetDateTime(sDateOpe));
            TobM := TobGrid.FindFirst(['TYPE', 'CODE'], ['@', TobL.getvalue('CODEBANQUE')], True);
            if Assigned(TobM) then TobM.PutValue(Dat, MntB);
          end
          else begin
            if Periodicite = tp_1 then
              Dat := GetTitreFromDate(TobL.GetDateTime(sDateOpe) - 1)
            else
              Dat := GetTitreFromDate(EncodeDate(TobL.GetValue('TE_ANNEE'), TobL.GetValue('TE_MOIS'), 1) - 1);
            TobG.PutValue(Dat, MntB);
          end;
          {Le seul cas où l'on met le solde Init au Matin, c'est si l'on part du 01/01 ou de janvier et
           que l'on est sur la première colonne}
          if (DateDepart = DebutAnnee(DateDepart)) and
             (((TobL.GetDateTime(sDateOpe) = DateDepart) and (Periodicite = tp_1)) or
              ((Periodicite = tp_30) and (DateDepart = EncodeDate(TobL.GetValue('TE_ANNEE'), TobL.GetValue('TE_MOIS'), 1)))) then begin
            if Periodicite = tp_30 then
              ListeRouge.Add('@;' + TobL.getvalue('CODEBANQUE') + ';' + GetTitreFromDate(EncodeDate(TobL.GetValue('TE_ANNEE'), TobL.GetValue('TE_MOIS'), 1)))
            else
              ListeRouge.Add('@;' + TobL.getvalue('CODEBANQUE') + ';' + GetTitreFromDate(TobL.GetDateTime(sDateOpe)));
          end
          else begin
            if Periodicite = tp_30 then
              ListeRouge.Add('-;' + TobL.getvalue('CODEBANQUE') + ';' + GetTitreFromDate(EncodeDate(TobL.GetValue('TE_ANNEE'), TobL.GetValue('TE_MOIS'), 1) - 1))
            else
              ListeRouge.Add('-;' + TobL.getvalue('CODEBANQUE') + ';' + GetTitreFromDate(TobL.GetDateTime(sDateOpe) - 1));
          end;
        end;

        if Periodicite = tp_1 then
          TobL := TobListe.FindNext(['TE_QUALIFORIGINE', 'TE_DATEVALEUR'], [CODEREINIT, DtVal], False)
        else
          TobL := TobListe.FindNext(['TE_QUALIFORIGINE', 'TE_ANNEE', 'TE_MOIS'], [CODEREINIT, Annee, Mois], False)
      end;
    finally
      Ferme(Q);
    end;
    TobL := TobListe.FindNext(['TE_QUALIFORIGINE'], [CODEREINIT], False);
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRSUIVITRESO.AfficherSoldes(DateDe: TDateTime);
{---------------------------------------------------------------------------------------}
var
  TobM,
  TobS,
  TobL,
  TobC,
  TobG : TOB;
  lBanque : string;
  Premier : Boolean;
  {LÉGENDE DES TYPES : ""          = Ligne contenant le libellé des trois grands soldes (Mat, Syn, Soir)
                       "$"         = Ligne de séparation entre les trois soldes
                       "@"         = Solde au matin pour une banque / compte
                       "+"         = Solde après synchro pour une banque / compte
                       "-"         = Solde au soir pour une banque / compte
                       "§"         = Total au matin pour une banque / compte
                       "£"         = Total après synchro pour une banque / compte
                       "µ"         = Total au soir pour une banque / compte }

    {------------------------------------------------------------------}
    procedure CreeTobDetail(TypeLig, CodeLig, Libelle, Nat : string; AvecSolde : Boolean = False);
    {------------------------------------------------------------------}
    var
      n : Integer;
      m : Double;
    begin
      TobG := TOB.Create('', TobGrid, -1);
      TobG.AddChampSupValeur('TYPE', TypeLig);
      TobG.AddChampSupValeur('CODE', CodeLig);
      TobG.AddChampSupValeur('NATURE'  , Nat);
      TobG.AddChampSupValeur('LIBELLE', Libelle);
      {Ajout des colonnes dates}
      for n := 0 to NbColonne - 1 do begin
        {Éventuelle récupération des soldes au matin pour le premier jour}
        if (n = 0) and AvecSolde then begin
          if DetailCpt then 
            {Affichage par compte, le montant retourné est en Devise pivot}
            m := GetSoldeInit(CodeLig, DateToStr(DateDe), Nature, not rbDateOpe.Checked, MntPivot) * txConv {FQ 10345}
          else
            {Affichage par banque, le montant retourné est en Devise pivot}
            m := GereSoldeBanque(CodeLig, DateToStr(DateDe), Nature, not rbDateOpe.Checked, False, GetControlText('REGROUPEMENT')) * txConv;
            
          if StrFPoint(m) = '0.001' then m := 0;
          TobG.AddChampSupValeur(RetourneCol(n), m);
        end
        else if (TypeLig <> '') and (TypeLig <> '$') then
          TobG.AddChampSupValeur(RetourneCol(n), 0.0)
        else
          TobG.AddChampSupValeur(RetourneCol(n), '')
      end;
    end;

    {------------------------------------------------------------------}
    procedure CumulerMontant(Code : string);
    {------------------------------------------------------------------}
    var
      i : Integer;
      m : Double;
      p : Double;
    begin

      for i := 1 to NbColonne - 1 do begin
        m := 0;
        p := 0;

        if Periodicite = tp_30 then
          TobL := TobListe.FindFirst(['CODEBANQUE', 'TE_QUALIFORIGINE', 'TE_ANNEE', 'TE_MOIS'], [Code, QUALIFCOMPTA, AnneeFromCol(i), MoisFromCol(i)], True)
        else
          TobL := TobListe.FindFirst(['CODEBANQUE', 'TE_QUALIFORIGINE', sDateOpe], [Code, QUALIFCOMPTA, DateDe + i], True);
        {Il y a eu des mouvements ce jour, on mémorise donc les références pour le DblClick}
        if Assigned(TobL) then begin
          ListeSolde.Add(RetourneCol(i) + ';+;' + Code);
          m := TobL.getvalue('MONTANT') * txConv;
        end;

        if Periodicite = tp_30 then
          TobC := TobListe.FindFirst(['CODEBANQUE', 'TE_QUALIFORIGINE', 'TE_ANNEE', 'TE_MOIS'], [Code, QUALIFTRESO, AnneeFromCol(i), MoisFromCol(i)], True)
        else
          TobC := TobListe.FindFirst(['CODEBANQUE', 'TE_QUALIFORIGINE', sDateOpe], [Code, QUALIFTRESO, DateDe + i], True);
        {Il y a eu des mouvements ce jour, on mémorise donc les références pour le DblClick}
        if Assigned(TobC) then begin
          ListeSolde.Add(RetourneCol(i) + ';-;' + Code);
          p := TobC.getvalue('MONTANT') * txConv;
        end;
        {Solde au matin : on reprend le solde de la veille}
        TobM.PutValue(RetourneCol(i), TobS.getvalue(RetourneCol(i - 1)));
        {Solde après synchro : on reprend le solde du matin auquel on ajoute éventuellement les écritures
         comptables synchronisées}
        TobG.PutValue(RetourneCol(i), TobM.getvalue(RetourneCol(i)) + m);

        {Solde du soir si l'on n'est pas sur un début de millésime}
        if not IsColReInit(i) then begin
          {Solde du soir : on reprend le solde du "midi" auquel on ajoute éventuellement les écritures
           de trésorerie}
          TobS.SetDouble(RetourneCol(i), TobG.GetDouble(RetourneCol(i)) + p);
        end
        {Solde du soir en début de millésime. Remarque, le GererSoldeInit calcul tous les soldes de millésimes
         en une seule fois, d'où le boolean "Premier" : pour toutes les autres banques / comptes, on n'écrira
         rien le 01/01 car cela aura été fait !!}
        else if Premier and IsColReInit(i) then begin
          {Gestion éventuelle d'un solde de réinitialisation}
          GererSoldeInit(txConv);
          Premier := False;
        end;
      end;
    end;

var
  I       : Integer;
  Code    : string;
  Montant : Double;
  ObjEnt  : TObjSuiv;
  IndC    : Integer;
begin
  {On vide la liste des soldes vides}
  ListeSolde.Clear;
  LibereListe(lIndice, False);
  SetMoisAnne(DateDe);
  ColRouge := -1;

  Code    := '';
  lBanque := '';
  Premier := True;

  IndC := 0;

  {-------------- 1° ÉTAPE : -----------------
   Constitution de la première série de tob filles de TobGrid, contenant les soldes aux matins}
  CreeTobDetail('', '', LIBMAT, '');
  Inc(IndC);
  for I := 0 to TobListe.Detail.Count - 1 do begin
    TobL := TobListe.Detail[I];
    lBanque := TobL.getvalue('CODEBANQUE');

    {Construction de la liste qui servira lors des autres étapes}
    if lIndice.IndexOf(lBanque) = - 1 then begin
      {12/05/04 : Suppression des Comptes / banques mouvementés}
      if DetailCpt then SupprimeCptePresent(TobL.getvalue('BQ_BANQUE'), lBanque)
                   else SupprimeCptePresent(lBanque, '');

      ObjEnt := GetObjSuiv(lBanque, '', TobL.getvalue('LIBELLE'));
      {Création des filles pour chaque compte / banque pour le solde du matin}
      CreeTobDetail('@', lBanque, ObjEnt.LibBque, '', True);

      {Recherche d'un éventuel cumul des synchronisations au premier jour de la période}
      if Periodicite = tp_30 then
        TobL := TobListe.FindFirst(['CODEBANQUE', 'TE_QUALIFORIGINE', 'TE_ANNEE', 'TE_MOIS'], [lBanque, QUALIFCOMPTA, AnneeDe, MoisDe], True)
      else
        TobL := TobListe.FindFirst(['CODEBANQUE', 'TE_QUALIFORIGINE', sDateOpe], [lBanque, QUALIFCOMPTA, DateDe], True);
      if TobL <> nil then begin
        ObjEnt.CumSynch := TobL.getvalue('MONTANT') * txConv;
        {Il y a eu des mouvements ce jour, on mémorise donc les références pour le DblClick}
        ListeSolde.Add(RetourneCol(0) + ';+;' + lBanque);
      end
      else
        ObjEnt.CumSynch := 0;

      {Recherche d'un éventuel cumul des opérations de trésorerie au premier jour de la période}
      if Periodicite = tp_30 then
        TobL := TobListe.FindFirst(['CODEBANQUE', 'TE_QUALIFORIGINE', 'TE_ANNEE', 'TE_MOIS'], [lBanque, QUALIFTRESO, AnneeDe, MoisDe], True)
      else
        TobL := TobListe.FindFirst(['CODEBANQUE', 'TE_QUALIFORIGINE', sDateOpe], [lBanque, QUALIFTRESO, DateDe], True);
      if TobL <> nil then begin
        ObjEnt.CumTreso := TobL.getvalue('MONTANT') * txConv;
        {Il y a eu des mouvements ce jour, on mémorise donc les références pour le DblClick}
        ListeSolde.Add(RetourneCol(0) + ';-;' + lBanque);
      end
      else
        ObjEnt.CumTreso := 0;

      ObjEnt.Ind1 := IndC;
      Inc(IndC);
    end;
  end;

  {12/05/04 : ajout des comptes / banqes non mouvementés sur la période}
  while ListeCpte.Detail.Count > 0 do begin
    if DetailCpt then begin
      lBanque := ListeCpte.Detail[0].GetString('COMPTE');
      ObjEnt := GetObjSuiv(lBanque, '', ListeCpte.Detail[0].getvalue('LIBELLE'))
    end else begin
      lBanque := ListeCpte.Detail[0].GetString('BANQUE');
      ObjEnt := GetObjSuiv(lBanque, '', ListeCpte.Detail[0].getvalue('LIBBANQUE'));
    end;

    {Création des filles pour chaque compte / banque pour le solde du matin}
    CreeTobDetail('@', lBanque, ObjEnt.LibBque, '', True);

//    ObjEnt.CumSynch := TobG.GetDouble(RetourneCol(0));
  //  ObjEnt.CumTreso := TobG.GetDouble(RetourneCol(0));
    ObjEnt.Ind1 := IndC;
    Inc(IndC);

    {Suppression du Compte / banque }
    SupprimeCptePresent(ListeCpte.Detail[0].GetString('BANQUE'), ListeCpte.Detail[0].GetString('COMPTE'));
  end;

  {Tob contenant les soldes généraux du matin}
  CreeTobDetail('§', '', 'Solde général initial', '');
  {Tob de séparation entre les solde du matin et ceux après synchro}
  CreeTobDetail('$', '', '', '');
  Inc(IndC, 2);

  {-------------- 2° ÉTAPE : -----------------
   Création de la seconde série de tob filles contenant les soldes après synchronisation}
  CreeTobDetail('', '', LIBSYN, '');
  Inc(IndC);
  for i := 0 to lIndice.Count - 1 do begin
    ObjEnt := TObjSuiv(lIndice.Objects[i]);
    CreeTobDetail('+', lIndice[i], ObjEnt.LibBque, '');
    TobG.PutValue(GetTitreFromDate(DateDe), TobGrid.Detail[ObjEnt.Ind1].GetDouble(GetTitreFromDate(DateDe)) + ObjEnt.CumSynch);
    ObjEnt.Ind2 := IndC;
    Inc(IndC);
  end;

  {Tob contenant les soldes généraux après synchronisation}
  CreeTobDetail('£', '', 'Solde général après synchro.', '');
  {Tob de séparation entre les solde après synchro de ceux après Treso}
  CreeTobDetail('$', '', '', '');

  {-------------- 3° ÉTAPE : -----------------
  {Création de la troisième série de tob filles contenant les soldes après Tréso}
  CreeTobDetail('', '', LIBTRE, '');
  for i := 0 to lIndice.Count - 1 do begin
    ObjEnt := TObjSuiv(lIndice.Objects[i]);
    CreeTobDetail('-', lIndice[i], ObjEnt.LibBque, '');
    TobG.PutValue(GetTitreFromDate(DateDe), TobGrid.Detail[ObjEnt.Ind2].getvalue(GetTitreFromDate(DateDe)) + ObjEnt.CumTreso);
  end;

  {Gestion éventuelle d'un solde de réinitialisation}
  if IsColReInit(0) then begin
    GererSoldeInit(txConv);
    Premier := False;
  end;

  {Tob contenant les soldes généraux après Tréso}
  CreeTobDetail('µ', '', 'Solde général final', '');

  {-------------- 4° ÉTAPE : -----------------
  Maintenant la TobGrid est construite et ses 4 premières colonnes remplies (les 3 fixes et le premier jour)
  reste à remplir les autres dates avec des systèmes de report du soir au matin}
  for i := 0 to lIndice.Count - 1 do begin
    Code := lIndice[i];
    TobM := TobGrid.FindFirst(['TYPE', 'CODE'], ['@', Code], True);
    TobG := TobGrid.FindFirst(['TYPE', 'CODE'], ['+', Code], True);
    TobS := TobGrid.FindFirst(['TYPE', 'CODE'], ['-', Code], True);
    if (TobG = nil) or (TobS = nil) or (TobM = nil)  then Break;
    CumulerMontant(Code);
  end;

  {Calcul des Totaux généraux pour les 3 soldes}
  for i := 0 to NbColonne - 1 do  begin
    TobG := TobGrid.FindFirst(['TYPE'], ['§'], True);
    Montant := TobGrid.Somme(RetourneCol(i), ['TYPE'], ['@'], False);
    TobG.PutValue(RetourneCol(i), Montant);

    TobG := TobGrid.FindFirst(['TYPE'], ['£'], True);
    Montant := TobGrid.Somme(RetourneCol(i), ['TYPE'], ['+'], False);
    TobG.PutValue(RetourneCol(i), Montant);

    TobG := TobGrid.FindFirst(['TYPE'], ['µ'], True);
    Montant := TobGrid.Somme(RetourneCol(i), ['TYPE'], ['-'], False);
    TobG.PutValue(RetourneCol(i), Montant);
  end;

  LibereListe(lIndice, False);
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRSUIVITRESO.AfficherCompte(DateDe : TDateTime);
{---------------------------------------------------------------------------------------}
var
  TobM,
  TobS,
  TobL,
  TobG : TOB;
  lBanque : string;
  {LÉGENDE DES TYPES : ""          = Ligne contenant le libellé d'une banque
                       "$"         = Ligne de séparation entre deux banques
                       "@"         = Solde avant synchronisation
                       "+"         = Solde après synchronisation
                       "-"         = Solde après Tréso}

    {------------------------------------------------------------------}
    procedure CreeTobDetail(TypeLig, CodeLig, Libelle, Nat : string; AvecSolde : Boolean = False);
    {------------------------------------------------------------------}
    var
      n : Integer;
      m : Double;
    begin
      TobG := TOB.Create('', TobGrid, -1);
      TobG.AddChampSupValeur('TYPE', TypeLig);
      TobG.AddChampSupValeur('CODE', CodeLig);
      TobG.AddChampSupValeur('NATURE', Nat);  {FQ 10345}
      TobG.AddChampSupValeur('LIBELLE', Libelle);
      {Ajout des colonnes dates}
      for n := 0 to NbColonne - 1 do begin
        {Éventuelle récupération des soldes au matin pour le premier jour}
        if (n = 0) and AvecSolde then begin
          if DetailCpt then 
            {Affichage par compte, le montant retourné est en Devise pivot}
            m := GetSoldeInit(CodeLig, DateToStr(DateDe), Nature, not rbDateOpe.Checked, MntPivot) * txConv {FQ 10345}
          else
            {Affichage par banque, le montant retourné est en Devise pivot}
            m := GereSoldeBanque(CodeLig, DateToStr(DateDe), Nature, not rbDateOpe.Checked, False, GetControlText('REGROUPEMENT')) * txConv; {FQ 10345}

          if StrFPoint(m) = '0.001' then m := 0;
          TobG.AddChampSupValeur(RetourneCol(n), m);
        end

        else if (TypeLig <> '') and (TypeLig <> '$') then
          TobG.AddChampSupValeur(RetourneCol(n), 0)

        else
          TobG.AddChampSupValeur(RetourneCol(n), '');
      end;
    end;

var
  I, N    : Integer;
  S, C, D : string;
  Montant : Double;
  Premier : Boolean;
begin
  {On vide la liste des soldes vides}
  ListeSolde.Clear;
  LibereListe(lIndice, False);
  SetMoisAnne(DateDe);

  ColRouge := -1;

  lBanque := '';
  Premier := True;

  {-------------- 1° ÉTAPE : -----------------
   Génération des tobs filles et affectation des quatre premiers champs : type
   code, libellé et date de départ}
  for I := 0 to TobListe.Detail.Count - 1 do begin
    TobL := TobListe.Detail[I];
    lBanque := TobL.getvalue('CODEBANQUE');

    {Construction de la liste qui servira lors des autres étapes}
    if lIndice.IndexOf(lBanque) = - 1 then begin
      {15/11/06 : FQ 10302 : Suppression des Comptes / banques mouvementés}
      if DetailCpt then SupprimeCptePresent(TobL.getvalue('BQ_BANQUE'), lBanque)
                   else SupprimeCptePresent(lBanque, '');

      lIndice.Add(lBanque);
      {Création de la fille pour le libellé de chaque compte / banque}
      CreeTobDetail('', lBanque, TobL.getvalue('LIBELLE'), '');
      {Création des filles pour chaque compte / banque pour le solde du matin}
      CreeTobDetail('@', lBanque, LIBMAT, '', True);

      {Récupération du solde récupérer dans CreeTobDetail }
      Montant := TobG.GetDouble(RetourneCol(0));
      {Recherche d'un éventuel cumul des synchronisations au premier jour de la période}
      if Periodicite = tp_30 then
        TobL := TobListe.FindFirst(['CODEBANQUE', 'TE_QUALIFORIGINE', 'TE_ANNEE', 'TE_MOIS'], [lBanque, QUALIFCOMPTA, AnneeDe, MoisDe], True)
      else
        TobL := TobListe.FindFirst(['CODEBANQUE', 'TE_QUALIFORIGINE', sDateOpe], [lBanque, QUALIFCOMPTA, DateDe], True);
      if TobL <> nil then begin
        Montant := Montant + TobL.getvalue('MONTANT') * txConv;
        {Il y a eu des mouvements ce jour, on mémorise donc les références pour le DblClick}
        ListeSolde.Add(RetourneCol(0) + ';+;' + lBanque);
      end;

      {Création de la tob contenant le solde après synchronisation}
      CreeTobDetail('+', lBanque, LIBSYN, '');
      {Affectation du solde après synchronisation pour le premier jour de la période}
      TobG.PutValue(RetourneCol(0), Montant);

      {Recherche d'un éventuel cumul des opérations de trésorerie au premier jour de la période}
      if Periodicite = tp_30 then
        TobL := TobListe.FindFirst(['CODEBANQUE', 'TE_QUALIFORIGINE', 'TE_ANNEE', 'TE_MOIS'], [lBanque, QUALIFTRESO, AnneeDe, MoisDe], True)
      else
        TobL := TobListe.FindFirst(['CODEBANQUE', 'TE_QUALIFORIGINE', sDateOpe], [lBanque, QUALIFTRESO, DateDe], True);
      if TobL <> nil then begin
        Montant := Montant + TobL.getvalue('MONTANT') * txConv;
        {Il y a eu des mouvements ce jour, on mémorise donc les références pour le DblClick}
        ListeSolde.Add(RetourneCol(0) + ';-;' + lBanque);
      end;
      {Création de la tob contenant le solde après Trésorerie}
      CreeTobDetail('-', lBanque, LIBTRE, '');
      {Affectation du solde après Trésorerie pour le premier jour de la période}
      TobG.PutValue(RetourneCol(0), Montant);

      {Création d'une tob de séparation entre deux banques/comptes}
      CreeTobDetail('$', '', '', '');
    end;
  end;

  {15/11/06 : FQ 10302 : Suppression des Comptes / banques mouvementés}
  while ListeCpte.Detail.Count > 0 do begin
    if DetailCpt then begin
      lBanque := ListeCpte.Detail[0].GetString('COMPTE');
      S := ListeCpte.Detail[0].GetString('LIBELLE');
    end else begin
      lBanque := ListeCpte.Detail[0].GetString('BANQUE');
      S := ListeCpte.Detail[0].GetString('LIBBANQUE');
    end;

    lIndice.Add(lBanque);
    {Création de la fille pour le libellé de chaque compte / banque}
    CreeTobDetail('', lBanque, S, '');
    {Création des filles pour chaque compte / banque pour le solde du matin}
    CreeTobDetail('@', lBanque, LIBMAT, '', True);
    Montant := TobG.GetDouble(RetourneCol(0));

    {Création de la tob contenant le solde après synchronisation}
    CreeTobDetail('+', lBanque, LIBSYN, '');
    {Affectation du solde après synchronisation pour le premier jour de la période}
    TobG.PutValue(RetourneCol(0), Montant);

    {Création de la tob contenant le solde après Trésorerie}
    CreeTobDetail('-', lBanque, LIBTRE, '');
    {Affectation du solde après Trésorerie pour le premier jour de la période}
    TobG.PutValue(RetourneCol(0), Montant);

    {Création d'une tob de séparation entre deux banques/comptes}
    CreeTobDetail('$', '', '', '');
    {Suppression du Compte / banque }
    SupprimeCptePresent(ListeCpte.Detail[0].GetString('BANQUE'), ListeCpte.Detail[0].GetString('COMPTE'));
  end;

  {Gestion éventuelle d'un solde de réinitialisation}
  if IsColReInit(0) then begin
    GererSoldeInit(txConv);
    Premier := False;
  end;

  {Suppression du dernier séparateur}
  FreeAndNil(TobG);

  {-------------- 2° ÉTAPE : -----------------
   Remplissage des montants pour les autres dates : on boucle sur chaque compte/ banque et
   pour chacun, on récupère les opérations et les soldes pour chaque jour de la période de
   traitemnt}
  for n := 0 to lIndice.Count - 1 do begin
    TobGrid.Detail.First;
    S := 'TYPE;CODE;';
    C := '@';
    D := lIndice[n];
    TobM := TobGrid.FindFirst(['TYPE', 'CODE'], [C, D], True);
    C := '+';
    TobG := TobGrid.FindFirst(['TYPE', 'CODE'], [C, D], True);
    C := '-';
    TobS := TobGrid.FindFirst(['TYPE', 'CODE'], [C, D], True);
    {J'espère que cela n'arrivera pas}
    if not Assigned(TobM) or not Assigned(TobG) or not Assigned(TobS) then Continue;

    for i := 1 to NbColonne - 1 do begin
      {1/ Gestion des soldes au matin : reprise du solde de la veille au soir}
      C := RetourneCol(i);
      D := RetourneCol(i - 1);
      TobM.PutValue(C, TobS.getvalue(D));

      {2/ Gestion des soldes après synchronisation : Solde du matin + écritures comptables}
      if Periodicite = tp_30 then
        TobL := TobListe.FindFirst(['CODEBANQUE', 'TE_QUALIFORIGINE', 'TE_ANNEE', 'TE_MOIS'], [lIndice[n], QUALIFCOMPTA, AnneeFromCol(i), MoisFromCol(i)], True)
      else
        TobL := TobListe.FindFirst(['CODEBANQUE', 'TE_QUALIFORIGINE', sDateOpe], [lIndice[n], QUALIFCOMPTA, DateDe + i], True);

      if Assigned(TobL) then begin
        Montant := TobL.getvalue('MONTANT') * txConv;
        {Il y a eu des mouvements ce jour, on mémorise donc les références pour le DblClick}
        ListeSolde.Add(RetourneCol(i) + ';+;' + lIndice[n]);
      end
      else
        Montant := 0;
      TobG.PutValue(C, Montant + TobM.getvalue(C));

      {3/ Gestion des soldes au soir : solde après synchro + écritures de tréso}
      if Periodicite = tp_30 then
        TobL := TobListe.FindFirst(['CODEBANQUE', 'TE_QUALIFORIGINE', 'TE_ANNEE', 'TE_MOIS'], [lIndice[n], QUALIFTRESO, AnneeFromCol(i), MoisFromCol(i)], True)
      else
        TobL := TobListe.FindFirst(['CODEBANQUE', 'TE_QUALIFORIGINE', sDateOpe], [lIndice[n], QUALIFTRESO, DateDe + i], True);

      if Assigned(TobL) then begin
        Montant := TobL.getvalue('MONTANT') * txConv;
        {Il y a eu des mouvements ce jour, on mémorise donc les références pour le DblClick}
        ListeSolde.Add(RetourneCol(i) + ';-;' + lIndice[n]);
      end
      else
        Montant := 0;

      {Solde du soir si l'on n'est pas sur un début de millésime}
      if not IsColReInit(i) then begin
        {Solde du soir : on reprend le solde du "midi" auquel on ajoute éventuellement les écritures
         de trésorerie}
        TobS.PutValue(C, Montant + TobG.getvalue(C));
      end
      {Solde du soir en début de millésime. Remarque, le GererSoldeInit calcul tous les soldes de millésimes
       en une seule fois, d'où le boolean "Premier" : pour toutes les autres banques / comptes, on n'écrira
       rien le 01/01 car cela aura été fait !!}
      else if Premier and IsColReInit(i) then begin
        {Gestion éventuelle d'un solde de réinitialisation}
        GererSoldeInit(txConv);
        Premier := False;
      end;
    end; {for i := 1}
  end; {for n := 0}

  LibereListe(lIndice, False);
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRSUIVITRESO.OnUpdate ;
{---------------------------------------------------------------------------------------}
var
  Q      : TQuery;
  SQL    : string;
  I      : Integer;
  sCpt,
  Tmp1   : string;
  TmpREI : string;
  wBanque: string;
  wDate  : string;
  sNat   : string;
  sDate  : string; {13/11/06 : Gestion mensuelle}
  gDate  : string;
  wNoDos : string; {07/02/07}
begin
  inherited ;
  DetailCpt := GetCheckBoxState('CKCOMPTE') in [cbChecked];

  if Periodicite = tp_30 then begin
    sDate :=  'YEAR(' + sDateOpe + ') TE_ANNEE, MONTH(' + sDateOpe + ') TE_MOIS, ';
    gDate :=  'YEAR(' + sDateOpe + '), MONTH(' + sDateOpe + ') ';
  end
  else begin
    sDate :=  sDateOpe + ', ';
    gDate :=  sDateOpe;
  end;

  sNat := GetControlText('MCNATURE');
  if (Pos('<<' , sNat) = 0) and (Trim(sNat) <> '') then
    Nature := 'AND TE_NATURE IN (' + GetClauseIn(sNat) + ')'
  else
    Nature := '';

  wDate  := sDateOpe + ' >= "' + UsDateTime(DateDepart) + '" AND ' + sDateOpe + ' <= "' + UsDateTime(DateFin) + '" ';

  TmpREI := Nature;

  {20/11/03 : Gestion de la nouvelle clause where sur les banques}
  if cbBanque.Value <> '' then begin
    wBanque := 'AND BQ_BANQUE IN (';
    Tmp1 := GetClauseIn(cbBanque.Value);
    wBanque := wBanque + Tmp1 + ') ';
  end
  else
    wBanque := '';

  Tmp1 := wBanque + sCpt + TmpREI;

  {Gestion éventuelle du regroupement multi sociétés}
  TmpREI := GetWhereDossier;
  if TmpREI <> '' then
    Tmp1 := Tmp1 + ' AND ' + TmpREI;

  {07/02/07 : On filtre banquecp sur les NoDossier du regroupement}  
  wNoDos := FiltreBanqueCp(tcp_Tous, '', GetControlText('REGROUPEMENT'));
  if wNoDos <> '' then wNoDos := ' AND ' + wNoDos;
  wNoDos := wNoDos;

  {Cependant l'ajout de la clause date présente l'avantage de très largement limité la requête et
   l'inconvénient de ne pas faire figurer les comptes non mouvementés sur la période : à méditer.
   Ici on se moque des lignes de réinitialisation, else seront gérées à la main}
  Tmp1 := ' WHERE ' + wDate + Tmp1 {+ 'AND TE_QUALIFORIGINE <> "' + CODEREINIT + '"'};

  if DetailCpt then begin                                          { FQ 10328 }
    SQL := 'SELECT TE_GENERAL CODEBANQUE, ' + sDate + ' SUM(' + ChpMontant + ') MONTANT, TE_QUALIFORIGINE, ' +
           ' BQ_LIBELLE LIBELLE, BQ_BANQUE FROM TRECRITURE' +
           ' LEFT JOIN BANQUECP ON TE_GENERAL = BQ_CODE' + Tmp1 + wNoDos + 
           ' GROUP BY TE_QUALIFORIGINE, ' + gDate + ', TE_GENERAL, BQ_LIBELLE, BQ_BANQUE';

    if rbBanque.Checked then
      SQL := SQL + ' ORDER BY BQ_BANQUE, TE_GENERAL, TE_QUALIFORIGINE, ' + gDate
    else
      SQL := SQL + ' ORDER BY TE_QUALIFORIGINE, BQ_BANQUE, TE_GENERAL, ' + gDate;
    {15/11/06 : FQ 10302 : Éventuelle création de la liste des comptes}
    CreerListeCompte;
  end
  else begin                                { FQ 10328 }
    SQL := 'SELECT ' + sDate + ' SUM(' + ChpMontant + ') MONTANT, TE_QUALIFORIGINE, ' +
           ' PQ_LIBELLE LIBELLE, BQ_BANQUE CODEBANQUE FROM TRECRITURE' +
           ' LEFT JOIN BANQUECP ON TE_GENERAL = BQ_CODE' + 
           ' LEFT JOIN BANQUES ON BQ_BANQUE = PQ_BANQUE' + Tmp1 + wNoDos +
           ' GROUP BY TE_QUALIFORIGINE, ' + gDate + ', BQ_BANQUE, PQ_LIBELLE';

    if rbBanque.Checked then
      SQL := SQL + ' ORDER BY PQ_LIBELLE, TE_QUALIFORIGINE, ' + gDate
    else
      SQL := SQL + ' ORDER BY TE_QUALIFORIGINE, PQ_LIBELLE, ' + gDate;
    {15/11/06 : FQ 10302 : Éventuelle création de la liste des comptes}
    CreerListeCompte(True);
  end;

  Q := OpenSQL(SQL, True);
  TobListe.LoadDetailDB('1', '', '', Q, False, True);
  Ferme(Q);

  {16/11/06 : FQ 10328 : Gestion des devises}
  txConv := 1;
  if ChpMontant = 'TE_MONTANT' then
    txConv := RetContreTaux(V_PGI.DateEntree, V_PGI.DevisePivot, Devise);

  if rbBanque.Checked then AfficherCompte(DateDepart)
                      else AfficherSoldes(DateDepart);

  Grid.ColCount := COL_DATEDE + NbColonne;

  {Formatage des montant : '#,##0.00'; Euros}
  SQL := StrFMask(CalcDecimaleDevise(Devise), '', True);
  {Formatage des dates}
  for I := COL_DATEDE to COL_DATEDE + NbColonne - 1 do
    Grid.ColFormats[I] := SQL;

  TobGrid.PutGridDetail(Grid, True, True, '', True);

  Grid.Cells[COL_LIBELLE, 0] := '';

  {Gestion du séparateur de ligne par solde entre chaque banque}
  for i := 0 to Grid.RowCount - 1 do
    if Grid.Cells[COL_TYPE, i] = '$' then Grid.RowHeights[i] := 3;

  SQL := 'Fiche de suivi Comptabilité / Trésorerie';
  if not MultiNature then begin
         if Pos('"' + na_Realise    + '"', Nature) > 0 then LibNature := ' réalisés'
    else if Pos('"' + na_Simulation + '"', Nature) > 0 then LibNature := ' de simulation'
                                                       else LibNature := ' prévisionnels';
    SQL := SQL + ' (flux' + LibNature + ')';
  end
  else
    LibNature := '';

  Ecran.Caption := TraduireMemoire(SQL);
  UpdateCaption(Ecran);

  Grid.Col := COL_DATEDE;
  Grid.Row := Grid.FixedRows;
  Grid.SetFocus;
end ;

{---------------------------------------------------------------------------------------}
procedure TOF_TRSUIVITRESO.CritereClick(Sender : TObject);
{---------------------------------------------------------------------------------------}
begin
  inherited;
  SetControlVisible('BGRAPH', not rbBanque.Checked);
  SetControlVisible('BGRAPH1', not rbBanque.Checked);
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRSUIVITRESO.bGraphClick(Sender : TObject);
{---------------------------------------------------------------------------------------}
var
  Titre  : string;
  sSolde : string;
  sDev   : string;
  n, p   : Integer;
  T, D   : TOB;
  sGraph : string;
begin
  Titre := 'Évolution des soldes pour la période du ' + DateToStr(ColToDate(0)) + ' au ' + DateToStr(ColToDate(NbColonne) - 1);

  if rbDateOpe.Checked then sSolde := 'Les montants sont calculés en date d''opération'
                       else sSolde := 'Les montants sont calculés en date de valeur';
  sDev  := 'Les montants sont éxprimés en ' + cbDevise.Items[cbDevise.ItemIndex];

  {Constitution de la tob des flux}
  T := Tob.Create('$$$', nil, -1);
  try
    {Chaine contenant le titre des séries}
    sGraph := 'DATE;SOLDEMAT;SOLDESYN;SOLDETRE;';

    for n := COL_DATEDE to Grid.ColCount - 1 do begin
      D := Tob.Create('$$$', T, -1);
      D.AddChampSup('DATE', False);
      D.PutValue('DATE', Grid.Cells[n, 0]);
      for p := 1 to Grid.RowCount - 1 do begin
        {Les lignes contenant les soldes des comptes contiennent dans la première colonne le libellé
         de leur banque et non "-", "+", "*", "" => le teste sur la longueur}
        if Grid.Cells[COL_TYPE, p] = '§' then
          D.AddChampSupValeur('SOLDEMAT', Grid.Cells[n, p])
        else if Grid.Cells[COL_TYPE, p] = '£' then
          D.AddChampSupValeur('SOLDESYN', Grid.Cells[n, p])
        else if Grid.Cells[COL_TYPE, p] = 'µ' then
          D.AddChampSupValeur('SOLDETRE', Grid.Cells[n, p]);
      end;
    end;

    {!!! Mettre sGraph en fin à cause des ";" qui traine dans la chaine}
    TRLanceFiche_TRGRAPHTROCTA(Titre + ';' + sDev + ';' + sSolde + ';' + sGraph, T);
    T.ClearDetail;
  finally
    FreeAndNil(T);
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRSUIVITRESO.bGraph1Click(Sender : TObject);
{---------------------------------------------------------------------------------------}
var
  Titre  : string;
  sSolde : string;
  sDev   : string;
  n      : Integer;
  T, D, L: TOB;
  sGraph : string;
begin
  Titre := 'Cumuls des opérations pour la période du ' + DateToStr(ColToDate(0)) + ' au ' + DateToStr(ColToDate(NbColonne) - 1);

  if rbDateOpe.Checked then sSolde := 'Les montants sont calculés en date d''opération'
                       else sSolde := 'Les montants sont calculés en date de valeur';
  sDev  := 'Les montants sont éxprimés en ' + cbDevise.Items[cbDevise.ItemIndex];

  {Constitution de la tob des flux}
  T := Tob.Create('$$$', nil, -1);
  try
    {Chaine contenant le titre des séries}
    sGraph := 'DATE;CUMTRE;CUMCPT;';

    for n := 0 to NbColonne - 1 do begin
      D := Tob.Create('$$$', T, -1);
      D.AddChampSupValeur('DATE', RetourneCol(n));
      if Periodicite = tp_30 then
        L := TobListe.FindFirst(['TE_QUALIFORIGINE', 'TE_ANNEE', 'TE_MOIS'], [QUALIFTRESO, AnneeFromCol(n), MoisFromCol(n)], True)
      else
        L := TobListe.FindFirst(['TE_QUALIFORIGINE', sDateOpe], [QUALIFTRESO, ColToDate(n)], True);
      if L <> nil then D.AddChampSupValeur('CUMTRE', L.getvalue('MONTANT'))
                  else D.AddChampSupValeur('CUMTRE', 0);

      if Periodicite = tp_30 then
        L := TobListe.FindFirst(['TE_QUALIFORIGINE', 'TE_ANNEE', 'TE_MOIS'], [QUALIFCOMPTA, AnneeFromCol(n), MoisFromCol(n)], True)
      else
        L := TobListe.FindFirst(['TE_QUALIFORIGINE', sDateOpe], [QUALIFCOMPTA, ColToDate(n)], True);
      if L <> nil then D.AddChampSupValeur('CUMCPT', L.getvalue('MONTANT'))
                  else D.AddChampSupValeur('CUMCPT', 0);
    end;

    {!!! Mettre sGraph en fin à cause des ";" qui traine dans la chaine}
    TRLanceFiche_TRGRAPHTROCTA(Titre + ';' + sDev + ';' + sSolde + ';' + sGraph, T);
    T.ClearDetail;
  finally
    FreeAndNil(T);
  end;
end;

{---------------------------------------------------------------------------------------}
function TOF_TRSUIVITRESO.GetWhereBqCpt : string;
{---------------------------------------------------------------------------------------}
var
  Prefix : string;
begin
  inherited GetWhereBqCpt;
  if cbBanque.Value <> '' then begin
    if DetailCpt then Prefix := 'BQ'
                 else Prefix := 'PQ';

    Result := Prefix + '_BANQUE IN (';
    Result := Result + GetClauseIn(cbBanque.Value);
    Result := Result + ')';
  end
end;

{13/11/06 : Adaptation à la gestion mensuelle
{---------------------------------------------------------------------------------------}
function TOF_TRSUIVITRESO.GetObjSuiv(CodBq, TitreDt, LibBq : string) : TObjSuiv;
{---------------------------------------------------------------------------------------}
var
  n : Integer;
begin
  n := lIndice.IndexOf({TitreDt + ';' + }CodBq);
  if (n = -1) or not Assigned(lIndice.Objects[n]) then begin
    Result := TObjSuiv.Create;
    Result.LibBque := LibBq;
    Result.CumSynch := 0;
    Result.CumTreso := 0;
    lIndice.AddObject({TitreDt + ';' + }CodBq, Result);
  end
  else
    Result := TObjSuiv(lIndice.Objects[n]);
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRSUIVITRESO.SetMoisAnne(DateRef: TDateTime);
{---------------------------------------------------------------------------------------}
var
  a, m, j : Word;
begin
  DecodeDate(DateRef, a, m, j);
  AnneeDe := a;
  MoisDe  := m;
end;

{---------------------------------------------------------------------------------------}
function TOF_TRSUIVITRESO.AnneeFromCol(n : Integer) : Integer;
{---------------------------------------------------------------------------------------}
var
  Dt : TDateTime;
  a, m, j : Word;
begin
  Dt := ColToDate(n);
  DecodeDate(Dt, a, m, j);
  Result := a;
end;

{---------------------------------------------------------------------------------------}
function TOF_TRSUIVITRESO.MoisFromCol(n : Integer) : Integer;
{---------------------------------------------------------------------------------------}
var
  Dt : TDateTime;
  a, m, j : Word;
begin
  Dt := ColToDate(n);
  DecodeDate(Dt, a, m, j);
  Result := m;
end;

{---------------------------------------------------------------------------------------}
function TOF_TRSUIVITRESO.IsColReInit(n : Integer) : Boolean;
{---------------------------------------------------------------------------------------}
begin
  Result := False;
  if n < 0 then Exit;
  if n = 0 then begin
    {          On part de janvier ou du 01/01}
    Result := (DateDepart = DebutAnnee(DateDepart)) or
    {          On part du 31/12}
              ((Periodicite = tp_1) and (DateDepart = FinAnnee(DateDepart))) or
    {          On part de décembre}
              ((Periodicite = tp_30) and (FinDeMois(DateDepart) = FinAnnee(DateDepart)));
  end
  else if Periodicite = tp_1 then
    Result := (DateDepart + n) = FinAnnee(DateDepart + n)
  else
    Result := (FinDeMois(ColToDate(n)) = FinAnnee(ColToDate(n))) and (n < NbColonne - 1);
end;

{30/11/06 : Comme la fonction GetSoldeBanque renvoie systématiquement le solde en devise Pivot, même si la sélection
 n'est pas en devise pivot et qu'elle correspond à la devise d'affichage (ce qui fait que txConv = 1), il faut gérer
 la devise spécifiquement
{---------------------------------------------------------------------------------------}
function TOF_TRSUIVITRESO.GereSoldeBanque(Banque, Endate, Nature : string; EnValeur, AgenceOk : Boolean; NoDossier : string) : Double;
{---------------------------------------------------------------------------------------}
begin
  {Récupèration du solde en devise pivot}
  Result := GetSoldeBanque(Banque, Endate, Nature, EnValeur, AgenceOk, NoDossier);
  {Si, on n'est pas en devise pivot et que le taux = 1, c'est que
   1/ soit tous les comptes sont dans la devises d'affichage
   2/ soit qu'il n'y a pas de saisie de taux}
  if not MntPivot and (txConv = 1) then
    Result := Result * RetContreTaux(V_PGI.DateEntree, V_PGI.DevisePivot, Devise)
  else
    Result := Result * txConv;
end;

initialization
  RegisterClasses([TOF_TRSUIVITRESO]);

end.



