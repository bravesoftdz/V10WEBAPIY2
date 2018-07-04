{-------------------------------------------------------------------------------------
    Version   |   Date  | Qui |   Commentaires
--------------------------------------------------------------------------------------
 0.91          15/10/03   JP   Création de l'unité
 6.00.014.001  17/09/04   JP   FQ 10128 : on cache les boutons inutiles
 7.05.001.001  24/10/06   JP   Gestion des filtres multi sociétés
 8.00.001.018  05/06/07   JP   FQ 10469 : désactvation des filtres et des boutons
--------------------------------------------------------------------------------------}

unit TRGRAPHSOLDE_TOF ;

interface

uses {$IFDEF VER150} variants,{$ENDIF}
  StdCtrls, Controls, Classes, Graphics, TeEngine, Chart, Windows,
  {$IFDEF EAGLCLIENT}
  MaineAGL,
  {$ELSE}
  FE_Main, db, {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF}
  {$ENDIF}
  SysUtils, HCtrls, HEnt1, GRS1, Series, GraphUtil, UTOF, UTob, UObjGen, ExtCtrls;


type
  TOF_TRGRAPHSOLDE = Class (TOF)
    procedure OnUpdate              ; override;
    procedure OnClose               ; override;
    procedure OnArgument(S : string); override;
  private
    ObjDevise : TObjDevise;
    TobSeries : TOB;
    UneSerie  : Boolean; {Si une comparaison avec une autre période est demandée}
    SoldeOk   : Boolean; {Si c'est un solde qui est demandé}
    SensDiff  : Boolean; {S'il y a des montants positifs et négatifs}
    function  RecupClauseIn (SansJointure : Boolean = False) : string;
    procedure RemplirTob;
    procedure RecupRequete  (var T : TOB; var TotalN, TotalN1 : Double);
    procedure RecupererTOB  (var T : TOB; var TotalN, TotalN1 : Double);
    procedure PeriodeOnClick(Sender : TObject);
    procedure CompareOnClick(Sender : TObject);
    procedure DeviseOnChange(Sender : TObject);
    procedure DessineLigne  (Sender : TObject);
    procedure MajLegendN    (Sender : TChartSeries; ValueIndex : Longint ; var MarkText : hString);
    procedure MajLegendN1   (Sender : TChartSeries; ValueIndex : Longint ; var MarkText : hString);
  end ;

procedure TRLanceFiche_HistoSolde(Dom, Fiche, Range, Lequel, Arguments : string);

implementation

uses
  Commun, Constantes, UProcGen, UObjFiltres;

{---------------------------------------------------------------------------------------}
procedure TRLanceFiche_HistoSolde(Dom, Fiche, Range, Lequel, Arguments : string);
{---------------------------------------------------------------------------------------}
begin
  AGLLanceFiche(Dom, Fiche, Range, Lequel, Arguments);
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRGRAPHSOLDE.OnUpdate ;
{---------------------------------------------------------------------------------------}
var
  F : TFGRS1;
  stTitre,
  stColonnes, stChampTitre,
  stCriteres, stWhere,
  stTitresCol1, stTitresCol2,
  stColGraph1, stColGraph2 : string;
  tstTitres                : Tstrings;
  ft                       : TFont;
begin
  inherited;
  RemplirTob;

  stCriteres := '';
  stColGraph1 := '';
  stColGraph2 := '';
  stTitresCol1 := '';
  stTitresCol2 := '';
  stTitre := '';

  tstTitres := TStringList.Create;
  ft := TFont.Create;
  try
    {Definition des fonts des titres}
    ft.Style := [fsBold];
    ft.Color := clBlue;
    ft.Size  := 12;

    F := TFGRS1(Ecran);
    if SoldeOk then TstTitres.Add ('Soldes des comptes bancaires')
               else TstTitres.Add ('Cumuls des mouvements bancaires');

    stColonnes := 'COMPTE;MONTANTN';
    stTitre := 'Compte;Montant';
    if not UneSerie then begin
      stColonnes := stColonnes  + '1;MONTANTN';
      F.FListe.ColCount := 3;
      stTitre := stTitre  + ' préc.;Montant';
    end
    else
      F.FListe.ColCount := 2;
    stColGraph1 := stColonnes;
    stColGraph2 := '';
    stChampTitre := stTitre;
    LanceGraph(F, TobSeries, '', stColonnes, stWhere , stTitre , stColGraph1,
                 stColGraph2, tstTitres, nil, TBarSeries , stChampTitre, False);

    if F.Fchart1.SeriesCount > 0 then begin
      {Évènements qui permet d'écrire au sommet des barres le pourcentage}
      if UneSerie then
        F.FChart1.SeriesList.Series[1].OnGetMarkText := MajLegendN
      else begin
        F.FChart1.SeriesList.Series[2].OnGetMarkText := MajLegendN;
        F.FChart1.SeriesList.Series[1].OnGetMarkText := MajLegendN1;
      end;

      {Pour dessiner des lignes horizontales sur le graphe}
      F.FChart1.SeriesList.Series[1].BeforeDrawValues := DessineLigne;

      {Les comptes ne servent sue de marques, on ne les affiche pas dans le graphe}
      F.FChart1.SeriesList.Series[0].Marks.Visible := False;
      F.FChart1.SeriesList.Series[0].DataSource := nil;

      {Titre des axes}
      F.FChart1.LeftAxis.Title.Caption := 'Montants';
      F.FChart1.BottomAxis.Title.Caption := 'Comptes';

      {Titre des légendes}
      if UneSerie then begin
        if SoldeOk then F.FChart1.SeriesList.Series[1].Title := 'Solde du ' + GetControlText('EDDATEAU')
                   else F.FChart1.SeriesList.Series[1].Title := 'Cumul du ' + GetControlText('EDDATEDEB') + ' au ' + GetControlText('EDDATEFIN');
      end
      else begin
        if SoldeOk then F.FChart1.SeriesList.Series[2].Title := 'Solde du ' + GetControlText('EDDATEAU')
                   else F.FChart1.SeriesList.Series[2].Title := 'Cumul du ' + GetControlText('EDDATEDEB') + ' au ' + GetControlText('EDDATEFIN');
        if SoldeOk then F.FChart1.SeriesList.Series[1].Title := 'Solde du ' + GetControlText('EDCOMPARE1')
                   else F.FChart1.SeriesList.Series[1].Title := 'Cumul du ' + GetControlText('EDCOMPARE1') + ' au ' + GetControlText('EDCOMPARE2');

      end;

      {Pour une question de place, on met la légende en bas de page}
      F.FChart1.Legend.Alignment := laBottom;

      {Pour ne pas afficher les comptes dans la légende : la série 0 concerne les comptes, on s'en moque dans la légende}
      F.FChart1.Legend.FirstValue := 1;
      F.FChart1.LeftAxis.LabelsFont.Color := clRed;

      {Font des titres des axes du graph}
      F.FChart1.LeftAxis  .Title.Font.Assign(ft);
      F.FChart1.BottomAxis.Title.Font.Assign(ft);
      F.FChart1.LeftAxis  .Title.Font.Color := clRed;

      ft.Size := 9;
      ft.Style := [];
      {Font des échelles des axes}
      F.FChart1.LeftAxis  .LabelsFont.Assign(ft);
      F.FChart1.BottomAxis.LabelsFont.Assign(ft);
      F.FChart1.LeftAxis  .LabelsFont.Color := clRed;
    end;
  finally
    FreeAndNil(ft);
    tstTitres.Free;
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRGRAPHSOLDE.OnArgument (S : String ) ;
{---------------------------------------------------------------------------------------}
begin
  inherited;
  Ecran.HelpContext := 150;
  SetControlVisible('BUNDO'  , False); {FQ 10128}
  SetControlVisible('BSAUVE' , False); {FQ 10128}
  SetControlVisible('BDELETE', False); {FQ 10128}

  ObjDevise := TObjDevise.Create(V_PGI.DateEntree);
  TCheckBox   (GetControl('CKCOMPARE')).OnClick := CompareOnClick;
  TRadioButton(GetControl('RBSOLDE'  )).OnClick := PeriodeOnClick;
  TRadioButton(GetControl('RBVOLUME' )).OnClick := PeriodeOnClick;
  SetControlText('CBDEVISE', V_PGI.DevisePivot);
  THValComboBox(GetControl('CBDEVISE')).OnChange := DeviseOnChange;
  TobSeries := TOB.Create('***', nil, -1);
  UneSerie := True;
  SoldeOk  := True;
  {24/10/06 : On filtre les comptes en fonction des sociétés du regroupement Tréso}
  THMultiValComboBox(GetControl('MCCOMPTE')).Plus := FiltreBanqueCp(THMultiValComboBox(GetControl('MCCOMPTE')).DataType, '', '');

  {05/06/07 : FQ 10469 : on cache les filtres}
  CacheFiltreGraph(Ecran);
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRGRAPHSOLDE.OnClose;
{---------------------------------------------------------------------------------------}
begin
  FreeAndNil(TobSeries);
  if Assigned(ObjDevise) then FreeAndNil(ObjDevise);
  inherited;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRGRAPHSOLDE.PeriodeOnClick(Sender : TObject);
{---------------------------------------------------------------------------------------}
begin
  SoldeOk := UpperCase(TComponent(Sender).Name) = 'RBSOLDE';

  SetControlEnabled('EDDATEAU', SoldeOk);
  SetControlEnabled('EDDATEDEB', not SoldeOk);
  SetControlEnabled('LBAU'     , not SoldeOk);
  SetControlEnabled('EDDATEFIN', not SoldeOk);
  SetControlVisible('LBCOMPAREAU', not SoldeOk);
  SetControlVisible('EDCOMPARE2' , not SoldeOk);
  if SoldeOk then SetControlText('LBCOMPARE', 'Avec le solde du')
             else SetControlText('LBCOMPARE', 'Avec la période du');
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRGRAPHSOLDE.CompareOnClick(Sender : TObject);
{---------------------------------------------------------------------------------------}
var
  CompOk : Boolean;
begin
  CompOk := TCheckBox(Sender).Checked;
  SetControlEnabled('LBCOMPAREAU', CompOk);
  SetControlEnabled('LBCOMPARE'  , CompOk);
  SetControlEnabled('EDCOMPARE1' , CompOk);
  SetControlEnabled('EDCOMPARE2' , CompOk);
  {Variable de classe pour savoir si on gère une ou deux séries}
  UneSerie := not CompOk;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRGRAPHSOLDE.DeviseOnChange(Sender : TObject);
{---------------------------------------------------------------------------------------}
begin
  AssignDrapeau(TImage(GetControl('IDEV')), THValComboBox(GetControl('CBDEVISE')).Value);
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRGRAPHSOLDE.RemplirTob;
{---------------------------------------------------------------------------------------}
var
  T : TOB;
  Q : TOB;
  n : Integer;
  SN  : Double;
  SN1 : Double;
begin
  TobSeries.Detail.Clear;
  Q := TOB.Create('***', nil, -1);
  try
    {S'il y a un filtre sur la nature et que l'on demande les soldes, le traitement est différent car
     il va falloir tenir compte des écritures d'initialisation et recalculer les soldes à la volée}
    if (GetControlText('MCNATURE') = '') or (Pos('<<', GetControlText('MCNATURE')) > 0) or not SoldeOk then
      RecupRequete(Q, SN, SN1)
    else
      RecupererTOB(Q, SN, SN1);

    for n := 0 to Q.Detail.Count - 1 do begin
      T := TOB.Create('***', TobSeries, -1);
      T.AddChampSupValeur('COMPTE'    , Q.Detail[n].GetValue('TE_GENERAL'));
      T.AddChampSupValeur('MONTANTN'  , FormatFloat('#,##0.00', Valeur(Q.Detail[n].GetValue('MONTANTN'))));
      T.AddChampSupValeur('MONTANTN1' , FormatFloat('#,##0.00', Valeur(Q.Detail[n].GetValue('MONTANTN1'))));
      if (SN <> 0) then T.AddChampSupValeur('POURCENTN', FormatFloat('#,##0.00', Abs(Valeur(Q.Detail[n].GetValue('MONTANTN'))) / SN * 100))
                   else T.AddChampSupValeur('POURCENTN', '0.00');
      if not UneSerie and (SN1 <> 0)then
        T.AddChampSupValeur('POURCENTN1', FormatFloat('#,##0.00', Abs(Valeur(Q.Detail[n].GetValue('MONTANTN1'))) / SN1 * 100))
      else
        T.AddChampSupValeur('POURCENTN1', '0.00');
    end;
  finally
    FreeAndNil(Q);
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRGRAPHSOLDE.RecupererTOB(var T : TOB; var TotalN, TotalN1 : Double);
{---------------------------------------------------------------------------------------}
var
  SQL : string;
  Q   : TQuery;
  QQ  : TQuery;
  O   : TOB;
  n   : Integer;
  Tmp : Double;
  Cot : Double;
  Tau : Double;
  Pos : Boolean;
  Neg : Boolean;
  Dev : string;
begin
  TotalN  := 0;
  TotalN1 := 0;
  Tmp     := 0;
  Pos     := False;
  Neg     := False;
  Dev     := THValComboBox(GetControl('CBDEVISE')).Value;
  {Récupération de la parité de la devise contre l'euro}
  Tau := RetPariteEuro(THValComboBox(GetControl('CBDEVISE')).Value, V_PGI.DateEntree);
  if Tau = 0 then Tau := 1;
  SQL := '';

  {Recherche des montants pour la colonne N}
  SQL := 'SELECT MAX(BQ_LIBELLE), SUM(TE_MONTANT), TE_GENERAL FROM TRECRITURE, BANQUECP WHERE TE_DATEVALEUR > "' +
               UsDateTime(DebutAnnee(StrToDate(GetControlText('EDDATEAU')))) +
               '" AND TE_DATEVALEUR <= "' + UsDateTime(StrToDate(GetControlText('EDDATEAU'))) + '" ' +
               RecupClauseIn + ' GROUP BY TE_GENERAL ORDER BY TE_GENERAL';

  Q := OpenSQL(SQL, True);
  try
    {Récupération des soldes d'initialisation}
    SQL := 'SELECT TE_SOLDEDEVVALEUR, TE_GENERAL FROM TRECRITURE WHERE TE_CODEFLUX = "' +
                   CODEREGULARIS + '" ' + RecupClauseIn(True) + 'AND TE_DATEVALEUR = "' +
                   UsDateTime(DebutAnnee(StrToDate(GetControlText('EDDATEAU')))) + '" ORDER BY TE_GENERAL';

    QQ := OpenSQL(SQL, True);
    {Calcul du total pour la colonne N et conversion des montants dans la devise d'affichage}
    Q.First;
    while not Q.EOF do begin
      {Création des tobs filles}
      O := TOB.Create('****', T, -1);
      O.AddChampSupValeur('TE_GENERAL', Q.Fields[0].AsString);
      Tmp := 0;

      {On recherche le montant de réinitialisation du compte en cours}
      if QQ.Locate('TE_GENERAL', Q.Fields[2].AsString, [loCaseInsensitive]) then begin
        {Conversion du montant dans la devise d'affichage}
        Tmp := QQ.Fields[0].AsFloat;
        Cot := ObjDevise.GetParite(ObjDevise.GetDeviseCpt(Q.Fields[2].AsString), Dev);
        if Cot <> 0 then Tmp := Tmp * Cot;
      end;

      Tmp := Tmp + Q.Fields[1].AsFloat / Tau;

      O.AddChampSupValeur('MONTANTN', Tmp);
      TotalN := TotalN + Abs(Tmp);

      {Gestion des différences de signes}
      if not Neg then Neg := Tmp < 0;
      if not Pos then Pos := Tmp > 0;

      Q.Next;
    end;
  finally
    if Assigned(QQ) then Ferme(QQ);
    if Assigned(Q ) then Ferme(Q);
  end;

  {Si une comparaison est demandée, recherche de montants pour la colonne N1}
  if not UneSerie then begin
    {Recherche des montants pour la colonne N1}
    SQL := 'SELECT MAX(BQ_LIBELLE), SUM(TE_MONTANT), TE_GENERAL FROM TRECRITURE, BANQUECP WHERE TE_DATEVALEUR > "' +
                 UsDateTime(DebutAnnee(StrToDate(GetControlText('EDCOMPARE1')))) +
                 '" AND TE_DATEVALEUR <= "' + UsDateTime(StrToDate(GetControlText('EDCOMPARE1'))) + '" ' +
                 RecupClauseIn + ' GROUP BY TE_GENERAL ORDER BY TE_GENERAL';

    Q := OpenSQL(SQL, True);
    try
      {Récupération des soldes d'initialisation}
      SQL := 'SELECT TE_SOLDEDEVVALEUR, TE_GENERAL FROM TRECRITURE WHERE TE_CODEFLUX = "' +
                     CODEREGULARIS + '" ' + RecupClauseIn(True) + 'AND TE_DATEVALEUR = "' +
                     UsDateTime(DebutAnnee(StrToDate(GetControlText('EDCOMPARE1')))) + '" ORDER BY TE_GENERAL';

      QQ := OpenSQL(SQL, True);

      {Calcul du total pour la colonne N1 et conversion des montants dans la devise d'affichage}
      Q.First;
      while not Q.EOF do begin
        {Ajout des montants N1 dans la TOB}
        for n := 0 to T.Detail.Count - 1 do begin
          if T.Detail[n].GetValue('TE_GENERAL') = Q.Fields[0].AsString then begin
            Tmp := 0;

            {On recherche le montant de réinitialisation du compte en cours}
            if QQ.Locate('TE_GENERAL', Q.Fields[2].AsString, [loCaseInsensitive]) then begin
              {Conversion du montant dans la devise d'affichage}
              Tmp := QQ.Fields[0].AsFloat;
              Cot := ObjDevise.GetParite(ObjDevise.GetDeviseCpt(Q.Fields[2].AsString), Dev);
              if Cot <> 0 then Tmp := Tmp * Cot;
            end;

            Tmp := Tmp + Q.Fields[1].AsFloat / Tau;


            T.Detail[n].AddChampSupValeur('MONTANTN1', Tmp);
            Break;
          end;
        end;
        TotalN1 := TotalN1 + Abs(Tmp);
        Q.Next;

        {Gestion des différences de signes}
        if not Neg then Neg := Tmp < 0;
        if not Pos then Pos := Tmp > 0;
      end;
    finally
      if Assigned(QQ) then Ferme(QQ);
      if Assigned(Q ) then Ferme(Q);
    end;
  end;

  {Il ne peut y avoir des montants de signes différents que si un solde est demandé !!}
  SensDiff := Pos and Neg;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRGRAPHSOLDE.RecupRequete(var T : TOB; var TotalN, TotalN1 : Double);
{---------------------------------------------------------------------------------------}
var
  SQL : string;
  Q   : TQuery;
  O   : TOB;
  n   : Integer;
  Tmp : Double;
  Cot : Double;
  Tau : Double;
  Pos : Boolean;
  Neg : Boolean;
begin
  TotalN  := 0;
  TotalN1 := 0;
  Tmp     := 0;
  Pos     := False;
  Neg     := False;

  {Récupération de la parité de la devise contre l'euro}
  Tau := RetPariteEuro(THValComboBox(GetControl('CBDEVISE')).Value, V_PGI.DateEntree);
  if Tau = 0 then Tau := 1;
  SQL := '';
  {Recherche des montants pour la colonne N}
  if SoldeOk then
    SQL := 'SELECT BQ_LIBELLE, TE_SOLDEDEVVALEUR, TE_COTATION FROM TRECRITURE, BANQUECP WHERE TE_CLEVALEUR IN ' +
           '(SELECT MAX(TE_CLEVALEUR) FROM TRECRITURE WHERE TE_DATEVALEUR <= "' + UsDateTime(StrToDate(GetControlText('EDDATEAU'))) +
           '" GROUP BY TE_GENERAL) ' + RecupClauseIn + ' ORDER BY TE_GENERAL'
  else
    SQL := 'SELECT MAX(BQ_LIBELLE), SUM(ABS(TE_MONTANT)) FROM TRECRITURE, BANQUECP WHERE TE_DATEVALEUR >= "' +
                 UsDateTime(StrToDate(GetControlText('EDDATEDEB'))) +
                 '" AND TE_DATEVALEUR <= "' + UsDateTime(StrToDate(GetControlText('EDDATEFIN'))) + '" ' +
                 RecupClauseIn + ' GROUP BY TE_GENERAL ORDER BY TE_GENERAL';

  Q := OpenSQL(SQL, True);

  {Calcul du total pour la colonne N et conversion des montants dans la devise d'affichage}
  Q.First;
  while not Q.EOF do begin
    {Création des tobs filles}
    O := TOB.Create('****', T, -1);
    O.AddChampSupValeur('TE_GENERAL', Q.Fields[0].AsString);

    {Conversion du montant dans la devise d'affichage}

    {Remarque : Dans la trésorerie, on part du principe MontantEur * Cotation = MontantDev.
                Cependant, historiquement, RetPariteEuro renvoie le taux (1 Dev = x.xx €)
                et non la cotation (1€ = x.xx Dev) : ici Cot est une cotation et tau un taux !!}
    if SoldeOk then begin
      Tmp := Q.Fields[1].AsFloat;
      Cot := Q.Fields[2].AsFloat;
      if Cot = 0 then Cot := 1;
      Tmp := (Tmp / Cot) / Tau;
    end
    else
      Tmp := Q.Fields[1].AsFloat / Tau;

    O.AddChampSupValeur('MONTANTN', Tmp);
    TotalN := TotalN + Abs(Tmp);

    {Gestion des différences de signes}
    if not Neg then Neg := Tmp < 0;
    if not Pos then Pos := Tmp > 0;

    Q.Next;
  end;
  Ferme(Q);

  {Si une comparaison est demandée, recherche de montants pour la colonne N1}
  if not UneSerie then begin
    {Recherche des montants pour la colonne N1}
    if SoldeOk then
      SQL := 'SELECT BQ_LIBELLE, TE_SOLDEDEVVALEUR, TE_COTATION FROM TRECRITURE, BANQUECP WHERE TE_CLEVALEUR IN ' +
                   '(SELECT MAX(TE_CLEVALEUR) FROM TRECRITURE WHERE TE_DATEVALEUR <= "' + UsDateTime(StrToDate(GetControlText('EDCOMPARE1'))) +
                   '" GROUP BY TE_GENERAL) ' + RecupClauseIn + ' ORDER BY TE_GENERAL'
    else
      SQL := 'SELECT MAX(BQ_LIBELLE), SUM(ABS(TE_MONTANT)) FROM TRECRITURE, BANQUECP WHERE TE_DATEVALEUR >= "' +
                   UsDateTime(StrToDate(GetControlText('EDCOMPARE1'))) +
                   '" AND TE_DATEVALEUR <= "' + UsDateTime(StrToDate(GetControlText('EDCOMPARE2'))) + '" ' +
                   RecupClauseIn + ' GROUP BY TE_GENERAL ORDER BY TE_GENERAL';

    Q := OpenSQL(SQL, True);

    {Calcul du total pour la colonne N1 et conversion des montants dans la devise d'affichage}
    Q.First;
    while not Q.EOF do begin
      {Ajout des montants N1 dans la TOB}
      for n := 0 to T.Detail.Count - 1 do begin
        if T.Detail[n].GetValue('TE_GENERAL') = Q.Fields[0].AsString then begin
          {Conversion du montant dans la devise d'affichage}
          if SoldeOk then begin
            Tmp := Q.Fields[1].AsFloat;
            Cot := Q.Fields[2].AsFloat;
            if Cot = 0 then Cot := 1;
            Tmp := (Tmp / Cot) / Tau;
          end
          else
            Tmp := Q.Fields[1].AsFloat * Tau;

          T.Detail[n].AddChampSupValeur('MONTANTN1', Tmp);
          Break;
        end;
      end;
      TotalN1 := TotalN1 + Abs(Tmp);
      Q.Next;

      {Gestion des différences de signes}
      if not Neg then Neg := Tmp < 0;
      if not Pos then Pos := Tmp > 0;

    end;
    Ferme(Q);
  end;

  {Il ne peut y avoir des montants de signes différents que si un solde est demandé !!}
  SensDiff := Pos and Neg;
end;

{---------------------------------------------------------------------------------------}
function TOF_TRGRAPHSOLDE.RecupClauseIn(SansJointure : Boolean = False) : string;
{---------------------------------------------------------------------------------------}
var
  s : string;
begin
  if not SansJointure then Result := ' AND BQ_CODE = TE_GENERAL ';

  s := THMultiValComboBox(GetControl('MCCOMPTE')).Value;
  if s <> '' then
    Result := Result + ' AND TE_GENERAL IN (' + GetClauseIn(s) + ') ';

  s := THMultiValComboBox(GetControl('MCNATURE')).Value;
  if s <> '' then
    Result := Result + ' AND TE_NATURE IN (' + GetClauseIn(s) + ') ';
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRGRAPHSOLDE.MajLegendN(Sender: TChartSeries; ValueIndex : Longint ;  var MarkText : hString);
{---------------------------------------------------------------------------------------}
begin
  MarkText := VarToStr(TobSeries.Detail[ValueIndex].GetValue('PourcentN')) + ' %';
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRGRAPHSOLDE.MajLegendN1(Sender: TChartSeries; ValueIndex : Longint ;  var MarkText : hString);
{---------------------------------------------------------------------------------------}
begin
  MarkText := VarToStr(TobSeries.Detail[ValueIndex].GetValue('PourcentN1')) + ' %';
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRGRAPHSOLDE.DessineLigne(Sender : TObject);
{---------------------------------------------------------------------------------------}
var
  YPosition  : Longint;
  Ecart, Inc : Double;
  Mini, Maxi : Double;
  n, p, k    : Integer;
    {---------------------------------------------------------------------}
    procedure Dessine(Zero : Boolean);
    {---------------------------------------------------------------------}
    begin
      with TFGRS1(Ecran).FChart1, Canvas do begin
        {Définition du crayon}
        if Zero then Pen.Width := 2
                else Pen.Width := 1;
        Pen.Style := psSolid;
        Pen.Color := clBlue;
        {Dessin de la ligne}
        MoveTo(ChartRect.Left, YPosition);
        LineTo(ChartRect.Left  + Width3D, YPosition - Height3D);
        LineTo(ChartRect.Right + Width3D, YPosition - Height3D);
        {Dessin d'un bout de ligne sur la gauche de l'axe pour faire le lien avec la marque}
        MoveTo(ChartRect.Left, YPosition);
        LineTo(ChartRect.Left - 5, YPosition);
      end
    end;

begin
  P := 0;
  with TFGRS1(Ecran).FChart1 do begin
    {On ne déssine une ligne pour la valeur nulle que s'il y a des montants de sens différents}
    if SensDiff then begin
      {Récupération de la position de 0 sur l'axe vertical}
      YPosition := Series[1].CalcYPosValue(0);
      Dessine(True);
    end;

    {Dessin des 7 principales marques}
    {On commence par calculer les extrêmes de la graduation et l'incrément, car à ce moment cela
     n'a pas été encore fait : Increment = 0 et Min et Max sont les valeurs Min et Max de la série}
    LeftAxis.CalcMinMax(Mini, Maxi);
    Inc   := LeftAxis.CalcIncrement;
    Ecart := Maxi - Mini;
    {On va mettre 7 lignes plus celui de zéro}
    Ecart := Ecart / 7;
    if Inc <> 0 then Ecart := Ecart / Inc
                else Exit;
    Ecart := Round(Ecart);
    {On gradue tous les n * Ecart * Increment}
    if not  SensDiff then begin
      for n := 1 to 7 do
        {Tant que le maximun n'est pas atteint, on dessine les lignes du côté positif}
        if n * Ecart * Inc < Maxi then begin
          YPosition := Series[1].CalcYPosValue(n * Ecart * Inc);
          Dessine(False);
        end
        {Tous les montants sont négatifs}
        else if n * Ecart * Inc * -1 > Mini then begin
          YPosition := Series[1].CalcYPosValue(n * Ecart * Inc * -1);
          Dessine(False);
        end;
    end else begin
      for n := 1 to 7 do begin
        {Tant que le maximun n'est pas atteint, on dessine les lignes du côté positif}
        if n * Ecart * Inc < Maxi then begin
          YPosition := Series[1].CalcYPosValue(n * Ecart * Inc);
          Dessine(False);
          p := n;
        end
        {... sinon on passe du côté négatif}
        else begin
          for k := 1 to 7 - p do
            if k * Ecart * Inc * -1 > Mini then begin
              YPosition := Series[1].CalcYPosValue(k * Ecart * Inc * -1);
              Dessine(False);
            end;
          Exit;
        end;
      end;
    end;
  end;
end;

initialization
  RegisterClasses ( [ TOF_TRGRAPHSOLDE ] ) ;

end.
