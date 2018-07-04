{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 28/04/2005
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : CPETATSYNTHESE ()
Mots clefs ... : TOF;CPETATSYNTHESE
*****************************************************************}
unit CPETATSYNTHESE_TOF;

interface

uses StdCtrls,
  Controls,
  Classes,
  ComCtrls,
{$IFNDEF EAGLCLIENT}
  db,
  dbtables,
  mul,
  fe_main,
  dialogs,
{$ELSE}
  eMul,
  MaineAGL,
{$ENDIF}
  uTob,
  forms,
  sysutils,
  HCtrls,
  HEnt1,
  HMsgBox,
  UTOF,
  HTB97,
  ParamSoc,
  { Compta }
  Ent1,
  uLibExercice,
  CritEdt,
  GridSynth,
  TOFMETH,
  uTofLookUpTob;

type
  TOF_CPETATSYNTHESE = class(TOF)
    procedure OnNew; override;
    procedure OnDelete; override;
    procedure OnUpdate; override;
    procedure OnLoad; override;
    procedure OnArgument(S: string); override;
    procedure OnDisplay; override;
    procedure OnClose; override;
    procedure OnCancel; override;
  private
    fTE: TTypeEtatSynthese;
    fFichierMaquette: string;
    fbModeSimple: boolean;
    fbAnalytique : boolean;
    function RecupereCritere(var Crit: TCritEdtPCL): integer;
    function RecupereCritereColonne ( Indice : integer;  var Col : TColPCL ) : boolean;
    function GetFichierMaquette: string;
    function stPlusMaquette: string;
    procedure OnModeClick(Sender: TObject);
    procedure OnComparatif3Click(Sender: TObject);
    procedure OnComparatif4Click(Sender: TObject);
    procedure OnElipsisComparatif3Click(Sender: TObject);
    procedure OnElipsisComparatif4Click(Sender: TObject);
    procedure OnChangeExercice ( Sender : TObject );
    procedure MajListePeriode(CB: THValCombobox);
    procedure ConstruitTobCompare(Tc: TOB);
    procedure SwapMode(bVersModeSimple: boolean);
    function RecalcFVariation(Indice: integer): string;
    procedure AjouteColonnePourcent(Indice: integer; var Col: TColPCL; ColPrec : TColPCL);
  end;

function ConvStDate(Exo: TExoDate; Debut, Fin: TDateTime; var stLibelle:
  string): string;
function QuelTitre(Date1, Date2: tDateTime): string;

const
  ERR_ESYMAQUETTE = 'Maquette introuvable';

procedure CPLanceEtatSynthese(TE: TTypeEtatSynthese);

implementation

const
  ESY_MAXMSG = 1;

var
  HMSG: array[1..ESY_MAXMSG] of string =
  {1}('Maquette introuvable. Vous devez la modifier.');

function NBP(St1, St2:string):Word;
var
    Ex:tExoDate;
    XX, YY, ZZ:Word;
    Date1, Date2:tDateTime;
begin
    Date1 := StrToDate(St1);
    Date2 := StrToDate(St2);
    FillChar(Ex, SizeOf(Ex), #0);
    EX.Deb := Date1;
    EX.Fin := Date2;
    NOMBREPEREXO(Ex, XX, YY, ZZ);
    Result := ZZ;
end;

procedure DecaleIndice(var St: string);
var
  St1, St2: string;
  i: Integer;
  C: Char;
begin
  if St = '' then
    Exit;
  St1 := St;
  i := Pos('COL', ST1);
  while i > 0 do
  begin
    St1 := FindEtReplace(ST1, 'COL', 'WWW', FALSE);
    St2 := Copy(St1, i + 3, 1);
    C := St2[1];
    case C of
      '2': C := '3';
      '3': C := '5';
      '4': C := '7';
    end;
    St[i + 3] := C;
    i := Pos('COL', ST1);
  end;
end;

function PourCentSurFormule(St: string; l: Integer): string;
var
  St1, St2, St3: string;
  k: Integer;
begin
  k := Pos('-', St);
  if k <= 0 then
  begin
    Result := '=100*[COL' + IntToStr(l) + ']/[BASE' + IntToStr(l) + ']';
    Exit;
  end;
  St1 := Copy(St, k + 1, Length(St) - k);
  St2 := Copy(St, 2, Length(St) - 1);
  St3 := '=100*(' + St2 + ')/(' + St1 + ')';
  DecaleIndice(St3);
  Result := St3;
end;

procedure CPLanceEtatSynthese(TE: TTypeEtatSynthese);
var
  stAction: string;
begin
  case TE of
    esSIG: stAction := 'SIG';
    esBIL: stAction := 'BIL';
    esCR: stAction := 'CR';
    esCRA: stAction := 'CRA';
    esSIGA: stAction := 'SIGA';
    esBILA: stAction := 'BILA';
  end;
  AGLLanceFiche('CP', 'CPETATSYNTHESE', '', '', stAction);
end;

procedure TOF_CPETATSYNTHESE.OnNew;
begin
  inherited;
end;

procedure TOF_CPETATSYNTHESE.OnDelete;
begin
  inherited;
end;

procedure TOF_CPETATSYNTHESE.OnUpdate;
var
  Crit: TCritEdtPCL;
  CritEdtChaine: TCritEdtChaine;
  TFormule, TFormat, TFormatDetail: array of string;
  i: integer;
begin
  inherited;
  { Récupération des critères renseignés par l'utilsateur }
  if (RecupereCritere(Crit) = 0) then
  begin
    { Etats chainés }
    Fillchar(CritEdtChaine, SizeOf(CritEdtChaine), #0);
    CritEdtChaine.Utiliser := False;

    SetLength(TFormule, Crit.NbColActif);
    SetLength(TFormat, Crit.NbColActif);
    SetLength(TFormatDetail, Crit.NbColActif);
    for i := 0 to Crit.NbColActif - 1 do
    begin
      TFormule[i] := Crit.Col[i].StFormule;
      TFormat[i] := '#,##0.00;; ;';
      TFormatDetail[i] := '#,##0.00;; ;';
    end;

    LanceLiasse(fFichierMaquette
      , Crit.AvecDetail
      , False
      , TFormule
      , TFormat
      , TFormatDetail
      , Crit
      , FALSE
      , CritEdtChaine
      );
  end
  else
    PGIBox(HMSG[1], ECRAN.Caption);
end;

procedure TOF_CPETATSYNTHESE.OnLoad;
begin
  inherited;
end;

procedure TOF_CPETATSYNTHESE.OnArgument(S: string);
begin
  inherited;

  if S = 'SIG' then
    fTE := esSIG
  else if S = 'BIL' then
    fTE := esBIL
  else if S = 'CR' then
    fTE := esCR
  else if S = 'CRA' then
    fTE := esCRA
  else if S = 'SIGA' then
    fTE := esSIGA
  else if S = 'BILA' then
    fTE := esBILA;

  fbModeSimple := True;
  fbAnalytique :=  ((fTE = esCRA) or (fTE=esSIGA));

  SetControlVisible('PANALYTIQUE',fbAnalytique);
  SetControlVisible('PMULTIDOSSIER',ExisteSQL('SELECT YMD_CODE FROM YMULTIDOSSIER'));
  SetControlVisible('FMAQUETTEFICHIER', not (ctxPCL in V_PGI.PGIContexte));
  SetControlVisible('FMAQUETTESTD', (ctxPCL in V_PGI.PGIContexte));
  SetControlProperty('FMAQUETTESTD', 'Plus', stPlusMaquette);
  SetControlText ('FMAQUETTESTD',IntToStr(GetParamSocSecur('SO_NUMPLANREF',0)));

  TToolBarButton97(GetControl('BMODE')).OnClick := OnModeClick;
  TCheckBox(GetControl('AVECCOMPARATIF3_A')).OnClick := OnComparatif3Click;
  TCheckBox(GetControl('AVECCOMPARATIF4_A')).OnClick := OnComparatif4Click;
  THCritMaskEdit(GetControl('COMPARE3')).OnElipsisClick := OnElipsisComparatif3Click;
  THCritMaskEdit(GetControl('COMPARE4')).OnElipsisClick := OnElipsisComparatif4Click;
  THValComboBox(GetControl('PERIODE1_A')).OnChange := OnChangeExercice;
  THValComboBox(GetControl('PERIODE2_A')).OnChange := OnChangeExercice;
  THValComboBox(GetControl('PERIODE3_A')).OnChange := OnChangeExercice;
  THValComboBox(GetControl('PERIODE4_A')).OnChange := OnChangeExercice;
  THValComboBox(GetControl('PERIODE1_S')).OnChange := OnChangeExercice;
  THValComboBox(GetControl('PERIODE2_S')).OnChange := OnChangeExercice;
  THValComboBox(GetControl('PERIODE1_S')).Vide := True;
  THValComboBox(GetControl('PERIODE1_S')).VideString := '<<'+TraduireMemoire('Aucun')+'>>';
  THValComboBox(GetControl('PERIODE2_S')).Vide := True;
  THValComboBox(GetControl('PERIODE2_S')).VideString := '<<'+TraduireMemoire('Aucun')+'>>';

  { Chargement des périodes }
  MajListePeriode(THValComboBox(GetControl('PERIODE1_S')));
  MajListePeriode(THValComboBox(GetControl('PERIODE2_S')));
  MajListePeriode(THValComboBox(GetControl('PERIODE1_A')));
  MajListePeriode(THValComboBox(GetControl('PERIODE2_A')));
  MajListePeriode(THValComboBox(GetControl('PERIODE3_A')));
  MajListePeriode(THValComboBox(GetControl('PERIODE4_A')));

  { Affichage du panel en mode simplifié }
  fbModeSimple := True;
  SetActiveTabSheet('PGENERALITE');
  SetActiveTabSheet('PSIMPLE');

  { Initialisation des zones }

  SetControlText('PERIODE1_A', CExerciceVersRelatif ( ctxExercice.EnCours.Code ) );
  if ctxExercice.Precedent.Code <> '' then
    SetControlText('PERIODE2_A', CExerciceVersRelatif ( ctxExercice.Precedent.Code ) )
  else
    SetControlText('PERIODE2_A', CExerciceVersRelatif ( ctxExercice.EnCours.Code ) );
  SetControlChecked('ETAT', True);
end;

procedure TOF_CPETATSYNTHESE.OnClose;
begin
  inherited;
end;

procedure TOF_CPETATSYNTHESE.OnDisplay();
begin
  inherited;
end;

procedure TOF_CPETATSYNTHESE.OnCancel();
begin
  inherited;
end;

function TOF_CPETATSYNTHESE.RecupereCritere(var Crit: TCritEdtPCL): integer;
var
  iCol: integer;
begin
  Result := 0;
  Fillchar(Crit, SizeOf(Crit), #0);

  { Maquette d'édition }
  fFichierMaquette := GetFichierMaquette;
  if not FileExists(fFichierMaquette) then
  begin
    Result := 1;
  end
  else
  begin
    Crit.TE := fTE;
    Crit.Resolution := '';
    Crit.ResolutionDetail := '';
    Crit.bEtat := (GetCheckBoxState('ETAT') = cbChecked);
    Crit.AvecDetail := (GetCheckBoxState('DETAIL') = cbChecked);
    Crit.AvecPourcent := (GetCheckBoxState('POURCENTAGE') = cbChecked);
    Crit.EnMonnaieOpposee := '';
    Crit.StTypEcr := 'N';
    iCol := 0;
    { Récupération des données de la colonne n°1 }
    if RecupereCritereColonne ( 1, Crit.Col[iCol] ) then
    begin
      Inc(iCol);
      if Crit.AvecPourcent then
      begin
        AjouteColonnePourcent ( 1, Crit.Col[iCol] , Crit.Col[iCol-1] );
        Inc ( iCol );
      end;
      { Récupération des données de la colonne n°2 }
      if RecupereCritereColonne ( 2, Crit.Col[iCol]) then
      begin
        Inc(iCol);
        if Crit.AvecPourcent then
        begin
          AjouteColonnePourcent ( 2, Crit.Col[iCol], Crit.Col[iCol-1] );
          Inc (iCol);
        end;
        { Récupération des données de la colonne n°3 }
        if RecupereCritereColonne ( 3, Crit.Col[iCol] ) then
        begin
          Inc (iCol);
          if Crit.AvecPourcent then
          begin
            AjouteColonnePourcent ( 3, Crit.Col[iCol], Crit.Col[iCol-1] );
            Inc (iCol);
          end;
          if not fbModeSimple then
          begin
            if RecupereCritereColonne ( 4, Crit.Col[iCol] ) then
            begin
              Inc (iCol);
              if Crit.AvecPourcent then
              begin
                AjouteColonnePourcent ( 4, Crit.Col[iCol], Crit.Col[iCol-1] );
                Inc (iCol);
              end;
            end;
          end;
        end;
      end;
    end;
    Crit.NbColActif := iCol;
    { Récupération des critères analytiques }
    Crit.stAxe := GetControlText('AXE');
    Crit.stSectionDe := GetControlText('SECTIONDE');
    Crit.stSectionA := GetControlText('SECTIONA');
    Crit.bUnEtatParSection := (GetCheckBoxState('BUNETATPARSECTION')=cbChecked);
  end;
end;

function TOF_CPETATSYNTHESE.GetFichierMaquette: string;
var
  Q: TQuery;
begin
  if ctxPCL in V_PGI.PGIContexte then
  begin
    Q := OpenSQL('SELECT * FROM STDMAQ WHERE STM_TYPEMAQ="' + stPlusMaquette +
      '" AND STM_NUMPLAN=' + GetControlText('FMAQUETTESTD'), True);
    if not Q.Eof then
    begin
      if Q.FindField('STM_PREDEFINI').AsString = 'DOS' then
        Result := ChangeStdDatPath('$DOS\' + stPlusMaquette + Format('%.03d',
          [Q.FindField('STM_NUMPLAN').AsInteger]) + '.txt', True)
      else if Q.FindField('STM_PREDEFINI').AsString = 'CEG' then
        Result := ChangeStdDatPath('$STD\' + stPlusMaquette + Format('%.03d',
          [Q.FindField('STM_NUMPLAN').AsInteger]) + '.txt', True)
      else
        Result := ChangeStdDatPath('$DAT\' + stPlusMaquette + Format('%.03d',
          [Q.FindField('STM_NUMPLAN').AsInteger]) + '.txt', True);
    end;
    Ferme(Q);
  end
  else
  begin
    Result := GetControlText('FMAQUETTEFICHIER');
  end;
end;

function TOF_CPETATSYNTHESE.stPlusMaquette: string;
begin
  case fTE of
    esCR, esCRA: Result := 'CR';
    esBIL, esBILA: Result := 'BIL';
    esSIG, esSIGA: Result := 'SIG';
  else
    Result := '';
  end;
end;

procedure TOF_CPETATSYNTHESE.SwapMode ( bVersModeSimple : boolean );
begin
  if bVersModeSimple then
  begin
    SetControlText ('PERIODE1_S', GetControlText('PERIODE1_A'));
    SetControlText ('PERIODE2_S', GetControlText('PERIODE2_A'));
    SetControlText ('SITUATION1_S', GetControlText('SITUATION1_A'));
    SetControlText ('SITUATION2_S', GetControlText('SITUATION2_A'));
    SetControlChecked ('AVECCOMPARATIF3_S', ( GetCheckBoxState ( 'AVECCOMPARATIF3_A') = cbChecked ) );
  end else
  begin
    SetControlText ('PERIODE1_A', GetControlText('PERIODE1_S'));
    SetControlText ('PERIODE2_A', GetControlText('PERIODE2_S'));
    SetControlText ('SITUATION1_A', GetControlText('SITUATION1_S'));
    SetControlText ('SITUATION2_A', GetControlText('SITUATION2_S'));
    SetControlChecked ('AVECCOMPARATIF3_A', ( GetCheckBoxState ( 'AVECCOMPARATIF3_S') = cbChecked ) );
    if ( GetCheckBoxState ( 'AVECCOMPARATIF3_S') = cbChecked ) then
      SetControlText ('COMPARE3','[COL1]-[COL2]');
  end;
end;

procedure TOF_CPETATSYNTHESE.OnModeClick(Sender: TObject);
begin
  if fbModeSimple then
  begin
    SetControlProperty('BMODE', 'Caption', 'Mode simplifié');
    TPageControl(GetControl('PCHOIXMODE')).ActivePage := TTabSheet(GetControl('PAVANCE'));
    SwapMode ( False );
  end
  else
  begin
    SetControlProperty('BMODE', 'Caption', 'Mode avancé');
    TPageControl(GetControl('PCHOIXMODE')).ActivePage := TTabSheet(GetControl('PSIMPLE'));
    SwapMode ( True);
  end;
  fbModeSimple := not fbModeSimple;
end;

procedure TOF_CPETATSYNTHESE.MajListePeriode(CB: THValCombobox);
begin
  { Chargement de la liste des exercices }
  CInitComboExercice(CB);
end;

procedure TOF_CPETATSYNTHESE.OnComparatif3Click(Sender: TObject);
var
  bChecked: boolean;
begin
  bChecked := (GetCheckBoxState('AVECCOMPARATIF3_A') = cbChecked);
  SetControlVisible('PERIODE3_A', not bChecked);
  SetControlVisible('LDATE3_A', not bChecked);
  SetControlVisible('DATE3_A', not bChecked);
  SetControlVisible('LDATE3_A_', not bChecked);
  SetControlVisible('DATE3_A_', not bChecked);
  SetControlVisible('SITUATION3_A', not bChecked);
  SetControlVisible('COMPARE3', bChecked);
end;

procedure TOF_CPETATSYNTHESE.OnComparatif4Click(Sender: TObject);
var
  bChecked: boolean;
begin
  bChecked := (GetCheckBoxState('AVECCOMPARATIF4_A') = cbChecked);
  SetControlVisible('PERIODE4_A', not bChecked);
  SetControlVisible('LDATE4_A', not bChecked);
  SetControlVisible('DATE4_A', not bChecked);
  SetControlVisible('LDATE4_A_', not bChecked);
  SetControlVisible('DATE4_A_', not bChecked);
  SetControlVisible('SITUATION4_A', not bChecked);
  SetControlVisible('COMPARE4', bChecked);
end;

procedure TOF_CPETATSYNTHESE.OnElipsisComparatif3Click(Sender: TObject);
var
  lTob, Tc: TOB;
begin
  Tc := TOB.Create('', nil, -1);
  try
    ConstruitTobCompare(Tc);
    lTob := LookUpTob(THCritMaskEdit(GetControl('COMPARE3')), Tc,
      'Choix du comparatif', 'LIBELLE', 'Libellé');
    if assigned(lTOB) then
      SetControlText('COMPARE3', lTOB.GetValue('FORMULE'));
  finally
    Tc.Free;
  end;
end;

procedure TOF_CPETATSYNTHESE.OnElipsisComparatif4Click(Sender: TObject);
var
  lTob, Tc: TOB;
begin
  Tc := TOB.Create('', nil, -1);
  try
    ConstruitTobCompare(Tc);
    lTob := LookUpTob(THCritMaskEdit(GetControl('COMPARE4')), Tc,
      'Choix du comparatif', 'LIBELLE;FORMULE', 'Libellé;Formule');
    if assigned(lTOB) then
      SetControlText('COMPARE4', lTOB.GetValue('FORMULE'));
  finally
    Tc.Free;
  end;
end;

procedure TOF_CPETATSYNTHESE.ConstruitTobCompare(Tc: TOB);
var
  T: TOB;
begin
  if Tc = nil then
    exit;
  T := TOB.Create('', Tc, -1);
  T.AddChampSupValeur('CODE', 'SIM');
  T.AddChampSupValeur('LIBELLE', 'Simple');
  T.AddChampSupValeur('FORMULE', RecalcFVariation( 0 ));

  T := TOB.Create('', Tc, -1);
  T.AddChampSupValeur('CODE', '12M');
  T.AddChampSupValeur('LIBELLE', 'Sur 12 mois');
  T.AddChampSupValeur('FORMULE', RecalcFVariation( 1 ));

  T := TOB.Create('', Tc, -1);
  T.AddChampSupValeur('CODE', 'PLP');
  T.AddChampSupValeur('LIBELLE', 'Ramené au plus petit');
  T.AddChampSupValeur('FORMULE', RecalcFVariation( 2 ));

  T := TOB.Create('', Tc, -1);
  T.AddChampSupValeur('CODE', 'PLG');
  T.AddChampSupValeur('LIBELLE', 'Ramené au plus grand');
  T.AddChampSupValeur('FORMULE', RecalcFVariation( 3 ));
end;

function ConvStDate(Exo: TExoDate; Debut, Fin: TDateTime; var stLibelle:
  string): string;
  function _ConvStUneDate(d: tDateTime): string;
  var
    YY, MM, DD: Word;
  begin
    DecodeDate(d, YY, MM, DD);
    Result := '(' + FormatFloat('00', DD) + '-' + FormatFloat('00', MM) + '-' +
      FormatFloat('0000', YY) + ')';
  end;

var
  i, j, k: Integer;
  st: string;
begin
  St := '';
  if (Exo.Deb = Debut) and (Exo.Fin = Fin) then
  begin
    if Exo.Code = ctxExercice.EnCours.Code then
    begin
      Result := 'N';
      stLibelle := 'N';
    end
    else if Exo.Code = ctxExercice.Precedent.Code then
    begin
      Result := 'N-';
      stLibelle := 'N-1';
    end
    else if Exo.Code = ctxExercice.Suivant.Code then
    begin
      Result := 'N+';
      stLibelle := 'N+1';
    end
    else
      for i := 1 to 5 do
        if Exo.Code = ctxExercice.ExoClo[i].Code then
        begin
          k := 1;
          for j := 5 downto i do
            if ctxExercice.ExoClo[j].code <> '' then
            begin
              St := St + '-';
              Inc(k);
            end;
          Result := 'N' + St;
          stLibelle := 'N-' + IntToStr(k);
        end;
    St := _ConvStUneDate(Debut);
    St := St + _ConvStUneDate(Fin);
    stLibelle := QuelTitre(Debut, Fin);
  end
  else
  begin
    St := _ConvStUneDate(Debut);
    St := St + _ConvStUneDate(Fin);
    Result := St;
    stLibelle := QuelTitre(Debut, Fin);
  end;
end;

function QuelTitre(Date1, Date2: tDateTime): string;
var
  YY2, MM2, DD2, YY1, MM1, DD1: Word;
  Deb, Fin: Boolean;
  St: string;
  stDuree: string;
  PremMois, PremAnnee, NbMois: Word;
begin
  DecodeDate(Date2, YY2, MM2, DD2);
  DecodeDate(Date1, YY1, MM1, DD1);
  St := '';
  Deb := (DebutDeMois(Date1) = Date1);
  Fin := (FinDeMois(Date2) = Date2);
  if ctxPCL in V_PGI.PGIContexte then
  begin
    NOMBREMOIS(Date1, Date2, PremMois, PremAnnee, NbMois);
    if NbMois > 0 then
      stDuree := IntToStr(NbMois) + TraduireMemoire(' mois')
    else
      stDuree := FormatFloat('#', Date2 - Date1) + TraduireMemoire(' jours');
  end;
  if (MM1 = MM2) and (YY1 = YY2) then
  begin
    if Deb and Fin then
    begin
      St := FormatDateTime('MMMM YYYY', Date1);
    end
    else
    begin
      St := FormatDateTime('MMM YYYY', Date1) + #13 + #10 + '(' +
        FormatFloat('00', DD1) + '-' +
        FormatFloat('00', DD2) + ')';
    end;
  end
  else
  begin
    if Deb and Fin then
    begin
      st := TraduireMemoire('du ') + FormatDateTime('dd/mm/yy',
        Date1) + #13 + #10 + TraduireMemoire('au ') +
        FormatDateTime('dd/mm/yy', Date2) + #13 + #10 + stDuree;
    end
    else
    begin
      St := TraduireMemoire('du ') + FormatDateTime('dd/mm/yy',
        Date1) + #13 + #10 + TraduireMemoire('au ') +
        FormatDateTime('dd/mm/yy', Date2) + #13 + #10 + stDuree;
    end;
  end;
  Result := St;
end;

procedure TOF_CPETATSYNTHESE.OnChangeExercice(Sender: TObject);
var
  St : string;
  Exo : TExoDate;
begin
  St := THValComboBox(Sender).Name;
  Exo := ctxExercice.QuelExoDate(CRelatifVersExercice(GetControlText(St)));
  SetControlText ('DATE'+St[8]+'_A',DateToStr(Exo.Deb));
  SetControlText ('DATE'+St[8]+'_A_',DateToStr(Exo.Fin));  
end;

function TOF_CPETATSYNTHESE.RecalcFVariation( Indice : integer ) : string;
var
    NbP1, nbP2:Word;
    NbpMin, NbpMax:Word;
    St:string;
begin
  nbP1 := NBP(GetControlText('DATE1_A'), GetControlText('DATE1_A_'));
  nbP2 := NBP(GetControlText('DATE2_A'), GetControlText('DATE2_A_'));
  NbPMin := Nbp1;
  NbPMax := Nbp2;
  if NbP1 > NbP2 then
  begin
      NbPMin := Nbp2;
      NbPMax := Nbp1;
  end;
  case Indice of
      0:
          begin
              St := '=[COL1]-[COL2]';
          end;
      1:
          begin
              St := '=(12*[COL1]/' + IntToStr(Nbp1) + ')-(12*[COL2]/' + IntToStr(Nbp2) + ')';
          end;
      2:
          begin
              St := '=(' + IntToStr(NbpMin) + '*[COL1]/' +
                  IntToStr(Nbp1) + ')-(' + IntToStr(NbpMin) + '*[COL2]/'
                  + IntToStr(Nbp2) + ')';
          end;
      3:
          begin
              St := '=(' + IntToStr(NbpMax) + '*[COL1]/' +
                  IntToStr(Nbp1) + ')-(' + IntToStr(NbpMax) + '*[COL2]/'
                  + IntToStr(Nbp2) + ')';
          end;
  end;
  Result := St;
end;

function TOF_CPETATSYNTHESE.RecupereCritereColonne(  Indice : integer; var Col: TColPCL): boolean;
var
  stSuffixe : string;
  Exo : TExoDate;
  stLibelle : string;
begin
  Result := False;
  if fbModeSimple then stSuffixe := IntToStr(Indice)+'_S'
  else stSuffixe := IntToStr(Indice)+'_A';
  if ((Indice >= 3) and (GetCheckBoxState('AVECCOMPARATIF'+stSuffixe)=cbChecked)) then
  begin
    if fbModeSimple then
    begin
      if (GetCheckBoxState('POURCENTAGE') = cbChecked) then Col.stFormule := '=[COL1]-[COL3]'
      else Col.stFormule := '=[COL1]-[COL2]';
    end else
    begin
      Col.StFormule := GetControlText('FORMULE'+IntToStr(Indice));
    end;
    Col.StTitre := 'Simple : ' + #13 + #10 + 'Variation en valeur';
    Col.IsFormule := True;
    Result := True;
  end else
  if ((fbModeSimple and (Indice <=2) ) or (not fbModeSimple)) then
  begin
    if (GetControlText('PERIODE'+stSuffixe)<> '' ) or (GetControlText('SITUATION'+stSuffixe)<> '' ) then
    begin
      Col.Actif := True;
      if GetControlText('SITUATION'+stSuffixe) <> '' then
          Col.BalSit := GetControlText('SITUATION'+stSuffixe)
      else
      begin
        Exo := ctxExercice.QuelExoDate(CRelatifVersExercice(GetControlText('PERIODE'+stSuffixe)));
        Col.stFormule := ConvStDate(Exo, Exo.Deb, Exo.Fin, stLibelle);
        Col.stTitre := stLibelle;
      end;
      Col.IsFormule := False;
      Result := True;
    end;
  end;
end;

procedure TOF_CPETATSYNTHESE.AjouteColonnePourcent ( Indice : integer; var Col: TColPCL; ColPrec : TColPCL);
begin
  if ColPrec.IsFormule then
    Col.stFormule := PourCentSurFormule(ColPrec.stFormule, (2*Indice)-1)
  else
    Col.stFormule := '=100*[COL' + IntToStr((2*Indice)-1) + ']/[BASE' + IntToStr((2*Indice)-1) + ']';
  Col.stTitre := '%';
  Col.IsFormule := True;
end;

initialization
  registerclasses([TOF_CPETATSYNTHESE]);
end.

