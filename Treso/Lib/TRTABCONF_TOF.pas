{-------------------------------------------------------------------------------------
   Version    |  Date  | Qui |   Commentaires
--------------------------------------------------------------------------------------
08.10.001.001  06/08/07   JP   Création de l'unité : Liste pour la gestion des tablettes des confidentialités
--------------------------------------------------------------------------------------}
unit TRTABCONF_TOF;

interface

uses
  Controls, Classes,
  {$IFNDEF EAGLCLIENT}
  FE_Main,
  {$ELSE}
  MaineAGL,
  {$ENDIF}
  SysUtils, HCtrls, UTOF, UTob;

type
  TOF_TRTABCONF = class (TOF)
    procedure OnArgument(S : string); override;
    procedure OnUpdate              ; override;
    procedure OnClose               ; override;
  private
    FTobTablette  : TOB;
    FChamp        : string;
    FCaption      : string;
    FEcranKeyDown : TKeyEvent;

    procedure GetEcran;
    procedure ChargeTOB;
    procedure SetEcran;
    procedure ShowLignes;
    procedure GereBasculement;
    function  RetourneSelection : string;
  public
    Grille       : THGrid;
    edCode       : THEdit;
    edLibelle    : THEdit;

    procedure OnEditKeyUp   (Sender : TObject; var Key : Word; Shift : TShiftState);
    procedure OnEcranKeyDown(Sender : TObject; var Key : Word; Shift : TShiftState);
  end;


function TrLanceFiche_TabletteConf(Range, Lequel, Arg : string) : string;


implementation

uses
  HEnt1, Vierge, ULibWindows, Windows, Constantes;

var
  gRange  : string;
  gLequel : string;

{---------------------------------------------------------------------------------------}
function TrLanceFiche_TabletteConf(Range, Lequel, Arg : string) : string;
{---------------------------------------------------------------------------------------}
begin
  gRange  := Range;
  gLequel := Lequel;
  Result := AglLanceFiche('TR', 'TRTABCONF', Range, Lequel, Arg);
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRTABCONF.OnArgument(S : string);
{---------------------------------------------------------------------------------------}
begin
  inherited;
  TFVierge(Ecran).Retour := '@@@@';
  
  {Récupération du champ à Traiter}
  FChamp := ReadTokenSt(S);
  {Initialisation des variables et méthodes d'interface}
  GetEcran;
  {construction de la Requête et chargement de la tob}
  ChargeTOB;
  {présentation de l'écran}
  SetEcran;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRTABCONF.OnClose;
{---------------------------------------------------------------------------------------}
begin
  if Assigned(FTobTablette) then FreeAndNil(FTobTablette);
  inherited;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRTABCONF.OnUpdate;
{---------------------------------------------------------------------------------------}
begin
  inherited;
  TFVierge(Ecran).Retour := RetourneSelection;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRTABCONF.OnEditKeyUp(Sender : TObject; var Key : Word; Shift : TShiftState);
{---------------------------------------------------------------------------------------}
begin
  ShowLignes;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRTABCONF.OnEcranKeyDown(Sender : TObject; var Key : Word; Shift : TShiftState);
{---------------------------------------------------------------------------------------}
begin
  case Key of
    VK_F12  : GereBasculement;
  else
    FEcranKeyDown(Sender, Key, Shift);
  end;
end;

{Affichage de l'écran
{---------------------------------------------------------------------------------------}
procedure TOF_TRTABCONF.SetEcran;
{--------------------------------------------------------------------------------------}
var
  Enreg : string;
  n     : Integer;
begin
  {Mise à jour du caption de l'écran}
  Ecran.Caption := FCaption;
  UpdateCaption(Ecran);

  {S'il n'y a pas de range ...}
  if gRange = '' then begin
    {... on affiche la Tob}
    FTobTablette.PutGridDetail(Grille, False, False, '', True);
    {Gestion d'une éventuelle pré-sélection}
    if gLequel <> '' then begin
      Enreg := ReadTokenSt(gLequel);
      while Enreg <> '' do begin
        for n := 1 to Grille.RowCount - 1 do begin
          if Grille.Cells[0, n] = Enreg then begin
            Grille.FlipSelection(n);
            Break;
          end;
        end;
        Enreg := ReadTokenSt(gLequel);
      end;
    end;
  end
  else begin
    {Gestion du range}
    edCode.Text := gRange;
    ShowLignes;
  end;

  Grille.SetFocus;
end;

{Constitution de la requête et chargement de la TOB
{---------------------------------------------------------------------------------------}
procedure TOF_TRTABCONF.ChargeTOB;
{---------------------------------------------------------------------------------------}
var
  Suf : string;
  SQL : string;
begin
  {On commence par récupérer le suffixe du Champ}
  Suf := FChamp;
  ReadTokenPipe(Suf, '_');
  if Suf = 'CODE' then Suf := FChamp;

  {Gestion du caption de la fiche}
       if Suf = 'BANQUE'     then FCaption := TraduireMemoire('Liste des banques')
  else if Suf = 'AGENCE'     then FCaption := TraduireMemoire('Liste des agences bancaires')
  else if Suf = 'NODOSSIER'  then FCaption := TraduireMemoire('Liste des dossiers')
  else if Suf = 'GENERAL'    then FCaption := TraduireMemoire('Liste des comptes généraux')

  else if Suf = 'CLASSEFLUX' then FCaption := TraduireMemoire('Liste des classes de flux')
  else if Suf = 'TYPEFLUX'   then FCaption := TraduireMemoire('Liste des type de flux')
  else if Suf = 'CODECIB'    then FCaption := TraduireMemoire('Liste des CIB')

  else if Suf = 'BQ_CODE'    then FCaption := TraduireMemoire('Liste des comptes bancaires');

  {Constitution de la requête en fonction du suffixe}
       if Suf = 'BANQUE'     then SQL := 'SELECT PQ_BANQUE AS CODE, PQ_LIBELLE AS LIBELLE FROM BANQUES'
  else if Suf = 'AGENCE'     then SQL := 'SELECT TRA_AGENCE AS CODE, TRA_LIBELLE AS LIBELLE FROM AGENCE'
  else if Suf = 'NODOSSIER'  then SQL := 'SELECT DOS_NODOSSIER AS CODE, DOS_LIBELLE AS LIBELLE FROM DOSSIER'
  else if Suf = 'GENERAL'    then SQL := 'SELECT G_GENERAL AS CODE, G_LIBELLE AS LIBELLE FROM GENERAUX'

  else if Suf = 'CLASSEFLUX' then SQL := 'SELECT CO_CODE AS CODE,CO_LIBELLE AS LIBELLE FROM COMMUN WHERE CO_TYPE = "TRC"'
  else if Suf = 'TYPEFLUX'   then SQL := 'SELECT TTL_TYPEFLUX AS CODE, TTL_LIBELLE AS LIBELLE FROM TYPEFLUX'
  else if Suf = 'CODECIB'    then SQL := 'SELECT TCI_CODECIB AS CODE TCI_LIBELLE AS LIBELLE FROM CIB WHERE TCI_BANQUE = "' + CODECIBREF + '"'

  else if Suf = 'BQ_CODE'    then SQL := 'SELECT BQ_CODE AS CODE, BQ_LIBELLE AS LIBELLE FROM BANQUECP';

  {Chargement de la requête}
  FTobTablette := TOB.Create('TIOIT', nil, -1);
  FTobTablette.LoadDetailFromSQL(SQL);
end;

{Assignation des pointeurs sur composants et évènements
{---------------------------------------------------------------------------------------}
procedure TOF_TRTABCONF.GetEcran;
{---------------------------------------------------------------------------------------}
begin
  Grille    := (GetControl('GRILLE'   ) as THGrid);
  edCode    := (GetControl('EDCODE'   ) as THEdit);
  edLibelle := (GetControl('EDLIBELLE') as THEdit);
  edCode.OnKeyUp := OnEditKeyUp;
  edLibelle.OnKeyUp := OnEditKeyUp;
  FEcranKeyDown   := Ecran.OnKeyDown;
  Ecran.OnKeyDown := OnEcranKeyDown;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRTABCONF.ShowLignes;
{---------------------------------------------------------------------------------------}

      {--------------------------------------------------------------}
      function _ATraiter(T : TOB) : Boolean;
      {--------------------------------------------------------------}

          {Ajout d'un éventuel caractère joker pour signifier le "like"
          {-----------------------------------------------------}
          function _CompleteChaine(Ch : string) : string;
          {-----------------------------------------------------}
          var
            C : Char;
          begin
            Result := '';
            if Length(Ch) > 0 then C := Ch[Length(Ch)]
                              else Exit;

            if not (C in ['%', '_', '?', '*']) then Result := Ch + '*'
                                               else Result := Ch;
          end;

      begin
        Result := edCode.Text = '';

        if not Result then Result := CompareAvecJoker(_CompleteChaine(edCode.Text), T.GetString('CODE'));

        if Result then begin
          Result := edLibelle.Text = '';
          if not Result then Result := CompareAvecJoker(_CompleteChaine(edLibelle.Text), T.GetString('LIBELLE'));
        end;
      end;

var
  n : Integer;
  p : Integer;
  T : TOB;
begin
  Grille.BeginUpdate;
  p := 1;
  {1/ On vide la grille}
  Grille.RowCount := 1;
  {2/ On Insère une ligne vide qui servira de base pour le InsertRow}
  Grille.RowCount := 2;
  {3/ Remplissage de la grille}
  for n := 0 to FTobTablette.Detail.Count - 1 do begin
    T := FTobTablette.Detail[n];
    if _ATraiter(T) then begin
      Inc(p);
      Grille.InsertRow(p);
      Grille.Cells[0, p - 1] := T.GetString('CODE');
      Grille.Cells[1, p - 1] := T.GetString('LIBELLE');
    end;
  end;
  {3/ On supprime une éventuelle ligne vide en fin de grille}
  Grille.RowCount := p;
  {4/ Le nombre de lignes fixes doit être strictement inférieures au nombre de lignes}
  if p > 1 then Grille.FixedRows := 1;

  Grille.EndUpdate;
end;

{Gestion du retour de la fiche : Attention, il ne doit y avoir de point virgule que s'il y
 a au moins deux enregistrements sélectionnés
{---------------------------------------------------------------------------------------}
function TOF_TRTABCONF.RetourneSelection : string;
{---------------------------------------------------------------------------------------}
var
  n : Integer;
begin
  Result := '';
  for n := 1 to Grille.RowCount - 1 do begin
    if Grille.IsSelected(n) then
      if Result <> '' then Result := Result + ';' + Grille.Cells[0, n]
                      else Result := Grille.Cells[0, n];
  end;
  if Pos(';', Result) > 0 then Result := Result + ';';
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRTABCONF.GereBasculement;
{---------------------------------------------------------------------------------------}
begin
  if Ecran.ActiveControl is THGrid then EdCode.SetFocus
                                   else Grille.SetFocus;

end;

initialization
  RegisterClasses([TOF_TRTABCONF]);

end.
