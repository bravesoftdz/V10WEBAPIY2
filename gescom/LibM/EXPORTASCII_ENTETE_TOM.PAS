{***********UNITE*************************************************
Auteur  ...... : JULIEN D
Créé le ...... : 15/02/2002
Modifié le ... : 04/06/2002
Description .. : Source TOM de la TABLE : EXPORTASCII_ENTETE
Suite ........ : (EXPORTASCII_ENTETE)
Mots clefs ... : TOM;EXPORTASCII_ENTETE
*****************************************************************}
unit EXPORTASCII_ENTETE_TOM;

interface

uses StdCtrls, Controls, Classes,
  {$IFDEF EAGLCLIENT}
  eFiche,
  {$ELSE}
  db, dbtables, Fiche,
  {$ENDIF}
  forms, sysutils, ComCtrls, HCtrls, HEnt1, HMsgBox, UTOM, UTob,
  M3FP, SaisieValDefautASCII, ParamSoc;

function ASCIIConstruireListeDeChamps(TOBASCII: TOB): string;
procedure ASCIIAjouteChampGaucheEtDate(TOBASCII: TOB);
function ASCIIConstruireUneLigne(Q: TQuery; TOBASCII: TOB; CodeExport, Separateur,
  DateFormat, BufferAlpha, BufferNumerique: string; BFormatDynamique: Boolean): string;
function ASCIIFormatChaine(Chaine, ValDefaut, DateFormat, BufferAlpha, BufferNumerique: string;
  Len: integer; BGauche: boolean = False; BDate: boolean = False; LongueurDynamique: boolean = False): string;
function RecupLibelle_ChampDollar(NomChamp: string; var TypeChamp: string): string;
function CreationJointure(Prefixe: string; var ListeDesChamps: string): string;

type
  TOM_EXPORTASCII_ENTETE = class(TOM)
    procedure OnNewRecord; override;
    procedure OnDeleteRecord; override;
    procedure OnUpdateRecord; override;
    procedure OnLoadRecord; override;
    procedure OnArgument(S: string); override;
    procedure OnClose; override;
  private
    LeRange, CodeExport, Prefixes, Champs_dollars: string;
    TOB_GCHAMPDISPO, TOB_GCHAMPSELECT: TOB;
    GCHAMPDISPO, GCHAMPSELECT: THGrid;
    procedure SetLastError(Num: integer; ou: string);
    procedure AffectePrefixes;
    procedure InitFiche;
    procedure DestroyFiche;
    procedure ChargeGrid;
    procedure AfficheGrid;
    procedure BClick(NameBouton: string);
    procedure AjouteUnChamp;
    procedure SupprimeUnChamp;
    procedure DeplaceChampVersLeHaut;
    procedure DeplaceChampVersLeBas;
    function ControleDonnees: boolean;
    procedure SupprimeAncienChamps;
    procedure NumeroteLignesTOB(GS: THGrid; TOBGrid: TOB);
    function LongueurParDefautChamp(TypeChamp: string): integer;
  end;

const
  // libellés des messages
  TexteMessage: array[1..8] of string = (
    {1}'Nature de paramétrage inexistante.'
    {2}, 'Ce code d''extraction existe déjà.'
    {3}, 'Vous devez renseigner un code.'
    {4}, 'Vous devez sélectionner au minimum un champ.'
    {5}, 'Vous devez renseigner un libellé.'
    {6}, 'Vous devez renseigner un séparateur'
    {7}, 'Vous devez renseigner un caractère de bourrage numérique'
    {8}, 'Vous devez renseigner un caractère de bourrage alpha'
    );
  Col_NOMCHAMP = 1;
  Col_DEBUTCHAMP = 2;
  Col_LENCHAMP = 3;

  TabChampSuppArticle: array[1..5, 1..3] of string = (
    ('$$_ARTLIBELLEDIM1', 'VARCHAR(35)', 'Libellé dimension 1'),
    ('$$_ARTLIBELLEDIM2', 'VARCHAR(35)', 'Libellé dimension 2'),
    ('$$_ARTLIBELLEDIM3', 'VARCHAR(35)', 'Libellé dimension 3'),
    ('$$_ARTLIBELLEDIM4', 'VARCHAR(35)', 'Libellé dimension 4'),
    ('$$_ARTLIBELLEDIM5', 'VARCHAR(35)', 'Libellé dimension 5')
    );

  TabChampSuppQTEFACT: array[1..3, 1..3] of string = (
    ('$$_QTEFACTABS', 'DOUBLE', 'Valeur absolue de GL_QTEFACT'),
    ('$$_QTEFACTSIGNE', 'VARCHAR(1)', 'Signe de GL_QTEFACT'),
    ('$$_SUPPQTEFACTZERO', 'VARCHAR(1)', 'Supprime enregistrement où GL_QTEFACT=0')
    );

implementation

function ASCIIConstruireListeDeChamps(TOBASCII: TOB): string;
var St: string;
  i: integer;
begin
  St := '';
  for i := 0 to TOBASCII.Detail.Count - 1 do
  begin
    if St <> '' then St := St + ',';
    St := St + TOBASCII.Detail[i].GetValue('GEX_NOMCHAMP');
  end;
  Result := St;
end;

procedure ASCIIAjouteChampGaucheEtDate(TOBASCII: TOB);
var i: integer;
  StChamp, StTypeChamp: string;
  Gauche: boolean;
begin
  for i := 0 to TOBASCII.Detail.Count - 1 do
  begin
    with TOBASCII.Detail[i] do
    begin
      StChamp := GetValue('GEX_NOMCHAMP');
      RecupLibelle_ChampDollar(StChamp, StTypeChamp);
      Gauche := (StTypeChamp = 'INTEGER') or (StTypeChamp = 'DOUBLE');
      AddChampSup('GAUCHE', False);
      PutValue('GAUCHE', Gauche);
      AddChampSup('DATE', False);
      PutValue('DATE', StTypeChamp = 'DATE');
    end;
  end;
end;

function ASCIIConstruireUneLigne(Q: TQuery; TOBASCII: TOB; CodeExport, Separateur,
  DateFormat, BufferAlpha, BufferNumerique: string; BFormatDynamique: Boolean): string;
var St, StChaine, StValDefaut, NomChamp: string;
  BGauche, BDate: Boolean;
  i, Longueur, Debut: integer;
begin
  St := '';
  for i := 0 to TOBASCII.Detail.Count - 1 do
  begin
    if St <> '' then St := St + Separateur;
    with TOBASCII.Detail[i] do
    begin
      StChaine := '';
      Debut := GetValue('GEX_DEBUT');
      NomChamp := GetValue('GEX_NOMCHAMP');

      if NomChamp = '$$_QTEFACTABS' then
        StChaine := FloatToStr(ABS(Q.FindField('GL_QTEFACT').AsFloat))
      else
        if NomChamp = '$$_QTEFACTSIGNE' then
      begin
        if Q.FindField('GL_QTEFACT').AsFloat >= 0 then StChaine := '+' else StChaine := '-';
      end
      else
        if Pos('$$_', NomChamp) > 0 then System.Delete(NomChamp, 1, 3);

      if StChaine = '' then StChaine := Q.FindField(NomChamp).AsString;
      //Réduit la chaîne en fonction de la valeur du début de chaîne
      StChaine := Copy(StChaine, Debut, Length(StChaine) + 1 - Debut);
      StValDefaut := GetValue('GEX_VALDEFAUT');
      Longueur := GetValue('GEX_LENCHAMP');
      BGauche := GetValue('GAUCHE');
      BDate := GetValue('DATE');
    end;
    St := St + ASCIIFormatChaine(StChaine, StValDefaut, DateFormat, BufferAlpha, BufferNumerique, Longueur, BGauche, BDate, BFormatDynamique);
  end;
  Result := St;
end;

function ASCIIFormatChaine(Chaine, ValDefaut, DateFormat, BufferAlpha, BufferNumerique: string;
  Len: integer; BGauche: boolean = False; BDate: boolean = False; LongueurDynamique: boolean = False): string;
var St: string;
begin
  St := Chaine;
  if LongueurDynamique then
  begin
    BufferAlpha := ' ';
    BufferNumerique := ' ';
  end;
  if St = '' then St := ValDefaut;
  try
    if BDate and (DateFormat <> '') then St := FormatDateTime(DateFormat, StrToDateTime(St));
  except
    PGIError('Erreur de format de date dans le modèle d''export', 'Export ASCII');
  end;
  while (Length(St) < Len) do
    if BGauche then St := BufferNumerique + St else St := St + BufferAlpha;
  if (Length(St) > Len) then St := Copy(St, 1, Len);
  if LongueurDynamique then St := Trim(St);
  Result := St;
end;

function RecupLibelle_ChampDollar(NomChamp: string; var TypeChamp: string): string;
var Q: TQuery;
  cpt: integer;
begin
  Result := '';
  if copy(NomChamp, 1, 3) = '$$_' then
  begin
    if copy(NomChamp, 1, 6) = '$$_ART' then
    begin
      for cpt := Low(TabChampSuppArticle) to High(TabChampSuppArticle) do
      begin
        if TabChampSuppArticle[cpt, 1] = NomChamp then
        begin
          TypeChamp := TabChampSuppArticle[cpt, 2];
          Result := TabChampSuppArticle[cpt, 3];
          Break;
        end;
      end;
    end
    else
      if copy(NomChamp, 1, 10) = '$$_QTEFACT' then
    begin
      for cpt := Low(TabChampSuppQTEFACT) to High(TabChampSuppQTEFACT) do
      begin
        if TabChampSuppQTEFACT[cpt, 1] = NomChamp then
        begin
          TypeChamp := TabChampSuppQTEFACT[cpt, 2];
          Result := TabChampSuppQTEFACT[cpt, 3];
          Break;
        end;
      end;
    end;
  end
  else
  begin
    Q := OpenSQL('SELECT DH_TYPECHAMP,DH_LIBELLE FROM DECHAMPS WHERE DH_NOMCHAMP="' + NomChamp + '"', True);
    TypeChamp := Q.FindField('DH_TYPECHAMP').AsString;
    Result := Q.FindField('DH_LIBELLE').AsString;
    Ferme(Q);
  end;
end;

function CreationJointure(Prefixe: string; var ListeDesChamps: string): string;
var StJointure: string;
  Num: integer;
begin
  StJointure := '';
  if Pos('GA_', ListeDesChamps) > 0 then
    StJointure := 'LEFT JOIN ARTICLE ON GA_ARTICLE=' + Prefixe + '_ARTICLE ';
  if Pos('GA2_', ListeDesChamps) > 0 then
    StJointure := StJointure + 'LEFT JOIN ARTICLECOMPL ON GA2_ARTICLE=' + Prefixe + '_ARTICLE ';
  if Pos('GPA_', ListeDesChamps) > 0 then
    StJointure := StJointure + 'LEFT JOIN PIECEADRESSE ON GPA_NATUREPIECEG=GP_NATUREPIECEG AND ' +
      'GPA_SOUCHE=GP_SOUCHE AND GPA_NUMERO=GP_NUMERO AND GPA_INDICEG=GP_INDICEG ' +
      'AND GPA_NUMLIGNE=0' + ' ';
  if Pos('$$_', ListeDesChamps) > 0 then
  begin
    for Num := 1 to 5 do
    begin
      if Pos('$$_ARTLIBELLEDIM' + IntToStr(Num), ListeDesChamps) > 0 then
      begin
        if Pos('LEFT JOIN ARTICLE ON', StJointure) = 0 then
          StJointure := StJointure + 'LEFT JOIN ARTICLE ON GA_ARTICLE=' + Prefixe + '_ARTICLE ';
        ListeDesChamps := StringReplace(ListeDesChamps, '$$_ARTLIBELLEDIM' + IntToStr(Num),
          'D' + IntToStr(Num) + '.GDI_LIBELLE AS ARTLIBELLEDIM' + IntToStr(Num), []);
        StJointure := StJointure +
          'LEFT JOIN DIMENSION D' + IntToStr(Num) + ' on D' + IntToStr(Num) + '.GDI_TYPEDIM="DI' +
          IntToStr(Num) + '" and D' + IntToStr(Num) + '.GDI_GRILLEDIM=GA_GRILLEDIM' +
          IntToStr(Num) + ' and D' + IntToStr(Num) + '.GDI_CODEDIM=GA_CODEDIM' + IntToStr(Num) + ' ';
      end;
    end;
    if Pos('$$_QTEFACT', ListeDesChamps) > 0 then
    begin
      if Pos('GL_QTEFACT', ListeDesChamps) = 0 then
      begin
        if ListeDesChamps = '' then ListeDesChamps := 'GL_QTEFACT'
        else ListeDesChamps := ListeDesChamps + ',GL_QTEFACT';
      end;
      if Pos('$$_QTEFACTABS', ListeDesChamps) > 0 then Delete(ListeDesChamps, Pos('$$_QTEFACTABS', ListeDesChamps), 14);
      if Pos('$$_QTEFACTSIGNE', ListeDesChamps) > 0 then Delete(ListeDesChamps, Pos('$$_QTEFACTSIGNE', ListeDesChamps), 16);
    end;
    if Pos('$$_SUPPQTEFACTZERO', ListeDesChamps) > 0 then Delete(ListeDesChamps, Pos('$$_SUPPQTEFACTZERO', ListeDesChamps), 19);
    if Copy(ListeDesChamps, Length(ListeDesChamps), 1) = ',' then Delete(ListeDesChamps, Length(ListeDesChamps), 1);
  end;
  Result := StJointure;
end;

procedure TOM_EXPORTASCII_ENTETE.AffectePrefixes;
begin
  Champs_dollars := '';
  if (LeRange = 'MPT') or (LeRange = 'GCF') then // Proposition de transferts Mode
    Prefixes := '"GA","GTL"'
  else
    if (LeRange = 'DIS') then //Export des stocks
    Prefixes := '"GQ"'
  else // 'TID' 'PAC' 'PVT'  Transferts + Pièces d'achat + Pièces de Ventes
  begin
    Prefixes := '"GP","GL","GA"';
    if GetParamSoc('SO_ARTLOOKORLI') then Prefixes := Prefixes + ',"GA2"';
    Prefixes := Prefixes + ',"GPA"';
  end;
end;

procedure TOM_EXPORTASCII_ENTETE.OnNewRecord;
begin
  inherited;
  CodeExport := '';
  GCHAMPDISPO.VidePile(True);
  GCHAMPSELECT.VidePile(True);
  ChargeGrid;
end;

procedure TOM_EXPORTASCII_ENTETE.OnDeleteRecord;
begin
  inherited;
  SupprimeAncienChamps;
  GCHAMPDISPO.VidePile(True);
  GCHAMPSELECT.VidePile(True);
  TFFiche(Ecran).Retour := 'DELETE';
end;

procedure TOM_EXPORTASCII_ENTETE.OnUpdateRecord;
var i: integer;
begin
  LastError := 0;
  inherited;
  // Mise à jour de la Tob "entête" avec les zones de saisie
  if ControleDonnees then
  begin
    // Récupération de la TOB des champs à reprendre
    for i := 0 to TOB_GCHAMPSELECT.Detail.Count - 1 do
    begin
      with TOB_GCHAMPSELECT.Detail[i] do
      begin
        PutValue('GEX_NATURE', LeRange);
        PutValue('GEX_CODEEX', CodeExport);
        PutValue('GEX_LENCHAMP', Valeur(GCHAMPSELECT.Cells[Col_LENCHAMP, i + 1]));
        PutValue('GEX_DEBUT', Valeur(GCHAMPSELECT.Cells[Col_DEBUTCHAMP, i + 1]));
      end;
    end;

    // Sauvegarde des champs sélectionnés
    if (DS.State in [dsEdit]) then SupprimeAncienChamps;
    TOB_GCHAMPSELECT.InsertDB(nil, True);

    TFFiche(Ecran).Retour := CodeExport;
  end;
end;

procedure TOM_EXPORTASCII_ENTETE.OnLoadRecord;
begin
  inherited;
  CodeExport := GetControlText('GAS_CODEEX');
  ChargeGrid;
end;

procedure TOM_EXPORTASCII_ENTETE.OnArgument(S: string);
begin
  inherited;
  LeRange := TFFiche(Ecran).FRange;
  AffectePrefixes;
  InitFiche;
end;

procedure TOM_EXPORTASCII_ENTETE.OnClose;
begin
  inherited;
  DestroyFiche;
end;

procedure TOM_EXPORTASCII_ENTETE.SetLastError(Num: integer; ou: string);
begin
  if ou <> '' then SetFocusControl(ou);
  LastError := -1;
  PGIError(TraduireMemoire(TexteMessage[Num]), Ecran.Caption);
end;

procedure TOM_EXPORTASCII_ENTETE.InitFiche;
begin
  TOB_GCHAMPDISPO := Tob.Create('DECHAMPS', nil, -1);
  TOB_GCHAMPSELECT := Tob.Create('Table EXPORTASCII', nil, -1);

  GCHAMPDISPO := THGrid(GetControl('GCHAMPDISPO'));
  GCHAMPSELECT := THGrid(GetControl('GCHAMPSELECT'));
end;

procedure TOM_EXPORTASCII_ENTETE.DestroyFiche;
begin
  GCHAMPDISPO.VidePile(True);
  GCHAMPSELECT.VidePile(True);

  if TOB_GCHAMPDISPO <> nil then TOB_GCHAMPDISPO.Free;
  if TOB_GCHAMPSELECT <> nil then TOB_GCHAMPSELECT.Free;
end;

procedure TOM_EXPORTASCII_ENTETE.ChargeGrid;
var Q: TQuery;
  TOBTemp, TobChamp, TobChampSelect: TOB;
  cpt: integer;
begin
  if LeRange = '' then
  begin
    SetLastError(1, '');
    exit;
  end;
  TOB_GCHAMPDISPO.ClearDetail;
  TOB_GCHAMPSELECT.ClearDetail;
  Q := OpenSQL('SELECT * FROM EXPORTASCII WHERE GEX_NATURE="' + LeRange + '" AND GEX_CODEEX="' + CodeExport + '" ORDER BY GEX_POSITION', True);
  Q.First;
  while not Q.Eof do
  begin
    TobChampSelect := TOB.Create('EXPORTASCII', TOB_GCHAMPSELECT, -1);
    TobChampSelect.PutValue('GEX_NATURE', Q.FindField('GEX_NATURE').AsString);
    TobChampSelect.PutValue('GEX_CODEEX', Q.FindField('GEX_CODEEX').AsString);
    TobChampSelect.PutValue('GEX_NOMCHAMP', Q.FindField('GEX_NOMCHAMP').AsString);
    TobChampSelect.PutValue('GEX_LENCHAMP', Q.FindField('GEX_LENCHAMP').AsInteger);
    TobChampSelect.PutValue('GEX_VALDEFAUT', Q.FindField('GEX_VALDEFAUT').Value);
    TobChampSelect.PutValue('GEX_POSITION', Q.FindField('GEX_POSITION').AsInteger);
    TobChampSelect.PutValue('GEX_DEBUT', Q.FindField('GEX_DEBUT').AsInteger);
    Q.Next;
  end;
  Ferme(Q);

  Q := OpenSQL('SELECT DH_NOMCHAMP,DH_TYPECHAMP,DH_LIBELLE FROM DECHAMPS WHERE DH_PREFIXE IN (' +
    Prefixes + ') AND DH_CONTROLE LIKE "%L%" ORDER BY DH_NOMCHAMP', True);
  Q.First;
  while not Q.Eof do
  begin
    TOBTemp := TOB_GCHAMPSELECT.FindFirst(['GEX_NATURE', 'GEX_CODEEX', 'GEX_NOMCHAMP'], [LeRange, CodeExport, Q.FindField('DH_NOMCHAMP').AsString], True);
    if TOBTemp = nil then
    begin
      TobChamp := Tob.Create('DECHAMPS', TOB_GCHAMPDISPO, -1);
      TobChamp.PutValue('DH_NOMCHAMP', Q.FindField('DH_NOMCHAMP').AsString);
      TobChamp.PutValue('DH_TYPECHAMP', Q.FindField('DH_TYPECHAMP').AsString);
      TobChamp.PutValue('DH_LIBELLE', Q.FindField('DH_LIBELLE').AsString);
    end;
    Q.Next;
  end;
  Ferme(Q);

  //Ajoute les champs $$ concernant les articles
  if Pos('"GA"', Prefixes) > 0 then
  begin
    for cpt := Low(TabChampSuppArticle) to High(TabChampSuppArticle) do
    begin
      TOBTemp := TOB_GCHAMPSELECT.FindFirst(['GEX_NATURE', 'GEX_CODEEX', 'GEX_NOMCHAMP'], [LeRange, CodeExport, TabChampSuppArticle[cpt, 1]], True);
      if TOBTemp = nil then
      begin
        TobChamp := Tob.Create('DECHAMPS', TOB_GCHAMPDISPO, -1);
        TobChamp.PutValue('DH_NOMCHAMP', TabChampSuppArticle[cpt, 1]);
        TobChamp.PutValue('DH_TYPECHAMP', TabChampSuppArticle[cpt, 2]);
        TobChamp.PutValue('DH_LIBELLE', TabChampSuppArticle[cpt, 3]);
      end;
    end;
  end;

  //Ajoute les champs $$ concernant les lignes
  if Pos('"GL"', Prefixes) > 0 then
  begin
    for cpt := Low(TabChampSuppQTEFACT) to High(TabChampSuppQTEFACT) do
    begin
      TOBTemp := TOB_GCHAMPSELECT.FindFirst(['GEX_NATURE', 'GEX_CODEEX', 'GEX_NOMCHAMP'], [LeRange, CodeExport, TabChampSuppQTEFACT[cpt, 1]], True);
      if TOBTemp = nil then
      begin
        TobChamp := Tob.Create('DECHAMPS', TOB_GCHAMPDISPO, -1);
        TobChamp.PutValue('DH_NOMCHAMP', TabChampSuppQTEFACT[cpt, 1]);
        TobChamp.PutValue('DH_TYPECHAMP', TabChampSuppQTEFACT[cpt, 2]);
        TobChamp.PutValue('DH_LIBELLE', TabChampSuppQTEFACT[cpt, 3]);
      end;
    end;
  end;

  AfficheGrid;
end;

procedure TOM_EXPORTASCII_ENTETE.AfficheGrid;
begin
  TOB_GCHAMPDISPO.PutGridDetail(GCHAMPDISPO, False, False, 'DH_NOMCHAMP;DH_TYPECHAMP;DH_LIBELLE', True);
  TOB_GCHAMPSELECT.PutGridDetail(GCHAMPSELECT, False, False, 'GEX_NOMCHAMP;GEX_DEBUT;GEX_LENCHAMP', True);
  GCHAMPSELECT.ColLengths[Col_NOMCHAMP] := -1;
  GCHAMPSELECT.ColLengths[Col_LENCHAMP] := 4;
  GCHAMPSELECT.ColTypes[Col_LENCHAMP] := 'R';
  GCHAMPSELECT.Col := Col_LENCHAMP;
end;

procedure TOM_EXPORTASCII_ENTETE.BClick(NameBouton: string);
begin
  if NameBouton = 'BRIGHT' then AjouteUnChamp
  else
    if NameBouton = 'BLEFT' then SupprimeUnChamp
  else
    if NameBouton = 'BUP' then DeplaceChampVersLeHaut
  else
    if NameBouton = 'BDOWN' then DeplaceChampVersLeBas
  else
    if NameBouton = 'BVALDEFAUT' then AppelSaisieValDefautASCII(TOB_GCHAMPSELECT);
  if not (DS.State in [dsInsert, dsEdit]) then DS.Edit;
end;

procedure TOM_EXPORTASCII_ENTETE.AjouteUnChamp;
var TobChampSelect, TobFille: TOB;
  IndiceChampRow: integer;
begin
  if GCHAMPDISPO.Row < 0 then exit;
  if TOB_GCHAMPDISPO.detail.count <= 0 then exit;

  //Sauvegarde les infos des champs sélectionnés
  if GCHAMPSELECT.Cells[1, 2] <> '' then
    TOB_GCHAMPSELECT.GetGridDetail(GCHAMPSELECT, GCHAMPSELECT.RowCount - 1, 'EXPORTASCII', ';GEX_NOMCHAMP;GEX_DEBUT;GEX_LENCHAMP');

  // récupération de la ligne sélectionnée
  IndiceChampRow := GCHAMPDISPO.Row;
  TobChampSelect := TOB_GCHAMPDISPO.FindFirst(['DH_NOMCHAMP'], [GCHAMPDISPO.Cells[1, IndiceChampRow]], False);

  // Création d'une ligne dans la TOB de reprise
  TobFille := TOB.Create('EXPORTASCII', TOB_GCHAMPSELECT, -1);
  TobFille.PutValue('GEX_NOMCHAMP', TobChampSelect.GetValue('DH_NOMCHAMP'));
  TobFille.PutValue('GEX_LENCHAMP', LongueurParDefautChamp(TobChampSelect.GetValue('DH_TYPECHAMP')));
  TobFille.PutValue('GEX_POSITION', TOB_GCHAMPSELECT.Detail.Count - 1);
  TobFille.PutValue('GEX_DEBUT', 1);

  // Suppression de la ligne sélectionnée dans la liste des champs
  GCHAMPDISPO.DeleteRow(IndiceChampRow);
  TobChampSelect.Free;

  // Affichage
  TOB_GCHAMPSELECT.PutGridDetail(GCHAMPSELECT, False, False, 'GEX_NOMCHAMP;GEX_DEBUT;GEX_LENCHAMP', True);

  GCHAMPSELECT.Row := TOB_GCHAMPSELECT.detail.count;
  if IndiceChampRow > 1 then
  begin
    if IndiceChampRow > TOB_GCHAMPDISPO.detail.count then GCHAMPDISPO.Row := TOB_GCHAMPDISPO.detail.count
    else GCHAMPDISPO.Row := IndiceChampRow;
  end;
  GCHAMPSELECT.Col := Col_LENCHAMP;
end;

function TOM_EXPORTASCII_ENTETE.LongueurParDefautChamp(TypeChamp: string): integer;
var LenDefault: integer;
begin
  if Pos('VARCHAR', TypeChamp) > 0 then LenDefault := StrToInt(Copy(TypeChamp, 9, Length(TypeChamp) - 9))
  else
    if (TypeChamp = 'INTEGER') or (TypeChamp = 'DOUBLE') or (TypeChamp = 'RATE') then LenDefault := 6
  else
    if TypeChamp = 'COMBO' then LenDefault := 3
  else
    if TypeChamp = 'DATE' then LenDefault := 10
  else
    if TypeChamp = 'BOOLEAN' then LenDefault := 1
  else
    if (TypeChamp = 'BLOB') or (TypeChamp = 'DATA') then LenDefault := 200
  else
    if Pos('LONGCHAR', TypeChamp) > 0 then LenDefault := StrToInt(Copy(TypeChamp, 10, Length(TypeChamp) - 10))
  else
    if Pos('CHAR', TypeChamp) > 0 then LenDefault := StrToInt(Copy(TypeChamp, 6, Length(TypeChamp) - 6))
  else LenDefault := 0;
  Result := LenDefault;
end;

procedure TOM_EXPORTASCII_ENTETE.SupprimeUnChamp;
var TobChampSelect, TobFille: TOB;
  TypeChamp, LibelleChamp: string;
  IndiceChampRow: integer;
begin
  if GCHAMPSELECT.Row < 0 then exit;
  if TOB_GCHAMPSELECT.detail.count <= 0 then exit;
  // récupération de la ligne sélectionnée
  IndiceChampRow := GCHAMPSELECT.Row;
  TobChampSelect := TOB_GCHAMPSELECT.Detail[IndiceChampRow - 1];
  // Création d'une ligne dans la TOB de reprise
  TobFille := Tob.Create('DECHAMPS', TOB_GCHAMPDISPO, -1);
  TobFille.PutValue('DH_NOMCHAMP', TobChampSelect.GetValue('GEX_NOMCHAMP'));
  LibelleChamp := RecupLibelle_ChampDollar(TobChampSelect.GetValue('GEX_NOMCHAMP'), TypeChamp);
  TobFille.PutValue('DH_TYPECHAMP', TypeChamp);
  TobFille.PutValue('DH_LIBELLE', LibelleChamp);
  // Suppression de la ligne sélectionnée dans la liste des champs
  TOB_GCHAMPSELECT.Detail[IndiceChampRow - 1].free;
  // Affichage
  AfficheGrid;
  GCHAMPDISPO.Row := TOB_GCHAMPDISPO.detail.count;
  if IndiceChampRow > 1 then
  begin
    if IndiceChampRow > TOB_GCHAMPSELECT.detail.count then GCHAMPSELECT.Row := TOB_GCHAMPSELECT.detail.count
    else GCHAMPSELECT.Row := IndiceChampRow;
  end;
  NumeroteLignesTOB(GCHAMPSELECT, TOB_GCHAMPSELECT);
end;

procedure TOM_EXPORTASCII_ENTETE.DeplaceChampVersLeHaut;
var TOBTemp: TOB;
  i: integer;
begin
  i := GCHAMPSELECT.Row;
  if i <= 1 then exit;
  if TOB_GCHAMPSELECT.detail.count <= 1 then exit;
  // Permutation dans le grid
  GCHAMPSELECT.ExchangeRow(i, i - 1);
  GCHAMPSELECT.Row := i - 1;

  // Inversemment des lignes de la TOB
  TOBTemp := TOB_GCHAMPSELECT.Detail[i - 2];
  TOB_GCHAMPSELECT.Detail[i - 2] := TOB_GCHAMPSELECT.Detail[i - 1];
  TOB_GCHAMPSELECT.Detail[i - 1] := TOBTemp;

  NumeroteLignesTOB(GCHAMPSELECT, TOB_GCHAMPSELECT);
end;

procedure TOM_EXPORTASCII_ENTETE.DeplaceChampVersLeBas;
var TOBTemp: TOB;
  i: integer;
begin
  i := GCHAMPSELECT.Row;
  if (i < 0) or (i = TOB_GCHAMPSELECT.detail.count) then exit;
  if TOB_GCHAMPSELECT.detail.count <= 1 then exit;
  // Permutation
  GCHAMPSELECT.ExchangeRow(i, i + 1);
  GCHAMPSELECT.Row := i + 1;

  // Inversemment des lignes de la TOB
  TOBTemp := TOB_GCHAMPSELECT.Detail[i];
  TOB_GCHAMPSELECT.Detail[i] := TOB_GCHAMPSELECT.Detail[i - 1];
  TOB_GCHAMPSELECT.Detail[i - 1] := TOBTemp;

  NumeroteLignesTOB(GCHAMPSELECT, TOB_GCHAMPSELECT);
end;

procedure TOM_EXPORTASCII_ENTETE.NumeroteLignesTOB(GS: THGrid; TOBGrid: TOB);
var i: integer;
  TOBL: TOB;
begin
  for i := 2 to GS.RowCount do
  begin
    TOBL := TOBGrid.Detail[i - 2];
    TOBL.PutValue('GEX_POSITION', i - 2);
  end;
end;

function TOM_EXPORTASCII_ENTETE.ControleDonnees: boolean;
begin
  Result := False;
  if LeRange = '' then exit;
  if (DS.State in [dsInsert]) then
  begin
    CodeExport := GetControlText('GAS_CODEEX');
    if CodeExport = '' then
    begin
      SetLastError(3, 'GAS_CODEEX');
      exit;
    end;
    if GetControlText('GAS_LIBELLE') = '' then
    begin
      SetLastError(5, 'GAS_LIBELLE');
      exit;
    end;
    if ExisteSQL('SELECT GAS_CODEEX FROM EXPORTASCII_ENTETE WHERE GAS_NATURE="' + LeRange + '" AND GAS_CODEEX="' + CodeExport + '"') then
    begin
      SetLastError(2, 'GAS_CODEEX');
      exit;
    end;
  end;
  if GetControlText('GAS_SEPARATEUR') = '' then
  begin
    SetLastError(6, 'GAS_SEPARATEUR');
    exit;
  end;
  if GetControlText('GAS_FORMATFIXE') = 'X' then
  begin
    if GetControlText('GAS_NUMERIQUE') = '' then
    begin
      SetLastError(7, 'GAS_NUMERIQUE');
      exit;
    end;
    if GetControlText('GAS_ALPHA') = '' then
    begin
      SetLastError(8, 'GAS_ALPHA');
      exit;
    end;
  end;
  if TOB_GCHAMPSELECT.Detail.count <= 0 then
  begin
    SetLastError(4, 'GAS_CODEEX');
    exit;
  end;
  if not TOB_GCHAMPSELECT.IsOneModifie then exit;
  Result := True;
end;

procedure TOM_EXPORTASCII_ENTETE.SupprimeAncienChamps;
var SQL: string;
begin
  SQL := 'Delete From EXPORTASCII WHERE GEX_NATURE="' + LeRange + '" AND GEX_CODEEX="' + CodeExport + '"';
  ExecuteSQL(SQL);
end;

//////////////////////////////////////////////////////////////////////////////

procedure AGLBClick(Parms: array of variant; nb: integer);
var F: TForm;
  OM: TOM;
begin
  F := TForm(Longint(Parms[0]));
  if (F is TFFiche) then OM := TFFiche(F).OM else OM := nil;
  if (OM is TOM_EXPORTASCII_ENTETE) then TOM_EXPORTASCII_ENTETE(OM).BClick(string(Parms[1]));
end;

initialization
  registerclasses([TOM_EXPORTASCII_ENTETE]);
  RegisterAglProc('BClick', True, 1, AGLBClick);
end.
