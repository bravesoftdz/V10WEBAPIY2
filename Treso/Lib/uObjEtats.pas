{Génération d'un état à partir d'une TOB ou d'une THGRID (compatible eAGL
--------------------------------------------------------------------------------------
  Version    |  Date   | Qui |   Commentaires
--------------------------------------------------------------------------------------
 8.01.001.001  19/12/03   JP   Création de l'unité
 8.01.001.025  19/07/07   SB   Correction récup alignement de la grille en génération auto
 8.10.001.012  04/10/07   JP   Gestion du DBIndicator pour les alignements
 8.10.004.001  06/11/07   GCO  Ajout du format "C" pour les VAL ayant +,-,*,/   
--------------------------------------------------------------------------------------}

unit uObjEtats;

interface

uses
  Classes, SysUtils, HCtrls, uTOB;

type
  TTypeAlign  = (ali_Centre, ali_Droite, ali_Gauche);

  TObjEtats = class(TObject)
  private
    FTobImp   : TOB;
    FTobCri   : TOB;
    TobDepart : TOB;
    AvecGrid  : Boolean;
    FTitreFen : string; {Titre de la fenêtre d'aperçu}
    FNbColonne : Byte;
    function  GetCriteres  : string;
    function  GetTitreEtat : string;
    procedure SetTitreEtat (Value : string);
    procedure RemplitTob   (Grille : THGrid = nil);
    procedure InitEtat     ;
    {Renvoie le nombre de champs d'une tob (réels et virtuels) et met à jour NbColonne
    function  GetNbChamp   (T : TOB) : Byte;}
    {Récupération du format d'affichage en fonction du type du champ}
    function  GetFormat    (NumChp : Integer; Fille : TOB) : string;
    {Récupération de l'alignement d'affichage en fonction du type du champ}
    function  GetAlignement(NumChp : Integer; Fille : TOB) : TTypeAlign;

  public
    ModeSilence : Boolean;
    {Les quatres booléens du LanceEtatTob}
    FDuplicata   : Boolean;
    FApercu      : Boolean;
    FDeuxPages   : Boolean;
    FListeExport : Boolean;

    constructor Create(aTitre : string; Grille : THGrid); overload;
    constructor Create(aTitre : string; LaTob  : TOB   ); overload;
    destructor Destroy; override;

    function  Imprimer : Boolean;
    {Mise à jour du titre de l'une des colonnes}
    procedure MajTitre(NumCol : Integer; Valeur  : string);
    {Alignement des colonnes de l'état}
    procedure MajAlign  (NumCol : Integer; aAlign  : TTypeAlign);
    {Format des colonnes de l'état}
    procedure MajFormat (NumCol : Integer; aFormat : string);

    class function GenereEtatTob    (aTob  : TOB   ; aTitre : string; AvecMsg : Boolean = True) : Boolean;
    class function GenereEtatGrille (aGrid : THGrid; aTitre : string; AvecMsg : Boolean = True) : Boolean;

    property TobImpression : TOB    read FTobImp;
    property TobCriteres   : TOB    read FTobCri;
    {Titre mis dans l'entête de l'état}
    property TitreEtat     : string read GetTitreEtat write SetTitreEtat;
    {Nombre de colonnes de l'état}
    property NbColonne     : Byte   read FNbColonne;
  end;


implementation


uses
  {$IFDEF EAGLCLIENT}
  UtileAgl,
  {$ELSE}
  EdtREtat,
  {$ENDIF EAGLCLIENT}
  HMsgBox, HEnt1, UTobDebug;


{---------------------------------------------------------------------------------------}
class function TObjEtats.GenereEtatTob(aTob  : TOB; aTitre : string; AvecMsg : Boolean = True) : Boolean;
{---------------------------------------------------------------------------------------}
var
  aObj : TObjEtats;
begin
  try
    aObj := TObjEtats.Create(aTitre, aTob);
    try
      {Si false, pas d'affichage des message}
      aObj.ModeSilence := not AvecMsg;
      {Lancement de l'état}
      Result := aObj.Imprimer;
    finally
      FreeAndNil(aObj);
    end;
  except
    on E : Exception do begin
      if AvecMsg then PGIError(TraduireMemoireToken('Impossible de générer l''état'));
      Result := False;
    end;
  end;
end;

{---------------------------------------------------------------------------------------}
class function TObjEtats.GenereEtatGrille(aGrid : THGrid; aTitre : string; AvecMsg : Boolean = True) : Boolean;
{---------------------------------------------------------------------------------------}
var
  aObj : TObjEtats;
begin
  try
    aObj := TObjEtats.Create(aTitre, aGrid);
    try
      {Si false, pas d'affichage des message}
      aObj.ModeSilence := not AvecMsg;
      {Lancement de l'état}
      Result := aObj.Imprimer;
    finally
      FreeAndNil(aObj);
    end;
  except
    on E : Exception do begin
      PGIError(TraduireMemoireToken('Impossible de générer l''état'));
      Result := False;
    end;
  end;
end;

{---------------------------------------------------------------------------------------}
constructor TObjEtats.Create(aTitre : string; Grille : THGrid);
{---------------------------------------------------------------------------------------}
begin
  FTitreFen := aTitre;
  AvecGrid  := True;
  FNbColonne := 0;
  {Préparation de la Tob d'impression}
  RemplitTob(Grille);
end;

{---------------------------------------------------------------------------------------}
constructor TObjEtats.Create(aTitre : string; LaTob : TOB);
{---------------------------------------------------------------------------------------}
begin
  AvecGrid  := False;
  TobDepart := LaTob;
  FTitreFen := aTitre;
  FNbColonne := 0;
  {Préparation de la Tob d'impression}
  RemplitTob;
end;

{---------------------------------------------------------------------------------------}
destructor TObjEtats.Destroy;
{---------------------------------------------------------------------------------------}
begin
  if Assigned(FTobImp) then FreeAndNil(FTobImp);
  if Assigned(FTobCri) then FreeAndNil(FTobCri);
  inherited;
end;

{---------------------------------------------------------------------------------------}
{***********A.G.L.***********************************************
Auteur  ...... : JP
Créé le ...... : 16/02/2007
Modifié le ... :   /  /    
Description .. : - LG - 16/02/2007 - ajout d'un readtoken pour gere les 
Suite ........ : format blanc si null
Mots clefs ... : 
*****************************************************************}
procedure TObjEtats.RemplitTob(Grille : THGrid = nil);
{---------------------------------------------------------------------------------------}
var
  i : Integer;
  n : Integer;
  p : Integer;
  k : Integer;
  j : Integer;
  d : Integer; {JP 04/10/07 : pour gérer le DBIndicator}
  F : TOB;
  s : string;
begin
  {Les quatres booléens du LanceEtatTob}
  FDuplicata   := False;
  FApercu      := True;
  FDeuxPages   := False;
  FListeExport := False;

  FTobImp := TOB.Create('IMPSTANDARD', nil, -1);
  FTobCri := TOB.Create('IMPCRITERES', nil, -1);
  {Par défaut le titre de l'état est le le Titre de la fenêtre d'édition}
  TitreEtat := FTitreFen;

  {Génération des la tob d'impression à partir de la grille}
  if AvecGrid then begin
    k := -1;
    FNbColonne := 0;
    {On n'imprime que les colonnes visibles : recherche du nombre de colonnes visibles et de
     l'indice de la dernière colonne visible}
    if Grille.DBIndicator then i := 1
                          else i := 0;
    d := i; {JP 04/10/07 : gestion du DBIndicator pour les alignements}

    for n := i to Grille.ColCount - 1 do
      if not((Grille.ColWidths[n] <= 1) and (Grille.ColLengths[n] <= 1)) then begin
        Inc(FNbColonne);
        if FNbColonne > 11 then Break;
        k := n;
      end;

    if (FNbColonne > 11) then begin
      if not ModeSilence then
        PGIInfo(TraduireMemoire('La grille a plus de 11 colonnes. Seules les 11 premières seront imprimées.'));
      FNbColonne := 11;
    end;

    {Remplissage de la tob d'impression}
    for p := 1 to Grille.RowCount - 1 do begin
      F := TOB.Create('_IMPSTANDARD', FTobImp, -1);
      {Pour les regroupements dans l'état}
      F.AddChampSupValeur('REG1', '1');
      F.AddChampSupValeur('REG2', '1');
      F.AddChampSupValeur('REG3', '1');
      F.AddChampSupValeur('REG4', '1');
      F.AddChampSupValeur('REG5', '1');
      F.AddChampSupValeur('REG6', '1');

      j := 0;
      for n := i to k do
        if not((Grille.ColWidths[n] <= 1) and (Grille.ColLengths[n] <= 1)) then begin
          {Récupération de l'éventuel format}
          if Grille.ColFormats[n] <> '' then
           begin
            s := Grille.ColFormats[n] ;
            s := '"' + ReadTokenSt(s) + '"' ;
           end
            else
             begin
               // GCO - 06/11/2007 - FQ 21723
               if (Pos('+', Grille.CellValues[n, p]) > 0) or
                  (Pos('-', Grille.CellValues[n, p]) > 0) or
                  (Pos('*', Grille.CellValues[n, p]) > 0) or
                  (Pos('/', Grille.CellValues[n, p]) > 0) then
                 s := '"C"'
               else
                 s := '';
             end;

          F.AddChampSupValeur('VAL' + IntToStr(j), s + Grille.CellValues[n, p]);
          Inc(j);
        end;
    end;

    {Met les valeurs par défaut des alignements des champs}
    InitEtat;
    {Remplissage de la tob des critères}
    j := 0;
    for n := i to k do
      if not((Grille.ColWidths[n] <= 1) and (Grille.ColLengths[n] <= 1)) then begin
        case Grille.ColAligns[n] of
          {19/07/07 SBO Correction récup alignement de la grille en génération auto
           04/10/07 JP : gestion du DBIndicator dans l'alignement}
          taLeftJustify  : MajAlign(j + d - i, ali_Gauche);//MajAlign(n - i, ali_Gauche);
          taRightJustify : MajAlign(j + d - i, ali_Droite);//MajAlign(n - i, ali_Droite);
          taCenter       : MajAlign(j + d - i, ali_Centre);//MajAlign(n - i, ali_Centre);
        else
          MajAlign(j, ali_Gauche);
        end;

        FTobCri.AddChampSupValeur('COL' + IntToStr(j), Grille.Cells[n, 0]);
        Inc(j);
      end;
  end

  {Si l'on part d'une tob}
  else begin
    if TobDepart.Detail.Count = 0 then Exit;

    if TobDepart.Detail[0].GetChampCount(ttcAll) > 11 then begin
      if not ModeSilence then
        PGIInfo(TraduireMemoire('La Tob a plus de 11 champs. Seuls les 11 premiers seront imprimés.'));
      FNbColonne := 11;
    end
    else
      FNbColonne := TobDepart.Detail[0].GetChampCount(ttcAll);

    {Remplissage de la tob d'impression}
    for p := 0 to TobDepart.Detail.Count - 1 do begin
      F := TOB.Create('_IMPSTANDARD', FTobImp, -1);
      {Pour les regroupements dans l'état}
      F.AddChampSupValeur('REG1', '1');
      F.AddChampSupValeur('REG2', '1');
      F.AddChampSupValeur('REG3', '1');
      F.AddChampSupValeur('REG4', '1');
      F.AddChampSupValeur('REG5', '1');
      F.AddChampSupValeur('REG6', '1');
      k := 0;
      j := 0;
      {Gestion des champs réels}
      for n := 1 to 11 do
        if TobDepart.Detail[p].GetNomChamp(n) <> '' then begin
          s := GetFormat(j, TobDepart.Detail[p]);
          F.AddChampSupValeur('VAL' + IntToStr(j), s + string(TobDepart.Detail[p].GetValeur(n)));
          Inc(j);
        end
        else begin
          {Récupération du nombre de colonnes disponibles pour les champs virtuels}
          k := FNbColonne - n + 1;
          Break;
        end;

      {Gestion des champs virtuels}
      for n := 1000 to 1000 + k - 1 do
        if TobDepart.Detail[p].GetNomChamp(n) <> '' then begin
          s := GetFormat(j, TobDepart.Detail[p]);
          F.AddChampSupValeur('VAL' + IntToStr(j), s + string(TobDepart.Detail[p].GetValeur(n)));
          Inc(j);
        end
        else
          Break;
    end;

    {Met les valeurs par défaut des alignements des champs}
    InitEtat;
    {Remplissage de la tob des critères}
    k := 0;
    j := 0;
    {Gestion des champs réels}
    for n := 0 to 10 do
      if TobDepart.Detail[0].GetNomChamp(n) <> '' then begin
        MajAlign(j, GetAlignement(j, TobDepart.Detail[0]));
        FTobCri.AddChampSupValeur('COL' + IntToStr(j), TobDepart.Detail[0].GetNomChamp(n));
        Inc(j);
      end
      else begin
        {Récupération du nombre de colonnes disponibles pour les champs virtuels}
        k := FNbColonne - n;
        Break;
      end;

    {Gestion des champs virtuels}
    for n := 1000 to 1000 + k - 1 do
      if TobDepart.Detail[0].GetNomChamp(n) <> '' then begin
        MajAlign(j, GetAlignement(j, TobDepart.Detail[0]));
        FTobCri.AddChampSupValeur('COL' + IntToStr(j), TobDepart.Detail[0].GetNomChamp(n));
        Inc(j);
      end
      else
        Break;
  end;

  {Le critère RUPTURE permet de savoir quelle bande détail imprimer ainsi le nombre de colonnes à afficher}
  FTobCri.AddChampSupValeur('RUPTURE', FNbColonne);
end;

{Met à jour une des valeurs de la Tob des critères
{---------------------------------------------------------------------------------------}
procedure TObjEtats.MajTitre(NumCol : Integer; Valeur : string);
{---------------------------------------------------------------------------------------}
begin
  if FTobCri.FieldExists('COL' + IntToStr(NumCol)) then
    FTobCri.SetString('COL' + IntToStr(NumCol), Valeur)
  else if not ModeSilence then
    PGIError(TraduireMemoire('Le champ "COL' + IntToStr(NumCol) + '" n''existe pas.'));
end;

{---------------------------------------------------------------------------------------}
function TObjEtats.Imprimer : Boolean;
{---------------------------------------------------------------------------------------}
begin
  Result := True;

  if (FNbColonne = 0) then begin
    if not ModeSilence then PGIError(TraduireMemoire('Impossible de récupérer le titre des colonnes.'));
  end
  else begin
    if (FTobImp.Detail.Count = 0) and not ModeSilence then
      PGIBox(TraduireMemoire('Il n''y a pas de données à imprimer.'))
    else begin
     {$IFNDEF TT}
      if V_PGI.SAV then
        TobDebug(FTobImp);
     {$ENDIF}   
      if FNbColonne < 8 then
        LanceEtatTob('E', 'ECT', 'P07', FTobImp, FApercu, FListeExport, FDeuxPages, nil, '', FTitreFen, FDuplicata, 0, GetCriteres)
      else
        LanceEtatTob('E', 'ECT', 'P11', FTobImp, FApercu, FListeExport, FDeuxPages, nil, '', FTitreFen, FDuplicata, 0, GetCriteres);
    end
  end;
end;

{---------------------------------------------------------------------------------------}
function TObjEtats.GetCriteres : string;
{---------------------------------------------------------------------------------------}
var
  n : Integer;
begin
  Result := '';
  for n := 0 to 30 do
    if FTobCri.GetNomChamp(1000 + n) <> '' then
      Result := Result + FTobCri.GetNomChamp(1000 + n) + '=' + string(FTobCri.GetValeur(1000 + n)) + '`';
end;

{---------------------------------------------------------------------------------------}
function TObjEtats.GetTitreEtat : string;
{---------------------------------------------------------------------------------------}
begin
  Result := '';
  if Assigned(FTobCri) then Result := FTobCri.GetString('TITRE');
end;

{---------------------------------------------------------------------------------------}
procedure TObjEtats.SetTitreEtat(Value : string);
{---------------------------------------------------------------------------------------}
begin
  if Assigned(FTobCri) then begin
    if FTobCri.FieldExists('TITRE') then
      FTobCri.SetString('TITRE', Value)
    else
      FTobCri.AddChampSupValeur('TITRE', Value);
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TObjEtats.InitEtat;
{---------------------------------------------------------------------------------------}
var
  n : Integer;
begin
  MajAlign(0, ali_Gauche);
  for n := 1 to FNbColonne do
    MajAlign (n, ali_Droite);
end;

{Alignement des colonnes de l'état
{---------------------------------------------------------------------------------------}
procedure TObjEtats.MajAlign(NumCol : Integer; aAlign : TTypeAlign);
{---------------------------------------------------------------------------------------}

    {------------------------------------------------------}
    function _GetAlign : string;
    {------------------------------------------------------}
    begin
      case aAlign of
        ali_Centre : Result := 'C';
        ali_Droite : Result := 'D';
        ali_Gauche : Result := 'G';
      end;
    end;

begin
  if NumCol >= FNbColonne then Exit;

  if FTobCri.FieldExists('ALI' + IntToStr(NumCol)) then
    FTobCri.SetString('ALI' + IntToStr(NumCol), _GetAlign)
  else
    FTobCri.AddChampSupValeur('ALI' + IntToStr(NumCol), _GetAlign);
end;

{Format des colonnes de l'état
{---------------------------------------------------------------------------------------}
procedure TObjEtats.MajFormat(NumCol: Integer; aFormat: string);
{---------------------------------------------------------------------------------------}
var
  n : Integer;
  s : string;
begin
  if NumCol >= FNbColonne then Exit;

  for n := 0 to FTobImp.Detail.Count - 1 do begin
    s := string(FTobImp.Detail[n].GetValeur(1006 + NumCol));
    if (s <> '') and (aFormat <> '') then
      FTobImp.Detail[n].PutValeur(1006 + NumCol, '"' + aFormat + '"' + s);
  end;
end;

(*
{Renvoie le nombre de champs d'une tob réelle et/ou VIRTUELLE
{---------------------------------------------------------------------------------------}
function TObjEtats.GetNbChamp(T : TOB): Byte;
{---------------------------------------------------------------------------------------}
var
  n : Integer;
begin
  Result := 0;
  for n := 0 to 20 do
    if Trim(T.GetNomChamp(n)) <> '' then Inc(Result)
                                    else Break;
  for n := 1000 to 1020 do
    if Trim(T.GetNomChamp(n)) <> '' then Inc(Result)
                                    else Break;
  FNbColonne := Result;
end;
*)

{Récupération du format à partir du format du champ
{---------------------------------------------------------------------------------------}
function TObjEtats.GetFormat(NumChp : Integer; Fille : TOB): string;
{---------------------------------------------------------------------------------------}
begin
  Result := '';
  if not Assigned(Fille.Items[NumChp]) then Exit;
  case Fille.Items[NumChp].CSType of
    CSTInteger : Result := '"0"';
    CSTDouble  : Result := '"#,##0.00"';
    CSTDate    : Result := '"dd/mm/yyyy"';
  end;
end;

{Récupération de l'alignement d'affichage en fonction du type du champ}
{---------------------------------------------------------------------------------------}
function TObjEtats.GetAlignement(NumChp : Integer; Fille : TOB): TTypeAlign;
{---------------------------------------------------------------------------------------}
begin
  Result := ali_Gauche;
  if not Assigned(Fille.Items[NumChp]) then Exit;
  case Fille.Items[NumChp].CSType of
    CSTInteger : Result := ali_Droite;
    CSTDouble  : Result := ali_Droite;
    CSTDate    : Result := ali_Centre;
    CSTBoolean : Result := ali_Centre;
  end;
end;


end.
