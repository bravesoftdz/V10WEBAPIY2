{-------------------------------------------------------------------------------------
    Version   |  Date   | Qui | Commentaires
--------------------------------------------------------------------------------------
08.10.001.011  02/08/07   JP   Création de l'unité : ajout de dossiers à un cash pooling
--------------------------------------------------------------------------------------}
unit TRADDCASHPOOLING_TOF;

interface

uses
  {$IFDEF EAGLCLIENT}
  MaineAGL,
  {$ELSE}
  FE_MAIN, {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
  {$ENDIF}
  StdCtrls, Controls, UTob, Classes, SysUtils, UTOF, ComCtrls, ParamSoc;


type
  PInfoTW = ^RInfoTW;
  RInfoTW = record
    NoDossier : string;
    NomBase   : string;
    Libelle   : string;
    NomPole   : string;
    MoulineBQ : Boolean;
    CanRemove : Boolean;
    CanPoolRe : Boolean;
  end;

  TOF_TRADDCASHPOOLING = class (TOF)
    procedure OnArgument (S : string); override;
    procedure OnUpdate               ; override;
    procedure OnClose                ; override;
    procedure OnLoad                 ; override;
  private
    TobInfos : TOB;

    procedure RemplirSource(aInfos : TOB);
    procedure RemplirDestination(aInfos : TOB);
    procedure ChargeInfos;
    function  MoulineDossier(NomBase, NoDossier, NomPool : string; MoulineBq : Boolean) : Boolean;
  protected {Contient tous les éléments de gestion des TreeView}
    procedure InitPInfoTW(var aPITW : PInfoTW; aInfos : TOB);
    procedure VideTreeView;
    procedure InitTreeView;
    {Enlève un objet de la ListView aLVS pour le mettre dans l'autre}
    procedure DeplaceObjet(aLVD : TTreeView; aObjS, aObjD : TTreeNode);
  public
    TVSource : TTreeView;
    TVDest   : TTreeView;

    procedure BPrendreOnClick(Sender : TObject);
    procedure BRendreOnClick (Sender : TObject);
    procedure BNewPoolOnClick(Sender : TObject);
    procedure BSupPoolOnClick(Sender : TObject);
    procedure TVDragDrop     (Sender, Source : TObject; X, Y : Integer);
    procedure TVDragOver     (Sender, Source : TObject; X, Y : Integer; State : TDragState; var Accept : Boolean);
  end;

procedure TRLanceFiche_AddCashPooling(Arguments : string);
Function CacheCash : Boolean ;

implementation

uses
  {$IFDEF EAGLCLIENT}
  MenuOLX,
  {$ELSE}
  MenuOLG,
  {$ENDIF EAGLCLIENT}
  HTB97, HEnt1, HCtrls, UtilPGI, URecupDos, HMsgBox;

const
  imi_BQOk   = 85 - 1; {Dossier n'appartenant pas un Pool, mais déjà mouliné}
  imi_NotBQ  = 14 - 1; {Dossier n'appartenant pas un Pool, mais avec BQ_GENERAL au lieu de BQ_CODE}
  imi_Pool   = 44 - 1; {Pool existant}
  imi_NewP   = 43 - 1; {Nouveau Pool}
  imi_Old    = 28 - 1; {Dossier appartenant à un Pool}
  imi_New    = 27 - 1; {Dossier mis dans un pool dans la session actuel}


{---------------------------------------------------------------------------------------}
procedure TRLanceFiche_AddCashPooling(Arguments : string);
{---------------------------------------------------------------------------------------}
var
  Q : TQuery;
  O : Boolean;
  B : string;
begin
  V_PGI.ZoomOle := True;
  try
    AglLanceFiche('TR', 'TRADDCASHPOOLING', '', '', Arguments);
  finally
    V_PGI.ZoomOle := False;
  end;
  {Après le traitement, on s'assure que ll'on peut rester sur la base courante}
  O := False;
  Q := OpenSQL('SELECT SOC_DATA FROM PARAMSOC WHERE SOC_NOM = "SO_TRBASETRESO"', True, -1, '', True);
  B := Q.FindField('SOC_DATA').AsString;
  if not Q.EOF then O := (B <> '') and (B <> V_PGI.SchemaName);
  Ferme(Q);
  {On ne peut plus travailler sur la base courant}
  if O then begin
    PGIInfo(TraduireMemoire('Vous avez rattacher le dossier au pôle "') + B + '".'#13#13 +
            TraduireMemoire('Vous ne pouvez plus travailler dans le dossier courant.') + #13 +
            TraduireMemoire('L''application va être fermée.'), 'Cash Pooling');
    {On se remet sur la mire de la connection}
    FMenuG.FermeSoc;
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRADDCASHPOOLING.InitPInfoTW(var aPITW : PInfoTW; aInfos : TOB);
{---------------------------------------------------------------------------------------}
begin
  if not Assigned(aInfos) then begin
    aPITW^.NoDossier := '';
    aPITW^.NomBase   := '';
    aPITW^.Libelle   := '';
    aPITW^.NomPole   := '';
    aPITW^.MoulineBQ := True;
    aPITW^.CanPoolRe := False;
  end else begin
    aPITW^.NoDossier := aInfos.GetString('SO_NODOSSIER');
    aPITW^.NomBase   := aInfos.GetString('SCHEMANAME');
    aPITW^.Libelle   := aInfos.GetString('SO_LIBELLE');
    aPITW^.NomPole   := aInfos.GetString('SO_TRBASETRESO');
    if aInfos.GetString('SO_TRMOULINEBQCODE') = '' then aPITW^.MoulineBQ := False
                                                   else aPITW^.MoulineBQ := aInfos.GetValue('SO_TRMOULINEBQCODE');
    if aInfos.GetString('SO_TRNOCASHPOOLING') = '' then aPITW^.MoulineBQ := False
                                                   else aPITW^.MoulineBQ := aPITW^.MoulineBQ and
                                                                            not aInfos.GetValue('SO_TRNOCASHPOOLING');
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRADDCASHPOOLING.VideTreeView;
{---------------------------------------------------------------------------------------}

    {---------------------------------------------------------------------}
    procedure _VideTreeView(TW : TTReeView);
    {---------------------------------------------------------------------}
    var
     n   : Integer;
     ITW : PInfoTW;
    begin
      for n := 0 to TW.Items.Count - 1  do begin
        ITW := TW.Items[n].Data;
        if ITW <> nil then Dispose(ITW);
        TW.Items[n].Data := nil;
      end ;
      TW.Items.Clear;
    end;

begin
  _VideTreeView(TVSource);
  _VideTreeView(TVDest);
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRADDCASHPOOLING.InitTreeView;
{---------------------------------------------------------------------------------------}
begin
  {Chargement des images PGI}
  if not Assigned(V_PGI.GraphList) then ChargeImageList;
  {Gestion des 2 TreeView}
  TVSource := (GetControl('TVSOURCE') as TTreeView);
  TVDest   := (GetControl('TVPOLE') as TTreeView);
  TVSource.Images := V_PGI.GraphList;
  TVDest  .Images := V_PGI.GraphList;
  TVDest  .ShowButtons := True;
  TVSource.ShowButtons := False;
  TVDest  .ReadOnly := True;
  TVSource.ReadOnly := True;
  TVDest  .DragMode := dmAutomatic;
  TVSource.DragMode := dmAutomatic;
  TVSource.OnDragDrop := TVDragDrop;
  TVSource.OnDragOver := TVDragOver;
  TVDest  .OnDragDrop := TVDragDrop;
  TVDest  .OnDragOver := TVDragOver;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRADDCASHPOOLING.DeplaceObjet(aLVD : TTreeView; aObjS, aObjD : TTreeNode);
{---------------------------------------------------------------------------------------}

    {-----------------------------------------------------------------------}
    function _PrepareStructure(aNeu : TTreeNode; LibOk : Boolean = True) : PInfoTW;
    {-----------------------------------------------------------------------}
    var
      ITWS : PInfoTW;
    begin
      System.New(Result);
      ITWS := PInfoTW(aNeu.Data);
      Result^.NoDossier := ITWS^.NoDossier;
      Result^.NomBase   := ITWS^.NomBase;
      Result^.Libelle   := ITWS^.Libelle;
      Result^.NomPole   := ITWS^.NomPole;
      Result^.MoulineBQ := ITWS^.MoulineBQ;
      Result^.CanRemove := ITWS^.CanRemove;
      Result^.CanPoolRe := ITWS^.CanPoolRe or (UpperCase(aLVD.Name) = 'TVSOURCE');
      {Suppression / Libération de l'objet source}
      if LibOk then begin
        Dispose(ITWS);
        aNeu.Data := nil;
        FreeAndNil(aNeu);
      end;
    end;

var
  O : TTreeNode;
  P : PInfoTW;
  L : string;
  n : Integer;
begin
  {On est sur un noeud que l'on déplace, il faut déplacer tous les dossiers qui lui sont liés.
   => Appel de la fonction de manière récursive ...}
  if PInfoTW(aObjS.Data)^.CanPoolRe then begin
    for n := TVDest.Items.Count - 1 downto 0 do begin
      O := TVDest.Items[n];
      if not Assigned(O) then Continue;
      {L'élément appartient au pôle que l'on déplace}
      if O.Parent = aObjS then begin
        {Il s'agit du dossier correspondant pôle en lui même ...}
        if PInfoTW(O.Data)^.NomBase = PInfoTW(aObjS.Data)^.NomBase then begin
          {... On se contente de détruire l'élément du TreeView}
          P := PInfoTW(O.Data);
          Dispose(P);
          O.Data := nil;
          FreeAndNil(O);
        end
        else
          {... Sinon, c'est qu'i s'agit d'un dossier du pôle, on le déplace}
          DeplaceObjet(aLVD, O, aObjD);
      end;
    end;
  end;

  L := PInfoTW(aObjS.Data)^.Libelle;
  P := _PrepareStructure(aObjS);

  {On rend un objet}
  if UpperCase(aLVD.Name) = 'TVSOURCE' then begin
    {Ajout de l'objet dans la grille 'SOURCE'}
    O := aLVD.Items.AddChildObject(nil, L, P);
    if P^.MoulineBQ then O.ImageIndex := imi_BQOk
                    else O.ImageIndex := imi_NotBQ;
  end
  else begin
    {Affectation du nom du pool}
    if Assigned(aObjD) then P^.NomPole := PInfoTW(aObjD.Data)^.NomPole
                       else P^.NomPole := P^.NomBase;
    {Ajout de l'objet dans la grille 'POLE'}
    if Assigned(aObjD) then begin
      P^.CanPoolRe := False;
      O := aLVD.Items.AddChildObject(aObjD, L, P);
      O.ImageIndex := imi_New;
    end else begin
      P^.CanPoolRe := True;
      O := aLVD.Items.AddChildObject(aObjD, 'Pôle : ' + P^.NomPole, P);
      O.ImageIndex := imi_NewP;
    end;

    {C'est que l'on vient de créer un nouveau pôle}
    if not Assigned(aObjD) then begin
      O.SelectedIndex := O.ImageIndex;
      {On crée l'élément correspondant au dossier du pôle}
      P := _PrepareStructure(O, False {False : on ne libère pas le Pôle !});
      P^.CanPoolRe := False;
      O := aLVD.Items.AddChildObject(O, P^.Libelle, P);
      O.ImageIndex := imi_New;
    end;
  end;

  O.SelectedIndex := O.ImageIndex;

  TVSource.Refresh;
  TVDest.Refresh;
  TVDest.FullExpand;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRADDCASHPOOLING.OnArgument(S : string);
{---------------------------------------------------------------------------------------}
begin
  inherited;
  Ecran.HelpContext := 50000160;

  SetControlEnabled('BVALIDER', V_PGI.Superviseur);

  (GetControl('BPRENDRE') as TToolbarButton97).OnClick := BPrendreOnClick;
  (GetControl('BRENDRE' ) as TToolbarButton97).OnClick := BRendreOnClick;
  (GetControl('BNEWPOOL') as TToolbarButton97).OnClick := BNewPoolOnClick;
  (GetControl('BSUPPOOL') as TToolbarButton97).OnClick := BSupPoolOnClick;
  {Récupère et initialise les 2 TreeView}
  InitTreeView;
  {Charge la tob contenant toutes les informations des dossiers du regroupement et les 2 TreeView}
  ChargeInfos;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRADDCASHPOOLING.OnClose;
{---------------------------------------------------------------------------------------}
begin
  if Assigned(TobInfos) then FreeAndNil(TobInfos);
  VideTreeView;
  inherited;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRADDCASHPOOLING.OnLoad;
{---------------------------------------------------------------------------------------}
begin
  inherited;
  TVDest.FullExpand;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRADDCASHPOOLING.OnUpdate;
{---------------------------------------------------------------------------------------}
var
  n   : Integer;
  Neu : TTreeNode;
  ITW : PInfoTW;
begin
  inherited;

  InitProgress(TraduireMemoire('Ajout de dossiers au Pôle'), 200);
  try
    BeginTrans;
    try
      for n := 0 to TVDest.Items.Count - 1 do begin
        Neu := TVDest.Items[n];
        if not Assigned(Neu.Data) then Continue;
        ITW := PInfoTW(Neu.Data);
        if not ITW^.CanRemove or ITW^.CanPoolRe then Continue;
          if not MoulineDossier(ITW^.NomBase, ITW^.NoDossier, ITW^.NomPole, ITW^.MoulineBQ) then begin
            RollBack;
            Break;
          end;
      end;

      {Création éventuelle de la banque et agence Comptes courants}
      CreeInfosTitreCourant;
      
      CommitTrans;
    except
      on E : Exception do begin
        PGIError(TraduireMemoire('Traitement interrompu avec le message :') + #13#13 + E.Message, Ecran.Caption);
        RollBack;
      end;
    end;
    {Mise à jour du champ qui a servi de champ temporaire pour les traitements}
    ExecuteSQL('UPDATE BANQUECP SET BQ_REPIMPAYEPRELV = ""');
  finally
    TermineProgress;
  end;

  PgiInfo(TraduireMemoire('Traitement terminé'), Ecran.Caption);
  {On recharge les grilles}
  VideTreeView;
  ChargeInfos;
  TVDest.FullExpand;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRADDCASHPOOLING.BRendreOnClick(Sender : TObject);
{---------------------------------------------------------------------------------------}
var
  Neu : TTreeNode;
  ITW : PInfoTW;
begin
  Neu := TVDest.Selected;
  if not Assigned(Neu) or not Assigned(Neu.Data) then Exit;

  ITW := PInfoTW(Neu.Data);

  {Si l'objet sélectionné appartient à la grille des pools et que le dossier est déjà un dossier en Cash Pooling}
  if not ITW^.CanRemove then Exit;

  {Déplace le dossier dans les arborescences}
  DeplaceObjet(TVSource, Neu, nil);
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRADDCASHPOOLING.BPrendreOnClick(Sender : TObject);
{---------------------------------------------------------------------------------------}
var
  NeuS : TTreeNode;
  NeuD : TTreeNode;
begin
  NeuS := TVSource.Selected;
  if not Assigned(NeuS) or not Assigned(NeuS.Data) then Exit;
  {On ne peut déplacer les dossiers qui non pas été moulinés}
  if NeuS.ImageIndex = imi_NotBQ then Exit;

  NeuD := TVDest.Selected;
  if not Assigned(NeuD) or not Assigned(NeuD.Data) then Exit;

  {On s'assure que l'élément de destination est bien un Pool et non un Dossier}
  if not (NeuD.ImageIndex in [imi_Pool, imi_NewP]) then Exit;

  {Déplace le dossier dans les arborescences}
  DeplaceObjet(TVDest, NeuS, NeuD);
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRADDCASHPOOLING.BNewPoolOnClick(Sender : TObject);
{---------------------------------------------------------------------------------------}
var
  NeuS : TTreeNode;
begin
  NeuS := TVSource.Selected;
  if not Assigned(NeuS) or not Assigned(NeuS.Data) then Exit;
  {On ne peut déplacer les dossiers qui non pas été moulinés}
  if NeuS.ImageIndex = imi_NotBQ then Exit;

  {Déplace le dossier dans les arborescences}
  DeplaceObjet(TVDest, NeuS, nil);
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRADDCASHPOOLING.BSupPoolOnClick(Sender : TObject);
{---------------------------------------------------------------------------------------}
var
  Neu : TTreeNode;
  ITW : PInfoTW;
begin
  Neu := TVDest.Selected;
  if not Assigned(Neu) or not Assigned(Neu.Data) then Exit;

  ITW := PInfoTW(Neu.Data);

  {Si l'objet sélectionné appartient à la grille des pools et que le dossier est déjà un dossier en Cash Pooling}
  if not ITW^.CanPoolRe then Exit;

  {Déplace le dossier dans les arborescences}
  DeplaceObjet(TVSource, Neu, nil);
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRADDCASHPOOLING.TVDragOver(Sender, Source : TObject; X, Y : Integer; State : TDragState; var Accept : Boolean);
{---------------------------------------------------------------------------------------}
var
  Neu : TTreeNode;
  ITW : PInfoTW;
begin
  Accept := False;
  if (Source = Sender) and (UpperCase((Sender as TTreeView).Name) = 'TVSOURCE') then Exit;
  if not (Source is TTreeView) or not (Sender is TTreeView) then Exit;

  {Sur le TreeView de destination, il faut être sur un noeud ...}
  Neu := (Sender as TTreeView).GetNodeAt(X, Y);
  {... de préférence de cash poooling et non sur un pool}
  if Assigned(Neu) and (Neu.ImageIndex in [imi_Old, imi_New]) then Exit;

  Neu := (Source as TTreeView).Selected;
  if not Assigned(Neu) or not Assigned(Neu.Data) then Exit;
  {On ne peut déplacer les dossiers qui non pas été moulinés}
  if Neu.ImageIndex = imi_NotBQ then Exit;

  ITW := PInfoTW(Neu.Data);
  {Si l'objet sélectionné appartient à la grille des pools et que le dossier est déjà un dossier en Cash Pooling}
  if not ITW^.CanRemove then Exit;

  Accept := True;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRADDCASHPOOLING.TVDragDrop(Sender, Source : TObject; X, Y : Integer);
{---------------------------------------------------------------------------------------}
var
  NeuS : TTreeNode;
  NeuD : TTreeNode;
begin
  {Récupération des noeud source et destination}
  NeuD := (Sender as TTreeView).GetNodeAt(X, Y);
  NeuS := (Source as TTreeView).Selected;

  {Déplace le dossier dans les arborescences}
  DeplaceObjet(Sender as TTreeView, NeuS, NeuD);
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRADDCASHPOOLING.RemplirDestination(aInfos : TOB);
{---------------------------------------------------------------------------------------}
var
  ITW : PInfoTW;

    {----------------------------------------------------------------------------}
    function _GetNodePool(NomPool : string) : TTreeNode;
    {----------------------------------------------------------------------------}
    var
      n : Integer;
    begin
      Result := nil;
      for n := 0 to TVDest.Items.Count - 1 do begin
        if TVDest.Items[n].Text = NomPool then begin
          Result := TVDest.Items[n];
          Break;
        end;
      end;
    end;

var
  Neu : TTreeNode;
begin
  System.New(ITW);
  {Remplissage du pointeur}
  InitPInfoTW(ITW, aInfos);
  Neu := _GetNodePool('Pôle : ' + aInfos.GetString('SO_TRBASETRESO'));
  if Assigned(Neu) then begin
    {On ne peut déplacer les éléments figurant déjà dans un pool}
    ITW^.CanRemove := False;
    ITW^.CanPoolRe := False;
    Neu := TVDest.Items.AddChildObject(Neu, ITW^.Libelle, ITW);
    Neu.ImageIndex := imi_Old;
    Neu.SelectedIndex := Neu.ImageIndex;
  end
  else begin
    {On ne peut déplacer un pool}
    ITW^.CanRemove := False;
    ITW^.CanPoolRe := False;
    Neu := TVDest.Items.AddChildObject(nil, 'Pôle : ' + aInfos.GetString('SO_TRBASETRESO'), ITW);
    Neu.ImageIndex := imi_Pool;
    Neu.SelectedIndex := Neu.ImageIndex;
    System.New(ITW);
    {Remplissage du pointeur}
    InitPInfoTW(ITW, aInfos);
    {On ne peut déplacer les éléments figurant déjà dans un pool}
    ITW^.CanRemove := False;
    ITW^.CanPoolRe := False;
    Neu := TVDest.Items.AddChildObject(Neu, ITW^.Libelle, ITW);
    Neu.ImageIndex := imi_Old;
    Neu.SelectedIndex := Neu.ImageIndex;
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_TRADDCASHPOOLING.RemplirSource(aInfos : TOB);
{---------------------------------------------------------------------------------------}
var
  ITW : PInfoTW;
  Neu : TTreeNode;
begin
  System.New(ITW);
  {Remplissage du pointeur}
  InitPInfoTW(ITW, aInfos);
  {On peut déplacer les éléments de la source dans les deux sens}
  ITW^.CanRemove := True;
  ITW^.CanPoolRe := True;
  Neu := TVSource.Items.AddChildObject(nil, ITW^.Libelle, ITW);
  if ITW^.MoulineBQ then Neu.ImageIndex := imi_BQOK
                    else Neu.ImageIndex := imi_NotBQ;
  Neu.SelectedIndex := Neu.ImageIndex;
end;

{Charge la tob contenant toutes les informations des dossiers du regroupement et les 2 TreeView
{---------------------------------------------------------------------------------------}
procedure TOF_TRADDCASHPOOLING.ChargeInfos;
{---------------------------------------------------------------------------------------}
var
  n : Integer;
  F : TOB;
begin
  TobInfos := RecupInfosSocietes('SO_TRBASETRESO;SO_NODOSSIER;SO_LIBELLE;SO_TRMOULINEBQCODE;So_TRNOCASHPOOLING;');
  for n := 0 to TobInfos.Detail.Count - 1 do begin
    F := TobInfos.Detail[n];
    if F.GetString('SO_TRBASETRESO') = '' then RemplirSource(F)
                                          else RemplirDestination(F);
  end;

end;

{---------------------------------------------------------------------------------------}
function TOF_TRADDCASHPOOLING.MoulineDossier(NomBase, NoDossier, NomPool : string; MoulineBq : Boolean) : Boolean;
{---------------------------------------------------------------------------------------}

    {------------------------------------------------------}
    procedure Finitions;
    {------------------------------------------------------}
    begin
      {Mise à jour de la table dossier avec les infos des ParamSoc}
      MajTableDossier(NomBase);
      SetParamSocDossier('SO_TRBASETRESO', NomPool, NomBase);
      Result := True;
    end;

begin
  Result := False;
  try
    {On remplace les valeurs des champs _GENERAL avec BQ_CODE}
    if PartageBanqueCP(NomBase)
    {On remplace les valeurs des champs _GENERAL avec BQ_CODE}
    and MajBQCode(NomBase, 'BQ_REPIMPAYEPRELV')
    {On Initialise les champs _NODOSSIER}
    and MajNoDossier(NomBase + ';') then begin
      {Fusionne les différentes tables de conditions du dossiers avec celles du pool}
      if FusionneConditions(NomBase, NomPool) then begin
        {Si on n'est sur le pool}
        if (NomBase = NomPool) then Finitions
        {Fusionne les tables de mouvements du dossier avec celles du pool}
        else if FusionneTableMvt(NomBase, NomPool) then Finitions;
      end;
    end;
  except
    on E : Exception do begin
      PgiError(TraduireMemoire('Une erreur est intervenu lors du traitement du dossier "') + NomBase +'",'#13 +
               TraduireMemoire('avec le message : ') + E.Message + #13#13 +
               TraduireMemoire('Ce dossier ne sera pas ajouté au Pool "') + NomPool + '".', Ecran.Caption);
    end;
  end;
end;

Function CacheCash : Boolean ;
// GP le 23/06/2008 : On cache le cash
BEGIN
Result:=TRUE ;
If Not estMultiSoc Then Exit ;
If Trim(GetParamSocSecur('SO_TRBASETRESO',''))='' Then Exit ;
Result:=FALSE ;
END ;

initialization
  RegisterClasses([TOF_TRADDCASHPOOLING]);

end.

