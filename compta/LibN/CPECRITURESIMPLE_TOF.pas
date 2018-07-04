{-------------------------------------------------------------------------------------
  Version    |  Date  | Qui |   Commentaires
--------------------------------------------------------------------------------------
08.01.001.009  14/02/07  JP  Création d'une saisie simplifiée plus générique que UTofEcrLet
08.00.001.025  17/07/07  JP  FQ 21129 : Correction de la gestion de la TVA
08.10.001.010  20/09/07  JP  FQ 21367 : si on passe d'un collectif à autre chose, il faut vider l'auxiliaire
                             FQ 21312 : par ricochet
08.10.001.013  10/10/07  JP  FQ 21055 : Meilleur gestion des bordereaux / Libre
--------------------------------------------------------------------------------------}
unit CPECRITURESIMPLE_TOF;

interface

uses
  {$IFDEF EAGLCLIENT}
  MaineAGL,
  {$ELSE}
  FE_Main,
  {$ENDIF}
  StdCtrls, Controls, Classes, SysUtils, UTOF, UTob, ULibEcrSimplifiee, ULibPieceCompta,
  ULibSaisiePiece, ULibEcriture;

type
  TObjEcritureSimplif = class;

  {JP 14/02/07 : création d'une classe ancêtre commune aux deux TOF, car les deux fiches
                 doivent rester IDENTIQUES}
  TOF_CPECRITURESIMPLE = class(TofAncetreEcr)
    procedure OnArgument(S : string); override;
    procedure OnClose               ; override;
    procedure OnUpdate              ; override;
    procedure OnLoad                ; override;
  private
    TSP : TSaisiePiece;

    function  PreparePiece : Boolean;
    procedure AffichePiece;
    procedure AfterChangeJournal;
    procedure AfterChangeGeneral;
    procedure MajEnabledTva;
  protected
    procedure PageControlChange        (Sender : TObject);
    procedure EdtGeneralOnExit         (Sender : TObject);
    procedure EdtAuxiliaireOnExit      (Sender : TObject);
    procedure EdtJournalOnExit         (Sender : TObject);
    procedure EdtJournalElipsisClick   (Sender : TObject); override;
    procedure EdtDateComptableExit     (Sender : TObject); override;
    procedure BtnVoirClick             (Sender : TObject); override;
    procedure EdtFolioExit             (Sender : TObject); override;
    procedure FormKeyDown              (Sender : TObject; var Key : Word; Shift : TShiftState); override;
    function  AssignEvent              : Boolean; override;
  public
    ObjEcriture : TObjEcritureSimplif;
    TobEntete   : Tob;
  end;

  TObjEcritureSimplif = class
    constructor Create(aEcranCaption : string = '');
    destructor  Destroy; override;
  private
    TobStructure  : TOB;
    FInfEcriture  : TInfoEcriture;
    FEcranCaption : string;
    FRegimeTVA    : string; {19/07/07 : FQ 21129}

    {Getter et Setter des properties}
    function  GetValeur(Indice : Integer) : Variant;
    procedure SetValeur(Indice : Integer; Value : Variant);
    function  GetInfo  : TInfoEcriture;
    function  GetRegimeTVA : string; {19/07/07 : FQ 21129}
    function  IsBordereau : Boolean;
    procedure GereErreur(Sender : TObject; Error : TRecError);
  public
    PieceCompta : TPieceCompta;
    {En cas de Multi sociétés}
    Dossier     : string;
    {Si on ne veut pas que l'utilisateur modifie les champs paramétrés, cad les properties de l'objet
     A False par défaut. A mettre à True avant de faire LanceSaisieSimplifie}
    Securise    : Boolean;
    {Si l'on veut saisir la pièce dans la grille : dans ce cas, il suffit juste d'initialiser
     l'entête de la pièce sur la première page. A False par défaut.
     A mettre à True avant de faire LanceSaisieSimplifie}
    ModeGrille  : Boolean;

    {Appel de la fiche de saisie simplifiée : Retourne True si une pièce a été préparée}
    function LanceSaisieSimplifie  : Boolean;
    {Crée la ligne 1 de la pièce à partir des valeurs des properties}
    procedure CreerLigne1(Ind : Integer);
    {Pour toute les lignes initialisation des champs standards de la pièce et gestion éventuelle de la TVA}
    procedure RemplitChpStd(Ind : Integer);
    {Mise à jour des champs liés a E_ECHE}
    procedure GereEche(Ind : Integer);
    {Gestion de la TVA de la pièce}
    procedure InsereTVA(Compte : string);
    {Retourne le compte de TVA : à appeler lorsque l'on est sur le compte HT}
    function  RecupCpteTva(GereTva : Boolean; Ind : Integer) : string;
    {19/07/07 : FQ 21129 : Retourne le taux de TVA}
    function  RecupTauxTva : string;

    {Création du TPieceCompta}
    function  CreatePieceCompta : TPieceCompta;

    {Pour créer une nouvelle property à l'objet, il faut
     1/ Créer la property dans la definition de la classe TObjEcritureSimplif
     2/ Créer une nouvelle constante (CHP_XXX) dans l'interface de ULibPointage
     3/ Ajouter le champ en respectant bien l'ordre des champs supplémentaires de la TOB}
    property Journal  : variant index CHP_JOURNAL  read GetValeur write SetValeur;
    property DateCpt  : variant index CHP_DATECPT  read GetValeur write SetValeur;
    property General  : variant index CHP_GENERAL  read GetValeur write SetValeur;
    property Tiers    : variant index CHP_TIERS    read GetValeur write SetValeur;
    property NatPiece : variant index CHP_NATPIECE read GetValeur write SetValeur;
    property Devise   : variant index CHP_DEVISE   read GetValeur write SetValeur;
    property Etab     : variant index CHP_ETAB     read GetValeur write SetValeur;
    property QualifP  : variant index CHP_QUALIFP  read GetValeur write SetValeur;
    property Libelle  : variant index CHP_LIBELLE  read GetValeur write SetValeur;
    property RefInt   : variant index CHP_REFINT   read GetValeur write SetValeur;
    property CreditD  : variant index CHP_CREDITD  read GetValeur write SetValeur;
    property DebitD   : variant index CHP_DEBITD   read GetValeur write SetValeur;
    property DateVal  : variant index CHP_DATEVAL  read GetValeur write SetValeur;
    property ModePaie : variant index CHP_MODEPAIE read GetValeur write SetValeur;
    property RIB      : variant index CHP_RIB      read GetValeur write SetValeur;

    property InfEcriture   : TInfoEcriture         read GetInfo   write FInfEcriture;
    property RegimeTVA     : string                read GetRegimeTVA write FRegimeTVA; {19/07/07 : FQ 21129}
    property EcranCaption  : string                read FEcranCaption;
    property BordereauOk   : Boolean               read IsBordereau;
  end;


implementation


uses
  {$IFNDEF EAGLCLIENT}{$IFNDEF DBXPRESS}dbtables{BDE},{$ELSE}uDbxDataSet,{$ENDIF} {$ENDIF EAGLCLIENT}
  {$IFDEf VER150} Variants, {$ENDIF}
  {$IFDEF MODENT1}
  CPProcMetier,
  {$ENDIF MODENT1}
  Forms, ComCtrls, HCtrls, HEnt1, HMsgBox, Vierge, SaisUtil, Ent1, Windows,
  AGLInit, uTobDebug, LookUp, ParamSoc, UProcGen, Messages;

     {-----------------------------------------------------------------------------}
     {-----------------------------------------------------------------------------}
     {                        TOF_CPECRITURESIMPLE                                 }
     {-----------------------------------------------------------------------------}
     {-----------------------------------------------------------------------------}

{---------------------------------------------------------------------------------------}
procedure TOF_CPECRITURESIMPLE.OnArgument(S : string);
{---------------------------------------------------------------------------------------}
begin
  inherited;
  PgcDetail.ActivePageIndex := 0;
  TobEntete   :=  TheTob;
  ObjEcriture := TObjEcritureSimplif(TheData);
  P2.Visible  := False;
  PgcDetail.Pages[1].TabVisible := False;
end;

{Sur le FormShow
{---------------------------------------------------------------------------------------}
procedure TOF_CPECRITURESIMPLE.OnLoad;
{---------------------------------------------------------------------------------------}
var
  DatC : TDateTime;
begin
  inherited;
  {Initialisation de la fiche}
  with ObjEcriture do begin
    if IsValidDate(VarToStr(DateCpt)) then DatC := VarToDateTime(DateCpt)
                                      else DatC := iDate1900;
                                      
    {Si on ne veut pas que l'utilisateur modifie les champs paramétrés}
    if Securise then begin
      SetControlEnabled('E_DATECOMPTABLE' , DatC <= iDate1900);
      SetControlEnabled('E_ETABLISSEMENT' , Etab = '');
      SetControlEnabled('E_JOURNAL'       , Journal = '');
      SetControlEnabled('E_NATUREPIECE'   , NatPiece = '');
      SetControlEnabled('TE_DATECOMPTABLE', DatC <= iDate1900);
      SetControlEnabled('TE_ETABLISSEMENT', Etab = '');
      SetControlEnabled('TE_JOURNAL'      , Journal = '');
      SetControlEnabled('TE_NATUREPIECE'  , NatPiece = '');
    end;

    {En mode saisie grille, on ne renseigne que l'entête}
    if ModeGrille then begin
      SetControlEnabled('E_GENERAL'     , False);
      SetControlEnabled('TE_GENERAL'    , False);
      SetControlEnabled('E_LIBELLE'     , False);
      SetControlEnabled('TE_LIBELLE'    , False);
    end;

    {Ces zones seront activées en fonction de la saisie}
    SetControlEnabled('E_TVA'           , False);
    SetControlEnabled('E_AUXILIAIRE'    , False);
    SetControlEnabled('TE_TVA'          , False);
    SetControlEnabled('TE_AUXILIAIRE'   , False);

    if DatC > iDate1900 then SetControlText('E_DATECOMPTABLE', DateToStr(DatC))
                        else SetControlText('E_DATECOMPTABLE', DateToStr(Date));

    SetControlText('E_JOURNAL'      , Journal);
    AfterChangeJournal;
    SetControlText('E_LIBELLE'      , Libelle);
    SetControlText('E_ETABLISSEMENT', Etab);
    SetControlText('E_NATUREPIECE'  , NatPiece);
    SetControlText('E_LIBELLE'      , Libelle);
    SetControlCaption('LGEN', '');
    SetControlCaption('LAUX', '');
    {On force l'uilisateur à changer de page pour voir le résultat de la pièce avant de valider}
    SetControlEnabled('BVALIDER', not ModeGrille);
    MajEnabledTva;
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPECRITURESIMPLE.OnClose;
{---------------------------------------------------------------------------------------}
begin
  if Assigned(TSP) then FreeAndNil(TSP);
  inherited;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPECRITURESIMPLE.OnUpdate;
{---------------------------------------------------------------------------------------}
begin
  inherited;
  if not Assigned(ObjEcriture.PieceCompta) or (PgcDetail.ActivePageIndex = 0) then
    PreparePiece;
    
  if ObjEcriture.PieceCompta.IsValidPiece then begin
    TFVierge(Ecran).Retour := 'OK';
    Ecran.Close;
  end;
end;

{---------------------------------------------------------------------------------------}
function TOF_CPECRITURESIMPLE.AssignEvent: Boolean;
{---------------------------------------------------------------------------------------}
begin
  Result := inherited AssignEvent;
  if not Result then Exit;
  EdtAuxiliaire.OnExit := EdtAuxiliaireOnExit;
  EdtGeneral   .OnExit := EdtGeneralOnExit;
  EdtJournal   .OnExit := EdtJournalOnExit;
  PgcDetail  .OnChange := PageControlChange;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPECRITURESIMPLE.PageControlChange(Sender: TObject);
{---------------------------------------------------------------------------------------}
begin
  BtnVoir.Enabled := PgcDetail.TabIndex <> 1;
  if ObjEcriture.ModeGrille then SetControlEnabled('BVALIDER', PgcDetail.TabIndex <> 0);
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPECRITURESIMPLE.BtnVoirClick(Sender: TObject);
{---------------------------------------------------------------------------------------}
begin
  inherited;
  if PreparePiece and (ObjEcriture.ModeGrille or ObjEcriture.PieceCompta.IsValidPiece) then
    AffichePiece;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPECRITURESIMPLE.EdtDateComptableExit(Sender: TObject);
{---------------------------------------------------------------------------------------}
begin
  inherited;
 {Gestion de la combo folio}
 RemplitComboFolio;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPECRITURESIMPLE.EdtFolioExit(Sender: TObject);
{---------------------------------------------------------------------------------------}
var
  Q : TQuery;
begin
  inherited;
  if ObjEcriture.BordereauOk then begin
    if (Trim(EdtFolio.Text) = '' ) then begin
      PGIError(TraduireMemoire('Le numéro de folio est incorrect'), Ecran.Caption);
      if EdtFolio.CanFocus then EdtFolio.SetFocus;
    end;
  end

  else if ObjEcriture.InfEcriture.Journal.GetValue('J_MODESAISIE') = 'LIB' then begin
    Q := OpenSql(' SELECT E_DEVISE FROM ECRITURE ' +
                 ' WHERE E_JOURNAL = "' + EdtJournal.Text + '"' +
                 ' AND E_EXERCICE = "' + QuelExo(EdtDateComptable.Text) + '" ' +
                 ' AND E_DATECOMPTABLE >= "'+USDateTime(StrToDate( EdtDateComptable.Text )) + '" ' +
                 ' AND E_DATECOMPTABLE <= "'+USDateTime(StrToDate( EdtDateComptable.Text )) + '" ' +
                 ' AND E_NUMEROPIECE =' + EdtFolio.Text +
                 ' AND E_QUALIFPIECE = "N" AND E_NUMLIGNE = 1', True);
    if not Q.Eof and (Q.FindField('E_DEVISE').asString <> V_PGI.DevisePivot) then begin
      PGIError(TraduireMemoire('Vous ne pouvez pas faire d''écriture compte à compte sur un journal libre en devise'), Ecran.Caption);
      if EdtFolio.CanFocus then EdtFolio.SetFocus;
    end;
  end;

  {10/10/07 : FQ 21055 : on s'assure que le folio n'est pas validé}
  if ExisteSQL('SELECT E_VALIDE FROM ECRITURE WHERE E_JOURNAL = "' + EdtJournal.Text + '"' +
               ' AND E_EXERCICE = "' + QuelExo(EdtDateComptable.Text) + '" ' +
               ' AND E_DATECOMPTABLE >= "' + USDateTime(DebutDeMois(StrToDate(EdtDateComptable.Text))) + '" ' +
               ' AND E_DATECOMPTABLE <= "' + USDateTime(FinDeMois(StrToDate(EdtDateComptable.Text))) + '" ' +
               ' AND E_NUMEROPIECE =' + EdtFolio.Text +
               ' AND E_QUALIFPIECE = "N" AND E_VALIDE = "X"') then begin
    PGIError(TraduireMemoire('Ce folio est validé. Veuillez en choisir un autre.'), Ecran.Caption);
    if EdtFolio.CanFocus then EdtFolio.SetFocus;
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPECRITURESIMPLE.FormKeyDown(Sender : TObject; var Key : Word; Shift : TShiftState);
{---------------------------------------------------------------------------------------}
begin
  case Key of
    VK_F6  : if PgcDetail.ActivePageIndex = 0 then inherited;
  else
    inherited;
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPECRITURESIMPLE.EdtAuxiliaireOnExit(Sender: TObject);
{---------------------------------------------------------------------------------------}
begin
  if not ObjEcriture.InfEcriture.LoadAux(EdtAuxiliaire.Text) then
    SendMessage((Sender as TCustomEdit).Handle, WM_KEYDOWN, VK_F5, 0);
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPECRITURESIMPLE.EdtGeneralOnExit(Sender: TObject);
{---------------------------------------------------------------------------------------}
var
  Cpte : string;
begin
  Cpte := EdtGeneral.Text;
  if ObjEcriture.InfEcriture.Compte.GetCompte(Cpte) = - 1 then
    SendMessage((Sender as TCustomEdit).Handle, WM_KEYDOWN, VK_F5, 0)
  {Le GetCompte complétant la valeur passée en paramètre, on peut donc avoir :
   Cpte = 4010000 et EdtGeneral.Text = 401}
  else if Cpte <> EdtGeneral.Text then
    EdtGeneral.Text := Cpte;

  AfterChangeGeneral;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPECRITURESIMPLE.EdtJournalOnExit(Sender: TObject);
{---------------------------------------------------------------------------------------}
begin
  AfterChangeJournal;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPECRITURESIMPLE.EdtJournalElipsisClick(Sender: TObject);
{---------------------------------------------------------------------------------------}
var
 lStWhere : string;
begin
  inherited;
  lStWhere := 'J_FERME="-" AND J_NATUREJAL <> "ANO" AND J_NATUREJAL <> "ANA" AND ' +
              'J_NATUREJAL <> "CLO" AND J_NATUREJAL <> "ODA"';
 if LookupList(EdtJournal, TraduireMemoire('Journaux'), 'JOURNAL', 'J_JOURNAL', 'J_LIBELLE', lStWhere, 'J_JOURNAL', True, 4) then
   AfterChangeJournal;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPECRITURESIMPLE.AfterChangeJournal;
{---------------------------------------------------------------------------------------}
begin
  ObjECriture.FInfEcriture.LoadJournal(EdtJournal.Text);
  ObjECriture.Journal := EdtJournal.Text;
  EdtFolio.Visible    := ObjEcriture.BordereauOk;
  lblFolio.Visible    := EdtFolio.Visible;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPECRITURESIMPLE.AfterChangeGeneral;
{---------------------------------------------------------------------------------------}
begin
  ObjECriture.InfEcriture.LoadCompte(EdtGeneral.Text);
  MajEnabledTva;
  EdtAuxiliaire.Enabled := ObjECriture.FInfEcriture.Compte.IsCollectif;
  {JP 20/09/07 : FQ 21367 : si on passe d'un collectif à autre chose, il faut vider l'auxiliaire
                 FQ 21312 : par ricochet}
  if not EdtAuxiliaire.Enabled then EdtAuxiliaire.Text := '';
  lblAuxiliaire.Enabled := EdtAuxiliaire.Enabled;
  FStNatGene            := ObjEcriture.InfEcriture.GetString('G_NATUREGENE');
  if EdtAuxiliaire.Enabled and (Ecran.ActiveControl = EdtLibelle) then
    EdtAuxiliaire.SetFocus;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPECRITURESIMPLE.MajEnabledTva;
{---------------------------------------------------------------------------------------}
var
  EnabledOk : Boolean;
begin
  EnabledOk := (Trim(EdtGeneral.Text) <> '') and (ObjEcriture.RegimeTVA <> '');
  EnabledOk := EnabledOk and ObjEcriture.InfEcriture.Compte.IsTvaAutorise(edtGeneral.Text);

  {Si la contre partie est un collectif (rappel : en pointage ce sera un 512), on s'assure
   que le tiers n'est pas un Salarie}
  if EnabledOk then begin
    ObjEcriture.InfEcriture.LoadCompte(ObjEcriture.General);
    if ObjEcriture.InfEcriture.Compte.IsCollectif then begin
      ObjEcriture.InfEcriture.LoadAux(ObjEcriture.Tiers);
      ObjEcriture.RegimeTVA := ObjEcriture.InfEcriture.GetString('T_REGIMETVA');
      EnabledOk := ObjEcriture.InfEcriture.GetString('T_NATUREAUXI') <> 'SAL';
    end;
  end;

  {19/07/07 : FQ 21129 : Chargement du taux de TVA}
  if EnabledOk then begin
    ObjEcriture.InfEcriture.LoadCompte(edtGeneral.Text);
    EdtTva.Text := ObjEcriture.RecupTauxTva;
  end;

  EdtTva.Enabled := EnabledOk;
  lblTva.Enabled := EnabledOk;
end;

{---------------------------------------------------------------------------------------}
function TOF_CPECRITURESIMPLE.PreparePiece : Boolean;
{---------------------------------------------------------------------------------------}
var
  AvecTva : Boolean;
  CpteTva : string;
  OldInd  : Integer;
  TmpInd  : Integer;
begin
  Result := True;
  {Mise à jour des champs de l'objet}
  with ObjECriture do begin
    Libelle  := EdtLibelle.Text;
    DateCpt  := StrToDate(EdtDateComptable.Text);
    Etab     := EdtEtablissement.Text;
    Journal  := EdtJournal.Text;
    NatPiece := EdtNaturePiece.Text;
  end;

  {Eventuelle création de l'objet PieceCompta}
  if not Assigned(ObjECriture.PieceCompta) then
    ObjECriture.PieceCompta := ObjECriture.CreatePieceCompta;

  with ObjECriture, PieceCompta do begin
    {1/ Initialisation de la Pièce}
    SetMultiEcheOff;
    InitPiece(Journal, VarToDateTime(DateCpt), NatPiece, ObjECriture.Devise, Etab, QualifP);
    InitSaisie;
    {Si on est en saisie libre ou borderau, on recherche si le folio existe pour saisir à la suite}
    if ModeSaisie <> msPiece then begin
      PutEntete('E_NUMEROPIECE', EdtFolio.Text);
      LoadFromSQL;
      {10/10/07 : ajout du lock}
      if not LockFolio(True) then Exit;
    end;

    if ModeGrille then Exit;

    {2/ Création de la ligne avec les infos de l'objet}
    NewRecord;
    TmpInd := CurIdx;
    CreerLigne1(TmpInd);
    PutValue(TmpInd, 'E_CONTREPARTIEGEN', EdtGeneral.Text);
    PutValue(TmpInd, 'E_CONTREPARTIEAUX', EdtAuxiliaire.Text);
    OldInd := TmpInd;

    {3/ Création avec la contrepartie à partir de la Saisie}
    NewRecord;
    TmpInd := CurIdx;
    PutValue(TmpInd, 'E_GENERAL'   , EdtGeneral.Text);
    PutValue(TmpInd, 'E_AUXILIAIRE', EdtAuxiliaire.Text);
    PutValue(TmpInd, 'E_CONTREPARTIEGEN', General);
    PutValue(TmpInd, 'E_CONTREPARTIEAUX', Tiers);
    {Gestion des échéances, TVA et des champs standards}
    RemplitChpStd(TmpInd);

    {Sur les 512, la TVA n'est pas mise à jour, on le fait manuellement}
    if GetValue(TmpInd, 'E_TVA') <> '' then PutValue(OldInd, 'E_TVA', GetValue(TmpInd, 'E_TVA'));

    InfEcriture.LoadCompte(EdtGeneral.Text); {19/07/07 : FQ 21129}
    CpteTva := RecupCpteTva(edtTva.Enabled and (Valeur(edtTva.text) <> 0), TmpInd);
    if (CpteTva <> '') then begin
      PutValue(TmpInd, 'E_DEBITDEV' , GetMntHTFromTTC(CreditD, Valeur(edtTva.Text) / 100, Devise.Decimale));
      PutValue(TmpInd, 'E_CREDITDEV', GetMntHTFromTTC(DebitD, Valeur(edtTva.Text) / 100, Devise.Decimale));
      AvecTva := True;
    end
    else begin
      PutValue(TmpInd, 'E_DEBITDEV' , CreditD);
      PutValue(TmpInd, 'E_CREDITDEV', DebitD);
      AvecTva := False;
    end;

    {4/ Eventuelle gestion de la TVA}
    if AvecTva then InsereTva(CpteTva);
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TOF_CPECRITURESIMPLE.AffichePiece;
{---------------------------------------------------------------------------------------}
begin
  //TobDebug(ObjECriture.PieceCompta);
  {Creation de l'objet de gestion de saisie en grille}
  if Assigned(TSP) then FreeAndNil(TSP);
  {Affichage de la Grille}
  TSP := TSaisiePiece.Create(Ecran, G, ObjECriture.PieceCompta);
  TSP.UpdateMasqueSaisie;
  TSP.AfficheLignes;

  PgcDetail.Pages[1].TabVisible := True;
  PgcDetail.TabIndex := 1;
  PgcDetail.OnChange(PgcDetail);
  TSP.SetFocus;
end;


     {-----------------------------------------------------------------------------}
     {-----------------------------------------------------------------------------}
     {                        TObjEcritureSimplif                                  }
     {-----------------------------------------------------------------------------}
     {-----------------------------------------------------------------------------}

{---------------------------------------------------------------------------------------}
constructor TObjEcritureSimplif.Create(aEcranCaption : string = '');
{---------------------------------------------------------------------------------------}
begin
  TobStructure := TOB.Create('****', nil, -1);
  InitStructureSimplifie(TobStructure);
  Dossier       := V_PGI.SchemaName;
  Securise      := False;
  ModeGrille    := False;
  FInfEcriture  := nil;
  PieceCompta   := nil;
  FEcranCaption := EcranCaption;
  if FEcranCaption = '' then FEcranCaption := TraduireMemoire('Saisie simplifiée');
end;

{---------------------------------------------------------------------------------------}
destructor TObjEcritureSimplif.Destroy;
{---------------------------------------------------------------------------------------}
begin
  {10/10/07 : ajout du lock}
  PieceCompta.UnLockFolio;
  if Assigned(TobStructure) then FreeAndNil(TobStructure);
  if Assigned(PieceCompta) then FreeAndNil(PieceCompta);
  if Assigned(FInfEcriture) then FreeAndNil(FInfEcriture);
  inherited;
end;

{---------------------------------------------------------------------------------------}
function TObjEcritureSimplif.GetValeur(Indice : Integer) : Variant;
{---------------------------------------------------------------------------------------}
begin
  if not Assigned(TobStructure) then Exit;
  Result := TobStructure.GetValeur(Indice);
end;

{---------------------------------------------------------------------------------------}
procedure TObjEcritureSimplif.SetValeur(Indice : Integer; Value : Variant);
{---------------------------------------------------------------------------------------}
var
  Val : string;
begin
  if not Assigned(TobStructure) then Exit;
  TobStructure.PutValeur(Indice, Value);
  {Mise à jour du TInfoEcriture}
  Val := VarToStr(Value);
  if Val <> '' then begin
         if Indice = CHP_JOURNAL then InfEcriture.LoadJournal(Val)
    else if Indice = CHP_GENERAL then InfEcriture.LoadCompte(Val)
    else if Indice = CHP_TIERS   then InfEcriture.LoadAux(Val);
  end;
end;

{---------------------------------------------------------------------------------------}
function TObjEcritureSimplif.GetInfo : TInfoEcriture;
{---------------------------------------------------------------------------------------}
begin
  if not Assigned (FInfEcriture) then
    FInfEcriture := TInfoEcriture.Create(Dossier);
  Result := FInfEcriture;
end;

{---------------------------------------------------------------------------------------}
function TObjEcritureSimplif.GetRegimeTVA : string;
{---------------------------------------------------------------------------------------}
begin
  if FRegimeTVA = '' then FRegimeTVA := GetParamSocSecur('SO_REGIMEDEFAUT', '');
  Result := FRegimeTVA;
end;

{---------------------------------------------------------------------------------------}
function TObjEcritureSimplif.LanceSaisieSimplifie : Boolean;
{---------------------------------------------------------------------------------------}
begin
  if V_PGI.SAV then TobDebug(TobStructure);
  TheTob  := TobStructure;
  TheData := TObjEcritureSimplif(Self);
  Result := (AglLanceFiche('CP', 'CPECRITRURESIMPLE', '', '', '') = 'OK') and (PieceCompta.Detail.Count > 0);
end;

{---------------------------------------------------------------------------------------}
procedure TObjEcritureSimplif.CreerLigne1(Ind : Integer);
{---------------------------------------------------------------------------------------}
begin
  with PieceCompta do begin
    PutValue(Ind, 'E_GENERAL'   , General);
    PutValue(Ind, 'E_AUXILIAIRE', Tiers);
    PutValue(Ind, 'E_CREDITDEV' , CreditD);
    PutValue(Ind, 'E_DEBITDEV'  , DebitD);
    {Maj des champs communs aux trois lignes de la pièce}
    RemplitChpStd(Ind);
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TObjEcritureSimplif.GereEche(Ind : Integer);
{---------------------------------------------------------------------------------------}
begin
  with PieceCompta do begin
    if GetValue(Ind, 'E_ECHE') = 'X' then begin
      {Si le mode de paiement n'est pas vide, on force celui renseigné, sinon ce sera le mode
       de paiement par défaut fourni par le noyau}
      if ModePaie <> '' then
        PutValue(Ind, 'E_MODEPAIE'       , ModePaie);
      {A priori, dans ce cas on est sur une écriture à partir d'un compte bancaire}
      if DateVal > iDate1900 then begin
        PutValue(Ind, 'E_DATEVALEUR'     , DateVal);
        PutValue(Ind, 'E_DATEECHEANCE'   , DateVal);
        PutValue(Ind, 'E_ORIGINEPAIEMENT', DateVal);
      end;
    end;
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TObjEcritureSimplif.RemplitChpStd(Ind : Integer);
{---------------------------------------------------------------------------------------}
begin
  {Gestion des échéances}
  GereEche(Ind);
  {Champs génériques}
  with PieceCompta do begin
    {En Bordereau et libre, la date comptable et la nature de pièce sont au niveau des lignes}
    if BordereauOk then begin
      PutValue(Ind, 'E_DATECOMPTABLE', DateCpt);
      PutValue(Ind, 'E_NATUREPIECE'  , NatPiece);
    end;
    PutValue(Ind, 'E_REGIMETVA', RegimeTVA);
    {Affectation de E_TVA et E_REGIMETVA}
    AffecteTVA(Ind);
    {Champs standards}
    PutValue(Ind, 'E_QUALIFORIGINE', QUALIFSIMPLIFIE);
    PutValue(Ind, 'E_VALIDE'       , '-');
    PutValue(Ind, 'E_IO'           , 'X');
    PutValue(Ind, 'E_CREERPAR'     , V_PGI.User);
    PutValue(Ind, 'E_LIBELLE'      , Libelle);
    PutValue(Ind, 'E_REFINTERNE'   , RefInt);
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TObjEcritureSimplif.InsereTVA(Compte : string);
{---------------------------------------------------------------------------------------}
var
  Ind : Integer;
begin
  with PieceCompta do begin
    NewRecord;
    Ind := CurIdx;
    PutValue(Ind, 'E_GENERAL'   , Compte);
    PutValue(Ind, 'E_CONTREPARTIEGEN', General);
    PutValue(Ind, 'E_CONTREPARTIEAUX', Tiers);
    {Solde la pièce}
    AttribSolde(Ind);
    {Gestion des échéances, TVA et des champs standards}
    RemplitChpStd(Ind);
  end;
end;

{Retourne le compte de TVA : à appeler lorsque l'on est sur le compte HT
{---------------------------------------------------------------------------------------}
function TObjEcritureSimplif.RecupCpteTva(GereTva : Boolean; Ind : Integer) : string;
{---------------------------------------------------------------------------------------}
var
  IsAchat : Boolean;
begin
  Result := '';
  if Assigned(PieceCompta) and Assigned(InfEcriture) and GereTva then
    with PieceCompta do begin
      IsAchat := InfEcriture.Compte.GetString('G_NATUREGENE') = 'CHA';

      if GetValue(CurIdx, 'E_TVAENCAISSEMENT') = 'X' then
        Result := Tva2Encais(GetValue(Ind, 'E_REGIMETVA'), GetValue(Ind, 'E_TVA'), IsAchat)
      else
        Result := Tva2Cpte(GetValue(Ind, 'E_REGIMETVA'), GetValue(Ind, 'E_TVA'), IsAchat);
    end;
end;

{19/07/07 : FQ 21129 : Retourne le taux de TVA
{---------------------------------------------------------------------------------------}
function TObjEcritureSimplif.RecupTauxTva : string;
{---------------------------------------------------------------------------------------}
var
  IsAchat : Boolean;
  E_TVA   : string;
begin
  Result := '';
  if Assigned(InfEcriture) then begin
    IsAchat := InfEcriture.Compte.GetString('G_NATUREGENE') = 'CHA';
    if InfEcriture.Compte.GetString('G_TVA') = '' then E_TVA := GetParamsocSecur('SO_CODETVAGENEDEFAULT',   '')
                                                  else E_TVA := InfEcriture.Compte.GetString('G_TVA');
    Result := FloatToStr(Tva2Taux(RegimeTVA, E_TVA, IsAchat) * 100);
  end;
end;

{---------------------------------------------------------------------------------------}
function TObjEcritureSimplif.IsBordereau : Boolean;
{---------------------------------------------------------------------------------------}
begin
  Result := (FInfEcriture.Journal.GetValue('J_MODESAISIE') = 'LIB') or
            (FInfEcriture.Journal.GetValue('J_MODESAISIE') = 'BOR');
end;

{---------------------------------------------------------------------------------------}
function TObjEcritureSimplif.CreatePieceCompta : TPieceCompta;
{---------------------------------------------------------------------------------------}
begin
  Result := TPieceCompta.CreerPiece(InfEcriture);
  Result.OnError := GereErreur;
end;

{---------------------------------------------------------------------------------------}
procedure TObjEcritureSimplif.GereErreur(Sender : TObject; Error : TRecError);
{---------------------------------------------------------------------------------------}
var
  MsgStruct : TMessageCompta;
begin
  if Trim(Error.RC_Message) <> '' then
    PGIInfo(Error.RC_Message, FEcranCaption)
  else if (Error.RC_Error <> RC_PASERREUR) then begin
    if not BordereauOk then MsgStruct := TMessageCompta.Create(FEcranCaption, msgSaisiePiece)
                       else MsgStruct := TMessageCompta.Create(FEcranCaption);
    try
      MsgStruct.Execute(Error.RC_Error);
    finally
      FreeAndNil(MsgStruct);
    end;
  end;

end;

initialization
  RegisterClasses([TOF_CPECRITURESIMPLE]);


end.
