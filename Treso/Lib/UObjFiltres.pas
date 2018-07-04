unit UObjFiltres;
{   FONCTIONNEMENT DE L'OBJET FILTRE POUR LES FICHES VIERGES

  1/ Déclarer une une vairable  de classe : ObjFiltre : TObjFiltre
  2/ de le OnArgument de la fiche
     21/ Déclarer une variable : Composants : TControlFiltre;

     22/ Composants.PopupF   := TPopUpMenu      (Getcontrol('POPF'));
         Composants.Filtres  := THValComboBox   (Getcontrol('FFILTRES'));
         Composants.Filtre   := TToolBarButton97(Getcontrol('BFILTRE'));
         Composants.PageCtrl := TPageControl    (Getcontrol('PCPAGECONTROL'));
         ObjFiltre := TObjFiltre.Create(Composants, 'TRSYNCHRONIS');
  3/ dans le OnLoad  : ObjFiltre.Charger
  4/ dans le OnClose :   if Assigned(ObjFiltre)    then FreeAndNil(ObjFiltre);

  AvantListByUser : Si l'on veut exécuter un traitement sur les évènements du filtre (Nouveau, supprimer ...)
                    ce booléen permet de préciser s'il faut effectuer le traitement de l'agl avant (False) ou
                    après (True) le code implémenté spécifiquement

  FFI_TABLE : Nom du filtre

  InChargement : permet de préciser si le filtre est en train d'être chargé => cela permet de désactiver
                 certaines parties de code des fiches

  ApresChangementFiltre : Procedure of Object exécutée à la fin du OnChangeFFiltres
}
interface

uses
  StdCtrls, Classes, HCtrls, HTB97, ComCtrls, UTob, SysUtils,
  Filtre, Menus, UListByUser, UTXml;

type
  TControlFiltre = record
    PopupF   : TPopUpMenu;
    Filtres  : THValComboBox;
    Filtre   : TToolBarButton97;
    PageCtrl : TPageControl;
  end;

  TObjFiltre = class (TObject)
  private
    {Gestion du filtre}
    PopF          : TPopUpMenu;
    FFiltres      : THValComboBox;
    BFiltre       : TToolBarButton97;
    PageControl   : TPageControl;
    FListeByUser  : TListByUser;
    FListeComboValuesNonAffecte : HString;

    procedure OnChangeFFiltres     (Sender : TObject);
    procedure InitAddFiltre        (T : TOB);
    procedure InitGetFiltre        (T : TOB);
    procedure InitSelectFiltre     (T : TOB);
    procedure UpgradeFiltre        (T : TOB);
    procedure ParseParamsFiltre    (Params : HTStrings);
    procedure UpdateComboFiltreAvances;

    {19/04/05 : Corrections de la gestions des pointeurs sur les évènements sur Popup}
    procedure LierClick    (Sender : TObject);
    procedure NouveauClick (Sender : TObject);
    procedure CreationClick(Sender : TObject);
    procedure RenommerClick(Sender : TObject);
    procedure ImporterClick(Sender : TObject);
    procedure ExporterClick(Sender : TObject);
    procedure SupprimeClick(Sender : TObject);
    procedure DupliqueClick(Sender : TObject);
    procedure EnregistClick(Sender : TObject);

    function  GetFiltreAcces : TFiltreAcces;
    procedure SetFiltreAcces ( Value : TFiltreAcces );

  public
    {Nom du filtre}
    FFI_TABLE     : string;
    InChargement  : Boolean;
    {19/04/05 : A True si le traitement "des pointeurs" ci-dessous doit être excécuter avant celui de l'AGL}
    AvantListByUser : Boolean;

    {19/04/05 : Nouveaux pointeurs sur les évènements sur Popup}
    LierFiltre     : procedure of object;
    NouveauFiltre  : procedure of object;
    CreationFiltre : procedure of object;
    RenommerFiltre : procedure of object;
    ImporterFiltre : procedure of object;
    ExporterFiltre : procedure of object;
    SupprimeFiltre : procedure of object;
    DupliqueFiltre : procedure of object;
    EnregistFiltre : procedure of object;

    {24/01/05 : Pointeur pour permettre un traitement après le changement de filtre : BChercheClick ...}
    AvantChangementFiltre : procedure of object;
    ApresChangementFiltre : procedure of object;

    constructor Create(Controles : TControlFiltre; NomFiche : string; FiltreDisabled : Boolean = False);
    destructor  Destroy; override;
    procedure   Charger;
    procedure   ForceSelectFiltre(NomFiltre : WideString; Appliquer : Boolean = True);

    property ForceAccessibilite : TFiltreAcces read GetFiltreAcces write SetFiltreAcces;
  end;


procedure CacheFiltreGraph(Ecran : TObject);


implementation

uses
  GRS1;

{JP 05/06/07 : Lorsque les graph sont alimentés par une Tob, la gestion des filtres doit être désactivée
{---------------------------------------------------------------------------------------}
procedure CacheFiltreGraph(Ecran : TObject);
{---------------------------------------------------------------------------------------}
begin
  TFGRS1(Ecran).BGraph  .Visible := False;
  TFGRS1(Ecran).BValider.Visible := False;
  TFGRS1(Ecran).FFiltres.Visible := False;
  TFGRS1(Ecran).BFiltre .Visible := False;
end;

{---------------------------------------------------------------------------------------}
constructor TObjFiltre.Create(Controles : TControlFiltre; NomFiche : string; FiltreDisabled : Boolean = False);
{---------------------------------------------------------------------------------------}
begin
  FListeComboValuesNonAffecte := '';
  InChargement := False;
  PageControl := Controles.PageCtrl;
  FFiltres    := Controles.Filtres;
  FFiltres.OnChange := OnChangeFFiltres;

  FFI_TABLE   := NomFiche;

  BFiltre      := Controles.Filtre;

  FListeByUser := TListByUser.Create(FFiltres, BFiltre, toFiltre, FiltreDisabled);
  with FListeByUser do begin
    OnSelect  := InitSelectFiltre;
    OnInitGet := InitGetFiltre;
    OnInitAdd := InitAddFiltre;
    OnUpgrade := UpgradeFiltre;
    OnParams  := ParseParamsFiltre;
    OnItemCreer       := CreationClick;
    OnItemEnregistrer := EnregistClick;
    OnItemRenommer    := RenommerClick;
    OnItemSupprimer   := SupprimeClick;
    OnItemDupliquer   := DupliqueClick;
    OnItemNouveau     := NouveauClick;
    OnItemImporter    := ImporterClick;
    OnItemExporter    := ExporterClick;
    OnItemLink        := LierClick;
  end;
  AvantListByUser := False;
end;

{---------------------------------------------------------------------------------------}
destructor TObjFiltre.Destroy;
{---------------------------------------------------------------------------------------}
begin
  if Assigned(FListeByUser) then FreeAndNil(FListeByUser);
  if Assigned(PopF) then PopF.Items.Clear;
  inherited Destroy;
end;

{---------------------------------------------------------------------------------------}
procedure TObjFiltre.Charger;
{---------------------------------------------------------------------------------------}
begin
  InChargement := True;
  if FListeByUser <> nil then
    FListeByUser.LoadDB(FFI_Table);
  InChargement := False;
end;

{---------------------------------------------------------------------------------------}
procedure TObjFiltre.OnChangeFFiltres(Sender: TObject);
{---------------------------------------------------------------------------------------}
begin
  if Assigned(AvantChangementFiltre) then
    AvantChangementFiltre;
    
  InChargement := True;
  FListeByUser.Select(FFiltres.Value);
  InChargement := False;
  {JP 24/01/05 : Pour pouvoir lancer un traitement après la mise à jour du filtre}
  if Assigned(ApresChangementFiltre) then
    ApresChangementFiltre;
end;

{---------------------------------------------------------------------------------------}
procedure TObjFiltre.InitAddFiltre(T: TOB);
{---------------------------------------------------------------------------------------}
var
  Lines : HTStrings ;
begin
  Lines := HTStringList.Create ;
  SauveCritMemoire(Lines, PageControl) ;
  FListeByUser.AffecteTOBFiltreMemoire(T, Lines);
  Lines.Free ;
end;

{---------------------------------------------------------------------------------------}
procedure TObjFiltre.InitGetFiltre(T: TOB);
{---------------------------------------------------------------------------------------}
var
  Lines : HTStrings ;
begin
  Lines := HTStringList.Create ;
  SauveCritMemoire(Lines, PageControl) ;
  FListeByUser.AffecteTOBFiltreMemoire(T, Lines);
  Lines.Free ;
end;

{---------------------------------------------------------------------------------------}
procedure TObjFiltre.InitSelectFiltre(T: TOB);
{---------------------------------------------------------------------------------------}
var
  Lines : HTStrings ;
  i : integer;
  stChamp, stVal : string;
begin
  if T = nil then Exit;
  Lines := HTStringList.Create ;
  for i:=0 to T.Detail.Count-1 do begin
    stChamp := T.Detail[i].GetValue('N');
    stVal   := T.Detail[i].GetValue('V');
    Lines.Add(stChamp + ';' + stVal);
  end;
  VideFiltre (FFiltres, PageControl, False);

  ChargeCritMemoire(Lines, PageControl, FListeComboValuesNonAffecte);
  UpdateComboFiltreAvances;
  Lines.Free ;
end;

{---------------------------------------------------------------------------------------}
procedure TObjFiltre.ParseParamsFiltre(Params: HTStrings);
{---------------------------------------------------------------------------------------}
var
  T : TOB;
begin
  FListeByUser.AddVersion;
  T := FListeByUser.Add;
  //en position 0 de Params se trouve le nom du filtre
  T.PutValue('NAME', XMLDecodeSt(Params[0])) ;
  T.PutValue('USER','---') ;
  Params.Delete(0);
  FListeByUser.AffecteTOBFiltreMemoire(T, Params);
end;

{---------------------------------------------------------------------------------------}
procedure TObjFiltre.UpgradeFiltre(T: TOB);
{---------------------------------------------------------------------------------------}
begin

end;

{---------------------------------------------------------------------------------------}
procedure TObjFiltre.UpdateComboFiltreAvances;
{---------------------------------------------------------------------------------------}
var
  St, stChamp, stVal: string;
  C : TComponent;
begin
  St := FListeComboValuesNonAffecte;
  while St <> '' do begin
    stVal := ReadTokenSt(St);
    stChamp := ReadTokenPipe(StVal,'|');
    C := PageControl.FindComponent(stChamp) ;
    if C = nil then Continue;
    if C is THValComboBox then begin
      if THValComboBox(C).Values.IndexOf(StVal)=-1 then begin
        THValComboBox(C).Values.Add(StVal);
        THValComboBox(C).Items.Add(StVal);
        THValComboBox(C).Value := stVal;
      end;
    end;
  end;
end;

{---------------------------------------------------------------------------------------}
procedure TObjFiltre.CreationClick(Sender : TObject);
{---------------------------------------------------------------------------------------}
begin
  {Traitement Agl si avant le traitement fonctionnel}
  if not AvantListByUser then FListeByUser.Creer;
  {Èventuel traitement fonctionnel}
  if Assigned(CreationFiltre) then CreationFiltre;
  {Traitement Agl si après le traitement fonctionnel}
  if AvantListByUser then FListeByUser.Creer;
  {Réinitialisation de AvantListByUser pour un prochain appel}
  AvantListByUser := False;
end;

{---------------------------------------------------------------------------------------}
procedure TObjFiltre.DupliqueClick(Sender : TObject);
{---------------------------------------------------------------------------------------}
begin
  {Traitement Agl si avant le traitement fonctionnel}
  if not AvantListByUser then FListeByUser.Duplicate;
  {Èventuel traitement fonctionnel}
  if Assigned(DupliqueFiltre) then DupliqueFiltre;
  {Traitement Agl si après le traitement fonctionnel}
  if AvantListByUser then FListeByUser.Duplicate;
  {Réinitialisation de AvantListByUser pour un prochain appel}
  AvantListByUser := False;
end;

{---------------------------------------------------------------------------------------}
procedure TObjFiltre.EnregistClick(Sender : TObject);
{---------------------------------------------------------------------------------------}
begin
  {Traitement Agl si avant le traitement fonctionnel}
  if not AvantListByUser then FListeByUser.Save;
  {Èventuel traitement fonctionnel}
  if Assigned(EnregistFiltre) then EnregistFiltre;
  {Traitement Agl si après le traitement fonctionnel}
  if AvantListByUser then FListeByUser.Save;
  {Réinitialisation de AvantListByUser pour un prochain appel}
  AvantListByUser := False;
end;

{---------------------------------------------------------------------------------------}
procedure TObjFiltre.ExporterClick(Sender : TObject);
{---------------------------------------------------------------------------------------}
begin
  {Traitement Agl si avant le traitement fonctionnel}
  if not AvantListByUser then FListeByUser.Exporter_;
  {Èventuel traitement fonctionnel}
  if Assigned(ExporterFiltre) then ExporterFiltre;
  {Traitement Agl si après le traitement fonctionnel}
  if AvantListByUser then FListeByUser.Exporter_;
  {Réinitialisation de AvantListByUser pour un prochain appel}
  AvantListByUser := False;
end;

{---------------------------------------------------------------------------------------}
procedure TObjFiltre.ImporterClick(Sender : TObject);
{---------------------------------------------------------------------------------------}
begin
  {Traitement Agl si avant le traitement fonctionnel}
  if not AvantListByUser then FListeByUser.Importer_;
  {Èventuel traitement fonctionnel}
  if Assigned(ImporterFiltre) then ImporterFiltre;
  {Traitement Agl si après le traitement fonctionnel}
  if AvantListByUser then FListeByUser.Importer_;
  {Réinitialisation de AvantListByUser pour un prochain appel}
  AvantListByUser := False;
end;

{---------------------------------------------------------------------------------------}
procedure TObjFiltre.LierClick(Sender : TObject);
{---------------------------------------------------------------------------------------}
begin
  {Èventuel traitement fonctionnel}
  if Assigned(LierFiltre) then LierFiltre;
  {Réinitialisation de AvantListByUser pour un prochain appel}
  AvantListByUser := False;
end;

{---------------------------------------------------------------------------------------}
procedure TObjFiltre.NouveauClick(Sender : TObject);
{---------------------------------------------------------------------------------------}
begin
  {Éventuel traitement fonctionnel si avant traitement Agl}
  if AvantListByUser and Assigned(NouveauFiltre) then NouveauFiltre;
  {Traitement Agl}
  VideFiltre(FFiltres, PageControl);
  FListeByUser.new;
  {Éventuel traitement fonctionnel si après traitement Agl}
  if not AvantListByUser and Assigned(NouveauFiltre) then NouveauFiltre;
  {Réinitialisation de AvantListByUser pour un prochain appel}
  AvantListByUser := False;
end;

{---------------------------------------------------------------------------------------}
procedure TObjFiltre.RenommerClick(Sender : TObject);
{---------------------------------------------------------------------------------------}
begin
  {Traitement Agl si avant le traitement fonctionnel}
  if not AvantListByUser then FListeByUser.Rename;
  {Èventuel traitement fonctionnel}
  if Assigned(RenommerFiltre) then RenommerFiltre;
  {Traitement Agl si après le traitement fonctionnel}
  if AvantListByUser then FListeByUser.Rename;
  {Réinitialisation de AvantListByUser pour un prochain appel}
  AvantListByUser := False;
end;

{---------------------------------------------------------------------------------------}
procedure TObjFiltre.SupprimeClick(Sender : TObject);
{---------------------------------------------------------------------------------------}
begin
  {Éventuel traitement fonctionnel si avant traitement Agl}
  if AvantListByUser and Assigned(SupprimeFiltre) then SupprimeFiltre;
  {Traitement Agl si avant le traitement fonctionnel}
  if FListeByUser.Delete then
    VideFiltre(FFiltres, PageControl);
  {Éventuel traitement fonctionnel si après traitement Agl}
  if not AvantListByUser and Assigned(SupprimeFiltre) then SupprimeFiltre;
  {Réinitialisation de AvantListByUser pour un prochain appel}
  AvantListByUser := False;
end;

{---------------------------------------------------------------------------------------}
function TObjFiltre.GetFiltreAcces: TFiltreAcces;
{---------------------------------------------------------------------------------------}
begin
  Result := FListeByUser.ForceAccessibilite;
end;

{---------------------------------------------------------------------------------------}
procedure TObjFiltre.SetFiltreAcces(Value: TFiltreAcces);
{---------------------------------------------------------------------------------------}
begin
  FListeByUSer.ForceAccessibilite := Value;
end;

{---------------------------------------------------------------------------------------}
procedure TObjFiltre.ForceSelectFiltre(NomFiltre : WideString; Appliquer : Boolean);
{---------------------------------------------------------------------------------------}
begin
  FListeByUser.SetCritere(NomFiltre);
  if Appliquer then
    FFiltres.OnChange(FFiltres);
end;

end.

