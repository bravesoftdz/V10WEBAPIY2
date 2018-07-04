{***********UNITE*************************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 05/03/2007
Modifié le ... : 05/03/2007
Description .. : Composant TAdresse
Suite ........ : Permet de gerer les differents formats d'adresse
Mots clefs ... :
*****************************************************************}
unit Adresse;

interface

uses Classes
     ,Controls
     ,HCtrls
     ,Forms
     ,uTob
     ,uTobDebug
     ,StdCtrls
     ,Types;
const
    // On pose le premier THLabel en 5.5
      HDEB          : integer = 5;
      VDEB          : integer = 5;
    // On espace horizontalement 2 THEdit de 5
      HSEP          : integer = 5;
    // On espace verticalement un THEdit et un THLabel de 5
      VSEP          : integer = 5;
    // On espace verticalement le THLabel et le THEdit de 5
      INTERSTICE    : integer = 5;
    // Hauteur du THEdit par defaut
      CHAMPHAUTEUR  : integer = 21;
    // Hauteur du THLabel par defaut
      LABELHAUTEUR  : integer = 13;
    // Longueur maximum d'un champs (en caractères)
      LONGUEURMAX   : integer = 50;

      
    // Pour gerer le TabOrder :
var   ChampTabOrder     : integer = 0;
      RealLigne         : integer = 0; 

type
  TChamp = class(TCollectionItem)
  protected
    FChamp       : THEdit;
    function  GetHeight : integer;
    procedure SetHeight(Pos : integer);
    function  GetLeft : integer;
    procedure SetLeft(Pos : integer);
    function  GetMaxLength : integer;
    procedure SetMaxLength(Pos : integer);
    function  GetName : string;
    procedure SetName(Name : string);
    function  GetTabOrder : integer;
    procedure SetTabOrder(Pos : integer);
    function  GetText : WideString;
    procedure SetText(Text : WideString);
    function  GetTop : integer;
    procedure SetTop(Pos : integer);
    function  GetWidth : integer;
    procedure SetWidth(Pos : integer);  
    function  GetParent : TWinControl;
    procedure SetParent(Pos : TWinControl);
    procedure SetPropertyDefault;     
    procedure SetLargeur(Fiche : TComponent ; Longueur : integer);
  published
    property    Height    : integer     read GetHeight      write SetHeight;
    property    Left      : integer     read GetLeft        write SetLeft;
    property    MaxLength : integer     read GetMaxLength   write SetMaxLength;
    property    Name      : string      read GetName        write SetName;
    property    Parent    : TWinControl read GetParent      write SetParent;
    property    Top       : integer     read GetTop         write SetTop;
    property    TabOrder  : integer     read GetTabOrder    write SetTabOrder;
    property    Text      : WideString  read GetText        write SetText;
    property    Width     : integer     read GetWidth       write SetWidth;
  public
    constructor Create(Collection : TCollection); override;
    destructor  Destroy; override;
    procedure   Assign(Source : TPersistent); override;
  end;

  TChamps = class(TCollection)
  protected
    FChamp             : THEdit;
    function GetOwner : TPersistent; override;
    function GetItem(Index : integer) : TChamp;
    procedure SetItem(Index : integer; Value : TChamp);
  public
    constructor Create(AChamp : THEdit);
    destructor Destroy; override;
    function Add : TChamp;
    function AddItem(Item : TChamp; Index : integer) : TChamp;
    function Insert(Index : integer) : TChamp;
    property Items[index : integer] : TChamp read GetItem write SetItem; default;
  end;

  TLigne = class(TCollectionItem)
  protected
    FTop    : integer;
    FLabel  : THLabel;
    FChamps : TChamps;
    procedure SetTop(Top : integer);
    procedure SetCaption(Caption : WideString);
    function  GetCaption : WideString;
    procedure SetWidth(Pos : integer);
    function  GetWidth : integer;
  published
    property Top     : integer      read FTop       write SetTop;
    property Caption : WideString   read GetCaption write SetCaption; 
    property Width   : integer      read GetWidth   write SetWidth;
  public
    constructor Create(Collection : TCollection); override;
    destructor  Destroy; override;
    procedure   Assign(Source : TPersistent); override;
    procedure   AddChamp(Fiche : TComponent);
    procedure   CreateLabel(Fiche : TComponent);
    function    CountChamps : integer;
    function    CalcLeft(NumChamp : integer) : integer;
  end;

  TLignes = class(TCollection)
  protected
    FChamps : TChamps;
    function  GetOwner : TPersistent; override;
    function  GetItem(Index : integer) : TLigne;
    procedure SetItem(Index : integer; Value : TLigne);
  public
    constructor Create(AChamps : TChamps);
    destructor  Destroy; override;
    function    Add : TLigne;
    function    AddItem(Item : TLigne; Index : integer) : TLigne;
    function    Insert(Index : integer) : TLigne;
    function    GetTop(NumLigne : integer) : integer;
    property    Items[index : integer] : TLigne read GetItem write SetItem; default;
  end;

  TAdresse = class(TScrollBox)
  protected
    // Fiche
    Fiche             : TComponent;
    // Identifiant du Paramétrage
    NumParametrage    : integer;   
    // Code ISO du Pays
    CodeISOPays       : string;
    // Type du paramétrage pour le Pays
    TypeParametrage   : integer;
    // Type d'adresse
    TypeAdresse       : string;
    LastTypeAdresse   : string;
    // Sous type d'adresse
    SSTypeAdresse     : integer;
    LastSSTypeAdresse : integer;
    // Date de saisie
    DateSaisie        : TDateTime;
    LastDateSaisie    : TDateTime;
    // Identifiant du Destinataire
    IDDest            : string;
    LastIDDest        : string;
    // Detail du parametrage
    ListeChampsParam  : TOB;
    // Lignes
    LigneDst          : TLignes;
    LigneDom          : TLignes;
    LigneVil          : TLignes;
    // Gestion d'erreur :
    FLastError        : integer;
    FLastErrorMsg     : string;
    // Gestion d'affichage
    FAffichage        : TMemo;

    // Permet de disposer les champs sur la fiche
    procedure PoseChampsFiche;
    // Creation de ligne
    procedure CreateLigne(var TypeLigne : TLignes ; NumLigne : integer);
    // Creation de champs
    procedure CreateChamp(var TypeLigne : TLignes ; NumLigne : integer ; PosLigne : integer ; Longueur : integer ; Valeur : string ; caption : string ; DetailTOB : integer);
    // Permet de gerer l'enregistrement des données
    procedure ChampChange (Sender : TObject);
    // Permet de renseigner l'ensemble des champs
    procedure MAJValeurChamp;
    procedure CCMouseWheelDown(Sender : TObject ; Shift : TShiftState ; MousePos : TPoint ; var Handled : Boolean);
    procedure CCMouseWheelUp(Sender : TObject ; Shift : TShiftState ; MousePos : TPoint ; var Handled : Boolean);
  published
    property Left     ;
    property Width    ;
    property Height   ;
    property Top      ;
    property Name     ;
    property TabOrder ;    
    property Affichage    : TMemo   read FAffichage write FAffichage;
    property LastError    : integer read FLastError;
    property LastErrorMsg : string  read FLastErrorMsg;  
    procedure SaveData(keepvalue : boolean = false);
    procedure SupprimerAdresse;
    procedure UpdateAffichage;
    function  GetDateSaisie : TDateTime;
    procedure ChangeDateSaisie(NewDateSaisie : TDateTime ; keepvalue : boolean = false);
    procedure ChangeTypeAdresse(NewTypeAdr : string ; NewSSTypeAdr : integer ; keepvalue : boolean = false);
    procedure ChangeDestinataire(NewDestinataire : string ; NewTypeAdr : string = 'VIDE' ;  NewSSTypeAdr : integer = -1 ; NewDateSaisie : TDateTime = 0 ; keepvalue : boolean = false);
  public
    // Creation de l'objet
    constructor Create(Owner : TComponent ; GUIDDest : string ; PaysISO : string ; TypePays : Integer ; TypeAdr : string ; SSTypeAdr : integer ; Date : TDateTime);  reintroduce;
    // Destruction de l'objet
    destructor  Destroy ; override;
end;



implementation
uses uLibGestionAdresse
{$IFNDEF EAGLCLIENT}
     {$IFNDEF DBXPRESS}
     , dbtables
     {$ELSE}
     ,uDbxDataSet
     {$ENDIF}   
{$ENDIF}
     , uLibApercu
     , SysUtils
     , Graphics;
const
  // libellés des messages d'erreur
  MESS : array[1..5] of string = (
    {1}'Le Paramétrage n''existe pas.',
    {2}'Erreur lors de la création d''une ligne sur la fiche',
    {3}'Erreur lors de la création d''un champ sur la fiche',
    {4}'Erreur lors de la création de la TOB',
    {5}'Aucun enregistrement à supprimer');

{***********A.G.L.************************************************
                        CLASSE TChamp
*****************************************************************}

constructor TChamp.Create(Collection : TCollection);
begin
  inherited create(Collection);
  FChamp := TChamps(Collection).FChamp;
end;

destructor TChamp.Destroy;
begin
  inherited;
end;

procedure TChamp.Assign(Source : TPersistent);
begin
  if Source is THEdit then
  begin
     FChamp := THEdit(Source);
  end
  else
     inherited;
end;

function TChamp.GetLeft : integer;
begin
  Result := FChamp.Left;
end;

procedure TChamp.SetLeft(Pos : integer);
begin
  FChamp.Left := Pos;
end;

function  TChamp.GetTop : integer;
begin
  Result := FChamp.Top;
end;

procedure TChamp.SetTop(Pos : integer);
begin
  FChamp.Top := Pos;
end;

function  TChamp.GetMaxLength : integer;
begin
  Result := FChamp.MaxLength;
end;

procedure TChamp.SetMaxLength(Pos : integer);
begin
  FChamp.MaxLength := Pos;
end;

function  TChamp.GetName : string;
begin
  Result := FChamp.Name;
end;

procedure TChamp.SetName(Name : string);
begin
  FChamp.Name := Name;
end;

function  TChamp.GetWidth : integer;
begin
  Result := FChamp.Width;
end;

procedure TChamp.SetWidth(Pos : integer);
begin
  FChamp.Width := Pos;
end;

function  TChamp.GetHeight : integer;
begin
  Result := FChamp.Height;
end;

procedure TChamp.SetHeight(Pos : integer);
begin
  FChamp.Height := Pos;
end;

function  TChamp.GetTabOrder : integer;
begin
  Result := FChamp.TabOrder;
end;

procedure TChamp.SetTabOrder(Pos : integer);
begin
  FChamp.TabOrder := Pos;
end;

function  TChamp.GetText : WideString;
begin
  Result := FChamp.Text;
end;

procedure TChamp.SetText(Text : WideString);
begin
  FChamp.Text := Text;
end;

function  TChamp.GetParent : TWinControl;
begin
  Result := FChamp.Parent;
end;

procedure TChamp.SetParent (Pos : TWinControl);
begin
  FChamp.Parent := Pos;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 07/03/2007
Modifié le ... :   /  /    
Description .. : Fonction qui renseigne les propriétés MaxLength, Width
Mots clefs ... : 
*****************************************************************}
procedure TChamp.SetLargeur(Fiche : TComponent ; Longueur : integer);
begin
  if Longueur < 1 then
     Longueur := LONGUEURMAX;

  MaxLength := Longueur;
  Width := TForm(Fiche).Canvas.TextWidth(StringOfChar('W',Longueur)); 
end;

{***********A.G.L.***********************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 07/03/2007
Modifié le ... :   /  /    
Description .. : Permet de renseigner les propriétés par defaut 
Suite ........ : (Hauteur,Text,Heigth...)
Mots clefs ... : 
*****************************************************************}
procedure TChamp.SetPropertyDefault;
begin
  FChamp.Text     := '';
  FChamp.TabOrder := TabOrder;
  FChamp.Height   := CHAMPHAUTEUR;
  FChamp.Visible  := true;
end;


{***********A.G.L.************************************************
                        CLASSE TCHAMPS
*****************************************************************}

constructor TChamps.Create(AChamp : THEdit);
begin
  inherited Create(TChamp);
  FChamp := AChamp;
end;

destructor TChamps.Destroy;
begin
  inherited;
end;

function TChamps.GetOwner : TPersistent;
begin
  Result := FChamp;
end;

function TChamps.GetItem(Index : integer) : TChamp;
begin
  Result := TChamp(inherited GetItem(Index));
end;

procedure TChamps.SetItem(Index : integer; Value : TChamp);
begin
  inherited SetItem(Index, Value);
end;

function TChamps.Add : TChamp;
begin
  Result := TChamp(inherited Add);
end;

function TChamps.AddItem(Item : TChamp; Index : integer) : TChamp;
begin
  if Item = nil then
     Result := Add
  else
     Result := Item;
  if Assigned(Result) then
  begin
     Result.Collection := Self;
     if Index < 0 then
        Index := Count - 1;
     Result.Index := Index;
  end;
end;

function TChamps.Insert(Index : integer) : TChamp;
begin
  Result := AddItem(nil, Index);
end;

{***********A.G.L.************************************************
                        CLASSE TLIGNE
*****************************************************************}

constructor TLigne.Create(Collection : TCollection);
begin
  inherited create(Collection);
  FChamps := TLignes(Collection).FChamps;
end;

destructor TLigne.Destroy;
begin
  inherited;
end;

procedure TLigne.Assign(Source : TPersistent);
begin
  if Source is TChamps then
  begin
     FChamps := TChamps(Source);
  end
  else
     inherited;
end;

procedure TLigne.SetTop(Top : integer);
begin
  FTop := Top; 
end;

procedure TLigne.SetCaption(Caption : WideString);
begin
  FLabel.Caption := Caption;
end;

function  TLigne.GetCaption : WideString;
begin
  Result := FLabel.Caption;
end;

procedure TLigne.SetWidth(Pos : integer);
begin
  FLabel.Width := Pos;
end;

function  TLigne.GetWidth : integer; 
begin
  Result := FLabel.Width;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 07/03/2007
Modifié le ... :   /  /
Description .. : Procedure permettant de rajouter un champ à la collection
Suite ........ : en lui donnant quelques propriétés par defaut
Mots clefs ... :
*****************************************************************}
procedure TLigne.AddChamp(Fiche : TComponent);
begin
  FChamps.Add;
  FChamps[FChamps.Count - 1].FChamp := THEdit.Create(Fiche);
  FChamps[FChamps.Count - 1].SetPropertyDefault;
  Inc(ChampTabOrder);
  FChamps[FChamps.Count - 1].SetTabOrder(ChampTabOrder);
end;

{***********A.G.L.***********************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 07/03/2007
Modifié le ... :   /  /    
Description .. : Création du label associé à la ligne
Mots clefs ... : 
*****************************************************************}
procedure TLigne.CreateLabel(Fiche : TComponent);
begin
  if not(assigned(FLabel)) then
     FLabel := THLabel.Create(Fiche);
  FLabel.Name := 'LABEL_' + IntToStr(RealLigne);
  FLabel.Top := Top;
  FLabel.Left := HDEB;
  FLabel.Caption := '';
  FLabel.Width := 0;
  FLabel.Height := LABELHAUTEUR;
  FLabel.Visible := true;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 07/03/2007
Modifié le ... :   /  /    
Description .. : Retourne le nombre de champ sur la ligne
Mots clefs ... : 
*****************************************************************}
function TLigne.CountChamps : integer;
begin
  Result := FChamps.Count ;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 07/03/2007
Modifié le ... :   /  /    
Description .. : Retourne la position Left du NumChamp
Mots clefs ... : 
*****************************************************************}
function TLigne.CalcLeft(NumChamp : integer) : integer;
var
  i : integer;
begin
  Result := HDEB;
  for i := 0 to NumChamp - 2 do
  begin
     Result := Result + FChamps[i].Width + HSEP;
  end;
end;

{***********A.G.L.************************************************
                        CLASSE TLIGNES
*****************************************************************}

constructor TLignes.Create(AChamps : TChamps);
begin
  inherited Create(TLigne);
  FChamps := AChamps;
end;

destructor TLignes.Destroy;
begin
  inherited;
end;

function TLignes.GetOwner : TPersistent;
begin
  Result := FChamps;
end;

function TLignes.GetItem(Index : integer) : TLigne;
begin
  Result := TLigne(inherited GetItem(Index));
end;

procedure TLignes.SetItem(Index : integer; Value : TLigne);
begin
  inherited SetItem(Index, Value);
end;

function TLignes.Add : TLigne;
begin
  Result := TLigne(inherited Add);
end;

function TLignes.AddItem(Item : TLigne; Index : integer) : TLigne;
begin
  if Item = nil then
     Result := Add
  else
     Result := Item;
  if Assigned(Result) then
  begin
     Result.Collection := Self;
     if Index < 0 then
        Index := Count - 1;
     Result.Index := Index;
  end;
end;

function TLignes.Insert(Index : integer) : TLigne;
begin
  Result := AddItem(nil, Index);
end;

{***********A.G.L.***********************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 07/03/2007
Modifié le ... :   /  /    
Description .. : Fonction qui retourne le TOP d'une ligne
Mots clefs ... : 
*****************************************************************}
function TLignes.GetTop(NumLigne : integer) : integer;
var
  i : integer;
begin
  Result := VDEB;
  for i := 0 to NumLigne - 2 do
     Result := Result + LABELHAUTEUR + INTERSTICE + CHAMPHAUTEUR + VSEP;
end;

{***********A.G.L.************************************************
                        CLASSE TADRESSE
*****************************************************************}


{***********A.G.L.************************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 05/03/2007
Modifié le ... :   /  /
Description .. : Creation du composant TAdresse
Mots clefs ... :
*****************************************************************}
constructor TAdresse.Create(Owner : TComponent ; GUIDDest : string ; PaysISO : string ; TypePays : Integer ; TypeAdr : string ; SSTypeAdr : integer ; Date : TDateTime);
begin
  // Creation du ScrollBox sur la fiche
  inherited Create(owner);
  // Valeurs par defaut
  with self do
  begin
    Align                 := alClient;  // A rendre parametrable.
    AutoScroll            := True;
    BorderStyle           := bsNone;
    Parent                := TWinControl(owner);
    ParentBackground      := True;
    VertScrollBar.Visible := True;
    Visible               := True;    
    // Evenement
    OnMouseWheelDown      := CCMouseWheelDown;
    OnMouseWheelUp        := CCMouseWheelUp;
  end;
  
  // On sauvegarde l'id de la Fiche parente
  Fiche := Owner;

  // Destinataire
  IDDest     := GUIDDest;  
  LastIDDest := GUIDDest;

  // Paramétrage
  CodeISOPays       := PaysISO;
  TypeParametrage   := TypePays;
  TypeAdresse       := TypeAdr;
  LastTypeAdresse   := TypeAdr;
  SSTypeAdresse     := SSTypeAdr;
  LastSSTypeAdresse := SSTypeAdr;
  DateSaisie        := Date;
  LastDateSaisie    := Date;
  NumParametrage    := GetNumParametrage(PaysISO,TypePays);
  if NumParametrage = 0 then
  begin
    // Le paramétrage indiqué n'existe pas
    FLastError := 1;
    FLastErrorMsg := MESS[LastError];
    Exit;
  end;

  ListeChampsParam := TOB.Create('Ma TOB',nil,-1);
  try
     // Chargement du paramétrage:
     ChargeTOBRefAdr(ListeChampsParam,NumParametrage);
     // Chargement des valeurs
     ChargeTOBValeur(ListeChampsParam,IDDest,NumParametrage,TypeAdresse,SSTypeAdresse,DateSaisie);
     // Positionnement des champs en fonction du parametrage
     PoseChampsFiche;
  except
    FLastError := 4;
    FLastErrorMsg := MESS[LastError];
    FreeAndNil(ListeChampsParam);
    Exit;
  end;                           
end;

{***********A.G.L.************************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 05/03/2007
Modifié le ... :   /  /
Description .. : Destruction du composant TAdresse
Mots clefs ... :
*****************************************************************}
destructor TAdresse.Destroy;
begin
  inherited;
  if assigned(FAffichage) then
     FAffichage.Clear;
  if assigned(ListeChampsParam) then
     FreeAndNil(ListeChampsParam);
  if assigned(LigneDst) then
     FreeAndNil(LigneDst);
  if assigned(LigneDom) then
     FreeAndNil(LigneDom);
  if assigned(LigneVil) then
     FreeAndNil(LigneVil);
end;

{***********A.G.L.***********************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 08/03/2007
Modifié le ... :   /  /    
Description .. : Permet de sauvegarder les changements
Mots clefs ... : 
*****************************************************************}
procedure TAdresse.SaveData(keepvalue : boolean = false);
begin
  if not(assigned(ListeChampsParam)) then Exit;
  if keepvalue then
     // On verifie que le parametrage n'a pas bougé
     if ( TypeAdresse   <> LastTypeAdresse   ) or
        ( SSTypeAdresse <> LastSSTypeAdresse ) or
        ( DateSaisie    <> LastDateSaisie    ) or
        ( IDDest        <> LastIDDest        ) then
     // Il a bougé on doit le mettre à jour.
     ModifTOBValeur(ListeChampsParam,IDDest,NumParametrage,TypeAdresse,SSTypeAdresse,DateSaisie,LastIDDest,LastTypeAdresse,LastSSTypeAdresse,LastDateSaisie);
  // Sauvegarde des champs
  SavTOBValeur(ListeChampsParam,IDDest,NumParametrage,TypeAdresse,SSTypeAdresse,DateSaisie,keepvalue);

  // On sauvegarde les anciennes valeurs
  LastTypeAdresse   := TypeAdresse;
  LastSSTypeAdresse := SSTypeAdresse;
  LastDateSaisie    := DateSaisie;
  LastIDDest        := IDDest;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 27/03/2007
Modifié le ... :   /  /    
Description .. : Permet de supprimer l'adresse de la base de donnée
Mots clefs ... : 
*****************************************************************}
procedure TAdresse.SupprimerAdresse;
var
  SQL      : string;
  Q        : TQuery;
  RealDate : TDateTime;
  i        : integer;
begin
  SQL := 'SELECT ##TOP 1## VAD_DATESAISIE FROM VALEURADRE ' +
         'WHERE VAD_PARCODE = "' + IntToStr(NumParametrage) + '" ' +
         'AND VAD_DEST = "' + IDDest + '" ' +
         'AND VAD_TYPEADR = "' + TypeAdresse + '" ' +
         'AND VAD_SSTYPEADR = "' + IntToStr(SSTypeAdresse) + '" ' +
         'AND VAD_DATESAISIE <= "' + usDateTime(DateSaisie) + '" ' +
         'ORDER BY VAD_DATESAISIE DESC';
  if not(ExisteSQL(SQL)) then
  begin
     FLastError := 5;
     FLastErrorMsg := MESS[LastError];
     Exit; // Aucun enregistrement à supprimer
  end;
  // On récupere la date exacte pour ne pas se tromper d'enreg par la suite.
  Q := OpenSQL(SQL,false);
{$IFDEF EAGLCLIENT}
  RealDate := Q.Detail[0].GetValue('VAD_DATESAISIE');
{$ELSE}
  RealDate := Q.Fields[0].AsDateTime;
{$ENDIF}
  ferme(Q);
  SQL := 'DELETE FROM VALEURADRE WHERE VAD_PARCODE = "' + IntToStr(NumParametrage) +
         '" AND VAD_DEST = "' + IDDest +
         '" AND VAD_TYPEADR = "' + TypeAdresse +
         '" AND VAD_SSTYPEADR = "' + IntToStr(SSTypeAdresse) +
         '" AND VAD_DATESAISIE = "' + usDateTime(RealDate) + '"';
  ExecuteSQL(SQL);

  // On remet tous les champs à vide
  for i := 0 to ListeChampsParam.Detail.Count - 1 do
     ListeChampsParam.Detail[i].PutValue('RAD_VALEUR','');
  MAJValeurChamp;
end;

{***********A.G.L.************************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 05/03/2007
Modifié le ... :   /  /
Description .. : Pose les champs necessaire sur la Fiche
Mots clefs ... :
*****************************************************************}
procedure TAdresse.PoseChampsFiche;
var
  i,j,k       : integer;
  cherche     : TOB;
begin
  // Pour chaque ligne
  // Destinataire
  LigneDst := TLignes.Create(nil);
  cherche := nil;
  for j := 1 to GetMaxLigne(ListeChampsParam,TypeLigneDestinataire) do
  begin
     CreateLigne(LigneDst,j);
     // Pour chaque champ de la ligne
     for i :=  1 to GetMaxPosLigne(ListeChampsParam,j,TypeLigneDestinataire) do
     begin
        for k := 0 to ListeChampsParam.Detail.Count - 1 do
        begin
           cherche := ListeChampsParam.Detail[k];
           if (cherche.GetValue('RAD_TYPELIGNE') = TypeLigneDestinataire) and
              (cherche.GetValue('RAD_NUMLIGNE')  = IntToStr(j)          ) and
              (cherche.GetValue('RAD_POSLIGNE')  = IntToStr(i)          ) and
              (cherche.GetValue('RAD_USE')       = 'X'                  ) then
              break;
          cherche := nil;
        end;
        if cherche = nil then continue;
        CreateChamp(LigneDst,j,i,cherche.GetValue('RAD_TAILLE'), cherche.GetValue('RAD_VALEUR'), cherche.GetValue('RAD_LIBELLE'), k);
     end;
  end;

  // Domiciliation
  LigneDom := TLignes.Create(nil);
  for j := 1 to GetMaxLigne(ListeChampsParam,TypeLigneDomiciliation) do
  begin
     CreateLigne(LigneDom,j);
     // Pour chaque champ de la ligne
     for i :=  1 to GetMaxPosLigne(ListeChampsParam,j,TypeLigneDomiciliation) do
     begin
        for k := 0 to ListeChampsParam.Detail.Count - 1 do
        begin
           cherche := ListeChampsParam.Detail[k];
           if (cherche.GetValue('RAD_TYPELIGNE') = TypeLigneDomiciliation) and
              (cherche.GetValue('RAD_NUMLIGNE')  = IntToStr(j)           ) and
              (cherche.GetValue('RAD_POSLIGNE')  = IntToStr(i)           ) and
              (cherche.GetValue('RAD_USE')       = 'X'                   ) then
              break; 
          cherche := nil;
        end;
        if cherche = nil then continue;
        CreateChamp(LigneDom,j,i,cherche.GetValue('RAD_TAILLE'), cherche.GetValue('RAD_VALEUR'), cherche.GetValue('RAD_LIBELLE'), k);
     end;
  end;

  // Ville
  LigneVil := TLignes.Create(nil);
  for j := 1 to GetMaxLigne(ListeChampsParam,TypeLigneVille) do
  begin
     CreateLigne(LigneVil,j);
     // Pour chaque champ de la ligne
     for i :=  1 to GetMaxPosLigne(ListeChampsParam,j,TypeLigneVille) do
     begin
        for k := 0 to ListeChampsParam.Detail.Count - 1 do
        begin
           cherche := ListeChampsParam.Detail[k];
           if (cherche.GetValue('RAD_TYPELIGNE') = TypeLigneVille) and
              (cherche.GetValue('RAD_NUMLIGNE')  = IntToStr(j)   ) and
              (cherche.GetValue('RAD_POSLIGNE')  = IntToStr(i)   ) and
              (cherche.GetValue('RAD_USE')       = 'X'           ) then
              break;     
          cherche := nil;
        end;
        if cherche = nil then continue;
        CreateChamp(LigneVil,j,i,cherche.GetValue('RAD_TAILLE'), cherche.GetValue('RAD_VALEUR'), cherche.GetValue('RAD_LIBELLE'), k);
     end;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 06/03/2007
Modifié le ... :   /  /
Description .. : Crée un champ dans le scroll box à la bonne ligne et à la
Suite ........ : bonne position
Mots clefs ... :
*****************************************************************}
procedure TAdresse.CreateLigne(var TypeLigne : TLignes ; NumLigne : integer);
var
  temp : integer ;
begin
  // Création de la ligne
  TypeLigne.Add;
  if NumLigne <> TypeLigne.Count then
  begin  
    FLastError := 2;
    FLastErrorMsg := MESS[FLastError];
  end;   
  Inc(RealLigne);
  // Création des champs
  TypeLigne[NumLigne-1].FChamps := TChamps.Create(nil);
  // En fonction du type ligne le calcul du top n'est pas le meme :
  temp := TypeLigne.GetTop(NumLigne);
  if TypeLigne = LigneDom then
  begin
     Temp := temp + LigneDst.GetTop(LigneDst.Count + 1 ) - HDEB ;
  end
  else if TypeLigne = LigneVil then
  begin
     Temp := temp + LigneDst.GetTop(LigneDst.Count + 1 ) - HDEB + LigneDom.GetTop(LigneDom.Count + 1) - HDEB;
  end; 
  TypeLigne[NumLigne-1].SetTop(temp);
  
  // Création du label
  TypeLigne[NumLigne-1].CreateLabel(Fiche);
  TypeLigne[NumLigne-1].FLabel.Parent := self;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 06/03/2007
Modifié le ... :   /  /
Description .. : Crée un champ dans le scroll box à la bonne ligne et à la
Suite ........ : bonne position
Mots clefs ... :
*****************************************************************}
procedure TAdresse.CreateChamp(var TypeLigne : TLignes ; NumLigne : integer ; PosLigne : integer ; Longueur : integer ; Valeur : string ; caption : string ; DetailTOB : integer);
var
  MonLabel : THLabel;
  MonChamp : TChamp;
begin
  // Creation du champ
  TypeLigne[NumLigne - 1].AddChamp(Fiche);
  if TypeLigne[NumLigne - 1].CountChamps <> PosLigne then
  begin
    FLastError := 3;
    FLastErrorMsg := MESS[FLastError];
  end;
  // Gestion du label
  MonLabel := TypeLigne[NumLigne-1].FLabel;
  if assigned(MonLabel) then
  begin
    if MonLabel.Caption = '' then
       MonLabel.Caption := caption
    else
       MonLabel.Caption := MonLabel.Caption + ', ' + caption;
    MonLabel.Width := TypeLigne[NumLigne-1].CalcLeft(PosLigne + 1) - HSEP; 
  end;

  // Gestion du champ
  MonChamp := TypeLigne[NumLigne-1].FChamps[PosLigne-1];
  MonChamp.SetLargeur(Fiche,Longueur);
  MonChamp.Left   := TypeLigne[NumLigne-1].CalcLeft(PosLigne);
  MonChamp.Top    := TypeLigne[NumLigne-1].Top + LABELHAUTEUR + INTERSTICE;
  MonChamp.Name   := 'CHAMP_' + IntToStr(DetailTOB);
  MonChamp.Text   := Valeur;
  MonChamp.Parent := Self;
  MonChamp.FChamp.OnChange := ChampChange;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 07/03/2007
Modifié le ... :   /  /    
Description .. : Procedure appellé sur le onChange des THEdit
Mots clefs ... : 
*****************************************************************}
procedure TAdresse.ChampChange(Sender : TObject);
var
 i,j : integer;
begin
  i := GetIndiceLigne(THEdit(Sender).Name);
  ListeChampsParam.Detail[i].PutValue('RAD_VALEUR',THEdit(Sender).Text);
  if assigned(FAffichage) then
  begin
     j := 1;
     if ListeChampsParam.Detail[i].GetValue('RAD_TYPELIGNE') = TypeLigneDestinataire then
        j := ListeChampsParam.Detail[i].GetValue('RAD_NUMLIGNE')
     else if ListeChampsParam.Detail[i].GetValue('RAD_TYPELIGNE') = TypeLigneDomiciliation then
        j := GetMaxLigne(ListeChampsParam,TypeLigneDestinataire) + ListeChampsParam.Detail[i].GetValue('RAD_NUMLIGNE')
     else if ListeChampsParam.Detail[i].GetValue('RAD_TYPELIGNE') = TypeLigneVille then
        j := GetMaxLigne(ListeChampsParam,TypeLigneDestinataire) + GetMaxLigne(ListeChampsParam,TypeLigneDomiciliation) + ListeChampsParam.Detail[i].GetValue('RAD_NUMLIGNE');
     UpdateChampLigneApercu(FAffichage,j,ListeChampsParam.Detail[i].GetValue('RAD_POSLIGNE'),THEdit(Sender).Text,#160);
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 20/03/2007
Modifié le ... :   /  /    
Description .. : Permet de mettre à jour le Memo d'affichage
Mots clefs ... : 
*****************************************************************}
procedure TAdresse.UpdateAffichage;
var
  i : integer;                               
  j : integer;
begin
  if not(assigned(FAffichage)) then Exit;
  for i := 0 to ListeChampsParam.Detail.Count - 1 do
  begin
    if ListeChampsParam.Detail[i].GetValue('RAD_USE') <> 'X' then continue;
    j := 1;
    if ListeChampsParam.Detail[i].GetValue('RAD_TYPELIGNE') = TypeLigneDestinataire then
        j := ListeChampsParam.Detail[i].GetValue('RAD_NUMLIGNE')
     else if ListeChampsParam.Detail[i].GetValue('RAD_TYPELIGNE') = TypeLigneDomiciliation then
        j := GetMaxLigne(ListeChampsParam,TypeLigneDestinataire) + ListeChampsParam.Detail[i].GetValue('RAD_NUMLIGNE')
     else if ListeChampsParam.Detail[i].GetValue('RAD_TYPELIGNE') = TypeLigneVille then
        j := GetMaxLigne(ListeChampsParam,TypeLigneDestinataire) + GetMaxLigne(ListeChampsParam,TypeLigneDomiciliation) + ListeChampsParam.Detail[i].GetValue('RAD_NUMLIGNE');
     UpdateChampLigneApercu(FAffichage,j,ListeChampsParam.Detail[i].GetValue('RAD_POSLIGNE'),ListeChampsParam.Detail[i].GetValue('RAD_VALEUR'),#160);
  end;
end;


{***********A.G.L.***********************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 26/03/2007
Modifié le ... :   /  /    
Description .. : Permet de changer la date de saisie sans avoir a recréer le 
Suite ........ : composant
Mots clefs ... : 
*****************************************************************}
procedure TAdresse.ChangeDateSaisie(NewDateSaisie : TDateTime ; keepvalue : boolean = false);
begin 
  DateSaisie := NewDateSaisie ;
  if not(assigned(ListeChampsParam)) or keepvalue then Exit;
  LastDateSaisie := DateSaisie;
  // Chargement des valeurs
  ChargeTOBValeur(ListeChampsParam,IDDest,NumParametrage,TypeAdresse,SSTypeAdresse,DateSaisie);
  MAJValeurChamp;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 20/03/2007
Modifié le ... :   /  /    
Description .. : Permet de changer de type adresse sans avoir à redessiner 
Suite ........ : les champs.
Mots clefs ... : 
*****************************************************************}
procedure TAdresse.ChangeTypeAdresse(NewTypeAdr : string ; NewSSTypeAdr : integer ; keepvalue : boolean = false);
begin
  TypeAdresse := NewTypeAdr ;
  SSTypeAdresse := NewSSTypeAdr ;
  if not(assigned(ListeChampsParam)) or keepvalue then Exit;
  LastTypeAdresse   := TypeAdresse;
  LastSSTypeAdresse := SSTypeAdresse;
  // Chargement des valeurs
  ChargeTOBValeur(ListeChampsParam,IDDest,NumParametrage,TypeAdresse,SSTypeAdresse,DateSaisie);
  MAJValeurChamp;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 21/03/2007
Modifié le ... :   /  /
Description .. : Procedure permettant de changer de destinataire sans avoir
Suite ........ : a replacer tous les champs.
Mots clefs ... :
*****************************************************************}
procedure TAdresse.ChangeDestinataire(NewDestinataire : string ; NewTypeAdr : string = 'VIDE' ;  NewSSTypeAdr : integer = -1 ; NewDateSaisie : TDateTime = 0 ; keepvalue : boolean = false);
begin
  if (NewTypeAdr <> 'VIDE') and (NewTypeAdr <> TypeAdresse) then
     TypeAdresse := NewTypeAdr ;
  if (NewSSTypeAdr <> -1) and (NewSSTypeAdr <> SSTypeAdresse) then
     SSTypeAdresse := NewSSTypeAdr;
  if (NewDateSaisie <> 0) and (NewDateSaisie <> DateSaisie) then
     DateSaisie := NewDateSaisie;
  IDDest := NewDestinataire;
  if not(assigned(ListeChampsParam)) or keepvalue then Exit;
  LastTypeAdresse := TypeAdresse;
  LastSSTypeAdresse := SSTypeAdresse;
  LastDateSaisie := DateSaisie;
  LastIDDest := IDDest;
  // Chargement des valeurs
  ChargeTOBValeur(ListeChampsParam,IDDest,NumParametrage,TypeAdresse,SSTypeAdresse,DateSaisie);
  MAJValeurChamp; 
end;

{***********A.G.L.***********************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 20/03/2007
Modifié le ... :   /  /    
Description .. : Permet de renseigner les champs du composant 
Suite ........ : sans avoir à les replacer sur la fiche. (appelé quand on 
Suite ........ : change de type de paramétrage)
Mots clefs ... : 
*****************************************************************}
procedure TAdresse.MAJValeurChamp;
var
  i        : integer;
  NomChamp : string;
begin
  for i := 0 to ListeChampsParam.Detail.Count - 1 do
  begin
     if ListeChampsParam.Detail[i].GetValue('RAD_USE') <> 'X' then continue;
     NomChamp := 'CHAMP_' + IntToStr(i);
     THEdit(Fiche.FindComponent(NomChamp)).Text := ListeChampsParam.Detail[i].GetValue('RAD_VALEUR');
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 28/03/2007
Modifié le ... : 28/03/2007
Description .. : Permet de restituer la bonne date de saisie de l'adresse
Mots clefs ... : 
*****************************************************************}
function TAdresse.GetDateSaisie : TDateTime;
begin
  Result := GetRealDate(IDDest,NumParametrage,DateSaisie,TypeAdresse,SSTypeAdresse)
end;

procedure TAdresse.CCMouseWheelDown;
begin
  VertScrollBar.Position  := VertScrollBar.ScrollPos + INTERSTICE + LABELHAUTEUR + CHAMPHAUTEUR + HSEP;
end;

procedure TAdresse.CCMouseWheelUp;
begin
  VertScrollBar.Position  := VertScrollBar.ScrollPos - INTERSTICE - LABELHAUTEUR - CHAMPHAUTEUR - HSEP;
end;

end.


