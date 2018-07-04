unit uLibGestionAdresse;
interface

uses controls
     ,classes
     ,forms
     ,uTob
     ,uTobDebug
     ,stdCtrls
{$IFNDEF EAGLCLIENT}
  {$IFNDEF DBXPRESS}
     ,dbtables
  {$ELSE}
     ,uDbxDataSet
  {$ENDIF}
{$ENDIF}
     ;

const             
// Combo Type Ligne
      TypeLigneDestinataire  : string = 'DST';
      TypeLigneDomiciliation : string = 'DOM';
      TypeLigneVille         : string = 'VIL';

// Combo Obligatoire
      ObligatoireObligatoire : string = 'OBL';
      ObligatoireFacultatif1 : string = 'FA1';
      ObligatoireFacultatif2 : string = 'FA2';

// Pour l'espacement entre les champs
   // Espace vertival
      INTERLIGNE : Integer = 30;
   // Espace horizontal
      ESPACEMENT : Integer = 10;


// Type champs
type
  TChampsAdresse = class
  private
     FIndice          : integer;
     FEcran           : TForm;
     FTypeLigne       : string;
     FLastTypeLigne   : string;
     FNumLigne        : integer;
     FLastNumLigne    : integer;
     FPosLigne        : integer;
     FLastPosLigne    : integer;
     FUtilise         : boolean;
     FReplace         : boolean;
     FDeplace         : boolean;
     procedure SetTypeLigne(const PTypeLigne : string);
     procedure SetNumLigne(const PNumLigne : integer);
     procedure SetPosLigne(const PPosLigne : integer);
  public
     property Indice           : integer read FIndice          write FIndice;
     property TypeLigne        : string  read FTypeLigne       write SetTypeLigne;
     property LastTypeLigne    : string  read FLastTypeLigne;
     property NumLigne         : integer read FNumLigne        write SetNumLigne;
     property LastNumLigne     : integer read FLastNumLigne;
     property PosLigne         : integer read FPosLigne        write SetPosLigne;
     property LastPosLigne     : integer read FLastPosLigne;
     property Utilise          : boolean read FUtilise         write FUtilise;
     property Replace          : boolean read FReplace         write FReplace;
     property Deplace          : boolean read FDeplace         write FDeplace;
     constructor Create (Ecran : TForm ; PIndice : integer);
     destructor  Destroy ; override;        
     procedure MAJForm(Controle : string = '');
     function GetNumLigne(Parametrage : TOB) : integer;
     function GetLastNumLigne(Parametrage : TOB) : integer;     
     function GetAddLigne(Parametrage : TOB) : integer;
  end;

  TLigneTOB = class(TOB)
  public
     Numero  : integer;
     constructor Create(LeNomTable : string; LeParent : TOB; IndiceFils : Integer); override;
     destructor  Destroy ; override;
     function    UpdateChamp(Num : integer;Valeur : string = '') : integer;
     procedure   UpdateValChamp(Num : integer;Valeur : string = '');
     procedure   DelChamp(Num : integer);
     function    GetMaxLigne : integer;
  end;

  TSectionApercu = class (TLigneTOB)
  private
  public
     constructor Create(LeNomTable : string; LeParent : TOB; IndiceFils : Integer); override;
     destructor  Destroy ; override;
     function GetLigne(var Num : integer) : TOB;
  end;

  TApercu = class
  private
  public
     SectionDst : TSectionApercu;
     SectionDom : TSectionApercu;
     SectionVil : TSectionApercu;
     Apercu     : TMemo;
     constructor Create;
     destructor  Destroy ; override;
     function GetSection(TypeLigne : string) : TSectionApercu;
     function UpdateLigApercu(Section : string;Ligne : integer) : integer;
     procedure DelLigApercu(Section : string;Ligne : integer);
     function GetRealLigne(Section : string;Ligne : integer) : integer;
     function GetMaxLigne : integer;
     function GetLigneFromApercu(Ligne : integer) : TOB;
  end;


// Procedures de chargements des enregistrements
procedure ChargeTOBRefAdr(var Champs : TOB;NumParametrage : integer = -1)  ;
procedure ChargeTOBValeur(var Champs : TOB;NumDestinataire : string = ''; NumParametrage : integer = -1 ; TypeAdr : string = '' ; SSTypeAdr : integer = -1 ; DateSaisie : TDateTime = 0);
procedure CopieDetail(Ref : TOB ; Det : TOB);
function  GetGoodValeur(var Q : TQuery ; Nom : string ) : string;
function  GetRealDate(IDDest : string ; NumParametrage : integer ; DateSaisie : TDateTime ; TypeAdresse : string ; SSTypeAdresse : integer ) : TDateTime ;

// Procedures de sauvegarde des enregistrements
procedure SavTOBRefAdr(var TOBChamps : TOB;NumParametrage : integer = -1)  ;
procedure SavTOBValeur(var TOBChamps : TOB;NumDestinataire : string = ''; NumParametrage : integer = -1 ; TypeAdr : string = '' ; SSTypeAdr : integer = -1 ; DateSaisie : TDateTime = 0 ; keepvalue : boolean = false);
procedure ModifTOBValeur(TOBChamps : TOB;NumDestinataire : string = ''; NumParametrage : integer = -1 ; TypeAdr : string = '' ; SSTypeAdr : integer = -1 ; DateSaisie : TDateTime = 0 ; LastNumDestinataire : string = '' ; LastTypeAdr : string = '' ; LastSSTypeAdr : integer = -1 ; LastDateSaisie : TDateTime = 0);
procedure Dupliquer(var TOBDst : TOB; TOBSrc : TOB ; Parametrage : integer);
// Procedure permettant de supprimer un paramétrage
procedure DelTOBRefAdr(Parametrage : integer);

// Fonction permettant de créer un composant
function CreationCopie(Component : TControl ; NomChamp : String ; TypeComponent : string):TControl;

function GetIndiceLigne(Name : string) : integer;
function GetTypeComposantLigne(Name : string) : string;
function GetMaxPosChamp(Ligne : TOB) : integer;
function GetMaxLigne(Parametre : TOB) : integer;overload;
function GetMaxLigne(Parametre : TOB ; TypeLigne : string) : integer; overload;
function GetMaxPosLigne(Parametre : TOB ; NumLigne : integer ; TypeLigne : string) : integer;
function GetMaxCode : integer;
function GetNumParametrage(PaysISO : string; TypePays : integer) : integer;
function isTypeExiste(Pays : string ; TypePays : integer ; Code : integer) : boolean;

// Restitution
procedure RestitutionAdresse( var Ligne : TOB ; var Memo : TMemo ; NumParametrage : integer = -1 ; Pays : string = '' ; TypePays : integer = -1 ; GUIDDest : string = '' ; TypeAdr : string = '' ; SSTypeAdr : integer = -1 ; DateSaisie : TDateTime = 0 ; NbLigne : integer = -1 );
procedure RestitutionAdresseType( var Ligne : TOB ; var Memo : TMemo ; NumParametrage : integer = -1 ; Pays : string = '' ; TypePays : integer = -1 ; GUIDDest : string = '' ; TypeAdr : string = '' ; SSTypeAdr : integer = -1 ; DateSaisie : TDateTime = 0 ; TypeLigne : string = '');
function CountLineUse ( Parametrage : TOB ) : integer;
procedure SupprimeFacultatif ( var Parametrage : TOB ; Facultatif : string );
procedure Restitue(Parametrage : TOB ; var Ligne : TOB ; var Memo : TMemo ; NbLigne : integer);


implementation
uses
     HCtrls
     ,SysUtils
     ,HEnt1
     ,HMsgBox
     ,HQry
     ,uLibApercu
        , DB;

{-----------------------------------------------------------------
------------------------------------------------------------------
                           CLASSE CHAMPS
------------------------------------------------------------------
-----------------------------------------------------------------}


{***********A.G.L.***********************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 19/02/2007
Modifié le ... : 19/02/2007
Description .. : Constructor de la classe Champs
Mots clefs ... :
*****************************************************************}
constructor TChampsAdresse.Create (Ecran : TForm ; PIndice : integer);
begin
  if not assigned(Ecran) then Exit;
  FIndice       := PIndice;
  FEcran        := Ecran;
  FTypeLigne     := THValComboBox(Ecran.FindComponent('TYPELIGNE_'+IntToStr(FIndice))).Value;
  FLastTypeLigne := THValComboBox(Ecran.FindComponent('LASTTYPELIGNE_'+IntToStr(FIndice))).Value;
  FNumLigne      := THSpinEdit(Ecran.FindComponent('NUMLIGNE_'+IntToStr(FIndice))).Value;
  FLastNumLigne  := THSpinEdit(Ecran.FindComponent('LASTNUMLIGNE_'+IntToStr(FIndice))).Value;
  FPosLigne      := THSpinEdit(Ecran.FindComponent('POSLIGNE_'+IntToStr(FIndice))).Value;
  FLastPosLigne  := THSpinEdit(Ecran.FindComponent('LASTPOSLIGNE_'+IntToStr(FIndice))).Value;
  FUtilise       := THCheckBox(Ecran.FindComponent('USE_'+IntToStr(FIndice))).Checked;
  FReplace       := THCheckBox(Ecran.FindComponent('REPLACE_'+IntToStr(FIndice))).Checked;
  FDeplace       := THCheckBox(Ecran.FindComponent('DEPLACE_'+IntToStr(FIndice))).Checked;
end;

destructor  TChampsAdresse.Destroy;
begin
  inherited;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 19/02/2007
Modifié le ... :   /  /
Description .. : Mis à jour du type ligne et du last type ligne
Mots clefs ... :
*****************************************************************}
procedure TChampsAdresse.SetTypeLigne (const PTypeLigne : string);
begin
  // Mis à jour des valeurs
  FTypeLigne := PTypeLigne;
  FLastTypeLigne := PTypeLigne;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 19/02/2007
Modifié le ... :   /  /
Description .. : Mis à jour du num ligne et du last num ligne
Mots clefs ... :
*****************************************************************}
procedure TChampsAdresse.SetNumLigne (const PNumLigne : integer);
begin
  // Mis à jour de la Form
  THSpinEdit(FEcran.FindComponent('NUMLIGNE_'+IntToStr(FIndice))).Value := FNumLigne;
  THSpinEdit(FEcran.FindComponent('LASTNUMLIGNE_'+IntToStr(FIndice))).Value := FLastNumLigne;
  // Mis à jour des valeurs
  FNumLigne := PNumLigne;
  FLastNumLigne := PNumLigne;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 19/02/2007
Modifié le ... :   /  /
Description .. : Mis à jour du pos ligne et du last pos ligne
Mots clefs ... :
*****************************************************************}
procedure TChampsAdresse.SetPosLigne (const PPosLigne : integer);
begin
  // Mis a jour de la form
  THSpinEdit(FEcran.FindComponent('POSLIGNE_'+IntToStr(FIndice))).Value := FPosLigne;
  THSpinEdit(FEcran.FindComponent('LASTPOSLIGNE_'+IntToStr(FIndice))).Value := LastPosLigne;
  // Mis à jour des valeurs
  FPosLigne := PPosLigne;
  FLastPosLigne := PPosLigne;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 14/02/2007
Modifié le ... :   /  /
Description .. : Retourne le numero de ligne dans l'apercu en fonction du
Suite ........ : type ligne
Mots clefs ... :
*****************************************************************}
function TChampsAdresse.GetNumLigne(Parametrage : TOB) : integer;
begin
  Result := NumLigne + GetAddLigne(Parametrage);
end;

{***********A.G.L.***********************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 14/02/2007
Modifié le ... :   /  /
Description .. : Retourne le numero de ligne dans l'apercu en fonction du
Suite ........ : type ligne
Mots clefs ... :
*****************************************************************}
function TChampsAdresse.GetLastNumLigne(Parametrage : TOB) : integer;
begin
  Result := LastNumLigne + GetAddLigne(Parametrage);
end;



{***********A.G.L.***********************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 20/02/2007
Modifié le ... :   /  /    
Description .. : Retourne le nombre de ligne à ajouter dans l'apercu
Mots clefs ... : 
*****************************************************************}
function TChampsAdresse.GetAddLigne(Parametrage : TOB) : integer;
var 
  ligne : TOB;
  Add   : integer;
begin
  Result := 0;
  Add    := 0;
  if TypeLigne = TypeLigneDestinataire then
  begin
     Exit;
  end
  else if TypeLigne = TypeLigneDomiciliation then
  begin
     Ligne := Parametrage.FindFirst(['RAD_TYPELIGNE','RAD_USE'],[TypeLigneDestinataire,'X'],true);
     while Ligne <> nil do
     begin
        if Ligne.GetValue('RAD_NUMLIGNE') > Add then
           Add := Ligne.GetValue('RAD_NUMLIGNE');
        Ligne := Parametrage.FindNext(['RAD_TYPELIGNE','RAD_USE'],[TypeLigneDestinataire,'X'],true);
     end;
     Result := Result + Add;
  end
  else if TypeLigne = TypeLigneVille then
  begin
     Ligne := Parametrage.FindFirst(['RAD_TYPELIGNE','RAD_USE'],[TypeLigneDestinataire,'X'],true);
     while Ligne <> nil do
     begin
        if Ligne.GetValue('RAD_NUMLIGNE') > Add then
           Add := Ligne.GetValue('RAD_NUMLIGNE');
        Ligne := Parametrage.FindNext(['RAD_TYPELIGNE','RAD_USE'],[TypeLigneDestinataire,'X'],true);
     end;
     Result := Result + Add;
     Add := 0;
     Ligne := Parametrage.FindFirst(['RAD_TYPELIGNE','RAD_USE'],[TypeLigneDomiciliation,'X'],true);
     while Ligne <> nil do
     begin
        if Ligne.GetValue('RAD_NUMLIGNE') > Add then
           Add := Ligne.GetValue('RAD_NUMLIGNE');
        Ligne := Parametrage.FindNext(['RAD_TYPELIGNE','RAD_USE'],[TypeLigneDomiciliation,'X'],true);
     end;
     Result := Result + Add;
  end;
end;



{***********A.G.L.***********************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 19/02/2007
Modifié le ... :   /  /
Description .. : Permet de mettre à jour l'ecran
Mots clefs ... :
*****************************************************************}
procedure TChampsAdresse.MAJForm(Controle : string = '');
begin
  if not assigned(FEcran) then Exit;
  if (Controle = 'TYPELIGNE') or (Controle = '') then
     THValComboBox(FEcran.FindComponent('TYPELIGNE_'+IntToStr(FIndice))).Value := FTypeLigne;
  if (Controle = 'LASTTYPELIGNE') or (Controle = '') then
     THValComboBox(FEcran.FindComponent('LASTTYPELIGNE_'+IntToStr(FIndice))).Value := FLastTypeLigne;
  if (Controle = 'NUMLIGNE') or (Controle = '') then
     THSpinEdit(FEcran.FindComponent('NUMLIGNE_'+IntToStr(FIndice))).Value := FNumLigne;
  if (Controle = 'LASTNUMLIGNE') or (Controle = '') then
     THSpinEdit(FEcran.FindComponent('LASTNUMLIGNE_'+IntToStr(FIndice))).Value := FLastNumLigne;
  if (Controle = 'POSLIGNE') or (Controle = '') then
     THSpinEdit(FEcran.FindComponent('POSLIGNE_'+IntToStr(FIndice))).Value := FPosLigne;
  if (Controle = 'LASTPOSLIGNE') or (Controle = '') then
     THSpinEdit(FEcran.FindComponent('LASTPOSLIGNE_'+IntToStr(FIndice))).Value := FLastPosLigne;
  if (Controle = 'USE') or (Controle = '') then
     THCheckBox(FEcran.FindComponent('USE_'+IntToStr(FIndice))).Checked := FUtilise;
end;


{-----------------------------------------------------------------
------------------------------------------------------------------
                           CLASSE LIGNE
------------------------------------------------------------------
-----------------------------------------------------------------}

{***********A.G.L.***********************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 19/02/2007
Modifié le ... :   /  /
Description .. : Constructor de la Classe TLigneTOB
Mots clefs ... :
*****************************************************************}
constructor TLigneTOB.Create(LeNomTable : string; LeParent : TOB; IndiceFils : Integer);
begin
  inherited Create(LeNomTable, LeParent, IndiceFils);
end;

{***********A.G.L.***********************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 19/02/2007
Modifié le ... :   /  /
Description .. : Destructeur de la Classe TLigneTOB
Mots clefs ... :
*****************************************************************}
destructor TLigneTOB.Destroy();
begin
  inherited;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 19/02/2007
Modifié le ... : 20/02/2007
Description .. : Permet de supprimer le Xème champ de la TOB en décalant
Suite ........ : les autres.
Mots clefs ... :
*****************************************************************}
procedure TLigneTOB.DelChamp(Num : integer);
begin
  // Suppression du champ
  DelChampSup(IntToStr(Num),false);
end;

{***********A.G.L.***********************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 19/02/2007
Modifié le ... : 20/02/2007
Description .. : Permet d'inserer le Xème champ en fin de TOB ou de 
Suite ........ : modifier sa valeur si existe deja
Mots clefs ... : 
*****************************************************************}
function TLigneTOB.UpdateChamp(Num : integer;Valeur : string = '') : integer;
var
  i      : integer;
begin
  i := GetChampCount(ttcSup);
  if i = 0 then i := 1;
  while (Num > GetChampCount(ttcSup)) do
  begin
     AddChampSup(IntToStr(i),false);
     Inc(i);
  end;

  Result := Num;
  if (Num = (GetChampCount(ttcSup) + 1)) then
  begin
     // On insere le champ en dernière place
     AddChampSup(IntToStr(Num),false);
  end;
  if not(FieldExists(IntToStr(Num))) then
     // Il faut ajouter le champ       
     AddChampSup(IntToStr(Num),false);

  // MAJ de la valeur
  if Valeur <> '' then              
     UpdateValChamp(Num,Valeur);
end;

{***********A.G.L.***********************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 19/02/2007
Modifié le ... : 20/02/2007
Description .. : Permet de modifier la valeur du Xème champ de la TOB.
Mots clefs ... :
*****************************************************************}
procedure TLigneTOB.UpdateValChamp(Num : integer;Valeur : string = '');
begin
  PutValue(IntToStr(Num),Valeur);
end;

{***********A.G.L.***********************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 20/02/2007
Modifié le ... :   /  /    
Description .. : Retourne le numero de la ligne la plus grande 
Mots clefs ... : 
*****************************************************************}
function TLigneTOB.GetMaxLigne : integer;
begin
  result := Detail.Count;
end;
                    
{-----------------------------------------------------------------
------------------------------------------------------------------
                           TSECTIONAPERCU
------------------------------------------------------------------
-----------------------------------------------------------------}
{***********A.G.L.***********************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 19/02/2007
Modifié le ... :   /  /
Description .. : Constructor de la Classe TSectionApercu
Mots clefs ... :
*****************************************************************}
constructor TSectionApercu.Create(LeNomTable : string; LeParent : TOB; IndiceFils : Integer);
begin
  inherited Create(LeNomTable, LeParent, IndiceFils);
end;

{***********A.G.L.***********************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 19/02/2007
Modifié le ... :   /  /
Description .. : Destructeur de la Classe TLigneTOB
Mots clefs ... :
*****************************************************************}
destructor TSectionApercu.Destroy();
begin
  inherited;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 19/02/2007
Modifié le ... : 20/02/2007
Description .. : Permet de recuperer la X ème ligne de la section ou la 
Suite ........ : dernière si elle n'existe pas.
Mots clefs ... : 
*****************************************************************}
function TSectionApercu.GetLigne(var Num : integer) : TOB;
begin
  while (Num > Detail.Count) do
  begin
     TLigneTOB.Create('',self,-1);
  end;
  Result := Detail[Num-1];
end;





{-----------------------------------------------------------------
------------------------------------------------------------------
                           TAPERCU
------------------------------------------------------------------
-----------------------------------------------------------------}
{***********A.G.L.***********************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 19/02/2007
Modifié le ... :   /  /
Description .. : Constructor de la Classe TSectionApercu
Mots clefs ... :
*****************************************************************}
constructor TApercu.Create;
begin
   SectionDst := TSectionApercu.Create('Section Destinataire', nil, -1);
   SectionDom := TSectionApercu.Create('Section Domiciliation', nil, -1);
   SectionVil := TSectionApercu.Create('Section Ville', nil, -1);
end;

{***********A.G.L.***********************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 19/02/2007
Modifié le ... :   /  /
Description .. : Destructeur de la Classe TLigneTOB
Mots clefs ... :
*****************************************************************}
destructor TApercu.Destroy();
begin
  inherited;
  if assigned(SectionDst) then
     FreeAndNil(SectionDst);
  if assigned(SectionDom) then
     FreeAndNil(SectionDom);
  if assigned(SectionVil) then
     FreeAndNil(SectionVil);
end;

{***********A.G.L.***********************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 19/02/2007
Modifié le ... : 21/02/2007
Description .. : Retourne la bonne section selon le type de section
Mots clefs ... :
*****************************************************************}
function TApercu.GetSection(TypeLigne : string) : TSectionApercu;
begin
  if TypeLigne = TypeLigneDestinataire then
     Result := SectionDst
  else if TypeLigne = TypeLigneDomiciliation then
     Result := SectionDom
  else if TypeLigne = TypeLigneVille then
     Result := SectionVil
  else
     Result := SectionDst;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 21/02/2007
Modifié le ... :   /  /    
Description .. : Met à jour une ligne de l'aperçu
Mots clefs ... : 
*****************************************************************}
function TApercu.UpdateLigApercu(Section : string;Ligne : integer) : integer;
var
  i : integer;
begin
  i := GetRealLigne(Section,Ligne);
  result := UpdateLigneApercu(Apercu,i,'',GetLigneFromApercu(i));
  for i := 1 to GetMaxLigne do
  begin
     UpdateLigneApercu(Apercu,i,'',GetLigneFromApercu(i));
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 21/02/2007
Modifié le ... :   /  /
Description .. : Supprime une ligne de l'apercu
Mots clefs ... :
*****************************************************************}
procedure TApercu.DelLigApercu(Section : string;Ligne : integer);
var
  Sec : TSectionApercu;
  Lig : TOB;    
  Lig2: TOB;
  i   : integer;
begin
  // MAJ de l'apercu
  DelLigneApercu(Apercu,GetRealLigne(Section,Ligne));

  // MAJ de la TOB
  Sec := GetSection(Section);
  while (ligne < Sec.Detail.Count) do
  begin
     Sec.Detail[Ligne].ChangeParent(Sec,Ligne-1);
     Inc(Ligne);
  end;
  Lig := Sec.Detail[Ligne-1];
  FreeAndNil(Lig);
  for i := Sec.Detail.Count-1 downto 0 do
  begin
     Lig2 := Sec.Detail[i];
     if Lig2.GetChampCount(ttcAll) = 0 then
        FreeAndNil(Lig2);
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 21/02/2007
Modifié le ... :   /  /    
Description .. : Retourne le vrai numero de ligne dans l'apercu
Mots clefs ... : 
*****************************************************************}
function TApercu.GetRealLigne(Section : string;Ligne : integer) : integer;
begin
  if Section = TypeLigneDestinataire then
  begin
     Result := Ligne;
  end
  else if Section = TypeLigneDomiciliation then
  begin
     Result := SectionDst.Detail.Count + Ligne;
  end
  else
  begin
     Result := SectionDom.Detail.Count + SectionDst.Detail.Count + Ligne;
  end
end;

{***********A.G.L.***********************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 21/02/2007
Modifié le ... : 21/02/2007
Description .. : Retourne le nombre total de ligne dans l'apercu
Mots clefs ... :
*****************************************************************}
function TApercu.GetMaxLigne : integer;
begin
  Result := SectionVil.Detail.Count + SectionDom.Detail.Count + SectionDst.Detail.Count;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 21/02/2007
Modifié le ... : 21/02/2007
Description .. : Retourne le nombre total de ligne dans l'apercu
Mots clefs ... :
*****************************************************************}
function TApercu.GetLigneFromApercu(Ligne : integer) : TOB;
begin
  if Ligne <= SectionDst.Detail.Count then
     Result := SectionDst.GetLigne(Ligne)
  else if Ligne <= (SectionDst.Detail.Count + SectionDom.Detail.Count) then
  begin
     Ligne := Ligne - SectionDst.Detail.Count;
     Result := SectionDom.GetLigne(Ligne);
  end
  else
  begin
     Ligne := Ligne - SectionDom.Detail.Count - SectionDst.Detail.Count;
     Result := SectionVil.GetLigne(Ligne);
  end
end;

{-----------------------------------------------------------------
------------------------------------------------------------------
                           AUTRES
------------------------------------------------------------------
-----------------------------------------------------------------}

{***********A.G.L.***********************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 14/02/2007
Modifié le ... :   /  /    
Description .. : Permet de mettre en TOB les données des necessaires au 
Suite ........ : paramétrages
Mots clefs ... : 
*****************************************************************}
procedure ChargeTOBRefAdr(var Champs : TOB;NumParametrage : integer = -1) ;
var
  i,j    : integer;
  Rech   : TOB; 
  SQL    : string;
  Q      : TQuery;
begin
  // On charge les enregistrements par défaut :
  Champs.LoadDetailDB('REFERADRE','','RAD_CODE',nil,false);
  // On rajoute le fait qu'aucune n'est utilisée pour le moment.
  Champs.Detail[0].AddChampSupValeur('RAD_USE','-',true);
  SQL := 'SELECT * FROM DETAILADRE WHERE DAD_PARCODE="' + IntToStr(NumParametrage) + '" ORDER BY DAD_REFCODE';
  if (NumParametrage <> -1) and ExisteSQL(SQL) then
  begin
     Q := OpenSQL(SQL,true);
     try
       Rech := TOB.Create('',nil,-1);
       Rech.LoadDetailDB('DETAILADRE','','',Q,false);
       // On va modifier les enregistrements par défaut selon le paramétrage
       i := 0;
       for j := 0 to Rech.Detail.Count -1 do
       begin
          while Champs.Detail[i].GetValue('RAD_CODE') < Rech.Detail[j].GetValue('DAD_REFCODE') do
             Inc(i);
          if Champs.Detail[i].GetValue('RAD_CODE') = Rech.Detail[j].GetValue('DAD_REFCODE') then
             CopieDetail(Champs.Detail[i],Rech.Detail[j]);
       end;
     finally
       Ferme(Q);
       if assigned(Rech) then
          Rech.Free;
     end;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 05/03/2007
Modifié le ... :   /  /    
Description .. : Permet de rajouter (si ce n'est pas deja le cas) un champs 
Suite ........ : valeur à la TOB champs et de le renseigner en fonction de 
Suite ........ : l'existant
Mots clefs ... : 
*****************************************************************}
procedure ChargeTOBValeur(var Champs : TOB;NumDestinataire : string = ''; NumParametrage : integer = -1 ; TypeAdr : string = '' ; SSTypeAdr : integer = -1 ; DateSaisie : TDateTime = 0);
var
  SQL     : string;
  Q       : TQuery;
  i       : integer;
  existe  : boolean;
  temp    : string;
  typetmp : string;
  lastdate: tdatetime;
  date    : tdatetime;
begin
  // Rajout des champs valeur si besoin.
  if not(Champs.Detail[0].FieldExists('RAD_VALEUR')) then
     Champs.Detail[0].AddChampSupValeur('RAD_VALEUR','',true);

  // On verifie que tous les type d'adresse renvoie la meme adresse
  typetmp := TypeAdr;
  existe := true;
  lastdate := 0;
  temp := ReadTokenSt(typetmp);
  while ( temp <> '' ) and existe do
  begin
     SQL := 'SELECT ##TOP 1## * FROM VALEURADRE '+
            'WHERE VAD_DEST = "' + NumDestinataire + '" ' +
            'AND VAD_PARCODE = "' + IntToStr(NumParametrage) + '" ' +
            'AND VAD_DATESAISIE <= "' + usDateTime(DateSaisie) + '" ' +
            'AND VAD_TYPEADR LIKE "%' + temp + '%" ' +
            'AND VAD_SSTYPEADR = "' + IntToStr(SSTypeAdr) + '"';
     SQL := SQL + ' ORDER BY VAD_DATESAISIE DESC';
     existe := ExisteSQL(SQL);
     if existe then
     begin
        Q := OpenSQL(SQL,false);
{$IFDEF EAGLCLIENT}
        date := Q.Detail[0].GetValue('VAD_DATESAISIE');
{$ELSE}
        for i := 0 to Q.FieldCount - 1 do
           if Q.Fields[i].FieldName = 'VAD_DATESAISIE' then
           begin
              date := Q.Fields[i].AsDateTime;
              break;
           end;
{$ENDIF}
        ferme(Q);
        if lastdate <> 0 then
           existe := (lastdate = date);
        lastdate := date;
     end;
     temp := ReadTokenSt(typetmp);
  end;

  if ( NumDestinataire <> '' ) and
     ( NumParametrage  <> -1 ) and
       existe                  then
  begin
     Q := OpenSQL(SQL,true);
     try
        for i := 0 to Champs.Detail.Count - 1 do
        begin
          if Champs.Detail[i].GetValue('RAD_USE') <> 'X' then continue;
          // Les champs de la tables VALEURADRE correspondent aux enregistrements
          // de la table REFERADRE par l'astuce suivante :
          // VAD_CHAMP_X ou X est la valeur de RAD_CODE
          Champs.Detail[i].PutValue('RAD_VALEUR',GetGoodValeur(Q,Champs.Detail[i].GetValue('RAD_CODE')));
        end;
     finally
        Ferme(Q);
     end;
  end
  else
  begin
     for i := 0 to Champs.Detail.Count - 1 do
        Champs.Detail[i].PutValue('RAD_VALEUR','');
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 23/02/2007
Modifié le ... :   /  /
Description .. : Permet de sauvegarder le paramétrage dans la BDD
Mots clefs ... :
*****************************************************************}
procedure SavTOBRefAdr(var TOBChamps : TOB;NumParametrage : integer = -1)  ;
var
  i         : integer;
  TOBTable  : TOB;
  TOBDelete : TOB;
  TOBUpdate : TOB; 
  TOBInsert : TOB;
  TOBTemp   : TOB; 
  TOBTemp2  : TOB;
  SQL       : string;      
  Q         : TQuery;
  value     : string;
begin
  TOBTable  := TOB.Create('DETAILADRE',nil,-1);
  TOBDelete := TOB.Create('DETAILADRE',nil,-1);
  TOBUpdate := TOB.Create('DETAILADRE',nil,-1);
  TOBInsert := TOB.Create('DETAILADRE',nil,-1);
  // Pour marquer les enregistrements traité
  if not(TOBChamps.Detail[0].FieldExists('RAD_TRAITE')) then
     TOBChamps.Detail[0].AddChampSupValeur('RAD_TRAITE','-',true);
  try
     if NumParametrage <> -1 then
     begin
        // On met à jour les enregistrements existants
        SQL := 'SELECT * FROM DETAILADRE WHERE DAD_PARCODE="' + IntToStr(NumParametrage) + '"';
        Q := OpenSQL(SQL,false);
        TOBTable.LoadDetailDB('DETAILADRE','','',Q,false);
        for i := 0 to TOBTable.Detail.Count - 1 do
        begin
           value := TOBTable.Detail[i].GetValue('DAD_REFCODE');      
           TOBTemp := TOBChamps.FindFirst(['RAD_CODE'],[value],true);
           if TOBTemp = nil then
           begin
              // Le paramétrage n'est plus dans la table référence on supprime
              if TOBDelete.GetValue('DAD_REFCODE') = 0 then
                 TOBTemp2 := TOBDelete
              else
                 TOBTemp2 := TOB.Create('DETAILADRE',TOBDelete,-1);
              Dupliquer(TOBTemp2,TOBTemp,NumParametrage);

              // On marque le champ comme traité
              TOBTemp.PutValue('RAD_TRAITE','X');
           end
           else if TOBTemp.GetValue('RAD_USE') = '-' then
           begin
              // On a desactiver la zone on supprime
              if TOBDelete.GetValue('DAD_REFCODE') = 0 then
                 TOBTemp2 := TOBDelete
              else
                 TOBTemp2 := TOB.Create('DETAILADRE',TOBDelete,-1);
              Dupliquer(TOBTemp2,TOBTemp,NumParametrage);

              // On marque le champ comme traité
              TOBTemp.PutValue('RAD_TRAITE','X');
           end
           else
           begin
              // Le champ existe on le MAJ
              TOBTemp2 := TOB.Create('DETAILADRE',TOBUpdate,-1);
              Dupliquer(TOBTemp2,TOBTemp,NumParametrage);
              // Pour prendre en compte le champs taille = 0  
              TOBTemp2.SetModifieField('DAD_TAILLE',true);

              // On marque le champ comme traité
              TOBTemp.PutValue('RAD_TRAITE','X');
           end;
        end;
        // On supprime les enregistrements
        if TOBDelete.GetValue('DAD_REFCODE') <> 0 then
           TOBDelete.DeleteDB(false);
        // On MAJ les enregistrements existants:
        if TOBUpdate.Detail.Count > 0 then
           TOBUpdate.UpdateDB(false);
     end;

     // On insert les nouveaux champs
     TOBTemp := TOBChamps.FindFirst(['RAD_TRAITE'],['-'],true);
     while (TOBTemp <> nil) do
     begin
        if TOBTemp.GetValue('RAD_USE') <> 'X' then
        begin
           // On passe au suivant
           TOBTemp := TOBChamps.FindNext(['RAD_TRAITE'],['-'],true);
           Continue;
        end;
        // Le champ existe on le MAJ
        if TOBInsert.GetValue('DAD_REFCODE') = 0 then
           TOBTemp2 := TOBInsert
        else
           TOBTemp2 := TOB.Create('DETAILADRE',TOBInsert,-1);
        Dupliquer(TOBTemp2,TOBTemp,NumParametrage);
        // On marque le champ comme traité
        TOBTemp.PutValue('RAD_TRAITE','X');
        // On passe au suivant
        TOBTemp := TOBChamps.FindNext(['RAD_TRAITE'],['-'],true);
     end;
     // On insert les enregistrements
     if TOBInsert.GetValue('DAD_REFCODE') <> 0 then
        TOBInsert.InsertDB(nil,false);
  finally
     TOBTable.Free;
     TOBDelete.Free;
     TOBUpdate.Free;
     TOBInsert.Free;
     Ferme(Q);
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 07/03/2007
Modifié le ... :   /  /    
Description .. : Permet d'enregistrer l'adresse d'un destinataire
Mots clefs ... : 
*****************************************************************}
procedure SavTOBValeur(var TOBChamps : TOB;NumDestinataire : string = ''; NumParametrage : integer = -1 ; TypeAdr : string = '' ; SSTypeAdr : integer = -1 ; DateSaisie : TDateTime = 0 ; keepvalue : boolean = false);
var
  TOBTable  : TOB;
  SQL       : string;
  Q         : TQuery;
  i         : integer;
  Name      : string;
  insert    : boolean;
  DateSav   : TDateTime;
begin
  SQL := 'SELECT ##TOP 1## * FROM VALEURADRE '+
         'WHERE VAD_DEST = "' + NumDestinataire + '" ' +
         'AND VAD_PARCODE = "' + IntToStr(NumParametrage) + '" ' +
         'AND VAD_SSTYPEADR = "' + IntToStr(SSTypeAdr) + '" ' +
         'AND VAD_DATESAISIE = "' + usDateTime(DateSaisie) + '" ' +
         'AND VAD_TYPEADR = "' + TypeAdr + '" ' +
         'ORDER BY VAD_DATESAISIE DESC';

  TOBTable  := TOB.Create('VALEURADRE',nil,-1);
  try
     if ExisteSQL(SQL) then
     begin
        // Il y a un paramétrage à mettre à jour
        Q := OpenSQL(SQL,false);
        TOBTable.LoadDetailDB('VALEURADRE','','',Q,false);
        insert := false;
        if keepvalue then
        begin
           // On met à jour les champs qui aurait put etre modifier
           TOBTable.Detail[0].PutValue('VAD_DATESAISIE',DateSaisie);
           TOBTable.Detail[0].PutValue('VAD_TYPEADR',TypeAdr);
           TOBTable.Detail[0].PutValue('VAD_SSTYPEADR',IntToStr(SSTypeAdr));
        end;
     end
     else
     begin
        insert := true;
        SQL := 'SELECT VAD_DATESAISIE FROM VALEURADRE '+
               'WHERE VAD_DEST = "' + NumDestinataire + '" ' +
               'AND VAD_PARCODE = "' + IntToStr(NumParametrage) + '" ' +
               'AND VAD_DATESAISIE < "' + usDateTime(DateSaisie) + '" ' +
               'AND VAD_TYPEADR = "' + TypeAdr + '" ' +
               'AND VAD_SSTYPEADR = "' + IntToStr(SSTypeAdr) + '"';
        SQL := SQL + ' ORDER BY VAD_DATESAISIE DESC';
        if ExisteSQL(SQL) then
        begin
           // On récupere la date exacte pour ne pas se tromper d'enreg par la suite.
           Q := OpenSQL(SQL,false);
{$IFDEF EAGLCLIENT}
           DateSav := Q.Detail[0].GetValue('VAD_DATESAISIE');
{$ELSE}
           DateSav := Q.Fields[0].AsDateTime;
{$ENDIF}
           ferme(Q);
           // Il existe un paramétrage plus ancien
           SQL := 'SELECT ##TOP 1## * FROM VALEURADRE '+
                  'WHERE VAD_DEST = "' + NumDestinataire + '" ' +
                  'AND VAD_PARCODE = "' + IntToStr(NumParametrage) + '" ' +
                  'AND VAD_DATESAISIE = "' + usDateTime(DateSav) + '" ' +
                  'AND VAD_TYPEADR = "' + TypeAdr + '" ' +
                  'AND VAD_SSTYPEADR = "' + IntToStr(SSTypeAdr) + '"';
           for i := 0 to TOBChamps.Detail.Count - 1 do
           begin
              if TOBChamps.Detail[i].GetValue('RAD_USE') = 'X' then
              begin
                 Name := 'VAD_CHAMP_' + ToString(TOBChamps.Detail[i].GetValue('RAD_CODE'));
                 SQL := SQL + ' AND ' + Name + '="' + TOBChamps.Detail[i].GetValue('RAD_VALEUR') + '"';
              end;
           end;
           SQL := SQL + ' ORDER BY VAD_DATESAISIE DESC';
           if ExisteSQL(SQL) then
           begin
              // L'adresse n'a pas changé
              Q := OpenSQL(SQL,false);
              TOBTable.LoadDetailDB('VALEURADRE','','',Q,false);
              // il faut juste mettre à jour la date ?? A VALIDER
              TOBTable.Detail[0].PutValue('VAD_DATESAISIE',DateSaisie);
              insert := false;
           end;
        end;
        if insert then
        begin
           // On renseigne les champs par defaut
           TOBTable.PutValue('VAD_PARCODE',IntToStr(NumParametrage));
           TOBTable.PutValue('VAD_DEST',NumDestinataire);
           TOBTable.PutValue('VAD_TYPEADR',TypeAdr);
           TOBTable.PutValue('VAD_SSTYPEADR',IntToStr(SSTypeAdr));
           TOBTable.PutValue('VAD_DATESAISIE',DateSaisie);
        end;
     end;
     for i := 0 to TOBChamps.Detail.Count - 1 do
     begin
       if TOBChamps.Detail[i].GetValue('RAD_USE') = 'X' then
       begin
          Name := 'VAD_CHAMP_' + ToString(TOBChamps.Detail[i].GetValue('RAD_CODE'));
          if TOBTable.Detail.Count > 0 then
             TOBTable.Detail[0].PutValue(Name,TOBChamps.Detail[i].GetValue('RAD_VALEUR'))
          else
             TOBTable.PutValue(Name,TOBChamps.Detail[i].GetValue('RAD_VALEUR'));
       end;
     end;
     if insert then
        TOBTable.InsertDB(nil)
     else
        TOBTable.UpdateDB;
  finally
     TOBTable.Free;
     if assigned(Q) then
        Ferme(Q);
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 27/02/2007
Modifié le ... :   /  /
Description .. : Permet de modifier le parametrage d'une adresse
Suite ........ : 
Mots clefs ... :
*****************************************************************}
procedure ModifTOBValeur(TOBChamps : TOB;NumDestinataire : string = ''; NumParametrage : integer = -1 ; TypeAdr : string = '' ; SSTypeAdr : integer = -1 ; DateSaisie : TDateTime = 0 ; LastNumDestinataire : string = '' ; LastTypeAdr : string = '' ; LastSSTypeAdr : integer = -1 ; LastDateSaisie : TDateTime = 0);
var
  SQL      : string;
  TOBTable : TOB;
  Q        : TQuery;
begin
  // Lecture de l’ancien paramètrage
  SQL := 'SELECT ##TOP 1## * FROM VALEURADRE '+
         'WHERE VAD_DEST = "' + LastNumDestinataire + '" ' +
         'AND VAD_PARCODE = "' + IntToStr(NumParametrage) + '" ' +
         'AND VAD_SSTYPEADR = "' + IntToStr(LastSSTypeAdr) + '" ' +
         'AND VAD_DATESAISIE <= "' + usDateTime(LastDateSaisie) + '" ' +
         'AND VAD_TYPEADR = "' + LastTypeAdr + '" ' +
         'ORDER BY VAD_DATESAISIE DESC';
  TOBTable  := TOB.Create('VALEURADRE',nil,-1);
  try
    Q := OpenSql (Sql, True) ;
    if not Q.Eof then
    begin
       with TobTable do
       begin
          if SelectDb('', Q) then // Utilisation de SelectDb(), comme cela pas de fille qui ne servent à rien…
          begin
             PutValue('VAD_PARCODE',IntToStr(NumParametrage));
             PutValue('VAD_DEST',NumDestinataire);
             PutValue('VAD_TYPEADR',TypeAdr);
             PutValue('VAD_SSTYPEADR',IntToStr(SSTypeAdr));
             PutValue('VAD_DATESAISIE',DateSaisie);
             UpdateDb();
          end;
       end;
    end
    else
    begin
       // On a pas d'ancien paramétrage
       SQL := 'SELECT ##TOP 1## * FROM VALEURADRE '+
              'WHERE VAD_DEST = "' + LastNumDestinataire + '" ' +
              'AND VAD_PARCODE = "' + IntToStr(NumParametrage) + '" ' +
              'AND VAD_SSTYPEADR = "' + IntToStr(LastSSTypeAdr) + '" ' +
              'AND VAD_DATESAISIE = "' + usDateTime(LastDateSaisie) + '" ' +
              'AND VAD_TYPEADR = "' + LastTypeAdr + '" ' +
              'ORDER BY VAD_DATESAISIE DESC';
       if not ExisteSQL(SQL) then
       begin
          with TobTable do
          begin
             // On renseigne les champs par defaut
             TOBTable.PutValue('VAD_PARCODE',IntToStr(NumParametrage));
             TOBTable.PutValue('VAD_DEST',NumDestinataire);
             TOBTable.PutValue('VAD_TYPEADR',TypeAdr);
             TOBTable.PutValue('VAD_SSTYPEADR',IntToStr(SSTypeAdr));
             TOBTable.PutValue('VAD_DATESAISIE',DateSaisie);
             UpdateDb();
          end;
       end;
    end;
  finally
     TobTable.Free ;
     If Assigned(Q) then Ferme(Q) ;
  end;
end;



{***********A.G.L.***********************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 23/02/2007
Modifié le ... :   /  /
Description .. : Permet de dupliquer la TOBSrc de format REFERADRE
Suite ........ : dans la TOBDst de format DETAILADRE
Mots clefs ... :
*****************************************************************}
procedure Dupliquer(var TOBDst : TOB; TOBSrc : TOB ; Parametrage : integer);
begin
  TOBDst.PutValue('DAD_REFCODE',TOBSrc.GetValue('RAD_CODE'));
  TOBDst.PutValue('DAD_PARCODE',IntToStr(Parametrage));
  TOBDst.PutValue('DAD_TYPELIGNE',TOBSrc.GetValue('RAD_TYPELIGNE'));
  TOBDst.PutValue('DAD_NUMLIGNE',TOBSrc.GetValue('RAD_NUMLIGNE'));
  TOBDst.PutValue('DAD_POSLIGNE',TOBSrc.GetValue('RAD_POSLIGNE'));
  TOBDst.PutValue('DAD_TAILLE',TOBSrc.GetValue('RAD_TAILLE'));
  TOBDst.PutValue('DAD_OBLIGATOIRE',TOBSrc.GetValue('RAD_OBLIGATOIRE'));
end;

{***********A.G.L.***********************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 29/03/2007
Modifié le ... :   /  /    
Description .. : Permet de récuperer la valeur d'un champ
Mots clefs ... : 
*****************************************************************}
function  GetGoodValeur(var Q : TQuery ; Nom : string ) : string;
var
  Indice  : integer;
begin
  Indice := GetIndiceLigne(Nom);
{$IFDEF EAGLCLIENT}
  Result := Q.Detail[0].GetValue('VAD_CHAMP_' + IntToStr(Indice));
{$ELSE}
  Result := Q.FieldValues['VAD_CHAMP_' + IntToStr(Indice)];
{$ENDIF}
end;

{***********A.G.L.***********************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 28/03/2007
Modifié le ... : 28/03/2007
Description .. : Permet de restituer la bonne date de saisie de l'adresse
Mots clefs ... :
*****************************************************************}
function GetRealDate(IDDest : string ; NumParametrage : integer ; DateSaisie : TDateTime ; TypeAdresse : string ; SSTypeAdresse : integer ) : TDateTime ;
var
  SQL : string;
  Q   : TQuery;
begin
  Result := 0;
  SQL := 'SELECT ##TOP 1## VAD_DATESAISIE FROM VALEURADRE '+
         'WHERE VAD_DEST = "' + IDDest + '" ' +
         'AND VAD_PARCODE = "' + IntToStr(NumParametrage) + '" ' +
         'AND VAD_DATESAISIE <= "' + usDateTime(DateSaisie) + '" ' +
         'AND VAD_TYPEADR = "' + TypeAdresse + '" ' +
         'AND VAD_SSTYPEADR = "' + IntToStr(SSTypeAdresse) + '" ' +
         'ORDER BY VAD_DATESAISIE DESC';
  Q := OpenSQL(SQL,true);
  try
     if not Q.eof then
     begin
{$IFDEF EAGLCLIENT}
        Result := Q.Detail[0].GetValue('VAD_DATESAISIE');
{$ELSE}
        Result := Q.Fields[0].AsDateTime;
{$ENDIF}
     end;
  finally
     Ferme(Q);
  end;
end;


{***********A.G.L.***********************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 29/03/2007
Modifié le ... : 29/03/2007
Description .. : Permet de supprimer un parametrage
Mots clefs ... : 
*****************************************************************}
procedure DelTOBRefAdr(Parametrage : integer);
var
  SQL : string;
begin
  SQL := 'DELETE FROM DETAILADRE WHERE DAD_PARCODE="' + IntToStr(Parametrage) + '"';
  ExecuteSQL(SQL);
end;

{***********A.G.L.***********************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 12/02/2007
Modifié le ... :   /  /
Description .. : Procedure qui permet de copier dans la TOB Ref les
Suite ........ : données de la TOB Det
Mots clefs ... : 
*****************************************************************}
procedure CopieDetail(Ref : TOB ; Det : TOB);
begin
  Ref.PutValue('RAD_TYPELIGNE',Det.GetValue('DAD_TYPELIGNE'));
  Ref.PutValue('RAD_NUMLIGNE',Det.GetValue('DAD_NUMLIGNE'));
  Ref.PutValue('RAD_POSLIGNE',Det.GetValue('DAD_POSLIGNE'));
  Ref.PutValue('RAD_TAILLE',Det.GetValue('DAD_TAILLE'));    
  Ref.PutValue('RAD_OBLIGATOIRE',Det.GetValue('DAD_OBLIGATOIRE'));
  Ref.PutValue('RAD_USE','X');
end;

{***********A.G.L.***********************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 13/02/2007
Modifié le ... :   /  /    
Description .. : Fonction permettant de dupliquer un composant
Mots clefs ... : 
*****************************************************************}
function CreationCopie(Component : TControl ; NomChamp : String ; TypeComponent : string):TControl;
var
  Stream   : TMemoryStream;
  DataType : String;
  Name     : String;
  value    : string;
begin
  result := nil;
  Stream := TMemoryStream.Create;
  try
    // Creation du composant :
    if UpperCase(TypeComponent) = 'THVALCOMBOBOX' then
    begin
        // On supprime des propriétés
        value := THValComboBox(Component).Value;
        THValComboBox(Component).Items.Clear;
        DataType := THValComboBox(Component).DataType;
        THValComboBox(Component).DataType := '';
        Result := THValComboBox.Create(Component.Owner);
    end
    else if UpperCase(TypeComponent) = 'THLABEL' then
    begin
        Result := THLabel.Create(Component.Owner);
    end
    else if UpperCase(TypeComponent) = 'THSPINEDIT' then
    begin
        Result := THSpinEdit.Create(Component.Owner);
    end
    else if UpperCase(TypeComponent) = 'THEDIT' then
    begin
        Result := THEdit.Create(Component.Owner);
    end
    else if UpperCase(TypeComponent) = 'THCHECKBOX' then
    begin
        Result := THCheckBox.Create(Component.Owner);
    end;
    // On vide temporairement le nom pour eviter les conflits.
    Name := Component.name;
    Component.name := '';

    // WriteComponent permet d'écrire toutes les propriétés publiée
    Stream.WriteComponent(Component);
    Stream.Seek(0,soFromBeginning);

    // ReadComponent permet de les lire en créant un nouveau composant
    Result := TControl(Stream.ReadComponent(Result));

    // On met à jour les propriétés qui ne l'ont pas été.
    Result.Parent := Component.Parent;
    Result.Name := NomChamp;
    Component.name := Name;
    if UpperCase(TypeComponent) = 'THVALCOMBOBOX' then
    begin
        // Cas specifique.
        THValComboBox(Result).DataType := DataType;
        THValComboBox(Component).DataType := DataType;
        THValComboBox(Result).Value := value;
        THValComboBox(Component).Value := value;
    end;
  finally
    Stream.Free;
  end;                   
end;

{***********A.G.L.***********************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 14/02/2007
Modifié le ... :   /  /    
Description .. : Retourne l'indice de la ligne en fonction du nom d'un des
Suite ........ : elements de cette ligne
Mots clefs ... : 
*****************************************************************}
function GetIndiceLigne(Name : string) : integer;
begin
  result := StrToInt(Copy(Name,Pos('_',Name)+1,2));
end;

{***********A.G.L.***********************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 14/02/2007
Modifié le ... :   /  /    
Description .. : Fonction qui retourne le type du composant
Mots clefs ... : 
*****************************************************************}
function GetTypeComposantLigne(Name : string) : string;
begin
  result := Copy(Name,0,Pos('_',Name)-1);
end;

{***********A.G.L.***********************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 23/02/2007
Modifié le ... :   /  /    
Description .. : Retourne le numéro du champ le plus elevé.
Mots clefs ... : 
*****************************************************************}
function GetMaxPosChamp(Ligne : TOB) : integer;
var
  i : integer;
begin
  Result := 0;
  for i := 1 to Ligne.GetChampCount(ttcReal) do
  begin
      if Ligne.GetNomChamp(i) = '' then continue;
      if StrToInt(Ligne.GetNomChamp(i)) > Result then
         Result := StrToInt(Ligne.GetNomChamp(i));
  end;
  for i := 1000 to Ligne.GetChampCount(ttcSup) + 999 do
  begin
      if Ligne.GetNomChamp(i) = '' then continue;
      if StrToInt(Ligne.GetNomChamp(i)) > Result then
         Result := StrToInt(Ligne.GetNomChamp(i));
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 06/03/2007
Modifié le ... : 06/03/2007
Description .. : Retourne le numéro de la ligne la plus elevé.
Mots clefs ... : 
*****************************************************************}
function GetMaxLigne(Parametre : TOB) : integer;
var
  i : integer;
  cherche : TOB;
  dst,dom,vil : integer;
begin
  dst := 0;
  dom := 0;
  vil := 0;
  for i := 0 to Parametre.Detail.Count - 1 do
  begin
     cherche := Parametre.Detail[i];
     if cherche.GetValue('RAD_USE') <> 'X' then continue;
     if ( cherche.GetValue('RAD_TYPELIGNE') = TypeLigneDestinataire ) then
        if ( cherche.GetValue('RAD_NUMLIGNE') > dst ) then
           dst := cherche.GetValue('RAD_NUMLIGNE')
     else if ( cherche.GetValue('RAD_TYPELIGNE') = TypeLigneDomiciliation ) then
        if ( cherche.GetValue('RAD_NUMLIGNE') > dom ) then
           dom := cherche.GetValue('RAD_NUMLIGNE')
     else if ( cherche.GetValue('RAD_TYPELIGNE') = TypeLigneVille ) then
        if ( cherche.GetValue('RAD_NUMLIGNE') > vil ) then
           vil := cherche.GetValue('RAD_NUMLIGNE');
  end;
  Result := dst+dom+vil;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 06/03/2007
Modifié le ... :   /  /    
Description .. : Retourne le nombre de ligne pour un type de ligne
Mots clefs ... : 
*****************************************************************}
function GetMaxLigne(Parametre : TOB ; TypeLigne : string) : integer;
var
  i : integer;
  cherche : TOB;
begin
  Result := 0;
  for i := 0 to Parametre.Detail.Count - 1 do
  begin
     cherche := Parametre.Detail[i]; 
     if cherche.GetValue('RAD_USE') <> 'X' then continue;
     if ( cherche.GetValue('RAD_TYPELIGNE') = TypeLigne ) and
        ( cherche.GetValue('RAD_NUMLIGNE') > result ) then
           result := cherche.GetValue('RAD_NUMLIGNE');
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 06/03/2007
Modifié le ... :   /  /    
Description .. : Retourne le nombre max de champs sur une ligne
Mots clefs ... : 
*****************************************************************}
function GetMaxPosLigne(Parametre : TOB ; NumLigne : integer ; TypeLigne : string) : integer;
var
  cherche : TOB;
begin
  Result := 0;
  cherche := Parametre.FindFirst(['RAD_NUMLIGNE','RAD_TYPELIGNE','RAD_USE'],[IntToStr(NumLigne),TypeLigne,'X'],TRUE) ;
  while cherche <> nil do
  begin
     if ( cherche.GetValue('RAD_POSLIGNE') > result ) then
        Result := cherche.GetValue('RAD_POSLIGNE');
     cherche := Parametre.FindNext(['RAD_NUMLIGNE','RAD_TYPELIGNE','RAD_USE'],[IntToStr(NumLigne),TypeLigne,'X'],TRUE) ;
  end;
end;
{***********A.G.L.***********************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 23/02/2007
Modifié le ... :   /  /
Description .. : Retourne le Max code de la table PARAMADRE
Mots clefs ... : 
*****************************************************************}
function GetMaxCode : integer;
var
  SQL    : string;
  Q      : TQuery;
begin
  SQL := 'SELECT MAX(PAD_CODE) FROM PARAMADRE';
  Q := OpenSQL(SQL,true);
  try
     if not(Q.Eof) then
        Result := Q.Fields[0].AsInteger
     else
        Result := 0;
  finally
     Ferme(Q);
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 05/03/2007
Modifié le ... :   /  /    
Description .. : Fonction qui retourne l'identifiant du paramétrage en 
Suite ........ : fonction du code ISO du Pays et du type pays
Mots clefs ... : 
*****************************************************************}
function GetNumParametrage(PaysISO : string; TypePays : integer) : integer;
var
  SQL    : string;
  Q      : TQuery;
begin
  SQL := 'SELECT PAD_CODE FROM PARAMADRE ' +
         'WHERE PAD_PAYS = "' + PaysISO + '" ' +
         'AND PAD_TYPE = "' + IntToStr(TypePays) + '"';
  Q := OpenSQL(SQL,true);
  try
     if not(Q.Eof) then
        Result := Q.Fields[0].AsInteger
     else
        Result := 0;
  finally
     Ferme(Q);
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 26/02/2007
Modifié le ... :   /  /    
Description .. : Indique si le Type est déjà utilisé pour ce pays
Mots clefs ... :
*****************************************************************}
function isTypeExiste(Pays : string ; TypePays : integer ; Code : integer) : boolean;
var
  SQL : string;
begin
  SQL := 'SELECT * FROM PARAMADRE WHERE ' +
         'PAD_PAYS="' + Pays + '" AND ' +
         'PAD_TYPE="' + IntToStr(TypePays) + '" AND ' +
         'PAD_CODE <> "' + IntToStr(Code) + '"';
  Result := ExisteSQL(SQL);
end;

{***********A.G.L.***********************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 22/03/2007
Modifié le ... :   /  /
Description .. : Permet de restituer un pavé adresse sous forme de ligne ou
Suite ........ : sous forme de Memo
Mots clefs ... :
*****************************************************************}
procedure RestitutionAdresse( var Ligne : TOB ; var Memo : TMemo ; NumParametrage : integer = -1 ; Pays : string = '' ; TypePays : integer = -1 ; GUIDDest : string = '' ; TypeAdr : string = '' ; SSTypeAdr : integer = -1 ; DateSaisie : TDateTime = 0 ; NbLigne : integer = -1 );
var
  monAdresse : TOB;
begin
  // Récuperation du parametrage :
  if NumParametrage = -1 then
     NumParametrage := GetNumParametrage(Pays,TypePays);
  if NumParametrage < 1 then
     Exit; // Pas de parametrage.
  monAdresse := TOB.Create('mon adresse',nil,-1);
  try
     // Chargement du paramétrage
     ChargeTOBRefAdr(monAdresse,NumParametrage);
     // Chargement des données
     ChargeTOBValeur(monAdresse,GUIDDest,NumParametrage,TypeAdr,SSTypeAdr,DateSaisie);

     // On vérifie le nombre de lignes à renvoyer
     if (NbLigne < 1) then Restitue(monAdresse,Ligne,Memo,CountLineUse(monAdresse))
     else if (CountLineUse(monAdresse) <= NbLigne) then Restitue(monAdresse,Ligne,Memo,NbLigne)
     else
     begin
        // Il y en a trop on supprime les Facultatives 2
        SupprimeFacultatif(monAdresse,ObligatoireFacultatif2);
        // On vérifie le nombre de lignes à renvoyer
        if CountLineUse(monAdresse) <= NbLigne then Restitue(monAdresse,Ligne,Memo,NbLigne)
        else
        begin
           // Il y en a trop on supprime les Facultatives 1
           SupprimeFacultatif(monAdresse,ObligatoireFacultatif1);
           // Peut importe le nombre restant on renvoie quoi qu'il arrive les lignes obligatoires.
           Restitue(monAdresse,Ligne,Memo,CountLineUse(monAdresse));
        end;
     end;                                  
  finally
     monAdresse.Free;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 27/03/2007
Modifié le ... :   /  /
Description .. : Permet de retourner les lignes Obligatoire
Suite ........ : Obligatoire + Facultative 1
Suite ........ : Obligatoire + Facultative 1 + Facultative 2
Mots clefs ... :
*****************************************************************}
procedure RestitutionAdresseType( var Ligne : TOB ; var Memo : TMemo ; NumParametrage : integer = -1 ; Pays : string = '' ; TypePays : integer = -1 ; GUIDDest : string = '' ; TypeAdr : string = '' ; SSTypeAdr : integer = -1 ; DateSaisie : TDateTime = 0 ; TypeLigne : string = '');
var
  monAdresse : TOB;
begin
  // Récuperation du parametrage :
  if NumParametrage = -1 then
     NumParametrage := GetNumParametrage(Pays,TypePays);
  if NumParametrage < 1 then
     Exit; // Pas de parametrage.
  monAdresse := TOB.Create('mon adresse',nil,-1);
  try
     // Chargement du paramétrage
     ChargeTOBRefAdr(monAdresse,NumParametrage);
     // Chargement des données
     ChargeTOBValeur(monAdresse,GUIDDest,NumParametrage,TypeAdr,SSTypeAdr,DateSaisie);

     // On vérifie le nombre de lignes à renvoyer
     if ( TypeLigne = ObligatoireFacultatif2 ) then
        // On renvoie tout
        Restitue(monAdresse,Ligne,Memo,CountLineUse(monAdresse))
     else
     begin
        // On supprime les lignes facultatives 2
        SupprimeFacultatif(monAdresse,ObligatoireFacultatif2);
        if ( TypeLigne = ObligatoireFacultatif1 ) then
           Restitue(monAdresse,Ligne,Memo,CountLineUse(monAdresse))
        else
        begin
           // On renvoie que les lignes obligatoires
           SupprimeFacultatif(monAdresse,ObligatoireFacultatif1);
           Restitue(monAdresse,Ligne,Memo,CountLineUse(monAdresse))
        end;
     end;
  finally
     monAdresse.Free;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 22/03/2007
Modifié le ... :   /  /
Description .. : Compte le nombre de ligne utilisées pour le paramétrage
Mots clefs ... :
*****************************************************************}
function CountLineUse ( Parametrage : TOB ) : integer;
var
  i     : integer;
  nbDst : integer;
  nbDom : integer;
  nbVil : integer;
begin
  nbDst := 0;
  nbDom := 0;
  nbVil := 0;
  for i := 0 to Parametrage.Detail.Count - 1 do
  begin
     if Parametrage.Detail[i].GetValue('RAD_USE') <> 'X' then continue;
     if Parametrage.Detail[i].GetValue('RAD_TYPELIGNE') = TypeLigneDestinataire then
     begin
        if StrToInt(Parametrage.Detail[i].GetValue('RAD_NUMLIGNE')) > nbDst then
           nbDst := StrToInt(Parametrage.Detail[i].GetValue('RAD_NUMLIGNE'));
     end
     else if Parametrage.Detail[i].GetValue('RAD_TYPELIGNE') = TypeLigneDomiciliation then
     begin
        if StrToInt(Parametrage.Detail[i].GetValue('RAD_NUMLIGNE')) > nbDom then
           nbDom := StrToInt(Parametrage.Detail[i].GetValue('RAD_NUMLIGNE'));
     end
     else if Parametrage.Detail[i].GetValue('RAD_TYPELIGNE') = TypeLigneVille then
     begin
        if StrToInt(Parametrage.Detail[i].GetValue('RAD_NUMLIGNE')) > nbVil then
           nbVil := StrToInt(Parametrage.Detail[i].GetValue('RAD_NUMLIGNE'));
     end;
  end;
  Result := nbDst + nbDom + nbVil;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 22/03/2007
Modifié le ... :   /  /    
Description .. : Enleve du pave adresse le type de champs Facultatif passé 
Suite ........ : en paramètre
Mots clefs ... : 
*****************************************************************}
procedure SupprimeFacultatif ( var Parametrage : TOB ; Facultatif : string );
var
  i         : integer;
  j         : integer;
  ligne     : integer;
  typeligne : string;
  Ok        : boolean;
begin
  for i := 0 to Parametrage.Detail.Count - 1 do
  begin
     if Parametrage.Detail[i].GetValue('RAD_USE') <> 'X' then continue;
     if ((Facultatif = ObligatoireFacultatif2) and (Parametrage.Detail[i].GetValue('RAD_OBLIGATOIRE') = Facultatif)) or
        ((Facultatif = ObligatoireFacultatif1) and (Parametrage.Detail[i].GetValue('RAD_OBLIGATOIRE') <> ObligatoireObligatoire)) then
     begin
        ligne := Parametrage.Detail[i].GetValue('RAD_NUMLIGNE');
        typeligne := Parametrage.Detail[i].GetValue('RAD_TYPELIGNE');
        Ok := true;
        // On recherche s'il n'y a pas un element plus elevé sur la ligne
        for j := 0 to Parametrage.Detail.Count - 1 do
        begin
           if (Parametrage.Detail[j].GetValue('RAD_USE')        <> 'X'       ) or
              (Parametrage.Detail[j].GetValue('RAD_TYPELIGNE')  <> typeligne ) or
              (Parametrage.Detail[j].GetValue('RAD_NUMLIGNE')   <> ligne     ) then  continue;
           if Facultatif = ObligatoireFacultatif2 then
              if Parametrage.Detail[j].GetValue('RAD_OBLIGATOIRE') <> Facultatif then
              begin
                 Ok := false;
                 break;
              end
           else if Facultatif = ObligatoireFacultatif1 then
              if Parametrage.Detail[j].GetValue('RAD_OBLIGATOIRE') <> ObligatoireObligatoire then
              begin
                 Ok := false;
                 break;
              end;
        end;
        if Ok then
           // Non on peut le supprimer.
           Parametrage.Detail[i].PutValue('RAD_USE','-');
     end;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 22/03/2007
Modifié le ... :   /  /
Description .. : Retourne le pave adresse soit dans un Memo soit dans une
Suite ........ : TOB
Mots clefs ... :
*****************************************************************}
procedure Restitue(Parametrage : TOB ; var Ligne : TOB ; var Memo : TMemo ; NbLigne : integer);
var
  i,j,k,l   : integer;
  TypeLigne : string;
  Res       : string;
  RealLigne : integer;
begin
  // Pour chaque type ligne ligne
  RealLigne := 0;
  for l := 1 to 3 do
  begin
     if l = 1 then TypeLigne := TypeLigneDestinataire
     else if l = 2 then TypeLigne := TypeLigneDomiciliation
     else if l = 3 then TypeLigne := TypeLigneVille;

     for i := 1 to GetMaxLigne(Parametrage,TypeLigne) do
     begin
        Res := ''; 
        Inc(RealLigne);
        // Pour chaque champ de la ligne
        for j := 1 to GetMaxPosLigne(Parametrage,i,TypeLigne) do
        begin
           // Pour chaque enregistrement du parametrage
           for k := 0 to Parametrage.Detail.Count - 1 do
           begin
              if (Parametrage.Detail[k].GetValue('RAD_USE')       = 'X'        ) and
                 (Parametrage.Detail[k].GetValue('RAD_TYPELIGNE') = TypeLigne  ) and
                 (Parametrage.Detail[k].GetValue('RAD_NUMLIGNE')  = IntToStr(i)) and
                 (Parametrage.Detail[k].GetValue('RAD_POSLIGNE')  = IntToStr(j)) then
              begin
                 if Res = '' then
                    Res := Parametrage.Detail[k].GetValue('RAD_VALEUR')
                 else
                    Res := Res + ' ' + Parametrage.Detail[k].GetValue('RAD_VALEUR');
                 break;
              end;
           end;
        end;                      
        if assigned(Ligne) then
        begin
           Ligne.AddChampSup('LIGNE_' + IntToStr(RealLigne),false);
           Ligne.PutValue('LIGNE_' + IntToStr(RealLigne),Res);
        end
        else if assigned(Memo) then
        begin
           Memo.Lines.Add(Res);
        end;
     end;
  end;
end;

end.
