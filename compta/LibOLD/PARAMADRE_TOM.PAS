{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 12/02/2007
Modifié le ... :   /  /
Description .. : Source TOM de la TABLE : PARAMADRE (PARAMADRE)
Mots clefs ... : TOM;PARAMADRE
*****************************************************************}
Unit PARAMADRE_TOM ;

Interface

Uses StdCtrls, 
     Controls, 
     Classes,
     DBGrids,
{$IFNDEF EAGLCLIENT}
     db,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
     Fiche,
     FichList,
     FE_Main, // AGLLanceFiche
     HDB,
{$else}
     MaineAGL, // AGLLanceFiche
     eFiche,
     eFichList,
{$ENDIF}
     forms, 
     sysutils, 
     ComCtrls,
     HCtrls, 
     HEnt1, 
     HMsgBox, 
     uLibGestionAdresse,
     UTOM, 
     UTob,
     UTobDebug,
     Windows ;
// Lancement de la fiche AGL
procedure YY_LanceParamAdresse;

Type
  TOM_PARAMADRE = Class (TOM)
    procedure OnNewRecord                ; override ;
    procedure OnDeleteRecord             ; override ;
    procedure OnUpdateRecord             ; override ;
    procedure OnAfterUpdateRecord        ; override ;
    procedure OnAfterDeleteRecord        ; override ;
    procedure OnLoadRecord               ; override ;
    procedure OnChangeField ( F: TField) ; override ;
    procedure OnArgument ( S: String )   ; override ;
    procedure OnClose                    ; override ;
    procedure OnCancelRecord             ; override ;
    private
       Champs     : TOB;
       Apercu     : TApercu;
       gLoad      : boolean;
       gInsertion : boolean;
       gReplace   : boolean; 
       gDeplace   : boolean; 
       gSave      : boolean;
       gNew       : boolean;
       gRecursif  : string;
       NumParametrage : integer;
       LibParametrage : string;
       PaysParametrage : string;
       TypeParametrage : integer;
       ScrollBar       : TScrollBox;
{$IFNDEF EAGLCLIENT}
       FListe          : THDBGrid;
{$ELSE}
       FListe          : THGrid;
{$ENDIF}
       // General
       procedure InitVariables;
       procedure InitVariablesLigne;
       procedure InitEvenements;
       procedure InitEvenementsLigne;
       procedure DesinitEvenements;
       procedure DechargeVar;     
       procedure SaveData(ask : boolean = true) ;
       procedure SetRecursif;
       procedure UnsetRecursif;       
       procedure ChargeChampsFiche ;
       procedure CreateChampsLigne(Indice : integer) ;
       procedure RenseigneChampsLigne(Ligne : TOB ;Indice : integer) ;
       procedure MAJComposants(Name : string;Valeur : string);
       procedure MAJTobSAV(Name : string) ; 
       procedure MAJLastChamp(Champ : TChampsAdresse);
       procedure ComposantChangePlace(var Champ : TChampsAdresse ; Name : string);
       function  AjouterChamp(Champ : TChampsAdresse) : boolean;
       procedure SupprimerChamp(Champ : TChampsAdresse);
       function  GetMaxPosLigne(Champ : TChampsAdresse) : integer;
       function  GetMaxLastPosLigne(Champ : TChampsAdresse) : integer;
       function  GetChampLigne(Champ : TChampsAdresse) : integer;
       function  GetChampLastLigne(Champ : TChampsAdresse) : integer;
        // Evenement ;
       procedure THEditOnExit(Sender : TObject) ;
       procedure THCheckBoxClick(Sender : TObject) ;
       procedure THValComboBoxChange(Sender : TObject);
       procedure THSpinEditChange(Sender : TObject);
       procedure THEditChange(Sender : TObject);
       procedure ModifLigne(Name : string) ;
       procedure LigneEnabled(Indice : integer; value : boolean);
       procedure CCMouseWheelDown(Sender : TObject ; Shift : TShiftState ; MousePos : TPoint ; var Handled : Boolean);
       procedure CCMouseWheelUp(Sender : TObject ; Shift : TShiftState ; MousePos : TPoint ; var Handled : Boolean);
    end ;
Implementation
uses
    HTB97,DBCtrls;


{***********A.G.L.***********************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 12/02/2007
Modifié le ... :   /  /    
Description .. : Lancement de la fiche AGL
Mots clefs ... : 
*****************************************************************}
procedure YY_LanceParamAdresse;
begin
  AGLLanceFiche('YY', 'PARAMADR', '', '', '');
end;

procedure TOM_PARAMADRE.OnNewRecord ;
var
  i : integer;
begin
  Inherited ;
  gNew := true;

  // On gere en interne le code
  SetControlEnabled('PAD_CODE',false);
  SetControlText('PAD_CODE',IntToStr(GetMaxCode+1));
  SetField('PAD_CODE',IntToStr(GetMaxCode+1));

  // On grise toutes les lignes   
  if assigned(Champs) then
     for i := 0 to Champs.Detail.Count -1 do
        LigneEnabled(i,false);

  // On met le bon focus
  SetFocusControl('PAD_LIBELLE');
end ;

procedure TOM_PARAMADRE.OnDeleteRecord ;
begin
  Inherited ;
  DelTOBRefAdr(StrToInt(GetControlText('PAD_CODE')));
end ;

procedure TOM_PARAMADRE.OnUpdateRecord ;
begin
  Inherited ;
  gNew := false;
  SaveData(false);
end ;

procedure TOM_PARAMADRE.OnAfterUpdateRecord ;
begin
  Inherited ;
end ;

procedure TOM_PARAMADRE.OnAfterDeleteRecord ;
begin
  Inherited ;
end ;

procedure TOM_PARAMADRE.SaveData(ask : boolean = true) ;
var
  retour : integer;
begin
  gSave := false;
  retour := mrYes;
  if ask then
     retour := PGIAsk('Voulez-vous enregistrer les modifications ?',LibParametrage);
  if retour = mrYes then
  begin
     // On sauvegarde les modifications :
     SavTOBRefAdr(Champs,NumParametrage);
  end;
end;

procedure TOM_PARAMADRE.CCMouseWheelDown;
begin
  ScrollBar.VertScrollBar.Position  := ScrollBar.VertScrollBar.ScrollPos + 10;
end;

procedure TOM_PARAMADRE.CCMouseWheelUp;
begin
  ScrollBar.VertScrollBar.Position  := ScrollBar.VertScrollBar.ScrollPos - 10;
end;

procedure TOM_PARAMADRE.THEditOnExit(Sender : TObject) ;
begin
  if THEdit(Sender).Name = 'PAD_LIBELLE' then
  begin
     LibParametrage := GetControlText('PAD_LIBELLE');
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 01/03/2007
Modifié le ... :   /  /    
Description .. : Chargement d'un enregistrement
Mots clefs ... : 
*****************************************************************}
procedure TOM_PARAMADRE.OnLoadRecord ;
begin
  Inherited ;  
  SetControlEnabled('PAD_CODE',false);          
  // On sauvegarde l'enregistrement précédent si besoin
  if gSave then SaveData;
  NumParametrage := StrToInt(GetControlText('PAD_CODE'));
  LibParametrage := GetControlText('PAD_LIBELLE');
  TypeParametrage := StrToInt(GetControlText('PAD_TYPE'));
  PaysParametrage := GetControlText('PAD_PAYS');
  DechargeVar;
  DesinitEvenements;
  InitVariablesLigne;
  gLoad := true;
  gInsertion := false;
  gReplace := false;
  gDeplace := false;
  // On charge la liste des champs en TOB
  ChargeTOBRefAdr(Champs,StrToInt(GetControlText('PAD_CODE')));
  // Pour chaque champs on place les valeurs.
  ChargeChampsFiche;
  // Initialisation des evenements
  InitEvenementsLigne;
  gLoad := false;
end ;

procedure TOM_PARAMADRE.OnChangeField ( F: TField ) ;
begin
  Inherited ;
end ;

procedure TOM_PARAMADRE.OnArgument ( S: String ) ;
begin
  Inherited ;    
  InitVariables;
  InitEvenements;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 13/02/2007
Modifié le ... :   /  /
Description .. : Affectation des variables
Mots clefs ... :
*****************************************************************}
procedure TOM_PARAMADRE.InitVariables;
begin
  ScrollBar := TScrollBox(GetControl('CC'));
{$IFNDEF EAGLCLIENT}
  FListe := THDBGrid(GetControl('FListe'));
{$ELSE}
  FListe := THGrid(GetControl('FListe'));
{$ENDIF}
end;

{***********A.G.L.***********************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 13/02/2007
Modifié le ... :   /  /
Description .. : Affectation des variables
Mots clefs ... :
*****************************************************************}
procedure TOM_PARAMADRE.InitVariablesLigne;
begin
  Champs := TOB.Create('REFERADRE',nil,-1);
  Apercu := TApercu.Create;
  Apercu.Apercu := TMemo(GetControl('APERCU'));
end;

{***********A.G.L.***********************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 13/02/2007
Modifié le ... : 14/02/2007
Description .. : Affectation des evenements
Mots clefs ... :
*****************************************************************}
procedure TOM_PARAMADRE.InitEvenements;
begin
  // Evenement sur la fiche :
{$IFNDEF EAGLCLIENT}
  THDBEdit(GetControl('PAD_LIBELLE')).OnExit := THEditOnExit;
  THDBEdit(GetControl('PAD_PAYS')).OnExit := THEditOnExit;
  THDBSpinEdit(GetControl('PAD_CODE')).OnChange := THSpinEditChange;
{$ELSE}
  THEdit(GetControl('PAD_LIBELLE')).OnExit := THEditOnExit;
  THEdit(GetControl('PAD_PAYS')).OnExit := THEditOnExit;    
  THSpinEdit(GetControl('PAD_CODE')).OnChange := THSpinEditChange; 
{$ENDIF}
  TScrollBox(GetControl('CC')).OnMouseWheelDown     := CCMouseWheelDown;
  TScrollBox(GetControl('CC')).OnMouseWheelUp       := CCMouseWheelUp;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 13/02/2007
Modifié le ... : 14/02/2007
Description .. : Affectation des evenements
Mots clefs ... :
*****************************************************************}
procedure TOM_PARAMADRE.InitEvenementsLigne;
var
  i : integer;
  Nom : string;
begin
  // Evenement sur les lignes :
  // Pour chaque ligne existante :
  i := 0;
  Nom := 'CHAMPS_' + IntToStr(i);
  while (GetControl(Nom) <> nil) do
  begin
     // Pour chaque composant de la ligne:
     // TypeLigne :
     Nom := 'TYPELIGNE_' + IntToStr(i);
     THValComboBox(GetControl(Nom)).OnChange := THValComboBoxChange;
     
     //  NumLigne :
     Nom := 'NUMLIGNE_' + IntToStr(i);
     THSpinEdit(GetControl(Nom)).OnChange := THSpinEditChange;
     
     // Position :
     Nom := 'POSLIGNE_' + IntToStr(i);
     THSpinEdit(GetControl(Nom)).OnChange := THSpinEditChange;
     
     // Taille :
     Nom := 'TAILLE_' + IntToStr(i);
     THEdit(GetControl(Nom)).OnChange := THEditChange;

     // Obligatoire :
     Nom := 'OBLIGATOIRE_' + IntToStr(i);
     THValComboBox(GetControl(Nom)).OnChange := THValComboBoxChange;
     
     // Utilisé :
     Nom := 'USE_' + IntToStr(i);
     THCheckBox(GetControl(Nom)).OnClick := THCheckBoxClick;

     Inc(i);
     Nom := 'CHAMPS_' + IntToStr(i);
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 23/02/2007
Modifié le ... :   /  /    
Description .. : Desinitilialisation
Mots clefs ... : 
*****************************************************************}
procedure TOM_PARAMADRE.DesinitEvenements;
var
  i : integer;
  Nom : string;
begin
  // Evenement sur la fiche :

  // Evenement sur les lignes :
  // Pour chaque ligne existante :
  i := 0;
  Nom := 'CHAMPS_' + IntToStr(i);
  while (GetControl(Nom) <> nil) do
  begin
     // Pour chaque composant de la ligne:
     // TypeLigne :
     Nom := 'TYPELIGNE_' + IntToStr(i);
     THValComboBox(GetControl(Nom)).OnChange := nil;
     
     //  NumLigne :
     Nom := 'NUMLIGNE_' + IntToStr(i);
     THSpinEdit(GetControl(Nom)).OnChange := nil;
     
     // Position :
     Nom := 'POSLIGNE_' + IntToStr(i);
     THSpinEdit(GetControl(Nom)).OnChange := nil;

     // Taille :
     Nom := 'TAILLE_' + IntToStr(i);
     THEdit(GetControl(Nom)).OnChange := nil;

     // Obligatoire :
     Nom := 'OBLIGATOIRE_' + IntToStr(i);
     THValComboBox(GetControl(Nom)).OnChange := nil;
     
     // Utilisé :
     Nom := 'USE_' + IntToStr(i);
     THCheckBox(GetControl(Nom)).OnClick := nil;

     Inc(i);
     Nom := 'CHAMPS_' + IntToStr(i);
  end;
end;
{***********A.G.L.***********************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 13/02/2007
Modifié le ... :   /  /    
Description .. : Evenement sur le CheckBox
Mots clefs ... : 
*****************************************************************}
procedure TOM_PARAMADRE.THCheckBoxClick(Sender : TObject) ;
var indice : integer;
begin
  gInsertion := THCheckBox(Sender).Checked;
  // On cherche l'indice de la ligne
  Indice := GetIndiceLigne(THCheckBox(Sender).Name);
  LigneEnabled(Indice,gInsertion);

  // On met à jour l'apercu
  ModifLigne(THCheckBox(Sender).Name);
  gInsertion := false;
end ;


{***********A.G.L.***********************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 14/02/2007
Modifié le ... :   /  /    
Description .. : Modification d'un ComboBox
Mots clefs ... : 
*****************************************************************}
procedure TOM_PARAMADRE.THValComboBoxChange(Sender : TOBject);
begin
  // On met à jour l'aperçu
  ModifLigne(THValComboBox(Sender).Name);
end;

{***********A.G.L.***********************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 14/02/2007
Modifié le ... :   /  /    
Description .. : Modification d'un SpinEdit
Mots clefs ... : 
*****************************************************************}
procedure TOM_PARAMADRE.THSpinEditChange(Sender : TOBject);
var
  Indice        : integer;
  Champ         : TChampsAdresse;
  TypeComposant : string;
  Name          : string;
  Sec           : TSectionApercu;
begin
  // On recupere le nom
  Name := THSpinEdit(Sender).Name;
  if Name = 'PAD_CODE' then
  begin
     if (TFFicheListe(ecran).Ta.State = dsInsert) and
        (GetControlText('PAD_CODE') <> IntToStr(GetMaxCode+1)) then
     begin
{$IFDEF EAGLCLIENT}
        THSpinEdit(GetControl('PAD_CODE')).Text := IntToStr(GetMaxCode+1);
        THSpinEdit(GetControl('PAD_CODE')).Value := GetMaxCode+1;
{$ELSE}
        THDBSpinEdit(GetControl('PAD_CODE')).Text := IntToStr(GetMaxCode+1);
        THDBSpinEdit(GetControl('PAD_CODE')).Value := GetMaxCode+1;
{$ENDIF}
     end;
     Exit;
  end;

  // On cherche l'indice de la ligne
  Indice := GetIndiceLigne(Name);
  // On cherche le type
  TypeComposant := GetTypeComposantLigne(Name);
  if GetControlText(Name) = '' then Exit;

  if TypeComposant = 'POSLIGNE' then
  begin
     Champ := TChampsAdresse.Create(Ecran,Indice);
     try
        if ( gInsertion or
           Champ.Replace or
           not(THSpinEdit(Sender).Value > GetMaxPosLigne(Champ))) and
           (THSpinEdit(Sender).Value > 0 ) then
           // On met à jour l'aperçu
           ModifLigne(Name)
        else
           // On remet sa valeur initiale
           MAJComposants(Name,IntToStr(Champ.LastPosLigne));
     finally
        FreeAndNil(Champ);
     end;
  end
  else if TypeComposant = 'NUMLIGNE' then
  begin
     gReplace := true;
     Champ := TChampsAdresse.Create(Ecran,Indice);
     try
        Sec := Apercu.GetSection(Champ.TypeLigne);
        if (THSpinEdit(Sender).Value < 1 ) then
           // On remet sa valeur initiale
           MAJComposants(Name,IntToStr(Champ.LastNumLigne))
        else if Champ.LastNumLigne > Champ.NumLigne then
           // On remonte d'une ligne
           ModifLigne(Name)
        else if gInsertion or Champ.Replace or
           ( not(THSpinEdit(Sender).Value > Sec.Detail.Count + 1) and
           (GetMaxLastPosLigne(Champ) > 1)) then
           // On met à jour l'aperçu
           ModifLigne(Name)
        else
           // On remet sa valeur initiale
           MAJComposants(Name,IntToStr(Champ.LastNumLigne));
     finally
        FreeAndNil(Champ);
     end;
     gReplace := false;
  end
  else
     ModifLigne(Name);

end;

{***********A.G.L.***********************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 14/02/2007
Modifié le ... :   /  /    
Description .. : Modification d'un THEdit
Mots clefs ... : 
*****************************************************************}
procedure TOM_PARAMADRE.THEditChange(Sender : TOBject);
begin
  if Trim(THEdit(Sender).Text) = '' then Exit;
  if not(isnumeric(Trim(THEdit(Sender).Text))) then
  begin
     THEdit(Sender).Text := '';
     Exit;
  end;
  try
    if (StrToInt(Trim(THEdit(Sender).Text)) < 0 ) or
       (StrToInt(Trim(THEdit(Sender).Text)) > 99) then
       MAJComposants(THEdit(Sender).Name,'1');
    // On met à jour l'aperçu
    ModifLigne(THEdit(Sender).Name);
  except
    MAJComposants(THEdit(Sender).Name,'1');;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 20/02/2007
Modifié le ... :   /  /    
Description .. : Permet de mettre à jour un composant si necessaire.
Mots clefs ... : 
*****************************************************************}
procedure TOM_PARAMADRE.MAJComposants(Name : string;Valeur : string);
var
  TypeComposant : string;
begin
  // On cherche le nom
  TypeComposant := GetTypeComposantLigne(Name);

  if (TypeComposant = 'NUMLIGNE') or (TypeComposant = 'POSLIGNE') then
     if StrToInt(Valeur) < 1 then
        Exit;
  if GetControlText(Name) <> Valeur then
     SetControlText(Name,Valeur);
end;

{***********A.G.L.***********************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 20/02/2007
Modifié le ... :   /  /    
Description .. : Mise à jour de la TOB Champs pour sauvegarder les valeurs
Mots clefs ... : 
*****************************************************************}
procedure TOM_PARAMADRE.MAJTobSAV(Name : string) ;
var
  Indice  : integer;
  TypeComposant : string;
begin
  if Trim(GetControlText(Name)) = '' then Exit;
  // On cherche l'indice de la ligne
  Indice := GetIndiceLigne(Name);
  // On cherche le nom
  TypeComposant := GetTypeComposantLigne(Name);
  // On met à jour la TOB
  Champs.Detail[Indice].PutValue('RAD_' + TypeComposant,GetControlText(Name));
end;

{***********A.G.L.***********************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 21/02/2007
Modifié le ... :   /  /    
Description .. : Permet de mettre à jour les champs LastPosLigne, 
Suite ........ : LastNumLigne et LastTypeLigne
Mots clefs ... : 
*****************************************************************}
procedure TOM_PARAMADRE.MAJLastChamp(Champ : TChampsAdresse);
begin
  MAJComposants('LASTTYPELIGNE_'+IntToStr(Champ.Indice),Champ.TypeLigne);
  MAJComposants('LASTPOSLIGNE_'+IntToStr(Champ.Indice),IntToStr(Champ.PosLigne));
  MAJComposants('LASTNUMLIGNE_'+IntToStr(Champ.Indice),IntToStr(Champ.NumLigne));
end;


{***********A.G.L.***********************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 20/02/2007
Modifié le ... : 20/02/2007
Description .. : Met le champ qui doit bouger à jour
Mots clefs ... : 
*****************************************************************}
procedure TOM_PARAMADRE.ComposantChangePlace(var Champ : TChampsAdresse ; Name : string);
var
  i             : integer;
  cherche       : TOB;
  Indice        : integer;
  TypeComposant : string;
begin
  // On cherche l'indice de la ligne
  Indice := GetIndiceLigne(Name);
  // On cherche le nom
  TypeComposant := GetTypeComposantLigne(Name);

  if not(Champ.Utilise) then Exit;
  // On recherche un champ du meme type ligne qui se deplacera
  for i := 0 to Champs.Detail.Count - 1 do
  begin
     cherche := Champs.Detail[i];
     if i = Indice then Continue;
     if ( cherche.GetValue('RAD_TYPELIGNE') = Champ.TypeLigne ) and
        ( cherche.GetValue('RAD_NUMLIGNE')  = Champ.NumLigne  ) and
        ( cherche.GetValue('RAD_POSLIGNE')  = Champ.PosLigne  ) and
        ( cherche.GetValue('RAD_USE')       = 'X'       ) then
     begin
        // CLICK SUR UTILISE
        if (Champ.LastPosLigne = Champ.PosLigne) and
           (Champ.LastNumLigne = Champ.NumLigne) and
           (Champ.LastTypeLigne = Champ.TypeLigne) then
        begin
           // On est en train d'ajouter un champ il faut donc le deplacer vers la droite
           SetRecursif;
           MAJComposants('POSLIGNE_'+IntToStr(i),IntToStr(Champ.PosLigne + 1));
           UnsetRecursif;
        end
        // CHANGEMENT POS LIGNE
        else if Champ.LastPosLigne < Champ.PosLigne then
        begin
           SetRecursif;
           if (Champ.PosLigne - Champ.LastPosLigne) > 1 then
              gDeplace := true;
           if (not(gReplace) or (GetChampLigne(Champ) = 2)) and (Length(gRecursif) < 2) then
              // On le deplace vers la gauche
              MAJComposants('POSLIGNE_'+IntToStr(i),IntToStr(Champ.PosLigne - 1))
           else
              MAJComposants('POSLIGNE_'+IntToStr(i),IntToStr(Champ.PosLigne + 1));
           if (Champ.PosLigne - Champ.LastPosLigne) > 1 then
              gDeplace := false;
           UnsetRecursif;
        end
        else if Champ.LastPosLigne > Champ.PosLigne then
        begin       
           SetRecursif;   
           if (Champ.LastPosLigne - Champ.PosLigne) > 1 then
              gDeplace := true;
           // On veux deplacer le champ vers la droite
           if not(gReplace) and not(gDeplace) then
              MAJComposants('REPLACE_'+IntToStr(i),'X');
           if gDeplace and (Length(gRecursif) > 1) then
              MAJComposants('POSLIGNE_'+IntToStr(i),IntToStr(Champ.PosLigne - 1))
           else
              MAJComposants('POSLIGNE_'+IntToStr(i),IntToStr(Champ.PosLigne + 1));
           if not(gReplace) then
              MAJComposants('REPLACE_'+IntToStr(i),'-');
           if (Champ.LastPosLigne - Champ.PosLigne) > 1 then
              gDeplace := false;
           UnsetRecursif;
        end
        // CHANGEMENT NUM LIGNE
        else if Champ.LastNumLigne <> Champ.NumLigne then
        begin
           SetRecursif;
           // On veux deplacer le champ vers la droite
           MAJComposants('POSLIGNE_'+IntToStr(i),IntToStr(Champ.PosLigne + 1));
           UnsetRecursif;
        end
        // CHANGEMENT TYPE LIGNE
        else
        begin                                        
           SetRecursif;
           // On veux deplacer le champ vers la droite
           MAJComposants('POSLIGNE_'+IntToStr(i),IntToStr(Champ.PosLigne + 1));  
           UnsetRecursif;
        end;
        Exit;
     end;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 20/02/2007
Modifié le ... :   /  /    
Description .. : Permet d'ajouter le champs dans l'aperçu
Mots clefs ... : 
*****************************************************************}
function TOM_PARAMADRE.AjouterChamp(Champ : TChampsAdresse) : boolean;
var
  Section   : TSectionApercu;
  NumLigne  : integer;
  PosLigne  : integer;
  Valeur    : string;
  Ligne     : TOB;
begin
  Result := true;
  // On récupere la bonne section
  Section := Apercu.GetSection(Champ.TypeLigne);
     
  // On recupere la bonne ligne
  NumLigne := Champ.NumLigne;
  Ligne := Section.GetLigne(NumLigne);

  // On ne doit pas ajouter le composant sur cette ligne
  // On le remet à jour avec la bonne valeur
  // Et on sort !
  if NumLigne <> Champ.NumLigne then
  begin
     MAJComposants('DEPLACE_' + IntToStr(Champ.Indice), 'X');
     MAJComposants('NUMLIGNE_'+ IntToStr(Champ.Indice),IntToStr(NumLigne));
     MAJComposants('DEPLACE_' + IntToStr(Champ.Indice), '-');
     Result := false;
     Exit;
  end;

  // On met à jour la TOB qui gere l'apercu
  if (Champ.PosLigne > GetChampLigne(Champ)) then
  begin
     // On ne doit remettre le composant à jour avec la bonne valeur
     MAJComposants('DEPLACE_' + IntToStr(Champ.Indice), 'X');
     MAJComposants('POSLIGNE_'+IntToStr(Champ.Indice),IntToStr(GetChampLigne(Champ)));
     MAJComposants('DEPLACE_' + IntToStr(Champ.Indice), '-');
     Result := false;
     Exit;
  end;
  Valeur := Champs.Detail[Champ.Indice].GetValue('RAD_ABREGE');
  PosLigne := TLigneTOB(Ligne).UpdateChamp(Champ.PosLigne,Valeur);
  if PosLigne <> Champ.PosLigne then
  begin
     // On ne doit remettre le composant à jour avec la bonne valeur
     MAJComposants('DEPLACE_' + IntToStr(Champ.Indice), 'X');
     MAJComposants('POSLIGNE_'+IntToStr(Champ.Indice),IntToStr(PosLigne));
     MAJComposants('DEPLACE_' + IntToStr(Champ.Indice), '-');
     Result := false;
     Exit;
  end;

  // On met à jour l'apercu
  Apercu.UpdateLigApercu(Champ.TypeLigne,Champ.NumLigne);
end;

{***********A.G.L.***********************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 20/02/2007
Modifié le ... :   /  /    
Description .. : Permet de supprimer un champ dans l'apercu
Mots clefs ... : 
*****************************************************************}
procedure TOM_PARAMADRE.SupprimerChamp(Champ : TChampsAdresse);
var
  Section     : TSectionApercu;
  NumLigne    : integer;
  MaxPosLigne : integer;
  Ligne       : TOB;
  Cherche     : TOB;
  i,j         : integer;
  replace     : boolean;
  isBreaked   : boolean;
begin
  Replace := False;
  isBreaked := False;
  MaxPosLigne := 0;
  // On récupere la bonne section
  Section := Apercu.GetSection(Champ.LastTypeLigne);

  // On recupere la bonne ligne
  NumLigne := Champ.LastNumLigne;
  Ligne := Section.GetLigne(NumLigne);

  // La Ligne n'existe pas on sort
  if NumLigne <> Champ.LastNumLigne then Exit;
  TLigneTOB(Ligne).DelChamp(Champ.LastPosLigne);

  if Ligne.GetChampCount(ttcSup) > 0 then
  begin
     // On doit mettre à jour tous les composants qui doivent glisser à gauche
     for i := Champ.LastPosLigne to GetMaxPosChamp(Ligne) do
     begin
        for j := 0 to Champs.Detail.Count - 1 do
        begin
           cherche := Champs.Detail[j];
           if j = Champ.Indice then Continue;
           if ( cherche.GetValue('RAD_TYPELIGNE') = Champ.LastTypeLigne ) and
              ( cherche.GetValue('RAD_NUMLIGNE')  = Champ.LastNumLigne  ) and
              ( cherche.GetValue('RAD_POSLIGNE')  = i + 1               ) and
              ( cherche.GetValue('RAD_USE')       = 'X'                 ) then
           begin
              MAJComposants('REPLACE_' + IntToStr(j), 'X');
              MAJComposants('POSLIGNE_' + IntToStr(j), IntToStr(i));
              MAJComposants('REPLACE_' + IntToStr(j), '-');
              Replace := true;
              if MaxPosLigne <  cherche.GetValue('RAD_POSLIGNE') then
                 MaxPosLigne := cherche.GetValue('RAD_POSLIGNE');
              isBreaked := true;
              break;
           end;
        end;
        if isBreaked then break;
     end;
     if not(Replace) then
     begin
        // On Supprime le champ de la TOB
        TLigneTOB(Ligne).DelChamp(Champ.LastPosLigne);
     end;
     if (TLigneTOB(Ligne).GetChampCount(ttcSup) - 999) > GetChampLastLigne(Champ) then
     begin
        // On Supprime le dernier champ de la ligne
        TLigneTOB(Ligne).DelChamp(MaxPosLigne);
     end;
     // On met à jour l'apercu
     Apercu.UpdateLigApercu(Champ.LastTypeLigne,Champ.LastNumLigne);
  end
  else
  begin
     if (Champ.Utilise and
        (Champ.NumLigne = Champ.LastNumLigne) and
        (Champ.TypeLigne = Champ.LastTypeLigne)) then Exit;
     // On doit mettre à jour tous les composants qui doivent "remonter"
     for i := (Champ.LastNumLigne + 1)  to Section.GetMaxLigne do
     begin
        for j := 0 to Champs.Detail.Count - 1 do
        begin
           cherche := Champs.Detail[j];
           if j = Champ.Indice then Continue;
           if ( cherche.GetValue('RAD_TYPELIGNE') = Champ.LastTypeLigne ) and
              ( cherche.GetValue('RAD_NUMLIGNE')  = i                   ) and
              ( cherche.GetValue('RAD_USE')       = 'X'                 ) then
           begin
              // On doit remplacer les champs :
              MAJComposants('REPLACE_' + IntToStr(j), 'X');
              MAJComposants('NUMLIGNE_' + IntToStr(j), IntToStr(i - 1));
              MAJComposants('REPLACE_' + IntToStr(j), '-');
              if cherche.GetValue('RAD_POSLIGNE') = IntToStr(Champ.PosLigne) then
                 replace := true;
              isBreaked := true;
              break;
           end;
        end;  
        if isBreaked then break;
     end;
     if Champs.Detail[Champ.Indice].GetValue('RAD_ABREGE') = Ligne.GetValue(IntToStr(Champ.LastPosLigne)) then
     begin
        // Le Champ n'a pas été remplacé.
        // On Supprime le champ de la TOB
        TLigneTOB(Ligne).DelChamp(Champ.LastPosLigne);
     end;
     if not(Replace) then
     begin
        // On doit supprimer la ligne
        Apercu.DelLigApercu(Champ.LastTypeLigne,Champ.LastNumLigne);
     end;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 13/02/2007
Modifié le ... : 14/02/2007
Description .. : Permet de mettre à jour l'aperçu apres modification d'un 
Suite ........ : composant de la ligne
Mots clefs ... : 
*****************************************************************}
procedure TOM_PARAMADRE.ModifLigne(Name : string) ;
var
  Indice        : integer;
  LastNumLigne  : integer;
  TypeComposant : string;
  Champ         : TChampsAdresse;
  LastLigne     : TOB;

begin
  if not(gLoad) then
  begin
     if not(gSave) then
     begin
        gSave := true;
        if not(gNew) then
           TFFicheListe(Ecran).Ta.Edit
     end;
  end;

  // On cherche l'indice de la ligne
  Indice := GetIndiceLigne(Name);
  // On cherche le nom
  TypeComposant := GetTypeComposantLigne(Name);

  // Gestion du chargement
  if gLoad then
  begin
     if (TypeComposant = 'USE') then
     begin
        if not(THCheckBox(GetControl(Name)).Checked) then Exit;
     end
     else Exit;
  end;

  // MAJ de la TOB
  MAJTobSAV(Name);

  if (TypeComposant = 'TAILLE'        ) or
     (TypeComposant = 'OBLIGATOIRE'   ) or
     (TypeComposant = 'LASTTYPELIGNE' ) or
     (TypeComposant = 'LASTPOSLIGNE'  ) or
     (TypeComposant = 'LASTNUMLIGNE'  ) then
     Exit; // Aucune modif de l'aperçu.

  // On récupere les Valeurs dont on a besoin
  Champ := TChampsAdresse.Create(Ecran,Indice);
  try
     LastNumLigne := Champ.LastNumLigne;
     LastLigne := Apercu.GetSection(Champ.LastTypeLigne).GetLigne(LastNumLigne);
     // On supprime le champ de son ancienne place
     if ( Champ.replace and LastLigne.FieldExists(IntToStr(Champ.LastPosLigne)) and
        (( Champ.PosLigne <> Champ.LastPosLigne) and (gRecursif = '')           or
        ( Champ.NumLigne <> Champ.LastNumLigne) and (gRecursif = '')))          or // Cas position <> numero champ TOB
        ( not(Champ.replace) and (TypeComposant <> 'POSLIGNE'))                 then // Cas on remplace
     begin
        if (TypeComposant <> 'USE') or not(Champ.Utilise) then
        begin
           if not(Champ.Deplace) then
           begin
              SupprimerChamp(Champ);
              gInsertion := true;
           end;
        end;
     end;

     // On cherche les composants à mettre à jour
     if not(Champ.replace) then ComposantChangePlace(Champ, Name);

     // On rajoute le champ à sa nouvelle place
     if Champ.Utilise then
        if AjouterChamp(Champ) then
           // On met a jour les champs last
           MAJLastChamp(Champ);

     gInsertion := false;
  finally
     FreeAndNil(Champ);
  end;
end ;


{***********A.G.L.***********************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 14/02/2007
Modifié le ... :   /  /
Description .. : Fermeture de la fiche
Mots clefs ... : 
*****************************************************************}
procedure TOM_PARAMADRE.OnClose ;
begin
  Inherited ;
  DechargeVar;
end ;

procedure TOM_PARAMADRE.DechargeVar;
begin
  if assigned(Champs) then
     FreeAndNil(Champs);
  if assigned(Apercu) then
  begin
     Apercu.Apercu.Clear;
     FreeAndNil(Apercu);
  end;
end ;

procedure TOM_PARAMADRE.OnCancelRecord ;
begin
  Inherited ;
  gSave := false;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 14/02/2007
Modifié le ... :   /  /    
Description .. : Charge les champs de la table Reference et en fonction 
Suite ........ : du paramétrage sur la fiche 
Mots clefs ... : 
*****************************************************************}
procedure TOM_PARAMADRE.ChargeChampsFiche ;
var
  i        : integer;
  NomChamp : string;
begin
  for i := 0 to Champs.Detail.Count - 1 do
  begin
     NomChamp := 'CHAMPS_' + IntToStr(i);

     // Création de la ligne
     if Ecran.FindComponent(NomChamp) = nil then
       CreateChampsLigne(i);

     // Renseignement des champs de la ligne
     RenseigneChampsLigne(Champs.Detail[i],i);

     // En fonction du CheckBox on va griser ou non la ligne :
     THCheckBoxClick(GetControl('USE_'+IntToStr(i)));
  end;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 14/02/2007
Modifié le ... :   /  /    
Description .. : Crée l'ensemble des composants d'une ligne
Mots clefs ... : 
*****************************************************************}
procedure TOM_PARAMADRE.CreateChampsLigne(Indice : integer) ;
var
  NomChamp : string;
begin
  // Le Label
  NomChamp := 'CHAMPS_';
  With TLabel(CreationCopie(GetControl(NomChamp + IntToStr(Indice - 1)),NomChamp + IntToStr(Indice),'THLabel')) Do
  Begin
    Top := Top+INTERLIGNE;
  End;
  // Type Ligne
  NomChamp := 'TYPELIGNE_';
  With THValComboBox(CreationCopie(GetControl(NomChamp + IntToStr(Indice - 1)),NomChamp + IntToStr(Indice),'THValComboBox')) Do
  Begin
    Top := Top+INTERLIGNE;
  End;
  // LAST Type Ligne
  NomChamp := 'LASTTYPELIGNE_';
  With THValComboBox(CreationCopie(GetControl(NomChamp + IntToStr(Indice - 1)),NomChamp + IntToStr(Indice),'THValComboBox')) Do
  Begin
    Top := Top+INTERLIGNE;
  End;
  // Num Ligne
  NomChamp := 'NUMLIGNE_';
  With THSpinEdit(CreationCopie(GetControl(NomChamp + IntToStr(Indice - 1)),NomChamp + IntToStr(Indice),'THSpinEdit')) Do
  Begin
    Top := Top+INTERLIGNE;
  End;  
  // LAST Num Ligne
  NomChamp := 'LASTNUMLIGNE_';
  With THSpinEdit(CreationCopie(GetControl(NomChamp + IntToStr(Indice - 1)),NomChamp + IntToStr(Indice),'THSpinEdit')) Do
  Begin
    Top := Top+INTERLIGNE;
  End;
  // Pos Ligne
  NomChamp := 'POSLIGNE_';
  With THSpinEdit(CreationCopie(GetControl(NomChamp + IntToStr(Indice - 1)),NomChamp + IntToStr(Indice),'THSpinEdit')) Do
  Begin
    Top := Top+INTERLIGNE;
  End;
  // LAST Pos Ligne
  NomChamp := 'LASTPOSLIGNE_';
  With THSpinEdit(CreationCopie(GetControl(NomChamp + IntToStr(Indice - 1)),NomChamp + IntToStr(Indice),'THSpinEdit')) Do
  Begin
    Top := Top+INTERLIGNE;
  End;
  // Taille
  NomChamp := 'TAILLE_';
  With THEdit(CreationCopie(GetControl(NomChamp + IntToStr(Indice - 1)),NomChamp + IntToStr(Indice),'THEdit')) Do
  Begin
    Top := Top+INTERLIGNE;
  End;
  // Obligatoire
  NomChamp := 'OBLIGATOIRE_';
  With THValComboBox(CreationCopie(GetControl(NomChamp + IntToStr(Indice - 1)),NomChamp + IntToStr(Indice),'THValComboBox')) Do
  Begin
    Top := Top+INTERLIGNE;
  End;
  // Utilisé
  NomChamp := 'USE_';
  With THCheckBox(CreationCopie(GetControl(NomChamp + IntToStr(Indice - 1)),NomChamp + IntToStr(Indice),'THCheckBox')) Do
  Begin
    Top := Top+INTERLIGNE;
  End;
  // Replace
  NomChamp := 'REPLACE_';
  With THCheckBox(CreationCopie(GetControl(NomChamp + IntToStr(Indice - 1)),NomChamp + IntToStr(Indice),'THCheckBox')) Do
  Begin
    Top := Top+INTERLIGNE;
  End;  
  // Deplace
  NomChamp := 'DEPLACE_';
  With THCheckBox(CreationCopie(GetControl(NomChamp + IntToStr(Indice - 1)),NomChamp + IntToStr(Indice),'THCheckBox')) Do
  Begin
    Top := Top+INTERLIGNE;
  End;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 14/02/2007
Modifié le ... :   /  /    
Description .. : Modifie les valeurs des composants d'une ligne
Mots clefs ... : 
*****************************************************************}
procedure TOM_PARAMADRE.RenseigneChampsLigne(Ligne : TOB ;Indice : integer) ;
var
  NomChamp : string;
begin
  // Le Label
  NomChamp := 'CHAMPS_' + IntToStr(Indice);
  SetControlText(NomChamp,Ligne.GetValue('RAD_LIBELLE'));

  // Type Ligne
  NomChamp := 'TYPELIGNE_' + IntToStr(Indice);
  SetControlText(NomChamp,Ligne.GetValue('RAD_TYPELIGNE'));
  
  // LAST Type Ligne
  NomChamp := 'LASTTYPELIGNE_' + IntToStr(Indice);
  SetControlText(NomChamp,Ligne.GetValue('RAD_TYPELIGNE'));

  // Num Ligne
  NomChamp := 'NUMLIGNE_' + IntToStr(Indice);
  SetControlText(NomChamp,Ligne.GetValue('RAD_NUMLIGNE'));
  
  // LAST Num Ligne
  NomChamp := 'LASTNUMLIGNE_' + IntToStr(Indice);
  SetControlText(NomChamp,Ligne.GetValue('RAD_NUMLIGNE'));

  // Pos Ligne
  NomChamp := 'POSLIGNE_' + IntToStr(Indice);
  SetControlText(NomChamp,Ligne.GetValue('RAD_POSLIGNE'));

  // LAST Pos Ligne
  NomChamp := 'LASTPOSLIGNE_' + IntToStr(Indice);
  SetControlText(NomChamp,Ligne.GetValue('RAD_POSLIGNE'));

  // Taille
  NomChamp := 'TAILLE_' + IntToStr(Indice);
  SetControlText(NomChamp,Ligne.GetValue('RAD_TAILLE'));

  // Obligatoire
  NomChamp := 'OBLIGATOIRE_' + IntToStr(Indice);
  SetControlText(NomChamp,Ligne.GetValue('RAD_OBLIGATOIRE'));

  // Utilisé
  NomChamp := 'USE_' + IntToStr(Indice);
  SetControlText(NomChamp,Ligne.GetValue('RAD_USE'));

end ;

{***********A.G.L.***********************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 13/02/2007
Modifié le ... :   /  /    
Description .. : Procedure permettant de griser ou dégriser tous les 
Suite ........ : elements d'une ligne.
Mots clefs ... : 
*****************************************************************}
procedure TOM_PARAMADRE.LigneEnabled(Indice : integer; Value : boolean) ;
var
  NomChamp : string;
begin                  
  // Le Label
  NomChamp := 'CHAMPS_' + IntToStr(Indice);
  SetControlEnabled(NomChamp,Value);

  // Type Ligne
  NomChamp := 'TYPELIGNE_' + IntToStr(Indice);
  SetControlEnabled(NomChamp,Value);

  // Num Ligne
  NomChamp := 'NUMLIGNE_' + IntToStr(Indice);
  SetControlEnabled(NomChamp,Value);

  // Pos Ligne
  NomChamp := 'POSLIGNE_' + IntToStr(Indice);
  SetControlEnabled(NomChamp,Value);

  // Taille
  NomChamp := 'TAILLE_' + IntToStr(Indice);
  SetControlEnabled(NomChamp,Value);

  // Obligatoire
  NomChamp := 'OBLIGATOIRE_' + IntToStr(Indice);
  SetControlEnabled(NomChamp,Value);
end;

{***********A.G.L.***********************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 21/02/2007
Modifié le ... :   /  /    
Description .. : Retourne la position maximum d'un champ sur la ligne
Mots clefs ... : 
*****************************************************************}
function TOM_PARAMADRE.GetMaxPosLigne(Champ : TChampsAdresse) : integer;
var
  i : integer;
  cherche : TOB;
begin
  Result := 0;
  for i := 0 to Champs.Detail.Count - 1 do
  begin
     cherche := Champs.Detail[i];
      if ( cherche.GetValue('RAD_TYPELIGNE') = Champ.TypeLigne ) and
         ( cherche.GetValue('RAD_NUMLIGNE')  = Champ.NumLigne  ) and
         ( cherche.GetValue('RAD_USE')       = 'X'             ) then
         Inc(Result);
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 21/02/2007
Modifié le ... :   /  /    
Description .. : Retourne la position maximum d'un champ sur la ligne
Mots clefs ... : 
*****************************************************************}
function TOM_PARAMADRE.GetMaxLastPosLigne(Champ : TChampsAdresse) : integer;
var
  i : integer;
  cherche : TOB;
begin
  Result := 0;
  for i := 0 to Champs.Detail.Count - 1 do
  begin
     cherche := Champs.Detail[i];
     if ( cherche.GetValue('RAD_TYPELIGNE') = Champ.TypeLigne     ) and
        ( cherche.GetValue('RAD_NUMLIGNE')  = Champ.LastNumLigne  ) and
        ( cherche.GetValue('RAD_POSLIGNE')  > Result              ) and
        ( cherche.GetValue('RAD_USE')       = 'X'                 ) then
        Result := cherche.GetValue('RAD_POSLIGNE')
  end;
end;

 {***********A.G.L.***********************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 21/02/2007
Modifié le ... :   /  /
Description .. : Retourne la position maximum d'un champ sur la ligne
Mots clefs ... :
*****************************************************************}
function TOM_PARAMADRE.GetChampLigne(Champ : TChampsAdresse) : integer;
var
  i : integer;
  cherche : TOB;
begin
  Result := 0;
  for i := 0 to Champs.Detail.Count - 1 do
  begin
     cherche := Champs.Detail[i];
     if ( cherche.GetValue('RAD_TYPELIGNE') = Champ.TypeLigne ) and
        ( cherche.GetValue('RAD_NUMLIGNE')  = Champ.NumLigne  ) and
        ( cherche.GetValue('RAD_USE')       = 'X'             ) then
        Inc(Result);
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 21/02/2007
Modifié le ... :   /  /
Description .. : Retourne la position maximum d'un champ sur la ligne
Mots clefs ... :
*****************************************************************}
function TOM_PARAMADRE.GetChampLastLigne(Champ : TChampsAdresse) : integer;
var
  i : integer;
  cherche : TOB;
begin
  Result := 0;
  for i := 0 to Champs.Detail.Count - 1 do
  begin
     cherche := Champs.Detail[i];
     if ( cherche.GetValue('RAD_TYPELIGNE') = Champ.TypeLigne     ) and
        ( cherche.GetValue('RAD_NUMLIGNE')  = Champ.LastNumLigne  ) and
        ( cherche.GetValue('RAD_USE')       = 'X'                 ) then
        Inc(Result);
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 21/02/2007
Modifié le ... :   /  /
Description .. : Retourne la position maximum d'un champ sur la ligne
Mots clefs ... :
*****************************************************************}
procedure TOM_PARAMADRE.SetRecursif;
begin
   gRecursif := gRecursif + '1';
end;

{***********A.G.L.***********************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 21/02/2007
Modifié le ... :   /  /
Description .. : Retourne la position maximum d'un champ sur la ligne
Mots clefs ... :
*****************************************************************}
procedure TOM_PARAMADRE.UnsetRecursif;
begin
   SetLength(gRecursif,Length(gRecursif)-1);
end;

Initialization
  registerclasses ( [ TOM_PARAMADRE ] ) ;
end.
