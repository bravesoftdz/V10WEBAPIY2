{***********UNITE*************************************************
Auteur  ...... :  NA
Créé le ...... : 13/02/2006
Modifié le ... :   /  /
Description .. : Source TOM de la TABLE : CODEPOPULATION (CODEPOPULATION)
Mots clefs ... : TOM;CODEPOPULATION
*****************************************************************}
{
PT1 19/02/07 NA Contrôle code population non existant en création
PT2 22/02/07 NA Contrôle transversalité : les critères sont communs si prédéfini = CEG ou STD, ou dossiers si préd = DOS
PT3 21/05/07 NA Contrôle que les critères sélectionnés sont renseignés dans la fiche salarié
PT4 28/06/07 NA Affichage des types de pop PRE et FOR uniquement si module sérialisé
PT5 23/08/07 NA FQ 14630 Ne pas afficher les critères "organisation", "combo", code stat
PT6 23/08/07 NA FQ 14389 Afficher le critère non renseigné dans la liste d'anomalie
PT7 21/12/07 GGU FQ 14684 Proposer une Saisie groupée pour les salariés dont les critères de population sont non renseignés.
PT9 11/03/08 GGU FQ 15297 ne pas permettre de minuscule dans les codes alphanumériques
PT10 10/07/08 NA FQ 15622 Affichage du type IDR si le module est sérialisé
}


unit CODEPOPULATION_TOM;

interface

uses StdCtrls,
  Controls,
  Classes,
 {$IFNDEF EAGLCLIENT}
  db,
 {$IFNDEF DBXPRESS}dbTables, {$ELSE}uDbxDataSet, {$ENDIF}
  Fiche,
  {$ELSE}
  eFiche,
 {$ENDIF}
  sysutils,
  ComCtrls,
  HCtrls,
  HMsgBox,
  PgOutils,
  PgOutils2,
  UTOM,
  Utob,
  UTobDebug,
  Hent1,
{$IFDEF EAGLCLIENT} //PT7
  MaineAGL,         //PT7
{$ELSE}             //PT7
  FE_main,          //PT7
{$ENDIF}            //PT7
  HDB;

type
  TOM_CODEPOPULATION = class(TOM)
    procedure OnNewRecord; override;
    procedure OnDeleteRecord; override;
    procedure OnUpdateRecord; override;
    procedure OnAfterUpdateRecord; override;
    procedure OnAfterDeleteRecord; override;
    procedure OnLoadRecord; override;
    procedure OnChangeField(F: TField); override;
    procedure OnArgument(S: string); override;
    procedure OnClose; override;
    procedure OnCancelRecord; override;
  private
    LectureSeule, CEG, STD, DOS : boolean;
    GblTypePop, LeTitre: string;
    bo, verif : boolean;

    listeCritereIdent, listeCodesSal : String; //PT7

    procedure Initfiche;
    procedure Initident;
  //  procedure Ordrepopulation (sender : TObject);
    procedure ChangeNbIdent (sender : TObject);
    procedure Changetypepop (sender : TObject);
    procedure controlecodepop (sender : TObject);     // pt1
  //  Function Critererenseigne(ListeCritere : String) : integer;// pt3  // pt6
    Function Critererenseigne(ListeCritere, listenomcritere : String) : integer; // pt6 
    Procedure OnValide1Click (Sender : TObject);
    Procedure OnSaisieSalarieClick(Sender : TObject);
  end;

implementation

Uses
  htb97, Entpaie;

procedure TOM_CODEPOPULATION.OnNewRecord;
begin
  inherited;
// SetControlEnabled('BSPOPULATION',False);
if (CEG = TRUE) then
    SetField('PPO_PREDEFINI', 'CEG')
  else
    SetField('PPO_PREDEFINI', 'DOS');
    setfocuscontrol('PPO_PREDEFINI');
end;

{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 07/03/2006
Modifié le ... :   /  /    
Description .. : Suppression code population
Mots clefs ... : 
*****************************************************************}
procedure TOM_CODEPOPULATION.OnDeleteRecord;
var
SQL1, COdepop,Sql2, Sql3,sql4, sql5, lib, Letitre1 : string;
OKSupp ,suppop: Boolean;
// Contrôle si la suppression est possible : pas de simulation existante pour ce code pop
begin
  suppop:= true;
  Codepop := Getfield('PPO_CODEPOP');
  lib := Getfield('PPO_LIBELLE');
  Letitre1 := Letitre + ' ' + codepop + ' ' + lib;

  inherited;
  Sql1 := 'Select PSR_CODESIMUL from RESULTSIMUL where PSR_CODEPOP = "'+Codepop+'"';
  if ExisteSql(Sql1) then
  begin
  PGIBOX('Suppression interdite : Simulations existantes', Ecran.Caption);
  Lasterror := 1;
  exit;
  end;

  OKsupp := true;

  Sql2 := 'Select PPR_CODEPOP from REGLESPOP where PPR_CODEPOP = "'+Codepop+'"';
  Sql3 := 'Select PPC_CODEPOP from ORDREPOPULATION where PPC_CODEPOP ="'+Codepop+'"';
  Sql4 := 'Select PHL_CODEPOP from PGHABILITATIONS where PHL_CODEPOP = "'+Codepop+'"';
  sql5 := 'Select PPC_POPULATION FROM ORDREPOPULATION, CODEPOPULATION , PGPARAMETRESASSOC WHERE' +
  ' PPC_CODEPOP = PPO_CODEPOP AND PGO_CODEASSOCIE= PPC_POPULATION AND PGO_TYPEPARAMETRE = "POP"';

  if ExisteSql(Sql3) then                   // si détail population existe
  begin
    if ExisteSql(Sql2) then              // si régles associées
    begin
      PGIBOX('Suppression impossible: il existe des populations et des règles IDR associées.', LeTitre1);
      suppop := false;
      lasterror := 1;
      exit;
    end;

     if suppop = true then
     begin
      if Existesql(sql4) then
      begin
        PGIBOX('Suppression impossible: il existe des populations et des habilitations associées.', LeTitre1);
        suppop := false;
        lasterror := 1;
        exit;
      end;
    end;

    if suppop = true then
    begin
      if Existesql(sql5) then
      begin
        PGIBOX('Suppression impossible: il existe des populations et des paramètres de paie associés.', LeTitre1);
        suppop := false;
        lasterror := 1;
        exit;
      end;
    end;

    if suppop = true  then
    begin
        Case PGIAskCancel('Attention : il existe des populations. Voulez vous les supprimer ?', LeTitre1) of
        mrYes   : ExecuteSql('Delete from ORDREPOPULATION where PPC_CODEPOP ="'+Codepop+'"');
        mrCancel: Oksupp := false;
        mrNo    : Oksupp := false;
        end;
    end;

    if Oksupp = false then
    begin
       lasterror := 1;
    end;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 07/03/2006
Modifié le ... :   /  /
Description .. : Contrôle de cohérence avant mise à jour
Mots clefs ... :
*****************************************************************}
procedure TOM_CODEPOPULATION.OnUpdateRecord;
{var
  nbident : integer;
  typepop, predefini : string;
  ident1 : integer;
  ident2 : integer;
  ident3 : integer;
  ident4 : integer;


i : integer; // nbident,
critere, sti, numident, listecritere, nomchamp :string;
Q1 : TQuery; }

begin
  inherited;
  SetField('PPO_CODEPOP', UpperCase(GetField('PPO_CODEPOP'))); //PT9
  if not verif then
  begin
    //Permet d'effectuer les contrôles blocants en cas de clique sur la croix et d'acceptation des modifications
    OnValide1Click(self);
    Lasterror := 1;
    exit;
  end;
  verif := False;
{  nbident := Getfield('PPO_NBIDENT');
  typepop := GetControlText('PPO_TYPEPOP');
  predefini := GetControlText('PPO_PREDEFINI');
  ident1 :=  strtoint(Getfield('PPO_IDENT1'));
  ident2 :=  strtoint(Getfield('PPO_IDENT2'));
  ident3 :=  strtoint(Getfield('PPO_IDENT3'));
  ident4 :=  strtoint(Getfield('PPO_IDENT4'));

  if predefini = 'DOS' then setfield('PPO_NODOSSIER', PgrendNodossier())  else
  setfield('PPO_NODOSSIER', '000000');

// Contrôle du type de population
if typepop = '' then begin
  PGIBOX('Vous devez renseigner: Type de population.', Letitre);
  setfocuscontrol('PPO_TYPEPOP');
  Lasterror := 1;
  exit;
end;

//  Contrôle des identifiants

  if nbident = 0 then begin
  PGIBOX('Vous devez renseigner: Nombre de critères.', Letitre);
  setfocuscontrol('PPO_NBIDENT');
  Lasterror := 1;
  exit;
  end;

  if (nbident = 1) and  (ident1 = 0) then begin
  PGIBOX('Vous devez renseigner : Critère 1.', Letitre);
  setfocuscontrol('PPO_IDENT1');
  Lasterror := 1;
  exit;
  end;

  //  si Nombre identifiant = 2//
  if (nbident = 2) then begin

      if ((ident1 = 0) or (ident2= 0)) then
      begin
      PGIBOX('Vous devez renseigner : Critère 1 et Critère 2.', Letitre);
      setfocuscontrol('PPO_IDENT1');
      if (ident1 <> 0) then setfocuscontrol('PPO_IDENT2');
      Lasterror := 1;
      exit;
      end
      else
        if ident1 = ident2 then begin
        PGIBOX('Les critères doivent être différents.', Letitre);
        setfocuscontrol('PPO_IDENT1');
        Lasterror := 1;
        exit;
      end;
  end;

  // si Nombre identifiant = 3 //
  if (nbident = 3) then begin

      if  ((ident1 = 0) or (ident2= 0) or (ident3 = 0)) then begin
      PGIBOX('Vous devez renseigner : Critère 1, Critère 2 et Critère 3.', Letitre);
      setfocuscontrol('PPO_IDENT1');
      if (ident3 = 0) then setfocuscontrol('PPO_IDENT3');
      if (ident2 = 0) then setfocuscontrol('PPO_IDENT2');
      Lasterror := 1;
      exit;
      end
      else
        if ((ident1 = ident2) or (ident1 = ident3) or (ident2 = ident3)) then begin
        PGIBOX('Les critères doivent être différents.', Letitre);
        setfocuscontrol('PPO_IDENT1');
        Lasterror := 1;
        exit;
        end;
   end;

  // si Nombre identifiant = 4 //
  if (nbident = 4) then begin

      if ((ident1 = 0) or (ident2= 0) or (ident3 = 0) or (ident4 = 0)) then begin
      PGIBOX('Vous devez renseigner : Critère 1, Critère 2, Critère 3 et Critère 4.', letitre);
      setfocuscontrol('PPO_IDENT1');
      if (ident4 = 0) then setfocuscontrol('PPO_IDENT4');
      if (ident3 = 0) then setfocuscontrol('PPO_IDENT3');
      if (ident2 = 0) then setfocuscontrol('PPO_ident2');
      Lasterror := 1;
      exit;
      end
      else
       if ((ident1 = ident2) or (ident1 = ident3) or (ident1 = ident4) or (ident2 = ident3)
       or (ident2 = ident4) or (ident3 = ident4)) then begin
       PGIBOX('Les critères doivent être différents.', Letitre);
       setfocuscontrol('PPO_IDENT1');
       Lasterror := 1;
       exit;
       end;
   end;

  listecritere := '';
//  nbident := getfield('PPO_NBIDENT');
  for i := 1 to nbident do
  begin
    Sti := IntToStr(i);
    numident := getfield('PPO_IDENT' +sti);
    Q1 := Opensql('select PAI_PREFIX,PAI_SUFFIX from PAIEPARIM where PAI_IDENT = "' + numident + '"', true);
    nomchamp := Q1.findfield('PAI_PREFIX').asstring + '_' + Q1.findfield('PAI_SUFFIX').asstring;
    ferme(Q1);
    listecritere := listecritere + ';' + nomchamp;
  end;
  listecritere := copy(listecritere, 2, length(listecritere));

  if bo and (critererenseigne(listecritere) <> 0) then
  begin
    bo := False;
    Lasterror := 1;
    exit;
  end;  }
end;

{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 06/04/2006
Modifié le ... :   /  /    
Description .. : Aprés Mise à jour
Mots clefs ... : 
*****************************************************************}
procedure TOM_CODEPOPULATION.OnAfterUpdateRecord;
// pt3
{var
nbident,i : integer;
critere, sti, numident, listecritere, nomchamp :string;
Q1 : TQuery;  }
begin
  inherited;
  AvertirTable('PGCRITEREPOPUL');
// SetControlEnabled('BSPOPULATION',True);
{  listecritere := '';
  nbident := getfield('PPO_NBIDENT');
  for i := 1 to nbident do
  begin
    Sti := IntToStr(i);
    numident := getfield('PPO_IDENT' +sti);
    Q1 := Opensql('select PAI_PREFIX,PAI_SUFFIX from PAIEPARIM where PAI_IDENT = "' + numident + '"', true);
    nomchamp := Q1.findfield('PAI_PREFIX').asstring + '_' + Q1.findfield('PAI_SUFFIX').asstring;
    ferme(Q1);
    listecritere := listecritere + ';' + nomchamp;
  end;
  listecritere := copy(listecritere, 2, length(listecritere));
  if critererenseigne(listecritere) = 0 then Ecran.Close; }
  Ecran.Close;
end;



{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 06/04/2006
Modifié le ... :   /  /    
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOM_CODEPOPULATION.OnAfterDeleteRecord;
begin
  inherited;
  AvertirTable('PGCRITEREPOPUL');
end;

{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 06/04/2006
Modifié le ... :   /  /    
Description .. :  Chargement
Mots clefs ... : 
*****************************************************************}
procedure TOM_CODEPOPULATION.OnLoadRecord;

begin
  inherited;
  bo := True;
  verif := False;
  GblTypePop := GetField('PPO_TYPEPOP');
  initfiche;   // initialise la fiche
end;

{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 04/04/2006
Modifié le ... :   /  /    
Description .. : Modification des champs 
Mots clefs ... : 
*****************************************************************}
procedure TOM_CODEPOPULATION.OnChangeField(F: TField);
var
pred : string;
begin
  inherited;

 // if DS.State in [DsEdit,DsInsert] then
 //   SetControlEnabled('BSPOPULATION',False)
 // else
 //   SetControlEnabled('BSPOPULATION',True);


  if (F.FieldName = 'PPO_PREDEFINI') and (ds.State = DsInsert)  then
  begin

    Pred := GetField('PPO_PREDEFINI');

    if Pred = '' then exit;
    AccesPredefini('TOUS', CEG, STD, DOS);
    if (Pred = 'CEG') and (CEG = FALSE) then
    begin
      PGIBox('Vous ne pouvez pas créer de population prédéfinie CEGID', Letitre);
      Pred := 'DOS';
      SetControlProperty('PPO_PREDEFINI', 'Value', Pred);
    end;
    if (Pred = 'STD') and (STD = FALSE) then
    begin
      PGIBox('Vous ne pouvez pas créer de population prédéfinie Standard', Letitre);
      Pred := 'DOS';
      SetControlProperty('PPO_PREDEFINI', 'Value', Pred);
    end;
    if Pred <> GetField('PPO_PREDEFINI') then SetField('PPO_PREDEFINI', pred);

    setfocuscontrol('PPO_PREDEFINI');
    initident;
  end;

end;

{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 06/04/2006
Modifié le ... :   /  /
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOM_CODEPOPULATION.OnArgument(S: string);
var
//Btn : TToolBarButton97;
//NumEdit : THDBSpinEdit;
Stplus : string;
{$IFNDEF EAGLCLIENT}
TYPEPOPX : THDBmultivalcombobox;
NumEdit : THDBSpinEdit;
codepop : THDBedit; // pt1
{$ELSE}
TYPEPOPX : THMultiValComboBox;
NumEdit : THSpinEdit;
codepop : Thedit;   // pt1
{$ENDIF}
begin
  inherited;
  AccesPredefini('TOUS', CEG, STD, DOS);

 // Btn := TToolBarButton97(GetControl('BSPOPULATION'));
 // if btn <> nil then Btn.Onclick := Ordrepopulation;
  Letitre := 'Code critère';

  // Afficher type population PRE (Presence) et FOR (formation) uniquement si les modules sont sérialisés
  // deb Pt4
  stplus := '';
  if not VH_PAIE.PGSeriaFormation then stplus := ' AND CO_CODE <> "FOR"';
  if not VH_PAIE.PGSeriaPresence then  stplus := stplus + ' AND CO_CODE <> "PRE"';
  if not VH_Paie.PgSeriaIDR then  stplus := stplus + ' AND CO_CODE <> "IDR"';    // pt10
  if stplus <> '' then SetControlProperty('PPO_TYPEPOP', 'Plus', stplus);
  // fin pt4

  {$IFNDEF EAGLCLIENT}
  TYPEPOPX := THDBmultivalcombobox(Getcontrol('PPO_TYPEPOP'));
  Numedit:= THDBSpinEdit(GetControl('PPO_NBIDENT'));
  codepop:= THDBEdit(GetControl('PPO_CODEPOP'));      // pt1
  {$ELSE}
  TYPEPOPX := THmultivalcombobox(Getcontrol('PPO_TYPEPOP'));
  Numedit:= THSpinEdit(GetControl('PPO_NBIDENT'));
  codepop:= THEdit(GetControl('PPO_CODEPOP'));      // pt1
  {$ENDIF}
  (GetControl('BValider1') as TToolbarButton97).OnClick := OnValide1Click;
  if Assigned(TYPEPOPX) then TYPEPOPX.Onchange := ChangeTypepop;
  if Assigned(TYPEPOPX) then TYPEPOPX.Onexit   := ChangeTypepop;
  if Assigned(codepop)  then codepop.Onexit    := Controlecodepop;  // pt1
  if Assigned(NumEdit)  then NumEdit.OnChange  := ChangeNbIdent;
  with (GetControl('BSAISIESALARIE') as TToolbarButton97) do
  begin
    OnClick := OnSaisieSalarieClick; //PT7
    enabled := False;
  end;
end;


procedure TOM_CODEPOPULATION.OnClose;
begin
  inherited;
end;

procedure TOM_CODEPOPULATION.OnCancelRecord;
begin
  inherited;
end;


{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 07/03/2006
Modifié le ... :   /  /    
Description .. : Initialisation de la fiche
Mots clefs ... :
*****************************************************************}
procedure TOM_CODEPOPULATION.initfiche;
var
  nbidentA, Pred, Nodos, Codepop, SQL : string;
  modif : boolean;
begin
if ds.state <> dsinsert then
begin
  // Gestion des prédéfinis
  LectureSeule := FALSE;

  if (Getfield('PPO_PREDEFINI') = 'CEG') then
  begin
    LectureSeule := (CEG = False);
    PaieLectureSeule(TFFiche(Ecran), (CEG = False));
 //   SetControlEnabled('BDelete', CEG);
  end;

  if (Getfield('PPO_PREDEFINI') = 'STD') then
  begin
    LectureSeule := (STD = False);
    PaieLectureSeule(TFFiche(Ecran), (STD = False));
    SetControlEnabled('BDelete', STD);
  end;

  if (Getfield('PPO_PREDEFINI') = 'DOS') then
  begin
    LectureSeule := False;
    PaieLectureSeule(TFFiche(Ecran), false);
    SetControlEnabled('BDelete', DOS);
  end;

  if (lectureseule = true) then Exit;

  SetControlEnabled('PPO_PREDEFINI', False);
  SetControlEnabled('PPO_CODEPOP', False);
 end;


  NbidentA := GetControlText('PPO_NBIDENT');

  Pred := Getfield('PPO_PREDEFINI');
  Nodos := Getfield('PPO_NODOSSIER');
  Codepop := Getfield('PPO_CODEPOP');

  setcontrolenabled('PPO_NBIDENT', false);
  setcontrolenabled('PPO_IDENT1', false);
  setcontrolenabled('PPO_IDENT2', false);
  setcontrolenabled('PPO_IDENT3', false);
  setcontrolenabled('PPO_IDENT4', false);
  setcontrolenabled('PPO_TYPEPOP', false);

  { si des détails population ont été saisis, le nombre d'identifiant,les identifiants et
   le type de population sont non modifiables  }


  SQL := 'Select PPC_POPULATION from ORDREPOPULATION Where PPC_PREDEFINI = "'+Pred+'" '+
  ' and PPC_NODOSSIER = "'+Nodos+'" and PPC_CODEPOP = "'+Codepop+'"';
  if not (ExisteSQL(SQL))  then
  begin
  if ds.state = dsbrowse then modif := false else modif := true;

  setcontrolenabled('PPO_TYPEPOP', true);
  setcontrolenabled('PPO_NBIDENT', true);


  // Selon le nombre d'identifiant, rend modifiable ou non l'identifiant 1,2,3,4
  if nbidentA = '1' then begin
    setcontrolenabled('PPO_IDENT1', true);
    if modif then
    begin
       setfield('PPO_IDENT2', '');
       setfield('PPO_IDENT3', '');
       setfield('PPO_IDENT4', '');
    end;
  end
  else
    if nbidentA = '2' then begin
      setcontrolenabled('PPO_IDENT1', true);
      setcontrolenabled('PPO_IDENT2', true);
      if modif then
      begin
        setfield('PPO_IDENT3', '');
        setfield('PPO_IDENT4', '');
      end;
    end
    else
      if nbidentA = '3' then begin
      setcontrolenabled('PPO_IDENT1', true);
      setcontrolenabled('PPO_IDENT2', true);
      setcontrolenabled('PPO_IDENT3', true);

      if modif then setfield('PPO_IDENT4', '');
      end
      else
       if nbidentA = '4' then begin
       setcontrolenabled('PPO_IDENT1', true);
       setcontrolenabled('PPO_IDENT2', true);
       setcontrolenabled('PPO_IDENT3', true);
       setcontrolenabled('PPO_IDENT4', true);
      end;

end;
       // initialise les identifiants
      initident;

end;


{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 07/04/2006
Modifié le ... :   /  /    
Description .. : Mise à jour des tablettes de chaque identifiant
Mots clefs ... : 
*****************************************************************}
procedure TOM_CODEPOPULATION.initident;
var
    Typepop, Typepopulation, St, StWhere , predefini: string;
    nbrparam : integer;

begin
// deb PT2
   predefini := getfield('PPO_PREDEFINI');
// fin PT2

  // Affichage des champs utilisable de PAIEPARIM selon le type de population choisi
  Typepop := Getcontroltext('PPO_TYPEPOP');
  Typepopulation := Typepop;

  StWhere :='';
  St := ReadTokenSt(Typepop); //IDR;CP
  While (St <> '') do
  Begin
     StWhere := StWhere + ' PAI_UTILISABLEPOP like "%'+St+'%" AND ';
     St := ReadTokenSt(Typepop);
  End;

  if StWhere <> '' then
  Begin
    // deb PT2 : si prédéfini = CEG ou STD : affichage des critères de type "Commun"
    //       si prédéfini = DOS        : affichage des critères de type "dosiier"
    if (predefini = 'CEG') or (predefini = 'STD') then
    stwhere := stwhere + 'PAI_PGTYPEUTILIS = "COM" AND ';
    // fin PT2

    // pt5 Affichage des codes travail , combo libre et code stat uniquement si ces champs sont gérés dans la fiche
    if VH_Paie.PGlibcodestat = '' then stwhere := stwhere + 'PAI_IDENT <> 161 AND ';
    nbrparam := VH_Paie.PGNbreStatorg;
    if nbrparam < 4 then
    begin
      Case nbrparam of
       0 : stwhere := stwhere + 'PAI_IDENT <> 157 AND PAI_IDENT <> 158 AND PAI_IDENT <> 159 AND PAI_IDENT <> 160 AND ';
       1 : stwhere := stwhere + 'PAI_IDENT <> 158 AND PAI_IDENT <> 159 AND PAI_IDENT <> 160 AND ';
       2 : stwhere := stwhere + 'PAI_IDENT <> 159 AND PAI_IDENT <> 160 AND ';
       3 : stwhere := stwhere + 'PAI_IDENT <> 160 AND ';
      end;
    end;
    nbrparam := VH_Paie.PGNBCombo;
    if nbrparam < 4 then
    begin
      Case nbrparam of
       0 : stwhere := stwhere + 'PAI_IDENT <> 345 AND PAI_IDENT <> 346 AND PAI_IDENT <> 347 AND PAI_IDENT <> 348 AND ';
       1 : stwhere := stwhere + 'PAI_IDENT <> 346 AND PAI_IDENT <> 347 AND PAI_IDENT <> 348 AND ';
       2 : stwhere := stwhere + 'PAI_IDENT <> 347 AND PAI_IDENT <> 348 AND ';
       3 : stwhere := stwhere + 'PAI_IDENT <> 348 AND ';
      end;
    end;
    // fin pt5

    StWhere := '('+Copy(StWhere,1,Length(StWhere)-4)+')';
    SetControlProperty('PPO_IDENT1','Plus',StWhere);
    SetControlProperty('PPO_IDENT2','Plus',StWhere);
    SetControlProperty('PPO_IDENT3','Plus',StWhere);
    SetControlProperty('PPO_IDENT4','Plus',StWhere);
  End;

end;

{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 07/03/2006
Modifié le ... :   /  /
Description .. : Bouton "Gestion des détails population"
Mots clefs ... :
*****************************************************************}
//Procedure TOM_CODEPOPULATION.Ordrepopulation (sender : TObject);
//Begin
//AglLancefiche('PAY', 'ORDREPOPUL_MUL', '', '',Getfield('PPO_CODEPOP')+';'+Getfield('PPO_PREDEFINI')+';'+
//Getfield('PPO_NODOSSIER')+';'+Getfield('PPO_LIBELLE'));
//initfiche
//end;

{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 06/04/2006
Modifié le ... :   /  /    
Description .. : Changement de la zone "Nombre identifiant"
Mots clefs ... : 
*****************************************************************}
procedure TOM_CODEPOPULATION.ChangeNbIdent(sender: TObject);
begin
initfiche;
end;

{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 06/04/2006
Modifié le ... :   /  /    
Description .. : Changement du type de population
Mots clefs ... : 
*****************************************************************}
procedure TOM_CODEPOPULATION.Changetypepop(sender: TObject);
begin
  if not (ds.state in [dsbrowse]) then
  begin

    if   GblTypePop <> getfield('PPO_TYPEPOP') then
    begin
//   SetControlEnabled('BSPOPULATION',False);
      if (GetFIELD('PPO_IDENT1') <> 0)
      then setFIELD('PPO_IDENT1',0);
      if (GetFIELD('PPO_IDENT2') <> 0)
      then setFIELD('PPO_IDENT2',0);
      if (GetFIELD('PPO_IDENT3') <> 0)
      then setFIELD('PPO_IDENT3',0);
      if (GETFIELD('PPO_IDENT4') <> 0)
      then setFIELD('PPO_IDENT4',0);

    initident;
    end;
  end;
end;

// Deb PT1
{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 19/02/2007
Modifié le ... :   /  /    
Description .. : Controle du code population
Mots clefs ... : 
*****************************************************************}
procedure TOM_CODEPOPULATION.Controlecodepop(sender: TObject);
var
codepop,sql, predefini : string;
i : integer;
popOK : boolean;

begin
// contrôle du code population

codepop := Getfield('PPO_CODEPOP');
predefini := Getfield('PPO_PREDEFINI');

if length(Codepop) <> 3 then begin
  PGIBOX('Le code critère doit comporter 3 caractères.', Letitre);
  setfocuscontrol('PPO_CODEPOP');
  Lasterror := 1;
  exit;
  end;


  for i := 1 to 3 do
  begin

  if (codepop[i] in ['0'..'9'])  or (codepop[i] in ['A'..'Z']) or (codepop[i] in ['a'..'z']) then
  popOK := true
  else
  popOK:= false;
  if not popOK then break;
  end;

  if not popOK then
  begin
  PGIBOX('Le code critère doit comporter uniquement des chiffres ou des lettres.', Letitre);
  setfocuscontrol('PPO_CODEPOP');
  Lasterror := 1;
  exit;
  end;


if  ((predefini = 'STD') and (Codepop[3] <> '1') and (Codepop[3] <> '3')) then begin
  PGIBOX('Le code critère doit se terminer par 1 ou 3.', Letitre);
  setfocuscontrol('PPO_CODEPOP');
  Lasterror := 1;
  exit;
  end;
if ((predefini = 'CEG') and (Codepop[3] <> '0') and (Codepop[3] <> '2') and (codepop[3] <> '4') and
    (codepop[3] <> '6') and (codepop[3] <> '8'))
  then begin
  PGIBOX('Le code critère doit se terminer par 0,2,4,6 ou 8', Letitre);
  setfocuscontrol('PPO_CODEPOP');
  Lasterror := 1;
  exit;
  end;
If ((predefini = 'DOS') and (Codepop[3] <> '5') and (Codepop[3] <> '7') and (Codepop[3] <> '9')) then begin
  PGIBOX('Le code critère doit se terminer par 5,7 ou 9.', Letitre);
  setfocuscontrol('PPO_CODEPOP');
  Lasterror := 1;
  exit;
  end;

  // contrôle code population non existant


    Sql := 'Select PPO_CODEPOP from CODEPOPULATION where ##PPO_PREDEFINI## PPO_CODEPOP = "'+codepop+'"';
    if ExisteSql(Sql) then
    begin
      PGIBox('Ce code critère existe déjà', Letitre);
      setfocuscontrol('PPO_CODEPOP');
    end;
end;
// fin PT1

// pt3
{***********A.G.L.***********************************************
Auteur  ...... : NA
Créé le ...... : 21/05/2007
Modifié le ... :   /  /    
Description .. : Contrôle que les critères utilisées sont bien renseignés dans 
Suite ........ : la fiche salarié
Mots clefs ... : 
*****************************************************************}
Function TOM_CODEPOPULATION.Critererenseigne(listeCritere, listenomcritere : String) : Integer; // pt6
var
critere, stwhere, salarie,nomsal,prenomsal, nomcritere : string;
datejour : Tdatetime;
QQ : TQuery;
Listeanomalies : Tlistbox;
anomalie : boolean;
Tob_CriterenonOK, T1 : Tob; // pt6

begin
  result := 0;
  Listeanomalies := TListbox(Getcontrol('TANOMALIE'));
  listeanomalies.Items.clear;
  anomalie := false;
  Datejour := Date;
  stwhere := '';
  critere := ReadTokenSt(listecritere);
  nomcritere := Readtokenst(listenomcritere); // pt6
  Tob_CriterenonOK := Tob.Create('Critères non renseignés',Nil,-1);   // pt6

  While (critere <> '') do
  begin

 // stwhere := stwhere + ' or '  + critere + '= ""';   pt6
   stwhere := critere + '= ""';     // pt6
    // critere := ReadTokenSt(listecritere); pt6
 // end; pt6
 //  stwhere :=copy(stwhere, 5, length(stwhere));  pt6

    QQ := opensql('Select PSA_SALARIE, PSA_LIBELLE,PSA_PRENOM FROM SALARIES WHERE '+stwhere+' and '+
              '(PSA_DATESORTIE >="'+usdatetime(datejour)+'" or PSA_DATESORTIE <="'+UsdateTime(Idate1900)+'" OR' +
              ' PSA_DATESORTIE IS NULL) and (PSA_DATEENTREE <= "'+usdatetime(datejour)+'")', true);

    // pt6
    // recherche nom du critère

    QQ.First;
    While not QQ.EOF DO
    begin
      T1 := Tob.Create('Critères non renseignés',Tob_CriterenonOK,-1);
      T1.AddChampSupValeur('SALARIE',QQ.Findfield('PSA_SALARIE').asstring);
      T1.AddChampSupValeur('NOM',QQ.Findfield('PSA_LIBELLE').asstring);
      T1.AddChampSupValeur('PRENOM',QQ.Findfield('PSA_PRENOM').asstring);
      T1.AddChampSupValeur('CRITERE', nomcritere);
      QQ.Next;
    end;
    ferme(QQ);
    critere := ReadTokenSt(listecritere);
    nomcritere := Readtokenst(listenomcritere);
   end;
  if Tob_CriterenonOK <> nil  then Tob_CriterenonOK.detail.sort('SALARIE');

//  While Not QQ.Eof Do
   T1:=Tob_CriterenonOK.FindFirst([''],[''], false );
   listeCodesSal := ''; //PT7
   while T1 <> nil do
   Begin
    salarie := T1.getvalue('SALARIE');

    listeCodesSal := listeCodesSal + ';' + salarie; // pt7

    nomsal :=  T1.Getvalue('NOM');
    prenomsal := T1.Getvalue('PRENOM');
    nomcritere := T1.Getvalue('CRITERE');
 //   salarie := QQ.findfield('PSA_SALARIE').asstring;
 //   nomsal :=  QQ.findfield('PSA_LIBELLE').asstring;
 //   prenomsal := QQ.findfield('PSA_PRENOM').asstring;
 // fin pt6
    anomalie := true;
    listeanomalies.items.Add('Le critère ' + nomcritere+  ' n''est pas renseigné pour le salarié ' + salarie + ' ' + nomsal + ' ' + prenomsal );
   // QQ.next; pt6
    T1:=Tob_CriterenonOK.FindNext([''],[''], false );  // pt6
   end;
   listeCodesSal     := copy(listeCodesSal, 2, length(listeCodesSal)); // pt7

 if anomalie then
 begin
    PGIINFO('Les critères sélectionnés ne sont pas renseignés dans la fiche salarié. Voir page Anomalies.', Ecran.Caption);
    (GetControl('PAGES') as TPageControl).ActivePageIndex := 1;
    result := 1;
 end;
 ferme(QQ);
 Freeandnil(Tob_CriterenonOK); // pt6
end;
// fin pt3

procedure TOM_CODEPOPULATION.OnValide1Click(Sender: TObject);
var
  nbident : integer;
  typepop, predefini : string;
  ident1 : integer;
  ident2 : integer;
  ident3 : integer;
  ident4 : integer;


i : integer; // nbident,
critere, sti, numident, listecritere, listenomcritere, nomchamp :string;
Q1 : TQuery;
begin
  SetFocusControl('PPO_LIBELLE');
  verif := True;
  nbident := Getfield('PPO_NBIDENT');
  typepop := GetControlText('PPO_TYPEPOP');
  predefini := GetControlText('PPO_PREDEFINI');
  ident1 :=  strtoint(Getfield('PPO_IDENT1'));
  ident2 :=  strtoint(Getfield('PPO_IDENT2'));
  ident3 :=  strtoint(Getfield('PPO_IDENT3'));
  ident4 :=  strtoint(Getfield('PPO_IDENT4'));

  if predefini = 'DOS' then setfield('PPO_NODOSSIER', PgrendNodossier())  else
  setfield('PPO_NODOSSIER', '000000');

// Contrôle du type de population
if typepop = '' then begin
  PGIBOX('Vous devez renseigner: Type de population.', Letitre);
  setfocuscontrol('PPO_TYPEPOP');
  Lasterror := 1;
  exit;
end;

//  Contrôle des identifiants

  if nbident = 0 then begin
  PGIBOX('Vous devez renseigner: Nombre de critères.', Letitre);
  setfocuscontrol('PPO_NBIDENT');
  Lasterror := 1;
  exit;
  end;

  if (nbident = 1) and  (ident1 = 0) then begin
  PGIBOX('Vous devez renseigner : Critère 1.', Letitre);
  setfocuscontrol('PPO_IDENT1');
  Lasterror := 1;
  exit;
  end;

  //  si Nombre identifiant = 2//
  if (nbident = 2) then begin

      if ((ident1 = 0) or (ident2= 0)) then
      begin
      PGIBOX('Vous devez renseigner : Critère 1 et Critère 2.', Letitre);
      setfocuscontrol('PPO_IDENT1');
      if (ident1 <> 0) then setfocuscontrol('PPO_IDENT2');
      Lasterror := 1;
      exit;
      end
      else
        if ident1 = ident2 then begin
        PGIBOX('Les critères doivent être différents.', Letitre);
        setfocuscontrol('PPO_IDENT1');
        Lasterror := 1;
        exit;
      end;
  end;

  // si Nombre identifiant = 3 //
  if (nbident = 3) then begin

      if  ((ident1 = 0) or (ident2= 0) or (ident3 = 0)) then begin
      PGIBOX('Vous devez renseigner : Critère 1, Critère 2 et Critère 3.', Letitre);
      setfocuscontrol('PPO_IDENT1');
      if (ident3 = 0) then setfocuscontrol('PPO_IDENT3');
      if (ident2 = 0) then setfocuscontrol('PPO_IDENT2');
      Lasterror := 1;
      exit;
      end
      else
        if ((ident1 = ident2) or (ident1 = ident3) or (ident2 = ident3)) then begin
        PGIBOX('Les critères doivent être différents.', Letitre);
        setfocuscontrol('PPO_IDENT1');
        Lasterror := 1;
        exit;
        end;
   end;

  // si Nombre identifiant = 4 //
  if (nbident = 4) then begin

      if ((ident1 = 0) or (ident2= 0) or (ident3 = 0) or (ident4 = 0)) then begin
      PGIBOX('Vous devez renseigner : Critère 1, Critère 2, Critère 3 et Critère 4.', letitre);
      setfocuscontrol('PPO_IDENT1');
      if (ident4 = 0) then setfocuscontrol('PPO_IDENT4');
      if (ident3 = 0) then setfocuscontrol('PPO_IDENT3');
      if (ident2 = 0) then setfocuscontrol('PPO_ident2');
      Lasterror := 1;
      exit;
      end
      else
       if ((ident1 = ident2) or (ident1 = ident3) or (ident1 = ident4) or (ident2 = ident3)
       or (ident2 = ident4) or (ident3 = ident4)) then begin
       PGIBOX('Les critères doivent être différents.', Letitre);
       setfocuscontrol('PPO_IDENT1');
       Lasterror := 1;
       exit;
       end;
   end;


  (GetControl('BSAISIESALARIE') as TToolbarButton97).enabled := False; //PT7

  listecritere := '';
  listenomcritere := '';
  listeCritereIdent := ''; //PT7
//  nbident := getfield('PPO_NBIDENT');
  for i := 1 to nbident do
  begin
    Sti := IntToStr(i);
    numident := getfield('PPO_IDENT' +sti);
    listeCritereIdent := ListeCritereIdent + ';' + numident; //PT7
    Q1 := Opensql('select PAI_PREFIX,PAI_SUFFIX , PAI_LIBELLE from PAIEPARIM where PAI_IDENT = "' + numident + '"', true);
    nomchamp := Q1.findfield('PAI_PREFIX').asstring + '_' + Q1.findfield('PAI_SUFFIX').asstring;
    listenomcritere := listenomcritere + ';' + Q1.Findfield('PAI_LIBELLE').asstring; // pt6
    ferme(Q1);
    listecritere := listecritere + ';' + nomchamp;
  end;
  listecritere := copy(listecritere, 2, length(listecritere));
  listenomcritere := copy(listenomcritere, 2, length(listenomcritere)); // pt6
  listeCritereIdent := copy(listeCritereIdent, 2, length(listeCritereIdent)); // pt7
  (GetControl('BSAISIESALARIE') as TToolbarButton97).enabled := True; //PT7

  if bo and (critererenseigne(listecritere, listenomcritere) <> 0) then    //PT7
  begin
    bo := False;
    Lasterror := 1;
    exit;
  end;

  (GetControl('BValider') as TToolbarButton97).OnClick(Self);
end;

procedure TOM_CODEPOPULATION.OnSaisieSalarieClick(Sender: TObject);   //PT7
begin
  AGllanceFiche('PAY', 'MULHISTOGROUPEE', '', '', 'CRITERE='+listeCritereIdent+'|SALARIES='+listeCodesSal);
end;

initialization
  registerclasses([TOM_CODEPOPULATION]);
end.

