{***********UNITE*************************************************
Auteur  ...... : GGU
Créé le ...... : 07/02/2007
Modifié le ... :   /  /
Description .. : Source TOM de la TABLE : TABLEDIMENT (TABLEDIMENT)
Mots clefs ... : TOM;TABLEDIMENT;PAIE
*****************************************************************}
{
PT1 05/04/2007 V7.02 GGU Alerte dès la sortie du champ de code Table si celui-ci
               est invalide
PT2 11/04/2007 V7.02 GGU Quand un des critères est vide stocker 0 à la place de null
               et quand on récupère 0 lors du chargement, afficher <<Aucun critère>>
PT3 02/05/2007 V7.02 GGU Gestion des tables dynamiques de type Variable
PT4 26/06/2007 V7.02 GGU FQ 14209 Dans la première colonne, après une saisie on doit faire deux fois TAB pour passer à la colonne suivante.
PT5 31/07/2007 V8 GGU gestion du niveau de saisie pour toutes les natures de table 
PT6 06/11/2007 FC V_80 : FQ 14908 Journal événement
PT7 21/11/2007 V8 GGU FQ14964 Problème d'affichage du premier critère dans certains cas
PT8 21/11/2007 V8 GGU FQ14969 listes déroulantes vides si option d'affichage 'Onglet sur plusieurs lignes' est cochée
PT9 26/11/2007 V8 GGU Correction d'un bug en multi dossier : quand on saisi une
                      table prédéfini, le champs dossier est renseigné avec le numéro de dossier, au lieu de 000000
PT10 12/12/2007 V8 GGU FQ 14451 nature de table : Variables --> remplacer accès liste par elipsis
PT11 13/12/2007 V8 GGU FQ 14620 Ajout d'une 4e nature de table : Elément national
}
Unit UTOMTABLEDIMENT;

Interface

Uses
 Controls, Classes,
{$IFNDEF EAGLCLIENT}
     db, Fiche, FichList,
     {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
{$else}
     eFiche, eFichList,
{$ENDIF}
     forms, UTOM, sysutils, HCtrls, HEnt1, HMsgBox, UTob, htb97, ed_tools, HDB,
     PAIETOM ; //PT6

//Debut PT10
Type
  TControlAbstractComboEdit = class
  private
    FControlToAbstract: TControl;
    FLinkTom: Tom;
    procedure SetDataType(const Value: String);
    procedure SetPlus(const Value: String);
    procedure SetVide(const Value: Boolean);
    function GetDataType: String;
    function GetPlus: String;
    function GetVide: Boolean;
    procedure SetControlToAbstract(const Value: TControl);
    function GetControlText: String;
    function GetVisible: Boolean;
    procedure SetControlText(const Value: String);
    procedure SetVisible(const Value: Boolean);
    procedure SetLinkTom(const Value: Tom);
    function GetEnabled: Boolean;
    procedure SetEnabled(const Value: Boolean);
  public
    Property LinkTom : Tom read FLinkTom write SetLinkTom;
    property ControlToAbstract : TControl read FControlToAbstract write SetControlToAbstract;
    Property Vide : Boolean read GetVide write SetVide;
    Property Visible : Boolean read GetVisible write SetVisible;
    Property Enabled : Boolean read GetEnabled write SetEnabled;
    Property Plus : String read GetPlus write SetPlus;
    Property DataType : String read GetDataType write SetDataType;
    Property ControlText : String read GetControlText write SetControlText;
    Procedure SetFocus;
    Constructor Create(LaTom : Tom); 
    Destructor Destroy; override;
  end;
//Fin PT10

Type
  TOM_TABLEDIMENT = Class (PGTOM) //PT6 class TOM devient PGTOM
    procedure OnNewRecord             ; override;
    procedure OnDeleteRecord          ; override;
    procedure OnUpdateRecord          ; override;
    procedure OnAfterUpdateRecord     ; override;
    procedure OnAfterDeleteRecord     ; override;
    procedure OnLoadRecord            ; override;
    procedure OnChangeField(F: TField); override;
    procedure OnArgument(S: String )  ; override;
    procedure OnClose                 ; override;
    procedure OnCancelRecord          ; override;
  private
    PDETAIL: THGrid;
    BtnIns, BtnDel, BtnAdd : TToolBarButton97;
    TobTableDynamique : TOB;
    Action : String;
    MultiLineOnglet : Boolean;
    TabletteCritere1, TabletteCritere2, PlusCritere1, PlusCritere2, WhereVar, WhereElt : String;
    BackTabletteCritere1, BackTabletteCritere2, BackPlusCritere1, BackPlusCritere2 : String;
    Default_NATURETABLE, Default_NIVSAIS, Default_VALNIV, Default_PREDEFINI, Default_DTVALID: String;
    Key : String;
    CEG, STD, DOS, Autorisation : Boolean; //Autorisations d'accès
    Trace: TStringList;                //PT6
    DerniereCreate: string;            //PT6
    LeStatut:TDataSetState;            //PT6
{$IFNDEF EAGLCLIENT}
    //Control_CRITERE1,
    Control_CRITERE2, Control_NATURETABLE, Control_NIVSAIS, Control_TYPERESULTAT, Control_PREDEFINI : THDBValComboBox;
    Control_TABLE : THDBEdit;
{$ELSE}
    //Control_CRITERE1,
    Control_CRITERE2, Control_NATURETABLE, Control_NIVSAIS, Control_TYPERESULTAT, Control_PREDEFINI : THValComboBox;
    Control_TABLE : THEdit;
{$ENDIF}
    Control_CRITERE1B : THValComboBox;
    Control_CRITERE1B_ : THEdit; //PT10
    Control_AbsCRITERE1B : TControlAbstractComboEdit; //PT10
    Control_COMBINCRITERE : TToolbarButton97;
    Control_VALNIV : THEdit;

    boAlerteInitGrid, boLoading, boControlTableAlreadyExist : Boolean;
    //Variables de la gestion des annulations
    BackTYPERESULTAT, BackTABLE, BackNATURETABLE, BackPREDEFINI, BackCRITERE1, BackCRITERE2 : String;
    isBackTYPERESULTAT, isBackNATURETABLE, isBackPREDEFINI, isBackCRITERE1, isBackCRITERE2 : Boolean;
    //Valeurs utilisées dans toute l'unité et mise à jour par les procédures "UpdateValues" et "UpdateVisibilites"
    stNATURETABLE, stNIVSAIS, stCRITERE1, stCRITERE2, stPREDEFINI, stVALNIV, stSENS1, stSENS2,
    stNODOSSIER, stCODTABL, stTYPECRITERE1, stTYPECRITERE2, stTYPERESULTAT : string;
    boDonneesSalarie, boNiveauEtab_ou_Convention, boCritere1AgeAnciennete, boCritere2AgeAnciennete, boCOMBINCRITERE : Boolean;

    function GetNbCriteres: integer;
    property NbCriteres : integer read GetNbCriteres;
    //Fonctions de la gestion des annulation et des mises à jour de l'affichage en fonction des données
    procedure OnEnterNATURETABLE    (Sender : TObject);
    procedure OnChangeNATURETABLE   (Sender : TObject);
    procedure OnExitNATURETABLE     (Sender : TObject);
    procedure OnEnterTABLE          (Sender : TObject);
    procedure OnExitTABLE           (Sender : TObject);
    procedure OnChangeTABLE         (Sender : TObject);
    procedure OnEnterTYPERESULTAT   (Sender : TObject);
    procedure OnChangeTYPERESULTAT  (Sender : TObject);
    procedure OnEnterPREDEFINI      (Sender : TObject);
    procedure OnChangePREDEFINI     (Sender : TObject);
    procedure OnEnterCRITERE1B      (Sender : TObject);
    procedure OnExitCRITERE1B       (Sender : TObject);
    procedure OnChangeCRITERE1B     (Sender : TObject);
    procedure OnEnterCRITERE2       (Sender : TObject);
    procedure OnExitCRITERE2        (Sender : TObject);
    procedure OnChangeNIVSAIS       (Sender : TObject);
    procedure OnChangeVALNIV        (Sender : TObject);
    procedure OnClickCOMBINCRITERE  (Sender : TObject);
    procedure InitGrid;
    procedure InitEntete;
    procedure PDetailCellEnter(Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
    procedure PDetailCellExit (Sender: TObject; var ACol, ARow: Integer; var Cancel: Boolean);
    procedure PDetailElipsisClick(Sender: TObject);
    procedure BtnInsClick(Sender: TObject);
    procedure BtnDelClick(Sender: TObject);
    procedure BtnAddClick(Sender: TObject);
    Procedure LoadGrid;
    Procedure SaveGrid;
    Procedure UpdateValues;
    Procedure UpdateVisibilites;
    function  UpdatePlusCriteres : Boolean;
    Procedure UpdateAutorisation;
    Procedure UpdateCriterePRNature(Force : Boolean);
    function  SetDetailsFieldsNotInGrid(T: TOB): integer;
    function  FindTabletteName(IDENT: integer): String;
    Procedure GetPrefixeCodeLibelleTablette(Tablette : String; Var Prefixe, Code, Libelle : String);
    procedure AffecteEvenements;
    procedure GridFormatCol0(ARow: Integer);
  end;

Implementation

uses
  lookup, DateUtils, PgOutils2, TntStdCtrls, TntGrids, wCommuns,
  hdtlinks, Pgoutils, grids, StrUtils, utobdebug, P5Util, Variants, Ent1,
  PGTablesDyna, P5Def;

const
  ListeDesChampsDuGrid : string = 'PTD_VALCRIT1;PTD_VALCRIT2;PTD_LIBELLECODE;PTD_RESULTAT';

function lpad(Str: String; Size : Integer; Pad : String = '0') : String;
begin
  result := Str;
  While length(result) < Size do
  begin
    result := Pad + result;
  end;
end;

Function CountCharInStr(ArrayChar : Array of char; Str : String) : Integer;
var
  i,j : Integer;
begin
  result := 0;
  for i := 0 to High(ArrayChar) do
  begin
    for j := 0 to Length(Str) do
    begin
      if ArrayChar[i] = str[j] then Result := Result +1;
    end;
  end;
end;

Function StrToNumStr(StrInit: String) : String;
begin
  result := '';
  if (StrInit <> '') and (IsNumeric(StrInit)) and (CountCharInStr(['.',','],StrInit) <= 1) then
  begin
    try
      StrToFloat(StrInit);
      result := StrInit;
    except
      if pos('.',StrInit) > 0 then
        result := StringReplace(StrInit,'.',',',[])
      else if pos(',',StrInit) > 0 then
        result := StringReplace(StrInit,',','.',[]);
    end;
  end;
end;

Procedure TOM_TABLEDIMENT.GridFormatCol0(ARow : Integer);
var
  tempstr : String;
begin
//PT11  if (Getfield('PTE_NATURETABLE') = 'VAR') or (Getfield('PTE_NATURETABLE') = 'COD') then
  if (Getfield('PTE_NATURETABLE') = 'ELT') or (Getfield('PTE_NATURETABLE') = 'VAR') or (Getfield('PTE_NATURETABLE') = 'COD') then  //PT11
  begin
    tempstr := PDETAIL.Cells [0,ARow];
    tempstr := StringReplace(tempstr,'.','0',[rfReplaceAll, rfIgnoreCase ]);
    tempstr := StringReplace(tempstr,',','0',[rfReplaceAll, rfIgnoreCase ]);
    tempstr := StrToNumStr(tempstr);
    if           (PDetail.ColFormats[0] = FormatC_DSA) then
    begin
      if Length(tempstr) < NbrC_DSA then tempstr := lpad(tempstr,NbrC_DSA)
    end else if  (PDetail.ColFormats[0] = FormatC_COD) then
    begin
      if Length(tempstr) < NbrC_COD then tempstr := lpad(tempstr,NbrC_COD);
    end;
    PDETAIL.Cells     [0,ARow] := tempstr;
    PDETAIL.CellValues[0,ARow] := tempstr;
  end;
end;

procedure TOM_TABLEDIMENT.OnNewRecord ;
begin
  Inherited;
  boLoading := True;
  boAlerteInitGrid := False;
  InitEntete;
  InitGrid;
  boLoading := False;
  UpdateVisibilites;
  boControlTableAlreadyExist := True;
end ;

procedure TOM_TABLEDIMENT.OnDeleteRecord ;
begin
  Inherited ;
  UpdateVisibilites;
  boLoading := True;
  //DEB PT6
  Trace := TStringList.Create ;
  Trace.Add('SUPPRESSION TABLE DYNAMIQUE '+GetField('PTE_CODTABL')+' '+ GetField('PTE_LIBELLE'));
  CreeJnalEvt('003','094','OK',nil,nil,Trace);
  FreeAndNil (Trace);
  //FIN PT6
end ;

procedure TOM_TABLEDIMENT.OnUpdateRecord ;
var
  i, j : integer;
  tempstr : String;
begin
  Inherited ;
  //Vérification du résultat de la dernière cellule (l'évènement OnExit de la cellule
  //n'est pas déclanché si on valide directement la saisie
  tempstr := StrToNumStr(PDETAIL.Cells[3,PDETAIL.Row]);
  PDETAIL.Cells     [3,PDETAIL.Row] := tempstr;
  PDETAIL.CellValues[3,PDETAIL.Row] := tempstr;

  GridFormatCol0(PDETAIL.Row);

{    if (Getfield('PTE_NATURETABLE') = 'VAR') or (Getfield('PTE_NATURETABLE') = 'COD') then
    begin
      tempstr := PDETAIL.Cells [0,PDETAIL.Row];
      tempstr := StringReplace(tempstr,'.','0',[rfReplaceAll, rfIgnoreCase ]);
      tempstr := StringReplace(tempstr,',','0',[rfReplaceAll, rfIgnoreCase ]);
      tempstr := StrToNumStr(tempstr);
      if           (PDetail.ColFormats[0] = FormatC_DSA) then
      begin
        if Length(tempstr) < NbrC_DSA then tempstr := lpad(tempstr,NbrC_DSA)
      end else if  (PDetail.ColFormats[0] = FormatC_COD) then
      begin
        if Length(tempstr) < NbrC_COD then tempstr := lpad(tempstr,NbrC_COD);
      end;
      PDETAIL.Cells     [0,PDETAIL.Row] := tempstr;
      PDETAIL.CellValues[0,PDETAIL.Row] := tempstr;
    end;
}
  SetFocusControl('PTE_LIBELLE');
  UpdateValues;
  //Vérification des données
  if stCODTABL = '' then
  begin
    LastErrorMsg := TraduireMemoire('Vous devez saisir le code de la table.');
    LastError := 1;
    SetFocusControl('PTE_CODTABL');
    Exit;
  end;
  if (not IsNumeric(stCODTABL)) or (pos('.',stCODTABL) > 0) or (pos(',',stCODTABL) > 0) or (length(stCODTABL) <> 3 ) then
  begin
    LastErrorMsg := TraduireMemoire('Le code de la table doit être un nombre à trois chiffres.');
    LastError := 101;
    SetControlText('PTE_CODTABL', BackTABLE);
    SetFocusControl('PTE_CODTABL');
    Exit;
  end;
//  if (stPREDEFINI = 'CEG') and (not (StrToInt(RightStr(stCODTABL,1)) mod 2 = 0) ) then
//  begin
//    LastErrorMsg := TraduireMemoire('Le code de la table doit être un nombre PAIR.');
//    LastError := 102;
//    SetControlText('PTE_CODTABL', BackTABLE);
//    SetFocusControl('PTE_CODTABL');
//    Exit;
//  end;
//  if (stPREDEFINI = 'STD') and
//     ( (    (StrToInt(RightStr(stCODTABL,1)) <> 1 )
//        and (StrToInt(RightStr(stCODTABL,1)) <> 3 )
//       )
//     ) then
//  begin
//    LastErrorMsg := TraduireMemoire('Le code de la table doit être un nombre se terminant par 1 ou 3.');
//    LastError := 102;
//    SetControlText('PTE_CODTABL', BackTABLE);
//    SetFocusControl('PTE_CODTABL');
//    Exit;
//  end;
//  if (stPREDEFINI = 'DOS') and
//     ( (    (StrToInt(RightStr(stCODTABL,1)) <> 5)
//        and (StrToInt(RightStr(stCODTABL,1)) <> 7)
//        and (StrToInt(RightStr(stCODTABL,1)) <> 9)
//       )
//     ) then
//  begin
//    LastErrorMsg := TraduireMemoire('Le code de la table doit être un nombre se terminant par 5, 7 ou 9.');
//    LastError := 102;
//    SetControlText('PTE_CODTABL', BackTABLE);
//    SetFocusControl('PTE_CODTABL');
//    Exit;
//  end;

  if stPREDEFINI = '' then
  begin
    LastErrorMsg := TraduireMemoire('Vous devez choisir votre type de prédéfini.');
    LastError := 2;
    SetFocusControl('PTE_PREDEFINI');
    Exit;
  end;
  if stNODOSSIER = '' then
  begin
    SetControlText('PTE_NODOSSIER','000000');
  end;
  if AGLStrToDate(GetControlText('PTE_DTVALID')) < EncodeDateDay(1900,10) then
  begin
    LastErrorMsg := TraduireMemoire('La date de validité est obligatoire.');
    LastError := 3;
    SetFocusControl('PTE_DTVALID');
    Exit;
  end;
  if stNATURETABLE = '' then
  begin
    LastErrorMsg := TraduireMemoire('Vous devez choisir la nature de la table.');
    LastError := 4;
    SetFocusControl('PTE_NATURETABLE');
    Exit;
  end;
  if stNATURETABLE = 'VAR' then   //PT3
  begin
    if stCRITERE1 = '' then
    begin
      LastErrorMsg := TraduireMemoire('Vous devez choisir la variable d''entrée.');
      //PT10 SetFocusControl('PTE_CRITERE1B');
      Control_AbsCRITERE1B.SetFocus; //PT10
      LastError := 6;
      Exit;
    end;
  end;
  if stNATURETABLE = 'ELT' then   //PT11
  begin
    if stCRITERE1 = '' then
    begin
      LastErrorMsg := TraduireMemoire('Vous devez choisir l''élément national d''entrée.');
      //PT10 SetFocusControl('PTE_CRITERE1B');
      Control_AbsCRITERE1B.SetFocus; //PT10
      LastError := 6;
      Exit;
    end;
  end;
//  if boDonneesSalarie then   //PT5
//  begin                      //PT5
    if stNIVSAIS = '' then
    begin
      LastErrorMsg := TraduireMemoire('Vous devez choisir le niveau de la saisie.');
      SetFocusControl('PTE_NIVSAIS');
      LastError := 5;
      Exit;
    end;
  if boDonneesSalarie then  //PT5
  begin                     //PT5
    if stCRITERE1 = '' then
    begin
      LastErrorMsg := TraduireMemoire('Vous devez choisir au moins un critère.');
      //PT10 SetFocusControl('PTE_CRITERE1B');
      Control_AbsCRITERE1B.SetFocus; //PT10
      LastError := 6;
      Exit;
    end;
  end;                     //PT5
//  end else begin         //PT5
//    SetControlText('PTE_NIVSAIS','GEN');    //PT5
//    SetControlText('PTE_VALNIVB','...');    //PT5
//  end;                                      //PT5
  if (    (stCRITERE1 = '1500')  // Age
       or (stCRITERE1 = '1501')) // Ancienneté
     and (stTYPECRITERE1 = '') then
  begin
    LastErrorMsg := TraduireMemoire('Vous devez choisir le type de votre premier critère.');
    SetFocusControl('PTE_TYPECRITERE1');
    LastError := 7;
    Exit;
  end;
  if (    (stCRITERE2 = '1500')  // Age
       or (stCRITERE2 = '1501')) // Ancienneté
     and (stTYPECRITERE2 = '') then
  begin
    LastErrorMsg := TraduireMemoire('Vous devez choisir le type de votre second critère.');
    SetFocusControl('PTE_TYPECRITERE2');
    LastError := 8;
    Exit;
  end;
  if (    (stCRITERE1 = '1500') or (stCRITERE1 = '1501'))
      and (stSENS1 = '') then
  begin
    LastErrorMsg := TraduireMemoire('Vous devez choisir le sens de votre premier critère.');
    SetFocusControl('PTE_SENS1');
    LastError := 9;
    Exit;
  end;
  if (    (stCRITERE2 = '1500') or (stCRITERE2 = '1501'))
      and (stSENS2 = '') then
  begin
    LastErrorMsg := TraduireMemoire('Vous devez choisir le sens de votre second critère.');
    SetFocusControl('PTE_SENS2');
    LastError := 10;
    Exit;
  end;
  if (stCRITERE1 = stCRITERE2) and (stCRITERE1 <> '') and (stCRITERE1 <> '0') then
  begin
    LastErrorMsg := TraduireMemoire('Les deux critères ne doivent pas être identiques.');
    SetFocusControl('PTE_TYPECRITERE2');
    LastError := 11;
    Exit;
  end;
  if stTYPERESULTAT = '' then
  begin
    LastErrorMsg := TraduireMemoire('Vous devez choisir le type de résultat.');
    SetFocusControl('PTE_TYPERESULTAT');
    LastError := 12;
    Exit;
  end;
  //Vérification des données du tableau :
  //Suppression des lignes vides
  for i := (PDetail.rowcount-1) downto 1 do
  begin
   if (PDETAIL.CellValues[0, i] = '') and (PDETAIL.CellValues[1, i] = '') and (PDETAIL.CellValues[2, i] = '') and (PDETAIL.CellValues[3, i] = '') then
     PDETAIL.DeleteRow(i);
  end;
  //Vérification des doublons
  PDETAIL.SortGrid(1,True);
  PDETAIL.SortGrid(0,True);
  if not boDonneesSalarie then
  begin
    for i := 2 to PDetail.rowcount-1 do
    begin
      if PDETAIL.CellValues[0, i] = PDETAIL.CellValues[0, i-1] then
      begin
        LastErrorMsg := TraduireMemoire('Erreur :'#10#13'Ce code existe déjà.');
        SetFocusControl('DETAILMINI');
        PDETAIL.Row := i;
        PDETAIL.Col := 0;
        LastError := 130;
        Exit;
      end;
    end;
  end else begin
    for i := 2 to PDetail.rowcount-1 do
    begin
      if (PDETAIL.CellValues[0, i] = PDETAIL.CellValues[0, i-1]) and (PDETAIL.CellValues[1, i] = PDETAIL.CellValues[1, i-1]) then
      begin
        LastErrorMsg := TraduireMemoire('Erreur :'#10#13'Ce code existe déjà.');
        SetFocusControl('DETAILMINI');
        PDETAIL.Row := i;
        PDETAIL.Col := 0;
        LastError := 130;
        Exit;
      end;
    end;
  end;

  for i := 1 to PDetail.rowcount-1 do
  begin
    for j := 0 to PDetail.Colcount-1 do
    begin
      if  (PDetail.ColFormats[j] = FormatC_DSA) then
      begin
        if Length(PDetail.CellValues[j, i]) < NbrC_DSA then
          PDetail.CellValues[j, i] := lpad(PDetail.CellValues[j, i],NbrC_DSA);
        if (not IsNumeric(PDetail.CellValues[j, i])) or (length(PDetail.CellValues[j, i]) <> NbrC_DSA) then
        begin
          LastErrorMsg := TraduireMemoire('Erreur :'#10#13'Le code doit être composé de '+IntToStr(NbrC_DSA)+' chiffres.');
          SetFocusControl('DETAILMINI');
          PDETAIL.Row := i;
          PDETAIL.Col := j;
          LastError := 131;
          Exit;
        end;
      end;
      if  (PDetail.ColFormats[j] = FormatC_COD) then
      begin
        if Length(PDetail.CellValues[j, i]) < NbrC_COD then
          PDetail.CellValues[j, i] := lpad(PDetail.CellValues[j, i],NbrC_COD);
        if (not IsNumeric(PDetail.CellValues[j, i])) or (length(PDetail.CellValues[j, i]) <> NbrC_COD) then
        begin
          LastErrorMsg := TraduireMemoire('Erreur :'#10#13'Le code doit être composé de '+IntToStr(NbrC_COD)+' chiffres.');
          SetFocusControl('DETAILMINI');
          PDETAIL.Row := i;
          PDETAIL.Col := j;
          LastError := 132;
          Exit;
        end;
      end;
    end;
    if (not boDonneesSalarie) and (PDetail.CellValues[2, i] = '') then
    begin
      LastErrorMsg := TraduireMemoire('Erreur :'#10#13'Vous devez renseigner le libellé pour chaque valeur du critères.');
      SetFocusControl('DETAILMINI');
      PDETAIL.Row := i;
      PDETAIL.Col := 2;
      LastError := 133;
      Exit;
    end;
    if PDetail.CellValues[3, i] = '' then
    begin
      LastErrorMsg := TraduireMemoire('Erreur :'#10#13'Vous devez renseigner le résultat pour chaque combinaison de critères.');
      SetFocusControl('DETAILMINI');
      PDETAIL.Row := i;
      PDETAIL.Col := 3;
      LastError := 133;
      Exit;
    end;
  end;
  if NbCriteres >= 1 then
  begin
    for i := 1 to PDetail.rowcount-1 do
    begin
      if PDetail.CellValues[0, i] = '' then
      begin
        LastErrorMsg := TraduireMemoire('Erreur :'#10#13'Vous devez renseigner la valeur du premier critère.');
        SetFocusControl('DETAILMINI');
        PDETAIL.Row := i;
        PDETAIL.Col := 0;
        LastError := 14;
        Exit;
      end;
    end;
  end;
  if NbCriteres >= 2 then
  begin
    for i := 1 to PDetail.rowcount-1 do
    begin
      if PDetail.CellValues[1, i] = '' then
      begin
        LastErrorMsg := TraduireMemoire('Erreur :'#10#13'Vous devez renseigner la valeur du second critère.');
        SetFocusControl('DETAILMINI');
        PDETAIL.Row := i;
        PDETAIL.Col := 1;
        LastError := 15;
        Exit;
      end;
    end;
  end;
  //Début PT2
  if (stCRITERE1 = '') or (stCRITERE1 = '0') then
  //PT10 SetControlText('PTE_CRITERE1B','0');
    Control_AbsCRITERE1B.SetControlText('0'); //PT10
  if (stCRITERE2 = '') or (stCRITERE2 = '0') then
    SetField('PTE_CRITERE2',0);
  //Fin PT2
  // Si l'update est validé, on force le boLoading pour éviter les questions
  // lors de la réinitialisation des champs qui a lieu a la fermeture de la fiche en CWAS
  boLoading := True;

  //DEB PT6
  if (DS.State = dsinsert) then
    DerniereCreate := GetField('PTE_CODTABL');
  LeStatut := DS.State;
  //FIN PT6
end ;

procedure TOM_TABLEDIMENT.OnAfterUpdateRecord ;
var
  even: boolean;
  LaTable,lecode, LeLibelle : String;
begin
  Inherited ;
  SaveGrid;
  if not boLoading then UpdateVisibilites;
  //DEB PT6
  LaTable := 'TABLEDIMENT';
  LeCode := 'PTE_CODTABL';
  LeLibelle := 'PTE_LIBELLE';
  Trace := TStringList.Create ;
  even := IsDifferent(dernierecreate,Latable,LeCode,LeLibelle,Trace,TFFicheListe(Ecran),LeStatut);
  FreeAndNil (Trace);
  //FIN PT6
end ;

procedure TOM_TABLEDIMENT.OnAfterDeleteRecord ;
begin
  Inherited ;
  ExecuteSQL('DELETE FROM TABLEDIMDET WHERE ##PTD_PREDEFINI## '+Key);
  if not boLoading then UpdateVisibilites;
end ;

procedure TOM_TABLEDIMENT.OnLoadRecord ;
var
  DateValid : TDateTime;
begin
  Inherited ;
//  if getField('PTE_CRITERE1') = 0 then SetField('PTE_CRITERE1B',0);
//  if getField('PTE_CRITERE2') = 0 then SetField('PTE_CRITERE2',0);
  boLoading := True;
  //Debut PT10
  UpdateCriterePRNature(True);
//  if GetField('PTE_NATURETABLE') = 'VAR' then //PT3
//  begin
//    Control_AbsCRITERE1B.ControlToAbstract := Control_CRITERE1B_;  //PT10
//    Control_CRITERE1B.Vide := False;
//    Control_CRITERE1B.Plus := '';
//    Control_CRITERE1B.DataType := 'PGVARIABLE';
//    SetControlText('PTE_CRITERE1B', lpad(GetField('PTE_CRITERE1'),4));
//    SetControlText('TPTE_CRITERE1',TraduireMemoire('Variable de paie'));
//  end else if GetField('PTE_NATURETABLE') = 'DSA' then begin
//    Control_AbsCRITERE1B.ControlToAbstract := Control_CRITERE1B;  //PT10
//    Control_CRITERE1B.Vide := True;
//    Control_CRITERE1B.Plus := 'AND ((PAI_IDENT <> 45) AND (PAI_IDENT <> 46) AND (PAI_IDENT <> 220) AND (PAI_IDENT <> 221))';
//    Control_CRITERE1B.DataType := 'PGCHAMPCRITERETABLE';
//    SetControlText('PTE_CRITERE1B',GetField('PTE_CRITERE1'));
//    SetControlText('TPTE_CRITERE1',TraduireMemoire('Critère 1'));
//  end;
  //Fin PT10

  UpdateValues;
  try
    DateValid := AGLStrToDate(getControlText('PTE_DTVALID'));
  except
    DateValid := DebutDeMois(now);
  end;
  Key := '(     PTD_CODTABL   ="'+stCODTABL  +'"'
        +'  AND PTD_NODOSSIER ="'+getControlText('PTE_NODOSSIER')+'"'  //@ PgRendNoDossier()
        +'  AND PTD_DTVALID   ="'+UsDateTime(DateValid)+'"'
        +'  AND PTD_NIVSAIS   ="'+stNIVSAIS  +'"'
        +'  AND PTD_VALNIV    ="'+stVALNIV +'"'
        +'  AND PTD_PREDEFINI ="'+stPREDEFINI+'")';
  LoadGrid;
  boLoading := False;
  UpdateVisibilites;

  if (DS.State in [dsInsert]) then
  else
    DerniereCreate := ''; //PT6

  // Gestion des concepts sur tables dynamiques
{  LectureSeule := FALSE;
  if (Getfield('PCT_PREDEFINI') = 'CEG') then
  begin
    LectureSeule := (CEG = False);
    if not CEG then PaieLectureSeule(TFFiche(Ecran), True);
    SetControlEnabled('BDelete', CEG);
  end
  else
    if (Getfield('PCT_PREDEFINI') = 'STD') then
    begin
      LectureSeule := (STD = False);
      if not STD then PaieLectureSeule(TFFiche(Ecran), True);
      SetControlEnabled('BDelete', DOS);
    end
    else
      if (Getfield('PCT_PREDEFINI') = 'DOS') then
      begin
        LectureSeule := (DOS = False);
        SetControlEnabled('BDelete', DOS);
      end;

   PaieConceptPlanPaie(Ecran); // Pour rendre invisible les boutons insert et dupliquer
}
end ;

procedure TOM_TABLEDIMENT.OnChangeField ( F: TField ) ;
begin
  Inherited ;
end ;

procedure TOM_TABLEDIMENT.OnArgument ( S: String ) ;
var
  Arg : String;
begin
  MultiLineOnglet := V_PGI.MultiLineOnglet; //PT8
  V_PGI.MultiLineOnglet := False; //PT8
  boControlTableAlreadyExist := False;
  boAlerteInitGrid := False;
  boLoading := True;
  Inherited;
  Arg := S;
  Action := Trim(ReadTokenPipe(Arg, ';')); //On récupère le type d'action
  if (Action = 'ACTION=MODIFICATION') then
    Action := 'MODIFICATION'
  else
    Action := 'CREATION';
  //On récupère la nature de la table sélectionnée dans le mul
  Default_NATURETABLE := Trim(ReadTokenPipe(Arg, ';'));
  //On récupère le niveau de saisie sélectionnée dans le mul
  Default_NIVSAIS     := Trim(ReadTokenPipe(Arg, ';'));
  //On récupère la valeur du niveau de saisie sélectionnée dans le mul
  Default_VALNIV      := Trim(ReadTokenPipe(Arg, ';'));
  if Default_VALNIV = '...' then Default_VALNIV := '';
  Default_PREDEFINI   := 'STD';
  Default_DTVALID := AGLDateToStr(DebutDeMois(now));
  //On récupère les composants
  PDetail:= (GetControl('DETAILMINI')   as THGrid);
  PDetail.ColAligns[3] := taCenter;

  BtnIns := (GetControl('BINSERT_LINE') as TToolBarButton97);
  BtnDel := (GetControl('BDEL_LINE')    as TToolBarButton97);
  BtnAdd := (GetControl('BADD_LINE')    as TToolBarButton97);
{$IFNDEF EAGLCLIENT}
  Control_TABLE          := (GetControl('PTE_CODTABL')         as THDBEdit);
  Control_NATURETABLE    := (GetControl('PTE_NATURETABLE')     as THDBValComboBox);
  Control_NIVSAIS        := (GetControl('PTE_NIVSAIS')         as THDBValComboBox);
  Control_TYPERESULTAT   := (GetControl('PTE_TYPERESULTAT')    as THDBValComboBox);
  Control_PREDEFINI      := (GetControl('PTE_PREDEFINI')       as THDBValComboBox);
  Control_CRITERE2       := (GetControl('PTE_CRITERE2')        as THDBValComboBox);
  Control_VALNIV         := (GetControl('PTE_VALNIVB')         as THEdit);
{$ELSE}
  Control_TABLE          := (GetControl('PTE_CODTABL')         as THEdit);
  Control_NATURETABLE    := (GetControl('PTE_NATURETABLE')     as THValComboBox);
  Control_NIVSAIS        := (GetControl('PTE_NIVSAIS')         as THValComboBox);
  Control_TYPERESULTAT   := (GetControl('PTE_TYPERESULTAT')    as THValComboBox);
  Control_PREDEFINI      := (GetControl('PTE_PREDEFINI')       as THValComboBox);
  Control_CRITERE2       := (GetControl('PTE_CRITERE2')        as THValComboBox);
  Control_VALNIV         := (GetControl('PTE_VALNIVB')         as THEdit);
{$ENDIF}
  Control_CRITERE1B      := (GetControl('PTE_CRITERE1B')       as THValComboBox);
  Control_COMBINCRITERE  := (GetControl('BTN_GENERATIONAUTO')  as TToolBarButton97);
  //Debut PT10
  Control_CRITERE1B_     := (GetControl('PTE_CRITERE1B_')      as THEdit);
  Control_AbsCRITERE1B   := TControlAbstractComboEdit.Create(self);
  Control_AbsCRITERE1B.ControlToAbstract := Control_CRITERE1B;
  //Fin PT10
  TobTableDynamique := TOB.Create('TABLEDIMENT mère', nil, -1);
  AccesPredefini('TOUS', CEG, STD, DOS);
  if Action = 'MODIFICATION' then
  begin
    boAlerteInitGrid := True;
  end else if Action = 'CREATION' then
  begin
  end;
  AffecteEvenements;
end ;

procedure TOM_TABLEDIMENT.OnClose ;
begin
  TobTableDynamique.Free;
  Inherited ;
  V_PGI.MultiLineOnglet := MultiLineOnglet; //PT8
  FreeAndNil(Control_AbsCRITERE1B);
end ;

procedure TOM_TABLEDIMENT.OnCancelRecord ;
begin
  Inherited ;
  if not boLoading then UpdateVisibilites;
end ;

procedure TOM_TABLEDIMENT.OnChangeNATURETABLE(Sender: TObject);
var
  SAVboAlerteInitGrid : Boolean;
begin
  if not boLoading then
  begin
    if not isBackNATURETABLE then
    begin
      if (not boAlerteInitGrid) or ((boAlerteInitGrid) and (PGIAsk( TraduireMemoire('La nature de la table a été modifiée.'#10#13'Toutes les données de la grille vont être supprimées.'#10#13#10#13'Voulez-vous continuer ?'),Ecran.Caption)
         = mrYes)) then
      begin
        SAVboAlerteInitGrid := boAlerteInitGrid;
        boAlerteInitGrid := False;
        isBackTYPERESULTAT := True;
        setControlText('PTE_TYPERESULTAT','');
        isBackTYPERESULTAT := False;
        ExecuteSQL('DELETE FROM TABLEDIMDET WHERE ##PTD_PREDEFINI## '+Key);
        InitGrid;
        setControlText('PTE_VALNIVB','...');
        boAlerteInitGrid := SAVboAlerteInitGrid;
      end else begin
        isBackNATURETABLE := True;
        SetControlText('PTE_NATURETABLE',BackNATURETABLE);
      end;
    end else begin
      isBackNATURETABLE := False;
    end;
  end;
  if not boLoading then UpdateVisibilites;
end;

procedure TOM_TABLEDIMENT.OnChangeNIVSAIS(Sender: TObject);
begin
  if not boLoading then UpdateVisibilites;
end;

procedure TOM_TABLEDIMENT.OnChangeTYPERESULTAT(Sender: TObject);
var
  i : integer;
begin
  if not boLoading then
  begin
    if not isBackTYPERESULTAT then
    begin
      if (not boAlerteInitGrid)
         or ((boAlerteInitGrid) and (PGIAsk( TraduireMemoire('Le type de résultat a été modifié.'#10#13'Les données de la colonne résultat de la grille vont être supprimées.'#10#13#10#13'Voulez-vous continuer ?'),Ecran.Caption)
         = mrYes)) then
      begin
        for i := 1 to PDetail.rowcount-1 do
        begin
          PDetail.Cells[3, i] := '';
          PDetail.CellValues[3, i] := '';
        end;
      end else begin
        isBackTYPERESULTAT := True;
        SetControlText('PTE_TYPERESULTAT',BackTYPERESULTAT);
      end;
    end else begin
      isBackTYPERESULTAT := False;
    end;
  end;
  if not boLoading then UpdateVisibilites;
end;

procedure TOM_TABLEDIMENT.OnEnterTYPERESULTAT(Sender: TObject);
begin
  UpdateValues;
  //On sauvegarde le type de résultata avant modif, pour pouvoir le remettre
  // si l'utilisateur annule la réinitialisation des données de la grille
  // lors du changement de type de résultat
  BackTYPERESULTAT := stTYPERESULTAT;
end;

procedure TOM_TABLEDIMENT.BtnAddClick(Sender: TObject);
begin
  boAlerteInitGrid := True;
  PDetail.RowCount := PDetail.RowCount + 5;
  PDetail.Col := 0;
end;

procedure TOM_TABLEDIMENT.BtnDelClick(Sender: TObject);
begin
  boAlerteInitGrid := True;
  if PDetail.Row <> 0 then PDETAIL.DeleteRow(PDETAIL.Row);
end;

procedure TOM_TABLEDIMENT.BtnInsClick(Sender: TObject);
begin
  boAlerteInitGrid := True;
  PDETAIL.InsertRow(PDETAIL.Row);
end;

// Les données de ACol et ARow contiennent les valeurs de la cellule
// précédente, et pas celles de la cellule dans laquelle
// on entre. Il faut utiliser PDetail.col et PDetail.row
procedure TOM_TABLEDIMENT.PDetailCellEnter(Sender: TObject; var ACol,
  ARow: Integer; var Cancel: Boolean);
var
  dif : Integer;
begin
  { Si on est dans une case invisible, on se déplace dans la case visible suivante (jusqu'a 3 cases plus loin).
    Si celle-ci est vide, on annule le déplacement.
    Ca permet de ne pas gêner les tabulations dans la grille. //PT4 FQ 14209 }
  if PDetail.colwidths[PDetail.col]<= 0 then
  begin
    dif := PDetail.col-ACol;
    if PDetail.colwidths[PDetail.col + dif]> 0 then
      PDetail.col := PDetail.col + dif
    else if PDetail.colwidths[PDetail.col + 2*dif]> 0 then
      PDetail.col := PDetail.col + 2*dif
    else if PDetail.colwidths[PDetail.col + 3*dif]> 0 then
      PDetail.col := PDetail.col + 3*dif
    else
      Cancel := True;
  end;  
  //on force l'update pour que les modifications des données de la grille soient prises
  // en compte même si on a pas modifié les données de l'entête. (permet de déclanché l'évènement Update lors de la validation)
  ForceUpdate;
  PDetail.ElipsisButton := false;
  if (PDetail.col = 3) then
  begin
    if (Getfield('PTE_TYPERESULTAT') = 'VAR') then
    begin
      PDetail.ElipsisButton := True;
    end else if (Getfield('PTE_TYPERESULTAT') = 'ELT') then
    begin
      PDetail.ElipsisButton := True;
    end;
  end;
end;

procedure TOM_TABLEDIMENT.PDetailElipsisClick(Sender: TObject);
begin
  if (Getfield('PTE_TYPERESULTAT') = 'VAR') then
    LookupList(PDetail,TraduireMemoire('Variable de Paie'),'VARIABLEPAIE','DISTINCT PVA_VARIABLE','PVA_LIBELLE',WhereVar,'PVA_VARIABLE',TRUE,-1);
  if (Getfield('PTE_TYPERESULTAT') = 'ELT') then
    LookupList(PDetail,TraduireMemoire('Elément national'),'ELTNATIONAUX','DISTINCT PEL_CODEELT','PEL_LIBELLE',WhereElt,'PEL_CODEELT',TRUE,-1);
end;

procedure TOM_TABLEDIMENT.InitGrid;
var
  i, j: Integer;
begin
  for j := 0 to PDetail.colcount-1 do
  begin
    PDetail.ColFormats[j] := '';
    for i := 1 to PDetail.rowcount-1 do
    begin
      PDetail.Cells[j, i] := '';
      PDetail.CellValues[j, i] := '';
    end;
  end;
  if Assigned(PDETAIL.ValCombo) then
  begin
    PDETAIL.Col := 3;
    PDETAIL.ValCombo.Hide;
    PDETAIL.ValCombo.Free;
  end;
  PDetail.Col := 0;
  PDetail.RowCount := 6;
  PDetail.Row := 1;
end;

procedure TOM_TABLEDIMENT.LoadGrid;
begin
  TobTableDynamique.ClearDetail;
  TobTableDynamique.LoadDetailDBFromSQL('TABLEDIMDET mère','SELECT * FROM TABLEDIMDET WHERE ##PTD_PREDEFINI## '+Key);
  //Permet d'affecter les combobox à la grille pour que l'affectation de cellValues
  // dans le putGridDetail soit interprétée comme un code, et non comme une valeur
  UpdateVisibilites;
  TobTableDynamique.PutGridDetail(PDETAIL,False,True,ListeDesChampsDuGrid,False);
end;

function TOM_TABLEDIMENT.SetDetailsFieldsNotInGrid(T : TOB) : integer;
var
  Value : String;
begin
  UpdateValues;
  Value := stCODTABL                      +'##S&p@r@t3vr##'
          +getControlText('PTE_NODOSSIER')+'##S&p@r@t3vr##'
          +getControlText('PTE_DTVALID')  +'##S&p@r@t3vr##'
          +stNIVSAIS                      +'##S&p@r@t3vr##'
          +stVALNIV                       +'##S&p@r@t3vr##'
          +stPREDEFINI;
  T.LoadFromSt('PTD_CODTABL,PTD_NODOSSIER,PTD_DTVALID,PTD_NIVSAIS,PTD_VALNIV,PTD_PREDEFINI',//,PTD_VALCRIT2',
                                         ',', Value, '##S&p@r@t3vr##');
end;

procedure TOM_TABLEDIMENT.SaveGrid;
begin
  TobTableDynamique.GetGridDetail(PDETAIL,PDETAIL.ColCount,'TABLEDIMDET',ListeDesChampsDuGrid,True);
  //On crée les champs qui ne sont pas dans la grille pour toutes les tob filles.
  TobTableDynamique.Detail[0].AddChampSup('PTD_CODTABL',True);
  TobTableDynamique.Detail[0].AddChampSup('PTD_PREDEFINI',True);
  TobTableDynamique.Detail[0].AddChampSup('PTD_NODOSSIER',True);
  TobTableDynamique.Detail[0].AddChampSup('PTD_DTVALID',True);
  TobTableDynamique.Detail[0].AddChampSup('PTD_NIVSAIS',True);
  TobTableDynamique.Detail[0].AddChampSup('PTD_VALNIV',True);
  TobTableDynamique.ParcoursTraitement([''],[''],False,SetDetailsFieldsNotInGrid);
//  TobDebug(TobTableDynamique);
  BeginTrans;
  try
    ExecuteSQL('DELETE FROM TABLEDIMDET WHERE ##PTD_PREDEFINI## '+Key);
    if TobTableDynamique.InsertOrUpdateDB then
      CommitTrans
    else
      RollBack;
  except
    RollBack;
  end;
end;

procedure TOM_TABLEDIMENT.InitEntete;
begin
  //DEB PT6
  if (not STD) then
    SetControlText('PTE_PREDEFINI','DOS')
  else
  //FIN PT6
    SetControlText('PTE_PREDEFINI',Default_PREDEFINI);
  SetControlText('PTE_NATURETABLE',Default_NATURETABLE);
  SetControlText('PTE_NIVSAIS',Default_NIVSAIS);
  SetControlText('PTE_VALNIVB',Default_VALNIV);
  SetControlText('PTE_DTVALID',Default_DTVALID);
end;

procedure TOM_TABLEDIMENT.OnChangePREDEFINI(Sender: TObject);
var
  SAVboAlerteInitGrid : Boolean;
begin
  if not boLoading then
  begin
    if not isBackPREDEFINI then
    begin
      if (not boAlerteInitGrid) or ((boAlerteInitGrid) and (PGIAsk( TraduireMemoire('Le type de prédéfini a été modifié.'#10#13'Toutes les données de la grille vont être supprimées.'#10#13#10#13'Voulez-vous continuer ?'),Ecran.Caption)
         = mrYes)) then
      begin
        SAVboAlerteInitGrid := boAlerteInitGrid;
        boAlerteInitGrid := False;
        if (GetControlText('PTE_PREDEFINI') = 'CEG') and (CEG = FALSE) then
        begin
          PGIBox('Vous ne pouvez pas créer de prédéfini CEGID.', Ecran.Caption);
          isBackPREDEFINI := True;
          SetControlText('PTE_PREDEFINI','DOS');
        end;
        if (GetControlText('PTE_PREDEFINI') = 'STD') and (STD = FALSE) then
        begin
          PGIBox('Vous ne pouvez pas créer de prédéfini Standard.', Ecran.Caption);
          isBackPREDEFINI := True;
          SetControlText('PTE_PREDEFINI','DOS');
        end;
        InitGrid;
        boAlerteInitGrid := SAVboAlerteInitGrid;
      end else begin
        isBackPREDEFINI := True;
        SetControlText('PTE_PREDEFINI',BackPREDEFINI);
      end;
    end else begin
      isBackPREDEFINI := False;
    end;
  end;
//Debut PT9
  if (GetControlText('PTE_PREDEFINI') = 'CEG') or (GetControlText('PTE_PREDEFINI') = 'STD') then
    SetControlText('PTE_NODOSSIER','000000')
  else
    SetControlText('PTE_NODOSSIER', PGRendNoDossier);
//Fin PT9
  if not boLoading then UpdateVisibilites;
end;

procedure TOM_TABLEDIMENT.OnEnterNATURETABLE(Sender: TObject);
begin
  UpdateValues;
  //On sauvegarde la nature de la table avant modif, pour pouvoir le remettre
  // si l'utilisateur annule la réinitialisation des données de la grille
  // lors du changement de type de résultat
  BackNATURETABLE := stNATURETABLE;
end;

procedure TOM_TABLEDIMENT.OnEnterPREDEFINI(Sender: TObject);
begin
  UpdateValues;
  //On sauvegarde le type de prédéfini avant modif, pour pouvoir le remettre
  // si l'utilisateur annule la réinitialisation des données de la grille
  // lors du changement de type de prédéfini
  BackPREDEFINI := stPREDEFINI;
end;

function TOM_TABLEDIMENT.FindTabletteName(IDENT : integer) : String;
var
  Q : TQuery;
  stQ : String;
begin
  result := '';
  try
    stQ := 'SELECT PAI_LIENASSOC FROM PAIEPARIM WHERE ##PAI_PREDEFINI## AND PAI_IDENT = '+intToStr(IDENT);
    if (stPREDEFINI = 'CEG') or (stPREDEFINI = 'STD') then
      stQ := stQ+' AND PAI_PGTYPEUTILIS = "COM" ';
    Q := OpenSQL(stQ, True, 1);
    if not Q.Eof then
      Result := Q.FindField('PAI_LIENASSOC').AsString;
    //La tablette PGLIBCOEFFICIENT indiquée dans PAIEPARIM n'est pas correcte,
    // on utilise PGCOEFFICIENT à la place
    if Result = 'PGLIBCOEFFICIENT' then Result := 'PGCOEFFICIENT';
    Ferme(Q);
  except
    result := '';
  end;
end;

procedure TOM_TABLEDIMENT.OnExitCRITERE1B(Sender: TObject);
var
  SAVboAlerteInitGrid : Boolean;
begin
  if    (not (stNATURETABLE = 'VAR')) and (not (stNATURETABLE = 'ELT'))  //PT11
    and (not boLoading )
//PT10    AND (BackCRITERE1 <> GetControlText('PTE_CRITERE1B')) then //PT3
    AND (BackCRITERE1 <>  Control_AbsCRITERE1B.GetControlText) then //PT3  //PT10
  begin
    if not isBackCRITERE1 then
    begin
      if (not boAlerteInitGrid) or ((boAlerteInitGrid) and (PGIAsk( TraduireMemoire('Le premier type de critère a été modifié.'#10#13'Toutes les données de la grille vont être supprimées.'#10#13#10#13'Voulez-vous continuer ?'),Ecran.Caption)
         = mrYes)) then
      begin
        SAVboAlerteInitGrid := boAlerteInitGrid;
        boAlerteInitGrid := False;
        InitGrid;
        boAlerteInitGrid := SAVboAlerteInitGrid;
      end else begin
        isBackCRITERE1 := True;
        //PT10 SetControlText('PTE_CRITERE1B',BackCRITERE1);
        Control_AbsCRITERE1B.ControlText := BackCRITERE1; //PT10
      end;
    end else begin
      isBackCRITERE1 := False;
    end;
  end;
//PT10  if GetControlText('PTE_CRITERE1B') <> '' then    //PT3
  if       Control_AbsCRITERE1B.ControlText <> '' then    //PT3   //PT10
    SetField('PTE_CRITERE1',Control_AbsCRITERE1B.ControlText)  //PT10
  else
    SetField('PTE_CRITERE1',0);
  if not boLoading then UpdateVisibilites;
end;

function TOM_TABLEDIMENT.GetNbCriteres: integer;
begin
  if (stCRITERE2 <> '0') and (stCRITERE2 <> '') then
    result := 2
  else if (stCRITERE1 <> '0') and (stCRITERE1 <> '') then
    result := 1
  else
    result := 0;
  if stNATURETABLE = 'VAR' then result := 0; //PT3
  if stNATURETABLE = 'ELT' then result := 0; //PT11
end;

procedure TOM_TABLEDIMENT.UpdateVisibilites;
var
  tempst : string;
  tempint : integer;
  TempPlusCritere1, TempPlusCritere2 : String;
begin
// PT8 Le problème est causé par quelque chose entre ici et la ligne 1120
  //Debut PT10
  UpdateCriterePRNature(False);
//  if GetField('PTE_NATURETABLE') = 'VAR' then //PT3
//  begin
//    if (Control_CRITERE1B.DataType <> 'PGVARIABLE') then
//    begin
//      Control_AbsCRITERE1B.ControlToAbstract := Control_CRITERE1B_;  //PT10
//      Control_CRITERE1B.Vide := False;
//      Control_CRITERE1B.Plus := '';
//      Control_CRITERE1B.DataType := 'PGVARIABLE';
//      SetControlText('TPTE_CRITERE1',TraduireMemoire('Variable de paie'));
//    end;
//  end else if GetField('PTE_NATURETABLE') = 'DSA' then begin
//    if (Control_CRITERE1B.DataType <> 'PGCHAMPCRITERETABLE') then
//    begin
//      Control_AbsCRITERE1B.ControlToAbstract := Control_CRITERE1B;  //PT10
//      Control_CRITERE1B.Vide := True;
//      Control_CRITERE1B.Plus := ' AND ((PAI_IDENT <> 45) AND (PAI_IDENT <> 46) AND (PAI_IDENT <> 220) AND (PAI_IDENT <> 221))';
//      Control_CRITERE1B.DataType := 'PGCHAMPCRITERETABLE';
//      SetControlText('TPTE_CRITERE1',TraduireMemoire('Critère 1'));
//    end;
//  end;
  //Fin PT10

  UpdateValues;
  TempPlusCritere1 := '';

  boLoading := True; //On force le loading pour ne pas déclancher les évènements de OnChange en boucle
  if stNIVSAIS = 'ETB' then    //Niveau etablisement
  begin
    setControlProperty('PTE_VALNIVB','Vide',True);
    setControlProperty('PTE_VALNIVB','VideString','<<Tous les établissements>>');
    setControlProperty('PTE_VALNIVB','DataType','TTETABLISSEMENT');
    setControlProperty('TPTE_VALNIV','Caption','Etablissement');
    TempPlusCritere1 :=' AND ((PAI_IDENT <> 45) AND (PAI_IDENT <> 46) AND (PAI_IDENT <> 220) AND (PAI_IDENT <> 221))';
    TempPlusCritere2 :=' AND ((PAI_IDENT <> 45) AND (PAI_IDENT <> 46) AND (PAI_IDENT <> 220) AND (PAI_IDENT <> 221))';
  end else if stNIVSAIS = 'CON' then   //Niveau Convention collective
  begin
    setControlProperty('PTE_VALNIVB','Vide',False);
    setControlProperty('PTE_VALNIVB','VideString','');
    setControlProperty('PTE_VALNIVB','DataType','PGCONVENTIONS');
    setControlProperty('TPTE_VALNIV','Caption','Convention');
    TempPlusCritere1 :='';
    TempPlusCritere2 :='';
  end else if stNIVSAIS = 'GEN' then   //Niveau Général
  begin
    setControlProperty('PTE_VALNIVB','DataType','');
    SetControlText('PTE_VALNIVB','...');
    TempPlusCritere1 :=' AND ((PAI_IDENT <> 45) AND (PAI_IDENT <> 46) AND (PAI_IDENT <> 220) AND (PAI_IDENT <> 221))';
    TempPlusCritere2 :=' AND ((PAI_IDENT <> 45) AND (PAI_IDENT <> 46) AND (PAI_IDENT <> 220) AND (PAI_IDENT <> 221))';
  end;

  if boDonneesSalarie then
    setControlProperty('PTE_TYPERESULTAT','Plus',' and (CO_LIBRE = "DSA" or CO_LIBRE = "")')
  else
    setControlProperty('PTE_TYPERESULTAT','Plus',' and CO_LIBRE = ""');
  if stPREDEFINI <> '' then
  begin
    setControlProperty('PTE_NIVSAIS','Plus',' and co_libre like "%'+stPREDEFINI+'%"');
    if (stPREDEFINI = 'CEG') or (stPREDEFINI = 'STD') then begin
      TempPlusCritere1 := TempPlusCritere1 + ' AND PAI_PGTYPEUTILIS = "COM"';
      TempPlusCritere2 := TempPlusCritere2 + ' AND PAI_PGTYPEUTILIS = "COM"';
    end;
  end;
  if (stNATURETABLE = 'VAR') or (stNATURETABLE = 'ELT') then //PT3 //PT11
  begin
    TempPlusCritere1 := '';
  end else begin
    //PT10 setControlProperty('PTE_CRITERE1B','Plus',TempPlusCritere1);
    Control_AbsCRITERE1B.Plus := TempPlusCritere1; //PT10
    setControlProperty('PTE_CRITERE2','Plus',TempPlusCritere2);
  end;
//PT8 le problème est causé par quelque chose entre ici et la ligne 1053
  if boDonneesSalarie then
  begin
    SetControlText('PTE_SENSINT','');
  end else begin
//PT5    SetControlText('PTE_NIVSAIS','GEN');
    if (stNATURETABLE <> 'VAR') AND (stNATURETABLE <> 'ELT') then //PT3 //PT11
//PT10      SetControlText('PTE_CRITERE1B','');
      Control_AbsCRITERE1B.ControlText := ''; //PT10
    SetControlText('PTE_CRITERE2','');
  end;
  if not boCritere1AgeAnciennete then
  begin
    SetControlText('PTE_SENS1','');
    SetControlText('PTE_TYPECRITERE1','');
  end;
  if not boCritere2AgeAnciennete then
  begin
    SetControlText('PTE_SENS2','');
    SetControlText('PTE_TYPECRITERE2','');
  end;
  if not boNiveauEtab_ou_Convention then
  begin
    SetControlText('PTE_VALNIVB','...');
  end;

  boLoading := False;
  //  On remet a jour les valeurs, parce que les changement de tablettes ont pu
  // changer certaines d'entre elles.
  UpdateValues;

  UpdateAutorisation;
  setControlProperty('PTE_NIVSAIS',     'Enabled',Autorisation);
//PT10  setControlProperty('PTE_CRITERE1B',   'Enabled',Autorisation);
  Control_AbsCRITERE1B.Enabled := Autorisation; //PT10
  setControlProperty('PTE_CRITERE2',    'Enabled',Autorisation);
  setControlProperty('PTE_NATURETABLE', 'Enabled',Autorisation);
  setControlProperty('PTE_TYPERESULTAT','Enabled',Autorisation);
  setControlProperty('PTE_LIBELLE',     'Enabled',Autorisation);
  setControlProperty('BINSERT_LINE',    'Enabled',Autorisation);
  setControlProperty('BADD_LINE',       'Enabled',Autorisation);
  setControlProperty('BDEL_LINE',       'Enabled',Autorisation);
  setControlProperty('BDelete',         'Visible',Autorisation and ((ds.state = dsBrowse) or (ds.state = dsEdit)));
  setControlProperty('BDelete',         'Enabled',Autorisation and ((ds.state = dsBrowse) or (ds.state = dsEdit)));
  setControlProperty('BInsert',         'Visible',(ds.state = dsBrowse) or (ds.state = dsEdit));
  setControlProperty('BInsert',         'Enabled',(ds.state = dsBrowse) or (ds.state = dsEdit));
  SetControlVisible('BDelete',Autorisation);
  SetControlVisible('BInsert',True);
  if boDonneesSalarie then
  begin
    tempst := RechDom(Control_AbsCRITERE1B.DataType,stCRITERE1,False);  //PT10
    PDETAIL.CellValues[0,0] := tempst;
    PDETAIL.Cells[0,0] := tempst;
    tempst := RechDom(Control_CRITERE2.DataType,stCRITERE2,False);
    PDETAIL.CellValues[1,0] := tempst;
    PDETAIL.Cells[1,0] := tempst;
    if NbCriteres = 0 then
    begin
      tempint := (PDETAIL.Width-24) div 10;
      PDETAIL.ColWidths[0] := 2*tempint;
      PDETAIL.ColWidths[1] := -1;
      PDETAIL.ColWidths[2] := 5*tempint;
      PDETAIL.ColWidths[3] := 3*tempint;
// la propriété ColEditable ne fonction pas en CWAS, alors on utilise FixedRow pour
// avoir un fonctionnement similaire.
      PDETAIL.ColEditables[0] := Autorisation;
      PDETAIL.ColEditables[1] := False;
      PDETAIL.ColEditables[2] := Autorisation;
      PDETAIL.ColEditables[3] := Autorisation;
      if Autorisation then
      begin
        if not (goEditing in PDETAIL.Options) then
          PDETAIL.Options := PDETAIL.Options + [goEditing]
      end else begin
        if goEditing in PDETAIL.Options then
          PDETAIL.Options := PDETAIL.Options - [goEditing]
      end;
    end;
    if NbCriteres = 1 then
    begin
      tempint := (PDETAIL.Width-24) div 10;
      PDETAIL.ColWidths[0] := 7*tempint;
      PDETAIL.ColWidths[1] := -1;
      PDETAIL.ColWidths[2] := -1;
      PDETAIL.ColWidths[3] := 3*tempint;
      PDETAIL.ColEditables[0] := Autorisation;
      PDETAIL.ColEditables[1] := False;
      PDETAIL.ColEditables[2] := False;
      PDETAIL.ColEditables[3] := Autorisation;
      if Autorisation then
      begin
        if not (goEditing in PDETAIL.Options) then
          PDETAIL.Options := PDETAIL.Options + [goEditing]
      end else begin
        if goEditing in PDETAIL.Options then
          PDETAIL.Options := PDETAIL.Options - [goEditing]
      end;
    end;
    if NbCriteres = 2 then
    begin
      tempint := (PDETAIL.Width-24) div 10;
      PDETAIL.ColWidths[0] := 4*tempint;
      PDETAIL.ColWidths[1] := 4*tempint;
      PDETAIL.ColWidths[2] := -1;
      PDETAIL.ColWidths[3] := 2*tempint;
      PDETAIL.ColEditables[0] := True;
      PDETAIL.ColEditables[1] := True;
      PDETAIL.ColEditables[2] := False;
      PDETAIL.ColEditables[3] := Autorisation;
      if Autorisation then
      begin
        if not (goEditing in PDETAIL.Options) then
          PDETAIL.Options := PDETAIL.Options + [goEditing]
      end else begin
        if goEditing in PDETAIL.Options then
          PDETAIL.Options := PDETAIL.Options - [goEditing]
      end;
    end;
    UpdatePlusCriteres;
    if ((stCRITERE1 <> '') and (stCRITERE1 <> '0'))
       and (NbCriteres >= 1)
       and not boCritere1AgeAnciennete then
    begin
      PDETAIL.ColLengths[0] := 17;
      PDetail.ColFormats[0]:='CB='+TabletteCritere1+'|'+PlusCritere1+'|';
    end else if ((stCRITERE1 <> '') and (stCRITERE1 <> '0'))
       and (NbCriteres >= 1)
       and boCritere1AgeAnciennete then
    begin
      PDETAIL.ColFormats[0] := FormatC_DSA;
      PDETAIL.ColLengths[0] := NbrC_DSA;
    end else begin
      PDETAIL.ColFormats[0] := '';
      PDETAIL.ColLengths[0] := 17;
    end;

    if ((stCRITERE2 <> '') and (stCRITERE2 <> '0'))
       and (NbCriteres >= 2)
       and not boCritere2AgeAnciennete then
    begin
      PDETAIL.ColLengths[1] := 17;
      PDetail.ColFormats[1]:='CB='+TabletteCritere2+'|'+PlusCritere2+'|';
    end else if ((stCRITERE2 <> '') and (stCRITERE2 <> '0'))
       and (NbCriteres >= 2)
       and boCritere2AgeAnciennete then
    begin
      PDETAIL.ColFormats[1] := FormatC_DSA;
      PDETAIL.ColLengths[1] := NbrC_DSA;
    end else begin
      PDETAIL.ColFormats[1] := '';
      PDETAIL.ColLengths[1] := 17;
    end;
  end else begin
    tempint := (PDETAIL.Width-24) div 10;
    PDETAIL.ColWidths[0] := 2*tempint;
    PDETAIL.ColWidths[1] := -1;
    PDETAIL.ColWidths[2] := 5*tempint;
    PDETAIL.ColWidths[3] := 3*tempint;
    PDETAIL.CellValues[0,0] := 'Code';
    PDETAIL.Cells[0,0] := 'Code';
    PDETAIL.ColEditables[0] := Autorisation;
    PDETAIL.ColEditables[1] := False;
    PDETAIL.ColEditables[2] := Autorisation;
    PDETAIL.ColEditables[3] := Autorisation;
    if (stNATURETABLE <> 'VAR') and (stNATURETABLE <> 'ELT') then //PT3 //PT11
    begin
      PDETAIL.ColFormats[0] := FormatC_COD;
      PDETAIL.ColLengths[0] := NbrC_COD;
    end else begin
      PDETAIL.ColFormats[0] := FormatC_VAR;
      PDETAIL.ColLengths[0] := NbrC_VAR;
//      PDETAIL.ColFormats[0] := '';
//      PDETAIL.ColLengths[0] := 17;
    end;
    try
      if Assigned(PDETAIL.ValCombo) then
      begin
        PDETAIL.Col := 3;
        PDETAIL.ValCombo.Hide;
        PDETAIL.ValCombo.Free;
      end;
    except
    end;
    if Autorisation then
    begin
        if not (goEditing in PDETAIL.Options) then
          PDETAIL.Options := PDETAIL.Options + [goEditing]
    end else begin
        if goEditing in PDETAIL.Options then
          PDETAIL.Options := PDETAIL.Options - [goEditing]
    end;
  end;
  if PDETAIL.RowCount < 2 then PDETAIL.RowCount := 2;
  PDETAIL.FixedRows := 1;
  setControlProperty('PTE_SENS1'         ,'Enabled',Autorisation);
  setControlProperty('PTE_SENS2'         ,'Enabled',Autorisation);
  setControlProperty('PTE_SENSINT'       ,'Enabled',Autorisation);
  setControlProperty('PTE_VALNIVB'       ,'Enabled',Autorisation);
  setControlProperty('PTE_TYPECRITERE1'  ,'Enabled',Autorisation);
  setControlProperty('PTE_TYPECRITERE2'  ,'Enabled',Autorisation);
  setControlProperty('BTN_GENERATIONAUTO','Enabled',Autorisation);
  setControlProperty('PTE_VALNIVB'       ,'Visible',boNiveauEtab_ou_Convention);
  setControlProperty('TPTE_VALNIV'       ,'Visible',boNiveauEtab_ou_Convention);
  setControlProperty('PTE_SENS1'         ,'Visible',boCritere1AgeAnciennete);
  setControlProperty('TPTE_SENS1'        ,'Visible',boCritere1AgeAnciennete);
  setControlProperty('PTE_SENS2'         ,'Visible',boCritere2AgeAnciennete);
  setControlProperty('TPTE_SENS2'        ,'Visible',boCritere2AgeAnciennete);
  setControlProperty('PTE_TYPECRITERE1'  ,'Visible',boCritere1AgeAnciennete);
  setControlProperty('TPTE_TYPECRITERE1' ,'Visible',boCritere1AgeAnciennete);
  setControlProperty('PTE_TYPECRITERE2'  ,'Visible',boCritere2AgeAnciennete);
  setControlProperty('TPTE_TYPECRITERE2' ,'Visible',boCritere2AgeAnciennete);
  setControlProperty('BTN_GENERATIONAUTO','Visible',boCOMBINCRITERE);
  setControlProperty('LBL_GENERATIONAUTO','Visible',boCOMBINCRITERE);
  setControlProperty('PTE_SENSINT'       ,'Visible',not boDonneesSalarie);
  setControlProperty('TPTE_SENSINT'      ,'Visible',not boDonneesSalarie);
  setControlProperty('PTE_NIVSAIS'       ,'Visible',True);  //PT5   boDonneesSalarie
  setControlProperty('TPTE_NIVSAIS'      ,'Visible',True);  //PT5
//PT10  setControlProperty('PTE_CRITERE1B'     ,'Visible',(boDonneesSalarie or (stNATURETABLE = 'VAR'))); //PT3
  Control_AbsCRITERE1B.Visible := (boDonneesSalarie or (stNATURETABLE = 'VAR') or (stNATURETABLE = 'ELT')); //PT3 //PT10 //PT11
  setControlProperty('TPTE_CRITERE1'     ,'Visible',(boDonneesSalarie or (stNATURETABLE = 'VAR') or (stNATURETABLE = 'ELT'))); //PT3   //PT11
  setControlProperty('PTE_CRITERE2'      ,'Visible',boDonneesSalarie);
  setControlProperty('TPTE_CRITERE2'     ,'Visible',boDonneesSalarie);

  //Gestion des concepts
  if   ((stPREDEFINI = 'CEG') and (not CEG))
    or ((stPREDEFINI = 'STD') and (not STD)) then
    PaieLectureSeule(TFFiche(Ecran), True);
  PaieConceptPlanPaie(Ecran);
end;

procedure TOM_TABLEDIMENT.UpdateValues;
var
  PrioriteTemp : Integer;
begin
  //Debut PT2
//PT10  if GetField('PTE_CRITERE1') = 0 then setControlText('PTE_CRITERE1B','');// else setControlText('PTE_CRITERE1',GetField('PTE_CRITERE1'));
  if GetField('PTE_CRITERE1') = 0 then Control_AbsCRITERE1B.ControlText := ''; // PT10
  if GetField('PTE_CRITERE2') = 0 then setControlText('PTE_CRITERE2','');
  //Fin PT2
  stNATURETABLE   := GetControlText('PTE_NATURETABLE');
  stCODTABL       := GetControlText('PTE_CODTABL');
  stNODOSSIER     := GetControlText('PTE_NODOSSIER');
  stNIVSAIS       := GetControlText('PTE_NIVSAIS');
  if not (VarType(GetField('PTE_CRITERE1')) = varNull ) then
    stCRITERE1    := intToStr(GetField('PTE_CRITERE1'))
  else
    stCRITERE1    := ''; // GetControlText('PTE_CRITERE1');            //PGLIBCOEFFICIENT -> PGCOEFFICIENT
  if not (VarType(GetField('PTE_CRITERE2')) = varNull ) then
    stCRITERE2    := intToStr(GetField('PTE_CRITERE2'))
  else
    stCRITERE2    := '';
//  stCRITERE2      := GetControlText('PTE_CRITERE2');
  stTYPECRITERE1  := GetControlText('PTE_TYPECRITERE1');
  stTYPECRITERE2  := GetControlText('PTE_TYPECRITERE2');
  stPREDEFINI     := GetControlText('PTE_PREDEFINI');
  if not boLoading then
  begin
    if GetControlText('PTE_VALNIVB')<> '' then
      SetControlText('PTE_VALNIV',GetControlText('PTE_VALNIVB'))
    else
      SetControlText('PTE_VALNIV','...');
  end else begin
    //On fait les affectations de datatype avant pour pouvoir mettre a jour la valeur de PTE_VALNIVB
    if stNIVSAIS = 'ETB' then    //Niveau etablisement
    begin
      setControlProperty('PTE_VALNIVB','Vide',True);
      setControlProperty('PTE_VALNIVB','VideString','<<Tous les établissements>>');
      setControlProperty('PTE_VALNIVB','DataType','TTETABLISSEMENT');
      setControlProperty('TPTE_VALNIV','Caption','Etablissement');
    end else if stNIVSAIS = 'CON' then   //Niveau Convention collective
    begin
      setControlProperty('PTE_VALNIVB','Vide',False);
      setControlProperty('PTE_VALNIVB','VideString','');
      setControlProperty('PTE_VALNIVB','DataType','PGCONVENTIONS');
      setControlProperty('TPTE_VALNIV','Caption','Convention');
    end else if stNIVSAIS = 'GEN' then   //Niveau Général
    begin
      setControlProperty('PTE_VALNIVB','DataType','');
    end;
    if GetControlText('PTE_VALNIV') <> '...' then
      SetControlText('PTE_VALNIVB',GetControlText('PTE_VALNIV'))
    else
      SetControlText('PTE_VALNIVB','');
  end;
  stVALNIV      := GetControlText('PTE_VALNIV');
  stTYPERESULTAT  := getControlText('PTE_TYPERESULTAT');
  stSENS1         := getControlText('PTE_SENS1');
  stSENS2         := getControlText('PTE_SENS2');
  boDonneesSalarie           := stNATURETABLE = 'DSA';
  boNiveauEtab_ou_Convention := (stNIVSAIS = 'ETB') or (stNIVSAIS = 'CON');
  boCritere1AgeAnciennete    := (stCRITERE1 = '1500') or (stCRITERE1 = '1501');
  boCritere2AgeAnciennete    := (stCRITERE2 = '1500') or (stCRITERE2 = '1501');
  boCOMBINCRITERE :=     ((stCRITERE1 <> '0') and (stCRITERE1 <> ''))
                     and (not boCritere1AgeAnciennete)
                     and (not boCritere2AgeAnciennete)
                     and (not (stNATURETABLE = 'ELT'))  //PT11
                     and (not (stNATURETABLE = 'VAR')); //PT3
  PrioriteTemp := 0;
  if      stPREDEFINI = 'CEG' then PrioriteTemp := 1000
  else if stPREDEFINI = 'STD' then PrioriteTemp := 2000
  else if stPREDEFINI = 'DOS' then PrioriteTemp := 3000;
  if      stNIVSAIS = 'GEN' then PrioriteTemp := PrioriteTemp + 100
  else if stNIVSAIS = 'CON' then PrioriteTemp := PrioriteTemp + 200
  else if stNIVSAIS = 'ETB' then PrioriteTemp := PrioriteTemp + 300;
  if (stNIVSAIS = 'CON') and (stVALNIV <> '000') then PrioriteTemp := PrioriteTemp + 20;
  if (stNIVSAIS = 'ETB') and (stVALNIV <> '...') then PrioriteTemp := PrioriteTemp + 30;
  SetControlText('PTE_PRIORITENIV',IntToStr(PrioriteTemp));
end;

procedure TOM_TABLEDIMENT.OnExitCRITERE2(Sender: TObject);
var
  SAVboAlerteInitGrid : Boolean;
begin
  if not boLoading AND (BackCRITERE2 <> GetControlText('PTE_CRITERE2')) then
  begin
    if not isBackCRITERE2 then
    begin
      if (not boAlerteInitGrid) or ((boAlerteInitGrid) and (PGIAsk( TraduireMemoire('Le second type de critère a été modifié.'#10#13'Toutes les données de la grille vont être supprimées.'#10#13#10#13'Voulez-vous continuer ?'),Ecran.Caption)
         = mrYes)) then
      begin
        SAVboAlerteInitGrid := boAlerteInitGrid;
        boAlerteInitGrid := False;
        InitGrid;
        boAlerteInitGrid := SAVboAlerteInitGrid;
      end else begin
        isBackCRITERE2 := True;
        SetControlText('PTE_CRITERE2',BackCRITERE2);
      end;
    end else begin
      isBackCRITERE2 := False;
    end;
  end;
  if not boLoading then UpdateVisibilites;
end;

procedure TOM_TABLEDIMENT.OnEnterCRITERE1B(Sender: TObject);
begin
  if (stNATURETABLE = 'VAR') or (stNATURETABLE = 'ELT') then Exit; //PT3  //PT11
  UpdateValues;
  //On sauvegarde le type de CRITERE1 avant modif, pour pouvoir le remettre
  // si l'utilisateur annule la réinitialisation des données de la grille
  // lors du changement de type de CRITERE1
  BackCRITERE1 := stCRITERE1;
end;

procedure TOM_TABLEDIMENT.OnEnterCRITERE2(Sender: TObject);
begin
  UpdateValues;
  //On sauvegarde le type de CRITERE2 avant modif, pour pouvoir le remettre
  // si l'utilisateur annule la réinitialisation des données de la grille
  // lors du changement de type de CRITERE2
  BackCRITERE2 := stCRITERE2;
end;

procedure TOM_TABLEDIMENT.OnClickCOMBINCRITERE(Sender: TObject);
var
  indexCritere1, indexCritere2 : integer;
  Q : TQuery;
  stQ : String;
  PrefixeTablette1, SQLT1, SQLT2, PrefixeTablette2 : String;
  RetourTablette1, RetourTablette2 : String;
  TOBTablette1, TOBTablette2 : TOB;
  Count1, Count2, TempRow : Integer;
  Mess, Mess2 : String;
  TabletteName : String;
begin
  SetFocusControl('PTE_LIBELLE');
  if PGIAsk( TraduireMemoire('Vous avez activé la combinaison automatique des critères.'#10#13'Toutes les données de la grille vont être supprimées.'#10#13#10#13'Voulez-vous continuer ?'),Ecran.Caption)
        = mrYes then
  begin
    if (stCRITERE1 = '1500') or (stCRITERE1 = '1501') or (stCRITERE2 = '1500') or (stCRITERE2 = '1501') then exit;
    //On réinitialise la grille, et on déroule les 2 critères
    InitGrid;
    PDETAIL.RowCount := 1;
    if (stCRITERE1 <> '0') and (stCRITERE1 <> '') then
    begin
      TabletteName := FindTabletteName(StrToInt(stCRITERE1));
      PrefixeTablette1 := '';
      RetourTablette1 := '';
      if TabletteName <> '' then
      begin
        try
          stQ := 'SELECT DO_CODE, DO_PREFIXE FROM DECOMBOS WHERE DO_COMBO="'+TabletteName+'"';
          Q := OpenSQL(stQ, True, 1);
          if not Q.Eof then
          begin
            PrefixeTablette1 := Q.FindField('DO_PREFIXE').AsString;
            RetourTablette1 := Q.FindField('DO_CODE').AsString;
          end;
          Ferme(Q);
        except
          PrefixeTablette1 := '';
        end;
      end;
      TOBTablette1 := TOB.Create(PrefixeToTable(PrefixeTablette1),nil,-1);
      if TabletteName <> '' then
      begin
        SQLT1 := GetTabletteSql(FindTabletteName(StrToInt(stCRITERE1)),PlusCritere1);
        TOBTablette1.LoadDetailDBFromSql(PrefixeToTable(PrefixeTablette1), SQLT1);
      end;
    end;
    if (stCRITERE2 <> '0') and (stCRITERE2 <> '') then
    begin
      TabletteName := FindTabletteName(StrToInt(stCRITERE2));
      PrefixeTablette2 := '';
      RetourTablette2 := '';
      if TabletteName <> '' then
      begin
        try
          stQ := 'SELECT DO_CODE, DO_PREFIXE FROM DECOMBOS WHERE DO_COMBO="'+TabletteName+'"';
          Q := OpenSQL(stQ, True, 1);
          PrefixeTablette2 := Q.FindField('DO_PREFIXE').AsString;
          RetourTablette2 := Q.FindField('DO_CODE').AsString;
          Ferme(Q);
        except
          PrefixeTablette2 := '';
        end;
      end;
      TOBTablette2 := TOB.Create(PrefixeToTable(PrefixeTablette2),nil,-1);
      if TabletteName <> '' then
      begin
        SQLT2 := GetTabletteSql(TabletteName,PlusCritere2);
        TOBTablette2.LoadDetailDBFromSql(PrefixeToTable(PrefixeTablette2), SQLT2);
      end;
    end;
    // On vire les combobox de la grille pour qu'une combo ne soit pas ouverte
    // au moment où on change l'"editabilité" de la colonne
    if Assigned(PDETAIL.ValCombo) then
    begin
      PDETAIL.Col := 3;
      PDETAIL.ValCombo.Hide;
      PDETAIL.ValCombo.Free;
    end;
    // Puis on réaffecte les combobox pour que les données soient gérés comme il
    // faut (PDETAIL.Cells contient les libellés et PDETAIL.cellValues les codes)
    UpdateVisibilites;
    if (stCRITERE1 <> '0') and (stCRITERE1 <> '') then
      Count1 := TOBTablette1.FillesCount(1)
    else
      Count1 := 1;
    if (stCRITERE2 <> '0') and (stCRITERE2 <> '') then
      Count2 := TOBTablette2.FillesCount(1)
    else
      Count2 := 1;
    PDETAIL.RowCount := (Count1*Count2)+1;

    InitMoveProgressForm(nil,Ecran.caption,'Traitement en cours, veuillez patienter...', PDETAIL.RowCount, True, true) ;
    Mess := RechDom(Control_AbsCRITERE1B.DataType,stCRITERE1,False)+' '+TraduireMemoire('valeur')+' ';   //PT10
    Mess2 := RechDom(Control_CRITERE2.DataType,stCRITERE2,False)+' '+TraduireMemoire('valeur')+' ';
//    indexCritere1 := 0;
//    indexCritere2 := 0;
    for indexCritere1 := 0 to (TOBTablette1.FillesCount(1)-1) do
    begin
      if (stCRITERE2 <> '0') and (stCRITERE2 <> '') then
      begin
        for indexCritere2 := 0 to (TOBTablette2.FillesCount(1)-1) do
        begin
          PDETAIL.Col := 3;
          TempRow := (indexCritere1 * Count2) + indexCritere2 + 1;
          PDETAIL.CellValues[0,TempRow] := TOBTablette1.Detail[indexCritere1].GetString(RetourTablette1);
          PDETAIL.CellValues[1,TempRow] := TOBTablette2.Detail[indexCritere2].GetString(RetourTablette2);
          PDETAIL.Cells     [3,TempRow] := '';
          PDETAIL.CellValues[3,TempRow] := '';
          if not MoveCurProgressForm(Mess+ IntToStr(indexCritere1)+' - '+Mess2+IntToStr(indexCritere2)) then
          begin
            ForceUpdate;
            UpdateVisibilites;
            FiniMoveProgressForm ;
            Exit;
          end;
        end;
      end else begin
        PDETAIL.Col := 3;
        TempRow := (indexCritere1 * Count2) + 1;  //+ indexCritere2 
        PDETAIL.CellValues[0,TempRow] := TOBTablette1.Detail[indexCritere1].GetString(RetourTablette1);
        PDETAIL.Cells     [3,TempRow] := '';
        PDETAIL.CellValues[3,TempRow] := '';
        if not MoveCurProgressForm(Mess+ IntToStr(indexCritere1)) then break;
      end;
    end;
  end;
  if Assigned(TOBTablette1) then FreeAndNil(TOBTablette1);
  if Assigned(TOBTablette2) then FreeAndNil(TOBTablette2);
  //Le forceUpdate permet d'activer les boutons de suppression et d'insertion de la fiche
  ForceUpdate;
  UpdateVisibilites;
  FiniMoveProgressForm ;
end;

procedure TOM_TABLEDIMENT.UpdateAutorisation;
begin
  Autorisation := False;
  if (stPREDEFINI = 'CEG') and (CEG) then
    Autorisation := True
  else if (stPREDEFINI = 'STD') and (STD) then
    Autorisation := True
  else if (stPREDEFINI = 'DOS') and (DOS) then
    Autorisation := True;
end;

procedure TOM_TABLEDIMENT.OnChangeTABLE(Sender: TObject);
begin
  if not boLoading then UpdateVisibilites;
end;

procedure TOM_TABLEDIMENT.OnEnterTABLE(Sender: TObject);
begin
  BackTABLE := GetControlText('PTE_CODTABL');
end;

//Debut PT1
procedure TOM_TABLEDIMENT.OnExitTABLE(Sender: TObject);
begin
  if stCODTABL = '   ' then exit;
  LastError := 0;
  if (not IsNumeric(stCODTABL)) or (pos('.',stCODTABL) > 0) or (pos(',',stCODTABL) > 0) or (length(stCODTABL) <> 3 ) then
  begin
    LastErrorMsg := TraduireMemoire('Le code de la table doit être un nombre à trois chiffres.');
    LastError := 101;
    SetControlText('PTE_CODTABL', BackTABLE);
    SetFocusControl('PTE_CODTABL');
    PGIError(LastErrorMsg,Ecran.Caption);
    Exit;
  end;
//  if (stPREDEFINI = 'CEG') and (not (StrToInt(RightStr(stCODTABL,1)) mod 2 = 0) ) then
//  begin
//    LastErrorMsg := TraduireMemoire('Le code de la table doit être un nombre PAIR.');
//    LastError := 102;
//    SetControlText('PTE_CODTABL', BackTABLE);
//    SetFocusControl('PTE_CODTABL');
//    PGIError(LastErrorMsg,Ecran.Caption);
//    Exit;
//  end;
//  if (stPREDEFINI = 'STD') and
//     ( (    (StrToInt(RightStr(stCODTABL,1)) <> 1 )
//        and (StrToInt(RightStr(stCODTABL,1)) <> 3 )
//       )
//     ) then
//  begin
//    LastErrorMsg := TraduireMemoire('Le code de la table doit être un nombre se terminant par 1 ou 3.');
//    LastError := 102;
//    SetControlText('PTE_CODTABL', BackTABLE);
//    SetFocusControl('PTE_CODTABL');
//    PGIError(LastErrorMsg,Ecran.Caption);
//    Exit;
//  end;
//  if (stPREDEFINI = 'DOS') and
//     ( (    (StrToInt(RightStr(stCODTABL,1)) <> 5)
//        and (StrToInt(RightStr(stCODTABL,1)) <> 7)
//        and (StrToInt(RightStr(stCODTABL,1)) <> 9)
//       )
//     ) then
//  begin
//    LastErrorMsg := TraduireMemoire('Le code de la table doit être un nombre se terminant par 5, 7 ou 9.');
//    LastError := 102;
//    SetControlText('PTE_CODTABL', BackTABLE);
//    SetFocusControl('PTE_CODTABL');
//    PGIError(LastErrorMsg,Ecran.Caption);
//    Exit;
//  end;
end;
//Fin PT1

procedure TOM_TABLEDIMENT.AffecteEvenements;
begin
  //On affecte les évenements aux composants de la fiche
  Control_TABLE.OnEnter           := OnEnterTABLE;
  Control_TABLE.OnChange          := OnChangeTABLE;
  Control_TABLE.OnExit            := OnExitTABLE;
  Control_NATURETABLE.OnEnter     := OnEnterNATURETABLE;
  Control_NATURETABLE.OnChange    := OnChangeNATURETABLE;
  Control_NATURETABLE.OnExit      := OnExitNATURETABLE;
  Control_NIVSAIS.OnChange        := OnChangeNIVSAIS;
  Control_VALNIV.OnChange         := OnChangeVALNIV;
  Control_TYPERESULTAT.OnEnter    := OnEnterTYPERESULTAT;
  Control_TYPERESULTAT.OnChange   := OnChangeTYPERESULTAT;
  Control_PREDEFINI.OnEnter       := OnEnterPREDEFINI;
  Control_PREDEFINI.OnChange      := OnChangePREDEFINI;
  Control_CRITERE1B.OnEnter       := OnEnterCRITERE1B;
  Control_CRITERE1B.OnExit        := OnExitCRITERE1B;
  Control_CRITERE1B.OnChange      := OnChangeCRITERE1B;
  Control_CRITERE1B_.OnEnter       := OnEnterCRITERE1B;   //PT10
  Control_CRITERE1B_.OnExit        := OnExitCRITERE1B;    //PT10
  Control_CRITERE1B_.OnChange      := OnChangeCRITERE1B;  //PT10
  Control_CRITERE2.OnEnter        := OnEnterCRITERE2;
  Control_CRITERE2.OnExit         := OnExitCRITERE2;
  Control_COMBINCRITERE.OnClick   := OnClickCOMBINCRITERE;
  BtnIns.OnClick                  := BTnInsClick;
  BtnDel.OnClick                  := BTnDelClick;
  BtnAdd.OnClick                  := BTnAddClick;
  PDetail.Oncellenter             := PdetailCellEnter;
  PDetail.OnCellExit              := PdetailCellExit;
  PDetail.OnElipsisClick          := PDetailElipsisClick;
end;

Procedure TOM_TABLEDIMENT.GetPrefixeCodeLibelleTablette(Tablette: String; Var Prefixe, Code, Libelle : String);
var
  Q : TQuery;
  stQ : String;
begin
  Prefixe := '';
  Code := '';
  Libelle := '';
  stQ := 'SELECT DO_CODE, DO_CHAMPLIB, DO_PREFIXE FROM DECOMBOS WHERE DO_COMBO="'+Tablette+'"';
  Q := OpenSQL(stQ, True, 1);
  try
    Prefixe := Q.FindField('DO_PREFIXE').AsString;
    Code := Q.FindField('DO_CODE').AsString;
    Libelle := Q.FindField('DO_CHAMPLIB').AsString;
  except
    Ferme(Q);
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : GGU
Créé le ...... : 05/03/2007
Modifié le ... :   /  /
Description .. : Met a jour les "Plus" pour les combo de la grille.
Suite ........ : Renvoie Vraie si les valeurs ont été modifiées, sinon,
Suite ........ : renvoie Faux.
Mots clefs ... :
*****************************************************************}
function TOM_TABLEDIMENT.UpdatePlusCriteres : Boolean;

  procedure UpdatePlusCritere(stCRITERE : String; var TabletteCritere, PlusCritere : String);
  var
    Prefixe, CodeTablette, Lib : String;
  begin
    PlusCritere := '';
    WhereVar :=  '';
    WhereElt :=  '';
    if stCRITERE <> '' then
    begin
      TabletteCritere := FindTabletteName(StrToInt(stCRITERE));
      GetPrefixeCodeLibelleTablette(TabletteCritere,Prefixe,CodeTablette,Lib);
      if stNIVSAIS = 'CON' then
      begin
        if (Prefixe <> '') and (_ChampExiste(PrefixeToTable(Prefixe),Prefixe+'_CONVENTION')) then
          PlusCritere := ' AND '+Prefixe+'_CONVENTION = "'+stVALNIV+'" ';
        WhereElt :=  ' AND PEL_CONVENTION = "'+stVALNIV+'" ';
      end;
      if stPREDEFINI = 'STD' then
      begin
        if (Prefixe <> '') and (_ChampExiste(PrefixeToTable(Prefixe),Prefixe+'_PREDEFINI')) then
          PlusCritere := PlusCritere + ' AND ('+Prefixe+'_PREDEFINI = "STD" or '+Prefixe+'_PREDEFINI = "CEG")';
        WhereVar := WhereVar + ' AND (PVA_PREDEFINI = "STD" or PVA_PREDEFINI = "CEG")';
        WhereElt := WhereElt + ' AND (PEL_PREDEFINI = "STD" or PEL_PREDEFINI = "CEG")';
      end else if stPREDEFINI = 'DOS' then
      begin
        if (Prefixe <> '') and ( _ChampExiste(PrefixeToTable(Prefixe),Prefixe+'_PREDEFINI')) then
        begin
          if PlusCritere <> '' then
            PlusCritere := ' AND ##'+Prefixe+'_PREDEFINI##' + PlusCritere
          else
            PlusCritere := ' AND ##'+Prefixe+'_PREDEFINI##';
        end;
        WhereVar := WhereVar + '##PVA_PREDEFINI##';
        WhereElt := WhereElt + '##PEL_PREDEFINI##';
      end;
    end;
  end;
begin
  UpdatePlusCritere(stCRITERE1, TabletteCritere1, PlusCritere1);
  UpdatePlusCritere(stCRITERE2, TabletteCritere2, PlusCritere2);
  if    (BackTabletteCritere1 = TabletteCritere1)
    and (BackPlusCritere1 = PlusCritere1)
    and (BackTabletteCritere2 = TabletteCritere2)
    and (BackPlusCritere2 = PlusCritere2) then
    result := False
  else begin
    result := True;
    BackTabletteCritere1 := TabletteCritere1;
    BackPlusCritere1     := PlusCritere1;
    BackTabletteCritere2 := TabletteCritere2;
    BackPlusCritere2     := PlusCritere2;
  end;
end;

procedure TOM_TABLEDIMENT.OnChangeVALNIV(Sender: TObject);
begin
  if not boLoading then UpdateVisibilites;
end;

procedure TOM_TABLEDIMENT.PDetailCellExit(Sender: TObject; var ACol,
  ARow: Integer; var Cancel: Boolean);
var
  tempstr : string;
begin

  if (ACol = 3) then
  begin
    if (Getfield('PTE_TYPERESULTAT') <> 'VAR') and (Getfield('PTE_TYPERESULTAT') <> 'ELT') then
    begin
      tempstr := StrToNumStr(PDETAIL.Cells     [3,ARow]);
      PDETAIL.Cells     [3,ARow] := tempstr;
      PDETAIL.CellValues[3,ARow] := tempstr;
    end;
  end;

  if (ACol = 0) then
  begin
    GridFormatCol0(ARow); 
{    if (Getfield('PTE_NATURETABLE') = 'VAR') or (Getfield('PTE_NATURETABLE') = 'COD') then
    begin
      tempstr := PDETAIL.Cells [0,ARow];
      tempstr := StringReplace(tempstr,'.','0',[rfReplaceAll, rfIgnoreCase ]);
      tempstr := StringReplace(tempstr,',','0',[rfReplaceAll, rfIgnoreCase ]);
      tempstr := StrToNumStr(tempstr);
      if           (PDetail.ColFormats[0] = FormatC_DSA) then
      begin
        if Length(tempstr) < NbrC_DSA then tempstr := lpad(tempstr,NbrC_DSA)
      end else if  (PDetail.ColFormats[0] = FormatC_COD) then
      begin
        if Length(tempstr) < NbrC_COD then tempstr := lpad(tempstr,NbrC_COD);
      end;
      PDETAIL.Cells     [0,ARow] := tempstr;
      PDETAIL.CellValues[0,ARow] := tempstr;
    end;
}
  end;

end;

procedure TOM_TABLEDIMENT.OnChangeCRITERE1B(Sender: TObject);
var
  tempC1, tempC1B : String;
begin
  if not boLoading then    //PT7
  begin
//PT10    tempC1B := GetControlText('PTE_CRITERE1B');
    tempC1B := Control_AbsCRITERE1B.ControlText; //PT10  
    if varType(GetField('PTE_CRITERE1')) = varNull then
    begin
      if tempC1B <> '' then
        SetField('PTE_CRITERE1',tempC1B)
      else
      SetField('PTE_CRITERE1',0);
      exit;
    end;
    tempC1 :=  GetField('PTE_CRITERE1');
// Debut PT11
//    if (    ((GetField('PTE_NATURETABLE') = 'VAR')  and  ( tempC1B <> lpad(tempC1,4) ) )
//         or ((GetField('PTE_NATURETABLE') <> 'VAR') and  ( tempC1B <> tempC1 ))
//       )
//       and not ((tempC1B = '') and (tempC1 = '0')) then
    if (    (((GetField('PTE_NATURETABLE') =  'VAR') or  (GetField('PTE_NATURETABLE') =  'ELT')) and (tempC1B <> lpad(tempC1,4)))
         or (((GetField('PTE_NATURETABLE') <> 'VAR') and (GetField('PTE_NATURETABLE') <> 'ELT')) and (tempC1B <> tempC1))
       )
       and not ((tempC1B = '') and (tempC1 = '0')) then
// Fin PT11
    begin
      if tempC1B <> '' then
        SetField('PTE_CRITERE1',tempC1B)
      else
        SetField('PTE_CRITERE1',0);
    end;
  end; 
end;

procedure TOM_TABLEDIMENT.OnExitNATURETABLE(Sender: TObject);
begin
  //Debut PT10
  UpdateCriterePRNature(False);
//  if GetField('PTE_NATURETABLE') = 'VAR' then //PT3
//  begin
//    if (Control_CRITERE1B.DataType <> 'PGVARIABLE') then
//    begin
//      Control_CRITERE1B.Vide := False;
//      Control_CRITERE1B.Plus := '';
//      Control_CRITERE1B.DataType := 'PGVARIABLE';
//      SetControlText('TPTE_CRITERE1',TraduireMemoire('Variable de paie'));
//    end;
//  end else begin
//    if (Control_CRITERE1B.DataType <> 'PGCHAMPCRITERETABLE') then
//    begin
//      Control_CRITERE1B.Vide := True;
//      Control_CRITERE1B.Plus := 'AND ((PAI_IDENT <> 45) AND (PAI_IDENT <> 46) AND (PAI_IDENT <> 220) AND (PAI_IDENT <> 221))';
//      Control_CRITERE1B.DataType := 'PGCHAMPCRITERETABLE';
//      SetControlText('TPTE_CRITERE1',TraduireMemoire('Critère 1'));
//    end;
//  end;
  //Fin PT10

end;

//Debut PT10
{ TControlAbstractComboEdit }

constructor TControlAbstractComboEdit.Create(LaTom: Tom);
begin
  inherited Create;
  LinkTom := LaTom;
end;

destructor TControlAbstractComboEdit.Destroy;
begin
  inherited;
end;

function TControlAbstractComboEdit.GetControlText: String;
begin
  if ControlToAbstract is THValComboBox then
    result := LinkTom.GetControlText('PTE_CRITERE1B')
  else if ControlToAbstract is THEdit then
    result := LinkTom.GetControlText('PTE_CRITERE1B_');
end;

function TControlAbstractComboEdit.GetDataType: String;
begin
  result := '';
  if ControlToAbstract is THValComboBox then
    result := (ControlToAbstract as THValComboBox).DataType
  else if ControlToAbstract is THEdit then
    result := (ControlToAbstract as THEdit).DataType;
end;

function TControlAbstractComboEdit.GetEnabled: Boolean;
begin
  result := False;
  if ControlToAbstract is THValComboBox then
    result := (ControlToAbstract as THValComboBox).Enabled
  else if ControlToAbstract is THEdit then
    result := (ControlToAbstract as THEdit).Enabled;
end;

function TControlAbstractComboEdit.GetPlus: String;
begin
  result := '';
  if ControlToAbstract is THValComboBox then
    result := (ControlToAbstract as THValComboBox).Plus
  else if ControlToAbstract is THEdit then
    result := (ControlToAbstract as THEdit).Plus;
end;

function TControlAbstractComboEdit.GetVide: Boolean;
begin
  result := False;
  if ControlToAbstract is THValComboBox then
    result := (ControlToAbstract as THValComboBox).Vide;
end;

function TControlAbstractComboEdit.GetVisible: Boolean;
begin
  result := False;
  if ControlToAbstract is THValComboBox then
    result := (ControlToAbstract as THValComboBox).Visible
  else if ControlToAbstract is THEdit then
    result := (ControlToAbstract as THEdit).Visible;
end;

procedure TControlAbstractComboEdit.SetControlText(const Value: String);
begin
  if ControlToAbstract is THValComboBox then
    LinkTom.SetControlText('PTE_CRITERE1B', Value)
  else if ControlToAbstract is THEdit then
    LinkTom.SetControlText('PTE_CRITERE1B_', Value);
end;

procedure TControlAbstractComboEdit.SetControlToAbstract(
  const Value: TControl);
begin
  { On cache l'ancien control }
  if ControlToAbstract is THValComboBox then
  begin
    (ControlToAbstract as THValComboBox).Visible := False;
  end else if ControlToAbstract is THEdit then
  begin
    (ControlToAbstract as THEdit).Visible := False;
    (ControlToAbstract as THEdit).Libelle.Visible := False;
  end;
  FControlToAbstract := Value;
  { On affiche le nouveau control }
  if ControlToAbstract is THValComboBox then
  begin
     (ControlToAbstract as THValComboBox).Visible := True;
  end else if ControlToAbstract is THEdit then
  begin
    (ControlToAbstract as THEdit).Visible := True;
    (ControlToAbstract as THEdit).Libelle.Visible := True;
  end;
end;

procedure TControlAbstractComboEdit.SetDataType(const Value: String);
begin
  if ControlToAbstract is THValComboBox then
    (ControlToAbstract as THValComboBox).DataType := Value
  else if ControlToAbstract is THEdit then
    (ControlToAbstract as THEdit).DataType := Value;
end;

procedure TControlAbstractComboEdit.SetEnabled(const Value: Boolean);
begin
  if ControlToAbstract is THValComboBox then
    (ControlToAbstract as THValComboBox).Enabled := Value
  else if ControlToAbstract is THEdit then
    (ControlToAbstract as THEdit).Enabled := Value;
end;

procedure TControlAbstractComboEdit.SetFocus;
begin
  if ControlToAbstract is THValComboBox then
    LinkTom.SetFocusControl('PTE_CRITERE1B')
  else if ControlToAbstract is THEdit then
    LinkTom.SetFocusControl('PTE_CRITERE1B_'); 
end;

procedure TControlAbstractComboEdit.SetLinkTom(const Value: Tom);
begin
  FLinkTom := Value;
end;

procedure TControlAbstractComboEdit.SetPlus(const Value: String);
begin
  if ControlToAbstract is THValComboBox then
    (ControlToAbstract as THValComboBox).Plus := Value
  else if ControlToAbstract is THEdit then
    (ControlToAbstract as THEdit).Plus := Value;
end;

procedure TControlAbstractComboEdit.SetVide(const Value: Boolean);
begin
  if ControlToAbstract is THValComboBox then
    (ControlToAbstract as THValComboBox).Vide := Value;
//  else if ControlToAbstract is THEdit then
//    (ControlToAbstract as THEdit).Vide := Value;
end;

procedure TControlAbstractComboEdit.SetVisible(const Value: Boolean);
begin
  if ControlToAbstract is THValComboBox then
    (ControlToAbstract as THValComboBox).Visible := Value
  else if ControlToAbstract is THEdit then
    (ControlToAbstract as THEdit).Visible := Value;
end;
//Fin PT10

//Debut PT10
procedure TOM_TABLEDIMENT.UpdateCriterePRNature(Force : Boolean);
begin
  if GetField('PTE_NATURETABLE') = 'VAR' then //PT3
  begin
    if Force or (Control_AbsCRITERE1B.DataType <> 'PGVARIABLE') then
    begin
      Control_AbsCRITERE1B.ControlToAbstract := Control_CRITERE1B_;  //PT10
  //    Control_CRITERE1B.Vide := False;
      Control_AbsCRITERE1B.Vide := False;
  //    Control_CRITERE1B.Plus := '';
      Control_AbsCRITERE1B.Plus := '';
  //    Control_CRITERE1B.DataType := 'PGVARIABLE';
      Control_AbsCRITERE1B.DataType := 'PGVARIABLE';
  //    SetControlText('PTE_CRITERE1B', lpad(GetField('PTE_CRITERE1'),4));
      if (GetField('PTE_CRITERE1') = 0) or (not Force) then
        Control_AbsCRITERE1B.ControlText := ''
      else
        Control_AbsCRITERE1B.ControlText := lpad(GetField('PTE_CRITERE1'),4);
      SetControlText('TPTE_CRITERE1',TraduireMemoire('Variable de paie'));
    end;
  end else if GetField('PTE_NATURETABLE') = 'ELT' then  //PT11
  begin
    if Force or (Control_AbsCRITERE1B.DataType <> 'PGELEMENTNAT') then
    begin
      Control_AbsCRITERE1B.ControlToAbstract := Control_CRITERE1B_; //PT10
      Control_AbsCRITERE1B.Vide := False;
      Control_AbsCRITERE1B.Plus := '';
      Control_AbsCRITERE1B.DataType := 'PGELEMENTNAT';
      if (GetField('PTE_CRITERE1') = 0) or (not Force) then
        Control_AbsCRITERE1B.ControlText := ''
      else
        Control_AbsCRITERE1B.ControlText := lpad(GetField('PTE_CRITERE1'),4);
      SetControlText('TPTE_CRITERE1',TraduireMemoire('Elément national'));
    end;
  end else if GetField('PTE_NATURETABLE') = 'DSA' then
  begin
    if Force or (Control_AbsCRITERE1B.DataType <> 'PGCHAMPCRITERETABLE') then
    begin
      Control_AbsCRITERE1B.ControlToAbstract := Control_CRITERE1B;  //PT10
  //    Control_CRITERE1B.Vide := True;
      Control_AbsCRITERE1B.Vide := True;
  //    Control_CRITERE1B.Plus := 'AND ((PAI_IDENT <> 45) AND (PAI_IDENT <> 46) AND (PAI_IDENT <> 220) AND (PAI_IDENT <> 221))';
      Control_AbsCRITERE1B.Plus := 'AND ((PAI_IDENT <> 45) AND (PAI_IDENT <> 46) AND (PAI_IDENT <> 220) AND (PAI_IDENT <> 221))';
  //    Control_CRITERE1B.DataType := 'PGCHAMPCRITERETABLE';
      Control_AbsCRITERE1B.DataType := 'PGCHAMPCRITERETABLE';
  //    SetControlText('PTE_CRITERE1B',GetField('PTE_CRITERE1'));
      if (GetField('PTE_CRITERE1') = 0) or (not Force) then
        Control_AbsCRITERE1B.ControlText := ''
      else
        Control_AbsCRITERE1B.ControlText := GetField('PTE_CRITERE1');
      SetControlText('TPTE_CRITERE1',TraduireMemoire('Critère 1'));
    end;
  end;
end;
//Fin PT10

Initialization
  registerclasses ( [ TOM_TABLEDIMENT ] ) ;
end.
