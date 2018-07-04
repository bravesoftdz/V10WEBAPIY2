{***********UNITE*************************************************
Auteur  ...... : Stéphane BOUSSERT
Créé le ...... : 28/03/2002
Modifié le ... : 29/04/2002
Description .. : Source TOM de la TABLE : SUIVCPTA (SUIVCPTA)
Mots clefs ... : TOM;SUIVCPTA
*****************************************************************}
Unit SUIVCPTA_TOM ;

//================================================================================
// Interface
//================================================================================
Interface

Uses
    StdCtrls,
    Controls,
    Classes,
    HMsgBox,
    forms,
    sysutils,
    ComCtrls,
    HCtrls,
    HEnt1,
{$IFDEF EAGLCLIENT}
    eFiche,
    MaineAGL,
    eFichList,
{$ELSE}
    DB,
    DBCtrls,
    {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
    Fiche,
    Fe_main,
    FichList,
{$ENDIF}
    UTob,
    UTOM,
    Choix,
    AtChComp,
    MajTable,
    SAISUTIL,
    LibChpLi,
    ULibEcriture,
    UtilSais,
    ed_tools
    ;

//==================================================
// Externe
//==================================================
procedure CPLanceFiche_Scenario(lequel : String = ''; arg : String = ''; Multi : boolean = false);
Procedure ParamScenario(Jal : string3 = ''; Nat : String3 = '');
Procedure MultiScenario;

//==================================================
// Definition de class
//==================================================
type
    TOM_SUIVCPTA = class (TOM)
        procedure OnNewRecord              ; override ;
        procedure OnDeleteRecord           ; override ;
        procedure OnUpdateRecord           ; override ;
        procedure OnAfterUpdateRecord      ; override ;
        procedure OnLoadRecord             ; override ;
        procedure OnChangeField(F: TField) ; override ;
        procedure OnArgument(S: String)    ; override ;
        procedure OnCancelRecord           ; override ;
        procedure OnClose                  ; override ;
    private
        mode : integer;

        Complement : array[1..10,1..10] of Char;
        CompLibre  : array[1..10,1..40] of Char;
        Radicaux   : array[1..10] of String[10];
        LibreEntete: array[1..40] of Char;
        Racines, Libres, LibreEnt : TGroupBox;
        GRacines : THGrid;
        ParamLibTOB : TOB;
        GeneCharge, CbFlaguer : Boolean;
        FBoModePiece : boolean ;
        procedure ForceChange(Sender: TObject);
        procedure OnEnterGRacinesRow(Sender: TObject; Ou: Longint; var Cancel: Boolean; Chg: Boolean);
        procedure OnExitGRacinesRow(Sender: TObject; Ou: Longint; var Cancel: Boolean; Chg: Boolean);
        procedure OnKeyPressGRacinesRow(Sender: TObject; var Key: Char);
        procedure OnClickCopierScenario(Sender: TObject);
        procedure OnClickAfficherInfoComp(Sender: TObject);
        procedure OnClickBGenerer(Sender: TObject);
        procedure OnClickQte( Sender : TObject ) ;

        procedure InitTableaux;
        procedure ChargeComplement;
        procedure SauveComplement;
        procedure MajComplement(ou : LongInt);
        procedure RAZRacines;
        procedure AttribLesHint;
        procedure AccesCoches(CT,CL : String ; VV : Boolean);
        procedure ChargeParamLibTOB;
        procedure GenereListe(Li : HTStrings);
        procedure ChercheAcces(Li : HTStrings);
        function  AjouteAcces(C:TCustomCheckBox; MemoSt, StOrigine, Radic: String; QuelTab: Byte): String;
        procedure VerifierAcces(Tout : boolean);
//        procedure LimiteAccesControles;
        procedure SelectionJournal;
        function  EnregOkMulti: Boolean;
        function  EstSaisieQte : Boolean ;
        procedure GereAccesCtrlQte ;
        {JP 16/10/07 : FQ 16149 : Gestion des restrictions établissement}
        procedure GereEtablissement;
    end;

//================================================================================
// Implementation
//================================================================================
Implementation

uses
  {$IFDEF MODENT1}
  CPTypeCons,
  {$ENDIF MODENT1}
  ENT1;


//==================================================
// Definition des Constante
//==================================================
Const MessageListe : array[1..19] of String = (
    {01}    'Voulez-vous enregistrer les modifications ??',
    {02}    'Désirez-vous créer un nouvel enregistrement ?',
    {03}    'Confirmez-vous la suppression de l''enregistrement ?',
    {04}    'Vous devez renseigner un journal et une nature de pièce.',
    {05}    'Le code que vous avez saisi existe déjà. Vous devez le modifier.',
    {06}    'L''enregistement est inaccessible',
    {07}    'Ce scénario existe déjà pour ',
    {08}    'un groupe d''utilisateurs !',
    {09}    'tous les groupes d''utilisateurs !',
    {10}    'Vous devez renseigner un groupe d''utilisateurs.',
    {11}    'Vous devez renseigner un journal.',
    {12}    'Vous devez renseigner une nature de pièce.',
    {13}    'Ce scénario existe déja !',
    {14}    'Il n''existe pas de scénario pour le journal et la nature de pièce sélectionnés !',
    {15}    'Scénario de saisie comptable : ',
    {16}    'Vous devez renseigner un établissement !',
    {17}    'Vous devez renseigner un type de pièce !',
    {18}    'Aucune information complémentaire n''a été sélectionné pour ce scénario !',
    {19}    'Choix d''un scénario pour la recopie'
            );

//==================================================
// fonctions hors class
//==================================================
{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 10/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
Procedure ParamScenario(Jal,Nat : String3);
var
    Arg, Lequel : String;
begin
    Arg := '';
    Lequel := '';

    if (Jal <> '') and (Nat <> '') then Lequel := '"' + V_PGI.Groupe + '";"' + Jal + '";"' + Nat +'"';

    CPLanceFiche_Scenario(lequel,arg,false);
end ;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 10/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
Procedure MultiScenario;
begin
    CPLanceFiche_Scenario('','',true);
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 10/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure CPLanceFiche_Scenario(lequel, arg : String ; Multi : boolean);
begin
    if (multi) then if (arg <> '') then arg := arg + ';MULTI' else arg := 'MULTI';
    AGLLanceFiche('CP','CPSCENARIO',Lequel,'','ACTION=MODIFICATION;' + Arg);
end;

//==================================================
// Evenements par default de la TOM
//==================================================
{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 10/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOM_SUIVCPTA.OnNewRecord;
var
    i : Integer ;
    Lib : String ;
begin
    inherited ;
    // Index
//    SetField('SC_NATUREPIECE','...');
    GereEtablissement;
    if GetControlText('SC_ETABLISSEMENT') = '' then
      SetField('SC_ETABLISSEMENT','...');
      
    SetField('SC_QUALIFPIECE','...');
    SetField('SC_USERGRP','...');
    // Scenario
    SetField('SC_OUVREECHE','X');
    SetField('SC_OUVREANAL','X');
    SetField('SC_CONTROLEQTE','-');
    SetField('SC_CONTROLETVA','RIE');
    SetField('SC_LETTRAGESAISIE','-');
    SetField('SC_NUMAXE','A1');
    SetField('SC_ALERTEDEV','JOU');
    SetField('SC_RIB','...');
    // Entete
    SetField('SC_DATEREFEXTERNE', '-');
    SetField('SC_REFLIBRE', '-');
    SetField('SC_AFFAIRE', '-');
    SetField('SC_REFINTERNE','-');
    SetField('SC_REFEXTERNE','-');
    SetField('SC_REFAUTO','');
    // Radicaux / compléments / Zones libre ligne et entetes
    Lib := '';
    for i := 1 to 40 do Lib := Lib+'-';

    SetField('SC_LIBREENTETE',Lib);

    for i := 1 to 10 do
    begin
        SetField('SC_COMPLEMENTS'+IntToStr(i),'----------');
        SetField('SC_RADICAL'+IntToStr(i),'');
        GRacines.Cells[1,i] := '';
        SetField('SC_COMPLIBRE'+IntToStr(i),Lib);
    end;

    { FQ 20872 BVE 27.09.07
      Deplacement de la fonction pour eviter de reprendre les anciennes valeurs }
    InitTableaux;
    { END FQ 20872 }

    if (Ecran <> nil) then
    begin
        SetControlEnabled('Cle',true);
        SetControlEnabled('Scenario',true);
        SetControlEnabled('Entete',true);
        SetControlEnabled('Racines',true);
        SetControlEnabled('Libres',true);

        RAZRacines;
    end;

end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 10/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOM_SUIVCPTA.OnDeleteRecord;
begin
    inherited;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 10/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOM_SUIVCPTA.OnUpdateRecord;
begin
    inherited ;
    // Choix du journal
    if (GetField('SC_JOURNAL') = '') then
    begin
        PGIBox(TraduireMemoire(MessageListe[11]),TraduireMemoire(Ecran.Caption));
        LastError := 1;
        Exit;
    end;
    // Choix de la nature de pièce
    if (GetField('SC_NATUREPIECE') = '') then
    begin
        PGIBox(TraduireMemoire(MessageListe[12]),TraduireMemoire(Ecran.Caption));
        LastError := 1;
        Exit;
    end;
    // Choix du groupe d'utilisateur
    if (GetField('SC_USERGRP') = '') then
    begin
        PGIBox(TraduireMemoire(MessageListe[10]),TraduireMemoire(Ecran.Caption));
        LastError := 1;
        Exit;
    end;
{$IFNDEF CCS3}
    // Choix de l'établissement
    if (GetField('SC_ETABLISSEMENT') = '') then
    begin
        PGIBox(TraduireMemoire(MessageListe[16]),TraduireMemoire(Ecran.Caption));
        LastError := 1;
        Exit;
    end;
    // Choix du type de pièce
    if (GetField('SC_QUALIFPIECE') = '') then
    begin
        PGIBox(TraduireMemoire(MessageListe[17]),TraduireMemoire(Ecran.Caption));
        LastError := 1;
        Exit;
    end;
{$ENDIF}
    // Remplissage auto
    if GetField('SC_ETABLISSEMENT') = '' then SetField('SC_ETABLISSEMENT','...');
    if GetField('SC_QUALIFPIECE') = '' then SetField('SC_QUALIFPIECE','...');
    // Mise en place des données complémentaires
    MajComplement(GRacines.row);
    SauveComplement;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 10/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOM_SUIVCPTA.OnAfterUpdateRecord;
begin
    inherited;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 10/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOM_SUIVCPTA.OnLoadRecord;
begin
    inherited;

    GeneCharge := true;
    CbFlaguer := false;
    if Ecran <> nil then SetControlEnabled('Cle',false);
    InitTableaux;
    GeneCharge := false;
    GereAccesCtrlQte ;    
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 10/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOM_SUIVCPTA.OnChangeField(F: TField);
begin
    inherited;

    if F.FieldName = 'SC_QUALIFPIECE' then VerifierAcces(false);
    if F.FieldName = 'SC_JOURNAL' then
      begin
      SelectionJournal; // MAJ Qualif_piece
      VerifierAcces(false);
      end ;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 10/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOM_SUIVCPTA.OnArgument(S: String);
var
    s1,s2 : string;
    i : Integer;
    C : TControl;
    tab : TPageControl;
    TS4,TS5 : TTabSheet;
  lRange : string;
begin
    inherited ;
  {Si l'on ne gère pas les établissement ...}
  if not VH^.EtablisCpta  then
    {... on affiche l'établissement par défaut}
    lRange := 'SC_ETABLISSEMENT = "' + VH^.EtablisDefaut + '"'

  {On gère l'établissement, avec restriction forcée}
  else if VH^.ProfilUserC[prEtablissement].ForceEtab then begin
    if VH^.ProfilUserC[prEtablissement].Etablissement <> '' then
      lRange := 'SC_ETABLISSEMENT IN ("' + VH^.ProfilUserC[prEtablissement].Etablissement + '", "...")'
    else
      lRange := 'SC_ETABLISSEMENT IN ("' + VH^.EtablisDefaut + '", "...")'
  end;

  {Affectation du filtre sur les établissements}
  TFFicheListe(Ecran).SetNewRange('', lRange);
  {On ferme la Table pour éviter un message après le OnArgument :
    "Ne peut effectuer ce traitement sur un ensemble de données ouvert"}
  {$IFNDEF EAGLCLIENT}
  TFFicheListe(Ecran).Ta.Close;
  {$ENDIF EAGLCLIENT}

    mode := 0;

    // recup de parametre
    s1 := uppercase(s1);
    while (s1 <> '') do
    begin
        s2 := ReadTokenSt(s1);
        if (s2 = 'MULTI') then mode := 1;
    end;

    tab := TPageControl(GetControl('TABLEMULTI')); if (not assigned(tab)) then exit;
    TS4 := TTabSheet(GetControl('TS4')); if (not assigned(TS4)) then exit;
    TS5 := TTabSheet(GetControl('TS5')); if (not assigned(TS5)) then exit;

    TS4.TabVisible := false;
    TS5.TabVisible := false;
    if (Mode = 0) then tab.ActivePage := TS4
    else tab.ActivePage := TS5;

    SetControlEnabled('TS1',(Mode = 0));
    SetControlEnabled('TS2',(Mode = 0));
    SetControlVisible('FListe',(Mode = 0));
    SetControlVisible('BGenerer',(Mode = 1));
    SetControlVisible('BAnnuler',(Mode = 0));
    SetControlVisible('BNouveau',(Mode = 0));
    SetControlVisible('BDelete',(Mode = 0));
    SetControlVisible('BInfoComp',(Mode = 0));
    SetControlVisible('BCopier',(Mode = 0));
    SetControlVisible('BImprimer',(Mode = 0));
    SetControlVisible('BValider',(Mode = 0));
    SetControlVisible('BAide',(Mode = 0));
    SetControlVisible('BFirst',(Mode = 0));
    SetControlVisible('BPrev',(Mode = 0));
    SetControlVisible('BNext',(Mode = 0));
    SetControlVisible('BLast',(Mode = 0));

    GeneCharge := true ;
    CbFlaguer := false ;

{$IFDEF EAGLCLIENT}
{$ELSE}
    ChangeSizeMemo(TMemoField(DS.FindField('SC_ATTRCOMP'))); // ?? SBO
{$ENDIF}           

    // Récupération des controles importants
    GRacines := THGrid(GetControl('GRacine',true));
    Racines := TGroupBox(GetControl('Racines',true));
    Libres := TGroupBox(GetControl('Libres',true));
    LibreEnt := TGroupBox(GetControl('LibreEnt',true));
    // Evènement sur la grille de saisie des codes racines
    GRacines.OnRowEnter := OnEnterGRacinesRow;
    GRacines.OnRowExit := OnExitGRacinesRow;
    GRacines.OnKeyPress := OnKeyPressGRacinesRow;
    
    // Evt boutons copier et infos
    TButton(GetControl('BGenerer',true)).onClick := OnClickBGenerer;
    TButton(GetControl('BCOPIER',true)).onClick := OnClickCopierScenario;
    TButton(GetControl('BINFOCOMP',true)).onClick := OnClickAfficherInfoComp;
    // Evt Checkbox Qtes
    TCheckBox( GetControl('LE7', True ) ).OnClick := OnClickQte ;
    TCheckBox( GetControl('LE8', True ) ).OnClick := OnClickQte ;

{$IFNDEF EAGLCLIENT}
    TDBCheckBox( GetControl('SC_REFINTERNE', True ) ).DataField     := 'SC_REFINTERNE';
    TDBCheckBox( GetControl('SC_LIBELLE', True ) ).DataField        := 'SC_LIBELLE';
    TDBCheckBox( GetControl('SC_REFEXTERNE', True ) ).DataField     := 'SC_REFEXTERNE';
    TDBCheckBox( GetControl('SC_AFFAIRE', True ) ).DataField        := 'SC_AFFAIRE';
    TDBCheckBox( GetControl('SC_REFLIBRE', True ) ).DataField       := 'SC_REFLIBRE';
    TDBCheckBox( GetControl('SC_DATEREFEXTERNE', True ) ).DataField := 'SC_DATEREFEXTERNE';
{$ENDIF EAGLCLIENT}

    // Valuer "Tous" dans les combos
//    THValComboBox(GetControl('SC_NATUREPIECE',true)).values[0] := '...';
    THValComboBox(GetControl('SC_USERGRP',true)).values[0] := '...';
    THValComboBox(GetControl('SC_QUALIFPIECE',true)).values[0] := '...';
    THValComboBox(GetControl('SC_ETABLISSEMENT',true)).values[0] := '...';
    // Mise en forme de la grille des Code racines
    GRacines.ColAligns[0] := TaCenter ;

    for i := 1 to 10 do GRacines.Cells[0,i] := IntToStr(i);
    // Evènement sur champs booléens pour les infos des lignes d'écritures
    for i := 0 to Racines.ControlCount-1 do
    begin
        C := Racines.Controls[i];
        if (C is TCheckBox) then TCheckBox(C).OnClick := ForceChange;
    end;
    // Evènement sur champs booléens pour les zones libres des lignes d'écritures
    for i := 0 to Libres.ControlCount-1 do
    begin
        C := Libres.Controls[i];
        if (C is TCheckBox) then TCheckBox(C).OnClick := ForceChange;
    end;
    // Evènement sur champs booleens pour les informations génériques libres d'entête
    for i := 0 to LibreEnt.ControlCount-1 do
    begin
        C := LibreEnt.Controls[i];
        if (C is TCheckBox) then TCheckBox(C).OnClick := ForceChange;
    end;
    // Paramétrage S3
{    if (EstSerie(S3) or EstSerie(S5)) then
      LimiteAccesControles
    else
}      AttribLesHint;
    // Première page active
    if (mode = 0) then SetActiveTabSheet('TS1')
    else SetActiveTabSheet('TS3');
//    Ecran.PopUpMenu := ADDMenuPop(Ecran.PopUpMenu,'',''); // ?? SBO

    GeneCharge := false;
    //SG6 11.03.05 Gestion en mode croisaxe on rend invisible l'axe préférentiel
    if VH^.AnaCroisaxe then
    begin
      SetControlVisible('TSC_NUMAXE',False);
      SetControlVisible('SC_NUMAXE', False);
    end;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 10/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOM_SUIVCPTA.OnClose;
begin
    Inherited;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 10/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOM_SUIVCPTA.OnCancelRecord;
begin
    Inherited;
end;

//==================================================
// Autres Evenements
//==================================================
{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 10/03/2003
Modifié le ... : 20/12/2006
Description .. : YMO
Mots clefs ... :
*****************************************************************}
procedure TOM_SUIVCPTA.ForceChange(Sender: TObject);
begin

// SBO 29/03/2007 : pb à l'ouverture...
// BVE 26.04.07 : plantage eAGL
if (DS = nil) or (DS.state=dsInactive) then Exit ;

{$IFDEF EAGLCLIENT}
    if (not GeneCharge) and (DS.state = dsBrowse) // Modif SBO 23/02/2004 : changement de mode que si nécessaire
      then DS.Edit;
{$ELSE}
    if (not GeneCharge) then ForceUpdate;
{$ENDIF}

// YMO 20/12/2006 Enregistrement systématique sur le click,...
// malgré le balayage automatique de tous les objets TH c'est nécessaire en 2 tiers,
// il faudrait revoir le fonctionnement pour ne mettre à jour que la case en question...
SauveComplement;

end;


{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 10/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOM_SUIVCPTA.OnEnterGRacinesRow(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
begin
    ChargeComplement;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 10/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOM_SUIVCPTA.OnExitGRacinesRow(Sender: TObject; Ou: Integer; var Cancel: Boolean; Chg: Boolean);
begin
    Radicaux[ou] := GRacines.Cells[1,Ou];
    ForceChange(nil);
    MajComplement(ou);
end;


{***********A.G.L.***********************************************
Auteur  ...... : YMO
Créé le ...... : 20/12/2006
Modifié le ... :   /  /
Description .. : Le fait de modifier une racine sans changer de ligne n'était pas détecté,
                même en validant (Onupdate pas déclenché)
Mots clefs ... :
*****************************************************************}
procedure TOM_SUIVCPTA.OnKeyPressGRacinesRow(Sender: TObject; var Key: Char);
begin
    //YMO 20/12/2006 Détection d'un changement si pas de modif par ailleurs, et sans changer de ligne
    if (not GeneCharge) then ForceUpdate;
    MajComplement(GRacines.Row);
    SauveComplement;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 10/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOM_SUIVCPTA.OnClickAfficherInfoComp(Sender: TObject);
var InfoComp : HTStrings;
begin
    InfoComp := HTStringList.Create;
{$IFDEF EAGLCLIENT}
    InfoComp.Text := (TMemoField(DS.FindField('SC_ATTRCOMP')).AsString) ;
{$ELSE}
    InfoComp.Assign(TMemoField(DS.FindField('SC_ATTRCOMP')));
{$ENDIF}

    if (InfoComp.Count <= 0) then GenereListe(InfoComp);
    ChercheAcces(InfoComp);

    if (CbFlaguer) then
    begin
        AttributComplementaires(InfoComp);
        if (DS.State = dsBrowse) then ForceUpdate;
{$IFDEF EAGLCLIENT}
        DS.FindField('SC_ATTRCOMP').asString := InfoComp.Text ;
{$ELSE}
        DS.FindField('SC_ATTRCOMP').Assign(InfoComp);
{$ENDIF}
    end
    else PGIInfo(MessageListe[18], Ecran.caption);

    InfoComp.Clear;
    InfoComp.Free;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 10/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOM_SUIVCPTA.OnClickCopierScenario(Sender: TObject);
var
    Q : TQuery;
    Num,i : integer;
    Champ, CodeN, Lib, req : String;
begin
    // Numéroter les enregistrements car pas de code unique
    Num := 0;
    Lib := '';

    Q := OpenSQL('Select SC_USERGRP, SC_JOURNAL, SC_NATUREPIECE, SC_QUALIFPIECE, SC_ETABLISSEMENT FROM SUIVCPTA ORDER BY SC_JOURNAL',false);
    While Not Q.EOF do
    begin
        Inc(Num);
        req := 'UPDATE SUIVCPTA SET SC_REFAUTO = "'+ IntToStr(Num) +'" WHERE SC_USERGRP = "' + Q.FindField('SC_USERGRP').asString + '" AND SC_JOURNAL = "' + Q.FindField('SC_JOURNAL').asString + '" AND SC_NATUREPIECE = "' + Q.FindField('SC_NATUREPIECE').asString + '" AND SC_QUALIFPIECE = "' + Q.FindField('SC_QUALIFPIECE').asString + '" AND SC_ETABLISSEMENT = "' + Q.FindField('SC_ETABLISSEMENT').asString + '"';
        ExecuteSQL(req);
        Q.Next;
    end;
    Ferme(Q);

    // Choisir le scénario
    CodeN := Choisir(MessageListe[19],'SUIVCPTA','SC_JOURNAL || " \ " || SC_NATUREPIECE || " \ " || SC_ETABLISSEMENT','SC_REFAUTO','','');
    if (CodeN = '') then Exit;

    // Copie des champs
    if (DS.State = dsBrowse) then ForceUpdate;
    Q := OpenSQL('SELECT * FROM SUIVCPTA WHERE SC_REFAUTO="' + CodeN + '"',true);
    if (Not Q.EOF) then
    begin
{$IFDEF EAGLCLIENT}
        for i := 0 to Q.FieldCount-1 do
{$ELSE}
        for i := 0 to Q.Fields.Count-1 do
{$ENDIF}
        begin
            Champ := Q.Fields[i].FieldName;
            if ((Champ = 'SC_JOURNAL') or (Champ = 'SC_NATUREPIECE') or (Champ = 'SC_ETABLISSEMENT') or (Champ = 'SC_QUALIFPIECE') or (Champ = 'SC_USERGRP')) then Continue;
            SetField(Champ,Q.Fields[i].AsVariant);
        end;
        InitTableaux;
    end;
    Ferme(Q);
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 10/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOM_SUIVCPTA.OnClickBGenerer(Sender: TObject);
begin
    EnregOkMulti;
end;

//==================================================
// Autres fonctions de la class
//==================================================
{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 10/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOM_SUIVCPTA.InitTableaux;
var
    i,j :  integer ;
    Compl : String[10];
    Lib : String[40];
begin
    // Valeurs par défaut
    FillChar(Radicaux,SizeOf(Radicaux),#0);
    FillChar(Complement,SizeOf(Complement),'-');
    FillChar(CompLibre,SizeOf(CompLibre),'-');
    FillChar(LibreEntete,SizeOf(LibreEntete),'-');
    // Si pas d'enregistrement, exit...
    if (DS.Eof and DS.Bof) then Exit ;

    for i := 1 to 10 do // Jusqu'à 10 racines par scénario...
    begin
        // Codes racines
        Radicaux[i] := Copy(GetField('SC_RADICAL'+IntToStr(i)),1,10);
        GRacines.Cells[1,i] := Radicaux[i];
        // Champs booléens pour les infos des lignes d'écritures
        Compl := GetField('SC_COMPLEMENTS'+IntToStr(i));
        for j := 1 to 10 do Complement[i,j] := Compl[j];
        // Champs booléens pour les zones libres des lignes d'écritures
        Lib := GetField('SC_COMPLIBRE'+IntToStr(i));
        for j := 1 to 40 do CompLibre[i,j] := Lib[j];
    end ;
    // Champs booleens pour les informations génériques libres d'entête
    Lib := GetField('SC_LIBREENTETE');
    for j := 1 to 40 do LibreEntete[j] := Lib[j];
    // Mise en place des données
    ChargeComplement;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 10/03/2003
Modifié le ... : 10/03/2003
Description .. : Mise en place dans l'interface des données contenues dans
Suite ........ : les tableaux correspondants aux paramètres traités en
Suite ........ : "masse"
Mots clefs ... :
*****************************************************************}
procedure TOM_SUIVCPTA.ChargeComplement;
var
    i : integer;
    C : TControl;
begin
    if (Ecran = nil) then Exit;

    // Champs booléens pour les infos des lignes d'écritures
    for i := 0 to Racines.ControlCount-1 do
    begin
        C := Racines.Controls[i];
        if (C is TCheckBox) then if (C.Tag > 0) then TCheckBox(C).Checked := (Complement[GRacines.Row,C.Tag] = 'X');
    end;
    // Champs booléens pour les zones libres des lignes d'écritures
    for i := 0 to Libres.ControlCount-1 do
    begin
        C := Libres.Controls[i];
        if (C is TCheckBox) then if (C.Tag > 0) then TCheckBox(C).Checked := (CompLibre[GRacines.Row,C.Tag] = 'X');
    end;
    // Champs booleens pour les informations génériques libres d'entête
    for i := 0 to LibreEnt.ControlCount-1 do
    begin
        C := LibreEnt.Controls[i];
        if (C is TCheckBox) then if (C.Tag > 0) then TCheckBox(C).Checked := (LibreEntete[C.Tag] = 'X');
    end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 10/03/2003
Modifié le ... :   /  /
Description .. : Mise en place dans le TDataSet (DS) des données saisies à
Suite ........ : l'interface  pour les paramètres traités en "masse"
Mots clefs ... :
*****************************************************************}
procedure TOM_SUIVCPTA.SauveComplement;
var
    i,j : Integer;
    Compl : String[10];
    Lib : String[40];
begin

  // SBO 29/03/2007 : pb à l'ouverture...
  if DS.state=dsInactive then Exit ;

    // Pour chacun des 10 codes racines...
    for i := 1 to 10 do
    begin
        // Code racine
        SetField('SC_RADICAL'+IntToStr(i),Radicaux[i]);
        // Champs booléens pour les infos des lignes d'écritures
        Compl := '--';
        for j := 3 to 10 do Compl := Compl+Complement[i,j];
        SetField('SC_COMPLEMENTS'+IntToStr(i),Compl);
        // Champs booléens pour les zones libres des lignes d'écritures
        Lib := '';
        for j := 1 to 40 do Lib := Lib+CompLibre[i,j];
        SetField('SC_COMPLIBRE'+IntToStr(i),Lib);
    end ;
    // Champs booleens pour les informations génériques libres d'entête
    Lib := '';
    for i := 1 to 40 do Lib := Lib+LibreEntete[i];
    SetField('SC_LIBREENTETE',Lib);
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 10/03/2003
Modifié le ... :   /  /
Description .. : Mise a jour des tableaux pour le code racine sélectionnés
Mots clefs ... :
*****************************************************************}
procedure TOM_SUIVCPTA.MajComplement(ou: Integer);
var
    i : Integer;
    C : TControl;
    Ch : Char;
begin
    // Mise a jour du code Racine
    Radicaux[ou] := GRacines.Cells[1,ou];
    // Champs booléens pour les infos des lignes d'écritures
    for i := 0 to Racines.ControlCount-1 do
    begin
        C := Racines.Controls[i];
        if (C is TCheckBox) and (C.Tag > 0) then
        begin
            if TCheckBox(C).Checked then Ch :='X'
            else Ch := '-';
            Complement[ou,C.Tag] := Ch;
        end;
    end;
    // Champs booléens pour les zones libres des lignes d'écritures
    for i := 0 to Libres.ControlCount-1 do
    begin
        C := Libres.Controls[i];
        if (C is TCheckBox) and (C.Tag > 0) then
        begin
            if TCheckBox(C).Checked then Ch := 'X'
            else Ch := '-';
            CompLibre[ou,C.Tag] := Ch;
        end;
    end;
    // Champs booleens pour les informations génériques libres d'entête
    for i := 0 to LibreEnt.ControlCount-1 do
    begin
        C := LibreEnt.Controls[i];
        if (C is TCheckBox) and (C.Tag > 0) then
        begin
            if TCheckBox(C).Checked then Ch := 'X'
            else Ch := '-';
            LibreEntete[C.Tag] := Ch;
        end;
    end;

  // Gestion accès contrôle des qtes
  GereAccesCtrlQte;

end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 10/03/2003
Modifié le ... :   /  /
Description .. : Vide les controles gérés en "masses"
Mots clefs ... :
*****************************************************************}
procedure TOM_SUIVCPTA.RAZRacines;
var
    i : integer;
    C: TControl;
begin
    // La grille des codes racines
    for i := 1 to 10 do GRacines.Cells[1,i] := '';
    // Champs booléens pour les infos des lignes d'écritures
    for i := 0 to Racines.ControlCount-1 do
    begin
        C := Racines.Controls[i];
        if (C is TCheckBox) then TCheckBox(C).Checked := false;
    end;
    // Champs booléens pour les zones libres des lignes d'écritures
    for i := 0 to Libres.ControlCount-1 do
    begin
        C := Libres.Controls[i];
        if (C is TCheckBox) then TCheckBox(C).Checked := false;
    end;
    // Champs booleens pour les informations génériques libres d'entête
    for i := 0 to LibreEnt.ControlCount-1 do
    begin
        C := LibreEnt.Controls[i];
        if (C is TCheckBox) then TCheckBox(C).Checked := false;
    end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 10/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOM_SUIVCPTA.AttribLesHint;
var
    S, CT, CL : String;
    VV : Boolean;
    i : integer;
begin
    VV := false;
    // Récupération du paramétrage des tables libres
    ChargeParamLibTOB;
    // Textes libres
    for i := 1 to 10 do
    begin
        CT := 'ET' + IntToStr(i);
        CL := 'LT' + IntToStr(i);
        S := 'E_LIBRETEXTE' + IntToStr(i-1);
        if (PersoChampAvecTOB(ParamLibTOB,'E',S,VV)) then
        begin
            SetControlProperty(CT,'Hint',S);
            SetControlProperty(CL,'Hint',S);
            AccesCoches(CT,CL,VV);
        end;
    end;
    // Tables libres
    for i := 1 to 4 do
    begin
        CT := 'ET' + IntToStr(i + 10);
        CL := 'LC' + IntToStr(i);
        S  := 'E_TABLE' + IntToStr(i-1);
        if (PersoChampAvecTOB(ParamLibTOB,'E',S,VV)) then
        begin
            SetControlProperty(CT,'Hint',S);
            SetControlProperty(CL,'Hint',S);
            AccesCoches(CT,Cl,VV);
        end;
    end;
    // Montants libres
    for i := 1 to 4 do
    begin
        CT := 'ET' + IntToStr(i+14);
        CL := 'LM' + IntToStr(i);
        S  := 'E_LIBREMONTANT' + IntToStr(i-1);
        if (PersoChampAvecTOB(ParamLibTOB,'E',S,VV)) then
        begin
            SetControlProperty(CT,'Hint',S);
            SetControlProperty(CL,'Hint',S);
            AccesCoches(CT,Cl,VV);
        end;
    end;
    // Choixs O/N libres
    for i := 1 to 2 do
    begin
        CT := 'ET' + IntToStr(i+18);
        CL := 'LB' + IntToStr(i);
        S  := 'E_LIBREBOOL' + IntToStr(i-1);
        if (PersoChampAvecTOB(ParamLibTOB,'E',S,VV)) then
        begin
            SetControlProperty(CT,'Hint',S);
            SetControlProperty(CL,'Hint',S);
            AccesCoches(CT,Cl,VV);
        end;
    end;
    // Dates libre
    CT := 'ET21';
    CL := 'LD1';
    S:='E_LIBREDATE';
    if PersoChampAvecTOB(ParamLibTOB,'E',S,VV) then
    begin
        SetControlProperty(CT,'Hint',S);
        SetControlProperty(CL,'Hint',S);
        AccesCoches(CT,Cl,VV);
    end;
    // Libération TOB
    ParamLibTOB.Free;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 10/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOM_SUIVCPTA.AccesCoches(CT, CL: String; VV: Boolean);
begin
    SetControlChecked(CT,VV);
    SetControlEnabled(CT,VV);
    SetControlChecked(CL,VV);
    SetControlEnabled(CL,VV);
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 10/03/2003
Modifié le ... :   /  /
Description .. : Charge les paramètres pour tables libres dans une TOB
Mots clefs ... :
*****************************************************************}
procedure TOM_SUIVCPTA.ChargeParamLibTOB;
begin
    ParamLibTOB := TOB.Create('Paramètres Tables Libres',nil,-1);
    ParamLibTOB.loadDetailDB('PARAMLIB','"E"','',nil,false);
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 10/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOM_SUIVCPTA.SelectionJournal;
var
    Q            : TQuery ;
    i            : integer ;
begin
    if Ecran = nil then Exit;
    // Maj du combo des natures de pieces
    Q := OpenSQL('Select J_NATUREJAL, J_MODESAISIE from JOURNAL Where J_JOURNAL="' + GetControlText('SC_JOURNAL') + '"',true);
    if Not Q.EOF then
      begin

      case CaseNatJal(Q.FindField('J_NatureJal').AsString) of
            tzJVente       : SetControlProperty('SC_NaturePiece','DataType','ttNatPieceVente');
            tzJAchat       : SetControlProperty('SC_NaturePiece','DataType','ttNatPieceAchat');
            tzJBanque      : SetControlProperty('SC_NaturePiece','DataType','ttNatPieceBanque');
            tzJEcartChange : SetControlProperty('SC_NaturePiece','DataType','ttNatPieceEcartChange');
            tzJOD          : SetControlProperty('SC_NaturePiece','DataType','ttNaturePiece');
        end;
//        THValComboBox(GetControl('SC_NATUREPIECE')).values[0] := '...';
      {YMO 13/12/2006 nécessaire uniquement en modification // en consultation, le 'OnChangeField' sur NATUREPIECE s'occupe
      de le rafraîchir // si on le laisse, le fait de faire le SetControlText remet la source de données en mode édition (version 2 tiers)}
      if (GetField('SC_NATUREPIECE') <> '') and (Ds.State in [dsInsert,dsEdit]) then
        SetControlText('SC_NATUREPIECE', GetField('SC_NATUREPIECE') );
      // MAJ accès zone pour journaux type Bor / LIB
      FBoModePiece := not ((Q.Findfield('J_MODESAISIE').AsString='BOR') or (Q.Findfield('J_MODESAISIE').AsString='LIB')) ;
      if ( TFFicheListe( Ecran ).FTypeAction <> taConsult ) then
        begin
          // Page scénario
          SetControlEnabled('SC_DATEOBLIGEE',     FBoModePiece);
          SetControlEnabled('SC_CONTROLETVA',     FBoModePiece);
          SetControlEnabled('SC_DOCUMENT',        FBoModePiece);
          SetControlEnabled('SC_VALIDE',          FBoModePiece);
          SetControlEnabled('SC_CONTROLEBUDGET',  FBoModePiece);
          // Page Entête
          SetControlEnabled('SC_REFINTERNE',      FBoModePiece);
          SetControlEnabled('SC_LIBELLE',         FBoModePiece);
          SetControlEnabled('SC_REFEXTERNE',      FBoModePiece);
          SetControlEnabled('SC_AFFAIRE',         FBoModePiece);
          SetControlEnabled('SC_REFLIBRE',        FBoModePiece);
          SetControlEnabled('SC_DATEREFEXTERNE',  FBoModePiece);
          for i := 1 to 21 do
            SetControlEnabled('ET'+IntToStr(i),   FBoModePiece);
        end;

      end;

    Ferme(Q);

//    If DS.State<>EtatInit then DS.State:=EtatInit;

end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 10/03/2003
Modifié le ... :   /  /
Description .. : Consruit chaîne descriptive du paramètrage d'infos
Suite ........ : complémentaires avec la formule : StOrigine + ":C.Hint;" +
Suite ........ : MemotSt + ";C.Enabled;C.Checked"
Mots clefs ... :
*****************************************************************}
function TOM_SUIVCPTA.AjouteAcces(C:TCustomCheckBox; MemoSt, StOrigine, Radic: String; QuelTab: Byte): String;
begin
    if (C.Enabled) then
    begin
        MemoSt := MemoSt + 'X;';

        if (Radic = '') then
        begin
            if (TCheckBox(C).Checked) then
            begin
                MemoSt := MemoSt + 'X;';
                CbFlaguer := true;
            end
            else MemoSt := MemoSt + '-;';
        end
        else
        begin
            Case QuelTab of
                1 : if (Complement[StrToInt(Radic),TCheckBox(C).Tag] = 'X') then
                    begin
                        MemoSt := MemoSt + 'X;';
                        CbFlaguer := true ;
                    end
                    else MemoSt:=MemoSt+'-;';
                2 : if (CompLibre[StrToInt(Radic),TCheckBox(C).Tag] = 'X') then
                    begin
                        MemoSt := MemoSt + 'X;';
                        CbFlaguer := true ;
                    end
                    else MemoSt := MemoSt + '-;';
            end;
        end;
    end
    else MemoSt := MemoSt + '-;-;';

    StOrigine := StOrigine + ':' + C.Hint + ';';
    MemoSt := StOrigine + MemoSt ;
    Result := MemoSt ;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 10/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOM_SUIVCPTA.ChercheAcces(Li: HTStrings);
var
    i,j : Integer;
    StChaine, StOrigine, StIndex, Radic, MemoSt, StControl, stModif, stOblig, stValDef, stValTL : String;
    cbCac : TControl;
begin
    for i := 0 to Li.Count-1 do
    begin
        StChaine := Li.Strings[i];           // Chaîne complète
        StOrigine := ReadTokenSt(StChaine);  // Origine (Entete libre ou Radical X)
        StControl := ReadTokenSt(StChaine);  // Nom du control
        // Décomposition et Recomposition de la chaîne MemoSt pour éviter grossissement (SBO)
        stOblig := ReadTokenSt(StChaine);
        stModif := ReadTokenSt(StChaine);
        stValDef := ReadTokenSt(StChaine);
        stValTL := ReadTokenSt(StChaine);
        MemoSt := StControl + ';' + stOblig + ';' + stModif + ';' + stValDef + ';' + stValTL +';';

        // Entête : Infos d'entête d'ecritures
        if (Pos('SC_',StControl) = 1) then
        begin
            cbCac := GetControl(StControl);
            if (cbCac <> Nil) then Li.Strings[i] := AjouteAcces(TCustomCheckBox(cbCac),MemoSt,StOrigine,'',0);
            continue;
        end;
        // Entête : Infos génériques libres
        if (Pos('ET',StControl) = 1) then
        begin
            cbCac := GetControl(StControl);
            if (cbCac <> Nil) then Li.Strings[i] := AjouteAcces(TCustomCheckBox(cbCac),MemoSt,StOrigine,'',0);
            continue;
        end;

        // Racines : informations lignes d'écritrures
        if ((Pos('LE',StControl) = 1) and (Pos('SC_',StControl) <= 0)) then
        begin
            StIndex := Copy(StControl,3,Length(StControl) - 2);
            for j := 0 to Racines.ControlCount-1 do
            begin
                cbCac := Racines.Controls[j];
                if (cbCac is TCheckBox) then if (TCheckBox(cbCac).Tag = StrToInt(StIndex)) then
                begin
                    if (Pos('SC_RADICAL',StOrigine) = 1) then Radic := Copy(StOrigine,11,2)
                    else Radic := '';
                    Li.Strings[i] := AjouteAcces(TCusTomCheckBox(cbCac),MemoSt,StOrigine,Radic,1);
                    Break;
                end;
            end;
            continue;
        end;

        // Racines : zones libres lignes d'écritrures
        if (Pos('ZLL',StControl) = 1) then
        begin
            StIndex := Copy(StControl,4,Length(StControl) - 3);
            for j := 0 to Libres.ControlCount-1 do
            begin
                cbCac := Libres.Controls[j];
                if (cbCac is TCheckBox) then if (TCheckBox(cbCac).Tag = StrToInt(StIndex)) then
                begin
                    if (Pos('SC_RADICAL',StOrigine) = 1) then Radic := Copy(StOrigine,11,2)
                    else Radic := '';
                    Li.Strings[i] := AjouteAcces(TCusTomCheckBox(cbCac),MemoSt,StOrigine,Radic,2);
                    Break;
                end;
            end;
        end;
    end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 10/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOM_SUIVCPTA.GenereListe(Li: HTStrings);
var
    i,j : Integer;
    entete : TWinControl;
begin
    Entete := TWinControl(GetControl('ENTETE'));
    for i := 0 to Entete.ControlCount - 1 do if (Entete.Controls[i] is TCustomCheckBox) then Li.Add('SC_LIBREENTETE;' + Entete.Controls[i].Name + ';-;X;;;');

    for i := 1 to 21 do Li.Add('SC_LIBREENTETE;ET' + IntToStr(i) + ';-;X;;;');

    for i := 1 to 10 do
    begin
        for j := 1 to 7 do Li.Add('SC_RADICAL'+IntToStr(i) + ';LE' + IntToStr(j+2) + ';-;X;;;');
        for j := 1 to 21 do Li.Add('SC_RADICAL'+IntToStr(i) + ';ZLL' + IntToStr(j) + ';-;X;;;');
    end ;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 10/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOM_SUIVCPTA.VerifierAcces(Tout: boolean);
begin
    if (Ecran = nil) then Exit;
    if ( TFFicheListe( Ecran ).FTypeAction = taConsult ) then Exit ;


    if ((GetField('SC_QUALIFPIECE') = 'N') or (GetField('SC_QUALIFPIECE') = '') or (GetField('SC_QUALIFPIECE') = '...')) then
    begin
        SetControlEnabled('SC_VALIDE',FBoModePiece);
        SetControlEnabled('SC_LETTRAGESAISIE',true);
    end
    else
    begin
        SetControlEnabled('SC_VALIDE',false);
        SetControlEnabled('SC_LETTRAGESAISIE',false);
        if (Tout) then
        begin
            SetControlChecked('SC_VALIDE',false);
            SetControlChecked('SC_LETTRAGESAISIE',false);
        end;
    end;

    if ((GetControlText('SC_RIB')='') and Tout) then SetControlText('SC_RIB','...');
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 10/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
{
procedure TOM_SUIVCPTA.LimiteAccesControles;
var
    i : Integer ;
begin
    // Zones invisibles en S3
{ // plus de gestion des série
    if (EstSerie(S3)) then
    begin
        SetControlVisible('SC_QUALIFPIECE',false);
        SetControlVisible('TSC_QUALIFPIECE',false);
        SetControlVisible('SC_ETABLISSEMENT',false);
        SetControlVisible('TSC_ETABLISSEMENT',false);
        SetControlVisible('SC_CONTROLEBUDGET',false);
        SetControlVisible('SC_NUMAXE',false);
        SetControlVisible('TSC_NUMAXE',false);
        SetControlVisible('SC_DOCUMENT',false);
        SetControlVisible('TSC_DOCUMENT',false);
    end ;

    // --> Panel LibreEnt :
    // Disable des champs booleens pour les informations génériques libres d'entête
    for i := 1 to 21 do SetControlEnabled('ET'+IntToStr(i), false);
    SetControlEnabled('Label7', false);
    SetControlEnabled('Label9', false);
    SetControlEnabled('Label10', false);
    SetControlEnabled('Label11', false);
    // Acces tables libres
    SetControlEnabled('ET11', true);
//    if (EstSerie(S5)) then
      SetControlEnabled('ET12', true);

    // --> Panel Libres :
    // Disable Champs booléens pour les zones libres des lignes d'écritures
    for i := 1 to 10 do SetControlEnabled('LT'+ IntToStr(i), false);
    for i := 1 to 4 do SetControlEnabled('LC' + IntToStr(i), false);
    for i := 1 to 4 do SetControlEnabled('LM' + IntToStr(i), false);
    for i := 1 to 2 do SetControlEnabled('LB' + IntToStr(i), false);
    SetControlEnabled('LD1', false);
    SetControlEnabled('Label1', false);
    SetControlEnabled('Label3', false);
    SetControlEnabled('Label5', false);
    SetControlEnabled('Label4', false);

    // Acces tables libres
    SetControlEnabled('LC1', true);
//    if (EstSerie(S5)) then
      SetControlEnabled('LC2', true);
//    if (EstSerie(S3)) then
    begin
        SetControlEnabled('LE7',false);
        SetControlEnabled('LE8',false);
    end ;
end;
}
{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 10/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
function TOM_SUIVCPTA.EnregOkMulti : Boolean;
var
    lStJournal,lStNature,lStJal,lStNat : string;
    lTOB : TOB;
    i : integer;
    lJournal : TZListJournal;
    JOURNALMULTI,NATUREPIECEMULTI : THMultiValComboBox;
begin
    result := false;
    lJournal := TZListJournal.Create;

    // recup des control
    JOURNALMULTI := THMultiValComboBox(GetControl('JOURNALMULTI')); if (not assigned(JOURNALMULTI)) then exit;
    NATUREPIECEMULTI := THMultiValComboBox(GetControl('NATUREPIECEMULTI')); if (not assigned(NATUREPIECEMULTI)) then exit;

    // si vide
    if (JOURNALMULTI.text = '') then
    begin
        PGIBox(TraduireMemoire(MessageListe[10]),TraduireMemoire(Ecran.Caption));
        Exit;
    end;
    if (NATUREPIECEMULTI.text = '') then
    begin
        PGIBox(TraduireMemoire(MessageListe[11]),TraduireMemoire(Ecran.Caption));
        Exit;
    end;

    // creation de la TOB
    lTOB := TOB.create('SUIVCPTA',nil,-1);
    lTOB.AddChampSup('_ETAT',false);

    // si tous
    if (JOURNALMULTI.Tous) then
    begin
        JOURNALMULTI.text := '';
        for i := 0 to JOURNALMULTI.Values.Count-1 do JOURNALMULTI.text := JOURNALMULTI.text+JOURNALMULTI.Values[i]+';';
    end;
    if (NATUREPIECEMULTI.Tous) then
    begin
        NATUREPIECEMULTI.text := '';
        for i := 0 to NATUREPIECEMULTI.Values.Count-1 do NATUREPIECEMULTI.text := NATUREPIECEMULTI.text+NATUREPIECEMULTI.Values[i]+';';
    end;

    // progress
    InitMoveProgressForm(ecran,'Génération en cours...','Génération en cours',10,true,false);

    // traitement !
    try
        if (DS.State in [dsBrowse]) then DS.Insert;
        MajComplement(GRacines.Row);
        SauveComplement;
        TOBM(lTOB).ChargeMvt(DS);

        lTOB.PutValue('SC_ATTRCOMP',GetField('SC_ATTRCOMP').AsString);
        lTOB.PutValue('SC_USERGRP','...');
        lTOB.PutValue('SC_ETABLISSEMENT','...');
        lTOB.PutValue('SC_QUALIFPIECE','...');

        lStJournal := JOURNALMULTI.text;
        lStJal := readtokenst(lStJournal);

        i := 0;

        while (lStJal <> '') do
        begin
            lStNature := NATUREPIECEMULTI.Text;
            lStNat := readtokenst(lStNature);
            while (lStNat <> '') do
            begin
                Inc(i);
                if i = 10 then i := 1;
                lJournal.Load([lStJal]);
                if (NATUREJALNATPIECEOK(lJournal.GetValue('J_NATUREJAL'),lStNat)) then
                begin
                    MoveCurProgressForm('journal : ' + lStJal + ' nature de pièce :' + lStNat);
                    lTOB.PutValue('SC_JOURNAL',UpperCase(lStJal));
                    lTOB.PutValue('SC_NATUREPIECE', UpperCase(lStNat));
                    lTOB.SetAllModifie(true);
                    lTOB.ChargeCle1;
                    lTOB.DeleteDB(false);
                    lTOB.InsertDB(nil);
                end;
                lStNat := readtokenst(lStNature);
            end ;
            lstJal := readtokenst(lStJournal);
        end;
    finally
    	FiniMoveProgressForm;
        FreeAndNil(lTOB);
        FreeAndNil(lJournal);
    end;
end;

procedure TOM_SUIVCPTA.OnClickQte(Sender: TObject);
begin
  if not Assigned(sender) then Exit ;
  if TCheckBox(Sender).Checked then
    SetControlEnabled('SC_CONTROLEQTE', True ) ;
end;

{***********A.G.L.***********************************************
Auteur  ...... : SBO
Créé le ...... : 24/03/2006
Modifié le ... :   /  /
Description .. : Restourne True si la saisie des qtyauntités à la ligne est
Suite ........ : demandé
Mots clefs ... :
*****************************************************************}
function TOM_SUIVCPTA.EstSaisieQte : Boolean ;
var i : integer ;
begin
  // Test pour contrôle qtés... décoché si accessible
  result := false ;
  for i := 1 to 10 do
    if ( Complement[i, 7]<>'-' ) or ( Complement[i, 8]<>'-' ) then
      begin
      result := true ;
      exit ;
      end ;
end;

procedure TOM_SUIVCPTA.GereAccesCtrlQte;
begin
  if EstSaisieQte then
    SetControlEnabled( 'SC_CONTROLEQTE', True )
    else begin
         SetControlEnabled( 'SC_CONTROLEQTE', False ) ;
         if ( DS.State <> dsBrowse ) and  ( GetField('SC_CONTROLEQTE') = 'X' ) then
           SetField('SC_CONTROLEQTE', '-' ) ;
         end ;
end;

{JP 15/10/07 : FQ 16149 : gestion des réstrictions Etablissements et à défaut des ParamSoc
{---------------------------------------------------------------------------------------}
procedure TOM_SUIVCPTA.GereEtablissement;
{---------------------------------------------------------------------------------------}
begin
  {Si l'on ne gère pas les établissement ...}
  if not VH^.EtablisCpta  then begin
    {... on affiche l'établissement par défaut}
    SetControlText('SC_ETABLISSEMENT', VH^.EtablisDefaut);
    SetField('SC_ETABLISSEMENT', VH^.EtablisDefaut);
    {... on désactive la zone}
    SetControlEnabled('SC_ETABLISSEMENT', False);
  end

  {On gère l'établissement, donc ...}
  else begin
    {... On commence par regarder les restrictions utilisateur}
    PositionneEtabUser(GetControl('SC_ETABLISSEMENT'));
    {... s'il n'y a pas de restrictions, on reprend le paramSoc
     JP 25/10/07 : FQ 19970 : Finalement on oublie l'option de l'établissement par défaut
    if GetControlText('SC_ETABLISSEMENT') = '' then begin
      {... on affiche l'établissement par défaut
      SetControlText('SC_ETABLISSEMENT', VH^.EtablisDefaut);
      SetField('SC_ETABLISSEMENT', VH^.EtablisDefaut);
      {... on active la zone
      SetControlEnabled('SC_ETABLISSEMENT', True);
    end;}
  end;
end;


//================================================================================
// Initialization
//================================================================================
Initialization
    registerclasses([TOM_SUIVCPTA]);
end.
