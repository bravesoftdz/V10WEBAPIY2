{***********UNITE*************************************************
Auteur  ...... : 
Créé le ...... : 30/06/2003
Modifié le ... :   /  /
Description .. : Source TOM de la TABLE : AXE (AXE)
Mots clefs ... : TOM;AXE
*****************************************************************}

Unit CPAXE_TOM;

//================================================================================
// Interface
//================================================================================
Interface

Uses
    StdCtrls, 
    Controls, 
    Classes, 
{$IFDEF EAGLCLIENT}
    MaineAGL,
    eFiche,
    eFichList,
{$ELSE}
    FE_Main,
    db,
    {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
    Fiche,
    FichList,
{$ENDIF}
    forms,
    sysutils,
    ComCtrls,
    HCtrls,
    Ent1,
    HEnt1,
    UtilPGI,
    HMsgBox,
    UTOM,
    ParamSoc,
    UTob  ,UentCommun
    ;

//==================================================
// Externe
//==================================================
Procedure FicheAxeAna(Lequel : String);

//==================================================
// Definition de class
//==================================================
Type
    TOM_AXE = Class (TOM)
        procedure OnNewRecord               ; override ;
        procedure OnDeleteRecord            ; override ;
        procedure OnUpdateRecord            ; override ;
        procedure OnAfterUpdateRecord       ; override ;
        procedure OnLoadRecord              ; override ;
        procedure OnChangeField(F : TField) ; override ;
        procedure OnArgument(S : String)    ; override ;
        procedure OnClose                   ; override ;
        procedure OnCancelRecord            ; override ;
    private
        LgSect : integer;
        ChangeSoc : boolean;

        procedure OnChangeX_BOURREANA(Sender : TObject);
        procedure OnChangeX_LONGSECTION(Sender : TObject);
        procedure OnDblClickX_SECTIONATTENTE(Sender : TObject);
        procedure OnEnterX_SECTIONATTENTE(Sender : TObject);
        procedure OnKeyPressX_SECTIONATTENTE(Sender : TObject ; var Key : Char);

        function  BourrageOk : Boolean;
        function  GeneAttenteOk : Boolean;
        function  GeneralNonVentilerSurAxe : Boolean;
        function  LongueurSectionOk : Boolean;
        function  SectionAttenteOk : Boolean;

        procedure RafraichiV;
    end;

//================================================================================
// Implementation
//================================================================================
Implementation

uses
  {$IFDEF MODENT1}
  CPTypeCons,
  CPProcMetier,
  {$ENDIF MODENT1}
  uProcGen;



//==================================================
// Definition des variable
//==================================================
var
    MESS : array [0..6] of string = (
        {00}    '0;Axes analytiques;Vous devez renseigner les règles de saisie.;W;O;O;O;',
        {01}    '1;Axes analytiques;Vous devez renseigner une section d''attente;W;O;O;O;',
        {02}    '2;Axes analytiques;Ce code section n''est pas rattaché à cet axe;W;O;O;O;',
        {03}    '3;Axes analytiques;Ce compte général n''est pas ventilable sur cet axe;W;O;O;O;',
        {04}    '4;Axes analytiques;La longueur des sections n''est plus modifiable.;W;O;O;O;',
        {05}    '5;Axes analytiques;La section d''attente doit être renseignée. Des comptes généraux sont ventilés sur cet axe.;W;O;O;O;',
        {06}    '6;Axes analytiques;Vous devez renseigner un caractère de bourrage;W;O;O;O;'
        );

//==================================================
// fonctions hors class
//==================================================
{***********A.G.L.***********************************************
Auteur  ...... : ?
Créé le ...... : 01/06/2007
Modifié le ... :   /  /    
Description .. : - LG - 01/06/2007 - suppression de la Serie S3
Mots clefs ... : 
*****************************************************************}
Procedure FicheAxeAna(Lequel : String);
begin
    if _Blocage(['nrCloture','nrBatch'],true,'nrBatch') then exit;

// LG -01/06/2007 - il n'y a plus de serie S3    
//    if EstSerie(S3) then range := 'A1'
//    else range := '';

    AGLLanceFiche('CP','CPAXES','',Lequel,'ACTION=MODIFICATION');
    
    _Bloqueur('nrBatch',false);
end;

//==================================================
// Evenements par default de la TOM
//==================================================
{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 01/07/2003
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... :
*****************************************************************}
procedure TOM_AXE.OnNewRecord;
begin
    Inherited;

    SetField('X_COMPTABLE','X');
    SetField('X_MODEREOUVERTURE','P');
    SetField('X_LONGSECTION',3);
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 01/07/2003
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOM_AXE.OnDeleteRecord;
begin
    Inherited;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 01/07/2003
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOM_AXE.OnUpdateRecord;
begin
    Inherited;

    LastError := 1;

    if DS.state in [dsEdit,dsInsert] then
    BEGIN
        if (not LongueurSectionOk) then begin SetControlText('X_LONGSECTION',IntToStr(LgSect)); exit; end;
        if (not SectionAttenteOk) then exit;
        if (not GeneralNonVentilerSurAxe) then exit;
        if (not GeneAttenteOk) then exit;
        if (not BourrageOk) then exit;
    END;

    LastError := 0;

    AvertirTable('AXE') ;
    RafraichiV;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 01/07/2003
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOM_AXE.OnAfterUpdateRecord;
begin
    Inherited;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 01/07/2003
Modifié le ... : 28/07/2006
Description .. : YMO : Affichage des comptes généraux ventilables sur
Suite ........ : L'axe en cours uniquement
Suite ........ : - FB 18630 - 28/07/2006 - On devrait interdire de faire un 
Suite ........ : changement de longueur de section sur l'axe paramétré
Suite ........ : pour la TVA.
Mots clefs ... : 
*****************************************************************}
procedure TOM_AXE.OnLoadRecord;
var
    fb : TFichierBase;
    OkOk : Boolean;
    i : Integer;
    axe : string;
    tmp : char;
begin
    Inherited;

    SetControlEnabled('X_SECTIONATTENTE',true);
    SetControlEnabled('BDelete',true);

    axe := GetField('X_AXE');
    if (Length(axe) >= 2) then tmp := axe[2] else tmp := '1';

    TFFicheListe(Ecran).OM.ModifAutorisee := not ( GetParamSocSecur('SO_CPPCLSAISIETVA',false) and ( axe = GetParamSocSecur('SO_CPPCLAXETVA', '') ) ) ;

    Case tmp of
        '1' : begin
                SetControlProperty('X_SECTIONATTENTE','DataType','TZSECTION');
                SetControlProperty('X_GENEATTENTE','DataType','TZGVENTIL1');
              end;
        '2' : begin
                SetControlProperty('X_SECTIONATTENTE','DataType','TZSECTION2');
                SetControlProperty('X_GENEATTENTE','DataType','TZGVENTIL2');
              end;
        '3' : begin
                SetControlProperty('X_SECTIONATTENTE','DataType','TZSECTION3');
                SetControlProperty('X_GENEATTENTE','DataType','TZGVENTIL3');
              end;
        '4' : begin
                SetControlProperty('X_SECTIONATTENTE','DataType','TZSECTION4');
                SetControlProperty('X_GENEATTENTE','DataType','TZGVENTIL4');
              end;
        '5' : begin
                SetControlProperty('X_SECTIONATTENTE','DataType','TZSECTION5');
                SetControlProperty('X_GENEATTENTE','DataType','TZGVENTIL5');
              end;
    end;

    fb := AxeToFb(GetField('X_AXE'));

    OkOk := true;
    For i:=1 To MaxSousPlan Do if (VH^.SousPlanAxe[fb,i].Code <> '') then OkOk := false;

    SetControlEnabled('X_STRUCTURE',OkOk);
    LgSect := StrToInt(GetField('X_LONGSECTION'));
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 01/07/2003
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOM_AXE.OnChangeField(F : TField);
begin
    Inherited;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 01/07/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOM_AXE.OnArgument(S : String);
var
    X_BOURREANA : THEdit;
    X_LONGSECTION : THSpinEdit;
    X_SECTIONATTENTE : THEdit;
    q : tquery;
begin
    Inherited;

    ChangeSoc := false;

    // ?? comment traduire ca en eAGL ??
//    if EstSerie(S3) then
//    begin
//        DS.SetRange(['A1'],['A1']);
//        SetControlVisible('X_SAISIETRANCHE',false);
//    end;

    if EstSerie(S3) then SetControlVisible('X_SAISIETRANCHE',false);

    // recup des control
    X_BOURREANA := THEdit(GetControl('X_BOURREANA',true));
    X_LONGSECTION := THSpinEdit(GetControl('X_LONGSECTION',true));
    X_SECTIONATTENTE := THEdit(GetControl('X_SECTIONATTENTE',true));

    // assignation des fonction o control
    X_BOURREANA.OnChange := OnChangeX_BOURREANA;
    X_LONGSECTION.OnChange := OnChangeX_LONGSECTION;
    X_SECTIONATTENTE.OnDblClick := OnDblClickX_SECTIONATTENTE;
    X_SECTIONATTENTE.OnEnter := OnEnterX_SECTIONATTENTE;
    X_SECTIONATTENTE.OnKeyPress := OnKeyPressX_SECTIONATTENTE;
            
    // verification du nb d'axe
    q := OpenSQL('SELECT COUNT(*) FROM AXE',true);
    if ((q.eof) or (q.Fields[0].AsInteger < 5)) then SetControlEnabled('Binsert',true);
    ferme(q);  
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 01/07/2003
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOM_AXE.OnClose;
begin
    Inherited;

    if (ChangeSoc) then ChargeMagHalley;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 01/07/2003
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOM_AXE.OnCancelRecord;
begin
    Inherited;
end;

//==================================================
// Autres Evenements
//==================================================
{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 01/07/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOM_AXE.OnDblClickX_SECTIONATTENTE(Sender : TObject);
begin
    inherited;

    ChangeSoc := true;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 01/07/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOM_AXE.OnKeyPressX_SECTIONATTENTE(Sender : TObject ; var Key : Char);
begin
    inherited;

    ChangeSoc := true;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 01/07/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOM_AXE.OnChangeX_LONGSECTION(Sender : TObject);
begin
    inherited;

    SetControlEnabled('X_SECTIONATTENTE',false);
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 01/07/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOM_AXE.OnChangeX_BOURREANA(Sender : TObject);
begin
    inherited;

    SetControlEnabled('X_SECTIONATTENTE',false);
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 01/07/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOM_AXE.OnEnterX_SECTIONATTENTE(Sender : TObject);
Var
    fb : TFichierBase ;
begin
    inherited;

    fb := AxeToFb(GetControlText('X_AXE'));
    VH^.Cpta[fb].Lg := StrToInt(GetControlText('X_LONGSECTION'));
end;

//==================================================
// Autres fonctions de la class
//==================================================
{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 01/07/2003
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
Function TOM_AXE.BourrageOk : Boolean;
begin
    result := false;

    if (GetControlText('X_BOURREANA') = '') or (GetControlText('X_BOURREANA') = ' ') then
    begin
        HShowMessage(MESS[6],'','');
        SetFocusControl('X_BOURREANA');
        exit;
    end;

    result := true;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 01/07/2003
Modifié le ... :   /  /
Description .. : 
Mots clefs ... : 
*****************************************************************}
Function TOM_AXE.GeneralNonVentilerSurAxe : Boolean;
var
    Q : TQuery;
    axe : string;
    tmp : char;
begin
    result := true;

    if (GetControlText('X_SECTIONATTENTE') <> '') then exit;

    axe := GetField('X_AXE');
    if (Length(axe) >= 2) then tmp := axe[2] else tmp := '1';

    Q := OpenSql('SELECT G_GENERAL FROM GENERAUX WHERE G_VENTILABLE' + tmp + '="X"',true);
    if (not Q.Eof) then
    begin
        HShowMessage(MESS[5],'','');
        SetFocusControl('X_SECTIONATTENTE');
        result := false;
    end;
    Ferme(Q);
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 01/07/2003
Modifié le ... :   /  /
Description .. : 
Mots clefs ... : 
*****************************************************************}
Function TOM_AXE.LongueurSectionOk : Boolean;
var
    Q : TQuery;
    fb : TFichierBase;
begin
    result := true;

    if (LgSect = StrtoInt(GetControlText('X_LONGSECTION'))) then exit;

    Q := OpenSql('SELECT S_SECTION FROM SECTION WHERE S_AXE="' + GetControlText('X_AXE') + '"',true);
    if (not Q.Eof) then
    begin
        if (Length(Q.Fields[0].AsString) <> StrToInt(GetControlText('X_LONGSECTION'))) then
        begin
            HShowMessage(MESS[4],'','');
            SetFocusControl('X_LONGSECTION');
            result := false;
            Ferme(Q);
            exit;
        end;
    end;
    Ferme(Q);

    fb := AxeToFb(GetControlText('X_AXE'));
    VH^.Cpta[fb].Lg := StrtoInt(GetControlText('X_LONGSECTION')) ; // Lhe le 17/03/1998
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 01/07/2003
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
Function TOM_AXE.GeneAttenteOk : Boolean;
var
    Q : TQuery;
    axe : string;
    tmp : char;
begin
    result := true;
    
    if (GetControlText('X_GENEATTENTE') = '') then exit;

    axe := GetField('X_AXE');
    if (Length(axe) >= 2) then tmp := axe[2] else tmp := '1';

    Q := OpenSql('SELECT G_GENERAL FROM GENERAUX WHERE G_GENERAL="' + GetControlText('X_GENEATTENTE') + '" AND G_VENTILABLE' + tmp + '="X"',true);
    if (Q.Eof) then
    begin
        HShowMessage(MESS[3],'','');
        SetFocusControl('X_GENEATTENTE');
        result := false;
    end;
    Ferme(Q);
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 01/07/2003
Modifié le ... :   /  /
Description .. : 
Mots clefs ... :
*****************************************************************}
Function TOM_AXE.SectionAttenteOk : Boolean;
var
    Q : TQuery;
begin
    result := true;
    
    if (GetControlText('X_SECTIONATTENTE') = '') then exit;

    Q := OpenSql('SELECT S_SECTION FROM SECTION WHERE S_SECTION="' + GetControlText('X_SECTIONATTENTE') + '" AND S_AXE="' + GetControlText('X_AXE') + '"',true);
    if (Q.Eof) then
    begin
        HShowMessage(MESS[2],'','');
        SetFocusControl('X_SECTIONATTENTE');
        result := false;
    end;
    Ferme(Q);
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 01/07/2003
Modifié le ... :   /  /    
Description .. :
Mots clefs ... : 
*****************************************************************}
Procedure TOM_AXE.RafraichiV;
var
    fb : TFichierBase;
begin
    fb := AxeToFb(getControlText('X_AXE'));

    VH^.Cpta[fb].Lg           := StrToInt(GetControlText('X_LONGSECTION'));
    VH^.Cpta[fb].Chantier     := (GetControlText('X_CHANTIER') = 'X');
    VH^.Cpta[fb].Structure    := (GetControlText('X_STRUCTURE') = 'X');
    VH^.Cpta[fb].Attente      := GetControlText('X_SECTIONATTENTE');
    VH^.Cpta[fb].Cb           := StrTOChr(GetControlText('X_BOURREANA'));
    VH^.Cpta[fb].AxGenAttente := GetControlText('X_GENEATTENTE');          // FQ20009 : raffraichissement des info de l'axe pour les autres traitements SBO 22/05/2007

end;

//================================================================================
// Initialization
//================================================================================
Initialization
    registerclasses([TOM_AXE]);
end.
