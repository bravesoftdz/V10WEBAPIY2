{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 24/04/2003
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : CPQRTVACTRL ()
Mots clefs ... : TOF;CPQRTVACTRL
*****************************************************************}

Unit CPQRTVACTRL_TOF;

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
    eQRS1,
{$ELSE}
    FE_Main,
    db,
    {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
    QRS1,
{$ENDIF}
    forms,
    sysutils,
    ComCtrls,
    HCtrls,
    Ent1,
    HEnt1,
    UTOF,
    extctrls,
    Windows,
    HTB97,
    TofMeth {JP 15/10/07 : FQ 16149 : le tof hérite maintenant de TOF_METH},
    paramsoc
    ;

//==================================================
// Externe
//==================================================
procedure EditionTvaHT;

//==================================================
// Definition de class
//==================================================
Type
    TOF_CPQRTVACTRL = Class(TOF_METH)
    public
        procedure OnNew                  ; override ;
        procedure OnDelete               ; override ;
        procedure OnUpdate               ; override ;
        procedure OnLoad                 ; override ;
        procedure OnArgument(S : String) ; override ;
        procedure OnDisplay              ; override ;
        procedure OnClose                ; override ;
        procedure OnCancel               ; override ;
    private
        procedure OnChangeFEXERCICE(Sender : TObject);
        procedure OnChangeFDEVISE(Sender : TObject);

        procedure OnChangeFNATURECPT(Sender : TObject);

        procedure OnClickFMONTANT(Sender : TObject);
        procedure OnClickFTYPETVA(Sender : TObject);
        procedure OnClickFNATTVA(Sender : TObject);

        procedure OnClickMONVALIDER(Sender : TObject);
        procedure FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);

        function  GetCritere : string;
    public
    private
        FGENERAL,FGENERAL_ : THEdit;
        FDEVISE,FNATURECPT,FEXERCICE : THValComboBox;

        critere : boolean;
    end;

//================================================================================
// Implementation
//================================================================================
Implementation

uses
  {$IFDEF MODENT1}
  ULibExercice,
  CPTypeCons,
  {$ENDIF MODENT1}
  HMsgBox;



//==================================================
// Definition de variables
//==================================================
var
    Mess : array [0..21] of string = (
    {00}    'Etat de contrôle de TVA sur encaissements',
    {01}    'Etat de contrôle de TVA sur immobilisations',
    {02}    'Etat de contrôle de TVA sur décaissements',
    {03}    '3;?caption?;Vous devez d''abord sélectionner une devise particulière.;E;O;O;O;',

    {00}    'Compte',
    {01}    'REPORT TOTAL GENERAL',
    {02}    'REPORT COMPTE',
    {06}    'Date cpta.',
    {07}    'Journal',
    {08}    'Pièce / Ligne',
    {09}    'Référence interne',
    {10}    'Libellé',
    {11}    'TVA',
    {12}    'D/E',
    {13}    'sur débit',
    {14}    'sur encaissement',
    {15}    'Débit',
    {16}    'Encais.',
    {18}    'Débits',
    {19}    'Encaissements',
    {20}    'Débits et encaissements',
    {21}    '21'
    );

//==================================================
// fonctions hors class
//==================================================
{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 24/04/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure EditionTvaHT;
begin
    AGLLanceFiche('CP','CPQRTVACTRL','','','');
end;

//==================================================
// Evenements par default de la TOF
//==================================================
{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 24/04/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPQRTVACTRL.OnNew;
begin
    Inherited;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 24/04/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPQRTVACTRL.OnDelete;
begin
    Inherited;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 24/04/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPQRTVACTRL.OnUpdate;
begin
    Inherited;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 24/04/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPQRTVACTRL.OnLoad;
begin
    Inherited;
    // click sur le bouton BValider avant le chargement des critères

    // YMO 11/04/2006 FQ 17630 le combo "Fetat" détermine l'état utilisé,
    // et non 'NatureEtat' et 'CodeEtat' sur le QRS1  
    THValComboBox(GetControl('FETAT')).Value:='CTR';

    // set du where
    if (not Critere) then SetControlText('XX_WHERE',GetCritere)
    else SetControlText('XX_WHERE','');
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 24/04/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPQRTVACTRL.OnArgument(S : String);
var
    FMONTANT,FTYPETVA,FNATTVA : TRadioGroup;
    MONVALIDER : TToolBarButton97;
begin
    Critere := false;

    Inherited;

    // reassigne du FormKeyDown
    Ecran.OnKeyDown := FormKeyDown;

    // recup des control
    FGENERAL := THEdit(GetControl('FGENERAL',true)); if (not assigned(FGENERAL)) then exit;
    FGENERAL_ := THEdit(GetControl('FGENERAL_',true)); if (not assigned(FGENERAL_)) then exit;
    FNATURECPT := THValComboBox(GetControl('FNATURECPT',true)); if (not assigned(FNATURECPT)) then exit;
    FDEVISE := THValComboBox(GetControl('FDEVISE',true)); if (not assigned(FDEVISE)) then exit;
    FNATURECPT := THValComboBox(GetControl('FNATURECPT',true)); if (not assigned(FNATURECPT)) then exit;
    FMONTANT := TRadioGroup(GetControl('FMONTANT',true)); if (not assigned(FMONTANT)) then exit;
    FTYPETVA := TRadioGroup(GetControl('FTYPETVA',true)); if (not assigned(FTYPETVA)) then exit;
    FNATTVA := TRadioGroup(GetControl('FNATTVA',true)); if (not assigned(FNATTVA)) then exit;
    FEXERCICE := THValComboBox(GetControl('FEXERCICE',true)); if (not assigned(FEXERCICE)) then exit;

    MONVALIDER := TToolBarButton97(GetControl('MONVALIDER',true)); if (not (assigned(MONVALIDER))) then exit;

    // set des fcts
    FMONTANT.OnClick := OnClickFMONTANT;
    FTYPETVA.OnClick := OnClickFTYPETVA;
    FNATTVA.OnClick := OnClickFNATTVA;
    FTYPETVA.OnClick := OnClickFTYPETVA;
    FNATURECPT.OnChange := OnChangeFNATURECPT;
    FEXERCICE.OnChange := OnChangeFEXERCICE;
    FDEVISE.OnChange := OnChangeFDEVISE;
    MONVALIDER.OnClick := OnClickMONVALIDER;

    // set des valeur par default
    FEXERCICE.Value := VH^.EnCours.Code;
    FDEVISE.ItemIndex := 0;
    {JP 15/10/07 : FQ 16149 : gestion des réstrictions Etablissements}
    if GetControlText('FETABLISSEMENT') = '' then
      THValComboBox(GetControl('FETABLISSEMENT', True)).ItemIndex := 0;

    if (FNATURECPT.ItemIndex = -1) then
    begin
        FNATURECPT.ItemIndex := 0;
        OnChangeFNATURECPT(FNATURECPT);
    end;

    OnClickFTYPETVA(FTYPETVA);
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 24/04/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPQRTVACTRL.OnClose;
begin
    Inherited;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 24/04/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPQRTVACTRL.OnDisplay;
begin
    Inherited;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 24/04/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPQRTVACTRL.OnCancel;
begin
    Inherited;
end;

//==================================================
// Autres Evenements
//==================================================
{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 24/04/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPQRTVACTRL.OnChangeFEXERCICE(Sender : TObject);
begin
    ExoToDates(FEXERCICE.Value,GetControl('FDATE'),GetControl('FDATE_'));
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 24/04/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPQRTVACTRL.OnChangeFDEVISE(Sender : TObject);
begin
    SetControlText('FCDEVISE',THValComboBox(Sender).Value);
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 24/04/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPQRTVACTRL.OnChangeFNATURECPT(Sender : TObject);
var
    C1,C2 : String ;
begin
    C1 := FGENERAL.Text;
    C2 := FGENERAL_.Text;

    FGENERAL.Clear;
    FGENERAL_.Clear;

    Case FNATURECPT.ItemIndex of
        0 :
            begin
                FGENERAL.DataType := 'TZGCHARGE';
                FGENERAL_.DataType := 'TZGCHARGE';
                Ecran.Caption := Mess[2];
                SetControlTExt('FTYPEEDT',Mess[2]);
            end;
        1 :
            begin
                FGENERAL.DataType := 'TZGIMMO';
                FGENERAL_.DataType := 'TZGIMMO';
                Ecran.Caption := Mess[1];
                SetControlTExt('FTYPEEDT',Mess[1]);
            end;
        2 :
            begin
                FGENERAL.DataType := 'TZGPRODUIT';
                FGENERAL_.DataType := 'TZGPRODUIT';
                Ecran.Caption := Mess[0];
                SetControlTExt('FTYPEEDT',Mess[0]);
            end;
    end;

    FGENERAL.Text := C1;
    FGENERAL_.Text := C2;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 24/04/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPQRTVACTRL.OnClickFNATTVA(Sender : TObject);
var
    index : integer;
begin
    // recup de l'index
    index := THRadioGroup(Sender).ItemIndex;
    // set d'une zone de text "pour recuperé le bon taux de TVA !"
    SetControlText('FNNATTVA',IntToStr(index));
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 24/04/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPQRTVACTRL.OnClickFTYPETVA(Sender : TObject);
var
    index : integer;
begin
    // recup de l'index
    index := THRadioGroup(Sender).ItemIndex;
    // set d'une zone de text "pour recuperé le bon taux de TVA !"
    SetControlText('FNTYPETVA',IntToStr(index));

    if (index = 0) then SetControlText('FNATURECPT','PRO')
    else SetControlText('FNATURECPT','CHA');
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 24/04/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPQRTVACTRL.OnClickFMONTANT(Sender : TObject);
var
    index : integer;
begin
    // recup de l'index
    index := THRadioGroup(Sender).ItemIndex;

    // test sur ledit index et set de la zone de text "pour selectionné les bon debit !"
    case (index) of
        0 : SetControlText('FTDEV','');
        1 :
            begin
                if (FDEVISE.Value = '') then
                begin
                    THRadioGroup(Sender).ItemIndex := 0;
                    index := 0;
                    HShowMessage(Mess[3],'','');
                end
                else SetControlText('FTDEV','DEV');
            end;
        2 : SetControlText('FTDEV','EURO');
        else SetControlText('FTDEV','');
    end;

    // set d'une autre zone de text "Pour affiche les bonnes devise"
    SetControlText('FNMONTANT',IntToStr(index));
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 28/04/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPQRTVACTRL.OnClickMONVALIDER(Sender : TObject);
var
    tmp : string;
    BVALIDER : TToolBarButton97;
begin
    // disable du btn valider
    SetControlEnabled('MONVALIDER',false);

    // get du btn valider
    BVALIDER := TToolBArButton97(GetControl('BVALIDER',true));

    // si on fait les Critere
    if (GetCheckBoxState('FRAPPELCRIT') = cbChecked) then
    begin
        // on sauvegarde l'etat en cour et on set les critere
        Critere := true;
        tmp := TFQRS1(Ecran).CodeEtat;
        // on passe sur l'etat critere
        TFQRS1(Ecran).CodeEtat := 'CR2';
        // on imprime
        BVALIDER.OnClick(BVALIDER);
        // on remet le bon etat et critere
        TFQRS1(Ecran).CodeEtat := tmp;
        Critere := false;
    end
    else
    begin
        BVALIDER.OnClick(BVALIDER);
    end;

    // re-enable du btn valider
    SetControlEnabled('MONVALIDER',true);
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 28/04/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPQRTVACTRL.FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
begin
    Case Key of
        VK_F9 :
            begin
                if Shift=[] then
                begin
                    Key := 0;
                    OnClickMONVALIDER(Nil);
                end;
            end;
    end;

    inherited;
end;

//==================================================
// Autres fonctions de la class
//==================================================
{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 28/04/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
function TOF_CPQRTVACTRL.GetCritere : string;
var
    Query : string;
    NatCpt : string;
    cpt1,cpt2 : string;
    exo,date1,date2 : string;
    exep : string;
    encaissement : integer;
    etablissement : string;
    devise : string;
begin
    Query := '';

    // recuperation des critere
    NatCpt := GetControlText('FNATURECPT');
    cpt1 := GetControlText('FGENERAL');
    cpt2 := GetControlText('FGENERAL_');
    exo := GetControlText('FEXERCICE');
    date1 := UsDateTime(StrToDate(GetControlText('FDATE')));
    date2 := UsDateTime(StrToDate(GetControlText('FDATE_')));
    exep := GetControlText('FEXCEP');
    encaissement := StrToInt(GetControlText('FNNATTVA'));
    etablissement := GetControlText('FETABLISSEMENT');
    devise := GetControlText('FDEVISE');

    // where de la requete
    Query := Query + 'E_EXERCICE="' + exo + '" AND E_DATECOMPTABLE>="' + date1 + '" AND E_DATECOMPTABLE<="' + date2 + '" AND ';
    if (not (etablissement = '')) then Query := Query + 'E_ETABLISSEMENT="' + etablissement + '" AND ';
    if (not (devise = '')) then Query := Query + 'E_DEVISE="' + devise + '")] AND ';
    Case encaissement Of
        0 : Query := Query + 'E_TVAENCAISSEMENT="-" AND ';
        1 : Query := Query + 'E_TVAENCAISSEMENT="X" AND ';
        2 : Query := Query + '(E_TVAENCAISSEMENT="X" OR E_TVAENCAISSEMENT="-") AND ';
    end;

    // creation de la sous-requete !
    Query := Query + 'E_GENERAL IN (SELECT G_GENERAL FROM GENERAUX WHERE ';

    // Nature de compte
    if (NatCpt = '') then Query := Query + '(G_NATUREGENE="CHA" OR G_NATUREGENE="PRO") AND '
    else
    begin   {FQ18449  18.12.2007  OD d'Escomptes du côté charges}
      Query := Query + '(G_NATUREGENE="' + NatCpt + '" ';

      If NatCpt='CHA' then Query := Query + 'OR ' else Query := Query + 'AND NOT ';

      Query := Query + '(E_JOURNAL="'+GetParamSocSecur('SO_CPJALESCOMPTE','')+'" AND E_QUALIFORIGINE="ESC")) AND ';
    end;
    // existance de TVA
    Query := Query + 'G_TVA<>"" AND ';
    // existance d'ecriture
    Query := Query + '((G_TOTDEBE<>0 OR G_TOTCREE<>0 OR G_TOTDEBANO<>0 OR G_TOTCREANO<>0) OR (EXISTS (SELECT E_GENERAL FROM ECRITURE WHERE E_GENERAL=G_GENERAL AND E_EXERCICE="' + exo + '" AND E_DATECOMPTABLE>="' + date1 + '" AND E_DATECOMPTABLE<="' + date2 + '" AND E_QUALIFPIECE="N"))) ';
    // encaisement ??
    Case encaissement Of
        0 : Query := Query + 'AND (EXISTS(SELECT E_AUXILIAIRE FROM ECRITURE WHERE E_GENERAL=G_GENERAL AND E_TVAENCAISSEMENT="-" AND E_EDITEETATTVA="-"))';
        1 : Query := Query + 'AND (EXISTS(SELECT E_AUXILIAIRE FROM ECRITURE WHERE E_GENERAL=G_GENERAL AND E_TVAENCAISSEMENT="X" AND E_EDITEETATTVA="-"))';
        2 : Query := Query + 'AND (EXISTS(SELECT E_AUXILIAIRE FROM ECRITURE WHERE E_GENERAL=G_GENERAL AND (E_TVAENCAISSEMENT="X" OR E_TVAENCAISSEMENT="-") AND E_EDITEETATTVA="-"))';
    end;
    // comptes ??
    if (cpt1 <> '') then Query := Query + 'AND G_GENERAL>="' + cpt1 + '" ';
    if (cpt2 <> '') then Query := Query + 'AND G_GENERAL<="' + cpt2 + '" ';
    // compte d'exeption
    if (exep <> '') then Query := Query + 'AND '+ AnalyseCompte(exep,fbGene,true,false);

//    Query := Query + ' ORDER BY G_GENERAL)';
    Query := Query + ' )';

    result := Query;
end;

//================================================================================
// Initialization
//================================================================================
Initialization
    registerclasses([TOF_CPQRTVACTRL]);
end.
