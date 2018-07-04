{***********UNITE*************************************************
Auteur  ...... : BPY
Créé le ...... : 07/04/2003
Modifié le ... : 27/02/2004
Description .. : Source TOF de la FICHE : CPQRTVAFAC ()
Suite ........ : 
Suite ........ : 06/06/2003 : Ajout d'une vue pour éviter les problème de
Suite ........ : jointures multiples sur la table ECRITURE. Utilisation des
Suite ........ : alias de la vue lors de la construction des critères
Suite ........ : 
Suite ........ : A FAIRE : ajouter test sur nombre de pages pour lancement
Suite ........ : auto des états suivant si aucune ligne à imprimer. Car
Suite ........ : actuellement reste bloquer sur bouton "suivant".
Suite ........ :
Suite ........ : BPY le 27/02/2004 : test ajouté ..... modification du 
Suite ........ : traitement des EAC pour collé au deux autre editions !
Mots clefs ... : TOF;CPQRTVAFAC
*****************************************************************}

Unit CPQRTVAFAC_TOF;

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
    HMsgBox,
    UTOF,
    UtilPGI,
    UTOB,
    CPMSGTVA_TOF,
    HStatus,
    HTB97,
    Windows,
    HPanel,
    UTILEDT,
    TofMeth
    ;

//==================================================
// Externe
//==================================================
procedure EditionTvaFac(D : boolean);
procedure EditionTvaAcE;

//==================================================
// Definition de class
//==================================================
Type
    TOF_CPQRTVAFAC = Class(TOF_Meth)
    public
        procedure OnNew                  ; override ;
        procedure OnDelete               ; override ;
        procedure OnUpdate               ; override ;
        procedure OnLoad                 ; override ;
        procedure OnArgument(S : String) ; override ;
        procedure OnDisplay              ; override ;
        procedure Onclose                ; override ;
        procedure OnCancel               ; override ;
    private
        procedure OnChangeFTVA(Sender : TObject);
        procedure OnChangeFEXERCICE(Sender : TObject);
        procedure OnChangeFDEVISE(Sender : TObject);

        procedure OnClickFMONTANT(Sender : TObject);
        procedure OnClickFTYPETVA(Sender : TObject);
        procedure OnClickMONVALIDER(Sender : TObject);
        procedure bSuivantClick(Sender : TObject);

        procedure FormKeyDown(Sender : TObject ; var Key : Word ; Shift : TShiftState);

        procedure OnExitFDATE(Sender : TObject);
        procedure OnExitFDATE_(Sender : TObject);

        procedure FinirPrint;
        procedure MajOkEditionFAC;
        procedure MajOkEditionEAC;
        function  GetCritereFAC (remplace : boolean) : string;
        function  GetCritereEAC : string;
        procedure FUniqAcompteClick (Sender : TObject);

    public
    private
        FUniqAcompte : THCheckBox ;
        FTVA,FDEVISE,FEXERCICE,FREGIMETVA : THValComboBox;
        DDeb,DFin : TDateTime;
        What : String3;
        TEdition : Integer;
        Critere : boolean;
        bExit, bContinue : boolean;
    end;

//================================================================================
// Implementation
//================================================================================
Implementation

uses
  {$IFDEF MODENT1}
  ULibExercice,
  {$ENDIF MODENT1}
  Paramsoc;

//==================================================
// Variable
//==================================================
var
    MsgBox : array [0..13] of string = (
    {00}    'TVA sur encaissements : factures directes',
    {01}    'TVA sur encaissements : factures par étapes',
    {02}    'Etat des acomptes',
    {03}    'Date d''échéance du',
    {04}    'Date comptables du',
    {05}    'Mise à jour impossible! Des écritures ont été modifiés par un autre utilisateur !',
    {06}    '6;?caption?;Vous devez d''abord sélectionner une devise particulière.;E;O;O;O;',
    {07}    '7;?caption?;Le paramétrage de la TVA est incorrect !;E;O;O;O;',
    {JP 02/06/05 : Suite à une remarque de MLL, un seul message suffit, dans la mesure où 8 et 9 sont
                   appelés consécutivement}
//    {08}    '8;?caption?;Le module n''est pas installé !;E;O;O;O;',
//    {09}    '9;?caption?;Vous pouvez paramétrer ce module dans les paramètres société;E;O;O;O;',
    {08}    '',
    {09}    '9;?caption?;Le module n''est pas installé.'#13#13 +
                        'Vous pouvez activer ce module dans les paramètres société.;E;O;O;O;',
    {10}    '10;?caption?;La fourchette de dates n''appartient pas à un exercice.;E;O;O;O;',
    {11}    '11;?caption?;Dates de début et de fin incohérentes.;E;O;O;O;',
    {12}    'Vous n''avez pas terminé l''impression de tous les documents.',
    {13}    'Edition des règlements non lettrés'
    );

//==================================================
// fonctions hors class
//==================================================
{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 08/04/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure EditionTvaFac(D : boolean);
begin
  {JP 02/06/05 : Même problème que pour la FQ 15355, ce test est effectué trop tard (OnArgument) :
                 je reprends le code de la fiche QRTvaFac.Pas}
  if not ExisteSQL('SELECT CC_CODE,CC_LIBELLE FROM CHOIXCOD WHERE CC_TYPE = "' +
                   VH^.DefCatTVA + '" AND (CC_LIBRE = "1" OR CC_LIBRE = "2"' +
                   ' OR CC_LIBRE = "3" OR CC_LIBRE = "4") ORDER BY CC_LIBRE') then begin
    HShowMessage(MsgBox[7], '', '');
    Exit;
  end;

  {JP 02/06/05 : FQ 15355 : Le test sur le module de TVA est remonté de OnArgument à EditionTvaFac}
  if (not VH^.OuiTvaEnc) then
    {JP 02/06/05 : Suite à une remarque de MLL, un seul message suffit}
    HShowMessage(MsgBox[9], '', '')
  else if (D) then
    AGLLanceFiche('CP', 'CPQRTVAFAC', '', '', 'ACTION=MODIFICATION;FAC')
  else
    AGLLanceFiche('CP', 'CPQRTVAFAC', '', '', 'ACTION=MODIFICATION;FET');
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 14/04/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure EditionTvaAcE;
begin
    AGLLanceFiche('CP','CPQRTVAFAC','','','ACTION=MODIFICATION;EAC');
end;

//==================================================
// Evenements par default de la TOF
//==================================================
{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 07/04/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPQRTVAFAC.OnNew;
begin
    Inherited;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 07/04/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPQRTVAFAC.OnDelete;
begin
    Inherited;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 07/04/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPQRTVAFAC.OnUpdate;
begin
    Inherited;

    // Après départ de l'impression, le bouton pour passer aux éditions suivantes est rendu accessible
    SetControlEnabled('BSuivant',true);
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 07/04/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPQRTVAFAC.OnLoad;
begin
    Inherited;

    // set du where
    if (not Critere) then
    begin
        if TEdition = 2 then SetControlText('XX_WHERE',GetCritereEAC)
        else SetControlText('XX_WHERE',GetCritereFAC (true));
    end
    else SetControlText('XX_WHERE','');
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 07/04/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPQRTVAFAC.OnArgument(S : String);
var
    MONVALIDER,BVALIDER,BEURO,BSUIVANT : TToolBarButton97;
    FMONTANT,FTYPETVA : THRadioGroup;
    FDATE,FDATE_ : THEdit;
    FETABLISSEMENT : THValComboBox;
    S1,S2 : string;
    NumTitre : ShortInt;
begin
    Critere := false;
    bContinue := True;

    Inherited;

    // recup de l'edition a faire dselon le parametre !
    S1 := UpperCase(S);
    S2 := ReadTokenSt(S1);                              // mode AGL
    S2 := ReadTokenSt(S1); What := S2;                  // Type d'edition

    // reassigne du FormKeyDown
    Ecran.OnKeyDown := FormKeyDown;

    // recup du combo d'exercice
    FEXERCICE := THValComboBox(GetControl('FEXERCICE',true)); if (not assigned(FEXERCICE)) then exit;
    FEXERCICE.OnChange := OnChangeFEXERCICE;

    // recup du combo de type de TVA
    FTVA := THValComboBox(GetControl('FTVA',true)); if (not assigned(FTVA)) then exit;
    FTVA.OnChange := OnChangeFTVA;
    FTVA.ItemIndex := 0;

    // recup du combo etablissement
    FETABLISSEMENT := THValComboBox(GetControl('FETABLISSEMENT',true)); if (not assigned(FETABLISSEMENT)) then exit;
   // FETABLISSEMENT.ItemIndex := 0;

    // recup du combo devise
    FDEVISE := THValComboBox(GetControl('FDEVISE',true)); if (not assigned(FDEVISE)) then exit;
    FDEVISE.OnChange := OnChangeFDEVISE;
    FDEVISE.ItemIndex := 0;

    // recup du combo Regime TVA
    FREGIMETVA := THValComboBox(GetControl('FREGIMETVA',true)); if (not assigned(FREGIMETVA)) then exit;
    FREGIMETVA.ItemIndex := 0;

    FUNIQACOMPTE := THCheckBox(GetControl('FUNIQACOMPTE',true)); if (not assigned(FUNIQACOMPTE)) then exit;

    FUNIQACOMPTE.OnClick :=FUniqAcompteClick ;

    // recup des deux THRadioGroup
    FMONTANT := THRadioGroup(GetControl('FMONTANT',true)); if (not assigned(FMONTANT)) then exit;
    FMONTANT.OnClick := OnClickFMONTANT;
    FTYPETVA := THRadioGroup(GetControl('FTYPETVA',true)); if (not assigned(FTYPETVA)) then exit;
    FTYPETVA.OnClick := OnClickFTYPETVA;

    // recup des deux date
    FDATE := THEdit(GetControl('FDATE',true)); if (not assigned(FDATE)) then exit;
    FDATE.OnExit := OnExitFDATE;
    FDATE_ := THEdit(GetControl('FDATE_',true)); if (not assigned(FDATE_)) then exit;
    FDATE_.OnExit := OnExitFDATE_;

    // bouton Valider
    BVALIDER := TToolBarButton97(GetControl('BVALIDER',true)); if (not assigned(BVALIDER)) then exit;
    MONVALIDER := TToolBarButton97(GetControl('MONVALIDER',true)); if (not assigned(MONVALIDER)) then exit;
    MONVALIDER.OnClick := OnClickMONVALIDER;
    MONVALIDER.Top := BVALIDER.Top;
    MONVALIDER.Left := BVALIDER.Left;

    // bouton suivant
    BEURO := TToolBarButton97(GetControl('BEURO',true));
    BSUIVANT := TToolBarButton97(GetControl('BSUIVANT',true));
    BSUIVANT.OnClick := bSuivantClick;
    BSUIVANT.Top := BEURO.Top;
    BSUIVANT.Left := BEURO.Left;
    BSUIVANT.Enabled := False ;

    {!! JP 02/06/05 : FQ 15355 : les deux tests sur le paramétrage de la TVA on été remontés dans EditionTvaFac}

    // modif des parametre selon le type
    if (What = 'FAC') then                  // Gestion directe des factures
    begin
        TEdition := 1;

        // modif de l'ecran
        if (IsInside(Ecran)) then THpanel(TFQRS1(Ecran).parent).insideTitle.Caption := MsgBox[0];
        TFQRS1(Ecran).Caption := MsgBox[0];
        TFQRS1(Ecran).HelpContext := 7650000;
        TFQRS1(Ecran).CodeEtat := 'FAC';

        // pour l'etat
        SetControlText('FTYPEFAC',MsgBox[0]);

        // changement des label
        SetControlText('TFDATE','&' + MsgBox[4]);
        SetControlText('FTDATE',MsgBox[4]);

        // champs a ajouté effacé !
        SetControlVisible('FDETAIL',true);
        SetControlVisible('FREPORT',false);
        SetControlVisible('FLIGNEREGIME',true);
        // proprieté a modifié !
        FEXERCICE.Vide := false;
        FEXERCICE.ReLoad;
        FEXERCICE.Value := VH^.EnCours.Code;
    end
    else if (What = 'FET') then             // Factures par étapes
    begin
        TEdition := 0;

        // modif de l'ecran
        if (IsInside(Ecran)) then THpanel(TFQRS1(Ecran).parent).insideTitle.Caption := MsgBox[1];
        TFQRS1(Ecran).Caption := MsgBox[1];
        TFQRS1(Ecran).HelpContext := 7651050;
        TFQRS1(Ecran).CodeEtat := 'FAC';

        // pour l'etat
        SetControlText('FTYPEFAC',MsgBox[1]);

        // changement des label
        SetControlText('TFDATE','&' + MsgBox[3]);
        SetControlText('FTDATE',MsgBox[3]);

        // champs a ajouté effacé !
        SetControlVisible('FDETAIL',true);
        SetControlVisible('FREPORT',false);
        SetControlVisible('FLIGNEREGIME',true);
        // proprieté a modifié !
        FEXERCICE.Vide := true;
        FEXERCICE.ReLoad;
        FEXERCICE.Value := VH^.EnCours.Code;
    end
    else if (What = 'EAC') then             // Etat des acomptes
    begin
        TEdition := 2;

        FUniqAcompte.Visible:=True;

        If FUniqAcompte.Checked then NumTitre:=2 else NumTitre:=13;

        // modif de l'ecran
        if (IsInside(Ecran)) then THpanel(TFQRS1(Ecran).parent).insideTitle.Caption := MsgBox[NumTitre];
        TFQRS1(Ecran).Caption := MsgBox[NumTitre];
        TFQRS1(Ecran).HelpContext := 7652000;
        TFQRS1(Ecran).CodeEtat := 'EAC';

        // pour l'etat
        SetControlText('FTYPEFAC',MsgBox[NumTitre]);

        // changement des label
        SetControlText('TFDATE','&' + MsgBox[4]);
        SetControlText('FTDATE',MsgBox[4]);

        // champs a ajouté effacé !
        SetControlVisible('FDETAIL',false);
        SetControlVisible('FREPORT',true);
        SetControlVisible('FLIGNEREGIME',false);
        // proprieté a modifié !
        FEXERCICE.Vide := false;
        FEXERCICE.ReLoad;
        FEXERCICE.Value := VH^.EnCours.Code;
    end
    else close;

    TFQRS1(Ecran).ChoixEtat  := True;
    TFQRS1(Ecran).ParamEtat  := True;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 07/04/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPQRTVAFAC.Onclose;
begin
    bExit := true;
    // Empêche de sortir si on n'a pas terminé l'impression de TOUS les documents
    if (not bContinue) then
    begin
        PGIBox(MsgBox[12],Ecran.Caption);
        LastError := -1;
        bExit := false;
    end;

    Inherited;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 07/04/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPQRTVAFAC.OnDisplay;
begin
    Inherited;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 07/04/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPQRTVAFAC.OnCancel;
begin
    Inherited;
end;

//==================================================
// Autres Evenements
//==================================================
{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 08/04/2003
Modifié le ... : 12/12/2007
Description .. :
Mots clefs ... : FQ21842
*****************************************************************}
procedure TOF_CPQRTVAFAC.OnChangeFTVA(Sender : TObject);
var
    Q : TQuery;
    req, FDossier : String;
begin
  {FQ21842 12.12.2007 YMO CPTVA Tablette partagée pour les index de E_ECHEENCX}
  If TableToBaseNum('CPTVA','TTE') >0 then
      FDossier := TabletteToBase('CPTVA');

  req:='SELECT CC_LIBRE FROM CHOIXCOD WHERE CC_TYPE="TX1" AND CC_CODE="' + THValComboBox(Sender).Value + '"';

  Q := OpenSelect( req, FDossier );

  if (not Q.EOF) then SetControlText('FNTVA',Q.Fields[0].AsString)
  else SetControlText('FNTVA','1');

  Ferme(Q);
  FreeAndNil(Q);
                    
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 08/04/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPQRTVAFAC.OnChangeFEXERCICE(Sender : TObject);
begin
    if (FEXERCICE.ItemIndex > 0) then
    begin
        ExoToDates(FEXERCICE.Value,GetControl('FDATE'),GetControl('FDATE_'));
        DDeb := StrToDate(GetControlText('FDATE'));
        DFin := StrToDate(GetControlText('FDATE_'));
    end
    else
    begin
        SetControlText('FDATE',StDate1900);
        SetControlText('FDATE_',StDate2099);
        DDeb := iDate1900;
        DFin := iDate2099;
    end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 09/04/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPQRTVAFAC.OnChangeFDEVISE(Sender : TObject);
begin
    SetControlText('FCDEVISE',THValComboBox(Sender).Value);
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 11/04/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPQRTVAFAC.OnClickFMONTANT(Sender : TObject);
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
                    SetControlText('FTDEV','');
                    HShowMessage(MsgBox[6],'','');
                end
                else SetControlText('FTDEV','DEV');
            end;
        else SetControlText('FTDEV','');
    end;

    // set d'une autre zone de text "Pour affiche les bonnes devise"
    SetControlText('FNMONTANT',IntToStr(index));
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 11/04/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPQRTVAFAC.OnClickFTYPETVA(Sender : TObject);
var
    index : integer;
begin
    // recup de l'index
    index := THRadioGroup(Sender).ItemIndex;

    // test sur ledit index et set de la zone de text "pour Critere de selection sur les natures de piece"
    case (index) of
        0 : SetControlText('FTNATPIECE','C');
        1 : SetControlText('FTNATPIECE','F');
        else SetControlText('FTNATPIECE','C');
    end;

    // set d'une autre zone de text "pour recuperé le bon taux de TVA !"
    SetControlText('FNTYPETVA',IntToStr(index));
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 11/04/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPQRTVAFAC.OnExitFDATE(Sender : TObject);
var
    d : TDateTime;
begin
    d := StrToDate(THEdit(Sender).Text);

    // date inferieur au debut d'exo
    if (d < DDeb) then
    begin
        HShowMessage(MsgBox[10],'','');
        THEdit(Sender).SetFocus;
    end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 11/04/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPQRTVAFAC.OnExitFDATE_(Sender : TObject);
var
    d1 : TDateTime;
    d2 : TDateTime;
begin
    d1 := StrToDate(GetControlText('FDATE'));
    d2 := StrToDate(THEdit(Sender).Text);

    // date supperieur a la fin d'exo
    if (d2 > DFin) then
    begin
        HShowMessage(MsgBox[10],'','');
        THEdit(Sender).SetFocus;
    end;
    // date inferieur a d1
    if (d1 > d2) then
    begin
        HShowMessage(MsgBox[11],'','');
        THEdit(Sender).SetFocus;
    end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 10/04/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPQRTVAFAC.OnClickMONVALIDER(Sender : TObject);
var
    tmp : string;
    i : integer;
    BVALIDER : TToolBArButton97;
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
        TFQRS1(Ecran).CodeEtat := 'CR1';
        // on imprime
        BVALIDER.OnClick(BVALIDER);
        // on remet le bon etat et critere
        TFQRS1(Ecran).CodeEtat := tmp;
        Critere := false;
    end
    else
    begin
        if (FTva.ItemIndex = 0) then
        begin // soit pas d'item soit '<<Tous>>' selectionné ... on execute pour tout les type
            bExit := False; // Empêche de sortir de la fiche
            for i := 1 to FTva.Items.Count - 1 do
            begin
                // changé l'index de TVA
                FTva.ItemIndex := i;
                OnChangeFTva(FTva);
                // valider pour de vrai
                BVALIDER.OnClick(BVALIDER);

                // apres impression
                FinirPrint;
            end;

            FTva.ItemIndex := 0;
        end
        else
        begin // un item autre que '<<Tous>>' selectionné ... on fait rien
            // valider pour de vrai
            BVALIDER.OnClick(BVALIDER);
            // apres impression
            FinirPrint;
        end;
    end;

    // re-enable du btn valider
    SetControlEnabled('MONVALIDER',true);
    // Re-disable du bouton suivant
    SetControlEnabled('BSUIVANT',false);

    // Permet de sortir de la fiche
//    bExit := True;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 11/04/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPQRTVAFAC.FormKeyDown(Sender: TObject; var Key: Word; Shift: TShiftState);
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

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 27/02/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPQRTVAFAC.bSuivantClick(Sender : TObject);
begin
    if (not bExit) then
        bContinue := true; // Passe au document suivant
end;

//==================================================
// Autres fonctions de la class
//==================================================
{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 10/04/2003
Modifié le ... : 10/06/2003
Description .. :
Suite ........ : 10/06/2003 : Distinction du traitement suivfant les états (a
Suite ........ : cause de la vue CPQRTVAFAC )
Mots clefs ... :
*****************************************************************}
procedure TOF_CPQRTVAFAC.FinirPrint;
var
    Query : string;
begin
    if TEdition = 2  then
    begin
        // test si des pages on ete imprimé ....
        Query := 'SELECT E_EXERCICE FROM ECRITURE E1 LEFT JOIN TIERS ON E_AUXILIAIRE=T_AUXILIAIRE LEFT JOIN TXCPTTVA ON TV_REGIME=E_REGIMETVA WHERE  TV_CODETAUX="' + GetControlText('FTVA') + '" ' + GetCritereEAC;
        if (ExisteSQL(Query)) then
        begin
            // Permet d'attendre que l'utilisateur click sur le bouton fermer pour passer à la pré-validation
            bContinue := False;
            while (true) do
            begin
                if (bContinue) then break;
                Application.ProcessMessages;
            end;

            // validation
            if (MsgTva(TEdition,(GetControlText('FNTYPETVA') = '1'),StrToDate(GetControlText('FDATE')),StrToDate(GetControlText('FDATE_')),GetControlText('FREGIMETVA'),GetControlText('FTVA'))) then
            begin
                if (Transactions(MajOkEditionEAC,1) <> OeOk) then MessageAlerte(MsgBox[5]);
            end;
        end;
    end
    else
    begin {27/11/2007 YMO Requête remplace la vue pour éviter modif SOCREF (ajout jointure sur JOURNAL)}
        Query := 'SELECT E1.E_EXERCICE, E1.E_DATECOMPTABLE, E1.E_DATEECHEANCE,'
        +' E2.E_NATUREPIECE, E2.E_ETABLISSEMENT, E2.E_ECHE, E2.E_AUXILIAIRE,E2.E_TYPEMVT,E1.E_GENERAL,'
        +' E1.E_NATUREPIECE,E1.E_EDITEETATTVA,E1.E_LETTRAGE'
        +' FROM ECRITURE E1 LEFT JOIN ECRITURE E2 ON E2.E_AUXILIAIRE=E1.E_AUXILIAIRE'
        +' AND E2.E_GENERAL=E1.E_GENERAL AND E2.E_ETATLETTRAGE=E1.E_ETATLETTRAGE AND E2.E_LETTRAGE=E1.E_LETTRAGE'
        +' LEFT JOIN GENERAUX ON E1.E_CONTREPARTIEGEN=G_GENERAL LEFT JOIN JOURNAL J1 ON E1.E_JOURNAL=J1.J_JOURNAL'
        +' LEFT JOIN JOURNAL J2 ON E2.E_JOURNAL=J2.J_JOURNAL LEFT JOIN TIERS ON E1.E_AUXILIAIRE=T_AUXILIAIRE'
        +' LEFT JOIN TXCPTTVA ON E2.E_REGIMETVA=TV_REGIME WHERE ' + GetCritereFAC (True);
        // test si des pages on ete imprimé ....
//        Query := 'SELECT EXERCICE1 FROM CPQRTVAFAC WHERE ' + GetCritereFAC (False);
        if (ExisteSQL(Query)) then
        begin
            // Permet d'attendre que l'utilisateur click sur le bouton fermer pour passer à la pré-validation
            bContinue := False;
            while (true) do
            begin
                if (bContinue) then break;
                Application.ProcessMessages;
            end;

            // validation
            if (MsgTva(TEdition,(GetControlText('FNTYPETVA') = '1'),StrToDate(GetControlText('FDATE')),StrToDate(GetControlText('FDATE_')),GetControlText('FREGIMETVA'),GetControlText('FTVA'))) then
            begin
                if (Transactions(MajOkEditionFAC,1) <> OeOk) then MessageAlerte(MsgBox[5]);
            end;
        end;
    end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 10/04/2003
Modifié le ... : 10/06/2003
Description .. :
Suite ........ : 06/06/2003 : Ajout d'une vue pour éviter les problème de
Suite ........ : jointures multiples sur la table ECRITURE. Utilisation des
Suite ........ : alias de la vue lors de la construction des critères
Suite ........ :
Suite ........ : YMO 28/06/2006 Modif pour rajouter les proratas ;
Suite ........ : pour ne pas modifier la vue, intégration de la vue dans la requête de l'état,
Suite ........ : et gestion dans le source des alias dans la clause WHERE
Mots clefs ... :
*****************************************************************}
Const
  ArrEcrit: array[0..30] of string =
  ( 'E1.E_JOURNAL','E1.E_NATUREPIECE','E1.E_DATECOMPTABLE','E1.E_DATEECHEANCE','E1.E_EXERCICE','E1.E_NATUREPIECE',
    'E1.E_EDITEETATTVA','E1.E_AUXILIAIRE','E1.E_GENERAL','E1.E_ETATLETTRAGE','E1.E_LETTRAGE','E1.E_NUMEROPIECE',
    'E1.E_NUMLIGNE','E2.E_JOURNAL','E2.E_NUMEROPIECE','E2.E_NUMLIGNE','E2.E_DATECOMPTABLE','E2.E_EXERCICE',
    ' E2.E_ECHEENC1',' E2.E_ECHEENC2',' E2.E_ECHEENC3',' E2.E_ECHEENC4','E2.E_NATUREPIECE','E2.E_ETABLISSEMENT',
    'E2.E_REGIMETVA','E2.E_GENERAL','E2.E_AUXILIAIRE','E2.E_ECHE','E2.E_TYPEMVT','J1.J_EFFET','J2.J_EFFET');

  ArrAlias: array[0..30] of string =
    ('JOURNAL1','NATUREPIECE1','DATECOMPTABLE1','DATEECHEANCE1','EXERCICE1','NATUREPIECE1','EDITEETATTVA1',
    'AUXILIAIRE1','GENERAL1','ETATLETTRAGE1','LETTRAGE1','NUMEROPIECE1','NUMLIGNE1','JOURNAL2','NUMEROPIECE2',
    'NUMLIGNE2','DATECOMPTABLE2','EXERCICE2',' ECHEENC1',' ECHEENC2',' ECHEENC3',' ECHEENC4','NATUREPIECE2',
    'ETABLISSEMENT2','REGIMETVA2','GENERAL2','AUXILIAIRE2','ECHE2','TYPEMVT2','JEFFET1','JEFFET2');

function TOF_CPQRTVAFAC.GetCritereFAC (remplace : boolean) : string;
var
    where, stTemp : string;
    StValid,etable,regime : String;
    Deductible : boolean;
    FiltreCompenTP: String;         {FP 19/05/2006 FQ17981}
    i :   integer;
begin
    where := '';

    // Recuperation des critere pour le XX_WHERE
    Deductible := (GetControlText('FNTYPETVA') = '1');
    etable := GetControlText('FETABLISSEMENT');
    regime := GetControlText('FREGIMETVA');

  {b FP 19/05/2006 FQ17981}
  FiltreCompenTP := '';{ YMO 28/06/2006 Mise en stand by des modifs pour fonctionnement tiers-payeurs
  if Deductible and (VH^.JalATP <> '') then          {Tiers payeur fournisseur
    FiltreCompenTP := FiltreCompenTP +
    ' OR (JOURNAL1="'+VH^.JalATP+'" AND NATUREPIECE1="FF" AND (G_NATUREGENE="COF" OR G_NATUREGENE="COD" OR G_NATUREGENE="COS") AND T_PAYEUR<>"")';
  if (not Deductible) and (VH^.JalVTP <> '') then    {Tiers payeur client
    FiltreCompenTP := FiltreCompenTP +
    ' OR (JOURNAL1="'+VH^.JalVTP+'" AND NATUREPIECE1="FC" AND (G_NATUREGENE="COC" OR G_NATUREGENE="COD" OR G_NATUREGENE="COS") AND T_PAYEUR<>"")';
  if TCompensation.IsCompensation then
    FiltreCompenTP := FiltreCompenTP +
    ' OR (JOURNAL1="'+TCompensation.GetJalCompensation+'" AND NATUREPIECE1="OD" AND (G_NATUREGENE="COF" OR G_NATUREGENE="COD" OR G_NATUREGENE="COS" OR G_NATUREGENE="COC"))';
  e FP 19/05/2006}
    // Creation de la clause WHERE

    // Factures directes et accompte : date comptable // Factures indirectes : date échéance
    if (TEdition > 0) then where := where + ' DATECOMPTABLE1>="' + usDateTime(StrToDate(GetControlText('FDATE'))) + '" AND DATECOMPTABLE1<="' + usdatetime(StrToDate(GetControlText('FDATE_'))) + '"'
    else where := where + ' DATEECHEANCE1>="' + usDateTime(StrToDate(GetControlText('FDATE'))) + '" AND DATEECHEANCE1<="' + usdatetime(StrToDate(GetControlText('FDATE_'))) + '"';

    where := where + ' AND ECHEENC' + GetControlText('FNTVA') + '<>0';
    where := where + ' AND TV_CODETAUX="' + GetControlText('FTVA') + '"';
// GP N° 23485 le 01/09/2008 Ca plante et ca me parait normal : et c'est du BO !!!!!
//  if remplace then
    where := where + ' AND (NATUREPIECE2="F' + GetControlText('FTNATPIECE')
    + '" OR NATUREPIECE2="A' + GetControlText('FTNATPIECE');

    if (TEdition=1) then
       where := where + '" OR (NATUREPIECE2="OD" AND J2.J_EFFET<>"X"))' {27/11/2007 YMO prendre les OD qui ne sont pas d'effets}
    else
       where := where + '" OR (NATUREPIECE2="OD" AND J2.J_EFFET="X"))';  {27/11/2007 YMO prendre uniquement les OD d'effets}


    if (not (etable = '')) then where := where + ' AND ETABLISSEMENT2="' + etable + '"';
    if (not (regime = '')) then where := where + ' AND REGIMETVA2="' + regime + '"';


    // si deductible
    if (Deductible) then stTemp := SQLCollFouEncTVA('')
    else stTemp := SQLCollCliEncTVA('');

    stTemp := 'AND '+StringReplace(stTemp, 'AND','(', [])+' OR (ECHE2="X" AND AUXILIAIRE2="" AND TYPEMVT2="TTC"))'; //YMO FQ17747 30/11/2006 Rajout TIC TID

    stTemp := FindEtReplace(stTemp,'E_GENERAL','GENERAL1', True) ; // remplacement du nom du champ par son alias
  //  stTemp := StringReplace(stTemp, 'E_GENERAL','E1.E_GENERAL', [rfReplaceAll]);
    where := where + ' ' + stTemp ;
    // Condition sur exercice uniquement en facture directe et accompte
    if ((TEdition > 0) and (FEXERCICE.ItemIndex > 0)) then where := where + ' AND EXERCICE1="' + FEXERCICE.Value + '"';

    // nature auxiliaire et de piece
    if (Deductible) then
    begin
        where := where + ' AND (T_NATUREAUXI="FOU" OR T_NATUREAUXI="AUC" OR T_NATUREAUXI IS NULL)';
        if (TEdition = 0) then
          where := where + ' AND ((NATUREPIECE1="RF" OR NATUREPIECE1="OF")'
          + ' AND ((G_NATUREGENE<>"BQE" AND G_NATUREGENE<>"CAI") OR (NATUREPIECE1="OD" AND JEFFET1<>"X")))' {FQ18754 06.12.2007 YMO OD d'effets dans les factures par étapes}
        else if (TEdition = 1) then
        begin  {Lek 230506 FQ 17981 Compensation et Tiers payeur}
          where := where + ' AND (((NATUREPIECE1="RF" OR NATUREPIECE1="OF")'
          + ' AND ((G_NATUREGENE="BQE" OR G_NATUREGENE="CAI") OR (NATUREPIECE1="OD" AND JEFFET1="X")))'
          + FiltreCompenTP+')';
        end;
    end
    else
    begin
        where := where + ' AND (T_NATUREAUXI="CLI" OR T_NATUREAUXI="AUD" OR T_NATUREAUXI IS NULL)';
        if (TEdition = 0) then where := where + ' AND ((NATUREPIECE1="RC" OR NATUREPIECE1="OD" OR NATUREPIECE1="OC") AND G_NATUREGENE<>"BQE" AND G_NATUREGENE<>"CAI")'
        else if (TEdition = 1) then
        begin  {Lek 230506 FQ 17981 Compensation et Tiers payeur}
          where := where + ' AND (((NATUREPIECE1="RC" OR NATUREPIECE1="OD" OR NATUREPIECE1="OC")'+
                            ' AND (G_NATUREGENE="BQE" OR G_NATUREGENE="CAI"))'+
                            FiltreCompenTP+')' ;
        end;
    end;

    // Etat d'edit de TVA
    StValid :=' AND (EDITEETATTVA1="-"';
    if (GetCheckBoxState('FPREVALIDTVA') = cbChecked) then StValid := StValid+' OR EDITEETATTVA1="X"';
    if (GetCheckBoxState('FVALIDTVA') = cbChecked) then StValid := StValid+' OR EDITEETATTVA1="#"';
    StValid := StValid + ')';
    where := where + StValid;
    // lettrage non vide
    where := where + ' AND LETTRAGE1<>""' ;

    // ajout du critere de ligne de facture corespondant
    where := where + ' AND EXISTS (SELECT E3.E_GENERAL FROM ECRITURE E3 WHERE E3.E_AUXILIAIRE=AUXILIAIRE1 AND E3.E_GENERAL=GENERAL1 AND E3.E_ETATLETTRAGE=ETATLETTRAGE1 AND E3.E_LETTRAGE=LETTRAGE1';
    // nature de Piece     {17.12.07 YMO prise en compte OD}
    where := where + ' AND (E3.E_NATUREPIECE="OD" OR E3.E_NATUREPIECE="F' + GetControlText('FTNATPIECE') + '" OR E3.E_NATUREPIECE="A' + GetControlText('FTNATPIECE') + '")';
    // si etablissement
    if (GetControlText('FETABLISSEMENT') <>'') then where := where + ' AND E3.E_ETABLISSEMENT="' + GetControlText('FETABLISSEMENT') + '"';
    // si regime TVA
    if (GetControlText('FREGIMETVA') <> '') then where := where + ' AND E3.E_REGIMETVA="' + GetControlText('FREGIMETVA') + '"';
    where := where + ' AND E3.E_ECHEENC' + GetControlText('FNTVA') + '<>0)';

    //YMO 28/06/2006 Mise en stand by des modifs pour fonctionnement tiers-payeurs
    {b FP 19/05/2006 FQ17981
    if TEdition = 1 then
      begin
      if Deductible and (VH^.JalATP <> '') then          {Tiers payeur fournisseur
        where := where + ' AND (JOURNAL2<>"'+VH^.JalATP+'" OR (JOURNAL2="'+VH^.JalATP+'" AND JOURNAL2<>JOURNAL1 AND NUMEROPIECE2<>NUMEROPIECE1 AND NUMLIGNE2<>NUMLIGNE1 AND DATECOMPTABLE2<>DATECOMPTABLE1 AND EXERCICE2<>EXERCICE1) )';
      if (not Deductible) and (VH^.JalVTP <> '') then          {Tiers payeur fournisseur
        where := where + ' AND (JOURNAL2<>"'+VH^.JalVTP+'" OR (JOURNAL2="'+VH^.JalVTP+'" AND JOURNAL2<>JOURNAL1 AND NUMEROPIECE2<>NUMEROPIECE1 AND NUMLIGNE2<>NUMLIGNE1 AND DATECOMPTABLE2<>DATECOMPTABLE1 AND EXERCICE2<>EXERCICE1) )';
      end;
    e FP 19/05/2006}

    if remplace then
    for i:=0 to 30 do
    begin
      Where := StringReplace(Where, ArrAlias[i], ArrEcrit[i], [rfReplaceAll]);
    end;

    // retour :
    result := where;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 10/04/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPQRTVAFAC.MajOkEditionFAC;
var
  QEcr : TQuery ;
  Query : string;
begin
    InitMove(100,'');
    Begintrans;
        Query := 'SELECT E1.E_EXERCICE, E1.E_JOURNAL, E1.E_DATECOMPTABLE, E1.E_NUMEROPIECE, E1.E_NUMLIGNE, E1.E_DATEECHEANCE, '
        +' E2.E_NATUREPIECE, E2.E_ETABLISSEMENT, E2.E_ECHE, E2.E_AUXILIAIRE,E2.E_TYPEMVT,E1.E_GENERAL,'
        +' E1.E_NATUREPIECE,E1.E_EDITEETATTVA,E1.E_LETTRAGE'
        +' FROM ECRITURE E1 LEFT JOIN ECRITURE E2 ON E2.E_AUXILIAIRE=E1.E_AUXILIAIRE'
        +' AND E2.E_GENERAL=E1.E_GENERAL AND E2.E_ETATLETTRAGE=E1.E_ETATLETTRAGE AND E2.E_LETTRAGE=E1.E_LETTRAGE'
        +' LEFT JOIN GENERAUX ON E1.E_CONTREPARTIEGEN=G_GENERAL LEFT JOIN JOURNAL J1 ON E1.E_JOURNAL=J1.J_JOURNAL'
        +' LEFT JOIN JOURNAL J2 ON E2.E_JOURNAL=J2.J_JOURNAL LEFT JOIN TIERS ON E1.E_AUXILIAIRE=T_AUXILIAIRE'
        +' LEFT JOIN TXCPTTVA ON E2.E_REGIMETVA=TV_REGIME WHERE E1.E_EDITEETATTVA<>"#" AND ' + GetCritereFAC (True);

// GP N° 23485 le 01/09/2008 : c'est de la modif à l'arrache et de toute façon, ca reste de la grosse merde....
//    Query := 'SELECT EXERCICE1, JOURNAL1, DATECOMPTABLE1, NUMEROPIECE1, NUMLIGNE1 FROM CPQRTVAFAC WHERE EDITEETATTVA1<>"#" AND ' + GetCritereFAC (false);
    QEcr := OpenSQL( Query , True);

    While not (QEcr.Eof) do
    begin
    (* GP N° 23485 le 01/09/2008 : c'est de la modif à l'arrache et de toute façon, ca reste de la grosse merde....
        Query := 'UPDATE ECRITURE SET E_EDITEETATTVA="X"'
               + ' WHERE E_EXERCICE="' + QEcr.Findfield('EXERCICE1').AsString + '"'
               + ' AND E_JOURNAL="' + QEcr.Findfield('JOURNAL1').AsString + '"'
               + ' AND E_DATECOMPTABLE="' + UsDateTime(QEcr.Findfield('DATECOMPTABLE1').AsDateTime) + '"'
               + ' AND E_NUMEROPIECE=' + QEcr.Findfield('NUMEROPIECE1').AsString
               + ' AND E_NUMLIGNE=' + QEcr.Findfield('NUMLIGNE1').AsString
               + ' AND E_EDITEETATTVA<>"#"';
     *)
        Query := 'UPDATE ECRITURE SET E_EDITEETATTVA="X"'
               + ' WHERE E_EXERCICE="' + QEcr.Fields[0].AsString + '"'
               + ' AND E_JOURNAL="' + QEcr.Fields[1].AsString + '"'
               + ' AND E_DATECOMPTABLE="' + UsDateTime(QEcr.Fields[2].AsDateTime) + '"'
               + ' AND E_NUMEROPIECE=' + QEcr.Fields[3].AsString
               + ' AND E_NUMLIGNE=' + QEcr.Fields[4].AsString
               + ' AND E_EDITEETATTVA<>"#"';
        ExecuteSQL(Query);

        QEcr.next;
        MoveCur(False);
    end;
    Ferme(QEcr);

    Committrans;
    FiniMove;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 27/02/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
function TOF_CPQRTVAFAC.GetCritereEAC: string;
var
    where : string;
    StValid,etable,regime : String;
    Deductible : boolean;
begin
    where := '';

    // Recuperation des critere pour le XX_WHERE
    Deductible := (GetControlText('FNTYPETVA') = '1');
    etable := GetControlText('FETABLISSEMENT');
    regime := GetControlText('FREGIMETVA');

    // si deductible
    if (Deductible) then where := where + ' ' + SQLCollFouEncTVA('')
    else where := where + ' ' + SQLCollCliEncTVA('');
    // Condition sur exercice uniquement en facture directe et accompte
    if ((TEdition > 0) and (FEXERCICE.ItemIndex > 0)) then where := where + ' AND E_EXERCICE="' + FEXERCICE.Value + '"';
    // Accompte : date comptable
    where := where + ' AND E_DATECOMPTABLE>="' + usDateTime(StrToDate(GetControlText('FDATE'))) + '" AND E_DATECOMPTABLE<="' + usdatetime(StrToDate(GetControlText('FDATE_'))) + '"' ;
    // pour les accompte : regim de TVA
    if ((TEdition = 2) and (FREGIMETVA.ItemIndex > 0)) then where := where + ' AND E_REGIMETVA="' + FREGIMETVA.Value + '"';
    // nature auxiliaire et de piece
    if (Deductible) then
    begin
        where := where + ' AND (T_NATUREAUXI="FOU" OR T_NATUREAUXI="AUC")';
        If FUNIQACOMPTE.Checked then  {14/09/2007 YMO Tous les règlements non lettrés}
          where := where + ' AND E_NATUREPIECE="OF"'
        else
          where := where + ' AND (E_NATUREPIECE="OF" OR E_NATUREPIECE="RF" '
          +'OR (E_NATUREPIECE="OD" '
          +'AND (E_GENERAL LIKE "'+GetParamSocSecur('SO_COLLCLIENC','')
          +'%" OR E_GENERAL LIKE "'+GetParamSocSecur('SO_COLLFOUENC','')+'%")))';

    end
    else
    begin
        where := where + ' AND (T_NATUREAUXI="CLI" OR T_NATUREAUXI="AUD")';
        If FUNIQACOMPTE.Checked then
          where := where + ' AND E_NATUREPIECE="OC"'
        else
          where := where + ' AND (E_NATUREPIECE="OC" OR E_NATUREPIECE="RC" '
          +'OR (E_NATUREPIECE="OD" '
          +'AND (E_GENERAL LIKE "'+GetParamSocSecur('SO_COLLCLIENC','')
          +'%" OR E_GENERAL LIKE "'+GetParamSocSecur('SO_COLLFOUENC','')+'%")))';
    end;

    // Etat d'edit de TVA
    StValid :=' AND (E_EDITEETATTVA="-"';
    if (GetCheckBoxState('FPREVALIDTVA') = cbChecked) then StValid := StValid+' OR E_EDITEETATTVA="X"';
    if (GetCheckBoxState('FVALIDTVA') = cbChecked) then StValid := StValid+' OR E_EDITEETATTVA="#"';
    StValid := StValid + ')';
    where := where + StValid;
    // lettrage non vide
    where := where + ' AND E_ETATLETTRAGE="AL"';

    // retour :
    result := where;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Yann MORENO
Créé le ...... : 18/10/2007
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... :
*****************************************************************}
procedure TOF_CPQRTVAFAC.FUniqAcompteClick (Sender : TObject);
var NumTitre : integer; 
begin
   If FUniqAcompte.Checked then NumTitre:=2 else NumTitre:=13;

   If (IsInside(Ecran)) then THpanel(TFQRS1(Ecran).parent).insideTitle.Caption := MsgBox[NumTitre];
   TFQRS1(Ecran).Caption := MsgBox[NumTitre];
   SetControlText('FTYPEFAC',MsgBox[NumTitre]);
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 27/02/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPQRTVAFAC.MajOkEditionEAC;
var
    QEcr : TQuery;
    Query : string;
begin
    InitMove(100,'');
    Begintrans;

    Query := 'SELECT E_EXERCICE, E_JOURNAL, E_DATECOMPTABLE, E_NUMEROPIECE, E_NUMLIGNE FROM ECRITURE LEFT JOIN TIERS ON E_AUXILIAIRE=T_AUXILIAIRE LEFT JOIN TXCPTTVA ON TV_REGIME=E_REGIMETVA '
           + 'WHERE E_EDITEETATTVA<>"#" AND TV_CODETAUX="' + GetControlText('FTVA') + '" ' + GetCritereEAC ;
    QEcr := OpenSQL( Query , True);

    While not QEcr.Eof do
    begin
        Query := 'UPDATE ECRITURE SET E_EDITEETATTVA="X"'
               + ' WHERE E_EXERCICE="' + QEcr.Findfield('E_EXERCICE').AsString + '"'
               + ' AND E_JOURNAL="' + QEcr.Findfield('E_JOURNAL').AsString + '"'
               + ' AND E_DATECOMPTABLE="' + UsDateTime(QEcr.Findfield('E_DATECOMPTABLE').AsDateTime) + '"'
               + ' AND E_NUMEROPIECE=' + QEcr.Findfield('E_NUMEROPIECE').AsString
               + ' AND E_NUMLIGNE=' + QEcr.Findfield('E_NUMLIGNE').AsString
               + ' AND E_EDITEETATTVA<>"#"';

        ExecuteSQL(Query);

        QEcr.next ;
        MoveCur(False);
    end;
    Ferme(QEcr);

    Committrans;
    FiniMove;
end;

//================================================================================
// Initialization
//================================================================================
Initialization
    registerclasses([TOF_CPQRTVAFAC]);
end.

