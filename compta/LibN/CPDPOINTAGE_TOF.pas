{***********UNITE*************************************************
Auteur  ...... : BPY
Créé le ...... : 05/09/2003
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : CPDPOINTAGE ()
Mots clefs ... : TOF;CPDPOINTAGE
*****************************************************************}

Unit CPDPOINTAGE_TOF;

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
    eMul,
{$ELSE}
    FE_Main,
    mul,
    db,
    {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
    HDB,
{$ENDIF}
    uTob,
    forms,
    sysutils,
    ComCtrls,
    HCtrls,
    HEnt1,
    HMsgBox,
    UTOF,
    Ent1,        // VH^.
    uLibWindows, // IIF
    HQRY,        // RecupWhereCritere
    HTB97,       // TToolBarButton97
    CPTESAV,     // RecalculTotPointeNew1
    Filtre,
    UListByUser,
    UTXml,
    UObjFiltres
    ;

//==================================================
// Definition de class
//==================================================
Type
    TOF_CPDPOINTAGE = Class (TOF)
        procedure OnLoad                 ; override ;
        procedure OnArgument(S : String) ; override ;
        procedure OnClose                ; override ;
    private
        FStEE_General : string;         // Attention ! peut contenir un GENERAL ou un JOURNAL
        FStEE_RefPointage : string;     // Contient la Réference de Pointage à utiliser
        FDtEE_DatePointage : TDatetime; // Contient la Date de Pointage

        TOBEcr : TOB;
        TOBExb : TOB;
        FListeECR : THGrid;
        FListeEXB : THGrid;

        TotSolde : THNumEdit;

        IsAllSelecting : boolean;
        IsCpteInDevise : boolean;
        FBoEnConsultation : Boolean;

        ObjFiltre : TObjFiltre;

        NomBase : string;

        Procedure OnFlipSelectionFListeECR(Sender : TObject);
        Procedure OnFlipSelectionFListeEXB(Sender : TObject);

        procedure OnClickBCherche(Sender : TObject);

        procedure OnClickBAgrandir(Sender : TObject);
        procedure OnClickBReduire(Sender : TObject);
        Procedure OnClickBOUVRIR(Sender : TObject);
        Procedure OnClickBMONSELECTALL(Sender : TObject);

        procedure OnChangeTotSolde(Sender : TObject);

        function RecupereCompte: string;

        function SelectAllEcr() : double;
        function SelectAllExb() : double;
    end;

//==================================================
// Fonctions d'ouverture de la fiche
//==================================================
procedure CPLanceFiche_DPointage(vStParam : string = '');


Implementation

uses
  Commun, UtilPgi;

//--------------------- Clé primaire Table ECRITURE --------------------------
// E_JOURNAL, E_EXERCICE, E_DATECOMPTABLE, E_NUMEROPIECE, E_NUMLIGNE, E_NUMECHE, E_QUALIFPIECE
//----------------------------------------------------------------------------

//==================================================
// Definition des Variables
//==================================================
var
    MESS : array [0..3] of string = (
    {00}    'Erreur : il n''y as pas de ligne de relevé bancaire correspondante.',
    {01}    'Erreur : il y as plusieurs ligne de relevé bancaire qui corresponde.',
    {02}    'Erreur lors du traitement. Vos modifications ne seront pas enregistrées.',
    {03}    'ERREUR'
            );

//==================================================
// Fonctions d'ouverture de la fiche
//==================================================
{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 05/09/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure CPLanceFiche_DPointage(vStParam : string = '');
begin
    AGLLanceFiche('CP','CPDPOINTAGE','','',vStParam);
end;

//==================================================
// Evenements par default de la TOF
//==================================================
{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 05/09/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPDPOINTAGE.OnLoad;
var
    titre : string;
begin
    Inherited;

    // set du caption ....
    titre := IIF(VH^.PointageJal,'Journal','Compte général');
    Ecran.Caption := titre + ' ' + FStEE_General + ' - Pointage au ' + DateToStr(FDtEE_DatePointage) + ' (' + FStEE_RefPointage + ')';
    UpDateCaption(Ecran);

    ObjFiltre.Charger;

    // on recharge les liste !!
    OnClickBCherche(nil);
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 05/09/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPDPOINTAGE.OnArgument(S : String);
var
  Composants : TControlFiltre;
  lQuery : TQuery;
  lst : string;
begin
    Inherited;
    // init
    IsAllSelecting := false;
    FBoEnConsultation := CEstPointageEnConsultationSurDossier;

    // recup les grid !
    FListeECR := THGrid(GetControl('FListeECR',true));
    FListeEXB := THGrid(GetControl('FListeEXB',true));

    // set des proprieté du THGrid
    FListeECR.ColAligns[2] := taCenter;
    FListeECR.ColAligns[3] := taCenter;
    FListeECR.ColAligns[6] := taCenter;
    FListeECR.ColAligns[7] := taRightJustify;
    FListeECR.ColAligns[8] := taRightJustify;

    // set des proprieté du THGrid
    FListeEXB.ColAligns[2] := taCenter;
    FListeEXB.ColAligns[3] := taCenter;
    FListeEXB.ColAligns[6] := taRightJustify;
    FListeEXB.ColAligns[7] := taRightJustify;

    // set des onExit des liste
    // GCO - 14/10/2004  - FQ 14804
    if FBoEnConsultation then
    begin
      FListeECR.MultiSelect := False;
      FListeEXB.MultiSelect := False;
    end
    else
    begin
      FListeECR.OnFlipSelection := OnFlipSelectionFListeECR;
      FListeEXB.OnFlipSelection := OnFlipSelectionFListeEXB;
    end;

    // recup du solde !
    TotSolde := THNumEdit(GetControl('TOTSOLDE',true));
    TotSolde.OnChange := OnChangeTotSolde;

    // assignation des btn
    TToolBarButton97(GetControl('BOUVRIR',true)).OnClick := OnClickBOUVRIR;
    TToolBarButton97(GetControl('BMONSELECTALL',true)).OnClick := OnClickBMONSELECTALL;
    TToolBarButton97(GetControl('BCherche',true)).OnClick := OnClickBCherche;
    TToolBarButton97(GetControl('BAgrandir',true)).OnClick := OnClickBAgrandir;
    TToolBarButton97(GetControl('BReduire',true)).OnClick := OnClickBReduire;

    // Compte GENERAL ou JOURNAL
    lSt := ReadTokenSt(S);
    FStEE_General := lSt;
    if (VH^.PointageJal) then FStEE_General := RecupereCompte;

    // Date de Pointage
    lSt := ReadTokenSt(S);
    FDtEE_DatePointage := StrToDate(lSt);

    // Référence de POINTAGE
    lSt := ReadTokenSt(S);
    FStEE_RefPointage := lSt;

    // set des valeur parametre dans les critere de selection
    SetControlText('E_GENERAL',FStEE_General);
    SetControlText('E_REFPOINTAGE',FStEE_RefPointage);
    SetControlText('CEL_GENERAL',FStEE_General);
    SetControlText('CEL_REFPOINTAGE',FStEE_RefPointage);

    if IsTresoMultiSoc then NomBase := ReadTokenSt(s)
                       else NomBase := '';

    // Cherche la devise du compte
    lQuery := OpenSQL('SELECT EE_DEVISE FROM ' + GetTableDossier(NomBase, 'EEXBQ') + ' WHERE EE_GENERAL="' + FStEE_General + '" AND EE_DATEPOINTAGE="' + UsDateTime(FDtEE_DatePointage) + '" AND EE_REFPOINTAGE="' + FStEE_RefPointage + '"', True);
    if (not lQuery.EOF) then
    begin
        if (lQuery.FindField('EE_DEVISE').AsString = 'EUR') then IsCpteInDevise := false
        else IsCpteInDevise := true;
    end;

  {JP 28/09/06 : Mise en place de l'objet Filtre}
  Composants.Filtres  := THValComboBox   (Getcontrol('FFILTRES'));
  Composants.Filtre   := TToolBarButton97(Getcontrol('BFILTRE'));
  Composants.PageCtrl := TPageControl    (Getcontrol('PAGES'));
  {Distinction des filtres pour le cas où le client utilise le traitement à la fois dans
   une Trésorerie MS et en comptabilité}
  if IsTresoMultiSoc then
    ObjFiltre := TObjFiltre.Create(Composants, Ecran.Name + 'TR')
  else
    ObjFiltre := TObjFiltre.Create(Composants, Ecran.Name);
  {JP 28/09/06 Il me parait plus propre de créer les tobs dans le ongarument que OnClickBCherche
   où elles étaient systématiquement recréées sans être libérées}
  TOBEcr := TOB.Create('£ECRITURE',nil,-1);
  TOBExb := TOB.Create('£EEXBQLIG',nil,-1);
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 05/09/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPDPOINTAGE.OnClose;
begin
    Inherited;
    If Assigned(ObjFiltre) then FreeAndNil(ObjFiltre);

    FreeAndNil(TOBEcr);
    FreeAndNil(TOBExb);
end;

//==================================================
// Autres Evenements
//==================================================
{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 19/05/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPDPOINTAGE.OnClickBCherche(Sender : TObject);
var
    WhereSQLEcr : string;
    WhereSQLExb : string;
begin
    Inherited;

    SourisSablier;

    {On vide les tobs}
    TOBEcr.ClearDetail;
    TOBExB.ClearDetail;

    {Récupération des clauses where}
    WhereSQLEcr := RecupWhereCritere(TPageControl(GetControl('PAGES',true)));
    WhereSQLExb := RecupWhereCritere(TPageControl(GetControl('TPAGESEXB',true)));

    TOBEcr.LoadDetailFromSQL('SELECT * FROM ' + GetTableDossier(NomBase, 'ECRITURE') + ' ' + WhereSQLEcr);
    if (IsCpteInDevise) then TOBEcr.PutGridDetail(FListeECR,false,false,'E_DATECOMPTABLE;E_NUMROPIECE;E_NUMLIGNE;E_REFINTERNE;E_LIBELLE;E_DEVISE;E_CREDITDEV;E_DEBITDEV;',true)
                        else TOBEcr.PutGridDetail(FListeECR,false,false,'E_DATECOMPTABLE;E_NUMROPIECE;E_NUMLIGNE;E_REFINTERNE;E_LIBELLE;E_DEVISE;E_CREDIT;E_DEBIT;',true);

    TOBExb.LoadDetailFromSQL('SELECT * FROM ' + GetTableDossier(NomBase, 'EEXBQLIG') + ' ' + WhereSQLExb);
    if (IsCpteInDevise) then TOBExb.PutGridDetail(FListeEXB,false,false,'CEL_DATEOPERATION;CEL_NUMRELEVE;CEL_NUMLIGNE;CEL_REFORIGINE;CEL_LIBELLE;CEL_CREDITDEV;CEL_DEBITDEV;',true)
                        else TOBExb.PutGridDetail(FListeEXB,false,false,'CEL_DATEOPERATION;CEL_NUMRELEVE;CEL_NUMLIGNE;CEL_REFORIGINE;CEL_LIBELLE;CEL_CREDITEURO;CEL_DEBITEURO;',true);

    SourisNormale;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 19/05/2004
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPDPOINTAGE.OnClickBAgrandir(Sender: TObject);
begin
    SetControlVisible('PAGES',false);
    SetControlVisible('BAgrandir',false);
    SetControlVisible('Breduire',true);
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 19/05/2004
Modifié le ... :   /  /    
Description .. :
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPDPOINTAGE.OnClickBReduire(Sender : TObject);
begin
    SetControlVisible('PAGES',true);
    SetControlVisible('BAgrandir',true);
    SetControlVisible('Breduire',false);
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 05/09/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPDPOINTAGE.OnClickBOUVRIR(Sender : TObject);
var
    i,j : integer;
    WhereSQLEcr : string;
    WhereSQLExb : string;
begin
    SourisSablier;

    try
        BeginTrans;

        //===================
        // liste des ecriture
        if ({FListeECR.AllSelected => marche pas le allselected}FListeECR.NbSelected = TOBEcr.Detail.Count) then // si tout ??
        begin
            // recup du where de remplisage ..... ;)
            WhereSQLEcr := RecupWhereCritere(TPageControl(GetControl('PAGES',true)));
            ExecuteSQL('UPDATE ' + GetTableDossier(NomBase, 'ECRITURE') + ' SET E_TRESOSYNCHRO = "MOD", E_REFPOINTAGE="", E_DATEPOINTAGE="' + UsDateTime(iDate1900) + '" ' + WhereSQLEcr);
        end
        else // si pas tout ....
        begin
            for j := 0 to FListeECR.NbSelected-1 do
            begin
                FListeECR.GotoLeBookMark(j);
                i := FListeECR.Row - 1;
                // traitement
                ExecuteSQL('UPDATE ' + GetTableDossier(NomBase, 'ECRITURE') + ' SET E_TRESOSYNCHRO = "MOD", E_REFPOINTAGE="", E_DATEPOINTAGE="' + UsDateTime(iDate1900) + '" WHERE E_JOURNAL="' + TOBEcr.Detail[i].GetValue('E_JOURNAL') + '" AND E_EXERCICE="' + TOBEcr.Detail[i].GetValue('E_EXERCICE') + '" AND E_DATECOMPTABLE="' + USDateTime(TOBEcr.Detail[i].GetValue('E_DATECOMPTABLE')) + '" AND E_NUMEROPIECE=' + IntToStr(TOBEcr.Detail[i].GetValue('E_NUMEROPIECE')) + ' AND E_NUMLIGNE=' + IntToStr(TOBEcr.Detail[i].GetValue('E_NUMLIGNE')) + ' AND E_NUMECHE=' + IntToStr(TOBEcr.Detail[i].GetValue('E_NUMECHE')) + ' AND E_QUALIFPIECE="' + TOBEcr.Detail[i].GetValue('E_QUALIFPIECE') + '"');
            end;
        end;

        FListeECR.AllSelected := false;
        FListeECR.ClearSelected;

        //===========================
        // liste des relevé banquaire
        if ({FListeEXB.AllSelected => marche pas le allselected}FListeEXB.NbSelected = TOBExb.Detail.Count) then // si tout ??
        begin
            // recup du where de remplisage ..... ;)
            WhereSQLExb := RecupWhereCritere(TPageControl(GetControl('TPAGESEXB',true)));
            ExecuteSQL('UPDATE ' + GetTableDossier(NomBase, 'EEXBQLIG') + ' SET CEL_DATEPOINTAGE="' + UsDateTime(iDate1900) + '" ' + WhereSQLExb);
        end
        else // si pas tout ....
        begin
            for j := 0 to FListeEXB.NbSelected-1 do
            begin
                FListeEXB.GotoLeBookMark(j);
                i := FListeEXB.Row - 1;
                ExecuteSQL('UPDATE ' + GetTableDossier(NomBase, 'EEXBQLIG') + ' SET CEL_DATEPOINTAGE="' + UsDateTime(iDate1900) + '" WHERE CEL_GENERAL="' + TOBExb.Detail[i].GetValue('CEL_GENERAL') + '" AND CEL_NUMRELEVE=' + IntToStr(TOBExb.Detail[i].GetValue('CEL_NUMRELEVE')) + ' AND CEL_NUMLIGNE=' + IntToStr(TOBExb.Detail[i].GetValue('CEL_NUMLIGNE')) );
            end;
        end;

        FListeEXB.AllSelected := false;
        FListeEXB.ClearSelected;

    // si ca plante
    except
        PGIBox(MESS[2],MESS[3]);
        RollBack;
        exit;
    end;

    // sinon !!! ;)
    commitTrans;
    // recalcul des cumul pointé !
    RecalculTotPointeNew1(FStEE_General, NomBase);

    SourisNormale;

    // on recharge les liste !!
    OnClickBCherche(nil);
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 19/09/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPDPOINTAGE.OnClickBMONSELECTALL(Sender : TObject);
begin
    IsAllSelecting := true;
    TotSolde.Value := SelectAllEcr() + SelectAllExb();
    IsAllSelecting := false;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 19/09/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPDPOINTAGE.OnChangeTotSolde(Sender : TObject);
begin
    if ((abs(TotSolde.value) < 0.01) and ((not (FListeECR.nbSelected = 0)) or (not (FListeEXB.nbSelected = 0)))) then SetControlEnabled('BOUVRIR',true) // FQ 13077
    else SetControlEnabled('BOUVRIR',false);
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 19/09/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
Procedure TOF_CPDPOINTAGE.OnFlipSelectionFListeECR(Sender : TObject);
begin
    // comptabilisation de la ligne
    if (TOBEcr.Detail.Count > 0) then
    begin
        if (IsCpteInDevise) then
        begin
            if (FListeECR.IsSelected(FListeECR.Row)) then TotSolde.Value := TotSolde.Value + ( TOBEcr.Detail[FListeECR.Row-1].GetValue('E_CREDITDEV') - TOBEcr.Detail[FListeECR.Row-1].GetValue('E_DEBITDEV') )
            else TotSolde.Value := TotSolde.Value - ( TOBEcr.Detail[FListeECR.Row-1].GetValue('E_CREDITDEV') - TOBEcr.Detail[FListeECR.Row-1].GetValue('E_DEBITDEV') );
        end
        else
        begin
            if (FListeECR.IsSelected(FListeECR.Row)) then TotSolde.Value := TotSolde.Value + ( TOBEcr.Detail[FListeECR.Row-1].GetValue('E_CREDIT') - TOBEcr.Detail[FListeECR.Row-1].GetValue('E_DEBIT') )
            else TotSolde.Value := TotSolde.Value - ( TOBEcr.Detail[FListeECR.Row-1].GetValue('E_CREDIT') - TOBEcr.Detail[FListeECR.Row-1].GetValue('E_DEBIT') );
        end;
    end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 19/09/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
Procedure TOF_CPDPOINTAGE.OnFlipSelectionFListeEXB(Sender : TObject);
begin
    // comptabilisation de la ligne
    if (TOBExb.Detail.Count > 0) then
    begin
        if (IsCpteInDevise) then
        begin
            if (FListeEXB.IsSelected(FListeEXB.Row)) then TotSolde.Value := TotSolde.Value + ( TOBExb.Detail[FListeEXB.Row-1].GetValue('CEL_CREDITDEV') - TOBExb.Detail[FListeEXB.Row-1].GetValue('CEL_DEBITDEV') )
            else TotSolde.Value := TotSolde.Value - ( TOBExb.Detail[FListeEXB.Row-1].GetValue('CEL_CREDITDEV') - TOBExb.Detail[FListeEXB.Row-1].GetValue('CEL_DEBITDEV') );
        end
        else
        begin
            if (FListeEXB.IsSelected(FListeEXB.Row)) then TotSolde.Value := TotSolde.Value + ( TOBExb.Detail[FListeEXB.Row-1].GetValue('CEL_CREDITEURO') - TOBExb.Detail[FListeEXB.Row-1].GetValue('CEL_DEBITEURO') )
            else TotSolde.Value := TotSolde.Value - ( TOBExb.Detail[FListeEXB.Row-1].GetValue('CEL_CREDITEURO') - TOBExb.Detail[FListeEXB.Row-1].GetValue('CEL_DEBITEURO') );
        end;
    end;
end;

//==================================================
// Autres fonctions de la class
//==================================================
{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 05/09/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
function TOF_CPDPOINTAGE.RecupereCompte: string;
var
    lQuery: TQuery;
begin
    Result := '';

    try
        lQuery := nil;
        try
            lQuery := OpenSql('SELECT G_GENERAL FROM GENERAUX LEFT JOIN JOURNAL ON J_CONTREPARTIE = G_GENERAL WHERE J_JOURNAL = "' + FStEE_General + '"',True);
            Result := lQuery.FindField('G_GENERAL').AsString;
        except
            on E: Exception do PgiError('Erreur de requête SQL : ' + E.Message, 'Fonction _RecupereCompte');
        end;

    finally
        Ferme(lQuery);
    end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 12/01/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
function TOF_CPDPOINTAGE.SelectAllEcr() : double;
var
    i : integer;
begin
    // selection
    FListeEcr.AllSelected := (not FListeEcr.AllSelected);

    // recalcul du sold
    result := 0.0;
    if (IsCpteInDevise) then
    begin
        if (TOBEcr.Detail.Count > 0) then for i := 1 to FListeEcr.RowCount do if (FListeEcr.IsSelected(i)) then result := result + ( TOBEcr.Detail[i-1].GetValue('E_CREDITDEV') - TOBEcr.Detail[i-1].GetValue('E_DEBITDEV') );
    end
    else
    begin
        if (TOBEcr.Detail.Count > 0) then for i := 1 to FListeEcr.RowCount do if (FListeEcr.IsSelected(i)) then result := result + ( TOBEcr.Detail[i-1].GetValue('E_CREDIT') - TOBEcr.Detail[i-1].GetValue('E_DEBIT') );
    end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 12/01/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
function TOF_CPDPOINTAGE.SelectAllExb() : double;
var
    i : integer;
begin
    // selection
    FLIsteExb.AllSelected := (not FLIsteExb.AllSelected);

    // recalcul du sold
    result := 0.0;
    if (IsCpteInDevise) then
    begin
        if (TOBExb.Detail.Count > 0) then for i := 1 to FLIsteExb.RowCount do if (FLIsteExb.IsSelected(i)) then result := result + ( TOBExb.Detail[i-1].GetValue('CEL_CREDITDEV') - TOBExb.Detail[i-1].GetValue('CEL_DEBITDEV') );
    end
    else
    begin
        if (TOBExb.Detail.Count > 0) then for i := 1 to FLIsteExb.RowCount do if (FLIsteExb.IsSelected(i)) then result := result + ( TOBExb.Detail[i-1].GetValue('CEL_CREDITEURO') - TOBExb.Detail[i-1].GetValue('CEL_DEBITEURO') );
    end;
end;

//================================================================================
// Initialization
//================================================================================
Initialization
    registerclasses([TOF_CPDPOINTAGE]);
    
end.
