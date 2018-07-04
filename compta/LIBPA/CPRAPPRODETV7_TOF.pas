{***********UNITE*************************************************
Auteur  ...... : BPY
Créé le ...... : 11/03/2004
Modifié le ... : 28/07/2005
Description .. : - BPY le 21/09/2004 - modification diverse et variée sur
Suite ........ : "l'etat recapitulatif si pointage complet" a la demande de la
Suite ........ : FFF
Suite ........ :
Suite ........ : SG6 19/11/04 : FQ 14969 (pb sous DB2) : Modification des
Suite ........ : stDate1900 par usdatetime(istdate1900)
Suite ........ :
Suite ........ : JP 28/07/05 : FQ 13731 : Gestion des comptes en devises
Mots clefs ... :
*****************************************************************}

Unit CPRAPPRODETV7_TOF;

//================================================================================
// Interface
//================================================================================
Interface

Uses
    StdCtrls,
    Controls,
    Classes,
{$IFDEF EAGLCLIENT}
    UTOB,
    eQRS1,
    MaineAGL,
{$ELSE}
    db,
    {$IFNDEF DBXPRESS} dbtables, {$ELSE} uDbxDataSet, {$ENDIF}
    QRS1,
    FE_Main,
{$ENDIF}
    forms,
    sysutils,
    ComCtrls,
    HCtrls,
    HEnt1,
    HMsgBox,
    UTOF,
    LookUp,
    Ent1,
    uLibWindows,
    Saisutil,
    TofMeth,
    Htb97,      //  TToolBarButton97
    CpteSav     // CTrouveContrePartie
    ;

//==================================================
// Externe
//==================================================
procedure CC_LanceFicheEtatRapproDetV7( vStArgument : string = '');

//==================================================
// Definition de class
//==================================================
Type
    TOF_CPRAPPRODETV7 = Class(TOF_METH)
        procedure OnNew                  ; override ;
        procedure OnDelete               ; override ;
        procedure OnUpdate               ; override ;
        procedure OnLoad                 ; override ;
        procedure OnArgument(S : String) ; override ;
        procedure OnClose                ; override ;

    private
        glbl_Date : THLabel;
        gtxt_DatePointage : THEdit;
        ggbx_Sens : THRadioGroup;
        gtxt_General : THEdit;
        FStCompteUtiliser : string;
        gtxt_Devise : THEdit;
        gtxt_DevPivot : THEdit;
        ggbx_Devise : THRadioGroup;
        gcbx_Recap : TCheckBox;
        
        XX_Rupture : THEdit;
        TE_General : THLabel;

        StType : string3;
        StDevise : string3;
        FStGeneral : string;

        FStArgument : string;

        procedure ggbx_DeviseClick(Sender : TObject);
        procedure OnExitGeneral(Sender : TOBject);
        procedure OnExitDatePointage(Sender : TOBject);
        procedure OnElipsisClickE_General(Sender : TObject);

        {JP 28/07/05 : FQ 13731 : Pour éxecuter OnExitGeneral après l'éventuel chargement du filtre}
        procedure FormAfterShow;
    end;

//================================================================================
// Implementation
//================================================================================
Implementation

//==================================================
// fonctions hors class
//==================================================
{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 11/03/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure CC_LanceFicheEtatRapproDetV7( vStArgument : string = '');
begin
    AGLLanceFiche('CP', 'CPRAPPRODETV7', '', '', vStArgument );
end;

//==================================================
// Evenements par default de la TOF
//==================================================
{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 11/03/2004
Modifié le ... : 26/12/2005 (GCO)
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPRAPPRODETV7.OnNew;
var lSt : string;
begin
  inherited;
  if FStArgument <> '' then
  begin
    lSt := ReadTokenSt(FStArgument);
    // GCO - 26/11/2004 - FQ 13180
    gtxt_General.SetFocus;
    gtxt_General.Text := lSt;
    FStGeneral := lSt;

    // Date de Pointage
    lSt := ReadTokenSt(FStArgument);
    gtxt_DatePointage.Text := lSt;

    // F9 Auto
    lSt := ReadtokenSt(FStArgument);
    if lSt = 'X' then
      TToolBarButton97(GetControl('BVALIDER', True)).Click;
  end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 11/03/2004
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPRAPPRODETV7.OnDelete;
begin
    Inherited;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 11/03/2004
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPRAPPRODETV7.OnUpdate;
var
    StV8 : String;
    Query, Query1, Query2 : string;
    lStControleDevise : string;
begin
    Inherited;
    // Remplissage du THEdit LECOMPTE
    SetControlText('LECOMPTE', FStCompteUtiliser);

//===== REQUETE 1 : ECRITURE =====
    // requete sur ECRITURE
    Query1 :=
          'SELECT 1 IDN,E_JOURNAL JAL,E_DATECOMPTABLE DCP,E_REFINTERNE REF,E_DEVISE DEV,E_GENERAL GEN,E_DATECOMPTABLE DVA,' +
          'E_NUMEROPIECE NUM,E_LIBELLE LIB,E_DEBIT DEN,E_CREDIT CRN,E_DEBITDEV DED,E_CREDITDEV CRD,E_AUXILIAIRE RPA,G_GENERAL,G_LIBELLE';

    //  pointage par journal ??
    if (VH^.PointageJal) then
    begin
        Query1 := Query1 +
              ' FROM GENERAUX,ECRITURE,JOURNAL WHERE J_JOURNAL=E_JOURNAL AND G_GENERAL=J_CONTREPARTIE'
            + ' AND E_QUALIFPIECE="N" AND (E_ECRANOUVEAU="N" OR E_ECRANOUVEAU="H")'
            + ' AND E_JOURNAL="' + gtxt_general.Text + '" AND E_GENERAL<>J_CONTREPARTIE';
    end
    else
    begin
        Query1 := Query1 +
              ' FROM GENERAUX,ECRITURE,JOURNAL WHERE J_JOURNAL=E_JOURNAL AND G_GENERAL=E_GENERAL'
            + ' AND E_QUALIFPIECE="N" AND (E_ECRANOUVEAU="N" OR E_ECRANOUVEAU="H")'
            + ' AND E_GENERAL="' + gtxt_general.Text + '"';
    end;
                                              
    // autre clause where : VH^.ExoV8
    StV8 := LWhereV8;
    if (StV8 <> '') then Query1 := Query1 + ' AND ' +StV8;

    // GCO - 07/11/2006 - FQ 17322
    Query1 := Query1 + ' AND E_CREERPAR <> "DET"';

    // selection uniquement des lignes d'ecriture plus vieille
    Query1 := Query1 + ' AND E_DATECOMPTABLE<="' + USDateTime(StrToDate(gtxt_DatePointage.Text)) + '"';

    if (gcbx_Recap.Checked) then
    begin
        // etat recapitulatif ....
        Query1 := Query1 + ' AND E_REFPOINTAGE="' + glbl_Date.Caption + '" AND (NOT E_DATEPOINTAGE="' + usdatetime(iDate1900) + '")';
    end
    else
    begin
        Query1 := Query1 + ' AND (((E_REFPOINTAGE<>"" OR E_REFPOINTAGE IS NOT NULL) AND E_DATEPOINTAGE>"' + USDateTime(StrToDate(gtxt_DatePointage.Text)) + '") OR (E_REFPOINTAGE="" OR E_REFPOINTAGE IS NULL))';
    end;

//===== REQUETE 2 : RELEVE DE COMPTE =====
    // requete sur EEXBQLIG
    Query2 :=
          'SELECT 2 IDN,CEL_CODEAFB JAL,CEL_DATEOPERATION DCP,CEL_REFPIECE REF,CEL_CODEAFB DEV,CEL_GENERAL GEN,CEL_DATEVALEUR DVA,CEL_NUMRELEVE' +
          ' NUM,CEL_LIBELLE LIB,CEL_DEBITEURO DEN,CEL_CREDITEURO CRN,CEL_DEBITDEV DEV,CEL_CREDITDEV CRD,CEL_REFPOINTAGE RPA,G_GENERAL,G_LIBELLE' +
          ' FROM GENERAUX,EEXBQLIG,EEXBQ WHERE EE_GENERAL=CEL_GENERAL AND EE_NUMRELEVE=CEL_NUMRELEVE AND G_GENERAL=CEL_GENERAL';

    // selection uniquement des lignes de releve plus vieille
    Query2 := Query2 + ' AND CEL_DATEOPERATION<="' + USDateTime(StrToDate(gtxt_DatePointage.Text)) + '"';

    if (gcbx_Recap.Checked) then
    begin
        // etat recapitulatif ....
        Query2 := Query2 + ' AND (NOT CEL_DATEPOINTAGE ="' + usdatetime(iDate1900) + '") AND CEL_GENERAL="' + gtxt_general.Text + '"';
    end
    else
    begin
        // element non pointé ...
        Query2 := Query2 + ' AND CEL_DATEPOINTAGE="' + usdatetime(iDate1900) + '" AND CEL_GENERAL="' + gtxt_general.Text + '"';
    end;

//===== VERIFICATION DEVISE =====
    // Devise des écritures à afficher
    if (not (StDevise = '')) then
    begin
        // On réutilise la requete sur ecriture pour vérifier la présence d'écriture avec devise différente
        lStControleDevise := Query1 + ' AND E_DEVISE <> "' + StDevise + '"';

        // Test d'existence des ecritures sur d'autres DEVISES que celle du compte de banque
        if ExisteSql(lStControleDevise) then
        begin
            if (VH^.PointageJal) then PgiInfo('Attention  : certaines écritures non pointées du journal ' + gtxt_general.Text + ' ont ' +
                                              'une monnaie de saisie différente de ' + StDevise , Ecran.Caption)
            else PgiInfo('Attention  : certaines écritures non pointées du compte ' + gtxt_general.Text + ' ont ' +
                         'une monnaie de saisie différente de ' + StDevise , Ecran.Caption);
        end;

        if (not (StDevise = V_PGI.DevisePivot)) then
        begin
            Query1 := Query1 + 'AND E_DEVISE ="' + StDevise + '"';
            Query2 := Query2 + 'AND EE_DEVISE ="' + StDevise + '"';
        end;
    end;

//===== UNION ALL et ORDER BY =====
    // set de la requete dans le whereSQL de la fiche
    Query := Query1 + ' UNION ALL ' + Query2 + ' ORDER BY IDN, G_GENERAL, DCP, NUM';
    TFQRS1(Ecran).WhereSQL := Query;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 11/03/2004
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPRAPPRODETV7.OnLoad;
var Q : TQuery;
    DatePointage : TDateTime;
    Solde, SoldeDev : Double;
    SoldeEuroB, SoldeDevB : Double;
    Exo : TExoDate;
    lStSql : string;
    notok : boolean;
begin
    Inherited;

    // Met une date par défaut si la date est incorrecte
    if (gtxt_DatePointage.Text = '') or not IsValidDate(gtxt_DatePointage.Text) then
    begin
        Q := OpenSQL('SELECT EE_DATEPOINTAGE, EE_REFPOINTAGE FROM EEXBQ WHERE EE_GENERAL="' + gtxt_General.Text + '" ORDER BY EE_DATEPOINTAGE DESC',True);
        if (not Q.EOF) then
        begin
            gtxt_DatePointage.Text := Q.Fields[0].AsString;
            glbl_Date.Caption      := Q.Fields[1].AsString;
        end
        else
        begin
            gtxt_DatePointage.Text := stDate1900;
            glbl_Date.Caption      := 'Aucune donnée';
        end;
        Ferme(Q);
    end;

    // Pour affichage dans l'état
    DatePointage := StrToDate(gtxt_DatePointage.Text);
    QuelDateDeExo(QUELEXODT(DatePointage),exo);

    // Solde Banquaire
    Q := OpenSQL('SELECT EE_NEWSOLDEDEBEURO - EE_NEWSOLDECREEURO, EE_NEWSOLDEDEB - EE_NEWSOLDECRE FROM EEXBQ WHERE EE_GENERAL="' + gtxt_General.Text + '" AND EE_REFPOINTAGE="' + glbl_Date.Caption + '" AND EE_DATEPOINTAGE="' + USDateTime(DatePointage) + '"',True);
    SoldeEuroB := Q.Fields[0].AsFloat;
    SoldeDevB := Q.Fields[1].AsFloat;
    Ferme(Q);

    // SET DES SOLDE SUR LA FICHE
    SetControlText('SOLDEEUROB',FloatToStr(SoldeEuroB));
    SetControlText('SOLDEDEVB',FloatToStr(SoldeDevB));

    // Calcul du solde comptable réel
    lStSql := 'SELECT SUM(E_CREDIT)-SUM(E_DEBIT), SUM(E_CREDITDEV)-SUM(E_DEBITDEV) ' +
              'FROM ECRITURE WHERE E_GENERAL = "' + FStCompteUtiliser + '" AND ' +
              'E_QUALIFPIECE = "N" AND (E_ECRANOUVEAU="N" OR E_ECRANOUVEAU="H") AND ' +
              'E_DATECOMPTABLE <= "' + USDateTime(DatePointage) + '" ';

    if (VH^.ExoV8.Code <> '') then
      lStSql := lStSql + 'AND E_DATECOMPTABLE >= "' + USDateTime(VH^.ExoV8.Deb) + '"';

    // BPY le 31/07/2004 => Fiche 13731 :  Les écritures ne sont pas restituées,
    // par contre le solde comptable en tient compte ce qui a pour effet de déséquilibrer l'état de rapprochement.
    if (not ((StDevise = '') or (StDevise = V_PGI.DevisePivot))) then
      lStSql := lStSql + 'AND E_DEVISE ="' + StDevise + '"';

    Q := OpenSQL(lStSql,True);
    Solde     := Q.Fields[0].AsFloat;
    SoldeDev  := Q.Fields[1].AsFloat;
    Ferme(Q);

    // SET DES SOLDE SUR LA FICHE
    SetControlText('SOLDE',FloatToStr(Solde));
    SetControlText('SOLDEDEV',FloatToStr(SoldeDev));

    // Sens ...
    if (ggbx_Sens.visible) then
    begin
        if (ggbx_Sens.value = '0') then SetControlText('SENS','1')  // sens Bancaire
        else SetControlText('SENS','2');                            // sans comptable
    end
    else SetControlText('SENS','0');

    // doit on faire un recap  ???
    if (not (StType = 'RIN')) then  // relevé manuel ou automatique
    begin
        notok := false;

        if (StType = 'INT') then    // relevé integré automatiquement
        begin
            // si il existe des lignes non pointées dans les lignes d'extrait bancaire alors pas de recap !
            // FQ 18478 : 28/08/2006 : modifications du test sur les relevé auto :
            //    on recherche la présence de lignes non pointées aussi sur les relevés précedents celui sélectionnés
            lStSql := 'SELECT CEL_GENERAL FROM EEXBQLIG'
                          + ' WHERE CEL_GENERAL="' + gtxt_General.Text + '" AND CEL_DATEPOINTAGE="' + usdatetime(iDate1900) + '"'
                          + ' AND ( CEL_NUMRELEVE <= (SELECT MAX(EE_NUMERO) FROM EEXBQ WHERE EE_GENERAL="' + gtxt_General.Text + '" '
                                                      + ' AND EE_REFPOINTAGE="' + glbl_Date.Caption + '" AND EE_ORIGINERELEVE="INT" ) ) ' ;
            if ExisteSQL( lStSQL ) then
              notok := true;
        end;

        // existe t'il des ecritures anterieures non pointées ??
        lStSql := 'SELECT E_DATEPOINTAGE FROM ECRITURE WHERE E_DATECOMPTABLE<="' + USDateTime(DatePointage) + '"' +
                  ' AND ((E_DATEPOINTAGE = "' + usdatetime(iDate1900) + '" AND E_REFPOINTAGE = "") OR ' +
                        '(E_DATEPOINTAGE > "' + USDateTime(DatePointage) + '" AND E_REFPOINTAGE IS NOT NULL))' +
                  ' AND E_QUALIFPIECE="N" AND (E_ECRANOUVEAU="N" OR E_ECRANOUVEAU="H")';

        if VH^.PointageJal then
          lStSql := lStSql + ' AND E_JOURNAL = "' + gtxt_General.Text + '" AND E_GENERAL <> "' + FStCompteUtiliser + '"'
        else
          lStsql := lStSql + ' AND E_GENERAL = "' + gtxt_General.Text + '"';

        if VH^.ExoV8.Code <> '' then
          lStSql := lStSql + ' AND E_DATECOMPTABLE >= "' + UsDateTime(VH^.ExoV8.Deb) + '"';

        if (ExisteSQL(lstSql)) then
          notok := true;

        if (notok) then
          gcbx_Recap.Checked := False
        else
          gcbx_Recap.Checked := True;

    end
    else gcbx_Recap.Checked := false;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 11/03/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPRAPPRODETV7.OnArgument(S : String );
begin
    Inherited;
    FStArgument := S;

    Ecran.HelpContext := 7619000 ; //RR FQ 16068 : 999999427;
    {JP 28/07/05 : FQ 13731 : Pour éxecuter OnExitGeneral après l'éventuel chargement du filtre}
    TFQRS1(Ecran).OnAfterFormShow := FormAfterShow;

    // Compte general
    gtxt_General := THEdit(GetControl('EGENERAL',True));
    gtxt_General.OnExit := OnExitGeneral;
    gtxt_General.OnElipsisClick := OnElipsisClickE_General;
    // libelle du compte general
    TE_General := THLabel(GetControl('TE_GENERAL', True));

    // date de la reference de pointage
    gtxt_DatePointage := THEdit(GetControl('DATEPOINTAGE', True));
    gtxt_DatePointage.EditMask := '!99/99/0000;1;_';
    gtxt_DatePointage.OnExit := OnExitDatePointage;
    // libelle de la reference pointage
    glbl_Date := THLabel(GetControl('LIBDATE',true));

    // Pour l'affichage en fonction de la devise choisie
    ggbx_Devise := THRadioGroup(GetControl('GBXDEVISE',True));
    ggbx_Devise.OnClick := ggbx_DeviseClick;

    {JP 04/05/05 : FQ 15482 : il ne faut mettre 'EUR' en dur}
    ggbx_Devise.Items[0] := VH^.LibDevisePivot;
    SetControlText('DEVISE', V_PGI.DevisePivot);
    SetControlText('DEVPIVOT', VH^.LibDevisePivot);
    {FIN FQ 15482 : il ne faut mettre 'EUR' en dur}

    // Sens si on a un relevé manuel
    ggbx_Sens := THRadioGroup(GetControl('RBTSENSETAT',true));

    // check box pour l'etat recap !
    gcbx_Recap := TCheckBox(GetControl('AFFRECAPITULATIF',true));

    // Critere de rupture
    XX_Rupture := THEdit(GetControl('XX_RUPTURE', True));

    // modif si pointage sur journal ou pas!
    if VH^.PointageJal then
    begin
        TE_General.Caption     := traduirememoire('Journal') ;
        gtxt_General.MaxLength := 3;
        XX_RUPTURE.Text        := 'E_JOURNAL';
    end
    else
    begin
        TE_General.Caption     := traduirememoire('Compte général') ;
        gtxt_General.MaxLength := VH^.CPta[fbGene].Lg;
        XX_RUPTURE.Text        := 'E_GENERAL';
    end;

    // init des variables
    SetControlText('LIBDATE','');
    gtxt_Devise := THEdit(GetControl('DEVISE',True));
    gtxt_DevPivot := THEdit(GetControl('DEVPIVOT',True));
    StType := 'RIN';
    StDevise := '';
    FStGeneral := '';
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 11/03/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPRAPPRODETV7.OnClose;
begin
    if (not (IsValidDate(GetControlText('DATEPOINTAGE')))) then SetControlText('DATEPOINTAGE',stDate1900);

    Inherited;
end;

//==================================================
// Autres Evenements
//==================================================
{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 11/03/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPRAPPRODETV7.ggbx_DeviseClick(Sender: TObject);
begin
    // Pour l'impression dans la monnaie sélectionnée
    if (ggbx_Devise.Value = '0') then
    begin
        {JP 04/05/05 : FQ 15482 : il ne faut mettre 'EUR' en dur}
        gtxt_Devise.Text := V_PGI.DevisePivot;
        gtxt_DevPivot.Text := VH^.LibDevisePivot;
    end
    else
    begin
        gtxt_Devise.Text := 'DEV';
        gtxt_DevPivot.Text := 'Devise';
    end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 11/03/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPRAPPRODETV7.OnExitGeneral(Sender: TOBject);
var
    lQuery : TQuery;
begin
    if Trim(gtxt_General.Text) = '' then Exit;

    if (not VH^.PointageJal) then
    begin
        if (length(gtxt_General.Text) < VH^.Cpta[fbGene].Lg) then
          OnElipsisClickE_General(sender);
    end;

    // changement de la clause plus de la date
    // FQ 19329 - CA 22/12/2006 - Suppression AND car inutile et pose problème en DB2
    //gtxt_DatePointage.Plus := 'AND EE_GENERAL="' + gtxt_General.Text + '"';
    gtxt_DatePointage.Plus := ' EE_GENERAL="' + gtxt_General.Text + '"';

    // Récupération du compte de contrepartie si on est en pointage sur JAL
    if VH^.PointageJal then
      FStCompteUtiliser := CTrouveContrePartie( gtxt_general.Text )
    else
      FStCompteUtiliser := gtxt_General.Text;

    lQuery := nil;
    try
        try
            // Recherche la devise du compte Général ou du Compte de Contrepartie
            {JP 28/07/05 : FQ 13731 : La requête se fait sur BQ_GENERAL et non BQ_CODE}
            lQuery := OpenSQL('SELECT BQ_DEVISE FROM BANQUECP WHERE BQ_GENERAL = "' + FStCompteUtiliser + '"', True);
            if (not lQuery.EOF) then
            begin
                StDevise := lQuery.FindField('BQ_DEVISE').AsString;

                if (StDevise = V_PGI.DevisePivot) then {JP 04/05/05 : FQ 15482}
                begin
                    ggbx_Devise.ItemIndex := 0;
                    ggbx_Devise.Enabled := true;
                end
                else
                begin
                    ggbx_Devise.ItemIndex := 1;
                    ggbx_Devise.Enabled := false;
                end;
            end;
        except
          on E : Exception do PgiError( 'Erreur de requête SQL : ' + E.Message , 'Attention' );
        end;
    finally
      Ferme( lQuery );
    end;

    // Cherche la référence de pointage la plus récente et met rien si ya pas !
    // Si et seulement si on change de compte general ! => Fiche n° 14016
    if (not (FStGeneral = gtxt_General.Text)) then
    begin
        try
           try
              lQuery := OpenSQL('SELECT EE_DATEPOINTAGE, EE_REFPOINTAGE, EE_ORIGINERELEVE FROM EEXBQ WHERE EE_GENERAL="' + gtxt_General.Text + '" ORDER BY EE_DATEPOINTAGE DESC',true);

              if (not lQuery.Eof) then
              begin
                  gtxt_DatePointage.Text := lQuery.FindField('EE_DATEPOINTAGE').AsString;
                  glbl_Date.Caption := lQuery.FindField('EE_REFPOINTAGE').AsString;
                  StType := lQuery.FindField('EE_ORIGINERELEVE').AsString;
                  ggbx_Sens.Visible := (not (StType = 'INT'));
              end
              else
              begin
                  gtxt_DatePointage.Text := '01/01/1900';
                  glbl_Date.Caption := '';
                  StType := 'RIN';
                  ggbx_Sens.Visible := false;
              end;
           except
                 on E : Exception do PgiError( 'Erreur de requête SQL : ' + E.Message, 'Attention' );
           end;
        finally
               Ferme(lQuery);
        end;
    end;

    FStGeneral := gtxt_General.Text;
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 16/08/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPRAPPRODETV7.OnExitDatePointage(Sender: TOBject);
var
    lQuery : TQuery;
begin
    lQuery := OpenSQL('SELECT EE_ORIGINERELEVE FROM EEXBQ WHERE EE_GENERAL="' + gtxt_General.Text + '" AND EE_DATEPOINTAGE="' + USDateTime(StrToDate(gtxt_DatePointage.Text)) + '"',true);

    if (not lQuery.Eof) then
    begin
        StType := lQuery.FindField('EE_ORIGINERELEVE').AsString;
        ggbx_Sens.Visible := (not (StType = 'INT'));
        gcbx_Recap.Enabled := true;
    end
    else
    begin
        StType := 'RIN';
        ggbx_Sens.Visible := false;
        gcbx_Recap.Checked := false;
        gcbx_Recap.Enabled := false;
    end;

    Ferme(lQuery);
end;

{***********A.G.L.***********************************************
Auteur  ...... : BPY
Créé le ...... : 11/03/2004
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPRAPPRODETV7.OnElipsisClickE_General(Sender: TObject);
var lSt : String;
begin
  if VH^.PointageJal then
  begin
    lSt := 'SELECT J_JOURNAL, J_LIBELLE FROM JOURNAL LEFT JOIN GENERAUX ON J_CONTREPARTIE=G_GENERAL WHERE J_NATUREJAL="BQE" AND G_POINTABLE="X"';
    LookUpList(THEdit(Sender), 'Journal', 'JOURNAL', 'J_JOURNAL', 'J_LIBELLE', '', 'J_JOURNAL', True, 0 , lSt)
  end
  else
  begin
    if (ctxTreso in V_PGI.PGIContexte) then
      LookUpList(THEdit(Sender), 'Compte général', 'GENERAUX', 'G_GENERAL', 'G_LIBELLE', 'G_NATUREGENE ="BQE" AND G_POINTABLE="X"', 'G_GENERAL', True, 0 )
    else
      LookUpList(THEdit(Sender), 'Compte général', 'GENERAUX', 'G_GENERAL', 'G_LIBELLE', 'G_POINTABLE="X"', 'G_GENERAL', True, 0 );
  end;
end;

{JP 28/07/05 : FQ 13731 : Pour éxecuter OnExitGeneral après l'éventuel chargement du filtre
{---------------------------------------------------------------------------------------}
procedure TOF_CPRAPPRODETV7.FormAfterShow;
{---------------------------------------------------------------------------------------}
begin
  if GetControlText('FFILTRES') <> '' then
    OnExitGeneral(gtxt_General);
end;

//================================================================================
// Initialization
//================================================================================
Initialization
    registerclasses([TOF_CPRAPPRODETV7]);
end.

