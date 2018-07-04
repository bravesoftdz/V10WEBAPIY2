{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 18/03/2003
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : POINTAGEMUL ()
Mots clefs ... : TOF;POINTAGEMUL
*****************************************************************}
unit uTOFPointageMulV7;

interface

uses
    StdCtrls,
    Controls,
    Classes,
{$IFDEF EAGLCLIENT}
    Maineagl, // AGLLanceFiche EAGL
    eMul, // TFMUL(ECRAN).BChercheClick( nil )
{$ELSE}
    Fe_main, // AGLLanceFiche AGL
    Mul, // TFMUL(ECRAN).BChercheClick( nil )
    DB,
    DBTables,
    Hdb,
{$ENDIF}
    forms,
    sysutils,
    ComCtrls,
    HCtrls,
    HEnt1,
    HMsgBox,
    UTOF,
    Htb97, // TToolBarButton97
    ParamSoc, // GetParamSoc
    Ent1, // VH^.
    uTob, // TOB
    uLibEcriture, // CEstSaisieOuverte
    uLibWindows, // IIF(
    Windows, // TheRect, VK_
    Menus, // TPopUpMenu
    LookUp, // LookUpList
    utilPGI, // _Blocage
    uTofPointageEcr, // CPLanceFiche_Pointag
    AGLInit, // TheMulQ
    ED_Tools, // InitMoveProgressForm
    HStatus, // InitMove
    CpteSav // CTrouveContrePartie
    ;

type
    TOF_POINTAGEMULV7 = class(TOF)
        procedure OnNew; override;
        procedure OnDelete; override;
        procedure OnUpdate; override;
        procedure OnLoad; override;
        procedure OnArgument(S : string); override;
        procedure OnDisplay; override;
        procedure OnClose; override;
        procedure OnCancel; override;

    public
{$IFDEF EAGLCLIENT}
        FListe : THGrid;
{$ELSE}
        FListe : THDBGrid;
{$ENDIF}

        TEE_GENERAL : THLabel;
        EE_GENERAL : THEdit;
        EE_DATEPOINTAGE : THEdit;
        EE_REFPOINTAGE : THEdit;

        Q : TQuery;
        BInsert : TToolBarButton97;
        BDelete : TToolBarButton97;
        BInfo : TToolBarButton97;
        BCherche : TToolBarButton97;
        BOuvrir : TToolBarButton97;
        BRechercher : TToolBarButton97;
        BEtatRappro : TToolBarButton97;

        POPF11 : TPopUpMenu;

    private
        FStArgument : string;
        FBoConsultation : Boolean;

        procedure OnDblClickFListe  (Sender : TObject);
        procedure OnClickBInsert    (Sender : TObject);
        procedure OnClickBDelete    (Sender : TObject);
        procedure OnClickBInfo      (Sender : TObject);
        procedure OnClickBEtatRappro(Sender : TObject);

        procedure OnElipsisClickEE_General(Sender : TObject);
        procedure OnElipsisClickEE_DatePointage(Sender : TObject);
        procedure OnElipsisClickEE_RefPointage(Sender : TObject);
        procedure OnExitEE_General(Sender : TObject);

        procedure OnKeyDownEcran(Sender : TObject; var Key : Word; Shift : TShiftState);

        procedure SuppressionReference;
    end;

/////////////////////////////////////////////////////////////////////
function CPLanceFiche_PointageMulV7(vStParam : string = '') : string;

////////////////////////////////////////////////////////////////
procedure CSupprimeRefPointageV7(vCompteOuJournal, vRefPointage : string; vDatePointage : TDateTime);

////////////////////////////////////////////////////////////////
procedure CDepointeEcriture(vCompteOuJournal, vRefPointage : string; vDatePointage : TDateTime);

///////////////////////////
procedure CRazModePointage;

implementation

uses uTOMEEXBQ,       // CPLanceFiche_PointageFic
     CPRAPPRODETV7_TOF; // CC_LanceFicheEtatRapproDet

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 19/03/2003
Modifié le ... :   /  /
Description .. : Lancement de la fiche CPPOINTAGEMUL
Mots clefs ... :
*****************************************************************}
function CPLanceFiche_PointageMulV7(vStParam : string = '') : string;
begin
  //if _Blocage(['nrCloture'], False, 'nrPointage') then Exit ;
  Result := AGLLanceFiche('CP', 'CPPOINTAGEMULV7', '', '', vStParam);
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 07/07/2003
Modifié le ... :   /  /
Description .. : Supprime toutes les références de pointage manuelles
Suite ........ : Enleve les infos de pointage des écritures associées
Suite ........ : Met à zéro les champs G_TOTDEB,etc.... de la Table GENERAUX
Mots clefs ... :
*****************************************************************}
procedure CRazModePointage;
var lQuery : TQuery;
    lStTitrePGIInfo : string;
begin
  if not _BlocageMonoPoste(True) then Exit;

  // GCO - 27/10/2004 FQ 14804
  if CEstPointageEnConsultationSurDossier then
  begin
    PgiInfo('Vous avez indiqué une liaison avec une comptabilité ' +
            RechDom('CPLIENCOMPTABILITE',GetParamSocSecur('SO_CPLIENGAMME', ''), False) +
            ' et la gestion du pointage ' + #10 +
            'est effectuée ' + RechDom('CPPOINTAGESX', GetParamSocSecur('SO_CPPOINTAGESX', ''), False) + '. ' +
            'Vous n''avez pas accès à cette commande.', 'Remise à zéro du pointage');
    Exit;
  end;

  lStTitrePGIInfo := 'Remise à zéro du pointage';
  if VH^.PointageJal then
    lStTitrePGIInfo := lStTitrePGIInfo + ' sur journal'
  else
    lStTitrePGIInfo := lStTitrePGIInfo + ' sur compte général';
  LstTitrePGIInfo:=TraduireMemoire(LstTitrePGIInfo) ;

  if PgiAsk('Attention : Vous allez supprimer le pointage du dossier. Confirmez vous le traitement ?', lStTitrePGIInfo) = MrNo then
    Exit;

  if PgiAsk('Attention : Le pointage va être complètement supprimé. Voulez vous abandonner le traitement ?', lStTitrePGIInfo) = MrYes then
    Exit;

  lQuery := nil;
  try
    try
      BeginTrans;
      InitMoveProgressForm(nil, lStTitrePGIInfo, 'Suppression', 3, True, True);

      MoveCurProgressForm('Suppression des références de pointage');
      ExecuteSql('DELETE FROM EEXBQ WHERE (EE_ORIGINERELEVE<>"INT" OR EE_ORIGINERELEVE IS NULL)');

      MoveCurProgressForm('Suppression du pointage sur les écritures');
      ExecuteSql('UPDATE ECRITURE SET E_REFPOINTAGE = "", ' +
{$IFDEF TRSYNCHRO}
      {FQ Tréso 10019 : maj du champ E_TRESOSYNCHRO, Pour que lors de prochaine synchronisation
                                  de la trésorerie le champ TE_DATERAPPRO puisse être mis à jour}
      'E_TRESOSYNCHRO = "MOD", ' +
{$ENDIF}
      'E_DATEPOINTAGE  = "' + UsDateTime(iDate1900) + '" ' + 'WHERE E_REFPOINTAGE <> ""');

      ExecuteSql('UPDATE EEXBQLIG SET CEL_DATEPOINTAGE = "' + UsDateTime(iDate1900) + '"' +
                 'WHERE CEL_DATEPOINTAGE <> "' + UsDateTime(iDate1900) + '"');

      MoveCurProgressForm('Suppression des totaux pointés des comptes généraux');
      ExecuteSql('UPDATE GENERAUX SET G_TOTDEBPTP = 0, G_TOTCREPTP = 0, G_TOTDEBPTD = 0, G_TOTCREPTD = 0');

      PgiInfo('Le traitement s''est correctement terminé.', lStTitrePGIInfo);
    except
      on E : Exception do
      begin
        PgiError('Erreur de requête SQL : ' + E.Message, 'Fonction : CRazModePointage');
        RollBack;
      end;
    end;

  finally
    Ferme(lQuery);
    CommitTrans;
    FiniMoveProgressForm;
    _DeblocageMonoPoste(True);
  end;
end;

////////////////////////////////////////////////////////////////////////////////

procedure CSupprimeRefPointageV7(vCompteOuJournal, vRefPointage : string; vDatePointage : TDateTime);
begin
    try
        // Suppression de la référence dans EEXBQ
        ExecuteSQL('DELETE FROM EEXBQ WHERE EE_GENERAL = "' + vCompteOuJournal + '"' +
            ' AND EE_REFPOINTAGE = "' + vRefPointage + '"' +
            ' AND EE_DATEPOINTAGE = "' + UsDateTime(vDatePointage) + '"' +
            ' AND EE_NUMERO=1');
    except
        on E : Exception do PgiError('Erreur de requête SQL : ' + E.Message, 'Fonction : CSupprimeRefPointage');
    end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 30/06/2003
Modifié le ... :   /  /
Description .. : Depointe les écritures d'une référence de pointage
Suite ........ : Si Pointage sur Compte --> vCompteJournal = Compte Général
Suite ........ : Si Pointage sur Journal --> vCompteJournal = Compte de Contreparie du journal
Mots clefs ... :
*****************************************************************}

procedure CDepointeEcriture(vCompteOuJournal, vRefPointage : string; vDatePointage : TDateTime);
var
    lContrePartie : string;
begin
    if VH^.PointageJal then
    begin
        lContrePartie := CTrouveContrePartie(vCompteOuJournal);
        if lContrePartie = '' then
        begin
            PGIError('Impossible de trouver le compte de contrepartie du journal : ' + vCompteOuJournal, 'Traitement annulé');
            Exit;
        end;
    end;

    try
        if VH^.PointageJal then
        begin
            // Suppression du pointage des écritures
            ExecuteSQL('UPDATE ECRITURE SET E_REFPOINTAGE = "", ' +
{$IFDEF TRSYNCHRO}
                {FQ Tréso 10019 : maj du champ E_TRESOSYNCHRO, Pour que lors de prochaine synchronisation
                                  de la trésorerie le champ TE_DATERAPPRO puisse être mis à jour}
                'E_TRESOSYNCHRO = "MOD", ' +
{$ENDIF}
                'E_DATEPOINTAGE = "01/01/1900" WHERE ' +
                'E_JOURNAL = "' + vCompteOuJournal + '" AND ' +
                'E_GENERAL <> "' + lContrePartie + '" AND ' +
                'E_REFPOINTAGE = "' + vRefPointage + '" AND ' +
                'E_DATEPOINTAGE = "' + UsDateTime(vDatePointage) + '"');
        end
        else
        begin
            // Suppression du pointage des écritures
            ExecuteSQL('UPDATE ECRITURE SET E_REFPOINTAGE = "", ' +
{$IFDEF TRSYNCHRO}
                {FQ Tréso 10019 : maj du champ E_TRESOSYNCHRO, Pour que lors de prochaine synchronisation
                                  de la trésorerie le champ TE_DATERAPPRO puisse être mis à jour}
                'E_TRESOSYNCHRO = "MOD", ' +
{$ENDIF}
                'E_DATEPOINTAGE = "01/01/1900" WHERE ' +
                'E_GENERAL = "' + vCompteOuJournal + '" AND ' +
                'E_REFPOINTAGE = "' + vRefPointage + '" AND ' +
                'E_DATEPOINTAGE = "' + UsDateTime(vDatePointage) + '"');
        end;
    except
        on E : Exception do PgiError('Erreur de requête SQL : ' + E.Message, 'Fonction : CDepointeEcriture');
    end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 19/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}

procedure TOF_POINTAGEMULV7.OnArgument(S : string);
begin
    inherited;
    Ecran.HelpContext := 7604000;
    FStArgument := S;

    TEE_GENERAL := THLabel(GetControl('TEE_GENERAL', True));
    EE_GENERAL := THEdit(GetControl('EE_GENERAL', True));
    EE_DATEPOINTAGE := THEdit(GetControl('EE_DATEPOINTAGE', True));
    EE_REFPOINTAGE := THEdit(GetControl('EE_REFPOINTAGE', True));

    EE_GENERAL.OnElipsisClick := OnElipsisClickEE_GENERAL;
    EE_DATEPOINTAGE.OnElipsisClick := OnElipsisClickEE_DATEPOINTAGE;
    EE_REFPOINTAGE.OnElipsisClick := OnElipsisClickEE_REFPOINTAGE;
    EE_General.OnExit := OnExitEE_General;

    // GCO - 16/01/2004 (Autorise l'accès aux références du pointage AUTO
    if not (ctxPCL in V_PGI.PGIContexte) then
        THEdit(GetControl('XX_WHERE')).Text := '';

{$IFDEF EAGLCLIENT}
    FListe := THGrid(GetControl('FLISTE', True));
{$ELSE}
    FListe := THDBGrid(GetControl('FLISTE', True));
{$ENDIF}
    FListe.OnDblClick := OnDblClickFListe;

    BInsert := TToolBarbutton97(GetControl('BINSERT', True));
    BInsert.OnClick := OnClickBInsert;

    BDelete := TToolBarbutton97(GetControl('BDELETE', True));
    BDelete.OnClick := OnClickBDelete;

    BInfo := TToolBarbutton97(GetControl('BINFO', True));
    BInfo.OnClick := OnClickBInfo;

    BCherche    := TToolBarButton97(GetControl('BCHERCHE', True));
    BOuvrir     := TToolBarButton97(GetControl('BOUVRIR', True));
    BRechercher := TToolBarButton97(GetControl('BRECHERCHER', True));

    BEtatRappro := TToolBarButton97(GetControl('BETATRAPPRO', True));
    BEtatRappro.OnClick := OnClickBEtatRappro;


    POPF11 := TPopUpMenu(GetControl('POPF11', True));
    POPF11.Items[0].OnClick := OnClickBInsert;
    POPF11.Items[1].OnClick := OnClickBDelete;
    POPF11.Items[2].OnClick := OnDblClickFListe;
    POPF11.Items[3].OnClick := OnClickBInfo;

    // a virer quand fiche dans SOCREF
    if PopF11.Items.Count = 5 then
      POPF11.Items[4].OnClick := OnClickBEtatRappro;

    AddMenuPop(PopF11, '', '');

    Ecran.OnKeyDown := OnKeyDownEcran;

    Q := TQuery(GetControl('Q', True));

    FBoConsultation := CEstPointageEnConsultationSurDossier;
    if FBoConsultation then
    begin
        SetControlEnabled('BINSERT', False);
        SetControlEnabled('BDELETE', False);
        SetControlEnabled('BVALIDER', False);
    end;
end;

////////////////////////////////////////////////////////////////////////////////

procedure TOF_POINTAGEMULV7.OnLoad;
var
    lStArg : string;
    F : TFMul;
begin
    inherited;
    if not (Ecran is TFMul) then exit;
    F := TFMul(Ecran);

    if FStArgument <> '' then
    begin
        lStArg := ReadTokenSt(FStArgument);
        if lStarg <> '' then
        begin
            EE_General.Text := ReadTokenSt(lStArg);
            EE_General.Enabled := False;
            Ecran.WindowState := wsMaximized;
        end;
    end;

    if not VH^.PointageJal then
    begin
        Ecran.Caption := 'Pointage sur compte';
        TEE_GENERAL.Caption := TraduireMemoire('Compte général') ;
        EE_GENERAL.MaxLength := VH^.CPta[fbGene].Lg;
    end
    else
    begin
        lStArg := ReadTokenSt(FStArgument);
        if lStarg <> '' then
         EE_REFPOINTAGE.Text := lStArg ;
        Ecran.Caption := 'Pointage sur journal';
        TEE_GENERAL.Caption := TraduireMemoire('Code journal') ;
        EE_GENERAL.MaxLength := 3;
        F.Q.Manuel := True; // Evite l'exécution de la requête lors de la maj de Q.Liste
        F.Q.Liste := 'CPEEXBQ2'; // Affectation de la liste ad-hoc
        F.Q.Manuel := False;
    end;
    Ecran.Caption:=traduireMemoire(Ecran.Caption) ;
    UpDateCaption(Ecran);
end;

////////////////////////////////////////////////////////////////////////////////

procedure TOF_POINTAGEMULV7.OnNew;
begin
    inherited;
end;

////////////////////////////////////////////////////////////////////////////////

procedure TOF_POINTAGEMULV7.OnDelete;
begin
    inherited;
end;

////////////////////////////////////////////////////////////////////////////////

procedure TOF_POINTAGEMULV7.OnUpdate;
begin
    inherited;
end;

////////////////////////////////////////////////////////////////////////////////

procedure TOF_POINTAGEMULV7.OnClose;
begin
    inherited;
end;

////////////////////////////////////////////////////////////////////////////////

procedure TOF_POINTAGEMULV7.OnDisplay();
begin
    inherited;
end;

////////////////////////////////////////////////////////////////////////////////

procedure TOF_POINTAGEMULV7.OnCancel();
begin
    inherited;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 19/03/2003
Modifié le ... :   /  /
Description .. : Suppression d'une ou plusieurs références de pointage
Mots clefs ... :
*****************************************************************}

procedure TOF_POINTAGEMULV7.OnClickBDelete(Sender : TObject);
begin
    // Empeche la suppression si une saisie est ouverte
    //if CEstSaisieOuverte( True ) then Exit;
    try
        if (FListe.AllSelected) or (FListe.nbSelected <> 0) then
        begin
            if PgiAsk('Attention, vous allez supprimer une référence de pointage.' + #13 + #10 +
                'Si des références suivantes ont déjà été créées, le seul moyen pour recréer' + #13 + #10 +
                'cette référence sera de supprimer toutes les références suivantes."' + #13 + #10 +
                'Confirmez vous la suppression des références de pointage sélectionnées ?', Ecran.Caption) = MrYes then
            begin
                if Transactions(SuppressionReference, 1) <> oeOk then
                    MessageAlerte('Traitement annulé. Erreur lors de la suppression des références de pointage.');
            end;
        end
        else
            PGIINFO('Vous n''avez pas sélectionné de référence.', Ecran.Caption);
        FListe.ClearSelected;
    finally
        //DeblocageMonoPoste(True);
        TFMUL(ECRAN).BChercheClick(nil);
    end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 19/03/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}

procedure TOF_POINTAGEMULV7.SuppressionReference;
var
    lCompteOuJournal : string;
    lRefPointage : string;
    lDatePointage : TDatetime;
    i : integer;
    lBoErreur,MsgDone : Boolean;
begin
    lBoErreur := False;
    MsgDone := false;

    try
        if FListe.AllSelected then
        begin // On travaille sur tous les enregistrements de la Query
            Q.First;
            while not Q.Eof do
            begin
                if (GetField('EE_ORIGINERELEVE') = 'INT') then
                begin
                    if (not MsgDone) then PGIBox('Il n''est pas possible de supprimer des référence de pointage automatique.');
                    MsgDone := true;
                end
                else
                begin
                    lCompteOuJournal := GetField('EE_GENERAL');
                    lRefPointage := GetField('EE_REFPOINTAGE');
                    lDatePointage := GetField('EE_DATEPOINTAGE');

                    CSupprimeRefPointageV7(lCompteOuJournal, lRefPointage, lDatePointage);
                    CDepointeEcriture(lCompteOuJournal, lRefPointage, lDatePointage);
                    RecalculTotPointeNew1(lCompteOuJournal);
                end;
                Q.Next;
            end;
        end
        else
        begin // On travaille sur les enregistrements sélectionnés
            for i := 0 to FListe.NbSelected - 1 do
            begin
                FListe.GotoLeBookMark(i);
                if (GetField('EE_ORIGINERELEVE') = 'INT') then
                begin
                    if (not MsgDone) then PGIBox('Il n''est pas possible de supprimer des référence de pointage automatique.');
                    MsgDone := true;
                end
                else
                begin
                    lCompteOuJournal := GetField('EE_GENERAL');
                    lRefPointage := GetField('EE_REFPOINTAGE');
                    lDatePointage := GetField('EE_DATEPOINTAGE');

                    CSupprimeRefPointageV7(lCompteOuJournal, lRefPointage, lDatePointage);
                    CDepointeEcriture(lCompteOuJournal, lRefPointage, lDatePointage);
                    RecalculTotPointeNew1(lCompteOuJournal);
                end;
            end;
        end;
    except
        on E : Exception do lBoErreur := True;
    end;

    if lBoErreur = False then PgiInfo('Le traitement s''est correctement terminé.', 'Suppression des références');
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 18/03/2003
Modifié le ... :   /  /
Description .. : Creation d'une référence de pointage
Mots clefs ... :
*****************************************************************}

procedure TOF_POINTAGEMULV7.OnClickBInsert(Sender : TObject);
{$IFNDEF IMP}
var
    lBoPresence : Boolean;
{$ENDIF}
begin
{$IFNDEF IMP}
    lBoPresence := False;
    if Trim(EE_General.Text) <> '' then
    begin
        if VH^.PointageJal then
        begin
            lBoPresence := Presence('JOURNAL', 'J_JOURNAL', EE_General.Text);
        end
        else
        begin
            if EE_General.Text <> '' then
                EE_General.Text := BourreEtLess(EE_General.Text, fbGene);

            lBoPresence := Presence('GENERAUX', 'G_GENERAL', EE_General.Text);
        end;

        if not lBoPresence then
        begin
            if VH^.PointageJal then
                PgiInfo('Vous ne pouvez pas créer une référence de pointage sur un journal inexistant.', Ecran.Caption)
            else
                PgiInfo('Vous ne pouvez pas créer une référence de pointage sur un compte inexistant.', Ecran.Caption);
            EE_General.SetFocus;
            EE_General.ElipsisClick(Sender);
        end;
    end;

    if (Trim(EE_General.Text) = '') or (lBoPresence) then
    begin
        CPLanceFiche_PointageFic('', 'ACTION=CREATION;' + EE_GENERAL.Text);
        TFMUL(Ecran).BChercheClick(nil);
    end;
{$ENDIF}
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 19/03/2003
Modifié le ... :   /  /
Description .. : Détail de la référence de Pointage
Mots clefs ... :
*****************************************************************}
procedure TOF_POINTAGEMULV7.OnClickBInfo(Sender : TObject);
begin
{$IFDEF EAGLCLIENT}
    TheMulQ := TFMul(Ecran).Q.TQ;
{$ELSE}
    TheMulQ := TFMul(Ecran).Q;
{$ENDIF}

    if TheMulQ.RecordCount <> 0 then
    begin
{$IFNDEF IMP}
        if (GetField('EE_ORIGINERELEVE') = 'INT') then CPLanceFiche_PointageFic(GetField('EE_GENERAL') + ';' + DateToStr(GetField('EE_DATEPOINTAGE')) + ';' + GetField('EE_REFPOINTAGE') + ';' + IntToStr(GetField('EE_NUMERO')), 'ACTION=CONSULTATION')
        else CPLanceFiche_PointageFic(GetField('EE_GENERAL') + ';' + DateToStr(GetField('EE_DATEPOINTAGE')) + ';' + GetField('EE_REFPOINTAGE') + ';' + IntToStr(GetField('EE_NUMERO')), 'ACTION=MODIFICATION');
{$ENDIF}
        TFMUL(ECRAN).BChercheClick(nil);
    end;
end;

{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 26/12/2005
Modifié le ... :   /  /    
Description .. : FQ 13182
Mots clefs ... :
*****************************************************************}
procedure TOF_POINTAGEMULV7.OnClickBEtatRappro(Sender: TObject);
begin
{$IFDEF EAGLCLIENT}
  TheMulQ := TFMul(Ecran).Q.TQ;
{$ELSE}
  TheMulQ := TFMul(Ecran).Q;
{$ENDIF}
  if TheMulQ.RecordCount <> 0 then
  begin
    CC_LanceFicheEtatRapproDetV7(GetField('EE_GENERAL') + ';' +
                               DateToStr(GetField('EE_DATEPOINTAGE'))+ ';X');
    TFMUL(ECRAN).BChercheClick(nil);
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 19/03/2003
Modifié le ... :   /  /
Description .. : Lancement de l'ecran de rapprochement bancaire
Mots clefs ... :
*****************************************************************}

procedure TOF_POINTAGEMULV7.OnDblClickFListe(Sender : TObject);
begin
{$IFDEF EAGLCLIENT}
    TheMulQ := TFMul(Ecran).Q.TQ;
{$ELSE}
    TheMulQ := TFMul(Ecran).Q;
{$ENDIF}

    if TheMulQ.RecordCount <> 0 then
    begin
        if (GetField('EE_ORIGINERELEVE') = 'INT') then
        begin
            PGIBox('Vous ne pouvez pas accéder au pointage manuel sur une référence de pointage automatique.'#13'Reportez vous à la commande Pointage automatique du menu Relevés.');
            CPLanceFiche_Pointage(GetField('EE_GENERAL') + ';' + DateToStr(GetField('EE_DATEPOINTAGE')) + ';' + GetField('EE_REFPOINTAGE') + ';' + IntToStr(GetField('EE_NUMERO')) + ';ACTION=CONSULTATION');
        end
        else CPLanceFiche_Pointage(GetField('EE_GENERAL') + ';' + DateToStr(GetField('EE_DATEPOINTAGE')) + ';' + GetField('EE_REFPOINTAGE') + ';' + IntToStr(GetField('EE_NUMERO')));
        TFMUL(ECRAN).BChercheClick(nil);
    end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 20/03/2003
Modifié le ... :   /  /
Description .. : Affichage des Comptes de Banque Pointable ou des Journaux
Suite ........ : de type Banque ayant leur compte de contrepartie Pointable
Mots clefs ... :
*****************************************************************}

procedure TOF_POINTAGEMULV7.OnElipsisClickEE_General(Sender : TObject);
var
    lSt : string;
begin
    if not VH^.PointageJal then
        LookUpList(THEdit(Sender), 'Compte général', 'GENERAUX', 'G_GENERAL', 'G_LIBELLE',
            'G_POINTABLE="X"', 'G_GENERAL', True, 0)
    else
    begin
        lSt := 'SELECT J_JOURNAL, J_LIBELLE FROM JOURNAL LEFT JOIN GENERAUX ON J_CONTREPARTIE=G_GENERAL ' +
            'WHERE J_NATUREJAL="BQE" AND G_POINTABLE="X"';{ ORDER BY J_JOURNAL' SG6 19/01/05 FQ 15135 }

        LookUpList(THEdit(Sender), 'Journal', 'JOURNAL', 'J_JOURNAL', 'J_LIBELLE', '', 'J_JOURNAL', True, 0, lSt);
    end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 19/05/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}

procedure TOF_POINTAGEMULV7.OnElipsisClickEE_DatePointage(Sender : TObject);
begin
    LookUpList(THEdit(Sender), 'Date de pointage', 'EEXBQ', 'EE_DATEPOINTAGE',
        'EE_REFPOINTAGE', 'EE_GENERAL="' + EE_General.Text + '"', 'EE_DATEPOINTAGE DESC', True, 0);
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 19/05/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}

procedure TOF_POINTAGEMULV7.OnElipsisClickEE_RefPointage(Sender : TObject);
begin
    LookUpList(THEdit(Sender), 'Référence de pointage', 'EEXBQ', 'EE_REFPOINTAGE',
        'EE_DATEPOINTAGE', 'EE_GENERAL="' + EE_General.Text + '"', 'EE_DATEPOINTAGE DESC', True, 0);
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 12/05/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}

procedure TOF_POINTAGEMULV7.OnKeyDownEcran(Sender : TObject; var Key : Word; Shift : TShiftState);
begin
    case key of

        Ord('A') : if Shift = [ssCtrl] then TToolBarButton97(GetControl('BSelectAll')).Click;

        Ord('N') : if Shift = [ssCtrl] then BInsert.Click;

        //Ctrl + H
        70 : if Shift = [ssCtrl] then
                BRechercher.Click;

        VK_DELETE : if Shift = [ssCtrl] then BDelete.Click;

        VK_F5 :
            begin
                if FListe.Focused then BOuvrir.Click;
            end;

        VK_F9 :
            begin
                BCherche.Click;
                if FListe.CanFocus then
                    FListe.SetFocus;
            end;

        VK_F10 : BOuvrir.Click;

        VK_F11 : POPF11.Popup(Mouse.CursorPos.x, Mouse.CursorPos.y);

        VK_F12 : if FListe.Focused then
                TPageControl(GetControl('PAGES', False)).SetFocus
            else
                FListe.SetFocus;

    end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.Privé.*****************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 20/02/2004
Modifié le ... : 20/02/2004
Description .. : Caractère de bourrage en sortie de la zone Général
Mots clefs ... :
*****************************************************************}

procedure TOF_POINTAGEMULV7.OnExitEE_General(Sender : TObject);
var
    lBoPresence : Boolean;
begin
    if Trim(EE_General.Text) = '' then Exit;

    if VH^.PointageJal then
    begin
        lBoPresence := Presence('JOURNAL', 'J_JOURNAL', EE_General.Text);
    end
    else
    begin
        if EE_General.Text <> '' then
            EE_General.Text := BourreEtLess(EE_General.Text, fbGene);

        lBoPresence := Presence('GENERAUX', 'G_GENERAL', EE_General.Text);
    end;

    if not lBoPresence then
    begin
        EE_General.SetFocus;
        EE_General.ElipsisClick(Sender);
    end;
end;

////////////////////////////////////////////////////////////////////////////////




initialization
    registerclasses([TOF_POINTAGEMULV7]);
end.
