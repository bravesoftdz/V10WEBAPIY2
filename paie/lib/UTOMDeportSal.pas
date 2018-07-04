{***********UNITE*************************************************
Auteur  ...... : JL
Créé le ...... : 04/12/2001
Modifié le ... :   /  /
Description .. : Source TOM de la TABLE : DEPORTSAL (DEPORTSAL)
Mots clefs ... : TOM;DEPORTSAL
*****************************************************************}
{
PT1   : 19/03/2002 VG V571 Gestion de la DADSU BTP : Ajout de champs dans la
                           table DEPORTSAL
PT2   : 17/05/2002 JL V582 Onglet MSA : effacement de l'item idem profil ds la
                           combo
PT3   : 21/05/2002 JL V582 Fermeture de la fiche sur suppression du dernier
                           enregistrement
PT4   : 01/07/2002 JL V585 Récupération des responsables sur le OnChangeField au
                           lieu du OnUpdateRecord + Reprise des champs
                           secrétaires de la table service
PT5   : 07/01/2003 JL V_42 Gestion accès formation et éléments variables selon
                           seria
PT6   : 13/03/2003 JL V_42 Fiche qualité 10503 zone PSE_MSAACTIVITE remplit
                           avec ETB mis en commentaire
PT7   : 13/05/2003 MF V_42 Gestion des tickets restaurant
PT8   : 09/10/2003 JL V_42 Passage en mode monofiche en création, car erreur
                           objet non trouver à cause requête.
PT9   : 09/10/2003 JL V_42 FQ : 10823 Initialisation activité idem établissement
                           + 20/10/03 modif requête pour recherche activité
PT10  : 26/03/2004 JL V_50 Getsion MSA EDI : nouveau champ PSE_MSATECHNIQUE
PT11  : 23/04/2004 JL V_50 Gestion MSA EDIT : La zone lieu de travail doit
                           contenir le code postal
                           -> ajout elipsis et gestion dans script, et affichage
                           libellé ville dans THLabel
PT12  : 17/05/2004 SB V_50 FQ 11288 Gestion acces zone resp. en fonction de la
                           gestion service
PT13  : 25/05/2004 VG V_50 Ajout de champs pour la DADS-U BTP - FQ N°11262
PT14  : 05/07/2004 VG V_50 Adaptation cahier des charges DADS-U V8R00
PT15  : 03/12/2004 JL V_60 Correction elipsisclick sur service en EAGL + affichage libellé responsables
PT16  : 22/04/2005 PH V_60 Filtrage salarié confidentialité établissement
PT17  : 19/05/2004 JL V_60 FQ 12283 Code act obligatoire msa
PT18  : 29/06/2005 JL V_60 Gestion responsable
PT19  : 04/04/2006 JL V_65 Ajout gestion UG MSA
PT20  : 09/06/2006 JL V_65 FQ 13216 pse_isnumident initialisé avec numéro SS
PT21  : 22/01/2007 MF V_720 Titres restaurant (saisie du code point de livraison ACCOR)
PT22  : 19/02/2007 MF V_720 Affichage du code point de tristribution ACCOR
PT23  : 31/05/2007 V_72 JL FQ 14021 Ajout apramsoc pour gestion intervenants exterieurs
PT24  : 20/12/2007 FC V810 Concepts accessibilité depuis la fiche salarié
PT25  : 13/02/2008 FL V_803 En multidossier, utilisation de la tablette SALARIEINT
PT27  : 21/02/2008 FL V_803 Si changement de service, mettre à jour les tables dépendantes
PT27  : 02/04/2008 FL V_803 Adaptation pour partage de la hiérarchie : restreindre l'affectation d'un service au même dossier
PT28  : 29/05/2008 NA V_850 Ticket restaurant : affichage du libellé "code succursale" au lieu de code de distribution si chq déjeuner
}
unit UTOMDeportSal;

interface

uses Controls, Classes,  forms, sysutils, ComCtrls, HTB97,
{$IFNDEF EAGLCLIENT}
  HDB,  db, {$IFNDEF DBXPRESS} dbTables, {$ELSE} uDbxDataSet, {$ENDIF} Fiche,
{$ELSE}
  MaineAGL, eFiche,
{$ENDIF}
   HCtrls, HEnt1, EntPaie, HMsgBox,  UTOB, UTOM,
  LookUp, ParamSoc;

type
  TOM_DEPORTSAL = class(TOM)
    procedure OnLoadRecord; override;
    procedure OnNewRecord; override;
    procedure OnChangeField(F: TField); override;
    procedure OnUpdateRecord; override;
    procedure OnAfterUpdateRecord; override;
    procedure OnArgument(S: string); override;
    procedure OnClose; override;
  private
    BSupp, BDelete: TToolBarButton97;
    ServiceInit, NomPage, ChampCle, Titre, Sal, Etat, Arg, TestTrim, TestAnnee, TestSalarie, Action, TAction: string;
    //    RespAbs:THLabel;
    P: TPageControl;
    BMsa, BIntermittents, BOrg, BBTP, BTckRestau, GestionUnique: Boolean; // PT7
    Retour: Boolean;
    Bouge, Requete: string;
    QDeportsal: TQuery;
    fournisseur:         string; // PT21
    
    procedure AfficheEcran(Sender: TObject);
    procedure Suppression(Sender: TObject);
    function TestExistence(ChampTest: string): Boolean;
    procedure premier(Sender: TObject);
    procedure dernier(Sender: TObject);
    procedure suivant(Sender: TObject);
    procedure precedent(Sender: TObject);
    procedure ServiceElipsisClick(Sender: TObject);
    procedure MetierBTPClick(Sender: TObject);
// d PT21
{$IFDEF EAGLCLIENT}
   procedure SaisieCodeLivr(Sender : TObject);
{$ENDIF}   
// f PT21
  end;

implementation

Uses PGOutilsFormation;

procedure TOM_DEPORTSAL.OnClose;
begin
  inherited;
  //If QDeportSal <> Nil then Ferme(QDeportSal);
end;

procedure TOM_DEPORTSAL.OnLoadRecord;
var
  TestBtp, TestInt, TestMsa, TestResp, TestExiste, TestTckRestau: Boolean; // PT7
  CodeBur, InfosCompl, swhere: string;
  Q: TQuery;
  PointLivr, TypPointLivr : string; // PT21
begin
  inherited;
  if Bouge = 'PREMIER' then
  begin
    TFFiche(Ecran).BFirst.Enabled := False;
    TFFiche(Ecran).BPrev.Enabled := False;
  end;
  if Bouge = 'DERNIER' then
  begin
    TFFiche(Ecran).BLast.Enabled := False;
    TFFiche(Ecran).BNext.Enabled := False;
  end;
  if (Bouge = 'PRECEDENT') and (Retour = True) then
  begin
    TFFiche(Ecran).BFirst.Enabled := False;
    TFFiche(Ecran).BPrev.Enabled := False;
  end;
  if (Bouge = 'SUIVANT') and (Retour = True) then
  begin
    TFFiche(Ecran).BLast.Enabled := False;
    TFFiche(Ecran).BNext.Enabled := False;
  end;
  if (Bouge = 'SUIVANT') or (Bouge = 'PREMIER') then
  begin
    QDeportSal.First;
    if QDeportSal.FindField('PSE_SALARIE').AsString = '' then // DEBUT PT3
    begin
      TFFiche(Ecran).Close;
      Exit;
    end; // FIN PT3
    while QDeportSal.FindField('PSE_SALARIE').AsString <> GetField('PSE_SALARIE') do
    begin
      if QDeportSal.Eof then
      begin
        if Retour = False then TFFiche(Ecran).QFiche.Next
        else TFFiche(Ecran).QFiche.Prior;
        if TFFiche(Ecran).QFiche.Eof then
        begin
          Retour := True;
          TFFiche(Ecran).BNext.Enabled := False;
          TFFiche(Ecran).BLast.Enabled := False;
          TFFiche(Ecran).QFiche.prior;
        end;
        if (Retour = True) and (TFFiche(Ecran).QFiche.Bof) then TFFiche(Ecran).Close;
      end;
      if QDeportSal.FindField('PSE_SALARIE').AsString <> GetField('PSE_SALARIE') then QDeportSal.Next;
    end;
  end;
  if (Bouge = 'PRECEDENT') or (Bouge = 'DERNIER') then
  begin
    QDeportSal.First;
    while QDeportSal.FindField('PSE_SALARIE').AsString <> GetField('PSE_SALARIE') do
    begin
      if QDeportSal.Eof then
      begin
        if Retour = False then TFFiche(Ecran).QFiche.Prior
        else TFFiche(Ecran).QFiche.Next;
        if TFFiche(Ecran).QFiche.Bof then
        begin
          TFFiche(Ecran).BFirst.Enabled := False;
          TFFiche(Ecran).BPrev.Enabled := False;
          Retour := True;
          TFFiche(Ecran).QFiche.Next;
        end;
        if (Retour = True) and (TFFiche(Ecran).QFiche.Eof) then TFFiche(Ecran).Close;
      end;
      if QDeportSal.FindField('PSE_SALARIE').AsString <> GetField('PSE_SALARIE') then QDeportSal.Next;
    end;
  end;
  if GestionUnique = False then // Si plusieurs cas sont gérés, alors on vérifie si l'enregistrement du salarié existe déja dans la base
  begin
    if Action = 'MODIFICATION' then SetControlEnabled(ChampCle, False);
    TestExiste := False; // initialisation des booléens à False
    TestInt := TestExistence('PSE_INTERMITTENT');
    TestMsa := TestExistence('PSE_MSA');
    TestResp := TestExistence('PSE_ORGANIGRAMME');
    TestBtp := TestExistence('PSE_BTP');
    TestTckRestau := TestExistence('PSE_TICKETREST'); // PT7
    if (Arg = 'RESP') and ((TestMsa = True) or (TestInt = True) or (TestBtp = True) or (TestTckRestau = True)) then TestExiste := True; // PT7
    if Arg = 'MSA' then
    begin
      if (TestInt = True) or (TestResp = True) or (TestBtp = True) or (TestTckRestau = True) then TestExiste := True; // PT7
    end;
    if (Arg = 'IS') and ((TestMsa = True) or (TestResp = True) or (TestBtp = True) or (TestTckRestau = True)) then TestExiste := True; // PT7
    if (Arg = 'BTP') and ((TestMsa = True) or (TestResp = True) or (TestInt = True) or (TestTckRestau = True)) then TestExiste := True; //@à
    if TestExiste = True then // Choix du bouton de suppression en fonction des tests d'existence de l'enregistrement
    begin
      BDelete.Enabled := False;
      BDelete.Visible := False;
      //DEB PT24
      if Action='CONSULTATION' then
      begin
        BSupp.Enabled := False;
        BSupp.Visible := False;
      end
      //FIN PT24
      else
      begin
        BSupp.Enabled := True;
        BSupp.Visible := True;
      end;
    end
    else
    begin
      //DEB PT24
      if Action='CONSULTATION' then
      begin
        BDelete.Enabled := False;
        BDelete.Visible := False;
      end
      //FIN PT24
      else
      begin
        BDelete.Enabled := True;
        BDelete.Visible := True;
      end;
      BSupp.Enabled := False;
      BSupp.Visible := False;
    end;
    if (Etat = 'NOUVEAU') or (TFFiche(Ecran).FTypeAction = taCreat) then
    begin
      ForceUpdate; // On force la passage en mode modification pour que le salmarié soit affecté
      BSupp.Enabled := False; // Suppression impossible car mode création
      BDelete.Enabled := False;
    end;
  end
  else // Si le salarié n'est géré que dans un cas
  begin
    if Action = 'MODIFICATION' then SetControlEnabled(ChampCle, False);
    SetControlVisible(ChampCle, True);
  end;
  Ecran.caption := Titre + GetField('PSE_SALARIE') + ' ' + RechDom('PGSALARIE', GetField('PSE_SALARIE'), FALSE);
  TFFiche(Ecran).DisabledMajCaption := True;
  SetActiveTabSheet(NomPage);
  //DEBUT PT20
  If Arg = 'IS' then
  begin
       if (Etat = 'NOUVEAU') or (action = 'CREATION') then
       begin
            Q := OpenSQL('SELECT PSA_NUMEROSS FROM SALARIES WHERE PSA_SALARIE="'+Sal+'"',True);
            If Not Q.Eof then
              if (Action<>'CONSULTATION') then //PT24
                SetField('PSE_ISNUMIDENT',Q.FindField('PSA_NUMEROSS').AsString);
            Ferme(Q);
       end;
  end;
  //FIN PT20
  if Arg = 'MSA' then
  begin
    if (GetField('PSE_MSATYPEACT') = '') AND (Action<>'CONSULTATION') then     //PT24
      SetField('PSE_MSATYPEACT', 'ETB');
    if (GetField('PSE_MSATYPUNITEG') = '') AND (Action<>'CONSULTATION') then     //PT24
      SetField('PSE_MSATYPUNITEG', 'ETB');
    if (Etat = 'NOUVEAU') or ((action = 'CREATION') AND (action<>'CONSULTATION')) then  //PT24
    begin
      SetField('PSE_MSATYPEACT', 'ETB'); // PT9
      Q := OpenSQL('SELECT PAT_CODEBUREAU FROM TAUXAT LEFT JOIN SALARIES ON PAT_ORDREAT=PSA_ORDREAT AND' +
        ' PAT_ETABLISSEMENT=PSA_ETABLISSEMENT WHERE PSA_SALARIE="' + Sal + '"', True);
      CodeBur := '';
      if not Q.eof then CodeBur := Q.FindField('PAT_CODEBUREAU').AsString; // PortageCWAS
      if CodeBur = 'BUR' then InfosCompl := 'ADMINISTRATIF'
      else InfosCompl := 'TECHNIQUE';
      Ferme(Q);
      SetField('PSE_MSAINFOSCOMPL', InfosCompl);
    end;
  end;
  if (Etat = 'NOUVEAU') AND (Action<>'CONSULTATION') then //PT24
  begin
    if Arg = 'RESP' then SetField('PSE_ORGANIGRAMME', 'X');
    if Arg = 'IS' then SetField('PSE_INTERMITTENT', 'X');
    if Arg = 'MSA' then SetField('PSE_MSA', 'X');
    if Arg = 'BTP' then SetField('PSE_BTP', 'X');
    if Arg = 'TCK' then
      SetField('PSE_TICKETREST', 'X'); // PT7
  end;
  if Arg = 'RESP' then ServiceInit := GetField('PSE_CODESERVICE');
  // DEBUT PT5
  if not VH_Paie.PgeAbsences then
  begin
    SetControlVisible('PSE_RESPONSVAR', False);
    SetControlVisible('LIBRESPVAR', False);
    SetControlVisible('TPSE_RESPONSVAR', False);
  end
  else
  begin
    SetControlVisible('PSE_RESPONSVAR', True);
    SetControlVisible('LIBRESPVAR', True);
    SetControlVisible('TPSE_RESPONSVAR', True);
  end;
  if not VH_Paie.PgSeriaFormation then
  begin
    SetControlVisible('PSE_RESPONSFOR', False);
    SetControlVisible('TPSE_RESPONSFOR', False);
  end
  else
  begin
    SetControlVisible('PSE_RESPONSFOR', True);
    SetControlVisible('TPSE_RESPONSFOR', True);
  end;
  // FIN PT5
  // d PT7
  if (Arg = 'TCK') then
  begin
// d PT21
  if ((VH_Paie.PGTicketRestau) and (fournisseur = '003')) then
     
    // ACCOR
    begin
      if (V_PGI.NumVersionBase < 800) then
      // (spécif V7)
      begin
        PointLivr:= GetField('PSE_MSAUNITEGES');
        TypPointLivr := GetField('PSE_MSATYPUNITEG');
{$IFDEF EAGLCLIENT}
        SetControlText ('PSE_MSAUNITEGES1',PointLivr);
        SetField ('PSE_MSAUNITEGES',PointLivr);
        SetControlText ('PSE_MSATYPUNITEG1',TypPointLivr);
        SetField ('PSE_MSATYPUNITEG',TypPointLivr);
{$ENDIF}
      end
      else
      // (à partir V8)
      begin
        PointLivr := GetField('ETB_TICKLIVR');

      end;
    end;
// f PT21

    {Type de ticket}
    if (GetField('PSE_TYPETYPTK') = '') AND (Action<>'CONSULTATION') then //PT24
      Setfield('PSE_TYPETYPTK', 'ETB');
    {Code de distribution}
    if (GetField('PSE_DISTRIBUTION') = '') then
    begin
      SetFocusControl('PSE_DISTRIBUTION'); {Champ obligatoire}
      if not (ds.state in [dsinsert, dsedit]) then ds.edit;
    end
// d PT22
    else
    begin
      Q := opensql('SELECT * FROM CHOIXCOD WHERE CC_TYPE="PCD" AND CC_ABREGE="'+GetField('PSE_DISTRIBUTION')+'"',TRUE);
      if not Q.EOF then
      begin
        SetControlText('DISTRIBACCOR', Q.findfield('CC_LIBRE').asstring);
      end;
      ferme (Q);
    end;
// f PT22
    {Personnalisation}
    if GetParamSoc('SO_PGPERSOTICKET') then
    begin
      Q := OpenSQL('SELECT PSA_LIBELLE, PSA_PRENOM FROM SALARIES ' +
        'WHERE PSA_SALARIE = "' + Sal + '"', True);
      if not Q.eof then
        if (Action<>'CONSULTATION') then //PT24
        begin
          SetField('PSE_PERSONNAL',
            Trim(Q.FindField('PSA_LIBELLE').AsString) + ' ' +
            Trim(Q.FindField('PSA_PRENOM').AsString));
          SetControlEnabled('PSE_PERSONNAL', True);
        end;
      Ferme(Q);
      ;
    end
    else
    begin
      if (Action<>'CONSULTATION') then //PT24
      begin
        SetField('PSE_PERSONNAL', '');
        SetControlEnabled('PSE_PERSONNAL', False);
      end;
    end;
  end;
  // f PT7

  //PT14
  if (Arg = 'BTP') then
  begin
    sWhere := 'SELECT PMB_METIERLIB' +
      ' FROM METIERBTP WHERE' +
      ' ##PMB_PREDEFINI##' +
      ' PMB_METIERBTP = "' + GetControlText('PSE_METIERBTP') + '"';
    Q := OpenSql(sWhere, TRUE);
    if (not Q.EOF) then
      SetControlText('LMETIERBTP_', Q.Fields[0].asstring)
    else
      SetControlText('LMETIERBTP_', '');
    Ferme(Q);
  end;
  //FIN PT14

  //Debut PT11
  Q := OpenSQL('Select O_VILLE from CODEPOST WHERE O_CODEPOSTAL="' + GetField('PSE_MSALIEUTRAV') + '"', True);
  if not Q.Eof then SetControlCaption('LIBVILLEMSA', Q.FindField('O_VILLE').AsString)
  else SetControlCaption('LIBVILLEMSA', '');
  Ferme(Q);
  //FIN PT11
  { DEB PT12 }
  SetControlEnabled('PSE_RESPONSABS', (not VH_Paie.PGResponsables));
  SetControlEnabled('PSE_RESPONSVAR', (not VH_Paie.PGResponsables));
  SetControlEnabled('PSE_RESPSERVICE', (not VH_Paie.PGResponsables));
  SetControlEnabled('PSE_RESPONSNDF', (not VH_Paie.PGResponsables));
  SetControlEnabled('PSE_RESPONSFOR', (not VH_Paie.PGResponsables));
  { FIN PT12 }
end;

procedure TOM_DEPORTSAL.OnUpdateRecord;
var
  Mess: string; // PT7
  PointLivr     : string;              //   PT21
begin
  inherited;
  //PT4 MAJ dans on change field, suppression du code le 10/02/03
  // d PT7
  Mess := '';
  if (Arg = 'TCK') and (GetField('PSE_DISTRIBUTION') = '') then
  begin
    LastError := 1;
    Mess := Mess + ' Code de distribution ';
    SetFocusControl('PSE_DISTRIBUTION');
  end;
  if LastError <> 0 then
  begin
    PGIBox('Vous devez renseigner :' + Mess, 'Information obligatoire');
  end;
// d PT21
  if (Arg = 'TCK') then
  // Titres restaurant
  begin
 if ((VH_Paie.PGTicketRestau) and (fournisseur = '003')) then

    // ACCOR   
    begin
      if (V_PGI.NumVersionBase < 800) then
      begin
//        PointLivr := GetField('PSE_MSAUNITEGES');
        PointLivr := GetControlText('PSE_MSAUNITEGES1');
      end
      else
      begin
        PointLivr := GetField('PSE_TICKLIVR');
      end;

        if ((not IsNumeric(PointLivr)) or (length(PointLivr) <> 6)) then
        begin
          LastError := 1;
          PgiBox ('Le code point de livraison doit être numérique et#13#10'+
                  'd''une longueur de 6 caractères ', TFFiche(Ecran).Caption);
          if (V_PGI.NumVersionBase < 800) then
          begin
            SetFocusControl('PSE_MSAUNITEGES1');
          end
          else
          begin
            SetFocusControl('PSE_TICKLIVR');
          end;
          exit;
        end;



    end;
  end;
// f PT21
  // f PT7
  //DEBUT PT17
  if Arg = 'MSA' then
  begin
       If GetField('PSE_MSAACTIVITE') = '' then
       begin
            PGIBox('Vous devez renseigner l''activité', 'Information obligatoire');
            LastError := 1;
            SetFocusControl('PSE_MSAACTIVITE');
       end;
  end;
  //FIN PT17
end;

procedure TOM_DEPORTSAL.OnAfterUpdateRecord;
begin
  inherited;

    //PT27 - Début
    If (Action = 'MODIFICATION') And (ServiceInit <> GetField('PSE_CODESERVICE')) Then
    Begin
        MAJResponsable(GetField('PSE_CODESERVICE'), 'VAR', GetField('PSE_RESPONSVAR'));
        MAJResponsable(GetField('PSE_CODESERVICE'), 'FOR', GetField('PSE_RESPONSFOR'));
    End;
    //PT27 - Fin

  if GestionUnique = False then
  begin
    if (Etat = 'NOUVEAU') or (TFFiche(Ecran).FTypeAction = taCreat) then
    begin
      BSupp.Enabled := True;
      BDelete.Enabled := True;
    end;
  end;
end;

procedure TOM_DEPORTSAL.OnNewRecord;
// d PT7
var
  Q: TQuery;
  Etab: string;
  // f PT7
begin
  inherited;
  SetField('PSE_MSATECHNIQUE', '-'); //PT10
  if Sal <> '' then SetField('PSE_SALARIE', Sal); // Récupération du code salarié (PSA_SALARIE)
  SetField('PSE_ASSISTVAR', '-');
  SetField('PSE_OKVALIDABS', '-');
  SetField('PSE_OKVALIDNDF', '-');
  SetField('PSE_OKVALIDVAR', '-');
  SetField('PSE_VEHICULESOC', '-');
  SetField('PSE_TICKETREST', '-');
  SetField('PSE_ADMINISTABS', '-');
  if Arg = 'IS' then
  begin
    SetField('PSE_INTERMITTENT', 'X');
    SetField('PSE_MSA', '-');
    SetField('PSE_ORGANIGRAMME', '-');
    SetField('PSE_BTP', '-');
    SetField('PSE_TICKETREST', '-'); // PT7
  end;
  if (Arg = 'MSA') then
  begin
    SetField('PSE_MSATYPEACT', 'ETB');
    SetField('PSE_MSA', 'X');
    SetField('PSE_ORGANIGRAMME', '-');
    SetField('PSE_INTERMITTENT', '-');
    SetField('PSE_BTP', '-');
    SetField('PSE_TICKETREST', '-'); // PT7
  end;
  if Arg = 'RESP' then
  begin
    SetField('PSE_ORGANIGRAMME', 'X');
    SetField('PSE_BTP', '-');
    SetField('PSE_INTERMITTENT', '-');
    SetField('PSE_MSA', '-');
    SetField('PSE_TICKETREST', '-'); // PT7
  end;
  if (Arg = 'BTP') then
  begin
    SetField('PSE_BTP', 'X');
    SetField('PSE_ORGANIGRAMME', '-');
    SetField('PSE_INTERMITTENT', '-');
    SetField('PSE_MSA', '-');
    SetField('PSE_TICKETREST', '-'); // PT7
  end;
  // d PT7
  if (Arg = 'TCK') then
  begin
    SetField('PSE_TYPETYPTK', 'ETB');
    SetField('PSE_BTP', '-');
    SetField('PSE_ORGANIGRAMME', '-');
    SetField('PSE_INTERMITTENT', '-');
    SetField('PSE_MSA', '-');
    SetField('PSE_TICKETREST', 'X');
// d PT21
    if (V_PGI.NumVersionBase < 800) then
    begin
      SetField('PSE_MSATYPUNITEG', 'ETB');
      SetControlText('PSE_MSATYPUNITEG1', 'ETB');
    end
    else
      SetField('PSE_TYPTICKLIVR', 'ETB');
// f PT21

    Q := OpenSQL('SELECT ETB_ETABLISSEMENT,PSA_SALARIE FROM ETABCOMPL ' +
      'LEFT JOIN SALARIES ON ETB_ETABLISSEMENT=PSA_ETABLISSEMENT ' +
      'WHERE PSA_SALARIE="' + Sal + '"', True);
    if not Q.eof then
      Etab := Q.FindField('ETB_ETABLISSEMENT').AsString;
    Ferme(Q);
// d PT21
    if (V_PGI.NumVersionBase < 800) then
      Q := OpenSQL('SELECT ETB_TYPTICKET, ETB_AUTRENUMERO FROM ETABCOMPL ' +
                   'WHERE ETB_ETABLISSEMENT = "' + Etab + '"', True)
    else
      Q := OpenSQL('SELECT ETB_TYPTICKET, ETB_TICKLIVR FROM ETABCOMPL ' +
                   'WHERE ETB_ETABLISSEMENT = "' + Etab + '"', True);

    if not Q.eof then
    begin
      SetField('PSE_TYPTICKET', Q.FindField('ETB_TYPTICKET').AsString);
      if (V_PGI.NumVersionBase < 800) then
      begin
        SetField('PSE_MSAUNITEGES', Q.FindField('ETB_AUTRENUMERO').AsString);
        SetControlText('PSE_MSAUNITEGES1',Q.FindField('ETB_AUTRENUMERO').AsString);
      end
      else
        SetField('PSE_TICKLIVR', Q.FindField('ETB_TICKLIVR').AsString);
    end;
    Ferme(Q);
    
    SetControlEnabled('PSE_TYPTICKET', False);
    if (V_PGI.NumVersionBase < 800) then
      SetControlEnabled('PSE_MSAUNITEGES1', False)
    else
      SetControlEnabled('PSE_TICKLIVR', False);   
// f PT21

    if GetParamSoc('SO_PGPERSOTICKET') then
    begin
      Q := OpenSQL('SELECT PSA_LIBELLE, PSA_PRENOM FROM SALARIES ' +
        'WHERE PSA_SALARIE = "' + Sal + '"', True);
      if not Q.eof then
        SetField('PSE_PERSONNAL',
          Trim(Q.FindField('PSA_LIBELLE').AsString) + ' ' +
          Trim(Q.FindField('PSA_PRENOM').AsString));
      SetControlEnabled('PSE_PERSONNAL', True);
      Ferme(Q);
    end
    else
    begin
      SetField('PSE_PERSONNAL', '');
      SetControlEnabled('PSE_PERSONNAL', False);
    end;
  end;
  // f PT7
end;

procedure TOM_DEPORTSAL.OnChangeField(F: TField);
var
  Q: TQuery;
  Sal: string;
  PointLivr   : string; // PT21
begin
  inherited;
  if Arg = 'RESP' then
  begin
    //PT4
  //  if (F.FieldName='PSE_RESPONSABS') then RespAbs.Caption:=RechDom ('PGSALARIE', GetField ('PSE_RESPONSABS'),FALSE); Mise en commentaire le 01/07/2002
    if (F.FieldName = 'PSE_CODESERVICE') then
    begin
      if (GetField('PSE_CODESERVICE') <> '') and (GetField('PSE_CODESERVICE') <> ServiceInit) then
      begin
        ForceUpdate;
        Q := OpenSQL('SELECT PGS_SECRETAIREABS,PGS_SECRETAIRENDF,PGS_SECRETAIREVAR,PGS_RESPONSABS,PGS_RESPONSNDF,PGS_RESPONSVAR,PGS_RESPONSFOR,PGS_RESPSERVICE' +
          ' FROM SERVICES WHERE PGS_CODESERVICE="' + GetField('PSE_CODESERVICE') + '"', True);
        if not Q.eof then // PortageCWAS
        begin
          //DEB PT24
          if (Action<>'CONSULTATION') then
          begin
            SetField('PSE_ASSISTABS', Q.FindField('PGS_SECRETAIREABS').AsString);
            SetField('PSE_ASSISTVAR', Q.FindField('PGS_SECRETAIREVAR').AsString);
            SetField('PSE_ASSISTNDF', Q.FindField('PGS_SECRETAIRENDF').AsString);
            SetField('PSE_RESPONSABS', Q.FindField('PGS_RESPONSABS').AsString);
            SetField('PSE_RESPSERVICE', Q.FindField('PGS_RESPSERVICE').AsString);
            SetField('PSE_RESPONSNDF', Q.FindField('PGS_RESPONSNDF').AsString);
            SetField('PSE_RESPONSVAR', Q.FindField('PGS_RESPONSVAR').AsString);
            SetField('PSE_RESPONSFOR', Q.FindField('PGS_RESPONSFOR').AsString);
          end;
          //FIN PT24
          // DEBUT PT15
{$IFDEF EAGLCLIENT}
          if Q.FindField('PGS_RESPONSABS').AsString <> '' then SetControlText('LBLNOMRESP', RechDom('PGSALARIE', Q.FindField('PGS_RESPONSABS').AsString, False));
          if Q.FindField('PGS_RESPONSABS').AsString <> '' then SetControlText('LIBRESPVAR', RechDom('PGSALARIE', Q.FindField('PGS_RESPONSABS').AsString, False));
{$ENDIF}
          // FIN PT15
        end;
        Ferme(Q);
      end;
      if GetField('PSE_CODESERVICE') = '' then
      begin
        //DEB PT24
        if (Action<>'CONSULTATION') then
        begin
          SetField('PSE_ASSISTABS', '');
          SetField('PSE_ASSISTVAR', '');
          SetField('PSE_ASSISTNDF', '');
          SetField('PSE_RESPONSABS', '');
          SetField('PSE_RESPSERVICE', '');
          SetField('PSE_RESPONSNDF', '');
          SetField('PSE_RESPONSVAR', '');
          SetField('PSE_RESPONSFOR', '');
        end;
        //FIN PT24
      end;
    end;
    //FIN PT4
  end;
  if arg = 'MSA' then
  begin
    if (F.FieldName = 'PSE_MSATYPEACT') then
    begin
      if GetField('PSE_MSATYPEACT') = 'ETB' then
      begin
        Sal := GetField('PSE_SALARIE');
        Q := OpenSQL('SELECT ETB_MSAACTIVITE FROM ETABCOMPL LEFT JOIN SALARIES ON ETB_ETABLISSEMENT=PSA_ETABLISSEMENT WHERE PSA_SALARIE="' + Sal + '"', True); //PT9
        if not Q.eof then
        begin
          if GetField('PSE_MSAACTIVITE') <> (Q.FindField('ETB_MSAACTIVITE').AsString) then
            if (Action<>'CONSULTATION') then //PT24
              SetField('PSE_MSAACTIVITE', Q.FindField('ETB_MSAACTIVITE').AsString); // PortageCWAS
        end;
        Ferme(Q);
        SetControlEnabled('PSE_MSAACTIVITE', False);
      end
      else SetControlEnabled('PSE_MSAACTIVITE', True);
    end;
    //DEBUT PT19
    if (F.FieldName = 'PSE_MSATYPUNITEG') then
    begin
      if GetField('PSE_MSATYPUNITEG') = 'ETB' then
      begin
        Sal := GetField('PSE_SALARIE');
        Q := OpenSQL('SELECT ETB_MSAUNITEGES FROM ETABCOMPL LEFT JOIN SALARIES ON ETB_ETABLISSEMENT=PSA_ETABLISSEMENT WHERE PSA_SALARIE="' + Sal + '"', True); //PT9
        if not Q.eof then
        begin
          if GetField('PSE_MSAUNITEGES') <> (Q.FindField('ETB_MSAUNITEGES').AsString) then
            if (Action<>'CONSULTATION') then //PT24
              SetField('PSE_MSAUNITEGES', Q.FindField('ETB_MSAUNITEGES').AsString);
        end;
        Ferme(Q);
        SetControlEnabled('PSE_MSAUNITEGES', False);
      end
      else SetControlEnabled('PSE_MSAUNITEGES', True);
    end;
    //FIN PT19
  end;
  // d PT7
  if arg = 'TCK' then
  begin
    if (F.FieldName = 'PSE_TYPETYPTK') then
    begin
      if (GetField('PSE_TYPETYPTK') = 'ETB') then
      begin
        Sal := GetField('PSE_SALARIE');
        Q := OpenSQL('SELECT ETB_TYPTICKET FROM ETABCOMPL ' +
          'LEFT JOIN SALARIES ON ETB_ETABLISSEMENT=PSA_ETABLISSEMENT' +
          ' LEFT JOIN DEPORTSAL ON PSE_SALARIE=PSA_SALARIE ' +
          'WHERE PSE_SALARIE="' + Sal + '"', True);
        if not Q.eof then
          if (Action<>'CONSULTATION') then //PT24
            SetField('PSE_TYPTICKET', Q.FindField('ETB_TYPTICKET').AsString); // PortageCWAS
        Ferme(Q);
        SetControlEnabled('PSE_TYPTICKET', False);
      end
      else
        SetControlEnabled('PSE_TYPTICKET', True);
      SetFocusControl('PSE_TYPTICKET');
    end;
// d PT22
    if (F.FieldName = 'PSE_DISTRIBUTION') then
    begin
      Q := opensql('SELECT * FROM CHOIXCOD WHERE CC_TYPE="PCD" AND CC_ABREGE="'+GetField('PSE_DISTRIBUTION')+'"',TRUE);
      if not Q.EOF then
      begin
        SetControlText('DISTRIBACCOR', Q.findfield('CC_LIBRE').asstring);
      end;
      ferme (Q);
    end;
// f PT22
// d PT21
    if (((F.FieldName = ('PSE_MSATYPUNITEG')) or (F.FieldName = ('PSE_TYPTICKLIVR'))) and
      (VH_Paie.PGTicketRestau) and (fournisseur = '003')) then

    begin
      if (V_PGI.NumVersionBase < 800) then
      begin
        if (GetField('PSE_MSATYPUNITEG') = 'ETB') then
        begin
          Sal := GetField('PSE_SALARIE');
          Q := OpenSQL('SELECT ETB_AUTRENUMERO FROM ETABCOMPL ' +
                       'LEFT JOIN SALARIES ON ETB_ETABLISSEMENT=PSA_ETABLISSEMENT' +
                       ' LEFT JOIN DEPORTSAL ON PSE_SALARIE=PSA_SALARIE ' +
                       'WHERE PSE_SALARIE="' + Sal + '"', True);
          if not Q.eof then
          begin
            SetField('PSE_MSAUNITEGES', Q.FindField('ETB_AUTRENUMERO').AsString);
{$IFDEF EAGLCLIENT}
            SetControlText('PSE_MSAUNITEGES1', Q.FindField('ETB_AUTRENUMERO').AsString);
{$ENDIF}
          end;
          Ferme(Q);
          SetControlEnabled('PSE_MSAUNITEGES1', False);
        end
        else
          SetControlEnabled('PSE_MSAUNITEGES1', True);

        SetFocusControl('PSE_MSAUNITEGES1');
      end
      else
      begin
        if (GetField('PSE_TYPTICKLIVR') = 'ETB') then      // PT21
        begin
          Sal := GetField('PSE_SALARIE');
          Q := OpenSQL('SELECT ETB_TICKLIVR FROM ETABCOMPL ' +
                       'LEFT JOIN SALARIES ON ETB_ETABLISSEMENT=PSA_ETABLISSEMENT' +
                       ' LEFT JOIN DEPORTSAL ON PSE_SALARIE=PSA_SALARIE ' +
                       'WHERE PSE_SALARIE="' + Sal + '"', True);
          if not Q.eof then
            if (Action<>'CONSULTATION') then //PT24
              SetField('PSE_TICKLIVR', Q.FindField('ETB_TICKLIVR').AsString);
          Ferme(Q);
          SetControlEnabled('PSE_TICKLIVR', False);
        end
        else
        SetControlEnabled('PSE_TICKLIVR', True);

      SetFocusControl('PSE_TICKLIVR');
      end;

    end;
    if (((F.FieldName = ('PSE_MSAUNITEGES')) or (F.FieldName = ('PSE_TICKLIVR'))) and
       (VH_Paie.PGTicketRestau) and (fournisseur = '003')) then

    begin
     if (V_PGI.NumVersionBase < 800) then
     begin
       PointLivr := GetField('PSE_MSAUNITEGES');
     end
     else
     begin
       PointLivr := GetField('PSE_TICKLIVR');
     end;



        if ((not IsNumeric(PointLivr)) or (length(PointLivr) <> 6)) then
        begin
          PgiBox ('Le code point de livraison doit être numérique et#13#10'+
                  'd''une longueur de 6 caractères ', TFFiche(Ecran).Caption);
          if (V_PGI.NumVersionBase < 800) then
          begin
            SetFocusControl('PSE_MSAUNITEGES1');
          end
          else
          begin
           SetFocusControl('PSE_TICKLIVR');
          end;
        end;

      

    end;
// f PT21
  end;
  // f PT7
  // DEB PT16
  if (F.FieldName = 'PSE_SALARIE') then
    begin
      if GetParamSocSecur('SO_PGINTERVENANTEXT', FALSE) Or PGBundleHierarchie then  //PT18 PT23 PT25
      begin
        if (RechDom ('PGSALARIEINT', GetField('PSE_SALARIE'), FALSE) = '') AND (GetField ('PSE_SALARIE') <> '') then
        begin
          PGIError ('Le salarié est inconnu', Ecran.Caption);
          if (Action<>'CONSULTATION') then //PT24
            SetField ('PSE_SALARIE', '');
        end;
      end
      else
      begin
         if (RechDom ('PGSALARIE', GetField('PSE_SALARIE'), FALSE) = '') AND (GetField ('PSE_SALARIE') <> '') then
         begin
          PGIError ('Le salarié est inconnu', Ecran.Caption);
          if (Action<>'CONSULTATION') then //PT24
            SetField ('PSE_SALARIE', '');
         end;
      end;
    end;
  // FIN PT16  
end;

procedure TOM_DEPORTSAL.OnArgument(S: string);
var
  TBtp, TResp, TInt, TMsa, TTckRestau: TTabSheet; // PT7
{$IFNDEF EAGLCLIENT}
  DBSal, MetierBTP, THService: THDBEdit;
  TypeAct: THDBValComboBox;
{$ELSE}
  DBSal, MetierBTP, THService: THEdit;
  TypeAct: THValComboBox;
{$ENDIF}

// d PT21
{$IFDEF EAGLCLIENT}
  Ticklivr : THEdit;
  TypTickLivr : THValComboBox;
{$ENDIF}
// f PT21

begin
  inherited;
  Bouge := 'AUCUN';
  Retour := False;
  TFFiche(Ecran).BFirst.OnClick := Premier;
  TFFiche(Ecran).BPrev.OnClick := Precedent;
  TFFiche(Ecran).BNext.OnClick := Suivant;
  TFFiche(Ecran).BLast.OnClick := Dernier;
  Sal := (Trim(ReadTokenSt(S)));
  Arg := Trim(ReadTokenPipe(S, ';'));
  Etat := Trim(ReadTokenPipe(S, ';'));
  Action := Trim(ReadTokenPipe(S, ';'));
  tAction := ReadtokenPipe(Action, '=');
  TestTrim := '';
  TestAnnee := '';
  TestSalarie := '';
  GestionUnique := False;
  BOrg := VH_Paie.PGRESPONSABLES;
  BMsa := VH_Paie.PGMsa;
  BIntermittents := VH_Paie.PGIntermittents; // récupération des paramètres sociétés
  {PT1
  If (BOrg=True) and (BMsa=False) and (BIntermittents=False) Then GestionUnique:=True;
  If (BOrg=False) and (BMsa=True) and (BIntermittents=False) Then GestionUnique:=True;
  If (BOrg=False) and (BMsa=False) and (BIntermittents=True) Then GestionUnique:=True;
  }
  BBTP := VH_Paie.PGBTP;
  BTckRestau := VH_Paie.PGTicketRestau; // PT7
  if (BOrg = True) and (BMsa = False) and (BIntermittents = False) and (BBTP = False) and (BTckRestau = False) then GestionUnique := True;
  if (BOrg = False) and (BMsa = True) and (BIntermittents = False) and (BBTP = False) and (BTckRestau = False) then GestionUnique := True;
  if (BOrg = False) and (BMsa = False) and (BIntermittents = True) and (BBTP = False) and (BTckRestau = False) then GestionUnique := True;
  if (BOrg = False) and (BMsa = False) and (BIntermittents = False) and (BBTP = True) and (BTckRestau = False) then GestionUnique := True;
  if (BOrg = False) and (BMsa = False) and (BIntermittents = False) and (BBTP = False) and (BTckRestau = True) then GestionUnique := True; // PT7
  //FIN PT1
  P := TPageControl(GetControl('PAGES'));
  if Action = 'CREATION' then TFFiche(Ecran).TypeAction := TaCreat;
  TInt := TTabSheet(GetControl('PINTERMITTENTS'));
  TMsa := TTabSheet(GetControl('PMSA'));
  Tresp := TTabSheet(GetControl('PGeneral'));
  TBtp := TTabSheet(GetControl('PBTP'));
  TTckRestau := TTabSheet(GetControl('PTICKETRESTAU')); // PT7
  if Arg = 'IS' then
  begin
    Requete := 'SELECT PSE_SALARIE FROM PGDEPORTSALINT';
    NomPage := 'PINTERMITTENTS';
    ChampCle := 'PSE_SALARIE_';
    Titre := 'Saisie intermittents du spectacle de : ';
    TResp.TabVisible := False;
    TInt.TabVisible := True;
    TMsa.TabVisible := False;
    TBtp.TabVisible := False;
    TTckRestau.TabVisible := False; // PT7
  end;
  if Arg = 'MSA' then
  begin
    Requete := 'SELECT PSE_SALARIE FROM PGDEPORTSALMSA';
    ChampCle := 'PSE_SALARIE__';
    NomPage := 'PMSA';
    Titre := 'Saisie MSA de : ';
    TResp.TabVisible := False;
    TInt.TabVisible := False;
    TMsa.TabVisible := True;
    TBtp.TabVisible := False;
    TTckRestau.TabVisible := False; // PT7
{$IFNDEF EAGLCLIENT}
    TypeAct := THDBValComboBox(GetControl('PSE_MSATYPEACT')); //DEBUT PT2
{$ELSE}
    TypeAct := THValComboBox(GetControl('PSE_MSATYPEACT'));
{$ENDIF}
    TypeAct.Items.Delete(2); // FIN PT2
  end;
  if Arg = 'RESP' then
  begin
    Requete := 'SELECT PSE_SALARIE FROM PGDEPORTSAL';
    ChampCle := 'PSE_SALARIE';
    NomPage := 'PGeneral';
    Titre := 'Saisie complémentaires de : ';
    TResp.TabVisible := True;
    TInt.TabVisible := False;
    TMsa.TabVisible := False;
    TBtp.TabVisible := False;
    TTckRestau.TabVisible := False; // PT7
    //   RespAbs:=THLabel(GetControl('LBLNOMRESP'));
{$IFNDEF EAGLCLIENT}
    THService := THDBEdit(GetControl('PSE_CODESERVICE'));
{$ELSE}
    THService := THEdit(GetControl('PSE_CODESERVICE'));
{$ENDIF}
    if THService <> nil then THService.OnElipsisClick := ServiceElipsisClick;
  end;
  if Arg = 'BTP' then
  begin
    Requete := 'SELECT PSE_SALARIE FROM PGDEPORTSALBTP';
    ChampCle := 'PSE_SALARIE___';
    NomPage := 'PBTP';
    Titre := 'Saisie complémentaires BTP de : ';
    TResp.TabVisible := False;
    TInt.TabVisible := False;
    TMsa.TabVisible := False;
    TBtp.TabVisible := True;
    TTckRestau.TabVisible := False; // PT7
    //PT13
{$IFNDEF EAGLCLIENT}
    MetierBTP := THDBEdit(GetControl('PSE_METIERBTP'));
{$ELSE}
    MetierBTP := THEdit(GetControl('PSE_METIERBTP'));
{$ENDIF}
    if (MetierBTP <> nil) then
      MetierBTP.OnElipsisClick := MetierBTPClick;
    //FIN PT13
  end;
  // d PT7
  if Arg = 'TCK' then
  begin
    Requete := 'SELECT PSE_SALARIE FROM PGDEPORTSALTCK';
    ChampCle := 'PSE_SALARIE___';
    NomPage := 'PTICKETRESTAU';
    Titre := 'Saisie paramétrage titre restaurant de : ';
    TResp.TabVisible := False;
    TInt.TabVisible := False;
    TMsa.TabVisible := False;
    TBtp.TabVisible := False;
    TTckRestau.TabVisible := True;

// d PT21
    // Tickets restaurant (ACCOR) - Code point de livraison
    fournisseur := '';
    if (VH_Paie.PGTicketRestau) then
    begin
      fournisseur := GetParamSocSecur ('SO_PGTYPECDETICKET','');
     if (fournisseur = '003') then

      // ACCOR  
      begin
        SetControlVisible('DISTRIBACCOR', True);     // PT22
       if (V_PGI.NumVersionBase < 800) then
       begin
         if (not VH_Paie.PGMsa) then
         // on ne traite le code point de livraion que si dossier non MSA
         begin
           SetControlVisible('TPSE_TICKLIVR', TRUE);
           SetControlVisible('PSE_MSATYPUNITEG1', TRUE);
           SetControlEnabled('PSE_MSATYPUNITEG1', TRUE);
           SetControlVisible('PSE_MSAUNITEGES1', TRUE);
           SetControlEnabled('PSE_MSAUNITEGES1', TRUE);
{$IFDEF EAGLCLIENT}
           Ticklivr := THEdit(GetControl('PSE_MSAUNITEGES1'));
           TickLivr.OnExit := SaisieCodeLivr;
           TypTickLivr := THValComboBox(GetControl('PSE_MSATYPUNITEG1'));
           TypTickLivr.OnExit := SaisieCodeLivr;
{$ENDIF}

         end
         else
        // dossier MSA - code point de livraison non traité
         begin
           SetControlVisible('TPSE_TICKLIVR', FALSE);
           SetControlVisible('PSE_MSATYPUNITEG1', FALSE);
           SetControlEnabled('PSE_MSATYPUNITEG1', FALSE);
           SetControlVisible('PSE_MSAUNITEGES1', FALSE);
           SetControlEnabled('PSE_MSAUNITEGES1', FALSE);
         end;
       end
       else
       begin
         SetControlVisible('TPSE_TICKLIVR', TRUE);
         SetControlVisible('PSE_TYPTICKLIVR', TRUE);
         SetControlEnabled('PSE_TYPTICKLIVR', TRUE);
         SetControlVisible('PSE_TICKLIVR', TRUE);

       end;
      end
      else
      // SODEXHO, NATEXIS , Chèque déjeuner
      begin
        SetControlVisible('DISTRIBACCOR', False);     // PT22
        if (V_PGI.NumVersionBase < 800) then
        begin
          SetControlVisible('TPSE_TICKLIVR', FALSE);
          SetControlVisible('PSE_MSATYPUNITEG1', FALSE);
          SetControlEnabled('PSE_MSATYPUNITEG1', FALSE);
          SetControlVisible('PSE_MSAUNITEGES1', FALSE);
          SetControlEnabled('PSE_MSAUNITEGES1', FALSE);
        end
        else
        begin
          SetControlVisible('TPSE_TICKLIVR', FALSE);
          SetControlVisible('PSE_TYPTICKLIVR', FALSE);
          SetControlEnabled('PSE_TYPTICKLIVR', FALSE);
          SetControlVisible('PSE_TICKLIVR', FALSE);
          SetControlEnabled('PSE_TICKLIVR', FALSE);
        end;
         // deb pt28
         if (fournisseur = '004') then
         setcontrolcaption('TPSE_DISTRIBUTION', 'Code distribution (succursale)');

      end;
    end;
// f PT21

  end;
  // f PT7

  // On prend le contrôle sur les 2 boutons supprimer : BSupp : Non géré par AGL et BDelete géré par l'AGL
  // Suivants les différents cas, l'un ou l'autre est accessible voir procedure OnLoad.
  BSupp := TToolBarButton97(GetControl('BDELETE1'));
  BDelete := TToolBarButton97(GetControl('BDELETE'));
  if GestionUnique = False then
  begin
    if BSupp <> nil then BSupp.OnClick := Suppression;
  end
  else
  begin
    TFFiche(Ecran).BInsert.Visible := True;
    BSupp.Visible := False;
    BDelete.Visible := True;
{$IFNDEF EAGLCLIENT}
    DBSal := THDBEdit(GetControl(ChampCle));
{$ELSE}
    DBSal := THEdit(GetControl(ChampCle));
{$ENDIF}
    if DBSal <> nil then DBSal.OnExit := AfficheEcran;
  end;
{$IFDEF EAGLCLIENT}
  QDeportSal := TFFiche(Ecran).FMulQ;
  TFFiche(Ecran).MonoFiche := FALSE;
{$ELSE}
  if Arg <> 'RESP' then
  begin
    if (action = 'CREATION') or (Etat = 'NOUVEAU') then //PT8
    begin
      TFFiche(Ecran).MonoFiche := TRUE;
    end
    else
    begin
      TFFiche(Ecran).MonoFiche := FALSE;
      QDeportSal := PrepareSQL(requete, False);
      RecupWhereSQL(TFFiche(Ecran).FMulQ, QDeportSal);
      QDeportSal.Open;
    end;
  end
  else
  begin
    QDeportSal := TFFiche(Ecran).FMulQ;
    TFFiche(Ecran).MonoFiche := TRUE;
  end;
{$ENDIF}

	//PT27
	If PGBundleHierarchie And (Not PGDroitMultiForm) Then 
		SetControlProperty(ChampCle, 'Plus', ' AND PSI_NODOSSIER="'+V_PGI.NoDossier+'"');

end;

procedure TOM_DEPORTSAL.AfficheEcran(Sender: TObject);
begin
  Ecran.caption := Titre + GetField('PSE_SALARIE') + ' ' + RechDom('PGSALARIE', GetField('PSE_SALARIE'), FALSE);
  TFFiche(Ecran).DisabledMajCaption := True;
end;

// Procédure qui gère la suppression si le salarié est présent dans plusieurs cas (Msa,Organigramme,Intermittent).

procedure TOM_DEPORTSAL.Suppression(Sender: TObject);
var
  SQL: string;
begin
  if Arg = 'IS' then SQL := 'UPDATE DEPORTSAL SET PSE_INTERMITTENT="-",PSE_ISNUMIDENT="",PSE_ISNUMASSEDIC="",PSE_ISCATEG="",PSE_ISRETRAITE="" WHERE PSE_SALARIE="' + GetField('PSE_SALARIE') + '"';
  if Arg = 'MSA' then SQL := 'UPDATE DEPORTSAL SET PSE_MSA="-",PSE_MSALIEUTRAV="",PSE_MSAINFOSCOMPL="",PSE_MSATECHNIQUE="-" WHERE PSE_SALARIE="' + GetField('PSE_SALARIE') + '"'; //PT10
  if Arg = 'RESP' then SQL := 'UPDATE DEPORTSAL SET PSE_ORGANIGRAMME="-" WHERE PSE_SALARIE="' + GetField('PSE_SALARIE') + '"';
  if Arg = 'BTP' then SQL := 'UPDATE DEPORTSAL SET PSE_BTP="-" WHERE PSE_SALARIE="' + GetField('PSE_SALARIE') + '"';
  if Arg = 'TCK' then
    SQL := 'UPDATE DEPORTSAL SET PSE_TICKETREST="-" WHERE PSE_SALARIE="' + GetField('PSE_SALARIE') + '"'; //@à
  case PGIAskCancel('Confirmez-vous la suppression de l''enregistrement ?', TFFiche(Ecran).Caption) of
    mrYes:
      begin
        ExecuteSQL(SQL);
        Suivant(nil);
      end;
    mrNo: exit;
    mrCancel: exit;
  end;
end;

// Fonction qui renvoi la valeur true si le salrié est dans la catégorie placé en paramètre.

function TOM_DEPORTSAL.TestExistence(ChampTest: string): Boolean;
begin
  Result := ExisteSQL('SELECT PSE_SALARIE FROM DEPORTSAL WHERE PSE_SALARIE="' + GetField('PSE_SALARIE') + '" and ' + ChampTest + '="X"');
end;

procedure TOM_DEPORTSAL.premier(Sender: TObject);
begin
{$IFNDEF EAGLCLIENT}
  TFFiche(Ecran).BLast.Enabled := true;
  TFFiche(Ecran).BNext.Enabled := true;
  Ferme(QDeportSal);
  QDeportSal := PrepareSQL(Requete, False);
  RecupWhereSQL(TFFiche(Ecran).FMulQ, QDeportSal);
  QDeportSal.Open;
  Bouge := 'PREMIER';
  Retour := False;
  TFFiche(Ecran).BFirstClick(TFFiche(Ecran).BFirst);
{$ENDIF}
end;

procedure TOM_DEPORTSAL.dernier(Sender: TObject);
begin
{$IFNDEF EAGLCLIENT}
  TFFiche(Ecran).BFirst.Enabled := true;
  TFFiche(Ecran).BPrev.Enabled := true;
  Ferme(QDeportSal);
  QDeportSal := PrepareSQL(Requete, False);
  RecupWhereSQL(TFFiche(Ecran).FMulQ, QDeportSal);
  QDeportSal.Open;
  Bouge := 'DERNIER';
  Retour := False;
  TFFiche(Ecran).BLastClick(TFFiche(Ecran).BLast);
{$ENDIF}
end;

procedure TOM_DEPORTSAL.suivant(Sender: TObject);
begin
{$IFNDEF EAGLCLIENT}
  TFFiche(Ecran).BFirst.Enabled := true;
  TFFiche(Ecran).BPrev.Enabled := true;
  Ferme(QDeportSal);
  QDeportSal := PrepareSQL(Requete, False);
  RecupWhereSQL(TFFiche(Ecran).FMulQ, QDeportSal);
  QDeportSal.Open;
  Bouge := 'SUIVANT';
  Retour := False;
  TFFiche(Ecran).BNextClick(TFFiche(Ecran).BNext);
{$ENDIF}
end;

procedure TOM_DEPORTSAL.precedent(Sender: TObject);
begin
{$IFNDEF EAGLCLIENT}
  TFFiche(Ecran).BLast.Enabled := true;
  TFFiche(Ecran).BNext.Enabled := true;
  Ferme(QDeportSal);
  QDeportSal := PrepareSQL(Requete, False);
  RecupWhereSQL(TFFiche(Ecran).FMulQ, QDeportSal);
  QDeportSal.Open;
  Bouge := 'PRECEDENT';
  Retour := False;
  TFFiche(Ecran).BPrevClick(TFFiche(Ecran).BPrev);
{$ENDIF}
end;

procedure TOM_DEPORTSAL.ServiceElipsisClick(Sender: TObject);
// PT21 var
// PT21   St: string;
Var Select,Where : String;
begin
  if Sender = nil then Exit;
  if GetField('PSE_SALARIE') = '' then Exit;
  //St:='SELECT PGS_CODESERVICE,PGS_NOMSERVICE,PHO_LIBELLE,PHO_NIVEAUH,PSA_LIBELLE FROM SERVICES'+
  //' LEFT JOIN SALARIES ON PSA_SALARIE=PGS_RESPSERVICE LEFT JOIN HIERARCHIE ON PHO_HIERARCHIE=PGS_HIERARCHIE';

  //LookupList(THDBEdit(Sender),'Liste des services','SERVICES','PGS_CODESERVICE','PGS_NOMSERVICE,PHO_NIVEAUH,PHO_LIBELLE,PSA_LIBELLE','PHO_NIVEAUH>0','PHO_NIVEAUH', True,-1,St);
  
  //PT27
  If PGBundleHierarchie Then
  Begin
  	Where := 'SERVICES LEFT JOIN INTERIMAIRES ON PSI_INTERIMAIRE=PGS_RESPSERVICE LEFT JOIN HIERARCHIE ON PHO_HIERARCHIE=PGS_HIERARCHIE ';
    // Cas particulier du superviseur qui peut affecter un salarié à un service non rattaché à un dossier
    If JaiLeDroitTag(46530) Then
    Begin
      	Where := Where + ' AND (PGS_NOMSERVICE LIKE "'+GetNoDossierSalarie(GetField('PSE_SALARIE'))+'%"';
        Where := Where + ' OR PGS_NOMSERVICE NOT LIKE "%#_%" ESCAPE "#")';
    End
    Else
   	    Where := Where + ' AND PGS_NOMSERVICE LIKE "'+GetNoDossierSalarie(GetField('PSE_SALARIE'))+'%"';
  	Select := 'PSI_LIBELLE';
  End
  Else
  Begin
  	Where := 'SERVICES LEFT JOIN SALARIES ON PSA_SALARIE=PGS_RESPSERVICE LEFT JOIN HIERARCHIE ON PHO_HIERARCHIE=PGS_HIERARCHIE';
  	Select := 'PSA_LIBELLE';
  End;
  	
  LookupList({$IFNDEF EAGLCLIENT}THDBEdit{$ELSE}THEdit{$ENDIF}(Sender), 'Liste des services', Where, 'PGS_CODESERVICE', 'PGS_NOMSERVICE,PHO_NIVEAUH,PHO_LIBELLE,'+Select, 'PHO_NIVEAUH>0', 'PHO_NIVEAUH', True, -1);

  //LookupList(THEdit(Sender),'Liste des services','SERVICES','PGS_CODESERVICE','PGS_NOMSERVICE,PHO_NIVEAUH,PHO_LIBELLE,PSA_LIBELLE','PHO_NIVEAUH>0','PHO_NIVEAUH', True,-1,St);
  //LookupList((Sender), 'Liste des services', 'SERVICES LEFT JOIN SALARIES ON PSA_SALARIE=PGS_RESPSERVICE LEFT JOIN HIERARCHIE ON PHO_HIERARCHIE=PGS_HIERARCHIE', 'PGS_CODESERVICE', 'PGS_NOMSERVICE,PHO_NIVEAUH,PHO_LIBELLE,PSA_LIBELLE', 'PHO_NIVEAUH>0', 'PHO_NIVEAUH', True, -1); // PT15
end;


//PT13
{***********A.G.L.Privé.*****************************************
Auteur  ...... : Vincent GALLIOT
Créé le ...... : 24/05/2004
Modifié le ... :   /  /
Description .. : Gestion de l'ellipsis du Métier BTP
Mots clefs ... : PAIE,PGDADSU,DEPORTSAL,DADSU
*****************************************************************}

procedure TOM_DEPORTSAL.MetierBTPClick(Sender: TObject);
var
  sWhere: string;
  QRechMetier: TQuery;
begin
  if (Sender = nil) then
    Exit;
  if GetField('PSE_SALARIE') = '' then
    Exit;

{$IFNDEF EAGLCLIENT}
  LookupList(THDBEdit(Sender), 'Liste des Métiers BTP', 'METIERBTP', 'PMB_METIERBTP',
    'PMB_CONVCOLL,PMB_SPECIALITE,PMB_METIERLIB', '', '', True, -1, '',
    tlDefault, 40);
{$ELSE}
  LookupList(THEdit(Sender), 'Liste des Métiers BTP', 'METIERBTP', 'PMB_METIERBTP',
    'PMB_CONVCOLL,PMB_SPECIALITE,PMB_METIERLIB', '', '', True, -1, '',
    tlDefault, 40);
{$ENDIF}
  sWhere := 'SELECT PMB_METIERLIB' +
    ' FROM METIERBTP WHERE' +
    ' ##PMB_PREDEFINI## PMB_METIERBTP = "' + GetControlText('PSE_METIERBTP') + '"';
  QRechMetier := OpenSql(sWhere, TRUE);
  if (not QRechMetier.EOF) then
    SetControlText('LMETIERBTP_', QRechMetier.Fields[0].asstring);
  Ferme(QRechMetier);
end;
//FIN PT13

// d PT21
{$IFDEF EAGLCLIENT}
procedure TOM_DEPORTSAL.SaisieCodeLivr(Sender : TObject);
begin
          ForceUpdate;
          SetField ('PSE_MSAUNITEGES',GetControlText('PSE_MSAUNITEGES1')) ;
          SetField ('PSE_MSATYPUNITEG',GetControlText('PSE_MSATYPUNITEG1')) ;
end;
{$ENDIF}
// f PT21
initialization
  registerclasses([TOM_DEPORTSAL]);
end.

