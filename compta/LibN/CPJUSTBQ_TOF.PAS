{***********UNITE*************************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 31/07/2002
Modifié le ... :   /  /
Description .. : Source TOF de la FICHE : CPJUSTBQ ()
Mots clefs ... : TOF;CPJUSTBQ
*****************************************************************}
Unit CPJUSTBQ_TOF ;

Interface

Uses StdCtrls, Controls, Classes,
{$IFDEF EAGLCLIENT}
     UTOB,
     MaineAGL,
     eQRS1,
{$ELSE}
     db,
     {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}
     FE_Main,
     QRS1,
{$ENDIF}
     forms,
     sysutils,
     ComCtrls,
     HCtrls,
     HMsgBox,
     UTOF,
     Ent1,
     Htb97,
     TofMeth,      // TOF_METH
     CritEdt,      // ClassCritEdt
     AGLInit,      // TheData
     uLibExercice, // CMajRequeteExercice
     HQry,         //RecupWhereCritere
     LookUp ;      // LookUpList

procedure CC_LanceFicheJustifPointage;

Type
  TOF_CPJUSTBQ = Class (TOF_METH)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure ChargementCritEdt        ; override;
    procedure OnArgument (S : String ) ; override ;
    procedure OnClose                  ; override ;

  private
    gcbo_Exercice    : THValComboBox;
    ggbx_Mouvements  : THRadioGroup;

    E_General        : THEdit;
    E_General_       : THEdit;
    XXGeneJal        : THEdit;
    XXGroupBy        : THEdit;

    E_Exercice       : THValComboBox;
    E_DateComptable  : THEdit;
    E_DateComptable_ : THEdit;

    DateExoPrecedent : THEdit;
    XX_Rupture       : THEdit;

    TE_General       : THLabel;
    TE_General_      : THLabel;

    Pages            : TPageControl;
    FFiltres         : THvalComboBox;

    procedure ggbx_MouvementsClick    ( Sender : TObject );
    procedure OnExitE_General         ( Sender : TObject );
    procedure OnElipsisClickE_General ( Sender : TObject );
    procedure OnChangeE_Exercice      ( Sender : TObject );
    procedure OnChangeFiltre          ( Sender : TObject ); override;

  end ;

Implementation

uses
  {$IFDEF MODENT1}
  CPProcMetier,
  CPTypeCons,
  {$ENDIF MODENT1}
  HEnt1;


////////////////////////////////////////////////////////////////////////////////
procedure CC_LanceFicheJustifPointage;
begin
  AGLLanceFiche('CP','CPJBANQ','','','');
end ;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... :   /  /
Modifié le ... : 25/10/2004
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPJUSTBQ.OnNew ;
var lExoDate : TExoDate;
begin
  inherited ;


  // GCO - 27/10/2003 Problèmes de rechargement des dates exemple
  // Exo 2004, D1 : 01/01/2003, D2 : 31/12/2003 enregistré dans le filtre
  lExoDate.Code := CRelatifVersExercice( gcbo_Exercice.Value );
  RempliExoDate( lExoDate );
  if gcbo_Exercice.ItemIndex > 0 then
  begin
    if ((StrToDate(E_DateComptable.Text) < lExoDate.Deb) or
        (StrToDate(E_DateComptable.Text) > lExoDate.Fin)) and
       ((StrToDate(E_DateComptable_.Text) < lExoDate.Deb) or
        (StrToDate(E_DateComptable_.Text) > lExoDate.Fin)) then
      OnChangeE_Exercice( nil );
  end;
end ;

////////////////////////////////////////////////////////////////////////////////
procedure TOF_CPJUSTBQ.OnDelete ;
begin
  Inherited ;
end ;

////////////////////////////////////////////////////////////////////////////////
procedure TOF_CPJUSTBQ.OnUpdate ;
var
  lStControleDevise : string;
  szWhere           : string;
  StV8              : string;
  C                 : THEdit;
begin
  Inherited ;

  TFQRS1(Ecran).WhereSQL := 'SELECT G_GENERAL, G_LIBELLE, BQ_BANQUE, PQ_LIBELLE, ' +
                            'BQ_DEVISE, BQ_DOMICILIATION, BQ_ETABBQ, BQ_GUICHET, ' +
                            'BQ_NUMEROCOMPTE, BQ_CLERIB, BQ_CODEBIC, E_GENERAL, ' +
                            'E_DATECOMPTABLE, E_NUMEROPIECE, E_LIBELLE, ' +
                            'E_DATEECHEANCE, E_REFINTERNE, E_REFPOINTAGE, '+
                            'E_DATEPOINTAGE, E_MODEPAIE, E_JOURNAL, E_EXERCICE, ' +
                            'E_NUMLIGNE, E_DEVISE, E_DEBIT, E_CREDIT, ' +
                            'E_DEBITDEV, E_CREDITDEV, EE_NEWSOLDEDEB, EE_NEWSOLDECRE, ' +
                            'EE_NEWSOLDEDEBEURO, EE_NEWSOLDECREEURO, J_JOURNAL, J_LIBELLE, ' +
                            'J_CONTREPARTIE FROM ECRITURE ' +
                            'LEFT JOIN JOURNAL ON J_JOURNAL = E_JOURNAL ';

  // Left Join différent sur GENERAUX en fonction du mode de pointage
  // Left Join différent sur EEXBQ en fonction du mode de pointage
  // GCO - 04/01/2006 FQ 15028
  if VH^.PointageJal then
  begin
    TFQRS1(Ecran).WhereSQL := TFQRS1(Ecran).WhereSQL +
                              ' LEFT JOIN GENERAUX ON G_GENERAL = J_CONTREPARTIE ' +
                              ' LEFT JOIN EEXBQ ON E_JOURNAL = EE_GENERAL';
  end
  else
  begin
    TFQRS1(Ecran).WhereSQL := TFQRS1(Ecran).WhereSQL +
                              ' LEFT JOIN GENERAUX ON G_GENERAL=E_GENERAL ' +
                              ' LEFT JOIN EEXBQ ON E_GENERAL = EE_GENERAL';
  end;

  TFQRS1(Ecran).WhereSQL := TFQRS1(Ecran).WhereSQL + ' AND E_REFPOINTAGE=EE_REFPOINTAGE' +
                            ' AND E_DATEPOINTAGE=EE_DATEPOINTAGE ';

  // Left Join identique pour le reste de la requête
  TFQRS1(Ecran).WhereSQL := TFQRS1(Ecran).WhereSQL + ' LEFT JOIN BANQUECP ON BQ_GENERAL=G_GENERAL AND BQ_NODOSSIER="'+V_PGI.NoDossier+'"' + // 19/10/2006 YMO Multisociétés
                                                     ' LEFT JOIN BANQUES ON BQ_BANQUE=PQ_BANQUE ';

  // Condition WHERE commune
  TFQRS1(Ecran).WhereSQL := TFQRS1(Ecran).WhereSQL + RecupWhereCritere( Pages );

  TFQRS1(Ecran).WhereSQL := TFQRS1(Ecran).WhereSQL + ' AND E_QUALIFPIECE="N" AND (E_ECRANOUVEAU="N" or E_ECRANOUVEAU="H")';

  // Modifie la clause Where
  if VH^.PointageJal then
  begin
    XXGENEJAL.Text         := ' E_JOURNAL = "' + E_General.Text + '"' ;
    XXGroupBy.text         := ' GROUP BY E_JOURNAL';
    TFQRS1(Ecran).WhereSQL := TFQRS1(Ecran).WhereSQL + ' AND E_JOURNAL >= "' + E_General.Text + '"'
                                                     + ' AND E_JOURNAL <= "' + E_General_.Text + '"'
                                                     + ' AND J_NATUREJAL = "BQE" AND G_POINTABLE = "X"'
                                                     + ' AND E_GENERAL <> J_CONTREPARTIE';
  end
  else
  begin
    XXGENEJAL.Text         := ' E_GENERAL = "' + E_General.Text + '"' ;
    XXGroupBy.Text         := ' GROUP BY E_GENERAL';
    TFQRS1(Ecran).WhereSQL := TFQRS1(Ecran).WhereSQL + ' AND E_GENERAL >= "' + E_General.Text + '"'
                                                     + ' AND E_GENERAL <= "' + E_General_.Text + '"'
// BPY le 28/07/2004 => Fiche n° 13999 : de nature banque et pointable !
                                                     + ' AND G_NATUREGENE = "BQE" AND G_POINTABLE = "X"';
  end;

  // VH^.ExoV8
  StV8 := LWhereV8 ;
  if StV8 <> '' then
  begin
    TFQRS1(Ecran).WhereSQL := TFQRS1(Ecran).WhereSQL + ' AND ' + StV8;
    C := THEdit(GetControl('STV8', True));
    if (C <> nil) then
      C.Text := ' AND ' + StV8;
  end;

  // Modifie la clause Where pour tenir compte des mouvements Pointés,Non pointés ou Tous
  if (ggbx_Mouvements.Value = '1') then
    szWhere := 'E_REFPOINTAGE <> ""' // Mouvements pointés
  else
    if (ggbx_Mouvements.Value = '2') then
      szWhere := 'E_REFPOINTAGE = ""'; // Mouvements non pointés

  if szWhere <> '' then
    TFQRS1(Ecran).WhereSQL := TFQRS1(Ecran).WhereSQL + ' AND ' + szWhere;

  // !Traduction de la notion relative de l'exercice
  TFQRS1(Ecran).WhereSQL := CMajRequeteExercice( gcbo_Exercice.Value, TFQRS1(Ecran).WhereSQL);

  // On se sert de la requête principal pour déterminer la présence d'écritures
  // das nue autre devise du compte
  lStControleDevise := TFQRS1(Ecran).WhereSQL;

  // Gestion des DEVISES
  TFQRS1(Ecran).WhereSQL := TFQRS1(Ecran).WhereSQL +
  ' AND ((BQ_DEVISE <> "' + V_PGI.DevisePivot + '" AND E_DEVISE = BQ_DEVISE)' +
  ' OR (BQ_DEVISE = "' + V_PGI.DevisePivot + '"))';

  lStControleDevise := lStControleDevise + ' AND (E_DEVISE <> BQ_DEVISE AND ' +
                            'BQ_DEVISE <> "' + V_PGI.DevisePivot + '")';

  // Ajout de l'ORDER BY
  if VH^.PointageJal then
  begin
    TFQRS1(Ecran).WhereSQL := TFQRS1(Ecran).WhereSQL + ' ORDER BY BQ_BANQUE, E_JOURNAL, E_DATEPOINTAGE, E_DATECOMPTABLE, E_NUMEROPIECE';
    lStControleDevise      := lStControleDevise + ' ORDER BY BQ_BANQUE, E_JOURNAL, E_DATEPOINTAGE, E_DATECOMPTABLE, E_NUMEROPIECE';
  end
  else
  begin
    TFQRS1(Ecran).WhereSQL :=  TFQRS1(Ecran).WhereSQL + ' ORDER BY BQ_BANQUE, E_GENERAL, E_DATEPOINTAGE, E_DATECOMPTABLE, E_NUMEROPIECE';
    lStControleDevise      :=  lStControleDevise + ' ORDER BY BQ_BANQUE, E_GENERAL, E_DATEPOINTAGE, E_DATECOMPTABLE, E_NUMEROPIECE';
  end;

  // Affichage si incohérence des DEVISES
  if ExisteSql( lStControleDevise ) then
  begin
    PgiInfo('Attention  : certaines écritures ont une monnaie de saisie ' +
            'différente de celle du compte de banque.', Ecran.Caption);
  end;
end ;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 06/09/2005
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPJUSTBQ.ChargementCritEdt;
begin
  inherited;
  if (TheData <> nil) and (TheData is ClassCritEdt) then
  begin
    E_General.Text  := ClassCritEdt(TheData).CritEdt.Cpt1;
    E_General_.Text := ClassCritEdt(TheData).CritEdt.Cpt2;

    E_Exercice.Value := CExerciceVersRelatif(ClassCritEdt(TheData).CritEdt.Exo.Code);
    // Date de Début de l'édition
    E_DateComptable.Text  := DateToStr(ClassCritEdt(TheData).CritEdt.Date1);
    // Date de Fin de l'édition
    E_DateComptable_.Text := DateToStr(ClassCritEdt(TheData).CritEdt.Date2);
    // Etablissement
    THValComboBox(GetControl('BQ_BANQUE',True)).Value := ClassCritEdt(TheData).CritEdt.Etab;
    // Devise
    THValComboBox(GetControl('BQ_DEVISE',True)).Value := ClassCritEdt(TheData).CritEdt.DeviseSelect;
    TheData := nil;
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 06/09/2005
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPJUSTBQ.OnLoad ;
var lQuery : TQuery;
    lStSql : string;
begin
  Inherited ;
  // Remplissage des bornes des comptes généraux
  lQuery := nil;
  if (Trim(E_General.Text) = '') or (Trim(E_General_.Text) = '') then
  begin
    try
      try
        if VH^.PointageJal then
          lStSql := 'SELECT MIN(J_JOURNAL)MINI, MAX(J_JOURNAL)MAXI FROM JOURNAL LEFT JOIN ' +
                    'GENERAUX ON J_CONTREPARTIE=G_GENERAL WHERE J_NATUREJAL="BQE" ' +
                    'AND G_POINTABLE="X"'
        else
          lStSql := 'SELECT MIN(G_GENERAL)MINI, MAX(G_GENERAL)MAXI FROM GENERAUX WHERE G_NATUREGENE = "BQE" AND G_POINTABLE="X"';

        lQuery := OpenSql( lStSql , True);
        if not lQuery.Eof then
        begin
          if Trim(E_General.Text) = '' then
            E_General.Text := lQuery.FindField('MINI').AsString;

          if Trim(E_General_.Text) = '' then
            E_General_.Text := lQuery.FindField('MAXI').AsString;
        end;
      except
        on E: Exception do PgiError('Erreur de requête SQL : ' + E.Message, 'Procédure : OnLoad');
      end;
    finally
      Ferme( lQuery );
    end;
  end;
  DateExoPrecedent.Text := DateToStr( StrToDate( E_DateComptable.Text ) - 1);

end ;

////////////////////////////////////////////////////////////////////////////////
procedure TOF_CPJUSTBQ.OnArgument (S : String ) ;
var
  cbo_Etablissement : THValComboBox;
  cbo_Devise : THValComboBox;
begin
  Inherited ;

  Ecran.HelpContext:=7622000;

  TE_GENERAL  := THLabel(GetControl('TE_GENERAL', True));
  TE_GENERAL_ := THLabel(GetControl('TE_GENERAL_', True));

  FFiltres := THValComboBox(GetControl('FFILTRES'));
  FFiltres.OnChange := OnChangeFiltre;

  E_General   := THEdit(GetControl('EGENERAL',True));
  E_General_  := THEdit(GetControl('EGENERAL_',True));
  XXGeneJal   := THEdit(GetControl('XXGENEJAL',True));
  XXGroupBy   := THEdit(GetControl('XXGROUPBY',True));

  E_Exercice       := THValComboBox(GetControl('E_EXERCICE', True));
  E_DateComptable  := THEdit(GetControl('E_DATECOMPTABLE',True));
  E_DateComptable_ := THEdit(GetControl('E_DATECOMPTABLE_',True));
  DateExoPrecedent := THEdit(GetControl('DATEEXOPRECEDENT',True));

  Pages            := TPageControl(GetControl('PAGES', True));
  XX_Rupture  := THEdit(GetControl('XX_RUPTURE',True));

  if VH^.PointageJal then
  begin
    E_General.MaxLength  := 3;
    E_General_.MaxLength := 3;
    TE_GENERAL.Caption   := 'Journaux de';
    TE_GENERAL_.Caption  := 'à';
    XX_RUPTURE.Text      := 'E_JOURNAL';
  end
  else
  begin
    E_General.MaxLength  := VH^.CPta[fbGene].Lg;
    E_General_.MaxLength := VH^.CPta[fbGene].Lg;
    TE_GENERAL.Caption   := 'Comptes généraux du';
    TE_GENERAL_.Caption  := 'au';
    XX_RUPTURE.Text      := 'E_GENERAL';
  end;

  E_General.OnElipsisClick  := OnElipsisClickE_General;
  E_General_.OnElipsisClick := OnElipsisClickE_General;
  E_General.OnExit          := OnExitE_General;

  // Dernier exercice par défaut
  gcbo_Exercice := THValComboBox(GetControl('E_EXERCICE',True));
  CInitComboExercice( gcbo_Exercice );
  if ( CtxPCl in V_PGI.PgiContexte ) and  ( VH^.CPExoRef.Code <>'' ) then
    gcbo_Exercice.Value := CExerciceVersRelatif(VH^.CPExoRef.Code)
  else
    gcbo_Exercice.Value := CExerciceVersRelatif(VH^.Entree.Code) ;
  gcbo_Exercice.OnChange := OnChangeE_Exercice;
  OnChangeE_Exercice(nil);

  // Etablissement : Tous
  cbo_Etablissement := THValComboBox(GetControl('BQ_BANQUE',True));
  cbo_Etablissement.ItemIndex := 0;

  // Devises : Toutes
  cbo_Devise := THValComboBox(GetControl('BQ_DEVISE',True));
  cbo_Devise.ItemIndex := 0;
  cbo_Devise.Enabled   := not VH^.PointageJal;

  // Mouvements
  ggbx_Mouvements := THRadioGroup(GetControl('GBXMOUVEMENTS',True));
  ggbx_Mouvements.OnClick := ggbx_MouvementsClick;
  ggbx_MouvementsClick(nil);

  // Onglet de DEV visible en mode SAV
  TTabSheet(GetControl('TABDEV', True)).TabVisible := V_PGI.Sav;
end ;

procedure TOF_CPJUSTBQ.OnClose ;
begin
  Inherited ;
end ;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... :   /  /
Modifié le ... : 24/06/2003
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPJUSTBQ.OnChangeE_Exercice(Sender: TObject);
begin
  CExoRelatifToDates ( gcbo_Exercice.Value, E_DATECOMPTABLE, E_DATECOMPTABLE_);
end;

////////////////////////////////////////////////////////////////////////////////

procedure TOF_CPJUSTBQ.ggbx_MouvementsClick(Sender: TObject);
begin
  if (ggbx_Mouvements.Value = '0') then
    SetControlText('MOUVEMENTS','T')
  else
    if (ggbx_Mouvements.Value = '1') then
      SetControlText('MOUVEMENTS','P')
    else
      SetControlText('MOUVEMENTS','N');
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 05/06/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPJUSTBQ.OnExitE_General(Sender: TObject);
begin
  if Trim(E_General.Text) = '' then Exit;

  if (not VH^.PointageJal) and (length(E_General.Text) <> VH^.CPta[fbGene].Lg) then
    E_General.Text := BourreEtLess(E_General.Text, fbGene);
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 24/06/2003
Modifié le ... : 24/06/2003
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPJUSTBQ.OnElipsisClickE_General(Sender: TObject);
var lSt : string;
begin
  if VH^.PointageJal then
  begin
    lSt := 'SELECT J_JOURNAL, J_LIBELLE FROM JOURNAL LEFT JOIN GENERAUX ON J_CONTREPARTIE=G_GENERAL ' +
           'WHERE J_NATUREJAL="BQE" AND G_POINTABLE="X"';

    LookUpList(THEdit(Sender), 'Journal', 'JOURNAL', 'J_JOURNAL', 'J_LIBELLE', '', 'J_JOURNAL', True, 0 , lSt)
  end
  else
    LookUpList(THEdit(Sender), 'Compte général', 'GENERAUX', 'G_GENERAL', 'G_LIBELLE', 'G_NATUREGENE="BQE" AND G_POINTABLE="X"', 'G_GENERAL', True, 0 );
end;

////////////////////////////////////////////////////////////////////////////////
procedure TOF_CPJUSTBQ.OnChangeFiltre(Sender: TObject);
var lExoDate : TExoDate;
begin
  TFQRS1(Ecran).FFiltresChange(FFiltres);
  // GCO - 27/10/2003 Problèmes de rechargement des dates exemple
  // Exo 2004, D1 : 01/01/2003, D2 : 31/12/2003 enregistré dans le filtre
  lExoDate.Code := CRelatifVersExercice( gcbo_Exercice.Value );
  RempliExoDate( lExoDate );
  if gcbo_Exercice.ItemIndex > 0 then
  begin
    if ((StrToDate(E_DateComptable.Text) < lExoDate.Deb) or
        (StrToDate(E_DateComptable.Text) > lExoDate.Fin)) and
       ((StrToDate(E_DateComptable_.Text) < lExoDate.Deb) or
        (StrToDate(E_DateComptable_.Text) > lExoDate.Fin)) then
      OnChangeE_Exercice( nil );
  end;
end ;

Initialization
  registerclasses ( [ TOF_CPJUSTBQ ] ) ;
end.
