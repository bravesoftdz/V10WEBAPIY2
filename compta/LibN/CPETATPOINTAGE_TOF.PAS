{***********UNITE*************************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 19/07/2002
Modifié le ... : 18/02/2004
Description .. : Source TOF de la FICHE : CPETATPOINTAGE ()
Mots clefs ... : TOF;CPETATPOINTAGE
*****************************************************************}
Unit CPEtatPointage_TOF;

Interface

Uses StdCtrls,
     Controls,
     Classes,
{$IFDEF EAGLCLIENT}
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
     HEnt1,
     HMsgBox,
     UTOF,
     uTob,          // TOB
     LookUp ;

procedure CC_LanceFicheEtatPointage;

Type
  TOF_CPETATPOINTAGE = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnClose                  ; override ;

  private
    gchk_Nature      : TCheckBox;
    gcbo_Banque      : THValComboBox;
    gcbo_Devise      : THValComboBox;
    gchk_RuptureEtab : TCheckBox;

    E_General        : THEdit;
    E_General_       : THEdit;

    XX_Rupture       : THEdit;

    E_DatePointage   : THEdit;
    E_DatePointage_  : THEdit;

    TE_General       : THLabel;
    TE_General_      : THLabel;

    procedure chk_NatureClick         (Sender: TObject);
    procedure gtxt_DatePointageClick  (Sender: TObject);
    procedure OnElipsisClickE_General (Sender: TObject);
    procedure OnExitE_General         (Sender: TObject);

  end ;

Implementation

uses
  {$IFDEF MODENT1}
  CPTypeCons,
  CPProcMetier,
  {$ENDIF MODENT1}
  Ent1,          // VH^.
  HQry;          // RecupWhereCritere



procedure CC_LanceFicheEtatPointage;
begin
  AGLLanceFiche('CP','CPETATPOINTAGE','','','');
end;

procedure TOF_CPETATPOINTAGE.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_CPETATPOINTAGE.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_CPETATPOINTAGE.OnUpdate ;
begin
  Inherited ;

  if VH^.PointageJal then
  begin
    TFQRS1(Ecran).WhereSQL := 'SELECT J_LIBELLE, E_GENERAL, E_DATECOMPTABLE, ' +
    'E_NUMEROPIECE, E_LIBELLE, E_DATEECHEANCE, E_REFINTERNE, E_JOURNAL, ' +
    'E_EXERCICE, E_NUMLIGNE, E_DEVISE, E_DEBIT, E_CREDIT, ' +
    'E_DEBITDEV, E_CREDITDEV, E_REFPOINTAGE, E_DATEPOINTAGE, ' +
    'E_MODEPAIE, (EE_NEWSOLDECRE-EE_NEWSOLDEDEB) SOLDEDEV, ' +
    '(EE_NEWSOLDECREEURO-EE_NEWSOLDEDEBEURO) SOLDEEURO, ' +
    'BQ_BANQUE, BQ_DEVISE, BQ_DOMICILIATION, BQ_ETABBQ, ' +
    'BQ_GUICHET, BQ_NUMEROCOMPTE, BQ_CLERIB, BQ_CODEBIC, ' +
    'G_GENERAL, G_LIBELLE FROM ECRITURE ' +
    'LEFT JOIN JOURNAL ON J_JOURNAL = E_JOURNAL ' +
    'LEFT JOIN BANQUECP ON BQ_GENERAL = J_CONTREPARTIE ' +
    'AND BQ_NODOSSIER="'+V_PGI.NoDossier+'" ' + // 19/10/2006 YMO Multisociétés
    'LEFT JOIN BANQUES on PQ_BANQUE = BQ_BANQUE ' +
    'LEFT JOIN GENERAUX ON G_GENERAL = J_CONTREPARTIE ' +
    'LEFT JOIN EEXBQ ON EE_GENERAL = E_JOURNAL AND E_REFPOINTAGE = EE_REFPOINTAGE ' +
    'WHERE E_QUALIFPIECE="N" and (E_ECRANOUVEAU="N" or E_ECRANOUVEAU="H") and ' +
    'E_REFPOINTAGE <> "" ';

    if Trim(E_GENERAL.Text) <> '' then
      TFQRS1(Ecran).WhereSQL := TFQRS1(Ecran).WhereSQL + ' AND E_JOURNAL >= "' + E_GENERAL.Text + '"' ;

    if Trim(E_GENERAL_.Text) <> '' then
      TFQRS1(Ecran).WhereSQL := TFQRS1(Ecran).WhereSQL + ' AND E_JOURNAL <= "' + E_GENERAL_.Text + '"' ;

    if IsValidDate(E_DatePointage.Text) then
      TFQRS1(Ecran).WhereSQL := TFQRS1(Ecran).WhereSQL + ' AND E_DATEPOINTAGE >= "' + UsDateTime(StrToDate(E_DatePointage.Text)) + '"' ;

    if IsValidDate(E_DatePointage_.Text) then
      TFQRS1(Ecran).WhereSQL := TFQRS1(Ecran).WhereSQL + ' AND E_DATEPOINTAGE <= "' + UsDateTime(StrToDate(E_DatePointage_.Text)) + '"' ;

    // GCO - 02/06/2006 - FQ 13289 - inversement E_REFPOINTAGE, E_DATEPOINTAGE,
    TFQRS1(Ecran).WhereSQL := TFQRS1(Ecran).WhereSQL + ' ORDER BY E_JOURNAL, E_DATEPOINTAGE, E_REFPOINTAGE, E_DATECOMPTABLE, E_NUMEROPIECE';
  end
  else
  begin
     TFQRS1(Ecran).WhereSQL := TFQRS1(Ecran).WhereSQL + ' AND BQ_NODOSSIER = "' + V_PGI.NoDossier + '" '; // 23/10/2006 YMO Multisociétés
     TFQRS1(Ecran).WhereSQL := TFQRS1(Ecran).WhereSQL + ' AND E_GENERAL >= "' + E_GENERAL.Text + '"' +
                                                       ' AND E_GENERAL <= "' + E_GENERAL_.Text + '"';
  end;

  // Si compte de nature banque
  if gchk_Nature.Checked then
  begin
    if gcbo_Banque.ItemIndex <> 0 then // Si une banque est sélectionnée
    begin
      // Modifie la clause Where
      TFQRS1(Ecran).WhereSQL := TFQRS1(Ecran).WhereSQL + ' AND BQ_BANQUE = "' + gcbo_Banque.Value + '"';
      if gcbo_Devise.ItemIndex <> 0 then
        TFQRS1(Ecran).WhereSQL := TFQRS1(Ecran).WhereSQL + ' AND BQ_DEVISE = "' + gcbo_Devise.Value + '"';
    end
    else
    begin
      if gcbo_Devise.ItemIndex <> 0 then
        TFQRS1(Ecran).WhereSQL := TFQRS1(Ecran).WhereSQL + ' AND BQ_DEVISE = "' + gcbo_Devise.Value + '"';
    end;
  end;
end ;

{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... :
Modifié le ... : 03/02/2004
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPETATPOINTAGE.OnLoad ;
var lQuery : TQuery;
begin
  Inherited ;
  lQuery := nil;
  try
    try
      // Si aucune donnée renseignée
      if VH^.PointageJal then
        lQuery := OpenSQL('SELECT MIN(J_JOURNAL) MINI, MAX(J_JOURNAL) MAXI FROM JOURNAL WHERE J_NATUREJAL="BQE"',True)
      else
        lQuery := OpenSQL('SELECT MIN(G_GENERAL) MINI, MAX(G_GENERAL) MAXI FROM GENERAUX WHERE G_POINTABLE ="X"',True);

      if not lQuery.Eof then
      begin
        if E_General.Text = '' then
          E_General.Text := lQuery.FindField('MINI').AsString;
        if E_General_.Text = '' then
         E_General_.Text := lQuery.FindField('MAXI').Asstring;
      end;
    except
      on E: Exception do PgiError('Erreur de requête SQL : ' + E.Message, 'Procédure : OnLoad');
    end;
  finally
    Ferme( lQuery );
  end;
end ;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... :
Modifié le ... : 05/06/2003
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPETATPOINTAGE.OnArgument (S : String ) ;
begin
  Inherited ;
  Ecran.HelpContext := 7616010;

  E_General        := THEdit(GetControl('EGENERAL',True));
  E_General_       := THEdit(GetControl('EGENERAL_',True));
  
  E_DatePointage   := THEdit(GetControl('E_DATEPOINTAGE',True));
  E_DatePointage_  := THEdit(GetControl('E_DATEPOINTAGE_',True));
  XX_Rupture       := THEdit(GetControl('XX_RUPTURE',True));
  TE_GENERAL       := THLabel(GetControl('TE_GENERAL', True));
  TE_GENERAL_      := THLabel(GetControl('TE_GENERAL_', True));

  gchk_Nature      := TCheckBox(GetControl('CHKNATURE',True));
  gcbo_Banque      := THValComboBox(GetControl('BQBANQUE',True));

  gchk_RuptureEtab := TCheckBox(GetControl('CHKRUPTUREPARETAB',True));
  gcbo_Devise      := THValComboBox(GetControl('BQ_DEVISE',True));

  gchk_Nature.OnClick  := chk_NatureClick;

  gcbo_Devise.ItemIndex := 0;
  gcbo_Devise.Enabled   := not VH^.PointageJal;
   
  E_General.OnElipsisClick      := OnElipsisClickE_General;
  E_General_.OnElipsisClick     := OnElipsisClickE_General;
  E_General.OnExit              := OnExitE_General;
  E_DatePointage.OnElipsisClick  := gtxt_DatePointageClick;
  E_Datepointage_.OnElipsisClick := gtxt_DatePointageClick;

  // Données par défaut
  gcbo_Banque.ItemIndex := 0;

  chk_NatureClick(nil);

  // GC - 18/06/2003
  if VH^.PointageJal then
  begin
    E_General.MaxLength  := 3;
    E_General_.MaxLength := 3;
    TE_GENERAL.Caption   := TraduireMemoire('Journaux de');
    TE_GENERAL_.Caption  := TraduireMemoire('à');
    XX_RUPTURE.Text      := 'E_JOURNAL';
    Gcbo_Banque.Enabled  := False;
    Gchk_Nature.Visible  := False;
    Gchk_Nature.Checked  := False;
  end
  else
  begin
    E_General.MaxLength  := VH^.CPta[fbGene].Lg;
    E_General_.MaxLength := VH^.CPta[fbGene].Lg;
    TE_GENERAL.Caption   := TraduireMemoire('Comptes généraux du');
    TE_GENERAL_.Caption  := TraduireMemoire('au');
    XX_RUPTURE.Text      := 'E_GENERAL';
  end;

end ;

procedure TOF_CPETATPOINTAGE.OnClose ;
begin
  Inherited ;
end ;

procedure TOF_CPETATPOINTAGE.chk_NatureClick(Sender: TObject);
var C : THEdit;
begin
  // Si on décoche les compte de nature banque
  if gcbo_Banque <> nil then
  begin
    // Grise les combos concernant les comptes de nature banque
    gcbo_Banque.Enabled := (gchk_Nature.Checked = True);
    gchk_RuptureEtab.Enabled := (gchk_Nature.Checked = True);
    gcbo_Devise.Enabled   := (gchk_Nature.Checked = True);
    if gchk_Nature.Checked = False then
      gcbo_Devise.ItemIndex := 0;
  end;

  C := THEdit(GetControl('G_NATUREGENE',True));
  if (c <> nil) then
    if gchk_Nature.Checked then
      C.Operateur := Egal
    else
    begin
      C.Operateur := Different;
      gchk_RuptureEtab.Checked := False;
    end;
end;

procedure TOF_CPETATPOINTAGE.gtxt_DatePointageClick(Sender: TObject);
var
  szCpte1,szCpte2,szWhere : String;
  lStSql : string;
begin
  szCpte1 := GetControlText('EGENERAL');
  szCpte2 := GetControlText('EGENERAL_');
  if (szCpte1<>'') then
    szWhere := 'EE_GENERAL>="' + szCpte1 + '"';

  if (szCpte2<>'') then
  begin
    if (szWhere<>'') then
      szWhere := szWhere + ' AND ';
    szWhere := szWhere + 'EE_GENERAL<="' + szCpte2 + '"';
  end;

  lStSql := 'SELECT DISTINCT EE_DATEPOINTAGE, EE_REFPOINTAGE FROM EEXBQ ' +
            'WHERE ' + szWhere;  // GCO - 27/10/2006

            //'WHERE ' + szWhere + ' ORDER BY EE_DATEPOINTAGE DESC';
            // GCO - 15/05/2006 - FQ 13289

  // Affiche le lookup
  LookUpList(THEdit(Sender), 'Date de pointage', 'EEXBQ', 'EE_DATEPOINTAGE', '',
    '', 'EE_DATEPOINTAGE DESC' , True, 0, lStSql );
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 18/06/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPETATPOINTAGE.OnElipsisClickE_General(Sender: TObject);
var lSt : string;
begin
  if VH^.PointageJal then
  begin
    lSt := 'SELECT J_JOURNAL, J_LIBELLE FROM JOURNAL LEFT JOIN GENERAUX ON J_CONTREPARTIE=G_GENERAL ' +
           'WHERE J_NATUREJAL="BQE" AND G_POINTABLE="X"';

    LookUpList(THEdit(Sender), 'Journal', 'JOURNAL', 'J_JOURNAL', 'J_LIBELLE', '', 'J_JOURNAL', True, 0 , lSt)
  end
  else
    LookUpList(THEdit(Sender), 'Compte général', 'GENERAUX', 'G_GENERAL', 'G_LIBELLE', 'G_POINTABLE="X"', 'G_GENERAL', True, 0 );
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 05/06/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPETATPOINTAGE.OnExitE_General(Sender: TObject);
begin
  if (Trim(E_General.Text) = '') then Exit;

  if (not VH^.PointageJal) and (length(E_General.Text) <> VH^.CPta[fbGene].Lg) then
    E_General.Text := BourreEtLess(E_General.Text, fbGene);
end;

////////////////////////////////////////////////////////////////////////////////

Initialization
  registerclasses ( [ TOF_CPETATPOINTAGE ] ) ;
end.
