{***********UNITE*************************************************
Auteur  ...... :
Créé le ...... : 19/07/2002
Modifié le ... : 27/02/2004
Description .. : Source TOF de la FICHE : CPRAPPRO ()
Mots clefs ... : TOF;CPRAPPRO
*****************************************************************}
Unit CPRAPPRO_TOF ;

Interface

Uses StdCtrls, Controls, Classes,
{$IFDEF EAGLCLIENT}
     UTOB,
     eQRS1,
     MaineAGL,
{$ELSE}
     db,
     {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}
     QRS1,
     FE_Main,
{$ENDIF}
     forms, sysutils, ComCtrls, HCtrls, HEnt1, HMsgBox, UTOF, LookUp, Ent1,
     uLibWindows,
     CpteSav,    // CTrouveContrePartie
     Saisutil;

procedure CC_LanceFicheEtatRappro;

Type
  TOF_CPRAPPRO = Class (TOF)
    procedure OnNew                    ; override ;
    procedure OnDelete                 ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure OnClose                  ; override ;

  private
    glbl_RefPointage   : THLabel;
    gtxt_DatePointage  : THEdit;
    gtxt_General       : THEdit;

    gtxt_Sens          : THEdit;
    grbt_SensEtat      : THRadioGroup;

    XX_Rupture         : THEdit;
    TE_General         : THLabel;

    AFFRecapitulatif   : TCheckBox;

    FStCompteUtiliser : string;
    
    procedure OnElipsisClickDatePointage ( Sender : TObject );
    procedure grbt_SensEtatClick         ( Sender : TObject );

    procedure OnExitGeneral              ( Sender : TOBject );
    procedure OnExitE_DatePointage       ( Sender : TOBject );
    procedure OnElipsisClickE_General    ( Sender : TObject );

    procedure RecupereRefPointage ( vLaDerniere : Boolean );
  end ;

Implementation

procedure CC_LanceFicheEtatRappro;
begin
  AGLLanceFiche('CP','CPRAPPRO','','','');
end;

procedure TOF_CPRAPPRO.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_CPRAPPRO.OnDelete ;
begin
  Inherited ;
end ;

procedure TOF_CPRAPPRO.OnUpdate ;
var lQuery : TQuery;
    lStDevise : string;
    lStControleDevise : string;
begin
  Inherited ;

  TFQRS1(Ecran).WhereSQL := 'SELECT E_GENERAL, E_JOURNAL, E_DATECOMPTABLE, '+
                            'E_NUMEROPIECE, E_REFINTERNE, E_LIBELLE, E_EXERCICE, '+
                            'E_NUMLIGNE, E_DEVISE, E_REFPOINTAGE, E_DATEPOINTAGE, '+
                            'E_DEBIT, E_CREDIT, E_DEBITDEV, E_CREDITDEV, '+
                            'BQ_DEVISE, BQ_DOMICILIATION, BQ_ETABBQ, BQ_GUICHET, '+
                            'BQ_NUMEROCOMPTE, BQ_CLERIB, G_LIBELLE, J_JOURNAL, '+
                            'J_LIBELLE FROM ECRITURE ';

  // Modifie la clause Where
  if VH^.PointageJal then
  begin
    TFQRS1(Ecran).WhereSQL := TFQRS1(Ecran).WhereSQL +
                              ' LEFT JOIN JOURNAL ON J_JOURNAL = E_JOURNAL' +
                              ' LEFT JOIN GENERAUX ON G_GENERAL = J_CONTREPARTIE' +
                              ' LEFT JOIN BANQUECP ON BQ_GENERAL = J_CONTREPARTIE AND BQ_NODOSSIER="'+V_PGI.NoDossier+'"'; // 19/10/2006 YMO Multisociétés
  end
  else
  begin
    TFQRS1(Ecran).WhereSQL := TFQRS1(Ecran).WhereSQL +
                              ' LEFT JOIN JOURNAL ON J_JOURNAL = E_JOURNAL' +
                              ' LEFT JOIN GENERAUX ON G_GENERAL = E_GENERAL'+
                              ' LEFT JOIN BANQUECP ON BQ_GENERAL = E_GENERAL AND BQ_NODOSSIER="'+V_PGI.NoDossier+'"'; // 19/10/2006 YMO Multisociétés
  end;

  TFQRS1(Ecran).WhereSQL := TFQRS1(Ecran).WhereSQL + ' WHERE E_QUALIFPIECE="N" AND (E_ECRANOUVEAU="N" OR E_ECRANOUVEAU="H")';

  // WHERE particuler si sur JOURNAL ou sur COMPTE
  if VH^.PointageJal then
  begin
    TFQRS1(Ecran).WhereSQL := TFQRS1(Ecran).WhereSQL + ' AND E_JOURNAL = "' + gtxt_general.Text + '"';
    TFQRS1(Ecran).WhereSQL := TFQRS1(Ecran).WhereSQL + ' AND E_GENERAL <> J_CONTREPARTIE';
  end
  else
  begin
    TFQRS1(Ecran).WhereSQL := TFQRS1(Ecran).WhereSQL + ' AND E_GENERAL = "' + gtxt_general.Text + '"';
    TFQRS1(Ecran).WhereSQL := TFQRS1(Ecran).WhereSQL + ' AND G_POINTABLE = "X"';
  end;

  // Gestion de l'ExoV8
  if VH^.ExoV8.Code <> '' then
    TFQRS1(Ecran).WhereSQL := TFQRS1(Ecran).WhereSQL + ' AND E_DATECOMPTABLE >= "' + UsDateTime(VH^.ExoV8.Deb) + '"';

  TFQRS1(Ecran).WhereSQL := TFQRS1(Ecran).WhereSQL + ' AND E_DATECOMPTABLE <= "' + USDateTime(StrToDate(gtxt_DatePointage.Text)) + '"';
  
  if not AffRecapitulatif.Checked then
    TFQRS1(Ecran).WhereSQL := TFQRS1(Ecran).WhereSQL + ' AND ((E_REFPOINTAGE <> "" and E_DATEPOINTAGE > "' + USDateTime(StrToDate(gtxt_DatePointage.Text)) + '" ) or (E_REFPOINTAGE = ""))';

  // Récupération de la devise du compte pour afficher que les écritures de cette devise
  lQuery := nil;
  try
    try
      lQuery := OpenSql('SELECT BQ_DEVISE FROM BANQUECP WHERE BQ_GENERAL = "' + FStCompteUtiliser
                       +'" AND BQ_NODOSSIER="'+V_PGI.NoDossier+'"', True); // 19/10/2006 YMO Multisociétés
      if not lQuery.Eof then
        lStDevise := lQuery.FindField('BQ_DEVISE').AsString;
    except
      on E : Exception do PgiError( 'Erreur de requête SQL : ' + E.Message, 'Procédure : OnUpdate' );
    end;
  finally
    Ferme( lQuery );
  end;

  // On réutilise la requete de l'état pour vérifier la présence des écritures avec devise différente
  lStControleDevise := TFQRS1(Ecran).WhereSQL;

  // Devise des écritures à afficher
  if lStDevise <> '' then
    TFQRS1(Ecran).WhereSQL := TFQRS1(Ecran).WhereSQL + ' AND E_DEVISE = "' + lStDevise + '"';

  lStControleDevise := lStControleDevise + ' AND E_DEVISE <> "' + lStDevise + '"';

  TFQRS1(Ecran).WhereSQL := TFQRS1(Ecran).WhereSQL + ' ORDER BY E_DATECOMPTABLE, E_NUMEROPIECE';
  lStControleDevise := lStControleDevise + ' ORDER BY E_DATECOMPTABLE, E_NUMEROPIECE';

  // Test d'existence des ecritures sur d'autres DEVISES que celle du compte de banque
  if ExisteSql( lStControleDevise ) then
  begin
    if VH^.PointageJal then
      PgiInfo('Attention  : certaines écritures non pointées du journal ' + gtxt_general.Text + ' ont ' +
              'une monnaie de saisie différente de ' + lStDevise , Ecran.Caption)
    else
      PgiInfo('Attention  : certaines écritures non pointées du compte ' + gtxt_general.Text + ' ont ' +
              'une monnaie de saisie différente de ' + lStDevise , Ecran.Caption);
  end;
end ;

procedure TOF_CPRAPPRO.OnLoad ;
var
  Q : TQuery;
  lStSql : string;
  DatePointage : TDateTime;
  Solde, SoldeDev  : Double;
  SoldeB, SoldeDevB : Double;
begin
  Inherited ;

  // On détermine sur quel compte GENERAL il faut travailler
  if VH^.PointageJal then
    FStCompteUtiliser := CTrouveContrePartie( gtxt_general.Text )
  else
    FStCompteUtiliser := gtxt_general.Text;


  // Met une date par défaut si la date est incorrecte
  if (gtxt_DatePointage.Text = '') or not IsValidDate(gtxt_DatePointage.Text) then
    RecupereRefPointage( True );

  // Récupération du nom de la référence de pointage
  OnExitE_DatePointage( nil );

  // Pour affichage dans l'état
  // Solde Bancaire
  try
    try
      Q := OpenSQL('SELECT EE_NEWSOLDEDEBEURO - EE_NEWSOLDECREEURO, EE_NEWSOLDEDEB - EE_NEWSOLDECRE FROM EEXBQ ' +
                   'WHERE EE_GENERAL = "' + gtxt_General.Text + '" ' +
                   'AND EE_REFPOINTAGE = "' + glbl_RefPointage.Caption + '" '+
                   'AND EE_DATEPOINTAGE = "' + USDateTime(StrToDate(gtxt_DatePointage.Text)) + '"' +
                   'ORDER BY EE_GENERAL', True);

      SoldeB     := Q.Fields[0].AsFloat;
      SoldeDevB  := Q.Fields[1].AsFloat;

      if gtxt_Sens.Text = 'Comptable' then
      begin
        THEdit(GetControl('SOLDEB',True)).Text    := FloatToStr(SoldeB);
        THEdit(GetControl('SOLDEDEVB',True)).Text := FloatToStr(SoldeDevB);
      end
      else
      begin
        THEdit(GetControl('SOLDE',True)).Text    := FloatToStr(SoldeB * -1);
        THEdit(GetControl('SOLDEDEV',True)).Text := FloatToStr(SoldeDevB * -1);
      end;

    except
      on E : Exception do PgiError( 'Erreur de requête SQL : ' + E.Message, 'Procédure : OnLoad' );
    end;
  finally
    Ferme(Q);
  end;

  // Solde comptable réel
  lStSql := 'SELECT SUM(E_CREDIT)-SUM(E_DEBIT), SUM(E_CREDITDEV)-SUM(E_DEBITDEV) ' +
            'FROM ECRITURE WHERE ';

  DatePointage := StrToDate(gtxt_DatePointage.Text);
  (*
  if QUELEXODT(DatePointage)=VH^.Suivant.Code then
  begin
    Q := OpenSQL('SELECT SUM(E_CREDIT)-SUM(E_DEBIT), SUM(E_CREDITDEV)-SUM(E_DEBITDEV) '+
                 'FROM ECRITURE WHERE E_GENERAL="'+gtxt_General.Text+'" AND E_QUALIFPIECE="N" ' +
                 'AND E_ECRANOUVEAU="N" AND E_DATECOMPTABLE>="'+USDateTime(GetDateDebExo(QUELEXODT(DatePointage)))+'" '+
                 'AND E_DATECOMPTABLE<="'+USDateTime(DatePointage)+'" AND E_EXERCICE="'+QUELEXODT(DatePointage)+'"' ,True);
    Solde     := SoldeB     + Q.Fields[0].AsFloat;
    SoldeDev  := SoldeDevB  + Q.Fields[1].AsFloat;
    Ferme(Q);
  end
  else
  begin *)
    lStSql := lStSql + 'E_GENERAL = "' + FStCompteUtiliser + '" AND E_QUALIFPIECE="N" ' +
                       'AND (E_ECRANOUVEAU="N" OR E_ECRANOUVEAU="H") ' +
                      // 'AND E_DATECOMPTABLE>="'+USDateTime(GetDateDebExo(QUELEXODT(DatePointage)))+'" '+
                       'AND E_DATECOMPTABLE<="'+USDateTime(DatePointage)+ '" ';

                       //+'" AND E_EXERCICE="'+QUELEXODT(DatePointage)+'"';

    if VH^.ExoV8.Code <> '' then
      lStSql := lStSql + 'AND E_DATECOMPTABLE >= "' + USDateTime(VH^.ExoV8.Deb) + '"';

    Q := OpenSQL( lStSql, True );
    Solde     := Q.Fields[0].AsFloat;
    SoldeDev  := Q.Fields[1].AsFloat;
    Ferme(Q);
  //end;

  if gtxt_Sens.Text ='Comptable' then
  begin
    THEdit(GetControl('SOLDE',True)).Text    := FloatToStr(Solde);
    THEdit(GetControl('SOLDEDEV',True)).Text := FloatToStr(SoldeDev);
  end
  else
  begin
    THEdit(GetControl('SOLDEB',True)).Text    := FloatToStr(-Solde);
    THEdit(GetControl('SOLDEDEVB',True)).Text := FloatToStr(-SoldeDev);
  end;
end ;

procedure TOF_CPRAPPRO.OnArgument (S : String );
begin
  Inherited ;
  Ecran.HelpContext := 7619000;

  gtxt_General                := THEdit(GetControl('EGENERAL',True));
  gtxt_General.OnExit         := OnExitGeneral;
  gtxt_General.OnElipsisClick := OnElipsisClickE_General;

  TE_General := THLabel(GetControl('TE_GENERAL', True));
  XX_Rupture := THEdit(GetControl('XX_RUPTURE', True));

  if VH^.PointageJal then
  begin
    TE_General.Caption     := 'Journal';
    gtxt_General.MaxLength := 3;
    XX_RUPTURE.Text        := 'E_JOURNAL';
  end
  else
  begin
    TE_General.Caption     := 'Compte général';
    gtxt_General.MaxLength := VH^.CPta[fbGene].Lg;
    XX_RUPTURE.Text        := 'E_GENERAL';
  end;

  gtxt_DatePointage := THEdit(GetControl('DATEPOINTAGE', True));
  gtxt_DatePointage.EditMask := '!99/99/0000;1;_';
  gtxt_DatePointage.OnElipsisClick := OnElipsisClickDatePointage;
  gtxt_DatePointage.OnExit := OnExitE_DatePointage;

  glbl_RefPointage := THLabel(GetControl('LBLDATEPOINTAGE',True));
  glbl_RefPointage.Caption := '';

  grbt_SensEtat := THRadioGroup(GetControl('RBTSENSETAT',True));
  grbt_SensEtat.OnClick := grbt_SensEtatClick;

  gtxt_Sens := THEdit(GetControl('SENS',True));
  grbt_SensEtatClick(nil);

  // GCO - 19/03/2004
  AFFRecapitulatif := TCheckBox(GetControl('AFFRECAPITULATIF', True));

  // Onglet de DEV visible en mode SAV
  TTabSheet(GetControl('TABDEV', True)).TabVisible := V_PGI.Sav;
end ;

procedure TOF_CPRAPPRO.OnClose ;
begin
  Inherited ;
end ;

////////////////////////////////////////////////////////////////////////////////
procedure TOF_CPRAPPRO.OnElipsisClickDatePointage(Sender: TObject);
begin
  // Affiche le lookup
  LookUpList(gtxt_DatePointage,Ecran.Caption,'EEXBQ','EE_DATEPOINTAGE','EE_REFPOINTAGE','EE_GENERAL="'+gtxt_General.Text+'"','EE_DATEPOINTAGE DESC',True,0);
end;

procedure TOF_CPRAPPRO.grbt_SensEtatClick(Sender: TObject);
begin
  // Pour l'état
  gtxt_Sens.Text := IIF( grbt_SensEtat.Value = '0', 'Comptable', 'Bancaire' );
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 05/06/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPRAPPRO.OnExitGeneral(Sender: TOBject);
begin
  if not VH^.PointageJal then
  begin
    if (Trim(gtxt_General.Text) <> '') and (length(gtxt_General.Text) <> VH^.CPta[fbGene].Lg) then
      gtxt_General.Text := BourreEtLess(gtxt_General.Text, fbGene);

    if ExisteSql('SELECT G_GENERAL FROM GENERAUX WHERE G_GENERAL = "' + Gtxt_General.Text + '"') then
    begin
      // Cherche la référence de pointage la plus récente
      if (gtxt_DatePointage.Text = DateToStr(iDate1900)) or not IsValidDate(gtxt_DatePointage.Text) then
        RecupereRefPointage( True);
    end
    else
    begin
      Gtxt_General.SetFocus;
      Gtxt_General.ElipsisClick( nil );
    end;
  end;
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 19/06/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPRAPPRO.OnElipsisClickE_General(Sender: TObject);
var lSt : String;
begin
  if VH^.PointageJal then
  begin
    lSt := 'SELECT J_JOURNAL, J_LIBELLE FROM JOURNAL LEFT JOIN GENERAUX ON J_CONTREPARTIE=G_GENERAL ' +
           'WHERE J_NATUREJAL="BQE" AND G_POINTABLE="X" ORDER BY J_JOURNAL';

    LookUpList(THEdit(Sender), 'Journal', 'JOURNAL', 'J_JOURNAL', 'J_LIBELLE', '', 'J_JOURNAL', True, 0 , lSt)
  end
  else
    LookUpList(THEdit(Sender), 'Compte général', 'GENERAUX', 'G_GENERAL', 'G_LIBELLE', 'G_NATUREGENE ="BQE"', 'G_GENERAL', True, 0 );
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 13/10/2003
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPRAPPRO.OnExitE_DatePointage(Sender: TOBject);
begin
  RecupereRefPointage( False );
end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 13/10/2003
Modifié le ... :   /  /    
Description .. : 
Mots clefs ... : 
*****************************************************************}
procedure TOF_CPRAPPRO.RecupereRefPointage ( vLaDerniere : Boolean );
var lQuery  : TQuery;
//    lStReference  : string;
//    lDate         : TDateTime;
//    lBoPointageOk : Boolean;
begin
//  lDate         := iDate1900;
//  lBoPointageOk := False;
  lQuery := nil;
  try
    try
      // Récupère la dernière référence du Compte sélectionné
//      lQuery := OpenSQL('SELECT EE_DATEPOINTAGE, EE_REFPOINTAGE, EE_AVANCEMENT FROM EEXBQ ' +
//                        'WHERE EE_GENERAL="' + gtxt_General.Text + '" ORDER BY EE_DATEPOINTAGE DESC', True);
//      if (not lQuery.Eof) then
//      begin
//        lStReference  := lQuery.FindField('EE_REFPOINTAGE').AsString;
///        lDate         := lQuery.FindField('EE_DATEPOINTAGE').AsDateTime;
//        lBoPointageOk := (lQuery.FindField('EE_AVANCEMENT').Asstring = 'X');
//      end;
//      Ferme(lQuery);

      if vLaDerniere then
        // Récupère la dernière référence du Compte sélectionné
        lQuery := OpenSQL('SELECT EE_DATEPOINTAGE, EE_REFPOINTAGE FROM EEXBQ ' +
                          'WHERE EE_GENERAL="' + gtxt_General.Text + '" ORDER BY EE_DATEPOINTAGE DESC', True)
      else
        // Récupère le libelle de la référence de pointage
        lQuery := OpenSQL('SELECT EE_DATEPOINTAGE, EE_REFPOINTAGE FROM EEXBQ ' +
                          'WHERE EE_GENERAL = "' + gtxt_General.Text + '" ' +
                          'AND EE_DATEPOINTAGE = "' + UsDateTime(StrToDate( gtxt_DatePointage.Text )) + '" ' +
                          'ORDER BY EE_GENERAL', True);

      if (not lQuery.EOF) then
      begin
//        AffRecapitulatif.Checked := False;

        if vLaDerniere then
          gtxt_DatePointage.Text := lQuery.FindField('EE_DATEPOINTAGE').AsString;

        glbl_RefPointage.Caption := lQuery.FindField('EE_REFPOINTAGE').AsString;

//        if (lStReference = lQuery.FindField('EE_REFPOINTAGE').AsString) and
//           (lDate        = lQuery.FindField('EE_DATEPOINTAGE').AsDateTime) and
//           (lBoPointageOk) then
//          AffRecapitulatif.Checked := True;
      end
      else
      begin
        if vLaDerniere then
          gtxt_DatePointage.Text := stDate1900;
        glbl_RefPointage.Caption := 'Aucune donnée';
      end;

    except
      on E : Exception do PgiError( 'Erreur de requête SQL : ' + E.Message, 'Procédure : RecupereRefPointage' );
    end;

  finally
    Ferme( lQuery );
  end;
end;

Initialization
  registerclasses ( [ TOF_CPRAPPRO ] ) ;
end.

