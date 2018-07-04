{***********UNITE*************************************************
Auteur  ...... : Stéphane BOUSSERT
Créé le ...... : 08/07/2002
Modifié le ... : 23/08/2005
Description .. : Source TOF de la FICHE : CPBALGEN ()
Suite ........ :
Suite ........ : Gestion du QRS1 des balances générales
Suite ........ :
Suite ........ : Depuis 25/07/2003 : les traitements communs aux
Suite ........ : balances ont été remontés dans une tof intermédiaire :
Suite ........ : TOF_METH --> TOF_CPBALANCE --> TOF_CPBALAUXI
Suite ........ :
Suite ........ : JP 01/07/05 : Gestion des caractères Joker : fonctions de
Suite ........ : base définies dans TofMeth
Suite ........ :
Suite ........ : GCO - 23/08/2005 - FQ 16404
Suite ........ :
Suite ........ : JP 03/10/07 : FQ 21560 : gestion de l'appel depuis
Suite ........ : CPTOTRUB_TOF
Suite ........ :
Mots clefs ... : TOF;CPBALGEN
*****************************************************************}

Unit CPBALGEN_TOF ;

Interface

Uses StdCtrls, Classes,
{$IFDEF EAGLCLIENT}
     MainEAgl, utob, eQRS1,
{$ELSE}
     db,
     {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}
     Fe_Main,
     QRS1,
{$ENDIF}
     sysutils, Spin, Ent1,
     HCtrls, UTOF,
     TofMeth,
     CPBALANCE_TOF,
     uLibExercice,  // CInitComboExercice(E_EXERCICE);
     uLibWindows,   // TraductionTHMultiValComboBox
     AGLInit,       // TheData
     Critedt,       // ClassCritEdt
     ParamSoc ;     // GetParamSocSecur, GetParamSocSecurSecur


procedure CPLanceFiche_BalanceGeneral ( Args : String = '' ) ;

Type
  TOF_CPBALGEN = Class (TOF_CPBALANCE)
    // ---------------
    // Onglet options
    DetailCollectif             : TCheckBox;

    // Rupture Cycle de révision
    CycleRevision               : THEdit;
    RupCycleRevision            : TCheckBox;
    SautCycleRevision           : TCheckBox;


    // EVT TOF
    procedure OnNew                    ; override ;
    procedure OnUpdate                 ; override ;
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;

    // EVT FICHE
    procedure TableLibreChanged ( Sender : TObject ) ; override ;
    procedure CompteOnExit      ( Sender : TObject ) ; override ;
    procedure RuptureClick      ( Sender : TObject ) ; override ;

    //  -> MAJ du champ CED_NATURE pour le calcul des totaux bilan / gestion
    procedure MajNatureCompte ;
    procedure RemplirEDTBALANCE ; override ;

    procedure ChargementCritEdt ; override ;

    // Gestion rupture sur cylcle de révision
    procedure RupCycleRevisionClick( Sender : TObject );

    function  GetConditionRupture : String ; override ;
    function  GenererRequeteBAL   : String ; override ;

    function  GetJointureCpt           : String ; override ;
    function  GetCompte2               : string ; override ;
    function  GetLibelle2              : string ; override ;
    function  GetCollectif             : string ; override ;
    function  GetGroupBy               : string ; override ;
    function  GetConditionSQLCptSpecif : string ; override ;

    private
      FBalRevision : boolean;
      FClauseWhGene : string; {JP 03/10/07 : FQ 21560}
  end ;

Implementation


uses
  {$IFDEF MODENT1}
  CPTypeCons,
  {$ENDIF MODENT1}
  Hent1;

  
//==============================================================================
procedure CPLanceFiche_BalanceGeneral ( Args : String ) ;
begin
	AGLLanceFiche('CP','CPBALGEN','','',Args);
end;

{TOF_CPBALGEN}

procedure TOF_CPBALGEN.OnNew ;
begin
  Inherited ;
end ;

procedure TOF_CPBALGEN.OnUpdate ;
begin
  Inherited ;
end ;

procedure TOF_CPBALGEN.OnLoad ;
begin

  Inherited ;

  if fBalRevision  then
  begin
    TFQRS1(Ecran).NatureEtat := 'BAR';
    TFQRS1(Ecran).CodeEtat   := 'BAR';
    SetControlText('TitreCol1','En cours : '+FormatDateTime('yyyy',DateFinPer)) ;

    // GCO - 02/08/2007 - FQ 21229 - Pas de risque, on compare toujours sur le précédent
    if VH^.Precedent.Code = '' then
      SetControlText('TitreCol2', '')
    else
      SetControlText('TitreCol2','Précédent : '+FormatDateTime('yyyy',DateFinCompar)) ;
  end;

  {JP 01/07/05 : on ne fait que l'auto-complétion que s'il n'y a pas de caractère joker}
  if not TestJoker(CompteDe.Text) then begin
    // Auto-remplissage si comptes non renseignés
    if (Trim(CompteDe.Text) = '') then
      CompteDe.Text := GetMinCompte('GENERAUX', 'G_GENERAL', NatureCpt.Value );
    if (Trim(CompteA.Text) = '') then
      CompteA.Text := GetMaxCompte('GENERAUX', 'G_GENERAL', NatureCpt.Value );
  end;

  // Les ruptures...
  if Rupture.Value <> 'SANS' then
    begin
    // rupture sur n° compte
    if (RuptureType.Value = 'RUPGROUPES') and (NivoRupture.value > 0) then
      AvecNivoRupture.Text := IntToStr(NivoRupture.value);
    // rupture sur champ libre généraux
    if (RuptureType.Value = 'RUPLIBRES') and (TableLibre.ItemIndex >= 0) then
      XXRupture.Text := 'G_TABLE' + GetNumTableLibre ;
    // rupture sur plan de correspondance
    if (RuptureType.Value = 'RUPCORRESP') and (Corresp.ItemIndex >= 0) then
      XXRupture.Text := 'G_CORRESP' + IntToStr(Corresp.ItemIndex + 1);
    end;
end ;

procedure TOF_CPBALGEN.OnArgument (S : String ) ;
begin
  FBalRevision := ( S = 'BAR' );
  if FBalRevision then
    TFQRS1(Ecran).FNomFiltre   := 'CPBALGENREV';

  // CONTROLES SPECIFIQUES
  // Rupture sur cycle de révision
  CycleRevision     := THEdit(GetControl('GCYCLEREVISION', True));
  RupCycleRevision  := TCheckBox(GetControl('RUPCYCLEREVISION', True));
  SautCycleRevision := TCheckBox(GetControl('SAUTCYCLEREVISION', True));


  // Controles communs
  Inherited ;

  // Rubrique d'aide
  Ecran.HelpContext:=7445000;

  // Options d'éditions
  DetailCollectif := TCheckBox(GetControl('DETAILCOLLECTIF', true));

  // Type de plan comptable :
  Case V_PGI.LaSerie of
    S7 :  begin
          Corresp.plus := 'AND (CO_CODE = "GE1" OR CO_CODE = "GE2")';
          end;
    else  begin
          Corresp.plus := 'AND CO_CODE = "GE1"';
          end;
    end;

  // GCO 08/08/2003
  CompteDe.MaxLength := VH^.CPta[fbGene].Lg;
  CompteA.MaxLength  := VH^.CPta[fbGene].Lg;

  // Type de balance de situation
  BalSit.Plus := ' BSI_TYPECUM = "GEN" ' ;

  // Gestion rupture sur cyle de révision
  // GCO - 23/08/2005 - FQ 16404
  RupCycleRevision.Enabled := VH^.Revision.Plan <> '';
  RupCycleRevision.OnClick := RupCycleRevisionClick ;
  RupCycleRevisionClick(nil);

  CycleRevision.Plus := ' AND CCY_EXERCICE = "' + VH^.EnCours.Code + '"';
end ;

//==============================================================================
//====================        EVT QRS1        ==================================
//==============================================================================

procedure TOF_CPBALGEN.TableLibreChanged(Sender: TObject);
begin
  if TableLibre.ItemIndex < 0 then
    begin
    LibreDe.DataType 	:= '';
    LibreA.DataType	:= '';
    LibreDe.Text 	:= '';
    LibreA.Text 	:= '';
    end
  else
    begin
    LibreDe.DataType 	:= 'TZNATGENE' + GetNumTableLibre ;
    LibreA.DataType	:= 'TZNATGENE' + GetNumTableLibre ;
    LibreDe.Text 	:= '';
    LibreA.Text 	:= '';
    end;
end;

//==============================================================================
//====================   TRAITEMENTS DES DONNEES  ==============================
//==============================================================================

procedure TOF_CPBALGEN.CompteOnExit(Sender: TObject);
begin
  if ( csDestroying in Ecran.ComponentState) then Exit ;
  if Trim(THEdit(Sender).Text) = '' then Exit;

  {JP 01/07/05 : on ne fait que l'auto-complétion que s'il n'y a pas de caractère joker}
  if HasJoker(Sender) then Exit;

  // Complétion auto du numéro de compte si possible
  if not CompleteAuto( Sender, fbGene ) then
    THEdit(Sender).ElipsisClick(nil);
end;

procedure TOF_CPBALGEN.MajNatureCompte;
var lQCpt          : TQuery ;
    lStCompte      : String ;
    lStNature      : String ;
    lStSO_FermeBil : String ;
begin
  lStSO_FermeBil := GetParamSocSecur('SO_FERMEBIL','') ;
  lQCpt := OpenSQL( 'SELECT CED_COMPTE FROM ' + GetTablePourBase('CEDTBALANCE')
                                       + ' WHERE CED_USER="' + V_PGI.User + '"', True ) ;
  While not lQCpt.Eof do
    begin
    lStCompte := lQCpt.FindField('CED_COMPTE').AsString ;
    lStNature := '' ;

    // Test Compte Bilan
    if ( ( lStCompte >= VH^.FBil[1].Deb ) and ( lStCompte <= VH^.FBil[1].Fin ) ) or // 1er groupe Bilan
       ( ( lStCompte >= VH^.FBil[2].Deb ) and ( lStCompte <= VH^.FBil[2].Fin ) ) or // 2ème groupe Bilan
       ( ( lStCompte >= VH^.FBil[3].Deb ) and ( lStCompte <= VH^.FBil[3].Fin ) ) or // 3ème groupe Bilan
       ( ( lStCompte >= VH^.FBil[4].Deb ) and ( lStCompte <= VH^.FBil[4].Fin ) ) or // 4ème groupe Bilan
       ( ( lStCompte >= VH^.FBil[5].Deb ) and ( lStCompte <= VH^.FBil[5].Fin ) ) or // 5ème groupe Bilan
       ( lStCompte = VH^.OuvreBil ) or ( lStCompte = lStSO_FermeBil )   // Compte Bilan (ouverture / fermeture)
       then lStNature := 'BIL'
    // Test Compte Charge
    else if ( ( lStCompte >= VH^.FCha[1].Deb ) and ( lStCompte <= VH^.FCha[1].Fin ) ) or // 1er groupe Bilan
            ( ( lStCompte >= VH^.FCha[2].Deb ) and ( lStCompte <= VH^.FCha[2].Fin ) ) or // 2ème groupe Bilan
            ( ( lStCompte >= VH^.FCha[3].Deb ) and ( lStCompte <= VH^.FCha[3].Fin ) ) or // 3ème groupe Bilan
            ( ( lStCompte >= VH^.FCha[4].Deb ) and ( lStCompte <= VH^.FCha[4].Fin ) ) or // 4ème groupe Bilan
            ( ( lStCompte >= VH^.FCha[5].Deb ) and ( lStCompte <= VH^.FCha[5].Fin ) )    // 5ème groupe Bilan
       then lStNature := 'CHA'
    // Test Compte Produit
    else if ( ( lStCompte >= VH^.FPro[1].Deb ) and ( lStCompte <= VH^.FPro[1].Fin ) ) or // 1er groupe Bilan
            ( ( lStCompte >= VH^.FPro[2].Deb ) and ( lStCompte <= VH^.FPro[2].Fin ) ) or // 2ème groupe Bilan
            ( ( lStCompte >= VH^.FPro[3].Deb ) and ( lStCompte <= VH^.FPro[3].Fin ) ) or // 3ème groupe Bilan
            ( ( lStCompte >= VH^.FPro[4].Deb ) and ( lStCompte <= VH^.FPro[4].Fin ) ) or // 4ème groupe Bilan
            ( ( lStCompte >= VH^.FPro[5].Deb ) and ( lStCompte <= VH^.FPro[5].Fin ) )    // 5ème groupe Bilan
       then lStNature := 'PRO'
       // Sinon Extra-comptable
       else lStNature := 'EXT' ;

    // Update CED_NATURE
    ExecuteSQL('UPDATE ' + GetTablePourBase('CEDTBALANCE') + ' SET CED_NATURE="' + lStNature
                         + '" WHERE CED_COMPTE="' + lStCompte + '" AND CED_USER="' + V_PGI.User + '"' ) ;

    // Au suivant
    lQCpt.Next ;

    end ;

  Ferme( lQCpt ) ;

end;

procedure TOF_CPBALGEN.RemplirEDTBALANCE ;
begin
  inherited;

  // MAJ de la nature des compte pour le calcul des totaux bilan / gestion
  if GetControlText('BILAN')='X' then
    MajNatureCompte ;

  // Prise en compte des comptes ferme
  // FQ 17430
  if GetControl('GENEFERMES') = nil then Exit;
  if ExecuteSQL('DELETE FROM CEDTBALANCE WHERE CED_USER = "' + V_PGI.User + '" ' + 
                'AND ((CED_DEBIT1=0 AND CED_CREDIT1=0) AND (CED_DEBIT2=0 AND CED_CREDIT2=0)) ' +
                'AND ((CED_COMPTE =(SELECT G_GENERAL FROM GENERAUX WHERE G_GENERAL = CED_COMPTE AND G_FERME = "X")) ' +
                'OR  (CED_COMPTE2=(SELECT G_GENERAL FROM GENERAUX WHERE G_GENERAL = CED_COMPTE2 AND G_FERME = "X")) ' +
                'AND CED_COMPTE="")') > 0 then
      THEdit(GetControl('GENEFERMES')).Text := 'X'
  else
      THEdit(GetControl('GENEFERMES')).Text := '-';   
end;

procedure TOF_CPBALGEN.RupCycleRevisionClick(Sender: TObject);
begin
  CycleRevision.Enabled     := (RupCycleRevision.Checked);
  SautCycleRevision.Enabled := (RupCycleRevision.Checked);
  if RupCycleRevision.Checked then
    NivoRupture.Value := 0 ;
end;

function TOF_CPBALGEN.GenererRequeteBAL: String;
begin
  // Champs

{  if FBoRegroupement and ( GetControlText('MULTIDOSSIER') <> '' )
    then result := 'SELECT SYSDOSSIER, CED_COMPTE CED_GENERAL, CED_NATURE, CED_LIBELLE'
    else }
  result := 'SELECT CED_COMPTE CED_GENERAL, CED_NATURE, CED_LIBELLE' ;

  result := result
               + ', CED_COMPTE2 CED_CYCLE, CED_LIBELLE2 CED_CYCLELIB'
               + ', CED_DEBIT1, CED_CREDIT1'
               + ', CED_DEBIT2, CED_CREDIT2'
               + ', CED_RUPTURE, CED_COLLECTIF ' ;

  // Récupération du sens dans le cas des balances de révision
  if fBalRevision  then
  begin
    result := result + ', G_SENS ';
    // Solde en cours en fonction du sens
    result := result + ', IIF(((G_SENS="D") OR (G_SENS="M")) , CED_DEBIT1-CED_CREDIT1, CED_CREDIT1-CED_DEBIT1) SOLDEE ';
    // Solde précédent en fonction du sens
    result := result + ', IIF(((G_SENS="D") OR (G_SENS="M")) , CED_DEBIT2-CED_CREDIT2, CED_CREDIT2-CED_DEBIT2) SOLDEP ';
  end;

  if Rupture.Value <> 'SANS' then
    begin

    // Rupture sur plan de correspondance
    if (RuptureType.Value = 'RUPCORRESP') and (Corresp.ItemIndex >= 0) then
      result := result + ', CR_LIBELLE RUPTURELIB '
    // Rupture sur tables libres
    else if (RuptureType.Value = 'RUPLIBRES') and (TableLibre.ItemIndex >= 0) then
      result := result + ', NT_LIBELLE RUPTURELIB '
    // pas de rupture
    else result := result + ', "&#@" RUPTURELIB ' ;

    end
  // pas de rupture
  else result := result + ', "&#@" RUPTURELIB ' ;

  // Tables
  result := result + ' FROM CEDTBALANCE ' ;
  if Rupture.Value <> 'SANS' then
    begin
    // Rupture classique
    if (RuptureType.Value = 'RUPCORRESP') and (Corresp.ItemIndex >= 0) then
      result := result + ' LEFT JOIN CORRESP ON CR_CORRESP=CED_RUPTURE '
                       + 'AND CR_TYPE="' + Corresp.Value + '"'
    else if (RuptureType.Value = 'RUPLIBRES') and (TableLibre.ItemIndex >= 0) then
      result := result + ' LEFT JOIN NATCPTE ON NT_NATURE=CED_RUPTURE '
                       + 'AND NT_TYPECPTE="' + TableLibre.Value + '"' ;
    end ;
  if fBalRevision  then
  begin
    // Ajout jointure sur la table des généraux pour récupérer le sens du compte
    result := result + ' LEFT JOIN GENERAUX ON G_GENERAL=CED_COMPTE '
  end;
  // Conditions
  result := result + ' WHERE CED_USER = "' + V_PGI.User + '"' ;

  // Tri
  if RupCycleRevision.Checked
    then result := result + ' ORDER BY CED_CYCLE, CED_RUPTURE, CED_GENERAL'
    else result := result + ' ORDER BY CED_RUPTURE, CED_GENERAL' ;


end;

procedure TOF_CPBALGEN.RuptureClick(Sender: TObject);
begin
  inherited;
  if Rupture.value <> 'SANS' then
    RupCycleRevision.Checked := False ;
end;

function TOF_CPBALGEN.GetConditionRupture : String ;
begin

  result := Inherited GetConditionRupture ;

  // Rupture spéciale sur cycle de révision
  if (RupCycleRevision.Checked) and (CycleRevision.Text <> '') then
  begin
    if TestJoker(CycleRevision.Text) then // GCO - 14/09/2007 - FQ 20913
      Result := Result + ' AND ' + ClauseAvecJoker(CycleRevision.Text, 'G_CYCLEREVISION')
    else
      Result := Result + ' AND G_CYCLEREVISION LIKE "' + CycleRevision.Text + '%"';
  end;

end;

////////////////////////////////////////////////////////////////////////////////
{***********A.G.L.***********************************************
Auteur  ...... : Gilles COSTE
Créé le ...... : 13/09/2005
Modifié le ... :   /  /
Description .. :
Mots clefs ... :
*****************************************************************}
procedure TOF_CPBALGEN.ChargementCritEdt;
begin
  if (TheData <> nil) and (TheData is ClassCritEdt) then
  begin
    {JP 03/10/07 : FQ 21560 : Gestion de l'appel depuis CPTOTRUB_TOF}
    FClauseWhGene := ClassCritEdt(TheData).CritEdt.FiltreSup;
    if FClauseWhGene <> '' then begin
      //NouvRechClick(nil);
      AvecComparatif.Checked := False;
      ClassCritEdt(TheData).CritEdt.FiltreSup := '';
    end;

    if (ClassCritEdt(TheData).CritEdt.BAL.CycleDeRevision <> '') and
       (RupCycleRevision <> nil) and
       (CycleRevision <> nil) then
    begin
      Rupture.ItemIndex := 1;
      RupCycleRevision.Checked := True;
      CycleRevision.Text := ClassCritEdt(TheData).CritEdt.BAL.CycleDeRevision;
    end;
  end;
  inherited;
end;

////////////////////////////////////////////////////////////////////////////////

function TOF_CPBALGEN.GetJointureCpt: String;
begin
  // Rupture sur révision
  if RupCycleRevision.Checked then
    result := result + ' LEFT JOIN CREVCYCLE ON ' +
              'CCY_CODECYCLE = G_CYCLEREVISION AND ' +
              'CCY_EXERCICE = "' + VH^.EnCours.Code + '"';
end;

function TOF_CPBALGEN.GetCompte2: string;
begin
  if RupCycleRevision.Checked
    then result := 'CCY_CODECYCLE'
    else result := '""';
end;

function TOF_CPBALGEN.GetLibelle2: string;
begin
  // GCO - 02/08/2007 - FQ 21227
  if RupCycleRevision.Checked then
    result := 'SUBSTRING(CCY_LIBELLECYCLE, 1,35)'
  else
    result := '""';
end;

function TOF_CPBALGEN.GetGroupBy: string;
begin
  result := inherited GetGroupBy;

  if RupCycleRevision.Checked
    then result := result + ',CCY_CODECYCLE, CCY_LIBELLECYCLE' ;

  if DetailCollectif.Checked
    then result := result + ',G_COLLECTIF' ;

end;

function TOF_CPBALGEN.GetCollectif: string;
begin
  if DetailCollectif.Checked
    then result := 'G_COLLECTIF'
    else result := '"-"';
end;

{JP 03/10/07 : FQ 21560 : depuis CPTOTRUB_TOF, on n'a pas une fourchette de comptes, mais
               une clause where avec des LIKE, BETWEEN. Cette fonction surchargée et appelée
               dans GetConditionSQLCpt permet de courcircuiter le traitement standard}
{---------------------------------------------------------------------------------------}
function TOF_CPBALGEN.GetConditionSQLCptSpecif : string;
{---------------------------------------------------------------------------------------}
begin
  Result := inherited GetConditionSQLCptSpecif;
  if Trim(FClauseWhGene) <> '' then Result := FClauseWhGene;
end;

Initialization
  registerclasses ( [ TOF_CPBALGEN ] ) ;
end.

