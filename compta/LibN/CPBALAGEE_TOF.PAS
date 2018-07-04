{***********UNITE*************************************************
Auteur  ...... : Stéphane BOUSSERT
Créé le ...... : 08/07/2002
Modifié le ... : 31/07/2003
Description .. : Source TOF de la FICHE : CPBALAGEE
Suite ........ :
Suite ........ : Gestion du QRS1 des balances auxiliaires
Suite ........ :
Suite ........ : Depuis 25/07/2003 : les traitements communs aux
Suite ........ : balances ont été remontés dans une tof intermédiaire :
Suite ........ : TOF_METH --> TOF_CPBALANCE --> TOF_CPBALAGEE
Suite ........ :
Suite ........ : JP 01/07/05 : Gestion des caractères Joker : fonctions de
Suite ........ : base définies dans TofMeth
Suite ........ :
Mots clefs ... : TOF;CPBALAGEE
*****************************************************************}
Unit CPBALAGEE_TOF;

Interface

Uses StdCtrls, Classes, Math,
{$IFDEF EAGLCLIENT}
     MainEAgl, utob, eQRS1,
{$ELSE}
     db,
     {$IFNDEF DBXPRESS}dbtables,{$ELSE}uDbxDataSet,{$ENDIF}
     Fe_Main, QRS1,
{$ENDIF}
     sysutils, Spin, Ent1,
     HCtrls, HEnt1, UTOF,
     TofMeth,
     hMsgBox,
     HQry,          // RecupWhereCritere
     HStatus,       // MoveCur
     AGLInit,       // TheData
     CritEdt,       // ClassCritEdt
     uLibExercice,  // CInitComboExercice(E_EXERCICE);
     CPBALANCE_TOF,
     Paramsoc ;

procedure CPLanceFiche_BalanceAgee ( Args : String = '' ) ;
procedure CPLanceFiche_BalanceVentilee ( Args : String = '' ) ;

Type
  TOF_CPBALAGEE = Class (TOF_CPBALANCE)

    ChoixTypeTaLi   : THRadioGroup;
    NatCptMulti     : ThMultiValComboBox ;

    Periodicite     : THValComboBox ;
    CollectifDe     : THEdit ;
    CollectifA      : THEdit ;

    DateArrete      : THEdit ;
    EnSituation     : TCheckBox ;

    Ecart           : THRadioGroup ;
    NbJEcart        : TSpinEdit ;

    TypePrevi       : TRadioButton ;
    TypeRetard      : TRadioButton ;

    // EVT TOF
    procedure OnLoad                   ; override ;
    procedure OnArgument (S : String ) ; override ;
    procedure ChargementCritEdt        ; override ;
    procedure SetTypeBalance           ; override ;

    // EVT FICHE
    procedure TableLibreChanged     (Sender : TObject);  override ;
    procedure CompteOnExit          (Sender: TObject) ;       override ;
    procedure EcartOnChange         (Sender : TObject) ;
    procedure NbJEcartOnChange      (Sender : TObject) ;
    procedure PeriodiciteOnChange   (Sender : TObject) ;
    procedure DateArreteOnChange    (Sender : TObject) ;
    procedure TypeBalVentilOnChange (Sender : TObject) ;
    procedure CollectifOnExit       (Sender: TObject) ; // FQ 18697
    procedure AuxiElipsisClick      (Sender: TObject) ;

    // INIT FORM
    procedure InitComposants ;  override ;
    procedure ParametreEtatType;

    // PARAM ETAT
    procedure ParamRuptures ;  override ;
    procedure ParamTitresCols ;  override ;
    procedure ParamComparatif ;  override ;
    procedure ParamDivers ;  override ;

    // NOUVELLES RUPTURES TABLES LIBRES TIERSCOMPL
    procedure InitChoixTableLibre;
    procedure OnClickChoixTypeTaLi  (Sender : TObject);
    function  GenererRequeteBAL : String; override ;
    Function  GetNumTableLibre : String ; override ;

    // SPECIF NATURE EN CHOIX MULTIPLE
    function  GetWhereNature : String ; override ;
    procedure InitZones ; override ;

    // GESTION DES PERIODES
    function  GetPeriodiciteIndex : Integer ;
    procedure CalculPeriodesAvecPeriodicite;
    procedure CalculPeriodesAvecNbJour;

    // TRAITEMENT
    // -> Procedure de calcul des totaux / soldes pour alimentation table temporaire
    procedure UpdateCumulsCEDTBALANCE ; override ;

    // -> Construit la requête de l'état 
    function  GetConditionSQLCpt                                           : String ; override ; // FQ 20697
    function  GetConditionSQLEcr                                           : String ; override ;
    function  GetConditionSQLPeriode( vInPer : Integer )                   : String ; override ;
    function  GetConditionSQLDateC  ( vInPer : Integer )                   : String ; override ;
    function  GetSQLCumul           ( vInPer : Integer ; vInCol : Integer ): String ; override ;

    // Nouvelles fonctions de maj des comptes
    function  GetJointureCpt                                               : String ; override ;

    // Tests autres méthodes (désactivé)
    procedure UpdateCumulsCEDTBALANCEbis ;

  private

    FNbPeriodes : Integer ;
    FDtArrete   : TDateTime ;
    FDtExoV8    : TDateTime ;

  end ;

Type TTabDate10 = Array[1..10] of TDateTime ; // tableau de dates pour le calcul des périodes

Implementation

uses
  {$IFDEF MODENT1}
  CPTypeCons,
  {$ENDIF MODENT1}
  uLibWindows, UTofMulParamGen; // TestJoker

//==============================================================================
procedure CPLanceFiche_BalanceAgee ( Args : String = '' ) ;
begin
  AGLLanceFiche('CP','CPBALAGEE','','','AGEE;' + Args );
end;
//==============================================================================
procedure CPLanceFiche_BalanceVentilee ( Args : String = '' ) ;
begin
  AGLLanceFiche('CP','CPBALAGEE','','','VENTIL;' + Args );
end;
//==============================================================================

{TOF_CPBALAGEE}

//==============================================================================
//====================         EVT TOF        ==================================
//==============================================================================


procedure TOF_CPBALAGEE.OnLoad ;
begin

  FDtArrete := StrToDate( DateArrete.Text ) ;

  Inherited ;

  if TypeBal = balVentil
    then SetControlText ( 'BALTYPE',  'VENTIL' )
    else SetControlText ( 'BALTYPE',  'AGEE' );

end ;

procedure TOF_CPBALAGEE.OnArgument (S : String ) ;
begin

  // Ajout type balance
  if pos( 'AGEE', S ) > 0
    then TypeBal := balAgee
    else TypeBal := balVentil ;

  // Date ExoV8
  if (VH^.ExoV8.Code <> '')
    then FDtExoV8 := VH^.ExoV8.Deb
    else FDtExoV8 := iDate1900 ;

  // Controles communs
  Inherited ;

  //=====> Param selon type de balance
  ParametreEtatType;

  // Type de plan comptable :
  Case V_PGI.LaSerie of
    S3 :  begin
          Corresp.plus := 'AND CO_CODE = "AU1"';
          end;
    else  begin
          Corresp.plus := 'AND (CO_CODE = "AU1" OR CO_CODE = "AU2")';
          end;
    end;

  // GCO 08/08/2003
  CompteDe.MaxLength := VH^.CPta[fbAux].Lg;
  CompteA.MaxLength  := VH^.CPta[fbAux].Lg;

  // SBO - 28/06/2004 - Initialisation du type de table libre pour les ruptures
  ChoixTypeTaLi.Visible := TL_TIERSCOMPL_Actif;
  ChoixTypeTaLi.ItemIndex := 0;
  InitChoixTableLibre;

  if GetParamSocSecur('SO_CPMULTIERS', false) then
  begin
    THEdit(GetControl('COMPTE', true)).OnElipsisClick:=AuxiElipsisClick;
    THEdit(GetControl('COMPTE_', true)).OnElipsisClick:=AuxiElipsisClick;
  end;

end ;

//==============================================================================
//====================        EVT QRS1        ==================================
//==============================================================================

{***********A.G.L.***********************************************
Auteur  ...... : YMO
Créé le ...... : 12/04/2007
Modifié le ... :   /  /
Description .. : Branchement de la fiche auxiliaire
Mots clefs ... :
*****************************************************************}
procedure TOF_CPBALAGEE.AuxiElipsisClick(Sender: TObject);
begin
     THEdit(Sender).text:= CPLanceFiche_MULTiers('M;' +THEdit(Sender).text + ';' +THEdit(Sender).Plus + ';');
end;

procedure TOF_CPBALAGEE.TableLibreChanged(Sender: TObject);
begin
  if TableLibre.ItemIndex < 0 then
    begin
    LibreDe.DataType := '';
    LibreA.DataType  := '';
    LibreDe.Text     := '';
    LibreA.Text      := '';
    end
  else
    begin
    if ChoixTypeTaLi.Value = 'AUX' then
      begin
      LibreDe.DataType := 'TZNATTIERS' + GetNumTableLibre ;
      LibreA.DataType  := 'TZNATTIERS' + GetNumTableLibre ;
      end
    else if ChoixTypeTaLi.Value = 'CLI' then
      begin
      LibreDe.DataType := 'GCLIBRETIERS' + GetNumTableLibre ;
      LibreA.DataType  := 'GCLIBRETIERS' + GetNumTableLibre ;
      end
    else if ChoixTypeTaLi.Value = 'SCL' then
      begin
      LibreDe.DataType := 'AFLRESSOURCE';
      LibreA.DataType  := 'AFLRESSOURCE';
      end;
    end;
end;

//==============================================================================
//====================   TRAITEMENTS DES DONNEES  ==============================
//==============================================================================

procedure TOF_CPBALAGEE.CompteOnExit(Sender: TObject);
begin
  if ( csDestroying in Ecran.ComponentState) then Exit ;
  if Trim(THEdit(Sender).Text) = '' then Exit;

  {JP 01/07/05 : on ne fait que l'auto-complétion que s'il n'y a pas de caractère joker}
  if HasJoker(Sender) then Exit;

  // Complétion auto du numéro de compte si possible
  if not CompleteAuto( Sender, fbAux ) then
    THEdit(Sender).ElipsisClick(nil);
end;

procedure TOF_CPBALAGEE.InitChoixTableLibre;
begin

  if ChoixTypeTaLi.Value = 'AUX' then
    begin
    SetControlProperty('TABLELIBRE','DataType','TTTABLESLIBRESAUX');
    SetControlProperty('TABLELIBRE','Plus','');
    end
  else if ChoixTypeTaLi.Value = 'CLI' then
    begin
    SetControlProperty('TABLELIBRE','DataType','GCZONELIBRETIE');
    SetControlProperty('TABLELIBRE','Plus','AND CC_CODE LIKE "CT%"');
    end
  else if ChoixTypeTaLi.Value ='SCL' then
    begin
    SetControlProperty('TABLELIBRE','DataType','GCZONELIBRETIE');
    SetControlProperty('TABLELIBRE','Plus','AND CC_CODE LIKE "CR%"');
    end;

  TableLibre.ReLoad;

  //SG6 14.03.05 Gestion des filtres FQ 14996
  if cLoadFiltre then Exit;
  
  TableLibre.ItemIndex := 0 ; 
  SetControlText ('LIBREDE','');
  SetControlText ('LIBREA','');

end;

procedure TOF_CPBALAGEE.OnClickChoixTypeTaLi(Sender: TObject);
begin
  InitChoixTableLibre;
  TableLibreChanged ( nil );
end;

function TOF_CPBALAGEE.GenererRequeteBAL: String;
begin
  // Champs
  result := 'SELECT CED_COMPTE CED_AUXILIAIRE, CED_NATURE, CED_LIBELLE' ;
  result := result
               + ', CED_DEBIT1, CED_CREDIT1, CED_DEBIT2, CED_CREDIT2'
               + ', CED_DEBIT3, CED_CREDIT3, CED_DEBIT4, CED_CREDIT4'
               + ', CED_DEBIT5, CED_CREDIT5, CED_DEBIT6, CED_CREDIT6'
               + ', CED_RUPTURE, CED_COLLECTIF ' ;

  if Rupture.Value <> 'SANS' then
    begin
    // Rupture sur compte de correspondance
    if (RuptureType.Value = 'RUPCORRESP') and (Corresp.ItemIndex >= 0) then
      result := result + ', CR_LIBELLE RUPTURELIB '
    // Rupture sur tables libres
    else if (RuptureType.Value = 'RUPLIBRES') and (TableLibre.ItemIndex >= 0) then
      begin
      if ChoixTypeTaLi.Value = 'AUX' then
        result := result + ', NT_LIBELLE RUPTURELIB '
      else if ChoixTypeTaLi.Value = 'CLI' then
        result := result + ', YX_LIBELLE RUPTURELIB '
      else if ChoixTypeTaLi.Value = 'SCL' then
        result := result + ', ARS_LIBELLE RUPTURELIB ' ;
      end
    // Nothing
    else result := result + ', "&#@" RUPTURELIB ' ;
    end
  // Pas de rupture
  else result := result + ', "&#@" RUPTURELIB ' ;

  // Tables
  result := result + ' FROM CEDTBALANCE ' ;
  if Rupture.Value <> 'SANS' then
    begin
    if (RuptureType.Value = 'RUPCORRESP') and (Corresp.ItemIndex >= 0) then
      result := result + ' LEFT JOIN CORRESP ON CR_CORRESP=CED_RUPTURE AND CR_TYPE="' + Corresp.Value + '"'
    else if (RuptureType.Value = 'RUPLIBRES') and (TableLibre.ItemIndex >= 0) then
      begin
      if ChoixTypeTaLi.Value = 'AUX' then
        result := result + ' LEFT JOIN NATCPTE ON NT_NATURE=CED_RUPTURE AND NT_TYPECPTE="' + TableLibre.Value + '"'
      else if ChoixTypeTaLi.Value = 'CLI' then
        result := result + ' LEFT JOIN CHOIXEXT ON YX_CODE=CED_RUPTURE AND YX_TYPE="LT' + GetNumTableLibre + '"'
      else if ChoixTypeTaLi.Value = 'SCL' then
        result := result + ' LEFT JOIN RESSOURCE ON ARS_RESSOURCE=CED_RUPTURE '
      end ;
    end ;

  // Conditions
  result := result + ' WHERE CED_USER = "' + V_PGI.User + '"' ;
  // Tri
  result := result + ' ORDER BY CED_RUPTURE, CED_COMPTE' ;

end;

function TOF_CPBALAGEE.GetNumTableLibre: String;
begin
  result := '' ;
  if TableLibre.ItemIndex < 0 then Exit ;
  result := Copy( TableLibre.Value, 3, 1 )  ;

end;

function TOF_CPBALAGEE.GetWhereNature: String;
var lStTous : String ;
    lStVal  : String ;
    lStText : String ;
    lTest   : String ;
begin
  result := '' ;
  lStTous := '<<' + TraduireMemoire('Tous') + '>>';
  if NatCptMulti.Tous or ( NatCptMulti.Text = lStTous ) or ( NatCptMulti.Text = '' ) then Exit ;

  lStText := NatCptMulti.Text ;
  ltest := NatCptMulti.getSQLValue ;

  while ( lStText <> '' ) do
    begin
    lStVal := readTokenSt(  lStText ) ;
    if Trim(lStVal) <> '' then
      if result = ''
        then result := '"' + lStVal + '"'
        else result := result + ',"' + lStVal + '"' ;
    end ;

  if result <> '' then
    result := ' T_NATUREAUXI IN (' + result + ') ' ;

end;

procedure TOF_CPBALAGEE.InitZones;
begin
  inherited;

  // Nature auxi
  NatCptMulti.text := '<<' + TraduireMemoire( 'Tous' ) + '>>' ;

  // Périodicité
  Periodicite.Value := '1M' ;

  // Date d'arrêté
  SetControlText('DATEARRETE', DateToStr(V_PGI.DateEntree) ) ;

  // Ecart
  Ecart.itemindex := 1 ;

  // Nb Jour d'écart
  NbJEcart.text   := '30' ;

  // Calcul des périodes
  EcartOnChange( nil ) ;

end;

procedure TOF_CPBALAGEE.ChargementCritEdt;
begin
  // FQ 15844 SBO 27/09/2005 : Application des critères sur nature auxi
  if (TheData <> nil) and (TheData is ClassCritEdt) then
    begin
    // Nature auxiliaire en multi :
    NatCptMulti.Text := ClassCritEdt(TheData).CritEdt.NatureCpt ;
//    NatureCptChanged( nil ) ;
    end ;

  inherited;

end;

procedure TOF_CPBALAGEE.InitComposants;
begin

  // CONTROLES SPECIFIQUES
  // Spécifique tables libres TIERSCOMPL
  ChoixTypeTaLi          := THRadioGroup(GetControl('CHOIXTYPETALI',True));
  ChoixTypeTaLi.OnClick  := OnClickChoixTypeTaLi;

  // Nature de compte en multi
  NatCptMulti	         := THMultiValComboBox(Getcontrol('NATURECPT', true));
  NatCptMulti.OnChange   := NatureCptChanged ;

  inherited;

  Periodicite           := THValComboBox(GetControl('PERIODICITE'));
  Periodicite.OnChange  := PeriodiciteOnChange;

  CollectifDe           := THEdit(GetControl('COLLECTIF'));
  CollectifDe.OnExit    := CollectifOnExit ; // FQ 18697
  CollectifA            := THEdit(GetControl('COLLECTIF_'));
  CollectifA.OnExit     := CollectifOnExit ; // FQ 18697

  DateArrete            := THEdit(GetControl('DATEARRETE'));
  DateArrete.OnChange   := DateArreteOnChange ;
  EnSituation           := TCheckBox(GetControl('ENSITUATION'));

  Ecart                 := THRadioGroup(GetControl('ECART'));
  Ecart.OnClick         := EcartOnChange ;
  NbJEcart              := TSpinEdit(GetControl('NBJECART'));
  NbJEcart.OnChange     := NbJEcartOnChange ;

  TypePrevi             := TRadioButton(GetControl('TYPEPREVI'));
  TypePrevi.OnClick     := TypeBalVentilOnChange ;
  TypeRetard            := TRadioButton(GetControl('TYPERETARD'));
  TypeRetard.OnClick    := TypeBalVentilOnChange ;

  ChoixTypeTaLi         := THRadioGroup(GetControl('CHOIXTYPETALI'));

end;

procedure TOF_CPBALAGEE.ParamComparatif;
begin

end;

procedure TOF_CPBALAGEE.ParamDivers;
var lStSQL : String ;
    lStLib : String ;
begin
  inherited;

  // Traduction des natures de pièces :
  TraductionTHMultiValComboBox( NatCptMulti, lStSQL, lStLib, 'T_NATUREAUXI' ) ;
  SetControlText( 'NATURECPTLIB', lStLib ) ;

end;

procedure TOF_CPBALAGEE.ParamRuptures;
begin
  inherited;

  // Les ruptures sur champ libre tiers
  if ( Rupture.Value <> 'SANS')  and (RuptureType.Value = 'RUPLIBRES') and (TableLibre.ItemIndex >= 0) then
    begin
      // Tables libres tiers
      if ChoixTypeTaLi.Value = 'AUX'
        then XXRupture.Text := 'T_TABLE' + GetNumTableLibre
      // Table libre clients
      else if ChoixTypeTaLi.Value = 'CLI'
        then XXRupture.Text := 'YTC_TABLELIBRETIERS' + GetNumTableLibre
      // Tables libres ressource
      else if ChoixTypeTaLi.Value = 'SCL'
        then XXRupture.Text := 'YTC_RESSOURCE' + IntToStr(TableLibre.ItemIndex + 1) ;
    end ;

end;

procedure TOF_CPBALAGEE.ParamTitresCols;
begin

  Case TypeBal of
    balAgee :   begin
                SetControlText( 'TITREETAT', 'Balance âgée' ) ;

                // 1ère Période
                SetControlText( 'TITRECOL1' ,  'Avant le' ) ;
                SetControlText( 'TITRECOL1_' , GetControlText('PERIODE1') ) ;

                // 2ème Période
                SetControlText( 'TITRECOL2' ,  'Du ' + GetControlText('PERIODE1') ) ;
                SetControlText( 'TITRECOL2_' , 'Au ' + GetControlText('PERIODE6') ) ;

                // 3ème Période
                SetControlText( 'TITRECOL3' ,  'Du ' + GetControlText('PERIODE2') ) ;
                SetControlText( 'TITRECOL3_' , 'Au ' + GetControlText('PERIODE7') ) ;

                // 4ème Période
                SetControlText( 'TITRECOL4' ,  'Du ' + GetControlText('PERIODE3') ) ;
                SetControlText( 'TITRECOL4_' , 'Au ' + GetControlText('PERIODE8') ) ;

                // 5ème Période
                SetControlText( 'TITRECOL5' ,  'Du ' + GetControlText('PERIODE4') ) ;
                SetControlText( 'TITRECOL5_' , 'Au ' + GetControlText('PERIODE9') ) ;

                // 6ème Période
                SetControlText( 'TITRECOL6' ,  'Du ' + GetControlText('PERIODE5') ) ;
                SetControlText( 'TITRECOL6_' , 'Au ' + GetControlText('PERIODE10') ) ;
                end ;

    balVentil : begin
                SetControlText( 'TITREETAT', 'Balance ventilée' ) ;

                // 1ère Période
                if TypePrevi.Checked then
                  begin
                  SetControlText( 'TITRECOL1' ,  'Déjà échus' ) ;
                  SetControlText( 'TITRECOL1_' , '' ) ;
                  end
                else
                  begin
                  SetControlText( 'TITRECOL1' ,  'Avant le' ) ;
                  SetControlText( 'TITRECOL1_' , GetControlText('PERIODE1') ) ;
                  end ;

                // 2ème Période
                SetControlText( 'TITRECOL2' ,  'Du ' + GetControlText('PERIODE1') ) ;
                SetControlText( 'TITRECOL2_' , 'Au ' + GetControlText('PERIODE6') ) ;

                // 3ème Période
                SetControlText( 'TITRECOL3' ,  'Du ' + GetControlText('PERIODE2') ) ;
                SetControlText( 'TITRECOL3_' , 'Au ' + GetControlText('PERIODE7') ) ;

                // 4ème Période
                SetControlText( 'TITRECOL4' ,  'Du ' + GetControlText('PERIODE3') ) ;
                SetControlText( 'TITRECOL4_' , 'Au ' + GetControlText('PERIODE8') ) ;

                // 5ème Période
                SetControlText( 'TITRECOL5' ,  'Du ' + GetControlText('PERIODE4') ) ;
                SetControlText( 'TITRECOL5_' , 'Au ' + GetControlText('PERIODE9') ) ;

                // 6ème Période
                if TypePrevi.Checked then
                  begin
                  SetControlText( 'TITRECOL6' ,  'Après le' ) ;
                  SetControlText( 'TITRECOL6_' , GetControlText('PERIODE9')  ) ;
                  end
                else
                  begin
                  { FQ 15731 BVE 19.04.07 }
                  if TypeRetard.Checked then
                     SetControlText( 'TITRECOL6' ,  'A échoir'  )
                  else
                     SetControlText( 'TITRECOL6' ,  'Déjà échus'  ) ;
                  { END FQ 15731 }
                  SetControlText( 'TITRECOL6_' , ''  ) ;
                  end

                end ;
    end ;

end;

procedure TOF_CPBALAGEE.ParametreEtatType;
var stTitre : String;
begin

  // BALANCE VENTILEE
  if TypeBal = balVentil then
    begin
    FNbPeriodes := 4 ;
    TFQRS1(Ecran).FNomFiltre := TFQRS1(Ecran).FNomFiltre + 'V' ;
    Ecran.HelpContext := 7556000;

    // Modif titre QRS1
    stTitre := 'Balance ventilée en situation';

    // Pour la balance ventilé seulement 4 périodes :
    SetControlVisible( 'TPERIODE5',   False);
    SetControlVisible( 'PERIODE5',    False);
    SetControlVisible( 'TPERIODE10',  False);
    SetControlVisible( 'PERIODE10',   False);

    //Affiche du type de balance (Prevision / Retard)
    SetControlVisible(  'TBALTYPE',   True);
    SetControlVisible(  'TYPEPREVI',  True);
    SetControlVisible(  'TYPERETARD', True);
    SetControlProperty( 'TYPERETARD', 'checked',  False);
    SetControlProperty( 'TYPEPREVI',  'checked',  True);

    end
  else
    // BALANCE AGEE
    begin

    FNbPeriodes := 5 ;
    Ecran.HelpContext := 7547000;

    // Modif titre QRS1
    stTitre := 'Balance âgée en situation';

    //Cache le type de balance ventilée (Prevision / Retard)
    SetControlVisible( 'TBALTYPE',   False);
    SetControlVisible( 'TYPEPREVI',  False);
    SetControlVisible( 'TYPERETARD', False);

    end;

  TFQRS1(Ecran).Caption := stTitre;
  Ecran.Caption :=TraduireMemoire(Ecran.Caption) ;
  UpdateCaption(Ecran);

end;

procedure TOF_CPBALAGEE.CalculPeriodesAvecPeriodicite;
var lInPerio : Integer ;
    i        : Integer ;
    an       : Word ;
    mois     : Word ;
    jour     : Word ;
    JMax     : Word ;
    DAT      : TDateTime ;
    lBoPrevi : Boolean;
    TabD     : TTabDate10;
begin
  // Initalisation et paramétrage suivant type état
  FillChar(TabD,SizeOf(TabD),#0);

  lBoPrevi := ( TypeBal = balVentil ) and ( TypePrevi.Checked ) ;
  if lBoPrevi
    then TabD[ 1 ]               := StrToDate( DateArrete.Text ) + 1   // si prévision, date d'arrete = date depart
    else TabD[ FNbPeriodes + 5 ] := StrToDate( DateArrete.Text ) ;     // sinon, date d'arrete = date de fin

  // Calcul en fonction de la périodicité choisie
  lInPerio := GetPeriodiciteIndex ;
  case lInPerio of
    // ===> Périodicité N mois
    0,1,2,3,4,5 :
      if lBoPrevi then
        // Pour chaque Fourchette de dates( en partant de la première) , en mensuel
        for i:=1 to FNbPeriodes do
          begin
          TabD[i+5] := PlusMois( TabD[i], (lInPerio+1) ) - 1 ;
          if i<FNbPeriodes then TabD[i+1]:=TabD[i+5]+1 ;
          end
      else
        // Pour chaque Fourchette de dates( en partant de la derniére) , en mensuel
        for i := FNbPeriodes downto 1 do
          begin
          // Borne supérieure
          DAT  := PlusMois( TabD[i+5], -(lInPerio+1) ) ; // On ote N Mois (la périodicité )
          // Borne inférieure
          DecodeDate( DAT, an, mois, jour ) ;
          JMax    := StrToInt( FormatDateTime( 'd', FinDeMois( EncodeDate(an,mois,1) ) ) ) ;
          TabD[i] := PlusMois( TabD[i+5], -(lInPerio+1) ) + ( JMax - jour ) + 1 ;
          // Borne supérieure précédente
          if i>1 then
            TabD[i+4] := TabD[i] - 1 ;
          end ;
    // ===> Périodicité de 15 jours
    6           :
      if lBoPrevi then
        // Pour chaque Fourchette de dates (en partant de la première) , en Quinzaine
        for i:=1 to FNbPeriodes do
           begin
           TabD[i+5]:=TabD[i]+ 15 ;
           if i<5 then TabD[i+1]:=TabD[i+5]+1 ;
           end
      else
        // Pour chaque Fourchette de dates( en partant de la derniére) , en Quinzaine
        for i:=5 downto 1 Do
           begin
           DAT  := TabD[ i + 5 ] ;
           DecodeDate( DAT, an, mois, jour ) ;
           JMax := StrToInt(FormatDateTime( 'd', FinDeMois( EncodeDate( an, mois, 1) ) ) ) ;
           if Jour<=15
             // Date départ = (Date d'arrivée - 15 jours) + 1 jour, si date avant le 15 du mois
             then TabD[i] := TabD[ i + 5 ] - 15 + 1
             // Date départ = (Date d'arrivée - (Nb jours Max du mois - 15 jours) ) + 1 jour, si date aprés le 15 du mois }
             else TabD[i] := TabD[ i + 5 ] - ( JMax - 15 ) + 1 ;
           if i > 1 then
             TabD[ i + 4 ] := TabD[ i ] - 1 ;
           end;
    // ===> Périodicité de 1 semaine
    7           :
      if lBoPrevi then
        // Pour chaque Fourchette de dates( en partant de la première) , en Hebdo
        for i:=1 to FNbPeriodes Do
          begin
          TabD[i+5]:=TabD[i]+7;
          if i<5 then TabD[i+1]:=TabD[i+5]+1 ;
          end
      else
        // Pour chaque Fourchette de dates( en partant de la derniére) , en Hebdo
	for i:=5 downto 1 Do
          begin
          // Date départ = (Date d'arrivée - (7 jours+ 1 jours) )
          TabD[i]:=TabD[i+5]-(7)+1 ;
          if i>1 then TabD[i+4]:=TabD[i]-1 ;
          end ;
  	end ;

  // Remplissage des contrôles
  for i := 1 to 10 do
    SetControlText( 'PERIODE' + IntToStr(i), DateToStr(TabD[i]) );

end;

procedure TOF_CPBALAGEE.CalculPeriodesAvecNbJour;
var lInNbJ   : Integer ;
    i        : Integer ;
    lBoPrevi : Boolean;
    TabD     : TTabDate10;
begin
  // Initalisation et paramétrage suivant type état
  FillChar(TabD,SizeOf(TabD),#0);

  lBoPrevi := ( TypeBal = balVentil ) and ( TypePrevi.Checked ) ;
  lInNbJ   := NbJEcart.value;

  // Calcul des périodes
  if lBoPrevi then
    // si prévision, date d'arrete = date depart
    begin
    TabD[1] := StrToDate(DateArrete.Text) + 1;
    for i:=1 to FNbPeriodes do
      begin
      TabD[i+5] := TabD[i] + lInNbJ;
      if i<5 then
        TabD[i+1] := TabD[i+5] + 1;
      end;
    end
  else
    // sinon, date d'arrete = date de fin
    begin
    TabD[ FNbPeriodes + 5 ] := StrToDate(DateArrete.Text);
    for i:=FNbPeriodes downto 1 do
      begin
      TabD[i] := TabD[ i + 5 ] - ( lInNbJ - 1 );
      if i>1 then
        TabD[i+4]:=TabD[i]-1;
      end;
    end;

	// Remplissage des contrôles
  for i:=1 to 10 do
    SetControlText('PERIODE'+IntToStr(i),DateToStr(TabD[i]));
end;

procedure TOF_CPBALAGEE.DateArreteOnChange(Sender: TObject);
begin

  if not IsValidDate( DateArrete.Text ) then exit;

  FDtArrete := StrToDate( DateArrete.Text ) ;

  if ( TypeBal = balVentil ) and ( TypePrevi.Checked )
    then SetControlText( 'PERIODE1', GetControlText('DATEARRETE') )
    else SetControlText( 'PERIODE' + IntToStr( FNbPeriodes + 5 ), DateArrete.Text ) ;

  EcartOnChange(nil);

end;

procedure TOF_CPBALAGEE.EcartOnChange(Sender: TObject);
begin

  SetControlVisible( 'PERIODES', Ecart.ItemIndex = 1 ) ;
  SetControlVisible( 'NBJECART', Ecart.ItemIndex <> 1  ) ;


  if Ecart.ItemIndex=0
    then NbJEcartOnChange(nil)
    else PeriodiciteOnChange(nil);

end;

procedure TOF_CPBALAGEE.NbJEcartOnChange(Sender: TObject);
begin
  if GetControlText('NBJECART')='' then exit;
  CalculPeriodesAvecNbJour;
end;

procedure TOF_CPBALAGEE.PeriodiciteOnChange(Sender: TObject);
begin
  CalculPeriodesAvecPeriodicite ;
end;


{***********A.G.L.***********************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 14/06/2007
Modifié le ... :   /  /    
Description .. : On ne prend en compte que les TIERS lettrable
Suite ........ : FQ 20697
Mots clefs ... : 
*****************************************************************}
function TOF_CPBALAGEE.GetConditionSQLCpt: String;
begin
  result := inherited GetConditionSQLCpt + ' AND T_LETTRABLE = "X"';
end;

function TOF_CPBALAGEE.GetConditionSQLEcr: String;
begin

  result := inherited GetConditionSQLEcr ;

  // On vire les A-Nouveaux non typé "H"
  Result := Result + ' AND E_ECRANOUVEAU<>"CLO" AND E_ECRANOUVEAU<>"OAN" ' ;

  // Condition sur les collectifs :
  result := result + ' AND ' + ConvertitCaractereJokers(CollectifDe, CollectifA, 'E_GENERAL' );

  // ====================
  // ==== LE LETTRAGE ===
  // ====================
  // Condition sur le lettrage
  if EnSituation.Checked
    then  result := result + ' AND ( (E_ETATLETTRAGE<>"TL" AND E_ETATLETTRAGE<>"RI")'
                              + ' OR (E_ETATLETTRAGE="TL" AND E_DATEPAQUETMAX>"'+ USDateTime(StrToDate(DateArrete.Text)) + '"))'
    else  result := result + ' AND E_ETATLETTRAGE<>"TL" AND E_ETATLETTRAGE<>"RI"';

end;

function TOF_CPBALAGEE.GetConditionSQLPeriode( vInPer : Integer ) : String;
var lStChp : String ;
    lDtDeb : TDateTime ;
    lDtFin : TDateTime ;
    function _GetSqlPeriode : String ;
      begin
      if TypeBal = balAgee then
        lDtDeb := max( lDtDeb, FDtExoV8 ) ;
      result := ' ' + lStChp + ' BETWEEN "' + USDateTime( lDtDeb ) + '" AND "' + USDateTime( lDtFin ) + '"' ;
      end ;
begin

  if TypeBal = balAgee
    then lStChp := 'E_DATECOMPTABLE'
    else lStChp := 'E_DATEECHEANCE' ;

  Case vInPer of

      // 1ère période
      1 : begin
          lDtDeb := FDtExoV8 ;
          lDtFin := StrToDate( GetControlText('PERIODE1') ) - 1 ;
          if TypeBal = balAgee
            then result := _GetSqlPeriode
            else result := ' ' + lStChp + ' <= "' + USDateTime( lDtFin ) + '" ' ;
          end ;

      // Périodes communes
      2..5 : begin
             lDtDeb := StrToDate( GetControlText('PERIODE' + IntToStr(vInPer - 1) ) ) ;
             lDtFin := StrToDate( GetControlText('PERIODE' + IntToStr(vInPer + 4) ) ) ;
             result := _GetSqlPeriode ;
             end ;

      // Pour la balance ventilée, correspond aux écritures > 5ème période
      6 : if TypeBal = balAgee then
            begin
            lDtDeb := StrToDate( GetControlText('PERIODE5') ) ;
            lDtFin := StrToDate( GetControlText('PERIODE10') ) ;
            result := _GetSqlPeriode ;
            end
          else result := ' ' + lStChp + ' > "' + USDateTime( StrToDate( GetControlText('PERIODE9') ) ) + '" ' ;
      end ;

  // Pour les balances ventilées, on ajoute le critère sur la date comptable
  //   car les périodes sont basées sur les dates d'échéances
  if ( TypeBal = balVentil)  and ( vInPer <> 0 ) then
    result := result + ' AND ' + GetConditionSQLDateC( 0 ) ;


end;

{***********A.G.L.***********************************************
Auteur  ...... : Brice Verchère
Créé le ...... : 03/10/2007
Modifié le ... :   /  /    
Description .. : Deplacement du traitement specifique de GetSQLCumum 
Suite ........ : pour les balances agée et ventilée ici :
Suite ........ : FQ 21569
Mots clefs ... : 
*****************************************************************}
function TOF_CPBALAGEE.GetSQLCumul(vInPer, vInCol: Integer): String;
begin
  // --> Champs
    Case vInCol of
  { FQ 17929 BVE 21.06.07 }
  // Crédit + débit
  0 : if Affichage.Value = 'EURO'
        then result := 'SELECT SUM(IIF((E_ETATLETTRAGE="PL" AND E_DEBIT <> 0),E_DEBIT - E_COUVERTURE, E_DEBIT)) D, ' +
                       'SUM(IIF((E_ETATLETTRAGE="PL" AND E_CREDIT <> 0),E_CREDIT - E_COUVERTURE, E_CREDIT)) C'
        else result := 'SELECT SUM(IIF((E_ETATLETTRAGE="PL" AND E_DEBITDEV <> 0),E_DEBITDEV - E_COUVERTUREDEV, E_DEBITDEV) D, ' +
                       'SUM(IIF((E_ETATLETTRAGE="PL" AND E_CREDITDEV <> 0),E_CREDITDEV - E_COUVERTUREDEV, E_CREDITDEV)) C';

  // Débit uniquement
  1 : if Affichage.Value = 'EURO'
        then result := 'SELECT SUM(IIF((E_ETATLETTRAGE="PL" AND E_DEBIT <> 0),E_DEBIT - E_COUVERTURE, E_DEBIT))'
        else result := 'SELECT SUM(IIF((E_ETATLETTRAGE="PL" AND E_DEBITDEV <> 0),E_DEBITDEV - E_COUVERTUREDEV, E_DEBITDEV))';

  // Crédit uniquement
  2 : if Affichage.Value = 'EURO'
        then result := 'SELECT SUM(IIF((E_ETATLETTRAGE="PL" AND E_CREDIT <> 0),E_CREDIT - E_COUVERTURE, E_CREDIT))'
        else result := 'SELECT SUM(IIF((E_ETATLETTRAGE="PL" AND E_CREDITDEV <> 0),E_CREDITDEV - E_COUVERTUREDEV, E_CREDITDEV))';
  end ;
  { END FQ 17929 }
  // --> Table
  result := result + ' FROM ' + GetTablePourBase('ECRITURE') ;

  // --> Conditions
  result := result + ' WHERE ' + GetConditionSQL( False, vInPer ) ;

  // ---> Lien Compte // champ CED_COMPTE
  result := result + ' AND E_AUXILIAIRE = CED_COMPTE ' ;

end;

function TOF_CPBALAGEE.GetConditionSQLDateC ( vInPer : Integer ) : String;
var lDtFin : TDateTime ;
begin
  // ==================
  // ==== LES DATES ===
  // ==================
  // borne supérieure en fonction du type de balance
  if ( TypeBal = balVentil ) then
    begin
    if ( TypePrevi.Checked )
      then lDtFin := StrToDate( GetControlText('PERIODE1' ) ) - 1 //FQ18802 YMO 18/09/2006 : Décalage d'un jour
      else lDtFin := StrToDate( GetControlText('PERIODE9') );     //FQ18802 YMO 18/09/2006 : 9 et non pas 8 (4 périodes et non pas 3)
    end
  else lDtFin := StrToDate( GetControlText('PERIODE10') );
{
  if TypeBal = balAgee
    then lDtFin := StrToDate( GetControlText('PERIODE10') )
    else lDtFin := StrToDate( GetControlText('PERIODE1' ) );
}
  // Borne inférieurs si ExoV8 défini
  if FDtExoV8 > iDate1900
    then result := ' E_DATECOMPTABLE BETWEEN "' + UsDateTime( FDtExoV8 ) + '" AND "' + UsDateTime( lDtFin ) + '" '
    else result := ' E_DATECOMPTABLE <= "' + UsDateTime( lDtFin ) + '" ' ;

end;

function TOF_CPBALAGEE.GetPeriodiciteIndex: Integer;
begin
  if Periodicite.Value = '1M' then
    result := 0
  else  if Periodicite.Value = '2M' then
    result := 1
  else  if Periodicite.Value = '3M' then
    result := 2
  else  if Periodicite.Value = '4M' then
    result := 3
  else  if Periodicite.Value = '5M' then
    result := 4
  else  if Periodicite.Value = '6M' then
    result := 5
  else  if Periodicite.Value = 'QUI' then
    result := 6
  else  if Periodicite.Value = 'SEM' then
    result := 7
  else result := -1 ;

end;

procedure TOF_CPBALAGEE.UpdateCumulsCEDTBALANCE;
Var lStReq      : String ;
    i           : Integer ;
    lStChp      : String ;
begin

  try  // finally
    try  // except

      // -----------------------------------------
      // Chargement de la liste des comptes cibles
      for i := 1 to 6 do // Pour les 6 périodes
        begin

        // Champ Débit
        lStChp := 'CED_DEBIT' + IntToStr( i ) ;

        lStReq := 'UPDATE ' + GetTablePourBase('CEDTBALANCE') + ' SET ' + lStChp + ' = ( ' + GetSQLCumul( i, 1 ) + ' )'
                                                              + ' WHERE CED_USER="'+V_PGI.User+'"' ;
        ExecuteSQL( lStReq ) ;
        // Le problème des requêtes précédentes est que les champs sont mis à NULL si le compte n'est pas mouvementé sur la période donnée
        ExecuteSQL('UPDATE ' + GetTablePourBase('CEDTBALANCE') + ' SET ' + lStChp + ' = 0 WHERE ' + lStChp + ' IS NULL AND CED_USER="' + V_PGI.User + '"' ) ;

        // Champ Crédit
        lStChp := 'CED_CREDIT' + IntToStr( i ) ;
        lStReq := 'UPDATE ' + GetTablePourBase('CEDTBALANCE') + ' SET ' + lStChp + ' = ( ' + GetSQLCumul( i, 2 ) + ' )'
                                                              + ' WHERE CED_USER="'+V_PGI.User+'"' ;
        ExecuteSQL( lStReq ) ;
        // Le problème des requêtes précédentes est que les champs sont mis à NULL si le compte n'est pas mouvementé sur la période donnée
        ExecuteSQL('UPDATE ' + GetTablePourBase('CEDTBALANCE') + ' SET ' + lStChp + ' = 0 WHERE ' + lStChp + ' IS NULL AND CED_USER="' + V_PGI.User + '"' ) ;

        end ;

    // Try Except
    except
      on E : Exception do
      begin
        PgiError( E.Message, Ecran.Caption );
      end;
    end;
  // Try Finally
  finally
//    FiniMoveProgressForm ;
  end ;

end;

procedure TOF_CPBALAGEE.SetTypeBalance;
begin
  TableEcr    := 'ECRITURE' ;
  PfEcr       := 'E' ;
  TableCpt    := 'TIERS' ;
  PfCpt       := 'T' ;
  ChampCpt    := 'AUXILIAIRE' ;
  ChampNatCpt := 'T_NATUREAUXI' ;
end;


procedure TOF_CPBALAGEE.UpdateCumulsCEDTBALANCEbis;
var lStCompte : String ;
    lStReq    : String ;
    i         : Integer ;
    QSum      : TQuery ;
begin

   try  // except

      // --> Calcul des totaux pour les 6 colonnes
      for i := 1 to 6 do
        begin

        // Crédit + débit
        if Affichage.Value = 'EURO'
          then lStReq := 'SELECT CED_COMPTE, CED_LIBELLE, SUM(E_DEBIT) D, SUM(E_CREDIT) C'
          else lStReq := 'SELECT CED_COMPTE, SUM(E_DEBITDEV) D, SUM(E_CREDITDEV) C';
        // --> Tables
        lStReq := lStReq + ' FROM ' + GetTablePourBase('ECRITURE') ;
        lStReq := lStReq + ' LEFT JOIN ' + GetTablePourBase('CEDTBALANCE') + ' ON CED_COMPTE = E_AUXILIAIRE' ;
        // --> Conditions
        lStReq := lStReq + ' WHERE ' + GetConditionSQL( False, i )
                         + ' AND CED_COMPTE IS NOT NULL '
                         + ' AND CED_USER = "' + V_PGI.User + '"' ;
        // --> Group by
        lStReq := lStReq + ' GROUP BY CED_COMPTE, CED_LIBELLE ' ;

        // ----------
        // Récupération des montants
        // ----------
        QSum := OpenSQL( lStReq, True);

    // Barre de progession
{
    InitMoveProgressForm( Ecran,
                          Ecran.Caption,
                          'Traitement des écritures en cours...',
                          100 ,
                          True,
                          False);
}

        // Parcours des comptes
        While not QSum.Eof do
          begin
          lStCompte   := QSum.FindField('CED_COMPTE').AsString ;
  //        lStLibelle  := QTiers.FindField('CED_LIBELLE').AsString ;
  //        if Not MoveCurProgressForm(lStCompte + ' - ' + lStLibelle) then break;

          MoveCur( False ) ;

          // -----------
          // REQUETE MAJ
          lStReq := 'UPDATE ' + GetTablePourBase('CEDTBALANCE') + ' SET '
                          + 'CED_DEBIT' + IntToStr(i) + '=' + VariantToSQL( QSum.FindField('D').AsFloat )
                        + ', CED_CREDIT'  + IntToStr(i) + '='  + VariantToSQL( QSum.FindField('C').AsFloat )
                        + ' WHERE CED_COMPTE = "' + lStCompte + '"'
                          + ' AND CED_USER = "' + V_PGI.User + '"' ;

          // --------
          // MAJ BASE
          ExecuteSQL( lStReq ) ;

          //---------
          // Suivant
          QSum.Next ;

          end ; // fin boucle sur le sum

        Ferme(QSum) ;

        end ; // fin boucle sur période

  // Try Except
  except
    on E : Exception do
    begin
      PgiError( E.Message, Ecran.Caption );
    end;
  end;


end;


function TOF_CPBALAGEE.GetJointureCpt: String;
begin
  // Table TIERSCOMPL pour rupture si besoin
  if ( Rupture.value <> 'SANS' )  and ( RuptureType.Value = 'RUPLIBRES' ) and
     ( ( ChoixTypeTaLi.Value = 'CLI' ) or ( ChoixTypeTaLi.Value = 'SCL' ) )
    then result := result + ' LEFT JOIN ' + GetTablePourBase('TIERSCOMPL') + ' ON YTC_AUXILIAIRE = T_AUXILIAIRE'
    else result := '' ;
end;

procedure TOF_CPBALAGEE.TypeBalVentilOnChange(Sender: TObject);
begin
  DateArreteOnChange( nil )
end;

procedure TOF_CPBALAGEE.CollectifOnExit(Sender: TObject);
begin
  if ( csDestroying in Ecran.ComponentState) then Exit ;
  // FQ 18697
  if Trim(THEdit(Sender).Text) = '' then Exit;

  {JP 01/07/05 : on ne fait que l'auto-complétion que s'il n'y a pas de caractère joker}
  if HasJoker(Sender) then Exit;

  // Complétion auto du numéro de compte si possible
  if not CompleteAuto( Sender, fbGene ) then
    THEdit(Sender).ElipsisClick(nil);
end;

Initialization
  registerclasses ( [ TOF_CPBALAGEE ] ) ;
end.
